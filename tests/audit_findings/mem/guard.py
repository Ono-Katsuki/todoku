"""監査用の見張り：書類を probe.py で動かす。nice 10・別のプロセス・常駐 1.5 GB か 1 道具 60 秒で殺す（1 度に 1 つだけ）。

使い方：guard.py 書類 [道具,道具|all] [秒] [GB]
殺された道具は記録して飛ばし、残りの道具を別のプロセスで続ける（max_kills まで）。
"""
import fcntl
import json
import os
import subprocess
import sys
import tempfile
import threading
import time

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.abspath(os.path.join(HERE, "..", "..", "..", "src"))
ALL = ["summary", "window", "fields", "grade", "audit", "fill", "view_page", "read_table",
       "calc_model", "calculate", "run_macro"]


def rss_kb(pid):
    try:
        return int(subprocess.run(["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True).stdout.strip() or 0)
    except ValueError:
        return 0


LOCK = os.path.join(tempfile.gettempdir(), "todoku_mem_guard.lock")


def one(path, tools, limit_s, limit_gb, env_extra=None, outdir=None):
    """1 度に 1 つだけ：ほかの見張りが動いていれば、終わるのを待つ（ファイルのロック）。"""
    with open(LOCK, "w") as lk:
        fcntl.flock(lk, fcntl.LOCK_EX)
        return _one(path, tools, limit_s, limit_gb, env_extra, outdir)


def _one(path, tools, limit_s, limit_gb, env_extra=None, outdir=None):
    """(各道具の記録, 殺された道具 or None, 理由, 見えた最大常駐 MB, 秒)"""
    env = dict(os.environ, PYTHONUNBUFFERED="1", PYTHONPATH=SRC, TODOKU_ISOLATE="0", **(env_extra or {}))
    cmd = ["nice", "-n", "10", sys.executable, os.path.join(HERE, "probe.py"), path, ",".join(tools)]
    if outdir:
        cmd.append(outdir)
    p = subprocess.Popen(cmd, env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, errors="replace")
    state = {"cur": None, "t0": time.time(), "recs": [], "tool_peak": 0}

    def reader():
        for line in p.stdout:
            try:
                d = json.loads(line)
            except ValueError:
                continue
            if d.get("ev") == "START":
                state["cur"], state["t0"] = d["tool"], time.time()
            elif d.get("ev") == "END":
                d["peak"] = state["tool_peak"]
                state["recs"].append(d)
                state["cur"] = None
                state["tool_peak"] = 0
    th = threading.Thread(target=reader, daemon=True)
    th.start()
    errs = []
    te = threading.Thread(target=lambda: errs.append(p.stderr.read()), daemon=True)
    te.start()
    peak = 0
    killed, why = None, ""
    start = time.time()
    while p.poll() is None:
        r = rss_kb(p.pid)
        peak = max(peak, r)
        state["tool_peak"] = max(state["tool_peak"], round(r / 1024))
        if r > limit_gb * 1048576:
            why = f"RSS>{limit_gb}GB"
        elif time.time() - state["t0"] > limit_s:
            why = f"time>{limit_s}s"
        if why:
            killed = state["cur"] or "?"
            p.kill()
            break
        time.sleep(0.03)
    p.wait()
    th.join(1)
    te.join(1)
    rc = p.returncode
    if not killed and rc != 0:
        # 自分で落ちた（セグメンテーション違反など）
        killed = state["cur"] or "?"
        why = f"died rc={rc}: {(errs[0] if errs else '')[-160:]}"
    return state["recs"], killed, why, round(peak / 1024), time.time() - start


def guarded(path, tools=None, limit_s=60, limit_gb=1.5, max_kills=2, env_extra=None, outdir=None):
    todo = list(tools or ALL)
    res, kills = [], []
    peak = 0
    while todo:
        recs, killed, why, pk, _dt = one(path, todo, limit_s, limit_gb, env_extra, outdir)
        peak = max(peak, pk)
        res += recs
        done = {r["tool"] for r in recs}
        if killed:
            kills.append({"tool": killed, "why": why, "rss": pk})
            done.add(killed)
            if len(kills) >= max_kills:
                res.append({"ev": "SKIP", "tools": [t for t in todo if t not in done]})
                break
        elif not recs:
            break
        todo = [t for t in todo if t not in done]
    return {"recs": res, "kills": kills, "peak": peak}


def main():
    path = sys.argv[1]
    tools = None if len(sys.argv) < 3 or sys.argv[2] == "all" else sys.argv[2].split(",")
    s = float(sys.argv[3]) if len(sys.argv) > 3 else 60
    gb = float(sys.argv[4]) if len(sys.argv) > 4 else 1.5
    r = guarded(path, tools, s, gb, max_kills=99)
    for x in r["recs"]:
        print(json.dumps({k: v for k, v in x.items() if k != "ev"}, ensure_ascii=False))
    for k in r["kills"]:
        print("KILLED", json.dumps(k, ensure_ascii=False))
    print(f"== {os.path.basename(path)} peak={r['peak']}MB")


if __name__ == "__main__":
    main()
