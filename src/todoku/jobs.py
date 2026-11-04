"""長い仕事（大きな書類の検査・大きな配列の集計）：1 回の道具の呼び出しの時間（limits.TIME_LIMIT）に縛られない、別のプロセスの仕事。

- audit_document を呼ぶと、仕事を始めて（すでに動いていれば、その進み具合を）すぐ返す。終わるまで何度でも呼べる
- 区切りごとに控え（state.json）を書くので、止まっても（PC を閉じた・メモリの上限など）次の呼び出しで続きから進む
- 控えの場所：TODOKU_JOBS_DIR（なければ ~/.cache/todoku/jobs）。ファイルの中身のうち、控えに入るのは検査で見つけた問題の短い字だけ。
  パスワードは控えに書かない（標準入力で仕事に渡す。止まった仕事を続けるときは、もう一度 password を渡す）
- 同じファイル（パス・大きさ・更新時刻・文字コード）の仕事は 1 つだけ
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import threading
import time
from pathlib import Path

from . import limits

SAVE_EVERY = 1.0  # 控えを書く間隔（秒）


def jobs_dir() -> Path:
    d = Path(os.environ.get("TODOKU_JOBS_DIR") or Path.home() / ".cache" / "todoku" / "jobs")
    d.mkdir(parents=True, exist_ok=True, mode=0o700)
    return d


def job_key(path: str, encoding: str | None) -> str:
    p = Path(str(path).split("!/")[0]).expanduser()
    st = p.stat()
    return hashlib.sha256(f"{p.resolve()}|{path}|{st.st_size}|{st.st_mtime_ns}|{encoding}".encode()).hexdigest()[:24]


def _state_path(key: str) -> Path:
    return jobs_dir() / f"{key}.json"


def load(key: str) -> dict | None:
    try:
        return json.loads(_state_path(key).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None


def save(key: str, record: dict) -> None:
    p = _state_path(key)
    tmp = p.with_suffix(f".{os.getpid()}.tmp")
    fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(record, f, ensure_ascii=False, default=str)
    os.replace(tmp, p)


def _alive(pid) -> bool:
    if not pid:
        return False
    try:
        os.kill(int(pid), 0)
    except (OSError, ValueError):
        return False
    # 自分が立ち上げた子が終わってゾンビになっているとき（kill(0) は成功する）を、片づける
    try:
        got, _st = os.waitpid(int(pid), os.WNOHANG)
        if got:
            return False
    except (ChildProcessError, OSError):
        pass
    return True


def status(path: str, encoding: str | None = None) -> dict | None:
    return load(job_key(path, encoding))


def start(path: str, password: str | None = None, encoding: str | None = None, restart: bool = False) -> dict:
    """仕事を始める（動いていれば何もしない。止まった仕事は控えの続きから）。いまの控えを返す。"""
    key = job_key(path, encoding)
    rec = load(key)
    if rec and rec.get("finished") and not restart:
        return rec
    if rec and _alive(rec.get("pid")) and not restart:
        return rec
    if restart:
        rec = None
        _state_path(key).unlink(missing_ok=True)
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    env = dict(os.environ)
    env["PYTHONPATH"] = os.pathsep.join([here] + [p for p in env.get("PYTHONPATH", "").split(os.pathsep) if p])
    new = dict(rec or {"key": key, "path": path, "encoding": encoding, "started": time.time(), "steps": 0})
    new["resumed"] = bool(rec)
    proc = subprocess.Popen([sys.executable, "-m", "todoku.jobs", key], stdin=subprocess.PIPE, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL, env=env, start_new_session=True)
    new["pid"] = proc.pid
    new.pop("stopped", None)
    save(key, new)
    proc.stdin.write(json.dumps({"path": path, "password": password, "encoding": encoding}).encode())
    proc.stdin.close()
    return new


# ---------------------------------------------------------------------- 大きな配列の集計（try_calculation の cells）
def calc_key(path: str, values: dict, cells: list) -> str:
    base = job_key(path, None)
    what = json.dumps({"values": values, "cells": sorted(cells)}, ensure_ascii=False, sort_keys=True, default=str)
    return "calc-" + hashlib.sha256(f"{base}|{what}".encode()).hexdigest()[:24]


def calc_status(path: str, values: dict, cells: list) -> dict | None:
    return load(calc_key(path, values, cells))


def calc_running(rec: dict | None) -> bool:
    return bool(rec) and not rec.get("finished") and _alive(rec.get("pid"))


def start_calc(path: str, password: str | None, values: dict, cells: list, scans: dict | None = None) -> dict:
    """集計の仕事を始める（動いていれば何もしない。止まった仕事は控えの続きから）。scans：ここまでの呼び出しで読み進めた控え。"""
    key = calc_key(path, values, cells)
    rec = load(key)
    if rec and (rec.get("finished") or _alive(rec.get("pid"))):
        return rec
    new = dict(rec or {"key": key, "path": path, "kind": "calc", "values": values, "cells": cells, "started": time.time()})
    merged = dict(new.get("scans") or {})
    for k, v in (scans or {}).items():
        if k not in merged or v.get("k", 0) > merged[k].get("k", 0):
            merged[k] = v
    new["scans"] = merged
    new.pop("stopped", None)
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    env = dict(os.environ)
    env["PYTHONPATH"] = os.pathsep.join([here] + [p for p in env.get("PYTHONPATH", "").split(os.pathsep) if p])
    proc = subprocess.Popen([sys.executable, "-m", "todoku.jobs", key], stdin=subprocess.PIPE, stdout=subprocess.DEVNULL,
                            stderr=subprocess.DEVNULL, env=env, start_new_session=True)
    new["pid"] = proc.pid
    save(key, new)
    proc.stdin.write(json.dumps({"path": path, "password": password}).encode())
    proc.stdin.close()
    return new


def _run_calc(key: str, rec: dict, args: dict) -> int:
    from .formula.lazyarr import ScanStore
    from .xlsx import model as xm

    os.nice(10)
    rec["pid"] = os.getpid()
    threading.Thread(target=_watch_memory, args=(key, rec), daemon=True).start()

    def save_scans(data: dict) -> None:
        rec["scans"] = data
        save(key, rec)

    store = ScanStore(rec.get("scans") or {})
    store.on_save = save_scans
    try:
        b = limits.Budget(total=None)
        b.calc_left = float("inf")
        with limits.call_budget(b):
            rec["results"] = xm.cells_for_job(args["path"], args.get("password"), rec.get("values") or {}, rec["cells"], store)
        rec["scans"] = store.data
        rec["finished"] = time.time()
        save(key, rec)
        return 0
    except Exception as e:  # noqa: BLE001  仕事の失敗は、控えで伝える（次の呼び出しで、続きから）
        rec["scans"] = store.data
        rec["stopped"] = f"途中で問題が起きました（{type(e).__name__}）：{str(e)[:500]}"
        save(key, rec)
        return 1


def _watch_memory(key: str, rec: dict) -> None:
    try:
        import resource
    except ImportError:
        return
    scale = 1 if sys.platform == "darwin" else 1024
    while True:
        if resource.getrusage(resource.RUSAGE_SELF).ru_maxrss * scale > limits.MEMORY_LIMIT:
            rec["stopped"] = (f"メモリを {limits.MEMORY_LIMIT / limits.MB:,.0f} MB より多く使ったので止めました。ここまでの控えは残っています。"
                              "もう一度 audit_document を呼ぶと続きから進みます（TODOKU_MEMORY_LIMIT_MB で上限を変えられます）")
            rec["skip_whole"] = True  # 書類の全体を一度に読む検査が重すぎた：次は分けて読む検査に（行き止まりにしない）
            save(key, rec)
            os._exit(86)
        time.sleep(0.2)


def run(key: str) -> int:
    from . import lazy

    args = json.loads(sys.stdin.read() or "{}")
    rec = load(key) or {"key": key}
    if rec.get("kind") == "calc":
        return _run_calc(key, rec, args)
    os.nice(10)
    rec["pid"] = os.getpid()
    threading.Thread(target=_watch_memory, args=(key, rec), daemon=True).start()
    last = {"t": 0.0}

    def report(text: str, done: int = 0, total: int = 0) -> None:
        rec["progress"] = {"text": text, "done": done, "total": total}
        now = time.monotonic()
        if now - last["t"] > SAVE_EVERY:
            last["t"] = now
            save(key, rec)

    limits.set_progress_reporter(report)
    try:
        with limits.call_budget(limits.Budget(total=None)):
            doc = lazy.open_doc(args["path"], args.get("password"), args.get("encoding"), force=True)
            state = rec.get("state") or doc.audit_begin()
            if rec.get("skip_whole"):
                state["skip_whole"] = True
            while not state.get("done"):
                state = doc.audit_step(state)
                rec["state"] = state
                rec["steps"] = rec.get("steps", 0) + 1
                now = time.monotonic()
                if now - last["t"] > SAVE_EVERY:
                    last["t"] = now
                    save(key, rec)
        rec["finished"] = time.time()
        try:
            import resource

            rec["peak_rss_mb"] = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss // (1 << 20 if sys.platform == "darwin" else 1 << 10)
        except ImportError:
            pass
        rec["progress"] = {"text": "終わりました"}
        save(key, rec)
        return 0
    except Exception as e:  # noqa: BLE001  仕事の失敗は、控えで伝える（次の呼び出しで、続きから）
        rec["stopped"] = f"途中で問題が起きました（{type(e).__name__}）：{str(e)[:500]}"
        save(key, rec)
        return 1


if __name__ == "__main__":
    sys.exit(run(sys.argv[1]))
