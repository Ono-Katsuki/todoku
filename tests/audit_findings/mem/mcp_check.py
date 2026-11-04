"""監査用：同じ書類を MCP サーバーの道（既定の isolate）で動かし、止まること・サーバーが生きていることを確かめる。

使い方：mcp_check.py 書類 [書類 ...]（道具名=open_document 既定。環境 MCP_TOOL、view_page は page=1）
プロセスの木の常駐を外から見張る（念のため 4 GB を超えたら全部を殺す）。
"""
import fcntl
import json
import os
import select
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.abspath(os.path.join(HERE, "..", "..", "..", "src"))
LOCK = os.path.join(tempfile.gettempdir(), "todoku_mem_guard.lock")


def tree_rss(root):
    out = subprocess.run(["ps", "-A", "-o", "pid=,ppid=,rss="], capture_output=True, text=True).stdout
    kids = {}
    rss = {}
    for line in out.splitlines():
        p, pp, r = line.split()
        kids.setdefault(int(pp), []).append(int(p))
        rss[int(p)] = int(r)
    total = {}
    st = [root]
    while st:
        p = st.pop()
        total[p] = rss.get(p, 0)
        st += kids.get(p, [])
    return total


def send(p, obj):
    p.stdin.write((json.dumps(obj) + "\n").encode())
    p.stdin.flush()


def recv(p, timeout, watch):
    end = time.time() + timeout
    peak = 0
    buf = b""
    while time.time() < end:
        r = tree_rss(p.pid)
        s = sum(r.values())
        peak = max(peak, s)
        if s > 4 * 1048576:
            for q in r:
                try:
                    os.kill(q, 9)
                except OSError:
                    pass
            return None, peak, "TREE>4GB killed by the check"
        rd, _, _ = select.select([p.stdout], [], [], 0.05)
        if rd:
            chunk = os.read(p.stdout.fileno(), 1 << 20)
            if not chunk:
                return None, peak, "server closed"
            buf += chunk
            if buf.endswith(b"\n"):
                return json.loads(buf.splitlines()[-1]), peak, ""
    return None, peak, "no answer"


def main():
    env = dict(os.environ, PYTHONPATH=SRC)
    env.pop("TODOKU_ISOLATE", None)
    tool = os.environ.get("MCP_TOOL", "open_document")
    with open(LOCK, "w") as lk:
        fcntl.flock(lk, fcntl.LOCK_EX)
        p = subprocess.Popen(["nice", "-n", "10", sys.executable, "-m", "todoku.mcp_server"], env=env, stdin=subprocess.PIPE,
                             stdout=subprocess.PIPE, stderr=subprocess.DEVNULL)
        send(p, {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2024-11-05", "capabilities": {}, "clientInfo": {"name": "audit", "version": "0"}}})
        recv(p, 30, None)
        send(p, {"jsonrpc": "2.0", "method": "notifications/initialized"})
        for i, path in enumerate(sys.argv[1:]):
            t0 = time.time()
            args = {"path": os.path.abspath(path)}
            if tool == "view_page":
                args["page"] = 1
            send(p, {"jsonrpc": "2.0", "id": 10 + i, "method": "tools/call", "params": {"name": tool, "arguments": args}})
            r, peak, why = recv(p, 400, None)
            dt = time.time() - t0
            text = ""
            if r:
                c = r.get("result", {}).get("content") or []
                text = (c[0].get("text") or "")[:110] if c else json.dumps(r)[:110]
            print(f"{os.path.basename(path)}: {dt:.1f}s tree_peak={peak // 1024}MB {why} isError={r and r.get('result', {}).get('isError')} {text}", flush=True)
            if why:
                break
            # サーバーが生きているか
            send(p, {"jsonrpc": "2.0", "id": 1000 + i, "method": "tools/list"})
            r2, _pk, why2 = recv(p, 30, None)
            print("   server alive:", bool(r2) and "tools" in r2.get("result", {}), why2, flush=True)
        p.kill()


if __name__ == "__main__":
    main()
