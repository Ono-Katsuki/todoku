"""道具の呼び出しを、別のプロセス（働き手）で動かす：壊れた・悪意のある書類で、MCP サーバーごと止まらないように。

- 働き手は、サーバーと同じ Python で動く子のプロセス 1 つ（はじめに立ち上げておき、呼び出しのたびには作らない。
  窓の控え（api._WINDOWS）も働き手の中に残るので、同じ書類を開いてから書くときは速いまま）。
- 1 回の呼び出しに、時間（limits.TIME_LIMIT 秒）とメモリ（limits.MEMORY_LIMIT）の上限を置く。超えたら働き手を止めて、
  エージェントには「この書類は大きすぎるか、壊れているため止めました：…」と返す。サーバーは動き続け、次の呼び出しのために
  働き手を立ち上げ直す。
- メモリ：Linux では働き手に RLIMIT_AS を置く（超えると MemoryError）。macOS は RLIMIT_AS・RLIMIT_DATA を守らない
  （setrlimit が失敗する）ので、働き手の中で最大の常駐メモリ（ru_maxrss）を見張り、サーバーの側でも ps で見張る。
- やりとりは、パイプの上の長さつきの JSON（結果は MCP にそのまま返すもので、もともと JSON にする）。
- TODOKU_ISOLATE=0 なら、別のプロセスにせず、サーバーの中で動かす（上限は、展開・描画の予算だけになる）。
"""

from __future__ import annotations

import json
import os
import select
import struct
import subprocess
import sys
import threading
import time

from . import limits

EXIT_MEMORY = 86  # 働き手が、メモリの上限を超えたと自分で気づいて終わったときの終了コード
EXIT_TIME = 87  # 働き手が、時間の上限を超えたと自分で気づいて終わったとき（サーバーの見張りが先に止めるのがふつう）


def enabled() -> bool:
    return os.environ.get("TODOKU_ISOLATE", "1") not in ("0", "false", "no", "")


# ---------------------------------------------------------------------- やりとり
def _send(fd: int, obj) -> None:
    data = json.dumps(obj, ensure_ascii=False, default=str).encode("utf-8")
    buf = struct.pack(">I", len(data)) + data
    while buf:
        n = os.write(fd, buf)
        buf = buf[n:]


def _recv_exact(fd: int, n: int) -> bytes | None:
    out = bytearray()
    while len(out) < n:
        b = os.read(fd, n - len(out))
        if not b:
            return None
        out += b
    return bytes(out)


def _recv(fd: int):
    head = _recv_exact(fd, 4)
    if head is None:
        return None
    body = _recv_exact(fd, struct.unpack(">I", head)[0])
    return None if body is None else json.loads(body.decode("utf-8"))


def _rss_kb(pid: int) -> int:
    try:
        out = subprocess.run(["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True, timeout=5).stdout
        return int(out.strip() or 0)
    except (OSError, ValueError, subprocess.SubprocessError):
        return 0


# ---------------------------------------------------------------------- サーバーの側
class Worker:
    def __init__(self):
        self.proc: subprocess.Popen | None = None
        self.req = self.res = -1
        self.lock = threading.Lock()

    def start(self) -> None:
        if self.proc is not None and self.proc.poll() is None:
            return
        self._close()
        r_req, w_req = os.pipe()
        r_res, w_res = os.pipe()
        here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # todoku のパッケージを含むフォルダ
        env = dict(os.environ)
        env["PYTHONPATH"] = os.pathsep.join([here] + [p for p in env.get("PYTHONPATH", "").split(os.pathsep) if p])
        self.proc = subprocess.Popen([sys.executable, "-m", "todoku.isolate", str(r_req), str(w_res)],
                                     pass_fds=(r_req, w_res), stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                     env=env, close_fds=True)
        os.close(r_req)
        os.close(w_res)
        self.req, self.res = w_req, r_res

    def _close(self) -> None:
        for fd in (self.req, self.res):
            if fd >= 0:
                try:
                    os.close(fd)
                except OSError:
                    pass
        self.req = self.res = -1

    def stop(self) -> None:
        if self.proc is not None and self.proc.poll() is None:
            self.proc.kill()
            try:
                self.proc.wait(5)
            except subprocess.TimeoutExpired:
                pass
        self.proc = None
        self._close()

    def call(self, name: str, args: dict) -> dict:
        """道具を働き手で動かして、MCP の結果（content・isError など）を返す。上限を超えたら、止めた理由の結果。"""
        with self.lock:
            self.start()
            try:
                _send(self.req, {"name": name, "args": args})
            except OSError:  # 働き手が、前の呼び出しの後に終わっていた：立ち上げ直して、もう一度
                self.stop()
                self.start()
                _send(self.req, {"name": name, "args": args})
            t0 = time.monotonic()
            next_ps = t0 + 1.0
            why = None
            while True:
                ready, _w, _x = select.select([self.res], [], [], 0.2)
                if ready:
                    msg = _recv(self.res)
                    if msg is not None:
                        return msg["result"]
                    try:
                        code = self.proc.wait(5) if self.proc else None
                    except subprocess.TimeoutExpired:
                        code = None
                    why = (_MEMORY() if code == EXIT_MEMORY else _TIME() if code == EXIT_TIME else
                           f"読む処理が途中で落ちました（終了コード {code}。入れ子が深すぎる・メモリが足りないなど）")
                    break
                now = time.monotonic()
                if now - t0 > limits.TIME_LIMIT:
                    why = _TIME()
                    break
                if now >= next_ps:  # 長い呼び出しだけ、外からもメモリを見張る（macOS は RLIMIT_AS を守らないため）
                    next_ps = now + 1.0
                    if self.proc is not None and _rss_kb(self.proc.pid) * 1024 > limits.MEMORY_LIMIT:
                        why = _MEMORY()
                        break
            self.stop()
            self.start()  # 次の呼び出しのために、すぐ立ち上げ直す（立ち上がるのを待たない）
            return stopped(why)


def _MEMORY() -> str:  # noqa: N802
    return f"メモリを {limits.MEMORY_LIMIT / limits.MB:,.0f} MB より多く使いました（環境変数 TODOKU_MEMORY_LIMIT_MB で変えられます）"


def _TIME() -> str:  # noqa: N802
    return f"{limits.TIME_LIMIT} 秒を超えても終わりませんでした（環境変数 TODOKU_TIME_LIMIT で変えられます）"


def stopped(why: str) -> dict:
    return {"content": [{"type": "text", "text": f"{limits.STOPPED}：{why}"}], "isError": True,
            "structuredContent": {"stopped": True, "reason": why}}


_WORKER: Worker | None = None


def worker() -> Worker:
    global _WORKER
    if _WORKER is None:
        _WORKER = Worker()
    return _WORKER


# ---------------------------------------------------------------------- 働き手の側
_BUSY = {"since": None}  # いま動いている呼び出しの始まり（time.monotonic）


def _watchdog() -> None:
    """働き手の中の見張り：メモリ（Linux は RLIMIT_AS も）・1 回の呼び出しの時間・親（サーバー）がいること。
    サーバーの側でも見張るが、サーバーが先に終わったとき（親がいなくなった）にも、働き手が動き続けないように。"""
    import resource

    cap = limits.MEMORY_LIMIT
    if sys.platform != "darwin" and hasattr(resource, "RLIMIT_AS"):  # macOS は RLIMIT_AS を守らない
        try:
            hard = resource.getrlimit(resource.RLIMIT_AS)[1]
            want = cap + 512 * limits.MB  # 読み込んだライブラリの分の余裕
            resource.setrlimit(resource.RLIMIT_AS, (want if hard == resource.RLIM_INFINITY else min(want, hard), hard))
        except (ValueError, OSError):
            pass
    parent = os.getppid()
    scale = 1 if sys.platform == "darwin" else 1024  # ru_maxrss：macOS はバイト、Linux は KB

    def watch():
        while True:
            if resource.getrusage(resource.RUSAGE_SELF).ru_maxrss * scale > cap:
                os._exit(EXIT_MEMORY)
            if os.getppid() != parent:
                os._exit(0)
            since = _BUSY["since"]
            if since is not None and time.monotonic() - since > limits.TIME_LIMIT + 10:
                os._exit(EXIT_TIME)
            time.sleep(0.05)

    threading.Thread(target=watch, daemon=True).start()


def serve(req: int, res: int) -> int:
    _watchdog()
    from . import mcp_server  # 先に読み込んでおく（はじめの呼び出しを速く）

    while True:
        msg = _recv(req)
        if msg is None:
            return 0
        _BUSY["since"] = time.monotonic()
        r = mcp_server.tool_result(msg["name"], msg.get("args") or {})
        _BUSY["since"] = None
        _send(res, {"result": r})


if __name__ == "__main__":
    sys.exit(serve(int(sys.argv[1]), int(sys.argv[2])))
