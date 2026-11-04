"""道具の呼び出しを、別のプロセス（働き手）で動かす：壊れた・悪意のある書類で、MCP サーバーごと止まらないように。

- 働き手は、サーバーと同じ Python で動く子のプロセス 1 つ（はじめに立ち上げておき、呼び出しのたびには作らない。
  窓の控え（api._WINDOWS）も働き手の中に残るので、同じ書類を開いてから書くときは速いまま）。
- 1 回の呼び出しに、時間（limits.TIME_LIMIT 秒）とメモリ（limits.MEMORY_LIMIT）の上限を置く。超えたら働き手を止めて、
  エージェントには「この書類は大きすぎるか、壊れているため止めました：…」と返す。サーバーは動き続け、次の呼び出しのために
  働き手を立ち上げ直す。
- メモリ：Linux では働き手に RLIMIT_AS と RLIMIT_DATA を置く（超えると MemoryError）。macOS はこれらを守らないので、働き手の中で
  最大の常駐メモリ（ru_maxrss）を見張り、サーバーの側でも 0.05 秒ごとに見張る（rss_bytes。C の中の 1 回の大きな確保は途中で
  止められないので、宣言の数から確保する前に、実際のデータの長さや予算と比べる：limits.Budget.count・alloc）。
- やりとりは、パイプの上の長さつきの JSON（結果は MCP にそのまま返すもので、もともと JSON にする）。
- TODOKU_ISOLATE=0 なら、別のプロセスにせず、サーバーの中で動かす（上限は、展開・描画の予算だけになる）。
- Windows（os.name == "nt"）では、パイプの受け渡し（pass_fds）・パイプの select・resource が使えないので、別のプロセスに
  せず、サーバーの中で動かす（標準エラーに一度知らせ、サーバーの説明 notice() にも書く）。このときは同じ呼び出しの予算に加えて、
  見張りのスレッド（soft_call）が時間とメモリを見て、Python の処理の切れ目に LimitExceeded を起こす（やわらかい止め方：
  C の中の 1 回の処理の途中では止まらない。TODOKU_ISOLATE=0 のときも同じ）。コマンドライン（cli）は api の同じ予算で動く。
- ps が使えないとき（外からのメモリの見張りができない）も、サーバーの立ち上げで一度確かめて（probe_ps）、標準エラーに一度
  知らせ、notice() に書く（黙って見張りを止めない）。
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
POLL = 0.05  # サーバーが働き手を見る間隔（秒）


_TOLD: set[str] = set()  # 標準エラーに一度だけ知らせたこと


def _tell_once(key: str, text: str) -> None:
    if key not in _TOLD:
        _TOLD.add(key)
        print(f"todoku: {text}", file=sys.stderr, flush=True)


_NO_ISOLATE_NT = ("Windows では、道具の呼び出しを別のプロセス（働き手）に分けられないので、サーバーの中で動かします。"
                  "壊れた書類には、展開・描画・項目の予算と、見張りのスレッドの時間とメモリの上限（Python の処理の切れ目で止める。"
                  "C の中の 1 回の処理の途中では止まらない）で備えます")
_NO_PS = ("ps が使えないので、働き手のメモリを外から見張れません（働き手の中の見張りだけになります。macOS では、"
          "メモリの上限を超えても止められないことがあります）" if sys.platform == "darwin" else
          "ps が使えないので、働き手の常駐メモリ（RSS）の外からの見張りは止まります（メモリの上限 RLIMIT_AS と、"
          "働き手の中の見張りは、そのまま効きます）")


def enabled() -> bool:
    if os.environ.get("TODOKU_ISOLATE", "1") in ("0", "false", "no", ""):
        return False
    if os.name == "nt":
        _tell_once("nt", _NO_ISOLATE_NT)
        return False
    return True


_PS_OK: bool | None = None  # ps で働き手のメモリを読めるか（サーバーの立ち上げで一度だけ確かめる）


def probe_ps() -> bool:
    """ps でプロセスの常駐メモリを読めるかを、このプロセス自身で一度だけ確かめる（読めなければ標準エラーに知らせる）。
    呼び出しの途中で初めて分かるのでなく、サーバーの説明（notice）に立ち上げの時から書けるように。"""
    global _PS_OK
    if _PS_OK is None:
        _PS_OK = _rss_kb(os.getpid()) > 0
        if not _PS_OK:
            _tell_once("ps", _NO_PS)
    return _PS_OK


def notice() -> str:
    """サーバーの説明に足す、上限が効かないことの知らせ（なければ空）。"""
    out = []
    on = os.environ.get("TODOKU_ISOLATE", "1") not in ("0", "false", "no", "")
    if os.name == "nt" and on:
        out.append(_NO_ISOLATE_NT + "。")
    if on and os.name != "nt":
        probe_ps()  # 働き手を外から見張る道具（ps）が使えるか：立ち上げの時に確かめる
    if "ps" in _TOLD:
        out.append(_NO_PS + "。")
    return "".join(out)


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
    except (OSError, ValueError, subprocess.SubprocessError) as e:
        _tell_once("ps", f"{_NO_PS}（{type(e).__name__}）")  # 0 を返すと見張りが黙って止まるので、知らせる
        return 0


def _darwin_rss(pid: int) -> int | None:
    """macOS：proc_pid_rusage（ps を立ち上げない）。常駐と、圧縮された分も含む使用量（phys_footprint）の大きい方。"""
    import ctypes

    class Info(ctypes.Structure):  # rusage_info_v0（sys/resource.h）
        _fields_ = [("uuid", ctypes.c_uint8 * 16), ("user", ctypes.c_uint64), ("system", ctypes.c_uint64),
                    ("wakeups", ctypes.c_uint64), ("interrupt", ctypes.c_uint64), ("pageins", ctypes.c_uint64),
                    ("wired", ctypes.c_uint64), ("resident", ctypes.c_uint64), ("footprint", ctypes.c_uint64),
                    ("start", ctypes.c_uint64), ("exit", ctypes.c_uint64)]
    info = Info()
    if _libc().proc_pid_rusage(pid, 0, ctypes.byref(info)) != 0:
        return None
    return max(info.resident, info.footprint)


def _libc():
    import ctypes

    global _LIBC
    if _LIBC is None:
        _LIBC = ctypes.CDLL("libSystem.dylib" if sys.platform == "darwin" else None)
    return _LIBC


_LIBC = None


def _windows_rss() -> int | None:
    """Windows：このプロセスの常駐（GetProcessMemoryInfo）。"""
    import ctypes
    from ctypes import wintypes

    class Counters(ctypes.Structure):
        _fields_ = [("cb", wintypes.DWORD), ("faults", wintypes.DWORD), ("peak", ctypes.c_size_t), ("working", ctypes.c_size_t),
                    ("a", ctypes.c_size_t), ("b", ctypes.c_size_t), ("c", ctypes.c_size_t), ("d", ctypes.c_size_t),
                    ("pagefile", ctypes.c_size_t), ("peak_pagefile", ctypes.c_size_t)]
    c = Counters()
    c.cb = ctypes.sizeof(c)
    psapi, k32 = ctypes.WinDLL("psapi"), ctypes.WinDLL("kernel32")
    k32.GetCurrentProcess.restype = wintypes.HANDLE
    if not psapi.GetProcessMemoryInfo(k32.GetCurrentProcess(), ctypes.byref(c), c.cb):
        return None
    return c.working


def rss_bytes(pid: int | None = None) -> int:
    """プロセスの常駐メモリ（バイト）。ps を立ち上げずに読める道（/proc・proc_pid_rusage・GetProcessMemoryInfo）を先に使う
    （短い間隔で見張るため）。読めなければ ps、それもだめなら 0。"""
    pid = os.getpid() if pid is None else pid
    try:
        if sys.platform == "darwin":
            v = _darwin_rss(pid)
            if v:
                return v
        elif os.name == "nt":
            if pid == os.getpid():
                return _windows_rss() or 0
            return 0
        else:
            with open(f"/proc/{pid}/statm", "rb") as f:
                return int(f.read().split()[1]) * os.sysconf("SC_PAGE_SIZE")
    except (OSError, ValueError, IndexError, AttributeError):
        pass
    return _rss_kb(pid) * 1024


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
            why = None
            at = None
            while True:
                ready, _w, _x = select.select([self.res], [], [], POLL)
                if ready:
                    msg = _recv(self.res)
                    if msg is not None and "progress" in msg:
                        at = msg["progress"]  # 働き手の進み具合（止まったときに、どこまで進んだかを言う）
                        continue
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
                # 外からもメモリを見張る（macOS は RLIMIT_AS を守らない。働き手の中の見張りは、C の中の確保の間は動けない）
                if self.proc is not None and rss_bytes(self.proc.pid) > limits.MEMORY_LIMIT:
                    why = _MEMORY()
                    break
            self.stop()
            self.start()  # 次の呼び出しのために、すぐ立ち上げ直す（立ち上がるのを待たない）
            return stopped(why, name, args, at)


def _MEMORY() -> str:  # noqa: N802
    return f"メモリを {limits.MEMORY_LIMIT / limits.MB:,.0f} MB より多く使いました（環境変数 TODOKU_MEMORY_LIMIT_MB で変えられます）"


def _TIME() -> str:  # noqa: N802
    return f"{limits.TIME_LIMIT} 秒を超えても終わりませんでした（環境変数 TODOKU_TIME_LIMIT で変えられます）"


_WINDOW_TOOLS = ("open_document", "read_window", "list_fields")  # 窓を作るところで時間がかかる道具
_RANGE_ARGS = ("pages", "sheet", "rows", "paragraphs", "records", "path_in", "cursor", "raw", "part")


def _window_args(name: str, args: dict | None) -> bool:
    """窓の全体を作る呼び出しか（範囲を頼んだ呼び出しは、範囲を小さくできる。パスワードは控えに残さないので、長い仕事にしない）。"""
    args = args or {}
    return name in _WINDOW_TOOLS and bool(args.get("path")) and not args.get("password") and not any(args.get(k) for k in _RANGE_ARGS)


def _job_result(rec: dict) -> dict:
    """窓を作っている長い仕事の進み具合（失敗ではない）。"""
    pr = rec.get("progress") or {}
    at = pr.get("text") or "始めたところ"
    if pr.get("total"):
        at += f"（{pr['done']:,}／{pr['total']:,}）"
    text = (f"窓を作るのに時間がかかるので、別のプロセスの長い仕事にしました（いま：{at}）。読めないのではありません。"
            "もう一度同じ呼び出しをすると、進み具合か結果が返ります")
    return {"content": [{"type": "text", "text": text}], "isError": False,
            "structuredContent": {"job": {"running": True, "progress": pr, "key": rec.get("key")}}}


def window_job_view(name: str, args: dict) -> dict | None:
    """同じ書類の窓を長い仕事で作っている間は、働き手を待たせず進み具合を返す。動いていなければ None。"""
    if not _window_args(name, args):
        return None
    from . import jobs

    try:
        rec = jobs.window_running(args["path"], args.get("encoding"))
    except OSError:
        return None
    return _job_result(rec) if rec is not None else None


def stopped(why: str, name: str = "", args: dict | None = None, at: dict | None = None) -> dict:
    """止めた理由と、どこまで進んだか、代わりに頼める小さな頼み方（行き止まりにしない）。"""
    from .lazy import smaller_request

    if why == _TIME() and _window_args(name, args):
        from . import jobs

        try:
            failed = jobs.window_failed(args["path"], args.get("encoding"))
            if failed:
                why = f"{why}。長い仕事でも止まりました：{failed}"
            elif not jobs.window_made(args["path"], args.get("encoding")):
                return _job_result(jobs.start_window(args["path"], args.get("encoding")))
        except (OSError, ValueError):
            pass  # 仕事を始められない：いままでどおり、止めた理由と頼み方を伝える
    hint = smaller_request(name, args or {})
    where = ""
    if at and at.get("text"):
        where = f"（{at['text']}" + (f"：{at['done']:,}／{at['total']:,}" if at.get("total") else "") + " まで進みました）"
    return {"content": [{"type": "text", "text": f"{limits.STOPPED}：{why}{where}\n{hint}"}], "isError": True,
            "structuredContent": {"stopped": True, "reason": why, "smaller_request": hint, **({"progress": at} if at else {})}}


class _SoftStop(limits.LimitExceeded):
    """見張りのスレッドが、動いているスレッドに起こす例外（引数なしで作れること：PyThreadState_SetAsyncExc）。"""

    reason = ""

    def __init__(self):
        super().__init__(_SoftStop.reason or "時間かメモリの上限を超えました", True)


def soft_call(fn, name: str, args: dict) -> dict:
    """別のプロセスにできないとき（Windows・TODOKU_ISOLATE=0）の道具の呼び出し：見張りのスレッドが時間とメモリを見て、Python の
    処理の切れ目に _SoftStop を起こす。C の中の 1 回の処理の途中では止まらない（そのために、予算と、確保の前の確かめがある）。"""
    import ctypes

    done = threading.Event()
    target = threading.get_ident()
    t0 = time.monotonic()

    def watch():
        while not done.wait(POLL * 2):
            why = None
            if time.monotonic() - t0 > limits.TIME_LIMIT:
                why = _TIME()
            elif rss_bytes() > limits.MEMORY_LIMIT / 2 and (limits.release_memory() or rss_bytes() > limits.MEMORY_LIMIT):
                why = _MEMORY()
            if why:
                _SoftStop.reason = why
                ctypes.pythonapi.PyThreadState_SetAsyncExc(ctypes.c_ulong(target), ctypes.py_object(_SoftStop))
                return

    w = threading.Thread(target=watch, daemon=True)
    w.start()
    try:
        return fn(name, args)
    except _SoftStop as e:  # fn の外で起きた（fn の中なら fn が LimitExceeded として結果にする）
        return stopped(e.what, name, args)
    finally:
        done.set()
        w.join(1)


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
    try:
        import resource
    except ImportError:  # resource のない環境（Windows）：働き手の中のメモリの見張りはできない
        resource = None
        _tell_once("resource", "resource が使えないので、働き手の中のメモリの見張りはしません")
    cap = limits.MEMORY_LIMIT
    if resource is not None and sys.platform != "darwin":  # macOS は守らない。Linux は超えると MemoryError（きれいに止まる）
        for name in ("RLIMIT_AS", "RLIMIT_DATA"):  # RLIMIT_DATA：Linux 4.7 から、匿名の大きな確保も数える（ファイルの写像は数えない）
            if not hasattr(resource, name):
                continue
            try:
                what = getattr(resource, name)
                hard = resource.getrlimit(what)[1]
                want = cap + 512 * limits.MB  # 読み込んだライブラリの分の余裕
                resource.setrlimit(what, (want if hard == resource.RLIM_INFINITY else min(want, hard), hard))
            except (ValueError, OSError):
                pass
    parent = os.getppid()

    def watch():
        next_release = 0.0
        while True:
            used = rss_bytes()
            if used > cap / 2 and time.monotonic() >= next_release:  # 使い終えた大きなメモリを返してから数え直す
                next_release = time.monotonic() + 0.5
                limits.release_memory()
                used = rss_bytes()
            if used > cap:
                os._exit(EXIT_MEMORY)
            if os.getppid() != parent:
                os._exit(0)
            since = _BUSY["since"]
            if since is not None and time.monotonic() - since > limits.TIME_LIMIT + 10:
                os._exit(EXIT_TIME)
            time.sleep(POLL)

    threading.Thread(target=watch, daemon=True).start()


def serve(req: int, res: int) -> int:
    _watchdog()
    from . import mcp_server  # 先に読み込んでおく（はじめの呼び出しを速く）

    last = {"t": 0.0}

    def report(text: str, done: int = 0, total: int = 0) -> None:  # 1 秒に 1 回まで、サーバーに進み具合を伝える
        now = time.monotonic()
        if now - last["t"] >= 1.0:
            last["t"] = now
            _send(res, {"progress": {"text": text, "done": done, "total": total}})

    limits.set_progress_reporter(report)

    while True:
        msg = _recv(req)
        if msg is None:
            return 0
        _BUSY["since"] = time.monotonic()
        r = mcp_server.tool_result(msg["name"], msg.get("args") or {})
        _BUSY["since"] = None
        limits.release_memory()  # 次の呼び出しまで、大きなメモリを持ち続けない
        _send(res, {"result": r})


if __name__ == "__main__":
    sys.exit(serve(int(sys.argv[1]), int(sys.argv[2])))
