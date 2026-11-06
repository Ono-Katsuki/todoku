"""長い仕事（大きな書類の検査・大きな配列の集計）：1 回の道具の呼び出しの時間（limits.TIME_LIMIT）に縛られない、別のプロセスの仕事。

- audit_document を呼ぶと、仕事を始めて（すでに動いていれば、その進み具合を）すぐ返す。終わるまで何度でも呼べる
- 区切りごとに控え（state.json）を書くので、止まっても（PC を閉じた・メモリの上限など）次の呼び出しで続きから進む
- 控えの場所：TODOKU_JOBS_DIR（なければ ~/.cache/todoku/jobs）。ファイルの中身のうち、控えに入るのは検査で見つけた問題の短い字と、集計の途中の値・結果だけ。
  パスワードは控えに書かない（標準入力で仕事に渡す。止まった仕事を続けるときは、もう一度 password を渡す）
- 同じファイル（パス・大きさ・更新時刻・文字コード）の仕事は 1 つだけ
- 窓を作る仕事（start_window）：open_document などが時間の上限に当たったとき。できた窓は控えの場所に置き（.window。自分だけが読める。
  書類の中身が入るので、大きさの合計を limits.WINDOW_CACHE までにする）、次の呼び出しで使う。パスワードのある書類ではしない
- 控えの消し方（申請の書類は個人情報なので、置きっぱなしにしない）：TODOKU_JOBS_TTL_HOURS（24 時間）より古い控えは、窓も記録も
  字を探した控えも消す（purge：この場所を使うときと MCP サーバーの立ち上げで）。元のファイルがなくなった仕事の控えも消す。動いている仕事は消さない。
  0 にすると、窓は読んだら、仕事の記録は終わった結果を返したら消す（読まれなかった控えは 1 時間で。大きな書類は開くたびに作り直す）。
  `todoku forget` で、動いている仕事のほかをすぐに消せる
- 仕事が落ちたとき：標準エラーを控えの場所に置き（.err。最後の 1 行を理由として伝える）、進まないまま MAX_TRIES 回止まったら、もう立ち上げ直さず
  失敗として返す（failed。「動いている」と返し続けない）。進んでいれば、何度止まっても続きから進む
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
MAX_TRIES = 3  # 進まないまま立ち上げる回数の上限


def _dir() -> Path:
    d = Path(os.environ.get("TODOKU_JOBS_DIR") or Path.home() / ".cache" / "todoku" / "jobs")
    d.mkdir(parents=True, exist_ok=True, mode=0o700)
    return d


DEFAULT_TTL_HOURS = 24.0
UNREAD_TTL = 3600.0  # 読んだら消す設定（0）で、読まれなかった控えを残す時間（秒）
PURGE_EVERY = 60.0  # 片づける間隔（秒）。同じプロセスで何度も一覧を読まない
_purged: dict[str, float] = {}


def jobs_dir() -> Path:
    d = _dir()
    now = time.monotonic()
    if now - _purged.get(str(d), -PURGE_EVERY) >= PURGE_EVERY:
        _purged[str(d)] = now
        purge()
    return d


def ttl_seconds() -> float:
    """控えを残す時間（秒）。0 は読んだら消す。読めない値は既定（消す側の既定に倒す）。"""
    try:
        h = float(os.environ.get("TODOKU_JOBS_TTL_HOURS", "") or DEFAULT_TTL_HOURS)
    except ValueError:
        h = DEFAULT_TTL_HOURS
    return max(h, 0.0) * 3600


def purge(now: float | None = None, everything: bool = False) -> list[str]:
    """古い控えを消す（消した名前を返す）。残す時間は最後に書いた時刻から。動いている仕事の控えは消さない。
    元のファイルがなくなった仕事の控えも消す（消したファイルの中身を控えに残さない）。everything：動いている仕事のほかは、すべて。"""
    now = time.time() if now is None else now
    ttl = ttl_seconds() or UNREAD_TTL
    gone: list[str] = []
    d = _dir()
    try:
        names = list(d.iterdir())
    except OSError:
        return gone
    running = set()
    source_gone = set()
    for p in names:
        if p.suffix != ".json" or p.name.startswith("anchor-"):
            continue
        try:
            rec = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            continue
        if not rec.get("finished") and not rec.get("stopped") and _alive(rec.get("pid")):
            running.add(p.stem)
        elif rec.get("path") and not Path(str(rec["path"]).split("!/")[0]).expanduser().exists():
            source_gone.add(p.stem)
    for p in names:
        stem = p.name.split(".", 1)[0]
        if stem in running:
            continue
        try:
            old = now - p.stat().st_mtime >= ttl
        except OSError:
            continue
        if everything or old or stem in source_gone:
            try:
                p.unlink(missing_ok=True)
                gone.append(p.name)
            except OSError:
                pass
    return gone


def read_once(rec: dict | None) -> None:
    """読んだら消す設定（TODOKU_JOBS_TTL_HOURS=0）：終わった仕事の結果を返したので、その記録を消す。"""
    if rec and rec.get("finished") and rec.get("key") and ttl_seconds() <= 0:
        _state_path(rec["key"]).unlink(missing_ok=True)
        _err_path(rec["key"]).unlink(missing_ok=True)


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


def _lower_priority() -> None:
    """ほかの処理を邪魔しないように優先度を下げる。下げられなくても（macOS は優先度がすでに低いと PermissionError）仕事は進める。"""
    try:
        os.nice(10)
    except OSError:
        pass


def _spawn(key: str) -> subprocess.Popen:
    """仕事のプロセスを立ち上げる。標準エラーは控えの場所へ（落ちた理由を伝えるため）。終わったらすぐ片づける：立ち上げたプロセスが生きている間
    ゾンビとして残ると、ほかのプロセスからは「動いている」と見える。"""
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    env = dict(os.environ)
    env["PYTHONPATH"] = os.pathsep.join([here] + [p for p in env.get("PYTHONPATH", "").split(os.pathsep) if p])
    fd = os.open(_err_path(key), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
    try:
        proc = subprocess.Popen([sys.executable, "-m", "todoku.jobs", key], stdin=subprocess.PIPE, stdout=subprocess.DEVNULL,
                                stderr=fd, env=env, start_new_session=True)
    finally:
        os.close(fd)
    threading.Thread(target=proc.wait, daemon=True).start()
    return proc


def _send(proc: subprocess.Popen, payload: dict) -> None:
    """仕事に頼む中身を標準入力で渡す。もう終わっていても（落ちた仕事は、次の呼び出しで失敗として分かる）ここでは失敗にしない。"""
    try:
        proc.stdin.write(json.dumps(payload).encode())
        proc.stdin.close()
    except OSError:
        pass


def _err_path(key: str) -> Path:
    return jobs_dir() / f"{key}.err"


def _last_error(key: str) -> str:
    """落ちた仕事の標準エラーの最後の行（例外の種類と理由）。"""
    try:
        with open(_err_path(key), "rb") as f:
            f.seek(0, os.SEEK_END)
            f.seek(max(f.tell() - 4096, 0))
            lines = f.read().decode("utf-8", "replace").splitlines()
    except OSError:
        return ""
    lines = [x.strip() for x in lines if x.strip()]
    return lines[-1][:500] if lines else ""


def _gave_up(key: str, rec: dict, done: int) -> bool:
    """止まった仕事を立ち上げ直す前に数える。done（進んだ量）が増えないまま MAX_TRIES 回止まったら、失敗として控えに書く（True：立ち上げ直さない）。"""
    if rec.get("tries_at") != done:
        rec["tries_at"] = done
        rec["tries"] = 0
        rec.pop("failed", None)
    rec["tries"] = rec.get("tries", 0) + 1
    if rec["tries"] < MAX_TRIES:
        return False
    if not rec.get("failed"):
        err = _last_error(key)
        rec["stopped"] = rec.get("stopped") or (f"仕事のプロセスが異常終了しました：{err}" if err else "仕事のプロセスが、理由を残さずに終わりました")
        rec["failed"] = True
    save(key, rec)
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
    if rec and not restart:
        # 仕事は止まっていた。読んだ後に終わったのかもしれない：古い控えで終わった結果を上書きして、やり直さない
        rec = load(key) or rec
        if rec.get("finished") or _gave_up(key, rec, rec.get("steps", 0)):
            return rec
    if restart:
        rec = None
        _state_path(key).unlink(missing_ok=True)
    new = dict(rec or {"key": key, "path": path, "encoding": encoding, "started": time.time(), "steps": 0})
    new["resumed"] = bool(rec)
    proc = _spawn(key)
    new["pid"] = proc.pid
    new.pop("stopped", None)
    save(key, new)
    _send(proc, {"path": path, "password": password, "encoding": encoding})
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
    new = dict(rec or {"key": key, "path": path, "kind": "calc", "cells": cells, "started": time.time()})  # 書く値は控えに残さない（標準入力で渡す）
    merged = dict(new.get("scans") or {})
    for k, v in (scans or {}).items():
        if k not in merged or v.get("k", 0) > merged[k].get("k", 0):
            merged[k] = v
    new["scans"] = merged
    if rec and _gave_up(key, new, sum(v.get("k", 0) for v in merged.values())):
        return new
    new.pop("stopped", None)
    proc = _spawn(key)
    new["pid"] = proc.pid
    save(key, new)
    _send(proc, {"path": path, "password": password, "values": values})
    return new


# ---------------------------------------------------------------------- 書いた後のファイルの、ブック全体の計算し直し（fill_document が未確認で返したもの）
def _proof_key(path: str) -> str:
    return "proof-" + job_key(path, None)


def note_proof(path: str, original: str, seeds: list, changed: list, marked: bool) -> None:
    """未確認で返したファイルと、確かめに要ること（セルの番地だけ。書いた値は控えに残さない）を控える。仕事は audit_document で始める。"""
    key = _proof_key(path)
    save(key, {"key": key, "kind": "proof", "path": path, "original": original, "seeds": [list(x) for x in seeds],
               "changed": [list(x) for x in changed], "marked": marked})


def proof_status(path: str) -> dict | None:
    try:
        return load(_proof_key(path))
    except OSError:
        return None


def start_proof(path: str, password: str | None = None) -> dict:
    """仕事を始める（動いていれば、終わっていれば、何もしない）。いまの控えを返す。"""
    key = _proof_key(path)
    rec = load(key) or {}
    if rec.get("finished") or _alive(rec.get("pid")):
        return rec
    if rec.get("pid"):  # 止まっていた
        rec = load(key) or rec
        if rec.get("finished") or _gave_up(key, rec, 0):
            return rec
    rec.pop("stopped", None)
    rec["started"] = rec.get("started") or time.time()
    proc = _spawn(key)
    rec["pid"] = proc.pid
    save(key, rec)
    _send(proc, {"password": password})
    return rec


def _run_proof(key: str, rec: dict, args: dict) -> int:
    from .api import _read
    from .book import open_book
    from .xlsx.reach import from_workbook, has_full_calc_mark, whole_calc

    _lower_priority()
    rec["pid"] = os.getpid()
    threading.Thread(target=_watch_memory, args=(key, rec), daemon=True).start()
    try:
        password = args.get("password")
        b = limits.Budget(total=None)
        b.calc_left = float("inf")
        with limits.call_budget(b):
            blob, name, _k, _lock = _read(rec["path"], password)
            wb = open_book(blob, name)
            graph = from_workbook(wb)  # 数式は元と同じ（書き込みの証明で確かめた）
            reach = graph.reach(tuple(x) for x in rec["seeds"])
            changed = {graph.key(*x) for x in rec["changed"]}

            def original():
                ob, on, _k2, _l2 = _read(rec["original"], password)
                return open_book(ob, on)

            calc, problems = whole_calc(wb, graph, reach, changed, rec.get("marked") or has_full_calc_mark(wb), original)
        rec["result"] = {"calc": calc, "problems": problems[:50]}
        rec["finished"] = time.time()
        try:
            import resource

            rec["peak_rss_mb"] = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss // (1 << 20 if sys.platform == "darwin" else 1 << 10)
        except ImportError:
            pass
        save(key, rec)
        return 0
    except Exception as e:  # noqa: BLE001  仕事の失敗は、控えで伝える
        rec["stopped"] = f"途中で問題が起きました（{type(e).__name__}）：{str(e)[:500]}"
        save(key, rec)
        return 1


# ---------------------------------------------------------------------- 窓を作る仕事（open_document が時間の上限に当たったとき）
def _window_key(path: str, encoding: str | None) -> str:
    return "win-" + job_key(path, encoding)


def _window_file(key: str) -> Path:
    return jobs_dir() / f"{key}.window"


def window_running(path: str, encoding: str | None) -> dict | None:
    rec = load(_window_key(path, encoding))
    return rec if rec and not rec.get("finished") and not rec.get("stopped") and _alive(rec.get("pid")) else None


def window_made(path: str, encoding: str | None) -> bool:
    """窓の仕事は終わっている（それでも時間の上限に当たるなら、窓のほかの所が遅い）。"""
    key = _window_key(path, encoding)
    rec = load(key)
    return bool(rec and rec.get("finished")) and _window_file(key).exists()


def window_failed(path: str, encoding: str | None) -> str | None:
    """窓の仕事が予算・メモリで止まった理由（なければ None）。同じファイルで仕事を始め直さない。"""
    rec = load(_window_key(path, encoding))
    return rec.get("stopped") if rec else None


def start_window(path: str, encoding: str | None) -> dict:
    """窓を作る仕事を始める（動いていれば、その控えを返す）。"""
    key = _window_key(path, encoding)
    rec = load(key)
    if rec and not rec.get("finished") and _alive(rec.get("pid")):
        return rec
    new = {"key": key, "kind": "window", "path": path, "encoding": encoding, "started": time.time()}
    if rec and not rec.get("finished"):
        new.update({k: rec[k] for k in ("tries", "tries_at", "stopped", "pid") if k in rec})
        if _gave_up(key, new, 0):
            return new
    proc = _spawn(key)
    new["pid"] = proc.pid
    save(key, new)
    _send(proc, {})
    return new


def window_result(path: str, encoding: str | None, sha: str):
    """長い仕事で作った窓（なければ None）。中身が同じファイルのものだけ。"""
    import pickle

    wf = _window_file(_window_key(path, encoding))
    try:
        with open(wf, "rb") as f:
            got_sha, out = pickle.load(f)
    except (OSError, ValueError, EOFError, pickle.PickleError, AttributeError, ImportError):
        return None
    if ttl_seconds() <= 0:
        wf.unlink(missing_ok=True)  # 残さない設定：読んだら消す
    return out if got_sha == sha else None


def _prune_windows(keep: Path) -> None:
    """置いた窓の大きさの合計を limits.WINDOW_CACHE までに（古いものから消す。いま作ったものは残す）。"""
    files = sorted((p for p in jobs_dir().glob("*.window") if p != keep), key=lambda p: p.stat().st_mtime, reverse=True)
    total = keep.stat().st_size
    for p in files:
        total += p.stat().st_size
        if total > limits.WINDOW_CACHE:
            p.unlink(missing_ok=True)


def _run_window(key: str, rec: dict) -> int:
    import pickle

    from . import api
    from .lazy.filebytes import sha256_hex

    _lower_priority()
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
        # 縛らないのは時間だけ。展開・描画・項目の予算は、ふつうの呼び出しと同じ（壊れた・悪意のある書類を、止まらない仕事にしない）
        with limits.call_budget():
            out = api.window(rec["path"], None, rec.get("encoding"))
            blob = api._read(rec["path"], None)[0]
        p = _window_file(key)
        tmp = p.with_suffix(f".{os.getpid()}.tmp")
        fd = os.open(tmp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "wb") as f:
            pickle.dump((sha256_hex(blob), out), f, protocol=pickle.HIGHEST_PROTOCOL)
        os.replace(tmp, p)
        _prune_windows(p)
        rec["finished"] = time.time()
        rec["progress"] = {"text": "終わりました"}
        save(key, rec)
        return 0
    except Exception as e:  # noqa: BLE001  仕事の失敗は、控えで伝える
        rec["stopped"] = f"途中で問題が起きました（{type(e).__name__}）：{str(e)[:500]}"
        save(key, rec)
        return 1


def _run_calc(key: str, rec: dict, args: dict) -> int:
    from .formula.lazyarr import ScanStore
    from .xlsx import model as xm

    _lower_priority()
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
            rec["results"] = xm.cells_for_job(args["path"], args.get("password"), args.get("values") or {}, rec["cells"], store)
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
    if rec.get("kind") == "window":
        return _run_window(key, rec)
    if rec.get("kind") == "proof":
        return _run_proof(key, rec, args)
    _lower_priority()
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
    from .embeds.lazy import whole

    try:
        with limits.call_budget(limits.Budget(total=None)), whole():  # 検査は、後回しにした埋め込みの表も読む
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
