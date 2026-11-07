"""ファイルをメモリに読まずに、bytes のように扱う（FileBytes）。何 GB の PDF でも、触った所のページだけがメモリに載る。

mmap なので、re・struct・hashlib にそのまま渡せる（Python 3.11 には __buffer__ がないので、mmap の子のクラスにする）。
ファイルの全部をなめる処理（ハッシュ・探す・比べる）は pread で少しずつ読むので、常駐メモリを増やさない（macOS の madvise は
常駐メモリを減らさない。減らすには fresh で写像を作り直す）。
"""

from __future__ import annotations

import contextlib
import contextvars
import hashlib
import mmap
import os
from .pread import O_BINARY, pread
import re
import tempfile
import weakref
from pathlib import Path

STEP = 1 << 20
MAP_MIN = 1 << 20  # これより小さなファイルは、読んだ方が速い

_MAPPED: contextvars.ContextVar = contextvars.ContextVar("todoku_mapped_read", default=False)
_TEMP_DIR: contextvars.ContextVar = contextvars.ContextVar("todoku_temp_dir", default=None)


@contextlib.contextmanager
def mapped_reads(temp_dir: str | Path | None = None):
    """この中で api._read が読むファイルは（小さなものを除き）FileBytes にする。temp_dir：書いた結果の一時ファイルを置く所。"""
    t1 = _MAPPED.set(True)
    t2 = _TEMP_DIR.set(str(temp_dir) if temp_dir else _TEMP_DIR.get())
    try:
        yield
    finally:
        _MAPPED.reset(t1)
        _TEMP_DIR.reset(t2)


def mapping_on() -> bool:
    return _MAPPED.get()


def _unlink_quiet(path: str) -> None:
    with contextlib.suppress(OSError):
        os.unlink(path)


class FileBytes(mmap.mmap):
    """読むだけの mmap に、bytes の主な読み方を足したもの。"""

    def __new__(cls, path):
        fd = os.open(path, os.O_RDONLY | O_BINARY)
        try:
            self = super().__new__(cls, fd, 0, access=mmap.ACCESS_READ)
        finally:
            os.close(fd)
        self.path = str(path)
        self._owned = None
        return self

    # ------------------------------------------------------------------ bytes の読み方
    def startswith(self, prefix, start: int = 0, end: int | None = None) -> bool:
        if isinstance(prefix, tuple):
            return any(self.startswith(p, start, end) for p in prefix)
        stop = start + len(prefix)
        if end is not None and stop > end:
            return False
        return self[start:stop] == prefix

    def endswith(self, suffix) -> bool:
        if isinstance(suffix, tuple):
            return any(self.endswith(s) for s in suffix)
        return self[len(self) - len(suffix):] == suffix if len(self) >= len(suffix) else False

    def index(self, sub, *a) -> int:
        i = self.find(sub, *a)
        if i < 0:
            raise ValueError("subsection not found")
        return i

    def __contains__(self, sub) -> bool:
        return self.scan_find(sub) >= 0

    def __eq__(self, other) -> bool:
        if other is self:
            return True
        if len(self) != len(other):
            return False
        return same_prefix(self, other)

    def __ne__(self, other) -> bool:
        return not self.__eq__(other)

    __hash__ = None

    def __add__(self, other) -> bytes:
        return bytes(self) + bytes(other)

    def __radd__(self, other) -> bytes:
        return bytes(other) + bytes(self)

    def __getattr__(self, name):
        # ほかの bytes の読み方（decode・split など）：丸ごと読むしかない所。メモリの上限は働き手の見張りで止まる
        if name.startswith("__") or not hasattr(bytes, name):
            raise AttributeError(name)
        return getattr(bytes(self), name)

    # ------------------------------------------------------------------ 少しずつ読む
    def chunks(self, a: int = 0, b: int | None = None, step: int = STEP):
        b = len(self) if b is None else b
        fd = os.open(self.path, os.O_RDONLY | O_BINARY)
        try:
            pos = a
            while pos < b:
                c = pread(fd, min(step, b - pos), pos)
                if not c:
                    return
                yield c
                pos += len(c)
        finally:
            os.close(fd)

    def scan_find(self, sub: bytes) -> int:
        at = 0
        tail = b""
        for c in self.chunks():
            buf = tail + c
            k = buf.find(sub)
            if k >= 0:
                return at - len(tail) + k
            keep = min(len(sub) - 1, len(buf)) if len(sub) > 1 else 0
            tail = buf[-keep:] if keep else b""
            at += len(c)
        return -1

    def sha256(self) -> str:
        h = hashlib.sha256()
        for c in self.chunks():
            h.update(c)
        return h.hexdigest()

    def fresh(self) -> "FileBytes":
        """同じファイルの新しい写像（古い写像を手放すと、触ったページの分の常駐メモリが減る）。"""
        new = FileBytes(self.path)
        new._owned = self._owned
        return new

    def adopt_name(self, new_path: str) -> None:
        """持ち主の一時ファイルが名前を変えたことを覚える（もう消さない）。"""
        self.path = new_path
        if self._owned is not None:
            self._owned.detach()
            self._owned = None


def sha256_hex(blob) -> str:
    """bytes でも FileBytes でも（FileBytes は常駐メモリを増やさずに）。"""
    if isinstance(blob, FileBytes):
        return blob.sha256()
    return hashlib.sha256(blob).hexdigest()


def sha256_file(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            c = f.read(STEP)
            if not c:
                return h.hexdigest()
            h.update(c)


def same_prefix(new, original) -> bool:
    """new の頭が original のバイト列と同じか（どちらも丸ごとメモリに読まない）。"""
    n = len(original)
    if len(new) < n:
        return False
    if not isinstance(new, FileBytes) and not isinstance(original, FileBytes):
        return new[:n] == original
    pos = 0
    while pos < n:
        e = min(pos + STEP, n)
        a = _pread(new, pos, e)
        if a != _pread(original, pos, e):
            return False
        pos = e
    return True


def _pread(blob, a: int, b: int) -> bytes:
    if isinstance(blob, FileBytes):
        for c in blob.chunks(a, b, step=b - a):
            return c
        return b""
    return bytes(blob[a:b])


def iter_matches(pattern: re.Pattern, blob, overlap: int = 128):
    """re.finditer と同じ順に、(絶対の位置, match) を返す。FileBytes は少しずつ読む（overlap より長い一致は見つけない）。"""
    if not isinstance(blob, FileBytes):
        for m in pattern.finditer(blob):
            yield m.start(), m
        return
    at = 0
    tail = b""
    done_to = 0  # ここより前で始まる一致は返した
    for c in blob.chunks(step=1 << 22):
        buf = tail + c
        base = at - len(tail)
        last = len(buf) - overlap if at + len(c) < len(blob) else len(buf)
        for m in pattern.finditer(buf):
            s = base + m.start()
            if s < done_to:
                continue
            if m.start() >= last:
                break
            yield s, m
        done_to = base + max(last, 0)
        tail = buf[max(last, 0):]
        at += len(c)


def appended_file(orig: FileBytes, tail: bytes | bytearray) -> FileBytes:
    """元のファイルの後ろに tail を足した一時ファイルを、少しずつ写して作り、写像で開く（使い終わると消える）。"""
    d = _TEMP_DIR.get() or os.path.dirname(orig.path) or None
    fd, name = tempfile.mkstemp(prefix=".todoku-append.", suffix=Path(orig.path).suffix, dir=d)
    try:
        with os.fdopen(fd, "wb") as f:
            for c in orig.chunks():
                f.write(c)
            f.write(tail)
            f.flush()
        out = FileBytes(name)
    except BaseException:
        _unlink_quiet(name)
        raise
    out._owned = weakref.finalize(out, _unlink_quiet, name)
    return out


def write_out(f, blob) -> None:
    """開いたファイルに blob を書く（FileBytes は少しずつ）。"""
    if isinstance(blob, FileBytes):
        for c in blob.chunks():
            f.write(c)
    else:
        f.write(blob)
