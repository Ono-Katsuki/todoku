"""複合ファイル（CFB）を、丸ごと読まずに読む：ストリームを、バイトの範囲で読める形（Stream）にする。

- FileBytes の上で cfb.CFB を開く（ディレクトリと FAT だけを読む）
- Stream：FAT のつながりを、続いたセクタの並びの表にして、範囲を pread で読む（メモリはセクタの数でなく、つながりの切れ目の数に比例）
- spool：大きなストリームを一時ファイルに写して FileBytes で開く（struct・スライスで読む窓の読み手を、そのまま使える）。元のファイルと同じ字が
  書かれた一時ファイルなので、開いたあとすぐ名前を消す（ほかから見えない）
"""

from __future__ import annotations

import bisect
import os
from .pread import O_BINARY, pread
import tempfile
from array import array

from ..cfb import CFB, ENDOFCHAIN, FREESECT
from .filebytes import FileBytes


class Stream:
    """複合ファイルの中の 1 つのストリーム。"""

    def __init__(self, cf: "CfbFile", name: str):
        self.cf = cf
        self.name = name
        e = cf.cfb._find(name)
        self.size = e.size
        self.mini = e.size < cf.cfb.cutoff and name != ""
        cfb = cf.cfb
        self.starts = array("Q")  # ストリームの中の位置
        self.at = array("Q")  # ファイルの位置
        unit = cfb.msize if self.mini else cfb.ssize
        if self.mini:
            self._data = bytes(cfb._mini_bytes(e.start, e.size))[:e.size]
            return
        self._data = None
        pos = 0
        last_end = -1
        s = e.start
        seen = 0
        fat = cfb.fat
        limit = len(fat)
        while s not in (ENDOFCHAIN, FREESECT) and s < limit and pos < e.size:
            off = (s + 1) * unit
            if off != last_end:
                self.starts.append(pos)
                self.at.append(off)
            last_end = off + unit
            pos += unit
            s = fat[s]
            seen += 1
            if seen > limit:
                break  # 輪になったつながり
        self.available = min(pos, e.size)

    def read(self, a: int, b: int) -> bytes:
        b = min(b, self.size)
        if a >= b:
            return b""
        if self._data is not None:
            return self._data[a:b]
        out = bytearray()
        i = bisect.bisect_right(self.starts, a) - 1
        fd = self.cf.fd
        pos = a
        while pos < b and i < len(self.starts):
            seg_end = self.starts[i + 1] if i + 1 < len(self.starts) else b
            take = min(b, seg_end) - pos
            if take > 0:
                c = pread(fd, take, self.at[i] + pos - self.starts[i])
                out += c
                if len(c) < take:
                    break
                pos += take
            i += 1
        return bytes(out)

    def chunks(self, a: int = 0, step: int = 1 << 20):
        pos = a
        while pos < self.size:
            c = self.read(pos, min(pos + step, self.size))
            if not c:
                return
            yield c
            pos += len(c)


class CfbFile:
    def __init__(self, path: str):
        self.path = path
        self.fb = FileBytes(path)
        self.cfb = CFB(self.fb)
        self.fd = os.open(path, os.O_RDONLY | O_BINARY)
        self.size = len(self.fb)
        self._streams = {n.lower(): n for n in self.cfb.streams()}

    def close(self) -> None:
        os.close(self.fd)

    def find(self, name: str) -> str | None:
        return self._streams.get(name.lower())

    def stream(self, name: str) -> Stream:
        return Stream(self, name)

    def spool(self, name: str):
        """ストリームの中身を bytes のように：小さければ bytes、大きければ一時ファイルの写像。"""
        st = self.stream(name)
        if st.size < 1 << 20:
            return st.read(0, st.size)
        fd, tmp = tempfile.mkstemp(prefix=".todoku-stream.")
        try:
            with os.fdopen(fd, "wb") as f:
                for c in st.chunks():
                    f.write(c)
            out = FileBytes(tmp)
        finally:
            os.unlink(tmp)
        return out


class SpoolCfb:
    """cfb.CFB と同じ問い合わせで、大きなストリームを写像で返す（.doc の読み手に渡す）。"""

    def __init__(self, cf: CfbFile):
        self._cf = cf
        self._memo: dict = {}

    def __getattr__(self, name):
        return getattr(self._cf.cfb, name)

    def read(self, path: str):
        got = self._memo.get(path)
        if got is None:
            got = self._memo[path] = self._cf.spool(path)
        return got
