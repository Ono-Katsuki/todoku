"""LZH・LHA（.lzh・.lha）：見出し（レベル 0・1・2・3）を読み、-lh0-・-lz4-（圧縮なし）と -lh4-〜-lh7- を解く。読むだけ。

-lh1-〜-lh3- と、それ以外の方式は解かない：見分けて、別の道具で展開する手順を返す（convert_first.py）。

見出しの形：
- レベル 0・1：[0] 見出しの大きさ、[1] 見出しの和、[2:7] 方式、[7:11] 圧縮後、[11:15] 元の大きさ、[15:19] 時刻、[19] 属性、[20] レベル、
  [21] 名前の長さ、名前、CRC-16。レベル 1 は OS 1 バイトと、次の拡張見出しの大きさ 2 バイトが続く（圧縮後の大きさは、拡張見出しを含む）
- レベル 2：[0:2] 見出し全体の大きさ、[21:23] CRC-16、[23] OS、[24:26] 次の拡張見出しの大きさ。拡張見出しの 0x01 は名前、0x02 は
  フォルダ（区切りは 0xFF）
- レベル 3：[0:2] 語の大きさ（4）、[21:23] CRC-16、[23] OS、[24:28] 見出し全体の大きさ、[28:32] 次の拡張見出しの大きさ（拡張見出しの大きさの欄も 4 バイト）

解いた結果は見出しの CRC-16 で確かめる（合わなければ読まない）。名前の文字コードは書庫に書かれていないので、ASCII・Shift_JIS・UTF-8 の順に
厳密に読めるもので読み、読み方を添える。Windows の自己解凍書庫（.exe の中の LZH）も、実行せずに中の書庫を読む。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

from ..jtd.lha import crc16

THRESHOLD = 3
NC, NT, TBIT, CBIT = 510, 19, 5, 9
# 方式 → (辞書のビット数, 位置の符号の数, 位置の符号の数のビット数)
_METHODS = {"-lh4-": (12, 14, 4), "-lh5-": (13, 14, 4), "-lh6-": (15, 16, 5), "-lh7-": (16, 17, 5)}
STORED = {"-lh0-", "-lz4-"}
MAX_MEMBER = 512 * 1024 * 1024
MAX_RATIO = 1000


class LzhError(ValueError):
    pass


@dataclass
class Entry:
    name: str
    named: str
    method: str
    csize: int
    size: int
    crc: int
    start: int  # 圧縮したデータの位置
    is_dir: bool


class _Bits:
    def __init__(self, data: bytes):
        self.data = data
        self.p = 0
        self.acc = 0
        self.n = 0
        self.over = 0

    def get(self, k: int) -> int:
        if k == 0:
            return 0
        while self.n < k:
            if self.p < len(self.data):
                self.acc = (self.acc << 8) | self.data[self.p]
                self.p += 1
            else:
                self.acc <<= 8
                self.over += 1
                if self.over > 64:
                    raise LzhError("圧縮したデータが途中で終わっています")
            self.n += 8
        self.n -= k
        v = self.acc >> self.n
        self.acc &= (1 << self.n) - 1
        return v


class _Huff:
    """正準のハフマン符号（短い順・同じ長さなら番号の順）。1 ビットずつ、長さごとの範囲で引く。"""

    def __init__(self, lengths: list[int], only: int | None = None):
        self.only = only
        if only is not None:
            return
        self.first = [0] * 18
        self.count = [0] * 18
        self.index = [0] * 18
        self.syms: list[int] = []
        code = 0
        for ln in range(1, 17):
            self.first[ln] = code
            self.index[ln] = len(self.syms)
            for sym, x in enumerate(lengths):
                if x == ln:
                    self.syms.append(sym)
            self.count[ln] = len(self.syms) - self.index[ln]
            code = (code + self.count[ln]) << 1
        if not self.syms:
            raise LzhError("ハフマンの表が空です")

    def decode(self, b: _Bits) -> int:
        if self.only is not None:
            return self.only
        code = 0
        for ln in range(1, 17):
            code = (code << 1) | b.get(1)
            i = code - self.first[ln]
            if 0 <= i < self.count[ln]:
                return self.syms[self.index[ln] + i]
        raise LzhError("ハフマンの符号が壊れています")


def _read_pt(b: _Bits, nn: int, nbit: int, special: int) -> _Huff:
    n = b.get(nbit)
    if n == 0:
        return _Huff([], only=b.get(nbit))
    if n > nn:
        raise LzhError("符号の表が大きすぎます")
    lengths: list[int] = []
    while len(lengths) < n:
        c = b.get(3)
        if c == 7:
            while b.get(1):
                c += 1
                if c > 16:
                    raise LzhError("符号の長さが壊れています")
        lengths.append(c)
        if len(lengths) == special:
            lengths += [0] * b.get(2)
    return _Huff((lengths + [0] * nn)[:nn])


def _read_c(b: _Bits, pt: _Huff) -> _Huff:
    n = b.get(CBIT)
    if n == 0:
        return _Huff([], only=b.get(CBIT))
    if n > NC:
        raise LzhError("符号の表が大きすぎます")
    lengths: list[int] = []
    while len(lengths) < n:
        c = pt.decode(b)
        if c == 0:
            lengths.append(0)
        elif c == 1:
            lengths += [0] * (b.get(4) + 3)
        elif c == 2:
            lengths += [0] * (b.get(CBIT) + 20)
        else:
            lengths.append(c - 2)
    return _Huff((lengths + [0] * NC)[:NC])


def unlh(data: bytes, size: int, method: str) -> bytes:
    """-lh4-〜-lh7-：辞書（dicbit）と位置の符号の数のほかは、同じ形（ブロックごとのハフマン）。"""
    dicbit, np, pbit = _METHODS[method]
    b = _Bits(data)
    out = bytearray()
    left = 0
    c_huff = p_huff = None
    while len(out) < size:
        if left == 0:
            left = b.get(16)
            if left == 0:
                raise LzhError("ブロックが空です")
            pt = _read_pt(b, NT, TBIT, 3)
            c_huff = _read_c(b, pt)
            p_huff = _read_pt(b, np, pbit, -1)
        left -= 1
        c = c_huff.decode(b)
        if c < 256:
            out.append(c)
            continue
        length = c - 256 + THRESHOLD
        j = p_huff.decode(b)
        if j:
            j = (1 << (j - 1)) + b.get(j - 1)
        dist = j + 1
        if dist > len(out) or dist > (1 << dicbit):
            raise LzhError("参照が、解いた範囲の外を指しています")
        start = len(out) - dist
        for k in range(length):
            out.append(out[start + k])
    return bytes(out[:size])


def _name(raw: bytes) -> tuple[str, str]:
    raw = raw.replace(b"\xff", b"/")  # 0xFF は Shift_JIS の字の中に出ないので、バイトのまま。\\ は字の 2 バイト目にも出るので、字にしてから
    if raw.isascii():
        return raw.decode("ascii").replace("\\", "/"), "ASCII"
    for enc, how in (("cp932", "Shift_JIS（推定）"), ("utf-8", "UTF-8（推定）")):
        try:
            return raw.decode(enc).replace("\\", "/"), how
        except UnicodeDecodeError:
            continue
    return raw.decode("latin-1").replace("\\", "/"), "Latin-1（どの文字コードでも読めません）"


def _exts(blob: bytes, pos: int, next_size: int, width: int = 2) -> tuple[bytes, bytes, int]:
    """拡張見出し：(名前, フォルダ, 拡張見出しの大きさの合計)。width：次の大きさの欄のバイト数（レベル 3 は 4）。"""
    name = b""
    folder = b""
    total = 0
    while next_size:
        if next_size < 1 + width or pos + next_size > len(blob):
            raise LzhError("拡張見出しが壊れています")
        t = blob[pos]
        body = blob[pos + 1:pos + next_size - width]
        if t == 0x01:
            name = body
        elif t == 0x02:
            folder = body
        total += next_size
        pos += next_size
        next_size = int.from_bytes(blob[pos - width:pos], "little")
    return name, folder, total


def _header_at(blob: bytes, pos: int) -> tuple[Entry, int] | None:
    """(項目, 次の見出しの位置)。終わりの印なら None。"""
    if pos >= len(blob) or blob[pos] == 0:
        return None
    if pos + 22 > len(blob):
        raise LzhError("見出しが途中で終わっています")
    level = blob[pos + 20]
    method = blob[pos + 2:pos + 7].decode("ascii", "replace")
    csize, size = struct.unpack_from("<II", blob, pos + 7)
    if level in (0, 1):
        hs = blob[pos]
        if pos + 2 + hs > len(blob):
            raise LzhError("見出しが途中で終わっています")
        if sum(blob[pos + 2:pos + 2 + hs]) & 0xFF != blob[pos + 1]:
            raise LzhError("見出しの和が合いません")
        nlen = blob[pos + 21]
        raw = blob[pos + 22:pos + 22 + nlen]
        crc = struct.unpack_from("<H", blob, pos + 22 + nlen)[0]
        data_at = pos + 2 + hs
        folder = b""
        if level == 1:
            next_size = struct.unpack_from("<H", blob, pos + hs)[0]
            n2, folder, total = _exts(blob, data_at, next_size)
            raw = n2 or raw
            data_at += total
            csize -= total
            if csize < 0:
                raise LzhError("見出しの大きさが壊れています")
    elif level == 2:
        total = struct.unpack_from("<H", blob, pos)[0]
        crc = struct.unpack_from("<H", blob, pos + 21)[0]
        next_size = struct.unpack_from("<H", blob, pos + 24)[0]
        raw, folder, _t = _exts(blob, pos + 26, next_size)
        data_at = pos + total
    elif level == 3:
        if blob[pos:pos + 2] != b"\x04\x00" or pos + 32 > len(blob):
            raise LzhError("レベル 3 の見出しが壊れています")
        crc = struct.unpack_from("<H", blob, pos + 21)[0]
        total, next_size = struct.unpack_from("<II", blob, pos + 24)
        raw, folder, _t = _exts(blob, pos + 32, next_size, 4)
        data_at = pos + total
    else:
        raise LzhError(f"見出しのレベル {level} は読めません")
    if data_at + csize > len(blob):
        raise LzhError("圧縮したデータが書庫の外に出ています")
    is_dir = method == "-lhd-"
    nm, how = _name(folder + raw)
    return Entry(nm, how, method, csize, size, crc, data_at, is_dir), data_at + csize


def start_of(blob: bytes) -> int | None:
    """書庫の始まりの位置：先頭か、自己解凍の .exe（MZ）の中の、和の合う最初の見出し。なければ None。"""
    def ok(p: int) -> bool:
        if p < 0 or p + 22 > len(blob) or blob[p + 2:p + 4] != b"-l" or blob[p + 6] != 0x2D or blob[p + 20] > 3:
            return False
        if blob[p + 20] in (0, 1):
            hs = blob[p]
            return hs >= 22 and p + 2 + hs <= len(blob) and sum(blob[p + 2:p + 2 + hs]) & 0xFF == blob[p + 1]
        return True

    if ok(0):
        return 0
    if blob[:2] == b"MZ":
        p = blob.find(b"-lh")
        while p != -1:
            if ok(p - 2):
                return p - 2
            p = blob.find(b"-lh", p + 1)
    return None


def entries(blob: bytes) -> list[Entry]:
    s = start_of(blob)
    if s is None:
        raise LzhError("LZH の書庫の見出しがありません")
    out = []
    pos = s
    while True:
        try:
            r = _header_at(blob, pos)
        except (struct.error, IndexError) as e:
            raise LzhError(f"見出しが壊れています（{e}）") from None
        if r is None:
            return out
        out.append(r[0])
        pos = r[1]


def extract(blob: bytes, e: Entry) -> bytes:
    if e.is_dir:
        raise LzhError(f"「{e.name}」はフォルダです")
    if e.size > MAX_MEMBER:
        raise LzhError(f"「{e.name}」は展開すると {e.size:,} バイトで、取り出せる大きさを超えます")
    data = blob[e.start:e.start + e.csize]
    if e.method in STORED:
        out = data
    elif e.method in _METHODS:
        if e.csize and e.size / e.csize > MAX_RATIO:
            raise LzhError(f"「{e.name}」は圧縮率が {e.size / e.csize:,.0f} 倍で、展開して大きくなる書庫として取り出しません")
        from ..limits import check_output

        check_output(e.size, f"「{e.name}」")
        out = unlh(data, e.size, e.method)
    else:
        from .convert_first import message

        raise LzhError(f"「{e.name}」の圧縮の方式 {e.method} は、とどくでは解きません。" + message("lzh"))
    if len(out) != e.size or crc16(out) != e.crc:
        raise LzhError(f"「{e.name}」を解いた結果が、見出しの CRC-16 と合いません（書庫が壊れています）")
    return out


def needs_other_tool(e: Entry) -> bool:
    """とどくが解かない方式か（別の道具で展開する）。"""
    return not e.is_dir and e.method not in STORED and e.method not in _METHODS
