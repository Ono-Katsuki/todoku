"""「圧縮して保存」した一太郎の文書（JSCompDocument のストリーム）を解く。外部のライブラリは使わない。

JSCompDocument：先頭の 2 バイトが中身の位置、"JustCompressedDocument"、その位置から LHA の書庫（レベル 1 の見出し、
方式 -lh5-）。書庫の中の 1 つのファイルが、圧縮しない元の一太郎の文書（複合ファイル）。

-lh5-：辞書 8 KB の LZ77 と、ブロックごとのハフマン符号（字と長さの符号 510 種・位置の符号 14 種・符号の長さの符号 19 種）。
ビットは上の桁から読む。ハフマン符号は、長さの短い順・同じ長さなら番号の順に割り当てる（正準の符号）。
解いた結果は、見出しの CRC-16（多項式 0xA001）と長さで確かめる（合わなければ JtdError）。
"""

from __future__ import annotations

import struct

from .book import JtdError

NC, NP, NT = 510, 14, 19  # 字と長さの符号・位置の符号・符号の長さの符号の数（-lh5-）
CBIT, PBIT, TBIT = 9, 4, 5
DICBIT, THRESHOLD = 13, 3


class _Bits:
    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0  # ビットの位置

    def peek(self, n: int) -> int:
        v = 0
        p = self.pos
        for _ in range(n):
            byte = self.data[p >> 3] if (p >> 3) < len(self.data) else 0
            v = (v << 1) | ((byte >> (7 - (p & 7))) & 1)
            p += 1
        return v

    def get(self, n: int) -> int:
        v = self.peek(n)
        self.pos += n
        return v


class _Huff:
    """正準のハフマン符号：長さの並び → 1 ビットずつたどる表。長さがすべて 0 なら、ただ 1 つの記号（ビットを使わない）。"""

    def __init__(self, lengths: list[int], only: int | None = None):
        self.only = only
        self.table: dict[tuple[int, int], int] = {}
        if only is not None:
            return
        code = 0
        for ln in range(1, 17):
            for sym, L in enumerate(lengths):
                if L == ln:
                    self.table[(ln, code)] = sym
                    code += 1
            code <<= 1
        self.maxlen = max(lengths) if lengths else 0

    def decode(self, bits: _Bits) -> int:
        if self.only is not None:
            return self.only
        code = 0
        for ln in range(1, self.maxlen + 1):
            code = (code << 1) | bits.get(1)
            sym = self.table.get((ln, code))
            if sym is not None:
                return sym
        raise JtdError("圧縮した文書のハフマン符号が壊れています")


def _read_pt(bits: _Bits, nn: int, nbit: int, special: int) -> _Huff:
    n = bits.get(nbit)
    if n == 0:
        return _Huff([], only=bits.get(nbit))
    lengths: list[int] = []
    while len(lengths) < n:
        c = bits.get(3)
        if c == 7:
            while bits.get(1):
                c += 1
        lengths.append(c)
        if len(lengths) == special:
            lengths += [0] * bits.get(2)
    lengths += [0] * (nn - len(lengths))
    return _Huff(lengths[:nn])


def _read_c(bits: _Bits, pt: _Huff) -> _Huff:
    n = bits.get(CBIT)
    if n == 0:
        return _Huff([], only=bits.get(CBIT))
    lengths: list[int] = []
    while len(lengths) < n:
        c = pt.decode(bits)
        if c == 0:
            lengths.append(0)
        elif c == 1:
            lengths += [0] * (bits.get(4) + 3)
        elif c == 2:
            lengths += [0] * (bits.get(CBIT) + 20)
        else:
            lengths.append(c - 2)
    lengths += [0] * (NC - len(lengths))
    return _Huff(lengths[:NC])


def unlh5(data: bytes, size: int) -> bytes:
    """-lh5- で縮めたデータ → 元のバイト列（size バイト）。"""
    bits = _Bits(data)
    out = bytearray()
    left = 0
    c_huff = p_huff = None
    limit = len(data) * 8 + 64
    while len(out) < size:
        if left == 0:
            left = bits.get(16)
            pt = _read_pt(bits, NT, TBIT, 3)
            c_huff = _read_c(bits, pt)
            p_huff = _read_pt(bits, NP, PBIT, -1)
            if left == 0:
                raise JtdError("圧縮した文書のブロックが空です")
        left -= 1
        c = c_huff.decode(bits)
        if c < 256:
            out.append(c)
        else:
            length = c - 256 + THRESHOLD
            j = p_huff.decode(bits)
            if j:
                j = (1 << (j - 1)) + bits.get(j - 1)
            dist = j + 1
            if dist > len(out):
                raise JtdError("圧縮した文書の参照が、解いた範囲の外を指しています")
            start = len(out) - dist
            for k in range(length):
                out.append(out[start + k])
        if bits.pos > limit:
            raise JtdError("圧縮した文書のデータが途中で終わっています")
    return bytes(out[:size])


def crc16(data: bytes) -> int:
    crc = 0
    for b in data:
        crc ^= b
        for _ in range(8):
            crc = (crc >> 1) ^ 0xA001 if crc & 1 else crc >> 1
    return crc


def unpack_jscomp(stream: bytes) -> bytes:
    """JSCompDocument のストリーム → 圧縮しない元の一太郎の文書（複合ファイル）。"""
    if len(stream) < 40 or stream[2:24] != b"JustCompressedDocument":
        raise JtdError("圧縮した一太郎の文書の見出しが想定と違います")
    at = struct.unpack_from("<H", stream, 0)[0]
    h = stream[at:]
    if len(h) < 24 or h[2:7] != b"-lh5-" or h[20] != 1:
        raise JtdError(f"圧縮の方式が想定（LHA レベル 1・-lh5-）と違います（{h[2:7]!r}）")
    hsize = h[0]
    packed, original = struct.unpack_from("<II", h, 7)
    nlen = h[21]
    crc = struct.unpack_from("<H", h, 22 + nlen)[0]
    pos = 2 + hsize
    ext_total = 0
    nxt = struct.unpack_from("<H", h, pos - 2)[0]
    while nxt:
        ext_total += nxt
        pos += nxt
        nxt = struct.unpack_from("<H", h, pos - 2)[0]
    body = h[pos:pos + packed - ext_total]
    out = unlh5(body, original)
    if len(out) != original or crc16(out) != crc:
        raise JtdError("圧縮した一太郎の文書を解いた結果が、見出しの長さ・CRC と合いません")
    if out[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise JtdError("圧縮した一太郎の文書の中身が、複合ファイルではありません")
    return out
