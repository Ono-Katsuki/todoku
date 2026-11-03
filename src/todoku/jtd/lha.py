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
    from ..limits import check_output

    check_output(size, "-lh5- で縮めたデータ")  # 解く大きさ（見出しの値）は、上限まで（out は size + 256 バイトまでにしかならない）
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
    from ..limits import budget, check_output

    check_output(original, "圧縮した一太郎の文書")  # 見出しの元の大きさ（展開の前に、上限と比べる）
    out = unlh5(body, original)
    budget().charge(len(out), "圧縮した一太郎の文書")
    if len(out) != original or crc16(out) != crc:
        raise JtdError("圧縮した一太郎の文書を解いた結果が、見出しの長さ・CRC と合いません")
    if out[:8] != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        raise JtdError("圧縮した一太郎の文書の中身が、複合ファイルではありません")
    return out


# ====================================================================== 縮める（-lh5-）
class _BitWriter:
    def __init__(self):
        self.buf = bytearray()
        self.acc = 0
        self.n = 0

    def put(self, nbits: int, value: int) -> None:
        if nbits <= 0:
            return
        self.acc = (self.acc << nbits) | (value & ((1 << nbits) - 1))
        self.n += nbits
        while self.n >= 8:
            self.n -= 8
            self.buf.append((self.acc >> self.n) & 0xFF)
        self.acc &= (1 << self.n) - 1

    def done(self) -> bytes:
        if self.n:
            self.buf.append((self.acc << (8 - self.n)) & 0xFF)
            self.acc, self.n = 0, 0
        return bytes(self.buf)


def _lengths(freq: list[int], limit: int = 16) -> list[int]:
    """出てくる数 → ハフマン符号の長さ（limit ビットまで）。長すぎるときは数を半分にして作り直す。"""
    import heapq

    f = list(freq)
    while True:
        heap = [(x, i, (i,)) for i, x in enumerate(f) if x > 0]
        out = [0] * len(f)
        if len(heap) < 2:
            return out
        heapq.heapify(heap)
        uid = len(f)
        depth = {i: 0 for _x, i, _s in heap}
        while len(heap) > 1:
            a = heapq.heappop(heap)
            b = heapq.heappop(heap)
            for s in a[2] + b[2]:
                depth[s] += 1
            heapq.heappush(heap, (a[0] + b[0], uid, a[2] + b[2]))
            uid += 1
        for s, d in depth.items():
            out[s] = d
        if max(out) <= limit:
            return out
        f = [(x + 1) // 2 if x else 0 for x in f]


def _codes(lengths: list[int]) -> list[int]:
    """長さ → 正準の符号（長さの短い順、同じ長さなら番号の順。_Huff と同じ割り当て）。"""
    codes = [0] * len(lengths)
    code = 0
    for ln in range(1, 17):
        for sym, L in enumerate(lengths):
            if L == ln:
                codes[sym] = code
                code += 1
        code <<= 1
    return codes


def _lz77(data: bytes):
    """LZ77（辞書 8 KB、長さ 3〜256）：('c', 字) と ('m', 長さ, 距離) の並び。"""
    out = []
    n = len(data)
    window = 1 << DICBIT
    maxlen = 256
    head: dict[bytes, list[int]] = {}
    i = 0
    while i < n:
        best_len, best_dist = 0, 0
        if i + THRESHOLD <= n:
            cand = head.get(data[i:i + THRESHOLD])
            if cand:
                lim = min(maxlen, n - i)
                for p in reversed(cand[-48:]):
                    d = i - p
                    if d > window:
                        break
                    k = THRESHOLD
                    while k < lim and data[p + k] == data[i + k]:
                        k += 1
                    if k > best_len:
                        best_len, best_dist = k, d
                        if k == lim:
                            break
        step = best_len if best_len >= THRESHOLD else 1
        for q in range(i, min(i + step, n - THRESHOLD + 1)):
            head.setdefault(data[q:q + THRESHOLD], []).append(q)
        if best_len >= THRESHOLD:
            out.append(("m", best_len, best_dist))
        else:
            out.append(("c", data[i]))
        i += step
    return out


def _write_pt(bw: _BitWriter, lengths: list[int], nbit: int, special: int) -> None:
    """_read_pt の逆：使う数・3 ビットの長さ（7 以上は 1 を続けて 0 で終わる）・special 番目の後ろの 0 の数（2 ビット）。"""
    n = len(lengths)
    while n and lengths[n - 1] == 0:
        n -= 1
    bw.put(nbit, n)
    i = 0
    while i < n:
        k = lengths[i]
        i += 1
        if k <= 6:
            bw.put(3, k)
        else:
            bw.put(3, 7)
            bw.put(k - 7, (1 << (k - 7)) - 1)
            bw.put(1, 0)
        if i == special:
            z = 0
            while i < n and z < 3 and lengths[i] == 0:
                i += 1
                z += 1
            bw.put(2, z)


def _c_tokens(clen: list[int]) -> list[tuple[int, int, int]]:
    """字と長さの符号の長さ → 符号の長さの符号の並び (記号, 足すビット数, 値)。_read_c の逆。"""
    n = len(clen)
    while n and clen[n - 1] == 0:
        n -= 1
    out = []
    i = 0
    while i < n:
        k = clen[i]
        i += 1
        if k:
            out.append((k + 2, 0, 0))
            continue
        cnt = 1
        while i < n and clen[i] == 0:
            i += 1
            cnt += 1
        if cnt <= 2:
            out += [(0, 0, 0)] * cnt
        elif cnt <= 18:
            out.append((1, 4, cnt - 3))
        elif cnt == 19:
            out += [(0, 0, 0), (1, 4, 15)]
        else:
            out.append((2, CBIT, cnt - 20))
    return out


def _single(bw: _BitWriter, nbit: int, sym: int) -> None:
    """使う記号が 1 つだけの表：数 0 と、その記号。"""
    bw.put(nbit, 0)
    bw.put(nbit, sym)


def lh5(data: bytes, block: int = 16384) -> bytes:
    """バイト列 → -lh5- で縮めたデータ（unlh5 で元に戻る）。"""
    toks = _lz77(data)
    bw = _BitWriter()
    for s in range(0, len(toks), block):
        part = toks[s:s + block]
        cf = [0] * NC
        pf = [0] * NP
        syms = []
        for t in part:
            if t[0] == "c":
                cf[t[1]] += 1
                syms.append((t[1], None))
            else:
                c = 256 + t[1] - THRESHOLD
                j = t[2] - 1
                p = j.bit_length()
                cf[c] += 1
                pf[p] += 1
                syms.append((c, (p, j)))
        bw.put(16, len(part))
        clen = _lengths(cf)
        used_c = [i for i, x in enumerate(cf) if x]
        ccode = None
        if len(used_c) == 1:
            _single(bw, TBIT, 0)
            _single(bw, CBIT, used_c[0])
        else:
            ct = _c_tokens(clen)
            tf = [0] * NT
            for sym, _nb, _v in ct:
                tf[sym] += 1
            tlen = _lengths(tf)
            used_t = [i for i, x in enumerate(tf) if x]
            tcode = None
            if len(used_t) == 1:
                _single(bw, TBIT, used_t[0])
            else:
                _write_pt(bw, tlen, TBIT, 3)
                tcode = _codes(tlen)
            n = len(clen)
            while n and clen[n - 1] == 0:
                n -= 1
            bw.put(CBIT, n)
            for sym, nb, v in ct:
                if tcode is not None:
                    bw.put(tlen[sym], tcode[sym])
                bw.put(nb, v)
            ccode = _codes(clen)
        plen = _lengths(pf)
        used_p = [i for i, x in enumerate(pf) if x]
        pcode = None
        if len(used_p) <= 1:
            _single(bw, PBIT, used_p[0] if used_p else 0)
        else:
            _write_pt(bw, plen, PBIT, -1)
            pcode = _codes(plen)
        for c, pos in syms:
            if ccode is not None:
                bw.put(clen[c], ccode[c])
            if pos is not None:
                p, j = pos
                if pcode is not None:
                    bw.put(plen[p], pcode[p])
                if p > 1:
                    bw.put(p - 1, j - (1 << (p - 1)))
    return bw.done()


def jscomp_parts(stream: bytes) -> dict:
    """JSCompDocument のストリーム → 部分（前置き・LHA の見出し・縮めたデータ・後ろ）と、見出しの値。"""
    if len(stream) < 40 or stream[2:24] != b"JustCompressedDocument":
        raise JtdError("圧縮した一太郎の文書の見出しが想定と違います")
    at = struct.unpack_from("<H", stream, 0)[0]
    h = stream[at:]
    if len(h) < 24 or h[2:7] != b"-lh5-" or h[20] != 1:
        raise JtdError(f"圧縮の方式が想定（LHA レベル 1・-lh5-）と違います（{h[2:7]!r}）")
    hsize = h[0]
    packed, original = struct.unpack_from("<II", h, 7)
    nlen = h[21]
    crc_at = 22 + nlen
    ext = struct.unpack_from("<H", h, 2 + hsize - 2)[0]
    return {"at": at, "hsize": hsize, "packed": packed, "original": original, "crc_at": crc_at, "ext": ext,
            "crc": struct.unpack_from("<H", h, crc_at)[0], "checksum": h[1],
            "prefix": stream[:at], "header": h[:2 + hsize], "data": h[2 + hsize:2 + hsize + packed],
            "tail": h[2 + hsize + packed:]}


def pack_jscomp(stream: bytes, native: bytes) -> bytes:
    """元の JSCompDocument のストリームの、中の文書だけを native に替えたストリーム。
    前置き（"JustCompressedDocument" とその後ろの値）・見出しの名前・日時・属性・OS の印・書庫の終わりの 0 は元のまま。
    見出しの「縮めたデータの長さ・元の長さ・CRC-16・見出しの合計（チェックサム）」だけを、新しいデータに合わせる。"""
    pt = jscomp_parts(stream)
    if pt["ext"]:
        raise JtdError("圧縮した文書の LHA の見出しに拡張の見出しがあります（意味を確かめていないので、縮め直しません）")
    packed = lh5(native)
    head = bytearray(pt["header"])
    struct.pack_into("<II", head, 7, len(packed), len(native))
    struct.pack_into("<H", head, pt["crc_at"], crc16(native))
    head[1] = sum(head[2:2 + pt["hsize"]]) & 0xFF
    return pt["prefix"] + bytes(head) + packed + pt["tail"]
