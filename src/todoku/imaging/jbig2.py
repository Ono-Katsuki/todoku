"""JBIG2（ITU-T T.88 | ISO/IEC 14492）の復号。PDF の JBIG2Decode（埋め込みの編成、ISO 32000-1 7.4.7）と、単独のファイルの編成。

実装する所：区切り（segment）の見出し（7.2）・ページの情報（7.4.8）・汎用領域（6.2：算術 MQ の 4 つの型・TPGD・AT、MMR）・
細かくする領域（6.3：2 つの型・TPGR）・記号の辞書（6.5：算術・ハフマン・細かくする／集める）・字の領域（6.4）・
模様の辞書（6.7）・網点の領域（6.6・附属書 C）・ハフマンの表（附属書 B：標準の B.1〜B.15 と、表の区切り）・
整数の算術符号（附属書 A）・JBIG2Globals。

画素は 1 画素 1 バイト（0 か 1。1 が黒）の行の並びで持つ：純粋な Python では、1 画素ずつ取り出す所はこれがいちばん速く、
重ねる所は行を大きな整数にして一度に計算できる。結果は page.px（0＝黒、255＝白）にして渡す（7.4.7 のとおり JBIG2 の 1 を黒に）。

予算：領域の大きさは limits.MAX_PIXELS まで。deadline（time.monotonic の時刻）を過ぎたら、そこまでで止めて、
描けた所まで返す（problems に書く）。
"""

from __future__ import annotations

import struct
import time

from ..limits import MAX_PIXELS, LimitExceeded
from .mq import MQ, NL, NM, QE, QS


class Jbig2Error(ValueError):
    pass


class _Stopped(Exception):
    """deadline を過ぎた（描けた所までを返す）。"""


class Bitmap:
    __slots__ = ("w", "h", "rows")

    def __init__(self, w: int, h: int, rows: list | None = None, fill: int = 0):
        self.w = w
        self.h = h
        self.rows = rows if rows is not None else [bytearray([fill]) * w for _ in range(h)]


def _check_size(w: int, h: int, what: str) -> None:
    if w < 0 or h < 0 or w * h > MAX_PIXELS:
        raise LimitExceeded(f"JBIG2 の{what}（{w:,} × {h:,} 画素）が、開ける大きさ {MAX_PIXELS:,} 画素を超えます")


# ====================================================================== 重ね方（7.4.1.5：OR・AND・XOR・XNOR・REPLACE）
def compose(dst: Bitmap, src: Bitmap, x: int, y: int, op: int) -> None:
    """src を dst の (x, y) に重ねる。"""
    sx0, sx1 = max(0, -x), min(src.w, dst.w - x)
    if sx1 <= sx0:
        return
    n = sx1 - sx0
    dx = x + sx0
    frm = int.from_bytes
    full = sx0 == 0 and sx1 == src.w
    one = int.from_bytes(b"\x01" * n, "big") if op == 3 else 0
    drows, srows = dst.rows, src.rows
    for j in range(max(0, -y), min(src.h, dst.h - y)):
        d = drows[y + j]
        s = srows[j]
        if not full:
            s = s[sx0:sx1]
        if op == 4:
            d[dx:dx + n] = s
            continue
        a, b = frm(d[dx:dx + n], "big"), frm(s, "big")
        if op == 0:
            if not b:
                continue
            v = a | b
        elif op == 1:
            v = a & b
        elif op == 2:
            v = a ^ b
        else:
            v = (a ^ b) ^ one
        d[dx:dx + n] = v.to_bytes(n, "big")


# ====================================================================== 汎用領域の算術符号（6.2）
_TPGD_CX = (0x9B25, 0x0795, 0x00E5, 0x0195)
# 型ごと：(今の行の字数, 上の行の字数, 上の行の右へ何画素, 2 つ上の行の字数, 2 つ上の行の右へ何画素)。AT の既定の位置も含む（図 3〜6）
_GB = ((4, 7, 3, 5, 2), (3, 6, 3, 4, 2), (2, 5, 2, 3, 1), (4, 6, 2, 0, 0))
_AT_NOMINAL = (((3, -1), (-3, -1), (2, -2), (-2, -2)), ((3, -1),), ((2, -1),), ((2, -1),))
_AT_BITS = ((4, 10, 11, 15), (3,), (2,), (4,))


def gb_context_size(template: int) -> int:
    return (1 << 16, 1 << 13, 1 << 10, 1 << 10)[template]


def decode_generic(mq: MQ, cx, w: int, h: int, template: int, tpgd: bool = False, at=None, skip: Bitmap | None = None,
                   deadline: float | None = None, partial: list | None = None) -> Bitmap:
    """算術符号の汎用領域（6.2.5）。cx：文脈の表。at：AT 画素の [(dx, dy)]（省くと既定の位置）。skip：飛ばす画素（0 と決まる）。"""
    _check_size(w, h, "汎用領域")
    at = list(at) if at else list(_AT_NOMINAL[template])
    nominal = all(tuple(a) == b for a, b in zip(at, _AT_NOMINAL[template]))
    if not nominal or skip is not None:
        return _generic_slow(mq, cx, w, h, template, tpgd, at, skip, deadline, partial)
    n0, n1, right1, n2, right2 = _GB[template]
    mask0, mask1, mask2 = (1 << n0) - 1, (1 << n1) - 1, (1 << n2) - 1
    s1, s2 = n0, n0 + n1
    tp = _TPGD_CX[template]
    zero = bytes(w + 8)
    r1 = r2 = zero
    rows: list = []
    ltp = 0
    # MQ の状態を、その場の変数に（mq.decode と同じ手順を、画素の繰り返しの中に書き写す。ここが画素ごとに通るいちばん熱い所）
    qs_t, qe_t, nm_t, nl_t = QS, QE, NM, NL
    a, c, ct = mq.a, mq.c, mq.ct
    data = mq.data
    bp = mq.bp
    clock = time.monotonic
    for y in range(h):
        if deadline is not None and y & 63 == 0 and clock() > deadline:
            mq.a, mq.c, mq.ct, mq.bp = a, c, ct, bp
            if partial is not None:
                partial.append(y)
            rows += [bytearray(w + 8) for _ in range(h - y)]
            break
        if tpgd:
            mq.a, mq.c, mq.ct, mq.bp = a, c, ct, bp
            ltp ^= mq.decode(cx, tp)
            a, c, ct, bp = mq.a, mq.c, mq.ct, mq.bp
            if ltp:
                rows.append(bytearray(r1))
                r2, r1 = r1, rows[-1]
                continue
        cur = bytearray(w + 8)
        l1 = 0
        for k in range(right1 + 1):
            l1 = (l1 << 1) | r1[k]
        l2 = 0
        if n2:
            for k in range(right2 + 1):
                l2 = (l2 << 1) | r2[k]
        l0 = 0
        i1, i2 = right1 + 1, right2 + 1
        for x, b1, b2 in zip(range(w), r1[i1:], r2[i2:]):
            i = (l2 << s2) | (l1 << s1) | l0
            s = cx[i]
            q = qe_t[s]
            qs = qs_t[s]
            a -= q
            if c < qs:
                if a < q:
                    d = s & 1
                    cx[i] = nm_t[s]
                else:
                    d = (s & 1) ^ 1
                    cx[i] = nl_t[s]
                a = q
            else:
                c -= qs
                if a >= 0x8000:
                    d = s & 1
                    if d:
                        cur[x] = 1
                    l0 = ((l0 << 1) | d) & mask0
                    l1 = ((l1 << 1) | b1) & mask1
                    l2 = ((l2 << 1) | b2) & mask2
                    continue
                if a < q:
                    d = (s & 1) ^ 1
                    cx[i] = nl_t[s]
                else:
                    d = s & 1
                    cx[i] = nm_t[s]
            while True:
                if ct == 0:
                    if data[bp] == 0xFF:
                        b = data[bp + 1]
                        if b > 0x8F:
                            c += 0xFF00
                            ct = 8
                        else:
                            bp += 1
                            c += b << 9
                            ct = 7
                    else:
                        bp += 1
                        c += data[bp] << 8
                        ct = 8
                a <<= 1
                c <<= 1
                ct -= 1
                if a >= 0x8000:
                    break
            if d:
                cur[x] = 1
            l0 = ((l0 << 1) | d) & mask0
            l1 = ((l1 << 1) | b1) & mask1
            l2 = ((l2 << 1) | b2) & mask2
        rows.append(cur)
        r2, r1 = r1, cur
    mq.a, mq.c, mq.ct, mq.bp = a, c, ct, bp
    return Bitmap(w, h, [r[:w] for r in rows])


def _generic_slow(mq, cx, w, h, template, tpgd, at, skip, deadline, partial) -> Bitmap:
    """AT 画素が既定の位置でないか、飛ばす画素のあるとき。AT の所のビットを、その位置の画素で置き換える。"""
    n0, n1, right1, n2, right2 = _GB[template]
    mask0, mask1, mask2 = (1 << n0) - 1, (1 << n1) - 1, (1 << n2) - 1
    s1, s2 = n0, n0 + n1
    tp = _TPGD_CX[template]
    bits = _AT_BITS[template]
    keep = ~sum(1 << b for b in bits)
    ats = [(a[0], a[1], b) for a, b in zip(at, bits)]
    rows: list = []
    zero = bytearray(w + 8)
    ltp = 0
    dec = mq.decode
    clock = time.monotonic
    for y in range(h):
        if deadline is not None and y & 15 == 0 and clock() > deadline:
            if partial is not None:
                partial.append(y)
            rows += [bytearray(w + 8) for _ in range(h - y)]
            break
        r1 = rows[y - 1] if y >= 1 else zero
        r2 = rows[y - 2] if y >= 2 else zero
        if tpgd:
            ltp ^= dec(cx, tp)
            if ltp:
                rows.append(bytearray(r1))
                continue
        cur = bytearray(w + 8)
        rows.append(cur)
        l1 = 0
        for k in range(right1 + 1):
            l1 = (l1 << 1) | r1[k]
        l2 = 0
        if n2:
            for k in range(right2 + 1):
                l2 = (l2 << 1) | r2[k]
        l0 = 0
        srow = skip.rows[y] if skip is not None else None
        for x in range(w):
            if srow is not None and srow[x]:
                d = 0
            else:
                i = ((l2 << s2) | (l1 << s1) | l0) & keep
                for dx, dy, b in ats:
                    xx, yy = x + dx, y + dy
                    if 0 <= xx < w and yy >= 0 and rows[yy][xx]:
                        i |= 1 << b
                d = dec(cx, i)
                if d:
                    cur[x] = 1
            l0 = ((l0 << 1) | d) & mask0
            l1 = ((l1 << 1) | r1[x + right1 + 1]) & mask1
            l2 = ((l2 << 1) | r2[x + right2 + 1]) & mask2
    return Bitmap(w, h, [r[:w] for r in rows])


# ====================================================================== 細かくする領域（6.3）
_TPGR_CX = (0x0100, 0x0080)


def gr_context_size(template: int) -> int:
    return (1 << 13, 1 << 10)[template]


def decode_refinement(mq: MQ, cx, w: int, h: int, template: int, ref: Bitmap, dx: int, dy: int, tpgr: bool = False, at=None) -> Bitmap:
    """6.3.5。ref：元の画像、(dx, dy)：元の画像のずれ。at：型 0 の AT 画素 2 つ（[今の領域, 元の画像]）。"""
    _check_size(w, h, "細かくする領域")
    a1 = tuple(at[0]) if at else (-1, -1)
    a2 = tuple(at[1]) if at else (-1, -1)
    rows: list = []
    rw, rh = ref.w, ref.h
    rrows = ref.rows
    zero = bytes(rw)

    def R(xx: int, yy: int) -> int:
        return rrows[yy][xx] if 0 <= yy < rh and 0 <= xx < rw else 0

    dec = mq.decode
    ltp = 0
    tp = _TPGR_CX[template]
    nominal = a1 == (-1, -1) and a2 == (-1, -1)
    for y in range(h):
        if tpgr:
            ltp ^= dec(cx, tp)
        prev = rows[y - 1] if y >= 1 else None
        cur = bytearray(w)
        rows.append(cur)
        ry = y - dy
        rm = rrows[ry - 1] if 0 <= ry - 1 < rh else zero
        r0 = rrows[ry] if 0 <= ry < rh else zero
        rp = rrows[ry + 1] if 0 <= ry + 1 < rh else zero
        for x in range(w):
            rx = x - dx
            if 1 <= rx < rw - 1:
                n00, n01, n02 = rm[rx - 1], rm[rx], rm[rx + 1]
                n10, n11, n12 = r0[rx - 1], r0[rx], r0[rx + 1]
                n20, n21, n22 = rp[rx - 1], rp[rx], rp[rx + 1]
            else:
                def g(row, q):
                    return row[q] if 0 <= q < rw else 0

                n00, n01, n02 = g(rm, rx - 1), g(rm, rx), g(rm, rx + 1)
                n10, n11, n12 = g(r0, rx - 1), g(r0, rx), g(r0, rx + 1)
                n20, n21, n22 = g(rp, rx - 1), g(rp, rx), g(rp, rx + 1)
            if ltp:
                # 6.3.5.6 の典型の予測：元の画像の周りの 3 × 3 が同じなら、その値
                t = n00 + n01 + n02 + n10 + n11 + n12 + n20 + n21 + n22
                if t == 0:
                    continue
                if t == 9:
                    cur[x] = 1
                    continue
            left = cur[x - 1] if x else 0
            up = prev[x] if prev is not None else 0
            upr = prev[x + 1] if prev is not None and x + 1 < w else 0
            if template == 0:
                if nominal:
                    ax = prev[x - 1] if prev is not None and x else 0
                    ar = n00
                else:
                    xx, yy = x + a1[0], y + a1[1]
                    ax = rows[yy][xx] if 0 <= xx < w and 0 <= yy <= y else 0
                    ar = R(rx + a2[0], ry + a2[1])
                i = left | upr << 1 | up << 2 | ax << 3 | n22 << 4 | n21 << 5 | n20 << 6 | n12 << 7 | n11 << 8 | n10 << 9 \
                    | n02 << 10 | n01 << 11 | ar << 12
            else:
                upl = prev[x - 1] if prev is not None and x else 0
                i = left | upr << 1 | up << 2 | upl << 3 | n22 << 4 | n21 << 5 | n12 << 6 | n11 << 7 | n10 << 8 | n01 << 9
            if dec(cx, i):
                cur[x] = 1
    return Bitmap(w, h, rows)


# ====================================================================== 整数の算術符号（附属書 A）
def _new_int_cx() -> bytearray:
    return bytearray(512)


def decode_int(mq: MQ, cx) -> int | None:
    """A.2。None は OOB。"""
    dec = mq.decode
    prev = 1

    def bit() -> int:
        nonlocal prev
        d = dec(cx, prev)
        prev = (prev << 1 | d) if prev < 256 else ((prev << 1 | d) & 511) | 256
        return d

    s = bit()
    if not bit():
        n, off = 2, 0
    elif not bit():
        n, off = 4, 4
    elif not bit():
        n, off = 6, 20
    elif not bit():
        n, off = 8, 84
    elif not bit():
        n, off = 12, 340
    else:
        n, off = 32, 4436
    v = 0
    for _ in range(n):
        v = v << 1 | bit()
    v += off
    if not s:
        return v
    return -v if v > 0 else None


def decode_iaid(mq: MQ, cx, length: int) -> int:
    prev = 1
    dec = mq.decode
    for _ in range(length):
        prev = prev << 1 | dec(cx, prev)
    return prev - (1 << length)


# ====================================================================== ビットを読む・ハフマン（附属書 B）
class BitReader:
    __slots__ = ("data", "pos", "end")

    def __init__(self, data: bytes, start: int = 0, end: int | None = None):
        self.data = data
        self.pos = start * 8
        self.end = (len(data) if end is None else end) * 8

    def bit(self) -> int:
        p = self.pos
        if p >= self.end:
            raise Jbig2Error("ハフマンのデータが途中で終わっています")
        self.pos = p + 1
        return (self.data[p >> 3] >> (7 - (p & 7))) & 1

    def bits(self, n: int) -> int:
        v = 0
        for _ in range(n):
            v = v << 1 | self.bit()
        return v

    def align(self) -> None:
        self.pos = (self.pos + 7) & ~7

    @property
    def byte(self) -> int:
        return self.pos >> 3


_LOW, _HIGH, _OOB = "low", "high", "oob"


class Huffman:
    """B.2 の表の行：(PREFLEN, RANGELEN, RANGELOW, 種類)。B.3 のとおり符号を割り当てる（同じ長さは行の順）。"""

    def __init__(self, lines: list):
        self.lines = lines
        self.codes: dict = {}
        maxlen = max((ln[0] for ln in lines), default=0)
        self.maxlen = maxlen
        counts = [0] * (maxlen + 2)
        for ln in lines:
            counts[ln[0]] += 1
        counts[0] = 0
        first = 0
        for cur in range(1, maxlen + 1):
            first = (first + counts[cur - 1]) << 1
            code = first
            for ln in lines:
                if ln[0] == cur:
                    self.codes[(cur, code)] = ln
                    code += 1

    def decode(self, br: BitReader) -> int | None:
        code = 0
        codes = self.codes
        for n in range(1, self.maxlen + 1):
            code = code << 1 | br.bit()
            ln = codes.get((n, code))
            if ln is not None:
                _p, rl, low, kind = ln
                if kind == _OOB:
                    return None
                if kind == _LOW:
                    return low - br.bits(32)
                return low + (br.bits(rl) if rl else 0)
        raise Jbig2Error("ハフマンの符号が表にありません")


def _table(spec: list, oob: bool = False) -> Huffman:
    lines = []
    for item in spec:
        p, r, low = item[:3]
        kind = item[3] if len(item) > 3 else ""
        if p:
            lines.append((p, r, low, kind))
    return Huffman(lines)


_STD_SPEC = {
    1: [(1, 4, 0), (2, 8, 16), (3, 16, 272), (3, 32, 65808, _HIGH)],
    2: [(1, 0, 0), (2, 0, 1), (3, 0, 2), (4, 3, 3), (5, 6, 11), (6, 32, 75, _HIGH), (6, 0, 0, _OOB)],
    3: [(8, 8, -256), (1, 0, 0), (2, 0, 1), (3, 0, 2), (4, 3, 3), (5, 6, 11), (8, 32, -257, _LOW), (7, 32, 75, _HIGH), (6, 0, 0, _OOB)],
    4: [(1, 0, 1), (2, 0, 2), (3, 0, 3), (4, 3, 4), (5, 6, 12), (5, 32, 76, _HIGH)],
    5: [(7, 8, -255), (1, 0, 1), (2, 0, 2), (3, 0, 3), (4, 3, 4), (5, 6, 12), (7, 32, -256, _LOW), (6, 32, 76, _HIGH)],
    6: [(5, 10, -2048), (4, 9, -1024), (4, 8, -512), (4, 7, -256), (5, 6, -128), (5, 5, -64), (4, 5, -32), (2, 7, 0), (3, 7, 128),
        (3, 8, 256), (4, 9, 512), (4, 10, 1024), (6, 32, -2049, _LOW), (6, 32, 2048, _HIGH)],
    7: [(4, 9, -1024), (3, 8, -512), (4, 7, -256), (5, 6, -128), (5, 5, -64), (4, 5, -32), (4, 5, 0), (5, 5, 32), (5, 6, 64),
        (4, 7, 128), (3, 8, 256), (3, 9, 512), (3, 10, 1024), (5, 32, -1025, _LOW), (5, 32, 2048, _HIGH)],
    8: [(8, 3, -15), (9, 1, -7), (8, 1, -5), (9, 0, -3), (7, 0, -2), (4, 0, -1), (2, 1, 0), (5, 0, 2), (6, 0, 3), (3, 4, 4), (6, 1, 20),
        (4, 4, 22), (4, 5, 38), (5, 6, 70), (5, 7, 134), (6, 7, 262), (7, 8, 390), (6, 10, 646), (9, 32, -16, _LOW), (9, 32, 1670, _HIGH),
        (2, 0, 0, _OOB)],
    9: [(8, 4, -31), (9, 2, -15), (8, 2, -11), (9, 1, -7), (7, 1, -5), (4, 1, -3), (3, 1, -1), (3, 1, 1), (5, 1, 3), (6, 1, 5), (3, 5, 7),
        (6, 2, 39), (4, 5, 43), (4, 6, 75), (5, 7, 139), (5, 8, 267), (6, 8, 523), (7, 9, 779), (6, 11, 1291), (9, 32, -32, _LOW),
        (9, 32, 3339, _HIGH), (2, 0, 0, _OOB)],
    10: [(7, 4, -21), (8, 0, -5), (7, 0, -4), (5, 0, -3), (2, 2, -2), (5, 0, 2), (6, 0, 3), (7, 0, 4), (8, 0, 5), (2, 6, 6), (5, 5, 70),
         (6, 5, 102), (6, 6, 134), (6, 7, 198), (6, 8, 326), (6, 9, 582), (6, 10, 1094), (7, 11, 2118), (8, 32, -22, _LOW),
         (8, 32, 4166, _HIGH), (2, 0, 0, _OOB)],
    11: [(1, 0, 1), (2, 1, 2), (4, 0, 4), (4, 1, 5), (5, 1, 7), (5, 2, 9), (6, 2, 13), (7, 2, 17), (7, 3, 21), (7, 4, 29), (7, 5, 45),
         (7, 6, 77), (7, 32, 141, _HIGH)],
    12: [(1, 0, 1), (2, 0, 2), (3, 1, 3), (5, 0, 5), (5, 1, 6), (6, 1, 8), (7, 0, 10), (7, 1, 11), (7, 2, 13), (7, 3, 17), (7, 4, 25),
         (8, 5, 41), (8, 32, 73, _HIGH)],
    13: [(1, 0, 1), (3, 0, 2), (4, 0, 3), (5, 0, 4), (4, 1, 5), (3, 3, 7), (6, 1, 15), (6, 2, 17), (6, 3, 21), (6, 4, 29), (6, 5, 45),
         (7, 6, 77), (7, 32, 141, _HIGH)],
    14: [(3, 0, -2), (3, 0, -1), (1, 0, 0), (3, 0, 1), (3, 0, 2)],
    15: [(7, 4, -24), (6, 2, -8), (5, 1, -4), (4, 0, -2), (3, 0, -1), (1, 0, 0), (3, 0, 1), (4, 0, 2), (5, 1, 3), (6, 2, 5), (7, 4, 9),
         (7, 32, -25, _LOW), (7, 32, 25, _HIGH)],
}
_STD_CACHE: dict = {}


def std_table(n: int) -> Huffman:
    t = _STD_CACHE.get(n)
    if t is None:
        t = _STD_CACHE[n] = _table(_STD_SPEC[n])
    return t


def custom_table(data: bytes) -> Huffman:
    """表の区切り（B.2）。"""
    flags = data[0]
    oob = flags & 1
    ps = ((flags >> 1) & 7) + 1
    rs = ((flags >> 4) & 7) + 1
    low, high = struct.unpack_from(">ii", data, 1)
    br = BitReader(data, 9)
    lines = []
    cur = low
    while cur < high:
        p = br.bits(ps)
        r = br.bits(rs)
        lines.append((p, r, cur, ""))
        cur += 1 << r
    p = br.bits(ps)
    lines.append((p, 32, low - 1, _LOW))
    p = br.bits(ps)
    lines.append((p, 32, high, _HIGH))
    if oob:
        p = br.bits(ps)
        lines.append((p, 0, 0, _OOB))
    return Huffman([ln for ln in lines if ln[0]])


# ====================================================================== MMR（T.6。汎用領域の MMR、6.2.6）
def decode_mmr(data: bytes, start: int, end: int, w: int, h: int) -> tuple[Bitmap, int]:
    """(画像, 読み終えた次のバイトの位置)。1 が黒。行の変わり目を T.6 のとおりにたどる（pdf/raster.py の部品を使う）。"""
    from ..pdf.raster import RasterError, _Bits, _g4_line

    _check_size(w, h, "MMR の領域")
    bits = _Bits(data[start:end])
    rows = []
    ref = [w, w]
    for _y in range(h):
        if bits.pos >= bits.n or bits.peek(24) == 0x001001:
            break
        try:
            cur = _g4_line(bits, ref, w)
        except RasterError as e:
            raise Jbig2Error(f"MMR の符号が読めません：{e}") from None
        row = bytearray(w)
        for i in range(0, len(cur) - 1, 2):
            a, b = min(cur[i], w), min(cur[i + 1], w)
            if b > a:
                row[a:b] = b"\x01" * (b - a)
        rows.append(row)
        ref = cur + [w, w]
    while len(rows) < h:
        rows.append(bytearray(w))
    if bits.pos + 24 <= bits.n and bits.peek(24) == 0x001001:
        bits.pos += 24
    return Bitmap(w, h, rows), start + ((bits.pos + 7) >> 3)


# ====================================================================== 字の領域（6.4）
class _TextParams:
    __slots__ = ("huff", "refine", "log_strips", "ref_corner", "transposed", "comb_op", "def_pixel", "ds_offset", "r_template", "r_at",
                 "fs", "ds", "dt", "rdw", "rdh", "rdx", "rdy", "rsize", "sym_code", "sym_code_len", "n_inst")


def _new_text_cx(code_len: int, r_template: int) -> dict:
    return {"dt": _new_int_cx(), "fs": _new_int_cx(), "ds": _new_int_cx(), "it": _new_int_cx(), "ri": _new_int_cx(),
            "rdw": _new_int_cx(), "rdh": _new_int_cx(), "rdx": _new_int_cx(), "rdy": _new_int_cx(),
            "id": bytearray(1 << (code_len + 1)), "gr": bytearray(gr_context_size(r_template))}


def decode_text_region(w: int, h: int, tp: _TextParams, syms: list, mq: MQ | None, br: BitReader | None, cx: dict | None,
                       data: bytes = b"", deadline: float | None = None) -> Bitmap:
    _check_size(w, h, "字の領域")
    out = Bitmap(w, h, fill=tp.def_pixel)
    strips = 1 << tp.log_strips
    huff = tp.huff
    if huff:
        t = tp.dt.decode(br)
    else:
        t = decode_int(mq, cx["dt"])
    stript = -(t or 0) * strips
    firsts = 0
    n = 0
    total = tp.n_inst
    nsyms = len(syms)
    clock = time.monotonic
    while n < total:
        if deadline is not None and clock() > deadline:
            raise _Stopped()
        dt = tp.dt.decode(br) if huff else decode_int(mq, cx["dt"])
        if dt is None:
            raise Jbig2Error("字の領域：行の位置が読めません")
        stript += dt * strips
        first = True
        curs = 0
        while True:
            if first:
                dfs = tp.fs.decode(br) if huff else decode_int(mq, cx["fs"])
                if dfs is None:
                    raise Jbig2Error("字の領域：行の頭の位置が読めません")
                firsts += dfs
                curs = firsts
                first = False
            else:
                ids = tp.ds.decode(br) if huff else decode_int(mq, cx["ds"])
                if ids is None:
                    break  # OOB：行の終わり
                curs += ids + tp.ds_offset
            if n >= total:
                break  # 仕様では OOB で行が終わる（数を超えて字が続くのは壊れたデータ）
            if strips == 1:
                curt = 0
            elif huff:
                curt = br.bits(tp.log_strips)
            else:
                curt = decode_int(mq, cx["it"]) or 0
            ti = stript + curt
            if huff:
                sid = tp.sym_code.decode(br)
            else:
                sid = decode_iaid(mq, cx["id"], tp.sym_code_len)
            if sid is None or not 0 <= sid < nsyms:
                raise Jbig2Error(f"字の領域：記号の番号 {sid} が辞書（{nsyms} 個）にありません")
            ri = 0
            if tp.refine:
                ri = br.bit() if huff else (decode_int(mq, cx["ri"]) or 0)
            ib = syms[sid]
            if ri:
                ib = _refine_instance(ib, tp, mq, br, cx, data)
            wi, hi = ib.w, ib.h
            corner, trans = tp.ref_corner, tp.transposed
            if not trans and corner >= 2:  # 右上・右下
                curs += wi - 1
            elif trans and not corner & 1:  # 左下・右下
                curs += hi - 1
            si = curs
            if not trans:
                x = si - wi + 1 if corner >= 2 else si
                y = ti if corner & 1 else ti - hi + 1
            else:
                x = ti - wi + 1 if corner >= 2 else ti
                y = si if corner & 1 else si - hi + 1
            compose(out, ib, x, y, tp.comb_op)
            if not trans and corner < 2:
                curs += wi - 1
            elif trans and corner & 1:
                curs += hi - 1
            n += 1
    return out


def _refine_instance(ib: Bitmap, tp: _TextParams, mq, br, cx, data) -> Bitmap:
    """6.4.11：細かくする字の形。"""
    huff = tp.huff
    if huff:
        rdw, rdh, rdx, rdy = tp.rdw.decode(br), tp.rdh.decode(br), tp.rdx.decode(br), tp.rdy.decode(br)
        rsize = tp.rsize.decode(br)
        br.align()
        start = br.byte
        m = MQ(data, start, start + rsize if rsize else None)
    else:
        rdw = decode_int(mq, cx["rdw"])
        rdh = decode_int(mq, cx["rdh"])
        rdx = decode_int(mq, cx["rdx"])
        rdy = decode_int(mq, cx["rdy"])
        m = mq
    if None in (rdw, rdh, rdx, rdy):
        raise Jbig2Error("字の領域：細かくする大きさが読めません")
    out = decode_refinement(m, cx["gr"], ib.w + rdw, ib.h + rdh, tp.r_template, ib, (rdw >> 1) + rdx, (rdh >> 1) + rdy, False, tp.r_at)
    if huff:
        br.pos = (start + rsize) * 8
    return out


# ====================================================================== 記号の辞書（6.5）
class SymbolDict:
    __slots__ = ("symbols", "gb", "gr")

    def __init__(self, symbols: list, gb=None, gr=None):
        self.symbols = symbols
        self.gb = gb
        self.gr = gr


def _ceil_log2(n: int) -> int:
    k = 0
    while (1 << k) < n:
        k += 1
    return k


def decode_symbol_dict(data: bytes, insyms: list, tables: list, retained: SymbolDict | None, deadline: float | None) -> SymbolDict:
    flags = struct.unpack_from(">H", data, 0)[0]
    pos = 2
    huff = flags & 1
    refagg = (flags >> 1) & 1
    sel_dh, sel_dw, sel_bm, sel_agg = (flags >> 2) & 3, (flags >> 4) & 3, (flags >> 6) & 1, (flags >> 7) & 1
    ctx_used, ctx_retained = (flags >> 8) & 1, (flags >> 9) & 1
    template = (flags >> 10) & 3
    r_template = (flags >> 12) & 1
    at = None
    if not huff:
        k = 4 if template == 0 else 1
        at = [struct.unpack_from(">bb", data, pos + 2 * i) for i in range(k)]
        pos += 2 * k
    r_at = None
    if refagg and not r_template:
        r_at = [struct.unpack_from(">bb", data, pos), struct.unpack_from(">bb", data, pos + 2)]
        pos += 4
    n_ex, n_new = struct.unpack_from(">II", data, pos)
    pos += 8
    if n_new > 1 << 24 or n_ex > (1 << 24) + len(insyms):
        raise LimitExceeded(f"JBIG2 の記号の辞書の記号の数（{n_new:,}）が多すぎます")
    tabs = list(tables)

    def pick(sel, choices):
        if sel >= len(choices) or choices[sel] is None:
            if not tabs:
                raise Jbig2Error("記号の辞書：ハフマンの表の区切りが足りません")
            return tabs.pop(0)
        return std_table(choices[sel])

    t_dh = t_dw = t_bm = t_agg = None
    if huff:
        t_dh = pick(sel_dh, (4, 5, None, None))
        t_dw = pick(sel_dw, (2, 3, None, None))
        t_bm = pick(sel_bm, (1, None))
        t_agg = pick(sel_agg, (1, None))
    gb = bytearray(gb_context_size(template)) if not huff else None
    gr = bytearray(gr_context_size(r_template)) if refagg else None
    if ctx_used and retained is not None:
        if retained.gb is not None and gb is not None and len(retained.gb) == len(gb):
            gb = bytearray(retained.gb)
        if retained.gr is not None and gr is not None and len(retained.gr) == len(gr):
            gr = bytearray(retained.gr)
    code_len = _ceil_log2(len(insyms) + n_new)
    if huff:
        code_len = max(code_len, 1)
    mq = br = None
    if huff:
        br = BitReader(data, pos)
    else:
        mq = MQ(data, pos)
    iadh, iadw, iaex, iaai = _new_int_cx(), _new_int_cx(), _new_int_cx(), _new_int_cx()
    tcx = _new_text_cx(code_len, r_template) if refagg else None
    if tcx is not None:
        tcx["gr"] = gr
    syms = list(insyms)
    new: list = []
    widths: list = []
    hc = 0
    clock = time.monotonic
    budget_px = 0
    while len(new) < n_new:
        if deadline is not None and clock() > deadline:
            raise _Stopped()
        dh = t_dh.decode(br) if huff else decode_int(mq, iadh)
        if dh is None:
            raise Jbig2Error("記号の辞書：高さが読めません")
        hc += dh
        if hc < 0 or hc > 65535:
            raise Jbig2Error("記号の辞書：高さが範囲の外です")
        symw = 0
        totw = 0
        first_of_class = len(new)
        while True:
            dw = t_dw.decode(br) if huff else decode_int(mq, iadw)
            if dw is None:
                break
            if len(new) >= n_new:
                raise Jbig2Error("記号の辞書：記号の数が宣言を超えます")
            symw += dw
            if symw < 0 or symw > 65535:
                raise Jbig2Error("記号の辞書：幅が範囲の外です")
            totw += symw
            budget_px += symw * hc
            if budget_px > MAX_PIXELS:
                raise LimitExceeded("JBIG2 の記号の辞書の画素の合計が大きすぎます")
            if not huff or refagg:
                if not refagg:
                    bm = decode_generic(mq, gb, symw, hc, template, False, at)
                else:
                    bm = _agg_symbol(symw, hc, syms, huff, mq, br, tcx, iaai, t_agg, r_template, r_at, code_len, data, deadline)
                new.append(bm)
                syms.append(bm)
            else:
                widths.append(symw)
                new.append(None)
                syms.append(None)
        if huff and not refagg:
            size = t_bm.decode(br)
            br.align()
            start = br.byte
            if size == 0:
                stride = (totw + 7) >> 3
                rows = []
                for j in range(hc):
                    row = data[start + j * stride:start + (j + 1) * stride]
                    rows.append(_unpack_row(row, totw))
                coll = Bitmap(totw, hc, rows)
                br.pos = (start + stride * hc) * 8
            else:
                coll, _e = decode_mmr(data, start, start + size, totw, hc)
                br.pos = (start + size) * 8
            x = 0
            for i in range(first_of_class, len(new)):
                sw = widths[i]
                bm = Bitmap(sw, hc, [r[x:x + sw] for r in coll.rows])
                new[i] = bm
                syms[len(insyms) + i] = bm
                x += sw
        if not huff or refagg:
            widths = []
    # 7.4.2.1.6 / 6.5.10 書き出しの印
    flagsx = []
    cur = 0
    total = len(insyms) + n_new
    while len(flagsx) < total:
        run = std_table(1).decode(br) if huff else decode_int(mq, iaex)
        if run is None or run < 0 or len(flagsx) + run > total or (run == 0 and len(flagsx) >= total):
            raise Jbig2Error("記号の辞書：書き出しの印が読めません")
        flagsx += [cur] * run
        cur ^= 1
    exported = [s for s, f in zip(syms, flagsx) if f]
    return SymbolDict(exported, gb if ctx_retained else None, gr if ctx_retained else None)


def _unpack_row(row: bytes, w: int) -> bytearray:
    out = bytearray()
    for b in row:
        out += _BYTE_BITS[b]
    return out[:w].ljust(w, b"\0")


_BYTE_BITS = [bytes((b >> (7 - i)) & 1 for i in range(8)) for b in range(256)]


def _agg_symbol(w, h, syms, huff, mq, br, tcx, iaai, t_agg, r_template, r_at, code_len, data, deadline) -> Bitmap:
    """6.5.8.2：既にある記号を細かくするか、集めて、新しい記号を作る。"""
    ninst = t_agg.decode(br) if huff else decode_int(mq, iaai)
    if ninst is None or ninst < 1:
        raise Jbig2Error("記号の辞書：集める数が読めません")
    if ninst > 1:
        tp = _TextParams()
        tp.huff, tp.refine, tp.log_strips, tp.ref_corner, tp.transposed = bool(huff), 1, 0, 1, 0
        tp.comb_op, tp.def_pixel, tp.ds_offset, tp.r_template, tp.r_at = 0, 0, 0, r_template, r_at
        tp.sym_code_len, tp.n_inst = code_len, ninst
        if huff:
            tp.fs, tp.ds, tp.dt = std_table(6), std_table(8), std_table(11)
            tp.rdw = tp.rdh = tp.rdx = tp.rdy = std_table(15)
            tp.rsize = std_table(1)
            tp.sym_code = _FixedCode(code_len)
        else:
            tp.fs = tp.ds = tp.dt = tp.rdw = tp.rdh = tp.rdx = tp.rdy = tp.rsize = tp.sym_code = None
        return decode_text_region(w, h, tp, syms, mq, br, tcx, data, deadline)
    if huff:
        sid = br.bits(code_len)
        rdx = std_table(15).decode(br)
        rdy = std_table(15).decode(br)
        rsize = std_table(1).decode(br)
        br.align()
        start = br.byte
        m = MQ(data, start, start + rsize if rsize else None)
    else:
        sid = decode_iaid(mq, tcx["id"], code_len)
        rdx = decode_int(mq, tcx["rdx"])
        rdy = decode_int(mq, tcx["rdy"])
        m = mq
    if sid >= len(syms) or syms[sid] is None:
        raise Jbig2Error(f"記号の辞書：細かくする元の記号 {sid} がありません")
    out = decode_refinement(m, tcx["gr"], w, h, r_template, syms[sid], rdx or 0, rdy or 0, False, r_at)
    if huff:
        br.pos = (start + rsize) * 8
    return out


class _FixedCode:
    """固定の長さの記号の番号（6.5.8.2.3）。"""

    def __init__(self, n: int):
        self.n = n

    def decode(self, br: BitReader) -> int:
        return br.bits(self.n)


def _symbol_id_table(br: BitReader, nsyms: int) -> Huffman:
    """7.4.3.1.7：記号の番号のハフマン符号。"""
    run_lines = [(br.bits(4), 0, i, "") for i in range(35)]
    run = Huffman([ln for ln in run_lines if ln[0]])
    lens: list[int] = []
    while len(lens) < nsyms:
        code = run.decode(br)
        if code is None:
            raise Jbig2Error("字の領域：記号の符号の長さが読めません")
        if code < 32:
            lens.append(code)
        elif code == 32:
            if not lens:
                raise Jbig2Error("字の領域：記号の符号の長さの繰り返しの前に長さがありません")
            lens += [lens[-1]] * (br.bits(2) + 3)
        elif code == 33:
            lens += [0] * (br.bits(3) + 3)
        else:
            lens += [0] * (br.bits(7) + 11)
    br.align()
    return Huffman([(ln, 0, i, "") for i, ln in enumerate(lens[:nsyms]) if ln])


# ====================================================================== 模様・網点（6.7・6.6・附属書 C）
def decode_pattern_dict(data: bytes) -> list:
    flags, pw, ph = data[0], data[1], data[2]
    graymax = struct.unpack_from(">I", data, 3)[0]
    mmr = flags & 1
    template = (flags >> 1) & 3
    w = (graymax + 1) * pw
    _check_size(w, ph, "模様の辞書")
    if mmr:
        coll, _e = decode_mmr(data, 7, len(data), w, ph)
    else:
        at = [(-pw, 0)] + ([(-3, -1), (2, -2), (-2, -2)] if template == 0 else [])
        coll = decode_generic(MQ(data, 7), bytearray(gb_context_size(template)), w, ph, template, False, at)
    return [Bitmap(pw, ph, [r[g * pw:(g + 1) * pw] for r in coll.rows]) for g in range(graymax + 1)]


def decode_halftone(data: bytes, w: int, h: int, patterns: list, deadline: float | None = None) -> Bitmap:
    flags = data[0]
    mmr = flags & 1
    template = (flags >> 1) & 3
    enable_skip = (flags >> 3) & 1
    op = (flags >> 4) & 7
    def_pixel = (flags >> 7) & 1
    gw, gh, gx, gy, rx, ry = struct.unpack_from(">IIiiHH", data, 1)
    pos = 21
    _check_size(gw, gh, "網点の格子")
    if not patterns:
        raise Jbig2Error("網点の領域：模様の辞書がありません")
    pw, ph = patterns[0].w, patterns[0].h
    bpp = max(1, _ceil_log2(len(patterns))) if len(patterns) > 1 else 0
    skip = None
    if enable_skip:
        skip = Bitmap(gw, gh)
        for mg in range(gh):
            row = skip.rows[mg]
            for ng in range(gw):
                x = (gx + mg * ry + ng * rx) >> 8
                y = (gy + mg * rx - ng * ry) >> 8
                if x + pw <= 0 or x >= w or y + ph <= 0 or y >= h:
                    row[ng] = 1
    # C.5：上の桁から順に、灰色符号のビットの面
    at = [(3 if template <= 1 else 2, -1)] + ([(-3, -1), (2, -2), (-2, -2)] if template == 0 else [])
    mq = MQ(data, pos) if not mmr else None
    cx = bytearray(gb_context_size(template)) if not mmr else None
    gray = [[0] * gw for _ in range(gh)]
    prev = None
    p = pos
    for _j in range(bpp - 1, -1, -1):
        if deadline is not None and time.monotonic() > deadline:
            raise _Stopped()
        if mmr:
            plane, p = decode_mmr(data, p, len(data), gw, gh)
        else:
            plane = decode_generic(mq, cx, gw, gh, template, False, at, skip, deadline)
        if prev is not None:
            frm = int.from_bytes
            for row, up in zip(plane.rows, prev.rows):
                row[:] = (frm(row, "big") ^ frm(up, "big")).to_bytes(gw, "big")
        prev = plane
        gray = [[g << 1 | b for g, b in zip(gr, row)] for gr, row in zip(gray, plane.rows)]
    out = Bitmap(w, h, fill=def_pixel)
    n = len(patterns)
    for mg, gr in enumerate(gray):
        if deadline is not None and mg & 15 == 0 and time.monotonic() > deadline:
            raise _Stopped()
        for ng, g in enumerate(gr):
            x = (gx + mg * ry + ng * rx) >> 8
            y = (gy + mg * rx - ng * ry) >> 8
            if x + pw <= 0 or x >= w or y + ph <= 0 or y >= h:
                continue
            compose(out, patterns[g if g < n else n - 1], x, y, op)
    return out


# ====================================================================== 区切り（7.2）
class _Segment:
    __slots__ = ("number", "type", "refs", "page", "length", "data")


def _parse_segments(data: bytes) -> list:
    """区切りの見出しとデータ。ファイルの編成（連続・ランダムアクセス、附属書 D）の頭があれば読み飛ばす。"""
    pos = 0
    random_access = False
    if data[:8] == b"\x97JB2\r\n\x1a\n":
        fl = data[8]
        random_access = not (fl & 1)
        pos = 9 + (0 if fl & 2 else 4)
    segs: list = []
    n = len(data)
    pending: list = []
    while pos + 11 <= n:
        sg = _Segment()
        sg.number = struct.unpack_from(">I", data, pos)[0]
        flags = data[pos + 4]
        sg.type = flags & 63
        pos += 5
        cnt = data[pos] >> 5
        if cnt == 7:
            cnt = struct.unpack_from(">I", data, pos)[0] & 0x1FFFFFFF
            pos += 4 + ((cnt + 8) >> 3)
        else:
            pos += 1
        if cnt > 1 << 20:
            raise Jbig2Error("区切りの参照の数が多すぎます")
        size = 1 if sg.number <= 256 else 2 if sg.number <= 65536 else 4
        fmt = {1: ">B", 2: ">H", 4: ">I"}[size]
        refs = []
        for _ in range(cnt):
            refs.append(struct.unpack_from(fmt, data, pos)[0])
            pos += size
        sg.refs = refs
        if flags & 0x40:
            sg.page = struct.unpack_from(">I", data, pos)[0]
            pos += 4
        else:
            sg.page = data[pos]
            pos += 1
        sg.length = struct.unpack_from(">I", data, pos)[0]
        pos += 4
        if random_access:
            pending.append(sg)
            segs.append(sg)
            sg.data = b""
            if sg.type == 51:
                break
            continue
        if sg.length == 0xFFFFFFFF:
            sg.length = _unknown_length(data, pos, sg)
        sg.data = data[pos:pos + sg.length]
        pos += sg.length
        segs.append(sg)
        if sg.type == 51:
            break
    if random_access:
        for sg in pending:
            if sg.length == 0xFFFFFFFF:
                sg.length = _unknown_length(data, pos, sg)
            sg.data = data[pos:pos + sg.length]
            pos += sg.length
    return segs


def _unknown_length(data: bytes, pos: int, sg) -> int:
    """7.2.7：長さの分からない汎用領域：終わりの印（算術 FF AC、MMR 00 00）と、その後の 4 バイトの行の数まで。"""
    if sg.type not in (36, 38, 39):
        raise Jbig2Error("長さの分からない区切りは、汎用領域だけです")
    mmr = data[pos + 17] & 1
    mark = b"\x00\x00" if mmr else b"\xff\xac"
    i = data.find(mark, pos + 18)
    if i < 0:
        return len(data) - pos
    return i + 2 + 4 - pos


# ====================================================================== 全体
class Page:
    """開いたページ。px：幅 × 高さのバイト（0＝黒、255＝白）。problems：描けなかった所（なければ空）。"""

    __slots__ = ("w", "h", "px", "problems", "xres", "yres", "stopped")

    def __init__(self, w, h, px, problems, xres=0, yres=0, stopped=False):
        self.w, self.h, self.px, self.problems, self.xres, self.yres, self.stopped = w, h, px, problems, xres, yres, stopped


_TO_GRAY = bytes([255, 0]) + bytes(254)


def _region_info(d: bytes) -> tuple:
    w, h, x, y, f = struct.unpack_from(">IIIIB", d, 0)
    return w, h, x, y, f & 7


def decode(data: bytes, globals_data: bytes | None = None, width: int | None = None, height: int | None = None,
           page_number: int | None = None, deadline: float | None = None) -> Page:
    """JBIG2 のデータ（と JBIG2Globals）を開いて、1 ページの画像を返す。width・height：PDF の画像の辞書の大きさ（ページの情報の
    高さが不明（ストライプ）のときと、ページの情報がないときに使う）。page_number：省くと、データの最初のページ。"""
    segs = []
    if globals_data:
        segs += _parse_segments(globals_data)
    segs += _parse_segments(data)
    if not segs:
        raise Jbig2Error("JBIG2 の区切りがありません")
    if page_number is None:
        page_number = next((s.page for s in segs if s.type == 48), 1)
    store: dict = {}  # 区切りの番号 → 記号の辞書・模様の辞書・ハフマンの表・中間の領域
    page: Bitmap | None = None
    problems: list[str] = []
    stopped = False
    xres = yres = 0
    grow = False  # 高さの宣言のないページ

    def ensure(hh: int) -> None:
        nonlocal page
        if page is not None and grow and hh > page.h:
            page.rows += [bytearray(page.w) for _ in range(hh - page.h)]
            page.h = hh

    for sg in segs:
        t = sg.type
        if sg.page not in (0, page_number):
            continue
        d = sg.data
        try:
            if t == 48:
                if page is not None:
                    continue
                pw, ph, xres, yres, fl, _st = struct.unpack_from(">IIIIBH", d, 0)
                if ph == 0xFFFFFFFF:
                    ph = height or 0
                    grow = True
                _check_size(pw, ph, "ページ")
                page = Bitmap(pw, ph, fill=(fl >> 2) & 1)
            elif t == 0:
                store[sg.number] = _symbol_dict(sg, d, store, deadline)
            elif t in (4, 6, 7):
                bm, info = _text_segment(d, sg, store, deadline)
                page = _place(page, store, sg, t, 4, bm, info, width, height, ensure)
            elif t == 16:
                store[sg.number] = decode_pattern_dict(d)
            elif t in (20, 22, 23):
                info = _region_info(d)
                pats = next((store[r] for r in sg.refs if isinstance(store.get(r), list)), [])
                bm = decode_halftone(d[17:], info[0], info[1], pats, deadline)
                page = _place(page, store, sg, t, 20, bm, info, width, height, ensure)
            elif t in (36, 38, 39):
                bm, info, part = _generic_segment(d, deadline)
                if part:
                    stopped = True
                page = _place(page, store, sg, t, 36, bm, info, width, height, ensure)
                if part:
                    break
            elif t in (40, 42, 43):
                bm, info = _refinement_segment(d, sg, store, page)
                page = _place(page, store, sg, t, 40, bm, info, width, height, ensure)
            elif t == 49:
                if sg.page == page_number:
                    break
            elif t == 53:
                store[sg.number] = custom_table(d)
            elif t in (50, 51, 52, 62):
                pass
            else:
                problems.append(f"区切りの種類 {t} は扱えません（番号 {sg.number}）")
        except _Stopped:
            stopped = True
            break
        except Jbig2Error as e:
            problems.append(f"区切り {sg.number}（種類 {t}）：{e}")
        except (IndexError, struct.error) as e:
            problems.append(f"区切り {sg.number}（種類 {t}）を読めません（{type(e).__name__}）")
    if stopped:
        problems.append("時間がかかりすぎるので、途中で止めました（描いていない所があります）")
    if page is None:
        if not problems:
            raise Jbig2Error("JBIG2：ページの区切りがありません")
        raise Jbig2Error("JBIG2：" + "／".join(problems[:3]))
    px = b"".join(page.rows).translate(_TO_GRAY)
    return Page(page.w, page.h, px, problems, xres, yres, stopped)


def _place(page, store, sg, t, base, bm, info, width, height, ensure):
    """領域を、中間（base＋0：控える）か、ページ（base＋2・＋3：重ねる）へ。"""
    if t == base:
        store[sg.number] = bm
        return page
    if page is None:
        page = Bitmap(width or (info[0] + info[2]), height or (info[1] + info[3]))
    ensure(info[3] + info[1])
    compose(page, bm, info[2], info[3], info[4])
    return page


def _generic_segment(d: bytes, deadline) -> tuple:
    info = _region_info(d)
    w, h = info[0], info[1]
    fl = d[17]
    mmr, template, tpgd, ext = fl & 1, (fl >> 1) & 3, (fl >> 3) & 1, (fl >> 4) & 1
    pos = 18
    if ext:
        raise Jbig2Error("12 個の AT 画素の型（T.88 の追補 2）は扱えません")
    part: list = []
    if mmr:
        bm, _e = decode_mmr(d, pos, len(d), w, h)
    else:
        k = 4 if template == 0 else 1
        at = [struct.unpack_from(">bb", d, pos + 2 * i) for i in range(k)]
        pos += 2 * k
        bm = decode_generic(MQ(d, pos), bytearray(gb_context_size(template)), w, h, template, bool(tpgd), at, None, deadline, part)
    return bm, info, bool(part)


def _refinement_segment(d: bytes, sg, store: dict, page) -> tuple:
    info = _region_info(d)
    w, h, x, y = info[:4]
    fl = d[17]
    template, tpgr = fl & 1, (fl >> 1) & 1
    pos = 18
    at = None
    if not template:
        at = [struct.unpack_from(">bb", d, pos), struct.unpack_from(">bb", d, pos + 2)]
        pos += 4
    ref = next((store[r] for r in sg.refs if isinstance(store.get(r), Bitmap)), None)
    if ref is None:
        if page is None:
            raise Jbig2Error("細かくする領域：元の画像がありません")
        ref = Bitmap(w, h, [bytearray(r[x:x + w].ljust(w, b"\0")) for r in page.rows[y:y + h]])
        ref.rows += [bytearray(w) for _ in range(h - len(ref.rows))]
    bm = decode_refinement(MQ(d, pos), bytearray(gr_context_size(template)), w, h, template, ref, 0, 0, bool(tpgr), at)
    return bm, info


def _symbol_dict(sg, d: bytes, store: dict, deadline) -> SymbolDict:
    insyms: list = []
    tables: list = []
    last = None
    for r in sg.refs:
        v = store.get(r)
        if isinstance(v, SymbolDict):
            insyms += v.symbols
            last = v
        elif isinstance(v, Huffman):
            tables.append(v)
    return decode_symbol_dict(d, insyms, tables, last, deadline)


def _text_segment(d: bytes, sg, store: dict, deadline) -> tuple:
    info = _region_info(d)
    w, h = info[0], info[1]
    fl = struct.unpack_from(">H", d, 17)[0]
    pos = 19
    tp = _TextParams()
    tp.huff = bool(fl & 1)
    tp.refine = (fl >> 1) & 1
    tp.log_strips = (fl >> 2) & 3
    tp.ref_corner = (fl >> 4) & 3
    tp.transposed = (fl >> 6) & 1
    tp.comb_op = (fl >> 7) & 3
    tp.def_pixel = (fl >> 9) & 1
    ds = (fl >> 10) & 31
    tp.ds_offset = ds - 32 if ds > 15 else ds
    tp.r_template = (fl >> 15) & 1
    hf = 0
    if tp.huff:
        hf = struct.unpack_from(">H", d, pos)[0]
        pos += 2
    tp.r_at = None
    if tp.refine and not tp.r_template:
        tp.r_at = [struct.unpack_from(">bb", d, pos), struct.unpack_from(">bb", d, pos + 2)]
        pos += 4
    tp.n_inst = struct.unpack_from(">I", d, pos)[0]
    pos += 4
    syms: list = []
    tables: list = []
    for r in sg.refs:
        v = store.get(r)
        if isinstance(v, SymbolDict):
            syms += v.symbols
        elif isinstance(v, Huffman):
            tables.append(v)
    code_len = _ceil_log2(len(syms))
    tp.sym_code_len = code_len
    tp.fs = tp.ds = tp.dt = tp.rdw = tp.rdh = tp.rdx = tp.rdy = tp.rsize = tp.sym_code = None
    mq = br = cx = None
    if tp.huff:
        def pick(sel, choices):
            if sel >= len(choices) or choices[sel] is None:
                if not tables:
                    raise Jbig2Error("字の領域：ハフマンの表の区切りが足りません")
                return tables.pop(0)
            return std_table(choices[sel])

        tp.fs = pick(hf & 3, (6, 7, None, None))
        tp.ds = pick((hf >> 2) & 3, (8, 9, 10, None))
        tp.dt = pick((hf >> 4) & 3, (11, 12, 13, None))
        tp.rdw = pick((hf >> 6) & 3, (14, 15, None, None))
        tp.rdh = pick((hf >> 8) & 3, (14, 15, None, None))
        tp.rdx = pick((hf >> 10) & 3, (14, 15, None, None))
        tp.rdy = pick((hf >> 12) & 3, (14, 15, None, None))
        tp.rsize = pick((hf >> 14) & 1, (1, None))
        br = BitReader(d, pos)
        tp.sym_code = _symbol_id_table(br, len(syms))
        cx = {"gr": bytearray(gr_context_size(tp.r_template))}
    else:
        mq = MQ(d, pos)
        cx = _new_text_cx(code_len, tp.r_template)
    bm = decode_text_region(w, h, tp, syms, mq, br, cx, d, deadline)
    return bm, info
