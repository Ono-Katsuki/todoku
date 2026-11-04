"""JPEG 2000 の符号ブロックの復号（ISO/IEC 15444-1 附属書 D：EBCOT の第 1 層）。純粋な Python で、いちばん時間のかかる所。

1 つの係数の印（flags）に、隣の 8 つの有意・上下左右の符号・自分の有意／見た／細かくした、を持つ（表で文脈を引く）。
1 判定ごとに関数を呼ぶと 2 倍ほど遅いので、MQ の復号の手順を、3 つの通り道（有意の伝播・細かくする・片付け）の中に書き写す。
Python には字句のマクロがないので、ソースの字を組み立てて compile する（下の _expand：@MQ(文脈, 結果) の行を手順に置き換える）。

符号ブロックの様式（表 A.19）：BYPASS（生のビット）・RESET・TERMALL・VCAUSAL・PTERM（読むだけなら扱いは同じ）・SEGSYM。
"""

from __future__ import annotations

import re

from .mq import NL, NM, QE, QS

BYPASS, RESET, TERMALL, VCAUSAL, PTERM, SEGSYM = 1, 2, 4, 8, 16, 32

# 印のビット
F_N, F_S, F_E, F_W, F_NE, F_NW, F_SE, F_SW = 1, 2, 4, 8, 0x10, 0x20, 0x40, 0x80
S_N, S_S, S_E, S_W = 0x100, 0x200, 0x400, 0x800
SIG, REFINED, VISIT, NEG = 0x1000, 0x2000, 0x4000, 0x8000
VSC_MASK = ~(F_S | F_SE | F_SW | S_S)  # VCAUSAL：次のストライプの係数は見ない

CTX_RL, CTX_UNI = 17, 18


def _zc_table(orient: int) -> bytes:
    """有意の文脈（表 D.1）。orient：0 LL・1 HL・2 LH・3 HH。添字は印の下位 8 ビット。"""
    out = bytearray(256)
    for f in range(256):
        h = bool(f & F_E) + bool(f & F_W)
        v = bool(f & F_N) + bool(f & F_S)
        d = bool(f & F_NE) + bool(f & F_NW) + bool(f & F_SE) + bool(f & F_SW)
        if orient == 3:
            hv = h + v
            n = (8 if d >= 3 else (7 if hv >= 1 else 6) if d == 2 else (5 if hv >= 2 else 4 if hv == 1 else 3) if d == 1
                 else (2 if hv >= 2 else 1 if hv == 1 else 0))
        else:
            if orient == 1:
                h, v = v, h
            n = (8 if h == 2 else (7 if v >= 1 else 6 if d >= 1 else 5) if h == 1
                 else (4 if v == 2 else 3 if v == 1 else 2 if d >= 2 else 1 if d == 1 else 0))
        out[f] = n
    return bytes(out)


ZC = [_zc_table(o) for o in range(4)]


def _sc_tables() -> tuple[bytes, bytes]:
    """符号の文脈と反転（表 D.2・D.3）。添字は 下位 4 ビット＝N S E W の有意、上位 4 ビット＝N S E W の符号（1 が負）。"""
    ctx, xr = bytearray(256), bytearray(256)
    for i in range(256):
        sig, neg = i & 15, i >> 4

        def part(s, n):
            return (0 if not sig & s else -1 if neg & s else 1)

        h = max(-1, min(1, part(F_E, 0) + part(F_W, 0)))
        v = max(-1, min(1, part(F_N, 0) + part(F_S, 0)))
        flip = 0
        if h < 0 or (h == 0 and v < 0):
            h, v = -h, -v
            flip = 1
        ctx[i] = 9 + (0 if (h, v) == (0, 0) else 1 if (h, v) == (0, 1) else 2 if (h, v) == (1, -1) else 3 if (h, v) == (1, 0) else 4)
        xr[i] = flip
    return bytes(ctx), bytes(xr)


SC_CTX, SC_XOR = _sc_tables()
_ORDERS: dict = {}


def _order(w: int, h: int) -> tuple:
    """走査の順（ストライプごと、列ごと、上から）：(全部の番号, 列の並び（4 行そろった列）, 端の行の列, 各ストライプの最後の行の印)。"""
    k = (w, h)
    o = _ORDERS.get(k)
    if o is None:
        st = w + 2
        flat = []
        cols = []
        lastrow = bytearray((w + 2) * (h + 2))
        for y0 in range(0, h, 4):
            ye = min(y0 + 4, h)
            for x in range(w):
                col = tuple((y + 1) * st + x + 1 for y in range(y0, ye))
                flat += col
                cols.append(col)
            for x in range(w):
                lastrow[(ye) * st + x + 1] = 1  # ストライプの最後の行（y = ye − 1）
        o = (flat, cols, lastrow)
        if len(_ORDERS) > 64:
            _ORDERS.clear()
        _ORDERS[k] = o
    return o


_T1_SRC = r'''
def decode_block(segs, w, h, numbps, orient, style):
    """segs：符号語の区切りの並び [(データ, 通り道の数)]。返す：(大きさ（小数 1 ビット）, 印)：(w+2)×(h+2) の平らな並び。"""
    st = w + 2
    n = st * (h + 2)
    flags = [0] * n
    mag = [0] * n
    zc = ZC[orient]
    sc_ctx, sc_xor = SC_CTX, SC_XOR
    qe_t, qs_t, nm_t, nl_t = QE, QS, NM, NL
    flat, cols, lastrow = _order(w, h)
    vsc = style & VCAUSAL
    cx = list(_INIT_CX)
    passtype = 2
    bp = numbps - 1
    done = 0
    for seg_data, seg_passes in segs:
        raw = bool(style & BYPASS) and passtype < 2 and bp <= numbps - 1 - 4
        data = seg_data + (b"\xff" * 1300 if raw else b"\xff\xff")  # データの終わりの先（生のビットは 1 ブロックで 8192 ビットまで）
        bp_i = 0
        if raw:
            rb = 0
            rc = 0
            rpos = 0
        else:
            c = data[0] << 16
            if data[0] == 0xFF:
                b = data[1]
                if b > 0x8F:
                    c += 0xFF00
                    ct = 8
                else:
                    bp_i = 1
                    c += b << 9
                    ct = 7
            else:
                bp_i = 1
                c += data[1] << 8
                ct = 8
            c <<= 7
            ct -= 7
            a = 0x8000
        for _p in range(seg_passes):
            if bp < 0:
                return mag, flags
            one = 1 << (bp + 1)
            half = one >> 1
            oph = one | half
            if passtype == 0:
                # 有意の伝播
                if raw:
                    for i in flat:
                        f = flags[i]
                        if vsc and lastrow[i]:
                            f &= VSC_MASK
                        if f & 0xFF and not f & SIG:
                            @RAW(d)
                            if d:
                                @RAW(sg)
                                _set_sig(flags, i, st, sg)
                                mag[i] = oph
                            flags[i] |= VISIT
                else:
                    for i in flat:
                        f = flags[i]
                        if vsc and lastrow[i]:
                            f &= VSC_MASK
                        if f & 0xFF and not f & SIG:
                            k = zc[f & 0xFF]
                            @MQ(k, d)
                            if d:
                                j = (f & 15) | ((f >> 4) & 0xF0)
                                k = sc_ctx[j]
                                @MQ(k, d)
                                sg = d ^ sc_xor[j]
                                @SET_SIG()
                                mag[i] = oph
                            flags[i] |= VISIT
            elif passtype == 1:
                # 細かくする
                todo = [i for i in flat if flags[i] & 0x5000 == 0x1000]
                if raw:
                    for i in todo:
                        @RAW(d)
                        mag[i] += half if d else -half
                        flags[i] |= REFINED
                else:
                    for i in todo:
                        f = flags[i]
                        if f & REFINED:
                            k = 16
                        else:
                            if vsc and lastrow[i]:
                                f &= VSC_MASK
                            k = 15 if f & 0xFF else 14
                            flags[i] |= REFINED
                        @MQ(k, d)
                        mag[i] += half if d else -half
            else:
                # 片付け（連続の符号を含む）
                for col in cols:
                    start = 0
                    if len(col) == 4:
                        i0, i1, i2, i3 = col
                        f3 = flags[i3]
                        if vsc:
                            f3 &= VSC_MASK
                        if not ((flags[i0] | flags[i1] | flags[i2] | f3) & 0x50FF):
                            @MQ(17, d)
                            if not d:
                                continue
                            @MQ(18, d)
                            r = d << 1
                            @MQ(18, d)
                            r |= d
                            i = col[r]
                            f = flags[i]
                            if vsc and lastrow[i]:
                                f &= VSC_MASK
                            j = (f & 15) | ((f >> 4) & 0xF0)
                            k = sc_ctx[j]
                            @MQ(k, d)
                            sg = d ^ sc_xor[j]
                            @SET_SIG()
                            mag[i] = oph
                            start = r + 1
                    for i in col[start:] if start else col:
                        f = flags[i]
                        if f & VISIT:
                            flags[i] = f & ~VISIT
                            continue
                        if f & SIG:
                            continue
                        if vsc and lastrow[i]:
                            f &= VSC_MASK
                        k = zc[f & 0xFF]
                        @MQ(k, d)
                        if d:
                            j = (f & 15) | ((f >> 4) & 0xF0)
                            k = sc_ctx[j]
                            @MQ(k, d)
                            sg = d ^ sc_xor[j]
                            @SET_SIG()
                            mag[i] = oph
                if style & SEGSYM:
                    for _s in range(4):
                        @MQ(18, d)
            if style & RESET and not raw:
                cx = list(_INIT_CX)
            passtype += 1
            if passtype == 3:
                passtype = 0
                bp -= 1
            done += 1
    return mag, flags
'''


def _set_sig(flags, i, st, sg) -> None:
    """生のビットの通り道で有意になった係数（速さの要らない所）。"""
    im = i - st
    ip = i + st
    if sg:
        flags[i] |= SIG | NEG
        flags[im] |= F_S | S_S
        flags[ip] |= F_N | S_N
        flags[i - 1] |= F_E | S_E
        flags[i + 1] |= F_W | S_W
    else:
        flags[i] |= SIG
        flags[im] |= F_S
        flags[ip] |= F_N
        flags[i - 1] |= F_E
        flags[i + 1] |= F_W
    flags[im - 1] |= F_SE
    flags[im + 1] |= F_SW
    flags[ip - 1] |= F_NE
    flags[ip + 1] |= F_NW


_MQ = """s = cx[{cx}]
q = qe_t[s]
qs = qs_t[s]
a -= q
if c < qs:
    if a < q:
        {d} = s & 1
        cx[{cx}] = nm_t[s]
    else:
        {d} = (s & 1) ^ 1
        cx[{cx}] = nl_t[s]
    a = q
    while True:
        if ct == 0:
            if data[bp_i] == 0xFF:
                b = data[bp_i + 1]
                if b > 0x8F:
                    c += 0xFF00
                    ct = 8
                else:
                    bp_i += 1
                    c += b << 9
                    ct = 7
            else:
                bp_i += 1
                c += data[bp_i] << 8
                ct = 8
        a <<= 1
        c <<= 1
        ct -= 1
        if a >= 0x8000:
            break
else:
    c -= qs
    if a >= 0x8000:
        {d} = s & 1
    else:
        if a < q:
            {d} = (s & 1) ^ 1
            cx[{cx}] = nl_t[s]
        else:
            {d} = s & 1
            cx[{cx}] = nm_t[s]
        while True:
            if ct == 0:
                if data[bp_i] == 0xFF:
                    b = data[bp_i + 1]
                    if b > 0x8F:
                        c += 0xFF00
                        ct = 8
                    else:
                        bp_i += 1
                        c += b << 9
                        ct = 7
                else:
                    bp_i += 1
                    c += data[bp_i] << 8
                    ct = 8
            a <<= 1
            c <<= 1
            ct -= 1
            if a >= 0x8000:
                break
"""

# 生のビット（D.6）：0xFF の次のバイトは 7 ビット
_RAW = """if rc == 0:
    if rb == 0xFF:
        rb = data[rpos]
        rc = 7
    else:
        rb = data[rpos]
        rc = 8
    rpos += 1
rc -= 1
{d} = (rb >> rc) & 1
"""

_SET_SIG = """im = i - st
ip = i + st
if sg:
    flags[i] |= 0x9000
    flags[im] |= 0x202
    flags[ip] |= 0x101
    flags[i - 1] |= 0x404
    flags[i + 1] |= 0x808
else:
    flags[i] |= 0x1000
    flags[im] |= 2
    flags[ip] |= 1
    flags[i - 1] |= 4
    flags[i + 1] |= 8
flags[im - 1] |= 0x40
flags[im + 1] |= 0x80
flags[ip - 1] |= 0x10
flags[ip + 1] |= 0x20
"""


def _expand(src: str) -> str:
    out = []
    for line in src.split("\n"):
        m = re.match(r"(\s*)@(MQ|RAW|SET_SIG)\((.*)\)\s*$", line)
        if not m:
            out.append(line)
            continue
        ind, what, args = m[1], m[2], [x.strip() for x in m[3].split(",")]
        body = _MQ.format(cx=args[0], d=args[1]) if what == "MQ" else _RAW.format(d=args[0]) if what == "RAW" else _SET_SIG
        out += [ind + ln for ln in body.rstrip("\n").split("\n")]
    return "\n".join(out)


# 文脈の初めの状態（表 D.7）：ZC の 0 は状態 4、RL は 3、UNI は 46、ほかは 0（状態の値は 2 × 番号 + MPS）
_INIT_CX = [2 * 4] + [0] * 16 + [2 * 3, 2 * 46]

_ns = {"ZC": ZC, "SC_CTX": SC_CTX, "SC_XOR": SC_XOR, "QE": QE, "QS": QS, "NM": NM, "NL": NL, "VSC_MASK": VSC_MASK, "SIG": SIG, "VISIT": VISIT,
       "REFINED": REFINED, "VCAUSAL": VCAUSAL, "BYPASS": BYPASS, "RESET": RESET, "SEGSYM": SEGSYM, "_order": _order, "_INIT_CX": _INIT_CX,
       "_set_sig": _set_sig}
exec(compile(_expand(_T1_SRC), "<jpx_t1>", "exec"), _ns)  # noqa: S102
decode_block = _ns["decode_block"]
