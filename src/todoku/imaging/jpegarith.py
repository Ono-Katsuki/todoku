"""JPEG（ITU-T T.81）の算術符号（附属書 D の QM 符号器・附属書 F.1.4・F.2.4 の順次・G.1.3・G.2.3 のプログレッシブ）と、可逆（附属書 H）のハフマン。

係数は jpegprog と同じ、成分ごとの平らな配列（ジグザグの順）。どちらも規格の流れ図のとおりで、速さより読みやすさを取った（算術符号の JPEG はめったにない）。
"""

from __future__ import annotations

# 表 D.3：(Qe, 次（LPS）, 次（MPS）, MPS を入れ替えるか)
_QE = (
    (0x5A1D, 1, 1, 1), (0x2586, 14, 2, 0), (0x1114, 16, 3, 0), (0x080B, 18, 4, 0), (0x03D8, 20, 5, 0), (0x01DA, 23, 6, 0),
    (0x00E5, 25, 7, 0), (0x006F, 28, 8, 0), (0x0036, 30, 9, 0), (0x001A, 33, 10, 0), (0x000D, 35, 11, 0), (0x0006, 9, 12, 0),
    (0x0003, 10, 13, 0), (0x0001, 12, 13, 0), (0x5A7F, 15, 15, 1), (0x3F25, 36, 16, 0), (0x2CF2, 38, 17, 0), (0x207C, 39, 18, 0),
    (0x17B9, 40, 19, 0), (0x1182, 42, 20, 0), (0x0CEF, 43, 21, 0), (0x09A1, 45, 22, 0), (0x072F, 46, 23, 0), (0x055C, 48, 24, 0),
    (0x0406, 49, 25, 0), (0x0303, 51, 26, 0), (0x0240, 52, 27, 0), (0x01B1, 54, 28, 0), (0x0144, 56, 29, 0), (0x00F5, 57, 30, 0),
    (0x00B7, 59, 31, 0), (0x008A, 60, 32, 0), (0x0068, 62, 33, 0), (0x004E, 63, 34, 0), (0x003B, 32, 35, 0), (0x002C, 33, 9, 0),
    (0x5AE1, 37, 37, 1), (0x484C, 64, 38, 0), (0x3A0D, 65, 39, 0), (0x2EF1, 67, 40, 0), (0x261F, 68, 41, 0), (0x1F33, 69, 42, 0),
    (0x19A8, 70, 43, 0), (0x1518, 72, 44, 0), (0x1177, 73, 45, 0), (0x0E74, 74, 46, 0), (0x0BFB, 75, 47, 0), (0x09F8, 77, 48, 0),
    (0x0861, 78, 49, 0), (0x0706, 79, 50, 0), (0x05CD, 48, 51, 0), (0x04DE, 50, 52, 0), (0x040F, 50, 53, 0), (0x0363, 51, 54, 0),
    (0x02D4, 52, 55, 0), (0x025C, 53, 56, 0), (0x01F8, 54, 57, 0), (0x01A4, 55, 58, 0), (0x0160, 56, 59, 0), (0x0125, 57, 60, 0),
    (0x00F6, 58, 61, 0), (0x00CB, 59, 62, 0), (0x00AB, 61, 63, 0), (0x008F, 61, 32, 0), (0x5B12, 65, 65, 1), (0x4D04, 80, 66, 0),
    (0x412C, 81, 67, 0), (0x37D8, 82, 68, 0), (0x2FE8, 83, 69, 0), (0x293C, 84, 70, 0), (0x2379, 86, 71, 0), (0x1EDF, 87, 72, 0),
    (0x1AA9, 87, 73, 0), (0x174E, 72, 74, 0), (0x1424, 72, 75, 0), (0x119C, 74, 76, 0), (0x0F6B, 74, 77, 0), (0x0D51, 75, 78, 0),
    (0x0BB6, 77, 79, 0), (0x0A40, 77, 48, 0), (0x5832, 80, 81, 1), (0x4D1C, 88, 82, 0), (0x438E, 89, 83, 0), (0x3BDD, 90, 84, 0),
    (0x34EE, 91, 85, 0), (0x2EAE, 92, 86, 0), (0x299A, 93, 87, 0), (0x2516, 86, 71, 0), (0x5570, 88, 89, 1), (0x4CA9, 95, 90, 0),
    (0x44D9, 96, 91, 0), (0x3E22, 97, 92, 0), (0x3824, 99, 93, 0), (0x32B4, 99, 94, 0), (0x2E17, 93, 86, 0), (0x56A8, 95, 96, 1),
    (0x4F46, 101, 97, 0), (0x47E5, 102, 98, 0), (0x41CF, 103, 99, 0), (0x3C3D, 104, 100, 0), (0x375E, 99, 93, 0), (0x5231, 105, 102, 0),
    (0x4C0F, 106, 103, 0), (0x4639, 107, 104, 0), (0x415E, 103, 99, 0), (0x5627, 105, 106, 1), (0x50E7, 108, 107, 0), (0x4B85, 109, 103, 0),
    (0x5597, 110, 109, 0), (0x504F, 111, 107, 0), (0x5A10, 110, 111, 1), (0x5522, 112, 109, 0), (0x59EB, 112, 111, 1),
    (0x5A1D, 113, 113, 0),  # 固定の確率（符号の字：F.1.4.3.1 の「Qe＝0x5A1D で変わらない」状態）
)
_FIXED = 113


class _QM:
    """算術符号の復号（D.2）。状態のバイト st[i]：下位 7 ビットが表の番号、最上位が MPS。詰め物を除いたバイトの並びを読み、尽きたら 0 を補う。"""

    __slots__ = ("buf", "pos", "a", "c", "ct")

    def __init__(self, buf: bytes):
        self.buf, self.pos = buf, 0
        self.a, self.c, self.ct = 0, 0, -16

    def bit(self, st: bytearray, i: int) -> int:
        a, c, ct = self.a, self.c, self.ct
        while a < 0x8000:
            ct -= 1
            if ct < 0:
                p = self.pos
                data = self.buf[p] if p < len(self.buf) else 0
                self.pos = p + 1
                c = (c << 8) | data
                ct += 8
                if ct < 0:
                    ct += 1
                    if ct == 0:
                        a = 0x8000  # 初めの 2 バイトを読んだ
            a <<= 1
        sv = st[i]
        qe, nl, nm, sw = _QE[sv & 0x7F]
        temp = a - qe
        a = temp
        temp <<= ct
        if c >= temp:
            c -= temp
            if a < qe:
                a = qe
                st[i] = (sv & 0x80) | nm
            else:
                a = qe
                st[i] = (sv & 0x80 ^ (0x80 if sw else 0)) | nl
                sv ^= 0x80
        elif a < 0x8000:
            if a < qe:
                st[i] = (sv & 0x80 ^ (0x80 if sw else 0)) | nl
                sv ^= 0x80
            else:
                st[i] = (sv & 0x80) | nm
        self.a, self.c, self.ct = a, c, ct
        return sv >> 7


def scan(segs, frame, coefs, sel, restart, progressive, ss, se, ah, al, hmax, vmax, dac) -> None:
    """1 つの走査。sel：(成分の番号, DC の表, AC の表)。dac：{(0, 表): (L, U)}・{(1, 表): Kx}（なければ既定 L=0・U=1・Kx=5）。"""
    ww, hh, comps = frame
    if len(sel) == 1:
        ci0 = sel[0][0]
        total = comps[ci0]["nx"] * comps[ci0]["ny"]
        order = None
    else:
        mcux = (ww + 8 * hmax - 1) // (8 * hmax)
        mcuy = (hh + 8 * vmax - 1) // (8 * vmax)
        total = mcux * mcuy
        order = mcux
    per = restart or total
    mi = 0
    p1, m1 = 1 << al, -(1 << al)
    fixed = bytearray([_FIXED])
    for buf in segs:
        if mi >= total:
            break
        qm = _QM(buf)
        bit = qm.bit
        dc_st: dict = {}
        ac_st: dict = {}
        for _ci, td, ta in sel:
            dc_st.setdefault(td, bytearray(64))
            ac_st.setdefault(ta, bytearray(256))
        pred = [0] * len(comps)
        ctx = [0] * len(comps)
        for _ in range(min(per, total - mi)):
            if order is None:
                ci, td, ta = sel[0]
                by, bx = divmod(mi, comps[ci]["nx"])
                blocks = ((ci, td, ta, by * comps[ci]["bw"] * 64 + bx * 64),)
            else:
                my, mx = divmod(mi, order)
                blocks = []
                for ci, td, ta in sel:
                    c = comps[ci]
                    for v in range(c["v"]):
                        for h in range(c["h"]):
                            blocks.append((ci, td, ta, ((my * c["v"] + v) * c["bw"] + mx * c["h"] + h) * 64))
            mi += 1
            for ci, td, ta, base in blocks:
                cf = coefs[ci]
                if ss == 0:
                    if ah == 0:
                        st = dc_st[td]
                        _dc(bit, st, ci, pred, ctx, dac.get((0, td), (0, 1)))
                        cf[base] = pred[ci] << al if progressive else pred[ci]
                    elif bit(fixed, 0):
                        cf[base] |= p1
                    if progressive:
                        continue
                st = ac_st[ta]
                kx = dac.get((1, ta), 5)
                lo, hi = (1, 63) if not progressive else (ss, se)
                if ah == 0:
                    _ac_first(bit, st, fixed, cf, base, lo, hi, kx, al)
                else:
                    _ac_refine(bit, st, fixed, cf, base, ss, se, p1, m1)


def _dc(bit, st: bytearray, ci: int, pred: list, ctx: list, lu: tuple) -> None:
    """図 F.19〜F.24：DC の差。"""
    s = ctx[ci]
    if not bit(st, s):
        ctx[ci] = 0
        return
    sign = bit(st, s + 1)
    s = s + 2 + sign
    m = bit(st, s)
    if m:
        s = 20
        while bit(st, s):
            m <<= 1
            if m >= 0x10000:
                return
            s += 1
    L, U = lu
    if m < ((1 << L) >> 1):
        ctx[ci] = 0
    elif m > ((1 << U) >> 1):
        ctx[ci] = 12 + sign * 4
    else:
        ctx[ci] = 4 + sign * 4
    v = m
    s += 14
    m >>= 1
    while m:
        if bit(st, s):
            v |= m
        m >>= 1
    v += 1
    pred[ci] += -v if sign else v


def _ac_first(bit, st: bytearray, fixed: bytearray, cf, base: int, lo: int, hi: int, kx: int, al: int) -> None:
    """図 F.20〜F.24（順次）と G.1.3.3（プログレッシブの初めの走査）。"""
    k = lo
    while k <= hi:
        s = 3 * (k - 1)
        if bit(st, s):
            return  # 終わり（EOB）
        while not bit(st, s + 1):
            s += 3
            k += 1
            if k > hi:
                return
        sign = bit(fixed, 0)
        s += 2
        m = bit(st, s)
        if m and bit(st, s):
            m = 2
            s = 189 if k <= kx else 217
            while bit(st, s):
                m <<= 1
                if m >= 0x10000:
                    return
                s += 1
        v = m
        s += 14
        m >>= 1
        while m:
            if bit(st, s):
                v |= m
            m >>= 1
        v += 1
        cf[base + k] = (-v if sign else v) * (1 << al)
        k += 1


def _ac_refine(bit, st: bytearray, fixed: bytearray, cf, base: int, ss: int, se: int, p1: int, m1: int) -> None:
    """図 G.7・G.8：AC を細かくする走査。"""
    kex = se
    while kex > 0 and not cf[base + kex]:
        kex -= 1
    k = ss
    while k <= se:
        s = 3 * (k - 1)
        if k > kex and bit(st, s):
            return
        while True:
            q = base + k
            cv = cf[q]
            if cv:
                if bit(st, s + 2):
                    cf[q] = cv + (m1 if cv < 0 else p1)
                break
            if bit(st, s + 1):
                cf[q] = m1 if bit(fixed, 0) else p1
                break
            s += 3
            k += 1
            if k > se:
                return
        k += 1


# ====================================================================== 可逆（附属書 H：ハフマン）
def lossless_scan(segs, frame, planes, sel, hts, restart, psv: int, pt: int, precision: int, hmax: int, vmax: int, err) -> None:
    """1 つの走査を、成分ごとの面（array）に開く。planes[ci]：(面, 幅, 高さ)。H.1.2：予測の 7 通り。区切りのすぐ後は予測を初めからやり直す。"""
    ww, hh, comps = frame
    if len(sel) == 1:
        ci0 = sel[0][0]
        cw, chh = planes[ci0][1], planes[ci0][2]
        total = cw * chh
        order = None
    else:
        mcux = (ww + hmax - 1) // hmax
        mcuy = (hh + vmax - 1) // vmax
        total = mcux * mcuy
        order = mcux
    per = restart or total
    mask = (1 << precision) - 1
    first = 1 << (precision - pt - 1)
    mi = 0
    for buf in segs:
        if mi >= total:
            break
        buf = buf + b"\0" * 8
        pos = 0
        fresh = {}  # 区切りのあと初めの行（成分ごと）：左の値か、行の頭は既定の値
        for _ in range(min(per, total - mi)):
            if order is None:
                ci, td = sel[0]
                y, x = divmod(mi, planes[ci][1])
                units = ((ci, td, x, y),)
            else:
                my, mx = divmod(mi, order)
                units = []
                for ci, td in sel:
                    c = comps[ci]
                    for v in range(c["v"]):
                        for h in range(c["h"]):
                            units.append((ci, td, mx * c["h"] + h, my * c["v"] + v))
            mi += 1
            for ci, td, x, y in units:
                tb = hts.get((0, td))
                if tb is None:
                    raise err("JPEG：可逆の走査のハフマン表がありません")
                plane, w, h = planes[ci]
                i = pos >> 3
                e = tb.lut[(int.from_bytes(buf[i:i + 4], "big") << (pos & 7) & 0xFFFFFFFF) >> 16]
                if not e:
                    raise err("JPEG：ハフマン符号が読めません")
                t = e & 255
                pos += e >> 8
                if t == 16:
                    d = 32768
                elif t:
                    i = pos >> 3
                    v = (int.from_bytes(buf[i:i + 4], "big") >> (32 - (pos & 7) - t)) & ((1 << t) - 1)
                    pos += t
                    d = v - (1 << t) + 1 if v < (1 << (t - 1)) else v
                else:
                    d = 0
                if x >= w or y >= h:
                    continue  # MCU の詰め物
                o = y * w + x
                start_row = fresh.setdefault(ci, y)
                if y == start_row:
                    p = first if x == 0 else plane[o - 1]
                elif x == 0:
                    p = plane[o - w]
                else:
                    a, b, c_ = plane[o - 1], plane[o - w], plane[o - w - 1]
                    if psv == 1:
                        p = a
                    elif psv == 2:
                        p = b
                    elif psv == 3:
                        p = c_
                    elif psv == 4:
                        p = a + b - c_
                    elif psv == 5:
                        p = a + ((b - c_) >> 1)
                    elif psv == 6:
                        p = b + ((a - c_) >> 1)
                    elif psv == 7:
                        p = (a + b) >> 1
                    else:
                        raise err(f"JPEG：可逆の予測の番号 {psv} は規格にありません")
                plane[o] = (p + d) & mask
