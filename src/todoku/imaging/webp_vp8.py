"""WebP の非可逆の画像（VP8 のキーフレーム、RFC 6386）を開く：ブール符号・係数のトークン・予測・逆 DCT・逆 WHT・ループフィルター。

純粋な Python で、1 メガ画素あたり数秒。色は YUV 4:2:0（色差は近い画素で広げる）。表は RFC 6386 の 13・11 章の値を 16 進で持つ（LICENSES/LICENSE-webm.txt）。
係数の読み方の一部（_large・_coeffs・_BANDS）と定数 101581 は libwebp（BSD 3-Clause。LICENSES/LICENSE-libwebp.txt）に基づく。
"""

from __future__ import annotations

from .formats import ImageError

_DEFAULT_COEF = bytes.fromhex(
    "808080808080808080808080808080808080808080808080808080808080808080fd88feffe4db8080808080bd81f2ffe3d5ffdb8080806a7ee3fcd6"
    "d1ffff8080800162f8ffece2ffff808080b585eefeddeaff9a8080804e86caf7c6b4ffdb80808001b9f9fff3ff8080808080b896f7ffece080808080"
    "804d6ed8ffece680808080800165fbfff1ff8080808080aa8bf1fcecd1ffff8080802574c4f3e4ffffff80808001ccfefff5ff8080808080cfa0faff"
    "ee8080808080806667e7ffd3ab80808080800198fcfff0ff8080808080b187f3ffeae180808080805081d3ffc2e080808080800101ff808080808080"
    "8080f601ff8080808080808080ff80808080808080808080c623eddfc1bba2a0919b3e832dc6ddacb0dc9dfcdd01442f92d095a7dda2ffdf800195f1"
    "ffdde0ffff808080b88deafddedcffc78080805163b5f2b0bef9caffff800181e8fdd6c5f2c4ffff806379d2fac9c6ffca808080175ba3f2aabbf7d2"
    "ffff8001c8f6ffeaff80808080806db2f1ffe7f5ffff8080802c82c9fdcdc0ffff8080800184effbdbd1ffa58080805e88e1fbdabeffff8080801664"
    "aef5baa1ffc780808001b6f9ffe8eb80808080807c8ff1ffe3ea8080808080234db5fbc1d3ffcd808080019df7ffece7ffff808080798debffe1e3ff"
    "ff8080802d63bcfbc3d9ffe08080800101fbffd5ff8080808080cb01f8ffff8080808080808901b1ffe0ff8080808080fd09f8fbcfd0ffc0808080af"
    "0de0f3c1b9f9c6ffff804911abdda1b3eca7ffea80015ff7fdd4b7ffff808080ef5af4fad3d1ffff8080809b4dc3f8bcc3ffff8080800118effbdadb"
    "ffcd808080c933dbffc4ba8080808080452ebeefc9daffe480808001bffbffff808080808080dfa5f9ffd5ff80808080808d7cf8ffff808080808080"
    "0110f8ffff808080808080be24e6ffecff80808080809501ff808080808080808001e2ff8080808080808080f7c0ff8080808080808080f080ff8080"
    "8080808080800186fcffff808080808080d53efaffff808080808080375dff8080808080808080808080808080808080808080808080808080808080"
    "808080808080808080808080ca18d5ebbabfdca0f0afff7e26b6e8a9b8e4aeffbb803d2e8adb97b2f0aaffd8800170e6fac7bff79fffff80a66de4fc"
    "d3d7ffae808080274da2e8acb4f5b2ffff800134dcf6c6c7f9dcffff807c4abff3b7c1faddffff80184782db9aaaf3b6ffff8001b6e1f9dbf0ffe080"
    "80809596e2fcd8cdffab8080801c6caaf2b7c2fedfffff800151e6fccccbffc08080807b66d1f7bcc4ffe9808080145f99f3a4adffcb80808001def8"
    "ffd8d58080808080a8aff6fcebcdffff8080802f74d7ffd3d4ffff8080800179ecfdd4d6ffff8080808d54d5fcc9caffdb8080802a50a0f0a2b9ffcd"
    "8080800101ff8080808080808080f401ff8080808080808080ee01ff8080808080808080"
)
_COEF_UPDATE = bytes.fromhex(
    "ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffb0f6ffffffffffffffffffdff1fcfffffffffffffffff9fdfdffff"
    "fffffffffffffff4fcffffffffffffffffeafefefffffffffffffffffdfffffffffffffffffffffff6feffffffffffffffffeffdfeffffffffffffff"
    "fffefffefffffffffffffffffff8fefffffffffffffffffbfffefffffffffffffffffffffffffffffffffffffffffdfefffffffffffffffffbfefeff"
    "fffffffffffffffefffefffffffffffffffffffefdfffefffffffffffffafffefffefffffffffffffeffffffffffffffffffffffffffffffffffffff"
    "ffffffffffffffffffffffffffffffffffffffffffffffffd9ffffffffffffffffffffe1fcf1fdfffffeffffffffeafaf1fafdfffdfefffffffffeff"
    "ffffffffffffffffdffefeffffffffffffffffeefdfefefffffffffffffffff8fefffffffffffffffff9feffffffffffffffffffffffffffffffffff"
    "fffffffffdfffffffffffffffffff7fefffffffffffffffffffffffffffffffffffffffffffdfefffffffffffffffffcffffffffffffffffffffffff"
    "fffffffffffffffffffffefefffffffffffffffffdfffffffffffffffffffffffffffffffffffffffffffffefdfffffffffffffffffaffffffffffff"
    "fffffffffeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffbafbfaffffffffffffffffea"
    "fbf4fefffffffffffffffbfbf3fdfefffefffffffffffdfeffffffffffffffffecfdfefffffffffffffffffbfdfdfefefffffffffffffffefeffffff"
    "fffffffffffefefefffffffffffffffffffffffffffffffffffffffffefffffffffffffffffffefefffffffffffffffffffeffffffffffffffffffff"
    "fffffffffffffffffffffffeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff"
    "ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff"
    "fffffffffffffffffffffffff8fffffffffffffffffffffafefcfefffffffffffffff8fef9fdfffffffffffffffffdfdfffffffffffffffff6fdfdff"
    "fffffffffffffffcfefbfefefffffffffffffffefcfffffffffffffffff8fefdfffffffffffffffffdfffefefffffffffffffffffbfeffffffffffff"
    "fffff5fbfefffffffffffffffffdfdfefffffffffffffffffffbfdfffffffffffffffffcfdfefffffffffffffffffffefffffffffffffffffffffcff"
    "fffffffffffffffff9fffefffffffffffffffffffffefffffffffffffffffffffdfffffffffffffffffaffffffffffffffffffffffffffffffffffff"
    "fffffffffffffffffffffffffffffeffffffffffffffffffffffffffffffffffffffffff"
)
_BMODES = bytes.fromhex(
    "e7783059737178987098b3407eaa762e465faf458f505552489b67383a0aabdabd110d98721a11a32cc3150aad791850c31a3e2c405590470a26abd5"
    "90221aaa2e371388a021ce473f14087272d00c09e251280b60b6541d102486b7598962656aa59448bb64829d6f204b504266a7634a3e28ea80293509"
    "b2f18d1a086b4a2b1a9249a631179d412669a033341f7380684f0c1bd9ff5711075744472c72330fba172f290e6eb6b71511c2422d1966c5bd171216"
    "585893962a2e2dc4cd2b61b775552623b33d2735c8571a152be8ab3822336872661d5d4d271c55ab3aa55a6240221674ce17222ba6496b36201a3301"
    "512b1f44196a1640ab24e1722213156684bc104c7c3e124e5f5539323033c165239fd76f592e6f3c941facdbe415126f70714d55b3ff267872282a01"
    "c4f5d10a196d582b1d8ca6d5252b9a3d3f1e9b432d4401d16450082b9a01331a478e4e4e10ff8022c5ab29280566d3b70401dd333211a8d1c0171952"
    "8a1f24ab1ba6262ce543573aa952731a3bb33f3b5ab43ba65d499a282815748fd12227af2f0f10b722df312db72e1121b706620f20b7392e16188001"
    "361125412049731c801780cd2803097333c01206df572509733b4d40152f68372cda09363582e2405a46cd2829171a39363970b8052926a6d51e221a"
    "8598740a2086271335dd1a722049ff1f0941ea020f0176494b200c33c0ffa02b33581f2343665537ba553815176f3bcd2d25c03726467c4966012262"
    "7d622a58685575af525f543559806471652d4b4f7b2f338051ab0139110547663935293126210d7939491a0155290a438a4d6e5a2f727315020a66ff"
    "a61706651d100a558065c41a39120a6666d522142b75140f24a38044011a663d472522351ff3c0453c472649771cde25442d8022012f0bf5ab3e1113"
    "469255373e46252b259a64a355a0013f095c881c4020c9554b0f090940ffb8771056061c0540ff19f8013808118489ff3774803a0f145287391a7928"
    "a4321f899a851923da33672c83837b1f069e5628408794e02db780161a1183f09a0e01d12d10155b40de0701c53815279b3c8a1766d5530c0d36c0ff"
    "442f1c551a555580802092ab120b073f90ab0404f6231b0a92aeab0c1a80be502363b4507e362d557e2f57b033291420654b808b769274805538290f"
    "b0ec5525093e471e117776ff11128a65263c8a37462b1a8e9224131eabff611b148a2d3d3edb0151bc4020291475978e1415a370130c3dc380300418"
)

_DC = (4, 5, 6, 7, 8, 9, 10, 10, 11, 12, 13, 14, 15, 16, 17, 17, 18, 19, 20, 20, 21, 21, 22, 22, 23, 23, 24, 25, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 37, 38, 39, 40, 41,
       42, 43, 44, 45, 46, 46, 47, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64, 65, 66, 67, 68, 69, 70, 71, 72, 73, 74, 75, 76, 76, 77, 78, 79, 80, 81, 82, 83, 84,
       85, 86, 87, 88, 89, 91, 93, 95, 96, 98, 100, 101, 102, 104, 106, 108, 110, 112, 114, 116, 118, 122, 124, 126, 128, 130, 132, 134, 136, 138, 140, 143, 145, 148, 151, 154, 157)
_AC = (4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50,
       51, 52, 53, 54, 55, 56, 57, 58, 60, 62, 64, 66, 68, 70, 72, 74, 76, 78, 80, 82, 84, 86, 88, 90, 92, 94, 96, 98, 100, 102, 104, 106, 108, 110, 112, 114, 116, 119, 122, 125, 128, 131,
       134, 137, 140, 143, 146, 149, 152, 155, 158, 161, 164, 167, 170, 173, 177, 181, 185, 189, 193, 197, 201, 205, 209, 213, 217, 221, 225, 229, 234, 239, 245, 249, 254, 259, 264, 269,
       274, 279, 284)
_ZIGZAG = (0, 1, 4, 8, 5, 2, 3, 6, 9, 12, 13, 10, 7, 11, 14, 15)
_BANDS = (0, 1, 2, 3, 6, 4, 5, 6, 6, 6, 6, 6, 6, 6, 6, 7, 0)
_CAT = ((173, 148, 140), (176, 155, 140, 135), (180, 157, 141, 134, 130), (254, 254, 243, 230, 196, 177, 153, 140, 133, 130, 129))
# 字の順（RFC 6386 11.2 の木を、番号の表にしたもの）：0 DC・1 TM・2 VE・3 HE・4 RD・5 VR・6 LD・7 VL・8 HD・9 HU（_BMODES の添字の順）
_B_DC, _B_TM, _B_VE, _B_HE, _B_RD, _B_VR, _B_LD, _B_VL, _B_HD, _B_HU = range(10)
_B_TREE = (-_B_DC, 1, -_B_TM, 2, -_B_VE, 3, 4, 6, -_B_HE, 5, -_B_RD, -_B_VR, -_B_LD, 7, -_B_VL, 8, -_B_HD, -_B_HU)
_DC_PRED, _TM_PRED, _V_PRED, _H_PRED = _B_DC, _B_TM, _B_VE, _B_HE
_B_PRED = 10
_NORM = [0] * 256
for _i in range(1, 256):
    _n = 0
    while (_i << _n) < 128:
        _n += 1
    _NORM[_i] = _n


class _Bool:
    """ブール符号の復号（RFC 6386 7.3）。"""

    __slots__ = ("d", "pos", "end", "value", "range", "bits")

    def __init__(self, d: bytes, start: int, end: int):
        self.d = d
        self.pos = start + 2
        self.end = end
        self.value = (d[start] << 8 | d[start + 1]) if start + 1 < end else (d[start] << 8 if start < end else 0)
        self.range = 255
        self.bits = 0

    def read(self, p: int) -> int:
        split = 1 + (((self.range - 1) * p) >> 8)
        big = split << 8
        if self.value >= big:
            self.value -= big
            rng = self.range - split
            r = 1
        else:
            rng = split
            r = 0
        if rng < 128:
            s = _NORM[rng]
            rng <<= s
            bits = self.bits + s
            if bits >= 8:
                bits -= 8
                pos = self.pos
                self.value = (self.value << (s - bits)) | (self.d[pos] if pos < self.end else 0)
                self.value <<= bits
                self.pos = pos + 1
            else:
                self.value <<= s
            self.bits = bits
        self.range = rng
        return r

    def lit(self, n: int) -> int:
        v = 0
        for _ in range(n):
            v = v << 1 | self.read(128)
        return v

    def flag_signed(self, n: int) -> int:
        """旗＋n ビット＋符号。旗が 0 なら 0。"""
        if not self.read(128):
            return 0
        v = self.lit(n)
        return -v if self.read(128) else v


def _clamp_q(q: int) -> int:
    return 0 if q < 0 else 127 if q > 127 else q


def _large(br: _Bool, p) -> int:
    rd = br.read
    if not rd(p[3]):
        if not rd(p[4]):
            return 2
        return 3 + rd(p[5])
    if not rd(p[6]):
        if not rd(p[7]):
            return 5 + rd(159)
        v = 7 + 2 * rd(165)
        return v + rd(145)
    b1 = rd(p[8])
    b0 = rd(p[9 + b1])
    cat = 2 * b1 + b0
    v = 0
    for t in _CAT[cat]:
        v += v + rd(t)
    return v + 3 + (8 << cat)


def _coeffs(br: _Bool, prob, ctx: int, dq: tuple, n: int, out: list) -> int:
    """1 つのブロックの係数（13 章）。返す：最後の 0 でない係数の次の番号（なければ n の初めの値以下）。"""
    rd = br.read
    p = prob[_BANDS[n]][ctx]
    while n < 16:
        if not rd(p[0]):
            return n
        while not rd(p[1]):
            n += 1
            if n == 16:
                return 16
            p = prob[_BANDS[n]][0]
        nxt = prob[_BANDS[n + 1]]
        if not rd(p[2]):
            v = 1
            p = nxt[1]
        else:
            v = _large(br, p)
            p = nxt[2]
        if rd(128):
            v = -v
        out[_ZIGZAG[n]] = v * dq[1 if n > 0 else 0]
        n += 1
    return 16


def _iwht(c: list) -> list:
    """逆 WHT（14.3）：16 個の DC。"""
    t = [0] * 16
    for i in range(4):
        a1, b1 = c[i] + c[12 + i], c[4 + i] + c[8 + i]
        c1, d1 = c[4 + i] - c[8 + i], c[i] - c[12 + i]
        t[i], t[4 + i], t[8 + i], t[12 + i] = a1 + b1, c1 + d1, a1 - b1, d1 - c1
    out = [0] * 16
    for i in range(4):
        q = 4 * i
        a1, b1 = t[q] + t[q + 3], t[q + 1] + t[q + 2]
        c1, d1 = t[q + 1] - t[q + 2], t[q] - t[q + 3]
        out[q], out[q + 1], out[q + 2], out[q + 3] = (a1 + b1 + 3) >> 3, (c1 + d1 + 3) >> 3, (a1 - b1 + 3) >> 3, (d1 - c1 + 3) >> 3
    return out


def _idct_add(c: list, plane: bytearray, pos: int, stride: int) -> None:
    """逆 DCT（14.4）と足し算。"""
    if not any(c[1:]):
        d = (c[0] + 4) >> 3
        if d:
            for y in range(4):
                q = pos + y * stride
                for x in range(4):
                    v = plane[q + x] + d
                    plane[q + x] = 0 if v < 0 else 255 if v > 255 else v
        return
    t = [0] * 16
    for i in range(4):
        i0, i4, i8, i12 = c[i], c[4 + i], c[8 + i], c[12 + i]
        a1, b1 = i0 + i8, i0 - i8
        t1 = (i4 * 35468) >> 16
        t2 = i12 + ((i12 * 20091) >> 16)
        c1 = t1 - t2
        t1 = i4 + ((i4 * 20091) >> 16)
        t2 = (i12 * 35468) >> 16
        d1 = t1 + t2
        t[i], t[12 + i], t[4 + i], t[8 + i] = a1 + d1, a1 - d1, b1 + c1, b1 - c1
    for r in range(4):
        q0 = 4 * r
        i0, i1, i2, i3 = t[q0], t[q0 + 1], t[q0 + 2], t[q0 + 3]
        a1, b1 = i0 + i2, i0 - i2
        t1 = (i1 * 35468) >> 16
        t2 = i3 + ((i3 * 20091) >> 16)
        c1 = t1 - t2
        t1 = i1 + ((i1 * 20091) >> 16)
        t2 = (i3 * 35468) >> 16
        d1 = t1 + t2
        q = pos + r * stride
        for x, e in enumerate(((a1 + d1 + 4) >> 3, (b1 + c1 + 4) >> 3, (b1 - c1 + 4) >> 3, (a1 - d1 + 4) >> 3)):
            v = plane[q + x] + e
            plane[q + x] = 0 if v < 0 else 255 if v > 255 else v


def _pred_block(p: bytearray, pos: int, st: int, n: int, mode: int, mb_x: int, mb_y: int) -> None:
    """16×16・8×8 の予測（12.2）。上の行・左の列は、端なら 127・129。"""
    above = bytes(p[pos - st:pos - st + n]) if mb_y else b"\x7f" * n
    left = [p[pos + y * st - 1] for y in range(n)] if mb_x else [129] * n
    if mode == _DC_PRED:
        shift = 4 if n == 16 else 3
        if mb_y and mb_x:
            v = (sum(above) + sum(left) + (1 << shift)) >> (shift + 1)
        elif mb_y:
            v = (sum(above) + (1 << (shift - 1))) >> shift
        elif mb_x:
            v = (sum(left) + (1 << (shift - 1))) >> shift
        else:
            v = 128
        row = bytes([v]) * n
        for y in range(n):
            p[pos + y * st:pos + y * st + n] = row
    elif mode == _V_PRED:
        for y in range(n):
            p[pos + y * st:pos + y * st + n] = above
    elif mode == _H_PRED:
        for y in range(n):
            p[pos + y * st:pos + y * st + n] = bytes([left[y]]) * n
    else:  # TM
        tl = p[pos - st - 1] if mb_y and mb_x else (127 if not mb_y else 129)
        for y in range(n):
            d = left[y] - tl
            p[pos + y * st:pos + y * st + n] = bytes([0 if a + d < 0 else 255 if a + d > 255 else a + d for a in above])


def _pred4(p: bytearray, pos: int, st: int, mode: int, A: list, L: list, P: int) -> None:
    """4×4 の予測（12.3）。A：上の 8 個、L：左の 4 個、P：左上。"""
    if mode == _B_DC:
        v = (sum(A[:4]) + sum(L) + 4) >> 3
        rows = [[v] * 4] * 4
    elif mode == _B_TM:
        rows = [[0 if L[r] + a - P < 0 else 255 if L[r] + a - P > 255 else L[r] + a - P for a in A[:4]] for r in range(4)]
    elif mode == _B_VE:
        rows = [[(P + 2 * A[0] + A[1] + 2) >> 2, (A[0] + 2 * A[1] + A[2] + 2) >> 2, (A[1] + 2 * A[2] + A[3] + 2) >> 2, (A[2] + 2 * A[3] + A[4] + 2) >> 2]] * 4
    elif mode == _B_HE:
        rows = [[(P + 2 * L[0] + L[1] + 2) >> 2] * 4, [(L[0] + 2 * L[1] + L[2] + 2) >> 2] * 4, [(L[1] + 2 * L[2] + L[3] + 2) >> 2] * 4, [(L[2] + 3 * L[3] + 2) >> 2] * 4]
    elif mode == _B_LD:
        d = [(A[i] + 2 * A[i + 1] + A[i + 2] + 2) >> 2 for i in range(6)] + [(A[6] + 3 * A[7] + 2) >> 2]
        rows = [d[r:r + 4] for r in range(4)]
    else:
        E = [L[3], L[2], L[1], L[0], P, A[0], A[1], A[2], A[3]]
        t3 = [(E[i] + 2 * E[i + 1] + E[i + 2] + 2) >> 2 for i in range(7)]  # t3[i]：E[i], E[i+1], E[i+2]
        t2 = [(E[i] + E[i + 1] + 1) >> 1 for i in range(8)]
        if mode == _B_RD:
            rows = [[t3[3], t3[4], t3[5], t3[6]], [t3[2], t3[3], t3[4], t3[5]], [t3[1], t3[2], t3[3], t3[4]], [t3[0], t3[1], t3[2], t3[3]]]
        elif mode == _B_VR:
            rows = [[t2[4], t2[5], t2[6], t2[7]], [t3[3], t3[4], t3[5], t3[6]], [t3[2], t2[4], t2[5], t2[6]], [t3[1], t3[3], t3[4], t3[5]]]
        elif mode == _B_HD:
            rows = [[t2[3], t3[3], t3[4], t3[5]], [t2[2], t3[2], t2[3], t3[3]], [t2[1], t3[1], t2[2], t3[2]], [t2[0], t3[0], t2[1], t3[1]]]
        elif mode == _B_VL:
            a2 = [(A[i] + A[i + 1] + 1) >> 1 for i in range(4)]
            a3 = [(A[i] + 2 * A[i + 1] + A[i + 2] + 2) >> 2 for i in range(6)]
            rows = [[a2[0], a2[1], a2[2], a2[3]], [a3[0], a3[1], a3[2], a3[3]], [a2[1], a2[2], a2[3], a3[4]], [a3[1], a3[2], a3[3], a3[5]]]
        else:  # HU
            a = (L[0] + L[1] + 1) >> 1
            b = (L[0] + 2 * L[1] + L[2] + 2) >> 2
            c = (L[1] + L[2] + 1) >> 1
            d = (L[1] + 2 * L[2] + L[3] + 2) >> 2
            e = (L[2] + L[3] + 1) >> 1
            f = (L[2] + 3 * L[3] + 2) >> 2
            g = L[3]
            rows = [[a, b, c, d], [c, d, e, f], [e, f, g, g], [g, g, g, g]]
    for r, row in enumerate(rows):
        p[pos + r * st:pos + r * st + 4] = bytes(row)


def _sc(v: int) -> int:
    return -128 if v < -128 else 127 if v > 127 else v


def _edge(p: bytearray, pos: int, across: int, along: int, count: int, limit: int, il: int, hev_t: int, mb_edge: bool) -> None:
    """ふつうのループフィルター（15.3）の 1 つの辺。"""
    for i in range(count):
        q = pos + i * along
        p3, p2, p1, p0 = p[q - 4 * across], p[q - 3 * across], p[q - 2 * across], p[q - across]
        q0, q1, q2, q3 = p[q], p[q + across], p[q + 2 * across], p[q + 3 * across]
        if abs(p0 - q0) * 2 + (abs(p1 - q1) >> 1) > limit:
            continue
        if abs(p3 - p2) > il or abs(p2 - p1) > il or abs(p1 - p0) > il or abs(q3 - q2) > il or abs(q2 - q1) > il or abs(q1 - q0) > il:
            continue
        hev = abs(p1 - p0) > hev_t or abs(q1 - q0) > hev_t
        if hev:
            a = _sc(_sc(p1 - q1) + 3 * (q0 - p0))
            f1 = _sc(a + 4) >> 3
            f2 = _sc(a + 3) >> 3
            v = q0 - f1
            p[q] = 0 if v < 0 else 255 if v > 255 else v
            v = p0 + f2
            p[q - across] = 0 if v < 0 else 255 if v > 255 else v
        elif mb_edge:
            w = _sc(_sc(p1 - q1) + 3 * (q0 - p0))
            a = (27 * w + 63) >> 7
            v = q0 - a
            p[q] = 0 if v < 0 else 255 if v > 255 else v
            v = p0 + a
            p[q - across] = 0 if v < 0 else 255 if v > 255 else v
            a = (18 * w + 63) >> 7
            v = q1 - a
            p[q + across] = 0 if v < 0 else 255 if v > 255 else v
            v = p1 + a
            p[q - 2 * across] = 0 if v < 0 else 255 if v > 255 else v
            a = (9 * w + 63) >> 7
            v = q2 - a
            p[q + 2 * across] = 0 if v < 0 else 255 if v > 255 else v
            v = p2 + a
            p[q - 3 * across] = 0 if v < 0 else 255 if v > 255 else v
        else:
            a = _sc(3 * (q0 - p0))
            f1 = _sc(a + 4) >> 3
            f2 = _sc(a + 3) >> 3
            v = q0 - f1
            p[q] = 0 if v < 0 else 255 if v > 255 else v
            v = p0 + f2
            p[q - across] = 0 if v < 0 else 255 if v > 255 else v
            a = (f1 + 1) >> 1
            v = q1 - a
            p[q + across] = 0 if v < 0 else 255 if v > 255 else v
            v = p1 + a
            p[q - 2 * across] = 0 if v < 0 else 255 if v > 255 else v


def _simple(p: bytearray, pos: int, across: int, along: int, count: int, limit: int) -> None:
    """簡単なループフィルター（15.2）。"""
    for i in range(count):
        q = pos + i * along
        p1, p0, q0, q1 = p[q - 2 * across], p[q - across], p[q], p[q + across]
        if abs(p0 - q0) * 2 + (abs(p1 - q1) >> 1) > limit:
            continue
        a = _sc(_sc(p1 - q1) + 3 * (q0 - p0))
        f1 = _sc(a + 4) >> 3
        f2 = _sc(a + 3) >> 3
        v = q0 - f1
        p[q] = 0 if v < 0 else 255 if v > 255 else v
        v = p0 + f2
        p[q - across] = 0 if v < 0 else 255 if v > 255 else v


def decode_yuv(d: bytes, w: int, h: int) -> tuple:
    """キーフレームを開く：(Y, U, V, Y の行の幅, U の行の幅)。"""
    if len(d) < 10:
        raise ImageError("WebP：VP8 のデータが短すぎます")
    tag = d[0] | d[1] << 8 | d[2] << 16
    if tag & 1:
        raise ImageError("WebP：VP8 のキーフレームではありません")
    part0 = tag >> 5
    if 10 + part0 > len(d):
        raise ImageError("WebP：VP8 の最初の区画が途中で切れています")
    br = _Bool(d, 10, 10 + part0)
    rd = br.read
    br.read(128)  # color_space
    br.read(128)  # clamping_type
    seg_on = rd(128)
    update_map = 0
    seg_abs = 0
    seg_q = [0] * 4
    seg_f = [0] * 4
    seg_probs = [255, 255, 255]
    if seg_on:
        update_map = rd(128)
        if rd(128):
            seg_abs = rd(128)
            seg_q = [br.flag_signed(7) for _ in range(4)]
            seg_f = [br.flag_signed(6) for _ in range(4)]
        if update_map:
            seg_probs = [br.lit(8) if rd(128) else 255 for _ in range(3)]
    simple = rd(128)
    level = br.lit(6)
    sharp = br.lit(3)
    ref_d = [0] * 4
    mode_d = [0] * 4
    delta_on = rd(128)
    if delta_on and rd(128):
        for i in range(4):
            if rd(128):
                ref_d[i] = _signed_after_flag(br, 6)
        for i in range(4):
            if rd(128):
                mode_d[i] = _signed_after_flag(br, 6)
    nparts = 1 << br.lit(2)
    # DCT の区画
    pos = 10 + part0
    sizes_at = pos
    pos += 3 * (nparts - 1)
    if pos > len(d):
        raise ImageError("WebP：VP8 の区画の大きさの表が途中で切れています")
    readers = []
    for i in range(nparts):
        if i < nparts - 1:
            n = d[sizes_at + 3 * i] | d[sizes_at + 3 * i + 1] << 8 | d[sizes_at + 3 * i + 2] << 16
        else:
            n = len(d) - pos
        e = min(len(d), pos + n)
        readers.append(_Bool(d, pos, e))
        pos = e
    q_base = br.lit(7)
    dy1 = br.flag_signed(4)
    dy2dc = br.flag_signed(4)
    dy2ac = br.flag_signed(4)
    duvdc = br.flag_signed(4)
    duvac = br.flag_signed(4)
    # 区分ごとの量子化（14.1）
    quants = []
    for s in range(4):
        q = q_base
        if seg_on:
            q = seg_q[s] if seg_abs else q_base + seg_q[s]
        q = _clamp_q(q)
        y1 = (_DC[_clamp_q(q + dy1)], _AC[q])
        y2 = (_DC[_clamp_q(q + dy2dc)] * 2, max(8, _AC[_clamp_q(q + dy2ac)] * 101581 >> 16))
        uv = (min(132, _DC[_clamp_q(q + duvdc)]), _AC[_clamp_q(q + duvac)])
        quants.append((y1, y2, uv))
    br.read(128)  # refresh_entropy_probs
    # 係数の確率の更新（13.4）
    probs = [[[list(_DEFAULT_COEF[(i * 8 + j) * 33 + k * 11:(i * 8 + j) * 33 + k * 11 + 11]) for k in range(3)] for j in range(8)] for i in range(4)]
    idx = 0
    for i in range(4):
        for j in range(8):
            for k in range(3):
                pp = probs[i][j][k]
                for t in range(11):
                    if rd(_COEF_UPDATE[idx]):
                        pp[t] = br.lit(8)
                    idx += 1
    skip_on = rd(128)
    skip_p = br.lit(8) if skip_on else 0
    mbw = (w + 15) >> 4
    mbh = (h + 15) >> 4
    sy, suv = mbw * 16, mbw * 8
    Y = bytearray(sy * mbh * 16)
    U = bytearray(suv * mbh * 8)
    V = bytearray(suv * mbh * 8)
    top_modes = [_B_DC] * (mbw * 4)
    top_nz = [[0] * 9 for _ in range(mbw)]  # Y 4・U 2・V 2・Y2
    finfo = []
    for my in range(mbh):
        left_modes = [_B_DC] * 4
        left_nz = [0] * 9
        tok = readers[my & (nparts - 1)]
        row_info = []
        for mx in range(mbw):
            seg = 0
            if update_map:
                seg = rd(seg_probs[1]) if not rd(seg_probs[0]) else 2 + rd(seg_probs[2])
            skip = rd(skip_p) if skip_on else 0
            i4 = not rd(145)
            b_modes = None
            if not i4:
                if rd(156):
                    ymode = _TM_PRED if rd(128) else _H_PRED
                else:
                    ymode = _V_PRED if rd(163) else _DC_PRED
                for k in range(4):
                    top_modes[mx * 4 + k] = ymode
                    left_modes[k] = ymode
            else:
                ymode = _B_PRED
                b_modes = [0] * 16
                for r in range(4):
                    ym = left_modes[r]
                    for c in range(4):
                        tm = top_modes[mx * 4 + c]
                        base = (tm * 10 + ym) * 9
                        i = _B_TREE[rd(_BMODES[base])]
                        while i > 0:
                            i = _B_TREE[2 * i + rd(_BMODES[base + i])]
                        ym = -i
                        top_modes[mx * 4 + c] = ym
                        b_modes[r * 4 + c] = ym
                    left_modes[r] = ym
            if not rd(142):
                uvmode = _DC_PRED
            elif not rd(114):
                uvmode = _V_PRED
            else:
                uvmode = _TM_PRED if rd(183) else _H_PRED
            # 係数
            q1, q2, quv = quants[seg]
            tn = top_nz[mx]
            coefs_y = None
            coefs_uv = None
            nonzero = False
            if not skip:
                coefs_y = [[0] * 16 for _ in range(16)]
                first = 0
                ac = probs[3]
                if not i4:
                    dc = [0] * 16
                    nz = _coeffs(tok, probs[1], tn[8] + left_nz[8], q2, 0, dc)
                    tn[8] = left_nz[8] = 1 if nz > 0 else 0
                    if nz > 0:
                        nonzero = True
                        wht = _iwht(dc)
                        for k in range(16):
                            coefs_y[k][0] = wht[k]
                    first = 1
                    ac = probs[0]
                for r in range(4):
                    for c in range(4):
                        nz = _coeffs(tok, ac, tn[c] + left_nz[r], q1, first, coefs_y[r * 4 + c])
                        f = 1 if nz > first else 0
                        tn[c] = left_nz[r] = f
                        if f:
                            nonzero = True
                coefs_uv = [[0] * 16 for _ in range(8)]
                for plane in (0, 1):
                    base = 4 + plane * 2
                    for r in range(2):
                        for c in range(2):
                            nz = _coeffs(tok, probs[2], tn[base + c] + left_nz[base + r], quv, 0, coefs_uv[plane * 4 + r * 2 + c])
                            f = 1 if nz > 0 else 0
                            tn[base + c] = left_nz[base + r] = f
                            if f:
                                nonzero = True
            else:
                # 飛ばした MB：Y2 の文脈は B_PRED でなければ 0 にし、B_PRED なら残す（13.3）
                for k in range(8):
                    tn[k] = left_nz[k] = 0
                if not i4:
                    tn[8] = left_nz[8] = 0
            # 予測と足し算
            py = (my * sy + mx) * 16
            if not i4:
                _pred_block(Y, py, sy, 16, ymode, mx, my)
                if coefs_y is not None:
                    for k in range(16):
                        _idct_add(coefs_y[k], Y, py + (k >> 2) * 4 * sy + (k & 3) * 4, sy)
            else:
                # 右上の 4 画素：上の MB の行（ブロックの行が下でも、MB の上の行を使う）
                if my:
                    if mx + 1 < mbw:
                        ar = list(Y[py - sy + 16:py - sy + 20])
                    else:
                        ar = [Y[py - sy + 15]] * 4
                else:
                    ar = [127] * 4
                for r in range(4):
                    for c in range(4):
                        pos4 = py + r * 4 * sy + c * 4
                        if r == 0 and not my:
                            A = [127] * 8
                        else:
                            A = list(Y[pos4 - sy:pos4 - sy + 8])
                            if c == 3:
                                A[4:8] = ar
                        if c == 0 and not mx:
                            L = [129] * 4
                        else:
                            L = [Y[pos4 + t * sy - 1] for t in range(4)]
                        if r == 0 and not my:
                            P = 127
                        elif c == 0 and not mx:
                            P = 129 if (r > 0 or my) else 127
                        else:
                            P = Y[pos4 - sy - 1]
                        _pred4(Y, pos4, sy, b_modes[r * 4 + c], A, L, P)
                        if coefs_y is not None:
                            _idct_add(coefs_y[r * 4 + c], Y, pos4, sy)
            pu = (my * suv + mx) * 8
            _pred_block(U, pu, suv, 8, uvmode, mx, my)
            _pred_block(V, pu, suv, 8, uvmode, mx, my)
            if coefs_uv is not None:
                for k in range(4):
                    _idct_add(coefs_uv[k], U, pu + (k >> 1) * 4 * suv + (k & 1) * 4, suv)
                    _idct_add(coefs_uv[4 + k], V, pu + (k >> 1) * 4 * suv + (k & 1) * 4, suv)
            # ループフィルターの強さ（15.1）
            lv = level
            if seg_on:
                lv = seg_f[seg] if seg_abs else level + seg_f[seg]
            lv = 0 if lv < 0 else 63 if lv > 63 else lv
            if delta_on:
                lv += ref_d[0]
                if i4:
                    lv += mode_d[0]
                lv = 0 if lv < 0 else 63 if lv > 63 else lv
            row_info.append((lv, bool(i4 or nonzero)))
        finfo.append(row_info)
    if level or any(any(x[0] for x in r) for r in finfo):
        _loop_filter(Y, U, V, sy, suv, mbw, mbh, finfo, simple, sharp)
    return Y, U, V, sy, suv


def _signed_after_flag(br: _Bool, n: int) -> int:
    v = br.lit(n)
    return -v if br.read(128) else v


def _loop_filter(Y, U, V, sy, suv, mbw, mbh, finfo, simple, sharp) -> None:
    for my in range(mbh):
        for mx in range(mbw):
            lv, inner = finfo[my][mx]
            if not lv:
                continue
            il = lv
            if sharp:
                il >>= 2 if sharp > 4 else 1
                if il > 9 - sharp:
                    il = 9 - sharp
            if il < 1:
                il = 1
            hev = 2 if lv >= 40 else 1 if lv >= 15 else 0
            mbl = (lv + 2) * 2 + il
            sbl = lv * 2 + il
            py = (my * sy + mx) * 16
            pu = (my * suv + mx) * 8
            if simple:
                if mx:
                    _simple(Y, py, 1, sy, 16, mbl)
                if inner:
                    for k in (4, 8, 12):
                        _simple(Y, py + k, 1, sy, 16, sbl)
                if my:
                    _simple(Y, py, sy, 1, 16, mbl)
                if inner:
                    for k in (4, 8, 12):
                        _simple(Y, py + k * sy, sy, 1, 16, sbl)
                continue
            if mx:
                _edge(Y, py, 1, sy, 16, mbl, il, hev, True)
                _edge(U, pu, 1, suv, 8, mbl, il, hev, True)
                _edge(V, pu, 1, suv, 8, mbl, il, hev, True)
            if inner:
                for k in (4, 8, 12):
                    _edge(Y, py + k, 1, sy, 16, sbl, il, hev, False)
                _edge(U, pu + 4, 1, suv, 8, sbl, il, hev, False)
                _edge(V, pu + 4, 1, suv, 8, sbl, il, hev, False)
            if my:
                _edge(Y, py, sy, 1, 16, mbl, il, hev, True)
                _edge(U, pu, suv, 1, 8, mbl, il, hev, True)
                _edge(V, pu, suv, 1, 8, mbl, il, hev, True)
            if inner:
                for k in (4, 8, 12):
                    _edge(Y, py + k * sy, sy, 1, 16, sbl, il, hev, False)
                _edge(U, pu + 4 * suv, suv, 1, 8, sbl, il, hev, False)
                _edge(V, pu + 4 * suv, suv, 1, 8, sbl, il, hev, False)


def decode_vp8(d: bytes, w: int, h: int) -> bytes:
    """VP8 のキーフレーム → RGB（w*h*3）。"""
    Y, U, V, sy, suv = decode_yuv(d, w, h)
    # 色の変換（WebP の YUV は BT.601 の限定の範囲。libwebp の固定小数の式と同じ）
    ky = [(19077 * v) >> 8 for v in range(256)]
    rv = [((26149 * v) >> 8) - 14234 for v in range(256)]
    gu = [(6419 * v) >> 8 for v in range(256)]
    gv = [((13320 * v) >> 8) - 8708 for v in range(256)]
    bu = [((33050 * v) >> 8) - 17685 for v in range(256)]
    out = bytearray(w * h * 3)
    half = [x >> 1 for x in range(w)]
    for y in range(h):
        yr = Y[y * sy:y * sy + w]
        b = (y >> 1) * suv
        ur = U[b:b + (w + 1) // 2]
        vr = V[b:b + (w + 1) // 2]
        uu = [ur[i] for i in half]
        vv = [vr[i] for i in half]
        o = y * w * 3
        line_y = [ky[v] for v in yr]
        r_ = bytes([0 if t < 0 else 255 if t > 16383 else t >> 6 for t in [a + rv[v] for a, v in zip(line_y, vv)]])
        g_ = bytes([0 if t < 0 else 255 if t > 16383 else t >> 6 for t in [a - gu[u] - gv[v] for a, u, v in zip(line_y, uu, vv)]])
        b_ = bytes([0 if t < 0 else 255 if t > 16383 else t >> 6 for t in [a + bu[u] for a, u in zip(line_y, uu)]])
        line = bytearray(w * 3)
        line[0::3], line[1::3], line[2::3] = r_, g_, b_
        out[o:o + w * 3] = line
    return bytes(out)
