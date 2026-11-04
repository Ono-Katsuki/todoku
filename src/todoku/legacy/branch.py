"""実行ファイルの飛び先の変換（BCJ・ARM など）と Delta・Swap を、元に戻す（7z の入れ物が使う）。

どれも、機械語の相対の飛び先を絶対の番地にして縮みやすくしたものを、相対に戻す変換（LZMA SDK の Bra.c・Bra86.c・BraIA64.c、
Delta.c の説明のとおり。LZMA SDK は public domain、XZ は 0BSD）。すべて ip（ストリームの先頭）＝0 から。
"""

from __future__ import annotations

M32 = 0xFFFFFFFF
_ALLOWED = (1, 1, 1, 0, 1, 0, 0, 0)
_BITNUM = (0, 1, 2, 2, 3, 3, 3, 3)


def x86(data: bytes) -> bytes:
    """BCJ（x86 の E8・E9）。"""
    d = bytearray(data)
    size = len(d)
    if size < 5:
        return bytes(d)
    pos = 0
    prev_pos = -1
    prev_mask = 0
    limit = size - 4
    ip = 5
    while True:
        p = pos
        while p < limit and d[p] & 0xFE != 0xE8:
            p += 1
        pos = p
        if p >= limit:
            break
        delta = pos - prev_pos
        if delta > 3:
            prev_mask = 0
        else:
            prev_mask = (prev_mask << (delta - 1)) & 7
            if prev_mask:
                b = d[p + 4 - _BITNUM[prev_mask]]
                if not _ALLOWED[prev_mask] or b == 0 or b == 0xFF:
                    prev_pos = pos
                    prev_mask = ((prev_mask << 1) & 7) | 1
                    pos += 1
                    continue
        prev_pos = pos
        if d[p + 4] == 0 or d[p + 4] == 0xFF:
            src = d[p + 4] << 24 | d[p + 3] << 16 | d[p + 2] << 8 | d[p + 1]
            while True:
                dest = (src - (ip + pos)) & M32
                if prev_mask == 0:
                    break
                idx = _BITNUM[prev_mask] * 8
                b = (dest >> (24 - idx)) & 255
                if b != 0 and b != 0xFF:
                    break
                src = dest ^ ((1 << (32 - idx)) - 1)
            d[p + 4] = 0xFF if dest >> 24 & 1 else 0
            d[p + 3] = (dest >> 16) & 255
            d[p + 2] = (dest >> 8) & 255
            d[p + 1] = dest & 255
            pos += 5
        else:
            prev_mask = ((prev_mask << 1) & 7) | 1
            pos += 1
    return bytes(d)


def arm(data: bytes) -> bytes:
    d = bytearray(data)
    for i in range(0, len(d) - 3, 4):
        if d[i + 3] == 0xEB:
            src = (d[i + 2] << 16 | d[i + 1] << 8 | d[i]) << 2
            dest = ((src - (i + 8)) & M32) >> 2
            d[i + 2] = (dest >> 16) & 255
            d[i + 1] = (dest >> 8) & 255
            d[i] = dest & 255
    return bytes(d)


def armt(data: bytes) -> bytes:
    d = bytearray(data)
    i = 0
    n = len(d) - 3
    while i < n:
        if d[i + 1] & 0xF8 == 0xF0 and d[i + 3] & 0xF8 == 0xF8:
            src = ((d[i + 1] & 7) << 19 | d[i] << 11 | (d[i + 3] & 7) << 8 | d[i + 2]) << 1
            dest = ((src - (i + 4)) & M32) >> 1
            d[i + 1] = 0xF0 | ((dest >> 19) & 7)
            d[i] = (dest >> 11) & 255
            d[i + 3] = 0xF8 | ((dest >> 8) & 7)
            d[i + 2] = dest & 255
            i += 2
        i += 2
    return bytes(d)


def ppc(data: bytes) -> bytes:
    d = bytearray(data)
    for i in range(0, len(d) - 3, 4):
        if d[i] >> 2 == 0x12 and d[i + 3] & 3 == 1:
            src = (d[i] & 3) << 24 | d[i + 1] << 16 | d[i + 2] << 8 | (d[i + 3] & 0xFC)
            dest = (src - i) & M32
            d[i] = 0x48 | ((dest >> 24) & 3)
            d[i + 1] = (dest >> 16) & 255
            d[i + 2] = (dest >> 8) & 255
            d[i + 3] = (d[i + 3] & 3) | (dest & 0xFC)
    return bytes(d)


def sparc(data: bytes) -> bytes:
    d = bytearray(data)
    for i in range(0, len(d) - 3, 4):
        if (d[i] == 0x40 and d[i + 1] & 0xC0 == 0) or (d[i] == 0x7F and d[i + 1] & 0xC0 == 0xC0):
            src = (d[i] << 24 | d[i + 1] << 16 | d[i + 2] << 8 | d[i + 3]) << 2 & M32
            dest = ((src - i) & M32) >> 2
            dest = (((-((dest >> 22) & 1)) << 22) & 0x3FFFFFFF) | (dest & 0x3FFFFF) | 0x40000000
            d[i:i + 4] = dest.to_bytes(4, "big")
    return bytes(d)


_IA64 = (0,) * 16 + (4, 4, 6, 6, 0, 0, 7, 7, 4, 4, 0, 0, 4, 4, 0, 0)


def ia64(data: bytes) -> bytes:
    d = bytearray(data)
    for i in range(0, len(d) - 15, 16):
        mask = _IA64[d[i] & 0x1F]
        bit_pos = 5
        for slot in range(3):
            if (mask >> slot) & 1:
                byte_pos = bit_pos >> 3
                bit_res = bit_pos & 7
                ins = int.from_bytes(d[i + byte_pos:i + byte_pos + 6], "little")
                norm = ins >> bit_res
                if (norm >> 37) & 0xF == 5 and (norm >> 9) & 7 == 0:
                    src = (norm >> 13) & 0xFFFFF
                    src |= ((norm >> 36) & 1) << 20
                    src <<= 4
                    dest = ((src - i) & M32) >> 4
                    norm &= ~(0x8FFFFF << 13)
                    norm |= (dest & 0xFFFFF) << 13
                    norm |= (dest & 0x100000) << (36 - 20)
                    ins = (ins & ((1 << bit_res) - 1)) | (norm << bit_res)
                    d[i + byte_pos:i + byte_pos + 6] = (ins & ((1 << 48) - 1)).to_bytes(6, "little")
            bit_pos += 41
    return bytes(d)


def arm64(data: bytes) -> bytes:
    d = bytearray(data)
    flag = 1 << 20
    mask = (1 << 24) - (flag << 1)
    for i in range(0, len(d) - 3, 4):
        v = int.from_bytes(d[i:i + 4], "little")
        if (v - 0x94000000) & 0xFC000000 == 0:
            c = i >> 2
            v = ((v - c) & 0x03FFFFFF) | 0x94000000
            d[i:i + 4] = v.to_bytes(4, "little")
            continue
        v = (v - 0x90000000) & M32
        if v & 0x9F000000 == 0:
            v = (v + flag) & M32
            if v & mask:
                continue
            z = (v & 0xFFFFFFE0) | (v >> 26)
            c = (i >> 9) & 0xFFFFFFF8
            z = (z - c) & M32
            v &= 0x1F
            v |= 0x90000000
            v |= (z << 26) & M32
            v |= 0x00FFFFE0 & ((z & ((flag << 1) - 1)) - flag)
            d[i:i + 4] = (v & M32).to_bytes(4, "little")
    return bytes(d)


def delta(data: bytes, dist: int) -> bytes:
    d = bytearray(data)
    for k in range(dist):
        prev = 0
        for i in range(k, len(d), dist):
            prev = (prev + d[i]) & 255
            d[i] = prev
    return bytes(d)


def swap(data: bytes, width: int) -> bytes:
    d = bytearray(data)
    for i in range(0, len(d) - width + 1, width):
        d[i:i + width] = d[i:i + width][::-1]
    return bytes(d)


def bcj2(main: bytes, call: bytes, jump: bytes, rc: bytes, size: int) -> bytes:
    """BCJ2：主な流れ・CALL の番地・JUMP の番地・範囲符号の 4 つの流れから 1 つに戻す。"""
    p = [1024] * 258
    if len(rc) < 5:
        raise ValueError("BCJ2 の範囲符号が短すぎます")
    code = int.from_bytes(rc[:5], "big") & M32
    rng = M32
    rp = 5
    out = bytearray()
    ip = cp = jp = 0
    prev = 0
    nmain = len(main)
    nrc = len(rc)
    while True:
        limit = min(nmain - ip, size - len(out))
        hit = False
        while limit:
            b = main[ip]
            out.append(b)
            if b & 0xFE == 0xE8 or (prev == 0x0F and b & 0xF0 == 0x80):
                hit = True
                break
            ip += 1
            prev = b
            limit -= 1
        if not hit or len(out) == size:
            break
        b = main[ip]
        ip += 1
        k = prev if b == 0xE8 else 256 if b == 0xE9 else 257
        t = p[k]
        bound = (rng >> 11) * t
        if code < bound:
            rng = bound
            p[k] = t + ((2048 - t) >> 5)
            if rng < 0x1000000:
                rng = (rng << 8) & M32
                code = ((code << 8) | (rc[rp] if rp < nrc else 0)) & M32
                rp += 1
            prev = b
            continue
        rng -= bound
        code -= bound
        p[k] = t - (t >> 5)
        if rng < 0x1000000:
            rng = (rng << 8) & M32
            code = ((code << 8) | (rc[rp] if rp < nrc else 0)) & M32
            rp += 1
        if b == 0xE8:
            v = call[cp:cp + 4]
            cp += 4
        else:
            v = jump[jp:jp + 4]
            jp += 4
        if len(v) < 4:
            raise ValueError("BCJ2 の番地の流れが足りません")
        dest = (int.from_bytes(v, "big") - (len(out) + 4)) & M32
        room = size - len(out)
        out += dest.to_bytes(4, "little")[:room]
        if room <= 4:
            break
        prev = dest >> 24
    if len(out) != size:
        raise ValueError("BCJ2 の結果の大きさが宣言と合いません")
    return bytes(out)
