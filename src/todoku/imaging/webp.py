"""WebP（RFC 9649）を開く：入れ物（RIFF・VP8X・ALPH・ANIM）、可逆の VP8L。非可逆の VP8 は webp_vp8.py（RFC 6386）。"""

from __future__ import annotations

import struct

from ..limits import inflate  # noqa: F401  （ほかの形式と同じ入口の並びのため）
from .formats import Frame, FrameInfo, ImageError, _on_white, check_size


def _chunks(blob: bytes, pos: int, end: int):
    while pos + 8 <= end:
        t = blob[pos:pos + 4]
        n = struct.unpack_from("<I", blob, pos + 4)[0]
        yield t, pos + 8, min(pos + 8 + n, end)
        pos += 8 + n + (n & 1)


def _parse(blob: bytes) -> dict:
    size = struct.unpack_from("<I", blob, 4)[0] + 8
    end = min(size, len(blob))
    d: dict = {"frames": [], "canvas": None, "bg": None}
    pending_alpha = None
    for t, s, e in _chunks(blob, 12, end):
        if t == b"VP8X":
            fl = blob[s]
            w = int.from_bytes(blob[s + 4:s + 7], "little") + 1
            h = int.from_bytes(blob[s + 7:s + 10], "little") + 1
            d["canvas"] = (w, h)
            d["flags"] = fl
        elif t == b"ALPH":
            pending_alpha = (s, e)
        elif t in (b"VP8 ", b"VP8L"):
            d["frames"].append({"x": 0, "y": 0, "kind": t, "s": s, "e": e, "alpha": pending_alpha, "blend": True, "dispose": 0})
            pending_alpha = None
        elif t == b"ANMF":
            x = int.from_bytes(blob[s:s + 3], "little") * 2
            y = int.from_bytes(blob[s + 3:s + 6], "little") * 2
            fl = blob[s + 15]
            alpha = None
            for t2, s2, e2 in _chunks(blob, s + 16, e):
                if t2 == b"ALPH":
                    alpha = (s2, e2)
                elif t2 in (b"VP8 ", b"VP8L"):
                    d["frames"].append({"x": x, "y": y, "kind": t2, "s": s2, "e": e2, "alpha": alpha, "blend": not fl & 2, "dispose": fl & 1})
    return d


def _frame_size(blob: bytes, f: dict) -> tuple[int, int]:
    s = f["s"]
    if f["kind"] == b"VP8L":
        if blob[s] != 0x2F:
            raise ImageError("WebP：VP8L の印がありません")
        v = int.from_bytes(blob[s + 1:s + 5], "little")
        return (v & 0x3FFF) + 1, ((v >> 14) & 0x3FFF) + 1
    # VP8：3 バイトのフレームの札、開始の符号 9d 01 2a、14 ビットの幅と高さ
    if blob[s + 3:s + 6] != b"\x9d\x01\x2a":
        raise ImageError("WebP：VP8 のキーフレームの印がありません")
    return struct.unpack_from("<H", blob, s + 6)[0] & 0x3FFF, struct.unpack_from("<H", blob, s + 8)[0] & 0x3FFF


def infos(blob: bytes) -> list[FrameInfo]:
    d = _parse(blob)
    if not d["frames"]:
        raise ImageError("WebP：画像のデータがありません")
    if d["canvas"]:
        return [FrameInfo(d["canvas"][0], d["canvas"][1]) for _ in d["frames"]]
    w, h = _frame_size(blob, d["frames"][0])
    return [FrameInfo(w, h)]


def _decode_plain(blob: bytes, f: dict) -> tuple[int, int, bytes, bytes | None]:
    """(幅, 高さ, RGB, 透け)。"""
    w, h = _frame_size(blob, f)
    check_size(w, h, "WebP")
    if f["kind"] == b"VP8L":
        px = decode_vp8l(blob[f["s"] + 5:f["e"]], w, h)
        out = bytearray(w * h * 3)
        arr = bytearray(w * h * 4)
        struct.pack_into(f"<{w * h}I", arr, 0, *px)  # B G R A
        out[0::3], out[1::3], out[2::3] = arr[2::4], arr[1::4], arr[0::4]
        return w, h, bytes(out), bytes(arr[3::4])
    from .webp_vp8 import decode_vp8

    rgb = decode_vp8(blob[f["s"]:f["e"]], w, h)
    alpha = None
    if f.get("alpha"):
        alpha = _alpha(blob[f["alpha"][0]:f["alpha"][1]], w, h)
    return w, h, rgb, alpha


def _alpha(d: bytes, w: int, h: int) -> bytes | None:
    """ALPH（RFC 9649 5.2.2）。"""
    if not d:
        return None
    head = d[0]
    comp = head & 3
    flt = (head >> 2) & 3
    if comp == 0:
        a = bytearray(d[1:1 + w * h].ljust(w * h, b"\0"))
    elif comp == 1:
        px = decode_vp8l(d[1:], w, h)
        a = bytearray((v >> 8) & 255 for v in px)  # 緑の成分に透けが入る
    else:
        raise ImageError("WebP：透けの圧縮の種類が規格にありません")
    if flt:
        for y in range(h):
            for x in range(w):
                i = y * w + x
                L = a[i - 1] if x else (a[i - w] if y else 0)
                T = a[i - w] if y else L
                TL = a[i - w - 1] if x and y else T
                if flt == 1:
                    p = L
                elif flt == 2:
                    p = T
                else:
                    p = max(0, min(255, L + T - TL))
                if x == 0 and y == 0:
                    p = 0
                a[i] = (a[i] + p) & 255
    return bytes(a)


def decode(blob: bytes, index: int = 0) -> Frame:
    d = _parse(blob)
    frames = d["frames"]
    if not frames:
        raise ImageError("WebP：画像のデータがありません")
    if not 0 <= index < len(frames):
        raise ImageError(f"WebP：画像は {len(frames)} 枚です")
    notes = []
    if not d["canvas"] or len(frames) == 1:
        w, h, rgb, alpha = _decode_plain(blob, frames[index])
        if alpha is not None and any(v != 255 for v in alpha):
            rgb = _on_white(rgb, alpha)
            notes.append("透けのある画像を、白い紙の上の色で示しました")
        return Frame(w, h, rgb, None, notes, [])
    W, H = d["canvas"]
    check_size(W, H, "WebP")
    canvas = bytearray(b"\xff" * (W * H * 3))
    for i in range(index + 1):
        f = frames[i]
        fw, fh, rgb, alpha = _decode_plain(blob, f)
        painted = _on_white(rgb, alpha) if alpha is not None and any(v != 255 for v in alpha) else rgb
        for y in range(fh):
            Y = f["y"] + y
            n = min(fw, W - f["x"])
            if Y < H and n > 0:
                canvas[(Y * W + f["x"]) * 3:(Y * W + f["x"] + n) * 3] = painted[y * fw * 3:y * fw * 3 + n * 3]
        if i < index and f["dispose"]:
            for y in range(fh):
                Y = f["y"] + y
                n = min(fw, W - f["x"])
                if Y < H and n > 0:
                    canvas[(Y * W + f["x"]) * 3:(Y * W + f["x"] + n) * 3] = b"\xff" * (n * 3)
    notes.append(f"動く WebP の {index + 1} 枚目（全部で {len(frames)} 枚）")
    return Frame(W, H, bytes(canvas), None, notes, [])


# ====================================================================== VP8L（RFC 9649 3）
_CODE_ORDER = (17, 18, 0, 1, 2, 3, 4, 5, 16, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15)
_DIST_MAP = ((0, 1), (1, 0), (1, 1), (-1, 1), (0, 2), (2, 0), (1, 2), (-1, 2), (2, 1), (-2, 1), (2, 2), (-2, 2), (0, 3), (3, 0), (1, 3), (-1, 3), (3, 1), (-3, 1),
             (2, 3), (-2, 3), (3, 2), (-3, 2), (0, 4), (4, 0), (1, 4), (-1, 4), (4, 1), (-4, 1), (3, 3), (-3, 3), (2, 4), (-2, 4), (4, 2), (-4, 2), (0, 5), (3, 4),
             (-3, 4), (4, 3), (-4, 3), (5, 0), (1, 5), (-1, 5), (5, 1), (-5, 1), (2, 5), (-2, 5), (5, 2), (-5, 2), (4, 4), (-4, 4), (3, 5), (-3, 5), (5, 3), (-5, 3),
             (0, 6), (6, 0), (1, 6), (-1, 6), (6, 1), (-6, 1), (2, 6), (-2, 6), (6, 2), (-6, 2), (4, 5), (-4, 5), (5, 4), (-5, 4), (3, 6), (-3, 6), (6, 3), (-6, 3),
             (0, 7), (7, 0), (1, 7), (-1, 7), (5, 5), (-5, 5), (7, 1), (-7, 1), (4, 6), (-4, 6), (6, 4), (-6, 4), (2, 7), (-2, 7), (7, 2), (-7, 2), (3, 7), (-3, 7),
             (7, 3), (-7, 3), (5, 6), (-5, 6), (6, 5), (-6, 5), (8, 0), (4, 7), (-4, 7), (7, 4), (-7, 4), (8, 1), (8, 2), (6, 6), (-6, 6), (8, 3), (5, 7), (-5, 7),
             (7, 5), (-7, 5), (8, 4), (6, 7), (-6, 7), (7, 6), (-7, 6), (8, 5), (7, 7), (-7, 7), (8, 6), (8, 7))


class _BR:
    __slots__ = ("d", "pos", "bits", "n")

    def __init__(self, d: bytes):
        self.d = d
        self.pos = 0
        self.bits = 0
        self.n = 0

    def read(self, k: int) -> int:
        while self.n < k:
            self.bits |= (self.d[self.pos] if self.pos < len(self.d) else 0) << self.n
            self.pos += 1
            self.n += 8
        v = self.bits & ((1 << k) - 1)
        self.bits >>= k
        self.n -= k
        return v


class _Code:
    """正準ハフマン。lut：下位 8 ビットで引く表（(記号 << 4) | 長さ。長さ 0 は表の外）。long：長い符号の dict（(長さ, 符号)→記号）。"""

    __slots__ = ("lut", "long", "single")

    def __init__(self, lengths: list):
        nz = [(ln, s) for s, ln in enumerate(lengths) if ln]
        self.single = nz[0][1] if len(nz) == 1 else None
        self.lut = None
        self.long = None
        if self.single is not None or not nz:
            if not nz:
                self.single = 0
            return
        lut = [0] * 256
        long_: dict = {}
        code = 0
        prev = 0
        for ln, s in sorted(nz):
            code <<= ln - prev
            prev = ln
            if code >= 1 << ln:
                raise ImageError("WebP：ハフマンの符号の長さが規格に合いません")
            if ln <= 8:
                rev = int(format(code, f"0{ln}b")[::-1], 2)
                v = (s << 4) | ln
                for hi in range(1 << (8 - ln)):
                    lut[rev | (hi << ln)] = v
            else:
                long_[(ln, code)] = s
            code += 1
        self.lut = lut
        self.long = long_


def _read_code(br: _BR, alphabet: int) -> _Code:
    if br.read(1):
        num = br.read(1) + 1
        first8 = br.read(1)
        lens = [0] * max(alphabet, 256)
        s0 = br.read(1 + 7 * first8)
        lens[s0] = 1
        if num == 2:
            s1 = br.read(8)
            lens[s1] = 1
        return _Code(lens)
    n = 4 + br.read(4)
    if n > 19:
        raise ImageError("WebP：符号の長さの符号の数が規格を超えます")
    cl = [0] * 19
    for i in range(n):
        cl[_CODE_ORDER[i]] = br.read(3)
    meta = _Code(cl)
    if br.read(1):
        nb = 2 + 2 * br.read(3)
        max_sym = 2 + br.read(nb)
        if max_sym > alphabet:
            raise ImageError("WebP：符号の数が規格を超えます")
    else:
        max_sym = alphabet
    lens = [0] * alphabet
    i = 0
    prev = 8
    while i < alphabet and max_sym > 0:
        max_sym -= 1
        s = _sym(br, meta)
        if s < 16:
            lens[i] = s
            i += 1
            if s:
                prev = s
        else:
            if s == 16:
                rep = 3 + br.read(2)
                val = prev
            elif s == 17:
                rep = 3 + br.read(3)
                val = 0
            else:
                rep = 11 + br.read(7)
                val = 0
            if i + rep > alphabet:
                raise ImageError("WebP：符号の長さが規格を超えます")
            for _ in range(rep):
                lens[i] = val
                i += 1
    return _Code(lens)


def _sym(br: _BR, c: _Code) -> int:
    if c.single is not None:
        return c.single
    while br.n < 8:
        br.bits |= (br.d[br.pos] if br.pos < len(br.d) else 0) << br.n
        br.pos += 1
        br.n += 8
    v = c.lut[br.bits & 255]
    if v:
        k = v & 15
        br.bits >>= k
        br.n -= k
        return v >> 4
    # 長い符号：1 ビットずつ（符号は上の桁から読む）
    code = 0
    for ln in range(1, 16):
        code = code << 1 | br.read(1)
        s = c.long.get((ln, code))
        if s is not None:
            return s
    raise ImageError("WebP：ハフマンの符号が読めません")


def _prefix(br: _BR, sym: int) -> int:
    if sym < 4:
        return sym + 1
    extra = (sym - 2) >> 1
    return ((2 + (sym & 1)) << extra) + br.read(extra) + 1


def _image_stream(br: _BR, w: int, h: int, is_argb: bool) -> list:
    """空間の符号の画像（5.2）。is_argb：主な画像（色のキャッシュの後に、ハフマンの組の分け方がある）。"""
    cache_bits = 0
    if br.read(1):
        cache_bits = br.read(4)
        if not 1 <= cache_bits <= 11:
            raise ImageError("WebP：色のキャッシュの大きさが規格にありません")
    cache = [0] * (1 << cache_bits) if cache_bits else None
    meta = None
    prefix_bits = 0
    groups = 1
    if is_argb and br.read(1):
        prefix_bits = br.read(3) + 2
        mw = _sub(w, prefix_bits)
        mh = _sub(h, prefix_bits)
        mp = _image_stream(br, mw, mh, False)
        meta = [(v >> 8) & 0xFFFF for v in mp]
        groups = max(meta) + 1 if meta else 1
    alpha_g = 256 + 24 + (1 << cache_bits if cache_bits else 0)
    tables = []
    for _g in range(groups):
        tables.append((_read_code(br, alpha_g), _read_code(br, 256), _read_code(br, 256), _read_code(br, 256), _read_code(br, 40)))
    total = w * h
    px = [0] * total
    i = 0
    mask = (1 << prefix_bits) - 1
    mw = _sub(w, prefix_bits) if meta else 0
    cg, cr, cb, ca, cd = tables[0]
    sym = _sym
    mul_shift = 32 - cache_bits
    last_cached = 0
    while i < total:
        if meta and (i % w) & mask == 0:
            y = i // w
            cg, cr, cb, ca, cd = tables[meta[(y >> prefix_bits) * mw + ((i % w) >> prefix_bits)]]
        s = sym(br, cg)
        if s < 256:
            r = sym(br, cr)
            b = sym(br, cb)
            a = sym(br, ca)
            px[i] = (a << 24) | (r << 16) | (s << 8) | b
            i += 1
        elif s < 280:
            length = _prefix(br, s - 256)
            dsym = sym(br, cd)
            dist = _prefix(br, dsym)
            if dist > 120:
                dist -= 120
            else:
                dx, dy = _DIST_MAP[dist - 1]
                dist = dx + dy * w
                if dist < 1:
                    dist = 1
            if dist > i:
                raise ImageError("WebP：距離がデータの前を指しています")
            length = min(length, total - i)
            if dist >= length:
                px[i:i + length] = px[i - dist:i - dist + length]
            else:
                for k in range(length):
                    px[i + k] = px[i + k - dist]
            i += length
        else:
            if cache is None:
                raise ImageError("WebP：色のキャッシュの符号が読めません")
            idx = s - 280
            while last_cached < i:
                v = px[last_cached]
                cache[((v * 0x1E35A7BD) & 0xFFFFFFFF) >> mul_shift] = v
                last_cached += 1
            px[i] = cache[idx]
            i += 1
        if cache is not None and meta is None and False:
            pass
    return px


def _sub(n: int, bits: int) -> int:
    return (n + (1 << bits) - 1) >> bits


def decode_vp8l(d: bytes, w: int, h: int) -> list:
    """VP8L の本体（札と大きさの後ろ、または ALPH の中）。w・h の画素：0xAARRGGBB の int の並び。"""
    br = _BR(d)
    xsize = w
    transforms = []
    seen = set()
    while br.read(1):
        t = br.read(2)
        if t in seen:
            raise ImageError("WebP：同じ変換が 2 度あります")
        seen.add(t)
        if t in (0, 1):
            bits = br.read(3) + 2
            data = _image_stream(br, _sub(xsize, bits), _sub(h, bits), False)
            transforms.append((t, bits, data, xsize))
        elif t == 2:
            transforms.append((t, 0, None, xsize))
        else:
            n = br.read(8) + 1
            tab = _image_stream(br, n, 1, False)
            for i in range(1, n):
                tab[i] = _add(tab[i], tab[i - 1])
            wb = 3 if n <= 2 else 2 if n <= 4 else 1 if n <= 16 else 0
            transforms.append((t, wb, tab, xsize))
            xsize = _sub(xsize, wb)
    px = _image_stream(br, xsize, h, True)
    for t, bits, data, tw in reversed(transforms):
        if t == 0:
            px = _inv_predict(px, tw, h, bits, data)
        elif t == 1:
            px = _inv_color(px, tw, h, bits, data)
        elif t == 2:
            px = [(v & 0xFF00FF00) | (((v & 0xFF0000) + ((v << 8) & 0xFF0000)) & 0xFF0000) | (((v & 255) + ((v >> 8) & 255)) & 255) for v in px]
        else:
            px = _inv_index(px, tw, h, bits, data)
    return px


def _add(p: int, q: int) -> int:
    return (((p & 0xFF00FF00) + (q & 0xFF00FF00)) & 0xFF00FF00) | (((p & 0x00FF00FF) + (q & 0x00FF00FF)) & 0x00FF00FF)


def _avg(a: int, b: int) -> int:
    return (((a ^ b) & 0xFEFEFEFE) >> 1) + (a & b)


def _clip(v: int) -> int:
    return 0 if v < 0 else 255 if v > 255 else v


def _predict(mode: int, L: int, T: int, TR: int, TL: int) -> int:
    if mode == 0:
        return 0xFF000000
    if mode == 1:
        return L
    if mode == 2:
        return T
    if mode == 3:
        return TR
    if mode == 4:
        return TL
    if mode == 5:
        return _avg(_avg(L, TR), T)
    if mode == 6:
        return _avg(L, TL)
    if mode == 7:
        return _avg(L, T)
    if mode == 8:
        return _avg(TL, T)
    if mode == 9:
        return _avg(T, TR)
    if mode == 10:
        return _avg(_avg(L, TL), _avg(T, TR))
    if mode == 11:
        d = 0
        for sh in (24, 16, 8, 0):
            tl = (TL >> sh) & 255
            d += abs(((T >> sh) & 255) - tl) - abs(((L >> sh) & 255) - tl)
        return L if d < 0 else T
    if mode == 12:
        out = 0
        for sh in (24, 16, 8, 0):
            out |= _clip(((L >> sh) & 255) + ((T >> sh) & 255) - ((TL >> sh) & 255)) << sh
        return out
    if mode == 13:
        av = _avg(L, T)
        out = 0
        for sh in (24, 16, 8, 0):
            a = (av >> sh) & 255
            out |= _clip(a + (a - ((TL >> sh) & 255)) // 2 if a >= ((TL >> sh) & 255) else a - (((TL >> sh) & 255) - a) // 2) << sh
        return out
    return 0xFF000000


def _inv_predict(px: list, w: int, h: int, bits: int, data: list) -> list:
    bw = _sub(w, bits)
    out = px
    # 先頭の行：最初は黒、残りは左
    out[0] = _add(out[0], 0xFF000000)
    for x in range(1, w):
        out[x] = _add(out[x], out[x - 1])
    mask = (1 << bits) - 1
    for y in range(1, h):
        base = y * w
        up = base - w
        out[base] = _add(out[base], out[up])
        row_modes = (y >> bits) * bw
        for x in range(1, w):
            mode = (data[row_modes + (x >> bits)] >> 8) & 15 if x & mask == 0 or x == 1 else mode
            i = base + x
            L = out[i - 1]
            T = out[up + x]
            TL = out[up + x - 1]
            TR = out[up + x + 1] if x + 1 < w else out[base]
            out[i] = _add(out[i], _predict(mode, L, T, TR, TL))
    return out


def _delta(t: int, c: int) -> int:
    if t >= 128:
        t -= 256
    if c >= 128:
        c -= 256
    return (t * c) >> 5


def _inv_color(px: list, w: int, h: int, bits: int, data: list) -> list:
    bw = _sub(w, bits)
    mask = (1 << bits) - 1
    for y in range(h):
        base = y * w
        row_e = (y >> bits) * bw
        e = 0
        for x in range(w):
            if x & mask == 0:
                e = data[row_e + (x >> bits)]
                g2r = e & 255
                g2b = (e >> 8) & 255
                r2b = (e >> 16) & 255
            v = px[base + x]
            g = (v >> 8) & 255
            r = (v >> 16) & 255
            b = v & 255
            r = (r + _delta(g2r, g)) & 255
            b = (b + _delta(g2b, g) + _delta(r2b, r)) & 255
            px[base + x] = (v & 0xFF00FF00) | (r << 16) | b
    return px


def _inv_index(px: list, w: int, h: int, bits: int, tab: list) -> list:
    n = len(tab)
    tb = tab + [0] * (256 - n) if n < 256 else tab
    if bits == 0:
        return [tb[(v >> 8) & 255] for v in px]
    per = 1 << bits
    bpp = 8 >> bits
    mask = (1 << bpp) - 1
    sw = _sub(w, bits)
    out = [0] * (w * h)
    for y in range(h):
        base = y * sw
        for x in range(w):
            v = (px[base + (x >> bits)] >> 8) & 255
            out[y * w + x] = tb[(v >> ((x & (per - 1)) * bpp)) & mask]
    return out
