"""PICT（Apple QuickDraw の絵。Inside Macintosh: Imaging With QuickDraw の付録 A）を画素に描く：版 1（1 バイトの命令）・版 2（2 バイトの命令）。

描くもの：ビットマップ（BitsRect・PackBitsRect・DirectBitsRect。1〜32 ビット、PackBits）・線・四角・角丸・楕円・円弧・多角形・領域・字・
QuickTime で埋め込んだ画像（JPEG・PNG・TIFF など、この部品で開ける形式）。描くのはメタファイルと同じ部品（metaraster._Player）。
描けないものは黙って飛ばさず、種類と数を problems に書く。

ファイルの決まり：先頭に 512 バイトのアプリケーションの見出しがあることもないこともある。その次に picSize（2）・picFrame（8）、版の命令。
"""

from __future__ import annotations

import struct

from .formats import Frame, FrameInfo, ImageError, check_size, decode

_QD_COLORS = {33: (0, 0, 0), 30: (255, 255, 255), 205: (221, 8, 6), 341: (0, 128, 17), 409: (0, 0, 212), 273: (2, 171, 234), 137: (242, 8, 132), 69: (252, 243, 5)}
BLACK = b"\xff" * 8


def _start(blob: bytes) -> int | None:
    """絵の始まり（picSize の位置）。版の命令を見て決める。"""
    for base in (512, 0):
        if len(blob) < base + 14:
            continue
        v = blob[base + 10:base + 14]
        if v == b"\x00\x11\x02\xff":
            return base
        if v[:2] == b"\x11\x01" and _sane(blob, base) and (base == 0 or not any(blob[:512])):
            return base
    return None


def _sane(blob: bytes, base: int) -> bool:
    t, l, b, r = struct.unpack_from(">hhhh", blob, base + 2)
    return b > t and r > l


def is_pict(blob: bytes) -> bool:
    return _start(blob) is not None


def _frame_of(blob: bytes) -> tuple[int, int, int, int, int]:
    base = _start(blob)
    if base is None:
        raise ImageError("PICT：版の命令がありません")
    t, l, b, r = struct.unpack_from(">hhhh", blob, base + 2)
    return base, t, l, b, r


def _resolution(blob: bytes, base: int) -> tuple | None:
    """版 2 の拡張の見出し（HeaderOp、版 −2）の解像度（dpi）。"""
    if blob[base + 10:base + 16] == b"\x00\x11\x02\xff\x0c\x00":
        ver = struct.unpack_from(">h", blob, base + 16)[0]
        if ver == -2:
            h, v = struct.unpack_from(">II", blob, base + 20)
            if h and v:
                return (h / 65536, v / 65536)
    return None


def _ext_frame(blob: bytes, base: int, t, l, b, r) -> tuple:
    """拡張の見出しの「元の解像度の枠」があれば、その画素の大きさを画像の大きさにする（72 dpi の枠を 300 dpi の画素で描く絵）。"""
    if blob[base + 10:base + 16] == b"\x00\x11\x02\xff\x0c\x00" and struct.unpack_from(">h", blob, base + 16)[0] == -2:
        y0, x0, y1, x1 = struct.unpack_from(">hhhh", blob, base + 28)
        if y1 > y0 and x1 > x0:
            return y0, x0, y1, x1
    return t, l, b, r


def infos(blob: bytes) -> list[FrameInfo]:
    base, t, l, b, r = _frame_of(blob)
    t, l, b, r = _ext_frame(blob, base, t, l, b, r)
    if b <= t or r <= l:
        raise ImageError("PICT：絵の枠の大きさが 0 です")
    return [FrameInfo(r - l, b - t, _resolution(blob, base) or (72.0, 72.0))]


def _packbits_row(d: bytes, want: int, unit: int = 1) -> bytes:
    """PackBits（TN1023）。unit 2 は 16 ビットの単位で繰り返す。"""
    out = bytearray()
    i, n = 0, len(d)
    while i < n and len(out) < want:
        c = d[i]
        i += 1
        if c < 128:
            k = (c + 1) * unit
            out += d[i:i + k]
            i += k
        elif c != 128:
            out += d[i:i + unit] * (257 - c)
            i += unit
    return bytes(out[:want]).ljust(want, b"\0")


def _expands_to(d: bytes, want: int, unit: int) -> bool:
    """PackBits がちょうど want バイトに開いて、データを使い切るか。"""
    i, n, got = 0, len(d), 0
    while i < n:
        c = d[i]
        i += 1
        if c < 128:
            got += (c + 1) * unit
            i += (c + 1) * unit
        elif c != 128:
            got += (257 - c) * unit
            i += unit
    return i == n and got == want


def _count_size(rd: "_Reader", row: int, want: int, unit: int) -> int:
    """行の長さの記録のバイト数。QuickDraw の決まりは rowBytes > 250 なら 2。ただしそうでなく書く道具もあるので、
    決まりどおりの読み方で行がちょうど開かないときだけ、もう一方を試す。"""
    first = 2 if row > 250 else 1
    for size in (first, 3 - first):
        head = rd.b[rd.p:rd.p + size]
        if len(head) < size:
            break
        n = int.from_bytes(head, "big")
        if _expands_to(rd.b[rd.p + size:rd.p + size + n], want, unit):
            return size
    return first


class _Reader:
    def __init__(self, blob: bytes, pos: int):
        self.b, self.p = blob, pos

    def u8(self) -> int:
        v = self.b[self.p]
        self.p += 1
        return v

    def s8(self) -> int:
        v = self.u8()
        return v - 256 if v >= 128 else v

    def u16(self) -> int:
        v = struct.unpack_from(">H", self.b, self.p)[0]
        self.p += 2
        return v

    def s16(self) -> int:
        v = struct.unpack_from(">h", self.b, self.p)[0]
        self.p += 2
        return v

    def u32(self) -> int:
        v = struct.unpack_from(">I", self.b, self.p)[0]
        self.p += 4
        return v

    def rect(self) -> tuple:
        v = struct.unpack_from(">hhhh", self.b, self.p)
        self.p += 8
        return v  # (上, 左, 下, 右)

    def point(self) -> tuple:
        y, x = struct.unpack_from(">hh", self.b, self.p)
        self.p += 4
        return x, y

    def rgb(self) -> tuple:
        r, g, b = struct.unpack_from(">HHH", self.b, self.p)
        self.p += 6
        return r >> 8, g >> 8, b >> 8

    def take(self, n: int) -> bytes:
        v = self.b[self.p:self.p + n]
        self.p += n
        return v


def _pixmap(rd: _Reader, row_bytes: int) -> dict:
    """PixMap の rowBytes の次から（bounds・版…pmTable・pmReserved）。"""
    bounds = rd.rect()
    rd.u16()  # pmVersion
    pack_type = rd.u16()
    rd.u32()  # packSize
    rd.u32()
    rd.u32()  # 解像度
    rd.u16()  # pixelType
    pixel_size = rd.u16()
    cmp_count = rd.u16()
    cmp_size = rd.u16()
    rd.u32()
    rd.u32()
    rd.u32()
    return {"row": row_bytes & 0x3FFF, "bounds": bounds, "pack": pack_type, "bits": pixel_size, "cc": cmp_count, "cs": cmp_size}


def _color_table(rd: _Reader) -> dict:
    rd.u32()
    flags = rd.u16()
    n = rd.u16()
    tab = {}
    for i in range(n + 1):
        idx = rd.u16()
        c = rd.rgb()
        tab[i if flags & 0x8000 else idx] = c
    return tab


def _bits_to_rgb(rd: _Reader, pm: dict, table: dict | None, fg, bg) -> bytes:
    """行ごとのデータ → RGB。"""
    t, l, b, r = pm["bounds"]
    w, h = r - l, b - t
    row = pm["row"]
    bits = pm["bits"]
    pack = pm["pack"]
    check_size(max(1, w), max(1, h), "PICT")
    out = bytearray(w * h * 3)
    cc = pm.get("cc", 1)
    # 1 行の展開後のバイト数
    if bits == 32 and pack == 2:
        want = w * 3
    elif bits == 32 and pack == 4:
        want = w * cc
    else:
        want = row
    packed = row >= 8 and pack != 1
    if bits == 32 and pack == 0:
        pack = 4 if cc in (3, 4) else 1
        packed = row >= 8
        want = w * cc if pack == 4 else row
    if bits == 32 and pack == 1:
        packed, want = False, row
    unit = 2 if bits == 16 and pack in (0, 3) else 1
    pal = None
    if table is not None or bits <= 8:
        if table is None:
            table = {0: bg, 1: fg}
        pal = [table.get(i, (0, 0, 0)) for i in range(256)]
        r_t, g_t, b_t = bytes(c[0] for c in pal), bytes(c[1] for c in pal), bytes(c[2] for c in pal)
    from .formats import _samples_row

    count_size = 0
    for y in range(h):
        if rd.p >= len(rd.b):
            break
        if packed:
            if not count_size:
                count_size = _count_size(rd, row, want, unit)
            n = rd.u16() if count_size == 2 else rd.u8()
            data = _packbits_row(rd.take(n), want, unit)
        else:
            data = rd.take(want).ljust(want, b"\0")
        o = y * w * 3
        line = bytearray(w * 3)
        if bits <= 8:
            v = data[:w] if bits == 8 else _samples_row(data, bits, w)
            line[0::3], line[1::3], line[2::3] = v.translate(r_t), v.translate(g_t), v.translate(b_t)
        elif bits == 16:
            vs = struct.unpack_from(f">{w}H", data.ljust(w * 2, b"\0"))
            exp = [(i * 255 + 15) // 31 for i in range(32)]
            line[0::3] = bytes([exp[(a >> 10) & 31] for a in vs])
            line[1::3] = bytes([exp[(a >> 5) & 31] for a in vs])
            line[2::3] = bytes([exp[a & 31] for a in vs])
        elif pack == 2:
            line = bytearray(data[:w * 3])
        elif pack == 4:  # 面ごと（cc＝4 なら A R G B）
            s = cc - 3
            line[0::3] = data[s * w:(s + 1) * w]
            line[1::3] = data[(s + 1) * w:(s + 2) * w]
            line[2::3] = data[(s + 2) * w:(s + 3) * w]
        else:  # 32 ビットの ARGB をそのまま
            d = data.ljust(w * 4, b"\0")
            line[0::3], line[1::3], line[2::3] = d[1::4][:w], d[2::4][:w], d[3::4][:w]
        out[o:o + w * 3] = line
    return bytes(out)


def _region(rd: _Reader) -> tuple:
    """領域（IM: QuickDraw 付録）：(外枠, 四角の並び)。行ごとの反転の並びを、(y0, x0, y1, x1) の四角に開く。"""
    start = rd.p
    size = rd.u16()
    box = rd.rect()
    end = start + size
    rects = []
    if size <= 10:
        rd.p = max(rd.p, end)
        return box, [box]
    active: list = []
    p = _Reader(rd.b, rd.p)
    last_y = None
    while p.p + 2 <= end:
        y = p.s16()
        if y == 32767:
            break
        if last_y is not None and active:
            for i in range(0, len(active) - 1, 2):
                rects.append((last_y, active[i], y, active[i + 1]))
        xs = []
        while p.p + 2 <= end:
            x = p.s16()
            if x == 32767:
                break
            xs.append(x)
        now = set(active) ^ set(xs)
        active = sorted(now)
        last_y = y
    rd.p = end
    return box, rects


class _Pict:
    def __init__(self, blob: bytes, base: int, t: int, l: int, b: int, r: int):
        from .metaraster import _Player

        self.blob = blob
        self.W, self.H = r - l, b - t
        check_size(self.W, self.H, "PICT")
        self.pl = _Player(self.W, self.H, (1.0, 0.0, 0.0, 1.0, -float(l), -float(t)))
        self.pl.st.brush = None
        self.t, self.l = t, l
        self.fg, self.bg = (0, 0, 0), (255, 255, 255)
        self.pen = (1, 1)
        self.pat = {"pen": BLACK, "fill": BLACK, "bk": b"\0" * 8}
        self.loc = (0, 0)
        self.last = {"rect": None, "rrect": None, "oval": None, "arc": None, "poly": None, "rgn": None}
        self.ov = (0, 0)
        self.org = (0, 0)
        self.font = {"size": 12, "face": 0, "id": 0}
        self.tloc = (0, 0)

    # ---- 色と型
    def _pat_color(self, which: str):
        """型の色：全部 1 は前景、全部 0 は背景。混ざる型は (色の関数)。"""
        pat = self.pat[which]
        if pat == BLACK:
            return self.fg, None
        if pat == b"\0" * 8:
            return self.bg, None
        fg, bg = self.fg, self.bg
        l, t = self.l, self.t

        def at(x, y):
            xx, yy = int(x) + l, int(y) + t
            return fg if (pat[yy & 7] >> (7 - (xx & 7))) & 1 else bg

        return None, at

    def _fill(self, polys, verb: int) -> None:
        from ..pdf.render.canvas import fill_rows, paint, paint_fn

        st = self.pl.st
        polys = [p for p in polys if len(p) >= 3]
        if not polys or st.clip.empty:
            return
        rows = fill_rows(polys, False, st.clip.rect)
        self.pl.drawn += 1
        if verb == 3:  # invert
            self._invert(rows)
            return
        which = {0: "pen", 1: "pen", 2: "bk", 4: "fill"}[verb]
        c, fn = self._pat_color(which)
        if fn is not None:
            paint_fn(self.pl.cv, rows, fn, 1.0, st.clip)
        else:
            paint(self.pl.cv, rows, c, 1.0, st.clip)

    def _invert(self, rows) -> None:
        cv = self.pl.cv
        clip = self.pl.st.clip
        for Y, runs in rows:
            if not clip.y0 <= Y < clip.y1:
                continue
            for x0, x1, a in runs:
                if a < 128:
                    continue
                x0, x1 = max(x0, clip.x0), min(x1, clip.x1)
                p0, p1 = (Y * cv.w + x0) * 3, (Y * cv.w + x1) * 3
                cv.buf[p0:p1] = bytes(255 - v for v in cv.buf[p0:p1])

    def P(self, x, y) -> tuple:
        return float(x - self.l), float(y - self.t)

    def line(self, x0, y0, x1, y1) -> None:
        """線：ペンの四角（左上が点）を動かした形。"""
        pw, ph = self.pen
        if pw <= 0 or ph <= 0:
            return
        pts = []
        for x, y in ((x0, y0), (x1, y1)):
            pts += [(x, y), (x + pw, y), (x + pw, y + ph), (x, y + ph)]
        hull = _hull([self.P(*p) for p in pts])
        self._fill([hull], 1)

    def rect_poly(self, t, l, b, r) -> list:
        return [self.P(l, t), self.P(r, t), self.P(r, b), self.P(l, b)]

    def frame_polys(self, outer: list, inner: list) -> list:
        return [outer, inner[::-1]]

    def shape(self, kind: str, verb: int, a) -> None:
        import math

        pw, ph = self.pen
        if kind == "rect":
            t, l, b, r = a
            if verb == 0:
                self._fill_ring(self.rect_poly(t, l, b, r), self.rect_poly(t + ph, l + pw, b - ph, r - pw))
            else:
                self._fill([self.rect_poly(t, l, b, r)], verb)
        elif kind == "oval":
            t, l, b, r = a
            outer = _ellipse(l, t, r, b, self.P)
            if verb == 0:
                self._fill_ring(outer, _ellipse(l + pw, t + ph, r - pw, b - ph, self.P))
            else:
                self._fill([outer], verb)
        elif kind == "rrect":
            t, l, b, r = a
            ow, oh = self.ov
            outer = _rrect(l, t, r, b, ow, oh, self.P)
            if verb == 0:
                self._fill_ring(outer, _rrect(l + pw, t + ph, r - pw, b - ph, max(0, ow - 2 * pw), max(0, oh - 2 * ph), self.P))
            else:
                self._fill([outer], verb)
        elif kind == "arc":
            (t, l, b, r), start, ang = a
            cx, cy, rx, ry = (l + r) / 2, (t + b) / 2, (r - l) / 2, (b - t) / 2
            n = max(8, int(abs(ang) / 5))
            arc = [self.P(cx + rx * math.sin(math.radians(start + ang * i / n)), cy - ry * math.cos(math.radians(start + ang * i / n))) for i in range(n + 1)]
            if verb == 0:
                from ..pdf.render.canvas import stroke_polys

                polys = stroke_polys([(arc, False)], max(pw, ph), 0, 0, 10.0, None, 0.0, 1.0)
                self._fill(polys, 1)
            else:
                self._fill([arc + [self.P(cx, cy)]], verb)
        elif kind == "poly":
            pts = a
            if verb == 0:
                for p, q in zip(pts, pts[1:]):
                    self.line(p[0], p[1], q[0], q[1])
            else:
                self._fill([[self.P(*p) for p in pts]], verb)
        elif kind == "rgn":
            _box, rects = a
            if verb == 0:
                self.pl.skip("領域の枠（塗りで描きました）")
            for t, l, b, r in rects:
                self._fill([self.rect_poly(t, l, b, r)], 1 if verb == 0 else verb)

    def _fill_ring(self, outer: list, inner: list) -> None:
        from ..pdf.render.canvas import fill_rows, paint, paint_fn

        st = self.pl.st
        if st.clip.empty:
            return
        rows = fill_rows([outer, inner] if len(inner) >= 3 else [outer], True, st.clip.rect)
        c, fn = self._pat_color("pen")
        self.pl.drawn += 1
        if fn is not None:
            paint_fn(self.pl.cv, rows, fn, 1.0, st.clip)
        else:
            paint(self.pl.cv, rows, c, 1.0, st.clip)

    # ---- 字
    def text(self, data: bytes, x: int, y: int) -> None:
        fid = self.font["id"]
        # スクリプトの番号の範囲（Inside Macintosh: Text）：0x4000〜 日本語（ほかは Mac OS Roman）
        enc = "shift_jis" if 0x4000 <= fid < 0x4200 else "mac_roman"
        s = data.decode(enc, "replace")
        pl = self.pl
        pl.st.font = {"h": float(self.font["size"]), "weight": 700 if self.font["face"] & 1 else 400, "italic": bool(self.font["face"] & 2), "esc": 0.0, "face": ""}
        pl.st.tc = self.fg
        pl.st.ta = 24  # 基線
        pl.draw_text(s, float(x), float(y), None)

    # ---- ビットマップ
    def draw_pixels(self, rgb: bytes, w: int, h: int, bounds, src, dst, mode: int, stencil: bool = False, rects: list | None = None) -> None:
        from ..pdf.render.image import RGBA, draw_image

        st = self.pl.st
        bt, bl, _bb, _br = bounds
        st_, sl, sb, sr = src
        dt, dl, db, dr = dst
        # 元の四角の、ビットマップの中の位置
        x0, y0 = max(0, sl - bl), max(0, st_ - bt)
        x1, y1 = min(w, sr - bl), min(h, sb - bt)
        if x1 <= x0 or y1 <= y0 or dr <= dl or db <= dt:
            return
        px = b"".join(rgb[(j * w + x0) * 3:(j * w + x1) * 3] for j in range(y0, y1))
        cw, ch = x1 - x0, y1 - y0
        X0, Y0 = dl - self.l, dt - self.t
        X1, Y1 = dr - self.l, db - self.t
        clip = st.clip
        if rects:
            from ..pdf.render.canvas import fill_rows

            polys = [[(l - self.l, t - self.t), (r - self.l, t - self.t), (r - self.l, b - self.t), (l - self.l, b - self.t)] for t, l, b, r in rects]
            clip = clip.with_rows(fill_rows(polys, False, clip.rect))
        al = None
        if stencil and mode in (1, 9, 33):  # srcOr・notPatOr：白は透ける
            al = px[0::3].translate(bytes(0 if v == 255 else 255 for v in range(256)))
        draw_image(self.pl.cv, RGBA(cw, ch, px, al), (float(X1 - X0), 0.0, 0.0, -float(Y1 - Y0), float(X0), float(Y1)), clip)
        self.pl.drawn += 1

    def bits(self, rd: _Reader, op: int, version: int) -> None:
        with_rgn = op in (0x91, 0x99, 0x9B)
        if op in (0x9A, 0x9B):
            rd.u32()  # baseAddr
        row = rd.u16()
        is_pix = bool(row & 0x8000) or op in (0x9A, 0x9B)
        table = None
        if is_pix:
            pm = _pixmap(rd, row)
            if op in (0x90, 0x91, 0x98, 0x99):
                table = _color_table(rd)
        else:
            pm = {"row": row & 0x3FFF, "bounds": rd.rect(), "pack": 0, "bits": 1, "cc": 1, "cs": 1}
        src, dst = rd.rect(), rd.rect()
        mode = rd.u16()
        rects = None
        if with_rgn:
            _b, rects = _region(rd)
        if op in (0x90, 0x91):
            pm["pack"] = 1  # 圧縮なし
        rgb = _bits_to_rgb(rd, pm, table, self.fg, self.bg)
        t, l, b, r = pm["bounds"]
        self.draw_pixels(rgb, r - l, b - t, pm["bounds"], src, dst, mode, not is_pix, rects)

    def quicktime(self, data: bytes, compressed: bool) -> None:
        """QuickTime の命令（0x8200）：行列・見出しのあとの、埋め込まれた画像。"""
        if not compressed:
            self.pl.skip("圧縮していない QuickTime の画像")
            return
        try:
            ma = struct.unpack_from(">9i", data, 2)
            matte_size = struct.unpack_from(">I", data, 38)[0]
            mask_size = struct.unpack_from(">I", data, 62)[0]
            at = 68 + matte_size + mask_size
            desc = struct.unpack_from(">I", data, at)[0]
            codec = data[at + 4:at + 8]
            iw, ih = struct.unpack_from(">HH", data, at + 32)
            body = data[at + desc:]
        except (struct.error, IndexError):
            self.pl.skip("読めない QuickTime の画像")
            return
        try:
            from .formats import sniff

            if sniff(body) is None:
                raise ImageError(codec.decode("latin-1"))
            fr = decode(body)
        except Exception as e:  # noqa: BLE001
            self.pl.skip(f"開けない QuickTime の画像（{codec.decode('latin-1')}：{type(e).__name__}）")
            return
        a, d, tx, ty = ma[0] / 65536, ma[4] / 65536, ma[6] / 65536, ma[7] / 65536
        iw, ih = iw or fr.w, ih or fr.h
        X0, Y0 = tx - self.l, ty - self.t
        X1, Y1 = X0 + iw * a, Y0 + ih * d
        from ..pdf.render.image import RGBA, draw_image

        draw_image(self.pl.cv, RGBA(fr.w, fr.h, fr.rgb, None), (X1 - X0, 0.0, 0.0, -(Y1 - Y0), X0, Y1), self.pl.st.clip)
        self.pl.drawn += 1

    def pixpat(self, rd: _Reader, which: str) -> None:
        typ = rd.u16()
        pat = rd.take(8)
        self.pat[which] = pat
        if typ == 1:
            row = rd.u16()
            pm = _pixmap(rd, row)
            table = _color_table(rd)
            t, l, b, r = pm["bounds"]
            pm["pack"] = pm["pack"] or 0
            rgb = _bits_to_rgb(rd, pm, table, self.fg, self.bg)
            w, h = r - l, b - t
            if w and h:
                # 模様の平均の色で塗る（近い形）
                n = w * h
                c = (sum(rgb[0::3]) // n, sum(rgb[1::3]) // n, sum(rgb[2::3]) // n)
                self.pat[which] = BLACK
                self._pixpat_color(which, c)
            self.pl.skip("色の模様（平均の色で描きました）")
        elif typ == 2:
            c = rd.rgb()
            self.pat[which] = BLACK
            self._pixpat_color(which, c)

    def _pixpat_color(self, which: str, c: tuple) -> None:
        # 模様の色は前景を変えない：型ごとの色として持つ代わりに、ふつう模様の次に塗るので前景に入れる
        if which != "bk":
            self.fg = c
        else:
            self.bg = c

    # ---- 命令の列
    def run(self, rd: _Reader, version: int, limit_ops: int = 2_000_000) -> None:
        ops = 0
        pl = self.pl
        n = len(rd.b)
        while rd.p < n:
            ops += 1
            if ops > limit_ops:
                pl.skip("命令が多すぎるので止めました")
                return
            if version == 2:
                if rd.p & 1 != self.align:
                    rd.p += 1
                if rd.p + 2 > n:
                    return
                op = rd.u16()
            else:
                op = rd.u8()
            if op == 0x00FF or (version == 1 and op == 0xFF):
                return
            try:
                self.op(rd, op, version)
            except (struct.error, IndexError):
                pl.skip(f"途中で切れた命令 0x{op:04X}")
                return

    def op(self, rd: _Reader, op: int, version: int) -> None:
        pl = self.pl
        if op == 0x0000 or op in (0x001C, 0x001E) or 0x17 <= op <= 0x19:
            return
        if op == 0x0001:
            box, rects = _region(rd)
            t, l, b, r = box
            from ..pdf.render.canvas import Clip, fill_rows

            base = Clip(0, 0, self.W, self.H)
            clip = base.with_rect(l - self.l, t - self.t, r - self.l, b - self.t)
            if len(rects) > 1:
                polys = [[(x0 - self.l, y0 - self.t), (x1 - self.l, y0 - self.t), (x1 - self.l, y1 - self.t), (x0 - self.l, y1 - self.t)] for y0, x0, y1, x1 in rects]
                clip = clip.with_rows(fill_rows(polys, False, clip.rect))
            pl.st.clip = clip
        elif op == 0x0002:
            self.pat["bk"] = rd.take(8)
        elif op == 0x0003:
            self.font["id"] = rd.u16()
        elif op == 0x0004:
            self.font["face"] = rd.u8()
        elif op == 0x0005 or op == 0x0008 or op == 0x0015 or op == 0x0016 or op == 0x00A0:
            rd.p += 2
        elif op == 0x0006:
            rd.p += 4
        elif op == 0x0007:
            h, w = rd.s16(), rd.s16()
            self.pen = (w, h)
        elif op == 0x0009:
            self.pat["pen"] = rd.take(8)
        elif op == 0x000A:
            self.pat["fill"] = rd.take(8)
        elif op == 0x000B:
            h, w = rd.s16(), rd.s16()
            self.ov = (w, h)
        elif op == 0x000C:
            dh, dv = rd.s16(), rd.s16()
            self.l += dh
            self.t += dv
            pl.base = (1.0, 0.0, 0.0, 1.0, -float(self.l), -float(self.t))
        elif op == 0x000D:
            self.font["size"] = rd.u16()
        elif op == 0x000E:
            self.fg = _QD_COLORS.get(rd.u32(), self.fg)
        elif op == 0x000F:
            self.bg = _QD_COLORS.get(rd.u32(), self.bg)
        elif op == 0x0010:
            rd.p += 8
        elif op == 0x0011:
            rd.p += 1 if version == 1 else 2
        elif op in (0x0012, 0x0013, 0x0014):
            self.pixpat(rd, {0x12: "bk", 0x13: "pen", 0x14: "fill"}[op])
        elif op == 0x001A:
            self.fg = rd.rgb()
        elif op == 0x001B:
            self.bg = rd.rgb()
        elif op in (0x001D, 0x001F):
            rd.p += 6
        elif op == 0x0020:
            y0, x0, y1, x1 = rd.s16(), rd.s16(), rd.s16(), rd.s16()
            self.line(x0, y0, x1, y1)
            self.loc = (x1, y1)
        elif op == 0x0021:
            y1, x1 = rd.s16(), rd.s16()
            self.line(self.loc[0], self.loc[1], x1, y1)
            self.loc = (x1, y1)
        elif op == 0x0022:
            y0, x0 = rd.s16(), rd.s16()
            dh, dv = rd.s8(), rd.s8()
            self.line(x0, y0, x0 + dh, y0 + dv)
            self.loc = (x0 + dh, y0 + dv)
        elif op == 0x0023:
            dh, dv = rd.s8(), rd.s8()
            x0, y0 = self.loc
            self.line(x0, y0, x0 + dh, y0 + dv)
            self.loc = (x0 + dh, y0 + dv)
        elif op == 0x0028:
            y, x = rd.s16(), rd.s16()
            c = rd.u8()
            self.tloc = (x, y)
            self.text(rd.take(c), x, y)
        elif op == 0x0029:
            dh = rd.u8()
            c = rd.u8()
            self.tloc = (self.tloc[0] + dh, self.tloc[1])
            self.text(rd.take(c), *self.tloc)
        elif op == 0x002A:
            dv = rd.u8()
            c = rd.u8()
            self.tloc = (self.tloc[0], self.tloc[1] + dv)
            self.text(rd.take(c), *self.tloc)
        elif op == 0x002B:
            dh, dv = rd.u8(), rd.u8()
            c = rd.u8()
            self.tloc = (self.tloc[0] + dh, self.tloc[1] + dv)
            self.text(rd.take(c), *self.tloc)
        elif 0x0030 <= op <= 0x0037 or 0x0038 <= op <= 0x003F:
            self._geom(rd, op, "rect", 0x30, 8)
        elif 0x0040 <= op <= 0x004F:
            self._geom(rd, op, "rrect", 0x40, 8)
        elif 0x0050 <= op <= 0x005F:
            self._geom(rd, op, "oval", 0x50, 8)
        elif 0x0060 <= op <= 0x006F:
            self._arc(rd, op)
        elif 0x0070 <= op <= 0x007F:
            self._poly(rd, op)
        elif 0x0080 <= op <= 0x008F:
            self._rgn(rd, op)
        elif op in (0x0090, 0x0091, 0x0098, 0x0099, 0x009A, 0x009B):
            self.bits(rd, op, version)
        elif op == 0x00A1:
            rd.p += 2
            n = rd.u16()
            rd.p += n
        elif op == 0x0C00:
            rd.p += 24
        elif op in (0x8200, 0x8201):
            n = rd.u32()
            self.quicktime(rd.take(n), op == 0x8200)
        elif op in (0x24, 0x25, 0x26, 0x27, 0x2C, 0x2D, 0x2E, 0x2F) or 0x92 <= op <= 0x97 or 0x9C <= op <= 0x9F or 0xA2 <= op <= 0xAF:
            n = rd.u16()
            data = rd.take(n)
            if op == 0x2C and len(data) >= 2:  # FontName：番号と名前
                self.font["id"] = struct.unpack(">H", data[:2])[0]
        elif 0xB0 <= op <= 0xCF:
            pass
        elif 0xD0 <= op <= 0xFE:
            n = rd.u32()
            rd.p += n
        elif 0x0100 <= op <= 0x7FFF:
            rd.p += (op >> 8) * 2
        elif 0x8000 <= op <= 0x80FF:
            pass
        else:
            n = rd.u32()
            rd.p += n

    def _verb(self, op: int, base: int) -> tuple[int, bool]:
        sub = op - base
        same = sub >= 8
        v = sub - 8 if same else sub
        return v, same

    def _geom(self, rd, op, kind, base, size) -> None:
        v, same = self._verb(op, base)
        if not same:
            self.last[kind] = rd.rect()
        if v > 4:
            if not same:
                pass
            return
        if self.last[kind] is not None:
            self.shape(kind, v, self.last[kind])

    def _arc(self, rd, op) -> None:
        v, same = self._verb(op, 0x60)
        if not same:
            self.last["arc"] = rd.rect()
        start, ang = rd.s16(), rd.s16()
        if v <= 4 and self.last["arc"] is not None:
            self.shape("arc", v, (self.last["arc"], start, ang))

    def _poly(self, rd, op) -> None:
        v, same = self._verb(op, 0x70)
        if not same:
            size = rd.u16()
            rd.rect()
            n = (size - 10) // 4
            pts = []
            for _ in range(max(0, n)):
                pts.append(rd.point())
            self.last["poly"] = pts
        if v <= 4 and self.last["poly"]:
            self.shape("poly", v, self.last["poly"])

    def _rgn(self, rd, op) -> None:
        v, same = self._verb(op, 0x80)
        if not same:
            self.last["rgn"] = _region(rd)
        if v <= 4 and self.last["rgn"] is not None:
            self.shape("rgn", v, self.last["rgn"])


def _hull(points: list) -> list:
    pts = sorted(set(points))
    if len(pts) < 3:
        return pts

    def cross(o, a, b):
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])

    lower: list = []
    for p in pts:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], p) <= 0:
            lower.pop()
        lower.append(p)
    upper: list = []
    for p in reversed(pts):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], p) <= 0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]


def _ellipse(l, t, r, b, P, n: int = 72) -> list:
    import math

    cx, cy, rx, ry = (l + r) / 2, (t + b) / 2, (r - l) / 2, (b - t) / 2
    return [P(cx + rx * math.cos(2 * math.pi * i / n), cy + ry * math.sin(2 * math.pi * i / n)) for i in range(n)]


def _rrect(l, t, r, b, ow, oh, P) -> list:
    import math

    rx, ry = min(ow / 2, (r - l) / 2), min(oh / 2, (b - t) / 2)
    out = []
    for cx, cy, a0 in ((r - rx, t + ry, -90), (r - rx, b - ry, 0), (l + rx, b - ry, 90), (l + rx, t + ry, 180)):
        for i in range(9):
            a = math.radians(a0 + 90 * i / 8)
            out.append(P(cx + rx * math.cos(a), cy + ry * math.sin(a)))
    return out


def render(blob: bytes) -> Frame:
    base, t, l, b, r = _frame_of(blob)
    res = _resolution(blob, base)
    et, el, eb, er = _ext_frame(blob, base, t, l, b, r)
    if eb <= et or er <= el:
        raise ImageError("PICT：絵の枠の大きさが 0 です")
    rd = _Reader(blob, base + 10)
    if blob[base + 10:base + 12] == b"\x11\x01":
        version = 1
        rd.p += 2
    else:
        version = 2
        rd.p += 4
    pict = _Pict(blob, base, et, el, eb, er)
    pict.align = (rd.p - 0) & 1
    pict.run(rd, version)
    pl = pict.pl
    problems = [f"{k}（{v} か所）" for k, v in pl.skipped.items()]
    notes = []
    if pl.text:
        notes.append("絵の中の字：" + "／".join(pl.text[:200]))
    if not pl.drawn:
        problems.append("描ける命令がありません（中身は読めません）")
    return Frame(pict.W, pict.H, bytes(pl.cv.buf), res or (72.0, 72.0), notes, problems)
