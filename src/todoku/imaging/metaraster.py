"""メタファイル（EMF：MS-EMF、WMF：MS-WMF）を画素に描く：線・多角形・楕円・字・ビットマップなど、よく使う記録。

描けない記録は、種類と数を problems で伝える（黙って飛ばさない）。字の記録の字は、描くほかに text で返す（エージェントが字として読める）。
EMF+ だけで描かれた絵は描けないので、中に入っている PNG・JPEG があれば取り出し、そうでなければ、描けないと伝える。
描くのは PDF の描画の部品（pdf/render/canvas.py）。
"""

from __future__ import annotations

import math
import struct

from .formats import Frame, ImageError, dib_frame

DEFAULT_DPI = 150
MAX_SIDE = 3000
_STOCK_BRUSH = {0: (255, 255, 255), 1: (192, 192, 192), 2: (128, 128, 128), 3: (64, 64, 64), 4: (0, 0, 0), 5: None}
_STOCK_PEN = {6: (255, 255, 255), 7: (0, 0, 0), 8: None}


def _rgb(c: int) -> tuple:
    return (c & 255, (c >> 8) & 255, (c >> 16) & 255)


class _State:
    __slots__ = ("world", "wo", "we", "vo", "ve", "mode", "pen", "brush", "font", "tc", "bk_op", "bk", "ta", "fill_eo", "cur", "clip")

    def __init__(self, clip):
        self.world = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        self.wo = (0.0, 0.0)
        self.we = (1.0, 1.0)
        self.vo = (0.0, 0.0)
        self.ve = (1.0, 1.0)
        self.mode = 1
        self.pen = {"rgb": (0, 0, 0), "w": 0.0, "style": 0}
        self.brush = (255, 255, 255)
        self.font = {"h": 12.0, "weight": 400, "italic": False, "esc": 0.0, "face": ""}
        self.tc = (0, 0, 0)
        self.bk_op = True
        self.bk = (255, 255, 255)
        self.ta = 0
        self.fill_eo = True
        self.cur = (0.0, 0.0)
        self.clip = clip

    def copy(self):
        o = _State.__new__(_State)
        for k in _State.__slots__:
            setattr(o, k, getattr(self, k))
        return o


class _Player:
    def __init__(self, W: int, H: int, base):
        from ..pdf.render.canvas import Canvas, Clip

        self.cv = Canvas(W, H)
        self.W, self.H = W, H
        self.base = base  # 装置の座標 → 画素
        self.st = _State(Clip(0, 0, W, H))
        self.stack: list = []
        self.objects: dict = {}
        self.path: list | None = None
        self.subs: list = []
        self.text: list[str] = []
        self.skipped: dict[str, int] = {}
        self.drawn = 0
        self.mm_unit = (1.0, 1.0)  # 1 mm の装置の単位

    # 座標
    def page_scale(self):
        s = self.st
        m = s.mode
        px, py = self.mm_unit
        if m == 1:
            return 1.0, 1.0
        if m in (2, 3, 4, 5, 6):
            mm = {2: 0.1, 3: 0.01, 4: 0.254, 5: 0.0254, 6: 25.4 / 1440}[m]
            return px * mm, -py * mm
        sx = s.ve[0] / s.we[0] if s.we[0] else 1.0
        sy = s.ve[1] / s.we[1] if s.we[1] else 1.0
        return sx, sy

    def P(self, x: float, y: float) -> tuple:
        s = self.st
        a, b, c, d, e, f = s.world
        wx, wy = a * x + c * y + e, b * x + d * y + f
        sx, sy = self.page_scale()
        dx = (wx - s.wo[0]) * sx + s.vo[0]
        dy = (wy - s.wo[1]) * sy + s.vo[1]
        m = self.base
        return (m[0] * dx + m[4], m[3] * dy + m[5])

    def unit_scale(self) -> float:
        s = self.st
        sx, sy = self.page_scale()
        a, b, c, d = s.world[:4]
        return math.sqrt(abs((a * d - b * c) * sx * sy * self.base[0] * self.base[3]))

    # 描く
    def fill(self, polys: list, eo: bool | None = None) -> None:
        from ..pdf.render.canvas import fill_rows, paint

        b = self.st.brush
        if b is None or self.st.clip.empty or not polys:
            return
        rows = fill_rows(polys, self.st.fill_eo if eo is None else eo, self.st.clip.rect)
        paint(self.cv, rows, b, 1.0, self.st.clip)
        self.drawn += 1

    def stroke(self, subs: list) -> None:
        from ..pdf.render.canvas import fill_rows, paint, stroke_polys

        pen = self.st.pen
        if pen["rgb"] is None or self.st.clip.empty or not subs:
            return
        w = max(1.0, pen["w"] * self.unit_scale())
        dash = None
        st = pen["style"] & 15
        if st == 5:
            return
        if st in (1, 2, 3, 4):
            k = max(1.0, w)
            dash = {1: [6 * k, 3 * k], 2: [k, 2 * k], 3: [6 * k, 2 * k, k, 2 * k], 4: [6 * k, 2 * k, k, 2 * k, k, 2 * k]}[st]
        polys = stroke_polys(subs, w, 0, 0, 10.0, dash, 0.0, 1.0)
        if polys:
            paint(self.cv, fill_rows(polys, False, self.st.clip.rect), pen["rgb"], 1.0, self.st.clip)
        self.drawn += 1

    def shape(self, pts_list: list, closed: bool = True, fill: bool = True) -> None:
        """論理の座標の多角形の並びを描く（パスの途中なら、パスに足す）。"""
        dev = [[self.P(x, y) for x, y in pts] for pts in pts_list if pts]
        if self.path is not None:
            self.path += [(p, closed) for p in dev]
            return
        if fill and closed:
            self.fill([p for p in dev if len(p) >= 3])
        self.stroke([(p, closed) for p in dev if len(p) >= 2])

    def box(self, x0, y0, x1, y1):
        return [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]

    def ellipse_pts(self, x0, y0, x1, y1, a0=0.0, a1=2 * math.pi, n=48):
        cx, cy, rx, ry = (x0 + x1) / 2, (y0 + y1) / 2, abs(x1 - x0) / 2, abs(y1 - y0) / 2
        return [(cx + rx * math.cos(a0 + (a1 - a0) * i / n), cy + ry * math.sin(a0 + (a1 - a0) * i / n)) for i in range(n + 1)]

    def bezier(self, p0, pts):
        """3 次のベジエ（p0 から、3 点ずつ）を折れ線に。"""
        out = [p0]
        cur = p0
        for i in range(0, len(pts) - 2, 3):
            c1, c2, e = pts[i:i + 3]
            for k in range(1, 17):
                t = k / 16
                u = 1 - t
                out.append((u ** 3 * cur[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t ** 3 * e[0],
                            u ** 3 * cur[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t ** 3 * e[1]))
            cur = e
        return out

    def draw_text(self, s: str, x: float, y: float, dx: list | None) -> None:
        if not s:
            return
        self.text.append(s)
        from ..pdf.render.canvas import fill_rows, flatten, paint
        from ..pdf.render.fontprog import CFFOutlines
        from ..pdf.render.fontmap import system_face

        st = self.st
        font = st.font
        h = abs(font["h"]) or 12.0
        k = self.unit_scale()
        size = h * k
        if st.ta & 1:  # TA_UPDATECP
            x, y = st.cur
        ang = -math.radians(font["esc"] / 10.0)  # 装置の座標は下向き
        ca, sa = math.cos(ang), math.sin(ang)
        bold = font["weight"] >= 600
        cjk = any(ord(ch) >= 0x2E80 for ch in s)
        kinds = (("cjk-sans-bold" if bold else "cjk-sans"), "fallback") if cjk else (("sans-bold" if bold else "sans"), "fallback")
        faces = [f for f in (system_face(n) for n in kinds) if f is not None]
        ox, oy = self.P(x, y)
        adv: list = []
        items = []
        for i, ch in enumerate(s):
            face = next((f for f in faces if f.lookup(ch) is not None), None)
            g = face.lookup(ch) if face else None
            if dx is not None and i < len(dx):
                a = dx[i] * k
            elif face is not None and g is not None and hasattr(face.outl, "advance") and face.outl.advance(g):
                a = face.outl.advance(g) * size
            else:
                a = size if ord(ch) >= 0x2E80 else size * 0.5
            adv.append(a)
            items.append((face, g, ch))
        total = sum(adv)
        align = st.ta & 6
        px = -total if align == 2 else -total / 2 if align == 6 else 0.0
        v = st.ta & 24
        py = 0.0 if v == 24 else (-0.2 * size if v == 8 else 0.85 * size)  # 基線・下・上（字の高さの 0.85 を基線まで）
        missing = 0
        paint_rgb = st.tc
        if not faces:
            missing = len(s)
        for (face, g, ch), a in zip(items, adv):
            if face is not None and g is not None and ch.strip():
                o = face.outl
                gm = o.glyph_matrix(g) if isinstance(o, CFFOutlines) else o.matrix
                # 字の座標（上向き）→ 画素（下向き）
                M = (gm[0] * size * ca, gm[0] * size * sa, gm[3] * size * sa, -gm[3] * size * ca,
                     ox + ca * px - sa * py, oy + sa * px + ca * py)
                polys = [p for p, _c in flatten(o.glyph(g), M) if len(p) >= 2]
                if polys and not st.clip.empty:
                    paint(self.cv, fill_rows(polys, False, st.clip.rect), paint_rgb, 1.0, st.clip)
            elif face is None and ch.strip():
                missing += 1
            px += a
        if missing:
            self.skip("字の形がこのパソコンの書体にない字（字は結果の text にあります）", missing)
        if st.ta & 1:
            st.cur = (x + total / (k or 1), y)
        self.drawn += 1

    def skip(self, what: str, n: int = 1) -> None:
        self.skipped[what] = self.skipped.get(what, 0) + n

    def image(self, fr: Frame, x: float, y: float, cx: float, cy: float, sx: int = 0, sy: int = 0, scx: int | None = None, scy: int | None = None) -> None:
        from ..pdf.render.image import RGBA, draw_image

        scx = fr.w if scx is None else scx
        scy = fr.h if scy is None else scy
        px = fr.rgb
        w, h = fr.w, fr.h
        if (sx, sy, scx, scy) != (0, 0, w, h) and scx > 0 and scy > 0 and sx >= 0 and sy >= 0 and sx + scx <= w and sy + scy <= h:
            px = b"".join(px[(j * w + sx) * 3:(j * w + sx + scx) * 3] for j in range(sy, sy + scy))
            w, h = scx, scy
        a0 = self.P(x, y)
        a1 = self.P(x + cx, y + cy)
        X0, X1, Y0, Y1 = min(a0[0], a1[0]), max(a0[0], a1[0]), min(a0[1], a1[1]), max(a0[1], a1[1])
        if X1 - X0 < 0.01 or Y1 - Y0 < 0.01:
            return
        draw_image(self.cv, RGBA(w, h, px, None), (X1 - X0, 0.0, 0.0, -(Y1 - Y0), X0, Y1), self.st.clip)
        self.drawn += 1


# ====================================================================== EMF
def _pts16(r: bytes, at: int, n: int) -> list:
    v = struct.unpack_from(f"<{2 * n}h", r, at)
    return list(zip(v[0::2], v[1::2]))


def _pts32(r: bytes, at: int, n: int) -> list:
    v = struct.unpack_from(f"<{2 * n}i", r, at)
    return list(zip(v[0::2], v[1::2]))


def emf(blob: bytes) -> Frame:
    if len(blob) < 88 or struct.unpack_from("<I", blob, 0)[0] != 1 or blob[40:44] != b" EMF":
        raise ImageError("EMF の見出しがありません")
    bx0, by0, bx1, by1, fx0, fy0, fx1, fy1 = struct.unpack_from("<8i", blob, 8)
    devw, devh, mmw, mmh = struct.unpack_from("<4i", blob, 72)
    fw_mm, fh_mm = (fx1 - fx0) / 100.0, (fy1 - fy0) / 100.0
    mm_x = devw / mmw if mmw > 0 else 96 / 25.4
    mm_y = devh / mmh if mmh > 0 else 96 / 25.4
    if fw_mm <= 0 or fh_mm <= 0:
        fw_mm, fh_mm = max(1, bx1 - bx0) / mm_x, max(1, by1 - by0) / mm_y
    W = max(1, int(fw_mm * DEFAULT_DPI / 25.4 + 0.5))
    H = max(1, int(fh_mm * DEFAULT_DPI / 25.4 + 0.5))
    k = min(1.0, MAX_SIDE / max(W, H))
    W, H = max(1, int(W * k)), max(1, int(H * k))
    # 装置の座標 → 画素：Frame（0.01 mm）の左上が画像の左上
    sx, sy = W / (fw_mm * mm_x), H / (fh_mm * mm_y)
    base = (sx, 0.0, 0.0, sy, -fx0 / 100.0 * mm_x * sx, -fy0 / 100.0 * mm_y * sy)
    pl = _Player(W, H, base)
    pl.mm_unit = (mm_x, mm_y)
    dpi = (W / fw_mm * 25.4, H / fh_mm * 25.4)
    pos = 0
    n = len(blob)
    plus = False
    blobs: list[bytes] = []
    while pos + 8 <= n:
        t, size = struct.unpack_from("<II", blob, pos)
        if size < 8 or pos + size > n or size % 4:
            pl.skip("壊れた記録で止めた")
            break
        r = blob[pos:pos + size]
        pos += size
        if t == 14:
            break
        if t == 70:
            if r[12:16] == b"EMF+":
                plus = True
                for sig in (b"\x89PNG\r\n\x1a\n", b"\xff\xd8\xff"):
                    i = r.find(sig, 16)
                    if i >= 0:
                        blobs.append(r[i:])
            continue
        try:
            _emf_record(pl, t, r)
        except (struct.error, IndexError, ValueError, ZeroDivisionError) as e:
            pl.skip(f"読めない記録（種類 {t}：{type(e).__name__}）")
    return _finish(pl, W, H, dpi, plus, blobs)


def _finish(pl: _Player, W: int, H: int, dpi, plus: bool, blobs: list) -> Frame:
    problems = [f"{k}（{v} か所）" for k, v in pl.skipped.items()]
    notes: list = []
    if pl.text:
        notes.append("絵の中の字：" + "／".join(pl.text[:200]))
    if not pl.drawn and blobs:
        from .formats import decode

        try:
            fr = decode(blobs[0])
            fr.notes.append("EMF+ の絵の中の、埋め込まれたビットマップを示しました")
            return fr
        except Exception:  # noqa: BLE001
            pass
    if not pl.drawn and plus:
        problems.append("EMF+ だけで描かれた絵で、描ける記録がありません（中身は読めません）")
    f = Frame(W, H, bytes(pl.cv.buf), dpi, notes, problems)
    return f


def _emf_record(pl: _Player, t: int, r: bytes) -> None:
    st = pl.st
    u32 = lambda p: struct.unpack_from("<I", r, p)[0]  # noqa: E731
    i32 = lambda p: struct.unpack_from("<i", r, p)[0]  # noqa: E731
    if t in (86, 87, 3, 4, 85, 88, 89, 2, 5, 6):
        big = t in (2, 3, 4, 5, 6)
        cnt = u32(24)
        pts = (_pts32 if big else _pts16)(r, 28, cnt)
        if t in (86, 3):
            pl.shape([pts], True)
        elif t in (87, 4):
            pl.shape([pts], False, False)
        elif t in (85, 2):
            pl.shape([pl.bezier(pts[0], pts[1:])], False, False)
        elif t in (88, 5):
            line = pl.bezier(st.cur, pts)
            pl.shape([line], False, False)
            st.cur = pts[-1]
        else:
            pl.shape([[st.cur] + pts], False, False)
            st.cur = pts[-1]
    elif t in (90, 91, 7, 8):
        big = t in (7, 8)
        np_, total = u32(24), u32(28)
        counts = struct.unpack_from(f"<{np_}I", r, 32)
        at = 32 + 4 * np_
        pts = (_pts32 if big else _pts16)(r, at, total)
        out = []
        i = 0
        for c in counts:
            out.append(pts[i:i + c])
            i += c
        pl.shape(out, t in (91, 8), t in (91, 8))
    elif t == 9:
        st.we = (float(i32(8)), float(i32(12)))
    elif t == 10:
        st.wo = (float(i32(8)), float(i32(12)))
    elif t == 11:
        st.ve = (float(i32(8)), float(i32(12)))
    elif t == 12:
        st.vo = (float(i32(8)), float(i32(12)))
    elif t == 17:
        st.mode = u32(8)
    elif t == 18:
        st.bk_op = u32(8) == 2
    elif t == 19:
        st.fill_eo = u32(8) == 1
    elif t == 22:
        st.ta = u32(8)
    elif t == 24:
        st.tc = _rgb(u32(8))
    elif t == 25:
        st.bk = _rgb(u32(8))
    elif t == 27:
        st.cur = (i32(8), i32(12))
        if pl.path is not None:
            pl.subs.append([st.cur])
    elif t == 54:
        x, y = i32(8), i32(12)
        pl.shape([[st.cur, (x, y)]], False, False)
        st.cur = (x, y)
    elif t == 33:
        pl.stack.append(st.copy())
    elif t == 34:
        n = i32(8)
        idx = len(pl.stack) + n if n < 0 else n - 1
        if 0 <= idx < len(pl.stack):
            pl.st = pl.stack[idx]
            del pl.stack[idx:]
    elif t == 35:
        st.world = struct.unpack_from("<6f", r, 8)
    elif t == 36:
        m = struct.unpack_from("<6f", r, 8)
        mode = u32(32)
        from ..pdf.render.canvas import mul

        if mode == 1:
            st.world = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        elif mode == 2:
            st.world = mul(m, st.world)
        elif mode == 3:
            st.world = mul(st.world, m)
        elif mode == 4:
            st.world = m
    elif t == 37:
        h = u32(8)
        if h & 0x80000000:
            k = h & 0x7FFFFFFF
            if k in _STOCK_BRUSH:
                st.brush = _STOCK_BRUSH[k]
            elif k in _STOCK_PEN:
                st.pen = {"rgb": _STOCK_PEN[k], "w": 0.0, "style": 0}
        else:
            o = pl.objects.get(h)
            if o is not None:
                if o[0] == "pen":
                    st.pen = o[1]
                elif o[0] == "brush":
                    st.brush = o[1]
                elif o[0] == "font":
                    st.font = o[1]
    elif t == 38:
        s_ = u32(12)
        pl.objects[u32(8)] = ("pen", {"rgb": None if s_ & 15 == 5 else _rgb(u32(24)), "w": float(i32(16)), "style": s_})
    elif t == 95:
        s_ = u32(28)
        pl.objects[u32(8)] = ("pen", {"rgb": None if s_ & 15 == 5 else _rgb(u32(40)), "w": float(u32(32)), "style": s_})
    elif t == 39:
        s_ = u32(12)
        c = _rgb(u32(16))
        if s_ == 2:
            pl.skip("模様のブラシ（単色で塗った）")
        pl.objects[u32(8)] = ("brush", None if s_ == 1 else c)
    elif t in (93, 94):
        pl.skip("ビットマップのブラシ（灰色で塗った）")
        pl.objects[u32(8)] = ("brush", (128, 128, 128))
    elif t == 82:
        face = r[40:104].decode("utf-16-le", "replace").split("\0")[0]
        pl.objects[u32(8)] = ("font", {"h": float(i32(12)), "weight": i32(28), "italic": bool(r[32]), "esc": float(i32(20)), "face": face})
    elif t == 40:
        pl.objects.pop(u32(8), None)
    elif t == 42:
        x0, y0, x1, y1 = struct.unpack_from("<4i", r, 8)
        pl.shape([pl.ellipse_pts(x0, y0, x1, y1)], True)
    elif t == 43:
        x0, y0, x1, y1 = struct.unpack_from("<4i", r, 8)
        pl.shape([pl.box(x0, y0, x1, y1)], True)
    elif t == 44:
        x0, y0, x1, y1, cw, ch = struct.unpack_from("<6i", r, 8)
        pl.shape([_round_rect(x0, y0, x1, y1, cw, ch)], True)
    elif t in (45, 46, 47, 55):
        x0, y0, x1, y1, sx, sy, ex, ey = struct.unpack_from("<8i", r, 8)
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        rx, ry = max(1e-6, abs(x1 - x0) / 2), max(1e-6, abs(y1 - y0) / 2)
        a0 = math.atan2((sy - cy) / ry, (sx - cx) / rx)
        a1 = math.atan2((ey - cy) / ry, (ex - cx) / rx)
        if a1 >= a0:
            a1 -= 2 * math.pi  # 既定は反時計回り（装置の座標は下向き）
        pts = pl.ellipse_pts(x0, y0, x1, y1, a0, a1)
        if t in (45, 55):
            pl.shape([pts], False, False)
        elif t == 46:
            pl.shape([pts], True)
        else:
            pl.shape([pts + [(cx, cy)]], True)
    elif t == 59:
        pl.path = []
    elif t == 60:
        pass
    elif t == 61:
        if pl.path:
            pts, _c = pl.path[-1]
            pl.path[-1] = (pts, True)
    elif t in (62, 63, 64):
        path = pl.path or []
        pl.path_done = path
        pl.path = None
        if t in (62, 63):
            pl.fill([p for p, _c in path if len(p) >= 3])
        if t in (63, 64):
            pl.stroke(path)
    elif t == 68:
        pl.path = None
    elif t == 30:
        x0, y0, x1, y1 = struct.unpack_from("<4i", r, 8)
        a, b = pl.P(x0, y0), pl.P(x1, y1)
        st.clip = st.clip.with_rect(min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1]))
    elif t in (29, 26, 67, 75, 28):
        if t not in (75, 28):
            pl.skip("切り抜きの一部（切り抜かずに描いた）")
    elif t in (83, 84):
        _emf_text(pl, r, t == 84)
    elif t == 108:
        x, y, cch, fu = i32(8), i32(12), u32(16), u32(20)
        at = 36 + (16 if not fu & 0x100 else 0)
        s = r[at:at + cch].decode("latin-1") if fu & 0x200 else r[at:at + 2 * cch].decode("utf-16-le", "replace")
        pl.draw_text(s, x, y, None)
    elif t in (76, 77, 114, 116):
        _emf_blt(pl, t, r)
    elif t == 81:
        x, y, sx, sy, cxs, cys, ob, _cb, of, _cf = struct.unpack_from("<4i2i4I", r, 24)
        cxd, cyd = i32(72), i32(76)
        if ob and of:
            fr = dib_frame(r, ob, None, of)
            pl.image(fr, x, y, cxd, cyd, sx, sy, cxs, cys)
    elif t == 80:
        x, y, sx, sy, cxs, cys, ob, _cb, of, _cf = struct.unpack_from("<4i2i4I", r, 24)
        if ob and of:
            fr = dib_frame(r, ob, None, of)
            pl.image(fr, x, y, cxs, cys, sx, fr.h - sy - cys if fr.h >= sy + cys else 0, cxs, cys)
    elif t in (13, 15, 16, 20, 21, 23, 48, 49, 50, 51, 52, 57, 58, 65, 66, 1, 98, 106, 115):
        pass
    else:
        pl.skip(f"描けない記録（EMF の種類 {t}）")


def _round_rect(x0, y0, x1, y1, cw, ch) -> list:
    rx, ry = abs(cw) / 2, abs(ch) / 2
    out = []
    for cx, cy, a in ((x1 - rx, y0 + ry, -90), (x1 - rx, y1 - ry, 0), (x0 + rx, y1 - ry, 90), (x0 + rx, y0 + ry, 180)):
        for i in range(9):
            t = math.radians(a + i * 90 / 8)
            out.append((cx + rx * math.cos(t), cy + ry * math.sin(t)))
    return out


def _emf_text(pl: _Player, r: bytes, wide: bool) -> None:
    x, y, n, off, _opt = struct.unpack_from("<iiIII", r, 36)
    odx = struct.unpack_from("<I", r, 72)[0]
    s = r[off:off + 2 * n].decode("utf-16-le", "replace") if wide else r[off:off + n].decode("cp932", "replace")
    dx = list(struct.unpack_from(f"<{n}i", r, odx)) if odx and odx + 4 * n <= len(r) else None
    pl.draw_text(s, x, y, dx)


def _emf_blt(pl: _Player, t: int, r: bytes) -> None:
    x, y, cx, cy = struct.unpack_from("<4i", r, 24)
    rop = struct.unpack_from("<I", r, 40)[0]
    sx, sy = struct.unpack_from("<2i", r, 44)
    ob, cb, of, _cf = struct.unpack_from("<4I", r, 84)
    if not cb or not of:
        c = {0x00FF0062: (255, 255, 255), 0x00000042: (0, 0, 0)}.get(rop)
        if c is None and pl.st.brush is not None:
            c = pl.st.brush
        if c is None:
            return
        keep = pl.st.brush
        pl.st.brush = c
        p = pl.P(x, y)
        q = pl.P(x + cx, y + cy)
        pl.fill([[p, (q[0], p[1]), q, (p[0], q[1])]], False)
        pl.st.brush = keep
        return
    fr = dib_frame(r, ob, None, of)
    if t == 76:
        scx, scy = cx, cy
    else:
        scx, scy = struct.unpack_from("<2i", r, 100)
    if rop not in (0x00CC0020, 0x00000000) and t in (76, 77):
        pl.skip("重ね方（ROP）のあるビットマップ（そのまま描いた）")
    pl.image(fr, x, y, cx, cy, sx, sy, scx, scy)


# ====================================================================== WMF
def wmf(blob: bytes) -> Frame:
    pos = 0
    inch = 0
    box = None
    if blob[:4] == b"\xd7\xcd\xc6\x9a":
        l_, t_, r_, b_, inch = struct.unpack_from("<4hH", blob, 6)
        box = (l_, t_, r_, b_)
        pos = 22
    if pos + 18 > len(blob):
        raise ImageError("WMF の見出しがありません")
    hsize = struct.unpack_from("<H", blob, pos + 2)[0]
    if hsize != 9:
        raise ImageError("WMF の見出しが規格に合いません")
    # 大きさ：置ける見出しがあれば、なければ窓の大きさの記録を先に探す
    win = _wmf_first_window(blob, pos + 18)
    if box is not None and inch:
        w_mm, h_mm = abs(box[2] - box[0]) / inch * 25.4, abs(box[3] - box[1]) / inch * 25.4
        org = (min(box[0], box[2]), min(box[1], box[3]))
        ext = (abs(box[2] - box[0]) or 1, abs(box[3] - box[1]) or 1)
    elif win:
        org, ext = win
        w_mm, h_mm = abs(ext[0]) / 96 * 25.4, abs(ext[1]) / 96 * 25.4
    else:
        raise ImageError("WMF の大きさが分かりません（置ける見出しも、窓の記録もありません）")
    W = max(1, int(w_mm * DEFAULT_DPI / 25.4 + 0.5))
    H = max(1, int(h_mm * DEFAULT_DPI / 25.4 + 0.5))
    k = min(1.0, MAX_SIDE / max(W, H))
    W, H = max(1, int(W * k)), max(1, int(H * k))
    pl = _Player(W, H, (1.0, 0.0, 0.0, 1.0, 0.0, 0.0))
    # WMF の装置の座標は、論理の座標そのまま（窓 → 画像）。ANISOTROPIC と同じ計算にする
    st = pl.st
    st.mode = 8
    st.wo = (float(org[0]), float(org[1]))
    st.we = (float(ext[0] or 1), float(ext[1] or 1))
    st.vo = (0.0, 0.0)
    st.ve = (float(W), float(H))
    placeable = box is not None
    dpi = (W / w_mm * 25.4, H / h_mm * 25.4) if w_mm and h_mm else None
    p = pos + 18
    n = len(blob)
    slots: list = []
    while p + 6 <= n:
        size, func = struct.unpack_from("<IH", blob, p)
        if size < 3 or p + size * 2 > n:
            pl.skip("壊れた記録で止めた")
            break
        r = blob[p + 6:p + size * 2]
        p += size * 2
        if func == 0:
            break
        try:
            _wmf_record(pl, func, r, slots, placeable, org, ext, W, H)
        except (struct.error, IndexError, ValueError, ZeroDivisionError) as e:
            pl.skip(f"読めない記録（種類 {func:#x}：{type(e).__name__}）")
    return _finish(pl, W, H, dpi, False, [])


def _wmf_first_window(blob: bytes, p: int):
    org, ext = None, None
    n = len(blob)
    while p + 6 <= n:
        size, func = struct.unpack_from("<IH", blob, p)
        if size < 3 or func == 0:
            break
        if func == 0x020B:
            y, x = struct.unpack_from("<2h", blob, p + 6)
            org = org or (x, y)
        elif func == 0x020C:
            y, x = struct.unpack_from("<2h", blob, p + 6)
            ext = ext or (x, y)
        if org and ext:
            break
        p += size * 2
    return (org or (0, 0), ext) if ext else None


def _slot(slots: list, obj) -> None:
    for i, v in enumerate(slots):
        if v is None:
            slots[i] = obj
            return
    slots.append(obj)


def _wmf_record(pl: _Player, f: int, r: bytes, slots: list, placeable: bool, org, ext, W, H) -> None:
    st = pl.st
    h16 = lambda p: struct.unpack_from("<h", r, p)[0]  # noqa: E731
    u16 = lambda p: struct.unpack_from("<H", r, p)[0]  # noqa: E731
    if f == 0x020B:
        if not placeable:
            st.wo = (float(h16(2)), float(h16(0)))
    elif f == 0x020C:
        if not placeable:
            st.we = (float(h16(2)) or 1.0, float(h16(0)) or 1.0)
    elif f in (0x020D, 0x020E, 0x0103, 0x0104, 0x0102 + 0x1000, 0x0107, 0x0231, 0x0037, 0x0035, 0x0234):
        pass
    elif f == 0x0102:
        st.bk_op = u16(0) == 2
    elif f == 0x0106:
        st.fill_eo = u16(0) == 1
    elif f == 0x012E:
        st.ta = u16(0)
    elif f == 0x0209:
        st.tc = _rgb(struct.unpack_from("<I", r, 0)[0])
    elif f == 0x0201:
        st.bk = _rgb(struct.unpack_from("<I", r, 0)[0])
    elif f == 0x001E:
        pl.stack.append(st.copy())
    elif f == 0x0127:
        if pl.stack:
            n = h16(0)
            idx = len(pl.stack) + n if n < 0 else n - 1
            if 0 <= idx < len(pl.stack):
                pl.st = pl.stack[idx]
                del pl.stack[idx:]
    elif f == 0x0214:
        st.cur = (h16(2), h16(0))
    elif f == 0x0213:
        x, y = h16(2), h16(0)
        pl.shape([[st.cur, (x, y)]], False, False)
        st.cur = (x, y)
    elif f == 0x041B:
        b, rr, t, l = struct.unpack_from("<4h", r, 0)
        pl.shape([pl.box(l, t, rr, b)], True)
    elif f == 0x0418:
        b, rr, t, l = struct.unpack_from("<4h", r, 0)
        pl.shape([pl.ellipse_pts(l, t, rr, b)], True)
    elif f == 0x061C:
        ch, cw, b, rr, t, l = struct.unpack_from("<6h", r, 0)
        pl.shape([_round_rect(l, t, rr, b, cw, ch)], True)
    elif f == 0x0324:
        n = u16(0)
        pl.shape([_pts16(r, 2, n)], True)
    elif f == 0x0325:
        n = u16(0)
        pl.shape([_pts16(r, 2, n)], False, False)
    elif f == 0x0538:
        np_ = u16(0)
        counts = struct.unpack_from(f"<{np_}H", r, 2)
        at = 2 + 2 * np_
        out = []
        for c in counts:
            out.append(_pts16(r, at, c))
            at += 4 * c
        pl.shape(out, True)
    elif f == 0x02FA:
        style, wx, _wy, color = struct.unpack_from("<HhhI", r, 0)
        _slot(slots, ("pen", {"rgb": None if style & 15 == 5 else _rgb(color), "w": float(wx), "style": style}))
    elif f == 0x02FC:
        style, color = struct.unpack_from("<HI", r, 0)
        _slot(slots, ("brush", None if style == 1 else _rgb(color)))
    elif f == 0x02FB:
        h_, _w, esc, _ori, weight = struct.unpack_from("<5h", r, 0)
        _slot(slots, ("font", {"h": float(h_), "weight": weight, "italic": bool(r[10]), "esc": float(esc), "face": ""}))
    elif f in (0x00F7, 0x01F9, 0x0142, 0x06FF):
        if f in (0x01F9, 0x0142):
            pl.skip("ビットマップのブラシ（灰色で塗った）")
            _slot(slots, ("brush", (128, 128, 128)))
        else:
            _slot(slots, ("other",))
    elif f == 0x012D:
        i = u16(0)
        o = slots[i] if i < len(slots) else None
        if o:
            if o[0] == "pen":
                st.pen = o[1]
            elif o[0] == "brush":
                st.brush = o[1]
            elif o[0] == "font":
                st.font = o[1]
    elif f == 0x01F0:
        i = u16(0)
        if i < len(slots):
            slots[i] = None
    elif f == 0x0521:
        n = u16(0)
        at = 2 + n + (n & 1)
        s = r[2:2 + n].decode("cp932", "replace")
        y, x = struct.unpack_from("<2h", r, at)
        pl.draw_text(s, x, y, None)
    elif f == 0x0A32:
        y, x, n, opt = struct.unpack_from("<hhhH", r, 0)
        at = 8 + (8 if opt & 6 else 0)
        s = r[at:at + n].decode("cp932", "replace")
        at2 = at + n + (n & 1)
        dx = list(struct.unpack_from(f"<{n}h", r, at2)) if at2 + 2 * n <= len(r) else None
        pl.draw_text(s, x, y, dx)
    elif f == 0x0F43:  # STRETCHDIB
        _rop, _usage, cys, cxs, sy, sx, cyd, cxd, y, x = struct.unpack_from("<IH8h", r, 0)
        fr = dib_frame(r, 22, None)
        pl.image(fr, x, y, cxd, cyd, sx, sy, cxs, cys)
    elif f == 0x0B41:  # DIBSTRETCHBLT
        _rop, cys, cxs, sy, sx, cyd, cxd, y, x = struct.unpack_from("<I8h", r, 0)
        fr = dib_frame(r, 20, None)
        pl.image(fr, x, y, cxd, cyd, sx, sy, cxs, cys)
    elif f == 0x0940:  # DIBBITBLT
        _rop, sy, sx, cy, cx, y, x = struct.unpack_from("<I6h", r, 0)
        fr = dib_frame(r, 16, None)
        pl.image(fr, x, y, cx, cy, sx, sy, cx, cy)
    elif f == 0x0D33:  # SETDIBTODEV
        _usage, _sc, _ss, sy, sx, cy, cx, y, x = struct.unpack_from("<H8h", r, 0)
        fr = dib_frame(r, 18, None)
        pl.image(fr, x, y, cx, cy, sx, sy, cx, cy)
    else:
        pl.skip(f"描けない記録（WMF の種類 {f:#06x}）")
