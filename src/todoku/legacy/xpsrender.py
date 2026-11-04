"""XPS のページを、エージェントが見る画像にする（view_page。ECMA-388 の FixedPage を、PDF と同じ描画の部品で描く）。

- 座標：FixedPage の 1 単位は 1/96 インチ。Canvas・Path・Glyphs の RenderTransform は親に重ねる。Clip は切り抜き
- Path：Data（省略の記法 M L H V C Q S A Z F0/F1、または PathGeometry・PathFigure の要素）。塗りは SolidColorBrush・ImageBrush。線は Stroke・StrokeThickness
- Glyphs：Indices（字形の番号・送り（1/100 em）・ずれ）と UnicodeString。FontUri の OpenType の字の形で描く。
  .odttf（難読化した書体）は、部品の名前の GUID のバイトを逆順にして先頭 32 バイトに XOR して戻す（ECMA-388 の Font Embedding）
- 画像：ImageBrush の ImageSource（PNG・JPEG・TIFF。imaging）。Viewbox → Viewport で写す
- 近い形で描くもの（結果の not_drawn に書く）：勾配は代表の色、VisualBrush は描かない、OpacityMask は無視、タイル張りは 1 枚目だけ、JPEG XR は灰色の枠

資源（FixedPage.Resources）の {StaticResource 鍵} は、ページと Canvas の辞書を引く。
"""

from __future__ import annotations

import base64
import math
import re
import struct
import time
import xml.etree.ElementTree as ET

from ..pdf.render.canvas import Canvas, Clip, fill_rows, flatten, mul, paint, scan, stroke_polys, transform_polys
from .xps import XpsError, _ln, _read_part, _resolve, pages

MAX_SIDE = 4000
MAX_PIXELS = 12_000_000
UNIT_MM = 25.4 / 96
_NUM = re.compile(r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")
_TOKEN = re.compile(r"[A-Za-z]|[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?")


def _floats(s: str) -> list[float]:
    return [float(x) for x in _NUM.findall(s or "")]


def _color(s: str | None):
    """(r, g, b, a)。#AARRGGBB・#RRGGBB・sc#（0〜1）。分からない色は None。"""
    if not s:
        return None
    s = s.strip()
    if s.startswith("#"):
        h = s[1:]
        try:
            if len(h) == 8:
                return int(h[2:4], 16), int(h[4:6], 16), int(h[6:8], 16), int(h[:2], 16) / 255
            if len(h) == 6:
                return int(h[:2], 16), int(h[2:4], 16), int(h[4:6], 16), 1.0
        except ValueError:
            return None
        return None
    if s.startswith("sc#"):
        v = _floats(s[3:])
        if len(v) == 4:
            return _lin(v[1]), _lin(v[2]), _lin(v[3]), v[0]
        if len(v) == 3:
            return _lin(v[0]), _lin(v[1]), _lin(v[2]), 1.0
    return None


def _lin(v: float) -> int:
    """scRGB（線形）→ sRGB の 0〜255。"""
    v = max(0.0, min(1.0, v))
    return round(255 * (12.92 * v if v <= 0.0031308 else 1.055 * v ** (1 / 2.4) - 0.055))


# ====================================================================== 図形
def _arc(x0, y0, rx, ry, phi, large, sweep, x1, y1) -> list:
    """楕円の弧（端点の形）を 3 次のベジェに（SVG の実装の付録の式）。"""
    if rx == 0 or ry == 0:
        return [(1, x1, y1)]
    rx, ry = abs(rx), abs(ry)
    co, si = math.cos(phi), math.sin(phi)
    dx, dy = (x0 - x1) / 2, (y0 - y1) / 2
    xp, yp = co * dx + si * dy, -si * dx + co * dy
    lam = (xp / rx) ** 2 + (yp / ry) ** 2
    if lam > 1:
        k = math.sqrt(lam)
        rx *= k
        ry *= k
    num = rx * rx * ry * ry - rx * rx * yp * yp - ry * ry * xp * xp
    den = rx * rx * yp * yp + ry * ry * xp * xp
    q = math.sqrt(max(0.0, num / den)) if den else 0.0
    if large == sweep:
        q = -q
    cxp, cyp = q * rx * yp / ry, -q * ry * xp / rx
    cx, cy = co * cxp - si * cyp + (x0 + x1) / 2, si * cxp + co * cyp + (y0 + y1) / 2
    a0 = math.atan2((yp - cyp) / ry, (xp - cxp) / rx)
    a1 = math.atan2((-yp - cyp) / ry, (-xp - cxp) / rx)
    d = a1 - a0
    if sweep and d < 0:
        d += 2 * math.pi
    elif not sweep and d > 0:
        d -= 2 * math.pi
    n = max(1, math.ceil(abs(d) / (math.pi / 2) - 1e-9))
    out = []
    step = d / n
    t = 4 / 3 * math.tan(step / 4)

    def pt(a, ox, oy):
        px, py = rx * (math.cos(a) + ox), ry * (math.sin(a) + oy)
        return co * px - si * py + cx, si * px + co * py + cy

    a = a0
    for _ in range(n):
        b = a + step
        p1 = pt(a, -t * math.sin(a), t * math.cos(a))
        p2 = pt(b, t * math.sin(b), -t * math.cos(b))
        p3 = pt(b, 0, 0)
        out.append((2, p1[0], p1[1], p2[0], p2[1], p3[0], p3[1]))
        a = b
    return out


def parse_data(s: str) -> tuple[list, bool]:
    """省略の記法 → (図形の命令, 偶奇の規則か)。F0 が偶奇（既定）、F1 が非ゼロ。"""
    toks = _TOKEN.findall(s or "")
    i = 0
    ops: list = []
    eo = True
    cx = cy = sx = sy = 0.0
    prev = None
    lc = (0.0, 0.0)
    cmd = ""

    def nums(n):
        nonlocal i
        v = [float(t) for t in toks[i:i + n]]
        if len(v) != n:
            raise ValueError
        i += n
        return v

    try:
        while i < len(toks):
            t = toks[i]
            if t.isalpha():
                cmd = t
                i += 1
                if cmd in "Zz":
                    ops.append((4,))
                    cx, cy = sx, sy
                    prev = cmd
                    continue
                if cmd in "Ff":
                    eo = nums(1)[0] == 0
                    continue
            if not cmd or cmd in "Zz":
                break
            c = cmd.upper()
            rel = cmd.islower()
            ox, oy = (cx, cy) if rel else (0.0, 0.0)
            if c == "M":
                x, y = nums(2)
                cx, cy = x + ox, y + oy
                sx, sy = cx, cy
                ops.append((0, cx, cy))
                cmd = "l" if rel else "L"
            elif c == "L":
                x, y = nums(2)
                cx, cy = x + ox, y + oy
                ops.append((1, cx, cy))
            elif c == "H":
                cx = nums(1)[0] + ox
                ops.append((1, cx, cy))
            elif c == "V":
                cy = nums(1)[0] + oy
                ops.append((1, cx, cy))
            elif c == "C":
                a, b, d, e, x, y = nums(6)
                lc = (d + ox, e + oy)
                ops.append((2, a + ox, b + oy, d + ox, e + oy, x + ox, y + oy))
                cx, cy = x + ox, y + oy
            elif c == "S":
                d, e, x, y = nums(4)
                a, b = (2 * cx - lc[0], 2 * cy - lc[1]) if prev and prev in "CcSs" else (cx, cy)
                lc = (d + ox, e + oy)
                ops.append((2, a, b, d + ox, e + oy, x + ox, y + oy))
                cx, cy = x + ox, y + oy
            elif c == "Q":
                a, b, x, y = nums(4)
                ops.append((3, a + ox, b + oy, x + ox, y + oy))
                cx, cy = x + ox, y + oy
            elif c == "A":
                rx, ry, rot, lg, sw, x, y = nums(7)
                ops += _arc(cx, cy, rx, ry, math.radians(rot), bool(lg), bool(sw), x + ox, y + oy)
                cx, cy = x + ox, y + oy
            else:
                break
            prev = cmd
    except (ValueError, IndexError):
        pass
    return ops, eo


def _pairs(s: str) -> list[tuple[float, float]]:
    v = _floats(s)
    return list(zip(v[0::2], v[1::2]))


def _geometry(el: ET.Element, scope: list) -> tuple[list, bool, tuple | None]:
    """PathGeometry → (命令, 偶奇か, Transform)。"""
    fig_src = el.get("Figures")
    ops: list = []
    eo = (el.get("FillRule") or "EvenOdd") != "NonZero"
    if fig_src:
        o2, _e = parse_data(fig_src)
        ops += o2
    m = _matrix_attr(el.get("Transform"), scope)
    for k in el:
        n = _ln(k.tag)
        if n == "PathGeometry.Transform":
            m = _child_matrix(k, scope) or m
        if n != "PathFigure":
            continue
        start = _pairs(k.get("StartPoint") or "0,0")
        if not start:
            continue
        cx, cy = start[0]
        ops.append((0, cx, cy))
        for seg in k:
            sn = _ln(seg.tag)
            if sn == "PolyLineSegment":
                for x, y in _pairs(seg.get("Points") or ""):
                    ops.append((1, x, y))
                    cx, cy = x, y
            elif sn == "PolyBezierSegment":
                p = _pairs(seg.get("Points") or "")
                for j in range(0, len(p) - 2, 3):
                    ops.append((2, p[j][0], p[j][1], p[j + 1][0], p[j + 1][1], p[j + 2][0], p[j + 2][1]))
                    cx, cy = p[j + 2]
            elif sn == "PolyQuadraticBezierSegment":
                p = _pairs(seg.get("Points") or "")
                for j in range(0, len(p) - 1, 2):
                    ops.append((3, p[j][0], p[j][1], p[j + 1][0], p[j + 1][1]))
                    cx, cy = p[j + 1]
            elif sn == "ArcSegment":
                pt = _pairs(seg.get("Point") or "")
                sz = _pairs(seg.get("Size") or "")
                if pt and sz:
                    ops += _arc(cx, cy, sz[0][0], sz[0][1], math.radians(float(seg.get("RotationAngle") or 0)),
                                (seg.get("IsLargeArc") or "").lower() == "true", (seg.get("SweepDirection") or "").lower() == "clockwise", pt[0][0], pt[0][1])
                    cx, cy = pt[0]
        if (k.get("IsClosed") or "").lower() == "true":
            ops.append((4,))
    return ops, eo, m


def _lookup(key: str, scope: list):
    for d in reversed(scope):
        if key in d:
            return d[key]
    return None


def _ref_key(s: str | None) -> str | None:
    if s and s.startswith("{"):
        m = re.match(r"\{\s*StaticResource\s+([^\s}]+)\s*\}", s)
        return m[1] if m else ""
    return None


def _matrix_attr(s: str | None, scope: list) -> tuple | None:
    if not s:
        return None
    k = _ref_key(s)
    if k is not None:
        el = _lookup(k, scope)
        if el is not None and _ln(el.tag) == "MatrixTransform":
            return _matrix_attr(el.get("Matrix"), scope)
        return None
    v = _floats(s)
    return tuple(v) if len(v) == 6 else None


def _child_matrix(prop: ET.Element, scope: list) -> tuple | None:
    for x in prop:
        if _ln(x.tag) == "MatrixTransform":
            return _matrix_attr(x.get("Matrix"), scope)
    return None


def _prop(el: ET.Element, name: str) -> ET.Element | None:
    tag = f"{_ln(el.tag)}.{name}"
    for k in el:
        if _ln(k.tag) == tag:
            return k
    return None


def _transform(el: ET.Element, scope: list) -> tuple | None:
    m = _matrix_attr(el.get("RenderTransform"), scope)
    p = _prop(el, "RenderTransform")
    if p is not None:
        m = _child_matrix(p, scope) or m
    return m


def _element_geometry(el: ET.Element, attr: str, scope: list) -> tuple[list, bool] | None:
    """Data・Clip の値（省略の記法か、資源の鍵か、子の PathGeometry）。"""
    s = el.get(attr)
    if s:
        k = _ref_key(s)
        if k is None:
            return parse_data(s)
        g = _lookup(k, scope)
        if g is None:
            return None
        ops, eo, m = _geometry(g, scope)
        return (_apply_ops(ops, m), eo)
    p = _prop(el, attr)
    if p is not None:
        for g in p:
            if _ln(g.tag) == "PathGeometry":
                ops, eo, m = _geometry(g, scope)
                return (_apply_ops(ops, m), eo)
    return None


def _apply_ops(ops: list, m) -> list:
    if not m:
        return ops
    a, b, c, d, e, f = m
    out = []
    for o in ops:
        if o[0] == 4:
            out.append(o)
            continue
        q = [o[0]]
        for i in range(1, len(o), 2):
            x, y = o[i], o[i + 1]
            q += [a * x + c * y + e, b * x + d * y + f]
        out.append(tuple(q))
    return out


def _deobfuscate(part: str, data: bytes) -> bytes:
    """難読化した書体（ECMA-388）：部品の名前の GUID の 32 桁を読み、バイトを逆順にして、先頭の 32 バイトに 2 回 XOR する。"""
    hexs = re.sub(r"[^0-9a-fA-F]", "", part.rsplit("/", 1)[-1].split(".")[0])
    if len(hexs) != 32 or len(data) < 32:
        raise XpsError(f"難読化した書体の名前（{part}）に GUID がありません")
    key = bytes.fromhex(hexs)[::-1]
    head = bytearray(data[:32])
    for i in range(16):
        head[i] ^= key[i]
        head[i + 16] ^= key[i]
    return bytes(head) + data[32:]


class _Font:
    def __init__(self, data: bytes, index: int = 0):
        from ..pdf.render.fontprog import CFFOutlines, TrueTypeOutlines, _cmaps, sfnt_tables

        off = 0
        if data[:4] == b"ttcf":
            n = struct.unpack_from(">I", data, 8)[0]
            if not 0 <= index < n:
                index = 0
            off = struct.unpack_from(">I", data, 12 + 4 * index)[0]
        tabs = sfnt_tables(data, off)
        self.tabs = tabs
        self.cff = None
        if "CFF " in tabs:
            self.prog = CFFOutlines(tabs["CFF "])
            self.cff = True
        else:
            self.prog = TrueTypeOutlines(tabs)
        head = tabs.get("head", b"")
        self.upem = struct.unpack_from(">H", head, 18)[0] if len(head) >= 20 else 1000
        self.cmaps = _cmaps(tabs.get("cmap", b""))
        hhea, hm = tabs.get("hhea"), tabs.get("hmtx", b"")
        self.adv: list[int] = []
        if hhea and len(hhea) >= 36:
            nh = min(struct.unpack_from(">H", hhea, 34)[0], len(hm) // 4)
            self.adv = [struct.unpack_from(">H", hm, 4 * i)[0] for i in range(nh)]

    def gid(self, cp: int) -> int:
        for key in ((3, 10), (3, 1), (0, 4), (0, 3), (0, 6), (0, 1), (0, 0), (1, 0)):
            g = self.cmaps.get(key, {}).get(cp)
            if g:
                return g
        sym = self.cmaps.get((3, 0))
        if sym:
            return sym.get(cp) or sym.get(0xF000 + (cp & 255)) or 0
        return 0

    def advance(self, gid: int) -> float:
        if not self.adv:
            return 0.5
        return self.adv[min(gid, len(self.adv) - 1)] / (self.upem or 1000)

    def outline(self, gid: int) -> tuple[list, tuple]:
        p = self.prog
        if self.cff:
            return p.glyph(gid), p.glyph_matrix(gid)
        return p.glyph(gid), p.matrix


_ENTRY = re.compile(r"^(?:\((\d+)(?::(\d+))?\))?([^,]*)(?:,([^,]*))?(?:,([^,]*))?(?:,([^,]*))?(?:,([^,]*))?$")


def _glyph_run(indices: str | None, text: str, font: _Font, em: float, ltr: bool) -> list[tuple[int, float, float]]:
    """(字形の番号, x, y) の並び。Indices の無い所は、字から cmap で引き、送りは書体の hmtx。"""
    chars = [ord(c) for c in text]
    if text.startswith("{}"):
        chars = chars[2:]
    out = []
    x = 0.0
    cp = 0  # 字の位置
    sign = 1.0 if ltr else -1.0
    entries = (indices or "").split(";") if indices else []
    i = 0
    while i < len(entries) or cp < len(chars):
        ent = entries[i] if i < len(entries) else ""
        i += 1
        m = _ENTRY.match(ent.strip())
        cc, gc, gs, adv, uo, vo = (m[1], m[2], m[3], m[4], m[5], m[6]) if m else (None, None, "", None, None, None)
        n_chars = int(cc) if cc else 1
        n_glyphs = int(gc) if gc else 1
        run = 0
        while run < n_glyphs:
            if run:
                ent2 = entries[i] if i < len(entries) else ""
                i += 1
                m2 = _ENTRY.match(ent2.strip())
                gs, adv, uo, vo = (m2[3], m2[4], m2[5], m2[6]) if m2 else ("", None, None, None)
            try:
                gid = int(gs) if gs and gs.strip() else None
            except ValueError:
                gid = None
            if gid is None:
                gid = font.gid(chars[cp]) if cp < len(chars) else 0
            try:
                a = float(adv) * em / 100 if adv and adv.strip() else font.advance(gid) * em
            except ValueError:
                a = font.advance(gid) * em
            try:
                dx = float(uo) * em / 100 if uo and uo.strip() else 0.0
                dy = float(vo) * em / 100 if vo and vo.strip() else 0.0
            except ValueError:
                dx = dy = 0.0
            out.append((gid, x + sign * dx, -dy))
            x += sign * a
            run += 1
        cp += n_chars
    return out


class _Painter:
    def __init__(self, z, names, part: str, cv: Canvas, base, deadline: float):
        self.z, self.names, self.part = z, names, part
        self.cv = cv
        self.base = base
        self.deadline = deadline
        self.fonts: dict[str, _Font | None] = {}
        self.images: dict[str, object] = {}
        self.not_drawn: list[str] = []
        self.stopped = False

    def note(self, s: str) -> None:
        if s not in self.not_drawn:
            self.not_drawn.append(s)

    # ---- 資源
    def resources(self, el: ET.Element, scope: list) -> list:
        p = _prop(el, "Resources")
        if p is None:
            return scope
        d: dict[str, ET.Element] = {}
        for rd in p:
            if _ln(rd.tag) != "ResourceDictionary":
                continue
            src = rd.get("Source")
            if src:
                try:
                    rd2 = ET.fromstring(_read_part(self.z, self.names, _resolve(self.part, src)))
                except (ET.ParseError, XpsError, KeyError):
                    self.note("資源の部品を読めません")
                    continue
                items = list(rd2)
            else:
                items = list(rd)
            for it in items:
                key = next((v for k, v in it.attrib.items() if k.rsplit("}", 1)[-1] == "Key"), None)
                if key:
                    d[key] = it
        return scope + [d] if d else scope

    def font(self, uri: str) -> _Font | None:
        if uri in self.fonts:
            return self.fonts[uri]
        path, _, frag = uri.partition("#")
        try:
            part = _resolve(self.part, path)
            data = _read_part(self.z, self.names, part)
            if part.lower().endswith(".odttf"):
                data = _deobfuscate(part, data)
            f: _Font | None = _Font(data, int(frag) if frag.isdigit() else 0)
        except (XpsError, KeyError, ValueError, struct.error, IndexError) as e:
            self.note(f"書体 {uri} を読めません（{e}）")
            f = None
        self.fonts[uri] = f
        return f

    # ---- 描く
    def clip_of(self, el: ET.Element, ctm, clip: Clip, scope: list) -> Clip:
        g = _element_geometry(el, "Clip", scope)
        if g is None:
            return clip
        ops, eo = g
        polys = [p for p, _c in flatten(ops, ctm) if len(p) >= 2]
        return clip.with_rows(scan(polys, eo, clip.rect)) if polys else Clip(0, 0, 0, 0, {})

    def brush_of(self, el: ET.Element, name: str, scope: list):
        """色の文字か、ブラシの要素（なければ None）。"""
        s = el.get(name)
        if s:
            k = _ref_key(s)
            return _lookup(k, scope) if k is not None else s
        p = _prop(el, name)
        if p is not None:
            for b in p:
                return b
        return None

    def solid(self, b):
        """ブラシ → (r, g, b, a)。勾配は最初の色。描けないものは None。"""
        if b is None:
            return None
        if isinstance(b, str):
            c = _color(b)
            if c is None:
                self.note(f"色 {b} を読めません")
            return c
        n = _ln(b.tag)
        if n == "SolidColorBrush":
            c = _color(b.get("Color"))
            if c is None:
                return None
            return c[0], c[1], c[2], c[3] * float(b.get("Opacity") or 1)
        if n in ("LinearGradientBrush", "RadialGradientBrush"):
            stops = [s for s in b.iter() if _ln(s.tag) == "GradientStop"]
            if stops:
                cs = [_color(s.get("Color")) for s in stops]
                cs = [c for c in cs if c]
                if cs:
                    self.note("勾配は代表の色で描きました")
                    return tuple(sum(c[i] for c in cs) // len(cs) for i in range(3)) + (float(b.get("Opacity") or 1),)
            return None
        if n == "VisualBrush":
            self.note("VisualBrush の中身は描いていません")
        return None

    def draw_image_brush(self, b: ET.Element, polys, eo, ctm, clip: Clip, alpha: float, scope: list) -> None:
        from ..imaging.formats import decode
        from ..pdf.render.image import RGBA, draw_image

        src = b.get("ImageSource") or ""
        path = src.split("{")[0].strip()
        vb = _floats(b.get("Viewbox") or "")
        vp = _floats(b.get("Viewport") or "")
        if len(vp) != 4 or vp[2] == 0 or vp[3] == 0:
            return
        c2 = clip.with_rows(scan(polys, eo, clip.rect))
        if c2.empty:
            return
        m = _matrix_attr(b.get("Transform"), scope)
        tp = _prop(b, "Transform")
        if tp is not None:
            m = _child_matrix(tp, scope) or m
        mt = mul(m, ctm) if m else ctm
        try:
            part = _resolve(self.part, path)
            hit = self.images.get(part)
            if hit is None:
                data = _read_part(self.z, self.names, part)
                fr = decode(data)
                hit = RGBA(fr.w, fr.h, fr.rgb, None)
                hit.dpi = fr.dpi  # type: ignore[attr-defined]
                self.images[part] = hit
        except Exception as e:  # noqa: BLE001  開けない画像は灰色の枠で示す（白いままにしない）
            from ..limits import LimitExceeded

            if isinstance(e, LimitExceeded):
                raise
            self.note(f"画像 {path} を開けません（{e}）")
            self.undrawn(polys, c2)
            return
        if (b.get("TileMode") or "None") != "None":
            self.note("タイル張りの画像は 1 枚目だけを描きました")
        dpi = getattr(hit, "dpi", None) or (96, 96)
        iw, ih = hit.w * 96 / (dpi[0] or 96), hit.h * 96 / (dpi[1] or 96)  # 画像の大きさ（1/96 インチ）
        if len(vb) != 4 or vb[2] == 0 or vb[3] == 0:
            vb = [0, 0, iw, ih]
        sx, sy = vp[2] / vb[2], vp[3] / vb[3]
        x0 = vp[0] - vb[0] * sx
        y0 = vp[1] - vb[1] * sy
        # 単位の四角（画像は y=1 が上）→ ページの座標 → 画面
        unit = (iw * sx, 0.0, 0.0, -ih * sy, x0, y0 + ih * sy)
        draw_image(self.cv, hit, mul(unit, mt), c2, alpha)

    def undrawn(self, polys, clip: Clip) -> None:
        rows = fill_rows(polys, False, clip.rect)
        paint(self.cv, rows, (205, 205, 205), 1.0, clip)

    def walk(self, el: ET.Element, ctm, clip: Clip, alpha: float, scope: list, depth: int = 0) -> None:
        if depth > 200:
            raise XpsError("Canvas の入れ子が深すぎます")
        for k in el:
            if time.monotonic() > self.deadline:
                self.stopped = True
                return
            t = _ln(k.tag)
            if t not in ("Canvas", "Path", "Glyphs"):
                continue
            m = _transform(k, scope)
            sc = self.resources(k, scope)
            mm = mul(m, ctm) if m else ctm
            a = alpha * float(k.get("Opacity") or 1)
            if k.get("OpacityMask") or _prop(k, "OpacityMask") is not None:
                self.note("OpacityMask（透明の型）は無視しました")
            c = self.clip_of(k, mm, clip, sc)
            if c.empty:
                continue
            if t == "Canvas":
                self.walk(k, mm, c, a, sc, depth + 1)
            elif t == "Path":
                self.path(k, mm, c, a, sc)
            else:
                self.glyphs(k, mm, c, a, sc)

    def path(self, el: ET.Element, ctm, clip: Clip, alpha: float, scope: list) -> None:
        g = _element_geometry(el, "Data", scope)
        if g is None:
            return
        ops, eo = g
        if not ops:
            return
        subs = flatten(ops, ctm)
        polys = [p for p, _c in subs if len(p) >= 2]
        fill = self.brush_of(el, "Fill", scope)
        if fill is not None and polys:
            if not isinstance(fill, str) and _ln(fill.tag) == "ImageBrush":
                self.draw_image_brush(fill, polys, eo, ctm, clip, alpha * float(fill.get("Opacity") or 1), scope)
            else:
                c = self.solid(fill)
                if c is not None and c[3] * alpha > 0:
                    paint(self.cv, fill_rows(polys, eo, clip.rect), c[:3], c[3] * alpha, clip)
        stroke = self.brush_of(el, "Stroke", scope)
        if stroke is not None:
            c = self.solid(stroke)
            if c is not None:
                w = float(el.get("StrokeThickness") or 1)
                cap = {"Flat": 0, "Round": 1, "Square": 2}.get(el.get("StrokeStartLineCap") or "Flat", 0)
                join = {"Miter": 0, "Round": 1, "Bevel": 2}.get(el.get("StrokeLineJoin") or "Miter", 0)
                dash = _floats(el.get("StrokeDashArray") or "")
                dash = [d * w for d in dash] or None
                phase = float(el.get("StrokeDashOffset") or 0) * w
                keep = flatten(ops, ctm, keep=True)
                sc = math.sqrt(abs(ctm[0] * ctm[3] - ctm[1] * ctm[2])) or 1.0
                sp = stroke_polys([(p, cl) for p, cl in keep if p], w, cap, join, float(el.get("StrokeMiterLimit") or 10), dash, phase, sc)
                dev = transform_polys(sp, ctm)
                paint(self.cv, fill_rows(dev, False, clip.rect), c[:3], c[3] * alpha, clip)

    def glyphs(self, el: ET.Element, ctm, clip: Clip, alpha: float, scope: list) -> None:
        uri = el.get("FontUri") or ""
        if not uri:
            self.note("書体の指定のない字があります")
            return
        font = self.font(uri)
        if font is None:
            return
        fill = self.brush_of(el, "Fill", scope)
        c = self.solid(fill) if fill is not None else None
        if c is None:
            if fill is not None and not isinstance(fill, str) and _ln(fill.tag) == "ImageBrush":
                self.note("画像で塗った字は黒で描きました")
                c = (0, 0, 0, 1.0)
            else:
                return
        em = abs(float(el.get("FontRenderingEmSize") or 0))
        if em <= 0:
            return
        ox, oy = float(el.get("OriginX") or 0), float(el.get("OriginY") or 0)
        ltr = int(el.get("BidiLevel") or 0) % 2 == 0
        if (el.get("IsSideways") or "").lower() == "true":
            self.note("横向きに置いた字（IsSideways）は、向きを変えずに描きました")
        sim = el.get("StyleSimulations") or "None"
        shear = 0.33 if "Italic" in sim else 0.0
        bold = "Bold" in sim
        run = _glyph_run(el.get("Indices"), el.get("UnicodeString") or "", font, em, ltr)
        out = []
        for gid, x, y in run:
            ops, fm = font.outline(gid)
            if not ops:
                continue
            gm = mul(fm, (em, 0.0, shear * em, -em, ox + x, oy + y))
            out += [p for p, _c in flatten(ops, mul(gm, ctm)) if len(p) >= 2]
        if not out:
            return
        paint(self.cv, fill_rows(out, False, clip.rect), c[:3], c[3] * alpha, clip)
        if bold:
            sc = math.sqrt(abs(ctm[0] * ctm[3] - ctm[1] * ctm[2])) or 1.0
            sp = stroke_polys([(p, True) for p in out], em / 24 * sc, 1, 1, 10.0, None, 0.0, 1.0)
            paint(self.cv, fill_rows(sp, False, clip.rect), c[:3], c[3] * alpha, clip)


def page_size(root: ET.Element) -> tuple[float, float]:
    w, h = float(root.get("Width") or 0), float(root.get("Height") or 0)
    if w <= 0 or h <= 0:
        raise XpsError("FixedPage の大きさがありません")
    return w, h


def render(blob: bytes, page_no: int, region_mm: list | None = None, dpi: float | None = None, time_limit: float = 60.0) -> dict:
    z, names, plist = pages(blob)
    if not 1 <= page_no <= len(plist):
        raise ValueError(f"page は 1〜{len(plist)} です")
    part = plist[page_no - 1]
    try:
        root = ET.fromstring(_read_part(z, names, part))
    except ET.ParseError as e:
        raise XpsError(f"ページの部品を XML として読めません（{e}）") from None
    w, h = page_size(root)
    pw, ph = w * UNIT_MM, h * UNIT_MM
    if region_mm is not None and not (isinstance(region_mm, list) and len(region_mm) == 4):
        raise ValueError("region は [左, 上, 右, 下]（mm）で指定してください")
    x0, y0, x1, y1 = (float(v) for v in (region_mm or [0, 0, pw, ph]))
    x0, x1 = sorted((x0, x1))
    y0, y1 = sorted((y0, y1))
    x0, y0, x1, y1 = max(0.0, x0), max(0.0, y0), min(pw, x1), min(ph, y1)
    if x1 - x0 < 1 or y1 - y0 < 1:
        raise ValueError(f"範囲が小さすぎるか、ページ（{pw:.0f}×{ph:.0f} mm）の外です")
    if dpi:
        s = float(dpi) / 25.4
    elif region_mm:
        s = 1600 / max(x1 - x0, y1 - y0)
    else:
        s = 100 / 25.4
    s = min(s, MAX_SIDE / max(x1 - x0, y1 - y0), (MAX_PIXELS / ((x1 - x0) * (y1 - y0))) ** 0.5)
    W, H = max(1, math.ceil((x1 - x0) * s - 1e-6)), max(1, math.ceil((y1 - y0) * s - 1e-6))
    cv = Canvas(W, H)
    k = UNIT_MM * s
    base = (k, 0.0, 0.0, k, -x0 * s, -y0 * s)
    pt = _Painter(z, names, part, cv, base, time.monotonic() + time_limit)
    pt.walk(root, base, Clip(0, 0, W, H), 1.0, pt.resources(root, [{}]))
    return {"rgb": cv.buf, "w": W, "h": H, "scale": s, "origin_mm": (x0, y0), "page_mm": (pw, ph), "dpi": s * 25.4,
            "not_drawn": pt.not_drawn, "stopped": pt.stopped, "pages": len(plist)}


def view_page(blob: bytes, page_no: int, region: list | None = None, marks: list | None = None, dpi: float | None = None) -> dict:
    """view_page と同じ形の結果。"""
    from ..pdf.pageimage import with_ruler

    try:
        r = render(blob, page_no, region, dpi)
    except KeyError as e:
        raise XpsError(f"部品がありません（{e}）") from None
    png, w, h, notes = with_ruler(r, marks)
    ox, oy = r["origin_mm"]
    pw, ph = r["page_mm"]
    text = (f"{page_no} ページ（{pw:.0f}×{ph:.0f} mm。XPS）の、左 {ox:.0f}〜{ox + r['w'] / r['scale']:.0f} mm・上 {oy:.0f}〜{oy + r['h'] / r['scale']:.0f} mm の範囲"
            f"（1 mm＝{r['scale']:.1f} 画素・{r['dpi']:.0f} dpi）。上と左の帯の数は、ページの左上からの mm です。小さな字は region（[左, 上, 右, 下] mm）で拡大して読みます。"
            "XPS には書きません。読めない所は推測で埋めず、本人に伝えてください。")
    if r["pages"] > 1:
        text += f" 全部で {r['pages']} ページです。"
    if r["stopped"]:
        text += " 時間の上限で止めました（途中まで描いた画像です）。"
    if r["not_drawn"]:
        text += " 描けなかったもの・近い形で描いたもの：" + "／".join(r["not_drawn"][:8]) + "。"
    if notes:
        text += " " + " ".join(notes)
    return {"text": text, "png_base64": base64.b64encode(png).decode("ascii"), "width": w, "height": h, "not_drawn": list(r["not_drawn"])[:8]}
