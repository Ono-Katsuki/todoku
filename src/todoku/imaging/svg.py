"""SVG（W3C SVG 1.1・SVG 2）の静的な部分を、PDF の描画の部品（pdf/render/canvas.py）で画素に描く。

描くもの：path（全部の命令）・rect・circle・ellipse・line・polyline・polygon・text・tspan・image（data: の URI）・g・svg・use・symbol・defs・switch、
transform、塗りと線（色・不透明度・fill-rule・線の端と角・破線）、clipPath、style の要素と class（型・.class・#id の簡単な選び方だけ）、グラデーション。
描けないものは黙って飛ばさない：外部の画像・フィルター・マスク・模様・字のパスなどは、種類と数を problems に書く。グラデーションは近い色で描いて、そう伝える。

動かさない：script は実行しない。アニメーションも始まりの形のまま。外のファイルや URL は読まない。ENTITY のある XML は開かない（字の爆弾）。

大きさ：SVG の 1 単位（CSS の px）は 96 分の 1 インチ。ベクターなので決まった解像度がない。300 dpi を「元の解像度」として画素の数を示す。
要る範囲（window）と画素の数（target_pixels）に合わせて描くので、範囲を狭めれば字もなめらかなまま拡大できる。
"""

from __future__ import annotations

import base64
import math
import re
import time
import xml.etree.ElementTree as ET
import zlib
from urllib.parse import unquote_to_bytes

from .formats import Frame, FrameInfo, ImageError

NOMINAL_DPI = 300.0
K = NOMINAL_DPI / 96.0  # 1 CSS px の、元の解像度の画素
MAX_ELEMENTS = 400_000
MAX_DEPTH = 64
_UNITS = {"": 1.0, "px": 1.0, "pt": 96 / 72, "pc": 16.0, "mm": 96 / 25.4, "cm": 96 / 2.54, "q": 96 / 101.6, "in": 96.0}
_HEAD = re.compile(rb"\s*(?:\xef\xbb\xbf)?\s*(?:<\?xml.*?\?>\s*)?(?:<!--.*?-->\s*|<!DOCTYPE[^>\[]*(?:\[.*?\])?\s*>\s*|<\?.*?\?>\s*)*<(?:[A-Za-z_][\w.-]*:)?svg[\s>/]", re.S)

_NAMES = """aliceblue f0f8ff antiquewhite faebd7 aqua 00ffff aquamarine 7fffd4 azure f0ffff beige f5f5dc bisque ffe4c4 black 000000 blanchedalmond ffebcd blue 0000ff
blueviolet 8a2be2 brown a52a2a burlywood deb887 cadetblue 5f9ea0 chartreuse 7fff00 chocolate d2691e coral ff7f50 cornflowerblue 6495ed cornsilk fff8dc crimson dc143c cyan 00ffff
darkblue 00008b darkcyan 008b8b darkgoldenrod b8860b darkgray a9a9a9 darkgreen 006400 darkgrey a9a9a9 darkkhaki bdb76b darkmagenta 8b008b darkolivegreen 556b2f darkorange ff8c00
darkorchid 9932cc darkred 8b0000 darksalmon e9967a darkseagreen 8fbc8f darkslateblue 483d8b darkslategray 2f4f4f darkslategrey 2f4f4f darkturquoise 00ced1 darkviolet 9400d3
deeppink ff1493 deepskyblue 00bfff dimgray 696969 dimgrey 696969 dodgerblue 1e90ff firebrick b22222 floralwhite fffaf0 forestgreen 228b22 fuchsia ff00ff gainsboro dcdcdc
ghostwhite f8f8ff gold ffd700 goldenrod daa520 gray 808080 green 008000 greenyellow adff2f grey 808080 honeydew f0fff0 hotpink ff69b4 indianred cd5c5c indigo 4b0082 ivory fffff0
khaki f0e68c lavender e6e6fa lavenderblush fff0f5 lawngreen 7cfc00 lemonchiffon fffacd lightblue add8e6 lightcoral f08080 lightcyan e0ffff lightgoldenrodyellow fafad2
lightgray d3d3d3 lightgreen 90ee90 lightgrey d3d3d3 lightpink ffb6c1 lightsalmon ffa07a lightseagreen 20b2aa lightskyblue 87cefa lightslategray 778899 lightslategrey 778899
lightsteelblue b0c4de lightyellow ffffe0 lime 00ff00 limegreen 32cd32 linen faf0e6 magenta ff00ff maroon 800000 mediumaquamarine 66cdaa mediumblue 0000cd mediumorchid ba55d3
mediumpurple 9370db mediumseagreen 3cb371 mediumslateblue 7b68ee mediumspringgreen 00fa9a mediumturquoise 48d1cc mediumvioletred c71585 midnightblue 191970 mintcream f5fffa
mistyrose ffe4e1 moccasin ffe4b5 navajowhite ffdead navy 000080 oldlace fdf5e6 olive 808000 olivedrab 6b8e23 orange ffa500 orangered ff4500 orchid da70d6 palegoldenrod eee8aa
palegreen 98fb98 paleturquoise afeeee palevioletred db7093 papayawhip ffefd5 peachpuff ffdab9 peru cd853f pink ffc0cb plum dda0dd powderblue b0e0e6 purple 800080 rebeccapurple 663399
red ff0000 rosybrown bc8f8f royalblue 4169e1 saddlebrown 8b4513 salmon fa8072 sandybrown f4a460 seagreen 2e8b57 seashell fff5ee sienna a0522d silver c0c0c0 skyblue 87ceeb
slateblue 6a5acd slategray 708090 slategrey 708090 snow fffafa springgreen 00ff7f steelblue 4682b4 tan d2b48c teal 008080 thistle d8bfd8 tomato ff6347 turquoise 40e0d0
violet ee82ee wheat f5deb3 white ffffff whitesmoke f5f5f5 yellow ffff00 yellowgreen 9acd32""".split()
NAMED = {_NAMES[i]: tuple(bytes.fromhex(_NAMES[i + 1])) for i in range(0, len(_NAMES), 2)}
INHERIT = ("fill", "fill-opacity", "fill-rule", "stroke", "stroke-opacity", "stroke-width", "stroke-linecap", "stroke-linejoin", "stroke-miterlimit",
           "stroke-dasharray", "stroke-dashoffset", "font-size", "font-family", "font-weight", "font-style", "text-anchor", "color", "visibility", "clip-rule", "dominant-baseline")
SKIP_REPORT = {"filter": "フィルター", "mask": "マスク", "marker": "線の印（marker）", "foreignObject": "foreignObject（中の HTML）", "pattern": "模様（pattern）", "textPath": "パスに沿う字（まっすぐに並べました）"}


def is_svg(blob: bytes) -> bool:
    """先頭を見て SVG か（gzip の .svgz も）。"""
    head = blob[:8192]
    if head[:2] == b"\x1f\x8b":
        try:
            head = zlib.decompressobj(31).decompress(blob, 8192)
        except zlib.error:
            return False
    if head[:2] in (b"\xff\xfe", b"\xfe\xff"):
        try:
            head = head.decode("utf-16").encode("utf-8")
        except UnicodeDecodeError:
            return False
    return bool(_HEAD.match(head))


def _text_of(blob: bytes) -> bytes:
    from ..limits import inflate

    if blob[:2] == b"\x1f\x8b":
        return inflate(blob, wbits=31, what="SVGZ の中身", tolerant=True)
    return blob


def _root(blob: bytes):
    data = _text_of(blob)
    if b"<!ENTITY" in data[:65536]:
        raise ImageError("SVG：ENTITY の宣言のある XML は開きません（字の爆弾を防ぐため）")
    try:
        return ET.fromstring(data)
    except ET.ParseError as e:
        raise ImageError(f"SVG：XML を読めません（{e}）") from None


def _local(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _num(v, default: float = 0.0, base: float = 0.0, font: float = 16.0) -> float:
    """長さ（単位つき）→ CSS px。% は base に対して。"""
    if v is None:
        return default
    m = re.match(r"\s*([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)\s*([A-Za-z%]*)", str(v))
    if not m:
        return default
    x, u = float(m[1]), m[2].lower()
    if u == "%":
        return x * base / 100
    if u == "em":
        return x * font
    if u == "ex":
        return x * font / 2
    return x * _UNITS.get(u, 1.0)


def _numbers(s: str | None) -> list[float]:
    return [float(x) for x in re.findall(r"[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?", s or "")]


def _size(root) -> tuple[float, float, tuple | None, list]:
    """(幅, 高さ, viewBox, 知らせ)：CSS px。"""
    vb = _numbers(root.get("viewBox"))
    box = tuple(vb[:4]) if len(vb) >= 4 and vb[2] > 0 and vb[3] > 0 else None
    notes = []
    ws, hs = root.get("width"), root.get("height")
    pct = lambda s: s is None or str(s).strip().endswith("%")  # noqa: E731
    w = None if pct(ws) else _num(ws)
    h = None if pct(hs) else _num(hs)
    if box:
        ar = box[2] / box[3]
        if w is None and h is None:
            w, h = box[2], box[3]
        elif w is None:
            w = h * ar
        elif h is None:
            h = w / ar
    if not w or not h or w <= 0 or h <= 0:
        w, h = w or 300.0, h or 150.0
        notes.append(f"大きさの指定がないので、{w:.0f}×{h:.0f} の大きさで描きました")
    return w, h, box, notes


def infos(blob: bytes) -> list[FrameInfo]:
    w, h, _b, _n = _size(_root(blob))
    return [FrameInfo(max(1, round(w * K)), max(1, round(h * K)), (NOMINAL_DPI, NOMINAL_DPI))]


def texts(blob: bytes) -> list[str]:
    """text の要素の字（見える順）。"""
    out = []
    for el in _root(blob).iter():
        if _local(el.tag) == "text":
            s = " ".join("".join(el.itertext()).split())
            if s:
                out.append(s)
    return out


# ====================================================================== 色
def _color(v: str | None, current: tuple = (0, 0, 0)):
    """色の字 → (r, g, b, a) か None（none・分からない）。"""
    if v is None:
        return None
    s = v.strip().lower()
    if s in ("none", "transparent", ""):
        return None
    if s == "currentcolor":
        return (*current, 1.0)
    if s[0] == "#":
        h = s[1:]
        if len(h) in (3, 4):
            h = "".join(c * 2 for c in h)
        if len(h) in (6, 8) and all(c in "0123456789abcdef" for c in h):
            b = bytes.fromhex(h)
            return (b[0], b[1], b[2], b[3] / 255 if len(b) == 4 else 1.0)
        return None
    if s in NAMED:
        return (*NAMED[s], 1.0)
    m = re.match(r"(rgba?|hsla?)\((.*)\)", s)
    if m:
        parts = re.split(r"[,\s/]+", m[2].strip())
        try:
            if m[1].startswith("rgb"):
                ch = [min(255.0, max(0.0, float(p[:-1]) * 2.55 if p.endswith("%") else float(p))) for p in parts[:3]]
                al = parts[3] if len(parts) > 3 else "1"
                return (int(ch[0] + 0.5), int(ch[1] + 0.5), int(ch[2] + 0.5), float(al[:-1]) / 100 if al.endswith("%") else float(al))
            hh, ss, ll = float(parts[0].replace("deg", "")) / 360 % 1, float(parts[1].rstrip("%")) / 100, float(parts[2].rstrip("%")) / 100
            import colorsys

            r, g, b = colorsys.hls_to_rgb(hh, ll, ss)
            al = parts[3] if len(parts) > 3 else "1"
            return (int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5), float(al[:-1]) / 100 if al.endswith("%") else float(al))
        except (ValueError, IndexError):
            return None
    return None


# ====================================================================== 行列
def _mul(m1, m2):
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (a1 * a2 + b1 * c2, a1 * b2 + b1 * d2, c1 * a2 + d1 * c2, c1 * b2 + d1 * d2, e1 * a2 + f1 * c2 + e2, e1 * b2 + f1 * d2 + f2)


def _transform(s: str | None):
    """transform の字 → 行列（先に書いたものが外側）。"""
    m = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    for name, args in re.findall(r"([a-zA-Z]+)\s*\(([^)]*)\)", s or ""):
        a = _numbers(args)
        t = None
        if name == "matrix" and len(a) >= 6:
            t = tuple(a[:6])
        elif name == "translate" and a:
            t = (1, 0, 0, 1, a[0], a[1] if len(a) > 1 else 0.0)
        elif name == "scale" and a:
            t = (a[0], 0, 0, a[1] if len(a) > 1 else a[0], 0, 0)
        elif name == "rotate" and a:
            c, s_ = math.cos(math.radians(a[0])), math.sin(math.radians(a[0]))
            t = (c, s_, -s_, c, 0, 0)
            if len(a) >= 3:
                t = _mul(_mul((1, 0, 0, 1, -a[1], -a[2]), t), (1, 0, 0, 1, a[1], a[2]))
        elif name == "skewx" or name == "skewX":
            t = (1, 0, math.tan(math.radians(a[0])), 1, 0, 0) if a else None
        elif name == "skewy" or name == "skewY":
            t = (1, math.tan(math.radians(a[0])), 0, 1, 0, 0) if a else None
        if t is not None:
            m = _mul(t, m)
    return m


# ====================================================================== パス
_PATH_TOK = re.compile(r"([MmLlHhVvCcSsQqTtAaZz])|([-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?)")
_ARGS = {"m": 2, "l": 2, "h": 1, "v": 1, "c": 6, "s": 4, "q": 4, "t": 2, "a": 7, "z": 0}


def _path_tokens(d: str):
    """命令と数の並び。a の旗（0・1）は間の区切りなしに続くことがあるので、ここで分ける。"""
    i, n = 0, len(d)
    cmd = None
    out = []
    while i < n:
        ch = d[i]
        if ch.isspace() or ch == ",":
            i += 1
            continue
        if ch in "MmLlHhVvCcSsQqTtAaZz":
            cmd = ch
            out.append(ch)
            i += 1
            count = 0
            continue
        if cmd is None:
            return out
        if cmd in "aA" and count % 7 in (3, 4) and ch in "01":
            out.append(float(ch))
            i += 1
            count += 1
            continue
        m = _PATH_TOK.match(d, i)
        if not m or m[2] is None:
            return out  # 読めない所までを描く（SVG の決まり：誤りの手前まで）
        out.append(float(m[2]))
        i = m.end()
        count += 1
    return out


def _arc_to_cubics(x0, y0, rx, ry, phi, fa, fs, x, y) -> list:
    """F.6.5：端点の形 → 中心の形 → 3 次のベジエ。"""
    if rx == 0 or ry == 0:
        return [(x0, y0, x, y, x, y)]
    rx, ry = abs(rx), abs(ry)
    c, s = math.cos(phi), math.sin(phi)
    dx, dy = (x0 - x) / 2, (y0 - y) / 2
    x1, y1 = c * dx + s * dy, -s * dx + c * dy
    lam = x1 * x1 / (rx * rx) + y1 * y1 / (ry * ry)
    if lam > 1:
        k = math.sqrt(lam)
        rx, ry = rx * k, ry * k
    num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1
    den = rx * rx * y1 * y1 + ry * ry * x1 * x1
    co = math.sqrt(max(0.0, num / den)) if den else 0.0
    if fa == fs:
        co = -co
    cx1, cy1 = co * rx * y1 / ry, -co * ry * x1 / rx
    cx, cy = c * cx1 - s * cy1 + (x0 + x) / 2, s * cx1 + c * cy1 + (y0 + y) / 2
    a0 = math.atan2((y1 - cy1) / ry, (x1 - cx1) / rx)
    a1 = math.atan2((-y1 - cy1) / ry, (-x1 - cx1) / rx)
    da = a1 - a0
    if fs and da < 0:
        da += 2 * math.pi
    elif not fs and da > 0:
        da -= 2 * math.pi
    n = max(1, math.ceil(abs(da) / (math.pi / 2) - 1e-9))
    out = []
    step = da / n
    t = 4 / 3 * math.tan(step / 4)
    for i in range(n):
        p0, p1 = a0 + i * step, a0 + (i + 1) * step
        c0, s0, c1, s1 = math.cos(p0), math.sin(p0), math.cos(p1), math.sin(p1)
        bx = [c0 - t * s0, c1 + t * s1, c1]
        by = [s0 + t * c0, s1 - t * c1, s1]
        pts = []
        for ux, uy in zip(bx, by):
            pts += [c * rx * ux - s * ry * uy + cx, s * rx * ux + c * ry * uy + cy]
        out.append(tuple(pts))
    return out


def path_subs(d: str, scale: float) -> list:
    """パスの字 → (点の並び, 閉じたか) の並び（利用者の座標）。scale：1 単位の画素（曲線を分ける細かさ）。"""
    toks = _path_tokens(d)
    subs = []
    pts: list = []
    i = 0
    n = len(toks)
    cx = cy = sx = sy = 0.0
    lc = None  # 直前の制御点（S・T のため）
    prev = ""
    cmd = None

    def curve(x1, y1, x2, y2, x, y):
        nonlocal cx, cy
        d_ = (abs(cx - x1) + abs(y1 - cy) + abs(x2 - x1) + abs(y2 - y1) + abs(x - x2) + abs(y - y2)) * scale
        k = max(2, min(64, int(math.sqrt(d_)) + 2))
        for j in range(1, k + 1):
            t = j / k
            u = 1 - t
            pts.append((u ** 3 * cx + 3 * u * u * t * x1 + 3 * u * t * t * x2 + t ** 3 * x, u ** 3 * cy + 3 * u * u * t * y1 + 3 * u * t * t * y2 + t ** 3 * y))
        cx, cy = x, y

    while i < n:
        t = toks[i]
        if isinstance(t, str):
            cmd = t
            i += 1
            if cmd in "zZ":
                if pts:
                    subs.append((pts, True))
                pts = []
                cx, cy = sx, sy
                prev = "z"
                lc = None
                continue
        if cmd is None:
            break
        low = cmd.lower()
        need = _ARGS[low]
        if need == 0 or i + need > n or any(isinstance(v, str) for v in toks[i:i + need]):
            break
        a = toks[i:i + need]
        i += need
        rel = cmd.islower()
        ox, oy = (cx, cy) if rel else (0.0, 0.0)
        if low == "m":
            if pts:
                subs.append((pts, False))
            cx, cy = a[0] + ox, a[1] + oy
            sx, sy = cx, cy
            pts = [(cx, cy)]
            cmd = "l" if rel else "L"  # 続きの組は線
            prev = "m"
            lc = None
            continue
        if not pts:
            pts = [(cx, cy)]
        if low == "l":
            cx, cy = a[0] + ox, a[1] + oy
            pts.append((cx, cy))
            lc = None
        elif low == "h":
            cx = a[0] + ox
            pts.append((cx, cy))
            lc = None
        elif low == "v":
            cy = a[0] + oy
            pts.append((cx, cy))
            lc = None
        elif low == "c":
            x2, y2 = a[2] + ox, a[3] + oy
            curve(a[0] + ox, a[1] + oy, x2, y2, a[4] + ox, a[5] + oy)
            lc = (x2, y2)
        elif low == "s":
            x1, y1 = (2 * cx - lc[0], 2 * cy - lc[1]) if lc and prev in "cs" else (cx, cy)
            x2, y2 = a[0] + ox, a[1] + oy
            curve(x1, y1, x2, y2, a[2] + ox, a[3] + oy)
            lc = (x2, y2)
        elif low in ("q", "t"):
            if low == "q":
                qx, qy = a[0] + ox, a[1] + oy
                ex, ey = a[2] + ox, a[3] + oy
            else:
                qx, qy = (2 * cx - lc[0], 2 * cy - lc[1]) if lc and prev in "qt" else (cx, cy)
                ex, ey = a[0] + ox, a[1] + oy
            curve(cx + 2 / 3 * (qx - cx), cy + 2 / 3 * (qy - cy), ex + 2 / 3 * (qx - ex), ey + 2 / 3 * (qy - ey), ex, ey)
            lc = (qx, qy)
        elif low == "a":
            ex, ey = a[5] + ox, a[6] + oy
            for c in _arc_to_cubics(cx, cy, a[0], a[1], math.radians(a[2]), bool(a[3]), bool(a[4]), ex, ey):
                curve(*c)
            lc = None
        prev = low
    if pts:
        subs.append((pts, False))
    return subs


def _ellipse(cx, cy, rx, ry, scale) -> list:
    n = max(16, min(128, int(math.pi * 2 * max(rx, ry) * scale / 4) + 8))
    return [(cx + rx * math.cos(2 * math.pi * i / n), cy + ry * math.sin(2 * math.pi * i / n)) for i in range(n)]


# ====================================================================== 描く
class _Painter:
    def __init__(self, root, W: int, H: int, base, deadline: float | None):
        from ..pdf.render.canvas import Canvas, Clip

        self.root = root
        self.cv = Canvas(W, H)
        self.W, self.H = W, H
        self.deadline = deadline
        self.skipped: dict = {}
        self.approx: dict = {}
        self.ids: dict = {}
        self.count = 0
        self.stopped = False
        self.css: list = []
        self.texts: list[str] = []
        self.images = 0
        for el in root.iter():
            i = el.get("id")
            if i and i not in self.ids:
                self.ids[i] = el
            if _local(el.tag) == "style" and el.text:
                self.css += _parse_css(el.text)
        vw, vh, box, _ = _size(root)
        self.vp = (box[2], box[3]) if box else (vw, vh)
        self.clip0 = Clip(0, 0, W, H)
        self.base = base

    def skip(self, what: str, n: int = 1) -> None:
        self.skipped[what] = self.skipped.get(what, 0) + n

    def style_of(self, el, parent: dict) -> dict:
        """継承する値と、この要素の値（属性 < style の要素 < style の属性。CSS の詳しさの順）。"""
        st = dict(parent)
        for k in ("opacity", "transform", "clip-path", "display"):
            st.pop(k, None)
        own: dict = {}
        for k, v in el.attrib.items():
            if "}" not in k:
                own[k] = v
        name = _local(el.tag)
        cls = (el.get("class") or "").split()
        eid = el.get("id")
        for spec, _order, sel, decl in self.css:
            if _css_match(sel, name, cls, eid):
                own.update(decl)
        if el.get("style"):
            own.update(_parse_decl(el.get("style")))
        if "font" in own:  # 短く書いた font：[太さ・斜体] 大きさ[/行の高さ] 書体
            m = re.search(r"((?:[\d.]+)(?:px|pt|pc|mm|cm|in|em|ex|%)?)(?:/\S+)?\s+(.+)$", own["font"])
            if m:
                head = own["font"][:m.start()]
                own.setdefault("font-size", m[1])
                own.setdefault("font-family", m[2])
                if "bold" in head:
                    own.setdefault("font-weight", "bold")
                if "italic" in head:
                    own.setdefault("font-style", "italic")
        for k, v in own.items():
            if v.strip() != "inherit":
                st[k] = v.strip()
        return st

    def number(self, st: dict, key: str, default: float, base: float = 0.0) -> float:
        fs = self.font_size(st)
        return _num(st.get(key), default, base, fs)

    def font_size(self, st: dict) -> float:
        return _num(st.get("font-size"), 16.0, 16.0, 16.0)

    def run(self, m) -> None:
        st = {"fill": "black", "stroke": "none", "stroke-width": "1", "color": "black", "font-size": "16"}
        self.element(self.root, st, m, self.clip0, 1.0, 0, True)

    def element(self, el, parent: dict, m, clip, op: float, depth: int, root: bool = False) -> None:
        name = _local(el.tag)
        if not name or self.stopped:
            return
        self.count += 1
        if self.count > MAX_ELEMENTS:
            self.stopped = True
            return
        if self.deadline is not None and self.count & 63 == 0 and time.monotonic() > self.deadline:
            self.stopped = True
            return
        if depth > MAX_DEPTH:
            self.skip("入れ子が深すぎる所")
            return
        if name in ("defs", "symbol", "clipPath", "linearGradient", "radialGradient", "style", "title", "desc", "metadata", "script", "stop"):
            return
        if name in SKIP_REPORT and name != "textPath":
            self.skip(SKIP_REPORT[name])
            return
        st = self.style_of(el, parent)
        if st.get("display") == "none":
            return
        for key, label in (("filter", "フィルター"), ("mask", "マスク")):
            if st.get(key) and st[key] != "none":
                self.skip(f"{label}（なしで描きました）")
        o = op
        if st.get("opacity"):
            o *= max(0.0, min(1.0, self.number(st, "opacity", 1.0)))
        if st.get("transform") and name != "svg":
            m = _mul(_transform(st["transform"]), m)
        cp = st.get("clip-path")
        if cp and cp != "none":
            clip = self.clip_of(cp, m, clip, st)
            if clip.empty:
                return
        if name == "svg":
            if not root:
                m, clip = self.nested(el, st, m, clip)
            for ch in el:
                self.element(ch, st, m, clip, o, depth + 1)
        elif name == "switch":
            for ch in el:
                if _local(ch.tag) and not ch.get("requiredExtensions"):
                    self.element(ch, st, m, clip, o, depth + 1)
                    break
        elif name in ("g", "a"):
            for ch in el:
                self.element(ch, st, m, clip, o, depth + 1)
        elif name == "use":
            self.use(el, st, m, clip, o, depth)
        elif name == "text":
            self.text(el, st, m, clip, o)
        elif name == "image":
            self.image(el, st, m, clip, o)
        elif name in ("path", "rect", "circle", "ellipse", "line", "polyline", "polygon"):
            self.shape(el, name, st, m, clip, o)
        elif name in ("animate", "animateTransform", "animateMotion", "set"):
            self.skip("アニメーション（始まりの形で描きました）")
        elif el.tag.startswith("{http://www.w3.org/2000/svg}") or "}" not in el.tag:
            self.skip(f"描けない要素 {name}")

    def nested(self, el, st, m, clip):
        x, y = _num(el.get("x")), _num(el.get("y"))
        w = _num(el.get("width"), self.vp[0], self.vp[0])
        h = _num(el.get("height"), self.vp[1], self.vp[1])
        m = _mul((1, 0, 0, 1, x, y), m)
        vb = _numbers(el.get("viewBox"))
        if len(vb) >= 4 and vb[2] > 0 and vb[3] > 0:
            m = _mul(_view_box(vb[:4], w, h, el.get("preserveAspectRatio")), m)
        return m, clip

    def use(self, el, st, m, clip, o, depth):
        href = el.get("href") or el.get("{http://www.w3.org/1999/xlink}href") or ""
        if not href.startswith("#"):
            self.skip("外のファイルを指す use")
            return
        target = self.ids.get(href[1:])
        if target is None:
            self.skip("指す先のない use")
            return
        m = _mul((1, 0, 0, 1, _num(el.get("x")), _num(el.get("y"))), m)
        tn = _local(target.tag)
        if tn == "symbol":
            vb = _numbers(target.get("viewBox"))
            w = _num(el.get("width"), vb[2] if len(vb) >= 4 else self.vp[0], self.vp[0])
            h = _num(el.get("height"), vb[3] if len(vb) >= 4 else self.vp[1], self.vp[1])
            if len(vb) >= 4 and vb[2] > 0 and vb[3] > 0:
                m = _mul(_view_box(vb[:4], w, h, target.get("preserveAspectRatio")), m)
            s2 = self.style_of(target, st)
            for ch in target:
                self.element(ch, s2, m, clip, o, depth + 1)
        else:
            self.element(target, st, m, clip, o, depth + 1)

    def subs_of(self, el, name: str, st: dict, scale: float) -> tuple[list, bool]:
        """図形 → (点の並び, 閉じたか) の並び。"""
        vw, vh = self.vp
        fs = self.font_size(st)
        if name == "path":
            return path_subs(st.get("d") or el.get("d") or "", scale), True
        if name == "rect":
            x, y = _num(el.get("x"), 0, vw, fs), _num(el.get("y"), 0, vh, fs)
            w, h = _num(el.get("width"), 0, vw, fs), _num(el.get("height"), 0, vh, fs)
            if w <= 0 or h <= 0:
                return [], True
            rx_, ry_ = el.get("rx"), el.get("ry")
            rx = _num(rx_ if rx_ is not None else ry_, 0, vw, fs)
            ry = _num(ry_ if ry_ is not None else rx_, 0, vh, fs)
            rx, ry = min(rx, w / 2), min(ry, h / 2)
            if rx > 0 and ry > 0:
                d = (f"M{x + rx} {y}H{x + w - rx}A{rx} {ry} 0 0 1 {x + w} {y + ry}V{y + h - ry}A{rx} {ry} 0 0 1 {x + w - rx} {y + h}"
                     f"H{x + rx}A{rx} {ry} 0 0 1 {x} {y + h - ry}V{y + ry}A{rx} {ry} 0 0 1 {x + rx} {y}Z")
                return path_subs(d, scale), True
            return [([(x, y), (x + w, y), (x + w, y + h), (x, y + h)], True)], True
        if name == "circle":
            r = _num(el.get("r"), 0, (vw + vh) / 2, fs)
            if r <= 0:
                return [], True
            return [(_ellipse(_num(el.get("cx"), 0, vw, fs), _num(el.get("cy"), 0, vh, fs), r, r, scale), True)], True
        if name == "ellipse":
            rx, ry = _num(el.get("rx"), 0, vw, fs), _num(el.get("ry"), 0, vh, fs)
            if rx <= 0 or ry <= 0:
                return [], True
            return [(_ellipse(_num(el.get("cx"), 0, vw, fs), _num(el.get("cy"), 0, vh, fs), rx, ry, scale), True)], True
        if name == "line":
            return [([(_num(el.get("x1"), 0, vw, fs), _num(el.get("y1"), 0, vh, fs)), (_num(el.get("x2"), 0, vw, fs), _num(el.get("y2"), 0, vh, fs))], False)], False
        v = _numbers(el.get("points"))
        pts = [(v[i], v[i + 1]) for i in range(0, len(v) - 1, 2)]
        if not pts:
            return [], True
        return [(pts, name == "polygon")], True

    def shape(self, el, name, st, m, clip, o) -> None:
        from ..pdf.render.canvas import scale_of

        if st.get("visibility") in ("hidden", "collapse"):
            return
        scale = scale_of(m) or 1.0
        subs, fillable = self.subs_of(el, name, st, scale)
        if not subs:
            return
        if el.find("{http://www.w3.org/2000/svg}animate") is not None:
            self.skip("アニメーション（始まりの形で描きました）")
        if fillable:
            polys = [[_pt(m, p) for p in pts] for pts, _c in subs if len(pts) >= 2]
            self.fill(polys, st, m, clip, o, subs)
        self.stroke(subs, st, m, clip, o)

    def fill(self, polys, st, m, clip, o, subs=None) -> None:
        from ..pdf.render.canvas import fill_rows, paint, paint_fn

        v = st.get("fill", "black")
        if v == "none" or not polys:
            return
        a = o * max(0.0, min(1.0, _num(st.get("fill-opacity"), 1.0)))
        rule = st.get("fill-rule") == "evenodd"
        grad = self.gradient(v, polys, m, st, a)
        rows = fill_rows(polys, rule, clip.rect)
        if grad is not None:
            paint_fn(self.cv, rows, grad[0], grad[1], clip)
            return
        c = self.paint_color(v, st)
        if c is not None:
            paint(self.cv, rows, c[:3], a * c[3], clip)

    def paint_color(self, v: str, st: dict):
        cur = _color(st.get("color"), (0, 0, 0))
        cur3 = cur[:3] if cur else (0, 0, 0)
        if v.startswith("url("):
            target = self.ids.get(_url_id(v))
            rest = v.split(")", 1)[1].strip() if ")" in v else ""
            if target is None:
                c = _color(rest, cur3)
                if not rest:
                    self.skip("指す先のない塗り")
                return c
            if _local(target.tag) == "pattern":
                self.skip("模様（pattern）の塗り（中間の灰色で描きました）")
                return (128, 128, 128, 1.0)
            return _color(rest, cur3)
        c = _color(v, cur3)
        if c is None and v.strip().lower() not in ("none", "transparent"):
            self.skip(f"読めない色 {v[:20]}（黒で描きました）")
            return (0, 0, 0, 1.0)
        return c

    def stops(self, g) -> list:
        """グラデーションの (位置, 色) の並び（href で指された先の stop も）。"""
        seen = 0
        while g is not None and seen < 16:
            found = [s for s in g if _local(s.tag) == "stop"]
            if found:
                break
            href = g.get("href") or g.get("{http://www.w3.org/1999/xlink}href") or ""
            g = self.ids.get(href[1:]) if href.startswith("#") else None
            seen += 1
        else:
            found = []
        out = []
        for s in found or []:
            d = dict(s.attrib)
            if s.get("style"):
                d.update(_parse_decl(s.get("style")))
            off = d.get("offset", "0").strip()
            t = float(off[:-1]) / 100 if off.endswith("%") else float(_numbers(off)[0] if _numbers(off) else 0)
            c = _color(d.get("stop-color", "black")) or (0, 0, 0, 0.0)
            al = c[3] * _num(d.get("stop-opacity"), 1.0)
            out.append((max(0.0, min(1.0, t)), c[:3], al))
        return out

    def gradient(self, v: str, polys, m, st, a):
        """グラデーションの塗り：(色の関数, 不透明度)。透けの変化は平均にする（そのことを伝える）。"""
        if not v.startswith("url("):
            return None
        g = self.ids.get(_url_id(v))
        if g is None or _local(g.tag) not in ("linearGradient", "radialGradient"):
            return None
        stops = self.stops(g)
        self.approx["グラデーション"] = self.approx.get("グラデーション", 0) + 1
        if not stops:
            return None
        if len(stops) == 1:
            t0 = stops[0]
            return (lambda x, y, c=t0[1]: c), a * t0[2]
        xs = [p[0] for pl in polys for p in pl]
        ys = [p[1] for pl in polys for p in pl]
        units = _attr_chain(self, g, "gradientUnits") or "objectBoundingBox"
        gt = _transform(_attr_chain(self, g, "gradientTransform"))
        linear = _local(g.tag) == "linearGradient"
        if units == "userSpaceOnUse":
            to = _mul(gt, m)
            base = self.vp
        else:
            # 塗る形の外枠（画面の座標で測った。回した形では少しずれる）→ 単位の四角
            x0, y0, x1, y1 = min(xs), min(ys), max(xs), max(ys)
            to = _mul(gt, (max(x1 - x0, 1e-6), 0, 0, max(y1 - y0, 1e-6), x0, y0))
            base = (1.0, 1.0)
        det = to[0] * to[3] - to[1] * to[2]
        if abs(det) < 1e-12:
            return None
        ia, ib, ic, id_ = to[3] / det, -to[1] / det, -to[2] / det, to[0] / det
        ie, if_ = to[4], to[5]
        ls = _stop_lookup(stops)
        al = sum(s[2] for s in stops) / len(stops)
        def attr(k, dflt):
            raw = _attr_chain(self, g, k)
            if raw is None:
                return dflt
            s = raw.strip()
            if s.endswith("%"):
                return float(s[:-1]) / 100 * (base[0] if k.startswith(("x", "c", "f")) and not k.endswith("y") else base[1])
            return _num(s, dflt)

        if linear:
            gx0, gy0, gx1, gy1 = attr("x1", 0.0), attr("y1", 0.0), attr("x2", base[0]), attr("y2", 0.0)
            dx, dy = gx1 - gx0, gy1 - gy0
            dd = dx * dx + dy * dy or 1e-9

            def at(X, Y):
                x, y = X - ie, Y - if_
                u, v_ = ia * x + ic * y, ib * x + id_ * y
                t = ((u - gx0) * dx + (v_ - gy0) * dy) / dd
                return ls(t)
        else:
            cx, cy = attr("cx", base[0] / 2), attr("cy", base[1] / 2)
            r = attr("r", base[0] / 2) or 1e-9

            def at(X, Y):
                x, y = X - ie, Y - if_
                u, v_ = ia * x + ic * y, ib * x + id_ * y
                return ls(math.hypot(u - cx, v_ - cy) / r)
        return at, a * al

    def stroke(self, subs, st, m, clip, o) -> None:
        from ..pdf.render.canvas import fill_rows, paint, paint_fn, scale_of, stroke_polys, transform_polys

        v = st.get("stroke", "none")
        if v == "none":
            return
        w = self.number(st, "stroke-width", 1.0, (self.vp[0] + self.vp[1]) / 2)
        if w <= 0:
            return
        scale = scale_of(m) or 1.0
        cap = {"round": 1, "square": 2}.get(st.get("stroke-linecap", ""), 0)
        join = {"round": 1, "bevel": 2}.get(st.get("stroke-linejoin", ""), 0)
        miter = _num(st.get("stroke-miterlimit"), 4.0)
        ds = st.get("stroke-dasharray")
        dash = None
        if ds and ds != "none":
            arr = _numbers(ds)
            if arr and any(x > 0 for x in arr):
                dash = arr if len(arr) % 2 == 0 else arr * 2
        phase = _num(st.get("stroke-dashoffset"), 0.0)
        good = [(pts, c) for pts, c in subs if pts]
        polys = stroke_polys(good, max(w, 0.5 / scale), cap, join, miter, dash, phase, scale)
        if not polys:
            return
        polys = transform_polys(polys, m)
        a = o * max(0.0, min(1.0, _num(st.get("stroke-opacity"), 1.0)))
        rows = fill_rows(polys, False, clip.rect)
        grad = self.gradient(v, polys, m, st, a)
        if grad is not None:
            paint_fn(self.cv, rows, grad[0], grad[1], clip)
            return
        c = self.paint_color(v, st)
        if c is not None:
            paint(self.cv, rows, c[:3], a * c[3], clip)

    def clip_of(self, v: str, m, clip, st):
        from ..pdf.render.canvas import fill_rows, scale_of

        target = self.ids.get(_url_id(v))
        if target is None or _local(target.tag) != "clipPath":
            if v.startswith("url("):
                self.skip("指す先のない clip-path")
            else:
                self.skip("形で書いた clip-path（切り抜かずに描きました）")
            return clip
        cm = m
        if target.get("clipPathUnits") == "objectBoundingBox":
            self.skip("clipPathUnits の objectBoundingBox（切り抜かずに描きました）")
            return clip
        if target.get("transform"):
            cm = _mul(_transform(target.get("transform")), cm)
        scale = scale_of(cm) or 1.0
        polys = []
        rule = False
        for ch in target:
            nm = _local(ch.tag)
            s2 = self.style_of(ch, st)
            tm = cm
            if ch.get("transform"):
                tm = _mul(_transform(ch.get("transform")), cm)
            if nm == "use":
                href = ch.get("href") or ch.get("{http://www.w3.org/1999/xlink}href") or ""
                ch = self.ids.get(href[1:]) if href.startswith("#") else None
                if ch is None:
                    continue
                nm = _local(ch.tag)
            if nm in ("path", "rect", "circle", "ellipse", "polygon"):
                subs, _f = self.subs_of(ch, nm, s2, scale)
                polys += [[_pt(tm, p) for p in pts] for pts, _c in subs if len(pts) >= 3]
                rule = rule or s2.get("clip-rule") == "evenodd"
            elif nm == "text":
                self.skip("字の形の clipPath（切り抜きに使いませんでした）")
        if not polys:
            return type(clip)(0, 0, 0, 0, {})
        return clip.with_rows(fill_rows(polys, rule, clip.rect))

    def text(self, el, st, m, clip, o) -> None:
        all_text = " ".join("".join(el.itertext()).split())
        if all_text:
            self.texts.append(all_text)
        if st.get("visibility") in ("hidden", "collapse"):
            return
        pen = [_first_num(el.get("x"), self.vp[0]), _first_num(el.get("y"), self.vp[1])]
        pen[0] += _first_num(el.get("dx"), self.vp[0])
        pen[1] += _first_num(el.get("dy"), self.vp[1])
        runs: list = []  # (字, 様式, 位置を決め直す x, y, dx, dy)
        self.collect(el, st, runs, True)
        # text-anchor は、位置を決め直すごとの「ひとかたまり」で揃える
        block: list = []
        for r in runs + [None]:
            if r is None or r[2] is not None or r[3] is not None:
                if block:
                    pen = self.draw_block(block, pen, m, clip, o)
                block = []
                if r is None:
                    break
                if r[2] is not None:
                    pen[0] = r[2]
                if r[3] is not None:
                    pen[1] = r[3]
            pen[0] += r[4] if not block else 0.0
            pen[1] += r[5] if not block else 0.0
            block.append(r)

    def collect(self, el, st, runs: list, first: bool) -> None:
        name = _local(el.tag)
        if name in ("title", "desc", "metadata"):
            return
        if name != "text":
            if name == "textPath":
                self.skip(SKIP_REPORT["textPath"])
            elif name == "a":
                pass
            elif name != "tspan":
                return
            st = self.style_of(el, st)
            if st.get("display") == "none":
                return
            x, y = el.get("x"), el.get("y")
            head = (_first_num(x, self.vp[0]) if x else None, _first_num(y, self.vp[1]) if y else None,
                    _first_num(el.get("dx"), self.vp[0]), _first_num(el.get("dy"), self.vp[1]))
            if any(v for v in head[2:]) or head[0] is not None or head[1] is not None:
                runs.append(("", st, *head))
        if el.text:
            s = " ".join(el.text.split()) if not el.text.isspace() else " "
            if el.text[0].isspace() and s and not s.startswith(" ") and not first:
                s = " " + s
            if el.text[-1].isspace() and s and not s.endswith(" "):
                s += " "
            runs.append((s, st, None, None, 0.0, 0.0))
        for ch in el:
            self.collect(ch, st, runs, False)
            if ch.tail:
                s = " ".join(ch.tail.split()) if not ch.tail.isspace() else " "
                if ch.tail[0].isspace() and not s.startswith(" "):
                    s = " " + s
                if ch.tail[-1].isspace() and not s.endswith(" "):
                    s += " "
                runs.append((s, st, None, None, 0.0, 0.0))

    def draw_block(self, block: list, pen: list, m, clip, o) -> list:
        from ..pdf.render.canvas import fill_rows, flatten, paint, paint_fn
        from ..pdf.render.fontprog import CFFOutlines

        laid = []
        total = 0.0
        for r in block:
            s, st = r[0], r[1]
            size = self.font_size(st)
            faces = _faces(st, s)
            for ch in s:
                face = next((f for f in faces if f.lookup(ch) is not None), None)
                g = face.lookup(ch) if face else None
                if face is not None and g is not None and hasattr(face.outl, "advance") and face.outl.advance(g):
                    a = face.outl.advance(g) * size
                else:
                    a = size if ord(ch) >= 0x2E80 else size * 0.5
                    if ch.strip() and face is None:
                        self.skip("字の形がこのパソコンの書体にない字（字は窓にあります）")
                laid.append((face, g, ch, size, st, total))
                total += a
        if not laid:
            return pen
        anchor = block[0][1].get("text-anchor", "start")
        sx = pen[0] - (total if anchor == "end" else total / 2 if anchor == "middle" else 0.0)
        y = pen[1]
        for face, g, ch, size, st, at in laid:
            if face is None or g is None or not ch.strip():
                continue
            if st.get("visibility") in ("hidden", "collapse"):
                continue
            base = block[0][1]
            o2 = face.outl
            gm = o2.glyph_matrix(g) if isinstance(o2, CFFOutlines) else o2.matrix
            shift = 0.0
            db = st.get("dominant-baseline") or base.get("dominant-baseline")
            if db in ("middle", "central"):
                shift = 0.35 * size
            elif db in ("hanging", "text-before-edge"):
                shift = 0.8 * size
            M = _mul((gm[0] * size, 0, 0, -gm[3] * size, sx + at, y + shift), m)
            polys = [p for p, _c in flatten(o2.glyph(g), M) if len(p) >= 2]
            if not polys or clip.empty:
                continue
            v = st.get("fill", "black")
            a = o * max(0.0, min(1.0, _num(st.get("fill-opacity"), 1.0)))
            rows = fill_rows(polys, False, clip.rect)
            grad = self.gradient(v, polys, m, st, a)
            if grad is not None:
                paint_fn(self.cv, rows, grad[0], grad[1], clip)
            elif v != "none":
                c = self.paint_color(v, st)
                if c is not None:
                    paint(self.cv, rows, c[:3], a * c[3], clip)
        return [sx + total, y]

    def image(self, el, st, m, clip, o) -> None:
        from ..pdf.render.image import RGBA, draw_image
        from .formats import decode

        href = el.get("href") or el.get("{http://www.w3.org/1999/xlink}href") or ""
        vw, vh = self.vp
        x, y = _num(el.get("x"), 0, vw), _num(el.get("y"), 0, vh)
        w, h = _num(el.get("width"), 0, vw), _num(el.get("height"), 0, vh)
        if not href.startswith("data:"):
            self.skip("外のファイルを指す画像（読みません。その所は描いていません）")
            self.hole(x, y, w, h, m, clip)
            return
        try:
            head, _, body = href[5:].partition(",")
            raw = base64.b64decode(re.sub(r"\s+", "", body)) if head.endswith(";base64") else unquote_to_bytes(body)
            fr = decode(raw, 0, None, self.deadline)
        except Exception as e:  # noqa: BLE001
            self.skip(f"開けない埋め込みの画像（{type(e).__name__}）")
            self.hole(x, y, w, h, m, clip)
            return
        if w <= 0 or h <= 0:
            w, h = (w or fr.w * (h / fr.h if h else 1)), (h or fr.h * (w / fr.w if w else 1))
        self.images += 1
        for p in fr.problems:
            self.skip(f"埋め込みの画像：{p}")
        t = _mul(_view_box((0, 0, fr.w, fr.h), w, h, el.get("preserveAspectRatio")), _mul((1, 0, 0, 1, x, y), m))
        # 画像の単位の四角 → 画面（画像の v は下から）
        draw_image(self.cv, RGBA(fr.w, fr.h, fr.rgb, None), _mul((fr.w, 0, 0, -fr.h, 0, fr.h), t), clip, o)

    def hole(self, x, y, w, h, m, clip) -> None:
        """描けない画像：灰色の枠と対角線（白いままにしない）。"""
        from ..pdf.render.canvas import fill_rows, paint, stroke_polys

        if w <= 0 or h <= 0:
            return
        pts = [_pt(m, p) for p in ((x, y), (x + w, y), (x + w, y + h), (x, y + h))]
        polys = stroke_polys([(pts, True), ([pts[0], pts[2]], False), ([pts[1], pts[3]], False)], 1.0, 0, 0, 10.0, None, 0.0, 1.0)
        if polys:
            paint(self.cv, fill_rows(polys, False, clip.rect), (150, 150, 150), 1.0, clip)


def _attr_chain(p: _Painter, g, key: str):
    """グラデーションの属性（href の先から受け継ぐ）。"""
    n = 0
    while g is not None and n < 16:
        v = g.get(key)
        if v is not None:
            return v
        href = g.get("href") or g.get("{http://www.w3.org/1999/xlink}href") or ""
        g = p.ids.get(href[1:]) if href.startswith("#") else None
        n += 1
    return None


def _stop_lookup(stops: list):
    """位置 → 色（256 段に前もって計算）。"""
    table = []
    for i in range(257):
        t = i / 256
        if t <= stops[0][0]:
            table.append(stops[0][1])
            continue
        for k in range(len(stops) - 1):
            a, b = stops[k], stops[k + 1]
            if t <= b[0]:
                u = 0 if b[0] <= a[0] else (t - a[0]) / (b[0] - a[0])
                table.append(tuple(int(p + (q - p) * u + 0.5) for p, q in zip(a[1], b[1])))
                break
        else:
            table.append(stops[-1][1])

    def look(t: float):
        return table[0 if t <= 0 else 256 if t >= 1 else int(t * 256)]

    return look


def _pt(m, p):
    return (m[0] * p[0] + m[2] * p[1] + m[4], m[1] * p[0] + m[3] * p[1] + m[5])


def _url_id(v: str) -> str:
    m = re.match(r"url\(\s*['\"]?#([^'\")\s]+)", v)
    return m[1] if m else ""


def _first_num(v, base: float) -> float:
    if not v:
        return 0.0
    return _num(re.split(r"[\s,]+", v.strip())[0], 0.0, base)


def _faces(st: dict, s: str) -> list:
    from ..pdf.render.fontmap import system_face

    fam = (st.get("font-family") or "").lower()
    bold = st.get("font-weight", "") in ("bold", "bolder", "600", "700", "800", "900")
    if "mono" in fam or "courier" in fam or "consol" in fam:
        kind = "mono"
    elif "serif" in fam.replace("sans-serif", "") or "times" in fam or "georgia" in fam or "mincho" in fam or "明朝" in fam:
        kind = "serif"
    else:
        kind = "sans"
    cjk = any(ord(c) >= 0x2E80 for c in s)
    order = []
    if cjk:
        order.append(("cjk-serif" if kind == "serif" else "cjk-sans") + ("-bold" if bold and kind != "serif" else ""))
    order += [kind + ("-bold" if bold else ""), "cjk-sans", "fallback"]
    return [f for f in (system_face(k) for k in order) if f is not None]


def _view_box(vb, w: float, h: float, par: str | None):
    """viewBox → 表示の四角の行列（7.8：preserveAspectRatio）。"""
    vx, vy, vw, vh = vb
    if vw <= 0 or vh <= 0:
        return (1, 0, 0, 1, 0, 0)
    sx, sy = w / vw, h / vh
    p = (par or "xMidYMid meet").split()
    align = p[0] if p else "xMidYMid"
    if align == "defer" and len(p) > 1:
        p = p[1:]
        align = p[0]
    if align != "none":
        s = max(sx, sy) if (len(p) > 1 and p[1] == "slice") else min(sx, sy)
        sx = sy = s
    tx, ty = -vx * sx, -vy * sy
    if "xMid" in align:
        tx += (w - vw * sx) / 2
    elif "xMax" in align:
        tx += w - vw * sx
    if "YMid" in align:
        ty += (h - vh * sy) / 2
    elif "YMax" in align:
        ty += h - vh * sy
    return (sx, 0, 0, sy, tx, ty)


# ====================================================================== CSS（style の要素：簡単な選び方だけ）
def _parse_decl(s: str) -> dict:
    out = {}
    for part in s.split(";"):
        k, sep, v = part.partition(":")
        if sep:
            out[k.strip().lower()] = v.replace("!important", "").strip()
    return out


def _parse_css(text: str) -> list:
    """(詳しさ, 順, 選び方, 宣言) の並び。@規則と子孫の選び方は使わない。"""
    text = re.sub(r"/\*.*?\*/", "", text, flags=re.S)
    text = text.replace("<![CDATA[", "").replace("]]>", "")
    out = []
    for i, (sels, body) in enumerate(re.findall(r"([^{}@]+)\{([^{}]*)\}", text)):
        decl = _parse_decl(body)
        for sel in sels.split(","):
            sel = sel.strip()
            if not sel or " " in sel or ">" in sel or ":" in sel or "[" in sel:
                continue
            spec = sel.count("#") * 100 + sel.count(".") * 10 + (0 if sel[0] in "#.*" else 1)
            out.append((spec, i, sel, decl))
    out.sort(key=lambda t: (t[0], t[1]))
    return out


def _css_match(sel: str, name: str, cls: list, eid: str | None) -> bool:
    if sel == "*":
        return True
    m = re.fullmatch(r"([A-Za-z][\w-]*)?((?:[.#][\w-]+)*)", sel)
    if not m:
        return False
    if m[1] and m[1] != name:
        return False
    for kind, val in re.findall(r"([.#])([\w-]+)", m[2]):
        if kind == "." and val not in cls:
            return False
        if kind == "#" and val != eid:
            return False
    return True


# ====================================================================== 入口
def render(blob: bytes, target_pixels: int | None = None, deadline: float | None = None, window: tuple | None = None) -> Frame:
    """SVG を描く。window：元の解像度（300 dpi）の画素の範囲。target_pixels：画像の全体を描いたときの画素の数（細かさを決める）。"""
    from ..limits import budget

    root = _root(blob)
    if _local(root.tag) != "svg":
        raise ImageError("SVG：いちばん外の要素が svg ではありません")
    vw, vh, box, notes = _size(root)
    FW, FH = max(1, round(vw * K)), max(1, round(vh * K))
    if target_pixels is None:
        target_pixels = min(FW * FH, 12_000_000)
    # 出す画素：元の解像度に対する倍率 k（元の解像度を超えない。範囲を狭めれば、画面に足りる細かさで描く）
    k = min(1.0, (target_pixels / (FW * FH)) ** 0.5)
    x0, y0, x1, y1 = (0, 0, FW, FH) if window is None else window
    x0, y0, x1, y1 = max(0, int(x0)), max(0, int(y0)), min(FW, int(x1)), min(FH, int(y1))
    if x1 <= x0 or y1 <= y0:
        x0, y0, x1, y1 = 0, 0, FW, FH
    W, H = max(1, round((x1 - x0) * k)), max(1, round((y1 - y0) * k))
    budget().charge(W * H * 3, "SVG の画素")
    sx = W / (x1 - x0) * K  # 1 CSS px の画素
    sy = H / (y1 - y0) * K
    base = (sx, 0.0, 0.0, sy, -x0 * W / (x1 - x0), -y0 * H / (y1 - y0))
    if box:
        base = _mul(_view_box(box, vw, vh, root.get("preserveAspectRatio")), base)
    if root.get("transform"):
        base = _mul(_transform(root.get("transform")), base)
    p = _Painter(root, W, H, base, deadline)
    p.run(base)
    notes = list(notes)
    if p.texts:
        notes.append("絵の中の字：" + "／".join(p.texts[:200]))
    if p.approx:
        notes.append("仕様どおりには描けず、近い色で描いたもの：" + "・".join(f"{a} {b} か所" for a, b in p.approx.items()) + "（透けの変化は平均にしています）")
    problems = [f"{a}（{b} か所）" for a, b in p.skipped.items()]
    if p.stopped:
        problems.append("要素が多すぎるか時間がかかりすぎるため、途中で止めました（描いていない所があります）")
    return Frame(W, H, bytes(p.cv.buf), (NOMINAL_DPI * W / (x1 - x0), NOMINAL_DPI * H / (y1 - y0)), notes, problems, (x0, y0, x1, y1) if window is not None else None)
