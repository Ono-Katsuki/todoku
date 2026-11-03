"""ページの描画の命令（コンテンツストリーム）を順に読み、文字（位置つき）・線・四角形・画像を取り出す。

位置はページの既定の座標（左下が原点、単位はポイント）で表す。ページの回転は layout.py で扱う。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .fonts import UNKNOWN, PdfFont, legacy_text, load_font
from .objects import Lexer, Name, PdfString, Ref, Stream, parse_object

IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def mul(m1, m2):
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (a1 * a2 + b1 * c2, a1 * b2 + b1 * d2, c1 * a2 + d1 * c2, c1 * b2 + d1 * d2,
            e1 * a2 + f1 * c2 + e2, e1 * b2 + f1 * d2 + f2)


def apply(m, x, y):
    return (m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5])


@dataclass
class Glyph:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size: float  # ページ上の文字の大きさ（ポイント）
    ox: float  # 文字の原点
    oy: float
    dx: float  # 送りの向き（ページ上）
    dy: float
    vertical: bool
    invisible: bool
    font: str
    color: tuple
    mcid: int | None = None
    actual: bool = False  # ActualText で置き換えた文字
    legacy: bool = False  # Unicode でない符号の書体で描いた字（取り出した字は元の文字ではない）
    adv: float = 0.0  # フォントの決める、この字の送りの幅（ページ上の長さ。字間・語間の調整は含まない）

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


@dataclass
class Seg:
    x0: float
    y0: float
    x1: float
    y1: float
    width: float
    filled: bool = False  # 細い塗りつぶしの四角を線として扱ったもの
    color: tuple = (0.0, 0.0, 0.0)


@dataclass
class Box:
    x0: float
    y0: float
    x1: float
    y1: float
    fill: tuple | None  # 塗りの色（RGB）
    stroke: bool


@dataclass
class Picture:
    x0: float
    y0: float
    x1: float
    y1: float
    name: str
    alt: str | None = None  # 印の付いた範囲（BDC）にある説明（Alt・ActualText）
    artifact: bool = False  # 飾り（Artifact）と印が付いている
    mc: tuple | None = None  # 内側の MCID：(持ち主, MCID)。持ち主は "page" か StructParents の番号
    objparent: int | None = None  # 画像（か、包むフォームの部品）の StructParent
    ctm: tuple | None = None  # 画像を置いた行列（画像の単位の四角 → ページの座標）。ページの見た目を描くのに使う
    image: object = None  # 画像の XObject（Stream）か、埋め込みの画像の (辞書, データ)


@dataclass
class _Mark:
    tag: str
    mcid: int | None
    owner: object
    alt: str | None
    artifact: bool


@dataclass
class PageContent:
    glyphs: list[Glyph] = field(default_factory=list)
    segs: list[Seg] = field(default_factory=list)
    boxes: list[Box] = field(default_factory=list)
    pictures: list[Picture] = field(default_factory=list)
    curves: list[tuple] = field(default_factory=list)  # 曲線を含む図形の外枠（丸など）
    unmapped: int = 0
    problems: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------- 字句
_TOK = re.compile(rb"""
 (?P<ws>[\x00\t\n\x0c\r ]+|%[^\r\n]*)
|(?P<num>[+-]?(?:\d+\.?\d*|\.\d+))(?![^\x00\t\n\x0c\r ()<>\[\]{}/%])
|(?P<name>/[^\x00\t\n\x0c\r ()<>\[\]{}/%]*)
|(?P<lit>\()
|(?P<dict><<|>>)
|(?P<hex><[0-9A-Fa-f\x00\t\n\x0c\r ]*>)
|(?P<arr>[\[\]])
|(?P<brace>[{}])
|(?P<op>[^\x00\t\n\x0c\r ()<>\[\]{}/%]+)
""", re.X)


def _hexstr(b: bytes) -> PdfString:
    h = re.sub(rb"[^0-9A-Fa-f]", b"", b)
    if len(h) % 2:
        h += b"0"
    s = PdfString(bytes.fromhex(h.decode("latin-1")))
    s.hexform = True
    return s


def tokens(data: bytes):
    """(種類, 値) を順に返す。種類: n 数 / s 文字列 / N 名前 / o 命令 / [ ] / << >> / I インライン画像のデータ。"""
    from ..limits import budget

    b = budget()  # 1 回の呼び出しで読む字句の数の合計（limits.MAX_PDF_OPS）：どの読み方で読んでも数える
    pos = 0
    n = len(data)
    match = _TOK.match
    while pos < n:
        m = match(data, pos)
        if m is None:
            pos += 1
            continue
        kind = m.lastgroup
        pos = m.end()
        if kind == "ws":
            continue
        b.ops += 1
        if b.ops > b.ops_total:
            raise b.ops_exceeded()
        t = m.group()
        if kind == "num":
            yield ("n", float(t) if b"." in t else int(t))
        elif kind == "name":
            yield ("N", Name(re.sub(rb"#([0-9A-Fa-f]{2})", lambda mm: bytes([int(mm[1], 16)]), t[1:]).decode("latin-1")))
        elif kind == "lit":
            lx = Lexer(data, m.start())
            s = lx._literal()
            pos = lx.pos
            yield ("s", s)
        elif kind == "hex":
            yield ("s", _hexstr(t[1:-1]))
        elif kind == "dict":
            yield (t.decode(), None)
        elif kind == "arr":
            yield (t.decode(), None)
        elif kind == "brace":
            continue
        else:
            op = t.decode("latin-1")
            if op == "ID":
                # インライン画像のデータ：空白1つの後から、空白に続く EI まで
                start = pos + 1
                e = start
                while True:
                    k = data.find(b"EI", e)
                    if k < 0:
                        k = n
                        break
                    before = data[k - 1:k]
                    after = data[k + 2:k + 3]
                    if before in (b" ", b"\n", b"\r", b"\t", b"\x00") and (after == b"" or after in b" \n\r\t\x00/[<(%"):
                        break
                    e = k + 2
                yield ("I", data[start:k])
                yield ("o", "EI")
                pos = k + 2
                continue
            yield ("o", op)


# ---------------------------------------------------------------------- 状態
class _GS:
    __slots__ = ("ctm", "lw", "fill", "stroke", "font", "fs", "tc", "tw", "th", "tl", "tr", "rise", "fontobj")

    def __init__(self):
        self.ctm = IDENT
        self.lw = 1.0
        self.fill = (0.0, 0.0, 0.0)
        self.stroke = (0.0, 0.0, 0.0)
        self.font = ""
        self.fontobj: PdfFont | None = None
        self.fs = 0.0
        self.tc = 0.0
        self.tw = 0.0
        self.th = 1.0
        self.tl = 0.0
        self.tr = 0
        self.rise = 0.0

    def copy(self):
        g = _GS.__new__(_GS)
        for k in _GS.__slots__:
            setattr(g, k, getattr(self, k))
        return g


def _gray(v):
    return (v, v, v)


def _cmyk(c, m, y, k):
    return ((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))


class Interpreter:
    def __init__(self, doc):
        self.doc = doc
        self.out = PageContent()

    def run_page(self, page: dict) -> PageContent:
        res = self.doc.resolve(page.get("Resources")) or {}
        data = self.doc.page_content(page)
        self._run(data, res, _GS(), depth=0, owner="page")
        self.out.unmapped = sum(f.unmapped for f in self._fonts_used.values()) if hasattr(self, "_fonts_used") else 0
        return self.out

    def run_stream(self, data: bytes, res: dict, ctm=IDENT) -> PageContent:
        gs = _GS()
        gs.ctm = ctm
        self._run(data, res or {}, gs, depth=0)
        return self.out

    # ------------------------------------------------------------------
    def _font(self, res: dict, name: str) -> PdfFont | None:
        fonts = self.doc.resolve(res.get("Font")) or {}
        ref = fonts.get(name) if isinstance(fonts, dict) else None
        if ref is None:
            return None
        f = load_font(self.doc, ref, name)
        if f is not None:
            self.__dict__.setdefault("_fonts_used", {})[id(f)] = f
        return f

    def _run(self, data: bytes, res: dict, gs: _GS, depth: int, outer: tuple = (), owner=None,
             objparent: int | None = None) -> None:
        if depth > 12:
            return
        from ..limits import MAX_PAGE_OPS, LimitExceeded

        drawing = self.__dict__.setdefault("_drawing", set())  # いま描いている途中のフォームの部品（自分を描く部品は描かない）
        doc = self.doc
        out = self.out
        stack: list = []
        gstack: list[_GS] = []
        path: list = []  # 部分の図形：[("m", x, y) ...] をページの座標で
        cur = None
        start = None
        tm = IDENT
        tlm = IDENT
        in_text = False
        mc_stack: list = []  # (mcid, actual_text, 先頭のglyphの位置)
        arr_stack: list = []
        dict_stack: list = []

        def push(v):
            if arr_stack:
                arr_stack[-1].append(v)
            elif dict_stack:
                dict_stack[-1].append(v)
            else:
                stack.append(v)

        def show(s: bytes, adjust_list=None):
            nonlocal tm
            f = gs.fontobj
            if f is None:
                return
            fs = gs.fs
            th = gs.th
            items = f.decode(s)
            mcid = mc_stack[-1][0] if mc_stack else None
            for code, text, w, is_space, vpos, vwidth in items:
                trm = mul((fs * th, 0, 0, fs, 0, gs.rise), mul(tm, gs.ctm))
                if f.vertical:
                    vx, vy = vpos
                    x0, x1 = -vx, -vx + (vwidth or 1.0)
                    y0, y1 = -vy + f.descent * 0 - 1.0 + 1.0 - 1.0 * 0, -vy + 1.0
                    # 縦書き：原点から下へ1字ぶん
                    y0, y1 = -vy, -vy + (f.ascent - f.descent)
                    y0, y1 = y1 - 1.0, y1
                    corners = [apply(trm, x0, y0), apply(trm, x1, y0), apply(trm, x0, y1), apply(trm, x1, y1)]
                    ty = w * fs + gs.tc + (gs.tw if is_space else 0)
                    ox, oy = apply(trm, 0, 0)
                    nx, ny = apply(mul((1, 0, 0, 1, 0, ty), mul(tm, gs.ctm)), 0, 0)
                    fx, fy = apply(mul((1, 0, 0, 1, 0, w * fs), mul(tm, gs.ctm)), 0, 0)
                    tm = mul((1, 0, 0, 1, 0, ty), tm)
                else:
                    x0, x1 = 0.0, w if w > 0 else 0.5
                    corners = [apply(trm, x0, f.descent), apply(trm, x1, f.descent), apply(trm, x0, f.ascent),
                               apply(trm, x1, f.ascent)]
                    tx = (w * fs + gs.tc + (gs.tw if is_space else 0)) * th
                    ox, oy = apply(trm, 0, 0)
                    fx, fy = apply(trm, w, 0)
                    tm = mul((1, 0, 0, 1, tx, 0), tm)
                    nx, ny = apply(mul((fs * th, 0, 0, fs, 0, gs.rise), mul(tm, gs.ctm)), 0, 0)
                xs = [c[0] for c in corners]
                ys = [c[1] for c in corners]
                a, b, c, d = trm[0], trm[1], trm[2], trm[3]
                size = (abs(a * d - b * c)) ** 0.5 if (a or b or c or d) else 0.0
                out.glyphs.append(Glyph(text, min(xs), min(ys), max(xs), max(ys), size, ox, oy, nx - ox, ny - oy,
                                        f.vertical, gs.tr in (3, 7), gs.font, gs.fill, mcid,
                                        legacy=legacy_text(f.legacy, text),
                                        adv=((fx - ox) ** 2 + (fy - oy) ** 2) ** 0.5))

        def adjust(v):
            nonlocal tm
            f = gs.fontobj
            if f is None:
                return
            if f.vertical:
                tm = mul((1, 0, 0, 1, 0, -v / 1000.0 * gs.fs), tm)
            else:
                tm = mul((1, 0, 0, 1, -v / 1000.0 * gs.fs * gs.th, 0), tm)

        def flush_path(stroke: bool, fill: bool):
            if not path:
                return
            ctm_scale = (abs(gs.ctm[0] * gs.ctm[3] - gs.ctm[1] * gs.ctm[2])) ** 0.5
            lw = gs.lw * ctm_scale if gs.lw > 0 else ctm_scale * 0.1
            for sub in path:
                pts = sub["pts"]
                if sub.get("curve"):
                    xs = [p[0] for p in pts]
                    ys = [p[1] for p in pts]
                    out.curves.append((min(xs), min(ys), max(xs), max(ys), stroke, fill))
                    if not sub.get("lines"):
                        continue
                rect = _as_rect(pts, sub.get("closed"))
                if fill and rect is not None:
                    x0, y0, x1, y1 = rect
                    if max(x1 - x0, y1 - y0) <= 2.5:
                        # 点（点線の1つぶ）：縦にも横にもつながりうるので、両方の向きの短い線として残す
                        out.segs.append(Seg(x0, (y0 + y1) / 2, x1, (y0 + y1) / 2, y1 - y0, True, gs.fill))
                        out.segs.append(Seg((x0 + x1) / 2, y0, (x0 + x1) / 2, y1, x1 - x0, True, gs.fill))
                        continue
                    if min(x1 - x0, y1 - y0) <= 2.5:
                        # 細い塗りつぶし：線として扱う
                        if x1 - x0 >= y1 - y0:
                            out.segs.append(Seg(x0, (y0 + y1) / 2, x1, (y0 + y1) / 2, y1 - y0, True, gs.fill))
                        else:
                            out.segs.append(Seg((x0 + x1) / 2, y0, (x0 + x1) / 2, y1, x1 - x0, True, gs.fill))
                        continue
                    out.boxes.append(Box(x0, y0, x1, y1, gs.fill, stroke))
                    if stroke:
                        _rect_edges(out, rect, lw, gs.stroke)
                    continue
                if stroke:
                    segs = sub.get("lines") or []
                    for (ax, ay, bx, by) in segs:
                        out.segs.append(Seg(ax, ay, bx, by, lw, False, gs.stroke))
                    if sub.get("closed") and len(pts) >= 2:
                        ax, ay = pts[-1]
                        bx, by = pts[0]
                        if (ax, ay) != (bx, by):
                            out.segs.append(Seg(ax, ay, bx, by, lw, False, gs.stroke))
                    if rect is not None and not fill:
                        out.boxes.append(Box(*rect, None, True))
                elif fill and len(pts) >= 3:
                    xs = [p[0] for p in pts]
                    ys = [p[1] for p in pts]
                    out.boxes.append(Box(min(xs), min(ys), max(xs), max(ys), gs.fill, False)) if False else None

        def picture(pts, name, own_parent=None):
            chain = outer + tuple(m[3] for m in mc_stack)
            inner = next(((m.owner, m.mcid) for m in reversed(chain) if m.mcid is not None and m.owner is not None), None)
            return Picture(min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts), name,
                           next((m.alt for m in reversed(chain) if m.alt), None), any(m.artifact for m in chain), inner,
                           own_parent if own_parent is not None else objparent)

        def new_sub(x, y):
            sub = {"pts": [(x, y)], "lines": [], "closed": False}
            path.append(sub)
            return sub

        sub = None
        inline_img = None
        for kind, val in tokens(data):
            # 1 ページ（とその部品）の字句の数の上限：部品が部品を何度も描いて、読み終わらない PDF を止める
            # （1 回の呼び出しの合計は tokens が数える）
            self._ops = getattr(self, "_ops", 0) + 1
            if self._ops > MAX_PAGE_OPS:
                raise LimitExceeded(f"1 ページの描画の字句（数・名前・命令）が {MAX_PAGE_OPS:,} 個を超えます（部品が何度も部品を描くなど）")
            if kind != "o":
                if kind == "[":
                    arr_stack.append([])
                elif kind == "]":
                    if arr_stack:
                        a = arr_stack.pop()
                        push(a)
                elif kind == "<<":
                    dict_stack.append([])
                elif kind == ">>":
                    if dict_stack:
                        items = dict_stack.pop()
                        d = {}
                        for i in range(0, len(items) - 1, 2):
                            k = items[i]
                            if isinstance(k, Name):
                                d[k.name] = items[i + 1]
                        push(d)
                elif kind == "I":
                    # インライン画像：BI の後の辞書（名前と値の組）とデータを覚えておく
                    d = {}
                    for i in range(0, len(stack) - 1, 2):
                        if isinstance(stack[i], Name):
                            d[stack[i].name] = stack[i + 1]
                    inline_img = (d, val)
                    stack.clear()
                else:
                    push(val)
                continue
            op = val
            try:
                if op == "Tj":
                    if stack and isinstance(stack[-1], (bytes, PdfString)):
                        show(stack[-1])
                elif op == "TJ":
                    if stack and isinstance(stack[-1], list):
                        for x in stack[-1]:
                            if isinstance(x, (bytes, PdfString)):
                                show(x)
                            elif isinstance(x, (int, float)):
                                adjust(x)
                elif op == "Td":
                    tx, ty = float(stack[-2]), float(stack[-1])
                    tlm = mul((1, 0, 0, 1, tx, ty), tlm)
                    tm = tlm
                elif op == "TD":
                    tx, ty = float(stack[-2]), float(stack[-1])
                    gs.tl = -ty
                    tlm = mul((1, 0, 0, 1, tx, ty), tlm)
                    tm = tlm
                elif op == "Tm":
                    tlm = tuple(float(x) for x in stack[-6:])
                    tm = tlm
                elif op == "T*":
                    tlm = mul((1, 0, 0, 1, 0, -gs.tl), tlm)
                    tm = tlm
                elif op == "'":
                    tlm = mul((1, 0, 0, 1, 0, -gs.tl), tlm)
                    tm = tlm
                    if stack and isinstance(stack[-1], (bytes, PdfString)):
                        show(stack[-1])
                elif op == '"':
                    gs.tw, gs.tc = float(stack[-3]), float(stack[-2])
                    tlm = mul((1, 0, 0, 1, 0, -gs.tl), tlm)
                    tm = tlm
                    if isinstance(stack[-1], (bytes, PdfString)):
                        show(stack[-1])
                elif op == "Tf":
                    name = stack[-2]
                    gs.fs = float(stack[-1])
                    gs.font = name.name if isinstance(name, Name) else ""
                    gs.fontobj = self._font(res, gs.font)
                elif op == "Tc":
                    gs.tc = float(stack[-1])
                elif op == "Tw":
                    gs.tw = float(stack[-1])
                elif op == "Tz":
                    gs.th = float(stack[-1]) / 100.0
                elif op == "TL":
                    gs.tl = float(stack[-1])
                elif op == "Tr":
                    gs.tr = int(stack[-1])
                elif op == "Ts":
                    gs.rise = float(stack[-1])
                elif op == "BT":
                    tm = tlm = IDENT
                    in_text = True
                elif op == "ET":
                    in_text = False
                elif op == "q":
                    gstack.append(gs.copy())
                elif op == "Q":
                    if gstack:
                        gs = gstack.pop()
                elif op == "cm":
                    m = tuple(float(x) for x in stack[-6:])
                    gs.ctm = mul(m, gs.ctm)
                elif op == "w":
                    gs.lw = float(stack[-1])
                elif op in ("g",):
                    gs.fill = _gray(float(stack[-1]))
                elif op == "G":
                    gs.stroke = _gray(float(stack[-1]))
                elif op == "rg":
                    gs.fill = tuple(float(x) for x in stack[-3:])
                elif op == "RG":
                    gs.stroke = tuple(float(x) for x in stack[-3:])
                elif op == "k":
                    gs.fill = _cmyk(*[float(x) for x in stack[-4:]])
                elif op == "K":
                    gs.stroke = _cmyk(*[float(x) for x in stack[-4:]])
                elif op in ("sc", "scn", "SC", "SCN"):
                    nums = [float(x) for x in stack if isinstance(x, (int, float))]
                    col = _gray(nums[0]) if len(nums) == 1 else tuple(nums[:3]) if len(nums) == 3 else _cmyk(*nums[:4]) if len(nums) == 4 else (0.0, 0.0, 0.0)
                    if op in ("sc", "scn"):
                        gs.fill = col
                    else:
                        gs.stroke = col
                elif op in ("cs", "CS"):
                    if op == "cs":
                        gs.fill = (0.0, 0.0, 0.0)
                    else:
                        gs.stroke = (0.0, 0.0, 0.0)
                elif op == "m":
                    x, y = apply(gs.ctm, float(stack[-2]), float(stack[-1]))
                    sub = new_sub(x, y)
                    cur = (x, y)
                    start = cur
                elif op == "l":
                    x, y = apply(gs.ctm, float(stack[-2]), float(stack[-1]))
                    if sub is None:
                        sub = new_sub(x, y)
                    else:
                        sub["lines"].append((cur[0], cur[1], x, y))
                        sub["pts"].append((x, y))
                    cur = (x, y)
                elif op in ("c", "v", "y"):
                    nums = [float(x) for x in stack[-(6 if op == "c" else 4):]]
                    pts = [apply(gs.ctm, nums[i], nums[i + 1]) for i in range(0, len(nums), 2)]
                    if sub is None:
                        sub = new_sub(*pts[-1])
                    sub["curve"] = True
                    sub["pts"].extend(pts)
                    cur = pts[-1]
                elif op == "h":
                    if sub is not None:
                        sub["closed"] = True
                        if start is not None:
                            cur = start
                elif op == "re":
                    x, y, w, h = [float(v) for v in stack[-4:]]
                    p0 = apply(gs.ctm, x, y)
                    p1 = apply(gs.ctm, x + w, y)
                    p2 = apply(gs.ctm, x + w, y + h)
                    p3 = apply(gs.ctm, x, y + h)
                    path.append({"pts": [p0, p1, p2, p3], "lines": [(p0[0], p0[1], p1[0], p1[1]), (p1[0], p1[1], p2[0], p2[1]),
                                                                  (p2[0], p2[1], p3[0], p3[1])], "closed": True, "re": True})
                    sub = None
                    cur = p0
                    start = p0
                elif op in ("S", "s"):
                    if op == "s" and sub is not None:
                        sub["closed"] = True
                    flush_path(True, False)
                    path.clear()
                    sub = None
                elif op in ("f", "F", "f*"):
                    flush_path(False, True)
                    path.clear()
                    sub = None
                elif op in ("B", "B*", "b", "b*"):
                    if op.startswith("b") and sub is not None:
                        sub["closed"] = True
                    flush_path(True, True)
                    path.clear()
                    sub = None
                elif op == "n":
                    path.clear()
                    sub = None
                elif op == "Do":
                    name = stack[-1].name if stack and isinstance(stack[-1], Name) else ""
                    xo = doc.resolve((doc.resolve(res.get("XObject")) or {}).get(name))
                    if isinstance(xo, Stream):
                        st = getattr(xo.dict.get("Subtype"), "name", "")
                        key = xo.num or id(xo)
                        if st == "Form" and key in drawing:
                            out.problems.append(f"フォームの部品 {name} が自分自身を描いています（その所は描きません）")
                        elif st == "Form":
                            m = xo.dict.get("Matrix")
                            m = tuple(float(doc.resolve(v)) for v in m) if isinstance(m, list) and len(m) == 6 else IDENT
                            sub_res = doc.resolve(xo.dict.get("Resources")) or res
                            g2 = gs.copy()
                            g2.ctm = mul(m, gs.ctm)
                            sps = doc.resolve(xo.dict.get("StructParents"))
                            sp = doc.resolve(xo.dict.get("StructParent"))
                            drawing.add(key)
                            try:
                                self._run(doc.decode(xo), sub_res, g2, depth + 1, outer + tuple(x[3] for x in mc_stack),
                                          sps if isinstance(sps, int) else None, sp if isinstance(sp, int) else objparent)
                            except (LimitExceeded, RecursionError, MemoryError):
                                raise
                            except Exception as e:  # noqa: BLE001
                                out.problems.append(f"フォームの部品 {name} を読めません: {e}")
                            finally:
                                drawing.discard(key)
                        elif st == "Image":
                            pts = [apply(gs.ctm, 0, 0), apply(gs.ctm, 1, 0), apply(gs.ctm, 0, 1), apply(gs.ctm, 1, 1)]
                            sp = doc.resolve(xo.dict.get("StructParent"))
                            pic = picture(pts, name, sp if isinstance(sp, int) else None)
                            pic.ctm, pic.image = gs.ctm, xo
                            if not self._thin_image(pic, xo.dict, None, gs, xo):
                                out.pictures.append(pic)
                elif op == "EI":
                    pts = [apply(gs.ctm, 0, 0), apply(gs.ctm, 1, 0), apply(gs.ctm, 0, 1), apply(gs.ctm, 1, 1)]
                    pic = picture(pts, "inline")
                    pic.ctm, pic.image = gs.ctm, inline_img
                    if not self._thin_image(pic, inline_img[0] if inline_img else {}, inline_img[1] if inline_img else b"", gs):
                        out.pictures.append(pic)
                    inline_img = None
                elif op == "gs":
                    name = stack[-1].name if stack and isinstance(stack[-1], Name) else ""
                    eg = doc.resolve((doc.resolve(res.get("ExtGState")) or {}).get(name))
                    if isinstance(eg, dict):
                        if "LW" in eg:
                            gs.lw = float(doc.resolve(eg["LW"]))
                        fnt = doc.resolve(eg.get("Font"))
                        if isinstance(fnt, list) and len(fnt) == 2:
                            gs.fontobj = load_font(doc, fnt[0], "gs")
                            gs.fs = float(doc.resolve(fnt[1]))
                elif op in ("BDC", "BMC"):
                    props = stack[-1] if op == "BDC" and stack else None
                    if isinstance(props, Name):
                        props = doc.resolve((doc.resolve(res.get("Properties")) or {}).get(props.name))
                    tag = stack[-2 if op == "BDC" else -1] if len(stack) >= (2 if op == "BDC" else 1) else None
                    tag = tag.name if isinstance(tag, Name) else ""
                    mcid = None
                    actual = None
                    alt = None
                    if isinstance(props, dict):
                        from .objects import decode_text

                        mc = props.get("MCID")
                        mcid = int(mc) if isinstance(mc, (int, float)) else None
                        at = props.get("ActualText")
                        if isinstance(at, (bytes, PdfString)):
                            actual = decode_text(at)
                        al = props.get("Alt")
                        alt = decode_text(al) if isinstance(al, (bytes, PdfString)) else None
                        alt = (alt or "").strip() or (actual or "").strip() or None
                    mark = _Mark(tag, mcid, owner, alt, tag == "Artifact")
                    if mcid is None and mc_stack:
                        mcid = mc_stack[-1][0]
                    mc_stack.append((mcid, actual, len(out.glyphs), mark))
                elif op == "EMC":
                    if mc_stack:
                        mcid, actual, first, _mark = mc_stack.pop()
                        if actual is not None:
                            gl = out.glyphs[first:]
                            if gl:
                                g0 = gl[0]
                                merged = Glyph(actual, min(g.x0 for g in gl), min(g.y0 for g in gl), max(g.x1 for g in gl),
                                               max(g.y1 for g in gl), g0.size, g0.ox, g0.oy, gl[-1].ox + gl[-1].dx - g0.ox,
                                               gl[-1].oy + gl[-1].dy - g0.oy, g0.vertical, g0.invisible, g0.font, g0.color,
                                               mcid, True, sum(g.adv for g in gl))
                                del out.glyphs[first:]
                                out.glyphs.append(merged)
            except LimitExceeded:
                raise
            except (IndexError, ValueError, TypeError, ZeroDivisionError):
                pass
            stack.clear()
        del in_text


_ABBR = {"AHx": "ASCIIHexDecode", "A85": "ASCII85Decode", "LZW": "LZWDecode", "Fl": "FlateDecode", "RL": "RunLengthDecode",
         "DCT": "DCTDecode", "CCF": "CCITTFaxDecode", "G": "DeviceGray", "RGB": "DeviceRGB", "CMYK": "DeviceCMYK",
         "W": "Width", "H": "Height", "BPC": "BitsPerComponent", "CS": "ColorSpace", "F": "Filter", "DP": "DecodeParms",
         "IM": "ImageMask", "D": "Decode"}


def _thin_image_impl(self, pic, d: dict, data, gs, xo=None) -> bool:
    """細い画像（罫線を画像で描いたもの）は、線として扱う。暗い画素の画像だけ。線にしたら True。"""
    w, h = pic.x1 - pic.x0, pic.y1 - pic.y0
    if not min(w, h) <= 2.5:
        return False
    dd = {_ABBR.get(k, k): v for k, v in d.items()}
    mask = dd.get("ImageMask") is True
    dark = True
    if mask:
        dark = max(gs.fill) < 0.5
    else:
        try:
            if xo is not None:
                raw = self.doc.decode(xo)
            else:
                from .objects import Stream as _S

                f = dd.get("Filter")
                if isinstance(f, Name):
                    f = Name(_ABBR.get(f.name, f.name))
                elif isinstance(f, list):
                    f = [Name(_ABBR.get(x.name, x.name)) if isinstance(x, Name) else x for x in f]
                raw = self.doc.decode(_S({"Filter": f, "DecodeParms": dd.get("DecodeParms")}, data), decrypt=False) if f else data
            cs = dd.get("ColorSpace")
            cs = _ABBR.get(cs.name, cs.name) if isinstance(cs, Name) else ""
            iw, ih = dd.get("Width"), dd.get("Height")
            bpc = dd.get("BitsPerComponent", 8)
            ncomp = {"DeviceRGB": 3, "DeviceGray": 1, "DeviceCMYK": 4}.get(cs)
            if ncomp and bpc == 8 and isinstance(iw, int) and isinstance(ih, int) and (ih == 1 or iw == 1) \
                    and len(raw) >= iw * ih * ncomp:
                # 1画素の幅の画像（図形を1行ずつ画像にしたもの）：暗い画素の続くところだけを線にする
                n = max(iw, ih)
                darks = []
                for k in range(n):
                    px = raw[k * ncomp:(k + 1) * ncomp]
                    v = sum(px) / ncomp
                    darks.append(v > 128 if cs == "DeviceCMYK" else v < 128)
                along_x = ih == 1
                span = (pic.x1 - pic.x0) if along_x else (pic.y1 - pic.y0)
                k = 0
                while k < n:
                    if not darks[k]:
                        k += 1
                        continue
                    j = k
                    while j < n and darks[j]:
                        j += 1
                    a, b = span * k / n, span * j / n
                    if along_x:
                        sub = Picture(pic.x0 + a, pic.y0, pic.x0 + b, pic.y1, pic.name)
                    else:
                        # 画像の行は上から（y が大きい方から）
                        sub = Picture(pic.x0, pic.y1 - b, pic.x1, pic.y1 - a, pic.name)
                    self._line_piece(sub)
                    k = j
                return True
            sample = raw[:4096]
            if sample:
                mean = sum(sample) / len(sample)
                dark = mean > 96 if cs == "DeviceCMYK" else mean < 160
        except Exception:  # noqa: BLE001
            dark = True
    if not dark:
        return True  # 白い細い画像は、見えない線（何もしない）
    self._line_piece(pic)
    return True


def _line_piece_impl(self, pic) -> None:
    w, h = pic.x1 - pic.x0, pic.y1 - pic.y0
    if max(w, h) < 3:
        # 点のような小さな画像（線を1画素ずつ描いたもの）：縦にも横にもつながりうる
        cx, cy = (pic.x0 + pic.x1) / 2, (pic.y0 + pic.y1) / 2
        self.out.segs.append(Seg(pic.x0, cy, pic.x1, cy, max(h, 0.1), True, (0.0, 0.0, 0.0)))
        self.out.segs.append(Seg(cx, pic.y0, cx, pic.y1, max(w, 0.1), True, (0.0, 0.0, 0.0)))
        return True
    if w >= h:
        self.out.segs.append(Seg(pic.x0, (pic.y0 + pic.y1) / 2, pic.x1, (pic.y0 + pic.y1) / 2, max(h, 0.1), True, (0.0, 0.0, 0.0)))
    else:
        self.out.segs.append(Seg((pic.x0 + pic.x1) / 2, pic.y0, (pic.x0 + pic.x1) / 2, pic.y1, max(w, 0.1), True, (0.0, 0.0, 0.0)))


Interpreter._thin_image = _thin_image_impl
Interpreter._line_piece = _line_piece_impl


def _as_rect(pts, closed) -> tuple | None:
    """4つの角が軸にそろった四角形なら (x0, y0, x1, y1)。"""
    p = list(pts)
    if len(p) == 5 and abs(p[0][0] - p[4][0]) < 0.01 and abs(p[0][1] - p[4][1]) < 0.01:
        p = p[:4]
    if len(p) != 4:
        return None
    xs = sorted({round(q[0], 2) for q in p})
    ys = sorted({round(q[1], 2) for q in p})
    if len(xs) > 2 or len(ys) > 2:
        # わずかなずれは許す
        xs2 = sorted(q[0] for q in p)
        ys2 = sorted(q[1] for q in p)
        if xs2[1] - xs2[0] > 0.3 or xs2[3] - xs2[2] > 0.3 or ys2[1] - ys2[0] > 0.3 or ys2[3] - ys2[2] > 0.3:
            return None
        return (xs2[0], ys2[0], xs2[3], ys2[3])
    x0, x1 = xs[0], xs[-1]
    y0, y1 = ys[0], ys[-1]
    for q in p:
        if not ((abs(q[0] - x0) < 0.3 or abs(q[0] - x1) < 0.3) and (abs(q[1] - y0) < 0.3 or abs(q[1] - y1) < 0.3)):
            return None
    return (min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))


def _rect_edges(out: PageContent, rect, lw, color) -> None:
    x0, y0, x1, y1 = rect
    out.segs += [Seg(x0, y0, x1, y0, lw, False, color), Seg(x0, y1, x1, y1, lw, False, color),
                 Seg(x0, y0, x0, y1, lw, False, color), Seg(x1, y0, x1, y1, lw, False, color)]


def page_content(doc, page: dict) -> PageContent:
    it = Interpreter(doc)
    pc = it.run_page(page)
    fonts = getattr(it, "_fonts_used", {})
    pc.unmapped = sum(1 for g in pc.glyphs if g.text == UNKNOWN)
    del fonts
    return pc
