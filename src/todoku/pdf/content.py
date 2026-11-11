"""ページの描画の命令（コンテンツストリーム）を順に読み、文字（位置つき）・線・四角形・画像を取り出す。

位置はページの既定の座標（左下が原点、単位はポイント）で表す。ページの回転は layout.py で扱う。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from .fonts import UNKNOWN, PdfFont, legacy_text, load_font
from .objects import Lexer, Name, PdfString, Stream

IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


def mul(m1, m2):
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (a1 * a2 + b1 * c2, a1 * b2 + b1 * d2, c1 * a2 + d1 * c2, c1 * b2 + d1 * d2,
            e1 * a2 + f1 * c2 + e2, e1 * b2 + f1 * d2 + f2)


def apply(m, x, y):
    return (m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5])


def _ink_all(gl) -> tuple | None:
    """まとめた字（ActualText）の輪郭の四角：まとめる前の字の四角を、1 つずつ並べたもの。1 つでも分からない字があれば None。"""
    if any(g.ink is None for g in gl):
        return None
    return tuple(b for g in gl for b in g.ink)


@dataclass(slots=True)  # 字はページに何万と並ぶ：1 字のメモリを小さく
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
    fontbase: str = ""  # 書体の名前（BaseFont）
    embedded: bool = False  # 書体のプログラムを埋め込んだ書体か
    mc: tuple | None = None  # 内側の番号つきの印：(持ち主, MCID)。持ち主は "page" か StructParents の番号（structadd）
    # 字形の輪郭の四角の並び（ページの座標の (x0, y0, x1, y1)。Interpreter.ink_of を渡したときだけ）：None は分からない、
    # () は墨のない字（空白など）。まとめた字（ActualText）は、まとめる前の字の四角を 1 つずつ
    ink: tuple | None = None

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
    part: bool = False  # 同じ道のほかの部分と重なる塗りの四角（枠の内と外・輪郭で描いた字の一部）：1 本の線として描いたものではない


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
    curve_paths: list[list] = field(default_factory=list)  # 曲線を含む図形（線・塗り）の点（制御点を含む。重なりの確かめ）
    # 曲線だけで閉じた図形（○・楕円。直線を含む角の丸い四角などは入れない）の外枠：(x0, y0, x1, y1, 線, 塗り)
    ovals: list[tuple] = field(default_factory=list)
    # 曲線と、長さのある直線とで閉じた図形（角の丸い四角など）の外枠：○かもしれない（字を囲む印にも枠にも使う）。窓で「？」にする
    maybe_ovals: list[tuple] = field(default_factory=list)
    unmapped: int = 0
    problems: list[str] = field(default_factory=list)
    # 既定の設定では画面に見えないレイヤー（optional content）の内容：[(レイヤーの見え方 layers.State, その内容 PageContent)]。
    # ふつうの内容（上の項目）には入れない：字は読み順に入れずに分けて伝え、線・画像は欄・画像にしない
    off: list[tuple] = field(default_factory=list)
    # 画面には見えるが、既定の設定では印刷に出ないレイヤーの字：[(レイヤーの見え方, 字)]。字は、ふつうの内容にも入っている
    screen_only: list[tuple] = field(default_factory=list)
    # 描かれているが、切り抜き（クリップ）の外にあって、画面にも印刷にも出ない字。ふつうの内容（glyphs）には入れない
    clipped: list[Glyph] = field(default_factory=list)
    # 同じく、切り抜きの外にあって見えない線・図形（道の部分の数）と画像の数。罫線・マス・画像の一覧には入れない
    clipped_shapes: int = 0
    clipped_pictures: int = 0


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
    __slots__ = ("ctm", "lw", "fill", "stroke", "font", "fs", "tc", "tw", "th", "tl", "tr", "rise", "fontobj", "clip", "ml")

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
        # 切り抜き（W・W*・部品の BBox）の道の外枠：(x0, y0, x1, y1)。切り抜きがなければ None。
        # 外枠は切り抜きの形を含むので、外枠の外にある字は、確かに見えない（外枠の中の字が見えるとは決めない）
        self.clip = None
        self.ml = 10.0  # 線の角のとがりの上限（M）：線の墨が、道からどこまで出うるか

    def copy(self):
        g = _GS.__new__(_GS)
        for k in _GS.__slots__:
            setattr(g, k, getattr(self, k))
        return g


_NO_AREA = (float("inf"), float("inf"), float("-inf"), float("-inf"))  # 何も通さない切り抜き


def _clip_and(cur, pts):
    """今の切り抜きの外枠と、道（その点の全部。曲線は制御点も：曲線は制御点の外枠から出ない）の外枠の重なり。"""
    x0, y0 = min(p[0] for p in pts), min(p[1] for p in pts)
    x1, y1 = max(p[0] for p in pts), max(p[1] for p in pts)
    if cur is not None:
        x0, y0, x1, y1 = max(x0, cur[0]), max(y0, cur[1]), min(x1, cur[2]), min(y1, cur[3])
    return (x0, y0, x1, y1) if x0 < x1 and y0 < y1 else _NO_AREA


def _cut_seg(ax, ay, bx, by, cl, half: float):
    """太さの半分が half の直線のうち、切り抜きの外枠 cl の中に墨がありうる部分（Liang–Barsky）。なければ None。
    線の向きには外枠のとおりに切り、線に直角な向きには太さの半分だけ外枠を広げる（外枠のすぐ外を通る線も、墨の半分は中に入る）。"""
    t0, t1 = 0.0, 1.0
    dx, dy = bx - ax, by - ay
    hx, hy = (0.0 if dy == 0 and dx != 0 else half), (0.0 if dx == 0 and dy != 0 else half)
    box = (cl[0] - hx, cl[1] - hy, cl[2] + hx, cl[3] + hy)
    for p_, q_ in ((-dx, ax - box[0]), (dx, box[2] - ax), (-dy, ay - box[1]), (dy, box[3] - ay)):
        if p_ == 0:
            if q_ < 0:
                return None
            continue
        r = q_ / p_
        if p_ < 0:
            t0 = max(t0, r)
        else:
            t1 = min(t1, r)
        if t0 > t1:
            return None
    if t0 == 0.0 and t1 == 1.0:
        return (ax, ay, bx, by)
    if t1 <= t0:
        return None  # 角に触れるだけ
    return (ax + t0 * dx, ay + t0 * dy, ax + t1 * dx, ay + t1 * dy)


def _open_at_page(cl, page):
    """線・塗りを切るための外枠：ページの枠の辺か、その外にある辺では切らない
    （ページの外へ出る線・塗りは、切り抜きがないときと同じに、出たままの形として扱う：ページの端で終わるマスの辺を失わない）。"""
    if cl is None or page is None:
        return cl
    inf = float("inf")
    r = (cl[0] if cl[0] > page[0] else -inf, cl[1] if cl[1] > page[1] else -inf,
         cl[2] if cl[2] < page[2] else inf, cl[3] if cl[3] < page[3] else inf)
    return None if r == (-inf, -inf, inf, inf) else r


def _outside(b, cl, m: float) -> bool:
    """四角 b が、余裕 m を置いても、切り抜きの外枠 cl の外にあるか。"""
    return b[2] + m <= cl[0] or b[0] - m >= cl[2] or b[3] + m <= cl[1] or b[1] - m >= cl[3]


def _gray(v):
    return (v, v, v)


def _cmyk(c, m, y, k):
    return ((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))


class Interpreter:
    def __init__(self, doc):
        self.doc = doc
        self.out = PageContent()
        self.written_forms: set[int] = set()
        from .layers import layers

        self.layers = layers(doc)
        self._main = self.out  # ふつうの内容（self.out は、画面に見えないレイヤーの中では、そのレイヤーの内容）
        self._layer = None  # いま読んでいる内容のレイヤーの見え方（重なったレイヤーは合わせたもの）。ふつうの内容は None
        # 字形の輪郭の四角を返す関数（placement が渡す）：(資源, 書体の名前, 書体, 符号, 字) → 字の原点からの (x0, y0, x1, y1)
        # （字の大きさを 1 とした長さ）。墨のない字は ()、分からなければ None
        self.ink_of = None
        self._clip_ink_fn = None
        self._page = None  # ページの枠（内容の座標）

    def run_page(self, page: dict) -> PageContent:
        res = self.doc.resolve(page.get("Resources")) or {}
        data = self.doc.page_content(page)
        from .pagedraw import verified_refs

        self.written_forms = verified_refs(self.doc, page)  # とどくが書いた値（欄の値として読む）：印刷された字として読まない
        from .pagetext import View

        gs = _GS()
        v = View(self.doc, page)
        gs.clip = (v.x0, v.y0, v.x1, v.y1)  # 見えるのは、ページの枠（CropBox を MediaBox に切り詰めたもの）の中だけ
        self._page = gs.clip
        self._run(data, res, gs, depth=0, owner="page")
        self.out.unmapped = sum(f.unmapped for f in self._fonts_used.values()) if hasattr(self, "_fonts_used") else 0
        for _st, c in self.out.off:
            self.out.problems += [x for x in c.problems if x not in self.out.problems]
        return self.out

    def run_stream(self, data: bytes, res: dict, ctm=IDENT) -> PageContent:
        gs = _GS()
        gs.ctm = ctm
        self._run(data, res or {}, gs, depth=0)
        return self.out

    def _clip_ink(self, res, name, pf, code, text):
        """切り抜きの外枠にかかる字の、字形の輪郭の四角（字の座標）。分からなければ None（見える字のままにする）。"""
        if self._clip_ink_fn is None:
            from .placement import outline_ink

            self._clip_ink_fn = outline_ink(self.doc)
        return self._clip_ink_fn(res, name, pf, code, text)

    # ------------------------------------------------------------------ レイヤー（optional content）
    def _layer_begin(self, oc):
        """内容の印（/OC）の始まり。既定の設定で、画面に見えない・印刷に出ない内容になるなら、戻すための控えを返す（変わらなければ None）。
        画面に見えない内容は、ふつうの内容と分けて集める（同じ見え方のレイヤーは、ページで 1 つにまとめる：行を組み立てるため）。"""
        st = self.layers.state(oc)
        if st is None:
            return None
        cur = self._layer
        eff = cur.both(st) if cur is not None else st
        if (eff.view, eff.print) == ((cur.view, cur.print) if cur is not None else (True, True)):
            return None
        saved = (cur, self.out, len(self.out.glyphs))
        self._layer = eff
        if not eff.view:
            got = next((c for s, c in self._main.off if s == eff), None)
            if got is None:
                got = PageContent()
                self._main.off.append((eff, got))
            self.out = got
        return saved

    def _layer_end(self, saved) -> None:
        cur, out, first = saved
        eff = self._layer
        if eff is not None and eff.view and not eff.print and self.out.glyphs[first:]:
            self._main.screen_only.append((eff, self.out.glyphs[first:]))
        self._layer, self.out = cur, out

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
        layered = self.layers.any  # レイヤーの設定のある書類だけ、内容の印（/OC）を見る
        stack: list = []
        gstack: list[_GS] = []
        path: list = []  # 部分の図形：[("m", x, y) ...] をページの座標で
        cur = None
        start = None
        tm = IDENT
        tlm = IDENT
        in_text = False
        side: list = [None]  # いまの文字の組（BT〜ET）で、すぐ前の字が見えない字か（None はまだ字がない）。後ろは、待たせている空白
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
            chain_ = outer + tuple(m[3] for m in mc_stack)
            mc_ = next(((m.owner, m.mcid) for m in reversed(chain_) if m.mcid is not None and m.owner is not None), None)
            ink_of = self.ink_of
            for code, text, w, is_space, vpos, vwidth in items:
                trm = mul((fs * th, 0, 0, fs, 0, gs.rise), mul(tm, gs.ctm))
                ink = None
                if ink_of is not None:
                    ink = ink_of(res, gs.font, f, code, text)
                    if ink:
                        gm = mul((1, 0, 0, 1, -vpos[0], -vpos[1]), trm) if f.vertical else trm
                        ps = [apply(gm, x, y) for x in (ink[0], ink[2]) for y in (ink[1], ink[3])]
                        ink = ((min(p[0] for p in ps), min(p[1] for p in ps), max(p[0] for p in ps), max(p[1] for p in ps)),)
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
                cl = gs.clip
                # 字の四角（送りの幅 × 書体の高さ）は墨の見積もり：字の大きさの半分の余裕を置いても外枠の外にある字だけを、見えない字にする
                hidden = False
                blank = False  # 外枠のすぐ外の、墨のない字（空白）：切り抜かれる墨がないので、同じ文字の組（BT〜ET）の字に合わせる
                if cl is not None and not (cl[0] <= min(xs) and max(xs) <= cl[2] and cl[1] <= min(ys) and max(ys) <= cl[3]):
                    # 字の四角（送りの幅 × 書体の高さ）は墨の見積もり：字の大きさの半分の余裕を置いても外枠の外にある字は、見えない。
                    # 外枠にかかる字・すぐ外の字は、字形の輪郭の四角（埋め込んだ書体）が外枠の外にあるときだけ、見えない字にする
                    box = (min(xs), min(ys), max(xs), max(ys))
                    hidden = _outside(box, cl, size / 2)
                    if not hidden:
                        emb = getattr(f, "embedded", False)  # 埋め込んでいない書体の字形は、見る道具が決める：分からない
                        ik = ink if emb else None
                        if ink_of is None and emb:
                            ik = self._clip_ink(res, gs.font, f, code, text)
                            if ik:
                                gm = mul((1, 0, 0, 1, -vpos[0], -vpos[1]), trm) if f.vertical else trm
                                ps = [apply(gm, x, y) for x in (ik[0], ik[2]) for y in (ik[1], ik[3])]
                                ik = ((min(p[0] for p in ps), min(p[1] for p in ps), max(p[0] for p in ps), max(p[1] for p in ps)),)
                        if ik:
                            hidden = all(_outside(b, cl, 0.0) for b in ik)
                        elif ik is not None:
                            blank = _outside(box, cl, 0.0)
                g = Glyph(text, min(xs), min(ys), max(xs), max(ys), size, ox, oy, nx - ox, ny - oy,
                          f.vertical, gs.tr in (3, 7), gs.font, gs.fill, mcid,
                          legacy=legacy_text(f.legacy, text),
                          adv=((fx - ox) ** 2 + (fy - oy) ** 2) ** 0.5,
                          fontbase=f.base, embedded=getattr(f, "embedded", False), mc=mc_, ink=ink)
                if blank:
                    # すぐ前の字と同じ側。前に字がなければ、見える側に置いておき、次の字が見えない字なら、そちらへ移す。
                    # 空白だけの組は、見える側のまま（前と同じ）
                    hidden = bool(side[0])
                    if side[0] is None:
                        side.append((out.glyphs, g))
                else:
                    if hidden and side[0] is None:
                        for lst, sp in side[1:]:
                            for i in range(len(lst) - 1, -1, -1):
                                if lst[i] is sp:
                                    del lst[i]
                                    self._main.clipped.append(sp)
                                    break
                    side[:] = [hidden]
                (self._main.clipped if hidden else out.glyphs).append(g)

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
            cl = _open_at_page(gs.clip, self._page)
            # 線の墨は、道から、太さの半分 × 角のとがりの上限（四角い端では √2）まで出うる
            reach = lw / 2 * max(gs.ml, 1.5) if stroke else 0.0
            bbs: list = []  # 道の部分ごとの外枠（要るときに 1 回だけ作る）

            def part(sub, rect) -> bool:
                """塗りの四角 rect が、同じ道のほかの部分と重なる（枠の内と外の四角・輪郭で描いた字の一部）。"""
                if len(path) < 2:
                    return False
                if not bbs:
                    for o in path:
                        ps = o["pts"]
                        bbs.append((min(p[0] for p in ps), min(p[1] for p in ps), max(p[0] for p in ps), max(p[1] for p in ps)))
                return any(o is not sub and bb[0] < rect[2] and bb[2] > rect[0] and bb[1] < rect[3] and bb[3] > rect[1]
                           for o, bb in zip(path, bbs))

            for sub in path:
                pts = sub["pts"]
                cut = False  # 外枠の外に出ている道：見えうるのは、外枠の中の部分だけ
                if cl is not None:
                    bb = (min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts))
                    if _outside(bb, cl, reach):
                        out.clipped_shapes += 1  # 切り抜きの外にあって見えない線・図形：罫線・マスにしない
                        continue
                    if bb[0] < cl[0] or bb[1] < cl[1] or bb[2] > cl[2] or bb[3] > cl[3]:
                        cut = True
                if sub.get("curve"):
                    xs = [p[0] for p in pts]
                    ys = [p[1] for p in pts]
                    out.curves.append((min(xs), min(ys), max(xs), max(ys), stroke, fill))
                    if (stroke or fill) and sub.get("poly") and len(out.curve_paths) < 20000:
                        out.curve_paths.append(sub["poly"])
                    w, h = max(xs) - min(xs), max(ys) - min(ys)
                    ends = abs(pts[0][0] - pts[-1][0]) + abs(pts[0][1] - pts[-1][1])
                    # 直線のうち、図形の外枠の対角線の 1% に足りないもの（長さ 0 の「x0 y0 l」で閉じる・曲線のつなぎ目の
                    # わずかな直線）は、形を変えないので数えない。作る道具が○の終わりに長さ 0 の直線を足すことがあり、
                    # それを直線と見ると、手で描いた○を読み落とす（1% は線の太さ・丸めの誤差より大きく、角の丸い四角の
                    # 辺よりずっと短い）
                    diag = (w * w + h * h) ** 0.5
                    lines = [s for s in sub.get("lines") or []
                             if ((s[2] - s[0]) ** 2 + (s[3] - s[1]) ** 2) ** 0.5 >= 0.01 * diag]
                    closed = sub.get("closed") or ends <= 0.02 * max(w, h, 1e-6)
                    if not lines and closed:
                        out.ovals.append((min(xs), min(ys), max(xs), max(ys), stroke, fill))  # 曲線だけで閉じた図形
                    elif closed:
                        # 曲線と長さのある直線とで閉じた図形（角の丸い四角など）：○かもしれない（落とさずに「？」で示す）
                        out.maybe_ovals.append((min(xs), min(ys), max(xs), max(ys), stroke, fill))
                    if not lines:
                        continue
                rect = _as_rect(pts, sub.get("closed"))
                if fill and rect is not None and cut:
                    # 塗りの四角は、外枠の中の部分だけにする。縁の線は、辺ごとに切る
                    x0, y0, x1, y1 = max(rect[0], cl[0]), max(rect[1], cl[1]), min(rect[2], cl[2]), min(rect[3], cl[3])
                    if stroke:
                        edges = PageContent()
                        _rect_edges(edges, rect, lw, gs.stroke)
                        for e in edges.segs:
                            c = _cut_seg(e.x0, e.y0, e.x1, e.y1, cl, lw / 2)
                            if c is not None:
                                out.segs.append(Seg(*c, lw, False, gs.stroke))
                    if x0 >= x1 or y0 >= y1:
                        continue
                    if max(x1 - x0, y1 - y0) <= 2.5:
                        pt = part(sub, rect)
                        out.segs.append(Seg(x0, (y0 + y1) / 2, x1, (y0 + y1) / 2, y1 - y0, True, gs.fill, pt))
                        out.segs.append(Seg((x0 + x1) / 2, y0, (x0 + x1) / 2, y1, x1 - x0, True, gs.fill, pt))
                    elif min(x1 - x0, y1 - y0) <= 2.5:
                        if x1 - x0 >= y1 - y0:
                            out.segs.append(Seg(x0, (y0 + y1) / 2, x1, (y0 + y1) / 2, y1 - y0, True, gs.fill, part(sub, rect)))
                        else:
                            out.segs.append(Seg((x0 + x1) / 2, y0, (x0 + x1) / 2, y1, x1 - x0, True, gs.fill, part(sub, rect)))
                    else:
                        out.boxes.append(Box(x0, y0, x1, y1, gs.fill, False))
                    continue
                if fill and rect is not None:
                    x0, y0, x1, y1 = rect
                    if max(x1 - x0, y1 - y0) <= 2.5:
                        # 点（点線の1つぶ）：縦にも横にもつながりうるので、両方の向きの短い線として残す
                        pt = part(sub, rect)
                        out.segs.append(Seg(x0, (y0 + y1) / 2, x1, (y0 + y1) / 2, y1 - y0, True, gs.fill, pt))
                        out.segs.append(Seg((x0 + x1) / 2, y0, (x0 + x1) / 2, y1, x1 - x0, True, gs.fill, pt))
                        continue
                    if min(x1 - x0, y1 - y0) <= 2.5:
                        # 細い塗りつぶし：線として扱う
                        if x1 - x0 >= y1 - y0:
                            out.segs.append(Seg(x0, (y0 + y1) / 2, x1, (y0 + y1) / 2, y1 - y0, True, gs.fill, part(sub, rect)))
                        else:
                            out.segs.append(Seg((x0 + x1) / 2, y0, (x0 + x1) / 2, y1, x1 - x0, True, gs.fill, part(sub, rect)))
                        continue
                    out.boxes.append(Box(x0, y0, x1, y1, gs.fill, stroke))
                    if stroke:
                        _rect_edges(out, rect, lw, gs.stroke)
                    continue
                if stroke:
                    segs = list(sub.get("lines") or [])
                    if sub.get("closed") and len(pts) >= 2 and pts[-1] != pts[0]:
                        segs.append((*pts[-1], *pts[0]))
                    for sg in segs:
                        if cut:
                            sg = _cut_seg(*sg, cl, lw / 2)
                            if sg is None:
                                out.clipped_shapes += 1
                                continue
                        out.segs.append(Seg(*sg, lw, False, gs.stroke))
                    if rect is not None and not fill and not cut:
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

        def end_path():
            nonlocal clip_next
            if clip_next:
                pts = [p for sp in path for p in sp["pts"]]
                if pts:  # 点のない道の切り抜きは、決めない（切り抜きを変えない）
                    gs.clip = _clip_and(gs.clip, pts)
                clip_next = False
            path.clear()

        def new_sub(x, y):
            sub = {"pts": [(x, y)], "lines": [], "closed": False}
            path.append(sub)
            return sub

        sub = None
        clip_next = False
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
                    side[:] = [None]
                elif op == "ET":
                    in_text = False
                    side[:] = [None]
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
                elif op == "M":
                    gs.ml = float(stack[-1])
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
                    if cur is not None:
                        # 曲線の形（重なりの確かめ）：3 次ベジエ曲線を折れ線にする。v は初めの制御点が今の点、y は後の制御点が終わりの点
                        c1, c2, p3 = ((pts[0], pts[1], pts[2]) if op == "c" else (cur, pts[0], pts[1]) if op == "v"
                                      else (pts[0], pts[1], pts[1]))
                        poly = sub.setdefault("poly", [cur])
                        if poly[-1] != cur:
                            poly.append(cur)  # 曲線のあいだの直線
                        for k in range(1, 13):
                            t = k / 12
                            u = 1 - t
                            poly.append((u ** 3 * cur[0] + 3 * u * u * t * c1[0] + 3 * u * t * t * c2[0] + t ** 3 * p3[0],
                                         u ** 3 * cur[1] + 3 * u * u * t * c1[1] + 3 * u * t * t * c2[1] + t ** 3 * p3[1]))
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
                elif op in ("W", "W*"):
                    clip_next = True  # 切り抜きは、次に道を描く（か捨てる）命令のあとで効く
                elif op in ("S", "s"):
                    if op == "s" and sub is not None:
                        sub["closed"] = True
                    flush_path(True, False)
                    end_path()
                    sub = None
                elif op in ("f", "F", "f*"):
                    flush_path(False, True)
                    end_path()
                    sub = None
                elif op in ("B", "B*", "b", "b*"):
                    if op.startswith("b") and sub is not None:
                        sub["closed"] = True
                    flush_path(True, True)
                    end_path()
                    sub = None
                elif op == "n":
                    end_path()
                    sub = None
                elif op == "Do":
                    name = stack[-1].name if stack and isinstance(stack[-1], Name) else ""
                    xo = doc.resolve((doc.resolve(res.get("XObject")) or {}).get(name))
                    if isinstance(xo, Stream):
                        # 部品の /OC：画面に見えないレイヤーの部品は、ふつうの内容に入れない
                        began = self._layer_begin(xo.dict.get("OC")) if layered else None
                        out = self.out
                        try:
                            st = getattr(xo.dict.get("Subtype"), "name", "")
                            key = xo.num or id(xo)
                            if st == "Form" and xo.num in self.written_forms:
                                pass
                            elif st == "Form" and key in drawing:
                                out.problems.append(f"フォームの部品 {name} が自分自身を描いています（その所は描きません）")
                            elif st == "Form":
                                m = xo.dict.get("Matrix")
                                m = tuple(float(doc.resolve(v)) for v in m) if isinstance(m, list) and len(m) == 6 else IDENT
                                sub_res = doc.resolve(xo.dict.get("Resources")) or res
                                g2 = gs.copy()
                                g2.ctm = mul(m, gs.ctm)
                                bb = doc.resolve(xo.dict.get("BBox"))
                                if isinstance(bb, list) and len(bb) == 4:  # 部品の内容は、BBox で切り抜かれる（8.10.2）
                                    bx0, by0, bx1, by1 = (float(doc.resolve(v)) for v in bb)
                                    g2.clip = _clip_and(gs.clip, [apply(g2.ctm, bx0, by0), apply(g2.ctm, bx1, by0),
                                                                  apply(g2.ctm, bx0, by1), apply(g2.ctm, bx1, by1)])
                                sps = doc.resolve(xo.dict.get("StructParents"))
                                sp = doc.resolve(xo.dict.get("StructParent"))
                                drawing.add(key)
                                keep = (self._layer, self.out)
                                try:
                                    self._run(doc.decode(xo), sub_res, g2, depth + 1, outer + tuple(x[3] for x in mc_stack),
                                              sps if isinstance(sps, int) else None, sp if isinstance(sp, int) else objparent)
                                except (LimitExceeded, RecursionError, MemoryError):
                                    raise
                                except Exception as e:  # noqa: BLE001
                                    out.problems.append(f"フォームの部品 {name} を読めません: {e}")
                                finally:
                                    drawing.discard(key)
                                    self._layer, self.out = keep  # 途中で読めなくなった部品の中のレイヤーを、外に持ち出さない
                            elif st == "Image":
                                pts = [apply(gs.ctm, 0, 0), apply(gs.ctm, 1, 0), apply(gs.ctm, 0, 1), apply(gs.ctm, 1, 1)]
                                sp = doc.resolve(xo.dict.get("StructParent"))
                                pic = picture(pts, name, sp if isinstance(sp, int) else None)
                                pic.ctm, pic.image = gs.ctm, xo
                                if gs.clip is not None and _outside((pic.x0, pic.y0, pic.x1, pic.y1), gs.clip, 0.0):
                                    out.clipped_pictures += 1  # 切り抜きの外にあって見えない画像
                                elif not self._thin_image(pic, xo.dict, None, gs, xo):
                                    out.pictures.append(pic)
                        finally:
                            if began is not None:
                                self._layer_end(began)
                            out = self.out
                elif op == "EI":
                    pts = [apply(gs.ctm, 0, 0), apply(gs.ctm, 1, 0), apply(gs.ctm, 0, 1), apply(gs.ctm, 1, 1)]
                    pic = picture(pts, "inline")
                    pic.ctm, pic.image = gs.ctm, inline_img
                    if gs.clip is not None and _outside((pic.x0, pic.y0, pic.x1, pic.y1), gs.clip, 0.0):
                        out.clipped_pictures += 1
                    elif not self._thin_image(pic, inline_img[0] if inline_img else {}, inline_img[1] if inline_img else b"", gs):
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
                    oc = None  # 内容の印が指すレイヤー（/OC /名前 BDC）
                    if isinstance(props, Name):
                        oc = (doc.resolve(res.get("Properties")) or {}).get(props.name)
                        props = doc.resolve(oc)
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
                    began = self._layer_begin(oc) if layered and tag == "OC" and oc is not None else None
                    out = self.out
                    mc_stack.append((mcid, actual, len(out.glyphs), mark, began))
                elif op == "EMC":
                    if mc_stack:
                        mcid, actual, first, _mark, began = mc_stack.pop()
                        if began is not None:
                            self._layer_end(began)
                            out = self.out
                        elif actual is not None:
                            gl = out.glyphs[first:]
                            if gl:
                                g0 = gl[0]
                                merged = Glyph(actual, min(g.x0 for g in gl), min(g.y0 for g in gl), max(g.x1 for g in gl),
                                               max(g.y1 for g in gl), g0.size, g0.ox, g0.oy, gl[-1].ox + gl[-1].dx - g0.ox,
                                               gl[-1].oy + gl[-1].dy - g0.oy, g0.vertical, g0.invisible, g0.font, g0.color,
                                               mcid=mcid, actual=True, legacy=any(g.legacy for g in gl), adv=sum(g.adv for g in gl),
                                               fontbase=g0.fontbase, embedded=g0.embedded, mc=g0.mc, ink=_ink_all(gl))
                                del out.glyphs[first:]
                                out.glyphs.append(merged)
            except LimitExceeded:
                raise
            except (IndexError, ValueError, TypeError, ZeroDivisionError):
                pass
            stack.clear()
        for m in reversed(mc_stack):  # 閉じていない内容の印は、この流れの終わりまで（次の内容まで隠さない。描く側と同じ）
            if m[4] is not None:
                self._layer_end(m[4])
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
    pc.unmapped = sum(1 for c in [pc] + [c for _st, c in pc.off] for g in c.glyphs if g.text == UNKNOWN)
    pc.unmapped += sum(1 for g in pc.clipped if g.text == UNKNOWN)
    del fonts
    return pc
