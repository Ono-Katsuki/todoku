"""PDF のページの字と線を、見た目の座標の字・罫線・マス・行にする（layout の 1〜3 の段階。決定的な処理だけで）。

- 座標：ページの回転・UserUnit を反映した、左上が原点・下向きが y の見た目の座標（View）
- 字：描かれた字形の位置と大きさ（Ch）。同じ位置に重ねて描いた同じ字（重ね書き）は 1 字にする（is_overprint）
- 罫線とマス：線を罫線に（rulings）、罫線の交わりから最小の四角のマスに（cells_from）
- 行：字を行に並べ、字間の空きは測った幅から空白として残す（make_lines）。縦書き・右から左・字の組み立ての順も
窓・書く所を見つける処理（layout.PageLayout）は、ここで作った字・マス・行を使う。
"""

from __future__ import annotations

import bisect
import math
import unicodedata
from dataclasses import dataclass, field

from .. import chars as _chars  # 字の種類（空白・下線の字・ダッシュ・点線・かっこの組・チェックの箱の字）：Unicode の資料の表
from ..limits import budget
from .content import PageContent
from .textcolor import hexcolor


# 印のない箱の字（chars.EMPTY_BOXES）：どれも箇条・見出しの印のこともある字で、チェックの印の所にはするが、規則では決めない（？）
BOX_GLYPHS = set(_chars.EMPTY_BOXES)

# ====================================================================== 座標
MM_PT = 72 / 25.4  # 1 mm のポイント（1 pt = 1/72 インチ）


def user_unit(doc, page: dict) -> float:
    """ページの UserUnit（ページの座標の 1 単位が何ポイントか。省いたときは 1。ISO 32000-1 14.11.1 表 30、PDF 1.6〜）。
    正の数でなければ 1 とみなす。"""
    v = doc.resolve(page.get("UserUnit")) if isinstance(page, dict) else None
    if isinstance(v, bool) or not isinstance(v, (int, float)) or not v > 0 or v != v or v == float("inf"):
        return 1.0
    return float(v)


def _box_numbers(doc, box) -> list[float] | None:
    """[左 下 右 上] の 4 つの数（壊れていて数でなければ None：ページの大きさは、次の候補か、既定の大きさ）。"""
    box = doc.resolve(box)
    if not isinstance(box, list) or len(box) != 4:
        return None
    vals = [doc.resolve(v) for v in box]
    return [float(v) for v in vals] if all(isinstance(v, (int, float)) for v in vals) else None


class View:
    """ページの座標 → 見た目の座標（左上が原点、y は下向き、回転を反映）。"""

    def __init__(self, doc, page: dict):
        box = _box_numbers(doc, page.get("CropBox")) or _box_numbers(doc, page.get("MediaBox")) or [0, 0, 595, 842]
        mb = page.get("MediaBox")
        if mb is not None:
            mb = _box_numbers(doc, mb)
            if mb:
                # CropBox は MediaBox の内側に切り詰める
                box = [max(box[0], min(mb[0], mb[2])), max(box[1], min(mb[1], mb[3])),
                       min(box[2], max(mb[0], mb[2])), min(box[3], max(mb[1], mb[3]))]
        self.x0, self.y0 = min(box[0], box[2]), min(box[1], box[3])
        self.x1, self.y1 = max(box[0], box[2]), max(box[1], box[3])
        rot = doc.resolve(page.get("Rotate"))
        rot = int(rot) % 360 if isinstance(rot, (int, float)) else 0
        self.rot = rot - rot % 90
        w, h = self.x1 - self.x0, self.y1 - self.y0
        self.width, self.height = (h, w) if self.rot in (90, 270) else (w, h)
        self.unit = user_unit(doc, page)

    @property
    def mm(self) -> float:
        """1 mm の長さ（ページの座標の単位）。ふつうは 72/25.4（1 単位 = 1/72 インチ）。UserUnit のあるページは、1 単位が
        UserUnit/72 インチなので、その分だけ短くする（ISO 32000-1 14.11.1 表 30）。"""
        return MM_PT / self.unit

    def pt(self, x: float, y: float) -> tuple[float, float]:
        r = self.rot
        if r == 0:
            return x - self.x0, self.y1 - y
        if r == 90:
            return y - self.y0, x - self.x0
        if r == 180:
            return self.x1 - x, y - self.y0
        return self.y1 - y, self.x1 - x

    def rect(self, x0, y0, x1, y1) -> tuple[float, float, float, float]:
        a = self.pt(x0, y0)
        b = self.pt(x1, y1)
        return min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1])

    def back(self, vx: float, vy: float) -> tuple[float, float]:
        """見た目の座標 → ページの座標。"""
        r = self.rot
        if r == 0:
            return vx + self.x0, self.y1 - vy
        if r == 90:
            return vy + self.x0, vx + self.y0
        if r == 180:
            return self.x1 - vx, vy + self.y0
        return self.x1 - vy, self.y1 - vx

    def back_rect(self, x0, y0, x1, y1) -> tuple[float, float, float, float]:
        a = self.back(x0, y0)
        b = self.back(x1, y1)
        return min(a[0], b[0]), min(a[1], b[1]), max(a[0], b[0]), max(a[1], b[1])


@dataclass(slots=True)  # 字はページに何万と並ぶ：1 字のメモリを小さく
class Ch:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    size: float
    vertical: bool  # 上から下へ読む（縦書き、または90度回した字）
    base: float  # 行をそろえる位置（横書きは原点の y、縦書きは字の中心の x）
    invisible: bool
    legacy: bool = False  # Unicode でない符号の書体の字
    seq: int = -1  # 描いた順（ページの描画の命令の中の何番目の字か）
    upward: bool = False  # 下から上へ進む字（左へ90度回した字。余白に回して印刷した字など）。行としては縦で、読む順は下から
    color: str = "000000"  # 字を描いた色 "RRGGBB"（模様で描いた字は textcolor.PATTERN）：pdf/textcolor.py
    faint: str = ""  # 描いても見えない字の理由（content.Glyph.faint）："same" は下の色と同じ色、"alpha" は不透明度 0

    @property
    def cx(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def cy(self) -> float:
        return (self.y0 + self.y1) / 2


def _is_cjk(ch: str) -> bool:
    """和文・中文のように、語の間を空けずに書く全角の字か（Unicode の東アジアの幅が W・F のもの）。
    ハングルは全角でも語の間を空白で区切るので含めない。アラビア文字の表示形など、幅の狭い字も含めない。"""
    if not ch:
        return False
    o = ord(ch[0])
    if o < 0x1100 or _is_hangul(o):
        return False
    if 0xFF61 <= o <= 0xFFDC or 0xE000 <= o <= 0xF8FF:
        return True  # 半角のかな、外字（和文の書類の外字は全角の字）
    return unicodedata.east_asian_width(ch[0]) in ("W", "F")


def _ruby_base(ch: str) -> bool:
    """ルビ（振り仮名・注音）のつく字：漢字（Unicode の名前が CJK UNIFIED / COMPATIBILITY IDEOGRAPH のもの）と、外字。"""
    if 0xE000 <= ord(ch) <= 0xF8FF:
        return True
    return unicodedata.name(ch, "").startswith(("CJK UNIFIED IDEOGRAPH", "CJK COMPATIBILITY IDEOGRAPH"))


def _is_hangul(o: int) -> bool:
    return 0x1100 <= o <= 0x11FF or 0x3130 <= o <= 0x318F or 0xA960 <= o <= 0xA97F or 0xAC00 <= o <= 0xD7FF


def _is_pua(ch: str) -> bool:
    """私用の符号（外字）：U+E000〜F8FF、第15・16面。"""
    return any(0xE000 <= ord(c) <= 0xF8FF or ord(c) >= 0xF0000 for c in ch)


# ====================================================================== 文字
def _box_overlap(a: "Ch", b: "Ch") -> float:
    """2つの字の箱の重なり：送りの向きの重なり（幅の狭いほうに対して）と、横切る向きの重なり（低いほうに対して）の小さいほう。"""
    def frac(a0, a1, b0, b1):
        inter = min(a1, b1) - max(a0, b0)
        return inter / max(1e-6, min(a1 - a0, b1 - b0)) if inter > 0 else 0.0
    along_x = frac(a.x0, a.x1, b.x0, b.x1)
    along_y = frac(a.y0, a.y1, b.y0, b.y1)
    return min(along_x, along_y)


def glyph_frame(g, view: "View | None" = None) -> tuple[float, float, float, float, float]:
    """字の原点 (x, y)・送りの向き (ux, uy)（長さ 1）・フォントの決める送りの幅。view があれば見た目の座標で。
    送りの幅が 0 の字（結合の記号など）は、字の箱の幅にする（content.py が 0.5 字ぶんの箱を置く）。"""
    if view is not None:
        ox, oy = view.pt(g.ox, g.oy)
        ex, ey = view.pt(g.ox + g.dx, g.oy + g.dy)
    else:
        ox, oy, ex, ey = g.ox, g.oy, g.ox + g.dx, g.oy + g.dy
    dx, dy = ex - ox, ey - oy
    n = (dx * dx + dy * dy) ** 0.5
    if n > 1e-9:
        ux, uy = dx / n, dy / n
    else:
        ux, uy = (0.0, 1.0) if g.vertical else (1.0, 0.0)
    w = g.adv
    if not w > 1e-9:
        w = max(abs((g.x1 - g.x0) * ux) + abs((g.y1 - g.y0) * uy), 1e-6)
    return ox, oy, ux, uy, w


def is_overprint(a: tuple, b: tuple, size_a: float, size_b: float) -> bool:
    """同じ字の重ね書き（太字に見せる重ね書き・色の層）か。a・b は glyph_frame。
    重ね書き＝同じ字を、ほぼ同じ所に描いたもの：2つの字の送りの幅（フォントが決める、その字の幅）が、狭い方の半分より多く
    重なる。続けて並べた字は、後の字が前の字の送りの終わりから始まり、重なりがない（か、詰めの分だけ）ので、重ね書きに
    ならない（幅の狭い「l」が続く「All」も、2つの字）。送りの向きに直角なずれは、字の大きさの 1/4 まで。"""
    ox, oy, ux, uy, w = a
    px, py = b[0] - ox, b[1] - oy
    along = px * ux + py * uy
    across = -px * uy + py * ux
    overlap = min(w, along + b[4]) - max(0.0, along)
    return (overlap > 0.5 * min(w, b[4]) and abs(across) < 0.25 * max(size_a, size_b)
            and abs(size_a - size_b) < 0.2 * max(size_a, size_b))


def drop_redrawn(hidden: list, glyphs: list, dropped: list | None = None) -> list:
    """塗りの下に隠れた字（hidden）のうち、同じ字を同じ所に、見える字としてもう一度描いてあるもの（塗りで消して描き直した
    ページの頭・足の字など）を除いて返す：重ね書き（is_overprint）と同じ決まりで、見える字の方を残し、隠れた方は 1 字に
    まとめたもの（"dup"）として数える。0.5pt に満たない字は、ここでは決めない（view_chars が "size" として落とす）。"""
    texts = {g.text for g in hidden if g.text and g.size >= 0.5}
    def frame(g):
        fr = glyph_frame(g)
        return fr if math.isfinite(fr[0] + fr[1] + fr[4]) else None  # 位置の決まらない字は、描き直しとは決めない

    cand = [(g, fr) for g in glyphs if g.text in texts and not g.invisible and g.size >= 0.5 and (fr := frame(g))]
    if not cand:
        return list(hidden)
    hid = [(g, frame(g) if g.text in texts and g.size >= 0.5 else None) for g in hidden]
    # 升目：重ね書きになる 2 つの字の原点は「広い方の送りの幅 + 大きさの 1/4」より離れないので、隣の升目までを見れば足りる
    cell = max(8.0, max(fr[4] for _g, fr in cand + [h for h in hid if h[1]]) + 0.25 * max(g.size for g, fr in cand + [h for h in hid if h[1]]))
    seen: dict[tuple, list] = {}
    for g, fr in cand:
        seen.setdefault((g.text, int(fr[0] // cell), int(fr[1] // cell)), []).append((fr, g.size))
    out = []
    for g, fr in hid:
        if fr is not None:
            kx, ky = int(fr[0] // cell), int(fr[1] // cell)
            if any(is_overprint(fr, ofr, g.size, osize) for ix in (kx - 1, kx, kx + 1) for iy in (ky - 1, ky, ky + 1)
                   for ofr, osize in seen.get((g.text, ix, iy), ())):
                if dropped is not None:
                    dropped.append(("dup", g.text))
                continue
        out.append(g)
    return out


def unseen_runs(glyphs: list, view: View) -> list[tuple[str, float, float, float]]:
    """読み順に入れない大きさの字（view_chars が外した字）を、ひと続きに描いた字ごとにまとめる：(字, 左, 上, 字の大きさ)。
    ひと続き：同じ文字の組（BT〜ET）の、同じ大きさの字で、前の字の原点から送りの向きに進んだ所の字（送りと直角に字の半分より
    多くずれる・後ろへ戻るなら、別のひと続き。送りのない字（大きさ 0 など）は、同じ所に重なるので、同じひと続き）。
    字は描いた順のまま（位置で空けた語の間は、空白にしない）。墨のある字のないひと続きは出さない。基線の上から、左からの順。"""
    runs: list[list] = []
    prev = None
    for g in glyphs:
        new = prev is None or g.bt != prev.bt or abs(g.size - prev.size) > 0.005
        if not new:
            n = math.hypot(prev.dx, prev.dy)
            if n > 0:
                vx, vy, ux, uy, half = g.ox - prev.ox, g.oy - prev.oy, prev.dx / n, prev.dy / n, prev.size / 2
                new = abs(vx * uy - vy * ux) > half or vx * ux + vy * uy < -half
        if new:
            runs.append([])
        runs[-1].append(g)
        prev = g
    out = []
    for run in runs:
        text = "".join(g.text for g in run).strip(" \u3000")
        if not text.strip():
            continue
        rs = [view.rect(g.x0, g.y0, g.x1, g.y1) for g in run if g.text.strip()]
        out.append((round(view.pt(run[0].ox, run[0].oy)[1], 2), text, min(r[0] for r in rs), min(r[1] for r in rs), run[0].size))
    out.sort(key=lambda r: (r[0], r[2]))
    return [r[1:] for r in out]


def view_chars(pc: PageContent, view: View, dropped: list | None = None, merges: list | None = None,
               unseen: list | None = None) -> list[Ch]:
    """unseen：読み順に入れない大きさの字（0.5 pt に満たない・ページより何倍も大きい）の、元の字（Glyph）を足していく一覧。"""
    out: list[Ch] = []
    seen: dict[tuple, list[tuple]] = {}
    cell = 8.0  # 探す範囲の区切り（重ね書きかは is_overprint が字の幅で決める。区切りは探す手間を減らすだけ）
    for seq, g in enumerate(pc.glyphs):
        if not g.text:
            continue
        x0, y0, x1, y1 = view.rect(g.x0, g.y0, g.x1, g.y1)
        fr = glyph_frame(g, view)
        ox, oy, ux, uy, _w = fr
        vertical = abs(uy) > abs(ux) * 1.5
        base = (x0 + x1) / 2 if vertical else oy
        if g.size < 0.5 or x1 - x0 > view.width * 2 or y1 - y0 > view.height * 2:
            if dropped is not None:
                dropped.append(("size", g.text))
            if unseen is not None:
                unseen.append(g)
            continue
        c = Ch(g.text, x0, y0, x1, y1, max(g.size, 1.0), vertical, base, g.invisible, g.legacy, seq, upward=vertical and uy < 0,
               color=hexcolor(g.color), faint=g.faint)
        # 同じ字を同じ位置に重ねたもの（太字に見せる重ね書き、色の層）は1字にする
        kx, ky = int(ox // cell), int(oy // cell)
        dup = None  # 重ね書きとした前の字
        for ix in (kx - 1, kx, kx + 1):
            for iy in (ky - 1, ky, ky + 1):
                for ofr, osize, och in seen.get((g.text, ix, iy), ()):
                    if is_overprint(ofr, fr, osize, g.size):
                        dup = och
                        break
                if dup:
                    break
            if dup:
                break
        if dup is not None:
            if not c.invisible:
                dup.color = c.color  # 重ねて描いた字は、後から描いた字（上の字）の色
                if dup.faint:
                    # 描いても見えない字（下の色と同じ色・不透明度 0）に、同じ字を見える描き方で重ねてある（白く塗った字に
                    # 黒い縁を重ねた字など）：見える字
                    dup.invisible, dup.faint = False, ""
            if dropped is not None:
                dropped.append(("dup", g.text))
            if merges is not None:
                # 字の箱（描いた字の形の範囲）の重なり：送りの幅（is_overprint）とは別の量。重ならない箱を1字にしたら、
                # 読む段階で知らせる（_read_flags の「重ね書きとして1字にした字」）
                merges.append((g.text, _box_overlap(dup, c), (c.x0, c.y0, c.x1, c.y1)))
            continue
        seen.setdefault((g.text, kx, ky), []).append((fr, g.size, c))
        out.append(c)
    _vertical_runs(out)
    return out


def _vertical_runs(chars: list[Ch]) -> None:
    """横書きの書体の字を、1字ずつ下へ積んで描いた縦書き（LibreOffice・Word の縦書きを PDF にしたもの）：
    描いた順に続く字が、同じ横の位置で、1字ぶん（字の大きさの 0.8〜1.5 倍）ずつ下へ進むものを、縦の行の字にする。
    3字以上続くものだけ（表の縦の見出しも同じ）。チェックの印（□）や、空白は含めない。"""
    run: list[Ch] = []

    def flush():
        if len(run) >= 3:
            for c in run:
                c.vertical = True
                c.base = c.cx
        run.clear()

    for c in chars:
        if c.vertical or c.text.isspace() or c.text in BOX_GLYPHS or len(c.text) > 2:
            flush()
            continue
        if run:
            p = run[-1]
            sz = max(p.size, c.size)
            if abs(c.cx - p.cx) <= 0.15 * sz and 0.8 * sz <= c.cy - p.cy <= 1.5 * sz and abs(c.size - p.size) < 0.2 * sz:
                run.append(c)
                continue
            flush()
        run.append(c)
    flush()


# ====================================================================== 罫線とマス
@dataclass
class Rule:
    pos: float  # 横線は y、縦線は x
    a: float  # 始まり（横線は x0、縦線は y0）
    b: float
    width: float = 0.5


def rulings(pc: PageContent, view: View) -> tuple[list[Rule], list[Rule]]:
    hs: list[Rule] = []
    vs: list[Rule] = []
    pw, ph = view.width, view.height
    for s in pc.segs:
        if all(c > 0.93 for c in s.color[:3]):
            continue  # 白い線は見えない
        x0, y0 = view.pt(s.x0, s.y0)
        x1, y1 = view.pt(s.x1, s.y1)
        if abs(y1 - y0) <= 0.8 and abs(x1 - x0) > 0.8:
            hs.append(Rule((y0 + y1) / 2, min(x0, x1), max(x0, x1), s.width))
        elif abs(x1 - x0) <= 0.8 and abs(y1 - y0) > 0.8:
            vs.append(Rule((x0 + x1) / 2, min(y0, y1), max(y0, y1), s.width))
        elif s.filled and abs(x1 - x0) <= 0.8 and abs(y1 - y0) <= 0.8:
            # 点線の小さな点：つながると線になる
            if abs(y1 - y0) < 1e-6:
                hs.append(Rule(y0, min(x0, x1), max(x0, x1), s.width))
            else:
                vs.append(Rule(x0, min(y0, y1), max(y0, y1), s.width))
    for b in pc.boxes:
        # 色の付いた四角（罫線のない表の、塗りつぶしたマス）の縁も、マスの境目として使う
        if b.fill is None or all(c > 0.93 for c in b.fill[:3]):
            continue
        x0, y0, x1, y1 = view.rect(b.x0, b.y0, b.x1, b.y1)
        if (x1 - x0) * (y1 - y0) > 0.6 * pw * ph:
            continue
        hs += [Rule(y0, x0, x1, 0.1), Rule(y1, x0, x1, 0.1)]
        vs += [Rule(x0, y0, y1, 0.1), Rule(x1, y0, y1, 0.1)]
    return _merge(hs), _merge(vs)


def _merge(rs: list[Rule], tol: float = 1.0, gap: float = 3.0) -> list[Rule]:
    """同じ高さ（位置）で重なる・ほぼつながる線を1本にする（破線もつなぐ）。"""
    if not rs:
        return []
    rs = sorted(rs, key=lambda r: (r.pos, r.a))
    groups: list[list[Rule]] = []
    for r in rs:
        if groups and abs(r.pos - groups[-1][0].pos) <= tol:
            groups[-1].append(r)
        else:
            groups.append([r])
    out = []
    for g in groups:
        g.sort(key=lambda r: r.a)
        pos = sum(r.pos for r in g) / len(g)
        cur = Rule(pos, g[0].a, g[0].b, g[0].width)
        for r in g[1:]:
            if r.a <= cur.b + gap:
                cur.b = max(cur.b, r.b)
            else:
                out.append(cur)
                cur = Rule(pos, r.a, r.b, r.width)
        out.append(cur)
    return [r for r in out if r.b - r.a >= 2.0]


@dataclass
class Cell:
    x0: float
    y0: float
    x1: float
    y1: float
    chars: list[Ch] = field(default_factory=list)
    lines: list = field(default_factory=list)

    @property
    def w(self) -> float:
        return self.x1 - self.x0

    @property
    def h(self) -> float:
        return self.y1 - self.y0

    def text(self) -> str:
        return join_lines(self.lines)

    def visible(self) -> bool:
        return any(not c.text.isspace() for c in self.chars)


def cells_from(hs: list[Rule], vs: list[Rule], tol: float = 2.0) -> list[Cell]:
    """線の交わりから、最小の四角（マス）を作る。"""
    if not hs or not vs:
        return []
    # 交点：縦線 v と横線 h が交わる点
    vs_sorted = sorted(vs, key=lambda r: r.pos)
    vpos = [r.pos for r in vs_sorted]
    pts: dict[tuple[int, int], tuple[float, float]] = {}  # (横線の番号, 縦線の番号) → 点
    by_h: dict[int, list[int]] = {}
    by_v: dict[int, list[int]] = {}
    for hi, h in enumerate(hs):
        lo = bisect.bisect_left(vpos, h.a - tol)
        hi_ = bisect.bisect_right(vpos, h.b + tol)
        budget().count(hi_ - lo, "罫線の交わり")
        for k in range(lo, hi_):
            v = vs_sorted[k]
            if v.a - tol <= h.pos <= v.b + tol:
                pts[(hi, k)] = (v.pos, h.pos)
                by_h.setdefault(hi, []).append(k)
                by_v.setdefault(k, []).append(hi)
    for hi in by_h:
        by_h[hi].sort(key=lambda k: vs_sorted[k].pos)
    for k in by_v:
        by_v[k].sort(key=lambda i: hs[i].pos)
    pos_h = {hi: [vs_sorted[k].pos for k in ks] for hi, ks in by_h.items()}
    pos_v = {k: [hs[i].pos for i in ks] for k, ks in by_v.items()}
    cells = []
    seen = set()
    for (hi, k), (x, y) in pts.items():
        # 右の交点（同じ横線の上）・下の交点：並びの中の位置から順に、要る所まで見る（交点ごとに線の上の全部の交点を見ると、3 乗になる）
        row, col = by_h[hi], by_v[k]
        right0 = bisect.bisect_right(pos_h[hi], x + tol)
        down0 = bisect.bisect_right(pos_v[k], y + tol)
        done = False
        for d_i in range(down0, len(col)):
            h2 = col[d_i]
            y2 = hs[h2].pos
            # 縦線 k が y から y2 までつながっているか
            v = vs_sorted[k]
            if not (v.a - tol <= y and v.b + tol >= y2):
                break
            for r_i in range(right0, len(row)):
                k2 = row[r_i]
                x2 = vs_sorted[k2].pos
                h = hs[hi]
                if not (h.a - tol <= x and h.b + tol >= x2):
                    break
                if (h2, k2) not in pts:
                    continue
                v2 = vs_sorted[k2]
                h_ = hs[h2]
                if v2.a - tol <= y and v2.b + tol >= y2 and h_.a - tol <= x and h_.b + tol >= x2:
                    key = (round(x, 1), round(y, 1), round(x2, 1), round(y2, 1))
                    if key not in seen:
                        seen.add(key)
                        cell = Cell(x, y, x2, y2)
                        cell.edges = (v, v2, h, h_)  # 左・右・上・下の線
                        cells.append(cell)
                    done = True
                    break
            if done:
                break
    return cells


def no_word_space(ch: str) -> bool:
    """語の間を空けずに書く字（行の折り返しで、空白なしでつなぐ字）：和文・中文の全角の字（_is_cjk）と、UAX #14 の行分割の種類
    SA（Complex Context）の文字（タイ・ラーオ・クメール・ビルマ・タイ系の文字：語の切れ目が字からは分からず、空白を入れない）。"""
    if not ch:
        return False
    o = ord(ch[0])
    if 0x0E00 <= o <= 0x0EFF or 0x1000 <= o <= 0x109F or 0x1780 <= o <= 0x17FF or 0x1950 <= o <= 0x19DF \
            or 0x1A20 <= o <= 0x1AAF or 0xA9E0 <= o <= 0xA9FF or 0xAA60 <= o <= 0xAADF or 0x19E0 <= o <= 0x19FF:
        return True
    return _is_cjk(ch)


def join_lines(lines) -> str:
    """マスの中の行をつなぐ。和文の折り返し（行の境の両側が全角）・タイ文字などの語の間を空けない文字の折り返しと、縦に1字ずつ
    並べた字は、空白なしでつなぐ。"""
    out = ""
    for ln in lines:
        t = ln.text.strip()
        if not t:
            continue
        if out and not (no_word_space(out[-1]) and no_word_space(t[0])):
            out += " "
        out += t
    return out


WIDE_BUCKETS = 4096  # 升目の索引に、1 つの四角を置く升の数の上限（ページの全体でも 1,000 升に届かない）


class CellIndex:
    def __init__(self, cells: list[Cell], bucket: float = 24.0):
        self.cells = cells
        self.bucket = bucket
        self.grid: dict[tuple[int, int], list[int]] = {}
        self.wide: list[int] = []  # 升目に置くと多すぎる大きなマス（ページより何桁も大きな宣言）：いつも見る
        for i, c in enumerate(cells):
            ax, bx = int(c.x0 // bucket), int(c.x1 // bucket)
            ay, by = int(c.y0 // bucket), int(c.y1 // bucket)
            if (bx - ax + 1) * (by - ay + 1) > WIDE_BUCKETS:
                self.wide.append(i)
                continue
            for gx in range(ax, bx + 1):
                for gy in range(ay, by + 1):
                    self.grid.setdefault((gx, gy), []).append(i)

    def at(self, x: float, y: float) -> Cell | None:
        best = None
        for i in [*self.grid.get((int(x // self.bucket), int(y // self.bucket)), ()), *self.wide]:
            c = self.cells[i]
            if c.x0 - 0.5 <= x <= c.x1 + 0.5 and c.y0 - 0.5 <= y <= c.y1 + 0.5:
                if best is None or c.w * c.h < best.w * best.h:
                    best = c
        return best

    def near(self, x0, y0, x1, y1) -> list[Cell]:
        out = {i: self.cells[i] for i in self.wide}
        b = self.bucket
        ax, bx = int(x0 // b), int(x1 // b)
        ay, by = int(y0 // b), int(y1 // b)
        if (bx - ax + 1) * (by - ay + 1) > len(self.grid):  # 大きな範囲：升目を 1 つずつ探すより、ある升目を見る
            for (gx, gy), items in self.grid.items():
                if ax <= gx <= bx and ay <= gy <= by:
                    for i in items:
                        out[i] = self.cells[i]
            return [out[i] for i in sorted(out)]
        for gx in range(ax, bx + 1):
            for gy in range(ay, by + 1):
                for i in self.grid.get((gx, gy), ()):
                    out[i] = self.cells[i]
        return list(out.values())


# ====================================================================== 行
@dataclass
class Line:
    chars: list[Ch]
    text: str
    pos: list[tuple[float, float]]  # text の1字ごとの (x0, x1)（縦書きは (y0, y1)）
    x0: float
    y0: float
    x1: float
    y1: float
    vertical: bool
    size: float

    def span_rect(self, a: int, b: int) -> tuple[float, float, float, float]:
        """text[a:b] の位置（右から左の字は、読む順と見た目の順が逆なので、範囲の字の両端で決める）。"""
        ps = self.pos[a:b] or [self.pos[min(a, len(self.pos) - 1)]]
        lo, hi = min(p[0] for p in ps), max(p[1] for p in ps)
        if self.vertical:
            return (self.x0, lo, self.x1, hi)
        return (lo, self.y0, hi, self.y1)


def make_lines(chars: list[Ch], overlaps: list | None = None) -> list[Line]:
    """overlaps：違う字が重なって描かれた行（2つの行を重ねて印刷したもの）を、分けた行の組で受け取る。"""
    out: list[Line] = []
    hor = [c for c in chars if not c.vertical]
    ver = [c for c in chars if c.vertical and not c.upward]
    up = [c for c in chars if c.vertical and c.upward]
    for vertical, group in ((False, hor), (True, ver), (True, up)):
        if not group:
            continue
        group.sort(key=lambda c: (c.base, c.x0 if not vertical else c.y0))
        rows: list[list[Ch]] = []
        total = biggest = 0.0  # いまの行の基準線の和と、いちばん大きな字（字ごとに行の全部を見ると、長い行で 2 乗になる）
        for c in group:
            if rows:
                last = rows[-1]
                ref = total / len(last)
                sz = max(biggest, c.size)
                if abs(c.base - ref) <= 0.45 * sz:
                    last.append(c)
                    total += c.base
                    biggest = sz
                    continue
            rows.append([c])
            total, biggest = c.base, c.size
        rows = [sub for row in rows for sub in _split_stacked(row, vertical)]
        for row in rows:
            row.sort(key=lambda c: (-c.y1 if c.upward else c.y0) if vertical else c.x0)
            row[:] = _marks_after_bases(row, vertical)
            parts = _split_overprinted(row)
            if len(parts) > 1:
                made = [_line(sorted(p, key=lambda c: (-c.y1 if c.upward else c.y0) if vertical else c.x0), vertical) for p in parts]
                for ln in made:
                    ln.overprint = id(made)  # 重ねて描かれた行の組（窓では「／」で区切る）
                if overlaps is not None:
                    overlaps.append(made)
                out += made
                continue
            out.append(_line(row, vertical))
    out = _vertical_stacks(out)
    # 横書きの行は上から。縦書きの列は、右の列から（横書きの行と混ざるときは、縦書きの列のまとまりを、いちばん上の列の位置に置く）
    hor_l = sorted((ln for ln in out if not ln.vertical), key=lambda ln: (ln.y0, ln.x0))
    ver_l = sorted((ln for ln in out if ln.vertical), key=lambda ln: (-ln.x1, ln.y0))
    if not ver_l:
        return hor_l
    top = min(ln.y0 for ln in ver_l)
    k = next((i for i, ln in enumerate(hor_l) if ln.y0 > top), len(hor_l))
    return hor_l[:k] + ver_l + hor_l[k:]


def _marks_after_bases(row: list[Ch], vertical: bool) -> list[Ch]:
    """結合する記号（Mn・Me）は、基の字の後ろの字（Unicode の決まり）。記号の箱の左端は次の字の左端より少し右に来ることがあり、位置で並べると
    次の字の後ろに回る（タイ語の「ชื่อ」が「ชอ่ื」に）。描いた順で基の字のすぐ後の記号は、位置によらず基の字の後ろに置く。
    左から右へ書く字（Unicode の方向 L）だけ：右から左の字は描く順が字の順とは限らない（_logical_order が決める）。"""
    if vertical or not any(len(c.text) == 1 and unicodedata.category(c.text) in ("Mn", "Me") for c in row):
        return row
    by_seq = {c.seq: c for c in row if c.seq >= 0}
    root: dict[int, Ch] = {}  # 記号 → 基の字
    after: dict[int, list[Ch]] = {}
    for c in sorted(row, key=lambda c: c.seq):
        if not (len(c.text) == 1 and unicodedata.category(c.text) in ("Mn", "Me")):
            continue
        p = by_seq.get(c.seq - 1)
        if p is None:
            continue
        b = root.get(id(p), p)
        if unicodedata.bidirectional(b.text[:1]) == "L" and b.x0 - 0.5 * b.size <= c.x0 <= b.x1 + 0.5 * b.size:
            root[id(c)] = b
            after.setdefault(id(b), []).append(c)
    if not root:
        return row
    out: list[Ch] = []
    for c in row:
        if id(c) not in root:
            out.append(c)
            out += after.get(id(c), [])
    return out


def _split_stacked(row: list[Ch], vertical: bool) -> list[list[Ch]]:
    """行の帯（いちばん大きな字で決まる）に入った、上下に積まれた別の行を分ける。大きな見出しの横に小さな字の 2 行があると、基準線の違う字が
    同じ帯に入り、字が 1 字ずつ交互に混ざる（「ADYEU NATAS…」）。1 つの行の字は送りの向きに進むので、送りの向きで大きく重なって基準線の違う
    字は、同じ行の字ではない。そういう字があるときだけ、字ごとの大きさで行に分け直す。"""
    if len(row) < 2:
        return [row]
    at = (lambda c: (c.y0, c.y1)) if vertical else (lambda c: (c.x0, c.x1))
    ink = sorted((c for c in row if not c.text.isspace() and not all(_is_mark(ch) for ch in c.text)), key=at)

    def stacked(a: Ch, b: Ch) -> bool:
        (a0, a1), (b0, b1) = at(a), at(b)
        inter = min(a1, b1) - max(a0, b0)
        return inter > 0.5 * max(1e-6, min(a1 - a0, b1 - b0)) and abs(a.base - b.base) > 0.45 * min(a.size, b.size)

    if not any(stacked(a, b) for a, b in zip(ink, ink[1:])) and not any(stacked(a, b) for a, b in zip(ink, ink[2:])):
        return [row]
    out: list[list[Ch]] = []
    for c in sorted(row, key=lambda c: (c.base, at(c)[0])):
        if out:
            last = out[-1]
            ref = sum(x.base for x in last) / len(last)
            size = sum(x.size for x in last) / len(last)
            if abs(c.base - ref) <= 0.45 * min(size, c.size):
                last.append(c)
                continue
        out.append([c])
    return out


def _clash(a: Ch, b: Ch) -> bool:
    """違う字の箱が大きく重なる（結合する記号は、基の字に重ねて描くものなので数えない）。"""
    if a.text == b.text or a.text.isspace() or b.text.isspace():
        return False
    if all(_is_mark(ch) for ch in a.text) or all(_is_mark(ch) for ch in b.text):
        return False
    return _box_overlap(a, b) > 0.5


def _grid_add(grid: dict, c: Ch, step: float) -> None:
    grid.setdefault(int(c.x0 // step), []).append(c)


def _clashes_grid(c: Ch, grid: dict, step: float) -> bool:
    """c が、grid（字の左の端の網の目 → 字）のどれかと _clash するか。箱が重なる字は、横の位置が字の幅の中にある。比べた数を予算に数える。"""
    k = int(c.x0 // step)
    for at in (k - 1, k, k + 1):
        cand = grid.get(at)
        if cand:
            budget().count(len(cand), "重ねて描いた行の字の比べ")
            if any(_clash(c, o) for o in cand):
                return True
    return False


def _split_overprinted(row: list[Ch]) -> list[list[Ch]]:
    """同じ行の帯に、違う字が重なって描かれているとき（2つの行を重ねて印刷した様式）、描いた順で行を分ける。
    描いた順に字をたどり、前の字と重なる所で切り、互いに重ならないまとまりどうしをつなぐ。"""
    if len(row) < 2:
        return [row]
    ink = [c for c in row if not c.text.isspace()]
    if not any(_clash(a, b) for a, b in zip(ink, ink[1:])) and not any(_clash(a, b) for a, b in zip(ink, ink[2:])):
        return [row]
    # 描いた順の続き：前の字より左へ戻った所（次の行を描き始めた所）で切る。約物を詰めた字（「・」「（」の箱が次の字と重なる）は、
    # 描いた順では左から右へ進み続けるので、切れない
    step = max((c.x1 - c.x0 for c in row), default=0.0) or 1.0  # 網の目の幅：いちばん広い字。重なる字は、同じか隣の目にある
    segs: list[list[Ch]] = []
    seg_grid: dict = {}
    for c in sorted(row, key=lambda c: c.seq):
        if segs:
            p = segs[-1][-1]
            back = c.x0 < p.x0 - 0.1 * max(c.size, p.size) if not c.vertical else c.y0 < p.y0 - 0.1 * max(c.size, p.size)
            if not (back and _clashes_grid(c, seg_grid, step)):
                segs[-1].append(c)
                _grid_add(seg_grid, c, step)
                continue
        segs.append([c])
        seg_grid = {}
        _grid_add(seg_grid, c, step)
    if len(segs) < 2:
        return [row]
    groups: list[list[Ch]] = []
    grids: list[dict] = []
    for sg in segs:
        for g, grid in zip(groups, grids):
            if not any(_clashes_grid(x, grid, step) for x in sg):
                g += sg
                for x in sg:
                    _grid_add(grid, x, step)
                break
        else:
            groups.append(list(sg))
            grids.append({})
            for x in sg:
                _grid_add(grids[-1], x, step)
    return groups


def _vertical_stacks(lines: list[Line]) -> list[Line]:
    """横書きの字を縦に1字ずつ積んだもの（表の縦の見出し「申／請／者」）は、上から読む1つの縦の行にする。
    同じ列の、回した字（縦書き用のかっこなど）も一緒にする。"""
    singles = [ln for ln in lines if not ln.vertical and len(ln.text.strip()) <= 2]
    if len(singles) < 3:
        return lines
    used: set[int] = set()
    out = []
    singles.sort(key=lambda ln: ((ln.x0 + ln.x1) / 2, ln.y0))
    for i, a in enumerate(singles):
        if id(a) in used:
            continue
        col = [a]
        cx = (a.x0 + a.x1) / 2
        for b in singles[i + 1:]:
            if id(b) in used:
                continue
            if abs((b.x0 + b.x1) / 2 - cx) <= 0.6 * max(a.size, b.size):
                col.append(b)
        col.sort(key=lambda ln: ln.y0)
        # 間が大きく空くところで切る
        runs = [[col[0]]]
        for b in col[1:]:
            if b.y0 - runs[-1][-1].y1 <= 1.6 * max(b.size, runs[-1][-1].size):
                runs[-1].append(b)
            else:
                runs.append([b])
        for run in runs:
            if len(run) >= 3:
                for ln in run:
                    used.add(id(ln))
                # 同じ列の縦書きの字も入れる
                x0 = min(ln.x0 for ln in run)
                x1 = max(ln.x1 for ln in run)
                y0 = min(ln.y0 for ln in run)
                y1 = max(ln.y1 for ln in run)
                extra = [ln for ln in lines if ln.vertical and id(ln) not in used and ln.x0 >= x0 - 3 and ln.x1 <= x1 + 3
                         and ln.y0 >= y0 - 3 and ln.y1 <= y1 + 3]
                for ln in extra:
                    used.add(id(ln))
                chars = sorted([c for ln in run + extra for c in ln.chars], key=lambda c: c.y0)
                st = _line(chars, True, stacked=True)
                st.stacked = len(run) + len(extra)  # 縦に積んだ字の数（窓で、縦の行として読んだことを示す）
                out.append(st)
    return [ln for ln in lines if id(ln) not in used] + out


# 和字の空白（U+3000 IDEOGRAPHIC SPACE）の幅：字の大きさの 1 字ぶん（1 em）。Unicode（第 6 章 6.2 の空白の字の表：
# 「width of ideographic (CJK) characters」）と JIS X 4051（日本語文書の組版方法：全角の空白は 1 字ぶん）で決まる幅。
# 字と字の間がこの幅以上空いていれば、字の間に 1 字を書ける空き（「第　号」「年　月　日」）とする。空きが U+3000 の字で
# 描かれても、半角の空白の並びでも、字の位置を動かしただけでも、測った幅で決める。欧文の語の間の空白は 0.25〜0.33 em
# （両端そろえで広げても 0.6 em ほど）で、この幅に届かない。プロポーショナルな和文の書体（MS P明朝 など）は U+3000 を
# 0.66 em ほどの送りで描くので、その 1 字だけの空きは 1 em に届かず、空きにしない（字の大きさに対する幅で決める）。
EM_SPACE = 1.0
EM_ROUND = 0.02  # 測った位置の丸め（座標・字の幅は小数 2 桁ほどで書かれる）の分だけ、1 em にわずかに届かないものも 1 em とする


def wide_gap(gap: float, size: float) -> bool:
    """字と字の間の空き（ポイント）が、和字の空白 1 字ぶん（1 em）以上か。size は両どなりの字の小さい方の大きさ。"""
    return gap >= (EM_SPACE - EM_ROUND) * size


def _spaces(parts: list, pos: list, a1: float, gap: float, sz: float, cjk: bool, had_space: bool,
            em: float | None = None) -> None:
    """字と字の間の空き（ポイント）を、空白の字にする。窓の空白の字は、測った幅を表す（gap_blanks はこれで空きを決める）：
    - 和字の空白 1 字ぶん（1 em。em：両どなりの字の小さい方の大きさ）以上の空き：全角の文字のまわりは全角の空白（1 つを 1 字ぶん）、ほかは
      2 つ以上の半角の空白（1 つを半字ぶん）。どちらも、字の間に 1 字を書ける空き
    - 1 em に届かない空き：半角の空白 1 つ（語の間の空白。空きではない）"""
    em = sz if em is None else em
    if not wide_gap(gap, em):
        if gap <= (0.3 if cjk else 0.15) * sz and not had_space:
            return
        parts.append(" ")
        pos.append((a1, a1 + gap))
        return
    if cjk:
        n = max(1, round(gap / sz))
        ch = "\u3000"
    else:
        n = max(2, round(gap / (0.5 * sz)))
        ch = " "
    step = gap / n
    for k in range(n):
        parts.append(ch)
        pos.append((a1 + k * step, a1 + (k + 1) * step))


def _line(row: list[Ch], vertical: bool, stacked: bool = False) -> Line:
    units: list[list] = []  # [字, [(位置)...], 字の Ch]（グリフごと。空白は1字ずつ）
    prev: Ch | None = None
    spaces = [c for c in row if c.text.isspace()]
    ink = [c for c in row if not c.text.isspace()] or row[:1]
    row = ink
    for c in row:
        if prev is not None:
            a1 = prev.y1 if vertical else prev.x1
            b0 = c.y0 if vertical else c.x0
            if c.upward:  # 下から上へ読む行：前の字の上端から、次の字の下端までが空き。向きをそろえて測る
                a1, b0 = -prev.y0, -c.y1
            gap = b0 - a1
            sz = max(prev.size, c.size)
            had_space = any((-x.cy if c.upward else x.cy if vertical else x.cx) > a1 - 0.5 and (-x.cy if c.upward else x.cy if vertical else x.cx) < b0 + 0.5
                            for x in spaces)
            cjk = _is_cjk(prev.text[-1:]) or _is_cjk(c.text[:1])
            em = min(prev.size, c.size)  # 和字の空白 1 字ぶんの幅（1 em）：両どなりの字の小さい方の大きさ
            if stacked:
                gap = gap - 0.8 * sz if gap > 1.5 * sz else 0.0
            if gap > 0 or had_space:
                sp_parts: list[str] = []
                sp_pos: list[tuple[float, float]] = []
                _spaces(sp_parts, sp_pos, a1, max(gap, 0.0), sz, cjk, had_space, em)
                units += [[t, [q], None] for t, q in zip(sp_parts, sp_pos)]
        t = c.text
        lo, hi = (c.y0, c.y1) if vertical else (c.x0, c.x1)
        step = (hi - lo) / max(1, len(t))
        units.append([t, [(lo + k * step, lo + (k + 1) * step) for k in range(len(t))] if len(t) > 1 else [(lo, hi)], c])
        prev = c
    rtl = False
    orphans: list[int] = []
    if not vertical and not stacked:
        units, rtl, orphans = _logical_order(units)
    text = "".join(u[0] for u in units)
    pos = [q for u in units for q in u[1]]
    ln = Line(row, text, pos, min(c.x0 for c in row), min(c.y0 for c in row), max(c.x1 for c in row),
              max(c.y1 for c in row), vertical, max(c.size for c in row))
    ln.rtl = rtl
    ln.orphans = orphans
    return ln


# Unicode の IndicPositionalCategory.txt（Unicode 15.1）で「Left」の字：子音の前（左）に描くが、読む順では子音の後ろの母音の記号。
# タイ文字・ラーオ文字の前に書く母音（Visual_Order_Left）は、Unicode でも見た目の順に並べるので含めない。
PRE_BASE = set(map(chr, [
    0x093F, 0x094E,  # デーヴァナーガリー ि ॎ
    0x09BF, 0x09C7, 0x09C8,  # ベンガル ি ে ৈ
    0x0A3F,  # グルムキー ਿ
    0x0ABF,  # グジャラート િ
    0x0B47,  # オリヤー େ
    0x0BC6, 0x0BC7, 0x0BC8,  # タミル ெ ே ை
    0x0D46, 0x0D47, 0x0D48,  # マラヤーラム െ േ ൈ
    0x0DD9, 0x0DDB,  # シンハラ ෙ ෛ
    0x1031, 0x1084,  # ビルマ ေ ႄ
    0x17C1, 0x17C2, 0x17C3,  # クメール េ ែ ៃ
    0x1A55, 0x1A6E, 0x1A6F, 0x1A70, 0x1A71, 0x1A72,  # タイ・タム
    0x1B3E, 0x1B3F,  # バリ
    0x1BA6,  # スンダ
    0xA9BA, 0xA9BB,  # ジャワ
    0xAA2F, 0xAA30, 0xAA34,  # チャム
]))


def _is_mark(ch: str) -> bool:
    return unicodedata.category(ch) in ("Mn", "Mc", "Me")


def _virama(ch: str) -> bool:
    return unicodedata.combining(ch) == 9  # 正準結合クラス 9（Virama）


def _logical_order(units: list[list]) -> tuple[list[list], bool, list[int]]:
    """見た目の順（左から右）のグリフ → 読む順。
    - 右から左の字（アラビア文字・ヘブライ文字）を含む行：UAX #9 のレベルで並べ替える（数は左から右のまま。グリフごとに動かす）
    - インド系の文字の、子音の前に描く母音の記号：子音（の結合）の後ろへ動かし、分けて描いた母音（ে…া）は NFC で1字にする
    戻り値：(並べた単位, 右から左の段落か, 前の字のない結合の記号の位置（字の番号）)。"""
    from .bidi import has_rtl, visual_to_logical

    texts = [u[0] for u in units]
    rtl = False
    if has_rtl(texts):
        units = _attach_marks(units)
        order, rtl = visual_to_logical([u[0] for u in units])
        units = [units[i] for i in order]
    elif any(ch in PRE_BASE for t in texts for ch in t):
        units = _move_pre_base(units)
    orphans = []
    at = 0
    for k, u in enumerate(units):
        t = u[0]
        if t and _is_mark(t[0]) and not t.isspace():
            before = units[k - 1][0] if k else ""
            if not before or before.isspace():
                orphans.append(at)
        at += len(t)
    return units, rtl, orphans


def _attach_marks(units: list[list]) -> list[list]:
    """結合する記号だけのグリフ（アラビア文字の母音の記号など）を、重なる基の字のグリフにつける（読む順で、基の字の後ろ）。"""
    out: list[list] = []
    for k, u in enumerate(units):
        t = u[0]
        if t and all(_is_mark(ch) for ch in t) and u[2] is not None:
            c = (u[1][0][0] + u[1][-1][1]) / 2
            cand = [v for v in (out[-1] if out else None, units[k + 1] if k + 1 < len(units) else None)
                    if v is not None and v[2] is not None and not v[0].isspace()]
            host = next((v for v in cand if v[1][0][0] - 0.5 <= c <= v[1][-1][1] + 0.5), cand[0] if cand else None)
            if host is not None:
                host[0] += t
                host[1] = host[1] + [host[1][-1]] * len(t)
                continue
        out.append(u)
    return out


def _move_pre_base(units: list[list]) -> list[list]:
    units = [list(u) for u in units]
    out: list[list] = []
    k = 0
    while k < len(units):
        u = units[k]
        if u[0] and u[0][0] in PRE_BASE and all(_is_mark(ch) for ch in u[0]) and k + 1 < len(units) \
                and units[k + 1][0] and not _is_mark(units[k + 1][0][0]) and not units[k + 1][0].isspace():
            # 子音の結合：次のグリフ（子音）と、Virama で終わる間は、その次の子音も
            j = k + 1
            cluster = [units[j]]
            while cluster[-1][0] and _virama(cluster[-1][0][-1]) and j + 1 < len(units) and units[j + 1][0] \
                    and not units[j + 1][0].isspace():
                j += 1
                cluster.append(units[j])
            # 子音につく記号（ヌクタなど、正準結合クラスが 0 でない Mn）も子音の側
            while j + 1 < len(units) and units[j + 1][0] and all(_is_mark(ch) and unicodedata.combining(ch) for ch in units[j + 1][0]):
                j += 1
                cluster.append(units[j])
            sign = u
            # 分けて描いた母音：前の記号と、子音の後ろの記号が、正準の合成で1字になる（ে + া → ো）
            if j + 1 < len(units) and units[j + 1][0] and _is_mark(units[j + 1][0][0]):
                comp = unicodedata.normalize("NFC", sign[0] + units[j + 1][0])
                if len(comp) < len(sign[0]) + len(units[j + 1][0]):
                    nxt = units[j + 1]
                    span = (min(sign[1][0][0], nxt[1][0][0]), max(sign[1][-1][1], nxt[1][-1][1]))
                    sign = [comp, [span] * len(comp), sign[2]]
                    j += 1
            out += cluster
            out.append(sign)
            k = j + 1
            continue
        out.append(u)
        k += 1
    return out
