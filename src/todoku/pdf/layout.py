"""PDFのページを、人が読む順に並べ、入力欄を見つける（決定的な処理だけで）。

1〜3 は pagetext（字・罫線・マス・行。名前はここからも使える）：
1. 文字・線をページの見た目の座標（左上が原点、下向きが y）に直す（ページの回転も反映）
2. 罫線から表のマス（セル）を組み立てる（線の交わりから最小の四角を作る）
3. 文字をマスに振り分け、マスの中・マスの外で行を作る（字間の空きは空白として残す）
4. 書く所を見つける（ひとつの決め方：書類の作りだけで。字の意味・見出しのあるなし・経験の閾値では決めない。下の「書く所」の節）：
   - 空のマス（と、同じ大きさの空のマスの並び）
   - 字の形の空欄（下線の字・点線・ダッシュの並び、かっこの中の空白）と、測った 1 字ぶん以上の空き（字の間・行の前後）
   - チェックの箱の字、下線の罫線
   どの所にも、事実（位置と大きさ mm・周りの字とその離れ mm・塗り）を添える。id は位置から決める（place_id）
5. 行を上から、行の中は左から並べた「読み順」を作る
"""

from __future__ import annotations

import re
import math
import unicodedata
from dataclasses import dataclass, field

from .. import chars as _chars  # 字の種類（空白・下線の字・ダッシュ・点線・かっこの組・チェックの箱の字）：Unicode の資料の表
from ..xlsx.window import label_text
from .content import PageContent
# 字・罫線・マス・行（pagetext）：ここからも同じ名前で使える（ほかのモジュール・テストの import のため）
from .pagetext import (  # noqa: F401
    WIDE_BUCKETS, _attach_marks, BOX_GLYPHS, _box_overlap, Cell, CellIndex, cells_from, Ch, _clash, EM_ROUND, EM_SPACE, glyph_frame,
    _is_cjk, _is_hangul, _is_mark, is_overprint, _is_pua, join_lines, Line, _line, _logical_order, make_lines, _merge,
    MM_PT, _move_pre_base, no_word_space, PRE_BASE, _ruby_base, Rule, rulings, _spaces, _split_overprinted, user_unit,
    _vertical_runs, _vertical_stacks, View, view_chars, _virama, wide_gap,
)


def _sentence_end(ch: str) -> bool:
    """文の終わりの字（Unicode の Sentence_Terminal：。 . ! ? । ۔ ؟ ። ։ など）。"""
    return _chars.is_sentence_terminal(ch)


# 記入済みの印。U+25A3 ▣（WHITE SQUARE CONTAINING BLACK SMALL SQUARE）は、とどくが ActiveX のどちらでもない箱に書く字
CHECKED_GLYPHS = set(_chars.CHECKED_BOXES) | set(_chars.MAYBE_CHECKED)  # ▣ は MAYBE_CHECKED に入っている


# ====================================================================== 書く所（場所）
# 書く所の決め方は、ひとつ（書類の作りだけで決める。字の意味・語・見出しのあるなし・経験で選んだ大きさの閾値では決めない）：
#   1. 空のマス：罫線で囲まれ、見える字のないマス（幅・高さが MIN_GLYPH 以上。それより細いものは字の入らないすき間として数だけ）。
#      同じ大きさの空のマスが横に隣り合って並ぶもの（2 つ以上）は、並び全体も 1 つの所（comb：1 字ずつ書くか、1 つの欄として
#      書くか）として示す（マスどうしは、それぞれの所のまま）
#   2. 字の形の空欄：下線の字（Unicode の Pc）の並びで幅が 1 字ぶん以上のもの、点線（leader）・ダッシュ（Pd）が 2 字以上続き
#      幅が 1 字ぶん以上のもの、かっこの組（BidiBrackets.txt）の中が空白だけのもの、同じ記号の字（一般カテゴリ So・Sm・Nl：
#      「○○」「〇〇」「××」）が 2 字以上続くもの（伏せ字の形。書くときは白い下地で覆う。行の字がその並びだけの行は、字で描いた線）
#   3. 測った空き：字と字の間、行の頭の前・行の終わりの後ろ（マスの縁・同じ高さの字・縦の罫線・ページの字の端まで）の空きで、
#      1 字ぶん（1 em：JIS X 4051・Unicode 6.2 の全角の空白の幅）以上のもの。組版の決まりでできる空き（文の終わりの字
#      Sentence_Terminal の後ろの行末・左右の空きが等しい中央そろえの行の両側・段落の 1 字下げ）は除く
#   4. チェックの箱の字（□ など。どれも「？」：箱か箇条の印かは決めない。記入済みの箱の字 ☑ ■ などは、書き換えられない所として）
#   5. 下線の罫線：マスの縁でも表の中の線でもない横線で、すぐ上に字のないもの（線の上の 1 行の帯）
# 行の所（2・3）は行のもの：書く四角は行の帯（line_band）で、字は行の基線に行の字の大きさで 1 行に書く。上下の空きは事実（room_mm）。
# 所は事実（位置・大きさ mm・周りの字とその離れ mm）と一緒に示し、書く所かどうか（意味）はエージェントが決める。
# 規則の欄（窓の［p1.t…］）と「？」（［p1.g…？］）の違いも作りで決める：描かれた空欄（空のマス・下線の罫線・下線の字・
# かっこの中）は規則の欄、字の間の空き（空白・点線・ダッシュ・行の前後の空き）・箱の字・空のマスの並びは「？」。
DASH_MIN = 2  # 点線・ダッシュの並び：同じ種類の字が 2 字以上続くもの（1 字だけのものは句読点・区切りの記号）
RULE_TOL = 2.0  # 罫線の交わりの許し（cells_from と同じ）：マスの大きさ・隣り合いを比べるときも、この幅の違いは同じとみなす
EDGE_PAD = 1.5  # 書く所の四角を、マスの縁（罫線）から離す幅（pt）：書いた字が罫線に触れないように（書き込みの安全のための余白）
TEXT_PAD = 0.2  # 書く所の四角を、隣の字から離す幅（字の大きさに対して）
# 窓に示す、見つけた形の名前（why の短い分類。測った大きさは facts に）
GAP_WHY = {"gap": "字の間の空き（1 字ぶん以上）", "lead": "行の頭の前の空き（1 字ぶん以上）",
           "trail": "行の終わりの後ろの空き（1 字ぶん以上）", "leader": "点線の並び", "dash": "ダッシュの並び",
           "connector": "下線の字の並び", "bracket": "かっこの組の中の空白", "blank_line": "空白の字だけの行",
           "placeholder": "同じ記号の字の並び（伏せ字の形）"}


def _char_kind(ch: str) -> str | None:
    if _chars.is_space(ch):
        return "space"
    if _chars.is_connector(ch):
        return "connector"
    if ch in _chars.LEADERS:
        return "leader"
    if _chars.is_dash(ch):
        return "dash"
    return None


def _runs(text: str) -> list[tuple[int, int, str]]:
    """同じ種類の字の並び：(始め, 終わり, 種類)（space・connector・leader・dash）。"""
    out = []
    i, n = 0, len(text)
    while i < n:
        k = _char_kind(text[i])
        if k is None:
            i += 1
            continue
        j = i + 1
        while j < n and _char_kind(text[j]) == k:
            j += 1
        out.append((i, j, k))
        i = j
    return out


# 1 字ずつの字を空白でつないだ並び（「氏　名」「事 業 所」。空白の長さは問わない）
_SPACED = re.compile(r"(?:\S[ 　]+)+\S")


def evenly_spaced(core: str) -> bool:
    """1 字ずつの字（Unicode の文字・数字：一般カテゴリ L・N）が、同じ長さの空白の並び（違いは 1 字まで）で 2 字以上つながる
    （字間を空けた「氏　　名」「事　業　所」）。形の事実（facts の spaced）として示すだけで、書く所かどうかは決めない。"""
    runs = [(a, b) for a, b, k in _runs(core) if k == "space"]
    if not runs or runs[0][0] == 0 or runs[-1][1] == len(core):
        return False
    toks = [core[a:b] for a, b in zip([0] + [y for _x, y in runs], [x for x, _y in runs] + [len(core)])]
    lens = [b - a for a, b in runs]
    return len(toks) >= 2 and all(len(t) == 1 and unicodedata.category(t)[0] in "LN" for t in toks) \
        and max(lens) - min(lens) <= 1


class Blank(tuple):
    """文字の中の空欄 (始め, 終わり)。structural：描かれた空欄（下線の字の並びで英数字に接しないもの・かっこの組の中が空白だけの
    もの）か。偽のものは字の間の空き（窓の「？」）。why は見つけた形（GAP_WHY の鍵）。2つの数の組としてそのまま使える。"""

    structural: bool
    why: str

    def __new__(cls, a: int, b: int, structural: bool = False, why: str = ""):
        t = super().__new__(cls, (a, b))
        t.structural = structural
        t.why = why
        return t


def underscore_touches_word(text: str, a: int, b: int) -> bool:
    """下線の字の並び（text[a:b] を含む、続いた並び全体）の前か後ろが、英数字・「/」「.」「-」か（名前・URL の一部の形）。"""
    def ident(ch: str) -> bool:
        return ch.isascii() and (ch.isalnum() or ch in ".-/")

    x, y = a, b
    while x > 0 and _chars.is_connector(text[x - 1]):
        x -= 1
    while y < len(text) and _chars.is_connector(text[y]):
        y += 1
    return (0 < x and ident(text[x - 1])) or (y < len(text) and ident(text[y]))


def _em_width(text: str, a: int, b: int) -> float:
    """字の幅（1 字ぶん＝1 em を 1 とする）：測った位置がないときの目安。UAX #11 の東アジアの幅が W・F の字は 1、ほかは 0.5。"""
    return sum(1.0 if unicodedata.east_asian_width(ch) in ("W", "F") else 0.5 for ch in text[a:b])


def _ends_sentence(core: str) -> bool:
    """行が文の終わりで終わるか：最後の字が Sentence_Terminal か、その後ろに閉じかっこ・引用符（一般カテゴリ Pe・Pf と
    QUOTATION MARK）だけが続くもの（UAX #29 の文の区切り：STerm Close* Sp*）。「（部位別に記入すること。）」も文の終わり。"""
    t = core.rstrip(" 　")
    while t and (unicodedata.category(t[-1]) in ("Pe", "Pf") or t[-1] in "\"'"):
        t = t[:-1]
    return bool(t) and _sentence_end(t[-1])


def gap_blanks(text: str, lead: int = 0, trail: int = 0, width=None, align: bool = True) -> list[Blank]:
    """行の字の中の書く所（上の決め方の 2・3 のうち、字の並びで決まるもの）。各要素は Blank。
    width(a, b)：text[a:b] の測った幅（1 em を 1 とする数）。省くと字の幅の目安（_em_width）で測る。
    - 空白の並び：1 em 以上（PDF の行の空白は測った空きから作る：1 em 以上の空きだけが全角の空白か 2 字以上の半角の空白になる）
    - 下線の字の並び：1 em 以上（英数字に接するものは「？」：名前の一部の形）
    - 点線・ダッシュ：同じ種類の字が 2 字以上続き、1 em 以上
    - かっこの組の中が空白だけのもの（幅は問わない）
    lead・trail：行の前後に足した空白の数（行の頭の前・行の終わりの後ろの空き。呼ぶ側が 1 em 以上のときだけ足す）。
    その空きは、文の終わりの字で終わる行の後ろ・左右の空きが等しい（差が 1 字に満たない）中央そろえの行の両側では示さない
    （align が偽なら、中央そろえは見ない：左右の端が描かれた端でないとき）。"""
    w = width or (lambda a, b: _em_width(text, a, b))
    n = len(text)
    core_a, core_b = lead, n - trail
    while core_a < core_b and _chars.is_space(text[core_a]):
        core_a += 1
    while core_b > core_a and _chars.is_space(text[core_b - 1]):
        core_b -= 1
    core = text[core_a:core_b]
    out: list[Blank] = []
    for a, b, kind in _runs(text):
        if kind == "connector":
            if w(a, b) >= EM_SPACE - EM_ROUND:
                out.append(Blank(a, b, not underscore_touches_word(text, a, b), "connector"))
            continue
        if kind in ("leader", "dash"):
            if b - a >= DASH_MIN and w(a, b) >= EM_SPACE - EM_ROUND:
                out.append(Blank(a, b, False, kind))
            continue
        if a < core_a or b > core_b:
            if not core:
                continue  # 字のない行
            if b <= core_a:
                out.append(Blank(a, b, False, "lead"))
            elif a >= core_b and not _ends_sentence(core):
                out.append(Blank(a, b, False, "trail"))  # 文の終わりの字で終わる行の後ろは、文の終わり
            continue
        opener = text[a - 1] if a > 0 else ""
        if opener and b < n and _chars.closing(opener) == text[b]:
            out.append(Blank(a, b, True, "bracket"))  # 「（　　）」「[ ]」：かっこの組の中が空白だけ
            continue
        if w(a, b) >= EM_SPACE - EM_ROUND:
            out.append(Blank(a, b, False, "gap"))
    # 同じ記号の字の並び（伏せ字の形）：一般カテゴリ So・Sm・Nl の同じ字が 2 字以上。箱の字（チェックの箱）は除く
    i = 0
    while i < n:
        ch = text[i]
        j = i + 1
        while j < n and text[j] == ch:
            j += 1
        if j - i >= 2 and unicodedata.category(ch) in ("So", "Sm", "Nl") and ch not in BOX_GLYPHS \
                and ch not in CHECKED_GLYPHS and text[i:j] != core:
            out.append(Blank(i, j, False, "placeholder"))
        i = j
    out.sort()
    edges = [x for x in out if x.why in ("lead", "trail")]
    if align and len(edges) == 2 and len(out) == 2 and abs(w(*edges[0]) - w(*edges[1])) < EM_SPACE:
        out = []  # 中央にそろえた行（左右の空きの差が 1 字に満たない）：両側の空きは、そろえ方の空き
    return out


MIN_WRITE = 6.0  # これより低い字の行の所は、小さな字の行として（書けるかは書き込みと同じ決め方で）示す
MIN_GLYPH = 4.0  # 書き込みのいちばん小さな字の大きさ（fill.layout_text）：幅か高さがこれに満たないマスには、どんな字も入らない


@dataclass
class Spot:
    """書く所（またはチェックの印）。座標は見た目の座標（pt）。facts は窓に示す事実（mm）。"""

    kind: str  # text / comb / inline / checkbox
    rect: tuple[float, float, float, float]
    label: str = ""
    reason: str = ""
    boxes: list = field(default_factory=list)  # 空のマスの並び：1字ずつの四角
    checked: bool = False
    row_label: str = ""
    column_label: str = ""
    line_text: str = ""  # 文字の中の空欄・チェックの印の行
    blank: tuple[int, int] | None = None
    cell: Cell | None = None
    id: str = ""
    size: float = 10.0
    value: object = None
    widget: object = None
    host: object = None  # 欄を書き入れる行（Line）またはマス（Cell）
    at: tuple[int, int] | None = None  # 行の中で置き換える範囲（文字の位置）
    confidence: str = "high"
    cover: str = ""  # 覆って書く元の字（伏せ字）：書くときは白い下地の欄で覆う（元の字は消さない）
    # 種類が形だけでは決まらない所：書き方の候補（初めが既定の書き方）。fill の as_kind で選ぶ
    kind_candidates: list = field(default_factory=list)
    alternative_to: object = None  # 重なる所（同じ所の別の読み方）：書けるのはどちらか一方（両方に書くと字が重なる）
    readonly: str = ""  # 書けない理由（小さすぎて1字も入らない・印刷された記入済みの印など）。欄としては示す
    shape: str = ""  # 見つけた作り：cell・cells・gap・lead・trail・connector・leader・dash・bracket・box・checked・underline
    anchor: tuple | None = None  # id を決める位置（作りの位置：マスの左上・空きの始まりと行の上。pt）
    facts: dict = field(default_factory=dict)  # 窓に示す事実（mm・字・離れ）
    write_size_from: str = ""  # 既定の字の大きさの出どころ（default_size）
    baseline: float | None = None  # 行の所：書く字の基線（見た目の y、pt。行の字と同じ基線に書く）


def _agent(sp: "Spot", why: str = "") -> "Spot":
    """字の間の空きなど、描かれた空欄でない所にする（窓に「？」）。why：見つけた形の分類。"""
    sp.confidence = "agent"
    if why:
        sp.reason = why
    return sp


# id の頭の字（作りの種類）：t 空のマス・m 空のマスの並び・g 行の中の空き（空白・下線の字・点線・かっこ・行の前後）・b 箱の字・
# u 下線の罫線。とどくが字を目印にして書いた所の w・印影の s・○ の c とは重ならない
ID_LETTER = {"cell": "t", "cells": "m", "gap": "g", "lead": "g", "trail": "g", "connector": "g", "leader": "g", "dash": "g",
             "bracket": "g", "blank_line": "g", "placeholder": "g", "box": "b", "checked": "b", "underline": "u"}
# とどくの所の名前（窓の id）：新しい形 p1.t45-112（頭の字・左・上の mm）と、これまでの形 p1.3・p1.a12（番号）
PLACE_ID = r"[tmgbu]\d+-\d+[a-z]?"
OLD_ID = r"a?\d+"


def place_id(page: int, sp: "Spot", mm: float) -> str:
    """所の id：p{ページ}.{頭の字}{左}-{上}（作りの位置 anchor を mm に丸めた数）。ページの中のほかの所が増えても減っても、
    この所の id は変わらない（位置だけで決まる）。同じ頭の字・同じ位置の所が 2 つあるときは、呼ぶ側が後ろに a・b… を足す。"""
    x, y = sp.anchor if sp.anchor is not None else sp.rect[:2]
    return f"p{page}.{ID_LETTER.get(sp.shape, 'g')}{max(0, round(x / mm))}-{max(0, round(y / mm))}"


class PageLayout:
    def line_glyphs(self) -> set[int]:
        """線の代わりに印刷された字（id）：下線の字・基線の点線の字が、同じ行で続く並び（1 字ぶん以上の幅。点線は 2 字以上。
        1 字だけの点は句読点）。基線か基線の下に描かれ、上に書く字を隠さないので、書く場所になる（重なりとして断らない）。"""
        cached = getattr(self, "_line_glyphs", None)
        if cached is not None:
            return cached
        out: set[int] = set()
        low = [c for c in self.chars if c.text and not c.vertical and all(_chars.is_low_line_char(ch) for ch in c.text)]
        low.sort(key=lambda c: (round(c.y0, 1), c.x0))
        run: list = []

        def flush() -> None:
            if not run:
                return
            wide = run[-1].x1 - run[0].x0 >= (EM_SPACE - EM_ROUND) * run[0].size
            connector = any(_chars.is_connector(ch) for ch in run[0].text)
            if wide and (connector or sum(len(c.text) for c in run) >= DASH_MIN):
                out.update(id(c) for c in run)

        for c in low:
            if run and abs(c.y0 - run[-1].y0) < 0.5 and c.x0 - run[-1].x1 < 0.5 * c.size:
                run.append(c)
            else:
                flush()
                run = [c]
        flush()
        self._line_glyphs = out
        return out

    def __init__(self, doc, page: dict, pc: PageContent, number: int, widgets=None):
        self.doc = doc
        self.page = page
        self.number = number
        self.view = View(doc, page)
        self.pc = pc
        self.dropped: list = []
        self.merges: list = []  # 重ね書きとして1字にした字：(字, 箱の重なり, 位置)
        self.chars = view_chars(pc, self.view, self.dropped, self.merges)
        hs, vs = rulings(pc, self.view)
        self.hs, self.vs = hs, vs
        # 斜めの線（マスを斜線で消す印）
        self.diagonals = []
        for sg in pc.segs:
            if all(c > 0.93 for c in sg.color[:3]):
                continue
            a, b = self.view.pt(sg.x0, sg.y0), self.view.pt(sg.x1, sg.y1)
            if abs(a[0] - b[0]) > 0.8 and abs(a[1] - b[1]) > 0.8:
                self.diagonals.append((a, b))
        cells = cells_from(hs, vs)
        # ページ全体を囲むだけの大きなマスは、マスとして扱わない（文字の行を分けないため）
        pw, ph = self.view.width, self.view.height
        # （幅か高さが 2pt に満たない「マス」は、二重線の間のすき間で、字の入る四角ではない）
        self.cells = [c for c in cells if c.w * c.h < 0.5 * pw * ph and c.w >= 2 and c.h >= 2]
        # 大きなマスのうち、中に字もほかのマスもないもの（ページの半分を超える、空の記入欄）は、枠ではないので候補に残す
        self._big_empty = [c for c in cells if c.w * c.h >= 0.5 * pw * ph]
        self.index = CellIndex(self.cells)
        self.free: list[Ch] = []
        for c in self.chars:
            cell = self.index.at(c.cx, c.cy)
            if cell is None:
                self.free.append(c)
            else:
                cell.chars.append(c)
        self.overprinted: list[list[Line]] = []  # 重ねて描かれ、分けた行の組
        self.added: list[str] = []  # 窓で足した字（ルビのかっこ）
        self.ruby_merges: list[tuple] = []  # ルビとして基の行につないだ所：(基の行の字, ルビの字, ルビの行の位置)
        for cell in self.cells:
            cell.lines = self._ruby(make_lines(cell.chars, self.overprinted)) if cell.chars else []
        self.free_lines = self._ruby(make_lines(self.free, self.overprinted))
        self.typical = _typical_size(self.chars)
        self.widgets = widgets or []
        self.fills: list[tuple] = []
        for b in pc.boxes:
            if b.fill is None:
                continue
            self.fills.append((*self.view.rect(b.x0, b.y0, b.x1, b.y1), b.fill))
        xs = [c.x0 for c in self.chars] + [r.a for r in hs] + [r.pos for r in vs]
        xe = [c.x1 for c in self.chars] + [r.b for r in hs] + [r.pos for r in vs]
        self.left = min(xs) if xs else 0.0
        self.right = max(xe) if xe else self.view.width
        self.spots: list[Spot] = []
        self.slivers: list[Cell] = []  # 字の入らない細い空のマス（数だけを窓に出す）
        # 欄にしない空き（マス, 向き, 大きさ pt）：字のあるマスの、字の下・上の 1 字ぶん以上の空き。窓に「［↓余白 n mm］」と出す
        self.room_marks: list[tuple] = []
        self._big_empty = [c for c in self._big_empty
                           if not any(c.x0 <= ch.cx <= c.x1 and c.y0 <= ch.cy <= c.y1 for ch in self.chars)
                           and not any(o.x0 >= c.x0 - 0.5 and o.x1 <= c.x1 + 0.5 and o.y0 >= c.y0 - 0.5 and o.y1 <= c.y1 + 0.5
                                       for o in self.cells)]
        self.drawn_text = self._drawn_text_regions()
        self._find_spots()
        if self.drawn_text:
            # 図形で描いた字の所：字の形の線や塗りからできたマス・空きは、書く所ではない
            m = 0.5 * max(self.typical, 10)  # 字の形の線（下線に見える横の線など）は、字の形の外へ少し出る
            # （中に字のある所は、飾りの形かもしれないので、欄の候補は消さない）
            self.spots = [k for k in self.spots if not any(
                _overlap(k.rect, (r[0] - m, r[1] - m, r[2] + m, r[3] + m)) > 0 for r, _n, mixed in self.drawn_text if not mixed)]
        self._circled()
        for k in self.spots:
            # 書き込みと同じ決め方で1字も入らない所も、消さずに示す（規則では決められない所・書けない所として）
            if not _holds_text(k):
                x0, y0, x1, y1 = k.rect
                _agent(k)
                k.readonly = f"小さい（幅 {(x1 - x0) / self.view.mm:.1f} mm・高さ {(y1 - y0) / self.view.mm:.1f} mm）ので、1字も入りません"
        self._assign_ids()
        self.flags: list[dict] = []  # 読む段階に関わる注意：{grade, title, detail, where}
        self._read_flags()

    def _drawn_text_regions(self) -> list[tuple]:
        """字を図形（曲線を含む塗りの形）にして描いたかもしれない所：字の大きさほどの（ページのふつうの字の大きさの 3 倍まで）
        小さな塗りの形のまとまりで、字の情報のないもの。いくつ並べば字かは決めない（1つでも、字の形かもしれない）。
        近くのまとまり（字3つぶんの内）は1つの所にまとめて、知らせる所が多くなりすぎないようにする。
        返すもの：(位置, 形の数, 中に字（字の情報）があるか)。字がある所も、飾りか字かは決めずに知らせる（エージェントが画像で決める）。"""
        typ = max(self.typical, 10)
        shapes = []
        for x0, y0, x1, y1, stroke, fill in self.pc.curves:
            if not fill:
                continue
            r = self.view.rect(x0, y0, x1, y1)
            w, h = r[2] - r[0], r[3] - r[1]
            if 0.3 <= max(w, h) <= 3 * typ:
                shapes.append(r)
        if not shapes:
            return []

        def cluster(items: list, gap: float) -> list[list]:
            groups: list[list] = []
            for r, n in sorted(items):
                hit = [g for g in groups
                       if g[0][0] - gap <= r[2] and r[0] <= g[0][2] + gap and g[0][1] - gap <= r[3] and r[1] <= g[0][3] + gap]
                if hit:
                    g = hit[0]
                    for o in hit[1:]:
                        g[0] = (min(g[0][0], o[0][0]), min(g[0][1], o[0][1]), max(g[0][2], o[0][2]), max(g[0][3], o[0][3]))
                        g[1] += o[1]
                        groups.remove(o)
                    g[0] = (min(g[0][0], r[0]), min(g[0][1], r[1]), max(g[0][2], r[2]), max(g[0][3], r[3]))
                    g[1] += n
                else:
                    groups.append([r, n])
            return groups

        # 字の間・行の間（字1字ぶんほど）でつながる形を1つのまとまりに。まとまりどうしが字3つぶんの内なら、同じ所にまとめる
        groups = cluster([(r, 1) for r in shapes], 1.2 * typ)
        while True:
            merged = cluster([(tuple(g[0]), g[1]) for g in groups], 3 * typ)
            if len(merged) == len(groups):
                break
            groups = merged
        out = []
        for box, n in groups:
            mixed = any(box[0] <= c.cx <= box[2] and box[1] <= c.cy <= box[3] for c in self.chars if not c.text.isspace())
            out.append((box, n, mixed))
        return out

    # ------------------------------------------------------------------ ルビ
    def _ruby(self, lines: list[Line]) -> list[Line]:
        """ルビ（振り仮名）：すぐ上にある、字の大きさが基の字の半分以下の横書きの行で、横の範囲が基の行の中に収まるもの
        （ルビの字は基の字の半分の大きさが決まり：JIS X 4051）。基の字の後ろに「（かつしか）」の形でつける（Word の窓と同じ）。
        ルビの行は、行としては出さない（見出しにも使わない）。"""
        hor = [ln for ln in lines if not ln.vertical and ln.text.strip()]
        if len(hor) < 2:
            return lines
        used: set[int] = set()
        for r in sorted(hor, key=lambda ln: ln.size):
            if id(r) in used:
                continue
            best = None
            for b in hor:
                if b is r or id(b) in used or r.size > 0.5 * b.size * 1.05:
                    continue
                if _SPACED.fullmatch(b.text.strip(" \u3000")) and re.search(r"[ \u3000]", b.text.strip(" \u3000")):
                    continue  # 字間を空けた見出し（「氏　名」）の上の小さな字は、ルビでなく見出し（「フリ　ガナ」）
                if not any(_ruby_base(ch) for ch in b.text):
                    continue  # ルビは漢字（注音は漢字の読み）につくもの：漢字のない行の上の小さな字（「(office use)」）は別の行のまま
                gap = b.y0 - r.y1
                if not (-0.15 * b.size <= gap <= 0.5 * b.size):
                    continue
                if not (b.x0 - 0.25 * b.size <= r.x0 and r.x1 <= b.x1 + 0.25 * b.size):
                    continue
                if best is None or gap < best[0]:
                    best = (gap, b)
            if best is None:
                continue
            b = best[1]
            # ルビの行の、空白で分かれたまとまりごとに（「やまだ　はなこ」→「山田」「花子」）、横の範囲に中心のある基の字の後ろへ
            groups = []
            for m in re.finditer(r"[^ \u3000]+", r.text):
                gx0, _gy0, gx1, _gy1 = r.span_rect(m.start(), m.end())
                groups.append((m[0], gx0, gx1))
            plan = []
            for ruby, gx0, gx1 in groups:
                idx = [k for k, q in enumerate(b.pos) if gx0 - 0.5 <= (q[0] + q[1]) / 2 <= gx1 + 0.5 and not b.text[k].isspace()]
                if not idx:
                    rc = (gx0 + gx1) / 2
                    idx = [min(range(len(b.pos)), key=lambda k: abs((b.pos[k][0] + b.pos[k][1]) / 2 - rc))]
                plan.append((max(idx) + 1, ruby))
            for at, ruby in sorted(plan, reverse=True):
                ins = "（" + ruby + "）"
                q = b.pos[at - 1]
                b.text = b.text[:at] + ins + b.text[at:]
                b.pos = b.pos[:at] + [(q[1], q[1])] * len(ins) + b.pos[at:]
                self.added += ["（", "）"]
            self.ruby_merges.append((b.text, r.text.strip(" \u3000"), (r.x0, r.y0, r.x1, r.y1)))
            b.chars = sorted(b.chars + r.chars, key=lambda c: c.x0)
            b.y0 = min(b.y0, r.y0)
            b.ruby = getattr(b, "ruby", []) + [g[0] for g in groups]
            used.add(id(r))
        return [ln for ln in lines if id(ln) not in used]

    # ------------------------------------------------------------------ 読む段階の注意
    def all_lines(self) -> list[Line]:
        return [ln for c in self.cells for ln in c.lines] + list(self.free_lines)

    def where(self, x0: float, y0: float, text: str = "") -> str:
        t = text.strip(" \u3000")
        return f"{self.number}ページ（左 {x0:.0f}・上 {y0:.0f} pt）" + (f"「{t[:24]}」" if t else "")

    def flag(self, grade: str, title: str, detail: str, where: list[str]) -> None:
        for f in self.flags:
            if f["title"] == title:
                f["where"] += [w for w in where if w not in f["where"]]
                return
        self.flags.append({"grade": grade, "title": title, "detail": detail, "where": list(where)})

    def _read_flags(self) -> None:
        legacy, pua = [], []
        for ln in self.all_lines():
            if any(c.legacy for c in ln.chars):
                legacy.append(self.where(ln.x0, ln.y0, ln.text))
            if _is_pua(ln.text):
                pua.append(self.where(ln.x0, ln.y0, ln.text))
        if legacy:
            self.flag("E", "字の符号が Unicode でない",
                      "Unicode 以前の書体（字の形を ASCII などの符号に割り当てたもの）で描いた字があります。窓の字は元の文字ではありません。"
                      "その位置の字は view_page で画像を見て読みます（この書体は、知られた書体の名前で見分けました。名前を知らない"
                      "書体は、見分けられないことがあります）", legacy)
        maybe = self._maybe_legacy()
        if maybe:
            self.flag("B", "字の符号が Unicode でない書体かもしれない",
                      "字の符号が Unicode でない書体かもしれない：書類の文字（ほかの字か、書類の言語 /Lang）はラテン文字でないのに、"
                      "埋め込んだ書体の字が、ラテン文字（ASCII など）だけになっています。Unicode 以前の書体（字の形を ASCII の符号に"
                      "割り当てたもの）なら、窓の字は元の文字ではありません。ふつうの欧文のこともあるので、その位置は view_page で"
                      "画像を見て、窓の字と同じかを確かめる（とどくは書体の名前を知らないので、決めていません）", maybe)
        if pua:
            self.flag("B", "外字（私用の符号の字）",
                      "外字：中身は view_page で画像を見て確かめる（私用の符号 U+E000〜F8FF などは、書類ごとに字の形が決まり、"
                      "字の意味は符号からは分かりません）", pua)
        if self.overprinted:
            many = [grp for grp in self.overprinted if len(grp) > 2]
            self.flag("C" if many else "B", "重なって描かれた行",
                      "違う字の行が同じ所に重ねて描かれています（人が見ても字が重なっています）。窓では、描いた順で別々の行に分けました"
                      + ("。3つ以上が重なる所は、分け方が正しいかを view_page で確かめる" if many else
                         "。どちらの行が生きているか（消した字か）は view_page で確かめる"),
                      [self.where(min(ln.x0 for ln in grp), min(ln.y0 for ln in grp), " ／ ".join(ln.text.strip() for ln in grp))
                       for grp in self.overprinted])
        ocr = [c for c in self.chars if c.invisible and not c.text.isspace()]
        if ocr:
            pics = [self.view.rect(p.x0, p.y0, p.x1, p.y1) for p in self.pc.pictures]
            on = [c for c in ocr if any(r[0] - 1 <= c.cx <= r[2] + 1 and r[1] - 1 <= c.cy <= r[3] + 1 for r in pics)]
            if on:
                self.flag("B", "OCR の字（画像の上の見えない字）",
                          "OCR の字はエージェントが画像と照らして確かめる（スキャンの画像の上に、文字の認識で作った見えない字が"
                          "重ねてあります。認識の誤りがあっても、字の情報からは分かりません）",
                          [f"{self.number}ページ（見えない字 {len(on)} 字）"])
        orphan = [self.where(ln.x0, ln.y0, ln.text) for ln in self.all_lines() if getattr(ln, "orphans", None)]
        if orphan:
            self.flag("C", "前の字のない結合の記号",
                      "母音の記号などの結合する記号が、基の字なしで行の頭や空白の後ろにあります。字の並びが正しくないおそれがあるので、"
                      "その位置は view_page で画像を見て確かめる", orphan)
        if self.drawn_text:
            # ページのほかに字があれば、窓はその所を除いて読める（外字と同じ B）。字がそれだけのページは、画像を見るほかない（E）
            has_text = any(not c.text.isspace() and not c.invisible for c in self.chars)
            self.flag("B" if has_text else "E", "図形で描いた字かもしれない所（字の情報がない）",
                      "図形で描いた字かもしれない所：view_page で確かめる。字の大きさほどの塗りの形（曲線を含む）があり、"
                      "そこには字の情報がないので、字だったとしても窓には出せません。字か飾り（記号・模様・ロゴ）かは、とどくは"
                      "決めません。注意書きなど大事な文のことがあるので、その位置は view_page でページの画像を見て読み、中身を推測で"
                      "埋めない（字の情報のない所の欄の候補は出していません。書く所なら、近くの字を目印にして書けます）",
                      [f"{self.number}ページ（左 {r[0]:.0f}・上 {r[1]:.0f}〜右 {r[2]:.0f}・下 {r[3]:.0f} pt、塗りの形 {n} 個"
                       + ("。中に字もある：飾りの形かもしれない" if mixed else "") + "）"
                       for r, n, mixed in self.drawn_text])
        if self.ruby_merges:
            # ルビ（振り仮名）とした小さな字の行：字の大きさと位置の決まり（JIS X 4051）でつないだが、小さな注の字のこともある
            self.flag("B", "ルビとして基の字につないだ字",
                      "すぐ上の小さな字の行を、ルビ（振り仮名）として基の字の後ろに「（）」でつなぎました（かっこは窓で足した字）。"
                      "小さな注の字・別の見出しのこともあるので、その位置は窓の字と view_page の画像で、ルビかを確かめる",
                      [self.where(r[0], r[1], f"{base.strip()[:16]} ← {ruby[:12]}") for base, ruby, r in self.ruby_merges])
        bad = [m for m in self.merges if m[1] <= 0.5]
        if bad:
            self.flag("C", "重ね書きとして1字にした字", "箱が重ならない字を、重ね書きとして1字にしました（読み落としのおそれ）",
                      [self.where(r[0], r[1], t) for t, _ov, r in bad])

    def _maybe_legacy(self) -> list[str]:
        """名前を知らない Unicode 以前の書体かもしれない書体：埋め込んだ書体（Type 3 を含む）で、このページで描いた字が
        ラテン文字だけ（字が1つ以上）のもの。書類の文字がラテン文字でないこと（このページのほかの書体の字の文字の種類か、
        Catalog の /Lang）が分かるときだけ。東アジアの文字（漢字・仮名・ハングル）の書類は含めない（fonts.LANG_SCRIPT）。
        返すもの：書体ごとに、はじめの行の位置と書体の名前。"""
        from .fonts import EAST_ASIAN, LANG_SCRIPT, letter_script

        glyphs = self.pc.glyphs
        by_font: dict[str, set] = {}
        first: dict[str, Ch] = {}
        embedded: dict[str, bool] = {}
        for c in self.chars:
            if c.legacy or c.invisible or not (0 <= c.seq < len(glyphs)):
                continue
            g = glyphs[c.seq]
            fb = g.fontbase or g.font
            for ch in c.text:
                sc = letter_script(ch)
                if sc:
                    by_font.setdefault(fb, set()).add(sc)
                    first.setdefault(fb, c)
            embedded[fb] = embedded.get(fb, False) or g.embedded
        others = {sc for scs in by_font.values() for sc in scs if sc != "LATIN"}
        if any(sc.startswith(EAST_ASIAN) for sc in others):
            return []
        lang = self.doc.resolve(self.doc.root.get("Lang")) if hasattr(self.doc, "root") else None
        from .objects import PdfString, decode_text

        lang = decode_text(lang) if isinstance(lang, (bytes, PdfString)) else ""
        want = LANG_SCRIPT.get(lang.split("-")[0].strip().lower(), "")
        out = []
        for fb, scs in by_font.items():
            if scs != {"LATIN"} or not embedded.get(fb):
                continue
            if not (others or want):
                continue
            c = first[fb]
            why = "ほかの字：" + "・".join(sorted(others)) if others else f"書類の言語 {lang}"
            out.append(self.where(c.x0, c.y0, "") + f" 書体 {fb.split('+')[-1]}（{why}）")
        return out

    def _circled(self) -> None:
        """ページに描かれた○（線の曲線）が、1 つの行の続いた字だけを囲んでいる所：self.circled に (行, 始め, 終わり, id) を。
        囲んでいるかは、○を書くとき・読み戻すときと同じ決め方（circle.encloses）。とどくが描いた○（印 TdkCircle）は、その形
        （楕円・四角）と id で、ほかは曲線だけで閉じた図形（pc.ovals）を、その外枠に内接する楕円で比べる。どの選択肢を選んだか（字の意味）は決めない。
        曲線と長さのある直線で閉じた図形（角の丸い四角：pc.maybe_ovals）が囲む字は、self.maybe_circled（○かもしれない：「？」）。"""
        from .circle import drawn, encloses

        self.circled = []
        self.maybe_circled = []
        shapes = []
        if not any(stroke for *_r, stroke, _f in self.pc.curves):
            return
        try:
            tdk = drawn(self.doc, self.number) if hasattr(self.doc, "page_content") else []
        except Exception:  # noqa: BLE001  ○の印を読めないときも、ほかの曲線は比べる
            tdk = []
        for i, kind, g, _lw in tdk:
            shapes.append((kind, g, i))
        # ○と比べるのは、曲線だけで閉じた図形（○・楕円：pc.ovals）。長さのある直線を含む曲線の閉じた図形（角の丸い四角：
        # pc.maybe_ovals）は、字を囲む○か枠か、形から言えないので、囲む字を self.maybe_circled（窓の「［○？］」）に分けて示す
        maybe: set[int] = set()  # shapes の番号
        for src, guess in ((getattr(self.pc, "ovals", []), False), (getattr(self.pc, "maybe_ovals", []), True)):
            for x0, y0, x1, y1, stroke, _fill in src:
                if not stroke:
                    continue
                r = self.view.rect(x0, y0, x1, y1)
                if r[2] - r[0] <= 0.5 or r[3] - r[1] <= 0.5:
                    continue
                if any(abs(r[0] - (g[0] - g[2] if k == "ellipse" else g[0])) < 0.5 and abs(r[1] - (g[1] - g[3] if k == "ellipse" else g[1])) < 0.5
                       for k, g, _i in shapes if _i):
                    continue  # とどくが描いた○（上で、形と id で比べた）
                if guess:
                    maybe.add(len(shapes))
                    shapes.append(("border", r, ""))  # 角の丸い四角：外枠の四角で比べる
                else:
                    shapes.append(("ellipse", ((r[0] + r[2]) / 2, (r[1] + r[3]) / 2, (r[2] - r[0]) / 2, (r[3] - r[1]) / 2), ""))
        if not shapes:
            return
        lines = [ln for ln in self.all_lines() if ln.text.strip()]
        for n, (kind, g, ident) in enumerate(shapes):
            inner = {id(c) for c in self.chars if not c.text.isspace() and encloses(kind, g, c)}
            if not inner:
                continue
            host = [ln for ln in lines if any(id(c) in inner for c in ln.chars)]
            if len(host) != 1 or sum(1 for c in host[0].chars if id(c) in inner) != len(inner):
                continue  # 2 つ以上の行にかかる（表の枠の角の丸みなど）：字を囲む○とは言えない
            ln = host[0]
            cs = [c for c in ln.chars if id(c) in inner]
            if ln.vertical:
                lo, hi = min(c.y0 for c in cs), max(c.y1 for c in cs)
            else:
                lo, hi = min(c.x0 for c in cs), max(c.x1 for c in cs)
            if any(id(c) not in inner and not c.text.isspace() and lo < (c.cy if ln.vertical else c.cx) < hi for c in ln.chars):
                continue  # 囲んだ字の間に、囲まれていない字がある
            idx = [k for k, q in enumerate(ln.pos) if not ln.text[k].isspace() and lo - 0.01 <= (q[0] + q[1]) / 2 <= hi + 0.01]
            if idx:
                (self.maybe_circled if n in maybe else self.circled).append((ln, min(idx), max(idx) + 1, ident))

    def cell_fill(self, c: Cell):
        """マスを塗っている色（いちばん上に描かれたもの）。なければ None。"""
        col = None
        for x0, y0, x1, y1, fill in self.fills:
            if x0 <= c.x0 + 1 and y0 <= c.y0 + 1 and x1 >= c.x1 - 1 and y1 >= c.y1 - 1:
                if (x1 - x0) * (y1 - y0) < 0.5 * self.view.width * self.view.height:
                    col = fill
        return col

    # ------------------------------------------------------------------ 書く所を見つける（決め方は「書く所（場所）」の節に）
    def _find_spots(self) -> None:
        self._index_text()
        # 3・4) 行の中の空き・字の形の空欄・箱の字（マスの中の行と、マスの外の行）
        for cell, lines in [(c, c.lines) for c in self.cells] + [(None, self.free_lines)]:
            for ln in lines:
                self._line_spots(ln, cell)
        # 1) 空のマス（と、同じ大きさの空のマスの並び）
        self._cell_spots()
        # 5) 下線の罫線
        self._underline_spots()
        # 欄にしない空き（マスの中の字の下・上の、1 字ぶん以上の空き）：窓の印「［↓余白 n mm］」で示す事実
        self._room_facts()
        # 重なる所：小さいほうの面積の 6 割（チェックの印どうしは 7 割）より多く重なる所は、同じ所の別の読み方（alternative_to）。
        # 両方に書くと字が重なるので、書き込みはどちらか一方だけを受ける（先に見つけた所を主にする）。重なりは測った事実で、
        # どちらが正しい読み方かは決めない
        kept = _RectIndex()
        for s in self.spots:
            chk = s.kind == "checkbox"
            lim = 0.7 if chk else 0.6
            k = next((k for k in kept.touching(s.rect) if (k.kind == "checkbox") == chk and _overlap(s.rect, k.rect) > lim), None)
            if k is not None:
                s.alternative_to = k
                continue
            kept.add(s)
        # ページの外：見えるのはページの枠（CropBox を MediaBox に切り詰めたもの）の中だけ。欄の四角を枠に切り詰め、
        # 枠の外にしかないもの（ページの外に描いた線・字から見つけたもの）は欄にしない
        self._clip_to_page()
        # 斜線で消したマスは欄にしない。印刷された字を含む欄は、字のない所に縮める（縮めると書けない大きさになるものは、欄にしない）
        self.spots = [sp for sp in self.spots if not self._struck(sp.rect) and self._clear_of_print(sp)]

    def _mm(self, v: float) -> float:
        return round(v / self.view.mm, 1)

    def _rect_mm(self, r) -> list[float]:
        return [self._mm(v) for v in r]

    def _index_text(self) -> None:
        """周りの字を探すための、字のあるもの（字のあるマス・マスの外の行）の一覧と、縦・横の帯ごとの索引。"""
        items = []
        for c in self.cells:
            if c.visible():
                ls = [ln for ln in c.lines if ln.text.strip()]
                items.append((c.x0, c.y0, c.x1, c.y1, join_lines(c.lines), max((ln.size for ln in ls), default=self.typical),
                              _rtl_cell(c)))
        for ln in self.free_lines:
            if ln.text.strip():
                items.append((ln.x0, ln.y0, ln.x1, ln.y1, ln.text.strip(), ln.size, bool(getattr(ln, "rtl", False))))
        self._items = items
        B = 24.0
        self._yb: dict[int, list[int]] = {}
        self._xb: dict[int, list[int]] = {}
        self._y_wide: list[int] = []  # ページより何桁も大きな宣言の字：帯に置かず、いつも見る
        self._x_wide: list[int] = []
        for i, it in enumerate(items):
            a, b = int(it[1] // B), int(it[3] // B)
            if b - a > WIDE_BUCKETS:
                self._y_wide.append(i)
            else:
                for k in range(a, b + 1):
                    self._yb.setdefault(k, []).append(i)
            a, b = int(it[0] // B), int(it[2] // B)
            if b - a > WIDE_BUCKETS:
                self._x_wide.append(i)
            else:
                for k in range(a, b + 1):
                    self._xb.setdefault(k, []).append(i)
        self._B = B
        self._line_of = {id(ch): ln for ln in self.all_lines() for ch in ln.chars}

    def _neighbours(self, r) -> dict:
        """所の四角 r の、左・上・右・下でいちばん近い字（字のあるマスの字・マスの外の行の字。元のまま）と、その離れ（mm）。
        左右は上下の範囲が重なるもの、上下は左右の範囲が重なるものだけ。どこまで離れていても示す（近いかはエージェントが見る）。"""
        x0, y0, x1, y1 = r
        B = self._B
        best: dict[str, tuple] = {}
        rows = _in_bands(self._yb, self._y_wide, int(y0 // B), int(y1 // B))
        for i in rows:
            it = self._items[i]
            if min(y1, it[3]) - max(y0, it[1]) <= 0:
                continue
            if it[2] <= x0 + RULE_TOL:
                side, d = "left", x0 - it[2]
            elif it[0] >= x1 - RULE_TOL:
                side, d = "right", it[0] - x1
            else:
                continue
            if side not in best or d < best[side][0]:
                best[side] = (max(d, 0.0), it)
        cols = _in_bands(self._xb, self._x_wide, int(x0 // B), int(x1 // B))
        for i in cols:
            it = self._items[i]
            if min(x1, it[2]) - max(x0, it[0]) <= 0:
                continue
            if it[3] <= y0 + RULE_TOL:
                side, d = "above", y0 - it[3]
            elif it[1] >= y1 - RULE_TOL:
                side, d = "below", it[1] - y1
            else:
                continue
            if side not in best or d < best[side][0]:
                best[side] = (max(d, 0.0), it)
        return {side: {"text": it[4], "mm": self._mm(d), "_size": it[5], "_rtl": it[6]} for side, (d, it) in best.items()}

    @staticmethod
    def _label_of(nb: dict) -> tuple[str, str, str]:
        """見出し（list_fields・本人に伝える一覧の名前）：左の字（右から左へ書く表では右の字）、なければ上、なければ右の字。
        どれを見出しにするかは、見出しの名前のための決まった順で、書く所かどうかの判断には使わない（周りの字はすべて facts に）。"""
        left = nb.get("left")
        right = nb.get("right")
        if right and right["_rtl"] and not (left and left["_rtl"]):
            left, right = right, left
        rl = label_text(left["text"]) if left else ""
        cl = label_text(nb["above"]["text"]) if nb.get("above") else ""
        return (rl or cl or (label_text(right["text"]) if right else "")), rl, cl

    def default_size(self, nb: dict | None = None, line: "Line | None" = None) -> tuple[float, str]:
        """書く字の既定の大きさ（ひとつの決まり）：行の中の所はその行の字の大きさ。行のない所（空のマス・下線の罫線）は、
        見出しにした周りの字（左、なければ上、なければ右）の大きさ。周りに字がなければ、ページのふつうの字の大きさ（字の大きさの
        中央の値）。所に入らないときは、書き込みが入る大きさに縮めて、縮めたことを知らせる（fill）。エージェントは size で変えられる。"""
        if line is not None:
            return line.size, "行の字の大きさ"
        for side, name in (("left", "左の字"), ("above", "上の字"), ("right", "右の字")):
            if nb and nb.get(side):
                return float(nb[side]["_size"]), f"{name}の大きさ"
        return self.typical, "ページのふつうの字の大きさ"

    def _facts_nb(self, nb: dict) -> dict:
        return {side: {"text": v["text"], "mm": v["mm"]} for side, v in nb.items()}

    def _cell_spots(self) -> None:
        """空のマス：1 つずつの所（字か印を書く。kind_candidates に text・checkbox）と、同じ大きさの空のマスが横に隣り合って
        並ぶもの（2 つ以上）の並び全体（comb：1 字ずつ書くか、1 つの欄として書くか。「？」）。大きさ・塗り・周りの字は facts に。"""
        empties = [c for c in self.cells if not c.visible()] + list(self._big_empty)
        cells = []
        for c in empties:
            if c.w < MIN_GLYPH or c.h < MIN_GLYPH:
                # 書き込みのいちばん小さな字（4pt）も入らない細いマス（枠と枠のすき間）：欄にはせず、ページごとの数を窓に出す
                self.slivers.append(c)
                continue
            cells.append(c)
        # 同じ大きさの空のマスが横に隣り合って並ぶもの（大きさ・隣り合いは、マスを作るときと同じ許し RULE_TOL で比べる）
        rows: list[list[Cell]] = []
        for c in sorted(cells, key=lambda c: (c.y0, c.x0)):
            if rows and abs(c.y0 - rows[-1][0].y0) <= RULE_TOL:
                rows[-1].append(c)
            else:
                rows.append([c])
        combs: list[Spot] = []
        for row in rows:
            row.sort(key=lambda c: c.x0)
            run: list[Cell] = []
            for c in row + [None]:
                if c is not None and run and abs(c.x0 - run[-1].x1) <= RULE_TOL and abs(c.w - run[0].w) <= RULE_TOL \
                        and abs(c.h - run[0].h) <= RULE_TOL and abs(c.y1 - run[0].y1) <= RULE_TOL:
                    run.append(c)
                    continue
                if len(run) >= 2:
                    rect = (run[0].x0, run[0].y0, run[-1].x1, run[-1].y1)
                    nb = self._neighbours(rect)
                    label, rl, cl = self._label_of(nb)
                    size, src = self.default_size(nb)
                    sp = Spot("comb", rect, label, "同じ大きさの空のマスの並び", [(x.x0, x.y0, x.x1, x.y1) for x in run],
                              row_label=rl, column_label=cl, size=size, host=run[0], cell=run[0],
                              kind_candidates=["comb", "text"], shape="cells", anchor=(run[0].x0, run[0].y0),
                              write_size_from=src)
                    sp.facts = {"rect_mm": self._rect_mm(rect), "cells": len(run),
                                "cell_mm": [self._mm(run[0].w), self._mm(run[0].h)], **self._facts_nb(nb)}
                    combs.append(_agent(sp))
                run = [c] if c is not None else []
        self.spots += combs
        for c in cells:
            rect = (c.x0, c.y0, c.x1, c.y1)
            nb = self._neighbours(rect)
            label, rl, cl = self._label_of(nb)
            size, src = self.default_size(nb)
            sp = Spot("text", rect, label, "空のマス", row_label=rl, column_label=cl, cell=c, size=size, host=c,
                      kind_candidates=["text", "checkbox"], shape="cell", anchor=(c.x0, c.y0), write_size_from=src)
            facts = {"rect_mm": self._rect_mm(rect), "size_mm": [self._mm(c.w), self._mm(c.h)]}
            fill = self.cell_fill(c)
            if fill is not None:
                facts["fill"] = "#" + "".join(f"{max(0, min(255, round(v * 255))):02X}" for v in list(fill)[:3])
            inner = sum(1 for o in self.index.near(c.x0, c.y0, c.x1, c.y1) if o is not c and o.x0 >= c.x0 - 0.5
                        and o.x1 <= c.x1 + 0.5 and o.y0 >= c.y0 - 0.5 and o.y1 <= c.y1 + 0.5)
            if inner:
                facts["inner_cells"] = inner  # 中にある、もっと小さなマスの数（入れ物のマス）
            sp.facts = {**facts, **self._facts_nb(nb)}
            self.spots.append(sp)

    def _chars_at_height(self, y: float, d: float) -> list:
        """字の中ほどの高さが y ± d の字（ページの字の順のまま）。高さの索引で引く（行ごとにページのすべての字をたどらない）。"""
        import bisect

        idx = self.__dict__.get("_cy_index")
        if idx is None or idx[2] != len(self.chars):
            order = sorted(range(len(self.chars)), key=lambda i: self.chars[i].cy)
            idx = self.__dict__["_cy_index"] = ([self.chars[i].cy for i in order], order, len(self.chars))
        keys, order, _n = idx
        return [self.chars[i] for i in sorted(order[bisect.bisect_left(keys, y - d):bisect.bisect_right(keys, y + d)])]

    def _bounds(self, ln: Line, cell: Cell | None) -> tuple[tuple, tuple]:
        """行の左右の端：(位置, 種類, その先の字)。種類は "cell"（マスの縁）・"text"（同じ高さのほかの字）・"ruling"（縦の罫線）・
        "page"（ページの字の端）。"""
        sz = ln.size
        if cell is not None:
            left, right = (cell.x0, "cell", ""), (cell.x1, "cell", "")
            for o in cell.lines:
                if o is ln or o.vertical or not o.text.strip(" 　") or not (o.y0 < ln.y1 - 0.2 * sz and o.y1 > ln.y0 + 0.2 * sz):
                    continue
                if o.x1 <= ln.x0 + 0.5 and o.x1 > left[0]:
                    left = (o.x1, "text", o.text.strip())
                if o.x0 >= ln.x1 - 0.5 and o.x0 < right[0]:
                    right = (o.x0, "text", o.text.strip())
            return left, right
        left, right = (self.left, "page", ""), (self.right, "page", "")
        own = {id(c) for c in ln.chars}
        for c in self._chars_at_height((ln.y0 + ln.y1) / 2, 0.6 * sz):
            if c.text.isspace() or id(c) in own or abs(c.cy - (ln.y0 + ln.y1) / 2) >= 0.6 * sz:
                continue
            if c.x1 <= ln.x0 + 0.5 and c.x1 > left[0]:
                o = self._line_of.get(id(c))
                left = (c.x1, "text", o.text.strip() if o is not None else c.text)
            if c.x0 >= ln.x1 - 0.5 and c.x0 < right[0]:
                o = self._line_of.get(id(c))
                right = (c.x0, "text", o.text.strip() if o is not None else c.text)
        for r in self.vs:
            if r.a - 1 <= ln.y0 and r.b + 1 >= ln.y1:
                if left[0] < r.pos <= ln.x0:
                    left = (r.pos, "ruling", "")
                if ln.x1 <= r.pos < right[0]:
                    right = (r.pos, "ruling", "")
        return left, right

    def _block_margin(self, ln: Line, cell: Cell | None) -> bool:
        """行の頭の前の空きが、組版の空きか：
        - 段落の 1 行目の字下げ（JIS X 4051：段落の頭は 1 字下げる）：すぐ下の行（行の間が 1 字ぶんまで）の頭より、この行の頭が
          読む順の後ろへずれていて、そのずれを 1 字ぶん（1 em）の数に丸めると 1（行の空白を作るときと同じ丸め）
        - すぐ上か下の行（行の間が 1 字ぶんまで）と行の頭がそろう（ずれを 1 em の数に丸めると 0）：段落・字の列の左の余白
        呼ぶのは、頭の前に描かれた端（マスの縁・縦の罫線・同じ高さの字）がない行（空きをページの字の端まで測った行）だけ"""
        sz = ln.size
        rtl = bool(getattr(ln, "rtl", False))
        for o in (cell.lines if cell is not None else self.free_lines):
            if o is ln or o.vertical or not o.text.strip(" \u3000"):
                continue
            below = 0 <= o.y0 - ln.y1 <= 1.0 * sz
            above = 0 <= ln.y0 - o.y1 <= 1.0 * sz
            if not (below or above):
                continue
            d = (o.x1 - ln.x1) if rtl else (ln.x0 - o.x0)
            if below and round(d / sz) == 1 and not (o.x1 < ln.x0 or o.x0 > ln.x1):
                return True
            if round(d / sz) == 0:
                return True
        return False

    def _line_spots(self, ln: Line, cell: Cell | None) -> None:
        text = ln.text
        # 4) チェックの箱の字（どの箱の字も同じ扱い）・記入済みの箱の字（書き換えられない所）
        for i, ch in enumerate(text):
            boxed = ch in BOX_GLYPHS
            if not boxed and ch not in CHECKED_GLYPHS:
                continue
            after = _option_text(text[i + 1:])
            before = _option_text(text[:i][::-1])[::-1] if i else ""
            x0, y0, x1, y1 = ln.span_rect(i, i + 1)
            sp = Spot("checkbox", (x0, y0, x1, y1), label_text(after or before) or "（選択肢の文字なし）", "箱の字",
                      line_text=text, cell=cell, size=ln.size, host=ln, at=(i, i + 1), shape="box", anchor=(x0, ln.y0))
            sp.facts = {"rect_mm": self._rect_mm((x0, y0, x1, y1)), "char": ch, "line": text.strip(" 　"),
                        "char_name": f"U+{ord(ch):04X} {unicodedata.name(ch, '')}", "after": after, "before": before}
            # 箱の字は、チェックの箱か箇条の印かを字の形では決められない（どの箱の字も「？」：表計算・Word の窓と同じ）
            _agent(sp)
            if not boxed:
                sp.checked = True
                sp.shape = "checked"
                sp.reason = "記入済みの箱の字"
                sp.readonly = f"印刷された記入済みの印「{ch}」は消せないので、書き換えられません"
                _agent(sp)
            self.spots.append(sp)
        if ln.vertical and len(text.strip(" \u3000")) > 1:
            return  # 縦書きの行（1 字だけの行は、向きが決まらないので横書きと同じに見る）
        # 2・3) 行の中の空き・字の形の空欄と、行の頭の前・行の終わりの後ろの空き（1 字ぶん以上）
        sz = ln.size
        (lb, lk, lt), (rb, rk, rt) = self._bounds(ln, cell)
        pad = {"cell": EDGE_PAD, "ruling": EDGE_PAD, "text": TEXT_PAD * sz, "page": 0.0}
        n = len(text)
        if not text.strip(" \u3000"):
            # 空白の字だけの行（字の形のない字を描いた行）：端から端までの空き（1 字ぶん以上のとき）。字の見えないマスの中の行は、
            # マスそのものが空のマスの所なので、行の所にはしない
            if cell is not None and not cell.visible():
                return
            x0, x1 = lb + pad[lk], rb - pad[rk]
            if wide_gap(rb - lb, sz) and x1 > x0:
                y0, y1, base = line_band(ln)
                sp = _agent(Spot("text" if cell is not None else "inline", (x0, y0, x1, y1), "［　］", GAP_WHY["blank_line"], line_text=text, blank=(0, n),
                                 cell=cell, size=sz, host=ln, at=(0, n), shape="blank_line", anchor=(lb, ln.y0),
                                 write_size_from="行の字の大きさ", baseline=base))
                sp.facts = {"rect_mm": self._rect_mm(sp.rect), "width_mm": self._mm(rb - lb), "width_em": round((rb - lb) / sz, 1),
                            "bound": {"kind": f"{lk}・{rk}", "text": " ／ ".join(t for t in (lt, rt) if t), "mm": 0.0},
                            "room_mm": self._room(ln, cell, x0, x1)}
                self.spots.append(sp)
            return
        rtl = bool(getattr(ln, "rtl", False))
        left_room, right_room = ln.x0 - lb, rb - ln.x1
        head_room, end_room = (right_room, left_room) if rtl else (left_room, right_room)
        lead = 1 if wide_gap(head_room, sz) else 0
        trail = 1 if wide_gap(end_room, sz) else 0
        if lead and (rk if rtl else lk) == "page" and self._block_margin(ln, cell):
            lead = 0  # 頭の前に描かれた端のない行の、段落の 1 字下げ・段落の左の余白：組版の空き
        padded = "\u3000" * lead + text + "\u3000" * trail

        def width(a: int, b: int) -> float:
            if lead and b <= lead:
                return head_room / sz
            if trail and a >= lead + n:
                return end_room / sz
            aa, bb = max(a - lead, 0), min(b - lead, n)
            if all(_chars.is_space(ch) for ch in text[aa:bb]):
                return _em_width(text, aa, bb)  # 行の空白は測った空きから作った字（1 em 以上の空きだけが 1 字ぶん以上になる）
            x0, _, x1, _ = ln.span_rect(aa, bb)
            return (x1 - x0) / sz

        # 中央そろえ（左右の空きが等しい）は、両側に描かれた端（マスの縁・罫線・字）があるときだけ見る：ページの字の端までの空きは、
        # 描かれた端ではないので、ページの中ほどに置いた見出し（「(申請者)住所」）を中央そろえとは言えない
        align = lk != "page" and rk != "page"
        band0, band1, base = line_band(ln)  # 所は行のもの：書く四角は行の帯（基線と字の高さ）。上下の空きは事実（room_mm）に
        for blank in gap_blanks(padded, lead, trail, width, align=align):
            a, b = blank
            a0, b0 = a - lead, b - lead
            why = blank.why
            on_left = False
            if why in ("lead", "trail"):
                at = (0, 0) if why == "lead" else (n, n)
                on_left = (why == "lead") != rtl
                if on_left:
                    x0, x1 = lb + pad[lk], ln.x0 - TEXT_PAD * sz
                    anchor, bound = (lb, ln.y0), {"kind": lk, "text": lt, "mm": self._mm(left_room)}
                else:
                    x0, x1 = ln.x1 + TEXT_PAD * sz, rb - pad[rk]
                    anchor, bound = (ln.x1, ln.y0), {"kind": rk, "text": rt, "mm": self._mm(right_room)}
                if x1 <= x0:
                    continue
                rect = (x0, band0, x1, band1)
                if any(k.cell is cell and k.shape in ("lead", "trail") and all(abs(u - v) <= 0.5 for u, v in zip(k.rect, rect))
                       for k in self.spots):
                    continue  # 同じマスのほかの行の前後の空きと同じ四角（縦に並ぶ行の左の、同じ空き）：同じ所は 1 つ
            else:
                at = (max(a0, 0), min(b0, n))
                if at[1] <= at[0]:
                    continue
                x0, _, x1, _ = ln.span_rect(*at)
                rect, anchor, bound = (x0, band0, x1, band1), (x0, ln.y0), None
                if why == "placeholder":
                    # PDF の字は消せない：伏せ字の上を白い下地の欄で覆って書く（元の字は残る）。覆う四角は字の箱より少し大きく
                    rect = (x0 - 0.3, band0, x1 + 0.3, band1)
            if rect[2] <= rect[0]:
                continue  # 空きがまったくない（縁や次の字に接している）
            orig = text[at[0]:at[1]]
            shown = text[:at[0]] + (orig if orig.strip(" 　") else "") + "［　］" + text[at[1]:]
            if why == "placeholder":
                shown = text[:at[0]] + "［　］" + text[at[1]:]
            # マスの中の行の前後の空きは字の欄（text）として書く。行の中の空きは inline（どちらも行の帯に、行の基線で書く）
            kind = "text" if why in ("lead", "trail") and cell is not None else "inline"
            sp = Spot(kind, rect, label_text(shown), GAP_WHY.get(why, "空き"), line_text=text, blank=at, cell=cell,
                      size=sz, host=ln, at=at, shape=why, anchor=anchor, write_size_from="行の字の大きさ",
                      cover=orig if why == "placeholder" else "", baseline=base)
            if not blank.structural:
                _agent(sp)
            # 幅は測った空き（行の前後の空きは、端から字までの空き。書く所の四角は、端と字から少し離したもの）
            w = (left_room if on_left else right_room) if why in ("lead", "trail") else rect[2] - rect[0]
            facts = {"rect_mm": self._rect_mm(rect), "width_mm": self._mm(w), "width_em": round(w / sz, 1),
                     "before": text[:at[0]].rstrip(" 　")[-12:], "after": text[at[1]:].lstrip(" 　")[:12],
                     "line": text.strip(" 　"), "room_mm": self._room(ln, cell, rect[0], rect[2])}
            if why in ("connector", "leader", "dash", "placeholder"):
                facts["glyph"] = orig[:1]
                facts["count"] = len(orig)
            if bound is not None:
                facts["bound"] = bound
            if why == "gap":
                core = text.strip(" 　")
                if evenly_spaced(core):
                    facts["spaced"] = True  # 行の字が 1 字ずつ同じ空白でつながる（字間を空けた見出しの形「氏　　名」と同じ形）
            sp.facts = facts
            self.spots.append(sp)

    def _room(self, ln: Line, cell: Cell | None, x0: float, x1: float) -> dict:
        """行の所（x0〜x1）の上下の空き（mm）：その横の範囲にかかる上下のほかの行の字・横の罫線まで（マスの中はマスの縁まで）。
        書く四角にはしない事実（広い所に書くなら box か where: below で）。上下に何もなければ、ページの端まで。"""
        top, bottom = (cell.y0, cell.y1) if cell is not None else (0.0, self.view.height)
        y0, y1 = ln.y0, ln.y1
        for o in (cell.lines if cell is not None else self.free_lines):
            if o is ln or not o.text.strip(" \u3000") or min(o.x1, x1) - max(o.x0, x0) <= 0:
                continue
            if o.y1 <= (y0 + y1) / 2:
                top = max(top, o.y1)
            elif o.y0 >= (y0 + y1) / 2:
                bottom = min(bottom, o.y0)
        if cell is None:
            for h in self.hs:
                if min(h.b, x1) - max(h.a, x0) <= 0:
                    continue
                if h.pos <= y0:
                    top = max(top, h.pos)
                elif h.pos >= y1:
                    bottom = min(bottom, h.pos)
        return {"above": self._mm(max(0.0, y0 - top)), "below": self._mm(max(0.0, bottom - y1))}

    def _room_facts(self) -> None:
        """字のあるマスの、字の下・上の空き（1 字ぶん以上）：欄にはせず、窓の印「［↓余白 n mm］」「［↑余白 n mm］」で示す
        （書くときは字を目印に {"after": …, "where": "below"} か、{"box": [左, 上, 右, 下]}（mm））。字は行の字の大きさで測る。"""
        for c in self.cells:
            ink = [x for x in c.lines if x.text.strip(" 　")]
            if not ink:
                continue
            meas = [x for x in ink if not x.vertical] or ink
            sz = max(x.size for x in meas)
            below = c.y1 - max(x.y1 for x in meas)
            above = min(x.y0 for x in meas) - c.y0
            if wide_gap(below, sz):
                self.room_marks.append((c, "↓", below))
            if wide_gap(above, sz):
                self.room_marks.append((c, "↑", above))

    def _struck(self, r) -> bool:
        """向かい合う角と角を結ぶ斜めの線がある（斜線で消した）マス。"""
        if not self.diagonals:
            return False
        x0, y0, x1, y1 = r
        tol = max(2.0, 0.12 * min(x1 - x0, y1 - y0))
        pairs = (((x0, y0), (x1, y1)), ((x1, y0), (x0, y1)))

        def near(p, q):
            return abs(p[0] - q[0]) <= tol and abs(p[1] - q[1]) <= tol

        for a, b in self.diagonals:
            for p, q in pairs:
                if (near(a, p) and near(b, q)) or (near(a, q) and near(b, p)):
                    return True
        return False

    def _clear_of_print(self, sp: Spot) -> bool:
        if sp.kind not in ("text", "inline", "space") or sp.cover:
            return True
        x0, y0, x1, y1 = sp.rect
        own = {id(c) for c in sp.host.chars} if isinstance(sp.host, Line) else set()  # 欄の行の字（下線の「＿」など）は欄の一部
        inside = [c for c in self.chars if not c.text.isspace() and id(c) not in own and x0 < c.cx < x1 and y0 < c.cy < y1]
        if not inside:
            return True
        # 上下の空き（字の帯のあいだ）と、左右の空きのうち、いちばん広い所
        need = max(MIN_WRITE, 0.9 * sp.size)
        best = None
        ys = sorted((c.y0, c.y1) for c in inside)
        edge = y0
        for a, b in ys + [(y1, y1)]:
            if a - edge >= need and (best is None or (a - edge) * (x1 - x0) > _area(best)):
                best = (x0, edge + 0.3, x1, a - 0.3)
            edge = max(edge, b)
        xs = sorted((c.x0, c.x1) for c in inside)
        edge = x0
        for a, b in xs + [(x1, x1)]:
            if a - edge >= max(need, 2 * sp.size) and y1 - y0 >= need and (best is None or (a - edge) * (y1 - y0) > _area(best)):
                best = (edge + 0.3, y0, a - 0.3, y1)
            edge = max(edge, b)
        if best is None:
            return False
        sp.rect = best
        sp.reason += "（印刷された字のない所に縮めた）"
        if sp.facts:
            sp.facts["rect_mm"] = self._rect_mm(best)
        return True

    def _clip_to_page(self) -> None:
        W, H = self.view.width, self.view.height

        def clip(r):
            x0, y0, x1, y1 = max(r[0], 0.0), max(r[1], 0.0), min(r[2], W), min(r[3], H)
            return (x0, y0, x1, y1) if x1 > x0 and y1 > y0 else None

        kept: list[Spot] = []
        for s in self.spots:
            r = clip(s.rect)
            if r is None:
                continue
            s.rect = r
            if s.boxes:
                s.boxes = [b2 for b2 in (clip(b) for b in s.boxes) if b2 is not None]
                if not s.boxes:
                    continue
            if getattr(s, "options", None):
                s.options = [(o, r2) for o, r2 in ((o, clip(r0)) for o, r0 in s.options) if r2 is not None]
                if not s.options:
                    continue
            kept.append(s)
        gone = {id(s) for s in self.spots} - {id(s) for s in kept}
        if gone:
            for s in kept:
                if s.alternative_to is not None and id(s.alternative_to) in gone:
                    s.alternative_to = None
        self.spots = kept

    def _assign_ids(self) -> None:
        """所の id（place_id）：作りの種類の頭の字と、作りの位置（mm）。同じ頭の字・同じ位置の所が 2 つ以上あるときだけ、
        右下の位置の順に、2 つ目から後ろに b・c… を足す。"""
        groups: dict[str, list[Spot]] = {}
        for sp in self.spots:
            groups.setdefault(place_id(self.number, sp, self.view.mm), []).append(sp)
        for base, sps in groups.items():
            sps.sort(key=lambda s: (round(s.rect[2], 1), round(s.rect[3], 1), s.kind))
            for k, sp in enumerate(sps):
                sp.id = base + ("" if k == 0 else chr(ord("a") + k))

    def _underline_spots(self) -> None:
        """下線の罫線：マスの縁でも表の中の線でもない横線で、線のすぐ上（1 行ぶん）に字のないもの。所は線の上の帯。
        線の左の、同じ高さの字（見出しの字のことが多い）と、その離れは facts に（離れていても示す）。"""
        lines = self.free_lines + [ln for c in self.cells for ln in c.lines]
        for h in self.hs:
            if self.index.at((h.a + h.b) / 2, h.pos - 3) is not None and self.index.at((h.a + h.b) / 2, h.pos + 3) is not None:
                continue  # 表の中の罫線
            near = None  # 線の左の、同じ高さ（字の下の端が線の近く）の、いちばん近い行
            for ln in lines:
                if ln.vertical:
                    continue
                if abs(ln.y1 - h.pos) <= 0.6 * ln.size and ln.x1 <= h.a + 0.5 * ln.size:
                    if near is None or ln.x1 > near.x1:
                        near = ln
            sz = near.size if near is not None else self.typical
            top = h.pos - 1.3 * sz  # 線の上の 1 行（字の大きさと、字の上下のあき）
            if any(h.a < c.cx < h.b and top < c.cy < h.pos and not c.text.isspace() for c in self.chars):
                continue  # 線の上に字がある（字に引いた下線）
            if self._cell_edge(h):
                continue  # マスの縁（表の外側の線）：下線ではない
            rect = (h.a, top, h.b, h.pos - 0.5)
            label = label_text(near.text.strip()) if near is not None else ""
            sp = Spot("text", rect, label, "下線の罫線", size=sz, host=near,
                      at=(len(near.text), len(near.text)) if near is not None else None, shape="underline",
                      anchor=(h.a, h.pos), write_size_from="線の左の字の大きさ" if near is not None else "ページのふつうの字の大きさ")
            sp.facts = {"rect_mm": self._rect_mm(rect), "length_mm": self._mm(h.b - h.a)}
            if near is not None:
                sp.facts["left"] = {"text": near.text.strip(), "mm": self._mm(max(0.0, h.a - near.x1))}
            self.spots.append(sp)

    def _cell_edge(self, h: Rule) -> bool:
        """横線 h が、どれかのマスの上か下の縁になっている。"""
        for c in self.index.near(h.a, h.pos - 2, h.b, h.pos + 2):
            if (abs(c.y0 - h.pos) <= 1.5 or abs(c.y1 - h.pos) <= 1.5) and c.x0 >= h.a - 1.5 and c.x1 <= h.b + 1.5:
                return True
        return False

    def _label_for(self, first: Cell, last: Cell) -> tuple[str, str, str]:
        """マス（または並び）の見出し：(見出し, 左の字, 上の字)。周りの字は _neighbours（facts と同じもの）。"""
        return self._label_of(self._neighbours((first.x0, first.y0, last.x1, last.y1)))


def line_band(ln: Line) -> tuple[float, float, float | None]:
    """行の所の、書く四角の上下と基線：(上, 下, 基線)。行のいちばん大きな字（ルビなどの小さな字を除く：基線と字の高さは本文の字）の
    高さに、字の下がり（ディセンダー）が切れないよう上下に字の大きさの 1/10 を足したもの。縦に読む行は、基線なし。"""
    main = [c for c in ln.chars if c.size >= ln.size - 0.01] or ln.chars
    pad = 0.1 * ln.size
    if ln.vertical:
        return ln.y0 - pad, ln.y1 + pad, None
    bases = sorted(c.base for c in main)
    return min(c.y0 for c in main) - pad, max(c.y1 for c in main) + pad, bases[len(bases) // 2]


def _rtl_cell(c: Cell) -> bool:
    lines = [ln for ln in c.lines if not ln.vertical and ln.text.strip()]
    return bool(lines) and all(getattr(ln, "rtl", False) for ln in lines)


def _option_text(t: str) -> str:
    """箱の字の後ろ（前の字は逆さにして渡す）の字：次の 1 字ぶん以上の空き（行の空白は測った空きから作るので、全角の空白か
    2 字以上の半角の空白）か、次の箱の字まで（元のまま）。1 字ずつを同じ空白でつないだ並び（「重　　度　　訪　　問」）は、並び全体。"""
    t = t.lstrip(" 　")
    m = re.match(r"\S(?:[ 　]+\S)+", t)
    if m and evenly_spaced(m[0]):
        chain = m[0]
        stop = next((i for i, ch in enumerate(chain) if ch in BOX_GLYPHS or ch in CHECKED_GLYPHS), len(chain))
        if stop == len(chain):
            return chain
    out = []
    i = 0
    while i < len(t):
        ch = t[i]
        if ch in BOX_GLYPHS or ch in CHECKED_GLYPHS:
            break
        if _chars.is_space(ch):
            j = i
            while j < len(t) and _chars.is_space(t[j]):
                j += 1
            if _em_width(t, i, j) >= EM_SPACE - EM_ROUND:
                break
            out.append(t[i:j])
            i = j
            continue
        out.append(ch)
        i += 1
    return "".join(out).strip(" 　")


def _inside(r, box) -> bool:
    cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
    return box[0] <= cx <= box[2] and box[1] <= cy <= box[3]


def _area(r) -> float:
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def _in_bands(bands: dict[int, list[int]], wide: list[int], a: int, b: int) -> set[int]:
    """帯 a〜b にある項目。幅の広い所（ページより何桁も大きな欄）は、番号を 1 つずつ探さず、ある帯を見る。"""
    if b - a >= len(bands):
        return {i for k, items in bands.items() if a <= k <= b for i in items} | set(wide)
    return {i for k in range(a, b + 1) for i in bands.get(k, ())} | set(wide)


class _RectIndex:
    """四角の索引：重なりうる所（入れた順）だけを比べる（線の交わりが何千とあるページでも、所どうしを総当たりしない）。
    四角は大きさの桁（2 の冪）ごとの升目に置くので、大小の四角が混ざっても、1 つの四角は 4 升までに収まる。"""

    def __init__(self):
        self.items: list = []
        self.grids: dict[int, dict[tuple[int, int], list[int]]] = {}
        self.odd: list[int] = []  # 無限大など、升目に置けない四角：いつも比べる

    def add(self, s) -> None:
        i = len(self.items)
        self.items.append(s)
        r = s.rect
        if not math.isfinite(r[0] + r[1] + r[2] + r[3]):
            self.odd.append(i)
            return
        level = math.frexp(max(r[2] - r[0], r[3] - r[1], 1.0))[1]
        cell = 2.0 ** level
        grid = self.grids.setdefault(level, {})
        for x in range(math.floor(r[0] / cell), math.floor(r[2] / cell) + 1):
            for y in range(math.floor(r[1] / cell), math.floor(r[3] / cell) + 1):
                grid.setdefault((x, y), []).append(i)

    def touching(self, r, pad: float = 0.0, own_size: bool = False) -> list:
        """四角 r に重なりうるもの（入れた順）。pad：r をこれだけ広げて探す。own_size：相手の大きさの分も広げて探す
        （「欄の高さの内に接する」のように、相手の大きさで届く所を探すとき）。"""
        if not math.isfinite(r[0] + r[1] + r[2] + r[3] + pad):
            return self.items
        found = set(self.odd)
        for level, grid in self.grids.items():
            cell = 2.0 ** level
            grow = pad + (cell if own_size else 0.0)
            x0, x1 = math.floor((r[0] - grow) / cell), math.floor((r[2] + grow) / cell)
            y0, y1 = math.floor((r[1] - grow) / cell), math.floor((r[3] + grow) / cell)
            if (x1 - x0 + 1) * (y1 - y0 + 1) > len(grid):  # 升目が多い：ある升目だけを見る
                for (x, y), ids in grid.items():
                    if x0 <= x <= x1 and y0 <= y <= y1:
                        found.update(ids)
            else:
                for x in range(x0, x1 + 1):
                    for y in range(y0, y1 + 1):
                        found.update(grid.get((x, y), ()))
        return [self.items[i] for i in sorted(found)]


def _overlap(r1, r2) -> float:
    x0, y0 = max(r1[0], r2[0]), max(r1[1], r2[1])
    x1, y1 = min(r1[2], r2[2]), min(r1[3], r2[3])
    if x1 <= x0 or y1 <= y0:
        return 0.0
    inter = (x1 - x0) * (y1 - y0)
    a = (r1[2] - r1[0]) * (r1[3] - r1[1])
    b = (r2[2] - r2[0]) * (r2[3] - r2[1])
    return inter / max(1e-6, min(a, b))


def _typical_size(chars: list[Ch]) -> float:
    """ページのふつうの字の大きさ：字の大きさの中央の値（字がなければ 10）。"""
    if not chars:
        return 10.0
    s = sorted(c.size for c in chars)
    return s[len(s) // 2]


def _holds_text(k: Spot) -> bool:
    """字を書く所に、書き込みと同じ決め方（いちばん小さな字の大きさで）で1字が入るか。
    入らない所（小さな字の図の中の空欄など）も欄としては示し、書けない理由を添える（呼ぶ側）。"""
    if k.kind not in ("text", "inline", "space"):
        return True
    from .fill import spot_fits

    return spot_fits("1", [round(v, 1) for v in k.rect], round(float(k.size or 10.5), 2), k.kind)
