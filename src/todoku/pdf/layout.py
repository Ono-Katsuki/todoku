"""PDFのページを、人が読む順に並べ、入力欄を見つける（決定的な処理だけで）。

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

import bisect
import re
import unicodedata
from dataclasses import dataclass, field

from .. import chars as _chars  # 字の種類（空白・下線の字・ダッシュ・点線・かっこの組・チェックの箱の字）：Unicode の資料の表
from ..xlsx.window import label_text
from .content import PageContent


def _sentence_end(ch: str) -> bool:
    """文の終わりの字（Unicode の Sentence_Terminal：。 . ! ? । ۔ ؟ ። ։ など）。"""
    return _chars.is_sentence_terminal(ch)


BOX_GLYPHS = set(_chars.EMPTY_BOXES)
# 箇条書きの印にもよく使う箱の字（影のついた四角 U+274F〜U+2752 ❏❐❑❒・U+25FB ◻・U+25FD ◽）：チェックの箱にも箇条の印にも
# 使う字。チェックの印の所にはするが、規則では決めない（？）
BULLET_BOXES = set(_chars.BULLET_LIKE_BOXES)  # 箇条書きの印にもよく使う箱（チェックの欄とは決まらない：「？」）
# 記入済みの印。U+25A3 ▣（WHITE SQUARE CONTAINING BLACK SMALL SQUARE）は、とどくが ActiveX のどちらでもない箱に書く字
CHECKED_GLYPHS = set(_chars.CHECKED_BOXES) | set(_chars.MAYBE_CHECKED) | {"▣"}


# ====================================================================== 座標
MM_PT = 72 / 25.4  # 1 mm のポイント（1 pt = 1/72 インチ）


def user_unit(doc, page: dict) -> float:
    """ページの UserUnit（ページの座標の 1 単位が何ポイントか。省いたときは 1。ISO 32000-1 14.11.1 表 30、PDF 1.6〜）。
    正の数でなければ 1 とみなす。"""
    v = doc.resolve(page.get("UserUnit")) if isinstance(page, dict) else None
    if isinstance(v, bool) or not isinstance(v, (int, float)) or not v > 0 or v != v or v == float("inf"):
        return 1.0
    return float(v)


class View:
    """ページの座標 → 見た目の座標（左上が原点、y は下向き、回転を反映）。"""

    def __init__(self, doc, page: dict):
        box = page.get("CropBox") or page.get("MediaBox") or [0, 0, 595, 842]
        box = [float(doc.resolve(v)) for v in doc.resolve(box)]
        if len(box) != 4:
            box = [0, 0, 595, 842]
        mb = page.get("MediaBox")
        if mb is not None:
            mb = [float(doc.resolve(v)) for v in doc.resolve(mb)]
            if len(mb) == 4:
                # CropBox は MediaBox の内側に切り詰める
                box = [max(box[0], min(mb[0], mb[2])), max(box[1], min(mb[1], mb[3])),
                       min(box[2], max(mb[0], mb[2])), min(box[3], max(mb[1], mb[3]))]
        self.x0, self.y0 = min(box[0], box[2]), min(box[1], box[3])
        self.x1, self.y1 = max(box[0], box[2]), max(box[1], box[3])
        rot = int(doc.resolve(page.get("Rotate")) or 0) % 360
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


@dataclass
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


def view_chars(pc: PageContent, view: View, dropped: list | None = None, merges: list | None = None) -> list[Ch]:
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
        if vertical and uy < 0:
            vertical = False  # 下から上へ進む字（左へ90度回した字）も、行としては横に扱わない
        base = (x0 + x1) / 2 if vertical else oy
        if g.size < 0.5 or x1 - x0 > view.width * 2 or y1 - y0 > view.height * 2:
            if dropped is not None:
                dropped.append(("size", g.text))
            continue
        c = Ch(g.text, x0, y0, x1, y1, max(g.size, 1.0), vertical, base, g.invisible, g.legacy, seq)
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
    cells = []
    seen = set()
    for (hi, k), (x, y) in pts.items():
        # 右の交点（同じ横線の上）
        rights = [k2 for k2 in by_h[hi] if vs_sorted[k2].pos > x + tol]
        downs = [h2 for h2 in by_v[k] if hs[h2].pos > y + tol]
        done = False
        for h2 in downs:
            y2 = hs[h2].pos
            # 縦線 k が y から y2 までつながっているか
            v = vs_sorted[k]
            if not (v.a - tol <= y and v.b + tol >= y2):
                break
            for k2 in rights:
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


class CellIndex:
    def __init__(self, cells: list[Cell], bucket: float = 24.0):
        self.cells = cells
        self.bucket = bucket
        self.grid: dict[tuple[int, int], list[int]] = {}
        for i, c in enumerate(cells):
            for gx in range(int(c.x0 // bucket), int(c.x1 // bucket) + 1):
                for gy in range(int(c.y0 // bucket), int(c.y1 // bucket) + 1):
                    self.grid.setdefault((gx, gy), []).append(i)

    def at(self, x: float, y: float) -> Cell | None:
        best = None
        for i in self.grid.get((int(x // self.bucket), int(y // self.bucket)), ()):
            c = self.cells[i]
            if c.x0 - 0.5 <= x <= c.x1 + 0.5 and c.y0 - 0.5 <= y <= c.y1 + 0.5:
                if best is None or c.w * c.h < best.w * best.h:
                    best = c
        return best

    def near(self, x0, y0, x1, y1) -> list[Cell]:
        out = {}
        b = self.bucket
        for gx in range(int(x0 // b), int(x1 // b) + 1):
            for gy in range(int(y0 // b), int(y1 // b) + 1):
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
    ver = [c for c in chars if c.vertical]
    for vertical, group in ((False, hor), (True, ver)):
        if not group:
            continue
        group.sort(key=lambda c: (c.base, c.x0 if not vertical else c.y0))
        rows: list[list[Ch]] = []
        for c in group:
            if rows:
                last = rows[-1]
                ref = sum(x.base for x in last) / len(last)
                sz = max(max(x.size for x in last), c.size)
                if abs(c.base - ref) <= 0.45 * sz:
                    last.append(c)
                    continue
            rows.append([c])
        for row in rows:
            row.sort(key=lambda c: c.y0 if vertical else c.x0)
            parts = _split_overprinted(row)
            if len(parts) > 1:
                made = [_line(sorted(p, key=lambda c: c.y0 if vertical else c.x0), vertical) for p in parts]
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


def _clash(a: Ch, b: Ch) -> bool:
    """違う字の箱が大きく重なる（結合する記号は、基の字に重ねて描くものなので数えない）。"""
    if a.text == b.text or a.text.isspace() or b.text.isspace():
        return False
    if all(_is_mark(ch) for ch in a.text) or all(_is_mark(ch) for ch in b.text):
        return False
    return _box_overlap(a, b) > 0.5


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
    segs: list[list[Ch]] = []
    for c in sorted(row, key=lambda c: c.seq):
        if segs:
            p = segs[-1][-1]
            back = c.x0 < p.x0 - 0.1 * max(c.size, p.size) if not c.vertical else c.y0 < p.y0 - 0.1 * max(c.size, p.size)
            if not (back and any(_clash(c, o) for o in segs[-1])):
                segs[-1].append(c)
                continue
        segs.append([c])
    if len(segs) < 2:
        return [row]
    groups: list[list[Ch]] = []
    for sg in segs:
        for g in groups:
            if not any(_clash(x, y) for x in sg for y in g):
                g += sg
                break
        else:
            groups.append(list(sg))
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
            gap = b0 - a1
            sz = max(prev.size, c.size)
            had_space = any((x.cy if vertical else x.cx) > a1 - 0.5 and (x.cy if vertical else x.cx) < b0 + 0.5 for x in spaces)
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
                and ch not in BULLET_BOXES and ch not in CHECKED_GLYPHS and text[i:j] != core:
            out.append(Blank(i, j, False, "placeholder"))
        i = j
    out.sort()
    edges = [x for x in out if x.why in ("lead", "trail")]
    if align and len(edges) == 2 and len(out) == 2 and abs(w(*edges[0]) - w(*edges[1])) < EM_SPACE:
        out = []  # 中央にそろえた行（左右の空きの差が 1 字に満たない）：両側の空きは、そろえ方の空き
    return out


MIN_WRITE = 6.0  # これより低い字の行の所は、小さな字の行として（書けるかは書き込みと同じ決め方で）示す
MIN_GLYPH = 4.0  # 書き込みのいちばん小さな字の大きさ（fill.layout_text）：幅か高さがこれに満たないマスには、どんな字も入らない


def _vov_rect(r: tuple[float, float, float, float], c) -> float:
    """矩形 r とマス c の、縦の重なりの割合（r の高さに対して）。"""
    h = r[3] - r[1]
    return max(0.0, min(r[3], c.y1) - max(r[1], c.y0)) / h if h > 0 else 0.0


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
    group: list = field(default_factory=list)  # （使わない：これまでの呼び出しのため）
    readonly: str = ""  # 書けない理由（小さすぎて1字も入らない・印刷された記入済みの印など）。欄としては示す
    weak: str = ""  # （使わない：これまでの呼び出しのため）
    gate: str = ""  # （使わない：これまでの呼び出しのため）
    shape: str = ""  # 見つけた作り：cell・cells・gap・lead・trail・connector・leader・dash・bracket・box・checked・underline
    anchor: tuple | None = None  # id を決める位置（作りの位置：マスの左上・空きの始まりと行の上。pt）
    facts: dict = field(default_factory=dict)  # 窓に示す事実（mm・字・離れ）
    write_size_from: str = ""  # 既定の字の大きさの出どころ（default_size）


def _agent(sp: "Spot", why: str = "", gate: str = "") -> "Spot":
    """字の間の空きなど、描かれた空欄でない所にする（窓に「？」）。why：見つけた形の分類。"""
    sp.confidence = "agent"
    if why:
        sp.reason = why
    if gate and not sp.gate:
        sp.gate = gate
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
        kept: list[Spot] = []
        for s in self.spots:
            chk = s.kind == "checkbox"
            lim = 0.7 if chk else 0.6
            k = next((k for k in kept if (k.kind == "checkbox") == chk and _overlap(s.rect, k.rect) > lim), None)
            if k is not None:
                s.alternative_to = k
                continue
            kept.append(s)
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
        for i, it in enumerate(items):
            for k in range(int(it[1] // B), int(it[3] // B) + 1):
                self._yb.setdefault(k, []).append(i)
            for k in range(int(it[0] // B), int(it[2] // B) + 1):
                self._xb.setdefault(k, []).append(i)
        self._B = B
        self._line_of = {id(ch): ln for ln in self.all_lines() for ch in ln.chars}

    def _neighbours(self, r) -> dict:
        """所の四角 r の、左・上・右・下でいちばん近い字（字のあるマスの字・マスの外の行の字。元のまま）と、その離れ（mm）。
        左右は上下の範囲が重なるもの、上下は左右の範囲が重なるものだけ。どこまで離れていても示す（近いかはエージェントが見る）。"""
        x0, y0, x1, y1 = r
        B = self._B
        best: dict[str, tuple] = {}
        rows = {i for k in range(int(y0 // B), int(y1 // B) + 1) for i in self._yb.get(k, ())}
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
        cols = {i for k in range(int(x0 // B), int(x1 // B) + 1) for i in self._xb.get(k, ())}
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
        for c in self.chars:
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
            boxed = ch in BOX_GLYPHS or ch in BULLET_BOXES
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
                y0, y1 = self._band(ln, cell, x0, x1)
                sp = _agent(Spot("text" if cell is not None else "inline", (x0, y0, x1, y1), "［　］", GAP_WHY["blank_line"], line_text=text, blank=(0, n),
                                 cell=cell, size=sz, host=ln, at=(0, n), shape="blank_line", anchor=(lb, ln.y0),
                                 write_size_from="行の字の大きさ"))
                sp.facts = {"rect_mm": self._rect_mm(sp.rect), "width_mm": self._mm(rb - lb), "width_em": round((rb - lb) / sz, 1),
                            "bound": {"kind": f"{lk}・{rk}", "text": " ／ ".join(t for t in (lt, rt) if t), "mm": 0.0}}
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
                y0, y1 = self._band(ln, cell, x0, x1)
                rect = (x0, y0, x1, y1)
                if any(k.cell is cell and k.shape in ("lead", "trail") and all(abs(u - v) <= 0.5 for u, v in zip(k.rect, rect))
                       for k in self.spots):
                    continue  # 同じマスのほかの行の前後の空きと同じ四角（縦に並ぶ行の左の、同じ空き）：同じ所は 1 つ
            else:
                at = (max(a0, 0), min(b0, n))
                if at[1] <= at[0]:
                    continue
                x0, _, x1, _ = ln.span_rect(*at)
                rect, anchor, bound = (x0, ln.y0, x1, ln.y1), (x0, ln.y0), None
                if why == "placeholder":
                    # PDF の字は消せない：伏せ字の上を白い下地の欄で覆って書く（元の字は残る）。覆う四角は字の箱より少し大きく
                    rect = (x0 - 0.3, ln.y0 - 0.1 * sz, x1 + 0.3, ln.y1 + 0.1 * sz)
            if rect[2] <= rect[0]:
                continue  # 空きがまったくない（縁や次の字に接している）
            orig = text[at[0]:at[1]]
            shown = text[:at[0]] + (orig if orig.strip(" 　") else "") + "［　］" + text[at[1]:]
            if why == "placeholder":
                shown = text[:at[0]] + "［　］" + text[at[1]:]
            # マスの中の行の前後の空きは、行と行の間の帯（行より高いこともある四角）：字の欄（text）として書く。行の中の空きは inline
            kind = "text" if why in ("lead", "trail") and cell is not None else "inline"
            sp = Spot(kind, rect, label_text(shown), GAP_WHY.get(why, "空き"), line_text=text, blank=at, cell=cell,
                      size=sz, host=ln, at=at, shape=why, anchor=anchor, write_size_from="行の字の大きさ",
                      cover=orig if why == "placeholder" else "")
            if not blank.structural:
                _agent(sp)
            # 幅は測った空き（行の前後の空きは、端から字までの空き。書く所の四角は、端と字から少し離したもの）
            w = (left_room if on_left else right_room) if why in ("lead", "trail") else rect[2] - rect[0]
            facts = {"rect_mm": self._rect_mm(rect), "width_mm": self._mm(w), "width_em": round(w / sz, 1),
                     "before": text[:at[0]].rstrip(" 　")[-12:], "after": text[at[1]:].lstrip(" 　")[:12],
                     "line": text.strip(" 　")}
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

    def _band(self, ln: Line, cell: Cell | None, x0: float, x1: float) -> tuple[float, float]:
        """行の前後の空き（x0〜x1）の上下：マスの中では、その横の範囲にかかる上下のほかの行の字（なければマスの縁）まで。
        マスの外の行は、行の高さだけ（上下に描かれた端がない）。"""
        if cell is None:
            return ln.y0, ln.y1
        pad = TEXT_PAD * ln.size
        top, bottom = cell.y0, cell.y1
        for o in cell.lines:
            if o is ln or not o.text.strip(" \u3000") or min(o.x1, x1) - max(o.x0, x0) <= 0:
                continue
            if o.y1 <= ln.y0 + pad:
                top = max(top, o.y1 + pad)
            elif o.y0 >= ln.y1 - pad:
                bottom = min(bottom, o.y0 - pad)
        return min(top, ln.y0), max(bottom, ln.y1)

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
        stop = next((i for i, ch in enumerate(chain) if ch in BOX_GLYPHS or ch in BULLET_BOXES or ch in CHECKED_GLYPHS), len(chain))
        if stop == len(chain):
            return chain
    out = []
    i = 0
    while i < len(t):
        ch = t[i]
        if ch in BOX_GLYPHS or ch in BULLET_BOXES or ch in CHECKED_GLYPHS:
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


def _vov(a: Cell, b: Cell) -> float:
    inter = min(a.y1, b.y1) - max(a.y0, b.y0)
    return inter / max(1e-6, min(a.h, b.h))


def _inside(r, box) -> bool:
    cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
    return box[0] <= cx <= box[2] and box[1] <= cy <= box[3]


def _area(r) -> float:
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


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
