"""PDFのページを、人が読む順に並べ、入力欄を見つける（決定的な処理だけで）。

1. 文字・線をページの見た目の座標（左上が原点、下向きが y）に直す（ページの回転も反映）
2. 罫線から表のマス（セル）を組み立てる（線の交わりから最小の四角を作る）
3. 文字をマスに振り分け、マスの中・マスの外で行を作る（字間の空きは空白として残す）
4. 入力欄を見つける：
   - 空のマス（左か上のマスの文字を見出しにする）。同じ大きさの空のマスが横に並ぶものは、1字ずつ書く欄（桁の枠）
   - 文字の中の空欄（「令和　　年　　月　　日」「（　　　）か月」）
   - 下線（見出しの右に引かれた線で、上に文字のないもの）
   - チェックの印（□ の字、または小さな正方形のマスで、右に選択肢の文字があるもの）
   - 文字の右に大きな空きがある1行のマス（「電話番号」と同じマスの右側）
5. 行を上から、行の中は左から並べた「読み順」を作る
"""

from __future__ import annotations

import bisect
import re
import unicodedata
from dataclasses import dataclass, field

from ..xlsx.window import inline_blanks, norm_label
from .content import Glyph, PageContent

BOX_GLYPHS = set("□☐◻▢❏❐❑❒⬜")
CHECKED_GLYPHS = set("☑☒■▣✓✔")


# ====================================================================== 座標
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
        dup = False
        for ix in (kx - 1, kx, kx + 1):
            for iy in (ky - 1, ky, ky + 1):
                for ofr, osize in seen.get((g.text, ix, iy), ()):
                    if is_overprint(ofr, fr, osize, g.size):
                        dup = True
                        break
                if dup:
                    break
            if dup:
                break
        if dup:
            if dropped is not None:
                dropped.append(("dup", g.text))
            if merges is not None:
                merges.append((g.text, 1.0, (c.x0, c.y0, c.x1, c.y1)))  # is_overprint：送りの幅が半分より多く重なる
            continue
        seen.setdefault((g.text, kx, ky), []).append((fr, g.size))
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


def join_lines(lines) -> str:
    """マスの中の行をつなぐ。和文の折り返し（行の境の両側が全角）と、縦に1字ずつ並べた字は、空白なしでつなぐ。"""
    out = ""
    for ln in lines:
        t = ln.text.strip()
        if not t:
            continue
        if out and not (_is_cjk(out[-1]) and _is_cjk(t[0])):
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


def _spaces(parts: list, pos: list, a1: float, gap: float, sz: float, cjk: bool, had_space: bool) -> None:
    """字と字の間の空き（ポイント）を、空白の字にする。全角の文字のまわりは、全角の空白1つを1字ぶんとする。"""
    if cjk:
        n = round(gap / sz)
        if n == 0 and (gap > 0.3 * sz or had_space):
            n = 1
        ch = "\u3000"
    else:
        if gap <= 0.15 * sz and not had_space:
            return
        n = 1 if gap < 1.5 * sz else max(2, round(gap / (0.5 * sz)))
        ch = " "
    if n <= 0:
        return
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
            if stacked:
                gap = gap - 0.8 * sz if gap > 1.5 * sz else 0.0
            if gap > 0 or had_space:
                sp_parts: list[str] = []
                sp_pos: list[tuple[float, float]] = []
                _spaces(sp_parts, sp_pos, a1, max(gap, 0.0), sz, cjk, had_space)
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


# ====================================================================== 入力欄
ADDRESS_UNITS = set("市区町村郡都道府県")
# 数の後ろに付く単位（その前の空きに数を書く）。「日／月」「時間／週」のような割合も
UNIT_ONLY = re.compile(r"(?:時間|日間|週間|か月|ヶ月|カ月|千円|万円|年度|[日時分秒年月週回人名件号円台本枚個歳才種級％%])"
                       r"(?:[／/](?:時間|[日時週月年回]))?")


def _unit_at(text: str, i: int) -> bool:
    """text の i 文字目から単位（「分」「日／月」「種」）があり、その後ろが字でない（「分類」「種類」は単位でない）。"""
    m = UNIT_ONLY.match(text, i)
    if not m:
        return False
    nxt = text[m.end():m.end() + 1]
    return not nxt or not ("\u3040" <= nxt <= "\u30ff" or "\u3400" <= nxt <= "\u9fff" or nxt.isalpha())


# 時刻の欄だけの行（「　　：　　～　　：　　」）、または時刻の範囲を含む行
_TIME_SKELETON = re.compile(r"[：:～~（）() \u3000]*[：:][：:～~（）() \u3000]*")
_TIME_RANGE = re.compile(r"[：:][ \u3000]{2,}[～~][ \u3000]{2,}[：:]")
# 字間を空けた見出し（「事　業　所」「氏　　名」）
_SPACED = re.compile(r"(?:\S[ \u3000]+)+\S")
DASHES = "－-‐−―–—"
LEAD_UNITS = ("年", "知事", "市長", "町長", "村長", "区長", "長　殿", "長殿")


class Blank(tuple):
    """文字の中の空欄 (始め, 終わり)。structural：字の形だけで決まる空欄（「＿＿＿」の下線の字、かっこの組「（　　）」の中が
    空白だけのもの）か。偽のものは、語（「市」「区長」「年」「：」など）を手がかりに見つけた空欄で、規則では欄と決められない
    （エージェントが判断する所として示す）。2つの数の組としてそのまま使える（a, b = 空欄・set・sorted）ので、これまでの呼び出し
    （Word など）はそのまま動く。"""

    structural: bool

    def __new__(cls, a: int, b: int, structural: bool = False):
        t = super().__new__(cls, (a, b))
        t.structural = structural
        return t


_BRACKETS = {"（": "）", "(": ")", "［": "］", "[": "]", "【": "】", "「": "」", "〔": "〕"}


def _structural(text: str, a: int, b: int) -> bool:
    """空欄 text[a:b] が、字の形だけで決まるもの：下線の字だけ、または かっこの組の中が空白だけ（「（　　）」）。
    英数字・「/」「.」「-」にはさまれた下線の字の並び（URL・ファイルの名前の「__icsFiles」）は、名前の一部のこともあるので除く。"""
    seg = text[a:b]
    if seg and all(ch in "_＿" for ch in seg):
        # 下線の字の並びが、英数字（「/」「.」「-」を含む）に接していない：書くための線（「氏名＿＿＿」「: ____」）。
        # 片側でも接していれば（「Name____」「_HVeg」）、語の一部（名前の区切り）のこともあるので、規則では決めない
        return not underscore_touches_word(text, a, b)
    if not seg or seg.strip(" 　"):
        return False
    i, j = a, b
    while i > 0 and text[i - 1] in " 　":
        i -= 1
    while j < len(text) and text[j] in " 　":
        j += 1
    return i > 0 and j < len(text) and _BRACKETS.get(text[i - 1]) == text[j]


def underscore_touches_word(text: str, a: int, b: int) -> bool:
    """下線の字の並び（text[a:b] を含む、続いた並び全体）の前か後ろが、英数字・「/」「.」「-」か。"""
    def ident(ch: str) -> bool:
        return ch.isascii() and (ch.isalnum() or ch in ".-/")

    x, y = a, b
    while x > 0 and text[x - 1] in "_＿":
        x -= 1
    while y < len(text) and text[y] in "_＿":
        y += 1
    return (0 < x and ident(text[x - 1])) or (y < len(text) and ident(text[y]))


def underscore_in_word(text: str, a: int, b: int) -> bool:
    """下線の字の並び（text[a:b] を含む、続いた並び全体）が、英数字・「/」「.」「-」にはさまれているか。"""
    def ident(ch: str) -> bool:
        return ch.isascii() and (ch.isalnum() or ch in ".-/")

    x, y = a, b
    while x > 0 and text[x - 1] in "_＿":
        x -= 1
    while y < len(text) and text[y] in "_＿":
        y += 1
    return 0 < x and y < len(text) and ident(text[x - 1]) and ident(text[y])


def pdf_blanks(text: str, lead: int = 0, trail: int = 0) -> list[Blank]:
    """文字の中の空欄（Excel と同じ規則）に、PDF でよく見る形を足す。各要素は Blank（(始め, 終わり) と structural）。
    lead・trail は、マスの縁までの空きを空白にして前後に足した数（作った空白なので、強い手がかりがあるときだけ空欄にする）。
    - 住所の欄：「　　　市　区」「　　　町　村」の前の空き、「都 道」「府 県」の前の空き、「市 区」の後の空き
    - 行の頭の空き（「　　　知事　殿」「　　　年　月　日」）、「（内線）」の前後の空き
    - 「　　　－　　　」のようにハイフンでつながる空き
    語（単位・年月日・「区長」・「：」・「－」など）を手がかりにしたものは structural が偽。"""
    return [Blank(a, b, _structural(text, a, b)) for a, b in _blank_spans(text, lead, trail)]


def _blank_spans(text: str, lead: int = 0, trail: int = 0) -> list[tuple[int, int]]:
    """pdf_blanks の中身：空欄の (始め, 終わり)。"""
    core = text[lead:len(text) - trail] if trail else text[lead:]
    out = set(inline_blanks(text))
    time_like = bool(_TIME_SKELETON.fullmatch(core) or _TIME_RANGE.search(core))
    spaced = bool(_SPACED.fullmatch(core.strip(" \u3000")))
    for m in re.finditer(r"[ \u3000]{2,}", text):
        a, b = m.start(), m.end()
        after = text[b:b + 2]
        before = text[a - 1:a] if a else ""
        a1, b1 = after[:1], before  # どちらも空の文字列のことがある（"" はどの文字列にも含まれるので、確かめてから比べる）
        if a1 and a1 in ADDRESS_UNITS and (len(after) < 2 or after[1] in ADDRESS_UNITS or after[1] in " \u3000・"):
            out.add((a, b))
        elif any(text[b:].startswith(u) for u in ("区長", "市長", "町長", "村長", "知事")):
            out.add((a, b))  # 「名古屋市　　　区長」の区の名前、「　　県知事」の県の名前
        elif a == 0 and any(text[b:].startswith(u) for u in LEAD_UNITS):
            out.add((a, b))
        elif a == 0 and a1 in ("（", "("):
            out.add((a, b))
        elif b == len(text) and b1 and (b1 in ADDRESS_UNITS or b1 == "〒"):
            out.add((a, b))
        elif a1 and a1 in DASHES and text[b + 1:b + 2] in (" ", "\u3000", ""):
            out.add((a, b))
        elif b1 and b1 in DASHES and (b < len(text) or a > 0):
            out.add((a, b))  # 「〒　　　−　　　　」の「−」の後ろ
        elif a1 and a1 in "：:" and time_like:
            out.add((a, b))  # 時刻「　　：　　～　　：　　」の時（「：」の前）
        elif b1 and b1 in "：:" and time_like:
            out.add((a, b))  # 時刻の分（「：」の後ろ）
        elif a1 and a1 in "～~" and a > 0:
            out.add((a, b))  # 「　　～　　」の「～」の前
        elif b1 and b1 in "～~" and b < len(text):
            out.add((a, b))  # 「～」の後ろ
        elif a > 0 and b1 and not b1.isspace() and _unit_at(text, b) and not spaced:
            out.add((a, b))  # 「片道　　分」「手帳　　種　　級」「　　日／月」の単位の前（「氏　　名」のような見出しは除く）
    kept: list[tuple[int, int]] = []
    for a, b in sorted(out):
        if trail and a >= len(text) - trail:
            # 縁までの空き（後ろ）：住所の単位の後、または「（内線）」のように かっこだけの文字の後
            before = text[a - 1:a]
            short_paren = core.strip(" \u3000").startswith(("（", "(")) and len(core.strip(" \u3000")) <= 10
            if not before or not (before in ADDRESS_UNITS or before == "〒" or before in DASHES or (before in "）)" and short_paren)
                                  or (before in "：:" and time_like)):
                continue
        if lead and b <= lead:
            after = text[b:b + 2]
            date_like = bool(re.search(r"[年月][ \u3000]+[月日]", core)) or bool(re.search(r"時[ \u3000]+分", core))
            unit_only = bool(UNIT_ONLY.fullmatch(core.strip(" \u3000")))
            if not after[:1] or not (after[:1] in ADDRESS_UNITS or after[:1] in "（(" or
                    any(text[b:].startswith(u) for u in LEAD_UNITS if u != "年") or
                    (after[:1] in "年月日時" and date_like) or unit_only or (after[:1] in "：:" and time_like)):
                continue
        if kept and a < kept[-1][1]:
            if b > kept[-1][1]:
                kept[-1] = (kept[-1][0], b)
            continue
        kept.append((a, b))
    return kept


_OPTION_GROUP = re.compile(r"[^・／/（）()\s\u3000]{1,10}(?:[\s\u3000]*[・／/][\s\u3000]*[^・／/（）()\s\u3000]{1,10}){1,5}")


MIN_WRITE = 6.0  # これより低い字の行の所は、小さな字の行として（書けるかは書き込みと同じ決め方で）示す
MIN_GLYPH = 4.0  # 書き込みのいちばん小さな字の大きさ（fill.layout_text）：幅か高さがこれに満たないマスには、どんな字も入らない


def _vov_rect(r: tuple[float, float, float, float], c) -> float:
    """矩形 r とマス c の、縦の重なりの割合（r の高さに対して）。"""
    h = r[3] - r[1]
    return max(0.0, min(r[3], c.y1) - max(r[1], c.y0)) / h if h > 0 else 0.0


def _cut_option(t: str) -> str:
    """チェックの印の右の文字から、選択肢の名前だけを取る（大きな空き・※・● で切る）。"""
    t = t.strip(" \u3000")
    # 字間を均等に空けた選択肢（「重　　度　　訪　　問　　介　　護」）：1字ずつの並びで、空きの長さがそろうものは1つの語
    m = re.match(r"\S(?:[ \u3000]+\S)+", t)
    if m:
        gaps = re.findall(r"[ \u3000]+", m[0])
        toks = re.split(r"[ \u3000]+", m[0])
        if len(toks) >= 3 and all(len(x) == 1 for x in toks) and max(map(len, gaps)) - min(map(len, gaps)) <= 1:
            return "".join(toks)
    m = re.search(r"[\u3000 ]{2,}|※|●|・・", t)
    if m and m.start() > 0:
        t = t[:m.start()]
    return t.strip(" \u3000")


@dataclass
class Spot:
    """入力欄（またはチェックの印）の候補。座標は見た目の座標。"""

    kind: str  # text / comb / inline / checkbox
    rect: tuple[float, float, float, float]
    label: str = ""
    reason: str = ""
    boxes: list = field(default_factory=list)  # 桁の枠：1字ずつの四角
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
    cover: str = ""  # 覆って書く元の字（伏せ字の「○○」）：書くときは白い下地の欄で覆う（元の字は消さない）
    # 種類が形だけでは決まらない所（小さな正方形のマス：チェックの印か1字の欄か、など）：候補の種類（初めが既定の書き方）
    kind_candidates: list = field(default_factory=list)
    alternative_to: object = None  # 同じ所の別の読み方（重なる先の候補）：どちらに書くかはエージェントが決める（両方には書けない）
    group: list = field(default_factory=list)  # 縦に並ぶ同じ形の空欄（「都 道／府 県」の前の空き）：1つの欄かもしれない組
    readonly: str = ""  # 書けない理由（小さすぎて1字も入らない・印刷された記入済みの印など）。欄としては示す
    gate: str = ""  # これまでの決め方では示していなかった所（どの決め方で外していたか）：数え上げのため


def _agent(sp: "Spot", why: str = "", gate: str = "") -> "Spot":
    """規則では決められない所にする（窓に「？」）。why：見つけた理由（測った大きさなど）。"""
    sp.confidence = "agent"
    if why:
        sp.reason = why
    if gate and not sp.gate:
        sp.gate = gate
    return sp


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
        self.room_marks: list[tuple] = []  # 欄にしない空き（行かマス, 向き, 大きさ pt）：窓に「［→余白 n pt］」と出す
        self._big_empty = [c for c in self._big_empty
                           if not any(c.x0 <= ch.cx <= c.x1 and c.y0 <= ch.cy <= c.y1 for ch in self.chars)
                           and not any(o.x0 >= c.x0 - 0.5 and o.x1 <= c.x1 + 0.5 and o.y0 >= c.y0 - 0.5 and o.y1 <= c.y1 + 0.5
                                       for o in self.cells)]
        self.drawn_text = self._drawn_text_regions()
        self._find_spots()
        if self.drawn_text:
            # 図形で描いた字の所：字の形の線や塗りからできたマス・空きは、書く所ではない
            m = 0.5 * max(self.typical, 10)  # 字の形の線（下線に見える横の線など）は、字の形の外へ少し出る
            self.spots = [k for k in self.spots if not any(
                _overlap(k.rect, (r[0] - m, r[1] - m, r[2] + m, r[3] + m)) > 0 for r, _n in self.drawn_text)]
        self._circled()
        for k in self.spots:
            # 書き込みと同じ決め方で1字も入らない所も、消さずに示す（規則では決められない所・書けない所として）
            if not _holds_text(k):
                x0, y0, x1, y1 = k.rect
                _agent(k, f"{k.reason}（小さい：幅 {x1 - x0:.1f}pt・高さ {y1 - y0:.1f}pt。1字も入らない）", "size")
                k.readonly = f"小さい（幅 {x1 - x0:.1f}pt・高さ {y1 - y0:.1f}pt）ので、1字も入りません"
        self.flags: list[dict] = []  # 読む段階に関わる注意：{grade, title, detail, where}
        self._read_flags()

    def _drawn_text_regions(self) -> list[tuple]:
        """字を図形（曲線を含む塗りの形）にして描いた所：字の大きさほどの小さな塗りの形が 10 個以上、間を空けずに
        並んでいて、その所に字（字の情報）がないもの。窓では字として読めないので、位置を示して知らせる。"""
        shapes = []
        for x0, y0, x1, y1, stroke, fill in self.pc.curves:
            if not fill:
                continue
            r = self.view.rect(x0, y0, x1, y1)
            w, h = r[2] - r[0], r[3] - r[1]
            if 0.3 <= max(w, h) <= 3 * max(self.typical, 10):
                shapes.append(r)
        if len(shapes) < 10:
            return []
        gap = 1.2 * max(self.typical, 10)  # 行の間（字1字ぶんほど）も、同じ所のつながりとする
        groups: list[list] = []
        for r in sorted(shapes):
            hit = [g for g in groups if g[0][0] - gap <= r[2] and r[0] <= g[0][2] + gap and g[0][1] - gap <= r[3] and r[1] <= g[0][3] + gap]
            if hit:
                g = hit[0]
                for o in hit[1:]:
                    g[0] = (min(g[0][0], o[0][0]), min(g[0][1], o[0][1]), max(g[0][2], o[0][2]), max(g[0][3], o[0][3]))
                    g[1] += o[1]
                    groups.remove(o)
                g[0] = (min(g[0][0], r[0]), min(g[0][1], r[1]), max(g[0][2], r[2]), max(g[0][3], r[3]))
                g[1] += 1
            else:
                groups.append([r, 1])
        out = []
        for box, n in groups:
            if n < 10:
                continue
            if any(box[0] <= c.cx <= box[2] and box[1] <= c.cy <= box[3] for c in self.chars if not c.text.isspace()):
                continue  # 字のある所（図形は飾り・表の模様など）
            out.append((box, n))
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
                      "その位置の字は view_page で画像を見て読みます", legacy)
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
            self.flag("B" if has_text else "E", "図形で描いた字（字の情報がない）",
                      "字の形を図形（線と塗り）にして描いた所があり、字の情報がないので、窓に字として出せません。注意書きなど大事な文の"
                      "ことがあるので、その位置は view_page でページの画像を見て読み、中身を推測で埋めない",
                      [f"{self.number}ページ（左 {r[0]:.0f}・上 {r[1]:.0f}〜右 {r[2]:.0f}・下 {r[3]:.0f} pt、字の形 {n} 個）"
                       for r, n in self.drawn_text])
        bad = [m for m in self.merges if m[1] <= 0.5]
        if bad:
            self.flag("C", "重ね書きとして1字にした字", "箱が重ならない字を、重ね書きとして1字にしました（読み落としのおそれ）",
                      [self.where(r[0], r[1], t) for t, _ov, r in bad])

    def _circled(self) -> None:
        """○で囲む選択肢のうち、ページに描かれた楕円（線の曲線の外枠に内接する楕円）が、その字だけを囲んでいるもの：選んだ字。"""
        from .circle import inside

        ells = []
        for x0, y0, x1, y1, stroke, _fill in self.pc.curves:
            if not stroke:
                continue
            r = self.view.rect(x0, y0, x1, y1)
            if r[2] - r[0] > 0.5 and r[3] - r[1] > 0.5:
                ells.append(((r[0] + r[2]) / 2, (r[1] + r[3]) / 2, (r[2] - r[0]) / 2, (r[3] - r[1]) / 2))
        if not ells:
            return

        def glyphs(r):
            # 選択肢の字の箱：選択肢の位置（行の高さの帯）の中に中心がある字の箱を合わせたもの。書くとき（circle.resolve）は
            # 字の箱を囲むので、行の高さ（行のいちばん大きい字の高さ）で比べると、囲んだ○でも角がはみ出す
            cs = [c for c in self.chars if not c.text.isspace() and r[0] <= c.cx <= r[2] and r[1] <= c.cy <= r[3]]
            if not cs:
                return r
            return (min(c.x0 for c in cs), min(c.y0 for c in cs), max(c.x1 for c in cs), max(c.y1 for c in cs))

        for sp in self.spots:
            if sp.kind != "circle":
                continue
            boxes = [(o, glyphs(r)) for o, r in sp.options]
            for cx, cy, a, b in ells:
                e = (cx, cy, a + 0.01, b + 0.01)  # 描いた座標の丸め（小数 3 桁）の分
                inner = [o for o, r in boxes if inside(e, r)]
                if len(inner) == 1:
                    sp.chosen = inner[0]
                    break

    def cell_fill(self, c: Cell):
        """マスを塗っている色（いちばん上に描かれたもの）。なければ None。"""
        col = None
        for x0, y0, x1, y1, fill in self.fills:
            if x0 <= c.x0 + 1 and y0 <= c.y0 + 1 and x1 >= c.x1 - 1 and y1 >= c.y1 - 1:
                if (x1 - x0) * (y1 - y0) < 0.5 * self.view.width * self.view.height:
                    col = fill
        return col

    # ------------------------------------------------------------------ 入力欄を見つける
    def _find_spots(self) -> None:  # noqa: C901
        cells = self.cells
        typ = self.typical
        used_cells = set()
        # 2) 文字の中の空欄・チェックの印（マスの中の行と、マスの外の行）
        for cell, lines in [(c, c.lines) for c in cells] + [(None, self.free_lines)]:
            before = len(self.spots)
            for ln in lines:
                self._line_spots(ln, cell)
            self._group_stacked(before)
        # 2') マスの外の行：見出しの右の余白、行の頭の空き（「　　知事　殿」）
        for ln in self.free_lines:
            self._free_line_spots(ln)
        # 1) 空のマス：同じ大きさの空のマスが横に並ぶものは「桁の枠」の候補
        empties = [c for c in cells if not c.visible()]
        rows: dict[int, list[Cell]] = {}
        for c in empties:
            if c.w >= 5 and c.h >= 5:
                rows.setdefault(round(c.y0), []).append(c)
        comb_cells: set[int] = set()
        undecided: list[tuple[Cell, str, str]] = []  # 規則では欄と決められない空のマス（マス・理由・外していた決め方）
        for y, row in rows.items():
            row.sort(key=lambda c: c.x0)
            run: list[Cell] = []
            skip: set[int] = set()
            seq = row + [None]
            for i, c in enumerate(seq):
                if c is not None and id(c) in skip:
                    continue
                if c is not None and run and abs(c.x0 - run[-1].x1) <= 1.5 and abs(c.w - run[0].w) <= 1.5 \
                        and abs(c.h - run[0].h) <= 1.5 and abs(c.y1 - run[0].y1) <= 1.5 and c.w <= 2.2 * c.h:
                    run.append(c)
                    continue
                wide_last = False
                if len(run) >= 3 and c is not None and abs(c.x0 - run[-1].x1) <= 1.5 and abs(c.h - run[0].h) <= 1.5 \
                        and abs(c.y1 - run[0].y1) <= 1.5 and run[0].w - 1.5 <= c.w <= 2 * run[0].w:
                    # 並びの最後のマスだけ幅が違う（表の右の縁まで広げたマス）：その右に続く空のマスがなく、
                    # 罫線の上でも同じ並びの続き（隣り合い、高さが同じ）なら、並びの最後の桁
                    nxt = seq[i + 1] if i + 1 < len(seq) else None
                    if nxt is None or abs(nxt.x0 - c.x1) > 1.5:
                        run.append(c)
                        skip.add(id(c))
                        wide_last = True
                        c = nxt
                        if c is not None:
                            skip.add(id(c))
                if len(run) >= 3 and run[0].w <= 2.2 * run[0].h:
                    for x in run:
                        comb_cells.add(id(x))
                    label, rl, cl = self._label_for(run[0], run[-1])
                    # 同じ大きさの空のマスの並びは、1字ずつ書く欄のことも、別々の欄（表の列）のこともある：形の大きさでは決めない
                    sp = Spot("comb", (run[0].x0, run[0].y0, run[-1].x1, run[-1].y1), label,
                              f"同じ大きさの空のマス {len(run)} 個が横に並ぶ（1字ずつ書く欄か、1つの欄か："
                              f"マスの幅 {run[0].w:.1f}pt・高さ {run[0].h:.1f}pt）",
                              [(x.x0, x.y0, x.x1, x.y1) for x in run], row_label=rl, column_label=cl,
                              size=min(run[0].h * 0.7, 14), host=run[0], cell=run[0], kind_candidates=["comb", "text"])
                    self.spots.append(_agent(sp, gate="kind"))
                    if wide_last:
                        sp.reason += f"（最後のマスは幅 {run[-1].w:.1f}pt）"
                    for x in run[1:]:
                        x.comb_of = run[0]
                run = [c] if c is not None else []
        for c in empties:
            if id(c) in comb_cells:
                continue
            if c.w < MIN_GLYPH or c.h < MIN_GLYPH:
                # 書き込みのいちばん小さな字（4pt）も入らない細いマス（枠と枠のすき間）：欄にはせず、ページごとの数を窓に出す
                self.slivers.append(c)
                continue
            if c.w < 5 or c.h < 5:
                undecided.append((c, f"小さな空のマス（幅 {c.w:.1f}pt・高さ {c.h:.1f}pt）", "small_cell"))
                continue
            # 小さな正方形のマス：チェックの印のことも、1字（数・印）を書く欄のこともある：大きさでは決めず、候補の種類を示す
            # （向かい合う2辺が、そのマスだけの短い線のもの。大きなマスの中に □ を描いたときにできる、□ の左右の余りのマスは除く）
            if 5 <= c.w <= 18 and 5 <= c.h <= 18 and abs(c.w - c.h) <= 0.35 * max(c.w, c.h) and _own_edges(c):
                right = self._text_right_of(c.x1, c.y0, c.y1, limit=max(2.5 * c.w, 12))
                size_note = f"（幅 {c.w:.1f}pt・高さ {c.h:.1f}pt：チェックの印か、1字を書く欄か）"
                if not right:
                    label, rl, cl = self._label_for(c, c)
                    sp = Spot("checkbox", (c.x0, c.y0, c.x1, c.y1), label or "（選択肢の文字なし）",
                              "小さな正方形のマス" + size_note, cell=c, host=c, row_label=rl, column_label=cl,
                              kind_candidates=["checkbox", "text"], size=_fit_size(c.h, typ))
                else:
                    sp = Spot("checkbox", (c.x0, c.y0, c.x1, c.y1), _cut_option(right) or right,
                              "小さな正方形のマスの右に選択肢" + size_note, cell=c, line_text=right, host=c,
                              kind_candidates=["checkbox", "text"], size=_fit_size(c.h, typ))
                self.spots.append(_agent(sp, gate="kind"))
                continue
            if c.w < 8 or c.h < 6:
                undecided.append((c, f"小さな空のマス（幅 {c.w:.1f}pt・高さ {c.h:.1f}pt）", "small_cell"))
                continue
            if any(o is not c and o.x0 >= c.x0 - 0.5 and o.x1 <= c.x1 + 0.5 and o.y0 >= c.y0 - 0.5 and o.y1 <= c.y1 + 0.5
                   for o in self.index.near(c.x0, c.y0, c.x1, c.y1)):
                # 中に小さなマス（□ など）を持つマスは、入れ物のことも、書く欄のこともある
                undecided.append((c, "中に小さなマスを持つ空のマス（入れ物か、書く欄か）", "container"))
                continue
            fill = self.cell_fill(c)
            gray = fill is not None and max(fill) - min(fill) < 0.04 and max(fill) < 0.97
            if gray:
                # 灰色に塗ったマスは「書かない欄」の印のことも、書く欄のこともある：規則では決めない
                undecided.append((c, "灰色に塗った空のマス（書かない欄の印のことも、書く欄のこともある）", ""))
                continue
            label, rl, cl = self._label_for(c, c)
            if not label:
                undecided.append((c, "見出しの見つからない空のマス", ""))
                continue
            sp = Spot("text", (c.x0, c.y0, c.x1, c.y1), label, "空のマス", row_label=rl, column_label=cl,
                      cell=c, size=_fit_size(c.h, typ), host=c)
            if fill is not None and max(fill) < 0.995:
                sp.reason += "（色の付いたマス）"
            self.spots.append(sp)
            used_cells.add(id(c))
        for c in self._big_empty:
            undecided.append((c, f"ページの半分を超える大きな空のマス（幅 {c.w:.0f}pt・高さ {c.h:.0f}pt。中に字もマスもない）",
                              "big_cell"))
        # 3) マスの中の見出しの右に、大きな空きがあるもの（1行のマスでも、見出しが何行かあるマスでも、行ごとに）
        for c in cells:
            if not c.visible():
                continue
            hor = [x for x in c.lines if not x.vertical]
            for ln in hor:
                self._right_room(c, ln, hor, typ)
            self._rest_of_cell(c, typ)
        # 4) 下線
        self._underline_spots()
        # ほかの欄と大きく重なる候補は、消さずに「同じ所の別の読み方」として残す（先に見つけたものを主にする）
        kept: list[Spot] = []
        for s in self.spots:
            if s.kind != "checkbox":
                k = next((k for k in kept if k.kind != "checkbox" and _overlap(s.rect, k.rect) > 0.6), None)
            else:
                k = next((k for k in kept if k.kind == "checkbox" and _overlap(s.rect, k.rect) > 0.7), None)  # 同じ □ を2重に
            if k is not None:
                s.alternative_to = k
                _agent(s, gate="overlap")  # 重なる先は alternative_to で示す（窓の種類の欄にも出す）
                continue
            kept.append(s)
        # 5) 規則では決められない所（窓に「？」をつけて示し、エージェントが読んで判断する）
        self._agent_spots(cells, undecided)
        # 6) ページの外：見えるのはページの枠（CropBox を MediaBox に切り詰めたもの）の中だけ。欄の四角を枠に切り詰め、
        #    枠の外にしかないもの（ページの外に描いた線・字から見つけたもの）は欄にしない
        self._clip_to_page()
        # 7) 斜線で消したマスは欄にしない。印刷された字を含む欄は、字のない所に縮める（縮めると書けない大きさになるものは、欄にしない）
        self.spots = [sp for sp in self.spots if not self._struck(sp.rect) and self._clear_of_print(sp)]


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
                if s.group:
                    s.group = [x for x in s.group if id(x) not in gone]
        self.spots = kept

    def _agent_spots(self, cells: list[Cell], undecided: list[tuple[Cell, str, str]]) -> None:  # noqa: C901
        """規則の型に合わないが、書き込む所かもしれない所：行の中の空き・見出しの後ろ・単位だけのマスなど
        （xlsx.window.agent_blanks と同じ見つけ方）と、見出しの見つからない空のマス・灰色の空のマス。
        先に見つけた候補と重なるものは、消さずに alternative_to（同じ所の別の読み方）として示す。"""
        from ..xlsx.window import AGENT_WHY, agent_blanks

        typ = self.typical
        primary = [k for k in self.spots if k.alternative_to is None]
        new: list[Spot] = []

        def over(rect, lim: float = 0.3):
            return next((k for k in primary + [x for x in new if x.alternative_to is None] if _overlap(rect, k.rect) > lim), None)

        def add(sp: Spot, gate: str = "") -> None:
            k = over(sp.rect)
            if k is not None:
                sp.alternative_to = k
                pass  # 重なる先は alternative_to で示す
                sp.gate = sp.gate or "overlap"
            elif gate:
                sp.gate = gate
            new.append(sp)

        for cell, lines in [(c, c.lines) for c in cells] + [(None, self.free_lines)]:
            for ln in lines:
                if ln.vertical or not ln.text.strip(" 　"):
                    continue
                small = ln.y1 - ln.y0 < MIN_WRITE  # 小さな字の行：書けるかは、書き込みと同じ決め方で確かめる（入らなければ書けない所）
                text = ln.text
                sz = ln.size
                found = [k.at for k in self.spots if k.host is ln and k.at is not None]
                # 同じ行の同じ位置に、もう候補がある（見出しの右の余白・マスの右の空き）：同じ所なので、2つ目は出さない
                ends = {k.at for k in self.spots + new if k.host is ln and k.at is not None}
                # 左右の端：マスの縁（マスの外は、ページの字の端）・同じ高さのほかの字・縦の罫線
                lb = cell.x0 if cell is not None else self.left
                rb = cell.x1 if cell is not None else self.right
                for o in (cell.lines if cell is not None else self.free_lines):
                    if o is ln or o.vertical or not (o.y0 < ln.y1 - 0.2 * sz and o.y1 > ln.y0 + 0.2 * sz):
                        continue
                    if o.x1 <= ln.x0 + 0.5:
                        lb = max(lb, o.x1)
                    if o.x0 >= ln.x1 - 0.5:
                        rb = min(rb, o.x0)
                if cell is None:
                    for r in self.vs:
                        if r.a - 1 <= ln.y0 and r.b + 1 >= ln.y1:
                            if ln.x1 <= r.pos < rb:
                                rb = r.pos
                            if lb < r.pos <= ln.x0:
                                lb = r.pos
                core_end = len(text.rstrip(" 　"))
                beside = None if cell is None else next(
                    (k for k in primary if k.kind != "space" and k.rect[0] >= cell.x1 - 1.5 and k.rect[0] - cell.x1 < 6 and _vov_rect(k.rect, cell) >= 0.5),
                    None)
                for a, b, why in agent_blanks(text, found, cell is not None):
                    gate = "small_line" if small else ""
                    if why == "placeholder":
                        # PDF の字は消せない：伏せ字の上を白い下地の欄で覆って書く（元の字は残る）
                        x0, y0, x1, y1 = ln.span_rect(a, b)
                        rect = (x0 - 0.3, y0 - 0.1 * sz, x1 + 0.3, y1 + 0.1 * sz)
                        add(Spot("inline", rect, norm_label(text[:a] + "［　］" + text[b:]), AGENT_WHY[why],
                                 line_text=text, blank=(a, b), cell=cell, size=sz, host=ln, at=(a, b),
                                 confidence="agent", cover=text[a:b]), gate)
                        continue
                    if (a, b) in ends or (a >= core_end and (len(text), len(text)) in ends):
                        continue  # 同じ所の候補が、もうある
                    reason = AGENT_WHY[why]
                    if why == "after_label" and cell is None:
                        gate = gate or "free_after_label"  # これまでは「見出しの右の余白」のときだけ示していた
                    alt = None
                    if why in ("after_label", "after_colon", "after_mark") and beside is not None:
                        alt = beside  # すぐ右のマスに規則の欄がある：値はそちらに書くことが多い
                    if a == b == 0 or (b <= len(text) - len(text.lstrip(" 　")) and a == 0):
                        x0, x1 = lb + 1.5, ln.x0 - 0.2 * sz  # 行の頭の前（マスの縁まで）
                    elif a >= core_end:
                        x0, x1 = ln.x1 + 0.2 * sz, rb - 1.5  # 行の終わりの後ろ（マスの縁・次の字まで）
                    elif a == b:
                        continue
                    else:
                        x0, _, x1, _ = ln.span_rect(a, b)
                    if x1 <= x0:
                        continue  # 空きがまったくない（縁や次の字に接している）
                    if x1 - x0 < 0.8 * sz:
                        reason += f"（空きは幅 {x1 - x0:.1f}pt：字 1 字（{sz:.1f}pt）より狭い）"
                        gate = gate or "narrow"
                    rect = (x0, ln.y0, x1, ln.y1)
                    sp = Spot("inline", rect, norm_label(text[:a] + "［　］" + text[b:]), reason,
                              line_text=text, blank=(a, b), cell=cell, size=sz, host=ln, at=(a, b), confidence="agent")
                    if alt is not None:
                        sp.alternative_to = alt
                        sp.reason += "（すぐ右のマスの欄の別の候補）"
                        sp.gate = gate or "beside"
                        new.append(sp)
                    else:
                        add(sp, gate)
        for c, why, gate in undecided:
            rect = (c.x0, c.y0, c.x1, c.y1)
            label, rl, cl = self._label_for(c, c)
            add(Spot("text", rect, label, why, row_label=rl, column_label=cl, cell=c, size=_fit_size(c.h, typ),
                     host=c, confidence="agent"), gate)
        self.spots += new

    def _rest_of_cell(self, c: Cell, typ: float) -> None:
        """字がマスの一部にだけあり、残りに書ける空きがあるマス。
        - 字が上の方にだけある：下の空き（「〒」の下に住所を書くなど）
        - 字が下の方にだけある：上の空き（「区」の上に区の名前を書くなど）
        - 字が右の方にだけある：左の空き（「□男 □女」の左に名前を書くなど）"""
        hor = [x for x in c.lines if not x.vertical]
        if not hor:
            return
        if any(sp.cell is c and sp.kind in ("text", "comb") for sp in self.spots):
            return  # このマスには、もう書く所がある
        bottom = max(x.y1 for x in hor)
        top = min(x.y0 for x in hor)
        left = min(x.x0 for x in hor)
        sz = max(x.size for x in hor)
        below = c.y1 - bottom
        above = top - c.y0
        leftroom = left - c.x0
        label, rl, _cl = self._label_for(c, c)
        own = norm_label(join_lines(hor))
        name = f"{rl}（{own}）" if rl and own else (rl or own)
        if c.w < 24:
            pass  # 細いマス：下の「規則の型に合わない空き」として示す
        elif below >= max(1.3 * sz, 12) and below >= 0.4 * c.h and c.w >= 60:
            self.spots.append(Spot("text", (c.x0 + 1.5, bottom + 0.5, c.x1 - 1.5, c.y1 - 0.5), name,
                                   "マスの中の字の下の空き", cell=c, size=_fit_size(min(below, 2.5 * sz), typ), host=c))
        elif above >= max(1.3 * sz, 10) and above >= 0.4 * c.h:
            self.spots.append(Spot("text", (c.x0 + 1.5, c.y0 + 0.5, c.x1 - 1.5, top - 0.5), name,
                                   "マスの中の字の上の空き", cell=c, size=_fit_size(min(above, 2.5 * sz), typ), host=c))
        elif leftroom >= max(48, 5 * sz) and leftroom >= 0.4 * c.w and not any(
                pdf_blanks("\u3000" * 3 + ln.text) for ln in hor) and not (
                all(getattr(ln, "rtl", False) for ln in hor) and self._empty_beside(c, "left")):
            self.spots.append(Spot("text", (c.x0 + 1.5, c.y0 + 0.5, left - 1.5, c.y1 - 0.5), rl or name,
                                   "マスの中の字の左の空き", cell=c, size=_fit_size(c.h, typ), host=c))
        else:
            # 規則の型（空きの大きさ・マスの幅）に合わない：欄にはせず、いちばん大きな空きの向きと大きさを窓の印で示す
            # （エージェントが字を目印にして書く。その行の字 1 字も入らない空きは、空きではない）
            rooms = [(below, "↓", c.w - 3), (above, "↑", c.w - 3), (leftroom, "←", c.h - 1)]
            depth, arrow, other = max(rooms, key=lambda r: r[0] * r[2])
            if depth >= sz + 1 and other >= sz:
                self.room_marks.append((c, arrow, depth))

    def _empty_beside(self, c: Cell, side: str) -> bool:
        """マスのすぐ左（または右）に、空のマスがあるか。"""
        if side == "left":
            return any(abs(o.x1 - c.x0) <= 1.5 and _vov(o, c) >= 0.5 and not o.visible()
                       for o in self.index.near(c.x0 - 5, c.y0, c.x0, c.y1))
        return any(abs(o.x0 - c.x1) <= 1.5 and _vov(o, c) >= 0.5 and not o.visible()
                   for o in self.index.near(c.x1, c.y0, c.x1 + 5, c.y1))

    def _left_room(self, c: Cell, ln: Line, hor: list, typ: float) -> None:
        """右から左へ書く見出し（アラビア文字など）：見出しと同じマスの、見出しの左の空き（_right_room を左右に映したもの）。"""
        sz = ln.size
        others = [o for o in hor if o is not ln]
        if any(o.y0 < ln.y1 - 0.2 * sz and o.y1 > ln.y0 + 0.2 * sz and o.x1 <= ln.x0 for o in others):
            return
        if len(hor) == 1:
            band = (c.y0, c.y1)
            ref = max(sz, 0.75 * typ)
            if c.h > 3.2 * ref or c.h < 1.1 * sz:
                return
        else:
            above = [o.y1 for o in others if o.y1 <= ln.y0 + 0.2 * sz]
            below = [o.y0 for o in others if o.y0 >= ln.y1 - 0.2 * sz]
            band = (max(above + [c.y0]), min(below + [c.y1]))
        room = ln.x0 - c.x0
        if self._empty_beside(c, "left"):
            return
        t = ln.text.strip(" \u3000")
        right_ok = c.x1 - ln.x1 <= 3 * sz and room >= 0.45 * c.w
        if right_ok and room >= max(48, 5 * sz) and len(t) <= 24 and not pdf_blanks(ln.text) \
                and not any(b in ln.text for b in BOX_GLYPHS):
            self.spots.append(Spot("text", (c.x0, band[0], ln.x0 - 1.5, band[1]), norm_label(t),
                                   "見出しと同じマスの左の空き（右から左へ書く見出し）", cell=c,
                                   size=_fit_size(band[1] - band[0], typ), host=ln, at=(len(ln.text), len(ln.text))))

    def _right_room(self, c: Cell, ln: Line, hor: list, typ: float) -> None:
        if getattr(ln, "rtl", False):
            return self._left_room(c, ln, hor, typ)
        sz = ln.size
        others = [o for o in hor if o is not ln]
        if any(o.y0 < ln.y1 - 0.2 * sz and o.y1 > ln.y0 + 0.2 * sz and o.x0 >= ln.x1 for o in others):
            return  # 同じ高さの右に、ほかの字がある（その間の空きは、行の中の空きとして見る）
        if any(abs(o.x0 - c.x1) <= 1.5 and _vov(o, c) >= 0.5 and not o.visible()
               for o in self.index.near(c.x1, c.y0, c.x1 + 5, c.y1)):
            return  # すぐ右に空のマスがある（値はそのマスに書く。そのマスは空のマスとして示す）
        gate_fail = []
        if len(hor) == 1:
            band = (c.y0, c.y1)
            ref = max(sz, 0.75 * typ)  # 「フリガナ」のような小さな見出しでも、書く字はふつうの大きさ
            if c.h > 3.2 * ref or c.h < 1.1 * sz:
                gate_fail.append(f"マスの高さ {c.h:.1f}pt・字 {sz:.1f}pt")
        else:
            above = [o.y1 for o in others if o.y1 <= ln.y0 + 0.2 * sz]
            below = [o.y0 for o in others if o.y0 >= ln.y1 - 0.2 * sz]
            band = (max(above + [c.y0]), min(below + [c.y1]))
        room = c.x1 - ln.x1
        t = ln.text.strip(" 　")
        # 大きな空きで分けた最後のまとまり（「上記のとおり証明します。　　　　委託者　事　業　所　の　住　所」の「委託者…」）
        seg = re.split(r"[ 　]{3,}", t)[-1]
        labelish = bool(_SPACED.fullmatch(seg) or re.search(r"(?:\S[ 　]){2,}\S$", seg)) or seg.endswith(("：", ":"))
        left_ok = ln.x0 - c.x0 <= 3 * sz and room >= 0.45 * c.w
        rect = (ln.x1 + 1.5, band[0], c.x1, band[1])
        if not gate_fail and (left_ok or (labelish and len(hor) >= 2 and room >= max(80, 8 * sz))) and room >= max(48, 5 * sz) \
                and len(seg if labelish else t) <= 16 + (8 if labelish else 0) \
                and not pdf_blanks(ln.text) and not any(b in ln.text for b in BOX_GLYPHS) \
                and not _OPTION_GROUP.search(ln.text):
            self.spots.append(Spot("text", rect, norm_label((seg if labelish else t)),
                                   "見出しと同じマスの右の空き", cell=c, size=_fit_size(band[1] - band[0], typ), host=ln,
                                   at=(len(ln.text), len(ln.text))))
            return
        # 規則の型（見出しの長さ・空きの幅・マスの高さ・選択肢や空欄のない行）に合わない：欄にはせず、空きの幅を窓の印で示す
        if rect[2] - rect[0] >= sz and rect[3] - rect[1] >= sz:
            self.room_marks.append((ln, "→", rect[2] - rect[0]))

    def _group_stacked(self, start: int) -> None:
        """同じマスの上下の行で、同じ横の位置にある空欄（「都 道／府 県」の前の空き）は、1つの欄かもしれない組として示す
        （1つにまとめず、それぞれの id のまま。どちらに書くか・まとめて1つの値かは、エージェントが判断する）。"""
        sp = self.spots[start:]
        inl = [x for x in sp if x.kind == "inline"]
        for i, a in enumerate(inl):
            for b in inl[i + 1:]:
                if abs(a.rect[0] - b.rect[0]) <= 3 and abs(a.rect[2] - b.rect[2]) <= 3 and 0 <= b.rect[1] - a.rect[3] <= 0.8 * a.size:
                    grp = a.group or [a]
                    if not any(x is b for x in grp):
                        grp.append(b)
                    for x in grp:
                        x.group = grp

    def _free_line_spots(self, ln: Line) -> None:
        """マスの外の行の、語を手がかりにした空き（規則では決められない所として示す）：
        行の頭の空き（「　　知事　殿」）・行の終わりの空き（「〒」「市」の後ろ）・見出しの右の余白。"""
        if ln.vertical:
            return
        t = ln.text.strip(" 　")
        sz = ln.size
        if not t:
            return
        # 同じ高さのほかの字（空白の字は、見えないので数えない）
        same_row = [c for c in self.chars if abs(c.cy - (ln.y0 + ln.y1) / 2) < 0.6 * sz and not c.text.isspace()
                    and c not in ln.chars]
        # 行の頭の空き：左に何もなければ、4字ぶんまで（「知事」「市長」「年」などの語で始まる行）
        if any(t.startswith(u) for u in LEAD_UNITS):
            left_bound = max([c.x1 for c in same_row if c.x1 <= ln.x0 + 0.5] + [self.left])
            for r in self.vs:
                if r.a - 1 <= ln.y0 and r.b + 1 >= ln.y1 and left_bound < r.pos <= ln.x0:
                    left_bound = r.pos
            room = ln.x0 - left_bound
            if room >= 1.3 * sz:
                w = min(room - 0.5 * sz, 5 * sz)
                why = "行の頭の空き（「" + next(u for u in LEAD_UNITS if t.startswith(u)) + f"」で始まる行。左の空き {room:.1f}pt）"
                self.spots.append(_agent(Spot("inline", (ln.x0 - w, ln.y0, ln.x0 - 0.3 * sz, ln.y1),
                                              norm_label("［　］" + t), why, line_text=ln.text, size=sz,
                                              host=ln, at=(0, 0)), gate="" if room >= 2.5 * sz else "lead_room"))
        right_bound = min([c.x0 for c in same_row if c.x0 >= ln.x1 - 0.5] + [self.right])
        for r in self.vs:
            if r.a - 1 <= ln.y0 and r.b + 1 >= ln.y1 and ln.x1 <= r.pos < right_bound:
                right_bound = r.pos
        room = right_bound - ln.x1
        # 行の終わりの空き：「住所　〒」「仙台市」「電話番号　　―　　―」の後ろ（同じ高さの次の字か、右の縁まで）
        last = t[-1:]
        if last == "〒" or last in DASHES or (last in ADDRESS_UNITS and len(t) <= 12):
            if room >= 3 * sz:
                self.spots.append(_agent(Spot("inline", (ln.x1 + 0.3 * sz, ln.y0, right_bound - 1, ln.y1),
                                              norm_label(t + "［　］"), f"行の終わりの空き（「{last}」の後ろ。空き {room:.1f}pt）",
                                              line_text=ln.text, size=sz, host=ln, at=(len(ln.text), len(ln.text)))))
                return
        if getattr(ln, "rtl", False):
            # 右から左へ書く見出し：見出しの左の余白（左に何もない）
            if 1 <= len(t) <= 15 and not pdf_blanks(ln.text):
                left_bound = max([c.x1 for c in same_row if c.x1 <= ln.x0 + 0.5] + [self.left])
                for r in self.vs:
                    if r.a - 1 <= ln.y0 and r.b + 1 >= ln.y1 and left_bound < r.pos <= ln.x0:
                        left_bound = r.pos
                if ln.x0 - left_bound >= max(80, 8 * sz) and left_bound <= self.left + 1:
                    self.spots.append(Spot("space", (left_bound, ln.y0 - 0.2 * sz, ln.x0 - sz, ln.y1 + 0.2 * sz),
                                           norm_label(t), "見出しの左の余白（右から左へ書く見出し。入力欄とは限らない）", size=sz,
                                           host=ln, at=(len(ln.text), len(ln.text)), confidence="agent"))
            return
        # 見出しの右の余白（短い見出しで、右に何もない）
        center = (ln.x0 + ln.x1) / 2
        aligned = any(o is not ln and not o.vertical and abs(o.x0 - ln.x0) <= 1.0 and abs(o.y0 - ln.y0) < 12 * sz
                      for o in self.free_lines)
        centered = abs(center - (self.left + self.right) / 2) < 0.02 * (self.right - self.left) and not aligned
        spaced = bool(re.fullmatch(r"(?:\S[　 ]+){2,}\S", t))  # 「事　業　所　の　住　所」字間を空けた見出し
        tight = not spaced and any(o is not ln and not o.vertical and abs(o.x0 - ln.x0) <= 1.0 and
                                   (0 <= o.y0 - ln.y1 < 0.6 * sz or 0 <= ln.y0 - o.y1 < 0.6 * sz) for o in self.free_lines)
        rect = (ln.x1 + sz, ln.y0 - 0.2 * sz, right_bound, ln.y1 + 0.2 * sz)
        if 1 <= len(t) <= 15 and not t.endswith(("。", "殿", "様", "御中")) and not pdf_blanks(ln.text) \
                and room >= max(80, 8 * sz) and not centered and (right_bound >= self.right - 1 or spaced) and not tight:
            self.spots.append(Spot("space", rect, norm_label(t), "見出しの右の余白（入力欄とは限らない）", size=sz, host=ln,
                                   at=(len(ln.text), len(ln.text)), confidence="agent"))
        elif rect[2] - rect[0] >= sz:
            # 見出しの右の余白の型（短い見出し・広い空き・中央でない・詰まった行でない）に合わない行の右の空き（文の終わりの
            # 余白のことも多い）：欄にはせず、空きの幅を窓の印で示す（エージェントが字を目印にして書く）
            self.room_marks.append((ln, "→", rect[2] - rect[0]))

    def _line_spots(self, ln: Line, cell: Cell | None) -> None:
        text = ln.text
        # チェックの印（□ の字）：右の文字を選択肢の見出しにする（☑ などの記入済みの字は、説明の文の中の記号のことが多い）
        marks = [i for i, ch in enumerate(text) if ch in BOX_GLYPHS]
        for n, i in enumerate(marks):
            end = marks[n + 1] if n + 1 < len(marks) else len(text)
            opt = _cut_option(text[i + 1:end])
            if not opt:
                # □ の右に文字がない：左の文字（「申請時に □」）を見出しにする
                opt = _cut_option(text[:i][::-1])[::-1] if i else ""
            x0, y0, x1, y1 = ln.span_rect(i, i + 1)
            self.spots.append(Spot("checkbox", (x0, y0, x1, y1), norm_label(opt) or "（選択肢の文字なし）", "□ の字",
                                   line_text=text, cell=cell, size=ln.size, host=ln, at=(i, i + 1)))
        # 記入済みの印（☑・■・✓ の字）：選んだ印のことも、箇条の印・説明の中の記号のこともある。PDF の字は消さないので
        # 書き換えられない所として、選んであるかもしれない印を示す
        for i, ch in enumerate(text):
            if ch not in CHECKED_GLYPHS:
                continue
            opt = _cut_option(text[i + 1:]) or (_cut_option(text[:i][::-1])[::-1] if i else "")
            x0, y0, x1, y1 = ln.span_rect(i, i + 1)
            sp = Spot("checkbox", (x0, y0, x1, y1), norm_label(opt) or "（選択肢の文字なし）",
                      f"記入済みの印「{ch}」（選んだ印か、箇条の印・説明の記号か）", checked=True, line_text=text, cell=cell,
                      size=ln.size, host=ln, at=(i, i + 1))
            sp.readonly = f"印刷された記入済みの印「{ch}」は消せないので、書き換えられません"
            self.spots.append(_agent(sp, gate="checked_glyph"))
        if ln.vertical:
            return
        # マスの中の行は、マスの縁（または同じ高さのほかの文字）までの空きも空白として扱う（「　　　都 道」「市 区　　　」）
        lead = trail = 0
        lb = rb = None
        if cell is not None:
            lb, rb = cell.x0, cell.x1
            for o in cell.lines:
                if o is ln or o.vertical or not (o.y0 < ln.y1 - 0.2 * ln.size and o.y1 > ln.y0 + 0.2 * ln.size):
                    continue
                if o.x1 <= ln.x0 + 0.5:
                    lb = max(lb, o.x1)
                if o.x0 >= ln.x1 - 0.5:
                    rb = min(rb, o.x0)
            # 縁までの空きが 1.5 字ぶん以上あれば空白にする（空欄にするかは pdf_blanks の強い手がかりで決める）
            lead = max(2, int((ln.x0 - lb - 1.5) // ln.size)) if ln.x0 - lb >= 1.5 * ln.size else 0
            trail = max(2, int((rb - ln.x1 - 1.5) // ln.size)) if rb - ln.x1 >= 1.5 * ln.size else 0
        padded = "\u3000" * lead + text + "\u3000" * trail
        for blank in pdf_blanks(padded, lead, trail):
            a, b = blank
            a0, b0 = a - lead, b - lead
            if a0 < 0 and b0 <= 0:
                x0, x1 = lb + 1.5, ln.x0 - 0.2 * ln.size
            elif a0 >= len(text):
                x0, x1 = ln.x1 + 0.2 * ln.size, rb - 1.5
            else:
                aa, bb = max(a0, 0), min(b0, len(text))
                if bb <= aa:
                    continue
                x0, _, x1, _ = ln.span_rect(aa, bb)
                if a0 < 0:
                    x0 = lb + 1.5
                if b0 > len(text):
                    x1 = rb - 1.5
            if x1 <= x0:
                continue  # 空きがまったくない（縁や次の字に接している）
            at = (max(a0, 0), max(0, min(b0, len(text))))
            orig = text[at[0]:at[1]]
            shown = text[:at[0]] + (orig if orig.strip(" 　") else "") + "［　］" + text[at[1]:]
            sp = Spot("inline", (x0, ln.y0, x1, ln.y1), norm_label(shown), "文字の中の空欄", line_text=text,
                      blank=at, cell=cell, size=ln.size, host=ln, at=at)
            if not blank.structural and text[at[0]:at[1]] and all(ch in "_＿" for ch in text[at[0]:at[1]]):
                # 英数字に接した下線の字（「Name____」「_HVeg」）：書くための線か、名前の一部か、字の形だけでは決まらない
                _agent(sp, "英数字に接した下線の字（" + _blank_hint(text, at) + "。名前の一部のこともある）")
            elif not blank.structural:
                # 語（単位・年月日・住所の単位・「：」など）を手がかりにした空欄：規則では欄と決められない
                _agent(sp, "語を手がかりにした文字の中の空欄（" + _blank_hint(text, at) + "）")
            if x1 - x0 < 0.8 * ln.size:
                _agent(sp, f"{sp.reason}（空きは幅 {x1 - x0:.1f}pt：字 1 字（{ln.size:.1f}pt）より狭い）", "narrow")
            self.spots.append(sp)
        self._era_choice(ln, cell)
        self._option_choice(ln, cell)

    def _option_choice(self, ln: Line, cell: Cell | None) -> None:
        """○で囲んで選ぶ選択肢の候補（規則では決められない所として示す。どれも字の並びの形だけで見つける）：
        - かっこの中を「・」「／」で区切った並び（「（父携帯・母携帯・自宅）」）。ただし、一つだけが句（かなの後ろに漢字の続く
          「通勤を除く」）の並びは、注のかっこなので除く
        - 選択肢だけの行（「有　・　無」）：区切りの前後を空けて並べたもの、またはマスの中にその行だけがあり、同じ高さの左の
          マスに字（見出し）があるもの
        - 番号の付いた並び（「1.新規　2.更新」「１　自家　２　借家」）：続き番号が 2 つ以上
        - 数だけの並び（「区分　１　２　３　４　５　６」）：空白で区切った続き番号が 3 つ以上
        どの選択肢も、ほかの欄の候補（文字の中の空欄など）に重なるものは除く（「年　月　日交付・更新」の「年　月　日交付」）。"""
        if ln.vertical:
            return
        text = ln.text
        groups = []
        for m in re.finditer(r"[（(]([^（）()]{2,60})[）)]", text):
            groups.append((m.start(1), m.group(1), True))
        whole = text.strip(" \u3000")
        # 後ろの「その他（　　）」の書く所は選択肢に含めない
        body = re.sub(r"[（(][ \u3000]*[）)]$", "", whole).rstrip(" \u3000")
        spaced_sep = bool(re.search(r"[ \u3000]+[・／/][ \u3000]+", body))
        if body and _OPTION_GROUP.fullmatch(body) and (spaced_sep or self._alone_with_label(ln, cell)):
            groups.append((text.index(body), body, False))
        for start, inner, paren in groups:
            parts = re.split(r"([・／/])", inner)
            opts = []
            pos = start
            for k, part in enumerate(parts):
                if k % 2 == 0:
                    word = part.strip(" \u3000")
                    if word:
                        a = pos + part.index(word)
                        opts.append((word, a, a + len(word)))
                pos += len(part)
            if not (2 <= len(opts) <= 6) or any(len(w) > 10 or re.search(r"[、。：:]", w) for w, _a, _b in opts):
                continue
            if any(re.fullmatch(r"[\d０-９]+", w) for w, _a, _b in opts):
                continue
            phrase = [bool(re.search(r"[\u3041-\u309f][\u3400-\u9fff]", w)) for w, _a, _b in opts]
            if paren and any(phrase) and not all(phrase):
                continue  # 注のかっこ（「（通学・通所・通勤を除く）」）：選択肢の並びでなく、句が混ざる
            head = text[:start].rstrip("（( \u3000")
            self._add_options(ln, cell, opts, head, "選択肢が「・」で並ぶ（どれかを○で囲む）")
        self._numbered_options(ln, cell)

    def _alone_with_label(self, ln: Line, cell: Cell | None) -> bool:
        """マスの中にその行だけがあり、同じ高さの左のマスに字がある（見出しの右の、選択肢だけのマス）。"""
        if cell is None or len([x for x in cell.lines if not x.vertical]) != 1:
            return False
        for c in self.index.near(cell.x0 - 300, cell.y0, cell.x0, cell.y1):
            if c is not cell and c.x1 <= cell.x0 + 1.5 and _vov(c, cell) >= 0.5 and c.visible():
                return True
        return False

    def _numbered_options(self, ln: Line, cell: Cell | None) -> None:
        text = ln.text
        if any(b in text for b in BOX_GLYPHS):
            return  # □ の付いた選択肢（チェックの印として示す）

        def num(t: str) -> int:
            return int(t.translate(str.maketrans("０１２３４５６７８９", "0123456789")))

        # 番号の付いた並び：番号（と「.」「．」「)」など）の後ろに語
        # （番号の前は、行の頭・空白・かっこ・区切りの字だけ。「第８条」の８は番号でない。○は番号を囲む）
        toks = [(num(m[1]), m.start(1), m.end(1), m[0].strip(" \u3000")) for m in re.finditer(
            r"(?:^|(?<=[\s\u3000（(「・／/]))([1-9１-９])[ \u3000]?[.．、,，)）:：]?[ \u3000]?[^\s\u3000\d０-９・／/、。（）()]{1,8}", text)]
        run = self._consecutive(toks)
        if len(run) >= 2:
            opts = [(t[3], t[1], t[2]) for t in run]
            self._add_options(ln, cell, opts, text[:run[0][1]], "番号の付いた選択肢が並ぶ（番号を○で囲む）", numbered=True)
            return
        # 数だけの並び：空白で区切った続き番号
        toks = [(num(m[0]), m.start(), m.end(), m[0]) for m in re.finditer(r"(?<![\d０-９\w])[\d０-９]{1,2}(?![\d０-９\w])", text)]
        run = self._consecutive(toks)
        if len(run) >= 3 and all(re.fullmatch(r"[ \u3000]+", text[a[2]:b[1]]) for a, b in zip(run, run[1:])):
            opts = [(t[3], t[1], t[2]) for t in run]
            self._add_options(ln, cell, opts, text[:run[0][1]], "数が空白で区切って並ぶ（どれかを○で囲む）")

    @staticmethod
    def _consecutive(toks: list) -> list:
        """続き番号のいちばん長い並び（1つずつ増える）。"""
        best: list = []
        cur: list = []
        for t in toks:
            if cur and t[0] == cur[-1][0] + 1:
                cur.append(t)
            else:
                cur = [t]
            if len(cur) > len(best):
                best = list(cur)
        return best

    def _add_options(self, ln: Line, cell: Cell | None, opts: list, head: str, why: str, numbered: bool = False) -> None:
        rects = [(w, ln.span_rect(a, b)) for w, a, b in opts]
        for k in self.spots:
            if k.kind == "circle" and k.host is ln and any(_overlap(r, o) > 0.3 for _w, r in rects for _t, o in k.options):
                return  # 同じ字を、ほかの決め方でもう選択肢にした
            if k.kind in ("inline", "text") and k.host is ln and any(_overlap(r, k.rect) > 0.05 for _w, r in rects):
                return  # 選択肢が、ほかの欄の候補（文字の中の空欄）に重なる
        label = norm_label(head.strip(" \u3000（(")[-30:])
        if not label and cell is not None:
            lab, rl, _cl = self._label_for(cell, cell)
            label = norm_label(rl or lab)
        label = label or "選択肢"
        r = (min(x[1][0] for x in rects), min(x[1][1] for x in rects), max(x[1][2] for x in rects), max(x[1][3] for x in rects))
        sp = Spot("circle", r, f"{label}（○で囲む）", why, cell=cell, host=ln, at=None, size=ln.size, confidence="agent")
        sp.options = rects
        sp.numbered = numbered  # 選択肢の名前は「番号と語」、○は番号を囲む
        self.spots.append(sp)

    def _era_choice(self, ln: Line, cell: Cell | None) -> None:
        """元号を○で囲んで選ぶ欄（「昭和／平成」を上下や左右に並べたもの）。マスの中の元号の字を集めて、1つの選ぶ欄にする。"""
        if cell is None or getattr(cell, "_era_done", False):
            return
        eras = []
        for o in cell.lines:
            for m in re.finditer(r"明治|大正|昭和|平成|令和", o.text):
                if o.vertical:
                    continue
                eras.append((m[0], o.span_rect(m.start(), m.end()), o))
        names = [e[0] for e in eras]
        if len(set(names)) < 2 or len(names) != len(set(names)):
            return
        cell._era_done = True
        r = (min(e[1][0] for e in eras), min(e[1][1] for e in eras), max(e[1][2] for e in eras), max(e[1][3] for e in eras))
        sp = Spot("circle", r, "元号（○で囲む）", "元号が並ぶ（どれかを○で囲む）", cell=cell, host=eras[0][2],
                  at=None, size=ln.size, confidence="agent")
        sp.options = [(n, rect) for n, rect, _ in eras]
        self.spots.append(sp)

    def _underline_spots(self) -> None:
        """下線：見出しの右に引かれた線で、上に文字のないもの。見出しが近くにない線・短い線も、
        マスの縁でなければ、測った長さを添えて候補にする（規則では決められない所）。"""
        lines = self.free_lines + [ln for c in self.cells for ln in c.lines]
        for h in self.hs:
            ln_len = h.b - h.a
            if ln_len > 0.9 * self.view.width:
                continue  # ページの幅いっぱいの線は、区切りの線
            if self.index.at((h.a + h.b) / 2, h.pos - 3) is not None and self.index.at((h.a + h.b) / 2, h.pos + 3) is not None:
                continue  # 表の中の罫線
            # 下線の左（同じ高さ）に見出しの文字がある
            label = None
            near = None  # 左にある、いちばん近い行（遠くても）
            for ln in lines:
                if ln.vertical:
                    continue
                if abs(ln.y1 - h.pos) <= 0.6 * ln.size and ln.x1 <= h.a + 0.5 * ln.size:
                    if h.a - ln.x1 <= 6 * ln.size and (label is None or ln.x1 > label.x1):
                        label = ln
                    if near is None or ln.x1 > near.x1:
                        near = ln
            sz = label.size if label is not None else (near.size if near is not None else self.typical)
            top = h.pos - 1.3 * sz
            if any(c.cx > h.a and c.cx < h.b and top < c.cy < h.pos and not c.text.isspace() for c in self.chars):
                continue  # 線の上に字がある（字に引いた下線）
            if label is not None and ln_len >= 24:
                self.spots.append(Spot("text", (h.a, top, h.b, h.pos - 0.5), norm_label(label.text.strip()), "見出しの右の下線",
                                       size=sz, host=label, at=(len(label.text), len(label.text))))
                continue
            if self._cell_edge(h):
                continue  # マスの縁（表の外側の線）：下線ではない
            notes = [f"線の長さ {ln_len:.1f}pt"]
            if label is None:
                notes.append("左に見出しの字がない" if near is None else f"左の字「{norm_label(near.text)[-12:]}」まで {h.a - near.x1:.0f}pt")
            host = label if label is not None else None
            sp = Spot("text", (h.a, top, h.b, h.pos - 0.5), norm_label((label or near).text.strip())[-30:] if (label or near) else "",
                      "下線（" + "・".join(notes) + "）", size=sz, host=host,
                      at=(len(label.text), len(label.text)) if label is not None else None)
            self.spots.append(_agent(sp, gate="underline"))

    def _cell_edge(self, h: Rule) -> bool:
        """横線 h が、どれかのマスの上か下の縁になっている。"""
        for c in self.index.near(h.a, h.pos - 2, h.b, h.pos + 2):
            if (abs(c.y0 - h.pos) <= 1.5 or abs(c.y1 - h.pos) <= 1.5) and c.x0 >= h.a - 1.5 and c.x1 <= h.b + 1.5:
                return True
        return False

    def _text_right_of(self, x: float, y0: float, y1: float, limit: float) -> str:
        best = None
        for ln in self.free_lines + [ln for c in self.index.near(x, y0, x + limit, y1) for ln in c.lines]:
            if ln.vertical:
                continue
            if ln.x0 >= x - 1 and ln.x0 - x <= limit and ln.y0 < y1 and ln.y1 > y0:
                if best is None or ln.x0 < best.x0:
                    best = ln
        if best is None:
            return ""
        t = best.text.strip()
        m = re.search(r"[　 ]{2,}", t)
        return norm_label(t[:m.start()] if m else t)

    def _label_for(self, first: Cell, last: Cell) -> tuple[str, str, str]:
        """空のマスの見出し：左の文字のあるマス（行の見出し）と、上の文字のあるマス（列の見出し）。"""
        left = None
        for c in self.index.near(first.x0 - 400, first.y0, first.x0, first.y1):
            if c.x1 <= first.x0 + 1.5 and c.visible() and _vov(c, first) >= 0.5:
                if left is None or c.x1 > left.x1:
                    left = c
        up = None
        width = last.x1 - first.x0
        for c in self.index.near(first.x0, first.y0 - 300, last.x1, first.y0):
            inter = min(c.x1, last.x1) - max(c.x0, first.x0)
            if c.y1 <= first.y0 + 1.5 and c.visible() and inter >= 0.6 * width and inter >= 0.5 * c.w \
                    and len(c.text()) <= 40:
                if up is None or c.y1 > up.y1:
                    up = c
        right = None
        for c in self.index.near(last.x1, first.y0, last.x1 + 400, first.y1):
            if c.x0 >= last.x1 - 1.5 and c.visible() and _vov(c, first) >= 0.5:
                if right is None or c.x0 < right.x0:
                    right = c
        if right is not None and _rtl_cell(right) and (left is None or not _rtl_cell(left)):
            left = right  # 右から左へ書く表：行の見出しは右のマス
        rl = norm_label(left.text()) if left else ""
        cl = norm_label(up.text()) if up else ""
        if up is not None and any(sp.cell is up and sp.kind in ("inline", "text") for sp in self.spots):
            cl = ""  # 上のマスは、それ自体が入力欄を含む（見出しではない）
        if rl and cl and rl != cl:
            # 左（右から左の表では右）がすぐ隣の見出しで、上の見出しが遠い（表の列の見出しでない）ときは、左だけ
            if min(abs(first.x0 - left.x1), abs(left.x0 - last.x1)) <= 1.5 and first.y0 - up.y1 > 3 * max(self.typical, 8):
                return rl, rl, cl
            return f"{rl}／{cl}", rl, cl
        return rl or cl, rl, cl


def _rtl_cell(c: Cell) -> bool:
    lines = [ln for ln in c.lines if not ln.vertical and ln.text.strip()]
    return bool(lines) and all(getattr(ln, "rtl", False) for ln in lines)


def _own_edges(c: Cell, slack: float = 2.5) -> bool:
    e = getattr(c, "edges", None)
    if not e:
        return True
    lv, rv, th, bh = e
    vert = lv.a >= c.y0 - slack and lv.b <= c.y1 + slack and rv.a >= c.y0 - slack and rv.b <= c.y1 + slack
    hor = th.a >= c.x0 - slack and th.b <= c.x1 + slack and bh.a >= c.x0 - slack and bh.b <= c.x1 + slack
    return vert or hor


def _vov(a: Cell, b: Cell) -> float:
    inter = min(a.y1, b.y1) - max(a.y0, b.y0)
    return inter / max(1e-6, min(a.h, b.h))


def _hov(a: Cell, first: Cell, last: Cell) -> float:
    inter = min(a.x1, last.x1) - max(a.x0, first.x0)
    return inter / max(1e-6, min(a.w, last.x1 - first.x0))


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
    if not chars:
        return 10.0
    s = sorted(c.size for c in chars)
    return s[len(s) // 2]


def _fit_size(h: float, typical: float) -> float:
    return max(6.0, min(h * 0.62, max(typical, 10.5), 14.0))


def _holds_text(k: Spot) -> bool:
    """字を書く所に、書き込みと同じ決め方（いちばん小さな字の大きさで）で1字が入るか。
    入らない所（小さな字の図の中の空欄など）も欄としては示し、書けない理由を添える（呼ぶ側）。"""
    if k.kind not in ("text", "inline", "space"):
        return True
    from .fill import spot_fits

    return spot_fits("1", [round(v, 1) for v in k.rect], round(float(k.size or 10.5), 2), k.kind)


_KIND_JA = {"text": "文字の欄", "comb": "1字ずつ書く欄", "inline": "文字の中の空欄", "checkbox": "チェックの印",
            "space": "余白", "circle": "○で囲む選択肢"}


def _ja_kind(k: Spot) -> str:
    return _KIND_JA.get(k.kind, k.kind)


def _blank_hint(text: str, at: tuple[int, int]) -> str:
    """語を手がかりにした空欄の、前後の字（どの語で見つけたかを、エージェントが確かめるため）。"""
    before = text[:at[0]].rstrip(" 　")[-4:]
    after = text[at[1]:].lstrip(" 　")[:4]
    return "・".join(x for x in ((f"前「{before}」" if before else ""), (f"後「{after}」" if after else "")) if x) or "行の空き"
