"""PDF：記入例と白紙の様式を照らして、入力欄の見落としがないかを確かめる。

記入例のページの字のうち、白紙の同じ位置にない字（書き込まれた値）をまとまりにし、
それが白紙で見つけた入力欄（空のマス・文字の中の空欄・下線・桁の枠・□・○で囲む選択など）の中にあるかを調べる。
入力欄の外にあるまとまりは「見落とした入力欄」の候補として、ページ・位置・まわりの文字と一緒に返す。
記入例で描き足された線（○で囲んだ印・チェックの印）も、入力欄のそばにあるかを調べる。

記入例の「記入例」の印や、「〜してください」の説明の文は、書き込まれた値ではないので数えない（別に数える）。

評価用の目安：この照らし合わせは、とどくの見つけ方を開発のときに測る道具（cli の example・xlsx/audit の記入例の検査）で、
エージェントに渡す窓・書き込みには使わない。説明の文・記入例の印を見分ける語の表（_is_note・_notes の「記入例」「ください」
など）と、様式の版が同じかの割合（same_version の 65%）は、測るための目安で、書類の決まり（仕様）から決めたものではない。
"""

from __future__ import annotations

import difflib
import math
import re
from dataclasses import dataclass, field

from .content import page_content
from .doc import PdfDocument
from .layout import Ch, PageLayout

MARKS = set("✓✔☑☒■●○〇◯レ×✕✗✘◉")


@dataclass
class Run:
    page: int
    text: str
    rect: tuple[float, float, float, float]
    size: float
    kind: str = "text"  # text / mark / drawing


@dataclass
class Missed:
    page: int
    text: str
    rect: list[float]
    near: str
    kind: str


@dataclass
class ExampleCheck:
    pages: int = 0
    added: int = 0
    covered: int = 0  # 規則で見つけた入力欄、または規則では決められない所（エージェントの判断）の中
    covered_agent: int = 0  # そのうち、規則では決められない所（［p1.a1？］）の中だけにあるもの
    notes: int = 0  # 記入例の印・説明の文
    missed: list[Missed] = field(default_factory=list)
    changed_template: int = 0  # 白紙にあって記入例で消えた字（版の違い・印の置き換え）
    template: int = 0  # 記入例にだけある字のうち、様式の字の並び（版の違いで動いたもの）
    excluded: list = field(default_factory=list)  # 数えなかったまとまり（理由・ページ・字・位置）
    blank_chars: int = 0  # 照らしたページの、白紙の字の数
    blank_text_chars: int = 0  # 白紙のすべてのページの字の数（0 なら、白紙は画像だけ）
    other_pages: list = field(default_factory=list)  # 白紙のどのページとも様式が合わない記入例のページ（説明のページなど）
    page_kept: list = field(default_factory=list)  # 照らしたページごとの、白紙の字が記入例の同じ位置に残った割合

    @property
    def same_version(self) -> bool:
        """照らしたどのページも、白紙の字の 65% 以上が記入例の同じ位置にある（様式の版が同じ）。65% は評価用の目安。"""
        return bool(self.page_kept) and min(k for _p, k in self.page_kept) >= 0.65

    @property
    def ok(self) -> bool:
        return not self.missed

    def headline(self) -> str:
        s = (f"記入例で書き込まれた所 {self.added} か所のうち {self.covered - self.covered_agent} か所が、規則で見つけた入力欄の中に、"
             f"{self.covered_agent} か所が、規則では決められない所（エージェントの判断）の中にあります"
             + (f"（記入例の印・説明の文 {self.notes} か所、様式の版の違いで動いた字 {self.template} か所は除外）"
                if self.notes or self.template else ""))
        if self.missed:
            s += f"。入力欄の外：{len(self.missed)} か所（見落としの候補）"
        if self.blank_chars and self.changed_template > 0.1 * self.blank_chars:
            s += f"。白紙と記入例で様式の字が {self.changed_template} 字違います（版が違うかもしれません）"
        if self.other_pages:
            s += f"。記入例の {len(self.other_pages)} ページは白紙のどのページとも様式が合わないので照らしていません（説明のページなど）"
        return s


def _match(blank: list[Ch], ex: list[Ch]) -> tuple[list[Ch], int]:
    """記入例の字のうち、白紙の同じ位置に同じ字がないもの。と、白紙で対応のない字の数。"""
    grid: dict[tuple, list[Ch]] = {}
    for c in blank:
        grid.setdefault((c.text, round(c.cx / 3), round(c.cy / 3)), []).append(c)
    used = set()
    added = []
    for c in ex:
        hit = None
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for b in grid.get((c.text, round(c.cx / 3) + dx, round(c.cy / 3) + dy), ()):
                    if id(b) in used:
                        continue
                    tol = max(1.5, 0.35 * max(b.size, c.size))
                    if abs(b.cx - c.cx) <= tol and abs(b.cy - c.cy) <= tol:
                        hit = b
                        break
                if hit:
                    break
            if hit:
                break
        if hit is None:
            added.append(c)
        else:
            used.add(id(hit))
    return added, sum(1 for b in blank if id(b) not in used and not b.text.isspace())


def _runs(chars: list[Ch], page: int) -> list[Run]:
    chars = [c for c in chars if not c.text.isspace()]
    hor = sorted([c for c in chars if not c.vertical], key=lambda c: (round(c.base), c.x0))
    ver = sorted([c for c in chars if c.vertical], key=lambda c: (round(c.base), c.y0))
    out: list[Run] = []
    for group, vertical in ((hor, False), (ver, True)):
        cur: list[Ch] = []
        for c in group:
            if cur:
                last = cur[-1]
                sz = max(last.size, c.size)
                same_line = abs(c.base - last.base) <= 0.5 * sz
                gap = (c.y0 - last.y1) if vertical else (c.x0 - last.x1)
                if same_line and -0.5 * sz <= gap <= 1.2 * sz:
                    cur.append(c)
                    continue
                out.append(_run(cur, page))
            cur = [c]
        if cur:
            out.append(_run(cur, page))
    return out


def _run(cs: list[Ch], page: int) -> Run:
    text = "".join(c.text for c in cs)
    r = Run(page, text, (min(c.x0 for c in cs), min(c.y0 for c in cs), max(c.x1 for c in cs), max(c.y1 for c in cs)),
            max(c.size for c in cs))
    if all(ch in MARKS for ch in text):
        r.kind = "mark"
    return r


def _is_note(r: Run) -> bool:
    """記入例の説明の文・印か（評価用の目安：語の表で見分ける。窓・書き込みには使わない）。"""
    t = r.text
    return any(w in t for w in ("記入例", "記載例", "作成例", "見本", "（例）", "(例)")) or \
        ("ください" in t and len(t) >= 10) or (t.endswith("。") and len(t) >= 20) or \
        bool(re.match(r"[→⇒←↓↑※＊]", t)) and (len(t) >= 6 or all(ch in "→⇒←↓↑※＊" for ch in t))


def _area(r) -> float:
    return max(0.0, r[2] - r[0]) * max(0.0, r[3] - r[1])


def _inter(a, b) -> float:
    x0, y0, x1, y1 = max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])
    return _area((x0, y0, x1, y1)) if x1 > x0 and y1 > y0 else 0.0


def _covered(r: Run, spots) -> bool:
    rx = ((r.rect[0] + r.rect[2]) / 2, (r.rect[1] + r.rect[3]) / 2)
    for sp in spots:
        x0, y0, x1, y1 = sp.rect
        pad = 2.0
        if sp.kind == "inline":
            pad = max(2.0, 0.6 * sp.size)  # 文字の中の空欄は、書いた字が空きより少しはみ出すことがある
        box = (x0 - pad, y0 - pad, x1 + pad, y1 + pad)
        if r.kind == "mark" or r.kind == "drawing":
            if sp.kind in ("checkbox", "circle"):
                cand = [sp.rect] + [rr for _o, rr in getattr(sp, "options", [])]
                for c in cand:
                    cx, cy = (c[0] + c[2]) / 2, (c[1] + c[3]) / 2
                    reach = max(c[2] - c[0], c[3] - c[1], 10.0)
                    if abs(cx - rx[0]) <= reach and abs(cy - rx[1]) <= reach:
                        return True
            if box[0] <= rx[0] <= box[2] and box[1] <= rx[1] <= box[3]:
                return True
            continue
        if box[0] <= rx[0] <= box[2] and box[1] <= rx[1] <= box[3]:
            return True
        if _area(r.rect) and _inter(r.rect, box) >= 0.5 * _area(r.rect):
            return True
    return False


def _near(lay: PageLayout, rect) -> str:
    """書き込まれた所のまわりの文字（同じ高さの左、なければ上）。"""
    cy = (rect[1] + rect[3]) / 2
    best = None
    for ln in lay.free_lines + [ln for c in lay.cells for ln in c.lines]:
        if ln.vertical:
            continue
        if ln.y0 - 2 <= cy <= ln.y1 + 2 and ln.x1 <= rect[0] + 2:
            if best is None or ln.x1 > best.x1:
                best = ln
    if best is None:
        for ln in lay.free_lines + [ln for c in lay.cells for ln in c.lines]:
            if ln.y1 <= rect[1] + 1 and ln.x0 < rect[2] and ln.x1 > rect[0]:
                if best is None or ln.y1 > best.y1:
                    best = ln
    return best.text.strip()[:60] if best is not None else ""


def _all_words(lay: PageLayout) -> dict[str, list]:
    """ページの語（空白で区切った3字以上のまとまり）：語 → 位置の一覧（上から順）。
    全角・半角の違い（２と2）は同じとみなす（目印にするだけで、字そのものは変えない）。"""
    import unicodedata

    found: dict[str, list] = {}
    for ln in lay.free_lines + [x for c in lay.cells for x in c.lines]:
        if ln.vertical:
            continue
        for m in re.finditer(r"[^\s\u3000]{3,}", ln.text):
            key = unicodedata.normalize("NFKC", m.group(0))
            found.setdefault(key, []).append(ln.span_rect(m.start(), m.end()))
    for v in found.values():
        v.sort(key=lambda r: (round(r[1]), r[0]))
    return found


def _words(lay: PageLayout) -> dict[str, tuple]:
    """1回だけ出てくる語：語 → 位置。"""
    return {k: v[0] for k, v in _all_words(lay).items() if len(v) == 1}


def _align(blay: PageLayout, elay: PageLayout):
    """記入例のページの座標 → 白紙のページの座標（ずれと縮み）。両方に1回だけ出てくる語を目印にする。
    目印が足りなければ None（そのまま重ねる）。"""
    wb, we = _words(blay), _words(elay)
    pairs = [(wb[k], we[k]) for k in wb.keys() & we.keys()]
    if len(pairs) < 4:
        return None
    best = None
    for _round in range(4):
        xs = [(e[0], b[0]) for b, e in pairs] + [(e[2], b[2]) for b, e in pairs]
        ys = [(e[1], b[1]) for b, e in pairs] + [(e[3], b[3]) for b, e in pairs]
        fit = []
        for pts in (xs, ys):
            n = len(pts)
            mx = sum(q[0] for q in pts) / n
            my = sum(q[1] for q in pts) / n
            vx = sum((q[0] - mx) ** 2 for q in pts)
            a = sum((q[0] - mx) * (q[1] - my) for q in pts) / vx if vx > 1e-9 else 1.0
            fit.append((a, my - a * mx))
        (ax, bx), (ay, by) = fit
        err = sorted(((abs(ax * e[0] + bx - b[0]) + abs(ay * e[1] + by - b[1]), b, e) for b, e in pairs), key=lambda x: x[0])
        med = err[len(err) // 2][0]
        keep = [(b, e) for d, b, e in err if d <= max(3.0, med * 3)]
        best = (ax, bx, ay, by, med)
        if len(keep) == len(pairs) or len(keep) < 4:
            break
        pairs = keep
    ax, bx, ay, by, med = best
    if med > 6 or not (0.3 < ax < 3 and 0.3 < ay < 3):
        return None
    return ax, bx, ay, by


class _Local:
    """ページの中の一部だけがずれた記入例（行が1つ増えた表など）のための、近くの目印によるずれの直し。
    全体の合わせ（_align）の後に残るずれを、近くの目印の語（両方に1回だけ出てくる語）から求める。"""

    R = 120.0  # 目印を探す範囲（pt）

    def __init__(self, blay: PageLayout, elay: PageLayout, t):
        # 両方に同じ回数（4回まで）出てくる語を、上からの順に組にする（表の見出しのくり返しも目印になる）
        wb, we = _all_words(blay), _all_words(elay)
        self.pts = []
        pairs = []
        for k in wb.keys() & we.keys():
            if len(wb[k]) == len(we[k]) <= 4:
                pairs += list(zip(wb[k], we[k]))
        for b, e in pairs:
            if t:
                ax, bx, ay, by = t
                e = (ax * e[0] + bx, ay * e[1] + by, ax * e[2] + bx, ay * e[3] + by)
            cx, cy = (e[0] + e[2]) / 2, (e[1] + e[3]) / 2
            self.pts.append((cx, cy, (b[0] + b[2]) / 2 - cx, (b[1] + b[3]) / 2 - cy))
        self.grid: dict[tuple[int, int], list] = {}
        for q in self.pts:
            self.grid.setdefault((int(q[0] // self.R), int(q[1] // self.R)), []).append(q)
        self.cache: dict[tuple[int, int], tuple[float, float]] = {}

    def shift(self, x: float, y: float) -> tuple[float, float]:
        key = (int(x // 10), int(y // 10))
        if key in self.cache:
            return self.cache[key]
        gx, gy = int(x // self.R), int(y // self.R)
        near = []
        for i in (-1, 0, 1):
            for j in (-1, 0, 1):
                for q in self.grid.get((gx + i, gy + j), ()):
                    d = math.hypot(q[0] - x, q[1] - y)
                    if d <= self.R:
                        near.append((d, q[2], q[3]))
        near.sort()
        near = near[:3]
        out = (0.0, 0.0)
        if len(near) >= 2:
            dxs = sorted(q[1] for q in near)
            dys = sorted(q[2] for q in near)
            mx, my = dxs[len(dxs) // 2], dys[len(dys) // 2]
            # 近くの目印が同じだけずれているときだけ直す（ばらばらなら、ずれではなく様式の違い）
            if all(abs(q[1] - mx) <= 2.5 and abs(q[2] - my) <= 2.5 for q in near) and math.hypot(mx, my) > 2.0:
                out = (mx, my)
        self.cache[key] = out
        return out

    def ch(self, c: Ch) -> Ch:
        dx, dy = self.shift(c.cx, c.cy)
        if not dx and not dy:
            return c
        return Ch(c.text, c.x0 + dx, c.y0 + dy, c.x1 + dx, c.y1 + dy, c.size, c.vertical,
                  c.base + (dx if c.vertical else dy), c.invisible)

    def rect(self, r):
        dx, dy = self.shift((r[0] + r[2]) / 2, (r[1] + r[3]) / 2)
        return (r[0] + dx, r[1] + dy, r[2] + dx, r[3] + dy)


def _moved(c: Ch, t) -> Ch:
    ax, bx, ay, by = t
    return Ch(c.text, ax * c.x0 + bx, ay * c.y0 + by, ax * c.x1 + bx, ay * c.y1 + by, c.size * (ax + ay) / 2,
              c.vertical, (ay * c.base + by) if not c.vertical else (ax * c.base + bx), c.invisible)


def _half(lay: PageLayout, x0: float, x1: float, shift: float) -> PageLayout:
    """見開きのページの半分を、1ページとして扱う（字と線を x0〜x1 に絞り、左へ shift だけ寄せる）。"""
    import copy

    h = copy.copy(lay)
    h.chars = [Ch(c.text, c.x0 - shift, c.y0, c.x1 - shift, c.y1, c.size, c.vertical, c.base - (shift if c.vertical else 0),
                  c.invisible) for c in lay.chars if x0 <= c.cx < x1]
    every = lay.free_lines + [ln for c in lay.cells for ln in c.lines]
    h.free_lines = [_shifted_line(ln, shift) for ln in every if x0 <= (ln.x0 + ln.x1) / 2 < x1]
    h.cells = []
    h._shift = shift
    h._range = (x0, x1)
    return h


def _shifted_line(ln, shift):
    import copy

    n = copy.copy(ln)
    n.x0 -= shift
    n.x1 -= shift
    return n


def _example_pages(e: PdfDocument, blank_width: float) -> list[tuple[PageLayout, object]]:
    """記入例のページ（見開きは左右の半分に分ける）：(ページの読み取り, ページの描画)。"""
    out = []
    for i, p in enumerate(e.pages()):
        pc = page_content(e, p)
        lay = PageLayout(e, p, pc, i + 1)
        if lay.view.width >= 1.8 * blank_width:
            half = lay.view.width / 2
            for k in range(2):
                lay.cells = lay.cells  # 表のマスは使わない（半分にすると切れるため）
                out.append((_half(lay, k * half, (k + 1) * half, k * half), pc))
        else:
            out.append((lay, pc))
    return out


def check(blank: bytes, example: bytes) -> ExampleCheck:
    b = PdfDocument(blank, "blank.pdf")
    e = PdfDocument(example, "example.pdf")
    out = ExampleCheck()
    pb = b.pages()
    blays = [PageLayout(b, p, page_content(b, p), i + 1) for i, p in enumerate(pb)]
    out.blank_text_chars = sum(1 for lay in blays for c in lay.chars if not c.text.isspace())
    ex_pages = _example_pages(e, blays[0].view.width if blays else 595)
    out.pages = len(ex_pages)
    for elay, pce in ex_pages:
        # 記入例のこのページが、白紙のどのページに当たるか（目印の行がいちばん多く合うページ）
        best = None
        for lay in blays:
            t = _align(lay, elay)
            if t is None:
                continue
            n = _anchors(lay, elay)
            if best is None or n > best[0]:
                best = (n, lay, t)
        if best is None:
            # 目印が合わない：ページの順で重ねる
            k = min(ex_pages.index((elay, pce)), len(blays) - 1)
            best = (0, blays[k], None)
        _n, lay, t = best
        i = lay.number - 1
        local = _Local(lay, elay, t)
        # 近くの目印によるずれの直しは、行ごとに同じだけ（字ごとに直すと、1つの行が切れてしまう）
        line_of: dict[int, object] = {}
        for ln in elay.free_lines + [x for c in elay.cells for x in c.lines]:
            for c in getattr(ln, "chars", ()):
                line_of[id(c)] = ln
        shift_of: dict[int, tuple[float, float]] = {}
        ex_chars = []
        for c in elay.chars:
            m = _moved(c, t) if t else c
            ln = line_of.get(id(c))
            if ln is not None:
                if id(ln) not in shift_of:
                    r = (ln.x0, ln.y0, ln.x1, ln.y1)
                    if t:
                        ax, bx, ay, by = t
                        r = (ax * r[0] + bx, ay * r[1] + by, ax * r[2] + bx, ay * r[3] + by)
                    shift_of[id(ln)] = local.shift((r[0] + r[2]) / 2, (r[1] + r[3]) / 2)
                dx, dy = shift_of[id(ln)]
                if dx or dy:
                    m = Ch(m.text, m.x0 + dx, m.y0 + dy, m.x1 + dx, m.y1 + dy, m.size, m.vertical,
                           m.base + (dx if m.vertical else dy), m.invisible)
            else:
                m = local.ch(m)
            ex_chars.append(m)
        added, removed = _match(lay.chars, ex_chars)
        page_chars = sum(1 for c in lay.chars if not c.text.isspace())
        if page_chars and removed > 0.75 * page_chars:
            # 白紙のページの字の4分の3より多くが記入例にない：同じ様式のページではない（記載要領・説明のページ）。
            # 版が違うだけの記入例（見出しや行が少し違う）は、字の3〜4割は同じ位置に残る
            out.other_pages.append(elay.number)
            continue
        out.changed_template += removed
        out.blank_chars += page_chars
        out.page_kept.append((elay.number, round(1 - removed / max(1, page_chars), 2)))
        runs = _runs(added, i + 1)
        # 記入例で描き足された線・図形（○で囲む印・チェックの印・吹き出し）
        shift = getattr(elay, "_shift", 0.0)
        rng = getattr(elay, "_range", None)
        blank_curves = [lay.view.rect(*c[:4]) for c in lay.pc.curves]
        callouts = []
        drawn = []
        for cv in pce.curves:
            r = elay.view.rect(*cv[:4])
            if rng and not (rng[0] <= (r[0] + r[2]) / 2 < rng[1]):
                continue
            r = (r[0] - shift, r[1], r[2] - shift, r[3])
            if t:
                ax, bx, ay, by = t
                r = (ax * r[0] + bx, ay * r[1] + by, ax * r[2] + bx, ay * r[3] + by)
            r = local.rect(r)
            if any(abs(r[0] - q[0]) < 2 and abs(r[1] - q[1]) < 2 and abs(r[2] - q[2]) < 2 and abs(r[3] - q[3]) < 2 for q in blank_curves):
                continue
            if (r[2] - r[0]) > 120 or (r[3] - r[1]) > 60:
                if _is_callout(r, lay):
                    callouts.append(r)  # 大きな図形：吹き出しなど
            elif 3 <= r[2] - r[0] and 3 <= r[3] - r[1]:
                drawn.append(r)
        for bx_ in pce.boxes:
            r = elay.view.rect(bx_.x0, bx_.y0, bx_.x1, bx_.y1)
            if rng and not (rng[0] <= (r[0] + r[2]) / 2 < rng[1]):
                continue
            r = (r[0] - shift, r[1], r[2] - shift, r[3])
            if t:
                ax, bx, ay, by = t
                r = (ax * r[0] + bx, ay * r[1] + by, ax * r[2] + bx, ay * r[3] + by)
            r = local.rect(r)
            if bx_.fill is not None and not all(v > 0.97 for v in bx_.fill[:3]) or bx_.stroke:
                if not any(_inter(r, f[:4]) > 0.9 * _area(r) for f in lay.fills) and _is_callout(r, lay):
                    callouts.append(r)
        for r_ in _marks_only(drawn):
            runs.append(Run(i + 1, "（描いた印）", r_, 10.0, "drawing"))
        pw, ph = lay.view.width, lay.view.height
        blank_lines = lay.free_lines + [x for c in lay.cells for x in c.lines]
        blank_text = "".join("".join(ln.text.split()) for ln in blank_lines)
        note_ids = _notes(runs)
        for r in runs:
            if r.kind == "text" and id(r) not in note_ids and _template_text(r, blank_text, lay, blank_lines):
                out.template += 1
                out.excluded.append(("様式の字", r.page, r.text[:30], [round(v) for v in r.rect]))
                continue
            cx, cy = (r.rect[0] + r.rect[2]) / 2, (r.rect[1] + r.rect[3]) / 2
            outside = not (0 <= cx <= pw and 0 <= cy <= ph)
            if id(r) in note_ids or outside:
                out.notes += 1
                why = "ページの外" if outside else "説明の文"
                out.excluded.append((why, r.page, r.text[:30], [round(v) for v in r.rect]))
                continue
            out.added += 1
            if _covered(r, [k for k in lay.spots if k.confidence != "agent"]):
                out.covered += 1
            elif _covered(r, [k for k in lay.spots if k.confidence == "agent"]):
                out.covered += 1
                out.covered_agent += 1
            else:
                out.missed.append(Missed(i + 1, r.text[:40], [round(v, 1) for v in r.rect], _near(lay, r.rect), r.kind))
    return out


def _marks_only(drawn: list) -> list:
    """描き足された線のうち、印（○・チェック）だけ。字の輪郭を線で描いたもの（同じ高さに小さな図形が並ぶ）は除く。"""
    out = []
    for r in drawn:
        h = r[3] - r[1]
        w = r[2] - r[0]
        if min(w, h) < 6.5 or w < 8:
            continue  # 字の画の一部ほどの小さな線（○で囲む印・チェックの印は、もっと大きい）
        row = [q for q in drawn if q is not r and abs((q[1] + q[3]) / 2 - (r[1] + r[3]) / 2) <= 0.3 * max(h, 4)
               and abs(q[0] - r[0]) <= 4 * max(h, 4) and abs((q[3] - q[1]) - h) <= 0.6 * max(h, 4)]
        if len(row) >= 2 and h < 14:
            continue  # 字の輪郭
        out.append(r)
    return out


def _is_callout(r, lay: PageLayout) -> bool:
    """記入例だけにある大きな図形のうち、吹き出し（白紙のマスと重ならず、ページの3割より小さい）。"""
    if _area(r) > 0.3 * lay.view.width * lay.view.height:
        return False
    for c in lay.cells:
        cr = (c.x0, c.y0, c.x1, c.y1)
        if abs(cr[0] - r[0]) < 4 and abs(cr[1] - r[1]) < 4 and abs(cr[2] - r[2]) < 4 and abs(cr[3] - r[3]) < 4:
            return False  # 白紙のマスと同じ形（表の描き方が違うだけ）
    return True


def _notes(runs: list[Run]) -> set[int]:
    """記入例の説明（吹き出しの文など）：縦に続く段落ごとに、説明の言い回しがあれば段落ぜんぶを説明とする。"""
    texts = [r for r in runs if r.kind == "text"]
    texts.sort(key=lambda r: (round(r.rect[0]), r.rect[1]))
    parent = {id(r): id(r) for r in texts}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for a in texts:
        for b in texts:
            if a is b:
                continue
            gap = b.rect[1] - a.rect[3]
            if abs(a.rect[0] - b.rect[0]) <= max(3.0, 0.4 * a.size) and -0.3 * a.size <= gap <= 0.8 * max(a.size, b.size):
                parent[find(id(a))] = find(id(b))
            # 同じ行の続き（吹き出しの文が、途中で字の書式が変わって切れたもの）
            same_line = abs((a.rect[1] + a.rect[3]) / 2 - (b.rect[1] + b.rect[3]) / 2) <= 0.4 * max(a.size, b.size)
            if same_line and 0 <= b.rect[0] - a.rect[2] <= 1.5 * max(a.size, b.size):
                parent[find(id(a))] = find(id(b))
    groups: dict[int, list[Run]] = {}
    for r in texts:
        groups.setdefault(find(id(r)), []).append(r)
    out = set()
    for g in groups.values():
        whole = "".join(x.text for x in g)
        if _is_note(Run(0, whole, (0, 0, 0, 0), 10)) or (len(g) >= 2 and re.search(r"ください|場合|必要|不要|できません|こちら|。", whole)) \
                or whole.strip() in ("※", "＊"):
            out.update(id(x) for x in g)
    return out


def _template_text(r: Run, blank_text: str, lay: PageLayout, blank_lines=()) -> bool:
    t = "".join(r.text.split())
    if len(t) >= 4 and t in blank_text:
        return True
    if len(t) >= 8:
        # 様式の版の違い：近くの白紙の行と、ほぼ同じ文
        cy = (r.rect[1] + r.rect[3]) / 2
        for ln in blank_lines:
            if abs((ln.y0 + ln.y1) / 2 - cy) <= 20:
                u = "".join(ln.text.split())
                if len(u) >= 6:
                    sm = difflib.SequenceMatcher(None, t, u, autojunk=False)
                    matched = sum(b.size for b in sm.get_matching_blocks())
                    if matched >= 0.7 * len(t):
                        return True  # 白紙の行の一部と、ほぼ同じ文
    if len(t) <= 3:
        cx, cy = (r.rect[0] + r.rect[2]) / 2, (r.rect[1] + r.rect[3]) / 2
        for c in lay.chars:
            if c.text in t and abs(c.cx - cx) <= 40 and abs(c.cy - cy) <= 40 and not c.text.isspace():
                # 近くに同じ字がある：様式の字が少し動いたもの（数字の値は除く）
                if not any(ch.isdigit() for ch in t):
                    return True
    return False


def _anchors(blay: PageLayout, elay: PageLayout) -> int:
    return len(_words(blay).keys() & _words(elay).keys())
