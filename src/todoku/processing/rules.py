"""形ごとの長方形を、規則（1 つのまとまりに 1 つ）にする。規則が読む範囲・規則のあいだの依存・順番・循環。"""

from __future__ import annotations

import bisect
from dataclasses import dataclass, field

from ..formula.parser import Binary, Func, MAX_COL, MAX_ROW, NameNode, Node, RefNode, StructRef, addr, iter_refs, parse, ref_text
from ..formula.values import Rng
from .scan import Rect, SheetScan
from .shape import Shape, Shaper, pattern, render

CELL_SEARCH_LIMIT = 2_000_000  # 循環を確かめるためにセルをたどる数の上限（超えたら「確かめられない」と言う）


def rect_text(r1: int, c1: int, r2: int, c2: int) -> str:
    a = addr(r1, c1)
    return a if (r1, c1) == (r2, c2) else f"{a}:{addr(r2, c2)}"


def key_text(sheet: str, r1: int, c1: int, r2: int | None = None, c2: int | None = None) -> str:
    return f"{sheet}!{rect_text(r1, c1, r1 if r2 is None else r2, c1 if c2 is None else c2)}"


@dataclass
class Read:
    """規則が読む 1 つの参照。"""

    node: Node  # RefNode か StructRef（origin から見た書き方）
    sheet: str  # 実際のシートの名前（外部は "[1]Sheet1"）
    origin: tuple[int, int]  # node を書いたセル
    table: tuple | None = None  # (Calc, シート) 構造化参照のとき
    name: str | None = None  # 名前を通して読むとき、その名前
    absolute: bool = False  # 規則のどのセルからも同じ所を読む
    external: bool = False
    missing: bool = False  # ないシート（#REF!）
    pattern: str = ""
    bounds: tuple[int, int] = (0, 0)
    lookup: set = field(default_factory=set)  # 表として読む関数（VLOOKUP・INDEX など）
    fixed: bool = False  # 書かれた場所をそのまま読む（名前の中の参照・配列数式は、セルによってずれない）

    def at(self, r: int, c: int) -> tuple[int, int, int, int] | None:
        """(r, c) のセルが読む長方形。"""
        n = self.node
        if isinstance(n, StructRef):
            calc, sheet = self.table
            got = calc.table_ref(n, sheet, r, c)
            if not isinstance(got, Rng):
                return None
            return got.r1, got.c1, got.r2, got.c2
        dr, dc = r - self.origin[0], c - self.origin[1]
        s = n if self.fixed else n.shifted(dr, dc)
        r1, r2 = (s.r1, s.r2) if s.r1 <= s.r2 else (s.r2, s.r1)
        c1, c2 = (s.c1, s.c2) if s.c1 <= s.c2 else (s.c2, s.c1)
        if r2 < 1 or c2 < 1 or r1 > MAX_ROW or c1 > MAX_COL:
            return None
        mr, mc = self.bounds
        if s.kind == "cols":
            r1, r2 = 1, max(mr, 1)
        elif s.kind == "rows":
            c1, c2 = 1, max(mc, 1)
        return max(r1, 1), max(c1, 1), min(r2, MAX_ROW), min(c2, MAX_COL)


@dataclass
class Rule:
    id: str
    si: int
    sheet: str
    shape: Shape
    kind: str
    members: list[Rect]
    r1: int
    c1: int
    r2: int
    c2: int
    origin: tuple[int, int]
    text: str  # 左上のセルの数式（ファイルのまま）
    count: int
    reads: list[Read] = field(default_factory=list)
    uses: set[str] = field(default_factory=set)  # この規則が読む規則
    used_by: set[str] = field(default_factory=set)
    exceptions: list[dict] = field(default_factory=list)
    exception_count: int = 0
    level: int = 0
    in_cycle: str | None = None
    note: dict | None = None

    @property
    def range(self) -> str:
        return rect_text(self.r1, self.c1, self.r2, self.c2)

    @property
    def ref(self) -> str:
        return f"{self.sheet}!{self.range}"

    def pattern_body(self) -> str:
        if self.shape.ast is None:
            return self.text
        many_r, many_c = self.moves()
        s = self.shape
        return render(s.ast, lambda n: pattern(n, s.r, s.c, many_r, many_c))

    def moves(self) -> tuple[bool, bool]:
        """(行に連動するか, 列に連動するか)。配列数式は 1 つの式なので、どちらも連動しない。"""
        if self.kind in ("array", "dynamic_array"):
            return False, False
        return self.r2 > self.r1, self.c2 > self.c1


def build_blocks(sheets: list[SheetScan], shaper: Shaper) -> list[Rule]:
    """同じ形の長方形を、まとまりにする。あいだが埋まっている（別の数式か定数が入っている）長方形は、同じ規則の例外として
    つなぐ。ただし規則のセルが例外より多いときだけ。空いたマスを挟むなら別のまとまり。"""
    rules: list[Rule] = []
    for si, sh in enumerate(sheets):
        by_key: dict[tuple, list[Rect]] = {}
        blocks: list[list[Rect]] = []
        for rc in sh.rects:
            if rc.kind in ("formula", "array_cell"):
                by_key.setdefault((rc.shape, rc.c1, rc.c2), []).append(rc)
            else:
                blocks.append([rc])
        stage = []
        for (_sid, c1, c2), rects in by_key.items():
            rects.sort(key=lambda x: x.r1)
            stage += _bridge(rects, lambda g, a, b: _gap_filled(sh, g.c1, g.c2, a + 1, b - 1), "r")
        # 同じ行の範囲で横に並ぶものも同じようにつなぐ（縦につないだものは、形が同じ外枠のときだけ）
        by_rows: dict[tuple, list[list[Rect]]] = {}
        for grp in stage:
            r1, r2 = min(x.r1 for x in grp), max(x.r2 for x in grp)
            by_rows.setdefault((grp[0].shape, r1, r2), []).append(grp)
        for (_sid, r1, r2), grps in by_rows.items():
            grps.sort(key=lambda g: g[0].c1)
            cur = grps[0]
            m = _cells(cur)
            gaps = 0
            for g in grps[1:]:
                hi = max(x.c2 for x in cur)
                lo = g[0].c1
                w = lo - hi - 1
                extra = w * (r2 - r1 + 1)
                n = _cells(g)
                if w == 0 or (w > 0 and _gap_filled_cols(sh, r1, r2, hi + 1, lo - 1) and m + n > gaps + extra):
                    cur = cur + g
                    m += n
                    gaps += extra
                else:
                    blocks.append(cur)
                    cur, m, gaps = g, n, 0
            blocks.append(cur)
        for grp in blocks:
            rules.append(_rule(si, sh.name, grp, shaper))
    rules.sort(key=lambda u: (u.si, u.r1, u.c1))
    for i, u in enumerate(rules):
        u.id = f"R{i + 1}"
    return rules


def _cells(grp: list[Rect]) -> int:
    return sum(x.cells for x in grp)


def _gap_filled(sh: SheetScan, c1: int, c2: int, r1: int, r2: int) -> bool:
    return all(sh.occupied.covers(c, r1, r2) for c in range(c1, c2 + 1))


def _gap_filled_cols(sh: SheetScan, r1: int, r2: int, c1: int, c2: int) -> bool:
    return all(sh.occupied.covers(c, r1, r2) for c in range(c1, c2 + 1))


def _bridge(rects: list[Rect], filled, _axis) -> list[list[Rect]]:
    out = []
    cur = [rects[0]]
    m = rects[0].cells
    gaps = 0
    w = rects[0].c2 - rects[0].c1 + 1
    for x in rects[1:]:
        last = cur[-1]
        h = x.r1 - last.r2 - 1
        extra = h * w
        if h <= 0 or (m + x.cells > gaps + extra and filled(last, last.r2, x.r1)):
            cur.append(x)
            m += x.cells
            gaps += max(extra, 0)
        else:
            out.append(cur)
            cur = [x]
            m, gaps = x.cells, 0
    out.append(cur)
    return out


def _rule(si: int, sheet: str, grp: list[Rect], shaper: Shaper) -> Rule:
    grp = sorted(grp, key=lambda x: (x.r1, x.c1))
    first = grp[0]
    shape = shaper.shapes[first.shape]
    r1 = min(x.r1 for x in grp)
    c1 = min(x.c1 for x in grp)
    r2 = max(x.r2 for x in grp)
    c2 = max(x.c2 for x in grp)
    if first.note and first.note.get("master"):
        origin = tuple(first.note["master"])
    else:
        origin = (first.r1, first.c1)
    text = first.text if first.text is not None else shape.text_at(*origin)
    return Rule("", si, sheet, shape, first.kind, grp, r1, c1, r2, c2, origin, text, _cells(grp), note=first.note)


# ---------------------------------------------------------------- 読む範囲
def analyse_reads(rules: list[Rule], sheets: list[SheetScan], wb) -> None:
    from ..xlsx.calc import Calc

    names = {s.name for s in sheets}
    lower = {s.name.lower(): s.name for s in sheets}
    bounds = {s.name: (s.max_row, s.max_col) for s in sheets}
    calc = None

    def real(sheet: str) -> tuple[str, bool, bool]:
        if sheet.startswith("["):
            return sheet, True, False
        if sheet in names:
            return sheet, False, False
        got = lower.get(sheet.lower())
        return (got, False, False) if got else (sheet, False, True)

    for u in rules:
        s = u.shape
        if s.ast is None:
            continue
        many_r, many_c = u.moves()
        is_array = u.kind in ("array", "dynamic_array")
        seen: dict = {}
        lookups = lookup_args(s.ast)

        def add(node, sheet, name, depth, fn=None):
            nonlocal calc
            if isinstance(node, NameNode):
                if depth > 8:
                    return
                d = wb.defined_name(node.name, node.sheet or sheet) if hasattr(wb, "defined_name") else None
                if d is None or d.ast is None:
                    return
                for x in iter_refs(d.ast):
                    add(x, d.sheet_scope or sheet, name or d.name, depth + 1, fn)
                return
            if isinstance(node, StructRef):
                if calc is None:
                    calc = Calc(wb)
                k = ("t", node.table, node.specials, node.col1, node.col2)
                if k not in seen:
                    seen[k] = Read(node, _table_sheet(wb, node, sheet), (s.r, s.c), (calc, sheet),
                                   name=name, absolute="#This Row" not in node.specials, pattern=render(node))
                    u.reads.append(seen[k])
                if fn:
                    seen[k].lookup.add(fn)
                return
            if not isinstance(node, RefNode):
                return
            sh_name, external, missing = real(node.sheet or sheet)
            k = (sh_name, node.r1, node.c1, node.r2, node.c2, node.abs_r1, node.abs_c1, node.abs_r2, node.abs_c2, name)
            if k in seen:
                if fn:
                    seen[k].lookup.add(fn)
                return
            fixed = bool(name) or is_array
            absolute = fixed or ((node.abs_r1 or node.kind == "cols") and (node.abs_r2 or node.kind == "cols")
                                      and (node.abs_c1 or node.kind == "rows") and (node.abs_c2 or node.kind == "rows"))
            pat = ref_text(node) if fixed else pattern(node, s.r, s.c, many_r, many_c)
            seen[k] = Read(node, sh_name, (s.r, s.c), None, name, absolute, external, missing, pat,
                           bounds.get(sh_name, (0, 0)), fixed=fixed)
            u.reads.append(seen[k])
            if fn:
                seen[k].lookup.add(fn)

        for ref in iter_refs(s.ast):
            add(ref, u.sheet, None, 0, lookups.get(id(ref)))
        for extra in literal_indirect(s.ast, u.sheet):  # INDIRECT("Sheet!A1")：字が決まっていれば静的に分かる
            add(extra, u.sheet, None, 0)


# 表の引数の位置（0 から）：表そのものを読む関数
TABLE_ARGS = {"VLOOKUP": (1,), "HLOOKUP": (1,), "LOOKUP": (1, 2), "INDEX": (0,), "MATCH": (1,), "XLOOKUP": (1, 2), "XMATCH": (1,)}


def lookup_args(ast: Node) -> dict[int, str]:
    """表として読まれる参照の節（id） → 関数の名前。"""
    from ..formula.parser import Postfix, Unary

    out: dict[int, str] = {}
    stack = [ast]
    while stack:
        n = stack.pop()
        if isinstance(n, Func):
            for i in TABLE_ARGS.get(n.name, ()):
                if i < len(n.args) and isinstance(n.args[i], (RefNode, NameNode, StructRef)):
                    out[id(n.args[i])] = n.name
            stack.extend(n.args)
        elif isinstance(n, Binary):
            stack.extend((n.left, n.right))
        elif isinstance(n, (Unary, Postfix)):
            stack.append(n.operand)
    return out


def literal_indirect(ast: Node, sheet: str) -> list[RefNode]:
    from ..formula.parser import Str, Unary, Postfix

    out = []
    stack = [ast]
    while stack:
        n = stack.pop()
        if isinstance(n, Func):
            if n.name == "INDIRECT" and n.args and isinstance(n.args[0], Str) and len(n.args) == 1:
                try:
                    got = parse(n.args[0].value)
                except Exception:  # noqa: BLE001
                    got = None
                if isinstance(got, RefNode):
                    got.abs_r1 = got.abs_c1 = got.abs_r2 = got.abs_c2 = True  # INDIRECT の字は、ずれない
                    out.append(got)
            stack.extend(n.args)
        elif isinstance(n, Binary):
            stack.extend((n.left, n.right))
        elif isinstance(n, (Unary, Postfix)):
            stack.append(n.operand)
    return out


def _table_sheet(wb, node: StructRef, sheet: str) -> str:
    t = (getattr(wb, "tables", None) or {}).get((node.table or "").lower())
    return t.sheet if t is not None else sheet


def read_rects(u: Rule, rd: Read) -> list[tuple[int, int, int, int]]:
    """規則のすべてのセルが読む長方形（まとまりごと）。"""
    out = []
    corners = u.members if not rd.absolute else u.members[:1]
    for m in corners:
        a = rd.at(m.r1, m.c1)
        b = rd.at(m.r2, m.c2)
        if a is None and b is None:
            continue
        if a is None:
            a = b
        if b is None:
            b = a
        rect = (min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]))
        if out and out[-1][1] == rect[1] and out[-1][3] == rect[3] and rect[0] <= out[-1][2] + 1:
            out[-1] = (min(out[-1][0], rect[0]), rect[1], max(out[-1][2], rect[2]), rect[3])
        else:
            out.append(rect)
    return out


# ---------------------------------------------------------------- 依存と順番
class RectIndex:
    """シートごとの長方形を、上の行の順に引く。"""

    def __init__(self):
        self.sheets: dict[str, tuple[list, list, list]] = {}
        self._raw: dict[str, list] = {}

    def add(self, sheet: str, r1: int, c1: int, r2: int, c2: int, who) -> None:
        self._raw.setdefault(sheet, []).append((r1, c1, r2, c2, who))

    def freeze(self) -> "RectIndex":
        for s, items in self._raw.items():
            items.sort(key=lambda x: (x[0], x[1]))
            tops = [x[0] for x in items]
            high = []
            m = 0
            for x in items:
                m = max(m, x[2])
                high.append(m)
            self.sheets[s] = (items, tops, high)
        return self

    def hits(self, sheet: str, r1: int, c1: int, r2: int, c2: int):
        got = self.sheets.get(sheet)
        if got is None:
            return
        items, tops, high = got
        i = bisect.bisect_right(tops, r2) - 1
        while i >= 0 and high[i] >= r1:
            x = items[i]
            if x[2] >= r1 and x[1] <= c2 and x[3] >= c1:
                yield x
            i -= 1


def link_rules(rules: list[Rule]) -> RectIndex:
    ix = RectIndex()
    for u in rules:
        for m in u.members:
            ix.add(u.sheet, m.r1, m.c1, m.r2, m.c2, u.id)
    ix.freeze()
    by_id = {u.id: u for u in rules}
    for u in rules:
        for rd in u.reads:
            if rd.external or rd.missing:
                continue
            for rect in read_rects(u, rd):
                for x in ix.hits(rd.sheet, *rect):
                    u.uses.add(x[4])
                    by_id[x[4]].used_by.add(u.id)
    readers = RectIndex()  # 読む所 → 読む規則（マクロが書く所と規則の関係に使う）
    for u in rules:
        for rd in u.reads:
            if not rd.external and not rd.missing:
                for rect in read_rects(u, rd):
                    readers.add(rd.sheet, *rect, u.id)
    ix.readers = readers.freeze()
    return ix


def order_rules(rules: list[Rule], ix: RectIndex) -> dict:
    """規則の順番（読むものが先）と循環。規則のあいだで依存が輪になっても、セルでは輪でない（累計など）ことが多いので、
    輪の中はセルの並びまで確かめる。"""
    by_id = {u.id: u for u in rules}
    sccs = _sccs(rules, by_id)
    order: list[str] = []
    cycles: list[dict] = []
    for comp in sccs:  # 読むものが先の順（Tarjan の出る順）
        comp_sorted = sorted(comp, key=lambda x: int(x[1:]))
        order += comp_sorted
        if len(comp) > 1 or comp[0] in by_id[comp[0]].uses:
            cycles.append(_classify(comp_sorted, by_id, ix))
    level: dict[str, int] = {}
    for comp in sccs:
        inside = set(comp)
        base = 0
        for x in comp:
            for y in by_id[x].uses:
                if y not in inside:
                    base = max(base, level[y] + 1)
        for x in comp:
            level[x] = base
            by_id[x].level = base
    return {"order": order, "cycles": cycles}


def _sccs(rules: list[Rule], by_id: dict) -> list[list[str]]:
    index: dict[str, int] = {}
    low: dict[str, int] = {}
    on: set[str] = set()
    stack: list[str] = []
    out: list[list[str]] = []
    n = 0
    for root in rules:
        if root.id in index:
            continue
        work = [(root.id, iter(sorted(by_id[root.id].uses, key=lambda x: int(x[1:]))))]
        index[root.id] = low[root.id] = n
        n += 1
        stack.append(root.id)
        on.add(root.id)
        while work:
            v, it = work[-1]
            for w in it:
                if w not in index:
                    index[w] = low[w] = n
                    n += 1
                    stack.append(w)
                    on.add(w)
                    work.append((w, iter(sorted(by_id[w].uses, key=lambda x: int(x[1:])))))
                    break
                if w in on:
                    low[v] = min(low[v], index[w])
            else:
                work.pop()
                if work:
                    low[work[-1][0]] = min(low[work[-1][0]], low[v])
                if low[v] == index[v]:
                    comp = []
                    while True:
                        w = stack.pop()
                        on.discard(w)
                        comp.append(w)
                        if w == v:
                            break
                    out.append(comp)
    return out


def _classify(comp: list[str], by_id: dict, ix: RectIndex) -> dict:
    """輪の種類：cell_cycle（セルでも輪）／ordered（セルの順に計算できる。累計など）／unproven。"""
    inside = set(comp)
    desc = {"rules": comp, "ranges": [by_id[x].ref for x in comp]}
    if _ordered_by_position(comp, by_id, ix, inside):
        return {**desc, "kind": "ordered", "note": "規則どうしは互いを読むが、セルではより前の行・列だけを読むので、輪ではない（累計など）"}
    got = _cell_cycle(comp, by_id, ix, inside)
    if got is None:
        return {**desc, "kind": "unproven", "note": f"たどるセルが多すぎて（{CELL_SEARCH_LIMIT:,} 以上）、セルでも輪かどうかを確かめていない"}
    if got:
        for x in comp:
            by_id[x].in_cycle = "cell_cycle"
        return {**desc, "kind": "cell_cycle", "cells": got[:32], "note": "セルが自分自身を（間接に）読む循環参照。Excel の反復計算の設定なしには値が決まらない"}
    return {**desc, "kind": "ordered", "note": "規則どうしは互いを読むが、セルをたどると輪ではない"}


def _ordered_by_position(comp: list[str], by_id: dict, ix: RectIndex, inside: set[str]) -> bool:
    """輪の中の読み取りが、すべて「シートの中で前のセル」だけを読むなら、セルに輪はない（行の順か列の順のどちらかで）。"""
    by_row = by_col = True
    for x in comp:
        u = by_id[x]
        for rd in u.reads:
            if rd.external or rd.missing:
                continue
            for m in u.members:
                for corner in ((m.r1, m.c1), (m.r2, m.c2)):
                    rect = rd.at(*corner)
                    if rect is None:
                        continue
                    hit = [h for h in ix.hits(rd.sheet, *rect) if h[4] in inside]
                    if not hit:
                        continue
                    if rd.sheet != u.sheet:
                        return False
                    # 読む所を、輪の中の規則の長方形に切り詰めて、セルから見たずれの最大
                    dr = max(min(rect[2], h[2]) for h in hit) - corner[0]
                    dc = max(min(rect[3], h[3]) for h in hit) - corner[1]
                    if not (dr < 0 or (dr <= 0 and dc < 0)):
                        by_row = False
                    if not (dc < 0 or (dc <= 0 and dr < 0)):
                        by_col = False
                    if not by_row and not by_col:
                        return False
    return by_row or by_col


def _cell_cycle(comp: list[str], by_id: dict, ix: RectIndex, inside: set[str]) -> list[str] | None:
    """輪の中のセルをたどる。輪があれば、その輪のセル（番地）。なければ []。上限を超えたら None。"""
    color: dict[tuple, int] = {}
    visits = 0

    def kids(k):
        nonlocal visits
        sheet, r, c, uid = k
        u = by_id[uid]
        for rd in u.reads:
            if rd.external or rd.missing:
                continue
            rect = rd.at(r, c)
            if rect is None:
                continue
            for h in ix.hits(rd.sheet, *rect):
                if h[4] not in inside:
                    continue
                for rr in range(max(rect[0], h[0]), min(rect[2], h[2]) + 1):
                    for cc in range(max(rect[1], h[1]), min(rect[3], h[3]) + 1):
                        visits += 1
                        if visits > CELL_SEARCH_LIMIT:
                            raise OverflowError
                        yield (rd.sheet, rr, cc, h[4])

    try:
        for uid in comp:
            u = by_id[uid]
            for m in u.members:
                for r0 in range(m.r1, m.r2 + 1):
                    for c0 in range(m.c1, m.c2 + 1):
                        root = (u.sheet, r0, c0, uid)
                        if root in color:
                            continue
                        visits += 1
                        if visits > CELL_SEARCH_LIMIT:
                            raise OverflowError
                        color[root] = 1
                        path = [root]
                        work = [kids(root)]
                        while work:
                            for k in work[-1]:
                                st = color.get(k)
                                if st == 1:
                                    i = path.index(k)
                                    return [f"{p[0]}!{addr(p[1], p[2])}" for p in path[i:]]
                                if st is None:
                                    color[k] = 1
                                    path.append(k)
                                    work.append(kids(k))
                                    break
                            else:
                                work.pop()
                                color[path.pop()] = 2
    except OverflowError:
        return None
    return []


def exceptions_of(rules: list[Rule], sheets: list[SheetScan], ix: RectIndex, src) -> None:
    """規則の範囲の中で、規則と違うセル（別の数式・定数）。"""
    limit = 200
    want: dict[int, set] = {}
    for u in rules:
        if len(u.members) < 2:
            continue
        sh = sheets[u.si]
        gaps = _gaps(u)
        found = []
        n = 0
        for g in gaps:
            for h in ix.hits(u.sheet, *g):
                if h[4] == u.id:
                    continue
                r1, c1, r2, c2 = max(g[0], h[0]), max(g[1], h[1]), min(g[2], h[2]), min(g[3], h[3])
                if r1 <= r2 and c1 <= c2:
                    n += (r2 - r1 + 1) * (c2 - c1 + 1)
                    if len(found) < limit:
                        found.append({"range": rect_text(r1, c1, r2, c2), "kind": "other_formula", "rule": h[4]})
            for c in range(g[1], g[3] + 1):
                for a, b in sh.const.spans(c, g[0], g[2]):
                    n += b - a + 1
                    if len(found) < limit:
                        rec = {"range": rect_text(a, c, b, c), "kind": "constant"}
                        found.append(rec)
                        if a == b:
                            want.setdefault(u.si, set()).add((a, c))
                            rec["_at"] = (a, c)
        u.exceptions = found
        u.exception_count = n
    got = {si: src.fetch(si, cells) for si, cells in want.items()}
    for u in rules:
        for rec in u.exceptions:
            at = rec.pop("_at", None)
            if at is not None:
                from .structure import plain

                rec["value"] = plain(got.get(u.si, {}).get(at))


def _gaps(u: Rule) -> list[tuple[int, int, int, int]]:
    """外枠の中で、規則のセルでない所の長方形。"""
    rows = sorted({m.r1 for m in u.members} | {m.r2 + 1 for m in u.members} | {u.r1, u.r2 + 1})
    out = []
    members = sorted(u.members, key=lambda m: m.r1)
    next_m = 0
    active: list[Rect] = []
    for a, b in zip(rows, rows[1:]):
        while next_m < len(members) and members[next_m].r1 <= a:
            active.append(members[next_m])
            next_m += 1
        active = [m for m in active if m.r2 >= a]
        spans = sorted((m.c1, m.c2) for m in active if m.r2 >= b - 1)
        c = u.c1
        for x, y in spans:
            if x > c:
                out.append((a, c, b - 1, x - 1))
            c = max(c, y + 1)
        if c <= u.c2:
            out.append((a, c, b - 1, u.c2))
    return out
