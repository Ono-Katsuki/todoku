"""書いた値がどこまで届くか（数式の連動）と、「書いた所と、その計算結果のほかは変えていない」の確かめ。

連動は計算せずに、数式の字（参照・定義名・テーブル）から求める。別のシート・範囲・定義名・テーブルも数える。
セルは次のどれかになる（道具は事実を出す。推測で連動とも非連動とも決めない）：
- written    書いたセル
- dependent  書いたセルを読む数式（連動の連動まで）。配列数式は、結果の範囲のセルみな
- dynamic    連動を決められない：読む先が計算で決まる（INDIRECT・OFFSET・範囲の演算子など）・読めない数式・相対の参照のある名前と、それを読む数式
- timed      開くたびに変わる（TODAY・RAND など）と、それを読む数式
- なし        書いた値に連動しない。値も数式も、元と同じでなければならない

確かめは 2 つ。
- changes_scope / files_scope（連動の確認）：書いた後のファイルの、元と違うセルが、書いたセルか連動するセルか。計算しないので速い
- whole_calc（連動する所の確認）：書いた後のファイルをブック全体で計算し直し、保存された値と照らす。書き込みの部分計算とは別に、全部の行を読んで計算する。
  計算には同じ計算器（formula/）を使うので、関数の実装の誤りまでは見つけない。それは関数の全数照合（calc.verify）で別に確かめる
"""

from __future__ import annotations

import re
import time
from array import array
from collections import Counter
from dataclasses import dataclass, field

from ..formula.functions import VOLATILE
from ..formula.parser import (
    ArrayLit,
    Binary,
    Func,
    NameNode,
    Node,
    Postfix,
    RefNode,
    StructRef,
    Unary,
    addr,
    parse,
    parse_cell,
)
from ..formula.values import Rng, Unsupported

_DYNAMIC_FUNCS = {"INDIRECT", "OFFSET", "EVALUATE"}  # 読む先が計算で決まる
_TIMED_FUNCS = (set(VOLATILE) - _DYNAMIC_FUNCS) | {"RANDARRAY", "RTD"}  # 開くたびに変わる
_SIZED_BY_FIRST = {"SUMIF": (0, 2), "AVERAGEIF": (0, 2), "LOOKUP": (1, 2)}  # 3 つ目の範囲は、左上だけが効く。読む大きさは比べる範囲と同じ
MAX_OUTPUT = 1 << 20  # 1 つの配列数式の結果の範囲を 1 マスずつ数える上限。超えるものは、連動を決められない数式とする
_LEVELS = 21  # 行の数（2^20）の木の高さ


def _fname(name: str) -> str:
    u = str(name).upper()
    for pre in ("_XLFN.", "_XLWS."):
        if u.startswith(pre):
            u = u[len(pre):]
    return u


@dataclass
class _Shape:
    """数式の字から読み取った、読む先の型（共有数式は同じ型を、ずらして使う）。"""

    refs: list = field(default_factory=list)
    sized: list = field(default_factory=list)  # (比べる範囲, 左上だけが効く範囲)
    names: list = field(default_factory=list)
    tables: list = field(default_factory=list)
    why: set = field(default_factory=set)  # 連動を決められない理由
    timed: bool = False


def _shape_of(ast: Node) -> _Shape:
    s = _Shape()
    stack = [ast]
    while stack:
        n = stack.pop()
        if isinstance(n, RefNode):
            s.refs.append(n)
        elif isinstance(n, NameNode):
            s.names.append(n)
        elif isinstance(n, StructRef):
            s.tables.append(n)
        elif isinstance(n, Func):
            f = _fname(n.name)
            if f in _DYNAMIC_FUNCS or f.startswith("GET."):
                s.why.add(f)
            if f in _TIMED_FUNCS:
                s.timed = True
            pair = _SIZED_BY_FIRST.get(f)
            if pair is not None and len(n.args) == 3:
                a, b = n.args[pair[0]], n.args[pair[1]]
                if isinstance(a, RefNode) and isinstance(b, RefNode):
                    s.sized.append((a, b))
                else:
                    s.why.add(f"{f}（範囲が参照の字でない）")
            stack.extend(n.args)
        elif isinstance(n, (Unary, Postfix)):
            stack.append(n.operand)
        elif isinstance(n, Binary):
            if n.op == ":":
                s.why.add("範囲の演算子")
            stack.extend((n.left, n.right))
        elif isinstance(n, ArrayLit):
            pass
    return s


def _key(si: int, r: int, c: int) -> int:
    return (si << 36) | (r << 16) | c


class Reach:
    """連動の結果。セルの鍵の集合と、ブック全体を決められなかった印。"""

    def __init__(self, graph: "Graph", written: set[int], dependent: set[int], dynamic: set[int], timed: set[int]):
        self.graph = graph
        self.written = written
        self.dependent = dependent
        self.dynamic = dynamic
        self.timed = timed

    def kind(self, key: int) -> str | None:
        if key in self.written:
            return "written"
        if key in self.dependent:
            return "dependent"
        if key in self.dynamic or self.graph.incomplete:
            return "dynamic"
        if key in self.timed:
            return "timed"
        return None

    def special(self) -> set[int]:
        return self.written | self.dependent | self.dynamic | self.timed


class Graph:
    """どのセルの値が変わると、どの数式が読むか。範囲は（行の木 × 列）に入れるので、書いたセルごとに全部の範囲を見ない。"""

    def __init__(self, wb):
        self.wb = wb  # 定義名とテーブルを読む
        self.sheets: dict[str, int] = {s.name.lower(): i for i, s in enumerate(wb.sheets)}
        self.names: list[str] = [s.name for s in wb.sheets]
        self.outputs: list = []  # 数式の番号 → 結果のセルの鍵か、配列数式の範囲 (si, r1, c1, r2, c2)
        self.single: dict[int, array] = {}
        self.rects: list[tuple[int, int, int, int, int]] = []
        self.rect_ids: dict[tuple, int] = {}
        self.rect_formulas: list[array] = []
        self.buckets: dict[int, array] = {}
        self.dynamic_roots: list[int] = []
        self.timed_roots: list[int] = []
        self.why: Counter = Counter()  # 連動を決められない理由 → 数式の数
        self.incomplete = False  # 数式を全部は読めなかった：どのセルも連動を決められない
        self._shapes: dict[int, tuple] = {}  # 構文木 → (構文木, 型)
        self._names: dict[tuple, tuple] = {}
        self._calc = None

    # ------------------------------------------------------------ 作る
    def key(self, sheet: str, r: int, c: int) -> int | None:
        si = self.sheets.get(str(sheet).lower())
        return None if si is None else _key(si, r, c)

    def unkey(self, k: int) -> tuple[str, int, int]:
        return self.names[k >> 36], (k >> 16) & 0xFFFFF, k & 0xFFFF

    def _shape(self, ast: Node) -> _Shape:
        got = self._shapes.get(id(ast))
        if got is None:
            got = self._shapes[id(ast)] = (ast, _shape_of(ast))
        return got[1]

    def add(self, sheet: str, r: int, c: int, ast: Node | None, dr: int = 0, dc: int = 0, output: str | None = None,
            why: str | None = None) -> None:
        """数式を 1 つ入れる。ast：None は読めない数式。output：配列数式の結果の範囲（「A1:C3」）。why：字を読む前に決められないと分かっているとき。"""
        si = self.sheets[sheet.lower()]
        fid = len(self.outputs)
        out: object = _key(si, r, c)
        reasons: set = {why} if why else set()
        if output:
            try:
                a, _, b = output.partition(":")
                r1, c1 = parse_cell(a.replace("$", ""))
                r2, c2 = parse_cell(b.replace("$", "")) if b else (r1, c1)
                if (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_OUTPUT:
                    reasons.add("結果の範囲がとても広い配列数式")
                elif (r1, c1) != (r2, c2):
                    out = (si, r1, c1, r2, c2)
            except (ValueError, KeyError):
                reasons.add("結果の範囲を読めない配列数式")
        self.outputs.append(out)
        timed = False
        if ast is None:
            reasons.add(why or "読めない数式")
        else:
            shape = self._shape(ast)
            reasons |= shape.why
            timed = shape.timed
            here = sheet
            for ref in shape.refs:
                self._add_ref(ref.shifted(dr, dc), here, fid, reasons)
            for a_, b_ in shape.sized:
                a2, b2 = a_.shifted(dr, dc), b_.shifted(dr, dc)
                self._add_rect(b2.sheet or here, b2.r1, b2.c1, b2.r1 + a2.r2 - a2.r1, b2.c1 + a2.c2 - a2.c1, fid, reasons)
            for node in shape.tables:
                self._add_table(node, here, r, c, fid, reasons)
            for node in shape.names:
                rects, why_n, timed_n = self._name_reads(node, here, frozenset())
                reasons |= why_n
                timed = timed or timed_n
                for (s2, r1, c1, r2, c2) in rects:
                    self._add_rect(s2, r1, c1, r2, c2, fid, reasons)
        if reasons:
            self.dynamic_roots.append(fid)
            self.why.update(reasons)
        if timed:
            self.timed_roots.append(fid)

    def _add_ref(self, ref: RefNode, here: str, fid: int, reasons: set) -> None:
        self._add_rect(ref.sheet or here, ref.r1, ref.c1, ref.r2, ref.c2, fid, reasons)

    def _add_rect(self, sheet: str, r1: int, c1: int, r2: int, c2: int, fid: int, reasons: set) -> None:
        if sheet.startswith("["):
            return  # 外部のブック：保存された値を読むだけで、書いた値は届かない
        si = self.sheets.get(sheet.lower())
        if si is None:
            reasons.add(f"シートが決まらない参照（{sheet}）")
            return
        if r1 == r2 and c1 == c2:
            self.single.setdefault(_key(si, r1, c1), array("I")).append(fid)
            return
        rect = (si, r1, c1, r2, c2)
        rid = self.rect_ids.get(rect)
        if rid is None:
            rid = self.rect_ids[rect] = len(self.rects)
            self.rects.append(rect)
            self.rect_formulas.append(array("I"))
            col = c1 if c1 == c2 else 0  # 1 列の範囲は列で分ける（2 列以上はまとめて、読むときに列を比べる）
            lo, hi = max(r1, 0), min(r2, (1 << 20)) + 1
            level = 0
            while lo < hi:
                if lo & 1:
                    self.buckets.setdefault(_bucket(si, col, level, lo), array("I")).append(rid)
                    lo += 1
                if hi & 1:
                    hi -= 1
                    self.buckets.setdefault(_bucket(si, col, level, hi), array("I")).append(rid)
                lo >>= 1
                hi >>= 1
                level += 1
        self.rect_formulas[rid].append(fid)

    def _calc_ctx(self):
        if self._calc is None:
            from .calc import Calc

            self._calc = Calc(self.wb)
        return self._calc

    def _add_table(self, node: StructRef, here: str, r: int, c: int, fid: int, reasons: set) -> None:
        rng = self._calc_ctx().table_ref(node, here, r, c)
        if isinstance(rng, Rng):
            self._add_rect(rng.sheet, rng.r1, rng.c1, rng.r2, rng.c2, fid, reasons)

    def _name_reads(self, node: NameNode, here: str, seen: frozenset) -> tuple[list, set, bool]:
        """名前が読む範囲：(範囲の並び, 決められない理由, 開くたびに変わるか)。相対の参照は使う所で決まるので、決められないとする。"""
        low = node.name.lower()
        key = (low, node.sheet or here)
        got = self._names.get(key)
        if got is not None:
            return got
        if low in seen:
            return [], {f"名前が自分を読む（{node.name}）"}, False
        try:
            d = self.wb.defined_name(node.name, node.sheet or here)
        except Unsupported as e:
            return [], {f"名前を読めない（{e}）"}, False
        if d is None:
            got = ([], set(), False)  # 定義のない名前（LET・LAMBDA の中の名前か #NAME?）：読む先はほかの引数の字にある
        elif d.ast is None:
            got = ([], {f"名前を読めない（{node.name}）"}, False)
        else:
            scope = d.sheet_scope or here
            s = self._shape(d.ast)
            rects: list = []
            why = set(s.why)
            timed = s.timed
            for ref in s.refs:
                if not (ref.abs_r1 and ref.abs_r2 and ref.abs_c1 and ref.abs_c2):
                    why.add("相対の参照のある名前")
                    continue
                rects.append((ref.sheet or scope, ref.r1, ref.c1, ref.r2, ref.c2))
            for a_, b_ in s.sized:
                rects.append((b_.sheet or scope, b_.r1, b_.c1, b_.r1 + a_.r2 - a_.r1, b_.c1 + a_.c2 - a_.c1))
            if s.tables:
                for t in s.tables:
                    if t.table is None or "#This Row" in t.specials:
                        why.add("使う所で決まるテーブルの参照のある名前")
                    else:
                        rng = self._calc_ctx().table_ref(t, scope, 1, 1)
                        if isinstance(rng, Rng):
                            rects.append((rng.sheet, rng.r1, rng.c1, rng.r2, rng.c2))
            for n2 in s.names:
                r2, w2, t2 = self._name_reads(n2, scope, seen | {low})
                rects += r2
                why |= w2
                timed = timed or t2
            got = (rects, why, timed)
        self._names[key] = got
        return got

    # ------------------------------------------------------------ たどる
    def _outputs(self, fid: int):
        out = self.outputs[fid]
        if isinstance(out, tuple):
            si, r1, c1, r2, c2 = out
            return (_key(si, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1))
        return (out,)

    def _closure(self, keys, fids=()) -> set[int]:
        """keys のセルと fids の数式から、読まれる先をたどった、数式の結果のセルの鍵（keys 自身は、読まれて数式の結果でなければ入らない）。"""
        seen_f: set[int] = set()
        fired: set[int] = set()
        out: set[int] = set()
        todo = list(keys)

        def fire(fid: int) -> None:
            if fid not in seen_f:
                seen_f.add(fid)
                for k in self._outputs(fid):
                    if k not in out:
                        out.add(k)
                        todo.append(k)

        for f in fids:
            fire(f)
        visited: set[int] = set()
        while todo:
            k = todo.pop()
            for f in self.single.get(k, ()):
                fire(f)
            si, r, c = k >> 36, (k >> 16) & 0xFFFFF, k & 0xFFFF
            for level in range(_LEVELS):
                i = r >> level
                b = _bucket(si, c, level, i)
                if b in self.buckets and b not in visited:
                    visited.add(b)
                    for rid in self.buckets[b]:
                        if rid not in fired:
                            fired.add(rid)
                            for f in self.rect_formulas[rid]:
                                fire(f)
                b = _bucket(si, 0, level, i)
                if b in self.buckets:
                    for rid in self.buckets[b]:
                        if rid not in fired:
                            rect = self.rects[rid]
                            if rect[2] <= c <= rect[4]:
                                fired.add(rid)
                                for f in self.rect_formulas[rid]:
                                    fire(f)
        return out

    def readers(self, keys) -> set[int]:
        """keys のセルを読む数式の結果のセル（連動の連動まで）。"""
        return self._closure(keys)

    def reach(self, written) -> Reach:
        """written：書いたセル (シート, 行, 列) の並び。"""
        keys = {k for k in (self.key(*w) for w in written) if k is not None}
        return Reach(self, keys, self._closure(keys), self._closure((), self.dynamic_roots), self._closure((), self.timed_roots))


def _bucket(si: int, col: int, level: int, i: int) -> int:
    return (si << 41) | (col << 26) | (level << 21) | i


def from_workbook(wb) -> Graph:
    """全部のセルを読んだブック（.xlsx・.xls）から。結果の範囲のある配列数式は親のセルだけを入れる（一部のセルは、親の結果）。"""
    g = Graph(wb)
    for sh in wb.sheets:
        for cell in sh.cells.values():
            if cell.array_master is not None:
                continue
            if cell.formula_kind == "dataTable":
                g.add(sh.name, cell.row, cell.col, None, output=cell.array_ref, why="データテーブル")
            elif cell.has_formula or cell.parse_error:
                g.add(sh.name, cell.row, cell.col, cell.ast, cell.dr, cell.dc, output=cell.array_ref if cell.formula_kind == "array" or cell.dynamic else None)
    return g


# 値のない <c .../> を飛ばし、ほかのセルの <f> を取り違えない
_FORMULA_CELL = re.compile(rb"<(?:\w+:)?c\b[^>]*?\br=\"([A-Z]+\d+)\"[^>]*?(?<!/)>(?:(?!</?(?:\w+:)?c\b).)*?<(?:\w+:)?f\b([^>]*?)(?:/>|>(.*?)</(?:\w+:)?f>)", re.S)
_STEP = 4096  # 一度に読む行の数


def _unesc(s: str) -> str:
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def from_lazy(doc) -> Graph:
    """大きなブック（lazy/xlsxdoc.XlsxLazy）から。ファイルの XML を行のかたまりごとに読む（書き込みの索引 FormulaIndex は使わない）。定義名とテーブルは、行のない小さなブックから読む。"""
    g = Graph(doc.book(doc.slice_blob({})))
    asts: dict[str, Node | None] = {}

    def ast_of(text: str):
        if text in asts:
            return asts[text]
        try:
            got = parse(text)
        except Exception:  # noqa: BLE001  読めない数式：連動は決められない
            got = None
        if len(asts) < 200_000:
            asts[text] = got
        return got

    for meta in doc.sheet_meta:
        ix = meta.get("index")
        if ix is None or not ix.formulas:
            continue
        name = meta["name"]
        shared: dict[bytes, tuple] = {}
        found = 0
        for i in range(0, len(ix.rows), _STEP):
            xml = ix.row_bytes(i, min(i + _STEP, len(ix.rows)))
            for m in _FORMULA_CELL.finditer(xml):
                found += 1
                r, c = parse_cell(m[1].decode())
                attrs = m[2]
                text = _unesc((m[3] or b"").decode("utf-8", "replace"))
                t = re.search(rb'\bt="(\w+)"', attrs)
                kind = t[1] if t else b"normal"
                ref = re.search(rb'\bref="([^"]+)"', attrs)
                output = ref[1].decode() if ref and kind in (b"array", b"dataTable") else None
                if kind == b"dataTable":
                    g.add(name, r, c, None, output=output, why="データテーブル")
                    continue
                if kind == b"shared":
                    si = re.search(rb'\bsi="(\d+)"', attrs)
                    if si and text:
                        shared[si[1]] = (text, r, c)
                    elif si and si[1] in shared:
                        text, r0, c0 = shared[si[1]]
                        g.add(name, r, c, ast_of(text), r - r0, c - c0)
                        continue
                    else:
                        g.add(name, r, c, None, why="親の見つからない共有数式")
                        continue
                g.add(name, r, c, ast_of(text), output=output)
        if found != ix.formulas:
            g.incomplete = True  # セルの番地のない数式など：全部は読めなかった
    return g


# ====================================================================== 確かめた範囲
@dataclass
class Scope:
    """確かめた範囲（数）と、見つけた問題。"""

    cells: int = 0  # ファイルにあるセル（書いた後）
    written: int = 0
    not_dependent: int = 0  # 書いた値に連動しないセル
    not_dependent_same: int = 0  # そのうち、元と同じだったセル
    dependent: int = 0  # 連動するセル
    dependent_changed: int = 0
    undetermined: int = 0  # 連動を決められないセル
    undetermined_changed: int = 0
    timed: int = 0  # 開くたびに変わるセル
    timed_changed: int = 0
    why: dict = field(default_factory=dict)
    problems: list = field(default_factory=list)
    changed: set = field(default_factory=set)  # 変わった数式のセルの鍵（全体の計算で使う）
    calc: dict = field(default_factory=dict)  # 全体の計算
    seeds: list = field(default_factory=list)  # 長い仕事に渡す、書いたセルと変わったセル
    changed_cells: list = field(default_factory=list)
    marked: bool = False
    book: object = field(default=None, repr=False, compare=False)  # 開いた、書いた後のブック（読むだけで使う。結果には出さない）

    def note_cells(self, graph: Graph, seeds: set) -> None:
        self.seeds = sorted(seeds)
        self.changed_cells = sorted(graph.unkey(k) for k in self.changed)

    def to_dict(self) -> dict:
        d = {"not_dependent": {"cells": self.not_dependent, "same": self.not_dependent_same},
             "written": {"cells": self.written},
             "dependent": {"cells": self.dependent, "changed": self.dependent_changed},
             "undetermined": {"cells": self.undetermined + self.timed, "changed": self.undetermined_changed + self.timed_changed,
                              "reads_changing_refs": self.undetermined, "changes_each_open": self.timed, "why": dict(self.why)},
             "recalculated": dict(self.calc)}
        return d

    def sentence(self) -> str:
        out = f"書いた値に連動しないセル {self.not_dependent:,} か所のうち {self.not_dependent_same:,} か所が元と同じ（数式の参照から求めました）"
        und = self.undetermined + self.timed
        if und:
            ch = self.undetermined_changed + self.timed_changed
            out += f"。連動を決められないセル {und:,} か所（読む先が計算で決まる・開くたびに変わる。そのうち値が変わったのは {ch:,} か所）"
        c = self.calc
        if c.get("status") == "done":
            out += f"。ブック全体を計算し直して、連動するセル {c['checked']:,} か所のうち {c['same']:,} か所が保存された値と一致"
            if c.get("cannot_calculate"):
                out += f"（計算できないセル {c['cannot_calculate']:,} か所は照らしていません）"
            if c.get("stale"):
                out += f"。書いた値をまだ反映していないセル {c['stale']:,} か所（開いたときに計算し直す印があります）"
        elif c.get("status") == "deferred":
            out += "。ブック全体の計算し直しとの照合は、まだしていません（未確認）"
        return out


def seeds_of(res) -> set[tuple[str, int, int]]:
    """書き込みの結果から、書いたセル (シート, 行, 列)。"""
    out = set()
    for sheet, cell in res.allowed_inputs:
        r, c = parse_cell(cell)
        out.add((sheet, r, c))
    return out


def changes_scope(wb, graph: Graph, reach: Reach, changes: list[dict]) -> Scope:
    """全部を読んだブックの道：証明が 2 つのファイルを読み直して比べた、値の変わったセル（Proof.cell_changes）を分ける。"""
    s = Scope()
    s.why = dict(graph.why)
    moved: dict[int, dict] = {}
    for ch in changes:
        r, c = parse_cell(ch["cell"])
        k = graph.key(ch["sheet"], r, c)
        if k is not None:
            moved[k] = ch
    present = set()
    for sh in wb.sheets:
        si = graph.sheets[sh.name.lower()]
        for (r, c) in sh.cells:
            k = _key(si, r, c)
            present.add(k)
            s.cells += 1
            _count(s, reach, k, k in moved)
    for k, ch in moved.items():
        if k not in present:
            s.cells += 1
            _count(s, reach, k, True)
        kind = reach.kind(k)
        if kind is None:
            s.problems.append(f"{ch['sheet']}!{ch['cell']}: 書いた値に連動しないセルの値が変わりました（{ch.get('before')!r} → {ch.get('after')!r}）")
        elif kind != "written":
            s.changed.add(k)
    return s


def _count(s: Scope, reach: Reach, k: int, changed: bool) -> None:
    kind = reach.kind(k)
    if kind is None:
        s.not_dependent += 1
        if not changed:
            s.not_dependent_same += 1
    elif kind == "written":
        s.written += 1
    elif kind == "dependent":
        s.dependent += 1
    elif kind == "dynamic":
        s.undetermined += 1
    else:
        s.timed += 1
    if changed and kind is not None:
        _count_changed(s, kind)


def _count_changed(s: Scope, kind: str) -> None:
    if kind == "dependent":
        s.dependent_changed += 1
    elif kind == "dynamic":
        s.undetermined_changed += 1
    elif kind == "timed":
        s.timed_changed += 1


_CELL = re.compile(rb"<(?:\w+:)?c\b[^>]*?\br=\"([A-Z]+\d+)\"[^>]*?(?:/>|>.*?</(?:\w+:)?c>)", re.S)
_CELL_F = re.compile(rb"<(?:\w+:)?f\b[^>]*?(?:/>|>.*?</(?:\w+:)?f>)", re.S)
_CELL_S = re.compile(rb"^<[^>]*?\bs=\"(\d+)\"")


def files_scope(doc_a, doc_b, graph: Graph, reach: Reach, renamed: set | None = None) -> Scope:
    """大きなブックの道：元と書いた後のファイルを、シートの XML で比べる。行のかたまりがバイト一致なら、中のセルは数えるだけ。
    違うかたまりは、セルごとに比べて分ける。doc_a・doc_b は開き直した XlsxLazy（書き込みの内部を使わない）。"""
    s = Scope()
    s.why = dict(graph.why)
    renamed = renamed or set()
    special = reach.special()
    rows_with: dict[int, set[int]] = {}
    for k in special:
        rows_with.setdefault(k >> 36, set()).add((k >> 16) & 0xFFFFF)
    by_name = {d["name"]: d for d in doc_b.sheet_meta}
    for meta in doc_a.sheet_meta:
        d2 = by_name.get(meta["name"])
        if d2 is None:
            s.problems.append(f"シート {meta['name']} がなくなりました")
            continue
        ia, ib = meta.get("index"), d2.get("index")
        if ia is None or ib is None:
            continue
        si = graph.sheets[meta["name"].lower()]
        marked = rows_with.get(si, ())
        pat = re.compile(b"<" + re.escape(ia.prefix) + rb"c[\s/>]")
        rows = ia.rows
        lo = 0  # 前のかたまりの続きの行
        for i in range(0, max(len(rows), 1), _STEP):
            last = i + _STEP >= len(rows)
            hi = (1 << 20) + 1 if last else rows[i + _STEP - 1]
            xa = _rows(ia, lo, hi)
            xb = _rows(ib, lo, hi)
            if xa == xb and (not marked or not _any_in(marked, lo, hi)):
                n = len(pat.findall(xa))
                s.cells += n
                s.not_dependent += n
                s.not_dependent_same += n
            else:
                _compare_cells(s, reach, graph, si, meta["name"], xa, xb, xa == xb, renamed)
            lo = hi + 1
    return s


def _rows(ix, lo: int, hi: int) -> bytes:
    i, j = ix.locate(lo, hi)
    return bytes(ix.row_bytes(i, j)) if i < j else b""


def _any_in(marked, lo: int, hi: int) -> bool:
    if hi - lo < len(marked):
        return any(r in marked for r in range(lo, hi + 1))
    return any(lo <= r <= hi for r in marked)


def _compare_cells(s: Scope, reach: Reach, graph: Graph, si: int, name: str, xa: bytes, xb: bytes, same: bool, renamed: set) -> None:
    ca = {m[1]: m[0] for m in _CELL.finditer(xa)}
    cb = ca if same else {m[1]: m[0] for m in _CELL.finditer(xb)}
    for ref in ca.keys() | cb.keys():
        a, b = ca.get(ref), cb.get(ref)
        r, c = parse_cell(ref.decode())
        k = _key(si, r, c)
        kind = reach.kind(k)
        if a == b:
            _count(s, reach, k, False)
            continue
        _count(s, reach, k, True)
        where = f"{name}!{ref.decode()}"
        if kind is None:
            s.problems.append(f"{where}: 書いた値に連動しないセルの値か数式が変わりました")
            continue
        if kind == "written":
            continue
        s.changed.add(k)
        if (name, ref.decode()) in renamed:
            continue
        fa = _CELL_F.search(a or b"")
        fb = _CELL_F.search(b or b"")
        if (fa[0] if fa else None) != (fb[0] if fb else None):
            s.problems.append(f"{where}: 数式が変わりました")
        sa, sb = _CELL_S.search(a or b""), _CELL_S.search(b or b"")
        if a is not None and b is not None and (sa[1] if sa else None) != (sb[1] if sb else None):
            s.problems.append(f"{where}: 書式が変わりました")


def prove_whole(wb, new_blob: bytes, name: str, seeds: set, changes: list[dict], full_calc: bool = False, seconds: float | None = None) -> Scope:
    """全部を読んだブックの道の確かめ：連動の確認と、ブック全体の計算し直し。wb は元のブック。
    full_calc：書き手が開いたときに計算し直す印を付けた。seconds：計算にかけてよい秒数（過ぎたら未確認として返す）。"""
    from ..book import open_book
    from ..limits import LimitExceeded

    graph = from_workbook(wb)
    reach = graph.reach(seeds)
    s = changes_scope(wb, graph, reach, changes)
    s.note_cells(graph, seeds)
    s.marked = full_calc
    try:
        wb_new = s.book = open_book(new_blob, name)
        s.marked = has_full_calc_mark(wb_new) or full_calc
        s.calc, bad = whole_calc(wb_new, graph, reach, s.changed, s.marked, lambda: wb, None if seconds is None else time.monotonic() + seconds)
    except LimitExceeded:  # この呼び出しの予算に入らない：長い仕事で確かめる
        s.calc = {"status": "deferred"}
        return s
    s.problems += bad
    return s


def add_scope(result: dict, scope: Scope) -> None:
    """書き込みの結果に、確かめた範囲を足す。proof_ok：問題があれば偽にする（いままでの意味のまま：確かめて、想定外の変更がない）。
    全体の計算を済ませていないときは、deferred（未確認）に書く。確かめたとは言わない。"""
    result["proof_scope"] = scope.to_dict()
    result["proof"] = (result.get("proof") or "") + "。" + scope.sentence()
    if scope.problems:
        result["proof_ok"] = False
        result["unexpected"] = list(result.get("unexpected") or []) + scope.problems[:20]
    if scope.calc.get("status") == "deferred":
        msg = ("未確認：ブック全体を計算し直して、書き込んだ計算結果と照らすこと（時間がかかるので、書き込みの中では済ませていません。"
               "audit_document に、書き出したファイルを渡すと確かめて、結果を返します）")
        result["deferred"] = list(result.get("deferred") or []) + [msg]
        result["warnings"] = list(result.get("warnings") or []) + [msg]


# ====================================================================== 全体の計算
def has_full_calc_mark(wb) -> bool:
    """開いたときに計算し直す印があるか（書いた値をまだ反映していないセルがあることを、ファイルが知らせている）。"""
    if getattr(wb, "kind", "xlsx") == "xls":
        return bool(getattr(wb, "full_calc_on_load", False))
    path = getattr(wb, "wb_path", None)
    if not path or not wb.zip.has(path):
        return False
    return bool(re.search(rb"<(?:\w+:)?calcPr\b[^>]*?\bfullCalcOnLoad=\"(?:1|true)\"", wb.zip.read(path)))


def whole_calc(wb_new, graph: Graph, reach: Reach, changed: set[int], marked: bool, open_original=None, deadline: float | None = None) -> tuple[dict, list[str]]:
    """書いた後のブック wb_new（全部の行を読んだもの）を計算し直し、保存された値と照らす：(範囲の数, 問題)。
    連動するセルの差は不合格（ただし、書いたときに値を変えなかったセルで、開いたときに計算し直す印のあるものは「まだ反映していない」と数える）。
    連動しないセルは元とバイト一致なので、差があれば元のファイルの差か、連動の見落としのどちらか：元のブックでも計算して見分ける。
    計算できない数式は計算できないと数え、推測で埋めない。deadline（time.monotonic の時刻）を過ぎたら、照らしたことにせず deferred を返す。"""
    from .calc import Calc, _run_deep, formula_cells, values_equal

    out = {"status": "done", "checked": 0, "same": 0, "cannot_calculate": 0, "stale": 0, "time_dependent": 0,
           "undetermined_checked": 0, "undetermined_same": 0, "others_checked": 0, "others_same": 0, "others_already_different": 0,
           "others_cannot_calculate": 0}
    problems: list[str] = []
    calc_o: list = []

    def original_value(sheet: str, r: int, c: int):
        if open_original is None:
            return None, False
        if not calc_o:
            calc_o.append(Calc(open_original()))
        co = calc_o[0]
        try:
            v = co.cell_value(sheet, r, c)
        except Exception:  # noqa: BLE001
            return None, False
        return v, co.status.get((sheet, r, c), ("ok", ""))[0] == "ok"

    def run():
        calc = Calc(wb_new)
        for n, cell in enumerate(formula_cells(wb_new)):
            if deadline is not None and not n % 256 and time.monotonic() > deadline:
                out["status"] = "deferred"
                return
            k = graph.key(cell.sheet, cell.row, cell.col)
            kind = reach.kind(k) if k is not None else "dynamic"
            where = f"{cell.sheet}!{cell.address}"
            if kind == "written":
                continue  # 数式の上に書いた値（数式は残らない）など
            key = (cell.sheet, cell.row, cell.col)
            try:
                got = calc.cell_value(*key)
                st = calc.status.get(key, ("ok", ""))[0] if cell.has_formula else "ok"
            except Exception:  # noqa: BLE001  計算できない数式は、数えて伝える
                st, got = "unsupported", None
            if kind == "dependent":
                if st not in ("ok", "volatile"):
                    out["cannot_calculate"] += 1
                elif k in reach.timed:
                    out["time_dependent"] += 1
                else:
                    out["checked"] += 1
                    if values_equal(cell.value, got):
                        out["same"] += 1
                    elif k not in changed and (marked or cell.value is None):
                        out["stale"] += 1
                    else:
                        problems.append(f"{where}: 保存された値 {cell.value!r} と、ブック全体を計算し直した値 {got!r} が違います")
            elif kind == "dynamic":
                if st not in ("ok", "volatile"):
                    out["cannot_calculate"] += 1
                elif k in reach.timed:
                    out["time_dependent"] += 1
                else:
                    out["undetermined_checked"] += 1
                    if values_equal(cell.value, got):
                        out["undetermined_same"] += 1
                    elif k in changed:
                        problems.append(f"{where}: 保存された値 {cell.value!r} と、ブック全体を計算し直した値 {got!r} が違います")
                    else:
                        out["stale"] += 1
            elif kind == "timed":
                out["time_dependent"] += 1
            elif st not in ("ok", "volatile"):
                out["others_cannot_calculate"] += 1
            else:
                out["others_checked"] += 1
                if values_equal(cell.value, got):
                    out["others_same"] += 1
                    continue
                before, ok = original_value(*key)
                if ok and not values_equal(before, got):
                    problems.append(f"{where}: 書いた値に連動しないと求めたセルですが、計算し直すと書いた値で変わります（{before!r} → {got!r}。連動の見落としです）")
                else:
                    out["others_already_different"] += 1

    _run_deep(run)
    return out, problems
