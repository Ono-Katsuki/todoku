"""規則のまわりの構造：入力・定数の値（パラメータ）・表（参照表）・出力・名前・分岐・動く参照・外部参照・隠れた所・シートのあいだの流れ。"""

from __future__ import annotations

import bisect

from ..formula.functions import EAGER, LAZY, VOLATILE
from ..formula.parser import (
    Binary,
    Func,
    NameNode,
    Node,
    Postfix,
    RefNode,
    Str,
    Unary,
    addr,
    iter_funcs,
    iter_refs,
)
from ..formula.values import ExcelError
from .rules import Read, RectIndex, Rule, key_text, read_rects, rect_text
from .scan import SheetScan
from .shape import render

LISTED = 64  # 1 つの入力について示す規則の数（数は n にすべて数える）
TABLE_INLINE_CELLS = 400  # 参照表の中身を結果にそのまま入れるマスの数（超える表は、はじめの行と read_table の頼み方）
TABLE_PREVIEW_ROWS = 5


def plain(v):
    """値を JSON に入れる形に。"""
    if isinstance(v, ExcelError):
        return v.code
    if isinstance(v, float) and v.is_integer() and abs(v) < 1e15:
        return int(v)
    return v


# ---------------------------------------------------------------- 区間の計算（列ごと）
def _merge(iv: list[tuple[int, int]]) -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for a, b in sorted(iv):
        if out and a <= out[-1][1] + 1:
            if b > out[-1][1]:
                out[-1] = (out[-1][0], b)
        else:
            out.append((a, b))
    return out


def _minus(iv: list[tuple[int, int]], cut: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """iv から cut を除く（どちらも _merge した並び）。"""
    out = []
    j = 0
    for a, b in iv:
        cur = a
        while j < len(cut) and cut[j][1] < cur:
            j += 1
        k = j
        while k < len(cut) and cut[k][0] <= b:
            if cut[k][0] > cur:
                out.append((cur, cut[k][0] - 1))
            cur = max(cur, cut[k][1] + 1)
            k += 1
        if cur <= b:
            out.append((cur, b))
    return out


def _to_rects(by_col: dict[int, list[tuple[int, int]]]) -> list[tuple[int, int, int, int]]:
    """列ごとの区間を長方形に（隣の列と区間が同じなら 1 つに）。"""
    open_: dict[tuple[int, int], tuple[int, int]] = {}  # (r1, r2) → (c1, c2)
    out = []
    for c in sorted(by_col):
        cur = set(by_col[c])
        for k in [k for k in open_ if k not in cur or open_[k][1] != c - 1]:
            out.append((k[0], open_[k][0], k[1], open_[k][1]))
            del open_[k]
        for k in by_col[c]:
            if k in open_:
                open_[k] = (open_[k][0], c)
            else:
                open_[k] = (c, c)
    for k, (c1, c2) in open_.items():
        out.append((k[0], c1, k[1], c2))
    out.sort()
    return out


def _clip_cols(rect, max_col: int):
    return range(rect[1], min(rect[3], max(max_col, rect[1])) + 1)


# ---------------------------------------------------------------- 見出しの事実
class Labels:
    """セルの左と上にある、いちばん近い字のセル（「見出し」と決めつけず、事実として示す）。"""

    def __init__(self, sheets: list[SheetScan], src):
        self.text_cols: dict[int, list[int]] = {}
        self.sheets = sheets
        self.src = src
        self.merges: dict[int, tuple[dict, dict]] = {}
        self.plans: dict[tuple[int, int, int], dict] = {}

    def _m(self, si: int):
        got = self.merges.get(si)
        if got is None:
            rows: dict[int, list] = {}
            cols: dict[int, list] = {}
            meta = self.src.meta.sheets[si]
            room = 1 << 20
            for r1, c1, r2, c2 in getattr(meta, "merges", ()):
                if not self.sheets[si].text.has(r1, c1):
                    continue
                room -= (r2 - r1 + 1) + (c2 - c1 + 1)
                if room < 0:
                    break
                for r in range(r1, r2 + 1):
                    rows.setdefault(r, []).append((c2, r1, c1))
                for c in range(c1, c2 + 1):
                    cols.setdefault(c, []).append((r2, r1, c1))
            got = self.merges[si] = (rows, cols)
        return got

    def plan(self, si: int, r: int, c: int) -> dict:
        k = (si, r, c)
        got = self.plans.get(k)
        if got is not None:
            return got
        sh = self.sheets[si]
        rows, cols = self._m(si)
        left = None
        cols_sorted = self.text_cols.get(si)
        if cols_sorted is None:
            cols_sorted = self.text_cols[si] = sorted(sh.text.cols)
        for i in range(bisect.bisect_left(cols_sorted, c) - 1, -1, -1):
            if sh.text.has(r, cols_sorted[i]):
                left = (r, cols_sorted[i])
                break
        for c2, r1, c1 in rows.get(r, ()):
            if c2 < c and (left is None or c1 > left[1]):
                left = (r1, c1)
        last = sh.text.last_at_or_before(c, r - 1)
        above = (last, c) if last else None
        for r2, r1, c1 in cols.get(c, ()):
            if r2 < r and (above is None or r1 > above[0]):
                above = (r1, c1)
        got = self.plans[k] = {"left": left, "above": above}
        return got

    def resolve(self) -> None:
        """計画したセルの字を、まとめて読む。"""
        want: dict[int, set] = {}
        for (si, _r, _c), p in self.plans.items():
            for at in p.values():
                if at:
                    want.setdefault(si, set()).add(at)
        self.got = {si: self.src.fetch(si, cells) for si, cells in want.items()}

    def facts(self, si: int, r: int, c: int) -> dict | None:
        p = self.plans.get((si, r, c))
        if not p:
            return None
        out = {}
        for side, at in p.items():
            if at:
                v = self.got.get(si, {}).get(at)
                if isinstance(v, str) and v.strip():
                    out[side] = {"cell": addr(*at), "text": v}
        return out or None


# ---------------------------------------------------------------- 分岐
def expr(node: Node, u: Rule) -> str:
    return render(node, u.ref_pattern)


BRANCHING = {"IF", "IFS", "SWITCH", "IFERROR", "IFNA", "CHOOSE"}


def branches(node: Node, u: Rule) -> dict | str:
    """IF・IFS・SWITCH の木。分岐のない式は字のまま。"""
    if isinstance(node, Func) and node.name in BRANCHING:
        a = node.args
        if node.name == "IF" and a:
            out = {"kind": "if", "condition": expr(a[0], u), "then": branches(a[1], u) if len(a) > 1 else "TRUE"}
            out["else"] = branches(a[2], u) if len(a) > 2 else "FALSE"
            return out
        if node.name == "IFS":
            return {"kind": "ifs", "cases": [{"condition": expr(a[i], u), "then": branches(a[i + 1], u)}
                                             for i in range(0, len(a) - 1, 2)]}
        if node.name == "SWITCH" and a:
            rest = a[1:]
            out = {"kind": "switch", "on": expr(a[0], u),
                   "cases": [{"value": expr(rest[i], u), "then": branches(rest[i + 1], u)} for i in range(0, len(rest) - 1, 2)]}
            if len(rest) % 2:
                out["default"] = branches(rest[-1], u)
            return out
        if node.name in ("IFERROR", "IFNA") and len(a) == 2:
            return {"kind": node.name.lower(), "value": branches(a[0], u), "on_error": branches(a[1], u)}
        if node.name == "CHOOSE" and a:
            return {"kind": "choose", "index": expr(a[0], u), "options": [branches(x, u) for x in a[1:]]}
    inner = _inner_branches(node, u)
    if inner:
        return {"kind": "expression", "text": expr(node, u), "contains": inner}
    return expr(node, u)


def _inner_branches(node: Node, u: Rule) -> list:
    out = []
    kids = []
    if isinstance(node, Func):
        kids = node.args
    elif isinstance(node, Binary):
        kids = [node.left, node.right]
    elif isinstance(node, (Unary, Postfix)):
        kids = [node.operand]
    for k in kids:
        if isinstance(k, Func) and k.name in BRANCHING:
            out.append(branches(k, u))
        else:
            out += _inner_branches(k, u)
    return out


def has_branch(node: Node) -> bool:
    return any(f in BRANCHING for f in iter_funcs(node))


# ---------------------------------------------------------------- 動く参照・計算のたびに変わる関数
def _name_funcs(node: Node, wb, sheet: str, depth: int = 0) -> set[str]:
    """数式と、数式が使う名前の中の関数。"""
    out = set(iter_funcs(node))
    if depth > 8 or not hasattr(wb, "defined_name"):
        return out
    for ref in iter_refs(node):
        if isinstance(ref, NameNode):
            d = wb.defined_name(ref.name, ref.sheet or sheet)
            if d is not None and d.ast is not None:
                out |= _name_funcs(d.ast, wb, d.sheet_scope or sheet, depth + 1)
    return out


def dynamic_refs(node: Node, u: Rule) -> list[dict]:
    """INDIRECT・OFFSET：静的に分かる所と分からない所。"""
    out = []
    stack = [node]
    while stack:
        n = stack.pop()
        if isinstance(n, Func):
            if n.name == "INDIRECT":
                a = n.args
                if a and isinstance(a[0], Str) and len(a) == 1:
                    out.append({"function": "INDIRECT", "expression": expr(n, u), "resolved": a[0].value})
                else:
                    out.append({"function": "INDIRECT", "expression": expr(n, u),
                                "unresolved": {"ref_text": expr(a[0], u) if a else ""}})
            elif n.name == "OFFSET":
                a = n.args
                un = {}
                for i, what in ((1, "rows"), (2, "cols"), (3, "height"), (4, "width")):
                    if i < len(a) and not _is_literal(a[i]):
                        un[what] = expr(a[i], u)
                d = {"function": "OFFSET", "expression": expr(n, u), "base": expr(a[0], u) if a else ""}
                if un:
                    d["unresolved"] = un
                else:
                    d["literal_arguments"] = True
                out.append(d)
            elif n.name == "CELL" or n.name == "INFO":
                out.append({"function": n.name, "expression": expr(n, u), "unresolved": {"environment": "開いている環境によって変わる"}})
            stack.extend(n.args)
        elif isinstance(n, Binary):
            if n.op == ":" and not (isinstance(n.left, RefNode) and isinstance(n.right, RefNode)):
                out.append({"function": "range operator", "expression": expr(n, u),
                            "unresolved": {"range": "範囲の端が計算で決まる"}})
            stack.extend((n.left, n.right))
        elif isinstance(n, (Unary, Postfix)):
            stack.append(n.operand)
    return out


def _is_literal(n: Node) -> bool:
    from ..formula.parser import Missing, Num

    if isinstance(n, Unary):
        return _is_literal(n.operand)
    return isinstance(n, (Num, Missing))


# ---------------------------------------------------------------- まとめ
def analyse(rules: list[Rule], sheets: list[SheetScan], src, ix: RectIndex, external_targets: dict) -> dict:
    wb = src.meta
    by_name = {s.name: (i, s) for i, s in enumerate(sheets)}
    labels = Labels(sheets, src)

    # 読む所を列ごとに
    reads_ix = RectIndex()  # 読む長方形 → 規則
    read_cols: dict[str, dict[int, list]] = {}  # シート → 列 → [(r1, r2)]（参照表も含むすべての読み取り）
    plain_rects: dict[str, set] = {}  # 参照表でない読み取りの長方形
    table_reads: dict[tuple, dict] = {}
    abs_cells: set[tuple] = set()
    missing: list[dict] = []
    hidden_used: dict[tuple, dict] = {}
    flow: dict[tuple, set] = {}

    for u in rules:
        for rd in u.reads:
            if rd.missing:
                missing.append({"rule": u.id, "reference": rd.pattern, "sheet": rd.sheet})
                continue
            if rd.external:
                continue
            if rd.sheet not in by_name:
                continue
            si, sh = by_name[rd.sheet]
            if rd.sheet != u.sheet:
                flow.setdefault((rd.sheet, u.sheet), set()).add(u.id)
            for rect in read_rects(u, rd):
                reads_ix.add(rd.sheet, *rect, u.id)
                cols = read_cols.setdefault(rd.sheet, {})
                for c in _clip_cols(rect, sh.max_col):
                    cols.setdefault(c, []).append((rect[0], rect[2]))
                if not rd.lookup:
                    plain_rects.setdefault(rd.sheet, set()).add(rect)
                if rd.lookup:
                    t = table_reads.setdefault((rd.sheet, rect), {"functions": set(), "used_by": set()})
                    t["functions"] |= rd.lookup
                    t["used_by"].add(u.id)
                elif rd.absolute and rect[0] == rect[2] and rect[1] == rect[3]:
                    abs_cells.add((rd.sheet, rect[0], rect[1]))
                _hidden(hidden_used, sh, wb.sheets[si] if si < len(wb.sheets) else None, rect, u.id)
    reads_ix.freeze()

    formula_iv: dict[str, dict[int, list]] = {}
    for u in rules:
        fcols = formula_iv.setdefault(u.sheet, {})
        for m in u.members:
            for c in range(m.c1, m.c2 + 1):
                fcols.setdefault(c, []).append((m.r1, m.r2))
    for fcols in formula_iv.values():
        for c, iv in fcols.items():
            fcols[c] = _merge(iv)

    # 入力：規則が読み、何も計算していないセル
    inputs, params = [], []
    dv_index = {s.name: _validations(wb, by_name, s.name) for s in sheets}
    links = _control_links(wb)
    for name, rects in plain_rects.items():
        si, sh = by_name[name]
        for rect in _input_rects(rects, formula_iv.get(name, {}), sh.max_col):
            for one in _split_for_params(rect, name, abs_cells, sh):
                r1, c1, r2, c2 = one
                n = (r2 - r1 + 1) * (c2 - c1 + 1)
                filled = sum(sh.const.count(c, r1, r2) for c in range(c1, c2 + 1))
                who = sorted({h[4] for h in reads_ix.hits(name, r1, c1, r2, c2)}, key=lambda x: int(x[1:]))
                labels.plan(si, r1, c1)
                d = {"sheet": name, "range": rect_text(r1, c1, r2, c2), "cells": n, "filled": filled,
                     "read_by": who[:LISTED], **({"read_by_n": len(who)} if len(who) > LISTED else {}),
                     "_at": (si, r1, c1)}
                dv = [_dv_fact(v) for v in dv_index[name] if _dv_hits(v, one)]
                if dv:
                    d["constraints"] = dv
                ctl = [x for k, x in links.get(name, {}).items() if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]
                if ctl:
                    d["controls"] = ctl
                if _rect_hidden(sh, wb.sheets[si] if si < len(wb.sheets) else None, one):
                    d["hidden"] = True
                if n == 1 and (name, r1, c1) in abs_cells and filled:
                    params.append(d)
                else:
                    inputs.append(d)

    # 値を取りに行くセル
    want: dict[int, set] = {}
    for d in inputs + params:
        if d["cells"] == 1 and d["filled"]:
            want.setdefault(d["_at"][0], set()).add(d["_at"][1:])
    got = {si: src.fetch(si, cells) for si, cells in want.items()}
    for d in inputs + params:
        if d["cells"] == 1 and d["filled"]:
            d["value"] = plain(got.get(d["_at"][0], {}).get(d["_at"][1:]))

    # 参照表
    tables = []
    tables_ids: dict[tuple, str] = {}
    for (name, rect), info in sorted(table_reads.items(), key=lambda kv: (by_name[kv[0][0]][0], kv[0][1])):
        si, sh = by_name[name]
        tables.append(_table_entry(name, si, sh, rect, info, src, ix, tables_ids))

    # 出力：読まれない規則のセル
    for name, cols in read_cols.items():
        for c, iv in cols.items():
            cols[c] = _merge(iv)
    for u in rules:
        cols = read_cols.get(u.sheet, {})
        by_col: dict[int, list] = {}
        for m in u.members:
            for c in range(m.c1, m.c2 + 1):
                left = _minus([(m.r1, m.r2)], cols.get(c, []))
                if left:
                    by_col.setdefault(c, []).extend(left)
        u.__dict__["outputs"] = _to_rects({c: _merge(iv) for c, iv in by_col.items()})
        labels.plan(u.si, *u.origin)

    labels.resolve()
    for d in inputs + params:
        at = d.pop("_at")
        f = labels.facts(*at)
        if f:
            d["labels"] = f

    # 計算のたびに変わる値（読む規則まで）
    volatile_rules = set()
    for u in rules:
        if u.shape.ast is None:
            continue
        funcs = _name_funcs(u.shape.ast, wb, u.sheet)
        u.__dict__["funcs"] = funcs
        v = sorted(funcs & VOLATILE)
        if v:
            u.__dict__["volatile"] = v
            volatile_rules.add(u.id)
    by_id = {u.id: u for u in rules}
    todo = list(volatile_rules)
    tainted = set(todo)
    while todo:
        for w in by_id[todo.pop()].used_by:
            if w not in tainted:
                tainted.add(w)
                todo.append(w)
    for w in tainted - volatile_rules:
        by_id[w].__dict__["depends_on_volatile"] = True

    unknown = {}
    for u in rules:
        funcs = u.__dict__.get("funcs", ())
        bad = sorted(f for f in funcs if f not in EAGER and f not in LAZY)
        if bad:
            u.__dict__["unknown_functions"] = bad
            for f in bad:
                unknown.setdefault(f, []).append(u.id)

    externals = []
    for u in rules:
        ext = [rd for rd in u.reads if rd.external]
        if ext:
            for rd in ext:
                idx = rd.sheet[1:rd.sheet.index("]")] if rd.sheet.startswith("[") else ""
                externals.append({"rule": u.id, "reference": rd.pattern, "book": external_targets.get(idx), "sheet": rd.sheet})

    return {"inputs": inputs, "parameters": params, "lookup_tables": tables,
            "missing_sheets": missing, "hidden_used": _hidden_list(hidden_used),
            "cross_sheet_flow": [{"from": a, "to": b, "rules": sorted(w, key=lambda x: int(x[1:]))[:LISTED], "rule_count": len(w)}
                                 for (a, b), w in sorted(flow.items())],
            "external_references": externals, "unknown_functions": {k: v for k, v in sorted(unknown.items())},
            "labels": labels}


def _input_rects(rects: set, formula_cols: dict, max_col: int) -> list[tuple[int, int, int, int]]:
    """読む長方形から、数式のセルを除いた所。重なる読み取りだけを 1 つにする（隣り合うだけの欄は、別の入力のまま）。"""
    items = sorted(rects)
    parent = list(range(len(items)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    ix = RectIndex()
    for i, x in enumerate(items):
        ix.add("", *x, i)
    ix.freeze()
    for i, x in enumerate(items):
        for h in ix.hits("", *x):
            a, b = find(i), find(h[4])
            if a != b:
                parent[a] = b
    comps: dict[int, list] = {}
    for i, x in enumerate(items):
        comps.setdefault(find(i), []).append(x)
    out = []
    for group in comps.values():
        by_col: dict[int, list] = {}
        for x in group:
            for c in _clip_cols(x, max_col):
                by_col.setdefault(c, []).append((x[0], x[2]))
        left = {}
        for c, iv in by_col.items():
            got = _minus(_merge(iv), formula_cols.get(c, []))
            if got:
                left[c] = got
        out += _to_rects(left)
    out.sort()
    return out


def _split_for_params(rect, name, abs_cells, sh):
    """絶対参照で読まれる値のあるセルは、1 つずつ（パラメータとして）分ける。"""
    r1, c1, r2, c2 = rect
    if (r1, c1) == (r2, c2):
        return [rect]
    hits = [(r, c) for s, r, c in abs_cells if s == name and r1 <= r <= r2 and c1 <= c <= c2 and sh.const.has(r, c)] if abs_cells else []
    if not hits:
        return [rect]
    hits_set = set(hits)
    out = [(r, c, r, c) for r, c in sorted(hits)]
    by_col: dict[int, list] = {}
    for c in range(c1, c2 + 1):
        rows = sorted(r for r, cc in hits_set if cc == c)
        iv = _minus([(r1, r2)], [(r, r) for r in rows])
        if iv:
            by_col[c] = iv
    out += _to_rects(by_col)
    out.sort()
    return out


def _hidden(acc: dict, sh: SheetScan, meta, rect, uid: str) -> None:
    if sh.state != "visible":
        d = acc.setdefault((sh.name, "sheet", sh.state), {"sheet": sh.name, "what": "sheet", "state": sh.state, "used_by": set()})
        d["used_by"].add(uid)
    h = sh.hidden_rows
    for i in range(0, len(h), 2):
        a, b = max(h[i], rect[0]), min(h[i + 1], rect[2])
        if a <= b:
            d = acc.setdefault((sh.name, "rows", h[i], h[i + 1]), {"sheet": sh.name, "what": "rows", "rows": f"{h[i]}:{h[i + 1]}", "used_by": set()})
            d["used_by"].add(uid)
    if meta is not None:
        for c in sorted(getattr(meta, "hidden_cols", ())):
            if rect[1] <= c <= rect[3]:
                from ..formula.parser import num_to_col

                d = acc.setdefault((sh.name, "col", c), {"sheet": sh.name, "what": "column", "column": num_to_col(c), "used_by": set()})
                d["used_by"].add(uid)


def _rect_hidden(sh: SheetScan, meta, rect) -> bool:
    if sh.state != "visible":
        return True
    h = sh.hidden_rows
    if any(max(h[i], rect[0]) <= min(h[i + 1], rect[2]) for i in range(0, len(h), 2)):
        return True
    return meta is not None and any(rect[1] <= c <= rect[3] for c in getattr(meta, "hidden_cols", ()))


def _hidden_list(acc: dict) -> list[dict]:
    out = []
    for d in acc.values():
        w = sorted(d["used_by"], key=lambda x: int(x[1:]))
        out.append({**{k: v for k, v in d.items() if k != "used_by"}, "used_by": w[:LISTED], "used_by_n": len(w)})
    return out


def _validations(wb, by_name, name) -> list:
    si = by_name[name][0]
    return list(getattr(wb.sheets[si], "validations", ())) if si < len(wb.sheets) else []


def _dv_hits(v, rect) -> bool:
    return any(a[0] <= rect[2] and a[2] >= rect[0] and a[1] <= rect[3] and a[3] >= rect[1] for a in v.sqref)


def _dv_fact(v) -> dict:
    d = {"type": v.type}
    for k in ("operator", "formula1", "formula2", "prompt", "error"):
        x = getattr(v, k, None)
        if x:
            d[k] = x
    d["blocks"] = v.blocks
    return d


def _control_links(wb) -> dict:
    out: dict[str, dict] = {}
    for s in wb.sheets:
        for c in getattr(s, "controls", ()):
            lc = getattr(c, "link_cell", None)
            if lc:
                out.setdefault(lc[0], {})[(lc[1], lc[2])] = {"kind": c.kind, "name": c.name, "text": c.text}
    return out


def _table_entry(name, si, sh, rect, info, src, ix, ids) -> dict:
    r1, c1, r2, c2 = rect
    rows, cols = r2 - r1 + 1, c2 - c1 + 1
    computed = sorted({h[4] for h in ix.hits(name, *rect)}, key=lambda x: int(x[1:]))
    d = {"sheet": name, "range": rect_text(*rect), "rows": rows, "cols": cols,
         "functions": sorted(info["functions"]), "used_by": sorted(info["used_by"], key=lambda x: int(x[1:]))[:LISTED],
         "used_by_n": len(info["used_by"])}
    if computed:
        d["computed_by"] = computed[:LISTED]
    hdr = None
    for row in (r1 - 1, r1):
        for c in range(c1, c2 + 1):
            if row >= 1 and sh.text.has(row, c):
                hdr = (row, c)
                break
        if hdr:
            break
    show_rows = rows if rows * cols <= TABLE_INLINE_CELLS else min(TABLE_PREVIEW_ROWS, rows)
    cells = {(r, c) for r in range(r1, r1 + show_rows) for c in range(c1, c2 + 1)}
    if hdr:
        cells.add(hdr)
    vals = src.fetch(si, cells)
    d["values" if show_rows == rows else "first_rows"] = [[plain(vals.get((r, c))) for c in range(c1, c2 + 1)]
                                                           for r in range(r1, r1 + show_rows)]
    if hdr and isinstance(vals.get(hdr), str):
        d["read_table"] = {"sheet": name, "header": vals[hdr]}
    else:
        d["read_table"] = None
        d["note"] = "表の上に字のセルがないので、read_table では引けません。read_window に sheet と rows を渡して読めます"
    return d


def read_json(u: Rule, rd: Read, ix: RectIndex) -> dict:
    d: dict = {"pattern": rd.pattern, "sheet": rd.sheet}
    if rd.external:
        d["external"] = True
        return d
    if rd.missing:
        d["missing_sheet"] = True
        return d
    rects = read_rects(u, rd)
    if rects:
        d["range"] = rect_text(*rects[0]) if len(rects) == 1 else [rect_text(*x) for x in rects[:LISTED]]
    if rd.name:
        d["name"] = rd.name
    if rd.absolute:
        d["absolute"] = True
    if rd.lookup:
        d["lookup"] = sorted(rd.lookup)
    return d


def rule_json(u: Rule, ix: RectIndex, labels: Labels, wb) -> dict:
    head = u.range
    if u.kind in ("array", "dynamic_array", "array_cell"):
        body = "{=" + u.pattern_body() + "}"
    elif u.kind == "data_table":
        body = "データテーブル（What-If 分析）"
    else:
        body = u.pattern_body()
    d = {"id": u.id, "sheet": u.sheet, "range": u.range, "kind": u.kind, "cells": u.count,
         "text": f"{head} = {body}", "formula": "=" + u.text if u.text and u.kind != "data_table" else None, "a1": "=" + u.pattern_body() if u.shape.ast is not None else None,
         "r1c1": "=" + u.shape.r1c1 if u.shape.ast is not None else None}
    if u.shape.error:
        d["unreadable"] = u.shape.error
    if u.note and u.kind == "data_table":
        d["definition"] = {k: v for k, v in u.note.items() if v}
    elif u.note and u.note.get("ref"):
        d["array_ref"] = u.note["ref"]
        if u.note.get("too_wide"):
            d["note"] = "配列数式の宣言された範囲が広すぎるので、親のセルだけを規則にしました"
    reads = [read_json(u, rd, ix) for rd in u.reads]
    if reads:
        d["reads"] = reads
    if u.uses:
        d["uses_rules"] = sorted(u.uses, key=lambda x: int(x[1:]))
    if u.used_by:
        d["read_by_rules"] = sorted(u.used_by, key=lambda x: int(x[1:]))[:LISTED]
        if len(u.used_by) > LISTED:
            d["read_by_rules_n"] = len(u.used_by)
    d["level"] = u.level
    outs = u.__dict__.get("outputs") or []
    if outs:
        d["outputs"] = [rect_text(*x) for x in outs[:LISTED]]
        if len(outs) > LISTED:
            d["outputs_n"] = len(outs)
    if u.exception_count:
        d["exceptions"] = u.exceptions
        d["exception_count"] = u.exception_count
    f = labels.facts(u.si, *u.origin)
    if f:
        d["labels"] = f
    ast = u.shape.ast
    if ast is not None:
        if has_branch(ast):
            d["branches"] = branches(ast, u)
        dyn = dynamic_refs(ast, u)
        if dyn:
            d["dynamic_refs"] = dyn
    for k in ("volatile", "depends_on_volatile", "unknown_functions"):
        v = u.__dict__.get(k)
        if v:
            d[k] = v
    if u.in_cycle:
        d["cycle"] = u.in_cycle
    return {k: v for k, v in d.items() if v is not None}


def names_json(wb, rules: list[Rule], sheets: list[SheetScan]) -> list[dict]:
    used: dict[str, set] = {}
    for u in rules:
        for rd in u.reads:
            if rd.name:
                used.setdefault(rd.name.lower(), set()).add(u.id)
        if u.shape.ast is not None:
            for ref in iter_refs(u.shape.ast):
                if isinstance(ref, NameNode):
                    used.setdefault(ref.name.lower(), set()).add(u.id)
    out = []
    for d in getattr(wb, "names", ()):
        rec = {"name": d.name, "refers_to": d.text}
        if d.sheet_scope:
            rec["scope"] = d.sheet_scope
        if d.hidden:
            rec["hidden"] = True
        if d.error:
            rec["unreadable"] = d.error
        w = used.get(d.name.lower())
        if w:
            rec["used_by"] = sorted(w, key=lambda x: int(x[1:]))[:LISTED]
        out.append(rec)
    return out


def key_of(sheet: str, r: int, c: int) -> str:
    return key_text(sheet, r, c)
