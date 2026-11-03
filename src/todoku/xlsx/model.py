"""Excel の様式を「データ」と「処理」に分ける：入力欄（データ）と、入力から結果を出す数式（処理）。

- 処理の一覧：入力欄に（数式をたどって）つながる数式のセル。見出し（窓と同じ決め方）と、番地を見出しに置き換えた式を添える。
  入力からの依存は、数式の参照（セル・範囲・名前・テーブル）を静的にたどる。INDIRECT・OFFSET のように参照先が計算で決まる
  数式は、静的にはたどれないので印を付ける（試算では、計算エンジンが実際の参照先で計算する）。
- 試算：入力の値の組から、ファイルを書かずに結果を計算する（書き込みと同じ受け取り方・入力規則・計算エンジン）。
"""

from __future__ import annotations

import bisect
from collections import defaultdict

from ..formula.numfmt import format_value
from ..formula.parser import NameNode, RefNode, StructRef, addr, iter_funcs, iter_refs, parse, parse_cell, tokenize
from ..formula.values import ExcelError, is_number
from .calc import _run_deep, recalc
from .workbook import Workbook

DYNAMIC = {"INDIRECT", "OFFSET"}


def _key(sheet: str, r: int, c: int) -> str:
    return f"{sheet}!{addr(r, c)}"


class _Index:
    """シートごとの、入力欄と数式のセル（範囲の参照に含まれるものを速く探す）。"""

    def __init__(self, cells: set[tuple[str, int, int]]):
        self.by_sheet: dict[str, list[tuple[int, int]]] = defaultdict(list)
        for s, r, c in cells:
            self.by_sheet[s].append((r, c))
        for v in self.by_sheet.values():
            v.sort()

    def within(self, sheet: str, r1: int, c1: int, r2: int, c2: int):
        rows = self.by_sheet.get(sheet, [])
        i = bisect.bisect_left(rows, (r1, -1))
        while i < len(rows) and rows[i][0] <= r2:
            r, c = rows[i]
            if c1 <= c <= c2:
                yield (sheet, r, c)
            i += 1


def _ref_cells(wb: Workbook, node, sheet: str, idx: _Index, depth: int = 0):
    """参照が指す（入力欄か数式の）セル。"""
    if isinstance(node, RefNode):
        sh = node.sheet or sheet
        if wb.sheet(sh) is None:
            return
        r1, c1, r2, c2 = node.r1, node.c1, node.r2, node.c2
        if node.kind == "cols":
            r1, r2 = 1, 1 << 20
        elif node.kind == "rows":
            c1, c2 = 1, 1 << 14
        yield from idx.within(sh, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
    elif isinstance(node, NameNode) and depth < 8:
        d = next((x for x in wb.names if x.name.lower() == node.name.lower()
                  and (x.sheet_scope in (None, sheet) or node.sheet)), None)
        if d is not None and d.text:
            try:
                ast = parse(d.text)
            except Exception:  # noqa: BLE001
                return
            for ref in iter_refs(ast):
                yield from _ref_cells(wb, ref, sheet, idx, depth + 1)
    elif isinstance(node, StructRef):
        # テーブル全体を、使う範囲とみなす（列の指定は、静的には広めにとる）
        t = (getattr(wb, "tables", None) or {}).get((node.table or "").lower())
        if t is not None:
            yield from idx.within(t.sheet, t.r1, t.c1, t.r2, t.c2)


def model(wb: Workbook, doc: dict) -> dict:
    """{inputs: 入力欄, steps: 入力につながる数式（処理）, results: その終わりの結果}。doc は窓（xlsx/window.build）。"""
    from .window import Layout

    inputs = {}
    for sh in doc["sheets"]:
        for f in sh.get("fields", []):
            if f.get("cell") and f.get("kind") != "inline" and not f.get("locked"):
                r, c = parse_cell(f["cell"].split(":")[0])
                inputs[(sh["name"], r, c)] = f
    formulas = {(s.name, r, c): cell for s in wb.sheets for (r, c), cell in s.cells.items() if cell.has_formula}
    idx = _Index(set(inputs) | set(formulas))
    uses: dict[tuple, set] = {}
    dynamic = set()
    for k, cell in formulas.items():
        if cell.ast is None:
            continue
        uses[k] = {x for ref in iter_refs(cell.ast) for x in _ref_cells(wb, ref, k[0], idx)}
        if any(f in DYNAMIC for f in iter_funcs(cell.ast)):
            dynamic.add(k)
    users = defaultdict(set)
    for k, us in uses.items():
        for u in us:
            users[u].add(k)
    # 入力につながる数式（入力から、使う側へたどる）
    seen = set(inputs)
    todo = list(inputs)
    while todo:
        for u in users.get(todo.pop(), ()):
            if u not in seen:
                seen.add(u)
                todo.append(u)
    steps_keys = sorted((k for k in seen if k in formulas), key=lambda k: ([s.name for s in wb.sheets].index(k[0]), k[1], k[2]))
    # 見出し（窓の入力欄と同じ決め方で、数式のセルにも）
    labels: dict[tuple, str] = {k: f.get("label") or "" for k, f in inputs.items()}
    lays = {}
    for k in steps_keys:
        sh = wb.sheet(k[0])
        if k[0] not in lays:
            lays[k[0]] = Layout(wb, sh)
        lay = lays[k[0]]
        box = next((b for b in lay.boxes(k[1]) if b.c1 <= k[2] <= b.c2 and b.kind == "formula"), None)
        if box is not None:
            try:
                labels[k] = lay.describe(box).get("label") or ""
            except Exception:  # noqa: BLE001
                labels[k] = ""
    def name_of(k) -> str:
        lab = labels.get(k)
        return f"「{lab}」" if lab else _key(*k)

    steps = []
    used_by_steps = set()
    for k in steps_keys:
        cell = formulas[k]
        used_by_steps |= uses.get(k, set())
        steps.append({"cell": _key(*k), "label": labels.get(k) or None, "formula": "=" + cell.formula,
                      "in_words": "＝" + _in_words(wb, cell.formula, k[0], idx, name_of),
                      "value": _show(wb, cell, cell.value),
                      "from_inputs": sorted(_key(*i) for i in uses.get(k, set()) if i in inputs),
                      "from_steps": sorted(_key(*i) for i in uses.get(k, set()) if i in formulas and i in seen),
                      **({"dynamic": "参照先が計算で決まる関数（INDIRECT・OFFSET）を含む。依存は静的にはたどれない"} if k in dynamic else {})})
    results = [s for s, k in zip(steps, steps_keys) if k not in used_by_steps]
    return {"inputs": [{"id": f["id"], "label": f.get("label"), "kind": f.get("kind"), "value": f.get("value"),
                        **({"options": f["options"]} if f.get("options") else {})} for f in inputs.values()],
            "steps": steps, "results": [{"cell": s["cell"], "label": s["label"], "value": s["value"]} for s in results],
            "dynamic": len(dynamic & set(steps_keys))}


def _in_words(wb: Workbook, formula: str, sheet: str, idx: _Index, name_of) -> str:
    """数式の番地を、見出しに置き換える（1つのセルを指す参照だけ。範囲は「番地の範囲」のまま）。"""
    try:
        toks = tokenize(formula)
    except Exception:  # noqa: BLE001
        return formula
    out = []
    for i, t in enumerate(toks):
        end = toks[i + 1].pos if i + 1 < len(toks) else len(formula)
        text = formula[t.pos:end]
        if t.kind == "ref" and isinstance(t.value, RefNode) and t.value.kind == "cell":
            k = (t.value.sheet or sheet, t.value.r1, t.value.c1)
            trail = text[len(text.rstrip()):]
            out.append(name_of(k) + trail)
        else:
            out.append(text)
    return "".join(out)


def _show(wb: Workbook, cell, v):
    if isinstance(v, ExcelError):
        return v.code
    if is_number(v):
        try:
            return format_value(v, wb.style(cell.style).num_fmt, wb.date1904)
        except Exception:  # noqa: BLE001
            return v
    return v


def calculate(wb: Workbook, doc: dict, values: dict) -> dict:
    """値の組（{欄の id: 値}）から、ファイルを書かずに結果を計算する。書き込みと同じ受け取り方と入力規則で。"""
    from .edit import EditError, plan_inputs, rule_issues
    from .calc import Calc

    plan_ = plan_inputs(wb, values)
    targets = plan_.targets

    def check():
        calc = Calc(wb, targets)
        return [i for (s, r, c), v in targets.items() for i in rule_issues(wb, s, r, c, v, calc)]

    issues = _run_deep(check)
    blocking = [i.message for i in issues if i.blocking]
    if blocking:
        raise EditError("入力規則に合いません（Excelでも受け付けない値です）:\n" + "\n".join(blocking))
    changed, problems = recalc(wb, targets)
    m = model(wb, doc)
    lab = {s["cell"]: s["label"] for s in m["steps"]}
    res_cells = {r["cell"] for r in m["results"]}
    out = []
    for (s, r, c), after in sorted(changed.items()):
        cell = wb.sheet(s).cells.get((r, c))
        k = _key(s, r, c)
        out.append({"cell": k, "label": lab.get(k), "before": _show(wb, cell, cell.value), "after": _show(wb, cell, after),
                    "result": k in res_cells})
    return {"inputs": {_key(*k): v for k, v in targets.items()}, "changed": out,
            "not_computed": [{"cell": _key(*k), "why": w} for k, w in problems.items()],
            "warnings": list(plan_.warnings) + [i.message for i in issues if not i.blocking]}


def summary(m: dict, limit: int = 40) -> list[dict]:
    """結果（ほかの数式に使われない数式）を、シートと見出しでまとめる（同じ形の行は1つに）。"""
    groups: dict[tuple, dict] = {}
    steps = {s["cell"]: s for s in m["steps"]}
    for r in m["results"]:
        sheet = r["cell"].split("!")[0]
        g = groups.setdefault((sheet, r["label"] or ""), {"sheet": sheet, "label": r["label"], "cells": [], "in_words": None})
        g["cells"].append(r["cell"])
        if g["in_words"] is None:
            g["in_words"] = steps[r["cell"]]["in_words"]
    out = sorted(groups.values(), key=lambda g: -len(g["cells"]))
    for g in out:
        g["count"] = len(g["cells"])
        g["cells"] = g["cells"][:5]
    return out[:limit]


def explain(m: dict, cell: str, depth: int = 6) -> dict:
    """1つのセルの計算を、入力欄までさかのぼる（見出しで書いた式の木）。"""
    steps = {s["cell"]: s for s in m["steps"]}
    inputs = {i["id"]: i for i in m["inputs"]}

    def walk(k: str, d: int, seen: set) -> dict:
        s = steps.get(k)
        if s is None:
            i = inputs.get(k)
            return {"cell": k, "input": True, "label": i.get("label") if i else None, "value": i.get("value") if i else None}
        node = {"cell": k, "label": s["label"], "in_words": s["in_words"], "formula": s["formula"], "value": s["value"],
                "inputs": s["from_inputs"]}
        if s.get("dynamic"):
            node["dynamic"] = s["dynamic"]
        if d > 0 and k not in seen:
            node["from"] = [walk(x, d - 1, seen | {k}) for x in s["from_steps"][:12]]
        return node

    if cell not in steps:
        raise ValueError(f"{cell} は、入力欄につながる数式のセルではありません")
    return walk(cell, depth, set())
