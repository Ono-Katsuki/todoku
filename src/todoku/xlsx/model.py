"""Excel の様式を「データ」と「処理」に分ける：入力欄（データ）と、入力から結果を出す数式（処理）。

- 処理の一覧：入力欄に（数式をたどって）つながる数式のセル。見出し（窓と同じ決め方）と、番地を見出しに置き換えた式を添える。
  入力からの依存は、数式の参照（セル・範囲・名前・テーブル）を静的にたどる。INDIRECT・OFFSET のように参照先が計算で決まる
  数式は、静的にはたどれないので印を付ける（試算では、計算エンジンが実際の参照先で計算する）。
- 試算：入力の値の組から、ファイルを書かずに結果を計算する（書き込みと同じ受け取り方・入力規則・計算エンジン）。
"""

from __future__ import annotations

import bisect
from collections import defaultdict

from ..formula.numfmt import LocaleDependent, format_value
from ..formula.parser import NameNode, RefNode, StructRef, addr, iter_funcs, iter_refs, parse, parse_cell, tokenize
from ..formula.values import ExcelError, Unsupported, is_number
from .calc import _run_deep, recalc
from .workbook import Workbook

DYNAMIC = {"INDIRECT", "OFFSET"}


def _key(sheet: str, r: int, c: int) -> str:
    return f"{sheet}!{addr(r, c)}"


class _Index:
    """入力欄と数式のセルに、シート・列・行の順の番号（位置）をふる。参照は、位置の区間のまま持つ
    （積み上げの合計 SUM(A$3:Ai) のように範囲が重なる数式が並んでも、マスごとの控えを作らない）。"""

    def __init__(self, cells: set[tuple[str, int, int]]):
        cols: dict[tuple[str, int], list[int]] = defaultdict(list)
        for s, r, c in cells:
            cols[(s, c)].append(r)
        self.keys: list[tuple[str, int, int]] = []
        self.cols: dict[str, tuple[list[int], list[tuple[list[int], int]]]] = {}  # シート → (列の並び, 列ごとの (行の並び, 最初の位置))
        by_sheet: dict[str, list[int]] = defaultdict(list)
        for s, c in cols:
            by_sheet[s].append(c)
        for s, cs in by_sheet.items():
            cs.sort()
            rows = []
            for c in cs:
                rs = sorted(cols[(s, c)])
                rows.append((rs, len(self.keys)))
                self.keys.extend((s, r, c) for r in rs)
            self.cols[s] = (cs, rows)
        self.pos = {k: i for i, k in enumerate(self.keys)}
        self._memo: dict[tuple, list[tuple[int, int]]] = {}

    def spans(self, sheet: str, r1: int, c1: int, r2: int, c2: int) -> list[tuple[int, int]]:
        """範囲の中のセルの位置の区間 [a, b)。同じ範囲は 1 度だけ計算する（隣の列まで続く区間は 1 つにつなぐ）。"""
        got = self._memo.get((sheet, r1, c1, r2, c2))
        if got is not None:
            return got
        out: list[tuple[int, int]] = []
        one = self.cols.get(sheet)
        if one is not None:
            cs, rows = one
            for i in range(bisect.bisect_left(cs, c1), bisect.bisect_right(cs, c2)):
                rs, base = rows[i]
                a, b = bisect.bisect_left(rs, r1) + base, bisect.bisect_right(rs, r2) + base
                if a >= b:
                    continue
                if out and out[-1][1] == a:
                    out[-1] = (out[-1][0], b)
                else:
                    out.append((a, b))
        self._memo[(sheet, r1, c1, r2, c2)] = out
        return out


def _ref_spans(wb: Workbook, node, sheet: str, idx: _Index, depth: int = 0):
    """参照が指す（入力欄か数式の）セルの、位置の区間。"""
    if isinstance(node, RefNode):
        sh = node.sheet or sheet
        if not wb.has_sheet(sh):
            return  # ないシートへの参照（#REF!）：つながるセルはない
        r1, c1, r2, c2 = node.r1, node.c1, node.r2, node.c2
        if node.kind == "cols":
            r1, r2 = 1, 1 << 20
        elif node.kind == "rows":
            c1, c2 = 1, 1 << 14
        yield from idx.spans(sh, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))
    elif isinstance(node, NameNode) and depth < 8:
        d = next((x for x in wb.names if x.name.lower() == node.name.lower()
                  and (x.sheet_scope in (None, sheet) or node.sheet)), None)
        if d is not None and d.text:
            try:
                ast = parse(d.text)
            except Exception:  # noqa: BLE001
                return
            for ref in iter_refs(ast):
                yield from _ref_spans(wb, ref, sheet, idx, depth + 1)
    elif isinstance(node, StructRef):
        # テーブル全体を、使う範囲とみなす（列の指定は、静的には広めにとる）
        t = (getattr(wb, "tables", None) or {}).get((node.table or "").lower())
        if t is not None:
            yield from idx.spans(t.sheet, t.r1, t.c1, t.r2, t.c2)


def _merged(spans: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """区間の和（重なりをなくし、位置の順に）。"""
    out: list[tuple[int, int]] = []
    for a, b in sorted(spans):
        if out and a <= out[-1][1]:
            if b > out[-1][1]:
                out[-1] = (out[-1][0], b)
        else:
            out.append((a, b))
    return out


class _Stabber:
    """区間を区間木に入れ、ある位置を含む区間の持ち主を、1 度ずつ取り出す（取り出した分は消える。全部で区間の数 × log）。"""

    def __init__(self, size: int):
        self.n = max(size, 1)
        self.nodes: dict[int, list[int]] = {}

    def add(self, a: int, b: int, who: int) -> None:
        a += self.n
        b += self.n
        while a < b:
            if a & 1:
                self.nodes.setdefault(a, []).append(who)
                a += 1
            if b & 1:
                b -= 1
                self.nodes.setdefault(b, []).append(who)
            a >>= 1
            b >>= 1

    def take(self, p: int) -> list[int]:
        out: list[int] = []
        i = p + self.n
        while i:
            got = self.nodes.pop(i, None)
            if got:
                out += got
            i >>= 1
        return out


LISTED = 64  # 1 つの処理について、つながる入力欄・処理を番地で示す数（数は n にすべて数える。積み上げの合計で 2 乗にならないように）


def _listed(spans: list[tuple[int, int]], next_at: list[int], name, total_upto: list[int]) -> tuple[list[str], int]:
    """区間の中の、印のある位置の名前（LISTED まで。字の順）と、印のある位置の数。next_at：その位置以降の次の印の位置。"""
    total = sum(total_upto[b] - total_upto[a] for a, b in spans)
    out: list[str] = []
    for a, b in spans:
        i = next_at[a]
        while i < b and len(out) < LISTED:
            out.append(name(i))
            i = next_at[i + 1]
        if len(out) >= LISTED:
            break
    out.sort()
    return out, total


def _marks(flags: bytearray) -> tuple[list[int], list[int]]:
    """印の並び → (その位置以降の次の印の位置, その位置より前の印の数)。"""
    n = len(flags)
    nxt = [n] * (n + 2)
    for i in range(n - 1, -1, -1):
        nxt[i] = i if flags[i] else nxt[i + 1]
    upto = [0] * (n + 1)
    for i in range(n):
        upto[i + 1] = upto[i] + flags[i]
    return nxt, upto


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
    uses: dict[tuple, list[tuple[int, int]]] = {}
    dynamic = set()
    users = _Stabber(len(idx.keys))
    for k, cell in formulas.items():
        if cell.ast is None:
            continue
        uses[k] = _merged([x for ref in iter_refs(cell.ast) for x in _ref_spans(wb, ref, k[0], idx)])
        me = idx.pos[k]
        for a, b in uses[k]:
            users.add(a, b, me)
        if any(f in DYNAMIC for f in iter_funcs(cell.ast)):
            dynamic.add(k)
    # 入力につながる数式（入力から、使う側へたどる）
    seen = bytearray(len(idx.keys))
    todo = [idx.pos[k] for k in inputs]
    for i in todo:
        seen[i] = 1
    while todo:
        for u in users.take(todo.pop()):
            if not seen[u]:
                seen[u] = 1
                todo.append(u)
    order = {s.name: i for i, s in enumerate(wb.sheets)}
    steps_keys = sorted((k for k in formulas if seen[idx.pos[k]]), key=lambda k: (order.get(k[0], -1), k[1], k[2]))
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

    is_input = bytearray(len(idx.keys))
    is_step = bytearray(len(idx.keys))
    for k in inputs:
        is_input[idx.pos[k]] = 1
    for k in steps_keys:
        is_step[idx.pos[k]] = 1
    in_next, in_upto = _marks(is_input)
    st_next, st_upto = _marks(is_step)
    names: dict[int, str] = {}

    def pos_name(i: int) -> str:
        got = names.get(i)
        if got is None:
            got = names[i] = _key(*idx.keys[i])
        return got

    steps = []
    used = [0] * (len(idx.keys) + 1)  # どこかの処理が使う位置（区間の出入りを数える）
    for k in steps_keys:
        cell = formulas[k]
        sp = uses.get(k, [])
        for a, b in sp:
            used[a] += 1
            used[b] -= 1
        from_in, n_in = _listed(sp, in_next, pos_name, in_upto)
        from_st, n_st = _listed(sp, st_next, pos_name, st_upto)
        steps.append({"cell": _key(*k), "label": labels.get(k) or None, "formula": "=" + cell.formula,
                      "in_words": "＝" + _in_words(wb, cell.formula, k[0], idx, name_of),
                      **_value_and_note(wb, cell),
                      "from_inputs": from_in, "from_steps": from_st,
                      **({"from_inputs_n": n_in} if n_in > len(from_in) else {}),
                      **({"from_steps_n": n_st} if n_st > len(from_st) else {}),
                      **({"dynamic": "参照先が計算で決まる関数（INDIRECT・OFFSET）を含む。依存は静的にはたどれない"} if k in dynamic else {})})
    depth = 0
    for i in range(len(idx.keys)):
        depth += used[i]
        used[i] = depth
    results = [s for s, k in zip(steps, steps_keys) if used[idx.pos[k]] == 0]
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


def _shown(wb: Workbook, cell, v) -> tuple:
    """(見えている字か値, 見え方について伝えること)。地域で見え方が変わる組み込みの表示形式（LocaleDependent）は、
    地域によらない書き方（ISO 8601）と、そう示す字。再現できない表示形式は、元の値と、再現できない理由。"""
    if isinstance(v, ExcelError):
        return v.code, None
    if is_number(v):
        fmt = wb.style(cell.style).num_fmt if cell is not None else "General"
        try:
            return format_value(v, fmt, wb.date1904), None
        except LocaleDependent as e:
            return e.shown, str(e)
        except Unsupported as e:
            return v, f"表示形式「{fmt}」は再現できません（{e}）。元の値を示します"
        except Exception as e:  # noqa: BLE001  壊れた表示形式
            return v, f"表示形式「{fmt}」を読めません（{type(e).__name__}）。元の値を示します"
    return v, None


def _show(wb: Workbook, cell, v):
    return _shown(wb, cell, v)[0]


def _value_and_note(wb: Workbook, cell) -> dict:
    v, note = _shown(wb, cell, cell.value)
    return {"value": v, **({"value_note": note} if note else {})}


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
        before, n1 = _shown(wb, cell, cell.value)
        shown_after, n2 = _shown(wb, cell, after)
        notes = list(dict.fromkeys(n for n in (n1, n2) if n))
        out.append({"cell": k, "label": lab.get(k), "before": before, "after": shown_after,
                    "result": k in res_cells, **({"note": "／".join(notes)} if notes else {})})
    return {"inputs": {_key(*k): v for k, v in targets.items()}, "changed": out,
            "not_computed": [{"cell": _key(*k), "why": w} for k, w in problems.items()],
            "warnings": list(plan_.warnings) + [i.message for i in issues if not i.blocking]}


def cell_values(wb: Workbook, values: dict, cells: list, store=None) -> list[dict]:
    """頼まれたセルの計算結果（ファイルは書かない）。動的配列のマスは、そのマスだけを計算する。集計が時間の中で終わらなければ、進んだ所を返す。"""
    from .calc import Calc
    from .edit import plan_inputs, resolve_target

    targets = plan_inputs(wb, values).targets if values else {}

    def run():
        calc = Calc(wb, targets, store)
        out = []
        for name in cells:
            s, r, c = resolve_target(wb, name)
            v = calc.cell_value(s, r, c)
            sh = wb.sheet(s)
            where = spill_of(calc, sh, r, c)
            if where is not None and sh.cells.get((r, c)) is None:
                v = where["value"]
            st, why = calc.status.get((s, r, c), ("ok", ""))
            if where is not None and where.get("master"):
                st, why = calc.status.get((s, *where["master"]), (st, why))
            d: dict = {"cell": _key(s, r, c)}
            un = calc.unfinished.get((s, r, c)) or (calc.unfinished.get((s, *where["master"])) if where and where.get("master") else None)
            if un is not None:
                d.update({"status": "unfinished", "why": why, "done": un.done, "total": un.total})
            elif st in ("ok", "volatile"):
                d.update({"status": "ok", "value": v.code if isinstance(v, ExcelError) else v})
            else:
                d.update({"status": st, "why": why})
            if where is not None and where.get("range"):
                d["spill"] = where["range"]
            note = calc.calc_notes.get((s, r, c))
            if note:
                d["note"] = note
            out.append(d)
        return out

    return _run_deep(run)


def spill_of(calc, sh, r: int, c: int) -> dict | None:
    """(r, c) に広がる配列（動的配列の親か、広がった先のマス）：{"master": 親, "range": 範囲, "value": このマスの値}。"""
    from ..formula.values import Arr

    cell = sh.cells.get((r, c))
    candidates = []
    if cell is not None and cell.has_formula and cell.array_master is None:
        candidates = [cell]
    elif cell is None or cell.array_master is not None:
        if cell is not None and cell.array_master is not None:
            candidates = [sh.cells[cell.array_master]]
        else:
            span = sh.array_span_at(r, c) if hasattr(sh, "array_span_at") else None
            candidates = [sh.cells[span]] if span else [m for (mr, mc), m in sh.cells.items()
                                                        if mr <= r and mc <= c and m.has_formula and m.array_master is None and getattr(m, "dynamic", False)]
    for m in candidates:
        calc.cell_value(m.sheet, m.row, m.col)
        arr = calc.memo.get(("__arr__", m.sheet, m.row, m.col))
        if not isinstance(arr, Arr):
            continue
        i, j = r - m.row, c - m.col
        if i < arr.rows and j < arr.cols:
            v = arr.at(i, j)
            return {"master": (m.row, m.col), "range": f"{addr(m.row, m.col)}:{addr(m.row + arr.rows - 1, m.col + arr.cols - 1)}",
                    "value": 0.0 if v is None else v}
    return None


def cells_for_job(path: str, password: str | None, values: dict, cells: list, store) -> list[dict]:
    """長い仕事の本体：ファイルを開いて、頼まれたセルを時間の上限なしに計算する。"""
    from .. import api

    wb = api.calc_workbook(path, password, list(values) + list(cells))
    return cell_values(wb, values, cells, store)


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
        if s.get("value_note"):
            node["value_note"] = s["value_note"]
        if s.get("dynamic"):
            node["dynamic"] = s["dynamic"]
        if d > 0 and k not in seen:
            node["from"] = [walk(x, d - 1, seen | {k}) for x in s["from_steps"][:12]]
        return node

    if cell not in steps:
        raise ValueError(f"{cell} は、入力欄につながる数式のセルではありません")
    return walk(cell, depth, set())
