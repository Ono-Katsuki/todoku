"""規則から、標準ライブラリだけで動く Python のモジュールを作る。

計算の決まりは vendor がとどくの計算エンジンのソースから切り出す（別の実装を持たない）。規則の数式は構文木のまま（平らな並びで）入れ、
エンジンの Evaluator がセルごとに、共有数式と同じやり方で（行・列のずれを渡して）計算する。
"""

from __future__ import annotations

import base64
import json
import zlib
from dataclasses import dataclass, field
from pathlib import Path

from ..formula.parser import ArrayLit, Binary, Bool, Err, Func, Missing, NameNode, Node, Num, Postfix, RefNode, Str, StructRef, Unary, iter_refs
from ..formula.values import ExcelError
from .rules import Rule, read_rects
from .scan import SheetScan
from .structure import _merge
from .vendor import Vendor, method_source, need_method

INLINE_DATA = 3000  # 入力の値をソースにそのまま書くセルの数。超えたら圧縮して 1 行に入れる
RESERVED = {"calculate", "RULES", "ORDER", "INPUTS", "PARAMETERS", "OUTPUTS", "UNREPRODUCIBLE", "VERIFICATION", "SOURCE", "main",
            "CellError", "NotReproducible", "excel_date", "json", "sys", "time", "threading", "re", "zlib", "base64", "datetime"}
DYNAMIC_FUNCS = {"INDIRECT", "OFFSET"}


@dataclass
class Generated:
    source: str
    facts: dict = field(default_factory=dict)
    placeholder: str = "\nVERIFICATION = None  # @@todoku-verification@@\n"  # 行の頭から（書類の字は改行を含まないので、ここだけに当たる）


def flat(node: Node) -> list:
    """構文木 → 後置の平らな並び（入れ子の深い式でも、Python の字句の上限に当たらない）。"""
    out: list = []
    stack: list = [node]
    while stack:
        n = stack.pop()
        if isinstance(n, tuple):  # 後で出す印
            out.append(n)
        elif isinstance(n, Num):
            out.append(("n", n.text))
        elif isinstance(n, Str):
            out.append(("s", n.value))
        elif isinstance(n, Bool):
            out.append(("b", n.value))
        elif isinstance(n, Err):
            out.append(("e", n.code))
        elif isinstance(n, Missing):
            out.append(("m",))
        elif isinstance(n, RefNode):
            flags = int(n.abs_r1) | int(n.abs_c1) << 1 | int(n.abs_r2) << 2 | int(n.abs_c2) << 3
            out.append(("R", n.sheet, n.r1, n.c1, n.r2, n.c2, flags, n.kind))
        elif isinstance(n, NameNode):
            out.append(("N", n.name, n.sheet))
        elif isinstance(n, StructRef):
            out.append(("S", n.table, tuple(n.specials), n.col1, n.col2))
        elif isinstance(n, Func):
            stack.append(("F", n.name, len(n.args)))
            stack.extend(reversed(n.args))
        elif isinstance(n, Unary):
            stack.append(("U", n.op))
            stack.append(n.operand)
        elif isinstance(n, Postfix):
            stack.append(("P", n.op))
            stack.append(n.operand)
        elif isinstance(n, Binary):
            stack.append(("B", n.op))
            stack.append(n.right)
            stack.append(n.left)
        elif isinstance(n, ArrayLit):
            stack.append(("A", len(n.rows), len(n.rows[0]) if n.rows else 0))
            for row in reversed(n.rows):
                stack.extend(reversed(row))
        else:
            raise TypeError(type(n))
    return out


def _lit(v) -> str:
    if isinstance(v, ExcelError):
        return f"_E({v.code!r})"
    return repr(v)


def _names_used(rules: list[Rule], wb) -> list:
    """規則が（たどって）使う名前。"""
    seen: dict[tuple, object] = {}
    todo = []
    for u in rules:
        if u.shape.ast is not None:
            todo += [(r, u.sheet) for r in iter_refs(u.shape.ast) if isinstance(r, NameNode)]
    while todo:
        ref, sheet = todo.pop()
        d = wb.defined_name(ref.name, ref.sheet or sheet) if hasattr(wb, "defined_name") else None
        if d is None or (d.name.lower(), d.sheet_scope) in seen:
            continue
        seen[(d.name.lower(), d.sheet_scope)] = d
        if d.ast is not None:
            todo += [(r, d.sheet_scope or sheet) for r in iter_refs(d.ast) if isinstance(r, NameNode)]
    return list(seen.values())


def _functions(nodes) -> set[str]:
    from ..formula.parser import iter_funcs

    out: set[str] = set()
    for n in nodes:
        out |= set(iter_funcs(n))
    return out


def _has_struct(nodes) -> bool:
    return any(isinstance(r, StructRef) for n in nodes for r in iter_refs(n))


def _collect_data(src, sheets: list[SheetScan], rules: list[Rule], all_cells: bool):
    """規則が読む定数の値：シート → {(行, 列): 値}。ほかのセルは持たない。"""
    reads: dict[str, dict[int, list]] = {}
    formula: dict[str, dict[int, list]] = {}
    widths = {s.name: s.max_col for s in sheets}
    for u in rules:
        cols = formula.setdefault(u.sheet, {})
        for m in u.members:
            for c in range(m.c1, m.c2 + 1):
                cols.setdefault(c, []).append((m.r1, m.r2))
        for rd in u.reads:
            if rd.external or rd.missing:
                continue
            for rect in read_rects(u, rd):
                rc = reads.setdefault(rd.sheet, {})
                for c in range(rect[1], min(rect[3], max(widths.get(rd.sheet, rect[1]), rect[1])) + 1):
                    rc.setdefault(c, []).append((rect[0], rect[2]))
    for d in (reads, formula):
        for cols in d.values():
            for c, iv in cols.items():
                cols[c] = _merge(iv)
    import bisect

    def inside(iv, r) -> bool:
        i = bisect.bisect_right(iv, (r, 1 << 30)) - 1
        return i >= 0 and iv[i][0] <= r <= iv[i][1]

    out: dict[str, dict] = {}
    for si, sh in enumerate(sheets):
        rc = reads.get(sh.name)
        if not all_cells and not rc:
            continue
        fc = formula.get(sh.name, {})
        got: dict = {}
        for r, c, v, _is_str in src.stored(si):
            if v is None:
                continue
            f = fc.get(c)
            if f and inside(f, r):
                continue
            if not all_cells:
                iv = rc.get(c)
                if not iv or not inside(iv, r):
                    continue
            got[(r, c)] = v
        out[sh.name] = got
    return out


def _phonetic(src, sheets, data: dict) -> dict:
    """PHONETIC が読むふりがな。"""
    out = {}
    wb = getattr(src, "wb", None)
    for si, sh in enumerate(sheets):
        cells = data.get(sh.name)
        if not cells:
            continue
        book = wb
        if book is None:
            rows = sorted({r for r, _c in cells})
            book = src.lz.rows_book({si: [(r, r) for r in rows]})
        for (r, c), v in cells.items():
            i = book.cell_sst.get((sh.name, r, c))
            if i is not None:
                got = book.phonetic.get(i, v)
                if got != v:
                    out[(sh.name, r, c)] = got
    return out


def generate(src, internal: dict, res: dict, name: str, password_note: bool = False) -> Generated:
    rules: list[Rule] = internal["rules"]
    sheets: list[SheetScan] = internal["sheets"]
    wb = src.meta
    names = _names_used(rules, wb)
    if any(u.shape.ast is not None and "INDIRECT" in _functions([u.shape.ast]) for u in rules):
        names = list(getattr(wb, "names", ()))  # INDIRECT は字で名前を指せる
    asts = [u.shape.ast for u in rules if u.shape.ast is not None] + [d.ast for d in names if d.ast is not None]
    funcs = _functions(asts)
    dynamic = bool(funcs & DYNAMIC_FUNCS)
    use_struct = _has_struct(asts)
    use_subtotal = bool(funcs & {"SUBTOTAL", "AGGREGATE"})
    # 計算の決まり：使う関数からたどる
    v = Vendor(RESERVED | {"_E", "_Ctx", "_Rule", "_T", "_build", "_Fail", "_WB", "_SheetView", "_CellView", "_DATA", "_NAMES", "_TABLES",
                           "_EXTERNAL", "_SHEETS", "_BOUNDS", "_HIDDEN", "_AT", "_R", "_READS", "_DATE1904", "_PHONETIC", "_a1", "_key",
                           "_parse_key", "_rule_cells", "_coerce", "_out", "_deep", "_load_data", "_load_extra", "_column", "_dump"})
    v.load_registry()
    v.need("formula.evaluator", "Evaluator")
    for n in ("Arr", "Rng", "ExcelError", "Unsupported"):
        v.need("formula.values", n)
    for n in ("Num", "Str", "Bool", "Err", "Missing", "RefNode", "NameNode", "StructRef", "Func", "Unary", "Postfix", "Binary", "ArrayLit"):
        v.need("formula.parser", n)
    v.need("formula.functions", "call")
    v.need("formula.functions", "EAGER")
    v.need("formula.functions", "LAZY")
    for f in sorted(funcs):
        v.want_function(f)
    v._drain()
    if use_struct:
        need_method("xlsx.calc", "Calc", "table_ref", v)
    vendored, final = v.source()
    regs = v.registrations(final)
    G = {n: v.global_name(final, m, n) for m, ns in (("formula.values", ("Arr", "Rng", "ExcelError", "Unsupported")),
                                                      ("formula.parser", ("Num", "Str", "Bool", "Err", "Missing", "RefNode", "NameNode", "StructRef",
                                                                          "Func", "Unary", "Postfix", "Binary", "ArrayLit")),
                                                      ("formula.evaluator", ("Evaluator",))) for n in ns}
    table_method = ""
    if use_struct:
        body = method_source("xlsx.calc", "Calc", "table_ref", v, final)
        table_method = "\n".join("    " + x for x in body.splitlines())

    data = _collect_data(src, sheets, rules, dynamic)
    phon = _phonetic(src, sheets, data) if "PHONETIC" in funcs else {}
    total = sum(len(x) for x in data.values())

    # 規則
    lines = []
    unrepro: dict[str, str] = {}
    for u in rules:
        if u.kind == "data_table":
            unrepro[u.id] = "データテーブル（What-If 分析）は、とどくでは計算しません"
        elif u.shape.ast is None:
            unrepro[u.id] = u.shape.error or "数式を読めません"
        elif u.__dict__.get("unknown_functions"):
            unrepro[u.id] = "とどくが計算できない関数：" + "、".join(u.__dict__["unknown_functions"])
        code = flat(u.shape.ast) if u.shape.ast is not None else None
        rects = [(m.r1, m.c1, m.r2, m.c2) for m in u.members]
        text = next((d["text"] for d in res["rules"] if d["id"] == u.id), "")
        lines.append(f"    # {u.id}  {_safe(u.sheet + '!' + text)}")
        lines.append(f"    _Rule({u.id!r}, {u.sheet!r}, {u.kind!r}, {(u.shape.r, u.shape.c)!r}, {u.origin!r}, {rects!r}, {code!r}),")
    rules_src = "\n".join(lines)
    order = res["order"]

    reads: dict[str, list] = {}
    for u in rules:
        for rd in u.reads:
            if not rd.external and not rd.missing:
                for rect in read_rects(u, rd):
                    reads.setdefault(rd.sheet, []).append(rect)
    for k in reads:
        reads[k] = sorted(set(reads[k]))

    names_src = repr([(d.name.lower(), d.sheet_scope, d.text, flat(d.ast) if d.ast is not None else None, d.error) for d in names])
    tables_src = "{}"
    if use_struct:
        ts = {k: (t.name, t.sheet, t.r1, t.c1, t.r2, t.c2, t.header_rows, t.totals_rows, list(t.columns))
              for k, t in getattr(wb, "tables", {}).items()}
        tables_src = repr(ts)
    ext_used = {rd.sheet for u in rules for rd in u.reads if rd.external}
    external_src = "{}"
    if ext_used or dynamic:
        ext = {}
        for i, sh in (getattr(wb, "external", None) or {}).items():
            ext[i] = {n: {k: x for k, x in cells.items()} for n, cells in sh.items()}
        external_src = _dict_lit(ext)
    hidden_src = "{}"
    if use_subtotal:
        hid = {}
        for si, sh in enumerate(sheets):
            h = list(sh.hidden_rows)
            meta = wb.sheets[si] if si < len(wb.sheets) else None
            af = getattr(meta, "autofilter", None)
            if h or af:
                hid[sh.name] = ([(h[i], h[i + 1]) for i in range(0, len(h), 2)], af)
        hidden_src = repr(hid)

    if total <= INLINE_DATA:
        data_src = _dict_lit(data)
        load = ""
    else:
        rows = [[s, [[r, c, _jv(x)] for (r, c), x in cells.items()]] for s, cells in data.items()]
        blob = base64.b85encode(zlib.compress(json.dumps(rows, ensure_ascii=False).encode("utf-8"), 6)).decode("ascii")
        data_src = "None  # 下の _DATA_PACKED を開いて使う（入力の値が多いので圧縮してある）"
        load = f'_DATA_PACKED = "{blob}"\n'

    inputs_meta = [{k: d[k] for k in ("sheet", "range", "cells", "filled", "value", "labels", "read_by", "constraints") if k in d} for d in res.get("inputs", [])]
    params_meta = [{k: d[k] for k in ("sheet", "range", "value", "labels", "read_by") if k in d} for d in res.get("parameters", [])]
    outs = [(u.sheet, *x) for u in rules for x in u.__dict__.get("outputs", [])]
    tmpl = Path(__file__).with_name("runtime_template.py.txt").read_text(encoding="utf-8")
    rule_lines = "\n".join(f"  {d['id']}  {_safe(d['sheet'] + '!' + d['text'])}" for d in res["rules"][:200])
    if len(res["rules"]) > 200:
        rule_lines += f"\n  …（全部で {len(res['rules'])} 個。RULES を見てください）"
    subs = {
        "@@SOURCE@@": _safe(name),
        "@@RULE_LINES@@": rule_lines,
        "@@VENDORED@@": vendored,
        "@@REGISTER@@": "\n".join(regs),
        "@@STDLIB@@": "\n".join(sorted(v.stdlib)),
        "@@RULES@@": rules_src,
        "@@ORDER@@": repr(order),
        "@@UNREPRO@@": repr(unrepro),
        "@@INPUTS@@": repr(inputs_meta),
        "@@PARAMETERS@@": repr(params_meta),
        "@@OUTPUTS@@": repr(outs),
        "@@READS@@": repr(reads),
        "@@DATA@@": data_src,
        "@@DATA_PACKED@@": load,
        "@@NAMES@@": names_src,
        "@@TABLES@@": tables_src,
        "@@EXTERNAL@@": external_src,
        "@@HIDDEN@@": hidden_src,
        "@@PHONETIC@@": repr(phon),
        "@@SHEETS@@": repr([s.name for s in sheets]),
        "@@BOUNDS@@": repr({s.name: (s.max_row, s.max_col) for s in sheets}),
        "@@DATE1904@@": repr(bool(getattr(wb, "date1904", False))),
        "@@TABLE_METHOD@@": table_method,
        "@@DYNAMIC@@": repr(dynamic),
        "@@FILENAME@@": repr(name),
        **{f"@@{k}@@": x for k, x in G.items()},
    }
    out = tmpl
    # 字の中に置き換えの印があっても壊れないよう、1 度なぞって置き換える
    import re

    out = re.sub(r"@@[A-Za-z0-9_]+@@", lambda m: subs.get(m[0], m[0]), out)
    facts = {"bytes": len(out.encode("utf-8")), "functions": sorted(funcs), "functions_included": sorted(v.functions),
             "unreproducible_rules": unrepro, "data_cells": total, "lines": out.count("\n") + 1}
    return Generated(out, facts)


def _safe(text: str) -> str:
    """書類から来た字を、ソースのコメント・説明の字に入れる形に（改行・引用符・逆斜線は字にする。書類の字がコードにならないように）。"""
    return json.dumps(text, ensure_ascii=False)[1:-1]


def _jv(v):
    if isinstance(v, ExcelError):
        return {"e": v.code}
    return v


def _dict_lit(d) -> str:
    def u(x):
        if isinstance(x, dict):
            return "{" + ", ".join(f"{u(k)}: {u(a)}" for k, a in x.items()) + "}"
        return _lit(x)

    return u(d)
