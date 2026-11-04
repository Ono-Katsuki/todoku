"""マクロ（VBA）の静的な読み取り：手続きのソース、読み書きするセル（字で決まる所）、いつ動くか、規則との関係。

動かさない（動かすのは run_macro）。参照が変数や計算で決まる所は「分からない」と書く（推測で埋めない）。
"""

from __future__ import annotations

import re

from ..formula.parser import addr, parse_cell
from ..vbarun import parse as P
from .rules import RectIndex, Rule

VALUE_PROPS = {"value", "value2", "formula", "formular1c1", "formulalocal", "formular1c1local", "formula2", "text"}
WRITE_METHODS = {"clearcontents", "clear", "delete", "insert", "autofill", "filldown", "fillright", "pastespecial", "sort",
                 "removeduplicates", "cut", "merge", "unmerge", "formulaarray", "autofilter"}
READ_METHODS = {"copy", "find", "select", "activate", "printout", "calculate"}
SHEET_WORDS = {"sheets", "worksheets"}
BOOK_WORDS = {"thisworkbook", "activeworkbook", "application", "workbooks"}
_EXCEL_FORMAT = {"interior", "font", "borders", "numberformat", "numberformatlocal", "locked", "hidden", "columnwidth", "rowheight",
                 "horizontalalignment", "verticalalignment", "wraptext", "style", "entirerow", "entirecolumn", "mergecells"}


class _Sheet:
    def __init__(self, name: str | None, how: str):
        self.name = name  # None は実行するまで決まらない
        self.how = how


class _Range:
    def __init__(self, sheet: _Sheet, ref: str | None, expr: str):
        self.sheet = sheet
        self.ref = ref  # 「B3」「B3:C9」「B?」など。None は決まらない
        self.expr = expr


def _src(n) -> str:
    """式を短い字に（分からない所を伝えるため）。"""
    if n is None:
        return ""
    if isinstance(n, P.Lit):
        return f'"{n.value}"' if isinstance(n.value, str) else str(n.value)
    if isinstance(n, P.Name):
        return n.raw or n.name
    if isinstance(n, P.Member):
        return (_src(n.obj) if n.obj is not None else "") + "." + (n.raw or n.name)
    if isinstance(n, P.Call):
        return _src(n.target) + "(" + ", ".join(_src(a.value) for a in n.args) + ")"
    if isinstance(n, P.Binary):
        return f"{_src(n.left)} {n.op} {_src(n.right)}"
    if isinstance(n, P.Unary):
        return f"{n.op}{_src(n.operand)}"
    return type(n).__name__


def _lit(n):
    if isinstance(n, P.Lit):
        return n.value
    if isinstance(n, P.Binary) and n.op == "&":
        a, b = _lit(n.left), _lit(n.right)
        if a is not None and b is not None:
            return f"{a}{b}"
    return None


def _col_of(n) -> int | None:
    v = _lit(n)
    if isinstance(v, (int, float)) and v >= 1:
        return int(v)
    if isinstance(v, str) and re.fullmatch(r"[A-Za-z]{1,3}", v):
        c = 0
        for ch in v.upper():
            c = c * 26 + ord(ch) - 64
        return c
    return None


def _word(n) -> str:
    return (n.name.lower() if isinstance(n, (P.Name, P.Member)) else "") if n is not None else ""


class _Walker:
    def __init__(self, module: str, sheet_of_module: str | None, wb, code_names: dict[str, str]):
        self.module = module
        self.home = sheet_of_module
        self.wb = wb
        self.code = code_names
        self.reads: list[dict] = []
        self.writes: list[dict] = []
        self.calls: list[str] = []
        self.formats: list[dict] = []
        self.unknown: list[str] = []
        self.env: dict[str, object] = {}
        self.with_: list[object] = []
        self.formula_writes: list[dict] = []

    def _default_sheet(self) -> _Sheet:
        if self.home:
            return _Sheet(self.home, "module")
        return self._active()

    def _active(self) -> _Sheet:
        """ActiveSheet：シートが 1 枚なら決まる。"""
        if len(self.wb.sheets) == 1:
            return _Sheet(self.wb.sheets[0].name, "only sheet")
        return _Sheet(None, "ActiveSheet")

    def sheet_from(self, n) -> _Sheet | None:
        if isinstance(n, P.Name):
            if n.name in self.env and isinstance(self.env[n.name], _Sheet):
                return self.env[n.name]
            if n.name == "activesheet":
                return self._active()
            if n.name in self.code:
                return _Sheet(self.code[n.name], "codename")
            return None
        if isinstance(n, P.Member):
            w = n.name.lower()
            if w == "activesheet":
                return self._active()
            obj = self._obj(n)
            if w in self.code and (obj is None or _word(obj) in BOOK_WORDS):
                return _Sheet(self.code[w], "codename")
            if n.bang:
                return None
            return None
        if isinstance(n, P.Call):
            w = _word(n.target)
            if w in SHEET_WORDS and n.args:
                v = _lit(n.args[0].value)
                if isinstance(v, str):
                    return _Sheet(v, "name")
                if isinstance(v, (int, float)) and 1 <= int(v) <= len(self.wb.sheets):
                    return _Sheet(self.wb.sheets[int(v) - 1].name, "index")
                return _Sheet(None, _src(n))
        return None

    def _obj(self, n: P.Member):
        if n.obj is not None:
            return n.obj
        return self.with_[-1][0] if self.with_ else None

    def range_from(self, n) -> _Range | None:
        """式が指す範囲（なければ None）。"""
        if isinstance(n, P.Name):
            got = self.env.get(n.name)
            if isinstance(got, _Range):
                return got
            if n.name in ("target",):
                return _Range(self._default_sheet(), None, "Target")
            if n.name == "selection":
                return _Range(self._active(), None, "Selection")
            return None
        if isinstance(n, P.Member) and n.name.lower() in ("value", "value2", "formula", "formular1c1", "text", "formulalocal", "formula2"):
            return self.range_from(n.obj) if n.obj is not None else (self.with_[-1][1] if self.with_ and self.with_[-1][1] else None)
        if isinstance(n, P.Member) and n.name.lower() == "activecell":
            return _Range(self._active(), None, "ActiveCell")
        if isinstance(n, P.Member) and n.name.lower() in ("usedrange", "currentregion", "entirerow", "entirecolumn"):
            return _Range(self._parent_sheet(n), None, _src(n))
        if isinstance(n, P.Call):
            t = n.target
            w = _word(t)
            if w == "range":
                sheet = self._parent_sheet(t)
                return self._range_args(n, sheet)
            if w == "cells":
                sheet = self._parent_sheet(t)
                return self._cells_args(n, sheet)
            if w == "columns" or w == "rows":
                sheet = self._parent_sheet(t)
                v = _lit(n.args[0].value) if n.args else None
                if isinstance(v, str):
                    return _Range(sheet, v, _src(n))
                if w == "columns":
                    c = _col_of(n.args[0].value) if n.args else None
                    if c:
                        from ..formula.parser import num_to_col

                        return _Range(sheet, f"{num_to_col(c)}:{num_to_col(c)}", _src(n))
                if w == "rows" and isinstance(v, (int, float)):
                    return _Range(sheet, f"{int(v)}:{int(v)}", _src(n))
                return _Range(sheet, None, _src(n))
            if w == "offset" or w == "resize":
                base = self.range_from(t.obj) if isinstance(t, P.Member) else None
                if base is not None:
                    return _Range(base.sheet, self._moved(base, n, w), _src(n))
            if w == "intersect":
                return None
        return None

    def _moved(self, base: _Range, n: P.Call, w: str) -> str | None:
        if base.ref is None or ":" in base.ref or "?" in base.ref:
            return None
        vals = [_lit(a.value) if a.value is not None else 0 for a in n.args]
        if any(not isinstance(v, (int, float)) for v in vals):
            return None
        try:
            r, c = parse_cell(base.ref.replace("$", ""))
        except Exception:  # noqa: BLE001
            return None
        vals = [int(v) for v in vals] + [0, 0]
        if w == "offset":
            return addr(r + vals[0], c + vals[1])
        h, wd = max(vals[0], 1), max(vals[1], 1)
        return f"{addr(r, c)}:{addr(r + h - 1, c + wd - 1)}"

    def _parent_sheet(self, t) -> _Sheet:
        if isinstance(t, P.Member):
            obj = self._obj(t)
            if obj is None:
                return self._default_sheet()
            if self.with_ and t.obj is None and self.with_[-1][2] is not None:
                return self.with_[-1][2]
            got = self.sheet_from(obj)
            if got is not None:
                return got
            base = self.range_from(obj)
            if base is not None:
                return base.sheet
            if _word(obj) in BOOK_WORDS:
                return self._default_sheet()
            return _Sheet(None, _src(obj))
        return self._default_sheet()

    def _range_args(self, n: P.Call, sheet: _Sheet) -> _Range:
        if not n.args:
            return _Range(sheet, None, _src(n))
        a = _lit(n.args[0].value)
        if isinstance(a, str):
            if len(n.args) == 2:
                b = _lit(n.args[1].value)
                if isinstance(b, str):
                    return _Range(sheet, f"{a}:{b}", _src(n))
                return _Range(sheet, None, _src(n))
            return _Range(sheet, a, _src(n))
        return _Range(sheet, None, _src(n))

    def _cells_args(self, n: P.Call, sheet: _Sheet) -> _Range:
        if len(n.args) < 2:
            return _Range(sheet, None, _src(n))
        rv = _lit(n.args[0].value)
        col = _col_of(n.args[1].value)
        from ..formula.parser import num_to_col

        if isinstance(rv, (int, float)) and col:
            return _Range(sheet, addr(int(rv), col), _src(n))
        if col:  # 列だけ決まる
            return _Range(sheet, f"{num_to_col(col)}?", _src(n))
        if isinstance(rv, (int, float)):
            return _Range(sheet, f"?{int(rv)}", _src(n))
        return _Range(sheet, None, _src(n))

    # ------------------------------------------------------------ 記録
    def _note(self, acc: list, rng: _Range, line: int, how: str | None = None) -> None:
        d = {"sheet": rng.sheet.name, "cell": rng.ref, "expression": rng.expr, "line": line}
        if rng.sheet.name is None:
            d["sheet_unknown"] = rng.sheet.how
        if rng.ref is None:
            d["cell_unknown"] = True
            self.unknown.append(rng.expr)
        if how:
            d["how"] = how
        k = (d["sheet"], d["cell"], d.get("how"))
        if not any((x["sheet"], x["cell"], x.get("how")) == k for x in acc):
            acc.append(d)

    def read_expr(self, n, line: int) -> None:
        """式の中の範囲を、読む所として記録する。"""
        if n is None or isinstance(n, (P.Lit, P.DateLit)):
            return
        rng = self.range_from(n)
        if rng is not None:
            self._note(self.reads, rng, line)
            if isinstance(n, P.Call):
                for a in n.args:
                    self.read_expr(a.value, line)
            return
        if isinstance(n, P.Member):
            if n.obj is not None:
                self.read_expr(n.obj, line)
            return
        if isinstance(n, P.Call):
            self._call_name(n)
            if isinstance(n.target, P.Member):
                self.read_expr(n.target, line)
            for a in n.args:
                self.read_expr(a.value, line)
            return
        if isinstance(n, P.Binary):
            self.read_expr(n.left, line)
            self.read_expr(n.right, line)
        elif isinstance(n, P.Unary):
            self.read_expr(n.operand, line)

    def _call_name(self, n: P.Call) -> None:
        t = n.target
        if isinstance(t, P.Name):
            self.calls.append(t.name)
        elif isinstance(t, P.Member) and t.name.lower() == "run" and n.args:
            v = _lit(n.args[0].value)
            self.calls.append(str(v).rsplit(".", 1)[-1].lower() if isinstance(v, str) else f"?{_src(n.args[0].value)}")

    def walk(self, body: list) -> None:
        for s in body or ():
            self.stmt(s)

    def stmt(self, s) -> None:
        line = getattr(s, "line", 0)
        if isinstance(s, P.Assign):
            self._assign(s, line)
        elif isinstance(s, P.CallStmt):
            self._call_stmt(s.expr, line)
        elif isinstance(s, P.If):
            for cond, body in s.branches:
                self.read_expr(cond, line)
                self.walk(body)
            self.walk(s.else_body)
        elif isinstance(s, P.For):
            self.read_expr(s.start, line)
            self.read_expr(s.end, line)
            self.walk(s.body)
        elif isinstance(s, P.ForEach):
            self.read_expr(s.coll, line)
            rng = self.range_from(s.coll)
            if rng is not None and isinstance(s.var, P.Name):
                self.env[s.var.name] = _Range(rng.sheet, None, f"{s.var.raw or s.var.name}（{rng.expr} の中のセル）")
            self.walk(s.body)
        elif isinstance(s, P.Do):
            if s.pre:
                self.read_expr(s.pre[1], line)
            if s.post:
                self.read_expr(s.post[1], line)
            self.walk(s.body)
        elif isinstance(s, P.Select):
            self.read_expr(s.expr, line)
            for _c, body in s.cases:
                self.walk(body)
        elif isinstance(s, P.With):
            rng = self.range_from(s.obj)
            sheet = self.sheet_from(s.obj)
            self.with_.append((s.obj, rng, sheet))
            self.walk(s.body)
            self.with_.pop()

    def _assign(self, s: P.Assign, line: int) -> None:
        if s.set and isinstance(s.target, P.Name):
            sheet = self.sheet_from(s.value)
            if sheet is not None:
                self.env[s.target.name] = sheet
                return
            rng = self.range_from(s.value)
            if rng is not None:
                self.env[s.target.name] = rng
                return
        self.read_expr(s.value, line)
        t = s.target
        if isinstance(t, P.Name) and isinstance(self.env.get(t.name), _Range):
            self._note(self.writes, self.env[t.name], line)
            return
        if isinstance(t, P.Member):
            prop = t.name.lower()
            if prop in VALUE_PROPS:
                rng = self.range_from(t.obj) if t.obj is not None else (self.with_[-1][1] if self.with_ else None)
                if rng is not None:
                    self._note(self.writes, rng, line, how=prop if prop != "value" else None)
                    if prop.startswith("formula"):
                        v = _lit(s.value)
                        d = {"sheet": rng.sheet.name, "cell": rng.ref, "line": line}
                        if v is not None:
                            d["formula"] = str(v)
                        else:
                            d["formula_expression"] = _src(s.value)
                        self.formula_writes.append(d)
                return
            base = self.range_from(t.obj) if t.obj is not None else (self.with_[-1][1] if self.with_ else None)
            if base is None and isinstance(t.obj, P.Member):
                base = self.range_from(t.obj.obj) if t.obj.obj is not None else (self.with_[-1][1] if self.with_ else None)
            if base is not None:
                self._note(self.formats, base, line, how=prop)
                return
            return
        rng = self.range_from(t)
        if rng is not None:  # Range("A1") = 1（既定の Value）
            self._note(self.writes, rng, line)

    def _call_stmt(self, e, line: int) -> None:
        if isinstance(e, P.Name):
            self.calls.append(e.name)
            return
        target = e.target if isinstance(e, P.Call) else e
        args = e.args if isinstance(e, P.Call) else []
        if isinstance(target, P.Member):
            w = target.name.lower()
            base = self.range_from(target.obj) if target.obj is not None else (self.with_[-1][1] if self.with_ else None)
            if base is not None and (w in WRITE_METHODS or w in READ_METHODS):
                if w in WRITE_METHODS:
                    self._note(self.writes, base, line, how=w)
                else:
                    self._note(self.reads, base, line, how=w)
                for a in args:
                    if a.name and a.name.lower() == "destination":
                        dst = self.range_from(a.value)
                        if dst is not None:
                            self._note(self.writes, dst, line, how="copy destination")
                return
            if w == "run" and args:
                self._call_name(e)
                return
        if isinstance(e, P.Call):
            self._call_name(e)
        for a in args:
            self.read_expr(a.value, line)
        if isinstance(target, P.Member) and target.obj is not None:
            self.read_expr(target.obj, line)


_PROC_RE = re.compile(r"^[ \t]*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([^\s(]+)", re.I | re.M)
_END_RE = re.compile(r"^[ \t]*End\s+(Sub|Function|Property)\b", re.I | re.M)


def _proc_sources(source: str) -> dict[str, str]:
    """手続きの名前（小文字） → その手続きのソース（書かれたまま）。"""
    out: dict[str, str] = {}
    for m in _PROC_RE.finditer(source):
        e = _END_RE.search(source, m.end())
        end = source.find("\n", e.end()) if e else -1
        text = source[m.start():(end if end != -1 else len(source))] if e else source[m.start():]
        out.setdefault(m[2].lower(), text.strip("\r\n"))
    return out


def _intersect_watches(proc) -> list[str]:
    """Worksheet_Change の見張る範囲：Intersect(Target, Range("B3:B9")) や Target.Address = "$B$3"（字で書かれているもの）。"""
    out: list[str] = []

    def walk(n):
        if isinstance(n, list):
            for x in n:
                walk(x)
            return
        if isinstance(n, tuple):
            for x in n:
                walk(x)
            return
        if isinstance(n, P.Call):
            if _word(n.target) == "intersect":
                for a in n.args:
                    v = a.value
                    if isinstance(v, P.Call) and _word(v.target) == "range" and v.args:
                        t = _lit(v.args[0].value)
                        out.append(str(t) if isinstance(t, str) else _src(v))
            for a in n.args:
                walk(a.value)
            walk(n.target)
            return
        if isinstance(n, P.Binary):
            if n.op == "=":
                left, right = n.left, n.right
                if isinstance(left, P.Member) and left.name.lower() == "address" and _word(left.obj) == "target":
                    t = _lit(right)
                    out.append(str(t).replace("$", "") if isinstance(t, str) else _src(right))
                if isinstance(left, P.Member) and left.name.lower() in ("column", "row") and _word(left.obj) == "target":
                    out.append(f"Target.{left.name} = {_src(right)}")
            walk(n.left)
            walk(n.right)
            return
        if isinstance(n, P.Unary):
            walk(n.operand)
            return
        if isinstance(n, P.Node) or hasattr(n, "__dataclass_fields__"):
            for f in n.__dataclass_fields__:
                x = getattr(n, f)
                if isinstance(x, (list, tuple)) or hasattr(x, "__dataclass_fields__"):
                    walk(x)

    walk(proc.body)
    seen = []
    for x in out:
        if x not in seen:
            seen.append(x)
    return seen


def analyse(wb, rules: list[Rule], ix: RectIndex, sheets) -> dict | None:
    from ..macros import EVENTS
    from ..vba import visible_source
    from ..vbarun.parse import parse_module

    has = getattr(wb, "has_macros", False)
    if not has:
        return None
    project = getattr(wb, "vba", None)
    if project is None:
        return {"present": True, "readable": False, "note": "マクロの入れ物を読めませんでした。何をするかは分かりません"}
    code = {s.code_name.lower(): s.name for s in wb.sheets if getattr(s, "code_name", None)}
    sheet_by_code = {s.code_name: s.name for s in wb.sheets if getattr(s, "code_name", None)}
    buttons: dict[str, list[dict]] = {}
    for s in wb.sheets:
        for c in getattr(s, "controls", ()):
            m = getattr(c, "macro", None)
            if m:
                proc = m.rsplit("!", 1)[-1].strip("'").rsplit(".", 1)[-1].lower()
                buttons.setdefault(proc, []).append({"sheet": s.name, "control": c.name, "text": c.text, "macro": m})
    for rec in _shape_macros(wb):
        proc = rec["macro"].rsplit("!", 1)[-1].strip("'").rsplit(".", 1)[-1].lower()
        buttons.setdefault(proc, []).append(rec)
    consts = getattr(project, "cc_constants", None)
    all_procs: list[dict] = []
    errors = []
    name_set = set()
    for m in project.modules:
        if not m.source:
            continue
        home = sheet_by_code.get(m.name)
        src = visible_source(m)
        texts = _proc_sources(src or "")
        try:
            mod = parse_module(m.name, m.source, consts)
        except Exception as e:  # noqa: BLE001  読めないモジュールは、ソースをそのまま示す
            errors.append({"module": m.name, "error": f"{type(e).__name__}: {e}", "source": src})
            continue
        for key, proc in mod.procs.items():
            w = _Walker(m.name, home, wb, code)
            try:
                w.walk(proc.body)
                failed = None
            except Exception as e:  # noqa: BLE001
                failed = f"{type(e).__name__}: {e}"
            base = proc.name.lower()
            name_set.add(base)
            event = EVENTS.get(base)
            trigger = {}
            if event:
                trigger["event"] = proc.raw
                trigger["when"] = event
                watch = _intersect_watches(proc)
                if base.endswith("change") or base.endswith("sheetchange"):
                    trigger["watches"] = watch or "字で書かれていない（どのセルの変更で動くかは、ソースを読んで判断）"
                    if home:
                        trigger["sheet"] = home
            elif m.kind in ("document", "form") and base.endswith("_click"):
                trigger["event"] = proc.raw
                trigger["when"] = f"ボタン「{proc.raw[:-6]}」を押したとき"
                if home:
                    trigger["sheet"] = home
            if base in buttons:
                trigger["buttons"] = buttons[base]
            d = {"module": m.name, "name": proc.raw, "ref": f"{m.name}.{proc.raw}", "kind": proc.kind,
                 "source": texts.get(base, ""), "reads": w.reads, "writes": w.writes,
                 "calls": sorted({c for c in w.calls}), "_w": w}
            if w.formula_writes:
                d["writes_formulas"] = w.formula_writes
            if w.formats:
                d["format_changes"] = w.formats
            if trigger:
                d["trigger"] = trigger
            if failed:
                d["analysis_failed"] = failed
            d["unknown"] = sorted(set(w.unknown))
            all_procs.append(d)
    procs_lower = {p["name"].lower() for p in all_procs}
    for p in all_procs:
        p.pop("_w")
        p["calls"] = [c for c in p["calls"] if c in procs_lower or c.startswith("?")]
        p["rules"] = _relations(p, ix, sheets)
        if not p["unknown"]:
            p.pop("unknown")
    triggers = [p["ref"] for p in all_procs if p.get("trigger")]
    out = {"present": True, "readable": True, "procedures": all_procs, "triggered": triggers,
           "note": "読み書きは字で決まる所だけ。cell_unknown・sheet_unknown の所は、実行するまで決まりません。ボタンのマクロを動かすには run_macro"}
    if errors:
        out["unparsed_modules"] = errors
    return out


_SHAPE_MACRO = re.compile(rb"<(?:\w+:)?(?:sp|pic|graphicFrame|grpSp)\b[^>]*?\bmacro=\"([^\"]+)\"")


def _shape_macros(wb) -> list[dict]:
    """図形に割り当てたマクロ（xdr:sp の macro）。形式が .xlsx のときだけ、図の部品から読む。"""
    out: list[dict] = []
    rels_of = getattr(wb, "rels", None)
    z = getattr(wb, "zip", None)
    if rels_of is None or z is None or not hasattr(z, "read"):
        return out
    for s in wb.sheets:
        path = getattr(s, "path", "")
        if not path:
            continue
        try:
            rels = rels_of(path)
        except Exception:  # noqa: BLE001
            continue
        for typ, target in rels.values():
            if typ != "drawing" or not z.has(target):
                continue
            for m in _SHAPE_MACRO.finditer(z.read(target)):
                out.append({"sheet": s.name, "control": "図形", "macro": m[1].decode("utf-8", "replace").replace("&apos;", "'")})
    return out


def _rect_of(ref: str) -> tuple[int, int, int, int] | None:
    t = ref.replace("$", "")
    try:
        a = t.split(":")
        r1, c1 = parse_cell(a[0])
        r2, c2 = parse_cell(a[-1])
    except Exception:  # noqa: BLE001  名前や列全体
        return None
    return r1, c1, r2, c2


def _relations(p: dict, ix: RectIndex, sheets) -> dict:
    """読み書きする所と規則の関係（字で決まる所だけ）。"""
    overwrites, feeds, reads_results = set(), set(), set()
    readers = _reader_index(ix)
    for w in p["writes"]:
        if w["sheet"] and w["cell"]:
            rect = _rect_of(w["cell"])
            if rect:
                overwrites |= {h[4] for h in ix.hits(w["sheet"], *rect)}
                feeds |= {h[4] for h in readers.hits(w["sheet"], *rect)}
    for w in p["reads"]:
        if w["sheet"] and w["cell"]:
            rect = _rect_of(w["cell"])
            if rect:
                reads_results |= {h[4] for h in ix.hits(w["sheet"], *rect)}

    def ids(s):
        return sorted(s, key=lambda x: int(x[1:]))

    d = {}
    if overwrites:
        d["overwrites_output_of"] = ids(overwrites)
    if feeds:
        d["feeds_input_of"] = ids(feeds)
    if reads_results:
        d["reads_output_of"] = ids(reads_results)
    return d


def _reader_index(ix: RectIndex) -> RectIndex:
    return getattr(ix, "readers", None) or RectIndex().freeze()
