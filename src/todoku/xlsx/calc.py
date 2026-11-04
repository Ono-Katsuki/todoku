"""ブック全体の再計算と、保存されている計算結果との照合（関数の全数照合）。"""

from __future__ import annotations

import sys
import threading
from dataclasses import dataclass, field

from ..formula.evaluator import Evaluator
from ..formula.parser import Node, addr, iter_funcs, iter_refs
from ..formula.values import Arr, ExcelError, Rng, Unfinished, Unsupported, is_number, to_15
from .workbook import Cell, Workbook


class Cycle(Exception):
    pass


@dataclass
class CellResult:
    sheet: str
    address: str
    formula: str
    status: str  # match / mismatch / unsupported / volatile / no_cache / cycle
    expected: object = None
    actual: object = None
    reason: str = ""


@dataclass
class VerifyReport:
    results: list[CellResult] = field(default_factory=list)

    def count(self, status: str) -> int:
        return sum(1 for r in self.results if r.status == status)

    @property
    def total(self) -> int:
        return len(self.results)

    @property
    def ok(self) -> bool:
        return all(r.status in ("match", "volatile") for r in self.results)

    def summary(self) -> dict:
        keys = ["match", "mismatch", "unsupported", "volatile", "no_cache", "cycle"]
        return {"total": self.total, **{k: self.count(k) for k in keys}}


class Calc:
    """ブックの計算コンテキスト。

    overrides に値を入れると、そのセルの値を差し替えて再計算する（記入のシミュレーション）。
    再計算できない数式のセルは、保存されている値を使って計算を続ける。
    scan_store：大きな配列の集計の控え（formula/lazyarr.ScanStore。長い仕事が続きから計算する）。
    """

    def __init__(self, wb: Workbook, overrides: dict[tuple[str, int, int], object] | None = None, scan_store=None):
        self.wb = wb
        self.date1904 = wb.date1904
        self.overrides = overrides or {}
        self.memo: dict[tuple[str, int, int], object] = {}
        self.status: dict[tuple[str, int, int], tuple[str, str]] = {}
        self.visiting: set[tuple[str, int, int]] = set()
        self.name_cache: dict[tuple[str, str | None], tuple[Node, str] | None] = {}
        self._frames: list[bool] = []  # 計算中のセルごとに「計算のたびに変わる値に依存したか」
        self._sheets = {s.name: s for s in wb.sheets}
        self._sheets_lower = {s.name.lower(): s for s in wb.sheets}
        self.scan_store = scan_store  # 大きな配列の集計の控え
        self.scan_seq: dict = {}  # 数式のセルごとの、集計の順の番号（呼び直しても同じ控えに当たるように）
        self.calc_notes: dict = {}  # 計算の結果についての注（式で答えた和など）
        self.unfinished: dict = {}  # 時間の中で終わらなかった集計：数式のセル → Unfinished

    # ------------------------------------------------------------ Context
    def _ext(self, sheet: str):
        """外部ブックのシート "[1]Sheet1" なら、その保存値の表。"""
        if not sheet.startswith("["):
            return None
        import re

        m = re.match(r"\[(\d+)\](.*)", sheet)
        if not m:
            return None
        return self.wb.external.get(int(m[1]), {}).get(m[2].lower())

    def has_sheet(self, sheet: str) -> bool:
        return sheet in self._sheets or sheet.lower() in self._sheets_lower or self._ext(sheet) is not None

    def _sheet(self, sheet: str):
        return self._sheets.get(sheet) or self._sheets_lower[sheet.lower()]

    def bounds(self, sheet: str) -> tuple[int, int]:
        ext = self._ext(sheet)
        if ext is not None:
            return max((r for r, _ in ext), default=1), max((c for _, c in ext), default=1)
        s = self._sheet(sheet)
        return s.max_row, s.max_col

    def resolve_name(self, name: str, sheet: str | None, here: str):
        key = (name.lower(), sheet or here)
        if key in self.name_cache:
            return self.name_cache[key]
        d = self.wb.defined_name(name, sheet or here)
        res = None
        if d is not None:
            if d.ast is None:
                raise Unsupported(f"名前 {name}: {d.error}")
            res = (d.ast, d.sheet_scope or here)
        self.name_cache[key] = res
        return res

    def table_ref(self, node, sheet: str, row: int, col: int):
        """構造化参照を、シート上の範囲にする。"""
        from ..formula.values import REF, VALUE, Rng, err

        if node.table is None:
            t = next((x for x in self.wb.tables.values() if x.sheet == sheet and x.r1 <= row <= x.r2
                      and x.c1 <= col <= x.c2), None)
        else:
            t = self.wb.tables.get(node.table.lower())
        if t is None:
            return err(REF)
        data1, data2 = t.r1 + t.header_rows, t.r2 - t.totals_rows
        sp = set(node.specials) or {"#Data"}
        if "#This Row" in sp:
            if not data1 <= row <= data2:
                return err(VALUE)
            r1 = r2 = row
        elif "#All" in sp:
            r1, r2 = t.r1, t.r2
        else:
            spans = []
            if "#Headers" in sp:
                if not t.header_rows:
                    return err(REF)
                spans.append((t.r1, t.r1 + t.header_rows - 1))
            if "#Data" in sp:
                spans.append((data1, data2))
            if "#Totals" in sp:
                if not t.totals_rows:
                    return err(REF)
                spans.append((t.r2 - t.totals_rows + 1, t.r2))
            r1, r2 = min(a for a, _ in spans), max(b for _, b in spans)
        names = [c.lower() for c in t.columns]
        if node.col1 is None:
            c1, c2 = t.c1, t.c2
        else:
            try:
                i1 = names.index(node.col1.lower())
                i2 = names.index(node.col2.lower()) if node.col2 is not None else i1
            except ValueError:
                return err(REF)
            c1, c2 = t.c1 + min(i1, i2), t.c1 + max(i1, i2)
        if r2 < r1:
            return err(REF)
        return Rng(t.sheet, r1, c1, r2, c2)

    def cells_in(self, sheet: str, r1: int, c1: int, r2: int, c2: int):
        """範囲の中の、ファイルに要素のあるセルと差し替えたセル（行の順）。広い範囲を空のマスまで読まないため。読めないときは None。"""
        if self._ext(sheet) is not None:
            return None
        sh = self._sheet(sheet)
        if not hasattr(sh, "existing_in") or getattr(sh, "array_spans", None) or getattr(sh, "virtual_cell", None) is not None:
            return None
        have = set(sh.existing_in(r1, c1, r2, c2))
        for (s, r, c) in self.overrides:
            if s == sh.name and r1 <= r <= r2 and c1 <= c <= c2:
                have.add((r, c))
        return sorted(have)

    def cell_value(self, sheet: str, r: int, c: int):
        ext = self._ext(sheet)
        if ext is not None:
            return ext.get((r, c))
        sh = self._sheet(sheet)
        key = (sh.name, r, c)
        if key in self.overrides:
            return self.overrides[key]
        cell = sh.cells.get((r, c))
        if cell is None:
            spans = getattr(sh, "array_spans", None)
            if spans:
                return self._span_value(sh, r, c)  # 広い配列数式のマス：マスを作らず、頼まれたマスだけ
            virtual = getattr(sh, "virtual_cell", None)
            cell = virtual(r, c) if virtual is not None else None
            if cell is None:
                return None
        if not cell.has_formula:
            return cell.value
        v = self.memo[key] if key in self.memo else self._compute(cell, key)
        if self._frames and self.status.get(key, ("ok", ""))[0] == "volatile":
            self._frames[-1] = True  # TODAY() などに依存する値を使った数式も、計算のたびに変わる
        return v

    # ------------------------------------------------------------ 計算
    def _compute(self, cell: Cell, key):
        if cell.array_master is not None:
            mr, mc = cell.array_master
            master = self._sheet(cell.sheet).cells[(mr, mc)]
            self.cell_value(cell.sheet, mr, mc)
            st = self.status.get((master.sheet, mr, mc), ("ok", ""))
            v = self._array_pick(master, cell) if st[0] in ("ok", "volatile") else cell.value
            self.status[key] = st
            self.memo[key] = v
            return v
        if cell.ast is None:
            self.status[key] = ("unsupported", cell.parse_error or "数式を読めませんでした")
            self.memo[key] = cell.value
            return cell.value
        if key in self.visiting:
            raise Cycle(f"{cell.sheet}!{cell.address}")
        self.visiting.add(key)
        self._frames.append(False)
        tainted = False
        try:
            is_array = cell.formula_kind == "array"
            ev = Evaluator(self, cell.sheet, cell.row, cell.col, cell.dr, cell.dc, array=is_array)
            v = ev.run(cell.ast)
            tainted = self._frames[-1]
            if isinstance(v, Arr):
                self.memo[("__arr__", cell.sheet, cell.row, cell.col)] = v
                v = v.at(0, 0) if v.rows and v.cols else ExcelError("#VALUE!")
            if isinstance(v, Rng):
                v = ev.intersect(v)
            if v is None:
                v = 0.0  # 空白セルを参照した数式は 0 を返す
            if ev.volatile:
                self.status[key] = ("volatile", "計算のたびに値が変わる関数")
            elif tainted:
                self.status[key] = ("volatile", "計算のたびに値が変わる関数（TODAY など）の結果を使っている")
            else:
                self.status[key] = ("ok", "")
        except Unfinished as e:
            v = cell.value
            self.status[key] = ("unsupported", str(e))
            self.unfinished[key] = e
        except Unsupported as e:
            v = cell.value
            self.status[key] = ("unsupported", str(e))
        except Cycle:
            v = cell.value
            self.status[key] = ("cycle", "循環参照")
        except RecursionError:
            v = cell.value
            self.status[key] = ("unsupported", "参照が深すぎます")
        finally:
            self.visiting.discard(key)
            self._frames.pop()
        self.memo[key] = v
        return v

    def _span_value(self, sh, r: int, c: int):
        m = sh.array_span_at(r, c)
        if m is None:
            return None
        master = sh.cells[m]
        self.cell_value(sh.name, *m)
        st = self.status.get((sh.name, *m), ("ok", ""))
        if st[0] not in ("ok", "volatile"):
            return None
        return self._array_pick(master, Cell(sh.name, r, c))

    def _array_pick(self, master: Cell, cell: Cell):
        arr = self.memo.get(("__arr__", master.sheet, master.row, master.col))
        if not isinstance(arr, Arr):
            return cell.value
        i, j = cell.row - master.row, cell.col - master.col
        if arr.rows == 1:
            i = 0
        if arr.cols == 1:
            j = 0
        if i < arr.rows and j < arr.cols:
            v = arr.at(i, j)
            return 0.0 if v is None else v
        return ExcelError("#N/A")


def values_equal(expected, actual) -> bool:
    if isinstance(expected, ExcelError) or isinstance(actual, ExcelError):
        return isinstance(expected, ExcelError) and isinstance(actual, ExcelError) and expected.code == actual.code
    if isinstance(expected, bool) or isinstance(actual, bool):
        return expected is actual or (isinstance(expected, bool) and isinstance(actual, bool) and expected == actual)
    if is_number(expected) and is_number(actual):
        return to_15(float(expected)) == to_15(float(actual))
    if expected is None:
        return actual in (None, "", 0.0)
    if isinstance(expected, str) and isinstance(actual, str):
        return expected == actual
    return False


def formula_cells(wb: Workbook):
    for sh in wb.sheets:
        for (r, c), cell in sorted(sh.cells.items()):
            if cell.has_formula or cell.parse_error:
                yield cell


def _run_deep(fn):
    """数式の依存が長い（累計など）とPythonの再帰が深くなるので、大きなスタックで動かす。"""
    result: dict = {}

    def target():
        try:
            result["v"] = fn()
        except BaseException as e:  # noqa: BLE001
            result["e"] = e

    old = sys.getrecursionlimit()
    sys.setrecursionlimit(200000)
    threading.stack_size(512 * 1024 * 1024)
    try:
        t = threading.Thread(target=target)
        t.start()
        t.join()
    finally:
        threading.stack_size(0)
        sys.setrecursionlimit(old)
    if "e" in result:
        raise result["e"]
    return result["v"]


def verify(wb: Workbook) -> VerifyReport:
    """すべての数式を計算し直し、Excelが保存した計算結果と照合する。"""

    def run():
        calc = Calc(wb)
        rep = VerifyReport()
        for cell in formula_cells(wb):
            key = (cell.sheet, cell.row, cell.col)
            actual = calc.cell_value(cell.sheet, cell.row, cell.col)
            st, reason = calc.status.get(key, ("unsupported", cell.parse_error or ""))
            formula = cell.formula or (f"{{配列数式の一部: {addr(*cell.array_master)}}}" if cell.array_master else "")
            expected = cell.value
            if st == "ok":
                if expected is None and cell.t != "str":
                    status = "no_cache"
                    reason = "計算結果が保存されていません"
                else:
                    status = "match" if values_equal(expected, actual) else "mismatch"
            else:
                status = st
            rep.results.append(CellResult(cell.sheet, cell.address, formula, status, expected, actual, reason))
        return rep

    return _run_deep(run)


def recalc(wb: Workbook, overrides: dict[tuple[str, int, int], object]):
    """値を差し替えたときに、値が変わる数式セルを返す。"""

    def run():
        new = Calc(wb, overrides)
        changed: dict[tuple[str, int, int], object] = {}
        problems: dict[tuple[str, int, int], str] = {}
        for cell in formula_cells(wb):
            key = (cell.sheet, cell.row, cell.col)
            if key in overrides:
                continue
            after = new.cell_value(cell.sheet, cell.row, cell.col)
            st = new.status.get(key, ("ok", ""))
            if st[0] in ("unsupported", "cycle"):
                if _depends_on(wb, cell, overrides):
                    problems[key] = st[1]
                continue
            if not values_equal(cell.value, after):
                changed[key] = after
            if cell.dynamic and cell.array_ref and _spill_changes(new, cell):
                # 動的配列の結果の大きさが変わる：こぼれる範囲（とその外のセル）が変わるので、ここでは書かずに
                # Excelで開いたときに計算し直させる
                problems[key] = "動的配列の結果の大きさが変わる（こぼれる範囲が変わる）"
        # こぼれる範囲が変わる動的配列の、ほかのセルも書かない
        for (sh_name, r, c), why in list(problems.items()):
            m = wb.sheet(sh_name).cells.get((r, c))
            if m is not None and m.dynamic and m.array_ref:
                for k2 in list(changed):
                    cc = wb.sheet(k2[0]).cells.get(k2[1:])
                    if k2[0] == sh_name and cc is not None and cc.array_master == (r, c):
                        problems[k2] = why
        return changed, problems

    return _run_deep(run)


def _spill_changes(calc: "Calc", master: Cell) -> bool:
    """動的配列の数式の結果が、ファイルに保存された範囲（ref）と違う大きさになるか。"""
    from ..formula.parser import parse_cell

    arr = calc.memo.get(("__arr__", master.sheet, master.row, master.col))
    a = (master.array_ref or "").split(":")
    r1, c1 = parse_cell(a[0])
    r2, c2 = parse_cell(a[-1])
    rows, cols = (arr.rows, arr.cols) if isinstance(arr, Arr) else (1, 1)
    return (rows, cols) != (r2 - r1 + 1, c2 - c1 + 1)


def _depends_on(wb: Workbook, cell: Cell, overrides) -> bool:
    """再計算できない数式が、書き換えたセルに依存しているかを（保守的に）調べる。"""
    if cell.ast is None:
        return True
    for ref in iter_refs(cell.ast):
        return True
    return False


def used_functions(wb: Workbook) -> dict[str, int]:
    out: dict[str, int] = {}
    seen = set()
    for cell in formula_cells(wb):
        if cell.ast is None:
            continue
        k = id(cell.ast)
        for f in iter_funcs(cell.ast):
            out[f] = out.get(f, 0) + 1
        seen.add(k)
    return out
