"""構文木を、Excelと同じ規則で計算する。"""

from __future__ import annotations

from fractions import Fraction
from typing import Callable, Protocol

from .parser import (
    MAX_COL,
    MAX_ROW,
    ArrayLit,
    Binary,
    Bool,
    Err,
    Func,
    Missing,
    NameNode,
    Node,
    Num,
    Postfix,
    RefNode,
    Str,
    StructRef,
    Unary,
)
from .. import limits
from .lazyarr import FlatView, LazyArr, RangeArr, make as make_lazy, poly_binary, want_lazy
from .dates import _CURRENT_1904
from .values import (
    NAME,
    REF,
    TEXT_MAX,
    VALUE,
    Arr,
    ExcelError,
    Rng,
    Unsupported,
    array_room,
    compare,
    is_big,
    err,
    to_number,
    to_text,
)


class Context(Protocol):
    def cell_value(self, sheet: str, r: int, c: int): ...

    def bounds(self, sheet: str) -> tuple[int, int]: ...

    def resolve_name(self, name: str, sheet: str | None, here: str) -> tuple[Node, str] | None: ...

    def has_sheet(self, sheet: str) -> bool: ...

    date1904: bool


class Evaluator:
    """1つのセルの数式を計算する。"""

    def __init__(self, ctx: Context, sheet: str, row: int, col: int, dr: int = 0, dc: int = 0, array: bool = False):
        self.ctx = ctx
        self.sheet = sheet
        self.row = row
        self.col = col
        self.dr = dr
        self.dc = dc
        self.array = array
        self.volatile = False
        self.date1904 = bool(getattr(ctx, "date1904", False))  # ブックの日付の数え方（DATE・TEXT などが使う）

    # -------------------------------------------------------------- 入口
    def run(self, node: Node):
        v = self.ev(node)
        if isinstance(v, Rng):
            if self.array:
                return self.materialize(v)
            v = self.intersect(v)
        return v

    def ev(self, node: Node):
        if _CURRENT_1904.get() is not self.date1904:
            token = _CURRENT_1904.set(self.date1904)  # この計算の間は、ブックの日付の数え方で日付を数える
            try:
                return self._ev(node)
            except ExcelError as e:
                return e
            finally:
                _CURRENT_1904.reset(token)
        try:
            return self._ev(node)
        except ExcelError as e:
            return e

    # -------------------------------------------------------------- 値の取り出し
    def materialize(self, rng: Rng) -> Arr:
        rng = self.clip(rng)
        if (rng.r2 - rng.r1 + 1) * (rng.c2 - rng.c1 + 1) > limits.MAX_ARRAY_CELLS:
            return RangeArr(self.ctx, rng)  # 広い範囲：リストに開かず、頼まれたマスだけ読む
        return Arr([[self.ctx.cell_value(rng.sheet, r, c) for c in range(rng.c1, rng.c2 + 1)]
                    for r in range(rng.r1, rng.r2 + 1)])

    def clip(self, rng: Rng) -> Rng:
        """列全体・行全体の参照を、実際に使われている範囲に縮める。"""
        if rng.r2 >= MAX_ROW or rng.c2 >= MAX_COL:
            mr, mc = self.ctx.bounds(rng.sheet)
            r2 = min(rng.r2, max(mr, rng.r1))
            c2 = min(rng.c2, max(mc, rng.c1))
            return Rng(rng.sheet, rng.r1, rng.c1, r2, c2)
        return rng

    def intersect(self, rng: Rng):
        """暗黙の共通部分：範囲を1つの値として使うとき、数式と同じ行・列のセルを取る。"""
        if rng.is_cell:
            return self.ctx.cell_value(rng.sheet, rng.r1, rng.c1)
        if rng.c1 == rng.c2 and rng.r1 <= self.row <= rng.r2:
            return self.ctx.cell_value(rng.sheet, self.row, rng.c1)
        if rng.r1 == rng.r2 and rng.c1 <= self.col <= rng.c2:
            return self.ctx.cell_value(rng.sheet, rng.r1, self.col)
        return err(VALUE)

    def scalar(self, v):
        if isinstance(v, Rng):
            return self.intersect(v)
        if isinstance(v, Arr):
            return v.data[0][0] if v.data and v.data[0] else err(VALUE)
        return v

    def values(self, v) -> list:
        """引数を値の平らなリストにする（範囲・配列はすべてのセル）。"""
        if isinstance(v, Rng):
            v = self.materialize(v)
        if isinstance(v, LazyArr):
            return FlatView(v)  # 大きな配列：作らずに順に読む
        if isinstance(v, Arr):
            return [x for row in v.data for x in row]
        return [v]

    def grid(self, v) -> Arr:
        if isinstance(v, Rng):
            return self.materialize(v)
        if isinstance(v, Arr):
            return v
        return Arr([[v]])

    # -------------------------------------------------------------- 計算本体
    def _ev(self, node: Node):
        if isinstance(node, Num):
            return node.value
        if isinstance(node, Str):
            return node.value
        if isinstance(node, Bool):
            return node.value
        if isinstance(node, Err):
            return err(node.code)
        if isinstance(node, Missing):
            return None
        if isinstance(node, RefNode):
            n = node.shifted(self.dr, self.dc)
            sheet = n.sheet or self.sheet
            if not self.ctx.has_sheet(sheet):
                return err(REF)
            if n.r1 < 1 or n.c1 < 1 or n.r2 > MAX_ROW or n.c2 > MAX_COL:
                return err(REF)
            return Rng(sheet, n.r1, n.c1, n.r2, n.c2)
        if isinstance(node, StructRef):
            if not hasattr(self.ctx, "table_ref"):
                raise Unsupported("構造化参照（テーブル）")
            return self.ctx.table_ref(node, self.sheet, self.row, self.col)
        if isinstance(node, NameNode):
            scope = getattr(self, "scope", None)
            if scope and node.sheet is None:
                import re as _re

                key = _re.sub(r"^(?i:_xlpm\.)", "", node.name).upper()
                if key in scope:
                    return scope[key]  # LET で決めた名前
            res = self.ctx.resolve_name(node.name, node.sheet, self.sheet)
            if res is None:
                return err(NAME)
            sub_node, sub_sheet = res
            sub = Evaluator(self.ctx, sub_sheet, self.row, self.col, 0, 0, self.array)
            v = sub.ev(sub_node)
            self.volatile |= sub.volatile
            return v
        if isinstance(node, ArrayLit):
            return Arr([[self._ev(x) for x in row] for row in node.rows])
        if isinstance(node, Unary):
            v = self.ev(node.operand)
            return self.map1(v, lambda x: -to_number(x) if node.op == "-" else x, node.op if node.op == "-" else "+")
        if isinstance(node, Postfix):
            v = self.ev(node.operand)
            return self.map1(v, lambda x: to_number(x) / 100, "%")
        if isinstance(node, Binary):
            if node.op == ":":
                return self.range_op(node)
            a = self.ev(node.left)
            b = self.ev(node.right)
            return self.map2(a, b, _BIN[node.op], node.op)
        if isinstance(node, Func):
            from .functions import call

            return call(self, node)
        raise Unsupported(f"未対応の構文: {type(node).__name__}")

    def range_op(self, node: Binary):
        a = self.ev(node.left)
        b = self.ev(node.right)
        if isinstance(a, ExcelError):
            return a
        if isinstance(b, ExcelError):
            return b
        if not isinstance(a, Rng) or not isinstance(b, Rng) or a.sheet != b.sheet:
            return err(VALUE)
        return Rng(a.sheet, min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2), max(a.c2, b.c2))

    # -------------------------------------------------------------- 配列への展開
    def map1(self, v, f: Callable, op: str = ""):
        if isinstance(v, Rng):
            v = self.materialize(v) if self.array else self.intersect(v)
        if isinstance(v, Arr):
            array_room(self, v.rows, v.cols)
            poly = None
            if v.poly is not None:  # 単項の - と %：多項式のまま
                poly = -v.poly if op == "-" else (v.poly.scaled(Fraction(1, 100)) if op == "%" else None)
            if want_lazy(v, (v.rows, v.cols)):
                at = v.at
                return make_lazy(self, v.rows, v.cols, lambda i, j: _safe(f, at(i, j)), poly)
            out = Arr([[_safe(f, x) for x in row] for row in v.data])
            out.poly = poly
            return out
        return _safe(f, v)

    def map2(self, a, b, f: Callable, op: str = ""):
        if isinstance(a, Rng):
            a = self.materialize(a) if self.array else self.intersect(a)
        if isinstance(b, Rng):
            b = self.materialize(b) if self.array else self.intersect(b)
        if not isinstance(a, Arr) and not isinstance(b, Arr):
            return _safe2(f, a, b)
        A = a if isinstance(a, Arr) else Arr([[a]])
        B = b if isinstance(b, Arr) else Arr([[b]])
        rows = max(A.rows, B.rows)
        cols = max(A.cols, B.cols)
        array_room(self, rows, cols)
        poly = poly_binary(op, a, b, rows, cols) if op and (A.poly is not None or B.poly is not None) else None

        if is_big(rows, cols):
            ar, ac, br, bc = A.rows, A.cols, B.rows, B.cols
            aat, bat = A.at, B.at
            na = err("#N/A")

            def cell(r, c):  # 1 行・1 列の配列は広げる。足りない所は #N/A
                rr, cc = (0 if ar == 1 else r), (0 if ac == 1 else c)
                x = aat(rr, cc) if rr < ar and cc < ac else na
                rr, cc = (0 if br == 1 else r), (0 if bc == 1 else c)
                y = bat(rr, cc) if rr < br and cc < bc else na
                return _safe2(f, x, y)

            return make_lazy(self, rows, cols, cell, poly)

        def pick(X: Arr, r, c):
            rr = 0 if X.rows == 1 else r
            cc = 0 if X.cols == 1 else c
            if rr >= X.rows or cc >= X.cols:
                return err("#N/A")
            return X.data[rr][cc]

        out = Arr([[_safe2(f, pick(A, r, c), pick(B, r, c)) for c in range(cols)] for r in range(rows)])
        out.poly = poly
        return out


def _safe(f, x):
    if isinstance(x, ExcelError):
        return x
    try:
        return f(x)
    except ExcelError as e:
        return e
    except ZeroDivisionError:
        return err("#DIV/0!")
    except OverflowError:
        return err("#NUM!")


def _safe2(f, a, b):
    if isinstance(a, ExcelError):
        return a
    if isinstance(b, ExcelError):
        return b
    try:
        return f(a, b)
    except ExcelError as e:
        return e
    except ZeroDivisionError:
        return err("#DIV/0!")
    except OverflowError:
        return err("#NUM!")


def _div(a, b):
    y = to_number(b)
    if y == 0:
        raise err("#DIV/0!")
    return to_number(a) / y


def _pow(a, b):
    x, y = to_number(a), to_number(b)
    if x == 0 and y == 0:
        raise err("#NUM!")
    if x < 0 and y != int(y):
        raise err("#NUM!")
    try:
        r = x ** y
    except ZeroDivisionError:
        raise err("#DIV/0!")
    if isinstance(r, complex):
        raise err("#NUM!")
    return float(r)


def _concat(a, b):
    x, y = to_text(a), to_text(b)
    if len(x) + len(y) > TEXT_MAX:
        raise err(VALUE)
    return x + y


_BIN: dict[str, Callable] = {
    "+": lambda a, b: to_number(a) + to_number(b),
    "-": lambda a, b: to_number(a) - to_number(b),
    "*": lambda a, b: to_number(a) * to_number(b),
    "/": _div,
    "^": _pow,
    "&": _concat,
    "=": lambda a, b: compare(a, b) == 0,
    "<>": lambda a, b: compare(a, b) != 0,
    "<": lambda a, b: compare(a, b) < 0,
    ">": lambda a, b: compare(a, b) > 0,
    "<=": lambda a, b: compare(a, b) <= 0,
    ">=": lambda a, b: compare(a, b) >= 0,
}
