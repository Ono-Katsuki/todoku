"""数式の「形」：行・列のずれ（R1C1 の相対）で同じ数式を、1 つの形にまとめる。

同じ列を 10 万行埋める数式は、番地の数字だけが行ごとに違う。数字を # に置き換えた字（型紙）が同じで、行に連動する数字と
行の差が同じなら、構文解析をしなくても同じ形だと分かる。解析は型紙ごとに 1 度だけなので、セルの数に比例する作業は
「字を 1 回なぞる」だけで済み、セルごとの控えを持たない。
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from ..formula.parser import (
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
    _BINARY_PREC,
    _PERCENT_PREC,
    _RANGE_PREC,
    _UNARY_PREC,
    _q,
    num_to_col,
    parse,
    ref_text,
    tokenize,
    unparse,
)

MAX_ROW = 1048576
_DIG = re.compile(r"\d+")


def render(node: Node, ref=ref_text) -> str:
    """構文木を字に戻す（かっこは、意味を変えないものだけ）。ref：参照の書き方。"""

    def prec(n: Node) -> int:
        if isinstance(n, Binary):
            return _RANGE_PREC if n.op == ":" else _BINARY_PREC[n.op]
        if isinstance(n, Unary):
            return _UNARY_PREC
        if isinstance(n, Postfix):
            return _PERCENT_PREC
        return 99

    def u(n: Node) -> str:
        if isinstance(n, RefNode):
            return ref(n)
        if isinstance(n, Func):
            return f"{n.name}({','.join(u(a) for a in n.args)})"
        if isinstance(n, Binary):
            p = prec(n)
            a, b = u(n.left), u(n.right)
            if prec(n.left) < p:
                a = f"({a})"
            if prec(n.right) <= p:
                b = f"({b})"
            return f"{a}{n.op}{b}"
        if isinstance(n, Unary):
            a = u(n.operand)
            return n.op + (f"({a})" if prec(n.operand) < _UNARY_PREC else a)
        if isinstance(n, Postfix):
            a = u(n.operand)
            return (f"({a})" if prec(n.operand) < _PERCENT_PREC else a) + n.op
        if isinstance(n, ArrayLit):
            return "{" + ";".join(",".join(u(x) for x in row) for row in n.rows) + "}"
        if isinstance(n, (Num, Str, Bool, Err, Missing, NameNode, StructRef)):
            return unparse(n)
        raise TypeError(n)

    return u(node)


def r1c1(n: RefNode, r: int, c: int) -> str:
    """参照を R1C1 の書き方にする（(r, c) から見た相対）。"""

    def row(v, a):
        return f"R{v}" if a else ("R" if v == r else f"R[{v - r}]")

    def col(v, a):
        return f"C{v}" if a else ("C" if v == c else f"C[{v - c}]")

    pre = f"{_q(n.sheet)}!" if n.sheet else ""
    if n.kind == "cols":
        return pre + col(n.c1, n.abs_c1) + (":" + col(n.c2, n.abs_c2) if (n.c1, n.abs_c1) != (n.c2, n.abs_c2) else "")
    if n.kind == "rows":
        return pre + row(n.r1, n.abs_r1) + (":" + row(n.r2, n.abs_r2) if (n.r1, n.abs_r1) != (n.r2, n.abs_r2) else "")
    a = row(n.r1, n.abs_r1) + col(n.c1, n.abs_c1)
    if n.kind == "cell":
        return pre + a
    return pre + a + ":" + row(n.r2, n.abs_r2) + col(n.c2, n.abs_c2)


def _off(name: str, d: int) -> str:
    return "{" + name + (f"{d:+d}" if d else "") + "}"


def pattern(n: RefNode, r: int, c: int, many_rows: bool, many_cols: bool) -> str:
    """参照を、行・列に連動する所だけ {row}・{col} で書く（1 行だけの範囲なら、行は番地のまま）。"""

    def row(v, a):
        if a:
            return f"${v}"
        return _off("row", v - r) if many_rows else str(v)

    def col(v, a):
        if a:
            return f"${num_to_col(v)}"
        return _off("col", v - c) if many_cols else num_to_col(v)

    pre = f"{_q(n.sheet)}!" if n.sheet else ""
    if n.kind == "cols":
        return pre + col(n.c1, n.abs_c1) + ":" + col(n.c2, n.abs_c2)
    if n.kind == "rows":
        return pre + row(n.r1, n.abs_r1) + ":" + row(n.r2, n.abs_r2)
    a = col(n.c1, n.abs_c1) + row(n.r1, n.abs_r1)
    if n.kind == "cell":
        return pre + a
    return pre + a + ":" + col(n.c2, n.abs_c2) + row(n.r2, n.abs_r2)


@dataclass
class Shape:
    """1 つの形：はじめて見たセル（r, c）から見た構文木。"""

    id: int
    r: int
    c: int
    text: str  # そのセルの数式（ファイルに書かれたまま）
    ast: Node | None
    r1c1: str  # 正規形（R1C1）
    error: str | None = None
    array: bool = False  # 配列数式（1 つの数式が範囲を計算する。形は親のセルごと）

    def text_at(self, r: int, c: int) -> str:
        """(r, c) のセルでの数式（相対の参照をずらす）。"""
        if self.ast is None or (r == self.r and c == self.c):
            return self.text
        dr, dc = r - self.r, c - self.c
        return render(self.ast, lambda n: ref_text(n.shifted(dr, dc)))


class Shaper:
    """数式の字 → 形の番号。型紙ごとに 1 度だけ解析する。"""

    def __init__(self):
        self.shapes: list[Shape] = []
        self.by_canon: dict[str, int] = {}
        self._plates: dict[str, tuple[bool, ...] | None] = {}  # 型紙 → 数字ごとの「行に連動するか」（読めない数式は None）
        self._keys: dict[tuple, int] = {}
        self._ast: dict[int, int] = {}  # 共有数式の親の構文木 → 形

    def new(self, text: str, r: int, c: int, ast: Node | None, error: str | None, array: bool = False) -> int:
        if ast is None:
            shown = text
            canon = "!" + text  # 読めない数式は、字がまったく同じものだけを同じ形にする
        else:
            shown = canon = render(ast, lambda n: r1c1(n, r, c))
        if array:
            canon = f"{{{canon}}}@{r},{c}"  # 配列数式は、親のセルごとに別の規則
        got = self.by_canon.get(canon)
        if got is not None:
            return got
        sid = len(self.shapes)
        self.shapes.append(Shape(sid, r, c, text, ast, shown, error, array))
        self.by_canon[canon] = sid
        return sid

    def _parse(self, text: str, r: int, c: int, array: bool = False) -> int:
        try:
            return self.new(text, r, c, parse(text), None, array)
        except Exception as e:  # noqa: BLE001  読めない数式は、読めないと規則に書く
            return self.new(text, r, c, None, f"{type(e).__name__}: {e}", array)

    def _plate(self, text: str) -> tuple[bool, ...] | None:
        """数字のならびごとに、行に連動する（$ のない行番号）か。"""
        try:
            toks = tokenize(text)
        except Exception:  # noqa: BLE001
            return None
        moving: set[int] = set()
        for i, t in enumerate(toks):
            if t.kind != "ref" or not isinstance(t.value, RefNode) or t.value.kind == "cols":
                continue
            end = toks[i + 1].pos if i + 1 < len(toks) else len(text)
            start = text.rfind("!", t.pos, end) + 1 or t.pos
            for m in _DIG.finditer(text, start, end):
                if text[m.start() - 1:m.start()] != "$":
                    moving.add(m.start())
        return tuple(m.start() in moving for m in _DIG.finditer(text))

    def of_text(self, text: str, r: int, c: int) -> int:
        """数式の字（= なし）→ 形の番号。"""
        plate_key = _DIG.sub("#", text)
        plate = self._plates.get(plate_key, False)
        if plate is False:
            plate = self._plates[plate_key] = self._plate(text)
        if plate is None:
            return self._parse(text, r, c)
        runs = _DIG.findall(text)
        if len(runs) != len(plate):
            return self._parse(text, r, c)
        ok = True
        key: list = [plate_key, c]
        for moving, d in zip(plate, runs):
            if moving:
                v = int(d)
                if not 1 <= v <= MAX_ROW:
                    ok = False
                    break
                key.append(v - r)
            else:
                key.append(d)
        if not ok:
            return self._parse(text, r, c)
        k = tuple(key)
        got = self._keys.get(k)
        if got is None:
            got = self._keys[k] = self._parse(text, r, c)
        return got

    def of_array(self, text: str, r: int, c: int) -> int:
        return self._parse(text, r, c, array=True)

    def of_ast(self, ast: Node, r0: int, c0: int) -> int:
        """共有数式の子：親の構文木と、親のセル。"""
        got = self._ast.get(id(ast))
        if got is None:
            got = self._ast[id(ast)] = self.new(unparse(ast), r0, c0, ast, None)
        return got
