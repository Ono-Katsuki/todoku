"""Excelの数式を構文木にする。

数式は「元の文字列」を正本として持ち、構文木は計算と説明のためだけに使う。
共有数式（t="shared"）は親の構文木に行・列のずれ（dr, dc）を与えて使い回す。
"""

from __future__ import annotations

import functools
import re
from dataclasses import dataclass, field

from .values import ERROR_CODES, Unsupported

MAX_ROW = 1048576
MAX_COL = 16384


def col_to_num(s: str) -> int:
    n = 0
    for ch in s.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


@functools.lru_cache(maxsize=1 << 15)
def num_to_col(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def addr(r: int, c: int) -> str:
    return f"{num_to_col(c)}{r}"


_CELL_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
_COL_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})$")
_ROW_RE = re.compile(r"^(\$?)(\d+)$")


@functools.lru_cache(maxsize=1 << 18)
def parse_cell(s: str) -> tuple[int, int]:
    """「B12」→ (12, 2)。同じ番地は何度も出てくるので、答えを控える（答えは変わらない値の組）。"""
    m = _CELL_RE.match(s)
    if not m:
        raise ValueError(s)
    return int(m[4]), col_to_num(m[2])


# ---------------------------------------------------------------- 構文木
@dataclass
class Node:
    pass


@dataclass
class Num(Node):
    value: float
    text: str


@dataclass
class Str(Node):
    value: str


@dataclass
class Bool(Node):
    value: bool


@dataclass
class Err(Node):
    code: str


@dataclass
class ArrayLit(Node):
    rows: list[list[Node]]


@dataclass
class RefNode(Node):
    """セルまたは範囲の参照。行・列ごとに絶対($)かどうかを持つ。

    列だけ（A:A）のときは r1=1, r2=MAX_ROW、行だけ（1:1）のときは c1=1, c2=MAX_COL。
    """

    sheet: str | None
    r1: int
    c1: int
    r2: int
    c2: int
    abs_r1: bool = False
    abs_c1: bool = False
    abs_r2: bool = False
    abs_c2: bool = False
    kind: str = "cell"  # cell / range / cols / rows

    def shifted(self, dr: int, dc: int) -> "RefNode":
        if dr == 0 and dc == 0:
            return self
        r1 = self.r1 if self.abs_r1 or self.kind == "cols" else self.r1 + dr
        r2 = self.r2 if self.abs_r2 or self.kind == "cols" else self.r2 + dr
        c1 = self.c1 if self.abs_c1 or self.kind == "rows" else self.c1 + dc
        c2 = self.c2 if self.abs_c2 or self.kind == "rows" else self.c2 + dc
        return RefNode(self.sheet, r1, c1, r2, c2, self.abs_r1, self.abs_c1, self.abs_r2, self.abs_c2, self.kind)


@dataclass
class NameNode(Node):
    name: str
    sheet: str | None = None


@dataclass
class Func(Node):
    name: str
    args: list[Node] = field(default_factory=list)


@dataclass
class Unary(Node):
    op: str
    operand: Node


@dataclass
class Postfix(Node):
    op: str  # '%'
    operand: Node


@dataclass
class Binary(Node):
    op: str
    left: Node
    right: Node


@dataclass
class Missing(Node):
    """省略された引数（=IF(A1,,1) など）。"""


@dataclass
class StructRef(Node):
    """構造化参照（テーブルの列を名前で指す参照）。Table1[金額]、Table1[[#This Row],[金額]] など。"""

    table: str | None  # None はその数式があるテーブル
    specials: tuple[str, ...] = ()  # "#All" "#Data" "#Headers" "#Totals" "#This Row"
    col1: str | None = None
    col2: str | None = None


_SPECIALS = {"#all": "#All", "#data": "#Data", "#headers": "#Headers", "#totals": "#Totals",
             "#this row": "#This Row", "#thisrow": "#This Row"}


def _unesc(x: str) -> str:
    # 列名の中の [ ] # ' は、前に ' を付けて書かれる
    return re.sub(r"'(.)", r"\1", x)


def _split_top(c: str) -> list[str]:
    """[a],[b] を、かっこの外の , で分ける（' のエスケープは飛ばす）。"""
    out, buf, depth, i = [], [], 0, 0
    while i < len(c):
        ch = c[i]
        if ch == "'" and i + 1 < len(c):
            buf.append(c[i:i + 2])
            i += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
        if ch == "," and depth == 0:
            out.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    out.append("".join(buf))
    return out


def _parse_struct(table: str | None, content: str) -> StructRef:
    c = content
    if c.strip() == "":
        return StructRef(table)
    if c.startswith("@"):
        rest = c[1:].strip()
        col = _unesc(rest[1:-1]) if rest.startswith("[") and rest.endswith("]") else _unesc(rest)
        return StructRef(table, ("#This Row",), col or None)
    if not c.lstrip().startswith("["):
        # 列名は空白も含めてそのまま（「Amount Claimed (£) 」のように末尾に空白のある列名もある）
        if c.strip().startswith("#") and c.strip().lower() in _SPECIALS:
            return StructRef(table, (_SPECIALS[c.strip().lower()],))
        return StructRef(table, (), _unesc(c))
    specials: list[str] = []
    cols: list[str] = []
    for item in _split_top(c.strip()):
        item = item.strip()
        parts = [x.strip() for x in re.split(r"\]\s*:\s*\[", item)] if "]:" in item.replace(" ", "") else [item]
        names = []
        for x in parts:
            x = x.strip()  # かっこの外の空白だけを除く
            if x.startswith("["):
                x = x[1:]
            if x.endswith("]") and not x.endswith("']"):
                x = x[:-1]
            names.append(x)
        for x in names:
            if x.startswith("#") and x.lower() in _SPECIALS:
                specials.append(_SPECIALS[x.lower()])
            elif x.startswith("@"):
                specials.append("#This Row")
                if x[1:]:
                    cols.append(_unesc(x[1:]))
            else:
                cols.append(_unesc(x))
    return StructRef(table, tuple(specials), cols[0] if cols else None, cols[1] if len(cols) > 1 else None)


def _read_bracket(s: str, i: int) -> tuple[str, int]:
    """s[i] の [ から、対応する ] までの中身と、その次の位置。"""
    depth, j = 0, i
    while j < len(s):
        ch = s[j]
        if ch == "'" and j + 1 < len(s):
            j += 2
            continue
        if ch == "[":
            depth += 1
        elif ch == "]":
            depth -= 1
            if depth == 0:
                return s[i + 1:j], j + 1
        j += 1
    raise SyntaxError("構造化参照の [ が閉じていません")


# ---------------------------------------------------------------- 字句解析
@dataclass
class Tok:
    kind: str  # num str bool err ref name func op lparen rparen comma semi lbrace rbrace
    value: object
    pos: int


_WORD_RE = re.compile(r"[A-Za-z_\\\u0080-￿$][\w.$\u0080-￿]*|\d[\w.$]*")
_NUM_RE = re.compile(r"(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?")
_SHEET_Q_RE = re.compile(r"'((?:[^']|'')+)'!")
_ROWRANGE_RE = re.compile(r"(\$?\d+):(\$?\d+)")


def _ref_from_parts(sheet, a: str, b: str | None) -> RefNode | None:
    """'A1', 'A1:B2', 'A:B', '1:2' を参照にする。参照でなければ None。"""
    ma = _CELL_RE.match(a)
    if b is None:
        if not ma:
            return None
        r, c = int(ma[4]), col_to_num(ma[2])
        if r > MAX_ROW or c > MAX_COL or r < 1:
            return None
        return RefNode(sheet, r, c, r, c, bool(ma[3]), bool(ma[1]), bool(ma[3]), bool(ma[1]), "cell")
    mb = _CELL_RE.match(b)
    if ma and mb:
        r1, c1, r2, c2 = int(ma[4]), col_to_num(ma[2]), int(mb[4]), col_to_num(mb[2])
        if max(r1, r2) > MAX_ROW or max(c1, c2) > MAX_COL:
            return None
        n = RefNode(sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2),
                    bool(ma[3]), bool(ma[1]), bool(mb[3]), bool(mb[1]), "range")
        return n
    ca, cb = _COL_RE.match(a), _COL_RE.match(b)
    if ca and cb:
        c1, c2 = col_to_num(ca[2]), col_to_num(cb[2])
        return RefNode(sheet, 1, min(c1, c2), MAX_ROW, max(c1, c2), True, bool(ca[1]), True, bool(cb[1]), "cols")
    ra, rb = _ROW_RE.match(a), _ROW_RE.match(b)
    if ra and rb:
        r1, r2 = int(ra[2]), int(rb[2])
        return RefNode(sheet, min(r1, r2), 1, max(r1, r2), MAX_COL, bool(ra[1]), True, bool(rb[1]), True, "rows")
    return None


def _first_token_text(s: str) -> str:
    """参照1つぶんの文字（A1、$A$1:$B$2、A:B、1:2）。"""
    m = re.match(r"\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+", s)
    if not m:
        raise Unsupported("外部ブックへの参照")
    return m[0]


def tokenize(src: str) -> list[Tok]:
    s = src
    i = 0
    toks: list[Tok] = []
    n = len(s)
    while i < n:
        ch = s[i]
        if ch in " \t\r\n":
            i += 1
            continue
        if ch == '"':
            j = i + 1
            buf = []
            while True:
                if j >= n:
                    raise SyntaxError("文字列が閉じていません")
                if s[j] == '"':
                    if j + 1 < n and s[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                buf.append(s[j])
                j += 1
            toks.append(Tok("str", "".join(buf), i))
            i = j + 1
            continue
        if ch == "#":
            for code in sorted(ERROR_CODES, key=len, reverse=True):
                if s.startswith(code, i):
                    toks.append(Tok("err", code, i))
                    i += len(code)
                    break
            else:
                raise SyntaxError(f"不明なエラー値: {s[i:i+8]}")
            # #REF!A1 のような形は参照エラーとして扱う
            continue
        if ch == "[":
            mx = re.match(r"\[(\d+)\]([^!\[\]'\s(),]*)!", s[i:])
            if mx and mx[2]:
                # 外部ブックへの参照：シート名を "[1]Sheet1" として扱い、保存されている値を使う
                sheet = f"[{mx[1]}]{mx[2]}"
                i += mx.end()
                text = _first_token_text(s[i:])
                a, _, b = text.partition(":")
                ref = _ref_from_parts(sheet, a, b or None)
                if ref is None:
                    raise Unsupported("外部ブックへの参照")
                toks.append(Tok("ref", ref, i))
                i += len(text)
                continue
            if re.match(r"\[\d+\]", s[i:]):
                raise Unsupported("外部ブックへの参照（名前など）")
            # その数式があるテーブルの列 [列名]・[@列名]
            content, j = _read_bracket(s, i)
            toks.append(Tok("ref", _parse_struct(None, content), i))
            i = j
            continue
        # シート名付き参照
        sheet = None
        m = _SHEET_Q_RE.match(s, i)
        if m:
            sheet = m[1].replace("''", "'")
            i = m.end()
        else:
            m = _WORD_RE.match(s, i)
            if m and m.end() < n and s[m.end()] == "!":
                word = m[0]
                if ":" in word:
                    raise Unsupported("複数シートにまたがる参照")
                sheet = word
                i = m.end() + 1
            elif m and m.end() < n and s[m.end()] == ":" and _COL_RE.match(m[0]) is None and _CELL_RE.match(m[0]) is None:
                # Sheet1:Sheet3!A1（3D参照）
                m3 = re.match(r"[\w.\u0080-￿]+:[\w.\u0080-￿]+!", s[i:])
                if m3:
                    raise Unsupported("複数シートにまたがる参照")
        if sheet is not None or (i < n and (s[i].isalnum() or s[i] in "_$\\" or ord(s[i]) >= 0x80)):
            # 行範囲 1:3
            mr = _ROWRANGE_RE.match(s, i)
            if mr:
                ref = _ref_from_parts(sheet, mr[1], mr[2])
                if ref is not None:
                    toks.append(Tok("ref", ref, i))
                    i = mr.end()
                    continue
            if sheet is None and s[i].isdigit() or (s[i] == "." and i + 1 < n and s[i + 1].isdigit()):
                mn = _NUM_RE.match(s, i)
                toks.append(Tok("num", mn[0], i))
                i = mn.end()
                continue
            if sheet is not None and s.startswith("#REF!", i):
                toks.append(Tok("err", "#REF!", i))
                i += 5
                continue
            m = _WORD_RE.match(s, i)
            if not m:
                raise SyntaxError(f"読めない文字: {s[i:i+10]}")
            word = m[0]
            j = m.end()
            if sheet is None and j < n and s[j] == "[":
                # 構造化参照 Table1[列名]
                content, j2 = _read_bracket(s, j)
                toks.append(Tok("ref", _parse_struct(word, content), i))
                i = j2
                continue
            # A1:B2 / A:B
            if j < n and s[j] == ":":
                m2 = _WORD_RE.match(s, j + 1)
                if m2:
                    word2 = m2[0]
                    if word2.endswith("!") is False and (j + 1 + len(word2) >= n or s[j + 1 + len(word2)] != "!"):
                        ref = _ref_from_parts(sheet, word, word2)
                        if ref is not None:
                            toks.append(Tok("ref", ref, i))
                            i = m2.end()
                            continue
            # 関数
            k = j
            if k < n and s[k] == "(" and sheet is None:
                # Excel 2010 以降の関数は、ファイルの中では _xlfn. や _xlws. を付けて書かれる
                fname = re.sub(r"^(?i:_xlfn\.|_xlws\.)+", "", word)
                toks.append(Tok("func", fname.upper() if fname != word else word, i))
                i = k + 1
                continue
            ref = _ref_from_parts(sheet, word, None)
            if ref is not None:
                toks.append(Tok("ref", ref, i))
                i = j
                continue
            if sheet is None and word.upper() in ("TRUE", "FALSE"):
                toks.append(Tok("bool", word.upper() == "TRUE", i))
                i = j
                continue
            toks.append(Tok("name", (word, sheet), i))
            i = j
            continue
        if s.startswith("<=", i) or s.startswith(">=", i) or s.startswith("<>", i):
            toks.append(Tok("op", s[i:i + 2], i))
            i += 2
            continue
        if ch in "+-*/^&=<>%:":
            toks.append(Tok("op", ch, i))
            i += 1
            continue
        if ch == "(":
            toks.append(Tok("lparen", ch, i))
        elif ch == ")":
            toks.append(Tok("rparen", ch, i))
        elif ch == ",":
            toks.append(Tok("comma", ch, i))
        elif ch == ";":
            toks.append(Tok("semi", ch, i))
        elif ch == "{":
            toks.append(Tok("lbrace", ch, i))
        elif ch == "}":
            toks.append(Tok("rbrace", ch, i))
        else:
            raise SyntaxError(f"読めない文字: {ch!r}")
        i += 1
    return toks


# ---------------------------------------------------------------- 構文解析
_BINARY_PREC = {
    "=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1,
    "&": 2,
    "+": 3, "-": 3,
    "*": 4, "/": 4,
    "^": 5,
}
_UNARY_PREC = 6  # Excelでは -2^2 = 4（単項マイナスが ^ より強い）
_PERCENT_PREC = 7
_RANGE_PREC = 8


class Parser:
    def __init__(self, src: str):
        self.src = src
        self.toks = tokenize(src)
        self.i = 0

    def peek(self) -> Tok | None:
        return self.toks[self.i] if self.i < len(self.toks) else None

    def next(self) -> Tok:
        t = self.peek()
        if t is None:
            raise SyntaxError("数式が途中で終わっています")
        self.i += 1
        return t

    def parse(self) -> Node:
        node = self.expr(0)
        if self.peek() is not None:
            raise SyntaxError(f"余分な記号: {self.peek().value!r}")
        return node

    def expr(self, min_prec: int) -> Node:
        left = self.prefix()
        while True:
            t = self.peek()
            if t is None or t.kind != "op":
                break
            op = t.value
            if op == "%":
                if _PERCENT_PREC < min_prec:
                    break
                self.next()
                left = Postfix("%", left)
                continue
            if op == ":":
                if _RANGE_PREC < min_prec:
                    break
                self.next()
                right = self.expr(_RANGE_PREC + 1)
                left = Binary(":", left, right)
                continue
            prec = _BINARY_PREC.get(op)
            if prec is None or prec < min_prec:
                break
            self.next()
            # ^ も含めてExcelは左結合
            right = self.expr(prec + 1)
            left = Binary(op, left, right)
        return left

    def prefix(self) -> Node:
        t = self.next()
        k = t.kind
        if k == "op" and t.value in "+-":
            operand = self.expr(_UNARY_PREC)
            return Unary(t.value, operand)
        if k == "num":
            return Num(float(t.value), t.value)
        if k == "str":
            return Str(t.value)
        if k == "bool":
            return Bool(t.value)
        if k == "err":
            return Err(t.value)
        if k == "ref":
            return t.value
        if k == "name":
            name, sheet = t.value
            return NameNode(name, sheet)
        if k == "lparen":
            inner = self.expr(0)
            if self.peek() is not None and self.peek().kind == "comma":
                raise Unsupported("参照の結合（A1:A2,B1:B2）")
            self._expect("rparen")
            return inner
        if k == "func":
            return self.func(t.value)
        if k == "lbrace":
            return self.array()
        raise SyntaxError(f"予期しない記号: {t.value!r}")

    def _expect(self, kind: str) -> Tok:
        t = self.next()
        if t.kind != kind:
            raise SyntaxError(f"{kind} が必要な位置に {t.value!r}")
        return t

    def func(self, raw_name: str) -> Node:
        name = raw_name.upper()
        for p in ("_XLFN._XLWS.", "_XLFN.", "_XLWS.", "_XLL."):
            if name.startswith(p):
                name = name[len(p):]
        args: list[Node] = []
        if self.peek() is not None and self.peek().kind == "rparen":
            self.next()
            return Func(name, args)
        while True:
            t = self.peek()
            if t is not None and t.kind in ("comma", "rparen"):
                args.append(Missing())
            else:
                args.append(self.expr(0))
            t = self.next()
            if t.kind == "rparen":
                break
            if t.kind != "comma":
                raise SyntaxError(f"関数の引数の区切りが不正: {t.value!r}")
        return Func(name, args)

    def array(self) -> Node:
        rows: list[list[Node]] = [[]]
        while True:
            t = self.next()
            if t.kind == "op" and t.value == "-":
                t2 = self.next()
                rows[-1].append(Num(-float(t2.value), "-" + t2.value))
            elif t.kind == "num":
                rows[-1].append(Num(float(t.value), t.value))
            elif t.kind == "str":
                rows[-1].append(Str(t.value))
            elif t.kind == "bool":
                rows[-1].append(Bool(t.value))
            elif t.kind == "err":
                rows[-1].append(Err(t.value))
            else:
                raise SyntaxError("配列定数が不正です")
            t = self.next()
            if t.kind == "rbrace":
                break
            if t.kind == "semi":
                rows.append([])
            elif t.kind != "comma":
                raise SyntaxError("配列定数の区切りが不正です")
        return ArrayLit(rows)


def parse(src: str) -> Node:
    if src.startswith("="):
        src = src[1:]
    return _parse_cached(src)


@functools.lru_cache(maxsize=200_000)
def _parse_cached(src: str) -> Node:
    # 同じ数式（共有数式・入力規則・条件付き書式は同じ式がくり返し出る）は1回だけ解析する
    return Parser(src).parse()


# ---------------------------------------------------------------- 構文木を文字列に戻す
def _q(sheet: str) -> str:
    if re.fullmatch(r"[A-Za-z_\u0080-￿][\w.\u0080-￿]*", sheet) and not _CELL_RE.match(sheet):
        return sheet
    return "'" + sheet.replace("'", "''") + "'"


def ref_text(n: RefNode) -> str:
    pre = f"{_q(n.sheet)}!" if n.sheet else ""

    def cell(r, c, ar, ac):
        return f"{'$' if ac else ''}{num_to_col(c)}{'$' if ar else ''}{r}"

    if n.kind == "cols":
        return f"{pre}{'$' if n.abs_c1 else ''}{num_to_col(n.c1)}:{'$' if n.abs_c2 else ''}{num_to_col(n.c2)}"
    if n.kind == "rows":
        return f"{pre}{'$' if n.abs_r1 else ''}{n.r1}:{'$' if n.abs_r2 else ''}{n.r2}"
    a = cell(n.r1, n.c1, n.abs_r1, n.abs_c1)
    if n.kind == "cell":
        return pre + a
    return pre + a + ":" + cell(n.r2, n.c2, n.abs_r2, n.abs_c2)


def unparse(node: Node, dr: int = 0, dc: int = 0) -> str:
    def u(n: Node) -> str:
        if isinstance(n, Num):
            return n.text
        if isinstance(n, Str):
            return '"' + n.value.replace('"', '""') + '"'
        if isinstance(n, Bool):
            return "TRUE" if n.value else "FALSE"
        if isinstance(n, Err):
            return n.code
        if isinstance(n, RefNode):
            return ref_text(n.shifted(dr, dc))
        if isinstance(n, NameNode):
            return (f"{_q(n.sheet)}!" if n.sheet else "") + n.name
        if isinstance(n, Func):
            return f"{n.name}({','.join(u(a) for a in n.args)})"
        if isinstance(n, Unary):
            return n.op + u(n.operand)
        if isinstance(n, Postfix):
            return u(n.operand) + n.op
        if isinstance(n, Binary):
            return f"{u(n.left)}{n.op}{u(n.right)}" if n.op == ":" else f"({u(n.left)}{n.op}{u(n.right)})"
        if isinstance(n, Missing):
            return ""
        if isinstance(n, ArrayLit):
            return "{" + ";".join(",".join(u(x) for x in row) for row in n.rows) + "}"
        if isinstance(n, StructRef):
            def esc(x):
                return re.sub(r"([\[\]#'])", r"'\1", x)
            items = [f"[{sp}]" for sp in n.specials]
            if n.col1 is not None:
                items.append(f"[{esc(n.col1)}]" + (f":[{esc(n.col2)}]" if n.col2 is not None else ""))
            inner = ",".join(items)
            return (n.table or "") + (f"[{inner}]" if len(items) > 1 else (f"[{inner[1:-1]}]" if items else "[]"))
        raise TypeError(n)

    return u(node)


def iter_refs(node: Node):
    """構文木に含まれる参照と名前をすべて返す。"""
    stack = [node]
    while stack:
        n = stack.pop()
        if isinstance(n, (RefNode, NameNode, StructRef)):
            yield n
        elif isinstance(n, Func):
            stack.extend(n.args)
        elif isinstance(n, (Unary, Postfix)):
            stack.append(n.operand)
        elif isinstance(n, Binary):
            stack.extend((n.left, n.right))


def iter_funcs(node: Node):
    stack = [node]
    while stack:
        n = stack.pop()
        if isinstance(n, Func):
            yield n.name
            stack.extend(n.args)
        elif isinstance(n, (Unary, Postfix)):
            stack.append(n.operand)
        elif isinstance(n, Binary):
            stack.extend((n.left, n.right))
