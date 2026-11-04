"""Excelの値の型と、型変換の規則。"""

from __future__ import annotations

import math
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal


class ExcelError(Exception):
    """セルの値としてのエラー（#DIV/0! など）。例外としても投げられる。"""

    def __init__(self, code: str):
        super().__init__(code)
        self.code = code

    def __eq__(self, other):
        return isinstance(other, ExcelError) and other.code == self.code

    def __hash__(self):
        return hash(self.code)

    def __repr__(self):
        return self.code


DIV0 = "#DIV/0!"
NA = "#N/A"
NAME = "#NAME?"
NULL = "#NULL!"
NUM = "#NUM!"
REF = "#REF!"
VALUE = "#VALUE!"
ERROR_CODES = {DIV0, NA, NAME, NULL, NUM, REF, VALUE, "#GETTING_DATA", "#SPILL!", "#CALC!"}


def err(code: str) -> ExcelError:
    return ExcelError(code)


class Unsupported(Exception):
    """この計算エンジンが再現できない（＝検証できない）もの。"""


class Volatile(Exception):
    """TODAY() など、計算するたびに値が変わるもの。"""


@dataclass(frozen=True)
class Rng:
    """シート上の範囲。値は必要になったときに取り出す。"""

    sheet: str
    r1: int
    c1: int
    r2: int
    c2: int

    @property
    def rows(self) -> int:
        return self.r2 - self.r1 + 1

    @property
    def cols(self) -> int:
        return self.c2 - self.c1 + 1

    @property
    def is_cell(self) -> bool:
        return self.r1 == self.r2 and self.c1 == self.c2


@dataclass
class Arr:
    """配列（{1,2;3,4} や範囲を計算した結果）。"""

    data: list[list]

    @property
    def rows(self) -> int:
        return len(self.data)

    @property
    def cols(self) -> int:
        return len(self.data[0]) if self.data else 0


BLANK = None  # 空白セル

TEXT_MAX = 32767  # 1 つのセルの字の数の上限（Excel の仕様。これを超える字を返す数式は #VALUE!）
_ARRAY_SMALL = 65536  # これ以下の配列は数えない


def array_room(ev, rows: int, cols: int) -> None:
    """rows 行 cols 列の配列を、作る前に確かめる。シートの大きさを超える配列は #NUM!（Excel の仕様）。
    シートに入る大きさでも、limits.MAX_ARRAY_CELLS を超える配列と、1 回の計算で作る配列の合計が上限を超えるものは、
    値を作らずに「計算しなかった」と伝える（Unsupported。Excel で開いたときに計算される）。"""
    n = rows * cols
    if n <= _ARRAY_SMALL:
        return
    if rows > 1048576 or cols > 16384:
        raise ExcelError("#NUM!")
    from ..limits import MAX_ARRAY_CELLS, MAX_ARRAY_TOTAL

    if n > MAX_ARRAY_CELLS:
        raise Unsupported(f"配列が大きすぎます（{rows:,} 行 × {cols:,} 列）。この数式は計算せず、Excel で開いたときに計算されます")
    ctx = ev.ctx
    made = getattr(ctx, "array_cells", 0) + n
    if made > MAX_ARRAY_TOTAL:
        raise Unsupported(f"1 回の計算で作る配列のマスが {MAX_ARRAY_TOTAL:,} を超えます。この数式は計算せず、Excel で開いたときに計算されます")
    ctx.array_cells = made


# ---------------------------------------------------------------- 数値の扱い
def to_15(x: float) -> float:
    """Excelは有効数字15桁で値を扱う。その桁に丸める。"""
    if x == 0 or not math.isfinite(x):
        return x
    return float(f"{x:.15g}")


def dec15(x: float) -> Decimal:
    return Decimal(f"{x:.15g}")


def round_half_up(x: float, digits: int) -> float:
    """ExcelのROUND（0から遠い方へ四捨五入、15桁で丸めてから処理）。"""
    d = dec15(x)
    q = Decimal(1).scaleb(-digits)
    sign = -1 if d < 0 else 1
    r = (abs(d)).quantize(q, rounding=ROUND_HALF_UP) * sign
    return float(r)


def number_to_text(x: float) -> str:
    """数値を文字列に（& や TEXT 以外での変換。Excelの「標準」に近い）。"""
    if isinstance(x, bool):
        return "TRUE" if x else "FALSE"
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    s = f"{x:.15g}"
    if "e" in s:
        mant, exp = s.split("e")
        e = int(exp)
        return f"{mant}E{'+' if e >= 0 else '-'}{abs(e):02d}"
    return s


_NUMBER = __import__("re").compile(r"([+-]?)\s*((?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)")


def parse_number(s: str) -> float | None:
    t = s.strip().replace(",", "")
    if not t:
        return None
    pct = t.endswith("%")
    if pct:
        t = t[:-1].rstrip()
    # 数の形だけを数にする（Excel は符号の後ろの空白を許す：「- 0.89」）。Python の float だけが読む形
    # （「1_000」「inf」「nan」）は数にしない
    m = _NUMBER.fullmatch(t)
    if not m:
        return None
    v = float(m[1] + m[2])
    if not math.isfinite(v):
        return None
    return v / 100 if pct else v


def to_number(v) -> float:
    """算術演算での型変換。"""
    if isinstance(v, ExcelError):
        raise v
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        n = parse_number(v)
        if n is None:
            n = _parse_date_text(v)
        if n is None:
            raise err(VALUE)
        return n
    raise err(VALUE)


def _parse_date_text(s: str) -> float | None:
    import datetime as dt
    import re

    m = re.fullmatch(r"\s*(\d{4})[/-](\d{1,2})[/-](\d{1,2})\s*", s)
    if not m:
        return None
    try:
        d = dt.date(int(m[1]), int(m[2]), int(m[3]))
    except ValueError:
        return None
    from .dates import date_to_serial

    return float(date_to_serial(d))


def to_text(v) -> str:
    if isinstance(v, ExcelError):
        raise v
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return number_to_text(float(v))
    return str(v)


def to_bool(v) -> bool:
    if isinstance(v, ExcelError):
        raise v
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        u = v.upper()
        if u == "TRUE":
            return True
        if u == "FALSE":
            return False
        raise err(VALUE)
    raise err(VALUE)


def is_number(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def compare(a, b) -> int:
    """Excelの比較（数値 < 文字列 < 論理値、文字列は大文字小文字を区別しない）。"""
    if isinstance(a, ExcelError):
        raise a
    if isinstance(b, ExcelError):
        raise b
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)

    def rank(x):
        if isinstance(x, bool):
            return 2
        if isinstance(x, str):
            return 1
        return 0

    ra, rb = rank(a), rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        a, b = a.lower(), b.lower()
    elif ra == 0:
        a, b = to_15(float(a)), to_15(float(b))
    return (a > b) - (a < b)
