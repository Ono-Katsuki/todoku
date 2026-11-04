"""ワークシート関数。行政の様式でよく使われるものから実装する。

未実装の関数は Unsupported を投げ、そのセルは「検証できなかった」として記録される。
"""

from __future__ import annotations

import datetime as dt
import math
import re
from decimal import ROUND_CEILING, ROUND_DOWN, ROUND_FLOOR, ROUND_UP, Decimal
from typing import Callable

from .dates import date_to_serial, serial_parts, serial_to_date, weekday_sun0
from .parser import Func, Missing, Node, iter_funcs
from .values import DIV0, NA, NUM, REF, VALUE, Arr, ExcelError, Rng, Unsupported, compare, dec15, err, is_number, parse_number, round_half_up, to_bool, to_number, to_text

EAGER: dict[str, Callable] = {}
LAZY: dict[str, Callable] = {}
# 引数を配列として計算する関数（範囲同士の演算を要素ごとに行う）
ARRAY_ARGS = {"SUMPRODUCT"}
VOLATILE = {"TODAY", "NOW", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO"}


def fn(name: str, lazy: bool = False):
    def deco(f):
        for n in name.split():
            (LAZY if lazy else EAGER)[n] = f
        return f

    return deco


def call(ev, node: Func):
    name = node.name
    if name in ("TODAY", "NOW", "RAND", "RANDBETWEEN"):
        ev.volatile = True
    if name in LAZY:
        return LAZY[name](ev, node.args)
    f = EAGER.get(name)
    if f is None:
        raise Unsupported(f"未対応の関数: {name}")
    if name in ARRAY_ARGS and not ev.array:
        saved = ev.array
        ev.array = True
        try:
            args = [ev.ev(a) for a in node.args]
        finally:
            ev.array = saved
    else:
        args = [ev.ev(a) for a in node.args]
    if ev.array and name in _arrays.SCALAR_FUNCS:
        # 1つの値を受け取る関数に、配列数式の中で範囲・配列が渡されたら、要素ごとに計算する
        out = _arrays.broadcast(ev, f, args)
        if out is not None:
            return out
    try:
        return f(ev, *args)
    except ExcelError as e:
        return e
    except ZeroDivisionError:
        return err(DIV0)
    except (OverflowError, ValueError):
        return err(NUM)
    except TypeError as e:
        if "positional argument" in str(e) or "arguments" in str(e):
            return err(VALUE)
        raise


def S(ev, v):
    """引数を1つの値にする。"""
    return ev.scalar(v)


def num(ev, v) -> float:
    return to_number(S(ev, v))


def opt(ev, v, default):
    if v is None:
        return default
    return v


# ================================================================ 集計
def _numbers(ev, args, *, count_text_bool_direct=True):
    """SUMなどが対象にする数値。範囲内の文字列・論理値は無視、直接の引数は変換する。"""
    out = []
    for a in args:
        if isinstance(a, (Rng, Arr)):
            for x in ev.values(a):
                if isinstance(x, ExcelError):
                    raise x
                if is_number(x):
                    out.append(float(x))
        else:
            if a is None:
                continue
            out.append(to_number(a))
    return out


def _sum(xs):
    # Excelは逐次加算。結果は15桁で比較するのでfsumではなく単純加算にする
    s = 0.0
    for x in xs:
        s += x
    return s


@fn("SUM")
def f_sum(ev, *args):
    return _sum(_numbers(ev, args))


@fn("PRODUCT")
def f_product(ev, *args):
    p = 1.0
    xs = _numbers(ev, args)
    if not xs:
        return 0.0
    for x in xs:
        p *= x
    return p


@fn("AVERAGE")
def f_average(ev, *args):
    xs = _numbers(ev, args)
    if not xs:
        raise err(DIV0)
    return _sum(xs) / len(xs)


@fn("MIN")
def f_min(ev, *args):
    xs = _numbers(ev, args)
    return min(xs) if xs else 0.0


@fn("MAX")
def f_max(ev, *args):
    xs = _numbers(ev, args)
    return max(xs) if xs else 0.0


@fn("COUNT")
def f_count(ev, *args):
    n = 0
    for a in args:
        if isinstance(a, (Rng, Arr)):
            n += sum(1 for x in ev.values(a) if is_number(x))
        else:
            try:
                to_number(a)
                n += 1
            except ExcelError:
                pass
    return float(n)


@fn("COUNTA")
def f_counta(ev, *args):
    n = 0
    for a in args:
        if isinstance(a, (Rng, Arr)):
            n += sum(1 for x in ev.values(a) if x is not None)
        else:
            n += 1
    return float(n)


@fn("COUNTBLANK")
def f_countblank(ev, rng):
    return float(sum(1 for x in ev.values(rng) if x is None or x == ""))


def _large(ev, rng, k):
    xs = sorted(_numbers(ev, [rng if isinstance(rng, (Rng, Arr)) else Arr([[rng]])]), reverse=True)
    i = int(num(ev, k))
    if i < 1 or i > len(xs):
        raise err(NUM)
    return xs[i - 1]


def _small(ev, rng, k):
    xs = sorted(_numbers(ev, [rng if isinstance(rng, (Rng, Arr)) else Arr([[rng]])]))
    i = int(num(ev, k))
    if i < 1 or i > len(xs):
        raise err(NUM)
    return xs[i - 1]


EAGER["LARGE"] = _large
EAGER["SMALL"] = _small


def _var(xs: list[float], sample: bool) -> float:
    n = len(xs)
    if n < (2 if sample else 1):
        raise err(DIV0)
    m = _sum(xs) / n
    return _sum([(x - m) ** 2 for x in xs]) / (n - 1 if sample else n)


@fn("VAR VAR.S")
def f_var(ev, *args):
    return _var(_numbers(ev, args), True)


@fn("VARP VAR.P")
def f_varp(ev, *args):
    return _var(_numbers(ev, args), False)


@fn("STDEV STDEV.S")
def f_stdev(ev, *args):
    return math.sqrt(_var(_numbers(ev, args), True))


@fn("STDEVP STDEV.P")
def f_stdevp(ev, *args):
    return math.sqrt(_var(_numbers(ev, args), False))


_SUBTOTAL = {1: "AVERAGE", 2: "COUNT", 3: "COUNTA", 4: "MAX", 5: "MIN", 6: "PRODUCT", 7: "STDEV", 8: "STDEVP",
             9: "SUM", 10: "VAR", 11: "VARP"}


@fn("SUBTOTAL")
def f_subtotal(ev, kind, *refs):
    """SUBTOTAL：範囲の中の SUBTOTAL の結果は数えない。101〜111 は手で隠した行も除く。
    オートフィルターで隠れた行は、どちらでも除く。"""
    k = int(num(ev, kind))
    hide_manual = k > 100
    base = k - 100 if hide_manual else k
    name = _SUBTOTAL.get(base)
    if name is None or not refs:
        raise err(VALUE)
    wb = getattr(ev.ctx, "wb", None)
    rows: list[list] = []
    for ref in refs:
        if not isinstance(ref, Rng):
            raise err(VALUE)
        rng = ev.clip(ref)
        sh = wb.sheet(rng.sheet) if wb is not None and wb.has_sheet(rng.sheet) else None
        for r in range(rng.r1, rng.r2 + 1):
            if sh is not None and r in sh.hidden_rows:
                filtered = sh.autofilter is not None and sh.autofilter[0] < r <= sh.autofilter[2]
                if hide_manual or filtered:
                    continue
            for c in range(rng.c1, rng.c2 + 1):
                cell = sh.cells.get((r, c)) if sh is not None else None
                if cell is not None and cell.ast is not None and any(
                        f in ("SUBTOTAL", "AGGREGATE") for f in iter_funcs(cell.ast)):
                    continue  # 入れ子の小計は数えない（二重に数えないため）
                rows.append([ev.ctx.cell_value(rng.sheet, r, c)])
    return EAGER[name](ev, Arr(rows or [[None]]))


@fn("SUMPRODUCT")
def f_sumproduct(ev, *args):
    grids = [ev.grid(a) for a in args]
    if not grids:
        raise err(VALUE)
    r, c = grids[0].rows, grids[0].cols
    for g in grids:
        if g.rows != r or g.cols != c:
            raise err(VALUE)
    total = 0.0
    for i in range(r):
        for j in range(c):
            p = 1.0
            for g in grids:
                x = g.data[i][j]
                if isinstance(x, ExcelError):
                    raise x
                p *= float(x) if is_number(x) else 0.0
            total += p
    return total


# ---------------------------------------------------------------- 条件
_WILD_CACHE: dict[str, re.Pattern] = {}


def _wild(pat: str) -> re.Pattern:
    p = _WILD_CACHE.get(pat)
    if p is None:
        out = []
        i = 0
        while i < len(pat):
            ch = pat[i]
            if ch == "~" and i + 1 < len(pat):
                out.append(re.escape(pat[i + 1]))
                i += 2
                continue
            out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
            i += 1
        p = re.compile("^" + "".join(out) + "$", re.I | re.S)
        _WILD_CACHE[pat] = p
    return p


def _num_eq(n: float):
    # 数値の条件は、数値に見える文字列にも一致する
    def p(x):
        if is_number(x):
            return compare(x, n) == 0
        if isinstance(x, str):
            v = parse_number(x)
            return v is not None and compare(v, n) == 0
        return False

    return p


def criteria(crit) -> Callable[[object], bool]:
    """COUNTIF/SUMIF の条件を判定関数にする。"""
    if isinstance(crit, ExcelError):
        # 条件がエラー値でも、エラーにはならない：同じエラー値のセルだけが当てはまる（Excelの挙動）
        return lambda x: isinstance(x, ExcelError) and x.code == crit.code
    if crit is None:
        # 空のセルを条件にすると 0 と一致する（Excelの挙動）
        return lambda x: is_number(x) and x == 0
    if isinstance(crit, bool):
        return lambda x: isinstance(x, bool) and x == crit
    if is_number(crit):
        return _num_eq(float(crit))
    s = str(crit)
    m = re.match(r"^(<=|>=|<>|<|>|=)?(.*)$", s, re.S)
    op, rhs = m[1] or "", m[2]
    n = parse_number(rhs)
    if n is None and rhs.upper() in ("TRUE", "FALSE"):
        b = rhs.upper() == "TRUE"
        if op in ("", "="):
            return lambda x: isinstance(x, bool) and x == b
        if op == "<>":
            return lambda x: not (isinstance(x, bool) and x == b)
    if op in ("", "="):
        if n is not None:
            return _num_eq(n)
        if rhs == "":
            return (lambda x: x is None or x == "") if op == "=" else (lambda x: x is None or x == "")
        if any(ch in rhs for ch in "*?~"):
            p = _wild(rhs)
            return lambda x: isinstance(x, str) and bool(p.match(x))
        low = rhs.lower()
        return lambda x: isinstance(x, str) and x.lower() == low
    if op == "<>":
        if rhs == "":
            return lambda x: x is not None and x != ""
        if n is not None:
            return lambda x: not (is_number(x) and compare(x, n) == 0)
        if any(ch in rhs for ch in "*?~"):
            p = _wild(rhs)
            return lambda x: not (isinstance(x, str) and p.match(x))
        low = rhs.lower()
        return lambda x: not (isinstance(x, str) and x.lower() == low)
    # 大小比較：数値は数値と、文字列は文字列と比べる
    if n is not None:
        def cmp_num(x, n=n, op=op):
            if not is_number(x):
                return False
            c = compare(x, n)
            return {"<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]

        return cmp_num

    def cmp_str(x, op=op, rhs=rhs):
        if not isinstance(x, str):
            return False
        c = compare(x, rhs)
        return {"<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]

    return cmp_str


def _same_shape(ev, base: Rng | Arr, other):
    """SUMIF の合計範囲は、条件範囲と同じ大きさに広げて使う。"""
    if isinstance(other, Rng) and isinstance(base, Rng):
        return ev.grid(Rng(other.sheet, other.r1, other.c1, other.r1 + base.rows - 1, other.c1 + base.cols - 1))
    return ev.grid(other)


def _flat(g: Arr):
    return [x for row in g.data for x in row]


def _lift(*crit_pos):
    """配列として計算している中で、条件に範囲や配列が渡されたら、その要素ごとに計算して配列で返す
    （SUMPRODUCT の中の COUNTIFS(A1:A9,A1:A9) など。Excelの動的配列と同じ）。"""

    def deco(f):
        def g(ev, *args):
            multi = [i for i in crit_pos if i < len(args) and isinstance(args[i], (Rng, Arr))
                     and (ev.grid(args[i]).rows * ev.grid(args[i]).cols > 1)]
            if not ev.array or not multi:
                return f(ev, *args)
            grids = {i: ev.grid(args[i]) for i in multi}
            rows = max(g_.rows for g_ in grids.values())
            cols = max(g_.cols for g_ in grids.values())
            out = []
            for r in range(rows):
                line = []
                for c in range(cols):
                    a = list(args)
                    for i, g_ in grids.items():
                        rr = 0 if g_.rows == 1 else r
                        cc = 0 if g_.cols == 1 else c
                        a[i] = g_.data[rr][cc] if rr < g_.rows and cc < g_.cols else err(NA)
                    try:
                        line.append(f(ev, *a))
                    except ExcelError as e:
                        line.append(e)
                out.append(line)
            return Arr(out)

        return g

    return deco


@fn("COUNTIF")
@_lift(1)
def f_countif(ev, rng, crit):
    if not isinstance(rng, (Rng, Arr)):
        raise err(VALUE)
    p = criteria(S(ev, crit))
    return float(sum(1 for x in ev.values(rng) if p(x)))


@fn("COUNTIFS")
@_lift(*range(1, 254, 2))
def f_countifs(ev, *args):
    if len(args) % 2:
        raise err(VALUE)
    masks = None
    for i in range(0, len(args), 2):
        vals = ev.values(args[i])
        p = criteria(S(ev, args[i + 1]))
        m = [p(x) for x in vals]
        if masks is None:
            masks = m
        elif len(m) != len(masks):
            raise err(VALUE)
        else:
            masks = [a and b for a, b in zip(masks, m)]
    return float(sum(masks or []))


@fn("SUMIF")
@_lift(1)
def f_sumif(ev, rng, crit, sum_rng=None):
    if not isinstance(rng, (Rng, Arr)):
        raise err(VALUE)
    p = criteria(S(ev, crit))
    keys = _flat(ev.grid(rng))
    vals = _flat(_same_shape(ev, rng, sum_rng)) if sum_rng is not None else keys
    total = 0.0
    for k, v in zip(keys, vals):
        if p(k):
            if isinstance(v, ExcelError):
                raise v
            if is_number(v):
                total += float(v)
    return total


@fn("AVERAGEIF")
def f_averageif(ev, rng, crit, avg_rng=None):
    p = criteria(S(ev, crit))
    keys = _flat(ev.grid(rng))
    vals = _flat(_same_shape(ev, rng, avg_rng)) if avg_rng is not None else keys
    xs = [float(v) for k, v in zip(keys, vals) if p(k) and is_number(v)]
    if not xs:
        raise err(DIV0)
    return _sum(xs) / len(xs)


def _ifs_mask(ev, pairs):
    mask = None
    for rng, crit in pairs:
        vals = ev.values(rng)
        p = criteria(S(ev, crit))
        m = [p(x) for x in vals]
        if mask is None:
            mask = m
        elif len(m) != len(mask):
            raise err(VALUE)
        else:
            mask = [a and b for a, b in zip(mask, m)]
    return mask


@fn("SUMIFS")
@_lift(*range(2, 255, 2))
def f_sumifs(ev, sum_rng, *args):
    if len(args) % 2:
        raise err(VALUE)
    vals = ev.values(sum_rng)
    mask = _ifs_mask(ev, [(args[i], args[i + 1]) for i in range(0, len(args), 2)])
    if len(mask) != len(vals):
        raise err(VALUE)
    total = 0.0
    for m, v in zip(mask, vals):
        if m:
            if isinstance(v, ExcelError):
                raise v
            if is_number(v):
                total += float(v)
    return total


def _xifs(kind):
    def f(ev, target, *args):
        if len(args) % 2:
            raise err(VALUE)
        vals = ev.values(target)
        mask = _ifs_mask(ev, [(args[i], args[i + 1]) for i in range(0, len(args), 2)])
        xs = [float(v) for m, v in zip(mask, vals) if m and is_number(v)]
        if kind == "AVERAGEIFS":
            if not xs:
                raise err(DIV0)
            return _sum(xs) / len(xs)
        if not xs:
            return 0.0
        return max(xs) if kind == "MAXIFS" else min(xs)

    return f


for _k in ("MAXIFS", "MINIFS", "AVERAGEIFS"):
    EAGER[_k] = _xifs(_k)


# ================================================================ 端数処理と数学
def _digits(ev, d) -> int:
    x = num(ev, d) if d is not None else 0.0
    return int(math.trunc(x))


@fn("ROUND")
def f_round(ev, x, d=None):
    return round_half_up(num(ev, x), _digits(ev, d))


def _round_mode(x: float, digits: int, mode) -> float:
    q = Decimal(1).scaleb(-digits)
    return float(dec15(x).quantize(q, rounding=mode))


@fn("ROUNDDOWN")
def f_rounddown(ev, x, d=None):
    return _round_mode(num(ev, x), _digits(ev, d), ROUND_DOWN)


@fn("ROUNDUP")
def f_roundup(ev, x, d=None):
    return _round_mode(num(ev, x), _digits(ev, d), ROUND_UP)


@fn("TRUNC")
def f_trunc(ev, x, d=None):
    return _round_mode(num(ev, x), _digits(ev, d), ROUND_DOWN)


@fn("INT")
def f_int(ev, x):
    return float(dec15(num(ev, x)).to_integral_value(rounding=ROUND_FLOOR))


@fn("ABS")
def f_abs(ev, x):
    return abs(num(ev, x))


@fn("SIGN")
def f_sign(ev, x):
    v = num(ev, x)
    return float((v > 0) - (v < 0))


@fn("MOD")
def f_mod(ev, a, b):
    x, y = num(ev, a), num(ev, b)
    if y == 0:
        raise err(DIV0)
    return float(dec15(x - y * math.floor(x / y)))


@fn("QUOTIENT")
def f_quotient(ev, a, b):
    x, y = num(ev, a), num(ev, b)
    if y == 0:
        raise err(DIV0)
    return float(math.trunc(x / y))


@fn("POWER")
def f_power(ev, a, b):
    from .evaluator import _pow

    return _pow(S(ev, a), S(ev, b))


@fn("SQRT")
def f_sqrt(ev, x):
    v = num(ev, x)
    if v < 0:
        raise err(NUM)
    return math.sqrt(v)


@fn("PI")
def f_pi(ev):
    return math.pi


def _mround_like(x: float, sig: float, mode) -> float:
    if sig == 0:
        return 0.0
    q = dec15(x) / dec15(sig)
    return float(q.to_integral_value(rounding=mode) * dec15(sig))


@fn("CEILING CEILING.MATH CEILING.PRECISE ISO.CEILING")
def f_ceiling(ev, x, sig=None, mode=None):
    v = num(ev, x)
    s = num(ev, sig) if sig is not None else 1.0
    if v > 0 and s < 0:
        raise err(NUM)
    if s == 0:
        return 0.0
    return _mround_like(v, abs(s) if v >= 0 or s > 0 else s, ROUND_CEILING if s > 0 or v >= 0 else ROUND_UP)


@fn("FLOOR FLOOR.MATH FLOOR.PRECISE")
def f_floor(ev, x, sig=None, mode=None):
    v = num(ev, x)
    s = num(ev, sig) if sig is not None else 1.0
    if v > 0 and s < 0:
        raise err(NUM)
    if s == 0:
        if sig is None:
            return 0.0
        raise err(DIV0)
    return _mround_like(v, abs(s) if v >= 0 or s > 0 else s, ROUND_FLOOR if s > 0 or v >= 0 else ROUND_DOWN)


@fn("MROUND")
def f_mround(ev, x, m):
    v, s = num(ev, x), num(ev, m)
    if s == 0:
        return 0.0
    if (v > 0 > s) or (v < 0 < s):
        raise err(NUM)
    from decimal import ROUND_HALF_UP

    return _mround_like(v, s, ROUND_HALF_UP)


# ================================================================ 論理
@fn("LET", lazy=True)
def f_let(ev, args: list[Node]):
    """LET(名前1, 値1, …, 計算)。名前は、ファイルの中では _xlpm.名前 と書かれる。"""
    from .parser import NameNode

    if len(args) < 3 or len(args) % 2 == 0:
        return err(VALUE)
    scope = dict(getattr(ev, "scope", None) or {})
    saved = getattr(ev, "scope", None)
    try:
        for i in range(0, len(args) - 1, 2):
            n = args[i]
            if not isinstance(n, NameNode):
                return err(VALUE)
            ev.scope = scope
            scope[_param(n.name)] = ev.ev(args[i + 1])
        ev.scope = scope
        return ev.ev(args[-1])
    finally:
        ev.scope = saved


def _param(name: str) -> str:
    return re.sub(r"^(?i:_xlpm\.)", "", name).upper()


@fn("IF", lazy=True)
def f_if(ev, args: list[Node]):
    if not args or len(args) > 3:
        return err(VALUE)
    c = ev.ev(args[0])
    if isinstance(c, Rng):
        c = ev.materialize(c) if ev.array else ev.intersect(c)
    if isinstance(c, Arr):
        t = ev.ev(args[1]) if len(args) > 1 else True
        f = ev.ev(args[2]) if len(args) > 2 else False
        t = ev.grid(t) if isinstance(t, (Rng, Arr)) else t
        f = ev.grid(f) if isinstance(f, (Rng, Arr)) else f

        def pick(src, i, j):
            if isinstance(src, Arr):
                return src.data[min(i, src.rows - 1)][min(j, src.cols - 1)]
            return src

        out = []
        for i, row in enumerate(c.data):
            r = []
            for j, x in enumerate(row):
                if isinstance(x, ExcelError):
                    r.append(x)
                    continue
                try:
                    r.append(pick(t, i, j) if to_bool(x) else pick(f, i, j))
                except ExcelError as e:
                    r.append(e)
            out.append(r)
        return Arr(out)
    if isinstance(c, ExcelError):
        return c
    try:
        cond = to_bool(c)
    except ExcelError as e:
        return e
    if cond:
        if len(args) < 2:
            return True
        if isinstance(args[1], Missing):
            return 0.0
        return ev.ev(args[1])
    if len(args) < 3:
        return False
    if isinstance(args[2], Missing):
        return 0.0
    return ev.ev(args[2])


@fn("IFS", lazy=True)
def f_ifs(ev, args):
    for i in range(0, len(args) - 1, 2):
        c = ev.scalar(ev.ev(args[i]))
        if isinstance(c, ExcelError):
            return c
        if to_bool(c):
            return ev.ev(args[i + 1])
    return err(NA)


@fn("IFERROR", lazy=True)
def f_iferror(ev, args):
    v = ev.ev(args[0])
    if isinstance(v, Rng) and not ev.array:
        v = ev.intersect(v)
    if isinstance(v, ExcelError):
        return ev.ev(args[1]) if not isinstance(args[1], Missing) else 0.0
    if isinstance(v, Arr):
        alt = None
        out = []
        for row in v.data:
            r = []
            for x in row:
                if isinstance(x, ExcelError):
                    if alt is None:
                        alt = ev.scalar(ev.ev(args[1]))
                    r.append(alt)
                else:
                    r.append(x)
            out.append(r)
        return Arr(out)
    return v


@fn("IFNA", lazy=True)
def f_ifna(ev, args):
    v = ev.scalar(ev.ev(args[0]))
    if isinstance(v, ExcelError) and v.code == NA:
        return ev.ev(args[1])
    return v


@fn("CHOOSE", lazy=True)
def f_choose(ev, args):
    i = ev.scalar(ev.ev(args[0]))
    if isinstance(i, ExcelError):
        return i
    k = int(to_number(i))
    if k < 1 or k >= len(args):
        return err(VALUE)
    return ev.ev(args[k])


def _bools(ev, args):
    out = []
    for a in args:
        if isinstance(a, (Rng, Arr)):
            for x in ev.values(a):
                if isinstance(x, ExcelError):
                    raise x
                if isinstance(x, bool) or is_number(x):
                    out.append(bool(x))
        else:
            out.append(to_bool(a))
    if not out:
        raise err(VALUE)
    return out


@fn("AND")
def f_and(ev, *args):
    return all(_bools(ev, args))


@fn("OR")
def f_or(ev, *args):
    return any(_bools(ev, args))


@fn("XOR")
def f_xor(ev, *args):
    return sum(_bools(ev, args)) % 2 == 1


@fn("NOT")
def f_not(ev, x):
    return not to_bool(S(ev, x))


@fn("TRUE")
def f_true(ev):
    return True


@fn("FALSE")
def f_false(ev):
    return False


# ================================================================ 情報
@fn("ISBLANK")
def f_isblank(ev, x):
    return S(ev, x) is None


@fn("ISNUMBER")
def f_isnumber(ev, x):
    return is_number(S(ev, x))


@fn("ISTEXT")
def f_istext(ev, x):
    return isinstance(S(ev, x), str)


@fn("ISNONTEXT")
def f_isnontext(ev, x):
    return not isinstance(S(ev, x), str)


@fn("ISLOGICAL")
def f_islogical(ev, x):
    return isinstance(S(ev, x), bool)


@fn("ISERROR")
def f_iserror(ev, x):
    return isinstance(S(ev, x), ExcelError)


@fn("ISERR")
def f_iserr(ev, x):
    v = S(ev, x)
    return isinstance(v, ExcelError) and v.code != NA


@fn("ISNA")
def f_isna(ev, x):
    v = S(ev, x)
    return isinstance(v, ExcelError) and v.code == NA


@fn("ISEVEN")
def f_iseven(ev, x):
    return int(math.trunc(num(ev, x))) % 2 == 0


@fn("ISODD")
def f_isodd(ev, x):
    return int(math.trunc(num(ev, x))) % 2 == 1


@fn("NA")
def f_na(ev):
    return err(NA)


@fn("N")
def f_n(ev, x):
    v = S(ev, x)
    if isinstance(v, ExcelError):
        return v
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    return float(v) if is_number(v) else 0.0


@fn("T")
def f_t(ev, x):
    v = S(ev, x)
    return v if isinstance(v, str) else ""


@fn("ROW")
def f_row(ev, ref=None):
    # 配列として計算するとき（配列数式・SUMPRODUCT の中など）、複数の行の範囲の ROW は、行の番号の縦の配列
    # （ROW(A2:A5) は {2;3;4;5}）。ふつうの数式では、最初の行の番号（Excel の仕様）
    if ref is None:
        return float(ev.row)
    if isinstance(ref, Rng):
        if ev.array and ref.rows > 1:
            return Arr([[float(r)] for r in range(ref.r1, ref.r2 + 1)])
        return float(ref.r1)
    raise err(VALUE)


@fn("COLUMN")
def f_column(ev, ref=None):
    # 配列として計算するとき、複数の列の範囲の COLUMN は、列の番号の横の配列（COLUMN(B1:D1) は {2,3,4}）
    if ref is None:
        return float(ev.col)
    if isinstance(ref, Rng):
        if ev.array and ref.cols > 1:
            return Arr([[float(c) for c in range(ref.c1, ref.c2 + 1)]])
        return float(ref.c1)
    raise err(VALUE)


@fn("ADDRESS")
def f_address(ev, row, col, abs_num=None, a1=None, sheet=None):
    from .parser import num_to_col

    r, c = int(num(ev, row)), int(num(ev, col))
    k = int(num(ev, abs_num)) if abs_num is not None else 1
    style_a1 = True if a1 is None else bool(to_bool(S(ev, a1)))
    if r < 1 or c < 1 or r > 1048576 or c > 16384 or k not in (1, 2, 3, 4):
        raise err(VALUE)
    if style_a1:
        ra = "$" if k in (1, 2) else ""
        ca = "$" if k in (1, 3) else ""
        text = f"{ca}{num_to_col(c)}{ra}{r}"
    else:
        rr = f"R{r}" if k in (1, 2) else f"R[{r}]"
        cc = f"C{c}" if k in (1, 3) else f"C[{c}]"
        text = rr + cc
    if sheet is not None:
        name = to_text(S(ev, sheet))
        if name:
            q = re.search(r"[^\w.]", name) or re.match(r"\d", name)
            text = (f"'{name.replace(chr(39), chr(39) * 2)}'" if q else name) + "!" + text
    return text


@fn("CELL")
def f_cell(ev, info, ref=None):
    """CELL：場所・内容・種類などを返す。ファイルの保存場所（filename）は環境で変わるので、計算のたびに変わる値として扱う。"""
    from .parser import num_to_col

    kind = to_text(S(ev, info)).lower()
    if ref is None:
        sheet, r, c = ev.sheet, ev.row, ev.col
    elif isinstance(ref, Rng):
        sheet, r, c = ref.sheet, ref.r1, ref.c1
    else:
        raise err(VALUE)
    if kind == "address":
        return f"${num_to_col(c)}${r}" if sheet == ev.sheet else f"{sheet}!${num_to_col(c)}${r}"
    if kind == "row":
        return float(r)
    if kind == "col":
        return float(c)
    if kind == "contents":
        v = ev.ctx.cell_value(sheet, r, c)
        return 0.0 if v is None else v
    if kind == "type":
        v = ev.ctx.cell_value(sheet, r, c)
        return "b" if v is None else ("l" if isinstance(v, str) else "v")
    if kind == "filename":
        ev.volatile = True  # 保存した場所によって変わる（Excelも開いた場所で計算し直す）
        wb = getattr(ev.ctx, "wb", None)
        name = (wb.filename or "Book1.xlsx").rsplit("/", 1)[-1] if wb is not None else "Book1.xlsx"
        return f"[{name}]{sheet}"
    if kind == "prefix":
        v = ev.ctx.cell_value(sheet, r, c)
        return "" if not isinstance(v, str) else "'"
    raise Unsupported(f"CELL の種類 {kind}")


@fn("ROWS")
def f_rows(ev, ref):
    return float(ev.grid(ref).rows if not isinstance(ref, Rng) else ref.rows)


@fn("COLUMNS")
def f_columns(ev, ref):
    return float(ev.grid(ref).cols if not isinstance(ref, Rng) else ref.cols)


# ================================================================ 検索
def _lookup_vector(vals: list, key, mode: int) -> int:
    """MATCH の位置（0始まり）。見つからなければ -1。"""
    if mode == 0:
        if isinstance(key, str) and any(ch in key for ch in "*?~"):
            p = _wild(key)
            for i, x in enumerate(vals):
                if isinstance(x, str) and p.match(x):
                    return i
            return -1
        for i, x in enumerate(vals):
            if x is None or isinstance(x, ExcelError):
                continue
            if _same_kind(x, key) and compare(x, key) == 0:
                return i
        return -1
    # 近似一致は二分探索（Excelと同じく、並んでいる前提）
    lo, hi = 0, len(vals) - 1
    best = -1
    while lo <= hi:
        mid = (lo + hi) // 2
        x = vals[mid]
        # 型の違う値・空白はとばす
        j = mid
        while j <= hi and (x is None or not _same_kind(x, key)):
            j += 1
            x = vals[j] if j <= hi else None
        if j > hi:
            hi = mid - 1
            continue
        c = compare(x, key)
        if mode == 1:
            if c <= 0:
                best = j
                lo = j + 1
            else:
                hi = mid - 1
        else:
            if c >= 0:
                best = j
                lo = j + 1
            else:
                hi = mid - 1
    return best


def _same_kind(a, b) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool)
    if isinstance(a, str) or isinstance(b, str):
        return isinstance(a, str) and isinstance(b, str)
    return is_number(a) and is_number(b)


@fn("MATCH")
def f_match(ev, key, arr, mode=None):
    k = S(ev, key)
    if isinstance(k, ExcelError):
        return k
    m = int(num(ev, mode)) if mode is not None else 1
    g = ev.grid(arr)
    if g.rows != 1 and g.cols != 1:
        raise err(NA)
    vals = _flat(g)
    i = _lookup_vector(vals, k, 0 if m == 0 else (1 if m > 0 else -1))
    if i < 0:
        raise err(NA)
    return float(i + 1)


def _approx_flag(ev, v) -> bool:
    if v is None:
        return True
    return to_bool(S(ev, v))


@fn("VLOOKUP")
def f_vlookup(ev, key, table, col, approx=None):
    k = S(ev, key)
    if isinstance(k, ExcelError):
        return k
    g = ev.grid(table)
    c = int(num(ev, col))
    if c < 1:
        raise err(VALUE)
    if c > g.cols:
        raise err(REF)
    keys = [row[0] for row in g.data]
    i = _lookup_vector(keys, k, 1 if _approx_flag(ev, approx) else 0)
    if i < 0:
        raise err(NA)
    v = g.data[i][c - 1]
    return 0.0 if v is None else v


@fn("HLOOKUP")
def f_hlookup(ev, key, table, row, approx=None):
    k = S(ev, key)
    if isinstance(k, ExcelError):
        return k
    g = ev.grid(table)
    r = int(num(ev, row))
    if r < 1:
        raise err(VALUE)
    if r > g.rows:
        raise err(REF)
    i = _lookup_vector(g.data[0], k, 1 if _approx_flag(ev, approx) else 0)
    if i < 0:
        raise err(NA)
    v = g.data[r - 1][i]
    return 0.0 if v is None else v


@fn("LOOKUP")
def f_lookup(ev, key, vec, result=None):
    k = S(ev, key)
    g = ev.grid(vec)
    if result is None:
        if g.rows >= g.cols:
            keys, res = [r[0] for r in g.data], [r[-1] for r in g.data]
        else:
            keys, res = g.data[0], g.data[-1]
    else:
        keys = _flat(g)
        res = _flat(ev.grid(result))
    i = _lookup_vector(keys, k, 1)
    if i < 0:
        raise err(NA)
    v = res[i] if i < len(res) else err(NA)
    return 0.0 if v is None else v


@fn("INDEX")
def f_index(ev, ref, r=None, c=None, area=None):
    if area is not None:
        raise Unsupported("INDEXの領域番号")
    rr = int(num(ev, r)) if r is not None else 0
    cc = int(num(ev, c)) if c is not None else 0
    if isinstance(ref, Rng):
        if ref.rows == 1 and c is None and r is not None:
            rr, cc = 1, rr
        if rr < 0 or cc < 0 or rr > ref.rows or cc > ref.cols:
            raise err(REF)
        if rr == 0 and cc == 0:
            return ref
        if rr == 0:
            return Rng(ref.sheet, ref.r1, ref.c1 + cc - 1, ref.r2, ref.c1 + cc - 1)
        if cc == 0:
            if ref.cols == 1:
                return Rng(ref.sheet, ref.r1 + rr - 1, ref.c1, ref.r1 + rr - 1, ref.c1)
            return Rng(ref.sheet, ref.r1 + rr - 1, ref.c1, ref.r1 + rr - 1, ref.c2)
        return Rng(ref.sheet, ref.r1 + rr - 1, ref.c1 + cc - 1, ref.r1 + rr - 1, ref.c1 + cc - 1)
    g = ev.grid(ref)
    if g.rows == 1 and c is None and r is not None:
        rr, cc = 1, rr
    if rr < 0 or cc < 0 or rr > g.rows or cc > g.cols:
        raise err(REF)
    if rr == 0 and cc == 0:
        return g  # 配列ぜんぶ
    if rr == 0:
        return Arr([[row[cc - 1]] for row in g.data]) if g.cols > 1 or c is not None else Arr([list(row) for row in g.data])
    if cc == 0:
        return Arr([list(g.data[rr - 1])]) if g.cols > 1 else g.data[rr - 1][0]
    return g.data[rr - 1][cc - 1]


@fn("OFFSET")
def f_offset(ev, ref, rows, cols, h=None, w=None):
    if not isinstance(ref, Rng):
        raise err(VALUE)
    dr, dc = int(num(ev, rows)), int(num(ev, cols))
    hh = int(num(ev, h)) if h is not None else ref.rows
    ww = int(num(ev, w)) if w is not None else ref.cols
    if hh < 1 or ww < 1:
        raise err(REF)
    r1, c1 = ref.r1 + dr, ref.c1 + dc
    if r1 < 1 or c1 < 1:
        raise err(REF)
    return Rng(ref.sheet, r1, c1, r1 + hh - 1, c1 + ww - 1)


@fn("INDIRECT")
def f_indirect(ev, text, a1=None):
    from .parser import RefNode, parse

    if a1 is not None and not to_bool(S(ev, a1)):
        raise Unsupported("INDIRECTのR1C1形式")
    t = to_text(S(ev, text))
    try:
        node = parse(t)
    except Exception:
        raise err(REF)
    if not isinstance(node, RefNode):
        res = ev.ctx.resolve_name(t, None, ev.sheet)
        if res is None:
            raise err(REF)
        return ev.ev(res[0])
    sheet = node.sheet or ev.sheet
    if not ev.ctx.has_sheet(sheet):
        raise err(REF)
    return Rng(sheet, node.r1, node.c1, node.r2, node.c2)


# ================================================================ 文字列
@fn("LEN")
def f_len(ev, x):
    return float(len(to_text(S(ev, x))))


@fn("LENB")
def f_lenb(ev, x):
    s = to_text(S(ev, x))
    return float(sum(1 if ord(ch) < 0x80 or 0xFF61 <= ord(ch) <= 0xFF9F else 2 for ch in s))


@fn("LEFT")
def f_left(ev, x, n=None):
    k = int(num(ev, n)) if n is not None else 1
    if k < 0:
        raise err(VALUE)
    return to_text(S(ev, x))[:k]


@fn("RIGHT")
def f_right(ev, x, n=None):
    k = int(num(ev, n)) if n is not None else 1
    if k < 0:
        raise err(VALUE)
    s = to_text(S(ev, x))
    return s[len(s) - k:] if k else ""


@fn("MID")
def f_mid(ev, x, start, n):
    s = to_text(S(ev, x))
    a, k = int(num(ev, start)), int(num(ev, n))
    if a < 1 or k < 0:
        raise err(VALUE)
    return s[a - 1:a - 1 + k]


@fn("CONCATENATE CONCAT")
def f_concat(ev, *args):
    out = []
    for a in args:
        if isinstance(a, (Rng, Arr)):
            for x in ev.values(a):
                out.append(to_text(x))
        else:
            out.append(to_text(a))
    return "".join(out)


@fn("TEXTJOIN")
def f_textjoin(ev, delim, ignore_empty, *args):
    d = to_text(S(ev, delim))
    ig = to_bool(S(ev, ignore_empty))
    parts = []
    for a in args:
        for x in (ev.values(a) if isinstance(a, (Rng, Arr)) else [a]):
            t = to_text(x)
            if ig and t == "":
                continue
            parts.append(t)
    return d.join(parts)


@fn("TRIM")
def f_trim(ev, x):
    return re.sub(r" +", " ", to_text(S(ev, x)).strip(" "))


@fn("UPPER")
def f_upper(ev, x):
    return to_text(S(ev, x)).upper()


@fn("LOWER")
def f_lower(ev, x):
    return to_text(S(ev, x)).lower()


@fn("PROPER")
def f_proper(ev, x):
    return to_text(S(ev, x)).title()


@fn("REPT")
def f_rept(ev, x, n):
    k = int(num(ev, n))
    if k < 0:
        raise err(VALUE)
    return to_text(S(ev, x)) * k


@fn("EXACT")
def f_exact(ev, a, b):
    return to_text(S(ev, a)) == to_text(S(ev, b))


@fn("SUBSTITUTE")
def f_substitute(ev, text, old, new, inst=None):
    s, o, n = to_text(S(ev, text)), to_text(S(ev, old)), to_text(S(ev, new))
    if o == "":
        return s
    if inst is None:
        return s.replace(o, n)
    k = int(num(ev, inst))
    if k < 1:
        raise err(VALUE)
    idx = -1
    for _ in range(k):
        idx = s.find(o, idx + 1)
        if idx < 0:
            return s
    return s[:idx] + n + s[idx + len(o):]


@fn("REPLACE")
def f_replace(ev, text, start, n, new):
    s = to_text(S(ev, text))
    a, k = int(num(ev, start)), int(num(ev, n))
    if a < 1 or k < 0:
        raise err(VALUE)
    return s[:a - 1] + to_text(S(ev, new)) + s[a - 1 + k:]


@fn("FIND")
def f_find(ev, needle, hay, start=None):
    n, h = to_text(S(ev, needle)), to_text(S(ev, hay))
    a = int(num(ev, start)) if start is not None else 1
    if a < 1 or a > len(h) + 1:
        raise err(VALUE)
    i = h.find(n, a - 1)
    if i < 0:
        raise err(VALUE)
    return float(i + 1)


@fn("SEARCH")
def f_search(ev, needle, hay, start=None):
    n, h = to_text(S(ev, needle)), to_text(S(ev, hay))
    a = int(num(ev, start)) if start is not None else 1
    if a < 1 or a > len(h) + 1:
        raise err(VALUE)
    p = _wild(n)
    pat = re.compile(p.pattern[1:-1], re.I | re.S)
    m = pat.search(h, a - 1)
    if not m:
        raise err(VALUE)
    return float(m.start() + 1)


@fn("VALUE")
def f_value(ev, x):
    v = S(ev, x)
    if v is None:
        return 0.0  # 空のセルは 0（Excelと同じ）
    if is_number(v):
        return float(v)
    if isinstance(v, ExcelError):
        return v
    return to_number(to_text(v))


@fn("TEXT")
def f_text(ev, x, fmt):
    from .numfmt import format_value

    v = S(ev, x)
    f = to_text(S(ev, fmt))
    if isinstance(v, ExcelError):
        return v
    if isinstance(v, str):
        n = parse_number(v)
        if n is not None:
            v = n
    return format_value(v, f)


@fn("CHAR")
def f_char(ev, x):
    k = int(num(ev, x))
    if k < 1 or k > 255:
        raise err(VALUE)
    return bytes([k]).decode("cp932", errors="replace") if k < 128 else bytes([k]).decode("cp1252", errors="replace")


@fn("CODE")
def f_code(ev, x):
    s = to_text(S(ev, x))
    if not s:
        raise err(VALUE)
    return float(ord(s[0])) if ord(s[0]) < 128 else float(int.from_bytes(s[0].encode("cp932", errors="replace"), "big"))


_ZEN = str.maketrans({chr(i): chr(i + 0xFEE0) for i in range(0x21, 0x7F)} | {" ": "　"})
_HAN = str.maketrans({chr(i + 0xFEE0): chr(i) for i in range(0x21, 0x7F)} | {"　": " "})


@fn("JIS DBCS")  # 同じ関数（日本語版の名前が JIS、英語版の名前が DBCS）
def f_jis(ev, x):
    import unicodedata

    s = to_text(S(ev, x))
    # 半角カナは NFKC で全角に。英数字は全角へ
    s = "".join(unicodedata.normalize("NFKC", ch) if 0xFF61 <= ord(ch) <= 0xFF9F else ch for ch in s)
    return s.translate(_ZEN)


def _kana_half_table() -> dict[str, str]:
    """全角カタカナ・記号 → 半角カナ。半角カナ（U+FF61〜U+FF9F）を NFKC で全角にした対応を逆に引く。
    濁点・半濁点のつく字（ガ・パ・ヴ）は、半角の字と濁点の2字になる。"""
    import unicodedata

    table: dict[str, str] = {}
    halves = [chr(c) for c in range(0xFF61, 0xFFA0)]
    for h in halves:
        f = unicodedata.normalize("NFKC", h)
        if len(f) == 1 and f not in table:
            table[f] = h
    for h in halves:
        for mark in ("\uff9e", "\uff9f"):
            f = unicodedata.normalize("NFKC", h + mark)
            if len(f) == 1 and f not in table:
                table[f] = h + mark
    return table


_KANA_HALF = _kana_half_table()


def _asc_text(s: str) -> str:
    """ASC：全角の英数字・記号・カタカナを半角にする（ひらがな・漢字はそのまま）。"""
    return "".join(_KANA_HALF.get(ch, ch) for ch in s.translate(_HAN))


@fn("ASC")
def f_asc(ev, x):
    return _asc_text(to_text(S(ev, x)))


# ================================================================ 日付
def _today():
    return dt.date.today()


@fn("TODAY")
def f_today(ev):
    ev.volatile = True
    return float(date_to_serial(_today()))


@fn("NOW")
def f_now(ev):
    ev.volatile = True
    n = dt.datetime.now()
    return float(date_to_serial(n.date())) + (n.hour * 3600 + n.minute * 60 + n.second) / 86400


@fn("DATE")
def f_date(ev, y, m, d):
    yy, mm, dd = int(num(ev, y)), int(num(ev, m)), int(num(ev, d))
    if yy < 1900:
        yy += 1900
    if yy > 9999:
        raise err(NUM)
    # 月・日のあふれは繰り上げる
    yy += (mm - 1) // 12
    mm = (mm - 1) % 12 + 1
    base = dt.date(yy, mm, 1)
    n = date_to_serial(base) + dd - 1
    if n < 0:
        raise err(NUM)
    return float(n)


@fn("YEAR")
def f_year(ev, x):
    return float(serial_parts(_serial(ev, x))[0])


@fn("MONTH")
def f_month(ev, x):
    return float(serial_parts(_serial(ev, x))[1])


@fn("DAY")
def f_day(ev, x):
    return float(serial_parts(_serial(ev, x))[2])


def _serial(ev, x) -> float:
    v = num(ev, x)
    if v < 0:
        raise err(NUM)
    return v


@fn("HOUR")
def f_hour(ev, x):
    v = _serial(ev, x)
    return float(int(round((v - math.floor(v)) * 86400)) // 3600)


@fn("MINUTE")
def f_minute(ev, x):
    v = _serial(ev, x)
    return float(int(round((v - math.floor(v)) * 86400)) % 3600 // 60)


@fn("SECOND")
def f_second(ev, x):
    v = _serial(ev, x)
    return float(int(round((v - math.floor(v)) * 86400)) % 60)


@fn("TIME")
def f_time(ev, h, m, s):
    t = num(ev, h) * 3600 + num(ev, m) * 60 + num(ev, s)
    if t < 0:
        raise err(NUM)
    return (t % 86400) / 86400


@fn("WEEKDAY")
def f_weekday(ev, x, kind=None):
    v = int(_serial(ev, x))
    k = int(num(ev, kind)) if kind is not None else 1
    # 1900 年の数え方：シリアル1 = 1900/1/1 は（Excel上）日曜。1900/2/29（シリアル60）が存在する扱いなので、曜日は通しで7日周期になる
    # 1904 年の数え方：実際の日付の曜日
    sun0 = weekday_sun0(v)
    if k in (1, 17):
        return float(sun0 + 1)
    if k in (2, 11):
        return float((sun0 + 6) % 7 + 1)
    if k == 3:
        return float((sun0 + 6) % 7)
    if 12 <= k <= 16:
        # 12: 火曜=1 … 16: 土曜=1（週の始まりの曜日をずらす）
        start = k - 10  # 火=2 水=3 木=4 金=5 土=6（日=0 として）
        return float((sun0 - start) % 7 + 1)
    raise err(NUM)


@fn("EDATE")
def f_edate(ev, start, months):
    d = serial_to_date(int(_serial(ev, start)))
    m = int(num(ev, months))
    y, mo = divmod(d.month - 1 + m, 12)
    y += d.year
    mo += 1
    import calendar

    day = min(d.day, calendar.monthrange(y, mo)[1])
    return float(date_to_serial(dt.date(y, mo, day)))


@fn("EOMONTH")
def f_eomonth(ev, start, months):
    d = serial_to_date(int(_serial(ev, start)))
    m = int(num(ev, months))
    y, mo = divmod(d.month - 1 + m, 12)
    y += d.year
    mo += 1
    import calendar

    return float(date_to_serial(dt.date(y, mo, calendar.monthrange(y, mo)[1])))


@fn("DATEDIF")
def f_datedif(ev, a, b, unit):
    d1 = serial_to_date(int(_serial(ev, a)))
    d2 = serial_to_date(int(_serial(ev, b)))
    u = to_text(S(ev, unit)).upper()
    if d1 > d2:
        raise err(NUM)
    if u == "D":
        return float((d2 - d1).days)
    months = (d2.year - d1.year) * 12 + d2.month - d1.month - (1 if d2.day < d1.day else 0)
    if u == "M":
        return float(months)
    if u == "Y":
        return float(months // 12)
    if u == "YM":
        return float(months % 12)
    if u == "MD":
        if d2.day >= d1.day:
            return float(d2.day - d1.day)
        import calendar

        pm = d2.month - 1 or 12
        py = d2.year if d2.month > 1 else d2.year - 1
        return float(calendar.monthrange(py, pm)[1] - d1.day + d2.day)
    raise Unsupported(f"DATEDIFの単位 {u}")


@fn("DAYS")
def f_days(ev, end, start):
    return float(int(_serial(ev, end)) - int(_serial(ev, start)))


@fn("DATEVALUE")
def f_datevalue(ev, x):
    v = S(ev, x)
    if not isinstance(v, str):
        raise err(VALUE)
    n = to_number(v)
    return float(int(n))


def _holidays(ev, h):
    if h is None:
        return set()
    return {int(x) for x in ev.values(h) if is_number(x)}


@fn("NETWORKDAYS")
def f_networkdays(ev, a, b, holidays=None):
    s, e = int(_serial(ev, a)), int(_serial(ev, b))
    hol = _holidays(ev, holidays)
    sign = 1
    if s > e:
        s, e, sign = e, s, -1
    n = 0
    for d in range(s, e + 1):
        wd = serial_to_date(d).weekday()
        if wd < 5 and d not in hol:
            n += 1
    return float(n * sign)


@fn("WORKDAY")
def f_workday(ev, a, days, holidays=None):
    d = int(_serial(ev, a))
    k = int(num(ev, days))
    hol = _holidays(ev, holidays)
    step = 1 if k >= 0 else -1
    while k != 0:
        d += step
        if serial_to_date(d).weekday() < 5 and d not in hol:
            k -= step
    return float(d)


# ================================================================ 未対応（明示）
@fn("HYPERLINK")
def f_hyperlink(ev, link, friendly=None):
    v = S(ev, friendly) if friendly is not None else S(ev, link)
    return 0.0 if v is None else v


@fn("PHONETIC")
def f_phonetic(ev, ref):
    """PHONETIC：セルの文字に付いているふりがな。ふりがながなければ、文字そのもの。"""
    if not isinstance(ref, Rng):
        raise err(VALUE)
    wb = getattr(ev.ctx, "wb", None)
    out = []
    for r in range(ref.r1, ref.r2 + 1):
        for c in range(ref.c1, ref.c2 + 1):
            v = ev.ctx.cell_value(ref.sheet, r, c)
            if not isinstance(v, str):
                continue
            i = wb.cell_sst.get((ref.sheet, r, c)) if wb is not None else None
            out.append(wb.phonetic.get(i, v) if i is not None else v)
    return "".join(out)


# 環境や外部によって決まるもの・まだ再現していないもの（明示的に「未対応」にする）
for _name in ("GETPIVOTDATA", "INFO", "WEBSERVICE", "FILTERXML", "LAMBDA", "RAND", "RANDBETWEEN"):
    EAGER.pop(_name, None)

# 動的配列の関数・XLOOKUP・XMATCH・通貨の書式（arrays.py で登録する）
from . import arrays as _arrays  # noqa: E402,F401
