"""動的配列の関数（SEQUENCE・FILTER・UNIQUE・SORT など）と、XLOOKUP・XMATCH、通貨の書式の関数。

Microsoft の関数の説明（support.microsoft.com）の規則どおりに計算する。配列を返す関数の結果は、
配列数式（動的配列はファイルの中で配列数式として書かれる）のセルに、左上から順に入る。
"""

from __future__ import annotations

from collections import Counter
from functools import cmp_to_key

from .functions import EAGER, LAZY, S, _same_kind, _wild, fn, num
from .values import NA, VALUE, Arr, ExcelError, Rng, compare, err, is_number, round_half_up, to_15, to_bool, to_number

CALC = "#CALC!"  # 空の配列（FILTER で何も残らないなど）


def _grid(ev, v) -> Arr:
    g = ev.grid(v)
    if g.rows == 0 or g.cols == 0:
        raise err(CALC)
    return g


def _int(ev, v, default: int | None = None) -> int:
    if v is None:
        if default is None:
            raise err(VALUE)
        return default
    x = S(ev, v)
    if isinstance(x, ExcelError):
        raise x
    return int(to_number(x))  # 小数は切り捨て（0 の方へ）


def _flag(ev, v, default: bool = False) -> bool:
    if v is None:
        return default
    x = S(ev, v)
    if isinstance(x, ExcelError):
        raise x
    return to_bool(x)


def _truthy(x) -> bool:
    if isinstance(x, ExcelError):
        raise x
    if x is None:
        return False
    if isinstance(x, bool):
        return x
    if is_number(x):
        return x != 0
    raise err(VALUE)  # 文字列は条件にならない


def _transpose(rows: list[list]) -> list[list]:
    return [list(c) for c in zip(*rows)] if rows else []


# ================================================================ 配列を作る・選ぶ
@fn("SEQUENCE")
def f_sequence(ev, rows, cols=None, start=None, step=None):
    r = _int(ev, rows)
    c = _int(ev, cols, 1)
    s = num(ev, start) if start is not None else 1.0
    st = num(ev, step) if step is not None else 1.0
    if r < 0 or c < 0:
        raise err(VALUE)
    if r == 0 or c == 0:
        raise err(CALC)
    if r * c > 1_048_576 * 16_384 or r > 1_048_576 or c > 16_384:
        raise err("#NUM!")
    return Arr([[to_15(s + (i * c + j) * st) for j in range(c)] for i in range(r)])


@fn("FILTER")
def f_filter(ev, array, include, if_empty=None):
    a = _grid(ev, array)
    inc = ev.grid(include)
    if inc.cols == 1 and inc.rows == a.rows:
        keep = [a.data[i] for i in range(a.rows) if _truthy(inc.data[i][0])]
        if keep:
            return Arr([list(r) for r in keep])
    elif inc.rows == 1 and inc.cols == a.cols:
        cols = [j for j in range(a.cols) if _truthy(inc.data[0][j])]
        if cols:
            return Arr([[row[j] for j in cols] for row in a.data])
    else:
        raise err(VALUE)
    if if_empty is None:
        raise err(CALC)
    return if_empty if isinstance(if_empty, (Arr, Rng)) else if_empty


def _ukey(x):
    """UNIQUE の比べ方：大文字・小文字は区別しない。数と文字列の「1」は別のもの。"""
    if isinstance(x, ExcelError):
        return ("x", x.code)
    if x is None:
        return ("e",)
    if isinstance(x, bool):
        return ("b", x)
    if is_number(x):
        return ("n", to_15(float(x)))
    return ("s", str(x).lower())


@fn("UNIQUE")
def f_unique(ev, array, by_col=None, exactly_once=None):
    a = _grid(ev, array)
    bc = _flag(ev, by_col)
    once = _flag(ev, exactly_once)
    items = _transpose(a.data) if bc else [list(r) for r in a.data]
    keys = [tuple(_ukey(x) for x in it) for it in items]
    count = Counter(keys)
    seen: set = set()
    out = []
    for it, k in zip(items, keys):
        if (once and count[k] > 1) or k in seen:
            continue
        seen.add(k)
        out.append(it)
    if not out:
        raise err(CALC)
    return Arr(_transpose(out) if bc else out)


def _sort_rank(x) -> int:
    # 並べ替えの順：数値 < 文字列 < 論理値 < エラー < 空白（空白はいつも最後）
    if x is None:
        return 4
    if isinstance(x, ExcelError):
        return 3
    if isinstance(x, bool):
        return 2
    if isinstance(x, str):
        return 1
    return 0


def _sort_cmp(a, b, order: int) -> int:
    ra, rb = _sort_rank(a), _sort_rank(b)
    if ra == 4 or rb == 4:
        return (ra == 4) - (rb == 4)  # 空白は昇順でも降順でも最後
    if ra != rb:
        c = -1 if ra < rb else 1
    elif ra == 3:
        c = 0
    else:
        c = compare(a, b)
    return c * order


def _sort_rows(rows: list[list], keys: list[tuple[list, int]]) -> list[list]:
    """keys：(各行のキーの値の並び, 1 か -1)。同じ値の行は、元の順のまま（安定）。"""
    idx = list(range(len(rows)))

    def cmp(i, j):
        for vals, order in keys:
            c = _sort_cmp(vals[i], vals[j], order)
            if c:
                return c
        return 0

    idx.sort(key=cmp_to_key(cmp))
    return [rows[i] for i in idx]


@fn("SORT")
def f_sort(ev, array, sort_index=None, sort_order=None, by_col=None):
    a = _grid(ev, array)
    bc = _flag(ev, by_col)
    rows = _transpose(a.data) if bc else [list(r) for r in a.data]
    width = len(rows[0]) if rows else 0
    idxs = [int(to_number(x)) for x in ev.values(sort_index)] if sort_index is not None else [1]
    orders = [int(to_number(x)) for x in ev.values(sort_order)] if sort_order is not None else [1]
    if len(orders) == 1 and len(idxs) > 1:
        orders = orders * len(idxs)
    if len(orders) != len(idxs) or any(i < 1 or i > width for i in idxs) or any(o not in (1, -1) for o in orders):
        raise err(VALUE)
    out = _sort_rows(rows, [([r[i - 1] for r in rows], o) for i, o in zip(idxs, orders)])
    return Arr(_transpose(out) if bc else out)


@fn("SORTBY")
def f_sortby(ev, array, *args):
    a = _grid(ev, array)
    if not args:
        raise err(VALUE)
    pairs = []
    for k in range(0, len(args), 2):
        by = ev.grid(args[k])
        order = _int(ev, args[k + 1], 1) if k + 1 < len(args) else 1
        if order not in (1, -1):
            raise err(VALUE)
        pairs.append((by, order))
    by_col = all(by.rows == 1 and by.cols == a.cols and a.cols > 1 for by, _o in pairs) and not \
        all(by.cols == 1 and by.rows == a.rows for by, _o in pairs)
    rows = _transpose(a.data) if by_col else [list(r) for r in a.data]
    keys = []
    for by, order in pairs:
        vals = by.data[0] if by_col else [r[0] for r in by.data]
        if (by_col and by.rows != 1) or (not by_col and by.cols != 1) or len(vals) != len(rows):
            raise err(VALUE)
        keys.append((list(vals), order))
    out = _sort_rows(rows, keys)
    return Arr(_transpose(out) if by_col else out)


def _flatten(ev, array, ignore, scan_by_col) -> list:
    a = _grid(ev, array)
    ig = _int(ev, ignore, 0)
    if ig not in (0, 1, 2, 3):
        raise err(VALUE)
    seq = [a.data[i][j] for j in range(a.cols) for i in range(a.rows)] if _flag(ev, scan_by_col) \
        else [x for row in a.data for x in row]
    out = [x for x in seq if not ((ig in (1, 3) and x is None) or (ig in (2, 3) and isinstance(x, ExcelError)))]
    if not out:
        raise err(CALC)
    return out


@fn("TOCOL")
def f_tocol(ev, array, ignore=None, scan_by_col=None):
    return Arr([[x] for x in _flatten(ev, array, ignore, scan_by_col)])


@fn("TOROW")
def f_torow(ev, array, ignore=None, scan_by_col=None):
    return Arr([_flatten(ev, array, ignore, scan_by_col)])


@fn("TRANSPOSE")
def f_transpose(ev, array):
    return Arr(_transpose(_grid(ev, array).data))


@fn("VSTACK")
def f_vstack(ev, *arrays):
    gs = [_grid(ev, a) for a in arrays]
    width = max(g.cols for g in gs)
    out = []
    for g in gs:
        for row in g.data:
            out.append(list(row) + [err(NA)] * (width - len(row)))
    return Arr(out)


@fn("HSTACK")
def f_hstack(ev, *arrays):
    gs = [_grid(ev, a) for a in arrays]
    height = max(g.rows for g in gs)
    out = [[] for _ in range(height)]
    for g in gs:
        for i in range(height):
            out[i] += list(g.data[i]) if i < g.rows else [err(NA)] * g.cols
    return Arr(out)


def _pick_index(k: int, n: int) -> int:
    if k == 0 or abs(k) > n:
        raise err(VALUE)
    return k - 1 if k > 0 else n + k


@fn("CHOOSECOLS")
def f_choosecols(ev, array, *nums):
    a = _grid(ev, array)
    if not nums:
        raise err(VALUE)
    cols = [_pick_index(int(to_number(x)), a.cols) for n_ in nums for x in ev.values(n_)]
    return Arr([[row[j] for j in cols] for row in a.data])


@fn("CHOOSEROWS")
def f_chooserows(ev, array, *nums):
    a = _grid(ev, array)
    if not nums:
        raise err(VALUE)
    rows = [_pick_index(int(to_number(x)), a.rows) for n_ in nums for x in ev.values(n_)]
    return Arr([list(a.data[i]) for i in rows])


def _take_range(k: int | None, n: int, drop: bool) -> range:
    if k is None:
        return range(n)
    if drop:
        if k >= 0:
            return range(min(k, n), n)
        return range(0, max(0, n + k))
    if k == 0:
        raise err(CALC)
    if k > 0:
        return range(0, min(k, n))
    return range(max(0, n + k), n)


def _take_drop(ev, array, rows, cols, drop: bool):
    a = _grid(ev, array)
    r = None if rows is None else _int(ev, rows)
    c = None if cols is None else _int(ev, cols)
    if not drop and c == 0:
        raise err(CALC)
    rr = _take_range(r, a.rows, drop)
    cc = _take_range(c, a.cols, drop)
    if len(rr) == 0 or len(cc) == 0:
        raise err(CALC)
    return Arr([[a.data[i][j] for j in cc] for i in rr])


@fn("TAKE")
def f_take(ev, array, rows=None, cols=None):
    return _take_drop(ev, array, rows, cols, False)


@fn("DROP")
def f_drop(ev, array, rows=None, cols=None):
    return _take_drop(ev, array, rows, cols, True)


@fn("EXPAND")
def f_expand(ev, array, rows=None, cols=None, pad=None):
    a = _grid(ev, array)
    r = _int(ev, rows, a.rows) if rows is not None else a.rows
    c = _int(ev, cols, a.cols) if cols is not None else a.cols
    if r < a.rows or c < a.cols:
        raise err(VALUE)
    p = S(ev, pad) if pad is not None else err(NA)
    out = [list(a.data[i]) + [p] * (c - a.cols) if i < a.rows else [p] * c for i in range(r)]
    return Arr(out)


def _vector(ev, v) -> list:
    g = _grid(ev, v)
    if g.rows != 1 and g.cols != 1:
        raise err(VALUE)
    return [x for row in g.data for x in row]


@fn("WRAPROWS")
def f_wraprows(ev, vector, count, pad=None):
    xs = _vector(ev, vector)
    n = _int(ev, count)
    if n < 1:
        raise err("#NUM!")
    p = S(ev, pad) if pad is not None else err(NA)
    rows = [xs[i:i + n] for i in range(0, len(xs), n)]
    rows[-1] = rows[-1] + [p] * (n - len(rows[-1]))
    return Arr(rows)


@fn("WRAPCOLS")
def f_wrapcols(ev, vector, count, pad=None):
    return Arr(_transpose(f_wraprows(ev, vector, count, pad).data))


# ================================================================ XLOOKUP・XMATCH
def _xfind(vals: list, key, match_mode: int, search_mode: int) -> int:
    """XLOOKUP・XMATCH の位置（0始まり、見つからなければ -1）。
    一致モード 0：完全一致（ワイルドカードは使わない）、-1：完全一致か次に小さい値、1：完全一致か次に大きい値、
    2：ワイルドカード。検索モード 1：先頭から、-1：末尾から、2・-2：並んでいる前提の二分探索。"""
    if isinstance(key, ExcelError):
        raise key
    if key is None:
        key = 0.0 if match_mode != 2 else ""
    n = len(vals)
    if search_mode in (2, -2):
        return _xbinary(vals, key, match_mode, search_mode)
    order = range(n) if search_mode == 1 else range(n - 1, -1, -1)
    if match_mode == 2 and isinstance(key, str):
        p = _wild(key)
        for i in order:
            x = vals[i]
            if isinstance(x, str) and p.match(x):
                return i
        return -1
    best = -1
    for i in order:
        x = vals[i]
        if x is None or isinstance(x, ExcelError) or not _same_kind(x, key):
            continue
        c = compare(x, key)
        if c == 0:
            return i
        if match_mode == -1 and c < 0 and (best < 0 or compare(x, vals[best]) > 0):
            best = i
        elif match_mode == 1 and c > 0 and (best < 0 or compare(x, vals[best]) < 0):
            best = i
    return best if match_mode in (-1, 1) else -1


def _xbinary(vals: list, key, match_mode: int, search_mode: int) -> int:
    asc = search_mode == 2
    lo, hi = 0, len(vals) - 1
    lower = upper = -1  # 次に小さい値・次に大きい値の位置
    while lo <= hi:
        mid = (lo + hi) // 2
        x = vals[mid]
        if x is None or isinstance(x, ExcelError) or not _same_kind(x, key):
            c = 1 if asc else -1  # 型の違う値は、後ろにあるものとして扱う
        else:
            c = compare(x, key)
        if c == 0:
            return mid
        if (c < 0) == asc:
            if c < 0:
                lower = mid
            else:
                upper = mid
            lo = mid + 1
        else:
            if c < 0:
                lower = mid
            else:
                upper = mid
            hi = mid - 1
    if match_mode == -1:
        return lower
    if match_mode == 1:
        return upper
    return -1


def _modes(ev, match_mode, search_mode) -> tuple[int, int]:
    m = _int(ev, match_mode, 0) if match_mode is not None else 0
    s = _int(ev, search_mode, 1) if search_mode is not None else 1
    if m not in (0, -1, 1, 2) or s not in (1, -1, 2, -2):
        raise err(VALUE)
    return m, s


@fn("XMATCH")
def f_xmatch(ev, key, look, match_mode=None, search_mode=None):
    m, s = _modes(ev, match_mode, search_mode)
    g = _grid(ev, look)
    if g.rows != 1 and g.cols != 1:
        raise err(VALUE)
    i = _xfind([x for row in g.data for x in row], S(ev, key), m, s)
    if i < 0:
        raise err(NA)
    return float(i + 1)


@fn("XLOOKUP")
def f_xlookup(ev, key, look, ret, if_not_found=None, match_mode=None, search_mode=None):
    m, s = _modes(ev, match_mode, search_mode)
    lg = _grid(ev, look)
    if lg.rows != 1 and lg.cols != 1:
        raise err(VALUE)
    vertical = lg.cols == 1 and not (lg.rows == 1 and lg.cols > 1)
    vals = [r[0] for r in lg.data] if vertical else list(lg.data[0])
    rg = ev.grid(ret)
    if (vertical and rg.rows != len(vals)) or (not vertical and rg.cols != len(vals)):
        raise err(VALUE)
    i = _xfind(vals, S(ev, key), m, s)
    if i < 0:
        if if_not_found is not None:
            return if_not_found
        raise err(NA)
    # 返す値：範囲なら、見つかった行（または列）の範囲（XLOOKUP(…):XLOOKUP(…) のように参照としても使える）
    if isinstance(ret, Rng):
        if vertical:
            return Rng(ret.sheet, ret.r1 + i, ret.c1, ret.r1 + i, ret.c2)
        return Rng(ret.sheet, ret.r1, ret.c1 + i, ret.r2, ret.c1 + i)
    if vertical:
        row = list(rg.data[i])
        return row[0] if len(row) == 1 else Arr([row])
    col = [[r[i]] for r in rg.data]
    return col[0][0] if len(col) == 1 else Arr(col)


# ================================================================ 通貨・桁区切りの書式
def _fmt_fixed(x: float, d: int, commas: bool) -> str:
    v = round_half_up(x, d)
    digits = max(d, 0)
    s = f"{abs(v):,.{digits}f}" if commas else f"{abs(v):.{digits}f}"
    return ("-" if v < 0 else "") + s


@fn("FIXED")
def f_fixed(ev, x, decimals=None, no_commas=None):
    d = _int(ev, decimals, 2)
    if d > 127:
        raise err(VALUE)
    return _fmt_fixed(num(ev, x), d, not _flag(ev, no_commas))


@fn("DOLLAR")
def f_dollar(ev, x, decimals=None):
    """DOLLAR：書式 $#,##0.00_);($#,##0.00)（負の数はかっこ）。"""
    d = _int(ev, decimals, 2)
    if d > 127:
        raise err(VALUE)
    s = _fmt_fixed(num(ev, x), d, True)
    return f"(${s[1:]})" if s.startswith("-") else f"${s}"


@fn("YEN")
def f_yen(ev, x, decimals=None):
    """YEN：円の通貨の書式（¥#,##0、負の数は -¥）。"""
    d = _int(ev, decimals, 0)
    if d > 127:
        raise err(VALUE)
    s = _fmt_fixed(num(ev, x), d, True)
    return f"-¥{s[1:]}" if s.startswith("-") else f"¥{s}"


# ================================================================ 配列の要素ごとの計算
# 1つの値を受け取る関数に、配列数式の中で範囲・配列が渡されたら、要素ごとに計算して配列で返す（Excelの動的配列と同じ）
SCALAR_FUNCS = set("""
ABS ASC CEILING CEILING.MATH CEILING.PRECISE CHAR CODE DATE DATEVALUE DAY DAYS EDATE EOMONTH EXACT FIND FLOOR
FLOOR.MATH FLOOR.PRECISE HOUR INT ISO.CEILING ISEVEN ISODD ISBLANK ISERR ISERROR ISLOGICAL ISNA ISNONTEXT ISNUMBER
ISTEXT JIS LEFT LEN LENB LOWER MID MINUTE MOD MONTH MROUND N NOT POWER PROPER QUOTIENT REPLACE REPT RIGHT ROUND
ROUNDDOWN ROUNDUP SEARCH SECOND SIGN SQRT SUBSTITUTE T TEXT TIME TRIM TRUNC UPPER VALUE WEEKDAY YEAR DATEDIF
DOLLAR YEN FIXED
""".split())


def broadcast(ev, f, args: list):
    """要素ごとの計算。大きさの違う配列は、1行・1列のものを広げ、足りない所は #N/A。"""
    grids = {i: ev.grid(a) for i, a in enumerate(args) if isinstance(a, (Rng, Arr))}
    multi = {i: g for i, g in grids.items() if g.rows * g.cols > 1}
    if not multi:
        return None
    rows = max(g.rows for g in multi.values())
    cols = max(g.cols for g in multi.values())
    out = []
    for r in range(rows):
        line = []
        for c in range(cols):
            a = list(args)
            for i, g in grids.items():
                rr = 0 if g.rows == 1 else r
                cc = 0 if g.cols == 1 else c
                a[i] = g.data[rr][cc] if rr < g.rows and cc < g.cols else err(NA)
            try:
                line.append(f(ev, *a))
            except ExcelError as e:
                line.append(e)
            except ZeroDivisionError:
                line.append(err("#DIV/0!"))
        out.append(line)
    return Arr(out)


__all__ = ["CALC", "SCALAR_FUNCS", "broadcast", "EAGER", "LAZY"]
