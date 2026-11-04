"""動的配列の関数（SEQUENCE・FILTER・UNIQUE・SORT など）と、XLOOKUP・XMATCH、通貨の書式の関数。

Microsoft の関数の説明（support.microsoft.com）の規則どおりに計算する。配列を返す関数の結果は、
配列数式（動的配列はファイルの中で配列数式として書かれる）のセルに、左上から順に入る。
"""

from __future__ import annotations

from bisect import bisect_right
from collections import Counter
from functools import cmp_to_key

from .functions import EAGER, LAZY, S, _same_kind, _wild, fn, num
from .lazyarr import LazyArr, Poly, make, tick, want_lazy
from .values import NA, VALUE, Arr, ExcelError, Rng, array_room, compare, err, is_big, is_number, round_half_up, to_15, to_bool, to_number, to_text

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
    array_room(ev, r, c)
    poly = _sequence_poly(s, st, r, c)
    if is_big(r, c):
        return make(ev, r, c, lambda i, j: to_15(s + (i * c + j) * st), poly)
    out = Arr([[to_15(s + (i * c + j) * st) for j in range(c)] for i in range(r)])
    out.poly = poly
    return out


def _sequence_poly(s: float, st: float, r: int, c: int):
    """SEQUENCE の値の多項式。マスの値は 15 桁に丸めるので、十進で 15 桁に入るときだけ。"""
    from fractions import Fraction

    a, d = Fraction(f"{s:.15g}"), Fraction(f"{st:.15g}")
    if float(a) != s or float(d) != st or max(a.denominator, d.denominator) > 1 << 40:
        return None
    if (abs(a) + abs(d) * r * c) * max(a.denominator, d.denominator) >= 10 ** 15:
        return None
    return Poly.linear(a, d * c, d)


@fn("FILTER")
def f_filter(ev, array, include, if_empty=None):
    a = _grid(ev, array)
    inc = ev.grid(include)
    if want_lazy(a, inc):
        return _filter_big(ev, a, inc, if_empty)
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


def _pick_rows(ev, a: Arr, rows: list[int]) -> Arr:
    """行を選んだ配列（並べ替えも）：選んだ番号だけを持つ。"""
    w = a.cols
    if is_big(len(rows), w):
        return make(ev, len(rows), w, lambda i, j: a.at(rows[i], j))
    return Arr([[a.at(i, j) for j in range(w)] for i in rows])


def _pick_cols(ev, a: Arr, cols: list[int]) -> Arr:
    if is_big(a.rows, len(cols)):
        return make(ev, a.rows, len(cols), lambda i, j: a.at(i, cols[j]))
    return Arr([[a.at(i, j) for j in cols] for i in range(a.rows)])


def _filter_big(ev, a: Arr, inc: Arr, if_empty):
    t0: list = []
    if inc.cols == 1 and inc.rows == a.rows:
        keep = []
        for i in range(a.rows):
            if not i % 4096:
                tick("FILTER", i, a.rows, t0)
            if _truthy(inc.at(i, 0)):
                keep.append(i)
        if keep:
            return _pick_rows(ev, a, keep)
    elif inc.rows == 1 and inc.cols == a.cols:
        keep = [j for j in range(a.cols) if _truthy(inc.at(0, j))]
        if keep:
            return _pick_cols(ev, a, keep)
    else:
        raise err(VALUE)
    if if_empty is None:
        raise err(CALC)
    return if_empty


def _digest(values) -> bytes:
    import hashlib

    h = hashlib.blake2b(digest_size=16)
    for x in values:
        h.update(repr(_ukey(x)).encode("utf-8", "surrogatepass"))
        h.update(b"\0")
    return h.digest()


def _unique_big(ev, a: Arr, bc: bool, once: bool) -> Arr:
    """行（か列）の全部を持たず、行ごとの指紋で比べる（行の数まで）。"""
    n, w = (a.cols, a.rows) if bc else (a.rows, a.cols)
    at = (lambda i, j: a.at(j, i)) if bc else a.at
    count: dict[bytes, int] = {}
    first: list[tuple[int, bytes]] = []
    t0: list = []
    for i in range(n):
        if not i % 256:
            tick("UNIQUE", i * w, n * w, t0)
        k = _digest(at(i, j) for j in range(w))
        if k not in count:
            count[k] = 0
            first.append((i, k))
        count[k] += 1
    keep = [i for i, k in first if not once or count[k] == 1]
    if not keep:
        raise err(CALC)
    return _pick_cols(ev, a, keep) if bc else _pick_rows(ev, a, keep)


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
    if want_lazy(a):
        return _unique_big(ev, a, bc, once)
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
    if want_lazy(a):
        return _sort_big(ev, a, bc, sort_index, sort_order)
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


def _sort_big(ev, a: Arr, bc: bool, sort_index, sort_order) -> Arr:
    """並べ替えの鍵の列だけを持ち、行は番号で並べる。"""
    n, width = (a.cols, a.rows) if bc else (a.rows, a.cols)
    idxs = [int(to_number(x)) for x in ev.values(sort_index)] if sort_index is not None else [1]
    orders = [int(to_number(x)) for x in ev.values(sort_order)] if sort_order is not None else [1]
    if len(orders) == 1 and len(idxs) > 1:
        orders = orders * len(idxs)
    if len(orders) != len(idxs) or any(i < 1 or i > width for i in idxs) or any(o not in (1, -1) for o in orders):
        raise err(VALUE)
    keys = []
    t0: list = []
    for k, (i, o) in enumerate(zip(idxs, orders)):
        tick("SORT", k, len(idxs), t0)
        keys.append(([a.at(x, i - 1) if not bc else a.at(i - 1, x) for x in range(n)], o))
    order = _sort_rows(list(range(n)), keys)
    return _pick_cols(ev, a, order) if bc else _pick_rows(ev, a, order)


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
    if want_lazy(a):
        n = a.cols if by_col else a.rows
        keys = []
        for by, order in pairs:
            if (by_col and by.rows != 1) or (not by_col and by.cols != 1) or by.rows * by.cols != n:
                raise err(VALUE)
            keys.append(([by.at(0, k) if by_col else by.at(k, 0) for k in range(n)], order))
        perm = _sort_rows(list(range(n)), keys)
        return _pick_cols(ev, a, perm) if by_col else _pick_rows(ev, a, perm)
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


def _flat_positions(ev, a: Arr, ig: int, by_col: bool, room: int):
    """平らにしたときの読み方：(個数, k → 値)。ignore のあるときは、残すマスの番号を持つ（シートに入る数まで）。"""
    n = a.rows * a.cols
    if by_col:
        h = a.rows
        at = lambda k: a.at(k % h, k // h)  # noqa: E731
    else:
        w = a.cols
        at = lambda k: a.at(k // w, k % w)  # noqa: E731
    if not ig:
        if n > room:
            raise err("#NUM!")
        return n, at
    keep = []
    t0: list = []
    for k in range(n):
        if not k % 4096:
            tick("TOCOL", k, n, t0)
        x = at(k)
        if not ((ig in (1, 3) and x is None) or (ig in (2, 3) and isinstance(x, ExcelError))):
            keep.append(k)
            if len(keep) > room:
                raise err("#NUM!")
    if not keep:
        raise err(CALC)
    return len(keep), lambda k: at(keep[k])


def _flatten_big(ev, array, ignore, scan_by_col, room: int):
    a = _grid(ev, array)
    ig = _int(ev, ignore, 0)
    if ig not in (0, 1, 2, 3):
        raise err(VALUE)
    return a, ig, _flat_positions(ev, a, ig, _flag(ev, scan_by_col), room)


def _is_lazy_arg(ev, v) -> bool:
    return isinstance(v, LazyArr) or (isinstance(v, Rng) and isinstance(ev.grid(v), LazyArr))


@fn("TOCOL")
def f_tocol(ev, array, ignore=None, scan_by_col=None):
    if _is_lazy_arg(ev, array):
        a, ig, (n, at) = _flatten_big(ev, array, ignore, scan_by_col, 1048576)
        return make(ev, n, 1, lambda i, _j: at(i), None, None if ig else [("arr", a)])
    return Arr([[x] for x in _flatten(ev, array, ignore, scan_by_col)])


@fn("TOROW")
def f_torow(ev, array, ignore=None, scan_by_col=None):
    if _is_lazy_arg(ev, array):
        a, ig, (n, at) = _flatten_big(ev, array, ignore, scan_by_col, 16384)
        return make(ev, 1, n, lambda _i, j: at(j), None, None if ig else [("arr", a)])
    return Arr([_flatten(ev, array, ignore, scan_by_col)])


@fn("TRANSPOSE")
def f_transpose(ev, array):
    g = _grid(ev, array)
    if want_lazy(g):
        array_room(ev, g.cols, g.rows)
        return make(ev, g.cols, g.rows, lambda i, j: g.at(j, i), g.poly.swapped() if g.poly is not None else None, [("arr", g)] if g.poly is None else None)
    return Arr(_transpose(g.data))


@fn("VSTACK")
def f_vstack(ev, *arrays):
    gs = [_grid(ev, a) for a in arrays]
    width = max(g.cols for g in gs)
    array_room(ev, sum(g.rows for g in gs), width)
    if is_big(sum(g.rows for g in gs), width):
        starts = []
        at = 0
        for g in gs:
            starts.append(at)
            at += g.rows
        na = err(NA)

        def cell(i, j):
            k = bisect_right(starts, i) - 1
            g = gs[k]
            return g.at(i - starts[k], j) if j < g.cols else na

        return make(ev, at, width, cell, None, [("arr", g) for g in gs] + [("const", na, sum(g.rows * (width - g.cols) for g in gs))])
    out = []
    for g in gs:
        for row in g.data:
            out.append(list(row) + [err(NA)] * (width - len(row)))
    return Arr(out)


@fn("HSTACK")
def f_hstack(ev, *arrays):
    gs = [_grid(ev, a) for a in arrays]
    height = max(g.rows for g in gs)
    array_room(ev, height, sum(g.cols for g in gs))
    if is_big(height, sum(g.cols for g in gs)):
        starts = []
        at = 0
        for g in gs:
            starts.append(at)
            at += g.cols
        na = err(NA)

        def cell(i, j):
            k = bisect_right(starts, j) - 1
            g = gs[k]
            return g.at(i, j - starts[k]) if i < g.rows else na

        return make(ev, height, at, cell, None, [("arr", g) for g in gs] + [("const", na, sum((height - g.rows) * g.cols for g in gs))])
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
    array_room(ev, a.rows, len(cols))
    if want_lazy(a, (a.rows, len(cols))):
        return _pick_cols(ev, a, cols)
    return Arr([[row[j] for j in cols] for row in a.data])


@fn("CHOOSEROWS")
def f_chooserows(ev, array, *nums):
    a = _grid(ev, array)
    if not nums:
        raise err(VALUE)
    rows = [_pick_index(int(to_number(x)), a.rows) for n_ in nums for x in ev.values(n_)]
    array_room(ev, len(rows), a.cols)
    if want_lazy(a, (len(rows), a.cols)):
        return _pick_rows(ev, a, rows)
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
    if is_big(len(rr), len(cc)):
        i0, j0 = rr.start, cc.start
        poly = a.poly.substitute(i0, j0) if a.poly is not None else None
        return make(ev, len(rr), len(cc), lambda i, j: a.at(i0 + i, j0 + j), poly)
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
    array_room(ev, r, c)
    p = S(ev, pad) if pad is not None else err(NA)
    if is_big(r, c):
        ar, ac, at = a.rows, a.cols, a.at
        return make(ev, r, c, lambda i, j: at(i, j) if i < ar and j < ac else p, None, [("arr", a), ("const", p, r * c - ar * ac)])
    out = [list(a.data[i]) + [p] * (c - a.cols) if i < a.rows else [p] * c for i in range(r)]
    return Arr(out)


def _vector(ev, v):
    g = _grid(ev, v)
    if g.rows != 1 and g.cols != 1:
        raise err(VALUE)
    if want_lazy(g):
        return ev.values(g)
    return [x for row in g.data for x in row]


@fn("WRAPROWS")
def f_wraprows(ev, vector, count, pad=None):
    xs = _vector(ev, vector)
    n = _int(ev, count)
    if n < 1:
        raise err("#NUM!")
    array_room(ev, (len(xs) + n - 1) // n, n)
    p = S(ev, pad) if pad is not None else err(NA)
    if is_big((len(xs) + n - 1) // n, n):
        m = len(xs)
        r = (m + n - 1) // n
        return make(ev, r, n, lambda i, j: xs[i * n + j] if i * n + j < m else p, None, [("arr", ev.grid(vector)), ("const", p, r * n - m)])
    rows = [xs[i:i + n] for i in range(0, len(xs), n)]
    rows[-1] = rows[-1] + [p] * (n - len(rows[-1]))
    return Arr(rows)


@fn("WRAPCOLS")
def f_wrapcols(ev, vector, count, pad=None):
    got = f_wraprows(ev, vector, count, pad)
    if isinstance(got, LazyArr):
        return f_transpose(ev, got)
    return Arr(_transpose(got.data))


# ================================================================ TEXTSPLIT・MAKEARRAY
def _delims(ev, v) -> list[str]:
    if v is None:
        return []
    return [to_text(x) for x in ev.values(v)] if isinstance(v, (Rng, Arr)) else [to_text(S(ev, v))]


@fn("TEXTSPLIT")
def f_textsplit(ev, text, col_delim=None, row_delim=None, ignore_empty=None, match_mode=None, pad_with=None):
    """TEXTSPLIT：行と列の区切りで字を分ける。大きさは字の数で決まる（行ごとの字の並びだけを持つ）。"""
    import re

    t = to_text(S(ev, text))
    cd, rd = _delims(ev, col_delim), _delims(ev, row_delim)
    if not (cd or rd) or "" in cd + rd:
        raise err(VALUE)
    drop = _flag(ev, ignore_empty)
    flags = re.I if _int(ev, match_mode, 0) == 1 else 0
    pad = S(ev, pad_with) if pad_with is not None else err(NA)

    def split(s: str, ds: list[str]) -> list[str]:
        if not ds:
            return [s]
        parts = re.split("|".join(re.escape(d) for d in sorted(ds, key=len, reverse=True)), s, flags=flags)
        return [p for p in parts if p] if drop else parts

    lines = [split(line, cd) for line in split(t, rd)]
    lines = [x for x in lines if x] if drop else lines
    if not lines:
        raise err(CALC)
    cols = max(len(x) for x in lines)
    array_room(ev, len(lines), cols)
    if is_big(len(lines), cols):
        return make(ev, len(lines), cols, lambda i, j: lines[i][j] if j < len(lines[i]) else pad)
    if len(lines) == 1 and cols == 1:
        return lines[0][0]
    return Arr([x + [pad] * (cols - len(x)) for x in lines])


@fn("MAKEARRAY", lazy=True)
def f_makearray(ev, args):
    """MAKEARRAY(行, 列, LAMBDA(r, c, 式))：マスの式を、頼まれたマスだけ計算する。"""
    from .evaluator import Evaluator
    from .functions import _param
    from .parser import Func, NameNode

    if len(args) != 3 or not isinstance(args[2], Func) or args[2].name != "LAMBDA" or len(args[2].args) != 3:
        return err(VALUE)
    rn, cn, body = args[2].args
    if not isinstance(rn, NameNode) or not isinstance(cn, NameNode):
        return err(VALUE)
    r = _int(ev, ev.ev(args[0]))
    c = _int(ev, ev.ev(args[1]))
    if r < 1 or c < 1:
        return err(VALUE)
    array_room(ev, r, c)
    base = dict(getattr(ev, "scope", None) or {})
    rk, ck = _param(rn.name), _param(cn.name)

    def cell(i, j):
        sub = Evaluator(ev.ctx, ev.sheet, ev.row, ev.col, ev.dr, ev.dc, ev.array)
        sub.scope = {**base, rk: float(i + 1), ck: float(j + 1)}
        v = sub.ev(body)
        return sub.scalar(v)

    if is_big(r, c):
        return make(ev, r, c, cell)
    return Arr([[cell(i, j) for j in range(c)] for i in range(r)])


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
    i = _xfind(ev.values(g), S(ev, key), m, s)
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
    vals = ev.values(lg) if isinstance(lg, LazyArr) else ([r[0] for r in lg.data] if vertical else list(lg.data[0]))
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
    array_room(ev, rows, cols)
    if is_big(rows, cols):
        def cell(r, c):
            a = list(args)
            for i, g in grids.items():
                rr = 0 if g.rows == 1 else r
                cc = 0 if g.cols == 1 else c
                a[i] = g.at(rr, cc) if rr < g.rows and cc < g.cols else err(NA)
            try:
                return f(ev, *a)
            except ExcelError as e:
                return e
            except ZeroDivisionError:
                return err("#DIV/0!")

        return make(ev, rows, cols, cell)
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
