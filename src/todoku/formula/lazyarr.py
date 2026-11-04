"""大きな配列：マスを作らず、頼まれたマスだけ計算する。

SEQUENCE・EXPAND・ROW(1:1048576)*COLUMN(A:XFD) のような動的配列は、シートに入るかぎり存在する（超えると #NUM!。Excel の仕様）。
ここでは配列を「(行, 列) → 値」の関数として持つ。作るのに要るのは大きさと計算の組み立てだけで、配列の大きさによらない。

- LazyArr.at(i, j)：1 つのマスを、Excel の要素ごとの計算と同じ浮動小数の計算で求める（0 始まり）
- 集計（SUM・COUNT・MAX…）：
    1. 形の分かる配列（Poly：行・列の番号の多項式）は、式で答える（closed）
    2. そのほかは、少しずつ順に読む（scan）。配列全体を持たない。時間の中で終わらなければ、進んだ所まで控えて Unfinished を投げる
       （控えから続きを計算できる：同じ集計を呼び直すと、控えの所から始まる。長い仕事は jobs.py）

小さな配列（values.LAZY_ABOVE まで）は、いままでどおりリストで作る（ふつうの様式の計算は変わらない）。
"""

from __future__ import annotations

import math
import time
from collections.abc import Sequence
from fractions import Fraction
from typing import Callable

from .. import limits
from . import values as _values
from .values import Arr, ExcelError, Rng, Unfinished, is_number

EXACT = 1 << 53  # これより小さな整数は double でちょうど表せる
SCAN_CHUNK = 4096  # 時間を見る間隔（マス）


class _RowView(Sequence):
    __slots__ = ("arr", "i")

    def __init__(self, arr, i):
        self.arr = arr
        self.i = i

    def __len__(self):
        return self.arr.cols

    def __getitem__(self, j):
        if isinstance(j, slice):
            return [self.arr.at(self.i, k) for k in range(*j.indices(self.arr.cols))]
        if j < 0:
            j += self.arr.cols
        if not 0 <= j < self.arr.cols:
            raise IndexError(j)
        return self.arr.at(self.i, j)

    def __iter__(self):
        a, i = self.arr, self.i
        return (a.at(i, j) for j in range(a.cols))


class _GridView(Sequence):
    """arr.data の代わり：data[i][j] と行の順の読みが、リストと同じに使える（マスは持たない）。"""

    __slots__ = ("arr",)

    def __init__(self, arr):
        self.arr = arr

    def __len__(self):
        return self.arr.rows

    def __getitem__(self, i):
        if isinstance(i, slice):
            return [_RowView(self.arr, k) for k in range(*i.indices(self.arr.rows))]
        if i < 0:
            i += self.arr.rows
        if not 0 <= i < self.arr.rows:
            raise IndexError(i)
        return _RowView(self.arr, i)

    def __iter__(self):
        a, n = self.arr, self.arr.rows
        t0: list = []
        for i in range(n):
            if not i % 256:
                tick("計算", i * a.cols, n * a.cols, t0)
            yield _RowView(a, i)


class FlatView(Sequence):
    """配列を行の順に平らにした並び（ev.values の代わり）。"""

    __slots__ = ("arr", "cols", "n")

    def __init__(self, arr: Arr):
        self.arr = arr
        self.cols = arr.cols
        self.n = arr.rows * arr.cols

    def __len__(self):
        return self.n

    def __getitem__(self, k):
        if isinstance(k, slice):
            return [self[x] for x in range(*k.indices(self.n))]
        if k < 0:
            k += self.n
        if not 0 <= k < self.n:
            raise IndexError(k)
        return self.arr.at(k // self.cols, k % self.cols)

    def __iter__(self):
        a, w, at = self.arr, self.cols, self.arr.at
        t0: list = []
        k = 0
        for i in range(a.rows):
            for j in range(w):
                if not k % SCAN_CHUNK:
                    tick("計算", k, self.n, t0)
                k += 1
                yield at(i, j)


class LazyArr(Arr):
    """fn(i, j) で値を求める配列。poly：行・列の番号の多項式で表せるとき（値はいつも数）。
    parts：同じマスの集まりに分けられるとき（[("arr", 配列) か ("const", 値, 個数)]。順は問わない数えものを、式で答える）。"""

    def __init__(self, rows: int, cols: int, fn: Callable, poly: "Poly | None" = None, parts: list | None = None):
        self._rows = rows
        self._cols = cols
        self._fn = fn
        self.poly = poly
        self.parts = parts

    @property
    def rows(self) -> int:
        return self._rows

    @property
    def cols(self) -> int:
        return self._cols

    @property
    def data(self):
        return _GridView(self)

    def at(self, i: int, j: int):
        return self._fn(i, j)

    def __eq__(self, other):
        return self is other

    __hash__ = None

    def __repr__(self):
        return f"LazyArr({self._rows}×{self._cols})"

    def nonblank(self) -> list | None:
        """空でないマスの平らな位置の並び（分かるとき）。"""
        return None


class RangeArr(LazyArr):
    """シートの広い範囲：マスの値は読むときに取り出す。"""

    def __init__(self, ctx, rng: Rng):
        self.ctx = ctx
        self.rng = rng
        cell_value, sheet, r1, c1 = ctx.cell_value, rng.sheet, rng.r1, rng.c1
        super().__init__(rng.rows, rng.cols, lambda i, j: cell_value(sheet, r1 + i, c1 + j))

    def nonblank(self):
        f = getattr(self.ctx, "cells_in", None)
        if f is None:
            return None
        got = f(self.rng.sheet, self.rng.r1, self.rng.c1, self.rng.r2, self.rng.c2)
        if got is None:
            return None
        r1, c1, w = self.rng.r1, self.rng.c1, self.rng.cols
        return [(r - r1) * w + (c - c1) for r, c in got]


# ---------------------------------------------------------------------- 多項式
def _power_sum(n: int, p: int) -> int:
    """0^p + 1^p + … + (n-1)^p（0^0 は 1）。"""
    if p == 0:
        return n
    # ファウルハーバーの公式（p は小さい）
    total = Fraction(0)
    for k in range(p + 1):
        total += Fraction(math.comb(p + 1, k)) * _bernoulli(k) * Fraction(n) ** (p + 1 - k)
    total /= p + 1
    return int(total)


_BERN: list[Fraction] = [Fraction(1)]


def _bernoulli(m: int) -> Fraction:
    while len(_BERN) <= m:
        n = len(_BERN)
        _BERN.append(-sum((Fraction(math.comb(n + 1, k)) * _BERN[k] for k in range(n)), Fraction(0)) / (n + 1))
    return _BERN[m]


MAX_DEGREE = 8


class Poly:
    """行 i・列 j（0 始まり）の多項式：{(p, q): 係数}。係数は正確な分数。"""

    __slots__ = ("t",)

    def __init__(self, t: dict):
        self.t = {k: v for k, v in t.items() if v != 0}

    @staticmethod
    def const(x) -> "Poly":
        return Poly({(0, 0): Fraction(x)})

    @staticmethod
    def linear(a, b, c) -> "Poly":
        return Poly({(0, 0): Fraction(a), (1, 0): Fraction(b), (0, 1): Fraction(c)})

    def __add__(self, o):
        t = dict(self.t)
        for k, v in o.t.items():
            t[k] = t.get(k, 0) + v
        return Poly(t)

    def __neg__(self):
        return Poly({k: -v for k, v in self.t.items()})

    def __sub__(self, o):
        return self + (-o)

    def __mul__(self, o):
        t: dict = {}
        for (a, b), v in self.t.items():
            for (c, d), w in o.t.items():
                t[(a + c, b + d)] = t.get((a + c, b + d), 0) + v * w
        return Poly(t)

    def degree(self) -> int:
        return max((p + q for p, q in self.t), default=0)

    def scaled(self, f: Fraction) -> "Poly":
        return Poly({k: v * f for k, v in self.t.items()})

    def bilinear(self) -> bool:
        return all(p <= 1 and q <= 1 for p, q in self.t)

    def total(self, rows: int, cols: int) -> Fraction:
        return sum((v * _power_sum(rows, p) * _power_sum(cols, q) for (p, q), v in self.t.items()), Fraction(0))

    def abs_total(self, rows: int, cols: int) -> Fraction:
        return sum((abs(v) * _power_sum(rows, p) * _power_sum(cols, q) for (p, q), v in self.t.items()), Fraction(0))

    def denominator(self) -> int:
        d = 1
        for v in self.t.values():
            d = d * v.denominator // math.gcd(d, v.denominator)
        return d

    def substitute(self, i_shift: int = 0, j_shift: int = 0, i_fixed: int | None = None, j_fixed: int | None = None) -> "Poly":
        """i → i + i_shift（i_fixed があれば i → i_fixed）。列も同じ。"""
        t: dict = {}
        for (p, q), v in self.t.items():
            ip = _binom_terms(p, i_shift) if i_fixed is None else {0: Fraction(i_fixed) ** p}
            jq = _binom_terms(q, j_shift) if j_fixed is None else {0: Fraction(j_fixed) ** q}
            for a, x in ip.items():
                for b, y in jq.items():
                    t[(a, b)] = t.get((a, b), 0) + v * x * y
        return Poly(t)

    def swapped(self) -> "Poly":
        return Poly({(q, p): v for (p, q), v in self.t.items()})


def _binom_terms(p: int, s: int) -> dict:
    return {k: Fraction(math.comb(p, k)) * Fraction(s) ** (p - k) for k in range(p + 1)}


def exact_doubles(poly: Poly, rows: int, cols: int) -> bool:
    """どのマスの値も、どの途中の和も double でちょうど表せるか（順に足した結果と、式の結果が、ビットまで同じになる）。"""
    d = poly.denominator()
    if d & (d - 1):
        return False  # 分母が 2 のべきでない（1/3 など）：マスの値が丸められる
    return poly.abs_total(rows, cols) * d < EXACT


def poly_of(v, rows: int, cols: int) -> Poly | None:
    """値を、rows 行 cols 列に広げたときの多項式（1 行・1 列は、Excel のとおり広げる）。"""
    if isinstance(v, bool) or v is None:
        return None
    if is_number(v):
        try:
            return Poly.const(v)
        except (ValueError, OverflowError):  # inf・nan
            return None
    if isinstance(v, Arr) and v.poly is not None:
        if v.rows == rows and v.cols == cols:
            return v.poly
        if (v.rows == 1 or v.rows == rows) and (v.cols == 1 or v.cols == cols):
            return v.poly.substitute(i_fixed=0 if v.rows == 1 and rows > 1 else None, j_fixed=0 if v.cols == 1 and cols > 1 else None)
        return None
    if isinstance(v, Arr) and v.rows == 1 and v.cols == 1:
        return poly_of(v.at(0, 0), rows, cols)
    return None


def poly_binary(op: str, a, b, rows: int, cols: int) -> Poly | None:
    """配列どうしの計算の結果の多項式（+ - * / ^）。分からなければ None（要素ごとに計算する）。"""
    pa, pb = poly_of(a, rows, cols), poly_of(b, rows, cols)
    if pa is None or pb is None:
        return None
    if op == "+":
        return pa + pb
    if op == "-":
        return pa - pb
    if op == "*":
        out = pa * pb
        return out if out.degree() <= MAX_DEGREE else None
    if op == "/":
        if pb.degree() == 0 and (0, 0) in pb.t:
            return pa.scaled(1 / pb.t[(0, 0)])
        return None
    if op == "^":
        if pb.degree() == 0 and (0, 0) in pb.t and pb.t[(0, 0)].denominator == 1 and 1 <= pb.t[(0, 0)].numerator <= MAX_DEGREE:
            out = pa
            for _ in range(int(pb.t[(0, 0)]) - 1):
                out = out * pa
            return out if out.degree() <= MAX_DEGREE else None
    return None


# ---------------------------------------------------------------------- 式で答える集計
class Stats:
    """配列の数えもの（平らに並べたときの、順によらないもの）。"""

    __slots__ = ("nums", "text", "blank", "empty", "errs", "total", "total2", "lo", "hi", "error", "exact")

    def __init__(self, nums=0, text=0, blank=0, empty=0, errs=0, total=Fraction(0), total2=Fraction(0), lo=None, hi=None, error=None, exact=True):
        self.nums = nums  # 数のマス
        self.text = text  # 空でない、数でないマス（文字・論理値。空の文字 "" も）
        self.blank = blank  # 空のマス
        self.empty = empty  # 空の文字 "" のマス（text の内数）
        self.errs = errs  # エラーのマス
        self.total = total  # 数の和（正確な分数）
        self.total2 = total2  # 数の 2 乗の和
        self.lo = lo  # 最小（分からなければ None）
        self.hi = hi
        self.error = error  # 配列の中のエラー（あれば、集計はこのエラー）
        self.exact = exact  # 順に足す計算と、ビットまで同じか

    def __add__(self, o: "Stats") -> "Stats":
        t = self.total + o.total
        d = t.denominator
        # 足し合わせた和も double でちょうど表せるときだけ、順に足す計算と同じ
        ok = self.exact and o.exact and not (d & (d - 1)) and (abs(self.total) + abs(o.total)) * d < EXACT
        return Stats(self.nums + o.nums, self.text + o.text, self.blank + o.blank, self.empty + o.empty, self.errs + o.errs, t,
                     self.total2 + o.total2,
                     _pick(min, self.lo, o.lo, self.nums, o.nums), _pick(max, self.hi, o.hi, self.nums, o.nums),
                     self.error or o.error, ok)


def _pick(f, a, b, na, nb):
    if not na:
        return b
    if not nb:
        return a
    if a is None or b is None:
        return None
    return f(a, b)


def constant_stats(v, n: int) -> Stats | None:
    """同じ値が n 個。"""
    if n == 0:
        return Stats()
    if isinstance(v, ExcelError):
        return Stats(errs=n, error=v)
    if v is None:
        return Stats(blank=n)
    if is_number(v):
        try:
            x = Fraction(v)
        except (ValueError, OverflowError):
            return None
        d = x.denominator
        return Stats(nums=n, total=x * n, total2=x * x * n, lo=float(v), hi=float(v), exact=not (d & (d - 1)) and abs(x) * n * d < EXACT)
    return Stats(text=n, empty=n if v == "" else 0)


def stats_of(a) -> Stats | None:
    """配列の数えものを、大きな配列のマスを読まずに求める。求められなければ None。"""
    if not isinstance(a, Arr):
        return None
    if a.poly is not None:
        r, c = a.rows, a.cols
        corners = [a.at(i, j) for i in (0, r - 1) for j in (0, c - 1)] if a.poly.bilinear() else None
        return Stats(nums=r * c, total=a.poly.total(r, c), total2=(a.poly * a.poly).total(r, c),
                     lo=min(corners) if corners else None, hi=max(corners) if corners else None, exact=exact_doubles(a.poly, r, c))
    if a.parts is not None:
        out = Stats()
        for p in a.parts:
            got = constant_stats(p[1], p[2]) if p[0] == "const" else stats_of(p[1])
            if got is None:
                return None
            if out.error is not None and got.error is not None and got.error != out.error:
                return None  # 違うエラーが混ざる：どれが先かは順に読まないと分からない
            out = out + got
        return out
    if isinstance(a, LazyArr) or a.rows * a.cols > 1 << 20:
        return None
    out = Stats()
    for row in a.data:
        for x in row:
            got = constant_stats(x, 1)
            if got is None:
                return None
            if out.error is not None and got.error is not None and got.error != out.error:
                return None
            out = out + got
    return out


def closed_float(x: Fraction) -> float:
    try:
        return float(x)
    except OverflowError:
        return math.inf if x > 0 else -math.inf


def make(ev, rows: int, cols: int, fn: Callable, poly=None, parts=None) -> LazyArr:
    """大きな配列を作る。あとで読むときも、ブックの日付の数え方で計算する。"""
    if getattr(ev, "date1904", False):
        from .dates import _CURRENT_1904

        inner = fn

        def fn(i, j):  # noqa: F811
            token = _CURRENT_1904.set(True)
            try:
                return inner(i, j)
            finally:
                _CURRENT_1904.reset(token)
    return LazyArr(rows, cols, fn, poly, parts)


def want_lazy(*arrays_or_sizes) -> bool:
    """作る配列の大きさ (行, 列) か、読む配列を渡す：大きな配列にするか。"""
    for x in arrays_or_sizes:
        if isinstance(x, LazyArr):
            return True
        if isinstance(x, tuple) and _values.is_big(*x):
            return True
    return False


# ---------------------------------------------------------------------- 少しずつ読む集計
class ScanStore:
    """大きな配列の集計の控え（JSON にできる）：鍵 → {"k": 読んだマスの数, "state": 途中の値, "total": 全体の数, "done": 終わったか}。"""

    def __init__(self, data: dict | None = None, frozen: bool = False):
        self.data = data if data is not None else {}
        self.frozen = frozen  # 長い仕事が動いている：読むだけ（ここでは計算を進めない）
        self.dirty = False
        self.on_save: Callable | None = None
        self._last = time.monotonic()

    def get(self, key: str):
        return self.data.get(key)

    def put(self, key: str, rec: dict) -> None:
        self.data[key] = rec
        self.dirty = True
        if self.on_save is not None and time.monotonic() - self._last > 1.0:
            self._last = time.monotonic()
            self.on_save(self.data)


def _store(ctx) -> ScanStore:
    st = getattr(ctx, "scan_store", None)
    if st is None:
        st = ScanStore()
        try:
            ctx.scan_store = st
        except AttributeError:
            pass
    return st


def _key(ev, tag: str, total: int) -> str:
    ctx = ev.ctx
    seq = getattr(ctx, "scan_seq", None)
    if seq is None:
        seq = {}
        try:
            ctx.scan_seq = seq
        except AttributeError:
            pass
    here = (ev.sheet, ev.row, ev.col)
    n = seq.get(here, 0)
    seq[here] = n + 1
    return f"{ev.sheet}!{ev.row},{ev.col}#{n}:{tag}:{total}"


def _clock():
    """いまの呼び出しの、計算に使ってよい時間の残り（秒）を持つもの。"""
    b = limits.budget()
    if b is limits._LOOSE:  # 呼び出しの外：計算ごとに持ち時間
        return limits.Budget()
    return b


def scan(ev, what: str, total: int, item: Callable, step: Callable, state: list, positions=None) -> list:
    """item(k)（k = 0..total-1）を順に読んで step(state, x) を呼ぶ。state は JSON にできるリスト（上書きしてよい）。
    positions：読むマスの位置の並び（空のマスを飛ばすとき）。途中で StopScan を投げると、そこで終わる。
    時間の中で終わらなければ、進んだ所を控えて Unfinished を投げる（同じ集計を呼び直すと、続きから）。"""
    store = _store(ev.ctx)
    key = _key(ev, what, total)
    got = store.get(key)
    if got is not None:
        if got.get("done"):
            return got["state"]
        start, state = got["k"], list(got["state"])
    else:
        start = 0
    n = total if positions is None else len(positions)
    if store.frozen:
        raise Unfinished(what, start if positions is None else _spot(positions, start, total), total, key)
    clock = _clock()
    left = limits.scan_seconds(clock)
    t0 = time.monotonic()
    k = start
    try:
        while k < n:
            stop = min(k + SCAN_CHUNK, n)
            if positions is None:
                for q in range(k, stop):
                    step(state, item(q))
            else:
                for q in positions[k:stop]:
                    step(state, item(q))
            k = stop
            now = time.monotonic()
            if now - t0 > left:
                limits.spend_scan(clock, now - t0)
                store.put(key, {"k": k, "state": state, "total": total, "what": what})
                raise Unfinished(what, k if positions is None else _spot(positions, k, total), total, key)
            if store.on_save is not None:
                store.put(key, {"k": k, "state": state, "total": total, "what": what})
            limits.progress(f"大きな配列の{what}", k, n)
    except StopScan:
        pass
    limits.spend_scan(clock, time.monotonic() - t0)
    store.put(key, {"k": n, "state": state, "total": total, "what": what, "done": True})
    return state


def _spot(positions, k: int, total: int) -> int:
    return total if k >= len(positions) else int(positions[k])


class StopScan(Exception):
    pass


def tick(what: str, done: int, total: int, t0: list) -> None:
    """控えをとらない、長い繰り返しの時間の見張り（FILTER・UNIQUE など。計算し直せば同じ所まで来る）。t0：[始めた時刻, 持ち時間]。"""
    if not t0:
        clock = _clock()
        t0 += [time.monotonic(), limits.scan_seconds(clock), clock]
    if time.monotonic() - t0[0] > t0[1]:
        limits.spend_scan(t0[2], time.monotonic() - t0[0])
        raise Unfinished(what, done, total)
    limits.progress(f"大きな配列の{what}", done, total)
