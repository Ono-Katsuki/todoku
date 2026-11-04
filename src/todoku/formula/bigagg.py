"""大きな配列・広い範囲の集計（SUM・COUNT・MAX・COUNTIF…）：配列を持たない。

1. 形の分かる配列は式で答える（lazyarr.stats_of）。どのマスの値も途中の和も double でちょうど表せるとき、式の答えは順に足した答えとビットまで同じ。
   ちょうど表せないとき（2^53 を超える和など）は、Excel と同じに行の順に足す（15 桁目まで Excel と同じになる。実機で確かめた）。
   ただし SEQUENTIAL_CELLS を超える配列は、Excel も持てない（試した版では 1 億マスで 10 GB を超える）ので、正確な和を返して、そう伝える。
2. そのほかは順に読む（lazyarr.scan）。途中まで進んだ所を控えるので、時間の中で終わらなくても続きから計算できる。
"""

from __future__ import annotations

import bisect
from fractions import Fraction

from ..limits import MAX_ARRAY_CELLS
from .lazyarr import FlatView, LazyArr, closed_float, constant_stats, exact_doubles, scan, stats_of
from .values import DIV0, NUM, Arr, ExcelError, Rng, err, is_number, to_number


def any_big(ev, args) -> bool:
    for a in args:
        if isinstance(a, LazyArr):
            return True
        if isinstance(a, Rng):
            r = ev.clip(a)
            if r.rows * r.cols > MAX_ARRAY_CELLS:
                return True
    return False


def _sources(ev, args) -> list:
    """引数：配列はそのまま、直接の値は to_number（SUM(TRUE) は 1）。空は飛ばす。"""
    out = []
    for a in args:
        if isinstance(a, Rng):
            out.append(ev.materialize(a))
        elif isinstance(a, Arr):
            out.append(a)
        elif a is not None:
            out.append(to_number(a))
    return out


def _stats(srcs):
    out = None
    for a in srcs:
        got = stats_of(a) if isinstance(a, Arr) else constant_stats(a, 1)
        if got is None:
            return None
        if out is None:
            out = got
        else:
            if out.error is not None and got.error is not None and got.error != out.error:
                return None
            out = out + got
    return out


def _cells(srcs) -> int:
    return sum(a.rows * a.cols if isinstance(a, Arr) else 1 for a in srcs)


class _Chain:
    """いくつかの引数を 1 つの並びとして読む。"""

    def __init__(self, srcs):
        self.srcs = srcs
        self.starts = []
        at = 0
        self.flats = []
        for a in srcs:
            self.starts.append(at)
            f = FlatView(a) if isinstance(a, Arr) else None
            self.flats.append(f)
            at += a.rows * a.cols if isinstance(a, Arr) else 1
        self.total = at

    def item(self, k: int):
        i = bisect.bisect_right(self.starts, k) - 1
        f = self.flats[i]
        return self.srcs[i] if f is None else f[k - self.starts[i]]

    def positions(self):
        """空のマスを飛ばせるときの、読む位置の並び（引数が 1 つの広い範囲のとき）。"""
        if len(self.srcs) == 1 and isinstance(self.srcs[0], LazyArr):
            return self.srcs[0].nonblank()
        return None


def _fold(ev, what: str, srcs, step, state: list, skip_blank: bool = True) -> list:
    ch = _Chain(srcs)
    return scan(ev, what, ch.total, ch.item, step, state, ch.positions() if skip_blank else None)


def _step_error(x) -> None:
    if isinstance(x, ExcelError):
        raise x


# ---------------------------------------------------------------------- SUM・AVERAGE・MIN・MAX・COUNT
SEQUENTIAL_CELLS = 1 << 28


def _exact_or_small(st, srcs) -> bool:
    """式の答えを使うか：順に足す答えと同じか、順に足す答えに意味がない大きさ（Excel も持てない）とき。"""
    return st.exact or _cells(srcs) > SEQUENTIAL_CELLS


def total(ev, args) -> float:
    srcs = _sources(ev, args)
    st = _stats(srcs)
    if st is not None:
        if st.error is not None:
            raise st.error
        if _exact_or_small(st, srcs):
            if not st.exact:
                _note(ev, "和は正確な値です（順に足す Excel の計算とは、15 桁より下で違うことがあります）")
            return closed_float(st.total)

    def step(s, x):
        if isinstance(x, ExcelError):
            raise x
        if is_number(x):
            s[0] += float(x)

    return _fold(ev, "合計", srcs, step, [0.0])[0]


def _note(ev, text: str) -> None:
    notes = getattr(ev.ctx, "calc_notes", None)
    if notes is not None:
        notes[(ev.sheet, ev.row, ev.col)] = text


def average(ev, args) -> float:
    srcs = _sources(ev, args)
    st = _stats(srcs)
    if st is not None and st.error is None and _exact_or_small(st, srcs):
        if not st.nums:
            raise err(DIV0)
        return closed_float(st.total) / st.nums if st.exact else closed_float(st.total / st.nums)
    if st is not None and st.error is not None:
        raise st.error

    def step(s, x):
        if isinstance(x, ExcelError):
            raise x
        if is_number(x):
            s[0] += float(x)
            s[1] += 1

    s = _fold(ev, "平均", srcs, step, [0.0, 0])
    if not s[1]:
        raise err(DIV0)
    return s[0] / s[1]


def extreme(ev, args, largest: bool) -> float:
    srcs = _sources(ev, args)
    st = _stats(srcs)
    if st is not None:
        if st.error is not None:
            raise st.error
        if not st.nums:
            return 0.0
        v = st.hi if largest else st.lo
        if v is not None:
            return float(v)

    def step(s, x):
        if isinstance(x, ExcelError):
            raise x
        if is_number(x):
            x = float(x)
            if s[0] is None or (x > s[0] if largest else x < s[0]):
                s[0] = x

    got = _fold(ev, "最大" if largest else "最小", srcs, step, [None])[0]
    return 0.0 if got is None else got


def product(ev, args) -> float:
    srcs = _sources(ev, args)

    def step(s, x):
        if isinstance(x, ExcelError):
            raise x
        if is_number(x):
            s[0] *= float(x)
            s[1] = 1

    s = _fold(ev, "積", srcs, step, [1.0, 0])
    return s[0] if s[1] else 0.0


def count(ev, args) -> float:
    srcs = [a if isinstance(a, Arr) else None for a in _arrays_only(ev, args)]
    direct = 0
    for a in args:
        if not isinstance(a, (Rng, Arr)):
            try:
                to_number(a)
                direct += 1
            except ExcelError:
                pass
    arrs = [a for a in srcs if a is not None]
    st = _stats(arrs) if arrs else None
    if st is not None:
        return float(st.nums + direct)
    return float(_fold(ev, "数の数え上げ", arrs, _count_num, [0])[0] + direct)


def _count_num(s, x) -> None:
    if is_number(x):
        s[0] += 1


def _arrays_only(ev, args) -> list:
    return [ev.materialize(a) if isinstance(a, Rng) else a for a in args if isinstance(a, (Rng, Arr))]


def counta(ev, args) -> float:
    arrs = _arrays_only(ev, args)
    direct = sum(1 for a in args if not isinstance(a, (Rng, Arr)))
    st = _stats(arrs) if arrs else None
    if st is not None:
        return float(st.nums + st.text + st.errs + direct)

    def step(s, x):
        if x is not None:
            s[0] += 1

    return float(_fold(ev, "数え上げ", arrs, step, [0])[0] + direct)


def countblank(ev, rng) -> float:
    arrs = _arrays_only(ev, [rng])
    st = _stats(arrs) if arrs else None
    if st is not None:
        return float(st.blank + st.empty)
    a = arrs[0]
    pos = a.nonblank() if isinstance(a, LazyArr) else None
    area = a.rows * a.cols

    def step(s, x):
        if x is None or x == "":
            s[0] += 1

    got = scan(ev, "空の数え上げ", area, FlatView(a).__getitem__, step, [0], pos)[0]
    return float(got + (area - len(pos) if pos is not None else 0))


def variance(ev, args, sample: bool) -> float:
    srcs = _sources(ev, args)
    st = _stats(srcs)
    if st is not None:
        if st.error is not None:
            raise st.error
        n = st.nums
        if n < (2 if sample else 1):
            raise err(DIV0)
        if not st.exact:
            _note(ev, "分散は正確な値です（Excel の計算とは、丸めの分だけ下の桁で違うことがあります）")
        return closed_float((st.total2 - st.total * st.total / n) / (n - 1 if sample else n))

    def step(s, x):
        if isinstance(x, ExcelError):
            raise x
        if is_number(x):
            s[0] += float(x)
            s[1] += 1

    total_, n = _fold(ev, "分散", srcs, step, [0.0, 0])
    if n < (2 if sample else 1):
        raise err(DIV0)
    m = total_ / n

    def step2(s, x):
        if is_number(x):
            s[0] += (float(x) - m) ** 2

    return _fold(ev, "分散", srcs, step2, [0.0])[0] / (n - 1 if sample else n)


# ---------------------------------------------------------------------- LARGE・SMALL
BINS = 4096


def kth(ev, rng, k: int, largest: bool) -> float:
    """大きい方（小さい方）から k 番目。配列を持たない：範囲を区切って数え、k 番目のある区切りへ狭めていく（数回読む）。"""
    srcs = _sources(ev, [rng])
    a = srcs[0] if srcs else None
    st = _stats(srcs)
    if st is not None and st.error is not None:
        raise st.error
    if st is None or st.lo is None or st.hi is None:
        s = _fold(ev, "順位", srcs, _step_range, [None, None, 0])
        lo, hi, n = s
    else:
        lo, hi, n = st.lo, st.hi, st.nums
    if k < 1 or k > n:
        raise err(NUM)
    if lo == hi:
        return float(lo)
    want = n - k + 1 if largest else k  # 小さい方から何番目か
    if isinstance(a, LazyArr) and a.poly is not None:
        got = _kth_progression(a, want)
        if got is not None:
            return got
    below = 0  # lo より小さい数
    lo, hi = float(lo), float(hi)
    while True:
        width = (hi - lo) / BINS
        if width <= 0 or lo + width == lo:
            return lo

        def step(s, x, lo=lo, hi=hi, width=width):
            if is_number(x):
                x = float(x)
                if lo <= x <= hi:
                    s[min(int((x - lo) / width), BINS - 1)] += 1

        counts = _fold(ev, "順位", srcs, step, [0] * BINS)
        at = below
        for b, c in enumerate(counts):
            if at + c >= want:
                break
            at += c
        if counts[b] <= 1 << 16:
            bl, bh = lo + b * width, (lo + (b + 1) * width if b < BINS - 1 else hi)

            def collect(s, x, bl=bl, bh=bh, last=b == BINS - 1):
                if is_number(x):
                    x = float(x)
                    if bl <= x and (x < bh or (last and x <= bh)):
                        s.append(x)

            vals = sorted(_fold(ev, "順位", srcs, collect, []))
            return vals[want - at - 1]
        below = at
        lo, hi = lo + b * width, (lo + (b + 1) * width if b < BINS - 1 else hi)


def _step_range(s, x) -> None:
    if isinstance(x, ExcelError):
        raise x
    if is_number(x):
        x = float(x)
        s[0] = x if s[0] is None or x < s[0] else s[0]
        s[1] = x if s[1] is None or x > s[1] else s[1]
        s[2] += 1


def _kth_progression(a: LazyArr, want: int) -> float | None:
    """等差の並びの配列（SEQUENCE・ROW・COLUMN）の、小さい方から want 番目。"""
    t = a.poly.t
    if not set(t) <= {(0, 0), (1, 0), (0, 1)}:
        return None
    base, b, c = t.get((0, 0), Fraction(0)), t.get((1, 0), Fraction(0)), t.get((0, 1), Fraction(0))
    r, w = a.rows, a.cols
    if c == 0 or w == 1:
        step, repeat, size = b, w, r
    elif b == 0 or r == 1:
        step, repeat, size = c, r, w
    elif b == c * w:
        step, repeat, size = c, 1, r * w
    elif c == b * r:
        step, repeat, size = b, 1, r * w
    else:
        return None
    if step == 0:
        return float(base)
    i = (want - 1) // repeat
    if step < 0:
        i = size - 1 - i
    return float(base + step * i)


# ---------------------------------------------------------------------- SUMPRODUCT
def sumproduct(ev, grids: list) -> float:
    r, c = grids[0].rows, grids[0].cols
    if all(g.poly is not None for g in grids):
        p = grids[0].poly
        for g in grids[1:]:
            p = p * g.poly
        if p.degree() <= 16:
            if exact_doubles(p, r, c) or r * c > SEQUENTIAL_CELLS:
                if not exact_doubles(p, r, c):
                    _note(ev, "和は正確な値です（順に足す Excel の計算とは、15 桁より下で違うことがあります）")
                return closed_float(p.total(r, c))
    if len(grids) == 1:
        return total(ev, grids)
    views = [FlatView(g) for g in grids]

    def item(k):
        return [v[k] for v in views]

    def step(s, xs):
        p = 1.0
        for x in xs:
            if isinstance(x, ExcelError):
                raise x
            p *= float(x) if is_number(x) else 0.0
        s[0] += p

    return scan(ev, "SUMPRODUCT", r * c, item, step, [0.0])[0]


# ---------------------------------------------------------------------- 条件
def where(ev, what: str, pairs: list, vals, step, state: list) -> list:
    """条件の組（並び, 条件）がすべて合う位置の vals を step に渡す。"""
    views = [ev.values(a) for a, _p in pairs]
    n = len(views[0])
    if any(len(v) != n for v in views):
        raise err("#VALUE!")
    if vals is not None and len(vals) != n:
        raise err("#VALUE!")
    preds = [p for _a, p in pairs]
    pos = None
    if len(pairs) == 1 and isinstance(pairs[0][0], (LazyArr, Rng)) and not preds[0](None):
        g = ev.grid(pairs[0][0])
        if isinstance(g, LazyArr):
            pos = g.nonblank()  # 空のマスは条件に合わない：飛ばす

    def item(k):
        for v, p in zip(views, preds):
            if not p(v[k]):
                return _NO
        return vals[k] if vals is not None else _YES

    def run(s, x):
        if x is not _NO:
            step(s, x)

    return scan(ev, what, n, item, run, state, pos)


_NO = object()
_YES = object()


def bools(ev, args) -> list:
    """AND・OR・XOR：[数, 真の数]。"""

    def step(s, x):
        if isinstance(x, ExcelError):
            raise x
        if isinstance(x, bool) or is_number(x):
            s[0] += 1
            s[1] += bool(x)

    state = [0, 0]
    for a in args:
        if isinstance(a, (Rng, Arr)):
            state = _fold(ev, "論理", _arrays_only(ev, [a]), step, state)
        else:
            from .values import to_bool

            state[0] += 1
            state[1] += to_bool(a)
    return state


