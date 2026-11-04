"""区切りの表：字形の表・幅の表のような「番号の範囲 → 値」を、1 つずつに広げずに持つ。

PDF やフォントの 12 バイトの宣言 1 つが、広げると 110 万個の項目になる。広げると、小さなファイルがメモリを使い切る。
ここでは範囲のまま持ち、引くときに二分探索で探す（メモリは宣言の数に比例し、宣言された幅に比例しない）。
範囲が重なるときは、後に書かれたものが勝つ（1 つずつ辞書に入れたのと同じ結果になるよう、重なりを初めに解いておく）。
"""

from __future__ import annotations

import bisect
import heapq
from collections.abc import Iterable, Mapping
from typing import Callable

from ..limits import budget


def flatten(entries: list[tuple[int, int, object]]) -> list[tuple[int, int, object]]:
    """(lo, hi, 値) の並び（後のものが勝つ）を、重ならない、lo の順の並びにする。O(n log n)。"""
    entries = [e for e in entries if e[0] <= e[1]]
    if all(entries[i][1] < entries[i + 1][0] for i in range(len(entries) - 1)):
        return entries  # ふつうの書類：はじめから順に並んでいて重ならない
    starts = sorted(range(len(entries)), key=lambda i: entries[i][0])
    points = sorted({e[0] for e in entries} | {e[1] + 1 for e in entries})
    out: list[tuple[int, int, object]] = []
    heap: list[tuple[int, int]] = []  # (-書かれた順, hi)
    s = 0
    for a, b in zip(points, points[1:]):  # 同じ値が続く小さな区間 [a, b - 1]
        while s < len(starts) and entries[starts[s]][0] <= a:
            i = starts[s]
            heapq.heappush(heap, (-i, entries[i][1]))
            s += 1
        while heap and heap[0][1] < a:
            heapq.heappop(heap)
        if not heap:
            continue
        v = entries[-heap[0][0]][2]
        if out and out[-1][2] is v and out[-1][1] == a - 1:
            out[-1] = (out[-1][0], b - 1, v)
        else:
            out.append((a, b - 1, v))
    return out


class RangeMap(Mapping):
    """(lo, hi, 値) の並びから作る、読むだけの辞書。value：(値, 番号) → その番号の値（None は「ない」）。省くと値そのもの。

    引く（get・in・[]）のは二分探索。items()・keys() は 1 つずつに広げるので、広げる数を予算（budget().count）で数える。"""

    __slots__ = ("_lo", "_hi", "_val", "value")

    def __init__(self, entries: Iterable[tuple[int, int, object]] = (), value: Callable[[object, int], object] | None = None):
        flat = flatten(list(entries))
        self._lo = [e[0] for e in flat]
        self._hi = [e[1] for e in flat]
        self._val = [e[2] for e in flat]
        self.value = value

    def segments(self) -> list[tuple[int, int, object]]:
        return list(zip(self._lo, self._hi, self._val))

    def get(self, k, default=None):
        i = bisect.bisect_right(self._lo, k) - 1
        if i < 0 or k > self._hi[i]:
            return default
        v = self._val[i] if self.value is None else self.value(self._val[i], k)
        return default if v is None else v

    def __getitem__(self, k):
        v = self.get(k, self)
        if v is self:
            raise KeyError(k)
        return v

    def __contains__(self, k) -> bool:
        return self.get(k, self) is not self

    def __bool__(self) -> bool:
        return bool(self._lo)

    def __len__(self) -> int:
        return sum(h - lo + 1 for lo, h in zip(self._lo, self._hi))

    def __iter__(self):
        budget().count(len(self), "範囲の表を広げた項目")
        for lo, hi, val in zip(self._lo, self._hi, self._val):
            for k in range(lo, hi + 1):
                if self.value is None or self.value(val, k) is not None:
                    yield k
