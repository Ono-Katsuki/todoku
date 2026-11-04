"""どの数式がどのセルを読むか：マスごとの控えを作らず、範囲（区間）のまま持つ。

積み上げの合計（SUM(A$3:A3)・SUM(A$3:A4)…）のように、範囲が重なる数式が並ぶと、マスごとに控えた数は行の数の 2 乗になる。
ここでは範囲を列ごとの区間として持ち、問われた列だけを 1 度なめて（行の順に区間の出入りを数える）、数と最初の数個を答える。
大きさは、数式の参照の数と、問われた列の数で決まる（範囲の広さによらない）。
"""

from __future__ import annotations

import bisect
from collections import defaultdict

from ..limits import MAX_DECLARED_CELLS, budget
from .places import ONE_DECLARATION

NARROW = 64  # これより幅の広い範囲は、列ごとに持たず、問われた列で見る


class Readers:
    """セル (r, c) → {"cells": [読む数式の番地（first 個まで）], "n": 読む数式の数}。dict の get・in・[] と同じに使える。
    places() は、範囲を実際のマスに広げた所（room まで）を返す。1 マスずつにするには広すぎる範囲は、wide に範囲のまま持つ（xlsx/places.py）。"""

    def __init__(self, first: int, room: int, note=None):
        self.first = first
        self.room = room
        self.note = note
        self.wide: list[tuple[int, int, int, int]] = []  # 1 マスずつにしない範囲（places のあとで決まる）
        self.who: list[tuple[str, int, int]] = []  # 数式の順 → 数式のセル
        self.boxes: list[tuple[int, int, int, int, int]] = []
        self._seen: set = set()
        self._narrow: dict[int, list] = defaultdict(list)
        self._wide: list = []
        self._cols: dict[int, tuple[list, list]] = {}
        self._names: dict[int, str] = {}

    def add_formula(self, sheet: str, r: int, c: int) -> int:
        self.who.append((sheet, r, c))
        return len(self.who) - 1

    def add_box(self, r1: int, c1: int, r2: int, c2: int, order: int) -> None:
        k = (r1, c1, r2, c2, order)
        if k in self._seen:
            return
        self._seen.add(k)
        self.boxes.append(k)
        if c2 - c1 < NARROW:
            for c in range(c1, c2 + 1):
                self._narrow[c].append(k)
        else:
            self._wide.append(k)
        self._cols.clear()

    def _column(self, c: int) -> tuple[list, list]:
        """列 c の、行の並びで切った区間：(切れ目の行, その行からの [数, 最初の数個, 答え])。"""
        got = self._cols.get(c)
        if got is not None:
            return got
        if self._wide:
            budget().count(len(self._wide), "数式の広い範囲")
        here = self._narrow.get(c, []) + [b for b in self._wide if b[1] <= c <= b[3]]
        events = defaultdict(list)
        for r1, _c1, r2, _c2, order in here:
            events[r1].append((1, order))
            events[r2 + 1].append((-1, order))
        starts: list[int] = []
        vals: list[list] = []
        active: list[int] = []  # 数式の順に並べる（同じ数式の重なる範囲は、1 つに数える）
        refs: dict[int, int] = {}
        last = None
        for r in sorted(events):
            for d, order in events[r]:
                if d > 0:
                    if refs.get(order, 0) == 0:
                        bisect.insort(active, order)
                    refs[order] = refs.get(order, 0) + 1
                else:
                    refs[order] -= 1
                    if refs[order] == 0:
                        del active[bisect.bisect_left(active, order)]
                        del refs[order]
            head = tuple(active[:self.first])
            if last is not None and last[0] == len(active) and last[1] == head:
                continue
            last = [len(active), head, None]  # 3 番目：答え（問われたときに作る。同じ区間のマスは同じものを返す）
            starts.append(r)
            vals.append(last)
        self._cols[c] = (starts, vals)
        return starts, vals

    def _name(self, order: int) -> str:
        got = self._names.get(order)
        if got is None:
            from .workbook import addr

            s, r, c = self.who[order]
            got = self._names[order] = f"{s}!{addr(r, c)}"
        return got

    def get(self, k: tuple[int, int], default=None):
        r, c = k
        if c not in self._narrow and not self._wide:
            return default
        starts, vals = self._column(c)
        i = bisect.bisect_right(starts, r) - 1
        if i < 0 or vals[i][0] == 0:
            return default
        v = vals[i]
        if v[2] is None:
            v[2] = {"cells": [self._name(o) for o in v[1]], "n": v[0]}
        return v[2]

    def __contains__(self, k) -> bool:
        return self.get(k) is not None

    def __getitem__(self, k):
        v = self.get(k)
        if v is None:
            raise KeyError(k)
        return v

    def places(self) -> list[tuple[int, int]]:
        """範囲を実際のマスに広げた所（重なりは 1 つに）。小さな範囲から room まで広げ、広げなかった範囲は wide に範囲のまま持つ
        （範囲の所：xlsx/places.py。ファイルにセルのあるマスは get で答えられるので、広げなくても読む数式は分かる）。"""
        self.wide = []
        room = self.room
        spans: dict[int, list[list[int]]] = defaultdict(list)  # 列 → 行の、重ならない区間（行の順）
        work = MAX_DECLARED_CELLS * 64  # 広げるために見る列の数の上限（幅の広い違う範囲が何百万も並ぶ、壊れた書類のため）
        same = {b[:4] for b in self.boxes}  # 同じ範囲を読む数式がいくつあっても、広げるのは 1 度
        for r1, c1, r2, c2 in sorted(same, key=lambda b: (b[2] - b[0] + 1) * (b[3] - b[1] + 1)):
            area = (r2 - r1 + 1) * (c2 - c1 + 1)
            if area > min(room, ONE_DECLARATION) or work < c2 - c1 + 1:
                self.wide.append((r1, c1, r2, c2))
                continue
            work -= c2 - c1 + 1
            grow = 0
            saved = []
            for c in range(c1, c2 + 1):
                merged, added = _insert(spans.get(c, []), r1, r2)
                grow += added
                saved.append((c, merged))
            if grow > room:
                self.wide.append((r1, c1, r2, c2))
                continue
            room -= grow
            for c, merged in saved:
                spans[c] = merged
        cells = [(r, c) for c, ivs in spans.items() for a, b in ivs for r in range(a, b + 1)]
        if self.wide:  # 範囲の所の中のマスは、範囲の所として示す（二重に出さない）
            wide = self.wide
            cells = [x for x in cells if not any(a <= x[0] <= b and c <= x[1] <= d for a, c, b, d in wide)]
        return cells


def _insert(ivs: list[list[int]], a: int, b: int) -> tuple[list[list[int]], int]:
    """行の区間の並び ivs に [a, b] を足した並びと、増えたマスの数。"""
    i = bisect.bisect_left(ivs, [a])
    if i > 0 and ivs[i - 1][1] >= a - 1:
        i -= 1
    j = i
    lo, hi = a, b
    had = 0
    while j < len(ivs) and ivs[j][0] <= b + 1:
        lo, hi = min(lo, ivs[j][0]), max(hi, ivs[j][1])
        had += ivs[j][1] - ivs[j][0] + 1
        j += 1
    return ivs[:i] + [[lo, hi]] + ivs[j:], (hi - lo + 1) - had
