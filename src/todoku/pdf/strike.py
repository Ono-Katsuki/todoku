"""PDF：字の上に描いた取り消し線。

PDF の取り消し線は、字の書式ではなく、字の上に描いた線（線・細い塗りの四角）なので、線と字の位置から決める。決まり：
- 行の向きに沿った直線（横書きの行は横の線、縦書きの行は縦の線）で、白くないもの。同じ高さで続く線は 1 本として見る。
  間の空いた線（破線・点線）は、空きが 1 字ぶん（字の大きさ）までなら、ひと続きとして見る
- 線が、字の四角の中央の半分を通る：横書きは、基線の上へ字の大きさの 0.38 倍の所（em の四角の中心。基線の下 0.12・
  上 0.88）から、上下に字の大きさの 1/4 まで。縦書きは、字の四角の左右の中心から左右に 1/4 まで。下線（基線の下）・マスの縁
  （字の四角の端）は入らない。線の太さは、字の大きさの 1/4 まで（それより太いものは、字を塗る帯）
- かかる字：幅の半分より多くが線の下にある字（字ごとに、その字の基線と大きさで確かめる）。ひと続きの線の両端は、かかる字の
  並びの端から 1 字ぶんより外へ出ない。線の長さの半分以上が、かかる字の並び（始めの字から終わりの字まで）の上にある
  （字を横切るだけの長い線・引き出しの線・字より長く続く破線・切り取りの線は入らない）
- 線の両端が、ほかの線・曲線につながらない（枠の辺・表の罫線・吹き出しの辺・直角の点線の 1 つぶは入らない）
- すき間なく並ぶ線（縞で塗った帯）は、並び全体を 1 本の太い線として見る
この決まりに合う線のかかる字が「取り消し線のある字」。合わない線は、何も言わない。
取り消し線と決めた線は、下線の罫線（書く所）にしない（layout.py）。
同じ道のほかの部分と重なる塗りの四角（枠の内と外・輪郭で描いた字の一部）、斜めの線・曲線、画像の中の線は見ない。
線を描いた順（字の下か上か）と、塗りで隠れているかは見ない。どうして取り消したか（意味）は決めない。"""

from __future__ import annotations

import bisect

EM_CENTRE = 0.38  # em の四角の中心：基線の上へ、字の大きさの何倍か（基線の下 0.12・上 0.88）
BAND = 0.25  # 中央の半分：中心から、字の大きさの 1/4 まで。線の太さも、字の大きさの 1/4 まで
JOIN = 0.5  # 同じ線の続きとして 1 本にする、すき間と高さの違い（pt）
TOUCH = 1.0  # 線の端が、ほかの線につながると見る近さ（pt。線の太さの半分を足す）
SLOPE = 0.05  # 行の向きに沿った線と見る傾き（およそ 3 度）まで


def _visible(color) -> bool:
    return not all(c > 0.93 for c in color[:3])


def _joined(rs: list[list[float]]) -> list[list[float]]:
    """[位置, 始め, 終わり, 太さ] の線を、同じ位置で続くものは 1 本に。位置の順に並べて返す。"""
    rs.sort()
    out: list[list[float]] = []
    level: list[list[float]] = []  # 同じ位置の線（位置の違いが JOIN まで）
    for r in rs + [None]:
        if r is not None and (not level or r[0] - level[-1][0] <= JOIN):
            level.append(r)
            continue
        level.sort(key=lambda q: q[1])
        cur = None
        for q in level:
            if cur is not None and q[1] <= cur[2] + JOIN:
                lo, hi = min(cur[0] - cur[3] / 2, q[0] - q[3] / 2), max(cur[0] + cur[3] / 2, q[0] + q[3] / 2)
                cur[0], cur[2], cur[3] = (lo + hi) / 2, max(cur[2], q[2]), hi - lo
            else:
                cur = list(q)
                out.append(cur)
        level = [r] if r is not None else []
    out.sort()
    return out


class _Ends:
    """ほかの線の点：線の端が、ほかの線・曲線につながるかを調べる。"""

    def __init__(self, hs, vs, diagonals, curve_points):
        self.across = {False: vs, True: hs}  # 横の線の端には縦の線、縦の線の端には横の線
        self.keys = {k: [r[0] for r in v] for k, v in self.across.items()}
        self.points: dict[tuple[int, int], list[tuple[float, float]]] = {}
        for p in list(curve_points) + [q for d in diagonals for q in d]:
            self.points.setdefault((int(p[0] // 4), int(p[1] // 4)), []).append(p)

    def touches(self, vertical: bool, pos: float, at: float, width: float, gap: float) -> bool:
        """線（位置 pos・太さ width）の端 at に、直角の線が通る・斜めの線か曲線の点がある。
        直角の線は、この線の太さより外へ出ているものか、同じ並びに gap までの空きで続きのあるもの（直角の点線の 1 つぶ）だけ：
        点線の 1 つぶは、縦の線でも横の線でもあるので、つぶだけでは、ほかの線としない。"""
        tol = TOUCH + width / 2
        rs, keys = self.across[vertical], self.keys[vertical]
        for r in rs[bisect.bisect_left(keys, at - tol - 3):bisect.bisect_right(keys, at + tol + 3)]:
            if abs(r[0] - at) > tol + r[3] / 2 or not r[1] - tol <= pos <= r[2] + tol:
                continue
            if r[1] < pos - width / 2 - 0.3 or r[2] > pos + width / 2 + 0.3:
                return True
            for q in rs[bisect.bisect_left(keys, r[0] - JOIN):bisect.bisect_right(keys, r[0] + JOIN)]:
                if 0 < r[1] - q[2] <= gap or 0 < q[1] - r[2] <= gap:
                    return True
        x, y = (pos, at) if vertical else (at, pos)
        for ix in range(int((x - tol) // 4), int((x + tol) // 4) + 1):
            for iy in range(int((y - tol) // 4), int((y + tol) // 4) + 1):
                for px, py in self.points.get((ix, iy), ()):
                    if abs(px - x) <= tol and abs(py - y) <= tol:
                        return True
        return False


def _stacked(rs: list[list[float]]) -> None:
    """すき間なく並ぶ線（縞で塗った帯：隣の線と太さの分だけ離れ、長さの半分より多くが重なる）は、太さを、並び全体の太さにする。
    rs は位置の順。"""
    group = list(range(len(rs)))  # 並びの先頭の番号
    widest = min(3.0, max((r[3] for r in rs), default=0.0))  # 細い線の縞だけを見る（太い線は、1 本で帯）
    for i, r in enumerate(rs):
        j = i + 1
        while j < len(rs) and rs[j][0] - r[0] <= (r[3] + widest) / 2 + 0.1:
            q = rs[j]
            if q[0] - r[0] <= (r[3] + q[3]) / 2 + 0.1 \
                    and min(r[2], q[2]) - max(r[1], q[1]) > 0.5 * min(r[2] - r[1], q[2] - q[1]):
                group[j] = group[i]
            j += 1
    span: dict[int, list[float]] = {}
    for i, r in enumerate(rs):
        g = span.setdefault(group[i], [r[0] - r[3] / 2, r[0] + r[3] / 2])
        g[0], g[1] = min(g[0], r[0] - r[3] / 2), max(g[1], r[0] + r[3] / 2)
    for i, r in enumerate(rs):
        r[3] = max(r[3], span[group[i]][1] - span[group[i]][0])


def _over(pieces: list[list[float]], lo: float, hi: float) -> float:
    """[lo, hi] のうち、線のかかる長さ。"""
    return sum(max(0.0, min(hi, q[2]) - max(lo, q[1])) for q in pieces)


def find(lay) -> tuple[list[tuple], list[tuple]]:
    """ページの配置（PageLayout）→ 取り消し線のある字：(行, 始め, 終わり) の並び（行の text の位置）と、
    その取り消し線：(縦か, 位置, 始め, 終わり) の並び（pt。左上が原点）。"""
    pc, view = lay.pc, lay.view
    hs: list[list[float]] = []
    vs: list[list[float]] = []
    diagonals = []
    for s in pc.segs:
        if s.part or not _visible(s.color):
            continue
        x0, y0 = view.pt(s.x0, s.y0)
        x1, y1 = view.pt(s.x1, s.y1)
        dx, dy = abs(x1 - x0), abs(y1 - y0)
        if dx > 0 and dy <= SLOPE * dx:
            hs.append([(y0 + y1) / 2, min(x0, x1), max(x0, x1), s.width])
        elif dy > 0 and dx <= SLOPE * dy:
            vs.append([(x0 + x1) / 2, min(y0, y1), max(y0, y1), s.width])
        elif dx > 0 and dy > 0:
            diagonals.append(((x0, y0), (x1, y1)))
    if not hs and not vs:
        return [], []
    hs, vs = _joined(hs), _joined(vs)
    _stacked(hs)
    _stacked(vs)
    keys = {False: [r[0] for r in hs], True: [r[0] for r in vs]}
    rules = {False: hs, True: vs}
    ends = None
    out = []
    lines: list[tuple] = []
    for ln in lay.all_lines():
        rs, ks = rules[ln.vertical], keys[ln.vertical]
        if not rs:
            continue
        # 字（グリフ）ごとの：行の向きの範囲・中央の位置・大きさ
        gl = []
        for c in ln.chars:
            if c.invisible or c.text.isspace() or c.size <= 0:
                continue
            lo, hi = (c.y0, c.y1) if ln.vertical else (c.x0, c.x1)
            if hi > lo:
                # 縦の行の字は、字の四角の左右の中心（横書きの書体の字を積んだ縦書きも、縦書きの書体の字も同じ）
                gl.append((lo, hi, (c.x0 + c.x1) / 2 if ln.vertical else c.base - EM_CENTRE * c.size, c.size))
        if not gl:
            continue
        gl.sort()
        near = rs[bisect.bisect_left(ks, min(g[2] - BAND * g[3] for g in gl)):bisect.bisect_right(ks, max(g[2] + BAND * g[3] for g in gl))]
        a_all, b_all = gl[0][0], max(g[1] for g in gl)
        if not any(r[2] > a_all and r[1] < b_all for r in near):
            continue
        em = max(g[3] for g in gl)
        struck: list[tuple[float, float]] = []  # 取り消し線のある字の範囲
        hulls: list[tuple[float, float]] = []  # その線の範囲（間の空白も、線の下）
        k = 0
        while k < len(near):
            # 同じ位置の線（near は位置の順）を、空きが 1 字ぶんまでのひと続きに分ける
            j = k + 1
            while j < len(near) and near[j][0] - near[j - 1][0] <= JOIN:
                j += 1
            level = sorted(near[k:j], key=lambda q: q[1])
            k = j
            chains: list[list[list[float]]] = []
            for q in level:
                if chains and q[1] <= max(p[2] for p in chains[-1]) + em:
                    chains[-1].append(q)
                else:
                    chains.append([q])
            for chain in chains:
                pos = sum(q[0] for q in chain) / len(chain)
                width = max(q[3] for q in chain)
                a, b = chain[0][1], max(q[2] for q in chain)
                hit = [g for g in gl if abs(pos - g[2]) <= BAND * g[3] and width <= BAND * g[3]
                       and _over(chain, g[0], g[1]) > 0.5 * (g[1] - g[0])]
                if not hit:
                    continue
                g0, g1 = hit[0][0], max(g[1] for g in hit)
                size = max(g[3] for g in hit)
                if g0 - a > size or b - g1 > size:
                    continue  # 字より長く続く線
                if _over(chain, g0, g1) < 0.5 * sum(q[2] - q[1] for q in chain):
                    continue  # 字を横切るだけの線
                if ends is None:
                    ends = _Ends(hs, vs, diagonals, [view.pt(*p[:2]) for path in pc.curve_paths for p in path])
                if ends.touches(ln.vertical, pos, a, width, size) or ends.touches(ln.vertical, pos, b, width, size):
                    continue  # 枠の辺・罫線・吹き出しの辺
                struck += [(g[0], g[1]) for g in hit]
                hulls.append((a, b))
                lines.append((ln.vertical, pos, a, b))
        if not struck:
            continue
        marks = []  # text の字ごと：2 取り消し線のある字・1 線の下の空白・0 ほか
        for ch, (p0, p1) in zip(ln.text, ln.pos):
            mid = (p0 + p1) / 2
            if ch.isspace():
                marks.append(1 if any(a <= mid <= b for a, b in hulls) else 0)
            else:
                marks.append(2 if any(lo - 0.01 <= mid <= hi + 0.01 for lo, hi in struck) else 0)
        i = 0
        while i < len(marks):
            if marks[i] != 2:
                i += 1
                continue
            j = last = i
            while j < len(marks) and marks[j]:
                if marks[j] == 2:
                    last = j
                j += 1
            out.append((ln, i, last + 1))
            i = j
    return out, lines
