"""図形を画素にする（ISO 32000-1 8.5：塗り・線・切り抜き）。

図形（パス）は命令の組の並び：(0, x, y) 移る・(1, x, y) 直線・(2, x1, y1, x2, y2, x3, y3) 3 次のベジェ・
(3, x1, y1, x2, y2) 2 次のベジェ（TrueType の字の形）・(4,) 閉じる。

塗り（8.5.3.3）：曲線を、画素の大きさで決めた本数の直線に分け、1 画素を縦に S 本の走査線で横切って、
各走査線で塗る区間（非ゼロ巻き数・偶奇の規則）を求め、区間の端は横の長さの割合で数える（なめらかな縁：アンチエイリアス）。
結果は、行ごとの「続き」（x0, x1, 濃さ 0〜255）の並び：同じ濃さの画素をまとめるので、広い塗りも速い。
線（8.5.3.2・8.4.3）：太さ・端の形（突き出さない・丸・四角）・角の形（とがり・丸・切り落とし）・とがりの限度・破線を、
利用者の座標で多角形にしてから写す（座標が縦横で違う倍率でも、仕様どおりの線の形）。1 画素より細い線は 1 画素で描く。
切り抜き（8.5.4）：四角（画素の境目）と、行ごとの続きの型（任意の形）の積。
"""

from __future__ import annotations

import math
import re

S = 4  # 1 画素を横切る走査線の数（縦のなめらかさ）
_INF = 1 << 30


# ====================================================================== 行列
def mul(m1, m2):
    a1, b1, c1, d1, e1, f1 = m1
    a2, b2, c2, d2, e2, f2 = m2
    return (a1 * a2 + b1 * c2, a1 * b2 + b1 * d2, c1 * a2 + d1 * c2, c1 * b2 + d1 * d2,
            e1 * a2 + f1 * c2 + e2, e1 * b2 + f1 * d2 + f2)


def inv(m):
    a, b, c, d, e, f = m
    det = a * d - b * c
    if abs(det) < 1e-14:
        return None
    return (d / det, -b / det, -c / det, a / det, (c * f - d * e) / det, (b * e - a * f) / det)


def scale_of(m) -> float:
    return math.sqrt(abs(m[0] * m[3] - m[1] * m[2]))


# ====================================================================== 曲線を直線に
def _ncurve(dd: float) -> int:
    # 等分した直線と曲線の差は、ほぼ dd / (8 n²)（dd：2 階の差の大きさ。画素）。差を 0.1 画素以下に
    if dd <= 0.4:
        return 1
    n = int(math.sqrt(dd * 1.25)) + 1
    return 64 if n > 64 else n


def flatten(ops, m, keep: bool = False) -> list[tuple[list, bool]]:
    """図形の命令 → (点の並び, 閉じたか) の並び。m：写す行列（曲線を分ける細かさは、写した先の画素で決める）。
    keep：点は写さずに元の座標で返す（線を利用者の座標で作るとき）。"""
    a, b, c, d, e, f = m
    out: list[tuple[list, bool]] = []
    pts: list = []
    closed = False
    cx = cy = sx = sy = 0.0

    def P(x, y):
        return (x, y) if keep else (a * x + c * y + e, b * x + d * y + f)

    for op in ops:
        k = op[0]
        if k == 1:
            cx, cy = op[1], op[2]
            if not pts:
                pts.append(P(sx, sy))
            pts.append(P(cx, cy))
        elif k == 0:
            if len(pts) > 0:
                out.append((pts, closed))
            pts, closed = [P(op[1], op[2])], False
            cx = sx = op[1]
            cy = sy = op[2]
        elif k == 2:
            if not pts:
                pts.append(P(cx, cy))
            x1, y1, x2, y2, x3, y3 = op[1:]
            # 2 階の差（写した先）
            ddx1, ddy1 = cx - 2 * x1 + x2, cy - 2 * y1 + y2
            ddx2, ddy2 = x1 - 2 * x2 + x3, y1 - 2 * y2 + y3
            dd = max(math.hypot(a * ddx1 + c * ddy1, b * ddx1 + d * ddy1), math.hypot(a * ddx2 + c * ddy2, b * ddx2 + d * ddy2))
            n = _ncurve(6 * dd)
            for i in range(1, n):
                t = i / n
                u = 1 - t
                w0, w1, w2, w3 = u * u * u, 3 * u * u * t, 3 * u * t * t, t * t * t
                pts.append(P(w0 * cx + w1 * x1 + w2 * x2 + w3 * x3, w0 * cy + w1 * y1 + w2 * y2 + w3 * y3))
            pts.append(P(x3, y3))
            cx, cy = x3, y3
        elif k == 3:
            if not pts:
                pts.append(P(cx, cy))
            x1, y1, x2, y2 = op[1:]
            ddx, ddy = cx - 2 * x1 + x2, cy - 2 * y1 + y2
            n = _ncurve(2 * math.hypot(a * ddx + c * ddy, b * ddx + d * ddy))
            for i in range(1, n):
                t = i / n
                u = 1 - t
                w0, w1, w2 = u * u, 2 * u * t, t * t
                pts.append(P(w0 * cx + w1 * x1 + w2 * x2, w0 * cy + w1 * y1 + w2 * y2))
            pts.append(P(x2, y2))
            cx, cy = x2, y2
        elif k == 4:
            if pts:
                closed = True
                out.append((pts, True))
                pts, closed = [], False
            cx, cy = sx, sy
    if pts:
        out.append((pts, closed))
    return out


def transform_polys(polys, m):
    a, b, c, d, e, f = m
    return [[(a * x + c * y + e, b * x + d * y + f) for x, y in p] for p in polys]


# ====================================================================== 走査
def _edges(polys) -> list:
    out = []
    for p in polys:
        n = len(p)
        if n < 2:
            continue
        x0, y0 = p[-1]
        for x1, y1 in p:
            if y1 != y0 and x0 == x0 and x1 == x1:  # NaN の点は使わない
                if y0 < y1:
                    out.append((y0, y1, x0, (x1 - x0) / (y1 - y0), 1))
                else:
                    out.append((y1, y0, x1, (x0 - x1) / (y0 - y1), -1))
            x0, y0 = x1, y1
    return out


def _rect_of(p):
    """軸にそろった四角の多角形なら (x0, y0, x1, y1)。"""
    if len(p) == 5 and p[0] == p[4]:
        p = p[:4]
    if len(p) != 4:
        return None
    (ax, ay), (bx, by), (cx, cy), (dx, dy) = p
    if (ay == by and bx == cx and cy == dy and dx == ax) or (ax == bx and by == cy and cx == dx and dy == ay):
        return min(ax, cx), min(ay, cy), max(ax, cx), max(ay, cy)
    return None


def rect_rows(x0, y0, x1, y1, cx0, cy0, cx1, cy1) -> list:
    """軸にそろった四角の塗り（面積で正確な縁の濃さ）。"""
    if x0 < cx0:
        x0 = cx0
    if y0 < cy0:
        y0 = cy0
    if x1 > cx1:
        x1 = cx1
    if y1 > cy1:
        y1 = cy1
    if x1 <= x0 or y1 <= y0:
        return []
    ix0, ix1 = int(math.floor(x0)), int(math.floor(x1))
    iy0, iy1 = int(math.floor(y0)), int(math.ceil(y1))
    rows = []
    for Y in range(iy0, iy1):
        cov = min(y1, Y + 1) - max(y0, Y)
        if cov <= 0:
            continue
        runs = []
        if ix0 == ix1:
            a = int(255 * cov * (x1 - x0) + 0.5)
            if a:
                runs.append((ix0, ix0 + 1, a))
        else:
            a = int(255 * cov * (ix0 + 1 - x0) + 0.5)
            if a:
                runs.append((ix0, ix0 + 1, a))
            if ix1 > ix0 + 1:
                runs.append((ix0 + 1, ix1, int(255 * cov + 0.5)))
            a = int(255 * cov * (x1 - ix1) + 0.5)
            if a:
                runs.append((ix1, ix1 + 1, a))
        if runs:
            rows.append((Y, runs))
    return rows


def scan(polys, evenodd: bool = False, clip=(-_INF, -_INF, _INF, _INF)) -> list:
    """多角形の並び（画素の座標）→ 行ごとの続き [(Y, [(x0, x1, 濃さ)])]。clip：(x0, y0, x1, y1) の整数の四角。"""
    cx0, cy0, cx1, cy1 = clip
    if len(polys) == 1:
        r = _rect_of(polys[0])
        if r is not None:
            return rect_rows(*r, cx0, cy0, cx1, cy1)
    edges = _edges(polys)
    if not edges:
        return []
    edges.sort()
    top = edges[0][0]
    bot = max(e[1] for e in edges)
    ylo = max(cy0, int(math.floor(top)))
    yhi = min(cy1, int(math.ceil(bot)))
    if yhi <= ylo:
        return []
    fx0, fx1 = float(cx0), float(cx1)
    out = []
    active: list = []
    i, n = 0, len(edges)
    # 上の切り抜きより上の辺を飛ばす
    ystart = ylo + 0.5 / S
    while i < n and edges[i][1] <= ystart:
        i += 1
    half = 0.5 / S
    inv_s = 1.0 / S
    for Y in range(ylo, yhi):
        ev: dict = {}
        parts: dict = {}
        hit = False
        for k in range(S):
            y = Y + k * inv_s + half
            while i < n and edges[i][0] <= y:
                if edges[i][1] > y:
                    active.append(edges[i])
                i += 1
            if not active:
                if i >= n:
                    break
                continue
            active = [e for e in active if e[1] > y]
            if not active:
                continue
            xs = sorted([(e[2] + (y - e[0]) * e[3], e[4]) for e in active])
            w = 0
            start = 0.0
            for x, dd in xs:
                if evenodd:
                    if w:
                        w = 0
                        a, b = start, x
                    else:
                        w = 1
                        start = x
                        continue
                else:
                    was = w
                    w += dd
                    if was == 0:
                        start = x
                        continue
                    if w != 0:
                        continue
                    a, b = start, x
                if a < fx0:
                    a = fx0
                if b > fx1:
                    b = fx1
                if b <= a:
                    continue
                hit = True
                ia, ib = int(math.floor(a)), int(math.floor(b))
                if ia == ib:
                    parts[ia] = parts.get(ia, 0.0) + (b - a)
                else:
                    parts[ia] = parts.get(ia, 0.0) + (ia + 1 - a)
                    if ib > ia + 1:
                        ev[ia + 1] = ev.get(ia + 1, 0) + 1
                        ev[ib] = ev.get(ib, 0) - 1
                    if b > ib:
                        parts[ib] = parts.get(ib, 0.0) + (b - ib)
        if not hit:
            if not active and i >= n:
                break
            continue
        keys = set(ev)
        keys.update(parts)
        keys.update([p + 1 for p in parts])
        ks = sorted(keys)
        cnt = 0
        runs = []
        last = len(ks) - 1
        for j, pos in enumerate(ks):
            cnt += ev.get(pos, 0)
            pa = parts.get(pos)
            if pa is not None:
                a = int((cnt + pa) * 255 / S + 0.5)
                if a > 0:
                    if a > 255:
                        a = 255
                    if runs and runs[-1][1] == pos and runs[-1][2] == a:
                        runs[-1] = (runs[-1][0], pos + 1, a)
                    else:
                        runs.append((pos, pos + 1, a))
            elif cnt > 0 and j < last:
                a = 255 if cnt >= S else cnt * 255 // S
                nxt = ks[j + 1]
                if runs and runs[-1][1] == pos and runs[-1][2] == a:
                    runs[-1] = (runs[-1][0], nxt, a)
                else:
                    runs.append((pos, nxt, a))
        if runs:
            out.append((Y, runs))
    return out


def _turn(p) -> float:
    """多角形の向き（面積の符号）。"""
    return sum(p[i - 1][0] * p[i][1] - p[i][0] * p[i - 1][1] for i in range(len(p)))


def _holes(polys, rs) -> bool:
    """向きの違う四角が重なる（巻き数の決まりで、重なりが穴になる：内と外の四角で描いた枠）。"""
    pos = [r for p, r in zip(polys, rs) if _turn(p) > 0]
    neg = [r for p, r in zip(polys, rs) if _turn(p) < 0]
    if not pos or not neg:
        return False
    if len(pos) * len(neg) > 4096:
        return True  # 1 組ずつ調べない：巻き数で数える道（scan）に任せる
    return any(a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3] for a in pos for b in neg)


def fill_rows(polys, evenodd, clip_rect, separate_rects: bool = False, wound: bool = False) -> list:
    """塗り：多角形 → 行ごとの続き。separate_rects：四角だけの並びなら、1 つずつ速く数える（重なりの縁は二重になる）。
    wound：多角形の向きが、書類の道の向きのまま（塗り）。向きの違う四角が重なる並びは、1 つずつ数えない
    （巻き数が 0 になる所は塗らない：8.5.3.3.2）。線の形の四角（どれも塗る）は、向きを見ない。"""
    if separate_rects and len(polys) > 1:
        rs = [_rect_of(p) for p in polys]
        if all(r is not None for r in rs) and not (wound and _holes(polys, rs)):
            rows = []
            for r in rs:
                rows += rect_rows(*r, *clip_rect)
            return rows
    return scan(polys, evenodd, clip_rect)


# ====================================================================== 線
def _dash(subs, arr, phase):
    """破線（8.4.3.6）：オンの区間の、開いた点の並び。"""
    total = sum(arr)
    if total <= 0:
        return subs
    out = []
    for pts, closed in subs:
        if closed:
            pts = pts + [pts[0]]
        # 位相
        k = 0
        rem = arr[0]
        on = True
        ph = phase % total if total > 0 else 0
        while ph > 0:
            if ph >= rem:
                ph -= rem
                k = (k + 1) % len(arr)
                rem = arr[k]
                on = not on
            else:
                rem -= ph
                ph = 0
        cur = [pts[0]] if on else None
        for (x0, y0), (x1, y1) in zip(pts, pts[1:]):
            seg = math.hypot(x1 - x0, y1 - y0)
            pos = 0.0
            while seg - pos > rem:
                pos += rem
                t = pos / seg
                p = (x0 + (x1 - x0) * t, y0 + (y1 - y0) * t)
                if on:
                    cur.append(p)
                    out.append((cur, False))
                    cur = None
                else:
                    cur = [p]
                on = not on
                k = (k + 1) % len(arr)
                rem = arr[k]
                if len(out) > 200000:
                    return out
            rem -= seg - pos
            if on:
                cur.append((x1, y1))
        if on and cur and len(cur) >= 2:
            out.append((cur, False))
    return out


def _ccw(p):
    s = 0.0
    x0, y0 = p[-1]
    for x1, y1 in p:
        s += x0 * y1 - x1 * y0
        x0, y0 = x1, y1
    return p if s >= 0 else p[::-1]


def _circle(x, y, r, n):
    return [(x + r * math.cos(2 * math.pi * i / n), y + r * math.sin(2 * math.pi * i / n)) for i in range(n)]


def stroke_polys(subs, w, cap=0, join=0, miter=10.0, dash=None, phase=0.0, dev_scale=1.0) -> list:
    """線の形の多角形（入力と同じ座標）。subs：flatten の (点の並び, 閉じたか)。dev_scale：この座標の 1 が何画素か（丸の細かさ）。"""
    if dash and any(v > 0 for v in dash):
        subs = _dash(subs, [max(0.0, float(v)) for v in dash], phase)
    hw = w / 2
    nc = max(8, min(64, int(math.pi * hw * dev_scale)))
    out = []
    for pts, closed in subs:
        q = [pts[0]]
        for p in pts[1:]:
            if abs(p[0] - q[-1][0]) > 1e-9 or abs(p[1] - q[-1][1]) > 1e-9:
                q.append(p)
        if closed and len(q) > 2 and abs(q[0][0] - q[-1][0]) < 1e-9 and abs(q[0][1] - q[-1][1]) < 1e-9:
            q.pop()
        if len(q) == 1:
            x, y = q[0]
            if cap == 1:
                out.append(_circle(x, y, hw, nc))
            elif cap == 2:
                out.append([(x - hw, y - hw), (x + hw, y - hw), (x + hw, y + hw), (x - hw, y + hw)])
            continue
        segs = list(zip(q, q[1:]))
        if closed:
            segs.append((q[-1], q[0]))
        dirs = []
        for (x0, y0), (x1, y1) in segs:
            ln = math.hypot(x1 - x0, y1 - y0)
            ux, uy = (x1 - x0) / ln, (y1 - y0) / ln
            dirs.append((ux, uy))
            nx, ny = -uy * hw, ux * hw
            out.append(_ccw([(x0 + nx, y0 + ny), (x1 + nx, y1 + ny), (x1 - nx, y1 - ny), (x0 - nx, y0 - ny)]))
        # 角
        nseg = len(segs)
        joins = range(nseg) if closed else range(1, nseg)
        for j in joins:
            (ux1, uy1), (ux2, uy2) = dirs[j - 1], dirs[j]
            vx, vy = segs[j][0]
            cross = ux1 * uy2 - uy1 * ux2
            dot = ux1 * ux2 + uy1 * uy2
            if abs(cross) < 1e-9 and dot > 0:
                continue
            if join == 1:
                out.append(_circle(vx, vy, hw, nc))
                continue
            sgn = -1.0 if cross > 0 else 1.0  # 外側：左に曲がるなら右側
            o1 = (-uy1 * hw * sgn, ux1 * hw * sgn)
            o2 = (-uy2 * hw * sgn, ux2 * hw * sgn)
            if join == 0:
                sx, sy = o1[0] + o2[0], o1[1] + o2[1]
                sl = math.hypot(sx, sy)
                if sl > 1e-12:
                    cos_half = (sx * o1[0] + sy * o1[1]) / (sl * hw)
                    if cos_half > 1e-9 and 1.0 / cos_half <= miter:
                        L = hw / cos_half
                        out.append(_ccw([(vx, vy), (vx + o1[0], vy + o1[1]), (vx + sx / sl * L, vy + sy / sl * L),
                                         (vx + o2[0], vy + o2[1])]))
                        continue
            out.append(_ccw([(vx, vy), (vx + o1[0], vy + o1[1]), (vx + o2[0], vy + o2[1])]))
        if not closed:
            for (x, y), (ux, uy) in ((q[0], (-dirs[0][0], -dirs[0][1])), (q[-1], dirs[-1])):
                if cap == 1:
                    out.append(_circle(x, y, hw, nc))
                elif cap == 2:
                    nx, ny = -uy * hw, ux * hw
                    ex, ey = x + ux * hw, y + uy * hw
                    out.append(_ccw([(x + nx, y + ny), (ex + nx, ey + ny), (ex - nx, ey - ny), (x - nx, y - ny)]))
    return out


# ====================================================================== 切り抜き
class Clip:
    """切り抜き：整数の四角（x0 ≤ x < x1, y0 ≤ y < y1）と、任意の形（行 → 続き）。"""

    __slots__ = ("x0", "y0", "x1", "y1", "mask")

    def __init__(self, x0, y0, x1, y1, mask=None):
        self.x0, self.y0, self.x1, self.y1 = x0, y0, x1, y1
        self.mask = mask

    @property
    def rect(self):
        return (self.x0, self.y0, self.x1, self.y1)

    @property
    def empty(self) -> bool:
        return self.x1 <= self.x0 or self.y1 <= self.y0

    def with_rect(self, x0, y0, x1, y1) -> "Clip":
        x0, y0 = max(self.x0, int(round(x0))), max(self.y0, int(round(y0)))
        x1, y1 = min(self.x1, int(round(x1))), min(self.y1, int(round(y1)))
        return Clip(x0, y0, max(x0, x1), max(y0, y1), self.mask)

    def with_rows(self, rows) -> "Clip":
        mask = {}
        for Y, runs in rows:
            if Y < self.y0 or Y >= self.y1:
                continue
            if self.mask is not None:
                old = self.mask.get(Y)
                if not old:
                    continue
                runs = intersect(runs, old)
            if runs:
                mask[Y] = runs
        if not mask:
            return Clip(0, 0, 0, 0, {})
        ys = list(mask)
        x0 = min(r[0][0] for r in mask.values())
        x1 = max(r[-1][1] for r in mask.values())
        return Clip(max(self.x0, x0), max(self.y0, min(ys)), min(self.x1, x1), min(self.y1, max(ys) + 1), mask)


def intersect(r1, r2, dx: int = 0) -> list:
    """2 つの続きの並びの重なり（濃さは掛ける）。r1 は dx だけずらして読む。"""
    out = []
    i = j = 0
    n1, n2 = len(r1), len(r2)
    while i < n1 and j < n2:
        a0, a1, aa = r1[i]
        a0 += dx
        a1 += dx
        b0, b1, ba = r2[j]
        lo = a0 if a0 > b0 else b0
        hi = a1 if a1 < b1 else b1
        if lo < hi:
            out.append((lo, hi, aa if ba >= 255 else ba if aa >= 255 else aa * ba // 255))
        if a1 < b1:
            i += 1
        else:
            j += 1
    return out


# ====================================================================== 塗る
class Canvas:
    def __init__(self, w: int, h: int, cover: bool = False):
        self.w, self.h = w, h
        self.buf = bytearray(b"\xff" * (w * h * 3))
        # 画素ごとの、描いたものの濃さ（0 は、まだ何も描いていない紙）。下の色の有無で結果の変わる重ね方（BM）のあるページだけ
        # 持つ（ふつうは None：塗るたびに控える手間をかけない）
        self.cover = bytearray(w * h) if cover else None

    def copy(self) -> "Canvas":
        """同じ中身の、別の絵。"""
        cv = Canvas.__new__(Canvas)
        cv.w, cv.h = self.w, self.h
        cv.buf = bytearray(self.buf)
        cv.cover = bytearray(self.cover) if self.cover is not None else None
        return cv


# ------------------------------------------------------------------ 重ね方（BM：ISO 32000-1 11.3.5）
def _hard(cb: float, cs: float) -> float:
    return cb * 2 * cs if cs <= 0.5 else cb + (2 * cs - 1) - cb * (2 * cs - 1)


def _soft(cb: float, cs: float) -> float:
    if cs <= 0.5:
        return cb - (1 - 2 * cs) * cb * (1 - cb)
    d = ((16 * cb - 12) * cb + 4) * cb if cb <= 0.25 else math.sqrt(cb)
    return cb + (2 * cs - 1) * (d - cb)


# 色の成分ごとに重ねる重ね方：下の色 cb・上の色 cs（0〜1）→ 重ねた色
_SEPARABLE = {
    "Multiply": lambda cb, cs: cb * cs,
    "Screen": lambda cb, cs: cb + cs - cb * cs,
    "Overlay": lambda cb, cs: _hard(cs, cb),
    "Darken": min,
    "Lighten": max,
    "ColorDodge": lambda cb, cs: 0.0 if cb <= 0 else 1.0 if cs >= 1 else min(1.0, cb / (1 - cs)),
    "ColorBurn": lambda cb, cs: 1.0 if cb >= 1 else 0.0 if cs <= 0 else 1 - min(1.0, (1 - cb) / cs),
    "HardLight": _hard,
    "SoftLight": _soft,
    "Difference": lambda cb, cs: abs(cb - cs),
    "Exclusion": lambda cb, cs: cb + cs - 2 * cb * cs,
}


def _lum(r: float, g: float, b: float) -> float:
    return 0.3 * r + 0.59 * g + 0.11 * b


def _set_lum(r: float, g: float, b: float, lum: float) -> tuple:
    d = lum - _lum(r, g, b)
    r, g, b = r + d, g + d, b + d
    lo, hi = min(r, g, b), max(r, g, b)
    if lo < 0:
        r, g, b = (lum + (c - lum) * lum / (lum - lo) for c in (r, g, b))
    if hi > 255:
        r, g, b = (lum + (c - lum) * (255 - lum) / (hi - lum) for c in (r, g, b))
    return r, g, b


def _set_sat(r: float, g: float, b: float, sat: float) -> tuple:
    c = [r, g, b]
    lo, mid, hi = sorted(range(3), key=c.__getitem__)
    if c[hi] > c[lo]:
        c[mid] = (c[mid] - c[lo]) * sat / (c[hi] - c[lo])
        c[hi] = sat
    else:
        c[mid] = c[hi] = 0.0
    c[lo] = 0.0
    return c[0], c[1], c[2]


def _sat(r: float, g: float, b: float) -> float:
    return max(r, g, b) - min(r, g, b)


# 3 つの成分を一緒に重ねる重ね方（11.3.5.3）：下の色 3 つ・上の色 3 つ（0〜255）→ 重ねた色
_NONSEPARABLE = {
    "Hue": lambda br, bg, bb, r, g, b: _set_lum(*_set_sat(r, g, b, _sat(br, bg, bb)), _lum(br, bg, bb)),
    "Saturation": lambda br, bg, bb, r, g, b: _set_lum(*_set_sat(br, bg, bb, _sat(r, g, b)), _lum(br, bg, bb)),
    "Color": lambda br, bg, bb, r, g, b: _set_lum(r, g, b, _lum(br, bg, bb)),
    "Luminosity": lambda br, bg, bb, r, g, b: _set_lum(br, bg, bb, _lum(r, g, b)),
}
BLEND_NAMES = frozenset(_SEPARABLE) | frozenset(_NONSEPARABLE)


class Blend:
    """重ね方（ExtGState の BM）の 1 つ。table：成分ごとの重ね方の表（上の色 × 256 + 下の色 → 重ねた色）。成分を一緒に重ねる
    重ね方は fn。下の色は、絵のその画素の色（ページの絵は、白い紙から始まる、透けない絵）。"""

    __slots__ = ("name", "table", "fn", "needs_cover")

    def __init__(self, name: str):
        self.name = name
        # まだ何も描いていない所では、重ねずに、そのままの色で塗る（11.3.6）。Multiply・Darken は、白い紙と重ねても同じ色
        self.needs_cover = name not in ("Multiply", "Darken")
        f = _SEPARABLE.get(name)
        self.fn = _NONSEPARABLE.get(name)
        self.table = None
        if f is not None:
            self.table = bytes(min(255, max(0, int(f(cb / 255, cs / 255) * 255 + 0.5))) for cs in range(256) for cb in range(256))

    def mix(self, br: int, bg: int, bb: int, r: int, g: int, b: int) -> tuple:
        t = self.table
        if t is not None:
            return t[r << 8 | br], t[g << 8 | bg], t[b << 8 | bb]
        r, g, b = self.fn(br, bg, bb, r, g, b)
        return (0 if r < 0 else 255 if r > 255 else int(r + 0.5), 0 if g < 0 else 255 if g > 255 else int(g + 0.5),
                0 if b < 0 else 255 if b > 255 else int(b + 0.5))


    def onto(self, br: int, bg: int, bb: int, ab: int, r: int, g: int, b: int) -> tuple:
        """下に描いたものの濃さが ab（0〜255）の画素に塗る色：(1 − αb)·Cs + αb·B(Cb, Cs)（11.3.6）。絵の色（br, bg, bb）は、
        下に描いたものを白い紙に重ねた色。"""
        if ab <= 0:
            return r, g, b
        if ab >= 255:
            return self.mix(br, bg, bb, r, g, b)
        w = (255 - ab) * 255
        m = self.mix(*(0 if v * 255 < w else min(255, (v * 255 - w) // ab) for v in (br, bg, bb)), r, g, b)
        na = 255 - ab
        return (r * na + m[0] * ab + 127) // 255, (g * na + m[1] * ab + 127) // 255, (b * na + m[2] * ab + 127) // 255


_BLENDS: dict = {}
_OVER: dict = {}
_COVER_RUNS = re.compile(rb"\xff+|\x00+|[^\x00\xff]")


def _over(a: int) -> bytes:
    """濃さ a で塗った後の「描いたものの濃さ」の表（前の濃さ → 後の濃さ）。"""
    t = _OVER.get(a)
    if t is None:
        t = _OVER[a] = bytes(c + ((255 - c) * a + 127) // 255 for c in range(256))
    return t


def blend_of(name: str) -> Blend | None:
    """名前の重ね方（Normal・Compatible・知らない名前は None）。"""
    if name not in BLEND_NAMES:
        return None
    if name not in _BLENDS:
        _BLENDS[name] = Blend(name)
    return _BLENDS[name]


def _paint_blend(cv: Canvas, rows, rgb, alpha: float, clip: Clip, dx: int, dy: int, blend: Blend) -> None:
    """paint と同じ所を、重ね方 blend で塗る：画素ごとに、下の色と重ねた色を、濃さで混ぜる。"""
    buf = cv.buf
    W = cv.w
    cx0, cy0, cx1, cy1 = clip.x0, clip.y0, clip.x1, clip.y1
    mask = clip.mask
    op = alpha >= 0.999
    t = blend.table
    full = [t[c << 8:(c << 8) + 256] for c in rgb] if t is not None else None
    luts: dict = {}
    plain: dict = {}
    seen: dict = {}
    r, g, b = rgb
    cover = cv.cover
    for Y, runs in rows:
        Y += dy
        if Y < cy0 or Y >= cy1:
            continue
        if mask is not None:
            mr = mask.get(Y)
            if not mr:
                continue
            runs = intersect(runs, mr, dx)
            ddx = 0
        else:
            ddx = dx
        base = Y * W
        for x0, x1, a in runs:
            x0 += ddx
            x1 += ddx
            if x0 < cx0:
                x0 = cx0
            if x1 > cx1:
                x1 = cx1
            if x1 <= x0:
                continue
            if not op:
                a = int(a * alpha + 0.5)
            if a <= 0:
                continue
            if a > 255:
                a = 255
            na = 255 - a
            # 下に描いたものの濃さで分ける：何も描いていない所（0）・透けずに描いてある所（255）・その間（縁の画素）
            parts = ((x0, x1, 255),) if cover is None else [
                (x0 + m.start(), x0 + m.end(), m.group()[0]) for m in _COVER_RUNS.finditer(bytes(cover[base + x0:base + x1]))]
            for s0, s1, ab in parts:
                p, q = (base + s0) * 3, (base + s1) * 3
                if ab == 0:  # 下の色がない：ふつうに塗る
                    lut = plain.get(a)
                    if lut is None:
                        lut = plain[a] = [bytes((cb * na + c * a + 127) // 255 for cb in range(256)) for c in rgb]
                elif ab == 255 and full is not None:  # 成分ごとの重ね方：下の色 → 塗った後の色の表を、成分ごとに引く
                    lut = luts.get(a)
                    if lut is None:
                        lut = luts[a] = full if a >= 255 else [bytes((cb * na + f[cb] * a + 127) // 255 for cb in range(256)) for f in full]
                else:
                    for p in range(p, q, 3):
                        under = bytes(buf[p:p + 3])
                        c = seen.get(under) if ab == 255 else None
                        if c is None:
                            c = blend.onto(under[0], under[1], under[2], ab, r, g, b)
                            if ab == 255:
                                seen[under] = c
                        buf[p] = (under[0] * na + c[0] * a + 127) // 255
                        buf[p + 1] = (under[1] * na + c[1] * a + 127) // 255
                        buf[p + 2] = (under[2] * na + c[2] * a + 127) // 255
                    continue
                buf[p:q:3] = buf[p:q:3].translate(lut[0])
                buf[p + 1:q:3] = buf[p + 1:q:3].translate(lut[1])
                buf[p + 2:q:3] = buf[p + 2:q:3].translate(lut[2])
            if cover is not None:
                cover[base + x0:base + x1] = b"\xff" * (x1 - x0) if a >= 255 else cover[base + x0:base + x1].translate(_over(a))


def paint(cv: Canvas, rows, rgb, alpha: float, clip: Clip, dx: int = 0, dy: int = 0, blend: Blend | None = None) -> None:
    """続きを、色 rgb（0〜255 の 3 つ）・不透明度 alpha で、切り抜きの中に塗る。dx, dy：続きをずらす量（字の型）。
    blend：重ね方（None はふつうの重ね方）。"""
    if blend is not None:
        _paint_blend(cv, rows, rgb, alpha, clip, dx, dy, blend)
        return
    buf = cv.buf
    W = cv.w
    cx0, cy0, cx1, cy1 = clip.x0, clip.y0, clip.x1, clip.y1
    mask = clip.mask
    r, g, b = rgb
    full = bytes((r, g, b))
    op = alpha >= 0.999
    cover = cv.cover
    for Y, runs in rows:
        Y += dy
        if Y < cy0 or Y >= cy1:
            continue
        if mask is not None:
            mr = mask.get(Y)
            if not mr:
                continue
            runs = intersect(runs, mr, dx)
            ddx = 0
        else:
            ddx = dx
        base = Y * W
        for x0, x1, a in runs:
            x0 += ddx
            x1 += ddx
            if x0 < cx0:
                x0 = cx0
            if x1 > cx1:
                x1 = cx1
            if x1 <= x0:
                continue
            if not op:
                a = int(a * alpha + 0.5)
            if a >= 255:
                buf[(base + x0) * 3:(base + x1) * 3] = full * (x1 - x0)
                if cover is not None:
                    cover[base + x0:base + x1] = b"\xff" * (x1 - x0)
            elif a > 0:
                na = 255 - a
                ra, ga, ba = r * a + 127, g * a + 127, b * a + 127
                p = (base + x0) * 3
                for p in range(p, p + 3 * (x1 - x0), 3):
                    buf[p] = (buf[p] * na + ra) // 255
                    buf[p + 1] = (buf[p + 1] * na + ga) // 255
                    buf[p + 2] = (buf[p + 2] * na + ba) // 255
                if cover is not None:
                    cover[base + x0:base + x1] = cover[base + x0:base + x1].translate(_over(a))


def _mix_same(under: bytes, over: bytes, a: int) -> bytes:
    """同じ長さの 2 つの画素の並びを、同じ濃さ a（1〜254）で混ぜる：(under × (255 − a) + over × a + 127) // 255。
    1 バイトずつ回さずに、1 バイトを 16 ビットの枠に置いた大きな整数で、一度に計算する（結果は、1 バイトずつの計算と同じ）。"""
    n = len(under)
    lane = bytearray(2 * n)
    lane[1::2] = under
    x = int.from_bytes(lane, "big")
    lane[1::2] = over
    y = int.from_bytes(lane, "big")
    one = int.from_bytes(b"\x00\x01" * n, "big")
    v = x * (255 - a) + y * a + one * 127  # 枠ごとに 65,152 以下
    v += one + ((v >> 8) & (one * 255))  # v // 255 = (v + 1 + (v >> 8)) >> 8（0 ≤ v < 65,535）
    return v.to_bytes(2 * n, "big")[0::2]


def lay(cv: Canvas, src: Canvas, rect, alpha: float, clip: Clip, blend: Blend | None = None) -> None:
    """別に描いた絵 src（透けて描くまとまり：buf は白い紙に重ねた色、cover は画素ごとの濃さ）を、cv に重ねる。
    rect：src に描いた範囲（画素の四角）。alpha：まとまりの全体の不透明度。clip：切り抜き（透けのマスクも）。
    blend：まとまりを置いたときの重ね方（None はふつうの重ね方）。"""
    W = cv.w
    x0, y0, x1, y1 = max(rect[0], clip.x0, 0), max(rect[1], clip.y0, 0), min(rect[2], clip.x1, W), min(rect[3], clip.y1, cv.h)
    if x1 <= x0 or y1 <= y0:
        return
    buf, sbuf, scov, cover = cv.buf, src.buf, src.cover, cv.cover
    mask = clip.mask
    ga = 255 if alpha >= 0.999 else int(alpha * 255 + 0.5)
    table = blend.table if blend is not None else None
    whole = ((x0, x1, 255),)
    nothing = bytes(x1 - x0)
    for Y in range(y0, y1):
        base = Y * W
        cov = bytes(scov[base + x0:base + x1])
        if cov == nothing:
            continue
        runs = whole
        if mask is not None:
            runs = mask.get(Y)
            if not runs:
                continue
        for m0, m1, ma in runs:
            if m0 < x0:
                m0 = x0
            if m1 > x1:
                m1 = x1
            if m1 <= m0:
                continue
            if ga != 255:
                ma = (ma * ga + 127) // 255
            if ma <= 0:
                continue
            for m in _COVER_RUNS.finditer(cov, m0 - x0, m1 - x0):
                c = cov[m.start()]
                if c == 0:
                    continue
                s0, s1 = base + x0 + m.start(), base + x0 + m.end()
                a = ma if c == 255 else (c * ma + 127) // 255
                if a <= 0:
                    continue
                p, q = s0 * 3, s1 * 3
                if c == 255 and blend is None:  # まとまりが透けずに描いた所を、ふつうに重ねる：行のまま
                    buf[p:q] = sbuf[p:q] if a == 255 else _mix_same(bytes(buf[p:q]), bytes(sbuf[p:q]), a)
                else:
                    na = 255 - a
                    w = (255 - c) * 255
                    for i in range(s0, s1):
                        p = i * 3
                        r, g, b = sbuf[p], sbuf[p + 1], sbuf[p + 2]
                        if c != 255:  # 白い紙に重ねた色から、まとまりの色に戻す
                            r, g, b = (0 if v * 255 < w else min(255, (v * 255 - w) // c) for v in (r, g, b))
                        if blend is not None:
                            ab = 255 if cover is None else cover[i]
                            if ab == 255 and table is not None:
                                r, g, b = table[r << 8 | buf[p]], table[g << 8 | buf[p + 1]], table[b << 8 | buf[p + 2]]
                            else:
                                r, g, b = blend.onto(buf[p], buf[p + 1], buf[p + 2], ab, r, g, b)
                        buf[p] = (buf[p] * na + r * a + 127) // 255
                        buf[p + 1] = (buf[p + 1] * na + g * a + 127) // 255
                        buf[p + 2] = (buf[p + 2] * na + b * a + 127) // 255
                if cover is not None:
                    cover[s0:s1] = b"\xff" * (s1 - s0) if a == 255 else cover[s0:s1].translate(_over(a))


def lift(part: Canvas, over: Canvas, under: Canvas, rect) -> None:
    """下と切り離さないまとまりを、下の絵 under の写しに描いた絵 over から、まとまりだけの色を取り出して part に入れる
    （11.4.8：C = Cn + (Cn − C0)·(α0/αg − α0)）。part.cover は、まとまりだけを描いたときの濃さ αg（先に描いてある）。
    どの絵の buf も、白い紙に重ねた色。"""
    W = part.w
    x0, y0, x1, y1 = max(rect[0], 0), max(rect[1], 0), min(rect[2], W), min(rect[3], part.h)
    if x1 <= x0 or y1 <= y0:
        return
    pbuf, pcov, obuf, ocov, ubuf, ucov = part.buf, part.cover, over.buf, over.cover, under.buf, under.cover
    nothing = bytes(x1 - x0)
    for Y in range(y0, y1):
        base = Y * W + x0
        cov = bytes(pcov[base:base + x1 - x0])
        if cov == nothing:
            continue
        for m in _COVER_RUNS.finditer(cov):
            c = cov[m.start()]
            if c == 0:
                continue
            i = base + m.start()
            p = i * 3
            if c == 255:  # まとまりが透けない所：重ねた後の色が、まとまりの色
                q = (base + m.end()) * 3
                pbuf[p:q] = obuf[p:q]
                continue
            a0 = 255 if ucov is None else ucov[i]
            an = 255 if ocov is None else ocov[i]
            if a0 == 0 or an == 0:  # 下に何もない：写しに描いた色が、そのまま、まとまりの色
                pbuf[p:p + 3] = obuf[p:p + 3]
                continue
            k = a0 / c - a0 / 255
            for j in (p, p + 1, p + 2):
                c0 = ubuf[j] if a0 == 255 else (ubuf[j] - 255 + a0) * 255 / a0  # 白い紙に重ねた色から、色に戻す
                cn = obuf[j] if an == 255 else (obuf[j] - 255 + an) * 255 / an
                v = cn + (cn - c0) * k
                pbuf[j] = int((0.0 if v < 0 else 255.0 if v > 255 else v) * c / 255 + 255 - c + 0.5)


def lay_over(cv: Canvas, src: Canvas, rect, alpha: float, clip: Clip) -> None:
    """cv の写しの上に描いた絵 src（下と切り離さない、透けて描くまとまり）を、cv に重ねる：画素ごとに、前の色と、まとまりを
    描いた後の色を、まとまりの全体の不透明度 alpha（と切り抜きの濃さ）で混ぜる（11.4.8 の式は、紙の白に重ねた色では、この形）。"""
    W = cv.w
    x0, y0, x1, y1 = max(rect[0], clip.x0, 0), max(rect[1], clip.y0, 0), min(rect[2], clip.x1, W), min(rect[3], clip.y1, cv.h)
    if x1 <= x0 or y1 <= y0:
        return
    buf, sbuf, cover, scov = cv.buf, src.buf, cv.cover, src.cover
    mask = clip.mask
    ga = 255 if alpha >= 0.999 else int(alpha * 255 + 0.5)
    whole = ((x0, x1, 255),)
    for Y in range(y0, y1):
        base = Y * W
        if buf[(base + x0) * 3:(base + x1) * 3] == sbuf[(base + x0) * 3:(base + x1) * 3] and (
                cover is None or cover[base + x0:base + x1] == scov[base + x0:base + x1]):
            continue  # まとまりが何も描かなかった行
        runs = whole
        if mask is not None:
            runs = mask.get(Y)
            if not runs:
                continue
        for m0, m1, a in runs:
            if m0 < x0:
                m0 = x0
            if m1 > x1:
                m1 = x1
            if m1 <= m0:
                continue
            if ga != 255:
                a = (a * ga + 127) // 255
            if a <= 0:
                continue
            p, q = (base + m0) * 3, (base + m1) * 3
            buf[p:q] = sbuf[p:q] if a == 255 else _mix_same(bytes(buf[p:q]), bytes(sbuf[p:q]), a)
            if cover is not None:
                p, q = base + m0, base + m1
                cover[p:q] = scov[p:q] if a == 255 else _mix_same(bytes(cover[p:q]), bytes(scov[p:q]), a)


def covered(cv: Canvas, rows, clip: Clip, dx: int = 0, dy: int = 0) -> list:
    """paint が塗る画素：(画素の番号（行 × 幅 + 列）, 墨の濃さ 1〜255) の並び。切り抜き・ずらし方は paint と同じ。
    塗る前の色（buf の 番号 × 3 から 3 つ）を見るために使う（書いた字が下地に紛れていないかの確かめ：pdf/inkcheck）。"""
    W = cv.w
    cx0, cy0, cx1, cy1 = clip.x0, clip.y0, clip.x1, clip.y1
    mask = clip.mask
    out = []
    for Y, runs in rows:
        Y += dy
        if Y < cy0 or Y >= cy1:
            continue
        if mask is not None:
            mr = mask.get(Y)
            if not mr:
                continue
            runs = intersect(runs, mr, dx)
            ddx = 0
        else:
            ddx = dx
        base = Y * W
        for x0, x1, a in runs:
            x0 += ddx
            x1 += ddx
            if x0 < cx0:
                x0 = cx0
            if x1 > cx1:
                x1 = cx1
            if x1 > x0 and a > 0:
                out += [(q, min(255, a)) for q in range(base + x0, base + x1)]
    return out


def paint_fn(cv: Canvas, rows, color_at, alpha: float, clip: Clip, blend: Blend | None = None) -> None:
    """画素ごとに色の変わる塗り（網かけ）：color_at(x, y) → (r, g, b) か None（塗らない）。blend：重ね方。"""
    buf = cv.buf
    W = cv.w
    mask = clip.mask
    cover = cv.cover
    for Y, runs in rows:
        if Y < clip.y0 or Y >= clip.y1:
            continue
        if mask is not None:
            mr = mask.get(Y)
            if not mr:
                continue
            runs = intersect(runs, mr)
        for x0, x1, a in runs:
            x0, x1 = max(x0, clip.x0), min(x1, clip.x1)
            a = int(a * alpha + 0.5)
            if a <= 0:
                continue
            na = 255 - a
            for x in range(x0, x1):
                c = color_at(x + 0.5, Y + 0.5)
                if c is None:
                    continue
                q = Y * W + x
                p = q * 3
                if blend is not None:
                    c = blend.onto(buf[p], buf[p + 1], buf[p + 2], 255 if cover is None else cover[q], c[0], c[1], c[2])
                buf[p] = (buf[p] * na + c[0] * a + 127) // 255
                buf[p + 1] = (buf[p + 1] * na + c[1] * a + 127) // 255
                buf[p + 2] = (buf[p + 2] * na + c[2] * a + 127) // 255
                if cover is not None and cover[q] < 255:
                    cover[q] += ((255 - cover[q]) * a + 127) // 255
