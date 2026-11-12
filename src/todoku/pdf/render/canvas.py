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


def fill_rows(polys, evenodd, clip_rect, separate_rects: bool = False) -> list:
    """塗り：多角形 → 行ごとの続き。separate_rects：四角だけの並びなら、1 つずつ速く数える（重なりの縁は二重になる）。"""
    if separate_rects and len(polys) > 1:
        rs = [_rect_of(p) for p in polys]
        if all(r is not None for r in rs):
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
    def __init__(self, w: int, h: int):
        self.w, self.h = w, h
        self.buf = bytearray(b"\xff" * (w * h * 3))


def paint(cv: Canvas, rows, rgb, alpha: float, clip: Clip, dx: int = 0, dy: int = 0) -> None:
    """続きを、色 rgb（0〜255 の 3 つ）・不透明度 alpha で、切り抜きの中に塗る。dx, dy：続きをずらす量（字の型）。"""
    buf = cv.buf
    W = cv.w
    cx0, cy0, cx1, cy1 = clip.x0, clip.y0, clip.x1, clip.y1
    mask = clip.mask
    r, g, b = rgb
    full = bytes((r, g, b))
    op = alpha >= 0.999
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
            elif a > 0:
                na = 255 - a
                ra, ga, ba = r * a + 127, g * a + 127, b * a + 127
                p = (base + x0) * 3
                for p in range(p, p + 3 * (x1 - x0), 3):
                    buf[p] = (buf[p] * na + ra) // 255
                    buf[p + 1] = (buf[p + 1] * na + ga) // 255
                    buf[p + 2] = (buf[p + 2] * na + ba) // 255


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


def paint_fn(cv: Canvas, rows, color_at, alpha: float, clip: Clip) -> None:
    """画素ごとに色の変わる塗り（網かけ）：color_at(x, y) → (r, g, b) か None（塗らない）。"""
    buf = cv.buf
    W = cv.w
    mask = clip.mask
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
                p = (Y * W + x) * 3
                buf[p] = (buf[p] * na + c[0] * a + 127) // 255
                buf[p + 1] = (buf[p + 1] * na + c[1] * a + 127) // 255
                buf[p + 2] = (buf[p + 2] * na + c[2] * a + 127) // 255
