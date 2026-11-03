"""書体のプログラムから、字の形（輪郭）を読む。

- TrueType（glyf・loca：2 次のベジェ、組み合わせの字。OpenType の仕様）
- CFF（Adobe TN 5176）と Type 2 の字の命令（TN 5177）：名前で引く書体と CID で引く書体（FDArray・FDSelect）、
  サブルーチン・ヒントの印（hintmask）・flex・endchar のアクセントの組み合わせ（seac）
- Type 1（Adobe Type 1 Font Format）：eexec の暗号を解き、Subrs・CharStrings の字の命令（hsbw・flex・seac など）

輪郭は canvas の図形の命令（0 移る・1 直線・2 3 次・3 2 次・4 閉じる）で、書体の単位で返す。matrix は書体の単位 → 1 em。
壊れた字は空の輪郭にする（止まらない）。
"""

from __future__ import annotations

import re
import struct

from ..sfnt import CFF_STD, _cff_dict, _cff_index
from ..fonts import STANDARD_ENCODING


class FontProgramError(ValueError):
    pass


# ====================================================================== TrueType
class TrueTypeOutlines:
    kind = "TrueType"

    def __init__(self, tabs: dict[str, bytes]):
        for need in ("head", "maxp", "loca", "glyf"):
            if need not in tabs:
                raise FontProgramError(f"TrueType の {need} の表がありません")
        head = tabs["head"]
        self.upem = struct.unpack_from(">H", head, 18)[0] or 1000
        long_ = struct.unpack_from(">h", head, 50)[0] == 1
        self.num = struct.unpack_from(">H", tabs["maxp"], 4)[0]
        loca = tabs["loca"]
        n = min(self.num + 1, len(loca) // (4 if long_ else 2))
        if long_:
            self.loca = list(struct.unpack_from(f">{n}I", loca, 0))
        else:
            self.loca = [x * 2 for x in struct.unpack_from(f">{n}H", loca, 0)]
        self.glyf = tabs["glyf"]
        self.matrix = (1.0 / self.upem, 0.0, 0.0, 1.0 / self.upem, 0.0, 0.0)
        self.tabs = tabs
        self._cache: dict[int, list] = {}
        self.cmaps = _cmaps(tabs.get("cmap", b""))
        hhea = tabs.get("hhea")
        hm = tabs.get("hmtx", b"")
        self.adv: list[int] = []
        if hhea and len(hhea) >= 36:
            nh = struct.unpack_from(">H", hhea, 34)[0]
            nh = min(nh, len(hm) // 4)
            self.adv = [struct.unpack_from(">H", hm, 4 * i)[0] for i in range(nh)]

    def advance(self, gid: int) -> float | None:
        """送りの幅（em）。"""
        if not self.adv:
            return None
        return self.adv[min(gid, len(self.adv) - 1)] / self.upem

    def glyph(self, gid) -> list:
        if not isinstance(gid, int) or gid < 0 or gid + 1 >= len(self.loca):
            return []
        hit = self._cache.get(gid)
        if hit is None:
            try:
                hit = self._glyph(gid, 0)
            except (struct.error, IndexError, ValueError):
                hit = []
            if len(self._cache) < 8000:
                self._cache[gid] = hit
        return hit

    def _glyph(self, gid: int, depth: int) -> list:
        if depth > 8 or gid + 1 >= len(self.loca):
            return []
        off, end = self.loca[gid], self.loca[gid + 1]
        if end <= off:
            return []
        g = self.glyf[off:end]
        nc = struct.unpack_from(">h", g, 0)[0]
        if nc >= 0:
            return _tt_simple(g, nc)
        ops: list = []
        p = 10
        while True:
            flags, sub = struct.unpack_from(">HH", g, p)
            p += 4
            if flags & 1:
                a1, a2 = struct.unpack_from(">hh", g, p)
                p += 4
            else:
                a1, a2 = struct.unpack_from(">bb", g, p)
                p += 2
            xx, yx, xy, yy = 1.0, 0.0, 0.0, 1.0
            if flags & 8:
                xx = yy = struct.unpack_from(">h", g, p)[0] / 16384
                p += 2
            elif flags & 0x40:
                xx, yy = (v / 16384 for v in struct.unpack_from(">hh", g, p))
                p += 4
            elif flags & 0x80:
                xx, yx, xy, yy = (v / 16384 for v in struct.unpack_from(">hhhh", g, p))
                p += 8
            dx, dy = (a1, a2) if flags & 2 else (0, 0)  # 点で合わせる組み合わせは、ずらさない
            for op in self._glyph(sub, depth + 1):
                if op[0] == 4:
                    ops.append(op)
                    continue
                q = [op[0]]
                for i in range(1, len(op), 2):
                    x, y = op[i], op[i + 1]
                    q += [xx * x + xy * y + dx, yx * x + yy * y + dy]
                ops.append(tuple(q))
            if not flags & 0x20:
                break
        return ops


def _tt_simple(g: bytes, nc: int) -> list:
    if nc == 0:
        return []
    ends = struct.unpack_from(f">{nc}H", g, 10)
    il = struct.unpack_from(">H", g, 10 + 2 * nc)[0]
    p = 12 + 2 * nc + il
    npts = ends[-1] + 1
    flags = []
    while len(flags) < npts:
        f = g[p]
        p += 1
        flags.append(f)
        if f & 8:
            r = g[p]
            p += 1
            flags += [f] * r
    flags = flags[:npts]
    xs = []
    v = 0
    for f in flags:
        if f & 2:
            d = g[p]
            p += 1
            v += d if f & 16 else -d
        elif not f & 16:
            v += struct.unpack_from(">h", g, p)[0]
            p += 2
        xs.append(v)
    ys = []
    v = 0
    for f in flags:
        if f & 4:
            d = g[p]
            p += 1
            v += d if f & 32 else -d
        elif not f & 32:
            v += struct.unpack_from(">h", g, p)[0]
            p += 2
        ys.append(v)
    ops: list = []
    s = 0
    for e in ends:
        pts = [(xs[i], ys[i], flags[i] & 1) for i in range(s, e + 1)]
        s = e + 1
        n = len(pts)
        if n < 2:
            continue
        ex = []
        for i in range(n):
            a, b = pts[i], pts[(i + 1) % n]
            ex.append(a)
            if not a[2] and not b[2]:
                ex.append(((a[0] + b[0]) / 2, (a[1] + b[1]) / 2, 1))
        k = next((i for i, q in enumerate(ex) if q[2]), None)
        if k is None:
            continue
        ex = ex[k:] + ex[:k]
        m = len(ex)
        ops.append((0, ex[0][0], ex[0][1]))
        i = 1
        while i < m:
            q = ex[i]
            if q[2]:
                ops.append((1, q[0], q[1]))
                i += 1
            else:
                r = ex[(i + 1) % m]
                ops.append((3, q[0], q[1], r[0], r[1]))
                i += 2
        ops.append((4,))
    return ops


def _cmaps(t: bytes) -> dict[tuple[int, int], dict[int, int]]:
    from ..sfnt import _subtable

    out: dict[tuple[int, int], dict[int, int]] = {}
    if len(t) < 4:
        return out
    try:
        n = struct.unpack_from(">H", t, 2)[0]
        for i in range(n):
            pid, eid, off = struct.unpack_from(">HHI", t, 4 + 8 * i)
            try:
                m = _subtable(t, off)
            except (struct.error, IndexError):
                m = {}
            if m:
                out.setdefault((pid, eid), m)
    except struct.error:
        pass
    return out


def sfnt_tables(data: bytes, offset: int = 0) -> dict[str, bytes]:
    out: dict[str, bytes] = {}
    n = struct.unpack_from(">H", data, offset + 4)[0]
    for i in range(n):
        tag, _chk, off, ln = struct.unpack_from(">4sIII", data, offset + 12 + 16 * i)
        out[tag.decode("latin-1")] = data[off:off + ln]
    return out


# ====================================================================== CFF
def _bias(n: int) -> int:
    return 107 if n < 1240 else 1131 if n < 33900 else 32768


class CFFOutlines:
    kind = "CFF"

    def __init__(self, data: bytes):
        if len(data) < 4:
            raise FontProgramError("CFF が短すぎます")
        p = data[2]
        _names, p = _cff_index(data, p)
        tops, p = _cff_index(data, p)
        self.strings, p = _cff_index(data, p)
        self.gsubrs, p = _cff_index(data, p)
        if not tops:
            raise FontProgramError("CFF の Top DICT がありません")
        top = _cff_dict(tops[0])
        self.data = data
        fm = top.get(1207)
        self.matrix = tuple(float(v) for v in fm) if fm and len(fm) == 6 else (0.001, 0.0, 0.0, 0.001, 0.0, 0.0)
        cs_off = top.get(17, [0])[0]
        self.charstrings, _ = _cff_index(data, cs_off) if cs_off else ([], 0)
        self.num = len(self.charstrings)
        self.is_cid = 1230 in top
        # 字ごとの私的な辞書（Subrs・幅の既定）
        self.fd_of: bytes | list | None = None
        self.privs: list[tuple[list, float, float]] = []
        self.fd_matrix: list = []
        if self.is_cid and 1236 in top:
            fds, _ = _cff_index(data, top[1236][0])
            for fd in fds:
                d = _cff_dict(fd)
                self.privs.append(self._private(d.get(18)))
                m = d.get(1207)
                self.fd_matrix.append(tuple(float(v) for v in m) if m and len(m) == 6 else None)
            self.fd_of = self._fdselect(top.get(1237, [0])[0])
        else:
            self.privs.append(self._private(top.get(18)))
        # 字の集まり（charset）
        sids = self._charset(top.get(15, [0])[0])
        self.names: dict[str, int] = {}
        self.cid_gid: dict[int, int] = {}
        if self.is_cid:
            self.cid_gid = {sid: g for g, sid in enumerate(sids)}
        else:
            for g, sid in enumerate(sids):
                nm = self._sid(sid)
                if nm and nm not in self.names:
                    self.names[nm] = g
        self.gid_name = {g: n for n, g in self.names.items()}
        self.encoding = self._encoding(top.get(16, [0])[0], sids)
        self._cache: dict[int, list] = {}

    def _sid(self, sid: int) -> str:
        if sid < len(CFF_STD):
            return CFF_STD[sid]
        k = sid - len(CFF_STD)
        return self.strings[k].decode("latin-1") if 0 <= k < len(self.strings) else ""

    def _private(self, ent):
        if not ent or len(ent) < 2:
            return ([], 0.0, 0.0)
        size, off = int(ent[0]), int(ent[1])
        d = _cff_dict(self.data[off:off + size])
        subrs = []
        if 19 in d:
            try:
                subrs, _ = _cff_index(self.data, off + int(d[19][0]))
            except (struct.error, IndexError):
                subrs = []
        return (subrs, float(d.get(20, [0])[0]), float(d.get(21, [0])[0]))

    def _fdselect(self, off: int):
        d = self.data
        if not off:
            return None
        fmt = d[off]
        if fmt == 0:
            return d[off + 1:off + 1 + self.num]
        if fmt == 3:
            nr = struct.unpack_from(">H", d, off + 1)[0]
            out = bytearray(self.num)
            q = off + 3
            for i in range(nr):
                first, fd = struct.unpack_from(">HB", d, q)
                nxt = struct.unpack_from(">H", d, q + 3)[0]
                for g in range(first, min(nxt, self.num)):
                    out[g] = fd
                q += 3
            return bytes(out)
        return None

    def _charset(self, off: int) -> list[int]:
        n = self.num
        d = self.data
        if off == 0:
            return list(range(n))  # ISOAdobe：番号 = SID
        if off in (1, 2):
            return list(range(n))
        sids = [0]
        fmt = d[off]
        q = off + 1
        try:
            if fmt == 0:
                for _ in range(n - 1):
                    sids.append(struct.unpack_from(">H", d, q)[0])
                    q += 2
            else:
                while len(sids) < n:
                    first = struct.unpack_from(">H", d, q)[0]
                    if fmt == 1:
                        left = d[q + 2]
                        q += 3
                    else:
                        left = struct.unpack_from(">H", d, q + 2)[0]
                        q += 4
                    sids += range(first, first + left + 1)
        except (struct.error, IndexError):
            pass
        return sids[:n]

    def _encoding(self, off: int, sids: list[int]) -> dict[int, int]:
        enc: dict[int, int] = {}
        if self.is_cid:
            return enc
        d = self.data
        if off > 1:
            try:
                fmt = d[off] & 0x7F
                q = off + 1
                if fmt == 0:
                    cnt = d[q]
                    for g in range(1, cnt + 1):
                        enc[d[q + g]] = g
                    q += 1 + cnt
                elif fmt == 1:
                    nr = d[q]
                    g = 1
                    q += 1
                    for _ in range(nr):
                        first, left = d[q], d[q + 1]
                        for k in range(left + 1):
                            enc[first + k] = g
                            g += 1
                        q += 2
                if d[off] & 0x80:  # 補い（同じ字に別の符号）
                    ns = d[q]
                    q += 1
                    for _ in range(ns):
                        code, sid = d[q], struct.unpack_from(">H", d, q + 1)[0]
                        nm = self._sid(sid)
                        if nm in self.names:
                            enc[code] = self.names[nm]
                        q += 3
            except (struct.error, IndexError):
                pass
            return enc
        # 0：標準の符号、1：Expert（名前で引く）
        for code, nm in enumerate(STANDARD_ENCODING):
            if nm and nm in self.names:
                enc[code] = self.names[nm]
        return enc

    def gid_for_cid(self, cid: int) -> int:
        if self.is_cid:
            return self.cid_gid.get(cid, 0)
        return cid

    def glyph_matrix(self, gid: int):
        if self.fd_of is not None and self.fd_matrix:
            k = self.fd_of[gid] if gid < len(self.fd_of) else 0
            m = self.fd_matrix[k] if k < len(self.fd_matrix) else None
            if m is not None:
                # FDArray の FontMatrix は、Top の FontMatrix に重ねる（TN 5176 18）
                from .canvas import mul

                top = self.matrix if self.matrix != (0.001, 0.0, 0.0, 0.001, 0.0, 0.0) else (1.0, 0, 0, 1.0, 0, 0)
                return mul(m, top)
        return self.matrix

    def glyph(self, gid) -> list:
        if isinstance(gid, str):
            gid = self.names.get(gid)
            if gid is None:
                return []
        if not isinstance(gid, int) or not 0 <= gid < self.num:
            return []
        hit = self._cache.get(gid)
        if hit is None:
            k = 0
            if self.fd_of is not None:
                k = self.fd_of[gid] if gid < len(self.fd_of) else 0
            subrs, dflt, nom = self.privs[k] if k < len(self.privs) else ([], 0.0, 0.0)

            def seac(code):
                nm = STANDARD_ENCODING[code] if 0 <= code < 256 else None
                g = self.names.get(nm) if nm else None
                return self.glyph(g) if g is not None and g != gid else []

            try:
                hit = type2(self.charstrings[gid], self.gsubrs, subrs, seac)
            except (IndexError, ValueError, struct.error, ZeroDivisionError, RecursionError):
                hit = []
            if len(self._cache) < 8000:
                self._cache[gid] = hit
        return hit


def type2(code: bytes, gsubrs, lsubrs, seac=None) -> list:
    """Type 2 の字の命令 → 輪郭。"""
    ops: list = []
    st: list = []
    state = {"x": 0.0, "y": 0.0, "stems": 0, "open": False, "first": True, "done": False}
    gb, lb = _bias(len(gsubrs)), _bias(len(lsubrs))
    trans: list = [0.0] * 32

    def move(dx, dy):
        if state["open"]:
            ops.append((4,))
        state["x"] += dx
        state["y"] += dy
        ops.append((0, state["x"], state["y"]))
        state["open"] = True

    def line(dx, dy):
        state["x"] += dx
        state["y"] += dy
        ops.append((1, state["x"], state["y"]))

    def curve(a, b, c, d, e, f):
        x0, y0 = state["x"], state["y"]
        x1, y1 = x0 + a, y0 + b
        x2, y2 = x1 + c, y1 + d
        x3, y3 = x2 + e, y2 + f
        ops.append((2, x1, y1, x2, y2, x3, y3))
        state["x"], state["y"] = x3, y3

    def width(n_expected_even: bool, n: int):
        # 最初の、積み上げを消す命令の前の余分な数は幅（使わない）
        if state["first"]:
            state["first"] = False
            if (n_expected_even and len(st) % 2) or (not n_expected_even and len(st) > n):
                del st[0]

    def run(c: bytes, depth: int):
        if depth > 12:
            raise ValueError("サブルーチンが深すぎます")
        i = 0
        L = len(c)
        while i < L:
            b0 = c[i]
            if b0 >= 32:
                if b0 <= 246:
                    st.append(b0 - 139)
                    i += 1
                elif b0 <= 250:
                    st.append((b0 - 247) * 256 + c[i + 1] + 108)
                    i += 2
                elif b0 <= 254:
                    st.append(-(b0 - 251) * 256 - c[i + 1] - 108)
                    i += 2
                else:
                    st.append(struct.unpack_from(">i", c, i + 1)[0] / 65536)
                    i += 5
                continue
            if b0 == 28:
                st.append(struct.unpack_from(">h", c, i + 1)[0])
                i += 3
                continue
            i += 1
            if b0 in (1, 3, 18, 23):  # hstem vstem hstemhm vstemhm
                width(True, 0)
                state["stems"] += len(st) // 2
                st.clear()
            elif b0 in (19, 20):  # hintmask cntrmask
                width(True, 0)
                state["stems"] += len(st) // 2
                st.clear()
                i += (state["stems"] + 7) // 8
            elif b0 == 21:
                width(False, 2)
                move(st[-2], st[-1])
                st.clear()
            elif b0 == 22:
                width(False, 1)
                move(st[-1], 0)
                st.clear()
            elif b0 == 4:
                width(False, 1)
                move(0, st[-1])
                st.clear()
            elif b0 == 5:
                for k in range(0, len(st) - 1, 2):
                    line(st[k], st[k + 1])
                st.clear()
            elif b0 in (6, 7):
                horiz = b0 == 6
                for v in st:
                    line(v, 0) if horiz else line(0, v)
                    horiz = not horiz
                st.clear()
            elif b0 == 8:
                for k in range(0, len(st) - 5, 6):
                    curve(*st[k:k + 6])
                st.clear()
            elif b0 == 24:
                k = 0
                while len(st) - k >= 8:
                    curve(*st[k:k + 6])
                    k += 6
                if len(st) - k >= 2:
                    line(st[k], st[k + 1])
                st.clear()
            elif b0 == 25:
                k = 0
                while len(st) - k >= 8:
                    line(st[k], st[k + 1])
                    k += 2
                if len(st) - k >= 6:
                    curve(*st[k:k + 6])
                st.clear()
            elif b0 == 26:
                k = 0
                dx1 = 0
                if len(st) % 2:
                    dx1 = st[0]
                    k = 1
                while len(st) - k >= 4:
                    curve(dx1, st[k], st[k + 1], st[k + 2], 0, st[k + 3])
                    dx1 = 0
                    k += 4
                st.clear()
            elif b0 == 27:
                k = 0
                dy1 = 0
                if len(st) % 2:
                    dy1 = st[0]
                    k = 1
                while len(st) - k >= 4:
                    curve(st[k], dy1, st[k + 1], st[k + 2], st[k + 3], 0)
                    dy1 = 0
                    k += 4
                st.clear()
            elif b0 in (30, 31):
                horiz = b0 == 31
                k = 0
                n = len(st)
                while n - k >= 4:
                    last = n - k == 5
                    a = st[k:k + 4]
                    e = st[k + 4] if last else 0
                    if horiz:
                        curve(a[0], 0, a[1], a[2], e, a[3])
                    else:
                        curve(0, a[0], a[1], a[2], a[3], e)
                    k += 5 if last else 4
                    horiz = not horiz
                st.clear()
            elif b0 in (10, 29):
                if not st:
                    raise ValueError("サブルーチンの番号がありません")
                idx = int(st.pop())
                subs, bias = (lsubrs, lb) if b0 == 10 else (gsubrs, gb)
                j = idx + bias
                if not 0 <= j < len(subs):
                    raise ValueError("サブルーチンの番号が範囲の外です")
                run(subs[j], depth + 1)
                if state["done"]:
                    return
            elif b0 == 11:
                return
            elif b0 == 14:
                if state["first"]:
                    state["first"] = False
                    if len(st) in (1, 5):
                        del st[0]
                if len(st) >= 4 and seac is not None:
                    adx, ady, bchar, achar = st[-4:]
                    if state["open"]:
                        ops.append((4,))
                        state["open"] = False
                    ops.extend(seac(int(bchar)))
                    for op in seac(int(achar)):
                        if op[0] == 4:
                            ops.append(op)
                        else:
                            ops.append(tuple([op[0]] + [v + (adx if j % 2 == 0 else ady) for j, v in enumerate(op[1:])]))
                elif state["open"]:
                    ops.append((4,))
                    state["open"] = False
                state["done"] = True
                return
            elif b0 == 12:
                b1 = c[i]
                i += 1
                if b1 == 35:
                    a = st[:13]
                    curve(*a[0:6])
                    curve(*a[6:12])
                elif b1 == 34:
                    dx1, dx2, dy2, dx3, dx4, dx5, dx6 = st[:7]
                    curve(dx1, 0, dx2, dy2, dx3, 0)
                    curve(dx4, 0, dx5, -dy2, dx6, 0)
                elif b1 == 36:
                    dx1, dy1, dx2, dy2, dx3, dx4, dx5, dy5, dx6 = st[:9]
                    curve(dx1, dy1, dx2, dy2, dx3, 0)
                    curve(dx4, 0, dx5, dy5, dx6, -(dy1 + dy2 + dy5))
                elif b1 == 37:
                    a = st[:11]
                    dx = sum(a[0:10:2])
                    dy = sum(a[1:10:2])
                    if abs(dx) > abs(dy):
                        d6x, d6y = a[10], -dy
                    else:
                        d6x, d6y = -dx, a[10]
                    curve(*a[0:6])
                    curve(a[6], a[7], a[8], a[9], d6x, d6y)
                elif b1 in (3, 4, 5, 9, 10, 11, 12, 14, 15, 18, 20, 21, 22, 23, 24, 26, 27, 28, 29, 30):
                    _arith(b1, st, trans)
                    continue
                st.clear()
            else:
                st.clear()

    run(code, 0)
    if state["open"] and not state["done"]:
        ops.append((4,))
    return ops


def _arith(op: int, st: list, trans: list) -> None:
    if op == 3:
        b, a = st.pop(), st.pop()
        st.append(1 if a and b else 0)
    elif op == 4:
        b, a = st.pop(), st.pop()
        st.append(1 if a or b else 0)
    elif op == 5:
        st.append(0 if st.pop() else 1)
    elif op == 9:
        st.append(abs(st.pop()))
    elif op == 10:
        b, a = st.pop(), st.pop()
        st.append(a + b)
    elif op == 11:
        b, a = st.pop(), st.pop()
        st.append(a - b)
    elif op == 12:
        b, a = st.pop(), st.pop()
        st.append(a / b if b else 0)
    elif op == 14:
        st.append(-st.pop())
    elif op == 15:
        b, a = st.pop(), st.pop()
        st.append(1 if a == b else 0)
    elif op == 18:
        st.pop()
    elif op == 20:
        i, v = int(st.pop()), st.pop()
        if 0 <= i < len(trans):
            trans[i] = v
    elif op == 21:
        i = int(st.pop())
        st.append(trans[i] if 0 <= i < len(trans) else 0)
    elif op == 22:
        v2, v1, s1, s2 = st.pop(), st.pop(), st.pop(), st.pop()
        st.append(s1 if v1 <= v2 else s2)
    elif op == 23:
        st.append(0.5)
    elif op == 24:
        b, a = st.pop(), st.pop()
        st.append(a * b)
    elif op == 26:
        st.append(abs(st.pop()) ** 0.5)
    elif op == 27:
        st.append(st[-1])
    elif op == 28:
        st[-1], st[-2] = st[-2], st[-1]
    elif op == 29:
        i = int(st.pop())
        st.append(st[-1 - max(0, i)])
    elif op == 30:
        j, n = int(st.pop()), int(st.pop())
        if n > 0:
            seg = st[-n:]
            j %= n
            st[-n:] = seg[-j:] + seg[:-j] if j else seg


# ====================================================================== Type 1
def _decrypt(data: bytes, r: int, skip: int) -> bytes:
    out = bytearray(len(data))
    for i, c in enumerate(data):
        out[i] = c ^ (r >> 8)
        r = ((c + r) * 52845 + 22719) & 0xFFFF
    return bytes(out[skip:])


class Type1Outlines:
    kind = "Type1"

    def __init__(self, data: bytes, length1: int | None = None):
        if data[:2] == b"\x80\x01":  # PFB の区切り
            data = _pfb(data)
        k = data.find(b"eexec")
        if k < 0:
            raise FontProgramError("Type 1 の eexec がありません")
        clear = data[:k]
        rest = data[k + 5:]
        while rest[:1] in (b"\r", b"\n", b" ", b"\t"):
            rest = rest[1:]
        if len(rest) >= 4 and all(chr(ch) in "0123456789abcdefABCDEF \r\n\t" for ch in rest[:16]):
            hexs = re.sub(rb"[^0-9A-Fa-f]", b"", rest[:])
            hexs = hexs[:len(hexs) // 2 * 2]
            rest = bytes.fromhex(hexs.decode("ascii"))
        priv = _decrypt(rest, 55665, 4)
        m = re.search(rb"/FontMatrix\s*\[([^\]]*)\]", clear)
        self.matrix = (0.001, 0.0, 0.0, 0.001, 0.0, 0.0)
        if m:
            try:
                v = [float(x) for x in m[1].split()]
                if len(v) == 6:
                    self.matrix = tuple(v)
            except ValueError:
                pass
        from ..sfnt import type1_encoding

        self.encoding_names = type1_encoding(data[:k + 5])
        m = re.search(rb"/lenIV\s+(-?\d+)", priv)
        self.leniv = int(m[1]) if m else 4
        self.subrs: list[bytes] = []
        self.charstrings: dict[str, bytes] = {}
        self._parse(priv)
        self.names = {n: n for n in self.charstrings}
        self._cache: dict[str, list] = {}

    def _parse(self, p: bytes) -> None:
        m = re.search(rb"/Subrs\s+(\d+)\s+array", p)
        pos = 0
        if m:
            n = int(m[1])
            self.subrs = [b""] * n
            pos = m.end()
            rx = re.compile(rb"dup\s+(\d+)\s+(\d+)\s+(\S+)\s")
            for _ in range(n):
                mm = rx.search(p, pos)
                if not mm:
                    break
                i, ln = int(mm[1]), int(mm[2])
                start = mm.end()
                if 0 <= i < n:
                    self.subrs[i] = self._cs(p[start:start + ln])
                pos = start + ln
        m = re.search(rb"/CharStrings\s+\d+\s+dict", p[pos:]) or re.search(rb"/CharStrings\s+\d+\s+dict", p)
        if not m:
            return
        base = pos if re.search(rb"/CharStrings\s+\d+\s+dict", p[pos:]) else 0
        q = base + m.end()
        rx = re.compile(rb"/([^\s/\[\]{}()<>%]+)\s+(\d+)\s+(\S+)\s")
        while True:
            mm = rx.search(p, q)
            if not mm or mm[3] in (b"def", b"end"):
                if not mm:
                    break
                q = mm.end()
                continue
            ln = int(mm[2])
            start = mm.end()
            self.charstrings[mm[1].decode("latin-1")] = self._cs(p[start:start + ln])
            q = start + ln
            if len(self.charstrings) > 70000:
                break

    def _cs(self, b: bytes) -> bytes:
        return _decrypt(b, 4330, self.leniv) if self.leniv >= 0 else b

    def glyph(self, name) -> list:
        if not isinstance(name, str):
            return []
        hit = self._cache.get(name)
        if hit is None:
            cs = self.charstrings.get(name)
            try:
                hit = self._run(cs, 0) if cs is not None else []
            except (IndexError, ValueError, ZeroDivisionError, RecursionError):
                hit = []
            if len(self._cache) < 8000:
                self._cache[name] = hit
        return hit

    def _run(self, code: bytes, depth: int) -> list:
        ops: list = []
        st: list = []
        ps: list = []
        s = {"x": 0.0, "y": 0.0, "open": False, "flex": None, "sbx": 0.0, "sby": 0.0}

        def move(x, y):
            if s["flex"] is not None:
                s["flex"].append((x, y))
                s["x"], s["y"] = x, y
                return
            if s["open"]:
                ops.append((4,))
            ops.append((0, x, y))
            s["x"], s["y"] = x, y
            s["open"] = True

        def line(x, y):
            if not s["open"]:
                ops.append((0, s["x"], s["y"]))
                s["open"] = True
            ops.append((1, x, y))
            s["x"], s["y"] = x, y

        def curve(x1, y1, x2, y2, x3, y3):
            if not s["open"]:
                ops.append((0, s["x"], s["y"]))
                s["open"] = True
            ops.append((2, x1, y1, x2, y2, x3, y3))
            s["x"], s["y"] = x3, y3

        def rcurve(a, b, c, d, e, f):
            x0, y0 = s["x"], s["y"]
            curve(x0 + a, y0 + b, x0 + a + c, y0 + b + d, x0 + a + c + e, y0 + b + d + f)

        def run(c: bytes, depth: int) -> bool:
            if depth > 12:
                raise ValueError("Subrs が深すぎます")
            i = 0
            L = len(c)
            while i < L:
                v = c[i]
                if v >= 32:
                    if v <= 246:
                        st.append(v - 139)
                        i += 1
                    elif v <= 250:
                        st.append((v - 247) * 256 + c[i + 1] + 108)
                        i += 2
                    elif v <= 254:
                        st.append(-(v - 251) * 256 - c[i + 1] - 108)
                        i += 2
                    else:
                        st.append(struct.unpack_from(">i", c, i + 1)[0])
                        i += 5
                    continue
                i += 1
                if v == 13:  # hsbw
                    s["sbx"] = st[0]
                    s["x"], s["y"] = st[0], 0.0
                    st.clear()
                elif v == 21:
                    move(s["x"] + st[0], s["y"] + st[1])
                    st.clear()
                elif v == 22:
                    move(s["x"] + st[0], s["y"])
                    st.clear()
                elif v == 4:
                    move(s["x"], s["y"] + st[0])
                    st.clear()
                elif v == 5:
                    line(s["x"] + st[0], s["y"] + st[1])
                    st.clear()
                elif v == 6:
                    line(s["x"] + st[0], s["y"])
                    st.clear()
                elif v == 7:
                    line(s["x"], s["y"] + st[0])
                    st.clear()
                elif v == 8:
                    rcurve(*st[:6])
                    st.clear()
                elif v == 30:
                    rcurve(0, st[0], st[1], st[2], st[3], 0)
                    st.clear()
                elif v == 31:
                    rcurve(st[0], 0, st[1], st[2], 0, st[3])
                    st.clear()
                elif v == 9:
                    if s["open"]:
                        ops.append((4,))
                        s["open"] = False
                    st.clear()
                elif v == 10:
                    k = int(st.pop())
                    if 0 <= k < len(self.subrs):
                        if run(self.subrs[k], depth + 1):
                            return True
                elif v == 11:
                    return False
                elif v == 14:
                    if s["open"]:
                        ops.append((4,))
                        s["open"] = False
                    return True
                elif v in (1, 3):
                    st.clear()
                elif v == 12:
                    w = c[i]
                    i += 1
                    if w == 7:  # sbw
                        s["sbx"], s["sby"] = st[0], st[1]
                        s["x"], s["y"] = st[0], st[1]
                        st.clear()
                    elif w == 6:  # seac
                        asb, adx, ady, bchar, achar = st[:5]
                        st.clear()
                        bn = STANDARD_ENCODING[int(bchar)] if 0 <= int(bchar) < 256 else None
                        an = STANDARD_ENCODING[int(achar)] if 0 <= int(achar) < 256 else None
                        if s["open"]:
                            ops.append((4,))
                            s["open"] = False
                        if bn:
                            ops.extend(self.glyph(bn))
                        if an:
                            dx, dy = adx - asb + s["sbx"], ady
                            for op in self.glyph(an):
                                if op[0] == 4:
                                    ops.append(op)
                                else:
                                    ops.append(tuple([op[0]] + [val + (dx if j % 2 == 0 else dy) for j, val in enumerate(op[1:])]))
                        return True
                    elif w == 12:  # div
                        b, a = st.pop(), st.pop()
                        st.append(a / b if b else 0)
                    elif w == 16:  # callothersubr
                        oth = int(st.pop())
                        n = int(st.pop())
                        args = [st.pop() for _ in range(n)][::-1] if n <= len(st) else []
                        if oth == 1:
                            s["flex"] = []
                        elif oth == 0 and s["flex"] is not None:
                            pts = s["flex"]
                            s["flex"] = None
                            if len(pts) >= 7:
                                curve(*pts[1], *pts[2], *pts[3])
                                curve(*pts[4], *pts[5], *pts[6])
                            elif pts:
                                line(*pts[-1])
                            ps.extend([s["y"], s["x"]])
                            continue
                        elif oth == 2:
                            pass
                        ps.extend(args[::-1] if oth != 3 else [3])
                    elif w == 17:  # pop
                        st.append(ps.pop() if ps else 0)
                    elif w == 33:  # setcurrentpoint
                        s["x"], s["y"] = st[0], st[1]
                        st.clear()
                    else:
                        st.clear()
                else:
                    st.clear()
            return False

        run(code, depth)
        if s["open"]:
            ops.append((4,))
        return ops


def _pfb(data: bytes) -> bytes:
    out = bytearray()
    i = 0
    while i + 6 <= len(data) and data[i] == 0x80 and data[i + 1] in (1, 2):
        ln = struct.unpack_from("<I", data, i + 2)[0]
        out += data[i + 6:i + 6 + ln]
        i += 6 + ln
    return bytes(out)
