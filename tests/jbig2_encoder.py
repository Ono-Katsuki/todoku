"""試験のための JBIG2 の符号化（ITU-T T.88 の符号化の手順どおり）。復号器とは別に、仕様から書いたもの。
pdftoppm（poppler の別の実装）で同じ画素に開けることを確かめて、符号化と復号の両方を確かめる。"""

from __future__ import annotations

import struct

from todoku.imaging.mq import _T  # 確率の表（T.88 表 E.1）
from todoku.pdf.raster import _BM, _BT, _WM, _WT, _XM


class MQEnc:
    def __init__(self):
        self.out = bytearray([0])  # 先頭の 1 バイトは BP = 開始の前（捨てる）
        self.a, self.c, self.ct = 0x8000, 0, 12

    def _byteout(self):
        o = self.out
        if o[-1] == 0xFF:
            o.append((self.c >> 20) & 0xFF)
            self.c &= 0xFFFFF
            self.ct = 7
        elif self.c < 0x8000000:
            o.append((self.c >> 19) & 0xFF)
            self.c &= 0x7FFFF
            self.ct = 8
        else:
            o[-1] += 1
            if o[-1] == 0xFF:
                self.c &= 0x7FFFFFF
                o.append((self.c >> 20) & 0xFF)
                self.c &= 0xFFFFF
                self.ct = 7
            else:
                o.append((self.c >> 19) & 0xFF)
                self.c &= 0x7FFFF
                self.ct = 8

    def _renorm(self):
        while True:
            self.a <<= 1
            self.c <<= 1
            self.ct -= 1
            if self.ct == 0:
                self._byteout()
            if self.a & 0x8000:
                break

    def encode(self, cx: dict, i: int, d: int):
        idx, mps = cx.get(i, (0, 0))
        q, nm, nl, sw = _T[idx]
        if d == mps:
            self.a -= q
            if not self.a & 0x8000:
                if self.a < q:
                    self.a = q
                else:
                    self.c += q
                cx[i] = (nm, mps)
                self._renorm()
            else:
                self.c += q
        else:
            self.a -= q
            if self.a < q:
                self.c += q
            else:
                self.a = q
            cx[i] = (nl, mps ^ sw)
            self._renorm()

    def flush(self, jbig2_marker: bool = True) -> bytes:
        t = self.c + self.a
        self.c |= 0xFFFF
        if self.c >= t:
            self.c -= 0x8000
        self.c <<= self.ct
        self._byteout()
        self.c <<= self.ct
        self._byteout()
        o = self.out
        if o[-1] != 0xFF:
            o.append(0xFF)
        if jbig2_marker:
            o.append(0xAC)
        return bytes(o[1:])


# ---------------------------------------------------------------------- 整数（附属書 A）
def enc_int(mq: MQEnc, cx: dict, v: int | None):
    prev = 1

    def bit(d):
        nonlocal prev
        mq.encode(cx, prev, d)
        prev = (prev << 1 | d) if prev < 256 else ((prev << 1 | d) & 511) | 256

    if v is None:
        s, a = 1, 0
    else:
        s, a = (1, -v) if v < 0 else (0, v)
    bit(s)
    for i, (n, off, nxt) in enumerate(((2, 0, 4), (4, 4, 20), (6, 20, 84), (8, 84, 340), (12, 340, 4436), (32, 4436, None))):
        if nxt is None or a < nxt:
            if i < 5:
                for _ in range(i):
                    bit(1)
                bit(0)
            else:
                for _ in range(5):
                    bit(1)
            for k in range(n - 1, -1, -1):
                bit(((a - off) >> k) & 1)
            return


def enc_iaid(mq: MQEnc, cx: dict, length: int, v: int):
    prev = 1
    for k in range(length - 1, -1, -1):
        d = (v >> k) & 1
        mq.encode(cx, prev, d)
        prev = prev << 1 | d


# ---------------------------------------------------------------------- 汎用領域（6.2）
NOMINAL = {0: [(3, -1), (-3, -1), (2, -2), (-2, -2)], 1: [(3, -1)], 2: [(2, -1)], 3: [(2, -1)]}
TPGD = {0: 0x9B25, 1: 0x0795, 2: 0x00E5, 3: 0x0195}


def _px(rows, x, y):
    return rows[y][x] if 0 <= y < len(rows) and 0 <= x < len(rows[y]) else 0


def gb_context(rows, x, y, t, at):
    def P(dx, dy):
        return _px(rows, x + dx, y + dy)

    if t == 0:
        A = [P(*a) for a in at]
        return (P(-1, 0) | P(-2, 0) << 1 | P(-3, 0) << 2 | P(-4, 0) << 3 | A[0] << 4 | P(2, -1) << 5 | P(1, -1) << 6 | P(0, -1) << 7
                | P(-1, -1) << 8 | P(-2, -1) << 9 | A[1] << 10 | A[2] << 11 | P(1, -2) << 12 | P(0, -2) << 13 | P(-1, -2) << 14 | A[3] << 15)
    if t == 1:
        return (P(-1, 0) | P(-2, 0) << 1 | P(-3, 0) << 2 | P(*at[0]) << 3 | P(2, -1) << 4 | P(1, -1) << 5 | P(0, -1) << 6 | P(-1, -1) << 7
                | P(-2, -1) << 8 | P(2, -2) << 9 | P(1, -2) << 10 | P(0, -2) << 11 | P(-1, -2) << 12)
    if t == 2:
        return (P(-1, 0) | P(-2, 0) << 1 | P(*at[0]) << 2 | P(1, -1) << 3 | P(0, -1) << 4 | P(-1, -1) << 5 | P(-2, -1) << 6
                | P(1, -2) << 7 | P(0, -2) << 8 | P(-1, -2) << 9)
    return (P(-1, 0) | P(-2, 0) << 1 | P(-3, 0) << 2 | P(-4, 0) << 3 | P(*at[0]) << 4 | P(1, -1) << 5 | P(0, -1) << 6 | P(-1, -1) << 7
            | P(-2, -1) << 8 | P(-3, -1) << 9)


def enc_generic(mq: MQEnc, cx: dict, rows, t=0, tpgd=False, at=None, skip=None):
    """rows：[[0/1,...]]。"""
    at = at or NOMINAL[t]
    done: list = []
    ltp = 0
    for y, row in enumerate(rows):
        if tpgd:
            same = row == (rows[y - 1] if y else [0] * len(row))
            mq.encode(cx, TPGD[t], ltp ^ int(same))
            ltp = int(same)
            done.append(row)
            if same:
                continue
        else:
            done.append(row)
        for x, d in enumerate(row):
            if skip and skip[y][x]:
                continue
            # 今の行は、左の画素だけが見える（AT も）
            vis = done[:-1] + [row[:x]]
            mq.encode(cx, gb_context(vis, x, y, t, at), d)


# ---------------------------------------------------------------------- 細かくする領域（6.3）
def gr_context(cur, ref, x, y, t, dx, dy, at):
    def C(a, b):
        return _px(cur, x + a, y + b)

    def R(a, b):
        return _px(ref, x - dx + a, y - dy + b)

    if t == 0:
        return (C(-1, 0) | C(1, -1) << 1 | C(0, -1) << 2 | C(*at[0]) << 3 | R(1, 1) << 4 | R(0, 1) << 5 | R(-1, 1) << 6 | R(1, 0) << 7
                | R(0, 0) << 8 | R(-1, 0) << 9 | R(1, -1) << 10 | R(0, -1) << 11 | R(*at[1]) << 12)
    return (C(-1, 0) | C(1, -1) << 1 | C(0, -1) << 2 | C(-1, -1) << 3 | R(1, 1) << 4 | R(0, 1) << 5 | R(1, 0) << 6 | R(0, 0) << 7
            | R(-1, 0) << 8 | R(0, -1) << 9)


def enc_refine(mq: MQEnc, cx: dict, rows, ref, t=0, dx=0, dy=0, tpgr=False, at=None):
    at = at or [(-1, -1), (-1, -1)]
    ltp = 0
    tp = 0x0100 if t == 0 else 0x0080
    for y, row in enumerate(rows):
        if tpgr:
            # その行で典型の予測がすべて当たるなら、予測を使う
            ok = True
            for x, d in enumerate(row):
                vals = [_px(ref, x - dx + a, y - dy + b) for a in (-1, 0, 1) for b in (-1, 0, 1)]
                if sum(vals) in (0, 9) and vals[0] != d:
                    ok = False
                    break
            mq.encode(cx, tp, ltp ^ int(ok))
            ltp = int(ok)
        for x, d in enumerate(row):
            if ltp:
                vals = [_px(ref, x - dx + a, y - dy + b) for a in (-1, 0, 1) for b in (-1, 0, 1)]
                if sum(vals) in (0, 9):
                    continue
            cur = rows[:y] + [row[:x]]
            mq.encode(cx, gr_context(cur, ref, x, y, t, dx, dy, at), d)


# ---------------------------------------------------------------------- MMR（T.6）
def _runcode(n: int, black: bool) -> str:
    term = {v: k for k, v in (_BT if black else _WT).items()}
    mk = {v: k for k, v in {**(_BM if black else _WM), **_XM}.items()}
    s = ""
    while n >= 2560:
        s += mk[2560]
        n -= 2560
    if n >= 64:
        s += mk[n // 64 * 64]
        n %= 64
    return s + term[n]


def enc_mmr(rows, eofb=True) -> bytes:
    w = len(rows[0]) if rows else 0
    bits: list[str] = []
    ref = [0] * w
    for row in rows:
        a0, color = -1, 0
        while a0 < w:
            i = max(a0, 0)
            while i < w and row[i] == color:
                i += 1
            a1 = i
            # b1：a0 より右で、色が color と逆へ変わる参照の行の変わり目
            b1 = w
            for i in range(a0 + 1, w):
                prev = ref[i - 1] if i else 0
                if ref[i] != prev and ref[i] != color:
                    b1 = i
                    break
            b2 = w
            for i in range(b1 + 1, w):
                if ref[i] != ref[i - 1]:
                    b2 = i
                    break
            if b2 < a1:
                bits.append("0001")  # 通過
                a0 = b2
                continue
            d = a1 - b1
            if -3 <= d <= 3:
                bits.append({0: "1", 1: "011", 2: "000011", 3: "0000011", -1: "010", -2: "000010", -3: "0000010"}[d])
                a0 = a1
                color ^= 1
            else:
                j = a1
                while j < w and row[j] != color:
                    j += 1
                bits.append("001" + _runcode(a1 - max(a0, 0), bool(color)) + _runcode(j - a1, not color))
                a0 = j
        ref = row
    if eofb:
        bits.append("000000000001" * 2)
    s = "".join(bits)
    s += "0" * (-len(s) % 8)
    return bytes(int(s[i:i + 8], 2) for i in range(0, len(s), 8))


# ---------------------------------------------------------------------- ハフマンを書く
class BitWriter:
    def __init__(self):
        self.s: list[str] = []
        self.n = 0

    def bits(self, v: int, n: int):
        if n:
            self.s.append(format(v & ((1 << n) - 1), f"0{n}b"))
            self.n += n

    def align(self):
        self.bits(0, -self.n % 8)

    def raw(self, b: bytes):
        self.align()
        for x in b:
            self.bits(x, 8)

    def bytes(self) -> bytes:
        self.align()
        s = "".join(self.s)
        return bytes(int(s[i:i + 8], 2) for i in range(0, len(s), 8))


def huff_put(bw: BitWriter, table, v: int | None):
    """todoku の Huffman の表（符号の割り当ては B.3）に値を書く。"""
    for (n, code), (_p, rl, low, kind) in table.codes.items():
        if v is None:
            if kind == "oob":
                bw.bits(code, n)
                return
            continue
        if kind == "oob":
            continue
        if kind == "low":
            if v <= low:
                bw.bits(code, n)
                bw.bits(low - v, 32)
                return
        elif kind == "high":
            if v >= low:
                bw.bits(code, n)
                bw.bits(v - low, 32)
                return
        elif low <= v < low + (1 << rl):
            bw.bits(code, n)
            bw.bits(v - low, rl)
            return
    raise ValueError(f"table has no code for {v}")


# ---------------------------------------------------------------------- 区切り
def segment(number: int, typ: int, data: bytes, refs=(), page: int = 1) -> bytes:
    assert len(refs) <= 4
    out = struct.pack(">IB", number, typ)
    out += bytes([len(refs) << 5])
    fmt = ">B" if number <= 256 else ">H"
    for r in refs:
        out += struct.pack(fmt, r)
    out += bytes([page]) + struct.pack(">I", len(data))
    return out + data


def page_info(w: int, h: int, fill: int = 0) -> bytes:
    return struct.pack(">IIIIBH", w, h, 0, 0, fill << 2, 0)


def region_info(w, h, x=0, y=0, op=0) -> bytes:
    return struct.pack(">IIIIB", w, h, x, y, op)


def at_bytes(at) -> bytes:
    return b"".join(struct.pack(">bb", *a) for a in at)


def generic_seg(rows, x=0, y=0, op=0, t=0, tpgd=False, at=None, mmr=False) -> bytes:
    w, h = len(rows[0]), len(rows)
    d = region_info(w, h, x, y, op) + bytes([(1 if mmr else 0) | t << 1 | (8 if tpgd else 0)])
    if mmr:
        return d + enc_mmr(rows)
    at = at or NOMINAL[t]
    m = MQEnc()
    enc_generic(m, {}, rows, t, tpgd, at)
    return d + at_bytes(at) + m.flush()


def sort_symbols(symbols):
    """辞書に並ぶ順（高さ、幅の順）。"""
    return sorted(symbols, key=lambda r: (len(r), len(r[0])))


def _classes(symbols):
    classes: dict = {}
    for s in symbols:
        classes.setdefault(len(s), []).append(s)
    return classes


def _export_runs(flags):
    runs = []
    cur = 0
    i = 0
    while i < len(flags):
        j = i
        while j < len(flags) and flags[j] == cur:
            j += 1
        runs.append(j - i)
        cur ^= 1
        i = j
    return runs


def symbol_dict_arith(symbols, t=0, at=None, n_in=0, exports=None, retain=False, reuse=None) -> tuple[bytes, dict]:
    """symbols：sort_symbols の順の [rows]。算術符号の記号の辞書。(データ, 残す文脈)。"""
    at = at or NOMINAL[t]
    fl = t << 10 | (1 << 9 if retain else 0) | (1 << 8 if reuse is not None else 0)
    n = len(symbols)
    flags = exports if exports is not None else [0] * n_in + [1] * n
    d = struct.pack(">H", fl) + at_bytes(at) + struct.pack(">II", sum(flags), n)
    m = MQEnc()
    iadh, iadw, iaex = {}, {}, {}
    gb = dict(reuse) if reuse is not None else {}
    hc = 0
    classes = _classes(symbols)
    for h in sorted(classes):
        enc_int(m, iadh, h - hc)
        hc = h
        sw = 0
        for s in classes[h]:
            enc_int(m, iadw, len(s[0]) - sw)
            sw = len(s[0])
            enc_generic(m, gb, s, t, False, at)
        enc_int(m, iadw, None)
    for run in _export_runs(flags):
        enc_int(m, iaex, run)
    return d + m.flush(), gb


def symbol_dict_huffman(symbols, std_table, mmr=True) -> bytes:
    """ハフマンの記号の辞書（表 B.4・B.2・B.1。まとめの画像は MMR か、BMSIZE=0 の圧縮なし）。"""
    d = struct.pack(">H", 1) + struct.pack(">II", len(symbols), len(symbols))
    bw = BitWriter()
    t_dh, t_dw, t_bm = std_table(4), std_table(2), std_table(1)
    hc = 0
    classes = _classes(symbols)
    for h in sorted(classes):
        huff_put(bw, t_dh, h - hc)
        hc = h
        sw = 0
        for s in classes[h]:
            huff_put(bw, t_dw, len(s[0]) - sw)
            sw = len(s[0])
        huff_put(bw, t_dw, None)
        coll = [[v for s in classes[h] for v in s[y]] for y in range(h)]
        if mmr:
            b = enc_mmr(coll, eofb=False)
            huff_put(bw, t_bm, len(b))
            bw.raw(b)
        else:
            huff_put(bw, t_bm, 0)
            bw.align()
            for row in coll:
                row8 = row + [0] * (-len(row) % 8)
                bw.raw(bytes(int("".join(map(str, row8[i:i + 8])), 2) for i in range(0, len(row8), 8)))
    t1 = std_table(1)
    for run in _export_runs([1] * len(symbols)):
        huff_put(bw, t1, run)
    return d + bw.bytes()


def text_flags(huff=False, refine=False, log_strips=0, corner=1, transposed=False, op=0, def_pixel=0, ds_offset=0, r_template=0) -> int:
    return (int(huff) | int(refine) << 1 | log_strips << 2 | corner << 4 | int(transposed) << 6 | op << 7 | def_pixel << 9
            | (ds_offset & 31) << 10 | r_template << 15)


def _strips(instances, strips, corner, transposed, syms):
    """instances：(x, y, 記号の番号[, 細かくした rows, rdx, rdy])。仕様の順に並べる：[(行の T, [(S の始まり, S の次, curt, id, extra)])]。"""
    rows: dict = {}
    for inst in instances:
        x, y, sid = inst[:3]
        extra = inst[3:] if len(inst) > 3 else None
        b = extra[0] if extra else syms[sid]
        w, h = len(b[0]), len(b)
        cx = x if corner < 2 else x + w - 1
        cy = y if corner & 1 else y + h - 1
        s, t = (cy, cx) if transposed else (cx, cy)
        # S の進み：左／上の基準の角は置いた後に、右／下の基準の角は置く前に大きさを足す
        ext = (h if transposed else w) - 1
        before = (corner >= 2) if not transposed else not corner & 1
        s0 = s - ext if before else s
        s1 = s if before else s + ext
        st = (t // strips) * strips
        rows.setdefault(st, []).append((s0, s1, t - st, sid, extra))
    return [(t, sorted(v, key=lambda e: e[0])) for t, v in sorted(rows.items())]


def text_seg_arith(w, h, instances, syms, log_strips=0, corner=1, transposed=False, op=0, def_pixel=0, ds_offset=0,
                   x=0, y=0, region_op=0, refine=False, r_template=0, r_at=None) -> bytes:
    from todoku.imaging.jbig2 import _ceil_log2

    fl = text_flags(False, refine, log_strips, corner, transposed, op, def_pixel, ds_offset, r_template)
    d = region_info(w, h, x, y, region_op) + struct.pack(">H", fl)
    if refine and not r_template:
        d += at_bytes(r_at or [(-1, -1), (-1, -1)])
    d += struct.pack(">I", len(instances))
    m = MQEnc()
    c = _text_ctx()
    _enc_text(m, c, instances, syms, _ceil_log2(len(syms)), log_strips, corner, transposed, ds_offset, refine, r_template, r_at)
    return d + m.flush()


def _text_ctx() -> dict:
    return {k: {} for k in ("dt", "fs", "ds", "it", "ri", "rdw", "rdh", "rdx", "rdy", "id", "gr")}


def _enc_text(m, c, instances, syms, code_len, log_strips=0, corner=1, transposed=False, ds_offset=0, refine=False, r_template=0, r_at=None):
    strips = 1 << log_strips
    enc_int(m, c["dt"], 0)
    stript = 0
    firsts = 0
    for st, items in _strips(instances, strips, corner, transposed, syms):
        enc_int(m, c["dt"], (st - stript) // strips)
        stript = st
        curs = None
        for s0, s1, curt, sid, extra in items:
            if curs is None:
                enc_int(m, c["fs"], s0 - firsts)
                firsts = s0
            else:
                enc_int(m, c["ds"], s0 - curs - ds_offset)
            if strips > 1:
                enc_int(m, c["it"], curt)
            enc_iaid(m, c["id"], code_len, sid)
            if refine:
                enc_int(m, c["ri"], 1 if extra else 0)
                if extra:
                    b, rdx, rdy = extra
                    ref = syms[sid]
                    rdw, rdh = len(b[0]) - len(ref[0]), len(b) - len(ref)
                    enc_int(m, c["rdw"], rdw)
                    enc_int(m, c["rdh"], rdh)
                    enc_int(m, c["rdx"], rdx)
                    enc_int(m, c["rdy"], rdy)
                    enc_refine(m, c["gr"], b, ref, r_template, (rdw >> 1) + rdx, (rdh >> 1) + rdy, False, r_at)
            curs = s1
        enc_int(m, c["ds"], None)


def sort_refagg(specs):
    return sorted(specs, key=lambda s: (len(s[0]), len(s[0][0])))


def symbol_dict_refagg(specs, insyms, r_template=0) -> bytes:
    """細かくする／集める辞書（SDREFAGG=1）。specs：sort_refagg の順の [(rows, [(x, y, 元の番号)])]。
    置くものが 1 つなら 6.5.8.2.2 の形、2 つ以上なら字の領域（6.5.8.2.1）。"""
    from todoku.imaging.jbig2 import _ceil_log2

    fl = 2 | r_template << 12
    n_in, n_new = len(insyms), len(specs)
    d = struct.pack(">H", fl) + at_bytes(NOMINAL[0])
    if not r_template:
        d += at_bytes([(-1, -1), (-1, -1)])
    flags = [0] * n_in + [1] * n_new
    d += struct.pack(">II", n_new, n_new)
    m = MQEnc()
    iadh, iadw, iaex, iaai = {}, {}, {}, {}
    c = _text_ctx()
    code_len = _ceil_log2(n_in + n_new)
    syms = list(insyms)
    hc = 0
    classes: dict = {}
    for s in specs:
        classes.setdefault(len(s[0]), []).append(s)
    for h in sorted(classes):
        enc_int(m, iadh, h - hc)
        hc = h
        sw = 0
        for rows, places in classes[h]:
            enc_int(m, iadw, len(rows[0]) - sw)
            sw = len(rows[0])
            enc_int(m, iaai, len(places))
            if len(places) == 1:
                x, y, sid = places[0]
                enc_iaid(m, c["id"], code_len, sid)
                enc_int(m, c["rdx"], x)
                enc_int(m, c["rdy"], y)
                enc_refine(m, c["gr"], rows, syms[sid], r_template, x, y)
            else:
                # 集める：字の領域として。細かくする字は「元の形」そのまま（RDW=RDH=0）。合わせた形が rows に等しいこと
                inst = [(x, y, sid, syms[sid], 0, 0) for x, y, sid in places]
                _enc_text(m, c, inst, syms, code_len, 0, 1, False, 0, True, r_template)
            syms.append(rows)
        enc_int(m, iadw, None)
    for run in _export_runs(flags):
        enc_int(m, iaex, run)
    return d + m.flush()


def text_seg_huffman(w, h, instances, syms, std_table, log_strips=0, corner=1, transposed=False, op=0, ds_offset=0) -> bytes:
    """ハフマンの字の領域（FS B.6・DS B.8・DT B.11）。記号の符号は、すべて同じ長さ。"""
    from todoku.imaging.jbig2 import Huffman, _ceil_log2

    strips = 1 << log_strips
    fl = text_flags(True, False, log_strips, corner, transposed, op, 0, ds_offset)
    d = region_info(w, h) + struct.pack(">H", fl) + struct.pack(">H", 0) + struct.pack(">I", len(instances))
    bw = BitWriter()
    # 7.4.3.1.7：35 個の連の符号の長さ。code_len の符号の長さだけに 1 ビットを割り当て、残りに 34（0 の連）を使わない素直な形
    code_len = max(1, _ceil_log2(len(syms)))
    lens = [0] * 35
    lens[code_len] = 1
    for i in range(35):
        bw.bits(lens[i], 4)
    run = Huffman([(1, 0, code_len, "")])
    for _ in syms:
        huff_put(bw, run, code_len)
    bw.align()
    sym = Huffman([(code_len, 0, i, "") for i in range(len(syms))])
    t_fs, t_ds, t_dt = std_table(6), std_table(8), std_table(11)
    # B.11 に 0 はないので、最初の T は 1 で始める（T の値は -1 × strips から）
    huff_put(bw, t_dt, 1)
    stript = -strips
    firsts = 0
    for st, items in _strips(instances, strips, corner, transposed, syms):
        huff_put(bw, t_dt, (st - stript) // strips)
        stript = st
        curs = None
        for s0, s1, curt, sid, _extra in items:
            if curs is None:
                huff_put(bw, t_fs, s0 - firsts)
                firsts = s0
            else:
                huff_put(bw, t_ds, s0 - curs - ds_offset)
            if strips > 1:
                bw.bits(curt, log_strips)
            huff_put(bw, sym, sid)
            curs = s1
        huff_put(bw, t_ds, None)
    return d + bw.bytes()


def pattern_dict(patterns, t=0, mmr=False) -> bytes:
    pw, ph = len(patterns[0][0]), len(patterns[0])
    coll = [[v for p in patterns for v in p[y]] for y in range(ph)]
    d = bytes([(1 if mmr else 0) | t << 1, pw, ph]) + struct.pack(">I", len(patterns) - 1)
    if mmr:
        return d + enc_mmr(coll)
    at = [(-pw, 0)] + ([(-3, -1), (2, -2), (-2, -2)] if t == 0 else [])
    m = MQEnc()
    enc_generic(m, {}, coll, t, False, at)
    return d + m.flush()


def halftone_seg(w, h, gray, n_patterns, pw, ph, gx=0, gy=0, rx=None, ry=0, t=0, mmr=False, op=0, def_pixel=0, skip=False) -> bytes:
    """gray：[[灰色の値]]（格子の行 mg・列 ng）。rx・ry は 1/256 画素。"""
    from todoku.imaging.jbig2 import _ceil_log2

    gh, gw = len(gray), len(gray[0])
    rx = (pw << 8) if rx is None else rx
    fl = (1 if mmr else 0) | t << 1 | (8 if skip else 0) | op << 4 | def_pixel << 7
    d = region_info(w, h) + bytes([fl]) + struct.pack(">IIiiHH", gw, gh, gx, gy, rx, ry)
    bpp = _ceil_log2(n_patterns)
    sk = None
    if skip:
        sk = [[int(((gx + m * ry + n * rx) >> 8) + pw <= 0 or ((gx + m * ry + n * rx) >> 8) >= w
                   or ((gy + m * rx - n * ry) >> 8) + ph <= 0 or ((gy + m * rx - n * ry) >> 8) >= h)
               for n in range(gw)] for m in range(gh)]
    at = [(3 if t <= 1 else 2, -1)] + ([(-3, -1), (2, -2), (-2, -2)] if t == 0 else [])
    planes = []
    for j in range(bpp - 1, -1, -1):
        planes.append([[(g >> j) & 1 for g in row] for row in gray])
    # 灰色符号：上の桁との XOR
    coded = [planes[0]] + [[[a ^ b for a, b in zip(r0, r1)] for r0, r1 in zip(planes[i], planes[i - 1])] for i in range(1, len(planes))]
    if sk:
        coded = [[[0 if sk[y][x] else v for x, v in enumerate(r)] for y, r in enumerate(p)] for p in coded]
    if mmr:
        return d + b"".join(enc_mmr(p) for p in coded)
    m = MQEnc()
    cx: dict = {}
    for p in coded:
        enc_generic(m, cx, p, t, False, at, sk)
    return d + m.flush()


def refinement_seg(rows, ref, x=0, y=0, op=4, t=0, tpgr=False, at=None) -> bytes:
    w, h = len(rows[0]), len(rows)
    d = region_info(w, h, x, y, op) + bytes([t | (2 if tpgr else 0)])
    if not t:
        d += at_bytes(at or [(-1, -1), (-1, -1)])
    m = MQEnc()
    enc_refine(m, {}, rows, ref, t, 0, 0, tpgr, at)
    return d + m.flush()


# ---------------------------------------------------------------------- 試験の PDF
def pdf_with_jbig2(data: bytes, w: int, h: int, globals_data: bytes | None = None, image_mask: bool = False, decode: str = "") -> bytes:
    """1 画素 = 1 ポイントの 1 ページの PDF（pdftoppm -r 72 で同じ画素の数になる）。"""
    objs = [b"<</Type/Catalog/Pages 2 0 R>>", b"<</Type/Pages/Kids[3 0 R]/Count 1>>",
            f"<</Type/Page/Parent 2 0 R/MediaBox[0 0 {w} {h}]/Contents 4 0 R/Resources<</XObject<</I 5 0 R>>>>>>".encode()]
    cont = f"q {w} 0 0 {h} 0 0 cm /I Do Q".encode()
    objs.append(b"<</Length %d>>stream\n" % len(cont) + cont + b"\nendstream")
    parms = b"/DecodeParms<</JBIG2Globals 6 0 R>>" if globals_data else b""
    kind = b"/ImageMask true" if image_mask else b"/ColorSpace/DeviceGray"
    objs.append(b"<</Type/XObject/Subtype/Image/Width %d/Height %d/BitsPerComponent 1%s/Filter/JBIG2Decode%s%s/Length %d>>stream\n"
                % (w, h, kind, parms, decode.encode(), len(data)) + data + b"\nendstream")
    if globals_data:
        objs.append(b"<</Length %d>>stream\n" % len(globals_data) + globals_data + b"\nendstream")
    out = bytearray(b"%PDF-1.5\n%\xe2\xe3\xcf\xd3\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for o in offs:
        out += b"%010d 00000 n \n" % o
    out += b"trailer<</Size %d/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, x)
    return bytes(out)
