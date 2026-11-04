"""JPEG 2000 Part 1（ISO/IEC 15444-1）の復号：PDF の JPXDecode と、.jp2・.j2k のファイル。

JP2 の箱（附属書 I：ihdr・bpcc・colr・pclr・cmap・cdef・res）と符号の流れ（附属書 A：SIZ・COD・COC・QCD・QCC・RGN・POC・PPM・PPT・
SOT・SOD）。第 2 層（附属書 B：タグの木・パケット・5 つの進み方・POC・SOP・EPH）、第 1 層（附属書 D：jpx_t1.py）、
量子化（附属書 E）、逆ウェーブレット（附属書 F：5/3・9/7）、成分の変換（附属書 G：RCT・ICT）、複数のタイル・成分・層。

時間：純粋な Python なので、画面に要る大きさより細かな解像度の段は開かない（reduce：ウェーブレットの段を減らす。仕様どおりの低い
解像度の画像が出る）。解像度の段を減らすのは、(1) 画面に要る画素の数（target_pixels）、(2) 第 1 層で読むバイトの数（MAX_T1_BYTES：
1 回の予算）と値の数（MAX_VALUES：メモリ）を超えないため。どちらもファイルの中身だけで決まる（同じ書類から同じ画像）。
一部は OpenJPEG（BSD 2-Clause。LICENSES/LICENSE-openjpeg.txt）に基づく。
使った解像度は Image.reduce・Image.notes で伝える。deadline を過ぎたら、残りの符号ブロックを読まず（細部がぼけるだけで、形は残る）、そう伝える。
"""

from __future__ import annotations

import struct
import time
from dataclasses import dataclass, field

from .. import limits
from ..limits import MAX_PIXELS, LimitExceeded
from . import jpx_t1
from .jpx_t1 import BYPASS, NEG, TERMALL

MAX_T1_BYTES = 3_000_000  # 第 1 層で読むバイトの数の上限（密な流れで 1 MB あたり 5 秒ほど：A4・150 dpi を 20 秒以内で開く大きさ。超えれば解像度を下げる）
MAX_VALUES = 12_000_000  # 復号で持つ値の数の上限（メモリ：値 1 つに 40 バイトほど）


class JpxError(ValueError):
    pass


@dataclass
class Image:
    w: int
    h: int
    planes: list  # 成分ごとの bytes（8 ビット、w × h に広げた）。パレットは開いたあと
    colorspace: str | None  # "sRGB"・"gray"・"sYCC"・"CMYK"・None（分からない）
    alpha: bytes | None
    reduce: int  # 解像度の段を何段減らしたか
    full_w: int
    full_h: int
    notes: list = field(default_factory=list)
    problems: list = field(default_factory=list)
    stopped: bool = False
    dpi: tuple | None = None  # (横, 縦) の 1 インチの画素（res の箱にあれば）
    bits: list = field(default_factory=list)  # 成分ごとのビットの数
    at_limit: bool = False  # 予算のために解像度を下げた（これより細かくは開かない。範囲 window を狭めれば細かくなる）
    window: tuple | None = None  # 開いた範囲（元の画素で [左, 上, 右, 下)）。None は全体


# ====================================================================== 箱（JP2）
@dataclass
class _JP2:
    codestream: bytes = b""
    colorspace: str | None = None
    icc_n: int = 0
    palette: list | None = None  # 列ごとの値の並び
    palette_bits: list | None = None
    cmap: list | None = None
    cdef: list | None = None
    bpc: list | None = None
    dpi: tuple | None = None


_ENUM = {16: "sRGB", 17: "gray", 18: "sYCC", 12: "CMYK", 14: "CIELab"}


def _boxes(data: bytes, pos: int, end: int):
    while pos + 8 <= end:
        ln, t = struct.unpack_from(">I4s", data, pos)
        hdr = 8
        if ln == 1:
            ln = struct.unpack_from(">Q", data, pos + 8)[0]
            hdr = 16
        elif ln == 0:
            ln = end - pos
        if ln < hdr:
            return
        yield t, pos + hdr, min(pos + ln, end)
        pos += ln


def _parse_jp2(data: bytes) -> _JP2:
    j = _JP2()
    if data[:2] == b"\xff\x4f":
        j.codestream = data
        return j
    if data[4:8] != b"jP  ":
        i = data.find(b"\xff\x4f\xff\x51")
        if i < 0:
            raise JpxError("JPEG 2000 の符号の流れが見つかりません")
        j.codestream = data[i:]
        return j
    for t, s, e in _boxes(data, 0, len(data)):
        if t == b"jp2c" and not j.codestream:
            j.codestream = data[s:e]
        elif t == b"jp2h":
            for t2, s2, e2 in _boxes(data, s, e):
                if t2 == b"colr" and j.colorspace is None and not j.icc_n:
                    meth = data[s2]
                    if meth == 1:
                        j.colorspace = _ENUM.get(struct.unpack_from(">I", data, s2 + 3)[0])
                    elif meth == 2 and e2 - s2 >= 3 + 20:
                        sig = data[s2 + 3 + 16:s2 + 3 + 20]
                        j.icc_n = {b"GRAY": 1, b"RGB ": 3, b"CMYK": 4, b"Lab ": 3}.get(sig, 0)
                elif t2 == b"bpcc":
                    j.bpc = list(data[s2:e2])
                elif t2 == b"pclr":
                    ne, npc = struct.unpack_from(">HB", data, s2)
                    bits = [data[s2 + 3 + i] for i in range(npc)]
                    p = s2 + 3 + npc
                    cols = [[0] * ne for _ in range(npc)]
                    for k in range(ne):
                        for i in range(npc):
                            nb = ((bits[i] & 0x7F) + 8) // 8
                            cols[i][k] = int.from_bytes(data[p:p + nb], "big")
                            p += nb
                    j.palette, j.palette_bits = cols, [(b & 0x7F) + 1 for b in bits]
                elif t2 == b"cmap":
                    j.cmap = [struct.unpack_from(">HBB", data, q) for q in range(s2, e2 - 3, 4)]
                elif t2 == b"cdef":
                    n = struct.unpack_from(">H", data, s2)[0]
                    j.cdef = [struct.unpack_from(">HHH", data, s2 + 2 + 6 * i) for i in range(n)]
                elif t2 == b"res ":
                    for t3, s3, e3 in _boxes(data, s2, e2):
                        if t3 in (b"resd", b"resc") and e3 - s3 >= 10 and (t3 == b"resd" or j.dpi is None):
                            vn, vd, hn, hd, ve, he = struct.unpack_from(">HHHHbb", data, s3)
                            if vd and hd:
                                j.dpi = (hn / hd * 10.0 ** he * 0.0254, vn / vd * 10.0 ** ve * 0.0254)
    if not j.codestream:
        raise JpxError("JPEG 2000 の符号の流れの箱（jp2c）がありません")
    return j


# ====================================================================== 符号の流れの見出し
@dataclass
class _Coding:
    precincts: list | None = None  # [(PPx, PPy)]
    sop: bool = False
    eph: bool = False
    prog: int = 0
    layers: int = 1
    mct: int = 0
    nl: int = 0
    xcb: int = 6
    ycb: int = 6
    style: int = 0
    reversible: bool = True


@dataclass
class _Quant:
    style: int = 0
    guard: int = 0
    vals: list = field(default_factory=list)  # (ε, μ)


def _ceildiv(a: int, b: int) -> int:
    return -(-a // b)


class _Tile:
    __slots__ = ("index", "cod", "coc", "qcd", "qcc", "rgn", "poc", "parts", "ppt", "tp_hdr", "comps", "x0", "y0", "x1", "y1", "ppm_parts")

    def __init__(self, index):
        self.index = index
        self.cod = None
        self.coc = {}
        self.qcd = None
        self.qcc = {}
        self.rgn = {}
        self.poc = []
        self.parts = []  # [(start, end)]
        self.ppt = []
        self.ppm_parts = []
        self.comps = None


class _Stream:
    """符号の流れを読む。"""

    def __init__(self, data: bytes):
        self.data = data
        self.cod = _Coding()
        self.cod_seen = False
        self.coc: dict = {}
        self.qcd = _Quant()
        self.qcc: dict = {}
        self.rgn: dict = {}
        self.poc: list = []
        self.ppm: dict = {}
        self.tiles: dict = {}
        self.problems: list = []
        self._read()

    def _spcod(self, b: bytes, pos: int, coding: _Coding, has_prec: bool) -> None:
        coding.nl = b[pos]
        coding.xcb = (b[pos + 1] & 15) + 2
        coding.ycb = (b[pos + 2] & 15) + 2
        coding.style = b[pos + 3]
        coding.reversible = b[pos + 4] == 1
        if coding.nl > 32 or coding.xcb > 10 or coding.ycb > 10 or coding.xcb + coding.ycb > 12:
            raise JpxError("JPEG 2000：符号ブロックの大きさか、分解の段の数が範囲の外です")
        if has_prec:
            coding.precincts = [(b[pos + 5 + i] & 15, b[pos + 5 + i] >> 4) for i in range(coding.nl + 1)]
        else:
            coding.precincts = None

    def _quant(self, b: bytes, pos: int) -> _Quant:
        s = b[pos]
        q = _Quant(s & 31, s >> 5)
        pos += 1
        if q.style == 0:
            q.vals = [(v >> 3, 0) for v in b[pos:]]
        else:
            q.vals = [(v >> 11, v & 0x7FF) for v in struct.unpack_from(f">{(len(b) - pos) // 2}H", b, pos)]
        return q

    def _marker(self, m: int, b: bytes, tile: _Tile | None) -> None:
        wide = self.csiz > 256
        cw = 2 if wide else 1

        def comp(p: int) -> int:
            return struct.unpack_from(">H", b, p)[0] if wide else b[p]

        if m == 0xFF52:  # COD
            c = _Coding()
            sc = b[0]
            c.sop, c.eph = bool(sc & 2), bool(sc & 4)
            c.prog, c.layers, c.mct = b[1], struct.unpack_from(">H", b, 2)[0], b[4]
            self._spcod(b, 5, c, bool(sc & 1))
            if tile is None:
                self.cod = c
                self.cod_seen = True
            else:
                tile.cod = c
        elif m == 0xFF53:  # COC
            cc = comp(0)
            sc = b[cw]
            base = tile.cod if tile is not None and tile.cod else self.cod
            c = _Coding(sop=base.sop, eph=base.eph, prog=base.prog, layers=base.layers, mct=base.mct)
            self._spcod(b, cw + 1, c, bool(sc & 1))
            (tile.coc if tile is not None else self.coc)[cc] = c
        elif m == 0xFF5C:  # QCD
            q = self._quant(b, 0)
            if tile is None:
                self.qcd = q
            else:
                tile.qcd = q
        elif m == 0xFF5D:  # QCC
            cc = comp(0)
            (tile.qcc if tile is not None else self.qcc)[cc] = self._quant(b, cw)
        elif m == 0xFF5E:  # RGN
            cc = comp(0)
            (tile.rgn if tile is not None else self.rgn)[cc] = b[cw + 1]
        elif m == 0xFF5F:  # POC
            sz = 5 + 2 * cw
            lst = tile.poc if tile is not None else self.poc
            for i in range(0, len(b) - sz + 1, sz):
                rs = b[i]
                cs = comp(i + 1)
                ly = struct.unpack_from(">H", b, i + 1 + cw)[0]
                re_ = b[i + 3 + cw]
                ce = comp(i + 4 + cw)
                if not wide and ce == 0:
                    ce = 256
                elif wide and ce == 0:
                    ce = 16384
                lst.append((rs, cs, ly, re_, ce, b[i + 4 + 2 * cw]))
        elif m == 0xFF60:  # PPM
            self.ppm[b[0]] = b[1:]
        elif m == 0xFF61 and tile is not None:  # PPT
            tile.ppt.append((b[0], b[1:]))

    def _read(self) -> None:
        d = self.data
        if d[:2] != b"\xff\x4f" or d[2:4] != b"\xff\x51":
            raise JpxError("JPEG 2000：SOC・SIZ で始まっていません")
        ln = struct.unpack_from(">H", d, 4)[0]
        vals = struct.unpack_from(">HIIIIIIIIH", d, 6)
        self.rsiz, self.xsiz, self.ysiz, self.xo, self.yo, self.xt, self.yt, self.xto, self.yto, self.csiz = vals
        if self.csiz == 0 or self.csiz > 64 or not self.xt or not self.yt:
            raise JpxError(f"JPEG 2000：成分の数 {self.csiz} は扱えません")
        self.comps = []
        for i in range(self.csiz):
            ss, xr, yr = d[42 + 3 * i], d[43 + 3 * i], d[44 + 3 * i]
            if not xr or not yr:
                raise JpxError("JPEG 2000：成分の間引きの値が 0 です")
            self.comps.append(((ss & 0x7F) + 1, bool(ss & 0x80), xr, yr))
        w, h = self.xsiz - self.xo, self.ysiz - self.yo
        if w <= 0 or h <= 0:
            raise JpxError("JPEG 2000：画像の大きさが 0 です")
        if w * h > MAX_PIXELS:
            raise LimitExceeded(f"JPEG 2000 の画像（{w:,} × {h:,} 画素）が、開ける大きさ {MAX_PIXELS:,} 画素を超えます")
        self.ntx = _ceildiv(self.xsiz - self.xto, self.xt)
        self.nty = _ceildiv(self.ysiz - self.yto, self.yt)
        if self.ntx * self.nty > 65535:
            raise JpxError("JPEG 2000：タイルの数が多すぎます")
        pos = 4 + ln
        tile = None
        n = len(d)
        ppm_order: list = []
        while pos + 2 <= n:
            m = struct.unpack_from(">H", d, pos)[0]
            if m == 0xFFD9:
                break
            if m == 0xFF90:  # SOT
                if pos + 12 > n:
                    self.problems.append("タイルの見出しが途中で切れています")
                    break
                _ls, isot, psot, tpsot, _tn = struct.unpack_from(">HHIBB", d, pos + 2)
                if isot >= self.ntx * self.nty:
                    self.problems.append(f"タイルの番号 {isot} が範囲の外です")
                    break
                tile = self.tiles.get(isot)
                if tile is None:
                    tile = self.tiles[isot] = _Tile(isot)
                start = pos
                pos += 12
                # タイルの部分の見出し
                while pos + 2 <= n:
                    mm = struct.unpack_from(">H", d, pos)[0]
                    if mm == 0xFF93:
                        pos += 2
                        break
                    ln2 = struct.unpack_from(">H", d, pos + 2)[0]
                    self._marker(mm, d[pos + 4:pos + 2 + ln2], tile)
                    pos += 2 + ln2
                end = start + psot if psot else n
                if end > n:
                    end = n
                    self.problems.append(f"タイル {isot} のデータが途中で切れています")
                if not psot and d[end - 2:end] == b"\xff\xd9":
                    end -= 2
                tile.parts.append((pos, end))
                ppm_order.append(tile)
                pos = end
                continue
            if m in (0xFF30, 0xFF31) or 0xFF30 <= m <= 0xFF3F:
                pos += 2
                continue
            if pos + 4 > n:
                break
            ln2 = struct.unpack_from(">H", d, pos + 2)[0]
            self._marker(m, d[pos + 4:pos + 2 + ln2], None)
            pos += 2 + ln2
        if not self.cod_seen:
            raise JpxError("JPEG 2000：COD の見出しがありません")
        if self.ppm:
            # Zppm の順につなぐと (Nppm, Ippm) の列になる（A.7.4）。タイルの部分の順に割り当てる
            blob = b"".join(self.ppm[k] for k in sorted(self.ppm))
            p = 0
            for t in ppm_order:
                if p + 4 > len(blob):
                    break
                nn = struct.unpack_from(">I", blob, p)[0]
                t.ppm_parts.append(blob[p + 4:p + 4 + nn])
                p += 4 + nn


# ====================================================================== 第 2 層
class _TagTree:
    """タグの木（B.10.2）。"""

    __slots__ = ("val", "low", "parent", "w")

    def __init__(self, w: int, h: int):
        self.w = w
        sizes = [(w, h)]
        while sizes[-1][0] > 1 or sizes[-1][1] > 1:
            a, b = sizes[-1]
            sizes.append(((a + 1) // 2, (b + 1) // 2))
        n = sum(a * b for a, b in sizes)
        self.val = [1 << 60] * n
        self.low = [0] * n
        parent = [-1] * n
        base = 0
        for lv in range(len(sizes) - 1):
            a, b = sizes[lv]
            nxt = base + a * b
            na = sizes[lv + 1][0]
            for y in range(b):
                for x in range(a):
                    parent[base + y * a + x] = nxt + (y // 2) * na + x // 2
            base = nxt
        self.parent = parent


class _Block:
    __slots__ = ("x0", "y0", "x1", "y1", "included", "zb", "lblock", "segs", "npasses")

    def __init__(self, x0, y0, x1, y1):
        self.x0, self.y0, self.x1, self.y1 = x0, y0, x1, y1
        self.included = False
        self.zb = 0
        self.lblock = 3
        self.segs = []  # [[データの部品の並び, 通り道の数, 上限]]
        self.npasses = 0


class _PBand:
    """ある precinct の、ある帯の符号ブロック。"""

    __slots__ = ("nw", "nh", "blocks", "incl", "zero", "band")

    def __init__(self, band, x0, y0, x1, y1, xcb, ycb):
        self.band = band
        if x1 <= x0 or y1 <= y0:
            self.nw = self.nh = 0
            self.blocks = []
            self.incl = self.zero = None
            return
        cx0, cy0 = x0 >> xcb, y0 >> ycb
        cx1, cy1 = _ceildiv(x1, 1 << xcb), _ceildiv(y1, 1 << ycb)
        self.nw, self.nh = cx1 - cx0, cy1 - cy0
        self.blocks = [_Block(max(x0, cx << xcb), max(y0, cy << ycb), min(x1, (cx + 1) << xcb), min(y1, (cy + 1) << ycb))
                       for cy in range(cy0, cy1) for cx in range(cx0, cx1)]
        self.incl = _TagTree(self.nw, self.nh)
        self.zero = _TagTree(self.nw, self.nh)


class _Band:
    __slots__ = ("orient", "x0", "y0", "x1", "y1", "numbps", "step", "blocks", "rows")

    def __init__(self, orient, x0, y0, x1, y1):
        self.orient, self.x0, self.y0, self.x1, self.y1 = orient, x0, y0, x1, y1
        self.blocks = []
        self.rows = None


class _Res:
    __slots__ = ("x0", "y0", "x1", "y1", "bands", "npw", "nph", "precincts", "ppx", "ppy", "done")


class _TileComp:
    __slots__ = ("x0", "y0", "x1", "y1", "res", "coding", "quant", "roi", "nl")


def _build_comp(st: _Stream, tile: _Tile, c: int, tx0, ty0, tx1, ty1, counter: list) -> _TileComp:
    ss, signed, xr, yr = st.comps[c]
    # 優先の順（A.6）：タイルの COC ＞ タイルの COD ＞ 全体の COC ＞ 全体の COD
    if c in tile.coc:
        coding = tile.coc[c]
    elif tile.cod is not None:
        coding = tile.cod
    elif c in st.coc:
        coding = st.coc[c]
    else:
        coding = st.cod
    quant = tile.qcc.get(c) or tile.qcd or st.qcc.get(c) or st.qcd
    tc = _TileComp()
    tc.coding, tc.quant = coding, quant
    tc.roi = tile.rgn.get(c, st.rgn.get(c, 0))
    tc.x0, tc.y0 = _ceildiv(tx0, xr), _ceildiv(ty0, yr)
    tc.x1, tc.y1 = _ceildiv(tx1, xr), _ceildiv(ty1, yr)
    nl = tc.nl = coding.nl
    tc.res = []
    for r in range(nl + 1):
        nd = nl - r
        res = _Res()
        k = 1 << nd
        res.x0, res.y0 = _ceildiv(tc.x0, k), _ceildiv(tc.y0, k)
        res.x1, res.y1 = _ceildiv(tc.x1, k), _ceildiv(tc.y1, k)
        ppx, ppy = coding.precincts[r] if coding.precincts else (15, 15)
        res.ppx, res.ppy = ppx, ppy
        res.npw = (_ceildiv(res.x1, 1 << ppx) - (res.x0 >> ppx)) if res.x1 > res.x0 else 0
        res.nph = (_ceildiv(res.y1, 1 << ppy) - (res.y0 >> ppy)) if res.y1 > res.y0 else 0
        counter[0] += res.npw * res.nph
        if counter[0] > 2_000_000:
            raise LimitExceeded("JPEG 2000 の区画（precinct）の数が多すぎます")
        res.bands = []
        orients = (0,) if r == 0 else (1, 2, 3)
        for o in orients:
            nb = nl if r == 0 else nl - r + 1
            xo, yo = (o & 1), (o >> 1)
            k2 = 1 << nb
            half = k2 >> 1
            bx0 = _ceildiv(tc.x0 - half * xo, k2) if nb else tc.x0
            by0 = _ceildiv(tc.y0 - half * yo, k2) if nb else tc.y0
            bx1 = _ceildiv(tc.x1 - half * xo, k2) if nb else tc.x1
            by1 = _ceildiv(tc.y1 - half * yo, k2) if nb else tc.y1
            band = _Band(o, bx0, by0, bx1, by1)
            # 量子化（E.1）
            if quant.style == 1:
                e0, mu = quant.vals[0] if quant.vals else (0, 0)
                # 導出（式 E-5）：(εb, μb) = (ε0 − NL + nb, μ0)
                eps = e0 - nl + nb
            else:
                bi = 0 if r == 0 else 3 * (r - 1) + o
                eps, mu = quant.vals[bi] if bi < len(quant.vals) else (quant.vals[-1] if quant.vals else (0, 0))
            band.numbps = quant.guard + eps - 1 + tc.roi
            gain = 0 if o == 0 else 2 if o == 3 else 1
            if coding.reversible or quant.style == 0:
                band.step = 1.0
            else:
                band.step = 2.0 ** (ss + gain - eps) * (1.0 + mu / 2048.0)
            res.bands.append(band)
        res.precincts = {}
        tc.res.append(res)
    return tc


def _precinct(res: _Res, tc: _TileComp, r: int, k: int) -> dict:
    """区画 k（行の順）の、帯ごとの符号ブロックを作る（要るときに）。"""
    p = res.precincts.get(k)
    if p is not None:
        return p
    coding = tc.coding
    ppx, ppy = res.ppx, res.ppy
    px = (res.x0 >> ppx) + k % res.npw
    py = (res.y0 >> ppy) + k // res.npw
    # 帯の座標での区画の大きさ：r > 0 では半分
    sx = ppx - 1 if r else ppx
    sy = ppy - 1 if r else ppy
    xcb = min(coding.xcb, sx)
    ycb = min(coding.ycb, sy)
    bands = []
    for b in res.bands:
        x0, y0 = max(b.x0, px << sx), max(b.y0, py << sy)
        x1, y1 = min(b.x1, (px + 1) << sx), min(b.y1, (py + 1) << sy)
        pb = _PBand(b, x0, y0, x1, y1, xcb, ycb)
        b.blocks += pb.blocks
        bands.append(pb)
    p = {"bands": bands, "next": 0}
    res.precincts[k] = p
    return p


def _packets(st: _Stream, tile: _Tile, tcs: list, tx0, ty0, tx1, ty1):
    """パケットの順（B.12）：(層, 解像度, 成分, 区画の番号)。POC があれば、その進み方を順に。"""
    cod = tile.cod or st.cod
    layers = cod.layers
    pocs = tile.poc or st.poc
    if pocs:
        ranges = [(p[5], p[0], p[3], p[1], p[4], p[2]) for p in pocs]
    else:
        ranges = [(cod.prog, 0, 33, 0, len(tcs), layers)]
    for prog, r0, r1, c0, c1, ly in ranges:
        c1 = min(c1, len(tcs))
        ly = min(ly, layers)
        maxres = max((tc.nl for tc in tcs[c0:c1]), default=0) + 1
        r1 = min(r1, maxres)
        if prog == 0:  # LRCP
            for l in range(ly):
                for r in range(r0, r1):
                    for c in range(c0, c1):
                        tc = tcs[c]
                        if r <= tc.nl:
                            res = tc.res[r]
                            for k in range(res.npw * res.nph):
                                yield l, r, c, k
        elif prog == 1:  # RLCP
            for r in range(r0, r1):
                for l in range(ly):
                    for c in range(c0, c1):
                        tc = tcs[c]
                        if r <= tc.nl:
                            res = tc.res[r]
                            for k in range(res.npw * res.nph):
                                yield l, r, c, k
        elif prog in (2, 3, 4):
            yield from _position_order(st, prog, tcs, r0, r1, c0, c1, ly, tx0, ty0, tx1, ty1)
        else:
            raise JpxError(f"JPEG 2000：進み方 {prog} は規格にありません")


def _position_order(st, prog, tcs, r0, r1, c0, c1, ly, tx0, ty0, tx1, ty1):
    """RPCL・PCRL・CPRL（B.12.1.3〜5）：画像の位置 (y, x) を、区画の頭の所だけたどる。"""

    def at(c, r, x, y):
        """(x, y) が成分 c・解像度 r の区画の頭なら、その番号。"""
        tc = tcs[c]
        if r > tc.nl:
            return None
        res = tc.res[r]
        if not res.npw or not res.nph:
            return None
        _ss, _sg, xr, yr = st.comps[c]
        nd = tc.nl - r
        ppx, ppy = res.ppx, res.ppy
        if not (y % (yr << (ppy + nd)) == 0 or (y == ty0 and ((res.y0 << nd) % (1 << (ppy + nd))) != 0)):
            return None
        if not (x % (xr << (ppx + nd)) == 0 or (x == tx0 and ((res.x0 << nd) % (1 << (ppx + nd))) != 0)):
            return None
        k = ((_ceildiv(x, xr << nd) >> ppx) - (res.x0 >> ppx)
             + res.npw * ((_ceildiv(y, yr << nd) >> ppy) - (res.y0 >> ppy)))
        return k

    def positions(cs, rs):
        xs: set = {tx0}
        ys: set = {ty0}
        for c in cs:
            tc = tcs[c]
            _ss, _sg, xr, yr = st.comps[c]
            for r in rs:
                if r > tc.nl:
                    continue
                res = tc.res[r]
                nd = tc.nl - r
                sx, sy = xr << (res.ppx + nd), yr << (res.ppy + nd)
                # 区画の数だけ足す（細かすぎる格子で手間が増えないように、その段にある頭だけ）
                x = _ceildiv(tx0, sx) * sx
                while x < tx1:
                    xs.add(x)
                    x += sx
                y = _ceildiv(ty0, sy) * sy
                while y < ty1:
                    ys.add(y)
                    y += sy
        return sorted(ys), sorted(xs)

    if prog == 2:  # RPCL
        for r in range(r0, r1):
            ys, xs = positions(range(c0, c1), [r])
            for y in ys:
                for x in xs:
                    for c in range(c0, c1):
                        k = at(c, r, x, y)
                        if k is not None:
                            for l in range(ly):
                                yield l, r, c, k
    elif prog == 3:  # PCRL
        ys, xs = positions(range(c0, c1), range(r0, r1))
        for y in ys:
            for x in xs:
                for c in range(c0, c1):
                    for r in range(r0, r1):
                        k = at(c, r, x, y)
                        if k is not None:
                            for l in range(ly):
                                yield l, r, c, k
    else:  # CPRL
        for c in range(c0, c1):
            ys, xs = positions([c], range(r0, r1))
            for y in ys:
                for x in xs:
                    for r in range(r0, r1):
                        k = at(c, r, x, y)
                        if k is not None:
                            for l in range(ly):
                                yield l, r, c, k


class _Trunc(Exception):
    pass


def _read_packets(st: _Stream, tile: _Tile, tcs: list, tx0, ty0, tx1, ty1, problems: list) -> None:
    """第 2 層：パケットの見出しと本体を読み、符号ブロックごとにデータの部品を集める。"""
    d = st.data
    cod = tile.cod or st.cod
    parts = tile.parts
    body = b"".join(d[s:e] for s, e in parts) if len(parts) != 1 else None
    if body is None:
        pos, end = parts[0]
        body = d
    else:
        pos, end = 0, len(body)
    hdr_stream = None
    if tile.ppm_parts:
        hdr_stream = b"".join(tile.ppm_parts)
    elif tile.ppt:
        hdr_stream = b"".join(b for _z, b in sorted(tile.ppt, key=lambda t: t[0]))
    hpos = 0
    sop, eph = cod.sop, cod.eph
    for l, r, c, k in _packets(st, tile, tcs, tx0, ty0, tx1, ty1):
        tc = tcs[c]
        res = tc.res[r]
        pct = _precinct(res, tc, r, k)
        if l < pct["next"]:
            continue  # POC：同じパケットは 2 度ない
        pct["next"] = l + 1
        if pos >= end and hdr_stream is None:
            return  # データはパケットの区切りで終わり（Psot より短いときは、見出しを読む所で伝えてある）
        if sop and pos + 6 <= end and body[pos] == 0xFF and body[pos + 1] == 0x91:
            pos += 6
        # 見出しのビットを読む（B.10.1：0xFF の次のバイトは 7 ビット）
        src = body if hdr_stream is None else hdr_stream
        p = pos if hdr_stream is None else hpos
        lim = end if hdr_stream is None else len(hdr_stream)
        buf = 0
        cnt = 0

        def bit() -> int:
            nonlocal buf, cnt, p
            if cnt == 0:
                if p >= lim:
                    raise _Trunc
                prev = buf
                buf = src[p]
                p += 1
                cnt = 7 if prev == 0xFF else 8
            cnt -= 1
            return (buf >> cnt) & 1

        def tag(tree: _TagTree, leaf: int, thresh: int) -> bool:
            val, low, parent = tree.val, tree.low, tree.parent
            stack = []
            n = leaf
            while parent[n] >= 0:
                stack.append(n)
                n = parent[n]
            lo = 0
            while True:
                if lo > low[n]:
                    low[n] = lo
                else:
                    lo = low[n]
                while lo < thresh and lo < val[n]:
                    if bit():
                        val[n] = lo
                    else:
                        lo += 1
                low[n] = lo
                if not stack:
                    break
                n = stack.pop()
            return val[n] < thresh

        adds = []
        try:
            if bit():
                for pb in pct["bands"]:
                    if not pb.blocks:
                        continue
                    for bi, blk in enumerate(pb.blocks):
                        if not blk.included:
                            inc = tag(pb.incl, bi, l + 1)
                        else:
                            inc = bit()
                        if not inc:
                            continue
                        if not blk.included:
                            blk.included = True
                            z = 1
                            while not tag(pb.zero, bi, z):
                                z += 1
                                if z > 80:
                                    raise JpxError("JPEG 2000：パケットの見出しが読めません（0 のビット面が多すぎます）")
                            blk.zb = z - 1
                        # 通り道の数（表 B.4）
                        if not bit():
                            np_ = 1
                        elif not bit():
                            np_ = 2
                        else:
                            v = bit() << 1 | bit()
                            if v < 3:
                                np_ = 3 + v
                            else:
                                v = 0
                                for _ in range(5):
                                    v = v << 1 | bit()
                                if v < 31:
                                    np_ = 6 + v
                                else:
                                    v = 0
                                    for _ in range(7):
                                        v = v << 1 | bit()
                                    np_ = 37 + v
                        while bit():
                            blk.lblock += 1
                        style = tc.coding.style
                        # 符号語の区切りに振り分ける（TERMALL：通り道ごと、BYPASS：10 通り道のあと 2・1・2・1）
                        left = np_
                        segs = blk.segs
                        while left > 0:
                            if not segs or segs[-1][1] >= segs[-1][2]:
                                if style & TERMALL:
                                    mx = 1
                                elif style & BYPASS:
                                    mx = 10 if not segs else (2 if segs[-1][2] in (1, 10) else 1)
                                else:
                                    mx = 1 << 30
                                segs.append([[], 0, mx])
                            sg = segs[-1]
                            take = min(left, sg[2] - sg[1])
                            nbits = blk.lblock + (31 - _clz32(take))
                            v = 0
                            for _ in range(nbits):
                                v = v << 1 | bit()
                            sg[1] += take
                            left -= take
                            adds.append((sg, v))
                        blk.npasses += np_
            # 見出しの終わり：次のバイトの境へ（0xFF で終わったときは詰めのバイトを読み飛ばす）
            if buf == 0xFF:
                p += 1
        except _Trunc:
            problems.append("データが途中で切れています（読めた層までで開きます）")
            return
        if eph and p + 2 <= lim and src[p] == 0xFF and src[p + 1] == 0x92:
            p += 2
        if hdr_stream is None:
            pos = p
        else:
            hpos = p
        for sg, v in adds:
            if pos + v > end:
                sg[0].append(body[pos:end])
                problems.append("データが途中で切れています（読めた層までで開きます）")
                return
            sg[0].append(body[pos:pos + v])
            pos += v


def _clz32(v: int) -> int:
    return 32 - v.bit_length()


# ====================================================================== 逆ウェーブレット（附属書 F）
def _synth_53(a: list, b: list, even: bool) -> list:
    """1 行：a（偶数の座標＝低域）、b（奇数の座標＝高域）→ 元の並び。"""
    na, nb = len(a), len(b)
    if na + nb == 1:
        return a if na else [v >> 1 for v in b]
    if even:
        lf = [b[0]] + b
        rt = b + [b[-1]]
    else:
        lf = b
        rt = b[1:] + [b[-1]]
    e = [x - ((p + q + 2) >> 2) for x, p, q in zip(a, lf, rt)]
    if even:
        lf2 = e
        rt2 = e[1:] + [e[-1]]
    else:
        lf2 = [e[0]] + e
        rt2 = e + [e[-1]]
    o = [y + ((p + q) >> 1) for y, p, q in zip(b, lf2, rt2)]
    out = [0] * (na + nb)
    if even:
        out[0::2] = e
        out[1::2] = o
    else:
        out[1::2] = e
        out[0::2] = o
    return out


_A, _B, _G, _D, _K = -1.586134342059924, -0.052980118572961, 0.882911075530934, 0.443506852043971, 1.230174104914001


def _synth_97(a: list, b: list, even: bool) -> list:
    na, nb = len(a), len(b)
    if na + nb == 1:
        return a if na else [v * 0.5 for v in b]
    ik = 1.0 / _K
    e = [x * _K for x in a]
    o = [x * ik for x in b]
    for coef_e, coef_o in ((_D, _G), (_B, _A)):
        if even:
            e = [x - coef_e * (p + q) for x, p, q in zip(e, [o[0]] + o, o + [o[-1]])]
            o = [y - coef_o * (p + q) for y, p, q in zip(o, e, e[1:] + [e[-1]])]
        else:
            e = [x - coef_e * (p + q) for x, p, q in zip(e, o, o[1:] + [o[-1]])]
            o = [y - coef_o * (p + q) for y, p, q in zip(o, [e[0]] + e, e + [e[-1]])]
    out = [0.0] * (na + nb)
    if even:
        out[0::2] = e
        out[1::2] = o
    else:
        out[1::2] = e
        out[0::2] = o
    return out


def _vsynth_53(lo: list, hi: list, even: bool) -> list:
    """縦：lo（偶数の行）・hi（奇数の行）。"""
    na, nb = len(lo), len(hi)
    if na + nb == 1:
        return lo if na else [[v >> 1 for v in r] for r in hi]
    if even:
        lf = [hi[0]] + hi
        rt = hi + [hi[-1]]
    else:
        lf = hi
        rt = hi[1:] + [hi[-1]]
    e = [[x - ((p + q + 2) >> 2) for x, p, q in zip(r, u, v)] for r, u, v in zip(lo, lf, rt)]
    if even:
        lf2, rt2 = e, e[1:] + [e[-1]]
    else:
        lf2, rt2 = [e[0]] + e, e + [e[-1]]
    o = [[y + ((p + q) >> 1) for y, p, q in zip(r, u, v)] for r, u, v in zip(hi, lf2, rt2)]
    out = [None] * (na + nb)
    if even:
        out[0::2] = e
        out[1::2] = o
    else:
        out[1::2] = e
        out[0::2] = o
    return out


def _vsynth_97(lo: list, hi: list, even: bool) -> list:
    na, nb = len(lo), len(hi)
    if na + nb == 1:
        return lo if na else [[v * 0.5 for v in r] for r in hi]
    ik = 1.0 / _K
    e = [[x * _K for x in r] for r in lo]
    o = [[x * ik for x in r] for r in hi]
    for coef_e, coef_o in ((_D, _G), (_B, _A)):
        if even:
            l1, r1 = [o[0]] + o, o + [o[-1]]
        else:
            l1, r1 = o, o[1:] + [o[-1]]
        e = [[x - coef_e * (p + q) for x, p, q in zip(r, u, v)] for r, u, v in zip(e, l1, r1)]
        if even:
            l2, r2 = e, e[1:] + [e[-1]]
        else:
            l2, r2 = [e[0]] + e, e + [e[-1]]
        o = [[y - coef_o * (p + q) for y, p, q in zip(r, u, v)] for r, u, v in zip(o, l2, r2)]
    out = [None] * (na + nb)
    if even:
        out[0::2] = e
        out[1::2] = o
    else:
        out[1::2] = e
        out[0::2] = o
    return out


def _inverse_level(ll: list, res: _Res, rev: bool) -> list:
    """1 段の 2D_SR（F.3.2）：横（HOR_SR）を先に、次に縦（VER_SR）。"""
    hl, lh, hh = res.bands
    w = res.x1 - res.x0
    h = res.y1 - res.y0
    if w <= 0 or h <= 0:
        return []
    even_x = not res.x0 & 1
    even_y = not res.y0 & 1
    syn = _synth_53 if rev else _synth_97
    zero = 0 if rev else 0.0
    nlo = ll
    n_even = (res.y1 + 1) // 2 - (res.y0 + 1) // 2
    n_odd = h - n_even
    na = (res.x1 + 1) // 2 - (res.x0 + 1) // 2
    nb = w - na
    zb = [zero] * nb
    za = [zero] * na
    low = [syn(nlo[i] if i < len(nlo) else za, hl.rows[i] if hl.rows else zb, even_x) for i in range(n_even)]
    high = [syn(lh.rows[i] if lh.rows else za, hh.rows[i] if hh.rows else zb, even_x) for i in range(n_odd)]
    return (_vsynth_53 if rev else _vsynth_97)(low, high, even_y)


# ====================================================================== 全体
def _block_to_band(blk: _Block, band: _Band, rev: bool, orient: int, style: int, roi: int) -> None:
    segs = [(b"".join(s[0]), s[1]) for s in blk.segs if s[1] > 0]
    if not segs:
        return
    numbps = band.numbps - blk.zb
    if numbps <= 0:
        return
    if numbps > 62:
        raise JpxError("JPEG 2000：ビット面の数が多すぎます")
    w, h = blk.x1 - blk.x0, blk.y1 - blk.y0
    mag, flags = jpx_t1.decode_block(segs, w, h, numbps, orient, style)
    st = w + 2
    rows = band.rows
    x0 = blk.x0 - band.x0
    y0 = blk.y0 - band.y0
    # 範囲だけを開くとき：帯の範囲の外に出る所は書かない
    bw, bh = band.x1 - band.x0, band.y1 - band.y0
    sk, en = max(0, -x0), min(w, bw - x0)
    if en <= sk:
        return
    thresh = 1 << (roi + 1) if roi else 0
    hs = band.step * 0.5
    for j in range(max(0, -y0), min(h, bh - y0)):
        b = (j + 1) * st + 1
        mm = mag[b + sk:b + en]
        if not any(mm):
            continue
        ff = flags[b + sk:b + en]
        if roi:
            mm = [m >> roi if m >= thresh else m for m in mm]
        if rev:
            rows[y0 + j][x0 + sk:x0 + en] = [-(m >> 1) if f & NEG else m >> 1 for m, f in zip(mm, ff)]
        else:
            rows[y0 + j][x0 + sk:x0 + en] = [-m * hs if f & NEG else m * hs for m, f in zip(mm, ff)]


def _output_size(st: _Stream, c: int, reduce: int) -> tuple:
    _ss, _sg, xr, yr = st.comps[c]
    k = 1 << reduce
    x0 = _ceildiv(_ceildiv(st.xo, xr), k)
    y0 = _ceildiv(_ceildiv(st.yo, yr), k)
    x1 = _ceildiv(_ceildiv(st.xsiz, xr), k)
    y1 = _ceildiv(_ceildiv(st.ysiz, yr), k)
    return x0, y0, x1, y1


def _comp_rects(st: _Stream, c: int, tc: _TileComp, win: tuple, rd: int) -> dict:
    """範囲（基準の格子の [左, 上, 右, 下)）を開くのに要る、解像度の段ごとの四角 {r: (x0, y0, x1, y1)}。
    逆ウェーブレットは近くの係数しか見ない（F.3：5/3 は ±2、9/7 は ±4）。段ごとにその分を足した四角を普通に合成すれば、足した分の内側は
    全体を開いたときと同じ値になる（四角の端の折り返しの影響は、足した分の中で消える）。"""
    _ss, _sg, xr, yr = st.comps[c]
    k = 1 << rd
    wx0, wy0 = win[0] // xr // k, win[1] // yr // k
    wx1, wy1 = _ceildiv(_ceildiv(win[2], xr), k), _ceildiv(_ceildiv(win[3], yr), k)
    m = 3 if tc.coding.reversible else 5
    rects = {}
    for r in range(tc.nl - rd, -1, -1):
        res = tc.res[r]
        A, C, B, D = max(res.x0, wx0 - m), max(res.y0, wy0 - m), min(res.x1, wx1 + m), min(res.y1, wy1 + m)
        if B <= A or D <= C:
            A, C, B, D = res.x0, res.y0, res.x0, res.y0
        rects[r] = (A, C, B, D)
        wx0, wy0, wx1, wy1 = (A + 1) >> 1, (C + 1) >> 1, (B + 1) >> 1, (D + 1) >> 1
    return rects


def _band_rect(o: int, r: int, rect: tuple) -> tuple:
    """解像度の段の四角のうち、帯 o（1 HL・2 LH・3 HH）の係数の四角。低域は偶数の座標、高域は奇数の座標（F.3.3）。r＝0 はそのまま。"""
    A, C, B, D = rect
    if r == 0:
        return rect
    x0, x1 = (A >> 1, B >> 1) if o & 1 else ((A + 1) >> 1, (B + 1) >> 1)
    y0, y1 = (C >> 1, D >> 1) if o & 2 else ((C + 1) >> 1, (D + 1) >> 1)
    return x0, y0, x1, y1


def _block_bytes(blk: _Block) -> int:
    return sum(len(x) for sg in blk.segs for x in sg[0])


def _plan_cost(st: _Stream, tiles: list, win: tuple, rd: int) -> tuple:
    """解像度を rd 段減らして範囲を開くときの、(第 1 層で読むバイト, 持つ値の数)。"""
    size = vals = 0
    for _t, tcs, *_r in tiles:
        for c, tc in enumerate(tcs):
            rects = _comp_rects(st, c, tc, win, min(rd, tc.nl))
            top = max(rects)
            A, C, B, D = rects[top]
            vals += (B - A) * (D - C)
            for r, rect in rects.items():
                for b in tc.res[r].bands:
                    x0, y0, x1, y1 = _band_rect(b.orient, r, rect)
                    if x1 <= x0 or y1 <= y0:
                        continue
                    for blk in b.blocks:
                        if blk.segs and blk.x0 < x1 and blk.x1 > x0 and blk.y0 < y1 and blk.y1 > y0:
                            size += _block_bytes(blk)
    return size, vals


def decode(data: bytes, target_pixels: int | None = None, reduce: int | None = None, deadline: float | None = None,
           max_t1_bytes: int | None = None, window: tuple | None = None) -> Image:
    """JPEG 2000 を開く。target_pixels：画面に要る画素の数（画像の全体で。これを下回らない範囲で解像度の段を減らす）。reduce：段の数を直接指定。
    window：[左, 上, 右, 下)（元の画素）を指定すると、そこに掛かるタイルと符号ブロックだけを開く（結果の画像は、その範囲だけ）。予算（MAX_T1_BYTES）は
    開く分に掛かるので、範囲を狭めれば、どの部分でも元の解像度まで開ける。"""
    j2 = _parse_jp2(data)
    st = _Stream(j2.codestream)
    problems = list(st.problems)
    notes: list = []
    max_t1_bytes = MAX_T1_BYTES if max_t1_bytes is None else max_t1_bytes
    W, H = st.xsiz - st.xo, st.ysiz - st.yo
    win = (st.xo, st.yo, st.xsiz, st.ysiz)
    if window is not None:
        fx = max(c[2] for c in st.comps)
        fy = max(c[3] for c in st.comps)
        wx0, wy0, wx1, wy1 = (int(v) for v in window)
        wx1, wy1 = min(W, max(wx1, wx0 + 1)), min(H, max(wy1, wy0 + 1))
        wx0, wy0 = max(0, min(wx0, W - 1)), max(0, min(wy0, H - 1))
        wx0, wy0, wx1, wy1 = wx0 + st.xo, wy0 + st.yo, wx1 + st.xo, wy1 + st.yo
        # 間引いた成分とそろうように、基準の格子で間引きの倍数に広げる
        win = (wx0 // fx * fx, wy0 // fy * fy, min(st.xsiz, _ceildiv(wx1, fx) * fx), min(st.ysiz, _ceildiv(wy1, fy) * fy))
    # タイルを読む（第 2 層は全部読む：符号ブロックのバイトの数が分かる。範囲の外のタイルは読まない）
    tiles = []
    counter = [0]
    seen = 0
    for idx in sorted(st.tiles):
        t = st.tiles[idx]
        p, q = idx % st.ntx, idx // st.ntx
        tx0 = max(st.xto + p * st.xt, st.xo)
        ty0 = max(st.yto + q * st.yt, st.yo)
        tx1 = min(st.xto + (p + 1) * st.xt, st.xsiz)
        ty1 = min(st.yto + (q + 1) * st.yt, st.ysiz)
        seen += 1
        if tx0 >= win[2] or tx1 <= win[0] or ty0 >= win[3] or ty1 <= win[1]:
            continue
        tcs = [_build_comp(st, t, c, tx0, ty0, tx1, ty1, counter) for c in range(st.csiz)]
        try:
            _read_packets(st, t, tcs, tx0, ty0, tx1, ty1, problems)
        except (IndexError, struct.error) as e:
            problems.append(f"タイル {idx} のパケットを読めません（{type(e).__name__}）")
        tiles.append((t, tcs, tx0, ty0, tx1, ty1))
    if not seen:
        raise JpxError("JPEG 2000：タイルのデータがありません")
    if not tiles:
        raise JpxError("JPEG 2000：指定の範囲にタイルのデータがありません")
    max_reduce = min(tc.nl for _t, tcs, *_r in tiles for tc in tcs)
    wide, tall = win[2] - win[0], win[3] - win[1]

    def cost(rd: int) -> tuple:
        return _plan_cost(st, tiles, win, rd)

    at_limit = False
    if reduce is None:
        rd = 0
        if target_pixels:
            while rd < max_reduce and ((W >> (rd + 1)) * (H >> (rd + 1))) >= target_pixels:
                rd += 1
        forced = rd
        kept = first = cost(rd)
        while rd < max_reduce and (kept[0] > max_t1_bytes or kept[1] > MAX_VALUES):
            rd += 1
            kept = cost(rd)
        if rd > forced:
            at_limit = True
            # 細かく見るには、バイトの数が予算に収まるまで範囲を狭める（辺の長さは面積の平方根）
            side = min(1.0, (max_t1_bytes / max(1, first[0])) ** 0.5, (MAX_VALUES / max(1, first[1])) ** 0.5) * 0.9
            notes.append(f"時間とメモリの予算の中で開くため、この範囲（{wide:,}×{tall:,} 画素）を解像度 {1 << rd} 分の 1 で開きました。"
                         f"{'元の細かさ' if not forced else f'{1 << forced} 分の 1'}で見るには、region で範囲の縦横を今の {max(side, 0.01):.2f} 倍くらいに狭めてください"
                         "（範囲を狭めれば、どの部分でも元の解像度まで開けます）")
    else:
        rd = min(reduce, max_reduce)
        kept = cost(rd)
    if kept[1] > MAX_VALUES * 2:
        raise LimitExceeded(f"JPEG 2000 の画像（{W:,} × {H:,} 画素・{st.csiz} 成分）は、分解の段が少なく、開ける大きさを超えます")
    limits.budget().charge(kept[1] * 8, "JPEG 2000 の係数")
    # 第 1 層のバイトの予算：解像度を下げきっても超えるなら、細かい段の後ろの通り道を読まない。低い解像度から順に読む
    budget = max_t1_bytes * 3 // 2 if kept[0] > max_t1_bytes else None
    stopped = False
    jobs = []
    for ti, (_t, tcs, *_r) in enumerate(tiles):
        for c, tc in enumerate(tcs):
            top = tc.nl - min(rd, tc.nl)
            rects = _comp_rects(st, c, tc, win, min(rd, tc.nl))
            for r in range(0, top + 1):
                res = tc.res[r]
                res.x0, res.y0, res.x1, res.y1 = rects[r]  # ここから先は、範囲の四角だけを合成する（第 2 層は読み終わった）
                for b in res.bands:
                    b.x0, b.y0, b.x1, b.y1 = _band_rect(b.orient, r, rects[r])
                    bw, bh = b.x1 - b.x0, b.y1 - b.y0
                    if bw > 0 and bh > 0:
                        zero = 0 if tc.coding.reversible else 0.0
                        b.rows = [[zero] * bw for _ in range(bh)]
                        for blk in b.blocks:
                            if blk.segs and blk.x0 < b.x1 and blk.x1 > b.x0 and blk.y0 < b.y1 and blk.y1 > b.y0:
                                jobs.append((r, ti, c, blk, b, tc))
    jobs.sort(key=lambda t: t[0])  # 低い解像度から（止めるときも、形は残る）
    spent = 0
    clock = time.monotonic
    skipped = 0
    for n, (_r, _ti, _c, blk, b, tc) in enumerate(jobs):
        if deadline is not None and n & 15 == 0 and clock() > deadline:
            stopped = True
            skipped = len(jobs) - n
            break
        size = _block_bytes(blk)
        if budget is not None and spent + size > budget:
            skipped += 1
            continue
        spent += size
        try:
            _block_to_band(blk, b, tc.coding.reversible, b.orient, tc.coding.style, tc.roi)
        except (IndexError, JpxError) as e:
            problems.append(f"符号ブロックを読めません（{type(e).__name__}）")
    if skipped:
        if stopped:
            problems.append(f"時間がかかりすぎるので、{skipped} 個の符号ブロックを読まずに開きました（細かい所がぼけています。region で範囲を狭めると読めます）")
        else:
            notes.append(f"予算の中で開くため、{skipped} 個の符号ブロックを読みませんでした（細かい所がぼけています。region で範囲を狭めると読めます）")
    # 出力の面
    sizes = [_window_size(st, c, rd, win) for c in range(st.csiz)]
    out_planes = [bytearray((s[2] - s[0]) * (s[3] - s[1])) for s in sizes]
    for _t, tcs, *_r in tiles:
        comps = []
        for tc in tcs:
            top = tc.nl - min(rd, tc.nl)
            rev = tc.coding.reversible
            res0 = tc.res[0]
            rows = res0.bands[0].rows or []
            for r in range(1, top + 1):
                # 1 つ下の段の四角は、合成のために足した分だけ広い：この段の低域の四角に切りそろえる
                cur, low = tc.res[r], tc.res[r - 1]
                lx0, ly0, lx1, ly1 = (cur.x0 + 1) >> 1, (cur.y0 + 1) >> 1, (cur.x1 + 1) >> 1, (cur.y1 + 1) >> 1
                if (lx0, ly0, lx1, ly1) != (low.x0, low.y0, low.x1, low.y1):
                    rows = [row[lx0 - low.x0:lx1 - low.x0] for row in rows[ly0 - low.y0:ly1 - low.y0]]
                rows = _inverse_level(rows, cur, rev)
                for b in cur.bands:
                    b.rows = None
            res0.bands[0].rows = None
            comps.append((tc, rows, top))
        _finish_tile(st, _t, comps, out_planes, sizes, rd)
        for tc, _rows, _top in comps:
            tc.res = None
    img = _finish_image(st, j2, out_planes, sizes, rd, notes, problems, stopped)
    img.at_limit = at_limit
    if window is not None:
        k = 1 << rd
        s0 = max(sizes, key=lambda s: s[2] - s[0])
        c0 = sizes.index(s0)
        xr, yr = st.comps[c0][2], st.comps[c0][3]
        img.window = (max(0, s0[0] * k * xr - st.xo), max(0, s0[1] * k * yr - st.yo),
                      min(W, s0[2] * k * xr - st.xo), min(H, s0[3] * k * yr - st.yo))
    return img


def _window_size(st: _Stream, c: int, rd: int, win: tuple) -> tuple:
    """成分 c の、解像度を rd 段減らした出力の面の四角（範囲に掛かる所）。"""
    x0, y0, x1, y1 = _output_size(st, c, rd)
    _ss, _sg, xr, yr = st.comps[c]
    k = 1 << rd
    wx0, wy0 = win[0] // xr // k, win[1] // yr // k
    wx1, wy1 = _ceildiv(_ceildiv(win[2], xr), k), _ceildiv(_ceildiv(win[3], yr), k)
    return max(x0, wx0), max(y0, wy0), max(max(x0, wx0) + 1, min(x1, wx1)), max(max(y0, wy0) + 1, min(y1, wy1))


def _finish_tile(st: _Stream, tile: _Tile, comps: list, out_planes: list, sizes: list, rd: int) -> None:
    """MCT（附属書 G）・DC のずらし・丸めて 8 ビットの面へ。"""
    cod = tile.cod or st.cod
    n = len(comps)
    rowsets = [c[1] for c in comps]
    if cod.mct and n >= 3 and all(len(rowsets[i]) == len(rowsets[0]) for i in (1, 2)):
        rev = comps[0][0].coding.reversible
        r0, r1, r2 = rowsets[0], rowsets[1], rowsets[2]
        if rows_same_width(r0, r1, r2):
            new0, new1, new2 = [], [], []
            for y0, y1, y2 in zip(r0, r1, r2):
                if rev:
                    g = [a - ((b + c) >> 2) for a, b, c in zip(y0, y1, y2)]
                    new0.append([c + v for c, v in zip(y2, g)])
                    new1.append(g)
                    new2.append([b + v for b, v in zip(y1, g)])
                else:
                    new0.append([a + 1.402 * c for a, c in zip(y0, y2)])
                    new1.append([a - 0.34413 * b - 0.71414 * c for a, b, c in zip(y0, y1, y2)])
                    new2.append([a + 1.772 * b for a, b in zip(y0, y1)])
            rowsets[0], rowsets[1], rowsets[2] = new0, new1, new2
    for c, (tc, _r, top) in enumerate(comps):
        rows = rowsets[c]
        if not rows:
            continue
        ss, signed, _xr, _yr = st.comps[c]
        res = tc.res[top]
        ox, oy, x1, y1 = sizes[c]
        stride = x1 - ox
        dst = out_planes[c]
        off = 0 if signed else 1 << (ss - 1)
        lo, hi = (-(1 << (ss - 1)), (1 << (ss - 1)) - 1) if signed else (0, (1 << ss) - 1)
        shift_show = 1 << (ss - 1) if signed else 0  # 符号つきの成分は、見せるために 0 始まりにずらす
        rev = tc.coding.reversible
        # 出力の面の四角に掛かる所だけ（範囲を開くときは、合成のために足した分が外に出る）
        sk = max(0, ox - res.x0)
        en = min(res.x1 - res.x0, x1 - res.x0)
        if en <= sk:
            continue
        lines = [(res.y0 + j, row[sk:en]) for j, row in enumerate(rows) if oy <= res.y0 + j < y1]
        if ss == 8:
            for Y, row in lines:
                if rev:
                    line = bytes([0 if v < 0 else 255 if v > 255 else v for v in [x + off + shift_show for x in row]])
                else:
                    t = off + shift_show + 0.5
                    line = bytes([0 if v < 0 else 255 if v > 255 else int(v) for v in [x + t for x in row]])
                p = (Y - oy) * stride + (res.x0 + sk - ox)
                dst[p:p + len(line)] = line
        else:
            # 8 ビットへ（16 ビットは上の桁、少ないビットは広げる）
            scale = 255.0 / ((1 << ss) - 1)
            for Y, row in lines:
                line = bytes([int((min(hi, max(lo, int(x + off + (0 if rev else 0.5)))) + shift_show) * scale + 0.5) for x in row])
                p = (Y - oy) * stride + (res.x0 + sk - ox)
                dst[p:p + len(line)] = line


def rows_same_width(*sets) -> bool:
    w = {len(s[0]) if s else 0 for s in sets}
    return len(w) == 1


def _upsample(px: bytes, sw: int, sh: int, W: int, H: int) -> bytes:
    if (sw, sh) == (W, H):
        return bytes(px)
    cols = [min(sw - 1, x * sw // W) for x in range(W)]
    rows = {}
    out = []
    for y in range(H):
        sy = min(sh - 1, y * sh // H)
        r = rows.get(sy)
        if r is None:
            base = sy * sw
            r = bytes([px[base + c] for c in cols])
            rows = {sy: r}
        out.append(r)
    return b"".join(out)


def _finish_image(st: _Stream, j2: _JP2, planes: list, sizes: list, rd: int, notes: list, problems: list, stopped: bool) -> Image:
    W = max(s[2] - s[0] for s in sizes)
    H = max(s[3] - s[1] for s in sizes)
    full = [_upsample(p, s[2] - s[0], s[3] - s[1], W, H) for p, s in zip(planes, sizes)]
    cs = j2.colorspace
    if cs is None and j2.icc_n:
        cs = {1: "gray", 3: "sRGB", 4: "CMYK"}.get(j2.icc_n)
    alpha = None
    if j2.palette and j2.cmap:
        # I.5.3.5：成分と列の対応（MTYP 0 はそのまま、1 はパレット）
        chans = []
        for cmp_, mtyp, pcol in j2.cmap:
            src = full[cmp_] if cmp_ < len(full) else None
            if src is None:
                continue
            if mtyp == 1 and pcol < len(j2.palette):
                bits = j2.palette_bits[pcol]
                col = j2.palette[pcol]
                # 添字は復号の値：8 ビットの面に写す前の値（1〜8 ビットの添字は、広げた分を戻す）
                ss = st.comps[cmp_][0]
                maxv = (1 << ss) - 1
                tab = bytearray(256)
                for v in range(256):
                    idx = v if ss >= 8 else (v * maxv + 127) // 255
                    x = col[min(idx, len(col) - 1)]
                    tab[v] = x >> (bits - 8) if bits > 8 else (x * 255 + ((1 << bits) - 1) // 2) // ((1 << bits) - 1)
                chans.append(src.translate(bytes(tab)))
            else:
                chans.append(src)
        full = chans
    if j2.cdef:
        # I.5.3.6：色の順と透け
        order: dict = {}
        for cn, typ, asoc in j2.cdef:
            if cn >= len(full):
                continue
            if typ in (1, 2) and alpha is None:
                alpha = full[cn]
            elif typ == 0 and 1 <= asoc <= len(full):
                order[asoc - 1] = full[cn]
        if order:
            full = [order[i] for i in sorted(order)]
    if cs == "sYCC" and len(full) >= 3:
        y_, cb, cr = full[:3]
        tr = [1.402 * (v - 128) for v in range(256)]
        tb = [1.772 * (v - 128) for v in range(256)]
        g1 = [-0.34413 * (v - 128) for v in range(256)]
        g2 = [-0.71414 * (v - 128) for v in range(256)]
        clip = lambda v: 0 if v < 0 else 255 if v > 255 else int(v + 0.5)  # noqa: E731
        r = bytes([clip(a + tr[c]) for a, c in zip(y_, cr)])
        g = bytes([clip(a + g1[b] + g2[c]) for a, b, c in zip(y_, cb, cr)])
        b_ = bytes([clip(a + tb[b]) for a, b in zip(y_, cb)])
        full = [r, g, b_] + full[3:]
        cs = "sRGB"
    if rd:
        notes.append(f"解像度の段を {rd} 段減らして開きました（元の {st.xsiz - st.xo}×{st.ysiz - st.yo} 画素の {1 << rd} 分の 1。開いた画像は {W}×{H} 画素）")
    return Image(W, H, full, cs, alpha, rd, st.xsiz - st.xo, st.ysiz - st.yo, notes, problems, stopped, j2.dpi, [c[0] for c in st.comps])


def info(data: bytes) -> dict:
    """見出しだけ読む：幅・高さ・成分の数。"""
    j2 = _parse_jp2(data)
    d = j2.codestream
    xs, ys, xo, yo = struct.unpack_from(">IIII", d, 8)
    return {"w": xs - xo, "h": ys - yo, "components": struct.unpack_from(">H", d, 40)[0], "dpi": j2.dpi, "colorspace": j2.colorspace}

