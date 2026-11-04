"""PDF の画像を、画素（白黒の濃さ）に開く：スキャンした紙の罫線を見つけるため、と、エージェントに画像を見せるため。

外部のライブラリを使わずに、PDF の画像の圧縮を開く（ISO 32000-1 8.9・7.4）：
- CCITTFaxDecode：ITU-T T.4（G3：1次元・2次元）と T.6（G4）。
- DCTDecode：JPEG（ITU-T T.81）のベースラインと拡張（ハフマン・順次）。濃さだけを使うので、輝度（Y）の成分だけを
  逆 DCT する（大きな画像は、係数の左上だけを使って 1/2・1/4 の大きさで開く：T.81 の DCT の定義どおりの縮小）。
- JBIG2Decode（ITU-T T.88）と JPXDecode（ISO/IEC 15444-1）：todoku.imaging の部品で開く。
- そのほか（FlateDecode・LZW など）は doc.decode が開いた画素を、色空間（Gray・RGB・CMYK・ICCBased・Indexed）と
  Decode の配列で濃さにする。

返す画像は「1行ずつ、左から、0＝黒〜255＝白」の濃さ（Gray）。
"""

from __future__ import annotations

import math
import struct
import zlib
from dataclasses import dataclass, field


class RasterError(ValueError):
    pass


@dataclass
class Gray:
    w: int
    h: int
    px: bytes  # w*h バイト（0＝黒〜255＝白）
    notes: list = field(default_factory=list)  # 開き方の知らせ（解像度を下げたなど）
    problems: list = field(default_factory=list)  # 描けなかった所

    def row(self, y: int) -> bytes:
        return self.px[y * self.w:(y + 1) * self.w]


# ====================================================================== CCITT（T.4・T.6）
def _table(spec: str) -> dict[str, int]:
    out = {}
    for item in spec.split():
        n, code = item.split(":")
        out[code] = int(n)
    return out


_WT = _table("""0:00110101 1:000111 2:0111 3:1000 4:1011 5:1100 6:1110 7:1111 8:10011 9:10100 10:00111 11:01000 12:001000
13:000011 14:110100 15:110101 16:101010 17:101011 18:0100111 19:0001100 20:0001000 21:0010111 22:0000011 23:0000100
24:0101000 25:0101011 26:0010011 27:0100100 28:0011000 29:00000010 30:00000011 31:00011010 32:00011011 33:00010010
34:00010011 35:00010100 36:00010101 37:00010110 38:00010111 39:00101000 40:00101001 41:00101010 42:00101011 43:00101100
44:00101101 45:00000100 46:00000101 47:00001010 48:00001011 49:01010010 50:01010011 51:01010100 52:01010101 53:00100100
54:00100101 55:01011000 56:01011001 57:01011010 58:01011011 59:01001010 60:01001011 61:00110010 62:00110011 63:00110100""")
_WM = _table("""64:11011 128:10010 192:010111 256:0110111 320:00110110 384:00110111 448:01100100 512:01100101 576:01101000
640:01100111 704:011001100 768:011001101 832:011010010 896:011010011 960:011010100 1024:011010101 1088:011010110
1152:011010111 1216:011011000 1280:011011001 1344:011011010 1408:011011011 1472:010011000 1536:010011001 1600:010011010
1664:011000 1728:010011011""")
_BT = _table("""0:0000110111 1:010 2:11 3:10 4:011 5:0011 6:0010 7:00011 8:000101 9:000100 10:0000100 11:0000101 12:0000111
13:00000100 14:00000111 15:000011000 16:0000010111 17:0000011000 18:0000001000 19:00001100111 20:00001101000
21:00001101100 22:00000110111 23:00000101000 24:00000010111 25:00000011000 26:000011001010 27:000011001011
28:000011001100 29:000011001101 30:000001101000 31:000001101001 32:000001101010 33:000001101011 34:000011010010
35:000011010011 36:000011010100 37:000011010101 38:000011010110 39:000011010111 40:000001101100 41:000001101101
42:000011011010 43:000011011011 44:000001010100 45:000001010101 46:000001010110 47:000001010111 48:000001100100
49:000001100101 50:000001010010 51:000001010011 52:000000100100 53:000000110111 54:000000111000 55:000000100111
56:000000101000 57:000001011000 58:000001011001 59:000000101011 60:000000101100 61:000001011010 62:000001100110
63:000001100111""")
_BM = _table("""64:0000001111 128:000011001000 192:000011001001 256:000001011011 320:000000110011 384:000000110100
448:000000110101 512:0000001101100 576:0000001101101 640:0000001001010 704:0000001001011 768:0000001001100
832:0000001001101 896:0000001110010 960:0000001110011 1024:0000001110100 1088:0000001110101 1152:0000001110110
1216:0000001110111 1280:0000001010010 1344:0000001010011 1408:0000001010100 1472:0000001010101 1536:0000001011010
1600:0000001011011 1664:0000001100100 1728:0000001100101""")
_XM = _table("""1792:00000001000 1856:00000001100 1920:00000001101 1984:000000010010 2048:000000010011 2112:000000010100
2176:000000010101 2240:000000010110 2304:000000010111 2368:000000011100 2432:000000011101 2496:000000011110
2560:000000011111""")
_MODES = {"0001": "P", "001": "H", "1": "V0", "011": "VR1", "000011": "VR2", "0000011": "VR3", "010": "VL1",
          "000010": "VL2", "0000010": "VL3"}
_VDELTA = {"V0": 0, "VR1": 1, "VR2": 2, "VR3": 3, "VL1": -1, "VL2": -2, "VL3": -3}
_EOL = "000000000001"


def _lookup(codes: dict[str, int]) -> dict[tuple[int, int], int]:
    """(符号の長さ, 符号の値) → 値。符号が互いに頭を共有しないこと（語頭符号）を確かめてから作る。"""
    ks = sorted(codes, key=len)
    for i, a in enumerate(ks):
        for b in ks[i + 1:]:
            if b.startswith(a):
                raise AssertionError(f"符号 {a} が {b} の頭になっています")
    return {(len(k), int(k, 2)): v for k, v in codes.items()}


_WHITE = _lookup({**_WT, **_WM, **_XM})
_BLACK = _lookup({**_BT, **_BM, **_XM})
_MODE = _lookup(_MODES)


class _Bits:
    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0
        self.n = len(data) * 8

    def peek(self, k: int) -> int:
        i = self.pos >> 3
        chunk = int.from_bytes(self.data[i:i + 4].ljust(4, b"\0"), "big")
        return (chunk >> (32 - (self.pos & 7) - k)) & ((1 << k) - 1)

    def skip(self, k: int) -> None:
        self.pos += k

    def code(self, table: dict, maxlen: int = 13) -> int | None:
        for ln in range(1, maxlen + 1):
            v = table.get((ln, self.peek(ln)))
            if v is not None:
                self.pos += ln
                return v
        return None

    def align(self) -> None:
        self.pos = (self.pos + 7) & ~7

    def at_eol(self) -> bool:
        return self.peek(12) == 1


def _run(bits: _Bits, black: bool) -> int:
    table = _BLACK if black else _WHITE
    total = 0
    while True:
        v = bits.code(table)
        if v is None:
            raise RasterError(f"CCITT：{'黒' if black else '白'}の長さの符号が読めません（{bits.pos} ビット目）")
        total += v
        if v < 64:
            return total


def ccitt(data: bytes, parms: dict) -> Gray:
    """CCITTFaxDecode（ISO 32000-1 表 11 の DecodeParms）。"""
    k = int(parms.get("K", 0))
    cols = int(parms.get("Columns", 1728))
    rows = int(parms.get("Rows", 0))
    align = bool(parms.get("EncodedByteAlign", False))
    black1 = bool(parms.get("BlackIs1", False))
    from ..limits import MAX_PIXELS

    if cols <= 0 or rows < 0 or cols * max(rows, 1) > MAX_PIXELS:
        raise RasterError(f"CCITT：画像が大きすぎます（{cols:,} × {rows:,}。開ける大きさは {MAX_PIXELS:,} 画素まで）")
    bits = _Bits(data)
    out = bytearray()
    ref = [cols, cols]  # 上の行の変わり目（白の行）
    y = 0
    while (rows == 0 or y < rows) and bits.pos < bits.n:
        if k >= 0:
            # G3：行の頭の EOL（と、2次元のときの 1次元／2次元の印）
            if align:
                while bits.pos < bits.n and bits.peek(12) != 1 and (bits.pos & 7) == 0 and bits.peek(8) == 0:
                    bits.skip(8)
            while bits.pos < bits.n and bits.peek(12) == 1:
                bits.skip(12)
                if bits.peek(12) == 1:  # RTC（EOL の続き）：終わり
                    bits.pos = bits.n
                    break
            if bits.pos >= bits.n:
                break
            two_d = False
            if k > 0:
                two_d = bits.peek(1) == 0
                bits.skip(1)
            cur = _g4_line(bits, ref, cols) if two_d else _g3_line(bits, cols)
        else:
            if align:
                bits.align()
            if bits.peek(24) == 0x001001:  # EOFB（EOL を2つ）：終わり
                break
            cur = _g4_line(bits, ref, cols)
        if (y + 1) * cols > MAX_PIXELS:
            raise RasterError(f"CCITT：画像が大きすぎます（{cols:,} 列で {y + 1:,} 行を超える。開ける大きさは {MAX_PIXELS:,} 画素まで）")
        row = bytearray(b"\xff" * cols)  # 白
        for i in range(0, len(cur) - 1, 2):
            a, b = min(cur[i], cols), min(cur[i + 1], cols)
            if b > a:
                row[a:b] = b"\x00" * (b - a)
        out += row
        ref = cur + [cols, cols]
        y += 1
        if k >= 0 and align:
            bits.align()
    if rows and y < rows:
        out += b"\xff" * cols * (rows - y)
        y = rows
    px = bytes(out)
    if black1:
        px = px.translate(bytes(255 - i for i in range(256)))
    return Gray(cols, y, px)


def _g3_line(bits: _Bits, cols: int) -> list[int]:
    cur: list[int] = []
    a = 0
    black = False
    while a < cols:
        a += _run(bits, black)
        cur.append(min(a, cols))
        black = not black
    return cur if len(cur) % 2 == 0 else cur + [cols]  # 変わり目（白→黒、黒→白）の組


def _g4_line(bits: _Bits, ref: list[int], cols: int) -> list[int]:
    """T.6 の2次元の符号で1行（変わり目の位置の並び：偶数番目が白→黒、奇数番目が黒→白）。"""
    cur: list[int] = []
    a0 = -1
    black = False
    ri = 0  # ref の探し始め
    while a0 < cols:
        # b1：a0 より右で、a0 の色と逆の色に変わる、上の行の最初の変わり目
        while ri > 0 and ref[ri - 1] > a0:
            ri -= 1
        while ref[ri] <= a0 or (ri % 2 == 1) != black:
            if ref[ri] <= a0:
                ri += 1
            else:
                ri += 1
            if ri >= len(ref) - 1:
                ri = len(ref) - 2
                break
        b1 = ref[ri]
        b2 = ref[ri + 1] if ri + 1 < len(ref) else cols
        mode = bits.code(_MODE, 7)
        if mode is None:
            raise RasterError(f"CCITT：2次元の符号が読めません（{bits.pos} ビット目）")
        if mode == "P":
            a0 = b2
        elif mode == "H":
            start = max(a0, 0)
            r1 = _run(bits, black)
            r2 = _run(bits, not black)
            a1 = start + r1
            a2 = a1 + r2
            cur += [min(a1, cols), min(a2, cols)]
            a0 = a2
        else:
            a1 = b1 + _VDELTA[mode]
            if a1 < 0 or (a1 < a0 and a0 >= 0):
                raise RasterError(f"CCITT：変わり目が前に戻ります（{bits.pos} ビット目）")
            cur.append(min(a1, cols))
            a0 = a1
            black = not black
    # 変わり目の組にする（終わりで黒のままなら、行の終わりで白に戻す）
    if len(cur) % 2 == 1:
        cur.append(cols)
    return cur


# ====================================================================== JPEG（T.81：ベースライン・拡張の順次）
_ZZ = [0, 1, 8, 16, 9, 2, 3, 10, 17, 24, 32, 25, 18, 11, 4, 5, 12, 19, 26, 33, 40, 48, 41, 34, 27, 20, 13, 6, 7, 14, 21,
       28, 35, 42, 49, 56, 57, 50, 43, 36, 29, 22, 15, 23, 30, 37, 44, 51, 58, 59, 52, 45, 38, 31, 39, 46, 53, 60, 61, 54, 47,
       55, 62, 63]


def _idct_matrix(n: int) -> list[list[float]]:
    """n×n に縮めた逆 DCT（T.81 A.3.3 の式を、係数の左上 n×n だけで、n 点で計算する）：M[x][u]。"""
    m = []
    for x in range(n):
        row = []
        for u in range(n):
            cu = math.sqrt(0.5) if u == 0 else 1.0
            row.append(cu * math.cos((2 * x + 1) * u * math.pi / (2 * n)) / 2)
        m.append(row)
    return m


class _Huff:
    """ハフマン表（T.81 C）：先の 16 ビットで引く表（値は 符号の長さ×256＋記号。0 は符号でない）。"""

    def __init__(self, counts: bytes, symbols: bytes):
        self.lut = [0] * 65536
        code = 0
        k = 0
        for ln in range(1, 17):
            for _ in range(counts[ln - 1]):
                lo = code << (16 - ln)
                hi = lo + (1 << (16 - ln))
                v = (ln << 8) | symbols[k]
                self.lut[lo:hi] = [v] * (hi - lo)
                k += 1
                code += 1
            code <<= 1


def _segments(data: bytes, start: int) -> list[bytes]:
    """走査の圧縮データを、リスタートマーカー（RSTn）で分け、詰め物（0xFF 0x00 → 0xFF）を取り除く。"""
    segs = []
    i = cur = start
    n = len(data)
    while True:
        j = data.find(b"\xff", i)
        if j < 0 or j + 1 >= n:
            segs.append(data[cur:].replace(b"\xff\x00", b"\xff"))
            return segs
        m = data[j + 1]
        if m == 0x00:
            i = j + 2
            continue
        if m == 0xFF:  # 詰めの 0xFF（B.1.1.2）
            i = j + 1
            continue
        segs.append(data[cur:j].replace(b"\xff\x00", b"\xff"))
        if 0xD0 <= m <= 0xD7:
            cur = i = j + 2
            continue
        return segs


def jpeg_window(window, ww: int, hh: int, hmax: int, vmax: int) -> tuple:
    """範囲（元の画素）を MCU の境に広げた [左, 上, 右, 下)。"""
    ax, ay = 8 * hmax, 8 * vmax
    x0, y0, x1, y1 = (int(v) for v in window)
    x1, y1 = min(ww, max(x1, x0 + 1)), min(hh, max(y1, y0 + 1))
    x0, y0 = max(0, min(x0, ww - 1)), max(0, min(y0, hh - 1))
    return x0 // ax * ax, y0 // ay * ay, min(ww, (x1 + ax - 1) // ax * ax), min(hh, (y1 + ay - 1) // ay * ay)


def jpeg_blocks(win: tuple, c: dict, hmax: int, vmax: int) -> tuple:
    """範囲を描くのに要る、成分 c のブロックの [左, 上, 右, 下)（なめらかに広げるので 1 つ余分に）。"""
    return (max(0, win[0] * c["h"] // (8 * hmax) - 1), max(0, win[1] * c["v"] // (8 * vmax) - 1),
            (win[2] * c["h"] + 8 * hmax - 1) // (8 * hmax) + 1, (win[3] * c["v"] + 8 * vmax - 1) // (8 * vmax) + 1)


def jpeg_gray(data: bytes, max_pixels: int = 4_000_000, want_rgb: bool = False, window: tuple | None = None):
    """JPEG の輝度（1成分のときはその成分）を濃さにする。大きな画像は 1/2・1/4・1/8 に縮めて開く。
    want_rgb：色のまま開く（RGB を返す。YCbCr・RGB・CMYK・YCCK を T.81・Adobe の印どおりに RGB にする）。
    window：want_rgb のとき、要る範囲（元の画素）。ハフマンの符号は順にしか読めないが、範囲の下で読むのをやめ、逆 DCT と色の変換は範囲だけ行う（RGB.window に開いた範囲）。"""
    if data[:2] != b"\xff\xd8":
        raise RasterError("JPEG ではありません（SOI がない）")
    pos = 2
    qt: dict[int, list[int]] = {}
    hts: dict[tuple[int, int], _Huff] = {}
    frame = None
    restart = 0
    adobe = None
    while pos < len(data):
        if data[pos] != 0xFF:
            pos += 1
            continue
        m = data[pos + 1]
        pos += 2
        if m in (0xD8, 0x01) or 0xD0 <= m <= 0xD7:
            continue
        if m == 0xD9:
            break
        ln = struct.unpack_from(">H", data, pos)[0]
        seg = data[pos + 2:pos + ln]
        if m == 0xDB:
            i = 0
            while i < len(seg):
                pq, tq = seg[i] >> 4, seg[i] & 15
                i += 1
                if pq:
                    vals = list(struct.unpack_from(">64H", seg, i))
                    i += 128
                else:
                    vals = list(seg[i:i + 64])
                    i += 64
                qt[tq] = vals
        elif m == 0xC4:
            i = 0
            while i < len(seg):
                tc, th = seg[i] >> 4, seg[i] & 15
                counts = seg[i + 1:i + 17]
                n = sum(counts)
                hts[(tc, th)] = _Huff(counts, seg[i + 17:i + 17 + n])
                i += 17 + n
        elif m in (0xC0, 0xC1):
            p, hh, ww, nf = seg[0], *struct.unpack_from(">HH", seg, 1), seg[5]
            if p != 8:
                from ..imaging.jpegprog import decode as multi_scan

                return multi_scan(data, max_pixels, want_rgb, window)  # 12 ビット（T.81 附属書 F）
            comps = []
            for j in range(nf):
                cid, hv, tq = seg[6 + 3 * j], seg[7 + 3 * j], seg[8 + 3 * j]
                comps.append({"id": cid, "h": hv >> 4, "v": hv & 15, "tq": tq})
            frame = (ww, hh, comps)
        elif m in (0xC2, 0xC3, 0xC9, 0xCA):
            from ..imaging.jpegprog import decode as multi_scan

            return multi_scan(data, max_pixels, want_rgb, window)  # プログレッシブ（附属書 G）・可逆（附属書 H）・算術符号（附属書 D）
        elif 0xC2 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC):
            raise RasterError(f"JPEG：この符号化（SOF{m - 0xC0}）は開けません（ベースライン・拡張の順次・プログレッシブ・可逆・算術符号だけ）")
        elif m == 0xDD:
            restart = struct.unpack_from(">H", seg, 0)[0]
        elif m == 0xEE and seg[:5] == b"Adobe":
            adobe = seg[11] if len(seg) > 11 else 0
        elif m == 0xDA:
            if frame is None:
                raise RasterError("JPEG：フレームの前に走査があります")
            return _scan(data, pos + ln, seg, frame, qt, hts, restart, adobe, max_pixels, want_rgb, window)
        pos += ln
    raise RasterError("JPEG：画像のデータがありません")


def _scan(data, start, seg, frame, qt, hts, restart, adobe, max_pixels, want_rgb=False, window=None):
    ww, hh, comps = frame
    ns = seg[0]
    sc = []
    for j in range(ns):
        cid, t = seg[1 + 2 * j], seg[2 + 2 * j]
        c = next(x for x in comps if x["id"] == cid)
        sc.append((c, hts.get((0, t >> 4)), hts.get((1, t & 15))))
    if ns != len(comps):
        from ..imaging.jpegprog import decode as multi_scan

        return multi_scan(data, max_pixels, want_rgb, window)  # 成分を分けた走査（順次）
    hmax = max(c["h"] for c in comps)
    vmax = max(c["v"] for c in comps)
    # 濃さに使う成分：1成分ならそれ、YCbCr（3成分・Adobe の変換 1 か、印がない）なら Y。RGB・CMYK は全部の成分から
    use_y = (len(comps) == 1 or (len(comps) == 3 and adobe != 0) or (len(comps) == 4 and adobe == 2)) and not want_rgb
    n = 8
    while n > 1 and (ww * n // 8) * (hh * n // 8) > max_pixels:
        n //= 2
    mat = _idct_matrix(n)
    mcux = (ww + 8 * hmax - 1) // (8 * hmax)
    mcuy = (hh + 8 * vmax - 1) // (8 * vmax)
    from ..limits import MAX_PIXELS

    blocks = mcux * mcuy * sum(c["h"] * c["v"] for c in comps)
    if (ww * n // 8) * (hh * n // 8) > MAX_PIXELS or blocks * 2 > (len(data) - start) * 8:
        # 1 つのブロックは少なくとも 2 ビット（DC の差の符号と、AC の終わりの符号）：データの足りない大きさは、書き換えた見出し
        raise RasterError(f"JPEG：画像の大きさ（{ww:,} × {hh:,}）に対して、データが足りないか、大きすぎます")
    planes = []
    for c in comps:
        bw, bh = mcux * c["h"], mcuy * c["v"]
        planes.append((c, bytearray(b"\x80" * (bw * n) * (bh * n)), bw * n))
    win = jpeg_window(window, ww, hh, hmax, vmax) if window is not None and want_rgb else None
    area = [jpeg_blocks(win, c, hmax, vmax) for c in comps] if win else None
    last_row = (win[3] + 8 * vmax - 1) // (8 * vmax) + 1 if win else mcuy  # この行より下は読まない（なめらかに広げるので 1 行余分に）
    segs = _segments(data, start)
    total = mcux * mcuy
    q = {id(c): qt.get(c["tq"]) for c in comps}
    zz = _ZZ
    per = restart or total
    mi = 0
    for buf in segs:
        if mi >= total or mi >= last_row * mcux:
            break
        buf = buf + b"\0" * 8
        pos = 0  # ビットの位置
        pred = [0] * len(comps)
        for _ in range(min(per, total - mi)):
            my, mx = divmod(mi, mcux)
            mi += 1
            if my >= last_row:
                break
            for ci, (c, dc_t, ac_t) in enumerate(sc):
                plane, stride = planes[ci][1], planes[ci][2]
                qv = q[id(c)]
                dcl, acl = dc_t.lut, ac_t.lut
                skip0 = use_y and ci > 0
                ar = area[ci] if area else None
                for by in range(c["v"]):
                    for bx in range(c["h"]):
                        skip = skip0
                        if ar and not (ar[0] <= mx * c["h"] + bx < ar[2] and ar[1] <= my * c["v"] + by < ar[3]):
                            skip = True
                        i = pos >> 3
                        e = dcl[(int.from_bytes(buf[i:i + 4], "big") << (pos & 7) & 0xFFFFFFFF) >> 16]
                        if not e:
                            raise RasterError("JPEG：ハフマン符号が読めません")
                        t = e & 255
                        pos += e >> 8
                        if t:
                            i = pos >> 3
                            v = (int.from_bytes(buf[i:i + 4], "big") >> (32 - (pos & 7) - t)) & ((1 << t) - 1)
                            pos += t
                            pred[ci] += v - (1 << t) + 1 if v < (1 << (t - 1)) else v
                        coef = [0] * 64
                        coef[0] = pred[ci] * qv[0]
                        k = 1
                        while k < 64:
                            i = pos >> 3
                            e = acl[(int.from_bytes(buf[i:i + 4], "big") << (pos & 7) & 0xFFFFFFFF) >> 16]
                            if not e:
                                raise RasterError("JPEG：ハフマン符号が読めません")
                            pos += e >> 8
                            sz = e & 15
                            if sz == 0:
                                if (e & 255) >> 4 == 15:
                                    k += 16
                                    continue
                                break
                            k += (e & 255) >> 4
                            if k > 63:
                                break
                            i = pos >> 3
                            v = (int.from_bytes(buf[i:i + 4], "big") >> (32 - (pos & 7) - sz)) & ((1 << sz) - 1)
                            pos += sz
                            if not skip:
                                coef[zz[k]] = (v - (1 << sz) + 1 if v < (1 << (sz - 1)) else v) * qv[k]
                            k += 1
                        if not skip:
                            _put_block(coef, mat, n, plane, stride, (mx * c["h"] + bx) * n, (my * c["v"] + by) * n)
    # 濃さにする（Y は成分の大きさをフレームの大きさに合わせる）
    W, H = ww * n // 8 or 1, hh * n // 8 or 1
    if want_rgb:
        return _jpeg_rgb(planes, hmax, vmax, W, H, adobe, win, n)
    if use_y:
        c, plane, stride = planes[0]
        fx, fy = hmax // c["h"], vmax // c["v"]
        if fx == 1 and fy == 1:
            px = b"".join(bytes(plane[y * stride:y * stride + W]) for y in range(H))
        else:
            px = b"".join(bytes(plane[(y // fy) * stride + x // fx] for x in range(W)) for y in range(H))
        return Gray(W, H, px)
    rows = []
    for y in range(H):
        vals = []
        for (c, plane, stride) in planes:
            fx, fy = hmax // c["h"], vmax // c["v"]
            vals.append([plane[(y // fy) * stride + x // fx] for x in range(W)])
        if len(planes) == 3:  # RGB（Adobe の変換 0）
            rows.append(bytes((r * 299 + g * 587 + b * 114) // 1000 for r, g, b in zip(*vals)))
        else:  # CMYK（Adobe の変換 0）・YCCK（変換 2 は上で Y にした）。Adobe の CMYK は反転して保存される
            inv = adobe is not None
            rows.append(bytes(max(0, min(255, (k if inv else 255 - k) - (255 - (cc if inv else 255 - cc)) // 3))
                              for cc, _m, _y, k in zip(*vals)))
    return Gray(W, H, b"".join(rows))


@dataclass
class RGB:
    w: int
    h: int
    px: bytes  # w*h*3 バイト（R G B）
    window: tuple | None = None  # 開いた範囲（元の画素の [左, 上, 右, 下)）。None は全体
    notes: list = field(default_factory=list)


def _smooth_plane(plane, stride: int, fx: int, fy: int, W: int, H: int, x0: int, y0: int, x1: int, y1: int) -> bytes:
    """間引いた色の面を、なめらかに広げる（三角のフィルター。libjpeg の fancy upsampling と同じ重み）。[x0, x1)×[y0, y1) を返す。
    2:1 の向きだけを広げ、面の幅は ceil(W / fx)・高さは ceil(H / fy) とする（ブロックの詰め物の画素を混ぜない）。"""
    cw, ch = (W + fx - 1) // fx, (H + fy - 1) // fy
    c0, c1 = max(0, x0 // fx - 1), min(cw, (x1 + fx - 1) // fx + 1)
    skip = x0 - c0 * fx
    out = []
    cache: dict = {}

    def row(j: int) -> bytes:
        r = cache.get(j)
        if r is None:
            r = bytes(plane[j * stride + c0:j * stride + c1])
            cache[j] = r
        return r

    n = x1 - x0
    for y in range(y0, y1):
        j = y // fy
        if fy == 2:
            nb = row(max(0, j - 1)) if not y & 1 else row(min(ch - 1, j + 1))
            sums = [3 * a + b for a, b in zip(row(j), nb)]
        else:
            sums = list(row(j))
        if fx == 2 and fy == 2:
            ev = [(3 * c + l_ + 8) >> 4 for c, l_ in zip(sums, sums[:1] + sums[:-1])]
            od = [(3 * c + r_ + 7) >> 4 for c, r_ in zip(sums, sums[1:] + sums[-1:])]
        elif fx == 2:
            ev = [(3 * c + l_ + 1) >> 2 for c, l_ in zip(sums, sums[:1] + sums[:-1])]
            od = [(3 * c + r_ + 2) >> 2 for c, r_ in zip(sums, sums[1:] + sums[-1:])]
        else:  # fy == 2 だけ
            bias = 2 if y & 1 else 1
            out.append(bytes([(v + bias) >> 2 for v in sums])[skip:skip + n])
            continue
        line = bytearray(2 * len(sums))
        line[0::2], line[1::2] = bytes(ev), bytes(od)
        out.append(bytes(line[skip:skip + n]))
        if len(cache) > 4:
            for k in [k for k in cache if k < j - 1]:
                del cache[k]
    return b"".join(out)


def _jpeg_rgb(planes, hmax, vmax, W, H, adobe, win=None, n: int = 8) -> RGB:
    """JPEG の成分 → RGB（T.81 には色の決まりがないので、JFIF の YCbCr と Adobe の APP14 の変換の印に従う。印がなく成分の番号が
    R・G・B なら RGB）。win：範囲（元の画素。n は 8 分のいくつに縮めたか）。"""
    FW, FH = W, H
    x0, y0, x1, y1 = 0, 0, W, H
    if win:
        x0, y0 = win[0] * n // 8, win[1] * n // 8
        x1, y1 = max(x0 + 1, min(W, win[2] * n // 8)), max(y0 + 1, min(H, win[3] * n // 8))
        W, H = x1 - x0, y1 - y0
    full = []
    for (c, plane, stride) in planes:
        fx, fy = hmax // c["h"], vmax // c["v"]
        if fx == 1 and fy == 1:
            full.append(b"".join(bytes(plane[y * stride + x0:y * stride + x1]) for y in range(y0, y1)))
        elif (fx, fy) in ((2, 2), (2, 1), (1, 2)) and n == 8:
            full.append(_smooth_plane(plane, stride, fx, fy, FW, FH, x0, y0, x1, y1))
        else:
            cols = [x // fx for x in range(x0, x1)]
            full.append(b"".join(bytes(map(plane.__getitem__, [(y // fy) * stride + c_ for c_ in cols])) for y in range(y0, y1)))
    nc = len(full)
    done = tuple(win) if win else None
    if nc == 1:
        g = full[0]
        out = bytearray(W * H * 3)
        out[0::3] = g
        out[1::3] = g
        out[2::3] = g
        return RGB(W, H, bytes(out), done)
    rgb_ids = adobe is None and nc == 3 and [p[0]["id"] for p in planes] == [82, 71, 66]
    ycc = (nc == 3 and adobe != 0 and not rgb_ids) or (nc == 4 and adobe == 2)
    out = bytearray(W * H * 3)
    if ycc:
        rcr = [int(1.402 * (v - 128)) for v in range(256)]
        gcb = [int(-0.344136 * (v - 128) * 65536) for v in range(256)]
        gcr = [int(-0.714136 * (v - 128) * 65536) for v in range(256)]
        bcb = [int(1.772 * (v - 128)) for v in range(256)]
        Y, Cb, Cr = full[0], full[1], full[2]

        def cl(v):
            return 0 if v < 0 else 255 if v > 255 else v
        rr = bytes(cl(y + rcr[cr]) for y, cr in zip(Y, Cr))
        gg = bytes(cl(y + ((gcb[cb] + gcr[cr]) >> 16)) for y, cb, cr in zip(Y, Cb, Cr))
        bb = bytes(cl(y + bcb[cb]) for y, cb in zip(Y, Cb))
        if nc == 3:
            out[0::3], out[1::3], out[2::3] = rr, gg, bb
            return RGB(W, H, bytes(out), done)
        # YCCK：YCC は反転した C M Y（Adobe）
        comps = [bytes(255 - v for v in rr), bytes(255 - v for v in gg), bytes(255 - v for v in bb), full[3]]
        inv = True
    else:
        if nc == 3:
            out[0::3], out[1::3], out[2::3] = full[0], full[1], full[2]
            return RGB(W, H, bytes(out), done)
        comps = full[:4]
        inv = adobe is not None
    # CMYK → RGB（Adobe の CMYK は、反転して保存される）
    if inv:
        cc, mm, yy, kk = (bytes(255 - v for v in p) for p in comps)
    else:
        cc, mm, yy, kk = comps
    out[0::3] = bytes((255 - c) * (255 - k) // 255 for c, k in zip(cc, kk))
    out[1::3] = bytes((255 - m) * (255 - k) // 255 for m, k in zip(mm, kk))
    out[2::3] = bytes((255 - y) * (255 - k) // 255 for y, k in zip(yy, kk))
    return RGB(W, H, bytes(out), done)


def _put_block(coef, mat, n, plane, stride, ox, oy) -> None:
    if not any(coef[1:]):
        v = coef[0] / 8 + 128
        b = 0 if v < 0 else 255 if v > 255 else int(v + 0.5)
        line = bytes([b]) * n
        for y in range(n):
            o = (oy + y) * stride + ox
            plane[o:o + n] = line
        return
    # 縮めた逆 DCT：係数の左上 n×n（f(x,y) = Σu Σv C(u)C(v)/4 F(u,v) cos… を n 点で。直流だけなら F/8 で、元と同じ濃さ）
    tmp = [[sum(mat[x][u] * coef[v * 8 + u] for u in range(n)) for x in range(n)] for v in range(n)]
    for y in range(n):
        o = (oy + y) * stride + ox
        row = bytearray(n)
        for x in range(n):
            s = sum(mat[y][v] * tmp[v][x] for v in range(n)) + 128
            row[x] = 0 if s < 0 else 255 if s > 255 else int(s + 0.5)
        plane[o:o + n] = row


# ====================================================================== PNG（エージェントに見せる画像）
def png(g: Gray) -> bytes:
    def chunk(t: bytes, b: bytes) -> bytes:
        return struct.pack(">I", len(b)) + t + b + struct.pack(">I", zlib.crc32(t + b) & 0xFFFFFFFF)

    raw = b"".join(b"\0" + g.row(y) for y in range(g.h))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", g.w, g.h, 8, 0, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 6)) + chunk(b"IEND", b""))


# ====================================================================== PDF の画像 → 濃さ
_FILTER_ABBR = {"AHx": "ASCIIHexDecode", "A85": "ASCII85Decode", "LZW": "LZWDecode", "Fl": "FlateDecode",
                "RL": "RunLengthDecode", "DCT": "DCTDecode", "CCF": "CCITTFaxDecode"}
_KEY_ABBR = {"W": "Width", "H": "Height", "BPC": "BitsPerComponent", "CS": "ColorSpace", "F": "Filter", "DP": "DecodeParms",
             "IM": "ImageMask", "D": "Decode"}
_CS_ABBR = {"G": "DeviceGray", "RGB": "DeviceRGB", "CMYK": "DeviceCMYK", "I": "Indexed"}


def jbig2_gray(doc, d: dict, data: bytes, names: list, deadline: float | None = None) -> Gray:
    """JBIG2Decode（ISO 32000-1 7.4.7）：埋め込みの編成。DecodeParms の JBIG2Globals を合わせて開く。1 が黒（0＝黒の濃さにして返す）。"""
    from ..imaging.jbig2 import Jbig2Error
    from ..imaging.pdfcodec import jbig2
    from .objects import Stream

    parms = doc.resolve(d.get("DecodeParms"))
    if isinstance(parms, list):
        i = names.index("JBIG2Decode")
        parms = doc.resolve(parms[i]) if i < len(parms) else None
    glob = doc.resolve(parms.get("JBIG2Globals")) if isinstance(parms, dict) else None
    gdata = doc.decode(glob) if isinstance(glob, Stream) else None
    try:
        page = jbig2(data, gdata, int(d.get("Width") or 0) or None, int(d.get("Height") or 0) or None, deadline,
                     cache=getattr(doc, "decryptor", None) is None)
    except Jbig2Error as e:
        raise RasterError(f"JBIG2 の画像を開けません：{e}") from None
    except (IndexError, struct.error) as e:
        raise RasterError(f"JBIG2 の画像を開けません（データが壊れています：{type(e).__name__}）") from None
    return Gray(page.w, page.h, page.px, [], [f"JBIG2：{x}" for x in page.problems])


def jpx_image(doc, data: bytes, target_pixels: int | None, deadline: float | None = None, window: tuple | None = None):
    """JPXDecode（ISO 32000-1 7.4.9）。imaging.jpx.Image を返す。window：要る範囲（元の画素）。"""
    from ..imaging.jpx import JpxError
    from ..imaging.pdfcodec import jpx

    try:
        return jpx(data, target_pixels, deadline, cache=getattr(doc, "decryptor", None) is None, window=window)
    except JpxError as e:
        raise RasterError(f"JPEG 2000 の画像を開けません：{e}") from None
    except (IndexError, struct.error) as e:
        raise RasterError(f"JPEG 2000 の画像を開けません（データが壊れています：{type(e).__name__}）") from None


def image_gray(doc, image, max_pixels: int = 4_000_000, deadline: float | None = None) -> Gray:
    """PDF の画像（XObject の Stream か、埋め込みの画像の (辞書, データ)）を濃さにする。開けない圧縮は RasterError。"""
    from .objects import Name, Stream

    if isinstance(image, Stream):
        d = {k: doc.resolve(v) for k, v in image.dict.items()}
        data = doc.decode(image)  # 画像の圧縮（DCT・CCITT）の手前まで開く
    elif isinstance(image, (tuple, list)) and len(image) == 2:
        d = {_KEY_ABBR.get(k, k): doc.resolve(v) for k, v in image[0].items()}
        f = d.get("Filter")
        fl = f if isinstance(f, list) else [f] if f is not None else []
        fl = [Name(_FILTER_ABBR.get(x.name, x.name)) if isinstance(x, Name) else x for x in fl]
        d["Filter"] = fl
        data = doc.decode(Stream({"Filter": fl, "DecodeParms": d.get("DecodeParms")}, image[1]), decrypt=False) if fl else image[1]
    else:
        raise RasterError("画像のデータがありません")
    flt = d.get("Filter")
    names = [getattr(x, "name", "") for x in (flt if isinstance(flt, list) else [flt] if flt is not None else [])]
    parms = d.get("DecodeParms")
    w, h = int(d.get("Width") or 0), int(d.get("Height") or 0)
    if "JBIG2Decode" in names:
        return _decode_array(jbig2_gray(doc, d, data, names, deadline), d, 1)
    if "JPXDecode" in names:
        im = jpx_image(doc, data, max_pixels, deadline)
        return Gray(im.w, im.h, _jpx_gray(im), list(im.notes), list(im.problems))
    if "CCITTFaxDecode" in names:
        p = parms[names.index("CCITTFaxDecode")] if isinstance(parms, list) else parms
        p = {k: doc.resolve(v) for k, v in (doc.resolve(p) or {}).items()}
        p.setdefault("Columns", w or 1728)
        p.setdefault("Rows", h)
        g = ccitt(data, p)
        return _decode_array(g, d, 1)
    if "DCTDecode" in names:
        return jpeg_gray(data, max_pixels)
    return _raw_gray(doc, data, d, w, h, max_pixels)


def _decode_array(g: Gray, d: dict, bpc: int) -> Gray:
    """Decode [1 0]（反転）と ImageMask（0 の所を塗る）。CCITT の画像は 1 ビット。"""
    dec = d.get("Decode")
    inv = isinstance(dec, list) and len(dec) >= 2 and float(dec[0]) > float(dec[1])
    if d.get("ImageMask") is True:
        inv = not inv  # 型抜き：0（既定）の所を、塗りの色（黒とみなす）で塗る
        g = Gray(g.w, g.h, g.px.translate(bytes(255 - i for i in range(256))), g.notes, g.problems)
    if inv:
        g = Gray(g.w, g.h, g.px.translate(bytes(255 - i for i in range(256))), g.notes, g.problems)
    return g


def _jpx_gray(im) -> bytes:
    """JPEG 2000 の画像の濃さ（SMask の用途：1 成分はそのまま。色は輝度に）。"""
    if len(im.planes) < 3:
        return bytes(im.planes[0])
    return bytes([(r * 77 + g * 151 + b * 28) >> 8 for r, g, b in zip(im.planes[0], im.planes[1], im.planes[2])])


def _colorspace(doc, cs) -> tuple[str, int, object]:
    """(種類, 成分の数, Indexed の表)。"""
    from .objects import Name, Stream

    cs = doc.resolve(cs)
    if isinstance(cs, Name):
        n = _CS_ABBR.get(cs.name, cs.name)
        return n, {"DeviceGray": 1, "CalGray": 1, "DeviceRGB": 3, "CalRGB": 3, "DeviceCMYK": 4}.get(n, 0), None
    if isinstance(cs, list) and cs:
        kind = getattr(doc.resolve(cs[0]), "name", "")
        kind = _CS_ABBR.get(kind, kind)
        if kind == "ICCBased":
            st = doc.resolve(cs[1])
            n = int(doc.resolve(st.dict.get("N")) or 3) if isinstance(st, Stream) else 3
            return {1: "DeviceGray", 3: "DeviceRGB", 4: "DeviceCMYK"}.get(n, "DeviceRGB"), n, None
        if kind in ("CalGray", "CalRGB", "Lab"):
            return ("DeviceGray", 1, None) if kind == "CalGray" else ("DeviceRGB", 3, None)
        if kind == "Indexed":
            base, nb, _ = _colorspace(doc, cs[1])
            lut = doc.resolve(cs[3])
            lut = doc.decode(lut) if isinstance(lut, Stream) else bytes(getattr(lut, "raw", lut) or b"")
            return "Indexed", 1, (base, nb, lut)
        if kind in ("Separation", "DeviceN"):
            return "Separation", 1 if kind == "Separation" else len(doc.resolve(cs[1]) or [1]), None
    return "DeviceGray", 1, None


def _raw_gray(doc, data: bytes, d: dict, w: int, h: int, max_pixels: int = 4_000_000) -> Gray:
    """圧縮していない（Flate などで開いた後の）画素の並び → 濃さ。幅 × 高さが max_pixels を超える画像は、
    行と列を同じ間隔で間引いて開く（画素の並びの大きさを作る前に決める：データより大きな幅・高さを書いた画像でも、
    幅 × 高さのメモリを使わない）。足りないデータは 0 とみなす。"""
    bpc = int(d.get("BitsPerComponent") or 1)
    if d.get("ImageMask") is True:
        kind, nc, idx = "DeviceGray", 1, None
        bpc = 1
    else:
        kind, nc, idx = _colorspace(doc, d.get("ColorSpace"))
    if not (w and h and nc):
        raise RasterError("画像の大きさか色空間が分かりません")
    if w < 0 or h < 0:
        raise RasterError("画像の大きさが負です")
    import math

    step = max(1, math.isqrt(w * h // max(1, max_pixels)))
    while (-(-w // step)) * (-(-h // step)) > max_pixels:
        step += 1
    stride = (w * nc * bpc + 7) // 8
    out = bytearray()
    maxv = (1 << bpc) - 1
    ow, oh = -(-w // step), -(-h // step)
    # 使う成分の番号（間引かなければ全部）
    comps = range(w * nc) if step == 1 else [x * nc + c for x in range(0, w, step) for c in range(nc)]

    def line(row: bytes) -> bytes:
        if bpc == 8:
            vals = row if step == 1 else bytes(row[i] for i in comps)
        elif bpc == 16:
            vals = row[0::2] if step == 1 else bytes(row[2 * i] for i in comps)
        else:
            vals = bytes((((row[(i * bpc) >> 3] >> (8 - bpc - (i * bpc) % 8)) & maxv) * 255 // maxv) if idx is None
                         else ((row[(i * bpc) >> 3] >> (8 - bpc - (i * bpc) % 8)) & maxv) for i in comps)
        if idx is not None:
            base, nb, lut = idx
            g = bytearray()
            for v in vals[:ow]:
                c = lut[v * nb:(v + 1) * nb] or b"\0" * nb
                g.append(_to_gray(base, c))
            return bytes(g)
        if nc == 1:
            return bytes(vals[:ow])
        return bytes(_to_gray(kind, vals[i * nc:(i + 1) * nc]) for i in range(ow))

    blank = None  # データのない行（0 とみなす）は、1 度だけ作る
    for y in range(0, h, step):
        if y * stride >= len(data):
            if blank is None:
                blank = line(bytes(stride))
            out += blank
            continue
        row = data[y * stride:(y + 1) * stride]
        if len(row) < stride:
            row = row + b"\0" * (stride - len(row))
        out += line(row)
    g = Gray(ow, oh, bytes(out))
    if kind == "Separation":  # 色の濃さ（1＝インクいっぱい）
        g = Gray(ow, oh, g.px.translate(bytes(255 - i for i in range(256))))
    return _decode_array(g, d, bpc)


def _to_gray(kind: str, c) -> int:
    if kind == "DeviceRGB" and len(c) >= 3:
        return (c[0] * 299 + c[1] * 587 + c[2] * 114) // 1000
    if kind == "DeviceCMYK" and len(c) >= 4:
        return max(0, 255 - min(255, (c[0] * 30 + c[1] * 59 + c[2] * 11) // 100 + c[3]))
    return c[0] if len(c) else 255
