"""単独の画像ファイルの形式を見分けて、フレーム（RGB の画素）に開く：JPEG・PNG・GIF・BMP・TIFF・WebP・JPEG 2000・JBIG2。

仕様：PNG（ISO/IEC 15948）・GIF89a・BMP（DIB、MS-WMF 2.2.2）・TIFF 6.0・WebP（RFC 9649・RFC 6386）。
どれも、開く前に幅 × 高さを limits.MAX_PIXELS で確かめ、展開は limits.inflate で数える。
透けのある画像は白い紙の上に置いた色にして返す（エージェントが見るのは紙の上の見た目。透けがあったことは notes で伝える）。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field
from itertools import accumulate

from ..limits import MAX_PIXELS, LimitExceeded, budget, inflate


class ImageError(ValueError):
    pass


@dataclass
class FrameInfo:
    w: int
    h: int
    dpi: tuple | None = None  # (横, 縦) 1 インチの画素。分からなければ None
    note: str = ""


@dataclass
class Frame:
    w: int
    h: int
    rgb: bytes  # w*h*3
    dpi: tuple | None = None
    notes: list = field(default_factory=list)
    problems: list = field(default_factory=list)


FORMAT_JA = {"emf": "EMF（メタファイル）", "wmf": "WMF（メタファイル）", "jpeg": "JPEG", "png": "PNG", "gif": "GIF", "bmp": "BMP", "tiff": "TIFF", "webp": "WebP", "jp2": "JPEG 2000", "j2k": "JPEG 2000（符号の流れ）",
             "jbig2": "JBIG2"}


def sniff(blob: bytes) -> str | None:
    """先頭のバイトで形式を見分ける。"""
    if blob[:3] == b"\xff\xd8\xff":
        return "jpeg"
    if blob[:8] == b"\x89PNG\r\n\x1a\n":
        return "png"
    if blob[:6] in (b"GIF87a", b"GIF89a"):
        return "gif"
    if blob[:2] == b"BM" and len(blob) > 18 and struct.unpack_from("<I", blob, 14)[0] in (12, 16, 40, 52, 56, 64, 108, 124):
        return "bmp"
    if blob[:4] in (b"II*\0", b"MM\0*", b"II+\0", b"MM\0+"):
        return "tiff"
    if blob[:4] == b"RIFF" and blob[8:12] == b"WEBP":
        return "webp"
    if blob[4:12] == b"jP  \r\n\x87\n":
        return "jp2"
    if blob[:4] == b"\xff\x4f\xff\x51":
        return "j2k"
    if blob[:8] == b"\x97JB2\r\n\x1a\n":
        return "jbig2"
    if len(blob) > 88 and blob[:4] == b"\x01\0\0\0" and blob[40:44] == b" EMF":
        return "emf"
    if blob[:4] == b"\xd7\xcd\xc6\x9a":
        return "wmf"
    return None


def heif_brand(blob: bytes) -> str | None:
    """HEIC・HEIF・AVIF（ISO BMFF の ftyp の銘柄）。"""
    if blob[4:8] != b"ftyp":
        return None
    n = struct.unpack_from(">I", blob, 0)[0] if len(blob) >= 12 else 0
    brands = {blob[8:12]} | {blob[i:i + 4] for i in range(16, min(n, 64, len(blob)), 4)}
    for b in (b"avif", b"avis"):
        if b in brands:
            return "AVIF"
    for b in (b"heic", b"heix", b"hevc", b"hevx", b"heim", b"heis", b"mif1", b"msf1"):
        if b in brands:
            return "HEIC"
    return None


def check_size(w: int, h: int, what: str) -> None:
    if w <= 0 or h <= 0:
        raise ImageError(f"{what}：画像の大きさが 0 です（{w} × {h}）")
    if w * h > MAX_PIXELS:
        raise LimitExceeded(f"{what} の画像（{w:,} × {h:,} 画素）が、開ける大きさ {MAX_PIXELS:,} 画素を超えます")
    budget().charge(w * h * 3, f"{what} の画素")


def info(blob: bytes) -> tuple[str, list[FrameInfo]]:
    """(形式, フレームの一覧)。画素は開かない。"""
    f = sniff(blob)
    if f is None:
        raise ImageError("画像のファイルとして見分けられません")
    if f == "jpeg":
        return f, [_jpeg_info(blob)]
    if f == "png":
        return f, [_png_info(blob)]
    if f == "gif":
        return f, _gif_infos(blob)
    if f == "bmp":
        return f, [_bmp_info(blob)]
    if f == "tiff":
        from .tiff import infos

        return f, infos(blob)
    if f == "webp":
        from .webp import infos as winfos

        return f, winfos(blob)
    if f in ("jp2", "j2k"):
        from . import jpx

        d = jpx.info(blob)
        return f, [FrameInfo(d["w"], d["h"], d["dpi"])]
    if f in ("emf", "wmf"):
        fr = decode(blob)
        return f, [FrameInfo(fr.w, fr.h, fr.dpi)]
    from . import jbig2

    return f, _jbig2_infos(blob, jbig2)


def decode(blob: bytes, index: int = 0, target_pixels: int | None = None, deadline: float | None = None) -> Frame:
    """index 番目のフレームを開く。target_pixels：画面に要る画素の数（JPEG・JPEG 2000 は、これより細かく開かない）。"""
    f = sniff(blob)
    if f == "jpeg":
        return _jpeg(blob, target_pixels)
    if f == "png":
        return _png(blob)
    if f == "gif":
        return _gif(blob, index)
    if f == "bmp":
        return _bmp(blob)
    if f == "tiff":
        from .tiff import decode as tdecode

        return tdecode(blob, index, target_pixels)
    if f == "webp":
        from .webp import decode as wdecode

        return wdecode(blob, index)
    if f in ("jp2", "j2k"):
        from . import jpx

        return _jpx_frame(jpx.decode(blob, target_pixels=target_pixels, deadline=deadline))
    if f in ("emf", "wmf"):
        from .metaraster import emf, wmf

        return emf(blob) if f == "emf" else wmf(blob)
    if f == "jbig2":
        from . import jbig2

        p = jbig2.decode(blob, page_number=index + 1, deadline=deadline)
        rgb = bytearray(p.w * p.h * 3)
        rgb[0::3] = rgb[1::3] = rgb[2::3] = p.px
        dpi = (p.xres * 0.0254, p.yres * 0.0254) if p.xres and p.yres else None
        return Frame(p.w, p.h, bytes(rgb), dpi, [], list(p.problems))
    raise ImageError("画像のファイルとして見分けられません")


def _jpx_frame(im) -> Frame:
    planes = im.planes
    n = len(planes)
    w, h = im.w, im.h
    out = bytearray(w * h * 3)
    if n >= 4 and im.colorspace == "CMYK":
        for c in range(3):
            out[c::3] = bytes([(255 - a) * (255 - k) // 255 for a, k in zip(planes[c], planes[3])])
    elif n >= 3:
        out[0::3], out[1::3], out[2::3] = planes[0], planes[1], planes[2]
    else:
        out[0::3] = out[1::3] = out[2::3] = planes[0]
    notes, problems = list(im.notes), list(im.problems)
    rgb = bytes(out)
    if im.alpha:
        rgb = _on_white(rgb, im.alpha)
        notes.append("透けのある画像を、白い紙の上の色で示しました")
    return Frame(w, h, rgb, im.dpi, notes, problems)


def _on_white(rgb: bytes, a: bytes) -> bytes:
    """透けを白の上に置いた色（RGB）に。"""
    if all(v == 255 for v in a):
        return rgb
    tab = [[(c * v + 255 * (255 - v) + 127) // 255 for c in range(256)] for v in range(256)]
    out = bytearray(len(rgb))
    for k in range(3):
        out[k::3] = bytes([tab[v][c] for c, v in zip(rgb[k::3], a)])
    return bytes(out)


# ====================================================================== JPEG（画素は pdf/raster.py）
def _jpeg_segments(blob: bytes):
    pos = 2
    n = len(blob)
    while pos + 4 <= n:
        if blob[pos] != 0xFF:
            pos += 1
            continue
        m = blob[pos + 1]
        if m == 0xFF:
            pos += 1
            continue
        if m == 0xDA or m == 0xD9:
            return
        if m in (0x01,) or 0xD0 <= m <= 0xD8:
            pos += 2
            continue
        ln = struct.unpack_from(">H", blob, pos + 2)[0]
        yield m, blob[pos + 4:pos + 2 + ln]
        pos += 2 + ln


def _exif_fields(seg: bytes) -> tuple[int, tuple | None]:
    """EXIF の (向き, (横の dpi, 縦の dpi))。"""
    if seg[:6] != b"Exif\0\0":
        return 1, None
    t = seg[6:]
    if t[:2] == b"II":
        e = "<"
    elif t[:2] == b"MM":
        e = ">"
    else:
        return 1, None
    try:
        off = struct.unpack_from(e + "I", t, 4)[0]
        cnt = struct.unpack_from(e + "H", t, off)[0]
        orient = 1
        xr = yr = None
        unit = 2
        for i in range(cnt):
            tag, typ, c, val = struct.unpack_from(e + "HHI4s", t, off + 2 + 12 * i)
            if tag == 0x0112:
                orient = struct.unpack(e + "H", val[:2])[0]
            elif tag == 0x0128:
                unit = struct.unpack(e + "H", val[:2])[0]
            elif tag in (0x011A, 0x011B) and typ == 5:
                p = struct.unpack(e + "I", val)[0]
                a, b = struct.unpack_from(e + "II", t, p)
                if b:
                    if tag == 0x011A:
                        xr = a / b
                    else:
                        yr = a / b
        dpi = None
        if xr and yr and unit in (2, 3):
            dpi = (xr, yr) if unit == 2 else (xr * 2.54, yr * 2.54)
        return orient if 1 <= orient <= 8 else 1, dpi
    except (struct.error, IndexError):
        return 1, None


def _jpeg_meta(blob: bytes) -> tuple[int, int, tuple | None, int]:
    w = h = 0
    dpi = None
    orient = 1
    for m, seg in _jpeg_segments(blob):
        if 0xC0 <= m <= 0xCF and m not in (0xC4, 0xC8, 0xCC) and len(seg) >= 5:
            h, w = struct.unpack_from(">HH", seg, 1)
        elif m == 0xE0 and seg[:5] == b"JFIF\0" and len(seg) >= 12 and dpi is None:
            unit, xd, yd = seg[7], *struct.unpack_from(">HH", seg, 8)
            if unit == 1 and xd and yd:
                dpi = (float(xd), float(yd))
            elif unit == 2 and xd and yd:
                dpi = (xd * 2.54, yd * 2.54)
        elif m == 0xE1:
            o2, d2 = _exif_fields(seg)
            if seg[:6] == b"Exif\0\0":
                orient = o2
                if d2 and dpi is None:
                    dpi = d2
    return w, h, dpi, orient


def _jpeg_info(blob: bytes) -> FrameInfo:
    w, h, dpi, orient = _jpeg_meta(blob)
    if not w or not h:
        raise ImageError("JPEG：フレームの見出しがありません")
    if orient >= 5:
        w, h = h, w
        dpi = (dpi[1], dpi[0]) if dpi else None
    return FrameInfo(w, h, dpi)


def _jpeg(blob: bytes, target_pixels: int | None) -> Frame:
    from ..pdf.raster import RasterError, jpeg_gray

    w, h, dpi, orient = _jpeg_meta(blob)
    if not w or not h:
        raise ImageError("JPEG：フレームの見出しがありません")
    check_size(w, h, "JPEG")
    try:
        im = jpeg_gray(blob, target_pixels * 4 if target_pixels else MAX_PIXELS, want_rgb=True)
    except RasterError as e:
        raise ImageError(str(e)) from None
    notes = []
    if im.w != w:
        notes.append(f"JPEG を {w}×{h} 画素の {w // im.w if im.w else 1} 分の 1 で開きました（region で拡大すると細かく開きます）")
    rgb, ow, oh = orient_rgb(im.px, im.w, im.h, orient)
    if orient != 1:
        notes.append(f"EXIF の向き（{orient}）のとおりに回しました")
    if orient >= 5 and dpi:
        dpi = (dpi[1], dpi[0])
    if dpi and im.w != w:
        dpi = (dpi[0] * im.w / w, dpi[1] * im.h / h)
    return Frame(ow, oh, rgb, dpi, notes, [])


def orient_rgb(px: bytes, w: int, h: int, o: int) -> tuple[bytes, int, int]:
    """TIFF・EXIF の向き 1〜8 のとおりに回す。"""
    if o == 1 or o not in range(1, 9):
        return px, w, h
    rows = [px[y * w * 3:(y + 1) * w * 3] for y in range(h)]
    if o in (2, 3):
        rows = [b"".join(r[i:i + 3] for i in range(len(r) - 3, -1, -3)) for r in rows]
    if o in (3, 4):
        rows = rows[::-1]
    if o in (2, 3, 4):
        return b"".join(rows), w, h
    cols = []
    for x in range(w):
        i = x * 3
        col = bytearray(h * 3)
        col[0::3] = bytes([r[i] for r in rows])
        col[1::3] = bytes([r[i + 1] for r in rows])
        col[2::3] = bytes([r[i + 2] for r in rows])
        cols.append(bytes(col))
    # cols[x] は、元の x 列を上から下へ。新しい画像の行にする
    if o == 5:  # 転置
        out = cols
    elif o == 6:  # 90 度右へ：新しい行 y は元の列 y を下から上へ
        out = [b"".join(c[i:i + 3] for i in range(len(c) - 3, -1, -3)) for c in cols]
    elif o == 7:  # 転置して 180 度
        out = [b"".join(c[i:i + 3] for i in range(len(c) - 3, -1, -3)) for c in cols[::-1]]
    else:  # 8：90 度左へ
        out = cols[::-1]
    return b"".join(out), h, w


# ====================================================================== PNG
def _png_chunks(blob: bytes):
    pos = 8
    n = len(blob)
    while pos + 8 <= n:
        ln, t = struct.unpack_from(">I4s", blob, pos)
        yield t, blob[pos + 8:pos + 8 + ln]
        pos += 12 + ln
        if t == b"IEND":
            return


def _png_info(blob: bytes) -> FrameInfo:
    w = h = 0
    dpi = None
    for t, d in _png_chunks(blob):
        if t == b"IHDR":
            w, h = struct.unpack_from(">II", d, 0)
        elif t == b"pHYs" and d[8] == 1:
            x, y = struct.unpack_from(">II", d, 0)
            dpi = (x * 0.0254, y * 0.0254)
        elif t == b"IDAT":
            break
    if not w or not h:
        raise ImageError("PNG：IHDR がありません")
    return FrameInfo(w, h, dpi)


_ADAM7 = ((0, 0, 8, 8), (4, 0, 8, 8), (0, 4, 4, 8), (2, 0, 4, 4), (0, 2, 2, 4), (1, 0, 2, 2), (0, 1, 1, 2))
_UNPACK = {b: [bytes((v >> (8 - b * (i + 1))) & ((1 << b) - 1) for i in range(8 // b)) for v in range(256)] for b in (1, 2, 4)}


def _unfilter(raw: bytes, pos: int, rows: int, stride: int, bpp: int) -> tuple[list[bytes], int]:
    """PNG のフィルター（9）を戻す。"""
    out = []
    prev = bytes(stride)
    for _y in range(rows):
        if pos >= len(raw):
            cur = bytes(stride)
            out.append(cur)
            prev = cur
            continue
        ft = raw[pos]
        row = raw[pos + 1:pos + 1 + stride]
        pos += 1 + stride
        if len(row) < stride:
            row = row + bytes(stride - len(row))
        if ft == 0:
            cur = row
        elif ft == 1:
            c = bytearray(stride)
            for k in range(bpp):
                c[k::bpp] = bytes([v & 255 for v in accumulate(row[k::bpp])])
            cur = bytes(c)
        elif ft == 2:
            cur = bytes([(a + b) & 255 for a, b in zip(row, prev)])
        elif ft == 3:
            c = bytearray(stride)
            for k in range(bpp):
                left = 0
                j = k
                for a, up in zip(row[k::bpp], prev[k::bpp]):
                    left = (a + ((left + up) >> 1)) & 255
                    c[j] = left
                    j += bpp
            cur = bytes(c)
        elif ft == 4:
            c = bytearray(stride)
            for k in range(bpp):
                a_ = c_ = 0
                j = k
                for v, b_ in zip(row[k::bpp], prev[k::bpp]):
                    p = a_ + b_ - c_
                    pa, pb, pc = abs(p - a_), abs(p - b_), abs(p - c_)
                    pr = a_ if pa <= pb and pa <= pc else b_ if pb <= pc else c_
                    a_ = (v + pr) & 255
                    c[j] = a_
                    c_ = b_
                    j += bpp
            cur = bytes(c)
        else:
            raise ImageError(f"PNG：行のフィルターの種類 {ft} は規格にありません")
        out.append(cur)
        prev = cur
    return out, pos


def _samples_row(row: bytes, depth: int, n: int) -> bytes:
    """1 行を、標本ごとに 1 バイトに（16 ビットは上の桁、1・2・4 ビットはそのままの値）。"""
    if depth == 8:
        return row[:n]
    if depth == 16:
        return row[0:2 * n:2]
    tab = _UNPACK[depth]
    out = bytearray()
    for b in row[:(n * depth + 7) // 8]:
        out += tab[b]
    return bytes(out[:n])


def _png(blob: bytes) -> Frame:
    ihdr = None
    plte = trns = None
    idat = []
    apng = False
    for t, d in _png_chunks(blob):
        if t == b"IHDR":
            ihdr = d
        elif t == b"PLTE":
            plte = d
        elif t == b"tRNS":
            trns = d
        elif t == b"IDAT":
            idat.append(d)
        elif t == b"acTL":
            apng = True
    if ihdr is None or len(ihdr) < 13:
        raise ImageError("PNG：IHDR がありません")
    w, h, depth, ct, _cm, _fm, il = struct.unpack_from(">IIBBBBB", ihdr, 0)
    check_size(w, h, "PNG")
    channels = {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}.get(ct)
    if channels is None or depth not in (1, 2, 4, 8, 16):
        raise ImageError(f"PNG：色の型 {ct}・ビットの深さ {depth} は規格にありません")
    raw = inflate(b"".join(idat), what="PNG の画素", tolerant=True)
    bpp = max(1, channels * depth // 8)
    frame = _png_info(blob)
    notes = []
    samples: list = []  # 行ごとの、標本 1 バイトの並び
    if il == 0:
        stride = (w * channels * depth + 7) // 8
        rows, _p = _unfilter(raw, 0, h, stride, bpp)
        samples = [_samples_row(r, depth, w * channels) for r in rows]
    elif il == 1:
        canvas = [bytearray(w * channels) for _ in range(h)]
        pos = 0
        for x0, y0, dx, dy in _ADAM7:
            pw = (w - x0 + dx - 1) // dx
            ph = (h - y0 + dy - 1) // dy
            if pw <= 0 or ph <= 0:
                continue
            stride = (pw * channels * depth + 7) // 8
            rows, pos = _unfilter(raw, pos, ph, stride, bpp)
            for j, r in enumerate(rows):
                s = _samples_row(r, depth, pw * channels)
                row = canvas[y0 + j * dy]
                for c in range(channels):
                    row[x0 * channels + c::dx * channels] = s[c::channels]
        samples = [bytes(r) for r in canvas]
    else:
        raise ImageError(f"PNG：インターレースの方法 {il} は規格にありません")
    maxv = (1 << min(depth, 8)) - 1
    out = bytearray(w * h * 3)
    alpha = None
    flat = b"".join(samples)
    if ct == 3:
        if not plte:
            raise ImageError("PNG：パレットがありません")
        pal = plte + bytes(768 - len(plte)) if len(plte) < 768 else plte
        for k in range(3):
            out[k::3] = flat.translate(pal[k::3][:256].ljust(256, b"\0"))
        if trns:
            alpha = flat.translate(trns[:256].ljust(256, b"\xff"))
    elif ct in (0, 4):
        g = flat[0::channels]
        if depth < 8:
            g = g.translate(bytes(v * 255 // maxv if v <= maxv else 255 for v in range(256)))
        out[0::3] = out[1::3] = out[2::3] = g
        if ct == 4:
            alpha = flat[1::2]
        elif trns and len(trns) >= 2:
            key = struct.unpack(">H", trns[:2])[0]
            key = key >> 8 if depth == 16 else key
            alpha = flat.translate(bytes(0 if v == key else 255 for v in range(256)))
    else:
        out[0::3], out[1::3], out[2::3] = flat[0::channels], flat[1::channels], flat[2::channels]
        if ct == 6:
            alpha = flat[3::4]
        elif trns and len(trns) >= 6:
            kr, kg, kb = struct.unpack(">HHH", trns[:6])
            if depth == 16:
                kr, kg, kb = kr >> 8, kg >> 8, kb >> 8
            key = bytes((kr, kg, kb))
            alpha = bytes([0 if flat[i:i + 3] == key else 255 for i in range(0, len(flat), 3)])
    rgb = bytes(out)
    if alpha is not None and any(v != 255 for v in alpha):
        rgb = _on_white(rgb, alpha)
        notes.append("透けのある画像を、白い紙の上の色で示しました")
    if apng:
        notes.append("動く PNG の、最初の画像を示しました")
    return Frame(w, h, rgb, frame.dpi, notes, [])


# ====================================================================== GIF
def _gif_blocks(blob: bytes, pos: int) -> tuple[bytes, int]:
    out = bytearray()
    n = len(blob)
    while pos < n:
        k = blob[pos]
        pos += 1
        if k == 0:
            break
        out += blob[pos:pos + k]
        pos += k
    return bytes(out), pos


def _gif_walk(blob: bytes):
    """(画像の位置, GCE) の順に。"""
    w, h, fl = struct.unpack_from("<HHB", blob, 6)
    pos = 13
    gct = None
    if fl & 0x80:
        n = 2 << (fl & 7)
        gct = blob[pos:pos + 3 * n]
        pos += 3 * n
    gce = (0, 0, None)
    frames = []
    size = len(blob)
    while pos < size:
        b = blob[pos]
        pos += 1
        if b == 0x3B:
            break
        if b == 0x21:
            lab = blob[pos] if pos < size else 0
            pos += 1
            if lab == 0xF9 and pos + 5 < size:
                f2 = blob[pos + 1]
                delay = struct.unpack_from("<H", blob, pos + 2)[0]
                gce = ((f2 >> 2) & 7, delay, blob[pos + 4] if f2 & 1 else None)
            _d, pos = _gif_blocks(blob, pos)
        elif b == 0x2C:
            if pos + 9 > size:
                break
            x, y, iw, ih, f3 = struct.unpack_from("<HHHHB", blob, pos)
            pos += 9
            lct = None
            if f3 & 0x80:
                n = 2 << (f3 & 7)
                lct = blob[pos:pos + 3 * n]
                pos += 3 * n
            min_code = blob[pos] if pos < size else 0
            pos += 1
            data, pos = _gif_blocks(blob, pos)
            frames.append({"x": x, "y": y, "w": iw, "h": ih, "interlace": bool(f3 & 0x40), "lct": lct, "min": min_code, "data": data, "gce": gce})
            gce = (0, 0, None)
        else:
            break
    return w, h, gct, frames, fl, (blob[11] if fl & 0x80 else 0)


def _gif_infos(blob: bytes) -> list[FrameInfo]:
    w, h, _g, frames, _f, _b = _gif_walk(blob)
    if not frames:
        raise ImageError("GIF：画像がありません")
    return [FrameInfo(w, h) for _ in frames]


def _gif_lzw(data: bytes, min_code: int, want: int) -> bytes:
    """GIF の LZW（下位のビットから、可変の長さ）。"""
    if not 1 <= min_code <= 11:
        raise ImageError("GIF：LZW の最小の符号の長さが範囲の外です")
    clear = 1 << min_code
    eoi = clear + 1
    base = [bytes([i]) for i in range(clear)] + [b"", b""]
    table = list(base)
    size = min_code + 1
    acc = 0
    nb = 0
    prev = None
    out = bytearray()
    for byte in data:
        acc |= byte << nb
        nb += 8
        while nb >= size:
            code = acc & ((1 << size) - 1)
            acc >>= size
            nb -= size
            if code == clear:
                table = list(base)
                size = min_code + 1
                prev = None
                continue
            if code == eoi:
                return bytes(out)
            if prev is None:
                if code >= len(table):
                    return bytes(out)
                ent = table[code]
            elif code < len(table):
                ent = table[code]
                if len(table) < 4096:
                    table.append(prev + ent[:1])
            elif code == len(table) and len(table) < 4096:
                ent = prev + prev[:1]
                table.append(ent)
            else:
                return bytes(out)
            out += ent
            prev = ent
            if len(table) == (1 << size) and size < 12:
                size += 1
            if len(out) >= want:
                return bytes(out[:want])
    return bytes(out)


def _gif(blob: bytes, index: int) -> Frame:
    W, H, gct, frames, fl, bg_index = _gif_walk(blob)
    check_size(W, H, "GIF")
    if not frames:
        raise ImageError("GIF：画像がありません")
    if not 0 <= index < len(frames):
        raise ImageError(f"GIF：画像は {len(frames)} 枚です")
    canvas = bytearray(b"\xff" * (W * H * 3))
    notes = []
    for i in range(index + 1):
        fr = frames[i]
        fw, fh = fr["w"], fr["h"]
        check_size(fw, fh, "GIF")
        ct = fr["lct"] or gct
        if not ct:
            raise ImageError("GIF：色の表がありません")
        px = _gif_lzw(fr["data"], fr["min"], fw * fh).ljust(fw * fh, b"\0")
        disposal, _delay, trans = fr["gce"]
        saved = bytes(canvas) if disposal == 3 else None
        pal = ct.ljust(768, b"\0")
        r, g, b = pal[0::3][:256], pal[1::3][:256], pal[2::3][:256]
        order = list(range(fh))
        if fr["interlace"]:
            order = list(range(0, fh, 8)) + list(range(4, fh, 8)) + list(range(2, fh, 4)) + list(range(1, fh, 2))
        for srow, y in enumerate(order):
            Y = fr["y"] + y
            if Y >= H:
                continue
            row = px[srow * fw:(srow + 1) * fw]
            n = min(fw, W - fr["x"])
            if n <= 0:
                continue
            base = (Y * W + fr["x"]) * 3
            if trans is None:
                line = bytearray(n * 3)
                line[0::3], line[1::3], line[2::3] = row[:n].translate(r), row[:n].translate(g), row[:n].translate(b)
                canvas[base:base + n * 3] = line
            else:
                for x in range(n):
                    v = row[x]
                    if v != trans:
                        canvas[base + x * 3:base + x * 3 + 3] = bytes((r[v], g[v], b[v]))
        if i < index:
            if disposal == 2:
                for y in range(fh):
                    Y = fr["y"] + y
                    n = min(fw, W - fr["x"])
                    if Y < H and n > 0:
                        canvas[(Y * W + fr["x"]) * 3:(Y * W + fr["x"] + n) * 3] = b"\xff" * (n * 3)
            elif disposal == 3 and saved is not None:
                canvas = bytearray(saved)
    if len(frames) > 1:
        notes.append(f"動く GIF の {index + 1} 枚目（全部で {len(frames)} 枚）")
    return Frame(W, H, bytes(canvas), None, notes, [])


# ====================================================================== BMP
def _bmp_head(blob: bytes) -> dict:
    off, hs = struct.unpack_from("<I", blob, 10)[0], struct.unpack_from("<I", blob, 14)[0]
    h: dict = {"off": off, "hs": hs}
    if hs == 12:
        w, ht, _pl, bpp = struct.unpack_from("<HHHH", blob, 18)
        comp = 0
        clr = 0
        xp = yp = 0
    else:
        w, ht, _pl, bpp, comp, _sz, xp, yp, clr = struct.unpack_from("<iiHHIIiiI", blob, 18)
    h.update(w=w, h=abs(ht), top=ht < 0, bpp=bpp, comp=comp, clr=clr, xp=xp, yp=yp)
    return h


def _bmp_info(blob: bytes) -> FrameInfo:
    h = _bmp_head(blob)
    dpi = (h["xp"] * 0.0254, h["yp"] * 0.0254) if h["xp"] > 0 and h["yp"] > 0 else None
    return FrameInfo(h["w"], h["h"], dpi)


def _mask_shift(m: int) -> tuple[int, int]:
    if not m:
        return 0, 0
    s = 0
    while not (m >> s) & 1:
        s += 1
    n = 0
    while (m >> (s + n)) & 1:
        n += 1
    return s, n


def _bmp(blob: bytes) -> Frame:
    return dib_frame(blob, 14, _bmp_info(blob).dpi, struct.unpack_from("<I", blob, 10)[0])


def dib_frame(blob: bytes, head_at: int, dpi, pixel_at: int | None = None) -> Frame:
    """DIB（BITMAPINFO＋画素）を開く：BMP・メタファイルの中の絵。pixel_at：画素の位置（省くと色の表の次）。"""
    hs = struct.unpack_from("<I", blob, head_at)[0]
    if hs == 12:
        w, ht, _pl, bpp = struct.unpack_from("<HHHH", blob, head_at + 4)
        comp = clr = 0
    else:
        w, ht, _pl, bpp, comp, _sz, _x, _y, clr = struct.unpack_from("<iiHHIIiiI", blob, head_at + 4)
    top = ht < 0
    h = abs(ht)
    check_size(w, h, "BMP")
    pal_at = head_at + hs
    masks = None
    if comp == 3 or comp == 6:
        if hs >= 52:
            masks = struct.unpack_from("<III", blob, head_at + 40)
        else:
            masks = struct.unpack_from("<III", blob, pal_at)
            pal_at += 12
    entry = 3 if hs == 12 else 4
    ncol = (clr or (1 << bpp)) if bpp <= 8 else 0
    pal = blob[pal_at:pal_at + ncol * entry]
    data_at = pixel_at if pixel_at is not None else pal_at + ncol * entry
    if comp in (4, 5):  # BI_JPEG・BI_PNG：画素が丸ごと JPEG・PNG
        return decode(blob[data_at:])
    b_tab = bytes(pal[0::entry]).ljust(256, b"\0")
    g_tab = bytes(pal[1::entry]).ljust(256, b"\0")
    r_tab = bytes(pal[2::entry]).ljust(256, b"\0")
    out = bytearray(w * h * 3)
    stride = ((w * bpp + 31) // 32) * 4
    if comp in (1, 2):  # RLE8・RLE4
        idx = _bmp_rle(blob, data_at, w, h, comp)
        for y in range(h):
            row = idx[y * w:(y + 1) * w]
            Y = y if top else h - 1 - y
            o = Y * w * 3
            line = bytearray(w * 3)
            line[0::3], line[1::3], line[2::3] = row.translate(r_tab), row.translate(g_tab), row.translate(b_tab)
            out[o:o + w * 3] = line
        return Frame(w, h, bytes(out), dpi)
    if comp not in (0, 3, 6):
        raise ImageError(f"BMP：圧縮の種類 {comp} は扱えません")
    if bpp not in (1, 4, 8, 16, 24, 32):
        raise ImageError(f"BMP：{bpp} ビットの画素は扱えません")
    for y in range(h):
        row = blob[data_at + y * stride:data_at + (y + 1) * stride]
        Y = y if top else h - 1 - y
        o = Y * w * 3
        line = bytearray(w * 3)
        if bpp in (1, 4):
            vals = _samples_row(row, bpp, w)
            line[0::3], line[1::3], line[2::3] = vals.translate(r_tab), vals.translate(g_tab), vals.translate(b_tab)
        elif bpp == 8:
            vals = row[:w].ljust(w, b"\0")
            line[0::3], line[1::3], line[2::3] = vals.translate(r_tab), vals.translate(g_tab), vals.translate(b_tab)
        elif bpp == 24:
            row = row[:w * 3].ljust(w * 3, b"\0")
            line[0::3], line[1::3], line[2::3] = row[2::3], row[1::3], row[0::3]
        elif bpp == 32 and masks is None:
            row = row[:w * 4].ljust(w * 4, b"\0")
            line[0::3], line[1::3], line[2::3] = row[2::4], row[1::4], row[0::4]
        else:
            m = masks or (0x7C00, 0x03E0, 0x001F)
            sh = [_mask_shift(x) for x in m]
            width = bpp // 8
            vals = struct.unpack_from(f"<{w}{'H' if bpp == 16 else 'I'}", row.ljust(w * width, b"\0"))
            for k in range(3):
                s, n = sh[k]
                mk = m[k]
                if n == 0:
                    continue
                top_v = (1 << n) - 1
                line[k::3] = bytes([((v & mk) >> s) * 255 // top_v for v in vals])
        out[o:o + w * 3] = line
    return Frame(w, h, bytes(out), dpi)


def _bmp_rle(blob: bytes, pos: int, w: int, h: int, comp: int) -> bytes:
    """RLE8・RLE4（下の行から）。"""
    out = bytearray(w * h)
    x = y = 0
    n = len(blob)
    four = comp == 2
    while pos + 1 < n and y < h:
        a, b = blob[pos], blob[pos + 1]
        pos += 2
        if a:
            if four:
                vals = bytes(((b >> 4) if i % 2 == 0 else (b & 15)) for i in range(a))
            else:
                vals = bytes([b]) * a
            k = min(len(vals), w - x)
            if k > 0:
                o = y * w + x
                out[o:o + k] = vals[:k]
            x += a
        elif b == 0:
            x = 0
            y += 1
        elif b == 1:
            break
        elif b == 2:
            x += blob[pos]
            y += blob[pos + 1]
            pos += 2
        else:
            if four:
                nb = (b + 1) // 2
                d = blob[pos:pos + nb]
                vals = bytes((d[i // 2] >> 4) if i % 2 == 0 else (d[i // 2] & 15) for i in range(b))
            else:
                nb = b
                vals = blob[pos:pos + nb]
            pos += nb + (nb & 1)
            k = min(len(vals), w - x)
            if k > 0 and y < h:
                o = y * w + x
                out[o:o + k] = vals[:k]
            x += b
    return bytes(out)


# ====================================================================== JBIG2（単独のファイル）
def _jbig2_infos(blob: bytes, jbig2) -> list[FrameInfo]:
    out = []
    for sg in jbig2._parse_segments(blob):
        if sg.type == 48 and len(sg.data) >= 16:
            w, h, xr, yr = struct.unpack_from(">IIII", sg.data, 0)
            dpi = (xr * 0.0254, yr * 0.0254) if xr and yr else None
            out.append(FrameInfo(w, h if h != 0xFFFFFFFF else 0, dpi))
    if not out:
        raise ImageError("JBIG2：ページの区切りがありません")
    return out
