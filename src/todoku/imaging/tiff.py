"""TIFF 6.0（と BigTIFF）を開く：圧縮なし・PackBits・LZW・Deflate・CCITT（1 次元・G3・G4）・JPEG（新旧）・JPEG 2000。複数のページ。

スキャナーの TIFF を読むための実装：ストリップとタイル、面ごとの標本（PlanarConfiguration）、水平の差分（Predictor 2）、
FillOrder、Photometric（白黒・パレット・RGB・CMYK・YCbCr・Lab）、Orientation、解像度（dpi）。
"""

from __future__ import annotations

import struct

from ..limits import budget, inflate
from .formats import Frame, FrameInfo, ImageError, _on_white, check_size, orient_rgb

_TYPE = {1: ("B", 1), 2: ("s", 1), 3: ("H", 2), 4: ("I", 4), 5: ("rat", 8), 6: ("b", 1), 7: ("s", 1), 8: ("h", 2), 9: ("i", 4), 10: ("srat", 8),
         11: ("f", 4), 12: ("d", 8), 13: ("I", 4), 16: ("Q", 8), 17: ("q", 8), 18: ("Q", 8)}
_REVERSE = bytes(int(f"{i:08b}"[::-1], 2) for i in range(256))


def _ifds(blob: bytes) -> tuple[str, list[dict]]:
    e = "<" if blob[:2] == b"II" else ">"
    big = struct.unpack_from(e + "H", blob, 2)[0] == 43
    if big:
        off = struct.unpack_from(e + "Q", blob, 8)[0]
    else:
        off = struct.unpack_from(e + "I", blob, 4)[0]
    out = []
    seen = set()
    while off and off not in seen and len(out) < 4096:
        seen.add(off)
        if off + 8 > len(blob):
            break
        tags: dict = {}
        if big:
            n = struct.unpack_from(e + "Q", blob, off)[0]
            p = off + 8
            esz = 20
        else:
            n = struct.unpack_from(e + "H", blob, off)[0]
            p = off + 2
            esz = 12
        if n > 4096 or p + n * esz > len(blob):
            break
        for i in range(n):
            q = p + i * esz
            tag, typ = struct.unpack_from(e + "HH", blob, q)
            if big:
                cnt = struct.unpack_from(e + "Q", blob, q + 4)[0]
                valpos = q + 12
                inline = 8
            else:
                cnt = struct.unpack_from(e + "I", blob, q + 4)[0]
                valpos = q + 8
                inline = 4
            t = _TYPE.get(typ)
            if t is None:
                continue
            total = cnt * t[1]
            if total > inline:
                valpos = struct.unpack_from(e + ("Q" if big else "I"), blob, valpos)[0]
            if valpos + total > len(blob) or cnt > 1 << 26:
                continue
            code = t[0]
            if code == "s":
                tags[tag] = blob[valpos:valpos + cnt]
            elif code in ("rat", "srat"):
                f = e + ("II" if code == "rat" else "ii") * cnt
                v = struct.unpack_from(f, blob, valpos)
                tags[tag] = [v[k] / v[k + 1] if v[k + 1] else 0.0 for k in range(0, 2 * cnt, 2)]
            else:
                tags[tag] = list(struct.unpack_from(f"{e}{cnt}{code}", blob, valpos))
        out.append(tags)
        q = p + n * esz
        if big:
            off = struct.unpack_from(e + "Q", blob, q)[0] if q + 8 <= len(blob) else 0
        else:
            off = struct.unpack_from(e + "I", blob, q)[0] if q + 4 <= len(blob) else 0
    return e, out


def _one(t: dict, tag: int, default=None):
    v = t.get(tag)
    if isinstance(v, list) and v:
        return v[0]
    return default if v is None or isinstance(v, list) else v


def _dpi(t: dict) -> tuple | None:
    x, y = _one(t, 282), _one(t, 283)
    unit = _one(t, 296, 2)
    if not x or not y or unit == 1:
        return None
    return (x, y) if unit == 2 else (x * 2.54, y * 2.54)


def _pages(blob: bytes) -> tuple[str, list[dict]]:
    """ページ（縮小画像の印 NewSubfileType の bit 0 のあるものは除く。全部が縮小なら全部）。"""
    e, ifds = _ifds(blob)
    pages = [t for t in ifds if not (_one(t, 254, 0) & 1) and 256 in t and 257 in t]
    return e, (pages or [t for t in ifds if 256 in t and 257 in t])


def infos(blob: bytes) -> list[FrameInfo]:
    _e, pages = _pages(blob)
    if not pages:
        raise ImageError("TIFF：画像のページがありません")
    out = []
    for t in pages:
        w, h = _one(t, 256, 0), _one(t, 257, 0)
        dpi = _dpi(t)
        if _one(t, 274, 1) >= 5:
            w, h = h, w
            dpi = (dpi[1], dpi[0]) if dpi else None
        out.append(FrameInfo(w, h, dpi))
    return out


# ---------------------------------------------------------------------- 圧縮
def _packbits(d: bytes, want: int) -> bytes:
    out = bytearray()
    i, n = 0, len(d)
    while i < n and len(out) < want:
        c = d[i]
        i += 1
        if c < 128:
            out += d[i:i + c + 1]
            i += c + 1
        elif c > 128:
            if i < n:
                out += d[i:i + 1] * (257 - c)
            i += 1
    return bytes(out)


def _lzw(d: bytes, want: int) -> bytes:
    """TIFF の LZW（上位のビットから。符号の切り替えは PDF の LZW と同じく 1 つ早い：EarlyChange）。"""
    base = [bytes([i]) for i in range(256)] + [b"", b""]
    table = list(base)
    size = 9
    acc = nb = 0
    prev = None
    out = bytearray()
    cap = min(want, budget().room())
    for byte in d:
        acc = (acc << 8) | byte
        nb += 8
        while nb >= size:
            nb -= size
            code = (acc >> nb) & ((1 << size) - 1)
            acc &= (1 << nb) - 1
            if code == 256:
                table = list(base)
                size = 9
                prev = None
                continue
            if code == 257:
                return bytes(out)
            if prev is None:
                if code >= len(table):
                    return bytes(out)
                ent = table[code]
            elif code < len(table):
                ent = table[code]
                table.append(prev + ent[:1])
            elif code == len(table):
                ent = prev + prev[:1]
                table.append(ent)
            else:
                return bytes(out)
            out += ent
            prev = ent
            if len(out) >= cap:
                return bytes(out[:cap])
            k = len(table) + 1  # 1 つ早く切り替える
            if k >= 2048:
                size = 12
            elif k >= 1024:
                size = 11
            elif k >= 512:
                size = 10
    return bytes(out)


def _undiff(row: bytearray, n: int, bits: int) -> None:
    """Predictor 2（水平の差分）。"""
    if bits == 8:
        for c in range(n):
            acc = 0
            j = c
            for v in row[c::n]:
                acc = (acc + v) & 255
                row[j] = acc
                j += n
    elif bits == 16:
        for c in range(n):
            acc = 0
            for x in range(c, len(row) // 2, n):
                v = row[2 * x] << 8 | row[2 * x + 1]
                acc = (acc + v) & 0xFFFF
                row[2 * x], row[2 * x + 1] = acc >> 8, acc & 255


# ---------------------------------------------------------------------- 開く
def decode(blob: bytes, index: int, target_pixels: int | None = None, deadline: float | None = None, window: tuple | None = None) -> Frame:
    e, pages = _pages(blob)
    if not 0 <= index < len(pages):
        raise ImageError(f"TIFF：ページは {len(pages)} 枚です")
    t = pages[index]
    w, h = _one(t, 256, 0), _one(t, 257, 0)
    check_size(w, h, "TIFF")
    comp = _one(t, 259, 1)
    photo = _one(t, 262, 1 if comp in (2, 3, 4) else 2 if _one(t, 277, 1) >= 3 else 1)
    spp = _one(t, 277, 1)
    bits_all = t.get(258) or [1]
    bits = bits_all[0]
    planar = _one(t, 284, 1)
    pred = _one(t, 317, 1)
    fill_order = _one(t, 266, 1)
    orient = _one(t, 274, 1)
    extra = t.get(338) or []
    notes: list = []
    problems: list = []
    sf = _one(t, 339, 1)
    if sf == 3:
        raise ImageError("TIFF：浮動小数の標本は扱えません")
    if bits not in (1, 2, 4, 8, 16):
        raise ImageError(f"TIFF：{bits} ビットの標本は扱えません")
    # 区画（ストリップかタイル）
    tiled = 322 in t
    if tiled:
        tw, tl = _one(t, 322), _one(t, 323)
        offs, cnts = t.get(324) or [], t.get(325) or []
        across = (w + tw - 1) // tw if tw else 0
        down = (h + tl - 1) // tl if tl else 0
    else:
        rps = _one(t, 278, h) or h
        rps = min(rps, h)
        tw, tl = w, rps
        offs, cnts = t.get(273) or [], t.get(279) or []
        across = 1
        down = (h + rps - 1) // rps
    if not offs:
        raise ImageError("TIFF：画像のデータの位置がありません")
    if not cnts:
        cnts = [len(blob) - offs[0]] if len(offs) == 1 else []
    planes = spp if planar == 2 else 1
    per = 1 if planar == 2 else spp
    stride_t = (tw * per * bits + 7) // 8
    # 開いた標本は、平面ごとに、行の並び（1 行は per × 幅 × bits ビット）
    full = [bytearray(((w * per * bits + 7) // 8) * h) for _ in range(planes)]
    stride_f = (w * per * bits + 7) // 8
    jpeg_rgb = None
    n_expected = across * down * planes
    for i in range(min(n_expected, len(offs))):
        pl = i // (across * down)
        k = i % (across * down)
        tx, ty = (k % across) * tw, (k // across) * tl
        size = cnts[i] if i < len(cnts) else 0
        raw = blob[offs[i]:offs[i] + size]
        if fill_order == 2:
            raw = raw.translate(_REVERSE)
        rows_here = min(tl, h - ty)
        want = stride_t * (tl if tiled else rows_here)
        if comp == 1:
            data = raw
        elif comp == 32773:
            data = _packbits(raw, want)
        elif comp == 5:
            data = _lzw(raw, want)
        elif comp in (8, 32946):
            data = inflate(raw, what="TIFF の Deflate", tolerant=True)
        elif comp in (2, 3, 4):
            data = _ccitt(raw, comp, _one(t, 292, 0), tw, tl if tiled else rows_here)
        elif comp == 7 or comp == 6:
            from ..pdf.raster import RasterError, jpeg_gray

            stream = _jpeg_stream(blob, t, raw, comp, photo)
            try:
                im = jpeg_gray(stream, 100_000_000, want_rgb=True)
            except RasterError as ex:
                raise ImageError(f"TIFF：JPEG の区画を開けません（{ex}）") from None
            if jpeg_rgb is None:
                jpeg_rgb = bytearray(b"\xff" * (w * h * 3))
            n = min(im.w, w - tx) * 3
            for y in range(min(im.h, h - ty)):
                o = ((ty + y) * w + tx) * 3
                jpeg_rgb[o:o + n] = im.px[y * im.w * 3:y * im.w * 3 + n]
            continue
        elif comp == 34712:
            from . import jpx

            if len(offs) == 1:  # 全体が 1 つの JPEG 2000：解像度と範囲を選んで開ける（予算は開く分に掛かる）
                from .formats import _jpx_frame, shown_rect, stored_rect

                im = jpx.decode(raw, target_pixels=target_pixels, deadline=deadline, window=stored_rect(window, w, h, orient) if window else None)
                fr = _jpx_frame(im)
                rgb, ow, oh = orient_rgb(fr.rgb, fr.w, fr.h, orient)
                return Frame(ow, oh, rgb, _dpi(t), notes + fr.notes, problems + fr.problems, shown_rect(im.window, w, h, orient) if im.window else None)
            im = jpx.decode(raw, max_t1_bytes=10**12, deadline=deadline)  # 区画に分かれたもの：区画を並べるので、解像度は下げない
            if jpeg_rgb is None:
                jpeg_rgb = bytearray(b"\xff" * (w * h * 3))
            ps = im.planes if len(im.planes) >= 3 else [im.planes[0]] * 3
            for y in range(min(im.h, h - ty)):
                n = min(im.w, w - tx)
                o = ((ty + y) * w + tx) * 3
                line = bytearray(n * 3)
                base = y * im.w
                line[0::3], line[1::3], line[2::3] = ps[0][base:base + n], ps[1][base:base + n], ps[2][base:base + n]
                jpeg_rgb[o:o + n * 3] = line
            continue
        else:
            raise ImageError(f"TIFF：圧縮の種類 {comp} は扱えません")
        dst = full[pl]
        for y in range(rows_here):
            row = bytearray(data[y * stride_t:(y + 1) * stride_t])
            if len(row) < stride_t:
                row += bytes(stride_t - len(row))
                if not problems:
                    problems.append("TIFF：画像のデータが途中で切れています（足りない所は 0 として示しました）")
            if pred == 2:
                _undiff(row, per, bits)
            if tx == 0 and stride_t == stride_f:
                dst[(ty + y) * stride_f:(ty + y + 1) * stride_f] = row
            else:
                bit0 = tx * per * bits
                if bit0 & 7:
                    continue
                o = (ty + y) * stride_f + bit0 // 8
                n = min(stride_t, stride_f - bit0 // 8)
                dst[o:o + n] = row[:n]
    if len(offs) < n_expected:
        problems.append("TIFF：画像のデータの区画が足りません（足りない所は白か黒のままです）")
    dpi = _dpi(t)
    if jpeg_rgb is not None:
        rgb = bytes(jpeg_rgb)
    else:
        rgb = _to_rgb(full, w, h, spp, per, bits, photo, t.get(320), extra, notes)
    rgb, ow, oh = orient_rgb(rgb, w, h, orient)
    if orient >= 5 and dpi:
        dpi = (dpi[1], dpi[0])
    if orient != 1:
        notes.append(f"TIFF の向き（{orient}）のとおりに回しました")
    return Frame(ow, oh, rgb, dpi, notes, problems)


def _ccitt(raw: bytes, comp: int, t4: int, cols: int, rows: int) -> bytes:
    """CCITT：1 ビットの行（TIFF の標本）にする。符号の「白の連」は標本の 0（TIFF 6.0 10 章。見え方は Photometric で決まる）。"""
    from ..pdf.raster import RasterError, ccitt

    parm = {"Columns": cols, "Rows": rows}
    if comp == 4:
        parm["K"] = -1
    elif comp == 3:
        parm["K"] = 1 if t4 & 1 else 0
        parm["EncodedByteAlign"] = bool(t4 & 4)
    else:  # 2：行ごとにバイトの境（EOL なし）
        parm["K"] = 0
        parm["EncodedByteAlign"] = True
    try:
        g = ccitt(raw, parm)
    except RasterError as e:
        raise ImageError(f"TIFF：{e}") from None
    # 濃さ（白の連＝255・黒の連＝0）→ 標本のビット
    stride = (cols + 7) // 8
    out = bytearray(stride * rows)
    for y in range(min(rows, g.h)):
        row = g.px[y * cols:(y + 1) * cols]
        val = 0
        for i, v in enumerate(row):
            val = val << 1 | (0 if v else 1)
            if i & 7 == 7:
                out[y * stride + (i >> 3)] = val
                val = 0
        if cols & 7:
            out[y * stride + (cols >> 3)] = val << (8 - (cols & 7))
    return bytes(out)


def _jpeg_stream(blob: bytes, t: dict, raw: bytes, comp: int, photo: int) -> bytes:
    """TIFF の JPEG の区画 → 単独の JPEG の流れ（JPEGTables を前に足す。RGB の JPEG には Adobe の印で「変換なし」を伝える）。"""
    if comp == 6:
        at, ln = _one(t, 513), _one(t, 514)
        if at and ln:
            return blob[at:at + ln]
        raise ImageError("TIFF：古い JPEG（圧縮 6）で、JPEG の流れの位置がありません")
    tables = t.get(347)
    s = raw
    if isinstance(tables, (bytes, bytearray)) and len(tables) > 4:
        tb = bytes(tables)
        if tb.endswith(b"\xff\xd9"):
            tb = tb[:-2]
        s = tb + raw[2:] if raw[:2] == b"\xff\xd8" else tb + raw
    if photo == 2 and b"Adobe" not in s[:256]:
        s = s[:2] + b"\xff\xee\x00\x0eAdobe\x00d\x00\x00\x00\x00\x00" + s[2:]
    return s


def _to_rgb(full: list, w: int, h: int, spp: int, per: int, bits: int, photo: int, cmap, extra, notes: list) -> bytes:
    from .formats import _samples_row

    stride = (w * per * bits + 7) // 8
    planes_8: list = []
    if len(full) == 1:
        rows = [_samples_row(bytes(full[0][y * stride:(y + 1) * stride]), bits, w * per) for y in range(h)]
        flat = b"".join(rows)
        planes_8 = [flat[c::per] for c in range(per)]
    else:
        for p in full:
            rows = [_samples_row(bytes(p[y * stride:(y + 1) * stride]), bits, w) for y in range(h)]
            planes_8.append(b"".join(rows))
    out = bytearray(w * h * 3)
    maxv = 255 if bits >= 8 else (1 << bits) - 1
    scale = bytes(min(255, v * 255 // maxv) for v in range(256)) if bits < 8 else None
    alpha = None
    n_color = {0: 1, 1: 1, 3: 1, 2: 3, 5: 4, 6: 3, 8: 3}.get(photo, spp)
    if spp > n_color and planes_8 and extra and extra[0] in (1, 2) and len(planes_8) > n_color:
        alpha = planes_8[n_color]
    if photo in (0, 1):
        g = planes_8[0]
        if scale:
            g = g.translate(scale)
        if photo == 0:
            g = g.translate(bytes(255 - v for v in range(256)))
        out[0::3] = out[1::3] = out[2::3] = g
    elif photo == 3:
        if not cmap or len(cmap) < 3 * (1 << bits):
            raise ImageError("TIFF：パレットの表がありません")
        n = 1 << bits
        for k in range(3):
            tab = bytes(cmap[k * n + i] >> 8 for i in range(n)).ljust(256, b"\0")
            out[k::3] = planes_8[0].translate(tab)
    elif photo == 2:
        for k in range(3):
            p = planes_8[min(k, len(planes_8) - 1)]
            out[k::3] = p.translate(scale) if scale else p
    elif photo == 5:
        c, m, y_, k_ = planes_8[:4]
        for i, pl in enumerate((c, m, y_)):
            out[i::3] = bytes([(255 - a) * (255 - k) // 255 for a, k in zip(pl, k_)])
    elif photo == 6:
        Y, cb, cr = planes_8[:3]
        out[0::3] = bytes([min(255, max(0, int(a + 1.402 * (c - 128)))) for a, c in zip(Y, cr)])
        out[1::3] = bytes([min(255, max(0, int(a - 0.34414 * (b - 128) - 0.71414 * (c - 128)))) for a, b, c in zip(Y, cb, cr)])
        out[2::3] = bytes([min(255, max(0, int(a + 1.772 * (b - 128)))) for a, b in zip(Y, cb)])
    else:
        raise ImageError(f"TIFF：色の型（Photometric {photo}）は扱えません")
    rgb = bytes(out)
    if alpha is not None:
        rgb = _on_white(rgb, alpha)
        notes.append("透けのある画像を、白い紙の上の色で示しました")
    return rgb
