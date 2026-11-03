"""画像を色（RGB）と透けで開き（ISO 32000-1 8.9・11.6.5）、ページの画素に写す（8.9.5 画像の座標：単位の四角）。

圧縮を開くのは raster.py（CCITT・JPEG・Flate など）。ここでは色空間と Decode・SMask・Mask（型抜き・色で隠す）・
ImageMask（塗りの色の型抜き）を扱い、写すときは、元の画像が画面の画素より細かければ箱の平均で縮めてから、
各画素の中心を画像に戻して引く（細い線を落とさない）。
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from operator import add

from ..objects import Name, Stream
from ..raster import _CS_ABBR, _FILTER_ABBR, _KEY_ABBR, RasterError, image_gray, jpeg_gray
from .canvas import Canvas, Clip, intersect, inv
from .colors import GRAY, colorspace


@dataclass
class RGBA:
    w: int
    h: int
    px: bytes  # w*h*3（R G B）。stencil のときは空
    alpha: bytes | None  # w*h（0＝透ける〜255＝不透明）
    stencil: bool = False  # 型抜き（ImageMask）：alpha の所を、塗りの色で塗る


def _image_dict(doc, image):
    if isinstance(image, Stream):
        return {k: doc.resolve(v) for k, v in image.dict.items()}, None
    if isinstance(image, (tuple, list)) and len(image) == 2:
        d = {_KEY_ABBR.get(k, k): doc.resolve(v) for k, v in image[0].items()}
        f = d.get("Filter")
        fl = f if isinstance(f, list) else [f] if f is not None else []
        d["Filter"] = [Name(_FILTER_ABBR.get(x.name, x.name)) if isinstance(x, Name) else x for x in fl]
        cs = d.get("ColorSpace")
        if isinstance(cs, Name):
            d["ColorSpace"] = Name(_CS_ABBR.get(cs.name, cs.name))
        elif isinstance(cs, list) and cs and isinstance(cs[0], Name):
            d["ColorSpace"] = [Name(_CS_ABBR.get(cs[0].name, cs[0].name))] + list(cs[1:])
        return d, image[1]
    raise RasterError("画像のデータがありません")


def _inv_table() -> bytes:
    return bytes(255 - i for i in range(256))


def image_rgba(doc, image, max_pixels: int = 16_000_000, res: dict | None = None, dct_pixels: int | None = None) -> RGBA:
    """PDF の画像（XObject の Stream か、埋め込みの (辞書, データ)）→ RGBA。"""
    d, inline = _image_dict(doc, image)
    flt = d.get("Filter")
    names = [getattr(x, "name", "") for x in (flt if isinstance(flt, list) else [flt] if flt is not None else [])]
    w, h = int(d.get("Width") or 0), int(d.get("Height") or 0)
    if w <= 0 or h <= 0:
        raise RasterError("画像の大きさが分かりません")
    if d.get("ImageMask") is True:
        g = image_gray(doc, image, max_pixels)  # 塗る所が 0（黒）
        return RGBA(g.w, g.h, b"", g.px.translate(_inv_table()), True)
    if "JBIG2Decode" in names or "JPXDecode" in names:
        raise RasterError(f"この画像の圧縮（{'JBIG2' if 'JBIG2Decode' in names else 'JPEG 2000'}）は、まだ開けません")
    cs = colorspace(doc, d.get("ColorSpace"), res)
    dec = d.get("Decode")
    dec = [float(doc.resolve(v)) for v in dec] if isinstance(dec, list) else None
    data = b""
    if "DCTDecode" in names:
        data = doc.decode(image) if isinstance(image, Stream) else doc.decode(Stream({"Filter": d["Filter"]}, inline), decrypt=False)
        im = jpeg_gray(data, min(max_pixels, dct_pixels or max_pixels), want_rgb=True)
        px = im.px
        if dec and len(dec) >= 2 and dec[0] > dec[1]:
            px = px.translate(_inv_table())
        if cs.n == 1 and cs.kind not in ("DeviceGray",):
            # 1 成分の色空間（Separation・Indexed など）の、濃さの JPEG
            tab = []
            for v in range(256):
                r, g, b = cs.rgb([v if cs.kind == "Indexed" else v / 255])
                tab.append(bytes((int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5))))
            px = b"".join(map(tab.__getitem__, px[0::3]))
        out = RGBA(im.w, im.h, px, None)
    elif "CCITTFaxDecode" in names:
        g = image_gray(doc, image, max_pixels)
        out = RGBA(g.w, g.h, _gray3(g.px), None)
    else:
        if isinstance(image, Stream):
            data = doc.decode(image)
        elif d.get("Filter"):
            data = doc.decode(Stream({"Filter": d["Filter"], "DecodeParms": d.get("DecodeParms")}, inline), decrypt=False)
        else:
            data = inline
        out = _raw_rgb(data, d, w, h, cs if cs.n else GRAY, dec, max_pixels)
    if isinstance(image, Stream):
        sm = doc.resolve(d.get("SMask"))
        mk = doc.resolve(d.get("Mask"))
        if isinstance(sm, Stream):
            a = image_gray(doc, sm, max(out.w * out.h, 1))
            out.alpha = _resample(a.px, a.w, a.h, out.w, out.h)
        elif isinstance(mk, Stream):
            a = image_gray(doc, mk, max(out.w * out.h, 1))  # 型抜き：隠す所（既定は 1）が白
            out.alpha = _resample(a.px.translate(_inv_table()), a.w, a.h, out.w, out.h)
        elif isinstance(mk, list) and len(mk) >= 2 * max(1, cs.n) and data and "DCTDecode" not in names:
            out.alpha = _color_key(data, d, w, h, cs, [int(doc.resolve(v)) for v in mk], out)
    return out


def _gray3(g: bytes) -> bytes:
    px = bytearray(len(g) * 3)
    px[0::3] = g
    px[1::3] = g
    px[2::3] = g
    return bytes(px)


def _resample(px: bytes, w: int, h: int, W: int, H: int) -> bytes:
    if (w, h) == (W, H):
        return bytes(px)
    cols = [min(w - 1, x * w // W) for x in range(W)]
    return b"".join(bytes(map(px.__getitem__, [min(h - 1, y * h // H) * w + c for c in cols])) for y in range(H))


_UNPACK: dict = {}


def _samples(row: bytes, bpc: int, n: int) -> list[int]:
    """1 行の標本（bpc ビット）を n 個の整数に（16 ビットは上位の 8 ビット）。"""
    if bpc == 8:
        return list(row[:n])
    if bpc == 16:
        return list(row[0:2 * n:2])
    maxv = (1 << bpc) - 1
    per = 8 // bpc
    tab = _UNPACK.get(bpc)
    if tab is None:
        tab = _UNPACK[bpc] = [tuple((b >> (8 - bpc * (k + 1))) & maxv for k in range(per)) for b in range(256)]
    out: list[int] = []
    for t in map(tab.__getitem__, row[:(n * bpc + 7) // 8]):
        out.extend(t)
    return out[:n]


def _raw_rgb(data: bytes, d: dict, w: int, h: int, cs, dec, max_pixels: int) -> RGBA:
    bpc = int(d.get("BitsPerComponent") or 8)
    if bpc not in (1, 2, 4, 8, 16):
        raise RasterError(f"画像の標本のビット数 {bpc} は扱えません")
    nc = cs.n
    step = max(1, math.isqrt(w * h // max(1, max_pixels)))
    while (-(-w // step)) * (-(-h // step)) > max_pixels:
        step += 1
    ow, oh = -(-w // step), -(-h // step)
    stride = (w * nc * bpc + 7) // 8
    maxv = 255 if bpc == 16 else (1 << bpc) - 1
    rng = cs.ranges()
    if not dec or len(dec) < 2 * nc:
        dec = [v for k in range(nc) for v in ((0.0, float(maxv)) if cs.kind == "Indexed" else rng[k])]
    out = bytearray()
    if nc == 1:
        tab = []
        for v in range(maxv + 1):
            x = dec[0] + v * (dec[1] - dec[0]) / maxv
            r, g, b = cs.rgb([round(x) if cs.kind == "Indexed" else x])
            tab.append(bytes((int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5))))
        for y in range(0, h, step):
            row = data[y * stride:(y + 1) * stride]
            if len(row) < stride:
                row = row + bytes(stride - len(row))
            vals = _samples(row, bpc, w)
            if step > 1:
                vals = vals[::step]
            out += b"".join(map(tab.__getitem__, vals))
        return RGBA(ow, oh, bytes(out), None)
    ident = all(abs(dec[2 * k] - rng[k][0]) < 1e-9 and abs(dec[2 * k + 1] - rng[k][1]) < 1e-9 for k in range(nc))
    cache: dict = {}
    for y in range(0, h, step):
        row = data[y * stride:(y + 1) * stride]
        if len(row) < stride:
            row = row + bytes(stride - len(row))
        vals = _samples(row, bpc, w * nc)
        if cs.kind == "DeviceRGB" and ident and bpc in (8, 16) and step == 1:
            out += bytes(vals)
            continue
        if cs.kind == "DeviceCMYK" and ident and bpc in (8, 16):
            c_, m_, y_, k_ = vals[0::nc * step], vals[1::nc * step], vals[2::nc * step], vals[3::nc * step]
            line = bytearray(len(k_) * 3)
            line[0::3] = bytes((255 - a) * (255 - k) // 255 for a, k in zip(c_, k_))
            line[1::3] = bytes((255 - a) * (255 - k) // 255 for a, k in zip(m_, k_))
            line[2::3] = bytes((255 - a) * (255 - k) // 255 for a, k in zip(y_, k_))
            out += line
            continue
        line = bytearray()
        for x in range(0, w, step):
            px = tuple(vals[x * nc:(x + 1) * nc])
            c = cache.get(px)
            if c is None:
                if cs.kind == "DeviceRGB" and ident:
                    c = bytes(v * 255 // maxv for v in px)
                else:
                    comps = [dec[2 * k] + px[k] * (dec[2 * k + 1] - dec[2 * k]) / maxv for k in range(nc)]
                    r, g, b = cs.rgb(comps)
                    c = bytes((int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5)))
                if len(cache) < 65536:
                    cache[px] = c
            line += c
        out += line
    return RGBA(ow, oh, bytes(out), None)


def _color_key(data: bytes, d: dict, w: int, h: int, cs, mk: list[int], out: RGBA) -> bytes | None:
    """色で隠す（Mask の配列、8.9.6.4）：どの成分も範囲の中の画素を隠す。"""
    bpc = int(d.get("BitsPerComponent") or 8)
    nc = cs.n
    if bpc > 8 or out.w != w or out.h != h:
        return None
    stride = (w * nc * bpc + 7) // 8
    alpha = bytearray()
    for y in range(h):
        vals = _samples(data[y * stride:(y + 1) * stride].ljust(stride, b"\0"), bpc, w * nc)
        for x in range(w):
            hid = all(mk[2 * k] <= vals[x * nc + k] <= mk[2 * k + 1] for k in range(nc))
            alpha.append(0 if hid else 255)
    return bytes(alpha)


def box_down(px: bytes, w: int, h: int, nc: int, kx: int, ky: int) -> tuple[bytes, int, int]:
    """画素を kx × ky の箱の平均で縮める（細い線を落とさない縮め方）。"""
    W, H = w // kx, h // ky
    if W < 1 or H < 1 or (kx == 1 and ky == 1):
        return px, w, h
    rowlen = w * nc
    kk = kx * ky
    out = bytearray(W * H * nc)
    div = [v // kk for v in range(256 * kk)]
    for Y in range(H):
        acc = None
        for r in range(ky):
            base = (Y * ky + r) * rowlen
            row = px[base:base + W * kx * nc]
            for i in range(kx):
                for c in range(nc):
                    part = row[i * nc + c::kx * nc]
                    if acc is None:
                        acc = [None] * nc
                    if acc[c] is None:
                        acc[c] = list(part)
                    else:
                        acc[c] = list(map(add, acc[c], part))
        o = Y * W * nc
        for c in range(nc):
            out[o + c:o + W * nc:nc] = bytes(map(div.__getitem__, acc[c]))
    return bytes(out), W, H


# ====================================================================== 写す
def draw_image(cv: Canvas, img: RGBA, m, clip: Clip, alpha: float = 1.0, fill=(0, 0, 0)) -> int:
    """画像（単位の四角 → 画面の画素の行列 m）を描く。描いた画素の数を返す。"""
    if clip.empty:
        return 0
    a, b, c, d, e, f = m
    # 画面の 1 画素あたりの、画像の画素（縮める割合）
    du = math.hypot(a, b)  # 画像の幅の向きの、画面の長さ
    dv = math.hypot(c, d)
    w, h = img.w, img.h
    px, al = img.px, img.alpha
    nc = 3
    kx = int(w / max(du, 1e-9)) if du > 0 else 1
    ky = int(h / max(dv, 1e-9)) if dv > 0 else 1
    if kx >= 2 or ky >= 2:
        kx, ky = max(1, kx), max(1, ky)
        if not img.stencil:
            px, w2, h2 = box_down(px, w, h, nc, kx, ky)
        if al is not None:
            al, w2, h2 = box_down(al, w, h, 1, kx, ky)
        w, h = w2, h2
    im = inv(m)
    if im is None or w <= 0 or h <= 0:
        return 0
    ia, ib, ic, id_, ie, if_ = im
    # 画像の四角の、画面の外枠
    corners = [(e, f), (a + e, b + f), (c + e, d + f), (a + c + e, b + d + f)]
    X0 = max(clip.x0, int(math.floor(min(p[0] for p in corners))))
    X1 = min(clip.x1, int(math.ceil(max(p[0] for p in corners))))
    Y0 = max(clip.y0, int(math.floor(min(p[1] for p in corners))))
    Y1 = min(clip.y1, int(math.ceil(max(p[1] for p in corners))))
    if X1 <= X0 or Y1 <= Y0:
        return 0
    buf = cv.buf
    W = cv.w
    mask = clip.mask
    fr, fg, fb = fill
    op = alpha >= 0.999
    n = 0
    axis = abs(ib) < 1e-12 and abs(ic) < 1e-12  # 回さない・傾けない画像
    for Y in range(Y0, Y1):
        yc = Y + 0.5
        # 行の中で u, v が [0,1) の x の範囲
        u0, v0 = ic * yc + ie, id_ * yc + if_  # x = 0 のとき（u = ia*x + ic*y + ie）
        lo, hi = X0 + 0.5, X1 - 0.5
        ok = True
        for k0, k1 in ((u0, ia), (v0, ib)):
            if abs(k1) < 1e-15:
                if not (0 <= k0 < 1):
                    ok = False
                    break
                continue
            t0, t1 = (0 - k0) / k1, (1 - k0) / k1
            if t0 > t1:
                t0, t1 = t1, t0
            lo, hi = max(lo, t0), min(hi, t1)
        if not ok or hi < lo:
            continue
        xs = max(X0, int(math.ceil(lo - 0.5)))
        xe = min(X1, int(math.floor(hi - 0.5)) + 1)
        if xe <= xs:
            continue
        spans = [(xs, xe, 255)]
        if mask is not None:
            mr = mask.get(Y)
            if not mr:
                continue
            spans = intersect(spans, mr)
        for sx0, sx1, ma in spans:
            if sx1 <= sx0:
                continue
            n += sx1 - sx0
            if axis:
                v = v0
                sy = int((1 - v) * h)
                sy = 0 if sy < 0 else h - 1 if sy >= h else sy
                cols = [int((ia * (x + 0.5) + u0) * w) for x in range(sx0, sx1)]
                cols = [0 if q < 0 else w - 1 if q >= w else q for q in cols]
                idx = [sy * w + q for q in cols]
            else:
                idx = []
                for x in range(sx0, sx1):
                    xc = x + 0.5
                    u, v = ia * xc + u0, ib * xc + v0
                    qx, qy = int(u * w), int((1 - v) * h)
                    qx = 0 if qx < 0 else w - 1 if qx >= w else qx
                    qy = 0 if qy < 0 else h - 1 if qy >= h else qy
                    idx.append(qy * w + qx)
            p0 = (Y * W + sx0) * 3
            if img.stencil:
                avals = list(map(al.__getitem__, idx))
                for j, av in enumerate(avals):
                    av = av * ma // 255
                    if not op:
                        av = int(av * alpha + 0.5)
                    if av <= 0:
                        continue
                    p = p0 + 3 * j
                    na = 255 - av
                    buf[p] = (buf[p] * na + fr * av + 127) // 255
                    buf[p + 1] = (buf[p + 1] * na + fg * av + 127) // 255
                    buf[p + 2] = (buf[p + 2] * na + fb * av + 127) // 255
                continue
            idx3 = [k for q in idx for k in (3 * q, 3 * q + 1, 3 * q + 2)]
            row = bytes(map(px.__getitem__, idx3))
            if al is None and ma >= 255 and op:
                buf[p0:p0 + len(row)] = row
                continue
            avals = list(map(al.__getitem__, idx)) if al is not None else [255] * len(idx)
            for j, av in enumerate(avals):
                if ma < 255:
                    av = av * ma // 255
                if not op:
                    av = int(av * alpha + 0.5)
                if av <= 0:
                    continue
                p = p0 + 3 * j
                q = 3 * j
                if av >= 255:
                    buf[p:p + 3] = row[q:q + 3]
                    continue
                na = 255 - av
                buf[p] = (buf[p] * na + row[q] * av + 127) // 255
                buf[p + 1] = (buf[p + 1] * na + row[q + 1] * av + 127) // 255
                buf[p + 2] = (buf[p + 2] * na + row[q + 2] * av + 127) // 255
    return n
