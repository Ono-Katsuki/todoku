"""画像を色（RGB）と透けで開き（ISO 32000-1 8.9・11.6.5）、ページの画素に写す（8.9.5 画像の座標：単位の四角）。

圧縮を開くのは raster.py（CCITT・JPEG・JBIG2・JPEG 2000・Flate など）。ここでは色空間と Decode・SMask・Mask（型抜き・色で隠す）・
ImageMask（塗りの色の型抜き）を扱い、写すときは、元の画像が画面の画素より細かければ面積の平均で画面の画素の数に縮めてから、
各画素の中心を画像に戻して引く（細い線を落とさない）。マスクは画像と別の解像度のことがあるので、細かいほうに合わせる。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from itertools import repeat
from operator import add, itemgetter, mul, rshift

from ...limits import budget
from ..objects import Name, Stream
from ..raster import _CS_ABBR, _FILTER_ABBR, _KEY_ABBR, RasterError, _thin_step, image_gray, jpeg_gray
from .canvas import Canvas, Clip, intersect, inv
from .colors import DCMYK, GRAY, colorspace


@dataclass
class RGBA:
    w: int
    h: int
    px: bytes  # w*h*3（R G B）。stencil のときは空
    alpha: bytes | None  # w*h（0＝透ける〜255＝不透明）
    stencil: bool = False  # 型抜き（ImageMask）：alpha の所を、塗りの色で塗る
    notes: list = field(default_factory=list)  # 開き方の知らせ（解像度を下げたなど）
    problems: list = field(default_factory=list)  # 描けなかった所
    scaled: tuple | None = None  # 画面の大きさに縮めたもの：((幅, 高さ), 画素, 透け)。同じ画像を同じ大きさで何度も描くとき、縮め直さない
    window: tuple | None = None  # 開いた範囲（元の画像の割合 [左, 上, 右, 下)。1 が右・下の端）。None は全体


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


def _color_transform(doc, d: dict, names: list) -> int | None:
    """DCTDecode の DecodeParms の ColorTransform（なければ None：既定は、3 成分が 1、4 成分が 0）。"""
    parms = doc.resolve(d.get("DecodeParms"))
    if isinstance(parms, list):
        i = names.index("DCTDecode")
        parms = doc.resolve(parms[i]) if i < len(parms) else None
    v = doc.resolve(parms.get("ColorTransform")) if isinstance(parms, dict) else None
    return int(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def _decode_changes(dec, cs) -> bool:
    """Decode が、色空間の既定の範囲と違う。"""
    if not dec or len(dec) < 2 * cs.n:
        return False
    rng = cs.ranges()
    return any(abs(dec[2 * k] - rng[k][0]) > 1e-9 or abs(dec[2 * k + 1] - rng[k][1]) > 1e-9 for k in range(cs.n))


def _shrink(g, screen_pixels: int | None):
    """画面より十分大きな白黒の画像（JBIG2・CCITT の 600 dpi のスキャンなど）は、色の面に広げる前に箱の平均で縮める（メモリ）。"""
    if not screen_pixels or g.w * g.h <= 4 * screen_pixels:
        return g.px, g.w, g.h
    k = math.isqrt(g.w * g.h // screen_pixels)
    return box_down(g.px, g.w, g.h, 1, k, k)


def image_rgba(doc, image, max_pixels: int = 16_000_000, res: dict | None = None, dct_pixels: int | None = None,
               deadline: float | None = None, view: tuple | None = None) -> RGBA:
    """PDF の画像（XObject の Stream か、埋め込みの (辞書, データ)）→ RGBA。
    view：画面に見える範囲（画像の割合 [左, 上, 右, 下)）。JPEG 2000・JPEG は、そこだけを開いて RGBA.window で知らせる（拡大して読むとき、
    予算は開く分に掛かるので、元の解像度まで開ける）。SMask・Mask は全体を開いて、同じ割合で切り取る（解像度は画像と別）。"""
    d, inline = _image_dict(doc, image)
    flt = d.get("Filter")
    names = [getattr(x, "name", "") for x in (flt if isinstance(flt, list) else [flt] if flt is not None else [])]
    w, h = int(d.get("Width") or 0), int(d.get("Height") or 0)
    if (w <= 0 or h <= 0) and "JPXDecode" not in names:
        raise RasterError("画像の大きさが分かりません")
    if d.get("ImageMask") is True:
        g = image_gray(doc, image, max_pixels, deadline)  # 塗る所が 0（黒）
        return RGBA(g.w, g.h, b"", g.px.translate(_inv_table()), True, g.notes, g.problems)
    if "JPXDecode" in names:
        return _jpx_rgba(doc, image, d, res, max_pixels, dct_pixels, deadline, view)
    cs = colorspace(doc, d.get("ColorSpace"), res)
    dec = d.get("Decode")
    dec = [float(doc.resolve(v)) for v in dec] if isinstance(dec, list) else None
    data = b""
    if "DCTDecode" in names:
        data = doc.decode(image) if isinstance(image, Stream) else doc.decode(Stream({"Filter": d["Filter"]}, inline), decrypt=False)
        dw, dh = w, h
        im = jpeg_gray(data, min(max_pixels, dct_pixels or max_pixels), want_rgb=True,
                       window=(int(view[0] * dw), int(view[1] * dh), math.ceil(view[2] * dw), math.ceil(view[3] * dh)) if view else None,
                       samples=True, transform=_color_transform(doc, d, names))
        px = im.px
        if im.cmyk is not None:
            # 4 成分の JPEG：フィルターの出す標本を、ほかの画像と同じに、Decode と画像の色空間で色にする（JPEG の印では反転しない）
            px = _raw_rgb(im.cmyk, {}, im.w, im.h, cs if cs.n == 4 else DCMYK, dec, im.w * im.h).px
        elif cs.n == 3 and (cs.kind != "DeviceRGB" or _decode_changes(dec, cs)):
            # 3 成分の JPEG で、色空間が RGB でない（Lab など）か、Decode が既定でない：標本として色にする
            px = _raw_rgb(px, {}, im.w, im.h, cs, dec, im.w * im.h).px
        elif cs.n != 3 and dec and len(dec) >= 2 and dec[0] > dec[1]:
            px = px.translate(_inv_table())
        if cs.n == 1 and cs.kind not in ("DeviceGray",):
            # 1 成分の色空間（Separation・Indexed など）の、濃さの JPEG
            tab = []
            for v in range(256):
                r, g, b = cs.rgb([v if cs.kind == "Indexed" else v / 255])
                tab.append(bytes((int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5))))
            px = b"".join(map(tab.__getitem__, px[0::3]))
        done = getattr(im, "window", None)
        out = RGBA(im.w, im.h, px, None, False, [], [], window=(done[0] / dw, done[1] / dh, done[2] / dw, done[3] / dh) if done else None)
    elif "CCITTFaxDecode" in names or "JBIG2Decode" in names:
        g = image_gray(doc, image, max_pixels, deadline)
        px, gw, gh = _shrink(g, dct_pixels)
        if cs.n == 1 and cs.kind not in ("DeviceGray", "CalGray", "ICCBased"):
            # Indexed・Separation の 1 ビットの画像：0 と 1 の色で塗る（濃さは 0 か 255）
            c0, c1 = [cs.rgb([v]) for v in (0, 1)]
            tab = [bytes(int((c0[k] + (c1[k] - c0[k]) * v / 255) * 255 + 0.5) for k in range(3)) for v in range(256)]
            rgb = b"".join(map(tab.__getitem__, px))
        else:
            rgb = _gray3(px)
        out = RGBA(gw, gh, rgb, None, False, g.notes, g.problems)
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
            a = _cropped(image_gray(doc, sm, max_pixels, deadline), out.window)
            _set_alpha(out, a.px, a.w, a.h, max_pixels, dct_pixels)
            matte = doc.resolve(sm.dict.get("Matte"))
            if isinstance(matte, list) and len(matte) == cs.n and (a.w, a.h) == (out.w, out.h):
                _unmatte(out, cs.rgb([float(doc.resolve(v)) for v in matte]))
        elif isinstance(mk, Stream):
            a = _cropped(image_gray(doc, mk, max_pixels, deadline), out.window)  # 型抜き：隠す所（既定は 1）が白
            _set_alpha(out, a.px.translate(_inv_table()), a.w, a.h, max_pixels, dct_pixels)
        elif isinstance(mk, list) and len(mk) >= 2 * max(1, cs.n) and data and "DCTDecode" not in names:
            out.alpha = _color_key(data, d, w, h, cs, [int(doc.resolve(v)) for v in mk], out)
    return out


def _cropped(g, win: tuple | None):
    """マスクを、画像の開いた範囲（割合）と同じ所に切り取る。"""
    if win is None:
        return g
    x0, y0 = int(win[0] * g.w), int(win[1] * g.h)
    x1, y1 = max(x0 + 1, min(g.w, math.ceil(win[2] * g.w))), max(y0 + 1, min(g.h, math.ceil(win[3] * g.h)))
    from ..raster import Gray

    return Gray(x1 - x0, y1 - y0, b"".join(g.px[y * g.w + x0:y * g.w + x1] for y in range(y0, y1)))


def _jpx_rgba(doc, image, d: dict, res, max_pixels: int, screen_pixels: int | None, deadline: float | None, view: tuple | None = None) -> RGBA:
    """JPXDecode（ISO 32000-1 7.4.9）：ColorSpace があれば JP2 の色の指定より優先、なければ JP2 の箱で決める。Decode は無視する。
    SMaskInData（表 89）：1・2 なら、JP2 の透けの成分を透けとして使う。SMask の Stream があれば、そちらを使う。"""
    from ..raster import image_gray as _ig
    from ..raster import jpx_image

    if isinstance(image, Stream):
        data = doc.decode(image)
    else:
        data = doc.decode(Stream({"Filter": d["Filter"]}, image[1]), decrypt=False)
    target = min(max_pixels, screen_pixels or max_pixels)
    # 画面の 2 倍の screen_pixels のうち、解像度の段は画面の画素を下回らないところまで下げる
    win = None
    if view is not None:
        from ...imaging import jpx as _jpx

        head = _jpx.info(data)
        win = (int(view[0] * head["w"]), int(view[1] * head["h"]), math.ceil(view[2] * head["w"]), math.ceil(view[3] * head["h"]))
    im = jpx_image(doc, data, max(1, target // 2), deadline, win)
    planes = im.planes
    n = len(planes)
    cs_obj = d.get("ColorSpace")
    if cs_obj is not None:
        cs = colorspace(doc, cs_obj, res)
    else:
        name = {"gray": "DeviceGray", "sRGB": "DeviceRGB", "CMYK": "DeviceCMYK"}.get(im.colorspace or "")
        if name is None:
            name = "DeviceGray" if n < 3 else "DeviceCMYK" if n == 4 and im.colorspace == "CMYK" else "DeviceRGB"
        cs = colorspace(doc, Name(name), res)
    nc = cs.n or 1
    if n < nc:
        planes = list(planes) + [planes[-1]] * (nc - n)
    if cs.kind == "Indexed" and im.bits and im.bits[0] < 8:
        top = (1 << im.bits[0]) - 1  # 面は 0〜255 に広げてあるので、添字に戻す
        planes = [planes[0].translate(bytes((v * top + 127) // 255 for v in range(256)))]
    px = _planes_rgb(planes[:nc], cs, im.w, im.h)
    fw_, fh_ = im.full_w, im.full_h
    out = RGBA(im.w, im.h, px, None, False, list(im.notes), list(im.problems),
               window=(im.window[0] / fw_, im.window[1] / fh_, im.window[2] / fw_, im.window[3] / fh_) if im.window else None)
    smd = doc.resolve(d.get("SMaskInData")) if isinstance(image, Stream) else 0
    if isinstance(image, Stream):
        sm = doc.resolve(d.get("SMask"))
        if isinstance(sm, Stream):
            a = _cropped(_ig(doc, sm, max_pixels, deadline), out.window)
            _set_alpha(out, a.px, a.w, a.h, max_pixels, screen_pixels)
        elif smd in (1, 2) and im.alpha is not None:
            out.alpha = bytes(im.alpha)
        else:
            mk = doc.resolve(d.get("Mask"))
            if isinstance(mk, Stream):
                a = _cropped(_ig(doc, mk, max_pixels, deadline), out.window)
                _set_alpha(out, a.px.translate(_inv_table()), a.w, a.h, max_pixels, screen_pixels)
    return out


def _planes_rgb(planes: list, cs, w: int, h: int) -> bytes:
    """8 ビットの成分の面 → RGB（w*h*3）。"""
    nc = len(planes)
    if nc == 1:
        p = planes[0]
        if cs.kind in ("DeviceGray", "CalGray") or (cs.kind == "ICCBased" and cs.n == 1):
            return _gray3(p)
        tab = []
        for v in range(256):
            r, g, b = cs.rgb([v if cs.kind == "Indexed" else v / 255])
            tab.append(bytes((int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5))))
        return b"".join(map(tab.__getitem__, p))
    if nc == 3 and cs.kind in ("DeviceRGB", "CalRGB", "ICCBased"):
        out = bytearray(w * h * 3)
        out[0::3], out[1::3], out[2::3] = planes[0], planes[1], planes[2]
        return bytes(out)
    if nc == 4 and cs.kind in ("DeviceCMYK", "ICCBased"):
        out = bytearray(w * h * 3)
        k_ = planes[3]
        for i in range(3):
            out[i::3] = bytes([(255 - a) * (255 - k) // 255 for a, k in zip(planes[i], k_)])
        return bytes(out)
    cache: dict = {}
    out = bytearray()
    rng = cs.ranges()
    for vals in zip(*planes):
        c = cache.get(vals)
        if c is None:
            r, g, b = cs.rgb([rng[i][0] + (rng[i][1] - rng[i][0]) * v / 255 for i, v in enumerate(vals)])
            c = bytes((int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5)))
            if len(cache) < 65536:
                cache[vals] = c
        out += c
    return bytes(out)


def _set_alpha(out: RGBA, a: bytes, aw: int, ah: int, max_pixels: int, screen_pixels: int | None) -> None:
    """透けのマスク（SMask・型抜きの Mask）を画像に付ける。マスクの解像度は画像と別に決まる（8.9.6.3・11.6.5.3）：
    2×2 の黒の画像に細かなマスクで字の形を持たせる書類もあるので、細かいほうに合わせる（粗いほうへ間引くと字が消える）。
    画面より大きなマスクだけ、画面の大きさの数倍まで箱の平均で縮める。"""
    limit = min(max_pixels, screen_pixels or max_pixels)
    if aw * ah > limit:
        k = math.isqrt(aw * ah // max(1, limit)) + 1
        a, aw, ah = box_down(a, aw, ah, 1, k, k)
    W, H = max(out.w, aw), max(out.h, ah)
    out.alpha = _resample(a, aw, ah, W, H)
    if (W, H) != (out.w, out.h):
        out.px = _resample(out.px, out.w, out.h, W, H, 3)
        out.w, out.h = W, H


def _unmatte(out: RGBA, matte) -> None:
    """Matte（11.6.5.3）：画素の色は、マスクで Matte の色と混ぜてある。元の色に戻す：c = m + (c′ − m) / α。"""
    mr, mg, mb = (int(v * 255 + 0.5) for v in matte)
    px, al = bytearray(out.px), out.alpha
    for i, av in enumerate(al):
        if av in (0, 255):
            continue
        p = i * 3
        px[p] = max(0, min(255, mr + (px[p] - mr) * 255 // av))
        px[p + 1] = max(0, min(255, mg + (px[p + 1] - mg) * 255 // av))
        px[p + 2] = max(0, min(255, mb + (px[p + 2] - mb) * 255 // av))
    out.px = bytes(px)


def _gray3(g: bytes) -> bytes:
    px = bytearray(len(g) * 3)
    px[0::3] = g
    px[1::3] = g
    px[2::3] = g
    return bytes(px)


def _resample(px: bytes, w: int, h: int, W: int, H: int, nc: int = 1) -> bytes:
    """いちばん近い画素で W × H に（nc：1 画素の成分の数）。"""
    if (w, h) == (W, H):
        return bytes(px)
    budget().count(W * nc, "画像を合わせる列の数")
    cols = [min(w - 1, x * w // W) * nc + k for x in range(W) for k in range(nc)]
    rows: dict[int, bytes] = {}  # 広げるときは同じ元の行が続く
    out = []
    for y in range(H):
        sy = min(h - 1, y * h // H)
        row = rows.get(sy)
        if row is None:
            base = sy * w * nc
            rows = {sy: bytes(map(px.__getitem__, [base + c for c in cols]))}
            row = rows[sy]
        out.append(row)
    return b"".join(out)


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
    step = _thin_step(w, h, max_pixels)
    ow, oh = -(-w // step), -(-h // step)
    stride = (w * nc * bpc + 7) // 8
    # 宣言の幅でなく、実際にあるデータの分だけ読む（描けるだけ描く。足りない所は 0 の標本の色で、場所は出力の画素の分だけ取る）
    w_read = min(w, len(data) * 8 // max(1, nc * bpc))
    stride_read = (w_read * nc * bpc + 7) // 8
    ow_read = -(-w_read // step)
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
        pad = tab[0] * (ow - ow_read)
        for y in range(0, h, step):
            if y * stride >= len(data):
                out += tab[0] * (ow * ((h - y + step - 1) // step))
                break
            row = data[y * stride:y * stride + stride_read]
            if len(row) < stride_read:
                row = row + bytes(stride_read - len(row))
            vals = _samples(row, bpc, w_read)
            if step > 1:
                vals = vals[::step]
            out += b"".join(map(tab.__getitem__, vals)) + pad
        return RGBA(ow, oh, bytes(out), None)
    ident = all(abs(dec[2 * k] - rng[k][0]) < 1e-9 and abs(dec[2 * k + 1] - rng[k][1]) < 1e-9 for k in range(nc))
    cache: dict = {}
    pad = b""
    if ow_read < ow or len(data) < stride * h:
        comps = [dec[2 * k] for k in range(nc)]  # 標本 0 の色
        r, g, b = cs.rgb(comps)
        pad = bytes((int(r * 255 + 0.5), int(g * 255 + 0.5), int(b * 255 + 0.5)))
    for y in range(0, h, step):
        if y * stride >= len(data):
            out += pad * (ow * ((h - y + step - 1) // step))
            break
        row = data[y * stride:y * stride + stride_read]
        if len(row) < stride_read:
            row = row + bytes(stride_read - len(row))
        vals = _samples(row, bpc, w_read * nc)
        if cs.kind == "DeviceRGB" and ident and bpc in (8, 16) and step == 1:
            out += bytes(vals) + pad * (ow - ow_read)
            continue
        if cs.kind == "DeviceCMYK" and ident and bpc in (8, 16):
            c_, m_, y_, k_ = vals[0::nc * step], vals[1::nc * step], vals[2::nc * step], vals[3::nc * step]
            line = bytearray(len(k_) * 3)
            line[0::3] = bytes((255 - a) * (255 - k) // 255 for a, k in zip(c_, k_))
            line[1::3] = bytes((255 - a) * (255 - k) // 255 for a, k in zip(m_, k_))
            line[2::3] = bytes((255 - a) * (255 - k) // 255 for a, k in zip(y_, k_))
            out += line + pad * (ow - ow_read)
            continue
        line = bytearray()
        for x in range(0, w_read, step):
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
        out += line + pad * (ow - ow_read)
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
    div = [v // kk for v in range(256 * kk)] if kk <= 1024 else None  # 表は小さな箱だけ（大きな箱で表を作ると、箱の広さ × 256 の場所を取る）
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
            out[o + c:o + W * nc:nc] = bytes(map(div.__getitem__, acc[c])) if div is not None else bytes([v // kk for v in acc[c]])
    return bytes(out), W, H


def _taps(n: int, N: int) -> list[tuple[int, int, int, int]]:
    """n 個を N 個に縮める（倍率は 1〜2）：出力の 1 個が覆う元の 3 個までの (番号, 重み 0, 重み 1, 重み 2)。重みの和は 256。"""
    r = n / N
    out = []
    for j in range(N):
        s, e = j * r, (j + 1) * r
        i = min(int(s), n - 1)
        a = (min(e, i + 1) - s) / r
        c = max(0.0, e - (i + 2)) / r
        w0, w2 = int(a * 256 + 0.5), int(c * 256 + 0.5)
        out.append((i, w0, 256 - w0 - w2, w2))
    return out


def area_down(px: bytes, w: int, h: int, nc: int, W: int, H: int) -> bytes:
    """画素の平均で W × H に縮める（W ≤ w、H ≤ h）：まず整数の倍率の箱の平均、残りの 1〜2 倍を面積の重みで。"""
    if (w, h) == (W, H):
        return px
    if nc == 3:
        g = px[0::3]
        if g == px[1::3] and g == px[2::3]:  # 灰色（スキャンの多く）は 1 成分で
            return _gray3(area_down(g, w, h, 1, W, H))
    kx, ky = max(1, w // W), max(1, h // H)
    if kx > 1 or ky > 1:
        px, w, h = box_down(px, w, h, nc, kx, ky)
    if (w, h) == (W, H):
        return px
    # 横：列ごとの重み
    rows: dict[int, bytes] = {}
    if w != W:
        tx = _taps(w, W)
        ix = [[min(w - 1, t[0] + d) * nc + c for t in tx for c in range(nc)] for d in range(3)]
        wx = [[t[1 + d] for t in tx for _c in range(nc)] for d in range(3)]
        get = [itemgetter(*v) if len(v) > 1 else (lambda r, q=v[0]: (r[q],)) for v in ix]

    def line(y: int) -> bytes:
        row = rows.get(y)
        if row is None:
            row = px[y * w * nc:(y + 1) * w * nc]
            if w != W:
                a, b, c = (g_(row) for g_ in get)
                row = bytes(map(rshift, map(add, map(add, map(mul, a, wx[0]), map(mul, b, wx[1])), map(mul, c, wx[2])), repeat(8)))
            rows[y] = row
        return row

    if h == H:
        return b"".join(line(y) for y in range(H))
    out = []
    for i, w0, w1, w2 in _taps(h, H):
        a, b = line(i), line(min(h - 1, i + 1))
        if w2:
            c = line(min(h - 1, i + 2))
            r = map(add, map(add, map(mul, a, repeat(w0)), map(mul, b, repeat(w1))), map(mul, c, repeat(w2)))
        else:
            r = map(add, map(mul, a, repeat(w0)), map(mul, b, repeat(w1)))
        out.append(bytes(map(rshift, r, repeat(8))))
        for y in [y for y in rows if y < i]:
            del rows[y]
    return b"".join(out)


# ====================================================================== 写す
def draw_image(cv: Canvas, img: RGBA, m, clip: Clip, alpha: float = 1.0, fill=(0, 0, 0), blend=None) -> int:
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
    W_t, H_t = max(1, math.ceil(du - 1e-6)), max(1, math.ceil(dv - 1e-6))
    if W_t < w or H_t < h:
        # 画面より細かい画像：面積の平均で、ちょうど画面の画素の数に縮める（整数の倍率でないと、間引きで細い線が切れる）
        W_t, H_t = min(W_t, w), min(H_t, h)
        if img.scaled is not None and img.scaled[0] == (W_t, H_t):
            px, al = img.scaled[1], img.scaled[2]
        else:
            if not img.stencil:
                px = area_down(px, w, h, nc, W_t, H_t)
            if al is not None:
                al = area_down(al, w, h, 1, W_t, H_t)
            img.scaled = ((W_t, H_t), px, al)
        w, h = W_t, H_t
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
    cover = cv.cover
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
            q0 = Y * W + sx0
            p0 = q0 * 3
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
                    cr, cg, cb = fr, fg, fb
                    if blend is not None:  # 重ね方（BM）：下の色と重ねた色を、濃さで混ぜる
                        cr, cg, cb = blend.onto(buf[p], buf[p + 1], buf[p + 2], 255 if cover is None else cover[q0 + j], fr, fg, fb)
                    buf[p] = (buf[p] * na + cr * av + 127) // 255
                    buf[p + 1] = (buf[p + 1] * na + cg * av + 127) // 255
                    buf[p + 2] = (buf[p + 2] * na + cb * av + 127) // 255
                    if cover is not None and cover[q0 + j] < 255:
                        cover[q0 + j] += ((255 - cover[q0 + j]) * av + 127) // 255
                continue
            idx3 = [k for q in idx for k in (3 * q, 3 * q + 1, 3 * q + 2)]
            row = bytes(map(px.__getitem__, idx3))
            if al is None and ma >= 255 and op and blend is None:
                buf[p0:p0 + len(row)] = row
                if cover is not None:
                    cover[q0:q0 + len(idx)] = b"\xff" * len(idx)
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
                if blend is not None:
                    cr, cg, cb = blend.onto(buf[p], buf[p + 1], buf[p + 2], 255 if cover is None else cover[q0 + j], row[q], row[q + 1], row[q + 2])
                    na = 255 - av
                    buf[p] = (buf[p] * na + cr * av + 127) // 255
                    buf[p + 1] = (buf[p + 1] * na + cg * av + 127) // 255
                    buf[p + 2] = (buf[p + 2] * na + cb * av + 127) // 255
                    if cover is not None and cover[q0 + j] < 255:
                        cover[q0 + j] += ((255 - cover[q0 + j]) * av + 127) // 255
                    continue
                if cover is not None and cover[q0 + j] < 255:
                    cover[q0 + j] += ((255 - cover[q0 + j]) * av + 127) // 255
                if av >= 255:
                    buf[p:p + 3] = row[q:q + 3]
                    continue
                na = 255 - av
                buf[p] = (buf[p] * na + row[q] * av + 127) // 255
                buf[p + 1] = (buf[p + 1] * na + row[q + 1] * av + 127) // 255
                buf[p + 2] = (buf[p + 2] * na + row[q + 2] * av + 127) // 255
    return n
