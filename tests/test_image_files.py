"""単独の画像ファイル（JPEG・PNG・GIF・BMP・TIFF・WebP）：画素が元と合うこと。Pillow・cwebp があれば別の実装の画像で確かめる（なければ飛ばす）。"""

from __future__ import annotations

import base64
import io
import shutil
import struct
import subprocess
import zlib

import pytest

from todoku import api
from todoku.grade import grade_blob
from todoku.imaging import formats as F
from todoku.imaging import standalone as S

try:
    from PIL import Image as PILImage
except ImportError:  # pragma: no cover
    PILImage = None
needs_pil = pytest.mark.skipif(PILImage is None, reason="Pillow がない（試験の画像を作るだけに使う）")


def _rgb(w=64, h=48):
    out = bytearray()
    for y in range(h):
        for x in range(w):
            out += bytes(((x * 4) & 255, (y * 5) & 255, ((x ^ y) * 3) & 255))
    return bytes(out)


def _gray(w=64, h=48):
    return bytes(((x * 3 + y * 2) & 255) for y in range(h) for x in range(w))


def _bits_img(w=64, h=48):
    return PILImage.frombytes("L", (w, h), bytes(0 if (x // 3 + y // 2) % 3 == 0 else 255 for y in range(h) for x in range(w))).convert("1", dither=PILImage.Dither.NONE)


# ====================================================================== 手で作る（外部の道具なし）
def _png(w, h, ct, depth, raw_rows, extra=b"", interlace=0):
    def ch(t, b):
        return struct.pack(">I", len(b)) + t + b + struct.pack(">I", zlib.crc32(t + b) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + ch(b"IHDR", struct.pack(">IIBBBBB", w, h, depth, ct, 0, 0, interlace)) + extra
            + ch(b"IDAT", zlib.compress(raw_rows)) + ch(b"IEND", b""))


def test_png_all_filters_exact():
    w, h = 8, 6
    src = [bytes((x * 10 + y * 5) & 255 for x in range(w * 3)) for y in range(h)]

    def filt(ft, cur, prev):
        out = bytearray()
        for i, v in enumerate(cur):
            a = cur[i - 3] if i >= 3 else 0
            b = prev[i]
            c = prev[i - 3] if i >= 3 else 0
            if ft == 0:
                p = 0
            elif ft == 1:
                p = a
            elif ft == 2:
                p = b
            elif ft == 3:
                p = (a + b) // 2
            else:
                q = a + b - c
                pa, pb, pc = abs(q - a), abs(q - b), abs(q - c)
                p = a if pa <= pb and pa <= pc else b if pb <= pc else c
            out.append((v - p) & 255)
        return bytes([ft]) + bytes(out)

    raw = b""
    prev = bytes(w * 3)
    for y, r in enumerate(src):
        raw += filt(y % 5, r, prev)
        prev = r
    fr = F.decode(_png(w, h, 2, 8, raw))
    assert fr.rgb == b"".join(src) and (fr.w, fr.h) == (w, h)


def test_png_palette_transparency_bits_and_phys():
    plte = struct.pack(">I", 6) + b"PLTE" + bytes([255, 0, 0, 0, 255, 0]) + struct.pack(">I", zlib.crc32(b"PLTE" + bytes([255, 0, 0, 0, 255, 0])) & 0xFFFFFFFF)
    trns = struct.pack(">I", 1) + b"tRNS" + b"\x00" + struct.pack(">I", zlib.crc32(b"tRNS\x00") & 0xFFFFFFFF)
    phys_b = struct.pack(">IIB", 11811, 11811, 1)
    phys = struct.pack(">I", 9) + b"pHYs" + phys_b + struct.pack(">I", zlib.crc32(b"pHYs" + phys_b) & 0xFFFFFFFF)
    raw = b"\0" + bytes([0b01000000])  # 2 画素：0（透け）・1（緑）
    fr = F.decode(_png(2, 1, 3, 1, raw, phys + plte + trns))
    assert fr.rgb == bytes([255, 255, 255, 0, 255, 0])  # 透けは白の上
    assert round(fr.dpi[0]) == 300 and fr.notes


def test_png_gray16_and_rgba():
    raw = b"\0" + struct.pack(">HH", 0x1234, 0xFF00)
    assert F.decode(_png(2, 1, 0, 16, raw)).rgb == bytes([0x12] * 3 + [0xFF] * 3)
    raw = b"\0" + bytes([10, 20, 30, 255, 10, 20, 30, 0])
    assert F.decode(_png(2, 1, 6, 8, raw)).rgb == bytes([10, 20, 30, 255, 255, 255])


def test_png_adam7_interlace():
    # 2×2 の濃さ：Adam7 の 1・6・7 番目の通し
    raw = b"\0" + bytes([1]) + b"\0" + bytes([2]) + b"\0" + bytes([3, 4])
    fr = F.decode(_png(2, 2, 0, 8, raw, interlace=1))
    assert fr.rgb[0::3] == bytes([1, 2, 3, 4])


def _gif(w, h, indexes, palette, min_code=2, interlace=False):
    # LZW：クリアの符号を入れて、辞書を増やさない素直な流れ
    bits = ""
    clear = 1 << min_code
    size = min_code + 1
    codes = [clear]
    for i, v in enumerate(indexes):
        codes.append(v)
        if (i + 1) % 2 == 0:
            codes.append(clear)
    codes.append(clear + 1)
    out = bytearray()
    acc = nb = 0
    for c in codes:
        acc |= c << nb
        nb += size
        while nb >= 8:
            out.append(acc & 255)
            acc >>= 8
            nb -= 8
    if nb:
        out.append(acc & 255)
    _ = bits
    head = b"GIF89a" + struct.pack("<HHBBB", w, h, 0x80 | 1, 0, 0) + palette
    img = b"," + struct.pack("<HHHHB", 0, 0, w, h, 0x40 if interlace else 0) + bytes([min_code])
    blocks = b""
    for i in range(0, len(out), 255):
        part = bytes(out[i:i + 255])
        blocks += bytes([len(part)]) + part
    return head + img + blocks + b"\0;"


def test_gif_lzw_and_palette():
    pal = bytes([255, 0, 0, 0, 255, 0, 0, 0, 255, 9, 9, 9])
    g = _gif(2, 2, [0, 1, 2, 3], pal[:12] if False else pal[:6] + pal[6:12])
    # 色の表の大きさの印は 2 << 1 = 4 色
    fr = F.decode(g)
    assert fr.rgb == bytes([255, 0, 0, 0, 255, 0, 0, 0, 255, 9, 9, 9])


def _bmp24(w, h, rows_rgb):
    stride = (w * 3 + 3) // 4 * 4
    body = b""
    for r in reversed(rows_rgb):
        line = bytearray()
        for i in range(0, len(r), 3):
            line += bytes((r[i + 2], r[i + 1], r[i]))
        body += bytes(line).ljust(stride, b"\0")
    dib = struct.pack("<IiiHHIIiiII", 40, w, h, 1, 24, 0, len(body), 11811, 11811, 0, 0)
    return b"BM" + struct.pack("<IHHI", 14 + 40 + len(body), 0, 0, 54) + dib + body


def test_bmp_24bit_bottom_up_and_dpi():
    rows = [bytes([1, 2, 3, 4, 5, 6]), bytes([7, 8, 9, 10, 11, 12])]
    fr = F.decode(_bmp24(2, 2, rows))
    assert fr.rgb == rows[0] + rows[1] and round(fr.dpi[0]) == 300


def test_bmp_rle8_and_palette():
    pal = bytes([0, 0, 255, 0, 0, 255, 0, 0])  # 0 赤・1 緑（BGRA）
    body = bytes([2, 0, 0, 0, 2, 1, 0, 1])  # 下の行：赤 2、行末、上の行：緑 2、終わり
    dib = struct.pack("<IiiHHIIiiII", 40, 2, 2, 1, 8, 1, len(body), 0, 0, 2, 0)
    off = 14 + 40 + len(pal)
    blob = b"BM" + struct.pack("<IHHI", off + len(body), 0, 0, off) + dib + pal + body
    fr = F.decode(blob)
    assert fr.rgb == bytes([0, 255, 0, 0, 255, 0, 255, 0, 0, 255, 0, 0])
    assert fr.dpi is None


def _tiff(w, h, data, comp=1, photo=1, spp=1, bits=8, extra=None, strips=None):
    tags = {256: (4, [w]), 257: (4, [h]), 258: (3, [bits] * spp), 259: (3, [comp]), 262: (3, [photo]), 277: (3, [spp]), 278: (4, [h]),
            282: (5, (300, 1)), 283: (5, (300, 1)), 296: (3, [2])}
    tags.update(extra or {})
    strips = strips or [data]
    n = len(tags) + 2
    base = 8 + 2 + 12 * n + 4
    blob_after = bytearray()
    offs = []
    p = base
    for s in strips:
        offs.append(p + len(blob_after))
        blob_after += s
    tags[273] = (4, offs)
    tags[279] = (4, [len(s) for s in strips])
    out = bytearray(b"II*\0" + struct.pack("<I", 8) + struct.pack("<H", n))
    pool = bytearray()
    pool_base = base + len(blob_after)
    for tag in sorted(tags):
        typ, val = tags[tag]
        if typ == 5:
            enc = struct.pack("<II", *val)
            cnt = 1
        else:
            fmt = {3: "H", 4: "I"}[typ]
            enc = struct.pack(f"<{len(val)}{fmt}", *val)
            cnt = len(val)
        if len(enc) <= 4:
            field = enc.ljust(4, b"\0")
        else:
            field = struct.pack("<I", pool_base + len(pool))
            pool += enc
        out += struct.pack("<HHI", tag, typ, cnt) + field
    out += struct.pack("<I", 0)
    return bytes(out) + bytes(blob_after) + bytes(pool)


def test_tiff_uncompressed_gray_and_dpi():
    d = _gray(16, 8)
    fr = F.decode(_tiff(16, 8, d))
    assert fr.rgb[0::3] == d and fr.dpi == (300.0, 300.0)


def test_tiff_packbits():
    # 8 画素：リテラル 3、繰り返し 5
    d = bytes([2, 1, 2, 3, 256 - 4, 9])
    fr = F.decode(_tiff(8, 1, d, comp=32773))
    assert fr.rgb[0::3] == bytes([1, 2, 3, 9, 9, 9, 9, 9])


def test_tiff_deflate_with_predictor():
    d = bytes([10, 1, 1, 1, 20, 2, 2, 2])
    fr = F.decode(_tiff(4, 2, zlib.compress(d), comp=8, extra={317: (3, [2])}, strips=[zlib.compress(d)]))
    assert fr.rgb[0::3] == bytes([10, 11, 12, 13, 20, 22, 24, 26])


def test_tiff_white_is_zero_one_bit():
    fr = F.decode(_tiff(8, 1, bytes([0b10100000]), photo=0, bits=1))
    assert fr.rgb[0::3] == bytes([0, 255, 0, 255, 255, 255, 255, 255])


def test_tiff_palette():
    cmap = [0x1000, 0xF000] + [0x2000, 0xE000] + [0x3000, 0xD000]
    fr = F.decode(_tiff(2, 1, bytes([0b01000000]), photo=3, bits=1, extra={320: (3, cmap)}))
    assert fr.rgb == bytes([0x10, 0x20, 0x30, 0xF0, 0xE0, 0xD0])


# ====================================================================== 別の実装の画像（Pillow）
@needs_pil
@pytest.mark.parametrize("kw", [{}, {"progressive": True}, {"subsampling": 0}, {"quality": 50, "progressive": True}])
def test_jpeg_baseline_and_progressive_close_to_pillow(kw):
    im = PILImage.frombytes("RGB", (64, 48), _rgb())
    b = io.BytesIO()
    im.save(b, "JPEG", quality=kw.pop("quality", 90), **kw)
    fr = F.decode(b.getvalue())
    ref = PILImage.open(io.BytesIO(b.getvalue())).convert("RGB").tobytes()
    assert (fr.w, fr.h) == (64, 48)
    # 色の間引き（4:2:0）の広げ方は実装で違う（todoku は近い画素、Pillow はなめらかに）ので、平均の差で比べる
    assert sum(abs(a - c) for a, c in zip(fr.rgb, ref)) / len(ref) <= 8


@needs_pil
def test_jpeg_gray_progressive_and_dpi_and_exif_orientation():
    im = PILImage.frombytes("L", (40, 20), _gray(40, 20))
    b = io.BytesIO()
    im.save(b, "JPEG", progressive=True, dpi=(200, 200))
    fr = F.decode(b.getvalue())
    assert (fr.w, fr.h) == (40, 20) and round(fr.dpi[0]) == 200
    # EXIF の向き 6：横長の画像を 90 度右へ
    ex = PILImage.Exif()
    ex[0x0112] = 6
    b2 = io.BytesIO()
    PILImage.frombytes("RGB", (64, 48), _rgb()).save(b2, "JPEG", exif=ex.tobytes(), quality=95)
    fi = F.info(b2.getvalue())[1][0]
    assert (fi.w, fi.h) == (48, 64)
    fr2 = F.decode(b2.getvalue())
    assert (fr2.w, fr2.h) == (48, 64)
    # PIL の回転と同じ向き
    ref = PILImage.open(io.BytesIO(b2.getvalue())).convert("RGB").transpose(PILImage.Transpose.ROTATE_270).tobytes()
    assert sum(abs(a - c) for a, c in zip(fr2.rgb, ref)) / len(ref) <= 8


@needs_pil
def test_gif_from_pillow_with_frames_and_interlace():
    a = PILImage.frombytes("RGB", (32, 16), _rgb(32, 16)).convert("P", palette=PILImage.Palette.ADAPTIVE, colors=64)
    c = PILImage.frombytes("RGB", (32, 16), _rgb(32, 16)[::-1]).convert("P", palette=PILImage.Palette.ADAPTIVE, colors=64)
    b = io.BytesIO()
    a.save(b, "GIF", save_all=True, append_images=[c], duration=100, disposal=1)
    blob = b.getvalue()
    assert len(F.info(blob)[1]) == 2
    b1 = PILImage.open(io.BytesIO(blob))
    assert F.decode(blob, 0).rgb == b1.convert("RGB").tobytes()
    b1.seek(1)
    assert F.decode(blob, 1).rgb == b1.convert("RGB").tobytes()


@needs_pil
@pytest.mark.parametrize("mode", ["1", "L", "P", "RGB", "RGBA", "I;16"])
def test_png_modes_from_pillow(mode):
    base = PILImage.frombytes("RGB", (37, 23), _rgb(37, 23))
    if mode == "1":
        im = _bits_img(37, 23)
    elif mode == "I;16":
        im = base.convert("L").convert("I;16")
    else:
        im = base.convert(mode)
    b = io.BytesIO()
    im.save(b, "PNG")
    fr = F.decode(b.getvalue())
    ref = PILImage.open(io.BytesIO(b.getvalue()))
    if mode == "RGBA":
        bg = PILImage.new("RGB", ref.size, (255, 255, 255))
        bg.paste(ref, mask=ref.split()[3])
        ref_px = bg.tobytes()
    elif mode == "I;16":
        return  # Pillow の I;16 の濃さの丸めは別（大きさと開けることだけ）
    else:
        ref_px = ref.convert("RGB").tobytes()
    assert fr.rgb == ref_px


@needs_pil
@pytest.mark.parametrize("mode,bits", [("1", 1), ("P", 8), ("RGB", 24)])
def test_bmp_from_pillow(mode, bits):
    base = PILImage.frombytes("RGB", (33, 21), _rgb(33, 21))
    im = _bits_img(33, 21) if mode == "1" else base.convert(mode)
    b = io.BytesIO()
    im.save(b, "BMP", dpi=(150, 150))
    fr = F.decode(b.getvalue())
    assert fr.rgb == PILImage.open(io.BytesIO(b.getvalue())).convert("RGB").tobytes()
    assert round(fr.dpi[0]) == 150


@needs_pil
@pytest.mark.parametrize("comp", ["raw", "packbits", "tiff_lzw", "tiff_adobe_deflate", "group4", "group3", "tiff_ccitt", "jpeg"])
def test_tiff_compressions_from_pillow(comp):
    base = PILImage.frombytes("RGB", (97, 61), _rgb(97, 61))
    im = _bits_img(97, 61) if comp in ("group4", "group3", "tiff_ccitt") else base
    b = io.BytesIO()
    try:
        im.save(b, "TIFF", compression=comp, dpi=(300, 300))
    except (OSError, KeyError, ValueError):
        pytest.skip(f"Pillow に {comp} がない")
    fr = F.decode(b.getvalue())
    ref = PILImage.open(io.BytesIO(b.getvalue())).convert("RGB").tobytes()
    worst = max(abs(a - c) for a, c in zip(fr.rgb, ref))
    assert worst <= (16 if comp == "jpeg" else 0), worst
    assert round(fr.dpi[0]) == 300


@needs_pil
def test_tiff_multipage_and_tiled_rgb():
    a = PILImage.frombytes("RGB", (40, 30), _rgb(40, 30))
    c = PILImage.frombytes("RGB", (40, 30), _rgb(40, 30)[::-1])
    b = io.BytesIO()
    a.save(b, "TIFF", save_all=True, append_images=[c], compression="tiff_lzw")
    blob = b.getvalue()
    assert len(F.info(blob)[1]) == 2
    assert F.decode(blob, 1).rgb == c.tobytes()


@pytest.mark.skipif(not shutil.which("tiffcp"), reason="tiffcp がない")
@needs_pil
def test_tiff_tiles_from_libtiff(tmp_path):
    a = PILImage.frombytes("RGB", (70, 50), _rgb(70, 50))
    src = tmp_path / "a.tif"
    a.save(src, "TIFF")
    out = tmp_path / "b.tif"
    r = subprocess.run(["tiffcp", "-t", "-w", "32", "-l", "32", "-c", "lzw:2", str(src), str(out)], capture_output=True, timeout=30)
    assert r.returncode == 0, r.stderr
    assert F.decode(out.read_bytes()).rgb == a.tobytes()


# ====================================================================== WebP
@needs_pil
@pytest.mark.parametrize("mode", ["RGB", "RGBA", "L"])
def test_webp_lossless_from_pillow(mode):
    base = PILImage.frombytes("RGB", (61, 43), _rgb(61, 43))
    im = base.convert(mode)
    if mode == "RGBA":
        im.putalpha(PILImage.frombytes("L", (61, 43), bytes(255 if x % 5 else 128 for y in range(43) for x in range(61))))
    b = io.BytesIO()
    try:
        im.save(b, "WEBP", lossless=True)
    except (OSError, KeyError):
        pytest.skip("Pillow に WebP がない")
    fr = F.decode(b.getvalue())
    ref = PILImage.open(io.BytesIO(b.getvalue()))
    if mode == "RGBA":
        bg = PILImage.new("RGB", ref.size, (255, 255, 255))
        bg.paste(ref, mask=ref.split()[3])
        want = bg.tobytes()
    else:
        want = ref.convert("RGB").tobytes()
    assert (fr.w, fr.h) == (61, 43)
    if mode == "RGBA":
        assert max(abs(x - y) for x, y in zip(fr.rgb, want)) <= 1
    else:
        assert fr.rgb == want


@needs_pil
def test_webp_lossless_palette_image():
    # 色の少ない画像は、色の索引の変換と、色のキャッシュを使う
    px = bytes(c for y in range(40) for x in range(50) for c in ((255, 0, 0) if (x + y) % 7 == 0 else (0, 0, 255) if x % 3 else (20, 200, 30)))
    im = PILImage.frombytes("RGB", (50, 40), px)
    b = io.BytesIO()
    im.save(b, "WEBP", lossless=True, method=6)
    assert F.decode(b.getvalue()).rgb == px


# ====================================================================== 書類として
def _png_file(tmp_path, name="a.png", dpi=True):
    phys = b""
    if dpi:
        b_ = struct.pack(">IIB", 11811, 11811, 1)
        phys = struct.pack(">I", 9) + b"pHYs" + b_ + struct.pack(">I", zlib.crc32(b"pHYs" + b_) & 0xFFFFFFFF)
    raw = b"".join(b"\0" + bytes([(x + y) & 255] * 3) * 1 for y in range(0) for x in range(0))
    w, h = 300, 200
    raw = b"".join(b"\0" + b"".join(bytes(((x * y) & 255,) * 3) for x in range(w)) for y in range(h))
    p = tmp_path / name
    p.write_bytes(_png(w, h, 2, 8, raw, phys))
    return p


def test_image_opens_as_a_document(tmp_path):
    p = _png_file(tmp_path)
    k, doc, parts = api.window(str(p))
    assert k == "image" and doc["pages"][0]["width_px"] == 300 and doc["pages"][0]["dpi"]
    assert "字の情報はありません" in parts[0][1] and "view_page" in parts[0][1] and "to_pdf" in parts[0][1]
    s = api.summary(str(p))
    assert s["kind"] == "image" and s["pages"] == 1 and s["fields"] == 0


def test_image_grade_is_read_e_write_d_with_to_pdf_reason(tmp_path):
    p = _png_file(tmp_path)
    g = grade_blob(p.read_bytes(), p.name)
    assert g.grade == "E" and g.write == "D"
    assert "view_page" in g.reasons[0].detail and "to_pdf" in g.write_reasons[0].detail


def test_view_page_of_image_has_mm_ruler_when_dpi_known(tmp_path):
    p = _png_file(tmp_path)
    v = api.view_page(str(p), 1)
    assert "mm" in v["text"] and "300 dpi" in v["text"]
    assert v["png_base64"] and v["width"] == 300 + 26
    # 範囲を拡大（mm）
    v2 = api.view_page(str(p), 1, region=[0, 0, 10, 10])
    assert v2["width"] > 200


def test_view_page_of_image_without_dpi_has_pixel_ruler_and_refuses_dpi(tmp_path):
    p = _png_file(tmp_path, dpi=False)
    v = api.view_page(str(p), 1)
    assert "画素" in v["text"] and "紙の上の mm は分かりません" in v["text"]
    with pytest.raises(api.ApiError):
        api.view_page(str(p), 1, dpi=200)
    with pytest.raises(api.ApiError):
        api.view_page(str(p), 2)


def test_fill_on_image_is_refused_with_the_way_out(tmp_path):
    p = _png_file(tmp_path)
    with pytest.raises(api.ApiError) as e:
        api.fill(str(p), {"@1": {"page": 1, "box": [1, 1, 5, 5], "value": "x"}}, str(tmp_path / "o.png"))
    assert "to_pdf" in str(e.value)


def test_to_pdf_then_write_on_the_pdf(tmp_path):
    p = _png_file(tmp_path)
    out = tmp_path / "wrapped.pdf"
    r = api.to_pdf(str(p), str(out))
    assert out.exists() and r["pages"][0]["size_mm"] == [25.4, 16.9] and not r["pages"][0]["dpi_assumed"]
    assert api.grade(str(out))["grade"] == "E"
    filled = tmp_path / "filled.pdf"
    res = api.fill(str(out), {"@1": {"page": 1, "box": [2, 2, 20, 8], "value": "ABC"}}, str(filled))
    assert res["proof_ok"]
    with pytest.raises(api.ApiError):
        api.to_pdf(str(p), str(out))  # すでにある


def test_to_pdf_without_dpi_says_it_assumed_a4(tmp_path):
    p = _png_file(tmp_path, dpi=False)
    r = api.to_pdf(str(p), str(tmp_path / "x.pdf"))
    assert r["pages"][0]["dpi_assumed"] and "A4" in r["say_to_user"]


@needs_pil
def test_to_pdf_keeps_jpeg_bytes(tmp_path):
    b = io.BytesIO()
    PILImage.frombytes("RGB", (64, 48), _rgb()).save(b, "JPEG", dpi=(100, 100))
    p = tmp_path / "a.jpg"
    p.write_bytes(b.getvalue())
    out = tmp_path / "j.pdf"
    api.to_pdf(str(p), str(out))
    assert b.getvalue() in out.read_bytes()


def test_audit_of_image_reports_ok(tmp_path):
    p = _png_file(tmp_path)
    d = api.audit(str(p))
    assert d["ok"] and d["grade"]["grade"] == "E"


def test_broken_heic_says_the_exact_conversion_never_pretends(tmp_path):
    p = tmp_path / "a.heic"
    p.write_bytes(struct.pack(">I", 24) + b"ftypheic" + b"\0\0\0\0mif1heic" + b"\0" * 64)
    with pytest.raises(api.ApiError) as e:
        api.window(str(p))
    assert "sips -s format jpeg" in str(e.value)
    g = api.grade(str(p))
    assert g["grade"] == "F" and "sips -s format jpeg" in g["reasons"][0]["detail"]


def test_png_bomb_is_stopped_by_declared_size():
    from todoku.limits import LimitExceeded

    blob = _png(200000, 200000, 2, 8, b"")
    with pytest.raises(LimitExceeded):
        F.decode(blob)


def test_view_image_png_base64_is_real_png(tmp_path):
    p = _png_file(tmp_path)
    v = S.view_frame(p.read_bytes(), 1)
    assert base64.b64decode(v["png_base64"])[:8] == b"\x89PNG\r\n\x1a\n"


# ====================================================================== WebP 非可逆（VP8）：dwebp -yuv と同じ画素
CWEBP, DWEBP = shutil.which("cwebp"), shutil.which("dwebp")
needs_webp = pytest.mark.skipif(not (CWEBP and DWEBP and PILImage), reason="cwebp・dwebp・Pillow がない")


@needs_webp
@pytest.mark.parametrize("args", [["-q", "75"], ["-q", "20", "-f", "60", "-sharpness", "3"], ["-q", "90", "-segments", "4", "-sns", "90"],
                                  ["-q", "50", "-nostrong"], ["-q", "50", "-strong"], ["-q", "40", "-f", "0"], ["-q", "70", "-segments", "1", "-sns", "0"]])
def test_vp8_lossy_decodes_exact_yuv(args, tmp_path):
    from todoku.imaging.webp import _frame_size, _parse
    from todoku.imaging.webp_vp8 import decode_yuv

    w, h = 83, 61
    # 写真のような濃淡と字のような線
    im = PILImage.frombytes("RGB", (w, h), bytes(c for y in range(h) for x in range(w)
                                                  for c in ((x * 3 + (y * y) // 7) & 255, (y * 4 + (x ^ y)) & 255, ((x * y) // 3) & 255)))
    src = tmp_path / "a.png"
    im.save(src)
    out = tmp_path / "a.webp"
    r = subprocess.run([CWEBP, "-quiet", *args, str(src), "-o", str(out)], capture_output=True, timeout=30)
    assert r.returncode == 0, r.stderr
    blob = out.read_bytes()
    yuv = tmp_path / "a.yuv"
    r = subprocess.run([DWEBP, "-quiet", str(out), "-yuv", "-o", str(yuv)], capture_output=True, timeout=30)
    assert r.returncode == 0, r.stderr
    ref = yuv.read_bytes()
    f = _parse(blob)["frames"][0]
    assert _frame_size(blob, f) == (w, h)
    Y, U, V, sy, suv = decode_yuv(blob[f["s"]:f["e"]], w, h)
    my = b"".join(bytes(Y[i * sy:i * sy + w]) for i in range(h))
    cw, ch = (w + 1) // 2, (h + 1) // 2
    mu = b"".join(bytes(U[i * suv:i * suv + cw]) for i in range(ch))
    mv = b"".join(bytes(V[i * suv:i * suv + cw]) for i in range(ch))
    assert my == ref[:w * h]
    assert mu == ref[w * h:w * h + cw * ch]
    assert mv == ref[w * h + cw * ch:]


@needs_webp
def test_vp8_rgb_close_to_pillow_and_opens_as_document(tmp_path):
    w, h = 60, 40
    im = PILImage.frombytes("RGB", (w, h), _rgb(w, h))
    b = io.BytesIO()
    im.save(b, "WEBP", quality=90)
    p = tmp_path / "a.webp"
    p.write_bytes(b.getvalue())
    fr = F.decode(b.getvalue())
    ref = PILImage.open(io.BytesIO(b.getvalue())).convert("RGB").tobytes()
    assert sum(abs(x - y) for x, y in zip(fr.rgb, ref)) / len(ref) <= 8
    k, _doc, _parts = api.window(str(p))
    assert k == "image"
    assert api.view_page(str(p), 1)["png_base64"]


@needs_webp
def test_vp8_with_alpha_chunk(tmp_path):
    im = PILImage.frombytes("RGB", (40, 30), _rgb(40, 30)).convert("RGBA")
    im.putalpha(PILImage.frombytes("L", (40, 30), bytes(255 if x < 20 else 0 for y in range(30) for x in range(40))))
    b = io.BytesIO()
    im.save(b, "WEBP", quality=90)
    fr = F.decode(b.getvalue())
    assert fr.notes and fr.rgb[3 * 39:3 * 40] == b"\xff\xff\xff"


# ====================================================================== JPEG 2000・JBIG2 のファイル、縦横の解像度の違う画像
@pytest.mark.skipif(not shutil.which("opj_compress"), reason="opj_compress がない")
def test_jp2_file_opens_as_a_document_and_view_page_reports_resolution(tmp_path):
    w, h = 400, 300
    pgm = tmp_path / "a.pgm"
    pgm.write_bytes(f"P5\n{w} {h}\n255\n".encode() + _gray(w, h))
    out = tmp_path / "a.jp2"
    r = subprocess.run(["opj_compress", "-i", str(pgm), "-o", str(out), "-n", "3"], capture_output=True, timeout=30)
    assert r.returncode == 0, r.stderr
    k, doc, _parts = api.window(str(out))
    assert k == "image" and doc["format"] == "jp2" and doc["pages"][0]["width_px"] == w
    v = api.view_page(str(out), 1)
    assert v["width"] == w + 26 and v["not_drawn"] == []
    assert api.grade(str(out))["grade"] == "E"


def test_jbig2_file_with_file_header_opens():
    import jbig2_encoder as E

    rows = [[(x // 3 + y) % 2 for x in range(32)] for y in range(16)]
    body = E.segment(0, 48, E.page_info(32, 16)) + E.segment(1, 38, E.generic_seg(rows))
    blob = b"\x97JB2\r\n\x1a\n" + bytes([1]) + (1).to_bytes(4, "big") + body
    fmt, frames = F.info(blob)
    assert fmt == "jbig2" and (frames[0].w, frames[0].h) == (32, 16)
    fr = F.decode(blob)
    assert fr.rgb[0::3] == bytes(0 if v else 255 for r in rows for v in r)


def test_fax_resolution_is_shown_in_paper_shape_with_a_true_ruler():
    # 204×98 dpi：紙の上では 1 画素が縦に 2 倍。目盛りの 1 mm の画素が縦横で同じになる
    d = _gray(204, 98)
    tags = {282: (5, (204, 1)), 283: (5, (98, 1))}
    blob = _tiff(204, 98, d, extra=tags)
    v = S.view_frame(blob, 1)
    assert "25×25 mm" in v["text"] or "26×25 mm" in v["text"]
    w, h = v["width"] - 26, v["height"] - 26
    assert abs(w - h) <= 2, (w, h)
