"""PDF の中の JPEG（DCTDecode）の色：フィルターの出す標本を、画像の Decode と色空間で色にする（ISO 32000-1 7.4.8・8.9.5）。

前は、Adobe の印のある 4 成分の JPEG を、PDF の中でも、復号のときに反転していた（JPEG のファイルの読み方）：Decode で反転を
指定していない画像・Decode [1 0 1 0 1 0 1 0] の画像のどちらかが、明るさの逆の絵になった（集めた PDF の 6,183 ページのうち 2 ページ）。
いまは、PDF の中では反転せず（YCCK は CMYK に直すだけ）、Decode と色空間で色にする。Adobe の印のない JPEG は、DecodeParms の
ColorTransform に従う。JPEG のファイルそのものの読み方は、前と同じ。
TIFF の中の JPEG の区画も、色の型のタグ（CMYK）が決める：Adobe の印があっても、反転しない（前は、反転して、どの色も黒にした）。

試験の JPEG は、ここで作る（どのブロックも DC だけの、8 × 8 画素の 1 色の絵）。画像は dpi=72：1 画素＝1 pt。
"""

from __future__ import annotations

import struct

import pdfmini
from test_pdf_render import _render, px
from todoku.imaging import tiff
from todoku.pdf.raster import jpeg_gray

RED_YCC = (76, 85, 255)  # 赤（255, 0, 0）の YCbCr


def _jpeg(samples: tuple, adobe: int | None = None) -> bytes:
    """成分ごとの値が samples の、8 × 8 画素の 1 色の基本の JPEG（T.81）。adobe：Adobe の印（APP14）の色の変換の値。"""
    seg = lambda m, b: b"\xff" + bytes([m]) + struct.pack(">H", len(b) + 2) + b
    n = len(samples)
    out = b"\xff\xd8"
    if adobe is not None:
        out += seg(0xEE, b"Adobe\x00\x64\x00\x00\x00\x00" + bytes([adobe]))
    out += seg(0xDB, b"\0" + bytes([8] * 64))  # 量子化はみな 8（DC は、画素の値 − 128 そのもの）
    out += seg(0xC0, b"\x08" + struct.pack(">HH", 8, 8) + bytes([n]) + b"".join(bytes([i + 1, 0x11, 0]) for i in range(n)))
    out += seg(0xC4, b"\x00" + bytes([0, 0, 0, 12] + [0] * 12) + bytes(range(12)))  # DC：分類 0〜11 を長さ 4 の符号に
    out += seg(0xC4, b"\x10" + bytes([1] + [0] * 15) + b"\x00")  # AC：終わりの印だけ
    out += seg(0xDA, bytes([n]) + b"".join(bytes([i + 1, 0]) for i in range(n)) + b"\x00\x3f\x00")
    bits = []
    for v in samples:
        dc = v - 128
        k = abs(dc).bit_length()
        bits.append(format(k, "04b"))
        if k:
            bits.append(format(dc if dc > 0 else dc + (1 << k) - 1, f"0{k}b"))
        bits.append("0")
    data = "".join(bits)
    data += "1" * (-len(data) % 8)
    return out + int(data, 2).to_bytes(len(data) // 8, "big").replace(b"\xff", b"\xff\x00") + b"\xff\xd9"


def _color(jpeg: bytes, space: str = "DeviceCMYK", extra: str = "") -> tuple:
    """JPEG を画像として置いたページの、画像の所の色。"""
    img = pdfmini.stream(jpeg, f"/Type /XObject /Subtype /Image /Width 8 /Height 8 /BitsPerComponent 8 /ColorSpace /{space} "
                         f"/Filter /DCTDecode {extra}")
    r = _render(pdfmini.build("q 200 0 0 200 100 500 cm /I Do Q", "<< /XObject << /I 2 0 R >> >>", [img]))
    assert not r.get("approx"), r.get("approx")
    return px(r, 200, 600)


def _is(c: tuple, want: str) -> bool:
    lo, hi = 40, 215
    r, g, b = c
    return {"red": r > hi and g < lo and b < lo, "black": max(c) < lo, "white": min(c) > hi,
            "cyan": r < lo and g > hi and b > hi}[want]


INVERT = "/Decode [1 0 1 0 1 0 1 0]"


def test_a_cmyk_jpeg_is_not_inverted_by_its_adobe_marker():
    """標本 (0, 255, 255, 0) は、C 0・M 1・Y 1・K 0 の赤（前は、Adobe の印を見て反転し、黒にした）。"""
    assert _is(_color(_jpeg((0, 255, 255, 0), adobe=0)), "red")
    assert _is(_color(_jpeg((0, 255, 255, 0))), "red")  # 印のない 4 成分：そのまま CMYK


def test_the_decode_array_inverts_a_cmyk_jpeg():
    """Photoshop の書く CMYK の JPEG：反転して入った標本を、Decode [1 0 …] で戻す。"""
    assert _is(_color(_jpeg((255, 0, 0, 255), adobe=0), extra=INVERT), "red")
    assert _is(_color(_jpeg((0, 255, 255, 0), adobe=0), extra=INVERT), "black")


def test_a_ycck_jpeg_becomes_cmyk_without_inversion():
    """YCCK（Adobe の印の変換 2）：YCC を RGB に直した補色が C・M・Y。K はそのまま。"""
    assert _is(_color(_jpeg((0, 128, 128, 0), adobe=2)), "black")  # RGB の黒 → C・M・Y がみな 1
    assert _is(_color(_jpeg((255, 128, 128, 0), adobe=2), extra=INVERT), "black")  # RGB の白 → 0 0 0 0 → Decode で反転 → みな 1
    c, m, y = (255 - v for v in (255, 0, 0))
    assert (c, m, y) == (0, 255, 255) and _is(_color(_jpeg((*RED_YCC, 0), adobe=2), extra="/Decode [1 0 1 0 1 0 0 1]"), "cyan")


def test_color_transform_of_the_filter_is_used_when_there_is_no_adobe_marker():
    ycc = "/DecodeParms << /ColorTransform 1 >>"
    # YCCK：YCC の赤 → C 0・M 255・Y 255、K 255 → Decode で反転 → C だけが 1
    assert _is(_color(_jpeg((*RED_YCC, 255)), extra=ycc + " " + INVERT), "cyan")
    assert not _is(_color(_jpeg((*RED_YCC, 255)), extra=INVERT), "cyan")  # 既定（4 成分は変換なし）：標本のまま CMYK
    # 3 成分：既定は YCC。ColorTransform 0 は、RGB の標本
    assert _is(_color(_jpeg(RED_YCC), "DeviceRGB"), "red")
    assert _is(_color(_jpeg((255, 0, 0)), "DeviceRGB", "/DecodeParms << /ColorTransform 0 >>"), "red")
    assert not _is(_color(_jpeg((255, 0, 0)), "DeviceRGB"), "red")
    # Adobe の印があれば、印が決める
    assert _is(_color(_jpeg(RED_YCC, adobe=1), "DeviceRGB", "/DecodeParms << /ColorTransform 0 >>"), "red")
    assert _is(_color(_jpeg(RED_YCC, adobe=1), "DeviceRGB", "/DecodeParms [<< /ColorTransform 0 >>]"), "red")


def test_the_decode_array_of_a_three_component_jpeg():
    assert _is(_color(_jpeg(RED_YCC), "DeviceRGB", "/Decode [1 0 1 0 1 0]"), "cyan")
    assert _is(_color(_jpeg(RED_YCC), "DeviceRGB", "/Decode [0 1 0 1 0 1]"), "red")


def test_a_jpeg_file_keeps_the_adobe_inversion():
    """JPEG のファイルそのもの（PDF の外）は、前と同じ：Adobe の印のある CMYK は、反転して入っている。"""
    im = jpeg_gray(_jpeg((255, 0, 0, 255), adobe=0), want_rgb=True)
    assert im.cmyk is None and _is(tuple(im.px[:3]), "red")
    im = jpeg_gray(_jpeg((0, 255, 255, 0), adobe=0), want_rgb=True)
    assert _is(tuple(im.px[:3]), "black")
    raw = jpeg_gray(_jpeg((0, 255, 255, 0), adobe=0), want_rgb=True, samples=True)
    assert raw.cmyk == bytes((0, 255, 255, 0)) * 64 and raw.px == b""


def _tiff(jpeg: bytes, photo: int = 5, n: int = 4) -> bytes:
    """8 × 8 画素の、JPEG で圧縮した区画 1 つの TIFF。"""
    tags = [(256, 3, 1, 8), (257, 3, 1, 8), (258, 3, n, None), (259, 3, 1, 7), (262, 3, 1, photo), (273, 4, 1, None),
            (277, 3, 1, n), (278, 3, 1, 8), (279, 4, 1, len(jpeg)), (284, 3, 1, 1)]
    bits_at = 8 + 2 + len(tags) * 12 + 4
    out = bytearray(b"II*\x00" + struct.pack("<IH", 8, len(tags)))
    for tag, typ, cnt, val in tags:
        v = struct.pack("<I", bits_at if tag == 258 else bits_at + 8) if val is None else \
            struct.pack("<HH", val, 0) if typ == 3 else struct.pack("<I", val)
        out += struct.pack("<HHI", tag, typ, cnt) + v
    return bytes(out + struct.pack("<I", 0) + struct.pack("<4H", 8, 8, 8, 8) + jpeg)


def test_a_cmyk_jpeg_strip_of_a_tiff_is_read_by_the_color_tag_of_the_tiff():
    """TIFF の CMYK は、0 が墨なし。JPEG の区画に Adobe の印があっても、ないときと同じ色（libtiff と同じ）。"""
    for adobe in (None, 0):
        f = tiff.decode(_tiff(_jpeg((0, 255, 255, 0), adobe=adobe)), 0)
        assert _is(tuple(f.rgb[:3]), "red"), adobe
        f = tiff.decode(_tiff(_jpeg((0, 0, 0, 0), adobe=adobe)), 0)
        assert _is(tuple(f.rgb[:3]), "white"), adobe
    f = tiff.decode(_tiff(_jpeg((255, 128, 128, 0), adobe=2)), 0)  # YCCK：YCC の白 → C・M・Y が 0
    assert _is(tuple(f.rgb[:3]), "white")
