"""ページの画像（view_page）が、見る道具と同じように描くか：図形・切り抜き・色・画像・字（埋め込んだ TrueType・CFF・
Type 1・型 3）・回したページ・注釈の見た目・見えない内容。書体は試験の中で手で組み立てる（このパソコンの書体に頼らない）。

画像は dpi=72 で描く：1 画素＝1 pt。ページの座標 (x, y) は、画像の (x, 842 − y)。
"""

from __future__ import annotations

import struct
import zlib

import pytest

import pdfmini
from todoku.pdf.doc import PdfDocument
from todoku.pdf.pageimage import render, view_page
from todoku.pdf.render import canvas, fontprog


def _render(blob: bytes, **kw) -> dict:
    return render(PdfDocument(blob, "t.pdf"), 1, dpi=72, **kw)


def px(r: dict, x: float, y: float) -> tuple[int, int, int]:
    """ページの座標 (x, y) の画素の色。"""
    k = r["scale"] / (72 / 25.4)
    X, Y = int(x * k + 1e-6), int((842 - y) * k + 1e-6)
    i = (Y * r["w"] + X) * 3
    return tuple(r["rgb"][i:i + 3])


WHITE, BLACK = (255, 255, 255), (0, 0, 0)


def dark(c) -> bool:
    return sum(c) < 200


# ====================================================================== 図形
def test_fill_rules_colors_and_clip():
    c = ("1 0 0 rg 50 700 100 80 re f "                                          # 赤い四角
         "0 0 1 rg 300 600 m 400 600 l 400 700 l 300 700 l h 320 620 m 380 620 l 380 680 l 320 680 l h f* "  # 偶奇：穴
         "0 1 0 rg 300 400 m 400 400 l 400 500 l 300 500 l h 320 420 m 380 420 l 380 480 l 320 480 l h f "  # 非ゼロ：穴なし
         "q 50 300 50 50 re W n 0 0 0 1 k 0 0 595 842 re f Q "                     # 切り抜き（CMYK の黒）
         "q 0.5 g 450 100 m 450 160 500 200 550 200 c 550 100 l h f Q")             # 曲線（灰色）
    r = _render(pdfmini.build(c, "<< >>", []))
    assert px(r, 100, 740) == (255, 0, 0)
    assert px(r, 310, 610) == (0, 0, 255) and px(r, 350, 650) == WHITE  # 偶奇の穴
    assert px(r, 350, 450) == (0, 255, 0)  # 非ゼロ：同じ向きの内側も塗る
    assert px(r, 75, 325) == BLACK and px(r, 120, 325) == WHITE and px(r, 75, 400) == WHITE  # 切り抜きの外は塗らない
    assert px(r, 540, 150) == (128, 128, 128) and px(r, 455, 190) == WHITE
    # 縁はなめらか（アンチエイリアス）：半分だけ覆う画素は中間の色
    r2 = _render(pdfmini.build("0 g 100.5 700 50 50 re f", "<< >>", []))
    v = px(r2, 100, 720)[0]
    assert 100 < v < 160, v


def test_strokes_width_caps_joins_and_dash():
    c = ("0 g 10 w 0 J 100 700 m 200 700 l S "   # 突き出さない端
         "2 J 100 600 m 200 600 l S "             # 四角の端：5 だけ突き出す
         "1 J 100 500 m 200 500 l S "             # 丸い端
         "3 w 0 J [10 10] 0 d 100 400 m 300 400 l S [] 0 d "  # 破線
         "0.1 w 100 300 m 300 300 l S")            # 細い線も 1 画素で見える
    r = _render(pdfmini.build(c, "<< >>", []))
    assert dark(px(r, 150, 703)) and not dark(px(r, 150, 708))  # 太さ 10（上下に 5）
    assert not dark(px(r, 97, 700)) and dark(px(r, 101, 700))
    assert dark(px(r, 97, 600)) and dark(px(r, 203, 600)) and not dark(px(r, 207, 600))
    assert dark(px(r, 97, 500)) and not dark(px(r, 96, 504.5))  # 丸い端：角は欠ける
    assert dark(px(r, 105, 400)) and not dark(px(r, 115, 400)) and dark(px(r, 125, 400))
    assert px(r, 200, 300)[0] < 255


def test_line_join_miter_and_bevel():
    r = _render(pdfmini.build("0 g 20 w 0 j 100 700 m 150 780 l 200 700 l S "
                              "2 j 300 700 m 350 780 l 400 700 l S", "<< >>", []))
    # とがり：頂点より上（150, 795）まで黒い。切り落とし：上は欠ける
    assert dark(px(r, 150, 795)) and not dark(px(r, 350, 795))


# ====================================================================== 画像
def test_images_rgb_stencil_and_smask():
    rgb = bytes([255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0])  # 2×2：赤 緑 / 青 黄
    alpha = bytes([255, 0, 255, 255])
    objs = [pdfmini.stream(zlib.compress(rgb), "/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB "
                           "/BitsPerComponent 8 /Filter /FlateDecode /SMask 3 0 R"),
            pdfmini.stream(alpha, "/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 8"),
            pdfmini.stream(bytes([0b01000000, 0b10000000]), "/Type /XObject /Subtype /Image /Width 2 /Height 2 /ImageMask true "
                           "/BitsPerComponent 1")]
    c = "q 100 0 0 100 100 600 cm /Im1 Do Q q 0 0 1 rg 100 0 0 100 300 600 cm /Im2 Do Q"
    r = _render(pdfmini.build(c, "<< /XObject << /Im1 2 0 R /Im2 4 0 R >> >>", objs))
    assert px(r, 125, 675) == (255, 0, 0)  # 左上
    assert px(r, 175, 675) == WHITE  # 右上は透ける（SMask 0）
    assert px(r, 125, 625) == (0, 0, 255) and px(r, 175, 625) == (255, 255, 0)
    # 型抜き：0 の所を塗りの色（青）で
    assert px(r, 325, 675) == (0, 0, 255) and px(r, 375, 675) == WHITE
    assert px(r, 325, 625) == WHITE and px(r, 375, 625) == (0, 0, 255)


def test_inline_image_and_indexed_colors():
    c = ("q 100 0 0 50 100 600 cm BI /W 2 /H 1 /CS [/I /RGB 1 <FF000000FF00>] /BPC 8 ID \x00\x01 EI Q")
    r = _render(pdfmini.build(c, "<< >>", []))
    assert px(r, 120, 625) == (255, 0, 0) and px(r, 180, 625) == (0, 255, 0)


# ====================================================================== 書体を手で組み立てる
def _index(items: list[bytes]) -> bytes:
    if not items:
        return b"\0\0"
    offs = [1]
    for it in items:
        offs.append(offs[-1] + len(it))
    return struct.pack(">HB", len(items), 4) + b"".join(struct.pack(">I", o) for o in offs) + b"".join(items)


def _num2(v: int) -> bytes:
    """Type 1・Type 2 の字の命令の数。"""
    if -107 <= v <= 107:
        return bytes([v + 139])
    if 108 <= v <= 1131:
        v -= 108
        return bytes([247 + v // 256, v % 256])
    if -1131 <= v <= -108:
        v = -v - 108
        return bytes([251 + v // 256, v % 256])
    raise ValueError(v)


def _square_ops(x0=100, w=400, h=700) -> bytes:
    """x0 から幅 w・高さ h の四角（rmoveto・rlineto の数と命令）。"""
    return (_num2(x0) + _num2(0) + bytes([21]) + _num2(w) + _num2(0) + bytes([5]) + _num2(0) + _num2(h) + bytes([5])
            + _num2(-w) + _num2(0) + bytes([5]))


def tiny_truetype() -> bytes:
    """字形 1：四角（100〜500 × 0〜700）、字形 2：2 次の曲線を含む三角。単位 1000/em。"""
    def simple(pts):
        n = len(pts)
        out = struct.pack(">hhhhh", 1, min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts),
                          max(p[1] for p in pts))
        out += struct.pack(">HH", n - 1, 0) + bytes([1 if p[2] else 0 for p in pts])
        px_, py_ = 0, 0
        xs = ys = b""
        for x, y, _on in pts:
            xs += struct.pack(">h", x - px_)
            ys += struct.pack(">h", y - py_)
            px_, py_ = x, y
        out += xs + ys
        return out + b"\0" * (len(out) % 2)

    glyphs = [b"", simple([(100, 0, 1), (500, 0, 1), (500, 700, 1), (100, 700, 1)]),
              simple([(100, 0, 1), (500, 0, 1), (300, 900, 0)])]
    loca, glyf = [], b""
    for g in glyphs:
        loca.append(len(glyf))
        glyf += g
    loca.append(len(glyf))
    head = struct.pack(">IIIIHHqqhhhhHHhhh", 0x10000, 0x10000, 0, 0x5F0F3CF5, 0, 1000, 0, 0, 0, 0, 1000, 1000, 0, 8, 2, 1, 0)
    tabs = {"head": head, "maxp": struct.pack(">IH", 0x5000, len(glyphs)), "loca": struct.pack(f">{len(loca)}I", *loca),
            "glyf": glyf, "hhea": struct.pack(">I" + "h" * 15 + "H", 0x10000, 800, -200, 0, *([0] * 12), len(glyphs)),
            "hmtx": b"".join(struct.pack(">Hh", 600, 0) for _ in glyphs)}
    n = len(tabs)
    out = struct.pack(">IHHHH", 0x10000, n, 0, 0, 0)
    off = 12 + 16 * n
    body = b""
    for tag, data in tabs.items():
        out += struct.pack(">4sIII", tag.encode(), 0, off + len(body), len(data))
        body += data + b"\0" * (-len(data) % 4)
    return out + body


def tiny_cff() -> bytes:
    """名前で引く CFF：字形 1 の名前は A（SID 34）。標準の符号で A（65）。"""
    hdr = bytes([1, 0, 4, 4])
    name = _index([b"Test"])
    top_len = 4 * 5 + 3  # 固定の長さ（数は 29 の 5 バイト）
    strings = _index([])
    gsubrs = _index([])
    head_len = len(hdr) + len(name) + len(_index([b"\0" * top_len])) + len(strings) + len(gsubrs)
    charset = bytes([0]) + struct.pack(">H", 34)
    cs = _index([bytes([14]), _num2(600) + _square_ops() + bytes([14])])
    off_charset = head_len
    off_cs = off_charset + len(charset)
    off_priv = off_cs + len(cs)

    def i5(v):
        return bytes([29]) + struct.pack(">i", v)

    top = i5(off_charset) + bytes([15]) + i5(off_cs) + bytes([17]) + i5(0) + i5(off_priv) + bytes([18])
    assert len(top) == top_len
    return hdr + name + _index([top]) + strings + gsubrs + charset + cs


def _encrypt(data: bytes, r: int) -> bytes:
    out = bytearray()
    for p in data:
        c = p ^ (r >> 8)
        r = ((c + r) * 52845 + 22719) & 0xFFFF
        out.append(c)
    return bytes(out)


def tiny_type1() -> bytes:
    """Type 1：/A は四角（hsbw の後、rmoveto・rlineto・closepath）。Subrs 0 は return だけ。"""
    a = _encrypt(b"\0\0\0\0" + _num2(0) + _num2(600) + bytes([13]) + _square_ops() + bytes([9, 14]), 4330)
    nd = _encrypt(b"\0\0\0\0" + _num2(0) + _num2(500) + bytes([13, 14]), 4330)
    sub = _encrypt(b"\0\0\0\0" + bytes([11]), 4330)
    priv = (b"dup /Private 8 dict dup begin /RD{string currentfile exch readstring pop}executeonly def "
            b"/ND{noaccess def}executeonly def /NP{noaccess put}executeonly def /lenIV 4 def\n"
            b"/Subrs 1 array\ndup 0 %d RD " % len(sub) + sub + b" NP\nND\n2 index /CharStrings 2 dict dup begin\n"
            b"/.notdef %d RD " % len(nd) + nd + b" ND\n/A %d RD " % len(a) + a + b" ND\nend\nend\nmark currentfile closefile\n")
    clear = (b"%!PS-AdobeFont-1.0: Tiny 001\n12 dict begin\n/FontName /Tiny def\n/FontMatrix [0.001 0 0 0.001 0 0] readonly def\n"
             b"/Encoding StandardEncoding def\ncurrentfile eexec\n")
    return clear + _encrypt(b"\0\0\0\0" + priv, 55665) + b"\n" + b"0" * 64 + b"\ncleartomark\n"


def _font_pdf(font_obj: str, extra: list, text: str = "BT /F1 100 Tf 100 600 Td (AA) Tj ET") -> bytes:
    return pdfmini.build(text, "<< /Font << /F1 2 0 R >> >>", [font_obj] + extra)


def _square_glyph_at(r: dict, x: float, y: float = 600, size: float = 100) -> bool:
    """原点 (x, y)・大きさ size の四角の字（0.1〜0.5 em × 0〜0.7 em）が描かれているか。"""
    s = size / 1000
    inside = px(r, x + 300 * s, y + 350 * s)
    left, above = px(r, x + 50 * s, y + 350 * s), px(r, x + 300 * s, y + 760 * s)
    return inside != WHITE and sum(inside) <= 255 and left == WHITE and above == WHITE


def test_embedded_truetype_subset_composite_and_simple():
    tt = tiny_truetype()
    ff = pdfmini.stream(zlib.compress(tt), f"/Filter /FlateDecode /Length1 {len(tt)}")
    t0 = ("<< /Type /Font /Subtype /Type0 /BaseFont /ABCDEF+Tiny /Encoding /Identity-H /DescendantFonts [<< /Type /Font "
          "/Subtype /CIDFontType2 /BaseFont /ABCDEF+Tiny /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> "
          "/DW 600 /CIDToGIDMap /Identity /FontDescriptor 3 0 R >>] >>")
    desc = "<< /Type /FontDescriptor /FontName /ABCDEF+Tiny /Flags 4 /FontBBox [0 0 600 900] /Ascent 800 /Descent -200 /FontFile2 4 0 R >>"
    r = _render(_font_pdf(t0, [desc, ff], "BT /F1 100 Tf 100 600 Td <00010001> Tj ET BT /F1 100 Tf 100 400 Td <0002> Tj ET"))
    assert _square_glyph_at(r, 100) and _square_glyph_at(r, 160)  # 送りの幅 600
    assert r["stats"]["glyphs_embedded"] == 3 and not r["substituted"]
    # 2 次の曲線の三角：底の中ほどは塗り、頂点の近くの両側は白い
    assert dark(px(r, 130, 410)) and px(r, 112, 480) == WHITE
    # 単純な TrueType（記号の書体：符号 = 字形の番号）
    simple = ("<< /Type /Font /Subtype /TrueType /BaseFont /Tiny /FirstChar 1 /LastChar 1 /Widths [600] "
              "/FontDescriptor 3 0 R >>")
    r = _render(_font_pdf(simple, [desc, ff], "BT /F1 100 Tf 100 600 Td <0101> Tj ET"))
    assert _square_glyph_at(r, 100) and _square_glyph_at(r, 160)


def test_embedded_cff_and_type1_and_type3():
    desc = "<< /Type /FontDescriptor /FontName /Tiny /Flags 32 /FontBBox [0 0 600 700] /Ascent 700 /Descent 0 {} 4 0 R >>"
    cff = tiny_cff()
    r = _render(_font_pdf("<< /Type /Font /Subtype /Type1 /BaseFont /Tiny /FirstChar 65 /LastChar 65 /Widths [600] "
                          "/FontDescriptor 3 0 R >>",
                          [desc.format("/FontFile3"), pdfmini.stream(cff, "/Subtype /Type1C")]))
    assert _square_glyph_at(r, 100) and _square_glyph_at(r, 160), "CFF"
    t1 = tiny_type1()
    r = _render(_font_pdf("<< /Type /Font /Subtype /Type1 /BaseFont /Tiny /FirstChar 65 /LastChar 65 /Widths [600] "
                          "/FontDescriptor 3 0 R >>",
                          [desc.format("/FontFile"), pdfmini.stream(t1, f"/Length1 {t1.find(b'eexec') + 6} /Length2 0 /Length3 0")]))
    assert _square_glyph_at(r, 100) and _square_glyph_at(r, 160), "Type 1"
    assert r["stats"]["glyphs_embedded"] == 2
    # 型 3：字の手続き（d1 なので、塗りの色は字を描くときの色）
    proc = pdfmini.stream("600 0 100 0 500 700 d1 100 0 400 700 re f")
    t3 = ("<< /Type /Font /Subtype /Type3 /FontBBox [0 0 600 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /sq 3 0 R >> "
          "/Encoding << /Type /Encoding /Differences [65 /sq] >> /FirstChar 65 /LastChar 65 /Widths [600] >>")
    r = _render(_font_pdf(t3, [proc], "1 0 0 rg BT /F1 100 Tf 100 600 Td (AA) Tj ET"))
    assert _square_glyph_at(r, 100) and _square_glyph_at(r, 160) and px(r, 130, 635) == (255, 0, 0)


def test_cff_and_type1_outline_readers():
    o = fontprog.CFFOutlines(tiny_cff())
    assert o.names.get("A") == 1 and o.encoding.get(65) == 1
    ops = o.glyph(1)
    assert ops[0] == (0, 100, 0) and (1, 500, 0) in ops and ops[-1] == (4,)
    t = fontprog.Type1Outlines(tiny_type1())
    assert set(t.charstrings) == {".notdef", "A"} and len(t.subrs) == 1
    assert t.glyph("A")[0] == (0, 100, 0)
    tt = fontprog.TrueTypeOutlines(fontprog.sfnt_tables(tiny_truetype()))
    assert tt.glyph(2)[-2][0] == 3  # 2 次の曲線


def test_text_render_modes_and_clip():
    tt = tiny_truetype()
    t0 = ("<< /Type /Font /Subtype /Type0 /BaseFont /T /Encoding /Identity-H /DescendantFonts [<< /Type /Font "
          "/Subtype /CIDFontType2 /BaseFont /T /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> "
          "/DW 600 /FontDescriptor 3 0 R >>] >>")
    desc = "<< /Type /FontDescriptor /FontName /T /Flags 4 /FontBBox [0 0 600 900] /Ascent 800 /Descent -200 /FontFile2 4 0 R >>"
    c = ("BT 3 Tr /F1 100 Tf 100 700 Td <0001> Tj ET "      # 見えない字
         "BT 1 Tr 4 w /F1 100 Tf 300 700 Td <0001> Tj ET "  # 線だけ：中は白い
         "q BT 7 Tr /F1 100 Tf 100 400 Td <0001> Tj ET 0 0 1 rg 0 0 595 842 re f Q")  # 字で切り抜く
    r = _render(_font_pdf(t0, [desc, pdfmini.stream(tt)], c))
    assert px(r, 130, 735) == WHITE
    assert dark(px(r, 310, 735)) and px(r, 330, 735) == WHITE
    assert px(r, 130, 435) == (0, 0, 255) and px(r, 50, 435) == WHITE and px(r, 130, 300) == WHITE


# ====================================================================== ページ・注釈・見えない内容
def test_rotated_page_and_crop_box():
    blob = pdfmini.build("1 0 0 rg 0 0 100 50 re f", "<< >>", [])
    r = _render(blob.replace(b"/MediaBox [0 0 595 842]", b"/MediaBox [0 0 595 842] /Rotate 90"))
    assert r["w"] > r["h"]
    # 回すと、ページの左下は、見た目の左上：赤い四角は左上に縦長に
    s = r["scale"] / (72 / 25.4)
    assert tuple(r["rgb"][(int(10 * s) * r["w"] + int(10 * s)) * 3:][:3]) == (255, 0, 0)
    assert tuple(r["rgb"][(int(10 * s) * r["w"] + int(80 * s)) * 3:][:3]) == WHITE
    assert tuple(r["rgb"][(int(80 * s) * r["w"] + int(10 * s)) * 3:][:3]) == (255, 0, 0)
    # CropBox：切り取った所だけ
    r = _render(blob.replace(b"/MediaBox [0 0 595 842]", b"/MediaBox [0 0 595 842] /CropBox [0 0 200 100]"))
    assert (r["w"], r["h"]) == (200, 100)
    assert tuple(r["rgb"][(90 * 200 + 10) * 3:][:3]) == (255, 0, 0)


def test_annotation_appearance_hidden_and_optional_content():
    ap = pdfmini.stream("0 0 1 rg 0 0 100 20 re f", "/Type /XObject /Subtype /Form /BBox [0 0 100 20] /Matrix [1 0 0 1 0 0]")
    ap2 = pdfmini.stream("1 0 0 rg 0 0 10 10 re f", "/Type /XObject /Subtype /Form /BBox [0 0 10 10]")
    objs = [ap, ap2,
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (a) /Rect [100 700 300 720] /AP << /N 2 0 R >> >>",  # 幅 2 倍に合わせる
            "<< /Type /Annot /Subtype /Square /F 2 /Rect [100 600 200 700] /AP << /N 3 0 R >> >>",  # 隠した注釈
            "<< /Type /OCG /Name (off) >>"]
    c = "/OC /L1 BDC 0 1 0 rg 100 400 100 100 re f EMC 0 g 300 400 50 50 re f"
    blob = pdfmini.build(c, "<< /Properties << /L1 6 0 R >> >>", objs, annots="/Annots [4 0 R 5 0 R]",
                         acroform="/OCProperties << /OCGs [6 0 R] /D << /OFF [6 0 R] >> >>")
    r = _render(blob)
    assert px(r, 110, 710) == (0, 0, 255) and px(r, 290, 710) == (0, 0, 255)  # Rect に合わせて広げる
    assert px(r, 150, 650) == WHITE  # 隠した注釈は描かない
    assert px(r, 150, 450) == WHITE and px(r, 320, 420) == BLACK  # OFF の内容は描かない


def test_written_field_is_drawn_with_glyph_shapes(tmp_path):
    from todoku import api

    src, out = tmp_path / "f.pdf", tmp_path / "o.pdf"
    src.write_bytes(pdfmini.build("0 G 1 w 100 300 200 20 re S", "<< >>", []))
    r = api.fill(str(src), {"@1": {"page": 1, "box": [40, 40, 110, 50], "value": "ABC"}}, str(out))
    assert r["proof_ok"], r
    rr = render(PdfDocument(out.read_bytes()), 1, dpi=150)
    s, W = rr["scale"], rr["w"]
    ink = [(x, y) for y in range(int(39 * s), int(51 * s)) for x in range(int(39 * s), int(111 * s))
           if rr["rgb"][(y * W + x) * 3] < 128]
    assert ink, "書いた字が描かれていません"
    xs, ys = [p[0] / s for p in ink], [p[1] / s for p in ink]
    assert 40 - 0.5 <= min(xs) and max(xs) <= 110.5 and 40 - 0.5 <= min(ys) and max(ys) <= 50.5
    assert max(xs) - min(xs) > 2 * (max(ys) - min(ys))  # 横に並ぶ 3 字


def test_missing_font_falls_back_to_boxes_and_says_so(monkeypatch):
    from todoku.pdf.render import fontmap

    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    blob = pdfmini.build("BT /F1 40 Tf 100 600 Td (AB) Tj ET", "<< /Font << /F1 2 0 R >> >>",
                         ["<< /Type /Font /Subtype /Type1 /BaseFont /NoSuchFont-Bold >>"])
    doc = PdfDocument(blob, "t.pdf")
    v = view_page(doc, 1)
    assert "灰色の四角" in v["text"] and "NoSuchFont-Bold" in v["text"]
    r = render(PdfDocument(blob, "t.pdf"), 1, dpi=72)
    assert r["stats"]["glyphs_box"] == 2


def test_view_page_png_and_dpi():
    import base64

    blob = pdfmini.build("0 g 100 700 50 50 re f", "<< >>", [])
    v = view_page(PdfDocument(blob, "t.pdf"), 1, dpi=50)
    png = base64.b64decode(v["png_base64"])
    assert png[:8] == b"\x89PNG\r\n\x1a\n" and "50 dpi" in v["text"]
    w, h = struct.unpack(">II", png[16:24])
    assert w == 26 + 414 and h == 26 + 585  # 帯（26 画素）＋ A4 の 50 dpi


def test_scan_rows_cover_exact_area():
    rows = canvas.scan([[(0.0, 0.0), (10.0, 0.0), (10.0, 10.0), (0.0, 10.0), (5.0, 5.0)]])
    assert rows
    total = sum((x1 - x0) * a for _y, runs in rows for x0, x1, a in runs) / 255
    tri = canvas.scan([[(0.0, 0.0), (8.0, 0.0), (0.0, 8.0)]])
    area = sum((x1 - x0) * a for _y, runs in tri for x0, x1, a in runs) / 255
    assert abs(area - 32) < 1.0 and total > 0
    with pytest.raises(ValueError):
        _num2(5000)
