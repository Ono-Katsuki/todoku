"""ページの画像が、書類の幅・マスク・線のとおりに描くか。描けないものを黙って捨てないか（view_page の忠実さ）。

画像は dpi=72 で描く：1 画素＝1 pt。ページの座標 (x, y) は、画像の (x, 842 − y)。書体は試験の中で手で組み立てるか、代わりの
書体が無いときの字の四角（送りの幅で描く）で測る（このパソコンの書体に頼らない）。
"""

from __future__ import annotations

import zlib

import pytest

import pdfmini
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fonts import load_font
from todoku.pdf.pageimage import render, view_page
from todoku.pdf.render import fontmap

WHITE = (255, 255, 255)


def _render(blob: bytes, **kw) -> dict:
    return render(PdfDocument(blob, "t.pdf"), 1, dpi=72, **kw)


def px(r: dict, x: float, y: float) -> tuple[int, int, int]:
    X, Y = int(x + 1e-6), int(842 - y + 1e-6)
    i = (Y * r["w"] + X) * 3
    return tuple(r["rgb"][i:i + 3])


def _cjk_font(enc: str, tounicode: str = "") -> str:
    """とどくが書く形の、埋め込まない CID フォント（半角 231〜632 は 500、ほかは 1000）。"""
    return (f"<< /Type /Font /Subtype /Type0 /BaseFont /HeiseiKakuGo-W5-{enc} /Encoding /{enc} {tounicode} /DescendantFonts [<< "
            "/Type /Font /Subtype /CIDFontType0 /BaseFont /HeiseiKakuGo-W5 "
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 2 >> "
            "/DW 1000 /W [231 632 500] >>] >>")


def _pf(font: str, extra: list | None = None) -> object:
    doc = PdfDocument(pdfmini.build("", "<< /Font << /F1 2 0 R >> >>", [font, *(extra or [])]), "t.pdf")
    res = doc.resolve(doc.pages()[0]["Resources"])
    return load_font(doc, doc.resolve(res["Font"])["F1"], "F1")


# ====================================================================== 幅：CID は CMap で決まる（9.7.4）
def test_predefined_cmap_kana_is_full_width():
    # 半角のひらがな（Adobe-Japan1 の 516〜）に取り違えると、幅が半分になって字が重なる
    pf = _pf(_cjk_font("UniJIS-UCS2-HW-H"))
    adv = [d[2] for d in pf.decode(bytes.fromhex("307E" "3072" "30CF" "8133" "0041" "0020"))]
    assert adv == [1.0, 1.0, 1.0, 1.0, 0.5, 0.5]  # ま・ひ・ハ・脳 は全角、A と空白は半角


def test_cmap_without_official_table_kana_is_full_width():
    pf = _pf(_cjk_font("90ms-RKSJ-H"))
    adv = [d[2] for d in pf.decode(bytes.fromhex("82DC" "83A2" "41" "B1"))]  # ま・ア（全角）、A、ｱ（半角のカタカナ）
    assert adv == [1.0, 1.0, 0.5, 0.5]


def test_width_comes_from_the_code_not_from_tounicode():
    # ToUnicode があっても CID は CMap の符号で決まる（Identity に落ちて DW にならない）
    tu = pdfmini.stream(pdfmini.tounicode({0x0041: "A"}))
    pf = _pf(_cjk_font("UniJIS-UCS2-HW-H", "/ToUnicode 3 0 R"), [tu])
    assert [d[2] for d in pf.decode(bytes.fromhex("0041"))] == [0.5]


def test_substitute_font_advances_by_the_pdf_widths_with_a_half_width_space(monkeypatch):
    # 代わりの書体の字の幅ではなく、PDF の幅で進む。字の四角（幅どおりの枠）の位置で確かめる
    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    c = ("BT /F1 40 Tf 100 600 Td <30DE3072> Tj /F2 40 Tf <20> Tj /F1 40 Tf <30C0> Tj ET")
    r = _render(pdfmini.build(c, "<< /Font << /F1 2 0 R /F2 3 0 R >> >>",
                              [_cjk_font("UniJIS-UCS2-HW-H"), "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica "
                               "/Encoding /WinAnsiEncoding >>"]))
    assert r["stats"]["glyphs_box"] >= 3
    # マ・ひ の四角は 100〜140・140〜180。半角の幅に取り違えると 120 に縁が来る
    assert px(r, 140, 620) != WHITE and px(r, 120, 620) == WHITE
    # 空白（Helvetica 278/1000）の分あいて、次の字は 180 + 11.12 から
    assert px(r, 185, 620) == WHITE and px(r, 191, 620) != WHITE


def test_word_spacing_and_tj_offsets_move_the_next_glyph():
    proc = pdfmini.stream("500 0 d0 0 0 500 700 re f")
    t3 = ("<< /Type /Font /Subtype /Type3 /FontBBox [0 0 500 700] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /sq 3 0 R >> "
          "/Encoding << /Type /Encoding /Differences [65 /sq] >> /FirstChar 32 /LastChar 65 "
          "/Widths [250 " + "0 " * 32 + "500] >>")
    c = ("BT /F1 100 Tf 40 Tw 100 600 Td (A A) Tj ET "  # 50＋25＋40 ＝ 115 だけあく
         "BT /F1 100 Tf 0 Tw 100 400 Td [(A) -200 (A)] TJ ET")  # 50＋20
    r = _render(pdfmini.build(c, "<< /Font << /F1 2 0 R >> >>", [t3, proc]))
    assert sum(px(r, 120, 630)) < 100 and px(r, 200, 630) == WHITE and sum(px(r, 220, 630)) < 100
    assert px(r, 160, 430) == WHITE and sum(px(r, 175, 430)) < 100


def test_gothic_name_wins_over_serif_flag(monkeypatch):
    seen = []
    real = fontmap.system_face

    def spy(kind):
        seen.append(kind)
        return real(kind)

    monkeypatch.setattr(fontmap, "system_face", spy)
    desc = "<< /Type /FontDescriptor /FontName /MS-Gothic /Flags 6 /FontBBox [0 0 1000 1000] /ItalicAngle 0 /Ascent 880 /Descent -120 /CapHeight 700 /StemV 80 >>"
    font = (f"<< /Type /Font /Subtype /Type0 /BaseFont /MS-Gothic /Encoding /90ms-RKSJ-H /DescendantFonts [<< /Type /Font "
            f"/Subtype /CIDFontType2 /BaseFont /MS-Gothic /CIDSystemInfo << /Registry (Adobe) /Ordering (Japan1) /Supplement 2 >> "
            f"/FontDescriptor {desc} >>] >>")
    _render(pdfmini.build("BT /F1 12 Tf 100 600 Td <82DC> Tj ET", "<< /Font << /F1 2 0 R >> >>", [font]))
    assert seen and seen[0] == "cjk-sans"


# ====================================================================== 画像：マスクの解像度は画像と別（8.9.6.3・11.6.5.3）
def _bits(rows: list[str]) -> bytes:
    out = bytearray()
    for r in rows:
        v = r.ljust((len(r) + 7) // 8 * 8, "0")
        out += bytes(int(v[i:i + 8], 2) for i in range(0, len(v), 8))
    return bytes(out)


TWO_BY_TWO_BLACK = "/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceGray /BitsPerComponent 1 "


def test_tiny_image_with_fine_soft_mask_keeps_the_shape():
    # 実際の書類の形：2×2 の黒い画像に、字の形を持つ細かな 1 ビットの SMask（解像度を画像に合わせると、字が消える）
    rows = ["0" * 80] * 40 + ["0" * 40 + "1" * 40] * 40
    mask = pdfmini.stream(zlib.compress(_bits(rows)), "/Type /XObject /Subtype /Image /Width 80 /Height 80 /ColorSpace /DeviceGray "
                          "/BitsPerComponent 1 /Filter /FlateDecode")
    img = pdfmini.stream(bytes(2), TWO_BY_TWO_BLACK + "/SMask 3 0 R")
    r = _render(pdfmini.build("q 80 0 0 80 100 600 cm /Im Do Q", "<< /XObject << /Im 2 0 R >> >>", [img, mask]))
    assert px(r, 160, 620) == (0, 0, 0)  # 右下：マスク 1
    assert px(r, 120, 620) == WHITE and px(r, 160, 660) == WHITE and px(r, 120, 660) == WHITE
    assert not r["skipped"]


def test_tiny_image_with_fine_explicit_mask_keeps_the_shape():
    rows = ["1" * 80] * 40 + ["1" * 40 + "0" * 40] * 40  # 型抜き：0 の所を描く
    mask = pdfmini.stream(zlib.compress(_bits(rows)), "/Type /XObject /Subtype /Image /Width 80 /Height 80 /ImageMask true "
                          "/BitsPerComponent 1 /Filter /FlateDecode")
    img = pdfmini.stream(bytes(2), TWO_BY_TWO_BLACK + "/Mask 3 0 R")
    r = _render(pdfmini.build("q 80 0 0 80 100 600 cm /Im Do Q", "<< /XObject << /Im 2 0 R >> >>", [img, mask]))
    assert px(r, 160, 620) == (0, 0, 0) and px(r, 120, 620) == WHITE and px(r, 160, 660) == WHITE


def test_soft_mask_matte_is_undone():
    # 色を黒（Matte）と半分混ぜた画素：元の色（赤 200）に戻す
    rgb = bytes([100, 0, 0])
    mask = pdfmini.stream(bytes([128]), "/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 "
                          "/Matte [0 0 0]")
    img = pdfmini.stream(rgb, "/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /SMask 3 0 R")
    r = _render(pdfmini.build("q 40 0 0 40 100 600 cm /Im Do Q", "<< /XObject << /Im 2 0 R >> >>", [img, mask]))
    # 白の上に、元の色 199 を半分の濃さで：255 − (255 − 199) / 2
    assert abs(px(r, 120, 620)[0] - 227) <= 2 and abs(px(r, 120, 620)[1] - 127) <= 2


def test_downscale_keeps_a_thin_line_at_non_integer_ratio():
    # 1.5 倍の縮小：画素の間引きでは 1 画素の線が消える行が出る。面積の平均なら、どの行にも残る
    w, h = 30, 30
    rows = [("0" * 7 + "1" + "0" * 22) for _ in range(h)]  # 0 が黒の 1 ビット画像の反転：1 の列だけ白
    inv = [r.replace("0", "x").replace("1", "0").replace("x", "1") for r in rows]  # 列 7 だけ黒（0）
    img = pdfmini.stream(_bits(inv), f"/Type /XObject /Subtype /Image /Width {w} /Height {h} /ColorSpace /DeviceGray /BitsPerComponent 1")
    r = _render(pdfmini.build("q 20 0 0 20 100 600 cm /Im Do Q", "<< /XObject << /Im 2 0 R >> >>", [img]))
    lowest = [min(px(r, x, y)[0] for x in range(100, 120)) for y in range(601, 620)]
    assert max(lowest) < 200, lowest  # 線はどの行でも見える


# ====================================================================== 線の調整（10.6.5）
def test_thin_rules_are_snapped_to_whole_pixels():
    c = ("0 G 0.5 w 100 700.5 m 300 700.5 l S "        # 細い横線：2 行にまたがって薄くならない
         "0.96 w 200 500.5 m 200 650.5 l S "           # 細い縦線
         "0 g 100 400.5 200 0.96 re f "                # 塗りの細い四角（罫線）
         "2 J 0.5 w [3 3] 0 d 100 300.5 m 300 300.5 l S")
    r = _render(pdfmini.build(c, "<< >>", []))
    for x, ys in ((150, range(699, 703)), (250, range(399, 403))):
        col = [px(r, x, y)[0] for y in ys]
        assert min(col) == 0 and sorted(col)[1] == 255, col  # ちょうど 1 画素が真っ黒、ほかは白
    row = [px(r, x, 600)[0] for x in range(198, 203)]
    assert min(row) == 0 and sorted(row)[1] == 255, row
    dashes = [px(r, x, y)[0] for x in range(100, 140) for y in (300, 301)]
    assert 0 in dashes and 255 in dashes  # 破線は切れている


# ====================================================================== 描けないものは伝える
def test_unsupported_image_is_marked_and_reported_with_its_place():
    img = pdfmini.stream(b"\x00" * 16, "/Type /XObject /Subtype /Image /Width 8 /Height 8 /ColorSpace /DeviceGray /BitsPerComponent 1 "
                         "/Filter /JBIG2Decode")
    blob = pdfmini.build("q 100 0 0 100 100 600 cm /Im Do Q", "<< /XObject << /Im 2 0 R >> >>", [img])
    r = _render(blob)
    assert r["skipped"] and "JBIG2" in r["skipped"][0] and "Im" in r["skipped"][0] and "左 35" in r["skipped"][0]
    assert sum(px(r, 150, 650)) < 3 * 200 and px(r, 130, 650) == WHITE  # 対角線の交わる所だけ灰色
    v = view_page(PdfDocument(blob, "t.pdf"), 1)
    assert any("JBIG2" in s for s in v["not_drawn"]) and "JBIG2" in v["text"]


def test_missing_font_xobject_and_bad_operator_are_reported():
    c = ("BT /Nope 12 Tf 100 700 Td (abc) Tj ET "  # 資源にない書体
         "/Ghost Do "                              # 資源にない部品
         "(x) w "                                   # 読めない命令
         "1 0 0 rg 100 500 50 50 re f")             # その後も描き続ける
    r = _render(pdfmini.build(c, "<< >>", []))
    assert px(r, 120, 520) == (255, 0, 0)
    joined = "／".join(r["approx"])
    assert "書体が見つからず" in joined and "Ghost" in joined and "命令（w）" in joined
    v = view_page(PdfDocument(pdfmini.build(c, "<< >>", []), "t.pdf"), 1)
    assert any("Ghost" in s for s in v["not_drawn"])


def test_unknown_glyph_of_a_non_embedded_font_is_a_box_not_nothing(monkeypatch):
    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    font = ("<< /Type /Font /Subtype /Type0 /BaseFont /NoFont /Encoding /Identity-H /DescendantFonts [<< /Type /Font "
            "/Subtype /CIDFontType2 /BaseFont /NoFont /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> >>] >>")
    r = _render(pdfmini.build("BT /F1 40 Tf 100 600 Td <0123> Tj ET", "<< /Font << /F1 2 0 R >> >>", [font]))
    assert r["stats"]["glyphs_box"] == 1


def test_field_without_appearance_and_need_appearances_are_reported():
    widget = "<< /Type /Annot /Subtype /Widget /FT /Tx /T (a) /V (HELLO) /Rect [100 700 300 724] /F 4 >>"
    blob = pdfmini.build("", "<< >>", [widget], annots="/Annots [2 0 R]", acroform="/AcroForm << /Fields [2 0 R] /NeedAppearances true >>")
    v = view_page(PdfDocument(blob, "t.pdf"), 1)
    assert any("値の入った欄 1 個" in s for s in v["not_drawn"]) and any("NeedAppearances" in s for s in v["not_drawn"])
    empty = "<< /Type /Annot /Subtype /Widget /FT /Tx /T (a) /Rect [100 700 300 724] /F 4 >>"
    v = view_page(PdfDocument(pdfmini.build("", "<< >>", [empty], annots="/Annots [2 0 R]", acroform="/AcroForm << /Fields [2 0 R] >>"), "t.pdf"), 1)
    assert v["not_drawn"] == []


def test_clean_page_has_nothing_in_not_drawn():
    v = view_page(PdfDocument(pdfmini.build("0 g 100 700 50 50 re f", "<< >>", []), "t.pdf"), 1)
    assert v["not_drawn"] == []


# ====================================================================== 窓を作れなくても、ページの画像は出す
@pytest.mark.parametrize("where", ["markdown_parts", "build"])
def test_view_page_survives_a_window_failure(tmp_path, monkeypatch, where):
    # 実例：grid のマスの行が members から全部外れて window._grid_blocks の max() が空になる書類（窓の Markdown でだけ起きる）
    from todoku import api
    from todoku.pdf import window

    def boom(*a, **k):
        raise ValueError("max() arg is an empty sequence")

    p = tmp_path / "f.pdf"
    p.write_bytes(pdfmini.build("0 g 100 700 50 50 re f", "<< >>", []))
    api._WINDOWS.clear()
    monkeypatch.setattr(window, where, boom)
    v = api.view_page(str(p), 1)
    assert v["png_base64"]
    if where == "build":
        assert any("青い枠" in s for s in v["not_drawn"])
    else:
        assert v["not_drawn"] == []  # 窓の JSON は作れるので、Markdown に頼らない
