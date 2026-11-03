"""PDF の読む窓：世界の文字・字の重なり・外字・OCR の層・選択肢の名前など（2026-11-03 の監査の項目）。

PDF は tests/pdfmini.py で手で組み立てる（字の見た目は確かめず、ToUnicode と位置から何を読むかを確かめる）。
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pdfmini  # noqa: E402

from todoku.grade import grade_blob  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.window import build  # noqa: E402


def window(blob: bytes) -> dict:
    return build(PdfDocument(blob, "t.pdf"))


def reading(blob: bytes) -> list[str]:
    return [ln for pg in window(blob)["pages"] for ln in pg["reading"]]


def flags(blob: bytes) -> dict[str, dict]:
    return {f["title"]: f for pg in window(blob)["pages"] for f in pg["read_flags"]}


def glyphs_pdf(glyphs: list[tuple[str, float, float]], widths: dict[str, int] | None = None, size: float = 12,
               base: str = "TestFont", extra: str = "") -> bytes:
    """字を1つずつ、決めた位置に置く（Type0・Identity-H・ToUnicode）。glyphs: (字, x, y)。widths：字ごとの送り幅（/1000 em）。"""
    chars = sorted({t for t, _x, _y in glyphs})
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    w = ""
    if widths:
        w = "/W [" + " ".join(f"{code[ch]} [{widths[ch]}]" for ch in chars if ch in widths) + "]"
    content = "\n".join(f"BT /F1 {size} Tf {x:.3f} {y:.3f} Td <{code[t]:04X}> Tj ET" for t, x, y in glyphs) + "\n" + extra
    objs = [pdfmini.type0(base, 3, w), pdfmini.stream(pdfmini.tounicode({c: ch for ch, c in code.items()}))]
    return pdfmini.build(content, "<< /Font << /F1 2 0 R >> >>", objs)


def word_glyphs(word: str, adv: dict[str, int], x: float = 40, y: float = 800, size: float = 14):
    out = []
    for ch in word:
        out.append((ch, x, y))
        x += adv[ch] * size / 1000
    return out


# ====================================================================== 1. 幅の狭い字の並び
ADV = {"C": 722, "a": 556, "s": 500, "t": 278, "i": 222, "l": 222, "o": 556, " ": 278, "W": 944, "m": 833, "n": 556}


def test_narrow_letters_side_by_side_are_not_overprint():
    blob = glyphs_pdf(word_glyphs("Castillo Williams illinois", ADV), ADV, size=14)
    assert "Castillo Williams illinois" in reading(blob)


def test_real_overprint_is_one_glyph():
    # 太字に見せる重ね書き：同じ字を 0.3pt ずらして2回
    g = word_glyphs("Williams", ADV)
    g2 = [(t, x + 0.3, y) for t, x, y in g]
    blob = glyphs_pdf(g + g2, ADV, size=14)
    assert reading(blob) == ["Williams"]


# ====================================================================== 2. アラビア文字の表示形・ハングルの空白
def test_arabic_presentation_forms_become_base_letters():
    blob = glyphs_pdf([("ﻣ", 100, 800), ("ﺤ", 92, 800), ("ﻤ", 84, 800), ("ﺪ", 76, 800)], {"ﻣ": 667, "ﺤ": 667, "ﻤ": 667, "ﺪ": 667})
    text = "".join(reading(blob))
    assert not any(0xFB50 <= ord(c) <= 0xFEFF for c in text)
    assert "　" not in text
    assert set(text) <= set("محمد")


def test_lam_alef_ligature_is_two_letters():
    from todoku.pdf.fonts import visible_text

    assert visible_text("ﻻ") == "لا"


def test_hangul_word_space_is_ascii():
    blob = glyphs_pdf([("홍", 40, 800), ("길", 52, 800), ("동", 64, 800), ("씨", 88, 800)])
    lines = reading(blob)
    assert lines == ["홍길동 씨"]


# ====================================================================== 3. 単純なフォントの符号・Unicode でない書体
def _simple_font_pdf(font: str, text: bytes) -> bytes:
    content = b"BT /F1 12 Tf 40 800 Td (" + text + b") Tj ET"
    return pdfmini.build(content.decode("latin-1"), "<< /Font << /F1 2 0 R >> >>", [font])


def test_simple_font_differences_without_names_are_unmapped():
    font = ("<< /Type /Font /Subtype /TrueType /BaseFont /ABCDEF+Custom /FirstChar 65 /LastChar 67 /Widths [500 500 500] "
            "/Encoding << /Type /Encoding /Differences [65 /g12 /g13 /g14] >> >>")
    w = window(_simple_font_pdf(font, b"ABC"))
    assert w["pages"][0]["unmapped"] == 3
    assert grade_blob(_simple_font_pdf(font, b"ABC"), "t.pdf").grade in "EF"  # 字に直せない本文の字


def test_simple_font_without_encoding_reads_ascii():
    font = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"
    assert reading(_simple_font_pdf(font, b"Name")) == ["Name"]


def test_legacy_indic_font_is_flagged_E():
    font = "<< /Type /Font /Subtype /TrueType /BaseFont /KrutiDev010 /Encoding /WinAnsiEncoding >>"
    blob = _simple_font_pdf(font, b"dey")
    f = flags(blob)
    assert "字の符号が Unicode でない" in f
    assert f["字の符号が Unicode でない"]["where"][0].startswith("1ページ（左 40")
    assert grade_blob(blob, "t.pdf").grade == "E"


def test_zawgyi_is_flagged():
    from todoku.pdf.fonts import legacy_encoding, legacy_text

    assert legacy_encoding("ABCDEF+Zawgyi-One") == "zawgyi"
    assert legacy_text("zawgyi", "မ")
    assert legacy_encoding("Bamini") == "ascii"
    assert legacy_encoding("FM-Abhaya") == "ascii"
    assert legacy_encoding("ArialMT") == ""
    assert legacy_encoding("Mangal") == ""


# ====================================================================== 7. 外字
def test_private_use_char_is_flagged_B():
    blob = glyphs_pdf([("山", 40, 800), ("", 52, 800), ("太", 64, 800)])
    f = flags(blob)
    assert "外字（私用の符号の字）" in f
    assert "view_page" in f["外字（私用の符号の字）"]["detail"]
    assert grade_blob(blob, "t.pdf").grade in "BC"


# ====================================================================== 9. OCR の層・重ねて描いた行・斜線のマス・選択肢の名前
def test_ocr_layer_over_image_is_B():
    img = pdfmini.stream("\xff" * 3, "/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8")
    chars = {"氏": 1, "名": 2}
    content = ("q 595 0 0 842 0 0 cm /Im1 Do Q\nBT 3 Tr /F1 12 Tf 40 800 Td <00010002> Tj ET")
    objs = [pdfmini.type0("TestFont", 3), pdfmini.stream(pdfmini.tounicode({c: ch for ch, c in chars.items()})), img]
    blob = pdfmini.build(content, "<< /Font << /F1 2 0 R >> /XObject << /Im1 4 0 R >> >>", objs)
    f = flags(blob)
    assert "OCR の字（画像の上の見えない字）" in f
    assert grade_blob(blob, "t.pdf").grade != "A"


def test_two_overprinted_lines_are_separated():
    # 「制限該当」と「支給可否」を、少しずらして重ねて描く（ずれは行の帯の内側）
    a = [(ch, 40 + 12 * i, 800) for i, ch in enumerate("制限該当")]
    b = [(ch, 41 + 12 * i, 797) for i, ch in enumerate("支給可否")]
    blob = glyphs_pdf(a + b)
    text = " ".join(reading(blob))
    assert "制限該当" in text and "支給可否" in text
    assert "重なって描かれた行" in flags(blob)


def test_struck_cell_is_not_a_field():
    # 見出しのマスと、空のマス2つ（右のマスは斜線で消してある）
    lines = ["40 700 m 340 700 l S", "40 680 m 340 680 l S", "40 680 m 40 700 l S", "140 680 m 140 700 l S",
             "240 680 m 240 700 l S", "340 680 m 340 700 l S", "240 680 m 340 700 l S"]
    blob = glyphs_pdf([("氏", 50, 685), ("名", 62, 685)], extra="\n".join(lines))
    fields = window(blob)["pages"][0]["fields"]
    rects = [f["rect"] for f in fields]
    assert any(abs(r[0] - 140) < 2 for r in rects)
    assert not any(abs(r[0] - 240) < 2 for r in rects)


def test_spaced_checkbox_label_is_whole():
    blob = glyphs_pdf([("□", 40, 800)] + [(ch, 60 + 36 * i, 800) for i, ch in enumerate("重度訪問")])
    fields = window(blob)["pages"][0]["fields"]
    assert [f["label"] for f in fields if f["kind"] == "checkbox"] == ["重度訪問"]


def test_field_box_with_printed_text_is_shrunk():
    from todoku.pdf.layout import PageLayout
    from todoku.pdf.content import page_content

    # 大きなマスの上の方に見出し、下の方に注意書き：書く所は、そのあいだ
    lines = ["40 500 m 400 500 l S", "40 700 m 400 700 l S", "40 500 m 40 700 l S", "400 500 m 400 700 l S"]
    g = [(ch, 50 + 12 * i, 685) for i, ch in enumerate("具体的内容")] + [(ch, 50 + 12 * i, 510) for i, ch in enumerate("※注意を書く")]
    blob = glyphs_pdf(g, extra="\n".join(lines))
    doc = PdfDocument(blob, "t.pdf")
    p = doc.pages()[0]
    lay = PageLayout(doc, p, page_content(doc, p), 1)
    for sp in lay.spots:
        if sp.kind in ("text", "inline", "space"):
            inside = [c for c in lay.chars if not c.text.isspace() and sp.rect[0] < c.cx < sp.rect[2] and sp.rect[1] < c.cy < sp.rect[3]]
            assert not inside, (sp.reason, sp.rect)


def test_radio_names_not_utf8_are_shown_raw_and_flagged():
    from todoku.pdf.form import name_text, name_unreadable

    raw = b"\x82\xcd\x82\xa210".decode("latin-1")
    assert name_unreadable(raw)
    assert name_text(raw) == "#82#CD#82#A210"
    assert name_text("Yes") == "Yes"
    assert name_text("はい".encode().decode("latin-1")) == "はい"


def test_option_text_with_undefined_pdfdoc_bytes_is_unreadable():
    from todoku.pdf.form import text_unreadable

    assert text_unreadable(b"\x9f\x40")  # 0x9F は PDFDocEncoding で決まっていない
    assert not text_unreadable(b"Yes")
    assert not text_unreadable("\ufeffはい".encode("utf-16-be"))
    assert text_unreadable(b"\xfe\xff\xd8\x00")  # 対のない代用符号
