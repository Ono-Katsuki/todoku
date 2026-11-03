"""独立した見直しで見つかった所の確かめ（小さな PDF で）：
- Unicode 以前の書体の名前は、知られた家族の名前で見分ける（似た名前の別の書体を含めない）。名前を知らない書体は、書類の文字の
  種類から「かもしれない」（読む B、位置つき）で知らせる
- ActualText で置き換えた字を、Unicode でない書体の字としない
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

import pdfmini  # noqa: E402
from test_pdf_scripts import flags  # noqa: E402

from todoku.pdf.fonts import legacy_encoding  # noqa: E402

MAYBE = "字の符号が Unicode でない書体かもしれない"


@pytest.mark.parametrize("name", ["Preeti", "Kantipur", "Sagarmatha", "Himalb", "PCSNEPALI", "PCS#20NEPALI", "Shree-Dev-0714",
                                  "PowerGeez", "VisualGeez", "ABCDEF+KrutiDev010", "Kruti Dev 010,Bold", "SutonnyMJ",
                                  "FM-Abhaya", "Bamini", "Joy", "Nudi01e", "ML-TT-Karthika", "Limon S1"])
def test_known_legacy_families_are_found(name):
    assert legacy_encoding(name) == "ascii"


@pytest.mark.parametrize("name", ["FM-Gothic", "DL-Paras", "Tab-Sans", "Joy Sans", "JoySans-Bold", "AlphaMJ", "ArialMT",
                                  "Mangal", "VisualGeezUnicode", "Agrandir", "Tamil Sangam MN"])
def test_similar_names_are_not_legacy(name):
    assert legacy_encoding(name) == ""


def _two_font_pdf(uni: str, latin: bytes, catalog: str = "") -> bytes:
    """F1：Unicode の字（Type0・ToUnicode）、F2：埋め込んだ（FontFile2 のある）名前を知らない TrueType の書体の字。"""
    chars = sorted(set(uni))
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    hexs = "".join(f"{code[c]:04X}" for c in uni)
    content = (f"BT /F1 12 Tf 40 800 Td <{hexs}> Tj ET\n" if uni else "") + \
        "BT /F2 12 Tf 40 700 Td (" + latin.decode("latin-1") + ") Tj ET"
    objs = [pdfmini.type0("TestFont", 3), pdfmini.stream(pdfmini.tounicode({c: ch for ch, c in code.items()})),
            "<< /Type /Font /Subtype /TrueType /BaseFont /ABCDEF+SomeOldFont /FirstChar 32 /LastChar 126 "
            "/Widths [" + " ".join(["500"] * 95) + "] /Encoding /WinAnsiEncoding /FontDescriptor 5 0 R >>",
            "<< /Type /FontDescriptor /FontName /ABCDEF+SomeOldFont /Flags 32 /FontBBox [0 0 1000 1000] /ItalicAngle 0 "
            "/Ascent 800 /Descent -200 /CapHeight 700 /StemV 80 /FontFile2 6 0 R >>",
            pdfmini.stream("not a real font")]
    return pdfmini.build(content, "<< /Font << /F1 2 0 R /F2 4 0 R >> >>", objs, acroform=catalog)


def test_unknown_font_with_latin_text_on_a_devanagari_page_is_maybe_legacy():
    f = flags(_two_font_pdf("नाम", b"dey"))
    assert MAYBE in f and f[MAYBE]["grade"] == "B"
    assert "SomeOldFont" in f[MAYBE]["where"][0] and "上 132" in f[MAYBE]["where"][0]
    assert "view_page" in f[MAYBE]["detail"]


def test_unknown_font_with_document_language_is_maybe_legacy():
    f = flags(_two_font_pdf("", b"dey", "/Lang (ne-NP)"))
    assert MAYBE in f and "ne-NP" in f[MAYBE]["where"][0]


def test_latin_text_in_japanese_or_latin_documents_is_not_flagged():
    assert MAYBE not in flags(_two_font_pdf("氏名", b"TEL"))
    assert MAYBE not in flags(_two_font_pdf("", b"Name"))
    assert MAYBE not in flags(_two_font_pdf("", b"TEL", "/Lang (ja-JP)"))


def test_known_legacy_reason_says_unknown_fonts_may_be_missed():
    font = "<< /Type /Font /Subtype /TrueType /BaseFont /Preeti /Encoding /WinAnsiEncoding >>"
    blob = pdfmini.build("BT /F1 12 Tf 40 800 Td (dey) Tj ET", "<< /Font << /F1 2 0 R >> >>", [font])
    f = flags(blob)
    assert f["字の符号が Unicode でない"]["grade"] == "E"
    assert "見分けられないことがあります" in f["字の符号が Unicode でない"]["detail"]


def test_actual_text_span_is_not_legacy():
    tu = pdfmini.stream(pdfmini.tounicode({1: "A", 2: "B"}))
    content = "/Span << /ActualText (Name) >> BDC BT /F1 12 Tf 40 800 Td <00010002> Tj ET EMC"
    blob = pdfmini.build(content, "<< /Font << /F1 2 0 R >> >>", [pdfmini.type0("TestFont", 3), tu])
    assert "字の符号が Unicode でない" not in flags(blob)


@pytest.mark.parametrize("inner, why", [("甲・乙・丙・丁・戊・己・庚", "7 個"), ("1・2・3", "数だけの語"),
                                         ("とても長い名前の選択肢です・短い", "11 字以上")])
def test_unusual_option_lists_are_kept_as_agent_candidates_with_reason(inner, why):
    from test_pdf_scripts import glyphs_pdf, window

    text = "種別（" + inner + "）"
    blob = glyphs_pdf([(ch, 40 + 12 * k, 700) for k, ch in enumerate(text)])
    circ = [f for f in window(blob)["pages"][0]["fields"] if f["kind"] == "circle"]
    assert circ and circ[0]["agent"] and why in circ[0]["why"], circ
