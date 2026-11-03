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


@pytest.mark.parametrize("inner", ["甲・乙・丙・丁・戊・己・庚", "1・2・3", "とても長い名前の選択肢です・短い"])
def test_printed_option_lists_stay_plain_text(inner):
    # 印刷された選択肢の並びは、欄（○で囲む選択肢）にしない：字の意味（どれが選択肢か）はエージェントが窓の字を読んで決め、
    # 字を指定して circle で書く
    from test_pdf_scripts import glyphs_pdf, window

    text = "種別（" + inner + "）"
    blob = glyphs_pdf([(ch, 40 + 12 * k, 700) for k, ch in enumerate(text)])
    w = window(blob)["pages"][0]
    assert not [f for f in w["fields"] if f["kind"] == "circle"]
    assert any(inner in ln for ln in w["reading"]), w["reading"]


# ====================================================================== 黙って飲み込まない
def test_signature_read_failure_is_a_write_C(monkeypatch):
    from todoku.pdf import signatures
    from todoku.pdf.doc import PdfDocument
    from test_pdf_scripts import glyphs_pdf

    def boom(_doc):
        raise RuntimeError("broken")

    monkeypatch.setattr(signatures, "read", boom)
    r = signatures.write_reasons(PdfDocument(glyphs_pdf([("委", 40, 700)]), "t.pdf"))
    assert r and r[0].grade == "C" and r[0].title == "署名を確かめられない"


def test_unreadable_annotation_appearance_is_recorded_and_flagged(monkeypatch):
    from todoku.pdf import content
    from test_pdf_scripts import window

    ap = pdfmini.stream("BT /F1 12 Tf 2 2 Td (x) Tj ET", "/Type /XObject /Subtype /Form /BBox [0 0 60 20]")
    blob = pdfmini.build("", "<< >>", [ap, "<< /Type /Annot /Subtype /Stamp /Rect [100 700 160 720] /AP << /N 2 0 R >> >>"],
                         annots="/Annots [3 0 R]")
    orig = content.Interpreter.run_stream

    def boom(self, data, res, *a, **k):
        if b"(x)" in data:
            raise RuntimeError("bad appearance")
        return orig(self, data, res, *a, **k)

    monkeypatch.setattr(content.Interpreter, "run_stream", boom)
    pg = window(blob)["pages"][0]
    assert pg["annotations"] and "RuntimeError" in pg["annotations"][0]["unread"]
    fl = {f["title"]: f for f in pg["read_flags"]}
    assert fl["見た目を読めない注釈"]["grade"] == "B" and "左 100" in fl["見た目を読めない注釈"]["where"][0]


def test_view_page_tells_about_annotation_it_could_not_draw(monkeypatch):
    from todoku.pdf import content
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.pageimage import view_page

    ap = pdfmini.stream("q 10 0 0 10 0 0 cm /Im1 Do Q", "/Type /XObject /Subtype /Form /BBox [0 0 60 20]")
    blob = pdfmini.build("", "<< >>", [ap, "<< /Type /Annot /Subtype /Stamp /Rect [100 700 160 720] /AP << /N 2 0 R >> >>"],
                         annots="/Annots [3 0 R]")
    orig = content.Interpreter.run_stream

    def boom(self, data, res, *a, **k):
        if b"/Im1" in data:
            raise RuntimeError("bad")
        return orig(self, data, res, *a, **k)

    monkeypatch.setattr(content.Interpreter, "run_stream", boom)
    v = view_page(PdfDocument(blob, "t.pdf"), 1, max_side=200)
    assert "見た目を描けなかった注釈" in v["text"] and "Stamp" in v["text"]


def test_written_id_missing_after_re_read_is_not_read_back_ok(monkeypatch, tmp_path):
    from todoku import api

    src = tmp_path / "a.csv"
    src.write_bytes("名前,年\n,\n".encode("utf-8"))
    out = tmp_path / "b.csv"
    orig = api.fields

    def drop(*a, **k):  # 書いた後のファイルの窓から、書いた欄が消えたことにする
        r = orig(*a, **k)
        return [f for f in r if not str(a[0]).endswith(".todoku-tmp.csv") or f["id"] != "R2.C1"]

    monkeypatch.setattr(api, "fields", drop)
    monkeypatch.setattr(api._open, "read_back", lambda *a, **k: {})
    r = api.fill(str(src), {"R2.C1": "山田"}, str(out))
    assert r["read_back_by"] == "window"
    assert not r["read_back_ok"] and any("読み戻せなかった" in d.get("why", "") for d in r["read_back_differences"])


def test_proof_read_back_formats_are_marked(tmp_path):
    from todoku import api

    src = tmp_path / "a.txt"
    src.write_bytes("氏名：\r\n".encode("cp932"))
    fl = api.fields(str(src))
    r = api.fill(str(src), {fl[0]["id"]: "山田"}, str(tmp_path / "b.txt"))
    assert r["read_back_by"] == "proof" and r["read_back_ok"]


def test_ods_recalculation_that_cannot_be_checked_is_told(monkeypatch, tmp_path):
    from test_openformats import SHEET, make_ods
    from todoku import api
    from todoku.xlsx import readback

    src = tmp_path / "a.ods"
    src.write_bytes(make_ods(SHEET))

    def broken(*_a, **_k):
        raise RuntimeError("broken")

    monkeypatch.setattr(readback, "check", broken)
    r = api.fill(str(src), {"Sheet1!B1": "山田"}, str(tmp_path / "b.ods"))
    assert "確かめられませんでした" in r["recalculation_unchecked"]
    assert any("計算し直し" in w for w in r["warnings"])
