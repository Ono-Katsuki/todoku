"""PDF に足す字を、同梱の書体（Noto Sans JP）から部分埋め込みする（作った書類で）。"""

from __future__ import annotations

import gzip
import hashlib
import subprocess
import sys
from pathlib import Path

import pytest

from pdfmini import build, stream
from test_fill_eval_fixes import _form, _widget
from test_pdf_write_fixes import signed
from todoku import api
from todoku.pdf import embedfont, glyphs
from todoku.pdf import fill as F
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import fill
from todoku.pdf.form import read_fields
from todoku.pdf.objects import Stream

ROOT = Path(__file__).resolve().parent.parent
HELV = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"


def _write(tmp_path, blob: bytes, values: dict, name: str = "o.pdf"):
    src, out = tmp_path / "a.pdf", tmp_path / name
    src.write_bytes(blob)
    return api.fill(str(src), values, str(out)), out


def _ap_fonts(doc, wd) -> dict:
    st = glyphs.widget_appearance(doc, wd)
    return doc.resolve(doc.resolve(st.dict["Resources"])["Font"])


def _descendant(doc, font: dict) -> dict:
    return doc.resolve(doc.resolve(font["DescendantFonts"])[0])


def _ap_text(blob: bytes) -> tuple[str, list]:
    doc = PdfDocument(blob)
    return glyphs.appearance_text(doc, glyphs.widget_appearance(doc, _widget(blob)))


# ---------------------------------------------------------------- 書体そのもの
def test_bundled_file_is_the_recorded_original():
    data = gzip.decompress(embedfont.BUNDLED_FILE.read_bytes())
    assert hashlib.sha256(data).hexdigest() == "6255ea4692e97ef27330304fb96fd650f9111f868cd585fcf2802d42823930e5"
    assert "6255ea4692e97ef2" in (ROOT / "THIRD_PARTY_NOTICES.md").read_text(encoding="utf-8")
    for p in (ROOT / "LICENSES" / "LICENSE-NotoSansJP-OFL.txt", embedfont.BUNDLED_FILE.parent / "LICENSE-NotoSansJP-OFL.txt"):
        text = p.read_text(encoding="utf-8")
        assert "SIL OPEN FONT LICENSE Version 1.1" in text and "Copyright 2014-2021 Adobe" in text
    assert embedfont.BUNDLED_FILE.stat().st_size < 4 * 1024 * 1024


def test_bundled_font_can_be_embedded():
    f = embedfont.bundled_font()
    assert f is not None and "glyf" in f.t and f.embeddable


def test_font_is_read_only_when_a_character_is_chosen():
    code = ("from todoku.pdf import embedfont, fill, glyphs;"
            "assert embedfont._read_bundled.cache_info().currsize == 0;"
            "glyphs.choose('山');"
            "assert embedfont._read_bundled.cache_info().currsize == 1")
    subprocess.run([sys.executable, "-c", code], check=True, env={"PYTHONPATH": str(ROOT / "src")})


def test_empty_setting_turns_embedding_off(monkeypatch):
    monkeypatch.setenv("TODOKU_PDF_FONT", "")
    assert glyphs.choose("山")[0] == "TdkJ" and glyphs.choose("A")[0] == "Helv"


# ---------------------------------------------------------------- 字ごとに書体を決める
@pytest.mark.parametrize("ch", ["山", "あ", "ア", "A", "é", "Ж", "Ω", "〒", "①", "ệ", "𠮟"])
def test_a_character_with_a_shape_is_embedded(ch):
    assert glyphs.choose(ch)[0] == embedfont.BUNDLED_KEY


@pytest.mark.parametrize("ch,key", [("ş", "TdkL"), ("한", "TdkK"), ("们", "TdkG")])
def test_a_character_without_a_shape_keeps_the_old_route(ch, key):
    assert embedfont.bundled_font().gid(ch) is None
    assert glyphs.choose(ch)[0] == key


def test_combining_marks_stay_refused():
    with pytest.raises(glyphs.GlyphError):
        glyphs.choose(chr(0x301))


def test_the_language_of_the_document_does_not_change_the_choice():
    with glyphs.cjk_order("zh-TW"):
        assert glyphs.choose("中")[0] == embedfont.BUNDLED_KEY


def test_characters_sharing_a_shape_have_different_codes():
    font = embedfont.bundled_font()
    assert font.cmap[0x301C] == font.cmap[0xFF5E]  # 「〜」と「～」は同じ字形
    assert glyphs.choose("〜")[1] != glyphs.choose("～")[1]


# ---------------------------------------------------------------- 書いた結果
def test_written_font_is_embedded_and_reported(tmp_path):
    r, out = _write(tmp_path, _form("100 700 500 730"), {"f": "山田 Hanako"})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert "fonts_not_embedded" not in r and "埋め込まず" not in r["say_to_user"]
    blob = out.read_bytes()
    doc = PdfDocument(blob)
    wd = _widget(blob)
    fonts = _ap_fonts(doc, wd)
    assert list(fonts) == ["TdkB"]
    font = doc.resolve(fonts["TdkB"])
    assert "NotoSansJP-Regular" in font["BaseFont"].name and font["Encoding"].name == "Identity-H"
    cid = _descendant(doc, font)
    assert cid["Subtype"].name == "CIDFontType2" and isinstance(doc.resolve(cid["CIDToGIDMap"]), Stream)
    assert "FontFile2" in doc.resolve(cid["FontDescriptor"])


def test_da_and_dr_name_the_drawn_font(tmp_path):
    r, out = _write(tmp_path, _form("100 700 500 730"), {"f": "Osaka"})
    assert r["proof_ok"]
    blob = out.read_bytes()
    doc = PdfDocument(blob)
    wd = _widget(blob)
    assert b"/TdkB 9 Tf" in bytes(wd["DA"])
    dr = doc.resolve(doc.resolve(doc.root["AcroForm"])["DR"])
    assert doc.resolve(doc.resolve(dr["Font"])["TdkB"]) is not None
    assert doc.resolve(_ap_fonts(doc, wd)["TdkB"]) == doc.resolve(doc.resolve(dr["Font"])["TdkB"])


def test_new_form_default_da_does_not_name_an_absent_font():
    blob = build("BT /F1 10 Tf 72 700 Td (Name) Tj ET", "<< /Font << /F1 2 0 R >> >>", [HELV])
    doc = PdfDocument(blob)
    res = fill(doc, {"@1": {"page": 1, "box": [40, 60, 150, 70], "value": "山田"}})
    new = PdfDocument(res.blob)
    af = new.resolve(new.root["AcroForm"])
    font = bytes(af["DA"]).split()[0].decode()[1:]
    assert font in new.resolve(new.resolve(af["DR"])["Font"])


@pytest.mark.parametrize("value", ["〜～", "一⼀", "長⻑", "-­‑", "・‧", "𠮟る", "Ｑ＆Ａ"])
def test_the_text_layer_reads_back_the_same_characters(tmp_path, value):
    r, out = _write(tmp_path, _form("100 700 500 730"), {"f": value})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    text, probs = _ap_text(out.read_bytes())
    assert text == value and not probs


def test_tounicode_is_exact_for_every_written_code():
    cids = sorted(int.from_bytes(glyphs.choose(c)[1], "big") for c in "〜～山")
    _t0, _cid, _desc, _data, tu, c2g, _set = embedfont.bundled_objects(cids)
    from todoku.pdf.fonts import CMap

    cmap = CMap.parse(tu)
    font = embedfont.bundled_font()
    for c in "〜～山":
        code = glyphs.choose(c)[1]
        assert cmap.to_unicode(code) == c
        n = int.from_bytes(code, "big")
        assert int.from_bytes(c2g[2 * n:2 * n + 2], "big") == font.cmap[ord(c)]


def test_only_used_glyphs_are_kept_and_the_output_stays_small(tmp_path):
    src = _form("100 700 500 730")
    r, out = _write(tmp_path, src, {"f": "山田花子 東京都日野市"})
    blob = out.read_bytes()
    assert r["proof_ok"] and blob.startswith(src)  # 元のバイト列はそのまま（追記）
    assert len(blob) - len(src) < 30_000
    doc = PdfDocument(blob)
    desc = doc.resolve(_descendant(doc, doc.resolve(_ap_fonts(doc, _widget(blob))["TdkB"]))["FontDescriptor"])
    sub = embedfont.TrueType("", doc.decode(doc.resolve(desc["FontFile2"])))
    full = embedfont.bundled_font()
    assert sub.glyph(full.cmap[ord("山")]) == full.glyph(full.cmap[ord("山")])
    assert not sub.glyph(full.cmap[ord("海")])


def test_missing_shapes_stay_listed(tmp_path):
    r, _out = _write(tmp_path, _form("100 700 500 730"), {"f": "山田 김 Şahin"})
    assert r["proof_ok"] and r["read_back_ok"]
    got = {x["font"]: x for x in r["fonts_not_embedded"]}
    assert got["HYGoThic-Medium"]["kind"] == "cjk_predefined" and got["Helvetica"]["kind"] == "standard14"
    assert "NotoSansJP" not in str(got)


def test_second_write_keeps_the_first_font_and_uses_a_new_name(tmp_path):
    extra = "<< /Type /Annot /Subtype /Widget /FT /Tx /T (g) /F 4 /Rect [100 650 500 680] /DA (/F5 9 Tf) /P 1 0 R >>"
    r, out = _write(tmp_path, _form("100 700 500 730", extra=extra), {"f": "山田"})
    assert r["proof_ok"]
    first = out.read_bytes()
    out2 = tmp_path / "o2.pdf"
    r2 = api.fill(str(out), {"g": "花子"}, str(out2))
    assert r2["proof_ok"] and r2["read_back_ok"], (r2["unexpected"], r2["read_back_differences"])
    blob = out2.read_bytes()
    assert blob.startswith(first)
    doc = PdfDocument(blob)
    fields = {f.name: doc.resolve(f.ref) for f in read_fields(doc)}
    assert b"/TdkB " in bytes(fields["f"]["DA"]) and b"/TdkB2 " in bytes(fields["g"]["DA"])
    fonts = doc.resolve(doc.resolve(doc.resolve(doc.root["AcroForm"])["DR"])["Font"])
    assert {"TdkB", "TdkB2"} <= set(fonts)
    assert glyphs.appearance_text(doc, glyphs.widget_appearance(doc, fields["f"]))[0] == "山田"
    assert glyphs.appearance_text(doc, glyphs.widget_appearance(doc, fields["g"]))[0] == "花子"


def test_transparent_text_of_a_circle_uses_the_font(tmp_path):
    from test_pdf_scripts import glyphs_pdf

    word = "はい"
    blob = glyphs_pdf([(ch, 100 + 12 * k, 700) for k, ch in enumerate(word)], {ch: 1000 for ch in word})
    r, out = _write(tmp_path, blob, {"@c1": {"circle": word}})
    assert r["proof_ok"], r["unexpected"]
    assert r["read_back_ok"], r["read_back_differences"]
    assert b"NotoSansJP-Regular" in out.read_bytes()


def test_drawn_value_on_a_page_without_a_form(tmp_path):
    blob = build("BT /F1 10 Tf 72 700 Td (Name) Tj ET", "<< /Font << /F1 2 0 R >> >>", [HELV])
    r, out = _write(tmp_path, blob, {"@1": {"page": 1, "box": [40, 60, 150, 70], "value": "山田 Hanako"}})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert "fonts_not_embedded" not in r
    assert b"NotoSansJP-Regular" in out.read_bytes()


def test_redraw_room_is_measured_in_helvetica_for_latin():
    """PDFKit は /DA の埋め込んだ書体を使わず、欧文を Helvetica で描く。"""
    want = sum(glyphs.helv_width(c) for c in "WWWW") * 10 / 1000 * F.REDRAW_SCALE + F.REDRAW_PAD
    assert F.pdfkit_needs("WWWW", 10) == pytest.approx(want)


# ---------------------------------------------------------------- 署名・PDF/A
@pytest.mark.parametrize("mode", ["approval", "p2"])
def test_signed_document_keeps_its_bytes(tmp_path, mode):
    src = signed(mode)
    r, out = _write(tmp_path, src, {"note": "山田"})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert out.read_bytes().startswith(src)
    assert "fonts_not_embedded" not in r


def _pdfa() -> bytes:
    meta = stream(b'<x:xmpmeta xmlns:x="adobe:ns:meta/"><pdfaid:part xmlns:pdfaid="http://www.aiim.org/pdfa/ns/id/">2</pdfaid:part></x:xmpmeta>')
    wid = "<< /Type /Annot /Subtype /Widget /FT /Tx /T (f) /F 4 /Rect [100 700 500 730] /DA (/Helv 9 Tf) /P 1 0 R >>"
    return build("", "<< >>", [wid, meta], annots="/Annots [2 0 R]", acroform="/AcroForm << /Fields [2 0 R] >> /Metadata 3 0 R")


def test_pdfa_gets_a_cidset(tmp_path):
    r, out = _write(tmp_path, _pdfa(), {"f": "山田"})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    blob = out.read_bytes()
    doc = PdfDocument(blob)
    desc = doc.resolve(_descendant(doc, doc.resolve(_ap_fonts(doc, _widget(blob))["TdkB"]))["FontDescriptor"])
    bits = doc.decode(doc.resolve(desc["CIDSet"]))
    for ch in "山田":
        n = int.from_bytes(glyphs.choose(ch)[1], "big")
        assert bits[n // 8] & (0x80 >> (n % 8))
    assert not any("PDF/A" in w for w in r["warnings"])


def test_pdfa_warns_when_a_character_cannot_be_embedded(tmp_path, monkeypatch):
    monkeypatch.setattr(embedfont, "system_font", lambda: None)
    r, _out = _write(tmp_path, _pdfa(), {"f": "한"})
    assert any("PDF/A" in w for w in r["warnings"]) and r["fonts_not_embedded"][0]["kind"] == "cjk_predefined"


def test_pdfa_prefers_the_system_font_over_a_font_that_is_not_embedded():
    font = embedfont.system_font()
    if font is None or font.gid("ş") is None:
        pytest.skip("ş の字形のある、このパソコンの書体がありません")
    with glyphs.embed_first(True):
        assert glyphs.choose("ş")[0] == "TdkE"
    assert glyphs.choose("ş")[0] == "TdkL"


def test_not_pdfa():
    assert not F.is_pdfa(PdfDocument(_form("100 700 500 730")))
    assert F.is_pdfa(PdfDocument(_pdfa()))


def test_a_document_in_another_language_is_told_about_the_kanji_shapes(tmp_path):
    blob = build("", "<< >>", ["<< /Type /Annot /Subtype /Widget /FT /Tx /T (f) /F 4 /Rect [100 700 500 730] /DA (/Helv 9 Tf) /P 1 0 R >>"],
                 annots="/Annots [2 0 R]", acroform="/AcroForm << /Fields [2 0 R] >> /Lang (zh-TW)")
    r, _out = _write(tmp_path, blob, {"f": "中文"})
    assert r["proof_ok"] and any("日本語の字形" in w for w in r["warnings"])
    r, _out = _write(tmp_path, _form("100 700 500 730"), {"f": "中文"}, "o2.pdf")
    assert not any("日本語の字形" in w for w in r["warnings"])


def test_da_keeps_the_colour_of_the_form(tmp_path):
    r, out = _write(tmp_path, _form("100 700 500 730", da="1 0 0 rg /F5 9 Tf"), {"f": "山田"})
    assert r["proof_ok"]
    assert bytes(_widget(out.read_bytes())["DA"]) == b"1 0 0 rg /TdkB 9 Tf"


def test_proof_catches_a_code_without_a_shape():
    blob = _form("100 700 500 730")
    res = fill(PdfDocument(blob), {"f": "山"})
    code = glyphs.choose("山")[1].hex().upper().encode()
    bad = res.blob.replace(b"<" + code + b">", b"<0001>")
    assert bad != res.blob
    from todoku.pdf.proof import prove

    pr = prove(blob, bad, {"f"})
    assert not pr.ok and any("字形" in u or "ToUnicode" in u for u in pr.unexpected), pr.unexpected
