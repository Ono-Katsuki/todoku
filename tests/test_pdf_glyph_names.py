"""PDF の字形の名前と、記号の書体の符号を、公開の一覧の全体で読む。

前は、①字形の名前（ToUnicode のない書体の Differences・埋め込みの書体の組み込みの符号）を、手で作った 168 個の表と、名前から
作る規則で読んでいた：Adobe Glyph List の 4,281 個のうち、読めたのは 505 個で、ギリシャ文字の大文字の名前（Alpha など 22 個）は
小文字として読んでいた。②埋め込みのない ZapfDingbats の符号は 17 個、Symbol は 165 個だけを読んでいた。③ZapfDingbats の名前
（a19・a20・a71・a73）を、どの書体でも記号として読んでいた（ほかの書体の a20 は、ただの番号の名前のことがある）。④埋め込みのない
Symbol に Differences があると、Differences にない符号をラテン文字として読んでいた。
いまは、Adobe Glyph List と ITC Zapf Dingbats Glyph List（Adobe の配るファイルのまま）で読み、符号は symfont.py の対応表の全体で読む。
"""

from __future__ import annotations

import gzip
import hashlib
from pathlib import Path

import pytest

import pdfmini
from todoku import symfont
from todoku.pdf import fonts
from todoku.pdf.content import page_content
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fonts import glyph_to_unicode

DATA = Path(fonts.__file__).parent / "data"
ZAPF = "<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats %s >>"
SYMBOL = "<< /Type /Font /Subtype /Type1 /BaseFont /Symbol %s >>"
HELV = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica %s >>"
UNREAD = "�"


def _text(font: str, s: str) -> str:
    blob = pdfmini.build(f"BT /F1 12 Tf 100 700 Td ({s}) Tj ET", "<< /Font << /F1 2 0 R >> >>", [font])
    doc = PdfDocument(blob, "t.pdf")
    return "".join(g.text for g in page_content(doc, doc.pages()[0]).glyphs)


def _diffs(names: str) -> str:
    return f"/Encoding << /Type /Encoding /Differences [65 {names}] >>"


# ====================================================================== 一覧のファイル
@pytest.mark.parametrize("name, sha, count", [
    ("glyphlist", "a3b2f61ced9f3644cc0d4ecde5c59df34ca286c689d9484a43a710a81c466789", 4281),
    ("zapfdingbats", "f6394e3cb8a447e84a1dad75d4baaf2aa7f45dc104faf369f4720e1a774ef2dc", 201),
])
def test_the_lists_are_the_files_adobe_distributes(name, sha, count):
    raw = gzip.decompress((DATA / f"{name}.txt.gz").read_bytes())
    assert hashlib.sha256(raw).hexdigest() == sha  # THIRD_PARTY_NOTICES.md の値
    assert sum(1 for ln in raw.decode("ascii").split("\n") if ln.strip() and not ln.startswith("#")) == count
    assert (DATA / "LICENSE-adobe-agl.txt").read_text(encoding="utf-8").startswith("Copyright 2002-2019 Adobe")


def test_names_of_the_list_are_read():
    agl, zapf = fonts._glyph_lists()
    assert len(zapf) == 201 and all(glyph_to_unicode(n, True) == u for n, u in zapf.items())
    kept = {"fi", "fl", "ff", "ffi", "ffl", "Omega"}  # 前からの読み方のまま（合字は字の並び、Omega はギリシャ文字）
    assert all(glyph_to_unicode(n) == u for n, u in agl.items() if n not in kept)
    assert len(agl) > 3600


@pytest.mark.parametrize("name, want", [
    ("Alpha", "Α"), ("alpha", "α"), ("Sigma", "Σ"), ("universal", "∀"), ("afii10017", "А"),
    ("Acyrillic", "А"), ("perpendicular", "⊥"), ("Scommaaccent", "Ș"), ("Abrevehookabove", "Ẳ"),
    ("dalethatafpatah", "דֲ"), ("uni3042", "あ"), ("u1F600", "\U0001f600"), ("A.sc", "A"), ("f_i", "fi"),
    ("fi", "fi"), ("Omega", "Ω"), ("Delta", "∆"),
])
def test_glyph_names(name, want):
    assert glyph_to_unicode(name) == want


@pytest.mark.parametrize("name", ["Asmall", "Acutesmall", "parenlefttp", "registersans", "a20", "a1", "g123", ""])
def test_names_without_a_character_stay_unread(name):
    """一覧が私用の領域の字にしている名前（小さな大文字・字の部品）と、ZapfDingbats でない書体の a20 は、決めない。"""
    assert glyph_to_unicode(name) == ""


# ====================================================================== ZapfDingbats
def test_codes_of_zapf_dingbats_that_were_not_in_the_short_table():
    assert _text(ZAPF % "", r"4!\050\200\215\254\376") == "✔✁✈❨❵①➾"


def test_every_code_of_the_two_fonts_is_the_code_of_the_public_table():
    for name, table in (("symbol", fonts.SYMBOL), ("zapfdingbats", fonts.ZAPF)):
        assert table == symfont.codes(name)
    assert len(fonts.SYMBOL) == 189 and len(fonts.ZAPF) == 202
    assert symfont.codes("Zapf Dingbats")[0x80] == "❨" and symfont.codes("NoSuchFont") == {}


def test_zapf_dingbats_names_in_differences():
    assert _text(ZAPF % _diffs("/a1 /a20 /a89 /a96"), "ABCD4") == "✁✔❨❵✔"
    assert _text(ZAPF.replace("/ZapfDingbats", "/ABCDEF+ZapfDingbats") % _diffs("/a12"), "A") == "☞"


def test_the_same_names_in_another_font_are_not_read_as_dingbats():
    """ZapfDingbats でない書体の a20・a1 は、ただの番号の名前のことがある：記号として読まない（読めない字として伝える）。"""
    assert _text(HELV % _diffs("/a20 /a1"), "ABC") == UNREAD + UNREAD + "C"


# ====================================================================== 名前で決まる字
def test_names_in_differences_are_read_by_the_list():
    got = _text(HELV % _diffs("/Alpha /alpha /universal /afii10017 /perpendicular /Scommaaccent /Omega /Asmall"), "ABCDEFGH")
    assert got == "Αα∀А⊥ȘΩ" + UNREAD


# ====================================================================== Symbol
def test_codes_of_symbol_that_were_not_in_the_table():
    assert _text(SYMBOL % "", r"\042a\346\375") == "∀α⎛⎬"


def test_symbol_with_differences_keeps_its_own_codes():
    """Differences にない符号は、Symbol の符号のまま（前は、ラテン文字の a として読んだ）。"""
    assert _text(SYMBOL % _diffs("/universal"), r"Aa\042") == "∀α∀"
