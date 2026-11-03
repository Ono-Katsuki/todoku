"""字の文字の種類（chars.script_of：Unicode の Scripts.txt）と、それを使う「Unicode でない書体かもしれない」の見分け。
字の名前の頭の語（MASCULINE・MICRO・MODIFIER・COMBINING など）を文字の種類と取り違えない。"""

from __future__ import annotations

import sys
import unicodedata
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from test_pdf_scripts import flags  # noqa: E402
from test_review_fixes import MAYBE, _two_font_pdf  # noqa: E402

from todoku import chars  # noqa: E402
from todoku.pdf.fonts import EAST_ASIAN, LANG_SCRIPT, letter_script  # noqa: E402


@pytest.mark.parametrize("ch,script", [
    ("A", "Latin"), ("º", "Latin"), ("ª", "Latin"), ("Ａ", "Latin"), ("é", "Latin"),
    ("µ", "Common"), ("ʼ", "Common"), ("1", "Common"), ("。", "Common"), ("ー", "Common"),
    ("́", "Inherited"), ("゙", "Inherited"),
    ("漢", "Han"), ("々", "Han"), ("あ", "Hiragana"), ("ア", "Katakana"), ("ｱ", "Katakana"), ("한", "Hangul"),
    ("न", "Devanagari"), ("ع", "Arabic"), ("ж", "Cyrillic"), ("ሀ", "Ethiopic"), ("က", "Myanmar"),
    ("͸", "Unknown"), ("\U0010FFFF", "Unknown"),
])
def test_script_of(ch, script):
    assert chars.script_of(ch) == script


def test_script_table_is_sorted_and_from_the_same_ucd():
    assert chars.UCD_VERSION == unicodedata.unidata_version
    prev = -1
    for a, b, i in chars._SCRIPTS:
        assert prev < a <= b and 0 <= i < len(chars.SCRIPT_NAMES)
        prev = b
    # 字の割り当てのある符号位置（未割り当て Cn・私用 Co・サロゲート Cs 以外）は、どれも表にある（Unknown にならない）
    for c in range(0, 0x30000, 7):
        if unicodedata.category(chr(c)) not in ("Cn", "Co", "Cs"):
            assert chars.script_of(chr(c)) != "Unknown", hex(c)


@pytest.mark.parametrize("ch", ["º", "ª"])
def test_ordinal_indicators_are_latin(ch):
    assert letter_script(ch) == "LATIN"


@pytest.mark.parametrize("ch", ["µ", "ʼ", "́", "1", "-"])
def test_neutral_letters_have_no_script(ch):
    assert letter_script(ch) == ""


def test_lang_and_east_asian_names_are_script_values():
    names = {n.upper() for n in chars.SCRIPT_NAMES}
    assert set(LANG_SCRIPT.values()) <= names
    assert set(EAST_ASIAN) <= names
    assert letter_script("漢").startswith(EAST_ASIAN) and letter_script("ｱ").startswith(EAST_ASIAN)


def test_ordinary_latin_text_with_ordinals_micro_and_accents_is_not_maybe_legacy():
    # ラテン文字の書類の「Nº」「ª」「µm」「ʼ」・結合の記号つきの e：どれもラテン文字か、どの文字の種類にも決まらない字
    assert MAYBE not in flags(_two_font_pdf("Nº 5 1ª µm donʼt é", b"Name"))


def test_devanagari_page_is_still_maybe_legacy():
    f = flags(_two_font_pdf("नाम", b"dey"))
    assert MAYBE in f and "DEVANAGARI" in f[MAYBE]["where"][0]
