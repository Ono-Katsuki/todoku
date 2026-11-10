"""箱の字（chars.py）：印刷された箱の字は、記入の箱のことも箇条・見出しの印のこともあり、字からは決まらない。
どの箱の字も規則のチェックの欄にせず「？」にする（チェックの欄と決めるのは、ファイルが部品として宣言しているときだけ）。"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from pdfmini import text_pdf  # noqa: E402

from todoku import chars  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.window import build as window  # noqa: E402


def test_box_sets():
    # 線の太さの違う白い四角（U+1F78E〜U+1F793：Wingdings・Wingdings 2 の箱の字）も、印のない箱の字
    assert set("☐□▢⬜❏❐❑❒◻◽") | {chr(c) for c in range(0x1F78E, 0x1F794)} == set(chars.EMPTY_BOXES)
    assert not chars.CERTAIN_EMPTY_BOXES  # 字だけで記入の箱と決まる字はない
    assert chars.BULLET_LIKE_BOXES == chars.EMPTY_BOXES
    # 太さ・書き方の違う、印のついた箱の字（U+2BBD・U+1F5F5・U+1F5F7・U+1F5F9）と、× の入った箱（U+2327）
    assert set("☑☒\u2bbd\U0001f5f5\U0001f5f7\U0001f5f9\u2327") == set(chars.CHECKED_BOXES)
    assert "▣" in chars.MAYBE_CHECKED and "▣" not in chars.BOXES  # 記入済みの箱のことも、箇条の印のこともある
    # 塗った四角（U+25FC・U+25FE・U+2B1B）・細い CHECK MARK（U+1F5F8）・中の四角の大きさの違う「▣」（U+1F794・U+1F795）も同じ
    assert set("■✓✔▣\u25fc\u25fe\u2b1b\U0001f5f8\U0001f794\U0001f795") == set(chars.MAYBE_CHECKED)
    assert all(chars.is_bullet_like_box(ch) for ch in chars.EMPTY_BOXES)


def test_glyph_facts_are_unicode_properties():
    f = chars.glyph_facts("□")
    assert f == {"char": "□", "code": "U+25A1", "name": "WHITE SQUARE", "category": "So",
                 "east_asian_width": f["east_asian_width"]}
    assert f["east_asian_width"] in ("A", "W", "N")


def _boxes(text: str) -> tuple[list[dict], str]:
    pg = window(PdfDocument(text_pdf([(text, 50, 700)]), "a.pdf"))["pages"][0]
    return [f for f in pg["fields"] if f["kind"] == "checkbox"], "\n".join(pg["reading"])


@pytest.mark.parametrize("ch", sorted(chars.EMPTY_BOXES))
def test_printed_box_is_an_agent_place(ch):
    boxes, text = _boxes(f"{ch}はい")
    assert boxes and boxes[0].get("agent"), boxes
    assert f"［{boxes[0]['id']}？］" in text
