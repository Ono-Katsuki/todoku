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
    assert set("☐□▢⬜❏❐❑❒◻◽") == set(chars.EMPTY_BOXES)
    assert not chars.CERTAIN_EMPTY_BOXES  # 字だけで記入の箱と決まる字はない
    assert chars.BULLET_LIKE_BOXES == chars.EMPTY_BOXES
    assert set("☑☒") == set(chars.CHECKED_BOXES)
    assert "▣" in chars.MAYBE_CHECKED and "▣" not in chars.BOXES  # 記入済みの箱のことも、箇条の印のこともある
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
