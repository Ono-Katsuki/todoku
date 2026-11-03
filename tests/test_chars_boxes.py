"""チェックの箱の字の分け方（chars.py）：記入の箱と決まる字と、箇条書きの印にもよく使う字（BULLET_LIKE_BOXES）。
PDF の窓は、箇条の印かもしれない箱の字を、規則のチェックの欄にせず「？」にする。"""

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
    assert set("☐□▢⬜") == set(chars.CERTAIN_EMPTY_BOXES)
    assert set("❏❐❑❒◻◽") == set(chars.BULLET_LIKE_BOXES)
    assert not chars.CERTAIN_EMPTY_BOXES & chars.BULLET_LIKE_BOXES
    assert chars.EMPTY_BOXES == chars.CERTAIN_EMPTY_BOXES | chars.BULLET_LIKE_BOXES
    assert set("☑☒") == set(chars.CHECKED_BOXES)
    assert "▣" in chars.MAYBE_CHECKED and "▣" not in chars.BOXES  # 記入済みの箱のことも、箇条の印のこともある
    assert all(chars.is_bullet_like_box(ch) for ch in "❏❐❑❒◻◽")
    assert not chars.is_bullet_like_box("□") and not chars.is_bullet_like_box("☐")


def _boxes(text: str) -> tuple[list[dict], str]:
    pg = window(PdfDocument(text_pdf([(text, 50, 700)]), "a.pdf"))["pages"][0]
    return [f for f in pg["fields"] if f["kind"] == "checkbox"], "\n".join(pg["reading"])


@pytest.mark.parametrize("ch", list("❏❐❑❒◻◽"))
def test_bullet_like_box_is_an_agent_place(ch):
    boxes, text = _boxes(f"{ch}はい")
    assert boxes and boxes[0].get("agent"), boxes
    assert "箇条書きの印にもよく使う字" in boxes[0]["why"] and f"U+{ord(ch):04X}" in boxes[0]["why"]
    assert boxes[0]["why"].count("箇条書きの印") == 1, boxes[0]["why"]  # 理由は 1 度だけ（重ねてつけない）
    assert f"［{boxes[0]['id']}？］" in text


@pytest.mark.parametrize("ch", list("☐□▢⬜"))
def test_certain_box_is_a_rule_checkbox(ch):
    boxes, _text = _boxes(f"{ch}はい")
    assert boxes and not boxes[0].get("agent"), boxes
