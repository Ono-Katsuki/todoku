"""Word：空白だけのふりがな（w:rt）は、窓に「なし（　）」のような字を作らない。読みのあるふりがなは、これまでどおり。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from test_word_font_marks import _md  # noqa: E402


def _ruby(base: str, rt: str) -> str:
    return (f"<w:p><w:r><w:t>障害：</w:t></w:r><w:r><w:ruby><w:rubyPr/><w:rt><w:r><w:t xml:space=\"preserve\">{rt}</w:t></w:r></w:rt>"
            f"<w:rubyBase><w:r><w:t>{base}</w:t></w:r></w:rubyBase></w:ruby></w:r><w:r><w:t>・あり</w:t></w:r></w:p>")


def test_blank_ruby_adds_no_characters():
    md = _md(_ruby("なし", "\u3000"))
    assert "障害：なし・あり" in md and "なし（" not in md


def test_ruby_with_reading_is_kept():
    assert "氏名（しめい）" in _md(_ruby("氏名", "しめい"))
