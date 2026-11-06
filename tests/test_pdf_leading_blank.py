"""頭に改行のある PDF も PDF として開く（ファイルの種類を「不明」にして断っていた）。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402
from todoku.book import kind_of  # noqa: E402


def _blob() -> bytes:
    return pdf([b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
                b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
                stream("", b"BT /F1 12 Tf 100 700 Td (Hello form) Tj ET"),
                b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"], 1)


def test_blank_before_the_header_is_still_a_pdf(tmp_path):
    p = tmp_path / "a.pdf"
    p.write_bytes(b"\n" + _blob())
    assert kind_of(p.read_bytes(), "a.pdf") == "pdf"
    assert "Hello form" in "\n".join(s for _n, s in api.window(str(p))[2])


def test_text_that_mentions_the_header_is_not_a_pdf():
    assert kind_of(b"memo: %PDF-1.4 is the header\n", "a.txt") != "pdf"
