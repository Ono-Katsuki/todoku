"""PDF：左へ90度回して印刷した字（余白の様式番号など。下から上へ進む）を、字の順のまま 1 行に読む。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402


def _blob(body: bytes) -> bytes:
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        stream("", body),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    ]
    return pdf(objs, 1)


def _window(tmp_path, body: bytes) -> str:
    p = tmp_path / "a.pdf"
    p.write_bytes(_blob(body))
    return "\n".join(s for _n, s in api.window(str(p))[2])


def test_text_rotated_left_reads_in_its_own_order(tmp_path):
    body = b"BT /F1 12 Tf 100 700 Td (Name of the applicant) Tj ET BT /F1 6 Tf 0 1 -1 0 20 300 Tm (Form 123 rev two) Tj ET"
    text = _window(tmp_path, body)
    assert "Form 123 rev two" in text and "Name of the applicant" in text


def test_text_rotated_right_still_reads_down(tmp_path):
    body = b"BT /F1 6 Tf 0 -1 1 0 20 600 Tm (Form 123 rev two) Tj ET"
    assert "Form 123 rev two" in _window(tmp_path, body)
