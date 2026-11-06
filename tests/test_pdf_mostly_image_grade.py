"""PDF の読む段階：ページの大半が画像で字がわずかなページ（字が少しだけあるスキャン）を A にしない。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402

WORDS = b"BT /F1 10 Tf 60 780 Td 14 TL " + b" ".join(b"(Applicant name and address of the person who applies for this permit today) '" for _ in range(45)) + b" ET "


def _grade(tmp_path, body: bytes):
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 6 0 R >> >> >>",
            stream("", body),
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            stream("/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", b"\x80")]
    p = tmp_path / "a.pdf"
    p.write_bytes(pdf(objs, 1))
    return api.grade(str(p))


def test_scan_with_a_few_printed_characters_is_not_a(tmp_path):
    g = _grade(tmp_path, b"q 595 0 0 800 0 20 cm /Im1 Do Q BT /F1 8 Tf 500 5 Td (CamScanner) Tj ET")
    assert g["grade"] == "C"
    assert any("大半が画像" in r["title"] for r in g["reasons"])


def test_text_page_with_a_small_logo_stays(tmp_path):
    g = _grade(tmp_path, WORDS + b"q 60 0 0 60 500 760 cm /Im1 Do Q")
    assert not any("大半が画像" in r["title"] for r in g["reasons"])


def test_page_with_half_or_less_image_is_not_flagged(tmp_path):
    g = _grade(tmp_path, b"q 595 0 0 400 0 0 cm /Im1 Do Q BT /F1 8 Tf 60 800 Td (Short heading) Tj ET")
    assert not any("大半が画像" in r["title"] for r in g["reasons"])
