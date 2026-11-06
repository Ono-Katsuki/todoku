"""PDF：字のないマスに画像が重なるとき、その事実を所の facts と窓で伝える（空のマスとだけ言わない）。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402

BOX = b"0.5 w 100 600 150 100 re S 300 600 150 100 re S "


def _blob(image: bool) -> bytes:
    body = BOX + b"BT /F1 10 Tf 100 720 Td (Crest) Tj 200 0 Td (Blank) Tj ET "
    if image:
        body += b"q 80 0 0 60 120 620 cm /Im1 Do Q"
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 6 0 R >> >> >>",
            stream("", body),
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
            stream("/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", b"\x00")]
    return pdf(objs, 1)


def _cells(tmp_path, image: bool) -> list[dict]:
    p = tmp_path / "a.pdf"
    p.write_bytes(_blob(image))
    doc = api.window(str(p))[1]
    return [f for f in doc["pages"][0]["fields"] if f.get("shape") == "cell"]


def test_only_the_cell_under_the_image_says_so(tmp_path):
    cells = _cells(tmp_path, True)
    assert sorted(bool(f["facts"].get("on_image")) for f in cells) == [False, True]


def test_no_image_no_fact(tmp_path):
    assert not any(f["facts"].get("on_image") for f in _cells(tmp_path, False))
