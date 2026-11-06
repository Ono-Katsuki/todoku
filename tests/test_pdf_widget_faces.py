"""PDF：欄の面の字（ボタンの字・電子署名の印の字）を読み順に出す。電子署名の値は、検証のバイト列でなく人の読む項目だけ。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402

FONT = "/Font << /F1 8 0 R >>"


def _blob() -> bytes:
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R 5 0 R] /SigFlags 3 /DA (/F1 0 Tf 0 g) >> >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Annots [4 0 R 5 0 R] /Resources << " + FONT.encode() + b" >> >>",
        b"<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /T (R1) /Rect [100 700 160 720] /F 4 /MK << /CA (Reset) >> >>",
        b"<< /Type /Annot /Subtype /Widget /FT /Sig /T (sig1) /V 6 0 R /Rect [100 600 300 660] /F 4 /AP << /N 7 0 R >> >>",
        b"<< /Type /Sig /Filter /Adobe.PPKLite /Name (Fictional Signer) /M (D:20260101000000Z) /ByteRange [0 1 2 3] /Contents <0082AABB> >>",
        stream("/Type /XObject /Subtype /Form /BBox [0 0 200 60] /Resources << /XObject << /FRM 9 0 R >> >>", b"q /FRM Do Q"),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
        stream(f"/Type /XObject /Subtype /Form /BBox [0 0 200 60] /Resources << {FONT} >>", b"BT /F1 9 Tf 2 30 Td (Digitally signed by Fictional Signer) Tj ET"),
    ]
    return pdf(objs, 1)


def _body(tmp_path) -> str:
    p = tmp_path / "a.pdf"
    p.write_bytes(_blob())
    text = "\n".join(s for _n, s in api.window(str(p))[2])
    return text


def test_button_and_signature_faces_are_in_the_reading_order(tmp_path):
    text = _body(tmp_path)
    assert any(ln.startswith("- ［R1］ Reset") for ln in text.splitlines())
    assert any("Digitally signed by Fictional Signer" in ln and "sig1" in ln for ln in text.splitlines() if ln.startswith("- ") and "署名（電子署名）" not in ln)


def test_signature_value_shows_people_facts_not_bytes(tmp_path):
    text = _body(tmp_path)
    assert "Name Fictional Signer" in text and "M D:20260101000000Z" in text
    assert "ByteRange" not in text and "Contents" not in text
