"""PDF：AcroForm にあるが、どのページにも置かれていない欄の値を、窓が黙って落とさない（書く欄にはしない）。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf  # noqa: E402

from todoku import api  # noqa: E402


def _blob() -> bytes:
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R 5 0 R] /DA (/Helv 0 Tf 0 g) >> >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] >>",
        b"<< /Type /Annot /Subtype /Widget /FT /Tx /T (Adress) /V (@postadress) /Rect [10 10 100 30] /F 4 >>",
        b"<< /Type /Annot /Subtype /Widget /FT /Tx /T (Tom) /Rect [10 40 100 60] /F 4 >>",
    ]
    return pdf(objs, 1)


def _src(tmp_path):
    p = tmp_path / "a.pdf"
    p.write_bytes(_blob())
    return p


def test_value_of_a_field_on_no_page_is_told(tmp_path):
    _k, doc, parts = api.window(str(_src(tmp_path)))
    text = "\n".join(s for _n, s in parts)
    assert "どのページにも置かれていない欄が 2 個" in text
    assert "「Adress」＝@postadress" in text
    assert {"name": "Tom", "value": ""} in doc["form"]["unplaced"]


def test_field_on_no_page_does_not_become_a_place_to_write(tmp_path):
    ids = {f["id"] for f in api.fields(str(_src(tmp_path)))}
    assert "Adress" not in ids
