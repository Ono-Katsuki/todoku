"""PDF：大きな見出しの横に小さな字の 2 行があるとき、2 行の字を 1 字ずつ交互に混ぜない。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402


def _blob(lines: list[tuple[float, float, float, str]]) -> bytes:
    body = b" ".join(f"BT /F1 {s} Tf {x} {y} Td ({t}) Tj ET".encode() for x, y, s, t in lines)
    objs = [
        b"<< /Type /Catalog /Pages 2 0 R >>",
        b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>",
        stream("", body),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>",
    ]
    return pdf(objs, 1)


def _window(tmp_path, lines) -> str:
    p = tmp_path / "a.pdf"
    p.write_bytes(_blob(lines))
    return "\n".join(s for _n, s in api.window(str(p))[2])


def test_two_small_lines_beside_a_big_heading_are_not_interleaved(tmp_path):
    text = _window(tmp_path, [(100, 763, 6.5, "AYUNTAMIENTO DE CACERES"), (100, 755.8, 6.5, "OFICINA DE ASUNTOS"),
                              (260, 759, 16, "SOLICITUD DE PRESTACION")])
    assert "AYUNTAMIENTO DE CACERES" in text and "OFICINA DE ASUNTOS" in text


def test_superscript_stays_on_its_line(tmp_path):
    text = _window(tmp_path, [(100, 700, 12, "Area in m"), (150, 704, 7, "2"), (156, 700, 12, "is measured")])
    assert "Area in m2 is measured" in text.replace("  ", " ")
