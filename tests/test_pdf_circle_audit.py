"""書いてある○の検査（pdf/audit.check_circles）：窓と書くときが同じ決め方（circle.encloses）で見ていることの照らし合わせ
（整合の検査）だと detail に書くこと。別に、描画の命令から直接、○の外枠と字の箱の中心で数え直すこと（raw_circles）。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from test_circle import _table_pdf, _write  # noqa: E402

from todoku.pdf import audit as pdf_audit  # noqa: E402
from todoku.pdf import circle  # noqa: E402
from todoku.pdf.content import page_content  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402


def _filled(tmp_path) -> PdfDocument:
    r, out = _write(tmp_path, _table_pdf(), "t.pdf", {"@c1": {"circle": "無", "near": "障害支援区分の認定"},
                                                      "@c2": {"circle": "２", "near": "性別"}})
    assert r["proof_ok"] and r["read_back_ok"]
    return PdfDocument(open(out, "rb").read())


def test_raw_circles_match_the_drawn_shapes_in_page_coordinates(tmp_path):
    doc = _filled(tmp_path)
    page = doc.pages()[0]
    raw = pdf_audit.raw_circles(doc, page)
    assert set(raw) == {i for i, *_x in circle.drawn(doc, 1)} == {"p1.c1", "p1.c2"}
    glyphs = page_content(doc, page).glyphs
    assert pdf_audit.raw_circled_text(glyphs, *raw["p1.c1"]) == "無"
    assert pdf_audit.raw_circled_text(glyphs, *raw["p1.c2"]) == "２"


def test_check_says_it_is_a_consistency_check_and_recounts_independently(tmp_path):
    c = pdf_audit.check_circles(_filled(tmp_path))
    assert c.ok and not c.agent, c
    assert "整合の検査" in c.detail and "正しさは確かめません" in c.detail
    assert "描画の命令から直接" in c.detail and "どれも窓の囲んだ字と同じでした" in c.detail


def test_independent_recount_that_disagrees_is_shown_for_the_agent(tmp_path, monkeypatch):
    doc = _filled(tmp_path)
    monkeypatch.setattr(pdf_audit, "raw_circled_text", lambda glyphs, kind, box: "有無")
    c = pdf_audit.check_circles(doc)
    assert c.ok and c.agent  # 値は窓で読めている：どちらが正しいかは画像で確かめる
    assert "2 個は窓の囲んだ字と違いました" in c.detail
    assert any("「有無」" in it and "view_page" in it for it in c.items), c.items
