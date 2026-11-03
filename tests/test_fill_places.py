"""書く所の見つけ方（72 の様式の試し書きから）：見出しは印刷された字そのまま・字間を空けた見出しは所にしない・
弱い候補の印・見落としていた所（マスの中の上の行の下の空き、行の頭の空き、空の段落など）。"""

from __future__ import annotations

import io

import pytest

from todoku.xlsx import window
from todoku.xlsx.workbook import Workbook

openpyxl = pytest.importorskip("openpyxl")
from openpyxl.styles import Protection  # noqa: E402


def _xlsx(build) -> dict:
    wb = openpyxl.Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    doc = window.build(Workbook(buf.getvalue(), "t.xlsx"))
    return {f["id"]: f for sh in doc["sheets"] for f in sh["fields"]}


def test_phrase_label_does_not_take_other_cells_text_as_prefix():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        for ref, v in {"A4": "なし・あり（１・２・３）", "B4": "計画作成日", "D4": "令和", "F4": "年", "H4": "月"}.items():
            ws[ref] = v
        for ref in ("C4", "E4", "G4"):
            ws[ref].protection = Protection(locked=False)
        ws.protection.sheet = True

    f = _xlsx(build)
    lab = f["S!E4"]["label"]
    # 見出しは、並びの字とすぐ前のマスの見出しだけ。左の別のマスの字（「なし・あり…」）を見出しの前に付けない
    assert "なし・あり" not in lab and "令和［▶］年" in lab
    assert "なし・あり" in (f["S!E4"].get("section") or "")


def test_pdf_circle_occurrence_counts_among_near_candidates(tmp_path):
    from todoku import api
    from test_circle import _pdf

    # 「通院頻度 月・週」の行と、上の「生年月日 年・月・日」の行。near と occurrence を合わせると、near に近い（同じ行の）
    # 候補の中の何番目か（Word・Excel と同じ）。文書全体の 1 番目（上の行の「月」）にはしない
    data = _pdf([(70, 700, 10, "生年月日"), (150, 700, 10, "年・月・日"), (70, 650, 10, "通院頻度"), (150, 650, 10, "月・週")])
    src = tmp_path / "t.pdf"
    src.write_bytes(data)
    r = api.fill(str(src), {"@c1": {"circle": "月", "near": "通院頻度", "occurrence": 1}}, str(tmp_path / "o.pdf"))
    assert "通院頻度 【月】・週" in r["circled"][0]["around"], r["circled"]
    with pytest.raises(api.ApiError, match="near"):
        api.fill(str(src), {"@c1": {"circle": "月", "near": "通院頻度", "occurrence": 2}}, str(tmp_path / "p.pdf"))
