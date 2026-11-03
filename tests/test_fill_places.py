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


def _docx_win(body: str):
    from test_docx import make_docx
    from todoku.docx.document import DocxDocument
    from todoku.docx.window import build, to_markdown

    win = build(DocxDocument(make_docx(body), "t.docx"))
    return {f["id"]: f for f in win["fields"]}, to_markdown(win)


def test_docx_first_empty_paragraph_in_cell_and_head_of_date_chain():
    from test_docx import p

    # 世田谷区の様式：氏名のマスの空の 1 行目（後ろの段落「個人番号」だけに字がある）
    body = ('<w:tbl><w:tr><w:tc>' + p("氏名") + '</w:tc><w:tc><w:p/>' + p("個人番号") + '</w:tc></w:tr></w:tbl>'
            + p("年　　月　　日") + p("氏　　名") + p("重　度　訪　問　介　護"))
    by, md = _docx_win(body)
    first = [f for f in by.values() if f.get("agent") and "初めの空の段落" in (f.get("why") or "")]
    assert first, by
    # 「年　　月　　日」で始まる段落の、初めの字の前（年を書く所）。2 字の「氏　　名」には出さない
    heads = [f for f in by.values() if "初めの字の前" in (f.get("why") or "")]
    assert len(heads) == 1 and heads[0]["label"].startswith("［"), heads
    # 字間を空けた 4 字以上の並び：id にせず〔字間〕
    assert "重　度　訪　問　介　護〔字間〕" in md
    assert not any("重" in (f.get("label") or "") and f.get("agent") for f in by.values())


def test_xlsx_head_of_date_chain_and_spaced_heading():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["B2"] = "年　　月　　日"
        ws["B3"] = "事　業　所　の　住　所"

    f = _xlsx(build)
    assert f["S!B2"]["blanks"] == 3 and "初めの字の前" in f["S!B2"]["why"], f["S!B2"]
    assert "S!B3" not in f


def _pdf_win(lines, rules: bytes = b""):
    from test_circle import _pdf
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.window import build, to_markdown

    win = build(PdfDocument(_pdf(lines, rules)))
    pg = win["pages"][0]
    return pg["fields"], "\n".join(pg["reading"]), to_markdown(win)


def test_pdf_spaced_heading_is_a_mark_not_places_but_two_letter_gap_stays():
    fields, reading, md = _pdf_win([(60, 700, 10, "重　度　訪　問　介　護"), (60, 660, 10, "第　　　号"),
                                    (60, 620, 10, "年　　月　　日"), (400, 580, 10, "備考")])
    # 4 字以上を同じ空白でつないだ並び：字の間は所（id）にせず、元の空白のまま「〔字間〕」の印
    assert "重　度　訪　問　介　護〔字間〕" in reading, reading
    assert "〔字間〕" in md and "字間を空けた見出しの形" in md
    # 2 字・3 字の並び（「第　　号」「年　　月　　日」）は、これまでどおり「？」の所
    assert any(f.get("agent") and f["label"].startswith("第") for f in fields), fields
    assert sum(1 for f in fields if f.get("agent") and "年" in f["label"] and "月" in f["label"]) >= 2


def test_pdf_band_below_top_line_and_free_line_room_are_offered():
    rules = b"0.8 w 60 640 m 400 640 l S 60 600 m 400 600 l S 60 640 m 60 600 l S 400 640 m 400 600 l S\n"
    # 葛飾区の様式：上の行「〒　…　電話番号」、右下に「(日中連絡先)」。住所は〒の行の下の、左の空きに書く
    fields, reading, _md = _pdf_win([(62, 628, 10, "〒"), (300, 628, 10, "電話番号"), (300, 603, 10, "(日中連絡先)"),
                                     (60, 500, 10, "申請者　　住　　所"), (500, 460, 10, "以上")], rules)
    band = [f for f in fields if "上の行" in (f.get("why") or "")]
    assert band and band[0]["rect"][0] < 70 and band[0]["rect"][2] < 300, fields
    room = [f for f in fields if "行の右の余白" in (f.get("why") or "")]
    assert room and room[0]["rect"][0] > 150, fields


def test_pdf_lead_room_before_year_is_not_taken_for_an_indent():
    # マスの中：上の行「年　　月　　日」は右に寄せてあり、下の行「申請日…」が左から始まる（3 字ぶん以上の差：字下げではない）
    rules = b"0.8 w 60 640 m 260 640 l S 60 600 m 260 600 l S 60 640 m 60 600 l S 260 640 m 260 600 l S\n"
    fields, reading, _md = _pdf_win([(110, 625, 10, "年　　月　　日"), (62, 605, 10, "申請日　　年　月　日")], rules)
    lead = [f for f in fields if "行の頭の前" in (f.get("why") or "") and f["rect"][2] <= 110]
    assert lead, fields


def test_pdf_weak_candidates_keep_ids_and_are_listed_compactly():
    fields, reading, md = _pdf_win([(60, 700, 10, "区分　１　２　３　４　５"), (400, 660, 10, "備考")])
    weak = [f for f in fields if f.get("weak")]
    assert weak and all(f.get("agent") for f in weak)
    assert all(f"［{f['id']}？" in reading for f in weak)  # 読み順には、ほかの所と同じ印で出ている
    assert "弱い候補" in md


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
