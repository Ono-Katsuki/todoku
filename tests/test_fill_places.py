"""書く所の見つけ方（72 の様式の試し書きから）：見出しは印刷された字そのまま・見落としていた所（マスの中の上の行の下の空き、
行の頭の空き、空の段落など）。PDF は、字間を空けた見出しの字の間も 1 字ぶん以上の空きなら所（形は facts）、弱い候補の印はない。"""

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
    # 「年　　月　　日」の初めの字の前は、所にしない（後ろの所の数で決めていた決まりをやめた）
    assert not any("初めの字の前" in (f.get("why") or "") for f in by.values())
    # 字間を空けた並び（重　度　訪　問　介　護）の字の間も、隠さずに「？」の所にし、並びの字数を事実（facts の chain）で示す。
    # 見出しか書く所かは、エージェントが判断する
    spaced = [f for f in by.values() if f.get("agent") and "重" in (f.get("label") or "")]
    assert spaced and all((f.get("facts") or {}).get("chain", {}).get("len") == 6 for f in spaced), spaced
    assert "〔字間〕" not in md


def test_xlsx_head_of_date_chain_and_spaced_heading():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["B2"] = "年　　月　　日"
        ws["B3"] = "事　業　所　の　住　所"

    f = _xlsx(build)
    # 字の間の空き（「？」）だけ。初めの字の前（以前の chain_head：後ろに所が 2 つ以上あるときだけ足していた）は所にしない。
    # 1 字ずつ同じ空白で並ぶ形は、事実（facts の chain：字数と字）として出す
    b2 = f["S!B2"]
    assert b2["blanks"] == 2 and b2["agent"], b2
    assert b2["spans"][0]["facts"]["chain"] == {"len": 3, "text": "年　　月　　日"}
    # 1 字ずつ同じ空白でつないだ 6 字の並び：以前は「字間を空けた見出し」として隠していた（4 字以上）。いまは字の間を
    # 「？」（spaced）として出し、並びの字数を事実にする（見出しか書く所かはエージェントが判断する）
    b3 = f["S!B3"]
    assert b3["blanks"] == 5 and {s["why_key"] for s in b3["spans"]} == {"spaced"}
    assert b3["spans"][0]["facts"]["chain"]["len"] == 6


def _pdf_win(lines, rules: bytes = b""):
    from test_circle import _pdf
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.window import build, to_markdown

    win = build(PdfDocument(_pdf(lines, rules)))
    pg = win["pages"][0]
    return pg["fields"], "\n".join(pg["reading"]), to_markdown(win)


def test_pdf_spaced_heading_gaps_are_places_with_the_shape_as_a_fact():
    fields, reading, md = _pdf_win([(60, 700, 10, "重　度　訪　問　介　護"), (60, 660, 10, "第　　　号"),
                                    (60, 620, 10, "年　　月　　日"), (400, 580, 10, "備考")])
    # 1 字ずつを同じ空白でつないだ並び（字間を空けた見出しの形）：字の間も 1 字ぶん以上の空き。見出しかどうか（形の意味）は
    # 決めず、形を facts の spaced と窓の「字間を空けた形」で示す。並びの長さで所にするかを変えない
    spaced = [f for f in fields if f.get("shape") == "gap" and f["facts"]["line"] == "重　度　訪　問　介　護"]
    assert len(spaced) == 5 and all(f["agent"] and f["facts"]["spaced"] for f in spaced), spaced
    assert "〔字間〕" not in reading and "字間を空けた形" in md
    # 2 字・3 字の並び（「第　　号」「年　　月　　日」）は、これまでどおり「？」の所
    assert any(f.get("agent") and f["label"].startswith("第") for f in fields), fields
    assert sum(1 for f in fields if f.get("agent") and "年" in f["label"] and "月" in f["label"]) >= 2


def test_pdf_band_below_top_line_and_free_line_room_are_offered():
    rules = b"0.8 w 60 640 m 400 640 l S 60 600 m 400 600 l S 60 640 m 60 600 l S 400 640 m 400 600 l S\n"
    # 葛飾区の様式：上の行「〒　…　電話番号」、右下に「(日中連絡先)」。住所は〒の行の下の、左の空きに書く
    fields, reading, _md = _pdf_win([(62, 628, 10, "〒"), (300, 628, 10, "電話番号"), (300, 603, 10, "(日中連絡先)"),
                                     (60, 500, 10, "申請者　　住　　所"), (500, 460, 10, "以上")], rules)
    # 「〒」と「電話番号」は同じ高さの 1 つの行：その間の 1 字ぶん以上の空き（住所を書く所）
    post = next(f for f in fields if f.get("shape") == "gap" and f["facts"]["before"] == "〒")
    assert post["rect"][0] < 80 and post["rect"][2] <= 300 and post["facts"]["after"] == "電話番号", post
    # 下の行「(日中連絡先)」の前の空き（マスの縁まで）：上下は、上の行の字からマスの縁まで（行と行の間の帯）
    head = next(f for f in fields if f.get("shape") == "lead" and f["facts"]["line"] == "(日中連絡先)")
    assert head["facts"]["bound"]["kind"] == "cell" and head["rect"][3] - head["rect"][1] > 10, head
    # マスの外の行の終わりの後ろ（ページの字の端まで）の空きも所
    room = [f for f in fields if f.get("shape") == "trail" and f["facts"]["line"].startswith("申請者")]
    assert room and room[0]["rect"][0] > 150 and room[0]["facts"]["bound"]["kind"] == "page", fields


def test_pdf_lead_room_before_year_is_not_taken_for_an_indent():
    # マスの中：上の行「年　　月　　日」は右に寄せてあり、下の行「申請日…」が左から始まる（3 字ぶん以上の差：字下げではない）
    rules = b"0.8 w 60 640 m 260 640 l S 60 600 m 260 600 l S 60 640 m 60 600 l S 260 640 m 260 600 l S\n"
    fields, reading, _md = _pdf_win([(110, 625, 10, "年　　月　　日"), (62, 605, 10, "申請日　　年　月　日")], rules)
    lead = [f for f in fields if f.get("shape") == "lead" and f["rect"][2] <= 110]
    assert lead, fields


def test_pdf_narrow_gaps_are_not_ranked_but_measured():
    # 1 字ぶんちょうどの空き（「区分　１　２」）も、ほかの空きと同じ所：弱い・強いの順は付けず、測った幅（mm・字）を示す
    fields, reading, md = _pdf_win([(60, 700, 10, "区分　１　２　３　４　５"), (400, 660, 10, "備考")])
    gaps = [f for f in fields if f.get("shape") == "gap"]
    assert len(gaps) == 5 and not any(f.get("weak") for f in fields)
    assert all(f["facts"]["width_em"] == 1.0 and f"［{f['id']}？" in reading for f in gaps)
    assert "弱い候補" not in md and "1.0字" in md


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


def test_docx_first_empty_paragraph_takes_the_cell_label(tmp_path):
    from test_docx import make_docx, p

    from todoku import api

    body = '<w:tbl><w:tr><w:tc>' + p("氏名") + '</w:tc><w:tc><w:p/>' + p("個人番号") + '</w:tc></w:tr></w:tbl>'
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(body))
    first = next(f for f in api.fields(str(src)) if "初めの空の段落" in (f.get("why") or ""))
    assert first["label"] == "氏名"  # 行の見出し（印刷された字のまま）。「［　］」だけにしない
    r = api.fill(str(src), {first["id"]: "山田　花子"}, str(tmp_path / "b.docx"), overwrite=True)
    assert "（近くの字「氏名」）：山田　花子" in r["say_to_user"], r["say_to_user"]
