"""第 8 回に残った所：near の作りの近さ・効かない align・セルの字の色・同じ内容の警告・Word のページの増え。"""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from todoku import api
from todoku.api import ApiError, _align_no_effect, _merged_notes

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def docx(tmp_path, body: str, name: str = "a.docx", app: str = "") -> Path:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   + ('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' if app else "")
                   + "</Relationships>")
        if app:
            z.writestr("docProps/app.xml", app)
        z.writestr("word/document.xml", f'<?xml version="1.0"?><w:document {W}><w:body>{body}'
                   '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1985" w:right="1701" w:bottom="1701" w:left="1701"/></w:sectPr></w:body></w:document>')
    p = tmp_path / name
    p.write_bytes(buf.getvalue())
    return p


def p(text: str = "") -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>' if text else "<w:p/>"


def row(*cells: str) -> str:
    return "<w:tr>" + "".join(f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{c}</w:tc>' for c in cells) + "</w:tr>"


def tbl(*rows: str) -> str:
    return '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>' + "".join(rows) + "</w:tbl>"


# ---------------------------------------------------------------- before + near
def _near_doc(tmp_path):
    # 同じ行の見出しのマスにも「年」があり、near の欄と同じマスの段落は「年」で始まる（文京区の様式の形）
    return docx(tmp_path, tbl(row(p("生年月日"), p("年　　月　　日") + p("氏名：　　　　"))))


def _near_id(path) -> str:
    return next(f["id"] for f in api.fields(str(path)) if "月" in str(f.get("label")))


def test_before_near_takes_the_candidate_in_nears_own_paragraph(tmp_path):
    src = _near_doc(tmp_path)
    near = _near_id(src)
    r = api.fill(str(src), {"@1": {"before": "年", "near": near, "value": "昭和60"}}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"]
    win = "\n".join(t for _n, t in api.window(str(tmp_path / "o.docx"))[2])
    assert "昭和60" in win and "生昭和60" not in win


def test_occurrence_still_counts_in_the_wide_scope(tmp_path):
    src = _near_doc(tmp_path)
    near = _near_id(src)
    r = api.fill(str(src), {"@1": {"before": "年", "near": near, "occurrence": 1, "value": "X"}}, str(tmp_path / "o.docx"))
    win = "\n".join(t for _n, t in api.window(str(tmp_path / "o.docx"))[2])
    assert r["proof_ok"] and "生［P1.w1＝X］年月日" in win


def test_two_candidates_in_nears_own_paragraph_are_still_refused(tmp_path):
    src = docx(tmp_path, tbl(row(p("生年月日"), p("年　　年　　日"))))
    near = next(f["id"] for f in api.fields(str(src)) if f["id"].startswith("P2."))
    with pytest.raises(ApiError, match="2 か所"):
        api.fill(str(src), {"@1": {"before": "年", "near": near, "value": "X"}}, str(tmp_path / "o.docx"))


# ---------------------------------------------------------------- align
def test_align_on_an_empty_paragraph_says_it_did_nothing(tmp_path):
    src = docx(tmp_path, tbl(row(p("時間"), p())))
    cid = next(f["id"] for f in api.fields(str(src)))
    r = api.fill(str(src), {cid: {"value": "50", "align": "right"}}, str(tmp_path / "o.docx"))
    # マスの欄に align：受けて足す空白がない、または形の違いで断る。受けたら黙らない
    assert r["aligned"][cid]["pad"] == [0, 0]
    assert any("効いていません" in w for w in r["warnings"])
    assert "効いていません" in r["say_to_user"]


def test_align_with_room_has_no_such_warning(tmp_path):
    assert _align_no_effect({"a": {"align": "right", "pad": [4, 0]}, "b": {"align": "left", "pad": [0, 0]}}) == []
    assert len(_align_no_effect({"a": {"align": "center", "pad": [0, 0]}})) == 1


# ---------------------------------------------------------------- 伝えることをまとめる
def test_merged_notes_keeps_places_and_says_numbers_differ():
    a = "書く所に残っていた書式は、書いた字に引き継がず書きました"
    got = _merged_notes([f"P1.c1: {a}", f"P2.c1: {a}", "p1.w1：重なります（左 1.5 mm）", "p1.w2：重なります（左 2.5 mm）", "ひとつだけ"])
    assert got[0] == f"2 か所（P1.c1・P2.c1）：{a}"
    assert "2 か所（p1.w1・p1.w2）：重なります（左 1.5 mm）（数字は 1 か所目のもの" in got[1]
    assert got[2] == "ひとつだけ"


# ---------------------------------------------------------------- ページの増え
APP = '<?xml version="1.0"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Pages>1</Pages></Properties>'


def test_page_growth_states_the_recorded_pages_and_the_lines(tmp_path):
    src = docx(tmp_path, tbl(row(p("氏名"), p())), app=APP)
    cid = next(f["id"] for f in api.fields(str(src)))
    r = api.fill(str(src), {cid: "とても長い値" * 12}, str(tmp_path / "o.docx"))
    assert r["wrap"] and r["page_growth"]["recorded_pages_before"] == 1 and r["page_growth"]["lines_added"] >= 1
    assert "ページが増えるかもしれません" in r["say_to_user"] and "1 ページ" in r["say_to_user"]


def test_no_page_note_without_new_lines(tmp_path):
    src = docx(tmp_path, tbl(row(p("氏名"), p())), app=APP)
    cid = next(f["id"] for f in api.fields(str(src)))
    r = api.fill(str(src), {cid: "山田"}, str(tmp_path / "o.docx"))
    assert "page_growth" not in r and "ページが増える" not in r["say_to_user"]


def test_page_note_without_a_record_says_so(tmp_path):
    src = docx(tmp_path, tbl(row(p("氏名"), p())))
    cid = next(f["id"] for f in api.fields(str(src)))
    r = api.fill(str(src), {cid: "とても長い値" * 12}, str(tmp_path / "o.docx"))
    assert r["page_growth"]["recorded_pages_before"] is None and "ページ数の記録はありません" in r["say_to_user"]


def test_doc_summary_page_count_is_read_as_the_property_set_says():
    import struct

    from todoku.pagecount import _doc_pages

    class Fake:
        def __init__(self, data):
            self.d = data

        def has(self, _n):
            return True

        def read(self, _n):
            return self.d

    body = struct.pack("<II", 16 + 8, 1) + struct.pack("<II", 14, 16) + struct.pack("<Ii", 3, 2)
    d = struct.pack("<HHI", 0xFFFE, 0, 0) + b"\0" * 16 + struct.pack("<I", 1) + b"\1" * 16 + struct.pack("<I", 48) + body
    import todoku.cfb as cfb
    keep = cfb.CFB
    cfb.CFB = lambda _b: Fake(d)
    try:
        assert _doc_pages(b"") == 2
    finally:
        cfb.CFB = keep


# ---------------------------------------------------------------- Excel の字の色
def test_excel_cell_colour_is_stated(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "氏名"
    ws["B1"].font = Font(color="FFFF0000")
    ws["B2"].font = Font(color="FF000000")
    src = tmp_path / "c.xlsx"
    wb.save(src)
    r = api.fill(str(src), {"Sheet!B1": "山田", "Sheet!B2": "花子"}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"]
    assert [x["id"] for x in r["written_formatting"]] == ["Sheet!B1"]
    assert "#FF0000" in r["say_to_user"]
