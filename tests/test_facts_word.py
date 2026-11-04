"""事実を出し、判断はエージェントに任せる（Word と同じ窓の形式・チェックの値・名前の決まり）：作った小さな書類で確かめる。

- チェックの欄の値：true / false と、その欄が宣言した状態の値だけ（言葉の意味では決めない。どの言語の言葉も受けない）
- 段落の中の所の名前：位置（前にある印刷された字の数）で決まり、ほかの所が増えても、前の所に書いた後も変わらない
- 古い数え順の名前（P1.1）は、同じ所を指すと決まるときだけ受ける
- 表のマスの見出し：左・上のマスの字をそのまま、間のマスの数と一緒に出す（距離で片方を落とさない）
- 字を目印にした書き込みの near の範囲：決まった距離の数を持たず、near_paras で指定できる
"""

from __future__ import annotations

import pytest

from test_docx import make_docx, p
from todoku.docx.document import DocxDocument
from todoku.docx.fill import FillError, fill
from todoku.docx.proof import prove
from todoku.docx.window import build
from todoku.inputcheck import InputError, checkbox_state, checkbox_value


# ---------------------------------------------------------------------- チェックの値
@pytest.mark.parametrize("word", ["はい", "あり", "有", "○", "yes", "oui", "sí", "예", "是", "да", "1", "x", "", "on"])
def test_checkbox_words_are_not_read_by_meaning(word):
    assert checkbox_state(word) is None
    with pytest.raises(InputError, match="true / false"):
        checkbox_value(word)


def test_checkbox_accepts_booleans_and_declared_states_only():
    assert checkbox_state(True) is True and checkbox_state(False) is False
    assert checkbox_state("true") is True and checkbox_state("FALSE") is False  # JSON の真偽値の綴り
    assert checkbox_state(1) is None and checkbox_state(None) is None
    # その欄が宣言した状態の値（PDF の状態の名前・XFA の items など）は、字のとおりに受ける
    assert checkbox_state("Ja", on=("Ja",), off=("Off",)) is True and checkbox_state("Off", on=("Ja",), off=("Off",)) is False
    assert checkbox_state("ja", on=("Ja",), off=("Off",)) is None
    with pytest.raises(InputError, match="Ja・Off"):
        checkbox_value("はい", on=("Ja",), off=("Off",))


def test_word_glyph_checkbox_refuses_words():
    blob = make_docx(p("☐ 同意する"))
    with pytest.raises(FillError, match="true / false"):
        fill(DocxDocument(blob, "a.docx"), {"P1.c0": "はい"})
    res = fill(DocxDocument(blob, "a.docx"), {"P1.c0": "☑"})  # 箱の字（宣言した状態の字）は受ける
    assert {f["id"]: f["value"] for f in build(DocxDocument(res.blob, "a.docx"))["fields"]}["P1.c0"] is True


# ---------------------------------------------------------------------- 位置で決まる名前
def _ids(blob: bytes) -> dict[str, str]:
    return {f["id"]: f["label"] for f in build(DocxDocument(blob, "a.docx"))["fields"]}


@pytest.mark.parametrize("text,label", [("Name: ＿＿＿＿", "Name: "), ("姓名：＿＿＿＿", "姓名："), ("성명: ＿＿＿＿", "성명: "),
                                        ("Имя: ＿＿＿＿", "Имя: "), ("الاسم: ＿＿＿＿", "الاسم: ")])
def test_place_name_counts_printed_characters(text, label):
    n = sum(1 for ch in label if not ch.isspace())
    assert f"P1.c{n}" in _ids(make_docx(p(text)))


def test_place_names_do_not_shift_when_other_places_appear_or_are_written():
    a = make_docx(p("氏名＿＿＿＿　住所＿＿＿＿"))
    b = make_docx(p("氏名＿＿＿＿　　　住所＿＿＿＿"))  # 前に「？」の空きが増えても
    ia, ib = _ids(a), _ids(b)
    assert {"P1.c2", "P1.c4"} <= set(ia) and {"P1.c2", "P1.c4"} <= set(ib)
    res = fill(DocxDocument(a, "a.docx"), {"P1.c2": "山田太郎"})  # 前の所に書いた後も、後ろの所の名前は同じ
    assert prove(a, res.blob, {x.id: x.value for x in res.patches}).ok
    after = _ids(res.blob)
    assert "P1.c2" in after and "P1.c4" in after
    res2 = fill(DocxDocument(res.blob, "a.docx"), {"P1.c4": "東京都"})
    got = {f["id"]: f["value"] for f in build(DocxDocument(res2.blob, "a.docx"))["fields"]}
    assert got["P1.c2"] == "山田太郎" and got["P1.c4"] == "東京都"


def test_adjacent_places_get_numbered_suffix():
    ids = _ids(make_docx(p("☐☐☐ 選ぶ")))
    assert {"P1.c0", "P1.c0.2", "P1.c0.3"} <= set(ids)


def test_old_ids_only_when_unambiguous():
    one = make_docx(p("氏名＿＿＿＿"))
    res = fill(DocxDocument(one, "a.docx"), {"P1.1": "山田"})  # その段落にその並びの所が 1 つだけ：同じ所
    assert {f["id"]: f["value"] for f in build(DocxDocument(res.blob, "a.docx"))["fields"]}["P1.c2"] == "山田"
    two = make_docx(p("氏名＿＿＿＿　住所＿＿＿＿"))
    with pytest.raises(FillError, match="P1.c2・"):
        fill(DocxDocument(two, "a.docx"), {"P1.1": "山田"})  # 2 つあるので、古い番号では決めない（いまの名前を示す）


# ---------------------------------------------------------------------- 表のマスの見出し
def _tbl(rows: list[list[str]]) -> str:
    return "<w:tbl>" + "".join("<w:tr>" + "".join(f"<w:tc>{p(c) if c else '<w:p/>'}</w:tc>" for c in r) + "</w:tr>"
                               for r in rows) + "</w:tbl>"


@pytest.mark.parametrize("row,col", [("氏名", "申請者"), ("Name", "Applicant"), ("성명", "신청인"), ("姓名", "申请人")])
def test_cell_gives_both_neighbours_verbatim_with_gaps(row, col):
    blob = make_docx(_tbl([["", col], ["", ""], ["", ""], [row, ""]]))
    f = {x["id"]: x for x in build(DocxDocument(blob, "a.docx"))["fields"]}["T1.R4.C2"]
    # 上の見出しは 2 行あいだが空いていても落とさない（距離は事実として出す）
    assert f["row_label"] == row and f["row_label_gap"] == 0
    assert f["column_label"] == col and f["column_label_gap"] == 2
    assert f["label"] == f"{row}／{col}"


def test_long_label_is_not_cut():
    long = "この欄には、申請者本人の氏名を、住民票に記載されたとおりに、楷書ではっきりと記入してください（外字は近い字で）"
    blob = make_docx(_tbl([[long, ""]]))
    f = {x["id"]: x for x in build(DocxDocument(blob, "a.docx"))["fields"]}["T1.R1.C2"]
    assert f["row_label"] == long and f["label"] == long
    blob = make_docx(p("☐ " + long))
    assert any(x["label"] == long for x in build(DocxDocument(blob, "a.docx"))["fields"])


# ---------------------------------------------------------------------- near の範囲
def test_near_scope_is_an_option():
    body = p("氏名＿＿＿＿") + "".join(p(f"説明 {k}") for k in range(6)) + p("備考　記入") + p("備考　記入")
    blob = make_docx(body)
    # near だけ：near の欄と同じ話（本文）の中。2 か所あるので決まらない
    with pytest.raises(FillError, match="2 か所"):
        fill(DocxDocument(blob, "a.docx"), {"@1": {"old": "備考　記入", "new": "備考　なし", "near": "P1.c2"}})
    # near_paras で範囲をしぼる（near の段落の前後 7 段落：8 段落目だけ）
    res = fill(DocxDocument(blob, "a.docx"), {"@1": {"old": "備考　記入", "new": "備考　なし", "near": "P1.c2", "near_paras": 7}})
    assert [x.id for x in res.patches] == ["P8.w1"]
    with pytest.raises(FillError, match="near_paras"):
        fill(DocxDocument(blob, "a.docx"), {"@1": {"old": "備考　記入", "new": "備考　なし", "near": "P1.c2", "near_paras": -1}})


# ---------------------------------------------------------------------- LibreOffice が書き出した小さな .doc
def test_libreoffice_doc_form_fields_and_small_streams(tmp_path):
    """FFData の位置が欄の命令の中の 0x01 の字にある .doc（LibreOffice）：欄の名前・選択肢を読み、マスの中の欄の見出しは
    いちばん近いマスの字（上のマスの字は column_label に別に出す）。ストリームが 4096 バイト未満（ミニストリーム）でも書ける。"""
    from pathlib import Path

    from todoku import api

    src = str(Path(__file__).parent / "fixtures" / "public" / "form.doc")
    f = {x["id"]: x for x in api.fields(src)}
    assert f["ctl1"]["options"][1:4] == ["新規", "変更", "廃止"] and f["ctl1"]["label"] == "区分"
    assert f["ctl1"]["column_label"] == "ここに氏名を入力" and "Tel" in f and "Mail" in f
    r = api.fill(src, {"Tel": "03-0000-0000", "ctl1": "変更", "Mail": True}, str(tmp_path / "o.doc"))
    assert r["proof_ok"] and r["read_back_ok"], r
    got = {x["id"]: x["value"] for x in api.fields(str(tmp_path / "o.doc"))}
    assert got["Tel"] == "03-0000-0000" and got["ctl1"] == "変更" and got["Mail"] is True


def _with_table_style(blob: bytes, left: int, right: int) -> bytes:
    import io
    import zipfile

    W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}>'
              '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:tblPr><w:tblCellMar>'
              f'<w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>')
    src = zipfile.ZipFile(io.BytesIO(blob))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            z.writestr(n, src.read(n))
        z.writestr("word/styles.xml", styles)
    return out.getvalue()


@pytest.mark.parametrize("styled", [False, True])
def test_wrap_estimate_reads_table_style_cell_margins(tmp_path, styled):
    """表のマスの余白：マス・表になければ、表のスタイル（既定の表のスタイル）の w:tblCellMar を読む（ECMA-376 17.7.6）。
    どこにも書かれていなければ 108 twip と仮に置き、見積もりの元にそう書く。"""
    from test_wrap_grid_circle import MSG, _docx_cell

    from todoku import api

    blob = _docx_cell("１　・　２　種", 2000)
    src = tmp_path / "t.docx"
    src.write_bytes(_with_table_style(blob, 0, 0) if styled else blob)
    r = api.fill(str(src), {"@1": {"old": "１　・　２　種", "new": "○１　・　２　種"}}, str(tmp_path / "o.docx"))
    hit = [w for w in r["warnings"] if MSG in w]
    assert hit, r["warnings"]
    if styled:
        assert "表のスタイル「TableNormal」" in hit[0] and "約 100 ポイント" in hit[0] and "仮に" not in hit[0]
    else:
        assert "仮に" in hit[0] and "約 89 ポイント" in hit[0]


def test_say_to_user_counts_unchosen_choices_and_question_places(tmp_path):
    """まだ空いている欄の伝え方：選んでいない選択肢の組と、残っている「？」の所も、数で伝える（一つずつは並べない）。"""
    from todoku import api

    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(p("氏名＿＿＿＿") + p("☐ 新規　☐ 変更") + p("☐ 同意する") + p("令和　　年　　月　　日")))
    r = api.fill(str(src), {"P1.c2": "山田"}, str(tmp_path / "o.docx"))
    assert r["unchosen_choice_groups"] == 2 and r["agent_places_left"] >= 3, r
    assert "選んでいない選択肢（チェックボックス・ラジオボタン）：2 組" in r["say_to_user"]
    assert "「？」の所" in r["say_to_user"]
    r = api.fill(str(src), {"P1.c2": "山田", "P2.c0": True}, str(tmp_path / "o2.docx"))
    assert r["unchosen_choice_groups"] == 1, r


# ---------------------------------------------------------------------- 表計算：所の id・離れた字・段階の符号
def _book(tmp_path, cells: dict) -> str:
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for a, v in cells.items():
        ws[a] = v
    path = tmp_path / "a.xlsx"
    wb.save(path)
    return str(path)


@pytest.mark.parametrize("text,vals,want", [("令和　　年　　月　　日", ["8", "4"], "令和8年4月"),
                                            ("Date:　　/　　/　　", ["2026", "11"], "Date:2026/11/")])
def test_sheet_span_ids_write_only_those_places(tmp_path, text, vals, want):
    from todoku import api

    src = _book(tmp_path, {"A1": "申請日", "B1": text})
    spans = [s["id"] for f in api.fields(src) for s in f.get("spans", [])]
    assert spans and all("@" in s for s in spans)
    r = api.fill(src, dict(zip(spans, vals)), str(tmp_path / "o.xlsx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert want in r["say_to_user"]


def test_sheet_far_text_is_not_presented_as_the_label():
    from todoku.api import far_label

    near = {"left": {"text": "氏名", "label": "氏名", "cell": "A12", "cols": 3, "mm": 60}}
    assert far_label({"id": "S!E12", "label": "氏名", "near": near}) == "E12（左に 3 列あけて『氏名』）"
    near = {"left": {"text": "氏名", "label": "氏名", "cell": "A12", "cols": 0, "mm": 0}}
    assert far_label({"id": "S!B12", "label": "氏名", "near": near}) is None  # すぐ隣のマスの字は、見出しのまま


def test_grade_reads_reason_codes_not_reason_text():
    from todoku.grade import sheet_input_reasons

    def grade_of(f):
        write: list = []
        sheet_input_reasons({"sheets": [{"fields": [f]}]}, [], write)
        return write[0].grade

    # 理由の字が何であっても（ほかの言葉に直しても）、符号で決まる
    assert grade_of({"id": "S!B1", "kind": "text", "reason": ["Data validation"], "reason_codes": ["validation"]}) == "A"
    assert grade_of({"id": "S!B1", "kind": "text", "reason": ["入力規則があるセル"], "reason_codes": ["formula_ref"]}) == "B"
