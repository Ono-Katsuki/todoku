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
