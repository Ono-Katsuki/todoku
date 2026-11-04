"""第 7 回の調べの残り：述べた意味（as）と近くの字の言い分け・字の前後への差し込み（before・after）・Word の○の既定・壊れた複合ファイルの言葉。
様式はここで作る（どの言語でも同じ）。"""

from __future__ import annotations

import re
import struct
import zipfile

import openpyxl
import pytest
from test_circle import _pdf as circle_pdf
from test_docx import make_docx, p
from test_pdf_near_widget import _GAP_ROWS, _GAP_WIDGETS, _form_pdf
from test_xls_controls import make_book

from todoku import api


def _xlsx(tmp_path, cells: dict, name="a.xlsx") -> str:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for k, v in cells.items():
        ws[k] = v
    path = tmp_path / name
    wb.save(path)
    return str(path)


def _docx(tmp_path, body: str, name="a.docx") -> str:
    path = tmp_path / name
    path.write_bytes(make_docx(body))
    return str(path)


def _cell(path: str, cell: str):
    return openpyxl.load_workbook(path)["S"][cell].value


# ====================================================================== 1. as と近くの字
@pytest.mark.parametrize("label,value,meaning", [("住所", "昭和60年4月1日", "生年月日"), ("Address", "1985-04-01", "Date of birth"),
                                                 ("주소", "1985년 4월 1일", "생년월일"), ("العنوان", "١٩٨٥ م", "تاريخ الميلاد")])
def test_as_is_shown_and_near_text_is_a_fact(tmp_path, label, value, meaning):
    src = _xlsx(tmp_path, {"A1": label})
    out = str(tmp_path / "o.xlsx")
    r_ = api.fill(src, {"S!B1": {"value": value, "as": meaning}}, out)
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    assert _cell(out, "B1") == value  # as は書類に書かない
    assert f"{meaning}：{value}（B1／近くの字「{label}」）" in r_["say_to_user"], r_["say_to_user"]
    assert r_["written_without_as"] == [] and r_["written"][0]["as"] == meaning


@pytest.mark.parametrize("label", ["住所", "Address", "주소"])
def test_without_as_the_near_text_is_not_the_field_name(tmp_path, label):
    src = _xlsx(tmp_path, {"A1": label})
    r_ = api.fill(src, {"S!B1": "昭和60年"}, str(tmp_path / "o.xlsx"))
    say = r_["say_to_user"]
    assert f"B1（近くの字「{label}」）：昭和60年" in say, say
    assert f"- {label}：" not in say
    assert r_["written_without_as"] == ["S!B1"]


def test_word_place_with_and_without_as(tmp_path):
    src = _docx(tmp_path, p("氏名：　　　") + p("住所：　　　"))
    ids = [f["id"] for f in api.fields(src)]
    r_ = api.fill(src, {ids[0]: {"value": "山田", "as": "申請者の氏名"}, ids[1]: "東京都"}, str(tmp_path / "o.docx"))
    say = r_["say_to_user"]
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    assert f"申請者の氏名：山田（{ids[0]}／近くの字「" in say, say
    assert f"{ids[1]}：住所：［東京都］" in say, say  # 行の字のとおりに、空欄に値を入れた形
    assert r_["written_without_as"] == [ids[1]]


def test_blank_in_a_line_is_shown_in_its_line(tmp_path):
    """行の字に値を入れた形は、見出しではなく行の字そのもの：as があれば意味を前に出す。"""
    src = _xlsx(tmp_path, {"A1": "令和　　年　　月"})
    spans = [s["id"] for f in api.fields(src) for s in f.get("spans", [])]
    r_ = api.fill(src, {spans[0]: {"value": "8", "as": "申請年"}, spans[1]: {"value": "4", "as": "申請月"}}, str(tmp_path / "o.xlsx"))
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    assert "申請年／申請月：令和8年4月" in r_["say_to_user"], r_["say_to_user"]
    assert not r_["written_without_as"]


def test_remaining_places_use_the_facts_wording(tmp_path):
    body = ("<w:tbl><w:tr><w:tc>" + p("住所") + "</w:tc><w:tc><w:p/></w:tc></w:tr>"
            "<w:tr><w:tc>" + p("電話") + "</w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>")
    src = _docx(tmp_path, body)
    r_ = api.fill(src, {"T1.R1.C2": {"value": "東京都", "as": "現住所"}}, str(tmp_path / "o.docx"))
    line = next(x for x in r_["say_to_user"].splitlines() if x.startswith("続きの空欄"))
    assert "T1.R2.C2（近くの字「電話」）" in line, line
    assert "現住所：東京都（T1.R1.C2／近くの字「住所」）" in r_["say_to_user"]
    assert r_["still_empty"][0]["id"] == "T1.R2.C2"


def test_csv_column_name_is_the_values_name(tmp_path):
    src = tmp_path / "a.csv"
    src.write_bytes("氏名,年齢\n山田,\n".encode())
    r_ = api.fill(str(src), {"R2.C2": {"value": "30", "as": "申請者の年齢"}}, str(tmp_path / "o.csv"))
    assert r_["proof_ok"] and "申請者の年齢：30（R2.C2／欄の名前「年齢（2行目）」）" in r_["say_to_user"], r_["say_to_user"]
    r2 = api.fill(str(src), {"R2.C2": "30"}, str(tmp_path / "o2.csv"))
    assert "年齢（2行目）：30" in r2["say_to_user"]


def test_as_must_be_text_and_value_only_dict_is_the_value(tmp_path):
    src = _xlsx(tmp_path, {"A1": "氏名"})
    with pytest.raises(api.ApiError, match="as は"):
        api.fill(src, {"S!B1": {"value": "x", "as": ""}}, str(tmp_path / "o.xlsx"))
    r_ = api.fill(src, {"S!B1": {"value": "x"}}, str(tmp_path / "o.xlsx"))
    assert r_["proof_ok"] and _cell(str(tmp_path / "o.xlsx"), "B1") == "x"


def test_try_calculation_takes_the_same_value_form(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"], ws["A2"], ws["B2"] = "単価", "合計", "=B1*2"
    path = tmp_path / "c.xlsx"
    wb.save(path)
    got = api.calculate(str(path), {"S!B1": {"value": 5, "as": "単価"}})
    assert any(c["cell"] == "S!B2" and str(c["after"]).strip() == "10" for c in got["changed"]), got


def test_circle_with_as_and_without(tmp_path):
    src = _docx(tmp_path, p("性別　男・女"))
    r_ = api.fill(src, {"@c1": {"circle": "女", "as": "申請者の性別"}}, str(tmp_path / "o.docx"))
    assert "申請者の性別：「女」" in r_["say_to_user"], r_["say_to_user"]
    assert r_["written_without_as"] == [] and r_["circled"][0]["as"] == "申請者の性別"
    r2 = api.fill(src, {"@c1": {"circle": "女"}}, str(tmp_path / "o2.docx"))
    assert r2["written_without_as"] == [r2["circled"][0]["id"]]


# ====================================================================== 2. 字の前・後ろへの差し込み
@pytest.mark.parametrize("text,mark,occ,value,want", [("年　月　日", "年", None, "令和8", "令和8年　月　日"),
                                                      ("年　月　日", "月", None, "11", "年　11月　日"),
                                                      ("Year　Month", "Year", None, "2026", "2026Year　Month"),
                                                      ("년 월 일", "월", None, "11", "년 11월 일"),
                                                      ("年度　年", "年", 2, "8", "年度　8年")])
def test_word_before(tmp_path, text, mark, occ, value, want):
    from todoku.docx.document import DocxDocument, iter_paras

    src = _docx(tmp_path, p(text))
    spec = {"before": mark, "value": value, **({"occurrence": occ} if occ else {}), "as": "日付"}
    out = str(tmp_path / "o.docx")
    r_ = api.fill(src, {"@1": spec}, out)
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    doc = DocxDocument(open(out, "rb").read(), "o.docx")
    texts = [q.text for _k, _p, blocks in doc.all_blocks() for q in iter_paras(blocks)]
    assert texts[0] == want, texts
    assert f"日付：{value}（{r_['anchored']['@1']}" in r_["say_to_user"], r_["say_to_user"]


def test_word_after_and_several_in_one_paragraph(tmp_path):
    from todoku.docx.document import DocxDocument, iter_paras

    src = _docx(tmp_path, p("年　月　日　氏名"))
    out = str(tmp_path / "o.docx")
    r_ = api.fill(src, {"@1": {"before": "年", "value": 8}, "@2": {"before": "月", "value": "4"},
                        "@3": {"before": "日", "value": "1"}, "@4": {"after": "氏名", "value": "　山田"}}, out)
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    doc = DocxDocument(open(out, "rb").read(), "o.docx")
    texts = [q.text for _k, _p, blocks in doc.all_blocks() for q in iter_paras(blocks)]
    assert texts[0] == "8年　4月　1日　氏名　山田", texts


@pytest.mark.parametrize("name", ["a.xlsx"])
def test_excel_before_and_after(tmp_path, name):
    src = _xlsx(tmp_path, {"A1": "年　月　日", "A2": "氏名", "A3": "年度"}, name)
    out = str(tmp_path / ("o" + name))
    r_ = api.fill(src, {"@1": {"before": "年", "cell": "S!A1", "value": "令和8", "as": "申請年"},
                        "@2": {"before": "月", "cell": "S!A1", "value": 4},
                        "@3": {"after": "氏名", "value": "　山田"}}, out)
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    assert _cell(out, "A1") == "令和8年　4月　日" and _cell(out, "A2") == "氏名　山田" and _cell(out, "A3") == "年度"
    assert "申請年：令和8年　4月　日（A1" in r_["say_to_user"], r_["say_to_user"]


def test_excel_before_occurrence_and_choice_is_listed(tmp_path):
    src = _xlsx(tmp_path, {"A1": "年　年", "A2": "年"})
    with pytest.raises(api.ApiError, match="3 か所"):
        api.fill(src, {"@1": {"before": "年", "value": "8"}}, str(tmp_path / "o.xlsx"))
    out = str(tmp_path / "o.xlsx")
    r_ = api.fill(src, {"@1": {"before": "年", "occurrence": 2, "value": "8"}}, out)
    assert r_["proof_ok"] and _cell(out, "A1") == "年　8年", r_


def test_xls_before(tmp_path):
    from todoku.book import open_book

    blob, _objs = make_book()
    src = tmp_path / "a.xls"
    src.write_bytes(blob)
    out = str(tmp_path / "o.xls")
    r_ = api.fill(str(src), {"@1": {"before": "内容", "value": "変更の", "as": "申請の種類"}}, out)
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    wb = open_book(open(out, "rb").read(), "o.xls")
    assert wb.sheets[0].cells[(5, 2)].value == "お申込みの変更の内容"
    assert "申請の種類：お申込みの変更の内容" in r_["say_to_user"], r_["say_to_user"]


def test_pdf_as_is_meaning_kind_is_kept_and_before_is_the_left(tmp_path):
    src = tmp_path / "a.pdf"
    src.write_bytes(_form_pdf(_GAP_ROWS, _GAP_WIDGETS))
    r_ = api.fill(str(src), {"name": {"value": "山田", "as": "申請者の氏名"}, "tel": "03"}, str(tmp_path / "o.pdf"))
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    assert "申請者の氏名：山田（name／近くの字「氏名」）" in r_["say_to_user"], r_["say_to_user"]
    assert "tel（近くの字「" in r_["say_to_user"] and r_["written_without_as"] == ["tel"], r_["say_to_user"]
    pl = tmp_path / "b.pdf"
    pl.write_bytes(circle_pdf([(72, 700, 10, "A"), (300, 700, 10, "RIGHT")]))
    got = api.fill(str(pl), {"@1": {"before": "RIGHT", "value": "L", "as": "印"}}, str(tmp_path / "o2.pdf"))
    assert got["proof_ok"] and got["read_back_ok"], got
    x = got["written_rects"]["@1"]
    assert x["x"] + x["w"] <= 300 / 72 * 25.4 + 0.5, x
    assert "印：L（" in got["say_to_user"]


def test_as_equal_to_a_kind_candidate_chooses_the_kind_and_kind_says_both():
    from todoku.api import _split_values

    cands = {"p1.m1-1": ["comb", "text"]}
    v, said = _split_values({"p1.m1-1": {"as": "text", "value": "100"}}, "pdf", cands)
    assert v == {"p1.m1-1": {"as": "text", "value": "100"}} and not said
    v, said = _split_values({"p1.m1-1": {"kind": "text", "as": "郵便番号", "value": "100"}}, "pdf", cands)
    assert v == {"p1.m1-1": {"as": "text", "value": "100"}} and said == {"p1.m1-1": "郵便番号"}
    v, said = _split_values({"p1.t1-1": {"as": "text", "value": "100"}}, "pdf", cands)  # 種類の候補のない欄の as は意味
    assert v == {"p1.t1-1": "100"} and said == {"p1.t1-1": "text"}
    v, said = _split_values({"S!B1": {"as": "text", "value": "100"}}, "xlsx")
    assert v == {"S!B1": "100"} and said == {"S!B1": "text"}


@pytest.mark.parametrize("spec,msg", [({"before": "年", "after": "月", "value": "1"}, "どちらか"),
                                      ({"before": "年"}, "value"), ({"before": "", "value": "1"}, "目印"),
                                      ({"before": "年", "value": True}, "value"), ({"before": "年", "value": "1", "page": 1}, "分からない指定")])
def test_insert_spec_refusals(tmp_path, spec, msg):
    src = _docx(tmp_path, p("年　月"))
    with pytest.raises(api.ApiError, match=msg):
        api.fill(src, {"@1": spec}, str(tmp_path / "o.docx"))


def test_no_extra_place_for_evenly_spaced_text(tmp_path):
    """字間をそろえた字の並びに、先頭の空きを作らない（目印の字で書く）。"""
    src = _docx(tmp_path, p("年　月　日"))
    assert not [f for f in api.fields(src) if re.fullmatch(r"P\d+\.c0(\.\d+)?", f["id"])]


def test_other_text_formats_take_the_same_insert(tmp_path):
    src = tmp_path / "a.txt"
    src.write_text("年　月\n", encoding="utf-8")
    out = tmp_path / "o.txt"
    r_ = api.fill(str(src), {"@1": {"before": "月", "value": "11", "as": "申請月"}}, str(out))
    assert r_["proof_ok"] and r_["read_back_ok"], r_
    assert out.read_text(encoding="utf-8") == "年　11月\n"
    assert "申請月：11" in r_["say_to_user"], r_["say_to_user"]


@pytest.mark.parametrize("name", ["a.odt", "a.rtf"])
def test_odt_and_rtf_take_the_same_insert(tmp_path, name):
    from test_openformats import ODT_AUTO, ODT_BODY, ODT_FORMS, RTF, make_odt

    src = tmp_path / name
    src.write_bytes(make_odt(ODT_BODY, ODT_AUTO, ODT_FORMS) if name == "a.odt" else RTF)
    r_ = api.fill(str(src), {"@1": {"after": "申請書", "value": "（新規）", "as": "申請の区分"}}, str(tmp_path / ("o" + name)))
    assert r_["proof_ok"] and r_["read_back_ok"] and r_["anchored"] == {"@1": "P1.w1"}, r_
    assert "申請の区分：（新規）（P1.w1" in r_["say_to_user"], r_["say_to_user"]


# ====================================================================== 3. Word の○の既定は字の囲み線
def test_word_circle_default_is_the_run_border_and_ellipse_is_explicit(tmp_path):
    src = _docx(tmp_path, p("有・無"))
    out = str(tmp_path / "o.docx")
    r_ = api.fill(src, {"@c1": {"circle": "有"}}, out)
    c = r_["circled"][0]
    assert r_["proof_ok"] and r_["read_back_ok"] and c["style"] == "border" and "w:bdr" in c["how"], r_
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    assert "w:bdr" in xml and "wp:anchor" not in xml
    assert "□で囲んだ" in r_["say_to_user"] and "○ではなく" in r_["say_to_user"]
    assert not any("見積もった" in w for w in r_.get("warnings") or [])  # 位置の見積もりの注意は、楕円のときだけ

    out2 = str(tmp_path / "o2.docx")
    r2 = api.fill(src, {"@c1": {"circle": "有", "style": "ellipse"}}, out2)
    assert r2["proof_ok"] and r2["read_back_ok"] and r2["circled"][0]["style"] == "ellipse" and r2["circled"][0]["uncertain"], r2
    assert "wp:anchor" in zipfile.ZipFile(out2).read("word/document.xml").decode()
    assert any("見積もった" in w for w in r2["warnings"]) and "○で囲んだ" in r2["say_to_user"]


# ====================================================================== 4. 壊れた複合ファイル
def _cfb_like() -> bytes:
    """ヘッダーの形だけ合って、中の記録の長さ・位置が合わない複合ファイル。"""
    head = bytearray(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(504))
    struct.pack_into("<HHHHH", head, 0x18, 0x3e, 3, 0xfffe, 9, 6)
    struct.pack_into("<I", head, 0x2c, 200)  # FAT のセクターが 200（ファイルにない）
    struct.pack_into("<I", head, 0x30, 100000)
    for i in range(109):
        struct.pack_into("<I", head, 0x4c + 4 * i, 100000 + i)
    return bytes(head)


@pytest.mark.parametrize("name", ["a.doc", "a.xls"])
@pytest.mark.parametrize("cut", [None, 100, 512, 700])
def test_broken_cfb_never_shows_the_internal_error(tmp_path, name, cut):
    blob = _cfb_like() if cut is None else (_cfb_like() + bytes(4096))[:cut]
    src = tmp_path / name
    src.write_bytes(blob)
    calls = [lambda: api.window(str(src)), lambda: api.fields(str(src)), lambda: api.audit(str(src)), lambda: api.grade(str(src)),
             lambda: api.summary(str(src)), lambda: api.fill(str(src), {"A1": "x"}, str(tmp_path / ("o" + name))),
             lambda: api.calc_model(str(src)), lambda: api.calculate(str(src), {"A1": 1})]
    for call in calls:
        try:
            got = call()
        except api.ApiError as e:
            got = str(e)
        text = str(got)
        assert not re.search(r"unpack|requires a buffer|struct\.error|bytes object", text), text[:400]

