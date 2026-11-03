"""世界調査の監査で見つかった、段階・欄の誤り（中身を見ない段階・書けない欄・確かめられない規則）。"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku import api
from todoku.book import kind_of
from todoku.grade import grade_blob

from test_docx import make_docx, p
from test_pdf import jp_font_objs, jp_text, make_pdf

openpyxl = pytest.importorskip("openpyxl")


def test_zip_named_csv_is_not_graded_by_its_name():
    # CSV をまとめた ZIP（統計局の配布など）：名前が .csv でも、名前でなく中のファイルで段階を決める
    def zip_of(files):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            for n, b in files.items():
                z.writestr(n, b)
        return buf.getvalue()

    blob = zip_of({"data.csv": "a,b\n1,2\n", "data_MetaData.csv": "x\n"})
    assert kind_of(blob, "data.csv") == "zip"
    g = grade_blob(blob, "data.csv")
    assert g.grade == "S" and all(r.title.startswith(("data.csv：", "data_MetaData.csv：")) for r in g.reasons), g.headline()
    g = grade_blob(zip_of({"x.bin": b"\x00\x01"}), "data.csv")
    assert g.grade == "F", g.headline()


def _blank_pdf(size: int) -> bytes:
    chars = "令和年月日"
    font, extra, codes = jp_font_objs(chars, 5)
    s = b"<" + b"".join(f"{codes[ch]:04X}".encode() if ch in codes else b"" for ch in "令和") + b">"
    c = (b"BT /F1 " + str(size).encode() + b" Tf 100 700 Td " + s + b" Tj " + str(4 * size).encode() + b" 0 Td "
         + jp_text(codes, "年") + b" Tj " + str(3 * size).encode() + b" 0 Td " + jp_text(codes, "月") + b" Tj "
         + str(3 * size).encode() + b" 0 Td " + jp_text(codes, "日") + b" Tj ET\n")
    return make_pdf(c, {"F1": font}, extra)


def test_blanks_in_tiny_text_are_shown_but_not_writable(tmp_path):
    # 書き込みと同じ決め方で1字も入らない空欄（1pt の字の図の注記など）も、消さずに「？」の所として示す。
    # 書けない理由（小さい）を添えて読み取り専用にし、書こうとすれば「入りきりません」と断る。監査は通る
    big, tiny = tmp_path / "big.pdf", tmp_path / "tiny.pdf"
    big.write_bytes(_blank_pdf(10))
    tiny.write_bytes(_blank_pdf(1))
    assert any(not f.get("readonly") for f in api.fields(str(big)))
    small = api.fields(str(tiny))
    assert small and all(f.get("agent") and f.get("readonly") and "小さい" in f.get("why_not", "") for f in small), small
    with pytest.raises(api.ApiError, match="入りきりません"):
        api.fill(str(tiny), {small[0]["id"]: "1"}, str(tmp_path / "out.pdf"))
    assert api.audit(str(tiny))["ok"]


def test_word_and_pdf_agree_on_structural_blanks(tmp_path):
    # かっこの組の中が空白だけの所（「（　　　）」）は、PDF と同じく作りで決まる空欄（規則の欄）。年月日の空欄は語からの推定（？）
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(p("電話（　　　　）") + p("令和　　年　　月　　日")))
    fs = api.fields(str(src))
    tel = [f for f in fs if f["label"].startswith("電話")]
    date = [f for f in fs if "令和" in f["label"]]
    assert tel and not any(f.get("agent") for f in tel), fs
    assert date and all(f.get("agent") for f in date), fs


def test_append_after_a_footnote_mark_uses_the_body_format(tmp_path):
    # 「人口※1」の後ろ（段落の終わりが脚注の印）：書けて、書いた字は印の上付きでなく本文の書式
    body = (p("届出書") + '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>人口</w:t></w:r><w:r><w:rPr><w:vertAlign w:val="superscript"/>'
            '</w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>')
    src, out = tmp_path / "a.docx", tmp_path / "b.docx"
    src.write_bytes(make_docx(body))
    fid = next(f["id"] for f in api.fields(str(src)) if f["label"].startswith("人口"))
    api.fill(str(src), {fid: "1200"}, str(out))
    xml = zipfile.ZipFile(out).read("word/document.xml").decode()
    run = xml[xml.rindex("<w:r>", 0, xml.index("1200")):xml.index("1200")]
    assert "<w:b/>" in run and "superscript" not in run
    assert api.audit(str(src))["ok"]


def _rule_book(formula: str) -> bytes:
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = "住所"
    dv = DataValidation(type="custom", formula1=formula, showErrorMessage=True)
    dv.add("B1")
    ws.add_data_validation(dv)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_dbcs_rule_is_checked_like_jis(tmp_path):
    # DBCS は JIS の英語版の名前：全角の字だけを受け付ける規則を、Excel と同じく確かめる
    src = tmp_path / "a.xlsx"
    src.write_bytes(_rule_book("B1=DBCS(B1)"))
    api.fill(str(src), {"S!B1": "大阪府１－１"}, str(tmp_path / "ok.xlsx"))
    with pytest.raises(api.ApiError):
        api.fill(str(src), {"S!B1": "ABC"}, str(tmp_path / "ng.xlsx"))
    assert "確かめられない入力規則" not in {r.title for r in grade_blob(src.read_bytes(), "a.xlsx").write_reasons}


def test_rule_with_unknown_function_lowers_write_grade():
    # 計算できない関数を使う入力規則：書いた値が合うかを確かめられないので、書く段階は A にしない
    g = grade_blob(_rule_book("NOSUCHFUNC(B1)"), "a.xlsx")
    assert g.write != "A" and any(r.title == "確かめられない入力規則" for r in g.write_reasons), g.headline()
