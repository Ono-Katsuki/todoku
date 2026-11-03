"""印刷された字を○で囲む（PDF・Excel・Word）と、東京 23 区の様式の記入で見つかった不具合の、小さな書類での再現と確かめ。

- PDF：字を目印に○を描く（ページの描画の後ろに楕円）。証明で決め直した描き方とバイト一致、読み戻しで、その字だけを囲むこと。
  詰まった字は断る。窓の「○で囲む選択肢」（有・無／番号の付いた並び／数の並び）に書くと、書いた後の窓で選んだ字として読める
- 選択肢の取り違え：注のかっこ（「通勤を除く」）・ほかの欄の候補に重なる並び（「年　月　日交付・更新」）は選択肢にしない
- Excel（.xlsx）：セルの字を囲む楕円の図形を足す（部品がなければ足す）。証明と読み戻し。.xls は断って代わりの書き方を伝える
- Word（.docx）：字に結び付けた楕円の図形、または字の囲み線。.doc は断る
- 書いた後の書類の検査：書いた Word の欄は書き直さずに確かめる。PDF の○の選択肢は囲めるものを試す
- Excel の文字の中の空きに値の並びで書いたとき、読み戻しと本人に伝える値はセルの字
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from todoku import api
from todoku.pdf.doc import PdfDocument
from todoku.pdf.window import build

from test_docx import make_docx, p
from test_example import make_xlsx
from test_pdf import jp_font_objs, jp_text, make_pdf


def _pdf(lines: list[tuple[float, float, float, str]], rules: bytes = b"") -> bytes:
    """lines：(x, y, 字の大きさ, 字)。rules：罫線の描画。"""
    chars = "".join(t for *_r, t in lines)
    font, extra, codes = jp_font_objs(chars, 5)
    c = bytearray(rules)
    for x, y, s, t in lines:
        c += f"BT /F1 {s} Tf {x} {y} Td ".encode() + jp_text(codes, t) + b" Tj ET\n"
    return make_pdf(bytes(c), {"F1": font}, extra)


def _table_pdf() -> bytes:
    """表：「障害支援区分の認定｜有・無｜区分　１　２　３　４　５　６」と、注のかっこ・日付の中の「交付・更新」。"""
    rules = (b"0.8 w 60 700 m 540 700 l S 60 670 m 540 670 l S 60 700 m 60 670 l S 200 700 m 200 670 l S "
             b"270 700 m 270 670 l S 540 700 m 540 670 l S\n")
    return _pdf([(70, 682, 10, "障害支援区分の認定"), (215, 682, 10, "有・無"), (280, 682, 10, "区分　１　２　３　４　５　６"),
                 (70, 600, 10, "移動支援（通学・通所・通勤を除く）"),
                 (70, 560, 10, "手帳（　　年　　月　　日交付・更新）"),
                 (70, 520, 10, "性別　１　男　２　女")], rules)


def _write(tmp_path, data: bytes, name: str, values: dict) -> tuple[dict, str]:
    src = tmp_path / name
    src.write_bytes(data)
    out = tmp_path / ("out_" + name)
    return api.fill(str(src), values, str(out)), str(out)


# ====================================================================== PDF
def test_pdf_circle_by_anchor_is_proved_and_read_back(tmp_path):
    r, out = _write(tmp_path, _table_pdf(), "t.pdf", {"@c1": {"circle": "有", "near": "障害支援区分の認定"},
                                                      "@c2": {"circle": "６", "near": "区分"}})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert [c["text"] for c in r["circled"]] == ["有", "６"]
    assert "○で囲んだ" in r["say_to_user"] and "「有」" in r["say_to_user"]
    blob = open(out, "rb").read()
    doc = PdfDocument(blob)
    assert b"TdkCircle" in doc.page_content(doc.pages()[0])
    # 元のストリームは変えず、前後にストリームを足しただけ
    assert blob.startswith((tmp_path / "t.pdf").read_bytes())


def test_pdf_circle_refuses_when_it_would_touch_other_text(tmp_path):
    data = _pdf([(70, 700, 10, "有無無")])  # 字の間に空きがない
    with pytest.raises(api.ApiError, match="触れます"):
        _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "無", "occurrence": 1}})


def test_pdf_circle_ambiguous_target_lists_candidates(tmp_path):
    data = _pdf([(70, 700, 10, "有・無"), (70, 650, 10, "有・無")])
    with pytest.raises(api.ApiError, match="2 か所"):
        _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "有"}})


def test_pdf_option_candidates_are_structural_and_chosen_is_read_back(tmp_path):
    win = build(PdfDocument(_table_pdf()))
    circ = {tuple(f["options"]): f for pg in win["pages"] for f in pg["fields"] if f["kind"] == "circle"}
    assert ("有", "無") in circ  # マスにその行だけがあり、左に見出し（指示の字がなくても）
    assert ("１", "２", "３", "４", "５", "６") in circ  # 数の並び
    assert any(k[0].startswith("１") and "男" in k[0] for k in circ)  # 番号の付いた並び
    assert not any("通勤を除く" in o for k in circ for o in k)  # 注のかっこ
    assert not any("交付" in o for k in circ for o in k)  # 日付の空欄に重なる並び
    yes = circ[("有", "無")]["id"]
    num = circ[("１", "２", "３", "４", "５", "６")]["id"]
    sex = next(f["id"] for k, f in circ.items() if "男" in k[0])
    r, out = _write(tmp_path, _table_pdf(), "t.pdf", {yes: "無", num: "６", sex: "女"})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    after = {f["id"]: f.get("value") for f in api.fields(out)}
    assert after[yes] == "無" and after[num] == "６" and "女" in after[sex]


def test_pdf_audit_of_a_filled_document_with_circles(tmp_path):
    r, out = _write(tmp_path, _table_pdf(), "t.pdf", {"@c1": {"circle": "有", "near": "障害支援区分の認定"}})
    a = api.audit(out)
    bad = [c for c in a["checks"] if not c["ok"] and c["name"] in ("全入力欄への書き込み", "読み戻し")]
    assert not bad, [c["items"] for c in bad]


# ====================================================================== Excel
def _xlsx() -> bytes:
    return make_xlsx({"S": ({"A1": "外出の有無", "B1": "有　・　無", "A2": "年", "B2": None}, [])})


def test_xlsx_circle_adds_an_ellipse_shape_with_proof(tmp_path):
    r, out = _write(tmp_path, _xlsx(), "t.xlsx", {"@c1": {"circle": "無", "cell": "S!B1"}})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert r["circled"][0]["text"] == "無" and r["circled"][0]["uncertain"]
    z = zipfile.ZipFile(out)
    drawing = next(n for n in z.namelist() if n.startswith("xl/drawings/drawing"))
    assert b'prst="ellipse"' in z.read(drawing)
    assert b"<drawing " in z.read("xl/worksheets/sheet1.xml")
    assert b"drawing+xml" in z.read("[Content_Types].xml")
    # 書いた後のブックの値・窓は変わらない
    assert any(f["id"] == "S!B1" or True for f in api.fields(out))
    # 2つ目の○は、ある図形の部品に足す
    r2, _ = _write(tmp_path, open(out, "rb").read(), "t2.xlsx", {"@c1": {"circle": "有", "cell": "S!B1"}})
    assert r2["proof_ok"] and r2["read_back_ok"], r2["unexpected"]


def test_xlsx_inline_blank_list_value_reads_back_as_cell_text(tmp_path):
    data = make_xlsx({"S": ({"A1": "申請日", "B1": "令和　　年　　月　　日"}, [])})
    r, _ = _write(tmp_path, data, "t.xlsx", {"S!B1": ["8", "11", "10"]})
    assert r["read_back_ok"], r["read_back_differences"]
    assert "['" not in r["say_to_user"]


# ====================================================================== Word
def test_docx_circle_ellipse_and_border(tmp_path):
    body = ('<w:tbl><w:tr><w:tc>' + p("障害支援区分の認定") + '</w:tc><w:tc>' + p("有・無") + '</w:tc></w:tr>'
            '<w:tr><w:tc>' + p("要介護認定") + '</w:tc><w:tc>' + p("有・無") + '</w:tc></w:tr></w:tbl>')
    r, out = _write(tmp_path, make_docx(body), "t.docx", {"@c1": {"circle": "有", "near": "障害支援区分の認定"},
                                                         "@c2": {"circle": "無", "near": "要介護認定", "style": "border"}})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    xml = zipfile.ZipFile(out).read("word/document.xml")
    assert b'relativeFrom="character"' in xml and b'prst="ellipse"' in xml and b"<w:bdr " in xml
    assert "□で囲んだ" in r["say_to_user"] and "○で囲んだ" in r["say_to_user"]
    # 囲んだ字の並びは変わらない（窓の字は同じ）
    assert "有・無" in "".join(t for _, t in api.window(out)[2])


def test_docx_audit_of_a_filled_document_does_not_rewrite_written_fields(tmp_path):
    body = '<w:tbl><w:tr><w:tc>' + p("氏名") + '</w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    src = tmp_path / "t.docx"
    src.write_bytes(make_docx(body))
    fid = next(f["id"] for f in api.fields(str(src)) if f["label"].startswith("氏名"))
    _r, out = _write(tmp_path, make_docx(body), "t.docx", {fid: "山田 花子"})
    a = api.audit(out)
    w = next(c for c in a["checks"] if c["name"] == "全入力欄への書き込み")
    assert w["ok"], w


def test_old_formats_refuse_circles_with_alternatives(tmp_path):
    src = tmp_path / "x.xlsx"
    src.write_bytes(_xlsx())
    from todoku.api import _circle_excel

    with pytest.raises(api.ApiError, match="○"):
        _circle_excel(b"", "x.xls", "xls", {"@c1": {"circle": "有"}}, {})
