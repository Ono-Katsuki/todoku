"""印刷された字を○で囲む（PDF・Excel・Word）と、東京 23 区の様式の記入で見つかった不具合の、小さな書類での再現と確かめ。

- PDF：字を目印に○を描く（ページの描画の後ろに楕円）。証明で決め直した描き方とバイト一致、読み戻しで、その字だけを囲むこと。
  窓は印刷された選択肢（有・無／番号の付いた並び／数の並び／元号）を欄として並べない（字の意味からは決めない）：エージェントが
  字を指定して circle で書き、書いた後の窓は、囲んだ字の後ろに「［p1.c1：○］」と出す。字の間が詰まった所では、すき間を 0 まで・
  縦横の比を細かく変えて、ほかの字の面に触れない楕円を探し、どれも触れるときだけ断って、四角の囲み線（style "border"）を勧める
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
    # 上下左右に字が詰まっている（行の間も字の間も空きがない）：どの楕円もほかの字の面に触れるので断り、四角の囲み線を勧める
    data = _pdf([(70, 710, 10, "有無無"), (70, 700, 10, "有無無"), (70, 690, 10, "有無無")])
    with pytest.raises(api.ApiError, match="触れます.*border"):
        _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "無", "occurrence": 3}})
    # 四角の囲み線は、字の箱に沿うので、ほかの字の面（箱の縁から 1/12 字の内側）に触れない
    r, out = _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "無", "occurrence": 3, "style": "border"}})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    win = build(PdfDocument(open(out, "rb").read()))
    assert any("無［p1.c1：○］" in ln for ln in win["pages"][0]["reading"]), win["pages"][0]["reading"]


def test_pdf_circle_fits_between_tight_neighbours(tmp_path):
    """23 区の様式（「補装具費（購入・借受け・修理）」）：囲む字の箱に、かっこ「（」の箱が重なる。すき間を 0 まで小さくし、
    縦長の楕円で、ほかの字の面に触れずに囲む。読み戻しも、書くときと同じ決め方（字の箱を囲み、ほかの字の面に触れない）で通る。"""
    data = _pdf([(70, 700, 10, "補装具費（購入・借受け・修理）支給申請書")])
    r, out = _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "購入"}})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    e = r["circled"][0]["ellipse_pt"]
    assert e[3] > e[2] / 2  # 縦に長い楕円（左右は字の箱にほぼ沿う）
    r, out = _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "購入", "pad": 0}})
    assert r["read_back_ok"], r["read_back_differences"]


def test_pdf_circle_ambiguous_target_lists_candidates(tmp_path):
    data = _pdf([(70, 700, 10, "有・無"), (70, 650, 10, "有・無")])
    with pytest.raises(api.ApiError, match="2 か所"):
        _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "有"}})


def test_pdf_printed_choices_stay_text_and_circled_text_is_read_back(tmp_path):
    # 印刷された選択肢（有・無／数の並び／番号の付いた並び／注のかっこ）は、欄として並べない：字のまま窓に出す。
    # 以前は字の並び（「・」の区切り・続き番号）から○で囲む選択肢の欄を作っていたが、それは字の意味の推定なので、エージェントが決める
    win = build(PdfDocument(_table_pdf()))
    assert not [f for pg in win["pages"] for f in pg["fields"] if f["kind"] == "circle"]
    text = "\n".join(win["pages"][0]["reading"])
    # 数の間の 1 字ぶん（1 em）の空きは「？」の所として示す（書く所かはエージェントが決める）。字はそのまま
    assert "有・無" in text and "１　２　３" in re.sub(r"［[^］]*？］", "　", text) and "通勤を除く" in text
    assert "印刷された選択肢は欄として並べていません" in "".join(t for _h, t in api.window(_save(tmp_path, _table_pdf()))[2])
    r, out = _write(tmp_path, _table_pdf(), "t.pdf", {"@c1": {"circle": "無", "near": "障害支援区分の認定"},
                                                      "@c2": {"circle": "６", "near": "区分"},
                                                      "@c3": {"circle": "２", "near": "性別"}})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    win = build(PdfDocument(open(out, "rb").read()))
    text = "\n".join(win["pages"][0]["reading"])
    assert "無［p1.c1：○］" in text and "６［p1.c2：○］" in text and "２［p1.c3：○］" in text, text
    assert {c["text"] for c in win["pages"][0]["circled"]} == {"無", "６", "２"}


def _save(tmp_path, data: bytes, name: str = "w.pdf") -> str:
    path = tmp_path / name
    path.write_bytes(data)
    return str(path)


def test_pdf_audit_of_a_filled_document_with_circles(tmp_path):
    # 書いた後の書類の検査：書いてある○は試しに書き直さず、いまの○を、書くときと同じ決め方で確かめる
    r, out = _write(tmp_path, _table_pdf(), "t.pdf", {"@c1": {"circle": "有", "near": "障害支援区分の認定"}})
    a = api.audit(out)
    bad = [c for c in a["checks"] if not c["ok"] and c["name"] in ("全入力欄への書き込み", "読み戻し", "書いてある○")]
    assert not bad, [c["items"] for c in bad]
    assert any(c["name"] == "書いてある○" and c["ok"] for c in a["checks"])


def test_pdf_era_circle_is_read_back_when_the_line_has_larger_chars(tmp_path):
    """元号の字が、同じ行のほかの字（「年」「月」「日」）より小さい：行の高さでなく字の箱で、囲んだ○を読む
    （目黒区・江戸川区の様式で、書いた「昭和」が窓で読めなかった）。"""
    rules = (b"0.8 w 60 700 m 540 700 l S 60 670 m 540 670 l S 60 700 m 60 670 l S 200 700 m 200 670 l S "
             b"540 700 m 540 670 l S\n")
    data = _pdf([(70, 682, 10, "生年月日"), (210, 682, 7.92, "昭和・平成"), (260, 682, 12, "　　年　　月　　日")], rules)
    r, out = _write(tmp_path, data, "t.pdf", {"@c1": {"circle": "昭和", "near": "生年月日"}})
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    win = build(PdfDocument(open(out, "rb").read()))
    assert [c["text"] for c in win["pages"][0]["circled"]] == ["昭和"]


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
    assert isinstance(api.fields(out), list)  # 書いた後も窓を作れる（「有　・　無」の 1 つの空白は所にしないので、欄の数は問わない）
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
