"""第6回の見直しで見つかった、証明・読み戻しを通ってしまう誤った書き出しと、誤った断りの確かめ（合成の文書）。

- Excel：数式セルの計算結果を差し替えるとき、属性のある <v xml:space="preserve"> も取り除き、<v> を一つにする。
  <v>・<f>・<is> が二つ以上あるセル（CT_Cell に合わない）は、証明で断り、読み戻しでも合うとしない
"""

from __future__ import annotations

import io
import re
import zipfile

import pytest

from todoku import api
from test_circle import _pdf, _write
from test_docx import make_docx
from test_docx import p as wp
from test_round5_office import _xlsx


# ---------------------------------------------------------------- Excel：<v> の二重化
def _vcount(blob: bytes, ref: str) -> int:
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        s = z.read("xl/worksheets/sheet1.xml").decode()
    m = re.search(r'<c r="' + ref + r'"[^>]*>(.*?)</c>', s, re.S)
    return len(re.findall(r"<v[ >]", m[1]))


def test_formula_value_with_space_attr_is_replaced_not_duplicated(tmp_path):
    rows = ('<row r="1"><c r="A1"/><c r="B1" t="str"><f>"区 "&amp;A1&amp;" 号 "</f>'
            '<v xml:space="preserve">区  号 </v></c></row>')
    src = tmp_path / "a.xlsx"
    src.write_bytes(_xlsx(rows))
    out = tmp_path / "o.xlsx"
    r = api.fill(str(src), {"S!A1": "業平"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    blob = out.read_bytes()
    assert _vcount(blob, "B1") == 1
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        s = z.read("xl/worksheets/sheet1.xml").decode()
    assert '<v xml:space="preserve">区 業平 号 </v>' in s


def test_proof_rejects_cell_with_two_values():
    from todoku.xlsx.proof import prove

    a = _xlsx('<row r="1"><c r="A1"><v>1</v></c></row>')
    b = _xlsx('<row r="1"><c r="A1"><v>2</v><v>1</v></c></row>')
    pr = prove(a, b, allowed_inputs={("S", "A1")})
    assert not pr.ok
    assert any("CT_Cell" in u for u in pr.unexpected)


def test_readback_rejects_cell_with_two_values():
    from todoku.xlsx.readback import check
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(_xlsx('<row r="1"><c r="A1"><v>2</v><v>1</v></c></row>'))
    assert wb.sheet("S").cells[(1, 1)].schema_error
    assert check(wb, {"S!A1": 2})["S!A1"][0] is False


# ---------------------------------------------------------------- Excel：幅に入らない数は「####」
def test_number_in_narrow_wrap_cell_predicts_hashes_not_wrap(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Alignment

    wb_ = openpyxl.Workbook()
    ws = wb_.active
    ws.title = "S"
    ws["A1"] = "年齢"
    ws["B1"].alignment = Alignment(wrap_text=True)
    ws.column_dimensions["B"].width = 1.2
    p = tmp_path / "n.xlsx"
    wb_.save(p)
    r = api.fill(str(p), {"S!B1": 30}, str(tmp_path / "o.xlsx"))
    assert any("####" in w and "S!B1" in w for w in r["warnings"]), r["warnings"]
    assert not any("折り返す見込み" in w for w in r["warnings"])
    # 「標準」の形式の小数は、桁を減らして入るので「####」にならない
    ws.column_dimensions["B"].width = 3
    wb_.save(p)
    r = api.fill(str(p), {"S!B1": 3.14159}, str(tmp_path / "o2.xlsx"))
    assert not any("####" in w for w in r["warnings"]), r["warnings"]


# ---------------------------------------------------------------- PDF：after の見出しのマスに読める大きさで入らない
def test_after_label_moves_to_the_value_cell_when_the_heading_cell_is_too_small(tmp_path):
    rules = (b"0.8 w 60 700 m 400 700 l S 60 680 m 400 680 l S 60 700 m 60 680 l S 140 700 m 140 680 l S "
             b"400 700 m 400 680 l S\n")
    data = _pdf([(70, 686, 10, "氏　　名"), (340, 686, 10, "番号")], rules)
    r, _ = _write(tmp_path, data, "t.pdf", {"@1": {"after": "氏　　名", "value": "山田　花子", "page": 1}})
    assert r["proof_ok"] and r["read_back_ok"]
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.anchor import resolve_write

    sp = resolve_write(PdfDocument(data), {"after": "氏　　名", "value": "山田　花子", "page": 1}, {}, {})
    x0, _y0, x1, _y1 = sp["rect"]
    assert x0 >= 140 and x1 <= 340, sp["rect"]  # 右のマスの、「番号」の手前
    assert not any("読みにくい" in w for w in r["warnings"]), r["warnings"]


# ---------------------------------------------------------------- PDF：後ろの字にくっつけるのは単位のような字だけ
def test_right_alignment_only_glues_to_unit_like_text():
    from todoku.pdf.fill import inline_align, unit_like

    assert unit_like("年") and unit_like("月　日") and unit_like("級") and unit_like("号")
    assert not unit_like("信用金庫") and not unit_like("（続柄）") and not unit_like("")
    assert inline_align({"label": "［ ］年"}) == "right"
    assert inline_align({"label": "［ ］信用金庫 支店"}) == "right_gap"
    assert inline_align({"label": "氏 名［ ］（続柄）", "rect": [0, 0, 200, 12], "size": 10}, 40) == "right_gap"


def test_right_gap_keeps_one_em_before_the_following_text(tmp_path):
    rules = b"0.8 w 60 700 m 400 700 l S 60 680 m 400 680 l S\n"
    data = _pdf([(300, 686, 10, "信用金庫")], rules)
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.window import build

    win = build(PdfDocument(data))
    spot = next(f for pg in win["pages"] for f in pg["fields"] if "信用金庫" in (f.get("label") or ""))
    r, out = _write(tmp_path, data, "t.pdf", {spot["id"]: "架空"})
    assert r["proof_ok"] and r["read_back_ok"]
    doc = PdfDocument(open(out, "rb").read())
    blob = open(out, "rb").read()
    m = re.findall(rb"1 0 0 1 ([\d.]+) ([\d.]+) Tm", blob)
    assert m
    x = float(m[-1][0])
    w = spot["rect"][2] - spot["rect"][0]
    # 字の右の端（x + 2 字）と、欄の右の端（＝「信用金庫」の手前）のあいだに、1 字ぶん（10pt 近く）の空き
    assert w - (x + 2 * float(spot.get("size") or 10)) >= 8, (x, w)
    assert doc is not None


# ---------------------------------------------------------------- ○：near の字の中の囲む字（circlepick）
def _two_rows_pdf() -> bytes:
    return _pdf([(70, 700, 10, "生年月日　大・昭・平・令　　年"), (70, 660, 10, "交付日　大・昭・平・令　　年"),
                 (70, 620, 10, "住所　都県市"), (70, 600, 10, "東 京 都　港区")])


def test_pdf_option_list_near_circles_the_item_inside_it(tmp_path):
    r, _ = _write(tmp_path, _two_rows_pdf(), "t.pdf", {"@c1": {"circle": "昭", "near": "生年月日　大・昭・平・令"}})
    assert r["proof_ok"] and r["read_back_ok"]
    assert "大・【昭】" in r["circled"][0]["around"] and "生年月日" in r["circled"][0]["context"]


def test_pdf_option_list_near_twice_asks_for_occurrence(tmp_path):
    with pytest.raises(api.ApiError, match="選択肢の並び.*occurrence"):
        _write(tmp_path, _two_rows_pdf(), "t.pdf", {"@c1": {"circle": "昭", "near": "大・昭・平・令"}})
    r, _ = _write(tmp_path, _two_rows_pdf(), "t.pdf", {"@c1": {"circle": "昭", "near": "大・昭・平・令", "occurrence": 2}})
    assert "交付日" in r["circled"][0]["context"]


def test_pdf_target_inside_near_word_does_not_jump_to_another_line(tmp_path):
    # 「都県市」の「都」（字の並びの中）：下の行の「東京都」へ黙って移らない。near の中の所も番号で選べる
    with pytest.raises(api.ApiError, match="別の行.*【都】県市"):
        _write(tmp_path, _two_rows_pdf(), "t.pdf", {"@c1": {"circle": "都", "near": "都県市"}})


def test_word_near_with_occurrence_circles_inside_a_word(tmp_path):
    # 「(３)市町村民税…」の「(３)」：後ろが字なので区切られていないが、番号で指せば囲む（PDF と同じ）
    body = wp("(２)非課税世帯で８０万円以下") + wp("(３)市町村民税非課税世帯で（２）以外のもの")
    r, _ = _write(tmp_path, make_docx(body), "t.docx",
                  {"@c1": {"circle": "(３)", "near": "（２）以外のもの", "occurrence": 1, "style": "border"}})
    assert r["proof_ok"] and r["read_back_ok"]
    assert "【(３)】市町村" in r["circled"][0]["around"]


def test_word_two_borders_with_the_same_heading_read_back(tmp_path):
    body = (wp("認定　有・無") + wp("認定　有・無"))
    r, _ = _write(tmp_path, make_docx(body), "t.docx",
                  {"@c1": {"circle": "有", "occurrence": 1, "style": "border"},
                   "@c2": {"circle": "無", "occurrence": 2, "style": "border"}})
    assert r["proof_ok"] and r["read_back_ok"], r["read_back_differences"]


def test_word_deletion_edit_is_not_an_audit_failure(tmp_path):
    body = wp("銀行　信金　　　　　　（支店コード）")
    r, out = _write(tmp_path, make_docx(body), "t.docx",
                    {"@1": {"old": "銀行", "new": "架空　銀行"}, "@2": {"old": "　　　　　　（支店", "new": "　　　（支店"}})
    assert r["proof_ok"] and r["read_back_ok"]
    a = api.audit(out)
    bad = [c for c in a["checks"] if not c["ok"] and "しおり" in str(c.get("items", ""))]
    assert not bad, bad


def test_word_option_list_near(tmp_path):
    body = wp("生年月日　S・H　年　月　日生") + wp("H　S　T")
    r, _ = _write(tmp_path, make_docx(body), "t.docx", {"@c1": {"circle": "S", "near": "S・H"}})
    assert r["proof_ok"] and r["read_back_ok"]
    assert "【S】・H" in r["circled"][0]["around"]
