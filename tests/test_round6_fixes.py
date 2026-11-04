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
def test_after_label_in_a_narrow_heading_cell_warns_with_the_free_place(tmp_path):
    rules = (b"0.8 w 60 700 m 400 700 l S 60 680 m 400 680 l S 60 700 m 60 680 l S 140 700 m 140 680 l S "
             b"400 700 m 400 680 l S\n")
    data = _pdf([(70, 686, 10, "氏　　名"), (340, 686, 10, "番号")], rules)
    r, _ = _write(tmp_path, data, "t.pdf", {"@1": {"after": "氏　　名", "value": "山田　花子", "page": 1}})
    assert r["proof_ok"] and r["read_back_ok"]
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.anchor import resolve_write

    sp = resolve_write(PdfDocument(data), {"after": "氏　　名", "value": "山田　花子", "page": 1}, {}, {})
    assert sp["rect"][2] <= 140  # 置く所は見出しのマスのまま（エージェントが決める）
    # 字がとても小さくなることと、右のマスの字のない所（mm）を知らせる
    assert any("狭い空き" in w and "右のマスの字のない所" in w and "box" in w for w in r["warnings"]), r["warnings"]


# ---------------------------------------------------------------- PDF：置き方はエージェントが決める（align・pad_mm）
def _gap_spot(data):
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.window import build

    win = build(PdfDocument(data))
    return next(f for pg in win["pages"] for f in pg["fields"] if "信用金庫" in (f.get("label") or ""))


def _drawn_x(out: str) -> float:
    m = re.findall(rb"1 0 0 1 ([\d.]+) ([\d.]+) Tm", open(out, "rb").read())
    return float(m[-1][0])


def test_default_is_left_and_align_right_is_honoured(tmp_path):
    rules = b"0.8 w 60 700 m 400 700 l S 60 680 m 400 680 l S\n"
    data = _pdf([(300, 686, 10, "信用金庫")], rules)
    spot = _gap_spot(data)
    r, out = _write(tmp_path, data, "t.pdf", {spot["id"]: "架空"})
    assert r["proof_ok"] and r["read_back_ok"]
    assert _drawn_x(out) <= 2.0  # 既定は左（枠の左の端から、決まった小さな余白）
    rect = r["written_rects"][spot["id"]]
    assert rect["align"] == "left" and rect["page"] == 1 and rect["w"] > 0 and rect["size_pt"] > 0
    r, out = _write(tmp_path, data, "t2.pdf", {spot["id"]: {"value": "架空", "align": "right", "pad_mm": 3}})
    assert r["proof_ok"] and r["read_back_ok"]
    w = spot["rect"][2] - spot["rect"][0]
    x = _drawn_x(out)
    assert w - (x + 2 * 10) == pytest.approx(3 * 72 / 25.4, abs=1.5), (x, w)  # 右の端から 3 mm
    assert r["written_rects"][spot["id"]]["align"] == "right"
    with pytest.raises(api.ApiError, match="align"):
        _write(tmp_path, data, "t3.pdf", {spot["id"]: {"value": "架空", "align": "justify"}})


# ---------------------------------------------------------------- PDF：字の行が縦に並ぶマスの欄
def _stack_pdf() -> bytes:
    rules = (b"0.8 w 60 700 m 300 700 l S 60 630 m 300 630 l S 60 700 m 60 630 l S 300 700 m 300 630 l S "
             b"20 700 m 60 700 l S 20 630 m 60 630 l S 20 700 m 20 630 l S\n")
    return _pdf([(25, 662, 10, "金融機関"), (200, 688, 10, "銀行"), (200, 674, 10, "信用金庫"), (200, 660, 10, "信用組合"),
                 (200, 646, 10, "農協")], rules)


def _stack_spots(data):
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.window import build

    win = build(PdfDocument(data))
    return [f for pg in win["pages"] for f in pg["fields"] if f.get("shape") == "lead"]


def test_value_beside_stacked_lines_has_a_place_on_each_line(tmp_path):
    # 字の行が縦に並ぶマス：行の前の空きは行ごとの所。どの行に続けて書くかは、所を選ぶことで決まる（行の高さの迷いがない）
    data = _stack_pdf()
    spots = _stack_spots(data)
    assert {f["facts"]["line"] for f in spots} == {"銀行", "信用金庫", "信用組合", "農協"}
    spot = next(f for f in spots if f["facts"]["line"] == "信用金庫")
    r, _ = _write(tmp_path, data, "t.pdf", {spot["id"]: "架空"})
    assert r["proof_ok"] and r["read_back_ok"]
    assert not any("行並ぶマス" in w for w in r["warnings"]), r["warnings"]
    got = r["written_rects"][spot["id"]]
    assert abs(got["y"] - round(spot["rect"][1] * 25.4 / 72, 1)) <= 0.1


# ---------------------------------------------------------------- PDF：一続きのマスの字の大きさをそろえる
def test_adjacent_cells_in_one_fill_get_one_font_size():
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.fill import Filler, _uniform_sizes

    doc = PdfDocument(_pdf([(70, 700, 10, "口座名義")]))
    f = Filler(doc)
    spots = {"p1.1": (1, {"id": "p1.1", "kind": "text", "rect": [100, 600, 128, 628], "size": 10.5}),
             "p1.2": (1, {"id": "p1.2", "kind": "text", "rect": [128, 600, 156, 628], "size": 10.5}),
             "p1.3": (1, {"id": "p1.3", "kind": "text", "rect": [156, 600, 167, 628], "size": 10.5})}
    _uniform_sizes(f, doc, {"p1.1": "ヤ", "p1.2": "マ", "p1.3": "コ"}, spots)
    assert f.size_override.get("p1.1") and f.size_override["p1.1"] == f.size_override["p1.2"] < 10.5
    assert "p1.3" not in f.size_override  # いちばん小さい欄は、もとの大きさのまま
    assert any("そろえました" in w for w in f.warnings)


# ---------------------------------------------------------------- PDF：重なりは、はっきりしたものだけ断り、ほかは位置を添えて知らせる
def test_write_across_a_curved_stamp_is_warned_with_the_position(tmp_path):
    # 「見本」の印のような曲線（半円）が、書く枠を横切る：断らずに、位置（mm）を添えて知らせる
    stamp = b"1 w 100 640 m 100 700 200 700 200 640 c S\n"
    data = _pdf([(70, 760, 10, "氏名")], stamp)
    top = (842 - 682) * 25.4 / 72
    r, _ = _write(tmp_path, data, "t.pdf", {"@1": {"page": 1, "box": [30, top - 3, 80, top + 3], "value": "山田 花子"}})
    assert r["proof_ok"] and r["read_back_ok"]
    assert any("曲線" in w and "mm" in w for w in r["warnings"]), r["warnings"]


def test_write_on_top_of_printed_text_is_still_refused(tmp_path):
    data = _pdf([(70, 700, 10, "見本見本見本")])
    top = (842 - 710) * 25.4 / 72
    left = 70 * 25.4 / 72
    with pytest.raises(api.ApiError, match="ページの字|の上に重なります"):
        _write(tmp_path, data, "t.pdf", {"@1": {"page": 1, "box": [left, top, left + 21, top + 5], "value": "山田花子"}})


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
