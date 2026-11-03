"""第4回の見直しの直し（表計算の書き込みと、どの形式にも共通の書き込みの入口）：合成した書類で確かめる。

A3 書いたセルの書式（見えない字は書かない・取り消し線は伝える）／A4 字ごとの書式を残す／A8 計算し直せない数式を伝える／
A9・C5 表示が書いた字と違う数／B1 収まらない字／C1 結合の中のセル／C2 同じセルへの2つの指示／C3 読み戻しのはい・いいえ／
D1 制御文字／D2 大きすぎる数／D4 _xHHHH_／D5 2つのうち1つの両方。
"""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest

from todoku import api

openpyxl = pytest.importorskip("openpyxl")
FIX = Path(__file__).parent / "fixtures"


def _xlsx(tmp_path, build, name="a.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    build(ws)
    p = tmp_path / name
    wb.save(p)
    return p


def _sheet_xml(path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/worksheets/sheet1.xml").decode("utf-8")


# ====================================================================== A4 字ごとの書式
def _rich(ws):
    from openpyxl.cell.rich_text import CellRichText, TextBlock
    from openpyxl.cell.text import InlineFont

    ws["A5"] = CellRichText([TextBlock(InlineFont(strike=True, color="FFFF0000"), "旧住所"), "　新住所：　　　　"])


def test_partial_write_keeps_struck_red_run(tmp_path):
    p = _xlsx(tmp_path, _rich)
    for vals in ({"S!A5": [None, "東京都千代田区1-1"]}, {"@1": {"old": "新住所：", "new": "新住所：東京", "cell": "S!A5"}}):
        r = api.fill(str(p), vals, str(tmp_path / "o.xlsx"), overwrite=True)
        assert r["proof_ok"] and r["read_back_ok"] and r["verified"], r
        xml = _sheet_xml(r["out_path"])
        cell = xml[xml.index('<c r="A5"'):]
        cell = cell[:cell.index("</c>")]
        assert "<strike" in cell and "FFFF0000" in cell and "<t>旧住所</t>" in cell
        assert "旧住所" not in cell.split("</r>", 1)[1]  # 書いた字は、取り消し線の <r> に入らない
        wb = openpyxl.load_workbook(r["out_path"], rich_text=True)
        assert str(wb["S"]["A5"].value).startswith("旧住所　新住所：東京")


def test_rich_proof_catches_lost_run_props():
    from todoku.xlsx.richtext import Rst, kept_problems

    before = Rst(b"", [("旧住所", b"<rPr><strike/></rPr>"), ("　新住所：", b"")])
    after_ok = Rst(b"", [("旧住所", b"<rPr><strike/></rPr>"), ("　新住所：東京", b"")])
    after_lost = Rst(b"", [("旧住所　新住所：東京", None)])
    assert kept_problems(before, after_ok) == []
    assert kept_problems(before, after_lost)


def test_insert_at_run_boundary_is_refused():
    from todoku.xlsx.richtext import Rst, splice

    rst = Rst(b"", [("旧", b"<rPr><strike/></rPr>"), ("新", b"")])
    with pytest.raises(ValueError, match="境目"):
        splice(rst, "旧X新")


def test_phonetic_kept_and_shifted():
    from todoku.xlsx.richtext import parse_rst, splice

    rst = parse_rst('<t>氏名　　様</t><rPh sb="0" eb="2"><t>シメイ</t></rPh><rPh sb="4" eb="5"><t>サマ</t></rPh>'
                    '<phoneticPr fontId="1"/>'.encode())
    inner, _props, dropped, _w = splice(rst, "氏名山田様")
    assert b'<rPh sb="0" eb="2">' in inner and b'<rPh sb="4" eb="5">' in inner and b"<phoneticPr" in inner
    assert dropped == []


def test_xls_partial_write_keeps_format_runs(tmp_path):
    from todoku.xls.book import XlsWorkbook

    r = api.fill(str(FIX / "r4_rich.xls"), {"S!A5": [None, "東京"]}, str(tmp_path / "o.xls"))
    assert r["verified"] and not r.get("written_formatting")  # 書いた字は、取り消し線のない字体の所に入る
    w = XlsWorkbook((tmp_path / "o.xls").read_bytes())
    i = w.rich_cells[("S", 5, 1)]
    assert w.shared_strings[i] == "旧住所　新住所：東京" and w.sst_runs[i] == [(0, 5), (3, 0)]
    with pytest.raises(api.ApiError, match="境目"):
        api.fill(str(FIX / "r4_rich.xls"), {"@1": {"old": "旧住所", "new": "旧住所X", "cell": "S!A5"}}, str(tmp_path / "o2.xls"))


def test_xls_rich_proof_catches_lost_runs(tmp_path):
    from todoku.xlsx.proof import _xls_rich_problems

    blob = (FIX / "r4_rich.xls").read_bytes()
    r = api.fill(str(FIX / "r4_rich.xls"), {"S!A5": {"value": "旧住所　新住所：東京", "overwrite_text": True}}, str(tmp_path / "o.xls"))
    assert _xls_rich_problems(blob, (tmp_path / "o.xls").read_bytes(), {("S", "A5")})  # ただの字で書き直すと書式が消える


# ====================================================================== A3 書いたセルの書式
def _styled(ws):
    from openpyxl.styles import Font, PatternFill

    ws["B1"].font = Font(color="FFFFFFFF")
    ws["B2"].number_format = ";;;"
    ws["B3"].font = Font(strike=True)
    ws["B5"].fill = PatternFill("solid", fgColor="FF000000")
    ws["A7"] = "金額　　　　円"
    ws["A7"].font = Font(strike=True)


@pytest.mark.parametrize("cell", ["S!B1", "S!B2", "S!B5"])
def test_invisible_cells_are_refused(tmp_path, cell):
    p = _xlsx(tmp_path, _styled)
    with pytest.raises(api.ApiError, match="見えなく"):
        api.fill(str(p), {cell: "山田"}, str(tmp_path / "o.xlsx"))
    r = api.fill(str(p), {cell: "山田"}, str(tmp_path / "o.xlsx"), force=True)  # 本人が確かめたうえでなら書ける
    assert r["written_formatting"]


def test_strike_is_reported(tmp_path):
    p = _xlsx(tmp_path, _styled)
    for vals in ({"S!B3": "東京"}, {"S!A7": ["1000"]}):
        r = api.fill(str(p), vals, str(tmp_path / "o.xlsx"), overwrite=True)
        assert r["verified"]
        assert any("取り消し線" in x for f in r["written_formatting"] for x in f["formatting"])
        assert "取り消し線" in r["say_to_user"]


def test_checkbox_link_cell_hidden_by_format_is_not_refused(tmp_path):
    # 部品につながったセルは書式で隠すのが決まり：ここでは見ない（部品のない合成なので、番地での書き込みは見える字で）
    p = _xlsx(tmp_path, _styled)
    r = api.fill(str(p), {"S!B4": "メモ"}, str(tmp_path / "o.xlsx"))
    assert r["verified"] and not r.get("written_formatting")


# ====================================================================== A8 計算し直せない数式
def test_stale_formula_is_told(tmp_path):
    p = _xlsx(tmp_path, lambda ws: (ws.__setitem__("A3", "計"), ws.__setitem__("B3", "=BAHTTEXT(B1)")))
    r = api.fill(str(p), {"S!B1": 100}, str(tmp_path / "o.xlsx"))
    assert r["stale_formulas"] and "S!B3" in r["stale_formulas"][0]
    assert "Excel で開いたときに計算し直されます" in r["say_to_user"] and "古い" in r["say_to_user"]


# ====================================================================== A9・C5 表示が書いた字と違う数
def _general(ws):
    ws["A1"] = "番号"
    ws["A5"] = "単価"
    ws["A6"] = "数量"
    ws["B7"] = "=B5*B6"


def test_twelve_digit_number_is_written_as_text(tmp_path):
    p = _xlsx(tmp_path, _general)
    for v in ("123456789012", 123456789012):
        r = api.fill(str(p), {"S!B1": v}, str(tmp_path / "o.xlsx"), overwrite=True)
        assert r["verified"]
        assert openpyxl.load_workbook(r["out_path"])["S"]["B1"].value == "123456789012"
        assert "1.23457E+11" in r["say_to_user"] and "E+11" not in r["written"][0]["value"]


def test_currency_and_percent_in_general_cells(tmp_path):
    p = _xlsx(tmp_path, _general)
    r = api.fill(str(p), {"S!B2": "¥1,000", "S!B3": "50%"}, str(tmp_path / "o.xlsx"))
    ws = openpyxl.load_workbook(r["out_path"])["S"]
    assert ws["B2"].value == "¥1,000" and ws["B3"].value == "50%"  # 計算に使われないセル：書いたとおりの字
    assert r["verified"] and "書いたとおりの字" in r["say_to_user"]
    r = api.fill(str(p), {"S!B5": "$100"}, str(tmp_path / "o2.xlsx"))  # 数式が参照するセル：数にして、見え方を伝える
    assert openpyxl.load_workbook(r["out_path"])["S"]["B5"].value == 100
    assert r["verified"] and "「100」と表示されます" in r["say_to_user"]


def test_formatted_cells_keep_numbers(tmp_path):
    def b(ws):
        ws["B1"].number_format = "0%"
        ws["B2"].number_format = "#,##0"
    p = _xlsx(tmp_path, b)
    r = api.fill(str(p), {"S!B1": "50%", "S!B2": "1,000"}, str(tmp_path / "o.xlsx"))
    ws = openpyxl.load_workbook(r["out_path"])["S"]
    assert ws["B1"].value == 0.5 and ws["B2"].value == 1000 and r["verified"]


def test_xls_and_ods_follow_the_same_rule(tmp_path):
    for f in ("r4_general.xls", "r4_general.ods"):
        r = api.fill(str(FIX / f), {"申請!B2": "123456789012"}, str(tmp_path / ("o." + f.split(".")[-1])), overwrite=True)
        assert r["verified"] and "1.23457E+11" in r["say_to_user"]
        assert r["written"][0]["value"] == "123456789012"


# ====================================================================== B1 収まらない字
def test_overflow_and_wrap_are_reported(tmp_path):
    from openpyxl.styles import Alignment

    def b(ws):
        ws["C3"] = "次"
        ws.column_dimensions["B"].width = 6
        ws["B4"].alignment = Alignment(wrap_text=True)
        ws.row_dimensions[4].height = 15
    p = _xlsx(tmp_path, b)
    r = api.fill(str(p), {"S!B3": "とても長い住所の文字列です東京都"}, str(tmp_path / "o.xlsx"))
    assert r["verified"] and r["fit"] and "C3 の字で切れて見える" in r["say_to_user"]
    r = api.fill(str(p), {"S!B4": "とても長い住所の文字列です東京都千代田区"}, str(tmp_path / "o.xlsx"), overwrite=True)
    assert r["fit"] and "下の行が切れて見える" in r["say_to_user"]
    r = api.fill(str(p), {"S!B5": "山田"}, str(tmp_path / "o.xlsx"), overwrite=True)
    assert not r.get("fit")


def test_too_long_text_is_refused(tmp_path):
    p = _xlsx(tmp_path, _general)
    with pytest.raises(api.ApiError, match="32,767"):
        api.fill(str(p), {"S!B1": "あ" * 32768}, str(tmp_path / "o.xlsx"))


# ====================================================================== C1・C2
def _merged(ws):
    ws["A9"] = "結合"
    ws.merge_cells("B9:D9")


def test_hidden_merge_cell_is_refused(tmp_path):
    p = _xlsx(tmp_path, _merged)
    with pytest.raises(api.ApiError, match="S!B9 に書いてください"):
        api.fill(str(p), {"S!C9": "x"}, str(tmp_path / "o.xlsx"))


@pytest.mark.parametrize("vals", [{"S!B9": "x", "S!B9:D9": "y"}, {"S!b1": "x", "S!B1": "y"}])
def test_two_ids_for_one_cell_are_refused(tmp_path, vals):
    p = _xlsx(tmp_path, _merged)
    with pytest.raises(api.ApiError, match="同じセル"):
        api.fill(str(p), vals, str(tmp_path / "o.xlsx"))


def test_two_anchors_in_one_cell_both_apply(tmp_path):
    p = _xlsx(tmp_path, lambda ws: ws.__setitem__("A1", "令和　　年　　月"))
    r = api.fill(str(p), {"@1": {"old": "令和　　年", "new": "令和8年", "cell": "S!A1"},
                          "@2": {"old": "　　月", "new": "11月", "cell": "S!A1"}}, str(tmp_path / "o.xlsx"))
    assert r["verified"] and openpyxl.load_workbook(r["out_path"])["S"]["A1"].value == "令和8年11月"


# ====================================================================== C3 読み戻しのはい・いいえ
def test_readback_bool_uses_writer_vocabulary():
    assert api._writer_bool("xlsx", "有") is True and api._writer_bool("xlsx", "☑") is True
    assert api._writer_bool("xlsx", "ON") is True and api._writer_bool("xlsx", "無") is False
    assert api._writer_bool("pdf", "レ") is True and api._writer_bool("pdf", "○") is True
    assert api._writer_bool("docx", "■") is True
    assert api._writer_bool("pdf", "たぶん") is None


# ====================================================================== D1・D2・D4
@pytest.mark.parametrize("bad", ["山田\x07太郎", "山田\ud800", "a\x00b", "a\x1bb"])
def test_control_chars_are_refused_for_every_format(tmp_path, bad):
    p = _xlsx(tmp_path, _general)
    with pytest.raises(api.ApiError, match="書けない字"):
        api.fill(str(p), {"S!B1": bad}, str(tmp_path / "o.xlsx"))
    with pytest.raises(api.ApiError, match="書けない字"):
        api.fill(str(p), {"S!B1": [bad]}, str(tmp_path / "o.xlsx"))
    t = tmp_path / "a.txt"
    t.write_text("氏名：\n", encoding="utf-8")
    with pytest.raises(api.ApiError, match="書けない字"):
        api.fill(str(t), {"@1": {"old": "氏名：", "new": "氏名：" + bad}}, str(tmp_path / "o.txt"))


def test_crlf_is_normalized():
    from todoku.inputcheck import normalize_values

    assert normalize_values({"a": "x\r\ny\rz", "b": ["p\r\n"]}) == {"a": "x\ny\nz", "b": ["p\n"]}


def test_huge_exponent_is_text_not_crash(tmp_path):
    p = _xlsx(tmp_path, _general)
    r = api.fill(str(p), {"S!B1": "1e400"}, str(tmp_path / "o.xlsx"))
    assert r["verified"] and openpyxl.load_workbook(r["out_path"])["S"]["B1"].value == "1e400"
    assert "扱える数の大きさ" in r["say_to_user"]


def test_xhhhh_text_is_escaped(tmp_path):
    from todoku.xlsx.workbook import Workbook

    p = _xlsx(tmp_path, _general)
    r = api.fill(str(p), {"S!B1": "a_x0041_b"}, str(tmp_path / "o.xlsx"))
    assert r["verified"] and "a_x005F_x0041_b" in _sheet_xml(r["out_path"])
    assert Workbook(Path(r["out_path"]).read_bytes()).sheet("S").cells[(1, 2)].value == "a_x0041_b"


def test_reader_decodes_xhhhh():
    import xml.etree.ElementTree as ET

    from todoku.xlsx.workbook import _rich_text

    ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    el = ET.fromstring(f'<is xmlns="{ns}"><t>a_x0041_b_x005F_x0041_</t></is>')
    assert _rich_text(el) == "aAb_x0041_"


# ====================================================================== D5 2つのうち1つの両方
def test_both_boxes_and_both_circles_warn(tmp_path):
    p = _xlsx(tmp_path, lambda ws: (ws.__setitem__("A2", "有・無"), ws.__setitem__("A6", "□はい　□いいえ")))
    r = api.fill(str(p), {"S!A6": [True, True]}, str(tmp_path / "o.xlsx"))
    assert r["verified"] and "両方に印" in r["say_to_user"]
    r = api.fill(str(p), {"S!A6": [True, None]}, str(tmp_path / "o.xlsx"), overwrite=True)
    assert "両方に印" not in r["say_to_user"]
    r = api.fill(str(p), {"@c1": {"circle": "有", "cell": "S!A2"}, "@c2": {"circle": "無", "cell": "S!A2"}},
                 str(tmp_path / "o.xlsx"), overwrite=True)
    assert "選択肢の全部" in r["say_to_user"]
