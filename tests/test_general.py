"""特定の様式に頼らない、一般的な書き方への対応のテスト。

公開されている様式を広く集めて検査したときに見つかった「とどくの側の一般的な原因」を、
小さなブックで再現して確かめる。正解はExcelの仕様どおりの値（手計算）で、ほかの表計算ソフトは使わない。
"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku.formula.numfmt import format_value
from todoku.formula.parser import parse, parse_cell
from todoku.xlsx import window
from todoku.xlsx.calc import Calc, verify
from todoku.xlsx.edit import EditError, fill, interpret_input
from todoku.xlsx.workbook import Workbook

openpyxl = pytest.importorskip("openpyxl")


def book(build) -> Workbook:
    wb = openpyxl.Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return Workbook(buf.getvalue(), "test.xlsx")


# ------------------------------------------------------------------ 数式
def test_structured_references():
    from openpyxl.worksheet.table import Table

    def build(wb):
        ws = wb.active
        ws.title = "明細"
        ws.append(["品名", "数量", "単価", "金額"])
        for name, q, p in [("A", 2, 100), ("B", 3, 250), ("C", 1, 80)]:
            ws.append([name, q, p, "=テーブル1[[#This Row],[数量]]*テーブル1[[#This Row],[単価]]"])
        ws.add_table(Table(displayName="テーブル1", ref="A1:D4"))
        ws["F1"] = "=SUM(テーブル1[金額])"
        ws["F2"] = "=COUNTA(テーブル1[#Headers])"
        ws["F3"] = "=SUM(テーブル1[[数量]:[単価]])"

    c = Calc(book(build))
    assert [c.cell_value("明細", r, 4) for r in (2, 3, 4)] == [200, 750, 80]
    assert (c.cell_value("明細", 1, 6), c.cell_value("明細", 2, 6), c.cell_value("明細", 3, 6)) == (1030, 4, 436)


def test_subtotal_skips_nested_and_hidden_rows():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        for r, v in enumerate([10, 20, 30, 40], start=1):
            ws.cell(r, 1, v)
        ws["A5"] = "=SUBTOTAL(9,A1:A4)"
        ws["A6"] = "=SUBTOTAL(9,A1:A5)"  # 入れ子の小計は数えない
        ws["B1"] = "=SUBTOTAL(109,A1:A4)"  # 手で隠した行（3行目）を除く
        ws["B2"] = "=SUBTOTAL(1,A1:A4)"  # 1〜11 は手で隠した行も数える
        ws.row_dimensions[3].hidden = True

    c = Calc(book(build))
    assert (c.cell_value("S", 5, 1), c.cell_value("S", 6, 1), c.cell_value("S", 1, 2), c.cell_value("S", 2, 2)) == (100, 100, 70, 25)


def test_let_weekday_address_cell():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = 5
        ws["B1"] = "=_xlfn.LET(_xlpm.x,A1*2,_xlpm.y,_xlpm.x+1,_xlpm.x*_xlpm.y)"
        ws["B2"] = "=WEEKDAY(DATE(2026,11,2),2)"  # 2026年11月2日は月曜
        ws["B3"] = "=WEEKDAY(DATE(2026,11,2),12)"  # 火曜を1とすると月曜は7
        ws["B4"] = '=ADDRESS(2,3,2,FALSE)'
        ws["B5"] = '=ADDRESS(1,1,1,TRUE,"Sheet 2")'
        ws["B6"] = '=CELL("row",A1)+CELL("col",C1)'
        ws["B7"] = '=_xlfn.IFS(A1>9,"大",A1>3,"中",TRUE,"小")'

    c = Calc(book(build))
    assert [c.cell_value("S", r, 2) for r in range(1, 8)] == [110, 1, 7, "R2C[3]", "'Sheet 2'!$A$1", 4, "中"]


def test_value_depending_on_today_is_volatile():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "=YEAR(TODAY())"
        ws["A2"] = "=A1+1"  # TODAY() の結果を使う数式も、計算のたびに変わる

    wb = book(build)
    c = Calc(wb)
    c.cell_value("S", 2, 1)
    assert c.status[("S", 1, 1)][0] == "volatile" and c.status[("S", 2, 1)][0] == "volatile"


def test_external_reference_uses_cached_value():
    # 外部ブックへのリンクは、保存されている値を使う（Excelがリンクを更新しないときと同じ）
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "=[1]元データ!$B$2*2"

    src = book(build).zip.original
    zin = zipfile.ZipFile(io.BytesIO(src))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "xl/workbook.xml":
                data = data.replace(b"</sheets>", b'</sheets><externalReferences><externalReference r:id="rIdX"/></externalReferences>')
                if b'xmlns:r=' not in data:
                    data = data.replace(b"<workbook ", b'<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ', 1)
            if item.filename == "xl/_rels/workbook.xml.rels":
                data = data.replace(b"</Relationships>", b'<Relationship Id="rIdX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/></Relationships>')
            z.writestr(item, data)
        z.writestr("xl/externalLinks/externalLink1.xml",
                   '<externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><externalBook>'
                   '<sheetNames><sheetName val="元データ"/></sheetNames><sheetDataSet><sheetData sheetId="0">'
                   '<row r="2"><cell r="B2"><v>21</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>')
    c = Calc(Workbook(out.getvalue()))
    assert c.cell_value("S", 1, 1) == 42


def test_strict_open_xml_namespaces_are_read():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "名前"
        ws["B1"] = 12

    src = book(build).zip.original
    zin = zipfile.ZipFile(io.BytesIO(src))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml") and item.filename.startswith("xl/"):
                data = data.replace(b"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
                                    b"http://purl.oclc.org/ooxml/spreadsheetml/main")
            z.writestr(item, data)
    wb = Workbook(out.getvalue())
    sh = wb.sheet("S")
    assert sh.cells[(1, 1)].value == "名前" and sh.cells[(1, 2)].value == 12


# ------------------------------------------------------------------ 入力
def test_circled_digits_are_not_turned_into_numbers():
    def build(wb):
        wb.active.title = "S"

    wb = book(build)
    assert interpret_input(wb, "S", 1, 1, "①")[0] == "①"  # NFKC なら 1 になってしまう
    assert interpret_input(wb, "S", 1, 1, "１２")[0] == 12.0  # 全角の数字は数値（Excelと同じ）
    assert interpret_input(wb, "S", 1, 1, "㈱とどく")[0] == "㈱とどく"


def test_rule_formula_is_relative_to_the_top_left_of_all_ranges():
    from openpyxl.worksheet.datavalidation import DataValidation

    def build(wb):
        ws = wb.active
        ws.title = "S"
        # 範囲を「B5 B2」の順に書いても、相対参照の基準は全体の左上（B2）
        dv = DataValidation(type="custom", formula1="ISNUMBER(B2)", showErrorMessage=True)
        dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange("B5 B2")
        ws.add_data_validation(dv)

    wb = book(build)
    fill(wb, {"S!B2": 3, "S!B5": 4})
    with pytest.raises(EditError):
        fill(wb, {"S!B5": "文字"})


def test_overlapping_rules_warn_instead_of_guessing():
    from openpyxl.worksheet.datavalidation import DataValidation

    def build(wb):
        ws = wb.active
        ws.title = "S"
        for f in ('"○"', '"甲,乙"'):
            dv = DataValidation(type="list", formula1=f, showErrorMessage=True)
            dv.add("C3")
            ws.add_data_validation(dv)

    wb = book(build)
    res = fill(wb, {"S!C3": "甲"})  # 片方の規則には合う：Excelがどちらを使うかは決まらないので、止めずに知らせる
    assert any("重なって" in w for w in res.warnings)
    with pytest.raises(EditError):
        fill(wb, {"S!C3": "丙"})  # どの規則にも合わない


# ------------------------------------------------------------------ 表示と窓
@pytest.mark.parametrize("fmt, value, text", [
    ("#,##0.##", 0, "0."), ("0.?", 0, "0. "), ("0.#?", 1.5, "1.5 "), ("???0", 12, "  12"), ("#", 0, ""),
])
def test_number_format_placeholders(fmt, value, text):
    assert format_value(float(value), fmt) == text


def test_two_row_records_are_summarized():
    def build(wb):
        ws = wb.active
        ws.title = "勤務表"
        ws["A1"] = "番号"
        ws["B1"] = "氏名"
        for i in range(30):
            r = 2 + i * 2
            ws.cell(r, 1, i + 1)
            ws.cell(r, 3, "シフト")
            ws.cell(r + 1, 3, "時間")
            ws.merge_cells(start_row=r, start_column=2, end_row=r + 1, end_column=2)
        ws.protection.sheet = True
        for i in range(30):
            ws.cell(2 + i * 2, 2).protection = openpyxl.styles.Protection(locked=False)

    doc = window.build(book(build))
    lines = doc["sheets"][0]["reading"]
    assert any("2行1組" in x["text"] for x in lines)  # 2行で1人分の表は、くり返しをまとめる
    assert len(lines) < 12


def test_large_window_is_split_into_readable_parts():
    def build(wb):
        for k in range(3):
            ws = wb.active if k == 0 else wb.create_sheet()
            ws.title = f"表{k + 1}"
            for r in range(1, 1500):
                ws.cell(r, 1, f"項目{k}-{r}の説明文です。これは読み順に出る文字です。")
                ws.cell(r, 2, r * 1.5)

    doc = window.build(book(build))
    parts = window.markdown_parts(doc, budget=60_000)
    assert len(parts) > 2 and parts[0][0] == "目次"
    assert max(len(t) for _, t in parts) <= 70_000
    whole = window.to_markdown(doc)
    assert all(f"項目2-{r}の" in whole for r in (1, 1499))


def test_side_by_side_tables_keep_their_own_row_labels():
    # 収入の表と支出の表が横に並ぶ予算様式（英語）。行の見出しは、その表の左の文字を使う
    def build(wb):
        ws = wb.active
        ws.title = "Budget"
        ws["A1"], ws["E1"] = "Income", "Expenditure"
        ws.merge_cells("A1:C1")
        ws.merge_cells("E1:G1")
        ws["A2"], ws["B2"], ws["C2"] = "Source", "Cash", "In-kind"
        ws["E2"], ws["F2"], ws["G2"] = "Item", "Cash", "In-kind"
        for r, (a, e) in enumerate([("Grant", "Staff"), ("Donation", "Venue"), ("Fees", "Printing")], start=3):
            ws.cell(r, 1, a)
            ws.cell(r, 5, e)
        ws.protection.sheet = True
        for r in range(3, 6):
            for c in (2, 3, 6, 7):
                ws.cell(r, c).protection = openpyxl.styles.Protection(locked=False)

    doc = window.build(book(build))
    f = {x["cell"]: x["label"] for x in doc["sheets"][0]["fields"]}
    assert f["B3"] == "Grant／Income／Cash"
    assert f["G4"] == "Venue／Expenditure／In-kind"
