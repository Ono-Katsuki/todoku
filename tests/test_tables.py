"""Excel の表を、見出しの列で行ごとに取り出す（結合セル・2行の見出し・表示形式）。"""

from __future__ import annotations

import csv
import io

import pytest

from todoku import api

openpyxl = pytest.importorskip("openpyxl")


def _book(tmp_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "居宅"
    ws["A1"] = "サービスコード表"
    ws["A3"], ws["C3"], ws["D3"], ws["F3"] = "サービスコード", "略称", "算定項目", "合成"
    ws.merge_cells("D3:E3")
    ws["A4"], ws["B4"], ws["F4"] = "種類", "項目", "単位数"
    rows = [(11, 1111, "身体日０．５", "(1)日中", 256), (11, 1112, "身体日０．５・２人", None, 256), (11, "J052", "身体日増", "(2)増分", 1080)]
    for i, (k, c, n, item, u) in enumerate(rows, 5):
        ws.cell(i, 1, k), ws.cell(i, 2, c), ws.cell(i, 3, n), ws.cell(i, 6, u)
        if item:
            ws.cell(i, 4, item)
    ws.merge_cells("D5:D6")  # 算定項目が2行にまたがる（結合）
    ws["F7"].number_format = "#,##0"
    ws["A9"] = "※注：これは表の後ろの注"
    other = wb.create_sheet("説明")
    other["A1"] = "見出しのないシート"
    p = tmp_path / "codes.xlsx"
    buf = io.BytesIO()
    wb.save(buf)
    p.write_bytes(buf.getvalue())
    return str(p)


def test_two_row_header_and_merged_cells(tmp_path):
    f = _book(tmp_path)
    t = api.read_table(f, "合成")
    assert t["columns"] == ["サービスコード／種類", "項目", "略称", "算定項目（D）", "算定項目（E）", "合成／単位数"]
    rows = t["all_rows"]
    assert [r["項目"] for r in rows[:3]] == ["1111", "1112", "J052"]
    assert rows[1]["算定項目（D）"] == "(1)日中" and rows[1]["_cells"]["算定項目（D）"] == "D5"  # 結合を画面どおりに広げる
    assert rows[2]["合成／単位数"] == "1,080"  # 表示形式どおり
    assert rows[-1]["_row"] == 9  # 表の後ろの注もそのまま返す
    assert t["sheet_count"] == 1


def test_csv_out_and_errors(tmp_path):
    f = _book(tmp_path)
    out = tmp_path / "t.csv"
    r = api.read_table(f, "合成", out_path=str(out))
    lines = list(csv.reader(io.StringIO(out.read_text(encoding="utf-8"))))
    assert lines[0][:3] == ["シート", "行", "サービスコード／種類"] and len(lines) == r["rows"] + 1
    with pytest.raises(api.ApiError):
        api.read_table(f, "存在しない見出し")
    with pytest.raises(api.ApiError):
        api.read_table(f, "合成", out_path=f)
