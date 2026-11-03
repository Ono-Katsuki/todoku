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
    # 1回目：見出しは字のある行だけ。下の行が2段目の見出しかは決めず、そのまま返す（エージェントが読んで決める）
    first = api.read_table(f, "合成")
    s0 = first["sheets"][0]
    assert s0["header_rows"] == [3, 3] and s0["rows_below_header"][0]["cells"] == {"A4": "種類", "B4": "項目", "F4": "単位数"}
    assert "header_rows" in first["check"]
    # 2回目：エージェントが「4行目は見出しの続き」と読んで、header_rows=2
    t = api.read_table(f, "合成", header_rows=2)
    assert t["sheets"][0]["header_rule"].startswith("エージェントの指定")
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
    r = api.read_table(f, "合成", header_rows=2, out_path=str(out))
    lines = list(csv.reader(io.StringIO(out.read_text(encoding="utf-8"))))
    assert lines[0][:3] == ["シート", "行", "サービスコード／種類"] and len(lines) == r["rows"] + 1
    with pytest.raises(api.ApiError):
        api.read_table(f, "存在しない見出し")
    with pytest.raises(api.ApiError):
        api.read_table(f, "合成", header_rows=2, out_path=f)


def test_exact_header_cell_wins_over_a_title(tmp_path):
    # 表題「…サービスコード表」の中の字ではなく、字が同じ見出しのセルを選ぶ
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "訪問介護サービスコード表"
    ws["A3"], ws["B3"] = "サービスコード", "単位"
    ws["A4"], ws["B4"] = "1111", 244
    p = tmp_path / "t.xlsx"
    buf = io.BytesIO()
    wb.save(buf)
    p.write_bytes(buf.getvalue())
    t = api.read_table(str(p), "サービスコード")
    assert t["sheets"][0]["header_rows"] == [3, 3] and t["all_rows"][0]["サービスコード"] == "1111"


def test_number_with_text_format_shows_like_excel():
    from todoku.formula.numfmt import format_value

    # 数のセルに「文字列」（@）の表示形式：Excel は標準の表示（「402313@」ではない）
    assert format_value(402313.0, "@") == "402313" and format_value(12.5, "@") == "12.5"
    assert format_value("abc", "@") == "abc" and format_value(12, '"@"0') == "@12"


def test_vertically_merged_header_is_structural(tmp_path):
    # 見出しのセルが縦に結合されていれば、その範囲までは規則で見出し（構造から確実に分かる）
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["B1"], ws["B2"] = "氏名", "連絡先", "電話"
    ws.merge_cells("A1:A2")
    ws["A3"], ws["B3"] = "山田", "03-0000-0000"
    p = tmp_path / "m.xlsx"
    buf = io.BytesIO()
    wb.save(buf)
    p.write_bytes(buf.getvalue())
    t = api.read_table(str(p), "氏名")
    assert t["sheets"][0]["header_rows"] == [1, 2] and t["columns"] == ["氏名", "連絡先／電話"]
    assert t["all_rows"][0] == {"_sheet": "Sheet", "_row": 3, "氏名": "山田", "連絡先／電話": "03-0000-0000", "_cells": {"氏名": "A3", "連絡先／電話": "B3"}}


def test_text_only_data_rows_are_not_taken_as_header(tmp_path):
    # 字だけのデータの行（名簿など）を、見出しと取り違えない
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["氏名", "所属"])
    ws.append(["山田", "総務"])
    ws.append(["佐藤", "経理"])
    p = tmp_path / "n.xlsx"
    buf = io.BytesIO()
    wb.save(buf)
    p.write_bytes(buf.getvalue())
    t = api.read_table(str(p), "氏名")
    assert t["sheets"][0]["header_rows"] == [1, 1] and [r["氏名"] for r in t["all_rows"]] == ["山田", "佐藤"]
