"""範囲の所：1 マスずつにしない広い宣言（入力規則・数式の参照・テーブル・配列数式・ODS の繰り返し）が、同じ事実を持つこと。"""

from __future__ import annotations

import io
import json
import tracemalloc
import zipfile

import pytest

from todoku import api, limits
from todoku.xlsx.places import Declared, decompose
from todoku.xlsx.calc import Calc
from todoku.xlsx.window import build
from todoku.xlsx.workbook import Workbook

from test_freeze_small_fixes import HUGE_DV, cells, xlsx  # 窓の試験の、小さなブックの作り方


def far(r: int = 200000) -> str:
    return f'<row r="{r}"><c r="A{r}"><v>1</v></c></row>'


def doc_of(body: str) -> dict:
    return build(Workbook(xlsx(body), "t.xlsx"))


DV = ('<dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="B3:B150000"><formula1>1</formula1></dataValidation></dataValidations>')
SUM_FAR = '<row r="1"><c r="A1"><f>SUM(C5:C150000)</f><v>0</v></c></row>'


def write_case(tmp_path, body: str):
    p = tmp_path / "t.xlsx"
    p.write_bytes(xlsx(body))
    return p


# ------------------------------------------------------------------ 区間の分け方
def rects(ps):
    return [(p.r1, p.c1, p.r2, p.c2, [d.what for d in p.sources]) for p in ps]


def test_decompose_disjoint_and_uniform():
    a = Declared(1, 1, 10, 2, "validation")
    b = Declared(5, 2, 20, 3, "formula_ref")
    ps = decompose([a, b])
    area = sum((p.r2 - p.r1 + 1) * (p.c2 - p.c1 + 1) for p in ps)
    assert area == 20 + 32 - 6  # 重なりは 1 度
    for x in ps:
        for y in ps:
            if x is not y:
                assert x.r2 < y.r1 or y.r2 < x.r1 or x.c2 < y.c1 or y.c2 < x.c1
    both = [p for p in ps if len(p.sources) == 2]
    assert [(p.r1, p.c1, p.r2, p.c2) for p in both] == [(5, 2, 10, 2)]
    # 同じ宣言の組の区間は、つながっていれば 1 つ
    ps = decompose([Declared(1, 1, 10, 5, "validation")])
    assert rects(ps) == [(1, 1, 10, 5, ["validation"])]


def test_decompose_many_declarations_stays_linear():
    ds = [Declared(i, 1, i + 100000, 16384, "formula_ref") for i in range(1, 3000)]
    ps = decompose(ds)
    assert len(ps) == len(ds) and all(p.overlap for p in ps)


# ------------------------------------------------------------------ 窓
def test_validation_place_carries_the_facts():
    doc = doc_of(cells(far()) + DV)
    sh = doc["sheets"][0]
    assert not doc["unread"]
    assert not [f for f in sh["fields"] if f["cell"] == "B4"]  # 1 マスずつにはしない
    p = sh["range_places"][0]
    assert p["id"] == "S!B3:B150000" and p["kind"] == "range"
    assert p["count"] == 150000 - 2 - sum(1 for f in sh["cells"] if f["cell"].startswith("B") and 3 <= int(f["cell"][1:]) <= 150000)
    assert p["why_key"] == ["validation"] and p["rules"][0]["type"] == "whole"
    assert "入力規則" in "\n".join(_md(doc))


def _md(doc):
    from todoku.xlsx.window import _md_sheet

    return _md_sheet(doc["sheets"][0])


def test_small_declarations_are_still_one_field_each():
    small = ('<dataValidations count="1"><dataValidation type="whole" sqref="B3:B6"><formula1>1</formula1></dataValidation></dataValidations>')
    doc = doc_of(cells('<row r="7"><c r="A7" s="1"/></row>') + small)
    assert "range_places" not in doc["sheets"][0]
    assert {"B3", "B4", "B5", "B6"} <= {f["cell"] for f in doc["sheets"][0]["fields"]}


def test_formula_reader_place_has_reader_summary_and_exact_count():
    body = cells(SUM_FAR + '<row r="10"><c r="C10"><v>5</v></c></row>' + far())
    doc = doc_of(body)
    sh = doc["sheets"][0]
    assert not doc["unread"]
    places = [p for p in sh["range_places"] if "formula_ref" in p["why_key"]]
    assert len(places) == 1
    p = places[0]
    assert p["cell"] == "C5:C150000" and p["agent"] is True
    assert p["count"] == 150000 - 4 - 1  # C10 だけファイルに要素がある
    assert p["referenced_by"]["n"] == 1 and p["referenced_by"]["cells"] == ["S!A1"]
    assert any(f["cell"] == "C10" and f.get("referenced_by") for f in sh["fields"])  # 要素のあるマスは 1 マスずつ


def test_overlap_of_validation_and_reader_is_split():
    dv = '<dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="C100000:D150000"><formula1>1</formula1></dataValidation></dataValidations>'
    wide = '<row r="200000"><c r="A200000"><v>1</v></c><c r="E200000"><v>1</v></c></row>'  # 使っている範囲を E 列まで
    doc = doc_of(cells(SUM_FAR + wide) + dv)
    places = doc["sheets"][0]["range_places"]
    both = [p for p in places if {"validation", "formula_ref"} <= set(p["why_key"])]
    assert [p["cell"] for p in both] == ["C100000:C150000"]
    assert both[0].get("rules") and both[0].get("referenced_by")
    assert not both[0].get("agent")  # 入力規則があるので入力欄


def test_hidden_rows_and_merges_are_not_counted():
    body = ('<row r="1"><c r="A1"><f>SUM(C5:C150000)</f><v>0</v></c></row><row r="6" hidden="1"><c r="C6"/></row>'
            '<row r="7"><c r="C7"/></row>' + far())
    merge = '<mergeCells count="1"><mergeCell ref="C100:D101"/></mergeCells>'
    doc = doc_of(cells(body) + merge)
    p = [x for x in doc["sheets"][0]["range_places"] if "formula_ref" in x["why_key"]][0]
    assert p["count"] == 149996 - 1 - 1 - 2  # 非表示の行 6・要素のある C7・結合の C100・C101


def test_no_memory_for_a_full_sheet_declaration():
    dv = '<dataValidations count="1"><dataValidation type="whole" sqref="A1:XFD1048576"><formula1>1</formula1></dataValidation></dataValidations>'
    tracemalloc.start()
    doc = doc_of(cells('<row r="1048576"><c r="XFD1048576"><v>1</v></c></row>') + dv)
    peak = tracemalloc.get_traced_memory()[1]
    tracemalloc.stop()
    assert peak < 64 << 20
    assert doc["sheets"][0]["range_places"][0]["count"] > 10 ** 10 and not doc["unread"]


# ------------------------------------------------------------------ 読む・書く
def test_list_fields_pages_into_the_place(tmp_path):
    p = write_case(tmp_path, cells(far()) + DV)
    whole = api.fields(str(p))
    assert any(f["kind"] == "range" and f["id"] == "S!B3:B150000" for f in whole)
    got = api.fields_range(str(p), {"sheet": "S", "rows": "1000-1002"})
    ids = {f["id"] for f in got["fields"]}
    assert {"S!B1000", "S!B1001", "S!B1002"} <= ids
    f = next(x for x in got["fields"] if x["id"] == "S!B1001")
    assert f["rules"][0]["type"] == "whole" and f["range_place"] == "S!B3:B150000"
    got = api.fields_range(str(p), {"sheet": "S", "rows": "1000-1100"}, limit=10)
    assert got["next"] and "1010" in json.dumps(got["next"])


def test_any_cell_inside_is_writable_and_read_back(tmp_path):
    p = write_case(tmp_path, cells(far()) + DV)
    out = tmp_path / "o.xlsx"
    r = api.fill(str(p), {"S!B77000": 5}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    doc = build(Workbook(out.read_bytes(), "o.xlsx"))
    sh = doc["sheets"][0]
    assert any(f["cell"] == "B77000" and f["value"] == 5 for f in sh["fields"])
    before = build(Workbook(p.read_bytes(), "t.xlsx"))["sheets"][0]["range_places"][0]["count"]
    assert sh["range_places"][0]["count"] == before - 1
    # 入力規則は、書くときにもマスごとに効く
    r = api.fill(str(p), {"S!B77000": "x"}, str(tmp_path / "bad.xlsx"))
    assert any("整数" in w for w in r["warnings"])


def test_doc_with_embedded_sheet_lists_places_and_rows(tmp_path):
    pytest.importorskip("todoku.docx.document")
    # 埋め込みの表（Word）：窓に範囲の所があり、「E1!シート」で行を頼める
    import todoku.embed as em

    win = {"sheets": [{"name": "S", "fields": [], "range_places": [{"id": "S!B3:B9", "cell": "B3:B9", "kind": "range", "count": 7, "why_key": ["validation"]}]}]}
    doc = type("D", (), {"kind": "docx", "zip": None})()
    info_key = em._WRITABLE
    k = None
    em._writable_key = lambda d, e: k
    em.writable = lambda d, e: (True, "")
    info = em.window_info(doc, {"n": 1}, win)
    assert info["range_places"][0]["id"] == "E1!S!B3:B9" and info["range_places"][0]["sheet"] == "S"
    assert info_key is em._WRITABLE


def test_grade_says_nothing_about_ranges_left_out(tmp_path):
    p = write_case(tmp_path, cells(far()) + DV)
    g = api.grade(str(p))
    assert not any(r["title"] in ("広げていない範囲", "読んでいない部分") for r in g["reasons"])
    a = api.audit(str(p))
    assert not any(c["name"] == "読んでいない部分" for c in a["checks"])


def test_markdown_window_shows_the_place(tmp_path):
    p = write_case(tmp_path, cells(far()) + DV)
    text = "\n".join(t for _n, t in api.window(str(p))[2])
    assert "S!B3:B150000（空のセル" in text and "読んでいない部分" not in text


# ------------------------------------------------------------------ 配列数式の範囲
ARR = '<row r="3"><c r="A3"><f t="array" ref="A3:XFD1048576">ROW(1:1048574)*COLUMN(A:XFD)</f><v>1</v></c></row>'


def test_wide_array_formula_is_a_place_and_cells_compute(tmp_path):
    wb = Workbook(xlsx(cells(ARR)), "t.xlsx")
    doc = build(wb)
    sh = doc["sheets"][0]
    assert not doc["unread"]
    p = sh["range_places"][0]
    assert p["why_key"] == ["array"] and p["master"] == "A3" and p["count"] == 1048574 * 16384
    c = Calc(wb)
    assert c.cell_value("S", 1048576, 16384) == 1048574 * 16384
    assert c.cell_value("S", 4, 2) == 4
    path = tmp_path / "a.xlsx"
    path.write_bytes(xlsx(cells(ARR)))
    r = api.calculate(str(path), {}, cells=["S!XFD1048576", "S!C5"])
    assert [x["value"] for x in r["cells"]] == [1048574 * 16384, 9]
    assert r["cells"][0]["spill"] == "A3:XFD1048576"
    assert str(1048574 * 16384) in r["say_to_user"].replace(",", "").replace(".0", "")


def test_aggregate_over_a_wide_array_range(tmp_path):
    body = cells('<row r="1"><c r="A1"><f>COUNT(C3:D1000)</f><v>0</v></c></row>' + ARR.replace("A3", "C3").replace('ref="C3:XFD1048576"', 'ref="C3:XFD1048576"'))
    wb = Workbook(xlsx(body), "t.xlsx")
    assert Calc(wb).cell_value("S", 1, 1) == 998 * 2


# ------------------------------------------------------------------ 長い仕事（集計の続き）
def test_try_calculation_continues_as_a_job(tmp_path, monkeypatch):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    monkeypatch.setattr(limits, "CALC_SECONDS", 0)
    formula = "SUM(MOD(ROW(1:400000),7))"
    body = cells(f'<row r="1"><c r="A1"><f t="array" ref="A1">{formula}</f><v>0</v></c></row>')
    p = write_case(tmp_path, body)
    r = api.calculate(str(p), {}, cells=["S!A1"])
    assert r["cells"][0]["status"] == "unfinished" and 0 < r["cells"][0]["done"] < 400000
    assert r["job"]["running"] and "まだ計算中" in r["say_to_user"]
    import time

    want = float(sum(i % 7 for i in range(1, 400001)))
    for _ in range(600):
        time.sleep(0.2)
        r = api.calculate(str(p), {}, cells=["S!A1"])
        if r["cells"][0]["status"] == "ok":
            break
    assert r["cells"][0]["value"] == want


# ------------------------------------------------------------------ OpenDocument の繰り返し
def _ods(table: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("META-INF/manifest.xml", '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>')
        z.writestr("content.xml", '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
                   'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
                   'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2">'
                   f'<office:body><office:spreadsheet><table:table table:name="S">{table}</table:table></office:spreadsheet></office:body></office:document-content>')
    return buf.getvalue()


def test_ods_repeats_are_a_range_and_every_cell_reads(tmp_path, monkeypatch):
    from todoku.odf import sheet as odf_sheet
    from todoku.odf.sheetwindow import build as ods_build

    monkeypatch.setattr(odf_sheet, "MAX_DECLARED_CELLS", 2000)
    table = ('<table:table-row><table:table-cell table:formula="of:=SUM([.B2:.B1048576])" office:value-type="float" office:value="0"/></table:table-row>'
             '<table:table-row table:number-rows-repeated="1048575"><table:table-cell/><table:table-cell office:value-type="float" office:value="2">'
             '<text:p>2</text:p></table:table-cell></table:table-row>')
    blob = _ods(table)
    wb = odf_sheet.OdsWorkbook(blob, "t.ods")
    assert len(wb.sheets[0].cells) <= 2010
    doc = ods_build(blob, "t.ods")
    sh = doc["sheets"][0]
    assert not doc["unread"]
    rep = [p for p in sh["range_places"] if p["why_key"] == ["repeat"]]
    assert rep and rep[0]["repeats_value"] == 2
    assert not [p for p in sh["range_places"] if p["why_key"] == ["formula_ref"] and p["count"]]  # 繰り返しのマスは空のセルに数えない
    c = Calc(wb)
    assert c.cell_value("S", 1048576, 2) == 2
    assert c.cell_value("S", 900000, 2) == 2
    # どのマスも書ける（行を分けて書き、読み直せる）
    p = tmp_path / "t.ods"
    p.write_bytes(blob)
    out = tmp_path / "o.ods"
    r = api.fill(str(p), {"S!B900000": 7}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    after = odf_sheet.OdsWorkbook(out.read_bytes(), "o.ods")
    assert Calc(after).cell_value("S", 900000, 2) == 7 and Calc(after).cell_value("S", 900001, 2) == 2
