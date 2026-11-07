"""処理の取り出し（extract_processing）：規則のまとめ方・構造・マクロ・書き出した Python のモジュール・PDF のフォーム。"""

from __future__ import annotations

import ast
import base64
import importlib.util
import io
import json
import shutil
import struct
import subprocess
import symtable
import sys
import zipfile
from pathlib import Path

import pytest

from todoku import api

openpyxl = pytest.importorskip("openpyxl")
FIX = Path(__file__).parent / "fixtures"
SRC = str(Path(__file__).parent.parent / "src")


# ---------------------------------------------------------------- 作り方
def raw_xlsx(sheets: dict[str, str], extra: dict[str, bytes | str] | None = None, wb_extra: str = "",
             states: dict[str, str] | None = None, rels_extra: str = "") -> bytes:
    """シートの中身（<sheetData> の中の XML、または <worksheet> 全体）をそのまま書いたブック。キャッシュの値も書いたまま。"""
    buf = io.BytesIO()
    ct = ['<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>']
    rels = ['<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    wb = ['<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>']
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for i, (name, body) in enumerate(sheets.items(), 1):
            st = (states or {}).get(name)
            wb.append(f'<sheet name="{name}" sheetId="{i}" r:id="rId{i}"' + (f' state="{st}"' if st else "") + "/>")
            rels.append(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>')
            ct.append(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
            xml = body if "<worksheet" in body else f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>{body}</sheetData></worksheet>'
            z.writestr(f"xl/worksheets/sheet{i}.xml", '<?xml version="1.0"?>' + xml)
        wb.append("</sheets>" + wb_extra + "</workbook>")
        rels.append(rels_extra + "</Relationships>")
        ct.append("</Types>")
        z.writestr("[Content_Types].xml", "".join(ct))
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", "".join(wb))
        z.writestr("xl/_rels/workbook.xml.rels", "".join(rels))
        for k, v in (extra or {}).items():
            z.writestr(k, v)
    return buf.getvalue()


def put(tmp_path, blob: bytes, name="b.xlsx") -> str:
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


def book(tmp_path, fill, name="b.xlsx", recalc=False) -> str:
    wb = openpyxl.Workbook()
    fill(wb)
    p = tmp_path / name
    wb.save(p)
    return lo(p, tmp_path / "lo") if recalc else str(p)


_LO = shutil.which("soffice")


def lo(path, out_dir) -> str:
    """LibreOffice に計算させて保存する（1 つずつ）。"""
    if not _LO:
        pytest.skip("LibreOffice なし")
    out_dir = Path(out_dir)
    out_dir.mkdir(exist_ok=True)
    r = subprocess.run([_LO, "--headless", f"-env:UserInstallation=file://{out_dir}/profile", "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                        "--outdir", str(out_dir), str(path)], capture_output=True, timeout=180)
    got = out_dir / Path(path).name
    assert got.exists(), r.stderr
    return str(got)


def load_module(path: str):
    spec = importlib.util.spec_from_file_location(f"gen_{abs(hash(path))}", path)
    m = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = m
    spec.loader.exec_module(m)
    return m


def rule(res: dict, text: str) -> dict:
    for d in res["rules"]:
        if d["text"] == text:
            return d
    raise AssertionError(f"{text!r} がありません: {[d['text'] for d in res['rules']]}")


def cached_cells(path: str) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for ws in wb:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out[f"{ws.title}!{c.coordinate}"] = c.value
    return out


# ---------------------------------------------------------------- 規則のまとめ方
def test_same_shape_becomes_one_rule(tmp_path):
    def fill(wb):
        ws = wb.active
        ws.title = "明細"
        ws.append(["単価", "数量", "金額", "累計"])
        for r in range(2, 502):
            ws.cell(r, 1, r)
            ws.cell(r, 2, 3)
            ws.cell(r, 3, f"=A{r}*B{r}")
            ws.cell(r, 4, f"=C{r}+D{r - 1}" if r > 2 else f"=C{r}")

    res = api.extract_processing(book(tmp_path, fill))
    assert res["summary"]["formula_cells"] == 1000
    assert res["summary"]["rules"] == 3
    c = rule(res, "C2:C501 = A{row}*B{row}")
    assert c["formula"] == "=A2*B2" and c["a1"] == "=A{row}*B{row}" and c["r1c1"] == "=RC[-2]*RC[-1]"
    assert c["cells"] == 500
    assert {x["range"] for x in c["reads"]} == {"A2:A501", "B2:B501"}
    d = rule(res, "D3:D501 = C{row}+D{row-1}")
    assert d["r1c1"] == "=RC[-1]+R[-1]C"
    assert res["cycles"] == [] or all(x["kind"] == "ordered" for x in res["cycles"])
    assert res["order"].index(c["id"]) < res["order"].index(d["id"])


def test_two_dimensional_fill_and_absolute_refs(tmp_path):
    def fill(wb):
        ws = wb.active
        for c in range(2, 8):
            ws.cell(1, c, c)
        for r in range(2, 12):
            ws.cell(r, 1, r)
            for c in range(2, 8):
                col = openpyxl.utils.get_column_letter(c)
                ws.cell(r, c, f"=$A{r}*{col}$1")

    res = api.extract_processing(book(tmp_path, fill))
    assert res["summary"]["rules"] == 1
    d = res["rules"][0]
    assert d["range"] == "B2:G11" and d["r1c1"] == "=RC1*R1C"
    assert d["a1"] == "=$A{row}*{col}$1"


def test_shared_and_plain_formulas_of_the_same_shape_merge():
    from todoku.processing.shape import Shaper

    s = Shaper()
    assert s.of_text("B3*C3", 3, 4) == s.of_text("B4*C4", 4, 4) == s.of_text("B5 * C5", 5, 4)
    assert s.of_text("$B$3*C6", 6, 4) != s.of_text("$B$4*C7", 7, 4)  # 絶対参照の行が違う
    assert s.of_text("B3+1", 3, 4) != s.of_text("B3+2", 3, 4)
    assert s.of_text('"3"&B3', 3, 4) != s.of_text('"4"&B4', 4, 4)  # 字の中の数字は動かない
    assert s.of_text("Sheet2!A3", 3, 1) == s.of_text("Sheet2!A4", 4, 1)
    assert s.of_text("A1", 1, 2) == s.of_text("B1", 1, 3)  # 列がずれても同じ形


def test_render_round_trips():
    from todoku.formula.parser import parse
    from todoku.processing.shape import render

    for f in ["A1+B1*C1", "(A1+B1)*C1", "A1-(B1-C1)", "A1/(B1/C1)", "2^3^2", "2^(3^2)", "-A1^2", "-(A1^2)", "A1%*2", "(A1+1)%", '"a""b"&C3', "SUM(A1:B2,C3)",
              "IF(A1>0,,1)", "{1,2;3,4}", "'シート 1'!$A$1:B2", "A:A", "1:3", "T[金額]", "T[[#This Row],[金額]]", "A1:INDEX(B:B,3)", "A1=B1<>C1", "-(-A1)", "+A1", "_xlfn.XLOOKUP(A1,B:B,C:C)"]:
        a = parse(f)
        assert parse(render(a)) == a, f


def test_exceptions_constant_and_other_formula(tmp_path):
    def fill(wb):
        ws = wb.active
        for r in range(1, 31):
            ws.cell(r, 1, r)
            ws.cell(r, 2, f"=A{r}*2")
        ws["B10"] = 999  # 定数で上書き
        ws["B20"] = "=A20+1"  # 違う数式

    res = api.extract_processing(book(tmp_path, fill))
    d = rule(res, "B1:B30 = A{row}*2")
    assert d["cells"] == 28 and d["exception_count"] == 2
    kinds = {x["range"]: x for x in d["exceptions"]}
    assert kinds["B10"]["kind"] == "constant" and kinds["B10"]["value"] == 999
    assert kinds["B20"]["kind"] == "other_formula"
    assert rule(res, "B20 = A20+1")["id"] == kinds["B20"]["rule"]
    assert res["summary"]["rules"] == 2


def test_blank_gap_makes_separate_blocks_with_the_same_shape(tmp_path):
    def fill(wb):
        ws = wb.active
        for r in list(range(1, 6)) + list(range(10, 15)):
            ws.cell(r, 1, r)
            ws.cell(r, 2, f"=A{r}*2")

    res = api.extract_processing(book(tmp_path, fill))
    assert res["summary"]["rules"] == 2 and res["summary"]["distinct_shapes"] == 1
    assert {d["range"] for d in res["rules"]} == {"B1:B5", "B10:B14"}


# ---------------------------------------------------------------- 構造
def test_lookup_tables_are_data_and_retrievable(tmp_path):
    def fill(wb):
        ws = wb.active
        ws.title = "計算"
        ws["A1"] = "区分"
        ws["A2"] = "B"
        ws["B2"] = '=VLOOKUP(A2,単価!$A$2:$B$4,2,FALSE)'
        ws["C2"] = '=INDEX(単価!$B$2:$B$4,MATCH(A2,単価!$A$2:$A$4,0))'
        t = wb.create_sheet("単価")
        t.append(["区分", "単価"])
        for k, v in (("A", 100), ("B", 200), ("C", 300)):
            t.append([k, v])

    path = book(tmp_path, fill, recalc=True)
    res = api.extract_processing(path)
    tables = {x["range"]: x for x in res["lookup_tables"]}
    assert tables["A2:B4"]["functions"] == ["VLOOKUP"]
    assert tables["A2:B4"]["values"] == [["A", 100], ["B", 200], ["C", 300]]
    assert tables["B2:B4"]["functions"] == ["INDEX"] and tables["A2:A4"]["functions"] == ["MATCH"]
    hint = tables["A2:B4"]["read_table"]
    got = api.read_table(path, hint["header"], sheet=hint["sheet"])
    assert got["rows"] >= 3
    # 参照表は入力の一覧に混ぜない
    assert not any(d["sheet"] == "単価" for d in res["inputs"])
    assert any(d["sheet"] == "計算" and d["range"] == "A2" for d in res["inputs"])


def test_inputs_outputs_parameters_labels_constraints(tmp_path):
    def fill(wb):
        ws = wb.active
        ws.title = "様式"
        ws["A1"] = "税率"
        ws["B1"] = 0.1
        ws["A3"] = "単価"
        ws["B3"] = 100
        ws["A4"] = "数量"
        ws["A5"] = "小計"
        ws["B5"] = "=B3*B4"
        ws["A6"] = "税込"
        ws["B6"] = "=ROUND(B5*(1+$B$1),0)"
        dv = openpyxl.worksheet.datavalidation.DataValidation(type="whole", operator="between", formula1="1", formula2="9", showErrorMessage=True)
        dv.add("B4")
        ws.add_data_validation(dv)
        ws.conditional_formatting.add("B6", openpyxl.formatting.rule.FormulaRule(formula=["B6>1000"], fill=openpyxl.styles.PatternFill(bgColor="FF0000")))

    res = api.extract_processing(book(tmp_path, fill, recalc=True))
    params = {d["range"]: d for d in res["parameters"]}
    assert params["B1"]["value"] == 0.1 and params["B1"]["labels"]["left"]["text"] == "税率"
    ins = {d["range"]: d for d in res["inputs"]}
    assert ins["B3"]["value"] == 100 and ins["B3"]["labels"]["left"]["text"] == "単価"
    assert ins["B4"]["filled"] == 0 and ins["B4"]["constraints"][0]["type"] == "whole"
    assert rule(res, "B6 = ROUND(B5*(1+$B$1),0)")["outputs"] == ["B6"]
    assert "outputs" not in rule(res, "B5 = B3*B4")
    assert res["conditional_formats"][0]["kind"] == "conditional_format" and res["conditional_formats"][0]["formulas"] == ["B6>1000"]
    assert res["data_validations"][0]["kind"] == "data_validation" and res["data_validations"][0]["formula1"] == "1"


def test_branches_are_shown_as_conditions(tmp_path):
    def fill(wb):
        ws = wb.active
        ws["A1"] = 5
        ws["B1"] = '=IF(A1>3,IFS(A1>9,"大",TRUE,"中"),"小")'
        ws["B2"] = '=ROUND(IF(A1>3,A1,0),0)'
        ws["B3"] = '=SWITCH(A1,1,"一",5,"五","他")'

    res = api.extract_processing(book(tmp_path, fill))
    b = rule(res, 'B1 = IF(A1>3,IFS(A1>9,"大",TRUE,"中"),"小")')["branches"]
    assert b["kind"] == "if" and b["condition"] == "A1>3" and b["else"] == '"小"'
    assert b["then"]["kind"] == "ifs" and b["then"]["cases"][0] == {"condition": "A1>9", "then": '"大"'}
    assert rule(res, "B2 = ROUND(IF(A1>3,A1,0),0)")["branches"]["contains"][0]["kind"] == "if"
    sw = next(d for d in res["rules"] if d["range"] == "B3")["branches"]
    assert sw["kind"] == "switch" and sw["default"] == '"他"' and sw["cases"][1] == {"value": "5", "then": '"五"'}


def test_cycles(tmp_path):
    def fill(wb):
        ws = wb.active
        ws["A1"] = "=B1+1"
        ws["B1"] = "=A1+1"
        # 互いに読むが、セルの順には計算できる
        ws["D1"] = 1
        ws["E1"] = 1
        for r in range(2, 20):
            ws.cell(r, 4, f"=E{r - 1}+1")
            ws.cell(r, 5, f"=D{r}+1")

    res = api.extract_processing(book(tmp_path, fill))
    kinds = {c["kind"] for c in res["cycles"]}
    assert "cell_cycle" in kinds and "ordered" in kinds
    bad = next(c for c in res["cycles"] if c["kind"] == "cell_cycle")
    assert set(bad["cells"]) == {"Sheet!A1", "Sheet!B1"}
    assert rule(res, "A1 = B1+1")["cycle"] == "cell_cycle"
    # 輪のないものは順番に並ぶ
    good = next(c for c in res["cycles"] if c["kind"] == "ordered")
    assert len(good["rules"]) == 2


def test_volatile_dynamic_external_hidden_names_crosssheet():
    sheet1 = (
        '<row r="1"><c r="A1"><v>2</v></c>'
        '<c r="B1"><f>TODAY()</f><v>46000</v></c>'
        '<c r="C1"><f>OFFSET(A1,0,0,COUNTA(A:A),1)</f><v>2</v></c>'
        '<c r="D1"><f>INDIRECT("隠し!A1")</f><v>7</v></c>'
        '<c r="E1"><f>INDIRECT(A1&amp;"!A1")</f><v>0</v></c>'
        '<c r="F1"><f>[1]外!A1*2</f><v>6</v></c>'
        '<c r="G1"><f>税率*A1</f><v>0.2</v></c>'
        '<c r="H1"><f>B1+1</f><v>46001</v></c></row>'
        '<row r="2" hidden="1"><c r="A2"><v>3</v></c><c r="I2"><f>A2*1</f><v>3</v></c></row>'
    )
    hidden = '<row r="1"><c r="A1"><v>7</v></c></row>'
    ext = ('<?xml version="1.0"?><externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><externalBook r:id="rId1"><sheetNames><sheetName val="外"/></sheetNames>'
           '<sheetDataSet><sheetData sheetId="0"><row r="1"><cell r="A1"><v>3</v></cell></row></sheetData></sheetDataSet></externalBook></externalLink>')
    extra = {"xl/externalLinks/externalLink1.xml": ext,
             "xl/externalLinks/_rels/externalLink1.xml.rels":
             '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="file:///D:/外部.xlsx" TargetMode="External"/></Relationships>'}
    wb_extra = '<externalReferences><externalReference r:id="rId9"/></externalReferences><definedNames><definedName name="税率">隠し!$A$2</definedName></definedNames>'
    rels = '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>'
    import tempfile

    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "v.xlsx"
        p.write_bytes(raw_xlsx({"本": sheet1, "隠し": hidden}, extra, wb_extra, {"隠し": "hidden"}, rels))
        res = api.extract_processing(str(p))
    by = {d["range"]: d for d in res["rules"]}
    assert by["B1"]["volatile"] == ["TODAY"]
    assert by["H1"]["depends_on_volatile"] is True
    off = by["C1"]["dynamic_refs"][0]
    assert off["function"] == "OFFSET" and off["unresolved"] == {"height": "COUNTA(A:A)"} and off["base"] == "A1"
    assert by["D1"]["dynamic_refs"][0]["resolved"] == "隠し!A1"
    assert any(x.get("range") == "A1" and x["sheet"] == "隠し" for x in by["D1"]["reads"])  # 字で決まる INDIRECT は静的に読む
    assert "unresolved" in by["E1"]["dynamic_refs"][0]
    e = res["external_references"][0]
    assert e["rule"] == by["F1"]["id"] and e["book"] == "file:///D:/外部.xlsx"
    names = {d["name"]: d for d in res["named_ranges"]}
    assert names["税率"]["used_by"] == [by["G1"]["id"]]
    assert any(x.get("name") == "税率" for x in by["G1"]["reads"])
    hid = res["hidden_used"]
    assert any(h["what"] == "sheet" and h["sheet"] == "隠し" and by["D1"]["id"] in h["used_by"] for h in hid)
    assert any(h["what"] == "rows" and h["rows"] == "2:2" for h in hid)
    flow = {(f["from"], f["to"]) for f in res["cross_sheet_flow"]}
    assert ("隠し", "本") in flow


def test_arrays_are_expressions_never_expanded():
    body = ('<row r="1"><c r="A1"><v>1</v></c><c r="C1"><f t="array" ref="C1:C3">A1:A3*2</f><v>2</v></c>'
            '<c r="E1" cm="1"><f t="array" ref="E1:E3">_xlfn.SEQUENCE(3)</f><v>1</v></c></row>'
            '<row r="2"><c r="A2"><v>2</v></c><c r="C2"><v>4</v></c><c r="E2"><v>2</v></c></row>'
            '<row r="3"><c r="A3"><v>3</v></c><c r="C3"><v>6</v></c><c r="E3"><v>3</v></c></row>')
    import tempfile

    with tempfile.TemporaryDirectory() as td:
        p = Path(td) / "a.xlsx"
        p.write_bytes(raw_xlsx({"S": body}))
        res = api.extract_processing(str(p))
        mod = Path(td) / "m.py"
        r2 = api.extract_processing(str(p), format="python", out_path=str(mod))
        assert r2["verification"]["stored"]["rate"] == "100%", r2["verification"]
        m = load_module(str(mod))
        got = m.calculate({"S!A1": 10}, all_cells=True)
        assert got["S!C1"] == 20 and got["S!C2"] == 4 and got["S!E3"] == 3
    a = rule(res, "C1:C3 = {=A1:A3*2}")
    assert a["kind"] == "array" and a["array_ref"] == "C1:C3" and a["cells"] == 3
    d = next(x for x in res["rules"] if x["range"] == "E1:E3")
    assert d["kind"] == "dynamic_array"
    assert res["summary"]["rules"] == 2
    assert not any(x["range"] == "C2" for x in res["inputs"])  # こぼれた値は定数ではない


def test_pivot_power_query_connections_are_listed_as_definitions(tmp_path):
    inner = io.BytesIO()
    with zipfile.ZipFile(inner, "w") as z:
        z.writestr("Formulas/Section1.m", 'section Section1;\nshared 集計 = let Source = Excel.CurrentWorkbook() in Source;')
    pkg = inner.getvalue()
    mash = struct.pack("<II", 0, len(pkg)) + pkg + b"\0" * 16
    dm = ('<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">' + base64.b64encode(mash).decode() + "</DataMashup>").encode("utf-16")
    sheet = ('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
             '<row r="1"><c r="A1" t="str"><v>x</v></c></row></sheetData></worksheet>')
    extra = {
        "xl/worksheets/_rels/sheet1.xml.rels": '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>',
        "xl/pivotTables/pivotTable1.xml": '<?xml version="1.0"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="P1" cacheId="1">'
        '<location ref="D3:F9"/><rowFields><field x="0"/></rowFields><dataFields><dataField name="合計" fld="1" subtotal="sum"/></dataFields></pivotTableDefinition>',
        "xl/pivotTables/_rels/pivotTable1.xml.rels": '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>',
        "xl/pivotCache/pivotCacheDefinition1.xml": '<?xml version="1.0"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<cacheSource type="worksheet"><worksheetSource ref="A1:B9" sheet="元"/></cacheSource><cacheFields><cacheField name="区分"/><cacheField name="金額"/></cacheFields></pivotCacheDefinition>',
        "customXml/item1.xml": dm,
        "xl/connections.xml": '<?xml version="1.0"?><connections xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<connection id="1" name="c" type="5"><dbPr connection="Password=secret" command="select 1"/></connection></connections>',
    }
    res = api.extract_processing(put(tmp_path, raw_xlsx({"S": sheet}, extra)))
    items = {d["kind"]: d for d in res["not_expressible_by_cells"]}
    p = items["pivot_table"]
    assert p["source"] == {"sheet": "元", "ref": "A1:B9"} and p["location"] == "D3:F9" and p["row_fields"] == ["区分"]
    assert p["data_fields"] == [{"name": "合計", "field": "金額", "function": "sum"}] and p["sheet"] == "S"
    assert p["expressible_by_cells"] is False
    assert "Excel.CurrentWorkbook" in items["power_query"]["m_source"]["Formulas/Section1.m"]
    c = items["data_connection"]
    assert c["command"] == "select 1" and "secret" not in json.dumps(res, ensure_ascii=False)


# ---------------------------------------------------------------- マクロ
def _macro_book():
    sys.path.insert(0, str(Path(__file__).parent))
    from test_vbarun import book as vba_book
    from todoku.xlsx.workbook import Control

    src = '''
Sub 集計()
    Dim ws As Worksheet
    Set ws = Sheets("入力")
    ws.Range("C1").Value = ws.Range("A1").Value * 2
    Range("D1").ClearContents
    With Worksheets("入力")
        .Cells(2, 3).Value = .Cells(2, 1).Value
    End With
    Dim i As Long
    For i = 1 To 3
        Cells(i, 5).Value = i
    Next i
    Range(Sheets("入力").Range("A1").Value).Value = 1
    Sheets("入力").Range("B1").Formula = "=A1*3"
    手伝い
End Sub

Sub 手伝い()
    Range("A5").Copy Destination:=Range("A6")
End Sub
'''
    sheet = '''
Private Sub Worksheet_Change(ByVal Target As Range)
    If Not Intersect(Target, Range("A1:A3")) Is Nothing Then
        Range("F1").Value = Range("A1").Value + 1
    End If
End Sub

Private Sub CommandButton1_Click()
    Range("G1").Value = 1
End Sub
'''
    wb = vba_book(src, {"A1": 1, "A2": 2, "B1": {"f": "A1*3"} if False else 5}, sheet)
    wb.sheets[0].controls.append(Control("Button", "ボタン1", "1", text="集計する", macro="[0]!Module1.集計"))
    return wb


def test_macros_static_reads_writes_triggers():
    from todoku.processing import macrostatic, rules as R, scan

    wb = _macro_book()
    src = scan.BookSource(wb)
    sheets, shaper = scan.scan_book(src)
    rules = R.build_blocks(sheets, shaper)
    R.analyse_reads(rules, sheets, wb)
    ix = R.link_rules(rules)
    m = macrostatic.analyse(wb, rules, ix, sheets)
    procs = {p["name"]: p for p in m["procedures"]}
    s = procs["集計"]
    writes = {(w["sheet"], w["cell"]) for w in s["writes"]}
    assert ("入力", "C1") in writes and ("入力", "C2") in writes and ("入力", "B1") in writes
    assert any(w["cell"] == "D1" and w.get("how") == "clearcontents" for w in s["writes"])
    assert any(w["cell"] == "E?" for w in s["writes"])  # 列だけ決まる
    assert ("入力", "A1") in {(r["sheet"], r["cell"]) for r in s["reads"]}
    assert any(w.get("cell_unknown") for w in s["writes"])  # Range(変数) は分からないと言う
    assert s["writes_formulas"][0]["formula"] == "=A1*3"
    assert "手伝い" in s["calls"]
    assert s["trigger"]["buttons"][0]["text"] == "集計する"
    assert "Sub 集計()" in s["source"] and s["source"].rstrip().endswith("End Sub")
    h = procs["手伝い"]
    assert any(w["cell"] == "A6" and w.get("how") == "copy destination" for w in h["writes"]) and any(r["cell"] == "A5" for r in h["reads"])
    c = procs["Worksheet_Change"]
    assert c["trigger"]["event"] == "Worksheet_Change" and c["trigger"]["watches"] == ["A1:A3"] and c["trigger"]["sheet"] == "入力"
    assert ("入力", "F1") in {(w["sheet"], w["cell"]) for w in c["writes"]}
    assert "ボタン" in procs["CommandButton1_Click"]["trigger"]["when"]
    assert set(m["triggered"]) >= {"Module1.集計", "Sheet1.Worksheet_Change"}


def test_macros_relate_to_rules():
    from todoku.processing import macrostatic, rules as R, scan

    sys.path.insert(0, str(Path(__file__).parent))
    from test_vbarun import book as vba_book
    from todoku.formula.parser import parse

    wb = vba_book('''
Sub T()
    Range("A1").Value = 3
    Range("C1").Value = 1
    Range("E1").Value = Range("B1").Value
End Sub
''', {"A1": 1})
    sh = wb.sheets[0]
    from todoku.xlsx.workbook import Cell

    c = Cell("入力", 1, 2, value=2.0)
    c.formula = "A1*2"
    c.ast = parse("A1*2")
    sh.cells[(1, 2)] = c
    c2 = Cell("入力", 1, 3, value=4.0)
    c2.formula = "B1*3"
    c2.ast = parse("B1*3")
    sh.cells[(1, 3)] = c2
    sh.max_col = 3
    src = scan.BookSource(wb)
    sheets, shaper = scan.scan_book(src)
    rules = R.build_blocks(sheets, shaper)
    R.analyse_reads(rules, sheets, wb)
    ix = R.link_rules(rules)
    m = macrostatic.analyse(wb, rules, ix, sheets)
    rel = next(p for p in m["procedures"] if p["name"] == "T")["rules"]
    by = {u.text: u.id for u in rules}
    assert by["A1*2"] in rel["feeds_input_of"]
    assert rel["overwrites_output_of"] == [by["B1*3"]]
    assert rel["reads_output_of"] == [by["A1*2"]]


# ---------------------------------------------------------------- 大きなブック
def _huge(path: Path, rows: int) -> None:
    """rows 行：A・B は値、C は共有数式、D はふつうの数式（1 行だけ違う形と定数の上書き）、最後に合計。"""
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=1) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                   'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="明細" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        with z.open("xl/worksheets/sheet1.xml", "w", force_zip64=True) as f:
            f.write(b'<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
            total = 0
            for r in range(1, rows + 1):
                a, b = r % 97 + 1, r % 13 + 1
                total += a * b
                c = (f'<f t="shared" ref="C1:C{rows}" si="0">A1*B1</f>' if r == 1 else '<f t="shared" si="0"/>') + f"<v>{a * b}</v>"
                if r == rows // 2:
                    d = "<v>7</v>"  # 定数で上書き
                elif r == rows // 3:
                    d = f"<f>A{r}+B{r}</f><v>{a + b}</v>"
                else:
                    d = f"<f>C{r}+A{r}</f><v>{a * b + a}</v>"
                f.write(f'<row r="{r}"><c r="A{r}"><v>{a}</v></c><c r="B{r}"><v>{b}</v></c><c r="C{r}">{c}</c><c r="D{r}">{d}</c></row>'.encode())
            r = rows + 1
            f.write(f'<row r="{r}"><c r="C{r}"><f>SUM(C1:C{rows})</f><v>{total}</v></c></row>'.encode())
            f.write(b"</sheetData></worksheet>")


_PEAK = """
import json, resource, sys
from todoku import api
res = api.extract_processing(sys.argv[1])
peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
peak = peak if sys.platform == "darwin" else peak * 1024
print(json.dumps({"peak": peak, "summary": res["summary"], "texts": [d["text"] for d in res["rules"]], "exc": [d.get("exception_count", 0) for d in res["rules"]]}, ensure_ascii=False))
"""


def _peak(path: Path) -> dict:
    p = subprocess.run([sys.executable, "-c", _PEAK, str(path)], capture_output=True, text=True, env={"PYTHONPATH": SRC, "PATH": "/usr/bin:/bin"}, timeout=300)
    assert p.returncode == 0, p.stderr[-2000:]
    return json.loads(p.stdout)


@pytest.mark.skipif(sys.platform == "win32", reason="resource（ピークのメモリ）は POSIX だけ")
def test_100k_rows_give_a_handful_of_rules_with_bounded_memory(tmp_path):
    small, big = tmp_path / "s.xlsx", tmp_path / "b.xlsx"
    _huge(small, 10_000)
    _huge(big, 100_000)
    a, b = _peak(small), _peak(big)
    assert b["summary"]["formula_cells"] == 100_000 + 100_000 - 1 + 1
    assert b["summary"]["rules"] == 4
    assert "C1:C100000 = A{row}*B{row}" in b["texts"]
    assert any(t.startswith("D1:D100000 = C{row}+A{row}") for t in b["texts"])
    assert sum(b["exc"]) == 2
    # セルが 10 倍でも、メモリはセルの数に比例しない（セルごとの控えなら数十 MB 増える）
    assert b["peak"] - a["peak"] < 24 * 1024 * 1024, (a["peak"], b["peak"])
    if sys.platform == "darwin":  # 全体のピークは C ライブラリの割り当てで変わる（Linux の glibc では大きく出る）。増え方は上で確かめる
        assert b["peak"] < 160 * 1024 * 1024


def test_100k_rows_python_module_verifies(tmp_path):
    p = tmp_path / "b.xlsx"
    _huge(p, 100_000)
    mod = tmp_path / "m.py"
    r = api.extract_processing(str(p), format="python", out_path=str(mod))
    v = r["verification"]
    assert v["ran"] and v["stored"]["compared"] == 200_000 - 0 and v["stored"]["rate"] == "100%", v
    assert v["evaluator"]["rate"] == "100%"
    assert v["rules_not_proven"] == 0
    assert mod.stat().st_size < 3_000_000  # 値は圧縮して 1 つのファイルに


# ---------------------------------------------------------------- 書き出した Python のモジュール
def _rich_fill(wb):
    from datetime import date

    ws = wb.active
    ws.title = "入力"
    ws["A1"] = "単価"
    ws["B1"] = 1234.5
    ws["A2"] = "数量"
    ws["B2"] = 3
    ws["A3"] = "率"
    ws["B3"] = 0.075
    ws["A4"] = "区分"
    ws["B4"] = "b"
    ws["A5"] = "日付"
    ws["B5"] = date(2024, 2, 29)
    ws["A6"] = "字"
    ws["B6"] = "  Todoku  Tool "
    t = wb.create_sheet("表")
    t.append(["区分", "係数", "名前"])
    for k, v, n in (("a", 1.5, "甲"), ("b", 2.25, "乙"), ("c", 3, "丙")):
        t.append([k, v, n])
    wb.defined_names["係数表"] = openpyxl.workbook.defined_name.DefinedName("係数表", attr_text="表!$A$2:$C$4")
    c = wb.create_sheet("計算")
    formulas = {
        "A1": "=入力!B1*入力!B2", "A2": "=ROUND(A1*入力!B3,0)", "A3": "=ROUNDDOWN(A1*入力!B3,1)", "A4": "=ROUNDUP(A1*入力!B3,-1)",
        "A5": "=VLOOKUP(入力!B4,係数表,2,FALSE)", "A6": "=A5*A1", "A7": "=INDEX(表!$C$2:$C$4,MATCH(入力!B4,表!$A$2:$A$4,0))",
        "A8": '=IF(A1>3000,"大",IF(A1>1000,"中","小"))', "A9": "=SUM(表!B2:B4)", "A10": "=AVERAGE(表!B2:B4)", "A11": "=MAX(A1:A4)-MIN(A1:A4)",
        "A12": '=COUNTIF(表!A2:A4,"b")', "A13": '=SUMIF(表!A2:A4,"<>a",表!B2:B4)', "A14": "=SUMPRODUCT(表!B2:B4,表!B2:B4)",
        "B1": "=YEAR(入力!B5)", "B2": "=MONTH(入力!B5)", "B3": "=DAY(入力!B5)", "B4": "=EOMONTH(入力!B5,1)", "B5": "=EDATE(入力!B5,12)",
        "B6": "=WEEKDAY(入力!B5)", "B7": "=DATE(2024,14,1)", "B8": '=DATEDIF(DATE(2020,1,15),入力!B5,"M")', "B9": '=TEXT(入力!B5,"yyyy/mm/dd")',
        "B10": '=TEXT(A1,"#,##0.00")', "C1": "=TRIM(入力!B6)", "C2": "=UPPER(C1)", "C3": "=LEN(入力!B6)", "C4": '=SUBSTITUTE(C1," ","-")', "C5": "=LEFT(C1,3)&RIGHT(C1,2)",
        "C6": '=MID(C1,2,3)', "C7": '=FIND("o",C1)', "C8": '=C1&"!"&A5', "D1": "=1/0", "D2": "=IFERROR(D1,-1)", "D3": "=ISERROR(D1)", "D4": "=NA()", "D5": "=_xlfn.IFNA(D4,0)",
        "D6": '=VLOOKUP("zzz",係数表,2,FALSE)', "D7": '="a"+1', "D9": "=MOD(-7,3)", "D10": "=INT(-2.5)", "D11": "=ROUND(2.5,0)", "D12": "=ROUND(-2.5,0)",
        "D13": "=ROUND(1.005,2)", "D14": "=0.1+0.2", "D15": "=CEILING(2.1,1)", "D16": "=FLOOR(2.9,1)", "D17": "=TRUE+1", "D18": "=AND(A1>0,A2>0)", "D19": "=OR(A1<0,A2<0)", "D20": "=NOT(D19)",
        "D21": "=COUNTA(表!A1:C4)", "D22": "=COUNT(表!A1:C4)", "D23": "=ABS(-3)", "D24": "=POWER(2,10)", "D25": "=2^0.5",
        "E1": "=SUM(入力!$B$1:$B$3)", "E2": "=ROW()", "E3": "=COLUMN()", "E4": "=ROW(A10)", "E5": "=CHOOSE(2,10,20,30)", "E6": "=SUM(A1:A2)*入力!$B$2",
    }
    for k, v in formulas.items():
        c[k] = v
    for r in range(20, 40):
        c.cell(r, 1, r)
        c.cell(r, 2, f"=A{r}*入力!$B$2")
        c.cell(r, 3, f"=SUM(B$20:B{r})")


@pytest.mark.skipif(not _LO, reason="LibreOffice なし")
def test_generated_module_equals_libreoffice_recalculation(tmp_path):
    base = tmp_path / "base.xlsx"
    wb = openpyxl.Workbook()
    _rich_fill(wb)
    wb.save(base)
    calc1 = lo(base, tmp_path / "lo1")
    mod = tmp_path / "m.py"
    res = api.extract_processing(calc1, format="python", out_path=str(mod))
    v = res["verification"]
    assert v["ran"] and v["stored"]["rate"] == "100%", json.dumps(v["mismatches"], ensure_ascii=False)
    assert v["evaluator"]["rate"] == "100%", json.dumps(v["mismatches"], ensure_ascii=False)
    # 入力を変えて、LibreOffice が計算し直した値と比べる
    m = load_module(str(mod))
    changes = {"B1": 2000.75, "B2": 7, "B3": 0.125, "B4": "c", "B6": "  ab  cd  "}
    wb2 = openpyxl.load_workbook(base)
    for k, x in changes.items():
        wb2["入力"][k] = x
    other = tmp_path / "changed.xlsx"
    wb2.save(other)
    calc2 = lo(other, tmp_path / "lo2")
    want = cached_cells(calc2)
    got = m.calculate({f"入力!{k}": x for k, x in changes.items()}, all_cells=True)
    assert len(got) > 80
    bad = []
    for k, x in got.items():
        y = want.get(k)
        if isinstance(x, str) and isinstance(y, str):
            ok = str(x) == y
        elif isinstance(x, m.CellError):
            ok = y == str(x)
        elif isinstance(y, (int, float)) and not isinstance(y, bool) and isinstance(x, (int, float)):
            ok = abs(float(x) - float(y)) <= 1e-9 * max(1.0, abs(float(y)))
        elif hasattr(y, "year"):
            ok = m.excel_date(x).date() == (y.date() if hasattr(y, "date") else y)
        else:
            ok = x == y or (x in (None, "") and y is None)
        if not ok:
            bad.append((k, x, y))
    assert not bad, bad[:20]


def test_generated_module_is_standalone(tmp_path):
    wb = openpyxl.Workbook()
    _rich_fill(wb)
    p = tmp_path / "b.xlsx"
    wb.save(p)
    mod = tmp_path / "m.py"
    res = api.extract_processing(str(p), format="python", out_path=str(mod))
    assert res["python"]["out_path"] == str(mod)
    src = mod.read_text(encoding="utf-8")
    assert "@@" not in src.replace("@@todoku-verification@@", "")  # 置き換え忘れがない
    roots = set()
    for n in ast.walk(ast.parse(src)):
        if isinstance(n, ast.Import):
            roots |= {a.name.split(".")[0] for a in n.names}
        elif isinstance(n, ast.ImportFrom):
            assert n.level == 0
            roots.add(n.module.split(".")[0])
    assert roots <= set(sys.stdlib_module_names), roots
    r = subprocess.run([sys.executable, "-I", "-S", str(mod)], capture_output=True, text=True, cwd=tmp_path, timeout=60)
    assert r.returncode == 0, r.stderr
    assert isinstance(json.loads(r.stdout), dict)
    # すでにあるファイルは、頼まれないと置き換えない
    with pytest.raises(api.ApiError):
        api.extract_processing(str(p), format="python", out_path=str(mod))
    api.extract_processing(str(p), format="python", out_path=str(mod), overwrite=True)


def test_inputs_are_checked_and_dates_become_serials(tmp_path):
    from datetime import date

    def fill(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = date(2024, 1, 31)
        ws["B1"] = "=EDATE(A1,1)"
        ws["C1"] = 5
        ws["D1"] = "未使用"

    path = book(tmp_path, fill, recalc=True)
    mod = tmp_path / "m.py"
    api.extract_processing(path, format="python", out_path=str(mod))
    m = load_module(str(mod))
    got = m.calculate({"S!A1": date(2023, 12, 31)})
    assert m.excel_date(got["S!B1"]).date() == date(2024, 1, 31)
    with pytest.raises(ValueError, match="数式のセル"):
        m.calculate({"S!B1": 1})
    with pytest.raises(ValueError, match="どの規則も読まない"):
        m.calculate({"S!D1": 1})
    with pytest.raises(KeyError):
        m.calculate({"無い!A1": 1})
    assert m.calculate({"S!B1": 3}, allow_override=True, outputs=["S!B1"]) == {"S!B1": 3.0}


def test_unknown_function_raises_for_those_outputs_and_says_why(tmp_path):
    body = ('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>WEBSERVICE("http://x")</f><v>9</v></c><c r="C1"><f>B1+1</f><v>10</v></c>'
            '<c r="D1"><f>A1*2</f><v>4</v></c><c r="E1"><f>IF(A1&gt;0,1,WEBSERVICE("y"))</f><v>1</v></c></row>')
    path = put(tmp_path, raw_xlsx({"S": body}))
    mod = tmp_path / "m.py"
    res = api.extract_processing(path, format="python", out_path=str(mod))
    assert "WEBSERVICE" in res["unknown_functions"]
    by = {d["range"]: d for d in res["rules"]}
    assert by["B1"]["verification"]["status"] == "not_proven" and "WEBSERVICE" in by["B1"]["verification"]["why"][0]
    assert by["C1"]["verification"]["status"] == "not_proven"  # 頼っている規則も
    assert by["D1"]["verification"]["status"] == "proven"
    assert by["E1"]["verification"]["status"] == "proven"  # 通らない枝の関数は計算に要らない
    assert res["verification"]["mismatch_count"] == 0
    m = load_module(str(mod))
    with pytest.raises(m.NotReproducible) as e:
        m.calculate()
    assert set(e.value.cells) == {"S!C1"} and "WEBSERVICE" in str(e.value)
    assert m.calculate(outputs=["S!D1"]) == {"S!D1": 4.0}
    assert m.calculate({"S!A1": 5}, on_unknown="mark", all_cells=True)["S!D1"] == 10
    assert m.calculate(on_unknown="mark")["S!C1"] == "#UNKNOWN!"
    assert m.calculate({"S!A1": -1}, outputs=["S!D1"]) == {"S!D1": -2.0}
    with pytest.raises(m.NotReproducible):
        m.calculate({"S!A1": -1}, outputs=["S!E1"])
    assert m.UNREPRODUCIBLE[by["B1"]["id"]].endswith("WEBSERVICE")


def test_single_cell_array_formulas_merge_into_one_rule(tmp_path):
    rows = "".join(f'<row r="{r}"><c r="A{r}"><v>{r}</v></c><c r="B{r}"><f t="array" ref="B{r}">SUM(A$1:A{r}*2)</f><v>{r * (r + 1)}</v></c></row>' for r in range(1, 6))
    path = put(tmp_path, raw_xlsx({"S": rows}))
    mod = tmp_path / "m.py"
    res = api.extract_processing(path, format="python", out_path=str(mod))
    assert res["summary"]["rules"] == 1
    d = res["rules"][0]
    assert d["kind"] == "array_cell" and d["text"] == "B1:B5 = {=SUM(A$1:A{row}*2)}"
    assert res["verification"]["stored"]["rate"] == "100%" and res["verification"]["evaluator"]["rate"] == "100%"
    assert load_module(str(mod)).calculate({"S!A1": 11}, outputs=["S!B5"]) == {"S!B5": 50.0}


_SLOW = """
import json, os, sys, time
from todoku import api
t0 = time.time()
r = api.extract_processing(sys.argv[1], format="python", out_path=sys.argv[2])
print(json.dumps({"secs": time.time() - t0, "notes": r["verification"]["notes"], "status": [d["verification"]["status"] for d in r["rules"]], "why": [d["verification"].get("why") for d in r["rules"]]}, ensure_ascii=False))
"""


def test_time_limit_stops_verification_and_says_so(tmp_path):
    rows = 60000
    with zipfile.ZipFile(tmp_path / "t.xlsx", "w", zipfile.ZIP_DEFLATED, compresslevel=1) as z:
        src = zipfile.ZipFile(io.BytesIO(raw_xlsx({"S": ""})))
        for n in src.namelist():
            if not n.startswith("xl/worksheets"):
                z.writestr(n, src.read(n))
        with z.open("xl/worksheets/sheet1.xml", "w", force_zip64=True) as f:
            f.write(b'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')
            total = 0
            for r in range(1, rows + 1):
                total += 1
                f.write(f'<row r="{r}"><c r="A{r}"><v>1</v></c><c r="B{r}"><f>SUM(A$1:A{r})</f><v>{total}</v></c></row>'.encode())
            f.write(b"</sheetData></worksheet>")
    p = subprocess.run([sys.executable, "-c", _SLOW, str(tmp_path / "t.xlsx"), str(tmp_path / "m.py")], capture_output=True, text=True, timeout=120,
                       env={"PYTHONPATH": SRC, "PATH": "/usr/bin:/bin", "TODOKU_TIME_LIMIT": "6"})
    assert p.returncode == 0, p.stderr[-1500:]
    d = json.loads(p.stdout.strip().splitlines()[-1])
    assert (tmp_path / "m.py").exists()
    if d["status"] != ["proven"]:  # 速いパソコンでは全部計算し終わることもある
        assert d["status"] == ["not_proven"] and "比べていない" in d["why"][0][0]
        assert any("上限" in n for n in d["notes"])


def test_every_mismatch_is_reported_per_rule(tmp_path):
    body = ('<row r="1"><c r="A1"><v>2</v></c><c r="B1"><f>A1*2</f><v>5</v></c></row>'
            '<row r="2"><c r="A2"><v>3</v></c><c r="B2"><f>A2*2</f><v>6</v></c></row>'
            '<row r="3"><c r="A3"><v>4</v></c><c r="B3"><f>A3*2</f></c></row>')
    path = put(tmp_path, raw_xlsx({"S": body}))
    mod = tmp_path / "m.py"
    res = api.extract_processing(path, format="python", out_path=str(mod))
    v = res["verification"]
    assert v["stored"] == {"compared": 2, "match": 1, "rate": "50.0000%"}
    assert v["evaluator"]["rate"] == "100%"
    assert v["mismatches"] == [{"cell": "S!B1", "rule": "R1", "kind": "stored", "stored": 5, "module": 4}]
    st = res["rules"][0]["verification"]
    assert st["status"] == "mismatch" and st["counts"]["no_stored"] == 1
    m = load_module(str(mod))
    assert m.VERIFICATION["not_proven"]["R1"]["status"] == "mismatch"


def test_cycle_and_data_table_are_not_proven(tmp_path):
    body = ('<row r="1"><c r="A1"><f>B1+1</f><v>0</v></c><c r="B1"><f>A1+1</f><v>0</v></c>'
            '<c r="D1"><f t="dataTable" ref="D1:E2" dt2D="1" r1="A1" r2="B1"/><v>1</v></c><c r="E1"><v>2</v></c></row>'
            '<row r="2"><c r="D2"><v>3</v></c><c r="E2"><v>4</v></c></row>')
    path = put(tmp_path, raw_xlsx({"S": body}))
    mod = tmp_path / "m.py"
    res = api.extract_processing(path, format="python", out_path=str(mod))
    by = {d["range"]: d for d in res["rules"]}
    assert by["A1"]["verification"]["status"] == "not_proven" and "循環" in by["A1"]["verification"]["why"][0]
    dt = by["D1:E2"]
    assert dt["kind"] == "data_table" and dt["verification"]["status"] == "not_proven" and dt["definition"]["r1"] == "A1"
    assert not any(d["range"].startswith("E") for d in res["inputs"])  # データテーブルの中の値は定数ではない
    m = load_module(str(mod))
    with pytest.raises(m.NotReproducible):
        m.calculate()


def test_tables_names_subtotal_and_deep_chain(tmp_path):
    def fill(wb):
        ws = wb.active
        ws.title = "S"
        ws.append(["品", "額", "倍"])
        for i in range(2, 6):
            ws.append([f"x{i}", i * 10, f"=[@額]*2"])
        tb = openpyxl.worksheet.table.Table(displayName="T1", ref="A1:C5")
        ws.add_table(tb)
        ws["E1"] = "=SUM(T1[額])"
        ws["E2"] = "=SUBTOTAL(9,B2:B5)"
        ws["E3"] = "=SUBTOTAL(109,B2:B5)"
        ws.row_dimensions[3].hidden = True
        ws["G1"] = 1
        for r in range(2, 3000):
            ws.cell(r, 7, f"=G{r - 1}+1")

    # ToDo は LibreOffice の構造化参照の保存の違いを避け、キャッシュのない形で確かめる（評価器との一致）
    p = book(tmp_path, fill)
    mod = tmp_path / "m.py"
    res = api.extract_processing(p, format="python", out_path=str(mod))
    v = res["verification"]
    assert v["ran"] and v["evaluator"]["rate"] == "100%", v["mismatches"][:5]
    assert rule(res, "G2:G2999 = G{row-1}+1")["cells"] == 2998
    m = load_module(str(mod))
    got = m.calculate({"S!B2": 100}, all_cells=True)
    assert got["S!E1"] == 100 + 30 + 40 + 50 and got["S!C2"] == 200
    assert got["S!E2"] == 220 and got["S!E3"] == 190  # 109 は隠した行（3 行目の 30）を除く
    assert m.calculate({"S!G1": 5}, outputs=["S!G2999"]) == {"S!G2999": 3003.0}


def test_all_registered_functions_can_be_vendored(tmp_path):
    from todoku.processing.vendor import Vendor

    v = Vendor(set())
    v.load_registry()
    v.need("formula.functions", "call")
    v.need("formula.functions", "EAGER")
    v.need("formula.functions", "LAZY")
    v.need("formula.evaluator", "Evaluator")
    for f in v.registry:
        v.want_function(f)
    v._drain()
    src, final = v.source()
    full = "from __future__ import annotations\n" + "\n".join(sorted(v.stdlib)) + "\n" + src + "\n" + "\n".join(v.registrations(final))
    import types

    mod = types.ModuleType("todoku_vendored_check")
    sys.modules[mod.__name__] = mod
    ns = mod.__dict__
    code = compile(full, "vendored", "exec")
    exec(code, ns)  # noqa: S102  とどくのソースから切り出したもの
    # 関数の中の大域の名前は、すべて定義されている（切り出し漏れがない）
    import builtins

    missing = set()

    def walk(t):
        for s in t.get_symbols():
            # Python 3.12 から、最上位の内包表記の変数はモジュールの表に入る（PEP 709）。代入される名前は除く
            if t.get_type() == "module" and s.is_assigned():
                continue
            if s.is_global() and s.is_referenced() and s.get_name() not in ns and not hasattr(builtins, s.get_name()):
                missing.add(s.get_name())
        for c in t.get_children():
            walk(c)

    walk(symtable.symtable(full, "vendored", "exec"))
    assert not missing, missing
    eager = ns[v.global_name(final, "formula.functions", "EAGER")]
    lazy = ns[v.global_name(final, "formula.functions", "LAZY")]
    assert set(v.registry) == set(eager) | set(lazy)


def test_source_text_cannot_become_code(tmp_path):
    """書類の字（シートの名前・数式の中の字）が、書き出したソースのコードにならない。"""
    evil = 'x\r__import__("os").system("touch PWNED")#'
    def fill(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = 1
        ws["B1"] = '=A1&"' + 'a"""\\' + '\n__import__(1)"'

    path = book(tmp_path, fill)
    mod = tmp_path / "m.py"
    res = api.extract_processing(path, format="python", out_path=str(mod))
    assert res["verification"]["ran"], res["verification"]["notes"]
    from todoku.processing.pygen import _safe

    assert "\r" not in _safe(evil) and "\n" not in _safe(evil)
    for line in mod.read_text(encoding="utf-8").splitlines():
        assert not line.startswith("__import__(")


# ---------------------------------------------------------------- 形式
@pytest.mark.parametrize("name", ["public/sheet.xlsx", "public/sheet.xls", "public/sheet.ods", "date1904.xls", "kami_meisai.xlsx", "r4_general.ods", "r4_general.xls"])
def test_excel_family_formats_verify(tmp_path, name):
    mod = tmp_path / "m.py"
    res = api.extract_processing(str(FIX / name), format="python", out_path=str(mod))
    v = res["verification"]
    assert v["ran"] and v["stored"]["compared"] > 0 and v["stored"]["rate"] == "100%", v
    assert v["evaluator"]["rate"] == "100%"


def test_xlsb_formulas(tmp_path):
    sys.path.insert(0, str(Path(__file__).parent))
    import xlsb_maker as M

    rows = b"".join([M.row(0), M.real(0, 2.0), M.real(1, 3.0), M.fnum(2, 6.0, M.p_ref(0, 0) + M.p_ref(0, 1) + M.MUL),
                     M.row(1), M.real(0, 4.0), M.real(1, 5.0), M.fnum(2, 20.0, M.p_ref(1, 0) + M.p_ref(1, 1) + M.MUL)])
    p = tmp_path / "b.xlsb"
    p.write_bytes(M.package({"S": M.sheet(rows)}))
    mod = tmp_path / "m.py"
    res = api.extract_processing(str(p), format="python", out_path=str(mod))
    assert res["kind"] == "xlsb" and res["summary"]["rules"] == 1
    assert res["rules"][0]["text"] == "C1:C2 = A{row}*B{row}"
    assert res["verification"]["stored"]["rate"] == "100%"


def test_other_formats_are_refused(tmp_path):
    p = tmp_path / "a.csv"
    p.write_text("a,b\n1,2\n", encoding="utf-8")
    with pytest.raises(api.ApiError, match="処理の取り出し"):
        api.extract_processing(str(p))
    with pytest.raises(api.ApiError, match="format"):
        api.extract_processing(str(FIX / "kami_meisai.xlsx"), format="sql")
    with pytest.raises(api.ApiError, match="out_path"):
        api.extract_processing(str(FIX / "kami_meisai.xlsx"), format="python")


def test_too_large_results_are_condensed(tmp_path):
    from todoku import processing as P

    path = str(FIX / "kami_meisai.xlsx")
    blob = Path(path).read_bytes()
    r = P.extract(path, blob, "kami_meisai.xlsx", "xlsx", budget_chars=2000)
    assert r["counts"]["rules"] == 14 and r["rules_shown"] < 14 and "out_path" in r["note"]


def test_wide_array_declaration_is_not_expanded(tmp_path):
    body = '<row r="1"><c r="A1"><f t="array" ref="A1:XFD1048576">1+1</f><v>2</v></c></row>'
    res = api.extract_processing(put(tmp_path, raw_xlsx({"S": body})))
    assert res["rules"][0]["range"] == "A1" and "note" in res["rules"][0]


def test_json_written_and_only_sheet(tmp_path):
    out = tmp_path / "p.json"
    res = api.extract_processing(str(FIX / "kami_meisai.xlsx"), out_path=str(out), sheet="明細書")
    assert res["out_path"] == str(out)
    full = json.loads(out.read_text(encoding="utf-8"))
    assert full["only_sheet"] == "明細書" and all(d["sheet"] == "明細書" for d in full["rules"])
    assert res["summary"]["rules"] == 14


# ---------------------------------------------------------------- PDF のフォーム
def test_pdf_calculation_rules_and_module(tmp_path):
    sys.path.insert(0, str(Path(__file__).parent))
    from pdf_forms_maker import build_all as make_forms_pdf

    p = tmp_path / "f.pdf"
    p.write_bytes(make_forms_pdf())
    res = api.extract_processing(str(p))
    calc = [d for d in res["rules"] if d["kind"] == "field_calculation"]
    assert any(d.get("operation") == "PRD" and d["inputs"] == ["qty", "price"] for d in calc)
    d = next(x for x in calc if x.get("operation") == "PRD")
    assert d["evaluated"] is True and d["text"] == f"{d['field']} = PRD(qty, price)"
    assert res["order"] and {x["field"] for x in res["inputs"]} >= {"qty", "price"}
    assert any(x["kind"] == "format" for x in res["field_scripts"])
    mod = tmp_path / "m.py"
    r2 = api.extract_processing(str(p), format="python", out_path=str(mod))
    assert r2["verification"]["ran"], r2["verification"]
    m = load_module(str(mod))
    assert m.calculate({"qty": 3, "price": "1,000"}, on_unknown="mark")[d["field"]] == 3000


def test_pdf_custom_javascript_is_listed_verbatim(tmp_path):
    sys.path.insert(0, str(Path(__file__).parent))
    from pdf_forms_maker import build_all as make_forms_pdf

    blob = make_forms_pdf()
    # 計算のスクリプトを独自の JavaScript に置き換えたもの（長さが変わるので作り直す：字を 1 か所だけ入れ替え）
    if b'AFSimple_Calculate\\("PRD"' not in blob:
        pytest.skip("様式の作りが違う")
    custom = blob.replace(b'AFSimple_Calculate\\("PRD", new Array \\("qty", "price"\\)\\);', b'event.value = this.getField\\("qty"\\).value;    ')
    p = tmp_path / "c.pdf"
    p.write_bytes(custom)
    res = api.extract_processing(str(p))
    d = next(x for x in res["rules"] if x["kind"] == "field_calculation")
    assert d["operation"] == "custom_javascript" and d["evaluated"] is False
    assert 'getField("qty")' in d["source"] and d["fields_named_in_source"] == ["qty"]
    mod = tmp_path / "m.py"
    r2 = api.extract_processing(str(p), format="python", out_path=str(mod))
    assert next(x for x in r2["rules"] if x["id"] == d["id"])["verification"]["status"] == "not_proven"
    m = load_module(str(mod))
    with pytest.raises(m.NotReproducible):
        m.calculate()


# ---------------------------------------------------------------- MCP
def test_mcp_tool(tmp_path):
    from todoku import mcp_server

    names = {t["name"] for t in mcp_server.TOOLS}
    assert "extract_processing" in names
    msgs = [{"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}},
            {"jsonrpc": "2.0", "id": 2, "method": "tools/call", "params": {"name": "extract_processing", "arguments": {"path": str(FIX / "kami_meisai.xlsx")}}}]
    p = subprocess.run([sys.executable, "-m", "todoku.mcp_server"], input="\n".join(json.dumps(m) for m in msgs) + "\n", capture_output=True, text=True,
                       env={"PYTHONPATH": SRC, "PATH": "/usr/bin:/bin"}, timeout=120)
    out = [json.loads(x) for x in p.stdout.splitlines() if x.strip()]
    res = next(x for x in out if x.get("id") == 2)["result"]
    assert res["structuredContent"]["summary"]["rules"] == 14
    assert "規則 14" in res["content"][0]["text"]
    mod = tmp_path / "m.py"
    msgs[1]["params"]["arguments"].update(format="python", out_path=str(mod))
    p = subprocess.run([sys.executable, "-m", "todoku.mcp_server"], input="\n".join(json.dumps(m) for m in msgs) + "\n", capture_output=True, text=True,
                       env={"PYTHONPATH": SRC, "PATH": "/usr/bin:/bin"}, timeout=120)
    out = [json.loads(x) for x in p.stdout.splitlines() if x.strip()]
    res = next(x for x in out if x.get("id") == 2)["result"]["structuredContent"]
    assert res["verification"]["stored"]["rate"] == "100%" and mod.exists()
