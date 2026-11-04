"""版を固める前の修正：宣言された広い範囲でメモリを使わない・壊れた .xls の生のエラー・PDF の種類の指定（kind）・字の囲み線の幅。"""

from __future__ import annotations

import io
import os
import random
import threading
import time
import tracemalloc
import zipfile

import pytest
from test_controls import make_book as xlsx_controls
from test_docx import make_docx
from test_pdf import sample_form
from test_wrap_grid_circle import _docx_cell
from test_xls_controls import make_book

from todoku import api
from todoku.wrapcheck import MESSAGE, count_lines
from todoku.xls.book import XlsWorkbook
from todoku.xlsx.calc import Calc
from todoku.xlsx.window import Box, Layout, _Cover, build
from todoku.xlsx.workbook import Workbook

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
STYLES = (f'<styleSheet {NS}><fonts count="1"><font><sz val="11"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
          '<border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
          '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/></cellXfs></styleSheet>')
LIMIT = 40 * 1024 * 1024  # 少しのセルの書類の窓に使ってよいメモリ
RUNAWAY = 400 * 1024 * 1024  # これを超えたら、PC を守るためにテストの途中で止める
CT = "application/vnd.openxmlformats-officedocument.spreadsheetml"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG = "http://schemas.openxmlformats.org/package/2006"


def xlsx(body: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", f'<?xml version="1.0"?><Types xmlns="{PKG}/content-types">'
                   f'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   f'<Override PartName="/xl/workbook.xml" ContentType="{CT}.sheet.main+xml"/>'
                   f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{CT}.worksheet+xml"/>'
                   f'<Override PartName="/xl/styles.xml" ContentType="{CT}.styles+xml"/></Types>')
        z.writestr("_rels/.rels", f'<?xml version="1.0"?><Relationships xmlns="{PKG}/relationships">'
                   f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", f'<?xml version="1.0"?><workbook {NS} xmlns:r="{REL}">'
                   '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", f'<?xml version="1.0"?><Relationships xmlns="{PKG}/relationships">'
                   f'<Relationship Id="rId1" Type="{REL}/worksheet" Target="worksheets/sheet1.xml"/>'
                   f'<Relationship Id="rId2" Type="{REL}/styles" Target="styles.xml"/></Relationships>')
        z.writestr("xl/styles.xml", STYLES)
        z.writestr("xl/worksheets/sheet1.xml", f'<?xml version="1.0"?><worksheet {NS}>{body}</worksheet>')
    return buf.getvalue()


def cells(extra: str = "") -> str:
    return ('<sheetData><row r="1"><c r="A1" t="inlineStr" s="1"><is><t>a</t></is></c><c r="B1" t="inlineStr" s="1"><is><t>b</t></is></c></row>'
            f'<row r="2"><c r="A2" s="1"/><c r="B2" s="1"/></row>{extra}</sheetData>')


def small_memory(fn):
    """fn のメモリの山が小さいこと。止まらないほど増えるときは、PC を守るためプロセスごと止める。"""
    done = threading.Event()

    def watch():
        while not done.wait(0.05):
            if tracemalloc.get_traced_memory()[0] > RUNAWAY:
                os.write(2, "メモリが増え続けるので止めました\n".encode())
                os._exit(99)

    tracemalloc.start()
    t = threading.Thread(target=watch, daemon=True)
    t.start()
    try:
        got = fn()
        peak = tracemalloc.get_traced_memory()[1]
    finally:
        done.set()
        t.join()
        tracemalloc.stop()
    assert peak < LIMIT, f"{peak / 1048576:.0f} MB"
    return got


HUGE_MERGE = '<mergeCells count="1"><mergeCell ref="D5:XFD1048576"/></mergeCells>'
HUGE_DV = '<dataValidations count="1"><dataValidation type="whole" sqref="A1:XFD1048576"><formula1>1</formula1></dataValidation></dataValidations>'
STRAY = '<row r="1048576"><c r="XFD1048576" s="1"/></row>'
SUM_ALL = '<row r="3"><c r="A3"><f>SUM(A4:XFD1048576)</f></c></row>'
TALL = "".join(f'<row r="{i + 5}"><c r="{chr(65 + i % 26)}{i + 5}" t="inlineStr" s="1"><is><t>m</t></is></c></row>' for i in range(100))
TALL_MERGES = ('<mergeCells count="100">' + "".join(f'<mergeCell ref="{chr(65 + i % 26)}{i + 5}:{chr(65 + i % 26)}1048576"/>' for i in range(100))
               + "</mergeCells>")


@pytest.mark.parametrize("body", [
    pytest.param('<dimension ref="A1:XFD1048576"/>' + cells(), id="dimension"),
    pytest.param('<cols><col min="1" max="16384" width="9" style="1"/></cols>' + cells(), id="column_border"),
    pytest.param('<cols><col min="1" max="2147483647" width="9" hidden="1"/></cols>' + cells(), id="column_huge"),
    pytest.param(cells() + HUGE_MERGE + HUGE_DV, id="merge_and_validation"),
    pytest.param(cells(STRAY) + HUGE_DV, id="far_cell_and_validation"),
    pytest.param(cells(SUM_ALL + STRAY), id="far_cell_and_formula"),
    pytest.param(cells(SUM_ALL) + HUGE_MERGE, id="merge_and_formula"),
    pytest.param(cells('<row r="3"><c r="A3"><f t="array" ref="A3:XFD1048576">A1</f></c></row>'), id="array"),
    pytest.param(cells(TALL) + TALL_MERGES, id="tall_merges"),
])
def test_declared_ranges_cost_no_memory(body):
    blob = xlsx(body)
    doc = small_memory(lambda: build(Workbook(blob, "t.xlsx")))
    assert doc["sheets"][0]["name"] == "S"


def test_declared_range_not_expanded_is_reported_and_small_one_still_is():
    doc = build(Workbook(xlsx(cells(STRAY) + HUGE_DV), "t.xlsx"))
    assert any("入力規則" in u and "広げていません" in u for u in doc["unread"])
    # 値も書式もないセルの入力規則でも、ふつうの広さなら欄（これまでと同じ）
    small = '<dataValidations count="1"><dataValidation type="whole" sqref="B3:B6"><formula1>1</formula1></dataValidation></dataValidations>'
    doc = build(Workbook(xlsx(cells('<row r="7"><c r="A7" s="1"/></row>') + small), "t.xlsx"))
    assert not doc["unread"]
    assert {"B3", "B4", "B5", "B6"} <= {f["cell"] for f in doc["sheets"][0]["fields"]}


def test_unread_reaches_the_markdown_window(tmp_path):
    p = tmp_path / "t.xlsx"
    p.write_bytes(xlsx(cells(STRAY) + HUGE_DV))
    assert any("読んでいない部分" in text for _title, text in api.window(str(p))[2])
    assert any(c["name"] == "読んでいない部分" and not c["ok"] for c in api.audit(str(p))["checks"])
    assert any(r["title"] == "広げていない範囲" and r["grade"] == "C" for r in api.grade(str(p))["reasons"])


def test_huge_range_in_a_formula_is_unsupported_not_a_blow_up():
    wb = Workbook(xlsx(cells(SUM_ALL + STRAY)), "t.xlsx")
    calc = Calc(wb)
    small_memory(lambda: calc.cell_value("S", 3, 1))
    assert calc.status[("S", 3, 1)][0] == "unsupported" and "広すぎます" in calc.status[("S", 3, 1)][1]


def test_cover_answers_like_a_table_of_rows():
    low = Box(2, 1, 4, 1, "text")
    tall = Box(1, 2, 1000, 2, "text")
    later = Box(3, 3, 5, 3, "text")
    cover = _Cover()
    for b in (low, tall, later):
        cover.add(b)
    assert cover.get(1) == [] and cover.get(2) == [tall]
    assert cover.get(3) == [low, tall] and cover.get(4) == [low, tall, later] and cover.get(5) == [tall, later]
    assert cover.get(1000) == [tall] and cover.get(1001, ["x"]) == ["x"]
    cover.drop(lambda b: b is low)
    assert cover.get(3) == [tall] and cover.boxes() == [tall, later]


def test_rows_below_a_tall_merge_find_the_merge_when_searching_up():
    merges = '<mergeCells count="1"><mergeCell ref="A1:A200"/></mergeCells>'
    wb = Workbook(xlsx(cells('<row r="300"><c r="C300" s="1"/></row>') + merges), "t.xlsx")
    lay = Layout(wb, wb.sheets[0])
    assert lay._next_row(150, (1, 1)) == 150 and lay._next_row(250, (1, 1)) == 200
    assert lay._next_row(250, (5, 5)) == 0


def test_control_declaring_the_whole_sheet_costs_no_memory():
    wb = XlsWorkbook(make_book()[0], "t.xls")
    for ctl in wb.sheets[0].controls:
        if ctl.anchor:
            ctl.anchor_to = (65536, 256)
    small_memory(lambda: build(wb))


def test_control_ending_above_its_start_does_not_break_the_window():
    wb = XlsWorkbook(make_book()[0], "t.xls")
    for ctl in wb.sheets[0].controls:
        if ctl.anchor:
            ctl.anchor, ctl.anchor_to = (20, ctl.anchor[1]), (2, ctl.anchor[1])
    assert build(wb)["sheets"]


def _ods(table: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet", zipfile.ZIP_STORED)
        z.writestr("META-INF/manifest.xml", '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">'
                   '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>'
                   '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>')
        z.writestr("content.xml", '<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
                   'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">'
                   f'<office:body><office:spreadsheet><table:table table:name="S">{table}</table:table></office:spreadsheet></office:body></office:document-content>')
    return buf.getvalue()


def test_ods_repeated_rows_and_cells_stop_at_the_budget(monkeypatch):
    from todoku.odf import sheet as odf_sheet
    from todoku.odf.sheetwindow import build as ods_build

    monkeypatch.setattr(odf_sheet, "MAX_DECLARED_CELLS", 2000)
    blob = _ods('<table:table-row table:number-rows-repeated="1048576"><table:table-cell table:number-columns-repeated="16384" '
                'office:value-type="string"><text:p>x</text:p></table:table-cell></table:table-row>')
    wb = odf_sheet.OdsWorkbook(blob, "t.ods")
    assert len(wb.sheets[0].cells) <= 2001 and wb.unexpanded
    assert any("読みました" in u for u in ods_build(blob, "t.ods")["unread"])


def test_ods_declared_span_and_blank_repeats_cost_no_memory():
    from todoku.odf.sheet import OdsWorkbook

    blob = _ods('<table:table-row table:number-rows-repeated="1048576"><table:table-cell table:number-columns-repeated="16384"/></table:table-row>'
                '<table:table-row><table:table-cell table:number-rows-spanned="1048576" table:number-columns-spanned="16384" '
                'office:value-type="string"><text:p>m</text:p></table:table-cell></table:table-row>')
    small_memory(lambda: OdsWorkbook(blob, "t.ods"))


# ====================================================================== 2. 壊れた .xls の生のエラー
def test_sheet_that_cannot_be_located_is_refused_in_words():
    from todoku.xls.write import XlsWriteError
    from todoku.xls import write as W

    wb = XlsWorkbook(make_book()[0], "t.xls")
    editor = next(v for v in vars(W).values() if isinstance(v, type) and hasattr(v, "_sheet_ranges"))
    ed = editor(wb)
    name = wb.sheets[0].name
    ed.sheet_at.pop(name)
    with pytest.raises(XlsWriteError, match="見つかりません"):
        ed._span(name)


def test_formula_reference_to_a_missing_sheet_is_ref_in_the_calculation_model(tmp_path):
    p = tmp_path / "t.xlsx"
    p.write_bytes(xlsx(cells('<row r="3"><c r="A3"><f>Pozitivnalica1!A1+A2</f><v>1</v></c></row>')))
    assert isinstance(api.calc_model(str(p)), dict)
    assert str(Calc(Workbook(p.read_bytes(), "t.xlsx")).cell_value("S", 3, 1)) == "#REF!"


def test_control_list_on_a_missing_sheet_has_no_options_and_no_raw_error(tmp_path):
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(xlsx_controls())) as zi, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for i in zi.infolist():
            d = zi.read(i.filename)
            if i.filename.endswith((".xml", ".vml")):
                d = d.replace(b"$A$1:$A$3", b"Pozitivnalica1!$A$1:$A$3")
            zo.writestr(i, d)
    p = tmp_path / "t.xlsx"
    p.write_bytes(out.getvalue())
    assert api.fields(str(p)) and api.audit(str(p))


def test_corrupted_xls_only_raises_words(tmp_path):
    """壊れた .xls（シートの名前のあたりと、ほかの所をこわす）で、KeyError などの生のエラーが検査から出ない。"""
    blob = make_book()[0]
    at = blob.find("申込".encode("utf-16-le"))
    rnd = random.Random(11)
    p = tmp_path / "a.xls"
    t0 = time.time()
    tried = 0
    while tried < 200 and time.time() - t0 < 10:
        b = bytearray(blob)
        for _ in range(4):
            b[rnd.randrange(max(0, at - 40), at + 60)] = rnd.randrange(256)
        for _ in range(8):
            b[rnd.randrange(512, len(b))] = rnd.randrange(256)
        p.write_bytes(bytes(b))
        tried += 1
        try:
            api.audit(str(p))
        except ValueError:  # ApiError・書けない（XlsWriteError）・上限（LimitExceeded）：言葉で返る
            pass
    assert tried


# ====================================================================== 3. PDF：種類の指定は kind
def _comb(tmp_path):
    p = tmp_path / "s.pdf"
    p.write_bytes(sample_form())
    return p, next(f for f in api.fields(str(p)) if f.get("kind") == "comb")


def test_kind_selects_and_as_says_what_was_written(tmp_path):
    p, comb = _comb(tmp_path)
    r = api.fill(str(p), {comb["id"]: {"kind": "text", "as": "郵便番号", "value": "100-0001"}}, str(tmp_path / "o.pdf"))
    assert any("罫線がかかる見込み" in w for w in r["warnings"])  # 1 つの欄として書いた
    assert r["written"][0]["as"] == "郵便番号" and "郵便番号" in r["say_to_user"]


def test_old_as_is_a_kind_only_when_it_equals_a_candidate(tmp_path):
    p, comb = _comb(tmp_path)
    r = api.fill(str(p), {comb["id"]: {"as": "text", "value": "100-0001"}}, str(tmp_path / "o.pdf"))
    assert any("罫線がかかる見込み" in w for w in r["warnings"]) and not r["written"][0].get("as")
    r = api.fill(str(p), {comb["id"]: {"as": "郵便番号", "value": "1000001"}}, str(tmp_path / "o2.pdf"))
    assert not any("罫線がかかる見込み" in w for w in r["warnings"]) and r["written"][0]["as"] == "郵便番号"


def test_kind_with_layout_options_goes_through(tmp_path):
    p, comb = _comb(tmp_path)
    r = api.fill(str(p), {comb["id"]: {"kind": "text", "as": "郵便番号", "value": "100-0001", "align": "left"}}, str(tmp_path / "o.pdf"))
    assert r["written"][0]["as"] == "郵便番号"


def test_nothing_advertises_as_for_the_kind():
    from pathlib import Path

    import todoku

    root = Path(todoku.__file__).parent
    old = ['{"as": "text"', '{\\"as\\": \\"text\\"', '\\"as\\": 種類', '"as": 種類']
    texts = {f: (root / f).read_text(encoding="utf8") for f in ("pdf/window.py", "mcp_server.py")}
    texts["SKILL"] = (root.parent.parent / "skills/shinsei/SKILL.md").read_text(encoding="utf8")
    for name, t in texts.items():
        assert not any(o in t for o in old), name
    assert '\\"kind\\": \\"text\\", \\"as\\": \\"郵便番号\\"' in texts["pdf/window.py"]


# ====================================================================== 4. 字の囲み線の幅
def test_count_lines_adds_the_extra_width():
    assert count_lines("あいう", 36, 12) == 1
    assert count_lines("あいう", 36, 12, extra_at=lambda k: 1.0 if k == 0 else 0.0) == 2


def test_docx_default_border_circle_that_wraps_is_warned(tmp_path):
    # 12pt の字が 8 字でちょうど埋まるマス（96pt）：○の囲み線（太さ 3/4 pt の左右）で、1.5pt あふれる
    src = tmp_path / "t.docx"
    src.write_bytes(_docx_cell("あいうえ・有・無", 2136))
    r = api.fill(str(src), {"@c": {"circle": "無"}}, str(tmp_path / "o.docx"))
    assert r["verified"]
    hit = [w for w in r.get("warnings") or [] if MESSAGE in w]
    assert hit and "が 1 行から 2 行" in hit[0] and "幅いっぱい" not in hit[0] and "字の囲み線" in hit[0], r.get("warnings")
    assert MESSAGE in r["say_to_user"]


def test_docx_wide_cell_border_has_no_warning(tmp_path):
    src = tmp_path / "t.docx"
    src.write_bytes(_docx_cell("あいうえ・有・無", 3000))
    r = api.fill(str(src), {"@c": {"circle": "無"}}, str(tmp_path / "o.docx"))
    assert not any(MESSAGE in w for w in r.get("warnings") or []), r.get("warnings")


def _bordered_run(bdr: str, text: str) -> bytes:
    run = f'<w:r><w:rPr>{bdr}<w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'
    return make_docx('<w:tbl><w:tblGrid><w:gridCol w:w="2136"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2136" w:type="dxa"/></w:tcPr>'
                     f'<w:p>{run}</w:p></w:tc></w:tr></w:tbl>'
                     '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:left="1701" w:right="1701" w:top="1985" w:bottom="1701"/></w:sectPr>')


@pytest.mark.parametrize("bdr,warned", [
    ('<w:bdr w:val="single" w:sz="8" w:space="4"/>', True),  # 左右に 1 + 4 pt ずつ
    ('<w:bdr w:val="none" w:sz="8" w:space="4"/>', False),
    ("", False),
])
def test_docx_written_text_in_a_bordered_run_counts_the_border(tmp_path, bdr, warned):
    src = tmp_path / "t.docx"
    src.write_bytes(_bordered_run(bdr, "あいうえお"))  # 12pt × 5 = 60pt → 12pt × 8 = 96pt（ちょうどマスの幅）
    r = api.fill(str(src), {"@1": {"old": "あいうえお", "new": "あいうえおかきく"}}, str(tmp_path / "o.docx"))
    # 字だけなら幅ちょうど（「幅いっぱいに近く」の言い方）で、囲み線の分があふれると、はっきり行が増える言い方
    assert any(MESSAGE in w and "幅いっぱい" not in w for w in r["warnings"]) is warned, r["warnings"]
