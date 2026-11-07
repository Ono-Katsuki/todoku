"""評価の試運転で見つかったことの試験：まとめた行・結合・字の書式・ふりがな・文字の並び順・規則の説明・道具の説明の長さ。"""

from __future__ import annotations

import io
import sys
import zipfile
from pathlib import Path

import pytest

from todoku import api, mcp_server
from todoku.processing.consistency import check, check_rule

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts" / "eval"))
sys.path.insert(0, str(ROOT / "scripts"))

openpyxl = pytest.importorskip("openpyxl")
LIMIT = 4096  # Claude Code（2.1.296）の CLAUDE_CODE_MAX_MCP_DESCRIPTION_LENGTH の既定：道具の説明とサーバーの説明の両方をこの字数で切る


def window_text(path) -> str:
    return "\n".join(s for _n, s in api.window(str(path))[2])


# ---------------------------------------------------------------- 道具の説明の長さ
def utf16_len(s: str) -> int:
    return len(s.encode("utf-16-le")) // 2  # Claude Code は JavaScript の length で数える


def test_every_tool_description_fits():
    long = {t["name"]: utf16_len(t["description"]) for t in mcp_server.TOOLS if utf16_len(t["description"]) > LIMIT}
    assert not long


def test_server_instructions_fit():
    assert utf16_len(mcp_server._instructions()) <= LIMIT


def test_fill_description_keeps_what_agents_need_up_front():
    d = next(t for t in mcp_server.TOOLS if t["name"] == "fill_document")["description"]
    for fact in ("out_path", "true/false", "as", "old", "circle", "seal", "consent", "overwrite_text", "allow_hidden", "view_page", "say_to_user", "help"):
        assert fact in d


def test_fill_help_returns_detail_without_writing(tmp_path):
    from todoku.fillhelp import TOPICS

    schema = next(t for t in mcp_server.TOOLS if t["name"] == "fill_document")["inputSchema"]["properties"]["help"]
    assert set(schema["enum"]) == {*TOPICS, "all"}
    out = tmp_path / "out.xlsx"
    r = mcp_server.tool_result("fill_document", {"path": "", "out_path": str(out), "values": {}, "help": "circle"})
    assert "isError" not in r or not r["isError"]
    assert "near" in r["content"][0]["text"] and "occurrence" in r["content"][0]["text"]
    assert not out.exists()
    allt = mcp_server.call_tool("fill_document", {"help": "all"})["content"][0]["text"]
    for v in TOPICS.values():
        assert v in allt


# ---------------------------------------------------------------- 規則の説明
def _same_shape_book(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(1, 12):
        ws.cell(r, 1, "✓")
        ws.cell(r, 2, "✓")
    ws["D1"] = '=COUNTIF(A1:B3,"✓")'
    ws["D5"] = '=COUNTIF(A5:B7,"✓")'  # 同じ形（R1C1）で、別の所を読む
    ws["D9"] = '=COUNTIF(A9:B11,"✓")'
    wb.save(path)


def test_rule_text_points_at_its_own_cells(tmp_path):
    p = tmp_path / "s.xlsx"
    _same_shape_book(p)
    r = api.extract_processing(str(p))
    rules = {d["range"]: d for d in r["rules"]}
    for cell, rng in (("D1", "A1:B3"), ("D5", "A5:B7"), ("D9", "A9:B11")):
        d = rules[cell]
        assert rng in d["text"] and rng in d["a1"] and d["reads"][0]["pattern"] == rng and d["reads"][0]["range"] == rng
    ec = r["explanation_check"]
    assert ec["ok"] and ec["checked"] == len(r["rules"]) and ec["checked"] >= 3


def test_consistency_check_catches_wrong_text():
    good = {"id": "R1", "kind": "formula", "range": "D5", "text": 'D5 = COUNTIF(A5:B7,"x")', "formula": '=COUNTIF(A5:B7,"x")',
            "a1": '=COUNTIF(A5:B7,"x")', "r1c1": '=COUNTIF(RC[-3]:R[2]C[-2],"x")', "reads": [{"pattern": "A5:B7", "range": "A5:B7", "sheet": "S"}]}
    assert check_rule(good, (5, 4)) == []
    bad = dict(good, text='D5 = COUNTIF(A1:B3,"x")', a1='=COUNTIF(A1:B3,"x")', reads=[{"pattern": "A1:B3", "range": "A5:B7", "sheet": "S"}])
    what = " ".join(check_rule(bad, (5, 4)))
    assert "r1c1" in what and "formula" not in what[:0] and "ファイルの数式" in what and "reads" in what
    res = check([bad], {"R1": (5, 4)}, 64)
    assert not res["ok"] and res["failed"] == 1 and res["failures"][0]["rule"] == "R1"


def test_consistency_check_cannot_pass_silently():
    res = check([{"id": "R1", "kind": "data_table", "range": "A1"}], {"R1": (1, 1)}, 64)
    assert res["checked"] == 0 and res["unchecked"] == 1 and res["unchecked_why"] == {"data_table": 1}
    res = check([], {}, 64)
    assert res["rules"] == 0 and res["checked"] == 0  # 規則がないことも、数でそのまま分かる


# ---------------------------------------------------------------- まとめた行・範囲の読み
def _collapsed_book(path):
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    dv = DataValidation(type="list", formula1='"a,b"')
    ws.add_data_validation(dv)
    ws["A1"] = "番号"
    ws["B1"] = "選ぶ"
    for r in range(2, 21):
        ws.cell(r, 1, r - 1)
        dv.add(f"B{r}")
        ws.cell(r, 3, "済")
        ws.cell(r, 4, f'=IF(B{r}="","入力済","")')
    ws.column_dimensions["D"].hidden = True
    wb.save(path)


def test_collapsed_rows_are_stated_and_reachable(tmp_path):
    p = tmp_path / "c.xlsx"
    _collapsed_book(p)
    t = window_text(p)
    line = next(ln for ln in t.splitlines() if "同じ形の行が" in ln)
    assert "字が違うのは A 列だけ" in line and "非表示の列も" in line and "read_window の sheet・rows" in line
    page, _doc = api.read_range(str(p), {"sheet": "S", "rows": "10-12"})
    body = page.text
    assert "同じ形の行が" not in body  # 頼んだ行は、まとめずに 1 行ずつ
    for r in (10, 11, 12):
        assert f"- {r}: " in body and f"［B{r}］" in body
    assert "- 13: " not in body  # 頼まなかった行は出さない（読んでいないので、作った字を出さない）


@pytest.mark.skipif(sys.platform == "win32", reason="評価の道具（scripts/eval）は POSIX だけ")
def test_expand_collapsed_counts_stated_rows():
    from read_compare import expand_collapsed

    md = "## シート「S」\n\n- 3: 1 ｜ 済 ｜ ［B3］\n- 4〜5: 3行目と同じ形の行が 2 行続きます（入力欄はすべて空欄。字が違うのは A 列だけで、ほかの字（非表示の列も）は同じ：A4＝2、A5＝3）　1 行ずつは read_window の sheet・rows で読めます\n"
    out = expand_collapsed(md)
    assert out.count("済") == 3 and "同じ形の行" not in out and "- 2" in out


@pytest.mark.skipif(sys.platform == "win32", reason="評価の道具（scripts/eval）は POSIX だけ")
def test_pdf_units_split_at_spaces_for_bigrams():
    from read_compare import _recall_units

    assert _recall_units({"kind": "pdf", "units": ["受 所", "氏名"]}) == ["受", "所", "氏名"]
    assert _recall_units({"kind": "docx", "units": ["受 所"]}) == ["受 所"]


def test_docx_probe_does_not_double_count_textbox(tmp_path):
    from probes import probe_docx

    w = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    xml = (f'<w:document {w}><w:body><w:p><w:r><w:t>本文</w:t></w:r><w:r><w:pict><w:txbxContent><w:p><w:r><w:t>箱の字</w:t></w:r></w:p>'
           '</w:txbxContent></w:pict></w:r></w:p></w:body></w:document>')
    p = tmp_path / "t.docx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("word/document.xml", xml)
    r = probe_docx(str(p))
    assert r["paragraphs"] == ["本文"] and r["textboxes"] == ["箱の字"]


# ---------------------------------------------------------------- 結合
def _merged_book(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for r in range(1, 6):
        ws.cell(r, 1, f"項目{r}")
        ws.cell(r, 2, f"値{r}")
    ws["D2"] = "申請者"  # 右の列の縦の結合：行の始めの字より右にある
    ws.merge_cells("D2:D4")
    ws["E2"] = "備考"
    wb.save(path)


def test_merged_value_reaches_every_row_it_covers(tmp_path):
    p = tmp_path / "m.xlsx"
    _merged_book(p)
    lines = window_text(p).splitlines()
    row = {int(ln[2:].split(":")[0]): ln for ln in lines if ln[:2] == "- " and ln[2:3].isdigit() and ":" in ln}
    assert "申請者〔D2:D4〕" in row[2]  # 始めの行に、かかる行の範囲の事実
    assert "項目3 ｜ 値3 ｜ （申請者）" in row[3] and "（申請者）" in row[4]  # かかる行に、列の位置で
    assert "申請者" not in row[1] and "申請者" not in row[5]


def test_read_table_states_merged_range(tmp_path):
    p = tmp_path / "t.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws.append(["区分", "名前"])
    ws["A2"] = "第一"
    ws.merge_cells("A2:A4")
    for r, n in ((2, "甲"), (3, "乙"), (4, "丙")):
        ws.cell(r, 2, n)
    ws["A5"] = "第二"
    ws["B5"] = "丁"
    wb.save(p)
    rows = api.read_table(str(p), "区分", "S")["all_rows"]
    assert [x["区分"] for x in rows] == ["第一", "第一", "第一", "第二"]
    assert [x.get("_merged", {}).get("区分") for x in rows] == ["A2:A4", "A2:A4", "A2:A4", None]
    assert rows[2]["_cells"]["区分"] == "A2"


# ---------------------------------------------------------------- 字の書式・ふりがな
def _styled_book(path):
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    for r in range(1, 8):
        ws.cell(r, 1, f"ふつうの字{r}")
    ws["B1"] = "取り消した字"
    ws["B1"].font = Font(strike=True)
    ws["B2"] = "赤い字"
    ws["B2"].font = Font(color="FFFF0000")
    wb.save(path)


def test_strike_and_color_are_stated(tmp_path):
    p = tmp_path / "f.xlsx"
    _styled_book(p)
    t = window_text(p)
    assert "〔取り消し線〕取り消した字" in t and "〔字の色 #FF0000〕赤い字" in t
    assert "〔字の色 #FF0000〕ふつう" not in t


def test_ordinary_window_gets_no_font_notes(tmp_path):
    p = tmp_path / "o.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "字"
    wb.save(p)
    t = window_text(p)
    assert "取り消し線" not in t and "字の色 #" not in t and "### ふりがな" not in t


def _sst_book(path, strings: list[str]) -> None:
    """共有文字列をそのまま書いたブック（A1 から順に）。"""
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    rows = "".join(f'<row r="{i}"><c r="A{i}" t="s"><v>{i - 1}</v></c></row>' for i in range(1, len(strings) + 1))
    parts = {
        "[Content_Types].xml": '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
        "_rels/.rels": f'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<?xml version="1.0"?><workbook xmlns="{main}" xmlns:r="{rel}"><sheets><sheet name="S" sheetId="1" r:id="r1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": f'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="{rel}/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": f'<?xml version="1.0"?><worksheet xmlns="{main}"><sheetData>{rows}</sheetData></worksheet>',
        "xl/sharedStrings.xml": f'<?xml version="1.0"?><sst xmlns="{main}" count="{len(strings)}" uniqueCount="{len(strings)}">' + "".join(f"<si>{x}</si>" for x in strings) + "</sst>",
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for k, v in parts.items():
            z.writestr(k, v)


def test_rich_runs_and_furigana_are_stated(tmp_path):
    p = tmp_path / "r.xlsx"
    _sst_book(p, ['<r><t>前半</t></r><r><rPr><strike/></rPr><t>後半</t></r>', '<r><t>地の字</t></r><r><rPr><color rgb="FFFF0000"/></rPr><t>赤</t></r>',
                  '<r><t>字</t></r><r><rPr><strike/></rPr><t xml:space="preserve">\n</t></r>', '<t>氏名</t><rPh sb="0" eb="2"><t>シメイ</t></rPh><phoneticPr fontId="0"/>'])
    t = window_text(p)
    assert "前半後半〔うち取り消し線：「後半」〕" in t  # 字そのものは続けて読める
    assert "地の字赤〔うち字の色 #FF0000：「赤」〕" in t
    assert "〔うち取り消し線：「」〕" not in t  # 字のない字の並び（改行）の書式は印にしない
    assert "### ふりがな" in t and "「氏名」シメイ" in t


def test_ods_ruby_and_strike_are_read(tmp_path):
    p = tmp_path / "r.fods"
    p.write_text("""<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" office:mimetype="application/vnd.oasis.opendocument.spreadsheet">
<office:automatic-styles><style:style style:name="ce1" style:family="table-cell"><style:text-properties style:text-line-through-style="solid"/></style:style></office:automatic-styles>
<office:body><office:spreadsheet><table:table table:name="S"><table:table-row><table:table-cell office:value-type="string"><text:p><text:ruby><text:ruby-base>氏名</text:ruby-base><text:ruby-text>シメイ</text:ruby-text></text:ruby></text:p></table:table-cell></table:table-row>
<table:table-row><table:table-cell table:style-name="ce1" office:value-type="string"><text:p>消した字</text:p></table:table-cell></table:table-row></table:table></office:spreadsheet></office:body></office:document>""", encoding="utf-8")
    t = window_text(p)
    assert "「氏名」シメイ" in t and "〔取り消し線〕消した字" in t


def test_furigana_uses_cell_text_not_whole_reading():
    from todoku.xlsx.fontfacts import reading_parts

    assert reading_parts("氏名　様", "シメイ　サマ") == ["シメイ", "サマ"]


# ---------------------------------------------------------------- 文字の並び順
def _order_book(path, formula: str, key="処遇Ⅱ"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    for i, h in enumerate(("処遇Ⅰ", "処遇Ⅱ", "処遇Ⅲ"), 1):
        ws.cell(1, i, h)
    ws["A2"] = key
    ws["B2"] = formula
    wb.save(path)


def test_text_order_lookup_is_told(tmp_path):
    p = tmp_path / "o.xlsx"
    _order_book(p, "=MATCH(A2,A1:C1)")
    t = window_text(p)
    assert "文字の並び順で結果が変わる数式 1 か所" in t and "B2" in t
    g = api.grade(str(p))
    assert g["grade"] == "C" and any(r["title"] == "文字の並び順で結果が変わる数式" for r in g["reasons"])


@pytest.mark.parametrize("formula,key", [("=MATCH(A2,A1:C1,0)", "処遇Ⅱ"), ("=MATCH(A2,A1:C1)", ""), ("=VLOOKUP(A2,A1:C1,2,FALSE)", "処遇Ⅱ"), ("=MATCH(5,A1:C1)", "処遇Ⅱ")])
def test_text_order_not_told_when_result_does_not_depend(tmp_path, formula, key):
    p = tmp_path / "o.xlsx"
    _order_book(p, formula, key)
    assert "文字の並び順" not in window_text(p)


# ---------------------------------------------------------------- 区間
@pytest.mark.skipif(sys.platform == "win32", reason="評価の道具（scripts/eval）は POSIX だけ")
def test_portal_and_eval_use_same_interval():
    import world_counts
    import world_portal
    from common import korn_graubard, stratified_estimate

    cells = [(1000, 0.8, 0.01, 20, 8), (500, 0.5, 0.02, 10, 4), (300, 1.0, 0.0, 12, 3)]
    a, b, c = world_portal.combine(cells), world_counts.est(cells), stratified_estimate(cells)
    assert a["ci"] == [round(x, 4) for x in c["ci"]] and b["p_ci"] == c["ci"] and a["n"] == b["n"] == 42
    assert not hasattr(world_portal, "wilson")
    # デザインの自由度が小さいほど、区間は広い（Korn & Graubard の補正）
    wide, plain = korn_graubard(0.3, 0.01, 200, df=4), korn_graubard(0.3, 0.01, 200)
    assert wide["n_eff"] < plain["n_eff"] and wide["ci"][0] < plain["ci"][0] and wide["ci"][1] > plain["ci"][1]


def test_blank_in_a_cell_keeps_the_characters_of_its_number_format(tmp_path):
    """表示形式が足した字（"〔"@"〕"）：文字の中の空欄を印に置き換えても、見えている字を窓から落とさない。"""
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws["A1"] = "○○養生した試験"
    ws["A1"].number_format = '"〔"@"〕"'
    p = tmp_path / "a.xlsx"
    wb.save(p)
    t = "\n".join(x[1] for x in api.window(str(p))[2])
    assert "〔A1〕〔［？○○］養生した試験〕" in t, t
    c = next(x for x in api.audit_whole(str(p))["checks"] if x["name"] == "読み落とし（Markdown の窓）")
    assert c["ok"], c["items"]
