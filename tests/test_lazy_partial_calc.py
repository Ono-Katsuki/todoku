"""大きなブックの「要る行だけの小さなブック」での計算：読んでいない行を空（0）として計算しない。保存された値を使うか、計算できないと伝える。"""

from __future__ import annotations

import zipfile

import pytest

from todoku import api, jobs, lazy, limits
from todoku.xlsx.workbook import Workbook

XL = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PK = "http://schemas.openxmlformats.org/package/2006/relationships"


@pytest.fixture(autouse=True)
def small_budget(monkeypatch, tmp_path):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2000)
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    jobs._purged.clear()
    lazy.clear()
    api._WINDOWS.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()


def _book(path, sheets: dict[str, str], names: str = ""):
    """sheets：シートの名前 → sheetData の中身（<row> の並び）。"""
    ct = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>')
    sh = "".join(f'<sheet name="{n}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, n in enumerate(sheets))
    wb = f'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{XL}" xmlns:r="{REL}"><sheets>{sh}</sheets>{names}</workbook>'
    wr = "".join(f'<Relationship Id="rId{i + 1}" Type="{REL}/worksheet" Target="worksheets/sheet{i + 1}.xml"/>' for i in range(len(sheets)))
    wr = f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PK}">{wr}<Relationship Id="rId99" Type="{REL}/styles" Target="styles.xml"/></Relationships>'
    st = (f'<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{XL}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
          '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PK}"><Relationship Id="rId1" Type="{REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wr)
        z.writestr("xl/styles.xml", st)
        for i, body in enumerate(sheets.values()):
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml", f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{XL}"><sheetData>{body}</sheetData></worksheet>')
    return path


def _filler(first: int, n: int) -> str:
    return "".join(f'<row r="{r}"><c r="A{r}"><v>{r}</v></c><c r="B{r}" t="inlineStr"><is><t>埋め草の字 {r}</t></is></c></row>' for r in range(first, first + n))


def _unit_price_book(path):
    """単価の形：C は別のシートの単価を読み、D は C と入力から差を計算する（D が読むのは同じ行だけ）。"""
    data = ('<row r="2"><c r="A2"><v>10</v></c><c r="B2"><f>A2*2</f><v>20</v></c>'
            "<c r=\"C2\"><f>'ほか'!A500</f><v>100</v></c>"
            '<c r="D2"><f>IF(C2-B2&gt;0,C2-B2,0)</f><v>80</v></c></row>') + _filler(3, 400)
    other = _filler(1, 499) + '<row r="500"><c r="A500"><v>100</v></c></row>'
    return _book(path, {"データ": data, "ほか": other})


def test_audit_does_not_call_a_formula_wrong_because_its_input_reads_another_sheet(tmp_path):
    """読む行が同じ区切りでも、その先の数式が別のシートを読むとき：計算の誤りと言わない。"""
    p = _unit_price_book(tmp_path / "a.xlsx")
    doc = lazy.open_doc(str(p), force=True)
    state = doc.audit_begin()
    while not state.get("done"):
        state = doc.audit_step(state)
    assert not state.get("formula_problems"), state["formula_problems"]
    check = next(c for c in state["checks"] if c["name"] == "数式の計算結果")
    assert check["ok"]
    assert state["formulas_ok"] >= 2  # B2 と D2 は計算して照らす（D2 は C2 の保存された値を使う）
    assert state["formulas_cross"] == 1  # C2 は照らしていないと数える


def test_audit_finds_a_real_difference_above_a_cross_sheet_formula(tmp_path):
    """別のシートを読む数式を読む数式でも、保存された値が本当に違えば、見つける。"""
    data = ("<row r=\"2\"><c r=\"C2\"><f>'ほか'!A500</f><v>100</v></c>"
            '<c r="D2"><f>C2+1</f><v>5</v></c></row>') + _filler(3, 400)
    other = _filler(1, 499) + '<row r="500"><c r="A500"><v>100</v></c></row>'
    p = _book(tmp_path / "a.xlsx", {"データ": data, "ほか": other})
    doc = lazy.open_doc(str(p), force=True)
    state = doc.audit_begin()
    while not state.get("done"):
        state = doc.audit_step(state)
    assert any("D2" in x and "101" in x for x in state.get("formula_problems", [])), state.get("formula_problems")


def test_fill_keeps_the_formula_that_reads_a_row_it_did_not_load(tmp_path):
    """連動する数式の先の数式が別のシートを読むとき：その値を 0 にして書かない（金額が変わる）。"""
    p = _unit_price_book(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    r = api.fill(str(p), {"データ!A2": 20}, str(out))
    assert r["streamed"] and r["proof_ok"], r
    assert r["recalculated"] == {"データ!B2": 40.0, "データ!D2": 60.0}, r["recalculated"]
    sh = Workbook(out.read_bytes(), "o.xlsx").sheet("データ")
    assert sh.cells[(2, 3)].value == 100  # 別のシートを読む数式は、そのまま
    assert sh.cells[(2, 4)].value == 60


def test_calculate_with_a_write_uses_the_stored_value_or_the_real_one(tmp_path):
    p = _unit_price_book(tmp_path / "a.xlsx")
    got = {c["cell"]: c for c in api.calculate(str(p), {"データ!A2": 20}, cells=["データ!D2", "データ!C2", "データ!B2"])["cells"]}
    assert got["データ!B2"]["value"] == 40
    assert got["データ!C2"]["value"] == 100
    assert got["データ!D2"]["value"] == 60


def test_calculate_wide_range_of_formulas_is_not_summed_as_zero(tmp_path):
    """広い範囲（読むセルをさかのぼらない広さ）の中の数式が別のシートを読む：0 と足さない。"""
    n = 4200
    rows = "".join(f"<row r=\"{r}\"><c r=\"A{r}\"><f>'ほか'!A{r}</f><v>1</v></c></row>" for r in range(1, n + 1))
    data = rows + f'<row r="{n + 1}"><c r="B{n + 1}"><f>SUM(A1:A{n})</f><v>{n}</v></c></row>'
    other = "".join(f'<row r="{r}"><c r="A{r}"><v>1</v></c></row>' for r in range(1, n + 1))
    p = _book(tmp_path / "a.xlsx", {"データ": data, "ほか": other})
    got = api.calculate(str(p), {}, cells=[f"データ!B{n + 1}"])["cells"][0]
    assert got.get("status") != "ok" or got["value"] == n, got


def _named_book(path):
    """D2 は名前（税率）と B2 を読む。名前の指す行は小さなブックに入らない。"""
    data = ('<row r="2"><c r="A2"><v>10</v></c><c r="B2"><f>A2*2</f><v>20</v></c>'
            '<c r="D2"><f>B2*税率</f><v>2</v></c><c r="E2"><f>D2+1</f><v>3</v></c></row>') + _filler(3, 400)
    other = _filler(1, 499) + '<row r="500"><c r="A500"><v>0.1</v></c></row>'
    names = "<definedNames><definedName name=\"税率\">'ほか'!$A$500</definedName></definedNames>"
    return _book(path, {"データ": data, "ほか": other}, names)


def test_fill_says_a_formula_with_a_name_was_not_recalculated(tmp_path):
    """名前の先を読まなかった数式：0 にせず、いまの値が書いた値を反映していないと伝える（連動の連動まで）。"""
    p = _named_book(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    r = api.fill(str(p), {"データ!A2": 20}, str(out))
    assert r["recalculated"] == {"データ!B2": 40.0}, r["recalculated"]
    stale = " ".join(r.get("stale_formulas") or [])
    assert "データ!D2" in stale or "D2" in stale, r
    sh = Workbook(out.read_bytes(), "o.xlsx").sheet("データ")
    assert sh.cells[(2, 4)].value == 2 and sh.cells[(2, 5)].value == 3


def test_kept_values_are_named_in_the_answer(tmp_path):
    n = 4200
    rows = "".join(f"<row r=\"{r}\"><c r=\"A{r}\"><f>'ほか'!A{r}</f><v>1</v></c></row>" for r in range(1, n + 1))
    data = rows + f'<row r="{n + 1}"><c r="B{n + 1}"><f>SUM(A1:A{n})</f><v>{n}</v></c></row>'
    other = "".join(f'<row r="{r}"><c r="A{r}"><v>1</v></c></row>' for r in range(1, n + 1))
    p = _book(tmp_path / "a.xlsx", {"データ": data, "ほか": other})
    got = api.calculate(str(p), {}, cells=[f"データ!B{n + 1}"])["cells"][0]
    assert got["value"] == n and "保存された値" in got["note"]


def test_whole_workbook_still_calculates_everything(tmp_path, monkeypatch):
    """小さなブックでないときは、いままでどおり全部を計算する（保存された値に頼らない）。"""
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 10 ** 9)
    lazy.clear()
    p = _unit_price_book(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    r = api.fill(str(p), {"データ!A2": 20}, str(out))
    assert not r.get("streamed")
    assert r["recalculated"] == {"データ!B2": 40.0, "データ!D2": 60.0}, r["recalculated"]
