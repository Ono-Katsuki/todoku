"""Excel の証明：「書いた所と、その計算結果のほかは変えていない」。わざと壊したファイルで落ちること、連動の求め方、確かめた範囲の数。"""

from __future__ import annotations

import re
import time
import zipfile

import pytest

from todoku import api, jobs, lazy, limits
from todoku.book import open_book
from todoku.formula.parser import parse_cell
from todoku.xlsx import edit as xedit
from todoku.xlsx import reach as xreach
from todoku.xlsx.proof import prove
from todoku.xlsx.workbook import Workbook

XL = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PK = "http://schemas.openxmlformats.org/package/2006/relationships"


@pytest.fixture(autouse=True)
def clean(monkeypatch, tmp_path):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    jobs._purged.clear()
    lazy.clear()
    api._WINDOWS.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()


def _book(path, sheets: dict[str, str], names: str = "", tables: dict[int, str] | None = None, calc: str = ""):
    """sheets：シートの名前 → sheetData の中身（<row> の並び）。tables：シートの番号 → テーブルの部品の中身。"""
    tables = tables or {}
    ct = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>')
    sh = "".join(f'<sheet name="{n}" sheetId="{i + 1}" r:id="rId{i + 1}"/>' for i, n in enumerate(sheets))
    wb = f'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{XL}" xmlns:r="{REL}"><sheets>{sh}</sheets>{names}{calc}</workbook>'
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
            parts = f'<tableParts count="1"><tablePart r:id="rId1"/></tableParts>' if i in tables else ""
            z.writestr(f"xl/worksheets/sheet{i + 1}.xml", f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{XL}" xmlns:r="{REL}"><sheetData>{body}</sheetData>{parts}</worksheet>')
            if i in tables:
                z.writestr(f"xl/tables/table{i + 1}.xml", tables[i])
                z.writestr(f"xl/worksheets/_rels/sheet{i + 1}.xml.rels",
                           f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{PK}"><Relationship Id="rId1" Type="{REL}/table" Target="../tables/table{i + 1}.xml"/></Relationships>')
    return path


def _filler(first: int, n: int) -> str:
    return "".join(f'<row r="{r}"><c r="A{r}"><v>{r}</v></c><c r="B{r}" t="inlineStr"><is><t>埋め草の字 {r}</t></is></c></row>' for r in range(first, first + n))


def _reach(path, written):
    wb = Workbook(path.read_bytes(), path.name)
    g = xreach.from_workbook(wb)
    return g, g.reach(written)


def _kind(g, reach, sheet, a1):
    r, c = parse_cell(a1)
    return reach.kind(g.key(sheet, r, c))


def _tamper(src, dst, part: str, old: str, new: str) -> None:
    with zipfile.ZipFile(src) as a, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as b:
        for i in a.infolist():
            data = a.read(i.filename)
            if i.filename == part:
                text = data.decode("utf-8")
                assert old in text, text[:400]
                data = text.replace(old, new).encode("utf-8")
            b.writestr(i, data)


# ------------------------------------------------------------------ 連動の求め方
def test_reach_follows_other_sheets_ranges_names_and_chains(tmp_path):
    body1 = ('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>5</v></c></row>'
             '<row r="2"><c r="C2"><f>A1*2</f><v>2</v></c><c r="D2"><f>B1*2</f><v>10</v></c></row>'
             '<row r="3"><c r="C3"><f>SUM(A1:B1)</f><v>6</v></c><c r="D3"><f>入力値+1</f><v>2</v></c><c r="E3"><f>B1</f><v>5</v></c></row>')
    body2 = ("<row r=\"1\"><c r=\"A1\"><f>Sheet1!C2+1</f><v>3</v></c><c r=\"B1\"><f>Sheet1!D2+1</f><v>11</v></c>"
             '<c r="C1"><f>A1+1</f><v>4</v></c></row>')
    p = _book(tmp_path / "a.xlsx", {"Sheet1": body1, "Sheet2": body2},
              names='<definedNames><definedName name="入力値">Sheet1!$A$1</definedName></definedNames>')
    g, r = _reach(p, {("Sheet1", 1, 1)})
    assert _kind(g, r, "Sheet1", "A1") == "written"
    for sheet, a1 in (("Sheet1", "C2"), ("Sheet1", "C3"), ("Sheet1", "D3"), ("Sheet2", "A1"), ("Sheet2", "C1")):
        assert _kind(g, r, sheet, a1) == "dependent", (sheet, a1)
    for sheet, a1 in (("Sheet1", "D2"), ("Sheet1", "E3"), ("Sheet1", "B1"), ("Sheet2", "B1")):
        assert _kind(g, r, sheet, a1) is None, (sheet, a1)


def test_reach_tables_shared_formulas_and_arrays(tmp_path):
    tbl = (f'<table xmlns="{XL}" id="1" name="表1" displayName="表1" ref="A1:B3"><tableColumns count="2">'
           '<tableColumn id="1" name="金額"/><tableColumn id="2" name="数"/></tableColumns></table>')
    body = ('<row r="1"><c r="A1" t="inlineStr"><is><t>金額</t></is></c><c r="B1" t="inlineStr"><is><t>数</t></is></c></row>'
            '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>7</v></c></row><row r="3"><c r="A3"><v>2</v></c><c r="B3"><v>8</v></c></row>'
            '<row r="5"><c r="D5"><f>SUM(表1[金額])</f><v>3</v></c><c r="E5"><f>SUM(表1[数])</f><v>15</v></c></row>'
            '<row r="6"><c r="D6"><f t="shared" ref="D6:D7" si="0">A2*2</f><v>2</v></c></row>'
            '<row r="7"><c r="D7"><f t="shared" si="0"/><v>4</v></c></row>'
            '<row r="8"><c r="F8"><f t="array" ref="F8:F9">A3*{1;2}</f><v>2</v></c></row>'
            '<row r="9"><c r="F9"><v>4</v></c></row>')
    p = _book(tmp_path / "a.xlsx", {"S": body}, tables={0: tbl})
    g, r = _reach(p, {("S", 3, 1)})  # A3 を書く
    assert _kind(g, r, "S", "D5") == "dependent"
    assert _kind(g, r, "S", "E5") is None
    assert _kind(g, r, "S", "D7") == "dependent"  # 共有数式の子は A3 を読む
    assert _kind(g, r, "S", "D6") is None  # 親は A2 を読む
    assert _kind(g, r, "S", "F8") == "dependent"
    assert _kind(g, r, "S", "F9") == "dependent"  # 配列数式の結果の範囲


def test_reach_sumif_reads_the_size_of_its_first_range(tmp_path):
    """SUMIF の 3 つ目は左上だけが効き、読む大きさは 1 つ目と同じ（字の参照より広い）。"""
    body = ('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>1</v></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row>'
            '<row r="4"><c r="D4"><f>SUMIF(A1:A2,1,B1)</f><v>3</v></c></row>')
    p = _book(tmp_path / "a.xlsx", {"S": body})
    g, r = _reach(p, {("S", 2, 2)})
    assert _kind(g, r, "S", "D4") == "dependent"


@pytest.mark.parametrize("formula", ["INDIRECT(&quot;A1&quot;)", "OFFSET(B1,0,-1)", "A1:INDEX(A1:A2,1)", "相対+1"])
def test_reach_does_not_guess_when_the_reference_is_computed(tmp_path, formula):
    """読む先が計算で決まる数式は、連動とも非連動とも決めない。その数式を読む数式も。"""
    body = ('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>9</v></c></row>'
            f'<row r="2"><c r="C2"><f>{formula}</f><v>1</v></c><c r="D2"><f>C2+1</f><v>2</v></c><c r="E2"><f>B1</f><v>9</v></c></row>')
    p = _book(tmp_path / "a.xlsx", {"S": body},
              names='<definedNames><definedName name="相対">S!B1</definedName></definedNames>')
    g, r = _reach(p, {("S", 1, 2)} if "OFFSET" not in formula else {("S", 1, 1)})
    assert _kind(g, r, "S", "C2") == "dynamic"
    assert _kind(g, r, "S", "D2") == "dynamic"
    assert g.why


def test_reach_marks_cells_that_change_each_open(tmp_path):
    body = ('<row r="1"><c r="A1"><v>1</v></c></row>'
            '<row r="2"><c r="C2"><f>TODAY()</f><v>46000</v></c><c r="D2"><f>C2+1</f><v>46001</v></c><c r="E2"><f>A1+TODAY()</f><v>46001</v></c></row>')
    p = _book(tmp_path / "a.xlsx", {"S": body})
    g, r = _reach(p, {("S", 1, 1)})
    assert _kind(g, r, "S", "C2") == "timed"
    assert _kind(g, r, "S", "D2") == "timed"
    assert _kind(g, r, "S", "E2") == "dependent"  # 書いたセルを読むので連動する（値は開くたびにも変わる）


def test_lazy_graph_agrees_with_the_workbook_graph(tmp_path, monkeypatch):
    """大きなブックの道（XML を行ごとに読む）と、全部を読んだブックの道で、同じ連動になる。"""
    body = ('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>5</v></c></row>'
            '<row r="6"><c r="D6"><f t="shared" ref="D6:D7" si="0">A1*2</f><v>2</v></c><c r="E6"><f>B1</f><v>5</v></c></row>'
            '<row r="7"><c r="D7"><f t="shared" si="0"/><v>0</v></c></row>'
            '<row r="8"><c r="F8"><f t="array" ref="F8:F9">A1*{1;2}</f><v>1</v></c></row><row r="9"><c r="F9"><v>2</v></c></row>'
            '<row r="10"><c r="G10"><f>INDIRECT(&quot;A1&quot;)</f><v>1</v></c></row>') + _filler(11, 50)
    other = '<row r="1"><c r="A1"><f>S!D6</f><v>2</v></c><c r="B1"><f>S!E6</f><v>5</v></c></row>'
    p = _book(tmp_path / "a.xlsx", {"S": body, "T": other})
    g1, r1 = _reach(p, {("S", 1, 1)})
    doc = lazy.open_doc(str(p), force=True)
    g2 = xreach.from_lazy(doc)
    r2 = g2.reach({("S", 1, 1)})
    names = lambda g, keys: {g.unkey(k) for k in keys}  # noqa: E731
    assert names(g1, r1.dependent) == names(g2, r2.dependent)
    assert names(g1, r1.dynamic) == names(g2, r2.dynamic)
    assert ("T", 1, 1) in names(g2, r2.dependent) and ("T", 1, 2) not in names(g2, r2.dependent)


# ------------------------------------------------------------------ わざと壊したファイル（全部を読んだブックの道）
def _small(path):
    body = ('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>5</v></c></row>'
            '<row r="2"><c r="C2"><f>A1*2</f><v>2</v></c><c r="D2"><f>Sheet2!A1+1</f><v>8</v></c></row>')
    return _book(path, {"Sheet1": body, "Sheet2": '<row r="1"><c r="A1"><v>7</v></c></row>'})


def _whole(original, new, seeds=frozenset({("Sheet1", 1, 1)})):
    a, b = original.read_bytes(), new.read_bytes()
    pr = prove(a, b)
    return xreach.prove_whole(Workbook(a, original.name), b, new.name, set(seeds), pr.cell_changes)


def test_honest_write_counts_the_scope(tmp_path):
    p = _small(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    res = api.fill(str(p), {"Sheet1!A1": 10}, str(out))
    assert res["proof_ok"], res
    scope = res["proof_scope"]
    assert scope["not_dependent"]["cells"] == scope["not_dependent"]["same"] == 3  # B1・D2・Sheet2!A1
    assert scope["written"]["cells"] == 1
    assert scope["dependent"]["cells"] == 1 and scope["dependent"]["changed"] == 1
    assert scope["recalculated"]["status"] == "done"
    assert scope["recalculated"]["checked"] == scope["recalculated"]["same"] == 1
    assert scope["undetermined"]["cells"] == 0
    assert "連動しないセル" in res["proof"]


def test_zeroing_an_unconnected_formula_fails(tmp_path):
    """実物の誤りの形：別のシートを読む、連動しない数式の値が 0 になったファイル。"""
    p = _small(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    assert api.fill(str(p), {"Sheet1!A1": 10}, str(out))["proof_ok"]
    bad = tmp_path / "bad.xlsx"
    _tamper(out, bad, "xl/worksheets/sheet1.xml", "<v>8</v>", "<v>0</v>")
    scope = _whole(p, bad)
    assert any("D2" in x and "連動しない" in x for x in scope.problems), scope.problems
    assert scope.not_dependent_same == scope.not_dependent - 1


def test_shifting_a_connected_formula_fails(tmp_path):
    p = _small(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    assert api.fill(str(p), {"Sheet1!A1": 10}, str(out))["proof_ok"]
    bad = tmp_path / "bad.xlsx"
    _tamper(out, bad, "xl/worksheets/sheet1.xml", "<v>20</v>", "<v>21</v>")
    scope = _whole(p, bad)
    assert any("C2" in x and "計算し直した" in x for x in scope.problems), scope.problems
    assert scope.calc["checked"] == 1 and scope.calc["same"] == 0


def test_a_writer_that_breaks_an_unconnected_formula_fails_the_fill(tmp_path, monkeypatch):
    """書き手に誤りがあるとき：fill_document の結果が不合格になり、提出してはいけない名前で書き出される。"""
    real = xedit.recalc

    def broken(wb, overrides):
        changed, problems = real(wb, overrides)
        changed[("Sheet1", 2, 4)] = 0.0  # D2 を 0 に
        return changed, problems

    monkeypatch.setattr(xedit, "recalc", broken)
    p = _small(tmp_path / "a.xlsx")
    res = api.fill(str(p), {"Sheet1!A1": 10}, str(tmp_path / "o.xlsx"))
    assert res["proof_ok"] is False
    assert any("D2" in x for x in res["unexpected"]), res["unexpected"]
    assert ".unverified" in res["out_path"]


def test_a_writer_that_shifts_a_connected_formula_fails_the_fill(tmp_path, monkeypatch):
    real = xedit.recalc

    def broken(wb, overrides):
        changed, problems = real(wb, overrides)
        changed[("Sheet1", 2, 3)] = 21.0  # C2 は 20
        return changed, problems

    monkeypatch.setattr(xedit, "recalc", broken)
    p = _small(tmp_path / "a.xlsx")
    res = api.fill(str(p), {"Sheet1!A1": 10}, str(tmp_path / "o.xlsx"))
    assert res["proof_ok"] is False
    assert any("C2" in x for x in res["unexpected"]), res["unexpected"]


def test_cannot_calculate_is_reported_not_filled(tmp_path):
    """計算できない関数：0 や推測で埋めず、計算できないと数える。"""
    body = ('<row r="1"><c r="A1"><v>1</v></c></row>'
            '<row r="2"><c r="C2" t="str"><f>_xlfn.NOSUCHFUNCTION(A1)</f><v>もとの字</v></c></row>')
    p = _book(tmp_path / "a.xlsx", {"S": body})
    res = api.fill(str(p), {"S!A1": 10}, str(tmp_path / "o.xlsx"))
    assert res["proof_scope"]["recalculated"]["cannot_calculate"] == 1
    assert res["proof_scope"]["recalculated"]["checked"] == 0
    wb = Workbook((tmp_path / "o.xlsx").read_bytes())
    assert wb.sheet("S").cells[(2, 3)].value == "もとの字"


def test_undetermined_is_reported_with_counts(tmp_path):
    body = ('<row r="1"><c r="A1"><v>1</v></c></row>'
            '<row r="2"><c r="C2"><f>INDIRECT(&quot;A1&quot;)</f><v>1</v></c><c r="D2"><f>TODAY()</f><v>1</v></c></row>')
    p = _book(tmp_path / "a.xlsx", {"S": body})
    res = api.fill(str(p), {"S!A1": 10}, str(tmp_path / "o.xlsx"))
    und = res["proof_scope"]["undetermined"]
    assert und["cells"] == 2
    assert und["reads_changing_refs"] == 1 and und["changes_each_open"] == 1
    assert "INDIRECT" in und["why"]
    assert "連動を決められない" in res["proof"]


def test_a_missed_connection_is_found_by_the_whole_calc(tmp_path, monkeypatch):
    """連動の求め方に見落としがあっても：計算し直すと書いた値で変わるセルが、値を変えないまま残っていれば不合格。"""
    p = _small(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    assert api.fill(str(p), {"Sheet1!A1": 10}, str(out))["proof_ok"]
    a, b = p.read_bytes(), out.read_bytes()
    wb = Workbook(a)
    g = xreach.from_workbook(wb)
    r = g.reach({("Sheet1", 1, 1)})
    r.dependent.clear()  # C2 を連動しないと誤って求めた
    s = xreach.changes_scope(wb, g, r, [])
    calc, bad = xreach.whole_calc(Workbook(xreach_stale(b)), g, r, s.changed, False, lambda: wb)
    assert any("C2" in x and "見落とし" in x for x in bad), (calc, bad)


def xreach_stale(b: bytes) -> bytes:
    """C2 の値を元の 2 に戻したブック（書いた値を反映していない）。"""
    import io

    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(b)) as a, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for i in a.infolist():
            data = a.read(i.filename)
            if i.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"<v>20</v>", b"<v>2</v>")
            z.writestr(i, data)
    return out.getvalue()


@pytest.mark.xfail(strict=True, reason="書き手が、連動しない数式も計算と違えば書いている（次の修正で直す）")
def test_an_already_different_unconnected_formula_does_not_fail(tmp_path):
    """連動しない数式の保存された値が、元から計算と違うとき：書き込みの誤りではない。別に数え、値はそのまま。"""
    body = ('<row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>5</v></c></row>'
            '<row r="2"><c r="C2"><f>B1*2</f><v>11</v></c></row>')  # 計算すると 10
    p = _book(tmp_path / "a.xlsx", {"S": body})
    res = api.fill(str(p), {"S!A1": 10}, str(tmp_path / "o.xlsx"))
    assert res["proof_ok"], res["unexpected"]
    assert res["proof_scope"]["recalculated"]["others_already_different"] == 1
    assert Workbook((tmp_path / "o.xlsx").read_bytes()).sheet("S").cells[(2, 3)].value == 11


# ------------------------------------------------------------------ 大きなブックの道
def _big(path):
    """連動しない、別のシートを読む数式（C2）と、連動する数式（D2）。読まない行がたくさんある。"""
    data = ('<row r="2"><c r="A2"><v>10</v></c>'
            "<c r=\"C2\"><f>'ほか'!A500</f><v>100</v></c>"
            '<c r="D2"><f>A2*2</f><v>20</v></c></row>') + _filler(3, 400)
    other = _filler(1, 499) + '<row r="500"><c r="A500"><v>100</v></c></row>'
    return _book(path, {"データ": data, "ほか": other})


@pytest.fixture
def small_budget(monkeypatch):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2000)


def _files(original, new, seeds=frozenset({("データ", 2, 1)})):
    doc = lazy.open_doc(str(original), force=True)
    from todoku.lazy.xlsxdoc import XlsxLazy
    from todoku.xlsx.workbook import sst_scope

    with sst_scope(doc.strings):
        g = xreach.from_lazy(doc)
    r = g.reach(set(seeds))
    after = XlsxLazy(str(new), new.name)
    try:
        return xreach.files_scope(doc, after, g, r)
    finally:
        after.close()


def test_big_book_honest_write_says_whole_calc_is_not_yet_checked(tmp_path, small_budget):
    p = _big(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    res = api.fill(str(p), {"データ!A2": 11}, str(out))
    assert res["streamed"] and res["proof_ok"], res
    scope = res["proof_scope"]
    assert scope["recalculated"]["status"] == "deferred"
    assert any("未確認" in x for x in res["deferred"])
    assert "未確認" in res["proof"]
    assert scope["not_dependent"]["cells"] == scope["not_dependent"]["same"] > 1000
    assert scope["dependent"] == {"cells": 1, "changed": 1}
    assert scope["written"]["cells"] == 1


def test_big_book_zeroing_an_unconnected_formula_fails(tmp_path, small_budget):
    p = _big(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    assert api.fill(str(p), {"データ!A2": 11}, str(out))["proof_ok"]
    bad = tmp_path / "bad.xlsx"
    _tamper(out, bad, "xl/worksheets/sheet1.xml", "<v>100</v>", "<v>0</v>")
    scope = _files(p, bad)
    assert any("C2" in x and "連動しない" in x for x in scope.problems), scope.problems


def test_big_book_writer_that_zeroes_an_unconnected_formula_fails_the_fill(tmp_path, small_budget, monkeypatch):
    """修正前の誤りの形：読まなかった行を空に見て、連動しない数式を 0 にした書き手。"""
    real = xedit.recalc

    def broken(wb, overrides):
        changed, problems = real(wb, overrides)
        changed[("データ", 2, 3)] = 0.0
        return changed, problems

    monkeypatch.setattr(xedit, "recalc", broken)
    p = _big(tmp_path / "a.xlsx")
    res = api.fill(str(p), {"データ!A2": 11}, str(tmp_path / "o.xlsx"))
    assert res["proof_ok"] is False
    assert any("C2" in x and "連動しない" in x for x in res["unexpected"]), res["unexpected"]
    assert ".unverified" in res["out_path"]


def test_big_book_changing_a_formula_text_fails(tmp_path, small_budget):
    p = _big(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    assert api.fill(str(p), {"データ!A2": 11}, str(out))["proof_ok"]
    bad = tmp_path / "bad.xlsx"
    _tamper(out, bad, "xl/worksheets/sheet1.xml", "<f>A2*2</f>", "<f>A2*3</f>")
    scope = _files(p, bad)
    assert any("D2" in x and "数式が変わりました" in x for x in scope.problems), scope.problems


def _wait_proof(path, tries=200):
    for _ in range(tries):
        d = api.audit(str(path))
        wp = d.get("written_proof")
        if wp and wp["status"] not in ("running",) and d["job"].get("finished", True):
            return d
        time.sleep(0.1)
    raise AssertionError("終わりませんでした")


def test_big_book_whole_calc_runs_as_a_long_job_from_audit(tmp_path, small_budget):
    """長い仕事：結果が出るまで「未確認」、終わったら数を返す。"""
    p = _big(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    assert api.fill(str(p), {"データ!A2": 11}, str(out))["proof_ok"]
    assert jobs.proof_status(str(out)) is not None
    first = api.audit(str(out))
    wp = first["written_proof"]
    if wp["status"] == "running":
        assert any("未確認" in x for x in first["deferred"])
    d = _wait_proof(out)
    assert d["written_proof"]["status"] == "done"
    calc = d["written_proof"]["recalculated"]
    assert calc["checked"] == calc["same"] == 1, calc
    assert any(c["name"].startswith("書いた後の計算結果") and c["ok"] for c in d["checks"])


def test_big_book_whole_calc_job_catches_a_shifted_value(tmp_path, small_budget):
    p = _big(tmp_path / "a.xlsx")
    out = tmp_path / "o.xlsx"
    res = api.fill(str(p), {"データ!A2": 11}, str(out))
    assert res["proof_ok"]
    bad = tmp_path / "bad.xlsx"
    _tamper(out, bad, "xl/worksheets/sheet1.xml", "<v>22</v>", "<v>23</v>")
    jobs.note_proof(str(bad), str(p), [("データ", 2, 1)], [("データ", 2, 4)], False)
    d = _wait_proof(bad)
    assert d["written_proof"]["status"] == "done"
    assert any("D2" in x for x in d["written_proof"]["problems"])
    assert d["ok"] is False


def test_whole_calc_deadline_is_not_a_pass(tmp_path):
    p = _small(tmp_path / "a.xlsx")
    wb = Workbook(p.read_bytes())
    g = xreach.from_workbook(wb)
    r = g.reach({("Sheet1", 1, 1)})
    calc, bad = xreach.whole_calc(wb, g, r, set(), False, None, deadline=time.monotonic() - 1)
    assert calc["status"] == "deferred" and not bad


def test_xls_write_keeps_the_scope(tmp_path):
    """古い Excel の道でも、同じ確かめが動く（ブックの作りが違っても、セルの持ち方は同じ）。"""
    from todoku.xlsx import reach

    p = _small(tmp_path / "a.xlsx")
    wb = open_book(p.read_bytes(), p.name)
    assert reach.from_workbook(wb).reach({("Sheet1", 1, 1)}).dependent


def test_reach_module_has_no_keyword_lists_about_documents():
    """連動の判断は数式の字の作りだけで決める（見出しの字は見ない）。"""
    src = (xreach.__file__ and open(xreach.__file__, encoding="utf-8").read())
    assert not re.search(r"(?:label|context)\b", src)
