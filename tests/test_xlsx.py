"""合格条件（構想書「検証設計」）をそのままテストにする。

往復：変換して戻したファイルのSHA-256が元と一致
編集：元との差分が、入力したセルとそれに依存する数式セルの値だけに収まる
関数：全セルで保存済みの計算結果（Excelが計算した値）と一致
記入：空の様式に記入例の入力を書き込むと、記入例でExcelが計算した値と全数で一致する
互換：書き出したファイルが、Excelが修復なしに開くための構造の条件を満たす

ほかの表計算ソフトは正解に使わない。正解はいつも、Excelが保存した値。
公開様式（corpus/forms）は scripts/fetch_forms.py で取得する。ないときは、そのテストを飛ばす。
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from todoku.formula.parser import parse_cell
from todoku.xlsx.calc import values_equal, verify
from todoku.xlsx.diagnose import diagnose
from todoku.xlsx.edit import EditError, bake, fill
from todoku.xlsx.lint import lint
from todoku.xlsx.proof import prove
from todoku.xlsx import window
from todoku.xlsx.workbook import Workbook, load

FX = Path(__file__).parent / "fixtures"
CORPUS = Path(__file__).parent.parent / "corpus"
KAMI = FX / "kami_meisai.xlsx"
MANIFEST = json.loads((Path(__file__).parent / "forms.json").read_text(encoding="utf-8"))["forms"]


def corpus_files():
    """コーパスの Excel のブック（.xlsx・.xlsm）。PK で始まっても Word の文書（記入例の組・暗号の試験の中身）は除く。"""
    if not CORPUS.exists():
        return []
    # world_tmp は世界調査の一時の置き場（調査が書いては消す）なので、試験の対象にしない
    # 名前が .xlsx でも中身が Excel でないもの（目録が形式を取り違えた ZIP など）は、中身で除く
    from todoku.book import kind_of

    return [p for p in sorted(CORPUS.rglob("*")) if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm")
            and "world_tmp" not in p.parts and kind_of(p.read_bytes(), p.name) == "xlsx"]


# ------------------------------------------------------------------ 往復
def example_pairs():
    pairs = {}
    for f in MANIFEST:
        if f["pair"]:
            pairs.setdefault(f["pair"], {})[f["role"]] = CORPUS / "forms" / f"{f['id']}.xlsx"
    return sorted(pairs.items())


@pytest.mark.parametrize("path", [KAMI] + corpus_files(), ids=lambda p: p.name)
def test_roundtrip_identical(path):
    blob = path.read_bytes()
    wb = Workbook(blob)
    assert hashlib.sha256(wb.zip.build()).digest() == hashlib.sha256(blob).digest()


def test_window_apply_without_edit_is_identical(tmp_path):
    from todoku.cli import main

    wb = load(str(KAMI))
    w = tmp_path / "w.json"
    import json

    w.write_text(json.dumps(window.build(wb), ensure_ascii=False, default=str), encoding="utf-8")
    out = tmp_path / "o.xlsx"
    assert main(["apply", str(w), "--original", str(KAMI), "-o", str(out)]) == 0
    assert out.read_bytes() == KAMI.read_bytes()


# ------------------------------------------------------------------ 関数
@pytest.mark.parametrize("path", corpus_files(), ids=lambda p: p.name)
def test_formulas_match_cached(path):
    """Excelで保存された実物の書類で、全数式がExcelの保存した計算結果と一致する。"""
    rep = verify(Workbook(path.read_bytes()))
    bad = [r for r in rep.results if r.status == "mismatch"]
    assert not bad, bad[:5]


@pytest.mark.parametrize("name,pair", example_pairs(), ids=lambda x: x if isinstance(x, str) else "")
def test_fill_matches_excel_example(name, pair):
    """空の様式に記入例の入力を書き込み、記入例でExcelが計算した値と全数で比べる。"""
    from todoku.xlsx.example import check

    if not all(p.exists() for p in pair.values()):
        pytest.skip("公開様式がありません（scripts/fetch_forms.py で取得）")
    r = check(pair["blank"].read_bytes(), pair["example"].read_bytes())
    assert r.inputs > 50 and r.formulas > 1000
    assert not r.mismatches, r.mismatches[:5]
    assert r.proof.ok, (r.proof.unexpected[:5], r.proof.lint_errors[:5])


@pytest.mark.parametrize("path", [KAMI] + corpus_files(), ids=lambda p: p.name)
def test_lint_has_no_false_positive_on_originals(path):
    """Excelで保存されたファイルには、構造の指摘が出ない（検査の誤りがないことの確認）。"""
    assert lint(path.read_bytes()).errors == []


def test_baked_fixture_values():
    """フィクスチャの計算結果（とどくが書き込んだ値）を、手計算の値と照らす。"""
    sh = load(str(KAMI)).sheet("明細書")
    v = lambda a: sh.cells[parse_cell(a)].value  # noqa: E731
    assert v("AA4") == 10.9  # 3級地
    assert v("I5") == 46296  # 2026年10月1日
    assert [v("S8"), v("S9"), v("S10")] == [663 * 20, 21 * 38, 30 * 18]
    assert v("S11") == ""  # 空欄の行は "" を返す
    assert v("S14") == 14598
    assert v("S15") == 6102  # ROUND(14598*0.418=6101.964)
    assert v("S16") == 20700
    assert v("S17") == 225630  # ROUNDDOWN(20700*10.9)
    assert v("S19") == 203067  # ROUNDDOWN(225630*0.9)
    assert v("S20") == 9300  # MIN(22563, 上限月額)
    assert v("S21") == "差額あり：13,263円"
    assert v("S22") == "令和8年10月"
    assert [v("S23"), v("S24"), v("S25")] == [2, 798 + 540, 192.08]


def test_bake_fills_missing_values(tmp_path):
    from openpyxl import Workbook as OWorkbook

    ow = OWorkbook()
    ws = ow.active
    ws["A1"], ws["A2"], ws["A3"] = 1234, 0.08, "=ROUNDDOWN(A1*A2,0)"
    ws["B1"] = '=TEXT(A3,"#,##0")&"円"'
    p = tmp_path / "nocache.xlsx"
    ow.save(p)
    wb = Workbook(p.read_bytes())
    assert verify(wb).summary()["no_cache"] == 2
    res = bake(wb)
    baked = Workbook(res.blob)
    assert baked.sheets[0].cells[(3, 1)].value == 98
    assert baked.sheets[0].cells[(1, 2)].value == "98円"
    assert verify(baked).summary()["match"] == 2
    pr = prove(p.read_bytes(), res.blob, set())
    assert pr.ok and pr.lint_errors == []


@pytest.mark.parametrize("expr,expected", [
    ("ROUND(2.675,2)", 2.68), ("ROUND(-2.5,0)", -3.0), ("ROUNDDOWN(-3.675,2)", -3.67), ("INT(-3.5)", -4.0),
    ("MOD(-7,3)", 2.0), ("-2^2", 4.0), ("0.1+0.2=0.3", True), ('"a"="A"', True), ("1<\"a\"", True),
    ('TEXT(1234567.891,"#,##0.00")', "1,234,567.89"), ('TEXT(DATE(2019,5,1),"ggge年m月d日")', "令和1年5月1日"),
    ('TEXT(DATE(2019,4,30),"ge.m.d")', "H31.4.30"), ("DATE(2026,14,35)", 46453.0), ("WEEKDAY(DATE(2026,11,2))", 2.0),
    ("ROUNDDOWN(10.9*1234,0)", 13450.0), ('IFERROR(1/0,"x")', "x"), ('SUBSTITUTE("a-b-c","-","+",2)', "a-b+c"),
    ('"a"&1.5&TRUE', "a1.5TRUE"), ("ROUND(1234.5678,-2)", 1200.0), ("MOD(5.5,2)", 1.5), ("1/3*3", 1.0),
    ("TRUNC(-3.5)", -3.0), ("ROUNDUP(3.001,0)", 4.0), ("CEILING(2.1,0.5)", 2.5), ("FLOOR(2.9,0.5)", 2.5),
    ('VALUE("1,234")', 1234.0), ('LEFT("とどく書類",2)', "とど"), ('MID("とどく書類",3,2)', "く書"),
    ('TRIM("  a   b  ")', "a b"), ('TEXT(0.256,"0.0%")', "25.6%"), ('TEXT(-1234.5,"#,##0")', "-1,235"),
    ("DATE(1900,3,1)", 61.0), ("DAY(DATE(2024,2,29))", 29.0), ('DATEDIF(DATE(1989,1,7),DATE(2024,2,29),"Y")', 35.0),
    ("EOMONTH(DATE(2019,4,30),1)", 43616.0), ("NETWORKDAYS(DATE(2026,11,1),DATE(2026,11,30))", 21.0),
])
def test_expressions(expr, expected):
    from todoku.formula.evaluator import Evaluator
    from todoku.formula.parser import parse
    from todoku.xlsx.calc import Calc

    wb = load(str(KAMI))
    v = Evaluator(Calc(wb), wb.sheets[0].name, 1, 1).run(parse(expr))
    assert values_equal(expected, v), (expr, v)


# ------------------------------------------------------------------ 編集
EDITS = {"明細書!O9": "25", "明細書!I4": "7級地", "明細書!A11": "送迎加算Ⅱ", "明細書!I11": "10", "明細書!O11": 5}


def test_fill_changes_only_inputs_and_dependents():
    wb = load(str(KAMI))
    res = fill(wb, EDITS)
    allowed = {(p.sheet, p.address) for p in res.patches if p.kind == "input"}
    pr = prove(KAMI.read_bytes(), res.blob, allowed)
    assert pr.ok, pr.unexpected
    assert pr.entries_changed == ["xl/worksheets/sheet1.xml"]
    assert pr.entries_identical == pr.entries_total - 1
    assert pr.sheets_outside_identical == pr.sheets_changed == 1  # 書いたセル・計算し直したセルのほかはバイト一致
    changed_formulas = {c["cell"] for c in pr.cell_changes if c["kind"] == "formula"}
    assert {"S9", "S11", "S14", "S17", "S19", "AA4"} <= changed_formulas
    # 書き込んだ後のファイルも、全数照合で一致する
    assert verify(Workbook(res.blob)).summary()["mismatch"] == 0


def test_fill_values_are_correct():
    res = fill(load(str(KAMI)), EDITS)
    sh = Workbook(res.blob).sheet("明細書")
    v = lambda a: sh.cells[__import__("todoku.formula.parser", fromlist=["parse_cell"]).parse_cell(a)].value  # noqa: E731
    assert v("AA4") == 10.18  # 7級地
    assert v("S9") == 21 * 25
    assert v("S14") == 663 * 20 + 21 * 25 + 30 * 18 + 10 * 5
    assert v("S15") == round(v("S14") * 0.418)
    assert v("S17") == int(v("S16") * 10.18)


def test_fill_refuses_formula_cells():
    with pytest.raises(EditError):
        fill(load(str(KAMI)), {"明細書!S14": 1})


def test_fill_merged_cell_inside_is_refused_with_anchor_hint():
    # I9:N9 の結合の途中（隠れたセル）：左上に黙って寄せず、どのセルのつもりかを確かめてもらう（第4回 C1）
    with pytest.raises(EditError, match="明細書!I9 に書いてください"):
        fill(load(str(KAMI)), {"明細書!K9": 22})
    res = fill(load(str(KAMI)), {"明細書!I9": 22})
    assert any(p.address == "I9" for p in res.patches)


def test_fill_refuses_label_in_merged_cell():
    with pytest.raises(EditError):
        fill(load(str(KAMI)), {"明細書!B26": "署名"})  # A26:R26 は見出し「印」


def test_validation_warning():
    res = fill(load(str(KAMI)), {"明細書!O8": 40})
    assert any("範囲外" in w for w in res.warnings)
    with pytest.raises(EditError):
        fill(load(str(KAMI)), {"明細書!O8": 40}, strict=True)


def test_fill_refuses_locked_cell_in_protected_sheet():
    # シート保護でロックされたセルは、Excelでも書き込めない
    with pytest.raises(EditError, match="ロック"):
        fill(load(str(KAMI)), {"明細書!AJ30": "追記"})


def test_fill_new_row_and_new_cell():
    # force=True はロックを越えて書く（行やセル要素を新しく作る処理を確かめる）
    res = fill(load(str(KAMI)), {"明細書!AJ30": "追記", "明細書!AH8": "備考"}, force=True)
    pr = prove(KAMI.read_bytes(), res.blob, {("明細書", "AJ30"), ("明細書", "AH8")})
    assert pr.ok, pr.unexpected
    assert pr.sheets_outside_identical == pr.sheets_changed == 1  # 新しく作った行の入れ物のほかは、バイト一致
    sh = Workbook(res.blob).sheet("明細書")
    assert sh.cells[(30, 36)].value == "追記"


def test_fill_several_new_cells_in_the_same_new_row():
    # ない行に2つのセル（列の逆の順で指示）と、同じ位置に入る2つの新しい行：1つの行にまとめ、行・列の順に並べる
    vals = {"明細書!AK60": "後", "明細書!AJ60": "前", "明細書!AJ62": "下", "明細書!AJ61": "上"}
    res = fill(load(str(KAMI)), vals, force=True)
    pr = prove(KAMI.read_bytes(), res.blob, {("明細書", k.split("!")[1]) for k in vals})
    assert pr.ok and not pr.lint_errors, (pr.unexpected, pr.lint_errors)
    sh = Workbook(res.blob).sheet("明細書")
    assert [sh.cells[(60, 36)].value, sh.cells[(60, 37)].value, sh.cells[(61, 36)].value, sh.cells[(62, 36)].value] == \
        ["前", "後", "上", "下"]


def test_proof_finds_a_byte_changed_outside_the_written_cells():
    # セルの中身・書式・結合では分からない変更（印刷の余白・ほかの部品）も、書いたセルの外のバイトの違いとして見つける
    from todoku.zipraw import RawZip

    res = fill(load(str(KAMI)), EDITS)
    allowed = {(p.sheet, p.address) for p in res.patches if p.kind == "input"}
    z = RawZip(res.blob)
    z.replace("xl/worksheets/sheet1.xml", z.read("xl/worksheets/sheet1.xml").replace(b'left="0.75"', b'left="0.76"'))
    pr = prove(KAMI.read_bytes(), z.build(), allowed)
    assert not pr.ok and any("ほかのバイトが変わって" in u for u in pr.unexpected)
    z = RawZip(res.blob)
    z.replace("docProps/app.xml", z.read("docProps/app.xml") + b" ")
    pr = prove(KAMI.read_bytes(), z.build(), allowed)
    assert not pr.ok and any("docProps/app.xml" in u for u in pr.unexpected)


# ------------------------------------------------------------------ 互換（構造）
@pytest.mark.parametrize("edits", [
    EDITS,
    {"明細書!AJ30": "追記", "明細書!AH8": "備考 ", "明細書!A40": "<&>\"'"},  # 新しい行・セル、XMLの特殊文字
    {"明細書!S18": True, "明細書!Y8": "改行\nあり"},
])
def test_filled_file_structure(edits):
    res = fill(load(str(KAMI)), edits, force=True)
    rep = lint(res.blob)
    assert rep.ok, rep.errors


# ------------------------------------------------------------------ 窓と診断
def test_window_lists_inputs_with_labels():
    doc = window.build(load(str(KAMI)))
    cells = {c["cell"]: c for c in doc["sheets"][0]["cells"]}
    assert cells["I4"]["input"] and cells["I4"]["label"] == "地域区分"
    assert cells["O9"]["rules"][0]["type"] == "whole"
    # 表の列は列の見出しで呼び、どの行かは context（その行の並び）で分かる
    assert cells["O9"]["label"] == "回数"
    f = {x["cell"]: x for x in doc["sheets"][0]["fields"]}
    assert "送迎加算Ⅰ" in f["O9"]["context"] and "［▶O9:R9＝38］" in f["O9"]["context"]
    assert cells["S14"]["formula"] == "=SUM(S8:S12)"
    assert doc["sheets"][0]["tables"][0]["columns"] == ["サービス内容", "単位数", "回数", "サービス単位数", "摘要"]


def test_diagnosis_finds_barriers():
    d = diagnose(load(str(KAMI)))
    codes = {f.code for f in d.findings}
    assert {"grid_paper", "spaced_label", "rounding_in_formulas"} <= codes
    # 押印は言葉の一致では決めない（打ち消しも拾うため）。段階の判定で、エージェントが確かめる問いとして出る
    assert "stamp" not in codes
    assert d.score()["reached"] == 1
    from todoku.grade import grade_blob

    g = grade_blob(KAMI.read_bytes(), KAMI.name)
    assert "cannot_do" in [c["id"] for c in g.agent_checks]


def test_drop_calc_chain_keeps_the_meaning_of_other_entries():
    from todoku.xlsx.edit import calc_chain_entries, drop_calc_chain

    xml = b'<calcChain xmlns="x"><c r="B3" i="1" l="1"/><c r="B4"/><c r="C1" i="2"/></calcChain>'
    new, n = drop_calc_chain(xml, {(1, "B3")})
    assert n == 1 and calc_chain_entries(new) == [(1, "B4"), (2, "C1")] and b'r="B4" i="1" l="1"' in new


def test_overwriting_a_formula_removes_only_its_calc_chain_entry():
    # 数式のセルに値を書いたら（マクロが書いたときなど）、calcChain.xml からそのセルの記録だけを取り除く
    from todoku.xlsx.edit import calc_chain_entries

    path = CORPUS / "wide" / "www.city.nagoya.jp" / "r8mousikomisyo.xlsm"
    if not path.exists():
        pytest.skip("corpus なし")
    wb = Workbook(path.read_bytes())
    sid_of = {sh.sheet_id: sh.name for sh in wb.sheets}
    sid, cell = calc_chain_entries(wb.zip.read(wb.calc_chain))[0]
    key = f"{sid_of[sid]}!{cell}"
    res = fill(wb, {key: 1}, allow_formula_overwrite=True, force=True)
    pr = prove(path.read_bytes(), res.blob, res.allowed_inputs, set(), {(sid_of[sid], cell)})
    assert pr.ok, (pr.unexpected[:5], pr.lint_errors[:5])
    after = Workbook(res.blob)
    entries = calc_chain_entries(after.zip.read(after.calc_chain))
    assert (sid, cell) not in entries and len(entries) == len(calc_chain_entries(wb.zip.read(wb.calc_chain))) - 1
