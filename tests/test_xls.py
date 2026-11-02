"""古いExcel（.xls）の読み書きのテスト。

公開されている .xls の様式（corpus/wide_xls、scripts/collect_forms.py --ext xls で取得）で確かめる。ないときは飛ばす。
正解はいつもExcelが保存した値（数式の保存値）と、書き込み前のファイルそのもの。
"""

from __future__ import annotations

import glob
from pathlib import Path

import pytest

from todoku.formula.parser import parse_cell
from todoku.xls.book import XlsWorkbook
from todoku.xls.write import XlsEditor
from todoku.xlsx import window
from todoku.xlsx.audit import run as audit
from todoku.xlsx.calc import verify
from todoku.xlsx.edit import fill
from todoku.xlsx.proof import prove
from todoku.xlsx.window import fill_blanks, inline_blanks

XLS = sorted(glob.glob(str(Path(__file__).parent.parent / "corpus" / "wide_xls" / "**" / "*.xls"), recursive=True))


def need():
    if not XLS:
        pytest.skip("古いExcelの様式がありません（scripts/collect_forms.py --ext xls --out corpus/wide_xls）")


@pytest.mark.parametrize("path", XLS, ids=lambda p: Path(p).name)
def test_xls_formulas_match_excel_and_rebuild_is_identical(path):
    blob = Path(path).read_bytes()
    wb = XlsWorkbook(blob, Path(path).name)
    s = verify(wb).summary()
    assert s["mismatch"] == 0 and s["unsupported"] == 0, s  # 数式はExcelの保存値と全数で一致
    assert XlsEditor(wb).build() == wb.data  # 書き込まずに組み立て直すと、ストリームはバイト一致


def test_xls_write_and_prove():
    need()
    path = next((p for p in XLS if Path(p).name == "saigai-yousiki1.xls"), XLS[0])
    blob = Path(path).read_bytes()
    wb = XlsWorkbook(blob, "x.xls")
    doc = window.build(wb)
    fields = [(s["name"], f) for s in doc["sheets"] for f in s["fields"] if f["kind"] in ("text", "number")][:40]
    vals = {f["id"]: (7 if f["kind"] == "number" else "とどく試験") for _, f in fields}
    res = fill(wb, vals)
    after = XlsWorkbook(res.blob)
    for k, v in vals.items():
        sh, a = k.rsplit("!", 1)
        assert after.sheet(sh).cells[parse_cell(a)].value == v
    assert verify(after).summary()["mismatch"] == 0  # 書いた後の計算結果も、計算し直した値と一致
    allowed = {(p.sheet, p.address) for p in res.patches if p.kind == "input"}
    pr = prove(blob, res.blob, allowed)
    assert pr.ok, (pr.unexpected[:5], pr.lint_errors[:5])
    assert pr.entries_identical == pr.entries_total - 1  # ブック以外のストリームはバイト一致


@pytest.mark.parametrize("name", ["saigai-yousiki1.xls", "731981.xls", "taiseijoukyouitiranhyou.xls"])
def test_xls_audit(name):
    need()
    path = next((p for p in XLS if Path(p).name == name), None)
    if path is None:
        pytest.skip(name)
    au = audit(Path(path).read_bytes(), name)
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]


@pytest.mark.parametrize("text, n", [
    ("令和　　年　　月　　日", 3), ("平成　年　月　日", 3), ("１．廃　業　年　月　日　　　　年　　月　　日", 3),
    ("第　　　　号", 1), ("北海道知事　　　　様", 0), ("電話　　　（　　　）　　　　", 3), ("Name: ________", 1),
    ("Total  (USD)", 0), ("事 業 所 番 号", 0), ("（事実発生日から 年間）", 1),
])
def test_inline_blanks(text, n):
    assert len(inline_blanks(text)) == n


def test_fill_blanks_keeps_the_rest_of_the_text():
    assert fill_blanks("令和　　年　　月　　日", ["8", "4", "1"]) == "令和8年4月1日"
    assert fill_blanks("１．廃　業　年　月　日　　　　年　　月　　日", ["2026", "11", "2"]) == "１．廃　業　年　月　日2026年11月2日"
