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
    # 字の形と並びだけで決める（字の意味（年・様・号など）は見ない）：2つ以上の空白の並び・下線の字・かっこの中の空白・点線
    ("令和　　年　　月　　日", 3), ("１．廃　業　年　月　日　　　　年　　月　　日", 7),
    ("第　　　　号", 1), ("電話　　　（　　　）　　　　", 3), ("Name: ________", 1),
    # 1 字ずつ同じ空白で 3 字以上（字間を空けた見出しの形）：「年　　月　　日」と形が同じなので落とさず、spaced の「？」
    ("事 業 所 番 号", 0), ("職　　業　　名", 2), ("事　　業　　所", 2), ("年　　月　　日", 2),
    ("第　　号", 1), ("金　　　　円", 1), ("月　　回", 1), ("平成　　年　　月　　日", 3),
    # 全角の空白（U+3000）は 1 つでも和字の空白 1 字ぶん（1 em：Unicode・JIS X 4051）の空き：「？」の所（「山田　太郎」の
    # ような語の区切りとは形で区別できないので、エージェントが前後の字から判断する）。半角の空白 1 つは語の区切りで、所にしない
    ("平成　年　月　日", 3), ("郡　市", 1), ("第　号", 1), ("山田　太郎", 1),
    ("　年　月　日", 2),  # 段落の頭の全角の空白 1 つは字下げ（字と字の間ではない）
    ("（事実発生日から 年間）", 0), ("Name here", 0),
    # 字間を空けた見出し・チェックの箱の字の後ろの、字間を空けた選択肢の字も spaced の「？」
    ("事　業　所", 2), ("□地　域　移　行　支　援", 5), ("☐居　宅　介　護　☐重　度　訪　問", 6),
    # 点線・ダッシュだけの長い行（区切りの飾りのことが多い）も落とさず decor の「？」
    ("..............................", 1), ("……………………", 1), ("－－－－－－－－", 1),
    # 2 字の間の空白（「氏　　名」と「第　　号」は形が同じ）・語の後ろの空き（「知事　　様」「Total  (USD)」）は、
    # 規則では決めない所（「？」）として示す（エージェントが前後の字から判断する）
    ("氏　　名", 1), ("番　　号", 1), ("北海道知事　　　　様", 1), ("Total  (USD)", 1),
    # 日本語以外の書類
    ("中華民國　　年　　月　　日", 3), ("民國　　年　　月　　日", 3), ("　　년　　월　　일", 3), ("　　시　　구　　동", 3),
    ("　　省　　市　　县", 3), ("Name ..............................", 1), ("Nom : …………………………", 1), ("नाम ...................", 1),
    ("[ ] Yes [ ] No", 2), ("( ) Male ( ) Female", 2), ("Signature ______", 1), ("الاسم: ____", 1), ("ชื่อ ____", 1),
    ("Date: ___/___/______", 3), ("https://www.moj.go.jp/MINJI/minji05_00589.html", 0), ("I am a b c", 0),
])
def test_inline_blanks(text, n):
    assert len(inline_blanks(text)) == n


def test_spaced_and_decor_kinds():
    from todoku.xlsx.window import AGENT_WHY, blank_places

    assert {k for _a, _b, k in blank_places("事　　業　　所")} == {"spaced"}
    assert {k for _a, _b, k in blank_places("□地　域　移　行")} == {"spaced"}
    assert {k for _a, _b, k in blank_places("　年　　月　　日")} == {"gap"}
    # 点線・ダッシュだけの行：鍵は長さでは分けない（以前の decor：8 字以上）。段落の字のすべてかは facts の whole_line
    got = blank_places("－－－－－－－－")
    assert [k for _a, _b, k in got] == ["dash"] and got[0].facts["whole_line"] and got[0].facts["len"] == 8
    assert "spaced" in AGENT_WHY


def test_fill_blanks_keeps_the_rest_of_the_text():
    assert fill_blanks("令和　　年　　月　　日", ["8", "4", "1"]) == "令和8年4月1日"
    # 番号のついた字間を空けた見出し（「１．廃　業　年　月　日」）の全角の空白 1 つも 1 字ぶんの空き（？）：書かない所は None
    # 空白だけの空きに書くときは、前後の印刷された字の側に元の空白を 1 字ずつ残す（後ろの字の字数では決めない）
    assert fill_blanks("１．廃　業　年　月　日　　　　年　　月　　日", [None] * 4 + ["2026", "11", "2"]) \
        == "１．廃　業　年　月　日　2026　年11月2日"
