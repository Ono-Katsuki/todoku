"""エージェントが書類を「全て正確に読み書きできるか」のテスト。

読む: ファイルの文字と数値が1つ残らず窓に出る。見出しは同じ言葉（氏名など）を取り違えない形で付く。
      条件付き書式で隠れている警告文は、見えているものとして伝えない。
書く: 値はExcelで手入力したときと同じ規則で受け取り、Excelが受け付けない値は受け付けない。
      チェックボックスは、Excelが保存したのと同じバイト列で切り替わる。
公開様式（corpus/forms）は scripts/fetch_forms.py で取得する。ないときは、そのテストを飛ばす。
"""

from __future__ import annotations

import datetime as dt
import io
from pathlib import Path

import pytest

from todoku.formula.numfmt import is_date_format
from todoku.formula.parser import parse_cell
from todoku.xlsx import window
from todoku.xlsx.audit import run as audit
from todoku.xlsx.edit import EditError, fill, interpret_input, parse_date_text
from todoku.xlsx.proof import prove
from todoku.xlsx.workbook import Workbook, load

FX = Path(__file__).parent / "fixtures"
KAMI = FX / "kami_meisai.xlsx"
FORMS = Path(__file__).parent.parent / "corpus" / "forms"


def form(name: str) -> Path:
    p = FORMS / f"{name}.xlsx"
    if not p.exists():
        pytest.skip("公開様式がありません（scripts/fetch_forms.py で取得）")
    return p


def fields_of(doc: dict, sheet_prefix: str) -> dict[str, dict]:
    sh = next(s for s in doc["sheets"] if s["name"].startswith(sheet_prefix))
    return {(f.get("cell") or f["control"]): f for f in sh["fields"]}


# ------------------------------------------------------------------ Excelと同じ受け付け
def test_invalid_choice_is_refused_like_excel():
    # 入力規則のエラーが「停止」のときは、Excelでも一覧にない値は入らない
    with pytest.raises(EditError, match="選択肢にありません"):
        fill(load(str(KAMI)), {"明細書!I4": "10級地"})
    res = fill(load(str(KAMI)), {"明細書!I4": "10級地"}, force=True)
    assert any("選択肢にありません" in w for w in res.warnings)
    assert not fill(load(str(KAMI)), {"明細書!I4": "1級地"}).warnings


@pytest.mark.parametrize("fmt, expected", [
    ("General", False), ('General"円"', False), ("0.00E+00", False), ("#,##0", False), ("@", False), ("0%", False),
    ("yyyy/m/d", True), ('[$-411]ggge"年"m"月"d"日"', True), ("[h]:mm", True), ("h:mm", True),
])
def test_is_date_format(fmt, expected):
    assert is_date_format(fmt) is expected


@pytest.mark.parametrize("text, ymd", [
    ("2026/4/1", (2026, 4, 1)), ("2026-04-01", (2026, 4, 1)), ("2026年4月1日", (2026, 4, 1)),
    ("令和8年4月1日", (2026, 4, 1)), ("R8.4.1", (2026, 4, 1)), ("令和元年5月1日", (2019, 5, 1)),
    ("平成31年4月30日", (2019, 4, 30)), ("２０２６／４／１", (2026, 4, 1)), ("2026/2/30", None), ("4月1日", None),
])
def test_parse_date_text(text, ymd):
    d = parse_date_text(text)
    assert (d.year, d.month, d.day) == ymd if ymd else d is None


@pytest.fixture(scope="module")
def typed_book() -> Workbook:
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "入力"
    ws["A1"].number_format = "yyyy/m/d"
    ws["B1"].number_format = "@"
    ws["C1"].number_format = "0%"
    ws["D1"] = 0  # 標準の数値のセル
    buf = io.BytesIO()
    wb.save(buf)
    return Workbook(buf.getvalue(), "typed.xlsx")


@pytest.mark.parametrize("cell, value, expected, note", [
    ("A1", "令和8年4月1日", 46113.0, None),  # 日付の書式のセル：日付の通し番号
    ("A1", "2026/4/1", 46113.0, None),
    ("B1", 123, "123", None),  # 文字列の書式のセル：入力したとおりの文字
    ("C1", "8", 0.08, None),  # パーセントの書式：Excelの「自動でパーセント入力」と同じ
    ("C1", "8%", 0.08, None),
    ("C1", 8, 8.0, "パーセント"),  # 数値で 8 を渡したら 800% になるので知らせる
    ("D1", "２０", 20.0, "全角"),  # 全角の数字は数値になる（Excelと同じ）
    ("D1", "1,234", "1,234", "書いたとおり"),  # 標準の表示形式の、計算に使わないセル：数だと「1234」と見えるので字のまま（第4回 C5）
    ("D1", "TRUE", True, None),
    ("D1", "0123", "0123", "先頭の 0"),  # 先頭の0を残す
    ("D1", "1234567890123456", "1234567890123456", "15桁"),
    ("D1", "2026/4/1", "2026/4/1", "日付"),  # 日付の書式でないセルに日付：書式は変えず文字のまま、知らせる
    ("D1", "", None, None),  # 空文字は消去
])
def test_interpret_input_like_excel(typed_book, cell, value, expected, note):
    r, c = parse_cell(cell)
    got, notes = interpret_input(typed_book, "入力", r, c, value)
    assert got == expected and type(got) is type(expected)
    assert (note is None and not notes) or (note is not None and any(note in n for n in notes)), notes


def test_new_cell_keeps_column_style():
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Font

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.column_dimensions["B"].font = Font(bold=True)  # 列の書式（セル要素はない）
    ws["A1"] = "見出し"
    buf = io.BytesIO()
    wb.save(buf)
    src = Workbook(buf.getvalue())
    sh = src.sheets[0]
    assert sh.style_of(5, 2) != 0
    res = fill(src, {"B5": "記入"})
    after = Workbook(res.blob).sheets[0]
    assert after.cells[(5, 2)].style == sh.style_of(5, 2)  # Excelと同じく列の書式を引き継ぐ
    assert prove(buf.getvalue(), res.blob, {(sh.name, "B5")}).ok


# ------------------------------------------------------------------ チェックボックス
@pytest.mark.corpus
def test_checkbox_write_matches_excel_bytes():
    blank, example = form("wakayama_r8keikaku"), form("wakayama_r8keikaku_rei")
    res = fill(Workbook(blank.read_bytes()), {f"別紙様式2-1 （総括表）!BA{r}": True for r in range(116, 123)})
    ours, ex = Workbook(res.blob), Workbook(example.read_bytes())
    parts = [p for p in res.changed_parts if "ctrlProps" in p or p.endswith(".vml")]
    assert len(parts) == 8
    for p in parts:
        assert ours.zip.read(p) == ex.zip.read(p), p  # Excelが保存した記入例とバイト一致
    allowed = {(p.sheet, p.address) for p in res.patches if p.kind == "input"}
    controls = {(p.sheet, p.address) for p in res.patches if p.kind == "control"}
    pr = prove(blank.read_bytes(), res.blob, allowed, controls)
    assert pr.ok and len(pr.control_changes) == 7, pr.unexpected


@pytest.mark.corpus
def test_checkbox_by_name_and_refuses_non_bool():
    blank = form("wakayama_r8keikaku")
    res = fill(Workbook(blank.read_bytes()), {"別紙様式2-1 （総括表）!Check Box 3": "はい"})
    sh = Workbook(res.blob).sheet("別紙様式2-1 （総括表）")
    assert sh.cells[parse_cell("BA118")].value is True  # つながったセルも一緒に変わる
    assert next(c for c in sh.controls if c.name == "Check Box 3").is_checked
    with pytest.raises(EditError, match="チェックボックス"):
        fill(Workbook(blank.read_bytes()), {"別紙様式2-1 （総括表）!BA118": "済"})


# ------------------------------------------------------------------ 見出し
@pytest.mark.corpus
def test_labels_distinguish_same_words():
    doc = window.build(Workbook(form("wakayama_r8keikaku").read_bytes()))
    basic = fields_of(doc, "基本情報")
    assert basic["L21"]["label"] == "法人 代表者／氏名"
    assert basic["L24"]["label"] == "書類作成 担当者／氏名"  # 同じ「氏名」でも、だれの氏名かが分かる
    assert basic["L15"]["label"] == "法人名／フリガナ"
    # 「〒｜［L17］｜-｜［Q17］」：欄の並びの後ろが字で終わらない（字にはさまれていない）ので並びにはせず、
    # いちばん近い字「-」をそのまま見出しにする（「〒」「-」を語として読まない）。並び全体は context に出る
    assert basic["Q17"]["label"] == "法人住所／-" and "〒 ｜ ［L17:O17］ ｜ - ｜ ［▶Q17:U17］" in basic["Q17"]["context"]
    assert basic["Q35"]["label"] == "事業所の所在地／都道府県"
    s21 = fields_of(doc, "別紙様式2-1")
    assert s21["E131"]["label"] == "令和［▶］年［　］月［　］日"
    assert s21["H131"]["label"] == "令和［　］年［▶］月［　］日"
    # ○を付ける選択肢（入力規則の選択肢が短い）でも、右の字を見出しにする決め方はやめた（選択肢の数・字数による判断）。
    # 右の項目の字は、近くの字（near.right）としてそのまま出る
    assert s21["E70"]["label"].startswith("入職促進に向けた取組")
    assert s21["E70"]["near"]["right"]["text"].startswith("③他産業からの転職者") and s21["E70"]["near"]["right"]["cols"] == 0
    assert s21["BA116"]["kind"] == "checkbox" and s21["BA116"]["label"].startswith("・処遇改善加算として給付される額は")
    assert s21["Q18"]["label"].startswith("②令和８年度の賃金改善の見込額 （①の額以上")  # 番号も文字も変えない
    s22 = fields_of(doc, "別紙様式2-2")
    # 見出しは、すぐ左の字「ヶ月）」（前の欄の並びの続き）もそのまま入れる（「ヶ月」を単位の語として除かない）
    assert s22["AF13"]["label"] == "ヶ月）／①月額賃金要件／月額賃金要件を満たす"
    # 「令和［R13］年［T13］月～令和［V13］年［X13］月｜（｜…｜ヶ月）」：字にはさまれた欄の並びをそのまま見出しにする。上の見出し
    # 「算定対象月 (c)」は、並びの後ろの字のマス（「（」「ヶ月）」）まで覆うので使わない（字の意味で除かない）。並びは context に出る
    assert s22["R13"]["label"] == "令和［▶］年［　］月～令和［　］年［　］月"


@pytest.mark.corpus
def test_labels_without_merged_sections():
    # 区分の見出しが最初の行にだけ書かれ、結合されていない様式（新潟県）
    doc = window.build(Workbook(form("niigata_490353").read_bytes()))
    basic = fields_of(doc, "基本情報")
    assert basic["M23"]["label"] == "法人名／名称"
    assert basic["M28"]["label"] == "法人代表者／氏名"
    # 郵便番号を1桁ずつ書くマス：見出しは、いちばん近い字をそのまま（「1マスに1文字」とは推定しない。並びは context に出る）
    assert basic["N24"]["label"] == "〒" and basic["Q24"]["label"] == "－"
    assert "〒 ｜ ［M24］ ｜ ［▶N24］ ｜ ［O24］ ｜ － ｜ ［Q24］" in basic["N24"]["context"]


@pytest.mark.corpus
def test_dependent_choices_follow_other_fields():
    # 市区町村の選択肢は、同じ行の都道府県で決まる。V37 の規則 INDIRECT($Q36) は範囲 V36:V40 の左上を基準に
    # 書かれているので、V37 では Q37 を見る（Excelと同じく相対参照をずらす）
    blank = Workbook(form("wakayama_r8keikaku").read_bytes())
    fill(blank, {"基本情報入力シート!Q37": "東京都", "基本情報入力シート!V37": "千代田区"})
    with pytest.raises(EditError, match="選択肢にありません"):
        fill(blank, {"基本情報入力シート!Q37": "東京都", "基本情報入力シート!V37": "札幌市"})


# ------------------------------------------------------------------ 見え方
@pytest.mark.corpus
def test_conditional_format_hides_warning_when_satisfied():
    name = "別紙様式3-1（処遇改善加算　総括表）"
    ex = window.build(Workbook(form("niigata_490358").read_bytes()))
    blank = window.build(Workbook(form("niigata_490353").read_bytes()))
    hid_ex = next(s for s in ex["sheets"] if s["name"] == name)["hidden_by_format"]
    hid_bl = next(s for s in blank["sheets"] if s["name"] == name)["hidden_by_format"]
    assert "AQ18" in hid_ex  # 記入例では要件を満たすので、警告文は灰色の地に灰色で見えない
    assert "AQ18" not in hid_bl
    line18 = next(x["text"] for x in next(s for s in ex["sheets"] if s["name"] == name)["reading"] if x.get("row") == 18)
    # いまは見えない警告文は、見える字としては読まない。ファイルにある字なので、印をつけて番地つきで出す
    shown, _, faded = line18.partition("〔条件付き書式でいまは見えない〕")
    assert "下回っています" not in shown and "〔AQ18〕" in faded and "下回っています" in faded


@pytest.mark.corpus
def test_values_hidden_by_merge_are_listed():
    doc = window.build(Workbook(form("wakayama_r8keikaku").read_bytes()))
    sh = next(s for s in doc["sheets"] if s["name"].startswith("別紙様式2-1"))
    e100 = next(c for c in sh["cells"] if c["cell"] == "E100")
    assert e100["hidden_by_merge"] == "B100:E101" and e100["value"] is False


@pytest.mark.corpus
def test_window_fits_in_context():
    doc = window.build(Workbook(form("wakayama_r8keikaku").read_bytes()))
    # 3,000か所を超える入力欄があっても、見えるシートは AI が一度に読める大きさ。非表示のシート（市町村の一覧など）の中身も
    # 窓に出す（印をつけて）ので、全体は、分けた部分のどれもが目安の大きさに収まること
    visible = sum(len("\n".join(window._md_sheet(sh))) for sh in doc["sheets"] if sh["state"] == "visible")
    assert visible < 60_000
    assert all(len(t) <= window.BUDGET for _, t in window.markdown_parts(doc))


# ------------------------------------------------------------------ 監査（読み落とし・全欄への書き込み・読み戻し）
def test_audit_fixture():
    au = audit(KAMI.read_bytes(), KAMI.name)
    assert au.ok, [(c.name, c.items[:3]) for c in au.checks if not c.ok]


@pytest.mark.corpus
@pytest.mark.parametrize("blank, example", [
    ("wakayama_r8keikaku", "wakayama_r8keikaku_rei"),
    ("niigata_490353", "niigata_490358"),
    ("wakayama_r8henkou", None),
    ("wakayama_r8tokurei", None),
])
def test_audit_public_forms(blank, example):
    ex = form(example).read_bytes() if example else None
    au = audit(form(blank).read_bytes(), blank, ex)
    assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]
