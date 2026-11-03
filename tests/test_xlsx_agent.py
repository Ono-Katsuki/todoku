"""Excel の窓：規則では決められない所（［B5？］・文字の中の［？］）を、黙って落とさず、エージェントに根拠つきで示す。

ファイルの作り（罫線・塗り・結合・入力規則・保護）で決まることは規則で、語の並びや近くの見出しから推定することは
エージェントの判断に任せる（窓では「？」）。"""

from __future__ import annotations

import io
from types import SimpleNamespace

import pytest

from todoku.grade import grade_blob, image_only_sheets
from todoku.xlsx import window
from todoku.xlsx.edit import fill
from todoku.xlsx.workbook import Workbook

openpyxl = pytest.importorskip("openpyxl")
from openpyxl.styles import Border, PatternFill, Side  # noqa: E402

THIN = Side(style="thin")


def blob_of(build) -> bytes:
    wb = openpyxl.Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def doc_of(blob: bytes) -> dict:
    return window.build(Workbook(blob, "t.xlsx"))


def fields_of(doc: dict) -> dict:
    return {f["id"]: f for sh in doc["sheets"] for f in sh["fields"]}


def test_bordered_blank_without_label_is_shown_as_agent_and_writable():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "申請書"
        ws["B5"].border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)  # 見出しのない枠
        ws["D7"].fill = PatternFill("solid", fgColor="FFFF00")  # 塗りだけの空のマス
        ws["F9"].border = Border(left=THIN)  # 1辺だけ（隣の表の外枠）は印にしない
        ws["H9"] = "x"

    blob = blob_of(build)
    doc = doc_of(blob)
    f = fields_of(doc)
    assert f["S!B5"]["agent"] and f["S!B5"]["provenance"] == "agent"
    assert f["S!B5"]["evidence"]["borders"] == ["左", "右", "上", "下"] and f["S!B5"]["why_key"] == ["blank_box"]
    assert f["S!D7"]["agent"] and f["S!D7"]["evidence"]["fill"] == "FFFF00"
    assert "S!F9" not in f
    reading = " ".join(x["text"] for x in doc["sheets"][0]["reading"])
    assert "［B5？］" in reading and "［D7？］" in reading
    assert {u["cell"] for u in doc["sheets"][0]["undecided"]} == {"B5", "D7"}
    md = window.to_markdown(doc)
    assert "規則では決められない所（2 か所" in md and "［B5？］" in md
    # エージェントが書く所と判断したら、ほかの入力欄と同じく番地で書ける
    res = fill(Workbook(blob, "t.xlsx"), {"S!B5": "とどく"})
    assert Workbook(res.blob, "t.xlsx").sheet("S").cells[(5, 2)].value == "とどく"


def test_labelled_box_stays_input_but_grades_b():
    # 枠と近くの見出しから決めた欄は入力欄のまま。ただし印（入力規則・保護）ではないので、書く段階は B（A ではない）
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "氏名"
        ws["B1"].border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    blob = blob_of(build)
    f = fields_of(doc_of(blob))
    assert not f["S!B1"].get("agent") and f["S!B1"]["reason"] == ["枠のある空欄"]
    g = grade_blob(blob, "t.xlsx")
    titles = {r.title: r.grade for r in g.write_reasons}
    assert titles.get("枠・塗りから推定した入力欄") == "B" and "印のある入力欄" not in titles, g.headline()
    assert g.grade == "A" and g.write == "B"  # 入力欄の見つけ方は、読む段階を下げない


def test_whole_column_validation_and_wide_formula_refs_are_clipped_not_skipped():
    from openpyxl.worksheet.datavalidation import DataValidation

    def build(wb):
        ws = wb.active
        ws.title = "S"
        for r in range(1, 9):
            ws.cell(r, 3, f"項目{r}")
        dv = DataValidation(type="list", formula1='"A,B"')
        dv.add("A1:A1048576")  # 列全体（5000 セルを超える範囲は、以前は飛ばしていた）
        ws.add_data_validation(dv)
        ws["E1"] = "=SUM(D1:D1000)"  # 400 セルを超える参照（以前は飛ばしていた）

    f = fields_of(doc_of(blob_of(build)))
    assert "入力規則があるセル" in f["S!A8"]["reason"] and "S!A9" not in f  # 使っている範囲（8 行目）までに切る
    assert "数式が参照している値" in f["S!D8"]["reason"] and "S!D9" not in f


def test_text_cell_gaps_are_agent_and_structural_blanks_are_rules():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "議長　　　　氏"  # 2つ以上の空白の並び：字の意味からは決めない（「？」）
        ws["A2"] = "令和　　年　　月　　日"  # 同じ（年・月・日の語は見ない）
        ws["A3"] = "電話（　　　　）"  # かっこの中の空白だけ：字の形で決まる
        ws["A4"] = "氏名＿＿＿＿"  # 下線：字の形で決まる

    blob = blob_of(build)
    doc = doc_of(blob)
    f = fields_of(doc)
    assert f["S!A1"]["agent"] and f["S!A1"]["spans"][0]["rule"] is False and "1 字ぶん以上の空白" in f["S!A1"]["why"]
    assert f["S!A2"]["agent"] and f["S!A2"]["blanks"] == 3 and not any(s["rule"] for s in f["S!A2"]["spans"])
    assert not f["S!A3"].get("agent") and f["S!A3"]["provenance"] == "layout"
    assert not f["S!A4"].get("agent") and f["S!A4"]["spans"][0]["rule"]
    reading = {x.get("row"): x["text"] for x in doc["sheets"][0]["reading"]}
    # 「？」の所の印は、元の空白をそのまま中に入れる（空きの大きさを印から読める）
    assert reading[1] == "〔A1〕議長　［？　　　］氏" and reading[2] == "〔A2〕令和［？　　］年［？　　］月［？　　］日"
    assert reading[3] == "〔A3〕電話（［　］）"
    # 書くのは、空欄ごとの値のリスト（書かない所は null）
    res = fill(Workbook(blob, "t.xlsx"), {"S!A1": ["山田"], "S!A2": ["8", "4", None]})
    sh = Workbook(res.blob, "t.xlsx").sheet("S")
    assert sh.cells[(1, 1)].value == "議長　山田氏" and sh.cells[(2, 1)].value == "令和8年4月　　日"


def test_after_label_is_not_cut_by_length_or_other_blanks():
    # 見出しの後ろの候補は、字数（2〜16 字）やほかの空きのあるなしでは落とさない。detail=True で、どれに当たるかを鍵で示す
    long = "令和八年度地域福祉活動推進事業補助金交付申請"
    assert window.agent_blanks(long) == [(len(long), len(long), "after_label")]
    assert window.agent_blanks(long, detail=True)[0][2] == "after_label_long"
    assert window.agent_blanks("氏名＿＿＿＿電話番号", found=[(2, 6)], detail=True)[-1][2] == "after_label_more"
    # 中に 1 字ぶん以上の空き（全角の空白 1 つ：wide_space）のある段落は、空きを別に示すので、段落の後ろは出さない
    # （「１　目的」の空きと後ろの両方を出さない。PDF・Word と同じ 1 em の決まり）
    assert window.agent_blanks("氏名＿＿＿＿　電話番号", found=[(2, 6)], detail=True) == []
    assert window.agent_blanks("１　目的") == [] and window.agent_blanks("1  Purpose") == []
    assert window.agent_blanks("1 Purpose") == [(9, 9, "after_label")]  # 半角の空白 1 つは語の区切り
    assert window.agent_blanks("住所", detail=True) == [(2, 2, "after_label")]
    assert window.agent_blanks("上記のとおり申請します。") == []  # 文の区切りのある段落は、見出しではない


def test_tables_are_layout_estimates():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws.append(["品名", "数量", "単価"])
        ws.append(["りんご", 3, 100])

    t = doc_of(blob_of(build))["sheets"][0]["tables"][0]
    assert t["provenance"] == "layout" and t["decide"] == "agent"


def test_image_only_sheet_is_decided_by_structure():
    pic = SimpleNamespace(kind="picture", text="")
    cell = lambda v: SimpleNamespace(value=v, has_formula=False)  # noqa: E731

    def sheet(name, cells, shapes=(pic,)):
        return SimpleNamespace(name=name, shapes=list(shapes), controls=[], cells=cells, in_print_area=lambda r, c: True)

    many_empty = {(r, c): cell(None) for r in range(1, 30) for c in range(1, 10)}  # 書式だけの空のセルが多くても画像だけ
    wb = SimpleNamespace(sheets=[sheet("絵", many_empty), sheet("字", {(1, 1): cell("申請書")}), sheet("なし", {}, ())])
    assert image_only_sheets(wb) == ["絵"]


@pytest.mark.parametrize("ch", list("❏❐❑❒◻◽"))
def test_bullet_like_box_in_a_cell_is_agent_like_word_and_pdf(ch):
    # 箇条書きの印にもよく使う箱の字は、マスの中でも規則のチェックの欄にしない（Word・PDF と同じ「？」：bullet_box）
    assert window.cell_blanks(f"{ch}はい") == [(0, 1, "bullet_box")]
    assert window.cell_blanks("□はい") == [(0, 1, None)] and window.cell_blanks("☑はい") == [(0, 1, None)]
    assert "箇条書きの印" in window.AGENT_WHY["bullet_box"]
    assert window.blank_token(f"{ch}はい", 0, 1, "bullet_box") == f"{ch}［？］"


def test_gap_token_keeps_the_spaces_and_labels_use_the_one_em_rule():
    # 「？」の所の印は、元の空白をそのまま中に入れる（全角の空白 1 つの空きと、もっと広い空きを、印から見分けられる）
    assert window.blank_token("第　号", 1, 2, "gap") == "［？　］"
    assert window.blank_token("A  B", 1, 3, "gap") == "［？  ］"
    assert window.blank_token("A……B", 1, 3, "leader") == "［？］"  # 空白でない所は、字を入れない
    assert window.blank_token("（　）", 1, 2, None) == "［　］"
    # 見出しの切り方も、窓の空きと同じ 1 em の決まり（全角の空白 1 つ・半角の空白 2 つ）。字間を空けた見出しはまとめる
    assert window.label_tail("１　目的") == "目的" and window.label_tail("Name  Date") == "Date"
    assert window.label_tail("住所　　氏　名") == "氏　名" and window.label_tail("東京都 港区") == "東京都 港区"
    assert window.label_head("１　目的") == "１" and window.label_head("事　業　所") == "事　業　所"


def test_word_label_is_cut_at_one_wide_space():
    # Word のフォームの欄の見出しも、1 字ぶんの空き（全角の空白 1 つ）で切る（以前は 2 つ以上の空白だけで切っていた）
    import sys
    from pathlib import Path

    sys.path.insert(0, str(Path(__file__).parent))
    from test_docx import make_docx

    from todoku.docx.document import DocxDocument
    from todoku.docx.window import build

    sdt = ('<w:p><w:r><w:t xml:space="preserve">番号　氏名</w:t></w:r><w:sdt><w:sdtPr><w:alias w:val=""/><w:tag w:val="nm"/>'
           '<w:text/></w:sdtPr><w:sdtContent><w:r><w:t></w:t></w:r></w:sdtContent></w:sdt></w:p>')
    win = build(DocxDocument(make_docx(sdt), "a.docx"))
    f = [x for x in win["fields"] if x.get("source") != "layout"]
    assert f and f[0]["label"] == "氏名", f
