"""書類の段階（S〜F）：誰が扱えば正確に読んで入れられるか。"""

from __future__ import annotations

import pytest

from todoku.grade import combine, grade_blob, macro_is_display_only

from test_docx import SAMPLE_BODY, make_docx, p
from test_example import HEAD, VALS, make_xlsx
from test_pdf import sample_form


def test_data_files_are_s_and_broken_files_are_f():
    assert grade_blob(b"name,age\nA,3\n", "list.csv").grade == "S"
    assert grade_blob(b"\x00\x01garbage", "x.bin").grade == "F"


def test_excel_with_marked_inputs_only_is_a():
    every = [("A2:A9", "list"), ("B2:B9", "textLength"), ("C2:C9", "decimal"), ("D2:D9", "textLength")]
    g = grade_blob(make_xlsx({"Budget": (HEAD, every)}), "a.xlsx")
    assert g.grade == "A", g.headline()


def test_unmarked_table_columns_are_write_b_not_read():
    # 入力欄の印のない表の空のセル：規則では決められない（エージェントが判断して、番地で書く）。書く所の見つけ方なので、
    # 書く段階の B。中身は規則で正確に読めるので、読む段階は A のまま
    g = grade_blob(make_xlsx({"Budget": (HEAD, VALS)}), "b.xlsx")
    # （表の行の空のマスは、ほかの「規則では決められない所」と同じく［B5？］として窓に出し、まとめて数える）
    assert g.grade == "A" and not any(r.title == "規則では決められない所" for r in g.reasons), g.headline()
    assert g.write == "B" and any(r.title == "規則では決められない所" and r.grade == "B" for r in g.write_reasons), g.headline()


def test_layout_fields_in_pdf_and_word_are_write_b_and_read_a():
    # 配置から見つけた欄・書く所かもしれない所（表題「申請書」の右の余白・見出しの後ろ）は、書く段階の B（解析して書く）。
    # 読む段階は中身が読めるかで決めるので、字が正確に読める様式は A
    for blob, name in ((sample_form(), "s.pdf"), (make_docx(SAMPLE_BODY), "s.docx")):
        g = grade_blob(blob, name)
        titles = {r.title: r.grade for r in g.write_reasons}
        assert titles.get("配置から推定した入力欄") == "B" and titles.get("規則では決められない所") == "B", g.headline()
        assert g.grade == "A" and g.write == "B" and not g.reasons, g.headline()


def test_places_rules_cannot_decide_are_c_and_writable():
    # 規則の型に合わない日付の空き・伏せ字・「住所：」の後ろ：規則では欄と決めず、エージェントが判断する所（C）。id で書ける
    from todoku.docx.document import DocxDocument
    from todoku.docx.fill import fill
    from todoku.docx.proof import prove
    from todoku.docx.window import build

    blob = make_docx(SAMPLE_BODY + p("令和 年　月") + p("〒　○○○－…　○○県○○市○○区……") + p("住所："))
    g = grade_blob(blob, "c.docx")
    assert g.grade == "A" and g.write == "B" and any(r.title == "規則では決められない所" for r in g.write_reasons), g.headline()
    win = build(DocxDocument(blob, "c.docx"))
    agent = {f["id"]: f for f in win["fields"] if f.get("agent")}
    # P1.a1：表題「申請書」の後ろ。P2.1〜P2.3：「令和　　年　　月　　日」の2つ以上の空白の並び（字の意味は見ないので「？」）。
    # P3.a1：マスの見出し「氏名」の後ろ（すぐ右の空のマス T1.R1.C2 の代わりの所として、消さずに示す）
    # P9.a1：「令和 年　月」の後ろ。P9.1：「年　月」の全角の空白 1 つ（1 字ぶんの空き）。半角の空白 1 つ（「令和 年」）は語の区切りで、
    # 所にしない（字のまま窓に出る。書くときは字を目印にする）。P10：同じ記号の並び（○○○・○○県○○市○○）・点線「……」・
    # 全角の空白 1 つの空き（「〒　○○○」「…　○○県」）
    assert set(agent) == {"P1.a1", "P2.1", "P2.2", "P2.3", "P3.a1", "P9.1", "P9.a1", "P10.1", "P10.2", "P10.3", "P10.a1", "P10.a2",
                          "P11.a1"}, agent
    assert agent["P3.a1"]["alternative_to"] == "T1.R1.C2"
    assert agent["P10.a2"]["placeholder"] == "○○県○○市○○"
    res = fill(DocxDocument(blob, "c.docx"), {"P9.a1": "4", "P10.a2": "東京都千代田", "P11.a1": "東京都"})
    pr = prove(blob, res.blob, {x.id: x.value for x in res.patches})
    assert pr.ok, pr.unexpected
    after = build(DocxDocument(res.blob, "c.docx"))
    lines = [ln for st in after["stories"] for ln in st["reading"]]
    assert "令和 年［P9.1？］月［P9.a1＝4］" in lines and "住所：［P11.a1＝東京都］" in lines, lines
    assert any("［P10.a2＝東京都千代田］" in ln and "［P10.a1？＝○○○］" in ln for ln in lines), lines


def test_meaning_is_left_to_the_agent_not_to_word_matching():
    # 押印・書く値のぼかしは、言葉の一致では決めない（規則の段階は変えない）。エージェントが確かめる問いとして返す
    # 「代表者氏名　　　　　　　印」の氏名の空きは、語と語の間の空白（規則では決められない所：書く段階の B）。言葉では F・D にしない
    g = grade_blob(make_docx(SAMPLE_BODY + p("代表者氏名　　　　　　　印") + p("利用者数　約　　　　人")), "f.docx")
    assert g.grade == "A" and not g.reasons, g.headline()
    assert any(r.title == "規則では決められない所" for r in g.write_reasons)
    assert [c["id"] for c in g.agent_checks] == ["cannot_do", "ui_only", "vague_value"]
    assert g.write == "B"  # 書く所を、配置から推定するかエージェントが決める


def test_agent_findings_are_combined_with_provenance():
    rule = grade_blob(make_docx(SAMPLE_BODY), "s.docx")
    # 押印は、書く段階の D（エージェントにはできない記入）。読む段階は変えない
    g = combine(rule, [{"check": "cannot_do", "title": "押印", "detail": "代表者の印", "where": "P9"}],
                answered=["ui_only", "vague_value"])
    assert g.grade == rule.grade == "A" and g.write == "D" and g.without_barriers == "B" and not g.agent_checks
    assert [r.by for r in g.write_reasons if r.grade == "D"] == ["agent"] and g.barriers
    assert "押印：エージェントの判断" in g.headline() and "それがなければ書く段階 B" in g.headline()
    d = combine(rule, [{"grade": "D", "title": "書く値をぼかす書き方", "where": ["P12.1"]}])
    assert d.grade == "D" and d.write == "B" and [c["id"] for c in d.agent_checks] == ["cannot_do", "ui_only"]  # まだ確かめていない問いは残る
    u = combine(rule, [{"score": "write", "grade": "C", "title": "電子申請の画面から入力"}])
    assert u.write == "C" and "ui_only" not in [c["id"] for c in u.agent_checks]
    with pytest.raises(ValueError):
        combine(rule, [{"grade": "Z"}])
    with pytest.raises(ValueError):
        combine(rule, [{"score": "write", "grade": "F"}])  # 書く段階は A〜D
    with pytest.raises(ValueError):
        combine(rule, [{"check": "nothing"}])


def test_write_grade_by_rules():
    # 開けない書類：読む F・書く D
    g = grade_blob(b"not a document", "x.docx")
    assert g.grade == "F" and g.write == "D" and "書く段階 D" in g.headline()
    d = grade_blob(b"a,b\n1,2\n", "x.csv")
    assert d.grade == "S" and d.write == "A" and d.to_dict()["write"]["grade"] == "A"


def test_display_only_macros_do_not_lower_the_grade():
    printing = ('Sub 印刷()\n    Sheets(Array("申請書", "明細")).Select\n    ActiveWindow.SelectedSheets.PrintOut Copies:=1, _\n'
                '        Collate:=True\n    Sheets("申請書").Select\nEnd Sub\n')
    assert macro_is_display_only(printing)
    assert not macro_is_display_only('Sub 消す()\n    Range("B3:F20").ClearContents\nEnd Sub\n')
    assert not macro_is_display_only('Sub 写す()\n    Range("A1").Value = Range("B1").Value\nEnd Sub\n')


def test_value_accepts_space_after_sign_like_excel():
    from todoku.formula.values import parse_number

    # 「▲ 0.89」を SUBSTITUTE で「- 0.89」にして VALUE に渡す書類（総務省の様式）。Excel は -0.89 と読む
    assert parse_number("- 0.89") == -0.89 and parse_number("1,234") == 1234.0
    assert parse_number("1_000") is None and parse_number("inf") is None and parse_number("nan") is None


def test_underscore_inside_a_url_is_not_a_blank():
    from todoku.xlsx.window import inline_blanks

    assert inline_blanks("https://www.moj.go.jp/MINJI/minji05_00589.html") == []
    assert inline_blanks("氏名___") == [(2, 5)]  # 下線の空欄は、これまでどおり
