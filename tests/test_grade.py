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


def test_unmarked_table_columns_are_b():
    g = grade_blob(make_xlsx({"Budget": (HEAD, VALS)}), "b.xlsx")
    assert g.grade == "B" and any(r.title == "入力欄か決められないセル" for r in g.reasons), g.headline()


def test_layout_fields_in_pdf_and_word_are_b():
    assert grade_blob(sample_form(), "s.pdf").grade == "B"
    assert grade_blob(make_docx(SAMPLE_BODY), "s.docx").grade == "B"


def test_meaning_is_left_to_the_agent_not_to_word_matching():
    # 押印・書く値のぼかしは、言葉の一致では決めない（規則の段階は B のまま）。エージェントが確かめる問いとして返す
    g = grade_blob(make_docx(SAMPLE_BODY + p("代表者氏名　　　　　　　印") + p("利用者数　約　　　　人")), "f.docx")
    assert g.grade == "B", g.headline()
    assert [c["id"] for c in g.agent_checks] == ["cannot_do", "vague_value"]


def test_agent_findings_are_combined_with_provenance():
    rule = grade_blob(make_docx(SAMPLE_BODY), "s.docx")
    g = combine(rule, [{"grade": "F", "title": "押印", "detail": "代表者の印", "where": "P9"}], answered=["vague_value"])
    assert g.grade == "F" and g.without_barriers == "B" and not g.agent_checks
    assert [r.by for r in g.reasons if r.grade == "F"] == ["agent"]
    assert "押印：エージェントの判断" in g.headline() and "それがなければ段階 B" in g.headline()
    d = combine(rule, [{"grade": "D", "title": "書く値をぼかす書き方", "where": ["P12.1"]}])
    assert d.grade == "D" and [c["id"] for c in d.agent_checks] == ["cannot_do"]  # まだ確かめていない問いは残る
    with pytest.raises(ValueError):
        combine(rule, [{"grade": "Z"}])


def test_display_only_macros_do_not_lower_the_grade():
    printing = ('Sub 印刷()\n    Sheets(Array("申請書", "明細")).Select\n    ActiveWindow.SelectedSheets.PrintOut Copies:=1, _\n'
                '        Collate:=True\n    Sheets("申請書").Select\nEnd Sub\n')
    assert macro_is_display_only(printing)
    assert not macro_is_display_only('Sub 消す()\n    Range("B3:F20").ClearContents\nEnd Sub\n')
    assert not macro_is_display_only('Sub 写す()\n    Range("A1").Value = Range("B1").Value\nEnd Sub\n')
