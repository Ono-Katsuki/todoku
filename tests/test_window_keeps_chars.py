"""窓は書類の字を一字も変えない：点線・下線の字を消さない。id と空欄の数え方は変わらない。"""

from __future__ import annotations

import io
import zipfile

import pytest

from todoku import api
from todoku.docx.document import DocxDocument
from todoku.docx.window import build, to_markdown
from todoku.xlsx import window as xw

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _docx(body: str) -> bytes:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0"?><w:document {NS}><w:body>{body}</w:body></w:document>')
    return b.getvalue()


def _p(t: str) -> str:
    return f'<w:p><w:r><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'


def test_word_leaders_stay_in_the_mark_and_ids_do_not_change():
    win = build(DocxDocument(_docx(_p("Name .......... end") + _p("Date ………… end") + _p("氏名＿＿＿＿") + _p("Tel ---- x")), "a.docx"))
    md = to_markdown(win)
    assert "Name ［P1.c4？..........］ end" in md
    assert "Date ［P2.c4？…………］ end" in md
    assert "氏名＿＿＿＿［P3.c2］" in md
    assert "Tel ［P4.c3？----］ x" in md
    ids = {f["id"] for f in win["fields"]}
    assert {"P1.c4", "P2.c4", "P3.c2", "P4.c3"} <= ids
    ctx = {f["id"]: f.get("context", "") for f in win["fields"]}
    assert "［P1.c4？..........］" in ctx["P1.c4"]  # 印の中に字があっても、欄の前後の字（context）が取れる


def test_write_to_a_leader_place_still_works(tmp_path):
    src = tmp_path / "a.docx"
    src.write_bytes(_docx(_p("Name .......... end")))
    r = api.fill(str(src), {"P1.c4": "Al"}, str(tmp_path / "b.docx"), overwrite=True)
    assert r["proof_ok"] and r["read_back_ok"], r


@pytest.mark.parametrize("cover,want", [("……", "［？……］"), ("......", "［？......］"), ("---", "［？---］"), ("　　", "［？　　］")])
def test_cell_token_keeps_chars_of_an_agent_place(cover, want):
    text = f"A{cover}B"
    assert xw.blank_token(text, 1, 1 + len(cover), "leader") == want
    assert xw.blank_token(text, 1, 1 + len(cover), "leader", mark=True) == want.replace("［", "［▶")


def test_cell_token_of_a_rule_blank():
    assert xw.blank_token("A＿＿B", 1, 3, None) == "＿＿［　］"
    assert xw.blank_token("（　）", 1, 2, None) == "［　］"
    assert xw.blank_token("A＿＿B", 1, 3, None, mark=True) == "＿＿［▶］"


def test_text_file_keeps_leaders(tmp_path):
    src = tmp_path / "a.txt"
    src.write_text("Name .......... end\n氏名＿＿＿＿\n", encoding="utf-8")
    k, _d, parts = api.window(str(src))
    text = "".join(x if isinstance(x, str) else str(x) for x in parts)
    assert ".........." in text and "＿＿＿＿" in text


def test_word_box_glyph_stays_the_printed_char():
    win = build(DocxDocument(_docx(_p("□ はい ❑ いいえ")), "a.docx"))
    md = to_markdown(win)
    assert "□［" in md and "❑［" in md


def test_word_control_shows_its_prompt_text_as_a_sample():
    sdt = ('<w:p><w:sdt><w:sdtPr><w:alias w:val="d"/><w:showingPlcHdr/><w:date/></w:sdtPr><w:sdtContent><w:r><w:t>From</w:t></w:r></w:sdtContent></w:sdt>'
           '<w:r><w:t xml:space="preserve"> to x</w:t></w:r></w:p>')
    md = to_markdown(build(DocxDocument(_docx(sdt), "a.docx")))
    assert "（見本の字「From」）" in md
    assert "＝From" not in md  # 値としては見せない


def test_pdf_printed_box_glyph_and_button_caption():
    import pdfmini
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.window import _box_glyph, build as pbuild

    class S:
        checked = False

    assert _box_glyph(S, "❑") == "❑" and _box_glyph(S, "") == "☐"
    S.checked = True
    assert _box_glyph(S, "☑") == "☑" and _box_glyph(S, "") == "☑" and _box_glyph(S, "□") == "☑"
    w = "<< /Type /Annot /Subtype /Widget /FT /Btn /Ff 65536 /T (R1) /Rect [100 700 160 720] /F 4 /MK << /CA (Reset) >> >>"
    blob = pdfmini.build("", "<< >>", [w], annots="/Annots [2 0 R]", acroform="/AcroForm << /Fields [2 0 R] >>")
    f = pbuild(PdfDocument(blob))["pages"][0]["fields"][0]
    assert f["label"] == "Reset"
