"""Word の取り消し線と字の色を、表計算と同じ印で窓に出す。既定の色には印を付けない。"""

from __future__ import annotations

import io
import struct
import zipfile

from todoku.doc.book import ICO, apply_look
from todoku.docx.document import DocxDocument
from todoku.docx.window import build, to_markdown

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
STYLES = ('<w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="auto"/></w:rPr></w:rPrDefault></w:docDefaults>'
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>'
          '<w:style w:type="paragraph" w:styleId="Red"><w:basedOn w:val="Normal"/><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>'
          '<w:style w:type="character" w:styleId="Gone"><w:rPr><w:strike/></w:rPr></w:style>')


def _docx(body: str, styles: str = STYLES) -> bytes:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0"?><w:document {NS}><w:body>{body}</w:body></w:document>')
        z.writestr("word/styles.xml", f'<?xml version="1.0"?><w:styles {NS}>{styles}</w:styles>')
    return b.getvalue()


NO_STRIKE = '<w:strike w:val="0"/>'


def _c(v: str) -> str:
    return f'<w:color w:val="{v}"/>'


def _r(t: str, rpr: str = "") -> str:
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>'


def _md(body: str, styles: str = STYLES) -> str:
    return to_markdown(build(DocxDocument(_docx(body, styles), "a.docx")))


def test_whole_paragraph_marks_are_before_the_text():
    md = _md(f"<w:p>{_r('ふつうの字です。とても長い本文です。')}</w:p>"
             f"<w:p>{_r('消した字。', '<w:strike/>')}</w:p><w:p>{_r('赤い字。', _c('FF0000'))}</w:p>"
             f"<w:p>{_r('二重。', '<w:dstrike/>')}</w:p>")
    assert "〔取り消し線〕消した字。" in md and "〔字の色 #FF0000〕赤い字。" in md and "〔取り消し線〕二重。" in md
    assert "〔字の色 #000000" not in md
    assert "- ふつうの字です。" in md
    assert "〔取り消し線〕・〔字の色 #FF0000〕" in md  # 記号の説明


def test_part_of_a_paragraph_is_named_after_it():
    md = _md(f"<w:p>{_r('前の字と')}{_r('赤', _c('ff0000'))}{_r('と')}{_r('消', '<w:strike/>')}。</w:p>")
    assert "〔うち字の色 #FF0000：「赤」〕" in md and "〔うち取り消し線：「消」〕" in md


def test_default_color_and_strike_off_get_no_mark():
    md = _md(f"<w:p>{_r('自動の色。本文が多い。', _c('auto'))}</w:p><w:p>{_r('黒。', _c('000000'))}</w:p>"
             f"<w:p>{_r('線なし。', NO_STRIKE)}</w:p>")
    assert "〔取り消し線" not in md and "〔字の色" not in md


def test_the_most_common_color_of_the_document_is_ordinary():
    md = _md(f"<w:p>{_r('青い字がいちばん多い本文です。', _c('0000FF'))}</w:p><w:p>{_r('黒。', _c('000000'))}</w:p>")
    assert "〔字の色 #0000FF" not in md and "〔字の色 #000000〕黒。" in md


def test_marks_come_through_styles():
    md = _md('<w:p><w:pPr><w:pStyle w:val="Red"/></w:pPr>' + _r("段落のスタイルが赤。") + "</w:p>"
             + "<w:p>" + _r("字のスタイルが線。", '<w:rStyle w:val="Gone"/>') + "</w:p>" + "<w:p>" + _r("ふつうの字がいちばん多い本文です。とても長い。") + "</w:p>")
    assert "〔字の色 #FF0000〕段落のスタイルが赤。" in md and "〔取り消し線〕字のスタイルが線。" in md


def test_direct_format_overrides_style():
    md = _md('<w:p><w:pPr><w:pStyle w:val="Red"/></w:pPr>' + _r("直した字。", '<w:color w:val="auto"/>') + "</w:p>"
             + "<w:p>" + _r("ふつうの字がいちばん多い本文です。") + "</w:p>")
    assert "〔字の色" not in md


def test_doc_sprms():
    assert apply_look((False, None), [(0x0837, bytes([1]))]) == (True, None)
    assert apply_look((False, None), [(0x2A53, bytes([1]))]) == (True, None)
    assert apply_look((True, None), [(0x0837, bytes([0x80]))]) == (True, None)
    assert apply_look((True, None), [(0x0837, bytes([0x81]))]) == (False, None)
    assert apply_look((False, None), [(0x2A42, bytes([6]))]) == (False, ICO[6]) and ICO[6] == "FF0000"
    assert apply_look((False, "FF0000"), [(0x2A42, bytes([0]))])[1] == "000000"
    assert apply_look((False, None), [(0x6870, struct.pack("<BBBB", 0x12, 0xAB, 0xCD, 0))])[1] == "12ABCD"
    assert apply_look((False, "FF0000"), [(0x6870, bytes([0, 0, 0, 0xFF]))])[1] == "000000"
