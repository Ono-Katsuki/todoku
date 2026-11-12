"""大きな .docx の道の「ふつうの字の色」：ふつうの道と同じ色（文書の全体で数えた、字数のいちばん多い色）。

前は、大きな .docx の道は数えずに黒をふつうとしていた：黒でない色がふつうの文書では、ふつうの道と違う字に色の印が付いた。
いまは、黒でない色になりうる字（色の指定・色を決めるスタイルの名指し）がどこにもなければ数えずに黒、あれば、初めて要るときに
全部のまとまりを、ふつうの道と同じ数え方で数える。
"""
import collections
import io
import re
import zipfile

import pytest

from todoku import lazy, limits
from todoku.docx.document import DocxDocument, iter_paras
from todoku.docx.fontfacts import count_colors, most_common_color
from todoku.docx.window import build
from todoku.lazy import docxdoc

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PLAIN = '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/>'
STYLES = (PLAIN + '<w:style w:type="paragraph" w:styleId="Red"><w:rPr><w:color w:val="FF0000"/></w:rPr></w:style>'
          '<w:style w:type="paragraph" w:styleId="RedChild"><w:basedOn w:val="Red"/></w:style>'
          '<w:style w:type="character" w:styleId="Blue"><w:rPr><w:color w:val="0000FF"/></w:rPr></w:style>')


def _docx(body: str, styles: str = STYLES, header: str = "") -> bytes:
    b = io.BytesIO()
    with zipfile.ZipFile(b, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>')
        rels = f'<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>'
        if header:
            rels += f'<Relationship Id="rId2" Type="{REL}/header" Target="header1.xml"/>'
            z.writestr("word/header1.xml", f'<?xml version="1.0"?><w:hdr {NS}>{header}</w:hdr>')
        z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   f'{rels}</Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0"?><w:document {NS}><w:body>{body}<w:sectPr/></w:body></w:document>')
        z.writestr("word/styles.xml", f'<?xml version="1.0"?><w:styles {NS}>{styles}</w:styles>')
    return b.getvalue()


def _p(t: str, rpr: str = "", ppr: str = "") -> str:
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r></w:p>'


def _c(v: str) -> str:
    return f'<w:color w:val="{v}"/>'


def _whole(blob: bytes) -> tuple[collections.Counter, list[str]]:
    """ふつうの道：色ごとの字数と、窓の行。"""
    d = DocxDocument(blob, "a.docx")
    n: collections.Counter = collections.Counter()
    for _k, _part, blocks in d.all_blocks():
        count_colors(iter_paras(blocks), n)
    return n, [ln for st in build(DocxDocument(blob, "a.docx"))["stories"] for ln in st["reading"]]


@pytest.fixture()
def large(tmp_path, monkeypatch):
    """大きな .docx の道で開く（小さなまとまりに分ける）。"""
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 300)
    monkeypatch.setattr(limits, "LAZY_MIN_BYTES", 0)
    monkeypatch.setattr(docxdoc, "CHUNK_BYTES", 700)

    def open_(blob: bytes):
        lazy.clear()
        p = tmp_path / "a.docx"
        p.write_bytes(blob)
        doc = lazy.open_doc(str(p), force=True)
        assert type(doc).__name__ == "DocxLazy"
        return doc

    yield open_
    lazy.clear()  # Windows は、開いたままのファイルを消せない


def _lines(doc) -> tuple[list[str], set[str]]:
    """大きな道で全部を読んだ行（段落の番号を除く）と、各回の頭の知らせ。"""
    body, heads, rng = [], set(), {}
    while True:
        pg = doc.read(rng, 400)
        body += pg.body
        heads.add(pg.text[:len(pg.text) - len("\n".join(pg.body))])
        if pg.next is None:
            return [re.sub(r"^〔段落 \d+(?:〜\d+)?〕", "", ln) for b in body for ln in b.split("\n")], heads
        rng = pg.next


def _no_counting(doc, monkeypatch) -> None:
    def boom(*a, **k):
        raise AssertionError("数えないはず")

    monkeypatch.setattr(doc, "color_counts", boom)


def test_a_document_whose_normal_color_is_not_black(large):
    blob = _docx("".join(_p(f"濃い灰色の段落 {i}。", _c("333333")) for i in range(40)) + _p("黒い段落。") + _p("赤い段落。", _c("FF0000")))
    n, normal = _whole(blob)
    assert most_common_color(n) == "333333" and normal[0] == "濃い灰色の段落 0。" and normal[-2:] == ["〔字の色 #000000〕黒い段落。", "〔字の色 #FF0000〕赤い段落。"]
    doc = large(blob)
    assert len(doc.chunks) > 3  # 黒い段落だけのまとまりもある
    assert doc.normal_font_color == "333333" and doc.color_counts() == n
    got, heads = _lines(doc)
    assert got == normal
    # 印のない字が黒とは限らない：概要と、各回の頭で、ふつうの色を伝える
    assert heads == {"（この文書のふつうの字の色は #333333：字の色の印のない字は、この色）\n\n"}
    assert doc.overview()["normal_font_color"] == "333333" and "ふつうの字の色は #333333" in doc.overview_text()


def test_a_document_without_any_color_is_black_without_counting(large, monkeypatch):
    blob = _docx("".join(_p(f"段落 {i}。") for i in range(40)) + _p("自動の色。", _c("auto")) + _p("黒の指定。", _c("000000")), PLAIN)
    n, normal = _whole(blob)
    assert set(n) == {"000000"}
    doc = large(blob)
    _no_counting(doc, monkeypatch)
    got, heads = _lines(doc)
    assert doc.normal_font_color == "000000" and got == normal and heads == {""}
    assert "normal_font_color" not in doc.overview() and "ふつうの字の色" not in doc.overview_text()


def test_colored_styles_that_the_body_does_not_name_are_not_counted(large, monkeypatch):
    # スタイルの表に色があっても、本文が名指さなければ、黒でない字はない
    blob = _docx("".join(_p(f"段落 {i}。") for i in range(30)))
    assert set(_whole(blob)[0]) == {"000000"}
    doc = large(blob)
    _no_counting(doc, monkeypatch)
    assert doc.normal_font_color == "000000"


@pytest.mark.parametrize("name, styles, para", [
    ("段落のスタイル", STYLES, lambda i: _p(f"段落 {i}。", "", '<w:pStyle w:val="Red"/>')),
    ("基のスタイル", STYLES, lambda i: _p(f"段落 {i}。", "", '<w:pStyle w:val="RedChild"/>')),
    ("字のスタイル", STYLES, lambda i: _p(f"段落 {i}。", '<w:rStyle w:val="Blue"/>')),
    ("既定の段落スタイル", '<w:style w:type="paragraph" w:default="1" w:styleId="N"><w:rPr><w:color w:val="00B050"/></w:rPr></w:style>',
     lambda i: _p(f"段落 {i}。")),
    ("文書の既定", '<w:docDefaults><w:rPrDefault><w:rPr><w:color w:val="7030A0"/></w:rPr></w:rPrDefault></w:docDefaults>' + PLAIN,
     lambda i: _p(f"段落 {i}。")),
])
def test_color_that_comes_from_styles(large, name, styles, para):
    blob = _docx("".join(para(i) for i in range(30)) + _p("黒い段落。", _c("000000")), styles)
    n, normal = _whole(blob)
    assert most_common_color(n) != "000000", name
    doc = large(blob)
    assert doc.normal_font_color == most_common_color(n) and doc.color_counts() == n
    assert _lines(doc)[0] == normal and normal[-1] == "〔字の色 #000000〕黒い段落。"


def test_colored_text_that_is_not_the_majority(large):
    blob = _docx("".join(_p(f"黒い段落 {i}。") for i in range(30)) + "".join(_p(f"赤い段落 {i}。", _c("FF0000")) for i in range(8)))
    n, normal = _whole(blob)
    doc = large(blob)
    assert doc._normal is None  # 色のある文書：数えて決める
    assert doc.normal_font_color == "000000" and doc.color_counts() == n
    got, heads = _lines(doc)
    assert got == normal and heads == {""} and "〔字の色 #FF0000〕赤い段落 0。" in got


def test_hidden_and_deleted_text_is_not_counted(large):
    hidden = "".join(_p("隠した赤い字がたくさんあります。" * 3, _c("FF0000") + "<w:vanish/>") for _ in range(20))
    gone = ('<w:p><w:del w:id="1" w:author="a" w:date="2026-01-01T00:00:00Z"><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr>'
            '<w:delText>消した赤い字がたくさんあります。消した赤い字がたくさんあります。</w:delText></w:r></w:del></w:p>') * 10
    blob = _docx(hidden + gone + "".join(_p(f"青い段落 {i}。", _c("0000FF")) for i in range(6)) + _p("黒。"))
    n, _normal = _whole(blob)
    doc = large(blob)
    assert doc.color_counts() == n and doc.normal_font_color == most_common_color(n) == "0000FF"


def test_text_outside_the_body_is_counted(large):
    # ヘッダーの字も数える（ふつうの道と同じ）：本文に色がなくても、ヘッダーの色がいちばん多ければ、その色
    header = "".join(_p("ヘッダーの赤い字がたくさんあります。" * 4, _c("FF0000")) for _ in range(6))
    blob = _docx("".join(_p(f"段落 {i}。") for i in range(12)), PLAIN, header)
    n, _normal = _whole(blob)
    assert most_common_color(n) == "FF0000"
    doc = large(blob)
    assert doc.normal_font_color == "FF0000" and doc.color_counts() == n
    blob = _docx("".join(_p(f"段落 {i}。") for i in range(12)), PLAIN, _p("ヘッダー"))
    assert large(blob)._normal == "000000"  # どこにも色がない


def test_the_colors_are_counted_once(large, monkeypatch):
    doc = large(_docx("".join(_p(f"赤い段落 {i}。", _c("FF0000")) for i in range(30))))
    calls = []
    real = doc.chunk_doc
    monkeypatch.setattr(doc, "chunk_doc", lambda i, real_=(): calls.append(i) or real(i, real_))
    assert doc.normal_font_color == "FF0000"
    first = len(calls)
    assert first == sum(1 for ch in doc.chunks if "story" not in ch)
    assert doc.normal_font_color == "FF0000" and doc.overview()["normal_font_color"] == "FF0000" and len(calls) == first
