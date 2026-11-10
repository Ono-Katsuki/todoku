"""記号の書体（Symbol・Wingdings・Wingdings 2・Wingdings 3・Webdings・Zapf Dingbats）で書いた字：書体ごとの対応表で読む。

対応表は src/todoku/symfont.py（Unicode の公開の対応：WG2 N4384、Adobe の symbol.txt・zdingbat.txt）。
- 対応表のある書体の字は、その字（Wingdings 2 の 0xA3 は □ の仲間、Symbol の m は μ）
- 文書が「記号の書体」と宣言した、対応表のない書体の字は、私用の領域の字（外字：読む B）
- 記号の書体でない書体の字は、そのまま
- 印をつける・外すときは、その書体の番号を書く（Wingdings は 0xFE・0xA8、Wingdings 2 は 0x52・0xA3）
"""

from __future__ import annotations

import io
import os
import zipfile

import pytest
from test_openformats import NS as ODF_NS
from test_openformats import B, make_rtf, sj

from todoku import api, symfont
from todoku.grade import grade_blob

FIXTURES = os.path.join(os.path.dirname(__file__), "fixtures")
W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


@pytest.fixture(autouse=True)
def clean():
    api._WINDOWS.clear()
    yield
    api._WINDOWS.clear()


# ====================================================================== 対応表
def test_tables_give_the_published_characters():
    assert symfont.decode("Symbol", 0x6D) == "μ" and symfont.decode("symbol", 0xF0B7) == "•"
    assert symfont.decode("Wingdings", 0xA8) == "\U0001F78E" and symfont.decode("Wingdings", 0xFE) == "\U0001F5F9"
    assert symfont.decode("Wingdings 2", 0xA3) == "□" and symfont.decode("Wingdings 2", 0x52) == "☑"
    assert symfont.decode("Wingdings  2", 0x52) == symfont.decode("wingdings 2", 0xF052)  # 名前は大文字・小文字と空白の数を見ない
    assert symfont.decode("Wingdings", 0xA8) != symfont.decode("Wingdings 2", 0xA8)  # 書体ごとに別の表
    assert symfont.decode("Arial", 0xA9) == "©"  # 記号の書体でない書体：字のまま
    # 対応表のない書体：文書が記号の書体と宣言したときだけ、私用の領域の字
    assert symfont.decode("Marlett", 0x61) == "a" and symfont.decode("Marlett", 0x61, {"marlett"}) == "\uf061"
    assert symfont.text("a b", "Symbol") == "α β"  # 空白は空白のまま
    assert symfont.box_codes("Wingdings") == (0xFE, 0xA8) and symfont.box_codes("Wingdings 2") == (0x52, 0xA3)
    assert symfont.box_codes("Symbol") is None


# ====================================================================== Word（.docx）
FONT_TABLE = ('<w:font w:name="Wingdings"><w:charset w:val="02"/></w:font><w:font w:name="Symbol"><w:charset w:val="02"/></w:font>'
              '<w:font w:name="Wingdings 2"><w:charset w:val="02"/></w:font><w:font w:name="MyDings"><w:charset w:val="02"/></w:font>'
              '<w:font w:name="Arial"><w:charset w:val="00"/></w:font>')
STYLES = ('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
          '<w:style w:type="paragraph" w:styleId="Greek"><w:name w:val="Greek"/><w:basedOn w:val="Normal"/>'
          '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:style>')
NUMBERING = ('<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A8;"/>'
             '<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
             '<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A3;"/>'
             '<w:rPr><w:rFonts w:ascii="Wingdings 2" w:hAnsi="Wingdings 2" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
             '<w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/>'
             '<w:rPr><w:rFonts w:ascii="MyDings" w:hAnsi="MyDings" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>'
             '<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>')


def make_docx(body: str, fonts: str = FONT_TABLE) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   f'<Relationship Id="rS" Type="{REL}/styles" Target="styles.xml"/>'
                   f'<Relationship Id="rN" Type="{REL}/numbering" Target="numbering.xml"/>'
                   f'<Relationship Id="rF" Type="{REL}/fontTable" Target="fontTable.xml"/></Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}</w:body></w:document>')
        z.writestr("word/styles.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:styles {W}>{STYLES}</w:styles>')
        z.writestr("word/numbering.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:numbering {W}>{NUMBERING}</w:numbering>')
        if fonts:
            z.writestr("word/fontTable.xml", f'<?xml version="1.0" encoding="UTF-8"?><w:fonts {W}>{fonts}</w:fonts>')
    return buf.getvalue()


def t(text: str, font: str = "") -> str:
    rpr = f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/></w:rPr>' if font else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def sym(font: str, code: str) -> str:
    return f'<w:r><w:sym w:font="{font}" w:char="{code}"/></w:r>'


def wp(*runs: str, ppr: str = "") -> str:
    return f"<w:p>{ppr}{''.join(runs)}</w:p>"


def bullet(num: int, text: str) -> str:
    return wp(t(text), ppr=f'<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num}"/></w:numPr></w:pPr>')


def _plain(path) -> list[str]:
    api._WINDOWS.clear()
    win = api.window(str(path))[1]
    return [x for st in win["stories"] for x in st["_plain"]]


def _fields(path) -> dict:
    api._WINDOWS.clear()
    return {f["id"]: f for f in api.fields(str(path))}


BODY = (wp(t("単位："), sym("Symbol", "F06D"), t("g　"), sym("Arial", "00A9"), t(" 2026"))
        + wp(t("abg", "Symbol"), t("　の値"))
        + wp(t("pr"), t("の積"), ppr='<w:pPr><w:pStyle w:val="Greek"/></w:pPr>')
        + wp(sym("Wingdings 2", "F0A3"), t("新規　"), sym("Wingdings", "F0A8"), t("変更"))
        + wp(t("印："), t("A", "MyDings"), sym("MyDings", "F042"))
        + bullet(1, "郵送を希望する") + bullet(2, "窓口で受け取る") + bullet(3, "そのほか"))
# 窓の字：μ・αβγ・πρ（Symbol）、□（Wingdings 2 の 0xA3）、U+1F78E（Wingdings の 0xA8）、対応表のない MyDings は私用の領域の字
READ = ["単位：μg　© 2026", "αβγ　の値", "πρの積", "□新規　\U0001F78E変更", "印：", "郵送を希望する", "窓口で受け取る", "そのほか"]


def _pua(g) -> list[str]:
    return [w for r in g.reasons if r.title.startswith("外字") for w in r.where]


def test_docx_reads_symbol_font_characters_by_the_font(tmp_path):
    from todoku.docx.audit import run
    from todoku.docx.document import DocxDocument

    blob = make_docx(BODY)
    src = tmp_path / "a.docx"
    src.write_bytes(blob)
    assert _plain(src) == READ
    assert [p.label for p in DocxDocument(blob).body][5:] == ["\U0001F78E", "□", ""]
    by = _fields(src)
    # 記号の字の □ も、箇条書きの印の □ も、印をつけられる所（対応表のない書体の印は、所にしない）
    assert {k for k, f in by.items() if f["kind"] == "checkbox"} == {"P4.c0", "P4.c2", "P6.b", "P7.b"}
    assert by["P4.c0"]["label"] == "新規" and by["P4.c2"]["label"] == "変更"
    g = grade_blob(blob, "a.docx")
    assert g.grade == "B" and _pua(g) == ["5 段落目「印：〔U+F041〕〔U+F042〕」", "段落の印「〔U+F0A7〕」：1 段落（8 段落目から）"]
    au = run(blob, "a.docx")
    assert au.ok, [(c.name, c.detail) for c in au.checks if not c.ok]


def test_docx_font_not_declared_as_symbol_keeps_its_characters(tmp_path):
    # 書体の一覧（fontTable）のない文書：MyDings は記号の書体と分からないので、「A」は A のまま
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(BODY, fonts=""))
    got = _plain(src)
    assert got[4] == "印：A" and got[:4] == READ[:4]  # 対応表のある書体は、宣言がなくても読む


def test_docx_checks_a_box_with_the_code_of_its_font(tmp_path):
    from todoku.docx.audit import run

    src, out, back = tmp_path / "a.docx", tmp_path / "o.docx", tmp_path / "b.docx"
    src.write_bytes(make_docx(BODY))
    res = api.fill(str(src), {"P4.c0": True, "P4.c2": True, "P6.b": True, "P7.b": True}, str(out))
    assert res["proof_ok"] and res["read_back_ok"] and not res.get("unexpected"), res
    with zipfile.ZipFile(out) as z:
        doc, num = z.read("word/document.xml").decode(), z.read("word/numbering.xml").decode()
    # Wingdings 2 は 0x52、Wingdings は 0xFE（同じ番号を書くと、別の書体では別の形になる）
    assert '<w:sym w:font="Wingdings 2" w:char="F052"/>' in doc and '<w:sym w:font="Wingdings" w:char="F0FE"/>' in doc
    assert '<w:lvlText w:val=""/><w:rPr><w:rFonts w:ascii="Wingdings"' in num
    assert '<w:lvlText w:val=""/><w:rPr><w:rFonts w:ascii="Wingdings 2"' in num
    assert _plain(out)[3] == "☑新規　\U0001F5F9変更"
    by = _fields(out)
    assert all(by[k]["value"] is True for k in ("P4.c0", "P4.c2", "P6.b", "P7.b"))
    assert run(out.read_bytes(), "o.docx").ok

    # 印のついた □ を外す：その書体の、印のない □ の番号に戻す
    src.write_bytes(make_docx(wp(sym("Wingdings 2", "F052"), t("新規　"), sym("Wingdings", "F0FE"), t("変更"))))
    api._WINDOWS.clear()
    by = _fields(src)
    assert by["P1.c0"]["value"] is True and by["P1.c2"]["value"] is True
    res = api.fill(str(src), {"P1.c0": False, "P1.c2": False}, str(back))
    assert res["proof_ok"] and res["read_back_ok"], res
    with zipfile.ZipFile(back) as z:
        doc = z.read("word/document.xml").decode()
    assert '<w:sym w:font="Wingdings 2" w:char="F0A3"/>' in doc and '<w:sym w:font="Wingdings" w:char="F0A8"/>' in doc
    assert _plain(back)[0] == "□新規　\U0001F78E変更"


# ====================================================================== Word 97-2003（.doc）
def test_doc_reads_symbol_font_characters_by_the_font():
    # tests/fixtures/symbol_fonts.doc：上の BODY の .docx を、LibreOffice で .doc にしたもの。
    # 書体の表には、同じ名前の書体が、記号の書体の宣言（chs 2）と、そうでない宣言の 2 つで並ぶ（Calibri・Arial）：
    # 宣言は書体の番号ごとに見る（名前で見ると、Calibri の本文の英数字がすべて外字になる）
    from todoku.doc.audit import run
    from todoku.doc.book import DocFile

    path = os.path.join(FIXTURES, "symbol_fonts.doc")
    blob = open(path, "rb").read()
    ft = DocFile(blob).font_table()
    assert ("Calibri", 2) in ft and any(n == "Calibri" and chs != 2 for n, chs in ft)
    got = _plain(path)
    # LibreOffice は、記号の字「©」（書体 Arial）を、記号の書体と宣言した Arial の字として書く：対応表のない書体なので外字
    assert got == ["単位：μg　 2026", *READ[1:]]
    by = _fields(path)
    assert by["P4.c0"]["kind"] == "checkbox" and by["P4.c2"]["kind"] == "checkbox"
    g = grade_blob(blob, "symbol_fonts.doc")
    assert g.grade == "B" and _pua(g) == ["1 段落目「単位：μg 〔U+F0A9〕 2026」", "5 段落目「印：〔U+F041〕〔U+F042〕」",
                                          "段落の印「〔U+F0A7〕」：1 段落（8 段落目から）"]
    au = run(blob, "symbol_fonts.doc")
    assert au.ok, [(c.name, c.detail) for c in au.checks if not c.ok]


# ====================================================================== RTF
def _rtf(*paras: str) -> bytes:
    f = B + "f"
    head = ("{" + B + "rtf1" + B + "ansi" + B + "ansicpg932" + B + "deff0" + B + "uc1{" + B + "fonttbl{" + f + "0" + B + "fcharset128 "
            + sj("ＭＳ 明朝") + ";}{" + f + "1" + B + "fcharset2 Symbol;}{" + f + "2" + B + "fcharset2 Wingdings 2;}{" + f + "3" + B
            + "fcharset2 MyDings;}{" + f + "4" + B + "fcharset0 Calibri;}{" + f + "5" + B + "fcharset2 Calibri;}{" + f + "6" + B
            + "fcharset0 Wingdings;}}")
    return "\r\n".join([head, *(B + "pard" + B + "plain" + B + "f0 " + x + B + "par" for x in paras), "}"]).encode("latin-1")


def test_rtf_reads_symbol_font_characters_by_the_font(tmp_path):
    from todoku.rtf.audit import run

    f, u = B + "f", B + "u"
    paras = [
        sj("単位：") + "{" + f + "1 " + B + "'6d}g",  # Symbol の 0x6D
        "{" + f + "1 abg}" + sj("の値"),  # Symbol の英字
        "{" + f + "2 " + u + "-3933?}" + sj("新規　") + "{" + f + "6 " + u + "-3928?}" + sj("変更"),  # U+F0A3（Wingdings 2）・U+F0A8（Wingdings）
        sj("印：") + "{" + f + "3 A}",  # 記号の書体と宣言した、対応表のない書体
        "{" + f + "4 Form 12}{" + f + "5 B}",  # 同じ名前の書体：記号の書体の宣言のない方の字は、そのまま
    ]
    blob = _rtf(*paras)
    src = tmp_path / "a.rtf"
    src.write_bytes(blob)
    assert _plain(src) == ["単位：μg", "αβγの値", "□新規　\U0001F78E変更", "印：", "Form 12"]
    by = _fields(src)
    assert by["P3.c0"]["kind"] == "checkbox" and by["P3.c2"]["kind"] == "checkbox"
    g = grade_blob(blob, "a.rtf")
    assert g.grade == "B" and _pua(g) == ["4 段落目「印：〔U+F041〕」", "5 段落目「Form 12〔U+F042〕」"]
    au = run(_rtf(*paras[:4]), "a.rtf")
    assert au.ok, [(c.name, c.detail) for c in au.checks if not c.ok]
    out = tmp_path / "o.rtf"
    res = api.fill(str(src), {"P3.c0": True}, str(out))
    assert res["proof_ok"] and res["read_back_ok"], res
    assert _fields(out)["P3.c0"]["value"] is True
    # 記号の書体の字のすぐ後ろに書いた字は、その書体を引き継がない（下の test_written_text_does_not_inherit_a_symbol_font）
    res = api.fill(str(src), {"P5.c7": "12"}, str(tmp_path / "x.rtf"))
    assert res["proof_ok"] and res["read_back_ok"], res
    assert _plain(tmp_path / "x.rtf")[4] == "Form 12\uf04212"


# ====================================================================== OpenDocument の文書（.odt）
ODT_FACES = ('<style:font-face style:name="Symbol" svg:font-family="Symbol" style:font-charset="x-symbol"/>'
             '<style:font-face style:name="Wingdings" svg:font-family="Wingdings" style:font-charset="x-symbol"/>'
             '<style:font-face style:name="MyDings" svg:font-family="MyDings" style:font-charset="x-symbol"/>'
             # 同じ書体の名前に、記号の書体の宣言と、そうでない宣言（LibreOffice が書く形）
             '<style:font-face style:name="Calibri" svg:font-family="Calibri" style:font-charset="x-symbol"/>'
             '<style:font-face style:name="Calibri1" svg:font-family="Calibri"/>')
ODT_STYLES = ('<style:style style:name="Body" style:family="paragraph"><style:text-properties style:font-name="Calibri1"/></style:style>'
              '<style:style style:name="Greek" style:family="paragraph"><style:text-properties style:font-name="Symbol"/></style:style>'
              '<style:style style:name="T1" style:family="text"><style:text-properties style:font-name="Symbol"/></style:style>'
              # 書体の宣言を通さず、字の書式そのものに書いた記号の書体
              '<style:style style:name="T2" style:family="text"><style:text-properties fo:font-family="&apos;Wingdings 2&apos;" '
              'style:font-charset="x-symbol"/></style:style>'
              '<style:style style:name="T3" style:family="text"><style:text-properties style:font-name="MyDings"/></style:style>'
              '<style:style style:name="T4" style:family="text"><style:text-properties style:font-name="Calibri"/></style:style>'
              '<style:style style:name="T5" style:family="text"><style:text-properties fo:font-family="Calibri"/></style:style>'
              '<text:list-style style:name="L1"><text:list-level-style-bullet text:level="1" text:bullet-char="&#xF0A8;">'
              '<style:text-properties style:font-name="Wingdings"/></text:list-level-style-bullet></text:list-style>')
ODT_TEXT = ('<text:p text:style-name="Body">単位：<text:span text:style-name="T1">&#xF06D;</text:span>g</text:p>'
            '<text:p text:style-name="Greek">pr<text:span text:style-name="T5">x</text:span>の積</text:p>'
            '<text:p text:style-name="Body"><text:span text:style-name="T2">&#xF0A3;</text:span>新規　'
            '<text:span text:style-name="T2">&#xF0A3;</text:span>変更</text:p>'
            '<text:p text:style-name="Body">印：<text:span text:style-name="T3">A</text:span><text:span text:style-name="T4">B</text:span></text:p>'
            '<text:list text:style-name="L1"><text:list-item><text:p text:style-name="Body">郵送を希望する</text:p></text:list-item></text:list>')


def _odt(text: str = ODT_TEXT) -> bytes:
    mime = "application/vnd.oasis.opendocument.text"
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), mime)
        z.writestr("META-INF/manifest.xml", '<?xml version="1.0" encoding="UTF-8"?><manifest:manifest '
                   'xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">'
                   f'<manifest:file-entry manifest:full-path="/" manifest:media-type="{mime}"/>'
                   '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
                   zipfile.ZIP_DEFLATED)
        z.writestr("content.xml", f'<?xml version="1.0" encoding="UTF-8"?><office:document-content {ODF_NS}>'
                   f'<office:font-face-decls>{ODT_FACES}</office:font-face-decls><office:automatic-styles>{ODT_STYLES}'
                   f'</office:automatic-styles><office:body><office:text>{text}</office:text></office:body></office:document-content>',
                   zipfile.ZIP_DEFLATED)
    return buf.getvalue()


def test_odt_reads_symbol_font_characters_by_the_font(tmp_path):
    from todoku.odf.text import OdtDocument
    from todoku.odf.textaudit import run

    blob = _odt()
    src = tmp_path / "a.odt"
    src.write_bytes(blob)
    # 段落のスタイルの書体（Greek）は段落の字に効き、字のスタイルの書体（T5：記号の書体でない）がそれを上書きする
    assert _plain(src) == ["単位：μg", "πρxの積", "□新規　□変更", "印：", "郵送を希望する"]
    assert OdtDocument(blob, "a.odt").body[4].label == "\U0001F78E"
    by = _fields(src)
    assert by["P3.c0"]["kind"] == "checkbox" and by["P3.c2"]["kind"] == "checkbox"
    g = grade_blob(blob, "a.odt")
    assert g.grade == "B" and _pua(g) == ["4 段落目「印：〔U+F041〕〔U+F042〕」"]
    au = run(blob, "a.odt")
    assert au.ok, [(c.name, c.detail) for c in au.checks if not c.ok]
    out = tmp_path / "o.odt"
    res = api.fill(str(src), {"P3.c2": True}, str(out))
    assert res["proof_ok"] and res["read_back_ok"], res
    assert _fields(out)["P3.c2"]["value"] is True and _fields(out)["P3.c0"]["value"] is False


# ====================================================================== 書いた字は、記号の書体を引き継がない
def _write(src, old: str, new: str, out) -> dict:
    api._WINDOWS.clear()
    return api.fill(str(src), {"@1": {"old": old, "new": new}}, str(out))


def _told(res: dict, font: str) -> bool:
    return any(f"記号の書体（{font}）" in w and "引き継がず" in w for w in res["warnings"])


def test_written_text_does_not_inherit_a_symbol_font(tmp_path):
    # 記号の書体の字のすぐ後ろに書いた字が、その書体を引き継ぐと、書いた字が、その書体の別の字（絵・ギリシア文字）になる。
    # 取り消し線・字の色と同じに、書いた字には引き継がず、そのことを伝える
    # ---- .docx：字の書式（w:rFonts の ascii・hAnsi）の記号の書体だけを外し、ほかの書式は残す
    body = (wp(t("単位："), t("m", "Symbol"))
            + wp(t("印："), '<w:r><w:rPr><w:rFonts w:ascii="MyDings" w:hAnsi="MyDings" w:eastAsia="ＭＳ 明朝"/><w:b/></w:rPr><w:t>A</w:t></w:r>')
            + wp(t("pr"), t("の積"), ppr='<w:pPr><w:pStyle w:val="Greek"/></w:pPr>'))
    src = tmp_path / "a.docx"
    src.write_bytes(make_docx(body))
    res = _write(src, "単位：μ", "単位：μAb", tmp_path / "o1.docx")
    assert res["proof_ok"] and res["read_back_ok"] and _told(res, "Symbol"), res
    assert _plain(tmp_path / "o1.docx")[0] == "単位：μAb"
    res = _write(src, "印：\uf041", "印：\uf041Ab", tmp_path / "o2.docx")
    assert res["proof_ok"] and res["read_back_ok"] and _told(res, "MyDings"), res
    xml = zipfile.ZipFile(tmp_path / "o2.docx").read("word/document.xml").decode("utf-8")
    assert '<w:r><w:rPr><w:rFonts w:eastAsia="ＭＳ 明朝"/><w:b/></w:rPr><w:t xml:space="preserve">Ab</w:t></w:r>' in xml
    # 段落のスタイルが決める記号の書体は外せない：書いた字と読み直した字が違うので、書き出さない
    res = _write(src, "πρ", "πρAb", tmp_path / "o3.docx")
    assert not res["read_back_ok"] and res["read_back_differences"] == [{"id": "P3.w1", "written": "Ab", "read_back": "Αβ"}]
    assert not (tmp_path / "o3.docx").exists()
    # ---- .doc：字の書体の命令（sprmCRgFtc0・sprmCRgFtc2）のうち、記号の書体のものを外す
    doc = os.path.join(FIXTURES, "symbol_fonts.doc")
    res = _write(doc, "αβγ", "αβγAb", tmp_path / "o1.doc")
    assert res["proof_ok"] and res["read_back_ok"] and _told(res, "Symbol"), res
    assert _plain(tmp_path / "o1.doc")[1] == "αβγAb　の値"
    res = _write(doc, "\uf041\uf042", "\uf041\uf042Ab", tmp_path / "o2.doc")
    assert res["proof_ok"] and res["read_back_ok"] and _told(res, "MyDings"), res
    # ---- RTF：記号の書体の前に選ばれていた書体（なければ既定の書体 \deff）に戻す
    f = B + "f"
    src = tmp_path / "a.rtf"
    src.write_bytes(_rtf(sj("印：") + f + "1 abg"))
    res = _write(src, "αβγ", "αβγAb", tmp_path / "o1.rtf")
    assert res["proof_ok"] and res["read_back_ok"] and _told(res, "Symbol") and "ＭＳ 明朝" in res["warnings"][0], res
    assert b"{" + (f + "0 Ab}").encode() in (tmp_path / "o1.rtf").read_bytes()
    assert _plain(tmp_path / "o1.rtf") == ["印：αβγAb"]
    # 戻す書体がない（既定の書体の指定がなく、その前に書体を選んでいない）：そのまま書き、読み直した字が違うので、書き出さない
    src = tmp_path / "b.rtf"
    src.write_bytes(_rtf(sj("印：") + f + "1 abg").replace((B + "deff0").encode(), b"").replace((B + "plain" + f + "0 ").encode(), (B + "plain ").encode()))
    assert _plain(src) == ["印：αβγ"]
    res = _write(src, "αβγ", "αβγAb", tmp_path / "o2.rtf")
    assert not res["read_back_ok"] and res["read_back_differences"] == [{"id": "P1.w1", "written": "Ab", "read_back": "Αβ"}]
    assert not (tmp_path / "o2.rtf").exists()
    # ---- .odt：記号の書体の外側で効いていた書体を、打ち消すスタイルに写す
    src = tmp_path / "a.odt"
    src.write_bytes(_odt())
    res = _write(src, "単位：μ", "単位：μAb", tmp_path / "o1.odt")
    assert res["proof_ok"] and res["read_back_ok"] and _told(res, "Symbol"), res
    xml = zipfile.ZipFile(tmp_path / "o1.odt").read("content.xml").decode("utf-8")
    assert '<style:style style:name="todoku_w1" style:family="text"><style:text-properties style:font-name="Calibri1"/></style:style>' in xml
    assert _plain(tmp_path / "o1.odt")[0] == "単位：μAbg"
    # 外側に、記号でない書体がない（段落のスタイルが記号の書体で、既定のスタイルは書体を決めていない）：書き出さない
    res = _write(src, "πρ", "πρAb", tmp_path / "o2.odt")
    assert not res["read_back_ok"] and not (tmp_path / "o2.odt").exists()


def test_cleaning_the_format_of_written_text_names_the_symbol_font():
    import struct

    from todoku.doc.book import DocFile
    from todoku.doc.fill import clean_chpx
    from todoku.docx.fill import clean_rpr

    # .docx：記号の書体の属性だけを外す。属性のなくなった w:rFonts・子のなくなった w:rPr は置かない
    rpr = b'<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:cs="Arial"/><w:sz w:val="21"/></w:rPr>'
    assert clean_rpr(rpr) == (b'<w:rPr><w:rFonts w:cs="Arial"/><w:sz w:val="21"/></w:rPr>', ["記号の書体（Wingdings）"])
    assert clean_rpr(b'<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>')[0] == b""
    # 文書が記号の書体と宣言した書体（対応表のない書体）は、宣言を渡したときだけ
    mine = b'<w:rPr><w:rFonts w:ascii="MyDings" w:hAnsi="Century"/></w:rPr>'
    assert clean_rpr(mine) == (mine, [])
    assert clean_rpr(mine, frozenset({symfont.key("MyDings")})) == (b'<w:rPr><w:rFonts w:hAnsi="Century"/></w:rPr>', ["記号の書体（MyDings）"])
    # .doc：記号の書体の番号の命令と、「字でないもの」の印（特別な字・書体と番号で書いた記号）を外す。ほかの命令は残す
    doc = DocFile(open(os.path.join(FIXTURES, "symbol_fonts.doc"), "rb").read())
    ft = doc.font_table()
    sym_ftc = next(i for i, (n, _chs) in enumerate(ft) if n == "Symbol")
    plain_ftc = next(i for i, (n, chs) in enumerate(ft) if chs != 2 and not symfont.known(n))
    size = struct.pack("<HH", 0x4A43, 21)
    grp = (struct.pack("<HH", 0x4A4F, sym_ftc) + struct.pack("<HH", 0x4A51, plain_ftc) + size + struct.pack("<HB", 0x0855, 1)
           + struct.pack("<HHHB", 0x6A09, sym_ftc, 0xF06D, 0)[:6])
    assert clean_chpx(grp, doc) == (struct.pack("<HH", 0x4A51, plain_ftc) + size, ["記号の書体（Symbol）"])
    assert clean_chpx(grp)[0] == grp[:12]  # 書体の表がなければ、書体は決めない（印だけ外す）


# ====================================================================== 数式（MathType）・外字の伝え方
def test_equation_characters_use_the_symbol_table():
    from todoku.embeds.mtef import _char

    # 記号・ギリシア文字の書体（4・5・6）の番号は Symbol の対応表で読む（m は μ、j は ϕ）
    assert [_char(face, 0x6D) for face in (4, 5, 6)] == ["μ", "μ", "μ"] and _char(5, 0x6A) == "ϕ"
    assert _char(1, 0x6D) == "m" and _char(1, 0xF06D) == "μ" and _char(4, 0x20) == " "
    assert _char(11, 0x41) == "〔MT Extra の字 0x41〕"  # 対応表を持たない書体の字は、番号で示す


def test_private_use_characters_are_told_in_every_text_format():
    # 外字（私用の領域の字）は、どの形式でも読む B：字の意味は符号からは分からない
    g = grade_blob("氏名：山".encode("utf-8"), "a.txt")
    assert g.grade == "B" and _pua(g) == ["1 段落目「氏名：〔U+E000〕山」"]
    g = grade_blob(make_rtf(B + "pard" + B + "f1 " + sj("氏名：") + B + "u-8192?" + B + "par"), "a.rtf")
    assert g.grade == "B" and _pua(g) == ["1 段落目「氏名：〔U+E000〕」"]
