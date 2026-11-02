"""Word に埋め込んだ Excel の表への書き込みのテスト：ブックのセルと、Word が表示・印刷に使う表の絵（EMF）の両方。

テストの中で、Excel の表を埋め込んだ小さな .docx（.xlsx の部品と、Excel と同じ描き方の EMF の絵）を組み立てて確かめ、
公開の様式（函館市の工賃向上計画の .doc。corpus/pairs、なければ飛ばす）でも確かめる。
"""

from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path

import pytest

from todoku.cfb import CFB, build_cfb
from todoku.docx.document import DocxDocument
from todoku.docx.fill import FillError, fill, written_of
from todoku.docx.proof import prove
from todoku.docx.window import build, to_markdown
from todoku.metafile import TA_RIGHT, Emf, ext_text_out, md4

CORPUS = Path(__file__).resolve().parent.parent / "corpus"
HAKODATE = CORPUS / "pairs" / "www.city.hakodate.hokkaido.jp"


def test_md4_rfc1320_vectors():
    # RFC 1320 の付録の試験値
    assert md4(b"").hex() == "31d6cfe0d16ae931b73c59d7e0c089c0"
    assert md4(b"abc").hex() == "a448017aaf21d8525fc10ae87aa6729d"
    assert md4(b"message digest").hex() == "d9130a8164549fe818874806e1c7014b"
    assert md4(b"1234567890" * 8).hex() == "e33b4ddc9c38f2199c3e7b164fcc0536"


def test_build_cfb_roundtrip():
    items = {"Workbook": b"w" * 5000, "\x01Ole": b"o" * 20, "_VBA_PROJECT_CUR/VBA/dir": b"d" * 100,
             "_VBA_PROJECT_CUR/PROJECT": b"p" * 4200}
    c = CFB(build_cfb(items))
    assert {p: c.read(p) for p in c.streams()} == items


# ====================================================================== テストの中で組み立てる .docx
NS = ('xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def _xlsx(b2=None) -> bytes:
    """3行3列の表：見出し（項目・金額・備考）、工賃の行（B2 が入力欄、#,##0）、合計（B3＝B2*2）。表の絵に映る範囲は A1:C3。"""
    strings = ["項目", "金額", "備考", "工賃", "円", "合計"]
    sst = "".join(f"<si><t>{s}</t></si>" for s in strings)
    b2c = f'<c r="B2" s="1"><v>{b2}</v></c>' if b2 is not None else '<c r="B2" s="1"/>'
    b3v = 0 if b2 is None else b2 * 2
    sheet = (f'<?xml version="1.0" encoding="UTF-8"?><worksheet {NS}><sheetFormatPr defaultRowHeight="15"/>'
             '<cols><col min="1" max="1" width="10" customWidth="1"/><col min="2" max="2" width="12" customWidth="1"/>'
             '<col min="3" max="3" width="8" customWidth="1"/></cols><sheetData>'
             '<row r="1" ht="15" customHeight="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="2" t="s"><v>1</v></c><c r="C1" s="2" t="s"><v>2</v></c></row>'
             f'<row r="2" ht="15" customHeight="1"><c r="A2" t="s"><v>3</v></c>{b2c}<c r="C2" s="2" t="s"><v>4</v></c></row>'
             f'<row r="3" ht="15" customHeight="1"><c r="A3" t="s"><v>5</v></c><c r="B3" s="1"><f>B2*2</f><v>{b3v}</v></c>'
             '<c r="C3" s="2" t="s"><v>4</v></c></row></sheetData></worksheet>')
    styles = (f'<?xml version="1.0" encoding="UTF-8"?><styleSheet {NS}>'
              '<fonts count="1"><font><sz val="11"/><name val="ＭＳ Ｐゴシック"/><family val="3"/><charset val="128"/></font></fonts>'
              '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>'
              '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
              '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              '<xf numFmtId="3" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyProtection="1"><protection locked="0"/></xf>'
              '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center"/></xf>'
              '</cellXfs></styleSheet>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                   '</Relationships>')
        z.writestr("xl/workbook.xml", f'<?xml version="1.0" encoding="UTF-8"?><workbook {NS}><bookViews><workbookView activeTab="0"/></bookViews>'
                   '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets><oleSize ref="A1:C3"/></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                   '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
                   '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
                   '</Relationships>')
        z.writestr("xl/styles.xml", styles)
        z.writestr("xl/sharedStrings.xml", f'<?xml version="1.0" encoding="UTF-8"?><sst {NS} count="7" uniqueCount="6">{sst}</sst>')
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return buf.getvalue()


def _rec(t: int, body: bytes) -> bytes:
    return struct.pack("<II", t, 8 + len(body)) + body


def _font(ih: int, height: int, face: str) -> bytes:
    lf = struct.pack("<iiiii8B", height, 0, 0, 0, 400, 0, 0, 0, 128, 0, 0, 0, 0x31)
    lf += face.encode("utf-16-le").ljust(64, b"\x00")
    elw = lf + b"\x00" * (128 + 64 + 64) + struct.pack("<II", 0x08007664, 0)  # LogFontExDv（設計の軸なし）
    return _rec(0x52, struct.pack("<I", ih) + elw)


def _patblt(x: int, y: int, cx: int, cy: int) -> bytes:
    """元の絵のない塗り（罫線）：EMR_BITBLT、PATCOPY。"""
    return _rec(0x4C, struct.pack("<4i4iIii6fIIIIII", x, y, x + cx - 1, y + cy - 1, x, y, cx, cy, 0x00F00021, 0, 0,
                                  1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0))


def _emf() -> bytes:
    """Excel が 96 dpi で描くのと同じ形の表の絵：列の幅 80・96・64、行の高さ 20、ＭＳ Ｐゴシック 11pt（h=-15）、余白 3。"""
    t = lambda x, y, s, dx: ext_text_out(x, y, s, dx, (x, y, x + sum(dx) - 1, y + 14), 2, (0.0, 0.0))  # noqa: E731
    body = [_rec(0x21, b""), _rec(0x1E, struct.pack("<4i", 0, 0, 241, 61)), _font(1, -15, "ＭＳ Ｐゴシック"),
            _rec(0x25, struct.pack("<I", 1)), _rec(0x12, struct.pack("<I", 1)),
            t(3, 3, "項目", [15, 15]), t(113, 3, "金額", [15, 15]), t(193, 3, "備考", [15, 15]),
            t(3, 23, "工賃", [15, 15]), t(200, 23, "円", [15]), t(3, 43, "合計", [15, 15]), t(165, 43, "0", [8]),
            t(200, 43, "円", [15]), _rec(0x25, struct.pack("<I", 0x8000000D)), _rec(0x22, struct.pack("<i", -1)),
            _rec(0x28, struct.pack("<I", 1))]
    for x in (0, 80, 176, 240):
        body.append(_patblt(x, 0, 1, 61))
    for y in (0, 20, 40, 60):
        body.append(_patblt(0, y, 241, 1))
    body.append(_rec(0x0E, struct.pack("<III", 0, 0, 20)))
    size = 108 + sum(len(r) for r in body)
    head = struct.pack("<II4i4iIIIIHHIIIiiiiIIIII", 1, 108, 0, 0, 240, 60, 0, 0, 6350, 1588, 0x464D4520, 0x10000, size,
                       len(body) + 1, 2, 0, 0, 0, 0, 1366, 768, 344, 194, 0, 0, 0, 344000, 194000)
    return head + b"".join(body)


def _docx(book: bytes, picture: bytes) -> bytes:
    w = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" '
         'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
    body = ('<w:p><w:r><w:t>工賃の表</w:t></w:r></w:p><w:p><w:r><w:object w:dxaOrig="3600" w:dyaOrig="900">'
            '<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:180pt;height:45pt" o:ole="">'
            '<v:imagedata r:id="rId8" o:title=""/></v:shape><o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" '
            'ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1774977835" r:id="rId9"/></w:object></w:r></w:p>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/><Default Extension="emf" ContentType="image/x-emf"/>'
                   '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'
                   '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                   '</Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                   '</Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {w}><w:body>{body}</w:body></w:document>')
        z.writestr("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.emf"/>'
                   '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" '
                   'Target="embeddings/Microsoft_Excel_Worksheet.xlsx"/></Relationships>')
        z.writestr("word/media/image1.emf", picture)
        z.writestr("word/embeddings/Microsoft_Excel_Worksheet.xlsx", book)
    return buf.getvalue()


def test_docx_embedded_table_write_redraws_picture():
    blob = _docx(_xlsx(), _emf())
    win = build(DocxDocument(blob, "e.docx"))
    e = win["embeds"][0]
    assert e["writable"], e["why"]
    md = to_markdown(win)
    assert "`E1!Sheet1!C4`" in md and "書くときの欄の名前" in md
    res = fill(DocxDocument(blob, "e.docx"), {"E1!Sheet1!B2": "１２３４"})  # 全角の数字も、Excel と同じく数値として受け取る
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    assert "入力 1・計算結果 1" in pr.headline() and "表の絵 1 枚" in pr.headline()
    after = DocxDocument(res.blob, "e.docx")
    wb = after.embedded_books()[0][1]
    assert wb.sheet("Sheet1").cells[(2, 2)].value == 1234 and wb.sheet("Sheet1").cells[(3, 2)].value == 2468
    pic = Emf(zipfile.ZipFile(io.BytesIO(res.blob)).read("word/media/image1.emf"))
    runs = {r.text: (r.x, r.y, r.align & 6, r.font.height) for r in pic.runs}
    # 右寄せ（#,##0 の数値）：右端は、元の「0」の右端（セルの右の境 176 − 余白 3）。行の下から 2 だけ上に字の下端
    assert runs["1,234"] == (173, 23, TA_RIGHT, -15) and runs["2,468"] == (173, 43, TA_RIGHT, -15)
    assert "0" not in runs  # 合計の元の字は消した
    assert {"項目", "金額", "備考", "工賃", "合計", "円"} <= set(runs)  # ほかの字はそのまま
    # 窓に読み戻せる
    got = {f["id"]: f["value"] for f in build(after)["embeds"][0]["fields"]}
    assert got.get("E1!Sheet1!B2") == 1234
    # 証明は、絵に書き足された別の字を見つける
    z = zipfile.ZipFile(io.BytesIO(res.blob))
    bad_pic = Emf(z.read("word/media/image1.emf"))
    tampered = bad_pic.rebuild(set(), [ext_text_out(10, 10, "X", None, (10, 10, 20, 20), 2, (0.0, 0.0))], bad_pic.n_handles)
    from todoku.zipraw import RawZip

    rz = RawZip(res.blob)
    rz.replace("word/media/image1.emf", tampered)
    assert not prove(blob, rz.build(), written_of(res.patches)).ok


def test_docx_embedded_table_refuses_what_it_cannot_draw():
    blob = _docx(_xlsx(), _emf())
    # 幅の分からない字が多く、欄に入るか確かめられない（入らなければ Excel では ####）：書かない
    with pytest.raises(FillError, match="入るか確かめられません"):
        fill(DocxDocument(blob, "e.docx"), {"E1!Sheet1!B2": 123456789012})
    with pytest.raises(FillError, match="埋め込み2 はありません"):
        fill(DocxDocument(blob, "e.docx"), {"E2!Sheet1!B2": 1})
    # 絵が EMF でない（写真）：書けないと窓に出し、書き込みも断る
    png = b"\x89PNG\r\n\x1a\n" + b"\x00" * 64
    blob2 = _docx(_xlsx(), png)
    e = build(DocxDocument(blob2, "e.docx"))["embeds"][0]
    assert not e["writable"] and "EMF" in e["why"]
    with pytest.raises(FillError, match="EMF"):
        fill(DocxDocument(blob2, "e.docx"), {"E1!Sheet1!B2": 5})


def test_docx_embedded_table_shown_as_icon():
    # アイコンで表示している埋め込み：セルに書いても Word に見える絵（アイコン）は変わらないので、ブックだけを書く
    png = b"\x89PNG\r\n\x1a\n" + b"\x00" * 64
    blob = _docx(_xlsx(), png)
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            data = src.read(n)
            z.writestr(n, data.replace(b'DrawAspect="Content"', b'DrawAspect="Icon"') if n == "word/document.xml" else data)
    blob = buf.getvalue()
    assert build(DocxDocument(blob, "e.docx"))["embeds"][0]["writable"]
    res = fill(DocxDocument(blob, "e.docx"), {"E1!Sheet1!B2": 7})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected
    z = zipfile.ZipFile(io.BytesIO(res.blob))
    assert z.read("word/media/image1.emf") == png


# ====================================================================== 公開の様式（函館市・.doc）
def _hakodate(name: str) -> bytes:
    p = HAKODATE / name
    if not p.exists():
        pytest.skip("corpus なし")
    return p.read_bytes()


def test_doc_embedded_table_write_redraws_both_pictures():
    from todoku.doc.audit import prove as prove_doc
    from todoku.doc.book import DocFile
    from todoku.doc.fill import fill as fill_doc
    from todoku.embed import Picture

    blob = _hakodate("b8e79d21c251_blank.doc")
    doc = DocFile(blob, "a.doc")
    loc = doc.embeds[0]["pictures"][0][1]
    res = fill_doc(doc, {"E1!Sheet1!C4": 9418000, "E1!Sheet1!D4": 9430000, "E1!Sheet1!F7": 5210000})
    pr = prove_doc(blob, res.blob, written_of(res.patches))
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    after = DocFile(res.blob, "a.doc")
    e = after.embeds[0]
    # 表示用の絵は Data の後ろに足し、絵の字の位置（sprmCPicLocation）だけを書き換えた。元の絵はそのまま残る
    new_loc = e["pictures"][0][1]
    assert new_loc == len(doc.data) and after.data[:len(doc.data)] == doc.data
    for emf in (Emf(after.cfb.read(e["storage"] + "/\x03EPRINT")), Emf(Picture(after.data, new_loc).metafile())):
        runs = {}
        for r in emf.runs:
            runs.setdefault(r.text, []).append(r)
        assert "9418000" in runs and "9430000" in runs and "5210000" in runs and "0%" in runs
    pr_ep = Emf(after.cfb.read(e["storage"] + "/\x03EPRINT"))
    c4 = next(r for r in pr_ep.runs if r.text == "9418000")
    d4 = next(r for r in pr_ep.runs if r.text == "9430000")
    # C4：元の「0」と同じ右端（2203）、同じ高さ。D4：縮小して全体を表示するセルなので、Excel と同じ規則で 92 → 85
    assert (c4.x, c4.y, c4.font.height) == (2203, 529, -92) and d4.font.height == -85
    old_pic = Picture(doc.data, loc)
    assert Picture(after.data, loc).metafile() == old_pic.metafile()
    assert md4(Picture(after.data, new_loc).metafile()) != md4(old_pic.metafile())
    # 書いた後の文書に、もう一度書ける（前に描き足した字も、セルの字として見つけて描き直す）。Word の欄と一緒にも書ける
    fid = next(f["id"] for f in build(after)["fields"] if f["kind"] == "text" and not f.get("agent"))
    res2 = fill_doc(DocFile(res.blob, "a.doc"), {"E1!Sheet1!D4": 100, fid: "とどく"})
    pr2 = prove_doc(res.blob, res2.blob, written_of(res2.patches))
    assert pr2.ok, pr2.unexpected
    win2 = build(DocFile(res2.blob, "a.doc"))
    vals = {f["id"]: f["value"] for f in win2["embeds"][0]["fields"]}
    assert vals["E1!Sheet1!D4"] == 100 and vals["E1!Sheet1!C4"] == 9418000
    assert {f["id"]: f["value"] for f in win2["fields"]}[fid] == "とどく"
    ep2 = {r.text for r in Emf(DocFile(res2.blob).cfb.read(e["storage"] + "/\x03EPRINT")).runs}
    assert "100" in ep2 and "9430000" not in ep2 and "9418000" in ep2


def test_docx_ole_bin_with_excel_pictures():
    # .docx に .xls の表を OLE の入れ物（oleObject1.bin）で埋め込んだもの：中身と絵は函館市の様式の埋め込みから作る
    from todoku.cfb import SubStorage
    from todoku.doc.book import DocFile
    from todoku.embed import Picture

    d = DocFile(_hakodate("b8e79d21c251_blank.doc"), "a.doc")
    e = d.embeds[0]
    binary = build_cfb(SubStorage(d.cfb, e["storage"]).contents())  # Workbook と \003EPRINT など
    base = _docx(b"", Picture(d.data, e["pictures"][0][1]).metafile())
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(base)) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n in src.namelist():
            data = src.read(n)
            if n.endswith(".xlsx"):
                n, data = "word/embeddings/oleObject1.bin", binary
            elif n == "word/_rels/document.xml.rels":
                data = data.replace(b"Microsoft_Excel_Worksheet.xlsx", b"oleObject1.bin").replace(b"relationships/package",
                                                                                                b"relationships/oleObject")
            elif n == "word/document.xml":
                data = data.replace(b"Excel.Sheet.12", b"Excel.Sheet.8")
            z.writestr(n, data)
    blob = buf.getvalue()
    assert build(DocxDocument(blob, "x.docx"))["embeds"][0]["writable"]
    res = fill(DocxDocument(blob, "x.docx"), {"E1!Sheet1!C4": 9418000, "E1!Sheet1!E4": 9440000})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected
    assert "表の絵 2 枚" in pr.headline()  # 入れ物の中の \003EPRINT と、表示用の image1.emf
    c = CFB(zipfile.ZipFile(io.BytesIO(res.blob)).read("word/embeddings/oleObject1.bin"))
    assert c.read("\x01CompObj") == CFB(binary).read("\x01CompObj")  # ほかのストリームはそのまま
    assert "9440000" in {r.text for r in Emf(c.read("\x03EPRINT")).runs}


def test_layout_reproduces_excel_drawing_of_the_filled_example():
    # 記入例（Excel が描いた絵）の字を、ここでの規則で描き直した位置・大きさと比べる（縮小して全体を表示するセルも含む）
    from todoku.doc.book import DocFile
    from todoku.embed import Picture
    from todoku.sheetpic import Layout, SheetModel

    doc = DocFile(_hakodate("db857f87dd1e_example.doc"), "b.doc")
    e, wb = doc.embedded_books()[0]
    model = SheetModel(wb)
    pics = [Emf(doc.cfb.read(e["storage"] + "/\x03EPRINT")), Emf(Picture(doc.data, e["pictures"][0][1]).metafile())]
    for emf in pics:
        lay = Layout(emf, model)
        assert len(lay.anchors) >= 75
        shrunk = 0
        for idx, rc in lay.anchors.items():
            run = lay._run(idx)
            p = lay.draw(rc, lay.texts[rc], 99, lay.texts)
            h = next(struct.unpack_from("<i", r, 12)[0] for r in p.records if struct.unpack_from("<I", r, 0)[0] == 0x52)
            al = next(struct.unpack_from("<I", r, 8)[0] for r in p.records if struct.unpack_from("<I", r, 0)[0] == 0x16)
            want_x = run.x + (run.width if al == 2 else run.width / 2 if al == 6 else 0)
            assert abs(p.ref[0] - want_x) <= 3 and abs(p.ref[1] - run.y) <= 4, (lay.texts[rc].text, p.ref, (want_x, run.y))
            assert abs(abs(h) - abs(run.font.height)) <= 1, lay.texts[rc].text
            if lay.texts[rc].shrink and abs(run.font.height) < abs(lay.font_scale() * 11) - 1:
                shrunk += 1
                if set(run.text) <= set("0123456789,.%"):
                    assert h == run.font.height, run.text  # 数値を縮小した大きさは Excel と同じ（字の幅がすべて分かる）
        assert shrunk >= 10


# ====================================================================== 公開の様式（corpus/wide_embed）
WIDE = CORPUS / "wide_embed"


def _wide(name: str) -> bytes:
    p = WIDE / name
    if not p.exists():
        pytest.skip("corpus なし")
    return p.read_bytes()


def test_doc_inline_picture_blip_is_found_by_property_order():
    # 表示用の絵の pib の値（4）は、図形の中の絵の入れ物の数（1）を超える：入れ物は BLIP の性質の順に並ぶ（MS-ODRAW 2.2.15）
    from todoku.doc.audit import prove as prove_doc
    from todoku.doc.book import DocFile
    from todoku.doc.fill import fill as fill_doc
    from todoku.embed import Picture

    blob = _wide("raw.githubusercontent.com/word_with_embeded.doc")
    doc = DocFile(blob, "w.doc")
    pic = Picture(doc.data, doc.embeds[0]["pictures"][0][1])
    assert pic.slot == 1 and pic.metafile()[40:44] == b" EMF"
    res = fill_doc(DocFile(blob, "w.doc"), {"E1!Sheet1!A2": "x"})
    pr = prove_doc(blob, res.blob, written_of(res.patches))
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    # A1 の字は列の外まではみ出して描かれ、そこの目盛り線は描かれていない：書き直すと線の見え方も変わるので、書かない
    with pytest.raises(FillError, match="はみ出して"):
        fill_doc(DocFile(blob, "w.doc"), {"E1!Sheet1!A1": "x"})


def test_print_picture_without_gridlines_right_aligned_number():
    # 印刷用の絵（EPRINT）には目盛り線がなく、字は右寄せの数字だけ：右寄せの字の右端（境 − 余白）は決まるので、数字は描ける。
    # 左寄せの字の位置は決まらないので、書かない
    blob = _wide("raw.githubusercontent.com/EmbeddedDocument.docx")
    res = fill(DocxDocument(blob, "e.docx"), {"E1!Sheet1!B1": 12})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected
    part = "word/embeddings/Microsoft_Office_Excel_97-2003_Worksheet1.xls"
    ep = Emf(CFB(zipfile.ZipFile(io.BytesIO(res.blob)).read(part)).read("\x03EPRINT"))
    run = next(r for r in ep.runs if r.text == "12")
    assert (run.x, run.y, run.align & 6) == (848, 7, TA_RIGHT)  # 元の「2」の右端（801＋47）、同じ高さ
    with pytest.raises(FillError, match="横の位置が決まりません"):
        fill(DocxDocument(blob, "e.docx"), {"E1!Sheet1!B1": "テスト"})


def test_embedded_workbook_showing_a_chart_sheet():
    # 選ばれているシートがグラフのシート：絵はグラフ。グラフが使うセルの値が変わらない書き込みだけを受け付ける
    from todoku.doc.audit import prove as prove_doc
    from todoku.doc.book import DocFile
    from todoku.doc.fill import fill as fill_doc

    blob = _wide("data.bodik.jp/1811.doc")
    sheet = "行財政マネジメントのための基本方針"
    with pytest.raises(FillError, match="グラフ"):
        fill_doc(DocFile(blob, "c.doc"), {f"E1!{sheet}!R4": 3})
    res = fill_doc(DocFile(blob, "c.doc"), {f"E1!{sheet}!O1": "テスト"})
    pr = prove_doc(blob, res.blob, written_of(res.patches))
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    assert "グラフを表示している埋め込み 1 個" in pr.embed_note


def test_xls_proof_with_a_chart_sheet_before_the_worksheets():
    # グラフのシート（Excel.Chart.8）が先にあるブック：シートの部分とシートを位置で結ぶ（ずれると変更を見落とす）
    from todoku.doc.book import DocFile
    from todoku.doc.fill import fill as fill_doc
    from todoku.embed import open_book
    from todoku.xls.proof import prove as prove_xls

    blob = _wide("data.bodik.jp/1844.doc")
    res = fill_doc(DocFile(blob, "g.doc"), {"E1!入力シート!H1": "テスト"})
    a, b = DocFile(blob, "a"), DocFile(res.blob, "b")
    ba, bb = open_book(a, a.embeds[0]), open_book(b, b.embeds[0])
    pr = prove_xls(ba.raw, bb.raw, {("入力シート", "H1")})
    assert [c["cell"] for c in pr.cell_changes] == ["H1"] and not pr.unexpected
    # 最後の部分の EOF の後ろのバイト（どの部分のレコードでもない）は、組み立て直さずにそのまま残す
    tail = ba.wb.data[1106950:]
    assert len(tail) == 21979 and bb.wb.data.endswith(tail)


def test_embedded_workbook_with_macros_that_do_not_run_on_change():
    # マクロはあるが、セルの値を変えたときに動くもの（Worksheet_Change など）はない：書ける（マクロがあることは知らせる）
    from todoku.embed import macro_block, open_book

    blob = _wide("data.bodik.jp/ikoutyousa.docx")
    doc = DocxDocument(blob, "m.docx")
    wb = open_book(doc, doc.embeds[0]).wb
    assert wb.has_macros and macro_block(wb) == ""


def test_text_boxes_in_a_drawing_canvas_are_all_read():
    # 描画キャンバス（wpc）の中の、いくつものテキストボックス：どれも窓に出す（Fallback の同じ中身は数えない）
    from todoku.docx.audit import check_reading

    doc = DocxDocument(_wide("data.bodik.jp/262064rpareport201903.docx"), "r.docx")
    win = build(doc)
    assert all(c.ok for c in check_reading(doc, win)), [c.items[:3] for c in check_reading(doc, win)]
    assert "RPAの適用により、正確性の向上や入力ミス等が減少した。" in to_markdown(win)


def test_empty_workbook_is_not_a_reading_failure():
    # どのシートにも値のあるセルがないブック（埋め込みだけ）：読み落としではない（XML を直接読んで確かめる）
    from todoku.xlsx.audit import run

    au = run(_wide("raw.githubusercontent.com/bug64512_embed.xlsx"), "bug64512_embed.xlsx")
    c = next(c for c in au.checks if c.name == "読み落とし（全体）")
    assert c.ok and au.ok


def test_word_document_shorter_cbmac_is_valid():
    # cbMac は WordDocument の中の意味のあるバイトの数：ストリームの方が長くてよい（MS-DOC 2.5.4）
    from todoku.doc.audit import lint
    from todoku.doc.book import DocFile

    doc = DocFile(_wide("raw.githubusercontent.com/word_with_embeded.doc"), "w.doc")
    assert doc.lw[0] < len(doc.wd) and lint(doc) == []
