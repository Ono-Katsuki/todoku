"""書類の中の画像：窓の一覧（id・位置・大きさ・説明・部品）と view_image。EMF・WMF を描く。PDF の注釈の画像。"""

from __future__ import annotations

import base64
import struct
import zipfile
import zlib

import pytest

from todoku import api
from todoku.imaging import docimages as D
from todoku.imaging import metaraster as M
from todoku.pdf.doc import PdfDocument


def _png(w=4, h=3, color=(200, 10, 20)):
    def ch(t, b):
        return struct.pack(">I", len(b)) + t + b + struct.pack(">I", zlib.crc32(t + b) & 0xFFFFFFFF)

    raw = b"".join(b"\0" + bytes(color) * w for _ in range(h))
    return b"\x89PNG\r\n\x1a\n" + ch(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + ch(b"IDAT", zlib.compress(raw)) + ch(b"IEND", b"")


def _gif():
    # 1×1 の赤
    return bytes.fromhex("474946383961010001008000 00ff0000000000 2c00000000010001000002024401003b".replace(" ", ""))


def _emf_rect(fill=(0, 0, 255)):
    """100×100 の絵に、塗った四角と字。"""
    recs = []

    def rec(t, body=b""):
        recs.append(struct.pack("<II", t, 8 + len(body)) + body)

    color = fill[0] | fill[1] << 8 | fill[2] << 16
    rec(39, struct.pack("<IIII", 1, 0, color, 0))  # CREATEBRUSHINDIRECT
    rec(37, struct.pack("<I", 1))  # SELECTOBJECT
    rec(43, struct.pack("<4i", 10, 10, 60, 60))  # RECTANGLE
    # EXTTEXTOUTW
    s = "AB"
    bytes_s = s.encode("utf-16-le")
    off = 76
    body = struct.pack("<4i", 0, 0, 0, 0) + struct.pack("<Iff", 1, 1.0, 1.0) + struct.pack("<iiIII", 10, 80, len(s), off, 0) + struct.pack("<4i", 0, 0, 0, 0) + struct.pack("<I", 0) + bytes_s
    rec(84, body)
    rec(14, struct.pack("<III", 0, 0, 20))  # EOF
    data = b"".join(recs)
    # 見出し（88 バイト）：Bounds 0,0,99,99、Frame 0.01 mm で 26.46 mm（96 dpi の 100 画素）
    head_body = struct.pack("<8i", 0, 0, 99, 99, 0, 0, 2646, 2646) + b" EMF" + struct.pack("<IIIHH", 0x10000, 88 + len(data), len(recs) + 1, 2, 0)
    head_body += struct.pack("<III", 0, 0, 0) + struct.pack("<4i", 96, 96, 25, 25)
    return struct.pack("<II", 1, 8 + len(head_body)) + head_body + data


def _wmf_rect():
    """置ける見出しつきの WMF：100×100 単位（1 インチ 100 単位）に、塗った四角。"""
    head = struct.pack("<IH4hHIH", 0x9AC6CDD7, 0, 0, 0, 100, 100, 100, 0, 0)
    recs = []

    def rec(f, body=b""):
        recs.append(struct.pack("<IH", (6 + len(body)) // 2, f) + body)

    rec(0x020B, struct.pack("<2h", 0, 0))
    rec(0x020C, struct.pack("<2h", 100, 100))
    rec(0x02FC, struct.pack("<HIh", 0, 0x0000FF00, 0))  # 緑のブラシ
    rec(0x012D, struct.pack("<H", 0))
    rec(0x041B, struct.pack("<4h", 60, 60, 10, 10))  # 下・右・上・左
    txt = b"Hi"
    rec(0x0521, struct.pack("<H", len(txt)) + txt + struct.pack("<2h", 80, 10))
    recs.append(struct.pack("<IH", 3, 0))
    data = b"".join(recs)
    mh = struct.pack("<HHHIHIH", 1, 9, 0x300, (18 + len(data)) // 2, 1, 16, 0)
    return head + mh + data


def _at(fr, x, y):
    i = (y * fr.w + x) * 3
    return tuple(fr.rgb[i:i + 3])


# ====================================================================== メタファイル
def test_emf_rectangle_and_text_are_drawn_and_text_is_returned():
    fr = M.emf(_emf_rect())
    assert fr.w > 100 and not fr.problems, fr.problems
    k = fr.w / 100
    assert _at(fr, int(30 * k), int(30 * k)) == (0, 0, 255)
    assert _at(fr, int(80 * k), int(30 * k)) == (255, 255, 255)
    assert any("AB" in n for n in fr.notes)


def test_wmf_rectangle_is_drawn():
    fr = M.wmf(_wmf_rect())
    k = fr.w / 100
    assert _at(fr, int(30 * k), int(30 * k)) == (0, 255, 0)
    assert any("Hi" in n for n in fr.notes)


def test_emf_stretchdibits_draws_the_bitmap():
    # 2×2 の 24 ビットの DIB を、全面に
    dib = struct.pack("<IiiHHIIiiII", 40, 2, 2, 1, 24, 0, 16, 0, 0, 0, 0)
    bits = bytes([0, 0, 255, 0, 255, 0]) + b"\0\0" + bytes([255, 0, 0, 255, 255, 255]) + b"\0\0"  # 下の行：赤・緑、上の行：青・白（BGR）
    off_bmi = 80
    body = struct.pack("<4i", 0, 0, 99, 99) + struct.pack("<6i", 0, 0, 0, 0, 2, 2) + struct.pack("<4I", off_bmi, len(dib), off_bmi + len(dib), len(bits))
    body += struct.pack("<II", 0, 0x00CC0020) + struct.pack("<2i", 100, 100) + dib + bits
    r = struct.pack("<II", 81, 8 + len(body)) + body
    eof = struct.pack("<II", 14, 20) + struct.pack("<III", 0, 0, 20)
    hb = struct.pack("<8i", 0, 0, 99, 99, 0, 0, 2646, 2646) + b" EMF" + struct.pack("<IIIHH", 0x10000, 0, 2, 0, 0) + struct.pack("<III", 0, 0, 0) + struct.pack("<4i", 96, 96, 25, 25)
    fr = M.emf(struct.pack("<II", 1, 8 + len(hb)) + hb + r + eof)
    assert _at(fr, 5, 5) == (0, 0, 255)  # 上の行の左：青
    assert _at(fr, fr.w - 5, fr.h - 5) == (0, 255, 0)  # 下の行の右：緑


def test_emf_unknown_record_is_reported():
    hb = struct.pack("<8i", 0, 0, 9, 9, 0, 0, 265, 265) + b" EMF" + struct.pack("<IIIHH", 0x10000, 0, 2, 0, 0) + struct.pack("<III", 0, 0, 0) + struct.pack("<4i", 96, 96, 25, 25)
    r = struct.pack("<II", 79, 8)  # PLGBLT
    fr = M.emf(struct.pack("<II", 1, 8 + len(hb)) + hb + r + struct.pack("<II", 14, 8))
    assert fr.problems and "79" in fr.problems[0]


def test_emf_plus_only_says_it_cannot_draw_and_extracts_embedded_png():
    hb = struct.pack("<8i", 0, 0, 9, 9, 0, 0, 265, 265) + b" EMF" + struct.pack("<IIIHH", 0x10000, 0, 2, 0, 0) + struct.pack("<III", 0, 0, 0) + struct.pack("<4i", 96, 96, 25, 25)
    com = struct.pack("<I", 8) + b"EMF+" + b"\x00" * 4
    com = com + b"\0" * (-len(com) % 4)
    r = struct.pack("<II", 70, 8 + len(com)) + com
    head = struct.pack("<II", 1, 8 + len(hb)) + hb
    fr = M.emf(head + r + struct.pack("<II", 14, 8))
    assert any("EMF+" in p for p in fr.problems)
    png = _png(3, 2)
    pad = b"\0" * (-len(png) % 4)
    com = struct.pack("<I", 4 + len(png) + len(pad)) + b"EMF+" + png + pad
    r = struct.pack("<II", 70, 8 + len(com)) + com
    fr = M.emf(head + r + struct.pack("<II", 14, 8))
    assert (fr.w, fr.h) == (3, 2) and fr.notes


# ====================================================================== docx
def _docx(tmp_path):
    p = tmp_path / "a.docx"
    doc = """<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>
<w:p><w:r><w:t>地図：</w:t></w:r><w:r><w:drawing><wp:inline><wp:extent cx="1800000" cy="900000"/><wp:docPr id="1" name="図" descr="案内図"/>
<a:graphic><a:graphicData uri="p"><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:r><w:drawing><wp:inline><wp:extent cx="360000" cy="360000"/><wp:docPr id="2" name="g"/>
<a:graphic><a:graphicData uri="p"><pic:pic><pic:blipFill><a:blip r:embed="rId2"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:p><w:r><w:drawing><wp:inline><wp:extent cx="360000" cy="360000"/><wp:docPr id="3" name="e"/>
<a:graphic><a:graphicData uri="p"><pic:pic><pic:blipFill><a:blip r:embed="rId3"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
</w:body></w:document>"""
    rels = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.gif"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image3.emf"/>
</Relationships>"""
    ct = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="emf" ContentType="image/x-emf"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""
    prels = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", prels)
        z.writestr("word/document.xml", doc)
        z.writestr("word/_rels/document.xml.rels", rels)
        z.writestr("word/media/image1.png", _png())
        z.writestr("word/media/image2.gif", _gif())
        z.writestr("word/media/image3.emf", _emf_rect())
        z.writestr("word/media/header9.png", _png(2, 2))  # 本文から指されていない
    return p


def test_docx_window_lists_every_image_with_id_size_alt_part(tmp_path):
    p = _docx(tmp_path)
    _k, doc, parts = api.window(str(p))
    imgs = {i["id"]: i for i in doc["images"]}
    assert imgs["図1"]["part"] == "word/media/image1.png" and imgs["図1"]["size_mm"] == [50.0, 25.0] and imgs["図1"]["alt"] == "案内図"
    assert imgs["図2"]["format"] == "gif" and imgs["図3"]["format"] == "emf"
    assert "word/media/header9.png" in imgs
    text = "\n".join(t for _n, t in parts)
    assert "### 書類の中の画像（4 個" in text and "案内図" in text and "view_image" in text


def test_view_image_png_is_returned_as_is(tmp_path):
    p = _docx(tmp_path)
    v = api.view_image(str(p), "図1")
    assert v["mime"] == "image/png" and base64.b64decode(v["data_base64"]) == _png() and (v["width"], v["height"]) == (4, 3)
    assert api.view_image(str(p), "word/media/image1.png")["id"] == "word/media/image1.png"


def test_view_image_gif_and_emf_become_png(tmp_path):
    p = _docx(tmp_path)
    v = api.view_image(str(p), "図2")
    assert v["mime"] == "image/png" and base64.b64decode(v["data_base64"])[:4] == b"\x89PNG"
    v = api.view_image(str(p), "図3")
    assert v["mime"] == "image/png" and "AB" in v["text"]


def test_view_image_unknown_id_lists_the_ids(tmp_path):
    p = _docx(tmp_path)
    with pytest.raises(api.ApiError) as e:
        api.view_image(str(p), "図99")
    assert "図1" in str(e.value)


# ====================================================================== xlsx
def _xlsx(tmp_path):
    p = tmp_path / "a.xlsx"
    ct = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>"""
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>""")
        z.writestr("xl/workbook.xml", """<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="様式" sheetId="1" r:id="rId1"/></sheets></workbook>""")
        z.writestr("xl/_rels/workbook.xml.rels", """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>""")
        z.writestr("xl/worksheets/sheet1.xml", """<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>氏名</t></is></c></row></sheetData><drawing r:id="rId1"/></worksheet>""")
        z.writestr("xl/worksheets/_rels/sheet1.xml.rels", """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>""")
        z.writestr("xl/drawings/drawing1.xml", """<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="印" descr="角印"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="720000" cy="720000"/></a:xfrm></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>""")
        z.writestr("xl/drawings/_rels/drawing1.xml.rels", """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>""")
        z.writestr("xl/media/image1.png", _png(5, 5))
    return p


def test_xlsx_window_and_view_image(tmp_path):
    p = _xlsx(tmp_path)
    _k, doc, parts = api.window(str(p))
    im = doc["images"][0]
    assert im["id"] == "様式!印" and im["part"] == "xl/media/image1.png" and im["alt"] == "角印" and im["size_mm"] == [20.0, 20.0]
    assert "B3" in im["where"]
    assert "様式!印" in parts[-1][1]
    v = api.view_image(str(p), "様式!印")
    assert base64.b64decode(v["data_base64"]) == _png(5, 5)


# ====================================================================== ODF・HWPX
def test_odt_list_and_view(tmp_path):
    p = tmp_path / "a.odt"
    content = """<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><office:body><office:text>
<text:p>本文<draw:frame draw:name="絵1" svg:width="5cm" svg:height="2cm"><draw:image xlink:href="Pictures/a.png"/><svg:desc>説明の字</svg:desc></draw:frame></text:p></office:text></office:body></office:document-content>"""
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.text")
        z.writestr("content.xml", content)
        z.writestr("META-INF/manifest.xml", """<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>""")
        z.writestr("Pictures/a.png", _png())
    _k, doc, _parts = api.window(str(p))
    im = doc["images"][0]
    assert im["id"] == "図1" and im["part"] == "Pictures/a.png" and im["size_mm"] == [50.0, 20.0] and im["alt"] == "説明の字"
    assert base64.b64decode(api.view_image(str(p), "図1")["data_base64"]) == _png()


def test_hwpx_list_from_raw_parts():
    import io

    b = io.BytesIO()
    with zipfile.ZipFile(b, "w") as z:
        z.writestr("mimetype", "application/hwp+zip")
        z.writestr("Contents/content.hpf", """<?xml version="1.0"?><opf:package xmlns:opf="http://www.idpf.org/2007/opf/"><opf:manifest><opf:item id="image1" href="BinData/image1.png" media-type="image/png"/></opf:manifest></opf:package>""")
        z.writestr("Contents/section0.xml", """<?xml version="1.0"?><hs:sec xmlns:hs="urn:hs" xmlns:hp="urn:hp" xmlns:hc="urn:hc"><hp:p><hp:run><hp:pic><hp:sz width="7200" height="3600"/><hc:img binaryItemIDRef="image1"/><hp:shapeComment>그림</hp:shapeComment></hp:pic></hp:run></hp:p></hs:sec>""")
        z.writestr("BinData/image1.png", _png())
    blob = b.getvalue()
    imgs = D.list_images("hwpx", blob, {})
    assert imgs[0]["id"] == "図1" and imgs[0]["part"] == "BinData/image1.png" and imgs[0]["size_mm"] == [25.4, 12.7] and imgs[0]["alt"] == "그림"
    assert base64.b64decode(D.get_image("hwpx", blob, imgs, "図1")["data_base64"]) == _png()


# ====================================================================== 旧形式・RTF
def test_officeart_blips_are_parsed_png_jpeg_and_emf():
    png = _png()
    rec_png = struct.pack("<HHI", 0x6E0 << 4, 0xF01E, 17 + len(png)) + b"\x11" * 16 + b"\xff" + png
    emf = _emf_rect()
    comp = zlib.compress(emf)
    meta = struct.pack("<I4i2iIBB", len(emf), 0, 0, 100, 100, 100, 100, len(comp), 0, 0xFE)
    rec_emf = struct.pack("<HHI", 0x3D4 << 4, 0xF01A, 16 + len(meta) + len(comp)) + b"\x22" * 16 + meta + comp
    inner = rec_png + rec_emf
    cont = struct.pack("<HHI", 0x000F, 0xF000, len(inner)) + inner
    got = D._blips(cont)
    assert got[0] == ("png", png) and got[1][0] == "emf" and got[1][1] == emf


def test_biff_drawing_group_is_joined_over_continue():
    a, b = b"x" * 10, b"y" * 5
    wb = struct.pack("<HH", 0x00EB, len(a)) + a + struct.pack("<HH", 0x003C, len(b)) + b + struct.pack("<HH", 0x000A, 0)
    assert D._biff_drawing_groups(wb) == a + b


def test_rtf_picture_is_listed_and_returned(tmp_path):
    p = tmp_path / "a.rtf"
    png = _png()
    p.write_bytes(b"{\\rtf1\\ansi body{\\pict\\pngblip\\picw4\\pich3\\picwgoal1440\\pichgoal720 " + png.hex().encode() + b"}}")
    blob = p.read_bytes()
    imgs = D.list_images("rtf", blob, {})
    assert imgs[0]["format"] == "png" and imgs[0]["size_mm"] == [25.4, 12.7]
    assert base64.b64decode(D.get_image("rtf", blob, imgs, "blip1")["data_base64"]) == png


# ====================================================================== PDF
def _pdf_with_image_and_stamp():
    raw = zlib.compress(bytes([10, 200, 30]) * 4)
    img = b"<</Type/XObject/Subtype/Image/Width 2/Height 2/BitsPerComponent 8/ColorSpace/DeviceRGB/Filter/FlateDecode/Length %d>>stream\n" % len(raw) + raw + b"\nendstream"
    ap_c = b"q 50 0 0 50 0 0 cm /I Do Q"
    ap = b"<</Type/XObject/Subtype/Form/BBox[0 0 50 50]/Resources<</XObject<</I 5 0 R>>>>/Length %d>>stream\n" % len(ap_c) + ap_c + b"\nendstream"
    cont = b"q 100 0 0 100 50 50 cm /I Do Q"
    objs = [b"<</Type/Catalog/Pages 2 0 R>>", b"<</Type/Pages/Kids[3 0 R]/Count 1>>",
            b"<</Type/Page/Parent 2 0 R/MediaBox[0 0 400 400]/Contents 4 0 R/Resources<</XObject<</I 5 0 R>>>>/Annots[7 0 R]>>",
            b"<</Length %d>>stream\n" % len(cont) + cont + b"\nendstream", img, ap,
            b"<</Type/Annot/Subtype/Stamp/Rect[200 200 250 250]/AP<</N 6 0 R>>/NM(hanko)>>"]
    out = bytearray(b"%PDF-1.5\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for o in offs:
        out += b"%010d 00000 n \n" % o
    out += b"trailer<</Size %d/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, x)
    return bytes(out)


def test_pdf_page_image_and_annotation_image_can_be_viewed(tmp_path):
    p = tmp_path / "a.pdf"
    p.write_bytes(_pdf_with_image_and_stamp())
    doc = PdfDocument(p.read_bytes(), "a.pdf")
    items = D.pdf_images(doc)
    assert [i["id"] for i in items] == ["p1.i1", "p1.a1"]
    v = api.view_image(str(p), "p1.a1")
    assert v["mime"] == "image/png" and (v["width"], v["height"]) == (2, 2) and "注釈" in v["text"]
    with pytest.raises(api.ApiError):
        api.view_image(str(p), "p9.i9")


def test_mcp_tools_exist_and_return_an_image_block(tmp_path):
    from todoku import mcp_server as M2

    names = {t["name"] for t in M2.TOOLS}
    assert {"view_image", "to_pdf"} <= names
    p = _docx(tmp_path)
    r = M2.tool_result("view_image", {"path": str(p), "image": "図1"})
    assert not r.get("isError") and r["content"][1]["type"] == "image" and r["content"][1]["mimeType"] == "image/png"


def test_bstore_entry_holds_the_blip_in_excel():
    png = _png()
    blip = struct.pack("<HHI", 0x6E0 << 4, 0xF01E, 17 + len(png)) + b"\x11" * 16 + b"\xff" + png
    fbse = bytes(36) + blip
    bse = struct.pack("<HHI", 0x0002 | (6 << 4), 0xF007, len(fbse)) + fbse
    store = struct.pack("<HHI", 0x000F | (1 << 4), 0xF001, len(bse)) + bse
    dgg = struct.pack("<HHI", 0, 0xF006, 8) + bytes(8)
    inner = dgg + store
    cont = struct.pack("<HHI", 0x000F, 0xF000, len(inner)) + inner
    assert D._blips(cont) == [("png", png)]


def test_rtf_picture_with_nested_groups_and_blip_uid():
    png = _png()
    rtf = (b"{\\rtf1{\\pict{\\*\\picprop{\\sp{\\sn x}{\\sv 75}}}\\pngblip\\bliptag-56558{\\*\\blipuid de49cf}" + png.hex().encode() + b"}}")
    im = D._rtf_pictures(rtf)[0]
    assert im["format"] == "png" and im["data"] == png


# ====================================================================== 字の形が分からない字
def _font_pdf(font_dict: bytes, text: bytes):
    cont = b"BT /F 40 Tf 50 300 Td " + text + b" Tj ET"
    objs = [b"<</Type/Catalog/Pages 2 0 R>>", b"<</Type/Pages/Kids[3 0 R]/Count 1>>",
            b"<</Type/Page/Parent 2 0 R/MediaBox[0 0 400 400]/Contents 4 0 R/Resources<</Font<</F 5 0 R>>>>>>",
            b"<</Length %d>>stream\n" % len(cont) + cont + b"\nendstream", font_dict]
    return objs


def _build(objs):
    out = bytearray(b"%PDF-1.5\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for o in offs:
        out += b"%010d 00000 n \n" % o
    out += b"trailer<</Size %d/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, x)
    return bytes(out)


CMAP = (b"/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CMapName /X def /CMapType 2 def 1 begincodespacerange <00> <FF> endcodespacerange "
        b"1 beginbfchar <41> <3042> endbfchar endcmap end end")


@pytest.mark.skipif(not __import__("os").path.exists("/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc"), reason="和文の書体がない")
def test_tounicode_glyph_is_drawn_with_any_available_font_not_a_box():
    from todoku.pdf.pageimage import render, view_page

    # 埋め込みのない未知の書体：字は ToUnicode で「あ」と分かる
    font = b"<</Type/Font/Subtype/TrueType/BaseFont/NoSuchFont-XYZ/FirstChar 65/LastChar 65/Widths[1000]/ToUnicode 6 0 R>>"
    objs = _font_pdf(font, b"(A)") + [b"<</Length %d>>stream\n" % len(CMAP) + CMAP + b"\nendstream"]
    pdf = _build(objs)
    doc = PdfDocument(pdf, "t.pdf")
    r = render(doc, 1)
    assert r["stats"]["glyphs_box"] == 0
    v = view_page(doc, 1)
    assert "字の形が分からない" not in v["text"]


def test_unknown_glyph_positions_are_reported_and_read_is_e(monkeypatch):
    from todoku.grade import grade_blob
    from todoku.pdf import pageimage
    from todoku.pdf.render import fontmap

    # どの書体にも字がないとき（このパソコンに書体がない環境）：灰色の四角と、位置と、読めないこと
    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    monkeypatch.setattr(fontmap, "_glyphs_cache", {}, raising=False)
    font = b"<</Type/Font/Subtype/Type0/BaseFont/NoSuchCid/Encoding/Identity-H/DescendantFonts[6 0 R]>>"
    cid = b"<</Type/Font/Subtype/CIDFontType2/BaseFont/NoSuchCid/CIDSystemInfo<</Registry(Adobe)/Ordering(Identity)/Supplement 0>>/DW 1000>>"
    objs = _font_pdf(font, b"<00410042>") + [cid]
    pdf = _build(objs)
    doc = PdfDocument(pdf, "t.pdf")
    v = pageimage.view_page(doc, 1)
    drawn = "／".join(v["not_drawn"])
    assert "字の形が分からない字" in drawn and "位置" in drawn
    assert "灰色の四角" in v["text"] and "推測で埋めず" in v["text"]
    g = grade_blob(pdf, "t.pdf")
    assert g.grade == "E", g.headline()


def test_pptx_window_and_view_image(tmp_path):
    p = tmp_path / "a.pptx"
    ns = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    slide = f"""<?xml version="1.0"?><p:sld {ns}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:pic><p:nvPicPr><p:cNvPr id="2" name="絵" descr="地図"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3600000" cy="1800000"/></a:xfrm></p:spPr></p:pic>
</p:spTree></p:cSld></p:sld>"""
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>""")
        z.writestr("_rels/.rels", """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>""")
        z.writestr("ppt/presentation.xml", f"""<?xml version="1.0"?><p:presentation {ns}><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst></p:presentation>""")
        z.writestr("ppt/_rels/presentation.xml.rels", """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>""")
        z.writestr("ppt/slides/slide1.xml", slide)
        z.writestr("ppt/slides/_rels/slide1.xml.rels", """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>""")
        z.writestr("ppt/media/image1.png", _png())
    _k, doc, _parts = api.window(str(p))
    im = doc["images"][0]
    assert im["id"] == "図1" and im["part"] == "ppt/media/image1.png" and im["size_mm"] == [100.0, 50.0] and im["alt"] == "地図"
    assert base64.b64decode(api.view_image(str(p), "図1")["data_base64"]) == _png()
