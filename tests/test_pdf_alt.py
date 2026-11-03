"""PDF の画像の説明（代わりの字）：段階とは別の、アクセシビリティの注意として示す。"""

from __future__ import annotations

from todoku.grade import grade_blob
from todoku.pdf.doc import PdfDocument
from todoku.pdf.window import build, to_markdown

IMG = b"q 40 0 0 40 100 700 cm BI /W 2 /H 2 /CS /G /BPC 8 ID \x00\xff\xff\x00 EI Q"


def make(content: bytes, tagged_alt: bytes | None = None, tagged: bool = False) -> bytes:
    """1ページの PDF。tagged なら、MCID 0 を図の要素（Alt は tagged_alt）に結ぶ構造の木を付ける。"""
    objs = [b"<< /Type /Catalog /Pages 2 0 R" + (b" /StructTreeRoot 5 0 R /MarkInfo << /Marked true >>" if tagged else b"") + b" >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << >> /Contents 4 0 R"
            + (b" /StructParents 0" if tagged else b"") + b" >>",
            b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream"]
    if tagged:
        objs.append(b"<< /Type /StructTreeRoot /K [6 0 R] /ParentTree << /Nums [0 [6 0 R]] >> /RoleMap << /Zu /Figure >> >>")
        objs.append(b"<< /Type /StructElem /S /Zu /P 5 0 R /Pg 3 0 R /K 0" + (b" /Alt " + tagged_alt if tagged_alt else b"") + b" >>")
    out = bytearray(b"%PDF-1.7\n")
    offs = []
    for i, body in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    x = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode() + b"".join(f"{o:010d} 00000 n \n".encode() for o in offs)
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{x}\n%%EOF\n".encode()
    return bytes(out)


def pics(blob: bytes) -> list[dict]:
    return build(PdfDocument(blob, "t.pdf"))["pages"][0]["pictures"]


def test_untagged_image_has_no_description():
    blob = make(IMG)
    p = pics(blob)
    assert len(p) == 1 and p[0]["alt"] is None and not p[0]["decorative"] and p[0]["where"] == "上の左"
    g = grade_blob(blob, "t.pdf")
    assert [a.title for a in g.alerts] == ["説明のない画像"]
    assert g.grade != "F"  # 段階は変えない（エージェントが読めるかとは別の注意）
    assert "説明（代わりの字）がありません" in to_markdown(build(PdfDocument(blob, "t.pdf")))
    assert "アクセシビリティの注意" in g.headline() and g.to_dict()["alerts"]


def test_alt_in_marked_content():
    p = pics(make(b"/Figure <</Alt (Map of the office)>> BDC " + IMG + b" EMC"))
    assert p[0]["alt"] == "Map of the office"


def test_artifact_is_decoration():
    blob = make(b"/Artifact BMC " + IMG + b" EMC")
    assert pics(blob)[0]["decorative"]
    assert grade_blob(blob, "t.pdf").alerts == []


def test_tagged_figure_alt_through_parent_tree_and_rolemap():
    blob = make(b"/Zu <</MCID 0>> BDC " + IMG + b" EMC", b"<FEFF5730 56F3>", tagged=True)
    p = pics(blob)
    assert p[0]["alt"] == "地図" and p[0]["tagged"]
    assert grade_blob(blob, "t.pdf").alerts == []


def test_tagged_figure_without_alt():
    p = pics(make(b"/Zu <</MCID 0>> BDC " + IMG + b" EMC", None, tagged=True))
    assert p[0]["tagged"] and p[0]["alt"] is None


def test_empty_alt_is_no_description():
    assert pics(make(b"/Zu <</MCID 0>> BDC " + IMG + b" EMC", b"( )", tagged=True))[0]["alt"] is None


def test_word_picture_without_description_is_an_alert():
    from test_docx import SAMPLE_BODY, make_docx

    pic = ('<w:p><w:r><w:t>地図</w:t></w:r><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:docPr id="1" name="図 1"{alt}/>'
           '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId99"/></pic:blipFill></pic:pic>'
           '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>')
    g = grade_blob(make_docx(SAMPLE_BODY + pic.format(alt="")), "p.docx")
    assert [a.title for a in g.alerts] == ["説明のない画像"] and "「地図」の段落" in g.alerts[0].where[0]
    ok = grade_blob(make_docx(SAMPLE_BODY + pic.format(alt=' descr="事業所の地図"')), "p.docx")
    # 説明のない本文の画像：中身を字として読めないので、画像の所はエージェントが見て確かめる（読む B）。説明があれば下げない
    assert g.grade == "B" and [r.title for r in g.reasons if r.grade == "B"] == ["説明のない本文の画像"]
    assert ok.alerts == [] and ok.grade == "A"


def test_old_word_picture_without_description_is_an_alert():
    import glob

    import pytest

    files = glob.glob("corpus/wide_embed/www.env.go.jp/aggregation.doc")
    if not files:
        pytest.skip("コーパスがありません")
    g = grade_blob(open(files[0], "rb").read(), "aggregation.doc")
    assert [a.title for a in g.alerts] == ["説明のない画像"] and g.alerts[0].where[0].startswith("図1")


def test_rtf_picture_description():
    from todoku.rtf.document import RtfDocument

    def rtf(prop: bytes) -> bytes:
        return (b"{\\rtf1\\ansi\\ansicpg932{\\fonttbl{\\f0 MS Mincho;}}\\pard \\f0 map\\par "
                b"{\\*\\shppict{\\pict{\\*\\picprop" + prop + b"}\\pngblip 89504e47}}"
                b"{\\nonshppict{\\pict\\wmetafile8 0100}}\\par}")

    bare = RtfDocument(rtf(b""), "a.rtf")
    assert [(p["n"], p["alt"]) for p in bare.pictures] == [(1, "")]  # 控え（\nonshppict）は数えない
    assert [a.title for a in grade_blob(rtf(b""), "a.rtf").alerts] == ["説明のない画像"]
    named = RtfDocument(rtf(b"{\\sp{\\sn wzDescription}{\\sv \\'92\\'6e\\'90\\'7d}}"), "a.rtf")
    assert named.pictures[0]["alt"] == "地図"
