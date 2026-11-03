"""試しの PowerPoint（.pptx）を作る：ECMA-376 の最小限の部品（presentation・スライド・レイアウト・マスター・ノート）。"""

from __future__ import annotations

import io
import zipfile
from xml.sax.saxutils import escape

NS = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def sp(paras: str, name: str = "テキスト", ph: str = "", sid: int = 2) -> str:
    phx = f'<p:ph type="{ph}"/>' if ph else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/><p:cNvSpPr/><p:nvPr>{phx}</p:nvPr></p:nvSpPr><p:spPr/>'
            f'<p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody></p:sp>')


def para(*runs: str, end: str = '<a:endParaRPr lang="ja-JP" sz="1800"/>') -> str:
    return "<a:p>" + "".join(f'<a:r><a:rPr lang="ja-JP"/><a:t>{escape(t)}</a:t></a:r>' for t in runs) + end + "</a:p>"


def table(rows: list[list[str]], sid: int = 5) -> str:
    trs = ""
    for r in rows:
        tcs = ""
        for c in r:
            if c == "<hmerge>":
                tcs += '<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p><a:endParaRPr lang="ja-JP"/></a:p></a:txBody><a:tcPr/></a:tc>'
                continue
            body = para(c) if c else '<a:p><a:endParaRPr lang="ja-JP" sz="1400"/></a:p>'
            tcs += f"<a:tc><a:txBody><a:bodyPr/><a:lstStyle/>{body}</a:txBody><a:tcPr/></a:tc>"
        trs += f'<a:tr h="370840">{tcs}</a:tr>'
    grid = "".join('<a:gridCol w="1000000"/>' for _ in rows[0])
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name="表"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>'
            '<p:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></p:xfrm><a:graphic>'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f"<a:tbl><a:tblPr/><a:tblGrid>{grid}</a:tblGrid>{trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>")


def pic(descr: str | None, sid: int = 9, decorative: bool = False) -> str:
    d = f' descr="{descr}"' if descr is not None else ""
    deco = ('<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"><adec:decorative '
            'xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst>') if decorative else ""
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{sid}" name="図 {sid}"{d}>{deco}</p:cNvPr><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
            '<p:blipFill><a:blip r:embed="rIdImg"/></p:blipFill><p:spPr/></p:pic>')


def group(*shapes: str) -> str:
    return ('<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="グループ"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
            + "".join(shapes) + "</p:grpSp>")


def _tree(shapes: str) -> str:
    return ('<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>'
            + shapes + "</p:spTree></p:cSld>")


def make(slides: list[str], notes: dict[int, str] | None = None, master_text: str = "", hidden: set[int] | None = None) -> bytes:
    """slides：スライドごとの図形の XML。notes：スライドの番号（1 から）→ ノートの図形。master_text：マスターの、プレースホルダーでない図形の字。"""
    notes = notes or {}
    hidden = hidden or set()
    files: dict[str, str] = {}
    ct = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="xml" ContentType="application/xml"/>',
          '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>']
    files["_rels/.rels"] = (f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                            f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>')
    ids = "".join(f'<p:sldId id="{256 + i}" r:id="rId{10 + i}"/>' for i in range(len(slides)))
    files["ppt/presentation.xml"] = (f'<?xml version="1.0" encoding="UTF-8"?><p:presentation {NS}>'
                                     '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
                                     f'<p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/></p:presentation>')
    prels = [f'<Relationship Id="rId1" Type="{REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    for i, shapes in enumerate(slides, 1):
        prels.append(f'<Relationship Id="rId{9 + i}" Type="{REL}/slide" Target="slides/slide{i}.xml"/>')
        show = ' show="0"' if i in hidden else ""
        files[f"ppt/slides/slide{i}.xml"] = f'<?xml version="1.0" encoding="UTF-8"?><p:sld {NS}{show}>{_tree(shapes)}</p:sld>'
        srels = [f'<Relationship Id="rId1" Type="{REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>',
                 f'<Relationship Id="rIdImg" Type="{REL}/image" Target="../media/image1.png"/>']
        if i in notes:
            srels.append(f'<Relationship Id="rId2" Type="{REL}/notesSlide" Target="../notesSlides/notesSlide{i}.xml"/>')
            files[f"ppt/notesSlides/notesSlide{i}.xml"] = (f'<?xml version="1.0" encoding="UTF-8"?><p:notes {NS}>{_tree(notes[i])}</p:notes>')
            ct.append(f'<Override PartName="/ppt/notesSlides/notesSlide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>')
        files[f"ppt/slides/_rels/slide{i}.xml.rels"] = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                                        + "".join(srels) + "</Relationships>")
        ct.append(f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
    files["ppt/_rels/presentation.xml.rels"] = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                                + "".join(prels) + "</Relationships>")
    files["ppt/slideLayouts/slideLayout1.xml"] = (f'<?xml version="1.0" encoding="UTF-8"?><p:sldLayout {NS}>'
                                                  + _tree(sp(para("タイトルを入力"), "タイトル", "title")) + "</p:sldLayout>")
    files["ppt/slideLayouts/_rels/slideLayout1.xml.rels"] = ('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                                             f'<Relationship Id="rId1" Type="{REL}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
    mshapes = sp(para("マスターのタイトルの書式"), "タイトル", "title") + (sp(para(master_text), "ロゴ", "", 7) if master_text else "")
    files["ppt/slideMasters/slideMaster1.xml"] = f'<?xml version="1.0" encoding="UTF-8"?><p:sldMaster {NS}>{_tree(mshapes)}</p:sldMaster>'
    ct.append("</Types>")
    files["[Content_Types].xml"] = "".join(ct)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", files.pop("[Content_Types].xml"))
        for n, x in files.items():
            z.writestr(n, x)
        z.writestr("ppt/media/image1.png", b"\x89PNG\r\n\x1a\n")
    return buf.getvalue()
