"""読むだけの古い形式・ほかのソフトの形式：手で作ったファイルで、窓・検査（読み落とし）・段階を確かめる。"""

from __future__ import annotations

import io
import struct
import zipfile

import pytest

from todoku import api
from todoku.book import kind_of


def _p(tmp_path, name, blob):
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


def _text(path):
    _k, _doc, parts = api.window(path)
    return "\n".join(t for _n, t in parts)


def _ok_audit(path):
    au = api.audit(path)
    assert au["ok"], [c for c in au["checks"] if not c["ok"]]
    return au


def _zip(files: dict) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in files.items():
            z.writestr(n, b)
    return buf.getvalue()


# ====================================================================== .ppt
def _ppt() -> bytes:
    from test_embeds import make_cfb, ppt_document

    d, cu = ppt_document()
    return make_cfb({"PowerPoint Document": d, "Current User": cu})


def test_ppt_reads_slide_text_and_is_read_only(tmp_path):
    blob = _ppt()
    assert kind_of(blob) == "ppt"
    p = _p(tmp_path, "a.ppt", blob)
    t = _text(p)
    assert "会議の資料" in t and "Agenda 2026" in t
    _ok_audit(p)
    g = api.grade(p)
    assert g["grade"] == "B" and g["write"]["grade"] == "D"
    with pytest.raises(api.ApiError, match="書きません|書けません"):
        api.fill(p, {"x": "y"}, str(tmp_path / "b.ppt"))


def test_ppt_encrypted_is_f_with_next_step(tmp_path):
    from test_embeds import make_cfb, ppt_document

    d, cu = ppt_document()
    cu = bytearray(cu)
    struct.pack_into("<I", cu, 12, 0xF3D1C4DF)
    p = _p(tmp_path, "a.ppt", make_cfb({"PowerPoint Document": d, "Current User": bytes(cu)}))
    g = api.grade(p)
    assert g["grade"] == "F" and "パスワード" in str(g["reasons"])


# ====================================================================== XPS
def _xps(page1: str, oxps: bool = False) -> bytes:
    ns = "http://schemas.openxps.org/oxps/v1.0" if oxps else "http://schemas.microsoft.com/xps/2005/06"
    rel = ("http://schemas.openxps.org/oxps/v1.0/fixedrepresentation" if oxps
           else "http://schemas.microsoft.com/xps/2005/06/fixedrepresentation")
    return _zip({
        "[Content_Types].xml": '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        "_rels/.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="{rel}" Target="/FixedDocSeq.fdseq"/></Relationships>',
        "FixedDocSeq.fdseq": f'<FixedDocumentSequence xmlns="{ns}"><DocumentReference Source="/Documents/1/FixedDoc.fdoc"/></FixedDocumentSequence>',
        "Documents/1/FixedDoc.fdoc": f'<FixedDocument xmlns="{ns}"><PageContent Source="Pages/1.fpage" Width="816" Height="1056"/><PageContent Source="Pages/2.fpage"/></FixedDocument>',
        "Documents/1/Pages/1.fpage/[0].piece": f'<FixedPage xmlns="{ns}" Width="816" Height="1056">{page1[:40]}',
        "Documents/1/Pages/1.fpage/[1].last.piece": f'{page1[40:]}</FixedPage>',
        "Documents/1/Pages/2.fpage": f'<FixedPage xmlns="{ns}" Width="816" Height="1056"><Glyphs FontUri="/f.ttf" Indices="1;2" OriginX="1" OriginY="1" FontRenderingEmSize="10" Fill="#000"/>'
                                     '<Path Data="M0,0"><Path.Fill><ImageBrush ImageSource="/a.jpg"/></Path.Fill></Path></FixedPage>',
    })


PAGE = ('<Glyphs UnicodeString="二行目" OriginX="10" OriginY="200" FontRenderingEmSize="12" FontUri="/f.ttf" Fill="#000"/>'
        '<Glyphs UnicodeString="右の字" OriginX="400" OriginY="100" FontRenderingEmSize="12" FontUri="/f.ttf" Fill="#000"/>'
        '<Glyphs UnicodeString="{}{題}&amp;" OriginX="10" OriginY="100" FontRenderingEmSize="12" FontUri="/f.ttf" Fill="#000"/>'
        '<Canvas RenderTransform="1,0,0,1,0,500"><Glyphs UnicodeString="下の字" OriginX="10" OriginY="10" FontRenderingEmSize="12" FontUri="/f.ttf" Fill="#000"/></Canvas>')


def test_xps_reads_text_in_position_order(tmp_path):
    for oxps in (False, True):
        blob = _xps(PAGE, oxps)
        assert kind_of(blob) == "xps"
        p = _p(tmp_path, "a.xps", blob)
        t = _text(p)
        assert "{題}& ｜ 右の字\n二行目\n下の字" in t
        assert "字の番号（Indices）だけ" in t and "絵 1 個" in t
        _ok_audit(p)
        g = api.grade(p)
        assert g["write"]["grade"] == "D" and g["grade"] == "C"


# ====================================================================== EPUB
def test_epub_reads_spine_order(tmp_path):
    blob = _zip({
        "mimetype": "application/epub+zip",
        "META-INF/container.xml": '<container xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/c.opf" media-type="application/oebps-package+xml"/></rootfiles></container>',
        "OEBPS/c.opf": '<package xmlns="http://www.idpf.org/2007/opf"><manifest><item id="b" href="b.xhtml" media-type="application/xhtml+xml"/><item id="a" href="a%20x.xhtml" media-type="application/xhtml+xml"/></manifest><spine><itemref idref="a"/><itemref idref="b"/></spine></package>',
        "OEBPS/a x.xhtml": '<html xmlns="http://www.w3.org/1999/xhtml"><head><title>題</title></head><body><p>第一章 &amp; 字</p></body></html>',
        "OEBPS/b.xhtml": "<html><body><p>第二章</p></body></html>",
    })
    assert kind_of(blob) == "epub"
    p = _p(tmp_path, "a.epub", blob)
    t = _text(p)
    assert t.index("第一章 & 字") < t.index("第二章") and "題" not in t
    _ok_audit(p)
    g = api.grade(p)
    assert g["grade"] == "B" and g["write"]["grade"] == "D"
