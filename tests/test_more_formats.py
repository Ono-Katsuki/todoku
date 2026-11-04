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


# ====================================================================== Excel 2〜4
def _r(t, body):
    return struct.pack("<HH", t, len(body)) + body


def _biff2() -> bytes:
    attr = bytes([0x40, 0x00, 0x40])  # ロック・書式 0・下の罫線
    g = bytes([0x24]) + struct.pack("<HB", 0xC000 | 1, 0) + bytes([0x1E]) + struct.pack("<H", 2) + b"\x05"  # A2*2
    fsum = bytes([0x25]) + struct.pack("<HHBB", 0xC000 | 1, 0xC000 | 1, 0, 0) + bytes([0x22, 1, 4])  # SUM(A2:A2)
    return b"".join([
        _r(0x0009, struct.pack("<HH", 2, 0x10)),
        _r(0x0042, struct.pack("<H", 932)),
        _r(0x001E, b"\x07General"),
        _r(0x0004, struct.pack("<HH", 0, 0) + attr + bytes([6]) + "申請".encode("cp932")),
        _r(0x0003, struct.pack("<HH", 1, 0) + attr + struct.pack("<d", 21.5)),
        _r(0x0002, struct.pack("<HH", 1, 1) + attr + struct.pack("<H", 7)),
        _r(0x0006, struct.pack("<HH", 2, 0) + attr + struct.pack("<d", 43.0) + b"\x00" + bytes([len(g)]) + g),
        _r(0x0006, struct.pack("<HH", 2, 1) + attr + struct.pack("<d", 21.5) + b"\x00" + bytes([len(fsum)]) + fsum),
        _r(0x000A, b""),
    ]).replace(bytes([6]) + "申請".encode("cp932"), bytes([4]) + "申請".encode("cp932"))


def _biff34(biff: int) -> bytes:
    bof = {3: 0x0209, 4: 0x0409}[biff]
    xf_t = {3: 0x0243, 4: 0x0443}[biff]
    fmt = (b"" if biff == 3 else b"\x00\x00") + b"\x07General"
    xf = bytes([0, 0, 0, 0]) + bytes(4) + struct.pack("<I", 1 << 16)  # ロックしない・下の罫線
    fnid = bytes([0x21]) + (bytes([24]) if biff == 3 else struct.pack("<H", 24))  # ABS
    g = bytes([0x24]) + struct.pack("<HB", 0xC000, 0) + fnid
    s = "字".encode("cp932")
    return b"".join([
        _r(bof, struct.pack("<HHH", biff << 8, 0x10, 0)),
        _r(0x0042, struct.pack("<H", 932)),
        _r(0x001E, fmt),
        _r(xf_t, xf),
        _r(0x0203, struct.pack("<HHH", 0, 0, 0) + struct.pack("<d", -3.0)),
        _r(0x0204, struct.pack("<HHH", 0, 1, 0) + struct.pack("<H", len(s)) + s),
        _r(0x0406 if biff == 4 else 0x0206, struct.pack("<HHH", 1, 0, 0) + struct.pack("<d", 3.0) + struct.pack("<H", 0) + struct.pack("<H", len(g)) + g),
        _r(0x000A, b""),
    ])


@pytest.mark.parametrize("make,biff", [(_biff2, 2), (lambda: _biff34(3), 3), (lambda: _biff34(4), 4)])
def test_excel_2_to_4(tmp_path, make, biff):
    blob = make()
    assert kind_of(blob) == "xls2"
    p = _p(tmp_path, "a.xls", blob)
    from todoku.book import open_book
    from todoku.xlsx.calc import verify

    wb = open_book(blob, "a.xls")
    assert wb.biff == biff and verify(wb).summary()["mismatch"] == 0
    assert any(c.formula for c in wb.sheets[0].cells.values())
    t = _text(p)
    assert ("申請" in t) if biff == 2 else ("字" in t)
    _ok_audit(p)
    g = api.grade(p)
    assert g["write"]["grade"] == "D" and g["grade"] in "AB"
    if biff == 2:
        assert wb.sheets[0].cells[(3, 1)].formula == "A2*2" and wb.sheets[0].cells[(3, 2)].formula == "SUM(A2:A2)"
    else:
        assert wb.sheets[0].cells[(2, 1)].formula == "ABS(A1)"
    with pytest.raises(api.ApiError, match="書きません"):
        api.fill(p, {"A1": "x"}, str(tmp_path / "b.xls"))


def test_excel_2_encrypted_is_f(tmp_path):
    blob = _r(0x0009, struct.pack("<HH", 2, 0x10)) + _r(0x002F, b"\0\0\0\0") + _r(0x000A, b"")
    assert api.grade(_p(tmp_path, "a.xls", blob))["grade"] == "F"


# ====================================================================== Lotus
def test_wk1_cells(tmp_path):
    s = "氏名".encode("cp932")
    blob = b"".join([_r(0, struct.pack("<H", 0x0406)),
                     _r(15, b"\xff" + struct.pack("<HH", 0, 0) + b"'" + s + b"\0"),
                     _r(14, b"\xff" + struct.pack("<HH", 1, 0) + struct.pack("<d", 1.5)),
                     _r(13, b"\xff" + struct.pack("<HH", 0, 1) + struct.pack("<h", -2)),
                     _r(16, b"\xff" + struct.pack("<HH", 1, 1) + struct.pack("<d", 9.0) + struct.pack("<H", 1) + b"\x03"),
                     _r(1, b"")])
    assert kind_of(blob) == "wk1"
    p = _p(tmp_path, "a.wk1", blob)
    t = _text(p)
    assert "氏名" in t and "Shift_JIS" in t
    _ok_audit(p)
    g = api.grade(p)
    assert g["grade"] == "C" and g["write"]["grade"] == "D"  # 読めない・Excel と意味の違う数式は保存された結果を読む


def test_wk3_and_works_and_docuworks_are_named_with_next_step(tmp_path):
    from test_embeds import make_cfb

    cases = [("a.123", b"\x00\x00\x1a\x00\x03\x10\x04\x00" + bytes(40), "Lotus", ".xlsx"),
             ("a.wps", make_cfb({"CONTENTS": b"CHNKWKS", "MN0": b"x"}), "Works", ".docx"),
             ("a.xdw", b"\x01\x02any bytes" + bytes(100), "DocuWorks", "DocuWorks Viewer Light"),
             ("a.xbd", b"\x01\x02any bytes" + bytes(100), "DocuWorks", "PDF")]
    for name, blob, word, step in cases:
        p = _p(tmp_path, name, blob)
        with pytest.raises(api.ApiError, match=word):
            api.window(p)
        g = api.grade(p)
        assert g["grade"] == "F" and step in str(g["reasons"]), name


def test_kingsoft_wps_with_worddocument_is_still_doc():
    from test_embeds import make_cfb

    assert kind_of(make_cfb({"WordDocument": bytes(1000)}), "a.wps") == "doc"


# ====================================================================== dBase
def _dbf(ldid=0x13, rows=None) -> bytes:
    fields = [(b"NAME", b"C", 12, 0), (b"NUM", b"N", 8, 2), (b"DAY", b"D", 8, 0), (b"OK", b"L", 1, 0), (b"NOTE", b"M", 10, 0)]
    rs = 1 + sum(f[2] for f in fields)
    rows = rows if rows is not None else [(b" ", "山田".encode("cp932").ljust(12), b"   12.50", b"20260401", b"T", b" " * 10),
                                          (b"*", b"deleted".ljust(12), b"    1.00", b"20260401", b"F", b" " * 10),
                                          (b" ", b"abc".ljust(12), b" " * 8, b" " * 8, b"?", b" " * 10)]
    hs = 32 + 32 * len(fields) + 1
    h = bytearray(32)
    h[0] = 0x03
    struct.pack_into("<IHH", h, 4, len(rows), hs, rs)
    h[29] = ldid
    d = bytes(h)
    for n, t, ln, dc in fields:
        d += n.ljust(11, b"\0") + t + bytes(4) + bytes([ln, dc]) + bytes(14)
    d += b"\x0d"
    for r in rows:
        d += b"".join(r)
    return d + b"\x1a"


def test_dbf_read_as_table(tmp_path):
    blob = _dbf()
    assert kind_of(blob, "a.dbf") == "dbf" and kind_of(blob) != "dbf"
    p = _p(tmp_path, "a.dbf", blob)
    s = api.summary(p)
    assert s["kind"] == "dbf" and s["sheets"] == ["a"]
    from todoku.book import open_book

    wb = open_book(blob, "a.dbf")
    sh = wb.sheets[0]
    assert [sh.cells[(1, c)].value for c in range(1, 6)] == ["NAME", "NUM", "DAY", "OK", "NOTE"]
    assert (sh.cells[(2, 1)].value, sh.cells[(2, 2)].value, sh.cells[(2, 3)].value, sh.cells[(2, 4)].value) == ("山田", 12.5, "2026-04-01", True)
    assert sh.cells[(3, 1)].value == "abc" and sh.max_row == 3
    assert any("削除" in u for u in sh.unread) and any("メモ" in u for u in sh.unread)
    _ok_audit(p)
    t = api.read_table(p, "NAME", "a")
    assert t["rows"] == 2
    g = api.grade(p)
    assert g["grade"] == "C" and g["write"]["grade"] == "B"


def test_dbf_rejects_inconsistent_header_and_guesses_with_note(tmp_path):
    from todoku.book import open_book

    bad = bytearray(_dbf())
    struct.pack_into("<H", bad, 10, 99)
    assert kind_of(bytes(bad), "a.dbf") != "dbf"
    wb = open_book(_dbf(ldid=0), "a.dbf")
    assert wb.encoding == "cp932" and any("推定" in u for u in wb.sheets[0].unread)


# ====================================================================== Access
@pytest.mark.parametrize("sizes", [(240, 10), (255, 1), (256, 1), (254, 254), (300, 20), (20, 700), (600, 700), (1200, 3)])
def test_jet3_long_rows_cross_256_boundaries(sizes):
    import jet_maker
    from todoku.book import open_book

    m = jet_maker.Maker(True)
    cols = [{"name": "番号", "type": 4, "num": 0, "fixed": 0, "size": 4}, {"name": "前", "type": 10, "num": 1, "var": 0, "size": 255},
            {"name": "空", "type": 10, "num": 2, "var": 1, "size": 255}, {"name": "後", "type": 10, "num": 3, "var": 2, "size": 255}]
    a, b = "あ" * (sizes[0] // 2), "z" * sizes[1]
    m.add_table("表", cols, [m.row(cols, {"番号": struct.pack("<i", 7), "前": m.text(a), "空": b"", "後": m.text(b)})])
    sh = open_book(m.build(), "a.mdb").sheets[0]
    assert sh.cells[(2, 1)].value == 7 and sh.cells[(2, 2)].value == a and sh.cells[(2, 4)].value == b


@pytest.mark.parametrize("jet3", [False, True])
def test_jet_tables(tmp_path, jet3):
    import jet_maker

    blob = jet_maker.sample(jet3)
    assert kind_of(blob) == "mdb"
    p = _p(tmp_path, "a.mdb", blob)
    assert api.summary(p)["sheets"] == ["コード表"]
    from todoku.book import open_book

    wb = open_book(blob, "a.mdb")
    sh = wb.sheets[0]
    assert [sh.cells[(1, c)].value for c in range(1, 8)] == ["番号", "単価", "可", "名前", "説明", "日付", "金額"]
    assert sh.cells[(2, 1)].value == 1 and sh.cells[(2, 2)].value == 12.5 and sh.cells[(2, 3)].value is True
    assert sh.cells[(2, 4)].value == "山田" and sh.cells[(2, 5)].value == "メモの字"
    assert sh.cells[(2, 6)].value == "2023-03-15" and sh.cells[(2, 7)].value == 1.2345
    assert sh.cells[(3, 1)].value == 3 and sh.cells[(3, 4)].value == "abc" and sh.cells[(3, 3)].value is False
    assert sh.max_row == 3  # 削除の印のある行は読まない
    _ok_audit(p)
    assert api.read_table(p, "番号", "コード表")["rows"] == 2
    g = api.grade(p)
    assert g["write"]["grade"] == "D" and g["grade"] in "ABC"
    with pytest.raises(api.ApiError, match="書きません"):
        api.fill(p, {"コード表!A2": 9}, str(tmp_path / "b.mdb"))


def test_jet_truncated_file_is_refused(tmp_path):
    import jet_maker

    blob = jet_maker.sample()[:4096 * 3]
    g = api.grade(_p(tmp_path, "a.mdb", blob))
    assert g["grade"] == "F"


# ====================================================================== Word 1.x・2.0
def _word2(text: bytes, lid=0x0411, flags=0, magic=0xA5DB, ftn: bytes = b"") -> bytes:
    fc = 0x180
    h = bytearray(fc)
    struct.pack_into("<HH", h, 0, magic, 0x2D)
    struct.pack_into("<H", h, 6, lid)
    struct.pack_into("<H", h, 0x0A, flags)
    struct.pack_into("<II", h, 0x18, fc, fc + len(text) + len(ftn))
    struct.pack_into("<III", h, 0x34, len(text), len(ftn), 0)
    return bytes(h) + text + ftn + bytes(32)


def test_word2_reads_text_tables_and_fields(tmp_path):
    t = ("申請書\r氏名\x07山田\x07\x07\r".encode("cp932") + b"\x13 PAGE \x141\x15\r")
    blob = _word2(t, ftn="脚注の字\r".encode("cp932"))
    assert kind_of(blob) == "word2"
    p = _p(tmp_path, "a.doc", blob)
    w = _text(p)
    assert "申請書" in w and "氏名 ｜ 山田" in w and "PAGE" not in w and "脚注の字" in w
    _ok_audit(p)
    g = api.grade(p)
    assert g["grade"] == "B" and g["write"]["grade"] == "D"


@pytest.mark.parametrize("flags,word", [(0x0100, "パスワード"), (0x0004, "高速保存")])
def test_word2_refuses_what_it_cannot_order(tmp_path, flags, word):
    p = _p(tmp_path, "a.doc", _word2(b"abc\r", flags=flags))
    g = api.grade(p)
    assert g["grade"] == "F" and word in str(g["reasons"])


# ====================================================================== OpenOffice.org 1.x
_SXW = ('<?xml version="1.0" encoding="UTF-8"?><!DOCTYPE office:document-content PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "office.dtd">'
        '<office:document-content xmlns:office="http://openoffice.org/2000/office" xmlns:text="http://openoffice.org/2000/text" xmlns:table="http://openoffice.org/2000/table" office:class="text" office:version="1.0">'
        '<office:body><text:h text:level="1">題の字</text:h><text:p>本文の<text:span>字</text:span></text:p>'
        '<table:table table:name="表1"><table:table-column/><table:table-row><table:table-cell><text:p>マスの字</text:p></table:table-cell></table:table-row></table:table>'
        '</office:body></office:document-content>')
_SXC = ('<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="http://openoffice.org/2000/office" xmlns:text="http://openoffice.org/2000/text" xmlns:table="http://openoffice.org/2000/table" office:class="spreadsheet" office:version="1.0">'
        '<office:body><table:table table:name="計算"><table:table-column/><table:table-row>'
        '<table:table-cell table:value-type="string"><text:p>氏名</text:p></table:table-cell>'
        '<table:table-cell table:value-type="float" table:value="12.5"><text:p>12.5</text:p></table:table-cell>'
        '</table:table-row></table:table></office:body></office:document-content>')


def test_sxw_and_sxc(tmp_path):
    # LibreOffice に、同じ中身に manifest を足したものを読ませ、同じ字になることを確かめた
    sxw = _zip({"mimetype": "application/vnd.sun.xml.writer", "content.xml": _SXW})
    assert kind_of(sxw) == "sxw"
    p = _p(tmp_path, "a.sxw", sxw)
    t = _text(p)
    assert "題の字" in t and "本文の字" in t and "マスの字" in t
    _ok_audit(p)
    g = api.grade(p)
    assert g["grade"] == "B" and g["write"]["grade"] == "D"
    sxc = _zip({"mimetype": "application/vnd.sun.xml.calc", "content.xml": _SXC})
    assert kind_of(sxc) == "sxc"
    p2 = _p(tmp_path, "a.sxc", sxc)
    t2 = _text(p2)
    assert "氏名" in t2 and "12.5" in t2
    assert api.summary(p2)["sheets"] == ["計算"]
    _ok_audit(p2)
    with pytest.raises(api.ApiError, match="書きません"):
        api.fill(p2, {"計算!B1": 3}, str(tmp_path / "b.sxc"))


# ====================================================================== iWork
def _vi(n):
    out = bytearray()
    while True:
        b = n & 0x7F
        n >>= 7
        out.append(b | 0x80 if n else b)
        if not n:
            return bytes(out)


def _f(num, v):
    return _vi(num << 3 | 2) + _vi(len(v)) + v if isinstance(v, bytes) else _vi(num << 3) + _vi(v)


def _snappy_literal_with_copy(raw: bytes) -> bytes:
    """字そのままのあとに、前の字の写し（タグ 1）を 1 つ入れた Snappy。"""
    n = len(raw)
    return _vi(n + 4) + bytes([(n - 1) << 2]) + raw + bytes([0x01 | (0 << 2) | (0 << 5), n & 0xFF]) if n < 60 else (_vi(n) + bytes([61 << 2]) + struct.pack("<H", n - 1) + raw)


def _iwa(msgs) -> bytes:
    """msgs：[(識別子, 型, 中身)]。"""
    plain = b""
    for ident, t, payload in msgs:
        info = _f(1, ident) + _f(2, _f(1, t) + _f(3, len(payload)))
        plain += _vi(len(info)) + info + payload
    body = _vi(len(plain)) + bytes([61 << 2]) + struct.pack("<H", len(plain) - 1) + plain
    return b"\x00" + struct.pack("<I", len(body))[:3] + body


def test_iwork_pages_text(tmp_path):
    body = _f(1, 0) + _f(3, "本文の一行目\n二行目".encode())
    box = _f(1, 3) + _f(3, "図形の字".encode())
    other = _f(1, 5)  # 字のないもの
    table = _f(3, _f(1, 1) + _f(3, "マスの字".encode()))
    blob = _zip({"Index/Document.iwa": _iwa([(7, 2001, box), (3, 2001, body), (5, 9999, other)]), "Index/Tables.iwa": _iwa([(9, 6005, table)]), "preview.jpg": b"x"})
    assert kind_of(blob) == "iwork"
    p = _p(tmp_path, "a.pages", blob)
    t = _text(p)
    assert t.index("本文の一行目") < t.index("図形の字") and "マスの字" in t and "公開していない" in t
    _ok_audit(p)
    g = api.grade(p)
    assert g["grade"] == "C" and g["write"]["grade"] == "D"


def test_snappy_copies():
    from todoku.legacy.iwork import unsnappy

    # 字 "abcd" のあと、距離 2・長さ 5 の写し（重なりあり）→ "abcdcdcdc"
    d = _vi(9) + bytes([3 << 2]) + b"abcd" + bytes([(5 - 4) << 2 | 1, 2])
    assert unsnappy(d) == b"abcdcdcdc"
    d2 = _vi(6) + bytes([1 << 2]) + b"ab" + bytes([(4 - 1) << 2 | 2]) + struct.pack("<H", 2)
    assert unsnappy(d2) == b"ababab"


def test_iwork_without_readable_parts_is_f(tmp_path):
    blob = _zip({"Index/Document.iwa": b"\x01garbage", "preview.jpg": b"x"})
    g = api.grade(_p(tmp_path, "a.numbers", blob))
    assert g["grade"] == "F" and "書き出して" in str(g["reasons"])


def test_docuworks_inside_a_zip_is_named(tmp_path):
    p = _p(tmp_path, "a.zip", _zip({"a.xdw": b"\x01" + bytes(50)}))
    g = api.grade(p)
    assert g["grade"] == "F" and "DocuWorks" in str(g["reasons"])


def test_xps_page_in_utf16(tmp_path):
    ns = "http://schemas.microsoft.com/xps/2005/06"
    page = f'<FixedPage xmlns="{ns}"><Glyphs UnicodeString="二重の字" OriginX="1" OriginY="1" FontRenderingEmSize="10" Fill="#000"/></FixedPage>'
    blob = _zip({
        "_rels/.rels": f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r" Type="{ns}/fixedrepresentation" Target="/s.fdseq"/></Relationships>',
        "s.fdseq": f'<FixedDocumentSequence xmlns="{ns}"><DocumentReference Source="d.fdoc"/></FixedDocumentSequence>',
        "d.fdoc": f'<FixedDocument xmlns="{ns}"><PageContent Source="1.fpage"/></FixedDocument>',
        "1.fpage": ('<?xml version="1.0" encoding="UTF-16"?>' + page).encode("utf-16"),
    })
    p = _p(tmp_path, "a.xps", blob)
    assert "二重の字" in _text(p)
    au = _ok_audit(p)
    assert "4 字" in au["checks"][0]["detail"]
