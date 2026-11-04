"""試験用の .xlsb を、MS-XLSB の仕様どおりにレコードから手で作る（Excel も LibreOffice も .xlsb を書き出せないため）。"""

from __future__ import annotations

import io
import struct
import zipfile


def rec(rid: int, body: bytes = b"") -> bytes:
    out = bytearray()
    if rid < 0x80:
        out.append(rid)
    else:
        out += bytes([(rid & 0x7F) | 0x80, (rid >> 7) & 0x7F])
    n = len(body)
    while True:
        b = n & 0x7F
        n >>= 7
        out.append(b | 0x80 if n else b)
        if not n:
            break
    return bytes(out) + body


def ws(s: str) -> bytes:
    return struct.pack("<I", len(s)) + s.encode("utf-16-le")


def cell(col: int, xf: int = 0) -> bytes:
    return struct.pack("<II", col, xf)


def row(r: int, flags: int = 0, height: int = 0) -> bytes:
    return rec(0, struct.pack("<IIHBBBI", r, 0, height, 0, flags, 0, 0))


def real(col, v, xf=0):
    return rec(5, cell(col, xf) + struct.pack("<d", v))


def rk_int(col, v, xf=0):
    return rec(2, cell(col, xf) + struct.pack("<I", (v << 2) | 2))


def st(col, s, xf=0):
    return rec(6, cell(col, xf) + ws(s))


def isst(col, i, xf=0):
    return rec(7, cell(col, xf) + struct.pack("<I", i))


def blank(col, xf=0):
    return rec(1, cell(col, xf))


def boolean(col, v, xf=0):
    return rec(4, cell(col, xf) + bytes([1 if v else 0]))


def err(col, code, xf=0):
    return rec(3, cell(col, xf) + bytes([code]))


def _fmla(rgce: bytes, extra: bytes = b"") -> bytes:
    return struct.pack("<H", 0) + struct.pack("<I", len(rgce)) + rgce + struct.pack("<I", len(extra)) + extra


def fnum(col, v, rgce, extra=b"", xf=0):
    return rec(9, cell(col, xf) + struct.pack("<d", v) + _fmla(rgce, extra))


def fstr(col, v, rgce, extra=b"", xf=0):
    return rec(8, cell(col, xf) + ws(v) + _fmla(rgce, extra))


# ------------------------------------------------------------ トークン
def p_ref(r, c, rel_r=True, rel_c=True):
    return bytes([0x24]) + struct.pack("<IH", r, c | (0x4000 if rel_c else 0) | (0x8000 if rel_r else 0))


def p_area(r1, r2, c1, c2):
    return bytes([0x25]) + struct.pack("<IIHH", r1, r2, c1 | 0xC000, c2 | 0xC000)


def p_ref3d(ixti, r, c):
    return bytes([0x3A]) + struct.pack("<HIH", ixti, r, c)


def p_int(v):
    return bytes([0x1E]) + struct.pack("<H", v)


def p_str(s):
    return bytes([0x17]) + struct.pack("<H", len(s)) + s.encode("utf-16-le")


def p_fn_var(iftab, argc):
    return bytes([0x22, argc]) + struct.pack("<H", iftab)


def p_fn(iftab):
    return bytes([0x21]) + struct.pack("<H", iftab)


def p_name(idx):
    return bytes([0x23]) + struct.pack("<I", idx)


ADD, MUL, CONCAT = b"\x03", b"\x05", b"\x08"


def rfx(r1, r2, c1, c2):
    return struct.pack("<IIII", r1, r2, c1, c2)


def name_rec(name, rgce, itab=0xFFFFFFFF, flags=0):
    return rec(39, struct.pack("<I", flags) + b"\x00" + struct.pack("<I", itab) + ws(name) + struct.pack("<I", len(rgce)) + rgce
               + struct.pack("<I", 0) + b"\xff\xff\xff\xff")


def bundle(sheet_id, relid, name, state=0):
    return rec(156, struct.pack("<II", state, sheet_id) + ws(relid) + ws(name))


def color_rgb(r, g, b):
    return bytes([2 << 1 | 1, 0, 0, 0, r, g, b, 255])


def styles() -> bytes:
    """xf 0：ふつう。xf 1：ロックしない（入力欄）。xf 2：表示形式 #,##0。xf 3：黄色の塗りと下の罫線、赤い字。"""
    font0 = struct.pack("<HHHHBBBB", 220, 0, 400, 0, 0, 2, 128, 0) + bytes([1 << 1, 8, 0, 0, 0, 0, 0, 255]) + b"\x00" + ws("MS Gothic")
    font1 = struct.pack("<HHHHBBBB", 220, 0, 400, 0, 0, 2, 128, 0) + color_rgb(255, 0, 0) + b"\x00" + ws("MS Gothic")
    nofill = struct.pack("<I", 0) + bytes(16) + bytes(48)
    yellow = struct.pack("<I", 1) + color_rgb(255, 255, 0) + bytes(8) + bytes(48)
    noborder = b"\x00" + bytes(50)
    bottom = b"\x00" + bytes(10) + bytes([1, 0]) + bytes(8) + bytes(30)
    def xf(parent, ifmt, ifont, ifill, ibrd, bits):
        return rec(47, struct.pack("<HHHHHBBHH", parent, ifmt, ifont, ifill, ibrd, 0, 0, bits, 0))
    return b"".join([
        rec(278),
        rec(615, struct.pack("<I", 1)), rec(44, struct.pack("<H", 164) + ws("#,##0\"円\"")), rec(616),
        rec(611, struct.pack("<I", 2)), rec(43, font0), rec(43, font1), rec(612),
        rec(603, struct.pack("<I", 2)), rec(45, nofill), rec(45, yellow), rec(604),
        rec(613, struct.pack("<I", 2)), rec(46, noborder), rec(46, bottom), rec(614),
        rec(626, struct.pack("<I", 1)), xf(0xFFFF, 0, 0, 0, 0, 0x1000), rec(627),
        rec(617, struct.pack("<I", 4)), xf(0, 0, 0, 0, 0, 0x1000), xf(0, 0, 0, 0, 0, 0),
        xf(0, 164, 0, 0, 0, 0x1000), xf(0, 0, 1, 1, 1, 0x1000), rec(618),
        rec(279),
    ])


def sst(strings: list[str]) -> bytes:
    return (rec(159, struct.pack("<II", len(strings), len(strings)))
            + b"".join(rec(19, b"\x00" + ws(s)) for s in strings) + rec(160))


def package(sheets: dict[str, bytes], names: list[bytes] | None = None, strings: list[str] | None = None,
            extern: bytes | None = None, extra: dict[str, bytes] | None = None, sheet_rels: dict[str, str] | None = None,
            states: dict[str, int] | None = None) -> bytes:
    """sheets：シートの名前 → シートの部品（レコードの並び）。"""
    wbrels = []
    bundles = []
    ct = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
          '<Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>',
          '<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>',
          '<Override PartName="/xl/styles.bin" ContentType="application/vnd.ms-excel.styles"/>']
    parts: dict[str, bytes] = {}
    for k, (nm, body) in enumerate(sheets.items(), 1):
        parts[f"xl/worksheets/sheet{k}.bin"] = body
        wbrels.append(f'<Relationship Id="rId{k}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{k}.bin"/>')
        bundles.append(bundle(k, f"rId{k}", nm, (states or {}).get(nm, 0)))
        ct.append(f'<Override PartName="/xl/worksheets/sheet{k}.bin" ContentType="application/vnd.ms-excel.worksheet"/>')
        if sheet_rels and nm in sheet_rels:
            parts[f"xl/worksheets/_rels/sheet{k}.bin.rels"] = sheet_rels[nm].encode()
    wbrels.append('<Relationship Id="rId100" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/>')
    if strings is not None:
        parts["xl/sharedStrings.bin"] = sst(strings)
        wbrels.append('<Relationship Id="rId101" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.bin"/>')
        ct.append('<Override PartName="/xl/sharedStrings.bin" ContentType="application/vnd.ms-excel.sharedStrings"/>')
    wb = (rec(131) + rec(143) + b"".join(bundles) + rec(144) + (extern or b"") + b"".join(names or []) + rec(132))
    parts["xl/workbook.bin"] = wb
    parts["xl/styles.bin"] = styles()
    parts["xl/_rels/workbook.bin.rels"] = (
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(wbrels) + "</Relationships>").encode()
    parts["_rels/.rels"] = (
        '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>').encode()
    parts["[Content_Types].xml"] = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                                    + "".join(ct) + "</Types>").encode()
    parts.update(extra or {})
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", parts.pop("[Content_Types].xml"))
        for n, b in parts.items():
            z.writestr(n, b)
    return buf.getvalue()


def sheet(rows: bytes, tail: bytes = b"", head: bytes = b"") -> bytes:
    return rec(129) + head + rec(145) + rows + rec(146) + tail + rec(130)


def sample() -> bytes:
    """2 シート：字・数・共有文字列・空・真偽・エラー・数式（SUM・積・別シート・名前・字の連結・共有数式）・結合・入力規則・保護・非表示の行。"""
    extern = (rec(353) + rec(357) + rec(362, struct.pack("<I", 1) + struct.pack("<Iii", 0, 1, 1)) + rec(354))
    names = [name_rec("税率", p_ref3d(0, 0, 0))]
    dv = (rec(573, struct.pack("<IIII", 0, 0, 0, 1))
          + rec(64, struct.pack("<I", 3 | 0x100 | 0x80000) + struct.pack("<I", 1) + rfx(4, 5, 1, 1)
                + ws("題") + ws("選んでください") + ws("入力の題") + ws("入力の説明")
                + struct.pack("<I", len(p_str("甲,乙"))) + p_str("甲,乙") + struct.pack("<I", 0)
                + struct.pack("<I", 0) + struct.pack("<I", 0))
          + rec(574))
    r = b"".join([
        row(0), isst(0, 0), isst(1, 1), st(2, "直接の字"),
        row(1), st(0, "単価"), real(1, 1234.5, 2), rk_int(2, 7),
        row(2), st(0, "合計"), fnum(1, 8641.5, p_ref(1, 1) + p_ref(1, 2) + MUL, xf=2),
        fnum(2, 1241.5, p_area(1, 1, 1, 2) + bytes([0x19, 0x10, 0, 0])),
        row(3), st(0, "税"), fnum(1, 0.1, p_name(1)), fstr(2, "単価円", p_ref(1, 0) + p_str("円") + CONCAT),
        row(4), st(0, "区分"), blank(1, 1), boolean(2, True),
        row(5), st(0, "備考"), blank(1, 1), err(2, 0x07),
        row(6, flags=0x10), st(0, "隠れた行"),
    ])
    s1 = sheet(r, tail=rec(177, struct.pack("<I", 1)) + rec(176, rfx(0, 0, 2, 3)) + rec(178) + dv)
    s2 = sheet(row(0) + real(0, 0.1, 0) + st(1, "税率"))
    return package({"申請": s1, "基準": s2}, names=names, strings=["品名", "数量"], extern=extern)
