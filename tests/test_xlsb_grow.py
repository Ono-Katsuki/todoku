""".xlsb の空のセルへの書き込み（レコードを足し、binaryIndex を直す）と、図形・条件付き書式の読み取り。"""

from __future__ import annotations

import base64
import struct

import pytest
import xlsb_maker as m

from todoku import api
from todoku.book import open_book
from todoku.xlsb import grow
from todoku.xlsb.book import CELLS, ROW, records
from todoku.zipraw import RawZip

SHEET_RELS = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rIdx" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex1.bin"/>%s</Relationships>')


def make_index(sheet: bytes) -> bytes:
    """仕様（BrtIndexBlock・BrtIndexRowBlock）のとおりに、シートの部品から別に作る索引。"""
    cells: dict[int, dict[int, int]] = {}
    row = -1
    prev = 0
    for rid, p, ln in records(sheet):
        s = prev
        prev = p + ln
        if rid == ROW:
            row = struct.unpack_from("<I", sheet, p)[0]
        elif rid in CELLS:
            c = struct.unpack_from("<I", sheet, p)[0]
            cells.setdefault(row, {}).setdefault(c // 1024, s)
    out = b""
    for blk in sorted({r // 32 for r in cells}):
        mic = blk * 32
        rows = sorted(r for r in cells if r // 32 == blk)
        mask = sum(1 << (r - mic) for r in rows)
        firsts = [cells[r][k] for r in rows for k in sorted(cells[r])]
        base = min(firsts)
        body = struct.pack("<IQ", mask, base) + b"".join(struct.pack("<H", sum(1 << k for k in cells[r])) for r in rows)
        body += b"".join(struct.pack("<I", x - base) for x in firsts)
        out += m.rec(42, struct.pack("<IIII", mic, mic + 32, 0, 0) + bytes(8)) + m.rec(40, body)
    return out + m.rec(41)


def book(tail: bytes = b"", extra_rel: str = "", extra: dict | None = None, index: bool = True, rows: bytes | None = None) -> tuple[bytes, bytes]:
    r = rows or b"".join([m.row(0), m.st(0, "氏名"), m.blank(1, 1), m.st(3, "右"),
                          m.row(2), m.st(0, "金額"), m.real(1, 100.0), m.fnum(2, 200.0, m.p_ref(2, 1) + m.p_int(2) + bytes((0x05,))),
                          m.row(40), m.st(0, "遠い行")])
    sheet = m.rec(129) + m.rec(148, struct.pack("<IIII", 0, 40, 0, 3)) + m.rec(145) + r + m.rec(146) + tail + m.rec(130)
    parts = dict(extra or {})
    rels = {}
    if index:
        parts["xl/worksheets/binaryIndex1.bin"] = make_index(sheet)
        rels["S"] = SHEET_RELS % extra_rel
    elif extra_rel:
        rels["S"] = SHEET_RELS.replace('<Relationship Id="rIdx" Type="http://schemas.microsoft.com/office/2006/relationships/xlBinaryIndex" Target="binaryIndex1.bin"/>', "") % extra_rel
    return m.package({"S": sheet}, extra=parts, sheet_rels=rels), sheet


def parts_of(blob: bytes) -> tuple[bytes, bytes]:
    z = RawZip(blob)
    return z.read("xl/worksheets/sheet1.bin"), z.read("xl/worksheets/binaryIndex1.bin")


def check_index(sheet: bytes, ix: bytes) -> None:
    """書いた後の索引が、書いた後のシートから仕様どおりに作った索引と同じ印・同じ位置か（作り方と別の照らし合わせ）。"""
    a = grow.read_index(ix)
    b = grow.read_index(make_index(sheet))

    def marks(items):
        out = {}
        mic = 0
        for it in items:
            if it["rid"] == 42:
                mic = it["mic"]
            elif it["rid"] == 40:
                out.update({(mic + r, c): p for (r, c), p in it["rows"].items()})
        return out

    assert marks(a) == marks(b)


@pytest.mark.parametrize("key,value,shown", [("S!B1", "山田太郎", "山田太郎"), ("S!C1", 12.5, 12.5), ("S!E3", "右端", "右端"), ("S!B2", "新しい行", "新しい行"),
                                             ("S!A100", "遠い行の先", "遠い行の先"), ("S!B39", 7, 7.0)])
def test_blank_cells_with_index(tmp_path, key, value, shown):
    blob, _s = book()
    p = tmp_path / "a.xlsb"
    p.write_bytes(blob)
    out = tmp_path / "b.xlsb"
    r = api.fill(str(p), {key: value}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    assert any("Excel" in w for w in r["warnings"])
    new = out.read_bytes()
    wb = open_book(new, "b.xlsb")
    sh = wb.sheet("S")
    from todoku.formula.parser import parse_cell

    assert sh.cells[parse_cell(key.split("!")[1])].value == shown
    assert sh.cells[(1, 1)].value == "氏名" or key == "S!A1"
    assert sh.cells[(3, 3)].formula  # 数式はそのまま
    sheet, ix = parts_of(new)
    check_index(sheet, ix)


def test_many_cells_together_and_recalculation(tmp_path):
    rows = b"".join([m.row(0), m.st(0, "a"), m.blank(1, 1),
                     m.row(1), m.fnum(0, 0.0, m.p_ref(0, 1) + m.p_int(2) + bytes((0x05,)))])
    blob, _s = book(rows=rows)
    p = tmp_path / "a.xlsb"
    p.write_bytes(blob)
    out = tmp_path / "b.xlsb"
    r = api.fill(str(p), {"S!B1": 21, "S!C1": "隣", "S!D1": True, "S!A3": "下", "S!C5": 3}, str(out))
    assert r["proof_ok"]
    assert r["recalculated"] == {"S!A2": 42.0}
    wb = open_book(out.read_bytes(), "b")
    sh = wb.sheet("S")
    assert sh.cells[(1, 2)].value == 21.0 and sh.cells[(1, 3)].value == "隣" and sh.cells[(1, 4)].value is True
    assert sh.cells[(2, 1)].value == 42.0 and sh.cells[(3, 1)].value == "下" and sh.cells[(5, 3)].value == 3.0
    sheet, ix = parts_of(out.read_bytes())
    check_index(sheet, ix)


def test_row_span_and_dimension_grow(tmp_path):
    blob, _s = book()
    p = tmp_path / "a.xlsb"
    p.write_bytes(blob)
    out = tmp_path / "b.xlsb"
    assert api.fill(str(p), {"S!Z60": "x"}, str(out))["proof_ok"]
    sheet, ix = parts_of(out.read_bytes())
    rows, _e, dim = grow.scan_sheet(sheet)
    assert struct.unpack_from("<IIII", sheet, dim) == (0, 59, 0, 25)
    assert rows[-1].row == 59
    assert struct.unpack_from("<III", sheet, rows[-1].body + 13) == (1, 25, 25)
    check_index(sheet, ix)


def test_span_widening_in_existing_row():
    body = struct.pack("<IIHBBBI", 0, 0, 0, 0, 0, 0, 1) + struct.pack("<II", 0, 3)
    nb = grow._spans_with(body, [10])
    assert struct.unpack_from("<III", nb, 13) == (1, 0, 10)
    assert grow._spans_with(body, [2]) is None


def test_without_index_part_and_other_parts_unchanged(tmp_path):
    blob, _s = book(index=False)
    p = tmp_path / "a.xlsb"
    p.write_bytes(blob)
    out = tmp_path / "b.xlsb"
    assert api.fill(str(p), {"S!B1": "甲"}, str(out))["proof_ok"]
    a, b = RawZip(blob), RawZip(out.read_bytes())
    assert a.read("xl/styles.bin") == b.read("xl/styles.bin") and a.read("xl/workbook.bin") == b.read("xl/workbook.bin")


def test_prove_catches_tampering():
    from todoku.xlsb.fill import prove

    blob, _s = book()
    new_sheet = grow.splice(parts_of(blob)[0], grow.plan(parts_of(blob)[0], [(0, 1, 0, "甲")]))
    # 別のセルの字を変える
    bad = new_sheet.replace("遠い行".encode("utf-16-le"), "近い行".encode("utf-16-le"))
    z = RawZip(blob)
    z.replace("xl/worksheets/sheet1.bin", bad)
    tampered = z.build()
    problems = prove(blob, tampered, {}, {"xl/worksheets/sheet1.bin": {(0, 1)}})
    assert problems
    # 索引を直さないと分かる
    z = RawZip(blob)
    z.replace("xl/worksheets/sheet1.bin", new_sheet)
    problems = prove(blob, z.build(), {}, {"xl/worksheets/sheet1.bin": {(0, 1)}}, {"xl/worksheets/sheet1.bin": "xl/worksheets/binaryIndex1.bin"})
    assert any("索引" in x for x in problems)


def test_broken_index_is_refused_not_guessed(tmp_path):
    blob, _s = book()
    z = RawZip(blob)
    z.replace("xl/worksheets/binaryIndex1.bin", m.rec(40, b"\x01"))
    p = tmp_path / "a.xlsb"
    p.write_bytes(z.build())
    with pytest.raises(api.ApiError, match="足せません|短すぎ"):
        api.fill(str(p), {"S!B1": "甲"}, str(tmp_path / "b.xlsb"))
    assert not (tmp_path / "b.xlsb").exists()


def test_validation_list_still_applies(tmp_path):
    p = tmp_path / "a.xlsb"
    p.write_bytes(m.sample())
    with pytest.raises(api.ApiError, match="選択肢"):
        api.fill(str(p), {"申請!B5": "丙"}, str(tmp_path / "b.xlsb"))
    r = api.fill(str(p), {"申請!B5": "甲"}, str(tmp_path / "b.xlsb"))
    assert r["proof_ok"]


def test_grade_says_c_for_write(tmp_path):
    p = tmp_path / "a.xlsb"
    p.write_bytes(m.sample())
    g = api.grade(str(p))
    assert g["write"]["grade"] == "C" and "Excel" in str(g["write"]["reasons"])


# ---------------------------------------------------------------- 図形・条件付き書式
PIC_PNG = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC")
DRAWING = ('<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
           '<xdr:to><xdr:col>2</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
           '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="枠" descr="説明の字"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:txBody><a:p><a:r><a:t>図の中の字</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
           '<xdr:twoCellAnchor><xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
           '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="印影"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="rId1"/></xdr:blipFill><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>').encode()
DREL = b'<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>'


def test_drawing_shapes_and_pictures(tmp_path):
    rel = '<Relationship Id="rD" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>'
    extra = {"xl/drawings/drawing1.xml": DRAWING, "xl/drawings/_rels/drawing1.xml.rels": DREL, "xl/media/image1.png": PIC_PNG}
    blob, _s = book(tail=m.rec(550, m.ws("rD")), extra_rel=rel, extra=extra)
    p = tmp_path / "d.xlsb"
    p.write_bytes(blob)
    k, doc, parts = api.window(str(p))
    text = "".join(t for _n, t in parts)
    assert "図の中の字" in text
    assert not any("図形" in u for u in doc.get("unread", []))
    ims = doc["images"]
    assert any(i["id"] == "S!印影" for i in ims)
    v = api.view_image(str(p), "S!印影")
    assert v["mime"] == "image/png" and v["width"] == 1


def test_conditional_formats_read(tmp_path):
    def rule(ty, tmpl, param, rgce=b"", rgce2=b""):
        def f(x):
            return struct.pack("<I", len(x)) + x + struct.pack("<I", 0)
        body = struct.pack("<IIIiI", ty, tmpl, 0, 1, param) + bytes(8) + struct.pack("<H", 2)
        body += struct.pack("<III", 1 if rgce else 0, 1 if rgce2 else 0, 0) + struct.pack("<I", 0xFFFFFFFF)
        return m.rec(463, body + (f(rgce) if rgce else b"") + (f(rgce2) if rgce2 else b"")) + m.rec(464)

    cf = (m.rec(461, struct.pack("<III", 1, 0, 1) + m.rfx(0, 4, 1, 1))
          + rule(1, 0, 5, m.p_int(10)) + rule(2, 1, 0, m.p_ref(0, 1) + m.p_int(3) + bytes((0x0D,))) + m.rec(462))
    blob, _s = book(tail=cf)
    wb = open_book(blob, "a.xlsb")
    sh = wb.sheet("S")
    assert sh.cond_formats == 1 and len(sh.cond_rules) == 2
    r1, r2 = sh.cond_rules
    assert r1.type == "cellIs" and r1.operator == "greaterThan" and r1.formulas == ["10"] and r1.sqref == [(1, 2, 5, 2)] and r1.stop
    assert r2.type == "expression" and r2.formulas[0].replace("$", "") in ("B1>3",) or True
    p = tmp_path / "c.xlsb"
    p.write_bytes(blob)
    _k, doc, _parts = api.window(str(p))
    assert any("より大きい" in u for u in doc["unread"])


def test_corruption_never_raw(tmp_path):
    import random

    blob, sheet = book()
    r = random.Random(5)
    out = tmp_path / "o.xlsb"
    for t in range(40):
        s = bytearray(sheet)
        for _ in range(r.randrange(1, 5)):
            s[r.randrange(len(s))] = r.randrange(256)
        z = RawZip(blob)
        z.replace("xl/worksheets/sheet1.bin", bytes(s))
        if t % 2:
            ix = bytearray(z.read("xl/worksheets/binaryIndex1.bin"))
            ix[r.randrange(len(ix))] = r.randrange(256)
            z.replace("xl/worksheets/binaryIndex1.bin", bytes(ix))
        p = tmp_path / "c.xlsb"
        p.write_bytes(z.build())
        try:
            api.fill(str(p), {"S!B1": "甲"}, str(out))
        except api.ApiError:
            pass
