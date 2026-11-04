"""実物のファイル（別の道具と照らした）で見つかった不具合の、自作の試験。実物そのものはリポジトリに入れない。

何と照らして見つかったかは、各試験の説明に書く。"""

from __future__ import annotations

import io
import struct
import zipfile

import xlsb_maker as m
from test_more_formats import _r

from todoku import api
from todoku.book import open_book
from todoku.formula.parser import parse_cell
from todoku.xlsb.audit import raw_values
from todoku.xlsb.grow import cell_positions
from todoku.zipraw import RawZip


def _p(tmp_path, name, blob):
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


# ---------------------------------------------------------------- .xlsb
def _frt_sheet(tail_rid: int = 1) -> bytes:
    """Excel が書く形：データの後の BrtFRTBegin〜End の中に、番号 1（セルの番号と同じ）のレコード。"""
    rows = m.row(0) + m.real(0, 1.5) + m.row(1) + m.real(0, 3.7477917436899863e-22)
    frt = m.rec(35, struct.pack("<I", 0x10E02)) + m.rec(tail_rid, bytes(9)) + m.rec(36)
    return m.sheet(rows, tail=frt)


def test_xlsb_record_in_future_frame_does_not_erase_last_row():
    """Excel の書いた .xlsb で、最後の行の値が消えていた（Excel の表示と照らして発見）。"""
    wb = open_book(m.package({"S": _frt_sheet()}), "a.xlsb")
    assert wb.sheet("S").cells[(2, 1)].value == 3.7477917436899863e-22


def test_xlsb_cells_only_inside_sheet_data_for_all_readers(tmp_path):
    blob = m.package({"S": _frt_sheet(tail_rid=5)})  # 番号 5（数のセルと同じ）でも
    wb = open_book(blob, "a.xlsb")
    assert set(wb.sheet("S").cells) == {(1, 1), (2, 1)}
    assert raw_values(wb)["S"] == {(1, 1), (2, 1)}  # 検査の数え方も同じ
    z = RawZip(blob)
    assert set(cell_positions(z.read("xl/worksheets/sheet1.bin"))) == {(0, 0), (1, 0)}
    assert api.audit(_p(tmp_path, "a.xlsb", blob))["checks"][0]["ok"]


def _rename_part(blob: bytes, old: str, new: str) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as a, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as b:
        for i in a.infolist():
            b.writestr(new if i.filename == old else i.filename, a.read(i.filename))
    return out.getvalue()


def test_part_names_do_not_depend_on_case():
    """OPC の部品の名前は大文字小文字を区別しない（Excel は開く。とどくは「共有文字列がありません」と止まっていた）。"""
    rows = m.row(0) + m.isst(0, 0)
    blob = _rename_part(m.package({"S": m.sheet(rows)}, strings=["共有の字"]), "xl/sharedStrings.bin", "xl/SharedStrings.bin")
    assert open_book(blob, "a.xlsb").sheet("S").cells[(1, 1)].value == "共有の字"


# ---------------------------------------------------------------- .xls の共有文字列
def _xls(strings: list[str], declared: int, cells: bytes) -> bytes:
    from test_embeds import make_cfb

    sst = struct.pack("<II", len(strings), declared) + b"".join(struct.pack("<HB", len(s), 1) + s.encode("utf-16-le") for s in strings)
    head = _r(0x0809, struct.pack("<HHHHII", 0x600, 5, 0, 0, 0, 6)) + _r(0x00E0, bytes(20))
    tail = _r(0x00FC, sst) + _r(0x000A, b"")
    bs = lambda at: _r(0x0085, struct.pack("<IBB", at, 0, 0) + bytes([1, 1]) + "S".encode("utf-16-le"))  # noqa: E731
    wbk = head + bs(len(head) + len(bs(0)) + len(tail)) + tail
    sh = _r(0x0809, struct.pack("<HHHHII", 0x600, 0x10, 0, 0, 0, 6)) + cells + _r(0x000A, b"")
    return make_cfb({"Workbook": wbk + sh})


def test_xls_sst_count_may_be_larger_than_the_strings():
    """宣言の数より字が少ない .xls を Excel は開く（とどくは開けなかった）。"""
    cell = _r(0x00FD, struct.pack("<HHHI", 0, 0, 0, 1))
    wb = open_book(_xls(["甲", "乙"], 9, cell), "a.xls")
    assert wb.sheet("S").cells[(1, 1)].value == "乙"


def test_xls_string_number_outside_the_table_is_reported_not_blank():
    cell = _r(0x00FD, struct.pack("<HHHI", 0, 0, 0, 5))
    wb = open_book(_xls(["甲", "乙"], 2, cell), "a.xls")
    sh = wb.sheet("S")
    assert any("A1" in u and "5" in u for u in sh.unread), sh.unread
    assert sh.cells[(1, 1)].value in (None, "")


# ---------------------------------------------------------------- Excel 4 の書式
def test_excel4_format_records_are_read():
    """BIFF4 の FORMAT は 0x041E（0x001E と読んでいて、すべての書式を落とし、(2) を 2、168,184,400 を 168184400 と見せていた。Excel の表示と照らして発見）。"""
    fmts = [b"\x07General", b"\x05#,##0", b"\x030.0"]
    xf = [bytes([0, i, 0, 0]) + bytes(4) + bytes(4) for i in range(3)]
    blob = b"".join([
        _r(0x0409, struct.pack("<HHH", 4 << 8, 0x10, 0)),
        _r(0x0042, struct.pack("<H", 932)),
        *[_r(0x041E, b"\x00\x00" + f) for f in fmts],
        *[_r(0x0443, x) for x in xf],
        _r(0x0203, struct.pack("<HHH", 0, 0, 1) + struct.pack("<d", 168184400.0)),
        _r(0x0203, struct.pack("<HHH", 1, 0, 2) + struct.pack("<d", 0.25)),
        _r(0x000A, b""),
    ])
    wb = open_book(blob, "a.xls")
    sh = wb.sheets[0]
    assert wb.style(sh.cells[(1, 1)].style).num_fmt == "#,##0"
    assert wb.style(sh.cells[(2, 1)].style).num_fmt == "0.0"


# ---------------------------------------------------------------- .ods
def _ods(cell_xml: str) -> bytes:
    content = ('<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
               'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
               'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:version="1.2"><office:body><office:spreadsheet>'
               f'<table:table table:name="S"><table:table-row>{cell_xml}</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>')
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("content.xml", content)
        z.writestr("META-INF/manifest.xml", '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>')
    return out.getvalue()


def _ods_value(tmp_path, cell_xml: str):
    p = _p(tmp_path, "a.ods", _ods(cell_xml))
    _k, doc, _parts = api.window(p)
    return next(c for c in doc["sheets"][0]["cells"] if parse_cell(c["cell"]) == (1, 1))["value"]


def test_ods_keeps_trailing_spaces(tmp_path):
    """LibreOffice が書く終わりの空白を落としていた（LibreOffice・Excel の読みと照らして発見）。"""
    x = '<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>'
    assert _ods_value(tmp_path, x.format("前後 <text:s/>")) == "前後  "
    assert _ods_value(tmp_path, x.format("字 ")) == "字 "
    assert _ods_value(tmp_path, x.format('<text:s text:c="2"/>字')) == "  字"


def test_ods_empty_string_value_with_visible_text(tmp_path):
    """office:string-value が空でも、段落に字があれば、それが見えている値（空のセルにしない）。"""
    x = '<table:table-cell office:value-type="string" table:formula="of:=1/0" office:string-value=""><text:p>#DIV/0!</text:p></table:table-cell>'
    assert _ods_value(tmp_path, x) == "#DIV/0!"
