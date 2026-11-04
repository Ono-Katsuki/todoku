"""ピボットの定義：.xls（公開仕様 [MS-XLS]）・.ods のデータパイロット。計算は再現しない。

.xls の試しは [MS-XLS] の並び（PIVOTCACHEDEFINITION・PIVOTCACHE・PIVOTCORE）のとおりに手で組み立てた記録。"""

from __future__ import annotations

import io
import struct
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))


def rec(t: int, body: bytes = b"") -> bytes:
    return struct.pack("<HH", t, len(body)) + body


def st_no_cch(s: str) -> bytes:
    return b"\x01" + s.encode("utf-16-le")


def u_str(s: str) -> bytes:
    return struct.pack("<H", len(s)) + st_no_cch(s)


class FakeCfb:
    def __init__(self, streams: dict):
        self.s = streams

    def has(self, n: str) -> bool:
        return n in self.s

    def read(self, n: str) -> bytes:
        return self.s[n]

    def streams(self):
        return {k: None for k in self.s}


BOF = rec(0x0809, struct.pack("<HHHH", 0x0600, 0x0005, 0, 0) + bytes(8))
EOF = rec(0x000A)
BOF_SHEET = rec(0x0809, struct.pack("<HHHH", 0x0600, 0x0010, 0, 0) + bytes(8))
BOF_CHART = rec(0x0809, struct.pack("<HHHH", 0x0600, 0x0020, 0, 0) + bytes(8))


def boundsheet(name: str, pos: int) -> bytes:
    return rec(0x0085, struct.pack("<IBB", pos, 0, 0) + bytes([len(name)]) + st_no_cch(name))


def sxview(name: str, ref=(4, 9, 3, 6), cache=0, dims=(4, 1, 1, 1, 1)) -> bytes:
    r1, r2, c1, c2 = ref
    b = struct.pack("<HHHH", r1, r2, c1, c2) + struct.pack("<HHHhH", r1 + 1, r1 + 2, c1 + 1, cache, 0) + struct.pack("<Hh", 8, -1)
    b += struct.pack("<hHHHh", *dims) + struct.pack("<HH", 1, 1) + struct.pack("<HH", 0, 0) + struct.pack("<HH", len(name), 0xFFFF)
    assert len(b) == 44
    return rec(176, b + st_no_cch(name))


def sxvd(axis: int, name: str | None = None) -> bytes:
    b = struct.pack("<HHHH", axis, 1, 1, 0)
    return rec(177, b + (struct.pack("<H", 0xFFFF) if name is None else struct.pack("<H", len(name)) + st_no_cch(name)))


def sxdi(fld: int, func: int, name: str) -> bytes:
    return rec(197, struct.pack("<hhhhhH", fld, func, 0, 0, 0, 0) + struct.pack("<H", len(name)) + st_no_cch(name))


def sxfdb(name: str) -> bytes:
    return rec(199, struct.pack("<HHHHHHH", 0, 0, 0, 0, 0, 0, 0) + u_str(name))


def workbook(view_records: bytes) -> tuple[FakeCfb, bytes]:
    dcon = rec(81, struct.pack("<HHBBH", 0, 44, 0, 3, 3) + b"\x01" + "\x02元表".encode("utf-16-le"))
    globals_ = BOF + rec(213, struct.pack("<H", 0x00A1)) + rec(227, struct.pack("<H", 1)) + dcon + boundsheet("集計", 0) + boundsheet("元表", 1) + EOF
    sheet1 = BOF_SHEET + BOF_CHART + rec(0x1002, bytes(16)) + EOF + view_records + EOF  # シートの中のグラフ（入れ子の BOF）
    sheet2 = BOF_SHEET + EOF
    cache = rec(198, bytes(20)) + sxfdb("区分") + sxfdb("年度") + sxfdb("担当") + sxfdb("金額") + EOF
    stream = globals_ + sheet1 + sheet2
    return FakeCfb({"Workbook": stream, "_SX_DB_CUR/00A1": cache}), stream


CORE = (sxview("集計表") + sxvd(1) + sxvd(2) + sxvd(4, "担当者") + sxvd(8)
        + rec(180, struct.pack("<h", 0)) + rec(180, struct.pack("<h", 1))
        + rec(182, struct.pack("<hhh", 2, 0x7FFD, 7)) + sxdi(3, 2, "平均金額") + rec(0x00F1, bytes(4)))


def test_xls_pivot_definition_is_read_from_the_records():
    from todoku.xls.pivot import views

    cfb, stream = workbook(CORE)
    v = views(cfb, stream)[0]
    assert v["name"] == "集計表" and v["sheet"] == "集計" and v["location"] == (4, 3, 9, 6)
    assert v["source"] == {"ref": (0, 0, 44, 3), "sheet": "元表"} and v["source_type"] == "worksheet"
    assert v["fields"] == ["区分", "年度", "担当者", "金額"]  # 項目の表示名があれば表示名、なければキャッシュの項目の名前
    assert v["row_fields"] == ["区分"] and v["column_fields"] == ["年度"] and v["page_fields"] == ["担当者"]
    assert v["data_fields"] == [{"name": "平均金額", "field": "金額", "function": "average"}]


def test_xls_columns_only_and_data_field_on_an_axis():
    from todoku.xls.pivot import views

    core = (sxview("列だけ", dims=(2, 0, 2, 0, 1)) + sxvd(2) + sxvd(8) + rec(180, struct.pack("<hh", 0, -2)) + sxdi(1, 0, "合計") + rec(0x00F1, bytes(4)))
    cfb, stream = workbook(core)
    v = views(cfb, stream)[0]
    assert v["row_fields"] == [] and v["column_fields"] == ["区分", "（値）"] and v["data_fields"][0]["function"] == "sum"


def test_xls_axis_continued_in_a_continue_record():
    from todoku.xls.pivot import views

    core = (sxview("続き", dims=(3, 3, 0, 0, 0)) + sxvd(1) + sxvd(1) + sxvd(1) + rec(180, struct.pack("<h", 0))
            + rec(0x003C, struct.pack("<hh", 1, 2)) + rec(0x00F1, bytes(4)))
    cfb, stream = workbook(core)
    assert views(cfb, stream)[0]["row_fields"] == ["区分", "年度", "担当"]


def test_two_views_on_one_sheet_stay_apart():
    from todoku.xls.pivot import views

    cfb, stream = workbook(CORE + sxview("2 つ目", dims=(1, 1, 0, 0, 0)) + sxvd(1) + rec(180, struct.pack("<h", 0)))
    got = views(cfb, stream)
    assert [v["name"] for v in got] == ["集計表", "2 つ目"] and got[1]["row_fields"] == ["区分"] and got[1]["data_fields"] == []


def test_extract_processing_reports_the_xls_definition(monkeypatch):
    from todoku.processing import extras
    import todoku.cfb as cfb_mod

    fake, _s = workbook(CORE)
    monkeypatch.setattr(cfb_mod, "CFB", lambda _b: fake)
    got = extras.not_cell_processing(b"x", None, "xls")
    assert got[0]["kind"] == "pivot_table" and got[0]["location"] == "D5:G10" and got[0]["source"]["ref"] == "A1:D45"
    assert got[0]["row_fields"] == ["区分"] and got[0]["expressible_by_cells"] is False


def test_unreadable_xls_pivot_is_said_not_silenced(monkeypatch):
    from todoku.processing import extras

    def boom(_b):
        raise ValueError("x")

    monkeypatch.setattr(extras, "_xls_pivots", boom)
    assert "読めませんでした" in extras.not_cell_processing(b"x", None, "xls")[0]["note"]


def _ods(body: str) -> bytes:
    T = 'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("content.xml", f'<?xml version="1.0"?><office:document-content {T}><office:body><office:spreadsheet>{body}</office:spreadsheet></office:body></office:document-content>')
    return buf.getvalue()


def test_ods_data_pilot_definition():
    from todoku.processing.extras import not_cell_processing

    body = ('<table:data-pilot-tables><table:data-pilot-table table:name="集計" table:target-range-address="S.D1:F9">'
            '<table:source-cell-range table:cell-range-address="S.A1:B9"/>'
            '<table:data-pilot-field table:source-field-name="区分" table:orientation="row"/>'
            '<table:data-pilot-field table:source-field-name="金額" table:orientation="data" table:function="sum"/>'
            '</table:data-pilot-table></table:data-pilot-tables>')
    d = not_cell_processing(_ods(body), None, "ods")[0]
    assert d["name"] == "集計" and d["source"] == {"ref": "S.A1:B9"} and d["location"] == "S.D1:F9"
    assert d["fields"][1] == {"name": "金額", "orientation": "data", "function": "sum"}
