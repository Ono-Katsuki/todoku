"""セルの読む向き（.xlsx の readingOrder、.xls の iReadOrder、.xlsb の iReadingOrder、.ods の style:writing-mode）。

寄せ方の決まっていないセル（標準）でも、読む向きが右から左のセルの字は右に寄るので、はみ出す字は画面の左へ伸びる。
右から左へ並ぶシート（列 A が右端）では、画面の左が、列の番号の増える向き。
fixtures/reading_order.xls・.ods は、同じ .xlsx（B1 読む向き 2・B3 読む向き 1・B5 読む向き 0・B7 左寄せで読む向き 2、
A の列に「氏名」）を LibreOffice 26.8 が書き出したファイル。
"""
from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path

import pytest
import xlsb_maker as xm

from todoku import api
from todoku.book import open_book
from todoku.xlsx import cellcheck

FIX = Path(__file__).parent / "fixtures"
NAME = "架空 太郎 " * 8


def _xlsx(tmp_path, rtl_sheet: bool = False, c_text: bool = False) -> str:
    openpyxl = pytest.importorskip("openpyxl")
    from openpyxl.styles import Alignment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "S"
    ws.sheet_view.rightToLeft = rtl_sheet
    for r, al in ((1, Alignment(readingOrder=2)), (3, Alignment(readingOrder=1)), (5, Alignment(readingOrder=0)),
                  (7, Alignment(horizontal="left", readingOrder=2)), (9, Alignment(horizontal="right"))):
        ws.cell(r, 1, "氏名")
        ws.cell(r, 2).alignment = al
        if c_text:
            ws.cell(r, 3, "備考")
    ws.cell(11, 2, "abc").alignment = Alignment(readingOrder=2)
    p = tmp_path / "a.xlsx"
    wb.save(p)
    return str(p)


def _fit(src: str, cell: str, out) -> list[str]:
    r = api.fill(src, {f"S!{cell}": NAME}, str(out), force=True)
    assert r["verified"], r["warnings"]
    return [f["message"] for f in r.get("fit", [])]


def test_formats_carry_the_reading_order(tmp_path):
    wb = open_book(Path(_xlsx(tmp_path)).read_bytes(), "a.xlsx")
    sh = wb.sheet("S")
    fm = cellcheck._formats(wb)
    assert [fm[sh.style_of(r, 2)].reading for r in (1, 3, 5, 7, 9)] == [2, 1, 0, 2, 0]
    for name in ("reading_order.xls",):
        wb = open_book((FIX / name).read_bytes(), name)
        sh = wb.sheet("S")
        fm = cellcheck._formats(wb)
        assert [(fm[sh.style_of(r, 2)].halign, fm[sh.style_of(r, 2)].reading) for r in (1, 3, 5, 7)] == [
            ("general", 2), ("general", 1), ("general", 0), ("left", 2)]


def test_xlsx_general_cell_read_right_to_left_overflows_to_the_left(tmp_path):
    src = _xlsx(tmp_path)
    m = _fit(src, "B1", tmp_path / "o1.xlsx")
    assert m and "隣の A1 の字で切れて見える見込み" in m[0] and "かもしれません" not in m[0]
    # 読む向きが 1・0 の標準のセルと、左寄せのセル（読む向き 2）は、右の空の C へ伸びる
    for i, cell in enumerate(("B3", "B5", "B7")):
        assert _fit(src, cell, tmp_path / f"o{i + 2}.xlsx") == []
    assert "隣の A9 の字で切れて見える見込み" in _fit(src, "B9", tmp_path / "o5.xlsx")[0]


def test_right_to_left_sheet_mirrors_the_column_direction(tmp_path):
    # 列 A が右端：左寄せは画面の右（A の側）へ、右寄せと、読む向き 2 の標準のセルは画面の左（C の側）へ
    src = _xlsx(tmp_path, rtl_sheet=True)
    assert "隣の A7 の字で切れて見える見込み" in _fit(src, "B7", tmp_path / "o1.xlsx")[0]
    assert _fit(src, "B9", tmp_path / "o2.xlsx") == [] and _fit(src, "B1", tmp_path / "o3.xlsx") == []
    src = _xlsx(tmp_path, rtl_sheet=True, c_text=True)
    assert "隣の C9 の字で切れて見える見込み" in _fit(src, "B9", tmp_path / "o4.xlsx")[0]
    assert "隣の C1 の字で切れて見える見込み" in _fit(src, "B1", tmp_path / "o5.xlsx")[0]
    # 寄せ方の決まっていない、読む向きが 2 でないセルは、前と同じに決めない
    m = _fit(src, "B5", tmp_path / "o6.xlsx")
    assert m and "かもしれません" in m[0] and "はみ出す向きを決めていない" in m[0]


@pytest.mark.parametrize("ext", ["xls", "ods"])
def test_xls_and_ods_written_by_libreoffice(tmp_path, ext):
    src = str(FIX / f"reading_order.{ext}")
    m = _fit(src, "B1", tmp_path / f"o1.{ext}")
    assert m and "隣の A1 の字で切れて見える見込み" in m[0]
    for i, cell in enumerate(("B3", "B5", "B7")):
        assert _fit(src, cell, tmp_path / f"o{i + 2}.{ext}") == []


def test_xlsb_reading_order_bits(tmp_path):
    blob = xm.package({"S": xm.sheet(xm.row(0) + xm.st(0, "氏名") + xm.blank(1, 1))})
    old = xm.rec(47, struct.pack("<HHHHHBBHH", 0, 0, 0, 0, 0, 0, 0, 0, 0))  # xf 1：ロックしないセル
    new = xm.rec(47, struct.pack("<HHHHHBBHH", 0, 0, 0, 0, 0, 0, 0, 2 << 10, 0))  # iReadingOrder（10〜11 ビット目）＝ 2
    buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as zin, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == "xl/styles.bin":
                assert data.count(old) == 1
                data = data.replace(old, new)
            zout.writestr(info, data)
    src = tmp_path / "a.xlsb"
    src.write_bytes(buf.getvalue())
    wb = open_book(src.read_bytes(), "a.xlsb")
    assert cellcheck._formats(wb)[wb.sheet("S").style_of(1, 2)].reading == 2
    m = _fit(str(src), "B1", tmp_path / "o.xlsb")
    assert m and "隣の A1 の字で切れて見える見込み" in m[0]


def test_picture_of_the_sheet_draws_right_to_left_general_text_at_the_right(tmp_path):
    from todoku.sheetpic import cell_text

    wb = open_book(Path(_xlsx(tmp_path)).read_bytes(), "a.xlsx")
    sh = wb.sheet("S")
    fm = cellcheck._formats(wb)
    assert cell_text(wb, sh, fm, 11, 2).halign == "right"  # 「abc」・標準・読む向き 2
    assert cell_text(wb, sh, fm, 1, 1).halign == "left"
