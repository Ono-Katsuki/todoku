""".xlsb（MS-XLSB）：手で作ったファイルを、窓・数式・欄・計算の説明・検査・段階・上書き（証明つき）で確かめる。"""

from __future__ import annotations

import struct
import zipfile

import pytest
import xlsb_maker as m

from todoku import api
from todoku.book import kind_of, open_book
from todoku.xlsb.book import XlsbError, XlsbWorkbook, records


def _write(tmp_path, blob=None, name="a.xlsb"):
    p = tmp_path / name
    p.write_bytes(blob or m.sample())
    return p


def test_kind_and_cells():
    blob = m.sample()
    assert kind_of(blob) == "xlsb"
    wb = open_book(blob, "a.xlsb")
    sh = wb.sheet("申請")
    assert sh.cells[(1, 1)].value == "品名" and sh.cells[(1, 3)].value == "直接の字"
    assert sh.cells[(2, 2)].value == 1234.5 and sh.cells[(2, 3)].value == 7.0
    assert sh.cells[(5, 3)].value is True and sh.cells[(6, 3)].value.code == "#DIV/0!"
    assert sh.merges == [(1, 3, 1, 4)] and 7 in sh.hidden_rows


def test_formulas_become_text_and_match_saved_values():
    wb = open_book(m.sample(), "a.xlsb")
    sh = wb.sheet("申請")
    assert sh.cells[(3, 2)].formula == "B2*C2"
    assert sh.cells[(3, 3)].formula == "SUM(B2:C2)"
    assert sh.cells[(4, 2)].formula == "税率"
    assert sh.cells[(4, 3)].formula == 'A2&"円"'
    assert [(d.name, d.text) for d in wb.names] == [("税率", "基準!$A$1")]
    from todoku.xlsx.calc import verify

    assert verify(wb).summary()["mismatch"] == 0


def test_styles_validation():
    wb = open_book(m.sample(), "a.xlsb")
    sh = wb.sheet("申請")
    assert wb.style(sh.cells[(2, 2)].style).num_fmt == '#,##0"円"'
    assert wb.style(sh.cells[(5, 2)].style).locked is False
    v = sh.validations[0]
    assert (v.sqref_text, v.type, v.formula1) == ("B5:B6", "list", '"甲,乙"')
    assert (v.prompt_title, v.prompt, v.error_title, v.error) == ("入力の題", "入力の説明", "題", "選んでください")
    wb2 = XlsbWorkbook(m.package({"S": m.sheet(m.row(0) + m.blank(0, 3))}), "b.xlsb")
    st = wb2.style(wb2.sheets[0].cells[(1, 1)].style)
    assert (st.font_color, st.fill_color, st.borders) == ("FF0000", "FFFF00", 8)


def test_shared_and_array_formulas():
    # 共有数式：B1 が本体（A1*2 を相対で）、B2 はそれを指す
    refn = bytes([0x2C]) + struct.pack("<IH", 0, (-1 & 0x3FFF) | 0xC000)
    body = refn + m.p_int(2) + m.MUL
    exp = bytes([0x01]) + struct.pack("<I", 0)
    col = struct.pack("<I", 1)
    rows = (m.row(0) + m.real(0, 3) + m.fnum(1, 6, exp, col) + m.rec(427, m.rfx(0, 1, 1, 1) + struct.pack("<I", len(body)) + body + struct.pack("<I", 0))
            + m.row(1) + m.real(0, 5) + m.fnum(1, 10, exp, col)
            + m.row(2) + m.fnum(0, 8, bytes([0x01]) + struct.pack("<I", 2), struct.pack("<I", 0))
            + m.rec(426, m.rfx(2, 2, 0, 0) + b"\x00" + struct.pack("<I", len(m.p_area(0, 1, 0, 0) + bytes([0x19, 0x10, 0, 0])))
                    + m.p_area(0, 1, 0, 0) + bytes([0x19, 0x10, 0, 0]) + struct.pack("<I", 0)))
    wb = XlsbWorkbook(m.package({"S": m.sheet(rows)}), "s.xlsb")
    sh = wb.sheets[0]
    assert sh.cells[(1, 2)].formula == "A1*2" and sh.cells[(2, 2)].formula.strip("()") == "A2*2"
    assert sh.cells[(3, 1)].formula == "SUM(A1:A2)" and sh.cells[(3, 1)].formula_kind == "array"
    from todoku.xlsx.calc import verify

    assert verify(wb).summary()["mismatch"] == 0


def test_window_api_fields_and_calc_model(tmp_path):
    p = _write(tmp_path)
    s = api.summary(str(p))
    assert s["kind"] == "xlsb" and s["sheets"] == ["申請", "基準"]
    k, doc, parts = api.window(str(p))
    text = "\n".join(t for _n, t in parts)
    assert "直接の字" in text and "1,235円" in text and "グラフ" not in text
    f = {x["id"]: x for x in api.fields(str(p))}
    b5 = next(v for i, v in f.items() if i.endswith("B5"))
    assert b5["options"] == ["甲", "乙"] and not b5.get("readonly")
    model = api.calc_model(str(p), "申請!B3")
    assert "B2" in str(model) and "C2" in str(model)
    t = api.read_table(str(p), "単価", "申請")
    assert t["rows"] >= 1
    r = api.calculate(str(p), {"申請!B2": 2000})
    assert "14000" in str(r) or "14,000" in str(r)


def test_audit_counts_and_grade(tmp_path):
    p = _write(tmp_path)
    au = api.audit(str(p))
    assert au["ok"], [c for c in au["checks"] if not c["ok"]]
    g = api.grade(str(p))
    assert g["write"]["grade"] == "C" and g["grade"] in "AB"
    assert any(".xlsb" in r["title"] for r in g["write"]["reasons"])


def test_overwrite_same_size_with_proof_and_recalculation(tmp_path):
    p = _write(tmp_path)
    out = tmp_path / "b.xlsb"
    r = api.fill(str(p), {"申請!B2": 2000}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert r["recalculated"] == {"申請!B3": 14000.0, "申請!C3": 2007.0}
    wb = open_book(out.read_bytes(), "b.xlsb")
    assert wb.sheet("申請").cells[(2, 2)].value == 2000.0 and wb.sheet("申請").cells[(3, 2)].value == 14000.0
    a, b = zipfile.ZipFile(p), zipfile.ZipFile(out)
    assert a.namelist() == b.namelist()
    for n in a.namelist():
        if n != "xl/worksheets/sheet1.bin":
            assert a.read(n) == b.read(n)
    assert len(a.read("xl/worksheets/sheet1.bin")) == len(b.read("xl/worksheets/sheet1.bin"))


def test_refuses_what_would_change_size(tmp_path):
    p = _write(tmp_path)
    out = tmp_path / "b.xlsb"
    with pytest.raises(api.ApiError, match="同じ字数"):
        api.fill(str(p), {"申請!C1": "もっと長い字にする"}, str(out))
    with pytest.raises(api.ApiError, match="共有"):
        api.fill(str(p), {"申請!A1": "品物"}, str(out))
    with pytest.raises(api.ApiError, match="数式"):
        api.fill(str(p), {"申請!B3": 1}, str(out))
    with pytest.raises(api.ApiError, match="表せません"):
        api.fill(str(p), {"申請!C2": 1.5}, str(out))
    assert not out.exists()


def test_same_length_text_and_broken_records(tmp_path):
    p = _write(tmp_path)
    out = tmp_path / "b.xlsb"
    r = api.fill(str(p), {"申請!C1": "別の字の字"[:4]}, str(out))
    assert r["proof_ok"] and open_book(out.read_bytes(), "b").sheet("申請").cells[(1, 3)].value == "別の字の"
    with pytest.raises(XlsbError):
        list(records(b"\x05\x7f\x00"))


def test_unsupported_formula_keeps_saved_value_and_grades_c(tmp_path):
    unknown = bytes([0x18, 0x19]) + bytes(12)  # 表の構造化参照（PtgList）：文字列にしない
    p = _write(tmp_path, m.package({"S": m.sheet(m.row(0) + m.fnum(0, 42, unknown))}))
    wb = open_book(p.read_bytes(), "a.xlsb")
    c = wb.sheets[0].cells[(1, 1)]
    assert c.value == 42 and "未対応" in c.parse_error
    assert api.grade(str(p))["grade"] == "C"


def test_drawing_is_reported_unread(tmp_path):
    p = _write(tmp_path, m.package({"S": m.sheet(m.row(0) + m.real(0, 1), tail=m.rec(550, m.ws("rId1")))}))
    k, doc, parts = api.window(str(p))
    assert any("図形" in u for u in doc["unread"])
    g = api.grade(str(p))
    assert g["grade"] == "C" and any(r["title"] == "読んでいない部分" for r in g["reasons"])
