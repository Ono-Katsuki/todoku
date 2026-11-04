"""表計算の窓の「事実」：所の facts・空のマスの罫線の辺と大きさ（mm）・近くの字とその離れ方・理由の鍵（reason_codes）・
マスの中の所の変わらない id（「@」＋字の位置）と、仕様どおりの色（テーマ・明るさの補正・色番号）・.xls の既定の列の幅。
どれも字の意味（語）を見ない。いろいろな言語の字で同じに動くことを確かめる。"""

from __future__ import annotations

import io
import struct
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from todoku import chars  # noqa: E402
from todoku.xlsx import window  # noqa: E402
from todoku.xlsx.workbook import Workbook, apply_tint, theme_colors  # noqa: E402

openpyxl = pytest.importorskip("openpyxl")
from openpyxl.styles import Border, Color, PatternFill, Side  # noqa: E402

THIN = Side(style="thin")


def _blob(build) -> bytes:
    wb = openpyxl.Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _fields(blob: bytes) -> dict:
    return {f["id"]: f for sh in window.build(Workbook(blob, "t.xlsx"))["sheets"] for f in sh["fields"]}


# ---------------------------------------------------------------- 文字の中の所の事実（どの言語でも同じ）
@pytest.mark.parametrize("text,before", [
    ("令和　　年　　月　　日", "令和"), ("　　년　　월　　일", ""), ("中華民國　　年　　月　　日", "中華民國"),
    ("Name:    Date", "Name:"), ("الاسم:    التاريخ", "الاسم:"), ("नाम    तारीख", "नाम"),
])
def test_place_facts_are_verbatim_and_measured(text, before):
    got = window.cell_blanks(text)
    assert got, text
    f = got[0].facts
    assert f["before"].endswith(before) or not before
    assert f["printed"] and set(f["printed"]) <= {" ", "　"}
    assert f["width_em"] == chars.width_em(f["printed"]) and f["len"] == len(f["printed"])


def test_gap_interior_keeps_one_space_on_each_printed_side_or_takes_the_whole_gap():
    t = "Name      Date"
    a, b, edges = window.gap_interior(t, 4, 10)
    assert (a, b, edges) == (5, 9, (True, True))
    assert window.gap_interior("年　　月", 1, 3) == (1, 3, (False, False))  # 1 字ぶんも残らないときは空き全体
    assert window.gap_interior(t, 4, 10, pad=(True, False)) == (5, 10, (True, False))  # 呼び手が決めて渡せる
    assert window.gap_interior("（　　　）", 1, 4) == (1, 4, (False, False))  # かっこの側には残さない


def test_old_hiding_rules_are_gone_but_chain_is_a_fact():
    # 1 字ずつ同じ空白でつないだ並び：字数で隠さない（heading_gap はいつも None）。並びの字数は facts の chain
    t = "就　労　定　着　支　援"
    assert window.heading_gap(t, 1, 2) is None and window.chain_head("年　　月　　日") is None
    places = window.blank_places(t)
    assert len(places) == 5 and {k for _a, _b, k in places} == {"spaced"}
    assert places[0].facts["chain"] == {"len": 6, "text": t}
    assert not window.weak_space("年　月", 1, 2)


def test_placeholder_runs_are_separate_places():
    # 「○○県○○市」：あいだの字の数で 1 つにまとめない（以前は 3 字までの字でつないでいた）
    got = window.agent_blanks("○○県○○市")
    assert [(a, b, k) for a, b, k in got if k == "placeholder"] == [(0, 2, "placeholder"), (3, 5, "placeholder")]
    assert got[1].facts["before"] == "○○県" and got[1].facts["after"] == "市"


@pytest.mark.parametrize("text,is_run", [("...", True), ("…", True), ("‥‥", True), ("..", False), ("‥", False),
                                         ("1.5", False), ("———", True), ("1-2", False)])
def test_leader_and_dash_length_follow_the_ellipsis_decomposition(text, is_run):
    # 点線の長さの基準は「…」（U+2026）の互換分解の点の数（3）。ダッシュも同じ長さ
    assert bool(window.blank_places(text)) == is_run, text


# ---------------------------------------------------------------- マスの中の所の、変わらない id
def test_span_ids_are_character_offsets_and_fill_by_offset():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "令和　　年　　月　　日"

    blob = _blob(build)
    f = _fields(blob)["S!A1"]
    ids = [s["id"] for s in f["spans"]]
    assert ids == [window.span_id("S!A1", s["at"]) for s in f["spans"]] and ids[0] == "S!A1@2"
    assert window.parse_span_id(ids[1]) == ("S!A1", f["spans"][1]["at"])
    assert window.parse_span_id("S!A1") is None
    text = "令和　　年　　月　　日"
    spans = [(s["at"], s["at"] + len(s["printed"])) for s in f["spans"]]
    assert window.fill_blanks(text, {"@2": "8", spans[2][0]: "1"}, spans) == "令和8年　　月1日"
    with pytest.raises(ValueError):
        window.fill_blanks(text, {"@3": "x"}, spans)


# ---------------------------------------------------------------- 空のマスの事実と、理由の鍵
@pytest.mark.parametrize("label", ["성명", "氏名", "Nom", "الاسم", "नाम"])
def test_bordered_cell_is_agent_with_side_size_and_distance_facts(label):
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = label
        ws["D1"].border = Border(bottom=THIN)  # 2 列あけた下線の欄
        ws["A5"] = label
        ws["A8"].fill = PatternFill("solid", fgColor="FFFF00")  # 2 行あけた塗りの欄

    f = _fields(_blob(build))
    d1, a8 = f["S!D1"], f["S!A8"]
    assert d1["agent"] and d1["reason_codes"] == ["blank_box"] and d1["evidence"]["sides"]["bottom"]
    assert d1["evidence"]["left"]["text"] == label and d1["evidence"]["left"]["cols"] == 2
    assert d1["evidence"]["left"]["mm"] > 20 and d1["evidence"]["size_mm"]["w"] > 0
    assert a8["evidence"]["fill"] == "FFFF00" and a8["evidence"]["above"]["rows"] == 3
    assert a8["evidence"]["above"]["mm"] > 10


def test_declared_inputs_carry_reason_codes():
    from openpyxl.worksheet.datavalidation import DataValidation

    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "区分"
        dv = DataValidation(type="list", formula1='"甲,乙"')
        ws.add_data_validation(dv)
        dv.add("B1")
        ws["A2"] = "数"
        ws["B3"] = "=B2*2"

    f = _fields(_blob(build))
    assert f["S!B1"]["reason_codes"] == ["validation"]
    # 数式の読むセルは宣言ではない：「？」の所で、読む数式は事実（referenced_by）
    assert f["S!B2"]["agent"] and f["S!B2"]["reason_codes"] == ["formula_ref"]
    assert f["S!B2"]["evidence"]["referenced_by"]["cells"] == ["S!B3"]
    # 選択肢の数・長さからは、欄の種類（○を付ける欄か）を決めない：見出しは左の字
    assert f["S!B1"]["options"] == ["甲", "乙"] and f["S!B1"]["label"] == "区分"


# ---------------------------------------------------------------- 色（ECMA-376）
def test_tint_follows_the_spec():
    assert apply_tint("FFFFFF", -0.05) == "F2F2F2" and apply_tint("FFFFFF", -0.15) == "D9D9D9"
    assert apply_tint("000000", 0.499984740745262) == "7F7F7F"  # Excel の「黒、テキスト 1、白 + 基本色 50%」
    assert apply_tint("4472C4", 0) == "4472C4"


def test_theme_tint_and_indexed_fills_are_resolved():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "x"
        ws["B1"].fill = PatternFill("solid", fgColor=Color(theme=0, tint=-0.15))
        ws["C1"].fill = PatternFill("solid", fgColor=Color(theme=4, tint=0.4))
        ws["D1"].fill = PatternFill("solid", fgColor=Color(indexed=10))

    wb = Workbook(_blob(build), "t.xlsx")
    sh = wb.sheet("S")
    theme = theme_colors(wb.xml("xl/theme/theme1.xml"))
    assert theme is not None and len(theme) == 12
    fills = [wb.style(sh.style_of(1, c)).fill_color for c in (2, 3, 4)]
    assert fills == [apply_tint(theme[0], -0.15), apply_tint(theme[4], 0.4), "FF0000"]


# ---------------------------------------------------------------- .xls の既定の列の幅（MS-XLS）
def test_xls_default_column_width_records():
    from test_xls_controls import cfb, rec, workbook

    from todoku.xls.book import XlsWorkbook

    data = workbook([{"name": "S", "cells": {(0, 0): "a"}, "defcolwidth": 10,
                      "records": [rec(0x0099, struct.pack("<H", 12 * 256 + 128))]}])
    sh = XlsWorkbook(cfb("Workbook", data)).sheet("S")
    assert sh.base_col_width == 10 and sh.default_col_width == 12.5


# ---------------------------------------------------------------- OpenDocument の長さ
def test_ods_length_parse_returns_none_when_unreadable():
    from todoku.odf.sheet import _pt

    assert _pt("1in") == 72 and _pt("2.54cm") == pytest.approx(72, rel=1e-3) and _pt("auto") is None and _pt("") is None


# ---------------------------------------------------------------- 見出しは字のまま・上に続く見出しと区分は事実（段の数・行の数で打ち切らない）
def _protected(ws, cells):
    from openpyxl.styles import Protection

    ws.protection.sheet = True
    for ref in cells:
        ws[ref].protection = Protection(locked=False)


@pytest.mark.parametrize("label,key", [("氏　　名", "氏名"), ("성　　명", "성명"), ("姓　　名", "姓名"),
                                       ("N a m e", "N a m e"), ("الاسم", "الاسم")])
def test_label_is_verbatim_and_compacted_copy_is_only_a_key(label, key):
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = label
        _protected(ws, ["B1"])

    f = _fields(_blob(build))["S!B1"]
    assert f["label"] == label  # 字間の空白を詰めない
    assert f.get("label_key", f["label"]) == key  # 詰めた写しは、探すための鍵として別に
    assert f["near"]["left"]["text"] == label and f["near"]["left"]["key"] == key


@pytest.mark.parametrize("parent,a,b", [("電話番号", "自宅", "携帯"), ("Phone", "Home", "Mobile"), ("هاتف", "منزل", "جوال"),
                                        ("전화", "자택", "휴대")])
def test_two_header_rows_are_reported_verbatim_with_spans_and_distance(parent, a, b):
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["B1"] = parent
        ws.merge_cells("B1:C1")
        ws["B2"], ws["C2"] = a, b
        _protected(ws, ["B3", "C3"])

    f = _fields(_blob(build))
    b3 = f["S!B3"]
    assert b3["label"] == a  # 見出しは、上のいちばん近い字だけ（上の段を幅の目安で選んで付けない）
    hd = b3["near"]["header"]
    assert [h["text"] for h in hd] == [a, parent]
    assert hd[0]["rows"] == 1 and hd[0]["cols"] == 1 and hd[1]["rows"] == 2 and hd[1]["cols"] == 2
    assert hd[1]["cell"] == "B1:C1" and hd[1]["mm"] > 0 and hd[0]["mm"] == 0


@pytest.mark.parametrize("head,rows", [("Name", ["x1", "x2", "x3"]), ("氏名", ["父", "母", "子"]), ("이름", ["가", "나", "다"])])
def test_repeated_rows_do_not_inherit_the_heading_above(head, rows):
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["B1"] = head
        for i, t in enumerate(rows, 2):
            ws.cell(i, 1, t)
        _protected(ws, ["B2", "B3", "B4"])

    b3 = _fields(_blob(build))["S!B3"]
    # 同じ形の行が続くだけ（ファイルに表の宣言はない）：見出しは前（左）の字。上の見出しは near の header に事実として
    assert b3["label"] == rows[1] and b3["label_method"] == "左の見出し"
    assert b3["near"]["header"][0]["text"] == head and b3["near"]["header"][0]["rows"] == 2


def test_section_heading_is_found_however_far_above_and_reported_with_distance():
    def build(wb):
        ws = wb.active
        ws.title = "S"
        ws["A1"] = "区分"
        for r in range(1, 13):
            ws.cell(r, 2, f"項目{r}")
        _protected(ws, [f"C{r}" for r in range(1, 13)])

    c12 = _fields(_blob(build))["S!C12"]
    # 区分の見出しは、同じ列割りの字の行が続くあいだ引き継ぐ（以前は 8 行で打ち切っていた）。何行上かは事実
    assert c12["label"] == "区分／項目12"
    sec = c12["near"]["section"]
    assert [(s["text"], s["cell"], s["rows"]) for s in sec] == [("区分", "A1", 11)] and sec[0]["mm"] > 40
