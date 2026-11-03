"""段階の検証（docs/段階の検証と比較.md）で見つかった不具合の再現：表計算の行のまとめ・非表示の中身・.xls の位置のない図形・
画像だけのシート・CSV の NUL・文字コード。どれも、この中で作る小さなファイルで確かめる（取ってきた書類は使わない）。"""

from __future__ import annotations

import io
import struct

import pytest

from todoku import api
from todoku.datafile.textenc import decode
from todoku.grade import grade_blob
from todoku.xlsx import window
from todoku.xlsx.audit import check_markdown
from todoku.xlsx.workbook import Workbook

openpyxl = pytest.importorskip("openpyxl")


def book(build) -> Workbook:
    wb = openpyxl.Workbook()
    build(wb)
    buf = io.BytesIO()
    wb.save(buf)
    return Workbook(buf.getvalue(), "test.xlsx")


def md_of(doc: dict) -> str:
    return "\n".join(t for _, t in window.markdown_parts(doc))


# ------------------------------------------------------------------ 1. 値のある行は、まとめに消えない
def test_data_rows_with_numbers_are_not_collapsed():
    """数だけが違うデータの行（以前は数を「#」と見て、同じ形の空欄の行として1行にまとめ、値が窓から消えた）。"""
    def build(wb):
        ws = wb.active
        ws.title = "表"
        ws.append(["地域", "人口", "世帯"])
        for i in range(40):
            ws.append(["地域", 1000 + i * 7, 300 + i])  # 文字の列は同じ、数の列だけが違う

    doc = window.build(book(build))
    sh = doc["sheets"][0]
    assert not any("rows" in x for x in sh["reading"]), [x for x in sh["reading"] if "rows" in x]
    md = md_of(doc)
    for i in range(40):
        assert f"{1000 + i * 7:,}" in md or str(1000 + i * 7) in md
    assert check_markdown(doc).ok


def test_blank_form_rows_are_summarized_with_every_differing_value():
    """番号だけが違い、入力欄が空の行（記入する表）は、まとめてよい。ただし番号は、まとめの行に番地つきで全部書く。"""
    def build(wb):
        ws = wb.active
        ws.title = "名簿"
        ws["A1"] = "番号"
        ws["B1"] = "氏名"
        for i in range(20):
            ws.cell(2 + i, 1, i + 1)
        ws.protection.sheet = True
        for i in range(20):
            ws.cell(2 + i, 2).protection = openpyxl.styles.Protection(locked=False)

    doc = window.build(book(build))
    sh = doc["sheets"][0]
    runs = [x for x in sh["reading"] if "rows" in x]
    assert len(runs) == 1 and runs[0]["rows"] == "3-21"
    assert [d["cell"] for d in runs[0]["differs"]] == [f"A{r}" for r in range(3, 22)]
    assert "A21＝20" in runs[0]["text"] and "字が違うのは A 列だけ" in runs[0]["text"]
    assert check_markdown(doc).ok


def test_markdown_check_catches_values_missing_from_the_reading():
    """監査の「読み落とし（Markdown の窓）」：読み順から消えた値・まとめで消えた値を見つける。"""
    def build(wb):
        ws = wb.active
        ws.append(["名前", "値"])
        ws.append(["a", 1])
        ws.append(["b", 2])

    doc = window.build(book(build))
    assert check_markdown(doc).ok
    sh = doc["sheets"][0]
    sh["reading"] = [x for x in sh["reading"] if x.get("row") != 3]
    ch = check_markdown(doc)
    assert not ch.ok and any("A3" in x for x in ch.items) and any("B3" in x for x in ch.items)
    sh["reading"].append({"rows": "3-3", "same_as": "2-2", "text": "2行目と同じ形の行が 1 行続きます"})
    ch = check_markdown(doc)
    assert not ch.ok and any("まとめの中" in x for x in ch.items)  # 2行目と字が違うのに、まとめた


def test_audit_reports_markdown_completeness():
    from todoku.xlsx.audit import run

    def build(wb):
        ws = wb.active
        for i in range(10):
            ws.append([f"項目{i}", i * 3])

    buf = io.BytesIO()
    wb = openpyxl.Workbook()
    build(wb)
    wb.save(buf)
    au = run(buf.getvalue(), "a.xlsx")
    ch = next(c for c in au.checks if c.name == "読み落とし（Markdown の窓）")
    assert ch.ok and "20/20" in ch.detail


# ------------------------------------------------------------------ 4. 非表示の中身は、印をつけて窓に出す
def test_hidden_rows_columns_sheets_and_merged_values_are_shown_marked():
    def build(wb):
        ws = wb.active
        ws.title = "様式"
        ws["A1"] = "申請書"
        ws["A2"] = "隠れた行の注意"
        ws.row_dimensions[2].hidden = True
        ws["C1"] = "隠れた列の字"
        ws.column_dimensions["C"].hidden = True
        ws["A4"] = "結合の左上"
        ws["B4"] = "結合の下の字"
        ws.merge_cells("A4:B4")
        lst = wb.create_sheet("選択肢")
        for i, v in enumerate(["東京都", "大阪府", "北海道"], 1):
            lst.cell(i, 1, v)
        lst.sheet_state = "hidden"

    wb = book(build)
    from todoku.xlsx.workbook import Cell

    # openpyxl は結合の下の値を保存しないので、読んだあとに足す（Excel のファイルには残っていることがある）
    wb.sheets[0].cells[(4, 2)] = Cell("様式", 4, 2, value="結合の下の字")
    doc = window.build(wb)
    md = md_of(doc)
    assert "〔非表示〕〔A2〕隠れた行の注意" in md
    assert "〔非表示〕〔C1〕隠れた列の字" in md
    assert "〔結合の下に隠れた字〕〔B4〕結合の下の字" in md
    assert "非表示のシート" in md and "東京都" in md and "北海道" in md
    # 非表示の字は、入力欄にも、見出しの推定にも使わない（配置には入れない）
    sh = doc["sheets"][0]
    assert not any(f.get("cell") in ("A2", "C1", "B4") for f in sh["fields"])
    assert check_markdown(doc).ok


# ------------------------------------------------------------------ 2. .xls：位置のない図形
def test_xls_shape_without_anchor_opens_and_keeps_text():
    """OfficeArt の図形に位置（0xF010 ClientAnchor）がない（グループの中の子の図形は 0xF00F だけ）と、以前は KeyError で
    ブック全体が開けなかった。位置は決めずに、字は残す。"""
    from test_xls_controls import Obj, cfb, workbook

    sheets = [{"name": "表", "cells": {(0, 0): "見出し", (1, 0): 12.0},
               "shapes": [{"obj": Obj(0x06, 1), "anchor": (2, 2, 4, 4), "name": "テキスト 1", "text": "グループの中の字"}]}]
    stream = workbook(sheets)
    anchor = struct.pack("<HHI", 0, 0xF010, 18)
    assert stream.count(anchor) == 1
    stream = stream.replace(anchor, struct.pack("<HHI", 0, 0xF00F, 18))  # 同じ長さの、子の図形の位置（読まない種類）にする
    blob = cfb("Workbook", stream)
    from todoku.xls.book import XlsWorkbook

    wb = XlsWorkbook(blob, "a.xls")
    shapes = wb.sheets[0].shapes
    assert [(s.text, s.anchor) for s in shapes] == [("グループの中の字", None)]
    doc = window.build(wb)
    md = md_of(doc)
    assert "グループの中の字" in md and "位置がファイルにない" in md


# ------------------------------------------------------------------ 3. 画像だけのシート
def test_image_only_sheet_ignores_print_area_and_charts():
    from todoku.xlsx.workbook import Shape

    def build(wb):
        ws = wb.active
        ws.title = "データ"
        for i in range(5):
            ws.append([f"行{i}", i])
        ws.print_area = "H20:J30"  # 使っている範囲の外の印刷範囲
        wb.create_sheet("グラフ")
        wb.create_sheet("写真")

    wb = book(build)
    data, chart, photo = wb.sheets
    data.shapes.append(Shape("picture", "図 1", "", "", (1, 5)))
    chart.shapes.append(Shape("chart", "グラフ 1", "", "", (1, 1)))
    photo.shapes.append(Shape("picture", "図 2", "", "", (1, 1)))
    assert window.image_only_sheets(wb) == ["写真"]


# ------------------------------------------------------------------ 5. CSV の NUL
def test_utf8_csv_with_nul_in_values_is_data():
    rows = ["名前;値"] + [f"Čaj {i};{i}\x00\x00" for i in range(30)]
    blob = ("\r\n".join(rows) + "\r\n").encode("utf-8")
    g = grade_blob(blob, "a.csv")
    assert g.grade == "S", g.reasons
    assert any(r.title == "制御文字" and "60" in r.detail for r in g.reasons)
    assert decode(blob).encoding == "utf-8"


def test_ascii_bytes_with_nul_are_still_binary():
    assert grade_blob(b"\x00\x01\x02abc", "a.csv").grade == "F"


# ------------------------------------------------------------------ 6. 文字コード
NOTE = "數據用戶須知\r\n本刊物所載的統計數字，均為臨時數字。\r\n"


def test_big5_is_not_read_as_latin1():
    d = decode(NOTE.encode("big5"))
    assert d.encoding != "latin-1" and d.basis != "none"
    names = {d.encoding} | {c["encoding"] for c in d.candidates}
    assert "cp950" in names
    if d.encoding != "cp950":
        # 決まった順の前の文字コード（日本の文字コード）でも読めるなら、決めずに候補として示す（黙って選ばない）
        assert d.undetermined and any(c["encoding"] == "cp950" and "數據用戶須知" in c["sample"] for c in d.candidates)
    else:
        assert d.text == NOTE


def test_ambiguous_encoding_is_handed_to_the_agent(tmp_path):
    """Shift_JIS としても GBK としても仕様どおりに読めるとき：決めずに候補を示し、段階は C（エージェントが決める）。"""
    text = "国民の祝日,名称\r\n1955/1/1,元日\r\n"
    blob = text.encode("cp932")
    d = decode(blob)
    assert d.encoding == "cp932" and d.text == text
    assert [c["encoding"] for c in d.candidates][:1] == ["gbk"]
    g = grade_blob(blob, "a.csv")
    assert g.grade == "C" and any(r.title == "文字コードを規則では決められない" for r in g.reasons)
    p = tmp_path / "a.csv"
    p.write_bytes(blob)
    _k, _doc, parts = api.window(str(p))
    md = parts[0][1]
    assert "文字コードを規則では決められません" in md and "GBK" in md


def test_shift_jis_decided_when_only_it_fits():
    """半角カナの後ろに「,」が続く（GBK・Big5・EUC では続くバイトになれない）：Shift_JIS だけが読める。"""
    text = "ｶﾅ,1\r\nｱｲｳ,2\r\n"
    d = decode(text.encode("cp932"))
    assert (d.encoding, d.basis, d.candidates) == ("cp932", "multibyte", [])
    assert grade_blob(text.encode("cp932"), "a.csv").grade == "S"


@pytest.mark.parametrize("enc, bom", [("utf-16-le", b""), ("utf-16-be", b""), ("utf-32-le", b"\xff\xfe\x00\x00"),
                                      ("utf-32-be", b"\x00\x00\xfe\xff")])
def test_utf16_32_with_and_without_bom(enc, bom):
    text = "name,value\r\n申請,1\r\n"
    d = decode(bom + text.encode(enc))
    assert d.encoding == enc and d.text == text and not d.undetermined
    assert d.byte_of(len(text)) == len(bom) + len(text.encode(enc))


def test_single_byte_text_gives_candidates_not_latin1():
    text = "Привет, мир\r\nДанные\r\n"
    d = decode(text.encode("cp1251"))
    assert d.encoding != "latin-1"
    names = [d.encoding] + [c["encoding"] for c in d.candidates]
    assert "cp1251" in names and len(names) > 1 and d.undetermined


def test_undecodable_bytes_are_said_not_silently_latin1(tmp_path):
    blob = b"abc\xff\xfe\xfd\x81\x81\r\n"  # どの多バイトにも合わず、1バイトでも C1 の字か割り当てのない字になる並び
    d = decode(blob)
    if d.basis == "none":
        p = tmp_path / "a.txt"
        p.write_bytes(blob)
        _k, _doc, parts = api.window(str(p))
        assert "文字コードを決められませんでした" in parts[0][1]
        assert any(r["title"] == "文字コードを決められない" for r in api.grade(str(p))["reasons"])
    else:
        assert d.undetermined or d.basis in ("multibyte", "singlebyte")
