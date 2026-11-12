"""基本多言語面の外の字（「𠮷」・絵文字：UTF-16 で 2 単位）のある .xls・.xlsb の字。

ファイルの中の字数と、字ごとの書式・ふりがなの位置は UTF-16 の単位で数える。前は、書くときに字（符号位置）の数を
字数として書いていたので、「𠮷田」を書くと読み直しが「𠮷」になり、読み直しで止めていた。
"""
from __future__ import annotations

import struct
from pathlib import Path

import xlsb_maker as m

from todoku import api, limits
from todoku.book import open_book
from todoku.xls import write
from todoku.xls.book import XlsWorkbook, _phonetic, cp_runs

FIX = Path(__file__).parent / "fixtures" / "symbol_fonts.xls"


def _book(path) -> XlsWorkbook:
    return XlsWorkbook(Path(path).read_bytes(), "o.xls")


def test_counts_are_utf16_units():
    assert write._xlstr("𠮷田")[:3] == struct.pack("<HB", 3, 1)
    assert write._xlstr("abc")[:3] == struct.pack("<HB", 3, 0)
    assert write._units("𠮷a😀b", 0) == 0 and write._units("𠮷a😀b", 1) == 2 and write._units("𠮷a😀b", 3) == 5
    assert cp_runs("𠮷a😀b", [(2, 5), (5, 6)]) == [(1, 5), (3, 6)]
    runs = [(1, 5)]
    assert cp_runs("abc", runs) is runs


def test_write_astral_text_to_xls_cell(tmp_path):
    out = tmp_path / "o.xls"
    r = api.fill(str(FIX), {"Sheet1!C1": "𠮷田", "Sheet1!C2": "a😀b", "Sheet1!C3": "😀"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"] and r["verified"]
    cells = _book(out).sheets[0].cells
    assert [cells[(k, 3)].value for k in (1, 2, 3)] == ["𠮷田", "a😀b", "😀"]


def test_rich_xls_cell_keeps_runs_after_astral_text(tmp_path):
    out = tmp_path / "o.xls"
    r = api.fill(str(FIX), {"@1": {"old": "Rich", "new": "𠮷😀", "cell": "Sheet1!B10"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    wb = _book(out)
    cell = wb.sheets[0].cells[(10, 2)]
    assert cell.value == "𠮷😀 ü done"
    # 字ごとの書体の境は、字の位置で読める（ファイルの中では UTF-16 の単位：「𠮷😀 」は 5 単位）
    assert wb.sym_runs(cell) == [(3, ""), (1, "Wingdings"), (5, "")]
    assert "𠮷😀 ✓ done" in api.window(str(out))[2][0][1]


def test_rich_astral_cell_on_the_large_sheet_path(tmp_path, monkeypatch):
    out = tmp_path / "o.xls"
    api.fill(str(FIX), {"@1": {"old": "Rich", "new": "𠮷😀", "cell": "Sheet1!B10"}}, str(out))
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)
    monkeypatch.setattr(limits, "PAGE_CHARS", 1_000)
    text = "".join(api.read_range(str(out), {"sheet": "Sheet1", "rows": rows})[0].text for rows in ("1-9", "10-14"))
    assert "𠮷😀 ✓ done" in text


def _ext(st: str, runs: list) -> bytes:
    """ExtRst（MS-XLS 2.5.109）：Phs・RPHSSub（読みの字）・PhRuns（読みの頭・親文字の頭・親文字の数）。"""
    sub = struct.pack("<HHH", len(runs), len(st), len(st)) + st.encode("utf-16-le")
    body = struct.pack("<HH", 0, 0x37) + sub + b"".join(struct.pack("<HHH", *x) for x in runs)
    return struct.pack("<HH", 1, len(body)) + body


def test_phonetic_positions_are_utf16_units():
    assert _phonetic("吉野の山", _ext("ヤマ", [(0, 3, 1)])) == "吉野のヤマ"
    assert _phonetic("𠮷野の山", _ext("ヤマ", [(0, 4, 1)])) == "𠮷野のヤマ"
    assert _phonetic("𠮷野の山", _ext("ヨシノヤマ", [(0, 0, 3), (3, 4, 1)])) == "ヨシノのヤマ"


def test_xlsb_same_length_text_with_astral(tmp_path):
    p = tmp_path / "a.xlsb"
    p.write_bytes(m.sample())
    out = tmp_path / "b.xlsb"
    # 「直接の字」は UTF-16 で 4 単位：「𠮷田山」（2 + 1 + 1）は同じ大きさで書ける
    r = api.fill(str(p), {"申請!C1": {"value": "𠮷田山", "overwrite_text": True}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    assert open_book(out.read_bytes(), "b.xlsb").sheet("申請").cells[(1, 3)].value == "𠮷田山"
