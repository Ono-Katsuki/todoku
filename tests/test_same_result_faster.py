"""速くした所は、速くする前と同じ結果を返す（結果は 1 字も変えない）。
控え（窓・範囲の値・表示の字）を使う道と、使わない道を、同じ書類でくらべる。"""

from __future__ import annotations

import json
import random

import pytest

from todoku import api, grade as grade_mod, mcp_server
from todoku.book import open_book
from todoku.formula.numfmt import _shown_number, format_value
from todoku.xlsx import window as xwin
from todoku.xlsx.calc import Calc

openpyxl = pytest.importorskip("openpyxl")


def _same(a, b) -> bool:
    return json.dumps(a, ensure_ascii=False, sort_keys=True, default=repr) == json.dumps(b, ensure_ascii=False, sort_keys=True, default=repr)


def _form(tmp_path, name="a.xlsx"):
    """見出しの下に空のマスが縦に並ぶ表・表を引く数式・ほかのシートを読む数式のあるブック。"""
    from openpyxl.styles import Border, Side

    box = Border(*(Side(style="thin"),) * 4)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "申請"
    ws["A1"], ws["B1"], ws["C1"], ws["D1"] = "番号", "氏名", "単価コード", "単価"
    for r in range(2, 42):
        ws.cell(r, 1, r - 1)
        for c in (2, 3):
            ws.cell(r, c).border = box
        ws.cell(r, 4, f"=IFERROR(VLOOKUP(C{r},表!$A$1:$B$30,2,FALSE),\"\")")
    ws["A43"], ws["D43"] = "合計", "=SUM(D2:D41)"
    ws["A44"], ws["D44"] = "今日", "=TODAY()"
    ws["A45"], ws["D45"] = "今日までの合計", "=SUM(D43:D44)"
    t = wb.create_sheet("表")
    for r in range(1, 31):
        t.cell(r, 1, f"K{r:02d}")
        t.cell(r, 2, r * 110)
    t["D1"] = "=申請!D43"
    p = tmp_path / name
    wb.save(p)
    return p


# ====================================================================== 段階：開いた窓を使う
def test_grade_with_the_opened_window_is_the_same(tmp_path):
    p = _form(tmp_path)
    api._WINDOWS.clear()
    fresh = api.grade(str(p))  # 窓の控えがない：段階の中で窓を作る
    api._WINDOWS.clear()
    win = api.window(str(p))
    again = api.grade(str(p))  # 開いた窓を使う
    assert _same(fresh, again)
    assert _same(api.window(str(p)), win)  # 段階を決めても、控えた窓は変わらない
    api._WINDOWS.clear()


def test_built_window_is_used_only_for_the_same_bytes():
    blob = b"abc"
    with grade_mod.built_window(blob, {"x": 1}):
        assert grade_mod._built(blob) == {"x": 1}
        assert grade_mod._built(bytes(bytearray(b"abc"))) is None  # 同じ字でも、渡された中身そのものでなければ使わない
    assert grade_mod._built(blob) is None


# ====================================================================== 窓の控え
def test_cached_window_is_the_same_as_a_fresh_one(tmp_path):
    p = _form(tmp_path)
    api._WINDOWS.clear()
    first = api.window(str(p))
    snap = next(iter(api._WINDOWS.values()))[1]
    assert isinstance(snap, bytes)
    second = api.window(str(p))
    assert second is not first and _same(first, second)
    second[1]["sheets"].clear()  # 受け取った側が変えても、控えは変わらない
    assert _same(api.window(str(p)), first)
    assert _same(api.fields(str(p)), api._fields_of(first[0], first[1]))
    api._WINDOWS.clear()


def test_summary_counts_the_fields_of_the_window_it_holds(tmp_path):
    p = _form(tmp_path)
    api._WINDOWS.clear()
    s = api.summary(str(p))
    fl = api.fields(str(p))
    assert s["fields"] == len(fl) > 40 and s["writable_fields"] == sum(1 for f in fl if not f.get("readonly"))
    api._WINDOWS.clear()


# ====================================================================== 範囲の値
def test_range_rows_equals_reading_cell_by_cell(tmp_path):
    p = _form(tmp_path)
    blob = p.read_bytes()
    plain, kept = Calc(open_book(blob, p.name)), Calc(open_book(blob, p.name))
    kept.ranges = {}
    for sheet, r1, c1, r2, c2 in (("申請", 1, 1, 45, 4), ("表", 1, 1, 30, 4), ("申請", 43, 4, 45, 4), ("申請", 50, 1, 52, 3)):
        want = [[plain.cell_value(sheet, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        assert plain.range_rows(sheet, r1, c1, r2, c2) == want
        got = kept.range_rows(sheet, r1, c1, r2, c2)
        assert got == want
        got[0][0] = "変えた"  # 受け取った側が変えても、控えは変わらない
        assert kept.range_rows(sheet, r1, c1, r2, c2) == want
    assert plain.status == kept.status


def test_kept_range_still_marks_the_reader_volatile(tmp_path):
    # 今日の日付のマスを含む範囲を、2 つの数式が読む：2 つ目は控えから読むが、「計算のたびに変わる」の印は同じに付く
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"], ws["A2"] = "=TODAY()", 5
    ws["B1"], ws["B2"] = "=SUM(A1:A2)", "=SUM(A1:A2)"
    ws["C1"], ws["C2"] = "=SUM(A2:A2)", "=SUM(A2:A2)"
    p = tmp_path / "v.xlsx"
    wb.save(p)
    blob = p.read_bytes()
    plain, kept = Calc(open_book(blob, p.name)), Calc(open_book(blob, p.name))
    kept.ranges = {}
    for c in (plain, kept):
        for r in (1, 2):
            c.cell_value("Sheet", r, 2), c.cell_value("Sheet", r, 3)
    assert kept.ranges  # 控えを使った
    assert plain.status == kept.status
    assert kept.status[("Sheet", 2, 2)][0] == "volatile" and kept.status.get(("Sheet", 2, 3), ("ok", ""))[0] != "volatile"


def test_window_does_not_change_when_ranges_are_not_kept(tmp_path, monkeypatch):
    p = _form(tmp_path)
    blob = p.read_bytes()
    with_keep = xwin.build(open_book(blob, p.name))
    monkeypatch.setattr("todoku.xlsx.calc._RANGES_ROOM", 0)  # 控える余地がない：どの範囲も読み直す
    assert _same(xwin.build(open_book(blob, p.name)), with_keep)


# ====================================================================== 数式が読むマス
def test_readers_of_all_sheets_at_once_equal_sheet_by_sheet(tmp_path):
    p = _form(tmp_path)
    wb = open_book(p.read_bytes(), p.name)
    together = xwin._formula_readers(wb, wb.sheets)
    for sh, got in zip(wb.sheets, together):
        alone = xwin._formula_readers(wb, [sh])[0]
        assert got.places() == alone.places()
        assert [got.get(k) for k in got.places()] == [alone.get(k) for k in alone.places()]
    assert together[0].places() and together[1].places()


# ====================================================================== 表示の字
def test_kept_number_text_equals_formatting_again():
    cases = [(1234.5, "#,##0.00"), (0.0, "0.00"), (-0.0, "0.00"), (-0.0, "General"), (0.0, "General"), (45000.0, "yyyy/m/d"),
             (45000.0, "[$-411]ggge\"年\"m\"月\"d\"日\""), (0.125, "0%"), (1e15, "General"), (-3.0, "#,##0;[Red]-#,##0"), (59.0, "yyyy/m/d"), (60.0, "yyyy/m/d")]
    for d in (False, True):
        for x, fmt in cases:
            want = _shown_number(x, fmt, d)
            assert format_value(x, fmt, d) == want
            assert format_value(x, fmt, d) == want  # 2 度目は控えから
    assert format_value(45000.0, "yyyy/m/d", False) != format_value(45000.0, "yyyy/m/d", True)
    assert format_value(True, "0.00", False) is not None and format_value(1, "0", False) == format_value(1.0, "0", False)


# ====================================================================== list_fields の大きさ
def test_fits_equals_measuring_the_whole_result():
    rnd = random.Random(7)
    words = ["氏名", "住所", "a\"b", "改行\nあり", "", "x" * 40, "１．一般　２．特定"]
    for _ in range(300):
        fl = [{"id": f"S!A{i}", "label": rnd.choice(words), "near": {"left": {"text": rnd.choice(words), "mm": rnd.random()}},
               "n": rnd.choice([None, 1, 2.5, True]), "list": [rnd.choice(words) for _ in range(rnd.randrange(3))]}
              for i in range(rnd.randrange(0, 12))]
        rest = {"path": rnd.choice(words), "note": [rnd.choice(words)] * rnd.randrange(3), "count": len(fl)}
        whole = mcp_server._size({**rest, "fields": fl})
        for budget in (0, 1, whole - 1, whole, whole + 1, whole * 2):
            assert mcp_server._fits(rest, fl, budget) == (whole <= budget)


# ====================================================================== 書いた後
def test_filled_file_is_opened_from_the_window_made_at_read_back(tmp_path):
    p = _form(tmp_path)
    out = tmp_path / "out.xlsx"
    api._WINDOWS.clear()
    r = api.fill(str(p), {"申請!B2": "山田花子", "申請!C2": "K03"}, str(out))
    assert r["proof_ok"] and r["verified"] and r["read_back_ok"]
    assert api._NAMED_AS.get() is None  # 読み戻しのあいだだけ
    names = [k[1] for k in api._WINDOWS]
    assert "out.xlsx" in names and not [n for n in names if n not in ("a.xlsx", "out.xlsx")]  # 一時ファイルの名前の窓を控えに残さない
    kept = api.window(str(out))  # 読み戻しで作った窓
    api._WINDOWS.clear()
    fresh = api.window(str(out))  # 書き出したファイルを、初めから開いた窓
    assert _same(kept, fresh)
    assert fresh[1]["source"]["file"] == "out.xlsx"
    api._WINDOWS.clear()


def test_fill_result_is_the_same_with_and_without_the_proved_book(tmp_path, monkeypatch):
    p = _form(tmp_path)
    api._WINDOWS.clear()
    a = api.fill(str(p), {"申請!B2": "山田花子", "申請!C2": "K03"}, str(tmp_path / "o1.xlsx"))
    from todoku.xlsx import reach

    real = reach.prove_whole

    def no_book(*args, **kw):  # 証明で開いたブックを渡さない：読み戻しで開き直す
        s = real(*args, **kw)
        s.book = None
        return s

    monkeypatch.setattr(reach, "prove_whole", no_book)
    api._WINDOWS.clear()
    b = api.fill(str(p), {"申請!B2": "山田花子", "申請!C2": "K03"}, str(tmp_path / "o2.xlsx"))
    assert (tmp_path / "o1.xlsx").read_bytes() == (tmp_path / "o2.xlsx").read_bytes()

    def strip(x):
        return json.loads(json.dumps(x, ensure_ascii=False, default=repr).replace("o2.xlsx", "o1.xlsx"))

    assert _same(strip(a), strip(b))
    api._WINDOWS.clear()


def test_checkbox_refusal_lists_declared_states_in_one_order():
    from todoku.inputcheck import InputError, checkbox_value

    said = set()
    for on in ({"☑", "■", "✓"}, {"✓", "☑", "■"}, frozenset("■✓☑")):
        with pytest.raises(InputError) as e:
            checkbox_value("はい", on, {"□", "☐"})
        said.add(str(e.value))
    assert len(said) == 1 and "■・☑・✓・□・☐" in said.pop()
    with pytest.raises(InputError, match="はい・いいえ"):  # 並びで渡された値は、その順のまま
        checkbox_value("x", ["はい"], ["いいえ"])
