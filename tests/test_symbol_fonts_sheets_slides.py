"""記号の書体（Wingdings・Symbol・Wingdings 2）で書いたセル・スライドの字：見える字で読み、書くときは、その書体で描かれる形を確かめる。

Excel（.xlsx・.xls の BIFF8）：セルの値（cell.value）は保存された字のまま（数式・計算はその字を使う）。窓・欄・書き込みの頼み方・
読み直し・検査は、書体で読んだ字（見える字）で行い、書く直前に保存する字へ戻す（xlsx/symcells.py）。
PowerPoint（.pptx）：字（a:r）の a:latin・a:sym の書体で読み、書いた字には記号の書体を引き継がない（pptx/fill.clean_rpr）。
どれも、ここで作ったファイル（.xls は、同じ .xlsx を LibreOffice が書き出したもの）で確かめる。
"""

from __future__ import annotations

import zipfile
from xml.sax.saxutils import escape

import pytest

import pptx_maker
import symbook_maker
from todoku import api, lazy, limits, symfont
from todoku.api import ApiError
from todoku.xls.book import XlsWorkbook
from todoku.xlsx.workbook import Workbook

FIX = __import__("pathlib").Path(__file__).parent / "fixtures"


@pytest.fixture()
def book(tmp_path):
    p = tmp_path / "a.xlsx"
    p.write_bytes(symbook_maker.make())
    return p


def _body(path) -> str:
    return "\n".join(text for _name, text in api.window(str(path))[2])


def _stored(path, *cells):
    blob = open(path, "rb").read()
    wb = XlsWorkbook(blob, "o.xls") if str(path).endswith(".xls") else Workbook(blob, "o.xlsx")
    got = [getattr(wb.sheets[0].cells.get(rc), "value", None) for rc in cells]
    if hasattr(wb, "zip"):
        getattr(wb.zip, "close", lambda: None)()
    return got


def _fill(path, values, tmp_path, name="out"):
    return api.fill(str(path), values, str(tmp_path / (name + path.suffix)))


# ====================================================================== 対応表の逆引き
def test_code_of_is_the_inverse_of_the_table():
    assert symfont.code_of("Wingdings", "🗹") == 0xFE
    assert symfont.code_of("Wingdings", "🞎") == 0xA8
    assert symfont.code_of("Symbol", "α") == 0x61
    assert symfont.code_of("wingdings  2", "☑") == 0x52
    for code in range(0x21, 0x100):
        ch = symfont.decode("Wingdings", code)
        if 0xF000 <= ord(ch) <= 0xF0FF or code == 0xA0 or 0x80 <= code <= 0x9F:
            continue
        got = symfont.code_of("Wingdings", ch)
        assert got is not None and symfont.decode("Wingdings", got) == ch


def test_code_of_has_no_code_for_spaces_other_fonts_and_unmapped_characters():
    assert symfont.code_of("Wingdings", " ") is None  # 空白は空白のまま読む
    assert symfont.code_of("Wingdings", "山") is None
    assert symfont.code_of("Arial", "a") is None
    assert symfont.code_of("", "a") is None


# ====================================================================== .xlsx：読む
def test_window_shows_the_characters_as_the_font_draws_them(book):
    body = _body(book)
    for line in ("Agree ｜ 〔B1〕🗹［？］", "Decline ｜ 〔B2〕🞎［？］", "Alpha ｜ αβγ", "Japanese ｜ 山田", "W2 ｜ 〔B6〕☑［？］☒［？］",
                 "PUA ｜ ✓", "Number ｜ 🗁", "Rich ｜ Rich ✓ done"):
        assert line in body, line
    for raw in ("þ", "¨", "abg", "RT", "ü", ""):
        assert raw not in body, raw


def test_cell_value_stays_the_stored_characters(book):
    assert _stored(book, (1, 2), (3, 2), (7, 2), (10, 2)) == ["þ", "abg", "", "Rich ü done"]
    wb = Workbook(book.read_bytes(), "a.xlsx")
    assert wb.has_sym
    assert wb.sym_runs(wb.sheets[0].cells[(10, 2)]) is not None
    assert wb.sym_runs(wb.sheets[0].cells[(1, 1)]) is None


def test_fields_use_the_seen_characters(book):
    fs = {f["id"]: f for f in api.fields(str(book))}
    assert fs["Sheet1!B1"]["value"] == "🗹" and fs["Sheet1!B2"]["value"] == "🞎" and fs["Sheet1!B6"]["value"] == "☑☒"
    assert [s["id"] for s in fs["Sheet1!B6"]["spans"]] == ["Sheet1!B6@0", "Sheet1!B6@1"]
    assert "þ" not in str(fs["Sheet1!B1"]) and "¨" not in str(fs["Sheet1!B2"])


def test_formulas_read_the_stored_characters(tmp_path):
    p = tmp_path / "f.xlsx"
    p.write_bytes(symbook_maker.make(formula=True))
    r = api.calculate(str(p), {"Sheet1!B2": True})
    assert r["inputs"] == {"Sheet1!B2": "🗹"}  # 伝える値は見える字
    assert [(c["cell"], c["before"], c["after"]) for c in r["changed"]] == [("Sheet1!D1", "1", "2")]  # =COUNTIF(B1:B2,"þ")
    got = api.calculate(str(p), {}, cells=["Sheet1!B1", "Sheet1!B3"])["cells"]
    assert [c["value"] for c in got] == ["🗹", "αβγ"]


def test_ordinary_workbooks_are_not_touched(tmp_path):
    import openpyxl
    wb = openpyxl.Workbook()
    wb.active["A1"] = "þ abg"
    p = tmp_path / "plain.xlsx"
    wb.save(p)
    assert not Workbook(p.read_bytes(), "plain.xlsx").has_sym
    assert "þ abg" in _body(p)


# ====================================================================== .xlsx：書く
def test_checking_a_box_writes_the_fonts_own_box(book, tmp_path):
    r = _fill(book, {"Sheet1!B2": True, "Sheet1!B1": False}, tmp_path)
    assert r["proof_ok"] and r["read_back_ok"]
    assert _stored(r["out_path"], (1, 2), (2, 2)) == ["¨", "þ"]
    body = _body(r["out_path"])
    assert "Agree ｜ 〔B1〕🞎［？］" in body and "Decline ｜ 〔B2〕🗹［？］" in body


def test_one_box_of_two_keeps_the_other_stored_character(book, tmp_path):
    r = _fill(book, {"Sheet1!B6": [None, False]}, tmp_path)
    assert r["read_back_ok"]
    assert _stored(r["out_path"], (6, 2)) == ["R£"]  # Wingdings 2 の、印のない箱


def test_a_symbol_is_written_as_the_fonts_code(book, tmp_path):
    r = _fill(book, {"Sheet1!B9": "🗹", "Sheet1!B3": {"value": "δ", "overwrite_text": True}}, tmp_path)
    assert r["read_back_ok"] and not [w for w in r["warnings"] if "記号の書体" in w]
    assert _stored(r["out_path"], (9, 2), (3, 2)) == ["þ", "d"]
    assert "Empty ｜ " in _body(r["out_path"]) and "Alpha ｜ δ" in _body(r["out_path"])


def test_an_ordinary_character_the_font_would_redraw_is_refused(book, tmp_path):
    with pytest.raises(ApiError) as e:
        _fill(book, {"Sheet1!B9": "A1"}, tmp_path)
    assert "記号の書体（Wingdings）" in str(e.value) and "「A」" in str(e.value) and "✌" in str(e.value)
    with pytest.raises(ApiError) as e:
        _fill(book, {"Sheet1!B3": {"value": "x", "overwrite_text": True}}, tmp_path, "o2")
    assert "記号の書体（Symbol）" in str(e.value) and "ξ" in str(e.value)


def test_characters_the_font_does_not_remap_are_kept_with_a_warning(book, tmp_path):
    r = _fill(book, {"Sheet1!B9": "山田"}, tmp_path)
    assert r["read_back_ok"]
    assert _stored(r["out_path"], (9, 2)) == ["山田"]
    assert any("「山田」" in w and "記号に読み替えない字" in w for w in r["warnings"])


def test_a_number_is_written_and_the_drawn_symbols_are_told(book, tmp_path):
    r = _fill(book, {"Sheet1!B9": 12}, tmp_path)
    assert r["read_back_ok"]
    assert _stored(r["out_path"], (9, 2)) == [12.0]
    assert any("「🗁🗎」の形で表示されます" in w for w in r["warnings"])


def test_rich_cell_keeps_each_characters_font(book, tmp_path):
    r = _fill(book, {"@1": {"old": "✓", "new": "🗹", "cell": "Sheet1!B10"}}, tmp_path)
    assert r["proof_ok"] and r["read_back_ok"]
    assert _stored(r["out_path"], (10, 2)) == ["Rich þ done"]
    wb = Workbook(open(r["out_path"], "rb").read(), "o.xlsx")
    assert [f for _n, f in wb.sym_runs(wb.sheets[0].cells[(10, 2)]) if f] == ["Wingdings"]
    assert "Rich 🗹 done" in _body(r["out_path"])
    # ふつうの書体の字の所に書いた ✓ は、✓ のまま（Wingdings の字は元のまま）
    r = _fill(book, {"@1": {"old": "done", "new": "✓", "cell": "Sheet1!B10"}}, tmp_path, "o2")
    assert r["read_back_ok"] and _stored(r["out_path"], (10, 2)) == ["Rich ü ✓"]
    with pytest.raises(ApiError) as e:
        _fill(book, {"@1": {"old": "✓", "new": "x", "cell": "Sheet1!B10"}}, tmp_path, "o3")
    assert "記号の書体（Wingdings）" in str(e.value)


def test_audit_trial_write_and_read_back_pass(book):
    from todoku.xlsx import audit

    checks = {c.name: c for c in audit.run(book.read_bytes(), "a.xlsx").checks}
    assert checks["全入力欄への書き込み"].ok, checks["全入力欄への書き込み"].items
    assert checks["読み戻し"].ok, checks["読み戻し"].items


def test_grade_does_not_count_a_decoded_private_use_character(book):
    from todoku import mcp_server

    g = mcp_server.tool_result("grade_document", {"path": str(book)})["structuredContent"]
    assert "外字" not in str(g.get("reasons")) and "私用" not in str(g.get("reasons"))


# ====================================================================== .xlsx：大きな表の道
def test_large_sheet_reads_and_writes_the_same_way(tmp_path, monkeypatch):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)
    monkeypatch.setattr(limits, "PAGE_CHARS", 1_000)
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    try:
        p = tmp_path / "big.xlsx"
        p.write_bytes(symbook_maker.make(extra=600))
        text = "".join(api.read_range(str(p), {"sheet": "Sheet1", "rows": rows})[0].text for rows in ("1-9", "10-14"))
        for line in ("Agree ｜ 〔B1〕🗹［？］", "Alpha ｜ αβγ", "PUA ｜ ✓", "Rich ｜ Rich ✓ done", "〔C11〕🗹［？］", "〔C12〕🞎［？］"):
            assert line in text, line
        assert "þ" not in text and "abg" not in text
        r = api.fill(str(p), {"Sheet1!B2": True, "Sheet1!C12": True, "Sheet1!B9": "🗹"}, str(tmp_path / "big_out.xlsx"))
        assert r["streamed"] and r["proof_ok"] and r["read_back_ok"]
        assert _stored(r["out_path"], (2, 2), (12, 3), (9, 2)) == ["þ", "þ", "þ"]
        with pytest.raises(ApiError) as e:
            api.fill(str(p), {"Sheet1!B9": "A"}, str(tmp_path / "big_out2.xlsx"))
        assert "記号の書体（Wingdings）" in str(e.value)
    finally:
        lazy.clear()
        api._WINDOWS.clear()


# ====================================================================== .xls（BIFF8）
def test_xls_reads_the_cell_font_and_the_rich_runs():
    p = FIX / "symbol_fonts.xls"
    wb = XlsWorkbook(p.read_bytes(), p.name)
    assert wb.has_sym
    assert wb.sheets[0].cells[(1, 2)].value == "þ"
    body = _body(p)
    for line in ("Agree ｜ 〔B1〕🗹［？］", "Decline ｜ 〔B2〕🞎［？］", "Alpha ｜ αβγ", "W2 ｜ 〔B6〕☑［？］☒［？］", "Rich ｜ Rich ✓ done"):
        assert line in body, line
    assert "þ" not in body and "abg" not in body


def test_xls_writes_the_fonts_code_and_keeps_rich_runs(tmp_path):
    p = FIX / "symbol_fonts.xls"
    r = _fill(p, {"Sheet1!B1": False, "Sheet1!B2": True, "Sheet1!B9": "🗹", "@1": {"old": "✓", "new": "🗹", "cell": "Sheet1!B10"}}, tmp_path)
    assert r["proof_ok"] and r["read_back_ok"]
    assert _stored(r["out_path"], (1, 2), (2, 2), (9, 2), (10, 2)) == ["¨", "þ", "þ", "Rich þ done"]
    wb = XlsWorkbook(open(r["out_path"], "rb").read(), "o.xls")
    assert [f for _n, f in wb.sym_runs(wb.sheets[0].cells[(10, 2)]) if f] == ["Wingdings"]
    with pytest.raises(ApiError) as e:
        _fill(p, {"Sheet1!B9": "A"}, tmp_path, "o2")
    assert "記号の書体（Wingdings）" in str(e.value)


# ====================================================================== .pptx
W = '<a:latin typeface="Wingdings" pitchFamily="2" charset="2"/>'
WS = '<a:sym typeface="Wingdings" panose="05000000000000000000" pitchFamily="2" charset="2"/>'
SYMBOL = '<a:latin typeface="Symbol" pitchFamily="18" charset="2"/>'


def _run(t: str, fonts: str = "") -> str:
    return f'<a:r><a:rPr lang="en-US" sz="2400">{fonts}</a:rPr><a:t>{escape(t)}</a:t></a:r>'


def _p(*runs: str) -> str:
    return "<a:p>" + "".join(runs) + '<a:endParaRPr lang="en-US" sz="2400"/></a:p>'


@pytest.fixture()
def deck(tmp_path):
    paras = [
        _p(_run("1 "), _run("þ", W), _run(" Agree")),
        _p(_run("2 "), _run("¨", W), _run(" Decline")),
        _p(_run("3 "), _run("abg", SYMBOL)),
        _p(_run("4 Next "), _run("", WS), _run(" step")),
        _p(_run("5 "), _run("q", WS)),  # a:sym だけ：私用の領域の字でない字は読み替えない
        _p(_run("6 "), _run("ü", '<a:latin typeface="Arial"/>' + WS)),
        _p(_run("7 "), _run("A", '<a:latin typeface="Marlett" charset="2"/>')),  # 対応表のない、記号と宣言した書体
        _p(_run("8 "), _run("ü", '<a:latin typeface="+mn-lt"/>')),  # テーマの書体への参照
    ]
    p = tmp_path / "a.pptx"
    p.write_bytes(pptx_maker.make([pptx_maker.sp("".join(paras))]))
    return p


def _lines(path) -> list[str]:
    return [ln for ln in _body(path).split("\n") if ln.startswith("- ") and ln[2:3].isdigit()]


def test_pptx_runs_are_read_by_their_symbol_font(deck):
    got = _lines(deck)
    assert got[0].startswith("- 1 🗹［P1.c1？］ Agree")
    assert got[1].startswith("- 2 🞎［P2.c1？］ Decline")
    assert got[2].startswith("- 3 αβγ")
    assert got[3].startswith("- 4 Next 🡪 step")
    assert got[4].startswith("- 5 q")
    assert got[5].startswith("- 6 ✓") and "ü" in got[5]
    assert got[6].startswith("- 7 ")  # 外字として伝える（私用の領域の字のまま）
    assert got[7].startswith("- 8 ü")


def test_pptx_written_characters_do_not_inherit_the_symbol_font(deck, tmp_path):
    r = _fill(deck, {"P2.c1": True, "P1.c1": False, "@1": {"old": "αβγ", "new": "αβδ"}}, tmp_path)
    assert r["proof_ok"] and r["read_back_ok"]
    assert sum("記号の書体（Wingdings）" in w and "記号の書体を外しました" in w for w in r["warnings"]) == 2
    assert any("記号の書体（Symbol）" in w for w in r["warnings"])
    got = _lines(r["out_path"])
    assert got[0].startswith("- 1 ☐［P1.c1？］ Agree") and got[1].startswith("- 2 ■［P2.c1？］ Decline") and got[2].startswith("- 3 αβδ")
    with zipfile.ZipFile(r["out_path"]) as z:
        xml = z.read("ppt/slides/slide1.xml").decode("utf-8")
    # 書き換えなかった字（ab）は Symbol の字のまま、書いた字（δ）は記号の書体のない字の要素
    assert f'<a:rPr lang="en-US" sz="2400">{SYMBOL}</a:rPr><a:t>ab</a:t>' in xml
    assert '<a:rPr lang="en-US" sz="2400"></a:rPr><a:t>δ</a:t>' in xml
    assert "Wingdings" not in xml.split("<a:t>4 Next")[0]


def test_pptx_writing_beside_a_symbol_run_keeps_it(deck, tmp_path):
    r = _fill(deck, {"@1": {"old": "3 αβγ", "new": "3 OK αβγ"}}, tmp_path)
    assert r["proof_ok"] and r["read_back_ok"] and not [w for w in r["warnings"] if "記号の書体" in w]
    assert _lines(r["out_path"])[2].startswith("- 3 OK αβγ")


def test_pptx_clean_rpr_removes_only_symbol_fonts():
    from todoku.pptx.fill import clean_rpr

    rpr = (b'<a:rPr lang="ja-JP" sz="1800"><a:latin typeface="Wingdings" charset="2"/><a:ea typeface="\xef\xbc\xad\xef\xbc\xb3 \xe6\x98\x8e\xe6\x9c\x9d"/>'
           b'<a:sym typeface="Symbol"/></a:rPr>')
    new, names = clean_rpr(rpr)
    assert new == b'<a:rPr lang="ja-JP" sz="1800"><a:ea typeface="\xef\xbc\xad\xef\xbc\xb3 \xe6\x98\x8e\xe6\x9c\x9d"/></a:rPr>'
    assert names == ["記号の書体（Wingdings）", "記号の書体（Symbol）"]
    plain = b'<a:rPr lang="ja-JP"><a:latin typeface="Arial"/><a:cs typeface="+mn-cs"/></a:rPr>'
    assert clean_rpr(plain) == (plain, [])


def test_pptx_audit_counts_the_same_characters(deck):
    from todoku.pptx import audit

    checks = {c.name: c for c in audit.run(deck.read_bytes(), "a.pptx").checks}
    assert checks["読み落とし（スライドの字）"].ok, checks["読み落とし（スライドの字）"].items
    assert checks["全入力欄への書き込み"].ok and checks["読み戻し"].ok
