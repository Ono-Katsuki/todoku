"""表のセルに字で指定した数と日付：ほかの地域の書き方・ほかの文字の数字を、推測せずに読む。

- 字で指定する数の書き方は 1 つ（小数点「.」・桁区切り「,」。JSON の数と、窓が示す数と同じ）
- 読み方が 1 つに決まる、ほかの書き方（1.234,56・1,5・1 234,56・12,34,567）と、ほかの文字の数字（١٢٣）は、数の入るセル
  （数の表示形式・数式から参照・数の入力規則）では数にして、どう読んだかを伝える。ほかのセルでは、書いたとおりの字
- 「1.234」「1,234」の形は、地域で値が 1000 倍違う：決まった書き方で読み、セルが同じ字で表示するときのほかは、読み方を伝える
- 通貨記号は Unicode の通貨記号（Sc）のどれでも、前でも後ろでも
- 日付：年が先頭の形に、ハングルの年月日・終わりの「.」・ほかの文字の数字。Excel の始まりの日より前は、字として書いて伝える
"""

import io
import zipfile

import pytest

from todoku import api
from todoku.xlsx.edit import interpret_input, parse_date_text, plain_digits

APR1 = 46113.0  # 2026-04-01（1900 年の数え方）
FORMATS = {1: "General", 2: "#,##0.00", 3: "yyyy/mm/dd", 4: "0%", 5: "@", 6: "General", 7: "#,##0", 8: "0.000", 9: "0"}


def _book_bytes(mac1904: bool = False) -> bytes:
    openpyxl = pytest.importorskip("openpyxl")
    wb_ = openpyxl.Workbook()
    if mac1904:
        wb_.epoch = openpyxl.utils.datetime.CALENDAR_MAC_1904
    ws = wb_.active
    for r, f in FORMATS.items():
        ws.cell(r, 1, "x" if f == "@" else 0).number_format = f
    ws["B6"] = "=A6*2"  # A6 は、数式から参照される標準のセル
    buf = io.BytesIO()
    wb_.save(buf)
    return buf.getvalue()


@pytest.fixture(scope="module")
def book():
    from todoku.book import open_book

    wb = open_book(_book_bytes(), "n.xlsx")
    return wb, wb.sheets[0].name


def _put(book, row, value):
    wb, name = book
    return interpret_input(wb, name, row, 1, value)


# ====================================================================== 数字と記号をそろえる
@pytest.mark.parametrize("typed,plain", [
    ("١٢٣", "123"), ("۱۲۳", "123"), ("१२३", "123"), ("๑๒๓", "123"), ("１２３", "123"),
    ("١٬٢٣٤٫٥", "1,234.5"), ("٥٪", "5%"), ("−5", "-5"), ("abc 氏名", "abc 氏名"),
])
def test_plain_digits(typed, plain):
    assert plain_digits(typed) == plain


# ====================================================================== ほかの書き方：読み方が 1 つに決まる形
OTHER = [("1.234,56", 1234.56), ("1 234,56", 1234.56), ("1 234,56", 1234.56), ("1 234,56", 1234.56), ("1'234.56", 1234.56),
         ("1.234.567", 1234567.0), ("1,5", 1.5), ("0,125", 0.125), ("1234,567", 1234.567), ("12,34,567.89", 1234567.89),
         ("123 456", 123456.0), ("١٢٣", 123.0), ("۱۲۳", 123.0), ("१२३", 123.0), ("١٬٢٣٤٫٥", 1234.5), ("−5", -5.0),
         ("₹1,00,000", 100000.0)]


@pytest.mark.parametrize("typed,x", OTHER)
def test_other_notation_is_text_in_a_plain_cell(book, typed, x):
    """標準の表示形式で、数式も入力規則も使わないセル：書いたとおりの字（Excel に手で入力したときと同じ）。"""
    assert _put(book, 1, typed) == (typed, [])


@pytest.mark.parametrize("typed,x", OTHER)
def test_other_notation_is_a_number_in_a_number_cell_and_told(book, typed, x):
    v, notes = _put(book, 2, typed)
    assert v == pytest.approx(x) and len(notes) == 1
    assert "数としての読み方が 1 つに決まる" in notes[0] and "と表示されます" in notes[0] and "JSON の数" in notes[0]


@pytest.mark.parametrize("typed,x", OTHER)
def test_other_notation_is_a_number_where_a_formula_uses_the_cell(book, typed, x):
    v, notes = _put(book, 6, typed)
    assert v == pytest.approx(x) and notes


def test_other_notation_note_shows_the_cell_text(book):
    _v, notes = _put(book, 2, "1.234,56")
    assert "「1,234.56」と表示されます" in notes[0]
    v, notes = _put(book, 4, "12,5")  # パーセントの書式：1 以上は ％ として受け取る（ふつうの書き方の数と同じ）
    assert v == pytest.approx(0.125) and "「13%」" in notes[0]


@pytest.mark.parametrize("typed", ["١٢٣", "1,5", "1.234,56", "1 234"])
def test_other_notation_is_not_a_number_in_a_date_cell(book, typed):
    """日付のセルでは数にしない（数にすると、その数の日付として表示される）。"""
    v, notes = _put(book, 3, typed)
    assert v == typed and len(notes) == 1 and "日付として読めない" in notes[0]


# ====================================================================== 「1.234」「1,234」：地域で値が違う形
def test_region_shape_dot_is_told_with_the_other_reading(book):
    for row, shown in ((1, "1.234"), (2, "1.23"), (7, "1"), (9, "1")):
        v, notes = _put(book, row, "1.234")
        assert v == 1.234 and len(notes) == 1, (row, notes)
        assert "「.」を小数点、「,」を桁区切りとして" in notes[0] and f"「{shown}」と表示されます" in notes[0]
        assert "1234 のつもりなら、JSON の数 1234 で指定" in notes[0]
    v, notes = _put(book, 7, "10.000")
    assert v == 10.0 and "10000 のつもりなら、JSON の数 10000 で" in notes[0]
    v, notes = _put(book, 2, "-1.234")
    assert v == -1.234 and "-1234 のつもりなら、JSON の数 -1234 で" in notes[0]


def test_region_shape_is_silent_when_the_cell_shows_the_typed_marks(book):
    assert _put(book, 8, "1.234") == (1.234, [])  # 0.000 のセル：「1.234」と表示される
    assert _put(book, 8, "10.000") == (10.0, [])
    assert _put(book, 7, "1,234") == (1234.0, [])  # #,##0 のセル：「1,234」と表示される
    assert _put(book, 2, "1,234") == (1234.0, [])  # #,##0.00 のセル：「1,234.00」
    assert _put(book, 2, "-10,000") == (-10000.0, [])


def test_region_shape_comma_is_told_when_the_cell_shows_no_comma(book):
    v, notes = _put(book, 8, "1,234")
    assert v == 1234.0 and "「1234.000」と表示されます" in notes[0] and "1.234 のつもりなら、JSON の数 1.234 で" in notes[0]
    v, notes = _put(book, 9, "10,000")
    assert v == 10000.0 and "「10000」と表示されます" in notes[0]


def test_region_shape_alternative_follows_percent(book):
    v, notes = _put(book, 4, "1,234")  # パーセントの書式：1,234 → 1234% → 12.34
    assert v == pytest.approx(12.34) and "1.234 のつもりなら、JSON の数 0.01234 で" in notes[0]
    v, notes = _put(book, 2, "1.234%")
    assert v == pytest.approx(0.01234) and "1234% のつもりなら、JSON の数 12.34 で" in notes[0]


@pytest.mark.parametrize("typed,x", [("1234.5", 1234.5), ("1,234.56", 1234.56), ("1,234,567", 1234567.0), ("12.5", 12.5), ("0.125", 0.125),
                                     ("1234", 1234.0), ("1e3", 1000.0), ("-5", -5.0), ("－5", -5.0), ("12.3456", 12.3456)])
def test_plain_numbers_are_silent_in_a_number_cell(book, typed, x):
    assert _put(book, 2, typed) == (x, [])


def test_json_numbers_are_never_reinterpreted(book):
    assert _put(book, 1, 1.234) == (1.234, []) and _put(book, 7, 1234) == (1234.0, [])


# ====================================================================== 通貨記号・％
@pytest.mark.parametrize("typed,x", [("€100", 100.0), ("100 €", 100.0), ("100€", 100.0), ("£100", 100.0), ("₩1,000", 1000.0),
                                     ("$100", 100.0), ("¥1,000", 1000.0), ("-€100", -100.0), ("€ 100", 100.0), ("5%", 0.05), ("5 %", 0.05)])
def test_any_currency_sign_before_or_after(book, typed, x):
    assert _put(book, 7, typed) == (pytest.approx(x), [])
    v, notes = _put(book, 1, typed)  # 標準のセル：Excel なら表示形式を変える入力。書式は変えず、書いたとおりの字にして伝える
    assert v == typed and notes


@pytest.mark.parametrize("typed", ["(100)", "100-", "1000円", "3丁目", "€€100", "€100$", "1.234,5,6", "1,23,4", "012,345", "1 23"])
def test_not_a_number(book, typed):
    """かっこの負の数・後ろの符号・通貨の語は読まない：数の入るセルでは、字として書いたことを伝える。"""
    assert _put(book, 1, typed)[0] == typed
    v, notes = _put(book, 2, typed)
    assert v == typed and len(notes) == 1 and "数として読めない" in notes[0]
    assert _put(book, 6, typed)[0] == typed


def test_text_without_digits_in_a_number_cell_is_not_told(book):
    assert _put(book, 2, "なし") == ("なし", []) and _put(book, 2, "-") == ("-", [])
    assert _put(book, 5, "1.234,56") == ("1.234,56", [])  # 文字列の書式のセル：書いたとおり


# ====================================================================== 日付
@pytest.mark.parametrize("typed", ["2026-04-01", "2026/4/1", "2026年4月1日", "2026년 4월 1일", "2026년4월1일", "2026.04.01", "2026.04.01.",
                                   "2026. 4. 1.", "٢٠٢٦/٠٤/٠١", "۲۰۲۶-۰۴-۰۱", "２０２６年４月１日", "令和8年4月1日", "R8.4.1"])
def test_year_first_dates(book, typed):
    import datetime as dt

    assert parse_date_text(typed) == dt.date(2026, 4, 1)
    assert _put(book, 3, typed) == (APR1, [])
    v, notes = _put(book, 1, typed)  # 日付の表示形式でないセル：字として書いて伝える
    assert v == typed and "日付として読めます" in notes[0]


@pytest.mark.parametrize("typed", ["April 1, 2026", "1 avril 2026", "民國115年4月1日", "2026-04-01T10:00", "20260401x"])
def test_dates_not_read(book, typed):
    assert parse_date_text(typed) is None
    v, notes = _put(book, 3, typed)
    assert v == typed and "ISO" in notes[0]


def test_dates_before_the_first_day_are_text_and_told(book):
    assert _put(book, 3, "1900-01-01") == (1.0, [])
    for typed in ("1899-12-31", "1447-10-13", "0033-04-01"):
        v, notes = _put(book, 3, typed)
        assert v == typed and len(notes) == 1 and "1900 年 1 月 1 日から）より前" in notes[0] and "西暦に直して" in notes[0]


def test_dates_before_the_first_day_1904():
    from todoku.book import open_book

    wb = open_book(_book_bytes(mac1904=True), "m.xlsx")
    if not wb.date1904:
        pytest.skip("1904 年から数える設定が保存されていません")
    name = wb.sheets[0].name
    assert interpret_input(wb, name, 3, 1, "1904-01-01") == (0.0, [])
    v, notes = interpret_input(wb, name, 3, 1, "1903-12-31")
    assert v == "1903-12-31" and "1904 年 1 月 1 日から）より前" in notes[0]


# ====================================================================== OpenDocument の表・書き込みの結果
def test_ods_number_and_dates():
    from test_openformats import SHEET, make_ods
    from todoku.odf.sheet import OdsWorkbook
    from todoku.odf.sheetfill import fill

    blob = make_ods(SHEET)

    def content(values):
        res = fill(OdsWorkbook(blob, "a.ods"), values)
        return zipfile.ZipFile(io.BytesIO(res.blob)).read("content.xml").decode(), res.warnings

    c, w = content({"Sheet1!B3": "1.234,56", "Sheet1!B4": "2026년 4월 1일"})
    assert 'office:value="1234.56"' in c and 'office:date-value="2026-04-01"><text:p>2026/4/1</text:p>' in c
    assert any("数としての読み方が 1 つに決まる" in x and "「1,235」" in x for x in w)
    c, w = content({"Sheet1!B3": "1.234"})
    assert 'office:value="1.234"' in c and any("1234 のつもりなら" in x for x in w)
    # 1900 年より前の日付：LibreOffice は日付として扱う。見える字は、書式の道具で作らず、値の字（ISO）
    c, w = content({"Sheet1!B4": "1899-12-31"})
    assert 'office:date-value="1899-12-31"><text:p>1899-12-31</text:p>' in c and any("見える字を書いた値のまま" in x for x in w)
    assert not any("より前" in x for x in w)


def test_fill_tells_how_a_region_shaped_number_was_read(tmp_path):
    src = tmp_path / "n.xlsx"
    src.write_bytes(_book_bytes())
    name = api.fields(str(src))[0]["id"].split("!")[0]
    r = api.fill(str(src), {f"{name}!A7": "1.234", f"{name}!A2": "1.234,56", f"{name}!A8": "1.234"}, str(tmp_path / "out.xlsx"), force=True)
    assert r["proof_ok"] and r["read_back_ok"], r
    told = r["say_to_user"].split("本人に伝えること：", 1)[1]
    assert "1234 のつもりなら、JSON の数 1234" in told and "「1,234.56」と表示されます" in told
    assert f"{name}!A8: " not in told  # 0.000 のセルは「1.234」と表示される：読み方について伝えることはない
