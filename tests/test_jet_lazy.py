"""Access の大きな表を分けて読む（どの行にも届く・行を黙って落とさない）と、表の取り出しが返さなかった行を伝えること。"""

from __future__ import annotations

import struct

import jet_maker
import pytest
from todoku import api, lazy, limits, mcp_server
from todoku.book import open_book
from todoku.legacy.jet import Jet, cell_text
from todoku.legacy.tablebook import TableBook
from todoku.xlsx.tables import read as tread

NUM = {"name": "番号", "type": 4, "num": 0, "fixed": 0, "size": 4}
TXT = {"name": "字", "type": 10, "num": 1, "var": 0, "size": 255}
DBL = {"name": "値", "type": 7, "num": 2, "fixed": 4, "size": 8}
COLS = [NUM, TXT, DBL]
PER_PAGE = 40


def _n(i: int) -> bytes:
    return struct.pack("<i", i)


def _pages_table(m: jet_maker.Maker, name: str, cols: list[dict], rows: list[bytes]) -> int:
    """何ページにもまたがる表（Maker.add_table は 1 ページなので、ここで分ける）。"""
    tdef = len(m.pages)
    m.pages.append(None)
    usage = len(m.pages)
    m.pages.append(None)
    data = []
    for i in range(0, len(rows), PER_PAGE):
        data.append(len(m.pages))
        m.pages.append(m._data_page(tdef, rows[i:i + PER_PAGE]))
    m.pages[usage] = m._usage_page(data)[0]
    m.pages[tdef] = m._tdef(cols, len(rows), usage << 8)
    m.catalog.append((tdef, name, 1, 0))
    return tdef


def _big(jet3: bool = False, n: int = 1000, null_every: int = 0) -> tuple[bytes, list[dict]]:
    m = jet_maker.Maker(jet3)
    values = []
    rows = []
    for i in range(1, n + 1):
        v = {} if null_every and i % null_every == 0 else {"番号": _n(i), "字": m.text(f"字{i}"), "値": struct.pack("<d", i + 0.5)}
        values.append(v)
        rows.append(m.row(COLS, v))
    _pages_table(m, "大きい", COLS, rows)
    m.add_table("小さい", [NUM], [m.row([NUM], {"番号": _n(7)})])
    return m.build(), values


@pytest.fixture
def path(tmp_path, monkeypatch):
    lazy.clear()
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 10)  # 大きな書類として扱う（中身の作りは本物と同じ）

    def write(blob: bytes, name: str = "a.mdb") -> str:
        p = tmp_path / name
        p.write_bytes(blob)
        return str(p)

    yield write
    lazy.clear()


def _window(args: dict):
    r = mcp_server.call_tool("read_window", args)
    return r["content"][0]["text"], r["structuredContent"]


def _table(args: dict) -> dict:
    return mcp_server.call_tool("read_table", args)["structuredContent"]


@pytest.mark.parametrize("jet3", [False, True])
def test_open_gives_the_overview_and_no_rows(path, jet3):
    p = path(_big(jet3)[0])
    r = mcp_server.call_tool("open_document", {"path": p})
    s = r["structuredContent"]
    assert s["large"] is True
    one = next(x for x in s["sheets"] if x["name"] == "大きい")
    assert one["rows"] == 1000 and one["last_row"] == 1001 and one["columns"] == ["番号", "字", "値"]
    text = r["content"][1]["text"]
    assert "| 大きい | 1,000 | 2〜1001 | 3 |" in text and "字500" not in text


def test_large_table_is_not_graded_down(path):
    g = api.grade(path(_big()[0]))
    assert g["grade"] == "A" and g["write"]["grade"] == "D"
    assert g["provisional"] is True  # 全部の行の形は、検査まで未確認


@pytest.mark.parametrize("jet3", [False, True])
def test_read_window_returns_exactly_the_range(path, jet3):
    p = path(_big(jet3)[0])
    text, sc = _window({"path": p, "sheet": "大きい", "rows": "501-503"})
    lines = [x for x in text.split("\n") if x.startswith("- ") and x[2].isdigit()]
    assert lines == ["- 501: 500 ｜ 字500 ｜ 500.5", "- 502: 501 ｜ 字501 ｜ 501.5", "- 503: 502 ｜ 字502 ｜ 502.5"]
    assert sc["done"] is True
    assert "〔列の名前〕番号 ｜ 字 ｜ 値" in text


@pytest.mark.parametrize("jet3", [False, True])
def test_read_table_returns_exactly_the_range(path, jet3):
    p = path(_big(jet3)[0])
    res = _table({"path": p, "header": "番号", "sheet": "大きい", "rows": "1000-1001"})
    assert [(r["_row"], r["番号"], r["字"], r["値"]) for r in res["all_rows"]] == [(1000, "999", "字999", "999.5"), (1001, "1000", "字1000", "1000.5")]
    assert res["rows"] == 2 and not res.get("next") and res["omitted_rows"] == 0


def test_range_decodes_only_its_pages(path, monkeypatch):
    p = path(_big()[0])
    doc = api.lazy_open(p, force=True)
    seen = []
    real = Jet.page
    monkeypatch.setattr(Jet, "page", lambda self, n: (seen.append(n), real(self, n))[1])
    doc.db._cache.clear()
    page, _ = api.read_range(p, {"sheet": "大きい", "rows": "900-901"})
    assert len(page.body) == 2
    assert len(set(seen)) <= 2  # 頼まれた行のあるページだけ


@pytest.mark.parametrize("jet3", [False, True])
def test_cursor_reaches_every_row_of_every_table(path, jet3):
    p = path(_big(jet3)[0])
    rng: dict = {}
    got: list[str] = []
    for _ in range(500):
        page, _d = api.read_range(p, rng, 1000)
        got += page.body
        if page.next is None:
            break
        rng = page.next
    assert page.next is None
    assert sum(1 for x in got if "字" in x and x.split(": ")[1].split(" ")[0].isdigit()) == 1000
    assert got[-1] == "- 2: 7"


def test_read_table_cursor_never_skips_rows_with_the_default_limit(path):
    p = path(_big()[0])
    args = {"path": p, "header": "番号", "sheet": "大きい"}
    rows: list[int] = []
    for _ in range(20):
        res = _table(args)
        rows += [r["_row"] for r in res["all_rows"]]
        if not res.get("next"):
            break
        args = {**args, **res["next"]}
    assert rows == list(range(2, 1002))


def test_out_path_writes_every_row(path, tmp_path):
    p = path(_big()[0])
    dst = tmp_path / "out.csv"
    res = _table({"path": p, "header": "番号", "sheet": "大きい", "out_path": str(dst)})
    lines = dst.read_text(encoding="utf-8").splitlines()
    assert res["rows"] == 1000 and len(lines) == 1001 and lines[-1] == "大きい,1001,1000,字1000,1000.5"


# ---------------------------------------------------------------- すべて NULL の行
@pytest.mark.parametrize("large", [True, False])
def test_all_null_records_are_kept(tmp_path, monkeypatch, large):
    lazy.clear()
    if large:
        monkeypatch.setattr(limits, "CONTEXT_BUDGET", 100)
    m = jet_maker.Maker(False)
    rows = [m.row(COLS, {"番号": _n(1)}), m.row(COLS, {}), m.row(COLS, {}), m.row(COLS, {"番号": _n(4)}), m.row(COLS, {})]
    m.add_table("表", COLS, rows)
    p = tmp_path / "a.mdb"
    p.write_bytes(m.build())
    res = _table({"path": str(p), "header": "番号", "sheet": "表"})
    assert [r["_row"] for r in res["all_rows"]] == [2, 3, 4, 5, 6]
    assert res["all_rows"][1].get("番号") is None and res["omitted_rows"] == 0
    page, _ = api.read_range(str(p), {"sheet": "表", "rows": "3-3"})
    assert page.body == ["- 3: 〔NULL〕 ｜ 〔NULL〕 ｜ 〔NULL〕"]
    if not large:
        whole = api.window(str(p))[2][0][1]
        assert "- 3: 〔NULL〕 ｜ 〔NULL〕 ｜ 〔NULL〕" in whole and "- 6: 〔NULL〕" in whole
    lazy.clear()


def test_trailing_all_null_rows_are_counted_in_the_sheet():
    m = jet_maker.Maker(False)
    m.add_table("表", COLS, [m.row(COLS, {"番号": _n(1)}), m.row(COLS, {})])
    sh = open_book(m.build(), "a.mdb").sheets[0]
    assert sh.max_row == 3 and sh.records == [(2, 3)]


# ---------------------------------------------------------------- 字と数をそのまま
@pytest.mark.parametrize("value,shown", [("末尾の空白 ", '"末尾の空白 "'), (" 頭", '" 頭"'), ("全角　", '"全角　"'), ("改\n行", '"改\\n行"'), ("a｜b", '"a｜b"'),
                                         ("", '""'), ("〔NULL〕", '"〔NULL〕"'), ('"引', '"\\"引"'), ("ふつう", "ふつう"), ("中 の 空白", "中 の 空白")])
def test_text_is_shown_unambiguously(value, shown):
    assert cell_text(value) == shown


def test_text_cells_keep_their_spaces(path):
    m = jet_maker.Maker(False)
    m.add_table("表", COLS, [m.row(COLS, {"番号": _n(1), "字": m.text("末尾の空白  ")})])
    p = path(m.build())
    page, _ = api.read_range(p, {"sheet": "表", "rows": "2"})
    assert page.body == ['- 2: 1 ｜ "末尾の空白  " ｜ 〔NULL〕']
    res = _table({"path": p, "header": "番号", "sheet": "表", "rows": "2"})
    assert res["all_rows"][0]["字"] == "末尾の空白  "


def test_empty_text_is_not_null(path):
    m = jet_maker.Maker(False)
    m.add_table("表", COLS, [m.row(COLS, {"番号": _n(1), "字": b""}), m.row(COLS, {"番号": _n(2)})])
    p = path(m.build())
    assert api.read_range(p, {"sheet": "表", "rows": "2-3"})[0].body == ['- 2: 1 ｜ "" ｜ 〔NULL〕', "- 3: 2 ｜ 〔NULL〕 ｜ 〔NULL〕"]
    rows = _table({"path": p, "header": "番号", "sheet": "表", "rows": "2-3"})["all_rows"]
    assert rows[0]["字"] == "" and "字" not in rows[1]


@pytest.mark.parametrize("v,text", [(0.123456789012345, "0.123456789012345"), (1234567890123456.0, "1234567890123456"), (1e22, "1e+22"), (-2.5, "-2.5"), (3.0, "3")])
def test_numbers_keep_every_digit(tmp_path, v, text):
    lazy.clear()
    m = jet_maker.Maker(False)
    m.add_table("表", COLS, [m.row(COLS, {"番号": _n(1), "値": struct.pack("<d", v)})])
    p = tmp_path / "a.mdb"
    p.write_bytes(m.build())
    assert _table({"path": str(p), "header": "番号", "sheet": "表"})["all_rows"][0]["値"] == text
    assert f"- 2: 1 ｜ 〔NULL〕 ｜ {text}" in api.window(str(p))[2][0][1]
    lazy.clear()


# ---------------------------------------------------------------- 形を確かめられない行
def _with_bad_row() -> bytes:
    m = jet_maker.Maker(False)
    bad = bytearray(m.row(COLS, {"番号": _n(2)}))
    bad[0:2] = struct.pack("<H", 0xFFFF)
    m.add_table("表", COLS, [m.row(COLS, {"番号": _n(1)}), bytes(bad), m.row(COLS, {"番号": _n(3)})])
    return m.build()


def test_bad_row_keeps_its_number_in_every_way_of_reading(path, tmp_path):
    blob = _with_bad_row()
    p = path(blob)
    page, _ = api.read_range(p, {"sheet": "表", "rows": "3-4"})
    assert page.body[0].startswith("- 3: 〔この行は形を確かめられないので読んでいません") and page.body[1].startswith("- 4: 3 ｜")
    res = _table({"path": p, "header": "番号", "sheet": "表"})
    assert [r["_row"] for r in res["all_rows"]] == [2, 3, 4] and "_unread" in res["all_rows"][1]
    sh = open_book(blob, "a.mdb").sheets[0]
    assert list(sh.bad_rows) == [3] and sh.cells[(4, 1)].value == 3


def test_audit_finds_bad_rows_and_grade_reports_them(path):
    p = path(_with_bad_row())
    doc = api.lazy_open(p, force=True)
    state = doc.audit_begin()
    while not state["done"]:
        state = doc.audit_step(state)
    check = state["checks"][0]
    assert check["ok"] is False and "3 行目" in check["items"][0]


def test_audit_passes_on_a_clean_table(path):
    doc = api.lazy_open(path(_big(n=200)[0]), force=True)
    state = doc.audit_begin()
    while not state["done"]:
        state = doc.audit_step(state)
    assert state["checks"][0]["ok"] is True and state["rows"] == 201


# ---------------------------------------------------------------- そのほか
def test_writing_is_refused_without_reading_the_whole_file(path, tmp_path):
    p = path(_big()[0])
    with pytest.raises(api.ApiError, match="Access のデータベースには書きません"):
        api.fill(p, {"大きい!A2": 1}, str(tmp_path / "b.mdb"))


def test_header_is_a_column_name_not_a_value(path):
    p = path(_big()[0])
    with pytest.raises(api.ApiError, match="見つかりません"):
        api.read_table(p, "字500", "大きい")


def test_unknown_table_and_rows_without_sheet_are_refused(path):
    p = path(_big()[0])
    with pytest.raises(api.ApiError, match="シート「丙」はありません"):
        api.read_range(p, {"sheet": "丙"})
    with pytest.raises(api.ApiError, match="sheet"):
        api.read_range(p, {"rows": "2-3"})
    with pytest.raises(api.ApiError, match="シート「丙」がありません"):
        api.read_table(p, "番号", "丙")
    with pytest.raises(api.ApiError, match="sheet"):
        api.read_table(p, "番号", rows="2-3")


def test_range_past_the_end_says_so(path):
    text, sc = _window({"path": path(_big()[0]), "sheet": "大きい", "rows": "5000-5001"})
    assert "1,001 行目まで" in text and sc["done"] is True


def test_one_long_row_is_split_and_can_be_rejoined(path):
    m = jet_maker.Maker(False)
    memo = "あ" * 1000
    m.add_table("表", COLS, [m.row(COLS, {"番号": _n(1), "字": m.text(memo[:255])}), m.row(COLS, {"番号": _n(2), "字": m.text("い" * 255)})] * 2)
    p = path(m.build())
    rng: dict = {"sheet": "表", "rows": "2-3"}
    got = ""
    for _ in range(50):
        page, _d = api.read_range(p, rng, 200)
        got += "".join(page.body)
        if page.next is None:
            break
        rng = page.next
    assert page.next is None and got.count("あ") == 255 and got.count("い") == 255


def test_password_is_asked_on_the_range_path_too(path):
    from todoku.legacy.jet import _rc4

    blob = bytearray(jet_maker.sample(True))  # Jet3：パスワードの欄は 0x42 からの字。0x18 から 126 バイトが RC4 で隠されている
    key = struct.pack("<I", 0x6B39DAC7)
    head = bytearray(_rc4(key, bytes(blob[0x18:0x18 + 126])))
    head[0x42 - 0x18:0x42 - 0x18 + 6] = b"secret"
    blob[0x18:0x18 + 126] = _rc4(key, bytes(head))
    p = path(bytes(blob))
    with pytest.raises(api.PasswordNeeded):
        api.read_range(p, {"sheet": "コード表"})
    with pytest.raises(api.PasswordWrong):
        api.read_range(p, {"sheet": "コード表"}, password="違う")
    assert api.read_range(p, {"sheet": "コード表"}, password="secret")[0].body


def test_encrypted_pages_are_counted_and_read(path):
    from test_jet_real_shapes import _page_key

    blob, _v = _big(n=300)
    p = path(_page_key(blob, 4096, 0xA7B773E0))
    page, _ = api.read_range(p, {"sheet": "大きい", "rows": "300-300"})
    assert page.body == ["- 300: 299 ｜ 字299 ｜ 299.5"]


def test_member_of_a_zip_is_read_in_ranges(path, tmp_path):
    import zipfile

    z = tmp_path / "a.zip"
    with zipfile.ZipFile(z, "w") as f:
        f.writestr("d/表.mdb", _big()[0])
    p = f"{z}!/d/表.mdb"
    assert api.summary(p)["large"] is True
    assert api.read_range(p, {"sheet": "大きい", "rows": "1001"})[0].body == ["- 1001: 1000 ｜ 字1000 ｜ 1000.5"]
    assert _table({"path": p, "header": "番号", "sheet": "大きい", "rows": "1001"})["all_rows"][0]["字"] == "字1000"


# ---------------------------------------------------------------- シート：返さなかった行を伝える
def _sheet_book(rows: dict[int, dict[int, str]]) -> TableBook:
    wb = TableBook(b"", "a", "wk1")
    sh = wb.add_sheet("表")
    for r, cells in rows.items():
        for c, v in cells.items():
            wb.put(sh, r, c, v)
    return wb


def test_sheet_reports_empty_rows_by_number():
    wb = _sheet_book({1: {1: "番号", 2: "名前"}, 2: {1: "1"}, 5: {2: "x"}, 8: {1: "2"}})
    t = tread(wb, "番号", "表")
    assert [r["_row"] for r in t["rows"]] == [2, 5, 8]
    om = t["sheets"][0]["omitted_rows"]
    assert om["count"] == 4 and om["no_values"] == {"count": 4, "rows": [[3, 4], [6, 7]], "what": om["no_values"]["what"]}
    assert t["omitted_rows"] == om["count"]


def test_sheet_reports_rows_with_values_only_outside_the_header_columns():
    wb = _sheet_book({1: {1: "番号"}, 2: {1: "1"}, 3: {5: "欄外"}, 4: {1: "2"}})
    t = tread(wb, "番号", "表")
    assert [r["_row"] for r in t["rows"]] == [2, 4]
    om = t["sheets"][0]["omitted_rows"]
    assert om["values_outside_columns"] == {"count": 1, "rows": [[3, 3]], "what": om["values_outside_columns"]["what"]} and "no_values" not in om


def test_sheet_trailing_used_range_is_reported():
    wb = _sheet_book({1: {1: "番号"}, 2: {1: "1"}})
    wb.sheets[0].max_row = 6  # 書式だけの行まで使っている範囲
    om = tread(wb, "番号", "表")["sheets"][0]["omitted_rows"]
    assert om["no_values"]["rows"] == [[3, 6]] and om["count"] == 4


def test_sheet_without_gaps_reports_nothing():
    t = tread(_sheet_book({1: {1: "番号"}, 2: {1: "1"}, 3: {1: "2"}}), "番号", "表")
    assert "omitted_rows" not in t["sheets"][0] and t["omitted_rows"] == 0


def test_blank_rows_reported_the_same_in_big_and_small_xlsx(tmp_path, monkeypatch):
    openpyxl = pytest.importorskip("openpyxl")
    lazy.clear()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表"
    ws["A1"] = "番号"
    for r in list(range(2, 2000)) + list(range(2100, 3000)):
        ws.cell(r, 1, r)
    ws.cell(2050, 3, "欄外")
    p = tmp_path / "a.xlsx"
    wb.save(p)
    small = api.read_table(str(p), "番号", "表")
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 100)
    lazy.clear()
    assert api.lazy_open(str(p), built_ok=False) is not None
    big = api.read_table(str(p), "番号", "表", limit=100000)
    for res in (small, big):
        om = res["sheets"][0]["omitted_rows"]
        assert om["no_values"]["rows"] == [[2000, 2049], [2051, 2099]]
        assert om["values_outside_columns"]["rows"] == [[2050, 2050]]
        assert res["rows"] == 2898
    lazy.clear()


def test_rows_range_on_a_small_book(tmp_path):
    openpyxl = pytest.importorskip("openpyxl")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表"
    ws["A1"] = "番号"
    for r in range(2, 50):
        ws.cell(r, 1, r)
    p = tmp_path / "a.xlsx"
    wb.save(p)
    res = _table({"path": str(p), "header": "番号", "sheet": "表", "rows": "10-11"})
    assert [r["_row"] for r in res["all_rows"]] == [10, 11]
