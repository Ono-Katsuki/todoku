"""Access の読み手：実物（公開の試験データ）で見つかった形を、自分で作ったファイルで確かめる。実物のファイルはリポジトリに入れない。"""

from __future__ import annotations

import struct

import jet_maker
import pytest
from todoku import api, mcp_server
from todoku.book import open_book
from todoku.legacy import jet
from todoku.legacy.jet import _date, _rc4


def _sheet(blob: bytes, name: str | None = None):
    wb = open_book(blob, "a.mdb")
    return wb, (wb.sheets[0] if name is None else next(s for s in wb.sheets if s.name == name))


def _rows(sh) -> list[list]:
    """読めた行（形を確かめられない行は番号を飛ばさないが、値はない）。"""
    return [[sh.cells[(r, c)].value if (r, c) in sh.cells else None for c in range(1, sh.max_col + 1)] for r in range(1, sh.max_row + 1) if r not in sh.bad_rows]


def _p(tmp_path, name: str, blob: bytes) -> str:
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


NUM = {"name": "番号", "type": 4, "num": 0, "fixed": 0, "size": 4}
TXT = {"name": "字", "type": 10, "num": 1, "var": 0, "size": 255}


def _table(jet3: bool, values: list[dict], cols=None) -> tuple[jet_maker.Maker, list[bytes], int]:
    m = jet_maker.Maker(jet3)
    cols = cols or [NUM, TXT]
    rows = [m.row(cols, v) for v in values]
    return m, rows, m.add_table("表", cols, rows)


def _n(i: int) -> bytes:
    return struct.pack("<i", i)


# ---------------------------------------------------------------- 列の 2 つの番号
@pytest.mark.parametrize("jet3", [False, True])
def test_deleted_column_null_bit_is_stable_id_not_position(jet3):
    """列を消すと並びは詰まるが、行の NULL の印のビットと位置の表の番号は元のまま。"""
    cols = [{"name": "A", "type": 10, "num": 0, "pos": 0, "var": 0, "size": 50},
            {"name": "C", "type": 10, "num": 2, "pos": 1, "var": 2, "size": 50}]
    m = jet_maker.Maker(jet3)
    # 消した列（番号 1）の字が行に残っている。C は NULL の行と、値のある行
    gone = [dict(cols[0], name="B", num=1, pos=1, var=1), dict(cols[1], pos=2)]
    full = [cols[0]] + gone
    r1 = m.row(full, {"A": m.text("a1"), "B": m.text("消えた列")})
    r2 = m.row(full, {"A": m.text("a2"), "C": m.text("c2")})
    m.add_table("表", cols, [r1, r2])
    _wb, sh = _sheet(m.build())
    assert _rows(sh) == [["A", "C"], ["a1", None], ["a2", "c2"]]


@pytest.mark.parametrize("jet3", [False, True])
def test_columns_ordered_by_position_then_id(jet3):
    m = jet_maker.Maker(jet3)
    cols = [{"name": "二", "type": 4, "num": 1, "pos": 2, "fixed": 4, "size": 4}, {"name": "一", "type": 4, "num": 0, "pos": 1, "fixed": 0, "size": 4}]
    m.add_table("表", cols, [m.row(cols, {"一": _n(1), "二": _n(2)})])
    assert _rows(_sheet(m.build())[1]) == [["一", "二"], [1, 2]]
    # システムの表のように並びの番号がすべて 0 なら、列の番号の順
    m = jet_maker.Maker(jet3)
    cols = [{"name": "後", "type": 4, "num": 1, "pos": 0, "fixed": 4, "size": 4}, {"name": "前", "type": 4, "num": 0, "pos": 0, "fixed": 0, "size": 4}]
    m.add_table("表", cols, [m.row(cols, {"前": _n(1), "後": _n(2)})])
    assert _rows(_sheet(m.build())[1]) == [["前", "後"], [1, 2]]


# ---------------------------------------------------------------- 行の位置
def _page(m: jet_maker.Maker, tdef: int, placed: list[tuple[bytes, int, int]]) -> bytes:
    """[(行, 位置, 印)] をそのままの位置に置いたページ。"""
    p = bytearray(m.pg)
    p[0], p[1] = 1, 1
    struct.pack_into("<I", p, 4, tdef)
    rc = 8 if m.jet3 else 12
    struct.pack_into("<H", p, rc, len(placed))
    for i, (r, at, fl) in enumerate(placed):
        p[at:at + len(r)] = r
        struct.pack_into("<H", p, rc + 2 + 2 * i, at | fl)
    return bytes(p)


@pytest.mark.parametrize("jet3", [False, True])
def test_row_offsets_need_not_descend(jet3):
    m, rows, tdef = _table(jet3, [{"番号": _n(1), "字": b"a" * 40}, {"番号": _n(2), "字": b"b" * 10}])
    a, b = rows
    # 1 つ目の行が、2 つ目の行より手前にある（行を書き直すとこうなる）
    m.pages[tdef + 2] = _page(m, tdef, [(a, m.pg - len(a) - len(b), 0), (b, m.pg - len(b), 0)])
    got = _rows(_sheet(m.build())[1])[1:]
    assert got == [[1, ("a" * 40 if jet3 else "慡" * 20)], [2, ("b" * 10 if jet3 else "扢" * 5)]]


@pytest.mark.parametrize("jet3", [False, True])
def test_empty_deleted_row_shares_position(jet3):
    m, rows, tdef = _table(jet3, [{"番号": _n(1), "字": m_text(jet3, "残る")}])
    r = rows[0]
    at = m.pg - len(r)
    m.pages[tdef + 2] = _page(m, tdef, [(b"", at, 0x8000), (r, at, 0)])
    assert _rows(_sheet(m.build())[1])[1:] == [[1, "残る"]]


def m_text(jet3: bool, s: str) -> bytes:
    return jet_maker.Maker(jet3).text(s)


@pytest.mark.parametrize("jet3", [False, True])
def test_pointer_row_is_followed_and_target_not_counted_twice(jet3):
    m, rows, tdef = _table(jet3, [{"番号": _n(1), "字": m_text(jet3, "ここ")}, {"番号": _n(2), "字": m_text(jet3, "移った")}])
    moved = len(m.pages)
    # 移った先のページは、この表のもの。行き先の行には消えた行と同じ印がある
    m.pages.append(_page(m, tdef, [(rows[1], m.pg - len(rows[1]), 0x8000)]))
    pointer = struct.pack("<I", moved << 8)
    a = rows[0]
    m.pages[tdef + 2] = _page(m, tdef, [(a, m.pg - len(a), 0), (pointer, m.pg - len(a) - 4, 0x4000)])
    m.pages[tdef + 1] = m._usage_page([tdef + 2, moved])[0]
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [[1, "ここ"], [2, "移った"]]
    assert not sh.unread


def test_pointer_to_nothing_is_reported():
    m, rows, tdef = _table(False, [{"番号": _n(1), "字": m_text(False, "ここ")}])
    a = rows[0]
    m.pages[tdef + 2] = _page(m, tdef, [(a, m.pg - len(a), 0), (struct.pack("<I", tdef + 2 << 8 | 9), m.pg - len(a) - 4, 0x4000)])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [[1, "ここ"]]
    assert any("1 行は形を確かめられない" in u and "表「表」" in u for u in sh.unread)


# ---------------------------------------------------------------- 位置の表
@pytest.mark.parametrize("jet3", [False, True])
def test_shrunk_row_has_gap_before_offset_table(jet3):
    """書き直して短くなった行：データの終わりと位置の表の間に余りがある。"""
    m = jet_maker.Maker(jet3)
    cols = [NUM, TXT]
    r = m.row(cols, {"番号": _n(1), "字": m.text("短い")})
    tail = (2 + 1 + 1) if jet3 else (4 + 2 + 1)  # 位置 2 つ・数・印 1 バイト
    m.add_table("表", cols, [r[:-tail] + b"\0\0" + r[-tail:]])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [[1, "短い"]] and not sh.unread


def test_jet3_carry_table_filler():
    """データの終わりがちょうど 255 の行：繰り上がりの表に 0xFF の詰め物が 1 つある。"""
    m = jet_maker.Maker(True)
    cols = [NUM, TXT]
    text = b"x" * 250  # 見出し 1 + 固定 4 + 250 = 255
    r = m.row(cols, {"番号": _n(7), "字": text})
    r = r[:-2] + b"\xff" + r[-2:]  # 数の直前に詰め物
    m.add_table("表", cols, [r])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [[7, "x" * 250]] and not sh.unread


@pytest.mark.parametrize("jet3", [False, True])
def test_broken_offset_table_reported_not_guessed(jet3):
    m = jet_maker.Maker(jet3)
    cols = [NUM, TXT]
    good = m.row(cols, {"番号": _n(1), "字": m.text("正しい行")})
    bad = bytearray(m.row(cols, {"番号": _n(2), "字": m.text("壊れた行")}))
    if jet3:
        bad[-4], bad[-3] = 0x05, 0xF0  # 位置が減る順になり、データの終わりも外れる
    else:
        bad[-7:-3] = struct.pack("<HH", 0x0FF0, 0x0006)
    m.add_table("表", cols, [good, bytes(bad)])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [[1, "正しい行"]]
    note = "".join(sh.unread)
    assert "1 行は形を確かめられない" in note and "mdb-export" in note


def test_row_with_too_many_columns_for_its_length_reported():
    m = jet_maker.Maker(False)
    cols = [NUM, TXT]
    bad = bytearray(m.row(cols, {"番号": _n(2), "字": m.text("行")}))
    bad[0:2] = struct.pack("<H", 0xFFFF)
    m.add_table("表", cols, [bytes(bad)])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [] and list(sh.bad_rows) == [2] and any("形を確かめられない" in u for u in sh.unread)


def test_missing_data_page_is_reported():
    m, _rows_, tdef = _table(False, [{"番号": _n(1), "字": m_text(False, "ある")}])
    m.pages[tdef + 1] = m._usage_page([tdef + 2, 200])[0]  # 200 ページ目はファイルにない
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [[1, "ある"]] and any("ファイルの外" in u for u in sh.unread)


def test_stated_row_count_mismatch_is_said():
    m, rows, tdef = _table(False, [{"番号": _n(1), "字": m_text(False, "あ")}])
    m.pages[tdef] = m._tdef([NUM, TXT], 5, (tdef + 1) << 8)
    _wb, sh = _sheet(m.build())
    assert any("行の数は 5" in u and "読めた行は 1" in u for u in sh.unread)


def test_broken_catalog_row_refuses_the_whole_list():
    """表の一覧が欠けたまま「表はこれだけ」と言わない。"""
    m = jet_maker.Maker(False)
    m.add_table("表", [NUM, TXT], [m.row([NUM, TXT], {"番号": _n(1), "字": m.text("あ")})])
    blob = bytearray(m.build())
    data = next(i for i in range(3, len(blob) // 4096) if blob[i * 4096] == 1 and struct.unpack_from("<I", blob, i * 4096 + 4)[0] == 2)
    off = struct.unpack_from("<H", blob, data * 4096 + 14)[0] & 0x1FFF
    blob[data * 4096 + off:data * 4096 + off + 2] = struct.pack("<H", 0xFFFF)
    with pytest.raises(jet.JetError, match="表の一覧"):
        open_book(bytes(blob), "a.mdb")


# ---------------------------------------------------------------- 値
def _decimal(v: int, neg: bool = False) -> bytes:
    groups = [(v >> 96) & 0xFFFFFFFF, (v >> 64) & 0xFFFFFFFF, (v >> 32) & 0xFFFFFFFF, v & 0xFFFFFFFF]
    return bytes([0x80 if neg else 0]) + struct.pack("<4I", *groups)


def test_decimal_precision_comes_before_scale():
    m = jet_maker.Maker(False)
    cols = [{"name": "数", "type": 16, "num": 0, "fixed": 0, "size": 17, "prec": 18, "scale": 6}]
    m.add_table("表", cols, [m.row(cols, {"数": _decimal(50325000)}), m.row(cols, {"数": _decimal(36222200, True)}), m.row(cols, {"数": _decimal(102030405060654321)})])
    assert [r[0] for r in _rows(_sheet(m.build())[1])[1:]] == [50.325, -36.2222, "102030405060.654321"]  # 倍精度に入らない桁は字で


def test_currency_keeps_every_digit():
    m = jet_maker.Maker(False)
    cols = [{"name": "額", "type": 5, "num": 0, "fixed": 0, "size": 8}]
    vals = [12345, -8345737500, 2 ** 63 - 1, -(2 ** 63)]
    m.add_table("表", cols, [m.row(cols, {"額": struct.pack("<q", v)}) for v in vals])
    assert [r[0] for r in _rows(_sheet(m.build())[1])[1:]] == [1.2345, -834573.75, "922337203685477.5807", "-922337203685477.5808"]


def test_single_shows_the_short_decimal():
    m = jet_maker.Maker(False)
    cols = [{"name": "単", "type": 6, "num": 0, "fixed": 0, "size": 4}]
    m.add_table("表", cols, [m.row(cols, {"単": struct.pack("<f", x)}) for x in (0.1, 804983.375, -0.00035134, 1e30)])
    assert [r[0] for r in _rows(_sheet(m.build())[1])[1:]] == [0.1, 804983.4, -0.00035134, 1e30]


@pytest.mark.parametrize("value,text", [(0.0, "1899-12-30"), (45000.0, "2023-03-15"), (45000.5, "2023-03-15 12:00:00"), (-1.5, "1899-12-29 12:00:00"),
                                        (-0.5, "1899-12-30 12:00:00"), (2958101.0, "9999-01-01"), (-98585.0, "1630-01-29"),
                                        (-657434.0, "0100-01-01"), (float("nan"), "nan"), (1e30, "1e+30")])
def test_dates(value, text):
    assert _date(value) == text


def test_ole_says_its_stored_length_and_empty_binary_is_empty():
    m = jet_maker.Maker(False)
    cols = [NUM, {"name": "絵", "type": 11, "num": 1, "var": 0, "size": 0}, {"name": "生", "type": 9, "num": 2, "var": 1, "size": 16}]
    head = struct.pack("<III", 70000 | 0x40000000, 0x1234, 0)
    m.add_table("表", cols, [m.row(cols, {"番号": _n(1), "絵": head, "生": b""})])
    _wb, sh = _sheet(m.build())
    assert sh.cells[(2, 2)].value == "〔OLE：70000 バイト。中身は読みません〕" and (2, 3) not in sh.cells
    assert any("表「表」" in u and "絵" in u for u in sh.unread)


def _ext(days: int, ticks: int) -> bytes:
    return f"{days:019d}:{ticks:019d}:7\0".encode("ascii")


def test_extended_date():
    m = jet_maker.Maker(False)
    cols = [{"name": "拡", "type": 20, "num": 0, "fixed": 0, "size": 42}]
    m.add_table("表", cols, [m.row(cols, {"拡": _ext(737592, 0)}), m.row(cols, {"拡": _ext(737954, 459000000000)}), m.row(cols, {"拡": _ext(737954, 819123456789)}),
                             m.row(cols, {"拡": _ext(36323, 459001234567)})])
    assert [r[0] for r in _rows(_sheet(m.build())[1])[1:]] == ["2020-06-17", "2021-06-14 12:45:00", "2021-06-14 22:45:12.3456789", "0100-06-14 12:45:00.1234567"]


def test_extended_date_broken_is_reported():
    m = jet_maker.Maker(False)
    cols = [NUM, {"name": "拡", "type": 20, "num": 1, "fixed": 4, "size": 42}]
    m.add_table("表", cols, [m.row(cols, {"番号": _n(1), "拡": b"x" * 42}), m.row(cols, {"番号": _n(2), "拡": _ext(737592, 864_000_000_000)}), m.row(cols, {"番号": _n(3), "拡": _ext(737592, 0)})])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [[3, "2020-06-17"]] and any("2 行は形を確かめられない" in u for u in sh.unread)


# ---------------------------------------------------------------- 計算列
def _calc(body: bytes) -> bytes:
    return bytes(16) + struct.pack("<I", len(body)) + body + bytes(3)


def _calc_cols() -> list[dict]:
    return [NUM, {"name": "氏名", "type": 10, "num": 1, "var": 0, "size": 509, "flags2": 0xC0},
            {"name": "月額", "type": 7, "num": 2, "var": 1, "size": 39, "flags2": 0xC0}]


def test_calculated_text_is_unframed_and_number_is_not_guessed():
    m = jet_maker.Maker(False)
    cols = _calc_cols()
    # 倍精度と宣言された列に、8 バイトの通貨の形が入っている（実物にある）
    m.add_table("表", cols, [m.row(cols, {"番号": _n(1), "氏名": _calc(b"\xff\xfeWayne, Bruce"), "月額": _calc(struct.pack("<q", 833333333))})])
    _wb, sh = _sheet(m.build())
    assert sh.cells[(2, 2)].value == "Wayne, Bruce"
    assert "読みません" in sh.cells[(2, 3)].value and not isinstance(sh.cells[(2, 3)].value, float)
    assert any("計算列（月額）" in u for u in sh.unread)


def test_calculated_long_text_in_memo_form():
    m = jet_maker.Maker(False)
    cols = [NUM, {"name": "長", "type": 10, "num": 1, "var": 0, "size": 0, "flags2": 0xC0}]
    body = _calc("長い字".encode("utf-16-le"))
    lval = len(m.pages)
    m.pages.append(m._data_page(0, [body]))
    inline = struct.pack("<III", len(body) | 0x80000000, 0, 0) + body
    outside = struct.pack("<III", len(body) | 0x40000000, lval << 8, 0)
    m.add_table("表", cols, [m.row(cols, {"番号": _n(1), "長": inline}), m.row(cols, {"番号": _n(2), "長": outside})])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [[1, "長い字"], [2, "長い字"]]


@pytest.mark.parametrize("raw", [b"\xff\xfeabc", bytes(16) + struct.pack("<I", 99) + b"ab" + bytes(3), b"\x01" * 16 + struct.pack("<I", 2) + b"ab" + bytes(3)])
def test_calculated_frame_broken_is_reported(raw):
    m = jet_maker.Maker(False)
    cols = _calc_cols()
    m.add_table("表", cols, [m.row(cols, {"番号": _n(1), "氏名": raw})])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [] and list(sh.bad_rows) == [2] and any("形を確かめられない" in u for u in sh.unread)


# ---------------------------------------------------------------- メモ
def _memo_table(jet3: bool, head_len: int, pieces: list[tuple[int, bytes]]) -> bytes:
    """pieces：(次の行の番号（-1 で終わり）, 字)。別のページに続きの行を置く。"""
    m = jet_maker.Maker(jet3)
    lval = len(m.pages)
    rows = [struct.pack("<I", (lval << 8 | nxt) if nxt >= 0 else 0) + body for nxt, body in pieces]
    m.pages.append(m._data_page(0, rows))
    cols = [NUM, {"name": "メモ", "type": 12, "num": 1, "var": 0, "size": 0}]
    head = struct.pack("<III", head_len, lval << 8, 0)
    m.add_table("表", cols, [m.row(cols, {"番号": _n(1), "メモ": head})])
    return m.build()


@pytest.mark.parametrize("jet3", [False, True])
def test_memo_chain_read_to_its_length(jet3):
    t = jet_maker.Maker(jet3).text("あいうえおかきくけこ")
    blob = _memo_table(jet3, len(t), [(1, t[:6]), (2, t[6:12]), (-1, t[12:] + b"\0\0\0\0")])
    assert _rows(_sheet(blob)[1])[1:] == [[1, "あいうえおかきくけこ"]]


def test_memo_chain_short_is_reported_not_truncated():
    t = jet_maker.Maker(False).text("あいうえお")
    blob = _memo_table(False, len(t) + 100, [(-1, t)])
    _wb, sh = _sheet(blob)
    assert _rows(sh)[1:] == [] and list(sh.bad_rows) == [2] and any("メモの続きが途中で切れて" in u for u in sh.unread)


def test_memo_chain_loop_is_reported():
    t = jet_maker.Maker(False).text("あ")
    blob = _memo_table(False, 1000, [(0, t)])
    _wb, sh = _sheet(blob)
    assert _rows(sh)[1:] == [] and list(sh.bad_rows) == [2] and any("輪" in u for u in sh.unread)


def test_memo_single_shorter_than_stated_is_reported():
    m = jet_maker.Maker(False)
    lval = len(m.pages)
    m.pages.append(m._data_page(0, [b"short"]))
    cols = [NUM, {"name": "メモ", "type": 12, "num": 1, "var": 0, "size": 0}]
    m.add_table("表", cols, [m.row(cols, {"番号": _n(1), "メモ": struct.pack("<III", 5000 | 0x40000000, lval << 8, 0)})])
    _wb, sh = _sheet(m.build())
    assert _rows(sh)[1:] == [] and list(sh.bad_rows) == [2] and any("メモの長さ" in u for u in sh.unread)


# ---------------------------------------------------------------- 字
def test_jet3_uses_the_files_code_page():
    m = jet_maker.Maker(True, code_page=1251)
    cols = [{"name": "A", "type": 10, "num": 0, "var": 0, "size": 20}]
    m.add_table("T", cols, [m.row(cols, {"A": "Привет".encode("cp1251")})])
    assert _sheet(m.build())[1].cells[(2, 1)].value == "Привет"


def test_jet3_unknown_code_page_says_so():
    m = jet_maker.Maker(True, code_page=59999)
    cols = [{"name": "A", "type": 10, "num": 0, "var": 0, "size": 20}]
    m.add_table("T", cols, [m.row(cols, {"A": b"abc"})])
    _wb, sh = _sheet(m.build())
    assert any("59999" in u for u in sh.unread)


def test_jet3_cp1252_gaps_keep_the_byte():
    m = jet_maker.Maker(True, code_page=1252)
    cols = [{"name": "A", "type": 10, "num": 0, "var": 0, "size": 20}]
    m.add_table("T", cols, [m.row(cols, {"A": b"a\x81b\x93c\x9d"})])
    assert _sheet(m.build())[1].cells[(2, 1)].value == "a\x81b“c\x9d"


# ---------------------------------------------------------------- 暗号
def _page_key(blob: bytes, pg: int, key: int) -> bytes:
    head = bytearray(blob[:pg])
    size = 126 if pg == 2048 else 128
    body = bytearray(_rc4(struct.pack("<I", 0x6B39DAC7), bytes(head[0x18:0x18 + size])))
    struct.pack_into("<I", body, 0x3E - 0x18, key)
    head[0x18:0x18 + size] = _rc4(struct.pack("<I", 0x6B39DAC7), bytes(body))
    out = bytearray(head)
    for n in range(1, len(blob) // pg):
        out += _rc4(struct.pack("<I", key ^ n), blob[n * pg:(n + 1) * pg])
    return bytes(out)


@pytest.mark.parametrize("jet3", [False, True])
def test_page_key_is_the_stored_value(jet3):
    blob = _page_key(jet_maker.sample(jet3), 2048 if jet3 else 4096, 0xA7B773E0)
    _wb, sh = _sheet(blob)
    assert sh.name == "コード表" and sh.cells[(2, 4)].value == "山田"


def test_encrypted_accdb_never_asks_for_a_password_it_cannot_use(tmp_path):
    """暗号化した .accdb では、パスワードの欄の字は本当のパスワードではない。聞いても解けないので、解けないと言う。"""
    blob = bytearray(jet_maker.sample(False))
    blob[4:19] = b"Standard ACE DB"
    blob[0x14] = 3
    key = struct.pack("<I", 0x6B39DAC7)
    body = bytearray(_rc4(key, bytes(blob[0x18:0x18 + 128])))
    struct.pack_into("<d", body, 0x72 - 0x18, 45000.25)
    mask = int(45000.25).to_bytes(4, "little")
    raw = bytearray("字に見える".encode("utf-16-le").ljust(40, b"\0"))
    for i in range(40):
        raw[i] ^= mask[i % 4]
    body[0x42 - 0x18:0x42 - 0x18 + 40] = raw
    blob[0x18:0x18 + 128] = _rc4(key, bytes(body))
    blob[2 * 4096] = 0x55
    p = _p(tmp_path, "a.accdb", bytes(blob))
    for pw in (None, "本人のパスワード"):
        with pytest.raises(api.ApiError, match="暗号化"):
            api.window(p, password=pw)


# ---------------------------------------------------------------- 上限と窓
def test_small_tables_survive_the_cell_budget(monkeypatch):
    monkeypatch.setattr(jet, "MAX_CELLS", 12)
    m = jet_maker.Maker(False)
    cols = [{"name": "A", "type": 4, "num": 0, "fixed": 0, "size": 4}]
    m.add_table("A大きい", cols, [m.row(cols, {"A": _n(i)}) for i in range(40)])
    m.add_table("Z小さい", cols, [m.row(cols, {"A": _n(7)})])
    wb = open_book(m.build(), "a.mdb")
    assert [s.name for s in wb.sheets] == ["A大きい", "Z小さい"]  # シートの順は名前の順のまま
    big, small = wb.sheets
    assert _rows(small)[1:] == [[7]] and not small.unread
    assert any("表「A大きい」" in u and "上限" in u for u in big.unread)


def _two_tables() -> bytes:
    m = jet_maker.Maker(False)
    cols = [{"name": "番号", "type": 4, "num": 0, "fixed": 0, "size": 4}, TXT]
    m.add_table("甲", cols, [m.row(cols, {"番号": _n(i), "字": m.text(f"甲の{i}")}) for i in range(1, 6)])
    m.add_table("乙", cols, [m.row(cols, {"番号": _n(i), "字": m.text(f"乙の{i}")}) for i in range(1, 4)])
    return m.build()


def _window(args: dict) -> tuple[str, dict]:
    r = mcp_server.call_tool("read_window", args)
    return r["content"][0]["text"], r["structuredContent"]


def test_read_window_sheet_returns_that_table_only(tmp_path):
    p = _p(tmp_path, "a.mdb", _two_tables())
    text, _sc = _window({"path": p, "sheet": "乙"})
    assert "## シート「乙」" in text and "乙の3" in text and "甲の1" not in text and "## シート「甲」" not in text


def test_read_window_rows_within_a_table(tmp_path):
    p = _p(tmp_path, "a.mdb", _two_tables())
    text, _sc = _window({"path": p, "sheet": "甲", "rows": "3-4"})
    assert "甲の2" in text and "甲の3" in text and "甲の1" not in text and "甲の4" not in text


def test_read_window_unknown_table_is_refused_not_replaced(tmp_path):
    p = _p(tmp_path, "a.mdb", _two_tables())
    with pytest.raises(api.ApiError, match="シート「丙」はありません.*甲"):
        _window({"path": p, "sheet": "丙"})


def test_read_window_rows_alone_is_refused(tmp_path):
    p = _p(tmp_path, "a.mdb", _two_tables())
    with pytest.raises(api.ApiError, match="sheet"):
        _window({"path": p, "rows": "1-2"})


def test_read_window_sheet_can_be_followed_to_the_end(tmp_path):
    p = _p(tmp_path, "a.mdb", _two_tables())
    seen = ""
    args = {"path": p, "sheet": "甲", "max_chars": 200}
    for _ in range(50):
        text, sc = _window(args)
        seen += text
        nxt = sc.get("next")
        if not nxt:
            break
        args = {"path": p, "max_chars": 200, **nxt}
    assert all(f"甲の{i}" in seen for i in range(1, 6)) and "乙の1" not in seen


def test_read_window_sheet_carries_that_tables_note(tmp_path):
    m = jet_maker.Maker(False)
    cols = [NUM, {"name": "絵", "type": 11, "num": 1, "var": 0, "size": 0}]
    m.add_table("写真", cols, [m.row(cols, {"番号": _n(1), "絵": struct.pack("<III", 10 | 0x40000000, 0, 0)})])
    m.add_table("他", [NUM], [m.row([NUM], {"番号": _n(1)})])
    p = _p(tmp_path, "a.mdb", m.build())
    assert "表「写真」" in _window({"path": p, "sheet": "写真"})[0]
    assert "表「写真」" not in _window({"path": p, "sheet": "他"})[0]


def test_read_table_on_one_table(tmp_path):
    p = _p(tmp_path, "a.mdb", _two_tables())
    res = mcp_server.call_tool("read_table", {"path": p, "header": "番号", "sheet": "乙"})["structuredContent"]
    assert res["rows"] == 3
    assert "乙の3" in str(res["all_rows"])
