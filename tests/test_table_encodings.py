"""8 ビットの字の表（dBASE .dbf・Lotus 1-2-3 .wk1・Excel 2〜4 の .xls）の文字コード：日本語と置かずに、ファイル全体で 1 つに決める。

- ファイルの宣言（.dbf：隣の .cpg・頭の言語の印。Excel 2〜4：CODEPAGE の記録）で全部の字を厳密に読めれば、それ
- 宣言がなく、全部の字が UTF-8 の形なら UTF-8
- ほかは、バイトからは決まらない：決まった順のいちばん前で示し、ほかの読み方を候補として並べ、決められないと伝える（読む B）
- エージェントが選んだ文字コード（encoding）は、全部の字を厳密に読めるときだけ使い、段階は下げない
- .dbf に ASCII でない字を書くのは、文字コードが決まっているときだけ
"""

from __future__ import annotations

import struct
import zipfile

import pytest

from test_more_formats import _dbf, _p, _r
from todoku import api, mcp_server
from todoku.book import open_book
from todoku.legacy.dbf import LDID, cpg_encoding

KO = ("홍길동", "김철수")
TOLD = "値の文字コードを規則では決められません"
TITLE = "字の文字コードを規則では決められない"
CHOSEN = "エージェントが選んだ文字コード"


def _row(name: bytes, note: bytes = b" " * 10):
    return (b" ", name.ljust(12)[:12], b"   12.50", b"20260401", b"T", note)


def _kdbf(ldid: int = 0) -> bytes:
    return _dbf(ldid=ldid, rows=[_row(s.encode("cp949")) for s in KO])


def _wk1(*labels: bytes) -> bytes:
    recs = [_r(0, struct.pack("<H", 0x0406))]
    for i, s in enumerate(labels):
        recs.append(_r(15, b"\xff" + struct.pack("<HH", 0, i) + b"'" + s + b"\0"))
    return b"".join(recs + [_r(1, b"")])


def _biff2(cp: int, *labels: bytes) -> bytes:
    attr = bytes([0x40, 0x00, 0x40])
    return b"".join([_r(0x0009, struct.pack("<HH", 2, 0x10))] + ([_r(0x0042, struct.pack("<H", cp))] if cp else [])
                    + [_r(0x001E, b"\x07General")]
                    + [_r(0x0004, struct.pack("<HH", i, 0) + attr + bytes([len(s)]) + s) for i, s in enumerate(labels)]
                    + [_r(0x000A, b"")])


def _md(path, **kw) -> str:
    _k, _doc, parts = api.window(path, **kw)
    return "\n".join(t for _n, t in parts)


def _te(path, **kw) -> dict:
    return api.window(path, **kw)[1].get("text_encoding") or {}


def _rows(path, header, **kw) -> list[dict]:
    return api.read_table(path, header, **kw)["all_rows"]


def _reasons(g: dict, side: str | None = None) -> list[tuple]:
    rs = g[side]["reasons"] if side else g["reasons"]
    return [(r["grade"], r["title"], r.get("by")) for r in rs]


# ====================================================================== dBASE：宣言がない
def test_dbf_without_declaration_is_told_with_candidates(tmp_path):
    p = _p(tmp_path, "k.dbf", _kdbf())
    te = _te(p)
    assert te["basis"] == "order" and te["count"] == 2
    cands = {c["encoding"]: c["sample"] for c in te["candidates"]}
    assert cands["cp949"] == KO[0] and "cp1251" in cands  # 韓国語としても、キリル文字としても、きれいに読める
    md = _md(p)
    assert TOLD in md and f"cp949：{KO[0]}" in md and "encoding で渡します" in md
    s = api.summary(p)["text_encoding"]
    assert s["basis"] == "order" and "notes" not in s and {"encoding": "cp949", "sample": KO[0]} in s["candidates"]
    g = api.grade(p)
    assert ("B", TITLE, "rules") in _reasons(g) and ("B", TITLE, "rules") in _reasons(g, "write")
    assert g["write"]["grade"] == "B"


def test_dbf_japanese_without_declaration_reads_as_before_but_is_told(tmp_path):
    p = _p(tmp_path, "j.dbf", _dbf(ldid=0))
    te = _te(p)
    assert te["encoding"] == "cp932" and te["basis"] == "order"  # Shift_JIS は順のいちばん前：字は前と同じ
    assert "山田" in _md(p) and TOLD in _md(p)
    assert not _te(_p(tmp_path, "d.dbf", _dbf()))  # 言語の印（0x13）がある：伝えることはない


def test_dbf_choice_reads_every_tool_without_downgrade(tmp_path):
    p = _p(tmp_path, "k.dbf", _kdbf())
    for name in ("euc-kr", "cp949", "UHC"):
        te = _te(p, encoding=name)
        assert te["basis"] == "given" and te["encoding"] == "cp949" and te["chosen_by"] == "agent" and "candidates" not in te
    md = _md(p, encoding="cp949")
    assert KO[0] in md and KO[1] in md and TOLD not in md and "エージェントが選んで指定した" in md
    g = api.grade(p, encoding="cp949")
    assert ("S", CHOSEN, "agent_choice") in _reasons(g) and not [r for r in _reasons(g) + _reasons(g, "write") if r[1] == TITLE]
    assert g["write"]["grade"] == "B" and _reasons(g, "write")[0][1] == "項目の幅の中で上書きする"  # 書く段階は、宣言のあるファイルと同じ
    assert [r["NAME"] for r in _rows(p, "NAME", encoding="cp949")] == list(KO)
    assert api.fields(p, encoding="cp949") == api.fields(p)  # 表に欄の一覧はない（番地で書く）：選んでも、呼び出しは通る
    assert api.audit_whole(p, encoding="cp949")["ok"] and api.audit_whole(p)["ok"]
    assert api.summary(p, encoding="cp949")["text_encoding"]["basis"] == "given"
    # 選択は、その呼び出しだけ：次の呼び出しは、また規則で読む
    assert _te(p)["basis"] == "order" and _rows(p, "NAME")[0]["NAME"] != KO[0]


@pytest.mark.parametrize("name,said", [("nope", "知りません"), ("utf-8", "厳密に読めません"), ("ascii", "厳密に読めません"),
                                       ("cp932", "厳密に読めません")])
def test_dbf_refuses_unknown_and_unreadable_choice(tmp_path, name, said):
    p = _p(tmp_path, "k.dbf", _kdbf())
    for call in (lambda: api.window(p, encoding=name), lambda: api.grade(p, encoding=name),
                 lambda: api.read_table(p, "NAME", encoding=name),
                 lambda: api.fill(p, {"A2": "x"}, str(tmp_path / "o.dbf"), encoding=name)):
        with pytest.raises(api.ApiError, match=said):
            call()
    assert not (tmp_path / "o.dbf").exists()


# ====================================================================== dBASE：宣言
def test_cpg_parser_and_language_ids():
    got = {s: cpg_encoding(s) for s in (b"UTF-8\n", b"65001", b"LDID/19", b"LDID/87", b"ANSI 1251", b"OEM 866", b"8859-2", b"88592",
                                         b"1252", b"Shift_JIS", b"EUC-KR", b"BIG5", b"", b"\xff\xfe", b"nope", b"0")}
    assert got == {b"UTF-8\n": "utf-8", b"65001": "utf-8", b"LDID/19": "cp932", b"LDID/87": None, b"ANSI 1251": "cp1251",
                   b"OEM 866": "cp866", b"8859-2": "iso8859-2", b"88592": "iso8859-2", b"1252": "cp1252", b"Shift_JIS": "cp932",
                   b"EUC-KR": "cp949", b"BIG5": "cp950", b"": None, b"\xff\xfe": None, b"nope": None, b"0": None}
    assert LDID[0x13] == "cp932" and LDID[0x65] == "cp866" and LDID[0x4E] == "cp949" and 0x57 not in LDID  # 0x57「ANSI」は宣言にしない
    for enc in LDID.values():
        "a".encode(enc)  # どの印も、この機械の Python が知っている文字コード


def test_dbf_language_id_declares(tmp_path):
    p = _p(tmp_path, "r.dbf", _dbf(ldid=0x65, rows=[_row("Москва".encode("cp866"))]))
    wb = open_book(open(p, "rb").read(), "r.dbf")
    assert wb.encoding == "cp866" and wb.text_reading.basis == "declared" and wb.sheets[0].cells[(2, 1)].value == "Москва"
    assert not _te(p) and TOLD not in _md(p)
    g = api.grade(p)
    assert not [r for r in _reasons(g) + _reasons(g, "write") if r[1] in (TITLE, CHOSEN)]
    # 宣言と違う読み方を選べる（厳密に読めるとき）。ファイルの宣言は、そのまま伝える
    te = _te(p, encoding="cp1251")
    assert te["basis"] == "given" and te["encoding"] == "cp1251" and "CP866" in te["declared"]
    assert "CP866" in "\n".join(te["notes"])


def test_dbf_declaration_that_cannot_read_is_told(tmp_path):
    raw = "Москва".encode("cp1251")
    with pytest.raises(UnicodeDecodeError):
        raw.ljust(12).decode("cp932")
    p = _p(tmp_path, "x.dbf", _dbf(ldid=0x13, rows=[_row(raw)]))  # Shift_JIS の印があるが、Shift_JIS として読めない
    te = _te(p)
    assert te["basis"] == "order" and te["encoding"] != "cp932" and "Shift_JIS" in te["declared"]
    assert te["declared"] in "\n".join(te["notes"]) and "cp1251" in {c["encoding"] for c in te["candidates"]}
    assert "Москва" in _md(p, encoding="cp1251")
    assert ("B", TITLE, "rules") in _reasons(api.grade(p))


def test_cpg_sibling_declares_and_cut_value_is_told(tmp_path):
    names = ["Αθήνα Πάτρα".encode("utf-8"), "Ωmega".encode("utf-8")]
    assert len(names[0][:12]) == 12 and names[0][:12].decode("utf-8", "ignore") == "Αθήνα "  # 12 バイト目は、2 バイトの字の前半
    blob = _dbf(ldid=0x13, rows=[_row(n) for n in names])
    p = _p(tmp_path, "cut.dbf", blob)
    (tmp_path / "cut.cpg").write_bytes(b"UTF-8\n")
    wb = open_book(blob, "cut.dbf")
    assert wb.encoding == "cp932" or wb.text_reading.basis != "declared"  # 隣のファイルを渡さなければ、.cpg は見えない
    md = _md(p)
    assert "Αθήνα" in md and "Ωmega" in md and "Πάτρα" not in md
    assert "終わりの字が途中で切れている値が 1 個" in md and "A2" in md.split("終わりの字が途中で切れている値")[1].split("\n")[0]
    assert not _te(p)  # .cpg は、頭の言語の印より先に見る宣言：伝えることはない


def test_cpg_inside_zip_and_utf8_shape(tmp_path):
    blob = _dbf(ldid=0, rows=[_row("Αθήνα".encode("utf-8"))])
    z = tmp_path / "s.zip"
    with zipfile.ZipFile(z, "w") as f:
        f.writestr("x/u.dbf", blob)
        f.writestr("x/U.CPG", b"65001")
        f.writestr("x/v.dbf", blob)
    assert "Αθήνα" in _md(f"{z}!/x/u.dbf") and not _te(f"{z}!/x/u.dbf")
    # 宣言がなくても、全部の字が UTF-8 の形なら UTF-8（決められないとは言わない）
    assert "Αθήνα" in _md(f"{z}!/x/v.dbf") and _te(f"{z}!/x/v.dbf").get("basis") not in ("order", "none")


# ====================================================================== dBASE：書く
def test_dbf_write_waits_for_the_choice(tmp_path):
    p = _p(tmp_path, "k.dbf", _kdbf())
    out = tmp_path / "o.dbf"
    with pytest.raises(api.ApiError, match="encoding で渡してください"):
        api.fill(p, {"A2": "박영희"}, str(out))
    assert not out.exists()
    r = api.fill(p, {"A2": "abc"}, str(out))  # ASCII の値は、選ぶ前でも書ける（どの文字コードでも同じバイト）
    assert r["read_back_ok"] and out.read_bytes().count(b"abc".ljust(12)) == 1
    out2 = tmp_path / "o2.dbf"
    r = api.fill(p, {"A2": "박영희"}, str(out2), encoding="euc-kr")
    assert r["read_back_ok"] and r["encoding"]["encoding"] == "cp949"
    new = out2.read_bytes()
    assert "박영희".encode("cp949").ljust(12) in new and KO[1].encode("cp949") in new and KO[0].encode("cp949") not in new
    assert [x["NAME"] for x in _rows(str(out2), "NAME", encoding="cp949")] == ["박영희", KO[1]]
    with pytest.raises(api.ApiError):  # 選んだ文字コードにない字は書かない
        api.fill(p, {"A2": "山田"}, str(tmp_path / "o3.dbf"), encoding="iso8859-7")
    assert not (tmp_path / "o3.dbf").exists()


def test_ascii_only_dbf_takes_the_writing_encoding(tmp_path):
    p = _p(tmp_path, "a.dbf", _dbf(ldid=0, rows=[_row(b"abc")]))
    assert not _te(p)
    out = tmp_path / "o.dbf"
    with pytest.raises(api.ApiError, match="encoding で渡してください"):
        api.fill(p, {"A2": "山田"}, str(out))
    assert not out.exists()
    r = api.fill(p, {"A2": "山田"}, str(out), encoding="cp932")
    assert r["read_back_ok"] and "山田".encode("cp932").ljust(12) in out.read_bytes()
    te = _te(p, encoding="cp932")
    assert te["basis"] == "given" and te["count"] == 0 and "ASCII でない字はありません" in te["notes"][0]
    g = api.grade(p, encoding="cp932")
    assert ("A", CHOSEN, "agent_choice") in _reasons(g, "write") and not [r for r in _reasons(g) if r[1] == CHOSEN]
    # 宣言のあるファイルは、選ばなくても宣言の文字コードで書く
    q = _p(tmp_path, "d.dbf", _dbf(ldid=0x13, rows=[_row(b"abc")]))
    r = api.fill(q, {"A2": "山田"}, str(tmp_path / "o2.dbf"))
    assert r["read_back_ok"] and "山田".encode("cp932").ljust(12) in (tmp_path / "o2.dbf").read_bytes()


# ====================================================================== Lotus 1-2-3
def test_wk1_is_read_as_one_file_not_per_string(tmp_path):
    a, b = "이름".encode("cp949"), "주소".encode("cp949")
    assert a.decode("cp932", "ignore") != "" and b"abc".isascii()
    p = _p(tmp_path, "k.wk1", _wk1(a, b, b"abc"))
    te = _te(p)
    assert te["basis"] == "order" and te["count"] == 2 and "cp949" in {c["encoding"] for c in te["candidates"]}
    wb = open_book(open(p, "rb").read(), "k.wk1")
    enc = wb.text_reading.encoding
    assert [c.value for c in wb.sheets[0].cells.values()] == [a.decode(enc), b.decode(enc), "abc"]  # どの字も、同じ 1 つの文字コード
    assert TOLD in _md(p) and ("B", TITLE, "rules") in _reasons(api.grade(p))
    md = _md(p, encoding="cp949")
    assert "이름" in md and "주소" in md and TOLD not in md
    g = api.grade(p, encoding="cp949")
    assert ("S", CHOSEN, "agent_choice") in _reasons(g) and not [r for r in _reasons(g) if r[1] == TITLE]
    assert _rows(p, "이름", encoding="cp949")[0]["이름"] == "주소"
    assert api.audit_whole(p, encoding="cp949")["ok"]


def test_ascii_wk1_has_nothing_to_choose(tmp_path):
    p = _p(tmp_path, "a.wk1", _wk1(b"abc", b"def"))
    assert not _te(p) and TOLD not in _md(p)
    with pytest.raises(api.ApiError, match="選べる字がありません"):
        api.window(p, encoding="cp949")


# ====================================================================== Excel 2〜4
def test_excel2_codepage_record_declares(tmp_path):
    s = "Москва".encode("cp1251")
    p = _p(tmp_path, "c.xls", _biff2(1251, s))
    wb = open_book(open(p, "rb").read(), "c.xls")
    assert wb.text_reading.basis == "declared" and wb.sheets[0].cells[(1, 1)].value == "Москва"
    assert not _te(p) and not [r for r in _reasons(api.grade(p)) if r[1] in (TITLE, CHOSEN)]
    with pytest.raises(api.ApiError, match="厳密に読めません"):
        api.window(p, encoding="utf-8")


def test_excel2_without_codepage_is_told_and_takes_choice(tmp_path):
    s = "Москва".encode("cp1251")
    p = _p(tmp_path, "n.xls", _biff2(0, s, "Тула".encode("cp1251")))
    te = _te(p)
    assert te["basis"] == "order" and "cp1251" in {c["encoding"] for c in te["candidates"]}
    assert ("B", TITLE, "rules") in _reasons(api.grade(p))
    for name in ("windows-1251", "cp1251"):
        assert "Москва" in _md(p, encoding=name) and "Тула" in _md(p, encoding=name)
    assert _rows(p, "Москва", encoding="cp1251")[0]["Москва"] == "Тула"
    assert ("S", CHOSEN, "agent_choice") in _reasons(api.grade(p, encoding="cp1251"))
    d = api.calc_model(p, encoding="cp1251")
    assert isinstance(d, dict)


# ====================================================================== ほかの形式・道具の入口
def test_other_kinds_still_refuse_encoding(tmp_path):
    import openpyxl

    x = tmp_path / "a.xlsx"
    wb = openpyxl.Workbook()
    wb.active["A1"] = "x"
    wb.save(x)
    with pytest.raises(api.ApiError, match="dbf|表"):
        api.window(str(x), encoding="cp949")


def test_mcp_tools_pass_the_choice(tmp_path):
    p = _p(tmp_path, "k.dbf", _kdbf())
    by = {t["name"]: t for t in mcp_server.TOOLS}
    for name in ("read_table", "explain_calculation", "try_calculation", "extract_processing"):
        assert "encoding" in by[name]["inputSchema"]["properties"]
    assert "値の文字コードを規則では決められません" in by["open_document"]["inputSchema"]["properties"]["encoding"]["description"]
    got = mcp_server.call_tool("read_table", {"path": p, "header": "NAME", "encoding": "euc-kr"})
    assert KO[0] in str(got) and KO[1] in str(got)
    got = mcp_server.call_tool("read_window", {"path": p})
    assert TOLD in str(got)
    got = mcp_server.call_tool("read_window", {"path": p, "encoding": "euc-kr"})
    assert KO[0] in str(got) and TOLD not in str(got)
    got = mcp_server.call_tool("grade_document", {"path": p, "encoding": "euc-kr"})
    assert CHOSEN in str(got) and TITLE not in str(got)
