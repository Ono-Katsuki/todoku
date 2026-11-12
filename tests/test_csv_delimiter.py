"""CSV・TSV の区切りを、エージェントが窓を読んで選ぶ（delimiter）。

規則で決まらない区切り（どの行も同じ数に分かれる区切りが 2 つある・「|」などの、規則が数えない字で区切ったファイル）を、
エージェントが選んで読み直せること。選んだ区切りは、その呼び出しの間だけ効き、段階は下げず、選んだのがエージェントである
ことを理由に残すこと。字の 2 バイト目が区切りと同じバイトになる文字コード（Shift_JIS の「ポ」は 0x83 0x7C）でも、
バイトの並びから読む道と、字の並びから読む道が同じ表になること。
"""

from __future__ import annotations

import random
import zipfile

import pytest

from todoku import api, jobs, lazy, limits, mcp_server
from todoku.datafile.csvfile import OTHER_SEPS, CsvFile, DelimiterError, choosing, chosen_delimiter, delimiters
from todoku.lazy.textstream import second_byte_leads

AMB = "1,5;2,5;3,5\r\n4,5;5,5;6,5\r\n".encode()
UNDECIDED = "区切りを規則では決められない"
CHOSEN = "エージェントが選んだ区切り"


@pytest.fixture(autouse=True)
def clean(monkeypatch, tmp_path):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()


def _file(tmp_path, name, data: bytes) -> str:
    p = tmp_path / name
    p.write_bytes(data)
    return str(p)


def _table(md: str) -> list[list[str]]:
    """窓の表の行（見出しの行と区切りの行を除く）。"""
    rows = [ln for ln in md.split("\n") if ln.startswith("| R")]
    return [[c.strip() for c in ln.strip("|").split(" | ")][1:] for ln in rows]


def _window(path, **k) -> tuple[dict, str]:
    _kind, doc, parts = api.window(path, **k)
    return doc, "\n".join(md for _t, md in parts)


def _reasons(g, side="read"):
    rs = g["reasons"] if side == "read" else g["write"]["reasons"]
    return [(r["grade"], r["title"], r["by"]) for r in rs]


def _tool(name, **args):
    return mcp_server.tool_result(name, args)


# ---------------------------------------------------------------------- 選んで読み直す
def test_chosen_delimiter_rereads_the_window(tmp_path):
    p = _file(tmp_path, "amb.csv", AMB)
    doc, md = _window(p)
    assert doc["format"]["delimiter"] == "," and "区切りを規則では決められません" in md
    assert "delimiter で渡して読み直し" in md  # どうすればよいかを、窓が言う
    assert _table(md)[0] == ["1", "5;2", "5;3", "5"]

    doc, md = _window(p, delimiter=";")
    fm = doc["format"]
    assert fm["delimiter"] == ";" and fm["delimiter_chosen_by"] == "エージェント" and fm["delimiter_by_rule"] == ","
    assert fm["delimiter_even_columns"] == 3 and "delimiter_also" not in fm
    assert _table(md)[:2] == [["1,5", "2,5", "3,5"], ["4,5", "5,5", "6,5"]]
    assert "エージェントが選んで指定した 「;」（セミコロン） で読みました" in md and "どの行も 3 列に分かれます" in md
    assert "区切りを規則では決められません" not in md


def test_choice_lasts_for_one_call_only(tmp_path):
    p = _file(tmp_path, "amb.csv", AMB)
    _window(p, delimiter=";")
    assert chosen_delimiter() is None
    assert _window(p)[0]["format"]["delimiter"] == ","  # 控えも、選んだ区切りごとに分かれる
    assert CsvFile(AMB, "amb.csv").delim == ","
    with pytest.raises(api.ApiError):
        api.window(p, delimiter="x")
    assert chosen_delimiter() is None


def test_fields_fill_and_proof_use_the_chosen_delimiter(tmp_path):
    p = _file(tmp_path, "amb.csv", AMB)
    assert [f["id"] for f in api.fields(p, delimiter=";")] == ["R2.C1", "R2.C2", "R2.C3", "R3.C1", "R3.C2", "R3.C3"]
    out = tmp_path / "out.csv"
    r = api.fill(p, {"R2.C3": "7,5", "R1.C2": "9;9"}, str(out), delimiter=";")
    assert r["read_back_ok"] is True
    # 値の中のカンマはそのまま、選んだ区切りの字は引用符で囲む。ほかのバイトは元のまま
    assert out.read_bytes() == b'1,5;"9;9";3,5\r\n4,5;5,5;7,5\r\n'
    assert _table(_window(str(out), delimiter=";")[1])[:2] == [["1,5", "9;9", "3,5"], ["4,5", "5,5", "7,5"]]


def test_grade_is_not_lowered_and_keeps_who_chose(tmp_path):
    p = _file(tmp_path, "amb.csv", AMB)
    g = api.grade(p)
    assert (g["grade"], g["write"]["grade"]) == ("B", "B") and ("B", UNDECIDED, "rules") in _reasons(g)
    g = api.grade(p, delimiter=";")
    assert (g["grade"], g["write"]["grade"]) == ("S", "A")
    assert ("S", CHOSEN, "agent_choice") in _reasons(g) and ("A", CHOSEN, "agent_choice") in _reasons(g, "write")
    assert not [r for r in _reasons(g) if r[1] == UNDECIDED]
    # 窓の読み方（カンマ）が正しいときは、その字を渡すと、エージェントが確かめた読み方として残る
    g = api.grade(p, delimiter=",")
    assert g["grade"] == "S" and ("S", CHOSEN, "agent_choice") in _reasons(g)
    assert api.window(p, delimiter=",")[1]["format"]["delimiter_even_columns"] == 4


def test_uneven_choice_is_kept_as_unchecked(tmp_path):
    p = _file(tmp_path, "rag.csv", "a,b\r\nc;d,e\r\nf\r\n".encode())
    doc, md = _window(p, delimiter=";")
    assert "delimiter_even_columns" not in doc["format"] and "行によって列の数が違うか、区切りのない行があります" in md
    g = api.grade(p, delimiter=";")
    assert g["grade"] == "S" and ("S", CHOSEN, "agent_unchecked") in _reasons(g)
    text = _tool("grade_document", path=p, delimiter=";")["content"][0]["text"]
    assert "（エージェントの選択。規則では確かめられない）" in text and "（エージェントの選択を規則で確かめた）" not in text
    text = _tool("grade_document", path=_file(tmp_path, "amb.csv", AMB), delimiter=";")["content"][0]["text"]
    assert "（エージェントの選択を規則で確かめた）" in text


def test_audit_reads_with_the_chosen_delimiter(tmp_path):
    p = _file(tmp_path, "amb.csv", AMB)
    a = api.audit(p, delimiter=";")
    assert a["ok"] and "マス 6 か所" in a["checks"][0]["detail"]
    assert "マス 8 か所" in api.audit(p)["checks"][0]["detail"]


# ---------------------------------------------------------------------- 規則が数えない字で区切ったファイル
def test_other_symbol_is_told_when_the_three_are_absent(tmp_path):
    p = _file(tmp_path, "pipe.csv", "名前|数|日\n山田|3|月\n田中|4|火\n".encode())
    assert delimiters("a|b\nc|d\n", "csv") == (",", [("|", 2)], False)
    doc, md = _window(p)
    fm = doc["format"]
    assert fm["delimiter"] == "," and fm["delimiter_absent"] is True
    assert fm["delimiter_also"] == [{"delimiter": "|", "delimiter_ja": "「|」（縦線）", "columns": 3}]
    assert "「,」・タブ・「;」がどの行にもないので、窓は 1 列です" in md
    g = api.grade(p)
    assert g["grade"] == "B" and [r["where"] for r in g["reasons"] if r["title"] == UNDECIDED] == [["|"]]

    doc, md = _window(p, delimiter="|")
    assert _table(md)[:3] == [["名前", "数", "日"], ["山田", "3", "月"], ["田中", "4", "火"]]
    g = api.grade(p, delimiter="|")
    assert g["grade"] == "S" and ("S", CHOSEN, "agent_choice") in _reasons(g)


def test_symbols_inside_values_are_not_decided(tmp_path):
    # 1 列の日付：「-」もそろう。値の一部か区切りかは、規則では決めない（窓は 1 列のまま）
    p = _file(tmp_path, "dates.csv", b"2024-01-05\n2024-02-11\n2025-12-31\n")
    doc, md = _window(p)
    assert doc["format"]["delimiter_also"] == [{"delimiter": "-", "delimiter_ja": "「-」", "columns": 3}]
    assert _table(md)[0] == ["2024-01-05"]
    g = api.grade(p, delimiter=",")  # 1 列と確かめた
    assert g["grade"] == "S" and ("S", CHOSEN, "agent_unchecked") in _reasons(g)


def test_no_note_when_nothing_else_splits_evenly(tmp_path):
    for data in (b"apple\nbanana\ncherry\n",            # 記号のない 1 列
                 b"a|b,c\nd|e\nf|g,h,i\n",               # カンマのある行がある：カンマで読む（「|」は数えない）
                 b"a,b\nc,d\n", b"a\tb\nc\td\n", b"a;b\nc;d\n"):
        doc, md = _window(_file(tmp_path, "x.csv", data))
        assert "delimiter_also" not in doc["format"] and "区切りを規則では決められません" not in md, data
        assert api.grade(_file(tmp_path, "x.csv", data))["grade"] == "S"
        api._WINDOWS.clear()


def test_sep_line_and_tsv_can_be_overridden(tmp_path):
    p = _file(tmp_path, "sep.csv", b"sep=;\r\na,b;c\r\n1,2;3\r\n")
    doc, md = _window(p)
    assert doc["format"]["delimiter"] == ";" and doc["format"]["delimiter_declared_by"]
    doc, md = _window(p, delimiter=",")
    fm = doc["format"]
    assert fm["delimiter"] == "," and fm["delimiter_by_rule"] == ";" and fm["delimiter_even_columns"] == 2  # 宣言の行は数えない
    assert "規則で読むと、区切りは 「;」（セミコロン） です（1行目の宣言）" in md and "区切りの宣言の行" in md
    t = _file(tmp_path, "t.tsv", b"a,b\tc\r\n1,2\t3\r\n")
    assert _window(t)[0]["format"]["delimiter"] == "\t"
    assert _table(_window(t, delimiter=",")[1])[0] == ["a", "b⇥c"]
    assert _window(t, delimiter="tab")[0]["format"]["delimiter"] == "\t" and _window(t, delimiter="\\t")[0]["format"]["delimiter"] == "\t"


# ---------------------------------------------------------------------- 選べない値・使えない形式
@pytest.mark.parametrize("bad", ["a", "1", "", "||", " ", '"', "\n", "\r", "、", "｜", "\x01", 5])
def test_values_that_cannot_be_a_delimiter_are_refused(tmp_path, bad):
    p = _file(tmp_path, "amb.csv", AMB)
    with pytest.raises(api.ApiError, match="ASCII の記号 1 字かタブ"):
        api.window(p, delimiter=bad)
    r = _tool("open_document", path=p, delimiter=bad)
    assert r["isError"] and "ASCII の記号 1 字かタブ" in r["content"][0]["text"]


def test_other_kinds_refuse_the_argument(tmp_path):
    j = _file(tmp_path, "a.json", b'{"a": 1}')
    x = _file(tmp_path, "a.txt", "あ;い\n".encode())
    for p in (j, x):
        for call in (lambda: api.window(p, delimiter=";"), lambda: api.fields(p, delimiter=";"), lambda: api.grade(p, delimiter=";"),
                     lambda: api.read_range(p, {}, delimiter=";")):
            with pytest.raises(api.ApiError, match="CSV・TSV にだけ指定できます"):
                call()
    z = tmp_path / "z.zip"
    with zipfile.ZipFile(z, "w") as f:
        f.writestr("amb.csv", AMB)
    r = _tool("open_document", path=str(z), delimiter=";")
    assert r["isError"] and "CSV・TSV にだけ指定できます" in r["content"][0]["text"]
    # 書庫の中の CSV は、その CSV を指して渡す
    r = _tool("grade_document", path=f"{z}!/amb.csv", delimiter=";")
    assert not r.get("isError") and CHOSEN in r["content"][0]["text"]


# ---------------------------------------------------------------------- 字の 2 バイト目が区切りと同じバイトになる文字コード
def test_shift_jis_pipe_is_not_cut_inside_a_character(tmp_path):
    text = "名前|数|メモ\r\nポチ|3|ソフト\r\nタマ|4|\"表|能\"\r\n"
    data = text.encode("cp932")
    assert b"\x83\x7c" in data  # 「ポ」の 2 バイト目は「|」と同じ
    p = _file(tmp_path, "sjis.csv", data)
    want = [["名前", "数", "メモ"], ["ポチ", "3", "ソフト"], ["タマ", "4", "表\\|能"]]
    doc, md = _window(p, delimiter="|")
    assert _table(md)[:3] == want and doc["format"]["encoding"] == "cp932"
    page, _doc = api.read_range(p, {"rows": [1, 3]}, delimiter="|")  # バイトの並びから読む道
    assert _table(page.text)[:3] == want
    with choosing("|"):
        d = lazy.open_doc(p, None, None, True)
        assert d.width == 3 and [d.record_values(i) for i in (1, 2, 3)] == [["名前", "数", "メモ"], ["ポチ", "3", "ソフト"], ["タマ", "4", "表|能"]]
    out = tmp_path / "o.csv"
    api.fill(p, {"R2.C2": "ポ|ポ"}, str(out), delimiter="|")
    assert out.read_bytes() == text.replace("ポチ|3|", 'ポチ|"ポ|ポ"|').encode("cp932")


@pytest.mark.parametrize("enc", ["cp932", "shift_jis_2004", "gbk", "gb18030", "big5", "big5hkscs", "johab", "cp949", "euc_jp", "utf-8",
                                 "utf-16-le"])
def test_byte_reading_matches_text_reading(tmp_path, enc):
    """バイトの並びから記録とマスを探す読み方（大きなファイル・範囲の読み方）が、字にしてから読む読み方と同じ表になる。"""
    rnd = random.Random(11)
    chars = []
    for cp in list(range(0x80, 0x3000, 7)) + list(range(0x3000, 0xA000, 3)) + list(range(0xAC00, 0xD7A4, 5)):
        try:
            if chr(cp).encode(enc).decode(enc) == chr(cp):
                chars.append(chr(cp))
        except UnicodeError:
            pass
    hot_seen = False
    for delim in [d for d in (*OTHER_SEPS, ";") if second_byte_leads(enc, d)] or ["|"]:
        db = delim.encode(enc)
        hot = [c for c in chars if db in c.encode(enc)] if enc != "utf-16-le" else []
        hot_seen = hot_seen or bool(hot)
        pool = hot * 3 + rnd.sample(chars, min(len(chars), 200))
        rows = []
        for _r in range(60):
            cells = []
            for _c in range(rnd.randint(1, 5)):
                v = "".join(rnd.choice(pool) if rnd.random() < 0.8 else rnd.choice(["a", "1", " ", ""]) for _ in range(rnd.randint(0, 6)))
                cells.append('"' + v + rnd.choice(["", delim, "\n", '""']) + '"' if rnd.random() < 0.15 else v)
            rows.append(delim.join(cells))
        blob = (rnd.choice(["\n", "\r\n"]).join(rows) + "\n").encode(enc)
        p = _file(tmp_path, f"r{ord(delim)}.csv", blob)
        with choosing(delim):
            want = [r.values for r in CsvFile(blob, "r.csv", "csv", enc).records]
            d = lazy.open_doc(p, None, enc, True)
            assert [d.record_values(i) for i in range(1, d.idx.count + 1)] == want, (enc, delim)
            assert d.width == max(len(r) for r in want), (enc, delim)
    # 2 バイト目になる文字コードでは、区切りと同じバイトを持つ字を、実際に入れて確かめている
    assert hot_seen or not (second_byte_leads(enc, "|") or second_byte_leads(enc, ";"))


def test_which_encodings_need_the_second_byte_check():
    assert second_byte_leads("cp932", "|") and second_byte_leads("gbk", "|") and second_byte_leads("big5", "|")
    assert second_byte_leads("johab", ";")  # Johab は、2 バイト目が 0x31 から：規則の「;」も、字の中に現れる
    for enc in ("cp932", "gbk", "gb18030", "big5", "johab", "cp949", "euc_jp", "euc_kr", "utf-8", "utf-16-le", "cp1252"):
        assert not second_byte_leads(enc, ",") and not second_byte_leads(enc, "\t"), enc
    for enc in ("cp949", "euc_jp", "euc_kr", "utf-8", "cp1252", "utf-16-le"):
        assert not second_byte_leads(enc, "|"), enc


def test_stateful_encoding_reads_other_symbols_through_a_utf8_copy(tmp_path):
    """切り替えの印のある文字コード：バイトの並びからは区切りを探せない（second_byte_leads は断る）ので、
    範囲の道は、UTF-8 に直した写しを読む。「ポ」は JIS で 0x25 0x5D、「ぬ」は 0x24 0x4C、「|」は 0x7C。"""
    p = _file(tmp_path, "jis.csv", "名前|数\n穐山|3\nポ|ぬ\n".encode("iso2022_jp"))
    assert _table(_window(p, encoding="iso2022_jp", delimiter="|")[1])[:2] == [["名前", "数"], ["穐山", "3"]]
    for enc in ("iso2022_jp", None):  # 規則も、切り替えの印から ISO-2022-JP と決める
        page, _doc = api.read_range(p, {"rows": [1, 3]}, encoding=enc, delimiter="|")
        assert _table(page.text)[:3] == [["名前", "数"], ["穐山", "3"], ["ポ", "ぬ"]]
    with pytest.raises(DelimiterError, match="範囲で分けて読むときに選べません"):
        second_byte_leads("iso2022_jp", "|")


# ---------------------------------------------------------------------- 大きなファイル（分けて読む道）と長い仕事
def test_large_csv_is_read_and_written_with_the_chosen_delimiter(tmp_path, monkeypatch):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)
    monkeypatch.setattr(limits, "PAGE_CHARS", 1_000)
    rows = [f"{i},5;x{i};{i * 2},25" for i in range(1, 801)]
    data = ("\r\n".join(rows) + "\r\n").encode()
    p = _file(tmp_path, "big.csv", data)
    r = _tool("open_document", path=p)
    assert r["structuredContent"]["large"] and r["structuredContent"]["format"]["delimiter"] == ","
    r = _tool("open_document", path=p, delimiter=";")
    sc = r["structuredContent"]
    assert sc["large"] and sc["format"]["delimiter"] == ";" and sc["format"]["delimiter_chosen_by"] == "エージェント"
    page, _doc = api.read_range(p, {"rows": [400, 401]}, delimiter=";")
    assert _table(page.text) == [["1,5", "x1", "2,25"], ["400,5", "x400", "800,25"], ["401,5", "x401", "802,25"]]  # 見出しの行と、頼んだ行
    page, _doc = api.read_range(p, {"rows": [400, 400]})  # 選ばなければ、規則の読み方（控えは別）
    assert _table(page.text)[-1] == ["400", "5;x400;800", "25"]
    out = tmp_path / "big_out.csv"
    res = api.fill(p, {"R400.C2": "書いた;値"}, str(out), delimiter=";")
    assert res["read_back_ok"] is True
    assert out.read_bytes() == data.replace(b"400,5;x400;", '400,5;"書いた;値";'.encode())
    g = api.grade(p, delimiter=";")
    assert CHOSEN in [r["title"] for r in g["reasons"]] and UNDECIDED not in [r["title"] for r in g["reasons"]]


def test_jobs_are_keyed_by_the_chosen_delimiter(tmp_path):
    p = _file(tmp_path, "amb.csv", AMB)
    plain = jobs.job_key(p, None)
    with choosing(";"):
        semi = jobs.job_key(p, None)
    with choosing("|"):
        pipe = jobs.job_key(p, None)
    assert len({plain, semi, pipe}) == 3 and jobs.job_key(p, None) == plain


def test_audit_job_runs_with_the_chosen_delimiter(tmp_path):
    import time

    p = _file(tmp_path, "amb.csv", AMB)
    for delim, cells in ((";", "マス 6 か所"), (None, "マス 8 か所")):
        args = {"path": p, "job": True, **({"delimiter": delim} if delim else {})}
        for _ in range(240):
            r = _tool("audit_document", **args)
            if not ((r.get("structuredContent") or {}).get("job") or {}).get("running"):
                break
            time.sleep(0.25)
        assert not r.get("isError") and cells in r["content"][0]["text"], r["content"][0]["text"][:300]


# ---------------------------------------------------------------------- 道具
def test_mcp_tools_take_delimiter(tmp_path):
    tools = {t["name"]: t for t in mcp_server.TOOLS}
    takes = {n for n, t in tools.items() if "delimiter" in t["inputSchema"]["properties"]}
    assert takes == {"open_document", "read_window", "list_fields", "fill_document", "audit_document", "grade_document"}
    p = _file(tmp_path, "amb.csv", AMB)
    assert _tool("open_document", path=p, delimiter=";")["structuredContent"]["format"]["delimiter"] == ";"
    r = _tool("read_window", path=p, delimiter=";")
    assert _table("\n".join(c["text"] for c in r["content"]))[:1] == [["1,5", "2,5", "3,5"]]
    assert '"id":"R2.C3"' in _tool("list_fields", path=p, delimiter=";")["content"][0]["text"]
    out = tmp_path / "m.csv"
    r = _tool("fill_document", path=p, values={"R2.C3": "7,5"}, out_path=str(out), delimiter=";")
    assert not r.get("isError") and out.read_bytes() == b"1,5;2,5;3,5\r\n4,5;5,5;7,5\r\n"
    assert len(mcp_server.INSTRUCTIONS.encode("utf-16-le")) // 2 < 4096 and "delimiter" in mcp_server.INSTRUCTIONS
