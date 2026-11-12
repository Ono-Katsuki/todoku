"""書庫の中の文字のファイルを、ファイルそのものと同じ範囲（rows・records・paragraphs・path_in）で読む。
頼まれた範囲の名前が、その書類の読み方にないときは断る（黙って、書類の頭から読まない）。"""
import io
import json
import zipfile

import pytest

from todoku import api, lazy, mcp_server

CSV = ("氏名,住所,金額,備考\r\n" + "".join(f"架空 {i},東京都架空区 {i} 番地,{i * 100},\r\n" for i in range(1, 2001))).encode("cp932")
TSV = ("x\ty\r\n" + "".join(f"{i}\t{i * 2}\r\n" for i in range(1, 301))).encode("utf-8")
JSONL = "".join(json.dumps({"n": i, "名前": f"架空 {i}"}, ensure_ascii=False) + "\n" for i in range(1, 501)).encode("utf-8")
TXT = "".join(f"{i} 行目の字\n" for i in range(1, 501)).encode("utf-8")
JSON = json.dumps({"a": [{"n": i} for i in range(50)], "名前": "架空"}, ensure_ascii=False).encode("utf-8")
FILES = {"data/a.csv": CSV, "t.tsv": TSV, "b.jsonl": JSONL, "c.txt": TXT, "d.json": JSON}


@pytest.fixture()
def box(tmp_path):
    with zipfile.ZipFile(tmp_path / "a.zip", "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in FILES.items():
            z.writestr(n, b)
    for n, b in FILES.items():
        (tmp_path / n.split("/")[-1]).write_bytes(b)
    return tmp_path


def _read(p, rng, n=4000, **kw):
    page, doc = api.read_range(str(p), rng, n, **kw)
    return page.text, page.next, page.covered, type(doc).__name__


def _tool(name, **a):
    r = mcp_server.call_tool(name, a)
    return json.dumps(r, ensure_ascii=False, sort_keys=True)


@pytest.mark.parametrize("member, rng", [
    ("data/a.csv", {"rows": "5-900"}), ("data/a.csv", {"rows": "1999-"}), ("data/a.csv", {}),
    ("t.tsv", {"rows": "100-120"}), ("b.jsonl", {"records": "5-7"}), ("c.txt", {"paragraphs": "5-7"}),
])
def test_member_reads_the_same_range_as_the_file(box, member, rng):
    inside = _read(f"{box / 'a.zip'}!/{member}", rng)
    outside = _read(box / member.split("/")[-1], rng)
    assert inside == outside
    # 続きの印も同じに使える
    if inside[1]:
        assert _read(f"{box / 'a.zip'}!/{member}", inside[1]) == _read(box / member.split("/")[-1], inside[1])


def test_member_rows_are_the_rows_asked_for(box):
    text, _next, _cov, kind = _read(f"{box / 'a.zip'}!/data/a.csv", {"rows": "1500-1501"})
    assert kind == "CsvDoc"
    assert "| R1500 | 架空 1499 | 東京都架空区 1499 番地 | 149900 |" in text and "R1501" in text and "R1502" not in text and "| R2 |" not in text
    fl = api.fields_range(f"{box / 'a.zip'}!/data/a.csv", {"rows": "1500-1500"})["fields"]
    assert [f["id"] for f in fl] == ["R1500.C1", "R1500.C2", "R1500.C3", "R1500.C4"]
    assert fl == api.fields_range(str(box / "a.csv"), {"rows": "1500-1500"})["fields"]
    text, *_ = _read(f"{box / 'a.zip'}!/b.jsonl", {"records": "400-400"})
    assert "L400" in text and "架空 400" in text and "L401" not in text


def test_small_json_is_read_by_pointer(box):
    """窓を作って読む小さな JSON（ファイルも、書庫の中も）：path_in で頼まれた所を読む（前は、黙って窓の頭を返していた）。"""
    for p in (box / "d.json", f"{box / 'a.zip'}!/d.json"):
        text, nxt, _cov, kind = _read(p, {"path": "/a/3"})
        assert kind == "JsonLazy" and "`/a/3/n` ＝ 3" in text and "/a/4" not in text
        text, nxt, _cov, _k = _read(p, {"path": "/a"}, 300)
        assert nxt and str(nxt["cursor"]).startswith("json:")
        more, *_ = _read(p, nxt, 300)
        assert "/a/" in more and more != text
        fl = api.fields_range(str(p), {"path": "/a/3"})["fields"]
        assert [f["id"] for f in fl] == ["/a/3/n"]
        # 範囲を頼まなければ、前と同じに窓の部分で読む
        assert _read(p, {"part": 1})[3] == "BuiltDoc"
    r = _tool("read_window", path=f"{box / 'a.zip'}!/d.json", path_in="/a/7")
    assert "/a/7/n" in r and "/a/8" not in r


def test_range_names_of_another_kind_are_refused(box):
    """頼まれた範囲の名前が、その書類にないとき：黙って頭から読まない。"""
    z = str(box / "a.zip")
    cases = [(box / "c.txt", {"pages": "2"}), (f"{z}!/c.txt", {"pages": "2"}), (box / "a.csv", {"paragraphs": "2-3"}),
             (box / "a.csv", {"sheet": "a", "rows": "2-3"}), (f"{z}!/data/a.csv", {"path": "/a"}), (box / "b.jsonl", {"pages": "1"}),
             (box / "d.json", {"records": "1-2"}), (box / "d.json", {"pages": "1"}), (f"{z}!/d.json", {"rows": "1-2"})]
    for p, rng in cases:
        with pytest.raises(api.ApiError, match="範囲では読めません"):
            api.read_range(str(p), rng, 500)
        with pytest.raises(api.ApiError, match="範囲では読めません"):
            api.fields_range(str(p), rng)
    with pytest.raises(api.ApiError, match="pages（ページ）.*paragraphs"):
        api.read_range(str(box / "c.txt"), {"pages": "2"}, 500)
    # その書類の名前は、前と同じに読める（JSON Lines・.txt は、行の範囲のどの名前でも）
    assert "L5" in _read(box / "b.jsonl", {"rows": "5-5"})[0]
    assert "5 行目の字" in _read(box / "c.txt", {"rows": "5-5"})[0]


def test_non_utf8_small_json_says_it_cannot_be_read_by_pointer(tmp_path):
    p = tmp_path / "s.json"
    p.write_bytes(json.dumps({"名前": ["架空", "太郎"]}, ensure_ascii=False).encode("cp932"))
    with pytest.raises(api.ApiError, match="path_in（JSON Pointer）の範囲では読めません"):
        api.read_range(str(p), {"path": "/名前/1"}, 500)
    assert _read(p, {})[3] == "BuiltDoc"


def test_large_member_is_handled_like_the_file(tmp_path):
    """窓の全体が予算に入らない、書庫の中の CSV：概要・範囲の読み・欄・段階が、ファイルそのものと同じ。書き込みは ZIP に入れ直す。"""
    rows = ["氏名,住所,金額,備考"] + [f"架空 {i},東京都架空区 {i} 番地,{i * 100}," for i in range(1, 20001)]
    blob = ("\r\n".join(rows) + "\r\n").encode("cp932")
    (tmp_path / "big.csv").write_bytes(blob)
    with zipfile.ZipFile(tmp_path / "big.zip", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("data/big.csv", blob)
    f, m = str(tmp_path / "big.csv"), f"{tmp_path / 'big.zip'}!/data/big.csv"
    for tool, a in (("open_document", {}), ("read_window", {"rows": "15000-15003"}), ("list_fields", {"rows": "15000-15001"}),
                    ("grade_document", {})):
        assert _tool(tool, path=m, **a) == _tool(tool, path=f, **a), tool
    assert "20,001 行" in _tool("open_document", path=m)
    d1 = lazy.open_doc(m, force=True)
    assert d1 is lazy.open_doc(m, force=True)  # 読み手は控える（呼び出しのたびに、取り出して数え直さない）
    assert d1.member and d1.whole_only and not d1.streams
    out = tmp_path / "out.zip"
    r = api.fill(m, {"R15000.C4": "確認済み"}, str(out), force=True)
    assert r.get("read_back_ok", True) and r["member"] == "data/big.csv"
    with zipfile.ZipFile(out) as z:
        new = z.read("data/big.csv")
    a, b = blob.split(b"\r\n"), new.split(b"\r\n")
    assert len(a) == len(b) and [i + 1 for i, (x, y) in enumerate(zip(a, b)) if x != y] == [15000]
    assert b[14999].decode("cp932") == "架空 14999,東京都架空区 14999 番地,1499900,確認済み"


def test_member_reader_follows_the_archive_file(tmp_path):
    """書庫を書き換えたら、控えた読み手は使わない。"""
    z = tmp_path / "a.zip"
    for word in ("前", "後"):
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("a.csv", ("見出し\r\n" + "".join(f"{word} {i}\r\n" for i in range(1, 60 + len(word) * (word == "後") * 7))).encode("utf-8"))
        z.write_bytes(buf.getvalue())
        assert f"{word} 5" in _read(f"{z}!/a.csv", {"rows": "6-6"})[0]


def test_nested_member_and_choices(box, tmp_path):
    inner = (box / "a.zip").read_bytes()
    with zipfile.ZipFile(tmp_path / "n.zip", "w") as z:
        z.writestr("in.zip", inner)
        z.writestr("sc.csv", b"a;b;c\r\n1;2;3\r\n4;5;6\r\n")
    assert "| R5 | 架空 4 |" in _read(f"{tmp_path / 'n.zip'}!/in.zip!/data/a.csv", {"rows": "5-5"})[0]
    r = _tool("read_window", path=f"{tmp_path / 'n.zip'}!/sc.csv", rows="2-2", delimiter=";")
    assert "| R2 | 1 | 2 | 3 |" in r
    # 文字コードの選択は、その呼び出しだけ
    assert "架空" not in _read(f"{box / 'a.zip'}!/data/a.csv", {"rows": "2-2"}, encoding="euc-kr")[0]
    assert "架空 1" in _read(f"{box / 'a.zip'}!/data/a.csv", {"rows": "2-2"})[0]


def test_stateful_encoding_member(tmp_path):
    rows = ["氏名,住所"] + [f"架空 {i},東京都架空区 {i} 番地" for i in range(1, 300)]
    blob = "\r\n".join(rows).encode("iso2022_jp")
    (tmp_path / "iso.csv").write_bytes(blob)
    with zipfile.ZipFile(tmp_path / "i.zip", "w") as z:
        z.writestr("iso.csv", blob)
    inside, outside = _read(f"{tmp_path / 'i.zip'}!/iso.csv", {"rows": "100-101"}), _read(tmp_path / "iso.csv", {"rows": "100-101"})
    assert inside == outside and "| R100 | 架空 99 | 東京都架空区 99 番地 |" in inside[0]


def test_small_book_fields_follow_sheet_and_rows():
    """窓を作って読む表（小さな .xls など）：list_fields も、読むときと同じに、頼まれたシート・行の欄だけを出す
    （前は、sheet・rows を見ずに全部の欄を返し、ない名前のシートでも同じだった）。開いた範囲（「2-」）でも読める。"""
    p = "tests/fixtures/symbol_fonts.xls"
    ids = lambda rng: [f["id"] for f in api.fields_range(p, rng)["fields"]]  # noqa: E731
    assert ids({}) == ids({"sheet": "Sheet1"}) == ["Sheet1!B1", "Sheet1!B2", "Sheet1!B4", "Sheet1!B6"]
    assert ids({"sheet": "Sheet1", "rows": "4-4"}) == ["Sheet1!B4"]
    assert ids({"sheet": "Sheet1", "rows": "2-"}) == ["Sheet1!B2", "Sheet1!B4", "Sheet1!B6"]
    assert ids({"sheet": "Sheet1", "rows": "-2"}) == ["Sheet1!B1", "Sheet1!B2"]
    for rng, why in (({"sheet": "ない"}, "シート「ない」はありません"), ({"rows": "2"}, "sheet（シートの名前）といっしょに"),
                     ({"pages": "1"}, "範囲では読めません"), ({"paragraphs": "1"}, "範囲では読めません")):
        with pytest.raises(api.ApiError, match=why):
            api.fields_range(p, rng)
        with pytest.raises(api.ApiError, match=why):
            api.read_range(p, rng, 500)
    text = api.read_range(p, {"sheet": "Sheet1", "rows": "9-"}, 3000)[0].text
    assert "- 9: Empty" in text and "- 10: Rich" in text and "- 8: " not in text
    text = api.read_range(p, {"sheet": "Sheet1", "rows": "-2"}, 3000)[0].text
    assert "- 2: Decline" in text and "- 4: " not in text


def test_big_member_is_kept_in_a_temporary_file(tmp_path, monkeypatch):
    """メモリに置く境より大きな中身は、一時ファイルに置いて読む（結果は同じ。閉じると消える）。"""
    from todoku.lazy import textstream

    monkeypatch.setattr(textstream, "RECODE_IN_MEMORY", 1000)
    (tmp_path / "a.csv").write_bytes(CSV)
    with zipfile.ZipFile(tmp_path / "a.zip", "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("a.csv", CSV)
    m = f"{tmp_path / 'a.zip'}!/a.csv"
    assert _read(m, {"rows": "1500-1600"}) == _read(tmp_path / "a.csv", {"rows": "1500-1600"})
    doc = lazy.open_doc(m, force=True)
    src = doc.src
    assert isinstance(src, textstream.SpilledSource) and src.size == len(CSV) and src.read(0, 6) == CSV[:6]
    for key in [k for k, v in lazy._CACHE.items() if v is doc]:
        lazy._CACHE.pop(key)
    src.close()
    assert src.fd == -1 and src._tmp is None
