"""大きな書類を分けて読み書きする：読む量は減らしてよいが、読めない所を作らない。

小さな上限で、遅延の道を通す（速い合成のファイルだけ）。どの形式でも「続きをたどり終えると全部読んだことになる」を確かめる。
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import re
import zipfile

import pytest

from todoku import api, jobs, lazy, limits, mcp_server


@pytest.fixture(autouse=True)
def small_budget(monkeypatch, tmp_path):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)
    monkeypatch.setattr(limits, "PAGE_CHARS", 1_000)
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    lazy.clear()
    api._WINDOWS.clear()
    yield
    lazy.clear()
    api._WINDOWS.clear()


def _pages(path, **rng):
    """cursor をたどって、すべてのページを集める。"""
    out = []
    args = dict(rng)
    while True:
        page, _doc = api.read_range(str(path), args, args.pop("max_chars", None) if False else rng.get("max_chars"))
        out.append(page)
        if page.next is None:
            return out
        assert page.remaining, "残りを言わない部分返しはしない"
        assert "続きは read_window に" in page.footer()
        args = dict(page.next)


def _text(res):
    return "\n".join(c["text"] for c in res["content"] if c.get("type") == "text")


# ---------------------------------------------------------------------- CSV
def _make_csv(tmp_path, n=2000, newline="\r\n", enc="utf-8", name="a.csv"):
    p = tmp_path / name
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator=newline)
    w.writerow(["番号", "氏名", "メモ"])
    rows = [["番号", "氏名", "メモ"]]
    for i in range(1, n):
        memo = "" if i % 10 == 0 else (f"引用,{i}" if i % 7 == 0 else (f"改\n行{i}" if i % 11 == 0 else f"備考{i}"))
        w.writerow([i, f"利用者{i}", memo])
        rows.append([str(i), f"利用者{i}", memo])
    p.write_bytes(buf.getvalue().encode(enc))
    return p, rows


def _cells_of_page(page):
    got = []
    for line in page.body:
        m = re.match(r"\| R(\d+)(?:（見出し）)? \| (.*) \|$", line, re.S)
        assert m, line[:80]
        got.append((int(m[1]), m[2]))
    return got


def test_csv_large_open_is_overview_not_whole(tmp_path):
    p, rows = _make_csv(tmp_path)
    res = mcp_server.tool_result("open_document", {"path": str(p)})
    assert res["structuredContent"]["large"] is True
    assert res["structuredContent"]["rows"] == len(rows)
    assert "rows" in _text(res)
    assert "利用者1999" not in _text(res)
    assert len(_text(res)) < 3000


def test_csv_small_behaves_as_before(tmp_path, monkeypatch):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 10 ** 9)
    p, rows = _make_csv(tmp_path, 20)
    res = mcp_server.tool_result("open_document", {"path": str(p)})
    assert not res["structuredContent"].get("large")
    assert res["structuredContent"]["rows"] == len(rows)
    assert "利用者19" in _text(res)


def test_csv_pages_cover_everything(tmp_path):
    from todoku.datafile.window import row_line

    p, rows = _make_csv(tmp_path)
    pages = _pages(p)
    seen = [x for pg in pages for x in pg.body]
    want = [row_line(f"R{r}" + ("（見出し）" if r == 1 else ""), vals, 3) for r, vals in enumerate(rows, 1)]
    assert seen == want
    assert len(pages) > 3
    assert all(len(pg.text) < 1400 for pg in pages)


def test_csv_range_and_blank_detection(tmp_path):
    p, rows = _make_csv(tmp_path)
    page, _ = api.read_range(str(p), {"rows": "100-105"})
    assert page.next is None and [x[0] for x in _cells_of_page(page)] == list(range(100, 106))
    assert "利用者99" in page.text and "番号" in page.text  # 1 行目（見出し）を添える
    res = api.fields_range(str(p), {"rows": "100-125"}, only_empty=True)
    ids = {f["id"] for f in res["fields"]}
    assert ids == {f"R{r + 1}.C3" for r in range(100, 125) if r % 10 == 0}
    assert res["blank_cells_in_range"] == len(ids)
    with pytest.raises(api.ApiError, match="ありません"):
        api.read_range(str(p), {"rows": "999999"})


def test_csv_fields_truncation_has_a_continuation(tmp_path):
    p, rows = _make_csv(tmp_path, 300)
    got = []
    args = {"rows": "2-300"}
    while True:
        res = api.fields_range(str(p), args, limit=40)
        got += [f["id"] for f in res["fields"] if f["row"] <= 300]
        if not res["next"]:
            break
        args = res["next"]
    assert got == [f"R{r}.C{c}" for r in range(2, 301) for c in (1, 2, 3)]


@pytest.mark.parametrize("enc,newline", [("utf-8", "\r\n"), ("utf-8-sig", "\n"), ("cp932", "\r\n"), ("utf-16", "\r\n")])
def test_csv_stream_fill_matches_classic_bytes(tmp_path, enc, newline):
    """同じ書き込みを、遅延の道といままでの道で：バイトまで同じ。"""
    p, rows = _make_csv(tmp_path, 500, newline, enc)
    values = {"R3.C2": "山田,太郎", "R100.氏名": "鈴木", "R401.C3": "", "R10.C3": "改\n行を書く", "R501.C1": "足す"}
    a = tmp_path / "lazy.csv"
    r = api.fill(str(p), values, str(a))
    assert r.get("streamed") and r["proof_ok"] and r["read_back_ok"] and r["verified"], r
    assert r["read_back_by"] == "offset"
    limits.CONTEXT_BUDGET, keep = 10 ** 9, limits.CONTEXT_BUDGET
    try:
        lazy.clear()
        b = tmp_path / "classic.csv"
        r2 = api.fill(str(p), values, str(b))
        assert not r2.get("streamed") and r2["proof_ok"]
    finally:
        limits.CONTEXT_BUDGET = keep
    assert a.read_bytes() == b.read_bytes()
    assert r["new_sha256"] == hashlib.sha256(a.read_bytes()).hexdigest()
    # 読み戻し：書いた後のファイルを範囲で読む
    lazy.clear()
    page, _ = api.read_range(str(a), {"rows": "3-3"})
    assert "山田,太郎" in page.text


def test_csv_proof_catches_change_in_untouched_range(tmp_path):
    from todoku.lazy.base import stream_patch, verify_ranges
    from todoku.lazy.textstream import Source

    p, _rows = _make_csv(tmp_path, 100)
    src = Source(str(p))
    dst = tmp_path / "o.csv"
    ranges, _a, _b = stream_patch(src, [(10, 11, b"Z")], dst)
    assert verify_ranges(dst, ranges) == []
    data = bytearray(dst.read_bytes())
    data[-3] ^= 1
    dst.write_bytes(bytes(data))
    assert verify_ranges(dst, ranges)


def test_csv_quoted_newline_across_chunks(tmp_path, monkeypatch):
    from todoku.lazy import textstream

    monkeypatch.setattr(textstream, "CHUNK", 64)
    p, rows = _make_csv(tmp_path, 300)
    lazy.clear()
    doc = lazy.open_doc(str(p), force=True)
    assert doc.idx.count == len(rows)
    assert doc.width == 3


def test_csv_huge_cell_is_split_with_marker_and_complete(tmp_path):
    p = tmp_path / "h.csv"
    cell = "字" * 5000
    p.write_text(f"a,b\n1,{cell}\n2,x\n", encoding="utf-8")
    pages = _pages(p, max_chars=800)
    joined = "".join(pg.body[0] if "〔長い行" in pg.body[0] else "" for pg in pages if pg.body)
    assert "〔長い行なので分けて出しています" in pages[0].text or any("〔長い行" in pg.text for pg in pages)
    text = "".join(re.sub(r"〔[^〕]*〕", "", x) for pg in pages for x in pg.body)
    assert text.count("字") == 5000
    assert joined


# ---------------------------------------------------------------------- JSON Lines
def _make_jsonl(tmp_path, n=800):
    p = tmp_path / "a.jsonl"
    p.write_text("".join(json.dumps({"id": i, "name": f"利用者{i}", "memo": None if i % 7 == 0 else "備考"}, ensure_ascii=False) + "\n"
                         for i in range(n)), encoding="utf-8")
    return p


def test_jsonl_pages_cover_everything_and_fill(tmp_path):
    p = _make_jsonl(tmp_path)
    pages = _pages(p)
    body = [x for pg in pages for x in pg.body]
    assert len(body) == 800 and body[5].startswith("- `L6`") and "利用者5" in body[5]
    res = api.fields_range(str(p), {"records": "8-8"}, only_empty=True)
    assert [f["id"] for f in res["fields"]] == ["L8/memo"]
    out = tmp_path / "o.jsonl"
    r = api.fill(str(p), {"L8/memo": "書いた", "L9/new": 3}, str(out))
    assert r["proof_ok"] and r["read_back_ok"] and r["streamed"], r
    lines = out.read_text(encoding="utf-8").splitlines()
    assert json.loads(lines[7])["memo"] == "書いた" and json.loads(lines[8])["new"] == 3
    old = p.read_text(encoding="utf-8").splitlines()
    assert [i for i, (x, y) in enumerate(zip(lines, old)) if x != y] == [7, 8]


def test_jsonl_bad_line_stays_readable(tmp_path):
    p = tmp_path / "b.jsonl"
    p.write_text('{"a":1}\n{oops\n{"a":3}\n' * 200, encoding="utf-8")
    body = [x for pg in _pages(p) for x in pg.body]
    assert len(body) == 600 and "JSON として読めません" in body[1] and "{oops" in body[1]


# ---------------------------------------------------------------------- .txt
def _make_txt(tmp_path, n=400):
    p = tmp_path / "a.txt"
    p.write_text("".join(f"第{i}行 氏名（　　　）の申請です。\n" for i in range(1, n + 1)), encoding="utf-8")
    return p


def test_txt_lines_cover_everything_same_ids_as_classic(tmp_path, monkeypatch):
    p = _make_txt(tmp_path)
    pages = _pages(p)
    body = [x for pg in pages for x in pg.body]
    assert len(body) == 400
    assert body[99].startswith("100｜第100行")
    lazy_ids = []
    args = {}
    while True:
        res = api.fields_range(str(p), args)
        lazy_ids += [f["id"] for f in res["fields"]]
        if not res["next"]:
            break
        args = res["next"]
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 10 ** 9)
    lazy.clear()
    classic = [f["id"] for f in api.fields(str(p))]
    assert lazy_ids == classic


def test_txt_fill_by_id_and_anchor(tmp_path):
    p = _make_txt(tmp_path)
    fid = api.fields_range(str(p), {"paragraphs": "300"})["fields"][0]["id"]
    out = tmp_path / "o.txt"
    r = api.fill(str(p), {fid: "山田"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    lines = out.read_text(encoding="utf-8").splitlines()
    assert "山田" in lines[299]
    assert lines[:299] == p.read_text(encoding="utf-8").splitlines()[:299]
    out2 = tmp_path / "o2.txt"
    r = api.fill(str(p), {"@1": {"old": "第250行 氏名", "new": "第250行 名前"}}, str(out2))
    assert r["proof_ok"] and "第250行 名前" in out2.read_text(encoding="utf-8")
    with pytest.raises(api.ApiError, match="か所あります"):
        api.fill(str(p), {"@1": {"old": "の申請です", "new": "x"}}, str(tmp_path / "o3.txt"))


@pytest.mark.parametrize("enc", ["utf-8", "cp932", "utf-16"])
def test_huge_single_line_is_split_and_complete(tmp_path, monkeypatch, enc):
    from todoku.lazy import textdocs

    monkeypatch.setattr(textdocs, "SEG_CHARS", 500)
    unit = "長い行の文章です。"
    p = tmp_path / "h.txt"
    text = unit * 1200
    p.write_bytes(text.encode(enc))
    pages = _pages(p)
    assert len(pages) > 5
    assert all("長い行を分けて" in pg.text or "長い行：" in pg.text for pg in pages)
    assert "".join(re.sub(r"^\d+｜〔[^〕]*〕", "", x) for pg in pages for x in pg.body) == text


def test_huge_line_field_by_char_position(tmp_path, monkeypatch):
    from todoku.lazy import textdocs

    monkeypatch.setattr(textdocs, "SEG_CHARS", 500)
    p = tmp_path / "h.txt"
    p.write_text("あいう" * 400 + "氏名（　　　）" + "えおか" * 400, encoding="utf-8")
    fs = []
    args = {}
    while True:
        res = api.fields_range(str(p), args)
        fs += res["fields"]
        if not res["next"]:
            break
        args = res["next"]
    x = [f for f in fs if ".x" in f["id"]]
    assert x, fs[:3]
    out = tmp_path / "o.txt"
    r = api.fill(str(p), {x[0]["id"]: "山田"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert "氏名（山田）" in out.read_text(encoding="utf-8")
    r = api.fill(str(p), {"@1": {"old": "氏名（", "new": "名前（"}}, str(tmp_path / "o2.txt"))
    assert r["proof_ok"] and "名前（" in (tmp_path / "o2.txt").read_text(encoding="utf-8")


# ---------------------------------------------------------------------- 文字コードの判定は同じ
@pytest.mark.parametrize("blob", [
    "番号,氏名\n1,山田\n".encode("cp932"), "番号,氏名\n".encode("utf-8"), b"a,b\n1,2\n", "﻿番号\n".encode("utf-8"),
    "番号,氏名\r\n".encode("utf-16"), "번호,이름\n".encode("cp949"), bytes(range(0x80, 0xff)), "品川,東京\n".encode("euc_jp"),
    "番号\n1\n".encode("utf-16-le"), b"a\x00b\x00\n\x00"])
def test_stream_detection_equals_decode(tmp_path, blob):
    from todoku.datafile.textenc import decode
    from todoku.lazy.textstream import Source, detect

    p = tmp_path / "x.csv"
    p.write_bytes(blob)
    want = decode(blob)
    got = detect(Source(str(p)))
    assert (got.encoding, got.bom, got.basis, got.ascii_only, got.controls) == (want.encoding, want.bom, want.basis, want.ascii_only, want.controls)
    assert [c["encoding"] for c in got.candidates] == [c["encoding"] for c in want.candidates]
    assert [c.get("sample") for c in got.candidates] == [c.get("sample") for c in want.candidates]


# ---------------------------------------------------------------------- 段階と検査
def test_csv_grade_honest_and_audit_job(tmp_path):
    import time

    p, rows = _make_csv(tmp_path, 3000)
    g = api.grade(str(p))
    assert g["grade"] == "S" and g["deferred"] == []
    au = api.audit(str(p))
    assert au["job"]["key"]
    t0 = time.time()
    while not au["job"]["finished"] and time.time() - t0 < 60:
        time.sleep(0.1)
        au = api.audit(str(p))
    assert au["job"]["finished"] and au["ok"], au
    assert any("別の読み方" in c["detail"] for c in au["checks"])


def test_txt_grade_says_what_was_deferred(tmp_path):
    p = _make_txt(tmp_path)
    g = api.grade(str(p))
    assert any("未確認" in x for x in g["deferred"]) and g["provisional"]
    assert "未確認" in g["headline"]


def test_audit_job_resumes_from_checkpoint(tmp_path):
    p, rows = _make_csv(tmp_path, 3000)
    doc = lazy.open_doc(str(p), force=True)
    key = jobs.job_key(str(p), None)
    state = doc.audit_step(doc.audit_begin())
    assert not state["done"] or doc.idx.count < 16384
    jobs.save(key, {"key": key, "path": str(p), "state": state, "steps": 1, "pid": 999999})
    au = api.audit(str(p))
    assert au["job"]["resumed"] is True


# ---------------------------------------------------------------------- 上限
def test_memory_limit_scales_and_file_size_no_dead_end(monkeypatch, tmp_path):
    assert limits.MEMORY_LIMIT >= 2048 * limits.MB
    p, rows = _make_csv(tmp_path, 300)
    monkeypatch.setattr(limits, "MAX_FILE", 4)
    lazy.clear()
    # メモリに丸ごと読めない大きさでも、分けて読める形式は読める
    res = mcp_server.tool_result("read_window", {"path": str(p), "rows": "2-3"})
    assert not res.get("isError") and "利用者1" in _text(res)
    out = tmp_path / "o.csv"
    res = mcp_server.tool_result("fill_document", {"path": str(p), "values": {"R2.C3": "x"}, "out_path": str(out)})
    assert not res.get("isError"), _text(res)


def test_stop_names_smaller_request():
    from todoku import isolate

    r = isolate.stopped("メモリを超えました", "read_window", {"path": "/x/a.pdf"})
    assert "ページを分けて読んでください" in r["content"][0]["text"] and r["structuredContent"]["smaller_request"]
    r = isolate.stopped("時間", "open_document", {"path": "/x/a.xlsx"})
    assert "rows" in r["content"][0]["text"]


def test_inflate_bomb_still_stopped():
    import zlib

    bomb = zlib.compress(b"\0" * (64 * limits.MB), 9)
    with limits.call_budget(limits.Budget(part=limits.MB)):
        with pytest.raises(limits.LimitExceeded):
            limits.inflate(bomb)
    s = limits.StreamInflate("部品", 15, compressed=len(bomb))
    s.cap = limits.MB
    with pytest.raises(limits.LimitExceeded):
        s.feed(bomb)
