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
    mc = rng.pop("max_chars", None)
    args = dict(rng)
    while True:
        page, _doc = api.read_range(str(path), args, mc)
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


# ---------------------------------------------------------------------- .xlsx
XL = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _make_xlsx(path, rows=400, extra_tail="", second=True):
    sst = [f"見出し{c}" for c in "ABCD"] + [f"利用者{i}" for i in range(rows)]
    si = "".join(f"<si><t>{s}</t></si>" for s in sst)
    body = ['<row r="1">' + "".join(f'<c r="{c}1" t="s"><v>{i}</v></c>' for i, c in enumerate("ABCD")) + "</row>"]
    for i in range(rows):
        r = i + 2
        cc = "" if i % 50 == 49 else f'<c r="C{r}"><v>{i * 3}</v></c>'
        body.append(f'<row r="{r}"><c r="A{r}"><v>{i}</v></c><c r="B{r}" t="s"><v>{i + 4}</v></c>{cc}<c r="D{r}"><f>A{r}*2</f><v>{i * 2}</v></c></row>')
    sheet = (f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{XL}"><cols><col min="1" max="4" width="12" customWidth="1"/></cols>'
             f'<sheetData>{"".join(body)}</sheetData>{extra_tail}</worksheet>')
    ct = ('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>')
    rel_ns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pk = "http://schemas.openxmlformats.org/package/2006/relationships"
    sheets = '<sheet name="データ" sheetId="1" r:id="rId1"/>' + ('<sheet name="二枚目" sheetId="2" r:id="rId4"/>' if second else "")
    wb = f'<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{XL}" xmlns:r="{rel_ns}"><sheets>{sheets}</sheets></workbook>'
    wr = (f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{pk}"><Relationship Id="rId1" Type="{rel_ns}/worksheet" Target="worksheets/sheet1.xml"/>'
          f'<Relationship Id="rId2" Type="{rel_ns}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="{rel_ns}/styles" Target="styles.xml"/>'
          + (f'<Relationship Id="rId4" Type="{rel_ns}/worksheet" Target="worksheets/sheet2.xml"/>' if second else "") + "</Relationships>")
    st = (f'<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{XL}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
          '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs></styleSheet>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", f'<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{pk}"><Relationship Id="rId1" Type="{rel_ns}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", wr)
        z.writestr("xl/styles.xml", st)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
        if second:
            z.writestr("xl/worksheets/sheet2.xml", f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{XL}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>二枚目の字</t></is></c><c r="B1"><f>データ!A2+1</f><v>1</v></c></row></sheetData></worksheet>')
        z.writestr("xl/sharedStrings.xml", f'<?xml version="1.0" encoding="UTF-8"?><sst xmlns="{XL}">{si}</sst>')
    return path


def test_xlsx_pages_cover_every_row_and_sheet(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", 300)
    res = mcp_server.tool_result("open_document", {"path": str(p)})
    assert res["structuredContent"]["large"] and res["structuredContent"]["sheets"][0]["rows"] == 301
    pages = _pages(p)
    text = "\n".join(pg.text for pg in pages)
    for i in range(300):
        assert f"利用者{i} " in text or f"利用者{i}\n" in text or f"利用者{i}｜" in text, i
    assert "二枚目の字" in text and "見出しA" in text
    assert all(len(pg.text) < 3500 for pg in pages)


def test_xlsx_row_range_and_fields_and_blanks(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", 300)
    page, _ = api.read_range(str(p), {"sheet": "データ", "rows": "100-102"})
    assert page.next is None and "利用者98" in page.text and "利用者200" not in page.text
    assert "見出しA" in page.text  # はじめの行を手がかりに添える
    res = api.fields_range(str(p), {"sheet": "データ", "rows": "50-52"})
    assert res["blank_cells_in_range"] == ["データ!C51"]
    with pytest.raises(api.ApiError, match="シートはありません"):
        api.read_range(str(p), {"sheet": "ない", "rows": "1-2"})


def test_xlsx_stream_fill_same_as_classic_and_recalculates(tmp_path, monkeypatch):
    p = _make_xlsx(tmp_path / "a.xlsx", 300)
    out = tmp_path / "o.xlsx"
    values = {"データ!A100": 1000, "データ!C51": "書いた", "データ!F400": "新しい行", "データ!C5": 7}
    r = api.fill(str(p), values, str(out))
    assert r["streamed"] and r["proof_ok"] and r["read_back_ok"] and r["verified"], r
    assert r["recalculated"] == {"データ!D100": 2000.0}, r["recalculated"]
    assert r["new_sha256"] == hashlib.sha256(out.read_bytes()).hexdigest()
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 10 ** 9)
    lazy.clear()
    api._WINDOWS.clear()
    b = tmp_path / "classic.xlsx"
    r2 = api.fill(str(p), values, str(b))
    assert not r2.get("streamed") and r2["proof_ok"]
    za, zb, zo = zipfile.ZipFile(out), zipfile.ZipFile(b), zipfile.ZipFile(p)
    assert za.namelist() == zb.namelist() == zo.namelist()
    assert za.testzip() is None
    for n in za.namelist():
        assert za.read(n) == zb.read(n), n  # いままでの道とバイトまで同じ
    # ほかの部品は生のバイトのまま
    assert za.read("xl/sharedStrings.xml") == zo.read("xl/sharedStrings.xml")
    from todoku.xlsx.workbook import Workbook

    wb = Workbook(out.read_bytes(), "o.xlsx")
    sh = wb.sheet("データ")
    assert sh.cells[(100, 1)].value == 1000 and sh.cells[(100, 4)].value == 2000 and sh.cells[(51, 3)].value == "書いた"
    assert sh.cells[(400, 6)].value == "新しい行" and sh.cells[(101, 2)].value == "利用者99"


def test_xlsx_dependency_closure_across_sheets_and_range(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", 200)
    r = api.fill(str(p), {"データ!A2": 5}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"], r
    assert r["recalculated"] == {"データ!D2": 10.0, "二枚目!B1": 6.0}


def test_xlsx_declared_cells_past_limit_are_reachable(tmp_path, monkeypatch):
    """宣言された広い範囲：上限を超えても、頼んだ行の範囲では欄として広げ、番地で書ける。"""
    from todoku.xlsx import window as xw

    tail = (f'<dataValidations count="1"><dataValidation type="whole" allowBlank="1" sqref="C1:C1048576"/></dataValidations>')
    p = _make_xlsx(tmp_path / "a.xlsx", 100, tail)
    monkeypatch.setattr(xw, "MAX_DECLARED_CELLS", 40)
    res = api.fields_range(str(p), {"sheet": "データ", "rows": "20-22"})
    ids = [f["id"] for f in res["fields"]]
    assert {"データ!C20", "データ!C21", "データ!C22"} <= set(ids), ids
    assert not any(re.search(r"C(1?\d|[3-9]\d)$", i) for i in ids if "!C" in i)
    r = api.fill(str(p), {"データ!G90": 5}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"] and r["read_back_ok"], r
    # 頼んだ範囲のほか（ファイルに行がない所）も番地で書ける
    r = api.fill(str(p), {"データ!A5000": 1}, str(tmp_path / "o2.xlsx"))
    assert r["proof_ok"] and r["read_back_ok"], r


def test_xlsx_audit_runs_in_steps(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", 300)
    doc = lazy.open_doc(str(p), force=True)
    state = doc.audit_begin()
    while not state["done"]:
        state = doc.audit_step(state)
        json.dumps(state)  # 控えに書ける
    assert all(c["ok"] for c in state["checks"]), state["checks"]
    assert state["formulas_ok"] == 300 and state["formulas_cross"] == 1
    assert state["still_deferred"]
    g = api.grade(str(p))
    assert g["provisional"] and any("未確認" in x for x in g["deferred"])


def test_xlsx_bomb_stops_in_stream(tmp_path):
    p = tmp_path / "b.xlsx"
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr("xl/workbook.xml", f'<workbook xmlns="{XL}"/>')
        z.writestr("xl/worksheets/sheet1.xml", b" " * (64 * limits.MB))
    from todoku.lazy.xlsxdoc import Member, package_kind

    assert package_kind(str(p)) == "xlsx"
    doc = lazy.textstream.Source(str(p))
    m = Member(doc, zipfile.ZipFile(p).getinfo("xl/worksheets/sheet1.xml"))
    keep = limits.MAX_INFLATE_PART
    limits.MAX_INFLATE_PART = limits.MB
    try:
        with pytest.raises(limits.LimitExceeded, match="爆弾"):
            for _ in m.chunks():
                pass
    finally:
        limits.MAX_INFLATE_PART = keep


# ---------------------------------------------------------------------- PDF
def _make_pdf(path, pages=12):
    import zlib

    objs = {3: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"}
    kids = []
    n = 4
    for i in range(pages):
        c, pg = n, n + 1
        n += 2
        raw = "\n".join(f"BT /F1 11 Tf 72 {760 - 14 * k} Td (Page {i + 1} line {k} text sample) Tj ET" for k in range(5)).encode()
        comp = zlib.compress(raw)
        objs[c] = b"<< /Length %d /Filter /FlateDecode >>\nstream\n" % len(comp) + comp + b"\nendstream"
        objs[pg] = b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 3 0 R >> >> /Contents %d 0 R >>" % c
        kids.append(pg)
    objs[1] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objs[2] = b"<< /Type /Pages /Count %d /Kids [" % pages + b" ".join(b"%d 0 R" % k for k in kids) + b"] >>"
    with open(path, "wb") as f:
        f.write(b"%PDF-1.5\n")
        off = {}
        for k in sorted(objs):
            off[k] = f.tell()
            f.write(b"%d 0 obj\n" % k + objs[k] + b"\nendobj\n")
        x = f.tell()
        f.write(b"xref\n0 %d\n0000000000 65535 f \n" % n)
        for k in range(1, n):
            f.write(b"%010d 00000 n \n" % off[k])
        f.write(b"trailer\n<< /Size %d /Root 1 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (n, x))
    return path


def test_pdf_lazy_open_pages_and_coverage(tmp_path, monkeypatch):
    from todoku.pdf.window import _md_page, build
    from todoku.pdf.doc import PdfDocument

    p = _make_pdf(tmp_path / "a.pdf")
    res = mcp_server.tool_result("open_document", {"path": str(p)})
    assert res["structuredContent"]["large"] and res["structuredContent"]["pages"] == 12
    assert "Page 12 line" not in _text(res)
    pages = _pages(p)
    got = "\n\n".join(x for pg in pages for x in pg.body)
    win = build(PdfDocument(p.read_bytes(), "a.pdf"))
    want = "\n\n".join("\n".join(_md_page(pg)) for pg in win["pages"])
    assert got == want
    one, _ = api.read_range(str(p), {"pages": "7-7"})
    assert "Page 7 line 0" in one.text and "Page 8" not in one.text and one.next is None


def test_pdf_small_max_chars_splits_page_with_no_loss(tmp_path):
    from todoku.pdf.window import _md_page, build
    from todoku.pdf.doc import PdfDocument

    p = _make_pdf(tmp_path / "a.pdf", 7)
    pages = _pages(p, max_chars=300)
    got = "".join(re.sub(r"\n〔\d+ ページ目はここで分けました〕$", "", x) for pg in pages for x in pg.body)
    win = build(PdfDocument(p.read_bytes(), "a.pdf"))
    want = "".join("\n".join(_md_page(pg)) for pg in win["pages"])
    assert got == want


def test_pdf_fill_only_lays_out_touched_page(tmp_path, monkeypatch):
    from todoku.pdf import window as pw

    p = _make_pdf(tmp_path / "a.pdf")
    calls = []
    real = pw.page_window

    def spy(doc, page, number, fields, layout=True):
        calls.append(number)
        return real(doc, page, number, fields, layout)

    monkeypatch.setattr(pw, "page_window", spy)
    lazy.clear()
    out = tmp_path / "o.pdf"
    r = api.fill(str(p), {"@1": {"after": "Page 9 line 4 text sample", "value": "ABC", "page": 9, "where": "below"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert out.read_bytes().startswith(p.read_bytes())
    assert "9" in r["empty_fields_note"]
    calls.clear()
    with pytest.raises(api.ApiError, match="page"):
        api.fill(str(p), {"@1": {"after": "Page 9 line 4", "value": "ABC"}}, str(tmp_path / "o2.pdf"))


def test_pdf_grade_says_deferred_and_audit_steps(tmp_path):
    p = _make_pdf(tmp_path / "a.pdf", 20)
    g = api.grade(str(p))
    assert "未確認：全ページの字の読み落とし" in g["deferred"] and "未確認" in g["headline"] and g["provisional"]
    doc = lazy.open_doc(str(p), force=True)
    state = doc.audit_begin()
    n = 0
    while not state["done"]:
        state = doc.audit_step(state)
        n += 1
    assert n == 21 and state["checks"][0]["ok"], state["checks"]


# ---------------------------------------------------------------------- Word・ほか
def _make_docx(path, n=200):
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    body = "".join(f'<w:p><w:r><w:t xml:space="preserve">第{i}段落 氏名（　　　）の説明です。</w:t></w:r></w:p>' for i in range(1, n + 1))
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        z.writestr("word/document.xml", f'<?xml version="1.0"?><w:document xmlns:w="{W}"><w:body>{body}<w:sectPr/></w:body></w:document>')
    return path


def test_word_paragraph_pages_cover_everything(tmp_path):
    p = _make_docx(tmp_path / "a.docx")
    res = mcp_server.tool_result("open_document", {"path": str(p)})
    assert res["structuredContent"]["large"] and res["structuredContent"]["paragraphs"] == 200
    pages = _pages(p, paragraphs="1-200")
    body = "\n".join(x for pg in pages for x in pg.body)
    assert all(f"第{i}段落" in body for i in range(1, 201))
    assert len(pages) > 2
    one, _ = api.read_range(str(p), {"paragraphs": "50-51"})
    assert "第50段落" in one.text and "第52段落" not in one.text and "［P50.c" in one.text
    res = api.fields_range(str(p), {"paragraphs": "50-50"})
    assert res["fields"] and all(f["id"].startswith("P50.") for f in res["fields"])
    r = api.fill(str(p), {res["fields"][0]["id"]: "山田"}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"], r


def test_built_documents_chars_cursor_covers_whole_window(tmp_path):
    """分けて読む道を持たない形式（ここでは JSON）でも、字の位置の cursor で全部を読める。"""
    p = tmp_path / "a.json"
    p.write_text(json.dumps({"items": [{"k": i, "v": f"値{i}"} for i in range(800)]}, ensure_ascii=False, indent=1), encoding="utf-8")
    pages = _pages(p)
    got = "".join(x for pg in pages for x in pg.body)
    doc = lazy.open_doc(str(p), force=True)
    assert got == doc.whole()
    assert all(f"値{i}" in got for i in range(800))
    # 1 行が長くても、字の位置で分けて進む
    q = tmp_path / "b.json"
    q.write_text(json.dumps({"s": "字" * 6000}, ensure_ascii=False), encoding="utf-8")
    assert "字" * 6000 in "".join(x for pg in _pages(q, max_chars=500) for x in pg.body)


def test_small_docs_also_accept_ranges(tmp_path, monkeypatch):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 10 ** 9)
    p, _rows = _make_csv(tmp_path, 30)
    res = mcp_server.tool_result("read_window", {"path": str(p), "rows": "5-6"})
    assert "利用者4" in _text(res) and "利用者9" not in _text(res)
    res = mcp_server.tool_result("list_fields", {"path": str(p), "rows": "11-12", "only_empty": True})
    assert [f["id"] for f in res["structuredContent"]["fields"]] == ["R11.C3"]
    q = _make_xlsx(tmp_path / "s.xlsx", 10)
    res = mcp_server.tool_result("read_window", {"path": str(q), "sheet": "データ", "rows": "3-4"})
    assert "利用者1" in _text(res)
    # part のいままでの読み方は変わらない
    res = mcp_server.tool_result("read_window", {"path": str(p)})
    assert "部分 1/1" in _text(res)


# ---------------------------------------------------------------------- 窓を作らずに読む・計算
def test_raw_reads_everything_without_a_window(tmp_path):
    p = _make_docx(tmp_path / "a.docx", 300)
    got = []
    args = {"raw": True}
    while True:
        page, _ = api.read_range(str(p), args, 500)
        got += page.body
        if page.next is None:
            break
        args = dict(page.next)
    text = "".join(got)
    assert all(f"第{i}段落" in text for i in range(1, 301))
    q = tmp_path / "b.json"
    src = json.dumps({"s": "字あ" * 4000}, ensure_ascii=False)
    q.write_text(src, encoding="utf-8")
    got, args = [], {"raw": True}
    while True:
        page, _ = api.read_range(str(q), args, 700)
        got += page.body
        if page.next is None:
            break
        args = dict(page.next)
    assert "".join(got) == src


def test_stop_for_built_format_names_raw():
    from todoku import isolate

    t = isolate.stopped("メモリ", "open_document", {"path": "/x/a.pptx"})["content"][0]["text"]
    assert "raw: true" in t
    t = isolate.stopped("時間", "audit_document", {"path": "/x/a.docx"})["content"][0]["text"]
    assert "job: true" in t


def test_xlsx_try_calculation_uses_only_needed_rows(tmp_path):
    p = _make_xlsx(tmp_path / "a.xlsx", 300)
    r = api.calculate(str(p), {"データ!A100": 1000})
    assert any(c["cell"].endswith("D100") and str(c["after"]).startswith("2000") for c in r["changed"]), r["changed"]
    e = api.calc_model(str(p), "データ!D100")
    assert "D100" in json.dumps(e, ensure_ascii=False)


def test_view_page_of_big_pdf_lays_out_one_page(tmp_path, monkeypatch):
    from todoku.pdf import window as pw

    p = _make_pdf(tmp_path / "a.pdf", 20)
    calls = []
    real = pw.page_window

    def spy(doc, page, number, fields, layout=True):
        calls.append(number)
        return real(doc, page, number, fields, layout)

    monkeypatch.setattr(pw, "page_window", spy)
    api._WINDOWS.clear()
    v = api.view_page(str(p), 7)
    assert v["png_base64"] and set(calls) <= {7}


def test_xlsx_shared_formula_master_outside_touched_rows(tmp_path):
    """共有数式の子のセルの計算：親の行が書く行の外でも、小さなブックに入れて計算し直す。空のセルの次のセルの式を取り違えない。"""
    p = _make_xlsx(tmp_path / "a.xlsx", 50)
    with zipfile.ZipFile(p) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    x = parts["xl/worksheets/sheet1.xml"].decode()
    x = re.sub(r"<f>A2\*2</f>", '<f t="shared" ref="D2:D51" si="0">A2*2</f>', x)
    x = re.sub(r"<f>A(\d+)\*2</f>", '<f t="shared" si="0"/>', x)
    x = x.replace('<c r="A30">', '<c r="E30" s="0"/><c r="A30">')
    parts["xl/worksheets/sheet1.xml"] = x.encode()
    with zipfile.ZipFile(p, "w", zipfile.ZIP_DEFLATED) as z:
        for n, b in parts.items():
            z.writestr(n, b)
    lazy.clear()
    r = api.fill(str(p), {"データ!A30": 100}, str(tmp_path / "o.xlsx"))
    assert r["proof_ok"], r
    assert r["recalculated"] == {"データ!D30": 200.0}


def test_part_on_large_stream_doc_points_to_ranges(tmp_path):
    p, _rows = _make_csv(tmp_path, 2000)
    res = mcp_server.tool_result("read_window", {"path": str(p), "part": 2})
    assert res.get("isError") and "範囲で読みます" in _text(res)


def test_password_is_not_a_cache_key():
    import inspect

    src = inspect.getsource(lazy._open_doc)
    assert "password is None" in src
