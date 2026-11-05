"""extract_processing・read_window の MCP の結果の大きさ：入らないときは読み減らすが、読めない所を作らない（省いたことと頼み方を、同じ結果に書く）。"""

from __future__ import annotations

import json

import pytest

from todoku import api, limits, mcp_server

openpyxl = pytest.importorskip("openpyxl")


def _book(path, sheets=2, shapes=30, rows=4):
    """形の違う数式の規則を shapes 個ずつ持つブック（1 つの規則は rows 行）。"""
    wb = openpyxl.Workbook()
    for s in range(sheets):
        ws = wb.active if s == 0 else wb.create_sheet()
        ws.title = f"表{s + 1}"
        ws["A1"] = "数量"
        for r in range(2, rows + 2):
            ws.cell(r, 1, r)
        for c in range(2, shapes + 2):
            ws.cell(1, c, f"項目{c}")
            for r in range(2, rows + 2):
                ws.cell(r, c, f"=$A{r}*{c}+{s}")
    wb.save(path)
    return str(path)


def _call(name="extract_processing", **args):
    r = mcp_server.tool_result(name, args)
    assert not r.get("isError"), r
    return r


def _chars(r) -> int:
    return sum(len(c["text"]) for c in r["content"])


def _walk(path, **args):
    pages = [_call(path=path, **args)]
    while pages[-1]["structuredContent"].get("next"):
        assert len(pages) < 300
        pages.append(_call(path=path, **pages[-1]["structuredContent"]["next"]))
    return pages


@pytest.fixture
def small(monkeypatch):
    monkeypatch.setattr(limits, "RESULT_CHARS", 6000)
    return 6000


# ---------------------------------------------------------------------- extract_processing
def test_extract_that_fits_is_unchanged(tmp_path):
    p = _book(tmp_path / "a.xlsx", sheets=1, shapes=2)
    r = _call(path=p)
    whole = api.extract_processing(p)
    assert r["structuredContent"] == whole
    assert [c["text"] for c in r["content"]][0] == whole["say_to_user"]
    assert "omitted" not in r["structuredContent"]


def test_api_output_does_not_depend_on_the_limit(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    res = api.extract_processing(p)
    assert len(res["rules"]) == 60 and "omitted" not in res and "reads" in res["rules"][0]


def test_extract_walks_every_rule_within_the_limit(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    whole = api.extract_processing(p)
    pages = _walk(p)
    assert all(_chars(r) <= small for r in pages)
    assert len(pages) > 1
    seen = [d for r in pages for d in r["structuredContent"]["rules"]]
    assert [d["id"] for d in seen] == [d["id"] for d in whole["rules"]]
    assert [(d["sheet"], d["range"], d["text"]) for d in seen] == [(d["sheet"], d["range"], d["text"]) for d in whole["rules"]]
    assert pages[0]["structuredContent"]["rules_total"] == len(whole["rules"])


def test_extract_keeps_the_ground_for_judging(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    out = tmp_path / "m.py"
    r = _call(path=p, format="python", out_path=str(out))
    res = r["structuredContent"]
    assert _chars(r) <= small
    assert r["content"][0]["text"] == res["say_to_user"] and "規則" in res["say_to_user"]
    assert res["summary"]["rules"] == 60 and res["limits"]
    assert res["python"]["out_path"] == str(out) and out.exists()
    v = res["verification"]
    assert v["rules_proven"] + v["rules_not_proven"] == 60 and "stored" in v


def test_python_is_written_once_and_next_does_not_write_again(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    out = tmp_path / "m.py"
    first = _call(path=p, format="python", out_path=str(out))
    before = out.stat().st_mtime_ns
    nxt = first["structuredContent"]["next"]
    assert "format" not in nxt and "out_path" not in nxt and "overwrite" not in nxt
    _call(path=p, **nxt)
    assert out.stat().st_mtime_ns == before


def test_omitted_rule_detail_can_be_fetched_by_id(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    whole = api.extract_processing(p)
    res = _call(path=p)["structuredContent"]
    assert "reads" not in res["rules"][0]
    detail = res["omitted"]["rule_detail"]
    assert detail["count_by_name"]["reads"] == 60 and "ids" in detail["how"]
    got = _call(path=p, ids=["R2", "R50", "ない"])["structuredContent"]
    assert got["rules"] == [whole["rules"][1], whole["rules"][49]]
    assert got["not_found"] == ["ない"]


def test_ids_too_large_for_one_result_continue_by_next(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    whole = api.extract_processing(p)
    want = [d["id"] for d in whole["rules"]]
    pages = _walk(p, ids=want)
    assert all(_chars(r) <= small for r in pages) and len(pages) > 1
    text = "".join(r["content"][1]["text"] for r in pages)
    assert json.loads(text)["rules"] == whole["rules"]


def test_every_omitted_section_can_be_fetched(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    whole = api.extract_processing(p)
    res = _call(path=p)["structuredContent"]
    names = res["omitted"]["sections"]["names"]
    for k, v in whole.items():
        if k == "rules" or k in res or v in (None, [], {}, "", False):
            continue
        assert k in names  # 黙って捨てない
        pages = _walk(p, section=k)
        assert all(_chars(r) <= small for r in pages)
        text = pages[0]["content"][0]["text"] if len(pages) == 1 else "".join(r["content"][1]["text"] for r in pages)
        assert json.loads(text) == {k: v}


def test_section_rules_gives_all_detail_in_order(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    whole = api.extract_processing(p)
    pages = _walk(p, section="rules")
    assert all(_chars(r) <= small for r in pages)
    assert json.loads("".join(r["content"][1]["text"] for r in pages))["rules"] == whole["rules"]


def test_sheet_narrows_rules_and_continues(tmp_path, small):
    p = _book(tmp_path / "a.xlsx", shapes=60)
    pages = _walk(p, sheet="表2")
    assert all(_chars(r) <= small for r in pages)
    seen = [d for r in pages for d in r["structuredContent"]["rules"]]
    assert len(seen) == 60 and {d["sheet"] for d in seen} == {"表2"}


def test_json_out_path_keeps_everything_and_says_where(tmp_path, small):
    p = _book(tmp_path / "a.xlsx")
    out = tmp_path / "all.json"
    r = _call(path=p, out_path=str(out))
    res = r["structuredContent"]
    assert _chars(r) <= small
    assert str(out) in res["omitted"]["all"]
    written = json.loads(out.read_text(encoding="utf-8"))
    assert len(written["rules"]) == 60 and "reads" in written["rules"][0]
    pages = _walk(p, **res["next"])
    assert sum(len(x["structuredContent"]["rules"]) for x in pages) + len(res["rules"]) == 60


def test_view_arguments_refuse_to_write(tmp_path):
    p = _book(tmp_path / "a.xlsx", sheets=1, shapes=2)
    out = tmp_path / "m.py"
    for extra in ({"skip": 1}, {"ids": ["R1"]}, {"section": "inputs"}):
        r = mcp_server.tool_result("extract_processing", {"path": p, "format": "python", "out_path": str(out), **extra})
        assert r["isError"] and "省いて" in r["content"][0]["text"]
    assert not out.exists()


def test_unknown_section_names_the_real_ones(tmp_path):
    p = _book(tmp_path / "a.xlsx", sheets=1, shapes=2)
    r = mcp_server.tool_result("extract_processing", {"path": p, "section": "ない節"})
    assert r["isError"] and "inputs" in r["content"][0]["text"]


def test_not_proven_rules_are_told_in_first_result(tmp_path, small, monkeypatch):
    p = _book(tmp_path / "a.xlsx")
    from todoku.processing import verify

    real = verify.run_and_compare

    def fake(*a, **k):
        rep = real(*a, **k)
        for i in range(1, 61):
            rep["rule_status"][f"R{i}"] = {"status": "proven", "counts": {}}
        rep["rule_status"]["R3"] = {"status": "not_proven", "why": ["比べたセルがない"], "counts": {}}
        return rep

    monkeypatch.setattr(verify, "run_and_compare", fake)
    res = _call(path=p, format="python", out_path=str(tmp_path / "m.py"))["structuredContent"]
    assert res["omitted"]["rule_verification"]["not_proven"]["R3"]["why"] == ["比べたセルがない"]
    assert next(d for d in res["rules"] if d["id"] == "R3")["verification"] == "not_proven"


def test_many_not_proven_rules_go_beside_the_module(tmp_path, small, monkeypatch):
    p = _book(tmp_path / "a.xlsx")
    from todoku.processing import verify

    real = verify.run_and_compare

    def fake(*a, **k):
        rep = real(*a, **k)
        for i in range(1, 61):
            rep["rule_status"][f"R{i}"] = {"status": "not_proven", "why": ["比べたセルがない"], "counts": {}}
        return rep

    monkeypatch.setattr(verify, "run_and_compare", fake)
    r = _call(path=p, format="python", out_path=str(tmp_path / "m.py"))
    told = r["structuredContent"]["omitted"]["rule_verification"]
    assert _chars(r) <= small and told["not_proven_count"] == 60
    assert len(json.loads(open(told["file"], encoding="utf-8").read())) == 60


def test_big_ground_values_are_cut_or_moved_and_told(tmp_path, small, monkeypatch):
    p = _book(tmp_path / "a.xlsx")
    real = api.extract_processing_both

    def fake(*a, **k):
        res, old = real(*a, **k)
        res = {**res, "verification": {"rules_proven": 60, "mismatch_count": 900,
                                       "mismatches": [{"cell": f"A{i}", "want": "x" * 20} for i in range(900)]},
               "cycles": [{"kind": "ordered", "cells": "y" * 40} for _ in range(900)]}
        return res, res

    monkeypatch.setattr(api, "extract_processing_both", fake)
    r = _call(path=p)
    res = r["structuredContent"]
    assert _chars(r) <= small
    assert res["verification"]["mismatch_count"] == 900
    assert res["omitted"]["cut"]["verification"]["lists"]["mismatches"]["total"] == 900
    assert len(res["verification"]["mismatches"]) < 900
    got = _walk(p, section="verification")
    assert json.loads("".join(x["content"][1]["text"] for x in got))["verification"]["mismatches"][899]["cell"] == "A899"


def test_one_huge_rule_still_fits(tmp_path, small, monkeypatch):
    p = _book(tmp_path / "a.xlsx", sheets=1, shapes=2)
    real = api.extract_processing_both

    def fake(*a, **k):
        res, old = real(*a, **k)
        res = {**res, "rules": [{**d, "text": "A = " + "長" * 30000} for d in res["rules"]]}
        return res, res

    monkeypatch.setattr(api, "extract_processing_both", fake)
    pages = _walk(p)
    assert all(_chars(r) <= small for r in pages)
    first = pages[0]["structuredContent"]["rules"][0]
    assert "ids" in first["text_cut"]
    got = _walk(p, ids=["R1"])
    assert all(_chars(r) <= small for r in got)
    assert len(json.loads("".join(x["content"][1]["text"] for x in got))["rules"][0]["text"]) == 30004


def test_pdf_form_result_fits(tmp_path, small, monkeypatch):
    real = api.extract_processing_both

    def fake(*a, **k):
        res = {"kind": "pdf", "summary": {"fields": 500, "rules": 400}, "say_to_user": "PDF のフォームの処理", "xfa": False,
               "fields": [{"name": f"欄{i}", "value": "x" * 30} for i in range(500)],
               "rules": [{"id": f"R{i}", "kind": "calculate", "field": f"欄{i}", "text": f"欄{i} = 1"} for i in range(400)]}
        return res, res

    monkeypatch.setattr(api, "extract_processing_both", fake)
    pages = _walk("f.pdf")
    assert all(_chars(r) <= small for r in pages)
    assert sum(len(r["structuredContent"]["rules"]) for r in pages) == 400
    assert "fields" in pages[0]["structuredContent"]["omitted"]["sections"]["names"]


# ---------------------------------------------------------------------- read_window
def _text_book(path, rows=120, width=40):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表"
    for r in range(1, rows + 1):
        ws.cell(r, 1, f"第{r}行：" + "あ" * width)
    wb.save(path)
    return str(path)


def _rw(path, **args):
    return mcp_server.tool_result("read_window", {"path": path, **args})


def test_window_that_fits_is_unchanged(tmp_path):
    p = _text_book(tmp_path / "a.xlsx", rows=5)
    r = _rw(p)
    _k, _d, parts = api.window(p)
    assert [c["text"] for c in r["content"]] == [f"部分 1/{len(parts)}：{parts[0][0]}", parts[0][1]]
    assert "structuredContent" not in r


def test_part_over_the_limit_is_cut_and_joins_to_the_same_text(tmp_path, small):
    p = _text_book(tmp_path / "a.xlsx")
    _k, _d, parts = api.window(p)
    assert len(parts) == 1 and len(parts[0][1]) > small  # 部分の番号は変わらない
    pages = [_rw(p, part=1)]
    while pages[-1]["structuredContent"].get("next"):
        pages.append(_rw(p, **pages[-1]["structuredContent"]["next"]))
    assert all(_chars(r) <= small for r in pages) and len(pages) > 1
    assert "".join(r["content"][1]["text"] for r in pages) == parts[0][1]
    assert pages[-1]["structuredContent"]["done"] is True


def test_cut_part_goes_on_to_the_next_part(tmp_path, small, monkeypatch):
    parts = [("目次", "目次の字"), ("シート", "\n".join("行" * 50 for _ in range(400))), ("残り", "残りの字")]
    monkeypatch.setattr(api, "window", lambda *a, **k: ("xlsx", {}, parts))
    monkeypatch.setattr(api, "lazy_open", lambda *a, **k: None)
    r = _rw("x.xlsx", part=2)
    pieces = [r]
    while "cursor" in (pieces[-1]["structuredContent"].get("next") or {}):
        pieces.append(_rw("x.xlsx", **pieces[-1]["structuredContent"]["next"]))
    assert "".join(x["content"][1]["text"] for x in pieces) == parts[1][1]
    assert pieces[-1]["structuredContent"]["next"] == {"part": 3}
    assert "part: 3" in pieces[-1]["content"][2]["text"]
    assert _rw("x.xlsx", part=3)["content"][1]["text"] == "残りの字"


def test_long_line_is_cut_without_losing_characters(tmp_path, small, monkeypatch):
    parts = [("全体", "長" * 20000)]
    monkeypatch.setattr(api, "window", lambda *a, **k: ("xlsx", {}, parts))
    monkeypatch.setattr(api, "lazy_open", lambda *a, **k: None)
    pages = [_rw("x.xlsx")]
    while pages[-1]["structuredContent"].get("next"):
        pages.append(_rw("x.xlsx", **pages[-1]["structuredContent"]["next"]))
    assert all(_chars(r) <= small for r in pages)
    assert "".join(r["content"][1]["text"] for r in pages) == parts[0][1]


def _range_walk(p, **args):
    pages = [_rw(p, **args)]
    while pages[-1]["structuredContent"].get("next"):
        assert len(pages) < 300
        pages.append(_rw(p, **pages[-1]["structuredContent"]["next"]))
    return pages


def test_range_reading_stays_within_the_limit_and_reads_all_rows(tmp_path, small):
    p = _text_book(tmp_path / "a.xlsx", rows=300)
    pages = _range_walk(p, sheet="表", rows="1-300")
    assert all(not r.get("isError") and _chars(r) <= small for r in pages)
    assert len(pages) > 1
    text = "".join(r["content"][0]["text"] for r in pages)
    for r in (1, 150, 300):
        assert f"第{r}行：" in text


def test_range_default_page_is_within_the_default_limit(tmp_path):
    p = _text_book(tmp_path / "a.xlsx", rows=2000)
    r = _rw(p, sheet="表", rows="1-2000")
    assert _chars(r) <= limits.RESULT_CHARS
    assert r["structuredContent"]["next"]


def test_max_chars_above_the_limit_is_held_to_the_limit(tmp_path, small):
    p = _text_book(tmp_path / "a.xlsx", rows=300)
    r = _rw(p, sheet="表", rows="1-300", max_chars=100000)
    assert _chars(r) <= small


def test_max_chars_below_the_limit_is_kept(tmp_path):
    p = _text_book(tmp_path / "a.xlsx", rows=300)
    r = _rw(p, sheet="表", rows="1-300", max_chars=1000)
    assert len(r["content"][0]["text"]) <= 1000


def test_overlong_page_from_a_range_is_cut_and_continues(tmp_path, small, monkeypatch):
    from todoku.lazy.base import Page

    text = "\n".join("行" * 50 for _ in range(400))

    def fake(path, rng, chars=None, pw=None, enc=None):
        if rng.get("cursor") == "after":
            return Page("おわり", "全部", None, "", ["おわり"]), None
        return Page(text, "1 ページ", {"pages": "2"}, "1 ページ", [text]), None

    monkeypatch.setattr(api, "read_range", fake)
    pieces = [_rw("x.pdf", pages="1")]
    while "cursor" in (pieces[-1]["structuredContent"].get("next") or {}):
        pieces.append(_rw("x.pdf", **pieces[-1]["structuredContent"]["next"]))
    assert all(_chars(r) <= small for r in pieces)
    assert "".join(r["content"][0]["text"] for r in pieces) == text
    assert pieces[-1]["structuredContent"]["next"] == {"pages": "2"}


def test_descriptions_stay_within_the_tool_limit():
    for t in mcp_server.TOOLS:
        assert len(t["description"]) <= 4096, t["name"]
