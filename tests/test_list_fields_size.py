"""list_fields の MCP の結果の大きさ：入らないときは読み減らすが、読めない所を作らない（省いたことと頼み方を、同じ結果に書く）。"""

from __future__ import annotations

import json

import pytest

from todoku import api, limits, mcp_server

openpyxl = pytest.importorskip("openpyxl")


def _book(path, rows=30):
    from openpyxl.styles import Border, Side
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表"
    line = Side(style="thin")
    ws["A1"] = "申請書の見出し"
    ws["B3"] = "有"  # 書いてある欄
    for r in range(3, rows + 3):
        ws.cell(r, 1, f"項目{r}")
        for c in range(2, 6):
            ws.cell(r, c).border = Border(left=line, right=line, top=line, bottom=line)
    dv = DataValidation(type="list", formula1='"有,無"')
    ws.add_data_validation(dv)
    dv.add("B3:B4")
    wb.save(path)
    return str(path)


def _call(**args):
    r = mcp_server.tool_result("list_fields", args)
    assert not r.get("isError"), r
    return r["structuredContent"]


def _walk(path, **args):
    """next をそのまま渡して、終わりまでたどる。"""
    pages = [_call(path=path, **args)]
    while pages[-1].get("next"):
        assert len(pages) < 200
        pages.append(_call(path=path, **pages[-1]["next"]))
    return pages


@pytest.fixture
def small_budget(monkeypatch):
    monkeypatch.setattr(limits, "RESULT_CHARS", 6000)
    return 6000


def test_result_that_fits_is_unchanged(tmp_path):
    p = _book(tmp_path / "a.xlsx", rows=3)
    r = mcp_server.tool_result("list_fields", {"path": p})
    whole = api.fields(p)
    assert r["structuredContent"] == {"fields": whole}
    assert r["content"][0]["text"] == json.dumps(whole, ensure_ascii=False, indent=1, default=str)


def test_too_large_keeps_what_writing_needs_and_says_how_to_get_the_rest(tmp_path, small_budget):
    p = _book(tmp_path / "a.xlsx")
    whole = {f["id"]: f for f in api.fields(p)}
    assert any("near" in f and "evidence" in f for f in whole.values())
    pages = _walk(p)
    for pg in pages:
        assert len(mcp_server._text(pg)["text"]) <= small_budget
        how = pg["omitted"]["detail"]["how"]
        assert "ids" in how and "detail" in how and "near" in how
    got = [f for pg in pages for f in pg["fields"]]
    assert len(pages) > 1 and [f["id"] for f in got] == list(whole)  # 分けても、どの欄も読める
    for f in got:
        # 省くのは根拠の事実だけ：ほかは一字も変えない
        assert f == {k: v for k, v in whole[f["id"]].items() if k not in mcp_server._FIELD_DETAIL}
    assert any(f.get("options") for f in got)
    counts = pages[0]["omitted"]["detail"]["count_by_name"]
    assert counts["near"] == sum(1 for f in whole.values() if "near" in f)
    assert "next" in pages[0]["omitted"]["more_fields"]


def test_only_empty_is_kept_while_going_on(tmp_path, small_budget):
    p = _book(tmp_path / "a.xlsx")
    whole = api.fields(p)
    empty = [f["id"] for f in whole if f.get("value") in (None, "", False)]
    assert len(empty) < len(whole)
    assert [f["id"] for pg in _walk(p, only_empty=True) for f in pg["fields"]] == empty


def test_ids_return_every_fact(tmp_path, small_budget):
    p = _book(tmp_path / "a.xlsx")
    whole = {f["id"]: f for f in api.fields(p)}
    slim = {f["id"] for pg in _walk(p) for f in pg["fields"]}
    one = next(i for i, f in whole.items() if "near" in f and i in slim)
    res = _call(path=p, ids=[one])
    assert res["fields"] == [whole[one]] and "near" in res["fields"][0]
    assert "omitted" not in res and "not_found" not in res
    res = _call(path=p, ids=[one, "表!Z999"])
    assert res["not_found"] == ["表!Z999"] and [f["id"] for f in res["fields"]] == [one]
    # たくさん頼んでも、分けて全部を返す
    pages = _walk(p, ids=list(whole))
    assert len(pages) > 1 and [f for pg in pages for f in pg["fields"]] == list(whole.values())
    assert all(len(mcp_server._text(pg)["text"]) <= small_budget for pg in pages)


def test_detail_returns_everything_in_steps(tmp_path, small_budget):
    p = _book(tmp_path / "a.xlsx")
    pages = _walk(p, detail=True)
    assert len(pages) > 1 and "detail" not in pages[0].get("omitted", {})
    assert [f for pg in pages for f in pg["fields"]] == api.fields(p)
    assert all(len(mcp_server._text(pg)["text"]) <= small_budget for pg in pages)


def test_range_request_is_kept_while_going_on(tmp_path, small_budget):
    p = _book(tmp_path / "a.xlsx")
    want = api.fields_range(p, {"sheet": "表", "rows": "3-25"})["fields"]
    pages = _walk(p, sheet="表", rows="3-25")
    assert len(pages) > 1 and all(pg["next"]["rows"] == "3-25" for pg in pages[:-1])
    assert [f["id"] for pg in pages for f in pg["fields"]] == [f["id"] for f in want]
    # 範囲と id でも、事実をすべて
    res = _call(path=p, sheet="表", rows="3-25", ids=[want[3]["id"]])
    assert res["fields"] == [want[3]]


def test_blank_cell_list_is_cut_with_how_to_ask(monkeypatch):
    monkeypatch.setattr(limits, "RESULT_CHARS", 2000)
    blanks = [f"表!B{i}" for i in range(1, 2001)]
    res = {"fields": [{"id": "表!A1", "label": "x"}], "blank_cells_in_range": blanks, "blank_cell_count": 2000, "next": None}
    rng = {"sheet": "表", "rows": "1-2000"}
    shown = mcp_server._fit_fields(res, {"path": "x", **rng}, rng)
    assert len(mcp_server._text(shown)["text"]) <= 2000
    assert shown["blank_cell_count"] == 2000 and shown["fields"] == res["fields"]
    assert shown["blank_cells_in_range"] == blanks[:len(shown["blank_cells_in_range"])]
    assert "rows" in shown["omitted"]["blank_cells_in_range"] and "2000" in shown["omitted"]["blank_cells_in_range"]


def test_ids_from_the_shorter_result_can_be_written(tmp_path, small_budget):
    p = _book(tmp_path / "a.xlsx")
    f = next(x for x in _call(path=p, only_empty=True)["fields"] if x.get("options"))
    out = tmp_path / "o.xlsx"
    r = mcp_server.tool_result("fill_document", {"path": p, "out_path": str(out), "values": {f["id"]: "有"}})
    assert r["structuredContent"]["proof_ok"] and r["structuredContent"]["read_back_ok"]


def test_fill_description_tells_about_circle_and_rewriting():
    d = next(t for t in mcp_server.TOOLS if t["name"] == "fill_document")["description"]
    head = d[:d.index("【値】")]  # はじめの段落（いつも読まれる所）に
    assert "○で囲める" in head and "help の circle" in head
    assert "old/new" in head and "help の text" in head
    assert mcp_server.call_tool("fill_document", {"help": "circle"})["content"][0]["text"].startswith("印刷された字を○で囲む")
    assert "old" in mcp_server.call_tool("fill_document", {"help": "text"})["content"][0]["text"]
