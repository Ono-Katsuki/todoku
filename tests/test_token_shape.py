"""MCP の結果の字を減らした所：減らしても、読める事実は減らない（行は 1 度だけ・すべての所は頼まれたとき・字下げなし）。"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from lines_back import restore
from todoku import api, limits, mcp_server

openpyxl = pytest.importorskip("openpyxl")

PUBLIC = Path(__file__).parent / "fixtures" / "public"
_FORMS = sorted(p.name for p in PUBLIC.iterdir() if p.suffix.lower() in (
    ".xlsx", ".xls", ".ods", ".docx", ".doc", ".rtf", ".odt", ".jtd", ".hwpx", ".pptx", ".csv", ".json", ".txt") or p.name in ("acroform.pdf", "flat.pdf"))


def _compact(obj) -> str:
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":"), default=str)


def _boxes(path, rows=20):
    """枠だけの空のマスが、1 行に 4 つずつ並ぶ表（入力欄と宣言する印はない）。"""
    from openpyxl.styles import Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "表"
    line = Side(style="thin")
    ws["A1"] = "申請書の見出し"
    for r in range(3, rows + 3):
        ws.cell(r, 1, f"項目{r}")
        for c in range(2, 6):
            ws.cell(r, c).border = Border(left=line, right=line, top=line, bottom=line)
    wb.save(path)
    return str(path)


# ====================================================================== 行は 1 度だけ
def test_line_of_moves_the_mark_only_to_one_place():
    head = "受付番号 ｜ ［▶B3？］ ｜ ［B30？］ ｜ ［W5:AJ5］ ｜ ［AD24：保護されたセル］"
    assert mcp_server._line_of(head, "表!B30") == "受付番号 ｜ ［B3？］ ｜ ［▶B30？］ ｜ ［W5:AJ5］ ｜ ［AD24：保護されたセル］"  # ［B3 は ［B30 の頭ではない
    assert mcp_server._line_of(head, "表!W5") == "受付番号 ｜ ［B3？］ ｜ ［B30？］ ｜ ［▶W5:AJ5］ ｜ ［AD24：保護されたセル］"
    assert mcp_server._line_of(head, "表!AD24") == "受付番号 ｜ ［B3？］ ｜ ［B30？］ ｜ ［W5:AJ5］ ｜ ［▶AD24：保護されたセル］"
    assert mcp_server._line_of(head, "表!C9") is None  # 行にない
    assert mcp_server._line_of("〔A1〕氏名［▶？　］　印［？　］", "表!A1@7") is None  # 番地のない所は決められない
    assert mcp_server._line_of("［B3］と［B3］", "表!B3") == "［B3］と［B3］"  # ▶ のない行は、そのまま
    assert mcp_server._line_of("［▶B3］と［B3］", "表!B3") is None  # 所が 2 つある


def test_share_lines_replaces_only_what_comes_back_exactly():
    row = "項目 ｜ ［▶B3？］ ｜ ［C3？］"
    fl = [{"id": "表!B3", "context": row}, {"id": "表!C3", "context": "項目 ｜ ［B3？］ ｜ ［▶C3？］"},
          {"id": "表!D3", "context": "項目 ｜ ［B3？］ ｜ ［C3？］ "},  # 1 字違う：行をそのまま返す
          {"id": "P1.c2", "context": "氏名［P1.c2？　］印［P1.c9？　］"}, {"id": "P1.c9", "context": "氏名［P1.c2？　］印［P1.c9？　］"},
          {"id": "P2.c1"}, {"id": "P1.c12", "context": "氏名［P1.c2？　］印［P1.c9？　］"}, {"id": "P3.c1", "context": ""}]
    before = _compact(fl)
    got = mcp_server._share_lines(fl)
    assert _compact(fl) == before  # 渡したものは変えない
    assert [("context_same_as" in f, f.get("context_same_as")) for f in got] == [
        (False, None), (True, "表!B3"), (False, None), (False, None), (True, "P1.c2"), (False, None), (True, "P1.c2"), (False, None)]
    assert list(got[1]) == ["id", "context_same_as"] and restore(got) == fl


@pytest.mark.parametrize("name", _FORMS)
def test_list_fields_gives_back_every_line(name, monkeypatch):
    p = str(PUBLIC / name)
    whole = api.fields(p)
    for budget in (limits.RESULT_CHARS, 3000):
        monkeypatch.setattr(limits, "RESULT_CHARS", budget)
        pages = [mcp_server.tool_result("list_fields", {"path": p})["structuredContent"]]
        while pages[-1].get("next"):
            assert len(pages) < 300
            pages.append(mcp_server.tool_result("list_fields", {"path": p, **pages[-1]["next"]})["structuredContent"])
        got = [f for pg in pages for f in restore(pg.get("fields") or [])]
        assert [f.get("id") for f in got] == [f.get("id") for f in whole]
        assert [f.get("context") for f in got] == [f.get("context") for f in whole]  # どの欄の行も、一字も違わずに戻る
        assert all(g == {k: v for k, v in w.items() if k in g} for g, w in zip(got, whole))  # 残した事実は、どれも元のまま
        for pg in pages:  # 行を持つ欄は、同じ結果の中の、前の欄
            seen = set()
            for f in pg.get("fields") or []:
                assert f.get("context_same_as") is None or f["context_same_as"] in seen
                seen.add(f.get("id"))


def test_shared_lines_make_the_result_smaller_and_detail_keeps_it(tmp_path):
    p = _boxes(tmp_path / "a.xlsx", rows=6)
    whole = api.fields(p)
    r = mcp_server.tool_result("list_fields", {"path": p})
    got = r["structuredContent"]["fields"]
    assert sum("context_same_as" in f for f in got) == 18 and sum("context" in f for f in got) == 6  # 6 行に 4 マスずつ
    assert len(_compact(r["structuredContent"])) < len(_compact({"fields": whole}))
    assert mcp_server.tool_result("list_fields", {"path": p, "detail": True})["structuredContent"]["fields"] == whole
    one = next(f["id"] for f in r["structuredContent"]["fields"] if "context_same_as" in f)
    by_id = mcp_server.tool_result("list_fields", {"path": p, "ids": [one]})["structuredContent"]["fields"]
    assert by_id == [next(f for f in whole if f["id"] == one)]  # ids は、行ごと返す


# ====================================================================== 字下げなし（上限に入る量は変えない）
def test_json_is_sent_without_indent_and_measured_as_before():
    obj = {"a": [1, {"b": "字"}], "c": None}
    assert mcp_server._text(obj)["text"] == '{"a":[1,{"b":"字"}],"c":null}'
    assert mcp_server._text("そのまま")["text"] == "そのまま"
    assert mcp_server._size(obj) == len(json.dumps(obj, ensure_ascii=False, indent=1, default=str)) > len(mcp_server._text(obj)["text"])


# ====================================================================== 段階：すべての所は、頼まれたとき
def _reasons(g: dict) -> list[dict]:
    return [*g["reasons"], *g["write"]["reasons"], *g.get("alerts", [])]


def test_grade_returns_all_places_only_when_asked(tmp_path):
    p = _boxes(tmp_path / "a.xlsx")
    whole = api.grade(p)
    many = [r for r in _reasons(whole) if "where_all" in r]
    assert many and all(len(r["where_all"]) > len(r["where"]) for r in many)  # 80 か所の理由がある
    r = mcp_server.tool_result("grade_document", {"path": p})
    short = r["structuredContent"]
    assert not [x for x in _reasons(short) if "where_all" in x] and "where_all: true" in short["omitted"]["where_all"]
    drop = lambda g: {**{k: v for k, v in g.items() if k != "omitted"},  # noqa: E731
                      "reasons": [{k: v for k, v in x.items() if k != "where_all"} for x in g["reasons"]],
                      "write": {**g["write"], "reasons": [{k: v for k, v in x.items() if k != "where_all"} for x in g["write"]["reasons"]]},
                      **({"alerts": [{k: v for k, v in x.items() if k != "where_all"} for x in g["alerts"]]} if "alerts" in g else {})}
    assert _compact(short | {"omitted": None}) == _compact(drop(whole) | {"omitted": None}) or _compact(drop(short)) == _compact(drop(whole))
    assert json.loads(r["content"][-1]["text"]) == json.loads(_compact(short))
    full = mcp_server.tool_result("grade_document", {"path": p, "where_all": True})["structuredContent"]
    assert _compact(full) == _compact(whole) and "omitted" not in full
    assert len(_compact(short)) < len(_compact(full))
    au = mcp_server.tool_result("audit_document", {"path": p})["structuredContent"]
    if isinstance(au.get("grade"), dict):  # 検査の結果の中の段階も同じ形
        assert not [x for x in _reasons(au["grade"]) if "where_all" in x]


def test_grade_without_many_places_says_nothing_is_omitted(tmp_path):
    p = _boxes(tmp_path / "a.xlsx", rows=2)
    whole = api.grade(p)
    assert not [r for r in _reasons(whole) if "where_all" in r]
    short = mcp_server.tool_result("grade_document", {"path": p})["structuredContent"]
    assert "omitted" not in short and _compact(short) == _compact(whole)


def test_grade_with_all_places_is_split_when_it_does_not_fit(tmp_path, monkeypatch):
    p = _boxes(tmp_path / "a.xlsx", rows=60)
    whole = _compact(api.grade(p))
    monkeypatch.setattr(limits, "RESULT_CHARS", 3000)
    short = mcp_server.tool_result("grade_document", {"path": p})
    assert "text" not in short["structuredContent"] or short["structuredContent"]["next"]  # 入らなければ、これも分ける
    pieces, ask = [], {"where_all": True}
    while ask:
        assert len(pieces) < 100
        r = mcp_server.tool_result("grade_document", {"path": p, **ask})
        sc = r["structuredContent"]
        assert len(_compact(sc)) <= 3000 and r["content"][1]["text"] == sc["text"]
        pieces.append(sc["text"])
        ask = sc["next"]
        assert sc["done"] == (ask is None) and (ask is None or ask["where_all"] is True)
    assert len(pieces) > 1 and "".join(pieces) == whole


# ====================================================================== 分けた JSON は、structuredContent でも読める
def test_sliced_json_text_is_in_structured_content(monkeypatch):
    monkeypatch.setattr(limits, "RESULT_CHARS", 1500)
    obj = {"rules": [{"name": f'引用符 " と \\ と改行\nのある値 {i}', "v": "あ" * 30} for i in range(120)]}
    whole, pieces, skip = _compact(obj), [], 0
    while skip is not None:
        assert len(pieces) < 200
        r = mcp_server._slice_json(obj, "規則", {"ids": ["rules"]}, skip)
        sc = r["structuredContent"]
        assert len(_compact(sc)) <= 1500  # \ の付く字があっても、上限に入る
        pieces.append(sc["text"])
        skip = sc["next"]["skip"] if sc["next"] else None
    assert "".join(pieces) == whole and json.loads("".join(pieces)) == obj


# ====================================================================== 本人に伝える字は 1 度だけ
def test_fill_says_it_once_in_the_text_and_keeps_it_in_the_result(tmp_path):
    p = _boxes(tmp_path / "a.xlsx", rows=3)
    r = mcp_server.tool_result("fill_document", {"path": p, "values": {"表!B3": "山田花子"}, "out_path": str(tmp_path / "o.xlsx")})
    assert not r.get("isError"), r
    say = r["structuredContent"]["say_to_user"]
    assert say and "山田花子" in say
    assert sum(c["text"].count(say) for c in r["content"]) == 1  # 字の結果（content）に 2 度は書かない
    assert sum(_compact(say)[1:-1] in c["text"] for c in r["content"]) == 0  # JSON の中にも、もう一度は入れない
    rest = json.loads(r["content"][-1]["text"])
    assert rest == {k: v for k, v in json.loads(_compact(r["structuredContent"])).items() if k != "say_to_user"}


# ====================================================================== 道具の説明
def test_tools_tell_the_new_shapes():
    by = {t["name"]: t for t in mcp_server.TOOLS}
    assert "context_same_as" in by["list_fields"]["description"]
    g = by["grade_document"]
    assert "where_all" in g["description"] and {"where_all", "skip"} <= set(g["inputSchema"]["properties"])
    full = by["open_document"]["inputSchema"]["properties"]["encoding"]["description"]
    for name in ("read_window", "list_fields", "fill_document", "audit_document", "grade_document"):
        d = by[name]["inputSchema"]["properties"]["encoding"]["description"]
        assert "encoding" in d and len(d) < len(full)  # 選び方は open_document とサーバーの説明に。ここは短く
    assert "encoding" in mcp_server.INSTRUCTIONS and "文字コードを規則では決められません" in full
