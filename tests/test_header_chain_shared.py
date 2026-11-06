"""上に続く見出しの並び（near の header）：何千行も字の続く列に結合の空きマスが並ぶシートでも、手間が行の数 × 欄の数にならない。

答えは、道を分け合わない素直なたどり方（下の _plain_chain）と 1 字も違わない。段の数・行の数で打ち切らない。"""

from __future__ import annotations

import copy
import json
import pickle

import pytest

from todoku import api
from todoku.xlsx import window as xw
from todoku.xlsx.window import _text

openpyxl = pytest.importorskip("openpyxl")

ROWS = 300


def _book(path, rows=ROWS):
    from openpyxl.styles import Border, Side

    line = Side(style="thin")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "一覧"
    ws["A1"] = "雑誌の一覧"
    ws["B1"] = "番号"
    ws["C1"] = "備考"
    ws.merge_cells("C1:D1")
    ws["F1"] = "広い見出し"
    ws.merge_cells("F1:H1")
    ws["F2"] = "左の列"
    ws["G2"] = "中の列"
    ws["H2"] = "右の列"
    ws.merge_cells("J1:J3")  # 字のない、縦の結合
    ws["J4"] = "縦の下"
    for r in range(3, rows + 3):
        ws.cell(r, 1, f"誌名{r}")
        ws.cell(r, 2, f"ISSN{r}")
        if r % 50 == 0:
            ws.cell(r, 2, r)  # 数の値：上に続く見出しは、ここを越えない
        # 字の列の上の、結合の空きマス（規則では欄と決められない所）
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.cell(r, 3).border = Border(left=line, right=line, top=line, bottom=line)
        ws.cell(r, 5).border = Border(left=line, right=line, top=line, bottom=line)
        if r % 7 == 0:
            ws.cell(r, 3, f"備考{r}")
        if r % 11 == 0:
            ws.cell(r, 6, f"列の字{r}")
        if r % 13 == 0:
            ws.merge_cells(start_row=r, start_column=7, end_row=r + 2, end_column=7)
            ws.cell(r, 7, f"縦の字{r}")
    ws.row_dimensions[5].height = 40
    ws.row_dimensions[9].hidden = True
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["E"].hidden = True
    wb.save(path)
    return path


def _plain_chain(lay, target, phrase):
    """素直なたどり方（欄ごとにシートの上まで見直す。答えの基準）。"""
    others = [b for b in lay.boxes(target.r) if b not in phrase and b.kind != "check"]
    lo = min(b.c1 for b in phrase) if phrase else target.c1
    hi = max(b.c2 for b in phrase) if phrase else target.c2
    out = []
    r = target.r - 1
    shape = lay.shape(target.r)
    span = (target.c1, target.c2)
    while r > 0:
        if not out:
            r = lay._skip_up(r, shape, span)
            if r <= 0:
                break
        r = lay._next_row(r, span)
        if r <= 0:
            break
        cand = lay._cand(r, span)
        if not out and any(b.kind == "input" for b in cand):
            break
        texts = [b for b in cand if _text(b)]
        if out and any(b.kind in ("input", "formula") or (b.kind == "text" and not isinstance(b.cell.value, str)) for b in cand):
            break
        if not texts:
            r -= 1
            continue
        h = texts[0]
        if not out:
            covered = [o for o in others if o.c1 <= h.c2 and o.c2 >= h.c1 and not (lo <= o.c1 and o.c2 <= hi)
                       and not (o.kind == "formula" and not o.text)]
            if covered:
                break
        if all(x is not h for _t, x in out):
            out.append((h.cell.value, h))
        span = (h.c1, h.c2)
        r = h.r - 1
    return out


def _layout(path):
    from todoku.book import open_book

    wb = open_book(path.read_bytes(), path.name)
    sh = wb.sheets[0]
    sh.__dict__[xw._MEMO] = {}
    return wb, sh, xw.Layout(wb, sh, xw._Options(wb).calc)


def test_chain_same_as_plain_walk_for_every_cell(tmp_path):
    wb, sh, lay = _layout(_book(tmp_path / "a.xlsx"))
    seen = deep = 0
    for r in range(1, sh.max_row + 2):
        for c in range(1, 12):
            target = xw.Box(r, c, r, c, "input", sh.cells.get((r, c)), "")
            want = _plain_chain(lay, target, [target])
            got = lay.header_chain(target, [target])
            assert [(t, id(x)) for t, x in got] == [(t, id(x)) for t, x in want], (r, c)
            facts = [lay._fact(x, target, t) for t, x in want]
            top = lay._header_top(target, [target])
            assert (lay._header_nodes().facts(top, r, c) if top >= 0 else []) == facts, (r, c)
            seen += 1
            deep = max(deep, len(want))
    assert seen > 3000 and deep >= 40  # 段を打ち切っていない（数の行のあいだの 49 行）


def test_near_header_is_built_only_when_read_and_reads_like_a_plain_dict(tmp_path):
    wb, sh, lay = _layout(_book(tmp_path / "a.xlsx"))
    box = next(b for b in lay.agent[120] if b.c1 == 3)
    d = lay.describe_agent(box)
    near = d["near"]
    assert isinstance(near, xw._Near) and isinstance(dict.get(near, "header"), xw._LazyHeader)
    # 値を読まない使い方では作らない
    assert list(near) and "header" in near and len(near) >= 1 and near not in (None, "", [], False)
    assert isinstance(dict.get(near, "header"), xw._LazyHeader)
    # 写しは、作らないまま写せる（段が何段あっても深くならない）
    blob = pickle.dumps(near, protocol=pickle.HIGHEST_PROTOCOL)
    back = pickle.loads(blob)
    assert isinstance(dict.get(back, "header"), xw._LazyHeader)

    fresh = pickle.loads(blob)
    plain = {k: (fresh[k]) for k in fresh}
    assert type(plain) is dict and isinstance(plain["header"], list) and len(plain["header"]) >= 10
    assert list(plain) == list(near)
    for kw in ({}, {"ensure_ascii": False}, {"indent": 1}, {"sort_keys": True}):  # C の JSON も Python の JSON も
        assert json.dumps(pickle.loads(blob), **kw) == json.dumps(plain, **kw)
    for make in (dict, lambda x: {**x}, lambda x: x.copy(), copy.copy, copy.deepcopy):
        got = make(pickle.loads(blob))
        assert got == plain and dict(got) == plain and list(got.values()) == list(plain.values())
    assert back == plain and plain == back and not (back != plain)
    assert pickle.loads(blob).get("header") == plain["header"]
    assert list(pickle.loads(blob).items()) == list(plain.items())
    assert pickle.loads(pickle.dumps(plain)) == pickle.loads(blob)


def test_range_read_work_grows_with_rows_not_rows_times_fields(tmp_path, monkeypatch):
    """1 回の範囲の読みで見る行の数：行 × 欄（欄ごとに上まで見直す）にならない。頼まれた行の外の欄の事実は作らない。"""
    p = _book(tmp_path / "big.xlsx", rows=1200)
    calls = {"next_row": 0, "fact": 0}
    real_next, real_fact = xw.Layout._next_row, xw._placed_fact

    def next_row(self, rr, span):
        calls["next_row"] += 1
        return real_next(self, rr, span)

    def fact(*a):
        calls["fact"] += 1
        return real_fact(*a)

    monkeypatch.setattr(xw.Layout, "_next_row", next_row)
    monkeypatch.setattr(xw, "_placed_fact", fact)
    page, _doc = api.read_range(str(p), {"sheet": "一覧", "rows": "1-1202"}, max_chars=25000)
    assert page.text
    assert calls["next_row"] < 40 * 1200  # 素直なたどり方では、1200 欄 × 数百行で何十万回
    assert calls["fact"] < 1000  # 読み順と目次だけ：欄の header の段の事実は作らない

    # 読み終えた後も、欄の一覧には段の事実がそのまま出る（打ち切っていない）
    res = api.fields_range(str(p), {"sheet": "一覧", "rows": "1190-1202"})
    deep = [f for f in res["fields"] if len((f.get("near") or {}).get("header") or []) > 1]
    assert deep


def test_list_fields_does_not_build_header_facts(tmp_path, monkeypatch):
    p = _book(tmp_path / "a.xlsx")
    calls = []
    real = xw._placed_fact
    monkeypatch.setattr(xw, "_placed_fact", lambda *a: (calls.append(1), real(*a))[1])
    fields = api.fields(str(p))
    assert fields
    made = len(calls)
    assert made < 1000  # 段の事実（欄の数 × 段の数）は、読まれるまで作らない
    text = json.dumps(fields, ensure_ascii=False)
    assert len(calls) > made and '"header"' in text
