"""大きいから読み減らすのはよいが、読めないのは困る：埋め込みの表は頼まれた所だけ解析し、遠くの行のセルで遅くならず、時間のかかる窓は長い仕事になる。"""

from __future__ import annotations

import io
import time
import zipfile

import pytest

from todoku import api, isolate, jobs, limits, mcp_server
from todoku.embed import writable
from todoku.docx.document import DocxDocument
from todoku.sheetpic import Layout, SheetModel, _RowTops
from todoku.xlsx.places import fits_window
from todoku.xlsx.window import cols_mm, rows_mm
from todoku.xlsx.workbook import Sheet, Workbook
from todoku.metafile import Emf

from test_embed import _docx, _emf, _xlsx
from test_range_places import DV, cells, xlsx


def _with_far_row(book: bytes, r: int = 200000) -> bytes:
    """_xlsx の表のブックの、遠くの行に 1 マス足す。"""
    src = zipfile.ZipFile(io.BytesIO(book))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for i in src.infolist():
            data = src.read(i.filename)
            if i.filename == "xl/worksheets/sheet1.xml":
                data = data.replace(b"</sheetData>", f'<row r="{r}"><c r="A{r}"><v>1</v></c></row></sheetData>'.encode())
                data = data.replace(b"<oleSize", b"<x").replace(b"oleSize", b"x")
            z.writestr(i, data)
    return out.getvalue()


def _no_ole_size(book: bytes) -> bytes:
    src = zipfile.ZipFile(io.BytesIO(book))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for i in src.infolist():
            data = src.read(i.filename)
            if i.filename == "xl/workbook.xml":
                import re

                data = re.sub(rb"<oleSize[^>]*/>", b"", data)
            z.writestr(i, data)
    return out.getvalue()


# ------------------------------------------------------------------ 大きさの和：行の数によらない
def test_rows_and_cols_sum_match_one_by_one():
    wb = Workbook(_xlsx(5), "t.xlsx")
    sh = wb.sheets[0]
    sh.hidden_rows.add(7)
    sh.row_heights[9] = 40.0
    sh.hidden_cols.add(5)
    sh.col_widths[6] = 20.0
    from todoku.xlsx.window import col_mm, row_mm

    assert rows_mm(sh, 2, 30) == pytest.approx(sum(row_mm(sh, r) for r in range(2, 31)))
    assert rows_mm(sh, 5, 4) == 0  # 空の並び（隣どうしのマスのあいだ）
    assert cols_mm(wb, sh, 1, 20) == pytest.approx(sum(col_mm(wb, sh, c) for c in range(1, 21)))
    assert cols_mm(wb, sh, 3, 2) == 0


def test_row_tops_match_adding_one_by_one():
    sh = Sheet.__new__(Sheet)
    sh.__dict__.update(default_row_height=15.0, row_heights={3: 20.0, 8: 7.5}, hidden_rows={5})
    tops = _RowTops(sh, 2, 12)
    y = 0.0
    for r in range(2, 13):
        assert tops[r] == pytest.approx(y)
        y += 0.0 if r == 5 else sh.row_heights.get(r, 15.0)


# ------------------------------------------------------------------ 表の絵：遠くの行のセルで遅くならない
def test_picture_analysis_does_not_scale_with_distance():
    book = _no_ole_size(_with_far_row(_xlsx(5)))
    wb = Workbook(book, "t.xlsx")
    model = SheetModel(wb)
    assert model.range[2] == 200000
    t = time.perf_counter()
    assert (200000, 1) in set(model.display_cells()) and (2, 1) in set(model.display_cells())
    try:
        Layout(Emf(_emf()), model)
    except ValueError:
        pass  # 範囲が絵と合わないと断ることはある（遅くならないことを確かめる）
    assert time.perf_counter() - t < 2.0


def test_embedded_far_row_opens_quickly_and_is_read(tmp_path):
    p = tmp_path / "far.docx"
    p.write_bytes(_docx(_with_far_row(_xlsx(5)), _emf()))
    t = time.perf_counter()
    r = mcp_server.call_tool("open_document", {"path": str(p)})
    assert time.perf_counter() - t < 10
    assert not r.get("isError")
    _k, doc, _parts = api.window(str(p))
    assert doc["embeds"][0]["markdown"]
    assert any("A200000" in ln or "200000:" in ln for ln in doc["embeds"][0]["markdown"])


def test_far_cell_outside_the_picture_is_refused_with_the_reason(tmp_path):
    """映る範囲の記録がなく、遠くのセルまでが絵に収まらないときは、速くなっても、いままでどおり書けないと伝える（読むのは読める）。"""
    p = tmp_path / "far.docx"
    p.write_bytes(_docx(_no_ole_size(_with_far_row(_xlsx(5))), _emf()))
    d = DocxDocument(p.read_bytes(), "far.docx")
    t = time.perf_counter()
    ok, why = writable(d, d.embeds[0])
    assert time.perf_counter() - t < 5
    assert not ok and "収まりません" in why


def test_writing_an_empty_cell_still_changes_the_picture(tmp_path):
    """要素のないマスだけを見るようにしても、空のマスに書いた字は「変わったセル」になる（要素は新しいブックの側にある）。"""
    from todoku.sheetpic import changed_cells

    old, new = SheetModel(Workbook(_xlsx(), "a.xlsx")), SheetModel(Workbook(_xlsx(5), "b.xlsx"))
    assert (2, 2) in {rc for rc, _a, _b in changed_cells(old, new)}


# ------------------------------------------------------------------ 窓に入らない埋め込み：頼まれた行だけ読む
@pytest.fixture
def small_budget(monkeypatch):
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2000)  # 本文は入り、表は入らない大きさ
    api._WINDOWS.clear()
    yield
    api._WINDOWS.clear()


def _big_book() -> bytes:
    rows = "".join(f'<row r="{r}"><c r="A{r}" t="inlineStr"><is><t>行{r}の字</t></is></c><c r="B{r}"><v>{r}</v></c></row>' for r in range(3, 3001))
    return xlsx(cells(rows))


def test_large_embed_is_deferred_not_lost(tmp_path, small_budget):
    p = tmp_path / "e.docx"
    p.write_bytes(_docx(_big_book(), _emf()))
    k, doc, parts = api.window(str(p))
    e = doc["embeds"][0]
    assert e["deferred"] and e["sheets"] == ["S"] and e["est_chars"] > 2000
    text = "\n".join(t for _n, t in parts)
    assert 'sheet: "E1!S"' in text and "行2000の字" not in text
    # 頼んだ行だけ読める
    page, _doc = api.read_range(str(p), {"sheet": "E1!S", "rows": "2000-2001"})
    assert "行2000の字" in page.text and "行1999の字" not in page.text and "E1!" in page.text
    # next をたどれば、すべての行を読める
    got, rng = [], {"sheet": "E1!S"}
    for _ in range(3100):
        page, _doc = api.read_range(str(p), rng, 20000)
        got.append(page.text)
        if not page.next:
            break
        rng = page.next
    assert not page.next and all(f"行{r}の字" in "".join(got) for r in (3, 1500, 3000))
    # 欄の範囲の読み方と、書き込み（欄の名前は同じ）
    res = api.fields_range(str(p), {"sheet": "E1!S", "rows": "1-5"})
    assert res["sheet"] == "E1!S"
    # 段階は「まだ確かめていない」と伝える（確かめたことにしない）
    g = api.grade(str(p))
    assert any("まだ確かめていない" in r["title"] for r in g["write"]["reasons"])


def test_small_embed_stays_in_the_window(tmp_path):
    api._WINDOWS.clear()
    p = tmp_path / "e.docx"
    p.write_bytes(_docx(_xlsx(5), _emf()))
    _k, doc, _parts = api.window(str(p))
    e = doc["embeds"][0]
    assert not e.get("deferred") and e["writable"] and any(f["id"] == "E1!Sheet1!B2" or f["cell"] == "B2" for f in e["fields"])


def test_embed_write_still_works_when_deferred(tmp_path, small_budget, monkeypatch):
    p = tmp_path / "e.docx"
    p.write_bytes(_docx(_xlsx(), _emf()))
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 0)  # どの埋め込みも後回し
    _k, doc, _parts = api.window(str(p))
    assert doc["embeds"][0]["deferred"]
    sheet = api._word(*api._read(str(p), None)[:3]).embedded_books()[0][1].sheets[0].name
    r = api.fill(str(p), {f"E1!{sheet}!B2": 7}, str(tmp_path / "o.docx"))
    assert r["proof_ok"] and r["read_back_ok"]


def test_read_range_unknown_embed_says_so(tmp_path):
    p = tmp_path / "e.docx"
    p.write_bytes(_docx(_xlsx(5), _emf()))
    with pytest.raises(api.ApiError, match="E9"):
        api.read_range(str(p), {"sheet": "E9!S"})


def test_sheet_spec():
    from todoku.embeds.lazy import split_sheet

    assert split_sheet("E1!シート") == (1, "シート") and split_sheet("E12") == (12, None) and split_sheet("S") is None


def test_writable_does_not_need_the_window(tmp_path):
    p = tmp_path / "e.docx"
    p.write_bytes(_docx(_xlsx(5), _emf()))
    d = DocxDocument(p.read_bytes(), "e.docx")
    assert writable(d, d.embeds[0]) == (True, "")


# ------------------------------------------------------------------ 時間のかかる窓は、失敗でなく長い仕事
def test_time_limit_on_open_becomes_a_job(tmp_path, monkeypatch):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    p = tmp_path / "e.docx"
    p.write_bytes(_docx(_xlsx(5), _emf()))
    r = isolate.stopped(isolate._TIME(), "open_document", {"path": str(p)})
    assert not r["isError"] and r["structuredContent"]["job"]["running"] and "読めないのではありません" in r["content"][0]["text"]
    key = r["structuredContent"]["job"]["key"]
    for _ in range(600):
        rec = jobs.load(key)
        if rec and (rec.get("finished") or rec.get("stopped")):
            break
        time.sleep(0.1)
    assert rec.get("finished"), rec
    assert isolate.window_job_view("open_document", {"path": str(p)}) is None
    api._WINDOWS.clear()
    called = []
    monkeypatch.setattr(api, "_window", lambda *a, **k: called.append(a))  # 作り直さず、仕事の窓を使う
    k, doc, parts = api.window(str(p))
    assert k == "docx" and doc["embeds"] and not called
    api._WINDOWS.clear()


def test_a_job_that_was_stopped_is_told_not_restarted(tmp_path, monkeypatch):
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(tmp_path / "jobs"))
    p = tmp_path / "e.docx"
    p.write_bytes(_docx(_xlsx(5), _emf()))
    key = jobs._window_key(str(p), None)
    jobs.save(key, {"key": key, "kind": "window", "path": str(p), "stopped": "展開したデータの合計が上限を超えます"})
    r = isolate.stopped(isolate._TIME(), "open_document", {"path": str(p)})
    assert r["isError"] and "展開したデータの合計" in r["content"][0]["text"]


def test_deferred_embed_still_shows_its_charts():
    from todoku.embeds.lazy import markdown

    lines = markdown({"n": 1, "est_chars": 9, "sheets": ["S"], "drawn": ["### グラフ1：題（棒）"]})
    assert "#### グラフ1：題（棒）" in lines


def test_range_requests_and_passwords_do_not_become_jobs():
    assert not isolate._window_args("read_window", {"path": "/x.docx", "paragraphs": "1-5"})
    assert not isolate._window_args("open_document", {"path": "/x.docx", "password": "p"})
    assert not isolate._window_args("fill_document", {"path": "/x.docx"})
    assert isolate._window_args("open_document", {"path": "/x.docx"})


# ------------------------------------------------------------------ 宣言：窓に入るかで決める（決め打ちの広さを持たない）
def test_declaration_choice_follows_the_context_budget(monkeypatch):
    assert fits_window(3, 2, 2000, 2)
    assert not fits_window(3, 2, 150000, 2)
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 1000)
    assert not fits_window(3, 2, 2000, 2)  # 窓の予算を変えれば、変わる


def test_declaration_fields_and_place_by_budget():
    from todoku.xlsx.window import build

    dv = DV.replace("B3:B150000", "B3:B400")
    d = build(Workbook(xlsx(cells(f'<row r="500"><c r="A500"><v>1</v></c></row>') + dv), "t.xlsx"))
    sh = d["sheets"][0]
    assert not sh.get("range_places") and sum(1 for f in sh["fields"] if f["cell"].startswith("B")) >= 398
