"""PDF の、読み順に入れない大きさの字（0.5 pt に満たない字・ページより何倍も大きい字）：字を落とさずに、分けて示す。

前は、その字の数だけを「読み順に入れていません」と伝え、字そのものは窓のどこにも出なかった（検査も、その字を数えなかった）。
いまは、ページの「見えない大きさの字」の節に、ひと続きに描いた字ごとに、位置（mm）と字の大きさ（pt）と一緒に示す：
読み順・入力欄には入れない。検査（読み落とし）は、その字も窓に出ていることを確かめる。
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from test_pdf_cover import MM, NAME, _audit_ok, _page, _reads_only_name, _write  # noqa: E402

from todoku import api, lazy, limits  # noqa: E402
from todoku.pdf import layout  # noqa: E402

TITLE = "### 見えない大きさの字"
TINY = b"BT /F1 0.2 Tf 150 720 Td (Secret words here) Tj 0 -0.3 Td (second line) Tj ET\n"


def _items(md: str) -> list[str]:
    """節の行（「- （左 …・上 … mm・… pt）字」）。"""
    i = md.find(TITLE)
    if i < 0:
        return []
    return [ln for ln in md[i:].split("\n### ")[0].split("\n") if ln.startswith("- （左 ")]


def test_text_smaller_than_half_a_point_is_shown_with_its_place_and_size(tmp_path):
    p = _write(tmp_path, NAME + TINY)
    pg, md = _page(p)
    assert _reads_only_name(pg) and pg["chars"] == len("Name:")  # 読み順・字の数には入れない
    assert [u["text"] for u in pg["unseen"]] == ["Secret words here", "second line"]  # 行が変わると、別のひと続き（上から）
    first, second = pg["unseen"]
    assert first["size_pt"] == second["size_pt"] == 0.2
    assert first["left_mm"] == second["left_mm"] == round(150 / MM, 1) and first["top_mm"] <= second["top_mm"]
    assert abs(first["top_mm"] - (pg["height"] - 720) / MM) < 0.2
    assert TITLE + "（25 字）" in md
    assert _items(md) == [f"- （左 {u['left_mm']}・上 {u['top_mm']} mm・0.2 pt）{u['text']}" for u in pg["unseen"]]
    assert "読み順に入れていません" not in md  # 前の、数だけの知らせ
    assert _audit_ok(p)


@pytest.mark.parametrize("how", [b"/F1 0 Tf 150 720 Td", b"/F1 12 Tf 0 Tz 150 720 Td", b"/F1 12 Tf 1 0 0 0 150 720 Tm"])
def test_text_that_has_no_ink_is_shown_as_size_zero(tmp_path, how):
    """大きさ 0・横の倍率 0・高さのない文字の行列：墨のない字。字は落とさない。"""
    p = _write(tmp_path, NAME + b"BT " + how + b" (Zero) Tj ET\n")
    pg, md = _page(p)
    assert _reads_only_name(pg)
    assert [(u["text"], u["size_pt"]) for u in pg["unseen"]] == [("Zero", 0.0)]
    assert len(_items(md)) == 1 and _items(md)[0].endswith("・0 pt）Zero")
    assert _audit_ok(p)


def test_a_glyph_many_times_larger_than_the_page_is_shown(tmp_path):
    p = _write(tmp_path, NAME + b"BT /F1 9000 Tf 150 720 Td (Big) Tj ET\n")
    pg, md = _page(p)
    assert _reads_only_name(pg)
    assert [(u["text"], u["size_pt"]) for u in pg["unseen"]] == [("Big", 9000.0)]
    assert _items(md)[0].endswith("・9000 pt）Big") and _audit_ok(p)


def test_text_of_half_a_point_or_more_stays_in_the_reading_order(tmp_path):
    p = _write(tmp_path, NAME + b"BT /F1 0.5 Tf 150 720 Td (Small) Tj ET\n")
    pg, md = _page(p)
    assert "unseen" not in pg and TITLE not in md
    assert any("Small" in ln for ln in pg["reading"]) and pg["chars"] == len("Name:Small")


@pytest.mark.parametrize("size, shown", [(b"0.497", 0.49), (b"0.003", 0.01), (b"0.25", 0.25)])
def test_the_size_is_not_rounded_to_zero_or_to_half_a_point(tmp_path, size, shown):
    pg, md = _page(_write(tmp_path, NAME + b"BT /F1 " + size + b" Tf 150 720 Td (Tiny) Tj ET\n"))
    assert [(u["text"], u["size_pt"]) for u in pg["unseen"]] == [("Tiny", shown)]
    assert _items(md)[0].endswith(f"・{shown:g} pt）Tiny")


def test_text_drawn_in_separate_text_objects_is_shown_separately(tmp_path):
    body = (b"BT /F1 0.3 Tf 150 720 Td (one) Tj ET\n" b"BT /F1 0.3 Tf 150.5 720 Td (two) Tj ET\n"
            b"BT /F1 0.3 Tf 150 700 Td (three ) Tj (four) Tj ET\n" b"BT /F1 0.3 Tf 150 690 Td (five) Tj /F1 0.4 Tf (six) Tj ET\n")
    pg, _md = _page(_write(tmp_path, NAME + body))
    # 同じ組で続けて描いた字は 1 つ、組が変わる・大きさが変わると別。字は描いた順のまま（基線の上から、左から）
    assert [(u["text"], u["size_pt"]) for u in pg["unseen"]] == [("one", 0.3), ("two", 0.3), ("three four", 0.3), ("five", 0.3), ("six", 0.4)]


def test_unseen_spaces_alone_are_not_told(tmp_path):
    p = _write(tmp_path, NAME + b"BT /F1 0.2 Tf 150 720 Td (   ) Tj ET\n")
    pg, md = _page(p)
    assert "unseen" not in pg and TITLE not in md and "読み順に入れていません" not in md
    assert pg["unseen_glyphs"] == 3  # 数は、前と同じに持つ
    assert _audit_ok(p)


def test_a_page_whose_only_text_is_unseen_points_to_the_section(tmp_path):
    p = _write(tmp_path, TINY)
    pg, md = _page(p)
    assert pg.get("blank") and pg["chars"] == 0
    assert "このページには字も画像もありません" in md and "見えない大きさの字は、下の節にあります" in md
    assert [u["text"] for u in pg["unseen"]] == ["Secret words here", "second line"]


def test_unseen_text_that_is_also_under_a_fill_is_shown_once(tmp_path):
    p = _write(tmp_path, NAME + TINY + b"140 700 100 50 re f\n")
    pg, md = _page(p)
    assert [u["text"] for u in pg["unseen"]] == ["Secret words here", "second line"]
    assert not pg.get("covered") and "### 塗りの下にあって見えない字" not in md
    assert _audit_ok(p)


def test_the_audit_fails_when_unseen_text_is_not_in_the_window(tmp_path, monkeypatch):
    """検査は、読み順に入れない大きさの字も、窓に出ていることを確かめる（前は、その字を数えなかった）。"""
    p = _write(tmp_path, NAME + TINY)
    assert _audit_ok(p)
    monkeypatch.setattr(layout, "unseen_runs", lambda glyphs, view: [])
    api._WINDOWS.clear()
    lazy.clear()
    assert not _audit_ok(p)
    api._WINDOWS.clear()
    lazy.clear()


def test_unseen_text_of_a_pdf_read_by_pages(tmp_path, monkeypatch):
    """ページの範囲で読む道（大きな PDF と同じ読み手）でも、同じ節が出る。"""
    p = _write(tmp_path, NAME + TINY)
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2_000)
    page, doc = api.read_range(p, {"pages": "1"}, 4000)
    assert type(doc).__name__ == "PdfLazy"
    assert TITLE + "（25 字）" in page.text and re.search(r"・0\.2 pt）Secret words here", page.text)
    lazy.clear()
