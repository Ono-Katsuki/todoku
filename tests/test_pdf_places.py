"""PDF の書く所：ひとつの決め方（作りだけ）・事実（mm）・位置から決める id（ほかの所が増減しても変わらない）。

どれも合成した様式（いくつかの言語・いくつかの組み方）で確かめる。実際の様式の数や形に合わせた値は使わない。
"""

from __future__ import annotations

import re

from todoku.pdf.audit import overprints
from todoku.pdf.content import page_content
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import fill
from todoku.pdf.layout import MM_PT, place_id
from todoku.pdf.window import build, to_markdown, unknown_id_message

from test_pdf_scripts import glyphs_pdf

MM = MM_PT


def _rect(x0, y0, x1, y1) -> str:
    return f"0.8 w {x0} {y0} m {x1} {y0} l S {x0} {y1} m {x1} {y1} l S {x0} {y0} m {x0} {y1} l S {x1} {y0} m {x1} {y1} l S"


def _row(text: str, x: float, y: float, step: float = 12) -> list[tuple[str, float, float]]:
    return [(ch, x + step * k, y) for k, ch in enumerate(text) if ch != " "]


def _fields(blob: bytes) -> list[dict]:
    return [f for pg in build(PdfDocument(blob, "t.pdf"))["pages"] for f in pg["fields"] if f.get("source") == "layout"]


def _table(label: str, extra_rows: list | None = None, extra_lines: str = "") -> bytes:
    # 見出しのマス（左）と空のマス（右）の 1 行の表
    lines = [_rect(60, 680, 160, 700), _rect(160, 680, 400, 700)]
    rows = _row(label, 64, 685) + (extra_rows or [])
    return glyphs_pdf(rows, extra="\n".join(lines) + "\n" + extra_lines)


# ====================================================================== 1. 空のマス：どれも同じ種類の所と、周りの字（mm）
def test_empty_cell_facts_in_mm_with_neighbours_in_every_direction():
    for label in ("氏名", "성명", "Name", "姓名"):
        f = next(x for x in _fields(_table(label)) if x.get("shape") == "cell")
        fa = f["facts"]
        assert fa["size_mm"] == [round(240 / MM, 1), round(20 / MM, 1)], fa
        assert fa["left"] == {"text": label, "mm": 0.0} and f["label"] == label
        assert fa["rect_mm"][:2] == [round(160 / MM, 1), round((842 - 700) / MM, 1)]
        assert not f.get("agent") and f["kind_candidates"] == ["text", "checkbox"]
        # 書く字の既定の大きさ：見出しにした左の字の大きさ（字のないマスには行の字の大きさがない）
        assert f["size"] == 12 and f["size_from"] == "左の字の大きさ"


def test_empty_cell_label_does_not_depend_on_distance_or_length():
    # 見出しのマスが遠くても・字が長くても、周りの字として示す（離れと字の数で見出しを捨てない）
    long = "この欄には申請者本人の現住所を住民票のとおりに書いてください"
    lines = [_rect(40, 600, 300, 640), _rect(320, 600, 560, 640)]  # 間に 20pt の離れ
    blob = glyphs_pdf(_row(long[:20], 44, 620) + _row(long[20:], 44, 606), extra="\n".join(lines))
    f = next(x for x in _fields(blob) if x.get("shape") == "cell")
    assert f["facts"]["left"]["text"].replace(" ", "") == long and f["facts"]["left"]["mm"] == round(20 / MM, 1)


def test_cells_under_a_header_row_take_the_header_as_above_and_rows_as_left():
    # 表の格子：列の見出し（上）と行の見出し（左）。間の空のマスを飛ばして、いちばん近い字のあるマスを示す
    lines = []
    for r in range(3):
        for c in range(3):
            lines.append(_rect(60 + 80 * c, 700 - 20 * (r + 1), 140 + 80 * c, 700 - 20 * r))
    rows = _row("月", 164, 685) + _row("火", 244, 685) + _row("朝", 64, 665) + _row("夜", 64, 645)
    fields = _fields(glyphs_pdf(rows, extra="\n".join(lines)))
    cells = {(f["facts"]["left"]["text"], f["facts"]["above"]["text"]): f for f in fields
             if f.get("shape") == "cell" and "left" in f["facts"] and "above" in f["facts"]}
    assert set(cells) == {("朝", "月"), ("朝", "火"), ("夜", "月"), ("夜", "火")}, cells
    assert cells[("夜", "火")]["facts"]["above"]["mm"] == round(20 / MM, 1)  # 間の空のマス（20pt）を飛ばした離れ
    md = to_markdown(build(PdfDocument(glyphs_pdf(rows, extra="\n".join(lines)), "t.pdf")))
    # 同じ大きさの空のマスのまとまり：id の作り（x・y）と、列の上の字・行の左の字の組
    assert "横 2 × 縦 2 に並ぶ" in md and "「月」" in md and "「夜」" in md, md


def test_equal_empty_cells_side_by_side_are_one_run_and_each_cell():
    lines = [_rect(60 + 20 * k, 680, 80 + 20 * k, 700) for k in range(5)]
    fields = _fields(glyphs_pdf(_row("〒", 40, 685), extra="\n".join(lines)))
    run = next(f for f in fields if f["kind"] == "comb")
    cells = [f for f in fields if f.get("shape") == "cell"]
    assert run["agent"] and run["facts"]["cells"] == 5 and len(cells) == 5
    assert all(c["alternative_to"] == run["id"] for c in cells) and sorted(run["cell_ids"]) == sorted(c["id"] for c in cells)


# ====================================================================== 2. 行の中の空き：1 字ぶん（1 em）以上の測った空き
def test_gaps_are_measured_in_em_in_several_scripts():
    cases = {"令和　年　月　日": 3, "　　년　　월　　일": 2, "Date:  /  /": 2, "省　　市　　县": 2}
    for text, want in cases.items():
        blob = glyphs_pdf(_row(text, 100, 700), {"　": 1000})
        gaps = [f for f in _fields(blob) if f.get("shape") == "gap"]
        assert len(gaps) == want, (text, [g["facts"] for g in gaps])
        assert all(g["facts"]["width_em"] >= 1 and g["agent"] for g in gaps)


def test_line_end_room_stops_at_a_cell_edge_and_after_a_sentence():
    # マスの中の見出しの右の空き：マスの縁まで（描かれた端）。文の終わりの字の後ろは、文の終わりなので所にしない
    # 文の終わりの字の後ろに閉じかっこが続く行（UAX #29：STerm Close*）も、文の終わり
    lines = [_rect(60, 640, 400, 700)]
    rows = _row("電話番号", 64, 685) + _row("記入してください。", 64, 665) + _row("（記入すること。）", 64, 645)
    fields = _fields(glyphs_pdf(rows, extra="\n".join(lines)))
    trail = [f for f in fields if f.get("shape") == "trail"]
    assert [f["facts"]["line"] for f in trail] == ["電話番号"], trail
    assert trail[0]["facts"]["bound"]["kind"] == "cell" and trail[0]["rect"][2] <= 400 - 1.5 + 0.01


def test_centred_title_between_drawn_edges_is_alignment_room():
    # 両側に描かれた端（マスの縁）があり、左右の空きの差が 1 字に満たない行：そろえ方の空き
    lines = [_rect(60, 680, 300, 700)]
    fields = _fields(glyphs_pdf(_row("申請書", 162, 685), extra="\n".join(lines)))
    assert not [f for f in fields if f.get("shape") in ("lead", "trail")], fields


def test_rtl_label_room_is_on_the_left():
    from test_pdf_order import _visual

    word = _visual("الاسم")
    lines = [_rect(40, 680, 340, 700)]
    blob = glyphs_pdf([(ch, 300 + 7 * k, 685) for k, ch in enumerate(word)], {ch: 583 for ch in set(word)}, extra="\n".join(lines))
    fields = _fields(blob)
    end = [f for f in fields if f.get("shape") == "trail"]
    assert len(end) == 1 and end[0]["rect"][2] <= 300 and end[0]["rect"][0] < 60, end  # 読む順の行の終わりは左


def test_underscore_and_bracket_are_drawn_blanks_and_box_is_left_to_the_agent():
    rows = _row("Signature ____", 60, 700, 7) + _row("( )", 60, 680, 7) + _row("□ Yes", 60, 660, 7)
    fields = _fields(glyphs_pdf(rows, {"_": 500, "(": 333, ")": 333, " ": 250}))
    by = {f.get("shape"): f for f in fields}
    assert not by["connector"].get("agent") and by["connector"]["facts"]["glyph"] == "_"
    assert not by["bracket"].get("agent")
    # 箱の字は、チェックの箱か箇条の印かを形で決めない（「？」。表計算・Word の窓と同じ）
    assert by["box"].get("agent") and by["box"]["facts"]["char"] == "□" and by["box"]["label"] == "Yes"


# ====================================================================== 3. 位置から決める id
def test_ids_come_from_position_and_kind():
    blob = _table("氏名")
    for f in _fields(blob):
        m = re.fullmatch(r"p1\.([tmgbu])(\d+)-(\d+)[a-z]?", f["id"])
        assert m, f["id"]
        if f.get("shape") == "cell":
            assert m[1] == "t" and (int(m[2]), int(m[3])) == (round(160 / MM), round((842 - 700) / MM))


def test_ids_do_not_shift_when_unrelated_places_are_added_or_removed():
    base = {f["id"]: f["facts"].get("left", {}).get("text") for f in _fields(_table("氏名"))}
    # ほかの所（上の行の日付の空き・別の空のマス）を足しても、元の所の id は同じ
    more = _table("氏名", _row("令和　年　月　日", 300, 760), _rect(60, 600, 200, 620))
    after = {f["id"] for f in _fields(more)}
    assert set(base) <= after, (base, after)
    assert len(after) > len(base)


def test_place_id_rounds_the_structural_anchor_to_mm():
    from todoku.pdf.layout import Spot

    sp = Spot("text", (0, 0, 10, 10), shape="cell", anchor=(28.35, 56.7))
    assert place_id(3, sp, MM) == "p3.t10-20"


def test_old_number_ids_are_kept_for_fields_written_before_and_refused_otherwise():
    # 前の版のとどくが「p1.3」の名前で書いた欄：その欄とほぼ同じ四角の所が 1 つだけなら、その所の id は「p1.3」のまま
    blob = _table("氏名")
    cell = next(f for f in _fields(blob) if f.get("shape") == "cell")
    res = fill(PdfDocument(blob, "t.pdf"), {cell["id"]: "山田"})
    # 名前だけを前の版の形に変えた書類を作る（増分の後ろの /T を書き換える）
    tagged = res.blob.replace(f"/T ({cell['id']})".encode(), b"/T (p1.3)")
    assert tagged != res.blob
    got = {f["id"]: f for f in _fields(tagged)}
    assert "p1.3" in got and got["p1.3"]["value"] == "山田" and cell["id"] not in got
    # 番号の名前を、窓にない所に渡したとき：読み替えずに断る知らせ
    assert "前の版" in unknown_id_message("p1.12") and "前の版" in unknown_id_message("p2.a7")
    assert unknown_id_message(cell["id"]) == ""


# ====================================================================== 4. 窓：1 行にまとめた事実、mm の印
def test_window_lists_line_places_per_line_with_measured_widths():
    blob = glyphs_pdf(_row("令和　　年　　月　　日", 100, 700), {"　": 1000})
    md = to_markdown(build(PdfDocument(blob, "t.pdf")))
    line = next(ln for ln in md.split("\n") if ln.startswith("- 行「令和"))
    assert line.count("字間") == 3 and "mm（2.0字）" in line, line


def test_room_below_text_in_a_cell_is_a_mark_in_mm():
    lines = [_rect(60, 620, 300, 700)]
    w = build(PdfDocument(glyphs_pdf(_row("住所", 64, 685), extra="\n".join(lines)), "t.pdf"))
    text = "\n".join(w["pages"][0]["reading"])
    assert re.search(r"［↓余白 \d+\.\d mm］", text), text
    assert "pt］" not in text


# ====================================================================== 5. 重ね書きの数え直し：窓と同じ決まり（layout.is_overprint）
def test_audit_counts_overprints_with_the_window_rule():
    blob = glyphs_pdf([("A", 100, 700), ("A", 100.2, 700), ("A", 120, 700)], {"A": 600})
    doc = PdfDocument(blob, "t.pdf")
    glyphs = page_content(doc, list(doc.pages())[0]).glyphs
    assert overprints(glyphs) == {"A": 1}


def test_grid_block_where_a_whole_row_is_left_out_of_the_group():
    """格子のある行のマスが、上・左の字が違うのでまとまりに入らないとき、窓が落ちない。"""
    from todoku.pdf.window import _grid_blocks

    def cell(x, y, above, left):
        return {"id": f"p1.t{x}-{y}", "page": 1, "source": "layout", "shape": "cell",
                "facts": {"rect_mm": [x, y, x + 10, y + 10], "size_mm": [10, 10],
                          "above": {"text": above, "mm": 1}, "left": {"text": left, "mm": 1}}}

    fields = [cell(10, 10, "甲", "一"), cell(20, 10, "乙", "一"), cell(10, 20, "丙", "二"), cell(20, 20, "丁", "二")]
    blocks, used = _grid_blocks(fields)
    assert used <= {f["id"] for f in fields}
