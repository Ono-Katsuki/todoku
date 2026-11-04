"""PDF の行の所：所は行のもの（書く四角は行の帯、字は行の基線に、1 行で）。上下の空きは事実（room_mm）。

重なる書き込みは断る・書いた後の書類の監査・○の near の行と、○が約物・罫線に触れるときの知らせも、ここで確かめる。
どれも合成した様式（いくつかの言語）で確かめる。実際の様式の数や形に合わせた値は使わない。
"""

from __future__ import annotations

import pytest

from todoku.pdf.audit import run as audit
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import fill, FillError
from todoku.pdf.layout import MM_PT, View
from todoku.pdf.placement import appearance_glyphs
from todoku.pdf.window import build

from test_circle import _pdf, _write
from test_pdf_places import _rect, _row
from test_pdf_scripts import glyphs_pdf

LABELS = ("電話番号", "전화번호", "Phone", "Tel")


def _fields(blob: bytes) -> dict:
    return {f["id"]: f for pg in build(PdfDocument(blob, "t.pdf"))["pages"] for f in pg["fields"] if f.get("source") == "layout"}


def _baseline_of(blob: bytes, name: str, page: int = 0) -> tuple[list[float], float]:
    """書いた欄の字の基線（見た目の y の一覧）と、字の大きさ。"""
    doc = PdfDocument(blob, "t.pdf")
    view = View(doc, doc.pages()[page])
    for a in doc.resolve(doc.pages()[page].get("Annots")) or []:
        ad = doc.resolve(a)
        if str(getattr(ad.get("T"), "text", ad.get("T"))) == name or name in str(ad.get("T")):
            gs = appearance_glyphs(doc, ad)
            return [view.pt(g.ox, g.oy)[1] for g in gs], gs[0].size
    raise AssertionError(name)


def _tall(*rows) -> bytes:
    # 高さ 120pt のマスの中に、行がいくつか（マスの縁までの上下の空きが大きい）
    return glyphs_pdf([g for text, x, y in rows for g in _row(text, x, y)], extra=_rect(60, 580, 500, 700))


@pytest.mark.parametrize("label", LABELS)
def test_room_after_a_line_in_a_tall_cell_is_the_line_band(label):
    blob = _tall((label, 70, 650))
    trail = next(f for f in _fields(blob).values() if f.get("shape") == "trail")
    x0, y0, x1, y1 = trail["rect"]
    # 書く四角は行の帯：行の字の高さほど。マスの縁（上下の 50pt・70pt）までは広げない
    assert y1 - y0 < 2 * 12 and y0 < 842 - 650 < y1 + 12
    assert x1 == pytest.approx(500 - 1.5, abs=0.1)
    room = trail["facts"]["room_mm"]
    assert room["above"] > 10 and room["below"] > 10  # 上下の空きは事実として示す
    res = fill(PdfDocument(blob, "t.pdf"), {trail["id"]: "03-0000-0000"})
    ys, size = _baseline_of(res.blob, trail["id"])
    assert all(y == pytest.approx(842 - 650, abs=0.2) for y in ys), ys  # 行の字と同じ基線
    assert size == pytest.approx(12)  # 行の字の大きさ


def test_rows_of_a_cell_are_separate_places_and_both_can_be_written():
    blob = _tall(("住所の地名", 70, 670), ("氏名", 90, 640), ("電話番号の欄", 80, 610))
    trails = sorted((f for f in _fields(blob).values() if f.get("shape") == "trail"), key=lambda f: f["rect"][1])
    assert [f["facts"]["line"] for f in trails] == ["住所の地名", "氏名", "電話番号の欄"]
    assert not any(f.get("alternative_to") for f in trails), trails  # 別の行の所は、同じ所の別の読み方ではない
    res = fill(PdfDocument(blob, "t.pdf"), {f["id"]: "山田" for f in trails})
    for f, y in zip(trails, (670, 640, 610)):
        assert _baseline_of(res.blob, f["id"])[0][0] == pytest.approx(842 - y, abs=0.2)


def test_room_after_spaced_heading_is_a_place_even_with_text_below():
    # 字間を空けた見出しの行：行の中の空きのほかに、行の後ろの空きも所。下の行の字があっても消えない
    blob = _tall(("氏    名", 200, 670), ("この書類の説明の文です", 70, 650))
    trail = [f for f in _fields(blob).values() if f.get("shape") == "trail" and f["facts"].get("line", "").startswith("氏")]
    gap = [f for f in _fields(blob).values() if f.get("shape") == "gap" and f["facts"].get("line", "").startswith("氏")]
    assert trail and gap


@pytest.mark.parametrize("value", ("昭和60", "1985"))
def test_narrow_gap_shrinks_on_one_line_instead_of_wrapping(value):
    # 字の間の 7 mm の空き（「生年月日　　年」の年の前）
    w = 7 * MM_PT
    blob = glyphs_pdf(_row("生年月日", 70, 650) + _row("年", 70 + 48 + w, 650), extra=_rect(60, 630, 500, 670))
    gap = next(f for f in _fields(blob).values() if f.get("shape") == "gap")
    res = fill(PdfDocument(blob, "t.pdf"), {gap["id"]: value})
    ys, size = _baseline_of(res.blob, gap["id"])
    assert len({round(y, 1) for y in ys}) == 1, ys  # 1 行
    assert size < 12
    assert res.patches[-1].rect_mm["size_pt"] == pytest.approx(size, abs=0.05)


def test_narrow_gap_below_readable_warns_with_size_and_box():
    w = 7 * MM_PT
    blob = glyphs_pdf(_row("生年月日", 70, 650) + _row("年", 70 + 48 + w, 650), extra=_rect(60, 630, 500, 670))
    gap = next(f for f in _fields(blob).values() if f.get("shape") == "gap")
    res = fill(PdfDocument(blob, "t.pdf"), {gap["id"]: "昭和六十"})
    assert any(gap["id"] in x and "pt" in x and "box" in x for x in res.warnings), res.warnings
    with pytest.raises(FillError, match="box"):
        fill(PdfDocument(blob, "t.pdf"), {gap["id"]: "昭和六十年四月一日生まれの方です"})


def test_wrapping_only_in_a_tall_area_and_says_so():
    # 字のない高いマス（何行も入る所）：折り返すときは知らせる
    blob = glyphs_pdf(_row("備考", 70, 640), extra=_rect(60, 580, 120, 700) + "\n" + _rect(120, 580, 200, 700))
    cell = next(f for f in _fields(blob).values() if f.get("shape") == "cell")
    res = fill(PdfDocument(blob, "t.pdf"), {cell["id"]: "とても長い値をこの枠に書いて折り返します"})
    assert any("折り返しました" in x for x in res.warnings), res.warnings
    short = fill(PdfDocument(blob, "t.pdf"), {cell["id"]: "短い"})
    assert not any("折り返しました" in x for x in short.warnings), short.warnings


def test_written_values_whose_ink_overlaps_are_refused():
    # 行の所の字と、0.8 字ぶんだけ下にずれた枠の字：字の面が重なる（下線に触れたとしない低い四角では見のがす）
    blob = glyphs_pdf(_row("氏名", 70, 650), extra=_rect(60, 600, 500, 680))
    trail = next(f for f in _fields(blob).values() if f.get("shape") == "trail")
    mm = 72 / 25.4
    x0, x1 = 120 / mm, 300 / mm
    base_y = (842 - 650) / mm  # 行の基線（mm）
    # 枠の高さと字の大きさ：枠の中ほどに字を置く。字の大きさ 10.5pt（3.7 mm）の字の基線が、行の基線の 0.8 字ぶん下に来る枠
    box = [x0, base_y + 0.8 * 12 / mm - 3.7, x1, base_y + 0.8 * 12 / mm + 1.0]
    with pytest.raises(FillError, match="重なります"):
        fill(PdfDocument(blob, "t.pdf"), {trail["id"]: "山田太郎山田太郎", "@1": {"page": 1, "box": box, "value": "山田太郎山田太郎"}})
    # 1 行ぶん離れていれば書ける
    far = [x0, base_y + 2.0, x1, base_y + 7.0]
    fill(PdfDocument(blob, "t.pdf"), {trail["id"]: "山田太郎", "@1": {"page": 1, "box": far, "value": "山田太郎"}})


def test_audit_of_a_filled_document_rewrites_narrow_tall_cells(tmp_path):
    # 幅の狭い高いマス（桁の枠）：書いた後の書類の監査の試しの書き込みは、そのマスに入る値で
    cells = "\n".join(_rect(100 + 14 * k, 600, 114 + 14 * k, 645) for k in range(5))
    blob = glyphs_pdf(_row("口座番号", 40, 620), extra=cells)
    tcells = [f for f in _fields(blob).values() if f.get("shape") == "cell"]
    assert len(tcells) == 5
    res = fill(PdfDocument(blob, "t.pdf"), {f["id"]: "1" for f in tcells})
    a = audit(res.blob, "t.pdf")
    writing = next(c for c in a.checks if c.name == "全入力欄への書き込み")
    assert writing.ok, (writing.detail, writing.items)


# ====================================================================== ○
TWO_ROWS = [(70, 680, 10, "身体障害者手帳"), (200, 680, 10, "第　号"),
            (70, 660, 10, "療育手帳"), (200, 660, 10, "都・県")]


def test_circle_near_on_another_line_warns_with_position_and_line(tmp_path):
    r, _ = _write(tmp_path, _pdf(TWO_ROWS), "n.pdf", {"@c": {"circle": "都", "near": "身体障害者手帳"}})
    assert [c["text"] for c in r["circled"]] == ["都"]
    w = next(x for x in r["warnings"] if "行にありません" in x)
    assert "mm" in w and "【都】・県" in w and "身体障害者手帳" in w


def test_circle_near_prefers_the_same_line_in_a_shared_cell(tmp_path):
    # 見出しのマスが 2 行ぶんの高さ：マスの外枠では、どちらの行の字も同じくらい近い。near の字の行の字を囲む
    rules = (b"0.8 w 60 700 m 540 700 l S 60 640 m 540 640 l S 60 700 m 60 640 l S 190 700 m 190 640 l S 540 700 m 540 640 l S\n")
    data = _pdf([(70, 680, 10, "手帳の種類"), (200, 680, 10, "甲・乙"), (200, 660, 10, "丙・乙")], rules)
    r, _ = _write(tmp_path, data, "s.pdf", {"@c": {"circle": "乙", "near": "手帳の種類"}})
    assert r["circled"][0]["context"].startswith("甲") or "甲" in r["circled"][0].get("around", ""), r["circled"]
    assert not any("行にありません" in x for x in r["warnings"]), r["warnings"]


def test_circle_near_tie_on_the_line_is_refused(tmp_path):
    data = _pdf([(250, 680, 10, "区分"), (170, 680, 10, "甲・乙"), (300, 680, 10, "甲・乙")])
    src = tmp_path / "t.pdf"
    src.write_bytes(data)
    with pytest.raises(Exception, match="occurrence"):
        fill(PdfDocument(data, "t.pdf"), {"@c": {"circle": "乙", "near": "区分"}})


def test_circle_touching_a_middle_dot_or_rule_warns_with_mm(tmp_path):
    # 詰まった「有・無」：○が「・」にかかる所は、mm を添えて知らせる
    r, _ = _write(tmp_path, _pdf([(100, 680, 10, "無"), (107.5, 680, 10, "・")]), "d.pdf", {"@c": {"circle": "無"}})
    assert r["read_back_ok"]
    assert any("・" in x and "縁に触れます" in x and "mm" in x for x in r["warnings"]), r["warnings"]
    # 字のすぐ横の縦の太い罫線：線に触れるだけでも知らせる
    rule = b"3 w 111 660 m 111 700 l S\n"
    r, _ = _write(tmp_path, _pdf([(100, 680, 10, "無")], rule), "v.pdf", {"@c": {"circle": "無"}})
    assert any("罫線" in x and "mm" in x for x in r["warnings"]), r["warnings"]
