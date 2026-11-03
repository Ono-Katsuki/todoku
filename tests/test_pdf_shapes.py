"""PDF の読む窓：桁の枠の最後のマス・図形で描いた字・注釈の見た目の字（23 区の試しで見つけたもの）。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pdfmini  # noqa: E402
from test_pdf_scripts import flags, glyphs_pdf, window  # noqa: E402


def test_comb_run_keeps_wider_last_box():
    # 口座番号：幅 20 のマス 6 個と、表の縁まで広げた幅 30 の最後のマス
    xs = [200 + 20 * i for i in range(7)] + [350]
    lines = ["200 700 m 350 700 l S", "200 680 m 350 680 l S"] + [f"{x} 680 m {x} 700 l S" for x in xs]
    lines += ["100 680 m 100 700 l S", "100 700 m 200 700 l S", "100 680 m 200 680 l S"]
    blob = glyphs_pdf([(ch, 110 + 12 * k, 685) for k, ch in enumerate("口座番号")], extra="\n".join(lines))
    combs = [f for f in window(blob)["pages"][0]["fields"] if f["kind"] == "comb"]
    assert combs and combs[0]["boxes"] == 7


def _glyph_shape(x: float, y: float) -> str:
    # 字の大きさほどの、曲線を含む塗りの形
    return f"{x} {y} m {x + 8} {y} l {x + 8} {y + 4} {x + 4} {y + 9} {x} {y + 9} c h f"


def test_text_drawn_as_shapes_is_flagged_and_not_candidates():
    shapes = "\n".join(_glyph_shape(60 + 10 * k, 780 - 14 * r) for r in range(3) for k in range(12))
    blob = glyphs_pdf([(ch, 60 + 12 * k, 600) for k, ch in enumerate("委任状")], extra=shapes)
    f = flags(blob)
    assert "図形で描いた字かもしれない所（字の情報がない）" in f
    fields = window(blob)["pages"][0]["fields"]
    assert not [x for x in fields if x["rect"][1] < 842 - 740]


DRAWN = "図形で描いた字かもしれない所（字の情報がない）"


def _grade(blob: bytes):
    from todoku.grade import grade_blob
    return grade_blob(blob, "a.pdf")


def test_few_drawn_shapes_are_flagged_with_position():
    # 形の数で決めない：9 個でも 1 個でも、字の情報のない字の大きさの塗りの形は、位置を示して知らせる（B）
    for n in (9, 1):
        shapes = "\n".join(_glyph_shape(60 + 10 * k, 780) for k in range(n))
        blob = glyphs_pdf([(ch, 60 + 12 * k, 600) for k, ch in enumerate("委任状")], extra=shapes)
        f = flags(blob)
        assert DRAWN in f and f[DRAWN]["grade"] == "B", n
        assert "view_page" in f[DRAWN]["detail"]
        assert len(f[DRAWN]["where"]) == 1 and "左 60" in f[DRAWN]["where"][0]
        assert _grade(blob).grade != "A"


def test_drawn_shapes_with_a_real_char_inside_are_still_flagged():
    shapes = "\n".join(_glyph_shape(60 + 10 * k, 780) for k in range(12))
    blob = glyphs_pdf([(ch, 60 + 12 * k, 600) for k, ch in enumerate("委任状")] + [("注", 100, 782)], extra=shapes)
    f = flags(blob)
    assert DRAWN in f and "中に字もある" in f[DRAWN]["where"][0]


def test_page_of_only_drawn_shapes_is_E():
    shapes = "\n".join(_glyph_shape(60 + 10 * k, 780) for k in range(5))
    blob = glyphs_pdf([], extra=shapes)
    assert flags(blob)[DRAWN]["grade"] == "E"


def test_nearby_drawn_clusters_merge_into_one_region_far_ones_stay_apart():
    near = [_glyph_shape(60 + 10 * k, 780) for k in range(3)] + [_glyph_shape(110 + 10 * k, 780) for k in range(3)]
    far = [_glyph_shape(60 + 10 * k, 300) for k in range(3)]
    blob = glyphs_pdf([(ch, 60 + 12 * k, 600) for k, ch in enumerate("委任状")], extra="\n".join(near + far))
    assert len(flags(blob)[DRAWN]["where"]) == 2


def test_page_without_filled_curves_has_no_drawn_flag():
    blob = glyphs_pdf([(ch, 60 + 12 * k, 600) for k, ch in enumerate("委任状")])
    assert DRAWN not in flags(blob)


def test_annotation_appearance_text_is_shown():
    tu = pdfmini.stream(pdfmini.tounicode({1: "要", 2: "注", 3: "意"}))
    ap = pdfmini.stream("BT /F1 12 Tf 2 2 Td <000100020003> Tj ET", "/Type /XObject /Subtype /Form /BBox [0 0 60 20] "
                        "/Resources << /Font << /F1 2 0 R >> >>")
    objs = [pdfmini.type0("TestFont", 3), tu, ap, "<< /Type /Annot /Subtype /Stamp /Rect [100 700 160 720] /AP << /N 4 0 R >> >>"]
    blob = pdfmini.build("", "<< /Font << /F1 2 0 R >> >>", objs, annots="/Annots [5 0 R]")
    ann = window(blob)["pages"][0]["annotations"]
    assert ann and ann[0].get("drawn") == "要注意"
