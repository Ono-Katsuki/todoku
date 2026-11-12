"""PDF の塗り・画像の下の字：2 つ以上の四角を合わせると全部が隠れる字。

- 字の四角（送りの幅 × 書体の高さ）の全部が、その字より後に描いた、透けない四角（塗り・画像）を合わせた範囲の中にある字は、
  1 つの四角の中にある字と同じに扱う：塗りだけで隠れるなら「塗りの下にあって見えない字」、画像を合わせて初めて隠れるなら
  「画像の下に描いた字」（読み順に残す）
- 四角の間にすき間がある・隅が残る・四角の 1 つが字より前・四角の 1 つが透ける、のどれかなら、見える字のまま
- 短い辺が 2.5 pt 以下の細い塗りは、並べても数えない（罫線として描いた四角）
"""

from __future__ import annotations

import random
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from test_pdf_cover import (DRAW, HIDDEN, IMG, NAME, TITLE, UNDER, XIM, _audit_ok, _content, _page,  # noqa: E402
                            _reads_only_name, _visible, _write)

from todoku.pdf.content import _union_covers  # noqa: E402


def _edge(tmp_path) -> tuple[float, float]:
    """「Hidden」の 1 つ目の「d」の中央の x と、字の四角の上下の中央の y。"""
    d = [g for g in _content(_write(tmp_path, HIDDEN)).glyphs if g.text == "d"][0]
    return (d.x0 + d.x1) / 2, (d.y0 + d.y1) / 2


def _halves(x: float, gap: float = 0.0) -> bytes:
    """140〜240 × 700〜750 の四角を、x で左右に分けた 2 つの塗り（gap は、間のすき間）。"""
    return b"140 700 %.3f 50 re f\n%.3f 700 %.3f 50 re f\n" % (x - 140, x + gap, 240 - x - gap)


def test_two_fills_side_by_side_hide_the_glyph_on_their_border(tmp_path):
    x, _y = _edge(tmp_path)
    p = _write(tmp_path, NAME + HIDDEN + _halves(x))
    pg, md = _page(p)
    assert _reads_only_name(pg) and [c["text"] for c in pg["covered"]] == ["Hidden"]
    assert TITLE + "（6 字）" in md
    assert _audit_ok(p)


def test_fills_that_overlap_hide_the_glyph_under_both(tmp_path):
    x, _y = _edge(tmp_path)
    p = _write(tmp_path, NAME + HIDDEN + b"140 700 %.3f 50 re f\n%.3f 700 %.3f 50 re f\n" % (x - 138, x - 2, 242 - x))
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]


def test_a_gap_between_the_fills_leaves_the_glyph_visible(tmp_path):
    x, _y = _edge(tmp_path)
    p = _write(tmp_path, NAME + HIDDEN + _halves(x, gap=0.5))
    pg, _md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["Hi den"]  # すき間の通る「d」だけが見える字（同じ行の字は 1 行で示す）
    assert pg["chars"] == len("Name:") + 1
    assert _audit_ok(p)


def test_four_fills_meeting_inside_the_glyph_hide_it(tmp_path):
    x, y = _edge(tmp_path)
    four = (b"140 700 %.3f %.3f re f\n" % (x - 140, y - 700) + b"%.3f 700 %.3f %.3f re f\n" % (x, 240 - x, y - 700)
            + b"140 %.3f %.3f %.3f re f\n" % (y, x - 140, 750 - y) + b"%.3f %.3f %.3f %.3f re f\n" % (x, y, 240 - x, 750 - y))
    p = _write(tmp_path, NAME + HIDDEN + four)
    pg, _md = _page(p)
    assert _reads_only_name(pg) and [c["text"] for c in pg["covered"]] == ["Hidden"]


def test_a_corner_left_open_leaves_the_glyph_visible(tmp_path):
    x, y = _edge(tmp_path)
    three = (b"140 700 %.3f %.3f re f\n" % (x - 140, y - 700) + b"%.3f 700 %.3f %.3f re f\n" % (x, 240 - x, y - 700)
             + b"140 %.3f %.3f %.3f re f\n" % (y, x - 140, 750 - y))  # 右上が空く：「d」の右上と「den」の上半分が見える
    p = _write(tmp_path, NAME + HIDDEN + three)
    pg, _md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["Hi"]
    assert pg["chars"] == len("Name:") + len("dden")


def test_a_fill_drawn_before_the_text_is_not_part_of_the_cover(tmp_path):
    x, _y = _edge(tmp_path)
    left, right = _halves(x).split(b"\n")[:2]
    p = _write(tmp_path, NAME + left + b"\n" + HIDDEN + right + b"\n")
    pg, _md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["den"]
    assert pg["chars"] == len("Name:") + len("Hid")


def test_a_fill_that_may_show_what_is_under_it_is_not_part_of_the_cover(tmp_path):
    x, _y = _edge(tmp_path)
    left, right = _halves(x).split(b"\n")[:2]
    p = _write(tmp_path, NAME + HIDDEN + left + b"\nq /G1 gs " + right + b" Q\n", b"/ExtGState << /G1 << /ca 0.5 >> >>")
    pg, _md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["Hi"]
    assert pg["chars"] == len("Name:") + len("dden")


def test_rectangles_of_one_path_that_touch_hide_the_glyph_on_their_border(tmp_path):
    x, _y = _edge(tmp_path)
    p = _write(tmp_path, NAME + HIDDEN + b"140 700 %.3f 50 re %.3f 700 %.3f 50 re f\n" % (x - 140, x, 240 - x))
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]


def test_thin_fills_side_by_side_hide_nothing(tmp_path):
    # 幅 2.5 pt の塗りを並べて、字の全部を覆う：細い塗りは、罫線として描いた四角なので数えない（見える字のまま）
    strips = b"".join(b"%.1f 700 2.5 50 re f\n" % (140 + 2.5 * i) for i in range(40))
    _visible(tmp_path, NAME + HIDDEN + strips)


def test_a_fill_and_an_image_together_put_the_glyph_under_the_image(tmp_path):
    # 左は塗り、右は画像：塗りだけの下の字は「塗りの下」、境の字と画像の下の字は「画像の下」（読み順に残す）
    x, _y = _edge(tmp_path)
    body = NAME + HIDDEN + b"140 700 %.3f 50 re f\n" % (x - 140) + b"q %.3f 0 0 50 %.3f 700 cm /Im1 Do Q\n" % (240 - x, x)
    p = _write(tmp_path, body, XIM, IMG)
    pg, _md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["Hi"]
    pc = _content(p)
    assert [g.text for g in pc.glyphs if g.under] == list("dden") and all(g.invisible for g in pc.glyphs if g.under)
    (f,) = [f for f in pg["read_flags"] if f["title"] == UNDER]
    assert "".join(w.split("「")[1].split("」")[0] for w in f["where"]) == "dden"
    assert _audit_ok(p)


def test_two_images_side_by_side_put_the_glyph_on_their_border_under_the_image(tmp_path):
    x, _y = _edge(tmp_path)
    body = (NAME + HIDDEN + b"q %.3f 0 0 50 140 700 cm /Im1 Do Q\n" % (x - 140)
            + b"q %.3f 0 0 50 %.3f 700 cm /Im1 Do Q\n" % (240 - x, x))
    p = _write(tmp_path, body, XIM, IMG)
    pg, md = _page(p)
    assert "covered" not in pg and TITLE not in md
    assert [g.text for g in _content(p).glyphs if g.under] == list("Hidden")
    assert any("Hidden" in ln for ln in pg["reading"])


def test_an_image_over_all_and_two_fills_over_all_is_told_as_under_the_fills(tmp_path):
    x, _y = _edge(tmp_path)
    p = _write(tmp_path, NAME + HIDDEN + DRAW + _halves(x), XIM, IMG)
    pg, _md = _page(p)
    assert _reads_only_name(pg) and [c["text"] for c in pg["covered"]] == ["Hidden"]
    assert not any(f["title"] == UNDER for f in pg.get("read_flags", []))


def test_a_glyph_larger_than_the_search_cell_follows_the_same_rule(tmp_path):
    big = b"BT /F1 96 Tf 150 620 Td (W) Tj ET\n"
    g = _content(_write(tmp_path, big)).glyphs[0]
    assert g.x1 - g.x0 > 36 and g.y1 - g.y0 > 36
    x = (g.x0 + g.x1) / 2
    two = b"120 560 %.3f 240 re f\n%.3f 560 %.3f 240 re f\n" % (x - 120, x, 400 - x)
    p = _write(tmp_path, NAME + big + two)
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["W"]
    gap = b"120 560 %.3f 240 re f\n%.3f 560 %.3f 240 re f\n" % (x - 120, x + 1, 399 - x)
    pg, _md = _page(_write(tmp_path, NAME + big + gap))
    assert "covered" not in pg and any("W" in ln for ln in pg["reading"])
    one = b"120 560 280 240 re f\n"
    assert [c["text"] for c in _page(_write(tmp_path, NAME + big + one))[0]["covered"]] == ["W"]


def test_union_of_rectangles_against_counting_unit_squares():
    box = (0, 0, 10, 10)
    assert _union_covers(box, [[0, 0, 5, 10], [5, 0, 10, 10]])
    assert _union_covers(box, [[0, 0, 5, 10], [5.015, 0, 10, 10]])  # ずれは 0.01 ずつまで
    assert not _union_covers(box, [[0, 0, 5, 10], [5.03, 0, 10, 10]])
    assert not _union_covers((0, 0, 0, 10), [[-1, 0, 1, 5], [-1, 5, 1, 10]]) and not _union_covers(box, [])
    rnd = random.Random(7)
    seen = {True: 0, False: 0}
    for _ in range(3000):
        rs = []
        for _n in range(rnd.randint(1, 6)):
            x, y = rnd.randint(-2, 2), rnd.randint(-2, 2)
            rs.append([x, y, x + rnd.randint(1, 6), y + rnd.randint(1, 6)])
        whole = all(any(r[0] <= x and x + 1 <= r[2] and r[1] <= y and y + 1 <= r[3] for r in rs)
                    for x in range(4) for y in range(4))
        assert _union_covers((0, 0, 4, 4), rs, 0) == whole, rs
        seen[whole] += 1
    assert min(seen.values()) > 100
