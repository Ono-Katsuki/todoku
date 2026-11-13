"""ページの画像の、模様の塗り（Tiling Pattern）：細かい模様・こまの多い模様を、こまの絵から、面積で重みを付けた平均の色で塗る。

前は、画面で 1.5 画素より小さいこまの模様と、こまが 4,000 個を超える塗りを、こまの中身を見ずに、半分の濃さの色（色のある模様は
灰色）で塗って「細かい模様」と伝えていた：オレンジ色の点線の罫線が灰色になり、濃い網も薄い網も同じ濃さになった。
いまは、こまを 1 つだけ別の絵に描き、画素ごとに、その画素にかかる所の平均の色と濃さで塗る。

画像は dpi=72 で描く：1 画素＝1 pt。ページの座標 (x, y) は、画像の (x, 842 − y)。試験は、ほかの道具（poppler など）に頼らない。
"""

from __future__ import annotations

import pytest

import pdfmini
from todoku.pdf.doc import PdfDocument
from todoku.pdf.pageimage import render
from todoku.pdf.render import paint

NOTE = "細かい模様"
CHECK = [[0 if (x + y) % 2 == 0 else 255 for x in range(8)] for y in range(8)]
DOTS = [[0 if (x % 4 == 0 and y % 4 == 0) else 255 for x in range(8)] for y in range(8)]
STRIPE = [[0 if x < 2 else 255 for x in range(8)] for _ in range(8)]
RES = "<< /Pattern << /P 2 0 R >> /ColorSpace << /PC [/Pattern /DeviceRGB] >> /ExtGState << /H << /ca 0.5 >> /M << /BM /Multiply >> >> >>"
FILL = "/Pattern cs /P scn 100 500 300 200 re f"
BLUE = "0 0 1 rg 50 450 200 300 re f "
ORANGE = (0xC6, 0x59, 0x11)


def _gray(rows) -> bytes:
    return pdfmini.stream(bytes(v for r in rows for v in r), "/Type /XObject /Subtype /Image /Width 8 /Height 8 /ColorSpace /DeviceGray /BitsPerComponent 8")


def _rgb(c) -> bytes:
    return pdfmini.stream(bytes(c) * 64, "/Type /XObject /Subtype /Image /Width 8 /Height 8 /ColorSpace /DeviceRGB /BitsPerComponent 8")


def _pat(content: str, bbox="0 0 1 1", step="1 1", matrix="0.75 0 0 0.75 0 0", ptype=1, res="<< /XObject << /I 3 0 R >> >>") -> bytes:
    xs, ys = step.split()
    return pdfmini.stream(content, f"/Type /Pattern /PatternType 1 /PaintType {ptype} /TilingType 1 /BBox [{bbox}] /XStep {xs} /YStep {ys} "
                                   f"/Matrix [{matrix}] /Resources {res}")


def _render(content: str, objs: list, dpi: float = 72, rotate: int = 0) -> dict:
    blob = pdfmini.build(content, RES, objs)
    if rotate:
        blob = blob.replace(b"/MediaBox [0 0 595 842]", b"/MediaBox [0 0 595 842] /Rotate %d" % rotate)
    return render(PdfDocument(blob, "t.pdf"), 1, dpi=dpi)


def _mean(r: dict, x0: int, y0: int, x1: int, y1: int) -> tuple:
    """画像の画素の四角（左上 (x0, y0)・右下 (x1, y1) の手前まで）の平均の色。"""
    w, rgb = r["w"], r["rgb"]
    s = [0, 0, 0]
    for y in range(y0, y1):
        row = bytes(rgb[(y * w + x0) * 3:(y * w + x1) * 3])
        for c in range(3):
            s[c] += sum(row[c::3])
    n = (x1 - x0) * (y1 - y0)
    return tuple(v / n for v in s)


def _px(r: dict, x: int, y: int) -> tuple:
    i = (y * r["w"] + x) * 3
    return tuple(r["rgb"][i:i + 3])


def _near(got, want, tol: float = 1.0) -> bool:
    return all(abs(g - w) <= tol for g, w in zip(got, want))


# ページの (100, 500)〜(400, 700) の塗りの、中ほどの画素の四角（72 dpi）
MID = (150, 180, 350, 300)
# 青い下地（x 50〜250）にかかる所と、白い紙の所
ON_BLUE, ON_WHITE = (120, 180, 240, 300), (270, 180, 380, 300)


def test_a_fine_colored_pattern_is_painted_with_the_color_of_its_tile():
    """こまの画像が 1 色（Office の点線の罫線の色）：その色で塗る。前は、灰色（#BFBFBF）だった。"""
    r = _render(FILL, [_pat("/I Do"), _rgb(ORANGE)])
    assert _px(r, 250, 240) == ORANGE and _px(r, 101, 143) == ORANGE and _px(r, 398, 340) == ORANGE
    assert _px(r, 98, 240) == (255, 255, 255) and _px(r, 402, 240) == (255, 255, 255)  # 塗りの外
    assert NOTE not in (r.get("approx") or {})


@pytest.mark.parametrize("rows, want", [(CHECK, 127.5), (DOTS, 255 * 60 / 64), (STRIPE, 255 * 0.75)])
def test_a_fine_pattern_is_painted_with_the_density_of_its_tile(rows, want):
    """網の濃さは、こまの中の黒い所の割合：市松は半分、点は 64 分の 4、縞は 4 分の 1。前は、どれも半分の灰色（191）だった。"""
    r = _render(FILL, [_pat("/I Do"), _gray(rows)])
    assert _near(_mean(r, *MID), (want,) * 3)
    assert NOTE not in (r.get("approx") or {})


def test_the_gaps_of_a_fine_pattern_show_what_is_under_it():
    """こまの 4 分の 1 だけを緑で塗る模様：残りの 4 分の 3 は、下の色（青・白）が見える。"""
    r = _render(BLUE + FILL, [_pat("0 0.5 0 rg 0 0 0.5 0.5 re f", res="<< >>")])
    assert _near(_mean(r, *ON_BLUE), (0, 127.5 / 4, 255 * 0.75))
    assert _near(_mean(r, *ON_WHITE), (255 * 0.75, 255 * 0.75 + 127.5 / 4, 255 * 0.75))


def test_a_fine_uncolored_pattern_takes_the_color_of_the_fill():
    """色のない模様（PaintType 2）：塗りの色（赤）で、こまの半分。"""
    r = _render(BLUE + "/PC cs 1 0 0 /P scn 100 500 300 200 re f", [_pat("0 0 0.5 1 re f", ptype=2, res="<< >>")])
    assert _near(_mean(r, *ON_BLUE), (127.5, 0, 127.5))
    assert _near(_mean(r, *ON_WHITE), (255, 127.5, 127.5))


def test_a_fine_pattern_with_opacity_and_blend_mode():
    r = _render(BLUE + "/H gs " + FILL, [_pat("/I Do"), _gray(CHECK)])  # 半分の透け方
    assert _near(_mean(r, *ON_BLUE), (63.75, 63.75, 191.25)) and _near(_mean(r, *ON_WHITE), (191.25,) * 3)
    r = _render(BLUE + "/M gs " + FILL, [_pat("/I Do"), _rgb(ORANGE)])  # 乗算
    assert _px(r, 200, 240) == (0, 0, 0x11) and _px(r, 300, 240) == ORANGE
    r = _render(BLUE + "/M gs " + FILL, [_pat("/I Do"), _gray(DOTS)])
    assert _near(_mean(r, *ON_BLUE), (0, 0, 255 * 60 / 64)) and _near(_mean(r, *ON_WHITE), (255 * 60 / 64,) * 3)


def test_a_fine_pattern_is_clipped():
    r = _render("200 550 100 100 re W n " + FILL, [_pat("/I Do"), _rgb(ORANGE)])
    assert _px(r, 250, 240) == ORANGE and _px(r, 150, 240) == (255, 255, 255) and _px(r, 250, 180) == (255, 255, 255)


@pytest.mark.parametrize("matrix", ["0 0.75 -0.75 0 0 0", "-0.75 0 0 0.75 3 1", "0.6 0.35 -0.35 0.6 3 7", "0.5 0 0 0.5 0 0", "0.01 0 0 0.01 0 0"])
def test_a_fine_pattern_turned_mirrored_slanted_or_very_small(matrix):
    """模様の向き・大きさが変わっても、濃さは同じ（縞は 4 分の 1）。"""
    r = _render(FILL, [_pat("/I Do", matrix=matrix), _gray(STRIPE)])
    assert _near(_mean(r, *MID), (191.25,) * 3, 1.5)
    assert NOTE not in (r.get("approx") or {})


@pytest.mark.parametrize("rotate", [90, 180, 270])
def test_a_fine_pattern_on_a_rotated_page(rotate):
    r = _render("/Pattern cs /P scn 0 0 595 842 re f", [_pat("/I Do"), _gray(STRIPE)], rotate=rotate)
    assert _near(_mean(r, 100, 100, 400, 400), (191.25,) * 3)


MANY = "/Pattern cs /P scn 0 0 595 842 re f"
SQUARES = "0 0 2 2 re f 1 0 0 rg 2 2 2 2 re f"  # こまの 4 分の 1 が黒、4 分の 1 が赤


def test_a_fill_of_many_tiles_is_the_same_picture_as_drawing_every_tile(monkeypatch):
    """こまが 4,000 個を超える塗り（こまは 4 画素）：こまの絵から塗っても、こまを 1 つずつ描いた絵と 1 画素も違わない。"""
    objs = [_pat(SQUARES, bbox="0 0 4 4", step="4 4", matrix="1 0 0 1 0 0", res="<< >>")]
    one = _render(MANY, objs)
    assert NOTE not in (one.get("approx") or {})
    assert [_px(one, x, 0) for x in range(4)] == [(0, 0, 0)] * 2 + [(255, 255, 255)] * 2  # ページの上の端は、こまの下の半分から
    assert [_px(one, x, 2) for x in range(4)] == [(255, 255, 255)] * 2 + [(255, 0, 0)] * 2
    monkeypatch.setattr(paint, "TILE_REPLAYS", 10 ** 9)
    assert bytes(_render(MANY, objs)["rgb"]) == bytes(one["rgb"])


@pytest.mark.parametrize("dpi", [50, 100, 130])
def test_the_density_of_a_fill_of_many_tiles_does_not_depend_on_the_resolution(dpi):
    """こまの辺が画素の整数倍でなくても（50 dpi では 2.78 画素）、色の割合は同じ。"""
    r = _render(MANY, [_pat(SQUARES, bbox="0 0 4 4", step="4 4", matrix="1 0 0 1 0 0", res="<< >>")], dpi=dpi)
    assert _near(_mean(r, 20, 20, r["w"] - 20, r["h"] - 20), (191.25, 127.5, 127.5))


SHAPES = "0 0 1 rg 2 3 m 30 9 l 12 28 l f 0.9 0.4 0 RG 3 w 0 20 m 32 4 l S 0 0.6 0 rg 20 18 9 11 re f"


def _far(a: bytes, b: bytes) -> int:
    """2 つの絵の、いちばん離れた色の値の差。"""
    return max(abs(x - y) for x, y in zip(a, b))


@pytest.mark.parametrize("matrix, far", [("1 0 0 1 0 0", 0), ("1 0 0 1 5 -7", 0), ("-1 0 0 1 0 0", 0), ("1 0 0 -1 3 2", 0), ("1.5 0 0 1.5 0 0", 0),
                                         ("2 0 0 1 0 0", 1), ("0 1 -1 0 0 0", 8), ("0 -1 1 0 11 6", 8)])
def test_a_tile_of_16_pixels_or_more_is_drawn_as_on_the_page(monkeypatch, matrix, far):
    """画面で 16 画素以上のこま（辺が画素の整数倍）：こまの絵から塗った絵は、こまを 1 つずつ描いた絵と同じ（三角・線・四角）。
    横に伸ばしたこまは、縁の画素の丸めが 1 だけ、90 度回したこまは、縁のなめらかさ（こまの絵は回す前の向きで描く）が 8 まで違う。"""
    objs = [_pat(SHAPES, bbox="0 0 32 32", step="32 32", matrix=matrix, res="<< >>")]
    each = bytes(_render(FILL, objs)["rgb"])
    assert each.count(b"\x00\x00\xff") > 5000  # 青い三角が描けている
    monkeypatch.setattr(paint, "TILE_REPLAYS", 9)  # こまを 1 つずつ描くのは 9 個まで：この塗りは、こまの絵から塗る
    r = _render(FILL, objs)
    assert NOTE not in (r.get("approx") or {})
    assert _far(bytes(r["rgb"]), each) <= far


def test_tiles_that_overlap_their_neighbours(monkeypatch):
    """BBox が歩みより大きい模様：隣のこまからはみ出した所も、こまの絵に入る。"""
    objs = [_pat("0 0 1 rg 0 0 48 16 re f 1 0 0 rg 0 0 16 48 re f", bbox="0 0 48 48", step="32 32", matrix="1 0 0 1 0 0", res="<< >>")]
    each = bytes(_render(FILL, objs)["rgb"])
    monkeypatch.setattr(paint, "TILE_REPLAYS", 9)  # こまを 1 つずつ描くのは 9 個まで：この塗りは、こまの絵から塗る
    assert bytes(_render(FILL, objs)["rgb"]) == each


def test_the_tile_picture_is_drawn_once_for_the_fills_of_a_page(monkeypatch):
    calls = []
    real = paint._Cell.__init__
    monkeypatch.setattr(paint._Cell, "__init__", lambda self, cv: (calls.append((cv.w, cv.h)), real(self, cv))[1])
    r = _render(FILL + " 100 100 300 100 re f 420 100 50 600 re f", [_pat("/I Do"), _gray(CHECK)])
    assert calls == [(16, 16)]
    assert _near(_mean(r, 150, 760 - 100, 350, 740), (127.5,) * 3)


def test_tiles_that_overlap_too_many_neighbours_are_still_told():
    """こまどうしの重なりが 4,000 個を超える模様（歩みの 100 倍の BBox）は、こまの絵を描かない：前と同じに、半分の濃さで塗って伝える。"""
    r = _render(FILL, [_pat("0 0 100 100 re f", bbox="0 0 100 100", step="1 1", res="<< >>")])
    assert (r.get("approx") or {}).get(NOTE) == 1
    assert _px(r, 250, 240) == (191, 191, 191)


def test_a_pattern_that_paints_with_itself_ends():
    objs = [_pat("/Pattern cs /P scn 0 0 1 1 re f", res="<< /Pattern << /P 2 0 R >> >>")]
    r = _render(FILL, objs)
    assert _px(r, 250, 240) == (255, 255, 255)
