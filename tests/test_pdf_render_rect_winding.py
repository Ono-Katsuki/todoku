"""ページの画像：四角だけの道を巻き数の決まりで塗るとき、向きの違う四角の重なりは塗らない（内と外の四角で描いた枠）。

前は、四角だけの並びを 1 つずつ塗っていた（速い道）：外の四角と、逆向きの内の四角で描いた枠は、枠の中の全部が塗りつぶされ、
view_page の画像でも、書いた字が見えるかの確かめ（inkcheck）でも、枠の中が 1 色の面になっていた。
いまは、向きの違う四角が重なる並びは、巻き数を数えて塗る。同じ向きの四角・重ならない四角・線の形は、前と同じ。
"""

from __future__ import annotations

import pdfmini
from test_pdf_render import BLACK, WHITE, _render, dark, px
from todoku.pdf.render import canvas

# 外の四角（幅も高さも負）と、逆向きの内の四角（幅だけ負）。枠の太さは 10 pt
FRAME = "0 g 400 700 -300 -500 re 390 210 -280 480 re f"


def test_a_frame_drawn_as_two_rectangles_of_opposite_direction_is_hollow():
    r = _render(pdfmini.build(FRAME, "<< >>", []))
    assert px(r, 250, 450) == WHITE and px(r, 115, 215) == WHITE and px(r, 385, 685) == WHITE  # 枠の中
    for x, y in ((105, 450), (395, 450), (250, 205), (250, 695)):  # 枠の 4 辺
        assert px(r, x, y) == BLACK
    assert px(r, 50, 450) == WHITE  # 枠の外


def test_the_frame_is_hollow_in_any_order_and_with_a_transform():
    for c in ("0 g 390 210 -280 480 re 400 700 -300 -500 re f",       # 内が先
              "0 g 100 200 300 500 re 110 690 280 -480 re f",        # 外が逆の向き、内はその逆
              "q 1 0 0 -1 0 900 cm 0 g 100 200 300 500 re 110 690 280 -480 re f Q",  # 上下を返す行列（向きは、どれも一緒に返る）
              "0 g 100 200 300 500 re 110 690 280 -480 re B"):       # 塗って線を引く
        r = _render(pdfmini.build(c, "<< >>", []))
        assert px(r, 250, 450) == WHITE, c
        assert px(r, 105, 450) == BLACK and px(r, 395, 450) == BLACK, c


def test_rectangles_of_the_same_direction_are_all_filled():
    # 同じ向き：重なりの巻き数は 2（塗る）。前と同じ
    r = _render(pdfmini.build("0 g 100 200 300 500 re 110 210 280 480 re f", "<< >>", []))
    assert px(r, 250, 450) == BLACK and px(r, 105, 450) == BLACK
    # 向きが違っても、重ならない四角は、どちらも塗る
    r = _render(pdfmini.build("0 g 100 600 100 100 re 400 700 -100 -100 re 250 500 -50 50 re f", "<< >>", []))
    assert px(r, 150, 650) == BLACK and px(r, 350, 650) == BLACK and px(r, 225, 525) == BLACK and px(r, 250, 650) == WHITE


def test_two_holes_and_a_filled_island():
    # 外の四角の中に、逆向きの穴が 2 つ。穴の中に、外と同じ向きの四角（島）
    c = "0 g 100 200 400 400 re 120 580 150 -360 re 330 580 150 -360 re 150 300 60 60 re f"
    r = _render(pdfmini.build(c, "<< >>", []))
    assert px(r, 110, 400) == BLACK and px(r, 300, 400) == BLACK and px(r, 490, 400) == BLACK
    assert px(r, 200, 500) == WHITE and px(r, 400, 500) == WHITE
    assert px(r, 180, 330) == BLACK  # 島


def test_the_even_odd_rule_and_strokes_are_unchanged():
    r = _render(pdfmini.build("0 g 100 200 300 500 re 110 210 280 480 re f*", "<< >>", []))
    assert px(r, 250, 450) == WHITE and px(r, 105, 450) == BLACK  # 偶奇：向きによらず穴
    # 線の形の四角は、向きを見ずに、どれも塗る（角の重なりも）
    r = _render(pdfmini.build("0 G 20 w 100 200 300 500 re S 400 100 m 100 100 l S", "<< >>", []))
    for x, y in ((100, 200), (400, 700), (105, 695), (250, 205), (250, 100)):
        assert dark(px(r, x, y)), (x, y)
    assert px(r, 250, 450) == WHITE


def test_a_thin_rectangle_keeps_its_direction():
    # 画素に合わせる細い四角（2 画素未満）も、向きは元のまま：細い穴は、穴のまま
    r = _render(pdfmini.build("0 g 100 200 300 500 re 250.75 210 -1.5 480 re f", "<< >>", []))
    assert px(r, 249, 450) == px(r, 250, 450) == WHITE and px(r, 200, 450) == BLACK and px(r, 300, 450) == BLACK
    r = _render(pdfmini.build("0 g 100 200 300 500 re 249.25 210 1.5 480 re f", "<< >>", []))
    assert px(r, 249, 450) == px(r, 250, 450) == BLACK  # 同じ向きの細い四角は、穴にならない


def test_many_rectangles_of_mixed_direction_are_counted_by_winding(monkeypatch):
    # 向きの違う四角の組が多い並びは、1 組ずつ調べずに、巻き数を数える道に任せる：結果は同じ
    cells = " ".join(f"{100 + 4 * i} 300 3 200 re" for i in range(70)) + " " + " ".join(f"{100 + 4 * i} 520 3 -40 re" for i in range(70))
    r = _render(pdfmini.build("0 g " + cells + " f", "<< >>", []))
    assert px(r, 101, 400) == BLACK and px(r, 101, 490) == WHITE and px(r, 101, 510) == BLACK and px(r, 103, 400) == WHITE and px(r, 101, 310) == BLACK
    seen = []
    real = canvas.scan
    monkeypatch.setattr(canvas, "scan", lambda *a, **k: seen.append(1) or real(*a, **k))
    _render(pdfmini.build("0 g " + cells + " f", "<< >>", []))
    assert seen  # 70 × 70 組：巻き数で数えた
    del seen[:]
    _render(pdfmini.build("0 g " + " ".join(f"{100 + 4 * i} 300 3 200 re" for i in range(70)) + " f", "<< >>", []))
    assert not seen  # 同じ向きだけ：速い道のまま
