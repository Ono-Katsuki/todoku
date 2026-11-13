"""ページの画像：透けて描くまとまり（透明グループのフォームの部品）を半透明で置いた所は、中のものも、その透け方で描く。

前は、まとまりの中で透け方を 1 に戻す書類（まとまりの中の透け方は 1 から始まるので、そう書く道具が多い）は、まとまりの全体が
透けずに描かれ、view_page の画像でも、書いた字と下地の明るさの比べ（inkcheck）でも、下地が実際より濃い色になっていた。
"""

from __future__ import annotations

import pdfmini
from test_pdf_render import BLACK, WHITE, _render, px

GROUP = "/Group << /Type /Group /S /Transparency /CS /DeviceGray >>"
RES = "<< /ExtGState << /H 3 0 R /F 4 0 R /Q 5 0 R >> /XObject << /Fm 2 0 R >> >>"
STATES = ["<< /Type /ExtGState /ca 0.5 /CA 0.5 >>", "<< /Type /ExtGState /ca 1 /CA 1 >>", "<< /Type /ExtGState /ca 0.5 /CA 0.5 >>"]


def _page(form: str, extra: str = "", draw: str = "q /H gs /Fm Do Q") -> dict:
    fm = pdfmini.stream(form, "/Type /XObject /Subtype /Form /BBox [0 0 595 842] " + extra
                        + " /Resources << /ExtGState << /F 4 0 R /Q 5 0 R >> >>")
    return _render(pdfmini.build(draw + " 0 g 400 100 50 50 re f", RES, [fm] + STATES))


def _grey(c, v: int) -> bool:
    return all(abs(x - v) <= 2 for x in c)


def test_a_group_placed_half_transparent_is_drawn_half_transparent():
    # まとまりの中で、透け方を 1 に戻してから描く
    r = _page("/F gs 0 g 100 600 100 100 re f 0 G 10 w 100 500 m 200 500 l S", GROUP)
    assert _grey(px(r, 150, 650), 128) and _grey(px(r, 150, 500), 128)
    assert px(r, 425, 125) == BLACK  # まとまりの外は、元の透け方
    # 透け方に触れないまとまりも、同じ（前と同じ）
    r = _page("0 g 100 600 100 100 re f", GROUP)
    assert _grey(px(r, 150, 650), 128)


def test_alpha_inside_the_group_multiplies():
    r = _page("/Q gs 0 g 100 600 100 100 re f", GROUP)  # 中の 0.5 × まとまりの 0.5
    assert _grey(px(r, 150, 650), 191)
    r = _page("/Q gs 0 g 100 600 100 100 re f", GROUP, draw="/Fm Do")  # 透けずに置いたまとまり：中の 0.5 だけ
    assert _grey(px(r, 150, 650), 128)


def test_a_group_inside_a_group():
    inner = pdfmini.stream("/F gs 0 g 100 600 100 100 re f", "/Type /XObject /Subtype /Form /BBox [0 0 595 842] " + GROUP
                           + " /Resources << /ExtGState << /F 4 0 R >> >>")
    outer = pdfmini.stream("/F gs /Q gs /In Do /F gs 0 g 300 600 100 100 re f", "/Type /XObject /Subtype /Form /BBox [0 0 595 842] " + GROUP
                           + " /Resources << /ExtGState << /F 4 0 R /Q 5 0 R >> /XObject << /In 6 0 R >> >>")
    r = _render(pdfmini.build("q /H gs /Fm Do Q", RES, [outer] + STATES + [inner]))
    assert _grey(px(r, 150, 650), 191)  # 外 0.5 × 内を置いた 0.5
    assert _grey(px(r, 350, 650), 128)  # 外のまとまりの中の、透け方 1 のもの


def test_a_form_that_is_not_a_group_is_unchanged():
    # まとまりでないフォームの部品は、ページにじかに描くのと同じ：中で透け方を 1 にすれば、透けない
    r = _page("/F gs 0 g 100 600 100 100 re f")
    assert px(r, 150, 650) == BLACK
    r = _page("0 g 100 600 100 100 re f")
    assert _grey(px(r, 150, 650), 128)
    assert px(r, 300, 300) == WHITE
