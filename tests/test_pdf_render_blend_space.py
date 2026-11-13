"""ページの画像：RGB でない色空間で重ねる所の重ね方を、伝える（ISO 32000-1 11.3.4・11.6.6）。

仕様は、Normal でない重ね方を、ページ・下と切り離すまとまりの色空間（Group の CS：CMYK など）で計算する。この描画は、いつも RGB で
重ねるので、CMYK で重ねる所の色は、ほかのビューアーの絵と違う。前は、何も伝えなかった。いまは、その重ね方を使った数を
「RGB でない色空間で重ねる所の重ね方（RGB で重ねた）」と伝える（絵は、前と同じ）。
"""

from __future__ import annotations

import pytest

import pdfmini
from test_pdf_render import _render, px
from test_pdf_render_transparency import B_FILL, FORM, ON_B, S_FILL
from todoku.pdf.render.paint import SPACE_NOTE

EG = "/ExtGState << /BM << /BM /Multiply >> /NM << /BM /Normal >> /H << /ca 0.5 >> >>"
BLEND = f"{B_FILL} q /BM gs {S_FILL} Q"


def _page(content: str, cs: str | None, objs: tuple = (), more: str = "") -> dict:
    group = f"/Group << /S /Transparency /CS /{cs} >>" if cs else ""
    return _render(pdfmini.build(content, f"<< {EG} {more} >>", list(objs), annots=group))


def _told(r: dict) -> int:
    return r.get("approx", {}).get(SPACE_NOTE, 0)


def _inner(cs: str | None, iso: bool) -> bytes:
    group = f"/Group << /S /Transparency {f'/CS /{cs}' if cs else ''}{' /I true' if iso else ''} >>"
    return pdfmini.stream(BLEND, FORM + group + f" /Resources << {EG} >>")


@pytest.mark.parametrize("cs, told", [("DeviceCMYK", 1), ("DeviceGray", 1), ("DeviceRGB", 0), (None, 0)])
def test_a_blend_mode_on_a_page_whose_group_is_not_rgb_is_told(cs, told):
    r = _page(BLEND, cs)
    assert _told(r) == told and list(r.get("approx", {})) == ([SPACE_NOTE] if told else [])
    assert px(r, *ON_B) == px(_page(BLEND, None), *ON_B)  # 絵は、RGB で重ねた色のまま


def test_normal_blending_and_opacity_in_a_cmyk_group_are_not_told():
    """伝えるのは、Normal でない重ね方だけ（透け方だけの重ね合わせは、伝えない）。"""
    assert not _page(f"{B_FILL} q /NM gs /H gs {S_FILL} Q", "DeviceCMYK").get("approx")
    assert not _page(f"{B_FILL} {S_FILL}", "DeviceCMYK").get("approx")


@pytest.mark.parametrize("page_cs, cs, iso, told", [
    (None, "DeviceCMYK", True, 1),     # 下と切り離すまとまり：自分の色空間で重ねる
    (None, "DeviceCMYK", False, 0),    # 切り離さないまとまり：親（ここではページ）の色空間
    ("DeviceCMYK", "DeviceRGB", True, 0),
    ("DeviceCMYK", "DeviceRGB", False, 1),
    ("DeviceCMYK", None, True, 1),     # 色空間のないまとまり：親の色空間
])
def test_a_group_blends_in_its_own_space_only_when_it_is_isolated(page_cs, cs, iso, told):
    r = _page("/X Do", page_cs, (_inner(cs, iso),), "/XObject << /X 2 0 R >>")
    assert _told(r) == told


def test_the_space_of_the_page_comes_back_after_an_isolated_group():
    r = _page(f"/X Do {BLEND}", None, (_inner("DeviceCMYK", True),), "/XObject << /X 2 0 R >>")
    assert _told(r) == 1  # まとまりの中の 1 回だけ。まとまりの後のページの重ね方は、数えない
    r = _page(f"/X Do {BLEND}", "DeviceCMYK", (_inner("DeviceRGB", True),), "/XObject << /X 2 0 R >>")
    assert _told(r) == 1  # ページの 1 回だけ
