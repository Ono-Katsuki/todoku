"""ページの画像：透けのマスク（SMask）・重ね方（BM）・透けて描くまとまりを、PDF の決まり（11.3〜11.6）のとおりに描く。

前は、透けのマスクのある所は、マスクを掛けずに描き、重ね方は、ふつうの重ね方で描くか、下に何もない所でも紙の白と重ねていた
（どちらも「近い形で描いたか、飛ばした」と伝えていた）。透けて描くまとまりを半透明で置いた所は、中のもの 1 つ 1 つを半透明で
描いていたので、まとまりの中で重なるものが、互いに透けていた。
いまは、マスクのまとまりを描いて、その明るさ・不透明度を、塗り・画像・字・置いたまとまりに掛ける。重ね方は、下に描いたものの
濃さ（何も描いていない紙は、透ける下地）と合わせて、16 個のどれも式のとおりに重ねる。透け方・重ね方・マスクで置いたまとまりは、
別の絵に描いてから、まとまりの全体を重ねる。

期待する色は、この試験の中で、決まりの式から小数で計算する（描画の表は使わない）。画像は dpi=72：1 画素＝1 pt。
"""

from __future__ import annotations

import math

import pytest

import pdfmini
from test_pdf_render import WHITE, _render, px
from todoku.pdf.render import canvas, paint

FORM = "/Type /XObject /Subtype /Form /BBox [0 0 595 842] "
BLUE, RED, GREEN, YELLOW, CYAN = (0, 0, 255), (255, 0, 0), (0, 255, 0), (255, 255, 0), (0, 255, 255)
BLUE_BOX = "0 0 1 rg 100 500 300 200 re f"  # 左半分（x 100〜250）と右半分（x 250〜400）を見る
LEFT, RIGHT = (175, 600), (325, 600)
TWO = "1 0 0 rg 100 500 200 200 re f 0 1 0 rg 200 550 200 200 re f"  # 重なる赤と緑：赤だけ (150, 520)・重なり (250, 600)・緑だけ (350, 700)
ONLY_RED, BOTH, ONLY_GREEN = (150, 520), (250, 600), (350, 700)


def _group(body: str, cs: str = "DeviceRGB", iso: bool = False, res: str = "") -> bytes:
    return pdfmini.stream(body, FORM + f"/Group << /S /Transparency /CS /{cs}{' /I true' if iso else ''} >>"
                          + (f" /Resources << {res} >>" if res else ""))


def _form(body: str, res: str = "") -> bytes:
    return pdfmini.stream(body, FORM + (f" /Resources << {res} >>" if res else ""))


def _near(c, want, tol: int = 2) -> bool:
    return all(abs(a - b) <= tol for a, b in zip(c, want))


def _over(under, over, a: float) -> tuple:
    return tuple(round(u * (1 - a) + o * a) for u, o in zip(under, over))


# ====================================================================== 透けのマスク
def _masked(mask_body: str, mask: str = "/S /Luminosity", content: str = "q /M gs " + BLUE_BOX + " Q", cs: str = "DeviceGray",
            mask_res: str = "", more: str = "", objs: tuple = ()) -> dict:
    g = _group(mask_body, cs, res=mask_res)
    res = f"<< /ExtGState << /M << /SMask << /Type /Mask {mask} /G 2 0 R >> >> /N << /SMask /None >> >> {more} >>"
    r = _render(pdfmini.build(content, res, [g, *objs]))
    assert not r.get("approx"), r.get("approx")  # 「近い形で描いた」とは言わない
    return r


def test_a_luminosity_mask_shows_only_where_the_mask_is_light():
    r = _masked("1 g 100 500 150 200 re f")  # 左半分だけが白。ほかは、マスクの下地の黒
    assert px(r, *LEFT) == BLUE and px(r, *RIGHT) == WHITE
    r = _masked("0.5 g 100 500 150 200 re f")  # 灰色：半分だけ出る
    assert _near(px(r, *LEFT), _over(WHITE, BLUE, 128 / 255)) and px(r, *RIGHT) == WHITE


def test_the_backdrop_colour_of_a_luminosity_mask():
    r = _masked("0 g 100 500 150 200 re f", "/S /Luminosity /BC [1]")  # 下地が白：描いていない所は、全部出る
    assert px(r, *LEFT) == WHITE and px(r, *RIGHT) == BLUE


def test_the_transfer_function_of_a_mask():
    tr = "/S /Luminosity /TR << /FunctionType 2 /Domain [0 1] /C0 [1] /C1 [0] /N 1 >>"  # 明るさを逆にする
    r = _masked("1 g 100 500 150 200 re f", tr)
    assert px(r, *LEFT) == WHITE and px(r, *RIGHT) == BLUE


def test_an_alpha_mask_uses_the_opacity_of_the_group():
    r = _masked("0 g 100 500 150 200 re f", "/S /Alpha")  # 色は見ない：描いた所が出る
    assert px(r, *LEFT) == BLUE and px(r, *RIGHT) == WHITE
    r = _masked("/H gs 0 g 100 500 150 200 re f", "/S /Alpha", mask_res="/ExtGState << /H << /ca 0.5 >> >>")
    assert _near(px(r, *LEFT), _over(WHITE, BLUE, 0.5)) and px(r, *RIGHT) == WHITE


def test_a_gradient_mask_fades():
    sh = ("/Shading << /S1 << /ShadingType 2 /ColorSpace /DeviceGray /Coords [100 0 400 0] "
          "/Function << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [1] /N 1 >> /Extend [true true] >> >>")
    r = _masked("/S1 sh", mask_res=sh)
    got = [px(r, x, 600) for x in (110, 175, 250, 325, 390)]
    assert all(_near(c, _over(WHITE, BLUE, (x - 100) / 300), 6) for c, x in zip(got, (110, 175, 250, 325, 390))), got
    assert [c[0] for c in got] == sorted((c[0] for c in got), reverse=True)  # 右へ行くほど濃い


def test_a_mask_is_taken_off_with_none():
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs 0 0 1 rg 100 500 300 100 re f /N gs 100 600 300 100 re f Q")
    assert px(r, 325, 550) == WHITE and px(r, 325, 650) == BLUE and px(r, 175, 550) == BLUE
    # q … Q の外には残らない
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs Q " + BLUE_BOX)
    assert px(r, *RIGHT) == BLUE


def test_a_mask_stays_where_it_was_set():
    # マスクの位置は、gs を呼んだときの座標で決まる：その後の cm では動かない。gs の前の cm では動く
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs 1 0 0 1 100 0 cm 0 0 1 rg 0 500 300 200 re f Q")
    assert px(r, *LEFT) == BLUE and px(r, *RIGHT) == WHITE
    r = _masked("1 g 100 500 150 200 re f", content="q 0.5 0 0 0.5 100 100 cm /M gs " + BLUE_BOX + " Q")
    assert px(r, 200, 400) == BLUE and px(r, 260, 400) == WHITE  # マスクの白は、ページの x 150〜225


def test_a_mask_and_a_clip_both_apply():
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs 150 520 200 150 re W n " + BLUE_BOX + " Q")
    assert px(r, 175, 600) == BLUE and px(r, 120, 600) == WHITE and px(r, 325, 600) == WHITE and px(r, 175, 690) == WHITE


def test_a_mask_on_an_image_and_on_a_shading():
    im = pdfmini.stream(bytes([0, 0, 255]), "/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8")
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs 300 0 0 200 100 500 cm /Im Do Q", more="/XObject << /Im 3 0 R >>", objs=(im,))
    assert px(r, *LEFT) == BLUE and px(r, *RIGHT) == WHITE
    sh = ("/Shading << /S1 << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 595 0] "
          "/Function << /FunctionType 2 /Domain [0 1] /C0 [0 0 1] /C1 [0 0 1] /N 1 >> /Extend [true true] >> >>")
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs /S1 sh Q", more=sh)
    assert px(r, *LEFT) == BLUE and px(r, *RIGHT) == WHITE and px(r, 50, 600) == WHITE


def test_a_mask_on_a_group_applies_to_the_whole_group():
    """透けて描くまとまりを、マスクのある所に置いた：中で SMask を外しても、まとまりの全体にマスクが掛かる（11.6.6）。
    まとまりでないフォームの部品は、ページにじかに描くのと同じ：中で外せば、掛からない。"""
    x = "/XObject << /X 3 0 R >>"
    inner = "/ExtGState << /N << /SMask /None >> >>"
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs /X Do Q", more=x, objs=(_group("/N gs " + BLUE_BOX, res=inner),))
    assert px(r, *LEFT) == BLUE and px(r, *RIGHT) == WHITE
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs /X Do Q", more=x, objs=(_form("/N gs " + BLUE_BOX, inner),))
    assert px(r, *LEFT) == BLUE and px(r, *RIGHT) == BLUE
    r = _masked("1 g 100 500 150 200 re f", content="q /M gs /X Do Q", more=x, objs=(_form(BLUE_BOX),))
    assert px(r, *LEFT) == BLUE and px(r, *RIGHT) == WHITE


def test_a_mask_whose_group_cannot_be_read_is_told():
    res = "<< /ExtGState << /M << /SMask << /Type /Mask /S /Luminosity >> >> >> >>"  # G がない
    r = _render(pdfmini.build("q /M gs " + BLUE_BOX + " Q", res, []))
    assert any("透けのマスク" in k for k in r.get("approx", {}))


# ====================================================================== 重ね方
B, S = (204, 102, 51), (102, 204, 153)  # 下の色と上の色（0・255 を避ける：0 と 1 の端で決まりが分かれる重ね方がある）
B_FILL, S_FILL = "0.8 0.4 0.2 rg 50 450 300 300 re f", "0.4 0.8 0.6 rg 100 500 300 200 re f"
ON_B, ON_PAPER = (200, 600), (375, 600)  # 上の色が、下の色の上にある所・何も描いていない紙の上にある所


def _lum(c):
    return 0.3 * c[0] + 0.59 * c[1] + 0.11 * c[2]


def _set_lum(c, lum):
    d = lum - _lum(c)
    c = [v + d for v in c]
    low, high = min(c), max(c)
    if low < 0:
        c = [lum + (v - lum) * lum / (lum - low) for v in c]
    if high > 1:
        c = [lum + (v - lum) * (1 - lum) / (high - lum) for v in c]
    return c


def _set_sat(c, sat):
    low, high = min(c), max(c)
    return [0.0] * 3 if high == low else [(v - low) * sat / (high - low) for v in c]


def _sat(c):
    return max(c) - min(c)


def _soft(b, s):
    if s <= 0.5:
        return b - (1 - 2 * s) * b * (1 - b)
    d = ((16 * b - 12) * b + 4) * b if b <= 0.25 else math.sqrt(b)
    return b + (2 * s - 1) * (d - b)


def _hard(b, s):
    return 2 * b * s if s <= 0.5 else 1 - 2 * (1 - b) * (1 - s)


SEPARABLE = {
    "Multiply": lambda b, s: b * s,
    "Screen": lambda b, s: b + s - b * s,
    "Overlay": lambda b, s: _hard(s, b),
    "Darken": min,
    "Lighten": max,
    "ColorDodge": lambda b, s: min(1.0, b / (1 - s)),
    "ColorBurn": lambda b, s: 1 - min(1.0, (1 - b) / s),
    "HardLight": _hard,
    "SoftLight": _soft,
    "Difference": lambda b, s: abs(b - s),
    "Exclusion": lambda b, s: b + s - 2 * b * s,
}
WHOLE = {
    "Hue": lambda b, s: _set_lum(_set_sat(s, _sat(b)), _lum(b)),
    "Saturation": lambda b, s: _set_lum(_set_sat(b, _sat(s)), _lum(b)),
    "Color": lambda b, s: _set_lum(s, _lum(b)),
    "Luminosity": lambda b, s: _set_lum(b, _lum(s)),
}
MODES = [*SEPARABLE, *WHOLE]


def _blend(name: str, b=B, s=S) -> tuple:
    """決まりの式（11.3.5）で重ねた色。"""
    fb, fs = [v / 255 for v in b], [v / 255 for v in s]
    out = [SEPARABLE[name](x, y) for x, y in zip(fb, fs)] if name in SEPARABLE else WHOLE[name](fb, fs)
    return tuple(round(v * 255) for v in out)


def _blended(name: str, content: str, more: str = "", objs: tuple = (), ca: str = "") -> dict:
    res = f"<< /ExtGState << /BM << /BM /{name} {ca} >> >> {more} >>"
    r = _render(pdfmini.build(content, res, list(objs)))
    assert not r.get("approx"), r.get("approx")
    return r


def test_the_reference_colours_differ_between_the_blend_modes():
    """試験そのものの確かめ：選んだ 2 色で、15 個の重ね方が、どれも違う色になる（取り違えても通る試験にしない）。"""
    got = {name: _blend(name) for name in MODES}
    assert len(set(got.values())) == len(MODES), got
    assert all(c != S for c in got.values())


@pytest.mark.parametrize("name", MODES)
def test_a_fill_is_blended_with_what_is_under_it(name):
    r = _blended(name, f"{B_FILL} q /BM gs {S_FILL} Q")
    assert _near(px(r, *ON_B), _blend(name), 3), (px(r, *ON_B), _blend(name))
    assert px(r, 75, 475) == B  # 下の色は、上の色のない所では変わらない
    # 何も描いていない紙は、透ける下地：どの重ね方でも、上の色がそのまま出る（紙の白とは重ねない）
    assert px(r, *ON_PAPER) == S


@pytest.mark.parametrize("name", MODES)
def test_a_half_transparent_fill_is_blended(name):
    r = _blended(name, f"{B_FILL} q /BM gs {S_FILL} Q", ca="/ca 0.5")
    assert _near(px(r, *ON_B), _over(B, _blend(name), 0.5), 3)
    assert _near(px(r, *ON_PAPER), _over(WHITE, S, 0.5))


@pytest.mark.parametrize("name", ["Multiply", "Screen", "Difference", "Luminosity"])
def test_an_image_a_shading_and_a_stroke_are_blended(name):
    im = pdfmini.stream(bytes(S), "/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8")
    r = _blended(name, f"{B_FILL} q /BM gs 300 0 0 200 100 500 cm /Im Do Q", "/XObject << /Im 2 0 R >>", (im,))
    assert _near(px(r, *ON_B), _blend(name), 3) and px(r, *ON_PAPER) == S
    sh = ("/Shading << /S1 << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [0 0 595 0] "
          "/Function << /FunctionType 2 /Domain [0 1] /C0 [0.4 0.8 0.6] /C1 [0.4 0.8 0.6] /N 1 >> /Extend [true true] >> >>")
    r = _blended(name, f"{B_FILL} q /BM gs 100 500 300 200 re W n /S1 sh Q", sh)
    assert _near(px(r, *ON_B), _blend(name), 3) and _near(px(r, *ON_PAPER), S)
    r = _blended(name, f"{B_FILL} q /BM gs 0.4 0.8 0.6 RG 40 w 100 600 m 400 600 l S Q")
    assert _near(px(r, *ON_B), _blend(name), 3) and px(r, *ON_PAPER) == S


def test_a_blend_over_a_half_transparent_backdrop():
    """下の色が半透明（紙が半分透ける）：下の色のある分だけ重ね、ない分は上の色（11.3.6 の式）。"""
    res = "<< /ExtGState << /H << /ca 0.5 >> /BM << /BM /Screen >> >> >>"
    r = _render(pdfmini.build("q /H gs 1 0 0 rg 50 450 300 300 re f Q q /BM gs 0 0 1 rg 100 500 300 200 re f Q", res, []))
    assert _near(px(r, *ON_B), (128, 0, 255))  # 0.5 × 青 ＋ 0.5 × Screen(赤, 青)
    assert px(r, *ON_PAPER) == BLUE and _near(px(r, 75, 475), (255, 128, 128))


def test_a_blend_mode_inside_a_form_and_set_back_to_normal():
    f = _form(f"/BM gs {S_FILL} /NM gs 0.4 0.8 0.6 rg 100 460 100 30 re f", "/ExtGState << /BM << /BM /Difference >> /NM << /BM /Normal >> >>")
    r = _blended("Multiply", f"{B_FILL} /X Do", "/XObject << /X 2 0 R >>", (f,))
    assert _near(px(r, *ON_B), _blend("Difference"), 3) and px(r, *ON_PAPER) == S
    assert px(r, 150, 475) == S  # Normal に戻した後は、ふつうに重ねる


def test_a_blend_mode_given_as_a_list_uses_the_first_known_name():
    res = "<< /ExtGState << /BM << /BM [/NoSuchMode /Multiply] >> >> >>"
    r = _render(pdfmini.build(f"{B_FILL} q /BM gs {S_FILL} Q", res, []))
    assert _near(px(r, *ON_B), _blend("Multiply"), 3) and not r.get("approx")


def test_an_unknown_blend_mode_is_told_and_drawn_normally():
    res = "<< /ExtGState << /BM << /BM /NoSuchMode >> >> >>"
    r = _render(pdfmini.build(f"{B_FILL} q /BM gs {S_FILL} Q", res, []))
    assert px(r, *ON_B) == S and any("重ね方" in k for k in r.get("approx", {}))


# ====================================================================== 透けて描くまとまり
EG = ("/ExtGState << /H << /ca 0.5 /CA 0.5 >> /MU << /BM /Multiply >> /SC << /BM /Screen >> /DI << /BM /Difference >> "
      "/HM << /ca 0.5 /BM /Multiply >> /K << /SMask << /Type /Mask /S /Luminosity /G 3 0 R >> >> >>")
BACK = "0 0 1 rg 50 450 400 350 re f"  # 青い下地（重なる赤と緑の全部の下）
MASK = _group("1 g 100 500 150 300 re f", "DeviceGray")  # x 100〜250 だけが出るマスク


def _placed(group: bytes, how: str = "/H gs", back: str = BACK, objs: tuple = (), told: bool = False) -> dict:
    res = f"<< {EG} /XObject << /X 2 0 R >> >>"
    r = _render(pdfmini.build(f"{back} q {how} /X Do Q", res, [group, MASK, *objs]))
    assert bool(r.get("approx")) == told, r.get("approx")
    return r


@pytest.mark.parametrize("iso", [False, True])
def test_things_inside_a_half_transparent_group_do_not_show_through_each_other(iso):
    """まとまりの全体を、1 つの絵として半透明で重ねる：重なった所は、上の緑だけが、下地に半分透ける（赤は透けない）。"""
    r = _placed(_group(TWO, iso=iso))
    assert _near(px(r, *ONLY_RED), _over(BLUE, RED, 0.5)) and _near(px(r, *ONLY_GREEN), _over(BLUE, GREEN, 0.5))
    assert _near(px(r, *BOTH), _over(BLUE, GREEN, 0.5))
    r = _placed(_group(TWO, iso=iso), back="")  # 何も描いていない紙の上
    assert _near(px(r, *ONLY_RED), _over(WHITE, RED, 0.5)) and _near(px(r, *BOTH), _over(WHITE, GREEN, 0.5))
    assert px(r, 450, 450) == WHITE


def test_a_clip_inside_a_group_does_not_cut_the_group():
    """まとまりの中の q … W … Q は、まとまりを重ねる範囲を狭めない。"""
    r = _placed(_group("q 100 500 100 100 re W n 1 0 0 rg 0 0 595 842 re f Q 0 1 0 rg 300 600 100 100 re f"), back="")
    assert _near(px(r, 150, 550), _over(WHITE, RED, 0.5)) and _near(px(r, 350, 650), _over(WHITE, GREEN, 0.5))
    assert px(r, 250, 550) == WHITE


def test_a_group_placed_inside_a_clip_is_cut():
    r = _placed(_group(TWO), how="150 520 200 120 re W n /H gs")
    assert _near(px(r, 250, 600), _over(BLUE, GREEN, 0.5)) and px(r, *ONLY_RED) == BLUE and px(r, *ONLY_GREEN) == BLUE


def test_a_blend_mode_inside_a_group_that_is_not_isolated_blends_with_the_page():
    """下と切り離さないまとまりの中の重ね方は、まとまりの下のページのものと重ねる。"""
    body = f"/DI gs {S_FILL}"
    r = _placed(_group(body, res=EG), how="", back=B_FILL)
    assert _near(px(r, *ON_B), _blend("Difference"), 3) and px(r, *ON_PAPER) == S
    r = _placed(_group(body, res=EG), back=B_FILL)  # 半透明で置く
    assert _near(px(r, *ON_B), _over(B, _blend("Difference"), 0.5), 3) and _near(px(r, *ON_PAPER), _over(WHITE, S, 0.5))


def test_a_blend_mode_inside_an_isolated_group_does_not_see_the_page():
    """下と切り離すまとまり（I）の中は、透ける下地から始まる：中の重ね方は、まとまりの中のものとだけ重ねる。"""
    r = _placed(_group(f"/DI gs {S_FILL}", iso=True, res=EG), how="", back=B_FILL)
    assert px(r, *ON_B) == S and px(r, *ON_PAPER) == S
    # まとまりの中に下の色があれば、それと重ねる
    r = _placed(_group(f"{B_FILL} /DI gs {S_FILL}", iso=True, res=EG), how="", back="")
    assert _near(px(r, *ON_B), _blend("Difference"), 3) and px(r, *ON_PAPER) == S
    r = _placed(_group(f"/DI gs {S_FILL}", iso=True, res=EG), back=B_FILL)  # 半透明で置く
    assert _near(px(r, *ON_B), _over(B, S, 0.5)) and _near(px(r, *ON_PAPER), _over(WHITE, S, 0.5))


def test_a_group_placed_with_a_blend_mode_is_blended_as_a_whole():
    """重ね方で置いたまとまり：重なった所は、上の緑だけを、下地と重ねる（赤と緑を、続けて 2 回重ねない）。"""
    r = _placed(_group(TWO), how="/MU gs", back="1 1 0 rg 50 450 400 350 re f")
    assert px(r, *ONLY_RED) == RED and px(r, *ONLY_GREEN) == GREEN  # Multiply(黄, 赤)＝赤、Multiply(黄, 緑)＝緑
    assert px(r, *BOTH) == GREEN  # 続けて重ねると、Multiply(赤, 緑)＝黒
    r = _placed(_group(TWO), how="/HM gs", back="1 1 0 rg 50 450 400 350 re f")
    assert _near(px(r, *BOTH), _over(YELLOW, GREEN, 0.5))
    r = _placed(_group(TWO, iso=True), how="/SC gs", back="")  # 何も描いていない紙の上：そのままの色
    assert px(r, *ONLY_RED) == RED and px(r, *BOTH) == GREEN


def test_a_group_under_a_mask_is_masked_as_a_whole():
    for iso in (False, True):
        r = _placed(_group(TWO, iso=iso), how="/K gs")
        assert px(r, *ONLY_RED) == RED and px(r, 225, 600) == GREEN  # マスクの白い所（x 100〜250）
        assert px(r, 275, 600) == BLUE and px(r, *ONLY_GREEN) == BLUE  # マスクの黒い所：下地のまま
    r = _placed(_group(f"/DI gs {S_FILL}", res=EG), how="/K gs", back=B_FILL)  # 中の重ね方は、ページのものと重ねる
    assert _near(px(r, *ON_B), _blend("Difference"), 3) and px(r, 300, 600) == B and px(r, *ON_PAPER) == WHITE


def test_groups_inside_groups():
    inner = _group(TWO)
    outer = _group("0 g 80 480 150 100 re f /H gs /In Do", res=EG + " /XObject << /In 4 0 R >>")
    r = _placed(outer, objs=(inner,))
    assert _near(px(r, 90, 490), _over(BLUE, (0, 0, 0), 0.5))  # 外のまとまりの黒
    assert _near(px(r, *BOTH), _over(BLUE, GREEN, 0.25))  # 内 0.5 × 外 0.5。赤は透けない
    assert _near(px(r, 150, 520), _over(BLUE, _over((0, 0, 0), RED, 0.5), 0.5))  # 黒の上の赤：内のまとまりの中で先に重なる
    # 下と切り離すまとまりの中の、切り離さないまとまり：中の重ね方は、外のまとまりの中のものとだけ重ねる
    inner = _group(f"/DI gs {S_FILL}", res=EG)
    outer = _group(f"{B_FILL} /In Do", iso=True, res=EG + " /XObject << /In 4 0 R >>")
    r = _placed(outer, back="1 1 0 rg 0 0 595 842 re f", objs=(inner,))
    assert _near(px(r, *ON_B), _over(YELLOW, _blend("Difference"), 0.5), 3) and _near(px(r, *ON_PAPER), _over(YELLOW, S, 0.5))


def test_a_group_placed_plainly_is_drawn_straight_onto_the_page():
    """透け方・重ね方・マスクなしで置いた、中に重ね方のないまとまりは、別の絵を作らない（じかに描いても同じ）。"""
    made = []
    real = canvas.Canvas.__init__

    def init(self, *a, **k):
        made.append(1)
        real(self, *a, **k)

    canvas.Canvas.__init__ = init
    try:
        r = _placed(_group(TWO), how="")
        assert len(made) == 1 and px(r, *BOTH) == GREEN
        made.clear()
        _placed(_group(TWO))
    finally:
        canvas.Canvas.__init__ = real


def test_groups_nested_too_deep_are_drawn_one_by_one_and_told():
    """別の絵を 3 枚まで重ねる。それより深いまとまりは、中の 1 つ 1 つを半透明で描いて、伝える。"""
    objs = [_group(TWO)]  # 5 番
    for n in range(4):  # 6〜9 番：どれも、前のまとまりを半透明で置く
        objs.append(_group("/H gs /In Do", res=f"{EG} /XObject << /In {5 + n} 0 R >>"))
    res = f"<< {EG} /XObject << /X 9 0 R >> >>"
    r = _render(pdfmini.build("q /H gs /X Do Q", res, [_group(""), MASK, _group(""), *objs]))
    assert any("別に描けない" in k for k in r.get("approx", {}))
    assert _near(px(r, *ONLY_RED), _over(WHITE, RED, 0.5 ** 5), 3)
    # 3 つまでの深さは、言わない
    res = f"<< {EG} /XObject << /X 7 0 R >> >>"
    r = _render(pdfmini.build("q /H gs /X Do Q", res, [_group(""), MASK, _group(""), *objs[:3]]))
    assert not r.get("approx") and _near(px(r, *BOTH), _over(WHITE, GREEN, 0.5 ** 3), 3)


def test_a_blend_inside_a_group_that_is_itself_blended():
    """重ね方で置いた、下と切り離さないまとまりの中の重ね方：中のものは、ページのものと重ねてから、まとまりだけの色を取り出し
    （11.4.8）、その色を、置いたときの重ね方でページに重ねる。"""
    yellow = "1 1 0 rg 50 450 400 350 re f"
    r = _placed(_group(f"/SC gs {TWO}", res=EG), how="/MU gs", back=yellow)
    # Screen(黄, 赤)＝黄、Multiply(黄, 黄)＝黄。中だけで重ねると、赤・緑のまま残る
    assert px(r, *ONLY_RED) == YELLOW and px(r, *BOTH) == YELLOW and px(r, *ONLY_GREEN) == YELLOW
    inner = _blend("Difference")
    r = _placed(_group(f"/DI gs {S_FILL}", res=EG), how="/MU gs", back=B_FILL)
    assert _near(px(r, *ON_B), _blend("Multiply", s=inner), 3)
    assert _near(px(r, *ON_PAPER), S)  # 何も描いていない紙の上：そのままの色
    # 中のものが半透明：まとまりの色は Difference(B, S)、濃さは 0.5
    r = _placed(_group(f"/H gs /DI gs {S_FILL}", res=EG), how="/MU gs", back=B_FILL)
    assert _near(px(r, *ON_B), _over(B, _blend("Multiply", s=inner), 0.5), 4)
    assert _near(px(r, *ON_PAPER), _over(WHITE, S, 0.5), 3)
    # まとまりも半透明で置く
    r = _placed(_group(f"/DI gs {S_FILL}", res=EG), how="/HM gs", back=B_FILL)
    assert _near(px(r, *ON_B), _over(B, _blend("Multiply", s=inner), 0.5), 3)
    # 中のまとまりも、重ね方で置く
    deep = _group(f"/DI gs {S_FILL}", res=EG)
    outer = _group("/SC gs /In Do", res=EG + " /XObject << /In 4 0 R >>")
    r = _placed(outer, how="/MU gs", back=B_FILL, objs=(deep,))
    assert _near(px(r, *ON_B), _blend("Multiply", s=_blend("Screen", s=inner)), 4)


def test_a_blend_inside_a_blended_group_too_large_to_draw_twice_is_told(monkeypatch):
    """2 回描く絵を持てない大きさ：中の重ね方を、まとまりの中のものとだけ重ねて（近い形）、伝える。"""
    r = _render(pdfmini.build("q /MU gs /X Do Q", f"<< {EG} /XObject << /X 2 0 R >> >>", [_group(TWO), MASK]))
    monkeypatch.setattr(paint, "APART_BYTES", r["w"] * r["h"] * 4 * 2 + 1000)  # 別の絵 1 枚は持てる
    r = _placed(_group(f"/SC gs {TWO}", res=EG), how="/MU gs", back="1 1 0 rg 50 450 400 350 re f", told=True)
    assert any("重ね方" in k for k in r["approx"])
    assert px(r, *ONLY_RED) == RED  # 近い形：赤は、黄と Screen で重ねていない


def test_a_knockout_group_is_drawn_as_an_ordinary_group_and_told():
    """限界：中のものどうしを重ねないまとまり（Knockout）は、ふつうのまとまりとして描いて、伝える。"""
    group = _group("1 0 0 rg 100 500 200 200 re f /H gs 0 1 0 rg 200 550 200 200 re f", res=EG).replace(b"/S /Transparency", b"/S /Transparency /K true")
    r = _placed(group, how="", told=True)
    assert list(r["approx"]) == ["透けて描くまとまり（Knockout）"]
    assert _near(px(r, *BOTH), _over(RED, GREEN, 0.5))  # 決まりでは、下地の青に緑を重ねた色
    assert not _placed(group.replace(b"/K true", b"/K false"), how="").get("approx")


def test_mixing_two_rows_with_one_alpha():
    under, over = bytes(range(256)) * 3, bytes(reversed(range(256))) * 3
    for a in (1, 37, 127, 128, 200, 254):
        assert canvas._mix_same(under, over, a) == bytes((u * (255 - a) + o * a + 127) // 255 for u, o in zip(under, over))
