"""PDF に入れた、字でないもの（チェック・ラジオの印、字を囲む ○、印影の画像）の、一部だけが見えない所。

前は、その四角の全体の画素だけを見ていた：どこか 1 か所でも見えれば、黒い帯に大半がかかる印も、後から描かれるものに大半が
隠れた画像も、何も伝えずに「書けた」と返していた。
いまは、描画が、その印・○・画像だけを、白い下地と黒い下地にも描く（下地が同じ色でも、墨の出る所が分かる）。白い下地に出る墨・
黒い下地に出る墨のどちらで見ても、ページの絵で色が変わった所が半分に満たなければ、位置（mm）つきで知らせる（止めない）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

import pdfmini  # noqa: E402
from test_pdf_ink_marks import CROSS, FORM, ZADB, _box, _choice, _fill, _mm, _seal, _square, _told  # noqa: E402

from todoku.pdf.render import fontmap  # noqa: E402

PART = "の半分より多く"


def _part(r: dict) -> list[str]:
    return [w for w in r["warnings"] if PART in w]


# ====================================================================== チェックの印（部品は 100〜112 pt、印の線は 102〜110 pt）
def test_a_mark_mostly_on_a_black_band_is_told_with_the_place_of_the_hidden_part(tmp_path):
    r = _fill(tmp_path, _box(content="0 g 90 690 18 40 re f"), {"a": True})  # 帯は 108 pt まで：印の 4 分の 3 が、同じ色の上
    assert r["read_back_ok"] and r["proof_ok"], r  # 止めない
    told = _part(r)
    assert len(told) == 1 and told[0].startswith("a：書いた印の半分より多く（墨の ") and "描いても見えません" in told[0]
    assert "同じ色の下地の上か、後から描かれるものの下" in told[0] and "view_page" in told[0]
    x0, _y0, x1, _y1 = _mm(told[0])  # 見えない所：印の左から、帯の右の端（108 pt = 38.1 mm）まで
    assert abs(x0 - 102 / 72 * 25.4) < 0.6 and abs(x1 - 108 / 72 * 25.4) < 0.6


def test_a_mark_whose_smaller_part_is_on_a_black_band_is_not_told(tmp_path):
    r = _fill(tmp_path, _box(content="0 g 90 690 14 40 re f"), {"a": True})  # 帯は 104 pt まで：印の 4 分の 1
    assert r["read_back_ok"] and not _told(r), r


def test_a_mark_mostly_on_a_dark_band_is_told_as_hard_to_see(tmp_path):
    r = _fill(tmp_path, _box(content="0 0 0.6 rg 90 690 18 40 re f"), {"a": True})
    assert r["read_back_ok"], r
    told = _part(r)
    assert len(told) == 1 and "下地との明るさの差が小さく" in told[0] and "見えにくいかもしれません" in told[0] and "1ページ" in told[0]


def _half_covered(to: int) -> bytes:
    """チェックボックスの後に描かれる、白く塗った注釈が、部品の左から to pt までを覆う。"""
    objs = ["<< /Type /Annot /Subtype /Widget /FT /Btn /T (a) /Rect [100 700 112 712] /F 4 /V /Off /AS /Off "
            "/AP << /N << /Yes 3 0 R /Off 4 0 R >> >> >>",
            pdfmini.stream(CROSS, FORM), pdfmini.stream("", FORM),
            f"<< /Type /Annot /Subtype /Square /Rect [95 695 {to} 720] /F 4 /AP << /N 6 0 R >> >>",
            pdfmini.stream(f"1 g 0 0 {to - 95} 25 re f", f"/Type /XObject /Subtype /Form /BBox [0 0 {to - 95} 25]")]
    return pdfmini.build("", "<< >>", objs, annots="/Annots [2 0 R 5 0 R]", acroform="/AcroForm << /Fields [2 0 R] >>")


def test_a_mark_mostly_under_something_drawn_later_is_told(tmp_path):
    r = _fill(tmp_path, _half_covered(108), {"a": True})
    assert r["read_back_ok"], r
    assert len(_part(r)) == 1 and "描いても見えません" in _part(r)[0]
    r = _fill(tmp_path, _half_covered(104), {"a": True})
    assert r["read_back_ok"] and not _told(r), r


@pytest.mark.parametrize("content", ["", "0 g 90 690 40 40 re f", "0 g 90 690 18 40 re f"], ids=["白いページ", "黒い帯", "黒い帯に半分"])
def test_a_mark_that_paints_its_own_ground_is_not_told(tmp_path, content):
    """「入」の見た目が、白い下地を塗ってから印を描く：下地がページと同じ色でも、印は見える（塗った下地を、見えない墨と数えない）。"""
    r = _fill(tmp_path, _box(on="1 g 0 0 12 12 re f " + CROSS, content=content), {"a": True})
    assert r["read_back_ok"] and not _told(r), r


def test_a_mark_whose_shape_is_not_on_this_machine_is_not_judged_by_the_gray_box(tmp_path, monkeypatch):
    """印の字の形がないと、描画は灰色の四角を描く：その四角の墨では決めない（前と同じに、印の色と下地の色で決める）。"""
    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    r = _fill(tmp_path, _box(on=ZADB, res=" /Resources << /Font << /ZaDb 5 0 R >> >>", content="0 g 90 690 18 40 re f"), {"a": True})
    assert not _part(r), r


# ====================================================================== 字を囲む ○（「Yes」は 200 pt から）
def test_a_circle_mostly_on_a_black_band_is_told(tmp_path):
    r = _fill(tmp_path, _choice("0 g 180 680 36 50 re f"), {"@c1": {"circle": "Yes"}})  # 帯は 216 pt まで
    assert r["read_back_ok"], r
    told = _part(r)
    assert len(told) == 1 and "書いた○の半分より多く" in told[0] and "描いても見えません" in told[0]
    box = r["circled"][0]["box_mm"]
    x0, y0, x1, y1 = _mm(told[0])
    assert box[0] - 0.5 <= x0 < x1 <= 216 / 72 * 25.4 + 0.5 and box[1] - 0.5 <= y0 < y1 <= box[3] + 0.5


def test_a_circle_that_only_touches_a_black_band_is_not_told(tmp_path):
    box = _fill(tmp_path, _choice(), {"@c1": {"circle": "Yes"}})["circled"][0]["box_mm"]
    edge = box[0] * 72 / 25.4 + 3  # ○ の左の端から 3 pt まで
    r = _fill(tmp_path, _choice(f"0 g 150 680 {edge - 150:.2f} 50 re f"), {"@c1": {"circle": "Yes"}})
    assert r["read_back_ok"] and not _told(r), r


# ====================================================================== 印影の画像（枠は、左 20〜40 mm・上 80〜100 mm）
RED = (255, 0, 0, 255)


def test_a_seal_image_mostly_on_a_ground_of_its_colour_is_told(tmp_path):
    r = _seal(tmp_path, _square(RED), "1 0 0 rg 50 550 50 80 re f")  # 赤い帯は 100 pt（35.3 mm）まで：枠の 4 分の 3
    assert r["read_back_ok"], r
    told = _part(r)
    assert len(told) == 1 and "書いた画像の半分より多く" in told[0] and "描いても見えません" in told[0]
    x0, y0, x1, y1 = _mm(told[0])
    assert 19.5 <= x0 < x1 <= 35.8 and 79.5 <= y0 < y1 <= 100.5


def test_a_seal_image_whose_smaller_part_is_on_a_ground_of_its_colour_is_not_told(tmp_path):
    r = _seal(tmp_path, _square(RED), "1 0 0 rg 50 550 20 80 re f")  # 赤い帯は 70 pt（24.7 mm）まで：枠の 4 分の 1
    assert r["read_back_ok"] and not _told(r), r
