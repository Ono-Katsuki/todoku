"""PDF の、見えない描き方（文字の描き方 3・7）の字のうち、画像の上にない字：読み順に残して、位置を伝える（読む B）。

前は、見えない描き方の字を伝えるのは、画像の上にあるとき（OCR の字）だけだった：画像の上にない見えない字は、
見える字と同じに読み、何も伝えなかった。いまは、その字の位置を「見えない描き方の字（画像の上にない字）」で伝える
（線や形で描いた字に重ねた検索用の字か、見せずに入れた字かは、とどくは決めない）。
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from test_pdf_cover import DRAW, NAME, XIM, IMG, _audit_ok, _page, _write  # noqa: E402

from todoku import api  # noqa: E402

OFF = "見えない描き方の字（画像の上にない字）"
OCR = "OCR の字（画像の上の見えない字）"
SECRET = b"BT 3 Tr /F1 12 Tf 150 720 Td (Hidden) Tj ET\n"


def _flags(path: str) -> dict:
    return {r["title"]: r for r in api.grade(path)["reasons"]}


def test_invisible_text_off_any_picture_is_told_with_its_place(tmp_path):
    p = _write(tmp_path, NAME + SECRET)
    pg, md = _page(p)
    assert any("Hidden" in ln for ln in pg["reading"])  # 読み順には残す（字は落とさない）
    f = _flags(p)
    assert OFF in f and OCR not in f and f[OFF]["grade"] == "B"
    assert len(f[OFF]["where"]) == 1 and "「Hidden」" in f[OFF]["where"][0] and "Name" not in f[OFF]["where"][0]
    assert "view_page" in f[OFF]["detail"] and OFF in md
    assert api.grade(p)["grade"] == "B" and _audit_ok(p)


def test_visible_text_alone_is_not_told(tmp_path):
    p = _write(tmp_path, NAME + b"BT /F1 12 Tf 150 720 Td (Shown) Tj ET\n")
    assert OFF not in _flags(p) and OFF not in _page(p)[1]


@pytest.mark.parametrize("mode", [b"3", b"7"])
def test_both_invisible_modes_are_told(tmp_path, mode):
    p = _write(tmp_path, NAME + b"BT " + mode + b" Tr /F1 12 Tf 150 720 Td (Hidden) Tj ET\n")
    assert OFF in _flags(p)


def test_invisible_text_on_a_picture_is_the_ocr_text(tmp_path):
    # 画像を先に描き、その上に見えない字：前と同じに「OCR の字」だけ
    p = _write(tmp_path, NAME + DRAW + SECRET, XIM, IMG)
    f = _flags(p)
    assert OCR in f and OFF not in f


def test_invisible_text_on_and_off_a_picture_is_told_separately(tmp_path):
    body = NAME + DRAW + SECRET + b"BT 3 Tr /F1 12 Tf 72 500 Td (Elsewhere) Tj ET\n"
    p = _write(tmp_path, body, XIM, IMG)
    f = _flags(p)
    assert OCR in f and OFF in f
    assert len(f[OFF]["where"]) == 1 and "「Elsewhere」" in f[OFF]["where"][0] and "Hidden" not in f[OFF]["where"][0]


def test_only_the_invisible_part_of_a_line_is_named(tmp_path):
    body = b"BT /F1 12 Tf 72 720 Td (Name: ) Tj 3 Tr (Hidden) Tj 0 Tr ( end) Tj ET\n"
    p = _write(tmp_path, body)
    f = _flags(p)
    assert OFF in f and "「Hidden」" in f[OFF]["where"][0]


def test_many_invisible_lines_are_counted(tmp_path):
    body = NAME + b"".join(b"BT 3 Tr /F1 10 Tf 72 %d Td (Line %d) Tj ET\n" % (700 - 20 * i, i) for i in range(15))
    p = _write(tmp_path, body)
    where = _flags(p)[OFF]["where"]
    assert len(where) == 13 and "「Line11」" in where[11] and where[-1].startswith("ほか")  # 12 行までを示し、残りはまとめる
