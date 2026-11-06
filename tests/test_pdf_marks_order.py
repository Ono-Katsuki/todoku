"""PDF：結合する記号の箱が次の字より少し右にあっても、基の字の後ろに置く（タイ語の字の順が変わっていた）。"""

from __future__ import annotations

from todoku.pdf.pagetext import Ch, make_lines


def _ch(text: str, x: float, seq: int, w: float = 6.0) -> Ch:
    return Ch(text=text, x0=x, y0=90.0, x1=x + w, y1=104.0, size=14.0, vertical=False, base=100.0, invisible=False, seq=seq)


def test_thai_marks_stay_after_their_base():
    # 描いた順は ช ื ่ อ（字の順）。記号の箱の左端（64.3・64.25）は次の字 อ の左端（64.2）より右
    chars = [_ch("ช", 58.6, 0), _ch("ื", 64.3, 1, 6.8), _ch("่", 64.25, 2, 6.8), _ch("อ", 64.2, 3)]
    assert make_lines(chars)[0].text == "ชื่อ"


def test_arabic_marks_are_left_to_the_logical_order():
    # 右から左の字は、描いた順でなく位置から字の順を決める（ここでは記号を動かさない）
    chars = [_ch("ل", 70.0, 0), _ch("َ", 66.0, 1, 3.0), _ch("ب", 60.0, 2)]
    assert set(make_lines(chars)[0].text) == {"ل", "َ", "ب"}


def test_latin_text_is_unchanged():
    chars = [_ch("a", 10.0, 0), _ch("b", 16.0, 1), _ch("c", 22.0, 2)]
    assert make_lines(chars)[0].text == "abc"
