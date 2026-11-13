"""書いた字が見えるかを、字（グリフ）ごとに確かめる。欄の見た目の字は、欄の書体の指定（/DA）の色で描く。

前は、欄の全体の画素だけを見ていた：欄の一部だけが黒い帯にかかる字は、ほかの字が見えるので、何も伝えなかった。字の色は、どの欄も
黒で描いていた：濃い下地に白い字と決めた欄（/MK の BG と /DA の色）は、黒い字になって読めなかった。
いまは、書いた欄の字を描くときに、字ごとに、墨の画素とその下の色を控える：墨がどれも下と同じ色の字は読み戻しの違い、墨の半分より多くが
比 3:1 未満の下地の上にある字と、描いた後に別のものが上に重なった字は、位置（mm）つきで知らせる。
"""
from __future__ import annotations

import random
import re

import pytest

import pdfmini
from todoku import api
from todoku.pdf import inkcheck
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import da_color
from todoku.pdf.form import read_fields
from todoku.pdf.render.canvas import Canvas, Clip, covered, paint, scan

LONG = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
ACRO = "/AcroForm << /Fields [2 0 R] /DA (/Helv 0 Tf 0 g) >>"
FONT = "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> >>"


def _fill(tmp_path, blob, values):
    src = tmp_path / "a.pdf"
    src.write_bytes(blob)
    out = tmp_path / "b.pdf"
    api._WINDOWS.clear()
    return api.fill(str(src), values, str(out), overwrite=True), out


def _widget(rect="100 700 300 716", extra="", da="/DA (/Helv 12 Tf 0 g)", ft="/FT /Tx"):
    return f"<< /Type /Annot /Subtype /Widget {ft} /T (a) /Rect [{rect}] /F 4 {da} {extra} >>"


def _form(content="", extra="", da="/DA (/Helv 12 Tf 0 g)", acro=ACRO, ft="/FT /Tx"):
    return pdfmini.build(content, "<< >>", [_widget(extra=extra, da=da, ft=ft)], annots="/Annots [2 0 R]", acroform=acro)


def _ap(out) -> bytes:
    doc = PdfDocument(out.read_bytes())
    f = next(x for x in read_fields(doc) if x.name == "a")
    wd = doc.resolve(f.widgets[0].ref)
    return doc.decode(doc.resolve(doc.resolve(wd["AP"])["N"]))


def _told(r) -> list[str]:
    """見え方についての知らせ。"""
    return [w for w in r["warnings"] if "明るさ" in w or "見え" in w or "重なって" in w]


def _mm(text: str) -> list[float]:
    return [float(v) for v in re.search(r"左 ([\d.]+)・上 ([\d.]+)・右 ([\d.]+)・下 ([\d.]+) mm", text).groups()]


# ====================================================================== 欄の一部だけが下地に紛れる
def test_letters_on_a_black_band_are_a_read_back_difference(tmp_path):
    # 欄の左 40pt だけが黒い帯の上：ほかの字は見えるので、欄の全体の画素では分からない
    r, _out = _fill(tmp_path, _form("0 g 90 690 40 40 re f"), {"a": LONG})
    assert r["read_back_ok"] is False and r["proof_ok"]
    d = next(x for x in r["read_back_differences"] if x["id"] == "a")
    assert "描いた字 26 個のうち" in d["read_back"] and "下地と同じ色" in d["read_back"] and "ほかの字は見えます" in d["read_back"]
    n = int(re.search(r"のうち (\d+) 個は、描いても見えません", d["read_back"])[1])
    assert 3 <= n <= 7  # 帯の中に入りきる字（12pt の字で、欄の左 30pt ほど）
    x0, y0, x1, y1 = _mm(d["read_back"])
    # 見えない字の位置は、帯の中（帯は、ページの左から 90〜130pt = 31.8〜45.9 mm）
    assert 35.0 <= x0 < x1 <= 46.0 and 44.0 <= y0 < y1 <= 50.5


def test_letters_on_a_dark_band_are_told_with_their_place(tmp_path):
    r, _out = _fill(tmp_path, _form("0 0 0.6 rg 90 690 40 40 re f"), {"a": LONG})
    assert r["read_back_ok"], r
    told = _told(r)
    assert len(told) == 1 and told[0].startswith("a：描いた字 26 個のうち") and "見えにくいかもしれません" in told[0]
    assert float(re.search(r"比 ([\d.]+):1 から", told[0])[1]) < 3.0
    x0, _y0, x1, _y1 = _mm(told[0])
    assert 35.0 <= x0 < x1 <= 46.0


def test_a_thin_line_through_the_letters_is_not_told(tmp_path):
    # 字にかかる細い罫線：字の墨のほとんどは白い下地の上
    r, _out = _fill(tmp_path, _form("0 g 0.5 w 100 706 m 300 706 l S"), {"a": "ABCDEFGH-"})
    assert r["read_back_ok"] and r["proof_ok"] and not _told(r), r


def test_a_plain_write_is_not_told(tmp_path):
    r, _out = _fill(tmp_path, _form(), {"a": LONG})
    assert r["read_back_ok"] and r["proof_ok"] and not _told(r), r


def test_a_value_drawn_into_a_flat_page_is_checked_letter_by_letter(tmp_path):
    """フォームのない PDF：とどくがページに描いた値（印つきの範囲）の字も、字ごとに見る。"""
    cells = "0.5 w 100 700 80 20 re S 180 700 220 20 re S BT /F1 10 Tf 105 706 Td (Name) Tj ET "

    def run(extra: str):
        src = tmp_path / "a.pdf"
        src.write_bytes(pdfmini.build(cells + extra, FONT, []))
        api._WINDOWS.clear()
        fid = next(f["id"] for f in api.fields(str(src)) if f["id"].startswith("p1.t"))
        return fid, api.fill(str(src), {fid: LONG}, str(tmp_path / "b.pdf"), overwrite=True)

    fid, r = run("")
    assert r["read_back_ok"] and not _told(r), r
    # マスの左の方に、濃い三角（マスの形は変えない）：その上の字だけを知らせる
    fid, r = run("0 0 0.6 rg 185 703 m 215 703 l 200 718 l f")
    told = _told(r)
    assert r["read_back_ok"] and len(told) == 1 and told[0].startswith(f"{fid}：描いた字 26 個のうち"), r
    x0, _y0, x1, _y1 = _mm(told[0])
    assert 65.0 <= x0 < x1 <= 76.0  # 三角は、ページの左から 185〜215pt = 65.3〜75.8 mm


# ====================================================================== 描いた後に、上に別のものが重なる
def _stamped(ops: str) -> bytes:
    ap = f"<< /Type /XObject /Subtype /Form /BBox [0 0 40 30] /Length {len(ops)} >>\nstream\n{ops}\nendstream"
    stamp = "<< /Type /Annot /Subtype /Stamp /Rect [95 695 135 725] /F 4 /AP << /N 4 0 R >> >>"
    return pdfmini.build("", "<< >>", [_widget(), stamp, ap], annots="/Annots [2 0 R 3 0 R]", acroform=ACRO)


@pytest.mark.parametrize("ops", ["1 0 0 rg 0 0 40 30 re f", "1 g 0 0 40 30 re f"])
def test_something_drawn_later_over_the_letters_is_told(tmp_path, ops):
    # 欄より後ろの注釈（スタンプ）の見た目が、欄の左の部分に重なる（白い塗りは、字を消す）
    r, _out = _fill(tmp_path, _stamped(ops), {"a": LONG})
    told = _told(r)
    assert len(told) == 1 and "後から描かれるもの" in told[0] and "重なっています" in told[0], r
    x0, _y0, x1, _y1 = _mm(told[0])
    assert 35.0 <= x0 < x1 <= 48.5  # スタンプは、ページの左から 95〜135pt = 33.5〜47.6 mm


def test_a_later_annotation_that_does_not_touch_the_letters_is_not_told(tmp_path):
    r, _out = _fill(tmp_path, _stamped("1 0 0 rg 0 0 2 2 re f"), {"a": LONG})
    assert r["read_back_ok"] and not _told(r), r


# ====================================================================== 字の色は、欄の書体の指定（/DA）の色
@pytest.mark.parametrize("da, op", [
    ("/Helv 12 Tf 0 g", b"0 g"), ("/Helv 12 Tf 1 g", b"1 g"), ("/Helv 12 Tf 1 1 0 rg", b"1 1 0 rg"),
    ("0 0 0.502 rg /Helv 12 Tf", b"0 0 0.502 rg"), ("/Helv 12 Tf 0.2 0.4 0 0.1 k", b"0.2 0.4 0 0.1 k"),
])
def test_the_text_is_drawn_in_the_color_of_the_default_appearance(tmp_path, da, op):
    r, out = _fill(tmp_path, _form(da=f"/DA ({da})", extra="/MK << /BG [0.5 0.5 0.5] >>"), {"a": "AB"})
    assert r["proof_ok"], r
    assert b"BT\n" + op + b"\n" in _ap(out)


def test_the_color_comes_from_the_form_when_the_field_has_none(tmp_path):
    blob = _form(da="", acro="/AcroForm << /Fields [2 0 R] /DA (/Helv 0 Tf 0 0 0.5 rg) >>")
    r, out = _fill(tmp_path, blob, {"a": "AB"})
    assert r["proof_ok"] and r["read_back_ok"] and b"BT\n0 0 0.5 rg\n" in _ap(out), r


def test_light_text_on_the_dark_ground_of_its_own_field_is_readable(tmp_path):
    # 濃い下地（/MK の BG）に白い字（/DA）と決めた欄：前は黒い字で描いて、読めなかった
    dark = "/MK << /BG [0 0 0.4] >>"
    r, out = _fill(tmp_path, _form(extra=dark, da="/DA (/Helv 12 Tf 1 g)"), {"a": "ABCDEFGH"})
    assert r["read_back_ok"] and r["proof_ok"] and not _told(r), r
    assert b"BT\n1 g\n" in _ap(out)
    # 濃い下地に黒い字と決めた欄は、そのとおりに描いて、見えにくいことを知らせる
    r, _out = _fill(tmp_path, _form(extra=dark), {"a": "ABCDEFGH"})
    told = _told(r)
    assert r["read_back_ok"] and len(told) == 1 and "描いた字 8 個のうち 8 個" in told[0], r


def test_white_text_on_a_white_page_is_a_read_back_difference(tmp_path):
    # 欄が白い字と決めていて、下地がない：決めたとおりに描くと見えないので、読み戻しの違いとして伝える
    r, _out = _fill(tmp_path, _form(da="/DA (/Helv 12 Tf 1 g)"), {"a": "AB"})
    assert r["read_back_ok"] is False and any("見えません" in d["read_back"] for d in r["read_back_differences"])


def test_choice_fields_use_the_color_too(tmp_path):
    for ff, value in ((131072, "Two"), (0, "Two")):  # コンボ・リスト
        extra = f"/Ff {ff} /Opt [(One) (Two) (Three)]"
        rect = "100 700 300 716" if ff else "100 660 300 716"
        blob = pdfmini.build("", "<< >>", [_widget(rect, extra, "/DA (/Helv 10 Tf 0.8 0 0 rg)", "/FT /Ch")], annots="/Annots [2 0 R]",
                             acroform=ACRO)
        r, out = _fill(tmp_path, blob, {"a": value})
        assert r["proof_ok"], r
        ap = _ap(out)
        assert b"BT\n0.8 0 0 rg\n" in ap and b"BT\n0 g\n" not in ap


@pytest.mark.parametrize("da, op", [
    ("", b"0 g"), ("/F 9 Tf", b"0 g"), ("/F 9 Tf .5 g", b"0.5 g"), ("1 0 0 rg /F 9 Tf 0 g", b"0 g"), ("0 g /F 9 Tf 1 0 0 rg", b"1 0 0 rg"),
    ("/F 9 Tf 0 0 0 1 k", b"0 0 0 1 k"), ("/F 9 Tf 2 -1 0 rg", b"1 0 0 rg"),  # 範囲の外の数は、0〜1 に
    ("/F 9 Tf 1 1 rg", b"0 g"),  # 数の足りない命令は、色としない
    ("q 1 0 0 RG /F 9 Tf", b"0 g"),  # 線の色（大文字）は、字の色ではない
    ("/F1 12 Tf 0.000 0.000 0.502 rg", b"0 0 0.502 rg"),
])
def test_the_color_of_a_default_appearance(da, op):
    assert da_color(da) == op


# ====================================================================== 部品
def test_covered_is_the_pixels_that_paint_changes():
    rnd = random.Random(5)
    for _ in range(40):
        cv = Canvas(60, 50)
        before = bytes(cv.buf)
        pts = [(rnd.uniform(-5, 65), rnd.uniform(-5, 55)) for _ in range(rnd.choice([3, 4, 5]))]
        rows = scan([pts])
        clip = Clip(rnd.randint(0, 10), rnd.randint(0, 10), rnd.randint(40, 60), rnd.randint(30, 50))
        if rnd.random() < 0.5:
            clip = clip.with_rows(scan([[(rnd.uniform(0, 60), rnd.uniform(0, 50)) for _ in range(3)]]))
        dx, dy = rnd.randint(-6, 6), rnd.randint(-6, 6)
        px = covered(cv, rows, clip, dx, dy)
        assert bytes(cv.buf) == before  # 見るだけで、塗らない
        paint(cv, rows, (0, 0, 0), 1.0, clip, dx, dy)
        changed = {q for q in range(60 * 50) if cv.buf[q * 3:q * 3 + 3] != before[q * 3:q * 3 + 3]}
        assert {q for q, _a in px} == changed and all(1 <= a <= 255 for _q, a in px)


def _probe(ground: dict, glyphs: list, later: dict | None = None):
    """小さな絵（幅 10）に、字（画素の番号と濃さの並び）を順に塗って、控えを作る。later：描き終えた後に上から塗る色。"""
    cv = Canvas(10, 4)
    for q, c in ground.items():
        cv.buf[q * 3:q * 3 + 3] = bytes(c)
    g = inkcheck._Glyphs()
    for px, rgb in glyphs:
        g.glyph(cv, px, rgb, 1.0)
        for q, _a in px:
            cv.buf[q * 3:q * 3 + 3] = bytes(rgb)
        g.painted(cv)
    for q, c in (later or {}).items():
        cv.buf[q * 3:q * 3 + 3] = bytes(c)
    return g.judge(bytes(cv.buf))


def test_a_letter_is_hidden_only_when_all_of_its_ink_is_on_its_own_color():
    black, white = (0, 0, 0), (255, 255, 255)
    a = [(q, 255) for q in (0, 1, 2)]
    b = [(q, 255) for q in (4, 5, 6)]
    n, hidden, dim, _low, over = _probe({q: black for q in (0, 1, 2, 4)}, [(a, black), (b, black)])
    assert n == 2 and hidden == [[0, 1, 2]] and dim == [] and over == []  # b は 3 つのうち 1 つだけ：半分より多くはない
    n, hidden, dim, low, _over = _probe({q: black for q in (0, 1, 4, 5)}, [(a, black), (b, black)])
    assert hidden == [] and dim == [[0, 1, 2], [4, 5, 6]] and low == 1.0  # 墨の 3 つのうち 2 つが、同じ色の上
    # 同じ字を 2 回重ねて描く（太字）：2 回目は、1 回目の墨の上でなく、1 回目の前の色と比べる
    n, hidden, dim, _low, over = _probe({}, [(a, black), (a, black)])
    assert n == 2 and hidden == [] and dim == [] and over == []
    # 白い下地に白い字
    assert _probe({}, [(a, white)])[1] == [[0, 1, 2]]


def test_a_letter_painted_over_later():
    black = (0, 0, 0)
    a = [(q, 255) for q in (0, 1, 2)]
    b = [(q, 255) for q in (4, 5, 6)]
    assert _probe({}, [(a, black), (b, black)], {0: (255, 0, 0), 1: (255, 0, 0)})[4] == [[0, 1, 2]]
    assert _probe({}, [(a, black), (b, black)], {0: (255, 0, 0)})[4] == []  # 3 つのうち 1 つ：半分より多くはない
    # 同じ欄の後ろの字が重なる所は、重なったものとしない（終わりに塗った色と比べる）
    assert _probe({}, [(a, black), ([(2, 255), (3, 255)], black)])[4] == []


def test_every_widget_of_a_field_on_the_page_is_looked_at(tmp_path):
    # 同じ欄の部品が 2 つ：2 つ目だけが黒い帯の上（前は、ページごとに、欄の 1 つ目の部品だけを見ていた）
    kid = "<< /Type /Annot /Subtype /Widget /Parent 2 0 R /Rect [{}] /F 4 >>"
    blob = pdfmini.build("0 g 90 590 300 40 re f", "<< >>",
                         ["<< /FT /Tx /T (a) /DA (/Helv 12 Tf 0 g) /Kids [3 0 R 4 0 R] >>", kid.format("100 700 300 716"),
                          kid.format("100 600 300 616")], annots="/Annots [3 0 R 4 0 R]", acroform=ACRO)
    r, _out = _fill(tmp_path, blob, {"a": "AB"})
    assert r["read_back_ok"] is False, r
    d = [x for x in r["read_back_differences"] if x["id"] == "a"]
    assert len(d) == 1 and "画素が 1 つも変わりません" in d[0]["read_back"]
    assert 76.0 <= _mm(d[0]["read_back"])[1] <= 82.0  # 2 つ目の部品の位置（ページの上から 226〜242pt = 79.7〜85.4 mm）


def test_the_check_mark_of_a_field_without_appearance_uses_the_color(tmp_path):
    box = "<< /Type /Annot /Subtype /Widget /FT /Btn /T (c) /Rect [100 700 112 712] /F 4 /DA (/ZaDb 0 Tf 0 0 0.5 rg) >>"
    blob = pdfmini.build("", "<< >>", [box], annots="/Annots [2 0 R]", acroform="/AcroForm << /Fields [2 0 R] /NeedAppearances true >>")
    src = tmp_path / "a.pdf"
    src.write_bytes(blob)
    fid = next(f["id"] for f in api.fields(str(src)))
    r = api.fill(str(src), {fid: True}, str(tmp_path / "b.pdf"), overwrite=True)
    assert r["proof_ok"], r
    doc = PdfDocument((tmp_path / "b.pdf").read_bytes())
    wd = doc.resolve(read_fields(doc)[0].widgets[0].ref)
    on = doc.resolve(doc.resolve(wd["AP"])["N"])
    ap = doc.decode(doc.resolve(next(v for k, v in on.items() if k != "Off")))
    assert b"BT\n0 0 0.5 rg\n" in ap


# ====================================================================== このパソコンに書体がないとき
NAME = "BT /F1 10 Tf 105 740 Td (Name of the applicant) Tj ET "


def _no_system_fonts(monkeypatch, own: bool = False):
    """このパソコンに代わりの書体がない（埋め込みでない書体の字は、灰色の四角になる）。own：書いた欄の字の形も分からない。"""
    from todoku.pdf.render import fontmap

    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    if own:
        monkeypatch.setattr(fontmap.Glyphs, "lookup", lambda self, code, text: None)


def _form_with_text(content: str = ""):
    return pdfmini.build(NAME + content, FONT, [_widget()], annots="/Annots [2 0 R]", acroform=ACRO)


def test_letters_are_checked_when_other_text_of_the_page_has_no_font(tmp_path, monkeypatch):
    """ページのほかの字の書体がなくても、書いた欄の字（形を描ける字）は、字ごとに見る（前は、ページに描けない書体が 1 つでもあると、
    字ごとには見なかった：書体のない機械では、黒い帯の上の字を、何も伝えずに書いていた）。"""
    _no_system_fonts(monkeypatch)
    r, _out = _fill(tmp_path, _form_with_text("0 g 90 690 40 40 re f"), {"a": LONG})
    assert r["read_back_ok"] is False and r["proof_ok"]
    d = next(x for x in r["read_back_differences"] if x["id"] == "a")
    assert "描いた字 26 個のうち" in d["read_back"] and "下地と同じ色" in d["read_back"]
    assert 3 <= int(re.search(r"のうち (\d+) 個は、描いても見えません", d["read_back"])[1]) <= 7
    r, _out = _fill(tmp_path, _form_with_text("0 0 0.6 rg 90 690 40 40 re f"), {"a": LONG})
    assert r["read_back_ok"] and len(_told(r)) == 1 and "見えにくいかもしれません" in _told(r)[0]
    r, _out = _fill(tmp_path, _form_with_text(), {"a": LONG})
    assert r["read_back_ok"] and not _told(r), r


def test_letters_without_a_shape_are_not_judged_by_the_gray_box(tmp_path, monkeypatch):
    """書いた字の形が分からない（描画は灰色の四角で示す）：白い下地の上の字を、四角の灰色の明るさで「見えにくい」と言わない。"""
    _no_system_fonts(monkeypatch, own=True)
    r, _out = _fill(tmp_path, _form_with_text(), {"a": LONG})
    assert r["read_back_ok"] and not _told(r), r


@pytest.mark.parametrize("ground, told", [("0 g", "違い"), ("0 0 0.6 rg", "知らせ")])
def test_letters_without_a_shape_are_judged_by_their_colour(tmp_path, monkeypatch, ground, told):
    """字の形はなくても、字の色は分かる：字の四角の下地がどこも字と同じ色なら見えない字、どこも比 3:1 未満なら見えにくい字。"""
    _no_system_fonts(monkeypatch, own=True)
    r, _out = _fill(tmp_path, _form_with_text(f"{ground} 90 690 40 40 re f"), {"a": LONG})
    note = [w for w in r["warnings"] if "字の形がこのパソコンにない" in w]
    assert len(note) == 1 and note[0].startswith("a：描いた字 26 個のうち 26 個は、") and "下地の色で確かめました" in note[0]
    if told == "違い":
        d = next(x for x in r["read_back_differences"] if x["id"] == "a")
        assert r["read_back_ok"] is False and "下地と同じ色" in d["read_back"]
        n = int(re.search(r"のうち (\d+) 個は、描いても見えません", d["read_back"])[1])
        x0, _y0, x1, _y1 = _mm(d["read_back"])
    else:
        assert r["read_back_ok"], r
        w = next(x for x in _told(r) if "見えにくいかもしれません" in x)
        n = int(re.search(r"のうち (\d+) 個は、字と下地", w)[1])
        x0, _y0, x1, _y1 = _mm(w)
    assert 2 <= n <= 7 and 35.0 <= x0 < x1 <= 46.5  # 四角が帯に入りきる字だけ


def test_letters_without_a_shape_under_a_later_annotation_are_told(tmp_path, monkeypatch):
    _no_system_fonts(monkeypatch, own=True)
    ops = "1 0 0 rg 0 0 105 30 re f"
    ap = f"<< /Type /XObject /Subtype /Form /BBox [0 0 105 30] /Length {len(ops)} >>\nstream\n{ops}\nendstream"
    stamp = "<< /Type /Annot /Subtype /Stamp /Rect [95 695 200 725] /F 4 /AP << /N 4 0 R >> >>"
    blob = pdfmini.build(NAME, FONT, [_widget(), stamp, ap], annots="/Annots [2 0 R 3 0 R]", acroform=ACRO)
    r, _out = _fill(tmp_path, blob, {"a": LONG})
    told = [w for w in r["warnings"] if "重なっています" in w]
    assert r["read_back_ok"] and len(told) == 1 and told[0].startswith("a：描いた字 26 個のうち"), r
