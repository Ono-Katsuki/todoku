"""PDF の字の色：字を描いた色（塗りの色。輪郭だけを描く字は線の色）を、ほかの形式と同じ印で、事実としてだけ伝える。

- 「ふつう」は、そのページで字数のいちばん多い色（同じ字数なら黒）。ふつうの色の字には印を付けない。黒でなければ、
  ページの normal_color と、読み順の頭の 1 行で伝える
- 行の字がみな同じ色なら、行の前に「〔字の色 #RRGGBB〕」。一部だけなら、その字のすぐ後ろに「〔うち字の色 #RRGGBB：「字」〕」
- 色は、選んだ色空間（cs・CS）で読む：Separation・DeviceN は色の変換の関数、Indexed は色の表、ICCBased は成分の数。
  色空間の初めの色（Separation は濃さ 1）も読む。模様（Pattern）で描いた字は「字の色（模様）」
- 見えない字（Tr 3）の色は見ない。重ねて描いた同じ字は、後から描いた字（上の字）の色
- 記法の説明は、字の色の印か、黒でないふつうの色のあるときだけ。字そのものは変えない
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

import pdfmini  # noqa: E402

from todoku import api, lazy  # noqa: E402
from todoku.pdf import strike  # noqa: E402
from todoku.pdf.content import page_content  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402

KEY = "〔字の色 #FF0000〕"
BODY = ("0 g", "申請者の氏名と住所を書いてください", 100, 700)  # 黒い字 17 字：ページのふつうの色を黒にする


def _pdf(runs, cs: str = "", objs=(), extra: str = "") -> bytes:
    """runs：(色の命令, 字, x, y)。cs：資源に足す辞書の項目（/ColorSpace など）。objs：4 番目からのオブジェクト。"""
    chars = sorted({ch for _o, t, _x, _y in runs for ch in t})
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    w = " ".join(f"{code[ch]} [500]" for ch in chars if ord(ch) < 0x2E80)
    parts = [f"{ops} BT /F1 12 Tf {x} {y} Td <{''.join(f'{code[ch]:04X}' for ch in t)}> Tj ET" for ops, t, x, y in runs]
    font = [pdfmini.type0("TestFont", 3, f"/W [{w}]" if w else ""), pdfmini.stream(pdfmini.tounicode({c: ch for ch, c in code.items()}))]
    return pdfmini.build("\n".join(parts) + "\n" + extra, f"<< /Font << /F1 2 0 R >> {cs} >>", font + list(objs))


def _open(tmp_path, blob: bytes):
    p = tmp_path / "a.pdf"
    p.write_bytes(blob)
    api._WINDOWS.clear()
    lazy.clear()
    _kind, win, parts = api.window(str(p))
    return win, win["pages"][0], "\n".join(md for _n, md in parts)


def _reading(tmp_path, runs, **kw) -> tuple[list[str], dict, dict]:
    win, pg, _md = _open(tmp_path, _pdf(runs, **kw))
    return pg["reading"], pg, win


# ---- 印


def test_whole_line_in_another_colour(tmp_path):
    lines, pg, win = _reading(tmp_path, [BODY, ("1 0 0 rg", "記入例", 100, 680)])
    assert lines[0] == "申請者の氏名と住所を書いてください"
    assert lines[1].startswith("〔字の色 #FF0000〕記入例")
    assert pg["color_marks"] == 1 and "normal_color" not in pg
    assert KEY in win["notation"]


def test_part_of_a_line(tmp_path):
    lines, pg, _win = _reading(tmp_path, [BODY, ("0 g", "氏名", 100, 680), ("1 0 0 rg", "山田太郎", 130, 680), ("0 g", "印", 200, 680)])
    assert "山田太郎〔うち字の色 #FF0000：「山田太郎」〕" in lines[1]
    assert lines[1].startswith("氏名") and "〔字の色 #FF0000〕" not in lines[1]
    assert lines[1].count("字の色") == 1 and pg["color_marks"] == 1


def test_two_colours_in_a_line(tmp_path):
    lines, pg, _win = _reading(tmp_path, [BODY, ("1 0 0 rg", "赤い字", 100, 680), ("0 0 1 rg", "青い字", 136, 680)])
    assert "赤い字〔うち字の色 #FF0000：「赤い字」〕青い字〔うち字の色 #0000FF：「青い字」〕" in lines[1]
    assert pg["color_marks"] == 2


def test_same_colour_across_a_gap_is_one_run(tmp_path):
    lines, _pg, _win = _reading(tmp_path, [BODY, ("0 g", "氏名", 100, 680), ("1 0 0 rg", "山田", 130, 680), ("1 0 0 rg", "太郎", 166, 680)])
    assert lines[1].count("〔うち字の色 #FF0000：") == 1
    assert "〔うち字の色 #FF0000：「山田 太郎」〕" in lines[1]


def test_black_only_page_has_no_marks(tmp_path):
    lines, pg, win = _reading(tmp_path, [BODY, ("0 0 0 rg", "二行目", 100, 680), ("0 0 0 1 k", "三行目", 100, 660)])
    assert not any("字の色" in ln for ln in lines)
    assert "color_marks" not in pg and "normal_color" not in pg
    assert KEY not in win["notation"]


def test_page_normal_colour_is_the_most_common(tmp_path):
    win, pg, md = _open(tmp_path, _pdf([("1 0 0 rg", "赤い字ばかりのページです", 100, 700), ("1 0 0 rg", "二行目も赤い", 100, 680),
                                        ("0 g", "黒い字", 100, 660)]))
    assert pg["normal_color"] == "#FF0000"
    assert not pg["reading"][0].startswith("〔") and not pg["reading"][1].startswith("〔")
    assert pg["reading"][2].startswith("〔字の色 #000000〕黒い字")
    assert "このページのふつうの字の色は #FF0000" in md
    assert KEY in win["notation"]


def test_same_count_prefers_black(tmp_path):
    lines, pg, _win = _reading(tmp_path, [("0 g", "黒い字", 100, 700), ("1 0 0 rg", "赤い字", 100, 680)])
    assert "normal_color" not in pg
    assert lines[1].startswith("〔字の色 #FF0000〕")


def test_page_all_in_one_colour_tells_only_the_normal_colour(tmp_path):
    lines, pg, win = _reading(tmp_path, [("0 0 1 rg", "青い字のページ", 100, 700)])
    assert pg["normal_color"] == "#0000FF" and "color_marks" not in pg
    assert lines[0].startswith("青い字のページ")
    assert KEY in win["notation"]


# ---- 色空間


def test_cmyk_and_gray(tmp_path):
    lines, _pg, _win = _reading(tmp_path, [BODY, ("0 1 1 0 k", "赤い字", 100, 680), ("0.5 g", "灰色", 100, 660)])
    assert lines[1].startswith("〔字の色 #FF0000〕")
    assert lines[2].startswith("〔字の色 #808080〕")


SEP = "/ColorSpace << /CS0 [/Separation /Spot /DeviceCMYK << /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 0.85 0 0] /N 1 >>] >>"


def test_separation_tint_goes_through_the_tint_function(tmp_path):
    """濃さ 1 は、特色のいちばん濃い色（前は、数 1 つを灰色として読み、白になっていた）。"""
    lines, _pg, _win = _reading(tmp_path, [BODY, ("/CS0 cs 1 scn", "特色の字", 100, 680), ("/CS0 cs 0.5 scn", "半分の濃さ", 100, 660)], cs=SEP)
    assert lines[1].startswith("〔字の色 #0026FF〕")
    assert lines[2].startswith("〔字の色 #8093FF〕")


def test_initial_colour_of_the_selected_space(tmp_path):
    """cs だけで色を言わないとき：Separation の初めの色は濃さ 1（黒ではない）。"""
    lines, _pg, _win = _reading(tmp_path, [BODY, ("/CS0 cs", "初めの色", 100, 680)], cs=SEP)
    assert lines[1].startswith("〔字の色 #0026FF〕")


def test_separation_line_is_not_white(tmp_path):
    """特色の濃さ 1 で描いた線は、白い線ではない：字の上にあれば取り消し線。"""
    win, pg, _md = _open(tmp_path, _pdf([("0 g", "申請者氏名", 100, 700)], cs=SEP, extra="/CS0 CS 1 SCN 0.5 w 100 704.5 m 160 704.5 l S"))
    assert [s["text"] for s in pg.get("struck", [])] == ["申請者氏名"]


def test_indexed(tmp_path):
    cs = "/ColorSpace << /CS0 [/Indexed /DeviceRGB 2 <000000 00FF00 0000FF>] >>"
    lines, _pg, _win = _reading(tmp_path, [BODY, ("/CS0 cs 1 sc", "緑の字", 100, 680), ("/CS0 cs 2 sc", "青い字", 100, 660)], cs=cs)
    assert lines[1].startswith("〔字の色 #00FF00〕")
    assert lines[2].startswith("〔字の色 #0000FF〕")


def test_iccbased_by_number_of_components(tmp_path):
    cs = "/ColorSpace << /CS0 [/ICCBased 4 0 R] /CS1 [/ICCBased 5 0 R] >>"
    objs = [pdfmini.stream(b"", "/N 3"), pdfmini.stream(b"", "/N 1")]
    lines, _pg, _win = _reading(tmp_path, [BODY, ("/CS0 cs 1 0 0 scn", "赤い字", 100, 680), ("/CS1 cs 0.5 scn", "灰色", 100, 660)],
                                cs=cs, objs=objs)
    assert lines[1].startswith("〔字の色 #FF0000〕")
    assert lines[2].startswith("〔字の色 #808080〕")


def test_devicen(tmp_path):
    fn = "<< /FunctionType 4 /Domain [0 1 0 1] /Range [0 1 0 1 0 1 0 1] /Length 22 >>\nstream\n{ 0 exch 0 exch 0 exch }\nendstream"
    cs = "/ColorSpace << /CS0 [/DeviceN [/Cyan /Magenta] /DeviceCMYK 4 0 R] >>"
    blob = _pdf([BODY, ("/CS0 cs 0 1 scn", "紅の字", 100, 680)], cs=cs, objs=[fn])
    doc = PdfDocument(blob)
    (g,) = [g for g in page_content(doc, doc.pages()[0]).glyphs if g.text == "紅"]
    assert g.color is not None and len(g.color) == 3 and g.color != (1.0, 1.0, 1.0)


def test_pattern_fill_has_no_single_colour(tmp_path):
    cs = "/ColorSpace << /CS0 [/Pattern] >> /Pattern << /P1 4 0 R >>"
    pat = pdfmini.stream("0 0 2 2 re f", "/PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 4 4] /XStep 4 /YStep 4 /Resources << >>")
    lines, _pg, _win = _reading(tmp_path, [BODY, ("/CS0 cs /P1 scn", "模様の字", 100, 680)], cs=cs, objs=[pat])
    assert lines[1].startswith("〔字の色（模様）〕模様の字")


def test_device_operator_after_a_selected_space(tmp_path):
    """g・rg・k は、色空間も決める：前に選んだ Separation は使わない。"""
    lines, _pg, _win = _reading(tmp_path, [BODY, ("/CS0 cs 1 scn 1 0 0 rg", "赤い字", 100, 680), ("/CS0 cs 1 scn 0 g 1 sc", "白い字", 100, 660)],
                                cs=SEP)
    assert lines[1].startswith("〔字の色 #FF0000〕")
    assert lines[2].startswith("〔字の色 #FFFFFF〕")


def test_colour_is_restored_with_the_graphics_state(tmp_path):
    lines, _pg, _win = _reading(tmp_path, [BODY, ("q /CS0 cs 1 scn Q", "黒い字", 100, 680)], cs=SEP)
    assert lines[1].startswith("黒い字")


# ---- 描き方


def test_outlined_text_uses_the_stroke_colour(tmp_path):
    lines, _pg, _win = _reading(tmp_path, [BODY, ("1 0 0 rg 0 0 1 RG 1 Tr", "輪郭の字", 100, 680), ("0 Tr 1 0 0 rg 0 0 1 RG 2 Tr", "塗りと輪郭", 100, 660)])
    assert lines[1].startswith("〔字の色 #0000FF〕")
    assert lines[2].startswith("〔字の色 #FF0000〕")


def test_white_text_is_told(tmp_path):
    lines, _pg, _win = _reading(tmp_path, [BODY, ("1 g", "白い字", 100, 680)])
    assert lines[1].startswith("〔字の色 #FFFFFF〕白い字")


def test_invisible_text_is_not_counted(tmp_path):
    """見えない字（Tr 3。画像の上に置いた読み取りの字など）の色は、ふつうの色にも印にも入れない。"""
    lines, pg, _win = _reading(tmp_path, [("0 g", "黒い字", 100, 700), ("1 0 0 rg 3 Tr", "見えない字がたくさんあります", 100, 680)])
    assert "normal_color" not in pg and "color_marks" not in pg
    assert not any("字の色" in ln for ln in lines)


def test_overprinted_glyph_takes_the_colour_drawn_last(tmp_path):
    """同じ字を同じ位置に重ねて描いたもの（影・縁取り）は 1 字：上に描いた字の色。"""
    lines, pg, _win = _reading(tmp_path, [BODY, ("1 g", "見出し", 100, 680), ("1 0 0 rg", "見出し", 100, 680)])
    assert lines[1].startswith("〔字の色 #FF0000〕見出し")
    assert lines[1].count("見出し") == 1 and pg["color_marks"] == 1


# ---- ほかの印と一緒に


def test_with_a_strikethrough(tmp_path):
    win, pg, _md = _open(tmp_path, _pdf([BODY, ("0 g", "氏名", 100, 680), ("1 0 0 rg", "山田太郎", 130, 680)],
                                        extra="0 g 0 G 0.5 w 130 684.5 m 178 684.5 l S"))
    line = pg["reading"][1]
    assert "山田太郎〔取り消し線：「山田太郎」〕〔うち字の色 #FF0000：「山田太郎」〕" in line
    assert [s["text"] for s in pg["struck"]] == ["山田太郎"]


def test_text_itself_is_unchanged(tmp_path):
    """読み落としの検査の数える字（窓に出した元の字）に、印の字は入らない。"""
    win, pg, _md = _open(tmp_path, _pdf([BODY, ("0 g", "氏名", 100, 680), ("1 0 0 rg", "山田太郎", 130, 680)]))
    plain = "".join(pg["_plain"])
    assert plain.count("山田太郎") == 1 and "字の色" not in plain


@pytest.mark.parametrize("vertical", [False, True])
def test_line_colours_follow_the_positions(tmp_path, vertical):
    """行の一部の色は、字の位置から決める（縦書きの書体の行も）。"""
    from todoku.pdf import textcolor
    from todoku.pdf.layout import PageLayout

    chars = sorted(set("氏名山田太郎"))
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    hx = lambda t: "".join(f"{code[ch]:04X}" for ch in t)  # noqa: E731
    content = (f"0 g BT /F1 12 Tf 100 700 Td <{hx('氏名')}> Tj ET 1 0 0 rg BT /F1 12 Tf "
               + ("100 676 Td" if vertical else "124 700 Td") + f" <{hx('山田太郎')}> Tj ET")
    font = [pdfmini.type0("TestFont", 3, vertical=vertical), pdfmini.stream(pdfmini.tounicode({c: ch for ch, c in code.items()}))]
    doc = PdfDocument(pdfmini.build(content, "<< /Font << /F1 2 0 R >> >>", font))
    page = doc.pages()[0]
    lay = PageLayout(doc, page, page_content(doc, page), 1)
    (ln,) = [ln for ln in lay.all_lines() if "山田太郎" in ln.text]
    assert ln.vertical == vertical
    per = textcolor.per_text(ln)
    assert [per[k] for k, ch in enumerate(ln.text) if not ch.isspace()] == ["000000"] * 2 + ["FF0000"] * 4
    a = ln.text.index("山")
    assert textcolor.line_marks(ln, "000000") == ("", [(a, a + 4, "FF0000")])
    assert strike.find(lay)[0] == []


def test_marks_do_not_change_how_lines_are_joined():
    """マスの中の行をつなぐとき、印（字の色・取り消し線）は字として数えない：印のあるなしで、空白の入り方は変わらない。"""
    from todoku.pdf.window import _join, _last_char

    assert _join(["申請する", "場合"]) == "申請する場合"
    assert _join(["〔字の色 #FF0000〕申請する", "〔字の色 #FF0000〕場合"]) == "〔字の色 #FF0000〕申請する〔字の色 #FF0000〕場合"
    assert _join(["申請する〔うち字の色 #FF0000：「する」〕", "場合"]) == "申請する〔うち字の色 #FF0000：「する」〕場合"
    assert _join(["申請する〔取り消し線：「する」〕〔うち字の色 #FF0000：「する」〕", "場合"]).endswith("〕場合")
    assert _join(["〔字の色（模様）〕申請", "場合"]) == "〔字の色（模様）〕申請場合"
    assert _join(["name〔うち字の色 #FF0000：「name」〕", "address"]) == "name〔うち字の色 #FF0000：「name」〕 address"
    assert _join(["〔字の色 #FF0000〕name", "〔字の色 #FF0000〕address"]) == "〔字の色 #FF0000〕name 〔字の色 #FF0000〕address"
    assert _last_char("「はい」〔うち字の色 #FF0000：「「はい」」〕") == "」"  # 字の中のかぎ括弧
    assert _last_char("申請する（注）") == "）" and _last_char("「申請」〕") == "〕"  # 印でない字は、そのまま


def test_lines_of_a_cell_join_the_same_with_and_without_colour(tmp_path):
    """同じ字を、黒で描いたページと赤で描いたページ：印を除くと、読み順の字は同じ。"""
    import re

    def reading(colour):
        runs = [BODY, (colour, "申請する", 100, 680), (colour, "場合", 100, 666)]
        _win, pg, _md = _open(tmp_path, _pdf(runs))
        return [re.sub(r"〔(?:うち)?字の色[^〕]*〕", "", ln) for ln in pg["reading"]]

    assert reading("0 g") == reading("1 0 0 rg")
