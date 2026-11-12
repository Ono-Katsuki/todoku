"""PDF の塗り・画像の下の字：字を描いた後から、透けない四角を重ねて隠した字を、見える内容として読まない。

- 字の四角（送りの幅 × 書体の高さ）の全部が、後から描いた透けない塗りの四角（軸にそろった四角形の塗り）の中にある字は、
  読み順・入力欄に入れず、ページの「塗りの下にあって見えない字」の節に、位置（mm）と一緒に分けて示す（字は 1 字も落とさない）
- 透けない＝不透明度 1・ふつうの重ね方・マスクなし・重ね刷りなし・模様でない色・四角の切り抜きの中・レイヤーの外。
  どれかが違う塗り、曲線・斜めの形の塗り、同じ道のほかの部分と重なる四角、線だけの四角は、隠すと決めない（見える字のまま）
- 四角にかかるだけの字は、見える字のまま
- 同じ字を同じ所に、見える字として描き直してあるときは、重ね書きと同じに 1 字にまとめる（隠れた方を、節に出さない）
- 後から描いた透けない画像の下の字は、読み順に残し、見えない字として「画像の下に描いた字」（読む B）で位置を伝える
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from test_robustness import _page_objs, _pdf, _stream  # noqa: E402

from todoku import api, lazy  # noqa: E402
from todoku.pdf.content import page_content  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402

NAME = b"BT /F1 12 Tf 72 720 Td (Name:) Tj ET\n"
HIDDEN = b"BT /F1 12 Tf 150 720 Td (Hidden) Tj ET\n"
BOX = b"140 700 100 50 re f\n"  # 「Hidden」の全部を中に入れる四角（Name: は外）
TITLE = "### 塗りの下にあって見えない字"
UNDER = "画像の下に描いた字（画面には出ない字）"
MM = 72 / 25.4
IMG = {6: _stream(b"/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", b"\x80")}
DRAW = b"q 100 0 0 50 140 700 cm /Im1 Do Q\n"  # 画像を、BOX と同じ四角に描く
XIM = b"/XObject << /Im1 6 0 R >>"


def _write(tmp_path, body: bytes, extra: bytes = b"", more: dict | None = None, name: str = "a.pdf") -> str:
    objs = _page_objs(body, extra)
    objs.update(more or {})
    p = tmp_path / name
    p.write_bytes(_pdf(objs))
    api._WINDOWS.clear()
    lazy.clear()
    return str(p)


def _page(path: str) -> tuple[dict, str]:
    _kind, win, parts = api.window(path)
    return win["pages"][0], "\n".join(md for _n, md in parts)


def _content(path: str):
    doc = PdfDocument(Path(path).read_bytes(), path)
    return page_content(doc, doc.pages()[0])


def _audit_ok(path: str) -> bool:
    return all(c["ok"] for c in api.audit(path)["checks"] if c["name"].startswith("読み落とし"))


def _reads_only_name(pg: dict) -> bool:
    """読み順にあるのは「Name:」だけ（塗りの四角は、空のマスとして欄になることがある：欄の印は字ではない）。"""
    import re
    return [t for t in (re.sub(r"［[^］]*］|[\s｜]", "", ln) for ln in pg["reading"]) if t] == ["Name:"]


def _visible(tmp_path, body: bytes, extra: bytes = b"", more: dict | None = None) -> None:
    """「Hidden」が、見える字のまま読み順にある（隠すと決めていない）。"""
    p = _write(tmp_path, body, extra, more)
    pg, md = _page(p)
    assert "covered" not in pg and TITLE not in md
    assert any("Hidden" in ln for ln in pg["reading"])
    assert not any(g.invisible or g.under for g in _content(p).glyphs if 149 < g.x0 < 190)
    assert not any(f["title"] == UNDER for f in pg.get("read_flags", []))
    assert _audit_ok(p)


# ---------------------------------------------------------------------- 塗り
def test_text_under_a_later_opaque_fill_is_told_apart_from_the_reading(tmp_path):
    p = _write(tmp_path, NAME + HIDDEN + BOX)
    pg, md = _page(p)
    assert _reads_only_name(pg) and pg["chars"] == len("Name:")
    (c,) = pg["covered"]
    assert c["text"] == "Hidden"
    assert abs(c["left_mm"] - 150 / MM) < 0.2 and 36 < c["top_mm"] < 43
    assert TITLE + "（6 字）" in md and f"- （左 {c['left_mm']}・上 {c['top_mm']} mm）Hidden" in md
    assert md.index("Hidden") > md.index("### 読み順")
    assert not any("Hidden" in str(f) for f in pg["fields"])  # 見出し・近くの字にもならない
    pc = _content(p)
    assert "".join(g.text for g in pc.covered) == "Hidden" and "".join(g.text for g in pc.glyphs) == "Name:"
    assert _audit_ok(p)


def test_the_colour_of_the_fill_does_not_matter(tmp_path):
    for colour in (b"1 g\n", b"0 g\n", b"1 0 0 rg\n", b"0 0 0 0 k\n"):
        p = _write(tmp_path, NAME + HIDDEN + colour + BOX)
        assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]


@pytest.mark.parametrize("paint", [b"f", b"F", b"f*", b"B", b"B*", b"b", b"b*"])
def test_every_filling_operator_covers(tmp_path, paint):
    p = _write(tmp_path, NAME + HIDDEN + b"140 700 100 50 re " + paint + b"\n")
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]
    assert _audit_ok(p)


def test_a_rectangle_drawn_with_lines_covers(tmp_path):
    p = _write(tmp_path, NAME + HIDDEN + b"140 700 m 240 700 l 240 750 l 140 750 l h f\n")
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]


def test_a_fill_drawn_before_the_text_hides_nothing(tmp_path):
    _visible(tmp_path, NAME + BOX + HIDDEN)


def test_text_drawn_over_the_fill_is_read_and_the_text_under_it_is_told_apart(tmp_path):
    p = _write(tmp_path, NAME + HIDDEN + b"1 g\n" + BOX + b"0 g BT /F1 12 Tf 150 730 Td (Example) Tj ET\n")
    pg, md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["Hidden"]
    assert any("Example" in ln for ln in pg["reading"]) and not any("Hidden" in ln for ln in pg["reading"])
    assert _audit_ok(p)


def test_only_the_glyphs_wholly_inside_the_fill_are_hidden(tmp_path):
    # 四角の右の辺が「Hidden」の途中を通る：中に全部ある字だけが隠れ、辺にかかる字と外の字は見える字のまま
    pc0 = _content(_write(tmp_path, HIDDEN))
    d1, d2 = [g for g in pc0.glyphs if g.text == "d"]
    edge = (d1.x0 + d1.x1) / 2  # 1 つ目の「d」の中央
    p = _write(tmp_path, NAME + HIDDEN + b"140 700 %.2f 50 re f\n" % (edge - 140))
    pg, _md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["Hi"]
    assert pg["chars"] == len("Name:") + len("dden")  # 辺にかかる「d」と、外の「den」は見える字
    assert _audit_ok(p)


def test_a_fill_shorter_than_the_glyph_box_hides_nothing(tmp_path):
    g = _content(_write(tmp_path, HIDDEN)).glyphs[0]
    h = (g.y1 - g.y0) * 0.8
    _visible(tmp_path, NAME + HIDDEN + b"140 %.2f 100 %.2f re f\n" % (g.y0, h))


@pytest.mark.parametrize("extra, before", [
    (b"/ExtGState << /G1 << /ca 0.5 >> >>", b"/G1 gs\n"),  # 半分透ける塗り
    (b"/ExtGState << /G1 << /BM /Multiply >> >>", b"/G1 gs\n"),  # 下の色と掛け合わせる重ね方
    (b"/ExtGState << /G1 << /SMask << /Type /Mask /S /Luminosity /G 7 0 R >> >> >>", b"/G1 gs\n"),  # マスク
    (b"/ExtGState << /G1 << /op true /OPM 1 >> >>", b"/G1 gs\n"),  # 重ね刷り
    (b"/Pattern << /P1 8 0 R >>", b"/Pattern cs /P1 scn\n"),  # 模様の塗り
    (b"/Properties << /L1 9 0 R >>", b"/OC /L1 BDC\n"),  # 既定では見えないレイヤーの中の塗り
])
def test_a_fill_that_may_show_what_is_under_it_hides_nothing(tmp_path, extra, before):
    more = {
        7: _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842] /Group << /S /Transparency /CS /DeviceGray >>", b"0.5 g 0 0 595 842 re f"),
        8: _stream(b"/Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 4 4] /XStep 4 /YStep 4 /Resources << >>",
                   b"0 0 1 1 re f"),
        9: b"<< /Type /OCG /Name (L) >>",
        1: b"<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [9 0 R] /D << /OFF [9 0 R] >> >> >>",
    }
    after = b"EMC\n" if before.startswith(b"/OC") else b""
    _visible(tmp_path, NAME + HIDDEN + b"q\n" + before + BOX + after + b"Q\n", extra, more)


def test_a_fill_in_a_layer_that_is_shown_by_default_covers(tmp_path):
    more = {9: b"<< /Type /OCG /Name (L) >>", 1: b"<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [9 0 R] /D << /ON [9 0 R] >> >> >>"}
    p = _write(tmp_path, NAME + HIDDEN + b"/OC /L1 BDC\n" + BOX + b"EMC\n", b"/Properties << /L1 9 0 R >>", more)
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]


def test_the_graphics_state_is_restored_after_a_transparent_fill(tmp_path):
    p = _write(tmp_path, NAME + HIDDEN + b"q /G1 gs 0 0 10 10 re f Q\n" + BOX, b"/ExtGState << /G1 << /ca 0.5 >> >>")
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]


@pytest.mark.parametrize("shape", [
    b"140 700 m 240 700 l 240 750 l h f\n",  # 三角
    b"140 700 m 240 705 l 240 755 l 140 750 l h f\n",  # 斜めの四角
    b"140 725 m 140 760 240 760 240 725 c 240 690 140 690 140 725 c h f\n",  # 曲線の形
    b"140 700 100 50 re S\n",  # 線だけの四角
    b"140 700 100 50 re n\n",  # 描かない道
    b"140 700 100 50 re 145 705 90 40 re f*\n",  # 同じ道の中の四角が重なる（偶奇の決まりで、内側は穴）
    b"140 700 100 50 re 145 705 90 40 re f\n",
    b"140 700 100 50 re W n\n",  # 切り抜きだけ
    b"/Sh1 sh\n",  # グラデーション
])
def test_a_shape_that_is_not_one_plain_rectangle_hides_nothing(tmp_path, shape):
    more = {7: b"<< /ShadingType 2 /ColorSpace /DeviceGray /Coords [0 0 595 0] /Function << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [1] /N 1 >> >>"}
    _visible(tmp_path, NAME + HIDDEN + shape, b"/Shading << /Sh1 7 0 R >>", more)


def test_a_fill_cut_by_a_rectangular_clip_covers_only_inside_the_clip(tmp_path):
    # 切り抜きが「Hidden」の左（x < 145）だけを残す：塗りは字に届かない
    _visible(tmp_path, NAME + HIDDEN + b"q 0 0 145 842 re W n\n" + BOX + b"Q\n")
    # 切り抜きが字の全部を中に入れる：塗りは字を隠す
    p = _write(tmp_path, NAME + HIDDEN + b"q 145 690 90 70 re W n\n" + BOX + b"Q\n")
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]


def test_a_fill_inside_a_clip_that_is_not_a_rectangle_hides_nothing(tmp_path):
    _visible(tmp_path, NAME + HIDDEN + b"q 100 650 m 300 650 l 200 800 l h W n\n" + BOX + b"Q\n")
    # 字の形の切り抜き（Tr 7）の中の塗り
    _visible(tmp_path, NAME + HIDDEN + b"q BT /F1 200 Tf 7 Tr 100 650 Td (O) Tj ET\n" + BOX + b"Q\n")


def test_a_fill_inside_a_form_covers_and_a_form_drawn_transparently_does_not(tmp_path):
    form = {7: _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842]", BOX)}
    p = _write(tmp_path, NAME + HIDDEN + b"/X1 Do\n", b"/XObject << /X1 7 0 R >>", form)
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hidden"]
    assert _audit_ok(p)
    # 部品を半分透ける描き方で描く：中の塗りも透ける
    _visible(tmp_path, NAME + HIDDEN + b"q /G1 gs /X1 Do Q\n", b"/XObject << /X1 7 0 R >> /ExtGState << /G1 << /ca 0.5 >> >>", form)


def test_text_inside_a_form_is_hidden_by_a_later_fill_of_the_page(tmp_path):
    form = {7: _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >>", HIDDEN)}
    p = _write(tmp_path, NAME + b"/X1 Do\n" + BOX, b"/XObject << /X1 7 0 R >>", form)
    pg, _md = _page(p)
    assert _reads_only_name(pg) and [c["text"] for c in pg["covered"]] == ["Hidden"]
    assert _audit_ok(p)


def test_a_small_fill_hides_nothing(tmp_path):
    # 字を通る細い塗り 1 本（罫線として描いた四角）は、字の四角の全部を覆わないので、字を隠さない
    _visible(tmp_path, NAME + HIDDEN + b"140 722 100 2 re f\n")


def test_invisible_text_under_a_fill_stays_as_it_is(tmp_path):
    p = _write(tmp_path, NAME + b"BT /F1 12 Tf 3 Tr 150 720 Td (Hidden) Tj ET\n" + BOX)
    pg, md = _page(p)
    assert "covered" not in pg and TITLE not in md
    pc = _content(p)
    assert not pc.covered and [g.text for g in pc.glyphs if g.invisible] == list("Hidden") and not any(g.under for g in pc.glyphs)
    assert _audit_ok(p)


def test_text_replaced_by_one_string_is_hidden_as_one(tmp_path):
    # ActualText でまとめた字：まとめた後の字が、四角の下にある
    body = NAME + b"/Span << /ActualText (Zed) >> BDC\n" + HIDDEN + b"EMC\n" + BOX
    p = _write(tmp_path, body)
    pg, _md = _page(p)
    assert _reads_only_name(pg) and [c["text"] for c in pg["covered"]] == ["Zed"]
    # まとめた字より前の四角は、後の字を隠さない
    p = _write(tmp_path, NAME + BOX + b"/Span << /ActualText (Zed) >> BDC\n" + HIDDEN + b"EMC\n")
    pg, _md = _page(p)
    assert "covered" not in pg and any("Zed" in ln for ln in pg["reading"])


def test_a_page_whose_text_is_all_under_fills_says_where_the_text_is(tmp_path):
    p = _write(tmp_path, HIDDEN + BOX)
    pg, md = _page(p)
    assert pg["chars"] == 0 and [c["text"] for c in pg["covered"]] == ["Hidden"]
    assert "塗りの下にあって見えない字は、下の節にあります" in md and TITLE + "（6 字）" in md
    assert _audit_ok(p)


# ---------------------------------------------------------------------- 描き直した字
def test_the_same_text_redrawn_at_the_same_place_is_read_once(tmp_path):
    p = _write(tmp_path, NAME + HIDDEN + b"1 g\n" + BOX + b"0 g\n" + HIDDEN)
    pg, md = _page(p)
    assert "covered" not in pg and TITLE not in md
    assert sum(ln.count("Hidden") for ln in pg["reading"]) == 1
    assert _audit_ok(p)


def test_the_same_text_redrawn_at_another_place_is_told_as_both(tmp_path):
    again = b"BT /F1 12 Tf 150 735 Td (Hidden) Tj ET\n"  # 1 行ぶん上
    p = _write(tmp_path, NAME + HIDDEN + b"1 g\n" + BOX + b"0 g\n" + again)
    pg, _md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["Hidden"]
    assert sum(ln.count("Hidden") for ln in pg["reading"]) == 1
    assert _audit_ok(p)


def test_text_overprinted_and_then_covered_is_told_once(tmp_path):
    # 太字に見せる重ね書き（同じ字を 2 回）を、後から塗りで隠す：節には 1 回だけ
    p = _write(tmp_path, NAME + HIDDEN + HIDDEN + BOX)
    pg, _md = _page(p)
    assert [c["text"] for c in pg["covered"]] == ["Hidden"]
    assert _audit_ok(p)


# ---------------------------------------------------------------------- 画像
def test_text_under_a_later_opaque_image_stays_in_the_reading_and_is_flagged(tmp_path):
    p = _write(tmp_path, NAME + HIDDEN + DRAW, XIM, IMG)
    pg, md = _page(p)
    assert "covered" not in pg and TITLE not in md
    assert any("Hidden" in ln for ln in pg["reading"])
    pc = _content(p)
    assert [g.text for g in pc.glyphs if g.under] == list("Hidden") and all(g.invisible for g in pc.glyphs if g.under)
    assert not any(g.under or g.invisible for g in pc.glyphs if g.x0 < 140)
    (f,) = [f for f in pg["read_flags"] if f["title"] == UNDER]
    assert f["grade"] == "B" and len(f["where"]) == 1 and "「Hidden」" in f["where"][0]
    assert _audit_ok(p)


def test_a_page_whose_text_is_all_under_an_image_says_so_in_one_line(tmp_path):
    p = _write(tmp_path, HIDDEN + DRAW, XIM, IMG)
    pg, _md = _page(p)
    (f,) = [f for f in pg["read_flags"] if f["title"] == UNDER]
    assert f["where"] == ["1ページ（ページの字のすべて：6 字）"]


def test_an_image_drawn_before_the_text_hides_nothing(tmp_path):
    _visible(tmp_path, NAME + DRAW + HIDDEN, XIM, IMG)


@pytest.mark.parametrize("image, draw", [
    (_stream(b"/Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1", b"\x00"), DRAW),  # 型抜き
    (_stream(b"/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /SMask 7 0 R", b"\x80"), DRAW),
    (_stream(b"/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /Mask [0 0]", b"\x80"), DRAW),
    (IMG[6], b"q 100 20 -20 50 150 690 cm /Im1 Do Q\n"),  # 斜めに置いた画像
    (IMG[6], b"q /G1 gs 100 0 0 50 140 700 cm /Im1 Do Q\n"),  # 半分透ける描き方
])
def test_an_image_that_may_show_what_is_under_it_hides_nothing(tmp_path, image, draw):
    more = {6: image, 7: _stream(b"/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", b"\x80")}
    _visible(tmp_path, NAME + HIDDEN + draw, XIM + b" /ExtGState << /G1 << /ca 0.5 >> >>", more)


def test_an_inline_image_covers_like_an_image(tmp_path):
    inline = b"q 100 0 0 50 140 700 cm BI /W 1 /H 1 /CS /G /BPC 8 ID \x80 EI Q\n"
    p = _write(tmp_path, NAME + HIDDEN + inline)
    pg, _md = _page(p)
    assert [g.text for g in _content(p).glyphs if g.under] == list("Hidden")
    assert any(f["title"] == UNDER for f in pg["read_flags"])


def test_a_fill_and_then_an_image_over_the_same_text_is_told_as_under_the_fill(tmp_path):
    # 塗りで隠した後に、画像を重ねる：いちばん後の覆いは画像だが、塗りが確かに隠している字は、読み順に戻さない
    p = _write(tmp_path, NAME + HIDDEN + BOX + DRAW, XIM, IMG)
    pg, _md = _page(p)
    assert _reads_only_name(pg) and [c["text"] for c in pg["covered"]] == ["Hidden"]
    assert not any(f["title"] == UNDER for f in pg.get("read_flags", []))
    assert _audit_ok(p)


# ---------------------------------------------------------------------- 空白
def test_a_space_under_a_fill_goes_with_the_text_of_its_text_object(tmp_path):
    # 空白には隠れる墨がない：同じ文字の組（BT〜ET）のすぐ前の字が見える字なら、四角の中の空白も見える字のまま
    tail = b"BT /F1 12 Tf 72 720 Td (Name:) Tj 78 0 Td ( ) Tj ET\n"  # 「Name:」の組の終わりの空白が、四角の中にある
    p = _write(tmp_path, tail + BOX)
    pc = _content(p)
    assert not pc.covered and "".join(g.text for g in pc.glyphs) == "Name: "
    # 隠れた字の間・後ろの空白は、隠れた字と一緒
    p = _write(tmp_path, NAME + b"BT /F1 12 Tf 150 720 Td (Hid den ) Tj ET\n" + BOX)
    pc = _content(p)
    assert "".join(g.text for g in pc.covered) == "Hid den " and "".join(g.text for g in pc.glyphs) == "Name:"
    assert [c["text"] for c in _page(p)[0]["covered"]] == ["Hid den"]
    # 組の初めの空白は、次の字と同じ側
    p = _write(tmp_path, NAME + b"BT /F1 12 Tf 150 720 Td ( Hidden) Tj ET\n" + BOX)
    assert "".join(g.text for g in _content(p).covered) == " Hidden"
    assert _audit_ok(p)


def test_a_text_object_of_spaces_only_under_a_fill_stays_as_it_is(tmp_path):
    p = _write(tmp_path, NAME + b"BT /F1 12 Tf 150 720 Td (   ) Tj ET\n" + BOX)
    pc = _content(p)
    assert not pc.covered and "".join(g.text for g in pc.glyphs) == "Name:   "
    # 画像の下でも同じ
    p = _write(tmp_path, NAME + b"BT /F1 12 Tf 150 720 Td (   ) Tj ET\n" + DRAW, XIM, IMG)
    assert not any(g.under or g.invisible for g in _content(p).glyphs)
