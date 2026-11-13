"""PDF の、描いても見えない字：下の色と同じ色で塗った字と、塗りの不透明度 0 で描いた字を、見える字として読まない。

前は、色で塗って描いた字は、下の色を見ずに、どれも見える字として読んでいた（白い字は「#FFFFFF」と伝えるだけ）。
いまは、字を描いた時点の、字の四角（送りの幅 × 書体の高さ）の下の色を、描いた順にたどって決める：
- 下がどこも字と同じ色（何も描いていないページの白の上の白い字・同じ色の透けない塗りの四角の上の字）なら、読み順に残したまま、
  見えない字として「下の色と同じ色の字」（読む B）で、行ごとの位置と一緒に伝える
- 字の四角に、ほかの色の塗り・字、線、四角でない塗り、画像、ぼかし、透ける塗りがかかるときは、決めない（見える字のまま）
- 塗りの不透明度（ca）が 0 の字は、「不透明度 0 で描いた字」（読む B）で伝える
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from test_pdf_cover import DRAW, IMG, NAME, XIM, _audit_ok, _content, _page, _write  # noqa: E402
from test_robustness import _stream  # noqa: E402

from todoku import api  # noqa: E402
from todoku.pdf import content  # noqa: E402

SAME = "下の色と同じ色の字（画面には出ない字）"
ALPHA = "不透明度 0 で描いた字（画面には出ない字）"
TEXT = b"BT /F1 12 Tf 150 720 Td (Hidden) Tj ET\n"
WHITE = b"1 g " + TEXT  # 白で塗る「Hidden」
BLACK_BOX = b"0 g 140 700 100 50 re f\n"  # 「Hidden」の全部を中に入れる黒い四角（Name: は外）
WHITE_BOX = b"1 g 140 700 100 50 re f\n"
WIDE_BLACK = b"0 g 100 690 300 80 re f\n"
SHADE = (b"/Shading << /Sh1 << /ShadingType 2 /ColorSpace /DeviceGray /Coords [0 0 595 0] "
         b"/Function << /FunctionType 2 /Domain [0 1] /C0 [0] /C1 [1] /N 1 >> >> >>")
TYPE3 = {7: b"<< /Type /Font /Subtype /Type3 /FontBBox [0 0 1000 1000] /FontMatrix [0.001 0 0 0.001 0 0] /CharProcs << /a 8 0 R >> "
            b"/Encoding << /Type /Encoding /Differences [97 /a] >> /FirstChar 97 /LastChar 97 /Widths [1000] >>",
         8: _stream(b"", b"1000 0 d0 0 0 1000 1000 re f")}


def _faint(path: str) -> list[tuple]:
    """「Hidden」の所の字の（字、見えない字か、見えない理由）。"""
    return [(g.text, g.invisible, g.faint) for g in _content(path).glyphs if 149 < g.x0 < 195 and 700 < g.y0 < 740]


def _flags(path: str) -> dict:
    return {r["title"]: r for r in api.grade(path)["reasons"]}


def _hidden(tmp_path, body: bytes, extra: bytes = b"", more: dict | None = None, kind: str = "same", text: str = "Hidden") -> dict:
    """「Hidden」が、読み順に残ったまま、見えない字として位置と一緒に伝わる。"""
    p = _write(tmp_path, NAME + body, extra, more)
    pg, md = _page(p)
    got = _faint(p)
    assert got and all(inv and fa == kind for _t, inv, fa in got), got
    assert any(text in ln for ln in pg["reading"])  # 字は落とさない
    title = SAME if kind == "same" else ALPHA
    f = _flags(p)
    assert title in f and f[title]["grade"] == "B" and (SAME if kind != "same" else ALPHA) not in f
    assert len(f[title]["where"]) == 1 and f"「{text}」" in f[title]["where"][0] and "Name" not in f[title]["where"][0]
    assert "view_page" in f[title]["detail"] and title in md
    assert api.grade(p)["grade"] == "B" and _audit_ok(p)
    return f[title]


def _visible(tmp_path, body: bytes, extra: bytes = b"", more: dict | None = None) -> None:
    """「Hidden」が、見える字のまま（見えないと決めていない）。"""
    p = _write(tmp_path, NAME + body, extra, more)
    pg, md = _page(p)
    got = _faint(p)
    assert got and not any(inv or fa for _t, inv, fa in got), got
    assert any("Hidden" in ln for ln in pg["reading"])
    f = _flags(p)
    assert SAME not in f and ALPHA not in f and SAME not in md and ALPHA not in md
    assert _audit_ok(p)


def test_white_text_on_the_bare_page_is_told_with_its_place(tmp_path):
    f = _hidden(tmp_path, WHITE)
    assert "ページの白" in f["detail"] and "とどくは決めません" in f["detail"]


@pytest.mark.parametrize("name, body", [
    ("黒い四角の上の白い四角の上の白い字", WIDE_BLACK + WHITE_BOX + WHITE),
    ("赤い四角の上の赤い字", b"1 0 0 rg 140 700 100 50 re f\n" + b"1 0 0 rg " + TEXT),
    ("黒い四角の上の黒い字", BLACK_BOX + b"0 g " + TEXT),
    ("CMYK の白", b"0 0 0 0 k " + TEXT),
    ("RGB の白い字と、灰色で指定した白い四角", WIDE_BLACK + WHITE_BOX + b"1 1 1 rg " + TEXT),
    ("字にかからない線", b"0 G 140 600 m 240 600 l S\n" + WHITE),
    ("字にかからない黒い四角", b"0 g 300 700 100 50 re f\n" + WHITE),
    ("黒い線で縁取った白い四角の中の白い字", b"0 G 1 g 140 700 100 50 re B\n" + WHITE),
    ("白い字の上に重ねた白い字", WHITE + WHITE),
    ("字の一部だけにかかる白い四角（残りは、ページの白）", b"1 g 140 700 30 50 re f\n" + WHITE),
    ("白い四角を 2 つ並べた上の白い字（下は、ページの白）", b"1 g 140 700 30 50 re f 170 700 70 50 re f\n" + WHITE),
])
def test_text_of_the_color_under_it_is_not_read_as_visible(tmp_path, name, body):
    _hidden(tmp_path, body)


@pytest.mark.parametrize("name, body, extra, more", [
    ("黒い四角の上の白い字", BLACK_BOX + WHITE, b"", None),
    ("何も描いていない所の赤い字", b"1 0 0 rg " + TEXT, b"", None),
    ("何も描いていない所の黒い字", TEXT, b"", None),
    ("少しだけ違う色の四角の上の白い字", b"0.98 g 140 700 100 50 re f\n" + WHITE, b"", None),
    ("字を通る線", b"0 G 140 722 m 240 722 l S\n" + WHITE, b"", None),
    ("黒い字の上に重ねた白い字", b"0 g " + TEXT + WHITE, b"", None),
    ("塗りと線で描く字（線は黒）", b"BT 1 g 0 G 2 Tr /F1 12 Tf 150 720 Td (Hidden) Tj ET\n", b"", None),
    ("線だけで描く白い字", b"BT 1 G 1 Tr /F1 12 Tf 150 720 Td (Hidden) Tj ET\n", b"", None),
    ("画像の上の白い字", DRAW + WHITE, XIM, IMG),
    ("ぼかしの塗りの上の白い字", b"/Sh1 sh\n" + WHITE, SHADE, None),
    ("三角形の塗りの上の白い字", b"0 g 140 700 m 260 700 l 140 760 l f\n" + WHITE, b"", None),
    ("黒い四角の上の、半分透ける白い四角の上の白い字", WIDE_BLACK + b"q /G1 gs " + WHITE_BOX + b"Q\n" + WHITE,
     b"/ExtGState << /G1 << /ca 0.5 >> >>", None),
    ("黒い四角の上の、字の一部だけにかかる白い四角の右の字", WIDE_BLACK + b"1 g 100 700 30 50 re f\n" + WHITE, b"", None),
    ("下の色と掛け合わせる重ね方の白い字", b"q /G1 gs " + WHITE + b"Q\n", b"/ExtGState << /G1 << /BM /Multiply >> >>", None),
    ("半分透ける黒い字", b"q /G1 gs " + TEXT + b"Q\n", b"/ExtGState << /G1 << /ca 0.5 >> >>", None),
    ("読めない不透明度の字", b"q /G1 gs " + TEXT + b"Q\n", b"/ExtGState << /G1 << /ca /Zero >> >>", None),
    ("不透明度 0 の、線だけで描く字（線の不透明度は別）", b"q /G1 gs BT 1 Tr /F1 12 Tf 150 720 Td (Hidden) Tj ET Q\n",
     b"/ExtGState << /G1 << /ca 0 >> >>", None),
])
def test_text_whose_ground_is_not_decided_stays_visible(tmp_path, name, body, extra, more):
    _visible(tmp_path, body, extra, more)


def test_only_the_letters_on_their_own_color_are_told(tmp_path):
    """黒い四角の上の、字の初めの 3 字だけの下にある白い四角：その 3 字だけが見えない。"""
    p = _write(tmp_path, NAME + WIDE_BLACK + b"1 g 140 700 30 50 re f\n" + WHITE)
    assert [(t, fa) for t, _inv, fa in _faint(p)] == [("H", "same"), ("i", "same"), ("d", "same"), ("d", ""), ("e", ""), ("n", "")]
    f = _flags(p)
    assert len(f[SAME]["where"]) == 1 and "「Hid」" in f[SAME]["where"][0]
    assert any("Hidden" in ln for ln in _page(p)[0]["reading"]) and _audit_ok(p)


def test_only_the_letter_that_a_line_crosses_is_left_undecided(tmp_path):
    """線だけの四角の左の縁（x = 160）が通る字（i）だけは決めない：縁のかからない字は、ページの白の上の白い字。"""
    p = _write(tmp_path, NAME + b"0 G 160 690 100 50 re S\n" + WHITE)
    assert [(t, fa) for t, _inv, fa in _faint(p)] == [("H", "same"), ("i", ""), ("d", "same"), ("d", "same"), ("e", "same"), ("n", "same")]
    assert _audit_ok(p)


def test_what_is_drawn_after_the_text_does_not_change_it(tmp_path):
    # 後から描くものは、字の画素にも下の画素にも同じにかかる：白い字は見えないまま
    _hidden(tmp_path, WHITE + b"0 G 140 722 m 240 722 l S\n")
    # 後から描いた透けない塗りの四角の下の字は、前と同じに「塗りの下にあって見えない字」（読み順に入れない）
    p = _write(tmp_path, NAME + WHITE + BLACK_BOX)
    pg, md = _page(p)
    assert pg.get("covered") and "### 塗りの下にあって見えない字" in md and SAME not in _flags(p) and not _faint(p)


@pytest.mark.parametrize("body", [b"q /G1 gs " + TEXT + b"Q\n", b"q /G1 gs " + WHITE + b"Q\n", BLACK_BOX + b"q /G1 gs " + WHITE + b"Q\n"])
def test_text_drawn_with_alpha_zero(tmp_path, body):
    f = _hidden(tmp_path, body, b"/ExtGState << /G1 << /ca 0 >> >>", kind="alpha")
    assert "ca" in f["detail"]


def test_text_put_together_by_actual_text(tmp_path):
    _hidden(tmp_path, b"/Span << /ActualText (Zed) >> BDC\n" + WHITE + b"EMC\n", text="Zed")
    # まとめた字の一部だけが見えないときは、見える字のまま
    p = _write(tmp_path, NAME + WIDE_BLACK + b"1 g 140 700 30 50 re f\n" + b"/Span << /ActualText (Zed) >> BDC\n" + WHITE + b"EMC\n")
    assert [(t, inv, fa) for t, inv, fa in _faint(p)] == [("Zed", False, "")] and SAME not in _flags(p)


def test_text_inside_a_form_xobject(tmp_path):
    form = {7: _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >>", WHITE)}
    _hidden(tmp_path, b"/X1 Do\n", b"/XObject << /X1 7 0 R >>", form)
    _visible(tmp_path, BLACK_BOX + b"/X1 Do\n", b"/XObject << /X1 7 0 R >>", form)
    # 部品の中の黒い四角の上に、ページの白い字
    box = {7: _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842]", BLACK_BOX)}
    _visible(tmp_path, b"/X1 Do\n" + WHITE, b"/XObject << /X1 7 0 R >>", box)


def test_type3_text_is_not_decided(tmp_path):
    """Type 3 の書体は、字の形の中で色を決められる：色で決めない。その字の下の字も、決めない。"""
    p = _write(tmp_path, NAME + b"BT 1 g /F3 12 Tf 150 720 Td (aaa) Tj ET\n", b"/Font << /F1 5 0 R /F3 7 0 R >>", TYPE3)
    assert SAME not in _flags(p) and not any(g.faint for g in _content(p).glyphs)
    p = _write(tmp_path, NAME + b"BT 0 g /F3 12 Tf 150 720 Td (aaaaa) Tj ET\n" + WHITE, b"/Font << /F1 5 0 R /F3 7 0 R >>", TYPE3)
    assert SAME not in _flags(p) and not any(g.faint for g in _content(p).glyphs)


def test_the_second_reading_gives_the_same_page(tmp_path, monkeypatch):
    """見えないかもしれない字のあるページは 2 回読む：字・線・四角・画像は、1 回読んだときと同じ（増えない）。"""
    body = NAME + WIDE_BLACK + WHITE_BOX + WHITE + DRAW + b"0 G 72 600 m 300 600 l S\n" + TEXT.replace(b"720", b"500")
    p = _write(tmp_path, body, XIM, IMG)
    two = _content(p)
    monkeypatch.setattr(content, "_same_color_watch", lambda pc: set())
    one = _content(p)
    assert any(g.faint for g in two.glyphs) and not any(g.faint for g in one.glyphs)
    key = lambda pc: ([(g.text, g.x0, g.y0, g.x1, g.y1, g.color) for g in pc.glyphs], pc.segs, pc.boxes,  # noqa: E731
                      [(q.x0, q.y0, q.x1, q.y1) for q in pc.pictures], len(pc.covered), len(pc.clipped), pc.problems, pc.unmapped)
    assert key(one) == key(two)


def test_a_page_without_such_text_is_read_once(tmp_path, monkeypatch):
    p = _write(tmp_path, NAME + BLACK_BOX + b"1 0 0 rg " + TEXT)
    runs = []
    real = content.Interpreter._run_page
    monkeypatch.setattr(content.Interpreter, "_run_page", lambda self, page: runs.append(1) or real(self, page))
    _content(p)
    assert len(runs) == 1
    p = _write(tmp_path, NAME + BLACK_BOX + WHITE, name="b.pdf")  # 白い字：下を確かめるために、もう 1 回読む
    runs.clear()
    _content(p)
    assert len(runs) == 2


def test_the_color_mark_stays_on_text_of_the_color_under_it(tmp_path):
    """見えない字でも色はある：窓の行の「〔字の色 #FFFFFF〕」の印は残す（どの字かが、行で分かる）。"""
    p = _write(tmp_path, NAME + WHITE)
    pg, _md = _page(p)
    assert any("#FFFFFF" in ln and "Hidden" in ln for ln in pg["reading"])
    # ふつうの色は、見える字で数える：見えない白い字がいちばん多くても、黒がふつう
    p = _write(tmp_path, NAME + b"BT 1 g /F1 12 Tf 150 720 Td (Hidden hidden hidden hidden) Tj ET\n", name="b.pdf")
    pg, _md = _page(p)
    assert "normal_color" not in pg and any("#FFFFFF" in ln for ln in pg["reading"]) and not any("#000000" in ln for ln in pg["reading"])
    # 不透明度 0 の字には、色の印を付けない（何色でも描かれない）
    p = _write(tmp_path, NAME + b"q /G1 gs 1 0 0 rg " + TEXT + b"Q\n", b"/ExtGState << /G1 << /ca 0 >> >>", name="c.pdf")
    pg, _md = _page(p)
    assert any("Hidden" in ln for ln in pg["reading"]) and not any("#FF0000" in ln for ln in pg["reading"])


def test_white_text_with_a_visible_copy_drawn_over_it_is_visible(tmp_path):
    """白く塗った字に、同じ字を黒い線で重ねてある（縁取りの字）：重ね書きとして 1 字にまとめた字は、見える字。"""
    p = _write(tmp_path, NAME + WHITE + b"BT 0 G 1 Tr /F1 12 Tf 150 720 Td (Hidden) Tj ET\n")
    pg, md = _page(p)
    assert sum(ln.count("Hidden") for ln in pg["reading"]) == 1
    assert SAME not in _flags(p) and SAME not in md and api.grade(p)["grade"] == "A"
    # 見える字の上に、同じ字を不透明度 0 で重ねても、見える字のまま
    p = _write(tmp_path, NAME + TEXT + b"q /G1 gs " + TEXT + b"Q\n", b"/ExtGState << /G1 << /ca 0 >> >>", name="b.pdf")
    assert ALPHA not in _flags(p) and sum(ln.count("Hidden") for ln in _page(p)[0]["reading"]) == 1
