"""PDF の切り抜き（クリップ）：描かれているが、切り抜きの外にあって見えない字・線を、見える内容として読まない。

- 切り抜き（W・W*、部品の BBox、ページの枠）の道の外枠の外にある字は、読み順・見出し・重なりの確かめに入れず、
  ページの「切り抜きで見えない字」の節に、位置（mm）と一緒に分けて示す（字は 1 字も落とさない）
- 外枠にかかる字は見える字のまま。外枠のすぐ外の字は、埋め込んだ書体の字形の輪郭が外にあるときだけ、見えない字にする
- 外枠のすぐ外の空白（墨のない字）は、同じ文字の組（BT〜ET）の隣の字と同じ側。空白だけの組は、見える側のまま
- 外枠の中にあるが、切り抜きの形（三角など）の外にある字は、決めない（見える字のまま）
- 外枠の外にある線は罫線にしない。外枠をまたぐ線・塗りの四角は、外枠の中の部分だけにする
- 線・塗りは、ページの端では切らない（ページの外へ出る線は、切り抜きがないときと同じ形のまま）
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from test_robustness import _page_objs, _pdf, _stream  # noqa: E402

from todoku import api, lazy  # noqa: E402
from todoku.pdf import content  # noqa: E402
from todoku.pdf.content import _cut_seg, page_content  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402

NAME = b"BT /F1 12 Tf 72 720 Td (Name:) Tj ET\n"
HIDDEN = b"BT /F1 12 Tf 150 720 Td (Hidden) Tj ET\n"
TITLE = "### 切り抜きで見えない字"
MM = 72 / 25.4


def _write(tmp_path, body: bytes, extra: bytes = b"", more: dict | None = None, page: bytes | None = None, name: str = "a.pdf") -> str:
    objs = _page_objs(body, extra)
    objs.update(more or {})
    if page is not None:
        objs[3] = page
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


def _rule(pg: dict) -> dict:
    (f,) = [f for f in pg["fields"] if f["id"].startswith("p1.u")]
    return f


# ---------------------------------------------------------------------- 字
def test_text_outside_the_clip_is_told_apart_from_the_reading(tmp_path):
    p = _write(tmp_path, NAME + b"q 0 0 50 50 re W n\n" + HIDDEN + b"Q\n")
    pg, md = _page(p)
    assert pg["reading"] == ["Name:"] and pg["chars"] == len("Name:")
    (c,) = pg["clipped"]
    assert c["text"] == "Hidden"
    assert abs(c["left_mm"] - 150 / MM) < 0.2 and 36 < c["top_mm"] < 43
    assert TITLE + "（6 字）" in md and f"- （左 {c['left_mm']}・上 {c['top_mm']} mm）Hidden" in md
    assert md.index("Hidden") > md.index("### 読み順")
    assert not any("Hidden" in str(f) for f in pg["fields"])  # 見出し・近くの字にもならない
    assert _audit_ok(p)


def test_without_a_clip_the_same_text_is_ordinary(tmp_path):
    pg, md = _page(_write(tmp_path, NAME + HIDDEN))
    assert "Hidden" in "\n".join(pg["reading"]) and "clipped" not in pg and TITLE not in md


def test_the_clip_takes_effect_after_the_path_is_ended_and_q_restores_it(tmp_path):
    # W のあと、道を終える命令（n）の前に切り抜きは効かない。Q のあとは、前の切り抜きに戻る
    body = b"q 0 0 50 50 re W n\n" + HIDDEN + b"Q\n" + NAME
    pg, _md = _page(_write(tmp_path, body))
    assert pg["reading"] == ["Name:"] and [c["text"] for c in pg["clipped"]] == ["Hidden"]
    # W だけで、道を終えていない：その道を塗る命令のあとで効く
    body = b"0 0 50 50 re W* f\n" + HIDDEN
    pg, _md = _page(_write(tmp_path, body, name="b.pdf"))
    assert [c["text"] for c in pg["clipped"]] == ["Hidden"]
    # 道のない W は、切り抜きを変えない（決めない）
    pg, _md = _page(_write(tmp_path, b"W n\n" + HIDDEN, name="c.pdf"))
    assert "Hidden" in "\n".join(pg["reading"]) and "clipped" not in pg


def test_nested_clips_narrow_and_the_matrix_applies(tmp_path):
    # 外の切り抜き（x 0〜300）と中の切り抜き（x 200〜500）：見えるのは x 200〜300 だけ
    body = (b"q 0 0 300 842 re W n q 200 0 300 842 re W n\n"
            b"BT /F1 12 Tf 72 720 Td (Left) Tj ET BT /F1 12 Tf 220 720 Td (Mid) Tj ET BT /F1 12 Tf 400 720 Td (Right) Tj ET Q Q\n")
    pg, _md = _page(_write(tmp_path, body))
    assert pg["reading"] == ["Mid"] and " ".join(c["text"] for c in pg["clipped"]).split() == ["Left", "Right"]
    # 行列（cm）で動かした道：切り抜きは、ページの座標で x 200〜260
    body = b"q 1 0 0 1 200 0 cm 0 0 60 842 re W n 1 0 0 1 -200 0 cm\n" + NAME + b"BT /F1 12 Tf 210 720 Td (In) Tj ET Q\n"
    pg, _md = _page(_write(tmp_path, body, name="b.pdf"))
    assert pg["reading"] == ["In"] and [c["text"] for c in pg["clipped"]] == ["Name:"]


def test_text_that_straddles_the_clip_stays_in_the_reading(tmp_path):
    # 「Hidden」（x 150〜187）の途中（x 170）までの切り抜き：外枠にかかる字は見える字のまま、すっかり外の字だけを分ける
    pg, _md = _page(_write(tmp_path, b"q 0 0 170 842 re W n\n" + HIDDEN + b"Q\n"))
    shown, hid = "".join(pg["reading"]), "".join(c["text"] for c in pg.get("clipped", []))
    assert shown + hid == "Hidden" and shown.startswith("Hid") and hid.endswith("n")


def test_a_glyph_just_outside_the_clip_needs_its_outline_to_be_called_hidden(tmp_path, monkeypatch):
    # 字の四角が外枠のすぐ外（字の大きさの半分の内）：埋め込んでいない書体は、字形が分からないので、見える字のまま
    body = b"q 0 0 148 842 re W n\n" + HIDDEN + b"Q\n"
    p = _write(tmp_path, body)
    assert "H" in "".join(_page(p)[0]["reading"])
    # 埋め込んだ書体：字形の輪郭の四角が外枠の外なら、見えない字
    real = content.load_font

    def embedded(doc, ref, name):
        f = real(doc, ref, name)
        f.embedded = True
        return f

    monkeypatch.setattr(content, "load_font", embedded)
    monkeypatch.setattr(content.Interpreter, "_clip_ink", lambda self, res, name, pf, code, text: (0.05, 0.0, 0.6, 0.7))
    pc = _content(p)
    assert "".join(g.text for g in pc.clipped) == "Hidden" and not pc.glyphs
    # 輪郭が外枠にかかる（字の原点より左に出る字形）：見える字のまま
    monkeypatch.setattr(content.Interpreter, "_clip_ink", lambda self, res, name, pf, code, text: (-0.3, 0.0, 0.6, 0.7))
    pc = _content(p)
    assert "".join(g.text for g in pc.glyphs).startswith("H")
    # 輪郭が分からない字：見える字のまま
    monkeypatch.setattr(content.Interpreter, "_clip_ink", lambda self, res, name, pf, code, text: None)
    assert "".join(g.text for g in _content(p).glyphs).startswith("H")


def test_a_space_just_outside_the_clip_goes_with_the_glyph_next_to_it(tmp_path, monkeypatch):
    # 空白には墨がない（切り抜かれるものがない）：同じ文字の組（BT〜ET）のすぐ前の字、なければ次の字と同じ側に入れる
    real = content.load_font

    def embedded(doc, ref, name):
        f = real(doc, ref, name)
        f.embedded = True
        return f

    monkeypatch.setattr(content, "load_font", embedded)
    monkeypatch.setattr(content.Interpreter, "_clip_ink",
                        lambda self, res, name, pf, code, text: () if text == " " else (0.05, 0.0, 0.6, 0.7))
    clip = b"q 0 0 148 842 re W n\n"

    def sides(body: bytes, name: str):
        pc = _content(_write(tmp_path, clip + body + b"Q\n", name=name))
        return "".join(g.text for g in pc.glyphs), "".join(g.text for g in pc.clipped)

    # 空白だけの組：見える側のまま（1 字ごとに切り抜く作りの、行の終わりの空白）
    assert sides(b"BT /F1 12 Tf 150 720 Td ( ) Tj ET\n", "a.pdf") == (" ", "")
    # 見えない字の後ろ・前の空白：見えない側
    assert sides(b"BT /F1 12 Tf 150 720 Td (A ) Tj ET\n", "b.pdf") == ("", "A ")
    assert sides(b"BT /F1 12 Tf 150 720 Td ( A) Tj ET\n", "c.pdf") == ("", " A")
    # 見える字の後ろの空白：見える側
    assert sides(b"BT /F1 12 Tf 139.9 720 Td [(A) -100 ( )] TJ ET\n", "d.pdf") == ("A ", "")
    # 組が変われば、前の組の字には合わせない
    assert sides(b"BT /F1 12 Tf 150 720 Td (A) Tj ET BT /F1 12 Tf 150 700 Td ( ) Tj ET\n", "e.pdf") == (" ", "A")


def test_inside_the_box_of_a_slanted_clip_is_not_decided(tmp_path):
    # 三角の切り抜き（外枠は x 0〜400・y 600〜800）：外枠の中で三角の外にある字は、決めない（見える字のまま）。外枠の外の字は分ける
    body = (b"q 0 600 m 400 600 l 0 800 l h W n\n" + b"BT /F1 12 Tf 300 780 Td (Corner) Tj ET\n"
            + b"BT /F1 12 Tf 72 500 Td (Below) Tj ET Q\n")
    pg, _md = _page(_write(tmp_path, body))
    assert pg["reading"] == ["Corner"] and [c["text"] for c in pg["clipped"]] == ["Below"]


def test_a_curved_clip_uses_its_control_points(tmp_path):
    # 曲線は、制御点の外枠から出ない：その外の字は、見えない
    body = b"q 0 0 m 0 100 100 100 100 0 c h W n\n" + HIDDEN + b"Q\n" + NAME
    pg, _md = _page(_write(tmp_path, body))
    assert pg["reading"] == ["Name:"] and [c["text"] for c in pg["clipped"]] == ["Hidden"]


def test_a_form_clips_its_content_to_its_bbox(tmp_path):
    form = _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /F1 5 0 R >> >>",
                   b"BT /F1 12 Tf 10 50 Td (In) Tj ET BT /F1 12 Tf 150 50 Td (Out) Tj ET")
    p = _write(tmp_path, NAME + b"q 1 0 0 1 200 600 cm /X1 Do Q\n", extra=b"/XObject << /X1 6 0 R >>", more={6: form})
    pg, _md = _page(p)
    assert "In" in " ".join(pg["reading"]) and "Out" not in " ".join(pg["reading"])
    (c,) = pg["clipped"]
    assert c["text"] == "Out" and abs(c["left_mm"] - 350 / MM) < 0.2
    assert _audit_ok(p)


def test_text_off_the_page_is_not_in_the_reading(tmp_path):
    # ページの枠（CropBox を MediaBox に切り詰めたもの）の外に描いた字
    body = NAME + b"BT /F1 12 Tf 700 720 Td (Margin) Tj ET BT /F1 12 Tf 72 900 Td (Above) Tj ET\n"
    pg, md = _page(_write(tmp_path, body))
    assert pg["reading"] == ["Name:"] and sorted(c["text"] for c in pg["clipped"]) == ["Above", "Margin"]
    above = next(c for c in pg["clipped"] if c["text"] == "Above")
    assert above["top_mm"] < 0  # ページの上の外
    page = (b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /CropBox [0 0 140 842] /Resources << /Font << /F1 5 0 R >> >>"
            b" /Contents 4 0 R >>")
    pg, _md = _page(_write(tmp_path, NAME + HIDDEN, page=page, name="b.pdf"))
    assert pg["reading"] == ["Name:"] and [c["text"] for c in pg["clipped"]] == ["Hidden"]


def test_a_page_with_only_hidden_text_says_where_the_text_is(tmp_path):
    pg, md = _page(_write(tmp_path, b"q 0 0 50 50 re W n\n" + HIDDEN + b"Q\n"))
    assert pg["chars"] == 0 and pg.get("blank") is True
    assert "切り抜きで見えない字は、下の節にあります" in md and TITLE + "（6 字）" in md


# ---------------------------------------------------------------------- 線・塗り
def test_a_rule_outside_the_clip_makes_no_place(tmp_path):
    shown = _write(tmp_path, NAME + b"100 600 m 300 600 l S\n")
    assert abs(_rule(_page(shown)[0])["facts"]["length_mm"] - 200 / MM) < 0.3
    hidden = _write(tmp_path, NAME + b"q 0 0 50 50 re W n 100 600 m 300 600 l S Q\n", name="b.pdf")
    pg, _md = _page(hidden)
    assert not [f for f in pg["fields"] if f["id"].startswith("p1.u")]
    assert _content(hidden).clipped_shapes == 1


def test_a_rule_that_crosses_the_clip_is_cut_to_the_visible_part(tmp_path):
    # x 100〜300 の線を、x 100〜200 の切り抜きの中に描く：見える罫線は 100 pt
    p = _write(tmp_path, NAME + b"q 100 500 100 200 re W n 100 600 m 300 600 l S Q\n")
    f = _rule(_page(p)[0])
    assert abs(f["facts"]["length_mm"] - 100 / MM) < 0.3 and abs(f["facts"]["rect_mm"][2] - 200 / MM) < 0.3
    # 細い塗りの四角で描いた線も同じ
    p = _write(tmp_path, NAME + b"q 100 500 100 200 re W n 100 599.5 200 1 re f Q\n", name="b.pdf")
    assert abs(_rule(_page(p)[0])["facts"]["length_mm"] - 100 / MM) < 0.3


def test_a_line_just_outside_the_clip_keeps_its_visible_half(tmp_path):
    # 太さ 2 pt の線が、外枠の縁の 0.5 pt 外を通る：墨の半分は中に入るので、線は残る。3 pt 外なら見えない
    pc = _content(_write(tmp_path, b"q 100 600 200 100 re W n 2 w 100 599.5 m 300 599.5 l S Q\n"))
    assert len(pc.segs) == 1 and pc.clipped_shapes == 0
    pc = _content(_write(tmp_path, b"q 100 600 200 100 re W n 2 w 150 597 m 250 597 l S Q\n", name="b.pdf"))
    assert not pc.segs and pc.clipped_shapes == 1


def test_a_filled_box_is_cut_to_the_clip(tmp_path):
    pc = _content(_write(tmp_path, b"q 100 600 100 50 re W n 0.8 g 50 580 400 200 re f Q\n"))
    (b,) = pc.boxes
    assert (b.x0, b.y0, b.x1, b.y1) == (100, 600, 200, 650) and b.stroke is False
    # 縁の線つきの四角：外枠の中にある辺の部分だけが線になり、四角（マスの枠）にはしない
    pc = _content(_write(tmp_path, b"q 100 600 100 50 re W n 150 620 200 100 re S Q\n", name="b.pdf"))
    assert not pc.boxes
    ends = sorted(tuple(sorted([(round(s.x0), round(s.y0)), (round(s.x1), round(s.y1))])) for s in pc.segs)
    assert ends == [((150, 620), (150, 650)), ((150, 620), (200, 620))]


def test_shapes_are_not_cut_at_the_edge_of_the_page(tmp_path):
    # ページの右へ出る四角（ページの端のつまみ）：ページの端では切らず、切り抜きがないときと同じ形のまま
    tab = b"560 400 60 40 re S 0.8 g -20 800 700 60 re f\n"
    plain = _content(_write(tmp_path, tab))
    for i, clip in enumerate((b"0 0 595 842 re W n\n", b"-10 -10 700 900 re W n\n")):  # ページの枠と同じ・ページより大きい切り抜き
        pc = _content(_write(tmp_path, b"q " + clip + tab + b"Q\n", name=f"c{i}.pdf"))
        assert [(s.x0, s.y0, s.x1, s.y1) for s in pc.segs] == [(s.x0, s.y0, s.x1, s.y1) for s in plain.segs]
        assert [(b.x0, b.y0, b.x1, b.y1, b.stroke) for b in pc.boxes] == [(b.x0, b.y0, b.x1, b.y1, b.stroke) for b in plain.boxes]
        assert pc.clipped_shapes == 0
    assert len(plain.segs) == 4 and any(b.stroke for b in plain.boxes) and plain.clipped_shapes == 0
    # ページの中にある辺では切る：左は切り抜きの辺（x 570）、右はページの外へ出たまま
    pc = _content(_write(tmp_path, b"q 570 0 100 842 re W n\n" + tab + b"Q\n", name="d.pdf"))
    assert sorted((round(min(s.x0, s.x1)), round(max(s.x0, s.x1))) for s in pc.segs) == [(570, 620), (570, 620), (620, 620)]
    (b,) = pc.boxes
    assert (b.x0, b.x1) == (570, 680)


def test_dashes_made_by_clipping_one_long_line_are_dashes(tmp_path):
    # 作る道具が、破線を「1 つぶごとに切り抜いて、長い線を塗る」で描く：見えるのは、つぶだけ
    body = b"".join(b"q %d 599 6 2 re W n 100 599.5 200 1 re f Q\n" % x for x in range(100, 300, 12))
    pc = _content(_write(tmp_path, body))
    assert len(pc.segs) == len(range(100, 300, 12)) and all(abs((s.x1 - s.x0) - 6) < 0.01 for s in pc.segs)


def test_cut_seg():
    cl = (10, 0, 50, 10)
    assert _cut_seg(0, 5, 100, 5, cl, 1.0) == (10, 5, 50, 5)  # 線の向きには、外枠のとおり
    assert _cut_seg(0, 10.5, 100, 10.5, cl, 1.0) == (10, 10.5, 50, 10.5)  # 直角な向きには、太さの半分まで
    assert _cut_seg(0, 12, 100, 12, cl, 1.0) is None
    assert _cut_seg(20, 2, 30, 8, cl, 1.0) == (20, 2, 30, 8)
    assert _cut_seg(0, 0, 100, 100, (10, 0, 50, 10), 0.0) is None  # 角に触れるだけ


# ---------------------------------------------------------------------- 書き込みの重なりの確かめ
def test_writing_over_hidden_text_is_not_refused(tmp_path):
    # 見えない字の上：重なる字はないので、書ける。同じ所に見える字があれば、断る
    box = {"@1": {"page": 1, "box": [52, 38, 90, 44], "value": "Taro"}}
    hidden = _write(tmp_path, NAME + b"q 0 0 50 50 re W n\n" + HIDDEN + b"Q\n")
    r = api.fill(hidden, box, str(tmp_path / "o1.pdf"))
    assert r["read_back_ok"] is not False and (tmp_path / "o1.pdf").exists()
    shown = _write(tmp_path, NAME + HIDDEN, name="b.pdf")
    try:
        api.fill(shown, box, str(tmp_path / "o2.pdf"))
    except api.ApiError as e:
        assert "ページの字（「Hidden」）" in str(e)
    else:
        raise AssertionError("見える字の上に書けた")
