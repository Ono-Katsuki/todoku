"""PDF のレイヤー（optional content）：既定の設定で画面に出ない内容を、ふつうの内容として読まない。

- 画面に出ないレイヤーの字は、読み順に入れず、ページの「既定では表示されないレイヤー」の節に分けて示す（読む B）
- そのレイヤーの線は欄にしない・画像は画像の一覧に入れない（数は伝える）
- 画面には出るが印刷に出ないレイヤーの字は、読み順に入れたまま知らせる（読む B）
- 描く側（view_page）と読む側（窓）が同じ決まり（pdf/layers.py）を使う
- レイヤーの設定のない書類・どのレイヤーも見える書類は、いままでどおり
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from test_robustness import _page_objs, _pdf, _stream  # noqa: E402

from todoku import api, lazy, limits  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.layers import layers  # noqa: E402

SHOWN = b"BT /F1 12 Tf 72 720 Td (Name:) Tj ET\n"
SAMPLE = b"BT /F1 12 Tf 150 720 Td (Taro Yamada) Tj ET\n"
TITLE = "既定では表示されないレイヤーの字・画像"


def _u16(s: str) -> bytes:
    return b"<FEFF" + s.encode("utf-16-be").hex().upper().encode() + b">"


def _doc(content: bytes, d: bytes = b"/OFF [6 0 R]", more: dict[int, bytes] | None = None, res: bytes = b"",
         ocprops: bool = True, g1: bytes = b"", g2: bytes = b"") -> bytes:
    """レイヤー 6（「記入例」）と 7（「Notes」）のある 1 ページの PDF。d は既定の設定（D）の中身。"""
    objs = _page_objs(content, b"/Properties << /L1 6 0 R /L2 7 0 R /M1 8 0 R >> " + res)
    if ocprops:
        objs[1] = b"<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R 7 0 R] /D << " + d + b" >> >> >>"
    objs[6] = b"<< /Type /OCG /Name " + _u16("記入例") + b" " + g1 + b" >>"
    objs[7] = b"<< /Type /OCG /Name (Notes) " + g2 + b" >>"
    objs[8] = b"<< /Type /OCMD /OCGs [6 0 R 7 0 R] /P /AllOn >>"
    objs.update(more or {})
    return _pdf(objs)


def _write(tmp_path, blob: bytes, name: str = "a.pdf") -> str:
    p = tmp_path / name
    p.write_bytes(blob)
    api._WINDOWS.clear()
    lazy.clear()
    return str(p)


def _page(path: str) -> tuple[dict, str]:
    _kind, win, parts = api.window(path)
    return win["pages"][0], "\n".join(md for _n, md in parts)


def _reading(pg: dict) -> str:
    return "\n".join(pg["reading"])


def _reasons(path: str) -> dict:
    return {r["title"]: r["grade"] for r in api.grade(path)["reasons"]}


def _audit_ok(path: str) -> bool:
    a = api.audit(path)
    return all(c["ok"] for c in a["checks"] if c["name"].startswith("読み落とし"))


# ---------------------------------------------------------------------- 画面に出ないレイヤー
def test_text_on_an_off_layer_is_told_apart_from_the_reading(tmp_path):
    p = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n"))
    pg, md = _page(p)
    assert "Name:" in _reading(pg) and "Taro Yamada" not in _reading(pg)
    (off,) = pg["off_layers"]
    assert off["layer"] == "記入例" and off["view"] is False and off["print"] is False
    assert off["lines"] == ["Taro Yamada"]
    assert f"### {TITLE}" in md and "レイヤー「記入例」（画面にも印刷にも出ない" in md and "  - Taro Yamada" in md
    assert md.index("Taro Yamada") > md.index("### 読み順")
    # 字の数：ページの字（読み順）には入れない。読み落としの検査は、分けて示した字も数える
    assert pg["chars"] == len("Name:")
    assert _reasons(p).get(TITLE) == "B"
    assert _audit_ok(p)
    # 書類の頭にも、どのレイヤーが出ないかを書く
    assert "画面に出ないレイヤー" in md.split("## ページ 1")[0] and "「記入例」" in md.split("## ページ 1")[0]


def test_the_same_layer_switched_on_reads_as_ordinary_text(tmp_path):
    p = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n", d=b""))
    pg, md = _page(p)
    assert "Taro Yamada" in _reading(pg) and not pg.get("off_layers")
    assert TITLE not in md and TITLE not in _reasons(p)
    assert _audit_ok(p)


def test_content_marks_without_a_layer_setting_are_ordinary(tmp_path):
    p = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n", ocprops=False))
    pg, md = _page(p)
    assert "Taro Yamada" in _reading(pg) and not pg.get("off_layers") and TITLE not in md


def test_lines_on_an_off_layer_make_no_field_and_images_are_counted(tmp_path):
    box = b"100 600 200 30 re S\n"
    img = _stream(b"/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", b"\x00")
    draw = b"q 100 0 0 100 300 500 cm /Im1 Do Q\n"
    shown = _write(tmp_path, _doc(SHOWN + box + draw, d=b"", more={9: img}, res=b"/XObject << /Im1 9 0 R >>"), "on.pdf")
    pg, _md = _page(shown)
    assert any(f.get("source") == "layout" for f in pg["fields"]) and pg["images"] == 1
    hidden = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + box + draw + b"EMC\n", more={9: img},
                                   res=b"/XObject << /Im1 9 0 R >>"), "off.pdf")
    pg, md = _page(hidden)
    assert not any(f.get("source") == "layout" for f in pg["fields"])
    assert pg["images"] == 0 and not pg.get("image_only")
    (off,) = pg["off_layers"]
    assert off["lines"] == [] and off["images"] == 1
    assert "字 0 字・画像 1 個" in md


def test_a_form_part_and_an_image_marked_with_an_off_layer(tmp_path):
    form = _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842] /OC 6 0 R /Resources << /Font << /F1 5 0 R >> >>", SAMPLE)
    img = _stream(b"/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /OC 6 0 R",
                  b"\x00")
    p = _write(tmp_path, _doc(SHOWN + b"/X1 Do q 100 0 0 100 300 500 cm /Im1 Do Q\nBT /F1 12 Tf 72 700 Td (After) Tj ET\n",
                              more={9: form, 10: img}, res=b"/XObject << /X1 9 0 R /Im1 10 0 R >>"))
    pg, _md = _page(p)
    assert "After" in _reading(pg) and "Taro Yamada" not in _reading(pg)
    (off,) = pg["off_layers"]
    assert off["lines"] == ["Taro Yamada"] and off["images"] == 1 and pg["images"] == 0
    assert _audit_ok(p)


@pytest.mark.parametrize("ocmd, hidden", [
    (b"/OCGs [6 0 R 7 0 R] /P /AllOn", True), (b"/OCGs [6 0 R 7 0 R] /P /AnyOn", False),
    (b"/OCGs [6 0 R 7 0 R] /P /AnyOff", False), (b"/OCGs [6 0 R 7 0 R] /P /AllOff", True),
    (b"/OCGs [6 0 R 7 0 R]", False), (b"/OCGs 6 0 R", True), (b"", False),
    # VE（見えるかの式）があれば、P より先
    (b"/OCGs [6 0 R 7 0 R] /P /AnyOn /VE [/And 6 0 R 7 0 R]", True),
    (b"/OCGs [6 0 R] /VE [/Not 6 0 R]", False),
    (b"/VE [/Or [/Not 7 0 R] [/And 6 0 R 7 0 R]]", True),
    (b"/VE [/Or [/Not 6 0 R] 6 0 R]", False),
])
def test_membership_marks(tmp_path, ocmd, hidden):
    p = _write(tmp_path, _doc(SHOWN + b"/OC /M1 BDC\n" + SAMPLE + b"EMC\n", more={8: b"<< /Type /OCMD " + ocmd + b" >>"}))
    pg, _md = _page(p)
    assert ("Taro Yamada" not in _reading(pg)) is hidden
    assert bool(pg.get("off_layers")) is hidden
    if hidden:
        assert "記入例" in pg["off_layers"][0]["layer"]


def test_base_state_off_hides_every_layer_not_switched_on(tmp_path):
    body = SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n/OC /L2 BDC BT /F1 12 Tf 72 680 Td (Second) Tj ET EMC\n"
    p = _write(tmp_path, _doc(body, d=b"/BaseState /OFF /ON [7 0 R]"))
    pg, _md = _page(p)
    assert "Second" in _reading(pg) and "Taro Yamada" not in _reading(pg)
    assert [o["layer"] for o in pg["off_layers"]] == ["記入例"]


def test_a_layer_inside_a_layer_and_a_mark_left_open(tmp_path):
    # 見えるレイヤーの中の、見えないレイヤー。閉じていない内容の印は、その流れ（フォームの部品）の終わりまで
    form = _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> "
                   b"/Properties << /L1 6 0 R >> >>", b"/OC /L1 BDC\n" + SAMPLE)
    body = (b"/OC /L2 BDC BT /F1 12 Tf 72 720 Td (Outer) Tj ET /OC /L1 BDC BT /F1 12 Tf 72 700 Td (Inner) Tj ET EMC "
            b"BT /F1 12 Tf 72 680 Td (Outer2) Tj ET EMC\n/X1 Do\nBT /F1 12 Tf 72 660 Td (Tail) Tj ET\n")
    p = _write(tmp_path, _doc(body, more={9: form}, res=b"/XObject << /X1 9 0 R >>"))
    pg, _md = _page(p)
    r = _reading(pg)
    assert "Outer" in r and "Outer2" in r and "Tail" in r and "Inner" not in r and "Taro Yamada" not in r
    (off,) = pg["off_layers"]
    assert sorted(off["lines"]) == ["Inner", "Taro Yamada"]
    assert _audit_ok(p)


# ---------------------------------------------------------------------- 印刷と画面の違い
PRINT_OFF = b"/Usage << /Print << /PrintState /OFF >> >>"
PRINT_ON = b"/Usage << /Print << /PrintState /ON >> >>"
VIEW_OFF = b"/Usage << /View << /ViewState /OFF >> >>"


def test_text_shown_on_screen_but_not_printed_is_flagged_in_place(tmp_path):
    p = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n", g1=PRINT_OFF,
                              d=b"/AS [<< /Event /Print /Category [/Print] /OCGs [6 0 R] >>]"))
    pg, md = _page(p)
    assert "Taro Yamada" in _reading(pg) and not pg.get("off_layers")
    assert _reasons(p).get("印刷には出ない字") == "B"
    assert "印刷には出ない字" in md and "「Taro Yamada」" in md and "レイヤー「記入例」" in md


def test_text_hidden_on_screen_but_printed_is_told_as_printed(tmp_path):
    p = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n", g1=PRINT_ON,
                              d=b"/OFF [6 0 R] /AS [<< /Event /Print /Category [/Print] /OCGs [6 0 R] >>]"))
    pg, md = _page(p)
    (off,) = pg["off_layers"]
    assert off["view"] is False and off["print"] is True
    assert "レイヤー「記入例」（画面に出ない・印刷には出る" in md


def test_a_usage_that_is_not_applied_changes_nothing(tmp_path):
    # Usage に PrintState があっても、D の AS がそのレイヤーに当てていなければ、既定の状態のまま
    p = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n", d=b"", g1=PRINT_OFF))
    assert "印刷には出ない字" not in _reasons(p)
    p = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n", g1=PRINT_OFF,
                              d=b"/AS [<< /Event /Print /Category [/Print] /OCGs [7 0 R] >>]"), "b.pdf")
    assert "印刷には出ない字" not in _reasons(p)


def test_view_usage_switches_a_layer_off_on_screen(tmp_path):
    p = _write(tmp_path, _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n", g1=VIEW_OFF,
                              d=b"/AS [<< /Event /View /Category [/View /Zoom] /OCGs [6 0 R] >>]"))
    pg, _md = _page(p)
    assert "Taro Yamada" not in _reading(pg)
    (off,) = pg["off_layers"]
    assert off["view"] is False and off["print"] is True


def test_layer_states_of_the_document():
    blob = _doc(SHOWN, g2=PRINT_OFF, d=b"/OFF [6 0 R] /AS [<< /Event /Print /Category [/Print] /OCGs [6 0 R 7 0 R] >>]")
    ly = layers(PdfDocument(blob, "a.pdf"))
    assert [(s.name, s.view, s.print) for s in ly.special()] == [("記入例", False, False), ("Notes", True, False)]
    assert not layers(PdfDocument(_doc(SHOWN, d=b""), "a.pdf")).special()
    assert not layers(PdfDocument(_doc(SHOWN, ocprops=False), "a.pdf")).any


# ---------------------------------------------------------------------- 描く側と同じ決まり
def _ink(path: str) -> int:
    from todoku.pdf import pageimage

    r = pageimage.render(PdfDocument(Path(path).read_bytes(), "a.pdf"), 1)
    return sum(1 for i in range(0, len(r["rgb"]), 3) if r["rgb"][i] < 128)


@pytest.mark.parametrize("d, g1, hidden", [
    (b"/OFF [6 0 R]", b"", True), (b"", b"", False),
    (b"/AS [<< /Event /View /Category [/View] /OCGs [6 0 R] >>]", VIEW_OFF, True),
    (b"/AS [<< /Event /Print /Category [/Print] /OCGs [6 0 R] >>]", PRINT_OFF, False),
])
def test_the_page_image_hides_what_the_window_tells_apart(tmp_path, d, g1, hidden):
    only = b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n"
    p = _write(tmp_path, _doc(only, d=d, g1=g1))
    pg, _md = _page(p)
    assert bool(pg.get("off_layers")) is hidden
    assert (_ink(p) == 0) is hidden


# ---------------------------------------------------------------------- フォームの欄
def test_a_form_field_on_an_off_layer_is_told_as_unseen(tmp_path):
    objs_more = {
        9: b"<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /Rect [150 700 300 720] /F 4 /P 3 0 R /OC 6 0 R >>",
    }
    objs = _page_objs(SHOWN)
    objs[1] = (b"<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [9 0 R] >> "
               b"/OCProperties << /OCGs [6 0 R] /D << /OFF [6 0 R] >> >> >>")
    objs[3] = objs[3].replace(b"/Contents 4 0 R", b"/Contents 4 0 R /Annots [9 0 R]")
    objs[6] = b"<< /Type /OCG /Name (Hidden fields) >>"
    objs.update(objs_more)
    p = _write(tmp_path, _pdf(objs))
    pg, md = _page(p)
    assert "［name：画面にも紙にも出ない］" in _reading(pg)
    f = next(f for f in api.fields(p) if f["id"] == "name")
    assert "レイヤー「Hidden fields」" in str(f)
    out = str(tmp_path / "out.pdf")
    with pytest.raises(Exception, match="画面にも紙にも出ない"):
        api.fill(p, {"name": "x"}, out)


# ---------------------------------------------------------------------- 大きな書類の道
def test_large_road_tells_off_layers_like_the_normal_window(tmp_path, monkeypatch):
    blob = _doc(SHOWN + b"/OC /L1 BDC\n" + SAMPLE + b"EMC\n")
    p = _write(tmp_path, blob)
    _pg, md = _page(p)
    normal = md[md.index("## ページ 1"):].strip()
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 300)
    api._WINDOWS.clear()
    lazy.clear()
    doc = lazy.open_doc(p, force=True)
    assert type(doc).__name__ == "PdfLazy"
    assert "画面に出ないレイヤー" in doc.overview_text() and "「記入例」" in doc.overview_text()
    from todoku import mcp_server

    got = mcp_server._read_window({"path": p, "pages": "1"})
    text = "\n".join(c["text"] for c in got["content"] if c.get("type") == "text")
    assert normal in text
    assert _reasons(p).get(TITLE) == "B"


# ---------------------------------------------------------------------- 閉じていない印の後ろに書く
BOX = b"0.5 w 140 705 200 30 re S\n"


@pytest.mark.parametrize("tail", [b"/OC /L1 BDC\n" + SAMPLE, b"/Artifact BMC\n", b"q /OC /L1 BDC /Span BMC\n"])
def test_writing_after_an_unclosed_mark_lands_outside_the_mark(tmp_path, tail):
    """元の描画が印（画面に出ないレイヤー・飾り）を閉じずに終わるページ：足した字が、その印の中に入らない。"""
    from todoku.pdf import circle

    p = _write(tmp_path, _doc(SHOWN + BOX + tail))
    fid = next(f["id"] for f in api.fields(p) if f["id"].startswith("p1.t"))
    out = str(tmp_path / "out.pdf")
    r = api.fill(p, {fid: "Hanako"}, out)
    assert r["proof_ok"] and r["read_back_ok"] and r["out_path"].endswith("out.pdf"), r.get("read_back_differences")
    doc = PdfDocument(open(r["out_path"], "rb").read(), "out.pdf")
    page = next(iter(doc.pages()))
    assert circle.open_depth(doc.page_content(page)) == (0, 0)
    pg, _md = _page(r["out_path"])
    assert f"［{fid}＝Hanako］" in _reading(pg)
