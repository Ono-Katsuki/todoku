"""PDF に入れた、字でないもの（チェック・ラジオの印、字を囲む ○、印影の画像）が見えるかを確かめる。

前は、書いた字だけを、書く前と書いた後の絵で比べていた：印・○・画像は、同じ色の下地の上でも、後から描かれるものの下でも、
印の見た目が何も描かなくても、何も伝えずに「書けた」と返していた。
いまは、その四角（印は部品、○ は外枠、画像は置いた枠）の画素を、書く前と書いた後で比べる：1 つも変わらなければ読み戻しの違い、
いちばん差のある所でも明るさの比が 3:1 未満なら、位置（mm）つきで知らせる。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

import pdfmini  # noqa: E402
from test_pdf import jp_font_objs, jp_text, make_pdf  # noqa: E402
from test_seal import png_bytes, ring_rgba  # noqa: E402

from todoku import api  # noqa: E402
from todoku.pdf.render import fontmap  # noqa: E402

FORM = "/Type /XObject /Subtype /Form /BBox [0 0 12 12]"
CROSS = "0 G 1.5 w 2 2 m 10 10 l 2 10 m 10 2 l S"  # 線で描いた印（書体は要らない）
ZADB = "q 0 g BT /ZaDb 10 Tf 1.5 2 Td (4) Tj ET Q"  # 書体の字で描いた印
COVER = ("<< /Type /Annot /Subtype /Square /Rect [95 695 120 720] /F 4 /AP << /N 7 0 R >> >>",
         pdfmini.stream("1 g 0 0 25 25 re f", "/Type /XObject /Subtype /Form /BBox [0 0 25 25]"))


def _fill(tmp_path, blob: bytes, values: dict) -> dict:
    src = tmp_path / "a.pdf"
    src.write_bytes(blob)
    api._WINDOWS.clear()
    return api.fill(str(src), values, str(tmp_path / "b.pdf"), overwrite=True)


def _box(on: str = CROSS, content: str = "", state: str = "Off", cover: bool = False, flags: int = 4, res: str = "") -> bytes:
    """チェックボックス 1 つ（a）。on：「入」の見た目の描画。cover：部品の後に描かれる、白く塗った注釈。"""
    objs = [f"<< /Type /Annot /Subtype /Widget /FT /Btn /T (a) /Rect [100 700 112 712] /F {flags} /V /{state} /AS /{state} "
            "/AP << /N << /Yes 3 0 R /Off 4 0 R >> >> >>",
            pdfmini.stream(on, FORM + res), pdfmini.stream("", FORM),
            "<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>"]
    annots = "/Annots [2 0 R]"
    if cover:
        objs += list(COVER)
        annots = "/Annots [2 0 R 6 0 R]"
    return pdfmini.build(content, "<< >>", objs, annots=annots, acroform="/AcroForm << /Fields [2 0 R] >>")


def _radio(second: str = CROSS) -> bytes:
    """ラジオの組（r）：A と B。second：B の「入」の見た目。"""
    kid = "<< /Type /Annot /Subtype /Widget /Parent 2 0 R /Rect [{r}] /F 4 /AS /Off /AP << /N << /{s} {n} 0 R /Off 7 0 R >> >> >>"
    objs = ["<< /FT /Btn /Ff 49152 /T (r) /V /Off /Kids [3 0 R 4 0 R] >>",
            kid.format(r="100 700 112 712", s="A", n=5), kid.format(r="140 700 152 712", s="B", n=6),
            pdfmini.stream(CROSS, FORM), pdfmini.stream(second, FORM), pdfmini.stream("", FORM)]
    return pdfmini.build("", "<< >>", objs, annots="/Annots [3 0 R 4 0 R]", acroform="/AcroForm << /Fields [2 0 R] >>")


def _told(r: dict) -> list[str]:
    return [w for w in r["warnings"] if "明るさ" in w or "見え" in w]


def _diff(r: dict, fid: str) -> str:
    return next(d["read_back"] for d in r["read_back_differences"] if d["id"] == fid)


def _mm(text: str) -> list[float]:
    return [float(v) for v in re.search(r"左 ([\d.]+)・上 ([\d.]+)・右 ([\d.]+)・下 ([\d.]+) mm", text).groups()]


# ====================================================================== チェック・ラジオの印
def test_a_plain_check_mark_is_not_told(tmp_path):
    r = _fill(tmp_path, _box(), {"a": True})
    assert r["read_back_ok"] and r["proof_ok"] and not _told(r), r


def test_a_check_whose_appearance_draws_nothing_is_a_read_back_difference(tmp_path):
    # 「入」の見た目が空：値は入るが、ページには何も出ない
    r = _fill(tmp_path, _box(on=""), {"a": True})
    assert r["proof_ok"] and r["read_back_ok"] is False
    d = _diff(r, "a")
    assert "印は入りましたが、描いても見えません" in d and "view_page" in d
    x0, y0, x1, y1 = _mm(d)  # 部品の四角（100〜112 pt・ページの上から 130〜142 pt）
    assert abs(x0 - 35.3) < 0.2 and abs(x1 - 39.5) < 0.2 and abs(y0 - 45.9) < 0.2 and abs(y1 - 50.1) < 0.2


def test_a_white_mark_on_white_is_a_read_back_difference(tmp_path):
    r = _fill(tmp_path, _box(on=CROSS.replace("0 G", "1 G")), {"a": True})
    assert r["read_back_ok"] is False and "描いても見えません" in _diff(r, "a")


def test_a_black_mark_on_a_black_band_is_a_read_back_difference(tmp_path):
    r = _fill(tmp_path, _box(content="0 g 90 690 40 40 re f"), {"a": True})
    assert r["read_back_ok"] is False and "同じ色の下地" in _diff(r, "a")


def test_a_black_frame_around_the_mark_is_not_a_black_ground(tmp_path):
    """外の四角と、逆向きの内の四角で描いた黒い枠（巻き数の決まりで、中は塗らない）の中の印は、白い下地の上の印。"""
    r = _fill(tmp_path, _box(content="0 g 90 690 40 40 re 129 691 -38 38 re f"), {"a": True})
    assert r["read_back_ok"] and r["proof_ok"] and not _told(r), r


@pytest.mark.parametrize("blob", [_box(on=CROSS.replace("0 G", "0.8 G")), _box(content="0 0 0.6 rg 90 690 40 40 re f")],
                         ids=["薄い灰色の印", "濃い青の下地"])
def test_a_mark_that_is_hard_to_see_is_told_with_its_place(tmp_path, blob):
    r = _fill(tmp_path, blob, {"a": True})
    assert r["read_back_ok"], r
    told = _told(r)
    assert len(told) == 1 and told[0].startswith("a：書いた印と下地の明るさの差が小さく") and "1ページ" in told[0]
    assert float(re.search(r"比 ([\d.]+):1", told[0])[1]) < 3.0
    assert abs(_mm(told[0])[0] - 35.3) < 0.2


def test_a_mark_under_something_drawn_later_is_a_read_back_difference(tmp_path):
    # 部品の後に描かれる、白く塗った注釈の下：書く前と書いた後で、絵が変わらない
    assert _fill(tmp_path, _box(cover=True), {"a": True})["read_back_ok"] is False
    r = _fill(tmp_path, _box(cover=True), {"a": True})
    assert "後から描かれるものの下" in _diff(r, "a")


def test_a_mark_that_was_already_on_is_not_looked_at(tmp_path):
    # 書く前から「入」の部品：絵は変わらなくてよい（とどくが入れた印ではない）
    r = _fill(tmp_path, _box(on="", state="Yes"), {"a": True})
    assert r["read_back_ok"] and not _told(r), r


def test_taking_a_mark_out_is_not_looked_at(tmp_path):
    r = _fill(tmp_path, _box(state="Yes"), {"a": False})
    assert r["read_back_ok"] and not _told(r), r


def test_the_chosen_radio_button_is_looked_at(tmp_path):
    src = tmp_path / "r.pdf"
    src.write_bytes(_radio())
    f = next(x for x in api.fields(str(src)) if x["kind"] == "radio")
    first, second = [o["value"] if isinstance(o, dict) else o for o in f["options"]][:2]
    r = _fill(tmp_path, _radio(), {f["id"]: second})
    assert r["read_back_ok"] and not _told(r), r
    # B の「入」の見た目が何も描かない：B を選ぶと違い、A を選ぶと、見るのは A だけ
    r = _fill(tmp_path, _radio(second=""), {f["id"]: second})
    assert r["read_back_ok"] is False
    x0 = _mm(_diff(r, f["id"]))[0]
    assert abs(x0 - 140 / 72 * 25.4) < 0.2  # B の部品の位置
    r = _fill(tmp_path, _radio(second=""), {f["id"]: first})
    assert r["read_back_ok"] and not _told(r), r


# ====================================================================== 印の字の形がこのパソコンにないとき
def test_a_mark_whose_font_is_missing_is_not_judged_by_its_brightness(tmp_path, monkeypatch):
    """描画は、形の分からない字を灰色の四角で描く：その明るさは印の明るさではないので、「見えにくい」とは言わない。"""
    res = " /Resources << /Font << /ZaDb 5 0 R >> >>"
    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    r = _fill(tmp_path, _box(on=ZADB, res=res), {"a": True})
    assert r["read_back_ok"] and not [w for w in _told(r) if "明るさ" in w], r
    # 後から描かれるものの下：違いとは決めずに、確かめられなかったことを知らせる
    r = _fill(tmp_path, _box(on=ZADB, res=res, cover=True), {"a": True})
    assert r["read_back_ok"], r
    told = [w for w in r["warnings"] if "印が見えるかを" in w]
    assert len(told) == 1 and "a" in told[0] and "確かめられませんでした" in told[0] and "view_page" in told[0]


def test_a_mark_drawn_as_a_placeholder_is_not_judged_by_its_brightness(tmp_path, monkeypatch):
    """代わりの書体に印の字がない（灰色の四角になる）ときも、明るさでは決めない。"""
    monkeypatch.setattr(fontmap.SystemFace, "lookup", lambda self, text: None)
    r = _fill(tmp_path, _box(on=ZADB, res=" /Resources << /Font << /ZaDb 5 0 R >> >>"), {"a": True})
    assert r["read_back_ok"] and not [w for w in _told(r) if "明るさ" in w], r


@pytest.mark.parametrize("ground, mark, told", [("0 g", "0 g", "違い"), ("0.25 g", "0 g", "知らせ"), ("1 g", "0 g", ""), ("0 g", "1 g", "")])
def test_a_mark_drawn_as_a_placeholder_is_judged_by_its_own_colour(tmp_path, monkeypatch, ground, mark, told):
    """印の字の形がなくても、印の色は分かる：下地がどこも印と同じ色なら違い、どこも明るさの比が 3:1 未満なら知らせる。"""
    monkeypatch.setattr(fontmap.SystemFace, "lookup", lambda self, text: None)
    on = f"q {mark} BT /ZaDb 10 Tf 1.5 2 Td (4) Tj ET Q"
    r = _fill(tmp_path, _box(on=on, content=f"{ground} 90 690 40 40 re f", res=" /Resources << /Font << /ZaDb 5 0 R >> >>"), {"a": True})
    if told == "違い":
        assert r["read_back_ok"] is False and "下地が、印と同じ色" in _diff(r, "a") and "描いても見えません" in _diff(r, "a")
        assert abs(_mm(_diff(r, "a"))[0] - 100 / 72 * 25.4) < 0.2
    elif told == "知らせ":
        dim = [w for w in r["warnings"] if "明るさ" in w]
        assert r["read_back_ok"] and len(dim) == 1 and dim[0].startswith("a：書いた印") and "view_page" in dim[0], r
    else:
        assert r["read_back_ok"] and not _told(r), r


def test_a_placeholder_mark_after_many_letters_without_a_shape_is_still_judged(tmp_path, monkeypatch):
    """ページの字の書体もないとき：描画が控える四角は、後から描いた 400 個（書いた印は、ページの中身の後に描かれる）。"""
    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    lines = "".join(f"BT /F1 6 Tf 50 {600 - 8 * i} Td ({'a' * 90}) Tj ET " for i in range(5))  # 形のない字 450 個
    objs = ["<< /Type /Annot /Subtype /Widget /FT /Btn /T (a) /Rect [100 700 112 712] /F 4 /V /Off /AS /Off "
            "/AP << /N << /Yes 3 0 R /Off 4 0 R >> >> >>",
            pdfmini.stream(ZADB, FORM + " /Resources << /Font << /ZaDb 5 0 R >> >>"), pdfmini.stream("", FORM),
            "<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>"]
    blob = pdfmini.build("0 g 90 690 40 40 re f " + lines, "<< /Font << /F1 5 0 R >> >>", objs, annots="/Annots [2 0 R]",
                         acroform="/AcroForm << /Fields [2 0 R] >>")
    r = _fill(tmp_path, blob, {"a": True})
    assert r["read_back_ok"] is False and "下地が、印と同じ色" in _diff(r, "a"), r

# ====================================================================== 印刷用のページに入れた印
def _flat(band: bytes = b"") -> bytes:
    font, extra, codes = jp_font_objs("申請書新規変更□", 5)
    c = (band + b"0 g BT /F1 14 Tf 70 700 Td " + jp_text(codes, "申請書") + b" Tj ET\nBT /F1 10 Tf 70 560 Td "
         + jp_text(codes, "□新規") + b" Tj 80 0 Td " + jp_text(codes, "□変更") + b" Tj ET\n")
    return make_pdf(c, {"F1": font}, extra)


def _flat_fill(tmp_path, band: bytes = b"") -> tuple[dict, str]:
    src = tmp_path / "a.pdf"
    src.write_bytes(_flat(band))
    fid = next(f["id"] for f in api.fields(str(src)) if f.get("kind") == "checkbox")
    return api.fill(str(src), {fid: True}, str(tmp_path / "b.pdf"), overwrite=True), fid


def test_a_mark_written_into_a_printed_box(tmp_path):
    r, _fid = _flat_fill(tmp_path)
    assert r["read_back_ok"] and not [w for w in _told(r) if "明るさ" in w], r


def test_a_mark_written_into_a_printed_box_on_a_black_band(tmp_path):
    r, fid = _flat_fill(tmp_path, b"0 g 60 550 30 30 re f\n")
    # 印の字を描く書体のないパソコンでは、違いとは決めずに、確かめられなかったことを知らせる
    unsure = [w for w in r["warnings"] if "印が見えるかを" in w and fid in w]
    if not unsure:
        assert r["read_back_ok"] is False and "描いても見えません" in _diff(r, fid)



def test_a_mark_in_a_printed_box_whose_letter_has_no_shape_either(tmp_path, monkeypatch):
    """印刷された □ の字の形もないとき：描画が □ の位置に描く灰色の枠を、印の下地の色に数えない（描画が足した印で、書類の色ではない）。"""
    monkeypatch.setattr(fontmap, "system_face", lambda kind: None)
    r, fid = _flat_fill(tmp_path, b"0 g 60 550 30 30 re f\n")
    assert r["read_back_ok"] is False and "下地が、印と同じ色" in _diff(r, fid), r
    r, fid = _flat_fill(tmp_path)
    assert r["read_back_ok"] and not _told(r), r


# ====================================================================== 字を囲む ○
def _choice(band: str = "") -> bytes:
    """「Status　Yes　No」の行。band：字の前に描く塗り（字の下地）。"""
    rows = [("Status", 70, 700), ("Yes", 200, 700), ("No", 260, 700)]
    code = {ch: i + 1 for i, ch in enumerate(sorted({ch for t, _x, _y in rows for ch in t}))}
    text = "\n".join(f"BT /F1 12 Tf {x} {y} Td <{''.join(f'{code[ch]:04X}' for ch in t)}> Tj ET" for t, x, y in rows)
    objs = [pdfmini.type0("TestFont", 3, "/W [" + " ".join(f"{c} [500]" for c in code.values()) + "]"),
            pdfmini.stream(pdfmini.tounicode({c: ch for ch, c in code.items()}))]
    return pdfmini.build(band + "\n0 g\n" + text + "\n", "<< /Font << /F1 2 0 R >> >>", objs)


def test_a_circle_on_white_is_not_told(tmp_path):
    r = _fill(tmp_path, _choice(), {"@c1": {"circle": "Yes"}})
    assert r["read_back_ok"] and r["proof_ok"] and not _told(r), r


def test_a_circle_on_a_black_band_is_a_read_back_difference(tmp_path):
    r = _fill(tmp_path, _choice("0 g 180 680 60 50 re f"), {"@c1": {"circle": "Yes"}})
    assert r["read_back_ok"] is False
    d = next(x for x in r["read_back_differences"] if "○" in x["written"])
    assert "○は入りましたが、描いても見えません" in d["read_back"] and "Yes" in d["written"]
    box = r["circled"][0]["box_mm"]
    assert all(abs(a - b) < 0.1 for a, b in zip(_mm(d["read_back"]), box))


def test_a_circle_on_a_dark_band_is_told(tmp_path):
    r = _fill(tmp_path, _choice("0 0 0.6 rg 180 680 60 50 re f"), {"@c1": {"circle": "Yes"}})
    assert r["read_back_ok"], r
    assert len(_told(r)) == 1 and "書いた○と下地の明るさの差が小さく" in _told(r)[0]


# ====================================================================== 印影の画像
def _square(px: tuple[int, int, int, int], n: int = 24) -> bytes:
    return png_bytes(n, n, 6, [bytes(px) * n for _ in range(n)])


def _seal(tmp_path, png: bytes, extra: str = "") -> dict:
    p = tmp_path / "seal.png"
    p.write_bytes(png)
    return _fill(tmp_path, _choice(extra), {"@1": {"seal": str(p), "page": 1, "box": [20, 80, 40, 100], "consent": True}})


def test_a_seal_image_on_white_is_not_told(tmp_path):
    r = _seal(tmp_path, ring_rgba()[0])
    assert r["read_back_ok"] and r["proof_ok"] and not _told(r), r


def test_a_pale_seal_image_is_told(tmp_path):
    r = _seal(tmp_path, _square((255, 225, 225, 255)))
    assert r["read_back_ok"], r
    told = _told(r)
    assert len(told) == 1 and "書いた画像と下地の明るさの差が小さく" in told[0]
    x0, y0, x1, y1 = _mm(told[0])
    assert 19.5 <= x0 < x1 <= 40.5 and 79.5 <= y0 < y1 <= 100.5


def test_a_seal_image_of_the_color_of_the_page_is_a_read_back_difference(tmp_path):
    r = _seal(tmp_path, _square((255, 255, 255, 255)))
    assert r["read_back_ok"] is False
    d = next(x for x in r["read_back_differences"] if "画像は入りましたが" in x["read_back"])
    assert "描いても見えません" in d["read_back"]
