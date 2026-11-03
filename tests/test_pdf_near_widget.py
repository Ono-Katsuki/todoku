"""配置から見つけた所が、元からあるフォームの欄（AcroForm）に重なる・すぐ隣にあるとき：消さずに「？」にして、欄の名前を添える。
欄の中に収まる所（欄の外に 1 字も入らない）だけを、その欄そのものとして外す。"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from pdfmini import build, stream, tounicode, type0  # noqa: E402

from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.window import build as window  # noqa: E402


def _form_pdf(rows, widgets) -> bytes:
    """rows: (字, x, y)（12pt。ASCII と記号は半角、ほかは全角）。widgets: (名前, 種類 Tx/Btn, [x0 y0 x1 y1])。"""
    chars = sorted({ch for t, _x, _y in rows for ch in t})
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    w = " ".join(f"{code[ch]} [500]" for ch in chars if ord(ch) < 0x2E80)
    parts = []
    for t, x, y in rows:
        hexs = "".join(f"{code[ch]:04X}" for ch in t)
        parts.append(f"BT /F1 12 Tf {x} {y} Td <{hexs}> Tj ET")
    objs = [type0("TestFont", 3, f"/W [{w}]" if w else ""), stream(tounicode({c: ch for ch, c in code.items()}))]
    first = len(objs) + 2
    refs = []
    for k, (nm, ft, rect) in enumerate(widgets):
        extra = "/V /Off /AS /Off" if ft == "Btn" else "/DA (/Helv 10 Tf 0 g)"
        objs.append(f"<< /Type /Annot /Subtype /Widget /FT /{ft} /T ({nm}) /Rect [{' '.join(map(str, rect))}] /P 1 0 R /F 4 {extra} >>")
        refs.append(f"{first + k} 0 R")
    return build("\n".join(parts), "<< /Font << /F1 2 0 R >> >>", objs, annots=f"/Annots [{' '.join(refs)}]",
                 acroform=f"/AcroForm << /Fields [{' '.join(refs)}] >>")


def test_box_before_a_real_text_field_is_kept_as_agent():
    # 「□ その他（［字の欄］）」：□ は消さない（「？」のチェック）。種類が違うので alternative_to にはせず、理由に欄の名前
    # □ 50〜56・他 56〜68・（ 68〜80・空き 80〜116・）116〜128：字の欄は 70〜114（□ から 14pt：欄の高さ 16pt の内）
    rows = [("□他（　　　）", 50, 700), ("□はい", 50, 670)]
    blob = _form_pdf(rows, [("other", "Tx", [70, 697, 114, 713])])
    pg = window(PdfDocument(blob, "a.pdf"))["pages"][0]
    by = {f["id"]: f for f in pg["fields"]}
    assert "other" in by
    near = [f for f in pg["fields"] if f["source"] == "layout" and "フォームの欄 other" in f.get("why", "")]
    boxes = [f for f in near if f["kind"] == "checkbox"]
    assert boxes, pg["fields"]
    assert all(f["agent"] and not f.get("alternative_to") for f in boxes), boxes
    text = "\n".join(pg["reading"])
    assert f"☐［{boxes[0]['id']}？］" in text, text
    # 欄から離れた「□はい」は、これまでどおり規則のチェック
    yes = [f for f in pg["fields"] if f["source"] == "layout" and f["kind"] == "checkbox" and f not in boxes]
    assert yes and not yes[0].get("agent"), yes
    # かっこの中の空き（欄に重なる字の所）は、字の欄と同じ所の別の読み方
    for f in near:
        if f["kind"] != "checkbox":
            assert f["alternative_to"] == "other" and f["agent"], f


def test_place_inside_a_real_field_is_the_field_itself():
    # 字の欄の中に収まる □（欄の外に 1 字も入らない）：その欄そのものなので、別の所として出さない
    rows = [("申込み", 50, 700), ("□", 120, 700)]
    blob = _form_pdf(rows, [("chk", "Btn", [118, 697, 128, 713])])
    pg = window(PdfDocument(blob, "a.pdf"))["pages"][0]
    lay = [f for f in pg["fields"] if f["source"] == "layout" and f["kind"] == "checkbox"]
    assert not lay, lay
    assert any(f["id"] == "chk" for f in pg["fields"])


# 電話番号［tel］（　　　　） / 氏名［name］　　年　　月　　日（欄 name は「氏名」と「年」の間の空きに置かれている）
_GAP_ROWS = [("電話番号", 50, 700), ("（　　　　）", 176, 700), ("氏名", 50, 660), ("　　年　　月　　日", 176, 660)]
_GAP_WIDGETS = [("tel", "Tx", [100, 697, 175, 713]), ("name", "Tx", [80, 657, 175, 673])]


def _gap_page():
    return window(PdfDocument(_form_pdf(_GAP_ROWS, _GAP_WIDGETS), "a.pdf"))["pages"][0]


def test_widget_in_the_gap_of_a_line_is_labelled_by_that_line():
    # 欄が、同じ高さの行の字の間に置かれているとき：見出しは、その行の欄より前の字（上の行ではない）
    by = {f["id"]: f for f in _gap_page()["fields"]}
    assert by["name"]["label"] == "氏名", by["name"]
    assert by["tel"]["label"] == "電話番号", by["tel"]


def test_gap_token_keeps_the_measured_spaces():
    # 「？」の空きの印に、行の空白（測った 24pt の空き＝12pt の字 2 つぶん）をそのまま入れる
    text = "\n".join(_gap_page()["reading"])
    assert "年［p1.a2？　　］月［p1.a3？　　］日" in text, text


def test_overlap_of_one_em_is_alternative_beside_is_not():
    lay = [f for f in _gap_page()["fields"] if f["source"] == "layout"]
    # 欄 name に 1 字ぶん以上重なる空き：同じ所の別の読み方
    over = [f for f in lay if f.get("alternative_to") == "name"]
    assert len(over) == 1 and over[0]["agent"] and "フォームの欄 name に重なる" in over[0]["why"], lay
    # 欄 tel のすぐ隣のかっこの中（重ならない・字の形で決まる空欄）：規則の空欄のまま。alternative_to にしない
    paren = [f for f in lay if "電話番号" in f["label"]]
    assert len(paren) == 1 and not paren[0].get("agent") and not paren[0].get("alternative_to"), paren
    assert "フォームの欄 tel" in paren[0]["reason"], paren


def test_beside_a_widget_is_agent_without_alternative():
    # 字の形で決まらない空き（測った空白）が、欄のすぐ隣にあるだけ：「？」で理由に欄の名前。alternative_to にはしない
    # 氏 50〜62・名 62〜74・空白 74〜122・様 122〜134：欄は 136〜200（空きから 14pt：欄の高さ 16pt の内）
    rows = [("氏名　　　　様", 50, 700)]
    pg = window(PdfDocument(_form_pdf(rows, [("nm", "Tx", [136, 697, 200, 713])]), "a.pdf"))["pages"][0]
    near = [f for f in pg["fields"] if f["source"] == "layout" and "フォームの欄 nm" in (f.get("why") or "")]
    assert near, pg["fields"]
    for f in near:
        assert f["agent"] and not f.get("alternative_to") and "すぐ隣" in f["why"], f


def test_writing_both_the_widget_and_its_alternative_is_refused():
    import pytest

    from todoku.pdf.fill import FillError, fill

    blob = _form_pdf(_GAP_ROWS, _GAP_WIDGETS)
    fields = _gap_page()["fields"]
    alt = next(f["id"] for f in fields if f.get("alternative_to") == "name")
    with pytest.raises(FillError) as e:
        fill(PdfDocument(blob, "a.pdf"), {"name": "山田", alt: "山田"})
    assert "name" in str(e.value) and alt in str(e.value) and "どちらか1つ" in str(e.value), str(e.value)
    # 一方だけなら書ける。重ならない隣の所（かっこの中）は、欄 tel と一緒に書ける
    paren = next(f["id"] for f in fields if "電話番号" in f["label"] and f["source"] == "layout")
    assert fill(PdfDocument(blob, "a.pdf"), {"name": "山田", "tel": "03", paren: "1234"}).patches


def test_say_to_user_names_the_widget_by_its_own_line(tmp_path):
    from todoku import api

    src = tmp_path / "a.pdf"
    src.write_bytes(_form_pdf(_GAP_ROWS, _GAP_WIDGETS))
    r = api.fill(str(src), {"name": "山田"}, str(tmp_path / "b.pdf"))
    assert "氏名：山田" in r["say_to_user"] and "電話番号 （ ）：山田" not in r["say_to_user"], r["say_to_user"]
