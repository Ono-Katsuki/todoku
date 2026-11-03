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
