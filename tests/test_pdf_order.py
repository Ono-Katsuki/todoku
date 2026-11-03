"""PDF の読む窓：字の並び（右から左へ書く字・子音の前に描く母音・縦書き・ルビ）。tests/test_pdf_scripts.py の続き。"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import pdfmini  # noqa: E402
from test_pdf_scripts import flags, glyphs_pdf, reading, window  # noqa: E402

from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.window import build  # noqa: E402


# ====================================================================== 4. 右から左へ書く字
def _visual(logical: str) -> list[str]:
    """試しの見た目の順（左から右）：論理順に UAX #9 の L2 を一度かけたもの（字ごと）。"""
    from todoku.pdf.bidi import bidi_class, levels, reorder

    texts = list(logical)
    cls = [bidi_class(t) if not t.isspace() else "N" for t in texts]
    strong = [c for c in cls if c in ("L", "R")]
    rtl = bool(strong) and strong[0] == "R"  # 論理順の最初の強い字
    return reorder(texts, levels(cls, rtl))


def test_bidi_round_trip():
    from todoku.pdf.bidi import visual_to_logical

    for logical in ["محمد علي", "رقم الهاتف 0312345678", "שלום עולם", "العنوان: شارع 12", "تاريخ 2026-11-03"]:
        vis = _visual(logical)
        order, rtl = visual_to_logical(vis)
        assert rtl
        assert "".join(vis[i] for i in order) == logical, logical


def test_bidi_ltr_paragraph_with_arabic_name():
    from todoku.pdf.bidi import visual_to_logical

    vis = list("Name: ") + list("دمحم")
    order, rtl = visual_to_logical(vis)
    assert not rtl
    assert "".join(vis[i] for i in order) == "Name: محمد"


def test_rtl_line_reads_in_logical_order_with_numbers_ltr():
    logical = "رقم 123"
    vis = _visual(logical)
    assert "".join(vis).startswith("123") and "".join(vis).endswith("مقر")
    g = [(ch, 40 + 7 * k, 800) for k, ch in enumerate(vis)]
    blob = glyphs_pdf(g, {ch: 583 for ch in set(vis)})
    assert reading(blob) == [logical]


def test_rtl_table_label_is_right_cell():
    # 右から左の表：右のマスに見出し「الاسم」、左のマスが空
    lines = ["40 700 m 340 700 l S", "40 680 m 340 680 l S", "40 680 m 40 700 l S", "240 680 m 240 700 l S",
             "340 680 m 340 700 l S"]
    word = _visual("الاسم")
    g = [(ch, 300 + 7 * k, 685) for k, ch in enumerate(word)]
    blob = glyphs_pdf(g, {ch: 583 for ch in set(word)}, extra="\n".join(lines))
    w = window(blob)
    fields = w["pages"][0]["fields"]
    left = [f for f in fields if abs(f["rect"][0] - 40) < 2]
    assert left and left[0]["label"] == "الاسم"
    assert w["pages"][0]["reading"][0].startswith("الاسم")


def test_anchor_before_for_rtl_label():
    from todoku.pdf.anchor import resolve_write

    word = _visual("الاسم")
    g = [(ch, 400 + 7 * k, 700) for k, ch in enumerate(word)]
    blob = glyphs_pdf(g, {ch: 583 for ch in set(word)}, extra="40 600 m 500 600 l S")  # ページの字の左端を決める線
    sd = resolve_write(PdfDocument(blob, "t.pdf"), {"after": "الاسم", "value": "محمد", "page": 1, "where": "before"})
    assert sd["rect"][2] <= 400 and sd["rect"][2] - sd["rect"][0] > 50


# ====================================================================== 5. 子音の前に描く母音
def test_bengali_split_vowel_recomposes():
    # কো を、ে（左）・ক・া（右）の3つのグリフで描く
    blob = glyphs_pdf([("ে", 40, 800), ("ক", 46, 800), ("া", 54, 800)])
    assert reading(blob) == ["কো"]


def test_recomposed_vowel_passes_reading_audit():
    from todoku.pdf.audit import check_reading

    doc = PdfDocument(glyphs_pdf([("ে", 40, 800), ("ক", 46, 800), ("া", 54, 800)]), "t.pdf")
    assert check_reading(doc, build(doc))[0].ok


def test_tamil_split_vowel_recomposes():
    blob = glyphs_pdf([("ெ", 40, 800), ("க", 47, 800), ("ா", 55, 800)])
    assert reading(blob) == ["கொ"]


def test_devanagari_i_after_conjunct():
    # क्षि：ि（左）・क्ष（1つのグリフ）
    blob = glyphs_pdf([("ि", 40, 800), ("क्ष", 45, 800)])
    assert reading(blob) == ["क्षि"]


def test_myanmar_e_moves_after_consonant():
    blob = glyphs_pdf([("ေ", 40, 800), ("က", 46, 800)])
    assert reading(blob) == ["ကေ"]


def test_orphan_dependent_sign_is_flagged():
    blob = glyphs_pdf([("ा", 40, 800), ("क", 60, 800)])
    assert "前の字のない結合の記号" in flags(blob)


# ====================================================================== 6. 縦書き
def test_stacked_vertical_columns_read_right_to_left():
    # 横書きの書体で、1字ずつ下へ積んだ縦書き（2列。右の列から）
    g = [(ch, 300, 800 - 12 * k) for k, ch in enumerate("申請者の")] + [(ch, 280, 800 - 12 * k) for k, ch in enumerate("住所欄")]
    text = " ".join(reading(glyphs_pdf(g)))
    assert "申請者の" in text and "住所欄" in text
    assert text.index("申請者の") < text.index("住所欄")


# ====================================================================== 8. ルビ
def test_ruby_is_attached_to_base():
    g = [("葛", 100, 780), ("飾", 114, 780), ("区", 128, 780)]
    ruby = [(ch, 100 + 7 * k, 795) for k, ch in enumerate("かつしか")]
    chars = sorted({t for t, _x, _y in g + ruby})
    code = {ch: i + 1 for i, ch in enumerate(chars)}
    content = "\n".join(f"BT /F1 14 Tf {x} {y} Td <{code[t]:04X}> Tj ET" for t, x, y in g)
    content += "\n" + "\n".join(f"BT /F1 7 Tf {x} {y} Td <{code[t]:04X}> Tj ET" for t, x, y in ruby)
    objs = [pdfmini.type0("TestFont", 3), pdfmini.stream(pdfmini.tounicode({c: ch for ch, c in code.items()}))]
    blob = pdfmini.build(content, "<< /Font << /F1 2 0 R >> >>", objs)
    assert reading(blob) == ["葛飾（かつしか）区"]
    from todoku.pdf.audit import check_reading

    doc = PdfDocument(blob, "t.pdf")
    assert check_reading(doc, build(doc))[0].ok
    # つないだ所は、読む B で位置を示して知らせる（エージェントが、ルビか注の字かを確かめる）
    fl = {f["title"]: f for f in build(doc)["pages"][0]["read_flags"]}
    r = fl["ルビとして基の字につないだ字"]
    assert r["grade"] == "B" and len(r["where"]) == 1 and "かつしか" in r["where"][0] and "左 100" in r["where"][0]
