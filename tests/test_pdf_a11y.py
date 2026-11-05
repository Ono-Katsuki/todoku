"""書いた字を読み上げに届ける：タグのある PDF の構造の木に足す（structadd）・値の控えは注釈にしない・○とチェックの印が示す字・/TU を整える。

実物の読み手（プレビューの読み上げの木・Chromium・pdf.js・PDFKit・pdftotext）では別に測った。ここでは、ファイルの形と、とどく自身の読み書き・証明を確かめる。
"""

from __future__ import annotations

import re

import pytest

from pdfmini import build, stream
from todoku.pdf import pagedraw, structadd
from todoku.pdf.content import page_content
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import FillError, appearance_readback, fill
from todoku.pdf.form import read_fields
from todoku.pdf.objects import Ref
from todoku.pdf.proof import prove
from todoku.pdf.window import build as build_window

HELV = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
PAGE = ("BT /F1 12 Tf 60 700 Td (Name) Tj ET\n0.5 w 50 690 300 30 re S\n"
        "BT /F1 12 Tf 60 640 Td (Date) Tj ET\n0.5 w 50 630 300 30 re S\n")
TAGGED_PAGE = ("/P <</MCID 0>> BDC BT /F1 12 Tf 60 700 Td (Name) Tj ET EMC\n0.5 w 50 690 300 30 re S\n"
               "/P <</MCID 1>> BDC BT /F1 12 Tf 60 640 Td (Date) Tj ET EMC\n0.5 w 50 630 300 30 re S\n")


OPTIONS = "BT /F1 12 Tf 60 700 Td (Male) Tj 80 0 Td (Female) Tj ET\n"


def plain() -> bytes:
    return build(PAGE, "<< /Font << /F1 2 0 R >> >>", [HELV])


def tagged(elem_extra: str = "", nums: str = "/Nums [0 8 0 R]", page_extra: str = "/StructParents 0", content: str = TAGGED_PAGE) -> bytes:
    objs = [HELV,
            "<< /Type /StructTreeRoot /K 4 0 R /ParentTree 7 0 R /ParentTreeNextKey 1 >>",
            "<< /Type /StructElem /S /Document /P 3 0 R /K [5 0 R 6 0 R] >>",
            f"<< /Type /StructElem /S /P /P 4 0 R /Pg 1 0 R /K 0 {elem_extra} >>",
            "<< /Type /StructElem /S /P /P 4 0 R /Pg 1 0 R /K 1 >>",
            f"<< {nums} >>",
            "[5 0 R 6 0 R]"]
    return build(content, "<< /Font << /F1 2 0 R >> >>", objs, annots=page_extra,
                 acroform="/StructTreeRoot 3 0 R /MarkInfo << /Marked true >>")


def spots(blob: bytes) -> list[dict]:
    return [fd for pg in build_window(PdfDocument(blob))["pages"] for fd in pg["fields"]]


def write(blob: bytes, values: dict):
    res = fill(PdfDocument(blob), values)
    pr = prove(blob, res.blob)
    assert pr.ok, pr.unexpected
    assert appearance_readback(PdfDocument(res.blob), res.patches) == []
    return res, PdfDocument(res.blob)


def elem_of(doc, mcid: int) -> dict:
    tree = structadd.Tree(doc)
    arr = doc.resolve(tree.entries[doc.pages()[0]["StructParents"]])
    return doc.resolve(arr[mcid])


# ---------------------------------------------------------------------- 値の控えは注釈にしない
def test_record_is_not_an_annotation_but_reads_and_rewrites():
    blob = plain()
    ids = [s["id"] for s in spots(blob)]
    res, doc = write(blob, {ids[0]: "Hanako"})
    assert not doc.resolve(doc.pages()[0].get("Annots"))
    f = next(x for x in read_fields(doc) if x.name == ids[0])
    assert f.value_text() == "Hanako" and f.widgets[0].page == 0 and f.widgets[0].on_screen
    res2, doc2 = write(res.blob, {ids[0]: "Taro"})
    assert next(x for x in read_fields(doc2) if x.name == ids[0]).value_text() == "Taro"
    assert not doc2.resolve(doc2.pages()[0].get("Annots"))
    win = build_window(doc2)
    assert any(fd["id"] == ids[0] and fd.get("value") == "Taro" for pg in win["pages"] for fd in pg["fields"])


def test_proof_flags_record_put_back_in_annots():
    """注釈に戻した（読み手が 2 度読む）ファイルは、証明が通さない。"""
    from todoku.pdf import fill as F

    blob = plain()
    orig = F.Filler._place

    def place(self, page_no, d, ap, name, state):
        orig(self, page_no, d, ap, name, state)
        self.page_annots.setdefault(page_no, []).append(self.new_fields[-1])

    F.Filler._place = place
    try:
        res = fill(PdfDocument(blob), {spots(blob)[0]["id"]: "x"})
    finally:
        F.Filler._place = orig
    assert any("注釈に入っています" in u for u in prove(blob, res.blob).unexpected)


def test_values_are_appended_in_reading_order():
    blob = plain()
    a, b = spots(blob)[:2]
    upper, lower = (a, b) if a["rect"][1] < b["rect"][1] else (b, a)
    res, doc = write(blob, {lower["id"]: "low", upper["id"]: "up"})
    order = [s["id"] for s in pagedraw.spans(doc, doc.pages()[0])]
    assert order == [upper["id"], lower["id"]]


# ---------------------------------------------------------------------- 構造の木
def test_value_gets_mcid_and_element_next_to_label():
    blob = tagged()
    s = spots(blob)[0]
    res, doc = write(blob, {s["id"]: "Hanako"})
    sp = pagedraw.spans(doc, doc.pages()[0])
    assert sp[0]["mcid"] == 2
    e = elem_of(doc, 2)
    assert e["S"].name == "Span" and e["K"] == 2 and e.get("ActualText") is None and e.get("Alt") is None
    par = doc.resolve(e["P"])
    assert par["S"].name == "P" and res.structure[0].how in ("見出し", "左", "上", "近く")
    assert isinstance(par["K"], list) and par["K"][0] in (0, 1) and isinstance(par["K"][1], Ref) and par["K"][1].num > 8
    assert b"ActualText" not in doc.page_content(doc.pages()[0])
    assert any("構造の木" in w for w in res.warnings)
    # 元の要素はそのまま読める。印刷された字は変わらない
    assert any(g.text == "N" for g in page_content(doc, doc.pages()[0]).glyphs)


def test_two_values_with_same_label_keep_reading_order():
    content = ("/P <</MCID 0>> BDC BT /F1 12 Tf 60 700 Td (Name) Tj 200 0 Td (Kana) Tj ET EMC\n"
               "0.5 w 50 690 150 30 re S\n0.5 w 200 690 150 30 re S\n/P <</MCID 1>> BDC BT /F1 12 Tf 60 500 Td (Other) Tj ET EMC\n")
    blob = tagged(content=content)
    ids = [s["id"] for s in spots(blob)]
    if len(ids) < 2:
        pytest.skip("この形では窓に空きが 2 つ出ない")
    res, doc = write(blob, {ids[1]: "b", ids[0]: "a"})
    assert len({e.id for e in res.structure}) == 2
    kids = doc.resolve(doc.resolve(elem_of(doc, 2)["P"])["K"])
    new = [k for k in kids if isinstance(k, Ref)]
    mc = [doc.resolve(k)["K"] for k in new]
    assert mc == sorted(mc)


def test_rewrite_keeps_structure_and_second_value_gets_next_mcid():
    blob = tagged()
    a, b = spots(blob)[:2]
    res, doc = write(blob, {a["id"]: "one"})
    res2, doc2 = write(res.blob, {a["id"]: "two", b["id"]: "three"})
    sp = {s["id"]: s["mcid"] for s in pagedraw.spans(doc2, doc2.pages()[0])}
    assert sp[a["id"]] == 2 and sp[b["id"]] == 3
    assert elem_of(doc2, 3)["K"] == 3


def test_page_without_structparents_and_missing_key_gets_new_key():
    blob = tagged(page_extra="", content=PAGE)
    res, doc = write(blob, {spots(blob)[0]["id"]: "x"})
    pg = doc.pages()[0]
    assert isinstance(pg["StructParents"], int) and pg["StructParents"] == 1
    tree = structadd.Tree(doc)
    assert doc.resolve(tree.root["ParentTreeNextKey"]) == 2
    e = doc.resolve(doc.resolve(tree.entries[1])[0])
    assert e["S"].name == "Span" and doc.resolve(e["P"])["S"].name == "Document"
    assert res.structure[0].how == "文書の終わり"
    assert any("文書の終わり" in w for w in res.warnings)


def test_parent_tree_with_kids_extends_leaf_and_limits():
    objs = [HELV,
            "<< /Type /StructTreeRoot /K 4 0 R /ParentTree 6 0 R /ParentTreeNextKey 5 >>",
            "<< /Type /StructElem /S /Document /P 3 0 R /K [5 0 R] >>",
            "<< /Type /StructElem /S /P /P 4 0 R /Pg 1 0 R /K 0 >>",
            "<< /Kids [7 0 R] >>",
            "<< /Limits [0 4] /Nums [4 8 0 R] >>",
            "[5 0 R]"]
    blob = build("/P <</MCID 0>> BDC BT /F1 12 Tf 60 700 Td (Name) Tj ET EMC\n0.5 w 50 690 300 30 re S\n", "<< /Font << /F1 2 0 R >> >>",
                 objs, annots="/StructParents 9", acroform="/StructTreeRoot 3 0 R")
    res, doc = write(blob, {spots(blob)[0]["id"]: "x"})
    leaf = doc.get(7)
    assert leaf["Limits"] == [0, 9] and leaf["Nums"][0] == 4 and leaf["Nums"][2] == 9
    assert doc.resolve(doc.get(3)["ParentTreeNextKey"]) == 10


@pytest.mark.parametrize("nums", ["/Nums 5", "/Nums [0]", "/Nums [(a) 8 0 R]"])
def test_broken_parent_tree_is_not_guessed(nums):
    blob = tagged(nums=nums)
    res, doc = write(blob, {spots(blob)[0]["id"]: "x"})
    assert pagedraw.spans(doc, doc.pages()[0])[0]["mcid"] is None
    assert any("足せませんでした" in w for w in res.warnings)
    assert res.structure == []


def test_unclosed_marks_on_page_are_not_nested():
    blob = tagged(content=TAGGED_PAGE + "/P <</MCID 7>> BDC\n")
    res, doc = write(blob, {spots(blob)[0]["id"]: "x"})
    assert pagedraw.spans(doc, doc.pages()[0])[0]["mcid"] is None
    assert any("閉じていません" in w for w in res.warnings)


def test_element_with_actualtext_is_not_entered():
    """中身を説明で置き換える要素の下に入れると、値が読まれない。親の段落の後ろに入れる。"""
    objs = [HELV,
            "<< /Type /StructTreeRoot /K 4 0 R /ParentTree 7 0 R /ParentTreeNextKey 1 >>",
            "<< /Type /StructElem /S /Document /P 3 0 R /K [5 0 R] >>",
            "<< /Type /StructElem /S /P /P 4 0 R /Pg 1 0 R /K [6 0 R] >>",
            "<< /Type /StructElem /S /Span /P 5 0 R /Pg 1 0 R /K 0 /ActualText (Name) >>",
            "<< /Nums [0 8 0 R] >>",
            "[6 0 R]"]
    blob = build("/Span <</MCID 0>> BDC BT /F1 12 Tf 60 700 Td (Name) Tj ET EMC\n0.5 w 50 690 300 30 re S\n", "<< /Font << /F1 2 0 R >> >>",
                 objs, annots="/StructParents 0", acroform="/StructTreeRoot 3 0 R")
    res, doc = write(blob, {spots(blob)[0]["id"]: "x"})
    e = elem_of(doc, 1)
    assert e["P"].num == 4
    kids = doc.resolve(doc.get(4)["K"])
    assert [k.num for k in kids][0] == 5 and kids[1].num == [n for n in doc.xref if isinstance(doc.get(n), dict) and doc.get(n).get("K") == 1][0]


def _tamper(blob: bytes, num: int, old: bytes, new: bytes) -> bytes:
    """足した部分のオブジェクトを同じ長さのまま書き換える。"""
    i = blob.rindex(f"\n{num} 0 obj".encode())
    j = blob.index(b"endobj", i)
    part = blob[i:j]
    assert old in part and len(old) == len(new)
    return blob[:i] + part.replace(old, new) + blob[j:]


def test_proof_rejects_changed_struct_elements():
    blob = tagged()
    res, doc = write(blob, {spots(blob)[0]["id"]: "x"})
    # 新しい要素の番号を変える：足した描画の範囲と合わない
    e = next(n for n in doc.xref if n > 8 and isinstance(doc.get(n), dict) and getattr(doc.get(n).get("Type"), "name", "") == "StructElem")
    bad = _tamper(res.blob, e, b"/K 2", b"/K 9")
    assert not prove(blob, bad).ok


def test_proof_rejects_removed_kid():
    from todoku.pdf.proof import _only_inserted

    assert _only_inserted([0, Ref(5)], [0, Ref(9), Ref(5)], {9})
    assert not _only_inserted([0, Ref(5)], [Ref(9), Ref(5)], {9})
    assert not _only_inserted([0, Ref(5)], [Ref(5), 0, Ref(9)], {9})


def test_empty_tree_gets_document_element_and_parent_tree():
    blob = build(PAGE, "<< /Font << /F1 2 0 R >> >>", [HELV, "<< /Type /StructTreeRoot >>"], acroform="/StructTreeRoot 3 0 R")
    res, doc = write(blob, {spots(blob)[0]["id"]: "x"})
    tree = structadd.Tree(doc)
    key = doc.pages()[0]["StructParents"]
    e = doc.resolve(doc.resolve(tree.entries[key])[0])
    assert doc.resolve(e["P"])["S"].name == "Document" and doc.resolve(e["P"])["P"].num == 3
    assert res.structure[0].how == "文書の終わり"


def test_signed_path_keeps_visible_field_and_artifact_circle(monkeypatch):
    """署名のある文書（may_draw が False）：見える欄のまま。○は飾りの印のまま。番号も透明な字も足さない。"""
    monkeypatch.setattr(pagedraw, "may_draw", lambda d, s=None: False)
    blob = tagged()
    s = spots(blob)[0]
    res = fill(PdfDocument(blob), {s["id"]: "x"})
    doc = PdfDocument(res.blob)
    assert doc.resolve(doc.pages()[0]["Annots"]) and res.structure == []
    assert not pagedraw.records(doc)
    assert any("フォームの欄として足しました" in w for w in res.warnings)
    blob2 = build(OPTIONS, "<< /Font << /F1 2 0 R >> >>", [HELV])
    res2 = fill(PdfDocument(blob2), {"@c": {"circle": "Female"}})
    data = PdfDocument(res2.blob).page_content(PdfDocument(res2.blob).pages()[0])
    assert b"/Artifact <</Type /Layout /TdkCircle" in data and b"TdkS" not in data
    pr = prove(blob2, res2.blob, anchored={"p1.c1": {"circle": "Female"}})
    assert pr.ok, pr.unexpected


# ---------------------------------------------------------------------- ○とチェックの印が示す字
def test_circle_says_what_it_encloses_and_window_does_not_read_it():
    blob = build(OPTIONS, "<< /Font << /F1 2 0 R >> >>", [HELV])
    res, doc = write_circle(blob, "Female")
    sp = pagedraw.spans(doc, doc.pages()[0])
    assert sp[0]["said"] == "○（選択：Female）"
    text = "".join(g.text for g in page_content(doc, doc.pages()[0]).glyphs)
    assert "選択" not in text and "Female" in text


def write_circle(blob, word):
    res = fill(PdfDocument(blob), {"@c": {"circle": word}})
    pr = prove(blob, res.blob, anchored={"p1.c1": {"circle": word}})
    assert pr.ok, pr.unexpected
    doc = PdfDocument(res.blob)
    assert appearance_readback(doc, res.patches) == []
    return res, doc


def test_circle_in_tagged_pdf_has_element():
    blob = tagged(content="/P <</MCID 0>> BDC " + OPTIONS.replace("ET\n", "ET EMC\n"))
    res, doc = write_circle(blob, "Female")
    sp = pagedraw.spans(doc, doc.pages()[0])
    assert sp[0]["mcid"] == 2 and sp[0]["said"] == "○（選択：Female）"
    assert elem_of(doc, 2)["K"] == 2


def test_circle_box_on_unread_page_says_nothing_and_warns():
    blob = build("", "<< >>", [HELV])
    res = fill(PdfDocument(blob), {"@c": {"circle": "X", "page": 1, "box": [20, 20, 40, 30]}})
    doc = PdfDocument(res.blob)
    assert pagedraw.spans(doc, doc.pages()[0])[0]["said"] is None
    assert any("囲んだ字を読めない" in w for w in res.warnings)


def test_visible_text_cannot_hide_behind_the_said_mark():
    """透明でない字のフォームに印をつけても、書類の字として読む。"""
    form = stream("q\n3 Tr\n0 Tr\nBT /F1 12 Tf 10 10 Td (Secret) Tj ET\nQ\n").decode("latin-1").replace(
        "<<", "<< /Type /XObject /Subtype /Form /BBox [0 0 100 100] /TdkSaid true /Resources << /Font << /F1 2 0 R >> >> ", 1)
    blob = build("/TdkS1 Do\n", "<< /Font << /F1 2 0 R >> /XObject << /TdkS1 3 0 R >> >>", [HELV, form])
    doc = PdfDocument(blob)
    assert not pagedraw.verified_refs(doc, doc.pages()[0])
    assert "Secret" in "".join(g.text for g in page_content(doc, doc.pages()[0]).glyphs)


def _check_form() -> bytes:
    return build("BT /F1 12 Tf 80 700 Td (Agree) Tj ET\n0.5 w 60 698 12 12 re S\n", "<< /Font << /F1 2 0 R >> >>", [HELV])


def test_check_mark_written_says_text_and_rewrite_off_removes_it():
    from todoku.pdf import fill as F
    from todoku.pdf.glyphs import appearance_text
    from todoku.pdf.layout import View

    blob = _check_form()
    f = F.Filler(PdfDocument(blob))
    f.add_widget(1, {"id": "c1", "kind": "checkbox", "rect": [60.0, 132.0, 72.0, 146.0]}, True, View(f.doc, f.pages[0]))
    out = f.finish()
    assert prove(blob, out).ok
    doc = PdfDocument(out)
    wd = next(wd for _r, wd in pagedraw.records(doc))
    assert "右の字：Agree" in appearance_text(doc, doc.resolve(wd[pagedraw.KEY]))[0]
    assert "Agree" == "".join(g.text for g in page_content(doc, doc.pages()[0]).glyphs).strip()  # 窓は印の字を読まない
    res2, doc2 = write(out, {"c1": False})
    wd2 = next(wd for _r, wd in pagedraw.records(doc2))
    assert "Agree" not in appearance_text(doc2, doc2.resolve(wd2[pagedraw.KEY]))[0]


def test_check_mark_says_text_beside_it_and_goes_with_the_mark():
    from todoku.pdf import circle
    from todoku.pdf.layout import PageLayout
    from todoku.pdf.content import page_content as pc

    blob = _check_form()
    doc0 = PdfDocument(blob)
    lay = PageLayout(doc0, doc0.pages()[0], pc(doc0, doc0.pages()[0]), 1)
    assert circle.check_said(lay, (60, 130, 72, 144)) == "（右の字：Agree）"
    assert circle.check_said(lay, (60, 500, 72, 514)) is None


def test_check_said_names_both_sides_when_equally_near():
    from todoku.pdf import circle
    from todoku.pdf.layout import PageLayout
    from todoku.pdf.content import page_content as pc

    blob = build("BT /F1 12 Tf 50 700 Td (Yes) Tj 60 0 Td (No) Tj ET\n", "<< /Font << /F1 2 0 R >> >>", [HELV])
    doc0 = PdfDocument(blob)
    lay = PageLayout(doc0, doc0.pages()[0], pc(doc0, doc0.pages()[0]), 1)
    xs = sorted(c.x0 for c in lay.chars)
    x_end_yes = max(c.x1 for c in lay.chars if c.x0 < 100)
    x_no = min(c.x0 for c in lay.chars if c.x0 > 100)
    mid = (x_end_yes + x_no) / 2
    y = lay.chars[0].cy
    assert circle.check_said(lay, (mid - 5, y - 5, mid + 5, y + 5)) == "（左の字：Yes、右の字：No）"
    assert xs


# ---------------------------------------------------------------------- /TU
@pytest.mark.parametrize("raw,want", [("氏名［　］", "氏名"), ("年［　］月　　日", "年 月 日"), ("住　　所", "住 所"), ("［　］", ""),
                                      ("（満［　］歳）", "（満 歳）"), ("a  b\tc", "a b c"), ("東京都" + "　" * 20 + "令和", "東京都 令和"), (None, "")])
def test_clean_label(raw, want):
    assert pagedraw.clean_label(raw) == want


def test_tu_is_clean_and_window_label_keeps_position():
    from todoku.pdf import fill as F

    blob = plain()
    s = spots(blob)[0]
    orig = F.Filler.add_widget

    def add(self, page_no, spot, value, view):
        orig(self, page_no, {**spot, "label": "年［　］月　　日"}, value, view)

    F.Filler.add_widget = add
    try:
        res = fill(PdfDocument(blob), {s["id"]: "x"})
    finally:
        F.Filler.add_widget = orig
    doc = PdfDocument(res.blob)
    wd = next(wd for _r, wd in pagedraw.records(doc))
    from todoku.pdf.objects import decode_text

    assert decode_text(wd["TU"]) == "年 月 日"
    assert next(f for f in read_fields(doc) if f.name == s["id"]).tu == "年［　］月　　日"


# ---------------------------------------------------------------------- pad_mm
def test_right_align_with_large_pad_stays_inside_or_is_refused():
    """余白（pad_mm）を決めたら、その内側に収める：字は小さくなるか、入らなければ断る。枠の端を越えない。"""
    from todoku.pdf.fill import text_width

    blob = plain()
    s = spots(blob)[0]
    text = "Hanako Yamada Hanako Yamada Hanako Yamada Hanako Yamada"
    res = fill(PdfDocument(blob), {s["id"]: {"value": text, "align": "right", "pad_mm": 20}})
    doc = PdfDocument(res.blob)
    wd = next(wd for _r, wd in pagedraw.records(doc))
    st = doc.resolve(doc.resolve(wd["AP"])["N"])
    data = doc.decode(st)
    vw = doc.resolve(st.dict["BBox"])[2]
    pad = 20 * 72 / 25.4
    for size, start, line in re.findall(rb"1 0 0 1 ([\d.]+) [\d.-]+ Tm\n/\w+ ([\d.]+) Tf \((.*?)\)", data) and [
            (float(m[2]), float(m[1]), m[3].decode()) for m in re.finditer(rb"1 0 0 1 ([\d.]+) [\d.-]+ Tm\n/\w+ ([\d.]+) Tf \((.*?)\)", data)]:
        assert start >= pad - 0.01
        assert start + text_width(line, size) <= vw - pad + 0.01
    with pytest.raises(FillError):
        fill(PdfDocument(blob), {s["id"]: {"value": "H" * 3000, "align": "right", "pad_mm": 20}})
