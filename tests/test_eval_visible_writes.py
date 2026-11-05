"""書いた値が「証明も読み戻しも通るのに見えない」穴（評価の世界の PDF で見つかった形を、作った書類で）。

- 高さの足りない 1 行の所：字が BBox で切られる → 収まる大きさにする
- どのページの注釈の一覧にも入っていない部品：どのビューアーも描かない → 隠れた欄として断る
- NeedAppearances：ビューアーが描き直して字化ける・消える → 描いた見た目のまま見せる（ほかの部品が見た目を持つとき）
- 書いた欄から /MK の枠が消える
- 同じ色の下地の上の字：画素が変わらない → 読み戻しの違いにする
"""

from __future__ import annotations

import pytest

import pdfmini
from todoku import api
from todoku.pdf import inkcheck, structadd
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import Update, layout_text, text_appearance
from todoku.pdf.form import read_fields
from todoku.pdf.objects import Name, Ref


def _fill(tmp_path, blob, values):
    src = tmp_path / "a.pdf"
    src.write_bytes(blob)
    out = tmp_path / "b.pdf"
    return api.fill(str(src), values, str(out), overwrite=True), out


def _ap(path_or_blob, name):
    blob = path_or_blob if isinstance(path_or_blob, bytes) else path_or_blob.read_bytes()
    doc = PdfDocument(blob)
    f = next(x for x in read_fields(doc) if x.name == name)
    wd = doc.resolve(f.widgets[0].ref)
    return doc, doc.decode(doc.resolve(doc.resolve(wd["AP"])["N"]))


def _widget(rect="100 700 300 716", extra="", name="a"):
    return f"<< /Type /Annot /Subtype /Widget /FT /Tx /T ({name}) /Rect [{rect}] /F 4 /DA (/Helv 12 Tf 0 g) {extra} >>"


def _form(content="", extra="", acro="", rect="100 700 300 716", annots="/Annots [2 0 R]"):
    return pdfmini.build(content, "<< >>", [_widget(rect, extra)], annots=annots, acroform=f"/AcroForm << /Fields [2 0 R] /DA (/Helv 0 Tf 0 g) {acro} >>")


# ====================================================================== 高さの足りない所
def test_one_line_text_is_shrunk_to_fit_the_height():
    lines, s = layout_text("山田花子", 300, 6, 12, multiline=False)
    assert lines == ["山田花子"] and s * 1.15 <= 6.5
    with pytest.raises(Exception):
        layout_text("山田花子", 300, 2, 12, multiline=False)


def test_low_field_is_not_clipped_by_its_box(tmp_path):
    r, out = _fill(tmp_path, _form(rect="100 700 300 708"), {"a": "AB"})
    assert r["proof_ok"] and r["read_back_ok"], r
    _doc, ap = _ap(out, "a")
    size = float(ap.split(b" Tf")[0].split()[-1])
    assert size * 1.15 <= 8.5
    assert any("小さく" in w or "読みにくい" in w for w in r["warnings"])


def test_baseline_is_kept_inside_the_box(tmp_path):
    # 行の基線にそろえるときも、字の箱が枠（BBox）を出ない高さに収める
    doc = PdfDocument(_form())
    up = Update(doc)
    fonts = lambda k: Ref(1, 0)  # noqa: E731
    ap, s = text_appearance(up, 200, 10, "AB", 8, fonts, multiline=False, baseline=0.0)
    ops = up.objs[ap.num][1].raw
    assert s == 8
    y = float(ops.split(b" Tm")[0].split()[-1]) if b" Tm" in ops else float(ops.split(b" Td")[0].split()[-1])
    assert s * 0.12 - 0.01 <= y <= 10 - s * 0.88 + 0.01


# ====================================================================== 一覧に入っていない部品
def test_widget_in_no_page_annots_is_refused_as_unseen(tmp_path):
    blob = _form(extra="/P 1 0 R", annots="")
    doc = PdfDocument(blob)
    w = read_fields(doc)[0].widgets[0]
    assert w.page == 0 and not w.listed and not w.on_screen and not w.printed
    with pytest.raises(api.ApiError) as e:
        _fill(tmp_path, blob, {"a": "AB"})
    assert "/Annots" in str(e.value) and "allow_hidden" in str(e.value)
    r, _out = _fill(tmp_path, blob, {"a": {"value": "AB", "allow_hidden": True}})
    assert r["proof_ok"] and r["read_back_ok"]
    assert any("画面にも紙にも出ない" in w for w in r["warnings"])


# ====================================================================== NeedAppearances
def test_need_appearances_is_cleared_when_every_other_widget_is_drawn(tmp_path):
    ap = pdfmini.stream("", "/Type /XObject /Subtype /Form /BBox [0 0 200 16]")
    other = _widget("100 600 300 616", "/AP << /N 4 0 R >>", "b")
    blob = pdfmini.build("", "<< >>", [_widget(), other, ap], annots="/Annots [2 0 R 3 0 R]",
                         acroform="/AcroForm << /Fields [2 0 R 3 0 R] /NeedAppearances true >>")
    r, out = _fill(tmp_path, blob, {"a": "山田"})
    assert r["proof_ok"] and r["read_back_ok"], r
    assert not any("NeedAppearances" in w for w in r["warnings"])
    doc = PdfDocument(out.read_bytes())
    assert doc.resolve(doc.resolve(doc.root["AcroForm"]).get("NeedAppearances")) is False


def test_need_appearances_stays_when_another_widget_has_no_appearance(tmp_path):
    other = _widget("100 600 300 616", "/V (keep)", "b")
    blob = pdfmini.build("", "<< >>", [_widget(), other], annots="/Annots [2 0 R 3 0 R]",
                         acroform="/AcroForm << /Fields [2 0 R 3 0 R] /NeedAppearances true >>")
    r, out = _fill(tmp_path, blob, {"a": "山田"})
    assert r["proof_ok"] and r["read_back_ok"], r
    assert any("NeedAppearances" in w for w in r["warnings"])
    doc = PdfDocument(out.read_bytes())
    assert doc.resolve(doc.resolve(doc.root["AcroForm"]).get("NeedAppearances")) is True


# ====================================================================== /MK の枠と下地
def test_border_and_background_of_the_field_stay_in_the_appearance(tmp_path):
    r, out = _fill(tmp_path, _form(extra="/MK << /BC [0.4 0.4 0.4] /BG [0.9] >> /BS << /W 2 >>"), {"a": "AB"})
    assert r["proof_ok"] and r["read_back_ok"], r
    _doc, ap = _ap(out, "a")
    assert b"0.4 0.4 0.4 RG" in ap and b"2 w" in ap and b"0.9 g" in ap
    assert ap.index(b"0.9 g") < ap.index(b"BT")  # 字より先に描く


def test_comb_field_draws_dividers_of_the_border(tmp_path):
    r, out = _fill(tmp_path, _form(extra="/Ff 16777216 /MaxLen 4 /MK << /BC [0] >>"), {"a": "AB"})
    assert r["proof_ok"] and r["read_back_ok"], r
    _doc, ap = _ap(out, "a")
    assert ap.count(b" l\nS") == 3


# ====================================================================== 見えない字
def test_black_text_on_black_is_a_read_back_difference(tmp_path):
    black = "0 g 90 690 300 40 re f"
    r, out = _fill(tmp_path, _form(content=black), {"a": "AB"})
    assert r["read_back_ok"] is False
    assert any("見えません" in d["read_back"] for d in r["read_back_differences"])
    assert not out.exists() or r["proof_ok"]


def test_text_on_a_dark_ground_is_told_as_low_contrast(tmp_path):
    r, _out = _fill(tmp_path, _form(content="0 0 1 rg 90 690 300 40 re f"), {"a": "AB"})
    assert r["read_back_ok"], r
    assert any("明るさの差" in w for w in r["warnings"])


def test_plain_write_has_no_visibility_complaint(tmp_path):
    r, _out = _fill(tmp_path, _form(), {"a": "AB"})
    assert r["proof_ok"] and r["read_back_ok"], r
    assert not any("明るさ" in w or "見えるか" in w for w in r["warnings"])


def test_rewriting_the_same_value_is_not_invisible(tmp_path):
    _r, out = _fill(tmp_path, _form(), {"a": "AB"})
    dst = tmp_path / "c.pdf"
    r = api.fill(str(out), {"a": "AB"}, str(dst), overwrite=True)
    assert r["read_back_ok"], r


def test_contrast_is_the_wcag_ratio():
    assert inkcheck.contrast((0, 0, 0), (255, 255, 255)) == pytest.approx(21.0)
    assert inkcheck.contrast((9, 9, 9), (9, 9, 9)) == pytest.approx(1.0)


# ====================================================================== 構造の木
def test_element_whose_parent_points_to_an_older_root_reaches_the_root():
    objs = {1: {"Type": Name("StructTreeRoot"), "K": [Ref(3, 0)]},
            2: {"Type": Name("StructTreeRoot")},  # 前の版の根
            3: {"Type": Name("StructElem"), "S": Name("Document"), "P": Ref(2, 0)},
            9: {"StructTreeRoot": Ref(1, 0)}}

    class D:
        trailer = {"Root": Ref(9, 0)}

        def resolve(self, o):
            return objs.get(o.num) if isinstance(o, Ref) else o

    chain = structadd.Tree(D()).chain(Ref(3, 0))
    assert [r.num for r, _d in chain] == [3]
