"""第4回の見直しで見つかった、PDF のフォームへの誤った書き込みの直し（合成した書類で確かめる）。

A1 部品ごとに「入」の名前が違うチェックの欄（/M と /F）：ひとつ選ぶ欄として、値と同じ名前の部品だけを入にする
A5 リスト：選んだ番号（/I）を値に合わせる・複数選べるリストにリストの値・リストの見た目（選択肢を並べ、選んだ行に色）
A6 1 行の欄：折り返さない。小さくした字の大きさを /DA にも書く。NeedAppearances を知らせる
A7 画面にも紙にも出ない欄（Hidden・NoView・Print なし）：書かない（allow_hidden で書くときは、出ないことを伝える）
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.fill import FillError, Update, appearance_readback, fill  # noqa: E402
from todoku.pdf.form import read_fields  # noqa: E402
from todoku.pdf.objects import Name  # noqa: E402
from todoku.pdf.proof import prove  # noqa: E402


def _build(name_da="/Helv 12 Tf 0 g", name_rect="100 755 180 771", name_flags=4, name_ff=0, amount_flags=4,
           need_app=False, list_ti="", af_da="/Helv 0 Tf 0 g") -> bytes:
    objs: list[bytes] = []

    def add(b):
        objs.append(b if isinstance(b, bytes) else b.encode("latin-1"))
        return len(objs)

    def set_(i, b):
        objs[i - 1] = b if isinstance(b, bytes) else b.encode("latin-1")

    cat, pages = add(""), add("")
    font = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    zadb = add("<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>")
    dr = f"<< /Font << /Helv {font} 0 R /ZaDb {zadb} 0 R >> >>"

    def apbox(on):
        return add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 12 12] /Resources {dr}",
                          b"q BT /ZaDb 10 Tf 2 2 Td (4) Tj ET Q" if on else b""))

    cs = add(stream("", b"BT /Helv 10 Tf 50 760 Td (Name) Tj ET BT /Helv 10 Tf 50 730 Td (Male) Tj ET "
                        b"BT /Helv 10 Tf 200 730 Td (Female) Tj ET BT /Helv 10 Tf 50 700 Td (Amount) Tj ET "
                        b"BT /Helv 10 Tf 50 600 Td (Plan) Tj ET BT /Helv 10 Tf 50 540 Td (Colors) Tj ET"))
    page = add("")
    ff = f" /Ff {name_ff}" if name_ff else ""
    name = add(f"<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /Rect [{name_rect}] /DA ({name_da}){ff} "
               f"/F {name_flags} /P {page} 0 R >>")
    par = add("")
    a_on, a_off, b_on, b_off = apbox(True), apbox(False), apbox(True), apbox(False)
    k1 = add(f"<< /Type /Annot /Subtype /Widget /Parent {par} 0 R /Rect [130 727 142 739] /AS /Off "
             f"/AP << /N << /M {a_on} 0 R /Off {a_off} 0 R >> >> /MK << /CA (4) >> /DA (/ZaDb 0 Tf 0 g) /F 4 /P {page} 0 R >>")
    k2 = add(f"<< /Type /Annot /Subtype /Widget /Parent {par} 0 R /Rect [240 727 252 739] /AS /Off "
             f"/AP << /N << /F {b_on} 0 R /Off {b_off} 0 R >> >> /MK << /CA (4) >> /DA (/ZaDb 0 Tf 0 g) /F 4 /P {page} 0 R >>")
    set_(par, f"<< /FT /Btn /T (gender) /V /Off /Kids [{k1} 0 R {k2} 0 R] >>")
    amt = add(f"<< /Type /Annot /Subtype /Widget /FT /Tx /T (amount) /Rect [100 695 250 711] /DA (/Helv 10 Tf 0 g) "
              f"/F {amount_flags} /P {page} 0 R >>")
    ap = add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 100 40] /Resources {dr}",
                    b"/Tx BMC q BT /Helv 10 Tf 2 28 Td (Alpha) Tj ET Q EMC"))
    lb = add(f"<< /Type /Annot /Subtype /Widget /FT /Ch /T (plan) /Rect [100 570 200 610] /Opt [(Alpha) (Beta) (Gamma)] "
             f"/V (Alpha) /I [0]{list_ti} /DA (/Helv 10 Tf 0 g) /AP << /N {ap} 0 R >> /F 4 /P {page} 0 R >>")
    ms = add(f"<< /Type /Annot /Subtype /Widget /FT /Ch /Ff 2097152 /T (colors) /Rect [100 470 200 550] "
             f"/Opt [(Red) (Green) (Blue)] /V [(Red) (Green)] /I [0 1] /DA (/Helv 10 Tf 0 g) /F 4 /P {page} 0 R >>")
    set_(page, f"<< /Type /Page /Parent {pages} 0 R /MediaBox [0 0 612 792] /Contents {cs} 0 R "
               f"/Resources << /Font << /Helv {font} 0 R >> >> /Annots [{name} 0 R {k1} 0 R {k2} 0 R {amt} 0 R {lb} 0 R {ms} 0 R] >>")
    set_(pages, f"<< /Type /Pages /Kids [{page} 0 R] /Count 1 >>")
    na = " /NeedAppearances true" if need_app else ""
    set_(cat, f"<< /Type /Catalog /Pages {pages} 0 R /AcroForm << /Fields [{name} 0 R {par} 0 R {amt} 0 R {lb} 0 R {ms} 0 R] "
              f"/DR {dr} /DA ({af_da}){na} >> >>")
    return pdf(objs, cat)


def _api(tmp_path, values, blob=None, **kw):
    src = tmp_path / "a.pdf"
    src.write_bytes(blob if blob is not None else _build(**kw))
    out = tmp_path / "b.pdf"
    return api.fill(str(src), values, str(out), overwrite=True), out


def _field(blob, name):
    return {f.name: f for f in read_fields(PdfDocument(blob))}[name]



# ====================================================================== A1
def test_checkbox_group_with_different_on_states_is_one_choice(tmp_path):
    r, out = _api(tmp_path, {"gender": "F"})
    assert r["proof_ok"] and r["read_back_ok"], r
    f = _field(out.read_bytes(), "gender")
    assert f.kind == "radio"
    assert f.value.name == "F"
    assert [w.state for w in f.widgets] == ["Off", "F"]  # Male の部品は切のまま
    r, out = _api(tmp_path, {"gender": "M"})
    assert [w.state for w in _field(out.read_bytes(), "gender").widgets] == ["M", "Off"]


def test_checkbox_group_bare_true_is_refused_with_states(tmp_path):
    with pytest.raises(api.ApiError) as e:
        _api(tmp_path, {"gender": True})
    assert "「M」" in str(e.value) and "「F」" in str(e.value)


def test_checkbox_group_false_turns_all_off(tmp_path):
    r, out = _api(tmp_path, {"gender": False})
    assert r["proof_ok"] and r["read_back_ok"], r
    f = _field(out.read_bytes(), "gender")
    assert f.value.name == "Off" and [w.state for w in f.widgets] == ["Off", "Off"]


def test_window_lists_the_on_states_as_options(tmp_path):
    src = tmp_path / "a.pdf"
    src.write_bytes(_build())
    fs = {f["id"]: f for f in api.fields(str(src))}
    assert fs["gender"]["kind"] == "radio"
    assert fs["gender"]["options"] == ["M", "F"]


def test_proof_and_readback_catch_both_widgets_on(tmp_path):
    """V は /F なのに、Male の部品も入：証明（つじつま）と読み戻し（部品ごとの /AS）の両方で見つける。"""
    orig = _build()
    doc = PdfDocument(orig)
    res = fill(doc, {"gender": "F"})
    bad_doc = PdfDocument(res.blob)
    up = Update(bad_doc)
    f = _field(res.blob, "gender")
    wd = up.get(f.widgets[0].ref)
    wd["AS"] = Name("M")
    bad = up.build()
    pr = prove(orig, bad, {"gender"})
    assert not pr.ok and any("gender" in u for u in pr.unexpected), pr.unexpected
    diff = appearance_readback(PdfDocument(bad), res.patches)
    assert any(d["id"] == "gender" for d in diff), diff


# ====================================================================== A5
def test_list_box_selection_index_follows_value(tmp_path):
    r, out = _api(tmp_path, {"plan": "Beta"})
    assert r["proof_ok"] and r["read_back_ok"], r
    doc = PdfDocument(out.read_bytes())
    f = _field(out.read_bytes(), "plan")
    node = doc.resolve(f.ref)
    assert f.value_text() == "Beta"
    assert node["I"] == [1]
    # 見た目：選択肢を並べて描き、選んだ行に色を付ける
    from todoku.pdf.glyphs import appearance_text, widget_appearance

    st = widget_appearance(doc, node)
    text, probs = appearance_text(doc, st)
    assert not probs and text.replace(" ", "").replace("\n", "") == "AlphaBetaGamma"
    assert b" rg" in doc.decode(st) and b" re f" in doc.decode(st)


def test_multiselect_list_takes_a_list_and_sorts_indices(tmp_path):
    r, out = _api(tmp_path, {"colors": ["Blue", "Red"]})
    assert r["proof_ok"] and r["read_back_ok"], r
    doc = PdfDocument(out.read_bytes())
    f = _field(out.read_bytes(), "colors")
    node = doc.resolve(f.ref)
    assert f.value_text() == ["Red", "Blue"]
    assert node["I"] == [0, 2]


def test_single_select_list_refuses_a_list(tmp_path):
    with pytest.raises(api.ApiError) as e:
        _api(tmp_path, {"plan": ["Alpha", "Beta"]})
    assert "1 つだけ" in str(e.value)


def test_list_scrolls_so_the_selection_is_visible(tmp_path):
    # 高さ 40pt・字 10pt：3 行見える。一番上の行（/TI）が 2 なら、選んだ Alpha（0）は見えないので、/TI を 0 にする
    r, out = _api(tmp_path, {"plan": "Alpha"}, list_ti=" /TI 2")
    assert r["proof_ok"] and r["read_back_ok"], r
    node = PdfDocument(out.read_bytes()).resolve(_field(out.read_bytes(), "plan").ref)
    assert node["TI"] == 0


def test_proof_catches_stale_list_index(tmp_path):
    orig = _build()
    res = fill(PdfDocument(orig), {"plan": "Beta"})
    up = Update(PdfDocument(res.blob))
    f = _field(res.blob, "plan")
    up.get(f.ref)["I"] = [0]
    pr = prove(orig, up.build(), {"plan"})
    assert any("/I" in u for u in pr.unexpected), pr.unexpected


# ====================================================================== A6
def test_shrunk_text_writes_the_used_size_into_da(tmp_path):
    r, out = _api(tmp_path, {"name": "Alexander Hamilton Jr"})
    assert r["proof_ok"] and r["read_back_ok"], r
    doc = PdfDocument(out.read_bytes())
    f = _field(out.read_bytes(), "name")
    da = bytes(doc.resolve(f.ref)["DA"]).decode()
    size = float(da.split()[1])
    assert size < 12
    ap = doc.decode(doc.resolve(doc.resolve(f.ref)["AP"]["N"]))
    assert f"{size:.2f} Tf".encode() in ap  # 見た目と同じ大きさ
    assert any("/DA" in w for w in r["warnings"])


def test_auto_size_da_stays_auto(tmp_path):
    # 字に十分な幅がある欄。狭い欄は、描き直すビューアーで欠けないよう決まった大きさにする（test_pdf_drawn）
    r, out = _api(tmp_path, {"name": "Alexander Hamilton Jr"}, name_rect="100 700 500 720", name_da="/Helv 0 Tf 0 g")
    assert r["proof_ok"] and r["read_back_ok"]
    doc = PdfDocument(out.read_bytes())
    assert bytes(doc.resolve(_field(out.read_bytes(), "name").ref)["DA"]).decode() == "/Helv 0 Tf 0 g"


def test_single_line_field_is_never_wrapped(tmp_path):
    # 高さ 60pt の 1 行の欄（Multiline の印なし）：以前は折り返して描いていた
    r, out = _api(tmp_path, {"name": "Alexander Hamilton Jr the Third"}, name_rect="100 700 180 760", name_da="/Helv 10 Tf 0 g")
    assert r["proof_ok"] and r["read_back_ok"]
    doc = PdfDocument(out.read_bytes())
    ap = doc.decode(doc.resolve(doc.resolve(_field(out.read_bytes(), "name").ref)["AP"]["N"]))
    assert ap.count(b" Tm") == 1  # 1 行


def test_multiline_field_still_wraps(tmp_path):
    r, out = _api(tmp_path, {"name": "Alexander Hamilton Jr the Third"}, name_rect="100 700 180 760",
                  name_da="/Helv 10 Tf 0 g", name_ff=4096)
    assert r["proof_ok"] and r["read_back_ok"]
    doc = PdfDocument(out.read_bytes())
    ap = doc.decode(doc.resolve(doc.resolve(_field(out.read_bytes(), "name").ref)["AP"]["N"]))
    assert ap.count(b" Tm") > 1


def test_single_line_too_long_is_refused(tmp_path):
    with pytest.raises(api.ApiError) as e:
        _api(tmp_path, {"name": "x" * 200}, name_rect="100 700 130 760")
    assert "1 行の欄" in str(e.value)


def test_need_appearances_is_reported(tmp_path):
    r, _out = _api(tmp_path, {"name": "Al"}, need_app=True)
    assert any("NeedAppearances" in w for w in r["warnings"])
    assert "NeedAppearances" in r["say_to_user"]


# ====================================================================== A7
def test_hidden_field_is_refused(tmp_path):
    with pytest.raises(api.ApiError) as e:
        _api(tmp_path, {"name": "Al"}, name_flags=2)
    assert "allow_hidden" in str(e.value) and "Hidden" in str(e.value)


def test_noview_field_is_refused(tmp_path):
    with pytest.raises(api.ApiError):
        _api(tmp_path, {"name": "Al"}, name_flags=32)  # 画面に出さない・印刷しない


def test_noview_but_printed_field_is_written_with_a_warning(tmp_path):
    r, _out = _api(tmp_path, {"name": "Al"}, name_flags=4 | 32)
    assert r["proof_ok"] and r["read_back_ok"]
    assert "画面には出ない" in r["say_to_user"]


def test_hidden_field_with_allow_hidden_is_written_and_told(tmp_path):
    r, out = _api(tmp_path, {"name": {"value": "Al", "allow_hidden": True}}, name_flags=2)
    assert r["proof_ok"] and r["read_back_ok"], r
    assert _field(out.read_bytes(), "name").value_text() == "Al"
    assert "紙に印刷されず" in r["say_to_user"]


def test_unprinted_field_is_written_with_a_warning(tmp_path):
    r, _out = _api(tmp_path, {"amount": "5"}, amount_flags=0)
    assert r["proof_ok"] and r["read_back_ok"]
    assert "紙に印刷されない" in r["say_to_user"]


def test_hidden_radio_widget_is_refused_but_visible_one_is_written(tmp_path):
    # 2 つめの部品（Female）だけを隠す
    src = tmp_path / "h.pdf"
    doc_blob = _build()
    i = doc_blob.index(b"/F 4 /P", doc_blob.index(b"/N << /F "))
    doc_blob = doc_blob[:i] + b"/F 2 /P" + doc_blob[i + 7:]
    src.write_bytes(doc_blob)
    with pytest.raises(api.ApiError):
        api.fill(str(src), {"gender": "F"}, str(tmp_path / "o.pdf"), overwrite=True)
    r = api.fill(str(src), {"gender": "M"}, str(tmp_path / "o.pdf"), overwrite=True)
    assert r["proof_ok"] and r["read_back_ok"]


def test_field_visibility_marks(tmp_path):
    fs = {f.name: f for f in read_fields(PdfDocument(_build(name_flags=2, amount_flags=0)))}
    assert fs["name"].visibility()[0] and fs["name"].visibility()[1] is None
    assert fs["amount"].visibility()[0] is None and fs["amount"].visibility()[1]
    assert fs["plan"].visibility() == (None, None)
