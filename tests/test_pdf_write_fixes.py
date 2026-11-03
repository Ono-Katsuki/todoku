"""PDF の書き込みの直し：LibreOffice の小数・描けない字・電子署名・AcroForm の書式と選択肢・字の目印の正規化・枠の字の大きさ。

どれも、テストの中で組み立てた小さな PDF で確かめる（外部のライブラリなし）。
"""

from __future__ import annotations

import re
import unicodedata

import pytest

from todoku import api
from todoku.pdf import glyphs
from todoku.pdf.acrojs import commit_pair, js_unescape, parse_call
from todoku.pdf.anchor import AnchorError, box_base_size, find_spans, resolve_write
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import FillError, appearance_readback, fill, inline_align
from todoku.pdf.glyphs import appearance_text
from todoku.pdf.proof import prove

from pdf_forms_maker import build_all, pdf, stream
from test_pdf import jp_font_objs, jp_text, make_pdf


def _write(tmp_path, blob: bytes, name: str = "in.pdf"):
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


# ====================================================================== 1. LibreOffice の PDF：ページの小数をそのまま
def libreoffice_like() -> bytes:
    """LibreOffice が書く形：MediaBox が 15 桁の小数（595.303937007874）。見出し「Name」の右が空いている。"""
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.303937007874 841.889763779528] /Resources << /Font << /F1 5 0 R >> >>"
            b" /Contents 4 0 R >>",
            stream("", b"0.5 w 60 740 m 540 740 l S 60 780 m 540 780 l S 60 740 m 60 780 l S 540 740 m 540 780 l S "
                   b"BT /F1 10 Tf 1 0 0 1 72 760 Tm (Name) Tj ET"),
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"]
    return pdf(objs, 1)


def test_libreoffice_floats_are_kept_byte_exact(tmp_path):
    src = _write(tmp_path, libreoffice_like())
    r = api.fill(src, {"@1": {"after": "Name", "value": "Ono"}}, str(tmp_path / "out.pdf"))
    assert r["proof_ok"], r["unexpected"]
    assert r["read_back_ok"]
    new = (tmp_path / "out.pdf").read_bytes()
    tail = new[len(libreoffice_like()):]
    assert b"595.303937007874" in tail and b"841.889763779528" in tail  # 書き換えたページの辞書でも、元の字のまま


def test_libreoffice_floats_survive_the_audit_write_probes(tmp_path):
    """書類の検査（audit）の試し書き：表の欄・字を目印にした書き込み・○（ページの描画を足す）のどれでも、
    書き換えたページの辞書の小数が元の字のまま（「MediaBox が変わっています」にならない）。"""
    chars = "氏名生年月日昭和・平成住所"
    font, (desc, tu), codes = jp_font_objs(chars, 6)
    t = lambda s: jp_text(codes, s).decode()  # noqa: E731
    content = ("0.5 w 60 760 m 540 760 l S 60 730 m 540 730 l S 60 700 m 540 700 l S 60 670 m 540 670 l S "
               "60 760 m 60 670 l S 200 760 m 200 670 l S 540 760 m 540 670 l S "
               f"BT /F1 10 Tf 70 742 Td {t('氏名')} Tj ET BT /F1 10 Tf 70 712 Td {t('生年月日')} Tj ET "
               f"BT /F1 10 Tf 210 712 Td {t('昭和・平成')} Tj ET BT /F1 10 Tf 70 682 Td {t('住所')} Tj ET").encode()
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595.303937007874 841.889763779528] /Resources << /Font << /F1 5 0 R >> >>"
            b" /Contents 4 0 R >>",
            stream("", content), font, desc, tu]
    src = _write(tmp_path, pdf(objs, 1))
    a = api.audit(src)
    bad = [c for c in a["checks"] if not c["ok"] and c["name"] in ("全入力欄への書き込み", "読み戻し", "字を目印にした書き込み")]
    assert not bad, [c["items"] for c in bad]
    # ○（ページの描画を足す）：書いて、書いた後の書類の検査も通る（ページの辞書の小数は元の字のまま）
    r = api.fill(src, {"@c1": {"circle": "昭和", "near": "生年月日"}}, str(tmp_path / "c.pdf"))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    a = api.audit(str(tmp_path / "c.pdf"))
    bad = [c for c in a["checks"] if not c["ok"] and c["name"] in ("全入力欄への書き込み", "読み戻し", "書いてある○")]
    assert not bad, [c["items"] for c in bad]


# ====================================================================== 2. 描けない字
@pytest.fixture
def no_embed(monkeypatch):
    """このパソコンの書体を埋め込まない（標準の書体だけで描く）。"""
    monkeypatch.setenv("TODOKU_PDF_FONT", "")


def test_glyph_choice_by_spec(no_embed):
    assert glyphs.choose("ş")[0] == "TdkL" and glyphs.choose("ł")[0] == "TdkL"  # Core14 の字形（/Differences）
    assert glyphs.choose("é")[0] == "Helv"
    assert glyphs.choose("Ж")[0] == "TdkJ" and glyphs.choose("Ж")[2] == 1000  # Adobe-Japan1 の全角の字形（補遺 0）
    assert glyphs.choose("ｱ")[2] == 500
    for ch in ("ό", "ς", "ễ", "ư"):
        with pytest.raises(glyphs.GlyphError, match="描けません"):
            glyphs.choose(ch)  # 平成角ゴシック（補遺 2）の範囲の外・どの字形にもない


def test_cjk_cid_respects_declared_supplement():
    cid, why = glyphs.cjk_cid("Adobe-Japan1", "ş")
    assert cid is None  # UniJIS-UCS2-HW-H にない（補遺 5 の字）
    assert glyphs.cjk_cid("Adobe-Japan1", "漢")[0] == 1533


def test_latin_extended_written_and_appearance_read_back(tmp_path):
    src = _write(tmp_path, build_all())
    r = api.fill(src, {"family": "Şükrü Öğüt İzmir Łódź Čapek"}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    doc = PdfDocument((tmp_path / "o.pdf").read_bytes())
    from todoku.pdf.form import read_fields

    f = next(x for x in read_fields(doc) if x.name == "family")
    st = glyphs.widget_appearance(doc, doc.resolve(f.widgets[0].ref))
    text, probs = appearance_text(doc, st)
    assert probs == [] and text.replace("\n", "") == "Şükrü Öğüt İzmir Łódź Čapek"


def test_undrawable_text_is_refused_with_reason(tmp_path, no_embed):
    src = _write(tmp_path, build_all())
    with pytest.raises(api.ApiError, match=r"family: 「ό」.*描けません"):
        api.fill(src, {"family": "Παπαδόπουλος"}, str(tmp_path / "o.pdf"))


def test_proof_catches_a_glyph_outside_the_supplement(tmp_path):
    """見た目の字が、書体の宣言した補遺の外の CID になるもの（白く抜ける）を、証明が見つける。"""
    doc = PdfDocument(build_all())
    res = fill(doc, {"family": "漢"})
    bad = res.blob.replace("漢".encode("utf-16-be").hex().upper().encode(), "ş".encode("utf-16-be").hex().upper().encode())
    pr = prove(doc.blob, bad, {"family"})
    assert not pr.ok and any("描けません" in u or "ありません" in u for u in pr.unexpected), pr.unexpected
    diffs = appearance_readback(PdfDocument(bad), res.patches)
    assert diffs and diffs[0]["id"] == "family"


@pytest.mark.fonts
def test_embedded_truetype_subset_for_greek_and_vietnamese(tmp_path):
    from todoku.pdf.embedfont import system_font

    if system_font() is None:
        pytest.skip("このパソコンに、埋め込める TrueType の書体がない")
    src = _write(tmp_path, build_all())
    value = "Παπαδόπουλος Nguyễn Đức"
    r = api.fill(src, {"family": value}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert any("埋め込みました" in w for w in r["warnings"])
    new = (tmp_path / "o.pdf").read_bytes()
    assert len(new) - len(build_all()) < 200_000  # 使った字形だけ
    doc = PdfDocument(new)
    from todoku.pdf.form import read_fields

    f = next(x for x in read_fields(doc) if x.name == "family")
    text, probs = appearance_text(doc, glyphs.widget_appearance(doc, doc.resolve(f.ref)))
    assert probs == [] and text.replace(chr(10), "") == value
    with pytest.raises(glyphs.GlyphError, match="シェーピング"):
        glyphs.choose(chr(0x301))  # 結合の字は、埋め込みでも描かない


# ====================================================================== 3. 電子署名
def signed(mode: str) -> bytes:
    """署名した形の PDF（署名の中身は 0 のまま。とどくが読むのは署名の辞書の決まりだけ）。
    mode：approval（承認の署名だけ）・p1・p2・p3（認証の署名 DocMDP）・lock（/Lock で欄 name を固定）。"""
    objs: list[bytes] = []

    def add(b):
        objs.append(b if isinstance(b, bytes) else b.encode("latin-1"))
        return len(objs)

    def set_(i, b):
        objs[i - 1] = b if isinstance(b, bytes) else b.encode("latin-1")

    cat, pages = add(""), add("")
    font = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    cs = add(stream("", b"BT /F1 10 Tf 1 0 0 1 72 760 Tm (Name) Tj 1 0 0 1 72 700 Tm (Note) Tj 1 0 0 1 72 600 Tm (Office) Tj ET"))
    p1 = add("")
    ap = add(stream("/Type /XObject /Subtype /Form /BBox [0 0 150 20]", b"/Tx BMC EMC"))
    name = add(f"<< /Type /Annot /Subtype /Widget /FT /Tx /T (name) /Rect [150 755 300 775] /P {p1} 0 R /DA (/Helv 10 Tf 0 g) /F 4 /AP << /N {ap} 0 R >> >>")
    note = add(f"<< /Type /Annot /Subtype /Widget /FT /Tx /T (note) /Rect [150 695 300 715] /P {p1} 0 R /DA (/Helv 10 Tf 0 g) /F 4 /AP << /N {ap} 0 R >> >>")
    ref = ""
    if mode in ("p1", "p2", "p3"):
        ref = f"/Reference [<< /Type /SigRef /TransformMethod /DocMDP /TransformParams << /Type /TransformParams /P {mode[1]} /V /1.2 >> >>]"
    sigv = add(f"<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 10 20 10] /Contents <00> {ref} >>")
    lock = " /Lock << /Type /SigFieldLock /Action /Include /Fields [(name)] >>" if mode == "lock" else ""
    sigf = add(f"<< /Type /Annot /Subtype /Widget /FT /Sig /T (office_sig) /V {sigv} 0 R /Rect [150 590 350 620] /P {p1} 0 R /F 132{lock} >>")
    set_(p1, f"<< /Type /Page /Parent {pages} 0 R /MediaBox [0 0 595 842] /Contents {cs} 0 R /Resources << /Font << /F1 {font} 0 R >> >> "
             f"/Annots [{name} 0 R {note} 0 R {sigf} 0 R] >>")
    set_(pages, f"<< /Type /Pages /Kids [{p1} 0 R] /Count 1 >>")
    perms = f" /Perms << /DocMDP {sigv} 0 R >>" if mode in ("p1", "p2", "p3") else ""
    set_(cat, f"<< /Type /Catalog /Pages {pages} 0 R /AcroForm << /Fields [{name} 0 R {note} 0 R {sigf} 0 R] /SigFlags 3 "
              f"/DR << /Font << /Helv {font} 0 R >> >> >>{perms} >>")
    return pdf(objs, cat)


def test_signed_pdf_tells_the_user(tmp_path):
    src = _write(tmp_path, signed("approval"))
    r = api.fill(src, {"note": "hello"}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and r["read_back_ok"]
    assert "署名のあとの変更" in r["say_to_user"]


def test_fieldmdp_lock_is_read_only(tmp_path):
    src = _write(tmp_path, signed("lock"))
    with pytest.raises(api.ApiError, match="name: 電子署名（office_sig）で固定された欄"):
        api.fill(src, {"name": "Ono"}, str(tmp_path / "o.pdf"))
    r = api.fill(src, {"note": "ok"}, str(tmp_path / "o2.pdf"))  # 固定されていない欄は書ける
    assert r["proof_ok"]


def test_docmdp_p2_allows_only_existing_fields(tmp_path):
    src = _write(tmp_path, signed("p2"))
    r = api.fill(src, {"note": "hello"}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and "署名のあとの変更" in r["say_to_user"]
    for v in ({"@1": {"after": "Office", "value": "extra", "page": 1, "where": "below"}},
              {"@1": {"page": 1, "box": [20, 200, 80, 210], "value": "boxed"}},
              {"@1": {"circle": "Office", "page": 1}}):
        with pytest.raises(api.ApiError, match="DocMDP P=2"):
            api.fill(src, v, str(tmp_path / "x.pdf"))


def test_docmdp_p1_refuses_and_grades_d(tmp_path):
    src = _write(tmp_path, signed("p1"))
    with pytest.raises(api.ApiError, match="変更を許さない"):
        api.fill(src, {"note": "hello"}, str(tmp_path / "o.pdf"))
    g = api.grade(src)
    assert g["write"]["grade"] == "D"
    assert any(r["title"] == "変更を許さない電子署名" for r in g["write"]["reasons"])
    assert api.grade(_write(tmp_path, signed("p2"), "p2.pdf"))["write"]["grade"] != "D"


# ====================================================================== 4. AcroForm
def test_js_string_escapes():
    assert js_unescape(r"¥") == "¥" and js_unescape(r"\x41\n\"") == 'A\n"'
    assert parse_call('AFNumber_Format(2, 0, 0, 0, "\\u00a5", true);')[1][4] == "¥"


def test_number_value_is_raw_and_appearance_is_formatted(tmp_path):
    src = _write(tmp_path, build_all())
    r = api.fill(src, {"amount": "1234.5"}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and r["read_back_ok"], r["read_back_differences"]
    from todoku.pdf.form import read_fields

    doc = PdfDocument((tmp_path / "o.pdf").read_bytes())
    f = next(x for x in read_fields(doc) if x.name == "amount")
    assert f.value_text() == "1234.5"  # 値はそのまま（Acrobat と同じ）
    st = glyphs.widget_appearance(doc, doc.resolve(f.ref))
    assert appearance_text(doc, st)[0] == "¥1,234.50"  # 見た目は書式どおり
    assert "¥1,234.50" in r["say_to_user"]
    assert commit_pair({"F": "AFPercent_Format(2, 0);"}, "25%") == ("0.25", "25.00%")


def test_combo_shows_display_text(tmp_path):
    src = _write(tmp_path, build_all())
    for given in ("Osaka", "27"):
        r = api.fill(src, {"pref": given}, str(tmp_path / "o.pdf"))
        assert r["proof_ok"] and r["read_back_ok"], r["read_back_differences"]
        assert "Osaka" in r["say_to_user"]
        doc = PdfDocument((tmp_path / "o.pdf").read_bytes())
        from todoku.pdf.form import read_fields

        f = next(x for x in read_fields(doc) if x.name == "pref")
        assert f.value_text() == "27"
        assert appearance_text(doc, glyphs.widget_appearance(doc, doc.resolve(f.ref)))[0] == "Osaka"


def test_calculated_field_is_refused(tmp_path):
    src = _write(tmp_path, build_all())
    with pytest.raises(api.ApiError, match=r"total: 計算で決まる欄です（「qty」「price」から計算）"):
        api.fill(src, {"total": "999"}, str(tmp_path / "o.pdf"))
    r = api.fill(src, {"qty": "3"}, str(tmp_path / "o.pdf"))  # 元の欄に書けば、計算し直して書く
    assert r["proof_ok"]


def test_bad_date_and_newline_are_refused_plainly(tmp_path):
    src = _write(tmp_path, build_all())
    with pytest.raises(api.ApiError, match="date: 「2026-13-45」は、ありえない日付です"):
        api.fill(src, {"date": "2026-13-45"}, str(tmp_path / "o.pdf"))
    with pytest.raises(api.ApiError, match="family: 1行の欄です"):
        api.fill(src, {"family": "a\nb"}, str(tmp_path / "o.pdf"))
    r = api.fill(src, {"notes": "a\nb"}, str(tmp_path / "o.pdf"))  # 複数行の欄は書ける
    assert r["proof_ok"] and r["read_back_ok"]


def test_need_appearances_checkbox_and_export_value(tmp_path):
    src = _write(tmp_path, build_all(with_ap=False, need_app=True))
    r = api.fill(src, {"agree": True}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    doc = PdfDocument((tmp_path / "o.pdf").read_bytes())
    from todoku.pdf.form import read_fields

    f = next(x for x in read_fields(doc) if x.name == "agree")
    assert f.checked and f.widgets[0].on == "Yes"  # 見た目がないとき：仕様の勧める Yes、入と切の見た目も足す
    src2 = _write(tmp_path, build_all(), "with_ap.pdf")
    r = api.fill(src2, {"agree": "Ja"}, str(tmp_path / "o2.pdf"))  # 「入」の状態の名前で指してもよい
    assert r["read_back_ok"]
    with pytest.raises(api.ApiError, match="plan: 選択肢の名前が書類にありません"):
        api.fill(src, {"plan": "B"}, str(tmp_path / "o3.pdf"))


# ====================================================================== 5. 字の目印の正規化
def test_find_spans_nfc_and_nfkc():
    nfd = unicodedata.normalize("NFD", "tên: Nguyễn")
    a, b = find_spans(nfd, "tên:")[0]
    assert nfd[a:b] == unicodedata.normalize("NFD", "tên:")
    a, b = find_spans(nfd, "Nguyễn")[0]
    assert unicodedata.normalize("NFC", nfd[a:b]) == "Nguyễn"
    assert find_spans("氏名（ＡＢＣ）", "ABC") == [(3, 6)]  # NFKC
    assert find_spans(unicodedata.normalize("NFD", "tê"), "te") == []  # 字の一部だけには合わせない


def test_anchor_not_found_message_has_no_blank(tmp_path):
    doc = PdfDocument(libreoffice_like())
    with pytest.raises(AnchorError) as e:
        resolve_write(doc, {"after": "tên:", "value": "x"})
    assert "「tên:」が見つかりません" in str(e.value)
    with pytest.raises(AnchorError) as e:
        resolve_write(doc, {"after": "tên:", "value": "x", "page": 1})
    assert "「tên:」が 1 ページに見つかりません" in str(e.value)


def test_anchor_written_after_nfd_label(tmp_path):
    label = unicodedata.normalize("NFD", "tên:")
    font, extra, codes = jp_font_objs(label, 5)
    content = (b"0.5 w 60 740 m 540 740 l S 60 780 m 540 780 l S 60 740 m 60 780 l S 540 740 m 540 780 l S "
               b"BT /F1 10 Tf 72 760 Td " + jp_text(codes, label) + b" Tj ET\n")
    src = _write(tmp_path, make_pdf(content, {"F1": font}, extra))
    r = api.fill(src, {"@1": {"after": "tên:", "value": "Ono"}}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])


# ====================================================================== 6. 置き方・字の大きさ
@pytest.mark.parametrize("label, want", [("〒［　］", "left"), ("［　］年", "right"), ("令和［　］年", "center"),
                                         ("東京都［▶　］区［　］", "center"), ("○○区［　］", "left")])
def test_inline_alignment_follows_the_neighbours(label, want):
    assert inline_align({"label": label}) == want


def test_box_size_follows_paper_and_box_height(tmp_path):
    class V:
        width, height = 841.89, 1190.55  # A3

    assert box_base_size(V) == pytest.approx(14.85, abs=0.01)
    V.width, V.height = 595.28, 841.89
    assert box_base_size(V) == 10.5
    src = _write(tmp_path, libreoffice_like())
    r = api.fill(src, {"@1": {"page": 1, "box": [20, 100, 40, 102.5], "value": "Ono"}},
                 str(tmp_path / "o.pdf"))
    assert any("読みにくい大きさ" in w for w in r["warnings"]), r["warnings"]


def test_q_right_alignment_is_used(tmp_path):
    blob = build_all().replace(b"/T (centered) /DA (/Helv 10 Tf 0 g) /Q 1", b"/T (centered) /DA (/Helv 10 Tf 0 g) /Q 2")
    assert b"/Q 2" in blob
    doc = PdfDocument(blob)
    res = fill(doc, {"centered": "R"})
    new = PdfDocument(res.blob)
    from todoku.pdf.form import read_fields

    f = next(x for x in read_fields(new) if x.name == "centered")
    data = new.decode(glyphs.widget_appearance(new, new.resolve(f.ref)))
    x = float(re.search(rb"1 0 0 1 ([\d.]+) [\d.]+ Tm", data)[1])
    assert x > 100  # 欄（幅 150）の右に寄せる
