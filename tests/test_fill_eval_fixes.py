"""白紙からの記入の評価で見つかった問題の確かめ（作った書類で）。"""

from __future__ import annotations

import struct

import pytest

from pdfmini import build
from todoku import api, chars
from todoku.pdf import fill as F
from todoku.pdf import glyphs
from todoku.pdf.content import page_content
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import fill
from todoku.pdf.form import read_fields
from todoku.pdf.layout import PageLayout
from todoku.pdf.proof import _rect_only_widened, prove

HELV = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
# 埋め込まない書体：見た目を作り直す表示は、この書体で字を描けない（×の四角）
SEGOE = ("<< /Type /Font /Subtype /TrueType /BaseFont /SegoeUI /Encoding /WinAnsiEncoding /FirstChar 32 /LastChar 33 "
         "/Widths [273 284] >>")


def _form(rect: str, da: str = "0 0 0 rg /F5 9 Tf", extra: str = "", mk: str = "") -> bytes:
    """2 3：書体、4：欄。extra があれば 5：もう 1 つの欄。"""
    wid = f"<< /Type /Annot /Subtype /Widget /FT /Tx /T (f) /F 4 /Rect [{rect}] /DA ({da}) {mk} /P 1 0 R >>"
    objs = [HELV, SEGOE, wid] + ([extra] if extra else [])
    refs = "4 0 R 5 0 R" if extra else "4 0 R"
    return build("", "<< >>", objs, annots=f"/Annots [{refs}]",
                 acroform=f"/AcroForm << /Fields [{refs}] /DR << /Font << /F5 3 0 R >> >> /NeedAppearances true >>")


def _widget(blob: bytes) -> dict:
    doc = PdfDocument(blob)
    f = next(x for x in read_fields(doc) if x.name == "f")
    return doc.resolve(f.widgets[0].ref)


def _da(blob: bytes) -> str:
    return bytes(_widget(blob)["DA"]).decode("latin-1")


def test_da_font_follows_the_drawn_font():
    """見た目を作り直す表示は /DA の書体で描く：埋め込まない書体のままだと、値は正しいのに字が×の四角になる。"""
    blob = _form("100 700 400 716")
    res = fill(PdfDocument(blob), {"f": "YAMADA"})
    assert prove(blob, res.blob, {"f"}).ok
    da = _da(res.blob)
    assert "/F5" not in da and "/TdkB 9 Tf" in da and "0 0 0 rg" in da  # 色と大きさはそのまま
    doc = PdfDocument(res.blob)
    fonts = doc.resolve(doc.resolve(doc.resolve(doc.root["AcroForm"])["DR"])["Font"])
    assert doc.resolve(fonts["TdkB"])["BaseFont"].name.endswith("+NotoSansJP-Regular")


def test_narrow_invisible_field_is_widened_so_redrawing_viewers_show_every_digit():
    """PDFKit は /V から描き直し、幅が足りないと字の終わりを描かない（生年月日の「01」が「0」）。枠も下地もない欄は四角を広げる。"""
    blob = _form("167 700 176.5 712")
    res = fill(PdfDocument(blob), {"f": "01"})
    assert prove(blob, res.blob, {"f"}).ok
    r = _widget(res.blob)["Rect"]
    assert r[0] == 167 and r[1] == 700 and r[3] == 712
    assert r[2] - r[0] >= F.pdfkit_needs("01", F._da_size(_da(res.blob)))
    assert not res.redraw_cut and any("広げました" in w for w in res.warnings)


def test_short_values_need_the_measured_room():
    # tests/audit_findings/pdfkit の測り方で、「01」は 6.5pt でも 9pt でも 字の幅 ＋ 7pt で初めて 2 字とも出た
    w = F.text_width("01", 6.5)
    assert F.pdfkit_needs("01", 6.5) >= w + 7
    long = "1234567890123456"
    helv = sum(glyphs.helv_width(c) for c in long) * 10 / 1000  # PDFKit は欧文を Helvetica で描き直す
    assert F.pdfkit_needs(long, 10) == helv * F.REDRAW_SCALE + F.REDRAW_PAD


def test_narrow_field_not_widened_over_another_field_and_the_fact_is_told(tmp_path):
    other = "<< /Type /Annot /Subtype /Widget /FT /Tx /T (g) /F 4 /Rect [178 700 190 712] /DA (/F5 9 Tf) /P 1 0 R >>"
    blob = _form("167 700 176.5 712", extra=other)
    src, out = tmp_path / "a.pdf", tmp_path / "o.pdf"
    src.write_bytes(blob)
    r = api.fill(str(src), {"f": "01"}, str(out))
    assert r["proof_ok"] and r["viewers_differ"][0]["id"] == "f" and r["viewers_differ"][0]["value"] == "01"
    assert "欠けて" in r["say_to_user"]
    assert _widget(out.read_bytes())["Rect"] == [167, 700, 176.5, 712]


def test_framed_field_keeps_its_rect():
    blob = _form("167 700 176.5 712", mk="/MK << /BC [0] >>")
    res = fill(PdfDocument(blob), {"f": "01"})
    assert _widget(res.blob)["Rect"] == [167, 700, 176.5, 712] and res.redraw_cut


def test_proof_accepts_only_a_widened_rect():
    doc = PdfDocument(_form("100 700 110 712"))
    old = {"Rect": [100, 700, 110, 712]}
    assert _rect_only_widened(doc, old, {"Rect": [100, 700, 120, 712]})
    assert not _rect_only_widened(doc, old, {"Rect": [100, 690, 120, 712]})  # 高さが変わる
    assert not _rect_only_widened(doc, old, {"Rect": [105, 700, 120, 712]})  # 元の四角を含まない
    assert not _rect_only_widened(doc, dict(old, MK={"BC": [0]}), {"Rect": [100, 700, 120, 712]})  # 枠のある欄


def test_values_not_written_now_are_told(tmp_path):
    """様式の既定の選択（書いていない欄）は、本人の答えのように見えるので、事実として伝える。"""
    race = ("<< /Type /Annot /Subtype /Widget /FT /Ch /Ff 131072 /T (race) /F 4 /Rect [100 650 300 670] /V (Asian) "
            "/Opt [(Asian) (Other)] /DA (/F5 9 Tf) /P 1 0 R >>")
    blob = _form("100 700 400 716", extra=race)
    src, out = tmp_path / "a.pdf", tmp_path / "o.pdf"
    src.write_bytes(blob)
    r = api.fill(str(src), {"f": "YAMADA"}, str(out))
    assert any(x["id"] == "race" and x["value"] == "Asian" for x in r["values_not_written_now"])
    assert "Asian" in r["say_to_user"] and "今回は書いていない" in r["say_to_user"]
    # 2 回目の書き込み：前の書き込みの値も一覧に出る（本人への報告が最後の書き込みだけにならない）
    out2 = tmp_path / "o2.pdf"
    r2 = api.fill(str(out), {"race": "Other"}, str(out2))
    assert any(x["id"] == "f" and x["value"] == "YAMADA" for x in r2["values_not_written_now"])
    assert "YAMADA" in r2["say_to_user"]
    assert not any(x["id"] == "race" for x in r2["values_not_written_now"])


# ---------------------------------------------------------------- 線の代わりに印刷された字
def _printed(lines: list[tuple[int, int, str]]) -> bytes:
    content = "".join(f"BT /F1 10 Tf {x} {y} Td ({t}) Tj ET\n" for x, y, t in lines)
    return build(content, "<< /Font << /F1 2 0 R >> >>", [HELV])


def _write_box(tmp_path, blob: bytes, y_pdf: float, x_mm=(36, 90)):
    src, out = tmp_path / "a.pdf", tmp_path / "o.pdf"
    src.write_bytes(blob)
    top = (842 - y_pdf) * 25.4 / 72
    return api.fill(str(src), {"@1": {"page": 1, "box": [x_mm[0], top - 3.5, x_mm[1], top + 0.5], "value": "Hanako"}}, str(out))


@pytest.mark.parametrize("line", ["_" * 40, "." * 60])
def test_box_on_printed_underscores_and_dots_is_written(tmp_path, line):
    r = _write_box(tmp_path, _printed([(100, 700, line)]), 700)
    assert r["proof_ok"] and r["read_back_ok"], line


@pytest.mark.parametrize("line", ["Name of person here", "Name ____________________"])
def test_real_text_is_still_refused(tmp_path, line):
    with pytest.raises(api.ApiError, match="ページの字|の上に重なります"):
        _write_box(tmp_path, _printed([(100, 700, line)]), 700)


def test_a_single_full_stop_is_not_a_line():
    doc = PdfDocument(_printed([(100, 700, "End.")]))
    pg = doc.pages()[0]
    assert not PageLayout(doc, pg, page_content(doc, pg), 1).line_glyphs()


def test_middle_dots_are_not_line_chars():
    """字の高さの中ほどに描かれる点（中点の並び）・ダッシュは字に重なるので、線の字にしない。"""
    assert chars.is_low_line_char("_") and chars.is_low_line_char(".") and chars.is_low_line_char("…")
    assert not chars.is_low_line_char("・") and not chars.is_low_line_char("·") and not chars.is_low_line_char("-")


# ---------------------------------------------------------------- 読み戻しの誤検出
def _doc(body: str, box: str = "") -> bytes:
    """本文とテキストボックスの流れを持つ Word 97 の文書。"""
    from test_round8_office import plain_doc

    blob = plain_doc(body + box)
    if not box:
        return blob
    from todoku.cfb import CFB
    from test_round8_office import _cfb_tree

    cfb = CFB(blob)
    wd = bytearray(cfb.read("WordDocument"))
    struct.pack_into("<i", wd, 64 + 4 * 3, len(body))
    struct.pack_into("<i", wd, 64 + 4 * 9, len(box))
    return _cfb_tree({"WordDocument": bytes(wd), "1Table": cfb.read("1Table")})


def test_doc_align_padding_is_not_a_difference(tmp_path):
    src, out = tmp_path / "a.doc", tmp_path / "o.doc"
    src.write_bytes(_doc("Name: ______________________ end\r"))
    fid = next(f["id"] for f in api.fields(str(src)) if f.get("kind") == "inline")
    r = api.fill(str(src), {fid: {"value": "Hanako", "align": "center"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r["read_back_differences"]
    assert any(d["id"] == fid and d["value"] == "Hanako" for d in r["written"])


def test_padding_helper_keeps_real_differences():
    assert api._only_added_padding("  novembro  ", "novembro", [2, 2])
    assert not api._only_added_padding("  novembre  ", "novembro", [2, 2])  # 字が違う
    assert not api._only_added_padding(" novembro  ", "novembro", [2, 2])  # 足した字数と違う
    assert not api._only_added_padding("  novembro  ", "novembro", None)  # align なし


def test_doc_textbox_write_reads_back(tmp_path, monkeypatch):
    src, out = tmp_path / "a.doc", tmp_path / "o.doc"
    src.write_bytes(_doc("Body text\r", "Name: ______________________ end\r"))
    fid = next(f["id"] for f in api.fields(str(src)) if f.get("kind") == "inline" and f["id"].startswith("P2."))
    r = api.fill(str(src), {fid: "Hanako"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r["read_back_differences"]
    # 本当に違うとき（段落に書いた字がない）は通さない
    from todoku.doc import fill as dfill

    monkeypatch.setattr(dfill, "read_back_outside", lambda *a, **k: {})
    bad = api.fill(str(src), {fid: "Hanako"}, str(tmp_path / "o2.doc"))
    assert not bad["read_back_ok"] and bad["read_back_differences"][0]["id"] == fid


def test_read_back_outside_finds_only_the_written_text():
    from todoku.doc.fill import read_back_outside

    blob = _doc("Body text\r", "Name: Hanako end\r")
    assert read_back_outside(blob, "a.doc", {"P2.c6": "Hanako"}) == {"P2.c6": "Hanako"}
    assert read_back_outside(blob, "a.doc", {"P2.c6": "Taro"}) == {}
    assert read_back_outside(blob, "a.doc", {"P1.c6": "Hanako"}) == {}  # 別の段落


def test_xlsx_text_in_a_date_formatted_cell_reads_back():
    """字のある日付の表示形式のセルに字として書いた所：セルの字が書いた字と同じなら合う。違う字は合わない。"""
    from types import SimpleNamespace

    from todoku.xlsx import readback as rb

    cell = SimpleNamespace(value="令和8年11月10日", style=0, schema_error=False)
    sh = SimpleNamespace(cells={(1, 1): cell})
    wb = SimpleNamespace(sheet=lambda n: sh, style=lambda s: SimpleNamespace(num_fmt="yyyy-mm-dd"), date1904=False,
                         sheet_names=["S"])
    import todoku.xlsx.edit as ed
    import todoku.xlsx.window as win

    orig = (ed.resolve_target, ed.interpret_input, win.display)
    ed.resolve_target = lambda w, k: ("S", 1, 1)
    ed.interpret_input = lambda *a, **k: (46336.0, [])
    win.display = lambda w, c: c.value
    try:
        assert rb.check(wb, {"S!A1": "令和8年11月10日"})["S!A1"][0] is True
        assert rb.check(wb, {"S!A1": "令和8年11月11日"})["S!A1"][0] is False
    finally:
        ed.resolve_target, ed.interpret_input, win.display = orig


# ---------------------------------------------------------------- 結果の大きさ
def _many_fields_pdf() -> bytes:
    widgets = [f"<< /Type /Annot /Subtype /Widget /FT /Tx /T (f{i}) /F 4 /Rect [100 {800 - 7 * i} 300 {806 - 7 * i}] /DA (/Helv 6 Tf) /P 1 0 R >>"
               for i in range(100)]
    refs = " ".join(f"{i + 3} 0 R" for i in range(100))
    return build("", "<< >>", [HELV] + widgets, annots=f"/Annots [{refs}]",
                 acroform=f"/AcroForm << /Fields [{refs}] /DR << /Font << /Helv 2 0 R >> >> >>")


def test_list_fields_limit_counts_fields_within_a_page(tmp_path):
    from todoku import mcp_server as m

    p = tmp_path / "m.pdf"
    p.write_bytes(_many_fields_pdf())
    got, ask, seen = [], {"pages": 1, "limit": 3}, 0
    while ask is not None and seen < 60:
        sc = m.call_tool("list_fields", {"path": str(p), **ask})["structuredContent"]
        assert len(sc["fields"]) <= 3
        got += [f["id"] for f in sc["fields"]]
        ask, seen = sc.get("next"), seen + 1
    assert len(got) == 100 == len(set(got))  # 続きで全部読める


def test_read_window_returns_text_in_structured_content(tmp_path):
    from todoku import mcp_server as m

    p = tmp_path / "t.pdf"
    p.write_bytes(_printed([(100, 700, "Printed words on the page")]))
    for args in ({"pages": 1}, {"max_chars": 1000}, {"pages": 1, "max_chars": 1000}):
        sc = m.call_tool("read_window", {"path": str(p), **args})["structuredContent"]
        assert "Printed words" in sc["text"], args


# ---------------------------------------------------------------- ○の位置
def test_long_word_in_a_narrow_row_gets_a_rounded_outline_not_an_ellipse_over_the_lines(tmp_path):
    """長い語の楕円は、左右の線か上下の線まで届く。字に沿った角の丸い囲みなら線のあいだに収まる。"""
    rules = ("0.5 w 103 700 m 103 722 l S 254 700 m 254 722 l S 103 722 m 254 722 l S 103 700 m 254 700 l S\n"
             "BT /F1 10 Tf 110 707 Td (Severe home visit care service) Tj ET\n")
    blob = build(rules, "<< /Font << /F1 2 0 R >> >>", [HELV])
    src, out = tmp_path / "a.pdf", tmp_path / "o.pdf"
    src.write_bytes(blob)
    r = api.fill(str(src), {"@c1": {"circle": "Severe home visit care service"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r["read_back_differences"]
    c = r["circled"][0]
    assert c.get("rounded") and not any("横切" in w or "かかる" in w or "触れ" in w for w in r["warnings"]), r["warnings"]
    assert "角の丸い" in r["say_to_user"]
    # 読み戻しは、書くときと同じ決め方で、言った字を囲むかを見る
    x0, y0, x1, y1 = c["border_pt"]
    assert x0 > 103 and x1 < 254 and y0 > 842 - 722 and y1 < 842 - 700


def test_after_label_writes_on_the_printed_line(tmp_path):
    """見出しのすぐ後ろが下線の字でも、見出しの右に書ける（下線の先に字があれば、そこまで）。"""
    src, out = tmp_path / "a.pdf", tmp_path / "o.pdf"
    src.write_bytes(_printed([(60, 700, "Name ______________________________ Age")]))
    r = api.fill(str(src), {"@1": {"after": "Name", "value": "Hanako"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    rect = r["written_rects"]["@1"]
    assert rect["x"] + rect["w"] < (60 + 170) * 25.4 / 72 + 40  # 「Age」の手前まで
    with pytest.raises(api.ApiError):  # 字の上（「Name」と「Age」が続く行）は断る
        src.write_bytes(_printed([(60, 700, "Name Age")]))
        api.fill(str(src), {"@1": {"after": "Name", "value": "Hanako Yamada Hanako"}}, str(tmp_path / "o2.pdf"))


def test_circle_said_text_uses_an_embedded_font_when_needed(tmp_path):
    """囲んだ字が標準の書体にないとき（キリル文字）：透明な字の書体の中身も書く（順序の誤りで、書体が空のまま証明で断られていた）。"""
    from test_pdf_scripts import glyphs_pdf

    word = "\u0434\u0430\u043d\u0456"
    blob = glyphs_pdf([(ch, 100 + 8 * k, 700) for k, ch in enumerate(word)], {ch: 600 for ch in word})
    src, out = tmp_path / "a.pdf", tmp_path / "o.pdf"
    src.write_bytes(blob)
    r = api.fill(str(src), {"@c1": {"circle": word, "style": "border"}}, str(out))
    assert r["proof_ok"], r["unexpected"]
    assert r["read_back_ok"], r["read_back_differences"]


# ---------------------------------------------------------------- 埋め込まない書体の事実
def test_fonts_not_embedded_are_reported(tmp_path, monkeypatch):
    monkeypatch.setenv("TODOKU_PDF_FONT", "")  # 埋め込まない経路（同梱の書体にない字が通る経路）
    src, out = tmp_path / "a.pdf", tmp_path / "o.pdf"
    src.write_bytes(_form("100 700 400 716"))
    r = api.fill(str(src), {"f": "山田 Hanako"}, str(out))
    got = {x["font"]: x for x in r["fonts_not_embedded"]}
    assert got["HeiseiKakuGo-W5"]["kind"] == "cjk_predefined" and got["HeiseiKakuGo-W5"]["ids"] == ["f"]
    assert got["Helvetica"]["kind"] == "standard14"
    assert "埋め込まず" in r["say_to_user"]


def test_latin_only_writes_do_not_warn_about_cjk_fonts(tmp_path, monkeypatch):
    monkeypatch.setenv("TODOKU_PDF_FONT", "")
    src, out = tmp_path / "a.pdf", tmp_path / "o.pdf"
    src.write_bytes(_form("100 700 400 716"))
    r = api.fill(str(src), {"f": "Hanako"}, str(out))
    assert [x["kind"] for x in r["fonts_not_embedded"]] == ["standard14"]
    assert "埋め込まず" not in r["say_to_user"]
