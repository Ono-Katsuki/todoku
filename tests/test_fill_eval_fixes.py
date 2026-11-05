"""白紙からの記入の評価で見つかった問題の確かめ（作った書類で）。"""

from __future__ import annotations

import pytest

from pdfmini import build
from todoku import api, chars
from todoku.pdf import fill as F
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
    assert "/F5" not in da and "/Helv 9 Tf" in da and "0 0 0 rg" in da  # 色と大きさはそのまま
    doc = PdfDocument(res.blob)
    fonts = doc.resolve(doc.resolve(doc.resolve(doc.root["AcroForm"])["DR"])["Font"])
    assert doc.resolve(fonts["Helv"])["BaseFont"].name == "Helvetica"


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
    assert F.pdfkit_needs(long, 10) == F.text_width(long, 10) * F.REDRAW_SCALE + F.REDRAW_PAD


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
