"""段階の検証（docs/段階の検証と比較.md）で見つかった不具合の、小さな PDF・表での再現と確かめ。

- 幅の狭い字（「l」）が続く「All」を、重ね書きとして1字にしない（重ね書きは、字の送りの幅が半分より多く重なるものだけ）
- 読み落としの検査が、重ね書きでない字を落としたことを見つける
- 字の「_」を、入力欄の印に置き換えて窓から消さない
- 相互参照の番号がずれた PDF でもページを見つける。ページが1つもない PDF は F（理由つき）
- ページの外の欄の候補は出さない
- 読む段階は入力欄の見つけ方で下げない。字に直せない字が行の頭の印だけなら注意（alerts）、本文の字なら E
- 開けない圧縮の画像だけのページは E（書類はスキャン）
"""

from __future__ import annotations

import zlib

from todoku.grade import grade_blob
from todoku.pdf import layout as pdf_layout
from todoku.pdf.audit import check_reading
from todoku.pdf.doc import PdfDocument
from todoku.pdf.window import build

from test_pdf import jp_font_objs, jp_text, make_pdf, tounicode

# 送りの幅：字は 500、「l」「i」は 222（Helvetica と同じ細さ）。1000 分の 1 字の単位
_W = [222 if chr(c) in "li" else 278 if chr(c) == " " else 500 for c in range(32, 127)]
HELV = (b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding /FirstChar 32 /LastChar 126 /Widths ["
        + " ".join(map(str, _W)).encode() + b"] >>")


def _reading(blob: bytes) -> list[str]:
    win = build(PdfDocument(blob, "t.pdf"))
    return [ln for pg in win["pages"] for ln in pg["reading"]]


def _line_pdf() -> bytes:
    c = b"BT /F1 12 Tf 72 700 Td (All Temperature Windows; collections delle) Tj ET\n"
    # 太字に見せる重ね書き：同じ字を 0.3pt ずらして2度描く（1字として読む）
    c += b"BT /F1 12 Tf 72 650 Td (Bold) Tj ET\nBT /F1 12 Tf 72.3 650 Td (Bold) Tj ET\n"
    return make_pdf(c, {"F1": HELV})


def test_narrow_letters_side_by_side_are_not_an_overprint():
    lines = _reading(_line_pdf())
    assert "All Temperature Windows; collections delle" in lines, lines
    assert any(ln.split(" ")[0] == "Bold" for ln in lines) and "BBoolldd" not in "".join(lines), lines  # 本当の重ね書きは1字
    win = build(PdfDocument(_line_pdf(), "t.pdf"))
    assert win["pages"][0]["merged_glyphs"] == 4


def test_reading_audit_catches_glyphs_dropped_as_false_overprints(monkeypatch):
    blob = _line_pdf()
    doc = PdfDocument(blob, "t.pdf")
    ok = check_reading(doc, build(doc))[0]
    assert ok.ok, ok.items
    assert "4 字" in ok.detail  # 本当の重ね書きを、字の位置と送りの幅から確かめて数える

    # これまでの判定（字の中心の差が字の大きさの 1/4 より小さければ重ね書き）に戻すと、「ll」の片方が落ちる。検査はそれを見つける
    def old_rule(a, b, size_a, size_b):
        return abs(a[0] - b[0]) < 0.25 * size_b and abs(a[1] - b[1]) < 0.25 * size_b and abs(size_a - size_b) < 0.2 * size_b

    monkeypatch.setattr(pdf_layout, "is_overprint", old_rule)
    doc = PdfDocument(blob, "t.pdf")
    bad = check_reading(doc, build(doc))[0]
    assert not bad.ok and any("「l」" in x and "重ねた字ではありません" in x for x in bad.items), bad.items


def test_literal_underscores_stay_in_the_window():
    c = b"BT /F1 11 Tf 40 700 Td (HWL_Spawn _HVeg_AllTs) Tj ET\n"
    c += b"BT /F1 11 Tf 40 650 Td (Name: ________ ) Tj ET\n"
    blob = make_pdf(c, {"F1": HELV})
    lines = _reading(blob)
    joined = "\n".join(lines)
    # 英数字に接した「_」は字のまま（候補にするときも、字は消さずに後ろに印を置く）
    assert "HWL_Spawn _" in joined and "HVeg_AllTs" in joined, lines
    # 英数字に接していない下線の並びは、書くための線（規則の欄）。下線の字は窓に残す
    assert "________［p1." in joined, lines
    doc = PdfDocument(blob, "t.pdf")
    assert check_reading(doc, build(doc))[0].ok  # 印に置き換えて消えた字があれば、読み落としとして見つかる


def _shifted_xref_pdf() -> bytes:
    """相互参照の表が「1 n」から始まるのに、最初の行が 0 番（空き）の項目：番号が1つずつずれて、どの位置も違う番号を指す。"""
    content = b"BT /F1 12 Tf 72 700 Td (Hello shifted xref) Tj ET"
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream", HELV]
    out = bytearray(b"%PDF-1.3\n")
    offs = []
    for i, body in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n1 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for o in offs:
        out += f"{o:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    return bytes(out)


def test_shifted_xref_finds_pages_and_can_be_written(tmp_path):
    from todoku import api

    blob = _shifted_xref_pdf()
    doc = PdfDocument(blob, "x.pdf")
    assert len(doc.pages()) == 1 and doc.repaired
    assert "Hello shifted xref" in _reading(blob)
    g = grade_blob(blob, "x.pdf")
    assert g.grade == "A", g.headline()
    src = tmp_path / "x.pdf"
    src.write_bytes(blob)
    res = api.fill(str(src), {"@1": {"after": "Hello", "value": "OK", "page": 1, "where": "below"}}, str(tmp_path / "y.pdf"))
    assert res["proof_ok"] and res["read_back_ok"], res
    out = (tmp_path / "y.pdf").read_bytes()
    assert out.startswith(blob)
    assert not PdfDocument(out, "y.pdf").repaired  # 足した相互参照に、すべてのオブジェクトの正しい位置を入れた


def test_pdf_without_pages_is_f_with_a_reason():
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [] /Count 0 >>"]
    out = bytearray(b"%PDF-1.4\n")
    offs = []
    for i, body in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(out)
    out += b"xref\n0 3\n0000000000 65535 f \n" + b"".join(f"{o:010d} 00000 n \n".encode() for o in offs)
    out += f"trailer\n<< /Size 3 /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    g = grade_blob(bytes(out), "e.pdf")
    assert g.grade == "F" and g.reasons[0].title == "ページがない" and g.write == "D", g.headline()


def test_field_candidates_outside_the_page_are_dropped():
    # ページ（MediaBox 200×200）の外に、見出しと下線を描く。ページの中にも同じ形を描く
    c = b"BT /F1 10 Tf 10 100 Td (Name) Tj ET\n0.8 w 45 98 m 150 98 l S\n"
    c += b"BT /F1 10 Tf -150 -40 Td (Name) Tj ET\n0.8 w -115 -42 m -10 -42 l S\n"
    blob = make_pdf(c, {"F1": HELV}).replace(b"/MediaBox [0 0 595 842]", b"/MediaBox [0 0 200 200]")
    win = build(PdfDocument(blob, "o.pdf"))
    fields = win["pages"][0]["fields"]
    assert fields, win["pages"][0]["reading"]
    for f in fields:
        x0, y0, x1, y1 = f["rect"]
        assert 0 <= x0 < x1 <= 200 and 0 <= y0 < y1 <= 200, f


def _unmapped_pdf(body: bool) -> bytes:
    """字に直せない字（ToUnicode にない符号 0x00FF）を、行の頭の印として（body=False）か、語の中に（body=True）置く。"""
    chars = "経済成長率は高い"
    font, extra, codes = jp_font_objs(chars, 5)
    t = lambda s: jp_text(codes, s)  # noqa: E731
    c = bytearray()
    if body:
        c += b"BT /F1 10 Tf 70 700 Td " + t("経済") + b" Tj <00FF> Tj " + t("成長率は高い") + b" Tj ET\n"
    else:
        for k in range(3):
            y = 700 - 20 * k
            c += f"BT /F1 10 Tf 70 {y} Td <00FF> Tj ET\n".encode()
            c += f"BT /F1 10 Tf 90 {y} Td ".encode() + t("経済成長率は高い") + b" Tj ET\n"
    return make_pdf(bytes(c), {"F1": font}, extra)


def test_unmappable_list_markers_are_alerts_not_grades():
    g = grade_blob(_unmapped_pdf(False), "m.pdf")
    assert g.grade == "A", g.headline()
    al = [a for a in g.alerts if a.title == "字に直せない行の頭の印"]
    assert al and "3 字" in al[0].detail and len(al[0].where) == 3 and "経済成長率は高い" in al[0].where[0], g.alerts


def test_unmappable_body_text_is_e_with_places():
    g = grade_blob(_unmapped_pdf(True), "b.pdf")
    assert g.grade == "E" and g.reasons[0].title == "字に直せない本文の字", g.headline()
    assert "経済" in g.reasons[0].where[0]


def test_scan_with_undecodable_compression_is_e():
    # 画像だけのページで、画像の圧縮が JBIG2（todoku は開けない）：書類はスキャンなので E。書く所は確かめられないので D
    img = b"<< /Type /XObject /Subtype /Image /Width 8 /Height 8 /ColorSpace /DeviceGray /BitsPerComponent 1 /Filter /JBIG2Decode /Length 4 >>\nstream\n\x00\x00\x00\x00\nendstream"
    content = b"q 595 0 0 842 0 0 cm /Im1 Do Q"
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>", b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream", img]
    out = bytearray(b"%PDF-1.5\n")
    offs = []
    for i, body in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + body + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode() + b"".join(f"{o:010d} 00000 n \n".encode() for o in offs)
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n".encode()
    g = grade_blob(bytes(out), "s.pdf")
    assert g.grade == "E" and "JBIG2Decode" in g.reasons[0].detail and g.write == "D", g.headline()


def test_csv_with_nul_in_values_is_still_data():
    g = grade_blob("名前;値\nA;1\x00\x00\nB;2\n".encode("utf-8"), "n.csv")
    assert g.grade == "S" and any(r.title == "値の中の NUL" for r in g.reasons), g.headline()


def test_read_grade_ignores_field_candidates_of_a_statistics_table():
    # 罫線の表で、見出しの行と空のマスのある統計表：空のマスは書く所かもしれない候補（書く段階の B）だが、中身は読めるので A
    c = bytearray(b"0.8 w ")
    for k in range(5):
        y = 700 - 20 * k
        c += f"60 {y} m 400 {y} l S ".encode()
    for x in (60, 200, 300, 400):
        c += f"{x} 700 m {x} 620 l S ".encode()
    c += b"\nBT /F1 10 Tf 70 685 Td (Region) Tj 140 0 Td (2023) Tj 100 0 Td (2024) Tj ET\n"
    c += b"BT /F1 10 Tf 70 665 Td (North) Tj 140 0 Td (12) Tj ET\n"
    c += b"BT /F1 10 Tf 70 645 Td (South) Tj 140 0 Td (7) Tj 100 0 Td (9) Tj ET\n"
    g = grade_blob(make_pdf(bytes(c), {"F1": HELV}), "t.pdf")
    assert g.grade == "A" and not g.reasons, g.headline()
    assert g.write == "B"


def test_chart_or_misplaced_print_area_is_not_an_image_only_sheet():
    from types import SimpleNamespace

    from todoku.grade import image_only_sheets

    cell = lambda v: SimpleNamespace(value=v, has_formula=False)  # noqa: E731
    pic = SimpleNamespace(kind="picture", text="", anchor=(1, 1))
    chart = SimpleNamespace(kind="chart", text="", anchor=(1, 1))

    def sheet(name, cells, shapes, area=None):
        inside = (lambda r, c: True) if area is None else (lambda r, c: area[0] <= r <= area[2] and area[1] <= c <= area[3])
        return SimpleNamespace(name=name, shapes=list(shapes), controls=[], cells=cells, in_print_area=inside)

    wb = SimpleNamespace(sheets=[
        sheet("グラフ", {}, [chart]),  # グラフだけ（値は別のシート）：画像ではない
        sheet("範囲の外", {(5, 2): cell("値")}, [pic], area=(101, 2, 119, 10)),  # 印刷範囲が絵にもセルにも重ならない
        sheet("絵", {(200, 1): cell("欄外の控え")}, [pic], area=(1, 1, 40, 10)),  # 印刷範囲の中は絵だけ
    ])
    assert image_only_sheets(wb) == ["絵"]


def test_tounicode_helper_still_used():
    # 部品の確かめ（test_pdf の組み立てを使う）
    assert b"beginbfchar" in tounicode({1: "あ"}) and zlib.compress(b"x")
