"""テスト用の PDF（架空の様式。個人情報なし）を tests/fixtures/public/ に作る。標準ライブラリだけ。

AcroForm・欄のない紙の様式・スキャン（字のない画像だけ）・回転したページ・暗号（RC4 128 ビット、パスワード todoku）・
縦書き・右から左の字（ヘブライ・アラビア）・インド系の字（デーヴァナーガリー）。
"""

from __future__ import annotations

import zlib

from pdfgen import (Encryption, Pdf, box, catalog, checkbox, combo, glyph_hex, helvetica, japanese, page, stream,
                    text_field, tounicode_font, utf16hex)


def jt(font: str, size: float, x: float, y: float, s: str) -> str:
    return f"BT /{font} {size:g} Tf {x:g} {y:g} Td <{utf16hex(s)}> Tj ET\n"


FORM_ROWS = [("氏名", "name"), ("住所", "address"), ("生年月日", "birth"), ("電話番号", "tel")]


def _form_page(with_fields: bool, p: Pdf, pg: int, pages: int, helv: int, ja: int) -> tuple[bytes, list[int]]:
    c = "0 g 0.5 w\n" + jt("F2", 16, 200, 790, "補助金交付申請書")
    c += jt("F2", 10, 400, 765, "令和　年　月　日")
    annots = []
    y = 720
    for label, name in FORM_ROWS:
        c += box(50, y - 6, 150, y + 18) + box(150, y - 6, 545, y + 18)
        c += jt("F2", 11, 56, y, label)
        if with_fields:
            annots.append(text_field(p, pg, name, (152, y - 4, 543, y + 16), tooltip=label))
        y -= 30
    c += box(50, y - 6, 150, y + 18) + box(150, y - 6, 545, y + 18) + jt("F2", 11, 56, y, "区分")
    if with_fields:
        annots.append(combo(p, pg, "kind", (152, y - 4, 300, y + 16), ["A", "B", "C"], tooltip="区分"))
    y -= 30
    c += box(60, y, 72, y + 12) + jt("F2", 11, 80, y + 1, "上記の内容に相違ありません。")
    if with_fields:
        annots.append(checkbox(p, pg, "agree", (60, y, 72, y + 12), tooltip="同意"))
    c += "BT /F1 8 Tf 50 60 Td (Synthetic test form for todoku. No personal data.) Tj ET\n"
    return c.encode("latin-1"), annots


def acroform(enc: Encryption | None = None) -> bytes:
    p = Pdf()
    cat, pages, pg = p.reserve(), p.reserve(), p.reserve()
    helv, ja = helvetica(p), japanese(p)
    c, annots = _form_page(True, p, pg, pages, helv, ja)
    page(p, pg, pages, c, {"F1": helv, "F2": ja, "Helv": helv}, annots)
    catalog(p, cat, pages, [pg], annots, helv=helv, lang="ja-JP")
    return p.build(cat, enc)


def flat() -> bytes:
    """欄のない紙の様式（罫線の表と見出しだけ）が2ページ。"""
    p = Pdf()
    cat, pages, pg1, pg2 = p.reserve(), p.reserve(), p.reserve(), p.reserve()
    helv, ja = helvetica(p), japanese(p)
    c, _ = _form_page(False, p, pg1, pages, helv, ja)
    page(p, pg1, pages, c, {"F1": helv, "F2": ja})
    c2 = "0 g 0.5 w\n" + jt("F2", 14, 230, 790, "事業計画書")
    y = 740
    for label in ("事業の名称", "実施の期間", "経費の総額"):
        c2 += box(50, y - 6, 180, y + 18) + box(180, y - 6, 545, y + 18) + jt("F2", 11, 56, y, label)
        y -= 30
    c2 += jt("F2", 10, 50, y - 10, "（注）金額は円単位で記入してください。")
    page(p, pg2, pages, c2.encode("latin-1"), {"F1": helv, "F2": ja})
    catalog(p, cat, pages, [pg1, pg2], lang="ja-JP")
    return p.build(cat)


def rotated() -> bytes:
    """/Rotate 90 のページ（横向きに見える様式）。字は 90 度回して置き、表示すると正立する。"""
    p = Pdf()
    cat, pages, pg = p.reserve(), p.reserve(), p.reserve()
    helv, ja = helvetica(p), japanese(p)
    c = "0 g 0.5 w\n"
    # 時計回りに 90 度回して表示するので、ユーザー空間の左（x が小さい所）が表示の上になる。
    # 字はユーザー空間で +y の向きに書く（表示では左から右）。Tm = [0 1 -1 0 x y]
    for i, s in enumerate(["受付票", "受付番号", "担当者名"]):
        x = 60 + 40 * i
        c += f"BT /F2 {14 if i == 0 else 11} Tf 0 1 -1 0 {x} 80 Tm <{utf16hex(s)}> Tj ET\n"
    c += "BT /F1 9 Tf 0 1 -1 0 200 80 Tm (Reference No.) Tj ET\n"
    annots = [text_field(p, pg, "receipt_no", (86, 180, 106, 600), tooltip="受付番号"),
              text_field(p, pg, "staff", (126, 180, 146, 600), tooltip="担当者名")]
    page(p, pg, pages, c.encode("latin-1"), {"F1": helv, "F2": ja, "Helv": helv}, annots, rotate=90)
    catalog(p, cat, pages, [pg], annots, helv=helv, lang="ja-JP")
    return p.build(cat)


def vertical() -> bytes:
    """縦書き（UniJIS-UTF16-V）。列は右から左。"""
    p = Pdf()
    cat, pages, pg = p.reserve(), p.reserve(), p.reserve()
    helv, jv = helvetica(p), japanese(p, vertical=True)
    cols = ["届出書", "私は次のとおり届け出ます。", "氏名", "本籍"]
    c = "0 g\n"
    x = 520
    for i, s in enumerate(cols):
        c += f"BT /F3 {16 if i == 0 else 12} Tf {x} 780 Td <{utf16hex(s)}> Tj ET\n"
        x -= 40
    c += "BT /F1 8 Tf 50 40 Td (Synthetic vertical-writing test form.) Tj ET\n"
    page(p, pg, pages, c.encode("latin-1"), {"F1": helv, "F3": jv})
    catalog(p, cat, pages, [pg], lang="ja-JP")
    return p.build(cat)


def _visual_font(p: Pdf, name: str, runs: list[list[str]]) -> tuple[int, list[list[int]]]:
    """runs：各行の字形（左から右に、見える順）。字形ごとに番号を振り、ToUnicode をつけた書体を作る。"""
    ids: dict[str, int] = {}
    out = []
    for r in runs:
        out.append([ids.setdefault(g, len(ids) + 1) for g in r])
    return tounicode_font(p, name, {v: k for k, v in ids.items()}), out


def rtl() -> bytes:
    """ヘブライ文字・アラビア文字の様式。字形は、見える順（左から右）に置く（多くの PDF と同じ）。"""
    p = Pdf()
    cat, pages, pg = p.reserve(), p.reserve(), p.reserve()
    helv = helvetica(p)
    he_title, he_name = "טופס בקשה", "שם מלא"
    ar_name = ["م", "س", "لا", "ا"]  # 「الاسم」を見える順に（ل と ا は合字の一つの字形）
    runs = [list(reversed(he_title)), list(reversed(he_name)), ar_name]
    font, glyphs = _visual_font(p, "SyntheticRTL", runs)
    c = "0 g 0.5 w\n"
    for (gl, y, size) in zip(glyphs, (790, 740, 700), (16, 12, 12)):
        c += f"BT /F4 {size} Tf {545 - len(gl) * size * 0.6:g} {y} Td <{glyph_hex(gl)}> Tj ET\n"
    c += box(200, 734, 450, 756) + box(200, 694, 450, 716)
    annots = [text_field(p, pg, "full_name", (202, 736, 448, 754), tooltip=he_name),
              text_field(p, pg, "ism", (202, 696, 448, 714), tooltip="الاسم")]
    c += "BT /F1 8 Tf 50 40 Td (Synthetic right-to-left test form.) Tj ET\n"
    page(p, pg, pages, c.encode("latin-1"), {"F1": helv, "F4": font, "Helv": helv}, annots)
    catalog(p, cat, pages, [pg], annots, helv=helv, lang="he")
    return p.build(cat)


def indic() -> bytes:
    """デーヴァナーガリー文字の様式。合字（त्र・न्द）は一つの字形、母音記号 ि は見える順で子音の前に置く。"""
    p = Pdf()
    cat, pages, pg = p.reserve(), p.reserve(), p.reserve()
    helv = helvetica(p)
    runs = [["आ", "व", "े", "द", "न", " ", "प", "त्र"],  # आवेदन पत्र
            ["न", "ा", "म"],  # नाम
            ["प", "त", "ा"],  # पता
            ["ि", "ह", "न्द", "ी"]]  # हिन्दी（見える順）
    font, glyphs = _visual_font(p, "SyntheticDeva", runs)
    c = "0 g 0.5 w\n"
    for gl, y, size in zip(glyphs, (790, 740, 700, 660), (16, 12, 12, 12)):
        c += f"BT /F5 {size} Tf 56 {y} Td <{glyph_hex(gl)}> Tj ET\n"
    c += box(150, 734, 450, 756) + box(150, 694, 450, 716)
    annots = [text_field(p, pg, "naam", (152, 736, 448, 754), tooltip="नाम"),
              text_field(p, pg, "pata", (152, 696, 448, 714), tooltip="पता")]
    c += "BT /F1 8 Tf 50 40 Td (Synthetic Devanagari test form.) Tj ET\n"
    page(p, pg, pages, c.encode("latin-1"), {"F1": helv, "F5": font, "Helv": helv}, annots)
    catalog(p, cat, pages, [pg], annots, helv=helv, lang="hi")
    return p.build(cat)


def scanned() -> bytes:
    """スキャンした紙（字のない灰色の画像だけ）。罫線の枠と、字の代わりの黒い塊を描く。"""
    w, h = 496, 702  # 60 dpi の A4
    px = bytearray([255]) * (w * h)

    def rect(x0, y0, x1, y1, v=0):
        for y in range(y0, y1):
            px[y * w + x0:y * w + x1] = bytes([v]) * (x1 - x0)

    def frame(x0, y0, x1, y1):
        rect(x0, y0, x1, y0 + 1)
        rect(x0, y1 - 1, x1, y1)
        rect(x0, y0, x0 + 1, y1)
        rect(x1 - 1, y0, x1, y1)

    rect(170, 40, 330, 54, 40)  # 表題の代わり
    for i in range(5):
        y = 100 + i * 28
        frame(40, y, 456, y + 26)
        rect(124, y, 125, y + 26)
        rect(48, y + 8, 100, y + 18, 60)
    img_data = zlib.compress(bytes(px), 9)
    p = Pdf()
    cat, pages, pg = p.reserve(), p.reserve(), p.reserve()
    img = p.add(f"<< /Type /XObject /Subtype /Image /Width {w} /Height {h} /ColorSpace /DeviceGray /BitsPerComponent 8 "
                f"/Filter /FlateDecode /Length {len(img_data)} >>\nstream\n".encode() + img_data + b"\nendstream")
    page(p, pg, pages, b"q 595 0 0 842 0 0 cm /Im0 Do Q\n", {}, xobjects={"Im0": img})
    catalog(p, cat, pages, [pg])
    return p.build(cat)


def encrypted() -> bytes:
    """acroform と同じ様式を、利用者のパスワード todoku で暗号にしたもの（RC4 128 ビット、R3）。"""
    return acroform(Encryption("todoku", "owner-only", bytes.fromhex("0123456789abcdef0123456789abcdef")))


ALL = {"acroform.pdf": acroform, "flat.pdf": flat, "rotated.pdf": rotated, "vertical.pdf": vertical,
       "rtl.pdf": rtl, "indic.pdf": indic, "scanned.pdf": scanned, "encrypted.pdf": encrypted}
