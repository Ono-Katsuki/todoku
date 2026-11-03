"""監査用の合成 PDF（AcroForm）を、手で書いたオブジェクトから作る（外部のライブラリなし）。

作るもの（out_dir に）：
  acro_all.pdf   1枚目：一行の欄・複数行・櫛（MaxLen 8・Comb）・中央揃え・字の大きさ 0（自動）・
                 チェック（書き出しの値が /Ja）・ラジオ（/A /B /C）・コンボ（選択肢に書き出しの値と表示が別）・リスト・
                 AFNumber の書式（JavaScript）・AFDate の書式・計算する欄（AFSimple_Calculate）・読み取り専用・必須・
                 階層の名前（applicant.name.family / applicant.name.given）・署名の欄（空）
                 2枚目：同じ名前の欄（同じ値を2ページに出す kids）と、2ページ目だけの欄
  need_app.pdf   見た目（/AP）のない欄だけで、/NeedAppearances true
"""
import sys
import zlib
from pathlib import Path


def pdf(objs: list[bytes], root: int) -> bytes:
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + o + b"\nendobj\n"
    x = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for o in offs:
        out += f"{o:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root {root} 0 R >>\nstartxref\n{x}\n%%EOF\n".encode()
    return bytes(out)


def stream(d: str, data: bytes) -> bytes:
    return f"<< {d} /Length {len(data)} >>\nstream\n".encode() + data + b"\nendstream"


def build_all(with_ap: bool = True, need_app: bool = False) -> bytes:
    objs: list[bytes] = []

    def add(b) -> int:
        objs.append(b if isinstance(b, bytes) else b.encode("latin-1"))
        return len(objs)

    def set_(i, b):
        objs[i - 1] = b if isinstance(b, bytes) else b.encode("latin-1")

    catalog = add("")
    pages = add("")
    font = add("<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>")
    zadb = add("<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>")
    dr = f"<< /Font << /Helv {font} 0 R /ZaDb {zadb} 0 R >> >>"

    def ap_text(w, h, txt=""):
        body = f"/Tx BMC q BT /Helv 10 Tf 2 {h / 2 - 3:.1f} Td ({txt}) Tj ET Q EMC".encode()
        return add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 {w} {h}] /Resources {dr}", body))

    def ap_box(w, h, on: bool):
        body = (b"q 0 0 1 rg BT /ZaDb 10 Tf 2 2 Td (4) Tj ET Q" if on else b"")
        return add(stream(f"/Type /XObject /Subtype /Form /BBox [0 0 {w} {h}] /Resources {dr}", body))

    # 1枚目の中身：見出しの字
    labels = [(50, 760, "Family name"), (50, 730, "Notes (multiline)"), (50, 680, "Postcode (comb 8)"),
              (50, 650, "Centered"), (50, 620, "Autosize"), (50, 590, "Agree"), (50, 560, "Plan A / B / C"),
              (50, 530, "Prefecture (combo)"), (50, 480, "Colour (list)"), (50, 440, "Amount (AFNumber)"),
              (50, 410, "Date (AFDate yyyy/mm/dd)"), (50, 380, "Qty"), (50, 350, "Price"), (50, 320, "Total (calc)"),
              (50, 290, "Office ID (read-only)"), (50, 260, "Phone (required)"), (50, 230, "Given name (applicant.name.given)"),
              (50, 160, "Signature")]
    c1 = "BT /F1 10 Tf " + " ".join(f"1 0 0 1 {x} {y} Tm ({t}) Tj" for x, y, t in labels) + " ET"
    c1s = add(stream("", c1.encode()))
    c2s = add(stream("", b"BT /F1 10 Tf 1 0 0 1 50 760 Tm (Family name again) Tj 1 0 0 1 50 700 Tm (Page 2 only) Tj ET"))
    p1 = add("")
    p2 = add("")
    fields_top = []
    annots1, annots2 = [], []

    def widget(rect, page, extra, parent=None, ap=None):
        d = f"<< /Type /Annot /Subtype /Widget /Rect [{rect}] /P {page} 0 R /F 4 /MK << /BC [0 0 0] >> {extra}"
        if parent:
            d += f" /Parent {parent} 0 R"
        if ap and with_ap:
            d += f" /AP << /N {ap} >>"
        return add(d + " >>")

    def text(name, rect, page, extra="", da="/Helv 10 Tf 0 g", value=None, w=150, h=20):
        v = f" /V ({value})" if value is not None else ""
        ap = f"{ap_text(w, h, value or '')} 0 R"
        i = widget(rect, page, f"/FT /Tx /T ({name}) /DA ({da}){v} {extra}", ap=ap)
        fields_top.append(i)
        (annots1 if page == p1 else annots2).append(i)
        return i

    text("family", "200 755 350 775", p1)
    text("notes", "200 700 450 745", p1, "/Ff 4096", w=250, h=45)
    text("postcode", "200 675 360 695", p1, "/Ff 16777216 /MaxLen 8", w=160)
    text("centered", "200 645 350 665", p1, "/Q 1")
    text("autosize", "200 615 350 635", p1, da="/Helv 0 Tf 0 g")
    # チェック：書き出しの値が /Ja
    on, off = ap_box(14, 14, True), ap_box(14, 14, False)
    i = widget("200 585 214 599", p1, f"/FT /Btn /T (agree) /V /Off /AS /Off /DA (/ZaDb 0 Tf 0 g) /MK << /CA (4) >>",
               ap=f"<< /Ja {on} 0 R /Off {off} 0 R >>")
    fields_top.append(i); annots1.append(i)
    # ラジオ：親と3つの子
    radio = add("")
    kids = []
    for n, (st, x) in enumerate((("A", 200), ("B", 240), ("C", 280))):
        on_, off_ = ap_box(14, 14, True), ap_box(14, 14, False)
        k = widget(f"{x} 555 {x + 14} 569", p1, "/AS /Off /MK << /CA (l) >>", parent=radio,
                   ap=f"<< /{st} {on_} 0 R /Off {off_} 0 R >>")
        kids.append(k); annots1.append(k)
    set_(radio, f"<< /FT /Btn /Ff 49152 /T (plan) /V /Off /Kids [{' '.join(f'{k} 0 R' for k in kids)}] /DA (/ZaDb 0 Tf 0 g) >>")
    fields_top.append(radio)
    # コンボ：書き出しの値と表示が別
    i = widget("200 525 350 545", p1, "/FT /Ch /Ff 131072 /T (pref) /Opt [[(13) (Tokyo)] [(27) (Osaka)] [(01) (Hokkaido)]] /DA (/Helv 10 Tf 0 g)",
               ap=f"{ap_text(150, 20)} 0 R")
    fields_top.append(i); annots1.append(i)
    i = widget("200 470 350 520", p1, "/FT /Ch /T (colour) /Opt [(Red) (Green) (Blue)] /DA (/Helv 10 Tf 0 g)", ap=f"{ap_text(150, 50)} 0 R")
    fields_top.append(i); annots1.append(i)
    # JavaScript の書式
    fmt = add("<< /S /JavaScript /JS (AFNumber_Format\\(2, 0, 0, 0, \"\\\\u00a5\", true\\);) >>")
    kst = add("<< /S /JavaScript /JS (AFNumber_Keystroke\\(2, 0, 0, 0, \"\\\\u00a5\", true\\);) >>")
    text("amount", "200 435 350 455", p1, f"/AA << /F {fmt} 0 R /K {kst} 0 R >>")
    dfmt = add("<< /S /JavaScript /JS (AFDate_FormatEx\\(\"yyyy/mm/dd\"\\);) >>")
    dk = add("<< /S /JavaScript /JS (AFDate_KeystrokeEx\\(\"yyyy/mm/dd\"\\);) >>")
    text("date", "200 405 350 425", p1, f"/AA << /F {dfmt} 0 R /K {dk} 0 R >>")
    qty = text("qty", "200 375 350 395", p1, value="2")
    price = text("price", "200 345 350 365", p1, value="100")
    calc = add("<< /S /JavaScript /JS (AFSimple_Calculate\\(\"PRD\", new Array \\(\"qty\", \"price\"\\)\\);) >>")
    total = text("total", "200 315 350 335", p1, f"/AA << /C {calc} 0 R >>", value="200")
    text("office_id", "200 285 350 305", p1, "/Ff 1", value="X-123")
    text("phone", "200 255 350 275", p1, "/Ff 2")
    # 階層の名前：applicant → name → family（2ページに kids）・given
    app = add("")
    nm = add("")
    fam_parent = add("")
    fk1 = widget("200 755 350 775", p2, "", parent=fam_parent, ap=f"{ap_text(150, 20)} 0 R")
    annots2.append(fk1)
    giv = widget("200 225 350 245", p1, "/FT /Tx /T (given) /DA (/Helv 10 Tf 0 g)", parent=nm, ap=f"{ap_text(150, 20)} 0 R")
    annots1.append(giv)
    fk0 = widget("400 225 550 245", p1, "", parent=fam_parent, ap=f"{ap_text(150, 20)} 0 R")
    annots1.append(fk0)
    set_(fam_parent, f"<< /FT /Tx /T (family) /Parent {nm} 0 R /DA (/Helv 10 Tf 0 g) /Kids [{fk0} 0 R {fk1} 0 R] >>")
    set_(nm, f"<< /T (name) /Parent {app} 0 R /Kids [{fam_parent} 0 R {giv} 0 R] >>")
    set_(app, f"<< /T (applicant) /Kids [{nm} 0 R] >>")
    fields_top.append(app)
    # 2ページ目だけの欄
    text("page2only", "200 695 350 715", p2)
    # 署名の欄（空）
    i = widget("200 140 400 180", p1, "/FT /Sig /T (sig1)")
    fields_top.append(i); annots1.append(i)
    co = add(f"<< /CO [{total} 0 R] >>")  # 計算の順
    res = f"<< /Font << /F1 {font} 0 R >> >>"
    set_(p1, f"<< /Type /Page /Parent {pages} 0 R /MediaBox [0 0 595 842] /Contents {c1s} 0 R /Resources {res} "
             f"/Annots [{' '.join(f'{a} 0 R' for a in annots1)}] >>")
    set_(p2, f"<< /Type /Page /Parent {pages} 0 R /MediaBox [0 0 595 842] /Contents {c2s} 0 R /Resources {res} "
             f"/Annots [{' '.join(f'{a} 0 R' for a in annots2)}] >>")
    set_(pages, f"<< /Type /Pages /Kids [{p1} 0 R {p2} 0 R] /Count 2 >>")
    na = " /NeedAppearances true" if need_app else ""
    set_(catalog, f"<< /Type /Catalog /Pages {pages} 0 R /AcroForm << /Fields [{' '.join(f'{f} 0 R' for f in fields_top)}] "
                  f"/DR {dr} /DA (/Helv 0 Tf 0 g) /CO [{total} 0 R]{na} >> >>")
    del co
    return pdf(objs, catalog)


if __name__ == "__main__":
    out = Path(sys.argv[1])
    out.mkdir(parents=True, exist_ok=True)
    (out / "acro_all.pdf").write_bytes(build_all())
    (out / "need_app.pdf").write_bytes(build_all(with_ap=False, need_app=True))
    print("ok")
