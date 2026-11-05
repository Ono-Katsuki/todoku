"""第5回の見直しで見つかった、PDF への誤った書き込みの直し（合成した書類で確かめる）。

1a 回した欄（/MK の R 90・270）と、回したページ（/Rotate 90・270）に足す欄：見た目の BBox を描く向きの大きさにし、
   Matrix で Rect にちょうど重ねる（ISO 32000-1 12.5.5 Algorithm 8.1）。証明は、BBox を Matrix で写した大きさが Rect と
   同じか・回す向きが /MK の R と同じかを確かめる。view_page は、見た目の字を、見た目のとおりの位置・大きさ・向きで描く
3a 作り手の権限の設定（暗号の辞書の P、表 22）が記入を許していない：断らずに書き、本人に伝え、書く段階を下げる
4a UserUnit：mm とページの座標の換算（枠・印影の大きさ・view_page・知らせる字の大きさ）に UserUnit を入れる
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))
from pdf_forms_maker import pdf, stream  # noqa: E402

from todoku import api  # noqa: E402
from todoku.pdf.doc import PdfDocument  # noqa: E402
from todoku.pdf.fill import Update, _rot_matrix, fill, text_appearance  # noqa: E402
from todoku.pdf.layout import MM_PT, View  # noqa: E402
from todoku.pdf.pageimage import render  # noqa: E402
from todoku.pdf.proof import prove  # noqa: E402

HELV = "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"


def _pdf(rotate: int = 0, widget: str | None = None, user_unit: float | None = None, need_app: bool = False,
         text: bytes = b"BT /Helv 10 Tf 60 760 Td (Name:) Tj ET", media: str = "0 0 595 842") -> bytes:
    """1 ページ（回転・UserUnit・欄 1 つを指定できる）。widget は欄の辞書の中身（/P と /Type・/Subtype は足す）。"""
    objs: list[bytes] = []

    def add(b):
        objs.append(b if isinstance(b, bytes) else b.encode("latin-1"))
        return len(objs)

    cat, pages, page = add(""), add(""), add("")
    font = add(HELV)
    zadb = add("<< /Type /Font /Subtype /Type1 /BaseFont /ZapfDingbats >>")
    cs = add(stream("", text))
    annots, fields = "", ""
    if widget is not None:
        w = add(f"<< /Type /Annot /Subtype /Widget {widget} /F 4 /P {page} 0 R >>")
        annots, fields = f" /Annots [{w} 0 R]", f"{w} 0 R"
    uu = f" /UserUnit {user_unit}" if user_unit is not None else ""
    objs[page - 1] = (f"<< /Type /Page /Parent {pages} 0 R /MediaBox [{media}] /Rotate {rotate}{uu} /Contents {cs} 0 R "
                      f"/Resources << /Font << /Helv {font} 0 R >> >>{annots} >>").encode()
    objs[pages - 1] = f"<< /Type /Pages /Kids [{page} 0 R] /Count 1 >>".encode()
    na = " /NeedAppearances true" if need_app else ""
    objs[cat - 1] = (f"<< /Type /Catalog /Pages {pages} 0 R /AcroForm << /Fields [{fields}] "
                     f"/DR << /Font << /Helv {font} 0 R /ZaDb {zadb} 0 R >> >> /DA (/Helv 0 Tf 0 g){na} >> >>").encode()
    return pdf(objs, cat)


def _widget(doc: PdfDocument, name: str) -> dict:
    from todoku.pdf import pagedraw

    # 値の控えは、注釈でなく /Fields にある
    for a in list(doc.resolve(doc.pages()[0].get("Annots")) or []) + [r for r, _wd in pagedraw.records(doc)]:
        d = doc.resolve(a)
        t = d.get("T")
        if t is not None and bytes(t).decode("latin-1") == name:
            return d
    raise AssertionError(name)


def _mapped(doc: PdfDocument, wd: dict, state: str | None = None) -> tuple[list, list, list]:
    """(Rect, BBox を Matrix で写した四角, Matrix)。"""
    n = doc.resolve(doc.resolve(wd["AP"])["N"])
    if state is not None:
        n = doc.resolve(n[state])
    bbox = [float(v) for v in doc.resolve(n.dict["BBox"])]
    m = [float(v) for v in doc.resolve(n.dict.get("Matrix", [1, 0, 0, 1, 0, 0]))]
    pts = [(m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]) for x in (bbox[0], bbox[2]) for y in (bbox[1], bbox[3])]
    box = [min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts)]
    return [float(v) for v in doc.resolve(wd["Rect"])], box, m


def _same_size(rect, box) -> bool:
    return abs((rect[2] - rect[0]) - (box[2] - box[0])) < 0.01 and abs((rect[3] - rect[1]) - (box[3] - box[1])) < 0.01


def _ink_mm(r: dict) -> tuple[float, float, float, float] | None:
    """描いた画像の、白くない画素の外枠（ページの左上からの mm）。"""
    w, h, rgb = r["w"], r["h"], r["rgb"]
    xs, ys = [], []
    for y in range(h):
        row = rgb[y * w * 3:(y + 1) * w * 3:3]
        for x, v in enumerate(row):
            if v < 250:
                xs.append(x)
                ys.append(y)
    if not xs:
        return None
    s, (ox, oy) = r["scale"], r["origin_mm"]
    return ox + min(xs) / s, oy + min(ys) / s, ox + (max(xs) + 1) / s, oy + (max(ys) + 1) / s


# ====================================================================== 1a 回した欄の見た目
@pytest.mark.parametrize("rot", [90, 180, 270])
def test_rotated_acroform_field_appearance_fits_rect(rot):
    # ページの字は置かない（ページの画像は字も描くので、欄の字だけの外枠を見る）
    blob = _pdf(rot, "/FT /Tx /T (name) /Rect [100 542 130 742] /DA (/Helv 10 Tf 0 g) /MK << /R %d >>" % rot
                if rot != 180 else "/FT /Tx /T (name) /Rect [100 542 300 572] /DA (/Helv 10 Tf 0 g) /MK << /R 180 >>",
                text=b"")
    res = fill(PdfDocument(blob), {"name": "ABC Yamada"})
    doc = PdfDocument(res.blob)
    rect, box, m = _mapped(doc, _widget(doc, "name"))
    assert _same_size(rect, box), (rect, box)
    ang = math.radians(rot)
    assert [round(v, 6) for v in m[:4]] == [round(math.cos(ang)), round(math.sin(ang)), -round(math.sin(ang)), round(math.cos(ang))]
    pr = prove(blob, res.blob, {"name"})
    assert pr.ok, pr.unexpected
    # 見た目のとおりに描くと、字は欄の中に、横に並ぶ（つぶれない・横倒しでない）
    r = render(doc, 1)
    ink = _ink_mm(r)
    view = View(doc, doc.pages()[0])
    vx0, vy0, vx1, vy1 = (v / view.mm for v in view.rect(*rect))
    assert ink is not None and r["glyphs"] == len("ABCYamada")
    assert vx0 - 0.3 <= ink[0] and ink[2] <= vx1 + 0.3 and vy0 - 0.3 <= ink[1] and ink[3] <= vy1 + 0.3, (ink, (vx0, vy0, vx1, vy1))
    assert ink[2] - ink[0] > 3 * (ink[3] - ink[1])  # 横書きの 1 行（字の高さより、ずっと幅が広い）


def test_rot_matrix_maps_view_box_exactly_onto_rect():
    for rot in (90, 180, 270):
        w, h = 30.0, 200.0
        vw, vh = (h, w) if rot in (90, 270) else (w, h)
        m = _rot_matrix(rot, w, h)
        pts = [(m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]) for x in (0, vw) for y in (0, vh)]
        assert (min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts)) == (0, 0, w, h)


def test_comb_appearance_on_rotated_field_uses_view_box():
    doc = PdfDocument(_pdf())
    up = Update(doc)
    fonts = {"Helv": up.new({})}
    comb = [(i * 20.0, 0, (i + 1) * 20.0, 30.0) for i in range(5)]  # 描く向き：幅 100・高さ 30
    ref, _s = text_appearance(up, 30.0, 100.0, "12345", 10, fonts, comb=comb, rot=90)
    st = up.objs[ref.num][1]
    assert st.dict["BBox"] == [0, 0, 100.0, 30.0] and st.dict["Matrix"] == [0, 1, -1, 0, 30.0, 0]


@pytest.mark.parametrize("rot", [90, 270])
def test_checkbox_without_appearance_on_rotated_widget(rot):
    blob = _pdf(rot, "/FT /Btn /T (ok) /Rect [100 500 112 520] /MK << /R %d /CA (4) >>" % rot, need_app=True)
    res = fill(PdfDocument(blob), {"ok": True})
    doc = PdfDocument(res.blob)
    wd = _widget(doc, "ok")
    for st in doc.resolve(doc.resolve(wd["AP"])["N"]):
        rect, box, m = _mapped(doc, wd, st)
        assert _same_size(rect, box)
        assert m[1] == (1 if rot == 90 else -1)
    assert prove(blob, res.blob, {"ok"}).ok


@pytest.mark.parametrize("rot", [90, 270])
def test_box_write_and_checkbox_on_rotated_page(tmp_path, rot):
    src, out = tmp_path / "r.pdf", tmp_path / "o.pdf"
    src.write_bytes(_pdf(rot))
    r = api.fill(str(src), {"@1": {"page": 1, "box": [60, 30, 120, 40], "value": "ABC 山田"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    doc = PdfDocument(out.read_bytes())
    rect, box, m = _mapped(doc, _widget(doc, "p1.w1"))
    assert _same_size(rect, box)
    assert doc.resolve(_widget(doc, "p1.w1")["MK"])["R"] == rot
    ink = _ink_mm(render(doc, 1, [50, 20, 130, 50]))
    assert ink is not None and 60 - 0.3 <= ink[0] and ink[2] <= 120.3 and 30 - 0.3 <= ink[1] and ink[3] <= 40.3
    assert ink[2] - ink[0] > 3 * (ink[3] - ink[1])


def _swap_once(blob: bytes, old: bytes, new: bytes) -> bytes:
    assert len(old) == len(new) and blob.count(old) == 1
    return blob.replace(old, new)


def test_proof_catches_a_stretched_or_misrotated_appearance():
    blob = _pdf(90, "/FT /Tx /T (name) /Rect [100 542 130 742] /DA (/Helv 10 Tf 0 g) /MK << /R 90 >>")
    res = fill(PdfDocument(blob), {"name": "ABC"})
    assert prove(blob, res.blob, {"name"}).ok
    tail = res.blob[len(blob):]
    # 以前の誤り：BBox を Rect の幅と高さのままにした（Algorithm 8.1 で、縦横別の倍率に伸び縮みする）
    bad = blob + _swap_once(tail, b"/BBox [0 0 200 30]", b"/BBox [0 0 30 200]")
    pr = prove(blob, bad, {"name"})
    assert not pr.ok and any("Algorithm 8.1" in u for u in pr.unexpected), pr.unexpected
    # 回す向きが /MK の R と違う（横倒しに見える）
    bad = blob + _swap_once(tail, b"/Matrix [0 1 -1 0 30 0]", b"/Matrix [0 -1 1 0 0 30]")
    pr = prove(blob, bad, {"name"})
    assert not pr.ok and any("90 度回す向き" in u for u in pr.unexpected), pr.unexpected


# ====================================================================== 3a 権限の設定（P）
def _encrypted(P: int, rev: int = 4, user: bytes = b"") -> bytes:
    import test_pdf_password as tpp

    old = tpp.P
    tpp.P = P
    try:
        return tpp.make_pdf(tpp.Enc(rev=rev, user=user, owner=b"owner-pass", cfm="AESV2" if rev == 4 else "V2"))
    finally:
        tpp.P = old


PRINT_ONLY = -3900  # 印刷だけ（変更・注釈と記入・欄への記入を許さない）
FILL_ONLY = -3900 | (1 << 8)  # ビット 9（欄への記入）だけ許す


def test_fill_on_a_form_whose_permissions_forbid_filling_is_told(tmp_path):
    src, out = tmp_path / "p.pdf", tmp_path / "o.pdf"
    src.write_bytes(_encrypted(PRINT_ONLY))
    r = api.fill(str(src), {"name": "Hanako"}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    assert [d["op"] for d in r["permissions_denied"]] == ["fill"]
    assert "権限の設定" in r["warnings"][0] and "権限の設定" in r["say_to_user"]
    g = api.grade(str(src))
    assert g["write"]["grade"] == "B" and any(x["title"].startswith("権限の設定") for x in g["write"]["reasons"])
    # 開くためのパスワードのある PDF：開くためのパスワードで開けば制限を受け、権限のパスワードで開けば受けない
    src.write_bytes(_encrypted(PRINT_ONLY, user=b"user-pass"))
    r = api.fill(str(src), {"name": "Hanako"}, str(tmp_path / "o2.pdf"), password="user-pass")
    assert [d["op"] for d in r["permissions_denied"]] == ["fill"]
    r = api.fill(str(src), {"name": "Hanako"}, str(tmp_path / "o3.pdf"), password="owner-pass")
    assert "permissions_denied" not in r and not any("権限の設定" in w for w in r["warnings"])
    assert r["proof_ok"] and r["read_back_ok"]


def test_fill_bit_allows_existing_fields_but_not_new_ones(tmp_path):
    src = tmp_path / "p.pdf"
    src.write_bytes(_encrypted(FILL_ONLY))
    r = api.fill(str(src), {"name": "Hanako"}, str(tmp_path / "o.pdf"))
    assert "permissions_denied" not in r
    r = api.fill(str(src), {"@1": {"page": 1, "box": [30, 100, 90, 110], "value": "X"}}, str(tmp_path / "o2.pdf"))
    assert [d["op"] for d in r["permissions_denied"]] == ["create"]
    # R 2 にはビット 9 がない：ビット 6 がなければ、欄への記入も許していない
    src.write_bytes(_encrypted(FILL_ONLY, rev=2))
    r = api.fill(str(src), {"name": "Hanako"}, str(tmp_path / "o3.pdf"))
    assert [d["op"] for d in r["permissions_denied"]] == ["fill"]


def test_unrestricted_pdf_has_no_permission_note(tmp_path):
    src = tmp_path / "p.pdf"
    src.write_bytes(_encrypted(-4))  # すべて許す
    r = api.fill(str(src), {"name": "Hanako"}, str(tmp_path / "o.pdf"))
    assert "permissions_denied" not in r
    assert not any(x["title"].startswith("権限の設定") for x in api.grade(str(src))["write"]["reasons"])


# ====================================================================== 4a UserUnit
def test_user_unit_box_seal_and_view_page(tmp_path):
    from test_seal import ring_rgba

    src, out, png = tmp_path / "u.pdf", tmp_path / "o.pdf", tmp_path / "seal.png"
    src.write_bytes(_pdf(0, user_unit=2, media="0 0 297.64 420.945", text=b"BT /Helv 5 Tf 30 400 Td (Name:) Tj ET"))
    png.write_bytes(ring_rgba()[0])
    unit_mm = 2 * 25.4 / 72  # ページの座標の 1 単位の mm
    r = api.fill(str(src), {"@1": {"page": 1, "box": [20, 20, 80, 30], "value": "ABC"},
                            "@2": {"seal": str(png), "page": 1, "box": [20, 40, 60, 80], "consent": True}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], r
    doc = PdfDocument(out.read_bytes())
    view = View(doc, doc.pages()[0])
    assert view.unit == 2 and abs(view.mm - MM_PT / 2) < 1e-9
    rect = [float(v) for v in doc.resolve(_widget(doc, "p1.w1")["Rect"])]
    assert abs((rect[2] - rect[0]) * unit_mm - 60) < 0.01 and abs((rect[3] - rect[1]) * unit_mm - 10) < 0.01
    pl = r["placed"][0]
    assert pl["box_mm"] == [20.0, 40.0, 60.0, 80.0] and "横 40 mm・縦 40 mm" in pl["say"]  # 丸い画像を、枠いっぱいに
    # 字の大きさ：A4 の紙（210×297 mm）の標準 10.5pt は、ページの座標では 5.25
    da = bytes(_widget(doc, "p1.w1")["DA"]).decode()
    assert float(da.split()[1]) <= 5.25 + 1e-6
    vp = api.view_page(str(out), 1)
    assert "210×297 mm" in vp["text"] and "UserUnit 2" in vp["text"]
    ink = _ink_mm(render(doc, 1, [15, 15, 85, 35]))
    assert ink is not None and 20 - 0.3 <= ink[0] and ink[2] <= 80.3 and 20 - 0.3 <= ink[1] and ink[3] <= 30.3


def test_user_unit_seal_size_and_small_text_warning(tmp_path):
    from test_seal import ring_rgba

    src, png = tmp_path / "u.pdf", tmp_path / "seal.png"
    src.write_bytes(_pdf(0, user_unit=2, media="0 0 297.64 420.945",
                         text=b"BT /Helv 5 Tf 30 400 Td (Name:) Tj ET BT /Helv 5 Tf 30 300 Td (Sign) Tj ET "
                              b"BT /Helv 5 Tf 260 300 Td (End) Tj ET"))
    png.write_bytes(ring_rgba()[0])
    r = api.fill(str(src), {"@1": {"seal": str(png), "after": "Sign", "size": 12, "consent": True}}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"], r
    w = r["placed"][0]["box_mm"][2] - r["placed"][0]["box_mm"][0]
    assert abs(w - 12) < 0.1 and "横 12 mm・縦 12 mm" in r["placed"][0]["say"]
    # 小さな枠に長い字：知らせる字の大きさは、紙の上の pt（ページの座標の大きさ × 2）
    r = api.fill(str(src), {"@1": {"page": 1, "box": [20, 20, 60, 23], "value": "ABCDEFGHIJKLMNOPQRSTUVWXYZ"}},
                 str(tmp_path / "o2.pdf"))
    doc = PdfDocument((tmp_path / "o2.pdf").read_bytes())
    s = float(bytes(_widget(doc, "p1.w1")["DA"]).decode().split()[1])
    msg = next(x for x in r["warnings"] if "pt" in x)
    assert f"{s * 2:.1f} pt" in msg, (s, msg)
