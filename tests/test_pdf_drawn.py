"""書いた字をページの描画に足す（pagedraw）：どのビューアーでも欠けずに見え、証明・読み戻し・窓・直しての書き直しが今までどおり働く。

Apple の PDFKit は文字の欄の見た目を使わず /V から描き直し、狭い欄では語の切れ目で折り返して字の終わりを描かない（21/46 の様式）。
Quartz は注釈を描かない。実物の表示での測り方は tests/audit_findings/pdfkit/。ここでは、ファイルの形と、とどく自身の読み書きを確かめる。
"""

from __future__ import annotations

import re
import os
import shutil
import subprocess

import pytest

from pdfmini import build, stream
from todoku import api
from todoku.pdf import fill as F, pagedraw
from todoku.pdf.content import page_content
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import FillError, fill
from todoku.pdf.form import read_fields
from todoku.pdf.objects import Name, Ref, Stream, text_string
from todoku.pdf.proof import prove
from todoku.pdf.window import build as build_window


def _helv() -> str:
    return "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"


def _form(rotate: int = 0) -> bytes:
    """「Name」の右に空きのある、罫線の表の 1 ページ（印刷用の PDF）。"""
    content = ("BT /F1 12 Tf 60 700 Td (Name) Tj ET\n0.5 w 50 690 300 30 re S\n"
               "BT /F1 12 Tf 60 640 Td (Date) Tj ET\n0.5 w 50 630 300 30 re S\n"
               "q 1 0 0 1 0 0 cm\n")  # 閉じていない q：足す描画は、ページの初めの座標で描く
    blob = build(content, "<< /Font << /F1 2 0 R >> >>", [_helv()])
    if rotate:
        blob = blob.replace(b"/MediaBox [0 0 595 842]", b"/Rotate %d /MediaBox [0 0 595 842]" % rotate)
        blob = _rexref(blob)
    return blob


def _rexref(blob: bytes) -> bytes:
    """手で書き換えた PDF の相互参照を作り直す。"""
    head = blob[:blob.index(b"xref\n0 ")]
    offs = {int(m[1]): m.start() for m in re.finditer(rb"(?m)^(\d+) 0 obj", head)}
    n = max(offs) + 1
    out = bytearray(head)
    x = len(out)
    out += f"xref\n0 {n}\n0000000000 65535 f \n".encode()
    for i in range(1, n):
        out += f"{offs[i]:010d} 00000 n \n".encode()
    root = re.search(rb"/Root (\d+) 0 R", blob)[1].decode()
    out += f"trailer\n<< /Size {n} /Root {root} 0 R >>\nstartxref\n{x}\n%%EOF\n".encode()
    return bytes(out)


def _ids(blob: bytes) -> dict:
    win = build_window(PdfDocument(blob))
    return {f["id"]: f for pg in win["pages"] for f in pg["fields"]}


def _write(blob: bytes, text: str = "H20", after: str = "Name"):
    return fill(PdfDocument(blob), {"@1": {"after": after, "value": text}})


def _proved(blob: bytes, res):
    anchored = {p.id: p.anchor for p in res.patches if p.anchor}
    pr = prove(blob, res.blob, {p.id for p in res.patches}, anchored)
    assert pr.ok, pr.unexpected + pr.lint_errors
    return pr


def test_written_text_is_page_content_and_a_hidden_record():
    blob = _form()
    res = _write(blob)
    assert res.blob.startswith(blob)  # 元のバイトはそのまま
    pr = _proved(blob, res)
    assert pr.drawn == [res.patches[0].id] and "書いた字 1 か所" in pr.headline()
    doc = PdfDocument(res.blob)
    fld = read_fields(doc)[0]
    wd = doc.resolve(fld.widgets[0].ref)
    assert wd["F"] == pagedraw.ANNOT_HIDDEN and fld.value_text() == "H20"
    page = doc.pages()[0]
    data = doc.page_content(page)
    # 元の描画の前に q、後ろで閉じていない q を戻してから描く
    assert data.startswith(b"q\n") and b"Do\nQ\nEMC" in data
    assert data.index(b"(Name) Tj") < data.index(b"BDC")
    key = re.search(rb"/(TdkW\d+) Do", data)[1].decode()
    xo = doc.resolve(doc.resolve(page["Resources"])["XObject"])[key]
    assert xo == wd[pagedraw.KEY] and xo == doc.resolve(wd["AP"])["N"]
    # 元の資源はそのまま残る
    assert "F1" in doc.resolve(doc.resolve(page["Resources"])["Font"])
    assert not any("Apple" in w or "プレビュー" in w for w in res.warnings)  # 表示の違いの知らせは、もう要らない


def test_drawn_value_not_read_as_printed_and_field_is_visible():
    blob = _form()
    res = _write(blob)
    doc = PdfDocument(res.blob)
    assert "H20" not in "".join(g.text for g in page_content(doc, doc.pages()[0]).glyphs)
    f = _ids(res.blob)[res.patches[0].id]
    assert f.get("value") == "H20" and not f.get("hidden") and not f.get("not_printed")
    assert read_fields(doc)[0].visibility() == (None, None)


def test_api_fill_proves_and_reads_back(tmp_path):
    src = tmp_path / "f.pdf"
    src.write_bytes(_form())
    out = tmp_path / "o.pdf"
    r = api.fill(str(src), {"@1": {"after": "Name", "value": "昭和60", "as": "生年"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert out.read_bytes().startswith(src.read_bytes())


def test_rewrite_same_id_replaces_what_is_drawn(tmp_path):
    src = tmp_path / "f.pdf"
    src.write_bytes(_form())
    one = tmp_path / "1.pdf"
    r1 = api.fill(str(src), {"@1": {"after": "Name", "value": "H20", "as": "年"}}, str(one))
    fid = r1["anchored"]["@1"]
    two = tmp_path / "2.pdf"
    r2 = api.fill(str(one), {fid: "H21"}, str(two))
    assert r2["proof_ok"] and r2["read_back_ok"], (r2["unexpected"], r2["read_back_differences"])
    assert two.read_bytes().startswith(one.read_bytes())
    doc = PdfDocument(two.read_bytes())
    fld = read_fields(doc)[0]
    wd = doc.resolve(fld.widgets[0].ref)
    assert fld.value_text() == "H21" and pagedraw.value_problem(doc, wd) is None
    page = doc.pages()[0]
    assert len(re.findall(rb"/TdkW\d+ Do", doc.page_content(page))) == 1  # 重ねて描かない


def test_second_write_on_other_spot_keeps_first(tmp_path):
    src = tmp_path / "f.pdf"
    src.write_bytes(_form())
    one = tmp_path / "1.pdf"
    r1 = api.fill(str(src), {"@1": {"after": "Name", "value": "H20", "as": "年"}}, str(one))
    two = tmp_path / "2.pdf"
    r2 = api.fill(str(one), {"@1": {"after": "Date", "value": "R8", "as": "日"}}, str(two))
    assert r2["proof_ok"] and r2["read_back_ok"], (r2["unexpected"], r2["read_back_differences"])
    doc = PdfDocument(two.read_bytes())
    data = doc.page_content(doc.pages()[0])
    assert len(set(re.findall(rb"/(TdkW\d+) Do", data))) == 2
    assert {f.value_text() for f in read_fields(doc)} == {"H20", "R8"}


def test_checkbox_drawn_and_toggled():
    blob = build("BT /F1 12 Tf 80 700 Td (\\225 Agree) Tj ET\n0.5 w 60 698 10 10 re S\n", "<< /Font << /F1 2 0 R >> >>", [_helv()])
    doc = PdfDocument(blob)
    f = F.Filler(doc)
    view = F.View(doc, doc.pages()[0])
    f.add_widget(1, {"id": "c1", "kind": "checkbox", "rect": [60.0, 134.0, 70.0, 144.0]}, True, view)
    b1 = f.finish()
    d1 = PdfDocument(b1)
    fld = read_fields(d1)[0]
    wd = d1.resolve(fld.widgets[0].ref)
    assert wd["F"] == pagedraw.ANNOT_HIDDEN and pagedraw.value_problem(d1, wd) is None
    assert b"Tj" in d1.decode(d1.resolve(wd[pagedraw.KEY]))
    # 外して書き直すと、ページの印も消える
    r2 = fill(d1, {fld.name: False})
    d2 = PdfDocument(r2.blob)
    wd2 = d2.resolve(read_fields(d2)[0].widgets[0].ref)
    assert wd2["AS"] == Name("Off") and pagedraw.value_problem(d2, wd2) is None
    assert b"Tj" not in d2.decode(d2.resolve(wd2[pagedraw.KEY]))
    pr = prove(b1, r2.blob, {fld.name})
    assert pr.ok, pr.unexpected


@pytest.mark.parametrize("rot", [90, 180, 270])
def test_rotated_page_draws_in_the_rect(rot):
    blob = _form(rot)
    doc0 = PdfDocument(blob)
    f = F.Filler(doc0)
    f.add_widget(1, {"id": "r1", "kind": "text", "rect": [100.0, 100.0, 300.0, 124.0], "size": 12.0}, "H20", F.View(doc0, doc0.pages()[0]))
    f.patches[-1].anchor = None
    out = f.finish()
    assert prove(blob, out, {"r1"}).ok
    doc = PdfDocument(out)
    page = doc.pages()[0]
    fld = read_fields(doc)[0]
    x0, y0, x1, y1 = fld.widgets[0].rect
    from todoku.pdf.content import Interpreter

    it = Interpreter(doc)  # 描画そのものを読む（読み分けなし）：字が、欄の四角の中にある
    pc = it.run_stream(doc.page_content(page), doc.resolve(page["Resources"]))
    gl = [g for g in pc.glyphs if g.text in "H20"]
    assert gl
    for g in gl:
        assert x0 - 1 <= g.ox <= x1 + 1 and y0 - 1 <= g.oy <= y1 + 1


def test_mark_alone_does_not_hide_printed_text():
    """/TdkField の印や、値と違う字を指す欄では、ページの字を隠せない（書類の作り手が読み手から字を隠せないように）。"""
    content = "/Span <</TdkField (x)>> BDC BT /F1 12 Tf 60 700 Td (Secret) Tj ET EMC\n"
    blob = build(content, "<< /Font << /F1 2 0 R >> >>", [_helv()])
    doc = PdfDocument(blob)
    assert "Secret" in "".join(g.text for g in page_content(doc, doc.pages()[0]).glyphs)
    # とどくの書いた形で、値だけを変えた欄：値と描いた字が違うので、描いた字は書類の字として読み、証明で気づく
    res = _write(_form())
    d = PdfDocument(res.blob)
    up = F.Update(d)
    wref = read_fields(d)[0].widgets[0].ref
    wd = up.get(wref)
    wd["V"] = text_string("X99")
    forged = up.build()
    d2 = PdfDocument(forged)
    assert "H" in "".join(g.text for g in page_content(d2, d2.pages()[0]).glyphs)
    w = read_fields(d2)[0].widgets[0]
    assert not w.drawn and not w.on_screen  # 隠した欄として扱う


def test_proof_catches_drawn_text_that_differs_or_is_missing():
    blob = _form()
    res = _write(blob)
    d = PdfDocument(res.blob)
    # ページの描画のストリームを別の命令に：照らして気づく
    pg = d.resolve(d.pages()[0]["_ref"])
    post = pg["Contents"][-1]
    up = F.Update(d)
    tampered = up.objs.setdefault(post.num, (0, Stream({}, b"Q\nQ\n")))
    assert tampered
    bad = up.build()
    pr = prove(blob, bad, {p.id for p in res.patches}, {p.id: p.anchor for p in res.patches if p.anchor})
    assert not pr.ok
    assert any("描" in u for u in pr.unexpected)


def _signed(tmp_path) -> tuple[bytes, list[int], bytes]:
    """本物の署名（承認の署名。自己署名の証明書で openssl が作る CMS）のある PDF と ByteRange。"""
    openssl = shutil.which("openssl")
    key, cert = tmp_path / "k.pem", tmp_path / "c.pem"
    subprocess.run([openssl, "req", "-x509", "-newkey", "rsa:2048", "-nodes", "-keyout", str(key), "-out", str(cert),
                    "-subj", "/CN=todoku test", "-days", "2"], check=True, capture_output=True)
    hole = 12000  # /Contents の 16 進の字数
    sigv = (f"<< /Type /Sig /Filter /Adobe.PPKLite /SubFilter /adbe.pkcs7.detached /ByteRange [0 0000000000 0000000000 0000000000] "
            f"/Contents <{'0' * hole}> >>")
    content = "BT /F1 12 Tf 60 700 Td (Name) Tj ET\n0.5 w 50 690 300 30 re S\n"
    sigf = "<< /Type /Annot /Subtype /Widget /FT /Sig /T (office_sig) /V 3 0 R /Rect [0 0 0 0] /P 1 0 R /F 132 >>"
    blob = bytearray(build(content, "<< /Font << /F1 2 0 R >> >>", [_helv(), sigv, sigf],
                           annots="/Annots [4 0 R]", acroform="/AcroForm << /Fields [4 0 R] /SigFlags 3 >>"))
    i = blob.index(b"/Contents <0000") + len(b"/Contents ")
    j = i + hole + 2
    br = [0, i, j, len(blob) - j]
    blob[blob.index(b"/ByteRange [") + 12:blob.index(b"/ByteRange [") + 12 + 32] = f"{br[0]} {br[1]:010d} {br[2]:010d} {br[3]:010d}".encode()[:34][:32].ljust(32)
    signed_part = bytes(blob[:i]) + bytes(blob[j:])
    data = tmp_path / "d.bin"
    data.write_bytes(signed_part)
    der = tmp_path / "s.der"
    subprocess.run([openssl, "cms", "-sign", "-binary", "-in", str(data), "-signer", str(cert), "-inkey", str(key), "-outform", "DER",
                    "-out", str(der), "-noattr"], check=True, capture_output=True)
    hexsig = der.read_bytes().hex().encode()
    assert len(hexsig) <= hole
    blob[i + 1:i + 1 + len(hexsig)] = hexsig
    return bytes(blob), br, cert.read_bytes()


def _verify(blob: bytes, br: list[int], cert: bytes, tmp_path) -> bool:
    openssl = shutil.which("openssl")
    a, b = blob[br[0]:br[0] + br[1]], blob[br[2]:br[2] + br[3]]
    hexsig = blob[br[1] + 1:br[2] - 1]
    if len(hexsig) % 2:
        hexsig += b"0"
    der = bytes.fromhex(hexsig.decode())
    # 署名の長さは DER の頭から読む（後ろの 0 を削って決めると、00 で終わる署名が欠けて、検証がときどき落ちる）
    n = der[1]
    size = 2 + n if n < 0x80 else 2 + (n & 0x7F) + int.from_bytes(der[2:2 + (n & 0x7F)], "big")
    (tmp_path / "v.der").write_bytes(der[:size])
    (tmp_path / "v.bin").write_bytes(a + b)
    (tmp_path / "v.pem").write_bytes(cert)
    r = subprocess.run([openssl, "cms", "-verify", "-binary", "-inform", "DER", "-in", str(tmp_path / "v.der"), "-content", str(tmp_path / "v.bin"),
                        "-CAfile", str(tmp_path / "v.pem"), "-out", os.devnull], capture_output=True)
    return r.returncode == 0


@pytest.mark.skipif(shutil.which("openssl") is None, reason="openssl が要る")
def test_signed_pdf_stays_a_visible_field_and_signature_verifies(tmp_path):
    blob, br, cert = _signed(tmp_path)
    assert _verify(blob, br, cert, tmp_path)
    src = tmp_path / "signed.pdf"
    src.write_bytes(blob)
    out = tmp_path / "out.pdf"
    r = api.fill(str(src), {"@1": {"after": "Name", "value": "H20", "as": "年"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    new = out.read_bytes()
    assert new.startswith(blob)
    assert _verify(new, br, cert, tmp_path)  # 署名の範囲のバイトは変わらない
    doc = PdfDocument(new)
    w = [f for f in read_fields(doc) if f.kind != "signature"][0].widgets[0]
    wd = doc.resolve(w.ref)
    assert wd["F"] == 4 and pagedraw.KEY not in wd  # ページの描画は変えない：欄として足す
    assert b"TdkW" not in doc.page_content(doc.pages()[0])
    assert any("プレビュー" in x for x in r["warnings"]) and any("電子署名" in x for x in r["warnings"])


def test_permission_modify_denied_keeps_field():
    res_doc = PdfDocument(_form())

    class Dec:
        def restrictions(self):
            return [("create", "欄の追加"), ("modify", "ページの描画の変更")]

    f = F.Filler(res_doc)
    res_doc.decryptor = Dec()
    assert pagedraw.may_draw(f.doc, f.sigs) is False
    res_doc.decryptor = None
    assert pagedraw.may_draw(f.doc, f.sigs) is True


def test_existing_tight_field_fits_redrawing_viewers():
    """元からある 1 行の欄：描き直すビューアーでも欠けないよう、入らない字は小さくし、/DA も同じ大きさにする（自動の大きさは決まった大きさに）。"""
    for da, expect_auto in ((b"/Helv 10 Tf 0 g", False), (b"/Helv 0 Tf 0 g", True)):
        wid = ("<< /Type /Annot /Subtype /Widget /FT /Tx /T (tight) /F 4 /Rect [100 700 121 716] /DA (%s) /P 1 0 R >>" % da.decode())
        blob = build("", "<< >>", [_helv(), wid], annots="/Annots [3 0 R]",
                     acroform="/AcroForm << /Fields [3 0 R] /DR << /Font << /Helv 2 0 R >> >> /DA (/Helv 0 Tf 0 g) >>")
        res = fill(PdfDocument(blob), {"tight": "H20"})
        assert prove(blob, res.blob, {"tight"}).ok
        doc = PdfDocument(res.blob)
        w = doc.resolve(read_fields(doc)[0].widgets[0].ref)
        assert w["F"] == 4  # 作り手の欄は、見える欄のまま（画面で直せる）
        size = F._da_size(bytes(doc.resolve(w.get("DA"))).decode("latin-1"))
        assert size is not None and F.pdfkit_needs("H20", size) <= 21.0 + 1e-6
        ap = doc.resolve(doc.resolve(w["AP"])["N"])
        assert f"{size:.2f} Tf".encode() in doc.decode(ap)


def test_existing_roomy_auto_field_stays_auto():
    wid = "<< /Type /Annot /Subtype /Widget /FT /Tx /T (wide) /F 4 /Rect [100 700 400 716] /DA (/Helv 0 Tf 0 g) /P 1 0 R >>"
    blob = build("", "<< >>", [_helv(), wid], annots="/Annots [3 0 R]",
                 acroform="/AcroForm << /Fields [3 0 R] /DR << /Font << /Helv 2 0 R >> >> /DA (/Helv 0 Tf 0 g) >>")
    res = fill(PdfDocument(blob), {"wide": "H20"})
    doc = PdfDocument(res.blob)
    w = doc.resolve(read_fields(doc)[0].widgets[0].ref)
    da = w.get("DA")
    assert da is None or b" 0 Tf" in bytes(doc.resolve(da))


def test_fit_redrawn_floor_keeps_size_when_it_cannot_fit():
    s, pinned = F._fit_redrawn("東京都千代田区一番町", 10.0, 20.0, 16.0, False)
    assert s == 10.0 and pinned is False  # 読める大きさまで小さくしても入らない：そのまま（出す前の知らせは別）


@pytest.mark.skipif(shutil.which("openssl") is None, reason="openssl が要る")
def test_even_page_content_on_signed_pdf_leaves_signature_bytes_verifying(tmp_path, monkeypatch):
    """（方針の材料）署名のある PDF にページの描画を足しても、増分なので openssl の検証は通る。とどくが欄のままにするのは、ビューアーの扱いのため。"""
    monkeypatch.setattr(pagedraw, "may_draw", lambda d, s=None: True)
    blob, br, cert = _signed(tmp_path)
    src = tmp_path / "signed.pdf"
    src.write_bytes(blob)
    out = tmp_path / "out.pdf"
    r = api.fill(str(src), {"@1": {"after": "Name", "value": "H20", "as": "年"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    assert _verify(out.read_bytes(), br, cert, tmp_path)
    assert b"TdkW" in PdfDocument(out.read_bytes()).page_content(PdfDocument(out.read_bytes()).pages()[0])


def test_encrypted_pdf_draws_with_the_same_key(tmp_path):
    from test_round5_pdf_fixes import _encrypted

    src = tmp_path / "e.pdf"
    src.write_bytes(_encrypted(-4))
    out = tmp_path / "o.pdf"
    r = api.fill(str(src), {"@1": {"page": 1, "box": [30, 100, 90, 110], "value": "H20", "as": "年"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    doc = PdfDocument(out.read_bytes())
    data = doc.page_content(doc.pages()[0])
    assert re.search(rb"/TdkW\d+ Do", data)  # 開いた後は、足した描画が読める（暗号はファイルの中）
    assert re.search(rb"/TdkW\d+ Do", out.read_bytes()) is None


def test_hidden_record_is_not_drawn_again_by_view_page():
    from todoku.pdf.render.paint import annotation_forms

    doc = PdfDocument(_write(_form()).blob)
    assert list(annotation_forms(doc, doc.pages()[0])) == []  # 字はページの描画の分だけ（見た目を重ねて描かない）


def test_existing_field_too_tight_to_fit_tells_the_viewer_difference():
    wid = "<< /Type /Annot /Subtype /Widget /FT /Tx /T (t) /F 4 /Rect [100 700 118 716] /DA (/Helv 10 Tf 0 g) /P 1 0 R >>"
    blob = build("", "<< >>", [_helv(), wid], annots="/Annots [3 0 R]",
                 acroform="/AcroForm << /Fields [3 0 R] /DR << /Font << /Helv 2 0 R >> >> /DA (/Helv 0 Tf 0 g) >>")
    res = fill(PdfDocument(blob), {"t": "昭和60"})
    assert any("プレビュー" in w for w in res.warnings)
