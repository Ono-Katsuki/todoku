"""本人の印影・署名の画像を置く（本人が同意したときだけ）：PNG・JPEG を開く、PDF に置く・証明・読み戻し・ページの画像。

画像は、テストの中で作った図形（輪と四角）。実在の人の印影は使わない。
"""

from __future__ import annotations

import base64
import struct
import zlib
from pathlib import Path

import pytest

from todoku import api, sealimage
from todoku.pdf.doc import PdfDocument
from todoku.pdf.fill import FillError, fill
from todoku.pdf.pageimage import render
from todoku.pdf.proof import prove

from test_pdf import jp_font_objs, jp_text, make_pdf
from test_scan import _scan_pdf

FIX = Path(__file__).parent / "fixtures"


# ---------------------------------------------------------------------- 図形の画像を作る
def png_bytes(w: int, h: int, ctype: int, rows: list[bytes], depth: int = 8, plte: bytes = b"", trns: bytes | None = None,
              interlace: int = 0, filt: int = 0) -> bytes:
    def chunk(t: bytes, b: bytes) -> bytes:
        return struct.pack(">I", len(b)) + t + b + struct.pack(">I", zlib.crc32(t + b) & 0xFFFFFFFF)

    if filt == 1:  # Sub のフィルターで縮める（戻す処理を確かめる）
        bpp = max(1, {0: 1, 2: 3, 3: 1, 4: 2, 6: 4}[ctype] * depth // 8)
        raw = b"".join(b"\x01" + bytes((r[i] - (r[i - bpp] if i >= bpp else 0)) & 0xFF for i in range(len(r))) for r in rows)
    elif filt == 2:  # Up
        prev = bytes(len(rows[0]))
        parts = []
        for r in rows:
            parts.append(b"\x02" + bytes((a - b) & 0xFF for a, b in zip(r, prev)))
            prev = r
        raw = b"".join(parts)
    else:
        raw = b"".join(b"\0" + r for r in rows)
    out = b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, depth, ctype, 0, 0, interlace))
    if plte:
        out += chunk(b"PLTE", plte)
    if trns is not None:
        out += chunk(b"tRNS", trns)
    return out + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def ring_rgba(n: int = 40) -> tuple[bytes, bytes, bytes]:
    """赤い輪（まわりは透ける）。(PNG, 開いた RGB, アルファ)。"""
    rows, rgb, alpha = [], bytearray(), bytearray()
    c = (n - 1) / 2
    for y in range(n):
        r = bytearray()
        for x in range(n):
            d = ((x - c) ** 2 + (y - c) ** 2) ** 0.5
            on = n * 0.36 <= d <= n * 0.48
            px = (200, 20, 20, 255) if on else (0, 0, 0, 0)
            r += bytes(px)
            rgb += bytes(px[:3])
            alpha.append(px[3])
        rows.append(bytes(r))
    return png_bytes(n, n, 6, rows, filt=1), bytes(rgb), bytes(alpha)


def seal_page() -> bytes:
    """「氏名　山田　　印」の行と、その下の「住所」の行がある、字のページ。"""
    font, extra, codes = jp_font_objs("氏名山田印住所", 5)
    c = (b"BT /F1 12 Tf 60 700 Td " + jp_text(codes, "氏名") + b" Tj 60 0 Td " + jp_text(codes, "山田") + b" Tj "
         b"140 0 Td " + jp_text(codes, "印") + b" Tj ET\n"
         b"BT /F1 12 Tf 60 600 Td " + jp_text(codes, "住所") + b" Tj ET\n")
    return make_pdf(c, {"F1": font}, extra)


# ---------------------------------------------------------------------- 画像を開く
def test_png_decodes_rgba_palette_gray_and_filters():
    data, rgb, alpha = ring_rgba(20)
    im = sealimage.decode(data)
    assert (im.kind, im.w, im.h, im.channels) == ("png", 20, 20, 3)
    assert im.samples == rgb and im.alpha == alpha and not im.opaque
    # パレット（4 ビット）と tRNS：番号 0 は透ける
    plte = bytes([255, 255, 255, 180, 0, 0])
    rows = [bytes([0x10, 0x10]), bytes([0x11, 0x00])]  # 4 画素×2 行：1,0,1,0 ／ 1,1,0,0
    im = sealimage.decode(png_bytes(4, 2, 3, rows, depth=4, plte=plte, trns=b"\x00"))
    assert im.samples == bytes([180, 0, 0, 255, 255, 255] * 2 + [180, 0, 0] * 2 + [255, 255, 255] * 2)
    assert im.alpha == bytes([255, 0, 255, 0, 255, 255, 0, 0])
    # 濃さ（Up のフィルター）：透けないので alpha は None
    rows = [bytes([0, 128, 255]), bytes([10, 20, 30])]
    im = sealimage.decode(png_bytes(3, 2, 0, rows, filt=2))
    assert im.samples == b"".join(rows) and im.alpha is None and im.opaque and im.channels == 1


def test_png_interlaced_and_broken_are_refused():
    rows = [bytes([0, 0, 0])]
    with pytest.raises(sealimage.SealError, match="インターレース"):
        sealimage.decode(png_bytes(1, 1, 2, rows, interlace=1))
    bad = bytearray(png_bytes(1, 1, 2, rows))
    bad[-20] ^= 0xFF  # IDAT の中を壊す → CRC が合わない
    with pytest.raises(sealimage.SealError):
        sealimage.decode(bytes(bad))
    with pytest.raises(sealimage.SealError, match="PNG か JPEG"):
        sealimage.decode(b"GIF89a....")


def test_jpeg_size_from_sof_and_progressive_refused():
    data = (FIX / "scan_box.jpg").read_bytes()
    im = sealimage.decode(data)
    assert (im.kind, im.w, im.h) == ("jpeg", 160, 96) and im.opaque and im.data == data
    i = data.index(b"\xff\xc0")
    prog = data[:i] + b"\xff\xc2" + data[i + 2:]
    with pytest.raises(sealimage.SealError, match="プログレッシブ"):
        sealimage.decode(prog)


# ---------------------------------------------------------------------- PDF に置く
def test_consent_is_required(tmp_path):
    png = tmp_path / "ring.png"
    png.write_bytes(ring_rgba()[0])
    blob = seal_page()
    for spec in ({"seal": str(png), "cover": "印"}, {"seal": str(png), "cover": "印", "consent": "yes"}):
        with pytest.raises(FillError, match="本人の同意"):
            fill(PdfDocument(blob, "s.pdf"), {"@1": spec})
    src = tmp_path / "s.pdf"
    src.write_bytes(blob)
    with pytest.raises(api.ApiError, match="consent"):
        api.fill(str(src), {"@1": {"signature": str(png), "after": "印"}}, str(tmp_path / "o.pdf"))
    assert not (tmp_path / "o.pdf").exists()


def test_seal_over_the_mark_on_a_text_page(tmp_path):
    data, rgb, alpha = ring_rgba()
    png = tmp_path / "ring.png"
    png.write_bytes(data)
    src, out = tmp_path / "s.pdf", tmp_path / "o.pdf"
    src.write_bytes(seal_page())
    r = api.fill(str(src), {"@1": {"seal": str(png), "cover": "印", "size": 12, "consent": True}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert r["anchored"] == {"@1": "p1.s1"}
    assert "置いた画像" in r["proof"]
    say = r["say_to_user"]
    assert "1ページの「印」の上に、本人の印影の画像を置きました（横 12 mm・縦 12 mm）" in say, say
    assert "実印" in say and "自署" in say and "様式の注意書き" in say
    assert "記入した内容（0 か所" not in say
    # 元のバイト列は先頭にそのまま。足した画像は、本人の画像の画素とアルファ
    orig, new = src.read_bytes(), out.read_bytes()
    assert new[:len(orig)] == orig
    doc = PdfDocument(new)
    annots = [doc.resolve(a) for a in doc.resolve(doc.pages()[0].get("Annots"))]
    st = [a for a in annots if a["Subtype"].name == "Stamp"]
    assert len(st) == 1
    ap = doc.resolve(st[0]["AP"]["N"])
    im = doc.resolve(ap.dict["Resources"]["XObject"]["TdkSeal"])
    assert doc.decode(im) == rgb and doc.decode(doc.resolve(im.dict["SMask"])) == alpha
    # 窓（読み直し）にも、スタンプの注釈として出る
    win = api.window(str(out))[1]
    assert any(a["kind"] == "Stamp" and "印影" in a["text"] for a in win["pages"][0]["annotations"])
    # ページの画像にも、置いた画像が描かれる（輪の上は赤い。輪の中は透けて、下の字「印」が見える）
    rr = render(doc, 1, max_side=600)
    box = r["placed"][0]["box_mm"]
    s = rr["scale"]
    cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2

    def px(x_mm, y_mm):
        i = (int(y_mm * s) * rr["w"] + int(x_mm * s)) * 3
        return tuple(rr["rgb"][i:i + 3])

    ring = px(cx + 12 * 0.42, cy)
    assert ring[0] > 150 and ring[1] < 80 and ring[2] < 80, ring
    assert px(cx + 12 * 0.3, cy) == (255, 255, 255)  # 輪と字の間は白い（透ける）
    v = api.view_page(str(out), 1)
    assert "p1.s1" in v["text"]


def test_seal_right_of_label_and_opaque_jpeg_does_not_hide_text(tmp_path):
    jpg = tmp_path / "box.jpg"
    jpg.write_bytes((FIX / "scan_box.jpg").read_bytes())
    src = tmp_path / "s.pdf"
    src.write_bytes(seal_page())
    out = tmp_path / "o.pdf"
    r = api.fill(str(src), {"@1": {"seal": str(jpg), "after": "山田", "consent": True},
                            "@2": {"signature": str(jpg), "after": "住所", "consent": True}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    assert "1ページの「山田」の右に、本人の印影の画像を置きました" in r["say_to_user"]
    assert "1ページの「住所」の右に、本人の署名の画像を置きました" in r["say_to_user"]
    doc = PdfDocument(out.read_bytes())
    ims = [o for n in doc.xref if isinstance(o := doc.get(n), object) and getattr(getattr(o, "dict", {}).get("Subtype"), "name", "") == "Image"]
    assert ims and all(doc.decode(o) == jpg.read_bytes() for o in ims)  # JPEG は、ファイルのバイト列のまま
    # 透けない JPEG で「山田」を隠す置き方は断る
    with pytest.raises(api.ApiError, match="隠れます"):
        api.fill(str(src), {"@1": {"seal": str(jpg), "page": 1, "box": [15, 40, 60, 55], "consent": True}}, str(tmp_path / "x.pdf"))


def test_seal_in_a_box_on_a_scanned_page(tmp_path):
    data, _rgb, _a = ring_rgba()
    png = tmp_path / "ring.png"
    png.write_bytes(data)
    src, out = tmp_path / "scan.pdf", tmp_path / "o.pdf"
    src.write_bytes(_scan_pdf(90))
    r = api.fill(str(src), {"@1": {"seal": str(png), "page": 1, "box": [10, 10, 30, 40], "consent": True}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
    pl = r["placed"][0]
    assert pl["id"] == "p1.s1" and pl["page"] == 1
    # 縦横の比を保って、枠の中央に：20×20 mm
    assert [round(v) for v in pl["box_mm"]] == [10, 15, 30, 35]
    assert "1ページの枠（左 10・上 10・右 30・下 40 mm）の中に、本人の印影の画像を置きました" in r["say_to_user"]
    v = api.view_page(str(out), 1)
    assert base64.b64decode(v["png_base64"])[:4] == b"\x89PNG" and "p1.s1" in v["text"]


def test_proof_catches_tampered_image_and_unclaimed_stamps(tmp_path):
    data, _rgb, _a = ring_rgba()
    png = tmp_path / "ring.png"
    png.write_bytes(data)
    blob = seal_page()
    spec = {"seal": str(png), "cover": "印", "size": 12, "consent": True}
    res = fill(PdfDocument(blob, "s.pdf"), {"@1": spec})
    assert [(p.id, p.kind) for p in res.patches] == [("p1.s1", "seal")]
    assert prove(blob, res.blob, {"p1.s1"}, {"p1.s1": spec}).ok
    # 言っていない画像：置いた所を伝えなければ、足したスタンプと画像は想定外
    bare = prove(blob, res.blob, {"p1.s1"}, {})
    assert not bare.ok and any("スタンプの注釈か画像" in u for u in bare.unexpected)
    # 本人の画像と違う画像（1画素だけ違う）で証明すると、画素が一致しない
    other, _x, _y = ring_rgba()
    rows = []
    n = 40
    rgba = bytearray(sealimage.decode(other).samples)
    al = sealimage.decode(other).alpha
    rgba[0] ^= 1
    for y in range(n):
        r = bytearray()
        for x in range(n):
            i = y * n + x
            r += rgba[i * 3:i * 3 + 3] + bytes([al[i]])
        rows.append(bytes(r))
    alt = tmp_path / "alt.png"
    alt.write_bytes(png_bytes(n, n, 6, rows))
    bad = prove(blob, res.blob, {"p1.s1"}, {"p1.s1": {**spec, "seal": str(alt)}})
    assert not bad.ok and any("画素" in u for u in bad.unexpected)
    # 埋めた JPEG のバイト列を書き換えると、証明が見つける（同じ長さで1バイトだけ）
    jpg = tmp_path / "box.jpg"
    jpg.write_bytes((FIX / "scan_box.jpg").read_bytes())
    jspec = {"seal": str(jpg), "after": "山田", "consent": True}
    res2 = fill(PdfDocument(blob, "s.pdf"), {"@1": jspec})
    assert prove(blob, res2.blob, {"p1.s1"}, {"p1.s1": jspec}).ok
    jb = jpg.read_bytes()
    k = res2.blob.index(jb, len(blob)) + len(jb) // 2
    tampered = res2.blob[:k] + bytes([res2.blob[k] ^ 0x55]) + res2.blob[k + 1:]
    t = prove(blob, tampered, {"p1.s1"}, {"p1.s1": jspec})
    assert not t.ok and any("バイト列" in u for u in t.unexpected)
    # 同意の印のない置き方の主張は、証明でも通らない
    nc = prove(blob, res.blob, {"p1.s1"}, {"p1.s1": {k_: v for k_, v in spec.items() if k_ != "consent"}})
    assert not nc.ok


def test_word_and_excel_refuse_until_supported(tmp_path):
    # Word・Excel には、証明つきで画像を置く方法をまだ作っていない：書かずに、はっきり断る（ファイルも書き出さない）
    from test_docx import make_docx, p

    png = tmp_path / "ring.png"
    png.write_bytes(ring_rgba()[0])
    docx = tmp_path / "a.docx"
    docx.write_bytes(make_docx(p("氏名　　　　印")))
    xlsx = tmp_path / "a.xlsx"
    xlsx.write_bytes((FIX / "kami_meisai.xlsx").read_bytes())
    for src, spec in ((docx, {"seal": str(png), "after": "印", "consent": True}),
                      (xlsx, {"seal": str(png), "cell": "A1", "consent": True})):
        out = tmp_path / ("o" + src.suffix)
        with pytest.raises(api.ApiError, match="いまは PDF の書類だけ"):
            api.fill(str(src), {"@1": spec}, str(out))
        assert not out.exists()


def test_seal_with_text_and_a_second_fill(tmp_path):
    # 字の書き込みと一緒に置き、置いた後のファイルにもう一度書く：番号は重ならず、前に置いた画像は変えない
    data, _rgb, _a = ring_rgba()
    png = tmp_path / "ring.png"
    png.write_bytes(data)
    src, mid, out = tmp_path / "s.pdf", tmp_path / "m.pdf", tmp_path / "o.pdf"
    src.write_bytes(seal_page())
    r = api.fill(str(src), {"@1": {"after": "住所", "value": "横浜市"},
                            "@2": {"seal": str(png), "cover": "印", "size": 10, "consent": True}}, str(mid))
    assert r["proof_ok"] and r["read_back_ok"] and r["anchored"] == {"@1": "p1.w1", "@2": "p1.s2"}, r
    assert "住所（の右）：横浜市" in r["say_to_user"] and "「印」の上に、本人の印影の画像" in r["say_to_user"]
    r2 = api.fill(str(mid), {"@1": {"signature": str(png), "page": 1, "box": [150, 20, 170, 40], "consent": True}}, str(out))
    assert r2["proof_ok"] and r2["read_back_ok"] and r2["anchored"] == {"@1": "p1.s3"}, r2
    v = api.view_page(str(out), 1)
    assert "p1.s2" in v["text"] and "p1.s3" in v["text"]


@pytest.mark.parametrize("case", ["R3-RC4-128", "R4-AES128", "R6-AES256"])
def test_seal_on_an_encrypted_pdf(case, tmp_path):
    # パスワードで保護された PDF：画像も、元と同じ鍵で暗号化して足す。証明は復号した画像で、本人の画像と照らす
    from test_pdf_password import USER, encrypted

    blob = encrypted(case)
    data, rgb, _alpha = ring_rgba()
    png = tmp_path / "ring.png"
    png.write_bytes(data)
    spec = {"seal": str(png), "after": "Name", "consent": True}
    res = fill(PdfDocument(blob, "e.pdf", password=USER), {"@1": spec})
    assert prove(blob, res.blob, {"p1.s1"}, {"p1.s1": spec}, password=USER).ok
    assert zlib.compress(rgb, 9) not in res.blob[len(blob):]  # 平文の画像は出てこない
    src, out = tmp_path / "e.pdf", tmp_path / "o.pdf"
    src.write_bytes(blob)
    r = api.fill(str(src), {"@1": spec}, str(out), password=USER)
    assert r["proof_ok"] and r["read_back_ok"], (r["unexpected"], r["read_back_differences"])
