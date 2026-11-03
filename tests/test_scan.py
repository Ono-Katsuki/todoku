"""スキャンした紙（画像だけのページ）：画像を開く（CCITT・JPEG・Flate）・ページの画像（目盛りと枠）・枠（mm）で書く。"""

from __future__ import annotations

import base64
import zlib
from pathlib import Path

from todoku import api
from todoku.pdf import raster
from todoku.pdf.doc import PdfDocument
from todoku.pdf.pageimage import render

FIX = Path(__file__).parent / "fixtures"


def _bits(s: str) -> bytes:
    s = s + "0" * (-len(s) % 8)
    return bytes(int(s[i:i + 8], 2) for i in range(0, len(s), 8))


def test_ccitt_g4_by_the_code_tables():
    # 16 画素の幅で2行：1行目は白だけ（V0）、2行目は 白4・黒8・白4（水平モード 001＋白4「1011」＋黒8「000101」、V0）、EOFB
    data = _bits("1" + "001" + "1011" + "000101" + "1" + "000000000001" * 2)
    g = raster.ccitt(data, {"K": -1, "Columns": 16, "Rows": 2})
    assert g.row(0) == b"\xff" * 16
    assert g.row(1) == b"\xff" * 4 + b"\x00" * 8 + b"\xff" * 4
    inv = raster.ccitt(data, {"K": -1, "Columns": 16, "Rows": 2, "BlackIs1": True})
    assert inv.row(1) == b"\x00" * 4 + b"\xff" * 8 + b"\x00" * 4


def test_ccitt_g3_one_dimensional():
    # K=0：EOL ＋ 白4「1011」黒8「000101」白4「1011」
    data = _bits("000000000001" + "1011" + "000101" + "1011")
    g = raster.ccitt(data, {"K": 0, "Columns": 16, "Rows": 1})
    assert g.row(0) == b"\xff" * 4 + b"\x00" * 8 + b"\xff" * 4


def test_jpeg_baseline_and_reduced_scale():
    data = (FIX / "scan_box.jpg").read_bytes()  # 160×96 の白に、黒い枠（x 20〜139・y 20〜76）
    g = raster.jpeg_gray(data)
    assert (g.w, g.h) == (160, 96)
    assert g.px[20 * g.w + 80] < 80 and g.px[50 * g.w + 80] > 200 and g.px[50 * g.w + 20] < 80
    half = raster.jpeg_gray(data, max_pixels=160 * 96 // 4)
    assert (half.w, half.h) == (80, 48) and half.px[25 * 80 + 40] > 200


def _box_bitmap() -> bytes:
    W, H = 160, 96
    px = bytearray(b"\xff" * W * H)
    for x in range(20, 140):
        for y in (20, 21, 75, 76):
            px[y * W + x] = 0
    for y in range(20, 77):
        for x in (20, 21, 138, 139):
            px[y * W + x] = 0
    return bytes(px)


def _scan_pdf(rotate: int = 0) -> bytes:
    """1ページの、画像だけの PDF（160×96 の Flate の画像を、200×120 pt のページいっぱいに置く）。"""
    img = zlib.compress(_box_bitmap())
    content = b"q 200 0 0 120 0 0 cm /Im1 Do Q"
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 120] /Rotate " + str(rotate).encode()
            + b" /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>",
            b"<< /Length " + str(len(content)).encode() + b" >>\nstream\n" + content + b"\nendstream",
            b"<< /Type /XObject /Subtype /Image /Width 160 /Height 96 /ColorSpace /DeviceGray /BitsPerComponent 8"
            b" /Filter /FlateDecode /Length " + str(len(img)).encode() + b" >>\nstream\n" + img + b"\nendstream"]
    out = bytearray(b"%PDF-1.7\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += f"{i} 0 obj\n".encode() + o + b"\nendobj\n"
    x = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    for o in offs:
        out += f"{o:010d} 00000 n \n".encode()
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{x}\n%%EOF\n".encode()
    return bytes(out)


def test_page_image_follows_the_page_and_rotation():
    for rot in (0, 90):
        doc = PdfDocument(_scan_pdf(rot))
        r = render(doc, 1, max_side=400)
        W, H, s = r["w"], r["h"], r["scale"]
        assert (W > H) == (rot == 0)
        # 枠の上の線：画像の y 20/96 → ページの上から 25 pt（＝8.8 mm）。回すと、左から
        def px(x_mm, y_mm):
            return r["rgb"][(int(y_mm * s) * W + int(x_mm * s)) * 3]
        top_mm = 120 * 20.5 / 96 * 25.4 / 72
        mid_mm = 200 * 80 / 160 * 25.4 / 72
        if rot == 0:
            assert px(mid_mm, top_mm) < 80 and px(mid_mm, top_mm + 5) > 200
        else:
            assert px(120 * 25.4 / 72 - top_mm, mid_mm) < 80


def test_write_into_a_box_on_a_scanned_page(tmp_path):
    src, out = tmp_path / "scan.pdf", tmp_path / "out.pdf"
    src.write_bytes(_scan_pdf())
    g = api.grade(str(src))
    assert g["grade"] == "E" and g["write"]["grade"] == "B", g["headline"]
    au = api.audit(str(src))
    img = next(c for c in au["checks"] if c["name"] == "画像だけのページ")
    assert img["ok"] and img["agent"]
    v = api.view_page(str(src), 1, marks=[{"box": [12, 12, 60, 30], "label": "1"}])
    assert base64.b64decode(v["png_base64"])[:8] == b"\x89PNG\r\n\x1a\n" and "mm" in v["text"]
    r = api.fill(str(src), {"@1": {"page": 1, "box": [12, 12, 60, 30], "value": "山田"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"] and r["anchored"] == {"@1": "p1.w1"}
    after = api.view_page(str(out), 1)
    assert "p1.w1「山田」" in after["text"]
    # ページの外・逆さの枠は断る
    import pytest

    with pytest.raises(api.ApiError):
        api.fill(str(src), {"@1": {"page": 1, "box": [60, 12, 12, 30], "value": "x"}}, str(tmp_path / "x.pdf"))
    with pytest.raises(api.ApiError):
        api.fill(str(src), {"@1": {"page": 1, "box": [0, 0, 500, 30], "value": "x"}}, str(tmp_path / "y.pdf"))
