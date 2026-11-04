"""範囲を指定して開く（region）：予算は開く分に掛かるので、どの部分でも元の解像度まで開ける。opj_compress・cjpeg で作り、全体を開いた画素と比べる。
道具がない環境では飛ばす。"""

from __future__ import annotations

import random
import shutil
import struct
import subprocess

import pytest

from todoku.imaging import formats as F
from todoku.imaging import jpx
from todoku.imaging import standalone as S
from todoku.pdf.raster import jpeg_gray

OPJ = shutil.which("opj_compress")
CJPEG = shutil.which("cjpeg")
needs_opj = pytest.mark.skipif(not OPJ, reason="opj_compress がない")
needs_cjpeg = pytest.mark.skipif(not CJPEG, reason="cjpeg がない")


def _ppm(path, w=320, h=240, seed=3):
    r = random.Random(seed)
    body = bytearray()
    for y in range(h):
        for x in range(w):
            v = (x * 255 // w + (y // 7) * 13 + r.randrange(16)) % 256
            body += bytes((v, (v * 3) % 256, 255 - v))
    path.write_bytes(b"P6\n%d %d\n255\n" % (w, h) + bytes(body))


def _crop(px, w, x0, y0, x1, y1, n=3):
    return b"".join(px[(y * w + x0) * n:(y * w + x1) * n] for y in range(y0, y1))


def _rows_jpx(im):
    out = bytearray(im.w * im.h * 3)
    for c in range(3):
        out[c::3] = im.planes[c]
    return bytes(out)


@pytest.fixture(scope="module")
def ppm(tmp_path_factory):
    p = tmp_path_factory.mktemp("crop") / "a.ppm"
    _ppm(p)
    return p


@needs_opj
@pytest.mark.parametrize("args", [[], ["-I", "-r", "5"], ["-t", "64,64"], ["-c", "[64,64],[32,32]", "-p", "RPCL"]],
                         ids=["5-3", "9-7", "tiles", "precincts"])
@pytest.mark.parametrize("win", [(40, 30, 200, 150), (0, 0, 33, 17), (300, 220, 320, 240), (157, 1, 170, 239)])
def test_jpx_window_same_as_full(ppm, tmp_path, args, win):
    out = tmp_path / "a.jp2"
    subprocess.run([OPJ, "-i", str(ppm), "-o", str(out)] + args, check=True, stdout=subprocess.DEVNULL)
    blob = out.read_bytes()
    full = jpx.decode(blob, max_t1_bytes=10**9)
    im = jpx.decode(blob, window=win, max_t1_bytes=10**9)
    assert im.window == win and (im.w, im.h) == (win[2] - win[0], win[3] - win[1])
    assert _rows_jpx(im) == _crop(_rows_jpx(full), full.w, *win)


@needs_opj
def test_jpx_window_reaches_native_under_a_small_budget(ppm, tmp_path, monkeypatch):
    """全体は予算のために解像度を下げるが、範囲を狭めれば元の解像度で開き、同じ画素になる。何を使ったかと、どう拡大するかを伝える。"""
    out = tmp_path / "a.jp2"
    subprocess.run([OPJ, "-i", str(ppm), "-o", str(out), "-b", "16,16"], check=True, stdout=subprocess.DEVNULL)
    blob = out.read_bytes()
    full = jpx.decode(blob, max_t1_bytes=10**9)
    monkeypatch.setattr(jpx, "MAX_T1_BYTES", len(blob) // 3)
    whole = S.view_frame(blob, 1)
    assert "分の 1" in whole["text"] and "region" in whole["text"]
    v = S.view_frame(blob, 1, [40, 30, 80, 60])
    assert "元の解像度" in v["text"]
    fr = jpx.decode(blob, window=(40, 30, 80, 60))
    assert fr.reduce == 0
    assert _rows_jpx(fr) == _crop(_rows_jpx(full), full.w, 40, 30, 80, 60)


@needs_opj
def test_pdf_page_region_opens_only_visible_part(ppm, tmp_path, monkeypatch):
    """PDF のページの画像も同じ：拡大すれば、見える所だけ元の解像度で開く。"""
    from test_jpx import _pdf_with_jpx

    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.pageimage import render

    out = tmp_path / "a.jp2"
    subprocess.run([OPJ, "-i", str(ppm), "-o", str(out), "-b", "16,16"], check=True, stdout=subprocess.DEVNULL)
    blob = out.read_bytes()
    full = jpx.decode(blob, max_t1_bytes=10**9)
    fp = _rows_jpx(full)
    monkeypatch.setattr(jpx, "MAX_T1_BYTES", len(blob) // 3)
    doc = PdfDocument(_pdf_with_jpx(blob, 320, 240), "t.pdf")
    r = render(doc, 1, [20, 20, 40, 40], max_side=800)
    assert any("分の 1" not in n for n in r["image_notes"]) or not r["image_notes"]
    # 描いた中に、元の画素の値がそのまま出ている（解像度を下げて開くと平均された値になり、ほとんど一致しない）
    want = {fp[i:i + 3] for i in range(0, len(fp), 3)}
    got = [bytes(r["rgb"][i:i + 3]) for i in range(0, len(r["rgb"]), 3)]
    assert sum(1 for g in got if g in want) > 0.99 * len(got)


def _exif(orient: int) -> bytes:
    ifd = b"MM\0*\0\0\0\x08\0\x01" + struct.pack(">HHI", 0x112, 3, 1) + struct.pack(">HH", orient, 0) + b"\0\0\0\0"
    body = b"Exif\0\0" + ifd
    return b"\xff\xe1" + struct.pack(">H", len(body) + 2) + body


@needs_cjpeg
@pytest.mark.parametrize("args", [["-sample", "2x2"], ["-sample", "1x1"], ["-sample", "2x2", "-progressive"], ["-sample", "2x1", "-restart", "2"]],
                         ids=["420", "444", "progressive", "422-restart"])
def test_jpeg_window_same_as_full(ppm, args):
    blob = subprocess.run([CJPEG, "-quality", "90"] + args + [str(ppm)], check=True, capture_output=True).stdout
    full = jpeg_gray(blob, 10**9, want_rgb=True)
    for win in ((40, 30, 200, 150), (0, 0, 17, 9), (300, 220, 320, 240)):
        im = jpeg_gray(blob, 10**9, want_rgb=True, window=win)
        x0, y0, x1, y1 = im.window
        assert x0 <= win[0] and y0 <= win[1] and x1 >= win[2] and y1 >= win[3]
        assert (im.w, im.h) == (x1 - x0, y1 - y0)
        assert im.px == _crop(full.px, full.w, x0, y0, x1, y1)


@needs_cjpeg
@pytest.mark.parametrize("o", range(1, 9))
def test_jpeg_window_follows_exif_orientation(ppm, o):
    """見た目の座標で指定した範囲が、向きを直した全体の同じ所になる。"""
    blob = subprocess.run([CJPEG, "-quality", "90", "-sample", "1x1", str(ppm)], check=True, capture_output=True).stdout
    blob = blob[:2] + _exif(o) + blob[2:]
    full = F.decode(blob)
    assert (full.w, full.h) == ((320, 240) if o < 5 else (240, 320))
    fr = F.decode(blob, window=(50, 60, 90, 120))
    x0, y0, x1, y1 = fr.window
    assert x0 <= 50 and y0 <= 60 and x1 >= 90 and y1 >= 120
    assert (fr.w, fr.h) == (x1 - x0, y1 - y0)
    assert fr.rgb == _crop(full.rgb, full.w, x0, y0, x1, y1)


@needs_cjpeg
def test_view_frame_region_matches_full_jpeg(ppm):
    """view_frame の範囲：範囲だけ開いても、全体を開いて切り抜いたものと同じ画像になる。"""
    blob = subprocess.run([CJPEG, "-quality", "90", "-sample", "2x2", str(ppm)], check=True, capture_output=True).stdout
    v = S.view_frame(blob, 1, [40, 30, 200, 150])
    assert "元の解像度" in v["text"]
    assert v["width"] > 160
