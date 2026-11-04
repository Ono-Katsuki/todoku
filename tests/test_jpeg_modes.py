"""JPEG（T.81）の算術符号・12 ビット・可逆と、4:2:0 のなめらかな広げ方：cjpeg（libjpeg-turbo）で作り、djpeg の画素と比べる。道具がなければ飛ばす。"""

from __future__ import annotations

import random
import shutil
import subprocess

import pytest

from todoku.imaging import standalone as S
from todoku.pdf.raster import jpeg_gray

CJPEG = shutil.which("cjpeg")
DJPEG = shutil.which("djpeg")
pytestmark = pytest.mark.skipif(not (CJPEG and DJPEG), reason="cjpeg・djpeg がない")


def _supports(flag: str) -> bool:
    return flag in subprocess.run([CJPEG, "-h"], capture_output=True, text=True).stderr


def _ppm(tmp_path, w=120, h=90, gray=False):
    r = random.Random(5)
    body = bytearray()
    for y in range(h):
        for x in range(w):
            v = ((x * 2 + y * 3) + r.randrange(12) + (60 if (x // 16 + y // 16) % 2 else 0)) % 256
            body += bytes((v,)) if gray else bytes((v, (x * 2) % 256, (y * 2) % 256))
    p = tmp_path / ("a.pgm" if gray else "a.ppm")
    p.write_bytes(b"%s\n%d %d\n255\n" % (b"P5" if gray else b"P6", w, h) + bytes(body))
    return p


def _ref(jpg, tmp_path):
    """djpeg の画素（PNM）→ (幅, 高さ, 最大値, 値の並び)。"""
    src = tmp_path / "x.jpg"
    src.write_bytes(jpg)
    out = subprocess.run([DJPEG, "-pnm", str(src)], check=True, capture_output=True).stdout
    pos = 0
    tok = []
    while len(tok) < 4:
        e = pos
        while out[e] not in b" \n":
            e += 1
        tok.append(out[pos:e])
        pos = e + 1
    w, h, top = int(tok[1]), int(tok[2]), int(tok[3])
    nc = 3 if tok[0] == b"P6" else 1
    if top > 255:
        vals = [int.from_bytes(out[i:i + 2], "big") for i in range(pos, len(out), 2)]
    else:
        vals = list(out[pos:])
    return w, h, top, nc, vals


def _mine(jpg):
    im = jpeg_gray(jpg, 10**9, want_rgb=True)
    return im


def _max_diff(im, w, h, top, nc, vals):
    assert (im.w, im.h) == (w, h)
    if nc == 1:
        vals = [v for v in vals for _ in range(3)]
    if top > 255:
        sh = len(bin(top)) - 2 - 8
        vals = [min(255, (v + (1 << (sh - 1))) >> sh) for v in vals]
    return max(abs(a - b) for a, b in zip(im.px, vals))


@pytest.mark.parametrize("args,tol", [
    (["-arithmetic", "-sample", "1x1"], 3),
    (["-arithmetic", "-sample", "2x2"], 3),
    (["-arithmetic", "-sample", "2x1", "-restart", "3"], 3),
    (["-arithmetic", "-progressive", "-sample", "2x2"], 3),
    (["-arithmetic", "-progressive", "-sample", "1x1", "-restart", "2"], 3),
    (["-progressive", "-sample", "2x2"], 3),
    (["-sample", "2x2"], 3),
    (["-sample", "2x1"], 3),
    (["-sample", "1x2"], 4),
], ids=["arith-444", "arith-420", "arith-422-restart", "arith-prog-420", "arith-prog-444-restart", "prog-420", "base-420", "base-422", "base-440"])
def test_lossy_modes_match_djpeg(tmp_path, args, tol):
    jpg = subprocess.run([CJPEG, "-quality", "85"] + args + [str(_ppm(tmp_path))], check=True, capture_output=True).stdout
    assert _max_diff(_mine(jpg), *_ref(jpg, tmp_path)) <= tol


def test_arithmetic_gray(tmp_path):
    jpg = subprocess.run([CJPEG, "-quality", "85", "-arithmetic", str(_ppm(tmp_path, gray=True))], check=True, capture_output=True).stdout
    assert _max_diff(_mine(jpg), *_ref(jpg, tmp_path)) <= 2


@pytest.mark.parametrize("args", [["-precision", "12"], ["-precision", "12", "-progressive"], ["-precision", "12", "-arithmetic"]], ids=["base", "prog", "arith"])
def test_12_bit(tmp_path, args):
    if not _supports("-precision"):
        pytest.skip("この cjpeg は 12 ビットを作れない")
    jpg = subprocess.run([CJPEG, "-quality", "85"] + args + [str(_ppm(tmp_path))], check=True, capture_output=True).stdout
    assert jpg[:2] == b"\xff\xd8"
    assert _max_diff(_mine(jpg), *_ref(jpg, tmp_path)) <= 3


@pytest.mark.parametrize("psv", range(1, 8))
def test_lossless_exact(tmp_path, psv):
    if not _supports("-lossless"):
        pytest.skip("この cjpeg は可逆を作れない")
    jpg = subprocess.run([CJPEG, "-lossless", str(psv), "-rgb", str(_ppm(tmp_path))], check=True, capture_output=True).stdout
    assert _max_diff(_mine(jpg), *_ref(jpg, tmp_path)) == 0


@pytest.mark.parametrize("arg", ["1,2", "4,3"])
def test_lossless_point_transform_and_gray(tmp_path, arg):
    if not _supports("-lossless"):
        pytest.skip("この cjpeg は可逆を作れない")
    jpg = subprocess.run([CJPEG, "-lossless", arg, "-restart", "5", str(_ppm(tmp_path, gray=True))], check=True, capture_output=True).stdout
    assert _max_diff(_mine(jpg), *_ref(jpg, tmp_path)) == 0


@pytest.mark.parametrize("prec", [12, 16])
def test_lossless_high_precision(tmp_path, prec):
    if not _supports("-lossless"):
        pytest.skip("この cjpeg は可逆を作れない")
    jpg = subprocess.run([CJPEG, "-lossless", "1", "-precision", str(prec), str(_ppm(tmp_path, gray=True))], check=True, capture_output=True).stdout
    assert _max_diff(_mine(jpg), *_ref(jpg, tmp_path)) <= 1


def test_unusual_jpeg_goes_into_pdf_as_pixels(tmp_path):
    """算術符号の JPEG は PDF の DCTDecode に入れられない：画素にして包み、ふつうの JPEG はそのまま入れる。"""
    ppm = _ppm(tmp_path)
    plain = subprocess.run([CJPEG, str(ppm)], check=True, capture_output=True).stdout
    arith = subprocess.run([CJPEG, "-arithmetic", str(ppm)], check=True, capture_output=True).stdout
    assert b"/DCTDecode" in S.to_pdf(plain)[0]
    pdf = S.to_pdf(arith)[0]
    assert b"/DCTDecode" not in pdf and b"/FlateDecode" in pdf


def test_view_frame_reports_nothing_lost_for_arithmetic(tmp_path):
    jpg = subprocess.run([CJPEG, "-arithmetic", str(_ppm(tmp_path))], check=True, capture_output=True).stdout
    v = S.view_frame(jpg, 1, [10, 10, 60, 50])
    assert v["not_drawn"] == [] and "元の解像度" in v["text"]
