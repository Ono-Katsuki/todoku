"""SVG と PICT を、画素の値で確かめる。SVG は手で書いた絵、PICT は手で組み立てた命令と、道具（magick・sips）があれば作ったファイル。"""

from __future__ import annotations

import base64
import random
import shutil
import struct
import subprocess
import zlib

import pytest

from todoku import api
from todoku.imaging import formats as F
from todoku.imaging import pict, svg
from todoku.imaging import standalone as S


def _px(fr, x, y):
    o = (y * fr.w + x) * 3
    return tuple(fr.rgb[o:o + 3])


# 1 CSS px を 1 画素で描く（target_pixels を CSS の大きさにする）
def _svg(body: str, w=100, h=100, extra="") -> bytes:
    return f'<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" width="{w}" height="{h}" {extra}>{body}</svg>'.encode()


def _draw(blob, w=100, h=100):
    return svg.render(blob, w * h)


def test_svg_sniff_and_info():
    b = _svg("<rect width='10' height='10'/>", 96, 192)
    assert F.sniff(b) == "svg"
    assert F.sniff(b'<?xml version="1.0"?>\n<!-- c -->\n<!DOCTYPE svg PUBLIC "x" "y">\n<svg xmlns="a"></svg>') == "svg"
    assert F.sniff(b"<html><svg></svg></html>") is None
    assert F.sniff(zlib.compress(b, wbits=31)) == "svg"
    fmt, fr = F.info(b)
    assert fmt == "svg" and (fr[0].w, fr[0].h) == (300, 600) and fr[0].dpi == (300.0, 300.0)
    # 紙の上の大きさ
    _f, fr = F.info(b'<svg xmlns="a" width="210mm" height="297mm"></svg>')
    assert abs(fr[0].w / fr[0].dpi[0] * 25.4 - 210) < 0.2 and abs(fr[0].h / fr[0].dpi[1] * 25.4 - 297) < 0.2


def test_svg_shapes_and_colors():
    fr = _draw(_svg("""<rect x="10" y="10" width="20" height="20" fill="red"/>
<circle cx="70" cy="20" r="10" fill="#00ff00"/>
<ellipse cx="20" cy="70" rx="15" ry="5" fill="rgb(0,0,255)"/>
<polygon points="60,60 90,60 90,90" fill="yellow"/>
<line x1="0" y1="50" x2="100" y2="50" stroke="black" stroke-width="2"/>"""))
    assert _px(fr, 20, 20) == (255, 0, 0) and _px(fr, 70, 20) == (0, 255, 0) and _px(fr, 20, 70) == (0, 0, 255)
    assert _px(fr, 85, 65) == (255, 255, 0) and _px(fr, 65, 85) == (255, 255, 255)
    assert _px(fr, 50, 50) == (0, 0, 0) and _px(fr, 50, 45) == (255, 255, 255)


@pytest.mark.parametrize("d,inside,outside", [
    ("M10 10H50V50H10Z", (30, 30), (60, 30)),
    ("m10 10 40 0 0 40 -40 0z", (30, 30), (5, 30)),
    ("M10 10L50 10 50 50 10 50Z", (30, 30), (60, 60)),
    ("M10,30 C10,10 50,10 50,30 S90,50 90,30 V90 H10 Z", (50, 60), (5, 5)),
    ("M10 50 Q50 0 90 50 T90 90 H10Z", (50, 60), (5, 5)),
    ("M20 50 A30 30 0 1 1 80 50 A30 30 0 1 1 20 50Z", (50, 50), (5, 5)),
    ("M20 50a30 30 0 1 1 60 0a30 30 0 1 1-60 0z", (50, 50), (5, 5)),
    ("M20 50a30,30 0 1160 0a30,30 0 11-60 0z", (50, 50), (5, 5)),  # 弧の旗が区切りなしに続く
])
def test_svg_path_commands(d, inside, outside):
    fr = _draw(_svg(f'<path d="{d}" fill="#123456"/>'))
    assert _px(fr, *inside) == (0x12, 0x34, 0x56) and _px(fr, *outside) == (255, 255, 255)


def test_svg_evenodd_and_transform_and_viewbox():
    ring = 'M10 10H90V90H10Z M30 30H70V70H30Z'
    fr = _draw(_svg(f'<path d="{ring}" fill="black" fill-rule="evenodd"/>'))
    assert _px(fr, 50, 50) == (255, 255, 255) and _px(fr, 20, 20) == (0, 0, 0)
    fr = _draw(_svg(f'<path d="{ring}" fill="black"/>'))
    assert _px(fr, 50, 50) == (0, 0, 0)
    # transform の順：後ろから先に効く
    fr = _draw(_svg('<g transform="translate(50 0)"><rect transform="scale(2)" x="5" y="5" width="10" height="10" fill="red"/></g>'))
    assert _px(fr, 65, 15) == (255, 0, 0) and _px(fr, 15, 15) == (255, 255, 255)
    fr = _draw(_svg('<rect transform="rotate(90 50 50)" x="10" y="45" width="30" height="10" fill="red"/>'))
    assert _px(fr, 50, 25) == (255, 0, 0) and _px(fr, 25, 50) == (255, 255, 255)
    # viewBox：10 倍して、余りを真ん中に
    fr = svg.render(_svg('<rect x="0" y="0" width="10" height="10" fill="red"/>', 100, 200, 'viewBox="0 0 10 10"'), 100 * 200)
    assert _px(fr, 50, 100) == (255, 0, 0) and _px(fr, 50, 20) == (255, 255, 255)


def test_svg_stroke_opacity_dash():
    fr = _draw(_svg('<rect x="20" y="20" width="60" height="60" fill="none" stroke="#0000ff" stroke-width="4"/>'))
    assert _px(fr, 50, 20) == (0, 0, 255) and _px(fr, 50, 50) == (255, 255, 255)
    fr = _draw(_svg('<rect x="0" y="0" width="100" height="100" fill="black" fill-opacity="0.5"/>'))
    assert abs(_px(fr, 50, 50)[0] - 128) <= 1
    fr = _draw(_svg('<g opacity="0.5"><rect x="0" y="0" width="100" height="100" fill="black" opacity="0.5"/></g>'))
    assert abs(_px(fr, 50, 50)[0] - 191) <= 2
    fr = _draw(_svg('<line x1="0" y1="50" x2="100" y2="50" stroke="black" stroke-width="4" stroke-dasharray="10 10"/>'))
    assert _px(fr, 5, 50) == (0, 0, 0) and _px(fr, 15, 50) == (255, 255, 255)


def test_svg_css_use_clip_inherit():
    fr = _draw(_svg("""<style>.a{fill:#ff0000} #b{fill:#00ff00} rect.c{fill:#0000ff}</style>
<rect class="a" x="0" y="0" width="10" height="10"/><rect id="b" class="a" x="20" y="0" width="10" height="10"/><rect class="c" x="40" y="0" width="10" height="10" fill="black"/>
<rect x="60" y="0" width="10" height="10" fill="black" style="fill:yellow"/>"""))
    assert _px(fr, 5, 5) == (255, 0, 0) and _px(fr, 25, 5) == (0, 255, 0) and _px(fr, 45, 5) == (0, 0, 255) and _px(fr, 65, 5) == (255, 255, 0)
    fr = _draw(_svg("""<defs><rect id="r" width="10" height="10"/><symbol id="s" viewBox="0 0 10 10"><rect width="10" height="10" fill="blue"/></symbol></defs>
<g fill="red"><use xlink:href="#r" x="20" y="20"/></g><use href="#s" x="50" y="50" width="20" height="20"/>"""))
    assert _px(fr, 25, 25) == (255, 0, 0) and _px(fr, 5, 5) == (255, 255, 255) and _px(fr, 65, 65) == (0, 0, 255)
    fr = _draw(_svg("""<clipPath id="c"><circle cx="50" cy="50" r="20"/></clipPath><rect width="100" height="100" fill="green" clip-path="url(#c)"/>"""))
    assert _px(fr, 50, 50) == (0, 128, 0) and _px(fr, 20, 20) == (255, 255, 255)


def test_svg_gradient_reported():
    fr = _draw(_svg("""<linearGradient id="g"><stop offset="0" stop-color="#000"/><stop offset="1" stop-color="#fff"/></linearGradient>
<rect width="100" height="100" fill="url(#g)"/>"""))
    assert _px(fr, 2, 50)[0] < 20 and _px(fr, 97, 50)[0] > 235 and abs(_px(fr, 50, 50)[0] - 128) < 10
    assert any("グラデーション" in n for n in fr.notes)
    fr = _draw(_svg("""<radialGradient id="g"><stop offset="0" stop-color="#f00"/><stop offset="1" stop-color="#00f"/></radialGradient>
<rect width="100" height="100" fill="url(#g)"/>"""))
    assert _px(fr, 50, 50)[0] > 240 and _px(fr, 99, 50)[2] > 235


def test_svg_embedded_image_and_external_reported():
    from test_image_files import _png

    raw = b"".join(b"\0" + bytes((10, 200, 30)) * 4 for _ in range(4))
    uri = "data:image/png;base64," + base64.b64encode(_png(4, 4, 2, 8, raw)).decode()
    fr = _draw(_svg(f'<image href="{uri}" x="20" y="20" width="40" height="40"/>'))
    assert _px(fr, 40, 40) == (10, 200, 30) and _px(fr, 10, 10) == (255, 255, 255)
    fr = _draw(_svg('<image href="http://example.com/a.png" x="20" y="20" width="40" height="40"/>'))
    assert any("外のファイル" in p for p in fr.problems)
    assert _px(fr, 20, 20)[0] < 255  # 灰色の枠で示す


def test_svg_text_drawn_and_listed(tmp_path):
    b = _svg('<text x="10" y="50" font-size="24">氏名 ABC<tspan x="10" y="80">住所</tspan></text>', 200, 100)
    fr = svg.render(b, 200 * 100)
    assert any("氏名 ABC" in n for n in fr.notes)
    from todoku.pdf.render.fontmap import system_face

    if system_face("sans") is not None or system_face("cjk-sans") is not None:
        dark = sum(1 for i in range(0, len(fr.rgb), 3) if fr.rgb[i] < 128)
        assert dark > 50
    win = S.build(b, "a.svg")
    assert win["kind"] == "image" and any("住所" in t for t in win["texts"])
    assert "氏名" in S.markdown_parts(win)[0][1]


def test_svg_unsupported_reported_not_silent():
    fr = _draw(_svg('<filter id="f"/><rect width="10" height="10" filter="url(#f)"/><foreignObject/><script>alert(1)</script>'))
    assert any("フィルター" in p for p in fr.problems) and any("foreignObject" in p for p in fr.problems)
    assert not any("script" in p for p in fr.problems)  # 実行しない：黙って無視してよい


def test_svg_entities_refused():
    bomb = b'<?xml version="1.0"?><!DOCTYPE svg [<!ENTITY a "aaaa"><!ENTITY b "&a;&a;&a;">]><svg xmlns="x"><text>&b;</text></svg>'
    with pytest.raises(F.ImageError):
        F.decode(bomb)


def test_svg_use_bomb_stops():
    body = '<defs><g id="a0"><rect width="1" height="1"/></g>' + "".join(f'<g id="a{i}">' + f'<use href="#a{i - 1}"/>' * 10 + "</g>" for i in range(1, 9)) + '</defs><use href="#a8"/>'
    fr = _draw(_svg(body))
    assert any("途中で止めました" in p or "深すぎる" in p for p in fr.problems)


def test_svg_zooms_smoothly_with_region():
    """ベクターは範囲を狭めれば細かく描く：縁がぼけず、範囲の外の画素を作らない。"""
    b = _svg('<rect x="50" y="50" width="25" height="25" fill="black"/>')
    v = S.view_frame(b, 1, [60 * 25.4 / 96, 60 * 25.4 / 96, 80 * 25.4 / 96, 80 * 25.4 / 96])
    assert v["width"] > 60
    fr = F.decode(b, 0, None, None, (150, 150, 300, 300))
    assert fr.window == (150, 150, 300, 300) and (fr.w, fr.h) == (150, 150)
    assert _px(fr, 10, 10) == (0, 0, 0) and _px(fr, 100, 100) == (255, 255, 255)


def test_svg_opens_as_a_document(tmp_path):
    p = tmp_path / "a.svg"
    p.write_bytes(_svg('<text x="1" y="10">こんにちは</text>'))
    kind, win, parts = api.window(str(p))
    assert kind == "image" and "こんにちは" in str(parts)
    pdf, listing, _n = S.to_pdf(p.read_bytes())
    assert pdf[:4] == b"%PDF" and listing


# ====================================================================== PICT：手で組み立てる
def _pict2(ops: bytes, w=100, h=100) -> bytes:
    head = b"\0" * 512 + struct.pack(">Hhhhh", 0, 0, 0, h, w) + b"\x00\x11\x02\xff" + b"\x0c\x00" + struct.pack(">h", -2) + b"\0\0" + struct.pack(">II", 72 << 16, 72 << 16)
    head += struct.pack(">hhhh", 0, 0, h, w) + b"\0\0\0\0"
    return head + ops + b"\x00\xff"


def _rect(t, l, b, r):
    return struct.pack(">hhhh", t, l, b, r)


def _rgb16(r, g, b):
    return struct.pack(">HHH", r * 257, g * 257, b * 257)


def test_pict_sniff_info():
    b = _pict2(b"")
    assert F.sniff(b) == "pict" and F.sniff(b[512:]) == "pict"
    fmt, fr = F.info(b)
    assert (fr[0].w, fr[0].h) == (100, 100) and fr[0].dpi == (72.0, 72.0)


def test_pict_vector_ops():
    ops = b"\x00\x1a" + _rgb16(255, 0, 0) + b"\x00\x31" + _rect(10, 10, 30, 30)  # paintRect
    ops += b"\x00\x1a" + _rgb16(0, 0, 255) + b"\x00\x51" + _rect(10, 60, 30, 90)  # paintOval
    ops += b"\x00\x1a" + _rgb16(0, 255, 0) + b"\x00\x07" + struct.pack(">hh", 3, 3) + b"\x00\x30" + _rect(50, 10, 90, 50)  # frameRect
    ops += b"\x00\x1a" + _rgb16(0, 0, 0) + b"\x00\x07" + struct.pack(">hh", 1, 1) + b"\x00\x20" + struct.pack(">hhhh", 60, 60, 60, 95)  # line
    poly = struct.pack(">H", 10 + 12) + _rect(70, 60, 95, 95) + struct.pack(">hhhhhh", 70, 60, 70, 95, 95, 95)
    ops += b"\x00\x1a" + _rgb16(255, 255, 0) + b"\x00\x71" + poly  # paintPoly
    fr = pict.render(_pict2(ops))
    assert _px(fr, 20, 20) == (255, 0, 0) and _px(fr, 75, 20) == (0, 0, 255) and _px(fr, 62, 12) == (255, 255, 255)
    assert _px(fr, 11, 70) == (0, 255, 0) and _px(fr, 30, 70) == (255, 255, 255)
    assert _px(fr, 70, 60) == (0, 0, 0) and _px(fr, 90, 75) == (255, 255, 0) and fr.problems == []


def test_pict_last_shape_and_patterns():
    ops = b"\x00\x1a" + _rgb16(255, 0, 0) + b"\x00\x31" + _rect(10, 10, 30, 30)
    ops += b"\x00\x1a" + _rgb16(0, 0, 255) + b"\x00\x39"  # 同じ四角をもう一度
    ops += b"\x00\x09" + b"\xaa\x55" * 4 + b"\x00\x1a" + _rgb16(0, 0, 0) + b"\x00\x31" + _rect(50, 50, 58, 58)
    fr = pict.render(_pict2(ops))
    assert _px(fr, 20, 20) == (0, 0, 255)
    assert {_px(fr, 50, 50), _px(fr, 51, 50)} == {(0, 0, 0), (255, 255, 255)}


def _packbits(row: bytes) -> bytes:
    """いちばん簡単な PackBits：リテラルだけ（128 バイトごと）。"""
    out = bytearray()
    for i in range(0, len(row), 128):
        chunk = row[i:i + 128]
        out.append(len(chunk) - 1)
        out += chunk
    return bytes(out)


def test_pict_direct_32_packtype4():
    w, h = 20, 10
    rows = b""
    for y in range(h):
        r = bytes((x * 10) & 255 for x in range(w))
        g = bytes((y * 20) & 255 for _ in range(w))
        b = bytes(99 for _ in range(w))
        pk = _packbits(r + g + b)
        rows += bytes((len(pk),)) + pk
    # 0x9A：baseAddr・rowBytes（高いビット）・bounds・版…
    pm = struct.pack(">IH", 0xFF, 0x8000 | (w * 4)) + _rect(0, 0, h, w) + struct.pack(">HHIIIHHHHIII", 0, 4, 0, 72 << 16, 72 << 16, 16, 32, 3, 8, 0, 0, 0)
    ops = b"\x00\x9a" + pm + _rect(0, 0, h, w) + _rect(0, 0, h, w) + struct.pack(">H", 0) + rows
    if len(ops) & 1:
        ops += b"\0"
    fr = pict.render(_pict2(ops, w, h))
    assert _px(fr, 3, 2) == (30, 40, 99) and _px(fr, 19, 9) == (190, 180, 99)


def test_pict_indexed_packbits_rect():
    w, h = 16, 4
    data = b""
    for y in range(h):
        row = bytes((x + y) % 2 for x in range(w))
        pk = _packbits(row)
        data += bytes((len(pk),)) + pk
    pm = struct.pack(">H", 0x8000 | w) + _rect(0, 0, h, w) + struct.pack(">HHIIIHHHHIII", 0, 0, 0, 72 << 16, 72 << 16, 0, 8, 1, 8, 0, 0, 0)
    ct = struct.pack(">IHH", 0, 0, 1) + struct.pack(">H", 0) + _rgb16(200, 10, 20) + struct.pack(">H", 1) + _rgb16(1, 2, 3)
    ops = b"\x00\x98" + pm + ct + _rect(0, 0, h, w) + _rect(0, 0, h, w) + struct.pack(">H", 0) + data
    if len(ops) & 1:
        ops += b"\0"
    fr = pict.render(_pict2(ops, w, h))
    assert _px(fr, 0, 0) == (200, 10, 20) and _px(fr, 1, 0) == (1, 2, 3) and _px(fr, 0, 1) == (1, 2, 3)


def test_pict_version1_bitmap_and_text():
    # 版 1：1 バイトの命令。1 ビットの BitsRect（0x90）
    head = struct.pack(">Hhhhh", 0, 0, 0, 16, 16) + b"\x11\x01"
    bits = bytes(x for y in range(8) for x in (0xF0, 0x0F))
    ops = b"\x90" + struct.pack(">H", 2) + _rect(0, 0, 8, 16) + _rect(0, 0, 8, 16) + _rect(0, 0, 8, 16) + struct.pack(">H", 0) + bits + b"\xff"
    b = head + ops
    assert F.sniff(b) == "pict"
    fr = pict.render(b)
    assert _px(fr, 0, 0) == (0, 0, 0) and _px(fr, 5, 0) == (255, 255, 255) and _px(fr, 12, 3) == (0, 0, 0)


def test_pict_embedded_jpeg(tmp_path):
    cj = shutil.which("cjpeg")
    if not cj:
        pytest.skip("cjpeg がない")
    ppm = tmp_path / "a.ppm"
    ppm.write_bytes(b"P6\n16 16\n255\n" + bytes((10, 200, 30)) * 256)
    jpg = subprocess.run([cj, "-quality", "95", str(ppm)], check=True, capture_output=True).stdout
    # 版(2)・行列(36)・matte の大きさ(4)・matte の四角(8)・描き方(2)・元の四角(8)・精度(4)・マスクの大きさ(4)・画像の説明(86)・データ
    mat = struct.pack(">9i", 1 << 16, 0, 0, 0, 1 << 16, 0, 20 << 16, 10 << 16, 1 << 30)
    head = struct.pack(">H", 0) + mat + struct.pack(">I", 0) + _rect(0, 0, 0, 0) + struct.pack(">H", 0) + _rect(0, 0, 16, 16) + struct.pack(">I", 0) + struct.pack(">I", 0)
    desc = struct.pack(">I", 86) + b"jpeg" + bytes(24) + struct.pack(">HH", 16, 16) + bytes(86 - 36)
    body = head + desc + jpg
    ops = b"\x82\x00" + struct.pack(">I", len(body)) + body
    if len(ops) & 1:
        ops += b"\0"
    fr = pict.render(_pict2(ops, 50, 50))
    assert fr.problems == []
    r, g, b = _px(fr, 28, 18)
    assert abs(r - 10) < 8 and abs(g - 200) < 8 and abs(b - 30) < 8
    assert _px(fr, 5, 5) == (255, 255, 255)


def test_pict_unknown_ops_skipped_by_size():
    ops = b"\x00\xa1" + struct.pack(">HH", 100, 4) + b"abcd" + b"\x01\x00" + b"\0\0" + b"\x80\x05" + b"\x00\xd0" + struct.pack(">I", 2) + b"zz"
    ops += b"\x00\x1a" + _rgb16(255, 0, 0) + b"\x00\x31" + _rect(1, 1, 9, 9)
    fr = pict.render(_pict2(ops, 10, 10))
    assert _px(fr, 5, 5) == (255, 0, 0)


@pytest.mark.parametrize("tool", ["magick", "sips"])
def test_pict_made_by_tools(tmp_path, tool):
    exe = shutil.which(tool)
    if not exe:
        pytest.skip(f"{tool} がない")
    r = random.Random(1)
    w, h = 64, 48
    body = bytearray()
    for y in range(h):
        for x in range(w):
            body += bytes(((x * 4) & 255, (y * 5) & 255, r.randrange(256) if (x // 8 + y // 8) % 2 else 40))
    ppm = tmp_path / "a.ppm"
    ppm.write_bytes(b"P6\n%d %d\n255\n" % (w, h) + bytes(body))
    out = tmp_path / "a.pict"
    if tool == "magick":
        c = subprocess.run([exe, str(ppm), f"PICT:{out}"], capture_output=True)
    else:
        png = tmp_path / "a.png"
        subprocess.run([exe, "-s", "format", "png", str(ppm), "--out", str(png)], capture_output=True)
        c = subprocess.run([exe, "-s", "format", "pict", str(png), "--out", str(out)], capture_output=True)
    if c.returncode or not out.exists() or out.stat().st_size < 100:
        pytest.skip(f"{tool} は PICT を作れない")
    fr = F.decode(out.read_bytes())
    assert (fr.w, fr.h) == (w, h) and fr.problems == []
    diff = sum(abs(a - b) for a, b in zip(fr.rgb, body)) / len(body)
    assert diff < 6
