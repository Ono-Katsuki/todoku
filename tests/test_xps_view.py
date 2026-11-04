"""XPS の view_page：手で作った FixedPage（図形・字・難読化した書体・画像）を描いて、画素で確かめる。"""

from __future__ import annotations

import io
import random
import struct
import uuid
import zipfile
import zlib

import pytest

from todoku import api
from todoku.imaging.formats import decode
from todoku.legacy import xpsrender

import base64

M = 26  # 目盛りの帯の画素


def glyph_square(x0, y0, x1, y1) -> bytes:
    pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1)]
    # TrueType：時計回りが塗り
    pts = pts[::-1]
    out = struct.pack(">hhhhh", 1, x0, y0, x1, y1) + struct.pack(">H", 3) + struct.pack(">H", 0)
    out += bytes((1,)) * 4
    px = py = 0
    xs = b""
    ys = b""
    for x, y in pts:
        xs += struct.pack(">h", x - px)
        ys += struct.pack(">h", y - py)
        px, py = x, y
    out += xs + ys
    return out + bytes(len(out) % 2)


def make_font() -> bytes:
    # 字形 0：空。字形 1（「A」）：全体の四角。字形 2（「B」）：下半分の四角
    glyphs = [b"", glyph_square(0, 0, 1000, 1000), glyph_square(0, 0, 1000, 500)]
    loca = []
    gl = b""
    for g in glyphs:
        loca.append(len(gl))
        gl += g
    loca.append(len(gl))
    head = bytearray(54)
    head[0:4] = b"\x00\x01\x00\x00"
    struct.pack_into(">H", head, 18, 1000)
    struct.pack_into(">h", head, 50, 1)
    maxp = struct.pack(">IH", 0x10000, len(glyphs)) + bytes(26)
    hhea = bytearray(36)
    struct.pack_into(">H", hhea, 34, 3)
    hmtx = struct.pack(">HhHhHh", 500, 0, 1000, 0, 2000, 0)
    groups = [(65, 66, 1)]
    cmap_sub = struct.pack(">HHIII", 12, 0, 16 + 12 * len(groups), 0, len(groups)) + b"".join(struct.pack(">III", *g) for g in groups)
    cmap = struct.pack(">HH", 0, 1) + struct.pack(">HHI", 3, 10, 12) + cmap_sub
    tabs = {"cmap": cmap, "glyf": gl, "head": bytes(head), "hhea": bytes(hhea), "hmtx": hmtx, "loca": struct.pack(f">{len(loca)}I", *loca), "maxp": maxp}
    n = len(tabs)
    off = 12 + 16 * n
    table = b""
    body = b""
    for t, d in tabs.items():
        table += t.encode() + struct.pack(">III", 0, off + len(body), len(d))
        body += d + bytes((4 - len(d) % 4) % 4)
    return struct.pack(">IHHHH", 0x10000, n, 0, 0, 0) + table + body


def obfuscate(font: bytes, guid: str) -> bytes:
    key = bytes.fromhex(guid.replace("-", ""))[::-1]
    head = bytearray(font[:32])
    for i in range(16):
        head[i] ^= key[i]
        head[i + 16] ^= key[i]
    return bytes(head) + font[32:]


def png(w: int, h: int, rgb: tuple[int, int, int]) -> bytes:
    raw = b"".join(b"\0" + bytes(rgb) * w for _ in range(h))

    def ch(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)

    return b"\x89PNG\r\n\x1a\n" + ch(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + ch(b"IDAT", zlib.compress(raw)) + ch(b"IEND", b"")


def make_xps(pages: list[str], parts: dict[str, bytes] | None = None) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("_rels/.rels", '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                                  '<Relationship Id="R0" Type="http://schemas.microsoft.com/xps/2005/06/fixedrepresentation" Target="/FixedDocSeq.fdseq"/></Relationships>')
        z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
        z.writestr("FixedDocSeq.fdseq", '<FixedDocumentSequence xmlns="http://schemas.microsoft.com/xps/2005/06"><DocumentReference Source="/Documents/1/F.fdoc"/></FixedDocumentSequence>')
        pcs = "".join(f'<PageContent Source="Pages/{i + 1}.fpage"/>' for i in range(len(pages)))
        z.writestr("Documents/1/F.fdoc", f'<FixedDocument xmlns="http://schemas.microsoft.com/xps/2005/06">{pcs}</FixedDocument>')
        for i, p in enumerate(pages):
            z.writestr(f"Documents/1/Pages/{i + 1}.fpage", p)
        for n, b in (parts or {}).items():
            z.writestr(n, b)
    return buf.getvalue()


NS = 'xmlns="http://schemas.microsoft.com/xps/2005/06" xmlns:x="http://schemas.microsoft.com/xps/2005/06/resourcedictionary-key"'


def page(body: str, w: int = 480, h: int = 480, res: str = "") -> str:
    return f'<FixedPage {NS} Width="{w}" Height="{h}">{res}{body}</FixedPage>'


def view(tmp_path, blob, n=1, **kw):
    p = tmp_path / "a.xps"
    p.write_bytes(blob)
    v = api.view_page(str(p), n, **kw)
    fr = decode(base64.b64decode(v["png_base64"]))
    return v, fr


def px_at(fr, v, x96: float, y96: float, dpi: float = 96) -> tuple[int, int, int]:
    """FixedPage の座標（1/96 インチ）の画素（dpi=96 で描いたとき 1 単位＝1 画素）。"""
    X, Y = int(x96 * dpi / 96) + M, int(y96 * dpi / 96) + M
    i = (Y * fr.w + X) * 3
    return tuple(fr.rgb[i:i + 3])


def test_paths_transform_clip_opacity(tmp_path):
    body = ('<Path Fill="#FF0000" Data="M 10,10 L 110,10 110,110 10,110 Z"/>'
            '<Canvas RenderTransform="1,0,0,1,200,0"><Path Fill="#FF00FF00" Data="M 0,0 H 50 V 50 H 0 z"/></Canvas>'
            '<Path Fill="#000080" Clip="M 300,300 L 350,300 350,350 300,350 Z" Data="M 280,280 L 400,280 400,400 280,400 Z"/>'
            '<Path Fill="#800000FF" Data="M 10,200 L 60,200 60,250 10,250 Z"/>'
            '<Path Stroke="#000000" StrokeThickness="4" Data="M 100,300 L 200,300"/>'
            '<Path Fill="#0000FF"><Path.Data><PathGeometry><PathFigure StartPoint="400,10" IsClosed="true"><PolyLineSegment Points="450,10 450,60 400,60"/></PathFigure></PathGeometry></Path.Data></Path>'
            '<Path Fill="#FF00FF" Data="M 400,100 A 20,20 0 1 1 440,100 A 20,20 0 1 1 400,100 Z"/>')
    v, fr = view(tmp_path, make_xps([page(body)]), dpi=96)
    assert fr.w - M == 480 and fr.h - M == 480
    assert px_at(fr, v, 50, 50) == (255, 0, 0)
    assert px_at(fr, v, 225, 25) == (0, 255, 0)
    assert px_at(fr, v, 325, 325) == (0, 0, 128)
    assert px_at(fr, v, 290, 290) == (255, 255, 255)  # 切り抜きの外
    r, g, b = px_at(fr, v, 30, 220)
    assert abs(r - 127) < 4 and abs(g - 127) < 4 and b == 255  # 半透明
    assert px_at(fr, v, 150, 300) == (0, 0, 0)
    assert px_at(fr, v, 425, 30) == (0, 0, 255)
    assert px_at(fr, v, 420, 100) == (255, 0, 255) and px_at(fr, v, 395, 80) == (255, 255, 255)  # 弧


def test_glyphs_with_obfuscated_font(tmp_path):
    guid = str(uuid.UUID(int=0x123456789ABCDEF0FEDCBA9876543210)).upper()
    font = make_font()
    parts = {f"Resources/{{{guid}}}.odttf": obfuscate(font, guid)}
    uri = f"/Resources/{{{guid}}}.odttf"
    # 字形の番号で（UnicodeString なし）。2 つ目は送りを 200/100 em で指定
    body = (f'<Glyphs FontUri="{uri}" FontRenderingEmSize="40" OriginX="20" OriginY="100" Fill="#FF0000" Indices="1,100;1,200;1"/>'
            f'<Glyphs FontUri="{uri}" FontRenderingEmSize="40" OriginX="20" OriginY="200" Fill="#0000FF" UnicodeString="AB"/>')
    v, fr = view(tmp_path, make_xps([page(body)], parts), dpi=96)
    assert px_at(fr, v, 40, 80) == (255, 0, 0)  # 1 つ目の字
    assert px_at(fr, v, 80, 80) == (255, 0, 0)  # 2 つ目（送り 100/100 em）
    assert px_at(fr, v, 110, 80) == (255, 255, 255)  # 2 つ目の字の後ろの空き（送り 200/100 em）
    assert px_at(fr, v, 150, 80) == (255, 0, 0)  # 3 つ目
    assert px_at(fr, v, 30, 170) == (0, 0, 255)  # A：書体の送り 1000/1000
    assert px_at(fr, v, 70, 170) == (255, 255, 255)  # B の上半分は空き（B は下半分だけ）
    assert px_at(fr, v, 70, 195) == (0, 0, 255)


def test_plain_ttf_and_missing_font(tmp_path):
    body = ('<Glyphs FontUri="/f.ttf" FontRenderingEmSize="40" OriginX="20" OriginY="100" Fill="#00FF00" UnicodeString="A"/>'
            '<Glyphs FontUri="/nofont.ttf" FontRenderingEmSize="40" OriginX="20" OriginY="200" Fill="#000000" UnicodeString="A"/>')
    v, fr = view(tmp_path, make_xps([page(body)], {"f.ttf": make_font()}), dpi=96)
    assert px_at(fr, v, 40, 80) == (0, 255, 0)
    assert any("nofont" in x for x in v["not_drawn"]) and "nofont" in v["text"]


def test_image_brush_and_viewbox(tmp_path):
    parts = {"Resources/i.png": png(10, 10, (10, 200, 30))}
    body = ('<Path Data="M 50,50 L 150,50 150,150 50,150 Z"><Path.Fill><ImageBrush ImageSource="/Resources/i.png" Viewbox="0,0,10,10" Viewport="50,50,100,100" TileMode="None" ViewboxUnits="Absolute" ViewportUnits="Absolute"/></Path.Fill></Path>'
            '<Path Data="M 200,200 L 260,200 260,260 200,260 Z"><Path.Fill><ImageBrush ImageSource="/Resources/none.png" Viewbox="0,0,10,10" Viewport="200,200,60,60"/></Path.Fill></Path>')
    v, fr = view(tmp_path, make_xps([page(body)], parts), dpi=96)
    assert px_at(fr, v, 100, 100) == (10, 200, 30)
    assert px_at(fr, v, 45, 45) == (255, 255, 255)
    assert px_at(fr, v, 230, 230) == (205, 205, 205)  # 開けない画像：灰色
    assert any("none.png" in x for x in v["not_drawn"])


def test_resources_and_gradient_note(tmp_path):
    res = ('<FixedPage.Resources><ResourceDictionary><SolidColorBrush x:Key="b" Color="#123456"/><PathGeometry x:Key="g"><PathFigure StartPoint="10,10" IsClosed="true"><PolyLineSegment Points="60,10 60,60 10,60"/></PathFigure></PathGeometry>'
           '<MatrixTransform x:Key="t" Matrix="1,0,0,1,100,0"/></ResourceDictionary></FixedPage.Resources>')
    body = ('<Path Fill="{StaticResource b}" Data="{StaticResource g}" RenderTransform="{StaticResource t}"/>'
            '<Path Data="M 300,10 L 360,10 360,60 300,60 Z"><Path.Fill><LinearGradientBrush><LinearGradientBrush.GradientStops><GradientStop Color="#FF0000" Offset="0"/><GradientStop Color="#FF0000" Offset="1"/></LinearGradientBrush.GradientStops></LinearGradientBrush></Path.Fill></Path>')
    v, fr = view(tmp_path, make_xps([page(body, res=res)]), dpi=96)
    assert px_at(fr, v, 130, 30) == (0x12, 0x34, 0x56)
    assert px_at(fr, v, 30, 30) == (255, 255, 255)
    assert px_at(fr, v, 330, 30) == (255, 0, 0)
    assert any("勾配" in x for x in v["not_drawn"])


def test_region_marks_and_pages(tmp_path):
    p1 = page('<Path Fill="#FF0000" Data="M 0,0 L 480,0 480,480 0,480 Z"/>')
    p2 = page('<Path Fill="#00FF00" Data="M 0,0 L 480,0 480,480 0,480 Z"/>')
    blob = make_xps([p1, p2])
    v, fr = view(tmp_path, blob, 2, region=[10, 10, 50, 50])
    assert fr.rgb[(fr.h // 2 * fr.w + fr.w // 2) * 3:(fr.h // 2 * fr.w + fr.w // 2) * 3 + 3] == b"\x00\xff\x00"
    assert "全部で 2 ページ" in v["text"]
    with pytest.raises(api.ApiError):
        view(tmp_path, blob, 3)
    with pytest.raises(api.ApiError):
        view(tmp_path, blob, 1, region=[1000, 1000, 2000, 2000])
    v, _ = view(tmp_path, blob, 1, marks=[{"box": [10, 10, 20, 20], "label": "x"}])
    assert v["width"] > 100


def test_deobfuscate_rejects_bad_name():
    with pytest.raises(ValueError):
        xpsrender._deobfuscate("Resources/font.odttf", bytes(40))


@pytest.mark.parametrize("seed", range(2))
def test_corruption_never_raw(tmp_path, seed):
    r = random.Random(seed)
    body = '<Path Fill="#FF0000" Data="M 10,10 L 110,10 A 5 5 0 1 1 30 30 Z"/><Glyphs FontUri="/f.ttf" FontRenderingEmSize="40" OriginX="20" OriginY="100" Fill="#00FF00" Indices="(1:2)1,5;x;1,9,9,9" UnicodeString="AB"/>'
    font = bytearray(make_font())
    blob0 = page(body)
    for t in range(25):
        f = bytearray(font)
        for _ in range(r.randrange(1, 12)):
            f[r.randrange(0, len(f))] = r.randrange(256)
        pg = blob0
        if t % 4 == 0:
            pos = r.randrange(0, len(pg))
            pg = pg[:pos] + r.choice(['"', "<", "x", "{", "9e9"]) + pg[pos + 1:]
        p = tmp_path / "c.xps"
        p.write_bytes(make_xps([pg], {"f.ttf": bytes(f)}))
        try:
            api.view_page(str(p), 1, dpi=48)
        except api.ApiError:
            pass


def test_non_page_kinds_still_refuse(tmp_path):
    p = tmp_path / "a.csv"
    p.write_bytes(b"x,y\n1,2\n")
    with pytest.raises(api.ApiError, match="view_page"):
        api.view_page(str(p), 1)
