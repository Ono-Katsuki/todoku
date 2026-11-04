"""JPEG 2000（ISO/IEC 15444-1）の復号：opj_compress（OpenJPEG、別の実装）で符号化した流れを、opj_decompress と同じ画素に開く。
opj_compress がない環境では、手で組み立てた最小の流れだけを確かめる（ほかは飛ばす）。"""

from __future__ import annotations

import random
import re
import shutil
import struct
import subprocess

import pytest

from todoku.imaging import jpx
from todoku.imaging import jpx_t1 as T1

OPJ = shutil.which("opj_compress")
DEC = shutil.which("opj_decompress")
needs_opj = pytest.mark.skipif(not (OPJ and DEC), reason="opj_compress がない")


def _picture(w, h, nc=1, seed=3, bits=8):
    """なめらかな濃淡に、線と字のような塊と少しの乱れ。"""
    r = random.Random(seed)
    top = (1 << bits) - 1
    planes = []
    for c in range(nc):
        rows = []
        for y in range(h):
            rows.append([min(top, max(0, int(((x * 3 + y * (2 + c) + 40 * c) % 256) * top / 255) + r.randrange(-3, 4)))
                         for x in range(w)])
        for _ in range(12):
            x, y, a, b = r.randrange(w), r.randrange(h), r.randrange(2, 40), r.randrange(1, 12)
            v = r.choice((0, top, top // 2))
            for j in range(y, min(h, y + b)):
                for i in range(x, min(w, x + a)):
                    rows[j][i] = v
        planes.append(rows)
    return planes


def _write_pnm(path, planes, bits=8):
    h, w = len(planes[0]), len(planes[0][0])
    nc = len(planes)
    top = (1 << bits) - 1
    head = f"{'P5' if nc == 1 else 'P6'}\n{w} {h}\n{top}\n".encode()
    body = bytearray()
    for y in range(h):
        for x in range(w):
            for c in range(nc):
                v = planes[c][y][x]
                if bits > 8:
                    body += struct.pack(">H", v)
                else:
                    body.append(v)
    path.write_bytes(head + bytes(body))


def _read_pnm(path):
    d = path.read_bytes()
    m = re.match(rb"(P[56])\s+(?:#[^\n]*\n)?(\d+)\s+(\d+)\s+(\d+)\s", d)
    kind, w, h, top = m[1], int(m[2]), int(m[3]), int(m[4])
    nc = 1 if kind == b"P5" else 3
    bps = 2 if top > 255 else 1
    body = d[-w * h * nc * bps:]
    planes = [bytearray(w * h) for _ in range(nc)]
    for c in range(nc):
        if bps == 1:
            planes[c][:] = body[c::nc]
        else:
            planes[c][:] = bytes(int(body[i * 2] * 256 + body[i * 2 + 1]) * 255 // top for i in range(c, w * h * nc, nc))
    return w, h, [bytes(p) for p in planes]


def _compress(tmp_path, planes, *args, bits=8, name="a.j2k"):
    src = tmp_path / ("s.pgm" if len(planes) == 1 else "s.ppm")
    _write_pnm(src, planes, bits)
    out = tmp_path / name
    r = subprocess.run([OPJ, "-i", str(src), "-o", str(out), *map(str, args)], capture_output=True, timeout=60)
    assert r.returncode == 0, r.stdout[-400:] + r.stderr[-400:]
    return out


def _reference(tmp_path, jp, reduce=0, nc=1):
    o = tmp_path / ("o.pgm" if nc == 1 else "o.ppm")
    a = [DEC, "-i", str(jp), "-o", str(o)]
    if reduce:
        a += ["-r", str(reduce)]
    r = subprocess.run(a, capture_output=True, timeout=60)
    assert r.returncode == 0, r.stdout[-400:] + r.stderr[-400:]
    return _read_pnm(o)


def _same(img, ref, tol=0):
    w, h, planes = ref
    assert (img.w, img.h) == (w, h)
    for a, b in zip(img.planes, planes):
        if tol == 0:
            assert a == b
        else:
            worst = max(abs(x - y) for x, y in zip(a, b))
            assert worst <= tol, worst


def _roundtrip(tmp_path, planes, *args, tol=0, reduce=0, bits=8, name="a.j2k"):
    jp = _compress(tmp_path, planes, *args, bits=bits, name=name)
    img = jpx.decode(jp.read_bytes(), reduce=reduce)
    assert not img.problems, img.problems
    _same(img, _reference(tmp_path, jp, reduce, len(planes)), tol)
    return img


# ====================================================================== 可逆（5/3）：画素が完全に一致する
@needs_opj
def test_lossless_gray_exact(tmp_path):
    p = _picture(83, 61)
    img = _roundtrip(tmp_path, p, "-n", 4)
    assert img.planes[0] == bytes(v for r in p[0] for v in r)


@needs_opj
def test_lossless_rgb_with_rct_exact(tmp_path):
    p = _picture(70, 50, 3)
    img = _roundtrip(tmp_path, p, "-n", 3)
    for c in range(3):
        assert img.planes[c] == bytes(v for r in p[c] for v in r)


@needs_opj
@pytest.mark.parametrize("prog", ["LRCP", "RLCP", "RPCL", "PCRL", "CPRL"])
def test_all_progressions_with_precincts_layers_tiles(prog, tmp_path):
    p = _picture(150, 110, 3)
    _roundtrip(tmp_path, p, "-p", prog, "-n", 3, "-r", "40,20,1", "-c", "[64,64],[32,32]", "-b", "16,16", "-t", "64,48")


@needs_opj
@pytest.mark.parametrize("flags", [1, 2, 4, 8, 16, 32, 63, 1 + 4, 1 + 8, 2 + 32])
def test_code_block_styles(flags, tmp_path):
    p = _picture(90, 70)
    _roundtrip(tmp_path, p, "-n", 3, "-M", flags)


@needs_opj
def test_sop_eph_markers_and_small_blocks(tmp_path):
    p = _picture(60, 60, 3)
    _roundtrip(tmp_path, p, "-SOP", "-EPH", "-b", "8,8", "-n", 2)


@needs_opj
def test_tile_parts_and_plt(tmp_path):
    p = _picture(100, 80)
    _roundtrip(tmp_path, p, "-t", "50,40", "-TP", "R", "-PLT", "-n", 2)


@needs_opj
def test_image_offset(tmp_path):
    p = _picture(61, 47)
    _roundtrip(tmp_path, p, "-d", "17,9", "-n", 2)


@needs_opj
def test_tile_offset_odd_origin(tmp_path):
    p = _picture(61, 47)
    _roundtrip(tmp_path, p, "-d", "17,9", "-T", "3,5", "-t", "30,30", "-n", 2)


@needs_opj
def test_region_of_interest_shift(tmp_path):
    p = _picture(80, 60)
    _roundtrip(tmp_path, p, "-ROI", "c=0,U=12", "-n", 2)


@needs_opj
def test_progression_order_change(tmp_path):
    p = _picture(100, 80, 3)
    _roundtrip(tmp_path, p, "-r", "20,5,1", "-n", 3, "-POC", "T1=0,0,1,3,3,CPRL/T1=0,0,3,4,3,CPRL")


def _split_packets(cs: bytes):
    """-SOP -EPH の 1 タイルの流れ：(主な見出し, タイルの見出し, [(パケットの見出し（EPH まで）, 本体)])。"""
    sot = cs.index(b"\xff\x90")
    sod = cs.index(b"\xff\x93", sot)
    pos = sod + 2
    end = len(cs) - 2  # EOC
    packets = []
    while pos < end:
        assert cs[pos:pos + 2] == b"\xff\x91"
        eph = cs.index(b"\xff\x92", pos)
        nxt = cs.find(b"\xff\x91", eph)
        nxt = end if nxt < 0 else nxt
        packets.append((cs[pos + 6:eph + 2], cs[pos:pos + 6], cs[eph + 2:nxt]))
        pos = nxt
    return cs[:sot], cs[sot:sod], packets


def _tile(hdr: bytes, extra: bytes, body: bytes) -> bytes:
    psot = len(hdr) + len(extra) + 2 + len(body)
    return hdr[:6] + struct.pack(">I", psot) + hdr[10:] + extra + b"\xff\x93" + body + b"\xff\xd9"


@needs_opj
def test_packed_packet_headers_ppt_and_ppm(tmp_path):
    p = _picture(70, 50, 3)
    jp = _compress(tmp_path, p, "-SOP", "-EPH", "-n", 2, "-r", "10,1")
    cs = jp.read_bytes()
    want = jpx.decode(cs).planes
    main, hdr, packets = _split_packets(cs)
    heads = b"".join(h for h, _s, _b in packets)
    body = b"".join(s + b for _h, s, b in packets)
    # PPT（A.7.5）：タイルの部分の見出しの中
    ppt = b"\xff\x61" + struct.pack(">HB", 3 + len(heads), 0) + heads
    img = jpx.decode(main + _tile(hdr, ppt, body))
    assert not img.problems and img.planes == want
    # PPM（A.7.4）：主な見出しの中。(Nppm, Ippm)
    ppm = b"\xff\x60" + struct.pack(">HB", 3 + 4 + len(heads), 0) + struct.pack(">I", len(heads)) + heads
    img = jpx.decode(main + ppm + _tile(hdr, b"", body))
    assert not img.problems and img.planes == want


@needs_opj
def test_subsampled_components_are_widened(tmp_path):
    p = _picture(80, 60, 3)
    src = tmp_path / "s.pgm"
    _write_pnm(src, [p[0]])
    # 成分ごとの間引きは opj_compress の -s（すべての成分に同じ間引き）：1 成分の間引きで大きさの計算を確かめる
    out = tmp_path / "s.j2k"
    r = subprocess.run([OPJ, "-i", str(src), "-o", str(out), "-s", "2,2", "-n", "2"], capture_output=True, timeout=60)
    assert r.returncode == 0, r.stderr
    img = jpx.decode(out.read_bytes())
    assert not img.problems
    ref = _reference(tmp_path, out, 0, 1)
    assert img.planes[0] == ref[2][0]


@needs_opj
def test_sixteen_bit(tmp_path):
    p = _picture(40, 30, bits=16)
    jp = _compress(tmp_path, p, "-n", 2, bits=16)
    img = jpx.decode(jp.read_bytes())
    assert img.planes[0][:40] == bytes(v * 255 // 65535 for v in p[0][0])[:40] or all(
        abs(a - v * 255 / 65535) <= 1 for a, v in zip(img.planes[0], [v for r in p[0] for v in r]))


# ====================================================================== 不可逆（9/7）：浮動小数の丸めの差だけ
@needs_opj
def test_lossy_97_gray_matches_openjpeg(tmp_path):
    _roundtrip(tmp_path, _picture(120, 90), "-I", "-r", "10", "-n", 4, tol=2)


@needs_opj
def test_lossy_97_rgb_ict_matches_openjpeg(tmp_path):
    _roundtrip(tmp_path, _picture(120, 90, 3), "-I", "-r", "12", "-n", 3, tol=3)


@needs_opj
def test_lossless_layers_are_all_read(tmp_path):
    p = _picture(80, 60)
    img = _roundtrip(tmp_path, p, "-r", "60,20,5,1", "-n", 3)
    assert img.planes[0] == bytes(v for r in p[0] for v in r)


# ====================================================================== 低い解像度で開く
@needs_opj
@pytest.mark.parametrize("reduce", [1, 2, 3])
def test_reduced_resolution_matches_openjpeg(reduce, tmp_path):
    img = _roundtrip(tmp_path, _picture(130, 97, 3), "-n", 4, reduce=reduce)
    assert img.reduce == reduce and img.notes


@needs_opj
def test_target_pixels_choose_the_resolution(tmp_path):
    jp = _compress(tmp_path, _picture(200, 160), "-n", 4)
    img = jpx.decode(jp.read_bytes(), target_pixels=200 * 160 // 16)
    assert img.reduce == 2 and (img.w, img.h) == (50, 40)
    assert (img.full_w, img.full_h) == (200, 160)


@needs_opj
def test_budget_forces_a_lower_resolution_and_says_so(tmp_path):
    jp = _compress(tmp_path, _picture(200, 160), "-n", 4)
    img = jpx.decode(jp.read_bytes(), max_t1_bytes=2000)
    assert img.reduce >= 1
    assert any("予算" in n for n in img.notes)


# ====================================================================== JP2 の箱
@needs_opj
def test_jp2_boxes_gray_and_rgb(tmp_path):
    img = _roundtrip(tmp_path, _picture(50, 40), "-n", 2, name="a.jp2")
    assert img.colorspace == "gray"
    img = _roundtrip(tmp_path, _picture(50, 40, 3), "-n", 2, name="b.jp2")
    assert img.colorspace == "sRGB"


def _box(t, b):
    return struct.pack(">I4s", 8 + len(b), t) + b


@needs_opj
def test_palette_and_resolution_boxes(tmp_path):
    # 4 色のパレットの画像：成分は 1 つ（添字）
    r = random.Random(1)
    idx = [[r.randrange(4) for _ in range(30)] for _ in range(20)]
    jp = _compress(tmp_path, [idx], "-n", 1)
    cs = jp.read_bytes()
    pal = [(255, 0, 0), (0, 255, 0), (0, 0, 255), (10, 20, 30)]
    pclr = struct.pack(">HB", 4, 3) + bytes([7, 7, 7]) + b"".join(bytes(c) for c in pal)
    cmap = b"".join(struct.pack(">HBB", 0, 1, i) for i in range(3))
    ihdr = struct.pack(">IIHBBBB", 20, 30, 1, 7, 7, 0, 0)
    res = _box(b"resc", struct.pack(">HHHHbb", 11811, 100, 11811, 100, 2, 2))  # 300 dpi = 11811 画素/m
    colr = bytes([1, 0, 0]) + struct.pack(">I", 16)
    jp2h = _box(b"ihdr", ihdr) + _box(b"colr", colr) + _box(b"pclr", pclr) + _box(b"cmap", cmap) + _box(b"res ", res)
    data = (_box(b"jP  ", b"\r\n\x87\n") + _box(b"ftyp", b"jp2 \0\0\0\0jp2 ") + _box(b"jp2h", jp2h) + _box(b"jp2c", cs))
    img = jpx.decode(data)
    assert len(img.planes) == 3 and img.colorspace == "sRGB"
    for y in range(20):
        for x in range(30):
            c = pal[idx[y][x]]
            assert tuple(p[y * 30 + x] for p in img.planes) == c
    assert img.dpi and round(img.dpi[0]) == 300


# ====================================================================== 表の確かめ（附属書 D）
def test_zero_coding_contexts_examples():
    z = T1.ZC
    # LL・LH：横に 2 つ → 8、HL は縦に 2 つ → 8、HH は斜めに 3 つ → 8
    assert z[0][T1.F_E | T1.F_W] == 8
    assert z[2][T1.F_E | T1.F_W] == 8
    assert z[1][T1.F_N | T1.F_S] == 8
    assert z[1][T1.F_E | T1.F_W] == 4
    assert z[3][T1.F_NE | T1.F_NW | T1.F_SE] == 8
    assert z[3][0] == 0 and z[0][0] == 0 and z[0][T1.F_NE] == 1


def test_sign_contexts_table():
    # 横の寄与 +1・縦の寄与 +1 → 13、反転なし。両方 −1 → 13、反転あり
    i = (T1.F_E | T1.F_N)
    assert T1.SC_CTX[i] == 13 and T1.SC_XOR[i] == 0
    i = (T1.F_E | T1.F_N) | ((T1.S_E | T1.S_N) >> 4)
    assert T1.SC_CTX[i] == 13 and T1.SC_XOR[i] == 1
    assert T1.SC_CTX[0] == 9 and T1.SC_XOR[0] == 0


def test_not_a_jpeg2000():
    with pytest.raises(jpx.JpxError):
        jpx.decode(b"hello world, not an image" * 4)


@needs_opj
def test_truncated_stream_is_opened_as_far_as_possible_and_says_so(tmp_path):
    jp = _compress(tmp_path, _picture(120, 90), "-n", 3, "-r", "20,5,1")
    d = jp.read_bytes()
    img = jpx.decode(d[:len(d) * 2 // 3])
    assert (img.w, img.h) == (120, 90)
    assert img.problems


# ====================================================================== PDF の中の JPEG 2000
def _pdf_with_jpx(data: bytes, w: int, h: int, extra: bytes = b"", more: tuple = ()) -> bytes:
    """1 画素 = 1 ポイントの 1 ページの PDF。"""
    objs = [b"<</Type/Catalog/Pages 2 0 R>>", b"<</Type/Pages/Kids[3 0 R]/Count 1>>",
            f"<</Type/Page/Parent 2 0 R/MediaBox[0 0 {w} {h}]/Contents 4 0 R/Resources<</XObject<</I 5 0 R>>>>>>".encode()]
    cont = f"0 1 0 rg 0 0 {w} {h} re f q {w} 0 0 {h} 0 0 cm /I Do Q".encode()
    objs.append(b"<</Length %d>>stream\n" % len(cont) + cont + b"\nendstream")
    objs.append(b"<</Type/XObject/Subtype/Image/Width %d/Height %d/Filter/JPXDecode%s/Length %d>>stream\n" % (w, h, extra, len(data))
                + data + b"\nendstream")
    objs += list(more)
    out = bytearray(b"%PDF-1.7\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 %d\n0000000000 65535 f \n" % (len(objs) + 1)
    for o in offs:
        out += b"%010d 00000 n \n" % o
    out += b"trailer<</Size %d/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF\n" % (len(objs) + 1, x)
    return bytes(out)


def _render(pdf: bytes, **kw):
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.pageimage import render

    return render(PdfDocument(pdf, "t.pdf"), 1, dpi=72, **kw)


@needs_opj
def test_pdf_page_draws_rgb_without_colorspace_key(tmp_path):
    p = _picture(60, 40, 3)
    jp = _compress(tmp_path, p, "-n", 2, name="c.jp2")
    r = _render(_pdf_with_jpx(jp.read_bytes(), 60, 40))
    assert not r["skipped"] and not r["problems"]
    want = bytearray(60 * 40 * 3)
    for c in range(3):
        want[c::3] = bytes(v for row in p[c] for v in row)
    assert r["rgb"] == bytes(want)


@needs_opj
def test_pdf_colorspace_key_overrides_the_jp2_box(tmp_path):
    p = _picture(40, 30)
    jp = _compress(tmp_path, p, "-n", 2, name="g.jp2")
    r = _render(_pdf_with_jpx(jp.read_bytes(), 40, 30, b"/ColorSpace/DeviceGray/BitsPerComponent 8"))
    assert r["rgb"][0::3] == bytes(v for row in p[0] for v in row)
    # Indexed の色空間：成分の値が添字
    idx = [[(x + y) % 2 for x in range(40)] for y in range(30)]
    jp = _compress(tmp_path, [idx], "-n", 1, name="i.j2k")
    cs = b"/ColorSpace[/Indexed/DeviceRGB 1<ff0000 0000ff>]"
    r = _render(_pdf_with_jpx(jp.read_bytes(), 40, 30, cs))
    assert r["rgb"][:6] == bytes((255, 0, 0, 0, 0, 255))


@needs_opj
def test_pdf_large_image_opens_at_the_resolution_of_the_screen_and_says_so(tmp_path):
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.pageimage import view_page

    p = _picture(1200, 900)
    jp = _compress(tmp_path, p, "-n", 3, "-r", "40")
    # 1 枚の画像を 100 ポイント四方に：画面に要る画素は少ない
    pdf = _pdf_with_jpx(jp.read_bytes(), 100, 75, b"/ColorSpace/DeviceGray")
    v = view_page(PdfDocument(pdf, "t.pdf"), 1, dpi=72)
    assert "解像度の段" in v["text"] and v["not_drawn"] == []


@needs_opj
def test_pdf_smask_in_data_uses_the_alpha_channel(tmp_path):
    # 2 成分（濃さ＋透け）：cdef の箱で 2 つ目を透けとする
    g = [[100 for _ in range(30)] for _ in range(20)]
    a = [[0 if x < 15 else 255 for x in range(30)] for _ in range(20)]
    j1 = _compress(tmp_path, [g], "-n", 1, name="x.j2k").read_bytes()
    j2 = _compress(tmp_path, [a], "-n", 1, name="y.j2k").read_bytes()
    _ = j2
    # OpenJPEG の PGM からは 2 成分を作れないので、RGB の流れに 4 つ目を付けず、透けは別の試験（_finish_image の cdef）で確かめる
    img = jpx._finish_image(jpx._Stream(j1), jpx._JP2(colorspace="gray", cdef=[(0, 0, 1), (1, 1, 0)]),
                            [bytes(v for r in g for v in r), bytes(v for r in a for v in r)], [(0, 0, 30, 20), (0, 0, 30, 20)], 0, [], [], False)
    assert len(img.planes) == 1 and img.alpha == bytes(v for r in a for v in r)


@needs_opj
def test_broken_jpx_stream_is_marked_not_guessed():
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.pageimage import view_page

    v = view_page(PdfDocument(_pdf_with_jpx(b"not a jpeg 2000 stream at all", 40, 30), "t.pdf"), 1)
    assert any("JPEG 2000" in s for s in v["not_drawn"])
