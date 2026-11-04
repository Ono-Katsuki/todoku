"""JBIG2（ITU-T T.88）の復号：試験の中の符号化で作ったデータを開き、元の画素と合うこと。
pdftoppm があれば、poppler の別の実装でも同じ画素に開けることを確かめる（符号化が仕様どおりであることの確かめ）。"""

from __future__ import annotations

import random
import re
import shutil
import subprocess

import jbig2_encoder as E
import pytest

from todoku.imaging import jbig2 as J
from todoku.limits import LimitExceeded

W, H = 96, 64


def _noise_page(seed=1, w=W, h=H):
    """字のような小さな塊と線の、白黒の画像（TPGD が効くように、同じ行も続く）。"""
    r = random.Random(seed)
    rows = [[0] * w for _ in range(h)]
    for _ in range(14):
        x, y, a, b = r.randrange(w), r.randrange(h), r.randrange(2, 16), r.randrange(1, 7)
        for j in range(y, min(h, y + b)):
            for i in range(x, min(w, x + a)):
                rows[j][i] = 1
    for _ in range(40):
        rows[r.randrange(h)][r.randrange(w)] ^= 1
    return rows


def _symbols(n=6, seed=2):
    r = random.Random(seed)
    out = []
    for k in range(n):
        h = 5 + k % 3 * 3
        w = 4 + (k * 3) % 7
        s = [[0] * w for _ in range(h)]
        for j in range(h):
            for i in range(w):
                s[j][i] = 1 if (i in (0, w - 1) or j in (0, h - 1) or r.random() < 0.2) else 0
        out.append(s)
    return E.sort_symbols(out)


def _rows(page: J.Page):
    return [[0 if page.px[y * page.w + x] else 1 for x in range(page.w)] for y in range(page.h)]


def _pdftoppm(pdf: bytes, tmp_path, w: int, h: int):
    """poppler の復号（pdfimages は描かずに画像をそのまま書き出すので、描くときの丸めが入らない）。"""
    exe = shutil.which("pdfimages")
    if not exe:
        return None
    p = tmp_path / "t.pdf"
    p.write_bytes(pdf)
    out = subprocess.run([exe, "-png", str(p), str(tmp_path / "o")], capture_output=True, timeout=60)
    if out.returncode != 0:
        return None  # poppler が開けない形（文脈を引き継ぐ辞書など）：ほかの実装で確かめる
    f = sorted(tmp_path.glob("o-*"))[-1]
    d = f.read_bytes()
    f.unlink()
    if f.suffix == ".pbm":
        parts = d.split(None, 3)
        pw, ph = int(parts[1]), int(parts[2])
        stride = (pw + 7) // 8
        body = d[-stride * ph:]
        return [[(body[y * stride + (x >> 3)] >> (7 - (x & 7))) & 1 for x in range(pw)] for y in range(ph)]
    if f.suffix == ".pgm":
        parts = d.split(None, 4)
        pw, ph = int(parts[1]), int(parts[2])
        body = d[-pw * ph:]
        return [[0 if body[y * pw + x] > 127 else 1 for x in range(pw)] for y in range(ph)]
    if f.suffix == ".jb2e":
        return None  # 復号せずに書き出した
    if f.suffix == ".png":
        import zlib

        i = 8
        idat = b""
        while i < len(d):
            n = int.from_bytes(d[i:i + 4], "big")
            t = d[i + 4:i + 8]
            if t == b"IHDR":
                pw, ph, bits = int.from_bytes(d[i + 8:i + 12], "big"), int.from_bytes(d[i + 12:i + 16], "big"), d[i + 16]
            elif t == b"IDAT":
                idat += d[i + 8:i + 8 + n]
            i += 12 + n
        raw = zlib.decompress(idat)
        assert bits == 1
        stride = (pw + 7) // 8 + 1
        out_rows = []
        prev = bytes(stride - 1)
        for y in range(ph):
            ft, line = raw[y * stride], bytearray(raw[y * stride + 1:(y + 1) * stride])
            assert ft in (0, 2)
            if ft == 2:
                line = bytearray((a + b) & 255 for a, b in zip(line, prev))
            prev = bytes(line)
            out_rows.append([1 - ((line[x >> 3] >> (7 - (x & 7))) & 1) for x in range(pw)])
        return out_rows
    raise AssertionError(f.suffix)


def _gs(pdf: bytes, tmp_path):
    """Ghostscript（jbig2dec）の復号。"""
    exe = shutil.which("gs")
    if not exe:
        return None
    p = tmp_path / "g.pdf"
    p.write_bytes(pdf)
    o = tmp_path / "g.pgm"
    out = subprocess.run([exe, "-q", "-dNOPAUSE", "-dBATCH", "-sDEVICE=pgmraw", "-r72", f"-sOutputFile={o}", str(p)], capture_output=True, timeout=60)
    assert out.returncode == 0, out.stderr
    d = o.read_bytes()
    m = re.match(rb"P5\s+(?:#[^\n]*\n)?(\d+)\s+(\d+)\s+255\s", d)
    pw, ph = int(m[1]), int(m[2])
    body = d[-pw * ph:]
    return [[0 if body[y * pw + x] > 127 else 1 for x in range(pw)] for y in range(ph)]


def _check(stream: bytes, expect, tmp_path, globals_data=None, oracle=True):
    """oracle：別の実装（poppler・Ghostscript）の少なくとも 1 つが、同じ画素に開くこと。入っていなければ飛ばす。"""
    h, w = len(expect), len(expect[0])
    page = J.decode(stream, globals_data, w, h)
    assert not page.problems, page.problems
    assert _rows(page) == expect
    if not oracle:
        return
    pdf = E.pdf_with_jbig2(stream, w, h, globals_data)
    others = [r for r in (_pdftoppm(pdf, tmp_path, w, h), _gs(pdf, tmp_path)) if r is not None]
    if others:
        assert expect in others, "別の実装と合わない（符号化か、復号の見立てが仕様と違う）"


def _page(*segs, w=W, h=H, fill=0):
    return E.segment(0, 48, E.page_info(w, h, fill)) + b"".join(segs)


# ====================================================================== 附属書 B の表
@pytest.mark.parametrize("n", range(1, 16))
def test_standard_tables_are_complete_prefix_codes_and_ranges_do_not_gap(n):
    spec = J._STD_SPEC[n]
    assert sum(2.0 ** -ln[0] for ln in spec) == 1.0  # 符号の空きも重なりもない（B.3）
    normal = [ln for ln in spec if len(ln) == 3 or ln[3] == ""]
    for a, b in zip(normal, normal[1:]):
        assert a[2] + (1 << a[1]) == b[2], (a, b)  # 範囲は続いている
    high = [ln for ln in spec if len(ln) > 3 and ln[3] == "high"]
    if high:
        assert high[0][2] == normal[-1][2] + (1 << normal[-1][1])
    low = [ln for ln in spec if len(ln) > 3 and ln[3] == "low"]
    if low:
        assert low[0][2] == normal[0][2] - 1


def test_custom_table_b2():
    # HTOOB=1・HTPS=3 ビット・HTRS=2 ビット、低 0・高 4 で範囲は [0,1)=長さ 1・[1,2)=長さ 2・[2,4)=長さ 3、低の外なし、高の外 長さ 3、OOB 長さ 3
    bw = E.BitWriter()
    for preflen, rangelen in ((1, 0), (2, 0), (3, 1)):
        bw.bits(preflen, 3)
        bw.bits(rangelen, 2)
    bw.bits(0, 3)  # 下の外
    bw.bits(4, 3)  # 上の外
    bw.bits(4, 3)  # OOB
    data = bytes([1 | (3 - 1) << 1 | (2 - 1) << 4]) + (0).to_bytes(4, "big") + (4).to_bytes(4, "big") + bw.bytes()
    t = J.custom_table(data)
    w = E.BitWriter()
    for v in (0, 1, 2, 3, 100, None):
        E.huff_put(w, t, v)
    br = J.BitReader(w.bytes())
    assert [t.decode(br) for _ in range(6)] == [0, 1, 2, 3, 100, None]


# ====================================================================== 整数（附属書 A）
def test_arith_integers_all_ranges():
    vals = [0, 1, 3, 4, 19, 20, 83, 84, 339, 340, 4435, 4436, 70000, -1, -4, -20, -84, -340, -4436, -70000, None, 5, None]
    m = E.MQEnc()
    cx: dict = {}
    for v in vals:
        E.enc_int(m, cx, v)
    idc: dict = {}
    for v in (0, 5, 31, 17):
        E.enc_iaid(m, idc, 5, v)
    data = m.flush()
    q = J.MQ(data)
    dc = bytearray(512)
    assert [J.decode_int(q, dc) for _ in vals] == vals
    id2 = bytearray(1 << 6)
    assert [J.decode_iaid(q, id2, 5) for _ in range(4)] == [0, 5, 31, 17]


# ====================================================================== 汎用領域
@pytest.mark.parametrize("t", [0, 1, 2, 3])
@pytest.mark.parametrize("tpgd", [False, True])
def test_generic_arith_all_templates(t, tpgd, tmp_path):
    rows = _noise_page(t + 1)
    rows[10] = list(rows[9])  # 同じ行（TPGD）
    rows[11] = list(rows[9])
    _check(_page(E.segment(1, 38, E.generic_seg(rows, t=t, tpgd=tpgd))), rows, tmp_path)


@pytest.mark.parametrize("t,at", [(0, [(5, -2), (-7, -1), (0, -3), (-5, 0)]), (1, [(-4, -2)]), (2, [(-6, 0)]), (3, [(4, -1)])])
def test_generic_arith_moved_at_pixels(t, at, tmp_path):
    rows = _noise_page(5 + t)
    _check(_page(E.segment(1, 38, E.generic_seg(rows, t=t, at=at))), rows, tmp_path)


def test_generic_mmr(tmp_path):
    rows = _noise_page(9)
    _check(_page(E.segment(1, 38, E.generic_seg(rows, mmr=True))), rows, tmp_path)


def test_regions_are_combined_on_the_page(tmp_path):
    a = _noise_page(10, 40, 30)
    b = _noise_page(11, 40, 30)
    for op in (0, 1, 2, 3, 4):
        expect = [[0] * W for _ in range(H)]
        for y in range(30):
            for x in range(40):
                expect[y + 5][x + 7] = a[y][x]
        for y in range(30):
            for x in range(40):
                u, v = expect[y + 5][x + 7], b[y][x]
                expect[y + 5][x + 7] = {0: u | v, 1: u & v, 2: u ^ v, 3: 1 - (u ^ v), 4: v}[op]
        s = _page(E.segment(1, 38, E.generic_seg(a, 7, 5)), E.segment(2, 38, E.generic_seg(b, 7, 5, op=op)))
        _check(s, expect, tmp_path)


def test_page_default_pixel_black(tmp_path):
    a = _noise_page(12, 20, 20)
    expect = [[1] * W for _ in range(H)]
    for y in range(20):
        for x in range(20):
            expect[y][x] = a[y][x]
    _check(_page(E.segment(1, 38, E.generic_seg(a, op=4)), fill=1), expect, tmp_path)


def test_unknown_page_height_with_stripes(tmp_path):
    rows = _noise_page(13)
    head = E.segment(0, 48, E.page_info(W, 0xFFFFFFFF)[:-2] + (0x8010).to_bytes(2, "big"))
    a = E.segment(1, 38, E.generic_seg(rows[:32]))
    st1 = E.segment(2, 50, (31).to_bytes(4, "big"))
    b = E.segment(3, 38, E.generic_seg(rows[32:], y=32))
    st2 = E.segment(4, 50, (63).to_bytes(4, "big"))
    _check(head + a + st1 + b + st2, rows, tmp_path)


def test_segment_of_unknown_length(tmp_path):
    rows = _noise_page(14)
    seg = E.segment(1, 38, E.generic_seg(rows))
    # 長さの欄を 0xFFFFFFFF に、データの後ろに行の数を足す（7.2.7）
    i = seg.index(E.generic_seg(rows)) - 4
    unk = seg[:i] + b"\xff\xff\xff\xff" + seg[i + 4:] + (len(rows)).to_bytes(4, "big")
    page = J.decode(_page(unk), None, W, H)
    assert _rows(page) == rows


def test_file_organisation_header_sequential():
    rows = _noise_page(15)
    body = _page(E.segment(1, 38, E.generic_seg(rows)))
    head = b"\x97JB2\r\n\x1a\n" + bytes([1]) + (1).to_bytes(4, "big")
    assert _rows(J.decode(head + body)) == rows


# ====================================================================== 記号の辞書・字の領域
def _text_expect(w, h, instances, syms, op=0, def_pixel=0):
    out = [[def_pixel] * w for _ in range(h)]
    for inst in instances:
        x, y, sid = inst[:3]
        b = inst[3] if len(inst) > 3 else syms[sid]
        for j, row in enumerate(b):
            for i, v in enumerate(row):
                if 0 <= y + j < h and 0 <= x + i < w:
                    u = out[y + j][x + i]
                    out[y + j][x + i] = {0: u | v, 1: u & v, 2: u ^ v, 3: 1 - (u ^ v)}[op]
    return out


def _place(syms):
    return [(2, 3, 0), (14, 4, 1), (30, 2, 2), (2, 25, 3), (20, 26, 4), (44, 24, 5), (60, 30, 1), (70, 5, 3)]


@pytest.mark.parametrize("t", [0, 1, 2, 3])
def test_symbols_and_text_arith(t, tmp_path):
    syms = _symbols()
    sd, _gb = E.symbol_dict_arith(syms, t=t)
    inst = _place(syms)
    tx = E.text_seg_arith(W, H, inst, syms)
    _check(_page(E.segment(1, 0, sd), E.segment(2, 6, tx, refs=[1])), _text_expect(W, H, inst, syms), tmp_path)


@pytest.mark.parametrize("corner", [0, 1, 2, 3])
@pytest.mark.parametrize("transposed", [False, True])
def test_text_corners_and_transposed(corner, transposed, tmp_path):
    syms = _symbols()
    sd, _gb = E.symbol_dict_arith(syms)
    inst = _place(syms)
    tx = E.text_seg_arith(W, H, inst, syms, corner=corner, transposed=transposed)
    _check(_page(E.segment(1, 0, sd), E.segment(2, 6, tx, refs=[1])), _text_expect(W, H, inst, syms), tmp_path)


@pytest.mark.parametrize("log_strips,ds", [(2, 0), (3, 3), (1, -4)])
def test_text_strips_and_offset(log_strips, ds, tmp_path):
    syms = _symbols()
    sd, _gb = E.symbol_dict_arith(syms)
    inst = _place(syms)
    tx = E.text_seg_arith(W, H, inst, syms, log_strips=log_strips, ds_offset=ds)
    _check(_page(E.segment(1, 0, sd), E.segment(2, 6, tx, refs=[1])), _text_expect(W, H, inst, syms), tmp_path)


@pytest.mark.parametrize("op", [1, 2, 3])
def test_text_combination_op_and_default_pixel(op, tmp_path):
    syms = _symbols()
    sd, _gb = E.symbol_dict_arith(syms)
    inst = [(2, 3, 0), (4, 4, 1), (30, 2, 2)]
    tx = E.text_seg_arith(50, 30, inst, syms, op=op, def_pixel=1)
    exp_tx = _text_expect(50, 30, inst, syms, op, 1)
    expect = [[0] * W for _ in range(H)]
    for y in range(30):
        expect[y][:50] = exp_tx[y]
    _check(_page(E.segment(1, 0, sd), E.segment(2, 6, tx, refs=[1])), expect, tmp_path)


def test_jbig2globals_hold_the_dictionary(tmp_path):
    syms = _symbols()
    sd, _gb = E.symbol_dict_arith(syms)
    glob = E.segment(1, 0, sd, page=0)
    inst = _place(syms)
    tx = E.text_seg_arith(W, H, inst, syms)
    _check(_page(E.segment(2, 6, tx, refs=[1])), _text_expect(W, H, inst, syms), tmp_path, glob)


def test_two_dictionaries_and_export_flags(tmp_path):
    syms = _symbols()
    a, b = syms[:3], E.sort_symbols(syms[3:])
    sd1, _g = E.symbol_dict_arith(a)
    # 2 つ目の辞書：1 つ目の 2 番目だけを引き継ぎ、新しい記号を足して書き出す
    sd2, _g = E.symbol_dict_arith(b, n_in=3, exports=[0, 1, 0] + [1] * len(b))
    all_ = a + [a[1]] + b  # 字の領域の番号の順：1 つ目の書き出し＋2 つ目の書き出し
    inst = [(1, 1, 0), (10, 2, 3), (22, 3, 4), (40, 30, 6), (60, 20, 2)]
    tx = E.text_seg_arith(W, H, inst, all_)
    seg = _page(E.segment(1, 0, sd1), E.segment(2, 0, sd2, refs=[1]), E.segment(3, 6, tx, refs=[1, 2]))
    _check(seg, _text_expect(W, H, inst, all_), tmp_path)


def test_retained_bitmap_contexts(tmp_path):
    syms = _symbols()
    a, b = syms[:3], E.sort_symbols(syms[3:])
    sd1, gb = E.symbol_dict_arith(a, retain=True)
    sd2, _ = E.symbol_dict_arith(b, n_in=3, exports=[1, 1, 1] + [1] * len(b), reuse=gb)
    all_ = a + a + b  # 1 つ目の書き出し＋2 つ目の書き出し（引き継いだ 3 つも書き出す）
    inst = [(2, 3, 0), (14, 4, 4), (30, 2, 2), (2, 25, 6), (20, 26, 7), (44, 24, 8)]
    tx = E.text_seg_arith(W, H, inst, all_)
    seg = _page(E.segment(1, 0, sd1), E.segment(2, 0, sd2, refs=[1]), E.segment(3, 6, tx, refs=[1, 2]))
    # 文脈の引き継ぎ（7.4.2.2）は、手元の poppler・jbig2dec が同じ流れに開けない（引き継がない流れでは合う）ので、仕様の手順と
    # 試験の符号化が一致することだけを確かめる
    _check(seg, _text_expect(W, H, inst, all_), tmp_path, oracle=False)


@pytest.mark.parametrize("mmr", [True, False])
def test_huffman_symbols_and_text(mmr, tmp_path):
    syms = _symbols()
    sd = E.symbol_dict_huffman(syms, J.std_table, mmr=mmr)
    inst = _place(syms)
    tx = E.text_seg_huffman(W, H, inst, syms, J.std_table)
    _check(_page(E.segment(1, 0, sd), E.segment(2, 6, tx, refs=[1])), _text_expect(W, H, inst, syms), tmp_path)


def test_huffman_text_with_strips(tmp_path):
    syms = _symbols()
    sd = E.symbol_dict_huffman(syms, J.std_table)
    inst = _place(syms)
    tx = E.text_seg_huffman(W, H, inst, syms, J.std_table, log_strips=2)
    _check(_page(E.segment(1, 0, sd), E.segment(2, 6, tx, refs=[1])), _text_expect(W, H, inst, syms), tmp_path)


@pytest.mark.parametrize("r_template", [0, 1])
def test_text_refinement(r_template, tmp_path):
    syms = _symbols()
    sd, _gb = E.symbol_dict_arith(syms)
    changed = [list(r) for r in syms[2]]
    changed[1][1] ^= 1
    changed[2][2] ^= 1
    bigger = [list(r) + [1] for r in syms[1]] + [[1] * (len(syms[1][0]) + 1)]
    inst = [(2, 3, 0), (14, 4, 2, changed, 0, 0), (30, 2, 1, bigger, 0, 0), (2, 25, 3), (40, 26, 2, changed, 1, -1)]
    tx = E.text_seg_arith(W, H, inst, syms, refine=True, r_template=r_template)
    _check(_page(E.segment(1, 0, sd), E.segment(2, 6, tx, refs=[1])), _text_expect(W, H, inst, syms), tmp_path)


# ====================================================================== 模様・網点
def _patterns(n=16, pw=4, ph=4):
    r = random.Random(5)
    out = []
    for g in range(n):
        p = [[0] * pw for _ in range(ph)]
        cells = [(x, y) for y in range(ph) for x in range(pw)]
        r.shuffle(cells)
        for x, y in cells[:g * pw * ph // max(1, n - 1)]:
            p[y][x] = 1
        out.append(p)
    return out


def _half_expect(w, h, gray, pats, gx, gy, rx, ry, op=0, def_pixel=0):
    out = [[def_pixel] * w for _ in range(h)]
    for mg, row in enumerate(gray):
        for ng, g in enumerate(row):
            x = (gx + mg * ry + ng * rx) >> 8
            y = (gy + mg * rx - ng * ry) >> 8
            for j, r in enumerate(pats[g]):
                for i, v in enumerate(r):
                    if 0 <= y + j < h and 0 <= x + i < w:
                        u = out[y + j][x + i]
                        out[y + j][x + i] = {0: u | v, 1: u & v, 2: u ^ v, 3: 1 - (u ^ v), 4: v}[op]
    return out


@pytest.mark.parametrize("t,mmr", [(0, False), (1, False), (2, False), (3, False), (0, True)])
def test_pattern_and_halftone(t, mmr, tmp_path):
    pats = _patterns()
    r = random.Random(3)
    gw, gh = W // 4, H // 4
    gray = [[r.randrange(16) for _ in range(gw)] for _ in range(gh)]
    pd = E.pattern_dict(pats, t=t, mmr=mmr)
    ht = E.halftone_seg(W, H, gray, len(pats), 4, 4, t=t, mmr=mmr)
    _check(_page(E.segment(1, 16, pd), E.segment(2, 22, ht, refs=[1])), _half_expect(W, H, gray, pats, 0, 0, 4 << 8, 0), tmp_path)


def test_halftone_angled_grid_with_skip(tmp_path):
    pats = _patterns(8)
    r = random.Random(4)
    gw, gh = 30, 30
    gray = [[r.randrange(8) for _ in range(gw)] for _ in range(gh)]
    gx, gy, rx, ry = -40 << 8, 10 << 8, 3 << 8, 1 << 8
    pd = E.pattern_dict(pats)
    ht = E.halftone_seg(W, H, gray, 8, 4, 4, gx, gy, rx, ry, skip=True)
    # 飛ばした格子の値は 0 と読む
    g2 = [[gray[m][n] if not (((gx + m * ry + n * rx) >> 8) + 4 <= 0 or ((gx + m * ry + n * rx) >> 8) >= W
                              or ((gy + m * rx - n * ry) >> 8) + 4 <= 0 or ((gy + m * rx - n * ry) >> 8) >= H) else 0
           for n in range(gw)] for m in range(gh)]
    _check(_page(E.segment(1, 16, pd), E.segment(2, 22, ht, refs=[1])), _half_expect(W, H, g2, pats, gx, gy, rx, ry), tmp_path)


# ====================================================================== 細かくする領域
@pytest.mark.parametrize("t,tpgr", [(0, False), (1, False), (0, True), (1, True)])
def test_refinement_region_of_the_page(t, tpgr, tmp_path):
    base = _noise_page(21)
    new = [list(r) for r in base]
    r = random.Random(1)
    for _ in range(30):
        new[r.randrange(H)][r.randrange(W)] ^= 1
    s = _page(E.segment(1, 38, E.generic_seg(base)), E.segment(2, 42, E.refinement_seg(new, base, t=t, tpgr=tpgr)))
    # 型 1 の TPGR の SLTP の文脈（図 15：元の画像の中心の画素だけが 1）は、poppler と jbig2dec で実装が揃っていない。
    # 図と pdf.js の値（0x0008：中心の画素）に合わせるので、別の実装では確かめない
    _check(s, new, tmp_path, oracle=not (t == 1 and tpgr))


def test_refinement_of_an_intermediate_region(tmp_path):
    base = _noise_page(22)
    new = [list(r) for r in base]
    new[5][5] ^= 1
    new[20][30] ^= 1
    s = _page(E.segment(1, 36, E.generic_seg(base)), E.segment(2, 42, E.refinement_seg(new, base, t=1), refs=[1]))
    _check(s, new, tmp_path)


# ====================================================================== 上限・壊れたデータ
def test_huge_declared_page_is_refused_not_allocated():
    s = E.segment(0, 48, E.page_info(200000, 200000))
    with pytest.raises(LimitExceeded):
        J.decode(s + E.segment(1, 38, E.generic_seg([[0]])))


def test_broken_segment_is_reported_with_the_rest_drawn():
    rows = _noise_page(30)
    good = E.segment(1, 38, E.generic_seg(rows))
    bad = E.segment(2, 6, b"\x00\x00\x00\x10")  # 短すぎる字の領域
    page = J.decode(_page(good, bad), None, W, H)
    assert _rows(page) == rows
    assert page.problems and "区切り 2" in page.problems[0]


def test_deadline_returns_partial_and_says_so():
    rows = _noise_page(31)
    page = J.decode(_page(E.segment(1, 38, E.generic_seg(rows))), None, W, H, deadline=0.0)
    assert page.stopped and page.problems


# ====================================================================== PDF の中の JBIG2：ページに描く
def _rendered_rows(pdf: bytes):
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.pageimage import render

    r = render(PdfDocument(pdf, "t.pdf"), 1, dpi=72)
    return r, [[0 if r["rgb"][(y * r["w"] + x) * 3] > 127 else 1 for x in range(r["w"])] for y in range(r["h"])]


def test_pdf_page_draws_a_generic_region():
    rows = _noise_page(40)
    r, got = _rendered_rows(E.pdf_with_jbig2(_page(E.segment(1, 38, E.generic_seg(rows))), W, H))
    assert not r["skipped"] and got == rows


def test_pdf_page_uses_jbig2globals_and_view_page_has_no_complaint():
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.pageimage import view_page

    syms = _symbols()
    sd, _gb = E.symbol_dict_arith(syms)
    inst = _place(syms)
    tx = E.text_seg_arith(W, H, inst, syms)
    pdf = E.pdf_with_jbig2(_page(E.segment(2, 6, tx, refs=[1])), W, H, E.segment(1, 0, sd, page=0))
    r, got = _rendered_rows(pdf)
    assert got == _text_expect(W, H, inst, syms)
    v = view_page(PdfDocument(pdf, "t.pdf"), 1)
    assert v["not_drawn"] == [] and "JBIG2" not in v["text"]


def test_pdf_decode_array_inverts_and_image_mask_paints_ones():
    rows = _noise_page(41)
    s = _page(E.segment(1, 38, E.generic_seg(rows)))
    _r, got = _rendered_rows(E.pdf_with_jbig2(s, W, H, decode="/Decode[1 0]"))
    assert got == [[1 - v for v in r] for r in rows]
    # ImageMask：JBIG2 の 1（黒）は、濃さ 0＝塗る。Decode [0 1] の既定では 0 の標本を塗る
    _r, got = _rendered_rows(E.pdf_with_jbig2(s, W, H, image_mask=True))
    assert got == rows


def test_broken_jbig2_stream_is_marked_not_guessed():
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.pageimage import view_page

    pdf = E.pdf_with_jbig2(b"\x00\x01\x02\x03", W, H)
    v = view_page(PdfDocument(pdf, "t.pdf"), 1)
    assert any("JBIG2" in s for s in v["not_drawn"])


# ====================================================================== 記号を細かくする・集める辞書
@pytest.mark.parametrize("r_template", [0, 1])
def test_symbol_dictionary_refinement_and_aggregation(r_template, tmp_path):
    base = _symbols(4)
    sd1, _gb = E.symbol_dict_arith(base)
    changed = [list(r) for r in base[1]]  # 1 つを細かくする（REFAGGNINST = 1）
    changed[1][1] ^= 1
    # 2 つを集めて 1 つに（REFAGGNINST = 2）
    a, b = base[0], base[2]
    agg_h = max(len(a), len(b))
    agg_w = len(a[0]) + len(b[0]) + 1
    new2 = [[0] * agg_w for _ in range(agg_h)]
    for y, r in enumerate(a):
        new2[y][:len(r)] = r
    for y, r in enumerate(b):
        new2[y][len(a[0]) + 1:len(a[0]) + 1 + len(r)] = r
    specs = E.sort_refagg([(changed, [(0, 0, 1)]), (new2, [(0, 0, 0), (len(a[0]) + 1, 0, 2)])])
    sd2 = E.symbol_dict_refagg(specs, base, r_template)
    news = [s[0] for s in specs]
    all_ = base + news
    inst = [(2, 3, 0), (20, 4, len(base)), (40, 5, len(base) + 1), (70, 30, 3)]
    tx = E.text_seg_arith(W, H, inst, all_)
    seg = _page(E.segment(1, 0, sd1), E.segment(2, 0, sd2, refs=[1]), E.segment(3, 6, tx, refs=[1, 2]))
    _check(seg, _text_expect(W, H, inst, all_), tmp_path)
