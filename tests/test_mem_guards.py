"""メモリと時間の守り：宣言された数で場所を取らない・項目の予算・範囲の表・総当たりをしない索引（小さな手作りのデータで、速く）。

範囲の表や索引は、1 つずつ広げた結果・総当たりの結果と同じになることを確かめる（読みの正しさを変えないこと）。
"""

from __future__ import annotations

import os
import random
import struct
import sys
import time
from pathlib import Path

import pytest

from todoku import isolate, limits
from todoku.cfb import MAGIC, CFB, CfbError
from todoku.pdf import doc as pdfdoc
from todoku.pdf import fonts, raster, sfnt
from todoku.pdf.audit import _KeptFrames
from todoku.pdf.layout import _overlap, _RectIndex
from todoku.pdf.pageimage import with_ruler
from todoku.pdf.pagetext import is_overprint
from todoku.pdf.ranges import RangeMap, flatten
from todoku.pdf.render import colors as render_colors
from todoku.pdf.render.fontprog import Type1Outlines as RenderType1
from todoku.pdf.render.image import _raw_rgb

sys.path.insert(0, str(Path(__file__).parent / "audit_findings" / "mem"))


# ---------------------------------------------------------------------- 予算
def test_items_budget_adds_up_and_scales_with_the_file():
    b = limits.Budget(total=1 << 30, items=100)
    b.count(60, "a")
    with pytest.raises(limits.LimitExceeded) as e:
        b.count(60, "b")  # 1 つずつは上限の内でも、合わせて止まる
    assert "b" in str(e.value) and "範囲" in str(e.value)  # 小さく頼み直す言葉がある
    big = limits.Budget(total=1 << 30, items=100)
    big.allow_file(1000)
    big.count(limits.ITEMS_PER_FILE_BYTE * 1000)  # 大きなファイルには広がる


def test_items_budget_outside_a_call_only_checks_each_request():
    loose = limits.Budget()
    for _ in range(3):
        loose.count(limits.MAX_ITEMS)  # 合計は数えない（呼び出しの外・長い仕事）
    with pytest.raises(limits.LimitExceeded):
        loose.count(limits.MAX_ITEMS + 1)


def test_inflate_total_leaves_room_below_the_memory_limit():
    assert limits.MAX_INFLATE_TOTAL < limits.MEMORY_LIMIT
    with pytest.raises(limits.LimitExceeded):
        limits.Budget().alloc(limits.MEMORY_LIMIT)


# ---------------------------------------------------------------------- PDF の宣言
def test_predictor_row_longer_than_data_takes_no_room():
    t = time.perf_counter()
    assert pdfdoc._predict(b"\x00abc", {"Predictor": 12, "Columns": 40_000_000_000}) == b""
    assert pdfdoc._predict(b"\x00abc", {"Predictor": 12, "Columns": 3}) == b"abc"
    assert pdfdoc._predict(b"abcd", {"Predictor": 2, "Columns": 0}) == b"abcd"
    assert time.perf_counter() - t < 1


def _xref_stream_pdf(index: bytes) -> bytes:
    obj = b"1 0 obj\n<< /Type /Catalog >>\nendobj\n"
    head = b"%PDF-1.7\n"
    pos = len(head) + len(obj)
    rows = struct.pack(">BHB", 0, 0, 0) + struct.pack(">BHB", 1, len(head), 0) + struct.pack(">BHB", 1, pos, 0)
    x = (b"2 0 obj\n<< /Type /XRef /Size 3 /W [1 2 1] /Index " + index + b" /Root 1 0 R /Length %d >>\nstream\n" % len(rows)
         + rows + b"\nendstream\nendobj\n")
    return head + obj + x + b"startxref\n%d\n%%%%EOF\n" % pos


def test_xref_stream_index_is_bounded_by_the_rows_that_exist():
    t = time.perf_counter()
    d = pdfdoc.PdfDocument(_xref_stream_pdf(b"[0 2000000000]"))
    assert time.perf_counter() - t < 1
    assert len(d.xref) == 3 and not d.repaired
    assert d.root.get("Type").name == "Catalog"


def test_cfb_mini_cutoff_must_be_4096():
    blob = bytearray(1024)
    blob[:8] = MAGIC
    struct.pack_into("<HHHHH", blob, 0x18, 0x3E, 3, 0xFFFE, 9, 6)
    struct.pack_into("<I", blob, 0x38, 0x7FFFFFFF)
    with pytest.raises(CfbError, match="4096"):
        CFB(bytes(blob))


def test_declared_image_size_draws_only_what_exists():
    d = {"BitsPerComponent": 8}
    t = time.perf_counter()
    im = _raw_rgb(b"\x80" * 100, d, 4_294_967_295, 4_294_967_295, render_colors.GRAY, None, 1000)
    assert time.perf_counter() - t < 2
    assert len(im.px) == im.w * im.h * 3 <= 3000
    # ふつうの画像は 1 画素も変わらない（足りない行は 0 の標本の色）
    im = _raw_rgb(bytes([10, 20, 30, 40]), d, 2, 3, render_colors.GRAY, None, 1000)
    assert im.px == bytes(v for g in (10, 20, 30, 40, 0, 0) for v in (g, g, g))
    rgb = _raw_rgb(bytes(range(12)), d, 2, 2, render_colors.colorspace(_Doc(), raster_name("DeviceRGB"), None), None, 1000)
    assert rgb.px == bytes(range(12))


class _Name:
    def __init__(self, name):
        self.name = name


def test_declared_gray_size_is_bounded_too():
    d = {"BitsPerComponent": 8}
    g = raster._raw_gray(_Doc(), b"\x80" * 10, {**d, "ColorSpace": raster_name("DeviceGray")}, 4_294_967_295, 3, 300)
    assert len(g.px) == g.w * g.h <= 300


def raster_name(n):
    from todoku.pdf.objects import Name

    return Name(n)


def test_type1_subrs_declaration_takes_no_room():
    priv = b"/Subrs 4000000000 array\ndup 0 3 RD abc NP\n/CharStrings 1 dict dup begin\n/a 1 RD x ND\nend\n"
    f = RenderType1.__new__(RenderType1)
    f.leniv = -1
    f.subrs = {}
    f.charstrings = {}
    f._parse(priv)
    assert f.subrs == {0: b"abc"} and "a" in f.charstrings


def test_ruler_ticks_are_limited_by_pixels_not_page_size():
    r = {"w": 40, "h": 40, "scale": 40 / 352_777_778, "origin_mm": (0, 0), "rgb": bytes(40 * 40 * 3)}
    t = time.perf_counter()
    with_ruler(r)
    assert time.perf_counter() - t < 1


# ---------------------------------------------------------------------- 範囲の表（広げた結果と同じ）
def _expanded(entries):
    out = {}
    for lo, hi, v in entries:
        for k in range(lo, hi + 1):
            out[k] = v
    return out


@pytest.mark.parametrize("seed", range(8))
def test_range_map_equals_expanding_one_by_one(seed):
    rnd = random.Random(seed)
    entries = []
    for i in range(rnd.randint(0, 40)):
        lo = rnd.randint(0, 200)
        entries.append((lo, lo + rnd.randint(-2, 30), f"v{i}"))
    want = _expanded(entries)
    m = RangeMap(entries)
    assert dict(m.items()) == want
    for k in range(-1, 240):
        assert m.get(k) == want.get(k)
        assert (k in m) == (k in want)
    assert all(a[1] < b[0] for a, b in zip(flatten(entries), flatten(entries)[1:]))


def test_range_map_counts_what_it_expands():
    m = RangeMap([(0, 10_000_000, 1)])
    with limits.call_budget(limits.Budget(total=1 << 30, items=1000)):
        assert m.get(5) == 1  # 引くだけなら広げない
        with pytest.raises(limits.LimitExceeded):
            list(m)


def _cmap12(groups) -> bytes:
    body = struct.pack(">HHIII", 12, 0, 16 + 12 * len(groups), 0, len(groups)) + b"".join(struct.pack(">III", *g) for g in groups)
    return struct.pack(">HH", 0, 1) + struct.pack(">HHI", 3, 10, 12) + body


def _font(tables: dict[str, bytes]) -> bytes:
    pos = 12 + 16 * len(tables)
    recs, body = b"", b""
    for tag, d in sorted(tables.items()):
        recs += struct.pack(">4sIII", tag.encode(), 0, pos + len(body), len(d))
        body += d
    return struct.pack(">IHHHH", 0x00010000, len(tables), 16, 0, 0) + recs + body


def test_cmap12_many_wide_groups_stay_small_and_match():
    groups = [(0x41, 0x60, 5), (0x50, 0x70, 100), (0x4E00, 0x4E10, 200), (0xD7FE, 0xE001, 300)]
    subs = sfnt.cmap_subtables(_font({"cmap": _cmap12(groups)}))[(3, 10)]
    want = {}
    for s, e, g0 in groups:
        for c in range(s, e + 1):
            want[c] = g0 + c - s
    assert dict(subs.items()) == want
    by_glyph = sfnt.unicode_by_glyph(_font({"cmap": _cmap12(groups)}))
    expect = {}
    for code, g in sorted(want.items()):
        if g not in expect and not 0xD800 <= code <= 0xDFFF:
            expect[g] = chr(code)
    assert by_glyph == expect
    # 12 バイトの区切りが 110 万字：1500 個でも、すぐ終わり、広げない
    t = time.perf_counter()
    huge = _font({"cmap": _cmap12([(i, 0x10FFFF, 1) for i in range(1500)])})
    assert sfnt.unicode_by_glyph(huge)
    assert time.perf_counter() - t < 5


def test_cmap4_matches_the_segments():
    seg = 3
    ends, starts, deltas, ro = [0x45, 0x62, 0xFFFF], [0x41, 0x61, 0xFFFF], [-0x40, 0, 1], [0, 4, 0]
    glyph_ids = [7, 0]
    body = struct.pack(">HHHHHHH", 4, 0, 0, seg * 2, 0, 0, 0) + struct.pack(f">{seg}H", *ends) + b"\0\0" + struct.pack(f">{seg}H", *starts)
    body += struct.pack(f">{seg}h", *deltas) + struct.pack(f">{seg}H", *ro) + struct.pack(">2H", *glyph_ids)
    m = sfnt._subtable(body, 0)
    assert dict(m.items()) == {0x41: 1, 0x42: 2, 0x43: 3, 0x44: 4, 0x45: 5, 0x61: 7}


def test_gsub_coverage_ignores_overlapping_ranges():
    cov = struct.pack(">HH", 2, 3) + struct.pack(">HHH", 5, 8, 0) + struct.pack(">HHH", 6, 10, 4) + struct.pack(">HHH", 0, 0xFFFF, 0)
    with limits.call_budget(limits.Budget(total=1 << 30, items=1 << 20)) as b:
        assert sfnt._coverage(cov, 0) == list(range(5, 0x10000))
        assert b.items < 0x10000


class _Doc:
    def resolve(self, o):
        return o


def test_cid_widths_are_range_tables_with_last_entry_winning():
    w = [1, 100_000_000, 500, 5, [600, 700], 3, 4_294_967_295, 800]
    t = fonts._cid_widths(_Doc(), w)
    assert t.get(1) == 500 and t.get(2) == 500
    assert t.get(3) == 800 and t.get(5) == 800 and t.get(4_000_000_000) == 800  # 後の宣言が勝つ
    assert t.get(0, 42) == 42
    t2 = fonts._cid_widths2(_Doc(), [0, 4_294_967_295, 1, 2, 3, 7, [4, 5, 6]])
    assert t2.get(99) == (1.0, 2.0, 3.0) and t2.get(7) == (4.0, 5.0, 6.0)


def test_cmap_ranges_match_expanding_and_stay_small():
    cm = fonts.CMap.parse(b"1 begincodespacerange <0000> <FFFF> endcodespacerange "
                          b"2 beginbfchar <0041> <0061> <0042> <0062> endbfchar "
                          b"2 beginbfrange <0042> <0044> <0070> <0100> <FFFF> <3042> endbfrange "
                          b"1 begincidrange <0000> <FFFF> 10 endcidrange 1 begincidrange <0010> <0020> 500 endcidrange "
                          b"1 begincidchar <0011> 7 endcidchar")
    assert cm.to_unicode(b"\x00\x41") == "a"
    assert cm.to_unicode(b"\x00\x42") == "p"  # 短い範囲は bfchar より後に書かれたものとして勝つ（これまでと同じ）
    assert cm.to_unicode(b"\x00\x44") == "r"
    assert cm.to_unicode(b"\x01\x01") == chr(0x3043)
    assert cm.to_unicode(b"\x41") is None  # 長さが違う
    assert cm.to_cid(b"\x00\x05") == 15 and cm.to_cid(b"\x00\x12") == 502 and cm.to_cid(b"\x00\x11") == 7
    t = time.perf_counter()
    many = fonts.CMap.parse(b"1 begincidrange " + b"".join(b"<%04X> <%04X> %d\n" % (i, i, i) for i in range(1, 65000)) * 2 + b" endcidrange")
    for _ in range(20000):
        assert many.to_cid(b"\xff\xfe") is None
    assert time.perf_counter() - t < 5


def test_codespace_split_is_not_ranges_times_chars():
    cs = b"".join(b"<%04X> <%04X>\n" % (0x8100 + i, 0x8100 + i) for i in range(5000)) + b"<00> <7F>"
    cm = fonts.CMap.parse(b"5001 begincodespacerange " + cs + b" endcodespacerange")
    t = time.perf_counter()
    got = cm.split(b"\x41\x81\x00" * 50000)
    assert time.perf_counter() - t < 3
    assert got[:2] == [b"\x41", b"\x81\x00"] and len(got) == 100000


# ---------------------------------------------------------------------- 総当たりをしない索引
def test_kept_frames_finds_every_overprint_candidate():
    rnd = random.Random(1)
    kept = []
    idx = _KeptFrames()
    for i in range(1500):
        w = rnd.choice([0.5, 3.0, 6.0, 12.0, 90.0])
        size = rnd.choice([1.0, 6.0, 12.0, 13.0, 40.0])
        ang = rnd.choice([(1.0, 0.0), (0.0, 1.0), (0.6, 0.8)])
        fr = (rnd.uniform(0, 300), rnd.uniform(0, 300), ang[0], ang[1], w)
        direct = any(is_overprint(f0, fr, s0, size) for f0, s0 in kept)
        via = any(is_overprint(f0, fr, s0, size) for f0, s0 in idx.near("a", fr, size))
        assert direct == via, i
        if not direct:
            kept.append((fr, size))
            idx.add("a", fr, size)


class _S:
    def __init__(self, rect):
        self.rect = rect


def test_rect_index_keeps_order_and_finds_all_overlaps():
    rnd = random.Random(2)
    kept: list[_S] = []
    idx = _RectIndex()
    for _ in range(800):
        x, y = rnd.uniform(0, 500), rnd.uniform(0, 700)
        w, h = rnd.choice([3, 10, 40, 300, 900]), rnd.choice([1, 10, 40, 900])
        s = _S((x, y, x + w, y + h))
        direct = next((k for k in kept if _overlap(s.rect, k.rect) > 0.6), None)
        via = next((k for k in idx.touching(s.rect) if _overlap(s.rect, k.rect) > 0.6), None)
        assert direct is via
        if direct is None:
            kept.append(s)
            idx.add(s)


# ---------------------------------------------------------------------- 見張り
def test_rss_is_read_without_starting_a_process():
    assert isolate.rss_bytes() > 1 << 20


def test_soft_watchdog_stops_python_with_a_smaller_request(monkeypatch):
    monkeypatch.setattr(limits, "TIME_LIMIT", 0.2)

    def spin(name, args):
        while True:
            pass
    t = time.perf_counter()
    r = isolate.soft_call(spin, "read_window", {"path": "a.pdf"})
    assert time.perf_counter() - t < 5
    assert r["isError"] and r["structuredContent"]["stopped"]
    assert "pages" in r["structuredContent"]["smaller_request"]


def test_soft_watchdog_leaves_quick_calls_alone():
    assert isolate.soft_call(lambda n, a: {"ok": n}, "x", {}) == {"ok": "x"}


def test_crafted_small_files_finish_quickly(tmp_path):
    """監査の手作りの書類のうち、宣言だけが大きいもの：予算の内で、すぐ終わる。"""
    os.environ.setdefault("CRAFT_ONLY", "1")
    import crafted

    from todoku import api

    for name in ("pdf_predictor_columns", "pdf_xrefstm_index", "pdf_type1_subrs", "pdf_cmap12_many", "pdf_gsub_coverage_ranges",
                 "pdf_w_many", "pdf_w2_many", "pdf_xfa_occur", "pdf_xfa_occur_nested", "pdf_xfa_transparent_occur"):
        p = tmp_path / f"{name}.pdf"
        p.write_bytes(crafted.CASES[name][0]())
        t = time.perf_counter()
        api.summary(str(p))
        assert time.perf_counter() - t < 30, name


def test_xfa_occur_beyond_data_is_stated_as_a_fact(tmp_path):
    os.environ.setdefault("CRAFT_ONLY", "1")
    import crafted

    from todoku import api

    p = tmp_path / "x.pdf"
    p.write_bytes(crafted.CASES["pdf_xfa_occur"][0]())
    _k, _d, parts = api.window(str(p))
    text = "\n".join(t for _n, t in parts)
    assert "2,000,000,000 回と宣言" in text


def test_comb_maxlen_declaration_takes_no_room(tmp_path):
    os.environ.setdefault("CRAFT_ONLY", "1")
    import crafted

    from todoku import api

    p = tmp_path / "c.pdf"
    p.write_bytes(crafted.CASES["pdf_comb_maxlen"][0]())
    fl = api.fields(str(p))
    r = api.fill(str(p), {fl[0]["id"]: "abc"}, str(tmp_path / "o.pdf"), overwrite=True)
    assert r.get("proof_ok") is not False
