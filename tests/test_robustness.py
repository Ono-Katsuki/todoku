"""壊れた・悪意のある書類への上限（limits・isolate）：小さな入力と、小さくした上限で確かめる（GB の大きさは作らない）。"""

from __future__ import annotations

import io
import json
import os
import struct
import sys
import time
import zipfile
import zlib
from pathlib import Path

import pytest

from todoku import api, limits
from todoku.limits import Budget, LimitExceeded, call_budget

sys.path.insert(0, str(Path(__file__).parent))


# ---------------------------------------------------------------------- 部品
def _pdf(objs: dict[int, bytes], root: int = 1) -> bytes:
    out = bytearray(b"%PDF-1.7\n%\xe2\xe3\xcf\xd3\n")
    offs = {}
    for n in sorted(objs):
        offs[n] = len(out)
        out += b"%d 0 obj\n" % n + objs[n] + b"\nendobj\n"
    x = len(out)
    size = max(objs) + 1
    out += b"xref\n0 %d\n0000000000 65535 f \n" % size
    for n in range(1, size):
        out += (b"%010d 00000 n \n" % offs[n]) if n in offs else b"0000000000 00000 f \n"
    out += b"trailer\n<< /Size %d /Root %d 0 R >>\nstartxref\n%d\n%%%%EOF\n" % (size, root, x)
    return bytes(out)


def _stream(d: bytes, data: bytes) -> bytes:
    return b"<< " + d + b" /Length %d >>\nstream\n" % len(data) + data + b"\nendstream"


def _page_objs(content: bytes = b"BT /F1 12 Tf 72 720 Td (Hello) Tj ET", extra: bytes = b"") -> dict[int, bytes]:
    return {
        1: b"<< /Type /Catalog /Pages 2 0 R >>",
        2: b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
        3: b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 5 0 R >> " + extra
           + b" >> /Contents 4 0 R >>",
        4: _stream(b"", content),
        5: b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    }


def _fanout_pdf() -> bytes:
    """フォームの部品が自分を 10 回描く（自分を描く所は描かないので、すぐ終わる）。"""
    objs = _page_objs(b"/X1 Do", b"/XObject << /X1 6 0 R >>")
    objs[6] = _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /X1 6 0 R >> >>",
                      b"/X1 Do " * 10 + b"BT /F1 12 Tf (a) Tj ET")
    return _pdf(objs)


def _chain_pdf(levels: int = 8, fan: int = 10) -> bytes:
    """別々のフォームの部品が、次の部品を fan 回ずつ描く（輪ではないが、fan ** levels 回になる）。"""
    objs = _page_objs(b"/X0 Do", b"/XObject << /X0 10 0 R >>")
    for i in range(levels):
        objs[10 + i] = _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 10 10] /Resources << /XObject << /X%d %d 0 R >> >>"
                               % (i + 1, 11 + i), (b"/X%d Do " % (i + 1)) * fan)
    objs[10 + levels] = _stream(b"/Type /XObject /Subtype /Form /BBox [0 0 10 10]", b"0 0 m 1 1 l S")
    return _pdf(objs)


def _raw_zip(members: list[tuple[str, bytes, int, int, int]], local_names: list[str] | None = None) -> bytes:
    out, cds = bytearray(), bytearray()
    for i, (name, data, crc, usize, method) in enumerate(members):
        nb = name.encode()
        lb = (local_names[i] if local_names else name).encode()
        off = len(out)
        out += struct.pack("<4sHHHHHIIIHH", b"PK\x03\x04", 20, 0x800, method, 0, 0, crc, len(data), usize, len(lb), 0) + lb + data
        cds += struct.pack("<4sHHHHHHIIIHHHHHII", b"PK\x01\x02", 20, 20, 0x800, method, 0, 0, crc, len(data), usize,
                           len(nb), 0, 0, 0, 0, 0, off) + nb
    cd = len(out)
    out += cds + struct.pack("<4sHHHHIIH", b"PK\x05\x06", 0, 0, len(members), len(members), len(cds), cd, 0)
    return bytes(out)


def _deflated(data: bytes) -> tuple[bytes, int]:
    c = zlib.compressobj(9, zlib.DEFLATED, -15)
    return c.compress(data) + c.flush(), zlib.crc32(data) & 0xFFFFFFFF


# ---------------------------------------------------------------------- 展開の予算
def test_inflate_counts_actual_bytes_per_part_and_per_call():
    blob = zlib.compress(b"\0" * 300_000)
    with call_budget(Budget(part=100_000, total=None)):
        with pytest.raises(LimitExceeded) as e:
            limits.inflate(blob)
    assert limits.STOPPED in str(e.value) and not e.value.total
    with call_budget(Budget(part=10**9, total=500_000)) as b:
        limits.inflate(blob)  # 300,000 バイト：合計の内
        with pytest.raises(LimitExceeded) as e:
            limits.inflate(blob)  # 合わせて 600,000：合計を超える
        assert e.value.total and b.inflated == 300_000
    # 途中で終わるデータ：tolerant なら読めた所まで、そうでなければ zlib.error
    cut = zlib.compress(b"abc" * 1000)[:-8]
    assert limits.inflate(cut, tolerant=True).startswith(b"abcabc")
    with pytest.raises(zlib.error):
        limits.inflate(cut)


def test_zip_member_with_lying_size_is_stopped_by_actual_bytes(monkeypatch):
    from todoku import ziparc

    monkeypatch.setattr(ziparc, "MAX_RATIO", 10)
    data, crc = _deflated(b"a,b\n" + b"1,2\n" * 2_000_000)  # 8 MB に展開、中央ディレクトリは 1,000 バイトと偽る
    blob = _raw_zip([("x.csv", data, crc, 1000, 8)])
    with pytest.raises((ziparc.ArchiveError, LimitExceeded)) as e:
        ziparc.extract(blob, "x.csv")
    assert "倍" in str(e.value) or limits.STOPPED in str(e.value)
    # OOXML の部品（zipraw）も、宣言でなく出てきたバイトで止める
    from todoku.zipraw import RawZip

    with call_budget(Budget(part=1_000_000)), pytest.raises(LimitExceeded):
        RawZip(blob).read("x.csv")


@pytest.mark.parametrize("case", ["duplicate", "case", "local_name", "overlap"])
def test_zip_with_two_readings_is_refused_and_graded_f(case, tmp_path):
    from todoku import ziparc
    from todoku.grade import grade_blob
    from todoku.zipraw import AmbiguousZip, RawZip

    a, ca = _deflated(b"name,amount\nyamada,\n")
    b, cb = _deflated(b"name,amount\nyamada,1000000\n")
    if case == "duplicate":
        blob = _raw_zip([("x.csv", a, ca, 20, 8), ("x.csv", b, cb, 27, 8)])
    elif case == "case":
        blob = _raw_zip([("x.csv", a, ca, 20, 8), ("X.CSV", b, cb, 27, 8)])
    elif case == "local_name":
        blob = _raw_zip([("x.csv", a, ca, 20, 8)], local_names=["y.csv"])
    else:  # 2 つの記録が同じバイトを指す
        one = _raw_zip([("x.csv", a, ca, 20, 8)])
        cd = one.index(b"PK\x01\x02")
        rec = one[cd:one.index(b"PK\x05\x06")]
        rec2 = rec.replace(b"x.csv", b"y.csv")
        blob = one[:cd] + rec + rec2 + struct.pack("<4sHHHHIIH", b"PK\x05\x06", 0, 0, 2, 2, len(rec) * 2, cd, 0)
    with pytest.raises(AmbiguousZip):
        RawZip(blob)
    with pytest.raises(ziparc.ArchiveError, match="二通り"):
        ziparc.members(blob)
    g = grade_blob(blob, "a.zip")
    assert g.grade == "F" and g.write == "D" and g.reasons[0].title == "ZIP の中身が二通りに読める"
    p = tmp_path / "a.zip"
    p.write_bytes(blob)
    with pytest.raises(ValueError, match="二通り"):
        api.window(str(p))


def test_zip_window_reads_the_central_directory_once():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
        for i in range(3000):
            z.writestr(f"d/{i}.csv", "a,b\n1,2\n")
    from todoku import ziparc

    t = time.time()
    doc, parts = ziparc.window(buf.getvalue(), "many.zip", api.KIND_JA)
    assert len(doc["members"]) == 3000 and time.time() - t < 10


def test_pdf_flate_bomb_is_stopped():
    from todoku.pdf.doc import PdfDocument

    objs = _page_objs()
    objs[4] = _stream(b"/Filter /FlateDecode", zlib.compress(b" " * 3_000_000))
    doc = PdfDocument(_pdf(objs))
    with call_budget(Budget(part=1_000_000)), pytest.raises(LimitExceeded):
        doc.page_content(doc.pages()[0])
    assert len(doc.page_content(doc.pages()[0])) == 3_000_000  # 上限の内なら、そのまま読む


def test_lh5_declared_size_is_capped():
    from todoku.jtd.lha import lh5, unlh5

    packed = lh5(b"x" * 5000)
    assert unlh5(packed, 5000) == b"x" * 5000
    with call_budget(Budget(part=1000)), pytest.raises(LimitExceeded):
        unlh5(packed, 5000)


# ---------------------------------------------------------------------- 輪・枝分かれ
def test_pdf_page_tree_cycle_and_xref_cycles():
    from todoku.pdf.doc import PdfDocument

    objs = _page_objs()
    objs[2] = b"<< /Type /Pages /Kids [2 0 R 2 0 R 3 0 R] /Count 1 >>"
    assert len(PdfDocument(_pdf(objs)).pages()) == 1
    b = _pdf(_page_objs())
    x = int(b.rsplit(b"startxref\n", 1)[1].split(b"\n")[0])
    b = b.replace(b"/Root 1 0 R ", b"/Root 1 0 R /Prev %d /XRefStm %d " % (x, x))
    assert len(PdfDocument(b).pages()) == 1


def test_pdf_object_stream_inside_itself_is_a_clean_error():
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.objects import PdfError

    body = b"<< /Type /Catalog /Pages 2 0 R >>"
    data = b"1 0 6 0 " + body + b" << /Type /ObjStm >>"
    out = bytearray(b"%PDF-1.7\n")
    offs = {}
    pages = _page_objs()
    for n in (2, 3, 4, 5):
        offs[n] = len(out)
        out += b"%d 0 obj\n" % n + pages[n] + b"\nendobj\n"
    offs[6] = len(out)
    out += b"6 0 obj\n" + _stream(b"/Type /ObjStm /N 2 /First 8", data) + b"\nendobj\n"
    entries = {1: (2, 6, 0), 6: (2, 6, 1), **{n: (1, offs[n], 0) for n in (2, 3, 4, 5)}}
    offs[7] = len(out)
    entries[7] = (1, offs[7], 0)
    rows = b"".join(struct.pack(">BIH", *entries.get(n, (0, 0, 0))) for n in range(8))
    out += b"7 0 obj\n" + _stream(b"/Type /XRef /Size 8 /W [1 4 2] /Root 1 0 R", rows) + b"\nendobj\n"
    out += b"startxref\n%d\n%%%%EOF\n" % offs[7]
    doc = PdfDocument(bytes(out))
    with pytest.raises(PdfError, match="自分自身"):
        doc.get(6)


def test_form_xobject_drawing_itself_is_drawn_once():
    from todoku.pdf.content import page_content
    from todoku.pdf.doc import PdfDocument

    doc = PdfDocument(_fanout_pdf())
    t = time.time()
    pc = page_content(doc, doc.pages()[0])
    assert time.time() - t < 5
    assert any("自分自身" in p for p in pc.problems)


def test_form_xobject_fanout_is_stopped_by_the_operator_budget(monkeypatch):
    from todoku.pdf.audit import raw_chars
    from todoku.pdf.content import page_content
    from todoku.pdf.doc import PdfDocument

    doc = PdfDocument(_chain_pdf())  # 10 ** 8 回
    monkeypatch.setattr(limits, "MAX_PAGE_OPS", 50_000)
    with pytest.raises(LimitExceeded):
        page_content(doc, doc.pages()[0])
    with call_budget(Budget(ops=50_000)), pytest.raises(LimitExceeded):
        raw_chars(doc, doc.pages()[0])  # 別の読み方（検査）も、同じ予算で止まる


def test_cfb_directory_cycle_and_minifat_loop():
    from todoku.cfb import CFB, build_cfb

    b = bytearray(build_cfb({"Workbook": b"\x09\x08" + b"\0" * 5000}))
    i = b.index("Root Entry".encode("utf-16-le"))
    struct.pack_into("<III", b, i + 128 + 68, 1, 1, 0xFFFFFFFF)  # 記録 1 の左と右が自分
    t = time.time()
    c = CFB(bytes(b))
    assert "Workbook" in c.streams() and time.time() - t < 2

    b = bytearray(build_cfb({"WordDocument": b"\xec\xa5" + b"\0" * 100, "1Table": b"\0" * 100}))
    c = CFB(bytes(b))
    i = b.index("Root Entry".encode("utf-16-le"))
    struct.pack_into("<Q", b, i + 120, 0)  # ミニストリームを空に
    mf = c.chain(c.minifat_start)[0]
    struct.pack_into("<I", b, (mf + 1) * c.ssize, 0)  # ミニ FAT[0] = 0（自分）
    for e in c.entries:
        if e.name in ("WordDocument", "1Table"):
            struct.pack_into("<I", b, b.index(e.name.encode("utf-16-le")) + 116, 0)
    from todoku.cfb import CfbError

    c2 = CFB(bytes(b))
    t = time.time()
    try:  # 輪：空のミニストリームで止まるか CfbError（以前は終わらなかった）
        assert len(c2.read("WordDocument")) <= 102
    except CfbError:
        pass
    assert time.time() - t < 2


# ---------------------------------------------------------------------- 画像・字句・見出し
def test_raster_huge_dimensions_are_downsampled_or_refused():
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.objects import Name, Stream
    from todoku.pdf.raster import RasterError, ccitt, image_gray

    doc = PdfDocument(_pdf(_page_objs()))
    st = Stream({"Width": 100000, "Height": 100000, "ColorSpace": Name("DeviceGray"), "BitsPerComponent": 8}, b"\0" * 10)
    t = time.time()
    g = image_gray(doc, st, max_pixels=10_000)
    assert g.w * g.h <= 10_000 and time.time() - t < 5
    with pytest.raises(RasterError, match="大きすぎ"):
        ccitt(b"\x00\x10\x01", {"K": 0, "Columns": 100000, "Rows": 100000})


def test_odf_tokenizer_fails_fast_on_unterminated_tags():
    from todoku.odf.xml import XmlError, parse

    head = (b'<?xml version="1.0"?><office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0">'
            b"<office:body><office:text>")
    for tail in (b"<" * 20000, b"<!--" * 20000):
        t = time.time()
        with pytest.raises(XmlError, match="閉じていない"):
            parse(head + tail)
        assert time.time() - t < 2
    assert parse(head + b"<!-- c --></office:text></office:body></office:document>")  # ふつうの注釈は読む


def test_long_paragraph_labels_are_bounded(tmp_path):
    from todoku.docx.window import LABEL_WINDOW

    p = tmp_path / "long.txt"
    p.write_text("氏名：" + "＿ " * 20000, encoding="utf-8")
    t = time.time()
    fl = api.fields(str(p))
    assert fl and time.time() - t < 30
    assert max(len(f.get("label") or "") for f in fl) <= 2 * LABEL_WINDOW + 10
    # 短い段落の見出しは、いままでどおり段落まるごと
    q = tmp_path / "short.txt"
    q.write_text("氏名：＿＿＿＿　住所：＿＿＿＿\n", encoding="utf-8")
    assert any("住所" in (f.get("label") or "") for f in api.fields(str(q)))


# ---------------------------------------------------------------------- マクロ
def _vba(src: str):
    from test_vbarun import book

    from todoku.vbarun import Unsupported, run_macro

    try:
        return run_macro(book(src), "Test")
    except Unsupported as e:
        return e


def test_vba_string_and_range_budgets(monkeypatch):
    from todoku.vbarun import Unsupported

    assert isinstance(_vba("Sub Test()\n Dim s\n s = Space(2000000000)\nEnd Sub\n"), Unsupported)
    monkeypatch.setattr(limits, "MAX_VBA_STRING", 1 << 16)
    t = time.time()
    r = _vba("Sub Test()\n Dim s, i\n s = \"x\"\n For i = 1 To 40\n  s = s & s\n Next\nEnd Sub\n")
    assert isinstance(r, Unsupported) and time.time() - t < 5
    assert isinstance(_vba("Sub Test()\n Range(\"A1:XFD1048576\").Value = 1\nEnd Sub\n"), Unsupported)
    assert isinstance(_vba("Sub Test()\n Dim a()\n ReDim a(1 To 100000, 1 To 100000)\nEnd Sub\n"), Unsupported)
    ok = _vba("Sub Test()\n Range(\"A1:B2\").Value = 1\nEnd Sub\n")
    assert not isinstance(ok, Unsupported) and len(ok.changed) == 4


# ---------------------------------------------------------------------- 結果の字・書き出し・控え
def test_errors_are_short_without_tracebacks_or_install_paths(monkeypatch, tmp_path):
    from todoku import mcp_server

    def boom(*a, **k):
        raise RuntimeError(f"壊れた {Path(api.__file__).resolve()} の中")

    monkeypatch.setattr(mcp_server, "call_tool", boom)
    r = mcp_server.tool_result("open_document", {"path": str(tmp_path / "a.pdf")})
    text = r["content"][0]["text"]
    assert r["isError"] and "Traceback" not in text and str(Path(api.__file__).parent) not in text and "RuntimeError" in text

    def deep(*a, **k):
        raise RecursionError("maximum recursion depth exceeded")

    monkeypatch.setattr(mcp_server, "call_tool", deep)
    text = mcp_server.tool_result("open_document", {"path": "x"})["content"][0]["text"]
    assert text.startswith(limits.STOPPED) and "RecursionError" in text


def test_file_size_cap(monkeypatch, tmp_path):
    p = tmp_path / "a.csv"
    p.write_text("a,b\n1,2\n")
    monkeypatch.setattr(limits, "MAX_FILE", 4)
    # 大きさだけでは断らない：丸ごと読まずに写像で読む（爆弾は展開の上限で止まる）
    k, _doc, parts = api.window(str(p))
    assert k == "csv" and "R2" in parts[0][1]


def test_window_cache_is_capped_by_bytes(monkeypatch, tmp_path):
    monkeypatch.setattr(limits, "WINDOW_CACHE", 20_000)
    api._WINDOWS.clear()
    for i in range(6):
        p = tmp_path / f"{i}.csv"
        p.write_text("id,name\n" + "".join(f"{j},名前{i}_{j}\n" for j in range(20)), encoding="utf-8")
        api.window(str(p))
    assert sum(len(v[1]) for v in api._WINDOWS.values()) <= 20_000
    assert api._WINDOWS  # 小さな窓は控える
    api._WINDOWS.clear()


def test_fill_output_is_atomic_refuses_overwrite_and_hardlinks(tmp_path):
    src = tmp_path / "a.csv"
    src.write_text("name,amount\nyamada,\n")
    out = tmp_path / "out.csv"
    r = api.fill(str(src), {"R2.C2": "5"}, str(out))
    assert r["verified"] and out.read_text().endswith("yamada,5\n")
    assert not [x for x in tmp_path.iterdir() if "todoku-tmp" in x.name]
    api.fill(str(src), {"R2.C2": "5"}, str(out))  # 同じ中身：そのまま
    with pytest.raises(api.ApiError, match="すでに別のファイル"):
        api.fill(str(src), {"R2.C2": "6"}, str(out))
    assert out.read_text().endswith("yamada,5\n")
    r = api.fill(str(src), {"R2.C2": "6"}, str(out), overwrite=True)
    assert out.read_text().endswith("yamada,6\n")
    link = tmp_path / "link.csv"
    os.link(src, link)
    with pytest.raises(api.ApiError, match="元のファイル"):
        api.fill(str(src), {"R2.C2": "7"}, str(link), overwrite=True)
    assert src.read_text() == "name,amount\nyamada,\n"


def test_fill_that_cannot_be_verified_is_named_unverified(monkeypatch, tmp_path):
    src = tmp_path / "a.csv"
    src.write_text("name,amount\nyamada,\n")
    real = api.fields

    def wrong(path, *a, **k):
        fl = real(path, *a, **k)
        if "todoku-tmp" in str(path):  # 読み戻しだけ、違う値に見せる
            fl = [{**f, "value": "違う"} if f["id"] == "R2.C2" else f for f in fl]
        return fl

    monkeypatch.setattr(api, "fields", wrong)
    monkeypatch.setattr(api._open, "read_back", lambda *a, **k: {"R2.C2": "違う"})
    r = api.fill(str(src), {"R2.C2": "5"}, str(tmp_path / "out.csv"))
    assert not r["verified"] and r["out_path"].endswith("out.unverified.csv")
    assert not (tmp_path / "out.csv").exists() and (tmp_path / "out.unverified.csv").exists()
    assert "提出しないでください" in r["say_to_user"]


# ---------------------------------------------------------------------- 別のプロセス（isolate）
def _handle(name, **args):
    from todoku.mcp_server import handle

    return handle({"jsonrpc": "2.0", "id": 1, "method": "tools/call", "params": {"name": name, "arguments": args}})["result"]


def test_isolated_calls_stop_on_time_and_memory_and_the_server_survives(monkeypatch, tmp_path):
    from todoku import isolate

    monkeypatch.setenv("TODOKU_ISOLATE", "1")
    monkeypatch.setenv("TODOKU_MAX_PAGE_OPS", "1000000000")  # 働き手の中の上限では止めず、時間とメモリで止める
    monkeypatch.setenv("TODOKU_MAX_PDF_OPS", "1000000000")
    isolate.worker().stop()  # この環境で立ち上げ直す
    csv = tmp_path / "a.csv"
    csv.write_text("name,amount\nyamada,\n")
    r = _handle("open_document", path=str(csv))
    assert not r.get("isError")
    t = time.time()
    for _ in range(3):
        _handle("open_document", path=str(csv))
    assert (time.time() - t) / 3 < 0.5  # 働き手は立ち上げたまま：控えた窓で速い

    slow = tmp_path / "slow.pdf"
    slow.write_bytes(_chain_pdf(9))  # 10 ** 9 回描く：上限の 1 秒で止める
    monkeypatch.setattr(limits, "TIME_LIMIT", 1)
    t = time.time()
    r = _handle("open_document", path=str(slow))
    assert r["isError"] and limits.STOPPED in r["content"][0]["text"] and "秒" in r["content"][0]["text"]
    assert time.time() - t < 10
    monkeypatch.setattr(limits, "TIME_LIMIT", 120)
    monkeypatch.setattr(limits, "MEMORY_LIMIT", 1024)  # 外からの見張り（ps）で、すぐ超える
    r = _handle("open_document", path=str(slow))
    assert r["isError"] and "メモリ" in r["content"][0]["text"]
    monkeypatch.setattr(limits, "MEMORY_LIMIT", 2048 * limits.MB)
    r = _handle("open_document", path=str(csv))  # 止めた後も、次の呼び出しは動く
    assert not r.get("isError")
    isolate.worker().stop()


def test_windows_falls_back_to_in_process_with_a_notice(monkeypatch, capsys):
    from todoku import isolate, mcp_server

    monkeypatch.setattr(isolate, "_TOLD", set())
    monkeypatch.setattr(isolate.os, "name", "nt")
    monkeypatch.setenv("TODOKU_ISOLATE", "1")
    assert isolate.enabled() is False
    assert "Windows" in capsys.readouterr().err
    r = mcp_server.handle({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}})
    assert "Windows" in r["result"]["instructions"]


def test_missing_ps_is_told_once_not_silent(monkeypatch, capsys):
    from todoku import isolate

    monkeypatch.setattr(isolate, "_TOLD", set())

    def no_ps(*_a, **_k):
        raise FileNotFoundError("ps")

    monkeypatch.setattr(isolate.subprocess, "run", no_ps)
    assert isolate._rss_kb(1) == 0 and isolate._rss_kb(1) == 0
    assert capsys.readouterr().err.count("ps が使えない") == 1
    assert "ps が使えない" in isolate.notice()
