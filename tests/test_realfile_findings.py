"""実物のファイル（別の道具と照らした）で見つかった不具合の、自作の試験。実物そのものはリポジトリに入れない。

何と照らして見つかったかは、各試験の説明に書く。"""

from __future__ import annotations

import io
import struct
import zipfile

import pytest
import xlsb_maker as m
from test_more_formats import _r

from todoku import api
from todoku.book import open_book
from todoku.formula.parser import parse_cell
from todoku.xlsb.audit import raw_values
from todoku.xlsb.grow import cell_positions
from todoku.zipraw import RawZip


def _p(tmp_path, name, blob):
    p = tmp_path / name
    p.write_bytes(blob)
    return str(p)


# ---------------------------------------------------------------- .xlsb
def _frt_sheet(tail_rid: int = 1) -> bytes:
    """Excel が書く形：データの後の BrtFRTBegin〜End の中に、番号 1（セルの番号と同じ）のレコード。"""
    rows = m.row(0) + m.real(0, 1.5) + m.row(1) + m.real(0, 3.7477917436899863e-22)
    frt = m.rec(35, struct.pack("<I", 0x10E02)) + m.rec(tail_rid, bytes(9)) + m.rec(36)
    return m.sheet(rows, tail=frt)


def test_xlsb_record_in_future_frame_does_not_erase_last_row():
    """Excel の書いた .xlsb で、最後の行の値が消えていた（Excel の表示と照らして発見）。"""
    wb = open_book(m.package({"S": _frt_sheet()}), "a.xlsb")
    assert wb.sheet("S").cells[(2, 1)].value == 3.7477917436899863e-22


def test_xlsb_cells_only_inside_sheet_data_for_all_readers(tmp_path):
    blob = m.package({"S": _frt_sheet(tail_rid=5)})  # 番号 5（数のセルと同じ）でも
    wb = open_book(blob, "a.xlsb")
    assert set(wb.sheet("S").cells) == {(1, 1), (2, 1)}
    assert raw_values(wb)["S"] == {(1, 1), (2, 1)}  # 検査の数え方も同じ
    z = RawZip(blob)
    assert set(cell_positions(z.read("xl/worksheets/sheet1.bin"))) == {(0, 0), (1, 0)}
    assert api.audit(_p(tmp_path, "a.xlsb", blob))["checks"][0]["ok"]


def _rename_part(blob: bytes, old: str, new: str) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(blob)) as a, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as b:
        for i in a.infolist():
            b.writestr(new if i.filename == old else i.filename, a.read(i.filename))
    return out.getvalue()


def test_part_names_do_not_depend_on_case():
    """OPC の部品の名前は大文字小文字を区別しない（Excel は開く。とどくは「共有文字列がありません」と止まっていた）。"""
    rows = m.row(0) + m.isst(0, 0)
    blob = _rename_part(m.package({"S": m.sheet(rows)}, strings=["共有の字"]), "xl/sharedStrings.bin", "xl/SharedStrings.bin")
    assert open_book(blob, "a.xlsb").sheet("S").cells[(1, 1)].value == "共有の字"


# ---------------------------------------------------------------- .xls の共有文字列
def _xls(strings: list[str], declared: int, cells: bytes) -> bytes:
    from test_embeds import make_cfb

    sst = struct.pack("<II", len(strings), declared) + b"".join(struct.pack("<HB", len(s), 1) + s.encode("utf-16-le") for s in strings)
    head = _r(0x0809, struct.pack("<HHHHII", 0x600, 5, 0, 0, 0, 6)) + _r(0x00E0, bytes(20))
    tail = _r(0x00FC, sst) + _r(0x000A, b"")
    bs = lambda at: _r(0x0085, struct.pack("<IBB", at, 0, 0) + bytes([1, 1]) + "S".encode("utf-16-le"))  # noqa: E731
    wbk = head + bs(len(head) + len(bs(0)) + len(tail)) + tail
    sh = _r(0x0809, struct.pack("<HHHHII", 0x600, 0x10, 0, 0, 0, 6)) + cells + _r(0x000A, b"")
    return make_cfb({"Workbook": wbk + sh})


def test_xls_sst_count_may_be_larger_than_the_strings():
    """宣言の数より字が少ない .xls を Excel は開く（とどくは開けなかった）。"""
    cell = _r(0x00FD, struct.pack("<HHHI", 0, 0, 0, 1))
    wb = open_book(_xls(["甲", "乙"], 9, cell), "a.xls")
    assert wb.sheet("S").cells[(1, 1)].value == "乙"


def test_xls_string_number_outside_the_table_is_reported_not_blank():
    cell = _r(0x00FD, struct.pack("<HHHI", 0, 0, 0, 5))
    wb = open_book(_xls(["甲", "乙"], 2, cell), "a.xls")
    sh = wb.sheet("S")
    assert any("A1" in u and "5" in u for u in sh.unread), sh.unread
    assert sh.cells[(1, 1)].value in (None, "")


# ---------------------------------------------------------------- Excel 4 の書式
def test_excel4_format_records_are_read():
    """BIFF4 の FORMAT は 0x041E（0x001E と読んでいて、すべての書式を落とし、(2) を 2、168,184,400 を 168184400 と見せていた。Excel の表示と照らして発見）。"""
    fmts = [b"\x07General", b"\x05#,##0", b"\x030.0"]
    xf = [bytes([0, i, 0, 0]) + bytes(4) + bytes(4) for i in range(3)]
    blob = b"".join([
        _r(0x0409, struct.pack("<HHH", 4 << 8, 0x10, 0)),
        _r(0x0042, struct.pack("<H", 932)),
        *[_r(0x041E, b"\x00\x00" + f) for f in fmts],
        *[_r(0x0443, x) for x in xf],
        _r(0x0203, struct.pack("<HHH", 0, 0, 1) + struct.pack("<d", 168184400.0)),
        _r(0x0203, struct.pack("<HHH", 1, 0, 2) + struct.pack("<d", 0.25)),
        _r(0x000A, b""),
    ])
    wb = open_book(blob, "a.xls")
    sh = wb.sheets[0]
    assert wb.style(sh.cells[(1, 1)].style).num_fmt == "#,##0"
    assert wb.style(sh.cells[(2, 1)].style).num_fmt == "0.0"


# ---------------------------------------------------------------- .ods
def _ods(cell_xml: str) -> bytes:
    content = ('<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
               'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" '
               'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0" office:version="1.2"><office:body><office:spreadsheet>'
               f'<table:table table:name="S"><table:table-row>{cell_xml}</table:table-row></table:table></office:spreadsheet></office:body></office:document-content>')
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        z.writestr("mimetype", "application/vnd.oasis.opendocument.spreadsheet")
        z.writestr("content.xml", content)
        z.writestr("META-INF/manifest.xml", '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>')
    return out.getvalue()


def _ods_value(tmp_path, cell_xml: str):
    p = _p(tmp_path, "a.ods", _ods(cell_xml))
    _k, doc, _parts = api.window(p)
    return next(c for c in doc["sheets"][0]["cells"] if parse_cell(c["cell"]) == (1, 1))["value"]


def test_ods_keeps_trailing_spaces(tmp_path):
    """LibreOffice が書く終わりの空白を落としていた（LibreOffice・Excel の読みと照らして発見）。"""
    x = '<table:table-cell office:value-type="string"><text:p>{}</text:p></table:table-cell>'
    assert _ods_value(tmp_path, x.format("前後 <text:s/>")) == "前後  "
    assert _ods_value(tmp_path, x.format("字 ")) == "字 "
    assert _ods_value(tmp_path, x.format('<text:s text:c="2"/>字')) == "  字"


def test_ods_empty_string_value_with_visible_text(tmp_path):
    """office:string-value が空でも、段落に字があれば、それが見えている値（空のセルにしない）。"""
    x = '<table:table-cell office:value-type="string" table:formula="of:=1/0" office:string-value=""><text:p>#DIV/0!</text:p></table:table-cell>'
    assert _ods_value(tmp_path, x) == "#DIV/0!"


# ---------------------------------------------------------------- .doc：BMP の外の字
def test_doc_cp_string_keeps_one_python_char_per_cp():
    """𠮷 は CP を 2 つ使う。Python の文字列で 1 字に数えると、後ろの字の書式と、本文と脚注の境が 1 つずつずれる。"""
    from todoku.doc.book import CP_PAD, cp_string

    assert cp_string("a𠮷b") == "a𠮷" + CP_PAD + "b"
    assert len(cp_string("𠮷𩸽😀")) == 6
    assert cp_string("a\ud842") == "a�"  # 組にならないものは U+FFFD
    assert cp_string("\ud842" + "\udfb7") == "𠮷" + CP_PAD  # ピースの境で分かれた組


def _doc97(text: str, deleted: tuple[int, int] | None = None) -> bytes:
    """Word 97 の .doc（1 つのピース）。deleted：変更履歴で削除した字の CP の範囲。"""
    from todoku.cfb import build_cfb

    fc0 = 2048
    tb = text.encode("utf-16-le")
    n = len(tb) // 2
    wd = bytearray(fc0 + len(tb))
    wd[fc0:] = tb
    struct.pack_into("<HHHHHHHI", wd, 0, 0xA5EC, 0x00C1, 0, 0x0411, 0, 0x1200, 0x00BF, 0)
    struct.pack_into("<H", wd, 0x20, 14)
    struct.pack_into("<H", wd, 0x3E, 22)
    struct.pack_into("<22i", wd, 0x40, len(wd), 0, 0, n, *([0] * 18))
    struct.pack_into("<H", wd, 0x98, 93)
    clx = b"\x02" + struct.pack("<I", 16) + struct.pack("<II", 0, n) + struct.pack("<HIH", 0, fc0, 0)
    table = clx
    struct.pack_into("<II", wd, 0x9A + 33 * 8, 0, len(clx))
    if deleted:
        pn = len(wd) // 512 + 1
        page = bytearray(512)
        cuts = [fc0, fc0 + 2 * deleted[0], fc0 + 2 * deleted[1], fc0 + len(tb)]
        struct.pack_into("<4I", page, 0, *cuts)
        page[16:19] = bytes([0, 200, 0])  # 真ん中の区間だけ書式あり（400 バイト目）
        page[400:404] = bytes([3, 0x00, 0x08, 0x01])  # sprmCFRMarkDel
        page[511] = 3
        wd += bytes(pn * 512 - len(wd)) + bytes(page)
        plc = struct.pack("<2I", cuts[0], cuts[-1]) + struct.pack("<I", pn)
        struct.pack_into("<II", wd, 0x9A + 12 * 8, len(table), len(plc))
        table += plc
    return build_cfb({"WordDocument": bytes(wd), "1Table": table})


def test_doc_marks_after_outside_bmp_char_stay_on_their_chars(tmp_path):
    """𠮷 の後ろの変更履歴が 1 字ずれ、消した字が本文に出て、本文の字が消えていた（LibreOffice が書いた .doc を、作った側の字と照らして発見）。"""
    text = "𠮷田𩸽の前削除字の後\r"
    units_before = len(text[:text.index("削")].encode("utf-16-le")) // 2
    p = _p(tmp_path, "a.doc", _doc97(text, (units_before, units_before + 3)))
    t = "\n".join(x[1] for x in api.window(p)[2])
    assert "𠮷田𩸽の前の後" in t
    assert "〔非表示の文字: 削除字〕" in t


def test_doc_stories_after_outside_bmp_char():
    """本文の 𠮷 の分だけ、脚注との境がずれて、脚注の頭の字が本文に入っていた。"""
    from todoku.cfb import build_cfb
    from todoku.doc.book import DocFile

    blob = _doc97("𠮷の本文\r脚注の字\r\r")
    z_cfb = __import__("todoku.cfb", fromlist=["CFB"]).CFB(blob)
    wd = bytearray(z_cfb.read("WordDocument"))
    struct.pack_into("<2i", wd, 0x40 + 3 * 4, 6, 6)  # 本文 6 単位（𠮷=2）・脚注 6 単位
    d = DocFile(build_cfb({"WordDocument": bytes(wd), "1Table": z_cfb.read("1Table")}), "a.doc")
    assert [p.text for p in d.body] == ["𠮷の本文"]
    assert [p.text for p in d.stories[0][2]][0] == "脚注の字"


def test_doc_edit_counts_utf16_units():
    """書く字の CP の数は UTF-16 の単位（字数で数えると、𠮷 を書いたときにピースの表がずれる。LibreOffice の読み戻しと照らした）。"""
    from todoku.doc.fill import Edit

    assert Edit(0, 0, "𩸽田", 0).n == 3


# ---------------------------------------------------------------- 画像
def _tiff(entries: dict, data: bytes, little: bool = True) -> bytes:
    e = "<" if little else ">"
    head = (b"II*\0" if little else b"MM\0*") + struct.pack(e + "I", 8)
    n = len(entries) + 1
    at = 8 + 2 + 12 * n + 4
    extra = b""
    rows = []
    entries = dict(entries)
    entries[273] = None
    for tag in sorted(entries):
        v = entries[tag]
        if tag == 273:
            rows.append(("off", tag))
            continue
        if isinstance(v, (list, tuple)):
            pos = at + len(extra)
            extra += struct.pack(e + f"{len(v)}H", *v)
            rows.append(struct.pack(e + "HHII", tag, 3, len(v), pos))
        else:
            rows.append(struct.pack(e + "HHI", tag, 3, 1) + struct.pack(e + "HH", v, 0))
    data_at = at + len(extra)
    body = struct.pack(e + "H", len(rows))
    for r in rows:
        body += struct.pack(e + "HHII", 273, 4, 1, data_at) if isinstance(r, tuple) else r
    return head + body + struct.pack(e + "I", 0) + extra + data


def test_tiff_16bit_little_endian_uses_the_high_byte():
    """II の 16 ビット TIFF で下の桁を色にしていた（画面全体が別の色に。ImageMagick の読みと照らして発見）。"""
    from todoku.imaging.formats import decode

    px = struct.pack("<3H", 0x8012, 0x4034, 0xFF56)
    blob = _tiff({256: 1, 257: 1, 258: (16, 16, 16), 259: 1, 262: 2, 277: 3, 278: 1, 279: 6}, px)
    assert decode(blob, 0).rgb == bytes([0x80, 0x40, 0xFF])
    be = _tiff({256: 1, 257: 1, 258: (16, 16, 16), 259: 1, 262: 2, 277: 3, 278: 1, 279: 6}, struct.pack(">3H", 0x8012, 0x4034, 0xFF56), little=False)
    assert decode(be, 0).rgb == bytes([0x80, 0x40, 0xFF])


def test_tiff_16bit_little_endian_predictor():
    from todoku.imaging.formats import decode

    # 2 画素・差分：1 画素目 0x0100、2 画素目はその差 0x0100 → 0x0200
    px = struct.pack("<2H", 0x0100, 0x0100)
    blob = _tiff({256: 2, 257: 1, 258: 16, 259: 1, 262: 1, 277: 1, 278: 1, 279: 4, 317: 2}, px)
    assert decode(blob, 0).rgb == bytes([1, 1, 1, 2, 2, 2])


def test_tiff_associated_alpha_on_white():
    """ExtraSamples 1（色に透けを掛けてある）を、掛けていないものと同じに混ぜて暗くしていた。"""
    from todoku.imaging.formats import decode

    px = bytes([100, 0, 0, 128])  # 掛けてある赤：白の紙の上では 100 + (255-128) = 227
    blob = _tiff({256: 1, 257: 1, 258: (8, 8, 8, 8), 259: 1, 262: 2, 277: 4, 278: 1, 279: 4, 338: 1}, px)
    assert decode(blob, 0).rgb[0] == 227


def test_ccitt_eol_with_fill_bits():
    """T4Options の 4（EOL の前の詰め物）を行のそろえと読んで開けなかった（G3 の TIFF）。"""
    from todoku.pdf.raster import ccitt

    # 8 列：白 8（10011）が 2 行。EOL（0 が 11 個と 1）の前に、終わりをバイトにそろえる詰め物
    bits = "0000" + "000000000001" + "10011" + "0000000" + "000000000001" + "10011"
    data = int(bits.ljust((len(bits) + 7) // 8 * 8, "0"), 2).to_bytes((len(bits) + 7) // 8, "big")
    g = ccitt(data, {"K": 0, "Columns": 8, "Rows": 2})
    assert g.px == b"\xff" * 16


def test_ccitt_rtc_ends_the_image():
    from todoku.pdf.raster import ccitt

    eol = "000000000001"
    bits = eol + "10011" + eol * 6
    data = int(bits.ljust((len(bits) + 7) // 8 * 8, "0"), 2).to_bytes((len(bits) + 7) // 8, "big")
    g = ccitt(data, {"K": 0, "Columns": 8})
    assert g.h == 1


class _Bits:
    """VP8L のビット列を、下の桁から書く。"""

    def __init__(self):
        self.acc = 0
        self.n = 0

    def put(self, v, k):
        self.acc |= v << self.n
        self.n += k

    def code(self, bits: str):  # ハフマンの符号は、上の桁から順に
        for ch in bits:
            self.put(int(ch), 1)

    def simple1(self, sym):  # 記号 1 つの符号（8 ビットの形）
        self.put(1, 1)
        self.put(0, 1)
        self.put(1, 1)
        self.put(sym, 8)

    def simple2(self, a, b):  # 記号 2 つの符号（1 ビット）
        self.put(1, 1)
        self.put(1, 1)
        self.put(1, 1)
        self.put(a, 8)
        self.put(b, 8)

    def bytes(self):
        return self.acc.to_bytes((self.n + 7) // 8 + 8, "little")


def test_webp_lossless_group_follows_position_after_a_copy():
    """距離の符号で区画をまたいだ後も、前の区画のハフマンの組で読み続けて、絵が壊れていた（sips の読みと照らして発見）。

    8×1 画素・4 画素ごとの区画。区画 0 は組 0（緑は 10 か距離の符号）、区画 1 は組 1（緑は 200 だけ）。
    1 画素目の後ろで 5 画素の距離の符号を使うと、区画 1 の途中（6 列目）に止まる。残りは組 1 で読む。"""
    from todoku.imaging.webp import decode_vp8l

    w = _Bits()
    w.put(0, 1)  # 変換なし
    w.put(0, 1)  # 色のキャッシュなし
    w.put(1, 1)  # ハフマンの組の分け方あり
    w.put(0, 3)  # 4 画素ごと
    # 分け方の画像（2×1）：緑が組の番号
    w.put(0, 1)
    w.simple2(0, 1)
    w.simple1(0)
    w.simple1(0)
    w.simple1(255)
    w.simple1(0)
    w.code("0")
    w.code("1")
    # 組 0 の緑：記号 10 と 260 に 1 ビットの符号。長さの符号の符号：18→0、1→10、17→11
    w.put(0, 1)
    w.put(0, 4)  # 4 つ
    for ln in (2, 1, 0, 2):  # 順は 17・18・0・1
        w.put(ln, 3)
    w.put(0, 1)
    w.code("11")  # 17：0 が 10 個（0〜9）
    w.put(7, 3)
    w.code("10")  # 記号 10 の長さ 1
    w.code("0")  # 18：0 が 138 個（11〜148）
    w.put(127, 7)
    w.code("0")  # 18：0 が 111 個（149〜259）
    w.put(100, 7)
    w.code("10")  # 記号 260 の長さ 1
    w.code("0")  # 18：0 が 19 個（261〜279）
    w.put(8, 7)
    w.simple1(0)
    w.simple1(0)
    w.simple1(255)
    w.simple1(1)  # 距離の記号 1：距離の表の 2 番目（左へ 1）
    # 組 1
    w.simple1(200)
    w.simple1(0)
    w.simple1(0)
    w.simple1(255)
    w.simple1(0)
    # 画素
    w.code("0")  # 緑 10
    w.code("1")  # 記号 260：長さの記号 4
    w.put(0, 1)  # 長さ 5
    px = decode_vp8l(w.bytes(), 8, 1)
    assert [(v >> 8) & 255 for v in px] == [10, 10, 10, 10, 10, 10, 200, 200]


def test_webp_animation_blends_over_the_previous_frame():
    """透けのある枠を、白に置いてから貼っていた（前の枠が見えない。ImageMagick の読みと照らして発見）。"""
    import todoku.imaging.webp as wp

    frames = {"canvas": (2, 1), "bg": None, "flags": 2, "frames": [
        {"x": 0, "y": 0, "kind": b"X", "s": 0, "e": 0, "alpha": None, "blend": True, "dispose": False, "n": 0},
        {"x": 0, "y": 0, "kind": b"X", "s": 0, "e": 0, "alpha": None, "blend": True, "dispose": False, "n": 1}]}
    plain = {0: (2, 1, bytes([10, 20, 30, 40, 50, 60]), None), 1: (2, 1, bytes([200, 200, 200, 200, 200, 200]), bytes([0, 255]))}
    orig_parse, orig_plain = wp._parse, wp._decode_plain
    try:
        wp._parse = lambda blob: frames
        wp._decode_plain = lambda blob, f: plain[f["n"]]
        got = wp.decode(b"", 1)
    finally:
        wp._parse, wp._decode_plain = orig_parse, orig_plain
    assert got.rgb == bytes([10, 20, 30, 200, 200, 200])


def test_wmf_without_placeable_header_is_recognised():
    """大きさの見出しのない WMF を「扱えない形式」としていた（META_HEADER から始まる標準の形）。"""
    from todoku.imaging.formats import sniff

    head = struct.pack("<HHHIHIH", 1, 9, 0x0300, 100, 0, 3, 0)
    assert sniff(head + struct.pack("<IH", 3, 0)) == "wmf"
    assert sniff(struct.pack("<HHHIHIH", 3, 9, 0x0300, 100, 0, 3, 0) + bytes(6)) != "wmf"


# ---------------------------------------------------------------- .ppt・.pptx
def _ppt_record(t, body=b"", ver=0, inst=0):
    return struct.pack("<HHI", ver | inst << 4, t, len(body)) + body


def test_ppt_comment_lines():
    """スライドのコメントを読んでいなかった（LibreOffice の読みと照らして発見）。"""
    from todoku.embeds.ppt import _comment_lines

    cm = _ppt_record(0x2EE0, _ppt_record(0x0FBA, "書いた人".encode("utf-16-le"), inst=0) + _ppt_record(0x0FBA, "直してください\r二行目".encode("utf-16-le"), inst=1), ver=15)
    blob = _ppt_record(0x138B, cm)
    tag = _ppt_record(0x138A, blob, ver=15)
    slide = _ppt_record(1006, _ppt_record(5000, tag, ver=15), ver=15)
    assert _comment_lines(slide, 0) == ["コメント（書いた人）：直してください ／ 二行目"]


def _cfb_ppt(extra: dict) -> bytes:
    from test_embeds import make_cfb

    return make_cfb(extra)


def test_ppt95_reads_text_content_and_says_what_it_could_not_read(tmp_path):
    """PowerPoint 95：「字はありません」と黙っていた（346 行の字がある実物。LibreOffice の読みと照らして発見）。"""
    summary = _summary_codepage(1252)
    p = _p(tmp_path, "a.ppt", _cfb_ppt({"PowerPoint Document": b"\x03\0\0\0\xff\xff\0\0" + bytes(32), "Current User": b"\x0b\0\0\0x", "Text_Content": b"Title\r\rBody \xe9\r",
                                        "\x05SummaryInformation": summary}))
    _k, doc, parts = api.window(p)
    t = "\n".join(x[1] for x in parts)
    assert "Title" in t and "Body é" in t
    assert any("区切り" in u for u in doc["unread"])


def _summary_codepage(cp: int) -> bytes:
    prop = struct.pack("<HxxH", 2, cp) + b"\0\0"
    sec = struct.pack("<II", 8 + 8 + len(prop), 1) + struct.pack("<II", 1, 16) + prop
    return struct.pack("<HHI", 0xFFFE, 0, 0x20005) + bytes(16) + struct.pack("<I", 1) + bytes(16) + struct.pack("<I", 48) + sec


def test_ppt_unreadable_structure_is_an_error_not_an_empty_document(tmp_path):
    p = _p(tmp_path, "a.ppt", _cfb_ppt({"PowerPoint Document": b"\x03\0\0\0\xff\xff\0\0" + bytes(32), "Current User": b"\x0b\0\0\0x"}))
    with pytest.raises(api.ApiError, match="読めず"):
        api.window(p)


def test_pptx_comments_old_and_new_forms(tmp_path):
    import pptx_maker

    blob = pptx_maker.make([pptx_maker.sp(pptx_maker.para("本文"))])
    out = io.BytesIO()
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    with zipfile.ZipFile(io.BytesIO(blob)) as a, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as b:
        for i in a.infolist():
            d = a.read(i.filename)
            if i.filename == "ppt/slides/_rels/slide1.xml.rels":
                d = d.replace(b"</Relationships>",
                              b'<Relationship Id="rC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>'
                              b'<Relationship Id="rC2" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/modern1.xml"/></Relationships>')
            b.writestr(i.filename, d)
        b.writestr("ppt/commentAuthors.xml", f'<p:cmAuthorLst xmlns:p="{P}"><p:cmAuthor id="0" name="書いた人"/></p:cmAuthorLst>')
        b.writestr("ppt/comments/comment1.xml", f'<p:cmLst xmlns:p="{P}"><p:cm authorId="0"><p:text>古い形のコメント</p:text></p:cm></p:cmLst>')
        b.writestr("ppt/comments/modern1.xml", '<p188:cmLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                                               '<p188:cm authorId="x"><p188:replyLst><p188:reply authorId="x"><p188:txBody><a:p><a:r><a:t>返信の字</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst>'
                                               '<p188:txBody><a:p><a:r><a:t>新しい形の</a:t></a:r><a:r><a:t>コメント</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>')
    p = _p(tmp_path, "a.pptx", out.getvalue())
    t = "\n".join(x[1] for x in api.window(p)[2])
    assert "コメント（書いた人）：古い形のコメント" in t
    assert "新しい形のコメント" in t and "返信の字" in t
    assert "本文" in t
