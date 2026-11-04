"""第8回：書いたファイルを実物の Word・Excel で開いて見つけたことの直し（合成した書類で確かめる）。"""

from __future__ import annotations

import io
import struct

from test_activex import _cfb_tree

from todoku import api
from todoku.cfb import CFB
from todoku.doc.book import FC_NAMES, DocFile
from todoku.doc.fill import DocWriter, Edit
from todoku.pdf.doc import PdfDocument
from todoku.pdf.form import FF_MULTILINE, read_fields
from todoku.pdf.pageimage import render


def plain_doc(text: str) -> bytes:
    """本文だけの Word 97 の文書（MS-DOC）。文字の書式・段落の書式の FKP は 1 枚ずつで、字のすぐ後ろで表が終わる。"""
    fc0 = 0x400
    wd = bytearray(0x400) + text.encode("utf-16-le")
    wd += bytes((-len(wd)) % 512)
    page = bytearray(512)
    struct.pack_into("<2I", page, 0, fc0, fc0 + 2 * len(text))
    page[511] = 1
    chpx_pn = len(wd) // 512
    wd += page
    cuts = [0] + [i + 1 for i, ch in enumerate(text) if ch == "\r"]
    page = bytearray(512)
    n = len(cuts) - 1
    struct.pack_into(f"<{n + 1}I", page, 0, *(fc0 + 2 * x for x in cuts))
    page[500] = 1
    for k in range(n):
        page[4 * (n + 1) + 13 * k] = 250
    page[511] = n
    papx_pn = len(wd) // 512
    wd += page
    table = bytearray()
    clx_at = len(table)
    table += b"\x02" + struct.pack("<I", 16) + struct.pack("<II", 0, len(text)) + struct.pack("<HIH", 0, fc0, 0)
    chpx_at = len(table)
    table += struct.pack("<III", fc0, fc0 + 2 * len(text), chpx_pn)
    papx_at = len(table)
    table += struct.pack("<III", fc0, fc0 + 2 * len(text), papx_pn)
    fib = bytearray(900)
    struct.pack_into("<HH", fib, 0, 0xA5EC, 0x00C1)
    struct.pack_into("<H", fib, 0x0A, 0x0200)
    struct.pack_into("<H", fib, 32, 14)
    struct.pack_into("<H", fib, 62, 22)
    lw = [0] * 22
    lw[0], lw[3] = len(wd), len(text)
    struct.pack_into("<22i", fib, 64, *lw)
    struct.pack_into("<H", fib, 152, 93)

    def put(name, fc, lcb):
        struct.pack_into("<II", fib, 154 + 8 * FC_NAMES.index(name), fc, lcb)

    put("Clx", clx_at, 21)
    put("PlcfBteChpx", chpx_at, 12)
    put("PlcfBtePapx", papx_at, 12)
    wd[:900] = fib
    return _cfb_tree({"WordDocument": bytes(wd), "1Table": bytes(table)})


def _chpx_chain(blob: bytes) -> list[tuple[int, int, int, int]]:
    """文字の書式の表の 1 区間ごとに (表の始め, 表の終わり, FKP の初めの字, FKP の終わりの字)。"""
    cfb = CFB(blob)
    wd, tb = cfb.read("WordDocument"), cfb.read("1Table")
    fc, lcb = struct.unpack_from("<II", wd, 154 + 8 * FC_NAMES.index("PlcfBteChpx"))
    n = (lcb - 4) // 8
    afc = struct.unpack_from(f"<{n + 1}I", tb, fc)
    pns = struct.unpack_from(f"<{n}I", tb, fc + 4 * (n + 1))
    out = []
    for i in range(n):
        page = wd[pns[i] * 512:(pns[i] + 1) * 512]
        crun = page[511]
        rgfc = struct.unpack_from(f"<{crun + 1}I", page, 0)
        out.append((afc[i], afc[i + 1], rgfc[0], rgfc[-1]))
    return out


def test_doc_character_format_table_stays_continuous_up_to_the_added_text():
    # 足した字はストリームの後ろ（FKP の後ろ）に入るので、元の表の終わりとの間にすき間ができる。
    # 表の区切りと FKP の字の範囲が合わないと、実物の Word は「ファイルを開こうとして、エラーが発生しました」で開かない
    doc = DocFile(plain_doc("氏名：\r住所：\r"), "a.doc")
    w = DocWriter(doc)
    w.add(Edit(3, 3, "山田", 3, None))
    blob = w.build()
    chain = _chpx_chain(blob)
    assert len(chain) == 2
    for lo, hi, first, last in chain:
        assert (lo, hi) == (first, last)
    assert all(a[1] == b[0] for a, b in zip(chain, chain[1:]))
    after = DocFile(blob, "a.doc")
    assert after.text.startswith("氏名：山田\r住所：\r")
    assert after.ccp["text"] == doc.ccp["text"] + 2


def _gray_jpeg(w: int, h: int, black) -> bytes:
    """濃さだけの基本の JPEG（T.81）。どのブロックも DC だけで、black(x, y) のブロックは黒、ほかは白。w・h は 8 の倍数。"""
    seg = lambda m, b: b"\xff" + bytes([m]) + struct.pack(">H", len(b) + 2) + b
    out = b"\xff\xd8" + seg(0xDB, b"\0" + bytes([8] * 64))  # 量子化はみな 8（DC は画素の値の差そのもの）
    out += seg(0xC0, b"\x08" + struct.pack(">HH", h, w) + b"\x01\x01\x11\x00")
    # DC：分類 0〜11 を長さ 4 の符号に。AC：終わりの印（0x00）だけ、長さ 1 の符号「0」
    out += seg(0xC4, b"\x00" + bytes([0, 0, 0, 12] + [0] * 12) + bytes(range(12)))
    out += seg(0xC4, b"\x10" + bytes([1] + [0] * 15) + b"\x00")
    out += seg(0xDA, b"\x01\x01\x00\x00\x3f\x00")
    bits, pred = [], 0
    for y in range(0, h, 8):
        for x in range(0, w, 8):
            dc = -128 if black(x, y) else 127
            diff, pred = dc - pred, dc
            n = abs(diff).bit_length()
            bits.append(format(n, "04b"))
            if n:
                bits.append(format(diff if diff > 0 else diff + (1 << n) - 1, f"0{n}b"))
            bits.append("0")
    data = "".join(bits)
    data += "1" * (-len(data) % 8)
    raw = int(data, 2).to_bytes(len(data) // 8, "big").replace(b"\xff", b"\xff\x00")
    return out + raw + b"\xff\xd9"


def _pdf_with_jpeg(jpeg: bytes, w: int, h: int) -> bytes:
    content = f"q {w} 0 0 {h} 0 0 cm /I Do Q".encode()
    objs = [b"<</Type/Catalog/Pages 2 0 R>>", b"<</Type/Pages/Kids[3 0 R]/Count 1>>",
            f"<</Type/Page/Parent 2 0 R/MediaBox[0 0 {w} {h}]/Contents 4 0 R/Resources<</XObject<</I 5 0 R>>>>>>".encode(),
            b"<</Length %d>>stream\n" % len(content) + content + b"\nendstream",
            f"<</Type/XObject/Subtype/Image/Width {w}/Height {h}/BitsPerComponent 8/ColorSpace/DeviceGray/Filter/DCTDecode/Length {len(jpeg)}>>stream\n".encode()
            + jpeg + b"\nendstream"]
    out = bytearray(b"%PDF-1.4\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 6\n0000000000 65535 f \n" + b"".join(b"%010d 00000 n \n" % o for o in offs)
    out += b"trailer\n<</Size 6/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF\n" % x
    return bytes(out)


def test_zoomed_page_draws_only_the_opened_part_of_a_jpeg_in_its_own_place():
    # JPEG の見える所だけを開くとき、開いた範囲を RGBA.window に渡す（別の欄に渡すと、切り取った絵がページ全体に引き伸ばされ、
    # スキャンの様式を拡大すると印刷の字が崩れて見える）
    w, h = 800, 800
    jpeg = _gray_jpeg(w, h, lambda x, y: x < 300 and y < 300)  # 左上だけ黒（真ん中に対して片寄らせる）
    doc = PdfDocument(_pdf_with_jpeg(jpeg, w, h), "t.pdf")
    mm = 25.4 / 72
    r = render(doc, 1, [500 * mm, 500 * mm, 700 * mm, 700 * mm], dpi=72)
    assert min(r["rgb"]) > 200  # 右下は白のまま
    r = render(doc, 1, [100 * mm, 100 * mm, 250 * mm, 250 * mm], dpi=72)
    assert max(r["rgb"]) < 60  # 左上は黒のまま
    # 境をまたぐ範囲：黒の終わりが、ページの同じ位置（横 300 pt）にある
    r = render(doc, 1, [250 * mm, 200 * mm, 500 * mm, 280 * mm], dpi=72)
    W = r["w"]
    vals = [r["rgb"][(r["h"] // 2 * W + x) * 3] for x in range(W)]
    edge = next(i for i, v in enumerate(vals) if v > 128)
    assert abs(edge - W / 5) <= 8


def _pdf_text(ops: bytes, w: int = 595, h: int = 842) -> bytes:
    objs = [b"<</Type/Catalog/Pages 2 0 R>>", b"<</Type/Pages/Kids[3 0 R]/Count 1>>",
            f"<</Type/Page/Parent 2 0 R/MediaBox[0 0 {w} {h}]/Contents 4 0 R/Resources<</Font<</F 5 0 R>>>>>>".encode(),
            b"<</Length %d>>stream\n" % len(ops) + ops + b"\nendstream",
            b"<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>"]
    out = bytearray(b"%PDF-1.4\n")
    offs = []
    for i, o in enumerate(objs, 1):
        offs.append(len(out))
        out += b"%d 0 obj\n" % i + o + b"\nendobj\n"
    x = len(out)
    out += b"xref\n0 6\n0000000000 65535 f \n" + b"".join(b"%010d 00000 n \n" % o for o in offs)
    out += b"trailer\n<</Size 6/Root 1 0 R>>\nstartxref\n%d\n%%%%EOF\n" % x
    return bytes(out)


def test_below_anchor_is_not_stopped_by_a_line_of_only_spaces(tmp_path):
    # マスの中に空白だけの行があっても、墨はないので「下に書く空きがありません」と断らない（窓は ↓余白 と示している）
    ops = (b"1 w 100 500 300 100 re S BT /F 12 Tf 110 585 Td (Addr) Tj ET BT /F 12 Tf 110 570 Td ( ) Tj ET "
           b"BT /F 12 Tf 110 510 Td (Tel) Tj ET")
    p = tmp_path / "a.pdf"
    p.write_bytes(_pdf_text(ops))
    r = api.fill(str(p), {"@1": {"after": "Addr", "where": "below", "page": 1, "value": "Tokyo"}}, str(tmp_path / "o.pdf"))
    assert r["proof_ok"] and r["read_back_ok"]


def test_value_drawn_on_two_lines_is_a_multiline_field_and_audit_accepts_the_output(tmp_path):
    # 2 行に描いた欄が 1 行の欄のままだと、見た目を作り直すソフトで 1 行に潰れ、書いた後の audit も「改行は書けません」になる
    ops = b"1 w 100 500 300 100 re S BT /F 12 Tf 110 610 Td (Name) Tj ET"
    p = tmp_path / "a.pdf"
    p.write_bytes(_pdf_text(ops))
    out = tmp_path / "o.pdf"
    r = api.fill(str(p), {"@1": {"page": 1, "box": [40, 90, 130, 120], "value": "Yamada\nHanako"}}, str(out))
    assert r["proof_ok"] and r["read_back_ok"]
    a = api.audit(str(out))
    bad = [c for c in a["checks"] if c["name"] == "全入力欄への書き込み" and not c["ok"]]
    assert not bad, bad
    doc = PdfDocument(out.read_bytes(), "o.pdf")
    multi = [f for f in read_fields(doc) if f.flags & FF_MULTILINE]
    assert multi


def test_doc_paragraph_place_takes_align_like_docx(tmp_path):
    # 説明書きのとおり、Word の段落の中の所は .doc でも align で書ける：空きの幅を保って、前後に半角の空白を足す
    p = tmp_path / "a.doc"
    p.write_bytes(plain_doc("申請日　　　　　　年\r"))
    place = next(f["id"] for f in api.fields(str(p)) if f["id"].startswith("P"))
    r = api.fill(str(p), {place: {"value": "8", "align": "right"}}, str(tmp_path / "o.doc"))
    assert r["proof_ok"] and r["read_back_ok"], r
    assert r["aligned"][place]["align"] == "right" and r["aligned"][place]["pad"][0] > 0
    after = DocFile((tmp_path / "o.doc").read_bytes(), "o.doc")
    assert " 8" in after.text
    # 指定がなければ、これまでのまま
    r = api.fill(str(p), {place: "8"}, str(tmp_path / "o2.doc"))
    assert r["proof_ok"] and r["read_back_ok"] and "aligned" not in r


def test_xlsb_appending_to_a_row_and_adding_the_next_row_together(tmp_path):
    # ある行の最後のセルの後ろと、その次の新しい行に、同じ書き込みで書く：足す位置が同じになっても、それぞれの行に入る
    import test_xlsb_grow as G

    blob, _ = G.book()
    p = tmp_path / "a.xlsb"
    p.write_bytes(blob)
    r = api.fill(str(p), {"S!E1": "右の空", "S!B2": "新しい行"}, str(tmp_path / "o.xlsb"))
    assert r["proof_ok"] and r["read_back_ok"]
    assert "recalculation_unchecked" not in r  # 書いた後のブックを読み直して確かめた
    k, doc, parts = api.window(str(tmp_path / "o.xlsb"))
    text = "\n".join(t for _, t in parts)
    assert "右の空" in text and "新しい行" in text


def test_written_pdf_field_warns_about_viewers_that_redraw_from_the_value(tmp_path):
    p = tmp_path / "a.pdf"
    p.write_bytes(_pdf_text(b"1 w 100 500 300 100 re S BT /F 12 Tf 110 610 Td (Name) Tj ET"))
    r = api.fill(str(p), {"@1": {"page": 1, "box": [40, 90, 130, 120], "value": "Yamada"}}, str(tmp_path / "o.pdf"))
    assert sum("Mac のプレビュー" in w for w in r["warnings"]) == 1
