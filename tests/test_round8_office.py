"""第8回：書いたファイルを実物の Word・Excel で開いて見つけたことの直し（合成した書類で確かめる）。"""

from __future__ import annotations

import struct

from test_activex import _cfb_tree

from todoku.cfb import CFB
from todoku.doc.book import FC_NAMES, DocFile
from todoku.doc.fill import DocWriter, Edit


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
