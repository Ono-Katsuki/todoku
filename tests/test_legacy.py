"""古い・独自の形式：一太郎（.jtd）、Excel 5.0/95（BIFF5）、Word 6.0/95。

小さな合成のファイル（このテストの中で作る）で、読む・字数を変えない上書き・証明・断り方を確かめ、
集めた様式（corpus/wide_jtd・wide_legacy。再配布しないので、なければ飛ばす）で、全体の検査を通す。
"""

from __future__ import annotations

import struct
from pathlib import Path

import pytest

from todoku.book import kind_of
from todoku.cfb import CFB, build_cfb

CORPUS = Path(__file__).resolve().parent.parent / "corpus"


# ====================================================================== 一太郎（合成）
def _text_units(text: str) -> list[int]:
    """窓の字の並び → 単位。「¶」は段落の終わり、「|」は表のマスの始まり（左右の位置つきの記録）、「⏎」は表の行の終わり。"""
    out: list[int] = []
    x = 0
    for ch in text:
        if ch == "¶":
            out.append(0x0A)
        elif ch == "⏎":
            out.append(0x0E)
            x = 0
        elif ch == "|":
            # 記録：0x1C 種類 長さ … 長さ 0 種類 0x1F（種類 0x30 は表のマス。中身の 2・3 番目が左右の位置）
            out += [0x1C, 0x30, 13, 0, x, x + 10, 0xFF, 0, 0x0C, 13, 0, 0x30, 0x1F]
            x += 12
        else:
            out.append(ord(ch))
    return out


def _ssmg(segments: list[tuple[int, bytes]]) -> bytes:
    bsize = 256
    blocks: list[bytes] = []
    directory = b""
    for sid, data in segments:
        idx = []
        for k in range(0, max(len(data), 1), bsize):
            idx.append(len(blocks))
            blocks.append(data[k:k + bsize].ljust(bsize, b"\0"))
        directory += struct.pack(">IIIII", sid, 0, len(data), len(idx), len(idx)) + struct.pack(f">{len(idx)}I", *idx)
    head = b"SsmgV.01" + struct.pack(">III", len(segments), bsize, len(blocks))
    return head + b"".join(blocks) + directory + b"\0\0\0\0"


def _textseg(text: str) -> bytes:
    u = _text_units(text)
    return b"TextV.01" + struct.pack(">I", len(u)) + struct.pack(f">{len(u)}H", *u) + b"\0\0\0\0\0"


def make_jtd(body: str, box: str | None = None) -> bytes:
    streams = {"\x04JSRV_SegmentInformation": b"VDA_DOC\0" + bytes(88), "Font": b"FontV.01" + bytes(8)}
    if box is not None:
        # 枠の置き場所：記録（種類 0、部品の種類 0x30、6 番目に枠の番号）と、中身が 0x02 の組
        anchor = "".join(map(chr, [0x1C, 0, 14, 0, 0x30, 0xFFFF, 0x0507, 0x12, 0, 0, 14, 0, 0, 0x1F,
                                   0x1C, 1, 7, 0, 0, 1, 0x1D, 2, 0x1E, 5, 0, 1, 0x1F]))
        body = anchor + body
        streams["LayoutBoxText"] = _ssmg([(0, _textseg(box))])
    streams["DocumentText"] = _ssmg([(0, _textseg(body))])
    return build_cfb(streams)


def test_jtd_reads_tables_boxes_and_overwrites_in_place():
    from todoku.docx.fill import FillError, written_of
    from todoku.jtd.audit import prove, raw_chars, run
    from todoku.jtd.book import JtdFile
    from todoku.jtd.fill import fill
    from todoku.jtd.window import build, to_markdown

    blob = make_jtd("申請書¶氏名　　　　　¶|住所|　　　　⏎|電話|¶⏎¶□同意する¶", box="受付印¶")
    assert kind_of(blob) == "jtd"
    doc = JtdFile(blob, "t.jtd")
    win = build(doc)
    md = to_markdown(win)
    assert "一太郎" in md and "受付印" in md and "住所 ｜" in md
    # 窓とは別の読み方で数えた字が、窓に全部出ている
    assert sum(raw_chars(doc).values()) == len("申請書氏名住所電話同意する受付印")
    fields = {f["id"]: f for f in win["fields"]}
    empty_cell = [f for f in fields.values() if f["kind"] == "text" and f.get("readonly")]
    assert empty_cell, "字のない空のマスは、書けない欄として示す"
    room = [f for f in fields.values() if not f.get("readonly") and f["kind"] != "checkbox"]
    target = room[0]
    res = fill(JtdFile(blob), {target["id"]: "山田"})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors + pr.readback
    assert len(res.blob) == len(blob) and sum(a != b for a, b in zip(blob, res.blob)) <= 4
    assert "山田" in "".join(p.text for _k, _p, bl in JtdFile(res.blob).all_blocks()
                             for p in __import__("todoku.docx.document", fromlist=["iter_paras"]).iter_paras(bl))
    # 字数を変える書き込み・字のない所への書き込みは断る
    with pytest.raises(FillError, match="字数"):
        fill(JtdFile(blob), {"@1": {"old": "申請書", "new": "申請"}})
    with pytest.raises(FillError):
        fill(JtdFile(blob), {empty_cell[0]["id"]: "x"})
    assert run(blob, "t.jtd").checks[0].ok


def test_jtd_rejects_broken_records():
    from todoku.jtd.book import JtdError, JtdFile

    bad = make_jtd("abc" + chr(0x1C) + chr(0x10) + chr(5) + "xy¶")
    with pytest.raises(JtdError):
        JtdFile(bad)


def test_lh5_roundtrip_crc():
    from todoku.jtd.lha import crc16

    assert crc16(b"123456789") == 0xBB3D  # CRC-16/ARC の確かめの値


JTD = sorted(CORPUS.glob("wide_jtd/**/*.jtd")) if CORPUS.exists() else []


@pytest.mark.parametrize("path", JTD[:40], ids=lambda p: p.name)
def test_jtd_corpus_reading_and_overwrite(path):
    from todoku.jtd.audit import run

    au = run(path.read_bytes(), path.name)
    checks = {c.name: c for c in au.checks}
    assert checks["読み落とし（文書の文字）"].ok, checks["読み落とし（文書の文字）"].items
    assert checks["ストリームの一覧"].ok, checks["ストリームの一覧"].items
    w = checks["全入力欄への書き込み"]
    assert w.ok or "圧縮" in w.detail, w.items


# ====================================================================== Excel 5.0/95（BIFF5）
def _rec(t: int, body: bytes) -> bytes:
    return struct.pack("<HH", t, len(body)) + body


def make_biff5() -> bytes:
    sjis = "氏名".encode("cp932")
    xf = struct.pack("<HHHBBII", 0, 0, 0x0000, 0, 0, 0, 0)  # 保護なし（書ける）セル
    glob_head = (_rec(0x0809, struct.pack("<HHHH", 0x0500, 0x0005, 0, 0)) + _rec(0x0042, struct.pack("<H", 932))
                 + _rec(0x0031, struct.pack("<HHHHHBBBB", 200, 0, 0x7FFF, 400, 0, 0, 0, 0x80, 0) + bytes([3]) + b"Foo")
                 + _rec(0x00E0, xf))
    name = "様式".encode("cp932")
    bs_len = 4 + 6 + 1 + len(name)
    glob_tail = _rec(0x000A, b"")
    sheet_pos = len(glob_head) + bs_len + len(glob_tail)
    bound = _rec(0x0085, struct.pack("<IBB", sheet_pos, 0, 0) + bytes([len(name)]) + name)
    sheet = (_rec(0x0809, struct.pack("<HHHH", 0x0500, 0x0010, 0, 0))
             + _rec(0x0204, struct.pack("<HHHH", 0, 0, 0, len(sjis)) + sjis)
             + _rec(0x0203, struct.pack("<HHHd", 1, 0, 0, 12.5))
             + _rec(0x0006, struct.pack("<HHH", 2, 0, 0) + struct.pack("<d", 25.0) + struct.pack("<HIH", 0, 0, 3) + b"\x1e\x02\x00")
             + _rec(0x001C, struct.pack("<HHH", 0, 0, 4) + "メモ".encode("cp932"))
             + _rec(0x000A, b""))
    return build_cfb({"Book": glob_head + bound + glob_tail + sheet})


def test_biff5_reads_cells_in_codepage_and_refuses_writing():
    from todoku.xls.book import XlsWorkbook
    from todoku.xls.proof import raw_texts
    from todoku.xlsx.edit import EditError, fill

    blob = make_biff5()
    assert kind_of(blob) == "xls"
    wb = XlsWorkbook(blob, "t.xls")
    assert wb.biff == 5 and wb.sheets[0].name == "様式"
    cells = wb.sheets[0].cells
    assert cells[(1, 1)].value == "氏名" and cells[(2, 1)].value == 12.5 and cells[(3, 1)].value == 25.0
    assert cells[(3, 1)].parse_error and "BIFF5" in cells[(3, 1)].parse_error  # 数式は保存値だけ（読み違えた式で計算しない）
    assert wb.sheets[0].comments[(1, 1)] == "メモ"
    assert ("様式!A1", "氏名") in raw_texts(blob)["セルの文字"]
    with pytest.raises(EditError, match="BIFF5"):
        fill(wb, {"様式!B1": "x"})


def test_dual_format_reads_biff8_and_book_agrees():
    p = next(iter(CORPUS.glob("wide_legacy/**/*SENTEI.xls")), None) if CORPUS.exists() else None
    if p is None:
        pytest.skip("corpus がありません")
    from todoku.xls.book import XlsWorkbook

    blob = p.read_bytes()
    w8 = XlsWorkbook(blob)
    assert w8.biff == 8 and w8.dual_format
    c = CFB(blob)
    w5 = XlsWorkbook(build_cfb({"Book": c.read("Book")}))
    assert w5.biff == 5
    for s8, s5 in zip(w8.sheets, w5.sheets):
        for k, cell in s8.cells.items():
            assert s5.cells[k].value == cell.value or (cell.value in (None, "") and s5.cells[k].value in (None, ""))


# ====================================================================== Word 6.0/95
def test_word95_is_refused_clearly():
    from todoku.doc.book import DocError, DocFile

    wd = struct.pack("<HH", 0xA5DC, 101) + bytes(0x300)
    blob = build_cfb({"WordDocument": wd})
    assert kind_of(blob) == "doc"
    with pytest.raises(DocError, match="Word 6.0/95"):
        DocFile(blob)
