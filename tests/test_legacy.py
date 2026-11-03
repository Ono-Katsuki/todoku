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
    assert "一太郎" in md and "受付印" in md and "住所［P4.c2？］ ｜ ［T1.R1.C2］" in md  # 見出し「住所」の後ろも、右のマスの代わりの所として示す（消さない）
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


# ====================================================================== 一太郎：字数を変える書き込み（合成）
def _seginfo(names: dict[str, int]) -> bytes:
    """ストリームの一覧：48 バイトの見出し（16 バイト目から 項目の大きさ 96・名前のある項目の数）、名前のない項目 1 つ、
    名前のある項目（名前・68 バイト目から大きさ・種類）。"""
    head = bytearray(b"VDA_DOC\x00" + bytes(40))
    struct.pack_into("<HH", head, 16, 96, len(names))
    head[20:36] = bytes.fromhex("00000100000801000010010002080100")
    out = bytes(head) + bytes(96)
    for name, size in names.items():
        item = bytearray(96)
        raw = name.encode("utf-16-le")
        item[:len(raw)] = raw
        struct.pack_into("<II", item, 68, size, 1)
        out += bytes(item)
    return out


def _textseg_runs(text: str, runs: list) -> bytes:
    """字の並び ＋ 文字の書式の区切り（("n", 数) は「00＋4 バイトの数」、("a", バイト列) は書式を変える FE … FF 00）。"""
    u = _text_units(text)
    tail = b"".join(b"\x00" + struct.pack(">I", v) if k == "n" else v for k, v in runs)
    return b"TextV.01" + struct.pack(">I", len(u)) + struct.pack(f">{len(u)}H", *u) + tail


BODY = "申請書¶|氏名|¶⏎|住所|　　⏎¶受付　　番号¶"
BOLD, PLAIN = b"\xfe\x02\x02\x01\x61\xff\x00", b"\xfe\x02\x02\x00\x00\xff\x00"


def make_jtd_full(body: str = BODY, runs=None, extra: dict | None = None, caret: int | None = None,
                  pieces: list[tuple[int, list]] | None = None, tc: bytes | None = None) -> bytes:
    """字数を変える書き込みのための、一太郎の文書に似せた合成のファイル（行・ページの割り付けの控え・字を打つ位置・
    表の位置の表・ストリームの一覧を持つ）。pieces（[(片の単位の数, 片の書式の区切り)]）を渡すと、本文を「QLSTV.01」の
    一覧と片に分けて持つ（集めた文書 61a7b1ebaf32 と同じ形：一覧は番号 0、片は 1, 2, …）。"""
    units = _text_units(body)
    n = len(units)
    runs = runs if runs is not None else [("n", n)]
    if pieces is None:
        text = _ssmg([(0, _textseg_runs(body, runs))])
    else:
        assert sum(m for m, _r in pieces) == n
        segs, at = [], 0
        lst = b"QLSTV.01" + struct.pack(">I", len(pieces))
        for j, (m, prs) in enumerate(pieces, 1):
            part = units[at:at + m]
            segs.append((j, struct.pack(f">{m}H", *part) + b"".join(b"\x00" + struct.pack(">I", v) if k == "n" else v for k, v in prs)))
            lst += struct.pack(">II", m, j)
            at += m
        text = _ssmg([(0, lst + bytes(4))] + segs)
    ed = b"\x00\x01\x00\x01" + struct.pack(">HI", 0x2000, 28) + struct.pack(">IIIII", 0x14, 0, n - 1 if caret is None else caret, 0, 0) \
        + bytes(8) + struct.pack(">HI", 0x2001, 4) + b"\x00\x01\x02\x03"
    rows = [i for i, x in enumerate(units) if x == 0x0E]
    # 表の範囲の項目：始まり・終わり・2 バイト・左右の位置・数 1 バイト・その数のバイト・13 バイト（集めた文書と同じ形）
    if tc is None:
        tc = b"TCntV.01" + struct.pack(">HIH", 1, 1, 0x24) + struct.pack(">II", 3, rows[-1]) \
            + b"\x01\x01" + bytes(4) + b"\x01\x03" + bytes(13) + b"\xff\xff"
    streams = {"Font": b"FontV.01" + bytes(8), "LineMark": b"\x09\x02" + bytes(30), "PageMark": bytes(40), "PaperMark": bytes(20),
               "DocumentText": text, "DocumentEditStyles": ed, "DocumentTextPositionTables": _ssmg([(0, tc)])}
    streams.update(extra or {})
    streams["\x04JSRV_SegmentInformation"] = _seginfo({k: len(v) for k, v in streams.items()})
    return build_cfb(streams)


def _paras(blob: bytes) -> list[str]:
    from todoku.docx.document import iter_paras
    from todoku.jtd.book import JtdFile

    return [p.text for _k, _p, bl in JtdFile(blob).all_blocks() for p in iter_paras(bl)]


def test_jtd_grows_empty_cell_and_proves_every_structure():
    from todoku.docx.fill import written_of
    from todoku.jtd.audit import prove
    from todoku.jtd.book import JtdFile, ssmg
    from todoku.jtd.edit import parse_runs, slots_of
    from todoku.jtd.fill import fill
    from todoku.jtd.window import build

    blob = make_jtd_full()
    doc = JtdFile(blob)
    win = build(doc)
    fields = {f["id"]: f for f in win["fields"]}
    empty = [f for f in fields.values() if f["kind"] == "text" and not f.get("maxlen") and not f.get("readonly")]
    assert empty, "字のない空のマスは、字を足して書ける欄にする"
    res = fill(JtdFile(blob), {empty[0]["id"]: "山田　太郎"})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors + pr.readback
    assert any("山田　太郎" in t for t in _paras(res.blob))
    c, c0 = CFB(res.blob), CFB(blob)
    # 行・ページの割り付けの控えは取り除き、ストリームの一覧からも項目を除く
    assert not any(c.has(n) for n in ("LineMark", "PageMark", "PaperMark"))
    names = [n for _b, n in slots_of(c.read("\x04JSRV_SegmentInformation"))[1]]
    assert "LineMark" not in names and "DocumentText" in names
    # 字を打つ位置・表の範囲の終わりは、足した 5 字だけ後ろへ。範囲の始まり（書いた所より前）はそのまま
    n0 = len(_text_units(BODY))
    assert struct.unpack_from(">I", c.read("DocumentEditStyles"), 18)[0] == n0 - 1 + 5
    tc = ssmg(c.read("DocumentTextPositionTables"), "t")[0].data
    rows = [i for i, x in enumerate(_text_units(BODY)) if x == 0x0E]
    assert struct.unpack_from(">II", tc, 16) == (3, rows[-1] + 5)
    # 文字の書式の区切り：1 つの区切りが 5 単位だけ長くなる
    seg = ssmg(c.read("DocumentText"), "d")[0].data
    n1 = struct.unpack_from(">I", seg, 8)[0]
    assert n1 == n0 + 5 and parse_runs(seg[12 + 2 * n1:]) == [("n", n0 + 5)]
    # ほかのストリームは 1 バイトも変えない
    assert c.read("Font") == c0.read("Font")
    assert "LineMark" in pr.headline() and "表の位置の表" in pr.headline()


def test_jtd_grow_placeholder_and_anchored_shorter():
    from todoku.docx.fill import written_of
    from todoku.jtd.audit import prove
    from todoku.jtd.book import JtdFile
    from todoku.jtd.fill import fill
    from todoku.jtd.window import build

    blob = make_jtd_full()
    win = build(JtdFile(blob))
    blank = [f for f in win["fields"] if f["kind"] == "text" and "住所" in str(f.get("label", ""))]
    assert blank
    # 空白 2 字のマスに、長い値（前の字の書式を引き継ぐ）
    res = fill(JtdFile(blob), {blank[0]["id"]: "東京都千代田区"})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors + pr.readback
    assert "東京都千代田区" in _paras(res.blob)
    # 字を目印にして、字数を減らす・足す
    res = fill(JtdFile(blob), {"@1": {"old": "受付　　番号", "new": "受付番号"}, "@2": {"old": "申請書", "new": "交付申請書"}})
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors + pr.readback
    texts = _paras(res.blob)
    assert "受付番号" in texts and "交付申請書" in texts


def test_jtd_grow_refuses_what_is_not_established():
    from todoku.docx.fill import FillError
    from todoku.jtd.book import JtdFile
    from todoku.jtd.edit import policy
    from todoku.jtd.fill import fill
    from todoku.jtd.window import build

    # 意味を確かめていないストリームがある文書：字数を変える書き込みを断り、窓でも「N 字まで」・変更できない欄にする
    blob = make_jtd_full(extra={"NumberingLink": b"\x00\x01\x00\x02" + bytes(40)})
    pol = policy(JtdFile(blob))
    assert not pol.ok and "NumberingLink" in pol.reasons[0]
    win = build(JtdFile(blob))
    empty = [f for f in win["fields"] if f["kind"] == "text" and f.get("readonly")]
    assert empty and "NumberingLink" in empty[0]["why_fixed"]
    with pytest.raises(FillError, match="NumberingLink"):
        fill(JtdFile(blob), {empty[0]["id"]: "山田"})


def test_jtd_grow_proof_catches_tampering():
    from todoku.cfb import CfbEditor
    from todoku.docx.fill import written_of
    from todoku.jtd.audit import prove
    from todoku.jtd.book import JtdFile, ssmg
    from todoku.jtd.fill import fill
    from todoku.jtd.window import build

    blob = make_jtd_full()
    win = build(JtdFile(blob))
    fid = next(f["id"] for f in win["fields"] if f["kind"] == "text" and not f.get("maxlen") and not f.get("readonly"))
    res = fill(JtdFile(blob), {fid: "山田"})
    w = written_of(res.patches)
    assert prove(blob, res.blob, w).ok
    c = CFB(res.blob)
    # 1) 変えてはいけないストリーム
    ed = CfbEditor(res.blob)
    ed.replace("Font", b"FontV.01" + b"\x01" * 8)
    assert not prove(blob, ed.blob(), w).ok
    # 2) 文字の書式の区切りの数
    d = bytearray(c.read("DocumentText"))
    seg = ssmg(bytes(d), "d")[0]
    n = struct.unpack_from(">I", seg.data, 8)[0]
    d[seg.offset(12 + 2 * n + 4)] ^= 1
    ed = CfbEditor(res.blob)
    ed.replace("DocumentText", bytes(d))
    pr = prove(blob, ed.blob(), w)
    assert not pr.ok and any("書式" in x for x in pr.unexpected)
    # 3) 取り除いたストリームの元のデータが、空きのミニセクタに残っている
    c0 = CFB(blob)
    e = c0._find("LineMark")
    out = bytearray(res.blob)
    root = c.chain(c.entries[0].start)
    m = e.start
    off = (root[m * 64 // 512] + 1) * 512 + (m * 64) % 512
    out[off:off + 4] = b"\x09\x02\x00\x01"
    assert not prove(blob, bytes(out), w).ok


def _fill_prove(blob: bytes, values: dict):
    from todoku.docx.fill import written_of
    from todoku.jtd.audit import prove
    from todoku.jtd.book import JtdFile
    from todoku.jtd.fill import fill

    res = fill(JtdFile(blob), values)
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors + pr.readback
    return res, pr


def _grow_field(blob: bytes) -> str:
    from todoku.jtd.book import JtdFile
    from todoku.jtd.window import build

    win = build(JtdFile(blob))
    return next(f["id"] for f in win["fields"] if f["kind"] == "text" and not f.get("maxlen") and not f.get("readonly"))


def test_jtd_reads_and_writes_body_split_into_pieces():
    """「QLSTV.01」の一覧と片：記録の途中で分けた本文を、つないで読み、上書き・字数を変える書き込みを証明する。"""
    from todoku.jtd.book import JtdFile, ssmg

    units = _text_units(BODY)
    rec = next(i for i, x in enumerate(units) if x == 0x1C)
    cut = rec + 5  # マスの記録の途中で分ける
    n = len(units)
    blob = make_jtd_full(pieces=[(cut, [("n", cut)]), (n - cut, [("n", n - cut)])])
    whole = make_jtd_full()
    assert _paras(blob) == _paras(whole)
    doc = JtdFile(blob)
    st = [t for t in doc.texts if t.stream == "DocumentText"]
    assert len(st) == 1 and [pc.n for pc in st[0].pieces] == [cut, n - cut]
    # 字数を変えない上書き（片の中の 2 バイトだけ）と、字数を変える書き込み
    _fill_prove(blob, {"@1": {"old": "受付　　番号", "new": "受付ＡＢ番号"}})
    res, pr = _fill_prove(blob, {_grow_field(blob): "山田　太郎"})
    assert any("山田　太郎" in t for t in _paras(res.blob))
    segs = {s.sid: s.data for s in ssmg(CFB(res.blob).read("DocumentText"), "d")}
    entries = [struct.unpack_from(">II", segs[0], 12 + 8 * k) for k in range(2)]
    assert entries == [(cut, 1), (n - cut + 5, 2)]  # 字を足した片の数だけ増える（分かれ目はそのまま）


def _long_body(k: int) -> str:
    return "申請書¶|氏名|¶⏎|住所|　　⏎¶" + "".join(f"第{i:04d}項の説明の文です。¶" for i in range(k)) + "受付　　番号¶"


def _big_body() -> str:
    """字の並び（書式の区切りは 40 バイトまで）が、1 つのセグメントで確かめた大きさ（42,466 バイト）のすぐ下になる本文。"""
    from todoku.jtd.edit import MAX_SEG_BYTES

    base = len(_text_units(_long_body(0)))
    return _long_body((MAX_SEG_BYTES - 12 - 40 - 2 * base) // (2 * len(_text_units(_long_body(1))) - 2 * base))


def test_jtd_splits_large_body_into_pieces_and_restates_format():
    """1 つのセグメントで確かめた大きさを超える本文は、「QLSTV.01」の一覧と、等しい大きさの片に分ける。分ける所より前の
    書式は、後ろの片の頭で言い直す（4 つの読み方のどれでも、どの字の書式も元のまま）。"""
    from todoku.jtd.book import JtdFile, ssmg
    from todoku.jtd.edit import MAX_PIECE_UNITS, MAX_SEG_BYTES
    from todoku.jtd.proof import MODES, _body_raw, _formats

    body = _big_body()
    units = _text_units(body)
    n = len(units)
    a = n // 3
    bold = b"\xfe\x02\x02\x01\x61\x14\x04\x80\x00\x00\x00\xff\x00"
    runs = [("n", a), ("a", bold), ("n", 7), ("a", PLAIN), ("n", n - a - 7)]
    blob = make_jtd_full(body, runs=runs)
    seg = ssmg(CFB(blob).read("DocumentText"), "d")[0].data
    assert len(seg) <= MAX_SEG_BYTES < len(seg) + 200
    res, pr = _fill_prove(blob, {_grow_field(blob): "山田　太郎" * 20})
    x = CFB(res.blob).read("DocumentText")
    kind, pieces, sid = _body_raw(x)
    assert kind == "list" and len(pieces) == 2 and sid == 0
    assert all(len(p[1]) <= MAX_PIECE_UNITS for p in pieces) and abs(len(pieces[0][1]) - len(pieces[1][1])) <= 1
    # 後ろの片の頭で、前の書式（重ねた値と、最後の区切り）を言い直している
    assert pieces[1][2][0][0] == "a" and pieces[1][2][1] == ("a", PLAIN)
    assert "分けて持つ形" in pr.headline()
    # 4 つの読み方のどれでも、書いた所の後ろの字の書式は元と同じ（証明でも確かめている）
    old = _body_raw(CFB(blob).read("DocumentText"))[1]
    for mode in MODES:
        assert _formats(old, mode)[-50:] == _formats(pieces, mode)[-50:]
    assert any("受付　　番号" in t for t in _paras(res.blob))
    # 分けた本文に、もう一度書ける（読み直して、片の分かれ目をそのままに書く）
    blob2 = res.blob
    _fill_prove(blob2, {"@1": {"old": "受付　　番号", "new": "受付番号"}})
    assert JtdFile(blob2).texts


def test_jtd_split_proof_catches_missing_restatement():
    from todoku.cfb import CfbEditor
    from todoku.docx.fill import written_of
    from todoku.jtd.audit import prove
    from todoku.jtd.book import JtdFile
    from todoku.jtd.edit import runs_bytes, ssmg_update
    from todoku.jtd.fill import fill
    from todoku.jtd.proof import _body_raw

    body = _big_body()
    n = len(_text_units(body))
    a = n // 3
    blob = make_jtd_full(body, runs=[("n", a), ("a", BOLD), ("n", n - a)])
    res = fill(JtdFile(blob), {_grow_field(blob): "山田　太郎" * 20})
    w = written_of(res.patches)
    assert prove(blob, res.blob, w).ok
    # 後ろの片の頭の言い直しを外す：「初めに戻る」読み方で、後ろの片の書式が変わる
    x = CFB(res.blob).read("DocumentText")
    _kind, pieces, _sid = _body_raw(x)
    sid, us, els = pieces[1]
    assert els[0] == ("a", BOLD)
    data = struct.pack(f">{len(us)}H", *us) + runs_bytes(els[1:])
    y, _info = ssmg_update(x, {sid: data})
    ed = CfbEditor(res.blob)
    ed.replace("DocumentText", y)
    pr = prove(blob, ed.blob(), w)
    assert not pr.ok and any("言い直して" in m or "書式" in m for m in pr.unexpected)


def test_jtd_grow_updates_marks_corrections_and_keeps_numbering():
    """しおりの位置（表の位置の表の MarkV）・校正の情報の位置をずらし、段落番号のつながり（番号でつながる）は変えない。"""
    from todoku.jtd.book import JtdFile, ssmg
    from todoku.jtd.edit import policy

    numbering = "".join(map(chr, [0x1C, 0, 13, 0, 0x0A, 0, 0x187, 0x2010, 0, 13, 0, 0, 0x1F]))
    correct = "".join(map(chr, [0x1C, 0, 10, 0, 0x41, 0x81, 10, 0, 0, 0x1F]))
    body = numbering + "申請書¶|氏名|¶⏎|住所|　　⏎¶受付" + correct + "　　番号¶"
    units = _text_units(body)
    corr_at = next(i for i in range(len(units)) if units[i] == 0x1C and units[i + 4] == 0x41)
    mark_at = units.index(ord("番"))
    nl = struct.pack(">HHIH", 1, 2, 44, 1) + struct.pack(">HHHHIHHHHHH", 0, 0x18, 0xFFFF, 0xFFFF, 1980, 0x100, 3, 1, 1, 0xFFFF, 0) \
        + bytes(6) + b"\xff\xff\x00\x00"
    cad = struct.pack(">HHIHH", 1, 6, 0, 2, 0) + struct.pack(">HHI", 0x81, 0, corr_at) + bytes(24) \
        + struct.pack(">HHI", 0x83, 0, 0xFFFFFFFF) + bytes(24) + bytes(20)
    tc = b"TCntV.01" + b"\x00\x00" + b"MarkV.01" + b"\x00\x00" + struct.pack(">II", mark_at, 0xFFFFFFFF)
    mt = b"MarkV.01" + struct.pack(">HIH", 1, 0, 7) + "_GoBack".encode("utf-16-be")
    blob = make_jtd_full(body, tc=tc, extra={"NumberingLink": nl, "CorrectAdditionalData": _ssmg([(1, cad)]), "MarkTag": mt})
    assert policy(JtdFile(blob)).ok, policy(JtdFile(blob)).reasons
    res, pr = _fill_prove(blob, {_grow_field(blob): "山田"})
    c = CFB(res.blob)
    t = ssmg(c.read("DocumentTextPositionTables"), "t")[0].data
    assert struct.unpack_from(">II", t, 20) == (mark_at + 2, 0xFFFFFFFF)
    d = ssmg(c.read("CorrectAdditionalData"), "c")[0].data
    assert struct.unpack_from(">I", d, 16)[0] == corr_at + 2 and struct.unpack_from(">I", d, 48)[0] == 0xFFFFFFFF
    assert c.read("NumberingLink") == nl and c.read("MarkTag") == mt
    assert "しおりの位置" in pr.headline() and "校正の情報" in pr.headline()
    # 段落番号のつながりが、本文の部品を番号で指していなければ（形を確かめられない）断る
    bad = nl[:32] + b"\x00\x09" + nl[34:]
    assert not policy(JtdFile(make_jtd_full(body, tc=tc, extra={"NumberingLink": bad, "MarkTag": mt}))).ok
    # 校正の情報の位置が、同じ番号の部品の記録の始まりでなければ断る
    cad2 = cad[:16] + struct.pack(">I", corr_at + 1) + cad[20:]
    pol = policy(JtdFile(make_jtd_full(body, tc=tc, extra={"CorrectAdditionalData": _ssmg([(1, cad2)]), "MarkTag": mt})))
    assert not pol.ok and "校正の情報" in pol.reasons[0]


def test_jtd_grow_allows_run_boundary_with_the_same_format():
    """書式の区切りの境でも、前と後ろの書式がどの読み方でも同じなら、字を足せる（同じ書式を 2 度言う区切り）。"""
    units = _text_units(BODY)
    cut = next(i for i in range(1, len(units)) if units[i] == 0x0A and units[i - 1] == 0x1F)
    blob = make_jtd_full(runs=[("n", 1), ("a", BOLD), ("n", cut - 1), ("a", BOLD), ("n", len(units) - cut), ("a", PLAIN)])
    res, _pr = _fill_prove(blob, {_grow_field(blob): "山田"})
    assert any("山田" in t for t in _paras(res.blob))


# 種類 0x14 だけを変える区切り（BOLD の後ろに置くと、重ねた値と最後の区切りが違う）
MARK = b"\xfe\x14\x04\x80\x00\x00\x00\xff\x00"


def _cell_end() -> int:
    """BODY の、マスの記録のすぐ後ろの段落の終わり（空のマス）の単位の位置。"""
    units = _text_units(BODY)
    return next(i for i in range(1, len(units)) if units[i] == 0x0A and units[i - 1] == 0x1F)


def _inserted_formats(old_blob: bytes, new_blob: bytes, q: int, count: int) -> None:
    """4 つの読み方のどれでも、足した字（新しい本文の q から count 個）が、元の q の単位（段落の終わりの印）と同じ書式で、
    ほかの字は元の書式のまま。"""
    from todoku.jtd.proof import MODES, _body_raw, _formats

    old = _body_raw(CFB(old_blob).read("DocumentText"))[1]
    new = _body_raw(CFB(new_blob).read("DocumentText"))[1]
    for mode in MODES:
        fa, fb = _formats(old, mode), _formats(new, mode)
        assert fb[q:q + count] == [fa[q]] * count, mode
        assert fb[:q] == fa[:q] and fb[q + count:] == fa[q:], mode


@pytest.mark.parametrize("runs_of, restated", [
    # 段落の終わりの印から太字（前の記録は初めの書式）：言い直すのは最後の区切りだけ
    (lambda cut, n: [("n", cut), ("a", BOLD), ("n", n - cut), ("a", PLAIN)], [("a", BOLD)]),
    # 区切りを重ねた値と最後の区切りが違う：重ねた値を 1 つにまとめた区切りと、最後の区切り
    (lambda cut, n: [("n", 1), ("a", BOLD), ("n", cut - 1), ("a", MARK), ("n", n - cut)],
     [("a", b"\xfe\x02\x02\x01\x61\x14\x04\x80\x00\x00\x00\xff\x00"), ("a", MARK)]),
    # 段落の終わりの印が、区切りの数の合計より後ろ（最後の書式のまま）
    (lambda cut, n: [("n", cut), ("a", BOLD)], [("a", BOLD)]),
])
def test_jtd_grow_at_format_boundary_restates_end_mark_format(runs_of, restated):
    """書式の切り替わる所（前の記録と段落の終わりの印の書式が違う、空のマス）に足す字は、それだけの区切りにして、
    段落の終わりの印の書式をすべて言い直して書く。どの読み方でも、足した字は終わりの印の書式で、ほかの字は元のまま。"""
    from todoku.jtd.book import JtdFile, ssmg
    from todoku.jtd.edit import parse_runs
    from todoku.jtd.window import build

    cut = _cell_end()
    n = len(_text_units(BODY))
    runs = runs_of(cut, n)
    blob = make_jtd_full(runs=runs)
    from todoku.grade import grade_blob

    win = build(JtdFile(blob))
    fid = _grow_field(blob)
    assert all("書式" not in f.get("why_fixed", "") for f in win["fields"])
    # 段階でも、書式の切り替わる所を書けない理由にしない
    g = grade_blob(blob, "boundary.jtd")
    assert g.write == "B" and not any("書式" in r.detail for r in g.reasons if r.grade == "D")
    res, pr = _fill_prove(blob, {fid: "山田"})
    assert any("山田" in t for t in _paras(res.blob))
    seg = ssmg(CFB(res.blob).read("DocumentText"), "d")[0].data
    n1 = struct.unpack_from(">I", seg, 8)[0]
    got = parse_runs(seg[12 + 2 * n1:])
    # 足す所の区切りの前に、言い直した書式と足した字の数。その後ろは元の区切りのまま
    j = next(i for i, (k, _v) in enumerate(runs) if k == "a" and sum(v for kk, v in runs[:i] if kk == "n") == cut)
    assert got == runs[:j] + restated + [("n", 2)] + runs[j:]
    _inserted_formats(blob, res.blob, cut, 2)
    assert any("書式の切り替わる所" in t for t in res.notes)
    assert "書式の切り替わる所に足した 1 か所" in pr.headline()


def test_jtd_grow_boundary_proof_catches_unrestated_and_wrong_format():
    """証明は、書式の切り替わる所に足した字が、言い直した区切りになっていること（書式が結果として同じでも）と、どの読み方
    でも決めた書式であることを確かめる。"""
    from todoku.cfb import CfbEditor
    from todoku.docx.fill import written_of
    from todoku.jtd.audit import prove
    from todoku.jtd.book import JtdFile, ssmg
    from todoku.jtd.edit import parse_runs, runs_bytes, ssmg_update
    from todoku.jtd.fill import fill

    cut = _cell_end()
    n = len(_text_units(BODY))
    runs = [("n", cut), ("a", BOLD), ("n", n - cut), ("a", PLAIN)]
    blob = make_jtd_full(runs=runs)
    res = fill(JtdFile(blob), {_grow_field(blob): "山田"})
    w = written_of(res.patches)
    assert prove(blob, res.blob, w).ok
    x = CFB(res.blob).read("DocumentText")
    seg = ssmg(x, "d")[0]
    n1 = struct.unpack_from(">I", seg.data, 8)[0]
    head = seg.data[:12 + 2 * n1]

    def swap(els) -> bytes:
        y, _info = ssmg_update(x, {seg.sid: head + runs_bytes(els)})
        ed = CfbEditor(res.blob)
        ed.replace("DocumentText", y)
        return ed.blob()

    assert parse_runs(seg.data[12 + 2 * n1:]) == [("n", cut), ("a", BOLD), ("n", 2), ("a", BOLD), ("n", n - cut), ("a", PLAIN)]
    # 1) 言い直さずに、終わりの印の区切りを増やしただけ（どの読み方でも書式は同じになるが、決めた形ではない）
    pr = prove(blob, swap([("n", cut), ("a", BOLD), ("n", n - cut + 2), ("a", PLAIN)]), w)
    assert not pr.ok and any("区切り" in m for m in pr.unexpected)
    # 2) 足した字を、前の記録の書式にした
    pr = prove(blob, swap([("n", cut + 2), ("a", BOLD), ("n", n - cut), ("a", PLAIN)]), w)
    assert not pr.ok and any("足した字の書式" in m for m in pr.unexpected)
    # 3) 言い直した区切りのバイト列が違う
    pr = prove(blob, swap([("n", cut), ("a", MARK), ("n", 2), ("a", BOLD), ("n", n - cut), ("a", PLAIN)]), w)
    assert not pr.ok and any("足した字の書式" in m for m in pr.unexpected)


def test_jtd_grow_boundary_in_split_body_pieces():
    """分けて持つ本文（片）の中の、書式の切り替わる所に足す：言い直した区切りは、その片の区切りに入る。"""
    from todoku.jtd.proof import _body_raw

    units = _text_units(BODY)
    n = len(units)
    cut = _cell_end()
    first = cut - 3  # マスの記録の途中で分ける。空のマスは後ろの片
    blob = make_jtd_full(pieces=[(first, [("n", first)]), (n - first, [("n", 3), ("a", BOLD), ("n", n - cut), ("a", PLAIN)])])
    res, pr = _fill_prove(blob, {_grow_field(blob): "山田"})
    pieces = _body_raw(CFB(res.blob).read("DocumentText"))[1]
    assert pieces[1][2] == [("n", 3), ("a", BOLD), ("n", 2), ("a", BOLD), ("n", n - cut), ("a", PLAIN)]
    _inserted_formats(blob, res.blob, cut, 2)
    assert "書式の切り替わる所に足した 1 か所" in pr.headline()


def _jscomp(native: bytes) -> bytes:
    from todoku.jtd.lha import crc16, lh5

    packed = lh5(native)
    name = b"C:\\Temp\\TempNativeFile"
    h = bytearray(b"\x00\x00-lh5-" + struct.pack("<II", len(packed), len(native)) + bytes(4) + b"\x20\x01"
                  + bytes([len(name)]) + name + struct.pack("<H", crc16(native)) + b"\x20\x00\x00")
    h[0] = len(h) - 2
    h[1] = sum(h[2:]) & 0xFF
    prefix = struct.pack("<H", 38) + b"JustCompressedDocument" + bytes(4) + struct.pack("<IIH", 15716028, 70970, 1)
    return prefix + bytes(h) + packed + b"\x00"


def test_jtd_compressed_document_is_recompressed_and_proved():
    from todoku.docx.fill import written_of
    from todoku.jtd.audit import prove
    from todoku.jtd.book import JtdFile
    from todoku.jtd.fill import fill
    from todoku.jtd.lha import jscomp_parts, unpack_jscomp
    from todoku.jtd.window import build

    native = make_jtd_full()
    stream = _jscomp(native)
    outer = build_cfb({"JSCompDocument": stream, "\x04JSRV_SegmentInformation": _seginfo({"JSCompDocument": len(stream)})})
    doc = JtdFile(outer)
    assert doc.compressed and doc.cfb.blob == native
    win = build(doc)
    grow = next(f["id"] for f in win["fields"] if f["kind"] == "text" and not f.get("maxlen") and not f.get("readonly"))
    for values in ({grow: "山田太郎"}, {"@1": {"old": "受付　　番号", "new": "受付ＡＢ番号"}}):  # 字数を変える・変えない
        res = fill(JtdFile(outer), values)
        pr = prove(outer, res.blob, written_of(res.patches))
        assert pr.ok, pr.unexpected + pr.lint_errors + pr.readback
        new_stream = CFB(res.blob).read("JSCompDocument")
        inner = unpack_jscomp(new_stream)  # CRC-16 と長さで確かめて解く
        assert JtdFile(inner).cfb.blob == inner
        assert jscomp_parts(new_stream)["prefix"] == stream[:38]
        assert "-lh5-" in pr.headline()


def test_lh5_encoder_roundtrip():
    import random

    from todoku.jtd.lha import lh5, unlh5

    rnd = random.Random(7)
    samples = [b"x", b"ab" * 5000, bytes(rnd.randrange(256) for _ in range(6000)),
               bytes(rnd.choice(b"\x00\x01ab") for _ in range(70000)), make_jtd_full()]
    for data in samples:
        assert unlh5(lh5(data), len(data)) == data


def test_cfb_editor_removes_streams_and_grows_mini_stream():
    from todoku.cfb import CfbEditor
    from todoku.jtd.proof import check_container

    streams = {f"S{k:02d}": bytes([k]) * (100 + k) for k in range(12)}
    streams["Big"] = b"B" * 5000
    blob = build_cfb(streams)
    ed = CfbEditor(blob)
    ed.replace("S03", b"3" * 5000)  # ミニストリームから、ふつうのセクタへ
    ed.replace("S04", b"4" * 900)  # ミニストリームの中で大きくなる
    ed.remove("S07")
    ed.remove("S08")
    out = ed.blob()
    c = CFB(out)
    assert c.read("S03") == b"3" * 5000 and c.read("S04") == b"4" * 900 and not c.has("S07") and not c.has("S08")
    assert all(c.read(k) == v for k, v in streams.items() if k not in ("S03", "S04", "S07", "S08"))
    assert check_container(blob, out, {"S03", "S04"}, {"S07", "S08"}) == []
    assert bytes([7]) * 107 not in out  # 取り除いたストリームのデータは残らない


JTD = sorted(CORPUS.glob("wide_jtd/**/*.jtd")) if CORPUS.exists() else []


@pytest.mark.corpus
@pytest.mark.parametrize("path", JTD[:40], ids=lambda p: p.name)
def test_jtd_corpus_reading_and_overwrite(path):
    from todoku.jtd.audit import run

    au = run(path.read_bytes(), path.name)
    checks = {c.name: c for c in au.checks}
    assert checks["読み落とし（文書の文字）"].ok, checks["読み落とし（文書の文字）"].items
    assert checks["ストリームの一覧"].ok, checks["ストリームの一覧"].items
    w = checks["全入力欄への書き込み"]
    assert w.ok, w.items  # 圧縮して保存した文書も、縮め直して書く
    assert checks["読み戻し"].ok, checks["読み戻し"].items
    g = checks.get("字数を変える、字を目印にした書き込み")
    assert g is None or g.ok, g.items


GROW = [p for p in JTD if p.name.split("_", 1)[1] in ("001188742.jtd", "1242588_001.jtd", "001189440.jtd")]


@pytest.mark.corpus
@pytest.mark.parametrize("path", GROW, ids=lambda p: p.name)
def test_jtd_corpus_fills_empty_cells(path):
    """集めた様式の、字のない空のマス（これまで書けなかった欄）に字を足し、証明・読み直しを確かめる。"""
    from todoku.docx.fill import written_of
    from todoku.jtd.audit import prove
    from todoku.jtd.book import JtdFile
    from todoku.jtd.fill import fill
    from todoku.jtd.window import build

    blob = path.read_bytes()
    win = build(JtdFile(blob))
    empty = [f["id"] for f in win["fields"] if f["kind"] == "text" and not f.get("readonly") and not f.get("maxlen")
             and not f.get("value")][:6]
    assert empty
    values = {fid: f"記入{k}" for k, fid in enumerate(empty)}
    res = fill(JtdFile(blob), values)
    assert res.ops, "字数を変える書き込みになる"
    pr = prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors + pr.readback
    texts = _paras(res.blob)
    assert all(any(v in t for t in texts) for v in values.values())
    assert not CFB(JtdFile(res.blob).cfb.blob).has("LineMark")


# 集めた様式の、書式の切り替わる所にある欄（前の記録と段落の終わりの印、または段落の頭の字の書式が違う）。
# 以前は足す字の書式を決められないとして断っていた所
BOUNDARY = [("057e70ea353d", "T2.R5.C4"), ("8cd333a10d8a", "T2.R3.C2"), ("056d416c7448", "T2.R3.C2"),
            ("4cede3b62917", "T1.R3.C3"), ("61a7b1ebaf32", "T1.R50.C1"), ("de1a26940ff8", "T4.R18.C20")]
# 以前の ("5c387711bc76", "P72.a1")（「円」だけのマスの、円の前）は、字の意味（単位の語）から決めていた所で、
# 字の形と並びだけで決める窓では示さない（書くときは字を目印にする）ので外した


@pytest.mark.corpus
@pytest.mark.parametrize("prefix, fid", BOUNDARY, ids=lambda x: x)
def test_jtd_corpus_fills_at_format_boundary(prefix, fid):
    """書式の切り替わる所の欄に字を足す：証明が通り、4 つの読み方のどれでも、足した字は足す所の単位（段落の終わりの印、
    段落の頭なら頭の字）の書式で、ほかの字は元の書式のまま。"""
    from todoku.jtd.book import JtdFile
    from todoku.jtd.proof import MODES, _body_raw, _formats
    from todoku.jtd.window import build

    path = next((p for p in JTD if p.name.startswith(prefix)), None)
    if path is None:
        pytest.skip("集めた様式（corpus/wide_jtd）がありません")
    blob = path.read_bytes()
    f = next(f for f in build(JtdFile(blob))["fields"] if f["id"] == fid)
    assert not f.get("readonly") and not f.get("maxlen"), f.get("why_fixed")
    res, pr = _fill_prove(blob, {fid: "記入した字"})
    assert len(res.ops) == 1
    op = res.ops[0]
    q, count = op.a, len(op.text)
    old = _body_raw(CFB(JtdFile(blob).cfb.blob).read("DocumentText"))[1]
    new = _body_raw(CFB(JtdFile(res.blob).cfb.blob).read("DocumentText"))[1]
    border = False
    for mode in MODES:
        fa, fb = _formats(old, mode), _formats(new, mode)
        border |= fa[q - 1] != fa[q]
        assert fb[q:q + count] == [fa[q]] * count, mode
        assert fb[:q] == fa[:q] and fb[q + count:] == fa[q:], mode
    assert border, "書式の切り替わる所の欄"
    assert "書式の切り替わる所に足した 1 か所" in pr.headline()
    assert any("記入した字" in t for t in _paras(res.blob))


# ====================================================================== Excel 5.0/95（BIFF5）
def _rec(t: int, body: bytes) -> bytes:
    return struct.pack("<HH", t, len(body)) + body


def _formula(r: int, c: int, cached, rgce: bytes, extra: bytes = b"") -> bytes:
    res = struct.pack("<d", cached) if isinstance(cached, float) else bytes([0, 0, 0, 0, 0, 0, 0xFF, 0xFF])
    return _rec(0x0006, struct.pack("<HHH", r, c, 0) + res + struct.pack("<HIH", 0, 0, len(rgce)) + rgce + extra)


def _ref(r: int, c: int, rel: bool = True) -> bytes:
    """tRefV（BIFF5）：行（相対の印つき）2 バイト・列 1 バイト。"""
    return b"\x44" + struct.pack("<HB", (0xC000 if rel else 0) | r, c)


def _sheet5(rows: dict[int, list[bytes]], links: bytes = b"", dims=(0, 10, 0, 6), base: int = 0) -> bytes:
    """BIFF5 のシート：BOF・INDEX・Link Table・DEFCOLWIDTH・DIMENSIONS・行のまとまり（ROW・セル・DBCELL）・EOF。
    INDEX と DBCELL の位置（ストリームの頭から。base はシートの BOF の位置）は、ここで（書き込み処理とは別に）数えて入れる。
    行は 32 行ずつのまとまり。"""
    head = _rec(0x0809, struct.pack("<HHHH", 0x0500, 0x0010, 0, 0))
    rws = sorted(rows)
    blocks: list[list[int]] = []
    for r in rws:
        if blocks and r // 32 == blocks[-1][0] // 32:
            blocks[-1].append(r)
        else:
            blocks.append([r])
    index_len = 4 + 12 + 4 * len(blocks)
    pre = links + _rec(0x0055, struct.pack("<H", 8))
    dim = _rec(0x0200, struct.pack("<HHHHH", *dims, 0))
    pos = base + len(head) + index_len
    defcol_at = pos + len(links)
    pos += len(pre) + len(dim)
    body = b""
    dbcell_at = []
    for blk in blocks:
        start = pos + len(body)
        row_recs = b"".join(_rec(0x0208, struct.pack("<HHHHHHI", r, 0, 6, 255, 0, 0, 0x100)) for r in blk)
        cells_pos = []
        cur = start + len(row_recs)
        cells = b""
        for r in blk:
            cells_pos.append(cur + len(cells))
            cells += b"".join(rows[r])
        here = start + len(row_recs) + len(cells)
        offs = []
        ref = start + 20
        for p in cells_pos:
            offs.append(p - ref)
            ref = p
        db = _rec(0x00D7, struct.pack("<I", here - start) + b"".join(struct.pack("<H", o) for o in offs))
        dbcell_at.append(here)
        body += row_recs + cells + db
    index = _rec(0x020B, struct.pack("<IHHI", 0, rws[0], rws[-1] + 1, defcol_at) + b"".join(struct.pack("<I", d) for d in dbcell_at))
    return head + index + pre + dim + body + _rec(0x000A, b"")


def make_biff5(extra_cells: bool = True) -> bytes:
    """様式（cp932）とデータ（別のシート）の 2 枚のブック。数式は保存値つき。"""
    xf = struct.pack("<HHHBBII", 0, 0, 0x0000, 0, 0, 0, 0)  # 保護なし（書ける）セル
    names = ["様式".encode("cp932"), b"Data"]
    glob_head = (_rec(0x0809, struct.pack("<HHHH", 0x0500, 0x0005, 0, 0)) + _rec(0x0042, struct.pack("<H", 932))
                 + _rec(0x0016, struct.pack("<H", 1)) + _rec(0x0017, bytes([len(names[0]), 3]) + names[0])
                 # 印刷範囲（名前）：様式!$A$1:$C$5（tArea3d、ブック全体の EXTERNSHEET の 1 番）
                 + _rec(0x0018, struct.pack("<HBBHHHBBBB", 0x20, 0, 1, 21, 0, 1, 0, 0, 0, 0) + b"\x06"
                        + b"\x3b" + struct.pack("<h", -1) + bytes(8) + struct.pack("<HHHHBB", 0, 0, 0, 4, 0, 2))
                 + _rec(0x0031, struct.pack("<HHHHHBBBB", 200, 0, 0x7FFF, 400, 0, 0, 0, 0x80, 0) + bytes([3]) + b"Foo")
                 + _rec(0x00E0, xf))
    sjis = "氏名".encode("cp932")
    s_tok = b"\x17" + bytes([len(sjis)]) + sjis
    ref3d = b"\x5a" + struct.pack("<h", -1) + bytes(8) + struct.pack("<HHHB", 1, 1, 0, 0)  # Data!$A$1
    rows = {
        0: [_rec(0x0204, struct.pack("<HHHH", 0, 0, 0, len(sjis)) + sjis), _rec(0x0201, struct.pack("<HHH", 0, 1, 0))],
        1: [_rec(0x0203, struct.pack("<HHHd", 1, 0, 0, 12.5))],
        2: [_formula(2, 0, 25.0, _ref(1, 0) + b"\x1e\x02\x00\x05"),
            _formula(2, 1, 19.5, _ref(1, 0) + ref3d + b"\x03"),
            _formula(2, 2, 1.0, _ref(0, 0, rel=False) + s_tok + b"\x0b\x1e\x01\x00\x1e\x00\x00\x42\x03\x01\x00"),
            _formula(2, 3, 3.0, b"\x60" + bytes(7) + b"\x42\x01\x04\x00",
                     bytes([3, 1, 0]) + b"".join(b"\x01" + struct.pack("<d", 1.0) for _ in range(3)))],
        # 文字の結果（STRING は 16 ビットの長さのバイト列）
        3: [_formula(3, 0, "s", _ref(0, 0) + s_tok + b"\x08"),
            _rec(0x0207, struct.pack("<H", 8) + "氏名氏名".encode("cp932"))],
    }
    if extra_cells:
        rows[4] = [_rec(0x0201, struct.pack("<HHH", 4, 0, 0))]
    links = _rec(0x0016, struct.pack("<H", 1)) + _rec(0x0017, bytes([len(names[1]), 3]) + names[1])
    glob_tail = _rec(0x000A, b"")
    bs_len = sum(4 + 4 + 2 + 1 + len(n) for n in names)
    pos = len(glob_head) + bs_len + len(glob_tail)
    s1 = _sheet5(rows, links, base=pos)
    s2 = _sheet5({0: [_rec(0x0203, struct.pack("<HHHd", 0, 0, 0, 7.0))]}, dims=(0, 1, 0, 1), base=pos + len(s1))
    bound = (_rec(0x0085, struct.pack("<IBB", pos, 0, 0) + bytes([len(names[0])]) + names[0])
             + _rec(0x0085, struct.pack("<IBB", pos + len(s1), 0, 0) + bytes([len(names[1])]) + names[1]))
    return build_cfb({"Book": glob_head + bound + glob_tail + s1 + s2})


def test_biff5_reads_cells_and_formulas_in_codepage():
    from todoku.xls.book import XlsWorkbook
    from todoku.xls.proof import lint, raw_texts
    from todoku.xlsx.calc import verify

    blob = make_biff5()
    assert kind_of(blob) == "xls" and not lint(blob).errors
    wb = XlsWorkbook(blob, "t.xls")
    assert wb.biff == 5 and [s.name for s in wb.sheets] == ["様式", "Data"]
    cells = wb.sheets[0].cells
    assert cells[(1, 1)].value == "氏名" and cells[(2, 1)].value == 12.5
    got = {k: (c.formula, c.parse_error) for k, c in cells.items() if c.has_formula}
    assert got == {(3, 1): ("A2*2", None), (3, 2): ("A2+Data!$A$1", None), (3, 3): ('IF($A$1="氏名",1,0)', None),
                   (3, 4): ("SUM({1,1,1})", None), (4, 1): ('A1&"氏名"', None)}, got
    assert cells[(4, 1)].value == "氏名氏名"
    # 計算し直すと、Excel の保存値と全数で一致する
    assert verify(wb).summary() == {"total": 5, "match": 5, "mismatch": 0, "unsupported": 0, "volatile": 0, "no_cache": 0, "cycle": 0}
    assert wb.sheets[0].print_area == "様式!$A$1:$C$5"
    assert ("様式!A1", "氏名") in raw_texts(blob)["セルの文字"]


@pytest.mark.parametrize("rgce, why", [
    (b"\x5a" + struct.pack("<h", -1) + bytes(8) + struct.pack("<HHHB", 0, 0, 0, 0), "食い違"),  # EXTERNSHEET は Data、番号は様式
    (b"\x5a" + struct.pack("<h", 3) + bytes(12) + struct.pack("<HB", 0, 0), "別のブック"),
    (b"\x5a" + struct.pack("<h", -1) + bytes(8) + struct.pack("<HHHB", 0, 1, 0, 0), "複数シート"),
    (b"\x44\x00\x00", "壊れ"),  # 番地が途中で終わる
    (b"\x1e\x01\x00\x41\x68\x01", "関数の番号"),  # 360（PHONETIC）は BIFF8 で足された関数
    (b"\x60" + bytes(7), "追加データ"),  # 配列の値がない
])
def test_biff5_refuses_formulas_it_cannot_establish(rgce, why):
    from todoku.formula.values import Unsupported
    from todoku.xls.book import XlsWorkbook
    from todoku.xls.ptg5 import Decoder5

    wb = XlsWorkbook(make_biff5())
    links = [bytes([3]) + b"Data"]
    with pytest.raises(Unsupported):
        Decoder5(wb, "cp932", links, 0).decode(rgce)
    del why


def test_biff5_write_label_number_and_dependent_formulas_with_proof():
    from todoku.xls.book import XlsWorkbook, records
    from todoku.xls.proof import lint
    from todoku.xlsx.calc import verify
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove

    blob = make_biff5()
    wb = XlsWorkbook(blob, "t.xls")
    res = fill(wb, {"様式!B1": "山田　太郎", "様式!A2": 20, "様式!A5": "東京都"})
    assert {(p.address, p.kind) for p in res.patches} >= {("B1", "input"), ("A2", "input"), ("A3", "formula"), ("B3", "formula")}
    after = XlsWorkbook(res.blob)
    c = after.sheets[0].cells
    assert c[(1, 2)].value == "山田　太郎" and c[(2, 1)].value == 20 and c[(5, 1)].value == "東京都"
    assert c[(3, 1)].value == 40 and c[(3, 2)].value == 27  # 依存する数式の保存値も書き直した
    assert verify(after).summary()["mismatch"] == 0
    # 字は LABEL（cp932 のバイト列）で書いた
    labels = [b for _p, t, b in records(after.data) if t == 0x0204]
    assert any("山田　太郎".encode("cp932") in b for b in labels)
    assert not lint(res.blob).errors
    pr = prove(blob, res.blob, res.allowed_inputs)
    assert pr.ok, pr.unexpected + pr.lint_errors
    assert pr.entries_identical == pr.entries_total - 1
    # 証明は、書き込み処理とは別に確かめる：計算結果を書き換えたファイル・指示していないセルを書いたファイルを見つける
    from todoku.cfb import CFB

    data = bytearray(CFB(res.blob).read("Book"))
    pos = data.find(struct.pack("<d", 40.0))
    data[pos:pos + 8] = struct.pack("<d", 41.0)
    bad = CFB(res.blob).replace("Book", bytes(data))
    pr2 = prove(blob, bad, res.allowed_inputs)
    assert any("計算し直すと" in u for u in pr2.unexpected), pr2.unexpected
    pr3 = prove(blob, res.blob, {("様式", "B1")})
    assert any("指示していない" in u for u in pr3.unexpected)


def test_biff5_write_refusals():
    from todoku.xls.book import XlsWorkbook
    from todoku.xlsx.edit import EditError, fill

    blob = make_biff5()
    with pytest.raises(EditError, match="cp932"):
        fill(XlsWorkbook(blob), {"様式!B1": "😀"})  # ブックの文字コードで書けない字
    with pytest.raises(EditError, match="DIMENSIONS"):
        fill(XlsWorkbook(blob), {"様式!H3": "x"})  # 使っている範囲の外
    with pytest.raises(EditError, match="ROW"):
        fill(XlsWorkbook(blob), {"様式!A7": "x"})  # 行のレコードがない行
    with pytest.raises(EditError, match="255"):
        fill(XlsWorkbook(blob), {"様式!B1": "あ" * 130})


@pytest.mark.corpus
def test_dual_format_reads_biff8_and_book_agrees():
    p = next(iter(CORPUS.glob("wide_legacy/**/*SENTEI.xls")), None) if CORPUS.exists() else None
    if p is None:
        pytest.skip("corpus がありません")
    from todoku.xls.book import XlsWorkbook

    blob = p.read_bytes()
    w8 = XlsWorkbook(blob)
    assert w8.biff == 8 and w8.dual_format
    w5 = XlsWorkbook(blob, stream="Book")
    assert w5.biff == 5
    for s8, s5 in zip(w8.sheets, w5.sheets):
        for k, cell in s8.cells.items():
            assert s5.cells[k].value == cell.value or (cell.value in (None, "") and s5.cells[k].value in (None, ""))


@pytest.mark.corpus
def test_dual_format_writes_both_streams_and_proves_both():
    p = next(iter(CORPUS.glob("wide_legacy/**/*SENTEI.xls")), None) if CORPUS.exists() else None
    if p is None:
        pytest.skip("corpus がありません")
    from todoku.cfb import CFB
    from todoku.xls.book import XlsWorkbook
    from todoku.xlsx import window
    from todoku.xlsx.edit import fill
    from todoku.xlsx.proof import prove

    blob = p.read_bytes()
    wb = XlsWorkbook(blob)
    fields = [f for s in window.build(wb)["sheets"] for f in s["fields"] if f["kind"] in ("text", "number")][:10]
    assert fields
    vals = {f["id"]: (3 if f["kind"] == "number" else "とどく") for f in fields}
    res = fill(wb, vals)
    assert sorted(res.changed_parts) == ["Book", "Workbook"]
    w8, w5 = XlsWorkbook(res.blob), XlsWorkbook(res.blob, stream="Book")
    for k, v in vals.items():
        sh, a = k.rsplit("!", 1)
        from todoku.formula.parser import parse_cell

        assert w8.sheet(sh).cells[parse_cell(a)].value == v and w5.sheet(sh).cells[parse_cell(a)].value == v
    pr = prove(blob, res.blob, res.allowed_inputs)
    assert pr.ok, pr.unexpected + pr.lint_errors
    # Book だけを古いままにしたファイルは、証明で見つかる
    stale = CFB(res.blob).replace("Book", CFB(blob).read("Book"), pad_ok=True)
    assert not prove(blob, stale, res.allowed_inputs).ok


# ====================================================================== Word 6.0/95
def _fkp(runs: list[tuple[int, int, bytes]], papx: bool) -> bytes:
    """512 バイトの FKP。runs は (FC の始まり, 終わり, grpprl)。段落は 7 バイトの BX、文字は 1 バイトの位置。"""
    page = bytearray(512)
    crun = len(runs)
    struct.pack_into(f"<{crun + 1}I", page, 0, *[r[0] for r in runs], runs[-1][1])
    top = 511
    for k, (_a, _b, g) in enumerate(runs):
        if not g:
            continue
        if papx:
            cw = (3 + len(g) + 1) // 2
            body = bytes([cw, 0, 0]) + g + bytes(2 * cw - 3 - len(g))
        else:
            body = bytes([len(g)]) + g
        top = (top - len(body)) & ~1
        page[top:top + len(body)] = body
        page[4 * (crun + 1) + (7 if papx else 1) * k] = top // 2
    page[511] = crun
    return bytes(page)


def make_word6(text: str, paras: dict[int, bytes] | None = None, chars: list[tuple[int, int, bytes]] | None = None,
               split: int | None = None, lid: int = 0x0411, nfib: int = 101, dop: bytes = bytes(16), enc: str = "cp932") -> bytes:
    """Word 6.0 の文書（仕様書の形のとおりに、ここで組み立てる）。text は文書の字（CP の順）。
    paras：段落の終わりの字の CP → その段落の grpprl。chars：(CP の始まり, 終わり, grpprl) の文字の書式。
    split を決めると、高速保存の形（ピースの表）にして、後ろの字を先に置く（字の並びとファイルの並びを変える）。"""
    raw = text.encode(enc)
    n = len(raw)
    fc_min = 0x200
    if split is None:
        body = raw
        cp_to_fc = list(range(fc_min, fc_min + n))
    else:
        body = raw[split:] + raw[:split]  # ファイルには後ろの字を先に置く
        cp_to_fc = [fc_min + (n - split) + i for i in range(split)] + [fc_min + i - split for i in range(split, n)]
    wd = bytearray(fc_min) + body
    wd += bytes((-len(wd)) % 512)
    lo, hi = fc_min, fc_min + n

    def runs_of(marks: dict[int, bytes]) -> list[tuple[int, int, bytes]]:
        cuts = sorted({lo, hi} | {f for fc in marks for f in (fc, fc + 1)})
        return [(a, b, marks.get(a, b"")) for a, b in zip(cuts, cuts[1:])]

    pmarks = {cp_to_fc[cp]: g for cp, g in (paras or {}).items()}
    cmarks: dict[int, bytes] = {}
    for a, b, g in chars or []:
        for cp in range(a, b):
            cmarks[cp_to_fc[cp]] = g
    pn_chp = len(wd) // 512
    wd += _fkp([r for r in runs_of(cmarks)], papx=False)
    pn_pap = len(wd) // 512
    wd += _fkp([r for r in runs_of(pmarks)], papx=True)
    tables: dict[str, bytes] = {"PlcfBteChpx": struct.pack("<IIH", lo, hi, pn_chp), "PlcfBtePapx": struct.pack("<IIH", lo, hi, pn_pap),
                                "Dop": dop}
    if split is not None:
        cps = [0, split, n]
        pcds = struct.pack("<HIH", 0, cp_to_fc[0], 0) + struct.pack("<HIH", 0, cp_to_fc[split], 0)
        plc = struct.pack("<3I", *cps) + pcds
        tables["Clx"] = b"\x02" + struct.pack("<H", len(plc)) + plc
    fib = bytearray(wd[:fc_min])
    struct.pack_into("<HHHH", fib, 0, 0xA5DC, nfib, 0x6000, lid)
    struct.pack_into("<H", fib, 0x0A, 0x0004 if split is not None else 0)
    for name, blob in tables.items():
        from todoku.doc.word6 import FIB6

        struct.pack_into("<II", fib, FIB6[name], len(wd), len(blob))
        wd += blob
    struct.pack_into("<III", fib, 0x18, fc_min, hi, len(wd))
    struct.pack_into("<i", fib, 0x34, n)
    struct.pack_into("<HHHH", fib, 0x18A, pn_chp, pn_pap, 1, 1)
    wd[:fc_min] = fib
    return build_cfb({"WordDocument": bytes(wd)})


IN_TABLE, ROW_END, HIDDEN = bytes([24, 1]), bytes([24, 1, 25, 1]), bytes([92, 1])


def _sample_word6(**kw) -> bytes:
    """申請書：本文・空欄・表（2 マス × 1 行）・□・テキストのフォームの欄・チェックボックスのフォームの欄・非表示の字。"""
    t = ("申請書\r氏名　　　　　　様\r住所\x07　　　　　　\x07\x07□同意する\r"
         "電話\x13 FORMTEXT \x14     \x15\r確認\x13 FORMCHECKBOX \x14\x15\r秘密\r")
    cp = lambda s: len(t[:t.index(s)].encode("cp932"))  # noqa: E731
    marks = [i for i, ch in enumerate(t) if ch in "\r\x07"]
    pos = [len(t[:i].encode("cp932")) for i in marks]
    paras = {pos[2]: IN_TABLE, pos[3]: IN_TABLE, pos[4]: ROW_END}
    hid = cp("秘密")
    return make_word6(t, paras, [(hid, hid + 4, HIDDEN)], **kw)


def test_word6_reads_text_tables_and_fields_in_codepage():
    from todoku.doc.book import DocFile
    from todoku.doc.word6 import Word6File, build, raw_chars, run
    from todoku.docx.window import to_markdown

    for kw in ({}, {"split": 13}):  # 高速保存（ピースの表で、ファイルの字の並びが違う）も同じに読む
        blob = _sample_word6(**kw)
        assert kind_of(blob) == "doc"
        doc = DocFile(blob, "t.doc")
        assert isinstance(doc, Word6File) and doc.enc == "cp932" and doc.complex == bool(kw)
        win = build(doc)
        md = to_markdown(win)
        assert "申請書" in md and "住所［P3.c2？］ ｜ ［T1.R1.C2］" in md and "同意する" in md and "〔非表示の文字: 秘密〕" in md, md
        assert "状態は読めません" in md  # チェックボックスの状態（FFDATA）は読まない
        assert sum(raw_chars(doc).values()) == len("申請書氏名様住所同意する電話確認秘密")  # 非表示の字も数える（窓に「非表示の文字」として出す）
        kinds = sorted(f["kind"] for f in win["fields"])
        assert "checkbox" in kinds and "text" in kinds, win["fields"]
        au = run(blob, "t.doc")
        assert au.ok, [(c.name, c.detail, c.items[:3]) for c in au.checks if not c.ok]


def test_word6_overwrites_in_place_with_proof():
    from todoku.cfb import CFB
    from todoku.doc.audit import prove as doc_prove
    from todoku.doc.book import DocFile
    from todoku.doc.fill import fill as doc_fill
    from todoku.doc.word6 import build
    from todoku.docx.document import iter_paras
    from todoku.docx.fill import FillError, written_of

    blob = _sample_word6()
    win = build(DocFile(blob))
    by_kind = {}
    for f in win["fields"]:
        by_kind.setdefault(f["kind"], []).append(f)
    cell = next(f for f in win["fields"] if f["id"].startswith("T"))
    inline = next(f for f in win["fields"] if f["id"].startswith("P2") and not f.get("readonly")
                  and f.get("maxlen_bytes", 0) >= 10)
    box = next(f for f in by_kind["checkbox"] if not f.get("readonly"))
    vals = {cell["id"]: "東京都", inline["id"]: "山田太郎", box["id"]: True}
    res = doc_fill(DocFile(blob), vals)
    pr = doc_prove(blob, res.blob, written_of(res.patches))
    assert pr.ok, pr.unexpected + pr.lint_errors + pr.readback
    # 変わったのは WordDocument の、書いた字のバイトだけ（大きさも同じ）
    a, b = CFB(blob).read("WordDocument"), CFB(res.blob).read("WordDocument")
    assert len(a) == len(b) and sum(x != y for x, y in zip(a, b)) <= len("東京都山田太郎■".encode("cp932"))
    texts = [p.text for _k, _p, bl in DocFile(res.blob).all_blocks() for p in iter_paras(bl)]
    assert "氏名　山田太郎　様" in texts and "東京都　　　" in texts and "■同意する" in texts, texts
    # 証明は、書き込み処理とは別に確かめる：書いていない字を変えたファイルを見つける
    wd = bytearray(b)
    k = wd.find("申請書".encode("cp932"))
    wd[k:k + 2] = "申".encode("cp932")[::-1]
    bad = CFB(res.blob).replace("WordDocument", bytes(wd))
    assert not doc_prove(blob, bad, written_of(res.patches)).ok
    # バイト数を超える値・文字コードで書けない字・字のない所は断る
    with pytest.raises(FillError, match="バイト"):
        doc_fill(DocFile(blob), {inline["id"]: "あ" * 20})
    with pytest.raises(FillError, match="cp932"):
        doc_fill(DocFile(blob), {inline["id"]: "😀"})


def test_word6_refuses_what_the_spec_does_not_establish():
    from todoku.doc.book import DocError, DocFile

    # 仕様書にない sprm の番号（長さが分からない）
    t = "あいう\r"
    with pytest.raises(DocError, match="sprm の番号 60"):
        DocFile(make_word6(t, {len(t.encode("cp932")) - 1: bytes([60, 1, 24, 1])}))
    # Macintosh の文字コード・決められない言語
    with pytest.raises(DocError, match="言語"):
        DocFile(make_word6(t, lid=0x0450))
    # Word 6.0 より古い版（Word 2.0）は、形を確かめていない
    wd = struct.pack("<HH", 0xA5DC, 45) + bytes(0x300)
    with pytest.raises(DocError, match="古い"):
        DocFile(build_cfb({"WordDocument": wd}))
    # 外字（利用者が作った字）は、どの字か決められない
    gaiji = make_word6("x\r")
    c = CFB(gaiji)
    w = bytearray(c.read("WordDocument"))
    w[0x200:0x202] = b"\xf0\x40"
    struct.pack_into("<I", w, 0x1C, 0x203)
    struct.pack_into("<i", w, 0x34, 3)
    with pytest.raises(DocError, match="外字"):
        DocFile(c.replace("WordDocument", bytes(w)))


def test_word6_protected_form_writes_only_form_fields():
    from todoku.doc.book import DocFile
    from todoku.doc.fill import fill as doc_fill
    from todoku.doc.word6 import build
    from todoku.docx.fill import FillError

    dop = bytearray(16)
    dop[7] = 0x02  # fProtEnabled（DOP の 6 バイト目からの語の 0x0200）
    blob = _sample_word6(dop=bytes(dop))
    doc = DocFile(blob)
    assert doc.protection == "forms" and doc.protection_enforced
    win = build(doc)
    text_ff = next(f for f in win["fields"] if f["kind"] == "text" and f.get("source") == "formfield")
    other = next(f for f in win["fields"] if f["id"].startswith("T"))
    res = doc_fill(DocFile(blob), {text_ff["id"]: "0312"})
    from todoku.doc.audit import prove as doc_prove
    from todoku.docx.fill import written_of

    assert doc_prove(blob, res.blob, written_of(res.patches)).ok
    with pytest.raises(FillError, match="保護"):
        doc_fill(DocFile(blob), {other["id"]: "x"})


def test_word6_through_the_api_and_grade(tmp_path):
    from todoku import api

    src = tmp_path / "w6.doc"
    src.write_bytes(_sample_word6())
    kind, win, _parts = api.window(str(src))
    assert kind == "doc" and win["word6"]["encoding"] == "cp932"
    cell = next(f for f in win["fields"] if f["id"].startswith("T"))
    out = tmp_path / "w6_filled.doc"
    res = api.fill(str(src), {cell["id"]: "大阪府"}, str(out))
    assert res["proof_ok"] and res["read_back_ok"], res
    g = api.grade(str(src))
    assert g["write"]["grade"] in ("B", "D") and any("状態の読めない" in r["title"] for r in g["reasons"]), g
