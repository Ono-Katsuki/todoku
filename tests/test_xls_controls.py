"""古いExcel（.xls、BIFF8）のフォームの部品（チェックボックス・オプションボタン・ドロップダウン・リストボックス・
スピンボタン・スクロールバー）を読み書きする。

公開されている .xls の様式（corpus/wide_xls）には部品がないので、ファイルを仕様書のとおりに組み立てて確かめる：
- 入れ物は複合ファイル（MS-CFB 版3：ヘッダー・FAT・ディレクトリ・ストリーム。4096バイト未満はミニストリーム）
- ブックのストリームは MS-XLS のレコード（BOF・Font・XF・BoundSheet8・SupBook・ExternSheet・MsoDrawingGroup・SST・
  ExtSST、シートは Index・計算と印刷の設定・DefColWidth・Dimensions・Row・セル・DBCell・Window2・EOF）
- 部品は、MsoDrawing（2.4.170：OfficeArt の図形。FSP・FOPT・OfficeArtClientAnchorSheet・OfficeArtClientData・
  OfficeArtClientTextbox）のすぐ後ろの Obj（2.4.181：FtCmo・FtCbls・FtRbo・FtSbs・ObjLinkFmla・FtCblsData・
  FtRboData・FtLbsData・FtGboData・ftEnd）と、見出しの TxO（2.4.329）と Continue
正解は、組み立てたときに置いた値と、状態のバイトの位置（組み立てる側で数えたもの）。
"""

from __future__ import annotations

import struct

import pytest

from todoku.cfb import CFB
from todoku.xls.book import XlsWorkbook, records
from todoku.xls.obj import parse as parse_obj
from todoku.xls.write import XlsEditor
from todoku.xlsx import window
from todoku.xlsx.audit import run as audit
from todoku.xlsx.calc import verify
from todoku.xlsx.edit import EditError, fill
from todoku.xlsx.proof import prove

MAXREC = 8224
FREE, END, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF
OT_CHECKBOX, OT_RADIO, OT_SPIN, OT_SCROLL, OT_LIST, OT_GBOX, OT_DROP = 0x0B, 0x0C, 0x10, 0x11, 0x12, 0x13, 0x14


# ====================================================================== 複合ファイル（MS-CFB）
def cfb(name: str, data: bytes) -> bytes:
    """ストリームが1つの複合ファイル（版3、セクタ512バイト）。4096バイト未満のストリームはミニストリーム（64バイト）に置く。"""
    size, mini, cutoff = 512, 64, 4096
    sectors: list[bytes] = []
    fat: list[int] = []

    def chain(blob: bytes) -> int:
        n = -(-len(blob) // size)
        first = len(sectors)
        for k in range(n):
            sectors.append(blob[k * size:(k + 1) * size].ljust(size, b"\0"))
            fat.append(first + k + 1 if k + 1 < n else END)
        return first if n else END

    root_start, root_size, minifat_start, n_minifat = END, 0, END, 0
    if len(data) < cutoff:
        n = -(-len(data) // mini)
        entries = [k + 1 if k + 1 < n else END for k in range(n)]
        entries += [FREE] * (-len(entries) % (size // 4))
        minifat_start = chain(struct.pack(f"<{len(entries)}I", *entries))
        n_minifat = len(entries) * 4 // size
        root_start = chain(data.ljust(n * mini, b"\0"))
        root_size = n * mini
        start = 0
    else:
        start = chain(data)

    def entry(nm: str, kind: int, child: int, first: int, length: int) -> bytes:
        raw = (nm + "\0").encode("utf-16-le")
        return (raw.ljust(64, b"\0") + struct.pack("<HBB", len(raw), kind, 1) + struct.pack("<III", NOSTREAM, NOSTREAM, child)
                + bytes(16) + struct.pack("<I", 0) + bytes(16) + struct.pack("<IQ", first, length))

    unused = bytes(64) + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) + bytes(36) + bytes(12)
    dir_start = chain(entry("Root Entry", 5, 1, root_start, root_size) + entry(name, 2, NOSTREAM, start, len(data)) + unused * 2)
    n_fat = 1
    while n_fat * (size // 4) < len(sectors) + n_fat:
        n_fat += 1
    fat_at = list(range(len(sectors), len(sectors) + n_fat))
    fat += [FATSECT] * n_fat
    fat += [FREE] * (n_fat * (size // 4) - len(fat))
    head = (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + bytes(16) + struct.pack("<HHHHH", 0x003E, 3, 0xFFFE, 9, 6) + bytes(6)
            + struct.pack("<IIIIIIIII", 0, n_fat, dir_start, 0, cutoff, minifat_start, n_minifat, END, 0)
            + struct.pack("<109I", *(fat_at + [FREE] * (109 - n_fat))))
    return head + b"".join(sectors) + struct.pack(f"<{len(fat)}I", *fat)


# ====================================================================== レコードの部品
def rec(t: int, body: bytes = b"") -> tuple[int, bytes]:
    return (t, body)


def xlstr(s: str, short: bool = False) -> bytes:
    """XLUnicodeString（short なら文字数が1バイトの ShortXLUnicodeString）。Latin-1 で書けないものは UTF-16。"""
    try:
        raw, high = s.encode("latin-1"), 0
    except UnicodeEncodeError:
        raw, high = s.encode("utf-16-le"), 1
    return (bytes([len(s)]) if short else struct.pack("<H", len(s))) + bytes([high]) + raw


def ptg_ref(r: int, c: int) -> bytes:
    """PtgRef（2.5.198.84、参照の形 0x24）：絶対参照（$C$1）。行・列は0から。"""
    return struct.pack("<BHH", 0x24, r, c)


def ptg_area(r1: int, c1: int, r2: int, c2: int) -> bytes:
    return struct.pack("<BHHHH", 0x25, r1, r2, c1, c2)


def ptg_area3d(ixti: int, r1: int, c1: int, r2: int, c2: int) -> bytes:
    return struct.pack("<BHHHHH", 0x3B, ixti, r1, r2, c1, c2)


def obj_fmla(rgce: bytes | None) -> bytes:
    """ObjFmla（2.5.187）：cbFmla・ObjectParsedFormula（cce・未使用4バイト・rgce）・偶数にする詰め物。"""
    if not rgce:
        return struct.pack("<H", 0)
    f = struct.pack("<H", len(rgce)) + bytes(4) + rgce
    f += b"\0" * (len(f) % 2)
    return struct.pack("<H", len(f)) + f


def ft(kind: int, body: bytes) -> bytes:
    return struct.pack("<HH", kind, len(body)) + body


class Obj:
    """OBJ の中身を部品の順に足していく。状態のバイトの位置（中身をつないだものの先頭から）も数える。"""

    def __init__(self, ot: int, oid: int, flags: int = 0x0011):
        self.ot, self.oid = ot, oid
        self.b = bytearray(ft(0x15, struct.pack("<HHH", ot, oid, flags) + bytes(12)))  # FtCmo
        self.at: dict[str, int] = {}

    def add(self, data: bytes, mark: str | None = None, off: int = 0) -> Obj:
        if mark:
            self.at[mark] = len(self.b) + off
        self.b += data
        return self

    def done(self) -> list[bytes]:
        return [bytes(self.b) + bytes(4)]  # reserved（ftEnd）


def checkbox(oid: int, checked: int, link: bytes | None) -> Obj:
    o = Obj(OT_CHECKBOX, oid).add(ft(0x0A, bytes(12)))  # FtCbls
    if link:
        o.add(struct.pack("<H", 0x14) + obj_fmla(link))  # ObjLinkFmla（ftCblsFmla）
    return o.add(ft(0x12, struct.pack("<HHHH", checked, 0, 0, 1)), "checked", 4)  # FtCblsData（fNo3d）


def radio(oid: int, checked: int, link: bytes | None, nxt: int, first: int) -> Obj:
    o = Obj(OT_RADIO, oid).add(ft(0x0A, bytes(12))).add(ft(0x0B, bytes(6)))  # FtCbls・FtRbo
    if link:
        o.add(struct.pack("<H", 0x14) + obj_fmla(link))
    o.add(ft(0x12, struct.pack("<HHHH", checked, 0, 0, 1)), "checked", 4)
    return o.add(ft(0x11, struct.pack("<HH", nxt, first)))  # FtRboData（idRadNext・fFirstBtn）


def sbs(val: int, lo: int, hi: int, inc: int = 1, page: int = 10, horiz: int = 0) -> bytes:
    """FtSbs（2.5.154）：unused1・iVal・iMin・iMax・dInc・dPage・fHoriz・dxScroll・フラグ（fDraw）。"""
    return ft(0x0C, bytes(4) + struct.pack("<hhhhhHhH", val, lo, hi, inc, page, horiz, 16, 0x0001))


def scroll(ot: int, oid: int, val: int, lo: int, hi: int, link: bytes | None, horiz: int = 0) -> Obj:
    o = Obj(ot, oid).add(sbs(val, lo, hi, horiz=horiz), "val", 8)
    if link:
        o.add(struct.pack("<H", 0x0E) + obj_fmla(link))  # ObjLinkFmla（ftSbsFmla）
    return o


def gbox(oid: int) -> Obj:
    return Obj(OT_GBOX, oid).add(ft(0x0F, struct.pack("<HHH", 0, 0, 1)))  # FtGboData


class ListObj(Obj):
    """リストボックス・ドロップダウン：FtSbs・ObjLinkFmla・FtLbsData（2.5.147）。OBJ はこの部品で終わる（ftEnd はない）。"""

    def __init__(self, ot: int, oid: int, fmla: bytes | None, lines: int, sel: int, link: bytes | None = None,
                 sel_type: int = 0, items: list[str] | None = None, bsels: list[int] | None = None, drop_text: str = "",
                 flags: int = 0x0011, lct: int | None = None):
        super().__init__(ot, oid, flags)
        self.add(sbs(0, 0, max(lines - 1, 0), page=8))
        if link:
            self.add(struct.pack("<H", 0x0E) + obj_fmla(link))
        self.lbs = (fmla, lines, sel, sel_type, items, bsels, drop_text, lct)

    def done(self) -> list[bytes]:
        fmla, lines, sel, sel_type, items, bsels, drop_text, lct = self.lbs
        bits = (0x02 if items is not None else 0) | 0x08 | (sel_type << 4)  # fValidPlex・fNo3d・wListSelType
        if lct is not None:
            bits |= 0x01 | (lct << 8)  # fUseCB と lct
        f = obj_fmla(fmla)
        fixed = f + struct.pack("<HHHH", lines, sel, bits, 0)  # cLines・iSel・フラグ・idEdit
        if self.ot == OT_DROP:  # LbsDropData：wStyle（0 は combo）・cLine・dxMin・str（奇数バイトなら1バイト足す）
            s = xlstr(drop_text)
            fixed += struct.pack("<HHH", 0, 8, 0) + s + b"\0" * (len(s) % 2)
        recs = [bytearray(self.b) + struct.pack("<HH", 0x13, 0)]
        start = len(recs[0])  # FtLbsData の欄の始まり（cbFContinued の後ろ）
        self.at["sel"] = start + len(f) + 2
        recs[0] += fixed
        total = len(recs[0])
        for s_ in items or []:  # rgLines：文字は1つのレコードの中に収める
            piece = xlstr(s_)
            if len(recs[-1]) + len(piece) > MAXREC:
                recs.append(bytearray())
            recs[-1] += piece
            total += len(piece)
        if sel_type:
            piece = bytes(bsels or [0] * lines)
            if len(recs[-1]) + len(piece) > MAXREC - 8:  # 入らないか、レコードの終わりの8バイトにかかるなら次へ
                recs.append(bytearray())
            self.at["bsels"] = total
            recs[-1] += piece
        # 1つのレコードに収まれば、Excel と同じく残りの最大の長さ（0x1FEE 以上になりうる長さ）、またがるなら「この中の長さ - 1」
        struct.pack_into("<H", recs[0], start - 2, 0x1FEE if len(recs) == 1 else len(recs[0]) - start - 1)
        return [bytes(r) for r in recs]


# ====================================================================== 図形（OfficeArt）
def oa(ver: int, inst: int, kind: int, body: bytes) -> bytes:
    """OfficeArtRecordHeader（recVer 4ビット・recInstance 12ビット・recType・recLen）と中身。"""
    return struct.pack("<HHI", ver | (inst << 4), kind, len(body)) + body


def shape_parts(spid: int, anchor: tuple[int, int, int, int], name: str | None, text: bool) -> tuple[bytes, bytes]:
    """図形の OfficeArtSpContainer を、OfficeArtClientData まで（この後ろに Obj）と、OfficeArtClientTextbox
    （この後ろに TxO）の2つに分けて返す。入れ物の長さは両方を合わせた長さ。"""
    fsp = oa(2, 0x00C9, 0xF00A, struct.pack("<II", spid, 0x0A00))  # msosptHostControl・fHaveAnchor・fHaveSpt
    props = [(0x007F, 0x01040104, b"")]  # 保護のフラグ
    if name:
        raw = (name + "\0").encode("utf-16-le")
        props.append((0x0380 | 0x8000, len(raw), raw))  # wzName（長さの変わる値は後ろに）
    fopt = oa(3, len(props), 0xF00B, b"".join(struct.pack("<HI", p, v) for p, v, _ in props) + b"".join(x for *_, x in props))
    c1, r1, c2, r2 = anchor
    anc = oa(0, 0, 0xF010, struct.pack("<HHHHHHHHH", 0, c1, 0, r1, 0, c2, 0, r2, 0))  # OfficeArtClientAnchorSheet
    data = oa(0, 0, 0xF011, b"")  # OfficeArtClientData
    tb = oa(0, 0, 0xF00D, b"") if text else b""  # OfficeArtClientTextbox
    body = fsp + fopt + anc + data
    return struct.pack("<HHI", 0x000F, 0xF004, len(body) + len(tb)) + body, tb


def txo(text: str) -> list[tuple[int, bytes]]:
    """TxO（2.4.329）：フラグ（左寄せ・上寄せ・fLockText）・rot・controlInfo か予約（どちらも6バイト）・cchText・cbRuns・
    ifntEmpty・fmla（なし）。続く Continue に XLUnicodeStringNoCch と TxORuns（Run 1つと TxOLastRun）。"""
    body = struct.pack("<HH", 0x0212, 0) + bytes(6) + struct.pack("<HHH", len(text), 16, 0) + struct.pack("<H", 0)
    return [rec(0x01B6, body), rec(0x003C, b"\x01" + text.encode("utf-16-le")),
            rec(0x003C, struct.pack("<HHHH", 0, 0, 0, 0) + struct.pack("<HHI", len(text), 0, 0))]


def drawing(shapes: list[dict]) -> list[tuple[int, bytes]]:
    """シートの図形の並び：MsoDrawing（OfficeArtDgContainer を図形ごとに分けたもの）・Obj・Continue・TxO。"""
    patriarch = oa(15, 0, 0xF004, oa(1, 0, 0xF009, bytes(16)) + oa(2, 0, 0xF00A, struct.pack("<II", 1024, 0x0005)))
    parts = [(shape_parts(1024 + s["obj"].oid, s["anchor"], s.get("name"), bool(s.get("text"))), s) for s in shapes]
    spgr_len = len(patriarch) + sum(len(a) + len(b) for (a, b), _ in parts)
    fdg = oa(0, 1, 0xF008, struct.pack("<II", len(shapes) + 1, 1024 + max(s["obj"].oid for s in shapes)))
    head = (struct.pack("<HHI", 0x000F, 0xF002, len(fdg) + 8 + spgr_len) + fdg
            + struct.pack("<HHI", 0x000F, 0xF003, spgr_len) + patriarch)
    out = []
    for k, ((first, tb), s) in enumerate(parts):
        out.append(rec(0x00EC, (head if k == 0 else b"") + first))
        objs = s["obj"].done()
        out.append(rec(0x005D, objs[0]))
        out += [rec(0x003C, x) for x in objs[1:]]
        if tb:
            out.append(rec(0x00EC, tb))
            out += txo(s["text"])
    return out


def drawing_group(n_shapes: int, last_spid: int) -> bytes:
    """MsoDrawingGroup の OfficeArtDggContainer（OfficeArtFDGGBlock：図形の番号の割り当て）。"""
    fdgg = oa(0, 0, 0xF006, struct.pack("<IIII", last_spid + 1, 2, n_shapes, 1) + struct.pack("<II", 1, n_shapes + 1))
    return struct.pack("<HHI", 0x000F, 0xF000, len(fdgg)) + fdgg


# ====================================================================== ブック
def _bof(dt: int) -> bytes:
    return struct.pack("<HHHHII", 0x0600, dt, 0x0DBB, 0x07CC, 0x000100C1, 0x00000406)


def workbook(sheets: list[dict], xti: list[tuple[int, int, int]] = ()) -> bytes:
    """sheets：[{"name", "cells": {(行, 列): 値}（0から。値は数・文字・真偽・("=", rgce, 保存値)）, "shapes": [...]}]。"""
    strings: list[str] = []
    refs = 0
    for sh in sheets:
        for v in sh["cells"].values():
            if isinstance(v, str):
                refs += 1
                if v not in strings:
                    strings.append(v)
    shapes = [s for sh in sheets for s in sh.get("shapes", [])]
    glob = [rec(0x0809, _bof(0x0005)), rec(0x00E1, struct.pack("<H", 1200)), rec(0x00C1, bytes(2)), rec(0x00E2),
            rec(0x005C, xlstr("todoku") + b" " * (112 - len(xlstr("todoku")))), rec(0x0042, struct.pack("<H", 1200)),
            rec(0x0161, bytes(2)), rec(0x013D, b"".join(struct.pack("<H", k + 1) for k in range(len(sheets)))),
            rec(0x0019, bytes(2)), rec(0x0012, bytes(2)), rec(0x0013, bytes(2)),
            rec(0x003D, struct.pack("<HHHHHHHHH", 0, 0, 0x4000, 0x2000, 0x0038, 0, 0, 1, 600)),
            rec(0x0040, bytes(2)), rec(0x008D, bytes(2)), rec(0x0022, bytes(2)), rec(0x000E, struct.pack("<H", 1)),
            rec(0x01B7, bytes(2)), rec(0x00DA, bytes(2))]
    glob += [rec(0x0031, struct.pack("<HHHHHBBBB", 220, 0, 0x7FFF, 400, 0, 0, 0, 128, 0) + xlstr("Arial", short=True))] * 4
    style_xf = struct.pack("<HHHBBBBIIH", 0, 0, 0xFFF5, 0x20, 0, 0, 0, 0, 0, 0x20C0)
    cell_xf = struct.pack("<HHHBBBBIIH", 0, 0, 0x0001, 0x20, 0, 0, 0, 0, 0, 0x20C0)
    glob += [rec(0x00E0, style_xf)] * 15 + [rec(0x00E0, cell_xf), rec(0x0293, struct.pack("<HBB", 0x8000, 0, 0xFF))]
    bound_at = len(glob)
    glob += [rec(0x0085, struct.pack("<IBB", 0, 0, 0) + xlstr(sh["name"], short=True)) for sh in sheets]
    glob.append(rec(0x008C, struct.pack("<HH", 81, 81)))
    if xti:
        glob += [rec(0x01AE, struct.pack("<HH", len(sheets), 0x0401)),
                 rec(0x0017, struct.pack("<H", len(xti)) + b"".join(struct.pack("<HHH", *x) for x in xti))]
    if shapes:
        glob.append(rec(0x00EB, drawing_group(len(shapes) + 1, 1024 + max(s["obj"].oid for s in shapes))))
    sst_at = len(glob)
    glob.append(rec(0x00FC, struct.pack("<II", refs, len(strings)) + b"".join(xlstr(s) for s in strings)))
    ext_at = len(glob)
    glob.append(rec(0x00FF, struct.pack("<H", 8) + bytes(8 * -(-len(strings) // 8))))
    glob.append(rec(0x000A))
    stream = list(glob)
    marks = []
    for sh in sheets:
        marks.append(len(stream))
        stream += _sheet(sh, strings)

    def offsets():
        out, pos = [], 0
        for _t, b in stream:
            out.append(pos)
            pos += 4 + len(b)
        return out

    offs = offsets()
    # 位置を持つレコード：BoundSheet8 の lbPlyPos・ExtSST・Index・DBCell（長さは変わらないので、一度で決まる）
    for k, m in enumerate(marks):
        t, b = stream[bound_at + k]
        stream[bound_at + k] = (t, struct.pack("<I", offs[m]) + b[4:])
    sst_body = stream[sst_at][1]
    pos, entries = 8, []
    for n, s in enumerate(strings):
        if n % 8 == 0:
            entries.append(struct.pack("<IHH", offs[sst_at] + 4 + pos, 4 + pos, 0))
        pos += len(xlstr(s))
    assert pos == len(sst_body)
    stream[ext_at] = (0x00FF, struct.pack("<H", 8) + b"".join(entries))
    for m in marks:
        _fix_index(stream, m, offs)
    return b"".join(struct.pack("<HH", t, len(b)) + b for t, b in stream)


def _cell_rec(r: int, c: int, v, strings: list[str]) -> tuple[int, bytes]:
    if isinstance(v, bool):
        return rec(0x0205, struct.pack("<HHHBB", r, c, 15, int(v), 0))
    if isinstance(v, str):
        return rec(0x00FD, struct.pack("<HHHI", r, c, 15, strings.index(v)))
    if isinstance(v, tuple):  # 数式：("=", rgce, 保存値)
        _, rgce, cached = v
        return rec(0x0006, struct.pack("<HHHdHIH", r, c, 15, cached, 0, 0, len(rgce)) + rgce)
    return rec(0x0203, struct.pack("<HHHd", r, c, 15, float(v)))


def _sheet(sh: dict, strings: list[str]) -> list[tuple[int, bytes]]:
    cells = sh["cells"]
    rows = sorted({r for r, _ in cells})
    out = [rec(0x0809, _bof(0x0010)), rec(0x020B, bytes(16 + 4 * len({r // 32 for r in rows})))]  # Index（あとで埋める）
    out += [rec(0x000D, struct.pack("<h", 1)), rec(0x000C, struct.pack("<H", 100)), rec(0x000F, struct.pack("<H", 1)),
            rec(0x0011, bytes(2)), rec(0x0010, struct.pack("<d", 0.001)), rec(0x005F, struct.pack("<H", 1)),
            rec(0x002A, bytes(2)), rec(0x002B, bytes(2)), rec(0x0082, struct.pack("<H", 1)), rec(0x0080, bytes(8)),
            rec(0x0225, struct.pack("<HH", 0, 255)), rec(0x0081, struct.pack("<H", 0x04C1)), rec(0x0014), rec(0x0015),
            rec(0x0083, bytes(2)), rec(0x0084, bytes(2)),
            rec(0x00A1, struct.pack("<HHHHHHHHddH", 9, 100, 1, 1, 1, 0x0044, 600, 600, 0.5, 0.5, 1)),
            rec(0x0055, struct.pack("<H", sh.get("defcolwidth", 8)))]
    out += sh.get("records", [])  # 足すレコード（STANDARDWIDTH・COLINFO など。DIMENSIONS の前）
    cols = [c for _, c in cells]
    out.append(rec(0x0200, struct.pack("<IIHHH", rows[0], rows[-1] + 1, min(cols), max(cols) + 1, 0)))
    for block in sorted({r // 32 for r in rows}):
        rs = [r for r in rows if r // 32 == block]
        for r in rs:
            cs = sorted(c for rr, c in cells if rr == r)
            out.append(rec(0x0208, struct.pack("<HHHHHHI", r, cs[0], cs[-1] + 1, 255, 0, 0, 0x0100)))
        for r in rs:
            for c in sorted(c for rr, c in cells if rr == r):
                out.append(_cell_rec(r, c, cells[(r, c)], strings))
        out.append(rec(0x00D7, bytes(4 + 2 * len(rs))))  # DBCell（あとで埋める）
    if sh.get("shapes"):
        out += drawing(sh["shapes"])
    out += [rec(0x023E, struct.pack("<HHHHHHHHH", 0x06B6, 0, 0, 0x40, 0, 0, 0, 0, 0)), rec(0x000A)]
    return out


def _fix_index(stream: list, start: int, offs: list[int]) -> None:
    """シートの Index（rwMic・rwMac・DefColWidth の位置・DBCell の位置）と、DBCell（最初の Row までの距離と、
    行ごとの最初のセルまでの距離：最初の行は最初の Row の終わりから、ほかは前の行の最初のセルから）を埋める。"""
    end = next(k for k in range(start + 1, len(stream)) if stream[k][0] == 0x000A)
    rows_seen, dbcells, block, first_cell = [], [], [], {}
    defcol = next(k for k in range(start, end) if stream[k][0] == 0x0055)
    for k in range(start, end):
        t, b = stream[k]
        if t == 0x0208:
            block.append(k)
            rows_seen.append(struct.unpack_from("<H", b, 0)[0])
        elif t in (0x0203, 0x00FD, 0x0205, 0x0006):
            first_cell.setdefault(struct.unpack_from("<H", b, 0)[0], k)
        elif t == 0x00D7:
            ref = offs[block[0]] + 4 + len(stream[block[0]][1])
            gaps = []
            for j in block:
                p = offs[first_cell[struct.unpack_from("<H", stream[j][1], 0)[0]]]
                gaps.append(p - ref)
                ref = p
            stream[k] = (t, struct.pack("<I", offs[k] - offs[block[0]]) + b"".join(struct.pack("<H", g) for g in gaps))
            dbcells.append(offs[k])
            block = []
    stream[start + 1] = (0x020B, struct.pack("<IIII", 0, min(rows_seen), max(rows_seen) + 1, offs[defcol])
                         + b"".join(struct.pack("<I", d) for d in dbcells))


# ====================================================================== 様式
SHEET = "申込"


def _rel(r: int, c: int) -> bytes:
    return struct.pack("<BHH", 0x44, r, c | 0xC000)  # PtgRef（値の形 0x44）の相対参照


def _int(n: int) -> bytes:
    return struct.pack("<BH", 0x1E, n)  # PtgInt


MUL, ADD = b"\x05", b"\x03"
F1 = _rel(0, 2) + _int(100) + MUL + _rel(0, 3) + _int(10) + MUL + ADD + _rel(0, 4) + ADD  # =C1*100+D1*10+E1
F2 = (_rel(0, 6) + _rel(0, 7) + _int(1000) + MUL + ADD + _rel(0, 9) + _int(100) + MUL + ADD
      + _rel(0, 8) + _int(10000) + MUL + ADD)  # =G1+H1*1000+J1*100+I1*10000


def make_book(radio_c1: float = 1, chain_first: bool = True) -> tuple[bytes, dict[int, Obj]]:
    """申込のシート（部品 12 個と、オートフィルターのボタン）と、選択肢のシート。(ファイル, id → 組み立てた OBJ)。"""
    objs = {
        1: gbox(1),
        2: radio(2, 1, ptg_ref(0, 2), 3, 1),
        3: radio(3, 0, ptg_ref(0, 2), 0, 0),
        4: ListObj(OT_DROP, 4, ptg_area(0, 0, 2, 0), 3, 1, link=ptg_ref(0, 3)),
        5: ListObj(OT_LIST, 5, ptg_area(0, 0, 2, 0), 3, 1, sel_type=1, bsels=[1, 0, 0]),
        6: scroll(OT_SPIN, 6, 3, 0, 10, ptg_ref(0, 4)),
        7: scroll(OT_SCROLL, 7, 50, 0, 100, ptg_ref(0, 6), horiz=1),
        8: ListObj(OT_LIST, 8, ptg_area3d(0, 0, 0, 3, 0), 4, 2, link=ptg_ref(0, 7)),
        9: radio(9, 0, ptg_ref(0, 8), 0, 0),  # 組の2番目（ファイルの中では先に置く）
        10: radio(10, 0, ptg_ref(0, 8), 9, 1 if chain_first else 0),  # 組の先頭 → 9
        11: ListObj(OT_DROP, 11, None, 3, 2, link=ptg_ref(0, 9), items=["赤", "青", "黄"]),
        12: checkbox(12, 1, ptg_ref(0, 10)),
        13: ListObj(OT_DROP, 13, None, 0, 0, flags=0x2101, lct=3),  # オートフィルターのボタン（fUIObj・lct=3）
    }
    names = {1: "Group Box 1", 2: "Option Button 2", 3: "Option Button 3", 4: "Drop Down 4", 5: "List Box 5",
             7: "Scroll Bar 7", 8: "List Box 8", 9: "Option Button 9", 10: "Option Button 10", 11: "Drop Down 11",
             12: "Check Box 12"}  # 6（スピン）は名前なし
    texts = {1: "性別", 2: "男", 3: "女", 9: "はい", 10: "いいえ", 12: "同意する"}
    anchors = {1: (5, 2, 9, 7), 2: (6, 3, 8, 4), 3: (6, 5, 8, 6), 4: (1, 9, 3, 10), 5: (1, 11, 3, 14), 6: (1, 16, 2, 18),
               7: (4, 16, 8, 17), 8: (4, 9, 6, 13), 9: (11, 3, 13, 4), 10: (11, 5, 13, 6), 11: (11, 9, 13, 10),
               12: (11, 12, 13, 13), 13: (0, 0, 1, 1)}
    cells = {(0, 0): "東京", (1, 0): "大阪", (2, 0): "名古屋", (0, 2): radio_c1, (0, 3): 1, (0, 4): 3,
             (0, 5): ("=", F1, radio_c1 * 100 + 10 + 3), (0, 6): 50, (0, 7): 2, (0, 9): 2, (0, 10): True,
             (1, 5): ("=", F2, 50 + 2000 + 200), (4, 1): "お申込みの内容"}
    shapes = [{"obj": objs[k], "anchor": anchors[k], "name": names.get(k), "text": texts.get(k)} for k in sorted(objs)]
    items = {(r, 0): s for r, s in enumerate(["北海道", "東北", "関東", "中部"])}
    data = workbook([{"name": SHEET, "cells": cells, "shapes": shapes}, {"name": "選択肢", "cells": items}],
                    xti=[(0, 1, 1)])
    return cfb("Workbook", data), objs


def make_long_list(n: int = 600) -> tuple[bytes, ListObj]:
    """項目がファイルの中にある（rgLines）複数選択のリストボックス。項目と bsels が Continue にまたがる。"""
    lst = ListObj(OT_LIST, 1, None, n, 2, sel_type=1, items=[f"選択肢{k:03d}" for k in range(1, n + 1)],
                  bsels=[1 if k in (2, n - 1) else 0 for k in range(1, n + 1)])
    data = workbook([{"name": "一覧", "cells": {(0, 0): "好きなもの"},
                      "shapes": [{"obj": lst, "anchor": (1, 1, 4, 12), "name": "List Box 1"}]}])
    return cfb("Workbook", data), lst


# ====================================================================== 確かめ
def _fields(blob: bytes) -> dict[str, dict]:
    return {f["id"]: f for sh in window.build(XlsWorkbook(blob))["sheets"] for f in sh["fields"]}


def _objs(blob: bytes) -> dict[int, bytes]:
    """ブックのストリームの OBJ（と続く Continue）の中身を、FtCmo の id ごとに。"""
    out, last = {}, None
    for _p, t, b in records(CFB(blob).read("Workbook")):
        if t == 0x005D:
            last = struct.unpack_from("<H", b, 6)[0]
            out[last] = b
        elif t == 0x003C and last is not None:
            out[last] += b
        else:
            last = None
    return out


def _allowed(res) -> tuple[set, dict]:
    return ({(p.sheet, p.address) for p in res.patches if p.kind == "input"},
            {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"})


def _tamper(blob: bytes, oid: int, off: int, new: bytes) -> bytes:
    """id が oid の OBJ（と続く Continue をつないだもの）の off バイト目から、new に書き換えたファイル（長さは同じ）。"""
    c = CFB(blob)
    data = bytearray(c.read("Workbook"))
    recs = list(records(bytes(data)))
    for k, (pos, t, b) in enumerate(recs):
        if t != 0x005D or struct.unpack_from("<H", b, 6)[0] != oid:
            continue
        for pos2, t2, b2 in recs[k:]:
            if pos2 != pos and t2 != 0x003C:
                break
            if off < len(b2):
                data[pos2 + 4 + off:pos2 + 4 + off + len(new)] = new
                return c.replace("Workbook", bytes(data))
            off -= len(b2)
    raise AssertionError(f"OBJ {oid} がありません")


def test_obj_records_follow_the_spec():
    # 読み方の確かめ：組み立てた OBJ を読むと、置いた値と状態の位置がそのまま出る
    blob, objs = make_book()
    got = _objs(blob)
    for oid, o in objs.items():
        info = parse_obj(got[oid])
        assert info.error is None and (info.ot, info.id) == (o.ot, oid)
        for key, off in o.at.items():
            assert info.at[key][0] == off, (oid, key)
    assert parse_obj(got[13]).ui_object and parse_obj(got[13]).lct == 3  # オートフィルターのボタン
    assert parse_obj(got[10]).rad_next == 9 and parse_obj(got[10]).first_btn
    assert parse_obj(got[11]).items == ["赤", "青", "黄"] and parse_obj(got[11]).drop_text == ""
    wb = XlsWorkbook(blob)
    assert XlsEditor(wb).build() == wb.data  # 書き込まずに組み立て直すと、ストリームはバイト一致


def test_controls_are_fields_with_options():
    blob, _ = make_book()
    sh = XlsWorkbook(blob).sheet(SHEET)
    assert [c.shape_id for c in sh.controls] == [str(k) for k in range(1, 13)]  # オートフィルターのボタンは部品ではない
    assert {c.name for c in sh.controls} >= {"Spin 6", "Check Box 12"}  # 名前のない図形は、種類と番号で呼ぶ
    f = _fields(blob)
    assert f["申込!C1"]["kind"] == "radio" and f["申込!C1"]["options"] == ["男", "女"] and f["申込!C1"]["value"] == "男"
    assert f["申込!C1"]["label"] == "性別"  # グループボックスの見出し
    assert f["申込!I1"]["kind"] == "radio" and f["申込!I1"]["options"] == ["いいえ", "はい"]  # 組の順は idRadNext のつながり
    assert f["申込!I1"]["value"] is None and "undecided" not in f["申込!I1"]
    assert f["申込!D1"]["kind"] == "select" and f["申込!D1"]["options"] == ["東京", "大阪", "名古屋"] and f["申込!D1"]["value"] == "東京"
    assert f["申込!List Box 5"]["kind"] == "multiselect" and f["申込!List Box 5"]["value"] == ["東京"]
    assert f["申込!E1"]["kind"] == "spin" and (f["申込!E1"]["min"], f["申込!E1"]["max"], f["申込!E1"]["value"]) == (0, 10, 3)
    assert f["申込!G1"]["kind"] == "spin" and (f["申込!G1"]["min"], f["申込!G1"]["max"], f["申込!G1"]["value"]) == (0, 100, 50)
    assert f["申込!H1"]["options"] == ["北海道", "東北", "関東", "中部"] and f["申込!H1"]["value"] == "東北"  # 別のシートの範囲
    assert f["申込!J1"]["options"] == ["赤", "青", "黄"] and f["申込!J1"]["value"] == "青"  # 項目がファイルの中にある
    assert f["申込!K1"]["kind"] == "checkbox" and f["申込!K1"]["value"] is True
    marks = "\n".join(x["text"] for x in window.build(XlsWorkbook(blob))["sheets"][0]["reading"])
    assert all(m in marks for m in ("◉男［C1］", "○女［C1］", "○はい［I1］", "▼東京", "≡東京", "⇕3", "⇕50", "≡東北", "▼青", "☑"))


def test_write_controls_updates_cells_states_and_formulas():
    blob, objs = make_book()
    values = {"申込!C1": "女", "申込!D1": "名古屋", "申込!List Box 5": ["大阪", "名古屋"], "申込!E1": 7, "申込!G1": 80,
              "申込!H1": "中部", "申込!I1": "はい", "申込!J1": "黄", "申込!K1": False}
    res = fill(XlsWorkbook(blob), values)
    inputs, ctl = _allowed(res)
    pr = prove(blob, res.blob, inputs, ctl)
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    assert pr.entries_identical == pr.entries_total - 1  # ブック以外のストリームはバイト一致
    assert len(pr.control_changes) == 10  # オプションボタン3つ（いいえ は選ばれないまま）・ほか7つ
    after = XlsWorkbook(res.blob)
    sh = after.sheet(SHEET)
    cell = {(r, c): x.value for (r, c), x in sh.cells.items()}
    assert (cell[(1, 3)], cell[(1, 4)], cell[(1, 5)], cell[(1, 7)], cell[(1, 8)], cell[(1, 9)], cell[(1, 10)]) == (2, 3, 7, 80, 4, 2, 3)
    assert cell[(1, 11)] is False
    assert cell[(1, 6)] == 2 * 100 + 3 * 10 + 7 and cell[(2, 6)] == 80 + 4 * 1000 + 3 * 100 + 2 * 10000  # つながったセルを使う数式
    assert verify(after).summary()["mismatch"] == 0
    st = {c.shape_id: c for c in sh.controls}
    assert (st["2"].checked, st["3"].checked, st["9"].checked, st["10"].checked, st["12"].checked) == (None, "Checked", "Checked", None, None)
    assert (st["4"].sel, st["5"].multi_sel, st["5"].sel, st["6"].val, st["7"].val, st["8"].sel, st["11"].sel) == (3, [2, 3], 2, 7, 80, 4, 3)
    f = _fields(res.blob)
    assert {k: f[k]["value"] for k in values} == {**values, "申込!E1": 7, "申込!K1": False}
    # OBJ は、状態のバイト（組み立てた側で数えた位置）だけが、書いた値になっている
    want = {2: {"checked": b"\0\0"}, 3: {"checked": b"\1\0"}, 9: {"checked": b"\1\0"}, 4: {"sel": b"\3\0"},
            5: {"sel": b"\2\0", "bsels": b"\0\1\1"}, 6: {"val": b"\7\0"}, 7: {"val": struct.pack("<h", 80)},
            8: {"sel": b"\4\0"}, 11: {"sel": b"\3\0"}, 12: {"checked": b"\0\0"}}
    before, now = _objs(blob), _objs(res.blob)
    for oid in objs:
        exp = bytearray(before[oid])
        for key, v in want.get(oid, {}).items():
            exp[objs[oid].at[key]:objs[oid].at[key] + len(v)] = v
        assert now[oid] == bytes(exp), oid


def test_proof_catches_changes_that_were_not_asked():
    blob, objs = make_book()
    res = fill(XlsWorkbook(blob), {"申込!D1": "名古屋", "申込!E1": 7})
    inputs, ctl = _allowed(res)
    assert prove(blob, res.blob, inputs, ctl).ok
    # 状態ではないバイト（スピンボタンの iMax）を変える
    bad = _tamper(res.blob, 6, objs[6].at["val"] + 4, struct.pack("<h", 99))
    pr = prove(blob, bad, inputs, ctl)
    assert not pr.ok and any("セル以外のレコードが変わりました" in u for u in pr.unexpected)
    # 指示していない部品の状態（チェックボックス）を変える：つながったセルとも食い違う
    bad = _tamper(res.blob, 12, objs[12].at["checked"], b"\0\0")
    pr = prove(blob, bad, inputs, ctl)
    assert any("指示していないチェックボックス「Check Box 12」" in u for u in pr.unexpected)
    # 指示した部品が、指示した状態と違う（ドロップダウンの iSel を 3 ではなく 2 に）
    bad = _tamper(res.blob, 4, objs[4].at["sel"], b"\2\0")
    pr = prove(blob, bad, inputs, ctl)
    assert any("書き込みを指示した 3 ではなく 2" in u for u in pr.unexpected)
    assert any("つながったセル $D$1 の値" in u for u in pr.unexpected)
    # 状態ではないバイト：リストボックスの項目の数（cLines。iSel のすぐ前）
    bad = _tamper(res.blob, 5, objs[5].at["sel"] - 2, b"\4\0")
    assert not prove(blob, bad, inputs, ctl).ok


def test_long_list_spans_continue_records():
    blob, lst = make_long_list()
    parts = [b for _p, t, b in records(CFB(blob).read("Workbook")) if t in (0x005D, 0x003C)]
    assert len(parts[0]) <= MAXREC and lst.at["bsels"] > len(parts[0])  # 複数選択の印は Continue の中
    f = _fields(blob)["一覧!List Box 1"]
    assert len(f["options"]) == 600 and f["value"] == ["選択肢002", "選択肢599"]
    res = fill(XlsWorkbook(blob), {"一覧!List Box 1": ["選択肢001", "選択肢600"]})
    inputs, ctl = _allowed(res)
    pr = prove(blob, res.blob, inputs, ctl)
    assert pr.ok, (pr.unexpected, pr.lint_errors)
    c = XlsWorkbook(res.blob).sheet("一覧").controls[0]
    assert c.multi_sel == [1, 600] and c.sel == 1
    assert _fields(res.blob)["一覧!List Box 1"]["value"] == ["選択肢001", "選択肢600"]
    # Continue の中の、状態ではないバイト（項目の文字）を変える
    bad = _tamper(res.blob, 1, lst.at["bsels"] - 4, "X".encode("utf-16-le"))
    assert not prove(blob, bad, inputs, ctl).ok


def test_write_refuses_values_outside_the_control():
    blob, _ = make_book()
    for bad in ({"申込!D1": "福岡"}, {"申込!E1": 11}, {"申込!G1": -1}, {"申込!List Box 5": ["福岡"]}, {"申込!J1": 4}):
        with pytest.raises(EditError):
            fill(XlsWorkbook(blob), bad)


def test_radio_groups_that_cannot_be_decided_are_not_written():
    # つながったセルの値（2）と、選ばれたボタン（1番目）が食い違う
    blob, _ = make_book(radio_c1=2)
    assert "undecided" in _fields(blob)["申込!C1"]
    with pytest.raises(EditError):
        fill(XlsWorkbook(blob), {"申込!C1": "女"})
    # 組の先頭の印（fFirstBtn）がない
    blob, _ = make_book(chain_first=False)
    f = _fields(blob)["申込!I1"]
    assert "fFirstBtn" in f["undecided"]
    with pytest.raises(EditError):
        fill(XlsWorkbook(blob), {"申込!I1": "はい"})


def test_audit_writes_and_reads_back_every_control():
    for blob in (make_book()[0], make_long_list()[0]):
        au = audit(blob, "controls.xls")
        bad = [(c.name, c.detail, c.items[:5]) for c in au.checks if not c.ok]
        assert not bad, bad
