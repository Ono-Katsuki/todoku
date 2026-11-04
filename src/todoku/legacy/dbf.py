"""dBase（.dbf。古い統計・地図の属性表で配られる）を、表のブックとして読む（読むだけ）。

- 見出し：[0] 版、[4:8] 件数、[8:10] 見出しの大きさ、[10:12] 1 件の大きさ、[29] 言語の印。32 バイト目から 32 バイトずつの項目（0x0D で終わる）
- 項目の型：C 字・N/F 数（字で書かれた数）・D 日付（YYYYMMDD）・L 真偽・I 整数・B 倍精度・Y 通貨・T 日時（FoxPro）。M・G は別のファイルのメモ（読まない）
- 削除の印（先頭が「*」）のある行は読まず、数を伝える
- 文字コード：言語の印（0x13・0x7B は Shift_JIS など）。印がなければ、全体が Shift_JIS として厳密に読めるかを見て、推定と書く
"""

from __future__ import annotations

import datetime as dt
import struct

from .tablebook import TableBook

LDID = {0x13: "cp932", 0x7B: "cp932", 0x01: "cp437", 0x02: "cp850", 0x03: "cp1252", 0x4D: "gbk", 0x4F: "big5", 0x79: "cp949", 0x7A: "gbk", 0x7D: "cp1255", 0x7E: "cp1256", 0x57: "cp1252", 0x58: "cp1252", 0x59: "cp1252"}
VERSIONS = {0x02, 0x03, 0x04, 0x05, 0x30, 0x31, 0x32, 0x43, 0x63, 0x83, 0x8B, 0x8E, 0xCB, 0xF5, 0xFB}


class DbfError(ValueError):
    pass


def parse_header(b: bytes):
    """(件数, 見出しの大きさ, 1 件の大きさ, 項目 [(名前, 型, 長さ, 小数)])。形が合わなければ DbfError。"""
    if len(b) < 33 or b[0] not in VERSIONS:
        raise DbfError("dBase の見出しではありません")
    n, hs, rs = struct.unpack_from("<IHH", b, 4)
    if hs < 33 or rs < 1 or hs > len(b):
        raise DbfError("dBase の見出しの大きさが合いません")
    fields = []
    p = 32
    off = 1
    while True:
        if p >= hs:
            raise DbfError("項目の終わりの印（0x0D）がありません")
        if b[p] == 0x0D:
            break
        if p + 32 > hs:
            raise DbfError("項目が見出しの外に出ています")
        raw = b[p:p + 11].split(b"\0", 1)[0]
        t = chr(b[p + 11])
        ln, dec = b[p + 16], b[p + 17]
        if t == "C" and b[0] in (0x30, 0x31, 0x43):  # FoxPro の長い字の項目は、長さの 2 バイト
            ln |= dec << 8
            dec = 0
        fields.append((raw, t, ln, dec))
        off += ln
        p += 32
    if off != rs:
        raise DbfError(f"項目の長さの合計（{off}）が、1 件の大きさ（{rs}）と合いません")
    if n * rs + hs > len(b) + 1:
        raise DbfError("件数のぶんのデータがありません（途中で切れています）")
    return n, hs, rs, fields


def is_dbf(b: bytes) -> bool:
    try:
        parse_header(b)
        return True
    except (DbfError, struct.error):
        return False


class DbfWorkbook(TableBook):
    def __init__(self, blob: bytes, filename: str = ""):
        super().__init__(blob, filename, "dbf")
        try:
            n, hs, rs, fields = parse_header(blob)
        except struct.error as e:
            raise DbfError(f"見出しが壊れています（{e}）") from None
        self.encoding, guessed = self._encoding(blob, n, hs, rs, fields)
        name = filename.rsplit("/", 1)[-1].rsplit(".", 1)[0] or "dbf"
        sh = self.add_sheet(name)
        for c, (raw, _t, _ln, _d) in enumerate(fields, 1):
            self.put(sh, 1, c, raw.decode(self.encoding, "replace"))
        memo = [raw.decode(self.encoding, "replace") for raw, t, _l, _d in fields if t in "MGP"]
        row = 1
        deleted = 0
        self.values = 0
        for i in range(n):
            p = hs + i * rs
            mark = blob[p]
            if mark == 0x2A:
                deleted += 1
                continue
            if mark == 0x1A:
                break
            row += 1
            q = p + 1
            for c, (_raw, t, ln, dc) in enumerate(fields, 1):
                v, st = self._value(blob[q:q + ln], t, dc)
                q += ln
                if v is not None:
                    self.put(sh, row, c, v, st)
                    self.values += 1
        if memo:
            sh.unread.append(f"メモの項目（{'・'.join(memo)}）の中身は別のファイル（.dbt・.fpt）にあり、読んでいません")
        if deleted:
            sh.unread.append(f"削除の印のある {deleted} 行は読んでいません")
        if guessed:
            sh.unread.append(f"文字コードの印がないので、{'Shift_JIS' if self.encoding == 'cp932' else 'Windows-1252'} と推定して読みました。字が正しいかはエージェントが確かめます")

    def _encoding(self, b: bytes, n: int, hs: int, rs: int, fields) -> tuple[str, bool]:
        ld = b[29]
        if ld in LDID:
            return LDID[ld], False
        for enc in ("ascii", "cp932"):
            try:
                b[hs:hs + n * rs].decode(enc)
                return ("cp932" if enc == "cp932" else "ascii"), enc == "cp932"
            except UnicodeDecodeError:
                continue
        return "cp1252", True

    def _value(self, raw: bytes, t: str, dc: int):
        if t == "C":
            s = raw.decode(self.encoding, "replace").rstrip(" \0")
            return (s or None), 0
        if t in "NF":
            s = raw.strip(b" \0*").decode("ascii", "replace")
            if not s or s == ".":
                return None, 0
            try:
                return float(s), 0
            except ValueError:
                return s, 0
        if t == "D":
            s = raw.strip(b" \0").decode("ascii", "replace")
            if len(s) == 8 and s.isdigit():
                try:
                    return f"{dt.date(int(s[:4]), int(s[4:6]), int(s[6:8])).isoformat()}", 0
                except ValueError:
                    return s, 0
            return (s or None), 0
        if t == "L":
            ch = raw[:1].upper()
            return (True if ch in (b"T", b"Y") else False if ch in (b"F", b"N") else None), 0
        if t == "I" and len(raw) == 4:
            return float(struct.unpack("<i", raw)[0]), 0
        if t in "BO" and len(raw) == 8:
            return struct.unpack("<d", raw)[0], 0
        if t == "Y" and len(raw) == 8:
            return struct.unpack("<q", raw)[0] / 10000, 0
        if t == "T" and len(raw) == 8:
            jd, ms = struct.unpack("<ii", raw)
            if jd == 0:
                return None, 0
            d = dt.date.fromordinal(jd - 1721425) if jd > 1721425 else None
            sec = ms // 1000
            return f"{d.isoformat()} {sec // 3600:02d}:{sec // 60 % 60:02d}:{sec % 60:02d}" if d else None, 0
        if t in "MGP":
            return None, 0
        return raw.decode(self.encoding, "replace").rstrip(" \0") or None, 0


def raw_values(wb) -> dict:
    """窓とは別に：データの部分を項目の長さで切り直し、空でないマスを数える。"""
    b = wb.zip.original
    n, hs, rs, fields = parse_header(b)
    out = set()
    row = 1
    for c in range(1, len(fields) + 1):
        out.add((1, c))
    for i in range(n):
        p = hs + i * rs
        if b[p] == 0x2A:
            continue
        if b[p] == 0x1A:
            break
        row += 1
        q = p + 1
        for c, (_r, t, ln, _d) in enumerate(fields, 1):
            cell = b[q:q + ln]
            q += ln
            if t in "MGP" or t in "ILBOYT":
                if t in "IBOY" or (t == "L" and cell[:1].upper() in (b"T", b"Y", b"F", b"N")) or (t == "T" and cell != bytes(ln)):
                    out.add((row, c))
                continue
            if cell.strip(b" \0") not in (b"", b".", b"*"):
                out.add((row, c))
    return {wb.sheets[0].name: out}
