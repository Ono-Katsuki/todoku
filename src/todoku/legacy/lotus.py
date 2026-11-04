"""Lotus 1-2-3（.wk1・.wks）を、表のブックとして読む（読むだけ）。.wk3・.wk4・.123 は見分けて、次の一手を伝える。

レコード：型 2 バイト・長さ 2 バイト。BOF（0）の版 0x0404（.wks）・0x0405・0x0406（.wk1）。セルは 13 INTEGER・14 NUMBER・15 LABEL・16 FORMULA（書式 1・列 2・行 2 の後ろ）。
数式の字の結果は続く 51 STRING。数式のトークンは字に直さない（保存された結果を読み、数式は「未対応」と示す）。
字のバイトは LICS で、8 ビットの字は Shift_JIS と推定する（読み方を窓に書く）。PASSWORD（75）があれば暗号として読まない。
"""

from __future__ import annotations

import struct

from .tablebook import TableBook

WK1_VERSIONS = {0x0404, 0x0405, 0x0406}
LABEL, INTEGER, NUMBER, FORMULA, STRING, PASSWORD = 15, 13, 14, 16, 51, 75


class LotusError(ValueError):
    pass


def is_wk1(b: bytes) -> bool:
    return len(b) >= 6 and b[:4] == b"\x00\x00\x02\x00" and struct.unpack_from("<H", b, 4)[0] in WK1_VERSIONS


def is_wk3plus(b: bytes) -> bool:
    return len(b) >= 6 and b[:4] == b"\x00\x00\x1a\x00" and b[5] == 0x10


class Wk1Workbook(TableBook):
    def __init__(self, blob: bytes, filename: str = ""):
        super().__init__(blob, filename, "wk1")
        name = filename.rsplit("/", 1)[-1].rsplit(".", 1)[0] or "Sheet1"
        sh = self.add_sheet(name)
        guessed = False
        pos = 0
        last = None
        while pos + 4 <= len(blob):
            t, ln = struct.unpack_from("<HH", blob, pos)
            b = blob[pos + 4:pos + 4 + ln]
            if len(b) < ln:
                raise LotusError("レコードが途中で終わっています")
            pos += 4 + ln
            if t == 1:
                break
            if t == PASSWORD:
                raise LotusError("パスワードで保護された Lotus のワークシートです。解いて保存し直したものか、PDF を頼んでください")
            if t == STRING and last is not None and ln >= 5:
                s, g = self._text(b[5:].split(b"\0", 1)[0])
                guessed |= g
                last.value = s
                last.t = "str"
                continue
            if t in (INTEGER, NUMBER, LABEL, FORMULA) and ln >= 5:
                c, r = struct.unpack_from("<HH", b, 1)
                last = None
                if t == INTEGER:
                    self.put(sh, r + 1, c + 1, float(struct.unpack_from("<h", b, 5)[0]))
                elif t == NUMBER:
                    self.put(sh, r + 1, c + 1, struct.unpack_from("<d", b, 5)[0])
                elif t == LABEL:
                    s, g = self._text(b[6:].split(b"\0", 1)[0])  # 先頭の 1 字は、そろえ方の印
                    guessed |= g
                    self.put(sh, r + 1, c + 1, s)
                else:
                    cell = self.put(sh, r + 1, c + 1, struct.unpack_from("<d", b, 5)[0])
                    cell.formula = "?"
                    cell.parse_error = "未対応: Lotus の数式（保存された結果だけを読みました）"
                    last = cell
        if guessed:
            sh.unread.append("文字コードがファイルに書かれていないので、字を Shift_JIS と推定しました。字が正しいかはエージェントが確かめます")

    @staticmethod
    def _text(raw: bytes) -> tuple[str, bool]:
        if raw.isascii():
            return raw.decode("ascii"), False
        try:
            return raw.decode("cp932"), True
        except UnicodeDecodeError:
            return raw.decode("cp1252", "replace"), True


def raw_values(wb) -> dict:
    b = wb.zip.original
    out = set()
    pos = 0
    while pos + 4 <= len(b):
        t, ln = struct.unpack_from("<HH", b, pos)
        if t == 1:
            break
        if t in (INTEGER, NUMBER, LABEL, FORMULA) and ln >= 5:
            c, r = struct.unpack_from("<HH", b, pos + 5)
            out.add((r + 1, c + 1))
        pos += 4 + ln
    return {wb.sheets[0].name: out}
