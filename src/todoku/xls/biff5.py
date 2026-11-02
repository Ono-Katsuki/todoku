"""Excel 5.0 / Excel 95（BIFF5・BIFF7、ストリーム名「Book」）のブックを読む。外部のライブラリは使わない。

レコードの形は OpenOffice.org の「Microsoft Excel File Format」（BIFF2〜BIFF8 の違いを並べた資料）で確かめた：
- 字はバイト列で、CODEPAGE（0x0042）の文字コード（日本語は 932）。LABEL・RSTRING・STRING は 16 ビットの長さ、
  SHEET（BOUNDSHEET）・FONT・FORMAT・HEADER・FOOTER は 8 ビットの長さ。共有文字列（SST）はない
- XF は 16 バイト（4〜5 バイト目に保護、8〜15 バイト目に塗りと罫線）。NOTE は字をレコードの中に持つ
- 数値・空・真偽・行・列幅・保護のレコードは BIFF8 と同じ形
数式のトークンは BIFF8 と長さが違う（tRef 4・tArea 7・tRef3d 18・tName 15 バイトなど）。読み違えた数式で計算すると
保存値と合わない値を作るので、数式は Excel が保存した計算結果だけを読み、数式そのものは「未対応」として示す
（計算で確かめられないので、段階は C）。書き込みは、まだしない（xlsx/edit.py で断る）。
"""

from __future__ import annotations

import codecs
import struct

from ..formula.values import ExcelError
from ..xlsx.workbook import Cell, DefinedName, Shape, Sheet, Style, _header_text

FORMULA_NOTE = "未対応: Excel 5.0/95（BIFF5）の数式。Excel が保存した計算結果だけを読みました"


def _codec(cp: int) -> str:
    name = {1200: "utf-16-le", 367: "ascii", 10000: "mac_roman", 32768: "mac_roman"}.get(cp, f"cp{cp}")
    try:
        codecs.lookup(name)
        return name
    except LookupError:
        return "cp1252"


class Biff5Reader:
    """XlsWorkbook（BIFF5 の版）の中身を埋める。wb は読み終わると xlsx と同じ形になる。"""

    def __init__(self, wb, it: list):
        self.wb = wb
        self.it = it
        self.enc = "cp1252"

    def s8(self, b: bytes, pos: int) -> tuple[str, int]:
        n = b[pos]
        return b[pos + 1:pos + 1 + n].decode(self.enc, errors="replace"), pos + 1 + n

    def s16(self, b: bytes, pos: int) -> tuple[str, int]:
        n = struct.unpack_from("<H", b, pos)[0]
        return b[pos + 2:pos + 2 + n].decode(self.enc, errors="replace"), pos + 2 + n

    def load(self) -> None:
        wb = self.wb
        bound = []
        for _pos, t, b in self.it:
            if t == 0x000A:
                break
            if t == 0x0042 and len(b) >= 2:
                self.enc = _codec(struct.unpack_from("<H", b, 0)[0])
            elif t == 0x0022:
                wb.date1904 = bool(struct.unpack_from("<H", b, 0)[0])
            elif t == 0x0031:  # FONT：色の番号は 4 バイト目（BIFF8 と同じ）
                wb._fonts.append(struct.unpack_from("<H", b, 4)[0])
            elif t == 0x041E:  # FORMAT：番号・8 ビットの長さの字
                wb._formats[struct.unpack_from("<H", b, 0)[0]] = self.s8(b, 2)[0]
            elif t == 0x0092:
                n = struct.unpack_from("<H", b, 0)[0]
                for k in range(min(n, len(wb._palette))):
                    wb._palette[k] = "".join(f"{x:02X}" for x in b[2 + 4 * k:5 + 4 * k])
            elif t == 0x00E0:
                self._xf(b)
            elif t == 0x0085:  # SHEET：位置・状態・種類・8 ビットの長さの名前
                lb = struct.unpack_from("<I", b, 0)[0]
                bound.append((lb, b[4] & 0x03, b[5], self.s8(b, 6)[0]))
            elif t == 0x0018:
                self._name(b)
            elif t == 0x0012:
                wb.workbook_protected = bool(struct.unpack_from("<H", b, 0)[0])
        for lb, state, dt, name in bound:
            st = {0: "visible", 1: "hidden", 2: "veryHidden"}.get(state, "visible")
            kind = {2: "chartsheet", 6: "vbmodule"}.get(dt, f"type{dt}")
            wb.sheets.append(Sheet(name=name, path=f"#{lb}", state=st if dt == 0 else f"non-worksheet:{kind}"))
        for (lb, _s, dt, _n), sh in zip(bound, wb.sheets):
            if dt == 0:
                self._sheet(sh, lb)

    def _xf(self, b: bytes) -> None:
        wb = self.wb
        ifnt, ifmt, prot = struct.unpack_from("<HHH", b, 0)
        area, lines = struct.unpack_from("<II", b, 8)
        fls = (area >> 16) & 0x3F
        fnt = ifnt - 1 if ifnt >= 4 else ifnt  # フォントの4番は欠番
        fc = wb._color(wb._fonts[fnt]) if 0 <= fnt < len(wb._fonts) else None
        fill = wb._color(area & 0x7F) if fls == 1 else None
        bottom, top, left, right = (area >> 22) & 7, lines & 7, (lines >> 3) & 7, (lines >> 6) & 7
        bits = (1 if left else 0) | (2 if right else 0) | (4 if top else 0) | (8 if bottom else 0)
        wb.styles.append(Style(wb._formats.get(ifmt, "General"), bool(prot & 1), bool(prot & 2), fc, fill, bits))

    def _name(self, b: bytes) -> None:
        grbit, _key, cch, _cce, _ixals, itab = struct.unpack_from("<HBBHHH", b, 0)
        raw = b[14:14 + cch].decode(self.enc, errors="replace")
        from .book import _BUILTIN_NAMES

        nm = "_xlnm." + _BUILTIN_NAMES.get(ord(raw[0]), f"builtin{ord(raw[0])}") if grbit & 0x20 and raw else raw
        names = [s.name for s in self.wb.sheets]
        scope = names[itab - 1] if 0 < itab <= len(names) else None
        self.wb.names.append(DefinedName(nm, "", scope, bool(grbit & 1), None, FORMULA_NOTE))

    def _sheet(self, sh: Sheet, lb: int) -> None:
        from .book import ERRORS, _rk, records

        wb = self.wb
        depth = 0
        last = None
        note_for = None
        rows = []
        for _pos, t, b in records(wb.data, lb):
            if t == 0x0809:
                depth += 1
                continue
            if t == 0x000A:
                depth -= 1
                if depth == 0:
                    break
                continue
            if depth == 1:
                rows.append((t, b))
        for i, (t, b) in enumerate(rows):
            if t == 0x0203:
                r, c, x = struct.unpack_from("<HHH", b, 0)
                wb._put(sh, r, c, x, struct.unpack_from("<d", b, 6)[0])
            elif t == 0x027E:
                r, c, x, v = struct.unpack_from("<HHHI", b, 0)
                wb._put(sh, r, c, x, _rk(v))
            elif t == 0x00BD:
                r, c0 = struct.unpack_from("<HH", b, 0)
                for k in range((len(b) - 6) // 6):
                    x, v = struct.unpack_from("<HI", b, 4 + 6 * k)
                    wb._put(sh, r, c0 + k, x, _rk(v))
            elif t in (0x0204, 0x00D6):  # LABEL・RSTRING：16 ビットの長さのバイト列
                r, c, x = struct.unpack_from("<HHH", b, 0)
                wb._put(sh, r, c, x, self.s16(b, 6)[0])
            elif t == 0x0201:
                r, c, x = struct.unpack_from("<HHH", b, 0)
                wb._put(sh, r, c, x, None)
            elif t == 0x00BE:
                r, c0 = struct.unpack_from("<HH", b, 0)
                for k in range((len(b) - 6) // 2):
                    wb._put(sh, r, c0 + k, struct.unpack_from("<H", b, 4 + 2 * k)[0], None)
            elif t == 0x0205:
                r, c, x, v, is_err = struct.unpack_from("<HHHBB", b, 0)
                wb._put(sh, r, c, x, ExcelError(ERRORS.get(v, "#N/A")) if is_err else bool(v))
            elif t == 0x0006:  # FORMULA：形は BIFF8 と同じ。保存された結果だけを読む
                r, c, x = struct.unpack_from("<HHH", b, 0)
                res = b[6:14]
                cell = wb._put(sh, r, c, x, None)
                if res[6:8] == b"\xff\xff":
                    cell.value = {1: bool(res[2]), 2: ExcelError(ERRORS.get(res[2], "#N/A")), 3: ""}.get(res[0], "")
                    if res[0] == 0:
                        cell.t = "str"
                else:
                    cell.value = struct.unpack_from("<d", res, 0)[0]
                cell.formula = "?"
                cell.formula_kind = "normal"
                cell.parse_error = FORMULA_NOTE
                last = cell
            elif t == 0x0207 and last is not None:
                last.value = self.s16(b, 0)[0]
            elif t == 0x0208:
                rw, _a, _b, miy = struct.unpack_from("<HHHH", b, 0)
                flags = struct.unpack_from("<I", b, 12)[0] if len(b) >= 16 else 0
                if flags & 0x20:
                    sh.hidden_rows.add(rw + 1)
                sh.row_heights[rw + 1] = (miy & 0x7FFF) / 20
            elif t == 0x007D:
                c1, c2, w, x, flags = struct.unpack_from("<HHHHH", b, 0)
                for c in range(c1, min(c2, 255) + 1):
                    sh.col_widths[c + 1] = w / 256
                    if flags & 0x01:
                        sh.hidden_cols.add(c + 1)
            elif t == 0x0012:
                sh.protected = bool(struct.unpack_from("<H", b, 0)[0])
            elif t in (0x0014, 0x0015) and b:
                txt = _header_text(self.s8(b, 0)[0])
                if txt:
                    sh.header_footer.append(txt)
            elif t == 0x001C:  # NOTE：行・列・字数・字（続きは行が 0xFFFF の NOTE）
                r, c, n = struct.unpack_from("<HHH", b, 0)
                txt = b[6:6 + n].decode(self.enc, errors="replace")
                if r == 0xFFFF and note_for is not None:
                    sh.comments[note_for] = sh.comments.get(note_for, "") + txt
                else:
                    note_for = (r + 1, c + 1)
                    sh.comments[note_for] = txt
            elif t == 0x01B6:  # TXO：図形の字は、後ろの CONTINUE にバイト列で入る
                n = struct.unpack_from("<H", b, 10)[0] if len(b) >= 12 else 0
                if n and i + 1 < len(rows) and rows[i + 1][0] == 0x003C:
                    txt = rows[i + 1][1][:n].decode(self.enc, errors="replace")
                    if txt.strip():
                        sh.shapes.append(Shape("text", "", txt, "", None))


def raw_texts5(wb) -> dict[str, list[tuple[str, str]]]:
    """窓とは別の読み方（BIFF5）：レコードを直接読み、セルの字と数値・メモ・図形の字・ヘッダーを拾う。"""
    from ..formula.parser import addr
    from .book import _rk, records

    r5 = Biff5Reader(wb, [])
    r5.enc = wb.biff5_encoding
    out: dict[str, list[tuple[str, str]]] = {"セルの文字": [], "セルの数値": [], "図形の文字": [], "メモ": [],
                                             "チェックボックスなどの文字": [], "入力規則の説明": [], "ヘッダー・フッター": []}
    for sh in wb.sheets:
        if sh.state.startswith("non-"):
            continue
        depth = 0
        prev = None
        for _pos, t, b in records(wb.data, int(sh.path[1:])):
            if t == 0x0809:
                depth += 1
                continue
            if t == 0x000A:
                depth -= 1
                if depth == 0:
                    break
                continue
            if depth != 1:
                continue
            if t in (0x0204, 0x00D6):
                r, c = struct.unpack_from("<HH", b, 0)
                s = r5.s16(b, 6)[0]
                if s:
                    out["セルの文字"].append((f"{sh.name}!{addr(r + 1, c + 1)}", s))
            elif t == 0x0203:
                r, c = struct.unpack_from("<HH", b, 0)
                out["セルの数値"].append((f"{sh.name}!{addr(r + 1, c + 1)}", repr(struct.unpack_from("<d", b, 6)[0])))
            elif t == 0x027E:
                r, c, _x, v = struct.unpack_from("<HHHI", b, 0)
                out["セルの数値"].append((f"{sh.name}!{addr(r + 1, c + 1)}", repr(_rk(v))))
            elif t == 0x00BD:
                r, c0 = struct.unpack_from("<HH", b, 0)
                for k in range((len(b) - 6) // 6):
                    out["セルの数値"].append((f"{sh.name}!{addr(r + 1, c0 + k + 1)}",
                                         repr(_rk(struct.unpack_from("<I", b, 6 + 6 * k)[0]))))
            elif t == 0x0006 and b[12:14] != b"\xff\xff":
                r, c = struct.unpack_from("<HH", b, 0)
                out["セルの数値"].append((f"{sh.name}!{addr(r + 1, c + 1)}", repr(struct.unpack_from("<d", b, 6)[0])))
            elif t == 0x0207:
                s = r5.s16(b, 0)[0]
                if s:
                    out["セルの文字"].append((f"{sh.name}（数式の結果）", s))
            elif t == 0x001C:
                n = struct.unpack_from("<H", b, 4)[0]
                s = b[6:6 + n].decode(r5.enc, errors="replace")
                if s.strip():
                    out["メモ"].append((f"{sh.name}（メモ）", s))
            elif t == 0x003C and prev == 0x01B6:
                pass
            elif t in (0x0014, 0x0015) and b:
                s = r5.s8(b, 0)[0]
                if s.strip():
                    out["ヘッダー・フッター"].append((f"{sh.name}（ヘッダー）", s))
            prev = t
    return out
