"""Lotus 1-2-3（.wk1・.wks）を、表のブックとして読む（読むだけ）。.wk3・.wk4・.123 は見分けて、次の一手を伝える。

レコード：型 2 バイト・長さ 2 バイト。BOF（0）の版 0x0404（.wks）・0x0405・0x0406（.wk1）。セルは 13 INTEGER・14 NUMBER・15 LABEL・16 FORMULA（書式 1・列 2・行 2 の後ろ）。
数式の字の結果は続く 51 STRING。数式は逆ポーランドのコード（Lotus の公開資料 WSFF2・WSFF5 の命令の表）を字に直す。相対の印（番地の最上位ビット。残りは符号つきの差）は
資料に書かれていない、広く知られた形。Excel と同じ意味の命令と関数だけの式は計算でも確かめ、ほかは字に直すだけ（未対応と示す）。読めないコードは推測せず「?」のまま保存値を読む。
.wk3・.wk4・.123 の形は Lotus の公開資料がないので読まず、次の一手を伝える。
字のバイトは LICS で、8 ビットの字は Shift_JIS と推定する（読み方を窓に書く）。PASSWORD（75）があれば暗号として読まない。
"""

from __future__ import annotations

import struct

from .tablebook import TableBook

WK1_VERSIONS = {0x0404, 0x0405, 0x0406}
LABEL, INTEGER, NUMBER, FORMULA, STRING, PASSWORD = 15, 13, 14, 16, 51, 75


class LotusError(ValueError):
    pass


class _Bad(ValueError):
    pass


_BIN = {0x09: ("+", 3), 0x0A: ("-", 3), 0x0B: ("*", 4), 0x0C: ("/", 4), 0x0D: ("^", 5), 0x0E: ("=", 1), 0x0F: ("<>", 1), 0x10: ("<=", 1), 0x11: (">=", 1), 0x12: ("<", 1), 0x13: (">", 1), 0x18: ("&", 2)}
# 命令 → (名前, 引数の数。None は引数の数のバイトが続く)
_FUNC = {0x1F: ("NA", 0), 0x20: ("ERR", 0), 0x21: ("ABS", 1), 0x22: ("INT", 1), 0x23: ("SQRT", 1), 0x24: ("LOG", 1), 0x25: ("LN", 1), 0x26: ("PI", 0), 0x27: ("SIN", 1), 0x28: ("COS", 1),
         0x29: ("TAN", 1), 0x2A: ("ATAN2", 2), 0x2B: ("ATAN", 1), 0x2C: ("ASIN", 1), 0x2D: ("ACOS", 1), 0x2E: ("EXP", 1), 0x2F: ("MOD", 2), 0x30: ("CHOOSE", None), 0x31: ("ISNA", 1), 0x32: ("ISERR", 1),
         0x33: ("FALSE", 0), 0x34: ("TRUE", 0), 0x35: ("RAND", 0), 0x36: ("DATE", 3), 0x37: ("NOW", 0), 0x38: ("PMT", 3), 0x39: ("PV", 3), 0x3A: ("FV", 3), 0x3B: ("IF", 3), 0x3C: ("DAY", 1), 0x3D: ("MONTH", 1),
         0x3E: ("YEAR", 1), 0x3F: ("ROUND", 2), 0x40: ("TIME", 3), 0x41: ("HOUR", 1), 0x42: ("MINUTE", 1), 0x43: ("SECOND", 1), 0x44: ("ISNUMBER", 1), 0x45: ("ISSTRING", 1), 0x46: ("LENGTH", 1), 0x47: ("VALUE", 1),
         0x48: ("STRING", 2), 0x49: ("MID", 3), 0x4A: ("CHAR", 1), 0x4B: ("CODE", 1), 0x4C: ("FIND", 3), 0x4D: ("DATEVALUE", 1), 0x4E: ("TIMEVALUE", 1), 0x4F: ("CELLPOINTER", 1),
         0x50: ("SUM", None), 0x51: ("AVG", None), 0x52: ("COUNT", None), 0x53: ("MIN", None), 0x54: ("MAX", None), 0x55: ("VLOOKUP", 3), 0x56: ("NPV", 2), 0x57: ("VAR", None), 0x58: ("STD", None), 0x59: ("IRR", 2),
         0x5A: ("HLOOKUP", 3), 0x5B: ("DSUM", 3), 0x5C: ("DAVG", 3), 0x5D: ("DCOUNT", 3), 0x5E: ("DMIN", 3), 0x5F: ("DMAX", 3), 0x60: ("DVAR", 3), 0x61: ("DSTD", 3), 0x62: ("INDEX", 3), 0x63: ("COLS", 1), 0x64: ("ROWS", 1),
         0x65: ("REPEAT", 2), 0x66: ("UPPER", 1), 0x67: ("LOWER", 1), 0x68: ("LEFT", 2), 0x69: ("RIGHT", 2), 0x6A: ("REPLACE", 4), 0x6B: ("PROPER", 1), 0x6C: ("CELL", 2), 0x6D: ("TRIM", 1), 0x6E: ("CLEAN", 1),
         0x6F: ("S", 1), 0x70: ("N", 1), 0x71: ("EXACT", 2), 0x73: ("INDIRECT", 1)}
# Excel と同じ意味の関数（計算で確かめる）。ほかは字に直すだけ（@INT は切り捨ての向き、@COUNT は字も数えるなど、意味が違う）
_SAME = {"SUM": "SUM", "AVG": "AVERAGE", "MIN": "MIN", "MAX": "MAX", "ABS": "ABS", "SQRT": "SQRT", "IF": "IF", "ROUND": "ROUND"}


def _rc(v: int, base: int, is_col: bool) -> str:
    rel = v & 0x8000
    x = v & 0x7FFF
    n = base + (x - 0x4000 if x >= 0x2000 else x) if rel else x
    if n < 0 or n > (255 if is_col else 8191):
        raise _Bad("番地が範囲の外です")
    from ..formula.parser import num_to_col

    if is_col:
        return ("" if rel else "$") + num_to_col(n + 1)
    return ("" if rel else "$") + str(n + 1)


def formula_text(code: bytes, col: int, row: int, text) -> tuple[str, bool]:
    """(字に直した式, Excel と同じ意味か)。読めなければ _Bad。text：字のバイト → 字。"""
    st: list[tuple[str, int]] = []
    same = True
    i = 0
    n = len(code)

    def need(k: int) -> None:
        if i + k > n:
            raise _Bad("コードが途中で終わっています")

    while True:
        need(1)
        op = code[i]
        i += 1
        if op == 3:
            break
        if op == 0:
            need(8)
            v = struct.unpack_from("<d", code, i)[0]
            i += 8
            st.append((repr(v) if v != int(v) or abs(v) >= 1e15 else str(int(v)), 99))
        elif op == 1:
            need(4)
            c, r = struct.unpack_from("<HH", code, i)
            i += 4
            st.append((_rc(c, col, True) + _rc(r, row, False), 99))
        elif op == 2:
            need(8)
            c1, r1, c2, r2 = struct.unpack_from("<HHHH", code, i)
            i += 8
            st.append((f"{_rc(c1, col, True)}{_rc(r1, row, False)}:{_rc(c2, col, True)}{_rc(r2, row, False)}", 99))
        elif op == 4:
            if not st:
                raise _Bad("括弧の中がありません")
            st[-1] = (f"({st[-1][0]})", 99)
        elif op == 5:
            need(2)
            st.append((str(struct.unpack_from("<h", code, i)[0]), 99))
            i += 2
        elif op == 6:
            e = code.find(b"\0", i)
            if e < 0:
                raise _Bad("字が途中で終わっています")
            s = text(code[i:e])
            i = e + 1
            st.append(('"' + s.replace('"', '""') + '"', 99))
        elif op in (8, 0x17):
            if not st:
                raise _Bad("引数が足りません")
            a, pa = st.pop()
            st.append((("-" if op == 8 else "+") + (a if pa >= 6 else f"({a})"), 6))
        elif op in _BIN:
            if len(st) < 2:
                raise _Bad("引数が足りません")
            b, pb = st.pop()
            a, pa = st.pop()
            sym, pr = _BIN[op]
            st.append((f"{a if pa >= pr else f'({a})'}{sym}{b if pb > pr else f'({b})'}", pr))
        elif op in (0x14, 0x15):
            if len(st) < 2:
                raise _Bad("引数が足りません")
            b, _pb = st.pop()
            a, _pa = st.pop()
            st.append((f"{'AND' if op == 0x14 else 'OR'}({a},{b})", 99))
        elif op == 0x16:
            if not st:
                raise _Bad("引数が足りません")
            st.append((f"NOT({st.pop()[0]})", 99))
        elif op in _FUNC:
            name, argc = _FUNC[op]
            if argc is None:
                need(1)
                argc = code[i]
                i += 1
            if len(st) < argc:
                raise _Bad("引数が足りません")
            args = [st.pop()[0] for _ in range(argc)][::-1]
            excel = _SAME.get(name)
            if excel is None:
                same = False
            st.append((f"{excel or '@' + name}({','.join(args)})", 99))
        else:
            raise _Bad(f"命令 0x{op:02X} は読めません")
    if len(st) != 1:
        raise _Bad("式が 1 つにまとまりません")
    return st[0][0], same


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
                    self._formula(cell, b, c, r)
                    last = cell
        if guessed:
            sh.unread.append("文字コードがファイルに書かれていないので、字を Shift_JIS と推定しました。字が正しいかはエージェントが確かめます")

    def _formula(self, cell, b: bytes, c: int, r: int) -> None:
        from ..formula.parser import parse

        try:
            size = struct.unpack_from("<H", b, 13)[0]
            text, same = formula_text(b[15:15 + size], c, r, lambda raw: self._text(raw)[0])
        except (_Bad, struct.error, IndexError) as e:
            cell.formula = "?"
            cell.parse_error = f"未対応: Lotus の数式（保存された結果だけを読みました：{e}）"
            return
        cell.formula = text
        cell.formula_kind = "normal"
        if same:
            try:
                cell.ast = parse(text)
                return
            except Exception:  # noqa: BLE001
                pass
        cell.parse_error = "未対応: Lotus の数式を字に直しました（Lotus の関数は Excel と意味の違うものがあるので、計算では確かめず、保存された結果を読みます）"

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
