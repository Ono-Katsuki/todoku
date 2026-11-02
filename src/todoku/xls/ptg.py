"""古いExcel（BIFF8）の数式（逆ポーランド記法のトークン列）を、数式の文字列に戻す。

文字列にしたあとは、xlsx と同じ構文解析・計算エンジンで計算する。トークンを読み違えると
計算結果が保存値と合わなくなるので、全数照合で必ず分かる。
"""

from __future__ import annotations

import struct

from ..formula.parser import num_to_col
from ..formula.values import Unsupported

# 関数の番号 → (名前, 引数の数。None は可変)
FTAB: dict[int, tuple[str, int | None]] = {
    0: ("COUNT", None), 1: ("IF", None), 2: ("ISNA", 1), 3: ("ISERROR", 1), 4: ("SUM", None), 5: ("AVERAGE", None),
    6: ("MIN", None), 7: ("MAX", None), 8: ("ROW", None), 9: ("COLUMN", None), 10: ("NA", 0), 11: ("NPV", None),
    12: ("STDEV", None), 13: ("DOLLAR", None), 14: ("FIXED", None), 15: ("SIN", 1), 16: ("COS", 1), 17: ("TAN", 1),
    18: ("ATAN", 1), 19: ("PI", 0), 20: ("SQRT", 1), 21: ("EXP", 1), 22: ("LN", 1), 23: ("LOG10", 1), 24: ("ABS", 1),
    25: ("INT", 1), 26: ("SIGN", 1), 27: ("ROUND", 2), 28: ("LOOKUP", None), 29: ("INDEX", None), 30: ("REPT", 2),
    31: ("MID", 3), 32: ("LEN", 1), 33: ("VALUE", 1), 34: ("TRUE", 0), 35: ("FALSE", 0), 36: ("AND", None),
    37: ("OR", None), 38: ("NOT", 1), 39: ("MOD", 2), 40: ("DCOUNT", 3), 41: ("DSUM", 3), 42: ("DAVERAGE", 3),
    43: ("DMIN", 3), 44: ("DMAX", 3), 45: ("DSTDEV", 3), 46: ("VAR", None), 47: ("DVAR", 3), 48: ("TEXT", 2),
    49: ("LINEST", None), 50: ("TREND", None), 51: ("LOGEST", None), 52: ("GROWTH", None), 56: ("PV", None),
    57: ("FV", None), 58: ("NPER", None), 59: ("PMT", None), 60: ("RATE", None), 61: ("MIRR", 3), 62: ("IRR", None),
    63: ("RAND", 0), 64: ("MATCH", None), 65: ("DATE", 3), 66: ("TIME", 3), 67: ("DAY", 1), 68: ("MONTH", 1),
    69: ("YEAR", 1), 70: ("WEEKDAY", None), 71: ("HOUR", 1), 72: ("MINUTE", 1), 73: ("SECOND", 1), 74: ("NOW", 0),
    75: ("AREAS", 1), 76: ("ROWS", 1), 77: ("COLUMNS", 1), 78: ("OFFSET", None), 82: ("SEARCH", None),
    83: ("TRANSPOSE", 1), 86: ("TYPE", 1), 97: ("ATAN2", 2), 98: ("ASIN", 1), 99: ("ACOS", 1), 100: ("CHOOSE", None),
    101: ("HLOOKUP", None), 102: ("VLOOKUP", None), 105: ("ISREF", 1), 109: ("LOG", None), 111: ("CHAR", 1),
    112: ("LOWER", 1), 113: ("UPPER", 1), 114: ("PROPER", 1), 115: ("LEFT", None), 116: ("RIGHT", None),
    117: ("EXACT", 2), 118: ("TRIM", 1), 119: ("REPLACE", 4), 120: ("SUBSTITUTE", None), 121: ("CODE", 1),
    124: ("FIND", None), 125: ("CELL", None), 126: ("ISERR", 1), 127: ("ISTEXT", 1), 128: ("ISNUMBER", 1),
    129: ("ISBLANK", 1), 130: ("T", 1), 131: ("N", 1), 140: ("DATEVALUE", 1), 141: ("TIMEVALUE", 1), 142: ("SLN", 3),
    143: ("SYD", 4), 144: ("DDB", None), 148: ("INDIRECT", None), 162: ("CLEAN", 1), 163: ("MDETERM", 1),
    164: ("MINVERSE", 1), 165: ("MMULT", 2), 167: ("IPMT", None), 168: ("PPMT", None), 169: ("COUNTA", None),
    183: ("PRODUCT", None), 184: ("FACT", 1), 189: ("DPRODUCT", 3), 190: ("ISNONTEXT", 1), 193: ("STDEVP", None),
    194: ("VARP", None), 195: ("DSTDEVP", 3), 196: ("DVARP", 3), 197: ("TRUNC", None), 198: ("ISLOGICAL", 1),
    199: ("DCOUNTA", 3), 204: ("YEN", None), 205: ("FINDB", None), 206: ("SEARCHB", None), 207: ("REPLACEB", 4),
    208: ("LEFTB", None), 209: ("RIGHTB", None), 210: ("MIDB", 3), 211: ("LENB", 1), 212: ("ROUNDUP", 2),
    213: ("ROUNDDOWN", 2), 214: ("ASC", 1), 215: ("JIS", 1), 216: ("RANK", None), 219: ("ADDRESS", None),
    220: ("DAYS360", None), 221: ("TODAY", 0), 222: ("VDB", None), 227: ("MEDIAN", None), 228: ("SUMPRODUCT", None),
    229: ("SINH", 1), 230: ("COSH", 1), 231: ("TANH", 1), 232: ("ASINH", 1), 233: ("ACOSH", 1), 234: ("ATANH", 1),
    235: ("DGET", 3), 244: ("INFO", 1), 247: ("DB", None), 252: ("FREQUENCY", 2), 261: ("ERROR.TYPE", 1),
    269: ("AVEDEV", None), 270: ("BETADIST", None), 271: ("GAMMALN", 1), 272: ("BETAINV", None),
    273: ("BINOMDIST", 4), 274: ("CHIDIST", 2), 275: ("CHIINV", 2), 276: ("COMBIN", 2), 277: ("CONFIDENCE", 3),
    278: ("CRITBINOM", 3), 279: ("EVEN", 1), 280: ("EXPONDIST", 3), 281: ("FDIST", 3), 282: ("FINV", 3),
    283: ("FISHER", 1), 284: ("FISHERINV", 1), 285: ("FLOOR", 2), 286: ("GAMMADIST", 4), 287: ("GAMMAINV", 3),
    288: ("CEILING", 2), 289: ("HYPGEOMDIST", 4), 290: ("LOGNORMDIST", 3), 291: ("LOGINV", 3),
    292: ("NEGBINOMDIST", 3), 293: ("NORMDIST", 4), 294: ("NORMSDIST", 1), 295: ("NORMINV", 3), 296: ("NORMSINV", 1),
    297: ("STANDARDIZE", 3), 298: ("ODD", 1), 299: ("PERMUT", 2), 300: ("POISSON", 3), 301: ("TDIST", 3),
    302: ("WEIBULL", 4), 303: ("SUMXMY2", 2), 304: ("SUMX2MY2", 2), 305: ("SUMX2PY2", 2), 306: ("CHITEST", 2),
    307: ("CORREL", 2), 308: ("COVAR", 2), 309: ("FORECAST", 3), 310: ("FTEST", 2), 311: ("INTERCEPT", 2),
    312: ("PEARSON", 2), 313: ("RSQ", 2), 314: ("STEYX", 2), 315: ("SLOPE", 2), 316: ("TTEST", 4), 317: ("PROB", None),
    318: ("DEVSQ", None), 319: ("GEOMEAN", None), 320: ("HARMEAN", None), 321: ("SUMSQ", None), 322: ("KURT", None),
    323: ("SKEW", None), 324: ("ZTEST", None), 325: ("LARGE", 2), 326: ("SMALL", 2), 327: ("QUARTILE", 2),
    328: ("PERCENTILE", 2), 329: ("PERCENTRANK", None), 330: ("MODE", None), 331: ("TRIMMEAN", 2), 332: ("TINV", 2),
    336: ("CONCATENATE", None), 337: ("POWER", 2), 342: ("RADIANS", 1), 343: ("DEGREES", 1), 344: ("SUBTOTAL", None),
    345: ("SUMIF", None), 346: ("COUNTIF", 2), 347: ("COUNTBLANK", 1), 350: ("ISPMT", 4), 351: ("DATEDIF", 3),
    352: ("DATESTRING", 1), 353: ("NUMBERSTRING", 2), 354: ("ROMAN", None), 358: ("GETPIVOTDATA", None),
    359: ("HYPERLINK", None), 360: ("PHONETIC", 1), 361: ("AVERAGEA", None), 362: ("MAXA", None), 363: ("MINA", None),
    364: ("STDEVPA", None), 365: ("VARPA", None), 366: ("STDEVA", None), 367: ("VARA", None),
}

_BINOP = {0x03: "+", 0x04: "-", 0x05: "*", 0x06: "/", 0x07: "^", 0x08: "&", 0x09: "<", 0x0A: "<=", 0x0B: "=",
          0x0C: ">=", 0x0D: ">", 0x0E: "<>", 0x0F: " ", 0x10: ",", 0x11: ":"}
_PREC = {":": 8, " ": 7, ",": 6, "neg": 5, "%": 4, "^": 3, "*": 2, "/": 2, "+": 1, "-": 1, "&": 0,
         "=": -1, "<": -1, ">": -1, "<=": -1, ">=": -1, "<>": -1}
ERRORS = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!", 0x1D: "#NAME?", 0x24: "#NUM!", 0x2A: "#N/A"}


def _cell(rw: int, colw: int, base: tuple[int, int] | None) -> str:
    """ptgRef の行・列（相対の印つき）を A1 の形にする。base があれば、相対は base からのずれ（ptgRefN）。"""
    col = colw & 0x3FFF
    rel_c = bool(colw & 0x4000)
    rel_r = bool(colw & 0x8000)
    r, c = rw, col
    if base is not None:
        if rel_r:
            r = base[0] + (rw - 0x10000 if rw >= 0x8000 else rw)
        if rel_c:
            c = base[1] + (col - 0x100 if col >= 0x80 and col < 0x100 else col)
        r %= 0x10000
        c %= 0x100
    return f"{'' if rel_c else '$'}{num_to_col(c + 1)}{'' if rel_r else '$'}{r + 1}"


def _q(sheet: str) -> str:
    import re

    if re.fullmatch(r"[A-Za-z_\u0080-￿][\w.\u0080-￿]*", sheet) and not re.match(r"^[A-Za-z]{1,3}\d+$", sheet):
        return sheet
    return "'" + sheet.replace("'", "''") + "'"


class Decoder:
    """book は、名前（NAME）・外部シート（EXTERNSHEET）・外部の名前を引くためのもの。"""

    def __init__(self, book):
        self.book = book

    def sheet_prefix(self, ixti: int) -> str:
        return self.book.xti_prefix(ixti)

    def decode(self, rgce: bytes, extra: bytes = b"", base: tuple[int, int] | None = None, rel: bool = False) -> str:
        """base は (行, 列)（0 始まり）。共有数式・入力規則・条件付き書式の相対参照の基準。

        rel=True のとき（共有数式・入力規則・条件付き書式）は、別シートへの参照（ptgRef3d）の相対の行・列も
        base からのずれとして読む。セルの数式と名前では、行・列はそのままの位置。"""
        st: list[tuple[str, int]] = []  # (文字列, 優先順位)
        arrays = []
        if extra:
            arrays = list(self._arrays(rgce, extra))
        ai = 0
        i = 0
        n = len(rgce)
        while i < n:
            ptg = rgce[i]
            i += 1
            base_id = ptg if ptg < 0x20 else 0x20 | (ptg & 0x1F)
            if ptg in _BINOP:
                if len(st) < 2:
                    raise Unsupported("数式のトークンが壊れています")
                b, pb = st.pop()
                a, pa = st.pop()
                op = _BINOP[ptg]
                prec = _PREC[op]
                if pa < prec:
                    a = f"({a})"
                if pb <= prec and op not in (":",):
                    b = f"({b})"
                st.append((f"{a}{op}{b}" if op != "," else f"({a},{b})", prec))
            elif ptg == 0x12:  # 単項の +
                a, pa = st.pop()
                st.append(("+" + a, _PREC["neg"]))
            elif ptg == 0x13:  # 単項の -
                a, pa = st.pop()
                st.append(("-" + (a if pa >= _PREC["neg"] else f"({a})"), _PREC["neg"]))
            elif ptg == 0x14:  # %
                a, pa = st.pop()
                st.append(((a if pa >= _PREC["%"] else f"({a})") + "%", _PREC["%"]))
            elif ptg == 0x15:  # かっこ
                a, _ = st.pop()
                st.append((f"({a})", 99))
            elif ptg == 0x16:  # 省略された引数
                st.append(("", 99))
            elif ptg == 0x17:  # 文字列
                cch = rgce[i]
                flags = rgce[i + 1]
                i += 2
                if flags & 1:
                    s = rgce[i:i + 2 * cch].decode("utf-16-le")
                    i += 2 * cch
                else:
                    s = rgce[i:i + cch].decode("latin-1")
                    i += cch
                st.append(('"' + s.replace('"', '""') + '"', 99))
            elif ptg == 0x18:
                raise Unsupported("拡張トークン（ptgElf）")
            elif ptg == 0x19:  # ptgAttr
                grbit = rgce[i]
                if grbit & 0x04:  # tAttrChoose：後ろに飛び先の表がある
                    cases = struct.unpack_from("<H", rgce, i + 1)[0]
                    i += 3 + 2 * (cases + 1)
                    continue
                i += 3
                if grbit & 0x10:  # tAttrSum：SUM(引数1つ)
                    a, _ = st.pop()
                    st.append((f"SUM({a})", 99))
                # tAttrSemi（計算のたびに変わる印）、tAttrIf、tAttrSkip、tAttrSpace は書式だけ
            elif ptg == 0x1C:
                st.append((ERRORS.get(rgce[i], "#N/A"), 99))
                i += 1
            elif ptg == 0x1D:
                st.append(("TRUE" if rgce[i] else "FALSE", 99))
                i += 1
            elif ptg == 0x1E:
                st.append((str(struct.unpack_from("<H", rgce, i)[0]), 99))
                i += 2
            elif ptg == 0x1F:
                v = struct.unpack_from("<d", rgce, i)[0]
                st.append((repr(v) if v != int(v) or abs(v) >= 1e15 else str(int(v)), 99))
                i += 8
            elif base_id == 0x20:  # 定数の配列
                i += 7
                st.append((arrays[ai] if ai < len(arrays) else "{0}", 99))
                ai += 1
            elif base_id == 0x21:  # 引数の数が決まった関数
                iftab = struct.unpack_from("<H", rgce, i)[0]
                i += 2
                name, argc = FTAB.get(iftab, (None, None))
                if name is None or argc is None:
                    raise Unsupported(f"関数の番号 {iftab}")
                args = [st.pop()[0] for _ in range(argc)][::-1]
                st.append((f"{name}({','.join(args)})", 99))
            elif base_id == 0x22:  # 引数の数が変わる関数
                argc = rgce[i] & 0x7F
                iftab = struct.unpack_from("<H", rgce, i + 1)[0] & 0x7FFF
                i += 3
                args = [st.pop()[0] for _ in range(argc)][::-1]
                if iftab == 255:  # アドイン関数・ユーザー定義関数：最初の引数が関数名
                    fname = args[0]
                    args = args[1:]
                else:
                    fname = FTAB.get(iftab, (None, None))[0]
                    if fname is None:
                        raise Unsupported(f"関数の番号 {iftab}")
                st.append((f"{fname}({','.join(args)})", 99))
            elif base_id == 0x23:  # 名前
                idx = struct.unpack_from("<I", rgce, i)[0]
                i += 4
                st.append((self.book.name_text(idx), 99))
            elif base_id == 0x24:  # セル
                rw, colw = struct.unpack_from("<HH", rgce, i)
                i += 4
                st.append((_cell(rw, colw, None), 99))
            elif base_id == 0x25:  # 範囲
                r1, r2, c1, c2 = struct.unpack_from("<HHHH", rgce, i)
                i += 8
                st.append((_area(r1, r2, c1, c2, None), 99))
            elif base_id in (0x26, 0x27, 0x28, 0x29):  # ptgMem*：後に続く式の長さだけを示す
                i += 6 if base_id in (0x26, 0x27) else 2
            elif base_id == 0x2A:
                i += 4
                st.append(("#REF!", 99))
            elif base_id == 0x2B:
                i += 8
                st.append(("#REF!", 99))
            elif base_id == 0x2C:  # 相対のセル（共有数式・入力規則）
                rw, colw = struct.unpack_from("<HH", rgce, i)
                i += 4
                st.append((_cell(rw, colw, base or (0, 0)), 99))
            elif base_id == 0x2D:
                r1, r2, c1, c2 = struct.unpack_from("<HHHH", rgce, i)
                i += 8
                st.append((_area(r1, r2, c1, c2, base or (0, 0)), 99))
            elif base_id == 0x39:  # 外部の名前（アドイン関数の名前など）
                ixti, idx = struct.unpack_from("<HH", rgce, i)
                i += 6
                st.append((self.book.extern_name(ixti, idx), 99))
            elif base_id == 0x3A:
                ixti, rw, colw = struct.unpack_from("<HHH", rgce, i)
                i += 6
                st.append((self.sheet_prefix(ixti) + _cell(rw, colw, base if rel else None), 99))
            elif base_id == 0x3B:
                ixti, r1, r2, c1, c2 = struct.unpack_from("<HHHHH", rgce, i)
                i += 10
                st.append((self.sheet_prefix(ixti) + _area(r1, r2, c1, c2, base if rel else None), 99))
            elif base_id == 0x3C:
                i += 6
                st.append(("#REF!", 99))
            elif base_id == 0x3D:
                i += 10
                st.append(("#REF!", 99))
            elif ptg == 0x01:
                raise Unsupported("共有数式・配列数式の本体を参照しています")
            else:
                raise Unsupported(f"数式のトークン 0x{ptg:02X}")
        if len(st) != 1:
            raise Unsupported("数式のトークンが壊れています")
        return st[0][0]

    def _arrays(self, rgce: bytes, extra: bytes):
        """rgcb（数式の後ろの追加データ）にある定数の配列を、出てくる順に文字列にする。"""
        pos = 0
        count = sum(1 for _ in self._array_ptgs(rgce))
        for _ in range(count):
            cols = extra[pos] + 1
            rows = struct.unpack_from("<H", extra, pos + 1)[0] + 1
            pos += 3
            out_rows = []
            for _r in range(rows):
                vals = []
                for _c in range(cols):
                    t = extra[pos]
                    pos += 1
                    if t == 0x01:
                        v = struct.unpack_from("<d", extra, pos)[0]
                        pos += 8
                        vals.append(str(int(v)) if v == int(v) else repr(v))
                    elif t == 0x02:
                        cch = struct.unpack_from("<H", extra, pos)[0]
                        flags = extra[pos + 2]
                        pos += 3
                        if flags & 1:
                            s = extra[pos:pos + 2 * cch].decode("utf-16-le")
                            pos += 2 * cch
                        else:
                            s = extra[pos:pos + cch].decode("latin-1")
                            pos += cch
                        vals.append('"' + s.replace('"', '""') + '"')
                    elif t == 0x04:
                        vals.append("TRUE" if extra[pos] else "FALSE")
                        pos += 8
                    elif t == 0x10:
                        vals.append(ERRORS.get(extra[pos], "#N/A"))
                        pos += 8
                    else:
                        pos += 8
                        vals.append("0")
                out_rows.append(",".join(vals))
            yield "{" + ";".join(out_rows) + "}"

    def _array_ptgs(self, rgce: bytes):
        # ptgArray の数だけ数える（トークンの長さを正しく読み飛ばす）
        i = 0
        while i < len(rgce):
            ptg = rgce[i]
            i += 1
            b = ptg if ptg < 0x20 else 0x20 | (ptg & 0x1F)
            if b == 0x20:
                yield i
                i += 7
                continue
            size = _SIZE.get(b)
            if ptg == 0x17:
                cch, flags = rgce[i], rgce[i + 1]
                i += 2 + (2 * cch if flags & 1 else cch)
            elif ptg == 0x19:
                grbit = rgce[i]
                if grbit & 0x04:
                    cases = struct.unpack_from("<H", rgce, i + 1)[0]
                    i += 3 + 2 * (cases + 1)
                else:
                    i += 3
            elif size is not None:
                i += size
            else:
                return


_SIZE = {0x01: 4, 0x02: 4, 0x03: 0, 0x04: 0, 0x05: 0, 0x06: 0, 0x07: 0, 0x08: 0, 0x09: 0, 0x0A: 0, 0x0B: 0, 0x0C: 0,
         0x0D: 0, 0x0E: 0, 0x0F: 0, 0x10: 0, 0x11: 0, 0x12: 0, 0x13: 0, 0x14: 0, 0x15: 0, 0x16: 0, 0x1C: 1, 0x1D: 1,
         0x1E: 2, 0x1F: 8, 0x21: 2, 0x22: 3, 0x23: 4, 0x24: 4, 0x25: 8, 0x26: 6, 0x27: 6, 0x28: 2, 0x29: 2, 0x2A: 4,
         0x2B: 8, 0x2C: 4, 0x2D: 8, 0x39: 6, 0x3A: 6, 0x3B: 10, 0x3C: 6, 0x3D: 10}


def _area(r1, r2, c1w, c2w, base) -> str:
    if base is None and r1 == 0 and r2 == 0xFFFF:
        # 列全体（A:A）
        ca = ("" if c1w & 0x4000 else "$") + num_to_col((c1w & 0xFF) + 1)
        cb = ("" if c2w & 0x4000 else "$") + num_to_col((c2w & 0xFF) + 1)
        return f"{ca}:{cb}"
    if base is None and (c1w & 0xFF) == 0 and (c2w & 0xFF) == 0xFF:
        # 行全体（1:3）
        ra = ("" if c1w & 0x8000 else "$") + str(r1 + 1)
        rb = ("" if c2w & 0x8000 else "$") + str(r2 + 1)
        return f"{ra}:{rb}"
    return f"{_cell(r1, c1w, base)}:{_cell(r2, c2w, base)}"
