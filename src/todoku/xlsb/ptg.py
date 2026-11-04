"""BIFF12（.xlsb、MS-XLSB 2.5.98）の数式のトークン列を、数式の文字列に戻す。

文字列にしたあとは、xlsx・xls と同じ構文解析・計算で確かめる（読み違えると、保存された計算結果と合わないので分かる）。
BIFF8 との違い（MS-XLSB 2.5.98.88 ほか）：
- 行は 4 バイト（20 ビット）、列は 14 ビット＋相対の印 2 つ（RgceLoc・RgceArea）
- 文字列（PtgStr）は長さ 2 バイト＋ UTF-16（印のバイトがない）
- 名前（PtgName）は 4 バイト。外部の名前（PtgNameX）は 2＋4 バイト
- 配列の定数は rgcb の PtgExtraArray（行数 4・列数 4・SerAr）。PtgExp の列は rgcb の PtgExtraCol
- 関数の番号（Ftab）は BIFF8 に 368〜484 が足された
読めないトークンは Unsupported（窓には保存された計算結果を出し、数式は「未対応」と示す）。
"""

from __future__ import annotations

import struct

from ..formula.parser import num_to_col
from ..formula.values import Unsupported
from ..xls.ptg import _BINOP, _PREC, ERRORS, FTAB, _q

# MS-XLSB 2.5.98.10 Ftab：BIFF8 の FTAB にない番号。名前と、決まった引数の数（None は可変：PtgFuncVar で書かれる）
_MORE = {368: ("BAHTTEXT", 1), 379: ("RTD", None), 380: ("CUBEVALUE", None), 381: ("CUBEMEMBER", None),
         382: ("CUBEMEMBERPROPERTY", None), 383: ("CUBERANKEDMEMBER", None), 384: ("HEX2BIN", None), 385: ("HEX2DEC", 1),
         386: ("HEX2OCT", None), 387: ("DEC2BIN", None), 388: ("DEC2HEX", None), 389: ("DEC2OCT", None),
         390: ("OCT2BIN", None), 391: ("OCT2HEX", None), 392: ("OCT2DEC", 1), 393: ("BIN2DEC", 1), 394: ("BIN2OCT", None),
         395: ("BIN2HEX", None), 396: ("IMSUB", 2), 397: ("IMDIV", 2), 398: ("IMPOWER", 2), 399: ("IMABS", 1),
         400: ("IMSQRT", 1), 401: ("IMLN", 1), 402: ("IMLOG2", 1), 403: ("IMLOG10", 1), 404: ("IMSIN", 1), 405: ("IMCOS", 1),
         406: ("IMEXP", 1), 407: ("IMARGUMENT", 1), 408: ("IMCONJUGATE", 1), 409: ("IMAGINARY", 1), 410: ("IMREAL", 1),
         411: ("COMPLEX", None), 412: ("IMSUM", None), 413: ("IMPRODUCT", None), 414: ("SERIESSUM", 4), 415: ("FACTDOUBLE", 1),
         416: ("SQRTPI", 1), 417: ("QUOTIENT", 2), 418: ("DELTA", None), 419: ("GESTEP", None), 420: ("ISEVEN", 1),
         421: ("ISODD", 1), 422: ("MROUND", 2), 423: ("ERF", None), 424: ("ERFC", 1), 425: ("BESSELJ", 2), 426: ("BESSELK", 2),
         427: ("BESSELY", 2), 428: ("BESSELI", 2), 429: ("XIRR", None), 430: ("XNPV", 3), 431: ("PRICEMAT", None),
         432: ("YIELDMAT", None), 433: ("INTRATE", None), 434: ("RECEIVED", None), 435: ("DISC", None), 436: ("PRICEDISC", None),
         437: ("YIELDDISC", None), 438: ("TBILLEQ", 3), 439: ("TBILLPRICE", 3), 440: ("TBILLYIELD", 3), 441: ("PRICE", None),
         442: ("YIELD", None), 443: ("DOLLARDE", 2), 444: ("DOLLARFR", 2), 445: ("NOMINAL", 2), 446: ("EFFECT", 2),
         447: ("CUMPRINC", 6), 448: ("CUMIPMT", 6), 449: ("EDATE", 2), 450: ("EOMONTH", 2), 451: ("YEARFRAC", None),
         452: ("COUPDAYBS", None), 453: ("COUPDAYS", None), 454: ("COUPDAYSNC", None), 455: ("COUPNCD", None),
         456: ("COUPNUM", None), 457: ("COUPPCD", None), 458: ("DURATION", None), 459: ("MDURATION", None),
         460: ("ODDLPRICE", None), 461: ("ODDLYIELD", None), 462: ("ODDFPRICE", None), 463: ("ODDFYIELD", None),
         464: ("RANDBETWEEN", 2), 465: ("WEEKNUM", None), 466: ("AMORDEGRC", None), 467: ("AMORLINC", None),
         469: ("ACCRINT", None), 470: ("ACCRINTM", None), 471: ("WORKDAY", None), 472: ("NETWORKDAYS", None), 473: ("GCD", None),
         474: ("MULTINOMIAL", None), 475: ("LCM", None), 476: ("FVSCHEDULE", 2), 477: ("CUBEKPIMEMBER", None),
         478: ("CUBESET", None), 479: ("CUBESETCOUNT", 1), 480: ("IFERROR", 2), 481: ("COUNTIFS", None), 482: ("SUMIFS", None),
         483: ("AVERAGEIF", None), 484: ("AVERAGEIFS", None)}
FTAB12 = {**FTAB, **_MORE}


def _ref(row: int, colw: int, base: tuple[int, int] | None) -> str:
    """row（4 バイト）と ColRelU（列 14 ビット・列の相対・行の相対）。base があれば、相対はそこからのずれ（RgceLocRel）。"""
    col = colw & 0x3FFF
    rel_c, rel_r = bool(colw & 0x4000), bool(colw & 0x8000)
    if base is not None:
        if rel_r:
            row = (base[0] + (row - (1 << 32) if row >= 1 << 31 else row)) % (1 << 20)
        if rel_c:
            col = (base[1] + (col - (1 << 14) if col >= 1 << 13 else col)) % (1 << 14)
    return f"{'' if rel_c else '$'}{num_to_col(col + 1)}{'' if rel_r else '$'}{row + 1}"


def _area(r1: int, r2: int, c1w: int, c2w: int, base) -> str:
    if (c1w & 0x3FFF) == 0 and (c2w & 0x3FFF) == 0x3FFF and base is None:  # 行全体
        return f"{'' if c1w & 0x8000 else '$'}{r1 + 1}:{'' if c2w & 0x8000 else '$'}{r2 + 1}"
    if r1 == 0 and r2 == 0xFFFFF and base is None:  # 列全体
        return (f"{'' if c1w & 0x4000 else '$'}{num_to_col((c1w & 0x3FFF) + 1)}:"
                f"{'' if c2w & 0x4000 else '$'}{num_to_col((c2w & 0x3FFF) + 1)}")
    return f"{_ref(r1, c1w, base)}:{_ref(r2, c2w, base)}"


class Decoder12:
    """book は、name_text・extern_name・xti_prefix を持つもの。"""

    def __init__(self, book):
        self.book = book

    def decode(self, rgce: bytes, extra: bytes = b"", base: tuple[int, int] | None = None) -> str:
        try:
            return self._decode(rgce, extra, base)
        except (struct.error, IndexError) as e:
            raise Unsupported(f"数式のトークンが壊れています（{e}）") from None

    def _decode(self, rgce: bytes, extra: bytes, base) -> str:
        st: list[tuple[str, int]] = []
        xp = 0  # rgcb の読む位置（トークンの順に並ぶ。MS-XLSB 2.5.98.87）
        i, n = 0, len(rgce)
        while i < n:
            ptg = rgce[i]
            i += 1
            bid = ptg if ptg < 0x20 else 0x20 | (ptg & 0x1F)
            if ptg in _BINOP:
                b, pb = st.pop()
                a, pa = st.pop()
                op = _BINOP[ptg]
                prec = _PREC[op]
                a = f"({a})" if pa < prec else a
                b = f"({b})" if pb <= prec and op != ":" else b
                st.append((f"{a}{op}{b}" if op != "," else f"({a},{b})", prec))
            elif ptg == 0x01:  # PtgExp はここには来ない（本体は book が引く）
                raise Unsupported("共有数式の指す先")
            elif ptg == 0x02:
                raise Unsupported("データテーブル（PtgTbl）")
            elif ptg == 0x12:
                a, _pa = st.pop()
                st.append(("+" + a, _PREC["neg"]))
            elif ptg == 0x13:
                a, pa = st.pop()
                st.append(("-" + (a if pa >= _PREC["neg"] else f"({a})"), _PREC["neg"]))
            elif ptg == 0x14:
                a, pa = st.pop()
                st.append(((a if pa >= _PREC["%"] else f"({a})") + "%", _PREC["%"]))
            elif ptg == 0x15:
                a, _pa = st.pop()
                st.append((f"({a})", 99))
            elif ptg == 0x16:
                st.append(("", 99))
            elif ptg == 0x17:
                cch = struct.unpack_from("<H", rgce, i)[0]
                i += 2
                s = rgce[i:i + 2 * cch].decode("utf-16-le", "replace")
                i += 2 * cch
                st.append(('"' + s.replace('"', '""') + '"', 99))
            elif ptg == 0x18:
                raise Unsupported("拡張トークン（表の構造化参照など）")
            elif ptg == 0x19:
                grbit = rgce[i]
                if grbit & 0x04:  # PtgAttrChoose：飛び先の表
                    cases = struct.unpack_from("<H", rgce, i + 1)[0]
                    i += 3 + 2 * (cases + 1)
                    continue
                i += 3
                if grbit & 0x10:
                    a, _pa = st.pop()
                    st.append((f"SUM({a})", 99))
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
            elif bid == 0x20:  # PtgArray：14 バイトの飛ばす所。値は rgcb
                i += 14
                text, xp = _array(extra, xp)
                st.append((text, 99))
            elif bid == 0x21:
                iftab = struct.unpack_from("<H", rgce, i)[0]
                i += 2
                name, argc = FTAB12.get(iftab, (None, None))
                if name is None or argc is None:
                    raise Unsupported(f"関数の番号 {iftab}")
                args = [st.pop()[0] for _ in range(argc)][::-1]
                st.append((f"{name}({','.join(args)})", 99))
            elif bid == 0x22:
                argc = rgce[i] & 0x7F
                iftab = struct.unpack_from("<H", rgce, i + 1)[0] & 0x7FFF
                i += 3
                args = [st.pop()[0] for _ in range(argc)][::-1]
                if iftab == 255:
                    fname, args = args[0], args[1:]
                    if fname.startswith("_xlfn."):
                        fname = fname[6:]
                else:
                    fname = FTAB12.get(iftab, (None, None))[0]
                    if fname is None:
                        raise Unsupported(f"関数の番号 {iftab}")
                st.append((f"{fname}({','.join(args)})", 99))
            elif bid == 0x23:
                st.append((self.book.name_text(struct.unpack_from("<I", rgce, i)[0]), 99))
                i += 4
            elif bid == 0x24:
                row, colw = struct.unpack_from("<IH", rgce, i)
                i += 6
                st.append((_ref(row, colw, None), 99))
            elif bid == 0x25:
                r1, r2, c1, c2 = struct.unpack_from("<IIHH", rgce, i)
                i += 12
                st.append((_area(r1, r2, c1, c2, None), 99))
            elif bid in (0x26, 0x27):  # PtgMemArea・PtgMemErr：4 バイトの飛ばす所と長さ。後の式がそのまま続く
                i += 6
                if bid == 0x26:  # PtgExtraMem：rgcb の count と範囲（読み飛ばす）
                    cnt = struct.unpack_from("<I", extra, xp)[0]
                    xp += 4 + 16 * cnt
            elif bid in (0x28, 0x29):
                i += 6 if bid == 0x28 else 2
            elif bid == 0x2A:
                i += 6
                st.append(("#REF!", 99))
            elif bid == 0x2B:
                i += 12
                st.append(("#REF!", 99))
            elif bid == 0x2C:
                row, colw = struct.unpack_from("<IH", rgce, i)
                i += 6
                st.append((_ref(row, colw, base or (0, 0)), 99))
            elif bid == 0x2D:
                r1, r2, c1, c2 = struct.unpack_from("<IIHH", rgce, i)
                i += 12
                st.append((_area(r1, r2, c1, c2, base or (0, 0)), 99))
            elif bid == 0x39:
                ixti, idx = struct.unpack_from("<HI", rgce, i)
                i += 6
                st.append((self.book.extern_name(ixti, idx), 99))
            elif bid == 0x3A:
                ixti, row, colw = struct.unpack_from("<HIH", rgce, i)
                i += 8
                st.append((self.book.xti_prefix(ixti) + _ref(row, colw, None), 99))
            elif bid == 0x3B:
                ixti, r1, r2, c1, c2 = struct.unpack_from("<HIIHH", rgce, i)
                i += 14
                st.append((self.book.xti_prefix(ixti) + _area(r1, r2, c1, c2, None), 99))
            elif bid == 0x3C:
                ixti = struct.unpack_from("<H", rgce, i)[0]
                i += 8
                st.append((self.book.xti_prefix(ixti) + "#REF!", 99))
            elif bid == 0x3D:
                ixti = struct.unpack_from("<H", rgce, i)[0]
                i += 14
                st.append((self.book.xti_prefix(ixti) + "#REF!", 99))
            else:
                raise Unsupported(f"トークン 0x{ptg:02X}")
        if len(st) != 1:
            raise Unsupported("数式のトークンが壊れています")
        return st[0][0]


def _array(x: bytes, p: int) -> tuple[str, int]:
    """PtgExtraArray（行数・列数・SerAr）→ {1,2;3,4}。SerAr は MS-XLSB 2.5.98.93〜97。"""
    rows, cols = struct.unpack_from("<II", x, p)
    p += 8
    if rows * cols > 1 << 20:
        raise Unsupported("配列の定数が大きすぎます")
    out = []
    for _r in range(rows):
        row = []
        for _c in range(cols):
            t = x[p]
            p += 1
            if t == 0:
                v = struct.unpack_from("<d", x, p)[0]
                p += 8
                row.append(repr(v) if v != int(v) or abs(v) >= 1e15 else str(int(v)))
            elif t == 1:
                cch = struct.unpack_from("<H", x, p)[0]
                p += 2
                row.append('"' + x[p:p + 2 * cch].decode("utf-16-le", "replace").replace('"', '""') + '"')
                p += 2 * cch
            elif t == 2:
                row.append("TRUE" if x[p] else "FALSE")
                p += 1
            elif t == 4:
                row.append(ERRORS.get(x[p], "#N/A"))
                p += 4
            else:
                raise Unsupported(f"配列の定数の種類 {t}")
        out.append(",".join(row))
    return "{" + ";".join(out) + "}", p


__all__ = ["Decoder12", "FTAB12", "_q"]
