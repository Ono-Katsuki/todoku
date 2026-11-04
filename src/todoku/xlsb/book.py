"""Excel のバイナリのブック（.xlsb、MS-XLSB）を読み、xlsx と同じ形（シート・セル・数式・入力規則）にする。

入れ物は xlsx と同じ OPC の ZIP で、部品（xl/workbook.bin・xl/worksheets/sheetN.bin・xl/styles.bin・xl/sharedStrings.bin）が
BIFF12 のレコードの並び（型 1〜2 バイト・長さ 1〜4 バイト。どちらも 7 ビットずつ。MS-XLSB 2.1.4）。
読んだあとは、窓・計算（explain_calculation・try_calculation）・表の取り出しを xlsx と同じ処理で行う。

図形は BrtDrawing の先の DrawingML（xlsx と同じ部品）を読む。条件付き書式は BrtBeginConditionalFormatting・BrtBeginCFRule の範囲と条件を読む。
読まないもの（窓の unread に書き、段階は C）：条件付き書式の見た目（BrtDXF の色など）・表の構造化参照の数式。
"""

from __future__ import annotations

import struct
import xml.etree.ElementTree as ET
from posixpath import dirname, join, normpath

from ..formula.numfmt import BUILTIN
from ..formula.parser import addr, parse
from ..formula.values import ExcelError, Unsupported
from ..xls.book import _DV_OP, _DV_TYPE, _PALETTE, _rk
from ..xls.ptg import ERRORS
from ..xlsx.workbook import Cell, CondRule, DefinedName, Sheet, Style, Validation, _area_ranges, _header_text, apply_tint, theme_colors
from ..zipraw import RawZip
from .ptg import Decoder12, _q

# MS-XLSB 2.3 の番号
ROW, C_BLANK, C_RK, C_ERROR, C_BOOL, C_REAL, C_ST, C_ISST, F_ST, F_NUM, F_BOOL, F_ERR = range(12)
C_RSTR = 62
CELLS = frozenset({C_BLANK, C_RK, C_ERROR, C_BOOL, C_REAL, C_ST, C_ISST, C_RSTR, F_ST, F_NUM, F_BOOL, F_ERR})
SST_ITEM, NAME, FONT, FMT, FILL, BORDER, XF = 19, 39, 43, 44, 45, 46, 47
COLINFO, DVAL = 60, 64
BUNDLESH, WBPROP, MERGE, BEGIN_DATA, END_DATA = 156, 153, 176, 145, 146
EXTERNSHEET, SUP_BOOK, SUP_SELF, SUP_SAME, SUP_TABS, SUP_PLACEHOLDER = 362, 355, 357, 358, 359, 361
ARR, SHR = 426, 427
BEGIN_CELLXFS, END_CELLXFS = 617, 618
HLINK, SHEETPROT, BOOKPROT, WSFMT, HEADERFOOTER = 494, 535, 534, 485, 479
DRAWING, LEGACYDRAWING, BEGIN_CF, BEGIN_CFRULE = 550, 551, 461, 463
_CF_OPER = {1: "between", 2: "notBetween", 3: "equal", 4: "notEqual", 5: "greaterThan", 6: "lessThan", 7: "greaterThanOrEqual", 8: "lessThanOrEqual"}
_CF_OPER_JA = {"between": "の間", "notBetween": "の外", "equal": "と同じ", "notEqual": "と違う", "greaterThan": "より大きい", "lessThan": "より小さい",
               "greaterThanOrEqual": "以上", "lessThanOrEqual": "以下"}
_CF_TYPE = {1: "cellIs", 3: "colorScale", 4: "dataBar", 5: "top10", 6: "iconSet"}
_CF_TEMPLATE = {1: "expression", 7: "uniqueValues", 8: "containsText", 9: "containsBlanks", 10: "notContainsBlanks", 11: "containsErrors", 12: "notContainsErrors",
                25: "aboveAverage", 26: "aboveAverage", 27: "duplicateValues", 29: "aboveAverage", 30: "aboveAverage"}
LISTPART = 661
C_AUTHOR, C_BEGIN, C_TEXT = 632, 635, 637

_ST = {0: "visible", 1: "hidden", 2: "veryHidden"}


class XlsbError(ValueError):
    pass


def records(data: bytes):
    """(型, 中身の始まり, 中身の長さ)。長さが部品の外に出たら XlsbError（読めた所まででごまかさない）。"""
    pos, n = 0, len(data)
    while pos < n:
        b0 = data[pos]
        pos += 1
        rid = b0 & 0x7F
        if b0 & 0x80:
            if pos >= n:
                raise XlsbError("レコードの型が途中で終わっています")
            rid |= (data[pos] & 0x7F) << 7
            pos += 1
        ln = shift = 0
        for _ in range(4):
            if pos >= n:
                raise XlsbError("レコードの長さが途中で終わっています")
            b = data[pos]
            pos += 1
            ln |= (b & 0x7F) << shift
            shift += 7
            if not b & 0x80:
                break
        if pos + ln > n:
            raise XlsbError("レコードが部品の外に出ています")
        yield rid, pos, ln
        pos += ln


def wstr(b, p: int) -> tuple[str, int]:
    """XLWideString（長さ 4 バイト・UTF-16）。"""
    cch = struct.unpack_from("<I", b, p)[0]
    p += 4
    if cch == 0xFFFFFFFF:  # XLNullableWideString の null
        return "", p
    if p + 2 * cch > len(b):
        raise XlsbError("文字列がレコードの外に出ています")
    return bytes(b[p:p + 2 * cch]).decode("utf-16-le", "replace"), p + 2 * cch


def rich(b, p: int = 0) -> str:
    """RichStr（MS-XLSB 2.5.122）の字。"""
    return wstr(b, p + 1)[0]


def _rfx(b, p: int = 0) -> tuple[int, int, int, int]:
    """UncheckedRfX（行の始め・終わり、列の始め・終わり）→ (r1, c1, r2, c2)、1 始まり。"""
    r1, r2, c1, c2 = struct.unpack_from("<IIII", b, p)
    return r1 + 1, c1 + 1, r2 + 1, c2 + 1


def _cf_text(rule: CondRule) -> str:
    where = "・".join(f"{addr(a, b)}:{addr(c, d)}" if (a, b) != (c, d) else addr(a, b) for a, b, c, d in rule.sqref[:3])
    what = rule.type
    if rule.type == "cellIs" and rule.operator:
        what = "〜".join(rule.formulas) + _CF_OPER_JA.get(rule.operator, rule.operator)
    elif rule.type == "expression" and rule.formulas:
        what = f"式 {rule.formulas[0]}"
    return f"{where} が {what}"


class XlsbWorkbook:
    kind = "xlsb"

    def __init__(self, blob: bytes, filename: str = "", strings: list[str] | None = None):
        """strings：読んである共有文字列（lazy/sheetdocs：小さなブックごとに読み直さない）。"""
        self.filename = filename
        self.zip = RawZip(blob)
        self.sheets: list[Sheet] = []
        self.names: list[DefinedName] = []
        self.shared_strings: list[str] = strings if strings is not None else []
        self._strings_given = strings is not None
        self.phonetic: dict = {}
        self.cell_sst: dict = {}
        self.rich_cells: dict = {}
        self.styles: list[Style] = []
        self.dxfs: list = []
        self.tables: dict = {}
        self.external: dict = {}
        self.external_links: list[str] = []
        self.date1904 = False
        self.calc_full_on_load = False
        self.workbook_protected = False
        self.has_macros = self.zip.has("xl/vbaProject.bin")
        self._vba = None
        self.wb_path = "xl/workbook.bin"
        self._supbooks: list[dict] = []
        self._xti: list[tuple[int, int, int]] = []
        self._name_raw: list[tuple] = []
        self.decoder = Decoder12(self)
        self.value_records: dict[tuple[str, int, int], tuple[str, int, int, int]] = {}  # (シート, 行, 列) → (部品, 型, 値の位置, 値の長さ)
        self.record_counts: dict[str, int] = {}  # 検査用：シートごとの、値のあるセルのレコードの数
        self.unread_notes: list[str] = []
        self._load()

    # ---------------------------------------------------------- xlsx と同じ問い合わせ
    @property
    def vba(self):
        if self._vba is None and self.has_macros:
            from ..vba import read_project

            try:
                self._vba = read_project(self.zip.read("xl/vbaProject.bin"))
            except Exception as e:  # noqa: BLE001
                self._vba = e
        return self._vba if not isinstance(self._vba, Exception) else None

    def sheet(self, name: str) -> Sheet:
        for s in self.sheets:
            if s.name == name:
                return s
        low = name.lower()
        for s in self.sheets:
            if s.name.lower() == low:
                return s
        raise KeyError(name)

    def has_sheet(self, name: str) -> bool:
        return any(s.name.lower() == name.lower() for s in self.sheets)

    def style(self, idx: int) -> Style:
        return self.styles[idx] if 0 <= idx < len(self.styles) else Style("General", True, False)

    def defined_name(self, name: str, sheet: str | None):
        low = name.lower()
        if sheet:
            for d in self.names:
                if d.name.lower() == low and d.sheet_scope == sheet:
                    return d
        for d in self.names:
            if d.name.lower() == low and d.sheet_scope is None:
                return d
        return None

    # ---------------------------------------------------------- 数式から引くもの
    def xti_prefix(self, ixti: int) -> str:
        if ixti >= len(self._xti):
            return "#REF!"
        isb, first, last = self._xti[ixti]
        sb = self._supbooks[isb] if isb < len(self._supbooks) else None
        if sb is None:
            return "#REF!"
        if sb["kind"] == "self":
            if first < 0 or first >= len(self.sheets):
                return "#REF!"
            if last != first:
                raise Unsupported("複数シートにまたがる参照")
            return _q(self.sheets[first].name) + "!"
        if sb["kind"] == "external":
            name = sb["sheets"][first] if 0 <= first < len(sb["sheets"]) else ""
            return "'" + f"[{isb + 1}]{name}".replace("'", "''") + "'!"
        raise Unsupported("外部の参照")

    def name_text(self, idx: int) -> str:
        return self._name_raw[idx - 1][0] if 1 <= idx <= len(self._name_raw) else "#NAME?"

    def extern_name(self, ixti: int, idx: int) -> str:
        if ixti < len(self._xti):
            isb = self._xti[ixti][0]
            if isb < len(self._supbooks):
                nm = self._supbooks[isb]["names"]
                if 1 <= idx <= len(nm):
                    return nm[idx - 1]
        return "#NAME?"

    # ---------------------------------------------------------- 読み込み
    def _rels(self, part: str) -> dict[str, tuple[str, str]]:
        """rId → (関係の種類の最後の語, 部品の名前)。"""
        path = join(dirname(part), "_rels", part.rsplit("/", 1)[-1] + ".rels")
        if not self.zip.has(path):
            return {}
        out = {}
        for r in ET.fromstring(self.zip.read(path)):
            t = r.get("Target") or ""
            if r.get("TargetMode") == "External":
                out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], t)
                continue
            out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], t.lstrip("/") if t.startswith("/") else normpath(join(dirname(part), t)))
        return out

    def _load(self) -> None:
        if not self.zip.has(self.wb_path):
            raise XlsbError("xl/workbook.bin がありません")
        rels = self._rels(self.wb_path)
        data = self.zip.read(self.wb_path)
        bundles = []
        for rid, p, ln in records(data):
            b = data[p:p + ln]
            if rid == BUNDLESH:
                st, _tab = struct.unpack_from("<II", b, 0)
                relid, q = wstr(b, 8)
                name, _q2 = wstr(b, q)
                bundles.append((st, relid, name))
            elif rid == WBPROP:
                self.date1904 = bool(struct.unpack_from("<I", b, 0)[0] & 1)
            elif rid == SUP_SELF or rid == SUP_SAME:
                self._supbooks.append({"kind": "self", "sheets": [], "names": []})
            elif rid == SUP_BOOK:
                relid = wstr(b, 0)[0]
                self.external_links.append(rels.get(relid, ("", relid))[1])
                self._supbooks.append({"kind": "external", "sheets": [], "names": []})
            elif rid == SUP_TABS and self._supbooks:
                cnt = struct.unpack_from("<I", b, 0)[0]
                q, sheets = 4, []
                for _ in range(cnt):
                    s, q = wstr(b, q)
                    sheets.append(s)
                self._supbooks[-1]["sheets"] = sheets
            elif rid == SUP_PLACEHOLDER and self._supbooks:
                self._supbooks[-1]["names"].append(wstr(b, 0)[0])
            elif rid == EXTERNSHEET:
                cnt = struct.unpack_from("<I", b, 0)[0]
                self._xti = [struct.unpack_from("<Iii", b, 4 + 12 * k) for k in range(cnt)]
            elif rid == NAME:
                flags = struct.unpack_from("<I", b, 0)[0]
                itab = struct.unpack_from("<I", b, 5)[0]
                name, q = wstr(b, 9)
                cce = struct.unpack_from("<I", b, q)[0]
                rgce = b[q + 4:q + 4 + cce]
                q += 4 + cce
                cb = struct.unpack_from("<I", b, q)[0]
                extra = b[q + 4:q + 4 + cb]
                if flags & 0x20 and not name.startswith("_xlnm."):  # fBuiltin
                    name = "_xlnm." + name
                self._name_raw.append((name, itab, rgce, extra, bool(flags & 1)))
        self._load_styles()
        self._load_sst()
        kinds = {}
        for st, relid, name in bundles:
            typ, part = rels.get(relid, ("", ""))
            kinds[name] = typ
            state = _ST.get(st, "visible")
            if typ != "worksheet":
                self.sheets.append(Sheet(name=name, path=part, state=f"non-worksheet:{typ or '不明'}"))
            else:
                self.sheets.append(Sheet(name=name, path=part, state=state))
        names = [s.name for s in self.sheets]
        for nm, itab, rgce, extra, hidden in self._name_raw:
            scope = names[itab] if itab < len(names) else None
            try:
                text = self.decoder.decode(rgce, extra) if rgce else ""
                ast, err = (parse(text) if text else None), None
            except Unsupported as e:
                text, ast, err = "", None, f"未対応: {e}"
            except Exception as e:  # noqa: BLE001
                text, ast, err = "", None, f"構文エラー: {e}"
            self.names.append(DefinedName(nm, text, scope, hidden, ast, err))
        for d in self.names:
            if d.name == "_xlnm.Print_Area" and d.sheet_scope:
                sh = self.sheet(d.sheet_scope)
                sh.print_area = d.text
                sh.print_ranges = _area_ranges(d.text)
        for sh in self.sheets:
            if sh.state.startswith("non-worksheet"):
                continue
            if not self.zip.has(sh.path):
                raise XlsbError(f"シート「{sh.name}」の部品がありません：{sh.path}")
            self._load_sheet(sh)

    def _load_styles(self) -> None:
        path = next((t for typ, t in self._rels(self.wb_path).values() if typ == "styles"), "xl/styles.bin")
        self._theme = None
        for typ, t in self._rels(self.wb_path).values():
            if typ == "theme" and self.zip.has(t):
                try:
                    self._theme = theme_colors(ET.fromstring(self.zip.read(t)))
                except ET.ParseError:
                    self._theme = None
        if not self.zip.has(path):
            return
        data = self.zip.read(path)
        fmts, fonts, fills, borders = dict(BUILTIN), [], [], []
        in_xfs = False
        for rid, p, ln in records(data):
            b = data[p:p + ln]
            if rid == FMT:
                fmts[struct.unpack_from("<H", b, 0)[0]] = wstr(b, 2)[0]
            elif rid == FONT:
                grbit = struct.unpack_from("<H", b, 2)[0]
                fonts.append((self._color(b[12:20]), bool(grbit & 0x08)))  # fStrikeout は 3 ビット目（MS-XLSB 2.4.690）
            elif rid == FILL:
                fls = struct.unpack_from("<I", b, 0)[0]
                fills.append(self._color(b[4:12]) if fls == 1 else None)
            elif rid == BORDER:
                # flags 1 バイトのあと、上・下・左・右・斜めの Blxf（線の種類 1・予約 1・色 8）
                bits = sum(v for v, off in ((4, 1), (8, 11), (1, 21), (2, 31)) if off < ln and b[off])
                borders.append(bits)
            elif rid == BEGIN_CELLXFS:
                in_xfs = True
            elif rid == END_CELLXFS:
                in_xfs = False
            elif rid == XF and in_xfs:
                _par, ifmt, ifnt, ifill, ibrd = struct.unpack_from("<HHHHH", b, 0)
                bits = struct.unpack_from("<H", b, 12)[0]
                fc, strike = fonts[ifnt] if ifnt < len(fonts) else (None, False)
                self.styles.append(Style(fmts.get(ifmt, "General"), bool(bits & 0x1000), bool(bits & 0x2000), fc,
                                         fills[ifill] if ifill < len(fills) else None,
                                         borders[ibrd] if ibrd < len(borders) else 0, strike))

    def _color(self, c: bytes) -> str | None:
        """BrtColor（8 バイト）：種類（0 自動・1 色番号・2 RGB・3 テーマ）・番号・濃淡・R・G・B・A。"""
        if len(c) < 8:
            return None
        typ = c[0] >> 1
        if typ == 2:
            return f"{c[4]:02X}{c[5]:02X}{c[6]:02X}"
        if typ == 1:
            i = c[1]
            if i < 8:
                return ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF"][i]
            return _PALETTE[i - 8] if 8 <= i < 8 + len(_PALETTE) else None
        if typ == 3 and self._theme and c[1] < len(self._theme):
            tint = struct.unpack_from("<h", c, 2)[0] / 32767
            base = self._theme[c[1]]
            return apply_tint(base, tint) if tint else base
        return None

    def _load_sst(self) -> None:
        if self._strings_given:
            return
        path = next((t for typ, t in self._rels(self.wb_path).values() if typ == "sharedStrings"), "xl/sharedStrings.bin")
        if not self.zip.has(path):
            return
        data = self.zip.read(path)
        for rid, p, ln in records(data):
            if rid == SST_ITEM:
                self.shared_strings.append(rich(data[p:p + ln]))

    def _put(self, sh: Sheet, row: int, b, xf_pos: int, value) -> Cell:
        c = struct.unpack_from("<I", b, xf_pos - 4)[0]
        xf = struct.unpack_from("<I", b, xf_pos)[0] & 0xFFFFFF
        cell = Cell(sh.name, row + 1, c + 1, style=xf)
        sh.cells[(row + 1, c + 1)] = cell
        cell.value = value
        if isinstance(value, str):
            cell.t = "s"
        elif isinstance(value, bool):
            cell.t = "b"
        sh.max_row = max(sh.max_row, row + 1)
        sh.max_col = max(sh.max_col, c + 1)
        return cell

    def _load_sheet(self, sh: Sheet) -> None:
        data = self.zip.read(sh.path)
        rels = self._rels(sh.path)
        row = -1
        shared: dict = {}
        arrays: dict = {}
        formulas = []
        n_vals = 0
        cf = 0
        cf_ranges: list = []
        cf_bad = 0
        last_formula = None
        for rid, p, ln in records(data):
            b = data[p:p + ln]
            if rid == ROW:
                row = struct.unpack_from("<I", b, 0)[0]
                ixfe = struct.unpack_from("<I", b, 4)[0]
                miy = struct.unpack_from("<H", b, 8)[0]
                flags = b[11]
                if flags & 0x10:
                    sh.hidden_rows.add(row + 1)
                if flags & 0x20:
                    sh.row_heights[row + 1] = miy / 20
                if flags & 0x40:
                    sh.row_styles[row + 1] = ixfe
            elif rid in CELLS:
                if row < 0:
                    raise XlsbError(f"シート「{sh.name}」：行の見出しのないセルがあります")
                last_formula = None
                pos = p  # 値を書き換えるための、部品の中の位置
                if rid == C_BLANK:
                    self._put(sh, row, b, 4, None)
                    self.value_records[(sh.name, row + 1, struct.unpack_from("<I", b, 0)[0] + 1)] = (sh.path, rid, pos + 8, 0)
                    continue
                n_vals += 1
                col = struct.unpack_from("<I", b, 0)[0] + 1
                if rid == C_RK:
                    self._put(sh, row, b, 4, _rk(struct.unpack_from("<I", b, 8)[0]))
                    vl = 4
                elif rid == C_REAL:
                    self._put(sh, row, b, 4, struct.unpack_from("<d", b, 8)[0])
                    vl = 8
                elif rid == C_BOOL:
                    self._put(sh, row, b, 4, bool(b[8]))
                    vl = 1
                elif rid == C_ERROR:
                    self._put(sh, row, b, 4, ExcelError(ERRORS.get(b[8], "#N/A")))
                    vl = 1
                elif rid == C_ST:
                    s = wstr(b, 8)[0]
                    self._put(sh, row, b, 4, s)
                    vl = ln - 8
                elif rid == C_ISST:
                    i = struct.unpack_from("<I", b, 8)[0]
                    if i >= len(self.shared_strings):
                        raise XlsbError(f"シート「{sh.name}」：共有文字列の番号 {i} がありません")
                    self._put(sh, row, b, 4, self.shared_strings[i])
                    vl = 4
                elif rid == C_RSTR:
                    self._put(sh, row, b, 4, rich(b, 8))
                    self.rich_cells[(sh.name, row + 1, col)] = None
                    vl = ln - 8
                else:  # 数式：保存された結果＋flags 2 バイト＋ CellParsedFormula
                    if rid == F_ST:
                        v, q = wstr(b, 8)
                    elif rid == F_NUM:
                        v, q = struct.unpack_from("<d", b, 8)[0], 16
                    elif rid == F_BOOL:
                        v, q = bool(b[8]), 9
                    else:
                        v, q = ExcelError(ERRORS.get(b[8], "#N/A")), 9
                    vl = q - 8
                    cell = self._put(sh, row, b, 4, v)
                    if rid == F_ST:
                        cell.t = "str"
                    q += 2
                    cce = struct.unpack_from("<I", b, q)[0]
                    rgce = b[q + 4:q + 4 + cce]
                    q += 4 + cce
                    cb = struct.unpack_from("<I", b, q)[0] if q + 4 <= ln else 0
                    extra = b[q + 4:q + 4 + cb]
                    formulas.append((cell, rgce, extra))
                    last_formula = cell
                self.value_records[(sh.name, row + 1, col)] = (sh.path, rid, pos + 8, vl)
            elif rid == SHR:
                rng = _rfx(b)
                cce = struct.unpack_from("<I", b, 16)[0]
                cb_at = 20 + cce
                cb = struct.unpack_from("<I", b, cb_at)[0] if cb_at + 4 <= ln else 0
                if last_formula is not None:
                    shared[(last_formula.row, last_formula.col)] = (rng, b[20:20 + cce], b[cb_at + 4:cb_at + 4 + cb])
            elif rid == ARR:
                rng = _rfx(b)
                cce = struct.unpack_from("<I", b, 17)[0]
                cb_at = 21 + cce
                cb = struct.unpack_from("<I", b, cb_at)[0] if cb_at + 4 <= ln else 0
                if last_formula is not None:
                    arrays[(last_formula.row, last_formula.col)] = (rng, b[21:21 + cce], b[cb_at + 4:cb_at + 4 + cb])
            elif rid == COLINFO:
                c1, c2, w, x, fl = struct.unpack_from("<IIIIH", b, 0)
                for c in range(c1, min(c2, 16383) + 1):
                    if c - c1 > 4096:  # 列全体の宣言：既定の幅の扱いにする
                        break
                    sh.col_widths[c + 1] = w / 256
                    if x:
                        sh.col_styles[c + 1] = x
                    if fl & 1:
                        sh.hidden_cols.add(c + 1)
            elif rid == WSFMT:
                dx = struct.unpack_from("<I", b, 0)[0]
                _cc, dy = struct.unpack_from("<HH", b, 4)
                if dx != 0xFFFFFFFF:
                    sh.default_col_width = dx / 256
                sh.default_row_height = dy / 20 or sh.default_row_height
            elif rid == MERGE:
                r1, c1, r2, c2 = _rfx(b)
                sh.merges.append((r1, c1, r2, c2))
            elif rid == SHEETPROT:
                sh.protected = bool(struct.unpack_from("<I", b, 2)[0])
            elif rid == DVAL:
                self._dval(sh, b)
            elif rid == HLINK:
                r1, c1, _r2, _c2 = _rfx(b)
                relid, q = wstr(b, 16)
                loc, q = wstr(b, q)
                target = rels.get(relid, ("", ""))[1] if relid else ""
                if loc:
                    target = f"{target}#{loc}" if target else loc
                sh.hyperlinks.append((addr(r1, c1), target))
            elif rid == HEADERFOOTER:
                q = 2
                for _ in range(6):
                    try:
                        s, q = wstr(b, q)
                    except (struct.error, XlsbError):
                        break
                    t = _header_text(s) if s else ""
                    if t:
                        sh.header_footer.append(t)
            elif rid == BEGIN_CF:
                cf += 1
                try:
                    n = struct.unpack_from("<I", b, 8)[0]
                    cf_ranges = [_rfx(b, 12 + 16 * k) for k in range(n)]
                except struct.error:
                    cf_ranges = []
                    cf_bad += 1
            elif rid == BEGIN_CFRULE:
                try:
                    self._cf_rule(sh, b, cf_ranges)
                except (struct.error, XlsbError, IndexError, ValueError, Unsupported):
                    cf_bad += 1
            elif rid == DRAWING:
                self._drawing(sh, wstr(b, 0)[0], rels)
            elif rid == LISTPART:
                pass
        self.record_counts[sh.name] = n_vals
        if cf:
            sh.cond_formats = cf
            shown = "／".join(_cf_text(x) for x in sh.cond_rules[:8])
            self._note(sh, f"条件付き書式 {cf} か所：{shown or '条件を読めません'}。見た目の変化（色など）は読んでいないので、色で見え方が変わる所は窓に出ません"
                       + (f"（{cf_bad} 個の条件は読めません）" if cf_bad else ""))
        for cell, rgce, extra in formulas:
            self._formula(sh, cell, rgce, extra, shared, arrays)
        for typ, part in rels.values():
            if typ == "comments" and self.zip.has(part):
                self._comments(sh, part)

    def _cf_rule(self, sh: Sheet, b: bytes, ranges: list) -> None:
        """BrtBeginCFRule（MS-XLSB 2.4.23）：iType・iTemplate・dxfId・iPri・iParam・予約 2 つ・印 2 バイト・cbFmla 3 つ・strParam・数式（CFParsedFormula）。"""
        ty, tmpl, dxf, pri, param = struct.unpack_from("<IIIiI", b, 0)
        flags = struct.unpack_from("<H", b, 28)[0]
        sizes = struct.unpack_from("<III", b, 30)
        s, q = wstr(b, 42)
        forms = []
        origin = (min(r[0] for r in ranges) - 1, min(r[1] for r in ranges) - 1) if ranges else (0, 0)
        for sz in sizes:
            if not sz:
                continue
            cce = struct.unpack_from("<I", b, q)[0]
            rgce = bytes(b[q + 4:q + 4 + cce])
            q += 4 + cce
            cb = struct.unpack_from("<I", b, q)[0]
            extra = bytes(b[q + 4:q + 4 + cb])
            q += 4 + cb
            forms.append(self.decoder.decode(rgce, extra, base=origin))
        kind = _CF_TYPE.get(ty) or ("timePeriod" if 15 <= tmpl <= 24 else _CF_TEMPLATE.get(tmpl, f"type{ty}"))
        rule = CondRule(list(ranges), kind, pri, None if dxf == 0xFFFFFFFF else dxf, forms, _CF_OPER.get(param) if ty == 1 else None, bool(flags & 2))
        rule.text = s  # type: ignore[attr-defined]
        sh.cond_rules.append(rule)

    def xml(self, name: str) -> ET.Element:
        """DrawingML の部品（xlsx と同じ XML）。"""
        from ..xlsx.workbook import Workbook

        return Workbook.xml(self, name)

    def _drawing(self, sh: Sheet, relid: str, rels: dict) -> None:
        part = rels.get(relid, ("", ""))[1]
        if not part or not self.zip.has(part):
            self._note(sh, "図形・グラフの部品が見つかりません（Excel で PDF にしたものか、.xlsx で保存し直したものを読んでください）")
            return
        from ..xlsx.workbook import Workbook

        try:
            Workbook._load_drawing(self, sh, part)
        except (ET.ParseError, AttributeError, ValueError, TypeError) as e:
            self._note(sh, f"図形・グラフの部品を読めません（{type(e).__name__}）")

    def _note(self, sh: Sheet, text: str) -> None:
        s = f"シート「{sh.name}」：{text}"
        if s not in sh.unread:
            sh.unread.append(s)

    def _comments(self, sh: Sheet, part: str) -> None:
        data = self.zip.read(part)
        at = None
        for rid, p, ln in records(data):
            b = data[p:p + ln]
            if rid == C_BEGIN:
                at = _rfx(b, 4)[:2]
            elif rid == C_TEXT and at:
                sh.comments[at] = rich(b)

    def _formula(self, sh, cell, rgce, extra, shared, arrays) -> None:
        try:
            if rgce[:1] == b"\x01":  # PtgExp：行は rgce、列は rgcb の PtgExtraCol
                mr = struct.unpack_from("<I", rgce, 1)[0] + 1
                mc = struct.unpack_from("<I", extra, 0)[0] + 1
                if (mr, mc) in arrays:
                    (r1, c1, r2, c2), arce, aextra = arrays[(mr, mc)]
                    if (cell.row, cell.col) == (mr, mc):
                        text = self.decoder.decode(arce, aextra)
                        cell.formula = text
                        cell.formula_kind = "array"
                        cell.array_ref = f"{addr(r1, c1)}:{addr(r2, c2)}"
                        cell.ast = parse(text)
                    else:
                        cell.array_master = (mr, mc)
                    return
                if (mr, mc) in shared:
                    _rng, srce, sextra = shared[(mr, mc)]
                    key = (sh.name, mr, mc)
                    cache = self.__dict__.setdefault("_shared_cache", {})
                    if key not in cache:
                        text = self.decoder.decode(srce, sextra, base=(mr - 1, mc - 1))
                        cache[key] = (parse(text), text)
                    ast, text = cache[key]
                    cell.formula_kind = "shared"
                    cell.dr, cell.dc = cell.row - mr, cell.col - mc
                    if cell.dr == 0 and cell.dc == 0:
                        cell.formula = text
                    else:
                        cell._shared_from = (ast, cell.dr, cell.dc)
                    cell.ast = ast
                    return
                cell.parse_error = "共有数式の本体が見つかりません"
                cell.formula = ""
                return
            text = self.decoder.decode(rgce, extra)
            cell.formula = text
            cell.formula_kind = "normal"
            cell.ast = parse(text)
        except Unsupported as e:
            cell.formula = cell.formula or "?"
            cell.parse_error = f"未対応: {e}"
        except Exception as e:  # noqa: BLE001
            cell.formula = cell.formula or "?"
            cell.parse_error = f"構文エラー: {e}"

    def _dval(self, sh: Sheet, b: bytes) -> None:
        """BrtDVal（MS-XLSB 2.4.356）：flags 4・範囲（UncheckedSqRfX）・字 4 つ（DValStrings）・数式 2 つ（DVParsedFormula）。"""
        flags = struct.unpack_from("<I", b, 0)[0]
        n = struct.unpack_from("<I", b, 4)[0]
        ranges = [_rfx(b, 8 + 16 * k) for k in range(n)]
        q = 8 + 16 * n
        strs = []
        for _ in range(4):
            s, q = wstr(b, q)
            strs.append(s or None)
        fs = []
        for _ in range(2):
            cce = struct.unpack_from("<I", b, q)[0]
            rgce = b[q + 4:q + 4 + cce]
            q += 4 + cce
            cb = struct.unpack_from("<I", b, q)[0]
            fs.append((rgce, b[q + 4:q + 4 + cb]))
            q += 4 + cb
        if not ranges:
            return
        origin = (min(r[0] for r in ranges) - 1, min(r[1] for r in ranges) - 1)
        typ = _DV_TYPE.get(flags & 0x0F, "any")
        f1 = f2 = None
        try:
            if fs[0][0]:
                f1 = self.decoder.decode(*fs[0], base=origin)
                if typ == "list" and f1.startswith('"') and "\x00" in f1:
                    f1 = f1.replace("\x00", ",")
            if fs[1][0]:
                f2 = self.decoder.decode(*fs[1], base=origin)
        except Exception:  # noqa: BLE001
            f1 = f1 or "#REF!"
        sh.validations.append(Validation(
            sqref=ranges, type=typ, operator=_DV_OP.get((flags >> 20) & 0x0F), formula1=f1, formula2=f2,
            allow_blank=bool(flags & 0x100), prompt=strs[3], prompt_title=strs[2], error=strs[1],
            sqref_text=" ".join(f"{addr(r1, c1)}:{addr(r2, c2)}" if (r1, c1) != (r2, c2) else addr(r1, c1) for r1, c1, r2, c2 in ranges),
            error_title=strs[0], error_style={0: "stop", 1: "warning", 2: "information"}.get((flags >> 4) & 0x07, "stop"),
            show_error=bool(flags & 0x80000), show_input=bool(flags & 0x40000)))


