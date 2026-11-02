"""古いExcel（.xls、BIFF8）を読み、xlsx と同じ形（シート・セル・数式・入力規則・チェックボックスなど）にする。

読んだあとは、窓・計算・診断・検査を xlsx と同じ処理で行う。元のファイルには一切手を加えない。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from ..cfb import CFB
from ..formula.numfmt import BUILTIN
from ..formula.parser import parse
from ..formula.values import ExcelError, Unsupported
from ..xlsx.workbook import (Cell, CondRule, Control, DefinedName, Dxf, Sheet, Style, Validation, _area_ranges,
                             _header_text)
from .ptg import ERRORS, Decoder

# 標準のパレット（色番号 8〜63）
_PALETTE = [
    "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "800000", "008000", "000080",
    "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080",
    "0066CC", "CCCCFF", "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF",
    "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99", "3366FF", "33CCCC", "99CC00", "FFCC00",
    "FF9900", "FF6600", "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399",
    "333333",
]
_BUILTIN_NAMES = {0x00: "Consolidate_Area", 0x01: "Auto_Open", 0x02: "Auto_Close", 0x03: "Extract", 0x04: "Database",
                  0x05: "Criteria", 0x06: "Print_Area", 0x07: "Print_Titles", 0x08: "Recorder", 0x09: "Data_Form",
                  0x0A: "Auto_Activate", 0x0B: "Auto_Deactivate", 0x0C: "Sheet_Title", 0x0D: "_FilterDatabase"}
_DV_TYPE = {0: "any", 1: "whole", 2: "decimal", 3: "list", 4: "date", 5: "time", 6: "textLength", 7: "custom"}
_DV_OP = {0: "between", 1: "notBetween", 2: "equal", 3: "notEqual", 4: "greaterThan", 5: "lessThan",
          6: "greaterThanOrEqual", 7: "lessThanOrEqual"}
_CF_OP = {1: "between", 2: "notBetween", 3: "equal", 4: "notEqual", 5: "greaterThan", 6: "lessThan",
          7: "greaterThanOrEqual", 8: "lessThanOrEqual"}
_OBJ_KIND = {0x0B: "CheckBox", 0x0C: "Radio", 0x14: "Drop", 0x12: "List", 0x10: "Spin", 0x11: "Scroll",
             0x07: "Button", 0x0E: "Label", 0x13: "GBox", 0x08: "Pict", 0x06: "Text", 0x19: "Note"}


class XlsError(ValueError):
    pass


# ====================================================================== 低いところ
def records(data: bytes, start: int = 0):
    """(位置, 種類, 中身) を順に返す。"""
    pos = start
    n = len(data)
    while pos + 4 <= n:
        t, ln = struct.unpack_from("<HH", data, pos)
        yield pos, t, data[pos + 4:pos + 4 + ln]
        pos += 4 + ln


class Seg:
    """CONTINUE でつながった中身を、続けて読む（文字列の途中で切れたときの扱いも含む）。"""

    def __init__(self, parts: list[bytes]):
        self.parts = parts or [b""]
        self.pi = 0
        self.pos = 0

    def _avail(self) -> int:
        return len(self.parts[self.pi]) - self.pos

    def _next(self) -> None:
        if self.pi + 1 >= len(self.parts):
            raise XlsError("文字列のデータが途中で終わっています")
        self.pi += 1
        self.pos = 0

    def read(self, n: int) -> bytes:
        out = bytearray()
        while n > 0:
            if self._avail() == 0:
                self._next()
            take = min(n, self._avail())
            out += self.parts[self.pi][self.pos:self.pos + take]
            self.pos += take
            n -= take
        return bytes(out)

    def u8(self) -> int:
        return self.read(1)[0]

    def u16(self) -> int:
        return struct.unpack("<H", self.read(2))[0]

    def u32(self) -> int:
        return struct.unpack("<I", self.read(4))[0]

    def chars(self, cch: int, high: bool) -> str:
        out = []
        while cch > 0:
            if self._avail() == 0:
                self._next()
                high = bool(self.parts[self.pi][self.pos] & 1)  # 続きの先頭には、文字の幅の印が入る
                self.pos += 1
            width = 2 if high else 1
            take = min(cch, self._avail() // width)
            if take == 0:
                raise XlsError("文字列のデータが壊れています")
            raw = self.parts[self.pi][self.pos:self.pos + take * width]
            out.append(raw.decode("utf-16-le", errors="replace") if high else raw.decode("latin-1"))
            self.pos += take * width
            cch -= take
        return "".join(out)


def xlstring(b: bytes, pos: int, cch_size: int = 2) -> tuple[str, int]:
    """XLUnicodeString（文字数 u16 または u8、印、文字）。(文字列, 次の位置)。"""
    if cch_size == 2:
        cch = struct.unpack_from("<H", b, pos)[0]
    else:
        cch = b[pos]
    pos += cch_size
    if cch == 0 and pos >= len(b):
        return "", pos
    flags = b[pos]
    pos += 1
    runs = ext = 0
    if flags & 0x08:
        runs = struct.unpack_from("<H", b, pos)[0]
        pos += 2
    if flags & 0x04:
        ext = struct.unpack_from("<I", b, pos)[0]
        pos += 4
    if flags & 0x01:
        s = b[pos:pos + 2 * cch].decode("utf-16-le", errors="replace")
        pos += 2 * cch
    else:
        s = b[pos:pos + cch].decode("latin-1")
        pos += cch
    pos += 4 * runs + ext
    return s, pos


def _rk(v: int) -> float:
    if v & 0x02:
        x = float(v >> 2) if not v & 0x80000000 else float((v >> 2) - (1 << 30))
    else:
        x = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    return x / 100 if v & 0x01 else x


def _empty(s: str | None) -> str | None:
    # 入力規則の文字は、空のとき "\x00" 1文字で書かれる
    return None if s in (None, "", "\x00") else s


# ====================================================================== ブック
@dataclass
class _ZipShim:
    """xlsx の RawZip の代わり（元のバイト列を持つだけ）。"""

    original: bytes

    def build(self) -> bytes:
        return self.original


@dataclass
class _Supbook:
    kind: str  # self / addin / external
    path: str = ""
    sheets: list[str] = field(default_factory=list)
    names: list[str] = field(default_factory=list)
    ext_index: int = 0  # 外部ブックの番号（[1] など）


class XlsWorkbook:
    kind = "xls"

    def __init__(self, blob: bytes, filename: str = "", cfb=None):
        """cfb：ほかの文書に埋め込んだブック（cfb.SubStorage）を読むとき。blob は外側のファイル。"""
        self.filename = filename
        self.cfb = cfb if cfb is not None else CFB(blob)
        self.zip = _ZipShim(blob)
        # Excel 97 と 95 の両方の形を持つブック（Workbook と Book）は、Excel 97 以降が読む Workbook（BIFF8）を読む
        name = next((n for n in self.cfb.streams() if n.lower() == "workbook"), None) or \
            next((n for n in self.cfb.streams() if n.lower() == "book"), None)
        if name is None:
            raise XlsError("Excelのブックのストリームがありません")
        self.stream_name = name
        self.data = self.cfb.read(name)
        self.sheets: list[Sheet] = []
        self.names: list[DefinedName] = []
        self.shared_strings: list[str] = []
        self.phonetic: dict[int, str] = {}
        self.cell_sst: dict[tuple[str, int, int], int] = {}
        self.styles: list[Style] = []
        self.dxfs: list[Dxf] = []
        self.tables: dict = {}
        self.external: dict[int, dict[str, dict[tuple[int, int], object]]] = {}
        self.external_links: list[str] = []
        self.date1904 = False
        self.calc_full_on_load = False
        self.workbook_protected = False
        streams = {n.split("/")[0].lower() for n in self.cfb.streams()}
        self.has_macros = "_vba_project_cur" in streams or any("vba" in n.lower() for n in self.cfb.streams())
        self._vba = None
        self.wb_path = name
        self._fonts: list[int] = []  # 色番号
        self._formats: dict[int, str] = dict(BUILTIN)
        self._palette = list(_PALETTE)
        self._supbooks: list[_Supbook] = []
        self._xti: list[tuple[int, int, int]] = []
        self._name_raw: list[tuple[str, int, bytes, bytes, bool]] = []
        self.decoder = Decoder(self)
        self.biff = 8
        self.biff5_encoding = "cp1252"
        self.dual_format = bool({n.lower() for n in self.cfb.streams()} >= {"workbook", "book"})
        self._load()

    # ------------------------------------------------------------------ xlsx と同じ問い合わせ
    @property
    def vba(self):
        """マクロ（VBA）のプロジェクト。古い Excel では、複合ファイルの _VBA_PROJECT_CUR の中にある。"""
        if self._vba is None and self.has_macros:
            from ..vba import read_project

            try:
                self._vba = read_project(self.blob if hasattr(self, "blob") else self.zip.original)
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
        try:
            self.sheet(name)
            return True
        except KeyError:
            return False

    def style(self, idx: int) -> Style:
        if 0 <= idx < len(self.styles):
            return self.styles[idx]
        return Style("General", True, False)

    def defined_name(self, name: str, sheet: str | None) -> DefinedName | None:
        low = name.lower()
        if sheet:
            for d in self.names:
                if d.name.lower() == low and d.sheet_scope == sheet:
                    return d
        for d in self.names:
            if d.name.lower() == low and d.sheet_scope is None:
                return d
        return None

    # ------------------------------------------------------------------ 数式の読み戻しに使う
    def xti_prefix(self, ixti: int) -> str:
        from .ptg import _q

        if ixti >= len(self._xti):
            return "#REF!"
        isb, first, last = self._xti[ixti]
        sb = self._supbooks[isb] if isb < len(self._supbooks) else None
        if sb is None:
            return "#REF!"
        if sb.kind == "self":
            if first < 0 or first >= len(self.sheets):
                return "#REF!"
            if last != first:
                raise Unsupported("複数シートにまたがる参照")
            return _q(self.sheets[first].name) + "!"
        if sb.kind == "external":
            name = sb.sheets[first] if 0 <= first < len(sb.sheets) else ""
            return f"[{sb.ext_index}]{name}!" if name and "'" not in name and " " not in name else \
                "'" + f"[{sb.ext_index}]{name}".replace("'", "''") + "'!"
        raise Unsupported("外部の参照")

    def name_text(self, idx: int) -> str:
        if 1 <= idx <= len(self._name_raw):
            return self._name_raw[idx - 1][0]
        return "#NAME?"

    def extern_name(self, ixti: int, idx: int) -> str:
        if ixti < len(self._xti):
            sb = self._supbooks[self._xti[ixti][0]]
            if 1 <= idx <= len(sb.names):
                n = sb.names[idx - 1]
                return n if sb.kind == "addin" else f"[{sb.ext_index}]!{n}"
        return "#NAME?"

    # ------------------------------------------------------------------ 読み込み
    def _load(self) -> None:
        data = self.data
        it = list(records(data))
        if not it or it[0][1] != 0x0809:
            raise XlsError("Excelのブックとして読めません")
        vers = struct.unpack_from("<H", it[0][2], 0)[0]
        if vers == 0x0500:
            # Excel 5.0 / 95（BIFF5・BIFF7）：字はバイト列、XF・NAME の形が違う（biff5.py）
            for _pos, t_, _b in it[:16]:
                if t_ == 0x002F:
                    raise XlsError("パスワードで保護された Excel 5.0/95 のブックです。この形の暗号は、まだ解けません")
            from .biff5 import Biff5Reader

            self.biff = 5
            r5 = Biff5Reader(self, it)
            r5.load()
            self.biff5_encoding = r5.enc
            return
        if vers != 0x0600:
            raise XlsError(f"BIFF8 ではない形式（版 0x{vers:04X}）です")
        # FilePass（0x002F）がある：ブックの中身は暗号（パスワードで解いてから開く。解かずに読むと意味のない字になる）
        for _pos, t_, _b in it[:16]:
            if t_ == 0x002F:
                raise XlsError("パスワードで保護されたブックです。パスワードで解いてから開いてください")
            if t_ == 0x0085:
                break
        bound: list[tuple[int, int, int, str]] = []
        i = 0
        sst_parts: list[bytes] | None = None
        cur_sb: _Supbook | None = None
        while i < len(it):
            pos, t, b = it[i]
            if t == 0x000A:  # EOF：ブック全体の部分の終わり
                break
            if t == 0x0022:
                self.date1904 = bool(struct.unpack_from("<H", b, 0)[0])
            elif t == 0x0031:  # FONT
                self._fonts.append(struct.unpack_from("<H", b, 4)[0])
            elif t == 0x041E:  # FORMAT
                ifmt = struct.unpack_from("<H", b, 0)[0]
                self._formats[ifmt] = xlstring(b, 2)[0]
            elif t == 0x0092:  # PALETTE
                n = struct.unpack_from("<H", b, 0)[0]
                for k in range(n):
                    r, g, bl = b[2 + 4 * k], b[3 + 4 * k], b[4 + 4 * k]
                    if k < len(self._palette):
                        self._palette[k] = f"{r:02X}{g:02X}{bl:02X}"
            elif t == 0x00E0:  # XF
                self._xf(b)
            elif t == 0x0085:  # BOUNDSHEET
                lb = struct.unpack_from("<I", b, 0)[0]
                state, dt = b[4] & 0x03, b[5]
                name = xlstring(b, 6, cch_size=1)[0]
                bound.append((lb, state, dt, name))
            elif t == 0x00FC:  # SST
                sst_parts = [b]
                j = i + 1
                while j < len(it) and it[j][1] == 0x003C:
                    sst_parts.append(it[j][2])
                    j += 1
                i = j
                continue
            elif t == 0x01AE:  # SUPBOOK
                cur_sb = self._supbook(b)
                self._supbooks.append(cur_sb)
            elif t == 0x0023 and cur_sb is not None:  # EXTERNNAME
                pos2 = 6 if cur_sb.kind != "addin" else 6
                try:
                    cur_sb.names.append(xlstring(b, pos2, cch_size=1)[0])
                except Exception:  # noqa: BLE001
                    cur_sb.names.append("#NAME?")
            elif t == 0x0059 and cur_sb is not None and cur_sb.kind == "external":  # XCT と CRN：外部ブックの保存値
                itab = struct.unpack_from("<H", b, 2)[0]
                count = struct.unpack_from("<h", b, 0)[0]
                sheet = cur_sb.sheets[itab] if itab < len(cur_sb.sheets) else str(itab)
                cache = self.external.setdefault(cur_sb.ext_index, {}).setdefault(sheet.lower(), {})
                j = i + 1
                while j < len(it) and it[j][1] == 0x005A and (count < 0 or j - i <= abs(count) + 1):
                    self._crn(it[j][2], cache)
                    j += 1
                i = j
                continue
            elif t == 0x0017:  # EXTERNSHEET
                n = struct.unpack_from("<H", b, 0)[0]
                self._xti = [struct.unpack_from("<Hhh", b, 2 + 6 * k) for k in range(n)]
            elif t == 0x0018:  # NAME
                self._name(b)
            elif t == 0x0012:  # PROTECT（ブックの構成の保護）
                self.workbook_protected = bool(struct.unpack_from("<H", b, 0)[0])
            i += 1
        if sst_parts is not None:
            self._sst(sst_parts)
        for lb, state, dt, name in bound:
            st = {0: "visible", 1: "hidden", 2: "veryHidden"}.get(state, "visible")
            if dt != 0:
                kind = {1: "macrosheet", 2: "chartsheet", 6: "vbmodule"}.get(dt, f"type{dt}")
                self.sheets.append(Sheet(name=name, path=f"#{lb}", state=f"non-worksheet:{kind}"))
            else:
                self.sheets.append(Sheet(name=name, path=f"#{lb}", state=st))
        names = [s.name for s in self.sheets]
        for nm, itab, rgce, extra, hidden in self._name_raw:
            scope = names[itab - 1] if 0 < itab <= len(names) else None
            try:
                text = self.decoder.decode(rgce, extra) if rgce else ""
                ast = parse(text) if text else None
                err = None
            except Unsupported as e:
                text, ast, err = "", None, f"未対応: {e}"
            except Exception as e:  # noqa: BLE001
                text, ast, err = "", None, f"構文エラー: {e}"
            d = DefinedName(nm, text, scope, hidden, ast, err)
            self.names.append(d)
        for d in self.names:
            if d.name == "_xlnm.Print_Area" and d.sheet_scope:
                sh = self.sheet(d.sheet_scope)
                sh.print_area = d.text
                sh.print_ranges = _area_ranges(d.text)
        for (lb, _s, dt, _n), sh in zip(bound, self.sheets):
            if dt == 0:
                self._load_sheet(sh, lb)

    def _xf(self, b: bytes) -> None:
        ifnt, ifmt, prot = struct.unpack_from("<HHH", b, 0)
        locked = bool(prot & 0x01)
        hidden = bool(prot & 0x02)
        fls = (struct.unpack_from("<I", b, 14)[0] >> 26) & 0x3F
        icv_fore = struct.unpack_from("<H", b, 18)[0] & 0x7F
        fnt = ifnt - 1 if ifnt >= 4 else ifnt  # フォントの4番は欠番
        fc = self._color(self._fonts[fnt]) if 0 <= fnt < len(self._fonts) else None
        fill = self._color(icv_fore) if fls == 1 else None
        lines = struct.unpack_from("<I", b, 10)[0]
        bits = sum(k for k, sh in ((1, 0), (2, 4), (4, 8), (8, 12)) if (lines >> sh) & 0x0F)
        self.styles.append(Style(self._formats.get(ifmt, "General"), locked, hidden, fc, fill, bits))

    def _color(self, icv: int) -> str | None:
        if icv in (0x7FFF, 0x40, 0x51):
            return None  # 自動（ふつうは黒）
        if icv == 0x41:
            return "FFFFFF"
        if 8 <= icv < 8 + len(self._palette):
            return self._palette[icv - 8]
        if icv < 8:
            return ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF"][icv]
        return None

    def _supbook(self, b: bytes) -> _Supbook:
        ctab, cch = struct.unpack_from("<HH", b, 0)
        if cch == 0x0401:
            return _Supbook("self")
        if cch == 0x3A01:
            return _Supbook("addin")
        pos = 2
        path, pos = xlstring(b, pos)
        sheets = []
        for _ in range(ctab):
            s, pos = xlstring(b, pos)
            sheets.append(s)
        n = sum(1 for s in self._supbooks if s.kind == "external") + 1
        self.external_links.append(path)
        return _Supbook("external", path, sheets, ext_index=n)

    def _crn(self, b: bytes, cache: dict) -> None:
        col_last, col_first, row = b[0], b[1], struct.unpack_from("<H", b, 2)[0]
        pos = 4
        for c in range(col_first, col_last + 1):
            t = b[pos]
            pos += 1
            if t == 0x01:
                cache[(row + 1, c + 1)] = struct.unpack_from("<d", b, pos)[0]
                pos += 8
            elif t == 0x02:
                s, pos = xlstring(b, pos)
                cache[(row + 1, c + 1)] = s
            elif t == 0x04:
                cache[(row + 1, c + 1)] = bool(b[pos])
                pos += 8
            elif t == 0x10:
                cache[(row + 1, c + 1)] = ExcelError(ERRORS.get(b[pos], "#N/A"))
                pos += 8
            else:
                pos += 8

    def _name(self, b: bytes) -> None:
        grbit, _key, cch, cce = struct.unpack_from("<HBBH", b, 0)
        itab = struct.unpack_from("<H", b, 8)[0]
        pos = 14
        flags = b[pos]
        pos += 1
        if flags & 0x01:
            raw = b[pos:pos + 2 * cch].decode("utf-16-le", errors="replace")
            pos += 2 * cch
        else:
            raw = b[pos:pos + cch].decode("latin-1")
            pos += cch
        if grbit & 0x20 and raw:
            nm = "_xlnm." + _BUILTIN_NAMES.get(ord(raw[0]), f"builtin{ord(raw[0])}")
        else:
            nm = raw
        rgce = b[pos:pos + cce]
        extra = b[pos + cce:]
        self._name_raw.append((nm, itab, rgce, extra, bool(grbit & 0x01)))

    def _sst(self, parts: list[bytes]) -> None:
        r = Seg(parts)
        r.u32()
        unique = r.u32()
        for i in range(unique):
            cch = r.u16()
            flags = r.u8()
            runs = r.u16() if flags & 0x08 else 0
            ext = r.u32() if flags & 0x04 else 0
            s = r.chars(cch, bool(flags & 0x01))
            if runs:
                r.read(4 * runs)
            if ext:
                ph = _phonetic(s, r.read(ext))
                if ph is not None:
                    self.phonetic[i] = ph
            self.shared_strings.append(s)

    # ------------------------------------------------------------------ シート
    def _load_sheet(self, sh: Sheet, lb: int) -> None:
        data = self.data
        rows = []
        depth = 0
        for pos, t, b in records(data, lb):
            if t == 0x0809:
                depth += 1
                if depth > 1:
                    rows.append((pos, -1, b))  # 埋め込みグラフなどの中は読まない
                    continue
            elif t == 0x000A:
                depth -= 1
                if depth == 0:
                    break
                continue
            if depth > 1:
                continue
            rows.append((pos, t, b))
        shared: dict[tuple[int, int], tuple[tuple[int, int, int, int], bytes, bytes]] = {}
        arrays: dict[tuple[int, int], tuple[tuple[int, int, int, int], bytes, bytes]] = {}
        formulas = []
        last_formula: Cell | None = None
        draw_parts: list[bytes] = []
        objs: list[dict] = []
        notes: list[tuple[int, int, int]] = []
        txo_for: dict[int, str] = {}
        last_obj: dict | None = None
        dv_list: list[Validation] = []
        cf_blocks: list = []
        i = 0
        while i < len(rows):
            pos, t, b = rows[i]
            if t == 0x0203:  # NUMBER
                r, c, x = struct.unpack_from("<HHH", b, 0)
                self._put(sh, r, c, x, struct.unpack_from("<d", b, 6)[0])
            elif t == 0x027E:  # RK
                r, c, x, v = struct.unpack_from("<HHHI", b, 0)
                self._put(sh, r, c, x, _rk(v))
            elif t == 0x00BD:  # MULRK
                r, c0 = struct.unpack_from("<HH", b, 0)
                n = (len(b) - 6) // 6
                for k in range(n):
                    x, v = struct.unpack_from("<HI", b, 4 + 6 * k)
                    self._put(sh, r, c0 + k, x, _rk(v))
            elif t == 0x00FD:  # LABELSST
                r, c, x, isst = struct.unpack_from("<HHHI", b, 0)
                s = self.shared_strings[isst] if isst < len(self.shared_strings) else ""
                self._put(sh, r, c, x, s)
                if isst in self.phonetic:
                    self.cell_sst[(sh.name, r + 1, c + 1)] = isst
            elif t in (0x0204, 0x00D6):  # LABEL / RSTRING
                r, c, x = struct.unpack_from("<HHH", b, 0)
                self._put(sh, r, c, x, xlstring(b, 6)[0])
            elif t == 0x0201:  # BLANK
                r, c, x = struct.unpack_from("<HHH", b, 0)
                self._put(sh, r, c, x, None)
            elif t == 0x00BE:  # MULBLANK
                r, c0 = struct.unpack_from("<HH", b, 0)
                n = (len(b) - 6) // 2
                for k in range(n):
                    self._put(sh, r, c0 + k, struct.unpack_from("<H", b, 4 + 2 * k)[0], None)
            elif t == 0x0205:  # BOOLERR
                r, c, x, v, is_err = struct.unpack_from("<HHHBB", b, 0)
                self._put(sh, r, c, x, ExcelError(ERRORS.get(v, "#N/A")) if is_err else bool(v))
            elif t == 0x0006:  # FORMULA
                r, c, x = struct.unpack_from("<HHH", b, 0)
                res = b[6:14]
                cell = self._put(sh, r, c, x, None)
                if res[6:8] == b"\xff\xff":
                    kind = res[0]
                    cell.value = {1: bool(res[2]), 2: ExcelError(ERRORS.get(res[2], "#N/A")), 3: ""}.get(kind, "")
                    if kind == 0:
                        cell.t = "str"
                else:
                    cell.value = struct.unpack_from("<d", res, 0)[0]
                cce = struct.unpack_from("<H", b, 20)[0]
                formulas.append((cell, b[22:22 + cce], b[22 + cce:]))
                last_formula = cell
            elif t == 0x0207 and last_formula is not None:  # STRING：直前の数式の結果の文字列
                last_formula.value = xlstring(b, 0)[0]
            elif t == 0x04BC:  # SHRFMLA
                r1, r2, c1, c2 = struct.unpack_from("<HHBB", b, 0)
                cce = struct.unpack_from("<H", b, 8)[0]
                shared[(r1, c1)] = ((r1, c1, r2, c2), b[10:10 + cce], b[10 + cce:])
            elif t == 0x0221:  # ARRAY
                r1, r2, c1, c2 = struct.unpack_from("<HHBB", b, 0)
                cce = struct.unpack_from("<H", b, 12)[0]
                arrays[(r1, c1)] = ((r1, c1, r2, c2), b[14:14 + cce], b[14 + cce:])
            elif t == 0x0236:  # TABLE（データテーブル）
                pass
            elif t == 0x0208:  # ROW
                rw, _a, _b, miy = struct.unpack_from("<HHHH", b, 0)
                flags = struct.unpack_from("<I", b, 12)[0]
                if flags & 0x20:
                    sh.hidden_rows.add(rw + 1)
                if flags & 0x40 or True:
                    sh.row_heights[rw + 1] = (miy & 0x7FFF) / 20
                if flags & 0x80:
                    sh.row_styles[rw + 1] = (flags >> 16) & 0xFFF
            elif t == 0x0225:  # DEFAULTROWHEIGHT
                sh.default_row_height = struct.unpack_from("<H", b, 2)[0] / 20
            elif t == 0x007D:  # COLINFO
                c1, c2, w, x, flags = struct.unpack_from("<HHHHH", b, 0)
                for c in range(c1, min(c2, 255) + 1):
                    sh.col_widths[c + 1] = w / 256
                    if x not in (0, 15):
                        sh.col_styles[c + 1] = x
                    if flags & 0x01:
                        sh.hidden_cols.add(c + 1)
            elif t == 0x00E5:  # MERGEDCELLS
                n = struct.unpack_from("<H", b, 0)[0]
                for k in range(n):
                    r1, r2, c1, c2 = struct.unpack_from("<HHHH", b, 2 + 8 * k)
                    sh.merges.append((r1 + 1, c1 + 1, r2 + 1, c2 + 1))
            elif t == 0x0012:  # PROTECT
                sh.protected = bool(struct.unpack_from("<H", b, 0)[0])
            elif t in (0x0014, 0x0015) and len(b) >= 3:  # HEADER / FOOTER
                txt = _header_text(xlstring(b, 0)[0])
                if txt:
                    sh.header_footer.append(txt)
            elif t == 0x01BE:  # DV
                v = self._dv(b)
                if v is not None:
                    dv_list.append(v)
            elif t == 0x01B0:  # CONDFMT
                cf_blocks.append(self._condfmt(b))
            elif t == 0x01B1 and cf_blocks:  # CF
                self._cf(b, cf_blocks[-1], sh)
            elif t == 0x00EC:  # MSODRAWING
                draw_parts.append(b)
                j = i + 1
                while j < len(rows) and rows[j][1] == 0x003C and rows[j - 1][1] == 0x00EC:
                    draw_parts.append(rows[j][2])
                    j += 1
            elif t == 0x005D:  # OBJ
                last_obj = self._obj(_obj_data(rows, i), sh)
                objs.append(last_obj)
            elif t == 0x01B6:  # TXO：直前の図形（OBJ）の文字
                cch = struct.unpack_from("<H", b, 10)[0]
                text = ""
                if cch and i + 1 < len(rows) and rows[i + 1][1] == 0x003C:
                    cont = rows[i + 1][2]
                    parts = [cont]
                    k = i + 2
                    seg = Seg(parts)
                    high = bool(seg.u8() & 0x01)
                    try:
                        text = seg.chars(cch, high)
                    except XlsError:
                        # 文字が次の CONTINUE にまたがる
                        more = [cont]
                        while k < len(rows) and rows[k][1] == 0x003C:
                            more.append(rows[k][2])
                            k += 1
                        s2 = Seg(more)
                        high = bool(s2.u8() & 0x01)
                        text = s2.chars(cch, high)
                if last_obj is not None:
                    last_obj["text"] = text
                    txo_for[last_obj["id"]] = text
            elif t == 0x001C:  # NOTE
                r, c, _g, idobj = struct.unpack_from("<HHHH", b, 0)
                notes.append((r, c, idobj))
            i += 1
        sh.validations = dv_list
        # 数式
        for cell, rgce, extra in formulas:
            self._formula(sh, cell, rgce, extra, shared, arrays)
        # メモ
        for r, c, idobj in notes:
            if idobj in txo_for:
                sh.comments[(r + 1, c + 1)] = txo_for[idobj]
        # 図形（位置と代替テキスト）とフォームコントロール
        anchors = _drawing_anchors(b"".join(draw_parts))
        k = 0
        for ob in objs:
            anchor = anchors[k] if k < len(anchors) else None
            k += 1
            self._place(sh, ob, anchor)

    def _put(self, sh: Sheet, r: int, c: int, xf: int, value) -> Cell:
        cell = sh.cells.get((r + 1, c + 1))
        if cell is None:
            cell = Cell(sh.name, r + 1, c + 1, style=xf)
            sh.cells[(r + 1, c + 1)] = cell
        cell.style = xf
        cell.value = value
        if isinstance(value, str):
            cell.t = "s"
        elif isinstance(value, bool):
            cell.t = "b"
        sh.max_row = max(sh.max_row, r + 1)
        sh.max_col = max(sh.max_col, c + 1)
        return cell

    def _formula(self, sh, cell, rgce, extra, shared, arrays) -> None:
        try:
            if rgce[:1] == b"\x01":  # ptgExp：共有数式か配列数式の本体を指す
                mr, mc = struct.unpack_from("<HH", rgce, 1)
                if (mr, mc) in arrays and arrays[(mr, mc)][0][0] <= cell.row - 1 <= arrays[(mr, mc)][0][2]:
                    (r1, c1, r2, c2), arce, aextra = arrays[(mr, mc)]
                    if (cell.row - 1, cell.col - 1) == (mr, mc):
                        text = self.decoder.decode(arce, aextra)
                        cell.formula = text
                        cell.formula_kind = "array"
                        cell.array_ref = f"{_a1(r1, c1)}:{_a1(r2, c2)}"
                        cell.ast = parse(text)
                    else:
                        cell.array_master = (mr + 1, mc + 1)
                    return
                if (mr, mc) in shared:
                    (r1, c1, r2, c2), srce, sextra = shared[(mr, mc)]
                    key = ("__shared__", mr, mc)
                    master = self._shared_cache.get(key) if hasattr(self, "_shared_cache") else None
                    if master is None:
                        text = self.decoder.decode(srce, sextra, base=(mr, mc), rel=True)
                        master = (parse(text), text)
                        if not hasattr(self, "_shared_cache"):
                            self._shared_cache = {}
                        self._shared_cache[key] = master
                    ast, text = master
                    cell.formula_kind = "shared"
                    cell.dr, cell.dc = cell.row - 1 - mr, cell.col - 1 - mc
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

    def _dv(self, b: bytes) -> Validation | None:
        flags = struct.unpack_from("<I", b, 0)[0]
        pos = 4
        strs = []
        for _ in range(4):
            s, pos = xlstring(b, pos)
            strs.append(_empty(s))
        f = []
        for _ in range(2):
            cce = struct.unpack_from("<H", b, pos)[0]
            pos += 4
            f.append(b[pos:pos + cce])
            pos += cce
        n = struct.unpack_from("<H", b, pos)[0]
        pos += 2
        ranges = []
        for _ in range(n):
            r1, r2, c1, c2 = struct.unpack_from("<HHHH", b, pos)
            pos += 8
            ranges.append((r1 + 1, c1 + 1, r2 + 1, c2 + 1))
        if not ranges:
            return None
        origin = (min(r[0] for r in ranges) - 1, min(r[1] for r in ranges) - 1)
        typ = _DV_TYPE.get(flags & 0x0F, "any")
        f1 = f2 = None
        try:
            if f[0]:
                if flags & 0x80 and f[0][:1] == b"\x17":  # 一覧を直接書いた選択肢（"a\0b\0c"）
                    items = self.decoder.decode(f[0])[1:-1].replace('""', '"').split("\x00")
                    f1 = '"' + ",".join(items) + '"'
                else:
                    f1 = self.decoder.decode(f[0], base=origin, rel=True)
            if f[1]:
                f2 = self.decoder.decode(f[1], base=origin, rel=True)
        except Exception:  # noqa: BLE001
            f1 = f1 or "#REF!"
        from ..formula.parser import addr

        v = Validation(sqref=ranges, type=typ, operator=_DV_OP.get((flags >> 20) & 0x0F),
                       formula1=f1, formula2=f2, allow_blank=bool(flags & 0x100), prompt=strs[2],
                       prompt_title=strs[0], error=strs[3],
                       sqref_text=" ".join(f"{addr(r1, c1)}:{addr(r2, c2)}" if (r1, c1) != (r2, c2) else addr(r1, c1)
                                           for r1, c1, r2, c2 in ranges),
                       error_title=strs[1], error_style={0: "stop", 1: "warning", 2: "information"}.get((flags >> 4) & 0x07, "stop"),
                       show_error=bool(flags & 0x80000), show_input=bool(flags & 0x40000))
        return v

    def _condfmt(self, b: bytes):
        n = struct.unpack_from("<H", b, 12)[0]
        ranges = []
        for k in range(n):
            r1, r2, c1, c2 = struct.unpack_from("<HHHH", b, 14 + 8 * k)
            ranges.append((r1 + 1, c1 + 1, r2 + 1, c2 + 1))
        return {"ranges": ranges, "count": 0}

    def _cf(self, b: bytes, block: dict, sh: Sheet) -> None:
        ct, cp, cce1, cce2 = struct.unpack_from("<BBHH", b, 0)
        flags = struct.unpack_from("<I", b, 6)[0]
        flags2 = struct.unpack_from("<H", b, 10)[0]
        pos = 12
        font = fill = fmt = None
        if flags & (1 << 25):  # 表示形式
            if flags2 & 0x01:
                cb = struct.unpack_from("<H", b, pos)[0]
                fmt = xlstring(b, pos + 2)[0]
                pos += cb
            else:
                pos += 2
        if flags & (1 << 26):  # フォント
            icv = struct.unpack_from("<i", b, pos + 80)[0]
            font = self._color(icv) if icv >= 0 else None
            pos += 118
        if flags & (1 << 27):
            pos += 8
        if flags & (1 << 28):
            pos += 8
        if flags & (1 << 29):  # 塗り
            fls = (struct.unpack_from("<H", b, pos)[0] >> 10) & 0x3F
            icvs = struct.unpack_from("<H", b, pos + 2)[0]
            fore, back = icvs & 0x7F, (icvs >> 7) & 0x7F
            fill = self._color(back if fls in (0, 1) else fore) if fls else None
            pos += 4
        if flags & (1 << 30):
            pos += 2
        rgce1, rgce2 = b[pos:pos + cce1], b[pos + cce1:pos + cce1 + cce2]
        origin = (min(r[0] for r in block["ranges"]) - 1, min(r[1] for r in block["ranges"]) - 1) if block["ranges"] else (0, 0)
        try:
            fs = [self.decoder.decode(x, base=origin, rel=True) for x in (rgce1, rgce2) if x]
        except Exception:  # noqa: BLE001
            return
        self.dxfs.append(Dxf(font, fill, fmt))
        block["count"] += 1
        sh.cond_formats += 1
        sh.cond_rules.append(CondRule(block["ranges"], "expression" if ct == 2 else "cellIs",
                                      len(sh.cond_rules) + 1, len(self.dxfs) - 1, fs, _CF_OP.get(cp)))

    def _obj(self, b: bytes, sh: Sheet) -> dict:
        """OBJ（b は続く CONTINUE もつないだ中身）を読む。部品の構造は obj.py（MS-XLS 2.4.181 Obj）のとおり。"""
        from .obj import parse as parse_obj

        info = parse_obj(b)
        ob = {"ot": info.ot, "id": info.id, "checked": info.checked, "link": None, "text": "", "fmla_range": None,
              "info": info}
        for key, rgce in (("link", info.link), ("fmla_range", info.fmla)):
            if rgce:
                try:
                    ob[key] = self.decoder.decode(rgce)  # つながったセル・選択肢の範囲
                except Exception:  # noqa: BLE001
                    ob[key] = "#REF!"
        return ob

    def _place(self, sh: Sheet, ob: dict, anchor) -> None:
        from ..xlsx.workbook import Shape, _link_cell

        kind = _OBJ_KIND.get(ob["ot"], "")
        info = ob.get("info")
        if kind in ("Drop", "List") and info is not None and info.ui_object:
            return  # オートフィルターなどのために Excel が自動で置いたボタン（入力欄ではない）
        pos = (anchor["r1"], anchor["c1"]) if anchor else None
        pos_to = (anchor["r2"], anchor["c2"]) if anchor else None
        if kind in ("CheckBox", "Radio", "Drop", "List", "Spin", "Scroll", "Button", "Label", "GBox"):
            c = Control(kind=kind, name=(anchor or {}).get("name") or f"{kind} {ob['id']}", shape_id=str(ob["id"]),
                        anchor=pos, anchor_to=pos_to, link=ob["link"], text=ob["text"] or "",
                        from_off=(anchor or {}).get("off1", 0.0), to_off=(anchor or {}).get("off2", 0.0))
            if kind in ("CheckBox", "Radio"):
                c.checked = {1: "Checked", 2: "Mixed"}.get(ob["checked"] or 0)
            if info is not None:
                _control_state(c, info, ob["fmla_range"])
            if c.link:
                c.link_cell = _link_cell(c.link.replace("$", ""), sh.name)
            sh.controls.append(c)
        elif kind == "Note":
            return
        else:
            alt = (anchor or {}).get("alt", "")
            sk = {"Pict": "picture", "Text": "text"}.get(kind, "shape")
            if ob["text"] or alt or sk == "picture":
                sh.shapes.append(Shape(sk if ob["text"] or sk == "picture" else "shape",
                                       (anchor or {}).get("name", ""), ob["text"], alt, pos))


def _a1(r: int, c: int) -> str:
    from ..formula.parser import addr

    return addr(r + 1, c + 1)


def _obj_data(rows: list, i: int) -> bytes:
    """OBJ の中身に、すぐ後ろに続く CONTINUE の中身をつないだもの（長いリストボックスの項目などは CONTINUE にまたがる）。"""
    parts = [rows[i][2]]
    k = i + 1
    while k < len(rows) and rows[k][1] == 0x003C:
        parts.append(rows[k][2])
        k += 1
    return b"".join(parts)


def _control_state(c: Control, info, fmla_range: str | None) -> None:
    """OBJ から読んだ部品の状態を、xlsx（CT_FormControlPr）と同じ欄に入れる。"""
    from .obj import SEL_TYPES

    if c.kind == "Radio":
        c.first_button = info.first_btn
        c.next_button = str(info.rad_next or 0)  # 同じ組の次のボタン（0 は次がない）
    if c.kind in ("Spin", "Scroll", "Drop", "List") and info.sbs is not None:
        c.val = info.sbs[0]  # リストボックス・ドロップダウンでは、上に隠れている行の数（xlsx の val と同じ）
        if c.kind in ("Spin", "Scroll"):
            c.vmin, c.vmax, c.inc, c.page = info.sbs[1:5]
    if c.kind in ("Drop", "List"):
        c.sel = info.sel
        c.fmla_range = fmla_range
        if not fmla_range and info.items is not None:
            c.items = list(info.items)  # 選択肢の文字がファイルの中にある（rgLines）
        if c.kind == "List":
            c.sel_type = SEL_TYPES.get(info.sel_type, "single")
        if info.bsels is not None:
            c.multi_sel = [k + 1 for k, x in enumerate(info.bsels) if x]


def _phonetic(base: str, ext: bytes) -> str | None:
    """ExtRst（ふりがなの情報）から、ふりがなを当てた読みを作る。"""
    try:
        cb = struct.unpack_from("<H", ext, 2)[0]
        p = 4 + 4  # Phs（ifnt, ph）
        crun, _cch = struct.unpack_from("<HH", ext, p)
        cch_st = struct.unpack_from("<H", ext, p + 4)[0]
        st = ext[p + 6:p + 6 + 2 * cch_st].decode("utf-16-le", errors="replace")
        q = p + 6 + 2 * cch_st
        runs = []
        for k in range(crun):
            ich_first, ich_mom, cch_mom = struct.unpack_from("<HHH", ext, q + 6 * k)
            runs.append((ich_first, ich_mom, cch_mom))
        if not st or not runs:
            return None
        out, at = [], 0
        for k, (f, m, n) in enumerate(runs):
            nxt = runs[k + 1][0] if k + 1 < len(runs) else len(st)
            out.append(base[at:m])
            out.append(st[f:nxt])
            at = m + n
        out.append(base[at:])
        del cb
        return "".join(out)
    except Exception:  # noqa: BLE001
        return None


def _drawing_anchors(data: bytes) -> list[dict]:
    """MSODRAWING（OfficeArt）から、図形ごとの位置（左上・右下のセル）・名前・代替テキストを、OBJ と同じ順に取り出す。"""
    out: list[dict] = []

    def walk(pos: int, end: int, cur: dict | None):
        while pos + 8 <= end:
            verinst, typ, ln = struct.unpack_from("<HHI", data, pos)
            ver = verinst & 0x0F
            body = pos + 8
            if typ == 0xF004:  # 図形のまとまり
                shape: dict = {}
                walk(body, min(body + ln, end), shape)
                if shape.get("client"):
                    out.append(shape)
            elif ver == 0x0F:
                walk(body, min(body + ln, end), cur)
            elif typ == 0xF010 and cur is not None and ln >= 18:  # 位置
                _f, c1, dx1, r1, dy1, c2, dx2, r2, dy2 = struct.unpack_from("<HHHHHHHHH", data, body)
                cur.update(r1=r1 + 1, c1=c1 + 1, r2=r2 + 1, c2=c2 + 1, off1=dy1 / 256 * 15, off2=dy2 / 256 * 15)
            elif typ == 0xF011 and cur is not None:  # この図形には OBJ がある
                cur["client"] = True
            elif typ == 0xF00B and cur is not None:  # 図形の性質（名前・代替テキスト）
                n = verinst >> 4
                props = []
                p = body
                for _ in range(n):
                    if p + 6 > body + ln:
                        break
                    opid, op = struct.unpack_from("<HI", data, p)
                    props.append((opid & 0x3FFF, bool(opid & 0x8000), op))
                    p += 6
                for pid, complex_, op in props:
                    if complex_:
                        raw = data[p:p + op]
                        p += op
                        if pid in (0x0380, 0x0381):
                            txt = raw.decode("utf-16-le", errors="replace").rstrip("\x00")
                            cur["name" if pid == 0x0380 else "alt"] = txt
            pos = body + ln

    walk(0, len(data), None)
    return out
