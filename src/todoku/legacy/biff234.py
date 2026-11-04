"""Excel 2.x・3.0・4.0（BIFF2・BIFF3・BIFF4 の 1 シートの .xls。複合ファイルでなく、レコードの並びそのもの）を読む（読むだけ）。

レコードの形は OpenOffice.org の「Microsoft Excel File Format」のとおり：
- BOF は 0x0009・0x0209・0x0409。セルのレコードは BIFF2 では 3 バイトの属性（2.5.13）、BIFF3・4 では XF の番号（BIFF8 と同じ形）
- 字はバイト列で、CODEPAGE の文字コード（なければ、読み方を推定して窓に書く）。LABEL の長さは BIFF2 が 1 バイト、BIFF3・4 が 2 バイト
- 数式は 3.4.7 の表の大きさのトークン（tFunc の関数の番号は BIFF2・3 が 1 バイト、BIFF4 が 2 バイト）。
  読めないトークンは Unsupported（窓には保存された結果を出す）
"""

from __future__ import annotations

import struct

from ..formula.parser import num_to_col, parse
from ..formula.values import ExcelError, Unsupported
from ..xls.book import _ZipShim, _rk
from ..xls.ptg import _BINOP, _PREC, ERRORS, FTAB
from ..xlsb.book import XlsbWorkbook
from ..xlsx.workbook import Cell, DefinedName, Sheet, Style, _header_text

BOFS = {0x0009: 2, 0x0209: 3, 0x0409: 4}
SHEET_TYPES = {0x10, 0x20, 0x40, 0x100}


class BiffError(ValueError):
    pass


def is_biff234(b: bytes) -> bool:
    if len(b) < 8:
        return False
    t, ln = struct.unpack_from("<HH", b, 0)
    if t not in BOFS or ln < 4 or ln > 16:
        return False
    return struct.unpack_from("<H", b, 6)[0] in SHEET_TYPES


def _records(d: bytes):
    pos = 0
    while pos + 4 <= len(d):
        t, ln = struct.unpack_from("<HH", d, pos)
        yield t, d[pos + 4:pos + 4 + ln]
        pos += 4 + ln


def _ref(raw: int, col: int, base=None) -> str:
    row = raw & 0x3FFF
    rel_r, rel_c = bool(raw & 0x8000), bool(raw & 0x4000)
    return f"{'' if rel_c else '$'}{num_to_col(col + 1)}{'' if rel_r else '$'}{row + 1}"


class Decoder234:
    def __init__(self, biff: int, enc, name_text):
        self.biff, self.enc, self.name_text = biff, enc, name_text

    def decode(self, g: bytes) -> str:
        try:
            return self._decode(g)
        except (struct.error, IndexError) as e:
            raise Unsupported(f"数式のトークンが壊れています（{e}）") from None

    def _decode(self, g: bytes) -> str:
        biff = self.biff
        st: list[tuple[str, int]] = []
        i, n = 0, len(g)
        while i < n:
            ptg = g[i]
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
            elif ptg == 0x12:
                a, _ = st.pop()
                st.append(("+" + a, _PREC["neg"]))
            elif ptg == 0x13:
                a, pa = st.pop()
                st.append(("-" + (a if pa >= _PREC["neg"] else f"({a})"), _PREC["neg"]))
            elif ptg == 0x14:
                a, pa = st.pop()
                st.append(((a if pa >= _PREC["%"] else f"({a})") + "%", _PREC["%"]))
            elif ptg == 0x15:
                a, _ = st.pop()
                st.append((f"({a})", 99))
            elif ptg == 0x16:
                st.append(("", 99))
            elif ptg == 0x17:
                cch = g[i]
                s = self.enc(g[i + 1:i + 1 + cch])
                i += 1 + cch
                st.append(('"' + s.replace('"', '""') + '"', 99))
            elif ptg == 0x19:
                grbit = g[i]
                if grbit & 0x04:
                    raise Unsupported("CHOOSE の飛び先の表")
                i += 2 if biff == 2 else 3
                if grbit & 0x10:
                    a, _ = st.pop()
                    st.append((f"SUM({a})", 99))
            elif ptg == 0x1C:
                st.append((ERRORS.get(g[i], "#N/A"), 99))
                i += 1
            elif ptg == 0x1D:
                st.append(("TRUE" if g[i] else "FALSE", 99))
                i += 1
            elif ptg == 0x1E:
                st.append((str(struct.unpack_from("<H", g, i)[0]), 99))
                i += 2
            elif ptg == 0x1F:
                v = struct.unpack_from("<d", g, i)[0]
                st.append((repr(v) if v != int(v) or abs(v) >= 1e15 else str(int(v)), 99))
                i += 8
            elif bid == 0x21:
                if biff == 4:
                    iftab = struct.unpack_from("<H", g, i)[0]
                    i += 2
                else:
                    iftab = g[i]
                    i += 1
                name, argc = FTAB.get(iftab, (None, None))
                if name is None or argc is None:
                    raise Unsupported(f"関数の番号 {iftab}")
                args = [st.pop()[0] for _ in range(argc)][::-1]
                st.append((f"{name}({','.join(args)})", 99))
            elif bid == 0x22:
                argc = g[i] & 0x7F
                if biff == 4:
                    iftab = struct.unpack_from("<H", g, i + 1)[0] & 0x7FFF
                    i += 3
                else:
                    iftab = g[i + 1]
                    i += 2
                args = [st.pop()[0] for _ in range(argc)][::-1]
                name = FTAB.get(iftab, (None, None))[0]
                if name is None:
                    raise Unsupported(f"関数の番号 {iftab}")
                st.append((f"{name}({','.join(args)})", 99))
            elif bid == 0x23:
                idx = struct.unpack_from("<H", g, i)[0]
                i += 7 if biff == 2 else 10
                st.append((self.name_text(idx), 99))
            elif bid == 0x24:
                raw, col = struct.unpack_from("<HB", g, i)
                i += 3
                st.append((_ref(raw, col), 99))
            elif bid == 0x25:
                r1, r2, c1, c2 = struct.unpack_from("<HHBB", g, i)
                i += 6
                st.append((f"{_ref(r1, c1)}:{_ref(r2, c2)}", 99))
            elif bid == 0x2A:
                i += 3
                st.append(("#REF!", 99))
            elif bid == 0x2B:
                i += 6
                st.append(("#REF!", 99))
            else:
                raise Unsupported(f"トークン 0x{ptg:02X}")
        if len(st) != 1:
            raise Unsupported("数式のトークンが壊れています")
        return st[0][0]


class Biff234Workbook(XlsbWorkbook):
    kind = "xls2"

    def __init__(self, blob: bytes, filename: str = ""):
        self.filename = filename
        self.zip = _ZipShim(blob)
        self.sheets: list[Sheet] = []
        self.names: list[DefinedName] = []
        self.shared_strings: list[str] = []
        self.phonetic = {}
        self.cell_sst = {}
        self.rich_cells = {}
        self.styles: list[Style] = []
        self.dxfs = []
        self.tables = {}
        self.external = {}
        self.external_links = []
        self.date1904 = False
        self.calc_full_on_load = False
        self.workbook_protected = False
        self.has_macros = False
        self._vba = None
        self.wb_path = ""
        self.value_records = {}
        self.record_counts = {}
        self._name_raw = []
        self.codepage: str | None = None
        self.guessed_codepage = False
        self.biff = 0
        self._load2(blob)

    @property
    def vba(self):
        return None

    def name_text(self, idx: int) -> str:
        return self._names_only[idx - 1] if 1 <= idx <= len(self._names_only) else "#NAME?"

    def _text(self, raw: bytes) -> str:
        if self.codepage:
            return raw.decode(self.codepage, "replace")
        if raw.isascii():
            return raw.decode("ascii")
        self.guessed_codepage = True
        try:
            return raw.decode("cp932")
        except UnicodeDecodeError:
            return raw.decode("cp1252", "replace")

    def _load2(self, blob: bytes) -> None:
        if not is_biff234(blob):
            raise BiffError("Excel 2〜4 のブックとして読めません")
        recs = list(_records(blob))
        self.biff = BOFS[recs[0][0]]
        if struct.unpack_from("<H", recs[0][1], 2)[0] == 0x100:
            sheet_names = [self._text(b[7:7 + b[6]]) for t, b in recs if t == 0x008E and len(b) > 6]
        else:
            sheet_names = []
        for t, b in recs:
            if t == 0x0042 and len(b) >= 2:
                cp = struct.unpack_from("<H", b, 0)[0]
                self.codepage = {932: "cp932", 1252: "cp1252", 367: "ascii", 437: "cp437", 0x8000: "mac_roman", 0x8001: "cp1252", 949: "cp949", 936: "gbk", 950: "cp950"}.get(cp)
                break
            if t == 0x002F:
                raise BiffError("パスワードで保護された古い Excel のブックです。この形の暗号は解けません。解いて保存し直したものか、PDF を頼んでください")
        self._fmts: list[str] = []
        self._fonts: list[int] = []
        stem = self.filename.rsplit("/", 1)[-1].rsplit(".", 1)[0] or "Sheet1"
        i = 0
        sh = None
        n_sheet = 0
        formulas: list = []
        names_raw: list = []
        ixfe = None
        last = None
        attr_styles: dict[bytes, int] = {}
        counts: dict[str, int] = {}
        for t, b in recs:
            if t in BOFS:
                dt = struct.unpack_from("<H", b, 2)[0]
                if dt == 0x10 or dt == 0x40:
                    n_sheet += 1
                    nm = sheet_names[n_sheet - 1] if n_sheet <= len(sheet_names) else (stem if n_sheet == 1 else f"{stem}{n_sheet}")
                    sh = Sheet(name=nm, path=f"#{n_sheet}")
                    self.sheets.append(sh)
                    counts[nm] = 0
                else:
                    sh = None
                continue
            try:
                if t == 0x001E:  # FORMAT
                    s = b[2:] if self.biff == 4 else b
                    self._fmts.append(self._text(s[1:1 + s[0]]) if s else "General")
                elif t in (0x0031, 0x0231):
                    self._fonts.append(0)
                elif t == 0x0043 and self.biff == 2:
                    pass  # XF：BIFF2 では、セルの属性に同じ内容がある
                elif t in (0x0243, 0x0443):
                    self._xf34(b, t)
                elif t in (0x0018, 0x0218):
                    if self.biff == 2:
                        ln, sz = b[3], b[4]
                        names_raw.append((self._text(b[5:5 + ln]), b[5 + ln:5 + ln + sz]))
                    else:
                        ln = b[3]
                        sz = struct.unpack_from("<H", b, 4)[0]
                        names_raw.append((self._text(b[6:6 + ln]), b[6 + ln:6 + ln + sz]))
                if sh is None:
                    continue
                if t == 0x0044:  # IXFE
                    ixfe = struct.unpack_from("<H", b, 0)[0]
                elif t == 0x0012:
                    sh.protected = bool(struct.unpack_from("<H", b, 0)[0])
                elif t in (0x0014, 0x0015) and b:
                    txt = _header_text(self._text(b[1:1 + b[0]]))
                    if txt:
                        sh.header_footer.append(txt)
                elif t == 0x0024 and len(b) >= 4:
                    w = struct.unpack_from("<H", b, 2)[0] / 256
                    for c in range(b[0], b[1] + 1):
                        sh.col_widths[c + 1] = w
                elif t == 0x007D:
                    c1, c2, w, x, fl = struct.unpack_from("<HHHHH", b, 0)
                    for c in range(c1, min(c2, 255) + 1):
                        sh.col_widths[c + 1] = w / 256
                        if fl & 1:
                            sh.hidden_cols.add(c + 1)
                elif t == 0x0208 and len(b) >= 16:
                    rw = struct.unpack_from("<H", b, 0)[0]
                    if struct.unpack_from("<I", b, 12)[0] & 0x20:
                        sh.hidden_rows.add(rw + 1)
                elif t == 0x001C:  # NOTE
                    r, c, ln = struct.unpack_from("<HHH", b, 0)
                    sh.comments[(r + 1, c + 1)] = self._text(b[6:6 + ln])
                elif t == 0x0207 or t == 0x0007:  # STRING：直前の数式の字の結果
                    if last is not None:
                        s = (b[2:2 + struct.unpack_from("<H", b, 0)[0]] if t == 0x0207 else b[1:1 + b[0]])
                        last.value = self._text(s)
                        last.t = "str"
                else:
                    cell, g = self._cell(sh, t, b, attr_styles, ixfe)
                    if cell is not None:
                        ixfe = None
                        if cell.value is not None or g is not None:
                            counts[sh.name] += 1
                        if g is not None:
                            formulas.append((cell, g))
                            last = cell
                        else:
                            last = None
            except (struct.error, IndexError) as e:
                raise BiffError(f"レコード 0x{t:04X} が壊れています（{e}）") from None
            i += 1
        self.record_counts = counts
        self._names_only = [n for n, _g in names_raw]
        dec = Decoder234(self.biff, self._text, self.name_text)
        for nm, g in names_raw:
            try:
                text = dec.decode(g) if g else ""
                self.names.append(DefinedName(nm, text, None, False, parse(text) if text else None, None))
            except Exception as e:  # noqa: BLE001
                self.names.append(DefinedName(nm, "", None, False, None, f"未対応: {e}"))
        for cell, g in formulas:
            try:
                text = dec.decode(g)
                cell.formula = text
                cell.formula_kind = "normal"
                cell.ast = parse(text)
            except Unsupported as e:
                cell.formula = cell.formula or "?"
                cell.parse_error = f"未対応: {e}"
            except Exception as e:  # noqa: BLE001
                cell.formula = cell.formula or "?"
                cell.parse_error = f"構文エラー: {e}"
        if not self.sheets:
            raise BiffError("ワークシートがありません")
        if self.guessed_codepage:
            self.sheets[0].unread.append("CODEPAGE の記録がないので、字の文字コードをバイトの並びから推定しました（Shift_JIS か Windows-1252）。字が正しいかはエージェントが確かめます")

    def _xf34(self, b: bytes, t: int) -> None:
        ifnt, ifmt = b[0], b[1]
        prot = b[2]
        borders = struct.unpack_from("<I", b, 8)[0]
        bits = sum(v for v, sh in ((4, 0), (1, 8), (8, 16), (2, 24)) if (borders >> sh) & 0x07)
        fmt = self._fmts[ifmt] if ifmt < len(self._fmts) else "General"
        self.styles.append(Style(fmt, bool(prot & 1), bool(prot & 2), None, None, bits))

    def _attr_style(self, a: bytes, cache: dict) -> int:
        if a in cache:
            return cache[a]
        fmt = self._fmts[a[1] & 0x3F] if (a[1] & 0x3F) < len(self._fmts) else "General"
        bits = sum(v for v, m in ((1, 0x08), (2, 0x10), (4, 0x20), (8, 0x40)) if a[2] & m)
        self.styles.append(Style(fmt, bool(a[0] & 0x40), bool(a[0] & 0x80), None, None, bits))
        cache[a] = len(self.styles) - 1
        return cache[a]

    def _cell(self, sh: Sheet, t: int, b: bytes, cache: dict, ixfe):
        if self.biff == 2 and t in (0x0001, 0x0002, 0x0003, 0x0004, 0x0005, 0x0006):
            r, c = struct.unpack_from("<HH", b, 0)
            xf = self._attr_style(b[4:7], cache)
            p = 7
            if t == 0x0001:
                v, g = None, None
            elif t == 0x0002:
                v, g = float(struct.unpack_from("<H", b, p)[0]), None
            elif t == 0x0003:
                v, g = struct.unpack_from("<d", b, p)[0], None
            elif t == 0x0004:
                v, g = self._text(b[p + 1:p + 1 + b[p]]), None
            elif t == 0x0005:
                v, g = (ExcelError(ERRORS.get(b[p], "#N/A")) if b[p + 1] else bool(b[p])), None
            else:
                v = self._result(b[p:p + 8])
                sz = b[p + 9]
                g = b[p + 10:p + 10 + sz]
        elif self.biff != 2 and t in (0x0201, 0x0203, 0x0204, 0x0205, 0x027E, 0x0206, 0x0406):
            r, c, xf = struct.unpack_from("<HHH", b, 0)
            p = 6
            g = None
            if t == 0x0201:
                v = None
            elif t == 0x0203:
                v = struct.unpack_from("<d", b, p)[0]
            elif t == 0x027E:
                v = _rk(struct.unpack_from("<I", b, p)[0])
            elif t == 0x0204:
                ln = struct.unpack_from("<H", b, p)[0]
                v = self._text(b[p + 2:p + 2 + ln])
            elif t == 0x0205:
                v = ExcelError(ERRORS.get(b[p], "#N/A")) if b[p + 1] else bool(b[p])
            else:
                v = self._result(b[p:p + 8])
                sz = struct.unpack_from("<H", b, 16)[0]
                g = b[18:18 + sz]
        else:
            return None, None
        cell = Cell(sh.name, r + 1, c + 1, style=xf)
        cell.value = v
        if isinstance(v, str):
            cell.t = "s"
        elif isinstance(v, bool):
            cell.t = "b"
        sh.cells[(r + 1, c + 1)] = cell
        sh.max_row = max(sh.max_row, r + 1)
        sh.max_col = max(sh.max_col, c + 1)
        return cell, g

    def _result(self, res: bytes):
        if res[6:8] == b"\xff\xff":
            kind = res[0]
            return {1: bool(res[2]), 2: ExcelError(ERRORS.get(res[2], "#N/A")), 3: ""}.get(kind, "")
        return struct.unpack_from("<d", res, 0)[0]


def raw_values(wb) -> dict[str, set[tuple[int, int]]]:
    """窓とは別に：セルのレコードの行・列を順に歩く（検査用）。"""
    out: dict[str, set] = {}
    blob = wb.zip.original
    current = None
    k = 0
    sheets = wb.sheets
    for t, b in _records(blob):
        if t in BOFS:
            dt = struct.unpack_from("<H", b, 2)[0]
            if dt in (0x10, 0x40):
                current = sheets[k].name
                out[current] = set()
                k += 1
            else:
                current = None
        elif current and t in (0x0002, 0x0003, 0x0004, 0x0005, 0x0006, 0x0203, 0x0204, 0x0205, 0x027E, 0x0206, 0x0406) and len(b) >= 4:
            r, c = struct.unpack_from("<HH", b, 0)
            out[current].add((r + 1, c + 1))
    return out
