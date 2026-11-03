"""OpenDocument の表計算（.ods・.ots・.fods）を読み、Excel（.xlsx）と同じ形（シート・セル・数式・入力規則・保護・結合・
非表示の行と列・書式）にする（ODF 1.3 第3部 9 章・16 章、第4部）。

読んだあとは、窓・入力欄の推定・数式の計算と照合を、Excel と同じ処理で行う。違うのは：
- セルの見える字は、ファイルに保存された字（text:p）をそのまま使う（ODF は見える字も保存している）
- 数式は OpenFormula を Excel の書き方にして計算する（odf.formula。元の数式も窓に出す）
- 入力規則は table:content-validation の条件（cell-content-is-in-list など）を、Excel の入力規則の形にする
行・列・セルの繰り返し（table:number-…-repeated）は、字や書式のある範囲だけを広げる（何百万行の空の行は広げない）。
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from ..formula.dates import date_to_serial
from ..formula.parser import parse
from ..formula.values import ExcelError, Unsupported
from ..xlsx.workbook import Cell, DefinedName, Shape, Sheet, Style, Validation
from .formula import OdfFormulaError, cell_address, range_address, to_excel
from .numstyle import excel_format
from .package import OdfError, OdfPackage, text_of
from .styles import Styles
from .text import INLINE, MARKS, SHAPES
from .xml import NS, El, Txt, q, top

_T = "{" + NS["text"] + "}"
_TABLE = "{" + NS["table"] + "}"
_OFFICE = "{" + NS["office"] + "}"
BIG = 64  # これより多く繰り返す、空のセル・行は広げない
ERRORS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")


class _Shim:
    """xlsx の窓が使う wb.zip.original の代わり。"""

    def __init__(self, blob: bytes):
        self.original = blob


@dataclass
class CellEl:
    el: El
    c0: int
    n: int
    covered: bool


@dataclass
class RowEl:
    el: El
    r0: int
    n: int
    cells: list[CellEl] = field(default_factory=list)


def para_text(el: El, data: bytes) -> str:
    """セルの中の段落（text:p）の字：ODF の空白の決まり（6.1.2）で作る。"""
    atoms: list = []

    def walk(n: El) -> None:
        for k in n.kids:
            if isinstance(k, Txt):
                atoms.extend(("c", c) for c in k.value(data))
                continue
            t = k.tag
            if t == _T + "s":
                atoms.append(("x", " " * int(k.get("text:c", "1") or 1)))
            elif t == _T + "tab":
                atoms.append(("x", "\t"))
            elif t == _T + "line-break":
                atoms.append(("x", "\n"))
            elif t in INLINE:
                walk(k)
            elif t in MARKS:
                continue
            elif t == _T + "ruby":
                base = k.find("text:ruby-base")
                if base is not None:
                    walk(base)
            elif t == _OFFICE + "annotation" or t in SHAPES:
                atoms.append(("sep",))
            elif t.startswith(_T) or t.startswith(_OFFICE):
                s = text_of(k, data)
                atoms.append(("x", s) if s else ("sep",))
            elif k.kids:
                walk(k)

    walk(el)
    out: list[str] = []
    prev_space = False
    at_start = True
    for a in atoms:
        if a[0] == "c":
            c = " " if a[1] in "\t\r\n" else a[1]
            if c == " ":
                if at_start or prev_space:
                    continue
                prev_space = True
            else:
                prev_space = False
            at_start = False
            out.append(c)
        else:
            prev_space = False
            at_start = False
            if a[0] == "x":
                out.append(("\x00" if a[1] == "" else "") + a[1])
    while out and out[-1] == " ":
        out.pop()
    return "".join(out).replace("\x00", "")


def cell_text(el: El, data: bytes) -> str:
    """セルの見える字：段落ごとに改行でつなぐ。"""
    paras = [k for k in el.els() if k.tag in (_T + "p", _T + "h")]
    return "\n".join(para_text(p, data) for p in paras)


def _bool(v: str | None) -> bool:
    return (v or "").strip().lower() == "true"


def _iso_date(s: str) -> dt.datetime | dt.date:
    s = s.strip()
    m = re.fullmatch(r"(-?\d{4,})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2})(?::(\d{2})(\.\d+)?)?)?(?:Z|[+-]\d{2}:\d{2})?", s)
    if not m:
        raise ValueError(f"日付を読めません：{s}")
    y, mo, d = int(m[1]), int(m[2]), int(m[3])
    if m[4] is None:
        return dt.date(y, mo, d)
    sec = int(m[6] or 0)
    frac = float(m[7] or 0)
    return dt.datetime(y, mo, d, int(m[4]), int(m[5]), sec, int(round(frac * 1e6)))


def _duration(s: str) -> float:
    """「PT12H30M00S」→ 日の割合。"""
    m = re.fullmatch(r"(-)?P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?", s.strip())
    if not m:
        raise ValueError(f"時刻を読めません：{s}")
    days = int(m[2] or 0) + (int(m[3] or 0) * 3600 + int(m[4] or 0) * 60 + float(m[5] or 0)) / 86400
    return -days if m[1] else days


def _borders(get) -> int:
    bits = 0
    all_ = get("fo:border")
    for k, side in ((1, "left"), (2, "right"), (4, "top"), (8, "bottom")):
        v = get("fo:border-" + side) or all_
        if v and v.strip() not in ("none", "hidden") and not v.strip().startswith(("0pt", "0cm", "0in", "0mm")):
            bits |= k
    return bits


def _color(v: str | None) -> str | None:
    if not v or v == "transparent":
        return None
    return v.lstrip("#").upper()


class OdsWorkbook:
    kind = "ods"

    def __init__(self, blob: bytes, filename: str = ""):
        self.blob = blob
        self.filename = filename
        self.pkg = OdfPackage(blob, filename)
        if self.pkg.kind not in ("ods", ""):
            raise OdfError(f"OpenDocument の表計算ではありません（{self.pkg.mimetype}）")
        self.zip = _Shim(blob)
        self.part = self.pkg.part_of("content.xml")
        self.data = self.pkg.xml(self.part)
        body = self.pkg.body()
        self.book = body.find("office:spreadsheet")
        if self.book is None:
            raise OdfError("表計算の本体（office:spreadsheet）がありません")
        self.st = Styles(self.pkg)
        self.sheets: list[Sheet] = []
        self.names: list[DefinedName] = []
        self.shared_strings: list[str] = []
        self.phonetic: dict = {}
        self.cell_sst: dict = {}
        self.styles: list[Style] = [Style("General", True, False)]
        self.style_names: list[str] = ["Default"]
        self.style_kind: list[str] = ["number"]
        self._style_idx: dict[str, int] = {}
        self.dxfs: list = []
        self.tables: dict = {}
        self.external: dict = {}
        self.external_links: list[str] = []
        self.date1904 = False
        self.calc_full_on_load = False
        self.workbook_protected = self.book.get("table:structure-protected") == "true"
        self.has_macros = self.pkg.has_macros
        self._vba = None
        self.wb_path = self.part
        self.shown: dict[tuple[str, int, int], str] = {}  # セルの見える字（text:p）
        self.odf_formula: dict[tuple[str, int, int], str] = {}  # 元の数式（OpenFormula）
        self.value_type: dict[tuple[str, int, int], str] = {}  # office:value-type
        self.rows: dict[str, list[RowEl]] = {}  # シート → 行の要素（書き込みで使う）
        self.unexpanded: list[str] = []  # 広げなかった、字のある繰り返し
        self.null_date = dt.date(1899, 12, 30)
        cs = self.book.find("table:calculation-settings")
        if cs is not None:
            nd = cs.find("table:null-date")
            if nd is not None and nd.get("table:date-value"):
                try:
                    self.null_date = _iso_date(nd.get("table:date-value"))
                except ValueError:
                    pass
        self.date1904 = self.null_date == dt.date(1904, 1, 1)
        self._validations = self._read_validations()
        for t in self.book.findall("table:table"):
            self._sheet(t)
        self._names(self.book, None)
        for sh, t in zip(self.sheets, self.book.findall("table:table")):
            self._names(t, sh.name)
        for d in self.names:
            if d.ast is None and d.error is None:
                try:
                    d.ast = parse(d.text)
                except Unsupported as e:
                    d.error = f"未対応: {e}"
                except Exception as e:  # noqa: BLE001
                    d.error = f"構文エラー: {e}"
        self._check_formats()

    # ------------------------------------------------------------------ xlsx と同じ問い合わせ
    @property
    def vba(self):
        return None

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
        return self.styles[idx] if 0 <= idx < len(self.styles) else self.styles[0]

    def defined_name(self, name: str, sheet: str | None):
        low = name.lower()
        local = [d for d in self.names if d.name.lower() == low and d.sheet_scope == sheet]
        if local:
            return local[0]
        glob = [d for d in self.names if d.name.lower() == low and d.sheet_scope is None]
        return glob[0] if glob else None

    def shown_text(self, cell: Cell) -> str | None:
        """xlsx の窓の display() が使う：保存された見える字。"""
        return self.shown.get((cell.sheet, cell.row, cell.col))

    # ------------------------------------------------------------------ 書式
    def style_index(self, name: str | None) -> int:
        name = name or "Default"
        if name in self._style_idx:
            return self._style_idx[name]
        st, part = self.st, self.part
        prop = lambda attr: st.prop(part, "table-cell", name, "style:table-cell-properties", attr)  # noqa: E731
        protect = prop("style:cell-protect") or "protected"
        ds = st.attr(part, "table-cell", name, "style:data-style-name")
        nel = st.number_style(part, ds)
        fmt, kind = excel_format(nel, self.pkg.xml(self._num_part(ds)), lambda n: st.number_style(part, n))
        s = Style(fmt or "General", protect not in ("none", ""), "hidden" in protect or "formula-hidden" in protect,
                  _color(st.prop(part, "table-cell", name, "style:text-properties", "fo:color")),
                  _color(prop("fo:background-color")), _borders(prop))
        idx = len(self.styles)
        self.styles.append(s)
        self.style_names.append(name)
        self.style_kind.append(kind if nel is not None else "number")
        self._style_idx[name] = idx
        if nel is not None and not fmt:
            s.unconvertible = True  # Excel の書き方にできない書式（書いた値の見える字は作らない）
        return idx

    def _num_part(self, ds: str | None) -> str:
        if self.pkg.flat:
            return self.part
        return self.part if ds and ds in self.st.numbers.get(self.part, {}) else "styles.xml" if self.pkg.has("styles.xml") else self.part

    def _check_formats(self) -> None:
        """書式ごとに、Excel の書き方にした表示形式で作った字が、保存された見える字と同じかを照らす。"""
        from ..formula.numfmt import format_value

        self.format_ok: dict[int, bool] = {}
        for sh in self.sheets:
            for (r, c), cell in sh.cells.items():
                if cell.value is None or isinstance(cell.value, (str, ExcelError)) or cell.style in self.format_ok and not self.format_ok[cell.style]:
                    continue
                shown = self.shown.get((sh.name, r, c))
                if shown is None:
                    continue
                st = self.style(cell.style)
                try:
                    got = format_value(cell.value, st.num_fmt) if not getattr(st, "unconvertible", False) else None
                except Exception:  # noqa: BLE001
                    got = None
                ok = got == shown
                self.format_ok[cell.style] = self.format_ok.get(cell.style, True) and ok

    # ------------------------------------------------------------------ 入力規則
    def _read_validations(self) -> dict[str, dict]:
        out: dict[str, dict] = {}
        cv = self.book.find("table:content-validations")
        if cv is None:
            return out
        for v in cv.findall("table:content-validation"):
            name = v.get("table:name", "")
            hm = v.find("table:help-message")
            em = v.find("table:error-message")
            em_macro = v.find("table:error-macro")
            d = {"condition": v.get("table:condition", ""), "allow_blank": v.get("table:allow-empty-cell", "true") != "false",
                 "base": v.get("table:base-cell-address"), "display_list": v.get("table:display-list", "unsorted"),
                 "prompt_title": hm.get("table:title") if hm is not None else None,
                 "prompt": "\n".join(para_text(p, self.data) for p in hm.findall("text:p")) if hm is not None else None,
                 "show_input": hm is not None and hm.get("table:display") == "true",
                 "error_title": em.get("table:title") if em is not None else None,
                 "error": "\n".join(para_text(p, self.data) for p in em.findall("text:p")) if em is not None else None,
                 "show_error": em is not None and em.get("table:display", "true") == "true",
                 "error_style": (em.get("table:message-type", "stop") if em is not None else "stop"),
                 "macro": em_macro is not None, "cells": []}
            out[name] = d
        return out

    def _validation(self, d: dict) -> Validation:
        typ, op, f1, f2 = condition(d["condition"])
        sq = _rects(d["cells"])
        v = Validation(sq, typ, op, f1, f2, d["allow_blank"], d["prompt"] or None, d["prompt_title"] or None, d["error"] or None,
                       " ".join(f"{_a(r1, c1)}" + (f":{_a(r2, c2)}" if (r1, c1) != (r2, c2) else "") for r1, c1, r2, c2 in sq),
                       d["error_title"] or None, d["error_style"] or "stop", d["show_error"], d["show_input"])
        v.odf_condition = d["condition"]
        return v

    # ------------------------------------------------------------------ シート
    def _sheet(self, t: El) -> None:
        name = t.get("table:name", f"Sheet{len(self.sheets) + 1}")
        disp = self.st.prop(self.part, "table", t.get("table:style-name"), "style:table-properties", "table:display")
        sh = Sheet(name=name, path=self.part, state="hidden" if disp == "false" else "visible")
        sh.protected = t.get("table:protected") == "true"
        sh.sheet_id = len(self.sheets) + 1
        pr = t.get("table:print-ranges")
        if pr:
            sh.print_area = pr
            for part in pr.split():
                try:
                    s_, r1, c1, r2, c2 = range_address(part)
                except (OdfFormulaError, ValueError):
                    continue
                if s_ in (None, name):
                    sh.print_ranges.append((r1, c1, r2, c2))
        self.sheets.append(sh)
        # 列
        col = 1
        col_style: dict[int, str] = {}
        for ce in _columns(t):
            n = int(ce.get("table:number-columns-repeated", "1") or 1)
            dcs = ce.get("table:default-cell-style-name")
            hidden = ce.get("table:visibility", "visible") != "visible"
            if n <= 1024:
                for k in range(n):
                    if dcs:
                        col_style[col + k] = dcs
                    if hidden:
                        sh.hidden_cols.add(col + k)
            col += n
        for c, s in col_style.items():
            if s != "Default":
                sh.col_styles[c] = self.style_index(s)
        # 行
        rows: list[RowEl] = []
        r = 1
        validations: dict[str, dict] = self._validations
        for re_ in _rows(t):
            n = int(re_.get("table:number-rows-repeated", "1") or 1)
            rec = RowEl(re_, r, n)
            rows.append(rec)
            hidden = re_.get("table:visibility", "visible") != "visible"
            rstyle = re_.get("table:style-name")
            height = self.st.prop(self.part, "table-row", rstyle, "style:table-row-properties", "style:row-height")
            dcs = re_.get("table:default-cell-style-name")
            c = 1
            meaningful = False
            styled = False  # 枠・塗り・保護の解除のある書式のセル（記入欄の行が何百行も続く様式）
            for ce in re_.els():
                if ce.tag not in (_TABLE + "table-cell", _TABLE + "covered-table-cell"):
                    continue
                m = int(ce.get("table:number-columns-repeated", "1") or 1)
                rec.cells.append(CellEl(ce, c, m, ce.tag == _TABLE + "covered-table-cell"))
                c += m
                if _has_content(ce):
                    meaningful = True
                elif m <= BIG and ce.get("table:style-name"):
                    s_ = self.style(self.style_index(ce.get("table:style-name")))
                    styled = styled or bool(s_.borders or s_.fill_color or not s_.locked)
            expand = n if (n <= BIG or meaningful or (styled and n <= 2000)) else 0
            if styled and not meaningful and n > 2000:
                # 枠・塗り・保護の解除のある、中身のない行の長い繰り返し：広げない（窓の大きさを抑える）。広げなかったことは伝える
                self.unexpanded.append(f"{name}：{r}行目から {n:,} 行（枠・塗りなどの書式だけの、中身のない行の繰り返し。"
                                       "空の欄の候補として広げていません）")
            if meaningful and n > 10000:
                self.unexpanded.append(f"{name}：{r}行目から {n:,} 行（同じ中身を繰り返す行。最初の 10,000 行だけ読みました）")
                expand = 10000
            for k in range(expand):
                rr = r + k
                if hidden:
                    sh.hidden_rows.add(rr)
                if height and _pt(height) is not None:
                    sh.row_heights[rr] = _pt(height)
                if dcs and dcs != "Default":
                    sh.row_styles[rr] = self.style_index(dcs)
                for cel in rec.cells:
                    self._cells(sh, cel, rr, dcs, col_style, validations, meaningful)
            r += n
        self.rows[name] = rows
        sh.max_row = max((rr for rr, _c in sh.cells), default=0)
        sh.max_col = max((cc for _r, cc in sh.cells), default=0)
        for d in validations.values():
            cells = [x for x in d["cells"] if x[0] == name]
            if cells:
                v = self._validation({**d, "cells": [(rr, cc) for _s, rr, cc in cells]})
                sh.validations.append(v)
        # シートに置いた図形（table:shapes）
        shapes = t.find("table:shapes")
        if shapes is not None:
            for s in shapes.els():
                self._shape(sh, s, None)
        self._header_footer(sh, t)

    def _cells(self, sh: Sheet, cel: CellEl, r: int, row_dcs: str | None, col_style: dict, validations: dict,
               meaningful_row: bool) -> None:
        el = cel.el
        content = _has_content(el)
        sname = el.get("table:style-name")
        n = cel.n
        if not content and n > BIG:
            return  # 広い範囲の、字のない書式だけのセル（行の終わりまでの塗りなど）は広げない
        if content and n > 1024:
            self.unexpanded.append(f"{sh.name}：{r}行目の {cel.c0} 列目から {n:,} 列（同じ中身。最初の 1,024 列だけ読みました）")
            n = 1024
        for k in range(n):
            c = cel.c0 + k
            style = sname or row_dcs or col_style.get(c)
            if not content and not sname and not el.get("table:content-validation-name"):
                continue  # 書式も中身もないセル（列・行の書式は、Sheet.style_of が使う）
            self._cell(sh, el, r, c, style, validations, cel.covered)

    def _cell(self, sh: Sheet, el: El, r: int, c: int, style: str | None, validations: dict, covered: bool) -> None:
        key = (sh.name, r, c)
        cell = Cell(sh.name, r, c, style=self.style_index(style) if style else 0)
        vt = el.get("office:value-type")
        cvt = el.get("calcext:value-type")
        text = cell_text(el, self.data)
        has_text = any(k.tag in (_T + "p", _T + "h") for k in el.els())
        if vt in ("float", "percentage", "currency"):
            try:
                cell.value = float(el.get("office:value", "0") or 0)
            except ValueError:
                cell.value = text
        elif vt == "date":
            try:
                d = _iso_date(el.get("office:date-value", ""))
                serial = (d.date() if isinstance(d, dt.datetime) else d) - self.null_date
                x = float(serial.days)
                if self.null_date == dt.date(1899, 12, 30):
                    x = float(date_to_serial(d.date() if isinstance(d, dt.datetime) else d, False))
                if isinstance(d, dt.datetime):
                    x += (d.hour * 3600 + d.minute * 60 + d.second + d.microsecond / 1e6) / 86400
                cell.value = x
            except ValueError:
                cell.value = text
        elif vt == "time":
            try:
                cell.value = _duration(el.get("office:time-value", ""))
            except ValueError:
                cell.value = text
        elif vt == "boolean":
            cell.value = _bool(el.get("office:boolean-value"))
        elif vt == "string":
            sv = el.get("office:string-value")
            cell.value = sv if sv is not None else text
            cell.t = "str"
        elif has_text and text:
            cell.value = text
        if cvt == "error" or (isinstance(cell.value, (int, float)) and text in ERRORS and el.get("table:formula")):
            cell.value = ExcelError(text) if text in ERRORS else ExcelError("#VALUE!")
        f = el.get("table:formula")
        if f:
            self.odf_formula[key] = f
            try:
                cell.formula = to_excel(f)
                cell.ast = parse(cell.formula)
            except (OdfFormulaError, Unsupported) as e:
                cell.formula = f  # 計算できない書き方：元の数式のまま（窓にも元の数式を出す）
                cell.ast = None
                cell.parse_error = str(e)
            except Exception as e:  # noqa: BLE001
                cell.ast = None
                cell.parse_error = f"構文エラー: {e}"
            if el.get("table:number-matrix-columns-spanned") or el.get("table:number-matrix-rows-spanned"):
                cell.formula_kind = "array"
        if vt:
            self.value_type[key] = vt
        if has_text or vt:
            self.shown[key] = text
        vn = el.get("table:content-validation-name")
        if vn and vn in validations and not covered:
            validations[vn]["cells"].append(key)
        if not covered:
            cs = int(el.get("table:number-columns-spanned", "1") or 1)
            rs = int(el.get("table:number-rows-spanned", "1") or 1)
            if cs > 1 or rs > 1:
                sh.merges.append((r, c, r + rs - 1, c + cs - 1))
        ann = el.find("office:annotation")
        if ann is not None:
            sh.comments[(r, c)] = "\n".join(para_text(p, self.data) for p in ann.findall("text:p"))
        for k in el.els():
            if k.tag in SHAPES:
                self._shape(sh, k, (r, c))
        sh.cells[(r, c)] = cell

    def _shape(self, sh: Sheet, el: El, at: tuple[int, int] | None) -> None:
        data = self.data
        if el.tag not in SHAPES and el.tag != q("draw:control"):
            return
        if el.tag == q("draw:g"):
            for k in el.els():
                self._shape(sh, k, at)
            return
        texts = []
        for k in el.iter():
            if k.tag in (_T + "p", _T + "h"):
                texts.append(para_text(k, data))
        title = el.find("svg:title")
        desc = el.find("svg:desc")
        alt = " ".join(x for x in ((text_of(title, data).strip() if title is not None else ""),
                                   (text_of(desc, data).strip() if desc is not None else "")) if x)
        kind = "text"
        if el.find("draw:image") is not None:
            kind = "picture"
        if el.find("draw:object") is not None or el.find("draw:object-ole") is not None:
            kind = "chart"
        if el.tag == q("draw:control"):
            kind = "control"
        if at is None:
            end = el.get("table:end-cell-address")
            if end:
                try:
                    _s, rr, cc = cell_address(end)
                    at = (rr, cc)
                except (OdfFormulaError, ValueError):
                    at = None
        sh.shapes.append(Shape(kind, el.get("draw:name", "") or "", "\n".join(t for t in texts if t), alt, at))

    def _header_footer(self, sh: Sheet, t: El) -> None:
        mp_name = self.st.attr(self.part, "table", t.get("table:style-name"), "style:master-page-name")
        part = self.pkg.part_of("styles.xml")
        ms = self.pkg.section(part, "master-styles") if (self.pkg.flat or self.pkg.has("styles.xml")) else None
        if ms is None:
            return
        data = self.pkg.xml(part)
        pages = ms.findall("style:master-page")
        mp = next((m for m in pages if m.get("style:name") == mp_name), pages[0] if pages else None)
        if mp is None:
            return
        for el in mp.els():
            if el.get("style:display") == "false":
                continue
            if el.local.startswith(("header", "footer")):
                parts = []
                for reg in el.iter():
                    if reg.tag in (_T + "p", _T + "h"):
                        s = para_text(reg, data)
                        if s:
                            parts.append(s)
                if parts:
                    sh.header_footer.append(("ヘッダー：" if el.local.startswith("header") else "フッター：") + "　".join(parts))

    # ------------------------------------------------------------------ 名前
    def _names(self, node: El, scope: str | None) -> None:
        ne = node.find("table:named-expressions")
        if ne is None:
            return
        for el in ne.els():
            name = el.get("table:name", "")
            if not name:
                continue
            if el.tag == _TABLE + "named-range":
                addr_ = el.get("table:cell-range-address", "")
                try:
                    text = to_excel("[" + addr_ + "]")
                except OdfFormulaError as e:
                    self.names.append(DefinedName(name, addr_, scope, False, None, str(e)))
                    continue
            elif el.tag == _TABLE + "named-expression":
                try:
                    text = to_excel(el.get("table:expression", ""))
                except OdfFormulaError as e:
                    self.names.append(DefinedName(name, el.get("table:expression", ""), scope, False, None, str(e)))
                    continue
            else:
                continue
            self.names.append(DefinedName(name, text, scope, False))


# ====================================================================== 入力規則の条件
def _pair(s: str) -> list[str]:
    """cell-content-is-between(値, 値) などの2つの値：ODF 1.2 第1部 19.595 table:condition の文法では「,」で分ける
    （LibreOffice もそう書く）。「;」で書いたものも読む。"""
    x = _args(s, ",")
    return x if len(x) > 1 else _args(s)


def _args(s: str, sep: str = ";") -> list[str]:
    """引数を、かっこ・引用符・角かっこの外の「;」（sep）で分ける。"""
    out, cur, depth, q_, br = [], [], 0, False, 0
    for ch in s:
        if ch == '"':
            q_ = not q_
        elif not q_:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
            elif ch == "[":
                br += 1
            elif ch == "]":
                br -= 1
            elif ch == sep and depth == 0 and br == 0:
                out.append("".join(cur).strip())
                cur = []
                continue
        cur.append(ch)
    if "".join(cur).strip():
        out.append("".join(cur).strip())
    return out


_OPS = {"<=": "lessThanOrEqual", ">=": "greaterThanOrEqual", "!=": "notEqual", "<>": "notEqual", "<": "lessThan",
        ">": "greaterThan", "=": "equal"}


def _call(s: str, name: str) -> str | None:
    """s の中の name(…) の中身（かっこの対応を見て）。"""
    i = s.find(name + "(")
    if i < 0:
        return None
    j = i + len(name) + 1
    depth, q_ = 1, False
    k = j
    while k < len(s):
        ch = s[k]
        if ch == '"':
            q_ = not q_
        elif not q_:
            if ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    return s[j:k]
        k += 1
    return None


def condition(cond: str) -> tuple[str, str | None, str | None, str | None]:
    """table:condition → (種類, 比べ方, 式1, 式2)。種類は Excel の入力規則と同じ名前（list・whole・decimal・date・time・
    textLength・custom）。読めない条件は custom（式1 に元の条件）。"""
    s = cond.strip()
    m = re.match(r"^([A-Za-z][\w.-]*):", s)
    if m:
        s = s[m.end():]
    typ = None
    for k, t in (("cell-content-is-whole-number", "whole"), ("cell-content-is-decimal-number", "decimal"),
                 ("cell-content-is-date", "date"), ("cell-content-is-time", "time")):
        if k + "()" in s:
            typ = t
    lst = _call(s, "cell-content-is-in-list")
    try:
        if lst is not None:
            items = _args(lst)
            lits = [x for x in items if re.fullmatch(r'"(?:[^"]|"")*"|-?\d+(?:\.\d+)?', x)]
            if len(lits) == len(items):
                vals = [x[1:-1].replace('""', '"') if x.startswith('"') else x for x in items]
                if any("," in v or '"' in v for v in vals):
                    return "list", None, "{" + ",".join('"' + v.replace('"', '""') + '"' for v in vals) + "}", None
                return "list", None, '"' + ",".join(vals) + '"', None
            return "list", None, to_excel(lst), None
        for name, op in (("cell-content-text-length-is-between", "between"), ("cell-content-text-length-is-not-between", "notBetween")):
            a = _call(s, name)
            if a is not None:
                x = _pair(a)
                return "textLength", op, to_excel(x[0]), to_excel(x[1]) if len(x) > 1 else None
        mm = re.search(r"cell-content-text-length\(\)\s*(<=|>=|!=|<>|<|>|=)\s*(.+)$", s)
        if mm:
            return "textLength", _OPS[mm[1]], to_excel(mm[2].strip()), None
        for name, op in (("cell-content-is-between", "between"), ("cell-content-is-not-between", "notBetween")):
            a = _call(s, name)
            if a is not None:
                x = _pair(a)
                return typ or "decimal", op, to_excel(x[0]), to_excel(x[1]) if len(x) > 1 else None
        mm = re.search(r"cell-content\(\)\s*(<=|>=|!=|<>|<|>|=)\s*(.+)$", s)
        if mm:
            return typ or "decimal", _OPS[mm[1]], to_excel(mm[2].strip()), None
        f = _call(s, "is-true-formula")
        if f is not None:
            return "custom", None, to_excel(f), None
    except OdfFormulaError:
        return "custom", None, cond, None
    if typ:
        return typ, None, None, None
    return "custom", None, cond, None


# ====================================================================== 補助
def _a(r: int, c: int) -> str:
    from ..formula.parser import addr

    return addr(r, c)


def _rects(cells: list[tuple[int, int]]) -> list[tuple[int, int, int, int]]:
    """セルの一覧 → 行ごとに続いた範囲の並び。"""
    out = []
    by_row: dict[int, list[int]] = {}
    for r, c in cells:
        by_row.setdefault(r, []).append(c)
    for r in sorted(by_row):
        cs = sorted(set(by_row[r]))
        start = prev = cs[0]
        for c in cs[1:]:
            if c == prev + 1:
                prev = c
                continue
            out.append((r, start, r, prev))
            start = prev = c
        out.append((r, start, r, prev))
    return out


def _pt(length: str) -> float | None:
    """ODF の長さ（数＋単位。ODF 1.2 第 1 部 18.3.18 length：cm・mm・in・pt・pc・px）の pt。読めなければ None
    （値を作らない。呼び手は、その長さのない所として扱う）。"""
    m = re.fullmatch(r"([\d.]+)(cm|mm|in|pt|pc|px)?", (length or "").strip())
    if not m:
        return None
    v = float(m[1])
    return {"cm": v * 28.3465, "mm": v * 2.83465, "in": v * 72, "pt": v, "pc": v * 12, "px": v * 0.75}.get(m[2] or "pt", v)


def _rows(t: El):
    for ch in t.els():
        if ch.tag == _TABLE + "table-row":
            yield ch
        elif ch.tag in (_TABLE + "table-header-rows", _TABLE + "table-rows", _TABLE + "table-row-group"):
            yield from _rows(ch)


def _columns(t: El):
    for ch in t.els():
        if ch.tag == _TABLE + "table-column":
            yield ch
        elif ch.tag in (_TABLE + "table-header-columns", _TABLE + "table-columns", _TABLE + "table-column-group"):
            yield from _columns(ch)


def _has_content(ce: El) -> bool:
    """字・値・数式・入力規則・コメント・図形・結合のあるセルか。"""
    if ce.get("office:value-type") or ce.get("table:formula") or ce.get("table:content-validation-name"):
        return True
    if ce.get("table:number-columns-spanned") or ce.get("table:number-rows-spanned"):
        return True
    return any(k.tag in (_T + "p", _T + "h", _OFFICE + "annotation") or k.tag in SHAPES for k in ce.els())
