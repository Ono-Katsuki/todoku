"""xlsx / xlsm を読み、セル・数式・入力規則・名前付き範囲などを取り出す。

ここでは読むだけで、元のファイルには一切手を加えない。
"""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from ..formula.parser import Node, addr, parse, parse_cell, unparse
from ..formula.values import ExcelError, Unsupported
from ..zipraw import RawZip

NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xm": "http://schemas.microsoft.com/office/excel/2006/main",
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
M = "{%s}" % NS["m"]
R = "{%s}" % NS["r"]


# Strict Open XML（ISO/IEC 29500 Strict）の名前空間 → 通常（Transitional）の名前空間
_STRICT = [
    (b"http://purl.oclc.org/ooxml/spreadsheetml/main", b"http://schemas.openxmlformats.org/spreadsheetml/2006/main"),
    (b"http://purl.oclc.org/ooxml/officeDocument/relationships", b"http://schemas.openxmlformats.org/officeDocument/2006/relationships"),
    (b"http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing", b"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
    (b"http://purl.oclc.org/ooxml/drawingml/main", b"http://schemas.openxmlformats.org/drawingml/2006/main"),
]


def parse_xml(data: bytes) -> ET.Element:
    """XMLを読む。Strict Open XML で書かれていても、同じ名前で読めるように名前空間をそろえる。"""
    if b"purl.oclc.org/ooxml" in data:
        for a, b in _STRICT:
            data = data.replace(a, b)
    return ET.fromstring(data)


def _q(tag: str) -> str:
    p, t = tag.split(":")
    return "{%s}%s" % (NS[p], t)


@dataclass
class Cell:
    sheet: str
    row: int
    col: int
    t: str | None = None  # XMLの t 属性
    style: int = 0
    value: object = None  # 保存されている値（数式なら最後に計算された値）
    formula: str | None = None  # 数式（共有数式は展開後の文字列）
    formula_kind: str | None = None  # normal / shared / array / dataTable
    array_ref: str | None = None
    ast: Node | None = None
    dr: int = 0
    dc: int = 0
    parse_error: str | None = None
    array_master: tuple[int, int] | None = None  # 配列数式の一部なら親セル

    @property
    def address(self) -> str:
        return addr(self.row, self.col)

    @property
    def has_formula(self) -> bool:
        return self.formula is not None or self.array_master is not None


@dataclass
class Validation:
    sqref: list[tuple[int, int, int, int]]
    type: str  # list / whole / decimal / date / textLength / custom ...
    operator: str | None
    formula1: str | None
    formula2: str | None
    allow_blank: bool
    prompt: str | None
    prompt_title: str | None
    error: str | None
    sqref_text: str = ""
    error_title: str | None = None
    error_style: str = "stop"  # stop / warning / information
    show_error: bool = False  # 誤った値を入れたときにExcelが警告を出すか
    show_input: bool = False

    def covers(self, r: int, c: int) -> bool:
        return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in self.sqref)

    @property
    def origin(self) -> tuple[int, int]:
        """入力規則の数式の相対参照の基準：範囲全体の左上（範囲が「N20 N8:O8」の順に書かれていても N8）。"""
        if not self.sqref:
            return (1, 1)
        return min(r[0] for r in self.sqref), min(r[1] for r in self.sqref)

    @property
    def blocks(self) -> bool:
        """Excelで手入力したとき、規則に合わない値を受け付けない（停止）か。"""
        return self.show_error and self.error_style in ("stop", "", None)


@dataclass
class Shape:
    kind: str  # text / picture / chart / control / group
    name: str
    text: str
    alt: str
    anchor: tuple[int, int] | None  # (row, col) 左上


@dataclass
class Control:
    """フォームコントロール（チェックボックス・オプションボタン・ドロップダウンなど）。"""

    kind: str  # CheckBox / Radio / Drop / List / Spin / Scroll / Button / Label / GBox / EditBox / ActiveX
    name: str
    shape_id: str
    anchor: tuple[int, int] | None = None  # 左上のセル（行, 列）
    anchor_to: tuple[int, int] | None = None
    link: str | None = None  # リンクするセル（書かれたまま）
    link_cell: tuple[str, int, int] | None = None
    checked: str | None = None  # Checked / Mixed / None（チェックなし）
    text: str = ""
    from_off: float = 0.0  # 左上のセルの上端からのずれ（ポイント）
    to_off: float = 0.0
    prop_part: str | None = None  # xl/ctrlProps/ctrlPropN.xml
    vml_part: str | None = None
    fmla_range: str | None = None

    @property
    def is_checked(self) -> bool:
        return self.checked == "Checked"


@dataclass
class Table:
    """テーブル（リストオブジェクト）。構造化参照で列を名前で指すのに使う。"""

    name: str
    sheet: str
    r1: int
    c1: int
    r2: int
    c2: int
    header_rows: int
    totals_rows: int
    columns: list[str]


@dataclass
class Sheet:
    name: str
    path: str
    state: str = "visible"
    cells: dict[tuple[int, int], Cell] = field(default_factory=dict)
    merges: list[tuple[int, int, int, int]] = field(default_factory=list)
    validations: list[Validation] = field(default_factory=list)
    protected: bool = False
    col_widths: dict[int, float] = field(default_factory=dict)
    default_col_width: float | None = None
    hidden_rows: set[int] = field(default_factory=set)
    hidden_cols: set[int] = field(default_factory=set)
    row_heights: dict[int, float] = field(default_factory=dict)
    default_row_height: float = 15.0
    col_styles: dict[int, int] = field(default_factory=dict)  # 列の既定の書式（セル要素がないセルに効く）
    row_styles: dict[int, int] = field(default_factory=dict)
    comments: dict[tuple[int, int], str] = field(default_factory=dict)
    shapes: list[Shape] = field(default_factory=list)
    hyperlinks: list[tuple[str, str]] = field(default_factory=list)
    print_area: str | None = None
    print_ranges: list[tuple[int, int, int, int]] = field(default_factory=list)
    controls: list[Control] = field(default_factory=list)
    header_footer: list[str] = field(default_factory=list)
    cond_formats: int = 0
    cond_rules: list[CondRule] = field(default_factory=list)
    autofilter: tuple[int, int, int, int] | None = None  # オートフィルターの範囲（ここで隠れた行は SUBTOTAL が常に除く）
    has_vml_controls: bool = False
    max_row: int = 0
    max_col: int = 0

    def merged_at(self, r: int, c: int) -> tuple[int, int, int, int] | None:
        for m in self.merges:
            if m[0] <= r <= m[2] and m[1] <= c <= m[3]:
                return m
        return None

    def style_of(self, r: int, c: int) -> int:
        """セルの書式の番号。セル要素がないときは、Excelと同じく行の書式、なければ列の書式を使う。"""
        cell = self.cells.get((r, c))
        if cell is not None:
            return cell.style
        if r in self.row_styles:
            return self.row_styles[r]
        return self.col_styles.get(c, 0)

    def is_hidden(self, r: int, c: int) -> bool:
        return r in self.hidden_rows or c in self.hidden_cols

    def in_print_area(self, r: int, c: int) -> bool:
        if not self.print_ranges:
            return True
        return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in self.print_ranges)


@dataclass
class DefinedName:
    name: str
    text: str
    sheet_scope: str | None
    hidden: bool
    ast: Node | None = None
    error: str | None = None


@dataclass
class Style:
    num_fmt: str
    locked: bool
    hidden: bool
    font_color: str | None = None  # "RRGGBB"（不明なら None＝自動の黒）
    fill_color: str | None = None  # 塗りつぶしの色（なければ None＝白）


@dataclass
class Dxf:
    """条件付き書式が当てる書式（変えるものだけ）。"""

    font_color: str | None = None
    fill_color: str | None = None
    num_fmt: str | None = None


@dataclass
class CondRule:
    sqref: list[tuple[int, int, int, int]]
    type: str
    priority: int
    dxf: int | None
    formulas: list[str]
    operator: str | None = None
    stop: bool = False

    @property
    def origin(self) -> tuple[int, int]:
        # 入力規則と同じく、範囲全体の左上を相対参照の基準にする
        return (min(r[0] for r in self.sqref), min(r[1] for r in self.sqref)) if self.sqref else (1, 1)

    def covers(self, r: int, c: int) -> bool:
        return any(r1 <= r <= r2 and c1 <= c <= c2 for r1, c1, r2, c2 in self.sqref)


class Workbook:
    def __init__(self, blob: bytes, filename: str = ""):
        self.filename = filename
        self.zip = RawZip(blob)
        self.sheets: list[Sheet] = []
        self.names: list[DefinedName] = []
        self.shared_strings: list[str] = []
        self.phonetic: dict[int, str] = {}  # 共有文字列の番号 → ふりがなを当てた読み（PHONETIC 関数の値）
        self._last_sst: int | None = None
        self.cell_sst: dict[tuple[str, int, int], int] = {}  # 共有文字列を使うセル → 番号
        self.styles: list[Style] = []
        self.dxfs: list[Dxf] = []
        self.tables: dict[str, Table] = {}
        self.external: dict[int, dict[str, dict[tuple[int, int], object]]] = {}  # [n] → シート名 → 保存値
        self.date1904 = False
        self.has_macros = self.zip.has("xl/vbaProject.bin")
        self.external_links = [n for n in self.zip.names() if n.startswith("xl/externalLinks/") and n.endswith(".xml")]
        self.calc_full_on_load = False
        self.workbook_protected = False
        self._load()

    # -------------------------------------------------------------- 補助
    def xml(self, name: str) -> ET.Element:
        return parse_xml(self.zip.read(name))

    def rels(self, part: str) -> dict[str, tuple[str, str]]:
        d, base = posixpath.split(part)
        rp = posixpath.join(d, "_rels", base + ".rels")
        if not self.zip.has(rp):
            return {}
        out = {}
        for rel in self.xml(rp):
            target = rel.get("Target", "")
            if rel.get("TargetMode") == "External":
                full = target
            elif target.startswith("/"):
                full = target[1:]
            else:
                full = posixpath.normpath(posixpath.join(d, target))
            out[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], full)
        return out

    def sheet(self, name: str) -> Sheet:
        for s in self.sheets:
            if s.name == name:
                return s
        lname = name.lower()
        for s in self.sheets:
            if s.name.lower() == lname:
                return s
        raise KeyError(name)

    def has_sheet(self, name: str) -> bool:
        try:
            self.sheet(name)
            return True
        except KeyError:
            return False

    # -------------------------------------------------------------- 読み込み
    def _load(self) -> None:
        wb_path = "xl/workbook.xml"
        root_rels = self.rels("")
        for typ, target in root_rels.values():
            if typ == "officeDocument":
                wb_path = target
        wb = self.xml(wb_path)
        pr = wb.find(M + "workbookPr")
        if pr is not None and pr.get("date1904") in ("1", "true"):
            self.date1904 = True
        cp = wb.find(M + "calcPr")
        if cp is not None and cp.get("fullCalcOnLoad") in ("1", "true"):
            self.calc_full_on_load = True
        self.workbook_protected = wb.find(M + "workbookProtection") is not None
        rels = self.rels(wb_path)
        self.wb_path = wb_path
        for typ, target in rels.values():
            if typ == "sharedStrings":
                self._load_shared_strings(target)
            elif typ == "styles":
                self._load_styles(target)
        sheets_el = wb.find(M + "sheets")
        for s in sheets_el if sheets_el is not None else []:
            rid = s.get(R + "id")
            typ, path = rels.get(rid, ("", ""))
            if typ != "worksheet":
                # グラフシート・ダイアログシートなどはセルを持たない
                self.sheets.append(Sheet(name=s.get("name"), path=path, state=f"non-worksheet:{typ}"))
                continue
            self.sheets.append(Sheet(name=s.get("name"), path=path, state=s.get("state", "visible")))
        er = wb.find(M + "externalReferences")
        for idx, e in enumerate(er if er is not None else [], start=1):
            typ, part = rels.get(e.get(R + "id"), ("", ""))
            if typ == "externalLink" and self.zip.has(part):
                self.external[idx] = _external_cache(self.xml(part))
        dn = wb.find(M + "definedNames")
        sheet_names = [s.name for s in self.sheets]
        for d in dn if dn is not None else []:
            scope = d.get("localSheetId")
            self.names.append(DefinedName(
                name=d.get("name"),
                text=d.text or "",
                sheet_scope=sheet_names[int(scope)] if scope is not None and int(scope) < len(sheet_names) else None,
                hidden=d.get("hidden") in ("1", "true"),
            ))
        for s in self.sheets:
            if s.state.startswith("non-worksheet"):
                continue
            self._load_sheet(s)
        for d in self.names:
            if d.name == "_xlnm.Print_Area" and d.sheet_scope:
                sh = self.sheet(d.sheet_scope)
                sh.print_area = d.text
                sh.print_ranges = _area_ranges(d.text)
            try:
                d.ast = parse(d.text)
            except Unsupported as e:
                d.error = f"未対応: {e}"
            except Exception as e:  # noqa: BLE001
                d.error = f"構文エラー: {e}"

    def _load_shared_strings(self, path: str) -> None:
        root = self.xml(path)
        for i, si in enumerate(root.findall(M + "si")):
            self.shared_strings.append(_rich_text(si))
            ph = _phonetic_text(si)
            if ph is not None:
                self.phonetic[i] = ph

    def _load_styles(self, path: str) -> None:
        from ..formula.numfmt import BUILTIN

        root = self.xml(path)
        fmts = dict(BUILTIN)
        nf = root.find(M + "numFmts")
        for f in nf if nf is not None else []:
            fmts[int(f.get("numFmtId"))] = f.get("formatCode")
        fonts = [_color(f.find(M + "color")) for f in (root.find(M + "fonts") or [])]
        fills = [_fill_color(f) for f in (root.find(M + "fills") or [])]
        xfs = root.find(M + "cellXfs")
        for xf in xfs if xfs is not None else []:
            fid = int(xf.get("numFmtId", "0"))
            prot = xf.find(M + "protection")
            locked = True
            hidden = False
            if prot is not None:
                locked = prot.get("locked", "1") not in ("0", "false")
                hidden = prot.get("hidden", "0") in ("1", "true")
            fo = int(xf.get("fontId", "0"))
            fi = int(xf.get("fillId", "0"))
            self.styles.append(Style(fmts.get(fid, "General"), locked, hidden,
                                     fonts[fo] if fo < len(fonts) else None, fills[fi] if fi < len(fills) else None))
        for d in root.find(M + "dxfs") or []:
            font = d.find(M + "font")
            nf = d.find(M + "numFmt")
            self.dxfs.append(Dxf(_color(font.find(M + "color")) if font is not None else None,
                                 _fill_color(d.find(M + "fill"), dxf=True) if d.find(M + "fill") is not None else None,
                                 nf.get("formatCode") if nf is not None else None))

    def style(self, idx: int) -> Style:
        if 0 <= idx < len(self.styles):
            return self.styles[idx]
        return Style("General", True, False)

    def _load_sheet(self, sh: Sheet) -> None:
        root = self.xml(sh.path)
        shared: dict[str, tuple[Cell, Node | None, str | None]] = {}
        pending_shared: list[tuple[Cell, str]] = []
        array_refs: list[tuple[Cell, tuple[int, int, int, int]]] = []
        fmt = root.find(M + "sheetFormatPr")
        if fmt is not None and fmt.get("defaultColWidth"):
            sh.default_col_width = float(fmt.get("defaultColWidth"))
        if fmt is not None and fmt.get("defaultRowHeight"):
            sh.default_row_height = float(fmt.get("defaultRowHeight"))
        cols = root.find(M + "cols")
        for c in cols if cols is not None else []:
            lo, hi = int(c.get("min")), int(c.get("max"))
            w = c.get("width")
            for i in range(lo, min(hi, lo + 16384) + 1):
                if w is not None:
                    sh.col_widths[i] = float(w)
                if c.get("style") not in (None, "0"):
                    sh.col_styles[i] = int(c.get("style"))
                if c.get("hidden") in ("1", "true"):
                    sh.hidden_cols.add(i)
        sd = root.find(M + "sheetData")
        for row in sd if sd is not None else []:
            rnum = int(row.get("r")) if row.get("r") else None
            if row.get("hidden") in ("1", "true") and rnum:
                sh.hidden_rows.add(rnum)
            if row.get("ht") and rnum:
                sh.row_heights[rnum] = float(row.get("ht"))
            if rnum and row.get("customFormat") in ("1", "true") and row.get("s"):
                sh.row_styles[rnum] = int(row.get("s"))
            next_col = 1
            for c in row.findall(M + "c"):
                ref = c.get("r")
                if ref:
                    r, col = parse_cell(ref)
                else:
                    r, col = rnum, next_col
                next_col = col + 1
                cell = Cell(sh.name, r, col, t=c.get("t"), style=int(c.get("s", "0")))
                v = c.find(M + "v")
                f = c.find(M + "f")
                self._last_sst = None
                cell.value = self._cell_value(cell.t, v, c)
                if self._last_sst is not None:
                    self.cell_sst[(sh.name, r, col)] = self._last_sst
                if f is not None:
                    kind = f.get("t", "normal")
                    cell.formula_kind = kind
                    text = f.text or ""
                    if kind == "shared":
                        si = f.get("si")
                        if text:
                            shared[si] = (cell, None, text)
                            cell.formula = text
                        else:
                            pending_shared.append((cell, si))
                    elif kind == "array":
                        cell.formula = text
                        cell.array_ref = f.get("ref")
                        if cell.array_ref:
                            a = cell.array_ref.split(":")
                            r1, c1 = parse_cell(a[0])
                            r2, c2 = parse_cell(a[-1])
                            array_refs.append((cell, (r1, c1, r2, c2)))
                    elif kind == "dataTable":
                        cell.formula = None
                        cell.parse_error = "データテーブル（What-If分析）"
                    else:
                        cell.formula = text
                sh.cells[(r, col)] = cell
                sh.max_row = max(sh.max_row, r)
                sh.max_col = max(sh.max_col, col)
        # 数式を構文木にする
        for cell in sh.cells.values():
            if cell.formula is not None and cell.formula_kind != "shared" or (cell.formula_kind == "shared" and cell.formula):
                try:
                    cell.ast = parse(cell.formula)
                except Unsupported as e:
                    cell.parse_error = f"未対応: {e}"
                except Exception as e:  # noqa: BLE001
                    cell.parse_error = f"構文エラー: {e}"
        for cell, si in pending_shared:
            master = shared.get(si)
            if master is None:
                cell.parse_error = "共有数式の親が見つかりません"
                continue
            m = master[0]
            cell.ast = m.ast
            cell.dr = cell.row - m.row
            cell.dc = cell.col - m.col
            cell.parse_error = m.parse_error
            if m.ast is not None:
                cell.formula = unparse(m.ast, cell.dr, cell.dc)
            else:
                cell.formula = m.formula
        for master, (r1, c1, r2, c2) in array_refs:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if (r, c) == (master.row, master.col):
                        continue
                    member = sh.cells.get((r, c))
                    if member is None:
                        member = Cell(sh.name, r, c)
                        sh.cells[(r, c)] = member
                    member.array_master = (master.row, master.col)
        mc = root.find(M + "mergeCells")
        for m in mc if mc is not None else []:
            a = m.get("ref").split(":")
            r1, c1 = parse_cell(a[0])
            r2, c2 = parse_cell(a[-1])
            sh.merges.append((r1, c1, r2, c2))
        dv = root.find(M + "dataValidations")
        for d in dv if dv is not None else []:
            sh.validations.append(_validation(d, M))
        # 新しいExcelは、別シートを参照するリストを extLst に入れる
        for d in root.iter(_q("x14:dataValidation")):
            sq = d.find(_q("xm:sqref"))
            f1 = d.find(_q("x14:formula1"))
            f2 = d.find(_q("x14:formula2"))
            v = _validation(d, "{%s}" % NS["x14"], sqref_text=sq.text if sq is not None else "")
            v.formula1 = f1.find(_q("xm:f")).text if f1 is not None and f1.find(_q("xm:f")) is not None else v.formula1
            v.formula2 = f2.find(_q("xm:f")).text if f2 is not None and f2.find(_q("xm:f")) is not None else v.formula2
            sh.validations.append(v)
        sh.protected = root.find(M + "sheetProtection") is not None and root.find(M + "sheetProtection").get("sheet") in ("1", "true")
        af = root.find(M + "autoFilter")
        if af is not None and af.get("ref"):
            rr = _sqref(af.get("ref"))
            sh.autofilter = rr[0] if rr else None
        sh.cond_formats = len(root.findall(M + "conditionalFormatting"))
        for cf in root.findall(M + "conditionalFormatting"):
            ranges = _sqref(cf.get("sqref", ""))
            for rule in cf.findall(M + "cfRule"):
                sh.cond_rules.append(CondRule(
                    ranges, rule.get("type", ""), int(rule.get("priority", "0")),
                    int(rule.get("dxfId")) if rule.get("dxfId") is not None else None,
                    [f.text or "" for f in rule.findall(M + "formula")], rule.get("operator"),
                    rule.get("stopIfTrue") in ("1", "true")))
        rels = self.rels(sh.path)
        for h in root.iter(M + "hyperlink"):
            rid = h.get(R + "id")
            target = rels.get(rid, ("", h.get("location", "")))[1] if rid else "#" + (h.get("location") or "")
            sh.hyperlinks.append((h.get("ref"), target))
        self._load_controls(sh, root, rels)
        hf = root.find(M + "headerFooter")
        for el in hf if hf is not None else []:
            t = _header_text(el.text or "")
            if t:
                sh.header_footer.append(t)
        for typ, target in rels.values():
            if typ == "comments" and self.zip.has(target):
                self._load_comments(sh, target)
            elif typ == "drawing" and self.zip.has(target):
                self._load_drawing(sh, target)
            elif typ == "vmlDrawing" and self.zip.has(target):
                vml = self.zip.read(target).decode("utf-8", errors="replace")
                if re.search(r'ObjectType="(Button|Checkbox|Drop|Radio|List|Spin|Scroll|GBox|Label|Edit)"', vml):
                    sh.has_vml_controls = True
            elif typ == "ctrlProp":
                sh.has_vml_controls = True
            elif typ == "table" and self.zip.has(target):
                t = self.xml(target)
                ref = t.get("ref", "A1").split(":")
                r1, c1 = parse_cell(ref[0])
                r2, c2 = parse_cell(ref[-1])
                cols = [_unxml(c.get("name", "")) for c in t.iter(M + "tableColumn")]
                name = t.get("displayName") or t.get("name") or ""
                self.tables[name.lower()] = Table(name, sh.name, r1, c1, r2, c2, int(t.get("headerRowCount", "1")),
                                                  int(t.get("totalsRowCount", "0")), cols)

    def _load_controls(self, sh: Sheet, root: ET.Element, rels: dict) -> None:
        """チェックボックスなどのフォームコントロールを、状態とリンク先のセルとともに読む。"""
        vml_path = None
        ld = root.find(M + "legacyDrawing")
        if ld is not None and ld.get(R + "id") in rels:
            vml_path = rels[ld.get(R + "id")][1]
        vml = {}
        if vml_path and self.zip.has(vml_path):
            vml = _vml_shapes(self.zip.read(vml_path).decode("utf-8", errors="replace"))
        seen = set()
        for ctl in root.iter(M + "control"):
            sid = ctl.get("shapeId", "")
            rid = ctl.get(R + "id")
            typ, part = rels.get(rid, ("", None)) if rid else ("", None)
            c = Control(kind="", name=ctl.get("name", ""), shape_id=sid)
            if typ == "ctrlProp" and part and self.zip.has(part):
                p = self.xml(part)
                c.prop_part = part
                c.kind = p.get("objectType", "")
                c.checked = p.get("checked") if p.get("checked") in ("Checked", "Mixed") else None
                c.link = p.get("fmlaLink")
                c.fmla_range = p.get("fmlaRange")
            elif typ:
                c.kind = "ActiveX"
                if part and self.zip.has(part):
                    m = re.search(rb'classid="\{([0-9A-Fa-f-]+)\}"', self.zip.read(part))
                    if m:
                        c.kind = "ActiveX " + _ACTIVEX.get(m[1].decode().upper(), "Control")
            pr = ctl.find(M + "controlPr")
            anc = pr.find(M + "anchor") if pr is not None else None
            if anc is not None:
                c.anchor = _xdr_pos(anc.find(M + "from"))
                c.anchor_to = _xdr_pos(anc.find(M + "to"))
                c.from_off = _xdr_row_off(anc.find(M + "from"))
                c.to_off = _xdr_row_off(anc.find(M + "to"))
            v = vml.get(sid)
            if v is not None:
                c.vml_part = vml_path
                c.text = v["text"]
                c.kind = c.kind or v["kind"]
                if c.prop_part is None:
                    c.checked = {"1": "Checked", "2": "Mixed"}.get(v["checked"] or "")
                c.link = c.link or v["link"]
                if c.anchor is None:
                    c.anchor, c.anchor_to, c.from_off, c.to_off = v["anchor"], v["anchor_to"], v["from_off"], v["to_off"]
            seen.add(sid)
            sh.controls.append(c)
        for sid, v in vml.items():
            if sid in seen or v["kind"] in ("Note", "") or v.get("dup"):
                continue
            sh.controls.append(Control(kind=v["kind"], name=v.get("name") or f"shape {sid}", shape_id=sid, anchor=v["anchor"],
                                       anchor_to=v["anchor_to"], link=v["link"], text=v["text"], vml_part=vml_path,
                                       from_off=v["from_off"], to_off=v["to_off"],
                                       checked={"1": "Checked", "2": "Mixed"}.get(v["checked"] or "")))
        for c in sh.controls:
            if c.link:
                c.link_cell = _link_cell(c.link, sh.name)

    def _cell_value(self, t, v, c):
        if t == "inlineStr":
            is_ = c.find(M + "is")
            return _rich_text(is_) if is_ is not None else ""
        if v is None or v.text is None:
            return "" if t == "str" and v is not None else None
        text = v.text
        if t == "s":
            i = int(text)
            if i in self.phonetic:
                self._last_sst = i
            return self.shared_strings[i] if i < len(self.shared_strings) else ""
        if t == "b":
            return text in ("1", "true")
        if t == "e":
            return ExcelError(text)
        if t in ("str", "d"):
            return text
        try:
            return float(text)
        except ValueError:
            return text

    def _load_comments(self, sh: Sheet, path: str) -> None:
        root = self.xml(path)
        cl = root.find(M + "commentList")
        for c in cl if cl is not None else []:
            r, col = parse_cell(c.get("ref"))
            t = c.find(M + "text")
            sh.comments[(r, col)] = _rich_text(t) if t is not None else ""

    def _load_drawing(self, sh: Sheet, path: str) -> None:
        root = self.xml(path)
        tops = []
        for top in root:
            if top.tag.endswith("}AlternateContent"):
                # 新しい形式の図形やフォームコントロールは mc:AlternateContent に包まれている
                choice = top.find(_q("mc:Choice"))
                tops.extend(list(choice) if choice is not None else [])
            else:
                tops.append(top)
        for anchor in tops:
            tag = anchor.tag.split("}")[-1]
            if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                continue
            if any(el.tag.endswith("}compatExt") for el in anchor.iter()):
                continue  # フォームコントロールの絵（コントロールとして別に読む）
            pos = None
            frm = anchor.find(_q("xdr:from"))
            if frm is not None:
                pos = (int(frm.find(_q("xdr:row")).text) + 1, int(frm.find(_q("xdr:col")).text) + 1)
            for kind, el in _drawing_items(anchor):
                nv = el.find(".//" + _q("xdr:cNvPr"))
                name = nv.get("name", "") if nv is not None else ""
                alt = (nv.get("descr") or nv.get("title") or "") if nv is not None else ""
                texts = []
                for p in el.iter(_q("a:p")):
                    line = "".join(t.text or "" for t in p.iter(_q("a:t")))
                    texts.append(line)
                text = "\n".join(texts).strip()
                sh.shapes.append(Shape(kind, name, text, alt, pos))

    # -------------------------------------------------------------- 名前
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


def _drawing_items(anchor: ET.Element):
    for el in anchor:
        tag = el.tag.split("}")[-1]
        if tag == "sp":
            yield ("text" if el.find(_q("xdr:txBody")) is not None else "shape"), el
        elif tag == "pic":
            yield "picture", el
        elif tag == "graphicFrame":
            yield "chart", el
        elif tag == "grpSp":
            # グループの中のグループも、たどって読む
            yield from _group_items(el)
        elif tag == "AlternateContent":
            # フォームコントロール（ボタンなど）は mc:AlternateContent に入る
            yield "control", el


def _group_items(grp: ET.Element):
    for sub in grp:
        st = sub.tag.split("}")[-1]
        if st == "sp":
            yield ("text" if sub.find(_q("xdr:txBody")) is not None else "shape"), sub
        elif st == "pic":
            yield "picture", sub
        elif st == "grpSp":
            yield from _group_items(sub)
        elif st == "graphicFrame":
            yield "chart", sub


def _phonetic_text(el: ET.Element) -> str | None:
    """ふりがな（rPh）を当てた読み。ふりがなが付いていない部分は、もとの文字のまま。"""
    rphs = el.findall(M + "rPh")
    if not rphs:
        return None
    base = _rich_text(el)
    out, pos = [], 0
    for rp in sorted(rphs, key=lambda x: int(x.get("sb", "0"))):
        sb, eb = int(rp.get("sb", "0")), int(rp.get("eb", "0"))
        t = rp.find(M + "t")
        out.append(base[pos:sb])
        out.append(t.text or "" if t is not None else base[sb:eb])
        pos = max(pos, eb)
    out.append(base[pos:])
    return "".join(out)


def _rich_text(el: ET.Element) -> str:
    # ふりがな（rPh）は除く
    out = []
    for child in el:
        tag = child.tag.split("}")[-1]
        if tag == "t":
            out.append(child.text or "")
        elif tag == "r":
            t = child.find(M + "t")
            if t is not None:
                out.append(t.text or "")
    return "".join(out)


# ActiveX コントロールの種類（classid）
_ACTIVEX = {
    "D7053240-CE69-11CD-A777-00DD01143C57": "CommandButton",
    "8BD21D40-EC42-11CE-9E0D-00AA006002F3": "CheckBox",
    "8BD21D50-EC42-11CE-9E0D-00AA006002F3": "OptionButton",
    "8BD21D10-EC42-11CE-9E0D-00AA006002F3": "TextBox",
    "8BD21D30-EC42-11CE-9E0D-00AA006002F3": "ComboBox",
    "8BD21D20-EC42-11CE-9E0D-00AA006002F3": "ListBox",
    "8BD21D60-EC42-11CE-9E0D-00AA006002F3": "ToggleButton",
    "978C9E23-D4B0-11CE-BF2D-00AA003F40D0": "Label",
    "79176FB0-B7F2-11CE-97EF-00AA006D2776": "SpinButton",
    "DFD181E0-5E2F-11CE-A449-00AA004A803D": "ScrollBar",
    "4C599241-6926-101B-9992-00000B65C6F9": "Image",
}

# 既定の色（Officeの標準テーマと、旧来の56色パレットのうち白黒）
_THEME = {0: "FFFFFF", 1: "000000", 2: "E7E6E6", 3: "44546A"}
_INDEXED = {8: "000000", 9: "FFFFFF", 64: "000000", 65: "FFFFFF"}


def _color(el: ET.Element | None) -> str | None:
    if el is None:
        return None
    if el.get("rgb"):
        return el.get("rgb")[-6:].upper()
    if el.get("theme") is not None and el.get("tint") in (None, "0", "0.0"):
        return _THEME.get(int(el.get("theme")))
    if el.get("indexed") is not None:
        return _INDEXED.get(int(el.get("indexed")))
    return None


def _fill_color(f: ET.Element | None, dxf: bool = False) -> str | None:
    if f is None:
        return None
    pf = f.find(M + "patternFill")
    if pf is None:
        return None
    if not dxf and pf.get("patternType") in (None, "none"):
        return None
    # セルの塗りは fgColor、条件付き書式（dxf）の塗りは bgColor に書かれる
    col = pf.find(M + ("bgColor" if dxf else "fgColor"))
    return _color(col)


def _sqref(text: str) -> list[tuple[int, int, int, int]]:
    out = []
    for part in text.split():
        a = part.split(":")
        try:
            r1, c1 = parse_cell(a[0])
            r2, c2 = parse_cell(a[-1])
        except ValueError:
            continue
        out.append((r1, c1, r2, c2))
    return out


def _validation(d: ET.Element, ns: str, sqref_text: str | None = None) -> Validation:
    sq = sqref_text if sqref_text is not None else d.get("sqref", "")
    ranges = []
    for part in sq.split():
        a = part.split(":")
        try:
            r1, c1 = parse_cell(a[0])
            r2, c2 = parse_cell(a[-1])
        except ValueError:
            continue
        ranges.append((r1, c1, r2, c2))
    f1 = d.find(ns + "formula1")
    f2 = d.find(ns + "formula2")
    return Validation(
        sqref=ranges,
        sqref_text=sq,
        type=d.get("type", "any"),
        operator=d.get("operator"),
        formula1=f1.text if f1 is not None and len(f1) == 0 else None,
        formula2=f2.text if f2 is not None and len(f2) == 0 else None,
        allow_blank=d.get("allowBlank") in ("1", "true"),
        prompt=d.get("prompt"),
        prompt_title=d.get("promptTitle"),
        error=d.get("error"),
        error_title=d.get("errorTitle"),
        error_style=d.get("errorStyle", "stop"),
        show_error=d.get("showErrorMessage") in ("1", "true"),
        show_input=d.get("showInputMessage") in ("1", "true"),
    )


def _xdr_pos(el: ET.Element | None) -> tuple[int, int] | None:
    if el is None:
        return None
    col = el.find(_q("xdr:col"))
    row = el.find(_q("xdr:row"))
    if col is None or row is None:
        return None
    return int(row.text) + 1, int(col.text) + 1


def _external_cache(root: ET.Element) -> dict[str, dict[tuple[int, int], object]]:
    """外部ブックへのリンク（externalLink）に保存されている、参照先の値。"""
    out: dict[str, dict[tuple[int, int], object]] = {}
    book = root.find(M + "externalBook")
    if book is None:
        return out
    sn = book.find(M + "sheetNames")
    names = [x.get("val", "") for x in sn] if sn is not None else []
    ds = book.find(M + "sheetDataSet")
    for sd in ds if ds is not None else []:
        sid = int(sd.get("sheetId", "0"))
        if sid >= len(names):
            continue
        cells: dict[tuple[int, int], object] = {}
        for row in sd.findall(M + "row"):
            for cell in row.findall(M + "cell"):
                v = cell.find(M + "v")
                if v is None or v.text is None or not cell.get("r"):
                    continue
                t = cell.get("t", "n")
                val: object
                if t == "b":
                    val = v.text in ("1", "true")
                elif t == "e":
                    val = ExcelError(v.text)
                elif t in ("s", "str", "inlineStr"):
                    val = v.text
                else:
                    try:
                        val = float(v.text)
                    except ValueError:
                        val = v.text
                cells[parse_cell(cell.get("r"))] = val
        out[names[sid].lower()] = cells
    for name in names:
        out.setdefault(name.lower(), {})
    return out


def _unxml(s: str) -> str:
    """_x000a_ のような書き方を、もとの文字に戻す。"""
    return re.sub(r"_x([0-9A-Fa-f]{4})_", lambda m: chr(int(m[1], 16)), s)


def _xdr_row_off(el: ET.Element | None) -> float:
    off = el.find(_q("xdr:rowOff")) if el is not None else None
    return int(off.text) / 12700 if off is not None and off.text else 0.0


def _vml_text(s: str) -> str:
    import html

    t = re.sub(r"<br\s*/?>", "\n", s, flags=re.I)
    t = re.sub(r"</(div|p)>", "\n", t, flags=re.I)
    t = re.sub(r"<[^>]+>", "", t)
    return html.unescape(t).strip()


def _vml_shapes(text: str) -> dict[str, dict]:
    """VML の図形（フォームコントロールとメモ）を、図形の番号ごとに取り出す。VMLは厳密なXMLとは限らないので正規表現で読む。"""
    out: dict[str, dict] = {}
    for m in re.finditer(r"<v:shape\b([^>]*)>(.*?)</v:shape>", text, re.S):
        attrs, body = m[1], m[2]
        ids = re.findall(r'\b(?:id|o:spid)="_x0000_s(\d+)"', attrs)
        cd = re.search(r'<x:ClientData\b[^>]*ObjectType="(\w+)"[^>]*>(.*?)</x:ClientData>', body, re.S)
        if not ids or not cd:
            continue

        def tag(n: str, _b=cd[2]):
            t = re.search(r"<x:" + n + r">(.*?)</x:" + n + ">", _b, re.S)
            return t[1].strip() if t else None

        pos = pos_to = None
        off = off_to = 0.0
        a = tag("Anchor")
        if a:
            nums = [int(x) for x in re.findall(r"\d+", a)]
            if len(nums) >= 8:
                # 左の列, ずれ(px), 上の行, ずれ(px), 右の列, ずれ, 下の行, ずれ
                pos, pos_to = (nums[2] + 1, nums[0] + 1), (nums[6] + 1, nums[4] + 1)
                off, off_to = nums[3] * 0.75, nums[7] * 0.75
        tb = re.search(r"<v:textbox\b[^>]*>(.*?)</v:textbox>", body, re.S)
        d = {"kind": {"Checkbox": "CheckBox", "Edit": "EditBox"}.get(cd[1], cd[1]), "checked": tag("Checked"),
             "link": tag("FmlaLink"), "anchor": pos, "anchor_to": pos_to, "from_off": off, "to_off": off_to,
             "text": _vml_text(tb[1]) if tb else ""}
        for i, sid in enumerate(dict.fromkeys(ids)):
            out[sid] = dict(d, dup=i > 0)
    return out


_LINK = re.compile(r"^(?:(?:'((?:[^']|'')+)'|([^'!]+))!)?\$?([A-Za-z]{1,3})\$?(\d+)$")


def _link_cell(link: str, here: str) -> tuple[str, int, int] | None:
    m = _LINK.match(link.strip().lstrip("="))
    if not m:
        return None
    sheet = (m[1].replace("''", "'") if m[1] else m[2]) or here
    try:
        r, c = parse_cell(m[3].upper() + m[4])
    except ValueError:
        return None
    return sheet, r, c


def _area_ranges(text: str) -> list[tuple[int, int, int, int]]:
    """印刷範囲（'シート'!$A$1:$AG$54 や $A:$AL、$1:$5）を範囲の一覧にする。"""
    from ..formula.parser import col_to_num

    out = []
    for part in re.split(r",(?=(?:[^']*'[^']*')*[^']*$)", text):
        ref = part.rsplit("!", 1)[-1].replace("$", "").strip()
        a, b = ref.split(":", 1) if ":" in ref else (ref, ref)
        corners = []
        for x in (a, b):
            m = re.fullmatch(r"([A-Za-z]*)(\d*)", x.strip())
            if not m or not x.strip():
                break
            corners.append((int(m[2]) if m[2] else None, col_to_num(m[1].upper()) if m[1] else None))
        if len(corners) != 2:
            continue
        (r1, c1), (r2, c2) = corners
        out.append((r1 or 1, c1 or 1, r2 or 1048576, c2 or 16384))
    return out


def _header_text(s: str) -> str:
    """ヘッダー・フッターの書式コード（&L &C &P など）を除いて、読める文字にする。"""
    t = s.replace("&&", "\x00")
    t = re.sub(r'&"[^"]*"', "", t)
    t = re.sub(r"&\d+", "", t)
    t = re.sub(r"&[LCR]", " ", t)
    for code, word in (("&P", "（ページ番号）"), ("&N", "（総ページ数）"), ("&D", "（日付）"), ("&T", "（時刻）"),
                       ("&F", "（ファイル名）"), ("&A", "（シート名）"), ("&Z", "（フォルダ）"), ("&G", "（画像）")):
        t = t.replace(code, word)
    t = re.sub(r"&[BIUESXYKO]", "", t)
    return re.sub(r"\s+", " ", t.replace("\x00", "&")).strip()


def load(path: str) -> Workbook:
    with open(path, "rb") as f:
        return Workbook(f.read(), filename=path)
