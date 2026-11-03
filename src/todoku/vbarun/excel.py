"""マクロから見える Excel：Application・Workbook・Worksheet・Range（todoku の Workbook の上に作る）。

セルの値：マクロが書いた値を重ね、数式のセルは、その値で計算し直した値（todoku の計算式の仕組み）を返す。
書いたセルは changed に残し、あとで「マクロが変えたセル」として、証明つきでファイルに書く。
書式・セル結合・シートの追加と削除・数式の書き込みなど、ここで再現しないものは Unsupported を投げる。
"""

from __future__ import annotations

import datetime as dt

from ..formula.parser import addr, col_to_num, num_to_col
from .interp import (EMPTY, MISSING, NOTHING, NeedsInput, Unsupported, VArray, VbaError, VDate, to_bool, to_double, to_long,
                     to_str)

MAX_ROW, MAX_COL = 1_048_576, 16_384
XL_UP, XL_DOWN, XL_LEFT, XL_RIGHT = -4162, -4121, -4159, 4161


def _empty_cell(v) -> bool:
    return v is None or v is EMPTY or v == ""


class Host:
    def __init__(self, wb, overrides: dict | None = None, now: dt.datetime | None = None):
        from ..xlsx.calc import Calc

        self.wb = wb
        self.over: dict = dict(overrides or {})
        self.changed: dict = {}  # マクロが書いたセル（シート, 行, 列）→ 値
        self._calc = None
        self._Calc = Calc
        self.rt = None
        self.volatile = False
        self._now = now
        self.events_enabled = True
        self.screen_updating = True
        self.display_alerts = True
        self.calculation = -4105
        self.status_bar = False
        self.printed: list[str] = []
        self.protected: dict[str, bool] = {sh.name: bool(sh.protected) for sh in wb.sheets}
        visible = [sh for sh in wb.sheets if sh.state == "visible" and not sh.state.startswith("non-")]
        self.active = (visible or wb.sheets)[0].name
        self.selection: Range | None = None
        self.clipboard: Range | None = None
        self.depth = 0  # 入れ子になったイベントの深さ
        # Find の設定（Excel は、指定のない設定に前回の Find の設定を使う）。はじめの値は Excel の既定
        self.find_state: dict = {"lookin": -4123, "lookat": 2, "searchorder": 1, "matchcase": False}
        self.app = Application(self)
        self.book = Workbook(self)

    # ---------------------------------------------------------------- セル
    def calc(self):
        if self._calc is None:
            self._calc = self._Calc(self.wb, self.over)
        return self._calc

    def sheet(self, name: str):
        for sh in self.wb.sheets:
            if sh.name.lower() == name.lower():
                return sh
        raise VbaError(9)

    def get(self, sheet: str, r: int, c: int):
        key = (sheet, r, c)
        if key in self.over:
            v = self.over[key]
            return EMPTY if v is None else v
        sh = self.sheet(sheet)
        cell = sh.cells.get((r, c))
        if cell is None:
            return EMPTY
        m = sh.merged_at(r, c)
        if m and (r, c) != (m[0], m[1]):
            return EMPTY
        if cell.has_formula:
            v = self.calc().cell_value(sheet, r, c)
        else:
            v = cell.value
        return _to_vba(v, self.wb, cell)

    def set(self, sheet: str, r: int, c: int, v, by_macro: bool = True) -> None:
        sh = self.sheet(sheet)
        if self.protected.get(sh.name) and self.wb.style(sh.style_of(r, c)).locked:
            raise VbaError(1004, "変更しようとしているセルは、保護されたシートにあります")
        m = sh.merged_at(r, c)
        if m and (r, c) != (m[0], m[1]):
            r, c = m[0], m[1]  # 結合セルの途中に書いた値は、左上のセルに入る
        val = _from_vba(v, self.wb)
        key = (sh.name, r, c)
        self.over[key] = val
        if by_macro:
            self.changed[key] = val
        self._calc = None

    def now(self) -> dt.datetime:
        self.volatile = True
        return self._now or dt.datetime.now()

    # ---------------------------------------------------------------- 名前
    def global_object(self, name: str):
        n = name.lower()
        if n == "application":
            return self.app
        if n in ("activesheet",):
            return Worksheet(self, self.active)
        if n in ("activeworkbook", "thisworkbook"):
            return self.book
        if n in ("worksheets", "sheets"):
            return Sheets(self)
        if n == "range":
            return _Callable(lambda args, named: Worksheet(self, self.active).range(args))
        if n == "cells":
            return _Callable(lambda args, named: Worksheet(self, self.active).cells(args), default=lambda: Worksheet(self, self.active).cells([]))
        if n in ("rows", "columns"):
            ws = Worksheet(self, self.active)
            return _Callable(lambda args, named: ws.vba_get(n, args, named), default=lambda: ws.vba_get(n, [], {}))
        if n == "activecell":
            return self.selection.cell(1, 1) if self.selection else Worksheet(self, self.active).cells([1, 1])
        if n == "selection":
            return self.selection or Worksheet(self, self.active).cells([1, 1])
        if n == "worksheetfunction":
            return WorksheetFunction(self)
        if n == "activewindow":
            return Window(self)
        if n == "debug":
            from .interp import _Debug

            return _Debug()
        if n == "workbooks":
            return _Callable(lambda args, named: self.book)
        for sh in self.wb.sheets:
            if sh.code_name and sh.code_name.lower() == n:
                return Worksheet(self, sh.name)
        if n in ("intersect", "union", "inputbox", "run", "calculate"):
            return _Callable(lambda args, named: self.app.vba_get(n, args, named))
        return None

    def me(self, module: str):
        if module.lower() == "thisworkbook":
            return self.book
        for sh in self.wb.sheets:
            if sh.code_name and sh.code_name.lower() == module.lower():
                return Worksheet(self, sh.name)
        raise VbaError(91)

    def new_object(self, cls: str):
        from .interp import Collection, Dictionary

        c = (cls or "").lower()
        if c == "collection":
            return Collection()
        if c in ("dictionary", "scripting.dictionary"):
            return Dictionary()
        raise Unsupported(f"New {cls}（クラスモジュールの部品）")

    def create_object(self, cls: str, node, fr):
        raise Unsupported(f"CreateObject(\"{cls}\")（外の部品）", node.line, fr.module)

    def iterate(self, obj) -> list:
        if isinstance(obj, Range):
            return [obj.cell(i, j) for i in range(1, obj.rows + 1) for j in range(1, obj.cols + 1)]
        if isinstance(obj, Sheets):
            return [Worksheet(self, sh.name) for sh in self.wb.sheets if not sh.state.startswith("non-")]
        if isinstance(obj, _Lines):
            return [obj.item(i) for i in range(1, obj.count + 1)]
        raise VbaError(438)

    def type_of(self, obj, cls: str) -> bool:
        return getattr(obj, "type_name", "").lower() == cls.split(".")[-1].lower()

    # ---------------------------------------------------------------- イベント
    def fire_change(self, rng: "Range") -> None:
        """セルの値が変わったときのイベント（Worksheet_Change と Workbook_SheetChange）。"""
        if not self.events_enabled or self.rt is None:
            return
        if self.depth > 50:
            raise VbaError(28, "スタック領域が不足しています（Change イベントが繰り返し呼ばれています）")
        sh = self.sheet(rng.sheet)
        self.depth += 1
        try:
            if sh.code_name:
                p = self.rt.find_proc(f"{sh.code_name}.worksheet_change")
                if p is not None:
                    from .interp import Ref

                    self.rt.call_proc(p, [Ref(rng)])
            p = self.rt.find_proc("thisworkbook.workbook_sheetchange")
            if p is not None:
                from .interp import Ref

                self.rt.call_proc(p, [Ref(Worksheet(self, rng.sheet)), Ref(rng)])
        finally:
            self.depth -= 1


def _to_vba(v, wb, cell):
    from ..formula.numfmt import is_date_format
    from ..formula.values import ExcelError

    if v is None:
        return EMPTY
    if isinstance(v, ExcelError):
        return v
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        fmt = wb.style(cell.style).num_fmt if cell is not None else "General"
        if is_date_format(fmt):
            # VBA の日付は 1900 年の数え方のシリアル値。1904 年の数え方のブックは 1462 日ずらす（ECMA-376 18.2.28 date1904）
            return VDate(float(v) + (1462 if getattr(wb, "date1904", False) else 0))
        return int(v) if float(v) == int(v) and abs(v) < 2**31 else float(v)
    return v


def _from_vba(v, wb=None):
    from ..formula.values import ExcelError

    if v is EMPTY or v is None:
        return None
    if isinstance(v, VDate):
        # 1904 年の数え方のブックのセルには、1462 日ずらしたシリアル値が入る
        return float(v) - (1462 if getattr(wb, "date1904", False) else 0)
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return float(v)
    if isinstance(v, (float, str, ExcelError)):
        return v
    raise VbaError(13)


class _Callable:
    """Range(…)・Cells(…) のように、呼んで使う名前。"""

    def __init__(self, f, default=None):
        self.f = f
        self.default = default

    def vba_call(self, args, named):
        return self.f(args, named)

    def vba_get(self, name, args, named):
        obj = self.default() if self.default else self.f([], {})
        return obj.vba_get(name, args, named)

    def vba_default(self):
        obj = self.default() if self.default else self.f([], {})
        return obj.vba_default() if hasattr(obj, "vba_default") else obj


# ====================================================================== Range
class Range:
    type_name = "Range"

    def __init__(self, host: Host, sheet: str, r1: int, c1: int, r2: int | None = None, c2: int | None = None):
        if r1 < 1 or c1 < 1 or (r2 or r1) > MAX_ROW or (c2 or c1) > MAX_COL:
            raise VbaError(1004, "範囲がシートの外です")
        self.host = host
        self.sheet = sheet
        self.r1, self.c1 = r1, c1
        self.r2, self.c2 = (r2 if r2 is not None else r1), (c2 if c2 is not None else c1)

    @property
    def rows(self) -> int:
        return self.r2 - self.r1 + 1

    @property
    def cols(self) -> int:
        return self.c2 - self.c1 + 1

    def same(self, other) -> bool:
        return isinstance(other, Range) and (self.sheet, self.r1, self.c1, self.r2, self.c2) == \
            (other.sheet, other.r1, other.c1, other.r2, other.c2)

    def cell(self, i: int, j: int) -> "Range":
        return Range(self.host, self.sheet, self.r1 + i - 1, self.c1 + j - 1)

    def address(self, row_abs=True, col_abs=True) -> str:
        def one(r, c):
            return ("$" if col_abs else "") + num_to_col(c) + ("$" if row_abs else "") + str(r)

        a = one(self.r1, self.c1)
        return a if (self.r1, self.c1) == (self.r2, self.c2) else f"{a}:{one(self.r2, self.c2)}"

    # ---------------------------------------------------------------- 値
    def get_value(self):
        if self.rows == 1 and self.cols == 1:
            return self.host.get(self.sheet, self.r1, self.c1)
        if self.rows * self.cols > 2_000_000:
            raise Unsupported("とても大きな範囲の値の読み込み")
        out = VArray([1, 1], [self.rows, self.cols])
        for i in range(self.rows):
            for j in range(self.cols):
                out.data[(i + 1, j + 1)] = self.host.get(self.sheet, self.r1 + i, self.c1 + j)
        return out

    def set_value(self, v) -> None:
        cells = [(self.r1 + i, self.c1 + j) for i in range(self.rows) for j in range(self.cols)]
        if isinstance(v, VArray):
            if len(v.lbounds) == 2:
                lo_r, lo_c = v.lbounds
                for i in range(self.rows):
                    for j in range(self.cols):
                        k = (lo_r + i, lo_c + j)
                        x = v.data.get(k, EMPTY) if i <= v.ubounds[0] - lo_r and j <= v.ubounds[1] - lo_c else _NA()
                        self.host.set(self.sheet, self.r1 + i, self.c1 + j, x)
            elif len(v.lbounds) == 1:
                lo = v.lbounds[0]
                for i in range(self.rows):
                    for j in range(self.cols):
                        x = v.data.get((lo + j,), EMPTY) if j <= v.ubounds[0] - lo else _NA()
                        self.host.set(self.sheet, self.r1 + i, self.c1 + j, x)
            else:
                raise VbaError(13)
        else:
            if hasattr(v, "vba_default"):
                v = v.vba_default()
            for r, c in cells:
                self.host.set(self.sheet, r, c, v)
        self.host.fire_change(self)

    def vba_default(self):
        return self.get_value()

    def text(self) -> str:
        from ..formula.numfmt import format_value
        from ..xlsx.window import display

        v = self.host.get(self.sheet, self.r1, self.c1)
        sh = self.host.sheet(self.sheet)
        cell = sh.cells.get((self.r1, self.c1))
        key = (sh.name, self.r1, self.c1)
        if key not in self.host.over and cell is not None and not cell.has_formula:
            return display(self.host.wb, cell, flag=False)
        fmt = self.host.wb.style(sh.style_of(self.r1, self.c1)).num_fmt
        raw = float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else (None if v is EMPTY else v)
        if raw is None:
            return ""
        if isinstance(raw, str):
            return raw
        if isinstance(raw, bool):
            return "TRUE" if raw else "FALSE"
        if isinstance(v, VDate) and self.host.wb.date1904:
            raw -= 1462  # VBA の日付（1900 年の数え方）を、ブックの数え方のシリアル値に
        return format_value(raw, fmt, self.host.wb.date1904)

    # ---------------------------------------------------------------- メンバー
    def vba_get(self, name: str, args: list, named: dict):
        n = name
        h = self.host
        if n in ("value", "value2"):
            return self.get_value()
        if n == "text":
            return self.text()
        if n in ("formula", "formulalocal", "formular1c1"):
            sh = h.sheet(self.sheet)
            cell = sh.cells.get((self.r1, self.c1))
            key = (sh.name, self.r1, self.c1)
            if key not in h.over and cell is not None and cell.formula:
                if n == "formular1c1":
                    raise Unsupported("Formula R1C1 の読み込み")
                return "=" + cell.formula
            v = self.get_value()
            return "" if v is EMPTY else to_str(v)
        if n == "hasformula":
            sh = h.sheet(self.sheet)
            return any((sh.cells.get((r, c)) is not None and sh.cells[(r, c)].has_formula and (sh.name, r, c) not in h.over)
                       for r in range(self.r1, self.r2 + 1) for c in range(self.c1, self.c2 + 1))
        if n == "row":
            return self.r1
        if n == "column":
            return self.c1
        if n in ("count", "countlarge"):
            return self.rows * self.cols
        if n == "rows":
            lines = _Lines(self, "rows")
            return lines.item(to_long(args[0])) if args else lines
        if n == "columns":
            lines = _Lines(self, "columns")
            return lines.item(_col_arg(args[0])) if args else lines
        if n == "address":
            ra = to_bool(named.get("rowabsolute", args[0] if len(args) > 0 and args[0] is not MISSING else True))
            ca = to_bool(named.get("columnabsolute", args[1] if len(args) > 1 and args[1] is not MISSING else True))
            return self.address(ra, ca)
        if n == "offset":
            dr = to_long(named.get("rowoffset", args[0] if len(args) > 0 and args[0] is not MISSING else 0))
            dc = to_long(named.get("columnoffset", args[1] if len(args) > 1 and args[1] is not MISSING else 0))
            return Range(h, self.sheet, self.r1 + dr, self.c1 + dc, self.r2 + dr, self.c2 + dc)
        if n == "resize":
            nr = named.get("rowsize", args[0] if len(args) > 0 and args[0] is not MISSING else MISSING)
            nc = named.get("columnsize", args[1] if len(args) > 1 and args[1] is not MISSING else MISSING)
            r = to_long(nr) if nr is not MISSING else self.rows
            c = to_long(nc) if nc is not MISSING else self.cols
            if r < 1 or c < 1:
                raise VbaError(1004)
            return Range(h, self.sheet, self.r1, self.c1, self.r1 + r - 1, self.c1 + c - 1)
        if n in ("cells", "item"):
            if not args:
                return self
            if len(args) == 1:
                k = to_long(args[0]) - 1
                return self.cell(k // self.cols + 1, k % self.cols + 1)
            return self.cell(to_long(args[0]), _col_arg(args[1]))
        if n == "range":
            sub = Worksheet(h, self.sheet).range(args)
            return Range(h, self.sheet, self.r1 + sub.r1 - 1, self.c1 + sub.c1 - 1, self.r1 + sub.r2 - 1, self.c1 + sub.c2 - 1)
        if n == "end":
            return self.end(to_long(args[0]))
        if n == "entirerow":
            return Range(h, self.sheet, self.r1, 1, self.r2, MAX_COL)
        if n == "entirecolumn":
            return Range(h, self.sheet, 1, self.c1, MAX_ROW, self.c2)
        if n == "currentregion":
            return self.current_region()
        if n in ("clearcontents", "clear"):
            self.clear()
            if n == "clear":
                h.rt.messages.append("（Clear の書式の消去は再現していません。値だけを消しました）")
            return EMPTY
        if n == "copy":
            dest = named.get("destination", args[0] if args and args[0] is not MISSING else MISSING)
            if dest is MISSING:
                h.clipboard = self
                return EMPTY
            self.copy_to(dest)
            return EMPTY
        if n == "pastespecial":
            if h.clipboard is None:
                raise VbaError(1004)
            paste = named.get("paste", args[0] if args else -4104)
            if paste not in (-4163, -4104, MISSING):
                raise Unsupported("PasteSpecial の、値と数式以外の貼り付け")
            h.clipboard.copy_to(self)
            return EMPTY
        if n in ("select", "activate"):
            h.active = self.sheet
            h.selection = self
            return EMPTY
        if n in ("worksheet", "parent"):
            return Worksheet(h, self.sheet)
        if n == "find":
            return self.find(args, named)
        if n in ("findnext", "findprevious"):
            st = h.find_state
            if "what" not in st:
                raise VbaError(91)
            after = args[0] if args and args[0] is not MISSING else named.get("after", self)
            prev = (n == "findprevious") != bool(st.get("prev")) if n == "findprevious" else bool(st.get("prev"))
            return st["range"].find([st["what"]], {}, after_default=after, direction_override=prev)
        if n == "mergecells":
            sh = h.sheet(self.sheet)
            return sh.merged_at(self.r1, self.c1) is not None
        if n == "mergearea":
            m = h.sheet(self.sheet).merged_at(self.r1, self.c1)
            return Range(h, self.sheet, *m) if m else self
        if n == "numberformat":
            sh = h.sheet(self.sheet)
            return h.wb.style(sh.style_of(self.r1, self.c1)).num_fmt
        if n == "locked":
            sh = h.sheet(self.sheet)
            return h.wb.style(sh.style_of(self.r1, self.c1)).locked
        if n in ("hidden",):
            sh = h.sheet(self.sheet)
            return self.r1 in sh.hidden_rows or self.c1 in sh.hidden_cols
        if n == "areas":
            return _Areas([self])
        if n in ("calculate",):
            return EMPTY
        raise Unsupported(f"Range の「{name}」")

    def vba_set(self, name: str, v) -> None:
        if name in ("value", "value2"):
            self.set_value(v)
            return
        if name in ("formula", "formulalocal"):
            s = to_str(v) if not isinstance(v, VArray) else None
            if s is not None and not s.startswith("="):
                self.set_value(v)  # 「=」で始まらない式は、値として入る
                return
            raise Unsupported("マクロからの数式の書き込み")
        raise Unsupported(f"Range の「{name}」への書き込み（書式などは再現しません）")

    # ---------------------------------------------------------------- 操作
    def clear(self) -> None:
        sh = self.host.sheet(self.sheet)
        r2 = min(self.r2, max(sh.max_row, max((k[1] for k in self.host.over if k[0] == sh.name), default=0)))
        c2 = min(self.c2, max(sh.max_col, max((k[2] for k in self.host.over if k[0] == sh.name), default=0)))
        for r in range(self.r1, r2 + 1):
            for c in range(self.c1, c2 + 1):
                cell = sh.cells.get((r, c))
                if (cell is not None and (cell.value is not None or cell.has_formula)) or (sh.name, r, c) in self.host.over:
                    if cell is not None and cell.has_formula and (sh.name, r, c) not in self.host.over:
                        raise Unsupported("数式のあるセルの消去（数式の削除は再現しません）")
                    self.host.set(sh.name, r, c, EMPTY)
        self.host.fire_change(self)

    def copy_to(self, dest) -> None:
        if not isinstance(dest, Range):
            raise VbaError(424)
        sh = self.host.sheet(self.sheet)
        for i in range(self.rows):
            for j in range(self.cols):
                cell = sh.cells.get((self.r1 + i, self.c1 + j))
                if cell is not None and cell.has_formula and (sh.name, self.r1 + i, self.c1 + j) not in self.host.over:
                    raise Unsupported("数式のあるセルのコピー（数式の書き込みは再現しません）")
                self.host.set(dest.sheet, dest.r1 + i, dest.c1 + j, self.host.get(self.sheet, self.r1 + i, self.c1 + j))
        self.host.fire_change(Range(self.host, dest.sheet, dest.r1, dest.c1, dest.r1 + self.rows - 1, dest.c1 + self.cols - 1))

    def end(self, direction: int) -> "Range":
        """Ctrl＋矢印と同じ動き。値のあるセルが続くならその端、続かないなら次に値のあるセル、なければシートの端。
        値のある範囲の外は、1つずつ見ずに飛ぶ（「Cells(Rows.Count, 1).End(xlUp)」が速いように）。"""
        dr, dc = {XL_UP: (-1, 0), XL_DOWN: (1, 0), XL_LEFT: (0, -1), XL_RIGHT: (0, 1)}.get(direction, (None, None))
        if dr is None:
            raise VbaError(5)
        h = self.host
        sh = h.sheet(self.sheet)
        used_r = max([sh.max_row] + [k[1] for k, v in h.over.items() if k[0] == sh.name and v is not None])
        used_c = max([sh.max_col] + [k[2] for k, v in h.over.items() if k[0] == sh.name and v is not None])

        def filled(rr, cc):
            if rr > used_r or cc > used_c:
                return False
            return not _empty_cell(h.get(self.sheet, rr, cc))

        def inside(rr, cc):
            return 1 <= rr <= MAX_ROW and 1 <= cc <= MAX_COL

        r, c = self.r1, self.c1
        nr, nc = r + dr, c + dc
        if not inside(nr, nc):
            return self
        if filled(r, c) and filled(nr, nc):
            while inside(nr + dr, nc + dc) and filled(nr + dr, nc + dc):
                nr, nc = nr + dr, nc + dc
            return Range(h, self.sheet, nr, nc)
        while inside(nr, nc):
            if dr > 0 and nr > used_r:
                return Range(h, self.sheet, MAX_ROW, nc)
            if dc > 0 and nc > used_c:
                return Range(h, self.sheet, nr, MAX_COL)
            if dr < 0 and nr > used_r:
                nr = used_r
                continue
            if dc < 0 and nc > used_c:
                nc = used_c
                continue
            if filled(nr, nc):
                return Range(h, self.sheet, nr, nc)
            nr, nc = nr + dr, nc + dc
        return Range(h, self.sheet, max(1, min(MAX_ROW, nr - dr)), max(1, min(MAX_COL, nc - dc)))

    def current_region(self) -> "Range":
        h = self.host
        r1, c1, r2, c2 = self.r1, self.c1, self.r2, self.c2

        def filled(rr, cc):
            return 1 <= rr <= MAX_ROW and 1 <= cc <= MAX_COL and not _empty_cell(h.get(self.sheet, rr, cc))

        changed = True
        while changed:
            changed = False
            if any(filled(r1 - 1, c) for c in range(c1 - 1, c2 + 2)):
                r1 -= 1
                changed = True
            if any(filled(r2 + 1, c) for c in range(c1 - 1, c2 + 2)):
                r2 += 1
                changed = True
            if any(filled(r, c1 - 1) for r in range(r1 - 1, r2 + 2)):
                c1 -= 1
                changed = True
            if any(filled(r, c2 + 1) for r in range(r1 - 1, r2 + 2)):
                c2 += 1
                changed = True
        return Range(h, self.sheet, max(1, r1), max(1, c1), r2, c2)

    def find(self, args, named, after_default=None, direction_override=None):
        """Range.Find（Excel の規則）：After のセルの次から探し、範囲の端で折り返して、最後に After を見る。
        LookIn・LookAt・SearchOrder・MatchCase は、指定がなければ前の Find の設定（はじめは xlFormulas・xlPart・
        xlByRows・大文字小文字を区別しない）。"""
        h = self.host
        names = ["what", "after", "lookin", "lookat", "searchorder", "searchdirection", "matchcase", "matchbyte"]
        opt = {}
        for i, n in enumerate(names):
            if n in named:
                opt[n] = named[n]
            elif i < len(args) and args[i] is not MISSING:
                opt[n] = args[i]
        if "what" not in opt:
            raise VbaError(5)
        st = h.find_state
        for k in ("lookin", "lookat", "searchorder", "matchcase"):
            if k in opt:
                st[k] = opt[k]
        lookin, lookat = to_long(st["lookin"]), to_long(st["lookat"])
        by_cols = to_long(st["searchorder"]) == 2
        match_case = to_bool(st["matchcase"])
        prev = direction_override if direction_override is not None else to_long(opt.get("searchdirection", 1)) == 2
        st["what"], st["range"], st["prev"] = opt["what"], self, prev
        what = to_str(opt["what"]) if not isinstance(opt["what"], (int, float)) or isinstance(opt["what"], bool) else to_str(opt["what"])
        after = opt.get("after", after_default)
        sh = h.sheet(self.sheet)
        used_r = max([sh.max_row] + [k[1] for k, v in h.over.items() if k[0] == sh.name and v is not None])
        used_c = max([sh.max_col] + [k[2] for k, v in h.over.items() if k[0] == sh.name and v is not None])
        r2, c2 = min(self.r2, used_r), min(self.c2, used_c)
        if r2 < self.r1 or c2 < self.c1:
            return NOTHING
        cells = ([(r, c) for c in range(self.c1, c2 + 1) for r in range(self.r1, r2 + 1)] if by_cols else
                 [(r, c) for r in range(self.r1, r2 + 1) for c in range(self.c1, c2 + 1)])
        start = (after.r1, after.c1) if isinstance(after, Range) else (self.r1, self.c1)
        try:
            k0 = cells.index(start)
        except ValueError:
            k0 = 0 if not prev else len(cells) - 1
        order = (cells[k0 + 1:] + cells[:k0 + 1]) if not prev else (cells[:k0][::-1] + cells[k0:][::-1])
        w = what if match_case else what.lower()
        for r, c in order:
            text = self._find_text(r, c, lookin)
            t = text if match_case else text.lower()
            if (lookat == 1 and t == w) or (lookat != 1 and w in t):
                return Range(h, self.sheet, r, c)
        return NOTHING

    def _find_text(self, r: int, c: int, lookin: int) -> str:
        """Find が見る文字：xlValues は表示の文字、xlFormulas は数式（定数は入れたときの文字）。"""
        h = self.host
        sh = h.sheet(self.sheet)
        key = (sh.name, r, c)
        cell = sh.cells.get((r, c))
        if lookin == -4123 and key not in h.over and cell is not None and cell.formula:
            return "=" + cell.formula
        if lookin == -4163:
            return Range(h, self.sheet, r, c).text()
        v = h.get(self.sheet, r, c)
        return "" if v is EMPTY else to_str(v)


def _NA():
    from ..formula.values import ExcelError

    return ExcelError("#N/A")


def _col_arg(v) -> int:
    if isinstance(v, str):
        return col_to_num(v.upper())
    return to_long(v)


class _Lines:
    """Rows・Columns：Count と、番号での行・列。"""

    type_name = "Range"

    def __init__(self, rng: Range, kind: str):
        self.rng = rng
        self.kind = kind

    @property
    def count(self) -> int:
        return self.rng.rows if self.kind == "rows" else self.rng.cols

    def item(self, i: int) -> Range:
        r = self.rng
        if self.kind == "rows":
            return Range(r.host, r.sheet, r.r1 + i - 1, r.c1, r.r1 + i - 1, r.c2)
        return Range(r.host, r.sheet, r.r1, r.c1 + i - 1, r.r2, r.c1 + i - 1)

    def vba_get(self, name, args, named):
        if name in ("count", "countlarge"):
            return self.count
        if name == "item":
            return self.item(to_long(args[0]))
        return self.rng.vba_get(name, args, named)

    def vba_default(self):
        return self.rng.vba_default()


class _Areas:
    def __init__(self, items):
        self.items = items

    def vba_get(self, name, args, named):
        if name == "count":
            return len(self.items)
        if name == "item":
            return self.items[to_long(args[0]) - 1]
        raise VbaError(438)


# ====================================================================== Worksheet・Workbook
class Worksheet:
    type_name = "Worksheet"

    def __init__(self, host: Host, name: str):
        self.host = host
        self.name = host.sheet(name).name

    def same(self, other) -> bool:
        return isinstance(other, Worksheet) and other.name == self.name

    def range(self, args) -> Range:
        if not args:
            raise VbaError(5)
        if len(args) >= 2 and isinstance(args[0], Range) and isinstance(args[1], Range):
            a, b = args[0], args[1]
            return Range(self.host, self.name, min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2), max(a.c2, b.c2))
        if len(args) >= 2:
            a = self.range([args[0]])
            b = self.range([args[1]])
            return Range(self.host, self.name, min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2), max(a.c2, b.c2))
        a = args[0]
        if isinstance(a, Range):
            return a
        text = to_str(a).replace("$", "").strip()
        if "," in text:
            raise Unsupported("離れた範囲（A1,C3）")
        if "!" in text:
            sheet, text = text.rsplit("!", 1)
            return Worksheet(self.host, sheet.strip("'")).range([text])
        named = self._name(text)
        if named is not None:
            return named
        parts = text.split(":")
        try:
            from ..formula.parser import parse_cell

            if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                return Range(self.host, self.name, int(parts[0]), 1, int(parts[1]), MAX_COL)
            if len(parts) == 2 and parts[0].isalpha() and parts[1].isalpha():
                return Range(self.host, self.name, 1, col_to_num(parts[0].upper()), MAX_ROW, col_to_num(parts[1].upper()))
            r1, c1 = parse_cell(parts[0].upper())
            r2, c2 = parse_cell(parts[-1].upper())
        except Exception:  # noqa: BLE001
            raise VbaError(1004, f"範囲「{to_str(a)}」が分かりません")
        return Range(self.host, self.name, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2))

    def _name(self, text: str) -> Range | None:
        for d in self.host.wb.names:
            if d.name.lower() == text.lower() and (d.sheet_scope is None or d.sheet_scope == self.name):
                ref = d.text.lstrip("=")
                if "!" in ref:
                    sheet, cells = ref.rsplit("!", 1)
                    return Worksheet(self.host, sheet.strip("'")).range([cells])
        return None

    def cells(self, args) -> Range:
        if not args:
            return Range(self.host, self.name, 1, 1, MAX_ROW, MAX_COL)
        if len(args) == 1:
            k = to_long(args[0]) - 1
            return Range(self.host, self.name, k // MAX_COL + 1, k % MAX_COL + 1)
        return Range(self.host, self.name, to_long(args[0]), _col_arg(args[1]))

    def vba_get(self, name: str, args: list, named: dict):
        h = self.host
        if name == "name":
            return self.name
        if name == "codename":
            return h.sheet(self.name).code_name or ""
        if name == "index":
            return next(i for i, s in enumerate(h.wb.sheets, 1) if s.name == self.name)
        if name == "range":
            return self.range(args)
        if name == "cells":
            return self.cells(args)
        if name in ("rows", "columns"):
            whole = Range(h, self.name, 1, 1, MAX_ROW, MAX_COL)
            return whole.vba_get(name, args, named)
        if name == "usedrange":
            sh = h.sheet(self.name)
            keys = [k for k in sh.cells if sh.cells[k].value is not None or sh.cells[k].has_formula]
            keys += [(k[1], k[2]) for k, v in h.over.items() if k[0] == sh.name and v is not None]
            if not keys:
                return Range(h, self.name, 1, 1)
            return Range(h, self.name, min(k[0] for k in keys), min(k[1] for k in keys), max(k[0] for k in keys),
                         max(k[1] for k in keys))
        if name == "activate":
            h.active = self.name
            if self.name not in (getattr(h, "selected_sheets", None) or []):
                h.selected_sheets = [self.name]
            return EMPTY
        if name == "select":
            h.active = self.name
            h.selected_sheets = [self.name]
            return EMPTY
        if name == "protect":
            h.protected[self.name] = True
            return EMPTY
        if name == "unprotect":
            h.protected[self.name] = False
            return EMPTY
        if name == "protectcontents":
            return h.protected.get(self.name, False)
        if name in ("printout", "printpreview"):
            h.printed.append(self.name)
            return EMPTY
        if name == "visible":
            st = h.sheet(self.name).state
            return -1 if st == "visible" else (0 if st == "hidden" else 2)
        if name in ("parent",):
            return h.book
        if name == "calculate":
            return EMPTY
        if name in ("next", "previous"):
            names = [s.name for s in h.wb.sheets]
            i = names.index(self.name) + (1 if name == "next" else -1)
            return Worksheet(h, names[i]) if 0 <= i < len(names) else NOTHING
        raise Unsupported(f"Worksheet の「{name}」")

    def vba_set(self, name: str, v) -> None:
        raise Unsupported(f"Worksheet の「{name}」への書き込み（シートの名前・表示などは再現しません）")

    def vba_default(self):
        raise VbaError(438)


class Sheets:
    type_name = "Sheets"

    def __init__(self, host: Host):
        self.host = host

    def item(self, key) -> Worksheet:
        sheets = [s for s in self.host.wb.sheets]
        if isinstance(key, str):
            return Worksheet(self.host, key)
        i = to_long(key)
        if not 1 <= i <= len(sheets):
            raise VbaError(9)
        return Worksheet(self.host, sheets[i - 1].name)

    def vba_call(self, args, named):
        if not args:
            return self
        if isinstance(args[0], VArray):
            return SheetGroup(self.host, [self.item(x).name for x in args[0].flat()])
        return self.item(args[0])

    def vba_get(self, name, args, named):
        if name == "count":
            return len(self.host.wb.sheets)
        if name == "item":
            return self.item(args[0])
        if name == "add":
            raise Unsupported("シートの追加")
        raise VbaError(438)


class SheetGroup:
    """まとめて選んだシート（Sheets(Array("A", "B"))）。"""

    type_name = "Sheets"

    def __init__(self, host: Host, names: list[str]):
        self.host = host
        self.names = names

    def vba_get(self, name, args, named):
        h = self.host
        if name in ("select",):
            h.selected_sheets = list(self.names)
            h.active = self.names[0]
            return EMPTY
        if name in ("printout", "printpreview"):
            h.printed.extend(self.names)
            return EMPTY
        if name == "count":
            return len(self.names)
        if name == "item":
            return Worksheet(h, self.names[to_long(args[0]) - 1])
        raise Unsupported(f"まとめて選んだシートの「{name}」")


class Window:
    """ActiveWindow：選んでいるシートと、画面の設定（値には関わらない）。"""

    type_name = "Window"

    def __init__(self, host: Host):
        self.host = host

    def vba_get(self, name, args, named):
        h = self.host
        if name == "selectedsheets":
            return SheetGroup(h, list(getattr(h, "selected_sheets", None) or [h.active]))
        if name in ("zoom", "scrollrow", "scrollcolumn", "freezepanes", "displaygridlines", "windowstate",
                    "displayheadings", "smallscroll", "largescroll"):
            return EMPTY
        if name == "activesheet":
            return Worksheet(h, h.active)
        raise Unsupported(f"ActiveWindow の「{name}」")

    def vba_set(self, name, v):
        if name in ("zoom", "scrollrow", "scrollcolumn", "freezepanes", "displaygridlines", "windowstate",
                    "displayheadings", "splitrow", "splitcolumn", "displayworkbooktabs"):
            return  # 画面の見え方の設定：値には関わらない
        raise Unsupported(f"ActiveWindow の「{name}」への書き込み")


class Workbook:
    type_name = "Workbook"

    def __init__(self, host: Host):
        self.host = host

    def vba_get(self, name, args, named):
        h = self.host
        if name in ("worksheets", "sheets"):
            s = Sheets(h)
            return s.vba_call(args, named) if args else s
        if name == "activesheet":
            return Worksheet(h, h.active)
        if name in ("name",):
            return h.wb.filename or "Book1.xlsm"
        if name in ("fullname",):
            return h.wb.filename or "Book1.xlsm"
        if name == "path":
            return ""
        if name in ("save",):
            return EMPTY
        if name in ("saveas", "savecopyas"):
            raise NeedsInput("save", "マクロが、別の名前でブックを保存しようとしています")
        if name == "close":
            return EMPTY
        if name in ("protect", "unprotect"):
            return EMPTY
        raise Unsupported(f"Workbook の「{name}」")

    def vba_set(self, name, v):
        if name == "saved":
            return
        raise Unsupported(f"Workbook の「{name}」への書き込み")


class Application:
    type_name = "Application"

    def __init__(self, host: Host):
        self.host = host

    def vba_get(self, name, args, named):
        h = self.host
        if name in ("activesheet", "activeworkbook", "thisworkbook", "worksheets", "sheets", "range", "cells", "rows",
                    "columns", "activecell", "selection", "worksheetfunction"):
            obj = h.global_object(name)
            if args and hasattr(obj, "vba_call"):
                return obj.vba_call(args, named)
            return obj
        if name == "screenupdating":
            return h.screen_updating
        if name == "displayalerts":
            return h.display_alerts
        if name == "enableevents":
            return h.events_enabled
        if name == "calculation":
            return h.calculation
        if name in ("calculate", "calculatefull", "wait", "cutcopymode", "statusbar", "doevents"):
            return EMPTY
        if name == "intersect":
            rs = [a for a in args if isinstance(a, Range)]
            if not rs:
                raise VbaError(424)
            cur = rs[0]
            for b in rs[1:]:
                if b.sheet != cur.sheet:
                    raise VbaError(1004)
                r1, c1 = max(cur.r1, b.r1), max(cur.c1, b.c1)
                r2, c2 = min(cur.r2, b.r2), min(cur.c2, b.c2)
                if r1 > r2 or c1 > c2:
                    return NOTHING
                cur = Range(h, cur.sheet, r1, c1, r2, c2)
            return cur
        if name == "inputbox":
            if h.rt.answers:
                return h.rt.answers.pop(0)
            raise NeedsInput("inputbox", to_str(args[0]) if args else "")
        if name == "run":
            n = to_str(args[0]).rsplit("!", 1)[-1].strip("'")
            return h.rt.run(n, list(args[1:]))
        if name == "version":
            return "16.0"
        if name in ("getopenfilename", "getsaveasfilename"):
            raise NeedsInput("file", "マクロが、ファイルを選ぶ画面を出そうとしています")
        if name == "username":
            raise NeedsInput("username", "マクロが、Excel の使用者の名前を読もうとしています")
        if name == "quit":
            from .interp import _EndProgram

            raise _EndProgram()
        if name == "ontime":
            raise Unsupported("Application.OnTime（時間を決めてマクロを動かす）")
        raise Unsupported(f"Application の「{name}」")

    def vba_set(self, name, v):
        h = self.host
        if name == "screenupdating":
            h.screen_updating = to_bool(v)
        elif name == "displayalerts":
            h.display_alerts = to_bool(v)
        elif name == "enableevents":
            h.events_enabled = to_bool(v)
        elif name == "calculation":
            h.calculation = to_long(v)
        elif name == "cutcopymode":
            if isinstance(v, bool) and not v:
                h.clipboard = None
        elif name in ("statusbar", "cursor", "interactive", "moveafterreturndirection", "displaystatusbar",
                      "asktoupdatelinks", "displayformulabar", "windowstate"):
            pass  # 画面の見え方の設定：値には関わらない
        else:
            raise Unsupported(f"Application の「{name}」への書き込み")


class WorksheetFunction:
    """Application.WorksheetFunction.X：todoku の計算式の関数で計算する（Excel と同じ規則）。"""

    type_name = "WorksheetFunction"

    def __init__(self, host: Host):
        self.host = host

    def vba_get(self, name, args, named):
        from ..formula.evaluator import Evaluator
        from ..formula.functions import EAGER
        from ..formula.values import Arr, ExcelError, Rng

        fname = name.upper()
        f = EAGER.get(fname)
        if f is None:
            raise Unsupported(f"WorksheetFunction.{name}")
        h = self.host
        ev = Evaluator(h.calc(), h.active, 1, 1)
        conv = []
        for a in args:
            if isinstance(a, Range):
                conv.append(Rng(a.sheet, a.r1, a.c1, a.r2, a.c2))
            elif isinstance(a, VArray):
                if len(a.lbounds) == 1:
                    conv.append(Arr([[_from_vba(x) for x in a.flat()]]))
                else:
                    lo_r, lo_c = a.lbounds
                    conv.append(Arr([[_from_vba(a.data.get((r, c), EMPTY)) for c in range(lo_c, a.ubounds[1] + 1)]
                                     for r in range(lo_r, a.ubounds[0] + 1)]))
            elif a is MISSING:
                conv.append(None)
            else:
                conv.append(_from_vba(a))
        try:
            v = f(ev, *conv)
        except ExcelError as e:
            v = e
        if isinstance(v, Rng):
            v = ev.intersect(v) if v.is_cell else Range(h, v.sheet, v.r1, v.c1, v.r2, v.c2)
        if isinstance(v, ExcelError):
            raise VbaError(1004, f"WorksheetFunction クラスの {name} プロパティを取得できません")
        if isinstance(v, Arr):
            out = VArray([1, 1], [v.rows, v.cols])
            for i, row in enumerate(v.data):
                for j, x in enumerate(row):
                    out.data[(i + 1, j + 1)] = EMPTY if x is None else x
            return out
        if isinstance(v, float) and v == int(v) and abs(v) < 2**31:
            return int(v)
        return EMPTY if v is None else v
