"""ブックを流して読み、数式を形ごとの長方形に、定数を列ごとの区間に畳む（セルごとの控えを持たない）。

- .xlsx：シートの XML を行の位置の索引（lazy.xlsxdoc.SheetIndex）で少しずつ読む
- ほかの形式（.xls・.xlsb・.ods など）：その形式の読み手が作ったブックのセルをなぞる（読み手がもともとセルを持っている）
"""

from __future__ import annotations

from array import array
from dataclasses import dataclass, field
from xml.parsers import expat

from ..formula.parser import parse_cell
from ..limits import MAX_DECLARED_CELLS
from ..formula.values import ExcelError
from .shape import Shaper

CHUNK_ROWS = 1024


@dataclass(slots=True)
class Rect:
    shape: int
    r1: int
    c1: int
    r2: int
    c2: int
    text: str | None = None  # 左上のセルの数式（ファイルに書かれたまま。共有数式の子なら None）
    kind: str = "formula"  # formula / array / dynamic_array / data_table / unreadable
    note: dict | None = None

    @property
    def cells(self) -> int:
        return (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1)


def _bisect_pairs(a: array, r: int) -> int:
    """r を含むか、r より下の最初の区間の位置（偶数の添字）。"""
    lo, hi = 0, len(a) // 2
    while lo < hi:
        m = (lo + hi) // 2
        if a[m * 2 + 1] < r:
            lo = m + 1
        else:
            hi = m
    return lo * 2


class Runs:
    """列ごとの行の区間（[始め, 終わり] の並び）。上から順に足す。"""

    __slots__ = ("cols",)

    def __init__(self):
        self.cols: dict[int, array] = {}

    def add(self, row: int, col: int) -> None:
        a = self.cols.get(col)
        if a is None:
            self.cols[col] = array("I", (row, row))
        elif a[-1] + 1 == row:
            a[-1] = row
        elif a[-1] != row:
            a.append(row)
            a.append(row)

    def covers(self, col: int, r1: int, r2: int) -> bool:
        """列の r1..r2 がすべて区間の中か。"""
        a = self.cols.get(col)
        if a is None:
            return False
        i = _bisect_pairs(a, r1)
        return i < len(a) and a[i] <= r1 and a[i + 1] >= r2

    def has(self, row: int, col: int) -> bool:
        return self.covers(col, row, row)

    def count(self, col: int, r1: int, r2: int) -> int:
        a = self.cols.get(col)
        if a is None:
            return 0
        n = 0
        i = _bisect_pairs(a, r1)
        while i < len(a) and a[i] <= r2:
            n += min(a[i + 1], r2) - max(a[i], r1) + 1
            i += 2
        return n

    def spans(self, col: int, r1: int, r2: int):
        a = self.cols.get(col)
        if a is None:
            return
        i = _bisect_pairs(a, r1)
        while i < len(a) and a[i] <= r2:
            yield max(a[i], r1), min(a[i + 1], r2)
            i += 2

    def last_at_or_before(self, col: int, row: int) -> int | None:
        """row 以前で、区間の中の一番下の行。"""
        a = self.cols.get(col)
        if not a:
            return None
        lo, hi = 0, len(a) // 2
        while lo < hi:
            m = (lo + hi) // 2
            if a[m * 2] <= row:
                lo = m + 1
            else:
                hi = m
        if lo == 0:
            return None
        return min(a[(lo - 1) * 2 + 1], row)


@dataclass
class SheetScan:
    name: str
    state: str = "visible"
    rects: list[Rect] = field(default_factory=list)
    const: Runs = field(default_factory=Runs)  # 値のある（数式でない）セル
    text: Runs = field(default_factory=Runs)  # そのうち、字のセル
    occupied: Runs = field(default_factory=Runs)  # 値か数式のあるセル
    hidden_rows: array = field(default_factory=lambda: array("I"))  # [始め, 終わり] の並び
    formulas: int = 0  # 数式の要素の数（配列数式は 1）
    formula_cells: int = 0  # 数式が値を決めるセルの数
    constants: int = 0
    max_row: int = 0
    max_col: int = 0


class _Feeder:
    """行を受け取り、長方形と区間に畳む。セル：(列, 数式の種類, 数式の字, 共有の番号, ref, 動的配列か, 定数か, 字か)。"""

    def __init__(self, sh: SheetScan, shaper: Shaper):
        self.sh = sh
        self.shaper = shaper
        self.open: dict[tuple, Rect] = {}
        self.shared: dict[str, int] = {}
        self.orphans: list[tuple[int, int, str]] = []
        self.blocks: list[tuple[int, int, int, int]] = []  # 配列数式・データテーブルの範囲（中のセルは定数ではない）

    def row(self, r: int, hidden: bool, cells: list) -> None:
        sh = self.sh
        if hidden:
            h = sh.hidden_rows
            if h and h[-1] + 1 == r:
                h[-1] = r
            elif not h or h[-1] != r:
                h.append(r)
                h.append(r)
        if not cells:
            return
        if r > sh.max_row:
            sh.max_row = r
        blocks = self.blocks
        if blocks:
            blocks[:] = [b for b in blocks if b[2] >= r]
        cur_sid = cur_c1 = cur_c2 = None
        cur_text = None
        occ = sh.occupied
        for col, fk, ftext, si, ref, dyn, is_const, is_text in cells:
            if col > sh.max_col:
                sh.max_col = col
            if fk is None:
                if not is_const:
                    continue
                if blocks and any(b[0] <= r <= b[2] and b[1] <= col <= b[3] for b in blocks):
                    continue  # 配列数式・データテーブルの範囲の中：値は数式が決める
                occ.add(r, col)
                sh.const.add(r, col)
                sh.constants += 1
                if is_text:
                    sh.text.add(r, col)
                continue
            occ.add(r, col)
            sh.formulas += 1
            if fk == "array":
                self._array(r, col, ftext, ref, dyn)
                continue
            if fk == "dataTable":
                self._data_table(r, col, ref, ftext)
                continue
            if fk == "shared":
                if ftext:
                    sid = self.shared[si] = self.shaper.of_text(ftext, r, col)
                else:
                    sid = self.shared.get(si)
                    if sid is None:
                        self.orphans.append((r, col, si))
                        continue
                    ftext = None
            elif fk == "ast":
                ast, r0, c0 = ftext
                sid = self.shaper.of_ast(ast, r0, c0)
                ftext = None
            elif fk == "bad":
                sid = self.shaper.new(ftext, r, col, None, ftext)
                ftext = None
            else:
                sid = self.shaper.of_text(ftext, r, col)
            sh.formula_cells += 1
            if sid == cur_sid and col == cur_c2 + 1:
                cur_c2 = col
                continue
            if cur_sid is not None:
                self._run(r, cur_sid, cur_c1, cur_c2, cur_text)
            cur_sid, cur_c1, cur_c2, cur_text = sid, col, col, ftext
        if cur_sid is not None:
            self._run(r, cur_sid, cur_c1, cur_c2, cur_text)

    def _run(self, r: int, sid: int, c1: int, c2: int, text) -> None:
        k = (sid, c1, c2)
        o = self.open.get(k)
        if o is not None and o.r2 == r - 1:
            o.r2 = r
            return
        if o is not None:
            self.sh.rects.append(o)
        self.open[k] = Rect(sid, r, c1, r, c2, text)

    def _block(self, ref: str | None, r: int, col: int) -> tuple[int, int, int, int]:
        if ref:
            a = ref.split(":")
            try:
                r1, c1 = parse_cell(a[0])
                r2, c2 = parse_cell(a[-1])
                return r1, c1, r2, c2
            except Exception:  # noqa: BLE001  壊れた ref：そのセルだけ
                pass
        return r, col, r, col

    def _declared(self, ref, r: int, col: int) -> tuple[tuple[int, int, int, int], bool]:
        """宣言された範囲。広すぎる宣言（壊れた・悪意のあるファイル）は広げず、親のセルだけにする。"""
        r1, c1, r2, c2 = self._block(ref, r, col)
        if (r2 - r1 + 1) * (c2 - c1 + 1) > MAX_DECLARED_CELLS or not (r1 <= r <= r2 and c1 <= col <= c2):
            return (r, col, r, col), True
        self.sh.max_row = max(self.sh.max_row, r2)
        self.sh.max_col = max(self.sh.max_col, c2)
        return (r1, c1, r2, c2), False

    def _array(self, r: int, col: int, text: str, ref, dyn: bool) -> None:
        (r1, c1, r2, c2), wide = self._declared(ref, r, col)
        sid = self.shaper.of_array(text, r, col)
        self.sh.formula_cells += (r2 - r1 + 1) * (c2 - c1 + 1)
        self.sh.rects.append(Rect(sid, r1, c1, r2, c2, text, "dynamic_array" if dyn else "array",
                                  {"ref": ref, "master": (r, col), **({"too_wide": True} if wide else {})}))
        if (r1, c1) != (r2, c2):
            self.blocks.append((r1, c1, r2, c2))

    def _data_table(self, r: int, col: int, ref, attrs) -> None:
        (r1, c1, r2, c2), _wide = self._declared(ref, r, col)
        sid = self.shaper.new(f"TABLE@{r},{col}", r, col, None, "データテーブル（What-If 分析）")
        self.sh.formula_cells += (r2 - r1 + 1) * (c2 - c1 + 1)
        self.sh.rects.append(Rect(sid, r1, c1, r2, c2, None, "data_table", {"ref": ref, **(attrs or {})}))
        self.blocks.append((r1, c1, r2, c2))

    def finish(self) -> None:
        for r, col, si in self.orphans:
            sid = self.shared.get(si)
            if sid is None:
                sid = self.shaper.new(f"shared#{si}", r, col, None, "共有数式の親が見つかりません")
            self.sh.formula_cells += 1
            self.sh.rects.append(Rect(sid, r, col, r, col))
        self.sh.rects.extend(self.open.values())
        self.open.clear()
        for rc in self.sh.rects:
            if rc.kind == "formula" and self.shaper.shapes[rc.shape].error:
                rc.kind = "unreadable"


def _is_text_type(t) -> bool:
    return t in ("s", "str", "inlineStr")


def decode_stored(t, v, inline, strings):
    """<c> の値（workbook._cell_value と同じ決まり）。"""
    if t == "inlineStr":
        return inline if inline is not None else ""
    if v is None:
        return None
    if t == "s":
        i = int(v)
        return strings[i] if strings is not None and 0 <= i < len(strings) else ""
    if t == "b":
        return v in ("1", "true")
    if t == "e":
        return ExcelError(v)
    if t == "str":
        return v
    if t == "d":
        return v
    try:
        return float(v)
    except ValueError:
        return v


class _RowXml:
    """<row> の並びのバイト列 → [(行, 隠れているか, [セル])]。セル：[列, t, cm, f の属性, f の字, v の字, is の字]。"""

    def __init__(self):
        self.next_row = 1

    def parse(self, xml: bytes) -> list:
        rows: list = []
        p = expat.ParserCreate()
        p.buffer_text = True
        cur = None  # 今の行のセル
        cell = None
        where = None
        ncol = 1
        in_rph = 0
        this = self

        def start(name, attrs):
            nonlocal cur, cell, where, ncol, in_rph
            local = name[name.find(":") + 1:]
            if local == "c":
                ref = attrs.get("r")
                if ref:
                    n = 0
                    for ch in ref:
                        if ch > "9":
                            n = n * 26 + ord(ch) - 64
                        else:
                            break
                    col = n
                else:
                    col = ncol
                ncol = col + 1
                cell = [col, attrs.get("t"), attrs.get("cm"), None, None, None, None]
            elif local == "f":
                cell[3] = attrs
                cell[4] = ""
                where = 4
            elif local == "v":
                if cell[5] is None:
                    cell[5] = ""
                    where = 5
            elif local == "is":
                cell[6] = ""
            elif local == "t":
                if cell is not None and cell[6] is not None and not in_rph:
                    where = 6
            elif local == "rPh":
                in_rph += 1
            elif local == "row":
                r = attrs.get("r")
                num = int(r) if r else this.next_row
                this.next_row = num + 1
                ncol = 1
                cur = []
                rows.append((num, attrs.get("hidden") in ("1", "true"), cur))

        def chars(s):
            if where is not None:
                cell[where] += s

        def end(name):
            nonlocal cell, where, in_rph
            local = name[name.find(":") + 1:]
            if local == "c":
                cur.append(cell)
                cell = None
                where = None
            elif local in ("f", "v", "t"):
                where = None
            elif local == "rPh":
                in_rph -= 1

        p.StartElementHandler = start
        p.CharacterDataHandler = chars
        p.EndElementHandler = end
        p.Parse(b"<x>" + xml + b"</x>", True)
        return rows


class XlsxSource:
    """.xlsx：シートを流して読む。"""

    def __init__(self, lz):
        self.lz = lz
        self.meta = lz.rows_book({})  # 行のないブック：名前・テーブル・入力規則・条件付き書式・マクロなど
        self.names = [d["name"] for d in lz.sheet_meta]
        self.states = [d["state"] for d in lz.sheet_meta]

    def _xml_rows(self, si: int):
        ix = self.lz.sheet_meta[si]["index"]
        if ix is None:
            return
        n = len(ix.rows)
        reader = _RowXml()
        for i in range(0, n, CHUNK_ROWS):
            reader.next_row = ix.rows[i]
            yield from reader.parse(ix.row_bytes(i, min(i + CHUNK_ROWS, n)))

    def feed(self, si: int, fd: _Feeder) -> None:
        for r, hidden, cells in self._xml_rows(si):
            out = []
            for col, t, cm, fa, ft, v, inl in cells:
                if fa is not None:
                    kind = fa.get("t", "normal")
                    if kind == "dataTable":
                        ft = {k: fa[k] for k in ("dt2D", "dtr", "r1", "r2") if k in fa}
                    out.append((col, kind, ft, fa.get("si"), fa.get("ref"), cm is not None, False, False))
                elif v is not None or inl is not None:
                    has_value = (v is not None and (v != "" or t == "str")) or inl is not None
                    out.append((col, None, None, None, None, False, has_value, _is_text_type(t)))
            fd.row(r, hidden, out)

    def stored(self, si: int):
        """(行, 列, 値, 字の数式の結果か) を行・列の順に。"""
        strings = self.lz.strings
        for r, _h, cells in self._xml_rows(si):
            for col, t, _cm, _fa, _ft, v, inl in cells:
                yield r, col, decode_stored(t, v, inl, strings), (t == "str")

    def fetch(self, si: int, cells) -> dict:
        """いくつかのセルの値：{(行, 列): 値}。セルのある行だけを読む。"""
        want = set(cells)
        ix = self.lz.sheet_meta[si]["index"]
        if ix is None or not want:
            return {}
        strings = self.lz.strings
        out = {}
        reader = _RowXml()
        rows = sorted({r for r, _ in want})
        spans = []
        a = b = rows[0]
        for x in rows[1:]:
            if x == b + 1:
                b = x
            else:
                spans.append((a, b))
                a = b = x
        spans.append((a, b))
        for a, b in spans:
            i, j = ix.locate(a, b)
            for k in range(i, j, CHUNK_ROWS):
                reader.next_row = ix.rows[k]
                for r, _h, row_cells in reader.parse(ix.row_bytes(k, min(k + CHUNK_ROWS, j))):
                    for col, t, _cm, _fa, _ft, v, inl in row_cells:
                        if (r, col) in want:
                            out[(r, col)] = decode_stored(t, v, inl, strings)
        return out


class BookSource:
    """読み手が作ったブック（.xls・.xlsb・.ods など）。"""

    def __init__(self, wb):
        self.meta = wb
        self.wb = wb
        self.names = [s.name for s in wb.sheets]
        self.states = [getattr(s, "state", "visible") for s in wb.sheets]

    def feed(self, si: int, fd: _Feeder) -> None:
        sh = self.wb.sheets[si]
        by_row: dict[int, list] = {}
        for (r, c) in sh.cells:
            by_row.setdefault(r, []).append(c)
        hidden = set(sh.hidden_rows)
        for r in sorted(set(by_row) | hidden):
            out = []
            for c in sorted(by_row.get(r, ())):
                cell = sh.cells[(r, c)]
                if cell.array_master is not None:
                    continue
                if cell.has_formula or cell.parse_error:
                    if cell.array_ref or cell.formula_kind == "array":
                        out.append((c, "array", cell.formula or "", None, cell.array_ref, cell.dynamic, False, False))
                    elif cell._shared_from is not None:
                        ast, dr, dc = cell._shared_from
                        out.append((c, "ast", (ast, r - dr, c - dc), None, None, False, False, False))
                    elif cell.formula is None or (cell.ast is None and cell.parse_error):
                        out.append((c, "bad", cell.parse_error or cell.formula or "", None, None, False, False, False))
                    else:
                        out.append((c, "normal", cell.formula, None, None, False, False, False))
                elif cell.value is not None:
                    out.append((c, None, None, None, None, False, True, isinstance(cell.value, str)))
            fd.row(r, r in hidden, out)

    def stored(self, si: int):
        sh = self.wb.sheets[si]
        for (r, c) in sorted(sh.cells):
            cell = sh.cells[(r, c)]
            yield r, c, cell.value, cell.t == "str"

    def fetch(self, si: int, cells) -> dict:
        sh = self.wb.sheets[si]
        return {k: sh.cells[k].value for k in cells if k in sh.cells}


def scan_book(src) -> tuple[list[SheetScan], Shaper]:
    shaper = Shaper()
    out = []
    for si, name in enumerate(src.names):
        sh = SheetScan(name, src.states[si])
        fd = _Feeder(sh, shaper)
        src.feed(si, fd)
        fd.finish()
        out.append(sh)
    return out, shaper
