"""表だけの形式（dBase・Lotus・Access など）を、xlsx と同じブックの形（シートとセル）にする土台。

窓・表の取り出し（read_table）・計算の説明は xlsx と同じ処理を使う。行は 1 から、表の見出しがあれば 1 行目。
"""

from __future__ import annotations

from ..xls.book import _ZipShim
from ..xlsb.book import XlsbWorkbook
from ..xlsx.workbook import Cell, Sheet, Style


class TableBook(XlsbWorkbook):
    def __init__(self, blob: bytes, filename: str, kind: str):
        self.kind = kind
        self.filename = filename
        self.zip = _ZipShim(blob)
        self.sheets: list[Sheet] = []
        self.names = []
        self.shared_strings = []
        self.phonetic = {}
        self.cell_sst = {}
        self.rich_cells = {}
        self.styles = [Style("General", True, False), Style("yyyy-mm-dd", True, False)]
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

    @property
    def vba(self):
        return None

    def add_sheet(self, name: str) -> Sheet:
        sh = Sheet(name=name, path=f"#{len(self.sheets) + 1}")
        self.sheets.append(sh)
        return sh

    def put(self, sh: Sheet, r: int, c: int, v, style: int = 0) -> Cell:
        """r・c は 1 始まり。"""
        cell = Cell(sh.name, r, c, style=style)
        cell.value = v
        if isinstance(v, str):
            cell.t = "s"
        elif isinstance(v, bool):
            cell.t = "b"
        sh.cells[(r, c)] = cell
        sh.max_row = max(sh.max_row, r)
        sh.max_col = max(sh.max_col, c)
        return cell
