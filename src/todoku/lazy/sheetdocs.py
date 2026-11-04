"""大きな .ods・.xlsb を分けて読む：シートの部品を流し読みして行の索引を作り、頼まれた行だけの小さなパッケージを、いままでの読み手で読む。

読む範囲・続きの印・欄は XlsxLazy と同じ（行の索引の形をそろえる）。書き込みは、いままでの api.fill（全体を読む）。
"""

from __future__ import annotations

import bisect
import io
import struct
import xml.parsers.expat as expat
import zipfile
from array import array

from .. import limits
from .textstream import Source
from .xlsxdoc import Member, SeekableInflate, XlsxLazy, _no_doctype


class _SheetBook(XlsxLazy):
    """.ods・.xlsb に共通：ZIP の部品・段階・検査。"""

    streams = False

    def _open_zip(self, path: str, name: str) -> None:
        self.path, self.name = path, name
        self.src = Source(path)
        self.z = zipfile.ZipFile(path)
        self.ratio = 0.25
        self._meta = None
        self._fx = None
        self.size = self.src.size

    def close(self) -> None:
        self.z.close()
        self.src.close()

    def _load_meta(self):
        return self._meta

    @property
    def sheet_meta(self):
        return self._meta

    @property
    def strings(self):
        return None

    def _blanks(self, si: int, i: int, k: int, ix) -> list[str]:
        return []

    def overview(self) -> dict:
        sheets = []
        for d in self._meta:
            ix = d["index"]
            one = {"name": d["name"], "state": d.get("state", "visible")}
            if ix is not None:
                one.update({"rows": len(ix.rows), "first_row": ix.rows[0] if ix.rows else None, "last_row": ix.rows[-1] if ix.rows else None,
                            "cells_with_value": ix.values, "formulas": ix.formulas})
            sheets.append(one)
        return {"sheets": sheets, "ranges": "sheet と rows（行の範囲）", "bytes": self.size}

    def overview_text(self) -> str:
        out = [f"# {self.name}", "", f"- 表計算のブック（{self.size:,} バイト）。全体を一度には出しません。read_window に sheet と rows（\"1-500\"）を渡して読みます"
               "（cursor をたどれば、すべてのシートのすべての行を読めます。窓を作るのが上限で止まるときは raw: true で字のまま読めます）。list_fields も同じ範囲です。",
               "", "## シート", "", "| シート | 行（範囲） | 値のあるセル | 数式 |", "|---|---|---|---|"]
        for s in self.overview()["sheets"]:
            if s.get("rows"):
                out.append(f"| {s['name']} | {s['rows']:,} 行（{s['first_row']}〜{s['last_row']}） | {s['cells_with_value']:,} | {s['formulas']:,} |")
            else:
                out.append(f"| {s['name']} | セルのないシート | | |")
        return "\n".join(out) + "\n"

    def grade(self):
        from ..grade import Reason, _final

        write = [Reason("B", "欄を範囲ごとに見つける", "大きなブックなので、欄は頼まれた行の範囲ごとに見つけます。どのセルにも番地で書けます")]
        g = _final([], write)
        g.deferred = ["未確認：全部のセルを窓に出すこと（audit_document で全部の行を確かめます）"]
        return g

    def audit_begin(self) -> dict:
        return {"sheet": 0, "next_i": 0, "cells": 0, "problems": [], "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        import os

        if state["sheet"] == 0 and state["next_i"] == 0 and not state["cells"] and os.path.getsize(self.path) <= limits.MAX_FILE // 8:
            from .. import api

            au = api.audit_whole(self.path, None, None)
            state["checks"], state["full"], state["done"] = au.get("checks", []), au, True
            return state
        si = state["sheet"]
        if si < len(self._meta):
            d = self._meta[si]
            ix = d["index"]
            if ix is None or state["next_i"] >= len(ix.rows):
                state["sheet"] = si + 1
                state["next_i"] = 0
                return state
            i = state["next_i"]
            k = min(len(ix.rows), i + 2048)
            sh = self._sheet_doc(si, ix.rows[i], ix.rows[k - 1])
            state["cells"] += sum(1 for ln in sh.get("reading", []) if ln.get("text"))
            state["next_i"] = k
            limits.progress(f"シート「{d['name']}」を確かめています", k, len(ix.rows))
            return state
        state["checks"] = [{"name": "読み落とし（セルの値）", "ok": True, "agent": False, "items": [],
                            "detail": f"全部のシートの全部の行を、窓の読み手で読みました（字のある行 {state['cells']:,}）"}]
        state["still_deferred"] = ["未確認：窓の読み手と別の読み方で照らすこと・書き込みの試し（大きすぎるので検査では試していません。fill_document のときに証明します）"]
        state["done"] = True
        return state


# ====================================================================== OpenDocument の表計算
_ROW_GROUPS = ("table-row", "table-header-rows", "table-rows", "table-row-group")


class OdsSheet:
    """1 つの table:table の索引。繰り返しの行（number-rows-repeated）は、はじめの行の番号で持つ。"""

    def __init__(self, name: str):
        self.name = name
        self.rows = array("I")
        self.offs = array("Q")
        self.reps = array("I")
        self.ends = array("Q")
        self.values = 0
        self.formulas = 0
        self.cells = 0
        self.max_col = 0
        self.tail = b""
        self.open_tag = (0, 0)  # table:table の開きタグの始めから、行の前まで
        self.pre_end = 0
        self.post_start = 0
        self.end = 0
        self.data_end = 0

    def locate(self, a: int, b: int) -> tuple[int, int]:
        i = bisect.bisect_right(self.rows, a) - 1
        if i < 0 or self.rows[i] + self.reps[i] - 1 < a:
            i += 1
        j = bisect.bisect_right(self.rows, b)
        return i, j


class OdsLazy(_SheetBook):
    kind = "ods"

    def __init__(self, path: str, name: str):
        self._open_zip(path, name)
        self.member = Member(self.src, self.z.getinfo("content.xml"))
        self.reader = SeekableInflate(self.member)
        self.sheets: list[OdsSheet] = []
        self._scan()
        self._meta = [{"name": s.name, "path": "content.xml", "state": "visible", "index": s} for s in self.sheets]
        self.est_chars = sum(s.values for s in self.sheets) * 16

    def _scan(self) -> None:
        p = expat.ParserCreate()
        p.buffer_text = True
        p.StartDoctypeDeclHandler = _no_doctype
        stack: list[str] = []
        cur: list[OdsSheet | None] = [None]
        next_row = [1]
        col = [0]
        head_end = [None]
        tables_end = [None]

        def start(name, attrs):
            local = name.rsplit(":", 1)[-1]
            sh = cur[0]
            if local == "table" and name.endswith("table:table") and sh is None and stack and stack[-1] == "spreadsheet":
                if head_end[0] is None:
                    head_end[0] = p.CurrentByteIndex
                sh = cur[0] = OdsSheet(attrs.get("table:name", f"Sheet{len(self.sheets) + 1}"))
                sh.open_tag = (p.CurrentByteIndex, 0)
                next_row[0] = 1
                self.sheets.append(sh)
            elif sh is not None:
                if local in _ROW_GROUPS and stack[-1] in ("table", "table-header-rows", "table-rows", "table-row-group"):
                    if not sh.pre_end:
                        sh.pre_end = p.CurrentByteIndex
                if local == "table-row":
                    n = int(attrs.get("table:number-rows-repeated", "1") or 1)
                    sh.rows.append(next_row[0])
                    sh.reps.append(n)
                    sh.offs.append(p.CurrentByteIndex)
                    next_row[0] += n
                    col[0] = 0
                elif local in ("table-cell", "covered-table-cell") and stack[-1] == "table-row":
                    m = int(attrs.get("table:number-columns-repeated", "1") or 1)
                    sh.cells += m
                    if attrs.get("office:value-type"):
                        sh.values += 1
                        col[0] += m
                        sh.max_col = max(sh.max_col, col[0])
                    else:
                        col[0] += m
                    if attrs.get("table:formula"):
                        sh.formulas += 1
            stack.append(local)

        def end(name):
            local = stack.pop()
            sh = cur[0]
            if sh is None:
                return
            if local == "table-row":
                sh.ends.append(_end_of(p))
            if local in _ROW_GROUPS and stack and stack[-1] == "table":
                sh.post_start = _end_of(p)
            if local == "table" and name.endswith("table:table") and stack and stack[-1] == "spreadsheet":
                sh.end = p.CurrentByteIndex
                if not sh.pre_end:
                    sh.pre_end = sh.end
                if not sh.post_start:
                    sh.post_start = sh.end
                cur[0] = None
                tables_end[0] = p.CurrentByteIndex

        buf = bytearray()
        at = [0]

        def _end_of(pp) -> int:
            i = pp.CurrentByteIndex
            k = buf.find(b">", i - at[0])
            return at[0] + k + 1 if k >= 0 else i

        p.StartElementHandler = start
        p.EndElementHandler = end
        total = 0
        size = self.member.size or 1
        for c in self.reader.scan():
            buf += c
            total += len(c)
            p.Parse(c, False)
            if len(buf) > 1 << 18:
                drop = len(buf) - (1 << 17)
                del buf[:drop]
                at[0] += drop
            limits.progress("content.xml の行の位置を調べています", total, size)
        p.Parse(b"", True)
        if head_end[0] is None or tables_end[0] is None:
            raise ValueError("表計算のシートがありません")
        self.head = self.reader.read(0, head_end[0])
        k = self.reader.read(tables_end[0], tables_end[0] + 64).find(b">")
        self.after_tables = tables_end[0] + k + 1
        self.tail = self.reader.read(self.after_tables, total)
        self.between = []
        for sh in self.sheets:
            sh.pre = self.reader.read(sh.open_tag[0], sh.pre_end)
            sh.post = self.reader.read(sh.post_start, sh.end + 64)
            sh.post = sh.post[:sh.post.find(b">", sh.end - sh.post_start) + 1]
            sh.data_end = sh.offs[-1] + 1 if sh.offs else 0

    def _mini(self, wanted: dict[int, list[tuple[int, int]]]) -> bytes:
        out = bytearray(self.head)
        for si, sh in enumerate(self.sheets):
            out += sh.pre
            at = 1
            done_upto = -1
            for a, b in sorted(wanted.get(si, [])):
                i, j = sh.locate(a, b)
                i = max(i, done_upto + 1)
                for k in range(i, j):
                    r = sh.rows[k]
                    if r > at:
                        out += b'<table:table-row table:number-rows-repeated="%d"/>' % (r - at)
                    out += self.reader.read(sh.offs[k], sh.ends[k])
                    at = r + sh.reps[k]
                    done_upto = k
            out += sh.post
        out += self.tail
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
            for i in self.z.infolist():
                n = i.filename
                if n == "content.xml":
                    z.writestr(n, bytes(out))
                elif n.lower().startswith("pictures/") or i.file_size > 1 << 24:
                    z.writestr(n, b"")
                else:
                    z.writestr(n, self._part(i))
        return buf.getvalue()

    def rows_book(self, wanted):
        from ..odf.sheet import OdsWorkbook

        return OdsWorkbook(self._mini(wanted), self.name)

    def _part(self, info) -> bytes:
        cache = self.__dict__.setdefault("_parts", {})
        got = cache.get(info.filename)
        if got is None:
            got = cache[info.filename] = Member(self.src, info).read()
        return got

    def _sheet_doc(self, si: int, a: int, b: int) -> dict:
        from ..odf.sheetwindow import build

        ix = self.sheets[si]
        wanted = [(a, b)]
        if ix.rows and ix.rows[0] < a:
            wanted.insert(0, (ix.rows[0], ix.rows[0]))
        win = build(self._mini({si: wanted}), self.name)
        return win["sheets"][si]

    def _no_rows_lines(self, si: int) -> list[str]:
        from ..odf.sheetwindow import build
        from ..xlsx.window import _md_sheet

        return _md_sheet(build(self._mini({}), self.name)["sheets"][si])


# ====================================================================== .xlsb
class XlsbSheet:
    def __init__(self):
        self.rows = array("I")
        self.offs = array("Q")
        self.ends = array("Q")
        self.values = self.formulas = self.cells = 0
        self.max_col = 0
        self.tail = b""
        self.data_end = 0
        self.head_bytes = b""
        self.tail_bytes = b""

    def locate(self, a: int, b: int) -> tuple[int, int]:
        return bisect.bisect_left(self.rows, a), bisect.bisect_right(self.rows, b)


def _walk_biff12(reader, sh: XlsbSheet) -> None:
    """レコードの頭をたどり、行（BrtRowHdr）の位置を集める。同じ番号の行・順に並ばない行があれば ValueError（窓を作る道で読む）。"""
    from ..xlsb.book import BEGIN_DATA, CELLS, END_DATA, F_BOOL, F_ERR, F_NUM, F_ST, ROW

    carry = bytearray()
    at = 0
    pos = 0
    data_begin = None
    data_end = None
    forms = {F_ST, F_NUM, F_BOOL, F_ERR}
    runs: list[tuple[int, int, int]] = []
    open_row = None

    def pump(final: bool) -> None:
        nonlocal pos, data_begin, data_end, open_row, at
        n = len(carry)
        while True:
            k = pos - at
            if k >= n or (not final and k + 16 > n):
                break  # 頭が途中で切れているかもしれない：次のかたまりで
            b0 = carry[k]
            k2 = k + 1
            rid = b0 & 0x7F
            if b0 & 0x80:
                rid |= (carry[k2] & 0x7F) << 7
                k2 += 1
            ln = shift = 0
            for _ in range(4):
                b = carry[k2]
                k2 += 1
                ln |= (b & 0x7F) << shift
                shift += 7
                if not b & 0x80:
                    break
            if rid == ROW:
                row = struct.unpack_from("<I", carry, k2)[0] + 1
                if open_row is not None:
                    runs.append((open_row[0], open_row[1], pos))
                open_row = (row, pos)
            elif rid == BEGIN_DATA:
                data_begin = pos + (k2 - k) + ln
            elif rid == END_DATA:
                if open_row is not None:
                    runs.append((open_row[0], open_row[1], pos))
                    open_row = None
                data_end = pos
            elif rid in CELLS:
                sh.cells += 1
                sh.values += 1
                if rid in forms:
                    sh.formulas += 1
                if k2 + 4 <= n:
                    sh.max_col = max(sh.max_col, struct.unpack_from("<I", carry, k2)[0] + 1)
            pos += (k2 - k) + ln
        drop = min(pos - at, len(carry))
        if drop > 0:
            del carry[:drop]
            at += drop

    for c in reader.scan():
        carry += c
        pump(False)
    pump(True)
    if open_row is not None:
        runs.append((open_row[0], open_row[1], data_end if data_end is not None else pos))
    if data_begin is None:
        raise ValueError("シートのデータがありません")
    if data_end is None:
        data_end = pos
    sh.data_begin, sh.data_end = data_begin, data_end
    runs.sort()
    last = -1
    cum = 0
    for row, a, b in runs:
        if row == last:
            raise ValueError("同じ番号の行があります")
        last = row
        sh.rows.append(row)
        sh.offs.append(a)
        sh.ends.append(b)
        cum += b - a


class XlsbLazy(_SheetBook):
    kind = "xlsb"

    def __init__(self, path: str, name: str):
        from ..xlsb.book import XlsbWorkbook

        self._open_zip(path, name)
        infos = self.z.infolist()
        self.parts: dict[str, bytes] = {}
        self.readers: dict[str, SeekableInflate] = {}
        sheet_parts = {i.filename for i in infos if i.filename.lower().startswith("xl/worksheets/") and i.filename.lower().endswith(".bin")}
        for i in infos:
            if i.filename in sheet_parts or i.filename.endswith("/"):
                continue
            low = i.filename.lower()
            self.parts[i.filename] = b"" if low.startswith(("xl/media/", "xl/embeddings/")) else Member(self.src, i).read()
        self.indexes: dict[str, XlsbSheet] = {}
        for part in sheet_parts:
            r = SeekableInflate(Member(self.src, self.z.getinfo(part)))
            sh = XlsbSheet()
            limits.progress(f"{part} の行の位置を調べています")
            _walk_biff12(r, sh)
            sh.head_bytes = r.read(0, sh.data_begin)
            sh.tail_bytes = r.read(sh.data_end, r.m.size)
            self.readers[part] = r
            self.indexes[part] = sh
        wb = XlsbWorkbook(self._mini({}, skeleton=True), name)
        self.strings_list = wb.shared_strings
        self._meta = []
        for s in wb.sheets:
            ix = self.indexes.get(s.path)
            self._meta.append({"name": s.name, "path": s.path, "state": s.state, "index": ix})
        self.parts = {n: (b"" if n.lower().endswith("sharedstrings.bin") else d) for n, d in self.parts.items()}
        self.est_chars = sum(x.values for x in self.indexes.values()) * 16

    def _mini(self, wanted: dict[int, list[tuple[int, int]]], skeleton: bool = False) -> bytes:
        buf = io.BytesIO()
        path_rows: dict[str, list] = {}
        if not skeleton:
            for si, spans in wanted.items():
                path_rows[self._meta[si]["path"]] = spans
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
            for i in self.z.infolist():
                n = i.filename
                if n in self.indexes:
                    sh = self.indexes[n]
                    out = bytearray(sh.head_bytes)
                    done = -1
                    for a, b in sorted(path_rows.get(n, [])):
                        x, y = sh.locate(a, b)
                        for k in range(max(x, done + 1), y):
                            out += self.readers[n].read(sh.offs[k], sh.ends[k])
                            done = k
                    out += sh.tail_bytes
                    z.writestr(n, bytes(out))
                elif n in self.parts:
                    z.writestr(n, self.parts[n])
        return buf.getvalue()

    def rows_book(self, wanted):
        return self._book(wanted)

    def _book(self, wanted):
        from ..xlsb.book import XlsbWorkbook

        return XlsbWorkbook(self._mini(wanted), self.name, strings=self.strings_list)

    def _sheet_doc(self, si: int, a: int, b: int) -> dict:
        from ..xlsx.window import _MEMO, _Options
        from ..xlsx.window import _sheet_doc as xw_sheet

        ix = self._meta[si]["index"]
        wanted = [(a, b)]
        if ix.rows and ix.rows[0] < a:
            wanted.insert(0, (ix.rows[0], ix.rows[0]))
        wb = self._book({si: wanted})
        sh = wb.sheets[si]
        sh.__dict__[_MEMO] = {}
        try:
            return xw_sheet(wb, sh, _Options(wb))
        finally:
            sh.__dict__.pop(_MEMO, None)

    def _no_rows_lines(self, si: int) -> list[str]:
        from ..xlsx.window import _md_sheet, build

        return _md_sheet(build(self._book({}))["sheets"][si])
