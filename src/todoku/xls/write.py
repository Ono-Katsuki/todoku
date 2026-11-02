"""古いExcel（.xls、BIFF8）に、元のファイルを正本にしたまま書き込む。

ブックのストリームをレコードの並びとして持ち、書き込むセルのレコードと、計算結果が変わる数式の
保存値と、フォームの部品の状態（OBJ の中のチェック・選んだ項目・値のバイト）だけを差し替える。ほかのレコードはバイト列のまま残す。レコードの長さが変わると位置がずれるので、
位置を持つレコード（BOUNDSHEET・INDEX・DBCELL・EXTSST）は、組み立て直したあとで正しい位置に直す。
複合ファイルの中のほかのストリームは、1バイトも変えない。
"""

from __future__ import annotations

import struct

from ..formula.values import ExcelError, is_number
from .book import ERRORS, XlsWorkbook, records
from .obj import ObjError, with_state

T_BOF, T_EOF, T_BOUNDSHEET, T_SST, T_EXTSST, T_CONTINUE = 0x0809, 0x000A, 0x0085, 0x00FC, 0x00FF, 0x003C
T_INDEX, T_DBCELL, T_ROW, T_FORMULA, T_STRING = 0x020B, 0x00D7, 0x0208, 0x0006, 0x0207
T_NUMBER, T_RK, T_MULRK, T_LABELSST, T_LABEL, T_RSTRING, T_BLANK, T_MULBLANK, T_BOOLERR = (
    0x0203, 0x027E, 0x00BD, 0x00FD, 0x0204, 0x00D6, 0x0201, 0x00BE, 0x0205)
T_DIMENSIONS, T_OBJ, T_RECALCID, T_DEFCOLWIDTH = 0x0200, 0x005D, 0x01C1, 0x0055
CELLS = {T_NUMBER, T_RK, T_LABELSST, T_LABEL, T_RSTRING, T_BLANK, T_BOOLERR, T_FORMULA}
MULTI = {T_MULRK, T_MULBLANK}
ERRCODE = {v: k for k, v in ERRORS.items()}
MAXREC = 8224


class XlsWriteError(ValueError):
    pass


class Rec:
    __slots__ = ("t", "b", "orig")

    def __init__(self, t: int, b: bytes, orig: bool = True):
        self.t = t
        self.b = b
        self.orig = orig  # 元のファイルのレコードそのまま


def _xlstr(s: str) -> bytes:
    """XLUnicodeString（u16 の文字数・印・文字）。Latin-1 で書けるものは1バイト、ほかは UTF-16。"""
    try:
        raw = s.encode("latin-1")
        return struct.pack("<HB", len(s), 0) + raw
    except UnicodeEncodeError:
        return struct.pack("<HB", len(s), 1) + s.encode("utf-16-le")


class XlsEditor:
    def __init__(self, wb: XlsWorkbook):
        self.wb = wb
        self.recs = [Rec(t, b) for _, t, b in records(wb.data)]
        self.sst_new: list[str] = []
        self.sst_index = {s: i for i, s in enumerate(wb.shared_strings) if s not in ()}
        self.n_sst = len(wb.shared_strings)
        self.changed = False
        self._sheet_ranges()

    # ------------------------------------------------------------------ 位置
    def _sheet_ranges(self) -> None:
        """シートの名前 → レコードの並びの中の [BOF, EOF] の位置。"""
        offsets = []
        pos = 0
        for r in self.recs:
            offsets.append(pos)
            pos += 4 + len(r.b)
        by_off = {o: i for i, o in enumerate(offsets)}
        self.sheet_at: dict[str, tuple[int, int]] = {}
        for sh in self.wb.sheets:
            if not sh.path.startswith("#"):
                continue
            start = by_off.get(int(sh.path[1:]))
            if start is None:
                continue
            depth = 0
            for j in range(start, len(self.recs)):
                t = self.recs[j].t
                if t == T_BOF:
                    depth += 1
                elif t == T_EOF:
                    depth -= 1
                    if depth == 0:
                        self.sheet_at[sh.name] = (start, j)
                        break

    def _find(self, sheet: str, r0: int, c0: int):
        """(レコードの番号, 複数セルのレコードなら列の位置) を返す。なければ None。"""
        a, z = self.sheet_at[sheet]
        depth = 0
        for j in range(a, z + 1):
            rec = self.recs[j]
            if rec.t == T_BOF:
                depth += 1
                continue
            if rec.t == T_EOF:
                depth -= 1
                continue
            if depth != 1 or len(rec.b) < 4:
                continue
            if rec.t in CELLS:
                r, c = struct.unpack_from("<HH", rec.b, 0)
                if (r, c) == (r0, c0):
                    return j, None
            elif rec.t in MULTI:
                r, cf = struct.unpack_from("<HH", rec.b, 0)
                cl = struct.unpack_from("<H", rec.b, len(rec.b) - 2)[0]
                if r == r0 and cf <= c0 <= cl:
                    return j, c0 - cf
        return None

    # ------------------------------------------------------------------ 書き込み
    def _sst_of(self, s: str) -> int:
        if s in self.sst_index:
            return self.sst_index[s]
        idx = self.n_sst + len(self.sst_new)
        self.sst_new.append(s)
        self.sst_index[s] = idx
        return idx

    def _cell_rec(self, r0: int, c0: int, xf: int, value) -> Rec:
        if value is None:
            return Rec(T_BLANK, struct.pack("<HHH", r0, c0, xf), False)
        if isinstance(value, bool):
            return Rec(T_BOOLERR, struct.pack("<HHHBB", r0, c0, xf, 1 if value else 0, 0), False)
        if isinstance(value, ExcelError):
            return Rec(T_BOOLERR, struct.pack("<HHHBB", r0, c0, xf, ERRCODE.get(value.code, 0x2A), 1), False)
        if is_number(value):
            return Rec(T_NUMBER, struct.pack("<HHHd", r0, c0, xf, float(value)), False)
        return Rec(T_LABELSST, struct.pack("<HHHI", r0, c0, xf, self._sst_of(str(value))), False)

    def set_cell(self, sheet: str, r: int, c: int, value, xf: int | None = None) -> None:
        r0, c0 = r - 1, c - 1
        hit = self._find(sheet, r0, c0)
        self.changed = True
        if hit is None:
            self._insert(sheet, r0, c0, self._cell_rec(r0, c0, xf or 15, value))
            return
        j, k = hit
        rec = self.recs[j]
        if k is None:
            old_xf = struct.unpack_from("<H", rec.b, 4)[0]
            new = self._cell_rec(r0, c0, old_xf, value)
            # 数式だったセル（上書きを許したとき）は、続く STRING も取り除く
            if rec.t == T_FORMULA:
                k = j + 1
                if k < len(self.recs) and self.recs[k].t in (0x04BC, 0x0221, 0x0236):
                    raise XlsWriteError("共有数式・配列数式の本体のセルは、ほかのセルの数式の元なので上書きできません")
                if k < len(self.recs) and self.recs[k].t == T_STRING:
                    del self.recs[k]
                    self._shift(sheet, -1)
            self.recs[j] = new
            return
        # MULRK / MULBLANK の途中：前・そのセル・後に分ける
        r, cf = struct.unpack_from("<HH", rec.b, 0)
        cl = struct.unpack_from("<H", rec.b, len(rec.b) - 2)[0]
        width = 6 if rec.t == T_MULRK else 2
        items = [rec.b[4 + width * i:4 + width * (i + 1)] for i in range(cl - cf + 1)]
        xf_here = struct.unpack_from("<H", items[k], 0)[0]
        out = []
        for part, start in ((items[:k], cf), (items[k + 1:], cf + k + 1)):
            if not part:
                out.append(None)
                continue
            if len(part) == 1:
                if rec.t == T_MULRK:
                    x, v = struct.unpack("<HI", part[0])
                    out.append(Rec(T_RK, struct.pack("<HHHI", r, start, x, v), False))
                else:
                    out.append(Rec(T_BLANK, struct.pack("<HHH", r, start, struct.unpack("<H", part[0])[0]), False))
            else:
                out.append(Rec(rec.t, struct.pack("<HH", r, start) + b"".join(part) + struct.pack("<H", start + len(part) - 1), False))
        repl = [x for x in (out[0], self._cell_rec(r0, c0, xf_here, value), out[1]) if x is not None]
        self.recs[j:j + 1] = repl
        self._shift(sheet, len(repl) - 1)  # レコードの数が変わったので、後ろのシートの位置をずらす

    def set_formula_value(self, sheet: str, r: int, c: int, value) -> None:
        hit = self._find(sheet, r - 1, c - 1)
        if hit is None or hit[1] is not None or self.recs[hit[0]].t != T_FORMULA:
            raise XlsWriteError(f"{sheet}!R{r}C{c} の数式のレコードが見つかりません")
        j = hit[0]
        b = bytearray(self.recs[j].b)
        # 結果の文字列（STRING）は、共有数式・配列数式の本体（SHRFMLA など）の後ろに置かれる
        k = j + 1
        while k < len(self.recs) and self.recs[k].t in (0x04BC, 0x0221, 0x0236):
            k += 1
        has_str = k < len(self.recs) and self.recs[k].t == T_STRING
        if isinstance(value, bool):
            b[6:14] = bytes([1, 0, 1 if value else 0, 0, 0, 0, 0xFF, 0xFF])
        elif isinstance(value, ExcelError):
            b[6:14] = bytes([2, 0, ERRCODE.get(value.code, 0x2A), 0, 0, 0, 0xFF, 0xFF])
        elif is_number(value):
            b[6:14] = struct.pack("<d", float(value))
        elif value == "" or value is None:
            b[6:14] = bytes([3, 0, 0, 0, 0, 0, 0xFF, 0xFF])
        else:
            b[6:14] = bytes([0, 0, 0, 0, 0, 0, 0xFF, 0xFF])
        self.recs[j] = Rec(T_FORMULA, bytes(b), False)
        is_text = isinstance(value, str) and value != ""
        if is_text:
            srec = Rec(T_STRING, _xlstr(value), False)
            if has_str:
                if self.recs[k].b != srec.b:
                    self.recs[k] = srec
            else:
                self.recs.insert(k, srec)
                self._shift(sheet, 1)
        elif has_str:
            del self.recs[k]
            self._shift(sheet, -1)
        self.changed = True

    def set_checkbox(self, sheet: str, shape_id: str, checked: bool) -> None:
        self.set_control(sheet, shape_id, {"checked": "Checked" if checked else None})

    def set_control(self, sheet: str, shape_id: str, change: dict) -> None:
        """フォームの部品の状態（チェック・選んだ項目・値・複数選択）のバイトだけを、OBJ（と続く CONTINUE）の中で書き換える。

        change は xlsx の部品の書き込みと同じ形（{"checked": ...}・{"sel": ...}・{"val": ...}・{"multi_sel": [...]}）。
        どのバイトが状態かは obj.py（MS-XLS 2.4.181 Obj）のとおりに読み、長さは変えない。"""
        idx = self._obj_at(sheet, shape_id)
        parts = [self.recs[k].b for k in idx]
        try:
            data = with_state(b"".join(parts), change)
        except ObjError as e:
            raise XlsWriteError(f"{sheet} の部品 {shape_id}: {e}") from None
        at = 0
        for k, part in zip(idx, parts):
            new = data[at:at + len(part)]
            at += len(part)
            if new != part:
                self.recs[k] = Rec(self.recs[k].t, new, False)
                self.changed = True

    def _obj_at(self, sheet: str, shape_id: str) -> list[int]:
        """シートの OBJ のうち FtCmo の id が shape_id のものと、すぐ後ろに続く CONTINUE の、レコードの番号。
        シートの中に埋め込んだグラフ（BOF〜EOF）の中の OBJ は、番号の数え方が別なので見ない。"""
        a, z = self.sheet_at[sheet]
        depth = 0
        for j in range(a, z + 1):
            rec = self.recs[j]
            if rec.t == T_BOF:
                depth += 1
                continue
            if rec.t == T_EOF:
                depth -= 1
                continue
            if depth != 1 or rec.t != T_OBJ or len(rec.b) < 8:
                continue
            ft, _cb, _ot, oid = struct.unpack_from("<HHHH", rec.b, 0)
            if ft == 0x15 and str(oid) == shape_id:
                k = j + 1
                while k <= z and self.recs[k].t == T_CONTINUE:
                    k += 1
                return list(range(j, k))
        raise XlsWriteError(f"{sheet} の部品 {shape_id} が見つかりません")

    def force_recalc(self) -> None:
        """開いたときにExcelがすべて計算し直すようにする（RECALCID を古い番号にする）。"""
        for j, rec in enumerate(self.recs):
            if rec.t == T_RECALCID:
                b = bytearray(rec.b)
                struct.pack_into("<I", b, 4, 0)
                self.recs[j] = Rec(T_RECALCID, bytes(b), False)
                self.changed = True
                return

    def _insert(self, sheet: str, r0: int, c0: int, new: Rec) -> None:
        """セルのレコードがないセル：同じ行の、列の順の位置に差し込む。行のレコードもなければ作る。"""
        a, z = self.sheet_at[sheet]
        last_cell_of_row = None
        first_after = None
        row_rec = None
        prev_rows_last_cell = None
        depth = 0
        for j in range(a, z + 1):
            rec = self.recs[j]
            if rec.t == T_BOF:
                depth += 1
                continue
            if rec.t == T_EOF:
                depth -= 1
                continue
            if depth != 1:
                continue
            if rec.t == T_ROW and struct.unpack_from("<H", rec.b, 0)[0] == r0:
                row_rec = j
            if rec.t in CELLS or rec.t in MULTI:
                r, c = struct.unpack_from("<HH", rec.b, 0)
                if r == r0:
                    if c < c0:
                        last_cell_of_row = j
                    elif first_after is None:
                        first_after = j
                elif r < r0:
                    prev_rows_last_cell = j
        if first_after is not None:
            at = first_after
        elif last_cell_of_row is not None:
            at = last_cell_of_row + 1
            while at < len(self.recs) and self.recs[at].t == T_STRING:
                at += 1
        else:
            if row_rec is None:
                self._insert_row(sheet, r0)
                a, z = self.sheet_at[sheet]
                return self._insert(sheet, r0, c0, new)
            # この行の最初のセル：前の行の最後のセルの次（なければ ROW の並びの後）
            if prev_rows_last_cell is not None and prev_rows_last_cell > row_rec:
                at = prev_rows_last_cell + 1
                while at < len(self.recs) and self.recs[at].t == T_STRING:
                    at += 1
            else:
                at = row_rec + 1
                while at < len(self.recs) and self.recs[at].t == T_ROW:
                    at += 1
        self.recs.insert(at, new)
        self._shift(sheet, 1)
        self._row_span(sheet, r0, c0)

    def _insert_row(self, sheet: str, r0: int) -> None:
        a, z = self.sheet_at[sheet]
        at = None
        for j in range(a, z + 1):
            rec = self.recs[j]
            if rec.t == T_ROW:
                if struct.unpack_from("<H", rec.b, 0)[0] > r0:
                    at = j
                    break
                at = j + 1
        if at is None:
            # 行のレコードがひとつもない：DIMENSIONS の後に置く
            at = next((j + 1 for j in range(a, z + 1) if self.recs[j].t == T_DIMENSIONS), a + 1)
        row = struct.pack("<HHHHHHI", r0, 0, 0, 255, 0, 0, 0x100)
        self.recs.insert(at, Rec(T_ROW, row, False))
        self._shift(sheet, 1)

    def _row_span(self, sheet: str, r0: int, c0: int) -> None:
        a, z = self.sheet_at[sheet]
        for j in range(a, z + 1):
            rec = self.recs[j]
            if rec.t == T_ROW and struct.unpack_from("<H", rec.b, 0)[0] == r0:
                rw, mic, mac = struct.unpack_from("<HHH", rec.b, 0)
                if mic == mac == 0:
                    mic, mac = c0, c0 + 1
                else:
                    mic, mac = min(mic, c0), max(mac, c0 + 1)
                b = bytearray(rec.b)
                struct.pack_into("<HH", b, 2, mic, mac)
                if bytes(b) != rec.b:
                    self.recs[j] = Rec(T_ROW, bytes(b), False)
                return

    def _shift(self, sheet: str, n: int) -> None:
        names = list(self.sheet_at)
        hit = False
        for name in names:
            a, z = self.sheet_at[name]
            if name == sheet:
                self.sheet_at[name] = (a, z + n)
                hit = True
            elif hit:
                self.sheet_at[name] = (a + n, z + n)

    # ------------------------------------------------------------------ 組み立て
    def build(self) -> bytes:
        recs = self._with_sst()
        for _round in range(6):
            sizes = [len(r.b) for r in recs]
            self._fix_blocks(recs, None)
            self._fix_extsst(recs, None)
            self._fix_boundsheets(recs)
            if [len(r.b) for r in recs] == sizes:
                # 長さが変わらなければ、位置はもう動かない（念のため、もう一度だけ位置を合わせる）
                self._fix_blocks(recs, None)
                self._fix_extsst(recs, None)
                self._fix_boundsheets(recs)
                break
        out = bytearray()
        for rec in recs:
            if len(rec.b) > MAXREC:
                raise XlsWriteError("レコードが長すぎます")
            out += struct.pack("<HH", rec.t, len(rec.b)) + rec.b
        return bytes(out)

    def _fix_boundsheets(self, recs: list[Rec]) -> None:
        """シートの BOF の位置（BOUNDSHEET の lbPlyPos）を、いまの並びに合わせる。"""
        offs = []
        pos = 0
        for rec in recs:
            offs.append(pos)
            pos += 4 + len(rec.b)
        bofs = []
        depth = 0
        for i, rec in enumerate(recs):
            if rec.t == T_BOF:
                if depth == 0 and i > 0:
                    bofs.append(offs[i])
                depth += 1
            elif rec.t == T_EOF:
                depth -= 1
        k = 0
        for i, rec in enumerate(recs):
            if rec.t == T_BOUNDSHEET and k < len(bofs):
                if struct.unpack_from("<I", rec.b, 0)[0] != bofs[k]:
                    b = bytearray(rec.b)
                    struct.pack_into("<I", b, 0, bofs[k])
                    recs[i] = Rec(T_BOUNDSHEET, bytes(b), False)
                k += 1

    def _with_sst(self) -> list[Rec]:
        """SST に文字列を足す（足すときは最後の CONTINUE の後ろに、決まりどおりに分けて続ける）。"""
        recs = list(self.recs)
        if not self.sst_new:
            return recs
        i = next(k for k, r in enumerate(recs) if r.t == T_SST)
        j = i + 1
        while j < len(recs) and recs[j].t == T_CONTINUE:
            j += 1
        parts = [bytearray(recs[k].b) for k in range(i, j)]
        total, unique = struct.unpack_from("<II", parts[0], 0)
        struct.pack_into("<II", parts[0], 0, total + len(self.sst_new), unique + len(self.sst_new))
        for s in self.sst_new:
            try:
                raw, high = s.encode("latin-1"), 0
            except UnicodeEncodeError:
                raw, high = s.encode("utf-16-le"), 1
            head = struct.pack("<HB", len(s), high)
            if len(parts[-1]) + len(head) + (2 if high else 1) > MAXREC:
                parts.append(bytearray())
            parts[-1] += head
            width = 2 if high else 1
            data = raw
            while data:
                room = (MAXREC - len(parts[-1])) // width * width
                if room <= 0:
                    parts.append(bytearray([high]))  # 文字の途中で切れたら、続きの先頭に文字の幅の印
                    continue
                parts[-1] += data[:room]
                data = data[room:]
        self.sst_parts = parts
        new = [Rec(T_SST, bytes(parts[0]), False)] + [Rec(T_CONTINUE, bytes(p), False) for p in parts[1:]]
        # 元の SST/CONTINUE と同じ中身の部分は「元のまま」として扱う
        for k in range(min(j - i, len(new))):
            if new[k].b == recs[i + k].b:
                new[k] = recs[i + k]
        return recs[:i] + new + recs[j:]

    def _fix_extsst(self, recs: list[Rec], offs_unused) -> None:
        """EXTSST（SST の索引）を、文字列の本当の位置に合わせて作り直す。"""
        k = next((i for i, r in enumerate(recs) if r.t == T_EXTSST), None)
        s = next((i for i, r in enumerate(recs) if r.t == T_SST), None)
        if k is None or s is None:
            return
        offs = []
        pos = 0
        for rec in recs:
            offs.append(pos)
            pos += 4 + len(rec.b)
        dsst = struct.unpack_from("<H", recs[k].b, 0)[0] or 8
        # SST の文字列の位置をたどる
        parts = [s]
        j = s + 1
        while j < len(recs) and recs[j].t == T_CONTINUE:
            parts.append(j)
            j += 1
        from .book import Seg

        seg = Seg([recs[p].b for p in parts])
        seg.u32()
        unique = seg.u32()
        entries = []
        for n in range(unique):
            if n % dsst == 0:
                rec_idx = parts[seg.pi]
                within = seg.pos + 4  # レコードの頭（4バイト）からの位置
                entries.append(struct.pack("<IHH", offs[rec_idx] + within, within, 0))
            cch = seg.u16()
            flags = seg.u8()
            runs = seg.u16() if flags & 0x08 else 0
            ext = seg.u32() if flags & 0x04 else 0
            seg.chars(cch, bool(flags & 0x01))
            if runs:
                seg.read(4 * runs)
            if ext:
                seg.read(ext)
        new = struct.pack("<H", dsst) + b"".join(entries)
        if new != recs[k].b:
            recs[k] = Rec(T_EXTSST, new, False)

    def _fix_blocks(self, recs: list[Rec], offs_unused) -> None:
        """DBCELL（行のまとまりの索引）と INDEX を、レコードの本当の位置に合わせて作り直す。"""
        for _round in range(2):  # 長さの変わらない書き直しなので、2回で位置が落ち着く
            offs = []
            pos = 0
            for rec in recs:
                offs.append(pos)
                pos += 4 + len(rec.b)
            depth = 0
            sheet_start = None
            dbcells: list[int] = []
            index_at = None
            defcol_at = None
            block_rows: list[int] = []
            cells_after: list[int] = []
            for i, rec in enumerate(recs):
                if rec.t == T_BOF:
                    depth += 1
                    if depth == 1:
                        sheet_start, dbcells, index_at, defcol_at = i, [], None, None
                        block_rows, cells_after = [], []
                    continue
                if rec.t == T_EOF:
                    if depth == 1 and index_at is not None:
                        self._write_index(recs, index_at, offs, dbcells, defcol_at)
                    depth -= 1
                    continue
                if depth != 1 or sheet_start is None:
                    continue
                if rec.t == T_INDEX:
                    index_at = i
                elif rec.t == T_DEFCOLWIDTH:
                    defcol_at = i
                elif rec.t == T_ROW:
                    if cells_after:
                        block_rows, cells_after = [], []
                    block_rows.append(i)
                elif rec.t in CELLS or rec.t in MULTI or rec.t == T_STRING:
                    cells_after.append(i)
                elif rec.t == T_DBCELL:
                    if block_rows:
                        self._write_dbcell(recs, i, offs, block_rows, cells_after)
                    dbcells.append(i)
                    block_rows, cells_after = [], []

    def _write_dbcell(self, recs, i, offs, rows, cells) -> None:
        """Excelと同じ形：ROW ごとに、その行の最初のセルの位置（セルのない行は、次のセルの位置）を、前の行からのずれで書く。"""
        if all(recs[j].orig for j in rows + cells):
            return  # このまとまりは変えていない（中の位置はすべて相対なので、そのまま）
        firsts: dict[int, int] = {}
        order = []
        for j in cells:
            rec = recs[j]
            if rec.t == T_STRING:
                continue
            r = struct.unpack_from("<H", rec.b, 0)[0]
            order.append((r, j))
            firsts.setdefault(r, j)
        pos_of = []
        for jr in rows:
            r = struct.unpack_from("<H", recs[jr].b, 0)[0]
            if r in firsts:
                pos_of.append(offs[firsts[r]])
            else:
                nxt = next((j for rr, j in order if rr > r), None)
                pos_of.append(offs[nxt] if nxt is not None else offs[i])
        ref = offs[rows[1]] if len(rows) > 1 else offs[rows[0]] + 4 + len(recs[rows[0]].b)
        offsets = []
        for p_ in pos_of:
            offsets.append(p_ - ref)
            ref = p_
        new = struct.pack("<I", offs[i] - offs[rows[0]]) + b"".join(struct.pack("<H", max(0, min(o, 0xFFFF))) for o in offsets)
        if new != recs[i].b:
            recs[i] = Rec(T_DBCELL, new, False)

    def _write_index(self, recs, i, offs, dbcells, defcol_at) -> None:
        b = bytearray(recs[i].b)
        n_old = (len(b) - 16) // 4
        if defcol_at is not None:
            struct.pack_into("<I", b, 12, offs[defcol_at])
        body = b"".join(struct.pack("<I", offs[d]) for d in dbcells)
        new = bytes(b[:16]) + body
        if n_old != len(dbcells) or new != recs[i].b:
            recs[i] = Rec(T_INDEX, new, False)
