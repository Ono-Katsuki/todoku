"""大きな .xls（BIFF8・BIFF5）を分けて読む。

- ブックの部分（globals）だけを読み、シートの位置（BoundSheet）を得る。シートごとに記録の頭だけをたどり、行ごとの記録の位置の表を作る（中身は読まない）
- 共有文字列（SST）：使うときに読む。はじめて使うときに 1 度だけたどって、64 個ごとの位置を控える。CONTINUE のまたぎも窓の読み手と同じ
- 窓：頼まれた行の記録と、シートの設定の記録（列の幅・結合・入力規則など）だけの小さなストリームを作り、いままでの XlsWorkbook と窓の読み手で読む。
  番地はファイルの記録のとおりなので、窓の字・欄の id は全体を読んだときと同じ
- 読む範囲・続きの印・欄は XlsxLazy と同じ（行の索引の形をそろえる）
- 書く：いままでの api.fill（ブックを全部読む）
"""

from __future__ import annotations

import bisect
import struct
from array import array

from .. import limits
from .base import LazyDoc, Page, RangeError, comma  # noqa: F401
from .cfbfile import CfbFile
from .xlsxdoc import XlsxLazy

CONTINUE = 0x003C
ROW_RECS = frozenset({0x0203, 0x027E, 0x00BD, 0x00FD, 0x0204, 0x00D6, 0x0201, 0x00BE, 0x0205, 0x0006, 0x0208})
VALUE_RECS = frozenset({0x0203, 0x027E, 0x00BD, 0x00FD, 0x0204, 0x00D6, 0x0205, 0x0006})
SKIP_RECS = frozenset({0x00D7, 0x020B})  # DBCELL・INDEX（位置の控え。窓の読み手は使わない）


def walk_records(st, start: int = 0, stop: int | None = None):
    """(位置, 種類, 中身の長さ)。頭だけを読む（1 MB ずつ控えて読む）。"""
    pos = start
    end = st.size if stop is None else stop
    buf = b""
    buf_at = 0
    while pos + 4 <= end:
        if pos < buf_at or pos + 16 > buf_at + len(buf):
            buf = st.read(pos, pos + (1 << 20))
            buf_at = pos
            if len(buf) < 4:
                return
        t, ln = struct.unpack_from("<HH", buf, pos - buf_at)
        yield pos, t, ln, buf, buf_at
        pos += 4 + ln


class SheetRows:
    """1 枚のシートの索引。形は xlsxdoc.SheetIndex に合わせる（rows・offs・locate）。"""

    def __init__(self, st, lb: int, name: str):
        self.st = st
        self.lb = lb
        self.m = None
        self.tail = b""  # xlsx の入力規則などの部分（使わない）
        runs_row = array("I")  # 記録のかたまり（連続した同じ行の記録）
        runs_at = array("Q")
        runs_end = array("Q")
        self.meta: list[tuple[int, int]] = []  # 設定の記録の位置 (始め, 終わり)
        self.cells = self.values = self.formulas = 0
        self.max_col = 0
        depth = 0
        last_kind = None  # 直前の記録の行き先（"row" か "meta"）。CONTINUE・STRING は直前に続く
        cur_row = -1
        self.bof = None
        for pos, t, ln, buf, buf_at in walk_records(st, lb):
            end = pos + 4 + ln
            if t == 0x0809:
                depth += 1
                if depth == 1:
                    self.bof = (pos, end)
                    continue
                self._add_meta(pos, end)
                last_kind = "meta"
                continue
            if t == 0x000A:
                depth -= 1
                if depth == 0:
                    self.end = end
                    break
                self._add_meta(pos, end)
                last_kind = "meta"
                continue
            if depth > 1:
                self._add_meta(pos, end)
                last_kind = "meta"
                continue
            if t in SKIP_RECS:
                last_kind = None
                continue
            if t in ROW_RECS and pos + 8 - buf_at <= len(buf):
                row = struct.unpack_from("<H", buf, pos + 4 - buf_at)[0] + 1
                if t != 0x0208:
                    self.cells += 1
                    col = struct.unpack_from("<H", buf, pos + 6 - buf_at)[0] + 1 if pos + 8 - buf_at <= len(buf) else 0
                    if col > self.max_col:
                        self.max_col = col
                    if t in VALUE_RECS:
                        self.values += 1
                    if t == 0x0006:
                        self.formulas += 1
                if row == cur_row and last_kind == "row" and runs_end[-1] == pos:
                    runs_end[-1] = end
                else:
                    runs_row.append(row)
                    runs_at.append(pos)
                    runs_end.append(end)
                cur_row = row
                last_kind = "row"
                continue
            if t in (CONTINUE, 0x0207) and last_kind == "row":  # 文字の続き・直前の数式の結果の文字列
                if runs_end[-1] == pos:
                    runs_end[-1] = end
                else:
                    runs_row.append(cur_row)
                    runs_at.append(pos)
                    runs_end.append(end)
                continue
            self._add_meta(pos, end)
            last_kind = "meta"
        order = sorted(range(len(runs_row)), key=lambda k: (runs_row[k], runs_at[k]))  # 行の順（ファイルの中で順に並んでいなくても）
        self.run_row = array("I", (runs_row[k] for k in order))
        self.run_at = array("Q", (runs_at[k] for k in order))
        self.run_end = array("Q", (runs_end[k] for k in order))
        self.rows = array("I")
        self.offs = array("Q")
        at = 0
        for k, r in enumerate(self.run_row):
            if not self.rows or self.rows[-1] != r:
                self.rows.append(r)
                self.offs.append(at)
            at += self.run_end[k] - self.run_at[k]
        self.data_end = at

    def _add_meta(self, a: int, b: int) -> None:
        if self.meta and self.meta[-1][1] == a:
            self.meta[-1] = (self.meta[-1][0], b)
        else:
            self.meta.append((a, b))

    def locate(self, a: int, b: int) -> tuple[int, int]:
        return bisect.bisect_left(self.rows, a), bisect.bisect_right(self.rows, b)

    def record_bytes(self, a: int, b: int) -> bytes:
        """行 a〜b の記録（ファイルの順は保つ）。"""
        i = bisect.bisect_left(self.run_row, a)
        j = bisect.bisect_right(self.run_row, b)
        idx = sorted(range(i, j), key=lambda k: self.run_at[k])
        return b"".join(self.st.read(self.run_at[k], self.run_end[k]) for k in idx)

    def meta_bytes(self) -> bytes:
        return b"".join(self.st.read(a, b) for a, b in self.meta)


class LazySst:
    """共有文字列。shared_strings・phonetic・sst_runs の代わりになる（使うときに読む）。"""

    STRIDE = 64

    def __init__(self, st, parts: list[tuple[int, int]]):
        self.st = st
        self.parts = parts  # 記録の中身の位置 (始め, 終わり)
        self.n = 0
        self.marks: array | None = None
        self._cache: dict[int, tuple] = {}
        self.phonetic = _Side(self, 1)
        self.sst_runs = _Side(self, 2)

    def _scan(self) -> None:
        from ..xls.book import Seg

        if self.marks is not None:
            return
        seg = Seg(_Parts(self))
        seg.u32()
        self.n = seg.u32()
        self.marks = array("Q")
        for i in range(self.n):
            if i % self.STRIDE == 0:
                self.marks.append((seg.pi << 20) | seg.pos)
            try:
                self._one(seg, decode=False)
            except ValueError:
                self.n = i
                break
            if i % 65536 == 0:
                limits.progress("共有文字列の位置を調べています", i, self.n)

    def _one(self, seg, decode: bool = True):
        from ..xls.book import _phonetic, cp_runs

        cch = seg.u16()
        flags = seg.u8()
        runs = seg.u16() if flags & 0x08 else 0
        ext = seg.u32() if flags & 0x04 else 0
        s = seg.chars(cch, bool(flags & 0x01))
        rl = None
        ph = None
        if runs:
            rb = seg.read(4 * runs)
            if decode:
                rl = cp_runs(s, [struct.unpack_from("<HH", rb, 4 * k) for k in range(runs)])
        if ext:
            eb = seg.read(ext)
            if decode:
                ph = _phonetic(s, eb)
        return s, ph, rl

    def get(self, i: int) -> tuple:
        from ..xls.book import Seg

        self._scan()
        got = self._cache.get(i)
        if got is not None:
            return got
        if len(self._cache) > 200_000:
            self._cache.clear()
        k = i - i % self.STRIDE
        mark = self.marks[k // self.STRIDE]
        seg = Seg(_Parts(self))
        seg.pi, seg.pos = mark >> 20, mark & 0xFFFFF
        for x in range(k, min(k + self.STRIDE, self.n)):
            one = self._one(seg)
            self._cache[x] = one
        return self._cache[i]

    def __len__(self) -> int:
        self._scan()
        return self.n

    def __getitem__(self, i):
        if isinstance(i, slice):
            return [self[k] for k in range(*i.indices(len(self)))]
        return self.get(i)[0]

    def __iter__(self):
        return (self[i] for i in range(len(self)))


class _Parts:
    """Seg に渡す、記録の中身の並び（使うときに読む。いくつかは控える）。"""

    def __init__(self, sst: LazySst):
        self.s = sst
        self.cache: dict[int, bytes] = {}

    def __len__(self) -> int:
        return len(self.s.parts)

    def __bool__(self) -> bool:
        return bool(self.s.parts)

    def __getitem__(self, i: int) -> bytes:
        got = self.cache.get(i)
        if got is None:
            a, b = self.s.parts[i]
            got = self.s.st.read(a, b)
            if len(self.cache) > 16:
                self.cache.clear()
            self.cache[i] = got
        return got


class _Side:
    """ふりがな（1）・字ごとの書式（2）：番号 → 値。"""

    def __init__(self, s: LazySst, k: int):
        self.s, self.k = s, k

    def __contains__(self, i) -> bool:
        return 0 <= i < len(self.s) and self.s.get(i)[self.k] is not None

    def __getitem__(self, i):
        return self.s.get(i)[self.k]

    def get(self, i, default=None):
        v = self.s.get(i)[self.k] if 0 <= i < len(self.s) else None
        return default if v is None else v

    def __setitem__(self, k, v):
        pass


class _MiniCfb:
    """小さなストリームだけの複合ファイルに見せる（XlsWorkbook に渡す）。"""

    def __init__(self, name: str, data: bytes, real):
        self._name, self._data, self._real = name, data, real

    def __getattr__(self, name):
        return getattr(self._real, name)

    def read(self, name: str) -> bytes:
        return self._data if name == self._name else self._real.read(name)


def _rec(t: int, body: bytes) -> bytes:
    return struct.pack("<HH", t, len(body)) + body


class XlsLazy(XlsxLazy):
    kind = "xls"
    streams = False

    def __init__(self, path: str, name: str, cf: CfbFile, stream: str, password: str | None = None):
        self.path = path
        self.name = name
        self.cf = cf
        self.st = cf.stream(stream)
        self.stream_name = stream
        self.password = password
        self.ratio = 0.25
        self._meta = None
        self._globals()
        self._fx = None
        self.est_chars = sum(d["index"].values for d in self.sheet_meta if d["index"] is not None) * 16

    def close(self) -> None:
        self.cf.close()

    @property
    def size(self) -> int:
        return self.cf.size

    def _globals(self) -> None:
        keep = []
        bound = []
        sst_parts: list[tuple[int, int]] = []
        in_sst = False
        pos0 = None
        for pos, t, ln, buf, buf_at in walk_records(self.st):
            if pos0 is None:
                if t != 0x0809:
                    raise ValueError("Excel のブックの頭（BOF）がありません")
                pos0 = pos
            end = pos + 4 + ln
            if t == 0x002F:
                raise ValueError("パスワードで保護されたブック")  # いままでの道で理由を伝える
            if t == 0x00FC:
                in_sst = True
                sst_parts.append((pos + 4, end))
                continue
            if t == CONTINUE and in_sst:
                sst_parts.append((pos + 4, end))
                continue
            in_sst = False
            body = self.st.read(pos + 4, end)
            if t == 0x0085:
                lb = struct.unpack_from("<I", body, 0)[0]
                from ..xls.book import xlstring

                try:
                    nm = xlstring(body, 6, cch_size=1)[0]
                except Exception:  # noqa: BLE001  BIFF5 の形：いままでの読み手が読む
                    n = body[6]
                    nm = body[7:7 + n].decode("cp932", "replace")
                bound.append((len(keep), lb, body[4] & 3, body[5], nm))
            keep.append((t, body))
            if t == 0x000A:
                break
        self.keep = keep
        self.bound = bound
        self.sst = LazySst(self.st, sst_parts) if sst_parts else None
        meta = []
        for _k, lb, state, dt, nm in bound:
            st = {0: "visible", 1: "hidden", 2: "veryHidden"}.get(state, "visible")
            d = {"name": nm, "path": f"#{lb}", "state": st, "index": None}
            if dt == 0:
                limits.progress(f"シート「{nm}」の記録の位置を調べています")
                d["index"] = SheetRows(self.st, lb, nm)
            else:
                d["state"] = {1: "macrosheet", 2: "chartsheet", 6: "vbmodule"}.get(dt, f"type{dt}")
            meta.append(d)
        self._meta = meta
        try:
            for d, sh in zip(meta, self.book_of({}).sheets):  # シートの名前は窓の読み手のとおり（BIFF5 の文字コードなど）
                d["name"] = sh.name
        except ValueError:
            raise

    @property
    def strings(self):
        return self.sst

    def _load_meta(self):
        return self._meta

    @property
    def sheet_meta(self):
        return self._meta

    def mini_stream(self, wanted: dict[int, list[tuple[int, int]]]) -> bytes:
        """小さなストリーム：ブックの部分（BoundSheet の位置を直す）と、頼まれたシートの頼まれた行だけ。"""
        by_keep = {b[0]: i for i, b in enumerate(self.bound)}
        # 2 回まわす：1 回目で長さを決める
        globals_len = sum(4 + len(b) for _t, b in self.keep)
        empty_bof = None
        sheets: list[bytes] = []
        for si, d in enumerate(self._meta):
            ix = d["index"]
            if ix is None:
                sheets.append(b"")
                continue
            bof = self.st.read(*ix.bof)
            empty_bof = empty_bof or bof
            if si in wanted:
                rows = b"".join(ix.record_bytes(a, b) for a, b in _merge(wanted[si]))
                sheets.append(bof + ix.meta_bytes() + rows + _rec(0x000A, b""))
            else:
                sheets.append(bof + _rec(0x000A, b""))
        out = bytearray()
        at = globals_len
        lbs = []
        for s in sheets:
            lbs.append(at)
            at += len(s)
        for k, (t, body) in enumerate(self.keep):
            if k in by_keep:
                si = by_keep[k]
                lb = lbs[si] if self._meta[si]["index"] is not None else 0
                body = struct.pack("<I", lb) + body[4:]
            out += _rec(t, body)
        for s in sheets:
            out += s
        return bytes(out)

    def book_of(self, wanted: dict[int, list[tuple[int, int]]]):
        from ..xls.book import XlsWorkbook

        data = self.mini_stream(wanted)
        return XlsWorkbook(self.cf.fb, self.name, cfb=_MiniCfb(self.stream_name, data, self.cf.cfb), strings=self.sst, stream=self.stream_name)

    def rows_book(self, wanted):
        return self.book_of(wanted)

    def _sheet_doc(self, si: int, a: int, b: int) -> dict:
        from ..xlsx.window import _MEMO, _Options
        from ..xlsx.window import _sheet_doc as xw_sheet

        ix = self._meta[si]["index"]
        first = ix.rows[0] if ix is not None and ix.rows else None
        wanted = [(a, b)]
        if first is not None and first < a:
            wanted.insert(0, (first, first))
        wb = self.book_of({si: wanted})
        sh = wb.sheets[si]
        sh.__dict__[_MEMO] = {}
        try:
            return xw_sheet(wb, sh, _Options(wb))
        finally:
            sh.__dict__.pop(_MEMO, None)

    def _no_rows_lines(self, si: int) -> list[str]:
        from ..xlsx.window import _md_sheet, build

        doc = build(self.book_of({}))
        return _md_sheet(doc["sheets"][si])

    def _blanks(self, si: int, i: int, k: int, ix) -> list[str]:
        return []  # .xls は空のマスの番地を、窓の欄（入力規則・空のマス）で示す

    def overview(self) -> dict:
        sheets = []
        for d in self._meta:
            one = {"name": d["name"], "state": d["state"]}
            ix = d["index"]
            if ix is not None:
                one.update({"rows": len(ix.rows), "first_row": ix.rows[0] if ix.rows else None, "last_row": ix.rows[-1] if ix.rows else None,
                            "cells": ix.cells, "cells_with_value": ix.values, "formulas": ix.formulas})
            sheets.append(one)
        return {"sheets": sheets, "ranges": "sheet と rows（行の範囲）", "bytes": self.size}

    def overview_text(self) -> str:
        out = [f"# {self.name}", "", f"- 古い Excel のブック（{self.size:,} バイト）。全体を一度には出しません。read_window に sheet と rows（\"1-500\"）を渡して読みます"
               "（cursor をたどれば、すべてのシートのすべての行を読めます。窓を作るのが上限で止まるときは raw: true で字のまま読めます）。list_fields も同じ範囲です。",
               "", "## シート", "", "| シート | 状態 | 行（範囲） | 値のあるセル | 数式 |", "|---|---|---|---|---|"]
        for s in self.overview()["sheets"]:
            if "rows" in s and s["rows"]:
                out.append(f"| {s['name']} | {s['state']} | {s['rows']:,} 行（{s['first_row']}〜{s['last_row']}） | {s['cells_with_value']:,} | {s['formulas']:,} |")
            else:
                out.append(f"| {s['name']} | {s['state']} | セルのないシート | | |")
        return "\n".join(out) + "\n"

    # ------------------------------------------------------------------ 段階・検査
    def grade(self):
        from ..grade import Reason, _final

        forms = sum(d["index"].formulas for d in self._meta if d["index"] is not None)
        total = sum(d["index"].cells for d in self._meta if d["index"] is not None)
        reasons = [Reason("A", "大きなブック", f"シート {len(self._meta)} 枚・セル {total:,}・数式 {forms:,}。記録の並びは全部調べました")]
        write = [Reason("B", "欄を範囲ごとに見つける", "大きなブックなので、欄は頼まれた行の範囲ごとに見つけます。どのセルにも番地で書けます")]
        g = _final(reasons, write)
        g.deferred = ["未確認：全部のセルの値を窓に出すこと（audit_document で全部の行を確かめます）"]
        return g

    def audit_begin(self) -> dict:
        return {"sheet": 0, "next_i": 0, "cells": 0, "problems": [], "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        import os

        if state["sheet"] == 0 and state["next_i"] == 0 and os.path.getsize(self.path) <= limits.MAX_FILE // 8 and not state.get("cells") and not state.get("skip_whole"):
            from .. import api

            au = api.audit_whole(self.path, self.password, None)
            state["checks"], state["full"], state["done"] = au.get("checks", []), au, True
            return state
        meta = self._meta
        si = state["sheet"]
        if si < len(meta):
            d = meta[si]
            ix = d["index"]
            if ix is None or state["next_i"] >= len(ix.rows):
                state["sheet"] = si + 1
                state["next_i"] = 0
                return state
            i = state["next_i"]
            k = min(len(ix.rows), i + 2048)
            ra, rb = ix.rows[i], ix.rows[k - 1]
            wb = self.book_of({si: [(ra, rb)]})
            sh = wb.sheets[si]
            # 窓の読み手とは別に、記録からセルの番地を拾って、窓の読み手が読んだことを確かめる
            want = _value_cells(ix.record_bytes(ra, rb))
            missing = [c for c in want if c not in sh.cells]
            state["cells"] += len(want)
            for r_, c_ in missing[:5]:
                if len(state["problems"]) < 30:
                    state["problems"].append(f"{d['name']}!{_a1(r_, c_)}：記録に値があるのに、窓の読み手が読んでいません")
            state["next_i"] = k
            limits.progress(f"シート「{d['name']}」を確かめています", k, len(ix.rows))
            return state
        state["checks"] = [{"name": "読み落とし（セルの値）", "ok": not state["problems"], "agent": False, "items": list(state["problems"]),
                            "detail": f"全部のシートの記録をたどり、値のあるセル {state['cells']:,} か所を、窓の読み手が読んだことを確かめました"}]
        state["still_deferred"] = ["未確認：書き込みの試し（大きすぎるので検査では試していません。fill_document のときに証明します）"]
        state["done"] = True
        return state


def _a1(r: int, c: int) -> str:
    from ..formula.parser import addr

    return addr(r, c)


def _value_cells(data: bytes) -> list[tuple[int, int]]:
    """行の記録の、値のあるセルの番地（1 から数える）。"""
    out = []
    pos = 0
    n = len(data)
    while pos + 4 <= n:
        t, ln = struct.unpack_from("<HH", data, pos)
        p = pos + 4
        if t in (0x0203, 0x027E, 0x00FD, 0x0204, 0x00D6, 0x0205, 0x0006) and ln >= 4:
            r, c = struct.unpack_from("<HH", data, p)
            out.append((r + 1, c + 1))
        elif t == 0x00BD and ln >= 6:
            r, c0 = struct.unpack_from("<HH", data, p)
            out += [(r + 1, c0 + 1 + k) for k in range((ln - 6) // 6)]
        pos = p + ln
    return out


def _merge(spans):
    out = []
    for a, b in sorted(spans):
        if out and a <= out[-1][1] + 1:
            out[-1] = (out[-1][0], max(out[-1][1], b))
        else:
            out.append((a, b))
    return out
