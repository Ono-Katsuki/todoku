"""大きな Access のデータベース（.mdb・.accdb）を分けて読む。上限は 1 回に返す量だけで、どの表のどの行にも届く。

- 開くとき：表の一覧と列の定義を読み、表ごとに、使うページと、ページごとの行の数の表だけを作る（ファイルの大きさに比例する小さな配列）
- 読むとき：頼まれた行のあるページだけを解く（ファイルは pread で読むので、常駐メモリは増えない。解いたページの控えは Jet.page の分だけ）
- シート 1 枚が表 1 つ。1 行目が列の名前、2 行目からレコード。行の番号はレコードの順で、読み方によらず同じ（消えた行は数えず、形を確かめられない行も番号を飛ばさない）
- 窓の行の書き方は legacy/jet.py の record_text（NULL の列・すべて NULL の行・前後の空白をそのまま示す）
- 表の取り出し（read_table）は rows_book の小さなブックで、いままでの処理を通す
- 書かない（読むだけ）。検査（audit_document）は全部の行を解いて、形を確かめられない行を数える
"""

from __future__ import annotations

import bisect
import re
import struct
from array import array

from .. import limits
from ..legacy.jet import (GUIDE, LEGEND, BadRow, Jet, JetError, RecordBook, _display_order, cell_text, check_password, column_notes,
                          database_notes, is_jet, open_catalog, record_text, stated_note)
from .base import LazyDoc, Page, RangeError, comma, parse_range
from .textstream import Source

CHARS_PER_CELL = 16  # 窓の字の見込み（XlsLazy と同じ。メモのある表は、小さければ測る）
JET_NAMES = (".mdb", ".accdb", ".mde", ".accde")


class _FileBlob:
    """ファイルを、メモリに読まずに bytes のように切り出す（Jet に渡す）。"""

    def __init__(self, src: Source):
        self.src = src

    def __len__(self) -> int:
        return self.src.size

    def __getitem__(self, k):
        if isinstance(k, slice):
            a, b, _step = k.indices(self.src.size)
            return self.src.read(a, b) if a < b else b""
        return self.src.read(k, k + 1)[0]


class TableIndex:
    """1 つの表の索引。rows はシートの行の番号（1 が列の名前）。"""

    def __init__(self, db: Jet, name: str, tdef: int):
        self.name = name
        self.tdef = tdef
        self.cols, d = db.table_def(tdef)
        self.order = _display_order(self.cols)
        self.stated = struct.unpack_from("<I", d, db.f.tab_rows)[0]
        self.pages = array("I")
        self.before = array("Q")  # そのページより前の行の数
        self.outside: list[int] = []  # ファイルの外のページ
        n = 0
        for pg in db.usage_pages(d):
            if pg == 0:
                continue
            if pg >= db.page_count:
                self.outside.append(pg)
                continue
            live = db.count_live(pg, tdef)
            if live:
                self.pages.append(pg)
                self.before.append(n)
                n += live
        self.records = n
        self.rows = range(1, n + 2)
        self.max_col = len(self.order)
        self.notes = column_notes(name, self.cols)
        if self.outside:
            self.notes.append(f"表「{name}」：使うページのうち {len(self.outside):,} ページはファイルの外です（例：ページ {self.outside[0]}）。そのページの行は読めません。{GUIDE}")
        if not self.outside and self.stated != n:
            self.notes.append(stated_note(name, self.stated, n))

    def locate(self, a: int, b: int) -> tuple[int, int]:
        return bisect.bisect_left(self.rows, a), bisect.bisect_right(self.rows, b)

    def has_memo(self) -> bool:
        return any(c.type == 12 for c in self.cols)


class JetLazy(LazyDoc):
    kind = "mdb"
    streams = True  # 書き込みは、丸ごと読まずに断る（fill）

    def __init__(self, path: str, name: str, password: str | None = None, blob=None):
        self.path, self.name, self.password = path, name, password
        self.src = None
        if blob is None:
            self.src = Source(path)
            blob = _FileBlob(self.src)
            if not is_jet(blob):
                self.close()
                raise JetError("Access のデータベースの見出しではありません")
        self.size = len(blob)
        state = check_password(blob, password)
        if state:
            from ..api import PasswordNeeded, PasswordWrong

            self.close()
            if state == "need":
                raise PasswordNeeded("このデータベースはパスワードで保護されています")
            raise PasswordWrong("パスワードが違います")
        self.db = Jet(blob)
        tables, linked, queries = open_catalog(self.db)
        self.linked, self.queries = linked, queries
        self.notes = database_notes(self.db, linked, queries)
        self._meta: list[dict] = []
        for tname, pg in sorted(tables, key=lambda x: x[0].lower()):
            d = {"name": tname[:64] or f"表{pg}", "table": tname, "state": "visible", "records": True, "index": None, "error": None}
            limits.progress(f"表「{tname}」のページを調べています")
            try:
                d["index"] = TableIndex(self.db, tname, pg)
            except (JetError, struct.error, IndexError, ValueError) as e:
                d["error"] = f"表「{tname}」を読めません（{e}）。値を推測しません。{GUIDE}"
            self._meta.append(d)
        cells = sum(d["index"].records * d["index"].max_col for d in self._meta if d["index"] is not None)
        self.est_chars = cells * CHARS_PER_CELL
        if self.est_chars <= limits.CONTEXT_BUDGET:
            self.est_chars = max(self.est_chars, self._measure())

    def close(self) -> None:
        if self.src is not None:
            self.src.close()

    @property
    def sheet_meta(self) -> list[dict]:
        return self._meta

    def _measure(self) -> int:
        """小さな表でも、長いメモの字があれば窓は大きい：メモのある表は、字を数える（セルの数が見込みに入るので、すぐ終わる）。"""
        total = 0
        for d in self._meta:
            ix = d["index"]
            if ix is None or not ix.has_memo():
                continue
            for _r, row in self.records(ix, 2, ix.rows[-1]):
                if not isinstance(row, BadRow):
                    total += sum(len(v) for v in row.values() if isinstance(v, str))
        return total

    # ------------------------------------------------------------------ 行を解く
    def records(self, ix: TableIndex, a: int, b: int):
        """シートの行 a〜b（2 以上）のレコード：(行の番号, 値の辞書か BadRow)。その行のあるページだけを解く。"""
        db = self.db
        k = max(a, 2) - 2  # 0 から数えたレコードの番号
        end = min(b, ix.rows[-1]) - 2
        pi = bisect.bisect_right(ix.before, k) - 1
        while k <= end and 0 <= pi < len(ix.pages):
            pg = ix.pages[pi]
            p = db.page(pg)
            live = db.live_rows(p)
            at = ix.before[pi]
            limit = ix.before[pi + 1] if pi + 1 < len(ix.before) else ix.records
            for s in live[k - at:limit - at]:
                if k > end:
                    return
                row = db.record(ix.tdef, p, s, ix.cols)
                if row is None:  # 索引のあとで消えた印を見つけることはない（読むだけのファイル）：番号は飛ばさない
                    row = BadRow("行の印が索引と合いません")
                elif isinstance(row, BadRow):
                    row.why = f"ページ {pg} の行 {s}：{row.why}"
                yield k + 2, row
                k += 1
            pi += 1

    def rows_book(self, wanted: dict[int, list[tuple[int, int]]]) -> RecordBook:
        """頼まれた行だけのブック（シートはすべての表の分。行の番号はそのまま）。"""
        wb = RecordBook(b"", self.name, "mdb")
        for si, d in enumerate(self._meta):
            sh = wb.add_sheet(d["name"])
            ix = d["index"]
            sh.records = []
            sh.bad_rows = {}
            if ix is None or si not in wanted:
                continue
            for a, b in _merge(wanted[si]):
                if a <= 1:
                    for c, col in enumerate(ix.order, 1):
                        wb.put(sh, 1, c, col.name)
                if b < 2:
                    continue
                sh.records.append((max(a, 2), b))
                for r, row in self.records(ix, a, b):
                    if isinstance(row, BadRow):
                        sh.bad_rows[r] = row.why
                        sh.max_row = max(sh.max_row, r)
                        continue
                    for c, col in enumerate(ix.order, 1):
                        v = row.get(col.name)
                        if v is not None:
                            wb.put(sh, r, c, v)
                    sh.max_row = max(sh.max_row, r)
        return wb

    # ------------------------------------------------------------------ 概要
    def overview(self) -> dict:
        sheets = []
        for d in self._meta:
            ix = d["index"]
            if ix is None:
                sheets.append({"name": d["name"], "state": d["state"], "error": d["error"]})
                continue
            sheets.append({"name": d["name"], "state": d["state"], "rows": ix.records, "first_row": 2 if ix.records else None,
                           "last_row": ix.rows[-1] if ix.records else None, "columns": [c.name for c in ix.order]})
        return {"sheets": sheets, "ranges": "sheet と rows（行の範囲。1 行目が列の名前、2 行目からレコード）", "bytes": self.size,
                "linked_tables": len(self.linked), "queries": self.queries}

    def overview_text(self) -> str:
        out = [f"# {self.name}", "", f"- Access のデータベース（{self.size:,} バイト。読むだけ）。全体を一度には出しません。read_window に sheet（表の名前）と rows（\"2-501\"）を渡すと、"
               "その範囲の行をそのまま返します（どの表のどの行にも届きます。cursor をたどれば、すべての表のすべての行を読めます）。read_table も sheet・rows で同じ範囲を返します。",
               "- 各表の 1 行目は列の名前、2 行目からレコードです。すべての列が NULL のレコードも 1 行です。"]
        out += [f"- **読んでいない部分**：{u}" for u in self.notes]
        out += ["", "## 表", "", "| 表 | レコードの数 | 行の範囲 | 列の数 | 列 |", "|---|---|---|---|---|"]
        for d in self._meta:
            ix = d["index"]
            if ix is None:
                out.append(f"| {d['name']} | 読めません | | | {d['error']} |")
                continue
            span = f"2〜{ix.rows[-1]}" if ix.records else "（レコードなし）"
            out.append(f"| {d['name']} | {ix.records:,} | {span} | {ix.max_col} | {'・'.join(c.name for c in ix.order)} |")
            out += [f"- **読んでいない部分**：{u}" for u in ix.notes]
        return "\n".join(out) + "\n"

    # ------------------------------------------------------------------ 読む
    def _sheet_at(self, sheet) -> int:
        if sheet is None:
            return 0
        for i, d in enumerate(self._meta):
            if d["name"] == sheet or d["table"] == sheet:
                return i
        for i, d in enumerate(self._meta):
            if d["name"].lower() == str(sheet).lower():
                return i
        raise RangeError(f"シート「{sheet}」はありません（ある名前：" + "・".join(d["name"] for d in self._meta[:30]) + "）")

    def read(self, rng: dict, max_chars: int) -> Page:
        cur = rng.get("cursor")
        if rng.get("part") is not None or str(cur or "").startswith("chars:"):
            return self._built(rng, max_chars)
        skip = 0
        if cur:
            m = re.fullmatch(r"sheet:(\d+):row:(\d+)(?::char:(\d+))?(?::to:(\d+))?", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
            si, a, skip = int(m[1]), int(m[2]), int(m[3] or 0)
            stop = int(m[4]) if m[4] else None
            walk = stop is None
        else:
            if rng.get("rows") is not None and rng.get("sheet") is None:
                raise RangeError("rows は sheet（シートの名前）といっしょに渡してください")
            si = self._sheet_at(rng.get("sheet"))
            a, stop = parse_range(rng.get("rows"), "行")
            a = 1 if a is None else a
            walk = rng.get("sheet") is None
        if si >= len(self._meta):
            if not self._meta and si == 0:
                return Page("利用者の表がありません", "表なし", None, "", [])
            raise RangeError(f"表は {len(self._meta)} 個です")
        d = self._meta[si]
        ix = d["index"]
        if ix is None:
            return self._end(si, walk, f"表「{d['name']}」（読めません）", [f"## シート「{d['name']}」", "", d["error"]], [])
        last = ix.rows[-1]
        if not walk and stop is None:
            stop = last  # 表を指して頼んだ：続きも、その表の中だけ
        b = last if stop is None else min(stop, last)
        if a > last:
            text = f"（表「{d['name']}」の行は {last:,} 行目までです。{a:,} 行目はありません）"
            return self._end(si, walk, f"表「{d['name']}」{a}〜{stop or a} 行目（行なし）", [text], [])
        head = [f"## シート「{d['name']}」", ""]
        first_look = not cur or (a == 1 and not skip)
        if first_look:
            head += [f"- **読んでいない部分**：{u}" for u in (self.notes if si == 0 else []) + ix.notes]
            if len(head) > 2:
                head.append("")
        head += ["### 読み順（上から）", "", LEGEND]
        if a > 1:
            head.append("〔列の名前〕" + " ｜ ".join(cell_text(c.name) for c in ix.order))
        head.append("")
        size = sum(len(x) + 1 for x in head)
        body: list[str] = []
        done = a - 1
        cut_row = None
        for r, text in self._lines(ix, a, b):
            line = f"- {r}: {text}"
            if skip and r == a:
                line = line[skip:]
            if len(line) > max_chars and not body:  # 1 行だけで大きすぎる：字の位置で分ける
                body.append(line[:max_chars])
                cut_row = (r, skip + max_chars)
                break
            if body and size + len(line) + 1 > max_chars:
                break
            body.append(line)
            size += len(line) + 1
            done = r
        to = f":to:{stop}" if stop is not None else ""
        table = f"表「{d['name']}」"
        if cut_row:
            cov = f"{table} {cut_row[0]} 行目の {skip + 1:,}〜{cut_row[1]:,} 字"
            return Page("\n".join(head + body), cov, {"cursor": f"sheet:{si}:row:{cut_row[0]}:char:{cut_row[1]}{to}"}, "この行の残りと後ろの行", body)
        cov = f"{table} {a}〜{done} 行目（表の行 1〜{last}）"
        if done < b:
            return Page("\n".join(head + body), cov, {"cursor": f"sheet:{si}:row:{done + 1}{to}"}, f"この表の {comma(b - done)} 行" + ("と、ほかの表" if walk and si + 1 < len(self._meta) else ""), body)
        return self._end(si, walk, cov, head + body, body)

    def _lines(self, ix: TableIndex, a: int, b: int):
        if a <= 1 <= b:
            yield 1, " ｜ ".join(cell_text(c.name) for c in ix.order)
        if b >= 2:
            for r, row in self.records(ix, max(a, 2), b):
                yield r, record_text(ix.order, row)

    def _end(self, si: int, walk: bool, cov: str, lines: list[str], body: list[str]) -> Page:
        text = "\n".join(lines)
        if walk and si + 1 < len(self._meta):
            return Page(text, cov, {"cursor": f"sheet:{si + 1}:row:1"}, f"ほかの表 {len(self._meta) - si - 1} 個", body)
        return Page(text, cov, None, "", body)

    def _built(self, rng: dict, max_chars: int) -> Page:
        """窓の全体の部分（part）・字の位置（小さなデータベースの窓の読み方）。"""
        if self.large:
            raise RangeError("大きなデータベースは part ではなく、sheet と rows で読みます（open_document の概要を見てください）")
        from .builtdoc import BuiltDoc

        return BuiltDoc(self.path, self.password, None).read(rng, max_chars)

    def whole(self) -> str:  # api.read_range が part を渡してよい印
        raise NotImplementedError

    # ------------------------------------------------------------------ 欄・書く
    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        return {"fields": [], "total_fields": 0, "next": None, "note": "Access のデータベースには書かないので、入力欄はありません"}

    def fill(self, values: dict, out_path: str, force: bool, overwrite: bool, said: dict | None) -> dict:
        from ..api import ApiError
        from ..morekinds import WRITE_NOTE

        raise ApiError(WRITE_NOTE["mdb"])

    # ------------------------------------------------------------------ 段階・検査
    def grade(self):
        from ..grade import Reason, _final
        from ..morekinds import KIND_JA, WRITE_NOTE

        reasons = [Reason("C", "読んでいない部分", u) for u in self.notes]
        for d in self._meta:
            if d["index"] is None:
                reasons.append(Reason("C", "読めない表", d["error"]))
            else:
                reasons += [Reason("C", "読んでいない部分", u) for u in d["index"].notes]
        write = [Reason("D", f"{KIND_JA['mdb']}には書かない", WRITE_NOTE["mdb"])]
        g = _final(reasons, write)
        g.deferred = ["未確認：全部の行の形（audit_document で全部の行を確かめます）"]
        return g

    def audit_begin(self) -> dict:
        return {"sheet": 0, "row": 2, "rows": 0, "cells": 0, "bad": 0, "problems": [], "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        si = state["sheet"]
        if si < len(self._meta):
            d = self._meta[si]
            ix = d["index"]
            if ix is None or state["row"] > ix.rows[-1]:
                state["sheet"] = si + 1
                state["row"] = 2
                return state
            b = min(ix.rows[-1], state["row"] + 2047)
            for r, row in self.records(ix, state["row"], b):
                state["rows"] += 1
                if isinstance(row, BadRow):
                    state["bad"] += 1
                    if len(state["problems"]) < 30:
                        state["problems"].append(f"{d['name']} の {r} 行目：{row.why}")
                else:
                    state["cells"] += sum(1 for v in row.values() if v is not None)
            state["row"] = b + 1
            limits.progress(f"表「{d['name']}」を確かめています", b, ix.rows[-1])
            return state
        state["checks"] = [{"name": "読み落とし（レコード）", "ok": not state["bad"], "agent": False, "items": list(state["problems"]),
                            "detail": f"全部の表の全部の行を解きました（行 {state['rows']:,}・値のある列 {state['cells']:,}。形を確かめられない行 {state['bad']:,}）"}]
        state["still_deferred"] = []
        state["done"] = True
        return state


def jet_member(path: str, password: str | None):
    """書庫の中の Access のファイル（名前で見分ける。丸ごと展開するので、ほかの形式には読まない）。"""
    from .. import ziparc
    from ..api import _read

    sp = ziparc.split(path)
    if sp is None or not ziparc.base_name(sp[1][-1]).lower().endswith(JET_NAMES):
        return None
    blob, name, k, _lock = _read(path, password)
    if k != "mdb":
        return None
    try:
        return JetLazy(path, name, password, blob)
    except (JetError, struct.error, IndexError):
        return None  # 窓を作る道で、その理由を伝える


def _merge(spans) -> list[tuple[int, int]]:
    out: list[list[int]] = []
    for a, b in sorted(spans):
        if out and a <= out[-1][1] + 1:
            out[-1][1] = max(out[-1][1], b)
        else:
            out.append([a, b])
    return [(a, b) for a, b in out]
