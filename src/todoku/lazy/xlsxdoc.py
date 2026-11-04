"""大きな .xlsx を分けて読み書きする。

- 読むときに一度だけ、シートの XML を流し読み（expat）して、行の番号とバイトの位置の索引を作る。展開は少しずつ進み、展開の途中の状態を
  およそ 4 MB ごとに控える（zlib のコピー）ので、あとで好きな行から展開し直せる（平文を一時ファイルに置かない）
- 共有文字列：展開したバイトと、各 <si> の位置の配列だけを持ち、使うときに 1 つずつ読む（LazyStrings）
- 窓：頼まれた行だけを持つ小さなブック（slice）を作り、いままでの窓の読み手（xlsx/window.py）で読む。だから窓の字・欄の決まりは同じ。
  セルの id は元の番地のまま。入力規則・条件付き書式の範囲は、頼まれた行に切ってから読む（シート全体に広げず、頼まれた範囲は必ず広げる）
- 数式：計算結果は保存された値を読む。書いたセルに連動する数式だけを計算し直す（依存の閉包：FormulaIndex）。
  計算に要る行が多すぎて小さなブックに入らないときは、計算せずに、開くときに計算し直す印を付けて、どのセルかを伝える
- 書く：小さなブックに、いままでの fill と証明をかけ、変わった行のバイトだけを、シートの XML を流して写す途中で置き換える。ほかの部品は生のバイトのまま写す
"""

from __future__ import annotations

import bisect
import hashlib
import io
import os
import re
import struct
import xml.parsers.expat as expat
import zipfile
import zlib
from array import array
from pathlib import Path

from .. import limits
from ..formula.parser import addr, parse_cell
from .base import LazyDoc, Page, RangeError, comma, parse_range, temp_beside
from .textstream import Source

CHECK_EVERY = 1 << 22  # 展開の状態を控える間隔（展開したバイト）


def package_kind(path: str) -> str | None:
    """ZIP の部品の名前だけで .xlsx か（丸ごと読まない）。ZIP64・同じ名前が 2 つの ZIP は、丸ごと読む道（いままでの確かめ）へ。"""
    try:
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
    except (zipfile.BadZipFile, OSError, ValueError):
        return None
    names = [i.filename for i in infos]
    if len(set(names)) != len(names) or any(i.file_size >= 0xFFFFFFFF or i.compress_size >= 0xFFFFFFFF or i.header_offset >= 0xFFFFFFFF for i in infos):
        return None
    s = set(names)
    if "[Content_Types].xml" not in s or "mimetype" in s or not any(n.startswith("xl/") for n in s) or any(n.startswith(("word/", "ppt/")) for n in s):
        return None
    return "xlsx"


class Member:
    def __init__(self, src: Source, info: zipfile.ZipInfo):
        self.name = info.filename
        self.info = info
        self.src = src
        self.compressed = info.compress_size
        self.size = info.file_size
        head = src.read(info.header_offset, info.header_offset + 30)
        n, m = struct.unpack("<HH", head[26:30])
        self.data_at = info.header_offset + 30 + n + m
        if head[:4] != b"PK\x03\x04" or src.read(info.header_offset + 30, info.header_offset + 30 + n).decode("utf-8", "replace") != \
                info.orig_filename and info.flag_bits & 0x800:
            raise ValueError(f"ZIP のローカルヘッダと中央ディレクトリの名前が合いません（{self.name}）")

    def read(self) -> bytes:
        if self.info.compress_type == 0:
            return self.src.read(self.data_at, self.data_at + self.compressed)
        return limits.inflate(self.src.read(self.data_at, self.data_at + self.compressed), -15, self.name)

    def chunks(self):
        """展開したバイトを少しずつ。出てくるバイトを数える（爆弾は比率の上限で止まる）。"""
        if self.info.compress_type == 0:
            yield from self.src.chunks(self.data_at, self.data_at + self.compressed)
            return
        st = limits.StreamInflate(self.name, -15, self.compressed)
        for c in self.src.chunks(self.data_at, self.data_at + self.compressed):
            while c:
                piece = st.feed(c, 1 << 20)
                c = st.d.unconsumed_tail
                if piece:
                    yield piece
        tail = st.d.flush()
        if tail:
            yield tail


class SeekableInflate:
    """展開の途中の状態を控えて、展開したバイトの好きな範囲を読む。メモリは控えの数に比例（4 MB ごとに、およそ 40 KB）。"""

    def __init__(self, m: Member):
        self.m = m
        self.marks: list[tuple[int, int, object]] = []  # (展開の位置, 圧縮の位置, 展開器のコピー)

    def scan(self):
        """はじめから順に、展開したバイトを出しながら控えを作る。"""
        m = self.m
        if m.info.compress_type == 0:
            yield from m.chunks()
            return
        st = limits.StreamInflate(m.name, -15, m.compressed)
        out = 0
        pos = m.data_at
        end = m.data_at + m.compressed
        self.marks = [(0, pos, st.d.copy())]
        since = 0
        while pos < end:
            c = m.src.read(pos, min(pos + (1 << 18), end))
            pos += len(c)
            while c:
                piece = st.feed(c, 1 << 20)
                c = st.d.unconsumed_tail
                if piece:
                    yield piece
                    out += len(piece)
                    since += len(piece)
                if since >= CHECK_EVERY and not c:
                    self.marks.append((out, pos, st.d.copy()))
                    since = 0
                if not piece:
                    break

    def read(self, a: int, b: int) -> bytes:
        m = self.m
        if m.info.compress_type == 0:
            return m.src.read(m.data_at + a, m.data_at + b)
        i = bisect.bisect_right([x[0] for x in self.marks], a) - 1
        out_at, pos, d0 = self.marks[max(i, 0)]
        d = d0.copy()
        end = m.data_at + m.compressed
        got = bytearray()
        at = out_at
        while pos < end and at < b:
            c = m.src.read(pos, min(pos + (1 << 18), end))
            pos += len(c)
            while c and at < b:
                piece = d.decompress(c, 1 << 20)
                c = d.unconsumed_tail
                if at + len(piece) > a:
                    got += piece[max(0, a - at):b - at]
                at += len(piece)
                if not piece:
                    break
        return bytes(got)


class LazyStrings:
    """共有文字列：展開したバイト（xml）と、各 <si> の始まりの位置。字は使うときに読む（いくつかは控える）。"""

    def __init__(self, xml: bytes):
        self.xml = xml
        self.offs = array("Q")
        self.ends = array("Q")
        self._cache: dict[int, str] = {}
        self.prefix = ""
        p = expat.ParserCreate()
        depth = [0]

        def start(name, attrs):
            depth[0] += 1
            if depth[0] == 2 and name.rsplit(":", 1)[-1] == "si":
                self.offs.append(p.CurrentByteIndex)

        def end(name):
            if depth[0] == 2 and name.rsplit(":", 1)[-1] == "si":
                self.ends.append(_end_index(xml, p.CurrentByteIndex))
            depth[0] -= 1

        p.StartElementHandler = start
        p.EndElementHandler = end
        p.StartDoctypeDeclHandler = _no_doctype
        p.Parse(xml, True)
        self.head = xml[:self.offs[0]] if self.offs else b"<sst/>"
        m = re.search(rb"<(\w+:)?sst\b([^>]*)>", self.head)
        self.wrap_open = (b"<" + (m[1] or b"") + b"sst" + m[2] + b">") if m else b"<sst>"
        self.wrap_close = b"</" + ((m[1] or b"") if m else b"") + b"sst>"
        self.phonetic = _Phonetic(self)

    def __len__(self) -> int:
        return len(self.offs)

    def si_xml(self, i: int) -> bytes:
        return self.xml[self.offs[i]:self.ends[i]]

    def __getitem__(self, i: int):
        if isinstance(i, slice):
            return [self[k] for k in range(*i.indices(len(self)))]
        got = self._cache.get(i)
        if got is not None:
            return got
        from ..xlsx.workbook import _rich_text, parse_xml

        root = parse_xml(self.wrap_open + self.si_xml(i) + self.wrap_close)
        text = _rich_text(root[0]) if len(root) else ""
        if len(self._cache) > 200_000:
            self._cache.clear()
        self._cache[i] = text
        return text

    def __iter__(self):
        return (self[i] for i in range(len(self)))


class _Phonetic:
    """共有文字列の番号 → ふりがな（読みが要るときだけ、その <si> を読む）。"""

    def __init__(self, s: LazyStrings):
        self.s = s

    def __contains__(self, i) -> bool:
        return 0 <= i < len(self.s) and b"rPh" in self.s.si_xml(i)

    def __setitem__(self, k, v):
        pass

    def get(self, i, default=None):
        if i not in self:
            return default
        from ..xlsx.workbook import _phonetic_text, parse_xml

        root = parse_xml(self.s.wrap_open + self.s.si_xml(i) + self.s.wrap_close)
        got = _phonetic_text(root[0])
        return got if got is not None else default


def _end_index(xml: bytes, at: int) -> int:
    """要素の閉じタグの終わり（expat の位置は閉じタグの始まり。自分で閉じたタグなら位置はその開きの始まり）。"""
    k = xml.find(b">", at)
    return k + 1


def _no_doctype(*a):
    raise ValueError("XML に DOCTYPE の宣言があるので読みません（Office の部品には現れないもので、展開の爆弾に使われます）")


class SheetIndex:
    """1 枚のシートの索引。"""

    def __init__(self, member: Member):
        self.m = member
        self.reader = SeekableInflate(member)
        self.rows = array("I")
        self.offs = array("Q")
        self.data_start = 0  # <sheetData> の中の始め
        self.data_end = 0  # </sheetData> の始め
        self.head = b""
        self.tail = b""
        self.self_closed = False
        self.cells = 0
        self.values = 0
        self.formulas = 0
        self.max_col = 0
        self.ascending = True
        self._scan()

    def _scan(self) -> None:
        p = expat.ParserCreate()
        p.buffer_text = True
        p.StartDoctypeDeclHandler = _no_doctype
        rows = self.rows
        offs = self.offs
        state = {"in_data": False, "next_row": 1, "col": 0, "depth": 0}
        kept = bytearray()
        size = self.m.size or 1

        def start(name, attrs):
            state["depth"] += 1
            local = name.rsplit(":", 1)[-1]
            if local == "c":
                self.cells += 1
                ref = attrs.get("r")
                if ref:
                    i = 0
                    while i < len(ref) and ref[i] > "9":
                        i += 1
                    n = 0
                    for ch in ref[:i]:
                        n = n * 26 + ord(ch) - 64
                    if n > self.max_col:
                        self.max_col = n
                return
            if local in ("v", "is"):
                self.values += 1
                return
            if local == "f":
                self.formulas += 1
                return
            if local == "row":
                r = attrs.get("r")
                num = int(r) if r else state["next_row"]
                if rows and num <= rows[-1]:
                    self.ascending = False
                rows.append(num)
                offs.append(p.CurrentByteIndex)
                state["next_row"] = num + 1
                return
            if local == "sheetData" and state["depth"] == 2:
                state["sheetdata_at"] = p.CurrentByteIndex

        def end(name):
            state["depth"] -= 1
            if name.rsplit(":", 1)[-1] == "sheetData" and state["depth"] == 1:
                self.data_end = p.CurrentByteIndex

        p.StartElementHandler = start
        p.EndElementHandler = end
        total = 0
        head_cap = 1 << 26  # 頭の部分（<sheetData> の前）を覚える
        for c in self.reader.scan():
            if "sheetdata_at" not in state and total < head_cap:
                kept += c
            total += len(c)
            p.Parse(c, False)
            limits.progress(f"シート {self.m.name} の行の位置を調べています", total, size)
        p.Parse(b"", True)
        self.total = total
        at = state.get("sheetdata_at")
        if at is None:
            raise ValueError(f"{self.m.name}：sheetData がありません")
        self.head = bytes(kept[:at]) if len(kept) >= at else self.reader.read(0, at)
        sd = self.reader.read(at, at + 256)
        close = sd.find(b">")
        self.sheetdata_tag = sd[:close + 1]
        self.self_closed = sd[close - 1:close] == b"/"
        self.data_start = at + close + 1
        if self.self_closed:
            self.data_end = self.data_start
            self.tail = self.reader.read(self.data_start, total)
        else:
            self.tail = self.reader.read(self.data_end, total)
        self.prefix = re.match(rb"<(\w+:)?", self.sheetdata_tag)[1] or b""

    def locate(self, a: int, b: int) -> tuple[int, int]:
        """行の番号の範囲 → 索引の位置の範囲 [i, j)。"""
        return bisect.bisect_left(self.rows, a), bisect.bisect_right(self.rows, b)

    def row_bytes(self, i: int, j: int) -> bytes:
        if i >= j:
            return b""
        end = self.offs[j] if j < len(self.offs) else self.data_end
        return self.reader.read(self.offs[i], end)

    def span(self, i: int, j: int) -> tuple[int, int]:
        return self.offs[i], (self.offs[j] if j < len(self.offs) else self.data_end)


_SQREF = re.compile(rb'(\bsqref=")([^"]*)(")')
_MERGE = re.compile(rb'<(?:\w+:)?mergeCell\b[^>]*\bref="([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?"[^>]*/>')


def _clip_sqref(tail: bytes, a: int, b: int) -> bytes:
    """入力規則・条件付き書式の範囲を、頼まれた行に切る（窓のためだけ。ファイルには書かない）。"""
    def clip(m):
        out = []
        for ref in m[2].split():
            mm = re.fullmatch(rb"\$?([A-Za-z]+)\$?(\d+)(?::\$?([A-Za-z]+)\$?(\d+))?", ref)
            if not mm:
                out.append(ref)
                continue
            r1 = int(mm[2])
            r2 = int(mm[4]) if mm[4] else r1
            if r2 < a or r1 > b:
                continue
            c1, c2 = mm[1], mm[3] or mm[1]
            x, y = max(r1, a), min(r2, b)
            out.append(c1 + str(x).encode() + (b":" + c2 + str(y).encode() if (c1, x) != (c2, y) else b""))
        return m[1] + b" ".join(out) + m[3]

    return _SQREF.sub(clip, tail)


class FormulaIndex:
    """数式の依存の索引（書くときに初めて作る）：どのセルの値が変わると、どの数式が読むか。"""

    def __init__(self, doc: "XlsxLazy"):
        from ..formula.parser import iter_refs, parse

        self.single: dict[int, array] = {}  # (シート, 行, 列) の鍵 → 数式の番号
        self.ranges: list[tuple[int, int, int, int, int, int]] = []  # (シート, r1, c1, r2, c2, 数式の番号)
        self.cells = array("Q")  # 数式の番号 → セルの鍵
        self.opaque: list[int] = []  # 依存を決められない数式（名前・表・間接参照）
        self.refs: dict[int, list] = {}
        names = {s["name"]: i for i, s in enumerate(doc.sheet_meta)}
        pat_c = re.compile(rb"<(?:\w+:)?c\b([^>]*?)(/?)>", re.S)
        pat_f = re.compile(rb"<(?:\w+:)?f\b([^>]*?)(?:/>|>(.*?)</(?:\w+:)?f>)", re.S)
        cache: dict[str, tuple] = {}

        def analyse(text: str):
            got = cache.get(text)
            if got is not None:
                return got
            try:
                ast = parse(text)
                out = []
                opaque = _has_opaque(ast)
                for ref in iter_refs(ast):
                    out.append((ref.sheet, ref.r1, ref.c1, ref.r2, ref.c2, getattr(ref, "abs_r1", True), getattr(ref, "abs_c1", True),
                                getattr(ref, "abs_r2", True), getattr(ref, "abs_c2", True)))
                got = (out, opaque)
            except Exception:  # noqa: BLE001  読めない数式：依存は決められない
                got = ([], True)
            if len(cache) < 100_000:
                cache[text] = got
            return got

        for si, meta in enumerate(doc.sheet_meta):
            ix = meta.get("index")
            if ix is None or not ix.formulas:
                continue
            shared: dict[bytes, tuple[str, int, int]] = {}
            # 行の範囲ごとに読む（XML を丸ごと持たない）
            n = len(ix.rows)
            step = 4096
            for i in range(0, n, step):
                xml = ix.row_bytes(i, min(i + step, n))
                if b"f" not in xml:
                    continue
                for m in re.finditer(rb"<(?:\w+:)?c\b[^>]*?\br=\"([A-Z]+\d+)\"[^>]*?>(?:(?!</(?:\w+:)?c>).)*?<(?:\w+:)?f\b([^>]*?)(?:/>|>(.*?)</(?:\w+:)?f>)", xml, re.S):
                    r, c = parse_cell(m[1].decode())
                    attrs = m[2]
                    text = (m[3] or b"").decode("utf-8", "replace")
                    if text:
                        text = _unesc(text)
                    t = re.search(rb'\bt="(\w+)"', attrs)
                    sim = re.search(rb'\bsi="(\d+)"', attrs)
                    dr = dc = 0
                    if t and t[1] == b"shared" and sim:
                        if text:
                            shared[sim[1]] = (text, r, c)
                        else:
                            mast = shared.get(sim[1])
                            if mast is None:
                                self._add(si, r, c, [], True, 0, 0)
                                continue
                            text, r0, c0 = mast
                            dr, dc = r - r0, c - c0
                    if t and t[1] in (b"dataTable",):
                        self._add(si, r, c, [], True, 0, 0)
                        continue
                    refs, opaque = analyse(text)
                    self._add(si, r, c, refs, opaque, dr, dc, names, si)

    def _add(self, si, r, c, refs, opaque, dr, dc, names=None, here=None):
        fid = len(self.cells)
        self.cells.append(_key(si, r, c))
        if opaque:
            self.opaque.append(fid)
        keep = []
        for (sheet, r1, c1, r2, c2, a1, b1, a2, b2) in refs:
            s2 = names.get(sheet, -1) if sheet else here
            if s2 == -1:
                continue
            r1b = r1 + (0 if a1 else dr)
            c1b = c1 + (0 if b1 else dc)
            r2b = r2 + (0 if a2 else dr)
            c2b = c2 + (0 if b2 else dc)
            keep.append((s2, r1b, c1b, r2b, c2b))
            if r1b == r2b and c1b == c2b:
                self.single.setdefault(_key(s2, r1b, c1b), array("I")).append(fid)
            else:
                self.ranges.append((s2, r1b, c1b, r2b, c2b, fid))
        self.refs[fid] = keep

    def dependents(self, cells: set[int]) -> set[int]:
        """cells（セルの鍵）に連動する数式のセルの鍵（連動の連動まで）。"""
        seen: set[int] = set()
        todo = list(cells)
        by_sheet: dict[int, list] = {}
        for rng in self.ranges:
            by_sheet.setdefault(rng[0], []).append(rng)
        while todo:
            k = todo.pop()
            si, r, c = _unkey(k)
            hits = list(self.single.get(k, ()))
            for (_s, r1, c1, r2, c2, fid) in by_sheet.get(si, ()):
                if r1 <= r <= r2 and c1 <= c <= c2:
                    hits.append(fid)
            for fid in hits:
                ck = self.cells[fid]
                if ck not in seen:
                    seen.add(ck)
                    todo.append(ck)
        return seen

    def precedents(self, formula_cells: set[int]) -> list[tuple[int, int, int]]:
        """数式が読む範囲：(シート, 行の始め, 行の終わり)。"""
        out = []
        for fid, ck in enumerate(self.cells):
            if ck in formula_cells:
                for (s2, r1, _c1, r2, _c2) in self.refs.get(fid, ()):
                    out.append((s2, r1, r2))
        return out


def _key(si: int, r: int, c: int) -> int:
    return (si << 36) | (r << 16) | c


def _unkey(k: int) -> tuple[int, int, int]:
    return k >> 36, (k >> 16) & 0xFFFFF, k & 0xFFFF


def _unesc(s: str) -> str:
    return s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&")


def _has_opaque(ast) -> bool:
    """依存を字から決められない数式か（名前・間接参照・OFFSET など）。"""
    from ..formula.parser import NameNode, iter_funcs

    def walk(n):
        if isinstance(n, NameNode):
            return True
        kids = getattr(n, "__dict__", {}).values()
        for x in kids:
            if isinstance(x, (list, tuple)):
                for y in x:
                    if hasattr(y, "__dict__") and walk(y):
                        return True
            elif hasattr(x, "__dict__") and walk(x):
                return True
        return False

    try:
        if any(f in ("INDIRECT", "OFFSET") for f in iter_funcs(ast)):
            return True
        return walk(ast)
    except Exception:  # noqa: BLE001
        return True


class XlsxLazy(LazyDoc):
    kind = "xlsx"
    streams = True

    def __init__(self, path: str, name: str):
        self.path = path
        self.name = name
        self.src = Source(path)
        self.z = zipfile.ZipFile(path)
        self.members = {i.filename: Member(self.src, i) for i in self.z.infolist()}
        self._strings: LazyStrings | None = None
        self._meta: list[dict] | None = None
        self._template: dict[str, bytes] | None = None
        self.ratio = 0.25  # 窓の字 ÷ シートの XML のバイト（測って直す）
        self._fx: FormulaIndex | None = None
        self.sst_part = None
        self.est_chars = self._estimate()

    def close(self) -> None:
        self.z.close()
        self.src.close()

    # ------------------------------------------------------------------ 見積もり
    def _estimate(self) -> int:
        """窓の字の数の見積もり：シートの頭 1 MB で値のあるセルの密度を測り、シートの XML の大きさに掛ける。"""
        sheets = [m for n, m in self.members.items() if n.startswith("xl/worksheets/") and n.endswith(".xml") and "_rels" not in n]
        total = sum(m.size for m in sheets)
        sst = next((m for n, m in self.members.items() if n.lower().endswith("sharedstrings.xml")), None)
        sst_size = sst.size if sst else 0
        if total + sst_size <= limits.CONTEXT_BUDGET:
            return 0  # 部品の字の全部でも予算に入る（宣言の大きさ。嘘なら展開の上限で止まる）
        est = 0
        for m in sheets:
            data = bytearray()
            for c in m.chunks():
                data += c
                if len(data) >= 1 << 20:
                    break
            cells = len(re.findall(rb"<(?:\w+:)?(?:v|is)\b", data))
            est += int(m.size * cells / max(len(data), 1) * 16)
        return est + sst_size // 4

    # ------------------------------------------------------------------ 部品
    @property
    def strings(self) -> LazyStrings | None:
        self._load_meta()
        return self._strings

    def _load_meta(self) -> list[dict]:
        if self._meta is not None:
            return self._meta
        from ..xlsx.workbook import M, R, parse_xml

        root_rels = self._rels("")
        wb_path = next((t for typ, t in root_rels.values() if typ == "officeDocument"), "xl/workbook.xml")
        self.wb_path = wb_path
        rels = self._rels(wb_path)
        root = parse_xml(self.members[wb_path].read())
        for typ, target in rels.values():
            if typ == "sharedStrings" and target in self.members:
                self.sst_part = target
                self._strings = LazyStrings(self.members[target].read())
        meta = []
        el = root.find(M + "sheets")
        for s in el if el is not None else []:
            typ, path = rels.get(s.get(R + "id"), ("", ""))
            d = {"name": s.get("name"), "path": path, "state": s.get("state", "visible"), "index": None}
            if typ == "worksheet" and path in self.members:
                d["index"] = SheetIndex(self.members[path])
            meta.append(d)
        self._meta = meta
        return meta

    @property
    def sheet_meta(self) -> list[dict]:
        return self._load_meta()

    def _rels(self, part: str) -> dict:
        import posixpath

        from ..xlsx.workbook import parse_xml

        d, base = posixpath.split(part)
        rp = posixpath.join(d, "_rels", base + ".rels")
        if rp not in self.members:
            return {}
        out = {}
        for rel in parse_xml(self.members[rp].read()):
            t = rel.get("Target", "")
            if rel.get("TargetMode") == "External":
                full = t
            elif t.startswith("/"):
                full = t[1:]
            else:
                full = posixpath.normpath(posixpath.join(d, t))
            out[rel.get("Id")] = (rel.get("Type", "").rsplit("/", 1)[-1], full)
        return out

    def _base_parts(self) -> dict[str, bytes]:
        """小さなブックの、シートと共有文字列のほかの部品（展開して控える）。"""
        if self._template is None:
            meta = self.sheet_meta
            skip = {d["path"] for d in meta if d["index"] is not None}
            if self.sst_part:
                skip.add(self.sst_part)
            self._template = {n: m.read() for n, m in self.members.items() if n not in skip}
        return self._template

    def slice_blob(self, wanted: dict[int, list[tuple[int, int]]], clip: dict[int, tuple[int, int]] | None = None) -> bytes:
        """小さなブック：wanted（シート → 行の番号の範囲の並び）の行だけ。ほかのシートは行のない形（列の幅・結合・図形などは残す）。"""
        buf = io.BytesIO()
        parts = self._base_parts()
        meta = self.sheet_meta
        by_path = {d["path"]: (si, d) for si, d in enumerate(meta) if d["index"] is not None}
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
            for n in self.members:
                if n in parts:
                    z.writestr(n, parts[n])
                elif n == self.sst_part:
                    z.writestr(n, b'<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>')
                elif n in by_path:
                    si, d = by_path[n]
                    ix = d["index"]
                    xml = bytearray(ix.head)
                    open_tag = ix.sheetdata_tag
                    if ix.self_closed:
                        open_tag = open_tag[:-2].rstrip() + b">"
                    xml += open_tag
                    for a, b in _merge_ranges(wanted.get(si, [])):
                        i, j = ix.locate(a, b)
                        xml += ix.row_bytes(i, j)
                    close = b"</" + ix.prefix + b"sheetData>"
                    if ix.self_closed:
                        xml += close
                    tail = ix.tail
                    if clip and si in clip:
                        tail = _clip_sqref(tail, *clip[si])
                    elif clip is not None:
                        tail = _clip_sqref(tail, 1, 0)  # 読まないシート：広げない
                    xml += tail
                    z.writestr(n, bytes(xml))
        return buf.getvalue()

    def book(self, blob: bytes):
        from ..xlsx.workbook import Workbook, sst_scope

        with sst_scope(self.strings):
            return Workbook(blob, self.name)

    # ------------------------------------------------------------------ 概要
    def overview(self) -> dict:
        sheets = []
        for d in self.sheet_meta:
            ix = d["index"]
            one = {"name": d["name"], "state": d["state"]}
            if ix is not None:
                one.update({"rows": len(ix.rows), "first_row": ix.rows[0] if ix.rows else None, "last_row": ix.rows[-1] if ix.rows else None,
                            "cells": ix.cells, "cells_with_value": ix.values, "formulas": ix.formulas, "last_column": addr(1, ix.max_col).rstrip("0123456789") if ix.max_col else None,
                            "merges": len(_MERGE.findall(ix.tail)), "validations": ix.tail.count(b"sqref=")})
            sheets.append(one)
        return {"sheets": sheets, "shared_strings": len(self.strings) if self.strings else 0, "ranges": "sheet と rows（行の範囲）",
                "bytes": self.src.size}

    def overview_text(self) -> str:
        out = [f"# {self.name}", "", f"- Excel のブック（{self.src.size:,} バイト）。全体を一度には出しません。read_window に sheet と rows（\"1-500\"）を渡して読みます"
               "（cursor をたどれば、すべてのシートのすべての行を読めます）。list_fields も同じ範囲で、その範囲の欄と空のマスを出します。",
               "- セルの id は シート!B3（元の番地のまま）。どのセルにも番地で書けます。書き込みは、変わった行のバイトだけを置き換え、ほかのバイトはそのままファイルを写します。",
               "- 入力規則・条件付き書式は、読んだ行の範囲で広げます（シート全体には広げません）。計算結果は保存された値です。書いたセルに連動する数式は、連動する分だけ計算し直します。",
               "", "## シート", "", "| シート | 状態 | 行（範囲） | 値のあるセル | 数式 | 結合 | 入力規則 |", "|---|---|---|---|---|---|---|"]
        for s in self.overview()["sheets"]:
            if "rows" in s:
                out.append(f"| {s['name']} | {s['state']} | {s['rows']:,} 行（{s['first_row']}〜{s['last_row']}） | {s['cells_with_value']:,} | {s['formulas']:,} | {s['merges']:,} | {s['validations']:,} |")
            else:
                out.append(f"| {s['name']} | {s['state']} | セルのないシート | | | | |")
        return "\n".join(out) + "\n"

    # ------------------------------------------------------------------ 読む
    def _sheet_at(self, sheet) -> int:
        meta = self.sheet_meta
        if sheet is None:
            return 0
        for i, d in enumerate(meta):
            if d["name"] == sheet:
                return i
        for i, d in enumerate(meta):
            if d["name"].lower() == str(sheet).lower():
                return i
        raise RangeError(f"「{sheet}」というシートはありません（シート：" + "、".join(d["name"] for d in meta[:30]) + "）")

    def _sheet_doc(self, si: int, a: int, b: int) -> dict:
        """シート si の行 a〜b の窓（xlsx/window.py の 1 枚のシートの形）。"""
        from ..xlsx.window import _MEMO, _Options
        from ..xlsx.window import _sheet_doc as xw_sheet

        ix = self.sheet_meta[si]["index"]
        first = ix.rows[0] if ix is not None and ix.rows else None
        wanted = [(a, b)]
        if first is not None and first < a:
            wanted.insert(0, (first, first))  # 見出しの手がかり：はじめの行を添える（出す行には数えない）
        blob = self.slice_blob({si: wanted}, {si: (a, b)})
        wb = self.book(blob)
        sh = wb.sheets[si]
        sh.__dict__[_MEMO] = {}
        try:
            with _scan_budget():
                return xw_sheet(wb, sh, _Options(wb))
        finally:
            sh.__dict__.pop(_MEMO, None)

    def read(self, rng: dict, max_chars: int) -> Page:
        from ..xlsx.window import _md_sheet

        meta = self.sheet_meta
        cur = rng.get("cursor")
        skip = 0
        stop_at = None  # 頼まれた行の終わり
        walk = True  # シートの指定がなければ、次のシートへ続ける
        if cur:
            m = re.fullmatch(r"sheet:(\d+):row:(\d+)(?::char:(\d+))?(?::to:(\d+))?", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
            si, a, skip = int(m[1]), int(m[2]), int(m[3] or 0)
            stop_at = int(m[4]) if m[4] else None
            walk = stop_at is None
        else:
            si = self._sheet_at(rng.get("sheet"))
            a, stop_at = parse_range(rng.get("rows"), "行")
            a = 1 if a is None else a
            walk = rng.get("sheet") is None and rng.get("rows") is None
        if si >= len(meta):
            raise RangeError(f"シートは {len(meta)} 枚です")
        d = meta[si]
        ix = d["index"]
        if ix is None or not ix.rows:
            return self._sheet_end(si, walk, f"シート「{d['name']}」（セルのないシート）", self._no_rows_lines(si))
        last_row = ix.rows[-1]
        b = last_row if stop_at is None else min(stop_at, last_row)
        i, j = ix.locate(a, b)
        if a > last_row and not rng.get("rows"):
            return self._sheet_end(si, walk, f"シート「{d['name']}」", [])
        if i >= j:
            page_lines = [f"（シート「{d['name']}」の {a}〜{b} 行目には、ファイルに行がありません。番地を指定すれば、書くことはできます）"]
            return Page("\n".join(page_lines), f"シート「{d['name']}」{a}〜{b} 行目（行なし）", None, "", [])
        # 窓の字の見込みから、読む行の数を決める（測った比率。字が多すぎれば、下で行を切り詰める）
        want_bytes = int(max_chars / max(self.ratio, 0.01))
        k = i + 1
        while k < j and (ix.offs[k] - ix.offs[i]) < want_bytes:
            k = min(j, k + max(1, (k - i)))
        k = min(k, j)
        ra = ix.rows[i]
        rb = ix.rows[k - 1]
        sh = self._sheet_doc(si, ra, rb)
        lines, last_done = _trim_sheet(sh, ra, rb, si == 0 and ra == ix.rows[0], ra == ix.rows[0], max_chars, _md_sheet)
        xml_bytes = max(ix.offs[k] if k < len(ix.offs) else ix.data_end, ix.offs[i] + 1) - ix.offs[i]
        chars = sum(len(x) + 1 for x in lines)
        if last_done == rb and chars > 0:
            self.ratio = max(0.001, min(4.0, chars / xml_bytes))
        text = "\n".join(lines)
        # 1 行だけでも大きすぎるとき：字の位置で分ける
        if len(text) - skip > max_chars and last_done == ra:
            piece = text[skip:skip + max_chars]
            cov = f"シート「{d['name']}」{ra} 行目（{skip + 1:,}〜{skip + max_chars:,} 字／全 {len(text):,} 字）"
            to = f":to:{stop_at}" if stop_at else ""
            return Page(piece, cov, {"cursor": f"sheet:{si}:row:{ra}:char:{skip + max_chars}{to}"}, "この行の残りと後ろの行", [piece])
        if skip:
            text = text[skip:]
            lines = text.split("\n")
        nxt_i = bisect.bisect_right(ix.rows, last_done)
        body = lines
        cov = f"シート「{d['name']}」{ra}〜{last_done} 行目（シートの行 {ix.rows[0]}〜{last_row}）"
        if nxt_i < j:
            to = f":to:{stop_at}" if stop_at else ""
            left = j - nxt_i
            return Page(text, cov, {"cursor": f"sheet:{si}:row:{ix.rows[nxt_i]}{to}"}, f"このシートの {comma(left)} 行" + (f"と、ほかのシート" if walk and si + 1 < len(meta) else ""), body)
        return self._sheet_end(si, walk, cov, body, text)

    def _no_rows_lines(self, si: int) -> list[str]:
        from ..xlsx.window import _md_sheet

        blob = self.slice_blob({})
        wb = self.book(blob)
        from ..xlsx.window import build

        doc = build(wb)
        return _md_sheet(doc["sheets"][si])

    def _sheet_end(self, si: int, walk: bool, cov: str, body: list[str], text: str | None = None) -> Page:
        meta = self.sheet_meta
        text = "\n".join(body) if text is None else text
        if walk and si + 1 < len(meta):
            return Page(text, cov, {"cursor": f"sheet:{si + 1}:row:1"}, f"ほかのシート {len(meta) - si - 1} 枚", body)
        return Page(text, cov, None, "", body)

    # ------------------------------------------------------------------ 欄
    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        from ..api import _field

        meta = self.sheet_meta
        cur = rng.get("cursor")
        if cur:
            m = re.fullmatch(r"fields:(\d+):(\d+):(\d+)", str(cur))
            if not m:
                raise RangeError(f"cursor「{cur}」はこの書類の続きの印ではありません")
            si, a, stop = int(m[1]), int(m[2]), int(m[3])
        else:
            si = self._sheet_at(rng.get("sheet"))
            a, stop = parse_range(rng.get("rows"), "行")
            a = a or 1
            stop = stop or 0
        d = meta[si]
        ix = d["index"]
        if ix is None or not ix.rows:
            return {"fields": [], "sheet": d["name"], "next": None}
        b = min(stop or ix.rows[-1], ix.rows[-1])
        i, j = ix.locate(a, b)
        # 1 回に見る行の数：limit に比例（幅の分だけ多めに見ないように、行の数で切る）
        k = min(j, i + max(1, limit // max(ix.max_col, 1)))
        if i >= j:
            rows_a, rows_b = a, b
        else:
            rows_a, rows_b = ix.rows[i], (ix.rows[k - 1] if k < j else b)
        sh = self._sheet_doc(si, rows_a, rows_b)
        got = [f for f in sh.get("fields", []) if _in_rows(f, rows_a, rows_b)]
        out = [_field(f, sheet=d["name"]) for f in got]
        if only_empty:
            out = [f for f in out if f.get("value") in (None, "", False)]
        # 空のマス（行の範囲で、ファイルにセルの値のない所）：どれも番地で書ける
        blanks = self._blanks(si, i, k if i < j else i, ix)
        res = {"fields": out, "sheet": d["name"], "looked_rows": f"{rows_a}〜{rows_b}", "blank_cells_in_range": blanks[:limit],
               "blank_cell_count": len(blanks), "last_column": addr(1, ix.max_col).rstrip("0123456789") if ix.max_col else None, "next": None,
               "note": "blank_cell_count は、見た行の中で、使っている列（シートの最後の列まで）に値のないセルの数。番地（シート!B3）で書けます。"}
        if k < j:
            res["next"] = {"cursor": f"fields:{si}:{ix.rows[k]}:{stop}"}
            res["remaining"] = f"{comma(j - k)} 行"
        return res

    def _blanks(self, si: int, i: int, k: int, ix: SheetIndex) -> list[str]:
        name = self.sheet_meta[si]["name"]
        out = []
        if i >= k:
            return out
        xml = ix.row_bytes(i, k)
        for m in re.finditer(rb"<(?:\w+:)?row\b[^>]*?(?:/>|>(.*?)</(?:\w+:)?row>)", xml, re.S):
            r = re.search(rb'\br="(\d+)"', m[0][:m[0].find(b">") + 1])
            if not r:
                continue
            have = set()
            for cm in re.finditer(rb"<(?:\w+:)?c\b[^>]*?\br=\"([A-Z]+)\d+\"[^>]*?(/?)>(.*?)(?=<(?:\w+:)?c\b|$)", m[1] or b"", re.S):
                if cm[2] != b"/" and re.search(rb"<(?:\w+:)?(?:v|is|f)\b", cm[3]):
                    have.add(cm[1].decode())
            rn = r[1].decode()
            for c in range(1, ix.max_col + 1):
                a = addr(1, c)[:-1]
                if a not in have:
                    out.append(f"{name}!{a}{rn}")
        return out

    # ------------------------------------------------------------------ 書く
    def _touched(self, values: dict) -> tuple[dict[int, set[int]], list[str], set[int]]:
        """書く所の行：シート → 行の番号。書く所のセルの鍵。"""
        rows: dict[int, set[int]] = {}
        cells: set[int] = set()
        problems = []
        names = {d["name"]: i for i, d in enumerate(self.sheet_meta)}
        lower = {k.lower(): v for k, v in names.items()}
        for key, value in values.items():
            k = str(key)
            if k.startswith("@"):
                spec = value if isinstance(value, dict) else {}
                near = spec.get("near") or spec.get("cell") or ""
                m = re.fullmatch(r"(.+)!\$?([A-Za-z]+)\$?(\d+)", str(near))
                if not m:
                    problems.append(f"{k}: 大きなブックでは、字を目印にした書き込みに、近くのセルの番地を near（シート!B3）で添えてください"
                                    "（read_window に sheet と rows を渡すと、番地が分かります）。番地で直接書くこともできます")
                    continue
                si = names.get(m[1], lower.get(m[1].lower()))
                if si is not None:
                    rows.setdefault(si, set()).add(int(m[3]))
                continue
            m = re.fullmatch(r"(.+)!\$?([A-Za-z]+)\$?(\d+)(?:@\d+)?", k)
            if not m:
                continue  # 部品の名前など（シートの行は要らない）
            si = names.get(m[1], lower.get(m[1].lower()))
            if si is None:
                problems.append(f"{k}: 「{m[1]}」というシートはありません")
                continue
            r = int(m[3])
            from ..formula.parser import col_to_num

            c = col_to_num(m[2].upper())
            rows.setdefault(si, set()).add(r)
            cells.add(_key(si, r, c))
            # 結合：結合の左上のセルに書く
            ix = self.sheet_meta[si]["index"]
            for mm in _MERGE.finditer(ix.tail if ix is not None else b""):
                r1 = int(mm[2])
                r2 = int(mm[4]) if mm[4] else r1
                c1 = col_to_num(mm[1].decode())
                c2 = col_to_num(mm[3].decode()) if mm[3] else c1
                if r1 <= r <= r2 and c1 <= c <= c2:
                    rows[si].add(r1)
                    cells.add(_key(si, r1, c1))
        return rows, problems, cells

    def _closure(self, rows: dict[int, set[int]], cells: set[int], warnings: list[str]):
        """書いたセルに連動する数式と、その計算に要る行を、小さなブックに入れる行に足す。入らないときは計算しない（開くときに計算する印）。"""
        if not any(d["index"] is not None and d["index"].formulas for d in self.sheet_meta):
            return {si: sorted(v) for si, v in rows.items()}, []
        if self._fx is None:
            self._fx = FormulaIndex(self)
        fx = self._fx
        dep = fx.dependents(cells)
        for k in dep:
            si, r, _c = _unkey(k)
            rows.setdefault(si, set()).add(r)
        wanted: dict[int, list[tuple[int, int]]] = {si: [(r, r) for r in sorted(v)] for si, v in rows.items()}
        far: list[str] = []
        room = limits.MEMORY_LIMIT // 16  # 小さなブックの XML の大きさの上限（セルの表はこの何倍かになる）
        used = 0
        # 連動する数式が読む行
        reads = fx.precedents(dep)
        for s2, r1, r2 in sorted(reads, key=lambda t: t[2] - t[1]):
            ix = self.sheet_meta[s2]["index"]
            if ix is None or not ix.rows:
                continue
            r2 = min(r2, ix.rows[-1])
            i, j = ix.locate(r1, r2)
            if i >= j:
                continue
            a, b = ix.span(i, j)
            if used + (b - a) > room:
                far.append(f"{self.sheet_meta[s2]['name']} の {r1}〜{r2} 行目")
                continue
            used += b - a
            wanted.setdefault(s2, []).append((r1, r2))
        # 開いて計算する：名前・間接参照のように依存を字から決められない数式は、計算し直さない（印を付けて伝える）
        for s, spans in wanted.items():
            wanted[s] = _merge_ranges(spans)
        if far:
            warnings.append("書いた値に連動する数式の計算に要る行が多いので、次の範囲を読まずに、Excel で開いたときに計算し直す印を付けました："
                            + "、".join(far[:10]) + "（その数式のいまの値は、書いた値を反映していません）")
        opaque = [fx.cells[f] for f in fx.opaque]
        return wanted, ([] if not opaque else opaque)

    def fill(self, values: dict, out_path: str, force: bool, overwrite: bool, said: dict | None) -> dict:
        from ..xlsx.edit import EditError, _group_span_ids
        from ..xlsx.edit import fill as xfill
        from ..xlsx.proof import prove
        from ..xlsx.readback import check as readback_check
        from ..xlsx.workbook import sst_scope
        from ..xlsx.window import parse_span_id
        from ..zipraw import RawZip

        rows, problems, cells = self._touched(values)
        if problems:
            raise EditError("\n".join(problems))
        warnings: list[str] = []
        wanted, opaque = self._closure(rows, cells, warnings)
        if isinstance(wanted, dict):
            spans = {si: [(x, x) if isinstance(x, int) else x for x in v] for si, v in wanted.items()}
        else:
            spans = {}
        full_calc_needed = bool(warnings) or bool(opaque)
        mini = self.slice_blob(spans)
        wb = self.book(mini)
        cell_values = _group_span_ids(values)
        readback_values = {k: v for k, v in values.items() if not (isinstance(k, str) and "!" in k and not k.startswith("@") and parse_span_id(k))}
        with sst_scope(self.strings):
            res = xfill(wb, cell_values, force=force, run_macros=False)
        if wb.has_macros:
            warnings.append("このブックにはマクロがあります。大きなブックの範囲の書き込みでは、セルの値を変えたときに動くマクロは動かしていません（本人が Excel で開いたときに動きます）")
        with sst_scope(self.strings):
            pr = prove(mini, res.blob, res.allowed_inputs, {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"},
                       {(p.sheet, p.address) for p in res.patches if p.kind == "macro"}, rich=set(res.rich_cells))
        za, zb = RawZip(mini), RawZip(res.blob)
        changed_sheets: dict[str, list[tuple[int, int, bytes]]] = {}
        other_parts: dict[str, bytes] = {}
        unexpected = list(pr.unexpected)
        by_path = {d["path"]: (si, d) for si, d in enumerate(self.sheet_meta) if d["index"] is not None}
        for n in zb.names():
            if n not in set(za.names()):
                other_parts[n] = zb.read(n)
                continue
            old = za.read(n)
            new = zb.read(n)
            if old == new:
                continue
            if n in by_path:
                changed_sheets[n] = self._row_edits(by_path[n][1]["index"], old, new, unexpected)
            else:
                other_parts[n] = new
        # 連動を確かめられない数式：開いたときに計算し直す印
        if full_calc_needed and self.wb_path not in other_parts and any(p.kind in ("input",) for p in res.patches):
            xml = other_parts.get(self.wb_path) or self.members[self.wb_path].read()
            marked = _with_full_calc(xml)
            if marked is not None:
                other_parts[self.wb_path] = marked
        dst = Path(out_path).expanduser().resolve()
        tmp = temp_beside(dst)
        try:
            report = stream_package(self, changed_sheets, other_parts, tmp)
            unexpected += report["problems"]
            back, shown = self._read_back(tmp, values, readback_values, res)
            for u in res.unverified:
                warnings.append(f"{u}")
            result = {"kind": "xlsx", "out_path": str(dst), "original_sha256": report["old_sha256"], "new_sha256": report["new_sha256"],
                      "proof_ok": pr.ok and not report["problems"], "unexpected": unexpected[:20], "structure_errors": pr.lint_errors[:20],
                      "warnings": list(res.warnings) + warnings, "streamed": True,
                      "proof": (f"{pr.headline()}。ファイル全体では、変わった行のほかは同じ：シートの XML の変えていない {report['ranges']} の範囲・{comma(report['same_bytes'])} バイトと、"
                                f"ほかの部品 {report['raw_parts']} つを、書いた後のファイルを読み直して比べました"),
                      "read_back_by": "offset", "read_back_ok": not back, "read_back_differences": back[:30],
                      "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"},
                      "rows_rewritten": sum(len(v) for v in changed_sheets.values())}
            if res.unverified:
                result["stale_formulas"] = list(res.unverified)
            if res.written_formatting:
                result["written_formatting"] = list(res.written_formatting)
            if res.fit:
                result["fit"] = list(res.fit)
                result["warnings"] += [x["message"] for x in res.fit]
            if res.anchored:
                result["anchored"] = dict(res.anchored)
            if opaque or warnings:
                result["formulas_recalculated_on_open"] = True
            from .textdocs import _finish_fill

            labels = {}
            _finish_fill(result, {k: shown.get(k, v) for k, v in readback_values.items()}, labels, said, tmp, dst, self.path, overwrite,
                         "list_fields に sheet と rows を渡すと、その範囲の欄と空のマスが分かります")
            return result
        finally:
            tmp.unlink(missing_ok=True)

    def _row_edits(self, ix: SheetIndex, old: bytes, new: bytes, unexpected: list[str]) -> list[tuple[int, int, bytes]]:
        """小さなブックのシート XML の前後を比べて、変わった行を、元のシート XML の範囲の置き換えにする。"""
        pa = _rows_in(old)
        pb = _rows_in(new)
        # 行のほかの部分（頭・終わり）は変わらないはず
        if old[:pa[0][1]] != new[:pb[0][1]] if pa and pb else False:
            unexpected.append("シートの頭の部分が変わっています")
        out: list[tuple[int, int, bytes]] = []
        olds = {r: old[a:b] for r, a, b in pa}
        for r, a, b in pb:
            data = new[a:b]
            if r in olds:
                if olds[r] != data:
                    i, j = ix.locate(r, r)
                    x, y = ix.span(i, j)
                    out.append((x, y, data))
            else:
                # 新しい行：行の番号の順の位置に差し込む
                i = bisect.bisect_left(ix.rows, r)
                at = ix.offs[i] if i < len(ix.offs) else ix.data_end
                out.append((at, at, data))
        if ix.self_closed and out:
            unexpected.append("元のシートに行がないので、行を差し込めませんでした")
        return out

    def _read_back(self, tmp: Path, values: dict, readback_values: dict, res) -> tuple[list[dict], dict]:
        """書いた後のファイルを開き直し、書いた行だけの小さなブックで、書いた値を読む（索引も、書いたファイルから作り直す）。"""
        from ..xlsx.readback import check as readback_check
        from ..xlsx.workbook import sst_scope

        after = XlsxLazy(str(tmp), self.name)
        try:
            rows, _p, _c = after._touched(values)
            mini = after.slice_blob({si: sorted((r, r) for r in v) for si, v in rows.items()})
            with sst_scope(after.strings):
                wb = after.book(mini)
                got = readback_check(wb, {k: v for k, v in readback_values.items() if not str(k).startswith("@")})
            back = []
            shown = {}
            for k, (ok, text) in got.items():
                shown[k] = text
                if not ok:
                    back.append({"id": k, "written": readback_values.get(k), "read_back": text})
            return back, shown
        finally:
            after.close()

    # ------------------------------------------------------------------ 段階
    def grade(self):
        from ..grade import Reason, _final

        meta = self.sheet_meta
        reasons: list = []
        write = [Reason("B", "欄を範囲ごとに見つける", "大きなブックなので、欄は頼まれた行の範囲ごとに見つけます。どのセルにも番地で書けます")]
        total = sum(d["index"].cells for d in meta if d["index"] is not None)
        forms = sum(d["index"].formulas for d in meta if d["index"] is not None)
        reasons.append(Reason("A", "大きなブック", f"シート {len(meta)} 枚・セル {total:,}・数式 {forms:,}。行の並びと XML の作りは全部調べました"))
        if "xl/vbaProject.bin" in self.members:
            reasons.append(Reason("B", "マクロ", "マクロがあります。窓の範囲で読めるかは、grade を全体で確かめる検査（audit_document）で確かめます"))
        g = _final(reasons, write)
        g.deferred = ["未確認：全部のセルの値を別の読み方で照らすこと・表示形式（audit_document で全部の行を確かめます）"]
        if forms:
            g.deferred.append("未確認：数式の計算結果と保存された値の照合（全部の数式）")
        return g

    # ------------------------------------------------------------------ 検査
    def audit_begin(self) -> dict:
        return {"sheet": 0, "next_i": 0, "cells": 0, "problems": [], "done": False, "checks": [], "formulas_ok": 0, "formulas_other": 0,
                "formulas_cross": 0, "unshown": 0}

    def audit_step(self, state: dict) -> dict:
        """1 区切り（行のかたまり）：窓の読み手と別の読み方（XML を標準の ElementTree で）で、値のあるセルが同じかを確かめる。"""
        meta = self.sheet_meta
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
            wb = self.book(self.slice_blob({si: [(ra, rb)]}))
            sh = wb.sheets[si]
            state["cells"] += self._audit_rows(ix, i, k, sh, state["problems"], d["name"])
            self._audit_formulas(wb, sh, ra, rb, state, d["name"])
            from ..xlsx.window import shown_text

            for cell in sh.cells.values():
                if cell.value is not None and shown_text(wb, cell)[1] is not None:
                    state["unshown"] += 1
            state["next_i"] = k
            limits.progress(f"シート「{d['name']}」を確かめています", k, len(ix.rows))
            return state
        total_rows = sum(len(d["index"].rows) for d in meta if d["index"] is not None)
        state["checks"] = [{"name": "読み落とし（セルの値）", "ok": not state["problems"], "agent": False, "items": list(state["problems"]),
                            "detail": f"全部のシートの {total_rows:,} 行・値のあるセル {state['cells']:,} か所を、窓の読み手と別の読み方で読み、同じ値であることを確かめました"}]
        cross = state["formulas_cross"]
        state["checks"].append({"name": "数式の計算結果", "ok": not state.get("formula_problems"), "agent": False, "items": state.get("formula_problems", []),
                                "detail": f"同じ区切りの行だけを読む数式 {state['formulas_ok']:,} 個を計算し直し、保存された値と同じことを確かめました"
                                          + (f"。計算し直せない・保存された値のない数式 {state['formulas_other']:,} 個" if state["formulas_other"] else "")
                                          + (f"。ほかの区切りの行を読む数式 {cross:,} 個は、保存された値のままで、照らしていません（書いて連動するときは、依存の閉包で計算し直します）" if cross else "")})
        if state["unshown"]:
            state["checks"].append({"name": "表示形式", "ok": True, "agent": True, "items": [],
                                    "detail": f"表示形式を再現できないセル {state['unshown']:,} か所：窓には元の値を出しています"})
        state["still_deferred"] = [f"未確認：ほかの区切りの行を読む数式 {cross:,} 個の計算結果と保存された値の照合"] if cross else []
        state["checks"].append(self._write_trial())
        state["done"] = True
        return state

    def _audit_formulas(self, wb, sh, ra: int, rb: int, state: dict, name: str) -> None:
        """同じ区切りの行だけを読む数式を計算し直して、保存された値と照らす。ほかの行を読む数式は、照らさなかったと数える。"""
        from ..formula.parser import iter_refs
        from ..xlsx.calc import Calc, _run_deep, values_equal

        cells = [c for c in sh.cells.values() if c.has_formula]
        if not cells:
            return
        calc = Calc(wb)
        bad = state.setdefault("formula_problems", [])

        def inside(c) -> bool:
            if c.ast is None:
                return False
            dr = getattr(c, "dr", 0) or 0
            for ref in iter_refs(c.ast):
                if ref.sheet not in (None, sh.name) or ref.r1 + dr < ra or ref.r2 + dr > rb:
                    return False
            return True

        def run():
            for c in cells:
                if not inside(c):
                    state["formulas_cross"] += 1
                    continue
                if c.value is None:
                    state["formulas_other"] += 1
                    continue
                try:
                    got = calc.cell_value(sh.name, c.row, c.col)
                except Exception:  # noqa: BLE001  計算できない数式は、数えて伝える
                    state["formulas_other"] += 1
                    continue
                if values_equal(c.value, got):
                    state["formulas_ok"] += 1
                elif len(bad) < 30:
                    bad.append(f"{name}!{addr(c.row, c.col)}：保存された値 {c.value!r} と、計算し直した値 {got!r} が違います")

        _run_deep(run)

    def _audit_rows(self, ix: SheetIndex, i: int, k: int, sh, problems: list[str], name: str) -> int:
        import xml.etree.ElementTree as ET

        wrap = b'<x xmlns:r="urn:r" xmlns:mc="urn:mc" xmlns:x14ac="urn:x14ac" xmlns:xr="urn:xr" xmlns="' + b'urn:m' + b'">'
        n = 0
        try:
            root = ET.fromstring(wrap + ix.row_bytes(i, k) + b"</x>")
        except ET.ParseError:
            return 0
        strings = self.strings
        for row in root:
            for c in row:
                if c.tag.rsplit("}", 1)[-1] != "c":
                    continue
                ref = c.get("r")
                v = next((x for x in c if x.tag.rsplit("}", 1)[-1] == "v"), None)
                if ref is None or v is None or v.text is None:
                    continue
                n += 1
                got = sh.cells.get(parse_cell(ref))
                t = c.get("t")
                want = None
                if t == "s":
                    want = strings[int(v.text)] if strings is not None and int(v.text) < len(strings) else ""
                elif t in (None, "n"):
                    try:
                        want = float(v.text)
                    except ValueError:
                        want = v.text
                else:
                    continue
                if (got is None or got.value != want) and len(problems) < 30:
                    problems.append(f"{name}!{ref}：窓の読み手（{got.value if got else None!r}）と別の読み方（{want!r}）で値が違います")
        return n

    def _write_trial(self) -> dict:
        import tempfile

        meta = self.sheet_meta
        for si, d in enumerate(meta):
            ix = d["index"]
            if ix is None or not ix.rows or d["state"] != "visible":
                continue
            col = addr(1, ix.max_col + 1).rstrip("0123456789") if ix.max_col < 16384 else "XFD"
            key = f"{d['name']}!{col}{ix.rows[len(ix.rows) // 2]}"
            with tempfile.TemporaryDirectory(prefix="todoku-audit-") as td:
                try:
                    res = self.fill({key: "テスト"}, str(Path(td) / "t.xlsx"), True, True, None)
                except ValueError as e:
                    return {"name": "書き込みの試し", "ok": False, "agent": False, "detail": f"{key} に書けませんでした：{e}", "items": []}
            return {"name": "書き込みの試し", "ok": bool(res.get("proof_ok") and res.get("read_back_ok")), "agent": False,
                    "detail": f"{key}（使っていない列）に試しの値を書き、{res.get('proof')}", "items": res.get("unexpected") or []}
        return {"name": "書き込みの試し", "ok": True, "agent": False, "detail": "書けるシートがありませんでした", "items": []}


def _with_full_calc(xml: bytes) -> bytes | None:
    m = re.search(rb"<(\w+:)?calcPr\b[^>]*?/?>", xml)
    if m:
        if b"fullCalcOnLoad" in m[0]:
            return None
        tag = m[0]
        new = tag[:-2].rstrip() + b' fullCalcOnLoad="1"/>' if tag.endswith(b"/>") else tag[:-1] + b' fullCalcOnLoad="1">'
        return xml[:m.start()] + new + xml[m.end():]
    end = re.search(rb"</(\w+:)?workbook>", xml)
    if not end:
        return None
    p = end[1] or b""
    return xml[:end.start()] + b"<" + p + b'calcPr fullCalcOnLoad="1"/>' + xml[end.start():]


_ROW = re.compile(rb"<(?:\w+:)?row\b[^>]*?(?:/>|>.*?</(?:\w+:)?row>)", re.S)


def _rows_in(xml: bytes) -> list[tuple[int, int, int]]:
    out = []
    prev = 0
    for m in _ROW.finditer(xml):
        head = xml[m.start():xml.index(b">", m.start())]
        r = re.search(rb'\br="(\d+)"', head)
        n = int(r[1]) if r else prev + 1
        out.append((n, m.start(), m.end()))
        prev = n
    return out


def _merge_ranges(spans) -> list[tuple[int, int]]:
    spans = sorted((x, x) if isinstance(x, int) else tuple(x) for x in spans)
    out: list[tuple[int, int]] = []
    for a, b in spans:
        if out and a <= out[-1][1] + 1:
            out[-1] = (out[-1][0], max(out[-1][1], b))
        else:
            out.append((a, b))
    return out


def _in_rows(f: dict, a: int, b: int) -> bool:
    fid = str(f.get("cell") or f.get("id") or "")
    m = re.search(r"\$?(\d+)(?:@\d+)?$", fid)
    return True if not m else a <= int(m[1]) <= b


class _scan_budget:
    """窓の読み手が宣言を広げてよい数は、1 回の小さなブックごと（頼まれた範囲は切ってあるので、頼まれた範囲は必ず広げる）。"""

    def __enter__(self):
        return self

    def __exit__(self, *a):
        return False


def _trim_sheet(sh: dict, a: int, b: int, first_sheet: bool, first_page: bool, max_chars: int, md) -> tuple[list[str], int]:
    """シートの窓を、頼まれた行に絞り、字の数が max_chars に入るように行を切り詰める。(Markdown の行, 読み終えた最後の行)。"""
    all_lines = sh.get("reading", [])
    reading = [ln for ln in all_lines if _line_rows(ln)[1] >= a]
    before = [dict(ln, text="〔シートのはじめの行（見出しの手がかり）〕" + ln.get("text", "")) for ln in all_lines if _line_rows(ln)[1] < a][:1]
    fields = [f for f in sh.get("fields", []) if _in_rows(f, a, b)]
    keep = dict(sh)
    if not first_page:
        keep["shapes"] = []
        keep["header_footer"] = []
        keep["charts"] = []
        keep["objects"] = []
    keep["notes"] = [n for n in sh.get("notes", []) if _in_rows({"cell": n.get("cell")}, a, b)]
    keep["reading"] = before + reading
    keep["fields"] = fields
    lines = md(keep)
    size = sum(len(x) + 1 for x in lines)
    last = b
    if size > max_chars and len(reading) > 1:
        # 読み順の行を、字の数が入るまで減らす（欄も同じ行までに）
        head = size - sum(len(ln.get("text", "")) + 12 for ln in reading)
        room = max(max_chars - max(head, 0), max_chars // 4)
        kept = []
        used = 0
        for ln in reading:
            n = len(ln.get("text", "")) + 12
            if kept and used + n > room:
                break
            kept.append(ln)
            used += n
        last = _line_rows(kept[-1])[1]
        keep["reading"] = before + kept
        keep["fields"] = [f for f in fields if _in_rows(f, a, last)]
        lines = md(keep)
    return lines, last


def _line_rows(line: dict) -> tuple[int, int]:
    if "row" in line:
        return line["row"], line["row"]
    a, _, b = str(line.get("rows", "0-0")).partition("-")
    return int(a), int(b or a)


def stream_package(doc: XlsxLazy, sheets: dict[str, list[tuple[int, int, bytes]]], parts: dict[str, bytes], tmp: Path) -> dict:
    """新しいファイルを書く：部品の順はそのまま。変わらない部品は生のバイトのまま。変わったシートは、行のバイトを置き換えながら流して圧縮し直す。"""
    src = doc.src
    infos = sorted(doc.z.infolist(), key=lambda i: i.header_offset)
    h_old = hashlib.sha256()
    h_new = hashlib.sha256()
    for c in src.chunks():
        h_old.update(c)
    problems: list[str] = []
    same_ranges: list[tuple[str, int, int, int, str]] = []  # (部品, 元の始め, 元の終わり, 書いた先の始め, ハッシュ)
    new_info: dict[str, tuple[int, int, int, int]] = {}  # 名前 → (ローカルヘッダの位置, crc, 圧縮, 展開)
    raw_parts = 0
    raw_hash: dict[str, tuple[str, int, int]] = {}
    with open(tmp, "wb") as f:
        def put(b: bytes) -> None:
            h_new.update(b)
            f.write(b)

        at = 0
        for info in infos:
            m = doc.members[info.filename]
            # 部品の前の、知らないバイト（ふつうはない）
            if info.header_offset > at:
                for c in src.chunks(at, info.header_offset):
                    put(c)
            head = src.read(info.header_offset, m.data_at)
            end = m.data_at + m.compressed
            flag = struct.unpack("<H", head[6:8])[0]
            desc = 0
            if flag & 0x08:
                sig = src.read(end, end + 4)
                desc = 16 if sig == b"PK\x07\x08" else 12
            pos = f.tell()
            if info.filename in sheets or info.filename in parts:
                # 書き直す部品
                buf = _Deflater()
                local = bytearray(head)
                if info.filename in sheets:
                    ix = next(d["index"] for d in doc.sheet_meta if d["path"] == info.filename)
                    _stream_sheet(ix, sheets[info.filename], buf, same_ranges, info.filename)
                else:
                    buf.feed(parts[info.filename])
                data_path, crc, csize, usize = buf.finish()
                local[8:10] = struct.pack("<H", 8)
                local[14:26] = struct.pack("<III", crc, csize, usize) if not flag & 0x08 else local[14:26]
                put(bytes(local))
                with open(data_path, "rb") as t:
                    while True:
                        c = t.read(1 << 20)
                        if not c:
                            break
                        put(c)
                os.unlink(data_path)
                if flag & 0x08:
                    put((b"PK\x07\x08" if desc == 16 else b"") + struct.pack("<III", crc, csize, usize))
                new_info[info.filename] = (pos, crc, csize, usize)
            else:
                h = hashlib.sha256()
                for c in src.chunks(info.header_offset, end + desc):
                    h.update(c)
                    put(c)
                raw_hash[info.filename] = (h.hexdigest(), pos, end + desc - info.header_offset)
                raw_parts += 1
                new_info[info.filename] = (pos, info.CRC, info.compress_size, info.file_size)
            at = end + desc
        extra_new = [n for n in parts if n not in doc.members]
        if extra_new:
            problems.append(f"増えた部品があります：{extra_new[:5]}")
        # 中央ディレクトリ
        cd_at = f.tell()
        pos = doc.z.start_dir
        count = 0
        size_cd = 0
        for info in doc.z.infolist():
            rec = src.read(pos, pos + 46)
            n, m_, k = struct.unpack("<HHH", rec[28:34])
            full = bytearray(src.read(pos, pos + 46 + n + m_ + k))
            pos += 46 + n + m_ + k
            off, crc, cs, us = new_info[info.filename]
            full[42:46] = struct.pack("<I", off)
            if info.filename in sheets or info.filename in parts:
                full[10:12] = struct.pack("<H", 8)
                full[16:28] = struct.pack("<III", crc, cs, us)
            put(bytes(full))
            size_cd += len(full)
            count += 1
        eocd = bytearray(src.read(doc.z.start_dir + _cd_size(doc), src.size))
        i = eocd.find(b"PK\x05\x06")
        if i < 0:
            problems.append("ZIP の終わりの記録が見つかりません")
        else:
            eocd[i + 8:i + 12] = struct.pack("<HH", count & 0xFFFF, count & 0xFFFF)
            eocd[i + 12:i + 20] = struct.pack("<II", size_cd, cd_at)
        put(bytes(eocd))
        f.flush()
        os.fsync(f.fileno())
    # 書いたファイルを読み直して比べる
    problems += _verify_package(doc, tmp, same_ranges, raw_hash)
    return {"problems": problems, "old_sha256": h_old.hexdigest(), "new_sha256": h_new.hexdigest(), "ranges": len(same_ranges),
            "same_bytes": sum(b - a for _n, a, b, _d, _h in same_ranges), "raw_parts": raw_parts}


def _cd_size(doc: XlsxLazy) -> int:
    pos = doc.z.start_dir
    n = 0
    for _ in doc.z.infolist():
        rec = doc.src.read(pos, pos + 46)
        a, b, c = struct.unpack("<HHH", rec[28:34])
        pos += 46 + a + b + c
        n += 46 + a + b + c
    return n


class _Deflater:
    """圧縮して一時ファイルへ（大きさの分からない部品を、メモリに置かない）。"""

    def __init__(self):
        import tempfile

        fd, self.path = tempfile.mkstemp(prefix="todoku-part-")
        self.f = os.fdopen(fd, "wb")
        self.z = zlib.compressobj(6, zlib.DEFLATED, -15)
        self.crc = 0
        self.usize = 0
        self.csize = 0

    def feed(self, b: bytes) -> None:
        self.crc = zlib.crc32(b, self.crc)
        self.usize += len(b)
        c = self.z.compress(b)
        if c:
            self.f.write(c)
            self.csize += len(c)

    def finish(self) -> tuple[str, int, int, int]:
        c = self.z.flush()
        self.f.write(c)
        self.csize += len(c)
        self.f.close()
        return self.path, self.crc & 0xFFFFFFFF, self.csize, self.usize


def _stream_sheet(ix: SheetIndex, edits: list[tuple[int, int, bytes]], out: _Deflater, same: list, name: str) -> None:
    """シートの XML を展開しながら、edits の範囲だけを置き換えて out に流す。"""
    edits = sorted(edits, key=lambda e: (e[0], e[1]))
    k = 0
    pos = 0  # いま流している元の位置
    done = 0  # 書いた先の位置
    seg_start = 0
    seg_dst = 0
    h = hashlib.sha256()

    def close_seg(end: int) -> None:
        nonlocal h
        same.append((name, seg_start, end, seg_dst, h.hexdigest()))
        h = hashlib.sha256()

    stream = ix.m.chunks()
    for c in stream:
        at = 0
        while at < len(c):
            if k < len(edits):
                x, y, new = edits[k]
                if pos + (len(c) - at) <= x:
                    piece = c[at:]
                    h.update(piece)
                    out.feed(piece)
                    pos += len(piece)
                    done += len(piece)
                    at = len(c)
                    continue
                if pos < x:
                    piece = c[at:at + (x - pos)]
                    h.update(piece)
                    out.feed(piece)
                    at += len(piece)
                    pos += len(piece)
                    done += len(piece)
                # pos == x：編集の範囲を飛ばして、新しいバイトを入れる
                close_seg(x)
                out.feed(new)
                done += len(new)
                skip = y - x
                k += 1
                seg_start, seg_dst = y, done
                while skip > 0:
                    take = min(skip, len(c) - at)
                    at += take
                    pos += take
                    skip -= take
                    if skip > 0:
                        c = next(stream)  # 編集の範囲が枠をまたぐ
                        at = 0
            else:
                piece = c[at:]
                h.update(piece)
                out.feed(piece)
                pos += len(piece)
                done += len(piece)
                at = len(c)
    close_seg(pos)
    limits.progress("シートを書き出しました", pos, pos)


def _verify_package(doc: XlsxLazy, tmp: Path, same: list, raw_hash: dict) -> list[str]:
    bad: list[str] = []
    src2 = Source(str(tmp))
    try:
        with zipfile.ZipFile(tmp) as z2:
            names2 = [i.filename for i in z2.infolist()]
            if names2 != [i.filename for i in doc.z.infolist()]:
                bad.append("ZIP の部品の並びが変わりました")
                return bad
            for name, (want, pos, size) in raw_hash.items():
                h = hashlib.sha256()
                for c in src2.chunks(pos, pos + size):
                    h.update(c)
                if h.hexdigest() != want:
                    bad.append(f"変えていない部品 {name} のバイトが変わっています")
            by_part: dict[str, list] = {}
            for name, a, b, d, want in same:
                by_part.setdefault(name, []).append((d, d + (b - a), want))
            for name, spans in by_part.items():
                m = Member(src2, z2.getinfo(name))
                pos = 0
                k = 0
                h = hashlib.sha256()
                crc = 0
                for c in m.chunks():
                    crc = zlib.crc32(c, crc)
                    at = 0
                    while at < len(c) and k < len(spans):
                        a, b, want = spans[k]
                        if pos < a:
                            step = min(a - pos, len(c) - at)
                        else:
                            step = min(b - pos, len(c) - at)
                            h.update(c[at:at + step])
                        at += step
                        pos += step
                        if pos >= b and k < len(spans) and pos >= spans[k][1]:
                            if h.hexdigest() != want:
                                bad.append(f"{name}：書き換えていない範囲（{a:,} バイト目から）のハッシュが違います")
                            h = hashlib.sha256()
                            k += 1
                    pos += len(c) - at
                if crc & 0xFFFFFFFF != z2.getinfo(name).CRC:
                    bad.append(f"{name}：CRC が合いません")
    finally:
        src2.close()
    return bad
