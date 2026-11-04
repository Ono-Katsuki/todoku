""".xlsb の空のセルへの書き込み：シートの部品にレコードを足し、位置の索引（binaryIndex）を合わせて直す。

- シート（MS-XLSB 2.1.7.62）：BrtRowHdr（行の昇順）の後ろに、列の昇順のセルのレコード。足すセルは BrtCellReal・BrtCellSt（字は共有文字列を変えずセルの中に持つ）・BrtCellBool
- BrtRowHdr（2.4.770）の rgBrtColspan：セルの列は、この範囲の和の中であること。外なら広げる。BrtWsDim（シートの使っている範囲）も広げる
- 索引の部品（Binary Index）：BrtIndexBlock（42。rwMic・rwMac、最大 32 行）→ BrtIndexRowBlock（40）→ … → BrtIndexPartEnd（41）。
  BrtIndexRowBlock（2.4.697）：grbitRowMask（4）・ibBaseOffset（8）・arrayColbitMask（行ごと 2。列の 1024 列ごとの有無）・arraySubBaseOffset（位置ごと 4）。
  位置は「ibBaseOffset + arraySubBaseOffset[x]」のシートの部品の中のバイトの位置
- 索引の直し方：すべての位置を、足したバイトのぶん写す。足したセルの行・列の範囲に印がなければ印と位置を足し、その範囲で最初のセルになるなら位置を差し替える。
  どのブロックにも入らない行なら、32 行のブロックを、行の順の所に足す

証明（prove）は、作り方とは別に、できた部品を読み直して確かめる。Excel で開いて確かめる手段はないので、結果の warnings に書く。
"""

from __future__ import annotations

import struct

from .book import ARR, BEGIN_DATA, C_BLANK, C_BOOL, C_REAL, C_ST, CELLS, END_DATA, ROW, SHR, records

WSDIM = 148
IDX_ROWBLOCK, IDX_END, IDX_BLOCK = 40, 41, 42
UNVERIFIED = ("空のセルへの書き込みは、シートのレコードを足し、位置の索引（binaryIndex）を仕様のとおりに直して、読み直して確かめました。"
              "Excel で開く確認は、とどくの手元ではできていません。提出の前に本人か発行元が Excel で開けることを確かめるか、.xlsx で保存し直したものに書いてください")


class GrowError(ValueError):
    pass


def rec(rid: int, body: bytes) -> bytes:
    out = bytearray()
    if rid < 0x80:
        out.append(rid)
    else:
        out += bytes(((rid & 0x7F) | 0x80, (rid >> 7) & 0x7F))
    n = len(body)
    while True:
        b = n & 0x7F
        n >>= 7
        out.append(b | 0x80 if n else b)
        if not n:
            return bytes(out) + body


class _Row:
    __slots__ = ("row", "start", "body", "ln", "cells", "end")

    def __init__(self, row, start, body, ln):
        self.row, self.start, self.body, self.ln = row, start, body, ln
        self.cells: list[tuple[int, int, int]] = []  # (列, レコードの始まり, レコードの終わり。SHR・ARR を含む)
        self.end = 0


def scan_sheet(data: bytes):
    """(行の並び, データの終わりの位置, BrtWsDim の中身の位置)。位置はレコードの見出しの始まり。"""
    rows: list[_Row] = []
    data_end = None
    dim = None
    begin = None
    prev = 0
    cur: _Row | None = None
    for rid, p, ln in records(data):
        start = prev
        prev = p + ln
        if rid == WSDIM and ln >= 16:
            dim = p
        elif rid == BEGIN_DATA:
            begin = prev
        elif rid == END_DATA:
            data_end = start
            if cur:
                cur.end = start
            cur = None
        elif rid == ROW:
            if cur:
                cur.end = start
            cur = _Row(struct.unpack_from("<I", data, p)[0], start, p, ln)
            rows.append(cur)
        elif rid in CELLS and cur is not None:
            cur.cells.append((struct.unpack_from("<I", data, p)[0], start, prev))
        elif rid in (SHR, ARR) and cur is not None and cur.cells:
            c, s, _e = cur.cells[-1]
            cur.cells[-1] = (c, s, prev)
    if data_end is None or begin is None:
        raise GrowError("シートのデータの区切り（BrtBeginSheetData・BrtEndSheetData）がありません")
    return rows, data_end, dim


def cell_record(col: int, xf: int, v) -> bytes:
    head = struct.pack("<II", col, xf & 0xFFFFFF)
    if isinstance(v, bool):
        return rec(C_BOOL, head + bytes((1 if v else 0,)))
    if isinstance(v, (int, float)):
        return rec(C_REAL, head + struct.pack("<d", float(v)))
    s = str(v)
    enc = s.encode("utf-16-le", "surrogatepass")
    return rec(C_ST, head + struct.pack("<I", len(enc) // 2) + enc)


def _spans_with(body: bytes, cols: list[int]) -> bytes | None:
    """BrtRowHdr の中身：cols が範囲の和に入るように。変わらなければ None。"""
    if len(body) < 17:
        raise GrowError("BrtRowHdr が短すぎます")
    n = struct.unpack_from("<I", body, 13)[0]
    if n > 16 or 17 + 8 * n > len(body):
        raise GrowError("BrtRowHdr の列の範囲が壊れています")
    if n == 0:
        return None  # 範囲を書いていない行（元のセルも範囲の外にならないよう、足さない）
    spans = [list(struct.unpack_from("<II", body, 17 + 8 * i)) for i in range(n)]
    changed = False
    for c in cols:
        if any(a <= c <= b for a, b in spans):
            continue
        changed = True
        best = min(spans, key=lambda s: min(abs(c - s[0]), abs(c - s[1])))
        best[0], best[1] = min(best[0], c), max(best[1], c)
    if not changed:
        return None
    spans.sort()
    merged: list[list[int]] = []
    for a, b in spans:  # 広げて重なった範囲は 1 つにする（範囲は重ならない）
        if merged and a <= merged[-1][1] + 1:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    return body[:13] + struct.pack("<I", len(merged)) + b"".join(struct.pack("<II", a, b) for a, b in merged) + body[17 + 8 * n:]


def plan(data: bytes, items: list[tuple[int, int, int, object]]) -> list[tuple[int, int, bytes]]:
    """items：(行（0 始まり）, 列（0 始まり）, 書式, 値)。(始め, 終わり, 置くバイト) の並び（重ならない）。"""
    rows, data_end, dim = scan_sheet(data)
    by_row = {r.row: r for r in rows}
    want: dict[int, dict[int, tuple[int, object]]] = {}
    for r, c, xf, v in items:
        want.setdefault(r, {})[c] = (xf, v)
    edits: dict[tuple[int, int], bytearray] = {}

    def put(s: int, e: int, b: bytes) -> None:
        edits.setdefault((s, e), bytearray()).extend(b)

    for r, wcols in sorted(want.items()):
        row = by_row.get(r)
        if row is None:
            continue
        for c in sorted(wcols):
            xf, v = wcols[c]
            hit = next((x for x in row.cells if x[0] == c), None)
            if hit is not None:
                put(hit[1], hit[2], cell_record(c, xf, v))
                continue
            pos = next((x[1] for x in row.cells if x[0] > c), row.end)
            put(pos, pos, cell_record(c, xf, v))
        nb = _spans_with(data[row.body:row.body + row.ln], sorted(wcols))
        if nb is not None:
            hdr_end = row.body + row.ln
            put(row.start, hdr_end, rec(ROW, nb))
    # 新しい行は、ある行の後ろに足す字より後に（同じ位置に足すとき、前の行の字が新しい行に入らないように）
    new_rows = sorted(r for r in want if r not in by_row)
    for r in new_rows:
        pos = next((x.start for x in rows if x.row > r), data_end)
        cols = sorted(want[r])
        body = struct.pack("<IIHBBBI", r, 0, 0, 0, 0, 0, 1) + struct.pack("<II", cols[0], cols[-1])
        put(pos, pos, rec(ROW, body) + b"".join(cell_record(c, *want[r][c]) for c in cols))
    if dim is not None:
        r1, r2, c1, c2 = struct.unpack_from("<IIII", data, dim)
        empty = r1 > r2 or c1 > c2
        rs = sorted(want)
        cs = sorted({c for w in want.values() for c in w})
        nr1, nr2 = (rs[0], rs[-1]) if empty else (min(r1, rs[0]), max(r2, rs[-1]))
        nc1, nc2 = (cs[0], cs[-1]) if empty else (min(c1, cs[0]), max(c2, cs[-1]))
        if (nr1, nr2, nc1, nc2) != (r1, r2, c1, c2):
            put(dim, dim + 16, struct.pack("<IIII", nr1, nr2, nc1, nc2))
    return sorted(((s, e, bytes(b)) for (s, e), b in edits.items()), key=lambda x: (x[0], x[1]))


def splice(data: bytes, edits: list[tuple[int, int, bytes]]) -> bytes:
    out = bytearray()
    p = 0
    for s, e, b in edits:
        if s < p:
            raise GrowError("書き込みの範囲が重なっています")
        out += data[p:s]
        out += b
        p = e
    out += data[p:]
    return bytes(out)


def map_pos(pos: int, edits: list[tuple[int, int, bytes]]) -> int:
    """元の位置 → 新しい位置（挿入は、その位置のレコードを後ろへ送る）。"""
    delta = 0
    for s, e, b in edits:
        if s == e:
            if s <= pos:
                delta += len(b)
        elif e <= pos:
            delta += len(b) - (e - s)
        elif s <= pos:
            return s + delta  # 置き換えたレコードの中：その始まり
    return pos + delta


def read_index(ix: bytes) -> list[dict]:
    """[{rid: 42, mic, mac, body} ・ {rid: 40, rows: {(行の添字, 列の範囲の添字): 絶対の位置}, base}・ほかのレコード（そのまま）]。"""
    out: list[dict] = []
    for rid, p, ln in records(ix):
        b = ix[p:p + ln]
        if rid == IDX_BLOCK:
            if ln < 16:
                raise GrowError("BrtIndexBlock が短すぎます")
            mic, mac = struct.unpack_from("<II", b, 0)
            out.append({"rid": rid, "mic": mic, "mac": mac, "raw": b})
        elif rid == IDX_ROWBLOCK:
            if ln < 12:
                raise GrowError("BrtIndexRowBlock が短すぎます")
            mask, base = struct.unpack_from("<IQ", b, 0)
            rows = [i for i in range(32) if mask >> i & 1]
            q = 12
            if q + 2 * len(rows) > ln:
                raise GrowError("BrtIndexRowBlock の列の印が足りません")
            cols = struct.unpack_from(f"<{len(rows)}H", b, q)
            q += 2 * len(rows)
            pos = {}
            for ri, cm in zip(rows, cols):
                for ci in range(16):
                    if cm >> ci & 1:
                        if q + 4 > ln:
                            raise GrowError("BrtIndexRowBlock の位置が足りません")
                        pos[(ri, ci)] = base + struct.unpack_from("<I", b, q)[0]
                        q += 4
            out.append({"rid": rid, "base": base, "rows": pos})
        else:
            out.append({"rid": rid, "raw": b})
    return out


def write_index(items: list[dict]) -> bytes:
    out = bytearray()
    for it in items:
        if it["rid"] == IDX_ROWBLOCK:
            pos = it["rows"]
            base = it["base"]
            rows = sorted({r for r, _c in pos})
            mask = sum(1 << r for r in rows)
            body = struct.pack("<IQ", mask, base)
            for r in rows:
                body += struct.pack("<H", sum(1 << c for rr, c in pos if rr == r))
            for key in sorted(pos):
                body += struct.pack("<I", pos[key] - base)
            out += rec(IDX_ROWBLOCK, body)
        else:
            out += rec(it["rid"], it["raw"])
    return bytes(out)


def cell_positions(data: bytes) -> dict[tuple[int, int], int]:
    """(行, 列) → セルのレコードの始まり（0 始まり）。"""
    out = {}
    row = -1
    prev = 0
    for rid, p, ln in records(data):
        start = prev
        prev = p + ln
        if rid == ROW:
            row = struct.unpack_from("<I", data, p)[0]
        elif rid in CELLS and rid != C_BLANK:
            out[(row, struct.unpack_from("<I", data, p)[0])] = start
    return out


def rewrite_index(ix: bytes, old_edits: list[tuple[int, int, bytes]], new_sheet: bytes, written: list[tuple[int, int]]) -> bytes:
    items = read_index(ix)
    blocks: list[tuple[int, int, dict]] = []  # 行の範囲と、対応する BrtIndexRowBlock
    i = 0
    while i < len(items):
        it = items[i]
        if it["rid"] == IDX_ROWBLOCK:
            it["base"] = map_pos(it["base"], old_edits)
            it["rows"] = {k: map_pos(v, old_edits) for k, v in it["rows"].items()}
            prev = items[i - 1] if i else None
            if prev is None or prev["rid"] != IDX_BLOCK:
                raise GrowError("BrtIndexRowBlock の前に BrtIndexBlock がありません")
            blocks.append((prev["mic"], prev["mac"], it))
        i += 1
    where = cell_positions(new_sheet)
    for r, c in written:
        pos = where[(r, c)]
        key_col = c // 1024
        hit = next((b for b in blocks if b[0] <= r < b[1]), None)
        if hit is None:
            mic = r - r % 32
            # 隣のブロックと重ならない範囲にする（ブロックは最大 32 行）
            lower = max((b[1] for b in blocks if b[1] <= r), default=0)
            upper = min((b[0] for b in blocks if b[0] > r), default=1048576)
            mic = max(mic, lower)
            mac = min(mic + 32, upper, 1048576)
            block = {"rid": IDX_BLOCK, "mic": mic, "mac": mac, "raw": struct.pack("<IIII", mic, mac, 0, 0) + bytes(((mac - mic + 32) // 32) * 4)}
            rowb = {"rid": IDX_ROWBLOCK, "base": pos, "rows": {}}
            at = next((k for k, x in enumerate(items) if x["rid"] == IDX_BLOCK and x["mic"] > mic), None)
            if at is None:
                at = next((k for k, x in enumerate(items) if x["rid"] == IDX_END), len(items))
            items[at:at] = [block, rowb]
            hit = (mic, mac, rowb)
            blocks.append(hit)
        mic, _mac, rb = hit
        key = (r - mic, key_col)
        old = rb["rows"].get(key)
        if old is None or pos < old:
            rb["rows"][key] = pos
        if pos < rb["base"]:
            rb["base"] = pos
    for _a, _b, rb in blocks:
        if rb["rows"] and rb["base"] > min(rb["rows"].values()):
            rb["base"] = min(rb["rows"].values())
    return write_index(items)


# ====================================================================== 証明（読み直して確かめる）
def prove_sheet(old: bytes, new: bytes, written: set[tuple[int, int]], value_spans: list[tuple[int, int]]) -> list[str]:
    """元のレコードがすべて、順のまま新しい部品にあるか（書いたセル・その行の見出し・BrtWsDim・計算し直した値の欄のほかは、バイト一致）。行と列の順。"""
    problems: list[str] = []
    new_recs = []
    prev = 0
    for rid, p, ln in records(new):
        new_recs.append((rid, prev, p, ln))
        prev = p + ln
    i = 0
    row = -1
    prev = 0
    for rid, p, ln in records(old):
        start = prev
        prev = p + ln
        raw = old[start:prev]
        if rid == ROW:
            row = struct.unpack_from("<I", old, p)[0]
        # 足したレコード（書いたセル・足した行）を飛ばす
        while i < len(new_recs):
            nr = new_recs[i]
            if rid in CELLS and rid == C_BLANK and (row, struct.unpack_from("<I", old, p)[0]) in written:
                break
            if nr[0] == rid and new[nr[1]:nr[2] + nr[3]] == raw:
                break
            if nr[0] == rid == WSDIM or (nr[0] == rid == ROW and struct.unpack_from("<I", new, nr[2])[0] == row):
                break
            if nr[0] == rid and rid in CELLS and new[nr[2]:nr[2] + 8] == old[p:p + 8]:  # 数式の保存値を直した所
                break
            if nr[0] in (ROW, C_REAL, C_ST, C_BOOL) and _added_ok(new, nr, written):
                i += 1
                continue
            problems.append(f"元のレコード（型 {rid}・位置 {start}）が、新しいシートに順のままありません")
            return problems
        if i >= len(new_recs):
            problems.append(f"元のレコード（型 {rid}・位置 {start}）がなくなっています")
            return problems
        nr = new_recs[i]
        if rid == C_BLANK and (row, struct.unpack_from("<I", old, p)[0]) in written:
            i += 1
            continue
        got = new[nr[1]:nr[2] + nr[3]]
        if got != raw:
            if rid == ROW:
                if nr[3] < 13 or new[nr[2]:nr[2] + 13] != old[p:p + 13]:
                    problems.append(f"行の見出し（{row + 1} 行）の、列の範囲のほかが変わっています")
            elif rid == WSDIM:
                pass
            elif rid in CELLS:
                diffs = [k for k in range(ln) if old[p + k] != new[nr[2] + k]] if ln == nr[3] else [-1]
                if -1 in diffs or not all(any(a <= p + k < b for a, b in value_spans) for k in diffs):
                    problems.append(f"セルのレコード（{row + 1} 行）の、値のほかが変わっています")
            else:
                problems.append(f"レコード（型 {rid}）のバイトが変わっています")
        i += 1
    while i < len(new_recs):
        if not _added_ok(new, new_recs[i], written):
            problems.append("書いたセルのほかにレコードが増えています")
            break
        i += 1
    # 行と列の順・列の範囲
    rows, _e, _d = scan_sheet(new)
    last = -1
    for r in rows:
        if r.row <= last:
            problems.append("行の順が崩れています")
        last = r.row
        cols = [c for c, _s, _x in r.cells]
        if cols != sorted(set(cols)):
            problems.append(f"{r.row + 1} 行のセルの順が崩れています")
        n = struct.unpack_from("<I", new, r.body + 13)[0] if r.ln >= 17 else 0
        spans = [struct.unpack_from("<II", new, r.body + 17 + 8 * k) for k in range(n)]
        if spans and any(not any(a <= c <= b for a, b in spans) for c in cols):
            problems.append(f"{r.row + 1} 行のセルが、行の見出しの列の範囲の外にあります")
    return problems


def _added_ok(new: bytes, nr, written: set[tuple[int, int]]) -> bool:
    rid, _s, p, _ln = nr
    if rid == ROW:
        r = struct.unpack_from("<I", new, p)[0]
        return any(x[0] == r for x in written)
    return rid in (C_REAL, C_ST, C_BOOL)


def prove_index(old_ix: bytes, new_ix: bytes, new_sheet: bytes, old_sheet: bytes, written: set[tuple[int, int]]) -> list[str]:
    """新しい索引のすべての位置がセルのレコードの始まりで、行と列の範囲が合うか。元の索引の印がすべて残り、増えた印は書いたセルのものか。"""
    problems = []
    starts = {}
    row = -1
    prev = 0
    for rid, p, ln in records(new_sheet):
        s = prev
        prev = p + ln
        if rid == ROW:
            row = struct.unpack_from("<I", new_sheet, p)[0]
            starts[s] = (rid, row, None)
        elif rid in CELLS:
            starts[s] = (rid, row, struct.unpack_from("<I", new_sheet, p)[0])

    def marks(items):
        out = {}
        mic = None
        for it in items:
            if it["rid"] == IDX_BLOCK:
                mic = it["mic"]
            elif it["rid"] == IDX_ROWBLOCK:
                for (ri, ci), pos in it["rows"].items():
                    out[(mic + ri, ci)] = pos
        return out

    new_m = marks(read_index(new_ix))
    old_m = marks(read_index(old_ix))
    for key, pos in new_m.items():
        hit = starts.get(pos)
        if hit is None:
            problems.append(f"索引の位置 {pos} が、シートのレコードの始まりではありません")
        elif hit[1] != key[0] or (hit[2] is not None and hit[2] // 1024 != key[1]):
            problems.append(f"索引の位置 {pos} が、印の行・列の範囲のセルを指していません")
    for key in old_m:
        if key not in new_m:
            problems.append(f"索引の印（{key[0] + 1} 行）がなくなっています")
    extra = {k for k in new_m if k not in old_m}
    want = {(r, c // 1024) for r, c in written}
    if extra - want:
        problems.append("書いたセルのほかの印が索引に増えています")
    return problems
