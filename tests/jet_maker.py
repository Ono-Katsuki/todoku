"""試験用の Access のデータベースを、ページの形（mdbtools の HACKING・Jackcess）のとおりに手で作る（Jet3・Jet4）。"""

from __future__ import annotations

import struct


class Maker:
    def __init__(self, jet3: bool = False, code_page: int = 932):
        self.jet3 = jet3
        self.pg = 2048 if jet3 else 4096
        self.pages: list[bytes | None] = [None, None, None]  # 0：見出し、1：予備、2：MSysObjects の定義
        self.code_page = code_page
        self.catalog: list[tuple[int, str, int, int]] = []

    # ------------------------------------------------------------ 値
    def text(self, s: str, compress: bool = False) -> bytes:
        if self.jet3:
            return s.encode("cp932" if self.code_page == 932 else "cp1252")
        if compress:
            return b"\xff\xfe" + s.encode("latin-1")
        return s.encode("utf-16-le")

    def name(self, s: str) -> bytes:
        return bytes([len(s.encode("cp932"))]) + s.encode("cp932") if self.jet3 else struct.pack("<H", len(s) * 2) + s.encode("utf-16-le")

    # ------------------------------------------------------------ 行
    def row(self, cols: list[dict], values: dict) -> bytes:
        """cols：{name, type, num, var or None, fixed or None, size}。"""
        ncols = len(cols)
        fixed_bytes = bytearray(sum(c["size"] for c in cols if c.get("var") is None))
        var_vals: list[bytes] = []
        mask = bytearray((ncols + 7) // 8)
        for c in cols:
            v = values.get(c["name"])
            if c["type"] == 1:
                if v:
                    mask[c["num"] // 8] |= 1 << (c["num"] % 8)
                continue
            if c.get("var") is None:
                if v is not None:
                    fixed_bytes[c["fixed"]:c["fixed"] + c["size"]] = v
                    mask[c["num"] // 8] |= 1 << (c["num"] % 8)
            else:
                while len(var_vals) <= c["var"]:
                    var_vals.append(b"")
                if v is not None:
                    var_vals[c["var"]] = v
                    mask[c["num"] // 8] |= 1 << (c["num"] % 8)
        hdr = bytes([ncols]) if self.jet3 else struct.pack("<H", ncols)
        out = bytearray(hdr) + fixed_bytes
        offs = []
        for vv in var_vals:
            offs.append(len(out))
            out += vv
        offs.append(len(out))
        tail = bytearray()
        nvar = len(var_vals)
        if any(c.get("var") is not None for c in cols):
            if self.jet3:
                # 位置は下の桁の 1 バイト。上の桁が k に増える最初の列を、繰り上がりの表に書く
                carries = [next(i for i, x in enumerate(offs) if x >> 8 >= k) for k in range(1, (offs[-1] >> 8) + 1)]
                tail += bytes(reversed([x & 0xFF for x in offs])) + bytes(reversed(carries)) + bytes([nvar])
            else:
                for x in reversed(offs):
                    tail += struct.pack("<H", x)
                tail += struct.pack("<H", nvar)
        return bytes(out) + bytes(tail) + bytes(mask)

    def _data_page(self, tdef: int, rows: list[bytes], flags: list[int] | None = None) -> bytes:
        p = bytearray(self.pg)
        p[0], p[1] = 1, 1
        struct.pack_into("<I", p, 4, tdef)
        rc = 8 if self.jet3 else 12
        struct.pack_into("<H", p, rc, len(rows))
        end = self.pg
        for i, r in enumerate(rows):
            start = end - len(r)
            p[start:end] = r
            struct.pack_into("<H", p, rc + 2 + 2 * i, start | ((flags or [0] * len(rows))[i]))
            end = start
        return bytes(p)

    def _tdef(self, cols: list[dict], nrows: int, usage_dp: int) -> bytes:
        p = bytearray(self.pg)
        p[0], p[1] = 2, 1
        j3 = self.jet3
        tab_rows, tab_cols, tab_idxs, tab_ridxs, usage, start = (12, 25, 27, 31, 35, 43) if j3 else (16, 45, 47, 51, 55, 63)
        struct.pack_into("<I", p, tab_rows, nrows)
        struct.pack_into("<H", p, tab_cols, len(cols))
        struct.pack_into("<I", p, usage, usage_dp)
        pos = start
        entry = 18 if j3 else 25
        for c in cols:
            e = bytearray(entry)
            e[0] = c["type"]
            e[1 if j3 else 5] = c["num"]
            struct.pack_into("<H", e, 3 if j3 else 7, c.get("var") or 0)
            struct.pack_into("<H", e, 5 if j3 else 9, c["num"])
            e[13 if j3 else 15] = 0 if c.get("var") is not None else 1
            struct.pack_into("<H", e, 14 if j3 else 21, c.get("fixed") or 0)
            struct.pack_into("<H", e, 16 if j3 else 23, c["size"])
            e[9 if j3 else 11] = c.get("scale", 0)
            p[pos:pos + entry] = e
            pos += entry
        for c in cols:
            n = self.name(c["name"])
            p[pos:pos + len(n)] = n
            pos += len(n)
        return bytes(p)

    def _usage_page(self, dp_pages: list[int]) -> tuple[bytes, int]:
        """型 1 のページに、型 0（ページの範囲を直接持つ）の使う頁の表を 1 行で置く。"""
        start = 0
        bits = bytearray(32)
        for pg in dp_pages:
            bits[(pg - start) // 8] |= 1 << ((pg - start) % 8)
        m = b"\x00" + struct.pack("<I", start) + bytes(bits)
        return self._data_page(0, [m]), 0

    def add_table(self, name: str, cols: list[dict], rows: list[bytes], flags: list[int] | None = None, catalog: bool = True,
                  typ: int = 1, cat_flags: int = 0) -> int:
        idx_tdef = len(self.pages)
        self.pages.append(None)
        idx_usage = len(self.pages)
        self.pages.append(None)
        idx_data = len(self.pages)
        self.pages.append(self._data_page(idx_tdef, rows, flags))
        self.pages[idx_usage] = self._usage_page([idx_data])[0]
        self.pages[idx_tdef] = self._tdef(cols, len(rows), idx_usage << 8)
        if catalog:
            self.catalog.append((idx_tdef, name, typ, cat_flags))
        return idx_tdef

    CAT = [{"name": "Id", "type": 4, "num": 0, "fixed": 0, "size": 4}, {"name": "Type", "type": 3, "num": 1, "fixed": 4, "size": 2},
           {"name": "Flags", "type": 4, "num": 2, "fixed": 6, "size": 4}, {"name": "Name", "type": 10, "num": 3, "var": 0, "size": 255}]

    def build(self) -> bytes:
        rows = [self.row(self.CAT, {"Id": struct.pack("<i", tid), "Type": struct.pack("<h", typ), "Flags": struct.pack("<i", fl),
                                    "Name": self.text(nm)}) for tid, nm, typ, fl in self.catalog]
        usage_i = len(self.pages)
        self.pages.append(None)
        data_i = len(self.pages)
        self.pages.append(self._data_page(2, rows))
        self.pages[usage_i] = self._usage_page([data_i])[0]
        self.pages[2] = self._tdef(self.CAT, len(rows), usage_i << 8)
        self.pages[1] = bytes(self.pg)
        head = bytearray(self.pg)
        head[:4] = b"\x00\x01\x00\x00"
        head[4:19] = b"Standard Jet DB"
        head[0x14] = 0 if self.jet3 else 1
        body = bytearray(126 if self.jet3 else 128)
        struct.pack_into("<H", body, 0x3C - 0x18, self.code_page)
        from todoku.legacy.jet import _rc4

        head[0x18:0x18 + len(body)] = _rc4(struct.pack("<I", 0x6B39DAC7), bytes(body))
        self.pages[0] = bytes(head)
        return b"".join(self.pages)


def sample(jet3: bool = False) -> bytes:
    m = Maker(jet3)
    cols = [{"name": "番号", "type": 4, "num": 0, "fixed": 0, "size": 4}, {"name": "単価", "type": 7, "num": 1, "fixed": 4, "size": 8},
            {"name": "可", "type": 1, "num": 2, "size": 0, "fixed": 0}, {"name": "名前", "type": 10, "num": 3, "var": 0, "size": 100},
            {"name": "説明", "type": 12, "num": 4, "var": 1, "size": 0}, {"name": "日付", "type": 8, "num": 5, "fixed": 12, "size": 8},
            {"name": "金額", "type": 5, "num": 6, "fixed": 20, "size": 8}]
    def memo(s: str) -> bytes:
        t = m.text(s)
        return struct.pack("<II", len(t) | 0x80000000, 0) + bytes(4) + t
    rows = [m.row(cols, {"番号": struct.pack("<i", 1), "単価": struct.pack("<d", 12.5), "可": True, "名前": m.text("山田"), "説明": memo("メモの字"),
                         "日付": struct.pack("<d", 45000.0), "金額": struct.pack("<q", 12345)}),
            m.row(cols, {"番号": struct.pack("<i", 2), "名前": m.text("deleted row")}),
            m.row(cols, {"番号": struct.pack("<i", 3), "名前": m.text("abc", compress=not jet3)})]
    m.add_table("コード表", cols, rows, flags=[0, 0x8000, 0])
    return m.build()
