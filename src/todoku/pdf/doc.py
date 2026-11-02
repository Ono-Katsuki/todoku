"""PDFの文書を読む：相互参照（xref 表・xref ストリーム・追記の連なり）、オブジェクト、ストリームの展開、暗号、ページ。

元のバイト列には手を加えない。書き込みは、末尾への追記（増分更新）で行う（form.py）。
"""

from __future__ import annotations

import re
import struct
import zlib

from .objects import Lexer, Name, PdfError, PdfString, Ref, Stream, parse_object


class PdfDocument:
    kind = "pdf"

    def __init__(self, blob: bytes, filename: str = ""):
        if not blob.lstrip()[:5] == b"%PDF-":
            raise PdfError("PDFではありません")
        self.blob = blob
        self.filename = filename
        self.xref: dict[int, tuple] = {}  # 番号 → ("o", 位置, 世代) / ("s", ストリームの番号, 何番目)
        self.trailer: dict = {}
        self.cache: dict[int, object] = {}
        self.objstm_cache: dict[int, list] = {}
        self.decryptor = None
        self.repaired = False
        self._read_xref()
        enc = self.trailer.get("Encrypt")
        if enc is not None:
            from .crypto import Decryptor

            encd = self.resolve(enc) if isinstance(enc, Ref) else enc
            ids = self.trailer.get("ID") or [b""]
            self.encrypt_ref = enc if isinstance(enc, Ref) else None
            self.decryptor = Decryptor(encd, bytes(ids[0]) if ids else b"")

    # ------------------------------------------------------------------ 相互参照
    def _read_xref(self) -> None:
        tail = self.blob[-2048:]
        m = list(re.finditer(rb"startxref\s+(\d+)", tail))
        try:
            if not m:
                raise PdfError("startxref がありません")
            pos = int(m[-1][1])
            seen = set()
            while pos is not None and pos not in seen and 0 <= pos < len(self.blob):
                seen.add(pos)
                pos = self._xref_section(pos)
            if not self.xref:
                raise PdfError("相互参照が空です")
        except Exception:  # noqa: BLE001
            self._repair()

    def _xref_section(self, pos: int):
        lx = Lexer(self.blob, pos)
        lx.skip()
        if self.blob.startswith(b"xref", lx.pos):
            lx.pos += 4
            while True:
                lx.skip()
                if self.blob.startswith(b"trailer", lx.pos):
                    lx.pos += 7
                    break
                k1, start = lx.token()
                k2, count = lx.token()
                if k1 != "num" or k2 != "num":
                    raise PdfError("xref 表が壊れています")
                lx.skip()
                for i in range(count):
                    m = re.match(rb"\s*(\d{1,10})\s+(\d{1,5})\s+([nf])", self.blob[lx.pos:lx.pos + 40])
                    if not m:
                        raise PdfError("xref 表の行が壊れています")
                    lx.pos += m.end()
                    num = start + i
                    if num not in self.xref and m[3] == b"n":
                        self.xref[num] = ("o", int(m[1]), int(m[2]))
                    elif num not in self.xref:
                        self.xref[num] = ("f", 0, 0)
            tr = parse_object(lx)
            for k, v in tr.items():
                self.trailer.setdefault(k, v)
            if "XRefStm" in tr:
                self._xref_section(int(tr["XRefStm"]))
            return int(tr["Prev"]) if "Prev" in tr else None
        # xref ストリーム
        obj = self._parse_at(pos)
        if not isinstance(obj, Stream) or getattr(obj.dict.get("Type"), "name", "") != "XRef":
            raise PdfError("xref ストリームではありません")
        d = obj.dict
        data = self.decode(obj, decrypt=False)
        w = [int(x) for x in d["W"]]
        size = int(d.get("Size", 0))
        index = [int(x) for x in d.get("Index", [0, size])]
        step = sum(w)
        p = 0
        for a in range(0, len(index), 2):
            start, count = index[a], index[a + 1]
            for i in range(count):
                fields = []
                for width in w:
                    fields.append(int.from_bytes(data[p:p + width], "big") if width else None)
                    p += width
                t = fields[0] if w[0] else 1
                num = start + i
                if num in self.xref:
                    continue
                if t == 1:
                    self.xref[num] = ("o", fields[1], fields[2] or 0)
                elif t == 2:
                    # 幅 0 の欄は、既定の値 0 を使う（PDF 32000 7.5.8.2）
                    self.xref[num] = ("s", fields[1] or 0, fields[2] or 0)
                else:
                    self.xref[num] = ("f", 0, 0)
        for k, v in d.items():
            if k not in ("Type", "W", "Index", "Length", "Filter", "DecodeParms"):
                self.trailer.setdefault(k, v)
        del step
        return int(d["Prev"]) if "Prev" in d else None

    def _repair(self) -> None:
        """相互参照が壊れているとき：ファイルを頭から読み、「n g obj」の位置を集め直す。"""
        self.repaired = True
        self.xref = {}
        for m in re.finditer(rb"(?<![\d])(\d{1,10})\s+(\d{1,5})\s+obj\b", self.blob):
            self.xref[int(m[1])] = ("o", m.start(), int(m[2]))
        for m in re.finditer(rb"trailer\s*<<", self.blob):
            try:
                tr = parse_object(Lexer(self.blob, m.end() - 2))
                self.trailer.update(tr)
            except Exception:  # noqa: BLE001
                pass
        if "Root" not in self.trailer:
            for num in self.xref:
                o = self.get(num)
                if isinstance(o, dict) and getattr(o.get("Type"), "name", "") == "Catalog":
                    self.trailer["Root"] = Ref(num, 0)
                    break

    # ------------------------------------------------------------------ オブジェクト
    def _parse_at(self, pos: int):
        lx = Lexer(self.blob, pos)
        k1, num = lx.token()
        k2, gen = lx.token()
        k3, kw = lx.token()
        if (k1, k2, k3, kw) != ("num", "num", "kw", "obj"):
            raise PdfError(f"位置 {pos} にオブジェクトがありません")
        obj = parse_object(lx)
        lx.skip()
        if isinstance(obj, dict) and self.blob.startswith(b"stream", lx.pos):
            p = lx.pos + 6
            if self.blob[p:p + 2] == b"\r\n":
                p += 2
            elif self.blob[p:p + 1] in (b"\n", b"\r"):
                p += 1
            length = obj.get("Length")
            if isinstance(length, Ref):
                try:
                    length = self.resolve(length)
                except Exception:  # noqa: BLE001
                    length = None
            raw = None
            if isinstance(length, int) and self.blob[p + length:p + length + 20].lstrip().startswith(b"endstream"):
                raw = self.blob[p:p + length]
            if raw is None:
                e = self.blob.find(b"endstream", p)
                raw = self.blob[p:e].rstrip(b"\r\n") if e >= 0 else self.blob[p:]
            return Stream(obj, raw, num, gen)
        return obj

    def get(self, num: int):
        if num in self.cache:
            return self.cache[num]
        ent = self.xref.get(num)
        obj = None
        if ent is None or ent[0] == "f":
            obj = None
        elif ent[0] == "o":
            try:
                obj = self._parse_at(ent[1])
            except Exception:  # noqa: BLE001
                obj = None
            if obj is not None and self.decryptor is not None and num != getattr(getattr(self, "encrypt_ref", None), "num", -1):
                obj = self._decrypt_strings(obj, num, ent[2])
        else:
            objs = self._objstm(ent[1])
            obj = objs[ent[2]] if ent[2] < len(objs) else None
        self.cache[num] = obj
        return obj

    def _objstm(self, num: int) -> list:
        if num in self.objstm_cache:
            return self.objstm_cache[num]
        st = self.get(num)
        out = []
        if isinstance(st, Stream):
            data = self.decode(st)
            n = int(st.dict.get("N", 0))
            first = int(st.dict.get("First", 0))
            lx = Lexer(data, 0)
            offs = []
            for _ in range(n):
                _k, onum = lx.token()
                _k, off = lx.token()
                offs.append(off)
            for off in offs:
                try:
                    out.append(parse_object(Lexer(data, first + off)))
                except Exception:  # noqa: BLE001
                    out.append(None)
        self.objstm_cache[num] = out
        return out

    def resolve(self, o):
        guard = 0
        while isinstance(o, Ref) and guard < 32:
            o = self.get(o.num)
            guard += 1
        return o

    def _decrypt_strings(self, obj, num, gen):
        dec = self.decryptor
        if isinstance(obj, PdfString):
            out = PdfString(dec.decrypt(bytes(obj), num, gen))
            out.hexform = obj.hexform
            return out
        if isinstance(obj, list):
            return [self._decrypt_strings(x, num, gen) for x in obj]
        if isinstance(obj, dict):
            return {k: self._decrypt_strings(v, num, gen) for k, v in obj.items()}
        if isinstance(obj, Stream):
            obj.dict = self._decrypt_strings(obj.dict, num, gen)
            return obj
        return obj

    # ------------------------------------------------------------------ ストリーム
    def decode(self, st: Stream, decrypt: bool = True) -> bytes:
        data = st.raw
        if decrypt and self.decryptor is not None and getattr(st.dict.get("Type"), "name", "") != "XRef":
            data = self.decryptor.decrypt(data, st.num, st.gen)
        filters = st.dict.get("Filter")
        parms = st.dict.get("DecodeParms")
        if filters is None:
            return data
        if not isinstance(filters, list):
            filters = [filters]
            parms = [parms]
        if not isinstance(parms, list):
            parms = [parms] * len(filters)
        for f, prm in zip(filters, parms + [None] * (len(filters) - len(parms))):
            f = self.resolve(f)
            prm = self.resolve(prm) if prm is not None else None
            name = getattr(f, "name", "")
            if name in ("FlateDecode", "Fl"):
                data = _inflate(data)
                data = _predict(data, prm)
            elif name in ("ASCIIHexDecode", "AHx"):
                h = re.sub(rb"\s", b"", data).split(b">")[0]
                data = bytes.fromhex((h + b"0" * (len(h) % 2)).decode("latin-1"))
            elif name in ("ASCII85Decode", "A85"):
                data = _a85(data)
            elif name in ("LZWDecode", "LZW"):
                data = _predict(_lzw(data), prm)
            elif name in ("RunLengthDecode", "RL"):
                data = _rle(data)
            else:
                return data  # 画像の圧縮（DCT など）はそのまま
        return data

    # ------------------------------------------------------------------ 文書の構造
    @property
    def root(self) -> dict:
        return self.resolve(self.trailer.get("Root")) or {}

    def pages(self) -> list[dict]:
        out: list[dict] = []

        def walk(node_ref, inherited, depth=0):
            node = self.resolve(node_ref)
            if not isinstance(node, dict) or depth > 64:
                return
            inh = dict(inherited)
            for k in ("Resources", "MediaBox", "CropBox", "Rotate"):
                if k in node:
                    inh[k] = node[k]
            t = getattr(node.get("Type"), "name", "")
            if t == "Pages" or ("Kids" in node and t != "Page"):
                for kid in self.resolve(node.get("Kids")) or []:
                    walk(kid, inh, depth + 1)
            else:
                page = dict(inh)
                page.update(node)
                page["_ref"] = node_ref if isinstance(node_ref, Ref) else None
                out.append(page)

        walk(self.root.get("Pages"), {})
        return out

    def page_content(self, page: dict) -> bytes:
        c = self.resolve(page.get("Contents"))
        if c is None:
            return b""
        if isinstance(c, Stream):
            return self.decode(c)
        if isinstance(c, list):
            return b"\n".join(self.decode(self.resolve(x)) for x in c if isinstance(self.resolve(x), Stream))
        return b""

    def info(self) -> dict:
        return self.resolve(self.trailer.get("Info")) or {}

    @property
    def max_num(self) -> int:
        return max([int(self.trailer.get("Size", 0)) - 1] + list(self.xref))


def _inflate(data: bytes) -> bytes:
    try:
        return zlib.decompress(data)
    except zlib.error:
        # 末尾が壊れている圧縮データ：読めるところまで読む
        d = zlib.decompressobj()
        out = bytearray()
        try:
            for i in range(0, len(data), 1024):
                out += d.decompress(data[i:i + 1024])
        except zlib.error:
            pass
        if not out:
            try:
                return zlib.decompress(data, -15)
            except zlib.error:
                return bytes(out)
        return bytes(out)


def _predict(data: bytes, prm) -> bytes:
    if not isinstance(prm, dict):
        return data
    pred = int(prm.get("Predictor", 1))
    if pred < 10:
        if pred == 2:
            colors = int(prm.get("Colors", 1))
            bpc = int(prm.get("BitsPerComponent", 8))
            cols = int(prm.get("Columns", 1))
            if bpc == 8:
                rowlen = colors * cols
                out = bytearray(data)
                for r in range(0, len(out), rowlen):
                    for i in range(r + colors, min(r + rowlen, len(out))):
                        out[i] = (out[i] + out[i - colors]) & 0xFF
                return bytes(out)
        return data
    colors = int(prm.get("Colors", 1))
    bpc = int(prm.get("BitsPerComponent", 8))
    cols = int(prm.get("Columns", 1))
    bpp = max(1, colors * bpc // 8)
    rowlen = (colors * bpc * cols + 7) // 8
    out = bytearray()
    prev = bytearray(rowlen)
    i = 0
    while i + 1 + rowlen <= len(data) + 0:
        ft = data[i]
        row = bytearray(data[i + 1:i + 1 + rowlen])
        if ft == 1:
            for x in range(bpp, rowlen):
                row[x] = (row[x] + row[x - bpp]) & 0xFF
        elif ft == 2:
            for x in range(rowlen):
                row[x] = (row[x] + prev[x]) & 0xFF
        elif ft == 3:
            for x in range(rowlen):
                left = row[x - bpp] if x >= bpp else 0
                row[x] = (row[x] + ((left + prev[x]) >> 1)) & 0xFF
        elif ft == 4:
            for x in range(rowlen):
                a = row[x - bpp] if x >= bpp else 0
                b = prev[x]
                c = prev[x - bpp] if x >= bpp else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                pr = a if pa <= pb and pa <= pc else (b if pb <= pc else c)
                row[x] = (row[x] + pr) & 0xFF
        out += row
        prev = row
        i += 1 + rowlen
    return bytes(out)


def _a85(data: bytes) -> bytes:
    d = re.sub(rb"\s", b"", data)
    if d.startswith(b"<~"):
        d = d[2:]
    d = d.split(b"~>")[0]
    out = bytearray()
    group = []
    for ch in d:
        if ch == ord("z") and not group:
            out += b"\x00\x00\x00\x00"
            continue
        group.append(ch - 33)
        if len(group) == 5:
            n = 0
            for g in group:
                n = n * 85 + g
            out += struct.pack(">I", n & 0xFFFFFFFF)
            group = []
    if group:
        k = len(group)
        group += [84] * (5 - k)
        n = 0
        for g in group:
            n = n * 85 + g
        out += struct.pack(">I", n & 0xFFFFFFFF)[:k - 1]
    return bytes(out)


def _lzw(data: bytes) -> bytes:
    out = bytearray()
    table = [bytes([i]) for i in range(256)] + [b"", b""]
    bits = 9
    buf = 0
    nbits = 0
    prev = None
    for byte in data:
        buf = (buf << 8) | byte
        nbits += 8
        while nbits >= bits:
            code = (buf >> (nbits - bits)) & ((1 << bits) - 1)
            nbits -= bits
            if code == 256:
                table = table[:258]
                bits = 9
                prev = None
                continue
            if code == 257:
                return bytes(out)
            if prev is None:
                entry = table[code]
            elif code < len(table):
                entry = table[code]
                table.append(prev + entry[:1])
            else:
                entry = prev + prev[:1]
                table.append(entry)
            out += entry
            prev = entry
            if len(table) + 1 >= (1 << bits) and bits < 12:
                bits += 1
    return bytes(out)


def _rle(data: bytes) -> bytes:
    out = bytearray()
    i = 0
    while i < len(data):
        n = data[i]
        if n == 128:
            break
        if n < 128:
            out += data[i + 1:i + 2 + n]
            i += 2 + n
        else:
            out += data[i + 1:i + 2] * (257 - n)
            i += 2
    return bytes(out)
