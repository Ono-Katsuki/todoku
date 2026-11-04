"""PDFの文書を読む：相互参照（xref 表・xref ストリーム・追記の連なり）、オブジェクト、ストリームの展開、暗号、ページ。

元のバイト列には手を加えない。書き込みは、末尾への追記（増分更新）で行う（form.py）。
"""

from __future__ import annotations

import re
import struct

# パスワードの例外（NeedsPassword・WrongPassword）は、ここからも使えるようにする
from .crypto import Decryptor, NeedsPassword, WrongPassword  # noqa: F401
from .objects import Lexer, Name, PdfError, PdfString, Ref, Stream, parse_object


class PdfDocument:
    kind = "pdf"
    light = False  # 大きな PDF（lazy）：大きなストリームは控えない（メモリがファイルの大きさに比例しないように）

    def __init__(self, blob: bytes, filename: str = "", password: str | bytes | None = None):
        """password：開くためのパスワード（ユーザー）か、権限のパスワード（所有者）。空のパスワードで開けるPDFには要らない。
        要るのに渡されていなければ NeedsPassword、合わなければ WrongPassword（NeedsPassword の一種）。パスワードは持たない。"""
        head = blob.lstrip()[:5] if isinstance(blob, (bytes, bytearray)) else blob[:1 << 16].lstrip()[:5]  # FileBytes：頭だけ
        if not head == b"%PDF-":
            raise PdfError("PDFではありません")
        self.blob = blob
        self.filename = filename
        self.xref: dict[int, tuple] = {}  # 番号 → ("o", 位置, 世代) / ("s", ストリームの番号, 何番目)
        self.trailer: dict = {}
        self.cache: dict[int, object] = {}
        self.objstm_cache: dict[int, list] = {}
        self.decryptor = None
        self.encrypt_ref = None
        self.repaired = False
        self._read_xref()
        enc = self.trailer.get("Encrypt")
        if enc is not None:
            failure = self._unlock(enc, password)
            password = None  # 例外の記録（トレースバックの変数）にも、パスワードを残さない
            if failure is not None:
                raise failure

    def _unlock(self, enc, password) -> NeedsPassword | None:
        """暗号の鍵を求める。パスワードが要る・合わないときは、新しく作った例外を返す（元の例外の記録には、パスワードを持つ変数が残るため）。"""
        encd = self.resolve(enc) if isinstance(enc, Ref) else enc
        if not isinstance(encd, dict):
            raise PdfError("暗号の辞書（Encrypt）が読めません")
        ids = self.resolve(self.trailer.get("ID"))
        first = self.resolve(ids[0]) if isinstance(ids, list) and ids else b""
        self.encrypt_ref = enc if isinstance(enc, Ref) else None
        try:
            self.decryptor = Decryptor(encd, bytes(first) if isinstance(first, bytes) else b"", password, self.resolve)
        except NeedsPassword as e:
            return type(e)()
        # 鍵の決まる前に読んだオブジェクト（相互参照の修復のときなど）は、暗号を解かずに読んでいるので、読み直す
        keep = self.encrypt_ref.num if self.encrypt_ref is not None else None
        self.cache = {k: v for k, v in self.cache.items() if k == keep}
        self.objstm_cache = {}
        return None

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
            xs = tr.get("XRefStm")
            if isinstance(xs, int) and xs != pos and xs not in self.__dict__.setdefault("_xrefstm_seen", set()):
                self._xrefstm_seen.add(xs)  # 自分や、読んだ xref ストリームを指す XRefStm は、もう読まない
                self._xref_section(xs)
            return int(tr["Prev"]) if "Prev" in tr else None
        # xref ストリーム
        obj = self._parse_at(pos)
        if not isinstance(obj, Stream) or getattr(obj.dict.get("Type"), "name", "") != "XRef":
            raise PdfError("xref ストリームではありません")
        d = obj.dict
        data = self.decode(obj, decrypt=False)
        w = [int(x) for x in d["W"]]
        if len(w) != 3 or min(w) < 0:
            raise PdfError("xref ストリームの欄の幅（W）が壊れています")
        size = int(d.get("Size", 0))
        index = [int(x) for x in d.get("Index", [0, size])]
        step = sum(w)
        if step <= 0:
            raise PdfError("xref ストリームの欄の幅（W）が壊れています")
        p = 0
        for a in range(0, len(index) - 1, 2):
            start = index[a]
            # 宣言された数（Index）は信じない：実際のデータに入っている行の数まで
            count = min(index[a + 1], (len(data) - p) // step)
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
        """相互参照が壊れているとき：ファイルを頭から読み、「n g obj」の位置を集め直す（同じ番号は後ろのもの＝追記で
        書き直したものを使う）。オブジェクトのストリーム（ObjStm）に入ったオブジェクトも、ストリームの中の番号の表から集める
        （ストリームより後ろに、同じ番号のオブジェクトがじかに書いてあれば、そちらを使う）。"""
        self.repaired = True
        self.xref = {}
        self.cache = {}
        self.objstm_cache = {}
        where: dict[int, int] = {}
        from ..lazy.filebytes import iter_matches

        for start, m in iter_matches(re.compile(rb"(?<![\d])(\d{1,10})\s+(\d{1,5})\s+obj\b"), self.blob):
            self.xref[int(m[1])] = ("o", start, int(m[2]))
            where[int(m[1])] = start
        for num, ent in sorted(self.xref.items(), key=lambda kv: kv[1][1]):
            if ent[0] != "o":
                continue
            try:
                st = self._parse_at(ent[1], num)
                if not (isinstance(st, Stream) and getattr(st.dict.get("Type"), "name", "") == "ObjStm"):
                    continue
                data = self.decode(st, decrypt=self.decryptor is not None)
                lx = Lexer(data, 0)
                for idx in range(int(st.dict.get("N", 0))):
                    k1, onum = lx.token()
                    k2, _off = lx.token()
                    if k1 != "num" or k2 != "num":
                        break
                    if onum != num and where.get(onum, -1) < ent[1]:
                        self.xref[onum] = ("s", num, idx)
                        where[onum] = ent[1]
            except Exception:  # noqa: BLE001
                continue
        for start, m in iter_matches(re.compile(rb"trailer\s*<<"), self.blob):
            try:
                tr = parse_object(Lexer(self.blob, start + m.end() - m.start() - 2))
                self.trailer.update(tr)
            except Exception:  # noqa: BLE001
                pass
        if "Root" not in self.trailer:
            for num in self.xref:
                o = self.get(num)
                if isinstance(o, dict) and getattr(o.get("Type"), "name", "") == "Catalog":
                    self.trailer["Root"] = Ref(num, 0)
                    break

    def verify_xref(self) -> None:
        """相互参照のすべての位置が、その番号のオブジェクトを指しているかを確かめ、1つでも違えば頭から読み直す
        （読むときは、読んだオブジェクトから順に確かめる。書き込み・証明では、はじめに全部を確かめる）。"""
        if self.repaired:
            return
        for num, ent in list(self.xref.items()):
            if ent[0] != "o":
                continue
            m = re.match(rb"\s*(\d+)\s+(\d+)\s+obj\b", self.blob[ent[1]:ent[1] + 40])
            if not m or int(m[1]) != num:
                self._repair()
                return

    # ------------------------------------------------------------------ オブジェクト
    def _parse_at(self, pos: int, expect: int | None = None):
        """pos のオブジェクトを読む。expect があれば、そこにあるのがその番号のオブジェクトか確かめる（違えば PdfError）。"""
        lx = Lexer(self.blob, pos)
        k1, num = lx.token()
        k2, gen = lx.token()
        k3, kw = lx.token()
        if (k1, k2, k3, kw) != ("num", "num", "kw", "obj"):
            raise PdfError(f"位置 {pos} にオブジェクトがありません")
        if expect is not None and num != expect:
            raise PdfError(f"位置 {pos} にあるのは、オブジェクト {expect} でなく {num} です")
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
                obj = self._parse_at(ent[1], num)
            except Exception:  # noqa: BLE001
                obj = None
                if not self.repaired:
                    # 相互参照の位置が、そのオブジェクトを指していない（番号のずれた xref 表・壊れた追記など）：表を信じず、
                    # ファイルを頭から読み直して位置を集め、読み直す
                    self._repair()
                    return self.get(num)
            if obj is not None and self.decryptor is not None and num != getattr(self.encrypt_ref, "num", -1):
                obj = self._decrypt_strings(obj, num, ent[2])
        else:
            objs = self._objstm(ent[1])
            obj = objs[ent[2]] if ent[2] < len(objs) else None
        if not (self.light and isinstance(obj, Stream) and len(obj.raw) > 1 << 16):
            self.cache[num] = obj
        return obj

    def _objstm(self, num: int) -> list:
        if num in self.objstm_cache:
            return self.objstm_cache[num]
        opening = self.__dict__.setdefault("_objstm_opening", set())
        if num in opening:
            # オブジェクトストリームが、自分自身（か、自分を読むのに要るストリーム）の中にあると相互参照が言う：読めない
            raise PdfError(f"オブジェクトストリーム {num} が、自分自身の中にあることになっています（相互参照が壊れています）")
        opening.add(num)
        try:
            return self._objstm_read(num)
        finally:
            opening.discard(num)

    def _objstm_read(self, num: int) -> list:
        st = self.get(num)
        out = []
        if isinstance(st, Stream):
            data = self.decode(st)
            n = int(st.dict.get("N", 0))
            first = int(st.dict.get("First", 0))
            lx = Lexer(data, 0)
            offs = []
            for _ in range(min(n, len(data) // 2)):  # 番号と位置の組は、どんなに短くても 2 バイトより長い
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
            out = PdfString(dec.decrypt_string(bytes(obj), num, gen))
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
            data = self.decryptor.decrypt_stream(data, st.num, st.gen, *self.crypt_of(st.dict, st.num))
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
            if name == "Crypt":
                continue  # 暗号フィルター（7.4.10）：暗号は上で解いた
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

    def crypt_of(self, d: dict, num: int) -> tuple[str | None, str]:
        """ストリームの暗号の決め方：(Filter の Crypt が指す暗号フィルターの名前, 種類（metadata・embedded・空）)。"""
        filters = self.resolve(d.get("Filter"))
        parms = self.resolve(d.get("DecodeParms"))
        if not isinstance(filters, list):
            filters, parms = [filters], [parms]
        elif not isinstance(parms, list):
            parms = [parms] * len(filters)
        crypt = None
        for i, f in enumerate(filters):
            if getattr(self.resolve(f), "name", "") == "Crypt":
                # 名前を省いたときは Identity（表14）
                prm = self.resolve(parms[i]) if i < len(parms) else None
                crypt = getattr(self.resolve(prm.get("Name")), "name", "Identity") if isinstance(prm, dict) else "Identity"
                break
        kind = ""
        t = getattr(self.resolve(d.get("Type")), "name", "")
        if t == "Metadata":
            # EncryptMetadata が決めるのは、文書のメタデータ（Catalog の Metadata）だけ（表21）
            md = self.root.get("Metadata")
            if isinstance(md, Ref) and md.num == num:
                kind = "metadata"
        elif t == "EmbeddedFile":
            kind = "embedded"
        return crypt, kind

    # ------------------------------------------------------------------ 文書の構造
    @property
    def root(self) -> dict:
        return self.resolve(self.trailer.get("Root")) or {}

    def pages(self) -> list[dict]:
        out: list[dict] = []
        seen: set[int] = set()  # たどったページの木の節（オブジェクトの番号）：輪や、同じ節を 2 度指す木を 1 度だけたどる

        def walk(node_ref, inherited, depth=0):
            if isinstance(node_ref, Ref):
                if node_ref.num in seen:
                    return
                seen.add(node_ref.num)
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
        if not out and not self.repaired and self.xref:
            # 入口（Root）からページの木をたどれない：相互参照が古い・壊れていることがある。頭から読み直して、もう一度たどる
            self._repair()
            seen.clear()
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
    """FlateDecode：出てくるバイトを数えながら展開する（上限：limits）。末尾が壊れている圧縮データは、読めるところまで読む
    （zlib の頭のないもの＝deflate だけのものも読む）。"""
    from ..limits import inflate

    out = inflate(data, 15, what="PDF のストリーム", tolerant=True)
    if not out:
        out = inflate(data, -15, what="PDF のストリーム", tolerant=True)
    return out


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
                if rowlen <= 0 or colors <= 0:
                    return data
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
    if rowlen <= 0 or rowlen >= len(data):
        return b"" if rowlen > 0 else data  # 宣言の 1 行より短いデータ：行が 1 つもない（宣言の大きさの場所を取らない）
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


class _Chunks:
    """展開したものを 4 MB ごとの塊で持つ（1 つの大きな塊を伸ばし続けると、写すたびに一時に何倍も使う）。"""

    def __init__(self):
        self.parts: list[bytes] = []
        self.cur = bytearray()
        self.n = 0

    def add(self, b: bytes) -> None:
        self.cur += b
        self.n += len(b)
        if len(self.cur) >= 1 << 22:
            self.parts.append(bytes(self.cur))
            self.cur.clear()

    def value(self) -> bytes:
        return b"".join(self.parts) + bytes(self.cur)


def _lzw(data: bytes) -> bytes:
    from ..limits import budget

    b = budget()
    cap = b.room()
    out = _Chunks()
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
                b.charge(out.n)
                return out.value()
            if prev is None:
                entry = table[code]
            elif code < len(table):
                entry = table[code]
                table.append(prev + entry[:1])
            else:
                entry = prev + prev[:1]
                table.append(entry)
            out.add(entry)
            prev = entry
            if out.n > cap:
                raise b.exceeded(cap, None, "PDF のストリーム（LZW）")
            if len(table) + 1 >= (1 << bits) and bits < 12:
                bits += 1
    b.charge(out.n)
    return out.value()


def _rle(data: bytes) -> bytes:
    from ..limits import budget

    b = budget()
    cap = b.room()
    out = _Chunks()
    i = 0
    while i < len(data):
        if out.n > cap:
            raise b.exceeded(cap, None, "PDF のストリーム（RunLength）")
        n = data[i]
        if n == 128:
            break
        if n < 128:
            out.add(data[i + 1:i + 2 + n])
            i += 2 + n
        else:
            out.add(data[i + 1:i + 2] * (257 - n))
            i += 2
    b.charge(out.n)
    return out.value()
