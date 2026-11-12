"""dBase（.dbf。古い統計・地図の属性表で配られる）を、表のブックとして読む。同じ大きさの上書きだけ書く。

- 見出し：[0] 版、[4:8] 件数、[8:10] 見出しの大きさ、[10:12] 1 件の大きさ、[29] 言語の印。32 バイト目から 32 バイトずつの項目（0x0D で終わる）
- 項目の型：C 字・N/F 数（字で書かれた数）・D 日付（YYYYMMDD）・L 真偽・I 整数・B 倍精度・Y 通貨・T 日時（FoxPro）。
- メモ（M）：同じ語幹の .fpt（FoxPro：見出しの 6 バイト目に大端の区画の大きさ、区画は型 4・長さ 4 の大端と中身）か .dbt（dBase III：512 バイトの区画で 0x1A まで。
  dBase IV：見出しの 20 バイト目に区画の大きさ、区画は FF FF 08 00・長さ 4）。ファイルの隣か、書庫の同じフォルダから読む。G・P（絵・OLE）の中身は読まない
- 書く：1 件の大きさも見出しも変えず、項目の幅の中で上書きする（字は空白で埋める。数は右に寄せる）。メモと絵の項目には書かない。変わったのが書いた項目のバイトだけであることを照らす
- 削除の印（先頭が「*」）のある行は読まず、数を伝える
- 文字コード：ファイルの字の全部を、1 つの文字コードで読む（namecode.decide_whole）。隣の .cpg（シェープファイルの決まり：
  文字コードの名前かコードページの番号）、見出しの言語の印（LDID）の順で、字の全部を厳密に読めた最初の宣言。宣言がないか、
  宣言では読めなければ、規則では決まらない：読める文字コードを決まった順に試し、最初の読み方で示して、候補を窓に並べる
  （エージェントが字を読んで選び、encoding で渡す）。項目の幅いっぱいの字は、終わりの字が幅で切れていることがあるので、
  終わりの途中までの字は読まずに、その数を伝える
"""

from __future__ import annotations

import datetime as dt
import struct

from .tablebook import TableBook

# 言語の印（見出しの 29 バイト目。LDID）→ 文字コード。dBASE・FoxPro・ESRI の表（dbfread と GDAL の表を合わせたもの）。
# 0x57（ANSI：書いた機械の文字コード）は、どの文字コードかをファイルが言わないので、宣言にしない。
# 0x68（Kamenický）・0x69（Mazovia）は、Python に表がないので、宣言にしない
LDID = {0x01: "cp437", 0x02: "cp850", 0x03: "cp1252", 0x04: "mac_roman", 0x08: "cp865", 0x09: "cp437", 0x0A: "cp850",
        0x0B: "cp437", 0x0D: "cp437", 0x0E: "cp850", 0x0F: "cp437", 0x10: "cp850", 0x11: "cp437", 0x12: "cp850",
        0x13: "cp932", 0x14: "cp850", 0x15: "cp437", 0x16: "cp850", 0x17: "cp865", 0x18: "cp437", 0x19: "cp437",
        0x1A: "cp850", 0x1B: "cp437", 0x1C: "cp863", 0x1D: "cp850", 0x1F: "cp852", 0x22: "cp852", 0x23: "cp852",
        0x24: "cp860", 0x25: "cp850", 0x26: "cp866", 0x37: "cp850", 0x40: "cp852", 0x4D: "gbk", 0x4E: "cp949",
        0x4F: "cp950", 0x50: "cp874", 0x58: "cp1252", 0x59: "cp1252", 0x64: "cp852", 0x65: "cp866", 0x66: "cp865",
        0x67: "cp861", 0x6A: "cp737", 0x6B: "cp857", 0x6C: "cp863", 0x78: "cp950", 0x79: "cp949", 0x7A: "gbk",
        0x7B: "cp932", 0x7C: "cp874", 0x7D: "cp1255", 0x7E: "cp1256", 0x86: "cp737", 0x87: "cp852", 0x88: "cp857",
        0x96: "mac_cyrillic", 0x97: "mac_latin2", 0x98: "mac_greek", 0xC8: "cp1250", 0xC9: "cp1251", 0xCA: "cp1254",
        0xCB: "cp1253", 0xCC: "cp1257"}


def cpg_encoding(data: bytes) -> str | None:
    """.cpg の中身（シェープファイルの決まり）→ 文字コード。コードページの番号（932・1252）、「8859」に続く番号（88591・8859-2）、
    「ANSI 1251」「OEM 866」の形、文字コードの名前（UTF-8・SJIS・Big5）。「LDID/19」は言語の印の番号。読めなければ None。"""
    from ..namecode import canonical, codec_of_codepage

    try:
        t = data[:64].decode("ascii").strip().strip("\0").strip()
    except UnicodeDecodeError:
        return None
    low = t.lower()
    if low.startswith("ldid/"):
        return LDID.get(int(low[5:])) if low[5:].isdigit() else None
    for head in ("ansi ", "oem "):
        if low.startswith(head):
            low = low[len(head):].strip()
    if low.startswith("8859"):
        num = low[4:].lstrip("-")
        return canonical(f"iso8859_{num}") if num.isdigit() else None
    if low.isdigit():
        return codec_of_codepage(int(low))
    return canonical(low)


VERSIONS = {0x02, 0x03, 0x04, 0x05, 0x30, 0x31, 0x32, 0x43, 0x63, 0x83, 0x8B, 0x8E, 0xCB, 0xF5, 0xFB}


class DbfError(ValueError):
    pass


def parse_header(b: bytes):
    """(件数, 見出しの大きさ, 1 件の大きさ, 項目 [(名前, 型, 長さ, 小数)])。形が合わなければ DbfError。"""
    if len(b) < 33 or b[0] not in VERSIONS:
        raise DbfError("dBase の見出しではありません")
    n, hs, rs = struct.unpack_from("<IHH", b, 4)
    if hs < 33 or rs < 1 or hs > len(b):
        raise DbfError("dBase の見出しの大きさが合いません")
    fields = []
    p = 32
    off = 1
    while True:
        if p >= hs:
            raise DbfError("項目の終わりの印（0x0D）がありません")
        if b[p] == 0x0D:
            break
        if p + 32 > hs:
            raise DbfError("項目が見出しの外に出ています")
        raw = b[p:p + 11].split(b"\0", 1)[0]
        t = chr(b[p + 11])
        ln, dec = b[p + 16], b[p + 17]
        if t == "C" and b[0] in (0x30, 0x31, 0x43):  # FoxPro の長い字の項目は、長さの 2 バイト
            ln |= dec << 8
            dec = 0
        fields.append((raw, t, ln, dec))
        off += ln
        p += 32
    if off != rs:
        raise DbfError(f"項目の長さの合計（{off}）が、1 件の大きさ（{rs}）と合いません")
    if n * rs + hs > len(b) + 1:
        raise DbfError("件数のぶんのデータがありません（途中で切れています）")
    return n, hs, rs, fields


def is_dbf(b: bytes) -> bool:
    try:
        parse_header(b)
        return True
    except (DbfError, struct.error):
        return False


class _Memo:
    def __init__(self, data: bytes, name: str, version: int):
        self.d = data
        self.name = name
        low = name.lower()
        self.fpt = low.endswith(".fpt")
        if self.fpt:
            self.block = struct.unpack_from(">H", data, 6)[0] if len(data) >= 8 else 0
        elif version in (0x8B, 0x8E, 0xCB, 0x7B) and len(data) >= 22:
            self.block = struct.unpack_from("<H", data, 20)[0] or 512
        else:
            self.block = 512
        if self.block <= 0:
            raise DbfError("メモのファイルの区画の大きさが壊れています")

    def read(self, n: int) -> bytes | None:
        """n 番の区画の中身（読めなければ None）。"""
        p = n * self.block
        d = self.d
        if p < 0 or p >= len(d):
            return None
        if self.fpt:
            if p + 8 > len(d):
                return None
            _t, ln = struct.unpack_from(">II", d, p)
            return d[p + 8:p + 8 + ln] if p + 8 + ln <= len(d) else None
        if d[p:p + 4] == b"\xff\xff\x08\x00":
            ln = struct.unpack_from("<I", d, p + 4)[0]
            return d[p + 8:p + ln] if 8 <= ln and p + ln <= len(d) else None
        e = d.find(b"\x1a", p)
        return d[p:e if e >= 0 else len(d)]


_CUT = object()  # 幅いっぱいで、厳密には読めない字の項目


class DbfWorkbook(TableBook):
    def __init__(self, blob: bytes, filename: str = ""):
        super().__init__(blob, filename, "dbf")
        try:
            n, hs, rs, fields = parse_header(blob)
        except struct.error as e:
            raise DbfError(f"見出しが壊れています（{e}）") from None
        store = self._memo_store(blob[0], filename) if any(t == "M" for _r, t, _l, _d in fields) else None
        self.text_reading = self._reading(blob, n, hs, rs, fields, filename, store)
        self.encoding = self.text_reading.encoding
        self.cut: list[tuple[int, int]] = []  # 項目の幅で、終わりの字が切れている値の (行, 列)
        name = filename.rsplit("/", 1)[-1].rsplit(".", 1)[0] or "dbf"
        sh = self.add_sheet(name)
        for c, (raw, _t, _ln, _d) in enumerate(fields, 1):
            self.put(sh, 1, c, raw.decode(self.encoding, "replace"))
        memo = [raw.decode(self.encoding, "replace") for raw, t, _l, _d in fields if t in "MGP"]
        self.fields_raw = fields
        self.hs, self.rs = hs, rs
        self.rowmap: dict[int, int] = {}
        bad_memo = 0
        row = 1
        deleted = 0
        self.values = 0
        for i in range(n):
            p = hs + i * rs
            mark = blob[p]
            if mark == 0x2A:
                deleted += 1
                continue
            if mark == 0x1A:
                break
            row += 1
            self.rowmap[row] = i
            q = p + 1
            for c, (_raw, t, ln, dc) in enumerate(fields, 1):
                v, st = self._value(blob[q:q + ln], t, dc)
                if v is _CUT:
                    v, was_cut = self._cut_text(blob[q:q + ln])
                    if was_cut:
                        self.cut.append((row, c))
                if t == "M" and store is not None:
                    v = self._memo_value(store, blob[q:q + ln])
                    if v is False:
                        v = None
                        bad_memo += 1
                q += ln
                if v is not None:
                    self.put(sh, row, c, v, st)
                    self.values += 1
        pictures = [raw.decode(self.encoding, "replace") for raw, t, _l, _d in fields if t in "GP"]
        if pictures:
            sh.unread.append(f"絵・OLE の項目（{'・'.join(pictures)}）の中身は読んでいません")
        if any(t == "M" for _r, t, _l, _d in fields):
            names = "・".join(raw.decode(self.encoding, "replace") for raw, t, _l, _d in fields if t == "M")
            if store is None:
                sh.unread.append(f"メモの項目（{names}）の中身は別のファイル（.dbt・.fpt）にあります。同じ名前のファイルが隣（書庫なら同じフォルダ）にないので読んでいません")
            elif bad_memo:
                sh.unread.append(f"メモのファイル「{store.name}」の {bad_memo} か所を読めません（壊れているか、別の版のファイルです）")
        if deleted:
            sh.unread.append(f"削除の印のある {deleted} 行は読んでいません")
        if self.cut:
            from ..formula.parser import addr

            sh.unread.append(f"項目の幅で、終わりの字が途中で切れている値が {len(self.cut)} 個あります（切れた字のバイトは字にならないので、"
                             f"窓に出していません：{'・'.join(addr(r, c) for r, c in self.cut[:8])}{' など' if len(self.cut) > 8 else ''}）")

    def _memo_store(self, version: int, filename: str) -> _Memo | None:
        from ..archives import SIBLING

        reader = SIBLING.get()
        if reader is None or not filename:
            return None
        stem = filename.rsplit("/", 1)[-1].rsplit(".", 1)[0]
        try:
            hit = reader(stem, (".fpt", ".dbt"))
            return _Memo(hit[0], hit[1], version) if hit else None
        except (DbfError, ValueError, struct.error, OSError):
            return None

    def _memo_value(self, store: _Memo, raw: bytes):
        """None は空、False は読めない。番号は字の 10 桁か、4 バイトの整数（Visual FoxPro）。"""
        if len(raw) == 4:
            n = struct.unpack("<I", raw)[0]
        else:
            s = raw.strip(b" \0")
            if not s:
                return None
            if not s.isdigit():
                return False
            n = int(s)
        if n == 0:
            return None
        data = store.read(n)
        if data is None:
            return False
        text = data.decode(self.encoding, "replace").rstrip("\x00")
        return text or None

    def _reading(self, b: bytes, n: int, hs: int, rs: int, fields, filename: str, store):
        """字（項目の名前・字の項目・メモ）の全部の、文字コードの読み方。削除の印のある行は読まないので、入れない。"""
        from ..archives import SIBLING
        from ..namecode import table_reading

        whole: dict[bytes, None] = {}
        cut: dict[bytes, None] = {}
        for raw, _t, _l, _d in fields:
            if not raw.isascii():
                (cut if len(raw) >= 10 else whole)[raw] = None
        cols, off = [], 1
        for _raw, t, ln, _d in fields:
            if t in "CM":
                cols.append((off, ln, t))
            off += ln
        if cols and (store is not None or not b[hs:hs + n * rs].isascii()):
            flags = b[hs:hs + n * rs:rs]  # 行の頭の 1 バイト（削除の印・終わりの印）
            end = flags.find(b"\x1a")
            rows = range(hs, hs + (len(flags) if end < 0 else end) * rs, rs)
            live = None if 0x2A not in flags[:len(rows)] else [p for p in rows if b[p] != 0x2A]
            for o, ln, t in cols:
                if t == "M" and store is None:
                    continue
                # 項目ごとに、違う値だけを見る（出てきた順）
                for raw in dict.fromkeys(b[p + o:p + o + ln] for p in (rows if live is None else live)):
                    if t == "M":
                        data = self._memo_bytes(store, raw)
                        if data and not data.isascii():
                            whole[data.rstrip(b"\0")] = None
                    elif not raw.isascii():
                        v = raw.rstrip(b" \0")
                        (cut if len(v) == ln else whole)[v] = None
        declared = []
        reader = SIBLING.get()
        if reader is not None and filename:
            try:
                hit = reader(filename.rsplit("/", 1)[-1].rsplit(".", 1)[0], (".cpg",))
            except (ValueError, OSError):
                hit = None
            if hit:
                declared.append((cpg_encoding(hit[0]), f"隣のファイル {hit[1]}"))
        if b[29] in LDID:
            declared.append((LDID[b[29]], f"言語の印 0x{b[29]:02X}"))
        return table_reading(b, whole, cut, declared, empty_ok=True)  # 書ける形式：ASCII だけのファイルに書く字の文字コードも選べる

    def _memo_bytes(self, store: _Memo, raw: bytes) -> bytes | None:
        if len(raw) == 4:
            n = struct.unpack("<I", raw)[0]
        else:
            s = raw.strip(b" \0")
            if not s.isdigit():
                return None
            n = int(s)
        return store.read(n) if n else None

    def _cut_text(self, raw: bytes) -> tuple[str | None, bool]:
        """幅いっぱいで厳密には読めない値：(字, 終わりの字が幅で切れていたか)。切れていれば、切れた字の前まで。"""
        import codecs

        try:
            s = codecs.getincrementaldecoder(self.encoding)("strict").decode(raw, False)
        except (UnicodeDecodeError, LookupError):
            return raw.decode(self.encoding, "replace").rstrip(" \0") or None, False
        return s.rstrip(" \0") or None, True

    def _value(self, raw: bytes, t: str, dc: int):
        if t == "C":
            try:
                s = raw.decode(self.encoding).rstrip(" \0")
            except UnicodeDecodeError:
                if raw[-1:] not in (b" ", b"\0"):
                    return _CUT, 0  # 幅いっぱいの字：終わりの字が切れているかを、呼んだ側で確かめる
                s = raw.decode(self.encoding, "replace").rstrip(" \0")
            return (s or None), 0
        if t in "NF":
            s = raw.strip(b" \0*").decode("ascii", "replace")
            if not s or s == ".":
                return None, 0
            try:
                return float(s), 0
            except ValueError:
                return s, 0
        if t == "D":
            s = raw.strip(b" \0").decode("ascii", "replace")
            if len(s) == 8 and s.isdigit():
                try:
                    return f"{dt.date(int(s[:4]), int(s[4:6]), int(s[6:8])).isoformat()}", 0
                except ValueError:
                    return s, 0
            return (s or None), 0
        if t == "L":
            ch = raw[:1].upper()
            return (True if ch in (b"T", b"Y") else False if ch in (b"F", b"N") else None), 0
        if t == "I" and len(raw) == 4:
            return float(struct.unpack("<i", raw)[0]), 0
        if t in "BO" and len(raw) == 8:
            return struct.unpack("<d", raw)[0], 0
        if t == "Y" and len(raw) == 8:
            return struct.unpack("<q", raw)[0] / 10000, 0
        if t == "T" and len(raw) == 8:
            jd, ms = struct.unpack("<ii", raw)
            if jd == 0:
                return None, 0
            d = dt.date.fromordinal(jd - 1721425) if jd > 1721425 else None
            sec = ms // 1000
            return f"{d.isoformat()} {sec // 3600:02d}:{sec // 60 % 60:02d}:{sec % 60:02d}" if d else None, 0
        if t in "MGP":
            return None, 0
        return raw.decode(self.encoding, "replace").rstrip(" \0") or None, 0


def raw_values(wb) -> dict:
    """窓とは別に：データの部分を項目の長さで切り直し、空でないマスを数える。"""
    b = wb.zip.original
    n, hs, rs, fields = parse_header(b)
    out = set()
    row = 1
    for c in range(1, len(fields) + 1):
        out.add((1, c))
    for i in range(n):
        p = hs + i * rs
        if b[p] == 0x2A:
            continue
        if b[p] == 0x1A:
            break
        row += 1
        q = p + 1
        for c, (_r, t, ln, _d) in enumerate(fields, 1):
            cell = b[q:q + ln]
            q += ln
            if t in "MGP" or t in "ILBOYT":
                if t in "IBOY" or (t == "L" and cell[:1].upper() in (b"T", b"Y", b"F", b"N")) or (t == "T" and cell != bytes(ln)):
                    out.add((row, c))
                continue
            if cell.strip(b" \0") not in (b"", b".", b"*"):
                out.add((row, c))
    return {wb.sheets[0].name: out}


# ====================================================================== 書く：同じ大きさの上書き
NO_WRITE = "メモ・絵の項目と、形の決まらない項目には書きません。CSV か .xlsx に書き出してもらえば書けます"


def encode_value(v, t: str, ln: int, dc: int, enc: str, where: str) -> bytes:
    """値 → 項目の幅ちょうどのバイト。入らなければ DbfError（切り詰めない）。書けない字・読み直すと違う字になる字も断る。"""
    if t == "C":
        from ..datafile.textenc import EncodingError, exact

        try:
            b = exact(str(v), enc)
        except EncodingError as e:
            raise DbfError(f"{where}：{e}") from None
        if len(b) > ln:
            raise DbfError(f"{where}：字が項目の幅（{ln} バイト）に入りません（{len(b)} バイト）。切り詰めません")
        return b.ljust(ln, b" ")
    if t in "NF":
        try:
            x = float(str(v).replace(",", "")) if not isinstance(v, bool) else None
        except ValueError:
            x = None
        if x is None or x != x or x in (float("inf"), float("-inf")):
            raise DbfError(f"{where}：数の項目です。{v!r} は書けません")
        s = f"{x:.{dc}f}"
        if float(s) != x:
            raise DbfError(f"{where}：小数 {dc} 桁の項目では {v!r} をちょうど表せません")
        if len(s) > ln:
            raise DbfError(f"{where}：数が項目の幅（{ln} 桁）に入りません")
        return s.rjust(ln).encode("ascii")
    if t == "D":
        s = str(v).strip().replace("-", "").replace("/", "")
        try:
            dt.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
        except ValueError:
            raise DbfError(f"{where}：日付の項目です。年月日（2026-11-03 の形）で書いてください") from None
        if len(s) != 8 or ln != 8:
            raise DbfError(f"{where}：日付の項目です。年月日（2026-11-03 の形）で書いてください")
        return s.encode("ascii")
    if t == "L":
        if isinstance(v, bool):
            return (b"T" if v else b"F").ljust(ln, b" ")
        raise DbfError(f"{where}：真偽の項目です。true か false を書いてください")
    if t == "I" and ln == 4 and isinstance(v, (int, float)) and not isinstance(v, bool) and float(v) == int(v) and -(1 << 31) <= v < 1 << 31:
        return struct.pack("<i", int(v))
    if t in "BO" and ln == 8 and isinstance(v, (int, float)) and not isinstance(v, bool):
        return struct.pack("<d", float(v))
    raise DbfError(f"{where}：この項目（型 {t}）には {v!r} を書けません。{NO_WRITE}")


def api_fill(blob: bytes, name: str, values: dict) -> dict:
    from ..formula.parser import addr
    from ..xlsx.edit import resolve_target

    wb = DbfWorkbook(blob, name)
    out = bytearray(blob)
    spans = []
    written = {}
    for key, v in values.items():
        try:
            sheet, r, c = resolve_target(wb, key)
        except Exception as e:  # noqa: BLE001
            raise DbfError(f"{key}：セルが分かりません（{e}）") from None
        where = f"{sheet}!{addr(r, c)}"
        if r == 1:
            raise DbfError(f"{where}：1 行目は項目の名前です。書きません")
        if r not in wb.rowmap or not 1 <= c <= len(wb.fields_raw):
            raise DbfError(f"{where}：表の外です（dBase の件数と項目は増やしません）")
        _raw, t, ln, dc = wb.fields_raw[c - 1]
        if t in "MGP":
            raise DbfError(f"{where}：{NO_WRITE}")
        pos = wb.hs + wb.rowmap[r] * wb.rs + 1 + sum(f[2] for f in wb.fields_raw[:c - 1])
        if t == "C" and wb.text_reading.undetermined and not str(v).isascii():
            raise DbfError(f"{where}：このファイルの字の文字コードは、規則では決まっていません（窓に候補が並んでいます）。"
                           "ASCII でない字を書く前に、字を読んで文字コードを選び、encoding で渡してください")
        if t == "C" and wb.encoding == "ascii" and not str(v).isascii():
            raise DbfError(f"{where}：このファイルの字は ASCII だけで、文字コードの宣言（言語の印・隣の .cpg）もないので、"
                           "ASCII でない字をどの文字コードで書くかが決まりません。このファイルを読むソフトの文字コードを本人に確かめて、"
                           "encoding で渡してください")
        out[pos:pos + ln] = encode_value(v, t, ln, dc, wb.encoding, where)
        spans.append((pos, pos + ln))
        written[(sheet, r, c)] = v
    new = bytes(out)
    problems = []
    if len(new) != len(blob):
        problems.append("ファイルの大きさが変わっています")
    spans.sort()
    i = 0
    for k in range(min(len(new), len(blob))):
        if new[k] != blob[k]:
            while i < len(spans) and spans[i][1] <= k:
                i += 1
            if i >= len(spans) or k < spans[i][0]:
                problems.append(f"書いた項目のほかのバイト（位置 {k}）が変わっています")
                break
    if wb.text_reading.basis == "given":
        from ..namecode import use_table_encoding

        use_table_encoding(wb.encoding, new)  # 書いた後のファイルも、選んだ文字コードで読む
    w2 = DbfWorkbook(new, name)
    for sh, sh2 in zip(wb.sheets, w2.sheets):
        for rc, cell in sh.cells.items():
            if (sh.name, *rc) not in written and sh2.cells.get(rc) is not None and sh2.cells[rc].value != cell.value:
                problems.append(f"{sh.name} の {rc} の値が変わっています")
    rb = {}
    for (s, r, c), _v in written.items():
        cell = w2.sheet(s).cells.get((r, c))
        rb[f"{s}!{addr(r, c)}"] = cell.value if cell is not None else None
    head = f"{len(written)} か所を項目の幅の中で上書きしました。ファイルの大きさは同じで、変わったのは書いた項目のバイトだけです" if not problems else "証明で、想定外の変更が見つかりました"
    return {"blob": new, "readback": rb, "labels": {},
            "result": {"proof_ok": not problems, "proof": head, "warnings": [], "unexpected": problems[:20], "structure_errors": [],
                       "recalculated": {}, "written_as": {k: str(x) for k, x in rb.items()}}}


def read_cells(blob: bytes, name: str, keys) -> dict:
    from ..xlsx.edit import resolve_target

    wb = DbfWorkbook(blob, name)
    out = {}
    for key in keys:
        try:
            s, r, c = resolve_target(wb, key)
        except Exception:  # noqa: BLE001
            continue
        cell = wb.sheet(s).cells.get((r, c))
        out[key] = cell.value if cell is not None else None
    return out
