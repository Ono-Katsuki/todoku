"""Access のデータベース（.mdb の Jet3・Jet4、.accdb の ACE）の表を読む（読むだけ）。統計・コード表が配られる形。

Microsoft の公開の仕様はないので、ページの形は、公開されている解説（mdbtools の HACKING・Jackcess）で知られている形と、実物の構造の観察による：
- ページ 0：版（0x14。0 が Jet3・1 が Jet4・2 以降が ACE）。0x18 からは RC4（鍵 0x6B39DAC7）で隠され、0x3C に文字コード、0x3E に
  ページの隠しの鍵（0 でなければ、ページごとに 鍵^ページ番号 の RC4。実物の「暗号化」した .mdb で確かめた）
- 表の定義（型 2。続きは next_pg）：MSysObjects はページ 2。行の数・列の数・使う頁の表の位置・列の並びと名前
- データのページ（型 1）：行の位置の表（上 2 ビットが印：0x8000 削除・0x4000 ほかの行へのポインタ。下 13 ビットが位置）。
  行の末尾に NULL の印（1 列 1 ビット）と、長さの変わる列の位置の表
- メモ：12 バイトの見出し（長さの上 2 ビット：0x80000000 同じ所・0x40000000 別のページ 1 つ・0 別のページの連なり）
- Jet4 の字：UTF-16。先頭が FF FE なら圧縮（1 バイトの字と 2 バイトの字を 0x00 で切り替える）。Jet3 の字：ファイルの文字コード

パスワード：ページ 0 の 0x42 の欄（Jet3 は 20 バイトの字。Jet4・ACE は 40 バイトの UTF-16 で、0x72 の作成日時（倍精度）の整数部の 4 バイトと XOR してある：
Jackcess の知られた形。Microsoft の公開の仕様はない）。空でなければ password を求めて照らす。字として読めない欄は、確かめられないと窓に書く。
Access 2007 以降の .accdb の「パスワードで暗号化」は別の形（ページの暗号）で、ここでは解けない（表が読めないときに、そう伝える。実物 5 つで確かめた）。

読まないもの（窓の unread に書く）：リンクした表の先・クエリ（数だけ）・OLE とバイナリの列の中身・複合データ型の添付・計算列のうち字でないもの
（宣言の型と入っている値の形が一致しない実物があるので、値を推測しない）。

実物（公開の試験データ）で観察した形：
- 列の定義には番号が 2 つある。1 つは行の NULL の印のビットで、列を消しても変わらない。もう 1 つは表の中の並びで、消すと詰まる（システムの表では 0 のまま）。
- 行の位置の表は、大きい順とは限らない。行の終わりは、同じページで次に大きい位置の手前。
- 書き直して短くなった行には、データの終わりと位置の表の間に余りがある。
- Jet3 の繰り上がりの表の 0xFF は「繰り上がりなし」の詰め物。
- 数値（Decimal）の精度は scale の手前のバイト。
- 計算列（列の 2 つ目の印が 0xC0）の値は、16 バイトの 0・4 バイトの長さ・値・3 バイトの 0。長い字はメモの形の中に同じ枠で入る。
- 拡張日付（型 20）は ASCII の字で「日数:100 ナノ秒の数:桁数」。日数は 0001-01-01 からの数。

行の形が確かめられないときは、値を推測せず、その行を読まずに行の数を窓に書く（別の道具で書き出してもらう）。
"""

from __future__ import annotations

import bisect
import datetime as dt
import codecs
import re
import struct
from decimal import Decimal

from .tablebook import TableBook

MAX_CELLS = 300_000
TYPE_JA = {1: "真偽", 2: "バイト", 3: "整数", 4: "長整数", 5: "通貨", 6: "単精度", 7: "倍精度", 8: "日時", 9: "バイナリ", 10: "字", 11: "OLE", 12: "メモ", 15: "GUID", 16: "数値", 18: "複合", 20: "拡張日付"}
GUIDE = "Access で開いて表を .xlsx か CSV に書き出したものか、別の道具（mdbtools の mdb-export など）で CSV にしたものを開いてください"
_C1_GAPS = frozenset(b"\x81\x8d\x8f\x90\x9d")  # cp1252 に字のないバイト：同じ番号の制御文字にする（バイトを失わない）
_EXT_DATE = re.compile(rb"^(\d{19}):(\d{19}):(\d)\0$")


class JetError(ValueError):
    pass


def check_password(blob: bytes, password: str | None) -> str:
    """"" 開いてよい・"need" パスワードが要る・"wrong" 違う。"""
    try:
        db = Jet(blob)
    except (JetError, struct.error, IndexError):
        return ""  # 壊れたファイルは、窓の所で伝える
    if db.pw_status != "ok":
        return ""
    try:
        db.catalog()
    except (JetError, struct.error, IndexError):
        return ""  # 表の一覧が読めない（.accdb の暗号化。パスワードの欄の字も本当のパスワードではない）：窓の所で、解けないと伝える
    if password is None:
        return "need"
    return "" if password == db.pw_value else "wrong"


def is_jet(b: bytes) -> bool:
    return len(b) >= 0x20 and b[:4] == b"\x00\x01\x00\x00" and b[4:19] in (b"Standard Jet DB", b"Standard ACE DB")


def _rc4(key: bytes, data: bytes) -> bytes:
    s = list(range(256))
    j = 0
    for i in range(256):
        j = (j + s[i] + key[i % len(key)]) & 255
        s[i], s[j] = s[j], s[i]
    out = bytearray(len(data))
    i = j = 0
    for k, c in enumerate(data):
        i = (i + 1) & 255
        j = (j + s[i]) & 255
        s[i], s[j] = s[j], s[i]
        out[k] = c ^ s[(s[i] + s[j]) & 255]
    return bytes(out)


class _Fmt:
    def __init__(self, jet3: bool):
        self.jet3 = jet3
        if jet3:
            self.pg = 2048
            self.row_count = 8
            self.tab_rows, self.tab_cols, self.tab_idxs, self.tab_ridxs = 12, 25, 27, 31
            self.usage, self.cols_start, self.ridx_size = 35, 43, 8
            self.c_prec, self.c_scale, self.c_flags, self.c_size, self.c_num, self.c_entry = 9, 10, 13, 16, 1, 18
            self.c_var, self.c_fixed, self.c_rowcol = 3, 14, 5
        else:
            self.pg = 4096
            self.row_count = 12
            self.tab_rows, self.tab_cols, self.tab_idxs, self.tab_ridxs = 16, 45, 47, 51
            self.usage, self.cols_start, self.ridx_size = 55, 63, 12
            self.c_prec, self.c_scale, self.c_flags, self.c_size, self.c_num, self.c_entry = 11, 12, 15, 23, 5, 25
            self.c_var, self.c_fixed, self.c_rowcol = 7, 21, 9


class _Col:
    __slots__ = ("name", "type", "num", "pos", "var", "fixed", "flags", "size", "scale", "prec", "is_fixed", "calc")


def _ascending(offsets: list[int]) -> bool:
    return all(a <= b for a, b in zip(offsets, offsets[1:]))


def _jet4_var_table(p: bytes, start: int, count_at: int) -> tuple[int, list[int]]:
    """Jet4：数（2 バイト）のすぐ前に、2 バイトの位置が数＋1 個、後ろの列から順に並ぶ。(数, 列の順の位置)。
    位置は増える順で、最後の位置（データの終わり）は位置の表の手前に収まる（書き直して短くなった行には、間に余りがある）。"""
    nvar = struct.unpack_from("<H", p, count_at)[0]
    table_at = count_at - 2 * (nvar + 1)
    if table_at <= start:
        raise JetError("行の、長さの変わる列の位置の表が行に収まりません")
    offsets = list(struct.unpack_from(f"<{nvar + 1}H", p, table_at))[::-1]
    if not _ascending(offsets) or offsets[-1] > table_at - start:
        raise JetError("行の、長さの変わる列の位置の表が、行のデータと合いません")
    return nvar, offsets


def _jet3_var_table(p: bytes, start: int, count_at: int) -> tuple[int, list[int]]:
    """Jet3：位置は 1 バイト（下の桁）で、後ろの列から順に並ぶ。上の桁は、数のすぐ前の「繰り上がりの表」が決める：表の字は列の番号で、
    その列から後ろの位置に 256 を足す。0xFF は繰り上がりなしの詰め物。繰り上がりの数は書かれていないので、ありうる数をすべて試し、
    位置が増える順に並び、最後の位置（データの終わり）が位置の表の始まりに一致するものを取る。一致がなければ、表の手前に収まるもの
    （書き直して短くなった行）。どちらも 1 通りに決まらなければ推測しない。(数, 列の順の位置)。"""
    nvar = p[count_at]
    found: list[list[int]] = []
    fits: list[list[int]] = []
    for jumps in range((count_at - start) // 256 + 1, -1, -1):
        table_at = count_at - jumps - nvar - 1
        if table_at <= start:
            continue
        carries = sorted(x for x in p[count_at - jumps:count_at] if x != 0xFF)
        if carries and carries[-1] > nvar:
            continue
        lows = p[table_at:count_at - jumps][::-1]
        offsets = [low + 256 * bisect.bisect_right(carries, col) for col, low in enumerate(lows)]
        if not _ascending(offsets):
            continue
        if offsets[-1] == table_at - start and offsets not in found:
            found.append(offsets)
        elif offsets[-1] < table_at - start and offsets not in fits:
            fits.append(offsets)
    found = found or fits
    if len(found) != 1:
        raise JetError("行の、長さの変わる列の位置の表が、行のデータと合いません" if not found else "行の、長さの変わる列の位置が 1 通りに決まりません")
    return nvar, found[0]


class Jet:
    def __init__(self, blob: bytes):
        if not is_jet(blob):
            raise JetError("Access のデータベースの見出しではありません")
        self.b = blob
        ver = blob[0x14]
        self.ver = ver
        self.f = _Fmt(ver == 0)
        self.page_count = len(blob) // self.f.pg
        head = bytearray(blob[:self.f.pg])
        size = 126 if self.f.jet3 else 128
        head[0x18:0x18 + size] = _rc4(struct.pack("<I", 0x6B39DAC7), bytes(head[0x18:0x18 + size]))
        self.code_page = struct.unpack_from("<H", head, 0x3C)[0]
        key = struct.unpack_from("<I", head, 0x3E)[0]
        self.key = key if key != 0xE15E01B9 else 0
        self.pw_status, self.pw_value = self._stored_password(head)
        self._cache: dict[int, bytes] = {}
        self._starts: tuple[bytes | None, list[int]] = (None, [])
        self.bad: dict[int, list[str]] = {}  # 表の定義のページ → 読めなかった行の理由
        self.enc, self.enc_known = _encoding(self.code_page)

    def _stored_password(self, head: bytearray) -> tuple[str, str]:
        """("none" 空・"ok" 読めた・"unknown" 確かめられない, 字)。"""
        if self.f.jet3:
            raw = bytes(head[0x42:0x42 + 20])
            if not any(raw):
                return "none", ""
            try:
                return "ok", raw.split(b"\0", 1)[0].decode({932: "cp932"}.get(struct.unpack_from("<H", head, 0x3C)[0], "cp1252"))
            except UnicodeDecodeError:
                return "unknown", ""
        raw = bytearray(head[0x42:0x42 + 40])
        try:
            mask = (int(struct.unpack_from("<d", head, 0x72)[0]) & 0xFFFFFFFF).to_bytes(4, "little")
        except (OverflowError, ValueError):
            return "unknown", ""
        for i in range(len(raw)):
            raw[i] ^= mask[i % 4]
        if not any(raw):
            return "none", ""
        try:
            s = bytes(raw).decode("utf-16-le").split("\0", 1)[0]
        except UnicodeDecodeError:
            return "unknown", ""
        if not s or any(ord(c) < 0x20 or 0xE000 <= ord(c) < 0xF900 for c in s):
            return "unknown", ""  # 字に見えない：パスワードの欄でないか、形が違う（推測しない）
        return "ok", s

    def page(self, n: int) -> bytes:
        if n in self._cache:
            return self._cache[n]
        o = n * self.f.pg
        if n < 0 or o + self.f.pg > len(self.b):
            raise JetError(f"ページ {n} がファイルの外です（途中で切れています）")
        p = self.b[o:o + self.f.pg]
        if n and self.key:
            p = _rc4(struct.pack("<I", self.key ^ n), p)
        if len(self._cache) > 512:
            self._cache.clear()
        self._cache[n] = p
        return p

    # ---------------------------------------------------------- 表の定義
    def tdef(self, n: int) -> bytes:
        p = self.page(n)
        if p[0] != 2:
            raise JetError(f"ページ {n} は表の定義ではありません")
        out = bytearray(p)
        nxt = struct.unpack_from("<I", p, 4)[0]
        seen = {n}
        while nxt:
            if nxt in seen:
                raise JetError("表の定義の続きが輪になっています")
            seen.add(nxt)
            q = self.page(nxt)
            if q[0] != 2:
                raise JetError("表の定義の続きが壊れています")
            out += q[8:]
            nxt = struct.unpack_from("<I", q, 4)[0]
        return bytes(out)

    def table_def(self, n: int) -> tuple[list[_Col], bytes]:
        d = self.tdef(n)
        f = self.f
        ncols = struct.unpack_from("<H", d, f.tab_cols)[0]
        nridx = struct.unpack_from("<I", d, f.tab_ridxs)[0]
        pos = f.cols_start + nridx * f.ridx_size
        cols = []
        for _ in range(ncols):
            e = d[pos:pos + f.c_entry]
            if len(e) < f.c_entry:
                raise JetError("列の定義が途中で終わっています")
            c = _Col()
            c.type = e[0]
            c.num = struct.unpack_from("<H", e, f.c_num)[0]  # 行の NULL の印のビット（列を消しても変わらない）
            c.pos = struct.unpack_from("<H", e, f.c_rowcol)[0]  # 表の中の並び
            c.var = struct.unpack_from("<H", e, f.c_var)[0]
            c.fixed = struct.unpack_from("<H", e, f.c_fixed)[0]
            c.flags = e[f.c_flags]
            c.size = struct.unpack_from("<H", e, f.c_size)[0]
            c.scale, c.prec = e[f.c_scale], e[f.c_prec]
            c.is_fixed = bool(c.flags & 1)
            c.calc = not f.jet3 and e[f.c_flags + 1] == 0xC0
            cols.append(c)
            pos += f.c_entry
        for c in cols:
            if f.jet3:
                ln = d[pos]
                c.name = d[pos + 1:pos + 1 + ln].decode(self.enc, "replace")
                pos += 1 + ln
            else:
                ln = struct.unpack_from("<H", d, pos)[0]
                c.name = self._ucs(d[pos + 2:pos + 2 + ln])
                pos += 2 + ln
        return cols, d

    @staticmethod
    def _ucs(raw: bytes) -> str:
        return raw.decode("utf-16-le", "replace")

    def text(self, raw: bytes) -> str:
        if self.f.jet3:
            if self.enc == "cp1252" and _C1_GAPS.intersection(raw):
                return "".join(chr(b) if b in _C1_GAPS else bytes((b,)).decode("cp1252") for b in raw)
            return raw.decode(self.enc, "replace")
        if raw[:2] == b"\xff\xfe":
            out = bytearray()
            comp = True
            i = 2
            n = len(raw)
            while i < n:
                c = raw[i]
                if c == 0:
                    comp = not comp
                    i += 1
                elif comp:
                    out += bytes((c, 0))
                    i += 1
                elif i + 1 < n:
                    out += raw[i:i + 2]
                    i += 2
                else:
                    break
            return bytes(out).decode("utf-16-le", "replace")
        return raw.decode("utf-16-le", "replace")

    # ---------------------------------------------------------- 使うページ
    def _row_range(self, p: bytes, i: int) -> tuple[int, int, int]:
        """(始まり, 終わり, 印)。行の終わりは、同じページで次に大きい位置の手前（位置の表は大きい順とは限らない）。"""
        rc = self.f.row_count
        off = struct.unpack_from("<H", p, rc + 2 + 2 * i)[0]
        if self._starts[0] is not p:
            n = struct.unpack_from("<H", p, rc)[0]
            self._starts = (p, sorted(x & 0x1FFF for x in struct.unpack_from(f"<{n}H", p, rc + 2)))
        starts = self._starts[1]
        start = off & 0x1FFF
        at = bisect.bisect_right(starts, start)
        end = (starts[at] if at < len(starts) else self.f.pg) - 1
        return start, end, off & 0xC000

    def _row_at(self, pg: int, i: int) -> bytes:
        p = self.page(pg)
        start, end, _f = self._row_range(p, i)
        return p[start:end + 1]

    def usage_pages(self, d: bytes) -> list[int]:
        f = self.f
        dp = struct.unpack_from("<I", d, f.usage)[0]
        m = self._row_at(dp >> 8, dp & 0xFF)
        if not m:
            return []
        out = []
        if m[0] == 0:
            startpg = struct.unpack_from("<I", m, 1)[0]
            for k, byte in enumerate(m[5:]):
                for bit in range(8):
                    if byte >> bit & 1:
                        out.append(startpg + k * 8 + bit)
        elif m[0] == 1:
            per = (f.pg - 4) * 8
            for j in range((len(m) - 1) // 4):
                mp = struct.unpack_from("<I", m, 1 + 4 * j)[0]
                if not mp:
                    continue
                bits = self.page(mp)
                if bits[0] != 5:
                    raise JetError("使うページの表が壊れています")
                for k, byte in enumerate(bits[4:]):
                    if byte:
                        for bit in range(8):
                            if byte >> bit & 1:
                                out.append(j * per + k * 8 + bit)
        else:
            raise JetError(f"使うページの表の形（{m[0]}）が読めません")
        return out

    def rows(self, tdef_pg: int, cols: list[_Col], d: bytes):
        """表の行を順に。形を確かめられない行は値を推測せず、self.bad[表の定義のページ] に理由を入れて飛ばす。"""
        f = self.f
        bad = self.bad[tdef_pg] = []
        for pg in self.usage_pages(d):
            if pg == 0:
                continue
            if pg >= self.page_count:
                bad.append(f"ページ {pg} がファイルの外です")
                continue
            p = self.page(pg)
            if p[0] != 1 or struct.unpack_from("<I", p, 4)[0] != tdef_pg:
                continue
            n = struct.unpack_from("<H", p, f.row_count)[0]
            for i in range(n):
                try:
                    q, start, end, flags = self._follow(tdef_pg, p, i)
                    if flags & 0x8000:
                        continue
                    row = self._crack(q, start, end, cols)
                except (JetError, struct.error, IndexError, ValueError) as e:
                    bad.append(f"ページ {pg} の行 {i}：{e}")
                    continue
                yield row

    def _follow(self, tdef_pg: int, p: bytes, i: int) -> tuple[bytes, int, int, int]:
        """行の位置。ほかの行へのポインタ（行番号 1 バイト・ページ 3 バイト）をたどる（行き先には消えた行と同じ印があるので、たどった先では見ない）。"""
        start, end, flags = self._row_range(p, i)
        if flags & 0x8000 or not flags & 0x4000:
            return p, start, end, flags
        seen = set()
        while True:
            dp = struct.unpack_from("<I", p, start)[0]
            if dp in seen or len(seen) > 16:
                raise JetError("ほかの行へのポインタが輪になっています")
            seen.add(dp)
            p = self.page(dp >> 8)
            if p[0] != 1 or struct.unpack_from("<I", p, 4)[0] != tdef_pg:
                raise JetError("ほかの行へのポインタの先が、この表のページではありません")
            if dp & 0xFF >= struct.unpack_from("<H", p, self.f.row_count)[0]:
                raise JetError("ほかの行へのポインタの先の行がありません")
            start, end, flags = self._row_range(p, dp & 0xFF)
            if not flags & 0x4000:
                return p, start, end, 0

    def _crack(self, p: bytes, start: int, end: int, cols: list[_Col]) -> dict:
        """1 つの行を列の値に。行の末尾から：NULL の印・長さの変わる列の数・（Jet3 は位置の繰り上がり）・位置の表。"""
        jet3 = self.f.jet3
        if end >= len(p) or start >= end:
            raise JetError("行の位置がページに収まりません")
        row_cols, col_start = (p[start], start + 1) if jet3 else (struct.unpack_from("<H", p, start)[0], start + 2)
        mask_sz = (row_cols + 7) // 8
        if row_cols == 0 or col_start + mask_sz > end + 1:
            raise JetError("行の列の数が行の長さと合いません")
        mask = p[end - mask_sz + 1:end + 1]
        nvar, var = 0, []
        data_end = end - mask_sz + 1
        if not all(c.is_fixed for c in cols):
            nvar, var = _jet3_var_table(p, start, end - mask_sz) if jet3 else _jet4_var_table(p, start, end - mask_sz - 1)
            data_end = start + var[-1]
        out: dict = {}
        for c in cols:
            nn = c.num // 8 < len(mask) and bool(mask[c.num // 8] >> (c.num % 8) & 1)
            if c.type == 1:
                out[c.name] = nn
                continue
            if not nn:
                continue
            if c.is_fixed:
                if c.num >= row_cols:
                    continue
                at = col_start + c.fixed
                if at + c.size > data_end:
                    raise JetError(f"列「{c.name}」が行のデータに収まりません")
                raw = p[at:at + c.size]
            else:
                if c.var >= nvar:
                    continue
                raw = p[start + var[c.var]:start + var[c.var + 1]]
            out[c.name] = self._calc_value(c, raw) if c.calc else self.value(c, raw)
        return out

    def _calc_value(self, c: _Col, raw: bytes):
        """計算列。字だけを読む（字でない列は、宣言の型と違う形の値が入っている実物があるので読まない）。"""
        if c.type not in (10, 12):
            return f"〔計算列（{TYPE_JA.get(c.type, f'型{c.type}')}）：結果の形を確かめられないので読みません〕"
        body = raw
        if c.type == 12 or c.size == 0:  # 長い字はメモの形（同じ行の中か、別の行）
            body = self._memo(raw)
        if len(body) < 23 or any(body[:16]):
            raise JetError(f"計算列「{c.name}」の値の枠が読めません")
        n = struct.unpack_from("<I", body, 16)[0]
        if len(body) != 23 + n:
            raise JetError(f"計算列「{c.name}」の値の長さが合いません")
        return self.text(body[20:20 + n])

    def _memo(self, raw: bytes) -> bytes:
        if len(raw) < 12:
            return b""
        head, dp = struct.unpack_from("<II", raw, 0)
        ln = head & 0x3FFFFFFF
        if head & 0x80000000:
            if len(raw) < 12 + ln:
                raise JetError("メモの長さが行と合いません")
            return raw[12:12 + ln]
        if head & 0x40000000:
            out = self._row_at(dp >> 8, dp & 0xFF)[:ln]
            if len(out) != ln:
                raise JetError("メモの長さが、入っている行と合いません")
            return out
        out = bytearray()
        seen = set()
        while dp and len(out) < ln:
            if dp in seen:
                raise JetError("メモの続きが輪になっています")
            seen.add(dp)
            r = self._row_at(dp >> 8, dp & 0xFF)
            dp = struct.unpack_from("<I", r, 0)[0]
            out += r[4:]
        if len(out) < ln:
            raise JetError("メモの続きが途中で切れています")
        return bytes(out[:ln])

    def value(self, c: _Col, raw: bytes):
        t = c.type
        if t == 2:
            return float(raw[0])
        if t == 3:
            return float(struct.unpack("<h", raw[:2])[0])
        if t == 4:
            return float(struct.unpack("<i", raw[:4])[0])
        if t == 5:
            v = struct.unpack("<q", raw[:8])[0]
            return _exact(abs(v), 4, v < 0)
        if t == 6:
            return _single(raw[:4])
        if t == 7:
            return struct.unpack("<d", raw[:8])[0]
        if t == 8:
            return _date(struct.unpack("<d", raw[:8])[0])
        if t == 10:
            return self.text(raw)
        if t == 12:
            return self.text(self._memo(raw))
        if t == 15:
            a, b_, c_ = struct.unpack("<IHH", raw[:8])
            return f"{{{a:08X}-{b_:04X}-{c_:04X}-{raw[8:10].hex().upper()}-{raw[10:16].hex().upper()}}}"
        if t == 16 and len(raw) >= 17:
            # 1 バイトの符号のあと、4 バイトずつの組（上の組から。組の中は下のバイトから）
            v = 0
            for k in range(4):
                v = (v << 32) | struct.unpack_from("<I", raw, 1 + 4 * k)[0]
            return _exact(v, c.scale, bool(raw[0] & 0x80))
        if t == 20:
            return _ext_date(raw)
        if t in (9, 11) and not raw:
            return ""
        if t == 11 and len(raw) >= 12:  # OLE の値は 12 バイトの見出しで、見出しの長さが中身の長さ
            n = struct.unpack_from("<I", raw, 0)[0] & 0x3FFFFFFF
            return f"〔OLE：{n} バイト。中身は読みません〕"
        return f"〔{TYPE_JA.get(t, f'型{t}')}：{len(raw)} バイト。中身は読みません〕"

    # ---------------------------------------------------------- 表の一覧
    def catalog(self) -> tuple[list[tuple[str, int]], list[str], int]:
        """(利用者の表 [(名前, 定義のページ)], リンクした表の名前, クエリの数)。"""
        cols, d = self.table_def(2)
        tables, linked = [], []
        queries = 0
        names = {c.name for c in cols}
        if not {"Name", "Type", "Id", "Flags"} <= names:
            raise JetError("表の一覧（MSysObjects）の列が読めません")
        for r in self.rows(2, cols, d):
            t = int(r.get("Type") or 0) & 0x7F
            fl = int(r.get("Flags") or 0) & 0xFFFFFFFF
            nm = r.get("Name") or ""
            if t == 1 and not fl & 0x80000002:
                tables.append((nm, int(r.get("Id") or 0) & 0x00FFFFFF))
            elif t in (4, 6):
                linked.append(nm)
            elif t == 5 and not nm.startswith("~"):
                queries += 1
        if self.bad.get(2):
            raise JetError(f"表の一覧（MSysObjects）の行を読めません（{self.bad[2][0]}）")
        return tables, linked, queries


def _encoding(code_page: int) -> tuple[str, bool]:
    """(Python の文字コードの名前, 確かな名前か)。0 は Access の既定の 1252。"""
    if code_page == 0:
        return "cp1252", True
    try:
        codecs.lookup(f"cp{code_page}")
    except LookupError:
        return "cp1252", False
    return f"cp{code_page}", True


def _exact(v: int, scale: int, neg: bool):
    """10 進の値を、倍精度で同じ値に戻るなら数で、戻らないなら字で（桁を失わない）。"""
    digits = str(v).rjust(scale + 1, "0")
    text = (digits[:-scale] + "." + digits[-scale:]) if scale else digits
    if neg:
        text = "-" + text
    x = float(text)
    return x if Decimal(repr(x)) == Decimal(text) else text


def _single(raw: bytes) -> float:
    """単精度の値を、同じ単精度に戻る一番短い 10 進に（0.1 を 0.10000000149 と見せない）。"""
    x = struct.unpack("<f", raw)[0]
    for digits in range(6, 10):
        y = float(f"{x:.{digits}g}")
        try:
            if struct.pack("<f", y) == raw:
                return y
        except OverflowError:
            break
    return x


def _ext_date(raw: bytes) -> str:
    """拡張日付：ASCII の「日数:100 ナノ秒の数:桁数」＋ 0。日数は 0001-01-01 から。"""
    m = _EXT_DATE.match(raw)
    if not m:
        raise JetError("拡張日付の字の形が読めません")
    day = dt.date.fromordinal(int(m[1]) + 1)
    ticks = int(m[2])
    if ticks >= 864_000_000_000:
        raise JetError("拡張日付の時刻が 1 日を超えています")
    sec, frac = divmod(ticks, 10_000_000)
    out = f"{day.year:04d}-{day.month:02d}-{day.day:02d}"
    if ticks:
        out += f" {sec // 3600:02d}:{sec // 60 % 60:02d}:{sec % 60:02d}"
        if frac:
            out += f".{frac:07d}".rstrip("0")
    return out


def _date(d: float) -> str:
    """日時：1899-12-30 からの日数。負の値も、整数の部分が日で、小数の部分の大きさが時刻（-1.5 は 1899-12-29 の 12:00）。"""
    try:
        whole = int(d)
        secs = round(abs(d - whole) * 86400)
        t = dt.datetime(1899, 12, 30) + dt.timedelta(days=whole, seconds=secs)
    except (OverflowError, ValueError):
        return str(d)
    day = f"{t.year:04d}-{t.month:02d}-{t.day:02d}"
    return f"{day} {t.hour:02d}:{t.minute:02d}:{t.second:02d}" if secs else day


class JetWorkbook(TableBook):
    def __init__(self, blob: bytes, filename: str = ""):
        super().__init__(blob, filename, "mdb")
        try:
            db = Jet(blob)
            from ..archives import current_password

            state = check_password(blob, current_password())
            if state:
                raise JetError("このデータベースはパスワードで保護されています。本人にパスワードを聞いて password に入れて開いてください"
                               if state == "need" else "パスワードが違います")
            try:
                tables, linked, queries = db.catalog()
            except (JetError, struct.error, IndexError) as e:
                if db.ver >= 2:
                    raise JetError(f"表の一覧を読めません（{e}）。Access 2007 以降の「パスワードで暗号化」した .accdb なら、とどくでは解けません。"
                                   "Access でパスワードを解除して保存し直したものか、表を .xlsx・CSV に書き出したものを開いてください") from None
                raise JetError(f"表の一覧を読めません（{e}）。{GUIDE}") from None
            self.db = db
            total = 0
            self.table_rows: dict[str, int] = {}
            todo = []
            for name, pg in sorted(tables, key=lambda x: x[0].lower()):
                self.table_rows[name] = 0
                todo.append((_stated_cells(db, pg), name, pg, self.add_sheet(name[:64] or f"表{pg}")))
            for _size, name, pg, sh in sorted(todo, key=lambda x: x[0]):  # 小さな表から読む（上限で、名前の順の後ろの小さな表が丸ごと消えないように）
                try:
                    total = self._read_table(db, sh, name, pg, total)
                except (JetError, struct.error, IndexError, ValueError) as e:
                    sh.unread.append(f"表「{name}」を読めません（{e}）。値を推測しません。{GUIDE}")
            if not self.sheets:
                sh = self.add_sheet("（表なし）")
                sh.unread.append("利用者の表がありません")
            if not db.enc_known:
                self.sheets[0].unread.append(f"データベースの文字コード（{db.code_page}）を知らないので、Jet3 の字は 1252 として読みました。字が正しいとは言えません")
            if db.pw_status == "unknown":
                self.sheets[0].unread.append("データベースのパスワードの欄を字として読めないので、パスワードが要るかどうかを確かめられません（表はそのまま読みました）")
            if linked or queries:
                self.sheets[0].unread.append("" + (f"リンクした表 {len(linked)} 個（{'・'.join(linked[:8])}）の先のデータはこのファイルにないので読めません。" if linked else "")
                                             + (f"クエリ {queries} 個は読んでいません（表だけを読みました）。" if queries else ""))
        except (struct.error, IndexError) as e:
            raise JetError(f"データベースのページが壊れています（{e}）") from None

    def _read_table(self, db: Jet, sh, name: str, pg: int, total: int) -> int:
        cols, d = db.table_def(pg)
        order = _display_order(cols)
        for c, col in enumerate(order, 1):
            self.put(sh, 1, c, col.name)
        nrows = 0
        skipped = 0
        for row in db.rows(pg, cols, d):
            if total >= MAX_CELLS:
                skipped += 1
                continue
            nrows += 1
            for c, col in enumerate(order, 1):
                v = row.get(col.name)
                if v is not None and v != "":
                    self.put(sh, nrows + 1, c, v)
                    total += 1
        self.table_rows[name] = nrows
        if skipped:
            sh.unread.append(f"表「{name}」：セルの数が読む上限（{MAX_CELLS:,}）を超えたので、{skipped:,} 行は読んでいません。{GUIDE}")
        bad = db.bad.get(pg) or []
        if bad:
            sh.unread.append(f"表「{name}」：{len(bad):,} 行は形を確かめられないので読んでいません（値を推測しません。例：{bad[0]}）。{GUIDE}")
        stated = struct.unpack_from("<I", d, db.f.tab_rows)[0]
        if not skipped and not bad and stated != nrows:
            sh.unread.append(f"表「{name}」：表の定義の行の数は {stated:,} ですが、読めた行は {nrows:,} です。行の数を確かめられません。{GUIDE}")
        bin_cols = [c.name for c in cols if c.type in (9, 11, 18)]
        if bin_cols:
            sh.unread.append(f"表「{name}」：バイナリ・OLE・添付の列（{'・'.join(bin_cols)}）の中身は読んでいません")
        calc_cols = [c.name for c in cols if c.calc and c.type not in (10, 12)]
        if calc_cols:
            sh.unread.append(f"表「{name}」：計算列（{'・'.join(calc_cols)}）の結果は読んでいません（結果の形を確かめられないので）。Access で開いて確かめてください")
        return total


def _stated_cells(db: Jet, pg: int) -> int:
    """表の定義に書かれた大きさ（行の数 × 列の数）。読む順を決めるだけに使う。読めない定義は最後に回す（読むときに理由を伝える）。"""
    try:
        d = db.tdef(pg)
        return struct.unpack_from("<I", d, db.f.tab_rows)[0] * struct.unpack_from("<H", d, db.f.tab_cols)[0]
    except (JetError, struct.error, IndexError):
        return 1 << 62


def _display_order(cols: list[_Col]) -> list[_Col]:
    """表の中の並び（システムの表では並びの番号が 0 のままなので、次に列の番号で）。"""
    return sorted(cols, key=lambda c: (c.pos, c.num))


def raw_values(wb) -> dict:
    """窓とは別に：同じページを、もう一度頭から歩いて行を数える（列ごとの NULL の印と、窓のセルを照らす）。"""
    db = wb.db
    out: dict = {}
    tables, _l, _q = db.catalog()
    for name, pg in tables:
        cols, d = db.table_def(pg)
        order = _display_order(cols)
        s = set()
        for c, col in enumerate(order, 1):
            s.add((1, c))
        r = 1
        for row in db.rows(pg, cols, d):
            r += 1
            if r - 1 > wb.table_rows.get(name, 0):
                break
            for c, col in enumerate(order, 1):
                v = row.get(col.name)
                if v is not None and v != "":
                    s.add((r, c))
        out[name[:64] or f"表{pg}"] = s
    return out
