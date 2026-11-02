"""Excel 5.0/95（BIFF5）の数式（逆ポーランド記法のトークン列）を、BIFF8 と同じ数式の文字列に戻す。

トークンの形は OpenOffice.org の「Microsoft Excel File Format」3 章（BIFF2〜BIFF8 の違いを並べたもの）で確かめた。
BIFF8 と違うところ：
- セルの番地は 3 バイト：行（下 14 ビット）と、行の 15・16 ビット目に「列が相対」「行が相対」の印、列（1 バイト）。
  範囲は 6 バイト（行・行・列・列）。共有数式（tRefN・tAreaN）では、相対の行は 14 ビット、列は 8 ビットの符号つきのずれ
- tName は 14 バイト（番号 2 バイト・未使用 12 バイト）。tNameX は 24 バイト
- 別のシートへの参照（tRef3d 17 バイト・tArea3d 20 バイト）：最初の 2 バイトが負なら同じブックで、その絶対値は
  そのシート（名前ならブック全体）の EXTERNSHEET の番号（1 始まり）。11・13 バイト目に、最初と最後のシートの番号（0 始まり）
- 文字列（tStr）は 8 ビットの長さのバイト列で、文字コードは CODEPAGE のもの
- 定数の配列（数式の後ろの追加データ）：列の数（0 は 256）・行の数（どちらも 1 を引かない）・値（文字列は 8 ビットの長さのバイト列）
読み違えた式で計算すると保存値と合わない値を作るので、確かめられない形（外部のブック・複数シート・名前の中の相対参照・
シートの番号と EXTERNSHEET の名前が食い違うもの・追加データの長さが合わないもの）は推測せずに Unsupported にする。
"""

from __future__ import annotations

import struct

from ..formula.parser import num_to_col
from ..formula.values import Unsupported
from .ptg import _BINOP, _PREC, ERRORS, FTAB, _q

# BIFF8 で新しく足された関数（Excel 5.0/95 のファイルには現れないはずなので、出てきたら読み違いを疑って断る）
_BIFF8_ONLY = set(range(358, 368))


class Decoder5:
    """links はこの数式の EXTERNSHEET の並び（セルならそのシートの、名前ならブック全体の）。

    sheet_index は数式のあるシートの番号（0 始まり。名前では None）。"""

    def __init__(self, book, enc: str, links: list[bytes | None], sheet_index: int | None, is_name: bool = False):
        self.book = book
        self.enc = enc
        self.links = links
        self.sheet_index = sheet_index
        self.is_name = is_name

    # ------------------------------------------------------------------ 番地
    @staticmethod
    def _cell(rw: int, col: int, base: tuple[int, int] | None) -> str:
        rel_r = bool(rw & 0x8000)
        rel_c = bool(rw & 0x4000)
        r = rw & 0x3FFF
        c = col
        if base is not None:
            if rel_r:
                r = base[0] + (r - 0x4000 if r >= 0x2000 else r)
            if rel_c:
                c = base[1] + (c - 0x100 if c >= 0x80 else c)
            if not (0 <= r < 0x4000 and 0 <= c < 0x100):
                raise Unsupported("共有数式の相対参照が、シートの外を指しています")
        return f"{'' if rel_c else '$'}{num_to_col(c + 1)}{'' if rel_r else '$'}{r + 1}"

    def _area(self, r1: int, r2: int, c1: int, c2: int, base) -> str:
        if base is None and (r1 & 0x3FFF) == 0 and (r2 & 0x3FFF) == 0x3FFF:
            ca = ("" if r1 & 0x4000 else "$") + num_to_col(c1 + 1)
            cb = ("" if r2 & 0x4000 else "$") + num_to_col(c2 + 1)
            return f"{ca}:{cb}"
        if base is None and c1 == 0 and c2 == 0xFF:
            ra = ("" if r1 & 0x8000 else "$") + str((r1 & 0x3FFF) + 1)
            rb = ("" if r2 & 0x8000 else "$") + str((r2 & 0x3FFF) + 1)
            return f"{ra}:{rb}"
        return f"{self._cell(r1, c1, base)}:{self._cell(r2, c2, base)}"

    def _sheet3d(self, ixals: int, first: int, last: int) -> str | None:
        """別のシートへの参照の「シート名!」。消えたシートは None（#REF!）。"""
        if ixals >= 0:
            raise Unsupported("別のブックへの参照（Excel 5.0/95）")
        k = -ixals
        if not (1 <= k <= len(self.links)) or self.links[k - 1] is None:
            raise Unsupported("別のシートへの参照の EXTERNSHEET が見つかりません")
        link = self.links[k - 1]
        if first == 0xFFFF or last == 0xFFFF:
            return None
        if first != last:
            raise Unsupported("複数シートにまたがる参照")
        sheets = self.book.sheets
        if first >= len(sheets):
            raise Unsupported("別のシートへの参照のシートの番号が、シートの数を超えています")
        kind = link[:1]
        if kind == b"\x03":
            name = link[1:].decode(self.enc)
            if name != sheets[first].name:
                raise Unsupported(f"別のシートへの参照で、シートの番号（{first}）と EXTERNSHEET の名前（{name}）が食い違います")
        elif kind == b"\x02":
            if self.sheet_index is None or first != self.sheet_index:
                raise Unsupported("「このシート」を指す EXTERNSHEET と、シートの番号が食い違います")
        elif kind == b"\x00":
            raise Unsupported("名前の中の、使う所のシートを指す参照（=!A1）")
        else:
            raise Unsupported("意味を確かめていない EXTERNSHEET（同じブックの、シートを決めない参照）")
        return _q(sheets[first].name) + "!"

    # ------------------------------------------------------------------ 本体
    def decode(self, rgce: bytes, extra: bytes = b"", base: tuple[int, int] | None = None, rel: bool = False) -> str:
        """base は共有数式の親のセル（行, 列。0 始まり）。tRefN・tAreaN の相対のずれの基準。"""
        del rel
        st: list[tuple[str, int]] = []
        need: list[str] = []  # 追加データを使うトークンの順
        i = 0
        n = len(rgce)

        def take(k: int) -> bytes:
            nonlocal i
            if i + k > n:
                raise Unsupported("数式のトークンが途中で終わっています")
            out = rgce[i:i + k]
            i += k
            return out

        def pop() -> tuple[str, int]:
            if not st:
                raise Unsupported("数式のトークンが壊れています")
            return st.pop()

        while i < n:
            ptg = rgce[i]
            i += 1
            bid = ptg if ptg < 0x20 else 0x20 | (ptg & 0x1F)
            if ptg in _BINOP:
                b, pb = pop()
                a, pa = pop()
                op = _BINOP[ptg]
                prec = _PREC[op]
                if pa < prec:
                    a = f"({a})"
                if pb <= prec and op != ":":
                    b = f"({b})"
                st.append((f"{a}{op}{b}" if op != "," else f"({a},{b})", prec))
            elif ptg == 0x12:
                a, _ = pop()
                st.append(("+" + a, _PREC["neg"]))
            elif ptg == 0x13:
                a, pa = pop()
                st.append(("-" + (a if pa >= _PREC["neg"] else f"({a})"), _PREC["neg"]))
            elif ptg == 0x14:
                a, pa = pop()
                st.append(((a if pa >= _PREC["%"] else f"({a})") + "%", _PREC["%"]))
            elif ptg == 0x15:
                a, _ = pop()
                st.append((f"({a})", 99))
            elif ptg == 0x16:
                st.append(("", 99))
            elif ptg == 0x17:  # tStr：8 ビットの長さのバイト列（CODEPAGE の文字コード）
                cch = take(1)[0]
                raw = take(cch)
                try:
                    s = raw.decode(self.enc)
                except UnicodeDecodeError:
                    raise Unsupported(f"数式の中の文字列が、文字コード {self.enc} で読めません") from None
                st.append(('"' + s.replace('"', '""') + '"', 99))
            elif ptg == 0x19:  # tAttr
                grbit = take(1)[0]
                if grbit & 0x04:  # tAttrChoose：選択肢の数（2 バイト）と飛び先の表
                    nc = struct.unpack("<H", take(2))[0]
                    take(2 * (nc + 1))
                    continue
                take(2)
                if grbit & 0x10:
                    a, _ = pop()
                    st.append((f"SUM({a})", 99))
                elif grbit & ~0x4B & 0xFF:  # 0x01 揮発・0x02 IF・0x08 飛び・0x40 空白 のほか（0x20 代入など）
                    raise Unsupported(f"数式の特別なトークン（tAttr 0x{grbit:02X}）")
            elif ptg == 0x1C:
                st.append((ERRORS.get(take(1)[0], "#N/A"), 99))
            elif ptg == 0x1D:
                st.append(("TRUE" if take(1)[0] else "FALSE", 99))
            elif ptg == 0x1E:
                st.append((str(struct.unpack("<H", take(2))[0]), 99))
            elif ptg == 0x1F:
                v = struct.unpack("<d", take(8))[0]
                st.append((repr(v) if v != int(v) or abs(v) >= 1e15 else str(int(v)), 99))
            elif bid == 0x20:  # tArray（7 バイトは未使用。値は後ろの追加データ）
                take(7)
                need.append("array")
                st.append((f"\x00ARRAY{len(need) - 1}\x00", 99))
            elif bid == 0x21:
                iftab = struct.unpack("<H", take(2))[0]
                name, argc = FTAB.get(iftab, (None, None))
                if name is None or argc is None or iftab in _BIFF8_ONLY:
                    raise Unsupported(f"関数の番号 {iftab}")
                if len(st) < argc:
                    raise Unsupported("数式のトークンが壊れています")
                args = [pop()[0] for _ in range(argc)][::-1]
                st.append((f"{name}({','.join(args)})", 99))
            elif bid == 0x22:
                argc = take(1)[0] & 0x7F
                iftab = struct.unpack("<H", take(2))[0]
                if iftab & 0x8000:
                    raise Unsupported("マクロの命令")
                if len(st) < argc:
                    raise Unsupported("数式のトークンが壊れています")
                args = [pop()[0] for _ in range(argc)][::-1]
                fname = FTAB.get(iftab, (None, None))[0]
                if iftab == 255 or fname is None or iftab in _BIFF8_ONLY:
                    raise Unsupported(f"関数の番号 {iftab}（アドイン・ユーザー定義の関数を含む）")
                st.append((f"{fname}({','.join(args)})", 99))
            elif bid == 0x23:  # tName：番号（1 始まり）・未使用 12 バイト
                idx = struct.unpack("<H", take(2))[0]
                take(12)
                if not (1 <= idx <= len(self.book._name_raw)):
                    raise Unsupported(f"名前の番号 {idx} がありません")
                st.append((self.book.name_text(idx), 99))
            elif bid in (0x24, 0x2C):  # tRef・tRefN
                if self.is_name:
                    raise Unsupported("名前の中の、シートを決めない参照")
                rw, col = struct.unpack("<HB", take(3))
                if bid == 0x2C and base is None:
                    raise Unsupported("共有数式ではない所の相対参照（tRefN）")
                st.append((self._cell(rw, col, base if bid == 0x2C else None), 99))
            elif bid in (0x25, 0x2D):  # tArea・tAreaN
                if self.is_name:
                    raise Unsupported("名前の中の、シートを決めない参照")
                r1, r2, c1, c2 = struct.unpack("<HHBB", take(6))
                if bid == 0x2D and base is None:
                    raise Unsupported("共有数式ではない所の相対参照（tAreaN）")
                st.append((self._area(r1, r2, c1, c2, base if bid == 0x2D else None), 99))
            elif bid in (0x26, 0x27, 0x28):  # tMemArea・tMemErr・tMemNoMem：未使用 4 バイトと、続く式の長さ
                take(6)
                if bid == 0x26:
                    need.append("mem")
            elif bid in (0x29, 0x2E, 0x2F):
                take(2)
            elif bid == 0x2A:
                take(3)
                st.append(("#REF!", 99))
            elif bid == 0x2B:
                take(6)
                st.append(("#REF!", 99))
            elif bid == 0x3A:  # tRef3d
                ixals = struct.unpack("<h", take(2))[0]
                take(8)
                first, last, rw, col = struct.unpack("<HHHB", take(7))
                if self.is_name and rw & 0xC000:
                    raise Unsupported("名前の中の相対参照")
                pre = self._sheet3d(ixals, first, last)
                st.append(((pre + self._cell(rw, col, None)) if pre else "#REF!", 99))
            elif bid == 0x3B:  # tArea3d
                ixals = struct.unpack("<h", take(2))[0]
                take(8)
                first, last, r1, r2, c1, c2 = struct.unpack("<HHHHBB", take(10))
                if self.is_name and (r1 | r2) & 0xC000:
                    raise Unsupported("名前の中の相対参照")
                pre = self._sheet3d(ixals, first, last)
                st.append(((pre + self._area(r1, r2, c1, c2, None)) if pre else "#REF!", 99))
            elif bid == 0x3C:  # tRefErr3d（17 バイト）
                take(17)
                st.append(("#REF!", 99))
            elif bid == 0x3D:  # tAreaErr3d（20 バイト）
                take(20)
                st.append(("#REF!", 99))
            elif bid == 0x39:
                raise Unsupported("ほかのシートの名前・外部の名前・アドインの関数（tNameX）")
            elif ptg in (0x01, 0x02):
                raise Unsupported("共有数式・配列数式・データテーブルの本体を指すトークン")
            else:
                raise Unsupported(f"数式のトークン 0x{ptg:02X}")
        if len(st) != 1:
            raise Unsupported("数式のトークンが壊れています")
        text = st[0][0]
        arrays = self._extra(need, extra)
        for k, a in enumerate(arrays):
            text = text.replace(f"\x00ARRAY{k}\x00", a)
        if "\x00" in text:
            raise Unsupported("定数の配列の値が見つかりません")
        return text

    def _extra(self, need: list[str], extra: bytes) -> list[str]:
        """追加データを、トークンの順に読む。tArray は定数の配列、tMemArea は範囲の並び（使わずに飛ばす）。
        最後まで読んで、長さがちょうど合わなければ断る（形を読み違えている）。"""
        pos = 0
        out = []

        def get(k: int) -> bytes:
            nonlocal pos
            if pos + k > len(extra):
                raise Unsupported("数式の追加データが途中で終わっています")
            b = extra[pos:pos + k]
            pos += k
            return b

        for kind in need:
            if kind == "mem":
                cnt = struct.unpack("<H", get(2))[0]
                get(6 * cnt)
                continue
            nc = get(1)[0] or 256
            nr = struct.unpack("<H", get(2))[0]
            rows = []
            for _r in range(nr):
                vals = []
                for _c in range(nc):
                    t = get(1)[0]
                    if t == 0x01:
                        v = struct.unpack("<d", get(8))[0]
                        vals.append(str(int(v)) if v == int(v) and abs(v) < 1e15 else repr(v))
                    elif t == 0x02:
                        cch = get(1)[0]
                        try:
                            s = get(cch).decode(self.enc)
                        except UnicodeDecodeError:
                            raise Unsupported(f"定数の配列の文字列が、文字コード {self.enc} で読めません") from None
                        vals.append('"' + s.replace('"', '""') + '"')
                    elif t == 0x04:
                        vals.append("TRUE" if get(8)[0] else "FALSE")
                    elif t == 0x10:
                        vals.append(ERRORS.get(get(8)[0], "#N/A"))
                    else:
                        raise Unsupported(f"定数の配列の値の種類 0x{t:02X}")
                rows.append(",".join(vals))
            out.append("{" + ";".join(rows) + "}")
        if pos != len(extra):
            raise Unsupported("数式の追加データの長さが合いません")
        return out
