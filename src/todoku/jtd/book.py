"""一太郎（.jtd・.jtt）の文書を読む。外部のライブラリは使わない。

一太郎の形式には公開された仕様がないので、集めた文書（法務局・文部科学省などの様式）そのものから確かめた構造だけを使う。
確かめたことは、すべての文書で「記録の両端の長さと種類が合う」「組の開きと閉じが対応する」ことを検査しながら読み、
合わなければ推測せずに読むのをやめる（JtdError）。

- 入れ物は複合ファイル（CFB）。本文は DocumentText、レイアウト枠（テキストボックス）の文字は LayoutBoxText、
  ヘッダー・フッター（柱）は Header、ページスタイルの柱は PageLayoutStyleHeader、脚注は Footnote のストリームにある
- これらは「Ssmg」の入れ物：先頭の 20 バイト（"SsmgV.01"・セグメントの数・ブロックの大きさ・ブロックの数。数はすべて
  ビッグエンディアン）・ブロックの並び・セグメントの目録（番号・0・バイト数・ブロックの数×2・ブロックの番号の並び）。
  セグメントは、目録のブロックを順につないで、バイト数で切ったもの
- 文字のセグメント：「TextV.01」・単位の数（4 バイト）・16 ビットの単位（ビッグエンディアンの UTF-16）。後ろに文字の書式が続く
- 単位：0x20 以上は字。0x0A 段落の終わり・0x0C 改ページ・0x0E 表の行の終わり・0x09 タブ・0x02 枠の置き場所の印。
  記録：0x1C 種類 長さ … 長さ 0 種類 0x1F（長さは両端を含む単位の数。後ろからもたどれる）。
  組：0x1C 種類 長さ … 0x1D〈中身〉0x1E 長さ 0 種類 0x1F（字の中の部品：均等割り・ルビ・リンク・フィールドなど）
- 記録の種類 0x30 は表のマスの始まり（2・3 番目の単位がマスの左右の位置）。0x00 は字の中の部品の始まりで、2 番目の単位が
  部品の種類（0x30 は枠の置き場所で、6 番目の単位が枠の番号＝LayoutBoxText のセグメントの番号。0x05 ルビ・0x48 リンク・
  0x35 フィールド）。0x10 は段落の書式、0x20 は表の範囲（どちらも字を持たない）
- 組の 3 つの単位 (0, A, B)：A=0 は表示する字。A=1・2 はルビ・リンク先・フィールドの命令などの添え書き
- 「圧縮して保存」した文書は JSCompDocument のストリームに、LHA（-lh5-）で縮めた元の文書が入っている（lha.py で解く）
読んだ結果は Word（.docx）と同じ形（段落・表・マス）にして、窓を共通にする。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

from ..cfb import CFB
from ..docx.document import Cell, Para, Seg, Table

SSMG = b"SsmgV.01"
TEXT = b"TextV.01"
# 文字のセグメントを持つストリームと、窓での名前
STREAMS = {"DocumentText": "body", "LayoutBoxText": "レイアウト枠", "Header": "header",
           "PageLayoutStyleHeader": "ページスタイルの柱", "Footnote": "footnotes"}
T_FRAME, T_RUBY, T_LINK, T_FIELD = 0x30, 0x05, 0x48, 0x35  # 字の中の部品の種類（記録の種類 0x00 の 2 番目の単位）


class JtdError(ValueError):
    pass


def is_jtd(names: set[str]) -> bool:
    """複合ファイルのストリームの名前（小文字）から、一太郎の文書かを見分ける。"""
    js = any(n.startswith("\x04jsrv_") for n in names)
    return js and ("documenttext" in names or "jscompdocument" in names)


@dataclass
class Segment:
    stream: str
    sid: int
    data: bytes  # ブロックをつないで、バイト数で切ったもの
    blocks: tuple[int, ...]
    bsize: int

    def offset(self, k: int) -> int:
        """セグメントの中の位置 → ストリームの中の位置。"""
        b, w = divmod(k, self.bsize)
        return 20 + self.blocks[b] * self.bsize + w


def ssmg(data: bytes, stream: str) -> list[Segment]:
    """Ssmg の入れ物 → セグメントの並び（目録の順）。"""
    if data[:8] != SSMG or len(data) < 20:
        raise JtdError(f"{stream} が Ssmg の入れ物ではありません")
    count, bsize, nblocks = struct.unpack_from(">III", data, 8)
    if not bsize or bsize % 2 or 20 + nblocks * bsize > len(data):
        raise JtdError(f"{stream} のブロックの大きさ・数が合いません")
    pos = 20 + nblocks * bsize
    out = []
    for _ in range(count):
        if pos + 20 > len(data):
            raise JtdError(f"{stream} のセグメントの目録が途中で終わっています")
        sid, _zero, ln, n, n2 = struct.unpack_from(">IIIII", data, pos)
        if n != n2 or pos + 20 + 4 * n > len(data):
            raise JtdError(f"{stream} のセグメント {sid} の目録が想定と違います")
        blocks = struct.unpack_from(f">{n}I", data, pos + 20)
        pos += 20 + 4 * n
        if any(b >= nblocks for b in blocks) or ln > n * bsize:
            raise JtdError(f"{stream} のセグメント {sid} のブロックが範囲の外です")
        body = b"".join(data[20 + b * bsize:20 + (b + 1) * bsize] for b in blocks)[:ln]
        out.append(Segment(stream, sid, body, tuple(blocks), bsize))
    return out


def text_units(seg: Segment) -> tuple[int, ...] | None:
    """文字のセグメントなら、その単位の並び。文字のセグメントでなければ None。"""
    if seg.data[:8] != TEXT:
        return None
    n = struct.unpack_from(">I", seg.data, 8)[0]
    if 12 + 2 * n > len(seg.data):
        raise JtdError(f"{seg.stream} のセグメント {seg.sid} の字の数が、セグメントの長さを超えています")
    return struct.unpack_from(f">{n}H", seg.data, 12)


def tokens(units):
    """単位の並び → (種類, 位置, 中身, 単位の数)。種類：ch（字）・ctl（制御の字）・rec（記録）・open・close（組の開き・閉じ）。
    記録と組は、両端の長さ・種類が合うことを確かめる（合わなければ JtdError）。"""
    i, n = 0, len(units)
    stack: list[int] = []
    while i < n:
        x = units[i]
        if x == 0x1C:
            if i + 3 > n:
                raise JtdError(f"位置 {i} の記録が途中で終わっています")
            cls, size = units[i + 1], units[i + 2]
            if size < 4 or i + size > n:
                raise JtdError(f"位置 {i} の記録の長さ（{size}）が範囲の外です")
            end = units[i + size - 1]
            body = units[i + 3:i + size - 1]
            if end == 0x1F:
                if size < 7 or units[i + size - 4] != size or units[i + size - 2] != cls:
                    raise JtdError(f"位置 {i} の記録の後ろの長さ・種類が、前と合いません")
                yield "rec", i, (cls, body), size
            elif end == 0x1D:
                stack.append(cls)
                yield "open", i, (cls, body), size
            else:
                raise JtdError(f"位置 {i} の記録の終わりの印が想定と違います（0x{end:04X}）")
            i += size
            continue
        if x == 0x1E:
            if i + 2 > n:
                raise JtdError(f"位置 {i} の組の閉じが途中で終わっています")
            size = units[i + 1]
            if size < 5 or i + size > n or units[i + size - 1] != 0x1F:
                raise JtdError(f"位置 {i} の組の閉じの形が想定と違います")
            cls = units[i + size - 2]
            if not stack or stack.pop() != cls:
                raise JtdError(f"位置 {i} の組の閉じが、開きと対応しません")
            yield "close", i, cls, size
            i += size
            continue
        yield ("ctl" if x < 0x20 else "ch"), i, x, 1
        i += 1
    if stack:
        raise JtdError("閉じていない組があります")


@dataclass
class TextStory:
    """1つの文字のセグメント。k は文書の中の番号（JRef.k）。"""

    k: int
    stream: str
    sid: int
    seg: Segment
    units: tuple[int, ...]

    def unit_offset(self, i: int) -> int:
        """単位の番号 → ストリームの中のバイトの位置。"""
        return self.seg.offset(12 + 2 * i)


@dataclass
class JRef:
    """段落の字の1片の、文書の中の位置：k 番目の文字のセグメントの、s 番目の単位から。"""

    k: int
    s: int


class JtdFile:
    kind = "jtd"

    def __init__(self, blob: bytes, filename: str = ""):
        self.blob = blob
        self.filename = filename
        self.compressed = False
        cfb = CFB(blob)
        names = {n.lower(): n for n in cfb.streams()}
        if "jscompdocument" in names and "documenttext" not in names:
            from .lha import unpack_jscomp

            native = unpack_jscomp(cfb.read(names["jscompdocument"]))
            cfb = CFB(native)
            names = {n.lower(): n for n in cfb.streams()}
            self.compressed = True
        self.cfb = cfb
        if "documenttext" not in names:
            raise JtdError("一太郎の本文（DocumentText）がありません")
        self.protection, self.protection_enforced = "", False
        self.controls: list = []
        self.embeds: list = []
        self.has_macros = False
        self.vba = None
        # 文字のセグメント（窓に出すものすべて）
        self.texts: list[TextStory] = []
        self.segments: dict[str, list[Segment]] = {}
        for name in cfb.streams():
            data = cfb.read(name)
            if data[:8] != SSMG:
                continue
            segs = ssmg(data, name)
            self.segments[name] = segs
            for sg in segs:
                u = text_units(sg)
                if u is not None:
                    self.texts.append(TextStory(len(self.texts), name, sg.sid, sg, u))
        self._used: set[int] = set()
        self.body: list = []
        self.stories: list[tuple[str, str, list]] = []
        self._build()

    # ------------------------------------------------------------------ .docx と同じ入口
    def all_blocks(self):
        yield "body", "DocumentText", self.body
        yield from self.stories

    def marks(self) -> dict:
        return {}  # 一太郎の文書には、書いた所の印（しおり）を足さない（書いた所は、証明の読み直しで確かめる）

    # ------------------------------------------------------------------ 組み立て
    def _story(self, stream: str, sid: int) -> TextStory | None:
        for t in self.texts:
            if t.stream == stream and t.sid == sid:
                return t
        return None

    def _build(self) -> None:
        for t in sorted((t for t in self.texts if t.stream == "DocumentText"), key=lambda t: t.sid):
            self._used.add(t.k)
            self.body += _Build(self, t, 0).run()
        order = list(STREAMS)
        rest = sorted((t for t in self.texts if t.stream != "DocumentText"),
                      key=lambda t: (order.index(t.stream) if t.stream in order else len(order), t.stream, t.sid))
        for t in rest:
            if t.k in self._used:
                continue  # 枠の置き場所から、本文などの中に読み込んだもの
            self._used.add(t.k)
            blocks = _Build(self, t, 0).run()
            kind = STREAMS.get(t.stream, f"その他の文字（{t.stream}）")
            self.stories.append((kind, f"{t.stream}#{t.sid}", blocks))

    def frame_blocks(self, fid: int, depth: int) -> list | None:
        """枠の番号 → 枠の中の段落（レイアウト枠の文字がなければ None：図などの枠）。"""
        t = self._story("LayoutBoxText", fid)
        if t is None or depth > 16:
            return None
        if t.k in self._used:
            return []  # 同じ枠を2度は読まない
        self._used.add(t.k)
        return _Build(self, t, depth).run()


class _Build:
    """1つの文字のセグメント → 段落と表（Word と同じ形）。"""

    def __init__(self, doc: JtdFile, story: TextStory, depth: int):
        self.doc, self.st, self.depth = doc, story, depth
        self.out: list = []
        self.para: Para | None = None
        self.table: Table | None = None
        self.row: list[Cell] | None = None
        self.cell: Cell | None = None
        self.obj: tuple[int | None, tuple] | None = None  # 直前の、字の中の部品（種類, 中身）
        self.groups: list[tuple[int, int, tuple | None]] = []  # 開いている組：(種類, A, 部品)
        self.mark = 0  # 今の段落が始まりうる位置（直前の段落の終わり・行の終わり・マスの始まりの記録の後ろ）

    def run(self) -> list:
        for kind, i, val, n in tokens(self.st.units):
            if kind == "ch":
                self._char(i, val)
            elif kind == "ctl":
                self._ctl(i, val)
            elif kind == "rec":
                self._rec(i, val, n)
            elif kind == "open":
                self._open(val)
            else:
                self._close()
        if self.para is not None and (self.para.segs or self.para.boxes):
            self._end_para(len(self.st.units))
        self._end_row()
        _columns(self.out)
        return self.out

    # 段落・表
    def _p(self) -> Para:
        if self.table is not None and self.cell is None:
            self.table, self.row = None, None  # 行の終わりの後ろに、マスでない字：表はここまで
        if self.para is None:
            self.para = Para(None, self.st.stream, [])
            self.para.cell = self.cell
            self.para.jstart = JRef(self.st.k, self.mark)
        return self.para

    def _end_para(self, at: int) -> None:
        """段落を閉じる。at は段落の終わりの印（段落の終わり・行の終わり・次のマスの記録・文字の終わり）の位置で、
        段落の後ろに字を足すときは、ここに入れる（p.jend）。"""
        p = self._p()
        p.jend = JRef(self.st.k, at)
        (self.cell.blocks if self.cell is not None else self.out).append(p)
        self.para = None

    def _end_row(self) -> None:
        if self.row is not None and self.table is not None:
            self.table.rows.append(self.row)
        self.row, self.cell = None, None

    # 単位
    def _annot(self) -> bool:
        return any(a != 0 for _c, a, _o in self.groups)

    def _char(self, i: int, x: int) -> None:
        p = self._p()
        ch = chr(x)
        if self._annot():
            p.segs.append(Seg(ch, "ruby"))  # 添え書き（ルビ・リンク先など）：窓には出すが、書き換える所にはしない
            return
        last = p.segs[-1] if p.segs else None
        if last is not None and last.kind == "t" and last.node.k == self.st.k and last.node.s + len(last.text) == i:
            last.text += ch
        else:
            p.segs.append(Seg(ch, "t", JRef(self.st.k, i)))

    def _ctl(self, i: int, x: int) -> None:
        if x == 0x0A:
            self._end_para(i)
            self.mark = i + 1
        elif x == 0x0E:
            if self.para is not None:
                self._end_para(i)
            self._end_row()
            self.mark = i + 1
        elif x == 0x0C:
            self._p().segs.append(Seg("\n", "br", JRef(self.st.k, i)))
        elif x == 0x09:
            self._p().segs.append(Seg("\t", "tab", JRef(self.st.k, i)))
        elif x == 0x02:
            obj = next((o for _c, _a, o in reversed(self.groups) if o is not None), None)
            if obj is None or obj[0] != T_FRAME or len(obj[1]) < 6:
                return
            p = self._p()
            blocks = self.doc.frame_blocks(obj[1][5], self.depth + 1)
            if blocks is None:
                p.segs.append(Seg("〔図〕", "other"))
            else:
                p.boxes += blocks

    def _rec(self, i: int, val: tuple, n: int) -> None:
        cls, body = val
        if cls == 0x30:
            # 表のマスの始まり：前のマスの段落を閉じ、表・行がなければ作る
            if self.para is not None:
                self._end_para(i)
            self.mark = i + n
            if self.table is None:
                self.table = Table(None, self.st.stream)
                self.out.append(self.table)
            if self.row is None:
                self.row = []
            c = Cell(None, len(self.table.rows), len(self.row), table=self.table)
            c.x = (body[1], body[2]) if len(body) > 2 else (0, 0)
            self.row.append(c)
            self.cell = c
        elif cls == 0x00:
            self.obj = (body[1] if len(body) > 1 else None, body)

    def _open(self, val: tuple) -> None:
        cls, body = val
        a = body[1] if len(body) > 1 else 0
        self.groups.append((cls, a, self.obj))
        if a:
            self._p().segs.append(Seg(_bracket(self.obj, a)[0], "other"))

    def _close(self) -> None:
        _cls, a, obj = self.groups.pop()
        if a:
            self._p().segs.append(Seg(_bracket(obj, a)[1], "other"))


def _bracket(obj, a: int) -> tuple[str, str]:
    """添え書きの組の、窓での囲み。"""
    t = obj[0] if obj else None
    if t == T_RUBY:
        return "（", "）"
    if t == T_LINK and a == 1:
        return "〔リンク先：", "〕"
    if t == T_FIELD and a == 1:
        return "〔フィールド：", "〕"
    return "〔", "〕"


def _columns(blocks) -> None:
    """表のマスの列の番号と幅：表のすべてのマスの左の位置を並べ、その順位を列の番号にする。"""
    for b in blocks:
        if isinstance(b, Table):
            xs = sorted({c.x[0] for row in b.rows for c in row})
            for r, row in enumerate(b.rows):
                for c in row:
                    c.row = r
                    c.col = xs.index(c.x[0])
                    c.span = max(1, sum(1 for x in xs if c.x[0] <= x < c.x[1]))
                    _columns(c.blocks)
        elif isinstance(b, Para) and b.boxes:
            _columns(b.boxes)


def para_units(doc: JtdFile, p: Para) -> list[tuple[int, int] | None]:
    """段落の文字（p.text）の1字ごとの位置：(文字のセグメントの番号, 単位の番号)。書き換えられない字（添え書きなど）は None。"""
    out: list[tuple[int, int] | None] = []
    for sg in p.segs:
        if sg.hidden:
            continue
        for k in range(len(sg.text)):
            if sg.kind == "t" and sg.node is not None:
                out.append((sg.node.k, sg.node.s + k))
            else:
                out.append(None)
    return out
