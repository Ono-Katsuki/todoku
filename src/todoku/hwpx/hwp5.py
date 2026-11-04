"""본 제품은 한글과컴퓨터의 한/글 문서 파일(.hwp) 공개 문서를 참고하여 개발하였습니다.

HWP 5（한글 97 の後の、한글の2進の形式。Hancom「한글 문서 파일 형식 5.0」）を読む（書かない）。

- 入れ物は CFB（複合文書）。FileHeader（先頭 32 バイトが「HWP Document File」、その後ろに版と属性：bit0 圧縮・bit1 암호・
  bit2 배포용）、DocInfo（文字の形 CHAR_SHAPE・段落の形 PARA_SHAPE など）、BodyText/Section0 …（本文の節）
- 圧縮した文書は、それぞれのストリームが raw deflate
- ストリームは「レコード」の並び：4 バイトの頭（タグ 10 bit・深さ 10 bit・大きさ 12 bit。大きさが 0xFFF なら次の 4 バイト）
- 段落：PARA_HEADER の子に PARA_TEXT（UTF-16 の字。0〜31 は制御の字：字の制御は 1 字、インラインと拡張の制御は 8 字）・
  PARA_CHAR_SHAPE・CTRL_HEADER（拡張の制御の字と同じ順）。表は CTRL_HEADER「tbl 」の子の TABLE と、マスごとの LIST_HEADER
  （マスの番地・結合）とその後ろの段落。ヘッダー/フッター・脚注・図形の中の字（글상자）も LIST_HEADER の後ろの段落
- 누름틀は CTRL_HEADER「%clk」（案内の字は Command の Direction）から、フィールドの終わり（制御の字 4）まで
配布用の文書（본문は ViewText に暗号化して入る）と、パスワードで保護した文書は、中身を読めない。
"""

from __future__ import annotations

import struct
import zlib

from ..cfb import CFB
from ..docx.document import Cell, Control, Para, Seg, Table

SIG = b"HWP Document File"
TAG_BEGIN = 0x10
DOCUMENT_PROPERTIES, ID_MAPPINGS, BIN_DATA, FACE_NAME, BORDER_FILL, CHAR_SHAPE = 0x10, 0x11, 0x12, 0x13, 0x14, 0x15
NUMBERING, BULLET, PARA_SHAPE = 0x17, 0x18, 0x19
PARA_HEADER, PARA_TEXT, PARA_CHAR_SHAPE, PARA_LINE_SEG, PARA_RANGE_TAG, CTRL_HEADER, LIST_HEADER = 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48
TABLE = 0x4D
SHAPE_COMPONENT = 0x4C
SHAPE_PICTURE = 0x55
SHAPE_OLE = 0x54
EQEDIT = 0x58
FORM_OBJECT = 0x5B
CTRL_DATA = 0x57

CHAR_CTRL = {0, 10, 13, 24, 25, 26, 27, 28, 29, 30, 31}  # 1 字
INLINE_CTRL = {4, 5, 6, 7, 8, 9, 19, 20}  # 8 字
# ほか（1・2・3・11・12・14・15・16・17・18・21・22・23）は拡張の制御（8 字。CTRL_HEADER と対応する）


class Hwp5Error(ValueError):
    pass


class Hwp5Locked(Hwp5Error):
    """中身を読めない HWP 5（配布用・パスワード）。title は段階の理由の名前。"""

    def __init__(self, title: str, msg: str):
        super().__init__(msg)
        self.title = title


class Rec:
    __slots__ = ("tag", "level", "data", "kids")

    def __init__(self, tag: int, level: int, data: bytes):
        self.tag, self.level, self.data = tag, level, data
        self.kids: list[Rec] = []

    def __repr__(self) -> str:
        return f"<Rec {self.tag:#x} L{self.level} {len(self.data)}B {len(self.kids)}k>"


def records(data: bytes) -> list[Rec]:
    out = []
    p = 0
    n = len(data)
    while p + 4 <= n:
        (h,) = struct.unpack_from("<I", data, p)
        p += 4
        tag, level, size = h & 0x3FF, (h >> 10) & 0x3FF, h >> 20
        if size == 0xFFF:
            if p + 4 > n:
                raise Hwp5Error("レコードの大きさを読めません")
            (size,) = struct.unpack_from("<I", data, p)
            p += 4
        if p + size > n:
            raise Hwp5Error("レコードが途中で切れています")
        out.append(Rec(tag, level, data[p:p + size]))
        p += size
    return out


def tree(recs: list[Rec]) -> list[Rec]:
    """深さ（level）で木にする：レコードの子は、後ろに続く、より深いレコード。"""
    roots: list[Rec] = []
    stack: list[Rec] = []
    for r in recs:
        while stack and stack[-1].level >= r.level:
            stack.pop()
        (stack[-1].kids if stack else roots).append(r)
        stack.append(r)
    return roots


def ctrl_id(data: bytes) -> str:
    if len(data) < 4:
        return ""
    (v,) = struct.unpack_from("<I", data, 0)
    return "".join(chr((v >> s) & 0xFF) for s in (24, 16, 8, 0))


def _wstr(data: bytes, at: int) -> tuple[str, int]:
    if at + 2 > len(data):
        return "", at
    (n,) = struct.unpack_from("<H", data, at)
    at += 2
    s = data[at:at + 2 * n].decode("utf-16-le", "replace")
    return s, at + 2 * n


class Hwp5Document:
    kind = "hwp"

    def __init__(self, blob: bytes, filename: str = "", partial: dict | None = None):
        """partial：本文の一部だけを読むとき（lazy）：{"cfb": 開いた複合ファイル, "records": レコードのバイト列（None なら本文を読まない）, "part": 節, "pictures": 前までの絵の数}。"""
        self.blob = blob
        self.filename = filename
        try:
            self.cfb = partial["cfb"] if partial else CFB(blob)
            head = self.cfb.read("FileHeader")
        except Exception as e:  # noqa: BLE001
            raise Hwp5Error(f"HWP 5 の入れ物（CFB）を読めません（{e}）") from None
        if head[:len(SIG)] != SIG:
            raise Hwp5Error("HWP 5 の文書ではありません（FileHeader の印がありません）")
        ver, props = struct.unpack_from("<II", head, 32)
        self.version = f"{ver >> 24}.{(ver >> 16) & 0xFF}.{(ver >> 8) & 0xFF}.{ver & 0xFF}"
        self.compressed = bool(props & 1)
        if props & 2:
            raise Hwp5Locked("パスワード", "パスワードで保護した HWP 5 の文書です（本文が暗号化されていて、いまは読めません）。"
                             "本人に、保護を外して保存した文書を頼んでください")
        names = {n.lower(): n for n in self.cfb.streams()}
        if props & 4 or any(n.startswith("viewtext/") for n in names):
            raise Hwp5Locked("配布用の文書", "配布用の文書（배포용 문서）です。本文が暗号化されていて（ViewText）、読めません。"
                             "作った人（発行した機関）に、配布用でない文書（HWP・HWPX）を頼んでください")
        if props & 16:
            raise Hwp5Locked("DRM", "DRM で保護した HWP 5 の文書です。読めません")
        self.parts = sorted((n for k, n in names.items() if k.startswith("bodytext/section")),
                            key=lambda n: int("".join(ch for ch in n.rsplit("/", 1)[-1] if ch.isdigit()) or 0))
        if not self.parts:
            raise Hwp5Error("本文の節（BodyText/Section0）がありません")
        self.part = self.parts[0]
        self.controls: list[Control] = []
        self.pictures: list[dict] = [{"n": i + 1, "stub": True} for i in range((partial or {}).get("pictures", 0))]
        self.notes: list[tuple[str, list]] = []
        self.comments: list = []
        self.headers: list[tuple[str, str, list]] = []
        self.unread: list[str] = []
        self.protection = ""
        self.protection_enforced = False
        self.has_macros = False
        self.vba = None
        self.scripts = []
        self.numbered = 0  # 番号・箇条書きの段落の数（HWP 5 の番号の字は、いまは窓に出さない）
        self._docinfo()
        self.body: list = []
        self.raw_texts: list[str] = []  # 検査用：すべての PARA_TEXT の字（制御の字を除く）
        if partial is not None:
            if partial.get("records") is not None:
                self.part = partial.get("part", self.part)
                self.body += self._paras(tree(records(partial["records"])), self.part, None)
        else:
            for part in self.parts:
                roots = tree(records(self._stream(part)))
                self.body += self._paras(roots, part, None)
        self.stories: list[tuple[str, str, list]] = []
        for kind, part, blocks in self.headers:
            self.stories.append((kind, part, blocks))
        for kind, blocks in self.notes:
            self.stories.append((kind, self.part, blocks))
        if self.comments:
            self.stories.append(("comments", self.part, self.comments))

    def _stream(self, name: str) -> bytes:
        raw = self.cfb.read(name)
        if self.compressed:
            try:
                return zlib.decompress(raw, -15)
            except zlib.error as e:
                raise Hwp5Error(f"{name} の圧縮を解けません（{e}）") from None
        return raw

    def all_blocks(self):
        yield "body", self.part, self.body
        for kind, part, blocks in self.stories:
            yield kind, part, blocks

    def marks(self) -> dict:
        return {}

    # ------------------------------------------------------------------ DocInfo
    def _docinfo(self) -> None:
        self.under: list[bool] = []
        self.para_head: list[int] = []
        try:
            recs = records(self._stream("DocInfo"))
        except (Hwp5Error, KeyError):
            return
        for r in recs:
            if r.tag == CHAR_SHAPE and len(r.data) >= 50:
                (attr,) = struct.unpack_from("<I", r.data, 46)
                self.under.append(((attr >> 2) & 3) == 1)
            elif r.tag == PARA_SHAPE and len(r.data) >= 4:
                (attr,) = struct.unpack_from("<I", r.data, 0)
                self.para_head.append((attr >> 23) & 3)

    # ------------------------------------------------------------------ 段落
    def _paras(self, nodes: list[Rec], part: str, cell: Cell | None) -> list:
        out: list = []
        for r in nodes:
            if r.tag == PARA_HEADER:
                out += self._para(r, part, cell)
        return out

    def _list_blocks(self, nodes: list[Rec], part: str, cell: Cell | None) -> list:
        """LIST_HEADER の後ろの段落（子のどこかの深さにある LIST_HEADER も）。"""
        out: list = []
        for r in nodes:
            if r.tag == PARA_HEADER:
                out += self._para(r, part, cell)
            elif r.tag != CTRL_HEADER:
                out += self._list_blocks(r.kids, part, cell)
        return out

    def _para(self, r: Rec, part: str, cell: Cell | None) -> list:
        p = Para(r, part, [], cell=cell)
        p.el = None
        if len(r.data) >= 10:
            (ps,) = struct.unpack_from("<H", r.data, 8)
            if ps < len(self.para_head) and self.para_head[ps]:
                self.numbered += 1
        text = b""
        shapes: list[tuple[int, int]] = []
        ctrls: list[Rec] = []
        for k in r.kids:
            if k.tag == PARA_TEXT:
                text = k.data
            elif k.tag == PARA_CHAR_SHAPE:
                shapes = [struct.unpack_from("<II", k.data, i) for i in range(0, len(k.data) - 7, 8)]
            elif k.tag == CTRL_HEADER:
                ctrls.append(k)
        after: list = []
        units = struct.unpack_from(f"<{len(text) // 2}H", text, 0) if text else ()
        fields: list[Control | None] = []
        raw = []
        i = 0
        ci = 0
        buf: list[str] = []

        def under_at(pos: int) -> bool:
            sid = 0
            for at, s in shapes:
                if at <= pos:
                    sid = s
            return sid < len(self.under) and self.under[sid]

        def flush(at: int) -> None:
            if buf:
                t = "".join(buf)
                buf.clear()
                f = next((x for x in reversed(fields) if x is not None), None)
                sg = Seg(t, "t", None, None, False, under_at(at), f)
                if f is not None and f.kind != "field":
                    f.result_segs.append(sg)
                p.segs.append(sg)

        def fixed(t: str) -> None:
            p.segs.append(Seg(t, "other", None, None, False, False, Control("field", "hwp5-fixed", "", None, part, p)))

        while i < len(units):
            u = units[i]
            if u >= 32:
                if 0xD800 <= u < 0xDC00 and i + 1 < len(units):
                    buf.append(chr(0x10000 + ((u - 0xD800) << 10) + (units[i + 1] - 0xDC00)))
                    raw.append(buf[-1])
                    i += 2
                    continue
                buf.append(chr(u))
                raw.append(chr(u))
                i += 1
                continue
            if u in CHAR_CTRL:
                if u == 10:
                    flush(i)
                    p.segs.append(Seg("\n", "br", None))
                elif u == 30:
                    buf.append(" ")
                elif u == 31:
                    buf.append("　")
                elif u == 24:
                    flush(i)
                    p.segs.append(Seg("-", "other", None))
                i += 1
                continue
            flush(i)
            if u in INLINE_CTRL:
                if u == 9:
                    p.segs.append(Seg("\t", "tab", None, None, False, under_at(i)))
                elif u == 4:  # フィールドの終わり
                    if fields:
                        f = fields.pop()
                        if f is not None and f.kind != "field":
                            self._close(f, p, part)
                i += 8
                continue
            # 拡張の制御：CTRL_HEADER と同じ順
            c = ctrls[ci] if ci < len(ctrls) else None
            ci += 1
            i += 8
            if c is None:
                continue
            cid = ctrl_id(c.data)
            if u == 3:  # フィールドの始まり
                fields.append(self._field(c, cid, p, part))
                continue
            if cid == "tbl ":
                after += self._table(c, part)
            elif cid in ("head", "foot"):
                blocks = self._list_blocks(c.kids, part, None)
                if blocks:
                    self.headers.append(("header" if cid == "head" else "footer", part, blocks))
            elif cid in ("fn  ", "en  "):
                kind = "footnotes" if cid == "fn  " else "endnotes"
                n = sum(1 for k_, _b in self.notes if k_ == kind) + 1
                p.segs.append(Seg(f"※{n}", "note", None))
                self.notes.append((kind, self._list_blocks(c.kids, part, None)))
            elif cid == "tcmt":
                self.comments += self._list_blocks(c.kids, part, None)
            elif cid == "eqed":
                eq = next((k for k in _walk(c) if k.tag == EQEDIT), None)
                script = _wstr(eq.data, 4)[0] if eq is not None else ""
                fixed(f"〔数式: {script.strip()}〕")
            elif cid == "gso ":
                self._shape(c, p, part, fixed)
            elif cid == "form":
                fo = next((k for k in _walk(c) if k.tag == FORM_OBJECT), None)
                props = _form_props(fo.data) if fo is not None else {}
                fixed(f"〔フォームの部品: {props.get('Caption') or props.get('Name') or ''}〕")
                self.unread.append("HWP 5 のフォームの部品（양식 개체）の値は、いまは読みません（HWPX で保存し直せば読めます）")
            elif cid == "tdut":
                main, at = _wstr(c.data, 4)
                sub, _ = _wstr(c.data, at)
                fixed(main + (f"（{sub}）" if sub else ""))
            elif cid == "tcps":
                s, _ = _wstr(c.data, 4)
                fixed(s)
        flush(len(units))
        for f in fields:
            if f is not None and f.kind != "field":
                self._close(f, p, part)
        self.raw_texts.append("".join(raw))
        return [p] + after

    def _field(self, c: Rec, cid: str, p: Para, part: str) -> Control | None:
        if cid == "%clk":
            # 属性（UINT32）・その他の属性（BYTE）・Command（WORD の長さと UTF-16）
            cmd, _ = _wstr(c.data, 9)
            direction = ""
            import re

            m = re.search(r"Direction:wstring:(\d+):", cmd)
            if m:
                direction = cmd[m.end():m.end() + int(m[1])].strip()
            ctl = Control("text", "hwp5-clickhere", "", c, part, p)
            ctl.direction = direction
            ctl.help_text = ""
            ctl.dirty = False
            self.controls.append(ctl)
            return ctl
        if cid in ("%hlk", "%bmk", "%unk") or not cid.startswith("%"):
            return None
        return Control("field", "hwp5-field", cid, c, part, p)

    def _close(self, ctl: Control, p: Para, part: str) -> None:
        content = "".join(s.text for s in ctl.result_segs)
        if not ctl.result_segs:
            sg = Seg("", "field", None, None, False, False, ctl)
            p.segs.append(sg)
            ctl.result_segs = [sg]
        ctl.placeholder = content == "" or (bool(ctl.direction) and content.rstrip() == ctl.direction)
        ctl.value = None if ctl.placeholder else content

    def _shape(self, c: Rec, p: Para, part: str, fixed) -> None:
        alt = _wstr(c.data, 44)[0].strip() if len(c.data) >= 46 else ""
        recs = list(_walk(c))
        p.boxes += self._list_blocks(c.kids, part, None)
        if any(k.tag == SHAPE_PICTURE for k in recs):
            self.pictures.append({"n": len(self.pictures) + 1, "alt": alt, "decorative": False, "para": p})
            fixed(f"〔図: {alt}〕" if alt else f"〔図{len(self.pictures)}（説明なし）〕")
        elif any(k.tag == SHAPE_OLE for k in recs):
            fixed(f"〔埋め込み: {alt or 'オブジェクト'}〕")
            self.unread.append(f"埋め込み（OLE）の中身は読めません（{alt or '説明なし'}）")

    def _table(self, c: Rec, part: str) -> list:
        t = Table(c, part)
        cells: list[Cell] = []
        kids = c.kids
        first = next((i for i, r in enumerate(kids) if r.tag == TABLE), 0)
        caption = self._list_blocks(kids[:first], part, None)  # 表の見出し（TABLE の前の LIST_HEADER と段落）
        k = first
        while k < len(kids):
            r = kids[k]
            if r.tag == LIST_HEADER and len(r.data) >= 16:
                (n_para,) = struct.unpack_from("<H", r.data, 0)
                col, row, cs, rs = struct.unpack_from("<HHHH", r.data, 8)
                cell = Cell(r, row, col, max(1, cs), "restart" if rs > 1 else "", table=t)
                cell.rowspan = max(1, rs)
                paras = []
                j = k + 1
                while j < len(kids) and kids[j].tag == PARA_HEADER and len(paras) < max(n_para, 1):
                    paras.append(kids[j])
                    j += 1
                blocks: list = []
                for pr in paras:
                    blocks += self._para(pr, part, cell)
                cell.blocks = blocks
                cells.append(cell)
                k = j
                continue
            k += 1
        nrows = max((x.row + x.rowspan for x in cells), default=0)
        rows: list[list[Cell]] = [[] for _ in range(nrows)]
        for x in cells:
            rows[x.row].append(x)
            for rr in range(x.row + 1, x.row + x.rowspan):
                rows[rr].append(Cell(x.node, rr, x.col, x.span, "continue", table=t))
        for row in rows:
            row.sort(key=lambda x: x.col)
        t.rows = rows
        return [t] + caption


def _walk(r: Rec):
    for k in r.kids:
        yield k
        yield from _walk(k)


def _form_props(data: bytes) -> dict:
    """양식 개체の属性の字（「Name:wstring:9:CheckBox1 Caption:wstring:…」の形）。"""
    import re

    out = {}
    try:
        s = data.decode("utf-16-le", "replace")
    except Exception:  # noqa: BLE001
        return out
    for m in re.finditer(r"(\w+):wstring:(\d+):", s):
        out.setdefault(m[1], s[m.end():m.end() + int(m[2])])
    return out
