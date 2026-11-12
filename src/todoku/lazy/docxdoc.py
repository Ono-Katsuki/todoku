"""大きな Word（.docx）を分けて読む：document.xml を流し読みして、本文の直下の要素（段落・表）の位置の索引だけを持つ。

- 開く：expat で流し読みして、ブロック（body の子）の位置・段落の数・表の数を数える。段落番号は流しながら数え（Numbering）、まとまり（chunk。およそ
  CHUNK_BYTES の XML）ごとに数えた値を控える。テキストボックス・コンテンツコントロール・フォームの欄・図・埋め込みのあるブロックだけは、そのブロックの木を
  作って数える（数え方を窓の読み手と同じにするため）。メモリは索引の分だけで、段落の数に比例して木を持たない
- 読む：頼まれた段落を含むまとまりだけの小さな文書を作り（頭と終わりの XML + ブロック + スタイルなどの部品）、いままでの読み手（DocxDocument・Spots）で読む。
  前までの数（base）を渡すので、段落・表の番号と欄の id、段落番号は全体を読んだときと同じ
- 検査：まとまりごとに、読み落としと書き込みの試しを行う（小さな文書の上で）。控えから続きを進められる
- 書く：いままでの api.fill（全体を読む）
"""

from __future__ import annotations

import bisect
import io
import re
import xml.parsers.expat as expat
import zipfile
import zlib
from array import array

from .. import limits
from .blocks import STORY_JA, BlocksLazy
from .textstream import Source
from .xlsxdoc import Member, SeekableInflate, _no_doctype

CHUNK_BYTES = 1 << 20  # 1 回の小さな文書の XML の大きさの目安
KEEP_PART = 1 << 24  # 小さな文書に写す部品の大きさの上限（これより大きな部品は、名前だけを残す）
_COMPLEX = frozenset(("txbxContent", "sdt", "customXml", "moveTo", "AlternateContent", "ffData", "object", "OLEObject", "control",
                      "blip", "imagedata", "graphicData", "pict"))
# フィールドの印（w:fldChar）を、本文の流れの順に数えるときにたどる要素（docx/document の _blocks・_inline と同じ道）
_FIELD_BLOCK = frozenset(("p", "tbl", "tr", "tc", "sdt", "sdtContent", "customXml", "ins", "moveTo"))
_FIELD_INLINE = frozenset(("hyperlink", "smartTag", "fldSimple", "bdo", "dir", "del", "moveFrom", "r"))
_STORY_KINDS = ("header", "footer", "footnotes", "endnotes", "comments")


def word_main(path: str) -> str | None:
    """分けて読める大きな Word なら、本体の部品の名前。小さな文書・ZIP に問題のある文書は None（いままでの道）。"""
    try:
        with zipfile.ZipFile(path) as z:
            infos = z.infolist()
            names = [i.filename for i in infos]
            if len({n.lower() for n in names}) != len(names) or any(i.flag_bits & 1 for i in infos):
                return None
            rels = z.read("_rels/.rels") if "_rels/.rels" in names else b""
            m = re.search(rb'Target="/?([^"]*document[^"]*\.xml)"', rels)
            main = m[1].decode() if m else "word/document.xml"
            if main not in names or not main.startswith("word/"):
                return None
            info = z.getinfo(main)
            if info.file_size < limits.LAZY_MIN_BYTES:
                return None
            settings = z.read("word/settings.xml") if "word/settings.xml" in names else b""
            if b"documentProtection" in settings and re.search(rb'edit="readOnly"', settings):
                return None  # 編集できる範囲（permStart〜permEnd）はブロックをまたぐので、全体で読む
            return main
    except (zipfile.BadZipFile, OSError, ValueError, KeyError):
        return None


class DocxLazy(BlocksLazy):
    kind = "docx"
    streams = False

    def __init__(self, path: str, name: str, main: str):
        from ..docx.document import DocxDocument

        self.path = path
        self.name = name
        self.main = main
        self.src = Source(path)
        self.z = zipfile.ZipFile(path)
        self.member = Member(self.src, self.z.getinfo(main))
        self.reader = SeekableInflate(self.member)
        self.size = self.src.size
        self._sha = None
        self.kept = self._kept_parts()
        self.offs = array("Q")  # ブロックの始まり
        self.para_first = array("Q")  # ブロックの前までの段落の数
        self.chunks: list[dict] = []  # まとまり：{"first": ブロック, "base": 前までの数, "para": 前までの段落の数}
        self.paragraphs = 0
        self.tables = 0
        self.body_paragraphs = 0
        self.chars = 0
        self.control_names: list[str] = []
        self.pictures: list[dict] = []
        self.object_stories: list[str] = []
        self.objects_blocks: list[int] = []  # 埋め込み・グラフ・SmartArt のあるブロック（窓の終わりの節を作るとき読む）
        self.tracked = 0
        self.tracked_ins = 0  # 変更履歴で挿入された字の所（w:ins・w:moveTo）の数
        self.head_xml = b""
        self.tail_xml = b""
        self.data_end = 0
        h = DocxDocument(self._mini(b"", ()), name, base={"stories": ()})
        self._h = h
        self.protection = h.protection
        self.protection_enforced = h.protection_enforced
        self.has_macros = h.has_macros
        self.story_list: list[dict] = []
        self._scan()
        self._stories()
        self.est_chars = int(self.chars * 1.25) + self.paragraphs * 12 + len(self.control_names) * 60
        self._h = None

    def close(self) -> None:
        self.z.close()
        self.src.close()

    @property
    def sha256(self) -> str:
        if self._sha is None:
            from .filebytes import sha256_file

            self._sha = sha256_file(self.path)
        return self._sha

    # ------------------------------------------------------------------ 部品
    def _kept_parts(self) -> dict[str, bytes]:
        """小さな文書に写す部品（本体のほか）。画像・大きな部品は名前だけ（中身のある部品は要るときに読む：_real）。"""
        out: dict[str, bytes] = {}
        for i in self.z.infolist():
            n = i.filename
            if n == self.main or n.endswith("/"):
                continue
            low = n.lower()
            if low.startswith(("word/media/", "word/embeddings/", "word/charts/", "word/diagrams/")) or i.file_size > KEEP_PART:
                if low.endswith(".rels") and i.file_size <= KEEP_PART:
                    out[n] = self._read_part(i)
                else:
                    out[n] = b""
            else:
                out[n] = self._read_part(i)
        return out

    def _read_part(self, info) -> bytes:
        return Member(self.src, info).read()

    def _mini(self, xml: bytes, real: tuple = ()) -> bytes:
        """小さな文書の ZIP。real：名前だけにしない部品。"""
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
            for n, data in self.kept.items():
                z.writestr(n, self._read_part(self.z.getinfo(n)) if n in real else data)
            z.writestr(self.main, xml or (self._empty_main()))
        return buf.getvalue()

    def _empty_main(self) -> bytes:
        # 頭の XML はまだ分からない：本体の頭だけを読んで、body の開きと閉じを作る
        if self.head_xml:
            return self.head_xml + self.tail_xml
        got = bytearray()
        for c in self.member.chunks():
            got += c
            m = re.search(rb"<(?:[\w.-]+:)?body\b[^>]*>", got)
            if m:
                head = bytes(got[:m.end()])
                root = re.findall(rb"<([\w.:-]+)", head)
                # 宣言のほかで最初の要素と body を閉じる
                first = next(x for x in root if not x.startswith(b"?"))
                body = re.match(rb"<([\w.:-]+)", m[0])[1]
                return head + b"</" + body + b"></" + first + b">"
        return b"<document/>"

    # ------------------------------------------------------------------ 索引
    def _scan(self) -> None:
        from ..docx.document import _carried_field, iter_paras
        from ..docx.window import _tables
        from ..docx.xmlpos import parse

        h = self._h
        numbering = h.numbering
        style_num: dict = {}

        def num_of_style(s):
            if s in style_num:
                return style_num[s]
            got = None
            x = s
            seen = 0
            while x and seen < 20:
                if x in h.style_num:
                    got = h.style_num[x]
                    break
                x = h.style_based.get(x, "")
                seen += 1
            style_num[s] = got
            return got

        p = expat.ParserCreate()
        p.buffer_text = True
        p.StartDoctypeDeclHandler = _no_doctype
        names: dict[str, str] = {}
        stack: list[str] = []
        st = {"at": 0, "complex": False, "paras": 0, "tables": 0, "pstyle": "", "numid": None, "ilvl": 0, "body": False, "body_end": None, "chars": 0, "del": 0, "ins": 0}
        pending: list = []
        fst: list[dict] = []  # 開いているフィールド（段落・まとまりをまたぐもの：目次など）
        fp = {"bad": 0, "p": 0, "instr": False, "at": ()}
        buf = bytearray()
        buf_at = 0
        tail_seen = {}
        size = self.member.size or 1
        data_start = {"v": None}

        def close_chunk_if_needed(at: int) -> None:
            n = len(self.offs)
            if not self.chunks or at - self.offs[self.chunks[-1]["first"]] >= CHUNK_BYTES:
                self.chunks.append({"first": n, "para": self.paragraphs, "base": self._base_now(numbering),
                                    "fields": [(f, f["phase"], f["instr"]) for f in fst]})

        def finish_block(at_end: int) -> None:
            nonlocal buf, buf_at
            a = st["at"]
            if st["complex"]:
                piece = bytes(buf[a - buf_at:at_end - buf_at]) if a >= buf_at else self.reader.read(a, at_end)
                xml = self.head_xml + piece + self.tail_guess
                root = parse(xml)
                top = root.kids[0]
                body = top.find("body")
                h._carry = [_carried_field({"phase": ph, "instr": ins, "locked": lk}) for ph, ins, lk in fp["at"]]
                blocks = h._blocks(body, self.main)
                paras = sum(1 for _ in iter_paras(blocks))
                tables = sum(1 for _ in _tables(blocks))
                self._after_tree(h, blocks, len(self.offs) - 1)
                self.paragraphs += paras
                self.tables += tables
            else:
                for num in pending:
                    numbering.label(num[0], num[1])
                self.paragraphs += st["paras"]
                self.tables += st["tables"]
            self.chars += st["chars"]
            pending.clear()
            st.update(complex=False, paras=0, tables=0, chars=0)

        def start(name, attrs):
            local = names.get(name)
            if local is None:
                local = names[name] = name.rsplit(":", 1)[-1]
            depth = len(stack)
            if depth == 1 and local == "body":
                st["body"] = True
                data_start["v"] = p.CurrentByteIndex
            elif depth == 2 and st["body"]:
                at = p.CurrentByteIndex
                if self.head_xml == b"":
                    self.head_xml = bytes(buf[:_tag_end(buf, data_start["v"] - buf_at)])
                    # 終わりの XML は、あとで本物に置き換える（木を作るときだけに使う）
                    self.tail_guess = b"</" + self._raw_names["body"] + b"></" + self._raw_names["top"] + b">"
                close_chunk_if_needed(at)
                self.offs.append(at)
                self.para_first.append(self.paragraphs)
                st["at"] = at
                fp["at"] = [(f["phase"], f["instr"], f["locked"]) for f in fst]
            if st["body"] and depth >= 2:
                if not fp["bad"] and (local in _FIELD_BLOCK or (fp["p"] and local in _FIELD_INLINE)):
                    if local == "p":
                        fp["p"] += 1
                else:
                    if not fp["bad"] and stack[-1] == "r":
                        if local == "fldChar":
                            typ = next((v for k, v in attrs.items() if k.rsplit(":", 1)[-1] == "fldCharType"), "")
                            if typ == "begin":
                                lock = next((v for k, v in attrs.items() if k.rsplit(":", 1)[-1] == "fldLock"), "")
                                fst.append({"phase": "instr", "instr": "", "locked": lock in ("1", "true", "on")})
                            elif typ == "separate" and fst:
                                fst[-1]["phase"] = "result"
                            elif typ == "end" and fst:
                                fst.pop()["closed"] = True
                        elif local == "instrText":
                            fp["instr"] = True
                    fp["bad"] += 1
                if local == "p":
                    if stack[-1] not in ("body", "tc"):
                        st["complex"] = True
                    st["paras"] += 1
                    st["pstyle"] = ""
                    st["numid"] = None
                    st["ilvl"] = 0
                elif local == "tbl":
                    if stack[-1] not in ("body", "tc"):
                        st["complex"] = True
                    st["tables"] += 1
                elif local == "tr":
                    if stack[-1] != "tbl":
                        st["complex"] = True
                elif local == "tc":
                    if stack[-1] != "tr":
                        st["complex"] = True
                elif local in _COMPLEX:
                    st["complex"] = True
                elif local == "pStyle":
                    if stack[-1] == "pPr":
                        st["pstyle"] = next((v for k, v in attrs.items() if k.rsplit(":", 1)[-1] == "val"), "")
                elif local == "numId":
                    if stack[-1] == "numPr":
                        st["numid"] = next((v for k, v in attrs.items() if k.rsplit(":", 1)[-1] == "val"), "")
                elif local == "ilvl":
                    if stack[-1] == "numPr":
                        st["ilvl"] = int(next((v for k, v in attrs.items() if k.rsplit(":", 1)[-1] == "val"), 0) or 0)
                elif local == "delText":
                    st["del"] += 1
                elif local in ("ins", "moveTo", "moveFrom"):
                    # 字の挿入（段落の中の w:ins）・移る前の字（w:moveFrom：字は w:t）：書式の中の印（段落の終わりの印・行・番号）は数えない
                    if stack[-1] not in ("rPr", "pPr", "trPr", "tcPr", "tblPr", "numPr", "sectPr"):
                        st["del" if local == "moveFrom" else "ins"] += 1
            if depth == 0:
                self._raw_names = {"top": name.encode()}
            elif depth == 1 and local == "body":
                self._raw_names["body"] = name.encode()
            stack.append(local)

        def end(name):
            local = stack.pop()
            depth = len(stack)
            if st["body"] and depth >= 2:
                if fp["bad"]:
                    fp["bad"] -= 1
                    fp["instr"] = False
                elif local == "p":
                    fp["p"] -= 1
            if local == "p" and st["body"] and depth >= 2:
                if st["numid"] is not None:
                    if st["numid"]:
                        pending.append((st["numid"], st["ilvl"]))
                elif st["pstyle"]:
                    num = num_of_style(st["pstyle"])
                    if num is not None and num[0]:
                        pending.append(num)
            if depth == 2 and st["body"]:
                at_end = _tag_end(buf, p.CurrentByteIndex - buf_at) + buf_at
                finish_block(at_end)
                nonlocal_trim(at_end)
            elif depth == 1 and local == "body":
                st["body"] = False
                st["body_end"] = p.CurrentByteIndex
                self.data_end = p.CurrentByteIndex

        def nonlocal_trim(upto: int) -> None:
            tail_seen["trim"] = upto

        def chars(data):
            if st["body"]:
                st["chars"] += len(data)
                if fp["instr"] and fst:
                    fst[-1]["instr"] += data

        p.StartElementHandler = start
        p.EndElementHandler = end
        p.CharacterDataHandler = chars
        total = 0
        crc = 0
        for c in self.reader.scan():
            buf += c
            total += len(c)
            crc = zlib.crc32(c, crc)
            p.Parse(c, False)
            limits.progress(f"{self.main} の段落の位置を調べています", total, size)
            t = tail_seen.pop("trim", None)
            if t is not None and t - buf_at > 0:
                del buf[:t - buf_at]
                buf_at = t
            if st["body_end"] is None and len(self.offs):
                if st["at"] > buf_at:
                    del buf[:st["at"] - buf_at]
                    buf_at = st["at"]
                if not st["complex"] and len(buf) > 1 << 22:
                    # 字だけの大きなブロック：手放す（あとで木が要るとわかったら、読み直す）。閉じタグの途中のために後ろは残す
                    drop = len(buf) - (1 << 16)
                    del buf[:drop]
                    buf_at += drop
        p.Parse(b"", True)
        if crc & 0xFFFFFFFF != self.member.info.CRC:
            raise ValueError(f"{self.main}：CRC が一致しません（ファイルが壊れています）")
        self.total = total
        if self.head_xml == b"":
            raise ValueError("本文のない文書です")  # いままでの道で読む
        at = st["body_end"] if st["body_end"] is not None else total
        self.tail_xml = self.reader.read(at, total) or self.tail_guess
        self.data_end = at
        self.body_paragraphs = self.paragraphs
        self.tracked = st["del"]
        self.tracked_ins = st["ins"]
        # まとまりの始まりで開いているフィールド（その時の phase・命令の字）と、後ろで閉じるか
        for ch in self.chunks:
            if "story" in ch:
                continue
            ch["fields"] = tuple({"phase": ph, "instr": ins, "locked": f["locked"], "closed": bool(f.get("closed"))}
                                 for f, ph, ins in ch["fields"])

    def _base_now(self, numbering) -> dict:
        h = self._h
        return {"para": self.paragraphs, "table": self.tables, "controls": len(self.control_names),
                "embeds": list(h.embeds), "charts": len(h.charts), "diagrams": len(h.diagrams), "pictures": len(h.pictures),
                "counters": {k: list(v) for k, v in numbering.counters.items()}, "stories": ()}

    def _after_tree(self, h, blocks, block_index: int) -> None:
        """木を作って読んだブロック：欄の名前・絵・埋め込みの控えだけを残す（木は手放す）。"""
        from ..docx.window import _pictures

        for c in h.controls:
            k = len(self.control_names) + 1
            name = c.name or f"ctl{k}"
            base = name
            j = 2
            used = self._used_names()
            while name in used:
                name = f"{base}#{j}"
                j += 1
            used.add(name)
            self.control_names.append(name)
        h.controls = []
        fresh = [p for p in h.pictures if not p.get("stub")]
        if fresh:
            keep = h.pictures
            h.pictures = fresh
            self.pictures += _pictures(h)
            h.pictures = keep
        if any(not e.get("stub") for e in h.embeds) or any(not e.get("stub") for e in h.charts) or any(not e.get("stub") for e in h.diagrams):
            if block_index >= 0:
                self.objects_blocks.append(block_index)
            else:
                self._story_has_objects = True
        h.pictures = [{"n": p["n"], "stub": True} for p in h.pictures]
        h.embeds = [{**{k: v for k, v in e.items() if k != "para"}, "stub": True} for e in h.embeds]
        h.charts = [{"n": e["n"], "stub": True} for e in h.charts]
        h.diagrams = [{"n": e["n"], "stub": True} for e in h.diagrams]

    def _used_names(self) -> set:
        u = self.__dict__.get("_used")
        if u is None:
            u = self._used = set(self.control_names)
        return u

    def _stories(self) -> None:
        """ヘッダー・フッター・脚注・コメント：部品ごとに 1 つのまとまり（本文の後ろの番号。いままでの読み手と同じ順）。"""
        from ..docx.document import _join, iter_paras
        from ..docx.window import _tables

        h = self._h
        for _rid, (typ, target) in h.rels.items():
            kind = typ.rsplit("/", 1)[-1]
            if kind not in _STORY_KINDS:
                continue
            name = _join(self.main, target)
            if not h.zip.has(name):
                continue
            base = self._base_now(h.numbering)
            base["stories"] = (name,)
            blocks = h._story(name)
            paras = sum(1 for _ in iter_paras(blocks))
            tables = sum(1 for _ in _tables(blocks))
            self._story_has_objects = False
            self.chars += sum(len(p.text) for p in iter_paras(blocks))
            self._after_tree(h, blocks, -1)
            if self._story_has_objects:
                self.object_stories.append(name)
            self.chunks.append({"story": name, "part": name, "kind": kind, "first": len(self.offs), "para": self.paragraphs, "base": base})
            self.story_list.append({"kind": kind, "part": name, "paragraphs": paras})
            self.paragraphs += paras
            self.tables += tables
            h.parts.pop(name, None)

    # ------------------------------------------------------------------ 小さな文書
    def _chunk_xml(self, i: int) -> bytes:
        ch = self.chunks[i]
        if "story" in ch:
            return self.head_xml + self.tail_xml
        first = ch["first"]
        nxt = self.chunks[i + 1]["first"] if i + 1 < len(self.chunks) else len(self.offs)
        a = self.offs[first] if first < len(self.offs) else self.data_end
        b = self.offs[nxt] if nxt < len(self.offs) else self.data_end
        return self.head_xml + self.reader.read(a, b) + self.tail_xml

    def chunk_doc(self, i: int, real: tuple = ()):
        from ..docx.document import DocxDocument

        return DocxDocument(self._mini(self._chunk_xml(i), real), self.name, base=self.base_of(i))

    def base_of(self, i: int) -> dict:
        base = dict(self.chunks[i]["base"])
        base["control_names"] = self.control_names[:base.pop("controls")]
        base["fields"] = self.chunks[i].get("fields", ())
        nxt = self.chunks[i + 1] if i + 1 < len(self.chunks) else {}
        # この部分の終わりで開いているフィールドが、後ろで閉じるか（最後のまとまりの後ろには、閉じる印がない）
        base["fields_end"] = tuple(f["closed"] for f in nxt["fields"]) if "fields" in nxt else None
        return base

    # ------------------------------------------------------------------ 概要
    def overview(self) -> dict:
        return {"bytes": self.size, "window_chars": self.est_chars, "paragraphs": self.paragraphs, "tables": self.tables,
                "fields_named": len(self.control_names), "pictures": len(self.pictures),
                "stories": [dict(s) for s in self.story_list[:40]],
                "protection": {"edit": self.protection, "enforced": self.protection_enforced},
                "ranges": "paragraphs（段落の範囲）か cursor"}

    def overview_text(self) -> str:
        lines = [f"# {self.name}", "", f"- 元ファイルの SHA-256: `{self.sha256}`（{self.size:,} バイト）、Word",
                 "- この文書は読むための窓です。書き込みは `todoku fill` で、書いた欄の部分だけを変えます。"]
        if self.protection_enforced and self.protection:
            lines.append(f"- **編集の制限: {self.protection}**")
        if self.has_macros:
            lines.append("- マクロがあります（窓の終わりの節にソースがあります）")
        lines += [f"- 段落 {self.paragraphs:,}・表 {self.tables:,}・名前のある入力欄 {len(self.control_names):,}・絵 {len(self.pictures):,}",
                  f"- 大きいので、ここでは本文を読んでいません（窓の全体は見積もりで {self.est_chars:,} 字）。read_window に paragraphs: \"1-200\" のように段落の範囲を渡して読みます"
                  "（cursor をたどれば、ヘッダー・脚注・埋め込みまで、すべてを読めます。max_chars で 1 回の大きさを変えられます）。"
                  "list_fields に paragraphs を渡すと、その範囲の欄を出します。"]
        if self.story_list:
            first = self.body_paragraphs
            where = []
            for s in self.story_list[:20]:
                where.append(f"{STORY_JA.get(s['kind'], s['kind'])}（{s['part']}）：段落 {first + 1}〜{first + s['paragraphs']}")
                first += s["paragraphs"]
            lines += ["", "## 本文のほかの部品", ""] + [f"- {w}" for w in where]
        return "\n".join(lines) + "\n"

    def has_extras(self) -> bool:
        return bool(self.objects_blocks or self.object_stories or self.has_macros or self.pictures or self._image_parts())

    def _image_parts(self) -> list[str]:
        from ..imaging.docimages import _IMG_EXT, _ext

        return [n for n in self.kept if _ext(n) in _IMG_EXT and n.startswith("word/")]

    def images_for_view(self) -> list[dict]:
        from ..imaging.docimages import list_images

        return list_images("docx", self._mini(b"", ()), {"pictures": self.pictures})

    def read_part(self, name: str) -> bytes:
        return self._read_part(self.z.getinfo(name))

    def _extras_doc(self):
        """埋め込み・グラフ・SmartArt のあるブロックだけの小さな文書（番号は、全体と同じ順に出てくる）。"""
        from ..docx.document import DocxDocument

        parts = [self.head_xml]
        for bi in self.objects_blocks:
            a = self.offs[bi]
            b = self.offs[bi + 1] if bi + 1 < len(self.offs) else self.data_end
            parts.append(self.reader.read(a, b))
        parts.append(self.tail_xml)
        real = tuple(n for n in self.kept if n.lower().startswith(("word/embeddings/", "word/charts/", "word/diagrams/"))
                     and self.z.getinfo(n).file_size <= limits.MAX_INFLATE_PART)
        return DocxDocument(self._mini(b"".join(parts), real), self.name, base={"stories": tuple(self.object_stories), "para": 0})

    def _make_extras(self) -> list[str]:
        from ..docx.window import _embeds, _macros
        from ..docx.window import to_markdown
        from ..embeds.word import objects as _objects
        from ..imaging.docimages import list_images, markdown

        d = self._extras_doc()
        win = {"source": {"file": self.name, "sha256": self.sha256, "bytes": self.size}, "notation": {}, "protection": {"edit": "", "enforced": False},
               "macros": _macros(d), "fields": [], "stories": []}
        win["embeds"] = _embeds(d)
        win["objects"] = _objects(d, shown={e["n"] for e in win["embeds"]})
        text = to_markdown(win)
        k = text.find("\n## ")
        lines = text[k + 1:].split("\n") if k >= 0 else []
        images = list_images("docx", self._mini(b"", ()), {"pictures": self.pictures})
        if images:
            lines += markdown(images).split("\n")
        return lines

    # ------------------------------------------------------------------ 段階
    def grade(self):
        from ..grade import Reason, _final

        reasons: list = []
        write = [Reason("B", "欄を範囲ごとに見つける", "大きな文書なので、欄は頼まれた段落の範囲ごとに見つけます")]
        if self.has_macros:
            write.append(Reason("C", "マクロ", "マクロがあります。窓の終わりの節のソースを読んで判断します"))
        if self.tracked:
            reasons.append(Reason("B", "変更履歴", f"変更履歴で削除された字が {self.tracked} か所に残っています"))
        if self.tracked_ins:
            from ..grade import inserted_reason

            reasons.append(Reason("B", "変更履歴", inserted_reason(self.tracked_ins, "か所")))
        g = _final(reasons, write)
        g.deferred = ["未確認：全段落の字の読み落とし・全部の欄への書き込みの試し（audit_document で全部のまとまりを確かめます）"]
        return g

    # ------------------------------------------------------------------ 検査
    def audit_begin(self) -> dict:
        return {"next_chunk": 0, "chars": 0, "fields": 0, "written": 0, "problems": [], "write_problems": [], "anchored": None, "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        from ..docx.audit import _count, anchored_check, check_writing, raw_texts
        from ..docx.document import DocxDocument, partial_base
        from ..docx.fill import fill
        from ..docx.proof import prove
        from ..docx.window import build

        i = state["next_chunk"]
        n = len(self.chunks)
        if i < n:
            with partial_base(self.base_of(i)):  # 書き込みの試しが開き直す文書も、番号は全体と同じ
                d = self.told_doc(i)
                win = build(d)
                raw = _count(raw_texts(d))
                shown = None
                from collections import Counter

                shown = Counter()
                for st in win["stories"]:
                    shown += _count(st["_plain"])
                state["chars"] += sum(raw.values())
                state["fields"] += len(win["fields"])
                for ch, k in list((raw - shown).items())[:5]:
                    if len(state["problems"]) < 30:
                        state["problems"].append(f"段落 {self.chunks[i]['para'] + 1}〜：「{ch}」が {k} 字、窓に出ていません")
                for ch, k in list((shown - raw).items())[:5]:
                    if len(state["problems"]) < 30:
                        state["problems"].append(f"段落 {self.chunks[i]['para'] + 1}〜：「{ch}」が {k} 字、元にないのに窓に出ています")
                # 書き込みの試し：このまとまりの欄すべて（小さな文書の上で）
                if win["fields"]:
                    for c in check_writing(self.told_doc(i), win):
                        state["written"] += 1 if c.ok else 0
                        if not c.ok and len(state["write_problems"]) < 30:
                            state["write_problems"] += [f"段落 {self.chunks[i]['para'] + 1}〜：{c.name}：{c.detail[:120]}"] + [str(x)[:160] for x in list(c.items)[:3]]
                if state["anchored"] is None and not (self.protection_enforced and self.protection) and "story" not in self.chunks[i]:
                    blob = d.blob
                    c = anchored_check(blob, self.name, DocxDocument, fill, prove)
                    state["anchored"] = {"name": c.name, "ok": c.ok, "agent": False, "detail": c.detail, "items": list(c.items)}
                del d, win
                state["next_chunk"] = i + 1
                limits.progress("まとまりごとに確かめています", i + 1, n)
                return state
        checks = [{"name": "読み落とし（文書の文字）", "ok": not state["problems"], "agent": False, "items": list(state["problems"]),
                   "detail": f"全 {self.paragraphs:,} 段落・字 {state['chars']:,} 字を、窓とは別の読み方（XML の文字の要素を直接）で拾い、まとまり {n} 個ごとに窓と比べました。"
                             f"見つけた欄 {state['fields']:,} か所"},
                  {"name": "窓の大きさ", "ok": True, "agent": False, "items": [],
                   "detail": f"窓の全体は見積もりで {self.est_chars:,} 字。段落の範囲で分けて読めます"},
                  {"name": "全入力欄への書き込み", "ok": not state["write_problems"], "agent": False, "items": list(state["write_problems"]),
                   "detail": f"全 {state['fields']:,} か所の欄に、まとまりごとの小さな文書の上で試しの値を書き、証明と読み戻しを確かめました（全体のファイルへの書き込みは fill_document のときに証明します）"}]
        if state["anchored"]:
            checks.append(state["anchored"])
        from ..docx.audit import macro_check
        from ..xlsx.audit import Check  # noqa: F401

        ex = self._extras_doc()
        from ..docx.window import _macros

        mc = macro_check(ex, {"macros": _macros(ex)})
        checks.append({"name": mc.name, "ok": mc.ok, "agent": mc.agent, "detail": mc.detail, "items": list(mc.items)})
        if self.objects_blocks:
            from ..docx.window import _embeds
            from ..embeds.audit import word_checks
            from ..embeds.word import objects as _objects

            w = {"embeds": _embeds(ex)}
            w["objects"] = _objects(ex, shown={e["n"] for e in w["embeds"]})
            for c in word_checks(ex, w):
                checks.append({"name": c.name, "ok": c.ok, "agent": c.agent, "detail": c.detail, "items": list(c.items)})
        state["checks"] = checks
        state["done"] = True
        return state


def _tag_end(buf: bytearray, pos: int) -> int:
    """閉じタグ・空要素のタグの終わりの次（属性の値の中の > を飛ばす）。"""
    q = None
    n = len(buf)
    i = pos
    while i < n:
        c = buf[i]
        if q is not None:
            if c == q:
                q = None
        elif c in (0x22, 0x27):
            q = c
        elif c == 0x3E:
            return i + 1
        i += 1
    return n

