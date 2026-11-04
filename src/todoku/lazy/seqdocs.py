"""単位ごとに小さく読む文書：.doc（ピースの表と段落の範囲）・HWP 5（節とレコード）・.odt・.hwpx（本文の直下の要素の位置）。"""

from __future__ import annotations

import io
import re
import struct
import xml.parsers.expat as expat
import zipfile
from array import array

from .. import limits
from .blocks import SeqLazy
from .cfbfile import CfbFile, SpoolCfb
from .textstream import Source
from .xlsxdoc import Member, SeekableInflate, _no_doctype

UNIT_CHARS = 1 << 16  # .doc：1 単位の字の数の目安
UNIT_BYTES = 1 << 20  # XML・レコード：1 単位のバイトの目安


def _stubs(items: list) -> list:
    return [{**{k: v for k, v in e.items() if k not in ("para",)}, "stub": True} for e in items]


# ====================================================================== .doc
class DocLazy(SeqLazy):
    """Word 97〜2003：字はピースの表から（メモリの字列）、段落は頼まれた字の範囲だけを作る。WordDocument などのストリームは写像で読む。"""

    kind = "doc"

    def __init__(self, path: str, name: str, cf: CfbFile):
        from ..doc.book import DocFile
        from ..docx.window import build, markdown_parts

        self.path, self.name, self.cf = path, name, cf
        self.size = cf.size
        self.spool = SpoolCfb(cf)
        self.window_fn = build
        self.md_fn = markdown_parts
        self.skel = DocFile(cf.fb, name, cfb=self.spool, partial={"range": ("body", 0, 0)})
        self.units = self._units()
        self.index_units()

    def _units(self) -> list[tuple]:
        """本文などを、表の外の段落の終わりで切る。"""
        d = self.skel
        c = d.ccp
        out = []
        at = 0
        for kind, key in (("body", "text"), ("footnotes", "ftn"), ("header", "hdd"), ("comments", "atn"),
                          ("endnotes", "edn"), ("textbox", "txbx"), ("header-textbox", "hdrtxbx")):
            n = c[key]
            if n <= 0:
                continue
            a, b = at, at + n
            at = b
            start = a
            for m in re.finditer("[\r\x07]", d.text[a:b]):
                e = a + m.start()
                if e + 1 - start < UNIT_CHARS or e + 1 >= b:
                    continue
                if d.para_props(e)["in_table"]:
                    continue
                out.append((kind, start, e + 1))
                start = e + 1
            if start < b or kind == "body":
                out.append((kind, start, b))
        return out or [("body", 0, 0)]

    def first_state(self) -> dict:
        return {"embeds": [], "pictures": 0}

    def unit_count(self) -> int:
        return len(self.units)

    def parse_unit(self, u: int, state: dict):
        from ..doc.book import DocFile

        return DocFile(self.cf.fb, self.name, cfb=self.spool, partial={"range": self.units[u], **state})

    def state_after(self, doc, state: dict) -> dict:
        return {"embeds": _stubs(doc.embeds), "pictures": len(doc.pictures)}

    def unit_has_objects(self, doc) -> bool:
        return any(not e.get("stub") for e in doc.embeds)

    def list_images(self) -> list[dict]:
        from ..imaging.docimages import list_images

        return list_images("doc", self.cf.fb, {})


# ====================================================================== HWP 5
class _StreamMember:
    """SeekableInflate に渡す、圧縮したストリーム（Member と同じ問い合わせ）。"""

    class _Info:
        compress_type = 8

    def __init__(self, st, compressed: bool):
        self.src = st
        self.name = st.name
        self.compressed = st.size
        self.size = 0
        self.data_at = 0
        self.info = self._Info()
        self.info.compress_type = 8 if compressed else 0

    def chunks(self):
        yield from self.src.chunks()


class HwpLazy(SeqLazy):
    kind = "hwp"

    def __init__(self, path: str, name: str, cf: CfbFile):
        from ..hwpx.hwp5 import Hwp5Document
        from ..hwpx.window import build5, markdown_parts

        self.path, self.name, self.cf = path, name, cf
        self.size = cf.size
        self.window_fn = build5
        self.md_fn = markdown_parts
        self.skel = Hwp5Document(cf.fb, name, partial={"cfb": cf.cfb, "records": None})
        self.readers = {p: SeekableInflate(_StreamMember(cf.stream(p), self.skel.compressed)) for p in self.skel.parts}
        self.units: list[tuple[str, int, int]] = []
        for part in self.skel.parts:
            self._index_section(part)
        self.index_units()

    def _index_section(self, part: str) -> None:
        """レコードの頭だけをたどり、いちばん上の段落の切れ目で単位に分ける。"""
        from ..hwpx.hwp5 import PARA_HEADER

        reader = self.readers[part]
        carry = b""
        at = 0  # carry の頭の位置
        start = 0
        pos = 0
        for c in reader.scan():
            carry += c
            while True:
                k = pos - at
                if k + 4 > len(carry):
                    break
                h = struct.unpack_from("<I", carry, k)[0]
                size = h >> 20
                head = 4
                if size == 0xFFF:
                    if k + 8 > len(carry):
                        break
                    size = struct.unpack_from("<I", carry, k + 4)[0]
                    head = 8
                tag, level = h & 0x3FF, (h >> 10) & 0x3FF
                if level == 0 and tag == PARA_HEADER and pos - start >= UNIT_BYTES:
                    self.units.append((part, start, pos))
                    start = pos
                pos += head + size
                if pos > at + len(carry):
                    break
            if pos > at:
                drop = min(pos - at, len(carry))
                carry = carry[drop:]
                at += drop
            limits.progress(f"{part} の段落の位置を調べています", pos, 0)
        self.units.append((part, start, pos))

    def first_state(self) -> dict:
        return {"pictures": 0}

    def unit_count(self) -> int:
        return len(self.units)

    def parse_unit(self, u: int, state: dict):
        from ..hwpx.hwp5 import Hwp5Document

        part, a, b = self.units[u]
        data = self.readers[part].read(a, b)
        return Hwp5Document(self.cf.fb, self.name, partial={"cfb": self.cf.cfb, "records": data, "part": part, **state})

    def state_after(self, doc, state: dict) -> dict:
        return {"pictures": len(doc.pictures)}

    def story_key(self, kind: str) -> int:
        return 0 if kind in ("header", "footer") else (2 if kind == "comments" else 1)

    def list_images(self) -> list[dict]:
        from ..imaging.docimages import list_images

        return list_images("hwp", self.cf.fb, {})


# ====================================================================== 本文の直下の要素で分ける XML（.odt・.hwpx）
class XmlSplit:
    """XML の部品を流し読みして、ある深さの要素の位置の索引を作る。head：その要素の前まで（keep の要素も含める）。"""

    def __init__(self, member: Member, depth: int, container: str | None, keep: frozenset[str] = frozenset(), chars: bool = True):
        self.m = member
        self.reader = SeekableInflate(member)
        self.offs = array("Q")
        self.end = 0
        self.head = b""
        p = expat.ParserCreate()
        p.buffer_text = True
        p.StartDoctypeDeclHandler = _no_doctype
        st = {"depth": 0, "first": None, "inside": container is None, "open_at": None}
        kept = bytearray()

        def start(name, attrs):
            d = st["depth"]
            st["depth"] = d + 1
            local = name.rsplit(":", 1)[-1]
            if d == depth - 1 and container is not None and local == container:
                st["inside"] = True
                return
            if d == depth and st["inside"]:
                if st["first"] is None and local in keep:
                    return
                if st["first"] is None:
                    st["first"] = p.CurrentByteIndex
                self.offs.append(p.CurrentByteIndex)

        def end(name):
            st["depth"] -= 1
            if st["depth"] == depth and st["inside"]:
                pass
            if st["depth"] == depth - 1 and st["inside"]:
                if self.end == 0:
                    self.end = p.CurrentByteIndex
                if container is not None:
                    st["inside"] = False

        p.StartElementHandler = start
        p.EndElementHandler = end
        total = 0
        size = member.size or 1
        for c in self.reader.scan():
            if st["first"] is None:
                kept += c
            total += len(c)
            p.Parse(c, False)
            limits.progress(f"{member.name} の位置を調べています", total, size)
        p.Parse(b"", True)
        self.total = total
        if not self.end:
            self.end = total
        first = st["first"] if st["first"] is not None else self.end
        self.head = bytes(kept[:first]) if len(kept) >= first else self.reader.read(0, first)
        self.tail = self.reader.read(self.end, total)

    def groups(self, limit: int = UNIT_BYTES) -> list[tuple[int, int]]:
        """[(要素の番号の始め, 終わり)]。"""
        out = []
        i = 0
        n = len(self.offs)
        while i < n:
            j = i + 1
            while j < n and self.offs[j] - self.offs[i] < limit:
                j += 1
            out.append((i, j))
            i = j
        return out or [(0, 0)]

    def mini(self, i: int, j: int) -> bytes:
        if i >= j:
            return self.head + self.tail
        a = self.offs[i]
        b = self.offs[j] if j < len(self.offs) else self.end
        return self.head + self.reader.read(a, b) + self.tail


class _PackagedLazy(SeqLazy):
    """ZIP の文書：小さな ZIP（部品は写し、画像は名前だけ）を作って読む。"""

    def _open_zip(self, path: str, name: str) -> None:
        self.path, self.name = path, name
        self.src = Source(path)
        self.size = self.src.size
        self.z = zipfile.ZipFile(path)
        self.kept: dict[str, bytes] = {}

    def _keep_parts(self, skip: set[str], blank_prefixes: tuple) -> None:
        for i in self.z.infolist():
            n = i.filename
            if n in skip or n.endswith("/"):
                continue
            if n.lower().startswith(blank_prefixes) or i.file_size > 1 << 24:
                self.kept[n] = b""
            else:
                self.kept[n] = Member(self.src, i).read()

    def _mini_zip(self, replace: dict[str, bytes]) -> bytes:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_STORED) as z:
            for i in self.z.infolist():
                n = i.filename
                if n in replace:
                    z.writestr(n, replace[n])
                elif n in self.kept:
                    z.writestr(n, self.kept[n])
                else:
                    z.writestr(n, b"")
        return buf.getvalue()

    def close(self) -> None:
        self.z.close()
        self.src.close()

    def unit_blob(self, u: int) -> bytes:
        raise NotImplementedError

    def list_images(self) -> list[dict]:
        """単位ごとに一覧を作ってつなぐ（図の番号は全体の順に直す）。"""
        from ..imaging.docimages import list_images

        out: list[dict] = []
        rest: dict[str, dict] = {}
        n = 0
        for u in range(self.unit_count()):
            for im in list_images(self.kind, self.unit_blob(u), {}):
                if str(im.get("id", "")).startswith("図"):
                    n += 1
                    k = im["id"][1:]
                    out.append({**im, "id": f"図{n}", "where": str(im.get("where", "")).replace(f"の {k} 番目", f"の {n} 番目")})
                else:
                    rest.setdefault(im["id"], im)
        used = {im.get("part") for im in out}
        return out + [im for i, im in rest.items() if i not in used]


_ODT_KEEP = frozenset(("forms", "tracked-changes", "variable-decls", "sequence-decls", "user-field-decls", "dde-connection-decls"))


class OdtLazy(_PackagedLazy):
    kind = "odt"

    def __init__(self, path: str, name: str):
        from ..odf.textwindow import build, markdown_parts

        self._open_zip(path, name)
        self.window_fn = build
        self.md_fn = markdown_parts
        self.part = "content.xml"
        self.split = XmlSplit(Member(self.src, self.z.getinfo("content.xml")), 3, "text", _ODT_KEEP)
        self.unit_list = self.split.groups()
        self._keep_parts({"content.xml"}, ("pictures/",))
        self.index_units()

    def first_state(self) -> dict:
        return {"list_counters": {}, "outline_counter": [0] * 11, "pictures": 0}

    def unit_count(self) -> int:
        return len(self.unit_list)

    def parse_unit(self, u: int, state: dict):
        from ..odf.text import OdtDocument

        i, j = self.unit_list[u]
        return OdtDocument(self._mini_zip({"content.xml": self.split.mini(i, j)}), self.name, partial={**state, "master": u == 0})

    def unit_blob(self, u: int) -> bytes:
        i, j = self.unit_list[u]
        return self._mini_zip({"content.xml": self.split.mini(i, j)})

    def state_after(self, doc, state: dict) -> dict:
        return {"list_counters": {k: list(v) for k, v in doc.list_counters.items()}, "outline_counter": list(doc.outline_counter), "pictures": len(doc.pictures)}

    def unit_has_objects(self, doc) -> bool:
        return False

    def story_key(self, kind: str) -> int:
        return 0 if kind in ("header", "footer") else (2 if kind == "comments" else 1)


class HwpxLazy(_PackagedLazy):
    kind = "hwpx"

    def __init__(self, path: str, name: str):
        from ..hwpx.document import HwpxDocument
        from ..hwpx.window import build, markdown_parts

        self._open_zip(path, name)
        self.window_fn = build
        self.md_fn = markdown_parts
        self._keep_parts(set(), ("bindata/",))
        names = [i.filename for i in self.z.infolist()]
        sections = sorted((n for n in names if re.fullmatch(r"Contents/section\d+\.xml", n, re.I)),
                          key=lambda n: int(re.search(r"(\d+)\.xml$", n, re.I)[1]))
        for s in sections:
            self.kept[s] = b""
        self.splits = {s: XmlSplit(Member(self.src, self.z.getinfo(s)), 1, None) for s in sections}
        self.unit_list = [(s, i, j) for s in sections for i, j in self.splits[s].groups()]
        self.empty = {s: sp.head + sp.tail for s, sp in self.splits.items()}
        self.has_master = any(n.startswith("Contents/masterpage") for n in names)
        self._cls = HwpxDocument
        self.index_units()

    def first_state(self) -> dict:
        return {"num_counters": {}, "outline_id": "", "pictures": 0}

    def unit_count(self) -> int:
        return len(self.unit_list)

    def parse_unit(self, u: int, state: dict):
        s, i, j = self.unit_list[u]
        rep = dict(self.empty)
        rep[s] = self.splits[s].mini(i, j)
        return self._cls(self._mini_zip(rep), self.name, partial={**state, "section": s, "master": u == len(self.unit_list) - 1})

    def unit_blob(self, u: int) -> bytes:
        s, i, j = self.unit_list[u]
        rep = dict(self.empty)
        rep[s] = self.splits[s].mini(i, j)
        return self._mini_zip(rep)

    def state_after(self, doc, state: dict) -> dict:
        return {"num_counters": {k: list(v) for k, v in doc.num_counters.items()}, "outline_id": doc.outline_id, "pictures": len(doc.pictures)}

    def unit_has_objects(self, doc) -> bool:
        return False

    def story_key(self, kind: str) -> int:
        return {"header": 0, "footer": 0, "footnotes": 1, "endnotes": 1, "comments": 2, "master": 3}.get(kind, 1)


# ====================================================================== PowerPoint
class PptxLazy(_PackagedLazy):
    """スライドごと：頼まれたスライドとそのノートだけを小さな ZIP に写して読む。"""

    kind = "pptx"

    def __init__(self, path: str, name: str):
        from ..docx.document import iter_paras
        from ..docx.window import _tables
        from ..pptx.document import PptxDocument

        self._open_zip(path, name)
        self._keep_parts(set(), ("ppt/media/", "ppt/embeddings/", "ppt/slides/slide", "ppt/notesslides/notesslide"))
        self.chunks = []
        self.templates: dict[str, list[str]] = {}
        self.unread_notes: list[str] = []
        self.pictures_seen: list[dict] = []
        skel = PptxDocument(self._mini_zip({}), name, partial={"slides": ()})
        order = list(skel.slide_parts)
        para = table = chars = 0
        pics = charts = dgm = 0
        for n, part in enumerate(order, 1):
            limits.progress("スライドを調べています", n, len(order))
            self.chunks.append({"para": para, "table": table, "n": n, "part": part, "kind": None, "state": (pics, charts, dgm)})
            d = self._open_slide(n, part, pics, charts, dgm)
            for _k, _p, blocks in d.all_blocks():
                para += sum(1 for _ in iter_paras(blocks))
                table += sum(1 for _ in _tables(blocks))
                chars += sum(len(p.text) for p in iter_paras(blocks))
            pics, charts, dgm = len(d.pictures), len(d.charts), len(d.diagrams)
            self.pictures_seen += [p for p in d.pictures if not p.get("stub")]
            self._remember_templates(d)
            self.unread_notes += [x for x in d.unread if x not in self.unread_notes]
        self.paragraphs, self.tables = para, table
        self.est_chars = int(chars * 1.25) + para * 12

    def _open_slide(self, n: int, part: str, pics: int, charts: int, dgm: int):
        from ..pptx.document import PptxDocument

        real = {part: Member(self.src, self.z.getinfo(part)).read()}
        rp = part.replace("slides/slide", "slides/_rels/slide") + ".rels"
        if rp in self.kept:
            for m in re.finditer(rb'Target="([^"]*notesSlide[^"]*)"', self.kept[rp]):
                from ..docx.document import _join

                nn = _join(part, m[1].decode())
                if nn in self.z.namelist():
                    real[nn] = Member(self.src, self.z.getinfo(nn)).read()
        return PptxDocument(self._mini_zip(real), self.name, partial={"slides": (n,), "pictures": pics, "charts": charts, "diagrams": dgm})

    def _remember_templates(self, d) -> None:
        from ..docx.window import Spots, _render_blocks
        from .blocks import SubDoc

        for part, blocks in d.templates:
            if part in self.templates:
                continue
            lines: list[str] = []
            _render_blocks(blocks, Spots(SubDoc(d, [])), {}, lines, [])
            self.templates[part] = lines

    def chunk_doc(self, i: int, real: tuple = ()):
        from .blocks import SubDoc

        ch = self.chunks[i]
        d = self._open_slide(ch["n"], ch["part"], *ch["state"])
        sub = SubDoc(d, list(d.all_blocks()), ch["para"], ch["table"])
        sub.templates = []
        return sub

    def build_window(self, doc) -> dict:
        from ..pptx.window import build

        return build(doc)

    def _make_extras(self) -> list[str]:
        from ..imaging.docimages import list_images, markdown
        from ..pptx.window import TEMPLATE

        lines: list[str] = []
        for part, ls in self.templates.items():
            if ls:
                lines += [f"## {TEMPLATE}（{part}）", ""] + [f"- {x}" for x in ls] + [""]
        pics = [{"id": f"図{p['n']}", "alt": p["alt"] or None, "decorative": p["decorative"],
                 "where": f"スライド {p['slide']}" + (f"の「{p['name']}」" if p["name"] else ""), "media": p.get("media", ""),
                 **({"size_mm": p["size_mm"]} if p.get("size_mm") else {})} for p in self.pictures_seen]
        images = list_images("pptx", self._mini_zip({}), {"pictures": pics})
        if images:
            lines += markdown(images).split("\n")
        return lines
