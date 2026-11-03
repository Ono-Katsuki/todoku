"""PowerPoint（.pptx・.pptm）を読む（ECMA-376 第1部 19 章 PresentationML・21 章 DrawingML）。外部のライブラリは使わない。

- スライドの順：presentation.xml の p:sldIdLst（r:id から関係でスライドの部品へ）。p:sld の show="0" は非表示のスライド
- スライドの字：p:spTree の図形（p:sp の p:txBody）・まとめた図形（p:grpSp、入れ子も）・表（p:graphicFrame の a:tbl。
  gridSpan・hMerge・rowSpan・vMerge）・AlternateContent（mc:Choice を読み、代わりの mc:Fallback は読まない）
- 段落（a:p）の字：a:r の a:t、a:fld の a:t（スライド番号・日付などの、PowerPoint が作る字）、a:br（行の区切り）
- 絵（p:pic、絵で塗った図形）：説明（cNvPr の descr）と、飾りの印（adec:decorative）。グラフ（c:chart）・SmartArt（dgm:relIds）
- ノート（スライドの関係の notesSlide）。スライドの型（レイアウト・マスター）の、プレースホルダーでない図形の字（スライドに出るもの）
読んだ結果は、Word（.docx）と同じ形（段落・表・マス）にして、窓・入力欄の推定（空のマス・空欄・□・「？」の所）を共通にする。
"""

from __future__ import annotations

import re

from ..docx.document import Cell, Control, Para, Seg, Table, _join
from ..docx.xmlpos import Node, parse
from ..zipraw import RawZip

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
A_T = frozenset({"{" + A + "}t"})
_STRICT_A = "http://purl.oclc.org/ooxml/drawingml/main"


class PptxError(ValueError):
    pass


def rels_of(z: RawZip, part: str) -> dict[str, tuple[str, str]]:
    """部品の関係：r:id → (種類, 先)。"""
    d, base = part.rsplit("/", 1) if "/" in part else ("", part)
    rp = f"{d}/_rels/{base}.rels" if d else f"_rels/{base}.rels"
    out = {}
    if z.has(rp):
        for m in re.finditer(rb"<(?:\w+:)?Relationship\b([^>]*)/?>", z.read(rp)):
            a = dict(re.findall(r'(\w+)="([^"]*)"', m[1].decode("utf-8", "replace")))
            if "Id" in a and a.get("TargetMode") != "External":
                out[a["Id"]] = (a.get("Type", ""), a.get("Target", ""))
    return out


def _rid(n: Node, local: str = "id") -> str:
    return n.attrs.get("{" + R + "}" + local, "")


def _on(v: str | None) -> bool:
    return v in ("1", "true", "on")


class PptxDocument:
    """PowerPoint の文書。Word の文書と同じ問い合わせ（all_blocks・controls・marks など）に答える。"""

    kind = "pptx"

    def __init__(self, blob: bytes, filename: str = ""):
        self.blob = blob
        self.filename = filename
        try:
            self.zip = RawZip(blob)
        except Exception as e:  # noqa: BLE001
            raise PptxError(f"ZIP として読めません：{e}") from None
        self.parts: dict[str, tuple[bytes, Node]] = {}
        self.pictures: list[dict] = []  # {"n", "alt", "decorative", "slide", "name"}
        self.charts: list[dict] = []
        self.diagrams: list[dict] = []
        self.unread: list[str] = []
        self.controls: list[Control] = []
        self.protection = ""
        self.protection_enforced = False
        self.has_macros = any(n.lower().endswith("vbaproject.bin") for n in self.zip.names())
        self.main = self._main_part()
        self.slides: list[dict] = []  # {"n", "part", "hidden", "notes", "blocks", "notes_blocks"}
        self.templates: list[tuple[str, list]] = []  # (レイアウト・マスターの部品, ブロック)
        self._slide = 0
        self._read()

    # ------------------------------------------------------------------ 部品
    @property
    def vba(self):
        if getattr(self, "_vba", None) is None and self.has_macros:
            from ..vba import read_project

            try:
                name = next(n for n in self.zip.names() if n.lower().endswith("vbaproject.bin"))
                self._vba = read_project(self.zip.read(name))
            except Exception as e:  # noqa: BLE001
                self._vba = e
        v = getattr(self, "_vba", None)
        return v if v is not None and not isinstance(v, Exception) else None

    def _main_part(self) -> str:
        rels = self.zip.read("_rels/.rels") if self.zip.has("_rels/.rels") else b""
        m = re.search(rb'Target="/?([^"]*presentation[^"]*\.xml)"', rels)
        name = m[1].decode() if m else "ppt/presentation.xml"
        if not self.zip.has(name):
            raise PptxError("PowerPoint の文書の本体（presentation.xml）がありません")
        return name

    def load(self, name: str) -> Node | None:
        if name in self.parts:
            return self.parts[name][1]
        if not self.zip.has(name):
            return None
        xml = self.zip.read(name)
        if xml.startswith((b"\xff\xfe", b"\xfe\xff")):
            raise PptxError(f"{name}：UTF-16 の XML は読みません")
        try:
            root = parse(xml, A_T | frozenset({"{" + _STRICT_A + "}t"}))
        except Exception as e:  # noqa: BLE001  expat の誤り
            raise PptxError(f"{name}：XML として読めません（{e}）") from None
        self.parts[name] = (xml, root)
        return root

    def xml(self, part: str) -> bytes:
        return self.parts[part][0]

    # ------------------------------------------------------------------ スライド
    def _read(self) -> None:
        root = self.load(self.main)
        pres = root.kids[0] if root is not None and root.kids else None
        lst = pres.find("sldIdLst") if pres is not None else None
        rels = rels_of(self.zip, self.main)
        shown_templates: list[str] = []
        for sid in (lst.findall("sldId") if lst is not None else []):
            typ, target = rels.get(_rid(sid), ("", ""))
            part = _join(self.main, target) if target else ""
            if not part or not self.zip.has(part):
                self.unread.append(f"スライドの部品（{target or _rid(sid)}）がありません")
                continue
            self._slide += 1
            sroot = self.load(part)
            sld = sroot.kids[0] if sroot.kids else None
            if sld is None:
                continue
            srels = rels_of(self.zip, part)
            s = {"n": self._slide, "part": part, "hidden": sld.get("show") in ("0", "false"), "notes": None,
                 "blocks": self._csld(sld, part, srels), "notes_blocks": []}
            for _rid_, (t, tg) in srels.items():
                kind = t.rsplit("/", 1)[-1]
                name = _join(part, tg)
                if kind == "notesSlide" and self.zip.has(name):
                    nroot = self.load(name)
                    if nroot is not None and nroot.kids:
                        s["notes"] = name
                        s["notes_blocks"] = self._csld(nroot.kids[0], name, rels_of(self.zip, name))
                elif kind == "slideLayout" and self.zip.has(name) and sld.get("showMasterSp") not in ("0", "false"):
                    shown_templates.append(name)
                    lroot = self.load(name)
                    lay = lroot.kids[0] if lroot is not None and lroot.kids else None
                    if lay is not None and lay.get("showMasterSp") not in ("0", "false"):
                        for t2, tg2 in rels_of(self.zip, name).values():
                            if t2.endswith("/slideMaster") and self.zip.has(_join(name, tg2)):
                                shown_templates.append(_join(name, tg2))
            self.slides.append(s)
        for name in dict.fromkeys(shown_templates):
            troot = self.load(name)
            if troot is None or not troot.kids:
                continue
            blocks = self._csld(troot.kids[0], name, rels_of(self.zip, name), template=True)
            if blocks:
                self.templates.append((name, blocks))

    def _csld(self, top: Node, part: str, rels: dict, template: bool = False) -> list:
        csld = top.find("cSld")
        tree = csld.find("spTree") if csld is not None else None
        if tree is None:
            return []
        return self._tree(tree, part, rels, template)

    def _tree(self, node: Node, part: str, rels: dict, template: bool) -> list:
        out: list = []
        for ch in node.kids:
            loc = ch.local
            if loc == "AlternateContent":
                choice = ch.find("Choice")  # 代わりの中身（Fallback）は、同じものの別の書き方なので読まない
                if choice is not None:
                    out += self._tree(choice, part, rels, template)
            elif loc == "grpSp":
                out += self._tree(ch, part, rels, template)
            elif loc == "sp":
                if template and _placeholder(ch):
                    continue  # スライドの型のプレースホルダー：スライドの側に字がなければ、スライドには出ない（見本の字）
                body = ch.find("txBody")
                if body is not None:
                    paras = [self._para(p, part, None) for p in body.findall("p")]
                    if _placeholder_type(ch) in ("title", "ctrTitle", "subTitle"):
                        for p_ in paras:
                            p_.style = "Title"  # タイトルのプレースホルダー（PresentationML の ph の type）：見出しの段落
                    out += paras
                sppr = ch.find("spPr")
                if not template and sppr is not None and sppr.find("blipFill") is not None:
                    out.append(self._picture(ch, part, "nvSpPr"))
            elif loc == "pic":
                if not template:
                    out.append(self._picture(ch, part, "nvPicPr"))
            elif loc == "graphicFrame" and not template:
                out += self._frame(ch, part, rels)
        return out

    def _picture(self, node: Node, part: str, nv: str) -> Para:
        pr = node.find(nv)
        c = pr.find("cNvPr") if pr is not None else None
        alt = ((c.get("descr") or "") if c is not None else "").strip()
        deco = c is not None and any(_on(x.get("val")) for x in c.iter("decorative"))
        blip = next(iter(node.iter("blip")), None)
        rid = (blip.get("embed") or "") if blip is not None else ""
        target = rels_of(self.zip, part).get(rid, ("", ""))[1] if rid else ""
        media = _join(part, target) if target else ""
        pic = {"n": len(self.pictures) + 1, "alt": alt, "decorative": deco, "slide": self._slide,
               "name": (c.get("name") or "") if c is not None else "", "media": media if self.zip.has(media) else ""}
        self.pictures.append(pic)
        mark = f"〔図: {alt}〕" if alt else ("" if deco else f"〔図{pic['n']}（説明なし）〕")
        pic["para"] = Para(node, part, [Seg(mark, "other", None)] if mark else [])
        return pic["para"]

    def _frame(self, ch: Node, part: str, rels: dict) -> list:
        g = ch.find("graphic")
        data = g.find("graphicData") if g is not None else None
        if data is None:
            return []
        nv = ch.find("nvGraphicFramePr")
        c = nv.find("cNvPr") if nv is not None else None
        alt = (c.get("descr") or "") if c is not None else ""
        out: list = []
        for k in data.kids:
            if k.local == "tbl":
                out.append(self._table(k, part))
            elif k.local in ("chart", "relIds"):
                rid = _rid(k) if k.local == "chart" else _rid(k, "dm")
                target = _join(part, rels[rid][1]) if rid in rels else ""
                if not target or not self.zip.has(target):
                    continue
                if k.local == "chart":
                    self.charts.append({"n": len(self.charts) + 1, "part": target, "alt": alt})
                    out.append(Para(k, part, [Seg(f"〔グラフ{len(self.charts)}〕", "other", None)]))
                else:
                    self.diagrams.append({"n": len(self.diagrams) + 1, "part": target, "alt": alt})
                    out.append(Para(k, part, [Seg(f"〔SmartArt{len(self.diagrams)}〕", "other", None)]))
            elif k.local == "oleObj" or any(x.local == "oleObj" for x in k.iter()):
                ole = next(x for x in k.iter() if x.local == "oleObj")
                prog = ole.get("progId") or "OLE"
                self.unread.append(f"スライド {self._slide} の埋め込み（{prog}）の中身は読んでいません")
                out.append(Para(k, part, [Seg(f"〔埋め込み：{prog}（中身は読んでいません）〕", "other", None)]))
        return out

    def _table(self, tbl: Node, part: str) -> Table:
        t = Table(tbl, part)
        for r, tr in enumerate(tbl.findall("tr")):
            row: list[Cell] = []
            for col, tc in enumerate(tr.findall("tc")):
                if _on(tc.get("hMerge")):
                    continue  # 左のマスの gridSpan に含まれるマス（字は左のマスの側にある）
                vm = "continue" if _on(tc.get("vMerge")) else ("restart" if int(tc.get("rowSpan") or 1) > 1 else "")
                cell = Cell(tc, r, col, int(tc.get("gridSpan") or 1), vm, [], t)
                body = tc.find("txBody")
                if body is not None:
                    cell.blocks = [self._para(p, part, cell) for p in body.findall("p")]
                row.append(cell)
            t.rows.append(row)
        return t

    def _para(self, p: Node, part: str, cell: Cell | None) -> Para:
        out = Para(p, part, [], cell=cell)
        out.end_rpr = p.find("endParaRPr")
        for k in p.kids:
            if k.local == "r":
                t = k.find("t")
                rpr = k.find("rPr")
                ul = rpr is not None and (rpr.get("u") or "none") != "none"
                if t is not None:
                    out.segs.append(Seg(t.text, "t", t, k, False, ul))
            elif k.local == "fld":
                # PowerPoint が作る字（スライド番号・日付など）：字として読むが、書き換えない（欄の中の字として扱う）
                t = k.find("t")
                if t is not None:
                    out.segs.append(Seg(t.text, "t", t, k, False, False, Control("text", "pptx-field", "", k, part)))
            elif k.local == "br":
                out.segs.append(Seg("\n", "br", k, None))
        return out

    # ------------------------------------------------------------------ Word と同じ問い合わせ
    def all_blocks(self):
        """(種類, 部品, ブロック)：スライドの順に、スライドとそのノート。種類は窓の節の名前にそのまま使う。"""
        for s in self.slides:
            yield f"スライド {s['n']}" + ("（非表示のスライド）" if s["hidden"] else ""), s["part"], s["blocks"]
            if s["notes"] and s["notes_blocks"]:
                yield f"スライド {s['n']} のノート", s["notes"], s["notes_blocks"]

    def marks(self) -> dict:
        """書き込んだ欄の印：PowerPoint の字には、Word のしおりのような見えない印がない（書いた所は、証明で読み直して確かめる）。"""
        return {}


def _placeholder_type(sp: Node) -> str:
    """プレースホルダーの種類（p:ph の type。省くと body）。プレースホルダーでなければ ""。"""
    nv = sp.find("nvSpPr")
    pr = nv.find("nvPr") if nv is not None else None
    ph = pr.find("ph") if pr is not None else None
    return "" if ph is None else (ph.get("type") or "body")


def _placeholder(sp: Node) -> bool:
    nv = sp.find("nvSpPr")
    pr = nv.find("nvPr") if nv is not None else None
    return pr is not None and pr.find("ph") is not None
