"""Word（.docx）の文書を読む：本文・表・ヘッダー/フッター・脚注・コメント・テキストボックス・段落番号・
旧来のフォームの欄（FORMTEXT など）・コンテンツコントロール・文書の保護。

段落の文字は、どの要素（w:t など）のどのバイト位置から来たかを覚えておき、書き込みを小さな差分にする。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..zipraw import RawZip
from .xmlpos import STRICT, W, Node, parse

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
MATH = "http://schemas.openxmlformats.org/officeDocument/2006/math"

WINGDINGS = {0xF06F: "□", 0xF0A8: "□", 0xF071: "❑", 0xF072: "❒", 0xF06E: "■", 0xF0FE: "☑", 0xF0FD: "☒", 0xF078: "☒",
             0xF0FC: "✓", 0xF0FB: "✗", 0xF06C: "●", 0xF0A1: "○", 0xF09F: "•", 0xF0A7: "▪", 0xF0D8: "➢", 0xF0E8: "➔",
             0xF0B2: "◆", 0xF075: "◆", 0xF0B7: "•"}
SYMBOL_FONT = {0xF0B7: "•", 0xF0A7: "§", 0xF0D7: "×", 0xF0B4: "×", 0xF0AE: "→", 0xF0AC: "←", 0xF0B0: "°", 0xF0B1: "±"}


class DocxError(ValueError):
    pass


@dataclass
class Seg:
    """段落の文字の1片。node は書き換える要素（w:t・w:sym など）。"""

    text: str
    kind: str  # t / tab / br / sym / field / note / other
    node: Node | None = None
    run: Node | None = None
    hidden: bool = False
    underline: bool = False
    field: "Control | None" = None


@dataclass
class Para:
    node: Node
    part: str
    segs: list[Seg]
    style: str = ""
    label: str = ""  # 段落番号（「１．」「(1)」「①」など）
    cell: "Cell | None" = None
    boxes: list["Block"] = field(default_factory=list)  # 段落の中のテキストボックスの中身

    @property
    def text(self) -> str:
        return "".join(s.text for s in self.segs if not s.hidden)

    def raw_text(self) -> str:
        """XML の文字の要素から来た字だけ（非表示の字も含む。窓の読み落としの検査に使う）。"""
        return "".join(s.text for s in self.segs if s.kind in ("t", "sym", "math", "ruby"))

    def hidden_text(self) -> str:
        return "".join(s.text for s in self.segs if s.hidden and s.kind in ("t", "sym"))

    def positions(self) -> list[tuple[int, int]]:
        """text の1字ごとに、(seg の番号, seg の中の位置)。"""
        out = []
        for i, s in enumerate(self.segs):
            if s.hidden:
                continue
            for k in range(len(s.text)):
                out.append((i, k))
        return out


@dataclass
class Cell:
    node: Node
    row: int
    col: int
    span: int = 1
    vmerge: str = ""  # "restart" / "continue" / ""
    blocks: list = field(default_factory=list)
    table: "Table | None" = None
    shade: str = ""

    def paras(self) -> list[Para]:
        out = []
        for b in self.blocks:
            if isinstance(b, Para):
                out.append(b)
            elif isinstance(b, Table):
                for row in b.rows:
                    for c in row:
                        out.extend(c.paras())
        return out

    def text(self) -> str:
        return " ".join(p.text.strip() for p in self.paras() if p.text.strip())


@dataclass
class Table:
    node: Node
    part: str
    rows: list[list[Cell]] = field(default_factory=list)


@dataclass
class Control:
    """入力欄：旧来のフォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN）またはコンテンツコントロール。"""

    kind: str  # text / checkbox / dropdown / date / combo / richtext / picture
    source: str  # formfield / sdt
    name: str
    node: Node  # fldChar（begin）または w:sdt
    part: str
    para: Para | None = None
    value: object = None
    options: list[str] = field(default_factory=list)
    checked: bool = False
    placeholder: bool = False
    maxlen: int | None = None
    date_format: str = ""
    result_segs: list[Seg] = field(default_factory=list)
    end: Node | None = None
    lock: str = ""
    help: str = ""
    ax: object = None  # ActiveX の部品（source が activex のとき、axfile.AxRef）


Block = Para | Table


class Numbering:
    """段落番号（numbering.xml）。"""

    def __init__(self, root: Node | None):
        self.abstract: dict[str, dict[int, dict]] = {}
        self.nums: dict[str, tuple[str, dict[int, dict]]] = {}
        self.counters: dict[str, list[int]] = {}
        if root is None:
            return
        top = root.kids[0] if root.kids else None
        if top is None:
            return
        for an in top.findall("abstractNum"):
            aid = an.get("abstractNumId")
            self.abstract[aid] = {int(lv.get("ilvl", 0)): self._lvl(lv) for lv in an.findall("lvl")}
            link = an.find("numStyleLink")
            if link is not None:
                self.abstract[aid]["_link"] = link.get("val")
        for n in top.findall("num"):
            nid = n.get("numId")
            an = n.find("abstractNumId")
            over = {}
            for lo in n.findall("lvlOverride"):
                il = int(lo.get("ilvl", 0))
                d = {}
                so = lo.find("startOverride")
                if so is not None:
                    d["start"] = int(so.get("val", 1))
                lv = lo.find("lvl")
                if lv is not None:
                    d.update(self._lvl(lv))
                over[il] = d
            self.nums[nid] = (an.get("val") if an is not None else "", over)

    @staticmethod
    def _lvl(lv: Node) -> dict:
        d = {}
        for k in ("start", "numFmt", "lvlText", "lvlRestart"):
            x = lv.find(k)
            if x is not None:
                d[k] = x.get("val")
        if "start" in d:
            d["start"] = int(d["start"] or 1)
        return d

    def label(self, num_id: str, ilvl: int) -> str:
        if num_id not in self.nums or num_id == "0":
            return ""
        aid, over = self.nums[num_id]
        levels = dict(self.abstract.get(aid, {}))
        key = aid
        cnt = self.counters.setdefault(key, [0] * 10)
        for i in range(ilvl + 1, 10):
            cnt[i] = 0
        lv = dict(levels.get(ilvl, {}))
        lv.update({k: v for k, v in over.get(ilvl, {}).items() if k != "start"})
        start = over.get(ilvl, {}).get("start", lv.get("start", 1))
        if cnt[ilvl] == 0:
            cnt[ilvl] = start
        else:
            cnt[ilvl] += 1
        text = lv.get("lvlText") or ""
        fmt = lv.get("numFmt") or "decimal"
        if fmt == "bullet":
            return _bullet(text)
        if fmt == "none":
            return re.sub(r"%\d", "", text)

        def rep(m):
            i = int(m[1]) - 1
            lvi = dict(levels.get(i, {}))
            lvi.update({k: v for k, v in over.get(i, {}).items() if k != "start"})
            n = cnt[i] if cnt[i] else lvi.get("start", 1)
            return number_text(n, lvi.get("numFmt") or "decimal")

        return re.sub(r"%(\d)", rep, text)


def _bullet(t: str) -> str:
    if not t:
        return "・"
    o = ord(t[0])
    if 0xF000 <= o <= 0xF0FF:
        return WINGDINGS.get(o) or SYMBOL_FONT.get(o) or "・"
    return t


_KANJI = "〇一二三四五六七八九"
_AIUEO = "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン"
_IROHA = "イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセスン"


def _roman(n: int) -> str:
    vals = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"),
            (9, "IX"), (5, "V"), (4, "IV"), (1, "I")]
    out = ""
    for v, s in vals:
        while n >= v:
            out += s
            n -= v
    return out


def _kanji(n: int) -> str:
    if n < 10:
        return _KANJI[n]
    if n < 100:
        t, o = divmod(n, 10)
        return ("" if t == 1 else _KANJI[t]) + "十" + (_KANJI[o] if o else "")
    return str(n)


def number_text(n: int, fmt: str) -> str:
    if fmt in ("decimal", "decimalHalfWidth"):
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt in ("decimalFullWidth", "decimalFullWidth2"):
        return str(n).translate(str.maketrans("0123456789", "０１２３４５６７８９"))
    if fmt == "decimalEnclosedCircle" or fmt == "decimalEnclosedCircleChinese":
        return chr(0x2460 + n - 1) if 1 <= n <= 20 else (chr(0x3251 + n - 21) if 21 <= n <= 35 else str(n))
    if fmt == "decimalEnclosedParen":
        return chr(0x2474 + n - 1) if 1 <= n <= 20 else f"({n})"
    if fmt == "decimalEnclosedFullstop":
        return chr(0x2488 + n - 1) if 1 <= n <= 20 else f"{n}."
    if fmt == "upperRoman":
        return _roman(n)
    if fmt == "lowerRoman":
        return _roman(n).lower()
    if fmt == "upperLetter":
        return chr(ord("A") + (n - 1) % 26) * ((n - 1) // 26 + 1)
    if fmt == "lowerLetter":
        return chr(ord("a") + (n - 1) % 26) * ((n - 1) // 26 + 1)
    if fmt in ("aiueo", "aiueoFullWidth"):
        ch = _AIUEO[(n - 1) % len(_AIUEO)]
        return ch if fmt == "aiueoFullWidth" else ch
    if fmt in ("iroha", "irohaFullWidth"):
        return _IROHA[(n - 1) % len(_IROHA)]
    if fmt in ("japaneseCounting", "chineseCounting", "japaneseLegal", "chineseCountingThousand", "taiwaneseCounting"):
        return _kanji(n)
    if fmt == "ideographTraditional":
        return "甲乙丙丁戊己庚辛壬癸"[(n - 1) % 10]
    if fmt == "ideographZodiac":
        return "子丑寅卯辰巳午未申酉戌亥"[(n - 1) % 12]
    if fmt == "ordinal":
        return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"
    return str(n)


class DocxDocument:
    kind = "docx"

    def __init__(self, blob: bytes, filename: str = ""):
        self.blob = blob
        self.filename = filename
        self.zip = RawZip(blob)
        self.embeds: list[dict] = []  # 埋め込んだ OLE の部品（word/embeddings の部品）
        self.charts: list[dict] = []  # グラフ（word/charts の部品）：本文の〔グラフ1〕の所
        self.diagrams: list[dict] = []  # SmartArt（word/diagrams のデータの部品）：本文の〔SmartArt1〕の所
        self.pictures: list[dict] = []  # 絵（a:blip・v:imagedata）：説明（代わりの字）と飾りの印。段階とは別の注意に使う
        self.main = self._main_part()
        self.parts: dict[str, tuple[bytes, Node]] = {}
        self.rels = self._rels(self.main)
        self.has_macros = any(n.lower().endswith("vbaproject.bin") for n in self.zip.names())
        styles = self._load("word/styles.xml")
        self.style_names: dict[str, str] = {}
        self.style_num: dict[str, tuple[str, int]] = {}
        self.style_based: dict[str, str] = {}
        self.style_outline: dict[str, int] = {}  # スタイルの見出しの段（w:outlineLvl。0〜8 が見出し、9 は本文）
        if styles is not None:
            for st in styles.kids[0].findall("style") if styles.kids else []:
                sid = st.get("styleId")
                nm = st.find("name")
                self.style_names[sid] = nm.get("val") if nm is not None else sid
                b = st.find("basedOn")
                if b is not None:
                    self.style_based[sid] = b.get("val")
                ppr = st.find("pPr")
                ol = ppr.find("outlineLvl") if ppr is not None else None
                if ol is not None and str(ol.get("val", "")).isdigit():
                    self.style_outline[sid] = int(ol.get("val"))
                npr = ppr.find("numPr") if ppr is not None else None
                if npr is not None:
                    nid = npr.find("numId")
                    il = npr.find("ilvl")
                    self.style_num[sid] = (nid.get("val") if nid is not None else "", int(il.get("val")) if il is not None else 0)
        self.numbering = Numbering(self._load("word/numbering.xml"))
        settings = self._load("word/settings.xml")
        self.protection = ""
        self.protection_enforced = False
        if settings is not None and settings.kids:
            dp = settings.kids[0].find("documentProtection")
            if dp is not None:
                self.protection = dp.get("edit", "")
                self.protection_enforced = dp.get("enforcement", "0") in ("1", "true", "on")
        self.controls: list[Control] = []
        self.body = self._story(self.main)
        self.stories: list[tuple[str, str, list]] = []  # (種類, 部品, ブロック)
        for rid, (typ, target) in self.rels.items():
            kind = typ.rsplit("/", 1)[-1]
            if kind in ("header", "footer", "footnotes", "endnotes", "comments"):
                name = _join(self.main, target)
                if self.zip.has(name):
                    self.stories.append((kind, name, self._story(name)))

    # ------------------------------------------------------------------ 部品

    @property
    def vba(self):
        """マクロ（VBA）のプロジェクト：モジュールとソース（todoku.vba.Project）。マクロがなければ None。"""
        if getattr(self, "_vba", None) is None and self.has_macros:
            from ..vba import read_project

            try:
                self._vba = read_project(self._vba_blob())
            except Exception as e:  # noqa: BLE001
                self._vba = e
        v = getattr(self, "_vba", None)
        return v if v is not None and not isinstance(v, Exception) else None

    def _vba_blob(self) -> bytes:
        name = next(n for n in self.zip.names() if n.lower().endswith("vbaproject.bin"))
        return self.zip.read(name)

    def _main_part(self) -> str:
        rels = self.zip.read("_rels/.rels") if self.zip.has("_rels/.rels") else b""
        m = re.search(rb'Target="/?([^"]*document[^"]*\.xml)"', rels)
        name = m[1].decode() if m else "word/document.xml"
        if not self.zip.has(name):
            raise DocxError("Word の文書の本体（document.xml）がありません")
        return name

    def _rels(self, part: str) -> dict[str, tuple[str, str]]:
        d, base = part.rsplit("/", 1) if "/" in part else ("", part)
        rp = f"{d}/_rels/{base}.rels" if d else f"_rels/{base}.rels"
        out = {}
        if self.zip.has(rp):
            for m in re.finditer(rb"<(?:\w+:)?Relationship\b([^>]*)/?>", self.zip.read(rp)):
                a = dict(re.findall(r'(\w+)="([^"]*)"', m[1].decode("utf-8", "replace")))
                if "Id" in a:
                    out[a["Id"]] = (a.get("Type", ""), a.get("Target", ""))
        return out

    def _load(self, name: str) -> Node | None:
        if name in self.parts:
            return self.parts[name][1]
        if not self.zip.has(name):
            return None
        xml = self.zip.read(name)
        root = parse(xml)
        self.parts[name] = (xml, root)
        return root

    def xml(self, part: str) -> bytes:
        return self.parts[part][0]

    # ------------------------------------------------------------------ 本文
    def _story(self, part: str) -> list:
        root = self._load(part)
        if root is None or not root.kids:
            return []
        top = root.kids[0]
        container = top.find("body") if top.local == "document" else top
        if container is None:
            return []
        if top.local in ("footnotes", "endnotes", "comments"):
            out = []
            for item in top.kids:
                if item.local in ("footnote", "endnote", "comment"):
                    if item.get("type") in ("separator", "continuationSeparator", "continuationNotice"):
                        continue
                    out += self._blocks(item, part)
            return out
        return self._blocks(container, part)

    def _blocks(self, node: Node, part: str, cell: Cell | None = None) -> list:
        out = []
        for ch in node.kids:
            loc = ch.local
            if loc == "p":
                out.append(self._para(ch, part, cell))
            elif loc == "tbl":
                out.append(self._table(ch, part))
            elif loc == "sdt":
                content = ch.find("sdtContent")
                ctl = self._sdt_control(ch, part, None)
                inner = self._blocks(content, part, cell) if content is not None else []
                if ctl is not None:
                    ctl.result_segs = [s for b in inner if isinstance(b, Para) for s in b.segs]
                    ctl.para = next((b for b in inner if isinstance(b, Para)), None)
                    for b in inner:
                        if isinstance(b, Para):
                            for s in b.segs:
                                s.field = s.field or ctl
                out += inner
            elif loc in ("customXml", "ins", "moveTo"):
                out += self._blocks(ch, part, cell)
        return out

    def _table(self, node: Node, part: str) -> Table:
        t = Table(node, part)
        for r, tr in enumerate([x for x in node.kids if x.local in ("tr", "sdt", "customXml")]):
            rows = [tr] if tr.local == "tr" else [x for x in tr.iter("tr")]
            for tr2 in rows:
                row = []
                col = 0
                for tc in [x for x in tr2.kids if x.local == "tc"] + [x for s in tr2.kids if s.local in ("sdt", "customXml") for x in s.iter("tc")]:
                    pr = tc.find("tcPr")
                    span = 1
                    vm = ""
                    shade = ""
                    if pr is not None:
                        gs = pr.find("gridSpan")
                        if gs is not None:
                            span = int(gs.get("val", 1) or 1)
                        v = pr.find("vMerge")
                        if v is not None:
                            vm = v.get("val", "continue") or "continue"
                        sh = pr.find("shd")
                        if sh is not None:
                            shade = sh.get("fill", "") or ""
                    c = Cell(tc, len(t.rows), col, span, vm, table=t, shade=shade)
                    c.blocks = self._blocks(tc, part, c)
                    row.append(c)
                    col += span
                t.rows.append(row)
        return t

    def _para(self, node: Node, part: str, cell: Cell | None) -> Para:
        p = Para(node, part, [], cell=cell)
        ppr = node.find("pPr")
        num = None
        if ppr is not None:
            ps = ppr.find("pStyle")
            if ps is not None:
                p.style = ps.get("val", "")
            npr = ppr.find("numPr")
            if npr is not None:
                nid = npr.find("numId")
                il = npr.find("ilvl")
                num = (nid.get("val") if nid is not None else "", int(il.get("val")) if il is not None else 0)
        if num is None and p.style:
            s = p.style
            seen = 0
            while s and seen < 20:
                if s in self.style_num:
                    num = self.style_num[s]
                    break
                s = self.style_based.get(s, "")
                seen += 1
        if num is not None and num[0]:
            p.label = self.numbering.label(num[0], num[1])
        self._inline(node, p, part, {"stack": []})
        return p

    def _inline(self, node: Node, p: Para, part: str, st: dict, deleted: bool = False) -> None:
        for ch in node.kids:
            loc = ch.local
            if loc == "r":
                self._run(ch, p, part, st, deleted)
            elif loc in ("hyperlink", "smartTag", "customXml", "ins", "moveTo", "fldSimple", "bdo", "dir"):
                if loc == "fldSimple":
                    st["simple"] = ch.get("instr", "")
                self._inline(ch, p, part, st, deleted)
            elif loc in ("del", "moveFrom"):
                self._inline(ch, p, part, st, True)
            elif loc == "sdt":
                ctl = self._sdt_control(ch, part, p)
                content = ch.find("sdtContent")
                if content is not None:
                    n0 = len(p.segs)
                    self._inline(content, p, part, st, deleted)
                    if ctl is not None:
                        ctl.result_segs = p.segs[n0:]
                        for s in ctl.result_segs:
                            s.field = ctl
                        if not ctl.result_segs:
                            p.segs.append(Seg("", "field", content, None, field=ctl))
                            ctl.result_segs = [p.segs[-1]]
            elif loc == "oMath" or loc == "oMathPara":
                txt = "".join(n.text for n in ch.iter("t"))
                if txt:
                    p.segs.append(Seg(txt, "math", None))

    def _run(self, r: Node, p: Para, part: str, st: dict, deleted: bool) -> None:
        rpr = r.find("rPr")
        hidden = deleted
        underline = False
        sym_font = ""
        if rpr is not None:
            v = rpr.find("vanish")
            if v is not None and v.get("val", "true") not in ("0", "false", "off"):
                hidden = True
            u = rpr.find("u")
            if u is not None and u.get("val", "single") not in ("none", "0"):
                underline = True
            rf = rpr.find("rFonts")
            if rf is not None:
                sym_font = rf.get("ascii", "") or rf.get("hAnsi", "") or ""
        stack: list = st["stack"]
        for ch in r.kids:
            loc = ch.local
            in_instr = bool(stack) and stack[-1]["phase"] == "instr"
            if loc == "fldChar":
                typ = ch.get("fldCharType")
                if typ == "begin":
                    stack.append({"phase": "instr", "instr": "", "begin": ch, "ctl": None, "run": r})
                elif typ == "separate" and stack:
                    f = stack[-1]
                    f["phase"] = "result"
                    f["ctl"] = self._form_field(f, part, p)
                    f["n0"] = len(p.segs)
                elif typ == "end" and stack:
                    f = stack.pop()
                    if f.get("ctl") is None and f["phase"] == "instr":
                        f["ctl"] = self._form_field(f, part, p)
                        f["n0"] = len(p.segs)
                    ctl = f.get("ctl")
                    if ctl is not None:
                        ctl.end = ch
                        ctl.result_segs = p.segs[f.get("n0", len(p.segs)):]
                        for s in ctl.result_segs:
                            s.field = s.field or ctl
                        if ctl.kind == "checkbox":
                            p.segs.append(Seg("☑" if ctl.checked else "☐", "field", ch, r, field=ctl))
                        elif not ctl.result_segs:
                            p.segs.append(Seg("", "field", ch, r, field=ctl))
                            ctl.result_segs = [p.segs[-1]]
                continue
            if loc == "instrText":
                if stack:
                    stack[-1]["instr"] += ch.text
                continue
            if in_instr:
                continue
            if loc == "t":
                p.segs.append(Seg(ch.text, "t", ch, r, hidden, underline))
            elif loc == "delText":
                continue
            elif loc == "tab" or loc == "ptab":
                p.segs.append(Seg("\t", "tab", ch, r, hidden, underline))
            elif loc in ("br", "cr"):
                p.segs.append(Seg("\n", "br", ch, r, hidden))
            elif loc == "noBreakHyphen":
                p.segs.append(Seg("‑", "other", ch, r, hidden))
            elif loc == "sym":
                code = int(ch.get("char", "0") or "0", 16)
                font = ch.get("font", "") or sym_font
                if code < 0xF000:
                    code2 = code + 0xF000
                else:
                    code2 = code
                if "wingding" in font.lower():
                    txt = WINGDINGS.get(code2, "◆")
                elif "symbol" in font.lower():
                    txt = SYMBOL_FONT.get(code2, chr(code2 - 0xF000) if code2 >= 0xF020 else "・")
                else:
                    txt = chr(code) if code < 0xE000 else "・"
                p.segs.append(Seg(txt, "sym", ch, r, hidden, underline))
            elif loc in ("footnoteReference", "endnoteReference"):
                p.segs.append(Seg(f"※{ch.get('id', '')}", "note", ch, r, hidden))
            elif loc in ("drawing", "pict", "object", "AlternateContent"):
                self._boxes(ch, p, part)
                axc = next(iter(ch.iter("control")), None) if loc in ("object", "pict") else None
                if axc is not None:
                    from ..axfile import word_control

                    word_control(self, ch, axc, part, p, r, hidden)  # ActiveX の部品（w:control）
                    continue
                for mark in self._graphics(ch, part):
                    p.segs.append(Seg(mark, "other", None, r, hidden))
                ole = next(iter(ch.iter("OLEObject")), None)
                e = self._add_embed(ole, part, p, ch) if ole is not None else None
                alt = [n.get("descr") or n.get("title") for n in ch.iter("docPr") if n.get("descr") or n.get("title")]
                if e is not None and "alt" not in e:
                    e["alt"] = next((n.get("alt") for n in ch.iter("shape") if n.get("alt")), "") or (alt[0] if alt else "")
                if e is not None:
                    p.segs.append(Seg(f"〔埋め込み{e['n']}：{e['label']}〕", "other", None, r, hidden))
                elif self._picture(ch, p, alt[0] if alt else ""):
                    pic = self.pictures[-1]
                    mark = f"〔図: {pic['alt']}〕" if pic["alt"] else f"〔図{pic['n']}（説明なし）〕" if not pic["decorative"] else ""
                    if mark:
                        p.segs.append(Seg(mark, "other", None, r, hidden))
                elif alt:
                    p.segs.append(Seg(f"〔図: {alt[0]}〕", "other", None, r, hidden))
            elif loc == "ruby":
                base = ch.find("rubyBase")
                rt = ch.find("rt")
                bt = "".join(n.text for n in base.iter("t")) if base is not None else ""
                rtt = "".join(n.text for n in rt.iter("t")) if rt is not None else ""
                p.segs.append(Seg(bt, "ruby", None, r, hidden))
                if rtt:
                    p.segs.append(Seg("（", "other", None, r, hidden))
                    p.segs.append(Seg(rtt, "ruby", None, r, hidden))
                    p.segs.append(Seg("）", "other", None, r, hidden))

    def _add_embed(self, ole: Node, part: str, p: Para, holder: Node | None = None) -> dict | None:
        """埋め込んだ OLE の部品（o:OLEObject）を、文書の埋め込みの一覧に足す：部品の置き場所（word/embeddings/…）を関係から引く。
        holder（w:object など）の中の絵（v:imagedata・a:blip）は、Word が表示・印刷に使う表の絵として覚える（images）。"""
        rid = ole.get("id")
        rels = self._rels(part)
        typ, target = rels.get(rid or "", ("", ""))
        if not target or ole.get("Type", "Embed") != "Embed":
            return None
        name = _join(part, target)
        if not self.zip.has(name):
            return None
        for e in self.embeds:
            if e["storage"] == name:
                return e
        images = []
        for n in list(holder.iter("imagedata")) + list(holder.iter("blip")) if holder is not None else []:
            t = rels.get(n.get("id") or n.get("embed") or "", ("", ""))[1]
            img = _join(part, t) if t else ""
            if img and self.zip.has(img) and img not in images:
                images.append(img)
        prog = ole.get("ProgID") or ""
        from ..embeds.ole import label_of

        label = label_of(prog)
        e = {"n": len(self.embeds) + 1, "storage": name, "prog": prog, "label": label, "para": p, "part": part,
             "images": images, "aspect": ole.get("DrawAspect") or "Content"}
        self.embeds.append(e)
        return e

    def _picture(self, node: Node, p: Para, alt: str) -> bool:
        """図の中に絵（a:blip・v:imagedata）があれば、絵の一覧に足す。代わりの中身（Fallback）とテキストボックスの中は数えない。"""
        found = None
        for n in list(node.iter("blip")) + list(node.iter("imagedata")):
            if not any(x.local in ("Fallback", "txbxContent") for x in _ancestors(n, node)):
                found = n
                break
        if found is None:
            return False
        rid = found.get("embed") or found.get("id") or ""
        target = self._rels(p.part).get(rid, ("", ""))[1] if rid else ""
        media = _join(p.part, target) if target else ""
        pr = next(iter(node.iter("docPr")), None)
        deco = pr is not None and any(x.local == "decorative" and x.get("val") in ("1", "true") for x in pr.iter("decorative"))
        if not alt:
            alt = next((s.get("alt") for s in node.iter("shape") if s.get("alt")), "") or ""  # VML の絵の説明
        self.pictures.append({"n": len(self.pictures) + 1, "alt": alt.strip(), "decorative": deco, "para": p,
                              "media": media if self.zip.has(media) else ""})
        return True

    def _graphics(self, node: Node, part: str) -> list[str]:
        """図の中のグラフ（c:chart・新しい種類の cx:chart）と SmartArt（dgm:relIds）を一覧に足し、本文に出す印を返す。
        AlternateContent の代わりの中身（Fallback）は数えない。"""
        out = []
        rels = None
        for g in node.iter("graphicData"):
            if any(x.local in ("Fallback", "txbxContent") for x in _ancestors(g, node)):
                continue  # 代わりの中身・テキストボックスの中（中身を段落として読むときに数える）
            for k in g.kids:
                if k.local not in ("chart", "relIds"):
                    continue
                rels = rels if rels is not None else self._rels(part)
                rid = k.get("id") if k.local == "chart" else k.get("dm")
                target = _join(part, rels.get(rid or "", ("", ""))[1]) if rid in rels else ""
                if not target or not self.zip.has(target):
                    continue
                pr = next((n for n in node.iter("docPr")), None)
                alt = (pr.get("descr") or pr.get("title") or "") if pr is not None else ""
                if k.local == "chart":
                    self.charts.append({"n": len(self.charts) + 1, "part": target, "alt": alt})
                    out.append(f"〔グラフ{len(self.charts)}〕")
                else:
                    self.diagrams.append({"n": len(self.diagrams) + 1, "part": target, "alt": alt})
                    out.append(f"〔SmartArt{len(self.diagrams)}〕")
        return out

    def embedded_books(self) -> list[tuple[dict, object]]:
        """埋め込んだ Excel のブック：(埋め込みの情報, ブック)。.xlsx の部品も、OLE の入れ物（.bin）の中のブックも読む。"""
        out = []
        for e in self.embeds:
            wb = self.embedded_book(e)
            if wb is not None:
                out.append((e, wb))
        return out

    def embedded_book(self, e: dict):
        """埋め込み1つの Excel のブック（Excel のブックでない・開けないときは None）。"""
        from ..cfb import CFB
        from ..xls.book import XlsWorkbook
        from ..xlsx.workbook import Workbook

        try:
            data = self.zip.read(e["storage"])
            if data[:4] == b"PK\x03\x04":
                return Workbook(data, e["storage"])
            if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                c = CFB(data)
                names = {n.lower(): n for n in c.streams()}
                if "workbook" in names or "book" in names:
                    return XlsWorkbook(data, e["storage"])
                if "package" in names:
                    return Workbook(c.read(names["package"]), e["storage"])
        except Exception:  # noqa: BLE001
            return None
        return None

    def _boxes(self, node: Node, p: Para, part: str) -> None:
        """図形・テキストボックスの中の文字（段落の後ろに付ける）。描画キャンバス・グループの中のテキストボックスも、
        すべて読む。AlternateContent は、Word が表示する Choice の方だけを読む（Fallback は同じ中身の代わりなので数えない）。
        テキストボックスの中のテキストボックスは、外の箱の段落の中で読む（ここでは二重に数えない）。"""
        for tx in node.iter("txbxContent"):
            ups = list(_ancestors(tx, node))
            if any(a.local in ("Fallback", "txbxContent") for a in ups):
                continue
            p.boxes += self._blocks(tx, part)

    # ------------------------------------------------------------------ 入力欄
    def _form_field(self, f: dict, part: str, p: Para) -> Control | None:
        instr = f["instr"].strip()
        word = instr.split()[0].upper() if instr.split() else ""
        if word not in ("FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN"):
            return None
        ff = f["begin"].find("ffData")
        name = ""
        ctl = Control({"FORMTEXT": "text", "FORMCHECKBOX": "checkbox", "FORMDROPDOWN": "dropdown"}[word], "formfield",
                      "", f["begin"], part, p)
        if ff is not None:
            n = ff.find("name")
            name = n.get("val", "") if n is not None else ""
            ht = ff.find("helpText") or ff.find("statusText")
            if ht is not None:
                ctl.help = ht.get("val", "")
            cb = ff.find("checkBox")
            if cb is not None:
                c = cb.find("checked")
                d = cb.find("default")
                if c is not None:
                    ctl.checked = c.get("val", "1") not in ("0", "false")
                elif d is not None:
                    ctl.checked = d.get("val", "0") not in ("0", "false")
            dd = ff.find("ddList")
            if dd is not None:
                ctl.options = [x.get("val", "") for x in dd.findall("listEntry")]
                res = dd.find("result")
                idx = int(res.get("val", 0)) if res is not None else 0
                ctl.value = ctl.options[idx] if 0 <= idx < len(ctl.options) else None
            ti = ff.find("textInput")
            if ti is not None:
                ml = ti.find("maxLength")
                if ml is not None and ml.get("val"):
                    ctl.maxlen = int(ml.get("val"))
                tp = ti.find("type")
                if tp is not None and tp.get("val") == "date":
                    ctl.kind = "date"
                fmt = ti.find("format")
                if fmt is not None:
                    ctl.date_format = fmt.get("val", "")
        ctl.name = name
        self.controls.append(ctl)
        return ctl

    def _sdt_control(self, sdt: Node, part: str, p: Para | None) -> Control | None:
        pr = sdt.find("sdtPr")
        if pr is None:
            return None
        kind = "richtext"
        options: list[str] = []
        date_fmt = ""
        checked = False
        for k in pr.kids:
            loc = k.local
            if loc == "text":
                kind = "text"
            elif loc == "date":
                kind = "date"
                df = k.find("dateFormat")
                date_fmt = df.get("val", "") if df is not None else ""
            elif loc in ("dropDownList", "comboBox"):
                kind = "dropdown" if loc == "dropDownList" else "combo"
                options = [li.get("displayText") or li.get("value") or "" for li in k.findall("listItem")]
            elif loc == "checkbox" and k.tag.startswith("{" + W14):
                kind = "checkbox"
                c = k.find("checked")
                checked = c is not None and c.get("val", "1") in ("1", "true")
            elif loc == "picture":
                kind = "picture"
            elif loc in ("docPartObj", "docPartList", "bibliography", "citation", "group", "equation"):
                return None
        alias = pr.find("alias")
        tag = pr.find("tag")
        lock = pr.find("lock")
        ctl = Control(kind, "sdt", (alias.get("val") if alias is not None else "") or (tag.get("val") if tag is not None else ""),
                      sdt, part, p, options=options, checked=checked, date_format=date_fmt,
                      placeholder=pr.find("showingPlcHdr") is not None, lock=lock.get("val", "") if lock is not None else "")
        self.controls.append(ctl)
        return ctl

    # ------------------------------------------------------------------ 全体
    def all_blocks(self):
        """(種類, 部品, ブロック)。本文が先、ヘッダー・フッター・脚注・コメントが後。"""
        yield "body", self.main, self.body
        for kind, part, blocks in self.stories:
            yield kind, part, blocks


def _ancestors(n: Node, stop: Node):
    """n の親を、stop まで順に。"""
    x = n.parent
    while x is not None and x is not stop:
        yield x
        x = x.parent


def iter_paras(blocks):
    for b in blocks:
        if isinstance(b, Para):
            yield b
            if b.boxes:
                yield from iter_paras(b.boxes)
        elif isinstance(b, Table):
            for row in b.rows:
                for c in row:
                    yield from iter_paras(c.blocks)


def _join(base: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    d = base.rsplit("/", 1)[0] if "/" in base else ""
    parts = (d + "/" + target).split("/") if d else target.split("/")
    out = []
    for x in parts:
        if x == "..":
            if out:
                out.pop()
        elif x and x != ".":
            out.append(x)
    return "/".join(out)


def strict_ns_used(xml: bytes) -> bool:
    return any(ns.encode() in xml[:2000] for ns in STRICT)
