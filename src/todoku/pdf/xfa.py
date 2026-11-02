"""PDFの XFA のフォーム（Adobe XML Forms Architecture 3.3）を読む。

XFA は、AcroForm の /XFA にある XML の部品（packet）：template（フォームの作り）・datasets（データ）・config など。
/XFA は「名前・ストリーム」の並びの配列か、XDP 全体の1つのストリーム。部品をつないだものが1つの XDP になる。

- 動的なフォーム（config の dynamicRender が required、Catalog の NeedsRendering）：ページの描画は代わりの表示で、
  フォームの中身は template と datasets にしかない
- 静的なフォーム：PDF の欄（AcroForm）があり、名前は XFA の名前の式（SOM。例 form1[0].#subform[0].LastName[0]）。
  Adobe Acrobat・Reader が見せる値は datasets から来る

読み方（仕様 XFA 3.3）：
- プロトタイプ（use・usehref）を、template を読むときに解く（6章「Prototypes」）
- データ（datasets の xfa:data）を、データの群（dataGroup）と値（dataValue）に分ける（4章）
- template とデータを結ぶ（4章「Basic Data Binding」・9章）：名前が同じものを、同じ名前の中の順番どおりに結ぶ（直接）。
  なければ、上の階層の結ばれたデータの群の中を探す（祖先の一致）。bind の match（once・global・dataRef・none）、
  名前のない subform・area は透明、くり返し（occur の min・max・initial）はデータの数で決まる
- 欄の id は、Form DOM の名前の式（名前のないものは #subform[k] のように、種類の中の順番）
"""

from __future__ import annotations

import re
import xml.parsers.expat as expat
from dataclasses import dataclass, field

from ..docx.xmlpos import _tag_end
from .objects import PdfString, Ref, Stream, decode_text

XFA_DATA = "http://www.xfa.org/schema/xfa-data/1.0/"
XHTML = "http://www.w3.org/1999/xhtml"
XSI = ("http://www.w3.org/2001/XMLSchema-instance", "http://www.w3.org/2000/10/XMLSchema-instance")
XDP = "http://ns.adobe.com/xdp/"


class XfaError(ValueError):
    pass


# ====================================================================== XML（バイトの位置つき）
class El:
    """XML の要素。kids は要素と文字（str）の並び。s〜se が開始タグ、es〜e が終了タグ（空要素は se == es == e）。"""

    __slots__ = ("tag", "ns", "attrs", "kids", "parent", "s", "se", "es", "e")

    def __init__(self, tag: str, ns: str, attrs: dict, s: int, parent: "El | None"):
        self.tag, self.ns, self.attrs, self.s, self.parent = tag, ns, attrs, s, parent
        self.kids: list = []
        self.se = self.es = self.e = -1

    def get(self, name: str, default=None):
        return self.attrs.get(name, default)

    @property
    def elements(self) -> list["El"]:
        return [k for k in self.kids if isinstance(k, El)]

    def find(self, tag: str) -> "El | None":
        return next((k for k in self.kids if isinstance(k, El) and k.tag == tag), None)

    def text(self) -> str:
        return "".join(k for k in self.kids if isinstance(k, str))

    def alltext(self) -> str:
        return "".join(k if isinstance(k, str) else k.alltext() for k in self.kids)

    def iter(self):
        stack = [self]
        while stack:
            n = stack.pop()
            yield n
            stack.extend(reversed(n.elements))

    def __repr__(self) -> str:
        return f"<{self.tag} {self.s}>"


_BARE_AMP = re.compile(rb"&(?!(?:[A-Za-z_][\w.-]*|#\d+|#x[0-9A-Fa-f]+);)")


def repair_xml(data: bytes) -> tuple[bytes, list[int]]:
    """エスケープされていない「&」（実際の様式のデータにある）を &amp; にする。(直したもの, 直した所の元の位置)。"""
    pos = [m.start() for m in _BARE_AMP.finditer(data)]
    return (_BARE_AMP.sub(b"&amp;", data), pos) if pos else (data, [])


def parse_xml(data: bytes) -> El:
    """expat で読む。名前空間は El.ns に、属性は名前空間のないものは名前だけ、あるものは {ns}名前。
    エスケープされていない「&」は直して読み、位置は元のバイト列の位置に戻す。"""
    fixed, amps = repair_xml(data)
    root = _parse(fixed)
    if amps:
        import bisect

        shift = [p + 4 * i for i, p in enumerate(amps)]  # 直したものの中の、足した「amp;」の位置の前

        def back(x: int) -> int:
            return x - 4 * bisect.bisect_right(shift, x)
        for e in root.iter():
            if e is not root:
                e.s, e.se, e.es, e.e = back(e.s), back(e.se), back(e.es), back(e.e)
    return root


def _parse(data: bytes) -> El:
    p = expat.ParserCreate(namespace_separator="}")
    p.buffer_text = True
    p.buffer_size = 1 << 16
    root = El("#root", "", {}, 0, None)
    stack = [root]

    def split(tag: str) -> tuple[str, str]:
        if "}" in tag:
            ns, local = tag.rsplit("}", 1)
            return local, ns
        return tag, ""

    def start(tag, attrs):
        local, ns = split(tag)
        a = {}
        for k, v in attrs.items():
            kl, kns = split(k)
            a[kl if not kns else "{" + kns + "}" + kl] = v
        s = p.CurrentByteIndex
        node = El(local, ns, a, s, stack[-1])
        node.se = _tag_end(data, s)
        stack[-1].kids.append(node)
        stack.append(node)

    def end(_tag):
        node = stack.pop()
        if data[node.se - 2:node.se] == b"/>":
            node.es = node.e = node.se
        else:
            node.es = p.CurrentByteIndex
            node.e = _tag_end(data, node.es)

    def chars(text):
        kids = stack[-1].kids
        if kids and isinstance(kids[-1], str):
            kids[-1] += text
        else:
            kids.append(text)

    p.StartElementHandler = start
    p.EndElementHandler = end
    p.CharacterDataHandler = chars
    try:
        p.Parse(data, True)
    except expat.ExpatError as e:
        raise XfaError(f"XFA の XML を読めません（{e}）") from None
    return root


# ====================================================================== 部品（packet）
@dataclass
class Packet:
    name: str
    ref: Ref | None  # 配列の形のときのストリーム
    data: bytes  # 展開したバイト列
    start: int = 0  # つないだ XDP の中の位置
    end: int = 0


class Xdp:
    """XFA の部品をつないだ XDP。mode は array（名前・ストリームの配列）か stream（1つのストリーム）。"""

    def __init__(self, doc):
        af = doc.resolve(doc.root.get("AcroForm"))
        x = af.get("XFA") if isinstance(af, dict) else None
        self.array_ref = x if isinstance(x, Ref) else None
        x = doc.resolve(x)
        self.packets: list[Packet] = []
        self.stream_ref = None
        if isinstance(x, Stream):
            self.mode = "stream"
            self.stream_ref = af.get("XFA") if isinstance(af.get("XFA"), Ref) else None
            self.packets.append(Packet("xdp", self.stream_ref, doc.decode(x)))
        elif isinstance(x, list):
            self.mode = "array"
            self.array = x
            for i in range(0, len(x) - 1, 2):
                nm = doc.resolve(x[i])
                ref = x[i + 1]
                st = doc.resolve(ref)
                name = decode_text(nm) if isinstance(nm, (bytes, PdfString)) else getattr(nm, "name", str(nm))
                self.packets.append(Packet(name, ref if isinstance(ref, Ref) else None,
                                           doc.decode(st) if isinstance(st, Stream) else b""))
        else:
            raise XfaError("/XFA が、配列でもストリームでもありません")
        buf = bytearray()
        for pk in self.packets:
            pk.start = len(buf)
            buf += pk.data
            pk.end = len(buf)
        self.combined = bytes(buf)
        self.root = self._parse()
        xdp = next((e for e in self.root.elements), None)
        if xdp is None:
            raise XfaError("XDP の要素がありません")
        self.xdp = xdp
        # 部品の要素（template・datasets・config など）。最初に出たもの
        self.parts: dict[str, El] = {}
        for e in xdp.elements:
            self.parts.setdefault(e.tag, e)
        self.encoding = self._encoding()

    def _parse(self) -> El:
        try:
            return parse_xml(self.combined)
        except XfaError:
            if self.mode != "array":
                raise
        # 部品ごとに XML の宣言などがあって、つなぐと XML にならないとき：宣言を除いて読み直す
        buf = bytearray()
        for pk in self.packets:
            d = re.sub(rb"^\xef\xbb\xbf", b"", pk.data)
            if pk is not self.packets[0]:
                d = re.sub(rb"^\s*<\?xml[^>]*\?>", lambda m: b" " * len(m[0]), d)
            pk.start = len(buf)
            buf += d
            pk.end = len(buf)
        self.combined = bytes(buf)
        return parse_xml(self.combined)

    def _encoding(self) -> str:
        m = re.match(rb'\s*<\?xml[^>]*encoding=["\']([A-Za-z0-9._-]+)', self.combined)
        return m[1].decode().lower() if m else "utf-8"

    def packet_of(self, el: El) -> Packet:
        for pk in self.packets:
            if pk.start <= el.s < pk.end:
                return pk
        raise XfaError("要素の部品が見つかりません")

    @property
    def names(self) -> list[str]:
        return [pk.name for pk in self.packets]


# ====================================================================== プロトタイプを解いた template
class T:
    """template の要素（プロトタイプを解いたもの）。src は元の要素（プロトタイプから受け継いだものは、その要素）。"""

    __slots__ = ("tag", "attrs", "kids", "src", "content")

    def __init__(self, tag: str, attrs: dict, kids: list, src: El, content: El):
        self.tag, self.attrs, self.kids, self.src = tag, attrs, kids, src
        self.content = content  # 文字の中身（data content）を持つ要素

    def get(self, name: str, default=None):
        return self.attrs.get(name, default)

    def find(self, tag: str) -> "T | None":
        return next((k for k in self.kids if k.tag == tag), None)

    def findall(self, tag: str) -> list["T"]:
        return [k for k in self.kids if k.tag == tag]

    def text(self) -> str:
        return self.content.text() if self.content is not None else ""

    def path(self, *tags) -> "T | None":
        cur = self
        for t in tags:
            cur = cur.find(t) if cur is not None else None
        return cur


_NOINHERIT = {"id", "name", "use", "usehref"}


class Template:
    def __init__(self, el: El):
        self.el = el
        self.ids: dict[str, El] = {}
        for e in el.iter():
            if "id" in e.attrs:
                self.ids.setdefault(e.attrs["id"], e)
        self.cache: dict[int, T] = {}
        self.unresolved: list[str] = []
        self.used_protos: set[int] = set()
        self._busy: set[int] = set()

    def resolve(self, el: El) -> T:
        k = id(el)
        if k in self.cache:
            return self.cache[k]
        proto = None
        if id(el) not in self._busy:
            self._busy.add(id(el))
            try:
                pe = self._proto_of(el)
                if pe is not None:
                    self.used_protos.add(id(pe))
                    proto = self.resolve(pe)
            finally:
                self._busy.discard(id(el))
        t = self._merge(el, proto)
        self.cache[k] = t
        return t

    def _merge(self, el: El, proto: T | None, extra: T | None = None) -> T:
        base = proto or extra
        if proto is not None and extra is not None:
            base = self._combine(proto, extra)
        attrs = {k: v for k, v in (base.attrs.items() if base else ()) if k not in _NOINHERIT}
        attrs.update({k: v for k, v in el.attrs.items() if k not in ("use", "usehref")})
        content = el if (el.text().strip() or base is None or el.ns == XHTML or el.tag == "exData") else base.content
        if el.tag in ("exData",) or el.ns == XHTML:
            kids = [self._plain(k) for k in el.elements] if (el.elements or base is None) else list(base.kids)
            return T(el.tag, attrs, kids, el, el if el.elements or el.text().strip() or base is None else base.content)
        own: dict[str, int] = {}
        kids = []
        bkids = list(base.kids) if base else []
        btype: dict[str, list[T]] = {}
        for b in bkids:
            btype.setdefault(b.tag, []).append(b)
        for c in el.elements:
            i = own.get(c.tag, 0)
            own[c.tag] = i + 1
            pk = btype.get(c.tag, [])
            pk = pk[i] if i < len(pk) else None
            if pk is None:
                kids.append(self.resolve(c))
            else:
                cp = self._proto_of(c)
                kids.append(self._merge(c, self.resolve(cp) if cp is not None else None, pk))
        seen: dict[str, int] = {}
        for b in bkids:
            j = seen.get(b.tag, 0)
            seen[b.tag] = j + 1
            if j >= own.get(b.tag, 0):
                kids.append(b)
        return T(el.tag, attrs, kids, el, content)

    def _combine(self, a: T, b: T) -> T:
        attrs = dict(b.attrs)
        attrs.update(a.attrs)
        return T(a.tag, attrs, a.kids or b.kids, a.src, a.content if a.text().strip() else b.content)

    def _plain(self, el: El) -> T:
        return T(el.tag, dict(el.attrs), [self._plain(k) for k in el.elements], el, el)

    def _proto_of(self, el: El) -> El | None:
        ref = el.attrs.get("usehref") or el.attrs.get("use")
        if not ref:
            return None
        target = None
        m = re.fullmatch(r"\.?#som\((.*)\)", ref.strip())
        if ref.startswith("#") and not ref.startswith("#som("):
            target = self.ids.get(ref[1:])
        elif m:
            target = self._som(m[1])
        elif "#" not in ref and "use" in el.attrs and not el.attrs.get("usehref"):
            target = self._som(ref)
        if target is None or target.tag != el.tag or target is el:
            self.unresolved.append(f"{el.tag}: {ref}")
            return None
        return target

    def _som(self, expr: str) -> El | None:
        """template の中の名前の式（$template.#subform.名前.名前）。proto・area・名前のない subform の中も探す。"""
        parts = [p for p in re.split(r"\.(?![^\[]*\])", expr.strip()) if p]
        if not parts:
            return None
        if parts[0] in ("$template", "xfa.template"):
            cur, parts = self.el, parts[1:]
        elif parts[0] == "xfa" and len(parts) > 1 and parts[1] == "template":
            cur, parts = self.el, parts[2:]
        else:
            # 相対の式：どこかにある、最初の名前の要素から
            cands = [e for e in self.el.iter() if e.attrs.get("name") == re.sub(r"\[.*\]$", "", parts[0])]
            if not cands:
                return None
            cur, parts = cands[0], parts[1:]
        for p in parts:
            m = re.fullmatch(r"(#?)([^\[]+)(?:\[(\d+)\])?", p)
            if not m:
                return None
            idx = int(m[3] or 0)
            if m[1]:
                hits = [e for e in cur.elements if e.tag == m[2]]
            else:
                hits = _named(cur, m[2])
            if idx >= len(hits):
                return None
            cur = hits[idx]
        return cur


def _named(el: El, name: str) -> list[El]:
    out = []
    for c in el.elements:
        if c.attrs.get("name") == name:
            out.append(c)
        elif c.tag in ("proto", "area", "variables") or (c.tag in ("subform", "exclGroup") and not c.attrs.get("name")):
            out.extend(_named(c, name))
    return out


# ====================================================================== データ（Data DOM）
class D:
    """データの節：kind は group（群）か value（値）。attr があれば属性の値（metadata）。"""

    __slots__ = ("name", "el", "kind", "kids", "parent", "attr", "null")

    def __init__(self, name, el, kind, parent, attr=None):
        self.name, self.el, self.kind, self.parent, self.attr = name, el, kind, parent, attr
        self.kids: list[D] = []
        self.null = False

    @property
    def value(self) -> str:
        if self.attr is not None:
            return self.el.attrs.get(self.attr, "")
        if self.null:
            return ""
        out = []
        for k in self.el.kids:
            if isinstance(k, str):
                out.append(k)
            elif k.ns == XHTML:
                out.append(xhtml_text(k))
            else:
                d = next((x for x in self.kids if x.el is k and x.attr is None), None)
                out.append(d.value if d is not None else k.alltext())
        return "".join(out)

    def path(self) -> list[tuple[str, int]]:
        """xfa:data からの道筋：(名前, 同じ名前の中の順番)。属性は ("@名前", 0)。"""
        out = []
        cur = self
        while cur is not None and cur.parent is not None:
            if cur.attr is not None:
                out.append(("@" + cur.name, 0))
            else:
                sib = [k for k in cur.parent.kids if k.attr is None and k.name == cur.name]
                out.append((cur.name, sib.index(cur)))
            cur = cur.parent
        return list(reversed(out))


def _excluded_attr(k: str) -> bool:
    if k.startswith("{"):
        ns = k[1:].split("}")[0]
        return ns in XSI or ns == XFA_DATA or ns == "http://www.w3.org/XML/1998/namespace"
    return k == "xmlns" or k.startswith("xmlns")


def build_data(el: El, parent: D | None, inside_value: bool = False) -> D:
    force = el.attrs.get("{" + XFA_DATA + "}dataNode")
    has_el = bool(el.elements)
    has_text = any(isinstance(k, str) and k.strip() for k in el.kids)
    if force in ("dataGroup", "dataValue"):
        kind = "group" if force == "dataGroup" else "value"
    elif inside_value or not has_el or has_text or any(k.ns == XHTML for k in el.elements):
        kind = "value"
    else:
        kind = "group"
    d = D(el.tag, el, kind, parent)
    nil = next((v for k, v in el.attrs.items() if k.startswith("{") and k[1:].split("}")[0] in XSI and k.endswith("}nil")), None)
    d.null = nil == "true"
    for k, v in el.attrs.items():
        if not _excluded_attr(k):
            d.kids.append(D(k.split("}")[-1], el, "value", d, attr=k))
    for c in el.elements:
        if c.ns == XHTML:
            continue  # 書式つきの文字の中身（値の一部）
        d.kids.append(build_data(c, d, kind == "value"))
    return d


class Data:
    def __init__(self, ds: El | None):
        self.datasets = ds
        self.data_el = None
        self.root: D | None = None  # xfa:data
        self.record: D | None = None
        self.others: list[D] = []  # xfa:data の外（一覧のデータなど）
        if ds is None:
            return
        for c in ds.elements:
            if c.tag == "data" and c.ns == XFA_DATA and self.data_el is None:
                self.data_el = c
            elif c.ns != "http://ns.adobe.com/data-description/":
                self.others.append(build_data(c, None))
        if self.data_el is not None:
            self.root = D("data", self.data_el, "group", None)
            for c in self.data_el.elements:
                self.root.kids.append(build_data(c, self.root))
            self.record = next((k for k in self.root.kids if k.kind == "group" and k.attr is None), None)

    def other(self, name: str) -> D | None:
        return next((d for d in self.others if d.name == name), None)


def xhtml_text(el: El) -> str:
    """書式つきの文字（XHTML）を、段落ごとに改行した平文にする。"""
    out: list[str] = []

    def walk(e: El):
        for k in e.kids:
            if isinstance(k, str):
                out.append(re.sub(r"\s+", " ", k))
            elif k.tag == "br":
                out.append("\n")
            else:
                block = k.tag in ("p", "div", "li", "ol", "ul", "h1", "h2", "h3", "h4", "h5", "h6")
                if block and out and not out[-1].endswith("\n"):
                    out.append("\n")
                walk(k)
                if block:
                    out.append("\n")

    walk(el)
    return re.sub(r" *\n *", "\n", "".join(out)).strip()


# ====================================================================== データの名前の式（SOM）
def data_som(expr: str, scope: D | None, data: Data) -> list[D] | None:
    """データの中の名前の式を解く。解けない形なら None、ないときは []。"""
    e = expr.strip()
    if not e:
        return None
    if e.startswith("!"):
        e = "xfa.datasets." + e[1:]
    parts = [p for p in re.split(r"\.(?![^\[]*\])", e) if p != ""]
    if not parts or any(p == "" for p in parts):
        return None
    head = parts[0]
    if head == "$data":
        cur, parts = [data.root] if data.root else [], parts[1:]
    elif head == "$record":
        cur, parts = [data.record] if data.record else [], parts[1:]
    elif head == "$":
        cur, parts = [scope] if scope else [], parts[1:]
    elif head == "xfa" and len(parts) > 1 and parts[1] == "datasets":
        parts = parts[2:]
        if not parts:
            return None
        if parts[0] == "data":
            cur, parts = [data.root] if data.root else [], parts[1:]
        else:
            o = data.other(re.sub(r"\[.*\]$", "", parts[0]))
            cur, parts = ([o] if o else []), parts[1:]
    elif head.startswith("$") or head.startswith("#"):
        return None
    else:
        # 相対の式：いまの所から上へ、最初の名前が見つかる所
        cur = []
        s = scope
        while s is not None:
            if _step([s], parts[0]):
                cur = [s]
                break
            s = s.parent
        if not cur:
            return []
    for p in parts:
        if p == "..":
            return None
        cur = _step(cur, p)
        if cur is None:
            return None
    return cur


def _step(nodes: list[D], seg: str) -> list[D] | None:
    m = re.fullmatch(r"(#?)([^\[\]]+)(?:\[(\*|\d+)\])?", seg)
    if not m:
        return None
    out = []
    for n in nodes:
        if m[1]:
            want = {"dataGroup": "group", "dataValue": "value"}.get(m[2])
            hits = [k for k in n.kids if k.kind == want and k.attr is None]
        else:
            hits = [k for k in n.kids if k.name == m[2] and k.attr is None] or [k for k in n.kids if k.name == m[2]]
        if m[3] == "*":
            out.extend(hits)
        else:
            i = int(m[3] or 0)
            if i < len(hits):
                out.append(hits[i])
    return out


# ====================================================================== Form DOM（データと結んだフォーム）
@dataclass
class Slot:
    """データの置き場所：あれば node、なければ親の置き場所と名前と順番（書くときに作る）。"""
    parent: "Slot | None"
    name: str
    index: int
    kind: str  # group / value
    node: D | None = None
    attr: bool = False

    def path(self) -> list[tuple[str, int]]:
        if self.node is not None:
            return self.node.path()
        base = self.parent.path() if self.parent is not None else []
        return base + [(("@" + self.name) if self.attr else self.name, self.index)]


class F:
    """Form DOM の節。"""

    __slots__ = ("t", "kind", "name", "som", "parent", "kids", "slot", "data", "hidden", "readonly", "page_area",
                 "instance", "occur", "scope_slot", "transparent")

    def __init__(self, t: T | None, kind: str, parent: "F | None"):
        self.t, self.kind, self.parent = t, kind, parent
        self.name = t.get("name", "") if t is not None else ""
        self.kids: list[F] = []
        self.som = ""
        self.slot: Slot | None = None
        self.data: D | None = None
        self.hidden = False
        self.readonly = False
        self.page_area = False
        self.instance = 0
        self.occur = None
        self.scope_slot: Slot | None = None
        self.transparent = False


def _occur(t: T) -> tuple[int, int, int]:
    o = t.find("occur")
    if o is None:
        return 1, 1, 1

    def num(k, default):
        v = o.get(k)
        try:
            return int(v) if v not in (None, "") else default
        except ValueError:
            return default
    mn = num("min", 1)
    mx = num("max", 1)
    ini = num("initial", mn)
    return mn, mx, ini


def bind_of(t: T) -> tuple[str, str]:
    b = t.find("bind")
    if b is None:
        return "once", ""
    return (b.get("match") or "once"), (b.get("ref") or "")


def ui_kind(t: T) -> str:
    ui = t.find("ui")
    if ui is not None:
        for k in ui.kids:
            if k.tag not in ("extras", "picture"):
                return k.tag
    v = t.find("value")
    c = v.kids[0].tag if v is not None and v.kids else "text"
    return {"date": "dateTimeEdit", "time": "dateTimeEdit", "dateTime": "dateTimeEdit", "decimal": "numericEdit",
            "integer": "numericEdit", "float": "numericEdit", "image": "imageEdit", "exData": "textEdit"}.get(c, "textEdit")


class Merge:
    """template とデータを結んで Form DOM を作る（consumeData の結び方）。"""

    def __init__(self, tmpl: Template, data: Data):
        self.tmpl, self.data = tmpl, data
        self.used: set[int] = set()  # 結んだデータ（once の結び方で、ほかに使えない）
        self.globals: dict[str, D] = {}
        self.alloc: dict[tuple[int, str], int] = {}
        self.gslots: dict[str, Slot] = {}
        self.empty = data.record is None
        root_el = next((e for e in tmpl.el.elements if e.tag == "subform"), None)
        if root_el is None:
            raise XfaError("template に subform がありません")
        rt = tmpl.resolve(root_el)
        self.root = F(rt, "subform", None)
        self.root.som = f"{rt.get('name') or '#subform'}[0]"
        rec = data.record
        self.root.slot = Slot(None, rt.get("name") or "form1", 0, "group", rec)
        self.root.data = rec
        if rec is not None:
            self.used.add(id(rec))
        self._fill(self.root, rt, self.root.slot, rec)
        self._ids(self.root)

    # ------------------------------------------------------------------ 中身
    def _fill(self, f: F, t: T, slot: Slot | None, scope: D | None) -> None:
        for c in t.kids:
            tag = c.tag
            if tag == "subform":
                self._subform(f, c, slot, scope)
            elif tag in ("subformSet", "area"):
                g = F(c, tag, f)
                g.transparent = True
                self._attrs(g, c)
                f.kids.append(g)
                self._fill(g, c, slot, scope)
            elif tag == "exclGroup":
                self._exclgroup(f, c, slot, scope)
            elif tag == "field":
                self._field(f, c, slot, scope)
            elif tag == "draw":
                d = F(c, "draw", f)
                self._attrs(d, c)
                f.kids.append(d)
            elif tag == "pageSet":
                ps = F(c, "pageSet", f)
                f.kids.append(ps)
                for pa in c.findall("pageArea"):
                    a = F(pa, "pageArea", ps)
                    a.page_area = True
                    ps.kids.append(a)
                    for k in pa.kids:
                        if k.tag in ("draw", "field", "exclGroup", "subform", "area"):
                            self._static(a, k)

    def _static(self, f: F, t: T) -> None:
        """ページの共通部分（pageArea）の中身：データとは結ばない。"""
        k = {"exclGroup": "exclGroup"}.get(t.tag, t.tag)
        n = F(t, k, f)
        n.page_area = True
        self._attrs(n, t)
        f.kids.append(n)
        for c in t.kids:
            if c.tag in ("draw", "field", "exclGroup", "subform", "area"):
                self._static(n, c)

    def _attrs(self, f: F, t: T) -> None:
        p = f.parent
        f.hidden = (t.get("presence") in ("hidden", "invisible", "inactive")) or (p is not None and p.hidden)
        f.readonly = (t.get("access") in ("readOnly", "protected", "nonInteractive")) or (p is not None and p.readonly)
        f.page_area = p is not None and p.page_area

    def _find(self, name: str, scope: D | None, kind: str, f: F, shared: bool = False) -> D | None:
        """名前の一致：いまの群（直接）→ 上の結ばれた群（祖先の一致）。"""
        tried = []
        s = scope
        cur = f
        while True:
            if s is not None and id(s) not in tried:
                tried.append(id(s))
                for k in s.kids:
                    if k.name == name and k.kind == kind and (k.attr is None or (s.kind == "value")) and \
                            (shared or id(k) not in self.used):
                        if k.attr is not None:
                            continue
                        return k
            cur = cur.parent
            while cur is not None and cur.data is None:
                cur = cur.parent
            if cur is None:
                return None
            s = cur.data

    def _subform(self, f: F, t: T, slot: Slot | None, scope: D | None) -> None:
        name = t.get("name")
        match, ref = bind_of(t)
        mn, mx, ini = _occur(t)
        named = bool(name) and t.get("scope") != "none" and match != "none"
        if match == "dataRef" and ref:
            nodes = data_som(ref, scope, self.data) or []
            nodes = [n for n in nodes if n.kind == "group"]
            count = ini if self.empty else max(mn, len(nodes) if mx < 0 else min(len(nodes), mx))
            for i in range(count):
                node = nodes[i] if i < len(nodes) else None
                sl = Slot(None, node.name, 0, "group", node) if node is not None else \
                    _ref_slot(ref.replace("[*]", f"[{i}]"), slot, self, "group")
                self._instance(f, t, i, sl, node, (mn, mx, ini))
            return
        if named:
            nodes = []
            limit = 10 ** 6 if mx < 0 else max(mx, mn)
            while len(nodes) < limit:
                n = self._find(name, scope, "group", f)
                if n is None:
                    break
                self.used.add(id(n))
                nodes.append(n)
            count = ini if self.empty else max(mn, len(nodes))
            for i in range(count):
                node = nodes[i] if i < len(nodes) else None
                sl = Slot(slot, name, 0 if node is not None else self._new_index(slot, scope, name), "group", node)
                self._instance(f, t, i, sl, node, (mn, mx, ini))
            return
        # 名前のない subform・scope="none"・bind none：データには透明。くり返しは、中身が新しく結ばれる間だけ
        count = ini if self.empty else mn
        i = 0
        while True:
            if i >= count:
                if self.empty or (mx >= 0 and i >= mx) or i >= 200:
                    break
                before = len(self.used)
                g = self._instance(f, t, i, slot, scope, (mn, mx, ini), transparent=True, trial=True)
                if len(self.used) == before:
                    f.kids.remove(g)
                    break
            else:
                self._instance(f, t, i, slot, scope, (mn, mx, ini), transparent=True)
            i += 1

    def _instance(self, f, t, i, slot, node, occur, transparent=False, trial=False) -> F:
        g = F(t, "subform", f)
        g.instance = i
        g.occur = occur
        g.transparent = transparent
        self._attrs(g, t)
        f.kids.append(g)
        if transparent:
            g.scope_slot = slot
            self._fill(g, t, slot, node)
        else:
            g.slot = slot
            g.data = node
            self._fill(g, t, slot, node)
        return g

    def _exclgroup(self, f: F, t: T, slot: Slot | None, scope: D | None) -> None:
        g = F(t, "exclGroup", f)
        self._attrs(g, t)
        f.kids.append(g)
        name = t.get("name")
        match, ref = bind_of(t)
        members = t.findall("field")
        if match == "none":
            for m in members:
                self._field(g, m, None, None, unbound=True)
            return
        node = None
        if match == "dataRef" and ref:
            nodes = data_som(ref, scope, self.data)
            node = nodes[0] if nodes else None
            g.slot = Slot(None, node.name, 0, "value", node) if node is not None else _ref_slot(ref, slot, self)
            for m in members:
                self._field(g, m, None, None, unbound=True)
            return
        elif name:
            node = self._find(name, scope, "value", g, shared=match == "global")
            if node is None:
                grp = self._find(name, scope, "group", g)
                if grp is not None:
                    # 長い形：群の中で、一つずつの欄が結ばれる
                    self.used.add(id(grp))
                    g.data = grp
                    gs = Slot(slot, name, 0, "group", grp)
                    for m in members:
                        self._field(g, m, gs, grp)
                    return
        if name or node is not None:
            if node is not None and match != "global":
                self.used.add(id(node))
            idx = 0 if node is not None else self._new_index(slot, scope, name or "")
            g.slot = Slot(slot, name or "", idx, "value", node)
            g.data = None
            for m in members:
                self._field(g, m, None, None, unbound=True)
        else:
            for m in members:
                self._field(g, m, slot, scope)

    def _new_index(self, slot: Slot | None, scope: D | None, name: str) -> int:
        """データのない欄・群を作るときの、同じ名前の中の順番：いまある同じ名前のデータの数と、先に決めた置き場所の数の後。"""
        key = (id(slot), name)
        n = self.alloc.get(key)
        if n is None:
            n = sum(1 for k in (scope.kids if scope is not None else []) if k.name == name and k.attr is None)
        self.alloc[key] = n + 1
        return n

    def _field(self, f: F, t: T, slot: Slot | None, scope: D | None, unbound: bool = False) -> None:
        x = F(t, "field", f)
        self._attrs(x, t)
        f.kids.append(x)
        if unbound:
            return  # ひとつ選ぶ欄の中（短い形）・bind none の中：一つずつは結ばない
        name = t.get("name")
        match, ref = bind_of(t)
        multi = ui_kind(t) == "choiceList" and (t.path("ui", "choiceList") or T("x", {}, [], None, None)).get("open") == "multiSelect"
        kind = "group" if multi else "value"
        if match == "none" or f.page_area:
            return
        if match == "dataRef" and ref:
            nodes = data_som(ref, scope, self.data)
            if nodes:
                x.slot = Slot(None, nodes[0].name, 0, kind, nodes[0])
            elif nodes is not None:
                x.slot = _ref_slot(ref, slot, self)
            return
        if not name:
            return  # 名前のない欄は、データに透明
        node = None
        if match == "global" and name in self.globals:
            node = self.globals[name]
        if node is None:
            node = self._find(name, scope, kind, x)
        if node is None and match == "global" and name in self.gslots:
            x.slot = self.gslots[name]  # 同じ名前の global の欄は、同じ置き場所を共有する
            return
        if node is None and match == "global" and self.data.root is not None:
            node = next((k for k in self.data.root.kids if k.name == name and k.kind == kind and k is not self.data.record), None)
        if node is not None:
            if match == "global":
                self.globals[name] = node
            else:
                self.used.add(id(node))
        idx = 0 if node is not None else self._new_index(slot, scope, name)
        x.slot = Slot(slot, name, idx, kind, node)
        if match == "global":
            self.gslots.setdefault(name, x.slot)

    # ------------------------------------------------------------------ id（名前の式）
    def _ids(self, root: F) -> None:
        # 名前のあるものは「名前[同じ名前の中の順番]」、名前のないもの（と area）は「#種類[その種類の中の順番]」。
        # 順番は、名前のないもの・area を透明として、すぐ上の透明でない入れ物の中で数える（仕様 3章「Transparent Nodes」
        # 「Reference by Class」）。静的なフォームの PDF の欄の名前も、この形
        def transparent(k: F) -> bool:
            return k.kind == "subformSet" or (k.kind in ("subform", "exclGroup", "area") and not k.name)

        def number(owner: F):
            names: dict[str, int] = {}
            classes: dict[str, int] = {}

            def visit(parent: F):
                for k in parent.kids:
                    n_cls = classes.get(k.kind, 0)
                    classes[k.kind] = n_cls + 1
                    if k.name and k.kind != "subformSet":
                        n = names.get(k.name, 0)
                        names[k.name] = n + 1
                        k.som = f"{parent.som}.{k.name}[{n}]"
                    else:
                        k.som = f"{parent.som}.#{k.kind}[{n_cls}]"
                    if transparent(k):
                        visit(k)
                    else:
                        number(k)
            visit(owner)
        number(root)


def _ref_slot(ref: str, slot: Slot | None, m: Merge, last_kind: str = "value") -> Slot | None:
    """dataRef の式から、作る置き場所を決める（$・$record・$data と、名前[番号] だけの式）。"""
    parts = [p for p in re.split(r"\.(?![^\[]*\])", ref.strip()) if p]
    if not parts:
        return None
    head = parts[0]
    if head == "$":
        cur, parts = slot, parts[1:]
    elif head == "$record":
        cur, parts = m.root.slot, parts[1:]
    elif head == "$data":
        cur, parts = None, parts[1:]
    elif head.startswith(("$", "!", "#")) or head == "xfa":
        return None
    else:
        cur = slot
    if not parts:
        return None
    for j, p in enumerate(parts):
        mm = re.fullmatch(r"([^\[\]#*.]+)(?:\[(\d+)\])?", p)
        if not mm:
            return None
        last = j == len(parts) - 1
        node = None
        if cur is not None and cur.node is not None or (cur is None and m.data.root is not None):
            base = cur.node if cur is not None else m.data.root
            hits = [k for k in base.kids if k.name == mm[1] and k.attr is None]
            i = int(mm[2] or 0)
            node = hits[i] if i < len(hits) else None
        cur = Slot(cur, mm[1], int(mm[2] or 0), last_kind if last else "group", node)
    return cur


# ====================================================================== 欄
KINDS = {"textEdit": "text", "passwordEdit": "text", "numericEdit": "number", "dateTimeEdit": "date",
         "checkButton": "checkbox", "signature": "signature", "button": "button", "imageEdit": "image", "barcode": "barcode"}
WRITABLE = {"text", "number", "date", "time", "datetime", "checkbox", "radio", "combo", "list"}
_NOT_BOUND = ("データに結びつかない欄です（bind が none か、名前がない）。Adobe Acrobat・Reader で書いた値が残る書き方が"
              "仕様で決まらないので、書きません")


@dataclass
class XField:
    id: str
    kind: str
    f: F  # Form DOM の節（欄か exclGroup）
    caption: str = ""
    tooltip: str = ""
    label: str = ""
    options: list[str] = field(default_factory=list)  # 一覧：見せる字。ひとつ選ぶ欄：選択肢の値
    exports: list[str] = field(default_factory=list)  # 一覧：データに書く値
    option_labels: list[str] = field(default_factory=list)  # ひとつ選ぶ欄の、選択肢の見出し
    on: str = ""  # チェックの「入」の値
    off: str = ""
    raw: str | None = None  # データの値（データがなければ template の既定の値）
    bound: bool = False
    value_type: str = "text"
    maxchars: int | None = None
    comb: int | None = None
    lead: int | None = None
    frac: int | None = None
    pictures: dict = field(default_factory=dict)  # format・validate・ui・bind の書式
    format_test: str = "warning"
    required: bool = False
    calculated: str | None = None  # 計算のスクリプトの override（error など）
    text_entry: bool = False
    multi: bool = False
    items_known: bool = True
    match: str = "once"
    why_not: str = ""  # 書けない理由
    hidden: bool = False
    readonly: bool = False
    locale: str = ""
    members: list = field(default_factory=list)  # ひとつ選ぶ欄の、選択肢の欄（F）
    assist: list[str] = field(default_factory=list)  # 説明（toolTip）と読み上げ（speak）の字
    default: str | None = None  # template の既定の値

    @property
    def writable(self) -> bool:
        return not self.why_not


def _text_of(t: T | None, shown: set) -> str:
    """value・caption などの中の字（文字か、書式つきの文字）。出した要素を shown に覚える。"""
    if t is None:
        return ""
    v = t if t.tag == "value" else t.find("value")
    if v is None:
        return ""
    for k in v.kids:
        if k.tag == "exData":
            body = k.content
            if body is None:
                return ""
            for e in body.iter():
                shown.add(id(e))
            return xhtml_text(body) if body.elements else body.text()
        if k.tag in ("text", "integer", "decimal", "float", "date", "time", "dateTime", "boolean"):
            if k.content is not None:
                shown.add(id(k.content))
            return k.text()
    return ""


class XfaForm:
    """XFA のフォームを読んだ結果：部品・template・データ・Form DOM・欄。"""

    def __init__(self, doc):
        self.doc = doc
        self.xdp = Xdp(doc)
        tel = self.xdp.parts.get("template")
        if tel is None:
            raise XfaError("template の部品がありません")
        self.template = Template(tel)
        self.data = Data(self.xdp.parts.get("datasets"))
        cfg = self.xdp.parts.get("config")
        dyn = False
        if cfg is not None:
            for e in cfg.iter():
                if e.tag == "dynamicRender" and e.text().strip() == "required":
                    dyn = True
        self.dynamic = dyn or doc.root.get("NeedsRendering") is True
        m = re.search(r"xfa-template/([\d.]+)/", tel.ns)
        self.version = m[1] if m else ""
        self.locales = {}
        ls = self.xdp.parts.get("localeSet")
        if ls is not None:
            for e in ls.elements:
                if e.tag == "locale" and e.get("name"):
                    self.locales[e.get("name")] = e
        self.merge = Merge(self.template, self.data)
        self.shown: set[int] = set()  # 窓に出した template の要素
        self.bound: set[int] = set()  # 欄の値として出したデータ
        self.fields: list[XField] = []
        self.by_id: dict[str, XField] = {}
        self.by_tid: dict[str, F] = {}
        for f in _walk(self.merge.root):
            src = f.t.src if f.t is not None else None
            tid = src.get("id") if src is not None else None
            if tid and tid not in self.by_tid:
                self.by_tid[tid] = f
        self._collect(self.merge.root)
        # 同じデータを共有する欄（global の結び方など）
        self.shared: dict[tuple, list[str]] = {}
        for x in self.fields:
            for sl in ([x.f.slot] if x.f.slot is not None else [m.slot for m in x.members if m.slot is not None]):
                self.shared.setdefault(tuple(sl.path()), []).append(x.id)

    # ------------------------------------------------------------------ 欄を集める
    def _collect(self, f: F) -> None:
        for k in f.kids:
            if k.kind == "exclGroup":
                self._add(self._radio(k))
            elif k.kind == "field":
                self._add(self._field(k))
            else:
                self._collect(k)

    def _add(self, x: XField) -> None:
        self.fields.append(x)
        self.by_id[x.id] = x

    def locale_of(self, f: F):
        from .xfapicture import Locale

        cur = f
        while cur is not None:
            if cur.t is not None and cur.t.get("locale"):
                name = cur.t.get("locale")
                return name, Locale(self.locales.get(name) or self.locales.get(name.replace("-", "_")))
            cur = cur.parent
        return "en_US", Locale(self.locales.get("en_US"))

    def _common(self, x: XField, t: T) -> None:
        f = x.f
        x.caption = _text_of(t.find("caption"), self.shown).strip()
        a = t.find("assist")
        if a is not None:
            for tag in ("toolTip", "speak"):
                e = a.find(tag)
                if e is not None and e.content is not None and e.text().strip():
                    self.shown.add(id(e.content))
                    x.assist.append(e.text().strip())
                    if not x.tooltip:
                        x.tooltip = e.text().strip()
        val = t.find("validate")
        if val is not None:
            x.required = val.get("nullTest") == "error"
            x.format_test = val.get("formatTest") or "warning"
            pic = val.find("picture")
            if pic is not None and pic.text().strip():
                x.pictures["validate"] = pic.text().strip()
                self.shown.add(id(pic.content))
        for where, path in (("format", ("format", "picture")), ("ui", ("ui", "picture")), ("bind", ("bind", "picture"))):
            p = t.path(*path)
            if p is not None and p.text().strip():
                x.pictures[where] = p.text().strip()
                self.shown.add(id(p.content))
        calc = t.find("calculate")
        if calc is not None:
            x.calculated = calc.get("override") or "error"
        x.hidden, x.readonly = f.hidden, f.readonly
        x.match = bind_of(t)[0]
        x.locale = self.locale_of(f)[0]
        x.label = x.caption or x.tooltip or ""

    def _field(self, f: F) -> XField:
        t = f.t
        ui = ui_kind(t)
        vt = t.find("value")
        vtype = vt.kids[0].tag if vt is not None and vt.kids else "text"
        kind = KINDS.get(ui, "text")
        if ui == "defaultUi":
            kind = {"date": "date", "time": "time", "dateTime": "datetime", "integer": "number", "decimal": "number",
                    "float": "number", "image": "image"}.get(vtype, "text")
        if kind == "date" and vtype in ("time", "dateTime"):
            kind = "time" if vtype == "time" else "datetime"
        x = XField(f.som, kind, f, value_type=vtype)
        self._common(x, t)
        if vt is not None and vt.kids:
            c = vt.kids[0]
            mc = c.get("maxChars")
            if mc and mc.isdigit() and int(mc) > 0:
                x.maxchars = int(mc)
            for k_, a_ in (("lead", "leadDigits"), ("frac", "fracDigits")):
                v_ = c.get(a_)
                if v_ and re.fullmatch(r"-?\d+", v_):
                    setattr(x, k_, int(v_))
            if vtype == "decimal" and x.frac is None:
                x.frac = 2
            if vtype == "exData" and kind == "text":
                x.kind = "richtext"
        comb = t.path("ui", "textEdit", "comb")
        if comb is not None:
            n = comb.get("numberOfCells")
            x.comb = int(n) if n and n.isdigit() and int(n) > 0 else x.maxchars
        default = _text_of(t, self.shown)
        x.default = default or None
        if ui == "checkButton":
            items = t.find("items")
            vals = [k.text() for k in items.kids] if items is not None else []
            if items is not None:
                for k in items.kids:
                    if k.content is not None:
                        self.shown.add(id(k.content))
            x.on = vals[0] if vals else "1"
            x.off = vals[1] if len(vals) > 1 else ("" if vals else "0")
        if ui == "choiceList":
            cl = t.path("ui", "choiceList")
            x.multi = cl is not None and cl.get("open") == "multiSelect"
            x.text_entry = cl is not None and cl.get("textEntry") == "1"
            x.kind = "list" if (x.multi or (cl is not None and cl.get("open") == "always")) else "combo"
            self._items(x, t)
        node = f.slot.node if f.slot is not None else None
        if node is not None:
            x.bound = True
            self.bound.add(id(node))
            if x.multi:
                x.raw = "\n".join(k.value for k in node.kids if k.attr is None)
                for k in node.kids:
                    self.bound.add(id(k))
            else:
                x.raw = node.value
                bp = x.pictures.get("bind")
                if bp and x.raw:
                    # データの書式（bind の picture）で書かれた値を、決まった形にする（仕様 4章「The picture property」）
                    from .xfapicture import parse_value

                    canon = parse_value(x.raw, bp, hint_of(x), self.locale_of(f)[1])
                    if canon is not None:
                        x.raw = canon
        else:
            x.raw = default if default != "" else None
        x.why_not = self._why_not(x, f)
        if not x.label:
            x.label = f.name or f.som
        return x

    def _items(self, x: XField, t: T) -> None:
        disp = save = None
        used: list[int] = []
        for it in t.findall("items"):
            vals = [k.text() for k in it.kids]
            used += [id(k.content) for k in it.kids if k.content is not None]
            if it.get("save") == "1" and save is None:
                save = vals
            elif disp is None:
                disp = vals
        bi = t.find("bindItems")
        if bi is not None and bi.get("ref"):
            nodes = data_som(bi.get("ref"), None, self.data) or []
            if nodes:
                disp = [_rel(n, bi.get("labelRef") or "$") for n in nodes]
                save = [_rel(n, bi.get("valueRef") or "$") for n in nodes]
                used = []  # template の選択肢は使われない（データの一覧に置き換わる）
        self.shown.update(used)
        if not disp and not save:
            saved = self._saved_items(x.f)
            if saved:
                disp, save = saved
        disp = disp if disp is not None else save
        save = save if save is not None else disp
        if not disp and not save:
            x.items_known = False
            return
        n = min(len(disp), len(save))
        x.options, x.exports = list(disp[:n]), list(save[:n])

    def _saved_items(self, f: F):
        """スクリプトが選択肢を入れる欄：Adobe が保存したフォームの状態（form の部品）にある選択肢。"""
        cur = self.xdp.parts.get("form")
        if cur is None:
            return None
        for seg in f.som.split("."):
            m = re.fullmatch(r"(#?)([^\[]+)\[(\d+)\]", seg)
            if not m:
                return None
            hits = [e for e in cur.elements if (e.tag == m[2] if m[1] else e.get("name") == m[2])]
            if int(m[3]) >= len(hits):
                return None
            cur = hits[int(m[3])]
        disp = save = None
        for it in (e for e in cur.elements if e.tag == "items"):
            vals = [k.text() for k in it.elements]
            if it.get("save") == "1" and save is None:
                save = vals
            elif disp is None:
                disp = vals
        if disp is None and save is None:
            return None
        return (disp if disp is not None else save), (save if save is not None else disp)

    def _radio(self, g: F) -> XField:
        t = g.t
        x = XField(g.som, "radio", g)
        self._common(x, t)
        x.members = [k for k in g.kids if k.kind == "field"]
        sel = None
        for m in x.members:
            mt = m.t
            items = mt.find("items")
            vals = [k.text() for k in items.kids] if items is not None else []
            if items is not None:
                for k in items.kids:
                    if k.content is not None:
                        self.shown.add(id(k.content))
            on = vals[0] if vals else ""
            cap = _text_of(mt.find("caption"), self.shown).strip()
            a = mt.find("assist")
            if a is not None:
                for tag in ("toolTip", "speak"):
                    e = a.find(tag)
                    if e is not None and e.content is not None and e.text().strip():
                        self.shown.add(id(e.content))
                        x.assist.append(e.text().strip())
                        cap = cap or e.text().strip()
            x.options.append(on)
            x.option_labels.append(cap or m.name or on)
            dflt = _text_of(mt, self.shown)
            if dflt:
                x.assist.append(dflt)
            if m.slot is not None and m.slot.node is not None:
                x.bound = True
                self.bound.add(id(m.slot.node))
                if m.slot.node.value == on and on != "":
                    sel = on
            elif m.slot is None and g.slot is None and dflt == on and on != "" and sel is None:
                sel = on
        if g.slot is not None and g.slot.node is not None:
            x.bound = True
            self.bound.add(id(g.slot.node))
            sel = g.slot.node.value or None
        elif g.slot is not None:
            sel = _text_of(t, self.shown) or None
        x.raw = sel
        x.why_not = self._why_not(x, g)
        if not x.label:
            x.label = g.name or g.som
        return x

    def _why_not(self, x: XField, f: F) -> str:
        if x.kind not in WRITABLE:
            return {"signature": "電子署名の欄です（字は書けません）", "button": "押すだけのボタンです",
                    "image": "画像の欄です（字は書けません）", "barcode": "ほかの欄から自動で作られるバーコードです",
                    "richtext": "書式つきの文字の欄です（書式を保って書く方法を確かめていないので、書きません）"}.get(
                x.kind, "書けない種類の欄です")
        if f.page_area:
            return "ページの共通部分（各ページの上下など）の欄で、データに結びつきません"
        if x.kind == "radio":
            if f.slot is None and not any(m.slot is not None for m in x.members):
                return _NOT_BOUND
            if not [o for o in x.options if o != ""]:
                return "選択肢の値（items）がない、ひとつ選ぶ欄です（何を書けばよいか、ファイルからは決まりません）"
        elif f.slot is None:
            return _NOT_BOUND
        if x.kind in ("combo", "list") and not x.items_known and not x.text_entry:
            return "選択肢がスクリプトで決まり、ファイルからは分からないので、書きません"
        return ""

    # ------------------------------------------------------------------ 値
    def value(self, x: XField):
        """窓に出す値：文字は決まった形の値（データの値）、チェックは True/False、ひとつ選ぶ欄は選んだ選択肢の値。"""
        if x.kind == "checkbox":
            return x.raw == x.on and x.on != ""
        if x.kind == "list" and x.multi:
            return [v for v in (x.raw or "").split("\n") if v] or None
        return x.raw if x.raw not in (None, "") else None

    def display(self, x: XField) -> str | None:
        """見た目の字（format の書式で整えた字）。整えられないときは None。"""
        from .xfapicture import format_value

        if x.raw in (None, "") or "format" not in x.pictures:
            return None
        return format_value(x.raw, x.pictures["format"], hint_of(x), self.locale_of(x.f)[1])


def hint_of(x: XField) -> str:
    if x.kind == "number" or x.value_type in ("integer", "decimal", "float"):
        return "num"
    if x.kind == "date" or x.value_type == "date":
        return "date"
    return "text"


def _rel(n: D, expr: str) -> str:
    """一覧のデータの1件の中の値（$ はその件、名前は子か属性）。"""
    if expr in ("$", ""):
        return n.value
    name = expr[2:] if expr.startswith("$.") else expr
    k = next((k for k in n.kids if k.name == name), None)
    return k.value if k is not None else ""


def _walk(f: F):
    stack = [f]
    while stack:
        n = stack.pop()
        yield n
        stack.extend(reversed(n.kids))


def read(doc) -> XfaForm | None:
    """XFA のフォームを読む（なければ None。読めないときは XfaError）。結果は文書に覚えておく。"""
    from .form import has_xfa

    cache = doc.__dict__.get("_xfa")
    if cache is not None:
        if isinstance(cache, XfaError):
            raise cache
        return cache if cache != "none" else None
    if not has_xfa(doc):
        doc.__dict__["_xfa"] = "none"
        return None
    try:
        x = XfaForm(doc)
    except XfaError as e:
        doc.__dict__["_xfa"] = e
        raise
    except (RecursionError, ValueError, KeyError, IndexError, AttributeError, TypeError) as e:
        err = XfaError(f"XFA を読めません（{type(e).__name__}: {e}）")
        doc.__dict__["_xfa"] = err
        raise err from None
    doc.__dict__["_xfa"] = x
    return x


# ====================================================================== 窓
KIND_JA = {"text": "文字", "number": "数", "date": "日付", "time": "時刻", "datetime": "日時", "checkbox": "チェック",
           "radio": "ひとつ選ぶ", "combo": "一覧から選ぶ", "list": "一覧から選ぶ", "signature": "署名（電子署名）",
           "button": "ボタン（押すだけ）", "image": "画像", "barcode": "バーコード（自動）", "richtext": "文字（書式つき）"}
_EVENTS_BY_USER = {"click", "mouseUp", "mouseDown", "mouseEnter", "mouseExit"}


def field_dict(x: XfaForm, f: XField) -> dict:
    """窓の入力欄（PDF の欄と同じ形）。"""
    d = {"id": f.id, "kind": f.kind, "label": f.label, "value": x.value(f), "source": "xfa",
         "provenance": "deterministic", "name": f.id}
    if f.tooltip and f.tooltip != f.label:
        d["tooltip"] = f.tooltip
    if f.kind in ("combo", "list"):
        d["options"] = list(f.options)
        if f.exports != f.options:
            d["export_values"] = list(f.exports)
    if f.kind == "radio":
        d["options"] = list(f.options)
        d["option_labels"] = list(f.option_labels)
    if f.kind == "checkbox":
        d["on_value"], d["off_value"] = f.on, f.off
    if f.maxchars:
        d["maxlen"] = f.maxchars
    if f.comb:
        d["comb"] = f.comb
    if f.assist:
        d["assist"] = list(dict.fromkeys(f.assist))
    if f.default and f.default != d["value"]:
        d["default"] = f.default
    rules = rules_of(f)
    if rules:
        d["rules"] = rules
    fmt = x.display(f)
    if fmt is not None and fmt != d["value"]:
        d["display"] = fmt
    if f.pictures.get("format") or f.kind in ("date", "number"):
        d["format"] = write_as(f)
    if f.required:
        d["required"] = True
    if f.hidden:
        d["hidden"] = True
    if f.calculated:
        d["calculated"] = f.calculated
    if not f.writable:
        d["readonly"] = True
        d["why_not"] = f.why_not
        if f.why_not == _NOT_BOUND:
            d["no_data"] = True  # データに結びつかない欄（bind none・名前なし）
    elif f.readonly:
        d["readonly"] = True
        d["why_not"] = "読み取り専用の欄です"
    elif f.calculated == "error":
        d["readonly"] = True
        d["why_not"] = "フォームのスクリプトが計算して決める欄です（Adobe で開くと計算し直され、上書きできません）"
    d["bound"] = f.bound
    if f.f.slot is not None:
        others = [i for i in x.shared.get(tuple(f.f.slot.path()), []) if i != f.id]
        if others:
            d["shared_with"] = others
            d.setdefault("rules", []).append("同じデータを共有する欄：" + "、".join(others[:5]))
    d["_sample"] = sample(x, f)
    if d["_sample"] is None and not d.get("why_not"):
        d["readonly"] = True
        d["why_not"] = ("選択肢がない一覧の欄です（書ける値が、ファイルからは決まりません）" if f.kind in ("combo", "list", "radio")
                        else "欄の決まり（文字数の上限と書式）が両立せず、どの値も決まりに合いません（書きません）")
    return d


def sample(x: XfaForm, f: XField):
    """検査で書く試しの値：欄の決まり（書式・桁・文字数）に合う値。"""
    from .xfapicture import alternatives

    if f.kind == "checkbox":
        return not x.value(f)
    if f.kind == "radio":
        opts = [o for o in f.options if o != "" and o != f.raw]
        return opts[-1] if opts else (f.options[0] if f.options else None)
    if f.kind in ("combo", "list"):
        opts = [o for o, e in zip(f.options, f.exports) if e not in ("", f.raw) and o.strip()]
        if not opts and f.text_entry:
            return "テスト"[:min(f.maxchars or 3, 3)]
        return opts[0] if opts else (f.options[0] if f.options else None)
    if f.kind == "date":
        return "2026-04-01"
    if f.kind == "time":
        return "10:30:00"
    if f.kind == "datetime":
        return "2026-04-01T10:30:00"
    loc = x.locale_of(f.f)[1]
    for k in ("validate", "bind", "ui", "format"):
        pic = f.pictures.get(k)
        if not pic:
            continue
        alts = alternatives(pic, hint_of(f), loc) or []
        for cat, body in alts:
            if cat == "date":
                return "2026-04-01"
            if cat == "text":
                from .xfapicture import _text_slots

                v = "".join({"9": "1", "A": "A", "X": "A", "O": "1", "0": "1", "t": "\t"}[s]
                            for kk, s in _text_slots(body) if kk == "sym")
                if v and len(v) <= min(f.maxchars or 999, f.comb or 999):
                    return v
            if cat == "num":
                return "1"
        if f.kind == "text" and alts:
            return None  # 書式と文字数の決まりが両立しない（どの値も書けない）
    if f.kind == "number" or f.value_type in ("integer", "decimal", "float"):
        return "1"
    n = min(f.maxchars or 99, f.comb or 99)
    return "テスト"[:n] if n >= 3 else "1" * n


def write_as(f: XField) -> str:
    """書くときの値の形（エージェントに示す）。"""
    if f.kind == "date":
        s = "日付は YYYY-MM-DD で書く（データにはこの形で入る）"
        if f.pictures.get("format"):
            s += f"。見た目の書式 {f.pictures['format']}"
        return s
    if f.kind == "number":
        s = "数は 1234.5 の形（区切りなし）で書く"
        if f.value_type == "integer":
            s = "整数で書く（区切りなし）"
        if f.pictures.get("format"):
            s += f"。見た目の書式 {f.pictures['format']}"
        return s
    return f"書式 {f.pictures.get('format') or f.pictures.get('validate')}"


def rules_of(f: XField) -> list[str]:
    out = []
    if f.maxchars:
        out.append(f"{f.maxchars} 字まで")
    if f.comb:
        out.append(f"1字ずつ {f.comb} マス")
    if f.value_type == "integer":
        out.append("整数")
    if f.lead is not None and f.lead >= 0:
        out.append(f"整数部 {f.lead} 桁まで")
    if f.frac is not None and f.frac >= 0 and f.value_type == "decimal":
        out.append(f"小数部 {f.frac} 桁まで")
    for k, ja in (("format", "見た目の書式"), ("validate", "確かめの書式"), ("ui", "入力の書式"), ("bind", "データの書式")):
        if f.pictures.get(k):
            out.append(f"{ja} {f.pictures[k]}")
    if f.required:
        out.append("必須")
    if f.calculated:
        out.append("スクリプトで計算される欄" + ("（上書きできない）" if f.calculated in ("error", "disabled") else ""))
    if f.kind == "checkbox":
        out.append(f"入の値 {f.on!r}・切の値 {f.off!r}")
    if f.kind in ("combo", "list") and f.text_entry:
        out.append("一覧にない値も書ける")
    if f.multi:
        out.append("いくつも選べる（改行で区切る）")
    if f.match in ("global", "dataRef"):
        out.append({"global": "同じ名前の欄と値を共有する（global）", "dataRef": "データの決まった所に結ぶ（dataRef）"}[f.match])
    return out


def _mark(x: XfaForm, f: XField) -> str:
    v = x.value(f)
    if f.kind == "checkbox":
        return ("☑" if v else "☐") + f"［{f.id}］"
    if f.kind == "radio":
        return "　".join(("◉" if f.raw == o and o != "" else "○") + f"［{f.id}＝{o}］{lab if lab != o else ''}"
                        for o, lab in zip(f.options, f.option_labels))
    if isinstance(v, list):
        v = "、".join(v)
    return f"［{f.id}＝{v}］" if v not in (None, "") else f"［{f.id}］"


def build_window(x: XfaForm) -> dict:
    """XFA の窓：入力欄・読み順（フォームの中の順）・ページの共通部分・データ・表示に使われない字・スクリプト。"""
    reading: list[str] = []
    master: list[str] = []
    plain: list[str] = []
    fields = [field_dict(x, f) for f in x.fields]

    def emit(out: list[str], line: str):
        line = line.strip()
        if line:
            out.append(line)
            plain.append(line)

    def draw_text(f: F) -> str:
        t = f.t
        v = t.find("value")
        if v is None:
            return ""
        for k in v.kids:
            if k.tag == "exData" and k.content is not None:
                for e in k.content.iter():
                    x.shown.add(id(e))
                return _rich(x, k.content) if k.content.elements else k.content.text()
            if k.tag in ("text", "integer", "decimal", "float", "date", "time", "dateTime", "boolean"):
                if k.content is not None:
                    x.shown.add(id(k.content))
                return k.text()
        return ""

    def walk(f: F, out: list[str], depth: int):
        for k in f.kids:
            hid = "（表示されない）" if k.hidden and not f.hidden else ""
            if k.kind == "pageSet":
                for pa in k.kids:
                    emit(master, f"【{pa.som}】（ページの共通部分）")
                    walk(pa, master, depth + 1)
                continue
            if k.kind == "draw":
                txt = draw_text(k)
                if txt.strip():
                    for i, line in enumerate(txt.split("\n")):
                        emit(out, (hid if i == 0 else "") + line)
                continue
            if k.kind in ("field", "exclGroup"):
                xf = x.by_id.get(k.som)
                if xf is None:
                    continue
                cap = xf.caption or ""
                mark = _mark(x, xf)
                if xf.kind == "checkbox":
                    line = f"{mark} {cap}"
                elif xf.kind == "radio":
                    line = (cap + " " if cap else "") + mark
                else:
                    line = (cap + " " if cap else "") + mark
                emit(out, hid + line)
                continue
            if k.kind == "subform" and k.name and not k.transparent:
                rep = k.occur is not None and (k.occur[1] != 1 or k.occur[0] != 1)
                if depth == 0 or rep:
                    emit(out, f"【{k.som}】" + (f"（くり返し：{k.occur[0]}〜{'上限なし' if k.occur[1] < 0 else k.occur[1]} 回）"
                                                if rep else ""))
            walk(k, out, depth + (0 if k.transparent else 1))

    walk(x.merge.root, reading, 0)
    data_lines = _unbound_data(x)
    other_lines = _other_data(x)
    scripts = _scripts(x)
    rest = _unshown(x)
    for line in data_lines + other_lines + rest:
        plain.append(line)
    plain_scripts = [s["source"] for s in scripts]
    return {"readable": True, "kind": "dynamic" if x.dynamic else "static", "version": x.version,
            "packets": x.xdp.names, "mode": x.xdp.mode, "fields": fields, "reading": reading, "master": master,
            "unbound_data": data_lines, "lists": other_lines, "unshown": rest, "scripts": scripts,
            "unresolved": list(dict.fromkeys(x.template.unresolved)),
            "_plain": plain, "_scripts": plain_scripts}


def _rich(x: XfaForm, body: El) -> str:
    """書式つきの文字。中に置いた欄（xfa:embed）は［欄＝値］にする。"""
    def walk(e: El, out: list):
        for k in e.kids:
            if isinstance(k, str):
                out.append(re.sub(r"\s+", " ", k))
                continue
            emb = k.attrs.get("{" + XFA_DATA + "}embed")
            if emb:
                tgt = x.by_tid.get(emb.lstrip("#"))
                xf = x.by_id.get(tgt.som) if tgt is not None else None
                out.append(_mark(x, xf) if xf is not None else f"［{emb}］")
                continue
            block = k.tag in ("p", "div", "li", "h1", "h2", "h3", "h4", "h5", "h6")
            if k.tag == "br":
                out.append("\n")
            if block and out and not "".join(out).endswith("\n"):
                out.append("\n")
            walk(k, out)
            if block:
                out.append("\n")
    out: list[str] = []
    walk(body, out)
    return re.sub(r" *\n *", "\n", "".join(out)).strip()


def _dpath(d: D) -> str:
    return ".".join(f"{n}[{i}]" if not n.startswith("@") else n for n, i in d.path())


def _unbound_data(x: XfaForm) -> list[str]:
    """xfa:data の中の、欄に結びつかない値（隠れた値・スクリプトの使う値など）。"""
    out = []
    if x.data.root is None:
        return out

    def walk(d: D):
        for k in d.kids:
            if k.kind == "value" and id(k) not in x.bound and k.value.strip():
                out.append(f"{_dpath(k)}：{k.value}")
            if k.kind == "group" or k.attr is None:
                walk(k)
    walk(x.data.root)
    return out


def _other_data(x: XfaForm) -> list[str]:
    """datasets の中の、xfa:data の外のデータ（選択肢の一覧など）。群ごとに1行。"""
    out = []

    def label(d: D) -> str:
        attrs = [f"{k.name}={k.value}" for k in d.kids if k.attr is not None]
        return d.name + (f"（{'・'.join(attrs)}）" if attrs else "")

    def walk(d: D, prefix: str):
        here = f"{prefix}.{label(d)}" if prefix else label(d)
        vals = [k for k in d.kids if k.attr is None and k.kind == "value"]
        if d.kind == "value":
            if d.value.strip() or any(k.attr is not None for k in d.kids):
                out.append(f"{here}：{d.value}")
            return
        if vals:
            out.append(f"{here}：" + "／".join(label(k) + ("＝" + k.value if k.value else "") for k in vals))
        for k in d.kids:
            if k.attr is None and k.kind == "group":
                walk(k, here)
    for d in x.data.others:
        walk(d, "")
    return out


def _scripts(x: XfaForm) -> list[dict]:
    """スクリプト（実行しない）：同じ中身はまとめ、どの欄のどの動きで動くかを並べる。"""
    found: dict[str, dict] = {}

    def owner_of(e: El) -> tuple[str, str]:
        kind = ""
        cur = e.parent
        while cur is not None and cur.tag not in ("field", "subform", "exclGroup", "draw", "pageArea", "template"):
            if cur.tag == "event":
                kind = f"{cur.get('activity') or ''}" + (f"（{cur.get('ref')}）" if cur.get("ref") else "")
            elif cur.tag in ("calculate", "validate", "variables"):
                kind = {"calculate": "計算", "validate": "確かめ", "variables": "共通の関数・変数"}[cur.tag]
            cur = cur.parent
        name = cur.get("name") if cur is not None else ""
        return f"{cur.tag if cur is not None else ''} {name or '（名前なし）'}", kind

    for e in x.template.el.iter():
        if e.tag != "script":
            continue
        src = e.text()
        if not src.strip():
            continue
        who, kind = owner_of(e)
        if e.get("name"):
            kind = f"{kind} {e.get('name')}".strip()
        lang = "FormCalc" if "formcalc" in (e.get("contentType") or "").lower() else "JavaScript"
        ent = found.setdefault(src, {"source": src, "language": lang, "where": [], "automatic": False})
        ent["where"].append(f"{who}：{kind}")
        act = kind.split("（")[0]
        if act not in _EVENTS_BY_USER:
            ent["automatic"] = True
        x.shown.add(id(e))
    return list(found.values())


def _unshown(x: XfaForm) -> list[str]:
    """template の中の字で、ほかの所に出していないもの（説明の情報・使われない書式・プロトタイプの中など）。"""
    out = []
    for e in x.template.el.iter():
        if id(e) in x.shown or e.tag in ("image", "script"):
            continue
        t = e.text().strip()
        if not t or e.ns == XHTML and id(e) in x.shown:
            continue
        where = e.tag
        p = e.parent
        hops = 0
        while p is not None and hops < 4:
            where = f"{p.tag}{('「' + p.get('name') + '」') if p.get('name') else ''} › {where}"
            if p.tag in ("field", "subform", "draw", "exclGroup"):
                break
            p = p.parent
            hops += 1
        flat = re.sub(r"\s+", " ", t)
        out.append(f"{where}：{flat}")
    return out


# ====================================================================== Markdown
def _md(s) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def _kind_md(d: dict) -> str:
    k = KIND_JA.get(d["kind"], d["kind"])
    if d["kind"] == "radio":
        k += ": " + "／".join(o if lab == o else f"{o}（{lab}）" for o, lab in zip(d["options"], d["option_labels"]))
    elif d.get("options") is not None:
        opts = [str(o) for o in d["options"]]
        ex = d.get("export_values")
        if ex:
            opts = [o if o == e else f"{o}（値 {e}）" for o, e in zip(opts, ex)]
        k += ": " + ("／".join(opts[:20]) + (f"（ほか{len(opts) - 20}件）" if len(opts) > 20 else "") if opts else "（選択肢なし）")
    if d.get("rules"):
        k += "（" + "・".join(d["rules"]) + "）"
    if d.get("hidden"):
        k += "（表示されない）"
    if d.get("why_not"):
        k += f"（書けない：{d['why_not']}）"
    return k


def _value_md(d: dict) -> str:
    v = d.get("value")
    if d["kind"] == "checkbox":
        return "☑ あり" if v else "☐ なし"
    if v in (None, ""):
        return ""
    s = "、".join(map(str, v)) if isinstance(v, list) else str(v)
    if d.get("display"):
        s += f"（見た目 {d['display']}）"
    return s


def markdown_units(w: dict, placed: set[str]) -> list[tuple[str, list[str]]]:
    """XFA の窓の Markdown を、部分に分けられる単位で：[(題, 行)]。placed はページの表に出した欄。"""
    kind = "動的" if w["kind"] == "dynamic" else "静的"
    head = [f"## XFA のフォーム（{kind}なフォーム・XFA {w['version'] or '?'}）", ""]
    if w["kind"] == "dynamic":
        head.append("- Adobe Acrobat・Reader で開くと、ページの描画ではなく、ここにある中身が表示されます"
                    "（ページの描画は、XFA に対応しないソフトのための代わりの表示です）。")
    else:
        head.append("- ページの描画と PDF の欄がありますが、Adobe Acrobat・Reader が見せる値は、XFA のデータ（datasets）から来ます。"
                    "書き込みは、データと PDF の欄の両方に同じ値を書きます。")
    head.append(f"- 中身の部品：{', '.join(w['packets'])}")
    head.append("- 欄の id は、XFA の名前の式（例 form1[0].Page1[0].LastName[0]）です。書くときは、この id を指定します。"
                "日付は YYYY-MM-DD、数は区切りなしの形で書きます（データにはこの形で入り、見た目は欄の書式で整えられます）。")
    head.append("- スクリプトは実行しません（下の「スクリプト」にソースを出しています）。")
    if w["unresolved"]:
        head.append(f"- 解けなかったプロトタイプの参照 {len(w['unresolved'])} 件：" + "、".join(w["unresolved"][:5]))
    head.append("")
    units = []
    rest = [d for d in w["fields"] if d["id"] not in placed]
    lines = list(head)
    if rest:
        n_ok = sum(1 for d in rest if not d.get("why_not"))
        lines += [f"### XFA の入力欄（{len(rest)} か所・書ける欄 {n_ok}）", ""]
        if placed:
            lines += [f"（ページの入力欄の表に出した {len(placed)} 欄は、ここには出していません）", ""]
        lines += ["| 欄 | 見出し | 種類 | いまの値 |", "|---|---|---|---|"]
        for d in rest:
            lines.append(f"| {_md(d['id'])} | {_md(d['label'])} | {_md(_kind_md(d))} | {_md(_value_md(d))} |")
        lines.append("")
    units.append(("XFA の入力欄", lines))
    for title, key, note in (("XFA の読み順（フォームの中の順）", "reading", ""),
                             ("XFA のページの共通部分（各ページの上下など）", "master", ""),
                             ("XFA のデータの中の、欄に結びつかない値", "unbound_data", "（隠れた値・スクリプトが使う値など。道筋：値）"),
                             ("XFA の一覧のデータ（datasets の中の、データの外。選択肢の元など）", "lists", ""),
                             ("XFA の定義の中の、上に出していない字（読み上げの字・説明など）", "unshown", "")):
        if not w[key]:
            continue
        lines = [f"### {title}", ""] + ([note, ""] if note else [])
        lines += [f"- {_md(s)}" for s in w[key]] + [""]
        units.append((title, lines))
    if w["scripts"]:
        auto = sum(1 for s in w["scripts"] if s["automatic"])
        lines = [f"### XFA のスクリプト（{len(w['scripts'])} 種類・実行しません）", "",
                 f"自動で動くもの（開いた時・値を変えた時・計算・確かめなど）{auto} 種類、押した時だけ動くもの "
                 f"{len(w['scripts']) - auto} 種類。同じ中身はまとめて、動く所を並べています。", ""]
        for i, s in enumerate(w["scripts"], 1):
            where = "、".join(s["where"][:8]) + (f" ほか {len(s['where']) - 8} か所" if len(s["where"]) > 8 else "")
            lines += [f"#### スクリプト {i}（{s['language']}・{where}）", "", "```", s["source"].strip("\n"), "```", ""]
        units.append(("XFA のスクリプト", lines))
    return units
