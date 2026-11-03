"""XML を、要素と文字のバイトの位置つきで読む（書き込みを、元のバイト列への小さな差分にするため）。

要素は開始タグ [s, se) と終了タグ [es, e)（空要素 <x/> は se == es == e）。文字（Txt）は [s, e) と、字ごとのバイトの位置
（実体参照 &amp; や文字参照 &#x3000; は、元の字の範囲のまま）を持つ。名前空間は宣言どおりに解決する（接頭辞に頼らない）。
"""

from __future__ import annotations

import re

NS = {
    "office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
    "text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "draw": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
    "fo": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
    "xlink": "http://www.w3.org/1999/xlink",
    "dc": "http://purl.org/dc/elements/1.1/",
    "meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
    "number": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
    "svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0",
    "form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
    "config": "urn:oasis:names:tc:opendocument:xmlns:config:1.0",
    "manifest": "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0",
    "script": "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
    "of": "urn:oasis:names:tc:opendocument:xmlns:of:1.2",
    "calcext": "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0",
    "loext": "urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0",
    "xml": "http://www.w3.org/XML/1998/namespace",
}


def q(name: str) -> str:
    """「text:p」→「{urn:…:text:1.0}p」。"""
    p, local = name.split(":", 1)
    return "{" + NS[p] + "}" + local


class XmlError(ValueError):
    pass


class Txt:
    __slots__ = ("s", "e", "cdata", "parent", "_v")

    def __init__(self, s: int, e: int, cdata: bool, parent: "El"):
        self.s = s
        self.e = e
        self.cdata = cdata
        self.parent = parent
        self._v = None

    tag = "#text"

    def decode(self, data: bytes) -> tuple[str, list[int]]:
        """(字, 字ごとの始まりのバイトの位置 ＋ 終わりの位置)。"""
        if self.cdata:
            raw = data[self.s:self.e]
            return decode_plain(raw, self.s)
        return decode_text(data, self.s, self.e)

    def value(self, data: bytes) -> str:
        if self._v is None:
            self._v = self.decode(data)[0]
        return self._v


class El:
    __slots__ = ("tag", "attrs", "s", "se", "es", "e", "kids", "parent", "nsmap")

    def __init__(self, tag: str, attrs: dict, s: int, se: int, parent: "El | None", nsmap: dict):
        self.tag = tag
        self.attrs = attrs
        self.s = s
        self.se = se
        self.es = se
        self.e = se
        self.kids: list = []
        self.parent = parent
        self.nsmap = nsmap  # この要素で使える 接頭辞 → 名前空間

    @property
    def local(self) -> str:
        return self.tag.rsplit("}", 1)[-1]

    def get(self, name: str, default=None):
        """属性（「table:name」の形、または「{名前空間}名前」）。"""
        key = q(name) if ":" in name and not name.startswith("{") else name
        return self.attrs.get(key, default)

    def is_(self, name: str) -> bool:
        return self.tag == q(name)

    def els(self) -> list["El"]:
        return [k for k in self.kids if isinstance(k, El)]

    def find(self, name: str) -> "El | None":
        t = q(name)
        for k in self.kids:
            if isinstance(k, El) and k.tag == t:
                return k
        return None

    def findall(self, name: str) -> list["El"]:
        t = q(name)
        return [k for k in self.kids if isinstance(k, El) and k.tag == t]

    def iter(self, name: str | None = None):
        t = q(name) if name else None
        stack = [self]
        while stack:
            n = stack.pop()
            if t is None or n.tag == t:
                yield n
            stack.extend(reversed([k for k in n.kids if isinstance(k, El)]))

    def prefix_of(self, uri: str) -> str | None:
        for p, u in self.nsmap.items():
            if u == uri and p:
                return p
        return None

    def __repr__(self) -> str:
        return f"<{self.local} {self.s}-{self.e}>"


_TOK = re.compile(rb'<!--.*?-->|<!\[CDATA\[.*?\]\]>|<\?.*?\?>|<!DOCTYPE[^>]*>|<(?:[^>"\']|"[^"]*"|\'[^\']*\')*>', re.S)
_ENDS = ((b"<!--", b"-->", "注釈（<!--）"), (b"<![CDATA[", b"]]>", "CDATA"), (b"<?", b"?>", "処理命令（<?）"))


def _tokens(data: bytes):
    """タグ・注釈などを順に（「<」から次の「<」へ飛ぶ。1 つの「<」ごとに 1 度だけ照らすので、ファイルの長さに比例する時間で読む）。
    閉じていない「<」・「<!--」は、整っていない XML として XmlError（字として読み飛ばして、終わりまで探し直すことをしない）。"""
    find, match = data.find, _TOK.match
    i = find(b"<")
    while i >= 0:
        m = match(data, i)
        if m is None:
            raise XmlError(f"閉じていない「<」があります（{i} バイト目）")
        tok = m[0]
        for head, tail, what in _ENDS:
            if tok.startswith(head) and not tok.endswith(tail):
                raise XmlError(f"閉じていない{what}があります（{i} バイト目）")
        yield m
        i = find(b"<", m.end())


_NAME = re.compile(rb"</?\s*([^\s/>]+)")
_ATTR = re.compile(rb'([^\s=/<>]+)\s*=\s*(?:"([^"]*)"|\'([^\']*)\')')
_ENT = re.compile(r"&(#x[0-9a-fA-F]+|#[0-9]+|lt|gt|amp|quot|apos);")
_NAMED = {"lt": "<", "gt": ">", "amp": "&", "quot": '"', "apos": "'"}


def unescape_attr(raw: bytes) -> str:
    s = raw.decode("utf-8")
    s = s.replace("\t", " ").replace("\n", " ").replace("\r", " ")  # XML の属性の値の正規化（文字参照の改行は残る）
    return _ENT.sub(lambda m: _NAMED.get(m[1]) or chr(int(m[1][2:], 16) if m[1][1] in "xX" else int(m[1][1:])), s)


def decode_plain(raw: bytes, base: int) -> tuple[str, list[int]]:
    s = raw.decode("utf-8")
    offs = []
    at = base
    for ch in s:
        offs.append(at)
        o = ord(ch)
        at += 1 if o < 0x80 else 2 if o < 0x800 else 3 if o < 0x10000 else 4
    offs.append(at)
    return s, offs


def decode_text(data: bytes, s: int, e: int) -> tuple[str, list[int]]:
    raw = data[s:e]
    if b"&" not in raw:
        return decode_plain(raw, s)
    out: list[str] = []
    offs: list[int] = []
    text = raw.decode("utf-8")
    at = s
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if ch == "&":
            m = _ENT.match(text, i)
            if not m:
                raise XmlError(f"知らない実体参照があります（{text[i:i + 10]}）")
            g = m[1]
            out.append(_NAMED.get(g) or chr(int(g[2:], 16) if g[1] in "xX" else int(g[1:])))
            offs.append(at)
            at += len(m[0])  # 実体参照は ASCII
            i = m.end()
            continue
        out.append(ch)
        offs.append(at)
        o = ord(ch)
        at += 1 if o < 0x80 else 2 if o < 0x800 else 3 if o < 0x10000 else 4
        i += 1
    offs.append(at)
    return "".join(out), offs


def parse(data: bytes) -> El:
    """XML を読み、根（#root）を返す。整っていない XML は XmlError。"""
    root = El("#root", {}, 0, 0, None, {"xml": NS["xml"]})
    stack = [root]
    names = [b""]
    at = 0
    for m in _tokens(data):
        if m.start() > at:
            cur = stack[-1]
            if cur is not root:
                cur.kids.append(Txt(at, m.start(), False, cur))
            elif data[at:m.start()].strip():
                raise XmlError("根の要素の外に字があります")
        at = m.end()
        tok = m[0]
        if tok.startswith(b"<!--") or tok.startswith(b"<?") or tok.startswith(b"<!DOCTYPE"):
            continue
        if tok.startswith(b"<![CDATA["):
            cur = stack[-1]
            cur.kids.append(Txt(m.start() + 9, m.end() - 3, True, cur))
            continue
        if tok.startswith(b"</"):
            if len(stack) <= 1:
                raise XmlError("開始タグのない終了タグがあります")
            nm = _NAME.match(tok)
            if nm is None or nm[1] != names[-1]:
                raise XmlError(f"終了タグが、開始タグと合いません（{names[-1][:40]!r}）")
            names.pop()
            el = stack.pop()
            el.es = m.start()
            el.e = m.end()
            continue
        nm = _NAME.match(tok)
        if not nm:
            raise XmlError("タグを読めません")
        qname = nm[1].decode("utf-8")
        parent = stack[-1]
        nsmap = parent.nsmap
        raw_attrs = []
        body = tok[nm.end():]
        for am in _ATTR.finditer(body):
            k = am[1].decode("utf-8")
            v = am[2] if am[2] is not None else am[3]
            raw_attrs.append((k, v))
            if k == "xmlns" or k.startswith("xmlns:"):
                if nsmap is parent.nsmap:
                    nsmap = dict(parent.nsmap)
                nsmap[k[6:] if k.startswith("xmlns:") else ""] = unescape_attr(v)
        attrs = {}
        for k, v in raw_attrs:
            if k == "xmlns" or k.startswith("xmlns:"):
                continue
            if ":" in k:
                p, local = k.split(":", 1)
                uri = nsmap.get(p)
                key = "{" + uri + "}" + local if uri else k
            else:
                key = k
            attrs[key] = unescape_attr(v)
        if ":" in qname:
            p, local = qname.split(":", 1)
            uri = nsmap.get(p)
            tag = "{" + uri + "}" + local if uri else qname
        else:
            uri = nsmap.get("")
            tag = "{" + uri + "}" + qname if uri else qname
        el = El(tag, attrs, m.start(), m.end(), parent, nsmap)
        parent.kids.append(el)
        if tok.endswith(b"/>"):
            el.es = el.e = m.end()
        else:
            stack.append(el)
            names.append(nm[1])
    if len(stack) != 1:
        raise XmlError("閉じていない要素があります")
    return root


def top(root: El) -> El:
    """根の下の、最初の要素（文書の要素）。"""
    for k in root.kids:
        if isinstance(k, El):
            return k
    raise XmlError("XML に要素がありません")


# ====================================================================== 書く
def esc_text(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def esc_attr(s: str) -> str:
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")
            .replace("\n", "&#10;").replace("\r", "&#13;").replace("\t", "&#9;"))


def set_attr(tag: bytes, qname: str, value: str | None) -> bytes:
    """開始タグのバイト列の、属性 qname（接頭辞つき）の値を変える。value が None なら属性を取り除く。なければ足す。"""
    pat = re.compile(rb'(\s+)' + re.escape(qname.encode()) + rb'\s*=\s*(?:"[^"]*"|\'[^\']*\')')
    m = pat.search(tag)
    if value is None:
        return tag[:m.start()] + tag[m.end():] if m else tag
    new = f'{qname}="{esc_attr(value)}"'.encode()
    if m:
        return tag[:m.start()] + m[1] + new + tag[m.end():]
    end = len(tag) - (2 if tag.endswith(b"/>") else 1)
    while end > 0 and tag[end - 1:end] in (b" ", b"\t", b"\n", b"\r"):
        end -= 1
    return tag[:end] + b" " + new + tag[end:]


def well_formed(data: bytes) -> str | None:
    """標準の XML の読み手（expat）で読めるか。読めなければ、その理由。"""
    import xml.parsers.expat as expat

    p = expat.ParserCreate(namespace_separator="}")
    try:
        p.Parse(data, True)
    except expat.ExpatError as e:
        return str(e)
    return None
