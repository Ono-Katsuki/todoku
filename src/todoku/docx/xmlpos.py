"""XMLを、要素ごとのバイト位置つきで読む（書き換えを元のバイト列への小さな差分にするため）。

expat（Python に組み込みのXMLの読み手）の「いまのバイト位置」を使い、各要素の
開始タグ [s, se) と終了タグ [es, e) の位置を記録する。空要素（<x/>）は se == es == e。
"""

from __future__ import annotations

import re
import xml.parsers.expat as expat

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
STRICT = {
    "http://purl.oclc.org/ooxml/wordprocessingml/main": W,
    "http://purl.oclc.org/ooxml/officeDocument/relationships": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "http://purl.oclc.org/ooxml/drawingml/main": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
}


class Node:
    __slots__ = ("tag", "attrs", "s", "se", "es", "e", "kids", "parent", "text")

    def __init__(self, tag: str, attrs: dict, s: int, parent: "Node | None"):
        self.tag = tag
        self.attrs = attrs
        self.s = s
        self.se = -1
        self.es = -1
        self.e = -1
        self.kids: list[Node] = []
        self.parent = parent
        self.text = ""

    @property
    def local(self) -> str:
        return self.tag.rsplit("}", 1)[-1]

    def get(self, local: str, default=None):
        """属性（名前空間を問わず、名前の後ろの部分で引く）。"""
        for k, v in self.attrs.items():
            if k.rsplit("}", 1)[-1] == local:
                return v
        return default

    def find(self, local: str) -> "Node | None":
        for k in self.kids:
            if k.local == local:
                return k
        return None

    def findall(self, local: str) -> list["Node"]:
        return [k for k in self.kids if k.local == local]

    def iter(self, local: str | None = None):
        stack = [self]
        while stack:
            n = stack.pop()
            if local is None or n.local == local:
                yield n
            stack.extend(reversed(n.kids))

    def __repr__(self) -> str:
        return f"<{self.local} {self.s}-{self.e}>"


def _tag_end(xml: bytes, pos: int) -> int:
    """pos から始まるタグの '>' の次の位置（属性の値の中の '>' は飛ばす）。"""
    q = None
    i = pos
    n = len(xml)
    while i < n:
        c = xml[i]
        if q is not None:
            if c == q:
                q = None
        elif c in (0x22, 0x27):
            q = c
        elif c == 0x3E:
            return i + 1
        i += 1
    return n


def parse(xml: bytes, more_text: frozenset[str] = frozenset()) -> Node:
    """more_text：字を集める要素を足す（名前空間つきの名前。PowerPoint の a:t など）。"""
    p = expat.ParserCreate(namespace_separator="}")
    p.buffer_text = True
    root = Node("#root", {}, 0, None)
    stack = [root]
    text_tags = {"{" + W + "}" + t for t in ("t", "instrText", "delText", "delInstrText")} | {
        "{http://schemas.openxmlformats.org/officeDocument/2006/math}t"} | set(more_text)

    def norm(tag: str) -> str:
        if "}" in tag:
            ns, local = tag.split("}", 1)
            ns = STRICT.get(ns, ns)
            return "{" + ns + "}" + local
        return tag

    def start(tag, attrs):
        tag = norm(tag)
        s = p.CurrentByteIndex
        node = Node(tag, {norm(k): v for k, v in attrs.items()}, s, stack[-1])
        node.se = _tag_end(xml, s)
        stack[-1].kids.append(node)
        stack.append(node)

    def end(tag):
        node = stack.pop()
        if xml[node.se - 2:node.se] == b"/>":
            # 空要素（<x/>）：開始タグだけ
            node.es = node.e = node.se
        else:
            i = p.CurrentByteIndex
            node.es = i
            node.e = _tag_end(xml, i)

    def chars(data):
        node = stack[-1]
        if node.tag in text_tags:
            node.text += data

    p.StartElementHandler = start
    p.EndElementHandler = end
    p.CharacterDataHandler = chars
    p.Parse(xml, True)
    return root


def escape(text: str) -> bytes:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").encode("utf-8")


def prefix_of(xml: bytes, ns: str = W) -> str:
    """その名前空間に使われている接頭辞（w: など）。"""
    m = re.search(rb'xmlns:([A-Za-z0-9_\-]+)="' + re.escape(ns.encode()) + rb'"', xml)
    if m:
        return m[1].decode()
    for strict, tr in STRICT.items():
        if tr == ns:
            m = re.search(rb'xmlns:([A-Za-z0-9_\-]+)="' + re.escape(strict.encode()) + rb'"', xml)
            if m:
                return m[1].decode()
    return "w"
