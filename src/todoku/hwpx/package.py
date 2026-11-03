"""HWPX のパッケージ（OWPML・KS X 6101:2018 の「패키지」）：ZIP の部品を読む。

- mimetype（先頭・圧縮なし）：application/hwp+zip
- Contents/content.hpf（OPF の manifest と spine）：本文の節（Contents/section0.xml …）の順
- Contents/header.xml：文字の形（charPr：下線・囲み線）・段落の形（paraPr：番号）・番号の書式（numbering・bullet）
- META-INF/manifest.xml：暗号化した部品（odf:encryption-data。パスワードで保護した文書・配布用の文書）
- Scripts/：文書のマクロ（JScript）
書き込みは、変えた部品（本文の節の XML）だけを ZIP の中で置き換える（ほかの部品は、圧縮されたバイト列のまま。zipraw.RawZip）。
"""

from __future__ import annotations

import re

from ..odf.xml import El, XmlError, parse, top
from ..zipraw import RawZip

MIME = "application/hwp+zip"
HNS = "http://www.hancom.co.kr/hwpml/"  # 한글の名前空間（2011/paragraph・2011/section・2011/head・2016/paragraph など）


class HwpxError(ValueError):
    pass


class HwpxEncrypted(HwpxError):
    """パスワードで保護した文書・配布用の文書（部品が暗号化されている）。"""


def L(el) -> str:
    """한글の名前空間の要素の名前（hp:p → "p"）。ほかの名前空間の要素は ""。"""
    t = getattr(el, "tag", "")
    if t.startswith("{" + HNS):
        return t[t.index("}") + 1:]
    return ""


def kids(el: El, local: str | None = None) -> list[El]:
    return [k for k in el.kids if isinstance(k, El) and (local is None or L(k) == local)]


def kid(el: El, local: str) -> El | None:
    for k in el.kids:
        if isinstance(k, El) and L(k) == local:
            return k
    return None


def is_hwpx(blob: bytes) -> bool:
    """ZIP の mimetype の部品が application/hwp+zip か（なければ Contents/content.hpf と Contents/section0.xml があるか）。"""
    try:
        z = RawZip(blob)
        if z.has("mimetype"):
            return z.read("mimetype").decode("ascii", "replace").strip() == MIME
        return z.has("Contents/content.hpf") and any(re.fullmatch(r"Contents/section\d+\.xml", n) for n in z.names())
    except Exception:  # noqa: BLE001
        return False


class HwpxPackage:
    def __init__(self, blob: bytes, filename: str = ""):
        self.blob = blob
        self.filename = filename
        try:
            self.zip = RawZip(blob)
        except Exception as e:  # noqa: BLE001
            raise HwpxError(f"ZIP として読めません（{e}）") from None
        self.names = self.zip.names()
        self.mimetype = self.zip.read("mimetype").decode("ascii", "replace").strip() if self.zip.has("mimetype") else ""
        self.trees: dict[str, tuple[bytes, El]] = {}
        self.encrypted: list[str] = []
        if self.zip.has("META-INF/manifest.xml"):
            man = self.zip.read("META-INF/manifest.xml")
            for m in re.finditer(rb"<(?:\w+:)?file-entry\b(?:[^>\"']|\"[^\"]*\"|'[^']*')*?(?:/>|>(.*?)</(?:\w+:)?file-entry>)", man, re.S):
                if m[1] and b"encryption-data" in m[1]:
                    p = re.search(rb'full-path="([^"]*)"', m[0])
                    if p:
                        self.encrypted.append(p[1].decode("utf-8", "replace"))
        self.sections = self._sections()
        if not self.sections:
            raise HwpxError("本文の節（Contents/section0.xml）がありません")
        if any(s in self.encrypted for s in self.sections) or any(n.lower().endswith(".xml") and n in self.encrypted for n in self.names):
            raise HwpxEncrypted("部品が暗号化された HWPX です（パスワードで保護した文書か、配布用の文書）。いまは開けません。"
                                "本人に、暗号化しないで保存した HWPX を頼んでください")

    def _sections(self) -> list[str]:
        """本文の節の部品の名前（content.hpf の spine の順。なければ番号の順）。"""
        found = [n for n in self.names if re.fullmatch(r"Contents/section\d+\.xml", n)]
        order: list[str] = []
        if self.zip.has("Contents/content.hpf"):
            try:
                root = parse(self.zip.read("Contents/content.hpf"))
                items = {}
                for el in top(root).iter():
                    if el.local == "item" and el.attrs.get("id") and el.attrs.get("href"):
                        items[el.attrs["id"]] = el.attrs["href"].lstrip("/")
                for el in top(root).iter():
                    if el.local == "itemref":
                        h = items.get(el.attrs.get("idref", ""))
                        if h in found and h not in order:
                            order.append(h)
            except (XmlError, UnicodeDecodeError):
                order = []
        rest = sorted((n for n in found if n not in order), key=lambda n: int(re.search(r"(\d+)", n)[1]))
        return order + rest

    # ------------------------------------------------------------------ 部品
    def has(self, name: str) -> bool:
        return self.zip.has(name)

    def tree(self, name: str) -> El | None:
        if name in self.trees:
            return self.trees[name][1]
        if not self.zip.has(name):
            return None
        data = self.zip.read(name)
        try:
            root = parse(data)
        except (XmlError, UnicodeDecodeError) as e:
            raise HwpxError(f"{name} を XML として読めません（{e}）") from None
        self.trees[name] = (data, root)
        return root

    def xml(self, name: str) -> bytes:
        if name not in self.trees:
            self.tree(name)
        return self.trees[name][0]

    def scripts(self) -> list[tuple[str, str]]:
        """文書のマクロ（Scripts/ の部品の JScript）：(名前, ソース)。"""
        out = []
        for n in self.names:
            if n.startswith("Scripts/") and not n.endswith("/"):
                raw = self.zip.read(n)
                if raw[:2] in (b"\xff\xfe", b"\xfe\xff"):
                    enc = "utf-16"
                elif len(raw) > 1 and raw[1:2] == b"\x00":
                    enc = "utf-16-le"  # 한글は BOM のない UTF-16（リトルエンディアン）で保存する
                else:
                    enc = "utf-8"
                out.append((n, raw.decode(enc, "replace")))
        return out

    # ------------------------------------------------------------------ 書く
    def rebuilt(self, changed: dict[str, bytes]) -> bytes:
        z = RawZip(self.blob)
        for name, data in changed.items():
            z.replace(name, data)
        return z.build() if z.modified_names() else self.blob


def script_has_code(src: str) -> bool:
    """JScript のソースに、中身のある処理があるか（コメントと、中身の空の function の定義を除いて、字が残るか）。
    한글は、マクロのない文書にも空の手続き（function OnDocument_New() { }）を保存する。"""
    s = re.sub(r"/\*.*?\*/", "", src, flags=re.S)
    s = re.sub(r"//[^\n]*", "", s)
    s = re.sub(r"\bfunction\s+[\w$]+\s*\([^)]*\)\s*\{\s*\}", "", s)
    # 한글が書く、部品を名前で指す宣言（var X = Document.….ItemFromName("X");）：処理ではない
    s = re.sub(r"\bvar\s+[\w$]+\s*=\s*[\w$.]+(?:\(\s*\"[^\"]*\"\s*\))?\s*;", "", s)
    return bool(s.strip(" \t\r\n﻿\x00;"))


__all__ = ["HwpxPackage", "HwpxError", "HwpxEncrypted", "L", "kids", "kid", "is_hwpx", "MIME", "HNS", "top", "script_has_code"]
