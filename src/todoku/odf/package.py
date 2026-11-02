"""OpenDocument のパッケージ（ODF 1.3 第2部）：ZIP の部品（content.xml・styles.xml・settings.xml・META-INF/manifest.xml）と、
1つの XML にまとめた形（.fodt・.fods。office:document）を、同じ問い合わせで読む。

書き込みは、変えた部品だけを ZIP の中で置き換える（ほかの部品は、圧縮されたバイト列のまま。zipraw.RawZip）。
"""

from __future__ import annotations

import re

from ..zipraw import RawZip
from .xml import NS, El, XmlError, parse, q, top

MIME_KIND = {"application/vnd.oasis.opendocument.text": "odt", "application/vnd.oasis.opendocument.text-template": "odt",
             "application/vnd.oasis.opendocument.text-master": "odt",
             "application/vnd.oasis.opendocument.spreadsheet": "ods",
             "application/vnd.oasis.opendocument.spreadsheet-template": "ods"}
FLAT = "(1つの XML)"


class OdfError(ValueError):
    pass


class OdfEncrypted(OdfError):
    """パスワードで暗号化された OpenDocument（manifest の encryption-data）。"""


class OdfPackage:
    def __init__(self, blob: bytes, filename: str = ""):
        self.blob = blob
        self.filename = filename
        self.flat = blob[:2] != b"PK"
        self.trees: dict[str, tuple[bytes, El]] = {}
        if self.flat:
            self.zip = None
            self.mimetype = ""
            root = self._parse(FLAT, blob)
            doc = top(root)
            self.mimetype = doc.get("office:mimetype", "")
            self.encrypted: list[str] = []
            self.names = [FLAT]
        else:
            try:
                self.zip = RawZip(blob)
            except Exception as e:  # noqa: BLE001
                raise OdfError(f"ZIP として読めません（{e}）") from None
            self.names = self.zip.names()
            self.mimetype = self.zip.read("mimetype").decode("ascii", "replace").strip() if self.zip.has("mimetype") else ""
            self.encrypted = []
            if self.zip.has("META-INF/manifest.xml"):
                man = self.zip.read("META-INF/manifest.xml")
                for m in re.finditer(rb"<manifest:file-entry\b(?:[^>\"']|\"[^\"]*\"|'[^']*')*?(?:/>|>(.*?)</manifest:file-entry>)", man, re.S):
                    if m[1] and b"encryption-data" in m[1]:
                        p = re.search(rb'full-path="([^"]*)"', m[0])
                        if p:
                            self.encrypted.append(p[1].decode("utf-8", "replace"))
                if not self.mimetype:
                    mm = re.search(rb'full-path="/"[^>]*?media-type="([^"]+)"|media-type="([^"]+)"[^>]*?full-path="/"', man)
                    if mm:
                        self.mimetype = (mm[1] or mm[2]).decode("ascii", "replace")
            if "content.xml" in self.encrypted:
                raise OdfEncrypted("パスワードで暗号化された OpenDocument です（いまは開けません。本人に、暗号化しないで保存した"
                                   "ファイルを頼んでください）")
        self.kind = MIME_KIND.get(self.mimetype, "")

    def _parse(self, name: str, data: bytes) -> El:
        try:
            root = parse(data)
        except (XmlError, UnicodeDecodeError) as e:
            raise OdfError(f"{name} を XML として読めません（{e}）") from None
        self.trees[name] = (data, root)
        return root

    # ------------------------------------------------------------------ 部品
    def has(self, name: str) -> bool:
        return self.flat and name == FLAT or (not self.flat and self.zip.has(name))

    def tree(self, name: str) -> El | None:
        """部品の根（1つの XML の形では、どの部品の名前でも、その XML）。"""
        if self.flat:
            return self.trees[FLAT][1]
        if name in self.trees:
            return self.trees[name][1]
        if not self.zip.has(name):
            return None
        return self._parse(name, self.zip.read(name))

    def xml(self, name: str) -> bytes:
        if self.flat:
            return self.trees[FLAT][0]
        if name not in self.trees:
            self.tree(name)
        return self.trees[name][0]

    def part_of(self, name: str) -> str:
        """書き換える部品の名前（1つの XML の形では、いつも同じ）。"""
        return FLAT if self.flat else name

    def section(self, part: str, local: str) -> El | None:
        """部品の文書の要素の下の office:…（automatic-styles・styles・master-styles・body・settings など）。"""
        root = self.tree(part)
        if root is None:
            return None
        doc = top(root)
        return doc.find("office:" + local)

    def body(self) -> El:
        b = self.section("content.xml", "body")
        if b is None:
            raise OdfError("文書の本体（office:body）がありません")
        return b

    def settings(self) -> dict[str, str]:
        """settings.xml の config-item（名前 → 値）。文書の設定（フォームの保護など）を読む。"""
        part = "settings.xml"
        sec = self.section(part, "settings") if (self.flat or (self.zip is not None and self.zip.has(part))) else None
        out: dict[str, str] = {}
        if sec is None:
            return out
        data = self.xml(part)
        for it in sec.iter("config:config-item"):
            name = it.get("config:name", "")
            if name and name not in out:
                out[name] = "".join(k.value(data) for k in it.kids if not isinstance(k, El))
        return out

    @property
    def has_macros(self) -> bool:
        if self.flat:
            return bool(re.search(rb"<office:scripts\b[^>]*>.*?<office:script\b", self.blob[:2_000_000], re.S))
        return any(n.startswith(("Basic/", "Scripts/")) for n in self.names)

    # ------------------------------------------------------------------ 書く
    def rebuilt(self, changed: dict[str, bytes]) -> bytes:
        """変えた部品だけを置き換えたファイル。"""
        if self.flat:
            return changed.get(FLAT, self.blob)
        z = RawZip(self.blob)
        for name, data in changed.items():
            z.replace(name, data)
        return z.build() if z.modified_names() else self.blob


def text_of(el: El, data: bytes) -> str:
    """要素の中の字を、そのままつなぐ（空白の正規化はしない。説明・名前など、段落でない所の字に使う）。"""
    out = []
    stack = [el]
    while stack:
        n = stack.pop()
        if isinstance(n, El):
            if n.tag == q("text:s"):
                out.append(" " * int(n.get("text:c", "1") or 1))
            elif n.tag == q("text:tab"):
                out.append("\t")
            elif n.tag == q("text:line-break"):
                out.append("\n")
            stack.extend(reversed(n.kids))
        else:
            out.append(n.value(data))
    return "".join(out)


__all__ = ["NS", "OdfPackage", "OdfError", "OdfEncrypted", "text_of", "FLAT", "MIME_KIND"]
