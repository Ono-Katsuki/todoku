"""Office（OPC）・OpenDocument・OCF（HWPX など）のパッケージの電子署名を見つける。

電子署名は、署名したときの部品のバイト列（のハッシュ）に結び付いている。とどくは書いた部品だけを書き換え、署名の部品は
そのまま残すので、書いた後のファイルの署名は「改ざんされた」ものとして無効になる。黙って無効にしないよう、書き込みの結果・
段階・確かめ（audit）で伝える。

- OPC（.docx・.xlsx・.pptx など）：パッケージの関係（_rels/.rels）の digital-signature/origin、または _xmlsignatures/ の部品
  （ECMA-376 第2部 13）
- OpenDocument：META-INF/documentsignatures.xml（ODF 1.2 第3部 4.2）
- OCF（HWPX・EPUB）：META-INF/signatures.xml
マクロ（VBA）だけの署名（vbaProjectSignature.bin・ODF の META-INF/macrosignatures.xml）は、マクロのプロジェクトだけに
かかり、セル・本文に書いても無効にならないので、ここでは数えない。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

ORIGIN_REL = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"

MESSAGE = "この書類には電子署名があります。書き込むと署名は無効になり、もう一度署名が要ります"


@dataclass
class PackageSignature:
    kind: str  # opc / odf / ocf
    parts: list[str] = field(default_factory=list)  # 署名の部品
    count: int = 0  # 署名の数（分かるとき）

    @property
    def message(self) -> str:
        what = {"opc": "Office の電子署名", "odf": "OpenDocument の電子署名", "ocf": "パッケージの電子署名"}[self.kind]
        n = f" {self.count} 件" if self.count else ""
        return (f"{MESSAGE}（{what}{n}：{'・'.join(self.parts[:4])}）。署名した人に、書いた後のファイルに"
                "署名し直してもらってください（とどくは署名できません）")

    def to_dict(self) -> dict:
        return {"kind": self.kind, "parts": self.parts, "count": self.count, "message": self.message}


def find(blob: bytes) -> PackageSignature | None:
    """ZIP のパッケージの電子署名（なければ、ZIP でなければ None）。読めない ZIP も None（形式ごとの読み方で扱う）。"""
    if blob[:2] != b"PK":
        return None
    from .zipraw import RawZip

    try:
        z = RawZip(blob)
        names = z.names()
    except Exception:  # noqa: BLE001
        return None
    return find_in(names, z.read)


def find_in(names, read) -> PackageSignature | None:
    """ZIP の中の名前の並びと、1 つを読む関数から、パッケージの電子署名を探す（大きな書類は、全体をメモリに載せずに探す）。"""
    lower = {n.lower(): n for n in names}
    if "[content_types].xml" in lower:
        sig_parts = [n for n in names if n.lower().startswith("_xmlsignatures/") and not n.endswith("/")]
        origin = False
        if "_rels/.rels" in lower:
            try:
                rels = read(lower["_rels/.rels"])
                origin = bool(re.search(rb'Type="' + re.escape(ORIGIN_REL.encode()) + rb'"', rels))
            except Exception:  # noqa: BLE001
                origin = False
        sigs = [n for n in sig_parts if re.fullmatch(r"(?i)_xmlsignatures/sig[^/]*\.xml", n)]
        if origin or sigs:
            return PackageSignature("opc", sigs or sig_parts or ["_xmlsignatures/origin.sigs"], len(sigs))
        return None
    for name, kind in (("meta-inf/documentsignatures.xml", "odf"), ("meta-inf/signatures.xml", "ocf")):
        if name in lower:
            try:
                n = len(re.findall(rb"<(?:[\w.-]+:)?Signature\b", read(lower[name])))
            except Exception:  # noqa: BLE001
                n = 0
            return PackageSignature(kind, [lower[name]], n)
    return None
