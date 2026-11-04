"""EPUB（.epub、OCF＋OPF＋XHTML）：本文の字を読む（読むだけ）。

- META-INF/container.xml → OPF（rootfile）→ spine の順に、manifest の項目の XHTML を読む。字の取り出しは markup（HTML の字）
- 暗号（META-INF/encryption.xml で本文の項目が暗号化）は読めないと伝える（フォントの難読化だけなら本文は読める）
"""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from posixpath import dirname, join, normpath
from urllib.parse import unquote

from .. import readonlydoc as rd
from ..zipraw import RawZip

KIND_JA = "EPUB（.epub。読むだけ）"
WRITE = "EPUB には書きません。電子書籍の形で、入力欄がありません"
MIME = "application/epub+zip"


class EpubError(ValueError):
    pass


def is_epub(z: RawZip, names: set) -> bool:
    if "mimetype" not in names:
        return False
    try:
        return z.read("mimetype").decode("ascii", "replace").strip() == MIME
    except Exception:  # noqa: BLE001
        return False


def _ln(t: str) -> str:
    return t.rsplit("}", 1)[-1]


def spine(z: RawZip) -> list[str]:
    names = set(z.names())
    if "META-INF/container.xml" not in names:
        raise EpubError("META-INF/container.xml がありません")
    try:
        root = ET.fromstring(z.read("META-INF/container.xml"))
        rf = next((e.get("full-path") for e in root.iter() if _ln(e.tag) == "rootfile"), None)
        if not rf or rf not in names:
            raise EpubError("OPF の部品がありません")
        opf = ET.fromstring(z.read(rf))
    except ET.ParseError as e:
        raise EpubError(f"EPUB の案内の部品を読めません（{e}）") from None
    items = {i.get("id"): i for i in opf.iter() if _ln(i.tag) == "item"}
    out = []
    for r in opf.iter():
        if _ln(r.tag) == "itemref":
            it = items.get(r.get("idref"))
            if it is not None and "html" in (it.get("media-type") or "") or (it is not None and (it.get("href") or "").lower().endswith((".html", ".xhtml", ".htm"))):
                out.append(normpath(join(dirname(rf), unquote(it.get("href") or ""))))
    return out


def _encrypted(z: RawZip) -> list[str]:
    if "META-INF/encryption.xml" not in set(z.names()):
        return []
    try:
        root = ET.fromstring(z.read("META-INF/encryption.xml"))
    except ET.ParseError:
        return []
    return [unquote(e.get("URI") or "") for e in root.iter() if _ln(e.tag) == "CipherReference"]


def build(blob: bytes, name: str) -> dict:
    from .. import markup

    z = RawZip(blob)
    names = set(z.names())
    enc = set(_encrypted(z))
    secs, unread = [], []
    for k, part in enumerate(spine(z), 1):
        if part in enc:
            unread.append(f"{part}：暗号化された本文は読めません（DRM）")
            continue
        if part not in names:
            unread.append(f"{part}：部品がありません")
            continue
        win = markup.build(z.read(part), part.rsplit("/", 1)[-1] + (".html" if not part.lower().endswith((".html", ".htm", ".xhtml")) else ""))
        secs.append((f"{k}. {part}", list(win["lines"])))
    return rd.make("epub", KIND_JA, name, blob, secs, "元のファイルが正本。目次の順（spine）に本文の字を読みました（字の飾り・絵は読みません）。" + WRITE, unread)


def grade(blob: bytes, name: str):
    from ..grade import _locked

    try:
        win = build(blob, name)
    except (EpubError, ValueError) as e:
        return _locked("読めない", str(e), "読めないので、書けません")
    if not win["sections"]:
        return _locked("字のない EPUB", "読める本文がありません" + ("：" + "／".join(win["unread"][:3]) if win["unread"] else ""), "読めないので、書けません")
    return rd.grade(win, "B", "EPUB は字を読む", "本文の XHTML の字を読みます（表の形・字の大きさ・リンクの先は読みません）", "EPUB には書かない", WRITE)


def audit(blob: bytes, name: str):
    import html

    from ..xlsx.audit import Audit, Check

    try:
        win = build(blob, name)
        z = RawZip(blob)
        parts = spine(z)
    except (EpubError, ValueError) as e:
        au = Audit(name)
        au.checks.append(Check("読む", False, str(e)))
        return au
    raw = []
    enc = set(_encrypted(z))
    names = set(z.names())
    for part in parts:
        if part in enc or part not in names:
            continue
        t = z.read(part).decode("utf-8-sig", "replace")
        t = re.sub(r"(?is)<(script|style|template|noscript|head)\b.*?</\1\s*>", " ", t)
        t = re.sub(r"(?s)<!--.*?-->", " ", t)
        t = re.sub(r"(?s)<!\[CDATA\[(.*?)\]\]>", r"\1", t)
        raw.append(html.unescape(re.sub(r"(?s)<[^>]*>", " ", t)))
    return rd.audit(win, name, raw, f"本文の部品 {len(parts)} 個", WRITE)
