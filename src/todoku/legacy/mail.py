"""メール（.eml、RFC 5322・MIME。.msg、MS-OXMSG）：差出人・宛先・件名・本文を読み、添付ファイルを中のファイルとして指す（読むだけ）。

- .eml：標準の email で解く。本文は text/plain（なければ text/html の字）。charset は宣言のとおり（iso-2022-jp など）。添付は「a.eml!/名前」
- .msg：複合ファイルの __substg1.0_XXXXYYYY のストリーム（XXXX がプロパティ、YYYY が型。001F は UTF-16、001E は 8 ビットの字、0102 はバイナリ）。
  添付は __attach_version1.0_#NNNNNNNN の入れ物。8 ビットの字の文字コードは PR_INTERNET_CPID（0x3FDE）か PR_MESSAGE_CODEPAGE（0x3FFD）
"""

from __future__ import annotations

import email
import email.policy
import re
import struct
from dataclasses import dataclass, field


class MailError(ValueError):
    pass


@dataclass
class Mail:
    headers: list[tuple[str, str]]
    body: str
    attachments: list[tuple[str, bytes]] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def _unique(names: list[str], name: str) -> str:
    name = (name or "添付").replace("\\", "/").rsplit("/", 1)[-1] or "添付"
    if name not in names:
        return name
    stem, dot, ext = name.rpartition(".")
    k = 2
    while True:
        cand = f"{stem}({k}).{ext}" if dot else f"{name}({k})"
        if cand not in names:
            return cand
        k += 1


def _html_text(s: str) -> str:
    from ..markup import _Text

    p = _Text(False)
    p.feed(s)
    p.close()
    return "\n".join(x.strip() for x in p.lines if x.strip())


def read_eml(blob: bytes) -> Mail:
    try:
        m = email.message_from_bytes(blob, policy=email.policy.default)
    except Exception as e:  # noqa: BLE001
        raise MailError(f"メールとして読めません（{type(e).__name__}）") from None
    heads = []
    for key in ("From", "To", "Cc", "Date", "Subject"):
        v = m.get(key)
        if v is not None:
            heads.append((key, str(v)))
    plain: list[str] = []
    html_parts: list[str] = []
    att: list[tuple[str, bytes]] = []
    names: list[str] = []
    notes: list[str] = []
    for part in m.walk():
        if part.is_multipart():
            continue
        ctype = part.get_content_type()
        fn = part.get_filename()
        disp = part.get_content_disposition()
        if fn or disp == "attachment" or not ctype.startswith("text/"):
            try:
                data = part.get_payload(decode=True) or b""
            except Exception:  # noqa: BLE001
                data = b""
            nm = _unique(names, fn or f"添付.{ctype.rsplit('/', 1)[-1]}")
            names.append(nm)
            att.append((nm, data))
            continue
        try:
            text = part.get_content()
        except (LookupError, ValueError):
            raw = part.get_payload(decode=True) or b""
            text = raw.decode("latin-1")
            notes.append(f"本文の文字コード（{part.get_content_charset()}）を読めません。バイトをそのまま字にしました")
        (plain if ctype == "text/plain" else html_parts).append(text)
    if plain:
        body = "\n".join(plain)
    elif html_parts:
        body = "\n".join(_html_text(x) for x in html_parts)
        notes.append("本文は HTML だけなので、タグの間の字を読みました")
    else:
        body = ""
    return Mail(heads, body, att, notes)


def _prop_text(c, base: str, tag: int, cp: str) -> str | None:
    for typ, enc in (("001F", "utf-16-le"), ("001E", cp)):
        n = f"{base}__substg1.0_{tag:04X}{typ}"
        if c.has(n):
            return c.read(n).decode(enc, "replace").rstrip("\x00")
    return None


CODEPAGES = {932: "cp932", 65001: "utf-8", 50220: "iso-2022-jp", 51932: "euc_jp", 1252: "cp1252", 20127: "ascii"}


def _codepage(c) -> str:
    """__properties_version1.0（MS-OXMSG 2.4）の固定長の値の並びから、文字コードのプロパティを引く。見出し 32 バイト、1 つ 16 バイト。"""
    n = "__properties_version1.0"
    if not c.has(n):
        return "cp932"
    d = c.read(n)
    found = {}
    for pos in range(32, len(d) - 15, 16):
        typ, pid, _flags, val = struct.unpack_from("<HHII", d, pos)
        if typ == 3 and pid in (0x3FDE, 0x3FFD):
            found[pid] = val
    for pid in (0x3FDE, 0x3FFD):
        if found.get(pid) in CODEPAGES:
            return CODEPAGES[found[pid]]
    return "cp932"


def read_msg(blob: bytes) -> Mail:
    from ..cfb import CFB

    c = CFB(blob)
    cp = _codepage(c)
    heads = []
    for key, tag in (("From", 0x0C1A), ("差出人のアドレス", 0x0C1F), ("To", 0x0E04), ("Cc", 0x0E03), ("Subject", 0x0037)):
        v = _prop_text(c, "", tag, cp)
        if v:
            heads.append((key, v))
    body = _prop_text(c, "", 0x1000, cp) or ""
    notes = []
    if not body:
        h = c.read("__substg1.0_10130102") if c.has("__substg1.0_10130102") else None
        if h:
            body = _html_text(h.decode(cp, "replace"))
            notes.append("本文は HTML だけなので、タグの間の字を読みました")
        elif c.has("__substg1.0_10090102"):
            notes.append("本文は圧縮した RTF（PR_RTF_COMPRESSED）だけで、読んでいません")
    atts, names = [], []
    stores = sorted({n.split("/")[0] for n in c.streams() if n.lower().startswith("__attach_version1.0_#")})
    for st in stores:
        base = st + "/"
        nm = _prop_text(c, base, 0x3707, cp) or _prop_text(c, base, 0x3704, cp)
        data_n = f"{base}__substg1.0_37010102"
        nm = _unique(names, nm or "添付")
        if not c.has(data_n):
            notes.append(f"添付「{nm}」は埋め込みのオブジェクトで、取り出していません")
            continue
        names.append(nm)
        atts.append((nm, c.read(data_n)))
    return Mail(heads, body, atts, notes)


_HEAD_JA = {"From": "差出人", "To": "宛先", "Cc": "写し", "Date": "日時", "Subject": "件名"}


def lines_of(m: Mail) -> list[str]:
    out = [f"{_HEAD_JA.get(k, k)}：{v}" for k, v in m.headers]
    out += [""] + re.split(r"\r\n|\r|\n", m.body)
    return out


def is_msg(names: set) -> bool:
    return any(n.lower().startswith(("__properties_version1.0", "__substg1.0_")) for n in names)
