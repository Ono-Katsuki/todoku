"""メール（.eml、RFC 5322・MIME。.msg、MS-OXMSG）：差出人・宛先・件名・本文を読み、添付ファイルを中のファイルとして指す（読むだけ）。

- .eml：標準の email で解く。本文は text/plain（なければ text/html の字）。charset は宣言のとおり（iso-2022-jp など）。添付は「a.eml!/名前」
- .msg：複合ファイルの __substg1.0_XXXXYYYY のストリーム（XXXX がプロパティ、YYYY が型。001F は UTF-16、001E は 8 ビットの字、0102 はバイナリ）。
  本文は PR_BODY、なければ PR_HTML、なければ PR_RTF_COMPRESSED（MS-OXRTFCP の LZ を解き、RTF の部品で字にする）。
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
    rawnames: dict = field(default_factory=dict)  # 添付の名前 → 文字コードの印のない名前のバイト（.msg の 8 ビットの字）
    text_encoding: dict = field(default_factory=dict)  # .msg の 8 ビットの字の読み方（namecode.Reading.info()）


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


# MS-OXRTFCP 2.1.2.1：LZ の辞書の初めの字（207 字）
_RTF_DICT = (b"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier"
             b"{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx")
_MAX_RTF = 64 * 1024 * 1024


def _crc_table():
    out = []
    for i in range(256):
        c = i
        for _ in range(8):
            c = (c >> 1) ^ 0xEDB88320 if c & 1 else c >> 1
        out.append(c)
    return out


def rtf_decompress(d: bytes) -> bytes:
    """PR_RTF_COMPRESSED（MS-OXRTFCP）：16 バイトの見出し（COMPSIZE・RAWSIZE・COMPTYPE・CRC）と、制御バイトつきの LZ（4 KB の輪の辞書）。"""
    if len(d) < 16:
        raise MailError("圧縮した RTF の見出しが短すぎます")
    comp, raw, typ, crc = struct.unpack_from("<IIII", d, 0)
    if typ == 0x414C454D:  # MELA：圧縮していない
        return d[16:16 + raw]
    if typ != 0x75465A4C:  # LZFu
        raise MailError("圧縮した RTF の形（COMPTYPE）が読めません")
    if raw > _MAX_RTF:
        raise MailError("圧縮した RTF が大きすぎます")
    end = min(len(d), comp + 4)
    t = _crc_table()
    c = 0
    for b in d[16:end]:
        c = t[(c ^ b) & 255] ^ (c >> 8)
    if c != crc:
        raise MailError("圧縮した RTF の CRC が合いません")
    dic = bytearray(_RTF_DICT) + bytearray(4096 - len(_RTF_DICT))
    wp = len(_RTF_DICT)
    out = bytearray()
    p = 16
    while p < end:
        ctl = d[p]
        p += 1
        for bit in range(8):
            if p >= end:
                break
            if ctl >> bit & 1:
                if p + 2 > end:
                    raise MailError("圧縮した RTF が途中で終わっています")
                v = d[p] << 8 | d[p + 1]
                p += 2
                off, ln = v >> 4, (v & 15) + 2
                if off == wp:  # 終わりの印
                    return bytes(out)
                for k in range(ln):
                    ch = dic[(off + k) & 4095]
                    out.append(ch)
                    dic[wp] = ch
                    wp = (wp + 1) & 4095
            else:
                ch = d[p]
                p += 1
                out.append(ch)
                dic[wp] = ch
                wp = (wp + 1) & 4095
            if len(out) > raw + 4096:
                raise MailError("圧縮した RTF が宣言より長くなっています")
    return bytes(out)


def rtf_text(rtf: bytes) -> str:
    """RTF の字（\\fromhtml の包みも、RTF として読めば本文の字になる。{\\*\\htmltag} のタグは字にしない）。"""
    from ..rtf.document import RtfDocument, paragraphs

    return "\n".join(p.text for p in paragraphs(RtfDocument(rtf, "本文.rtf")))


_MESSAGE_CP, _INTERNET_CP = 0x3FFD, 0x3FDE  # PR_MESSAGE_CODEPAGE（8 ビットの字のプロパティ）・PR_INTERNET_CPID（本文の HTML）


def _codepages(c) -> dict:
    """__properties_version1.0（MS-OXMSG 2.4）の固定長の値の並びから、コードページのプロパティを引く：{プロパティ: codec}。
    見出し 32 バイト、1 つ 16 バイト。Python が表を持たない番号は入れない（ほかの番号の表で読まない）。"""
    from ..namecode import codec_of_codepage

    n = "__properties_version1.0"
    out = {}
    if not c.has(n):
        return out
    d = c.read(n)
    for pos in range(32, len(d) - 15, 16):
        typ, pid, _flags, val = struct.unpack_from("<HHII", d, pos)
        if typ == 3 and pid in (_MESSAGE_CP, _INTERNET_CP):
            codec = codec_of_codepage(val)
            if codec and not codec.startswith(("utf-16", "utf-32")):
                out[pid] = codec
    return out


def _html_decode(h: bytes, cps: dict, notes: list) -> str:
    """本文の HTML（バイト）を字にする：PR_INTERNET_CPID → PR_MESSAGE_CODEPAGE の順に、厳密に読めた文字コード。
    どちらもなければ、文字のファイルと同じ決め方（datafile/textenc：BOM・UTF-8 の形・meta の charset・読める文字コードの全部）。"""
    from ..datafile import textenc

    for pid in (_INTERNET_CP, _MESSAGE_CP):
        if pid in cps:
            try:
                return h.decode(cps[pid])
            except UnicodeDecodeError:
                continue
    dec = textenc.decode(h)
    notes.extend(x.lstrip("- ") if x.startswith("- ") else x for x in dec.note_lines())
    return dec.text


def read_msg(blob: bytes, encoding: str | None = None) -> Mail:
    """Outlook の .msg。字のプロパティは、UTF-16（001F）があればそれ。8 ビットの字（001E）は、PR_MESSAGE_CODEPAGE（なければ
    PR_INTERNET_CPID）で厳密に読めれば、それ。コードページのプロパティがないか、読めなければ、namecode の決め方（決められなければ、
    候補を notes に並べる）。encoding：エージェントが選んだ、8 ビットの字の文字コード。"""
    from .. import namecode
    from ..cfb import CFB

    c = CFB(blob)
    cps = _codepages(c)
    wide: dict = {}
    narrow: dict = {}

    def load(base: str, tag: int) -> None:
        n = f"{base}__substg1.0_{tag:04X}"
        if c.has(n + "001F"):
            wide[(base, tag)] = c.read(n + "001F").decode("utf-16-le", "replace").rstrip("\x00")
        elif c.has(n + "001E"):
            narrow[(base, tag)] = c.read(n + "001E").rstrip(b"\x00")

    head_tags = (("From", 0x0C1A), ("差出人のアドレス", 0x0C1F), ("To", 0x0E04), ("Cc", 0x0E03), ("Subject", 0x0037))
    stores = sorted({n.split("/")[0] for n in c.streams() if n.lower().startswith("__attach_version1.0_#")})
    for _key, tag in head_tags + (("", 0x1000),):
        load("", tag)
    for st in stores:
        for tag in (0x3707, 0x3704):
            load(st + "/", tag)
    declared = [(cps[pid], what) for pid, what in ((_MESSAGE_CP, "メールのコードページ PR_MESSAGE_CODEPAGE"),
                                                   (_INTERNET_CP, "メールのコードページ PR_INTERNET_CPID")) if pid in cps]
    reading = namecode.decide(list(narrow.values()), encoding, declared)
    seven = next((e for e, _w in declared if e.startswith("iso2022")), None)

    def text(base: str, tag: int) -> str | None:
        if (base, tag) in wide:
            return wide[(base, tag)]
        raw = narrow.get((base, tag))
        if raw is None:
            return None
        if seven and b"\x1b" in raw and raw.isascii():  # ISO-2022 の字（7 ビット。エスケープで切り替える）
            try:
                return raw.decode(seven)
            except UnicodeDecodeError:
                pass
        return reading.text(raw)[0]

    heads = []
    for key, tag in head_tags:
        v = text("", tag)
        if v:
            heads.append((key, v))
    body = text("", 0x1000) or ""
    notes = [x[2:] if x.startswith("- ") else x for x in reading.note_lines(
        "8 ビットの字の項目（差出人・件名・本文・添付の名前）", "選んだら、その文字コードの名前を encoding で渡してください。")]
    if not body:
        h = c.read("__substg1.0_10130102") if c.has("__substg1.0_10130102") else None
        if h:
            body = _html_text(_html_decode(h, cps, notes))
            notes.append("本文は HTML だけなので、タグの間の字を読みました")
        elif c.has("__substg1.0_10090102"):
            try:
                body = rtf_text(rtf_decompress(c.read("__substg1.0_10090102")))
                notes.append("本文は RTF（PR_RTF_COMPRESSED）だけなので、RTF を解いて字を読みました")
            except (MailError, ValueError, IndexError) as e:
                notes.append(f"本文の圧縮した RTF（PR_RTF_COMPRESSED）を解けません（{e}）")
    atts, names, rawnames = [], [], {}
    for st in stores:
        base = st + "/"
        tag = 0x3707 if text(base, 0x3707) else 0x3704
        nm = _unique(names, text(base, tag) or "添付")
        data_n = f"{base}__substg1.0_37010102"
        if not c.has(data_n):
            notes.append(f"添付「{nm}」は埋め込みのオブジェクトで、取り出していません")
            continue
        names.append(nm)
        atts.append((nm, c.read(data_n)))
        raw = narrow.get((base, tag))
        if raw is not None and not raw.isascii():
            rawnames[nm] = raw
    return Mail(heads, body, atts, notes, rawnames, reading.info())


_HEAD_JA = {"From": "差出人", "To": "宛先", "Cc": "写し", "Date": "日時", "Subject": "件名"}


def lines_of(m: Mail) -> list[str]:
    out = [f"{_HEAD_JA.get(k, k)}：{v}" for k, v in m.headers]
    out += [""] + re.split(r"\r\n|\r|\n", m.body)
    return out


def is_msg(names: set) -> bool:
    return any(n.lower().startswith(("__properties_version1.0", "__substg1.0_")) for n in names)
