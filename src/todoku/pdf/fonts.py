"""PDFのフォント：文字列のバイトを、文字（Unicode）と送り幅に直す。

文字の決め方（上から順に）:
  1. ToUnicode（フォントが持つ、文字コード → 文字 の対応）
  2. 定義済みの CMap（90ms-RKSJ-H や UniJIS-UCS2-H など）：文字コードそのものが Shift_JIS や UTF-16 なので、その符号で読む
  3. 文字集合（Adobe-Japan1 など）の公式の対応表で CID から文字を決める（data/cid2unicode.json.gz、Adobe の公開資料）
  4. 埋め込まれたフォントのプログラムの cmap・グリフ名
  5. 単純なフォントは Encoding と Differences のグリフ名
どれでも決まらない文字は「読めない文字」として数え、窓と検査に出す（黙って捨てない）。
"""

from __future__ import annotations

import bisect
import gzip
import json
import re
import unicodedata
from functools import lru_cache
from pathlib import Path

from .objects import Name, PdfString, Ref, Stream

UNKNOWN = "�"

# ====================================================================== 単純なフォントの符号
STANDARD_ENCODING: list[str | None] = [None] * 256
for _i, _n in enumerate(
        "space exclam quotedbl numbersign dollar percent ampersand quoteright parenleft parenright asterisk plus comma "
        "hyphen period slash zero one two three four five six seven eight nine colon semicolon less equal greater question "
        "at A B C D E F G H I J K L M N O P Q R S T U V W X Y Z bracketleft backslash bracketright asciicircum underscore "
        "quoteleft a b c d e f g h i j k l m n o p q r s t u v w x y z braceleft bar braceright asciitilde".split()):
    STANDARD_ENCODING[0x20 + _i] = _n
for _c, _n in {
        0xA1: "exclamdown", 0xA2: "cent", 0xA3: "sterling", 0xA4: "fraction", 0xA5: "yen", 0xA6: "florin", 0xA7: "section",
        0xA8: "currency", 0xA9: "quotesingle", 0xAA: "quotedblleft", 0xAB: "guillemotleft", 0xAC: "guilsinglleft",
        0xAD: "guilsinglright", 0xAE: "fi", 0xAF: "fl", 0xB1: "endash", 0xB2: "dagger", 0xB3: "daggerdbl",
        0xB4: "periodcentered", 0xB6: "paragraph", 0xB7: "bullet", 0xB8: "quotesinglbase", 0xB9: "quotedblbase",
        0xBA: "quotedblright", 0xBB: "guillemotright", 0xBC: "ellipsis", 0xBD: "perthousand", 0xBF: "questiondown",
        0xC1: "grave", 0xC2: "acute", 0xC3: "circumflex", 0xC4: "tilde", 0xC5: "macron", 0xC6: "breve", 0xC7: "dotaccent",
        0xC8: "dieresis", 0xCA: "ring", 0xCB: "cedilla", 0xCD: "hungarumlaut", 0xCE: "ogonek", 0xCF: "caron",
        0xD0: "emdash", 0xE1: "AE", 0xE3: "ordfeminine", 0xE8: "Lslash", 0xE9: "Oslash", 0xEA: "OE", 0xEB: "ordmasculine",
        0xF1: "ae", 0xF5: "dotlessi", 0xF8: "lslash", 0xF9: "oslash", 0xFA: "oe", 0xFB: "germandbls"}.items():
    STANDARD_ENCODING[_c] = _n


def _codec_table(codec: str) -> list[str | None]:
    out: list[str | None] = []
    for c in range(256):
        try:
            out.append(bytes([c]).decode(codec))
        except UnicodeDecodeError:
            out.append(None)
    return out


WINANSI = _codec_table("cp1252")
for _c in (0x81, 0x8D, 0x8F, 0x90, 0x9D):
    WINANSI[_c] = "•"  # PDF の WinAnsiEncoding で未定義の位置は bullet
WINANSI[0xA0] = " "
WINANSI[0xAD] = "-"
MACROMAN = _codec_table("mac_roman")
MACROMAN[0xDB] = "¤"
PDFDOC = _codec_table("latin-1")

SYMBOL: dict[int, str] = {
    0x20: " ", 0x21: "!", 0x22: "∀", 0x23: "#", 0x24: "∃", 0x25: "%", 0x26: "&", 0x27: "∋", 0x28: "(", 0x29: ")",
    0x2A: "∗", 0x2B: "+", 0x2C: ",", 0x2D: "−", 0x2E: ".", 0x2F: "/", 0x3A: ":", 0x3B: ";", 0x3C: "<", 0x3D: "=",
    0x3E: ">", 0x3F: "?", 0x40: "≅", 0x5B: "[", 0x5C: "∴", 0x5D: "]", 0x5E: "⊥", 0x5F: "_", 0x60: "‾", 0x7B: "{",
    0x7C: "|", 0x7D: "}", 0x7E: "∼", 0xA0: "€", 0xA1: "ϒ", 0xA2: "′", 0xA3: "≤", 0xA4: "⁄", 0xA5: "∞", 0xA6: "ƒ",
    0xA7: "♣", 0xA8: "♦", 0xA9: "♥", 0xAA: "♠", 0xAB: "↔", 0xAC: "←", 0xAD: "↑", 0xAE: "→", 0xAF: "↓", 0xB0: "°",
    0xB1: "±", 0xB2: "″", 0xB3: "≥", 0xB4: "×", 0xB5: "∝", 0xB6: "∂", 0xB7: "•", 0xB8: "÷", 0xB9: "≠", 0xBA: "≡",
    0xBB: "≈", 0xBC: "…", 0xBF: "↵", 0xC0: "ℵ", 0xC1: "ℑ", 0xC2: "ℜ", 0xC3: "℘", 0xC4: "⊗", 0xC5: "⊕", 0xC6: "∅",
    0xC7: "∩", 0xC8: "∪", 0xC9: "⊃", 0xCA: "⊇", 0xCB: "⊄", 0xCC: "⊂", 0xCD: "⊆", 0xCE: "∈", 0xCF: "∉", 0xD0: "∠",
    0xD1: "∇", 0xD2: "®", 0xD3: "©", 0xD4: "™", 0xD5: "∏", 0xD6: "√", 0xD7: "⋅", 0xD8: "¬", 0xD9: "∧", 0xDA: "∨",
    0xDB: "⇔", 0xDC: "⇐", 0xDD: "⇑", 0xDE: "⇒", 0xDF: "⇓", 0xE0: "◊", 0xE1: "〈", 0xE2: "®", 0xE3: "©", 0xE4: "™",
    0xE5: "∑", 0xF1: "〉", 0xF2: "∫"}
for _i, _ch in enumerate("0123456789"):
    SYMBOL[0x30 + _i] = _ch
for _i, _ch in enumerate("ΑΒΧΔΕΦΓΗΙϑΚΛΜΝΟΠΘΡΣΤΥςΩΞΨΖ"):
    SYMBOL[0x41 + _i] = _ch
for _i, _ch in enumerate("αβχδεφγηιϕκλμνοπθρστυϖωξψζ"):
    SYMBOL[0x61 + _i] = _ch

ZAPF: dict[int, str] = {0x20: " ", 0x33: "✓", 0x34: "✔", 0x35: "✕", 0x36: "✖", 0x37: "✗", 0x38: "✘", 0x48: "★",
                        0x6C: "●", 0x6E: "■", 0x6F: "❏", 0x70: "❐", 0x71: "❑", 0x72: "❒", 0x73: "▲", 0x74: "▼",
                        0x75: "◆"}

# グリフ名 → 文字（よく使う名前。アクセント付きの文字は下の規則で作る）
_NAMES = {
    "space": " ", "exclam": "!", "quotedbl": '"', "numbersign": "#", "dollar": "$", "percent": "%", "ampersand": "&",
    "quotesingle": "'", "quoteright": "’", "quoteleft": "‘", "parenleft": "(", "parenright": ")", "asterisk": "*",
    "plus": "+", "comma": ",", "hyphen": "-", "period": ".", "slash": "/", "colon": ":", "semicolon": ";", "less": "<",
    "equal": "=", "greater": ">", "question": "?", "at": "@", "bracketleft": "[", "backslash": "\\", "bracketright": "]",
    "asciicircum": "^", "underscore": "_", "grave": "`", "braceleft": "{", "bar": "|", "braceright": "}",
    "asciitilde": "~", "quotedblleft": "“", "quotedblright": "”", "quotesinglbase": "‚", "quotedblbase": "„",
    "guillemotleft": "«", "guillemotright": "»", "guilsinglleft": "‹", "guilsinglright": "›", "endash": "–",
    "emdash": "—", "bullet": "•", "ellipsis": "…", "dagger": "†", "daggerdbl": "‡", "perthousand": "‰",
    "trademark": "™", "copyright": "©", "registered": "®", "degree": "°", "plusminus": "±", "multiply": "×",
    "divide": "÷", "minus": "−", "mu": "µ", "paragraph": "¶", "section": "§", "brokenbar": "¦", "currency": "¤",
    "cent": "¢", "sterling": "£", "yen": "¥", "Euro": "€", "euro": "€", "florin": "ƒ", "exclamdown": "¡",
    "questiondown": "¿", "ordfeminine": "ª", "ordmasculine": "º", "logicalnot": "¬", "macron": "¯", "acute": "´",
    "cedilla": "¸", "dieresis": "¨", "circumflex": "ˆ", "tilde": "˜", "ring": "˚", "caron": "ˇ", "breve": "˘",
    "dotaccent": "˙", "hungarumlaut": "˝", "ogonek": "˛", "periodcentered": "·", "middot": "·", "onehalf": "½",
    "onequarter": "¼", "threequarters": "¾", "onesuperior": "¹", "twosuperior": "²", "threesuperior": "³", "fi": "fi",
    "fl": "fl", "ff": "ff", "ffi": "ffi", "ffl": "ffl", "dotlessi": "ı", "germandbls": "ß", "ae": "æ", "AE": "Æ", "oe": "œ",
    "OE": "Œ", "oslash": "ø", "Oslash": "Ø", "thorn": "þ", "Thorn": "Þ", "eth": "ð", "Eth": "Ð", "lslash": "ł",
    "Lslash": "Ł", "nbspace": " ", "nonbreakingspace": " ", "sfthyphen": "­", "fraction": "⁄",
    "zero": "0", "one": "1", "two": "2", "three": "3", "four": "4", "five": "5", "six": "6", "seven": "7", "eight": "8",
    "nine": "9", "dcroat": "đ", "Dcroat": "Đ", "lozenge": "◊", "notequal": "≠", "infinity": "∞", "lessequal": "≤",
    "greaterequal": "≥", "partialdiff": "∂", "summation": "∑", "product": "∏", "integral": "∫", "radical": "√",
    "approxequal": "≈", "Delta": "∆", "Omega": "Ω", "pi": "π", "arrowleft": "←", "arrowright": "→", "arrowup": "↑",
    "arrowdown": "↓", "check": "✓", "checkmark": "✓", "a20": "✔", "a19": "✓", "a71": "●", "a73": "■", "H22073": "□",
    "H18533": "●", "H18543": "▪", "H18551": "▫", "filledbox": "■", "whitesquare": "□", "circle": "○", "blackcircle": "●",
    "lamda": "λ", "Lamda": "Λ", "franc": "₣", "Idotaccent": "İ", "kgreenlandic": "ĸ", "napostrophe": "ŉ", "longs": "ſ",
}
_ACCENTS = {"acute": "ACUTE", "grave": "GRAVE", "circumflex": "CIRCUMFLEX", "dieresis": "DIAERESIS", "tilde": "TILDE",
            "ring": "RING ABOVE", "cedilla": "CEDILLA", "caron": "CARON", "macron": "MACRON", "breve": "BREVE",
            "ogonek": "OGONEK", "dotaccent": "DOT ABOVE", "hungarumlaut": "DOUBLE ACUTE", "commaaccent": "CEDILLA",
            "slash": "STROKE"}


@lru_cache(maxsize=4096)
def glyph_to_unicode(name: str) -> str:
    """グリフ名から文字を決める（Adobe Glyph List の規則：uniXXXX、uXXXX、合字の _、.以降の飾りの名前）。"""
    if not name:
        return ""
    base = name.split(".")[0]
    if "_" in base and base not in _NAMES:
        parts = [glyph_to_unicode(p) for p in base.split("_")]
        return "".join(parts) if all(parts) else ""
    if base in _NAMES:
        return _NAMES[base]
    m = re.fullmatch(r"uni((?:[0-9A-F]{4})+)", base)
    if m:
        s = m[1]
        try:
            return "".join(chr(int(s[i:i + 4], 16)) for i in range(0, len(s), 4))
        except ValueError:
            return ""
    m = re.fullmatch(r"u([0-9A-F]{4,6})", base)
    if m:
        v = int(m[1], 16)
        return chr(v) if v <= 0x10FFFF else ""
    if len(base) == 1 and base.isascii() and base.isalpha():
        return base
    for suf, acc in _ACCENTS.items():
        if base.endswith(suf) and 1 <= len(base) - len(suf) <= 2:
            letter = base[: -len(suf)]
            case = "CAPITAL" if letter[0].isupper() else "SMALL"
            try:
                return unicodedata.lookup(f"LATIN {case} LETTER {letter.upper()} WITH {acc}")
            except KeyError:
                pass
    for case, word in (("SMALL", base), ("CAPITAL", base)):
        if case == "CAPITAL" and not base[:1].isupper():
            continue
        try:
            return unicodedata.lookup(f"GREEK {case} LETTER {word.upper()}")
        except KeyError:
            pass
    try:
        return unicodedata.lookup(base.replace("-", " "))
    except KeyError:
        return ""


# ====================================================================== CMap
class CMap:
    """埋め込み・ToUnicode の CMap（codespacerange、bfchar・bfrange、cidchar・cidrange）。"""

    def __init__(self):
        self.codespace: list[tuple[bytes, bytes]] = []
        self.uni: dict[bytes, str] = {}
        self.uni_ranges: list[tuple[int, int, int, int, bytes]] = []  # (lo, hi, 長さ, 始まりの値, 始まりの文字)
        self.cid: dict[bytes, int] = {}
        self.cid_ranges: list[tuple[int, int, int, int]] = []  # (lo, hi, 長さ, 始まりの CID)
        self.wmode = 0
        self.name = ""
        self._lens: list[int] | None = None

    @classmethod
    def parse(cls, data: bytes, resolver=None) -> "CMap":
        cm = cls()
        cm._parse(data, resolver)
        cm.uni_ranges.sort()
        cm.cid_ranges.sort()
        return cm

    def _parse(self, data: bytes, resolver) -> None:
        text = data
        m = re.search(rb"/CMapName\s*/([^\s/]+)", text)
        if m:
            self.name = m[1].decode("latin-1")
        m = re.search(rb"/WMode\s+(\d)", text)
        if m:
            self.wmode = int(m[1])
        m = re.search(rb"/([^\s/]+)\s+usecmap", text)
        if m and resolver is not None:
            parent = resolver(m[1].decode("latin-1"))
            if parent is not None:
                self.codespace += parent.codespace
                self.uni.update(parent.uni)
                self.uni_ranges += parent.uni_ranges
                self.cid.update(parent.cid)
                self.cid_ranges += parent.cid_ranges
        for block in re.findall(rb"begincodespacerange(.*?)endcodespacerange", text, re.S):
            for lo, hi in re.findall(rb"<([0-9A-Fa-f\s]+)>\s*<([0-9A-Fa-f\s]+)>", block):
                lo_b, hi_b = _hb(lo), _hb(hi)
                if len(lo_b) == len(hi_b) and lo_b:
                    self.codespace.append((lo_b, hi_b))
        for block in re.findall(rb"beginbfchar(.*?)endbfchar", text, re.S):
            for src, dst in re.findall(rb"<([0-9A-Fa-f\s]*)>\s*(<[0-9A-Fa-f\s]*>|/[^\s<>\[\]/]+)", block):
                self.uni[_hb(src)] = _dst(dst)
        for block in re.findall(rb"beginbfrange(.*?)endbfrange", text, re.S):
            for m2 in re.finditer(rb"<([0-9A-Fa-f\s]*)>\s*<([0-9A-Fa-f\s]*)>\s*(<[0-9A-Fa-f\s]*>|\[[^\]]*\])", block):
                lo_b, hi_b = _hb(m2[1]), _hb(m2[2])
                lo, hi = int.from_bytes(lo_b, "big"), int.from_bytes(hi_b, "big")
                dst = m2[3]
                n = len(lo_b)
                if dst.startswith(b"["):
                    for i, d in enumerate(re.findall(rb"<([0-9A-Fa-f\s]*)>", dst)):
                        if lo + i > hi:
                            break
                        self.uni[(lo + i).to_bytes(n, "big")] = _utf16(_hb(d))
                elif hi >= lo:
                    raw = _hb(dst[1:-1])
                    if hi - lo < 256:
                        base = int.from_bytes(raw, "big") if raw else 0
                        for i in range(hi - lo + 1):
                            self.uni[(lo + i).to_bytes(n, "big")] = _utf16(_inc(raw, i))
                    else:
                        self.uni_ranges.append((lo, hi, n, 0, raw))
        for block in re.findall(rb"begincidchar(.*?)endcidchar", text, re.S):
            for src, cid in re.findall(rb"<([0-9A-Fa-f\s]*)>\s*(\d+)", block):
                self.cid[_hb(src)] = int(cid)
        for block in re.findall(rb"begincidrange(.*?)endcidrange", text, re.S):
            for lo, hi, cid in re.findall(rb"<([0-9A-Fa-f\s]*)>\s*<([0-9A-Fa-f\s]*)>\s*(\d+)", block):
                lo_b, hi_b = _hb(lo), _hb(hi)
                self.cid_ranges.append((int.from_bytes(lo_b, "big"), int.from_bytes(hi_b, "big"), len(lo_b), int(cid)))

    def lengths(self) -> list[int]:
        if self._lens is None:
            self._lens = sorted({len(lo) for lo, _ in self.codespace}) or [1]
        return self._lens

    def split(self, s: bytes) -> list[bytes]:
        """文字列を文字コードに分ける（codespacerange に合う長さで）。"""
        lens = self.lengths()
        if lens == [2] and len(self.codespace) <= 1:
            return [s[i:i + 2] for i in range(0, len(s) - 1, 2)]
        out = []
        i = 0
        n = len(s)
        while i < n:
            took = 0
            for ln in lens:
                c = s[i:i + ln]
                if len(c) < ln:
                    break
                for lo, hi in self.codespace:
                    if len(lo) == ln and all(lo[k] <= c[k] <= hi[k] for k in range(ln)):
                        took = ln
                        break
                if took:
                    break
            if not took:
                # どの範囲にも合わない：最初のバイトが合う範囲の長さ、それもなければ1バイト
                for lo, hi in self.codespace:
                    if lo[0] <= s[i] <= hi[0]:
                        took = len(lo)
                        break
                took = took or lens[0]
            out.append(s[i:i + took])
            i += took
        return out

    def to_unicode(self, code: bytes) -> str | None:
        u = self.uni.get(code)
        if u is not None:
            return u
        if self.uni_ranges:
            v = int.from_bytes(code, "big")
            k = bisect.bisect_right(self.uni_ranges, (v, 1 << 62)) - 1
            while k >= 0:
                lo, hi, n, _b, raw = self.uni_ranges[k]
                if lo <= v <= hi and n == len(code):
                    return _utf16(_inc(raw, v - lo))
                if hi < v - 0x10000:
                    break
                k -= 1
        return None

    def to_cid(self, code: bytes) -> int | None:
        c = self.cid.get(code)
        if c is not None:
            return c
        if self.cid_ranges:
            v = int.from_bytes(code, "big")
            k = bisect.bisect_right(self.cid_ranges, (v, 1 << 62)) - 1
            while k >= 0:
                lo, hi, n, cid = self.cid_ranges[k]
                if lo <= v <= hi and n == len(code):
                    return cid + v - lo
                k -= 1
                if k >= 0 and self.cid_ranges[k][1] < v - 0x10000:
                    break
        return None


def _hb(b: bytes) -> bytes:
    h = re.sub(rb"\s", b"", b)
    if len(h) % 2:
        h += b"0"
    return bytes.fromhex(h.decode("latin-1"))


def _inc(raw: bytes, i: int) -> bytes:
    if not raw:
        return raw
    v = int.from_bytes(raw, "big") + i
    try:
        return v.to_bytes(len(raw), "big")
    except OverflowError:
        return raw


def _utf16(b: bytes) -> str:
    if len(b) == 1:
        return chr(b[0])
    try:
        return b.decode("utf-16-be")
    except UnicodeDecodeError:
        return b.decode("utf-16-be", errors="replace")


def _dst(d: bytes) -> str:
    if d.startswith(b"/"):
        return glyph_to_unicode(d[1:].decode("latin-1"))
    return _utf16(_hb(d[1:-1]))


# ====================================================================== 定義済みの CMap
def predefined(name: str):
    """定義済みの CMap の名前 → (文字コードの符号, 分け方, 縦書きか, 比例幅か)。わからなければ None。"""
    vertical = name.endswith("-V") or name == "V"
    base = name[:-2] if name.endswith(("-H", "-V")) else ("" if name in ("H", "V") else name)
    prop = "msp" in base or ("UCS2" in base and "HW" not in base) or "UTF" in base
    if name in ("Identity-H", "Identity-V"):
        return ("identity", 2, vertical, False)
    if "RKSJ" in base:
        return ("cp932", "sjis", vertical, prop)
    if base in ("", "Add", "Ext", "NWP"):
        return ("jis", 2, vertical, False)
    if base.startswith("EUC"):
        return ("euc_jp", "euc", vertical, False)
    if base.startswith(("UniJIS", "UniGB", "UniCNS", "UniKS", "UniAKR")):
        if "UCS2" in base:
            return ("utf-16-be", 2, vertical, prop)
        if "UTF16" in base:
            return ("utf-16-be", "utf16", vertical, prop)
        if "UTF32" in base:
            return ("utf-32-be", 4, vertical, prop)
        if "UTF8" in base:
            return ("utf-8", "utf8", vertical, prop)
    if base.startswith(("GBK2K",)):
        return ("gb18030", "gb18030", vertical, False)
    if base.startswith(("GBK", "GBKp")):
        return ("gbk", "dbcs", vertical, False)
    if base.startswith(("GB-EUC", "GBpc-EUC", "GBT-EUC", "GBTpc-EUC")):
        return ("gb2312", "dbcs", vertical, False)
    if base.startswith(("GB", "GBT")):
        return ("gb2312-7bit", 2, vertical, False)
    if base.startswith("HKscs") or base.startswith("HK"):
        return ("big5hkscs", "dbcs", vertical, False)
    if base.startswith(("B5", "ETen", "ETHK")):
        return ("cp950", "dbcs", vertical, False)
    if base.startswith("KSCms-UHC"):
        return ("cp949", "dbcs", vertical, False)
    if base.startswith(("KSC-EUC", "KSCpc-EUC")):
        return ("euc_kr", "dbcs", vertical, False)
    if base.startswith("KSC-Johab"):
        return ("johab", "dbcs", vertical, False)
    if base.startswith("KSC"):
        return ("ksc-7bit", 2, vertical, False)
    return None


def split_predefined(s: bytes, how) -> list[bytes]:
    out = []
    i = 0
    n = len(s)
    if isinstance(how, int):
        return [s[k:k + how] for k in range(0, n - how + 1, how)]
    while i < n:
        b = s[i]
        if how == "sjis":
            ln = 2 if (0x81 <= b <= 0x9F or 0xE0 <= b <= 0xFC) else 1
        elif how == "euc":
            ln = 3 if b == 0x8F else (2 if (b == 0x8E or 0xA1 <= b <= 0xFE) else 1)
        elif how == "dbcs":
            ln = 2 if b >= 0x81 else 1
        elif how == "gb18030":
            ln = (4 if i + 1 < n and 0x30 <= s[i + 1] <= 0x39 else 2) if b >= 0x81 else 1
        elif how == "utf16":
            ln = 4 if 0xD8 <= b <= 0xDB else 2
        elif how == "utf8":
            ln = 1 if b < 0xC0 else 2 if b < 0xE0 else 3 if b < 0xF0 else 4
        else:
            ln = 1
        out.append(s[i:i + ln])
        i += ln
    return out


def decode_predefined(code: bytes, codec: str) -> str | None:
    try:
        if codec == "jis":
            return bytes(b | 0x80 for b in code).decode("euc_jp")
        if codec == "gb2312-7bit":
            return bytes(b | 0x80 for b in code).decode("gb2312")
        if codec == "ksc-7bit":
            return bytes(b | 0x80 for b in code).decode("euc_kr")
        return code.decode(codec)
    except (UnicodeDecodeError, LookupError):
        return None


# ====================================================================== CID → 文字（Adobe の公式の対応表）
@lru_cache(maxsize=1)
def _cid_data() -> dict:
    p = Path(__file__).parent / "data" / "cid2unicode.json.gz"
    try:
        return json.loads(gzip.decompress(p.read_bytes()))
    except OSError:
        return {}


@lru_cache(maxsize=8)
def cid_table(ros: str) -> dict[int, str]:
    runs = _cid_data().get(ros)
    out: dict[int, str] = {}
    if not runs:
        return out
    for cid, n, u in runs:
        if n == 0:
            out[cid] = u
        else:
            for k in range(n):
                out[cid + k] = chr(u + k)
    return out


@lru_cache(maxsize=8)
def cid_inverse(ros: str) -> dict[str, int]:
    inv: dict[str, int] = {}
    for cid, u in sorted(cid_table(ros).items()):
        inv.setdefault(u, cid)
    return inv


def ros_name(info: dict | None, doc=None) -> str:
    if not isinstance(info, dict):
        return ""
    reg = info.get("Registry", b"")
    ordg = info.get("Ordering", b"")
    if doc is not None:
        reg, ordg = doc.resolve(reg), doc.resolve(ordg)
    reg = reg.decode("latin-1") if isinstance(reg, (bytes, PdfString)) else str(reg or "")
    ordg = ordg.decode("latin-1") if isinstance(ordg, (bytes, PdfString)) else str(ordg or "")
    return f"{reg}-{ordg}"


# ====================================================================== 標準の14書体の幅（埋め込みも Widths もないとき）
_HELV = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278] + [556] * 10 + [
    278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778,
    722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222,
    222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584]


def std14_width(base: str, code: int) -> float:
    b = base.split("+")[-1].lower()
    if "courier" in b:
        return 600
    if 32 <= code <= 126:
        w = _HELV[code - 32]
        if "times" in b or "roman" in b or "serif" in b:
            return w * 0.93
        return w
    return 556


# ====================================================================== フォント
class PdfFont:
    def __init__(self, doc, fd: dict, key: str = ""):
        self.doc = doc
        self.key = key
        self.subtype = getattr(fd.get("Subtype"), "name", "")
        self.base = getattr(fd.get("BaseFont"), "name", "") or ""
        self.composite = self.subtype == "Type0"
        self.vertical = False
        self.unmapped = 0
        self.matrix = (0.001, 0, 0, 0.001, 0, 0)
        self.tounicode: CMap | None = None
        tu = doc.resolve(fd.get("ToUnicode"))
        if isinstance(tu, Stream):
            try:
                self.tounicode = CMap.parse(doc.decode(tu))
            except Exception:  # noqa: BLE001
                self.tounicode = None
        self._cache: dict[bytes, list] = {}
        if self.composite:
            self._init_composite(fd)
        else:
            self._init_simple(fd)

    # ------------------------------------------------------------------ 複合フォント（Type0）
    def _init_composite(self, fd: dict) -> None:
        doc = self.doc
        enc = doc.resolve(fd.get("Encoding"))
        self.cmap: CMap | None = None
        self.pre = None
        if isinstance(enc, Name):
            self.pre = predefined(enc.name)
            self.enc_name = enc.name
        elif isinstance(enc, Stream):
            try:
                self.cmap = CMap.parse(doc.decode(enc), resolver=lambda n: None)
            except Exception:  # noqa: BLE001
                self.cmap = None
            self.enc_name = self.cmap.name if self.cmap else ""
            if self.cmap and not self.cmap.codespace:
                pre = predefined(self.cmap.name)
                if pre:
                    self.pre = pre
                    self.cmap = None
            if self.cmap is not None:
                um = re.search(rb"/([^\s/]+)\s+usecmap", doc.decode(enc))
                if um and not self.cmap.codespace:
                    self.pre = predefined(um[1].decode("latin-1"))
        else:
            self.enc_name = ""
            self.pre = ("identity", 2, False, False)
        if self.cmap is not None:
            self.vertical = self.cmap.wmode == 1
        elif self.pre:
            self.vertical = self.pre[2]
        dfs = doc.resolve(fd.get("DescendantFonts")) or []
        df = doc.resolve(dfs[0]) if dfs else {}
        df = df if isinstance(df, dict) else {}
        self.cid_subtype = getattr(df.get("Subtype"), "name", "")
        self.ros = ros_name(doc.resolve(df.get("CIDSystemInfo")), doc)
        self.dw = float(doc.resolve(df.get("DW")) or 1000)
        self.widths = _cid_widths(doc, doc.resolve(df.get("W")))
        dw2 = doc.resolve(df.get("DW2")) or [880, -1000]
        self.dw2 = (float(dw2[0]), float(dw2[1])) if isinstance(dw2, list) and len(dw2) >= 2 else (880.0, -1000.0)
        self.w2 = _cid_widths2(doc, doc.resolve(df.get("W2"))) if self.vertical else {}
        desc = doc.resolve(df.get("FontDescriptor")) or {}
        self._metrics(desc)
        self.gid_uni: dict[int, str] | None = None
        self.cid2gid = None
        self._desc = desc
        c2g = doc.resolve(df.get("CIDToGIDMap"))
        if isinstance(c2g, Stream):
            try:
                self.cid2gid = doc.decode(c2g)
            except Exception:  # noqa: BLE001
                self.cid2gid = None
        self._ros_table = None

    def _font_program_unicode(self) -> dict[int, str]:
        if self.gid_uni is None:
            self.gid_uni = {}
            desc = self._desc
            ff = self.doc.resolve(desc.get("FontFile2")) or self.doc.resolve(desc.get("FontFile3"))
            if isinstance(ff, Stream):
                try:
                    data = self.doc.decode(ff)
                    from . import sfnt

                    if data[:4] in (b"\x00\x01\x00\x00", b"true", b"OTTO", b"ttcf"):
                        self.gid_uni = sfnt.unicode_by_glyph(data)
                    else:
                        names, _enc, is_cid = sfnt.cff_glyph_info(data)
                        if not is_cid:
                            self.gid_uni = {g: glyph_to_unicode(n) for g, n in names.items() if glyph_to_unicode(n)}
                except Exception:  # noqa: BLE001
                    self.gid_uni = {}
        return self.gid_uni

    def _colr_text(self, cid: int) -> str | None:
        """色の付いた字（絵文字など）の重ねる層：もとの字の文字にする（同じ位置に重なった層は、窓で1字にまとめる）。"""
        if not hasattr(self, "_colr"):
            self._colr = {}
            ff = self.doc.resolve(self._desc.get("FontFile2")) if isinstance(self._desc, dict) else None
            if isinstance(ff, Stream):
                from . import sfnt

                try:
                    self._colr = sfnt.colr_layers(self.doc.decode(ff))
                except Exception:  # noqa: BLE001
                    self._colr = {}
        base = self._colr.get(self._gid(cid))
        if base is None:
            return None
        if self.cid2gid is None:
            code = base.to_bytes(2, "big")
        else:
            code = next((c.to_bytes(2, "big") for c in range(len(self.cid2gid) // 2) if self._gid(c) == base), None)
        if code is not None and self.tounicode is not None:
            u = self.tounicode.to_unicode(code)
            if u:
                return u
        return self._font_program_unicode().get(base)

    def _gid(self, cid: int) -> int:
        if self.cid2gid is not None:
            k = 2 * cid
            if k + 2 <= len(self.cid2gid):
                return (self.cid2gid[k] << 8) | self.cid2gid[k + 1]
            return 0
        return cid

    def _decode_composite(self, s: bytes) -> list:
        out = []
        if self.cmap is not None:
            codes = self.cmap.split(s)
        elif self.pre is not None:
            codes = split_predefined(s, self.pre[1])
        else:
            codes = [s[i:i + 2] for i in range(0, len(s) - 1, 2)]
        for code in codes:
            text = None
            cid = None
            if self.cmap is not None:
                cid = self.cmap.to_cid(code)
            elif self.pre is not None and self.pre[0] == "identity":
                cid = int.from_bytes(code, "big")
            if self.tounicode is not None:
                text = self.tounicode.to_unicode(code)
            if text is None and self.pre is not None and self.pre[0] != "identity":
                text = decode_predefined(code, self.pre[0])
                if cid is None and text:
                    cid = self._cid_from_text(text, len(code))
            if text is None and cid is not None:
                text = self._cid_text(cid)
            if text is None and cid is not None:
                text = self._colr_text(cid)
            if cid is None:
                cid = int.from_bytes(code, "big")
            if text is None or text == "":
                if text is None:
                    self.unmapped += 1
                    text = UNKNOWN
            if self.vertical:
                w = self.w2.get(cid)
                adv = (w[0] if w else self.dw2[1]) / 1000.0
                vx = (w[1] if w else self.widths.get(cid, self.dw) / 2.0) / 1000.0
                vy = (w[2] if w else self.dw2[0]) / 1000.0
                out.append((code, text, adv, len(code) == 1 and code == b" ", (vx, vy), self.widths.get(cid, self.dw) / 1000.0))
            else:
                out.append((code, text, self.widths.get(cid, self.dw) / 1000.0, len(code) == 1 and code == b" ", None, None))
        return out

    def _cid_from_text(self, text: str, nbytes: int) -> int | None:
        if not self.ros:
            return None
        inv = cid_inverse(self.ros.replace("Adobe-", "Adobe-", 1))
        cid = inv.get(text)
        if cid is not None and self.ros == "Adobe-Japan1" and not self.pre[3] and 1 <= cid <= 95:
            cid += 230  # 比例幅でない CMap の半角の文字（AJ1 では 231〜325）
        return cid

    def _cid_text(self, cid: int) -> str | None:
        # 埋め込まれた TrueType は、まずフォント自身の cmap（グリフに付いた文字）で決める
        if self.cid_subtype == "CIDFontType2":
            u = self._font_program_unicode().get(self._gid(cid))
            if u:
                return u
        if self._ros_table is None:
            self._ros_table = cid_table(self.ros) if self.ros in ("Adobe-Japan1", "Adobe-GB1", "Adobe-CNS1", "Adobe-Korea1", "Adobe-KR") else {}
        u = self._ros_table.get(cid)
        if u:
            # TrueType で CIDToGIDMap が Identity のとき、CID はグリフ番号のことがある（Adobe-Japan1 と書かれていても）
            if self.cid_subtype == "CIDFontType2" and self.cid2gid is None and self._font_program_unicode():
                return None
            return u
        if self.cid_subtype != "CIDFontType2":
            u = self._font_program_unicode().get(cid)
            if u:
                return u
        if cid == 0:
            return ""
        return None

    # ------------------------------------------------------------------ 単純なフォント
    def _init_simple(self, fd: dict) -> None:
        doc = self.doc
        self.first = int(doc.resolve(fd.get("FirstChar")) or 0)
        ws = doc.resolve(fd.get("Widths"))
        self.widths_list = [float(doc.resolve(w) or 0) for w in ws] if isinstance(ws, list) else None
        desc = doc.resolve(fd.get("FontDescriptor")) or {}
        self._metrics(desc)
        self.missing = float(doc.resolve(desc.get("MissingWidth")) or 0) if isinstance(desc, dict) else 0
        if self.subtype == "Type3":
            fm = doc.resolve(fd.get("FontMatrix"))
            if isinstance(fm, list) and len(fm) == 6:
                self.matrix = tuple(float(doc.resolve(x)) for x in fm)
        flags = int(doc.resolve(desc.get("Flags")) or 0) if isinstance(desc, dict) else 0
        symbolic = bool(flags & 4)
        enc = doc.resolve(fd.get("Encoding"))
        base_name = enc.name if isinstance(enc, Name) else None
        diffs: dict[int, str] = {}
        if isinstance(enc, dict):
            be = doc.resolve(enc.get("BaseEncoding"))
            base_name = be.name if isinstance(be, Name) else None
            arr = doc.resolve(enc.get("Differences")) or []
            code = 0
            for x in arr:
                x = doc.resolve(x)
                if isinstance(x, (int, float)):
                    code = int(x)
                elif isinstance(x, Name):
                    diffs[code] = x.name
                    code += 1
        bname = self.base.split("+")[-1]
        table: list[str | None]
        builtin: dict[int, str] = {}
        if base_name == "WinAnsiEncoding":
            table = list(WINANSI)
        elif base_name == "MacRomanEncoding":
            table = list(MACROMAN)
        elif base_name == "StandardEncoding":
            table = [glyph_to_unicode(n) if n else None for n in STANDARD_ENCODING]
        elif base_name == "MacExpertEncoding":
            table = [None] * 256
        elif "Symbol" in bname and not diffs and base_name is None:
            table = [SYMBOL.get(i) for i in range(256)]
        elif "Dingbats" in bname and base_name is None:
            table = [ZAPF.get(i) for i in range(256)]
        else:
            # 組み込みの符号：埋め込まれたフォントから読む
            builtin = self._builtin_encoding(desc)
            if builtin:
                table = [glyph_to_unicode(builtin[i]) if i in builtin else None for i in range(256)]
            elif symbolic and self.subtype == "TrueType":
                table = [None] * 256
            else:
                table = [glyph_to_unicode(n) if n else None for n in STANDARD_ENCODING] if self.subtype == "Type1" else list(WINANSI)
        for c, n in diffs.items():
            if 0 <= c < 256:
                table[c] = glyph_to_unicode(n) or None
        self.table = table
        self.symbolic_tt = symbolic and self.subtype == "TrueType"
        self._desc = desc
        self.gid_uni = None

    def _builtin_encoding(self, desc) -> dict[int, str]:
        if not isinstance(desc, dict):
            return {}
        from . import sfnt

        try:
            ff = self.doc.resolve(desc.get("FontFile"))
            if isinstance(ff, Stream):
                return sfnt.type1_encoding(self.doc.decode(ff))
            ff3 = self.doc.resolve(desc.get("FontFile3"))
            if isinstance(ff3, Stream):
                names, enc, is_cid = sfnt.cff_glyph_info(self.doc.decode(ff3))
                if not is_cid:
                    return {c: names.get(g, "") for c, g in enc.items() if names.get(g)}
        except Exception:  # noqa: BLE001
            return {}
        return {}

    def _tt_symbol_text(self, code: int) -> str | None:
        """記号の TrueType（Wingdings など）：フォントの cmap → グリフ名・Unicode の cmap で決める。"""
        if self.gid_uni is None:
            self.gid_uni = {}
            self._sym = {}
            ff = self.doc.resolve(self._desc.get("FontFile2")) if isinstance(self._desc, dict) else None
            if isinstance(ff, Stream):
                from . import sfnt

                try:
                    data = self.doc.decode(ff)
                    self.gid_uni = sfnt.unicode_by_glyph(data)
                    self._sym = sfnt.symbol_cmap(data)
                except Exception:  # noqa: BLE001
                    pass
        g = self._sym.get(code) if self._sym else None
        if g is not None:
            u = self.gid_uni.get(g)
            if u and not 0xF000 <= ord(u[0]) <= 0xF0FF:
                return u
        return None

    def _decode_simple(self, s: bytes) -> list:
        out = []
        tu = self.tounicode
        for b in s:
            code = bytes([b])
            text = tu.to_unicode(code) if tu is not None else None
            if text is None:
                text = self.table[b]
                if text is None and self.symbolic_tt:
                    text = self._tt_symbol_text(b)
                if text is None:
                    if b == 0x20:
                        text = " "
                    elif 0x21 <= b <= 0x7E and not self.symbolic_tt:
                        text = chr(b)
                    else:
                        self.unmapped += 1
                        text = UNKNOWN
            if self.widths_list is not None and 0 <= b - self.first < len(self.widths_list):
                w = self.widths_list[b - self.first]
            elif self.widths_list is not None:
                w = self.missing
            else:
                w = std14_width(self.base, b)
            if self.subtype == "Type3":
                w = w * self.matrix[0] * 1000.0 / 1000.0
                out.append((code, text, w, b == 0x20, None, None))
            else:
                out.append((code, text, w / 1000.0, b == 0x20, None, None))
        return out

    # ------------------------------------------------------------------ 共通
    def _metrics(self, desc) -> None:
        desc = desc if isinstance(desc, dict) else {}
        asc = self.doc.resolve(desc.get("Ascent"))
        dsc = self.doc.resolve(desc.get("Descent"))
        a = float(asc) / 1000.0 if isinstance(asc, (int, float)) and asc else 0.85
        d = float(dsc) / 1000.0 if isinstance(dsc, (int, float)) and dsc else -0.15
        if a <= 0 or a > 1.5:
            a = 0.85
        if d > 0 or d < -0.8:
            d = -0.15
        self.ascent, self.descent = a, d

    def decode(self, s: bytes) -> list:
        """(文字コード, 文字, 送り幅（em）, 単語の区切りの空白か, 縦書きの位置ずれ, 縦書きの横幅)。"""
        hit = self._cache.get(s)
        if hit is not None:
            return hit
        out = self._decode_composite(s) if self.composite else self._decode_simple(s)
        out = [(c, visible_text(t), w, sp, v, vw) for c, t, w, sp, v, vw in out]
        if len(self._cache) < 20000:
            self._cache[s] = out
        return out


# 見た目は同じ字なのに、ToUnicode が別の符号を指していることがある（康熙部首「⽇」→「日」、縦書き用の約物「︓」→「：」）。
# 人が見る字に合わせる（どの字を直したかは数えて、検査に出す）。
_VERT = {0xFE10: "，", 0xFE11: "、", 0xFE12: "。", 0xFE13: "：", 0xFE14: "；", 0xFE15: "！", 0xFE16: "？", 0xFE17: "〖",
         0xFE18: "〗", 0xFE19: "…", 0xFE30: "‥", 0xFE31: "—", 0xFE32: "–", 0xFE33: "＿", 0xFE34: "＿", 0xFE35: "（",
         0xFE36: "）", 0xFE37: "｛", 0xFE38: "｝", 0xFE39: "〔", 0xFE3A: "〕", 0xFE3B: "【", 0xFE3C: "】", 0xFE3D: "《",
         0xFE3E: "》", 0xFE3F: "〈", 0xFE40: "〉", 0xFE41: "「", 0xFE42: "」", 0xFE43: "『", 0xFE44: "』", 0xFE47: "［",
         0xFE48: "］"}


def visible_text(t: str) -> str:
    if not t:
        return t
    out = []
    for ch in t:
        o = ord(ch)
        if 0x2F00 <= o <= 0x2FDF or 0x2E80 <= o <= 0x2EFF:
            n = unicodedata.normalize("NFKC", ch)
            out.append(n if len(n) == 1 else ch)
        elif o in _VERT:
            out.append(_VERT[o])
        else:
            out.append(ch)
    return "".join(out)


def _cid_widths(doc, w) -> dict[int, float]:
    out: dict[int, float] = {}
    if not isinstance(w, list):
        return out
    i = 0
    while i < len(w):
        a = doc.resolve(w[i])
        b = doc.resolve(w[i + 1]) if i + 1 < len(w) else None
        if isinstance(b, list):
            for k, x in enumerate(b):
                out[int(a) + k] = float(doc.resolve(x) or 0)
            i += 2
        else:
            c = doc.resolve(w[i + 2]) if i + 2 < len(w) else 0
            if isinstance(a, (int, float)) and isinstance(b, (int, float)) and b - a < 70000:
                for cid in range(int(a), int(b) + 1):
                    out[cid] = float(c or 0)
            i += 3
    return out


def _cid_widths2(doc, w) -> dict[int, tuple[float, float, float]]:
    out: dict[int, tuple[float, float, float]] = {}
    if not isinstance(w, list):
        return out
    i = 0
    while i < len(w):
        a = doc.resolve(w[i])
        b = doc.resolve(w[i + 1]) if i + 1 < len(w) else None
        if isinstance(b, list):
            vals = [float(doc.resolve(x) or 0) for x in b]
            for k in range(0, len(vals) - 2, 3):
                out[int(a) + k // 3] = (vals[k], vals[k + 1], vals[k + 2])
            i += 2
        else:
            vals = [float(doc.resolve(x) or 0) for x in w[i + 2:i + 5]]
            if len(vals) == 3 and isinstance(a, (int, float)) and isinstance(b, (int, float)) and b - a < 70000:
                for cid in range(int(a), int(b) + 1):
                    out[cid] = (vals[0], vals[1], vals[2])
            i += 5
    return out


def load_font(doc, ref_or_dict, key: str = "") -> PdfFont | None:
    cache = doc.__dict__.setdefault("_fonts", {})
    k = ref_or_dict.num if isinstance(ref_or_dict, Ref) else id(ref_or_dict)
    if k in cache:
        return cache[k]
    fd = doc.resolve(ref_or_dict)
    f = PdfFont(doc, fd, key) if isinstance(fd, dict) else None
    cache[k] = f
    return f
