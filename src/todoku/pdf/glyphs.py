"""PDF に書く字を、どの書体で描くか決め、描いた見た目（/AP のストリーム）を読み直して、どの字も描ける字形かを確かめる。

書体は埋め込まない（増分更新を小さく保つ）。描けるのは、仕様で字形が決まっている所だけ：
- Helvetica（標準の 14 書体）：WinAnsiEncoding の字と、標準の 14 書体の字形の表（Adobe Core14 AFM の 315 字形）にある
  ラテン文字の拡張（ş ğ İ ł ę ř ő ű ă ć Č など）。拡張の字は、/Differences で字形の名前を指した別の書体（TdkL）で描く
- CID フォント（平成角ゴシックなど）：定義済みの CMap（UniJIS-UCS2-HW-H など、Adobe の公開資料の対応表）で決まる CID が、
  宣言した字形の集まりの補遺（Supplement）の範囲にあり、幅が決まっているもの。範囲の外の CID は、ビューアーが持つ字形に
  ないことがあり、白く抜ける（証明も読み戻しも /V を読むので、気づかない）
どれでも描けない字は、白く抜けた見た目を書かずに、理由を添えて断る。
"""

from __future__ import annotations

import gzip
import json
import unicodedata
from functools import lru_cache
from pathlib import Path

from .objects import Lexer, Name, PdfString, Stream


class GlyphError(ValueError):
    pass


# ====================================================================== Helvetica（標準の 14 書体）
# WinAnsiEncoding にない字で、Core14 の AFM（Helvetica・Times・Courier の 315 字形）にある字形：Unicode → 字形の名前
_EXT_NAMES = (
    "Amacron amacron Abreve abreve Aogonek aogonek Cacute cacute Ccaron ccaron Dcaron dcaron Dcroat dcroat Emacron emacron "
    "Edotaccent edotaccent Eogonek eogonek Ecaron ecaron Gbreve gbreve Gcommaaccent gcommaaccent Imacron imacron Iogonek iogonek "
    "Idotaccent dotlessi Kcommaaccent kcommaaccent Lacute lacute Lcommaaccent lcommaaccent Lcaron lcaron Lslash lslash Nacute nacute "
    "Ncommaaccent ncommaaccent Ncaron ncaron Omacron omacron Ohungarumlaut ohungarumlaut Racute racute Rcommaaccent rcommaaccent "
    "Rcaron rcaron Sacute sacute Scedilla scedilla Tcommaaccent tcommaaccent Tcaron tcaron Umacron umacron Uring uring "
    "Uhungarumlaut uhungarumlaut Uogonek uogonek Zacute zacute Zdotaccent zdotaccent Scommaaccent scommaaccent "
    "breve caron dotaccent hungarumlaut ogonek ring minus fraction Delta lozenge greaterequal lessequal notequal partialdiff "
    "radical summation").split()
_EXT_CODES = (
    [0x100, 0x101, 0x102, 0x103, 0x104, 0x105, 0x106, 0x107, 0x10C, 0x10D, 0x10E, 0x10F, 0x110, 0x111, 0x112, 0x113,
     0x116, 0x117, 0x118, 0x119, 0x11A, 0x11B, 0x11E, 0x11F, 0x122, 0x123, 0x12A, 0x12B, 0x12E, 0x12F,
     0x130, 0x131, 0x136, 0x137, 0x139, 0x13A, 0x13B, 0x13C, 0x13D, 0x13E, 0x141, 0x142, 0x143, 0x144,
     0x145, 0x146, 0x147, 0x148, 0x14C, 0x14D, 0x150, 0x151, 0x154, 0x155, 0x156, 0x157,
     0x158, 0x159, 0x15A, 0x15B, 0x15E, 0x15F, 0x162, 0x163, 0x164, 0x165, 0x16A, 0x16B, 0x16E, 0x16F,
     0x170, 0x171, 0x172, 0x173, 0x179, 0x17A, 0x17B, 0x17C, 0x218, 0x219,
     0x2D8, 0x2C7, 0x2D9, 0x2DD, 0x2DB, 0x2DA, 0x2212, 0x2044, 0x2206, 0x25CA, 0x2265, 0x2264, 0x2260, 0x2202,
     0x221A, 0x2211])
assert len(_EXT_NAMES) == len(_EXT_CODES)
EXT = {chr(u): n for u, n in zip(_EXT_CODES, _EXT_NAMES)}  # 字 → 字形の名前
EXT_FIRST = 33  # TdkL の文字コード：33 から、上の並びの順
EXT_CODE = {ch: EXT_FIRST + i for i, ch in enumerate(EXT)}
# 幅（Helvetica の AFM。基の字と違うものだけ。ほかは基の字の幅）
_EXT_W = {"dcaron": 643, "lcaron": 299, "tcaron": 317, "Dcroat": 722, "dcroat": 556, "Lslash": 556, "lslash": 222,
          "dotlessi": 278, "Idotaccent": 278, "imacron": 278, "breve": 333, "caron": 333, "dotaccent": 333,
          "hungarumlaut": 333, "ogonek": 333, "ring": 333, "minus": 584, "fraction": 167, "Delta": 612, "lozenge": 471,
          "greaterequal": 549, "lessequal": 549, "notequal": 549, "partialdiff": 476, "radical": 453, "summation": 600}
# Helvetica の幅（AFM）：ASCII の 32〜126
HELV_ASCII = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278] + [556] * 10 + [
    278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778,
    722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222,
    222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584]
# WinAnsiEncoding の 128〜255 の幅（Helvetica の AFM。基の字で決まらないもの）
_WIN_W = {"€": 556, "‚": 222, "ƒ": 556, "„": 333, "…": 1000, "†": 556, "‡": 556, "ˆ": 333, "‰": 1000, "‹": 333, "Œ": 1000,
          "‘": 222, "’": 222, "“": 333, "”": 333, "•": 350, "–": 556, "—": 1000, "˜": 333, "™": 1000, "›": 333, "œ": 944,
          "\xa0": 278, "¡": 333, "¢": 556, "£": 556, "¤": 556, "¥": 556, "¦": 260, "§": 556, "¨": 333, "©": 737, "ª": 370,
          "«": 556, "¬": 584, "\xad": 333, "®": 737, "¯": 333, "°": 400, "±": 584, "²": 333, "³": 333, "´": 333, "µ": 556,
          "¶": 537, "·": 278, "¸": 333, "¹": 333, "º": 365, "»": 556, "¼": 834, "½": 834, "¾": 834, "¿": 611, "Æ": 1000,
          "Ð": 722, "×": 584, "Ø": 778, "Þ": 667, "ß": 611, "æ": 889, "ð": 556, "÷": 584, "ø": 611, "þ": 556}


def _base_width(ch: str) -> int | None:
    base = unicodedata.normalize("NFD", ch)[0]
    if " " <= base <= "~":
        return HELV_ASCII[ord(base) - 32]
    return None


def helv_width(ch: str) -> float:
    """Helvetica の字の幅（1000 分の 1）。"""
    if " " <= ch <= "~":
        return HELV_ASCII[ord(ch) - 32]
    if ch in _WIN_W:
        return _WIN_W[ch]
    name = EXT.get(ch)
    if name in _EXT_W:
        return _EXT_W[name]
    return _base_width(ch) or 556


def winansi(ch: str) -> int | None:
    try:
        b = ch.encode("cp1252")
    except UnicodeEncodeError:
        return None
    return b[0] if len(b) == 1 else None


# ====================================================================== CID フォント
# 字形の集まり → (資源の名前, 書体, CMap, 宣言する補遺)
CJK = {
    "Adobe-Japan1": ("TdkJ", "HeiseiKakuGo-W5", "UniJIS-UCS2-HW-H", 2),
    "Adobe-Korea1": ("TdkK", "HYGoThic-Medium", "UniKS-UCS2-H", 1),
    "Adobe-GB1": ("TdkG", "STSong-Light", "UniGB-UCS2-H", 2),
    "Adobe-CNS1": ("TdkC", "MSung-Light", "UniCNS-UCS2-H", 0),
}
# 補遺ごとの、最後の CID（Adobe Technical Note #5078・#5093・#5079・#5080）
SUPPLEMENT_LAST = {
    "Adobe-Japan1": [8283, 8358, 8719, 9353, 15443, 20316, 23057, 23059],
    "Adobe-Korea1": [9332, 18154, 18351],
    "Adobe-GB1": [7716, 9896, 22126, 22352, 29063, 30283],
    "Adobe-CNS1": [14098, 17407, 17600, 18845, 18964, 19087, 19155, 19178],
}
CMAP_ROS = {v[2]: k for k, v in CJK.items()}


@lru_cache(maxsize=1)
def _uni2cid_data() -> dict:
    p = Path(__file__).parent / "data" / "uni2cid.json.gz"
    try:
        return json.loads(gzip.decompress(p.read_bytes()))
    except OSError:
        return {}


@lru_cache(maxsize=8)
def uni2cid(cmap: str) -> dict[int, int]:
    """定義済みの CMap（UCS-2 の横書き）：Unicode → CID。"""
    name = cmap
    out: dict[int, int] = {}
    for u, n, c in _uni2cid_data().get(name, []):
        for k in range(n):
            out[u + k] = c + k
    return out


def cid_width(ros: str, cid: int, ch: str) -> float | None:
    """CID の幅（1000 分の 1）。決まらない（比例幅など）なら None。
    Adobe-Japan1 は CID の範囲で決まる（1〜230 比例幅・231〜632 半角・633〜 全角）。ほかの集まりは、字の東アジアの幅
    （Unicode の East Asian Width）が全角（W・F）か半角（H）のときだけ決める。"""
    if ros == "Adobe-Japan1":
        if 231 <= cid <= 632:
            return 500.0
        if cid >= 633:
            return 1000.0
        return None
    eaw = unicodedata.east_asian_width(ch)
    if eaw in ("W", "F"):
        return 1000.0
    if eaw == "H":
        return 500.0
    return None


def cjk_cid(ros: str, ch: str, supplement: int | None = None) -> tuple[int | None, str]:
    """その集まりの書体で、字 ch を描く CID（描けなければ None と理由）。"""
    _res, _base, cmap, sup = CJK[ros]
    sup = sup if supplement is None else supplement
    cid = uni2cid(cmap).get(ord(ch))
    if not cid:
        return None, f"{cmap} に字がありません"
    last = SUPPLEMENT_LAST[ros][min(sup, len(SUPPLEMENT_LAST[ros]) - 1)]
    if cid > last:
        return None, f"{ros} の CID {cid} は、宣言した補遺 {sup}（CID {last} まで）の外です"
    if cid_width(ros, cid, ch) is None:
        return None, f"{ros} の CID {cid} は比例幅で、幅が決まりません"
    return cid, ""


def cjk_font(ros: str) -> dict:
    res, base, cmap, sup = CJK[ros]
    reg, order = ros.split("-")
    desc = {"Type": Name("FontDescriptor"), "FontName": Name(base), "Flags": 4, "FontBBox": [-100, -250, 1100, 900],
            "ItalicAngle": 0, "Ascent": 880, "Descent": -120, "CapHeight": 700, "StemV": 80}
    cid = {"Type": Name("Font"), "Subtype": Name("CIDFontType0"), "BaseFont": Name(base),
           "CIDSystemInfo": {"Registry": PdfString(reg.encode()), "Ordering": PdfString(order.encode()), "Supplement": sup},
           "FontDescriptor": desc, "DW": 1000}
    if ros == "Adobe-Japan1":
        cid["W"] = [231, 632, 500]  # 半角（英数字・カタカナなど）は 500
    return {"Type": Name("Font"), "Subtype": Name("Type0"), "BaseFont": Name(f"{base}-{cmap}"), "Encoding": Name(cmap),
            "DescendantFonts": [cid]}


def ext_font() -> tuple[dict, bytes]:
    """TdkL：Helvetica で、WinAnsiEncoding にない拡張の字形を /Differences で指す書体。(辞書, ToUnicode の CMap)。"""
    diffs: list = [EXT_FIRST] + [Name(EXT[ch]) for ch in EXT]
    font = {"Type": Name("Font"), "Subtype": Name("Type1"), "BaseFont": Name("Helvetica"),
            "Encoding": {"Type": Name("Encoding"), "BaseEncoding": Name("WinAnsiEncoding"), "Differences": diffs}}
    lines = [f"<{EXT_CODE[ch]:02X}> <{ord(ch):04X}>" for ch in EXT]
    body = ["/CIDInit /ProcSet findresource begin", "12 dict begin", "begincmap",
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def", "/CMapName /TdkL-UCS def",
            "/CMapType 2 def", "1 begincodespacerange", "<00> <FF>", "endcodespacerange"]
    for i in range(0, len(lines), 100):
        chunk = lines[i:i + 100]
        body += [f"{len(chunk)} beginbfchar", *chunk, "endbfchar"]
    body += ["endcmap", "CMapName currentdict /CMap defineresource pop", "end", "end"]
    return font, ("\n".join(body) + "\n").encode("latin-1")


# ====================================================================== 字ごとに書体を決める
ORDER = ("Adobe-Japan1", "Adobe-Korea1", "Adobe-GB1", "Adobe-CNS1")


def choose(ch: str) -> tuple[str, bytes, float]:
    """字 → (書体の資源の名前, 描く文字コード, 幅（1000 分の 1）)。描けなければ GlyphError。"""
    if ord(ch) > 0xFFFF:
        raise GlyphError(f"「{ch}」（U+{ord(ch):X}）は、この PDF に描けません：基本多言語面の外の字で、"
                         "埋め込まない書体の CMap（UCS-2）では指せません")
    if not ("＀" <= ch <= "￯"):
        c = winansi(ch)
        if c is not None:
            return "Helv", bytes([c]), helv_width(ch)
        if ch in EXT_CODE:
            return "TdkL", bytes([EXT_CODE[ch]]), helv_width(ch)
    why = []
    for ros in ORDER:
        cid, reason = cjk_cid(ros, ch)
        if cid is not None:
            return CJK[ros][0], ch.encode("utf-16-be"), cid_width(ros, cid, ch)
        why.append(reason)
    if ch in ("\n", "\r", "\t"):
        raise GlyphError(f"制御の字（U+{ord(ch):04X}）は描けません")
    name = unicodedata.name(ch, "")
    raise GlyphError(f"「{ch}」（U+{ord(ch):04X}{' ' + name if name else ''}）は、この PDF に描けません："
                     "埋め込まない書体（Helvetica・平成角ゴシックなど）の、仕様で決まった字形にない字です"
                     "（書くと、その字が白く抜けます）")


def runs(text: str) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    for ch in text:
        key = choose(ch)[0]
        if out and out[-1][0] == key:
            out[-1] = (key, out[-1][1] + ch)
        else:
            out.append((key, ch))
    return out


def encode_run(key: str, run: str) -> bytes:
    return b"".join(choose(ch)[1] for ch in run)


def width(text: str) -> float:
    """字の並びの幅（1000 分の 1 の和を 1000 で割ったもの：字の大きさ 1 のときの幅）。"""
    return sum(choose(ch)[2] for ch in text) / 1000.0


def undrawable(text: str) -> list[str]:
    """描けない字の理由（重ねない）。"""
    out = []
    seen = set()
    for ch in text:
        if ch in seen or ch in "\n":
            continue
        seen.add(ch)
        try:
            choose(ch)
        except GlyphError as e:
            out.append(str(e))
    return out


# ====================================================================== 見た目を読み直す
_STD_FAMILIES = ("Helvetica", "Arial", "Times", "Courier")
WIN_TABLE = [None] * 256
for _c in range(256):
    try:
        WIN_TABLE[_c] = bytes([_c]).decode("cp1252")
    except UnicodeDecodeError:
        pass
_EXT_BY_NAME = {n: ch for ch, n in EXT.items()}


def _std_glyph_unicode(name: str) -> str | None:
    """標準の 14 書体の字形の名前 → 字（Core14 の 315 字形のどれか。なければ None）。"""
    if name in _EXT_BY_NAME:
        return _EXT_BY_NAME[name]
    from .fonts import glyph_to_unicode

    u = glyph_to_unicode(name)
    if u and len(u) == 1 and winansi(u) is not None:
        return u
    return None


class _FontInfo:
    def __init__(self, doc, fd):
        self.kind = "unknown"
        self.problem = ""
        self.name = ""
        self.table: dict[int, str | None] = {}
        self.ros = ""
        self.sup = 0
        self.cmap = ""
        if not isinstance(fd, dict):
            self.problem = "書体の辞書がありません"
            return
        st = getattr(fd.get("Subtype"), "name", "")
        base = getattr(doc.resolve(fd.get("BaseFont")), "name", "")
        self.name = base
        if st == "Type1" and not self._embedded(doc, fd):
            if base in ("ZapfDingbats", "Symbol"):
                self.kind = "symbolic"
                return
            if not base.split("-")[0].startswith(_STD_FAMILIES):
                self.problem = f"埋め込まない書体 {base} は、標準の 14 書体ではありません"
                return
            self.kind = "std"
            enc = doc.resolve(fd.get("Encoding"))
            names: dict[int, str | None] = {}
            base_enc = enc if isinstance(enc, Name) else (doc.resolve(enc.get("BaseEncoding")) if isinstance(enc, dict) else None)
            ben = getattr(base_enc, "name", "")
            if ben not in ("WinAnsiEncoding", ""):
                self.problem = f"書体の符号化 {ben} を、ここでは確かめられません"
                return
            if ben == "" and not isinstance(enc, dict):
                self.problem = "符号化の指定がない標準の書体（StandardEncoding）で、確かめられません"
                return
            for c in range(256):
                names[c] = WIN_TABLE[c]
            if isinstance(enc, dict):
                diffs = doc.resolve(enc.get("Differences")) or []
                code = 0
                for x in diffs:
                    if isinstance(x, int):
                        code = x
                    elif isinstance(x, Name):
                        names[code] = _std_glyph_unicode(x.name)
                        code += 1
            self.table = names
            return
        if st == "Type0":
            enc = doc.resolve(fd.get("Encoding"))
            df = doc.resolve(fd.get("DescendantFonts"))
            cf = doc.resolve(df[0]) if isinstance(df, list) and df else None
            if not isinstance(enc, Name) or enc.name not in CMAP_ROS or not isinstance(cf, dict) or self._embedded(doc, cf):
                self.problem = f"書体 {base} の CMap・埋め込みを、ここでは確かめられません"
                return
            info = doc.resolve(cf.get("CIDSystemInfo")) or {}
            from .fonts import ros_name

            self.ros = ros_name(info, doc)
            if self.ros != CMAP_ROS[enc.name]:
                self.problem = f"CMap {enc.name} と字形の集まり {self.ros} が合いません"
                return
            self.sup = int(doc.resolve(info.get("Supplement")) or 0)
            self.cmap = enc.name
            self.kind = "cid"
            return
        self.problem = f"書体 {base}（{st}）を、ここでは確かめられません"

    @staticmethod
    def _embedded(doc, fd) -> bool:
        desc = doc.resolve(fd.get("FontDescriptor"))
        return isinstance(desc, dict) and any(k in desc for k in ("FontFile", "FontFile2", "FontFile3"))

    def decode(self, s: bytes) -> tuple[str, list[str]]:
        """文字コード → (字, 描けない理由)。"""
        out, probs = [], []
        if self.kind == "symbolic":
            return "", []
        if self.kind == "std":
            for c in s:
                u = self.table.get(c)
                if u is None:
                    probs.append(f"文字コード {c} は、{self.name} の字形を指していません")
                    out.append("�")
                else:
                    out.append(u)
            return "".join(out), probs
        if self.kind == "cid":
            m = uni2cid(self.cmap)
            last = SUPPLEMENT_LAST[self.ros][min(self.sup, len(SUPPLEMENT_LAST[self.ros]) - 1)]
            for i in range(0, len(s) - 1, 2):
                u = int.from_bytes(s[i:i + 2], "big")
                ch = chr(u)
                cid = m.get(u)
                if not cid:
                    probs.append(f"「{ch}」は {self.cmap} にありません（{self.name} で描けません）")
                elif cid > last:
                    probs.append(f"「{ch}」（{self.ros} の CID {cid}）は、書体の補遺 {self.sup}（CID {last} まで）の外で、描けません")
                elif cid_width(self.ros, cid, ch) is None:
                    probs.append(f"「{ch}」（{self.ros} の CID {cid}）は幅の決まらない字形です")
                out.append(ch)
            return "".join(out), probs
        return "", [self.problem]


def appearance_text(doc, stream, resources=None) -> tuple[str, list[str]]:
    """見た目のストリーム（Form XObject）が描く字と、描けない字の理由。行は改行でつなぐ。
    ZapfDingbats など記号の書体の字は、字に数えない（チェックの印）。"""
    if not isinstance(stream, Stream):
        return "", []
    res = resources if resources is not None else doc.resolve(stream.dict.get("Resources"))
    fonts = doc.resolve(res.get("Font")) if isinstance(res, dict) else None
    fonts = fonts if isinstance(fonts, dict) else {}
    data = doc.decode(stream) if stream.num else stream.raw
    lx = Lexer(data, 0)
    stack: list = []
    cur = None
    cache: dict[str, _FontInfo] = {}
    text: list[str] = []
    probs: list[str] = []
    line_started = False
    while True:
        kind, val = lx.token()
        if kind == "eof":
            break
        if kind == "kw":
            op = val
            if op == "Tf" and len(stack) >= 2 and isinstance(stack[-2], Name):
                key = stack[-2].name
                if key not in cache:
                    fd = doc.resolve(fonts.get(key))
                    cache[key] = _FontInfo(doc, fd)
                cur = cache[key]
            elif op in ("Tj", "'", '"', "TJ"):
                strs = [x for x in (stack[-1] if op == "TJ" and stack and isinstance(stack[-1], list) else stack[-1:])
                        if isinstance(x, (bytes, PdfString))]
                for s in strs:
                    if cur is None:
                        probs.append("書体を決める前に字を描いています")
                        continue
                    t, p = cur.decode(bytes(s))
                    text.append(t)
                    probs += p
                line_started = True
            elif op in ("Tm", "Td", "TD", "T*") and line_started:
                text.append("\n")
                line_started = False
            elif op == "BI":
                break
            stack = []
            continue
        if kind == "[":
            arr = []
            while True:
                k2, v2 = lx.token()
                if k2 in ("]", "eof"):
                    break
                if k2 in ("str", "num", "name"):
                    arr.append(v2)
            stack.append(arr)
            continue
        if kind == "<<":
            depth = 1
            while depth:
                k2, _v = lx.token()
                if k2 == "eof":
                    break
                depth += {"<<": 1, ">>": -1}.get(k2, 0)
            continue
        stack.append(val)
    out = "".join(text).strip("\n")
    return out, list(dict.fromkeys(probs))


def widget_appearance(doc, widget: dict, state: str | None = None):
    """欄の部品（Widget）の通常の見た目（/AP /N）。状態を持つもの（チェックなど）は、その状態のもの。"""
    ap = doc.resolve(widget.get("AP")) if isinstance(widget, dict) else None
    n = doc.resolve(ap.get("N")) if isinstance(ap, dict) else None
    if isinstance(n, dict) and not isinstance(n, Stream):
        st = state or getattr(widget.get("AS"), "name", None)
        n = doc.resolve(n.get(st)) if st else None
    return n if isinstance(n, Stream) else None
