"""PDF に書く字を、どの書体で描くか決め、描いた見た目（/AP のストリーム）を読み直して、どの字も描ける字形かを確かめる。

字形のある字は、語や言語によらず、同梱の書体（Noto Sans JP。embedfont）から使った字形だけを埋め込んで描く。同梱の書体にない字は
次の順。埋め込まない書体で描けるのは、仕様で字形が決まっている所だけ：
- Helvetica（標準の 14 書体）：WinAnsiEncoding の字と、標準の 14 書体の字形の表（Adobe Core14 AFM の 315 字形）にある
  ラテン文字の拡張（ş ğ İ ł ę ř ő ű ă ć Č など）。拡張の字は、/Differences で字形の名前を指した別の書体（TdkL）で描く
- CID フォント（平成角ゴシックなど）：定義済みの CMap（UniJIS-UCS2-HW-H など、Adobe の公開資料の対応表）で決まる CID が、
  宣言した字形の集まりの補遺（Supplement）の範囲にあり、幅が決まっているもの。範囲の外の CID は、ビューアーが持つ字形に
  ないことがあり、白く抜ける（証明も読み戻しも /V を読むので、気づかない）
どれでも描けない字は、白く抜けた見た目を書かずに、理由を添えて断る。
字の組み立て（シェーピング）が要る行（アラビア文字・ヘブライ文字・インド系の文字・タイ文字）は shape.py で組み立て、
このパソコンの書体から使った字形だけを埋め込んで描く（行を元の字の ActualText で包む）。読み直すときは、その範囲の
字形の並びが、元の字をいま組み立てた並びと同じかを確かめる。

Helvetica の字幅は Adobe Core 14 AFM による。Copyright (c) 1985, 1987, 1989, 1990, 1997 Adobe Systems Incorporated. All Rights Reserved.
"""

from __future__ import annotations

import gzip
import json
import struct
import contextlib
import contextvars
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
          "Ð": 722, "×": 584, "Ø": 778, "Þ": 667, "ß": 611, "æ": 889, "ð": 556, "÷": 584, "ø": 611, "þ": 556,
          "ì": 278, "í": 278, "î": 278, "ï": 278}
# 字形の輪郭の四角（Helvetica の AFM の B：左・下・右・上。1000 分の 1）。「字のコード（16 進）:四角」の並び
_HELV_BOX_SRC = (
    "20:0,0,0,0 21:90,0,187,718 22:70,463,285,718 23:28,0,529,688 24:32,-115,520,775 25:39,-19,850,703 26:44,-15,645,718 "
    "27:59,463,132,718 28:68,-207,299,733 29:34,-207,265,733 2A:39,431,349,718 2B:39,0,545,505 2C:87,-147,191,106 "
    "2D:44,232,289,322 2E:87,0,191,106 2F:-17,-19,295,737 30:37,-19,519,703 31:101,0,359,703 32:26,0,507,703 "
    "33:34,-19,522,703 34:25,0,523,703 35:32,-19,514,688 36:38,-19,518,703 37:37,0,523,688 38:38,-19,517,703 "
    "39:42,-19,514,703 3A:87,0,191,516 3B:87,-147,191,516 3C:48,11,536,495 3D:39,115,545,390 3E:48,11,536,495 "
    "3F:56,0,492,727 40:147,-19,868,737 41:14,0,654,718 42:74,0,627,718 43:44,-19,681,737 44:81,0,674,718 45:86,0,616,718 "
    "46:86,0,583,718 47:48,-19,704,737 48:77,0,646,718 49:91,0,188,718 4A:17,-19,428,718 4B:76,0,663,718 4C:76,0,537,718 "
    "4D:73,0,761,718 4E:76,0,646,718 4F:39,-19,739,737 50:86,0,622,718 51:39,-56,739,737 52:88,0,684,718 53:49,-19,620,737 "
    "54:14,0,597,718 55:79,-19,644,718 56:20,0,647,718 57:16,0,928,718 58:19,0,648,718 59:14,0,653,718 5A:23,0,588,718 "
    "5B:63,-196,250,722 5C:-17,-19,295,737 5D:28,-196,215,722 5E:-14,264,483,688 5F:0,-125,556,-75 60:14,593,211,734 "
    "61:36,-15,530,538 62:58,-15,517,718 63:30,-15,477,538 64:35,-15,499,718 65:40,-15,516,538 66:14,0,262,728 "
    "67:40,-220,499,538 68:65,0,491,718 69:67,0,155,718 6A:-16,-210,155,718 6B:67,0,501,718 6C:67,0,155,718 6D:65,0,769,538 "
    "6E:65,0,491,538 6F:35,-14,521,538 70:58,-207,517,538 71:35,-207,494,538 72:77,0,332,538 73:32,-15,464,538 "
    "74:14,-7,257,669 75:68,-15,489,523 76:8,0,492,523 77:14,0,709,523 78:11,0,490,523 79:11,-214,489,523 7A:31,0,469,523 "
    "7B:42,-196,292,722 7C:94,-225,167,775 7D:42,-196,292,722 7E:61,180,523,326 A1:118,-195,215,523 A2:51,-115,513,623 "
    "A3:33,-16,539,718 A4:28,99,528,603 A5:3,0,553,688 A6:94,-150,167,700 A7:43,-191,512,737 A8:40,604,293,706 "
    "A9:-14,-19,752,737 AA:24,405,346,737 AB:97,108,459,446 AC:39,108,545,390 AE:-14,-19,752,737 AF:10,627,323,684 "
    "B0:54,411,346,703 B1:39,0,545,506 B2:4,281,323,703 B3:5,270,325,703 B4:122,593,319,734 B5:68,-207,489,523 "
    "B6:18,-173,497,718 B7:77,190,202,315 B8:45,-225,259,0 B9:43,281,222,703 BA:25,405,341,737 BB:97,108,459,446 "
    "BC:73,-19,756,703 BD:43,-19,773,703 BE:45,-19,810,703 BF:91,-201,527,525 C0:14,0,654,929 C1:14,0,654,929 "
    "C2:14,0,654,929 C3:14,0,654,917 C4:14,0,654,901 C5:14,0,654,931 C6:8,0,951,718 C7:44,-225,681,737 C8:86,0,616,929 "
    "C9:86,0,616,929 CA:86,0,616,929 CB:86,0,616,901 CC:-13,0,188,929 CD:91,0,292,929 CE:-6,0,285,929 CF:13,0,266,901 "
    "D0:0,0,674,718 D1:76,0,646,917 D2:39,-19,739,929 D3:39,-19,739,929 D4:39,-19,739,929 D5:39,-19,739,917 "
    "D6:39,-19,739,901 D7:39,0,545,506 D8:39,-19,740,737 D9:79,-19,644,929 DA:79,-19,644,929 DB:79,-19,644,929 "
    "DC:79,-19,644,901 DD:14,0,653,929 DE:86,0,622,718 DF:67,-15,571,728 E0:36,-15,530,734 E1:36,-15,530,734 "
    "E2:36,-15,530,734 E3:36,-15,530,722 E4:36,-15,530,706 E5:36,-15,530,756 E6:36,-15,847,538 E7:30,-225,477,538 "
    "E8:40,-15,516,734 E9:40,-15,516,734 EA:40,-15,516,734 EB:40,-15,516,706 EC:-13,0,184,734 ED:95,0,292,734 "
    "EE:-6,0,285,734 EF:13,0,266,706 F0:35,-15,522,737 F1:65,0,491,722 F2:35,-14,521,734 F3:35,-14,521,734 "
    "F4:35,-14,521,734 F5:35,-14,521,722 F6:35,-14,521,706 F7:39,-19,545,524 F8:28,-22,537,545 F9:68,-15,489,734 "
    "FA:68,-15,489,734 FB:68,-15,489,734 FC:68,-15,489,706 FD:11,-214,489,734 FE:58,-207,517,718 FF:11,-214,489,706 "
    "100:14,0,654,879 101:36,-15,530,684 102:14,0,654,926 103:36,-15,530,731 104:14,-225,654,718 105:36,-220,547,538 "
    "106:44,-19,681,929 107:30,-15,477,734 10C:44,-19,681,929 10D:30,-15,477,734 10E:81,0,674,929 10F:35,-15,655,718 "
    "110:0,0,674,718 111:35,-15,550,718 112:86,0,616,879 113:40,-15,516,684 116:86,0,616,901 117:40,-15,516,706 "
    "118:86,-220,633,718 119:40,-225,516,538 11A:86,0,616,929 11B:40,-15,516,734 11E:48,-19,704,926 11F:40,-220,499,731 "
    "122:48,-225,704,737 123:40,-220,499,822 12A:-17,0,296,879 12B:5,0,272,684 12E:-3,-225,211,718 12F:-31,-225,183,718 "
    "130:91,0,188,901 131:95,0,183,523 136:76,-225,663,718 137:67,-225,501,718 139:76,0,537,929 13A:67,0,264,929 "
    "13B:76,-225,537,718 13C:67,-225,167,718 13D:76,0,537,718 13E:67,0,311,718 141:-20,0,537,718 142:-20,0,242,718 "
    "143:76,0,646,929 144:65,0,491,734 145:76,-225,646,718 146:65,-225,491,538 147:76,0,646,929 148:65,0,491,734 "
    "14C:39,-19,739,879 14D:35,-14,521,684 150:39,-19,739,929 151:35,-14,521,734 152:36,-19,965,737 153:35,-15,902,538 "
    "154:88,0,684,929 155:77,0,332,734 156:88,-225,684,718 157:77,-225,332,538 158:88,0,684,929 159:61,0,352,734 "
    "15A:49,-19,620,929 15B:32,-15,464,734 15E:49,-225,620,737 15F:32,-225,464,538 160:49,-19,620,929 161:32,-15,464,734 "
    "162:14,-225,597,718 163:14,-225,257,669 164:14,0,597,929 165:14,-7,329,808 16A:79,-19,644,879 16B:68,-15,489,684 "
    "16E:79,-19,644,931 16F:68,-15,489,756 170:79,-19,644,929 171:68,-15,521,734 172:79,-225,644,718 173:68,-225,519,523 "
    "178:14,0,653,901 179:23,0,588,929 17A:31,0,469,734 17B:23,0,588,901 17C:31,0,469,706 17D:23,0,588,929 17E:31,0,469,734 "
    "192:-11,-207,501,737 218:49,-225,620,737 219:32,-225,464,538 2C6:21,593,312,734 2C7:21,593,312,734 2D8:13,595,321,731 "
    "2D9:121,604,212,706 2DA:75,572,259,756 2DB:73,-225,287,0 2DC:-4,606,337,722 2DD:31,593,409,734 2013:0,240,556,313 "
    "2014:0,240,1000,313 2018:65,470,169,725 2019:53,463,157,718 201A:53,-149,157,106 201C:38,470,307,725 "
    "201D:26,463,295,718 201E:26,-149,295,106 2020:43,-159,514,718 2021:43,-159,514,718 2022:18,202,333,517 "
    "2026:115,0,885,106 2030:7,-19,994,703 2039:88,108,245,446 203A:88,108,245,446 2044:-166,-19,333,703 20AC:0,0,0,0 "
    "2122:46,306,903,718 2202:13,-38,463,714 2206:6,0,608,688 2211:15,-10,586,706 2212:39,216,545,289 221A:-4,-80,458,762 "
    "2260:12,-35,537,551 2264:26,0,523,674 2265:26,0,523,674 25CA:10,0,462,728 "
)
HELV_BOX = {chr(int(k, 16)): tuple(int(v) / 1000.0 for v in b.split(","))
            for k, b in (it.split(":") for it in _HELV_BOX_SRC.split())}


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


def helv_box(ch: str) -> tuple | None:
    """Helvetica の字形の輪郭の四角（字の大きさ 1 のとき：左・下・右・上）。墨のない字（空白）は ()。表にない字は None。
    埋め込まない書体なので、見る道具は自分の Helvetica（か、同じ字幅の代わりの書体）で描く：その字形に近い値。"""
    b = HELV_BOX.get(ch)
    if b is None:
        return None
    return b if b[2] > b[0] and b[3] > b[1] else ()


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
# 漢字の字形の集まり（ROS）を、書類の言語（Catalog の /Lang、BCP 47）で先に選ぶ：同じ漢字でも、和文・中文（簡体・繁体）・
# 韓国の字形は違う。言語の分からない書類は、これまでどおり ORDER の順（Adobe-Japan1 から）
_LANG_ROS = {"ja": "Adobe-Japan1", "ko": "Adobe-Korea1", "zh-hant": "Adobe-CNS1", "zh-tw": "Adobe-CNS1", "zh-hk": "Adobe-CNS1",
             "zh-mo": "Adobe-CNS1", "zh-hans": "Adobe-GB1", "zh-cn": "Adobe-GB1", "zh-sg": "Adobe-GB1", "zh": "Adobe-GB1"}
_order = contextvars.ContextVar("todoku_cjk_order", default=ORDER)
# 埋め込まない書体を避ける文書か（PDF/A：書体をすべて埋め込む決まり）。このときは、このパソコンの書体を先に試す
_embed_first = contextvars.ContextVar("todoku_embed_first", default=False)


def order_for_lang(lang: str) -> tuple[str, ...]:
    """書類の言語（/Lang の BCP 47。「zh-Hant-TW」「ko-KR」など）→ 字形の集まりを選ぶ順。分からなければ ORDER。"""
    tags = [t.lower() for t in str(lang or "").replace("_", "-").split("-") if t]
    if not tags:
        return ORDER
    first = None
    if tags[0] == "zh":
        for t in tags[1:]:
            first = _LANG_ROS.get("zh-" + t)
            if first:
                break
    first = first or _LANG_ROS.get(tags[0])
    if first is None:
        return ORDER
    return (first,) + tuple(r for r in ORDER if r != first)


@contextlib.contextmanager
def cjk_order(lang: str):
    """この中で描く字は、書類の言語の字形の集まりから先に選ぶ（fill が書類の /Lang で包む）。"""
    token = _order.set(order_for_lang(lang))
    try:
        yield
    finally:
        _order.reset(token)


def first_order() -> str:
    """いまの書類の言語で先に選ぶ字形の集まり。"""
    return _order.get()[0]


@contextlib.contextmanager
def embed_first(on: bool):
    """この中で描く字は、埋め込まない書体より、このパソコンの書体を先に試す（fill が PDF/A の文書で包む）。"""
    token = _embed_first.set(on)
    try:
        yield
    finally:
        _embed_first.reset(token)


def choose(ch: str) -> tuple[str, bytes, float]:
    """字 → (書体の資源の名前, 描く文字コード, 幅（1000 分の 1）)。描けなければ GlyphError。
    同梱の書体に字形があれば、それで描く。なければ、これまでの順（Helvetica・定義済みの CMap・このパソコンの書体）。"""
    from .embedfont import BUNDLED_KEY, bundled_glyph

    b = bundled_glyph(ch)
    if b is not None:
        return BUNDLED_KEY, b[0], b[1]
    if _embed_first.get():
        e = _embedded(ch)
        if e is not None:
            return e
    if ord(ch) > 0xFFFF:
        raise GlyphError(f"「{ch}」（U+{ord(ch):X}）は、この PDF に描けません：基本多言語面の外の字で、"
                         "埋め込まない書体の CMap（UCS-2）では指せません")
    if not ("＀" <= ch <= "￯"):
        c = winansi(ch)
        if c is not None:
            return "Helv", bytes([c]), helv_width(ch)
        if ch in EXT_CODE:
            return "TdkL", bytes([EXT_CODE[ch]]), helv_width(ch)
    if ch in ("\n", "\r", "\t"):
        raise GlyphError(f"制御の字（U+{ord(ch):04X}）は描けません")
    wide = unicodedata.east_asian_width(ch) in ("W", "F", "H")
    emb = _embedded(ch)
    if emb is not None and not wide and unicodedata.category(ch).startswith("L"):
        return emb  # ギリシャ文字・キリル文字など、東アジアでない文字：埋め込んだ書体の比例幅の字形で（全角の字形より自然）
    for ros in _order.get():
        cid, _reason = cjk_cid(ros, ch)
        if cid is not None:
            return CJK[ros][0], ch.encode("utf-16-be"), cid_width(ros, cid, ch)
    if emb is not None:
        return emb
    name = unicodedata.name(ch, "")
    from .embedfont import system_font, usable

    tail = ""
    if system_font() is None:
        tail = "。このパソコンに、字形を埋め込める書体（Arial・DejaVu Sans など）も見つかりません"
    elif not usable(ch):
        tail = "。結合の字・右から左に書く字は、字の並べ替え・組み合わせ（シェーピング）が要るので、埋め込みでも描けません"
    else:
        tail = f"。このパソコンの書体（{system_font().name}）にも字形がありません"
    raise GlyphError(f"「{ch}」（U+{ord(ch):04X}{' ' + name if name else ''}）は、この PDF に描けません："
                     "埋め込まない書体（Helvetica・平成角ゴシックなど）の、仕様で決まった字形にない字です"
                     f"（書くと、その字が白く抜けます）{tail}")


def _embedded(ch: str):
    """埋め込む書体（TdkE）で描けるなら (TdkE, GID の 2 バイト, 幅)。"""
    from .embedfont import system_font, usable

    font = system_font()
    if font is None or not usable(ch):
        return None
    g = font.gid(ch)
    if g is None:
        return None
    return "TdkE", g.to_bytes(2, "big"), font.width(g)


def shaped(text: str) -> bool:
    """字の組み立て（シェーピング）をして描く行か（アラビア文字・インド系の文字など。shape.py）。"""
    from .shape import needs

    return needs(text)


def units(text: str) -> list[str]:
    """折り返してよい単位（組み立てる字は、音節・記号の付いた字の並びを切り離さない）。"""
    from .shape import units as _units

    return _units(text)


def line_ops(text: str, size: float) -> bytes:
    """組み立てて描く 1 行の命令（文字の命令の中、Tm の後に置く）。行全体を ActualText（論理の順の元の字）で包み、
    組み立てた字形は、位置の差を TJ の数と Ts（上下）で指す。"""
    from .shape import line

    ln = line(text)
    out = bytearray(b"/Span <</ActualText <FEFF" + text.encode("utf-16-be").hex().upper().encode() + b">>> BDC\n")
    for sg in ln.segs:
        out += f"/{sg.key} {size:.2f} Tf ".encode()
        if not sg.shaped:
            out += b"<" + encode_run(sg.key, sg.text).hex().upper().encode() + b"> Tj\n"
            continue
        parts: list[bytes] = []
        cur = 0.0
        rise = 0.0

        def flush():
            if parts:
                out.extend(b"[" + b" ".join(parts) + b"] TJ\n")
                parts.clear()

        for gid, x, y, w in sg.glyphs:
            if abs(y - rise) > 0.005:
                flush()
                out += f"{y * size / 1000:.3f} Ts\n".encode()
                rise = y
            adj = cur - x
            if abs(adj) >= 0.005:
                parts.append(_num(adj))
            parts.append(b"<%04X>" % gid)
            cur = x + w
        adj = cur - sg.width
        if abs(adj) >= 0.005:
            parts.append(_num(adj))
        flush()
        if rise:
            out += b"0 Ts\n"
    out += b"EMC\n"
    return bytes(out)


def _num(v: float) -> bytes:
    t = f"{v:.2f}".rstrip("0").rstrip(".")
    return (t if t not in ("-0", "") else "0").encode()


def runs(text: str) -> list[tuple[str, str]]:
    if shaped(text):
        from .shape import line

        return [(sg.key, sg.logical if sg.shaped else sg.text) for sg in line(text).segs]
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
    if shaped(text):
        from .shape import line

        return line(text).width / 1000.0
    return sum(choose(ch)[2] for ch in text) / 1000.0


@lru_cache(maxsize=65536)
def _glyph_y(font, gid: int) -> tuple[float, float] | None:
    """TrueType の字形の輪郭の、基線からの上の端・下の端（字の大きさ 1 のとき）。輪郭のない字形（空白など）は None。"""
    g = font.glyph(gid)
    if len(g) < 10:
        return None
    _n, _x0, y0, _x1, y1 = struct.unpack_from(">hhhhh", g, 0)
    return y1 / font.upem, y0 / font.upem


def _char_y(ch: str) -> tuple[float, float] | None:
    from .embedfont import BUNDLED_KEY, _read_bundled, system_font

    try:
        key = choose(ch)[0]
    except GlyphError:
        return None
    if key in ("Helv", "TdkL"):
        b = helv_box(ch)
        return (b[3], b[1]) if b else None
    font = _read_bundled() if key == BUNDLED_KEY else system_font() if key == "TdkE" else None
    g = font.gid(ch) if font is not None else None
    return _glyph_y(font, g) if g is not None else None


def ink_extent(text: str) -> tuple[float, float] | None:
    """1 行の字の墨の、基線からの上の端・下の端（字の大きさ 1 のとき。基線より下は負）：埋め込む書体の、字形の輪郭の四角から。
    埋め込まない Helvetica の字は、Adobe の AFM の四角から。ほかの埋め込まない書体で描く字（定義済みの CMap の書体：字形は、
    見る道具の書体で決まる）は数えない。輪郭の分かる字がなければ None。"""
    ys: list[tuple[float, float]] = []
    if shaped(text):
        from .shape import line

        for sg in line(text).segs:
            if not sg.shaped:
                ys += [y for y in map(_char_y, sg.text) if y]
                continue
            for gid, _x, dy, _w in sg.glyphs:
                y = _glyph_y(sg.font.tt, gid) if sg.font is not None else None
                if y:
                    ys.append((y[0] + dy / 1000.0, y[1] + dy / 1000.0))
    else:
        ys = [y for y in map(_char_y, text) if y]
    return (max(y[0] for y in ys), min(y[1] for y in ys)) if ys else None


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
        self.c2g = None
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
            if getattr(enc, "name", "") == "Identity-H" and isinstance(cf, dict) and self._embedded(doc, cf):
                self._init_embedded(doc, fd, cf)
                return
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

    def _init_embedded(self, doc, fd, cf) -> None:
        """埋め込んだ TrueType（CIDFontType2・Identity-H・CIDToGIDMap /Identity）：字形の番号が書体にあり、中身があるか。"""
        from .embedfont import TrueType
        from .fonts import CMap

        desc = doc.resolve(cf.get("FontDescriptor"))
        ff = doc.resolve(desc.get("FontFile2")) if isinstance(desc, dict) else None
        c2g = doc.resolve(cf.get("CIDToGIDMap"))
        if getattr(doc.resolve(cf.get("Subtype")), "name", "") != "CIDFontType2" or ff is None \
                or not (getattr(c2g, "name", "") == "Identity" or isinstance(c2g, Stream)):
            self.problem = f"書体 {self.name} の埋め込みの形を、ここでは確かめられません"
            return
        try:
            self.c2g = doc.decode(c2g) if isinstance(c2g, Stream) else None
            cache = doc.__dict__.setdefault("_tdk_embedded_fonts", {})
            self.tt = cache.get(ff.num) or cache.setdefault(ff.num, TrueType("", doc.decode(ff)))
        except Exception as e:  # noqa: BLE001
            self.problem = f"埋め込んだ書体 {self.name} を読めません（{e}）"
            return
        tu = doc.resolve(fd.get("ToUnicode"))
        self.tu = CMap.parse(doc.decode(tu)) if tu is not None and hasattr(tu, "raw") else None
        self.kind = "embedded"

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
        if self.kind == "embedded":
            for i in range(0, len(s) - 1, 2):
                g = int.from_bytes(s[i:i + 2], "big")
                if self.c2g is not None:  # 字ごとの CID（同梱の書体）→ 字形の番号
                    g = int.from_bytes(self.c2g[2 * g:2 * g + 2], "big")
                ch = (self.tu.to_unicode(s[i:i + 2]) if self.tu is not None else None) or "�"
                if g == 0 or g >= self.tt.num:
                    probs.append(f"「{ch}」の字形の番号 {g} が、埋め込んだ書体 {self.name} にありません")
                elif not self.tt.glyph(g) and not ch.isspace():
                    probs.append(f"「{ch}」の字形（番号 {g}）の中身が、埋め込んだ書体 {self.name} にありません")
                if ch == "�":
                    probs.append(f"字形の番号 {g} の字が、{_NO_TU}")
                out.append(ch)
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


_NO_TU = "ToUnicode で決まりません"
_EMPTY = "の中身が、埋め込んだ書体"


def _check_span(actual: str, drawn: list) -> tuple[list[str], bool]:
    """ActualText の範囲で、組み立てた書体で描いた字形：元の字をいま組み立てた字形と同じか・点線の丸を入れていないか。
    (理由, 組み立てた字形と同じだと確かめたか)。"""
    from .shape import font_by_name, verify

    probs = []
    seq = []
    fonts: dict[str, object] = {}
    for info, gid in drawn:
        nm = info.name.split("+", 1)[-1]
        if nm not in fonts:
            fonts[nm] = font_by_name(nm)
        sf = fonts[nm]
        if sf is None:
            continue
        seq.append((sf.name, gid))
        dot = sf.gid(0x25CC)
        if dot is not None and gid == dot and "\u25cc" not in actual:
            probs.append(f"「{actual}」の見た目に、元の字にない点線の丸（U+25CC）があります（組み立てられなかった並び）")
    ok = False
    if seq:
        why = verify(actual, seq)
        if why:
            probs.append(why)
        else:
            ok = not probs
    return probs, ok


def appearance_text(doc, stream, resources=None) -> tuple[str, list[str]]:
    """見た目のストリーム（Form XObject）が描く字と、描けない字の理由。行は改行でつなぐ。
    ZapfDingbats など記号の書体の字は、字に数えない（チェックの印）。
    ActualText の範囲（字を組み立てて描いた行）は、その字にし、描いた字形がその字を組み立てたものかを確かめる。"""
    from .objects import PdfError, _obj, decode_text

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
    drawn: list = []  # 埋め込んだ書体で描いた (書体, 字形)
    marks: list = []  # 印の付いた範囲：(text の位置, ActualText, drawn の位置, probs の位置)
    line_started = False
    while True:
        kind, val = lx.token()
        if kind == "eof":
            break
        if kind == "kw":
            op = val
            if op in ("BDC", "BMC"):
                props = stack[-1] if op == "BDC" and stack else None
                at = props.get("ActualText") if isinstance(props, dict) else None
                actual = decode_text(at) if isinstance(at, (bytes, PdfString)) else None
                marks.append((len(text), actual, len(drawn), len(probs)))
            elif op == "EMC" and marks:
                t0, actual, d0, p0 = marks.pop()
                if actual is not None:
                    sp, ok = _check_span(actual, drawn[d0:])
                    # 字を組み立てた字形だと確かめた範囲：字形ごとの ToUnicode・中身のない字形（書体の組み立てが使う空きの
                    # 字形）は、元の字（ActualText）と字形の並びで確かめたので、問題にしない
                    probs[p0:] = [p for p in probs[p0:] if _NO_TU not in p and not (ok and _EMPTY in p)]
                    probs += sp
                    text[t0:] = [actual]
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
                    if cur.kind == "embedded":
                        bs = bytes(s)
                        drawn += [(cur, int.from_bytes(bs[i:i + 2], "big")) for i in range(0, len(bs) - 1, 2)]
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
            try:
                stack.append(_obj(lx, "<<", None))
            except PdfError:
                break
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
