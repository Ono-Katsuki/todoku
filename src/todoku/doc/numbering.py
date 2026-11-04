"""古い Word（.doc）の段落番号（MS-DOC 2.9.147 LSTF・2.9.150 LVL・2.9.131 LFO、sprmPIlfo・sprmPIlvl）。

番号の字はファイルの本文になく、Word が数えて出す。読まないと「１ 申請者」の「１」が窓から消える。
数え方：リスト（lsid）ごと、深さごとに数える。深さの浅い番号が進むと、深い番号は始めに戻る（fNoRestart のときを除く）。
読めない番号の形は、1, 2, 3 にせず、number_text の〔…再現できません〕を出す。リストの表が読めなければ、番号は出さずに unread で伝える。
"""

from __future__ import annotations

import struct

from ..docx.document import WINGDINGS, number_text

# MSONFC（MS-DOC 2.9.157。順は ST_NumberFormat の並びと同じ）。ここにない番号は number_text が〔再現できません〕にする
NFC = {0: "decimal", 1: "upperRoman", 2: "lowerRoman", 3: "upperLetter", 4: "lowerLetter", 5: "ordinal", 6: "cardinalText", 7: "ordinalText",
       10: "ideographDigital", 11: "japaneseCounting", 12: "aiueo", 13: "iroha", 14: "decimalFullWidth", 15: "decimalHalfWidth", 16: "japaneseLegal",
       17: "japaneseDigitalTenThousand", 18: "decimalEnclosedCircle", 19: "decimalFullWidth2", 20: "aiueoFullWidth", 21: "irohaFullWidth", 22: "decimalZero"}
BULLET = 23
NONE = 255
_LVLF = 28
_SPRM_ILFO, _SPRM_ILVL = 0x460B, 0x260A


def _lvl(t: bytes, p: int) -> tuple[dict, int]:
    start, nfc, flags = struct.unpack_from("<iBB", t, p)
    cb_papx, cb_chpx = t[p + 24], t[p + 25]
    q = p + _LVLF + cb_papx + cb_chpx
    cch = struct.unpack_from("<H", t, q)[0]
    if cch > 9 * 2 + 256 or q + 2 + 2 * cch > len(t):
        raise ValueError("番号の字の長さが合いません")
    xst = t[q + 2:q + 2 + 2 * cch].decode("utf-16-le", "replace")
    return {"start": start, "nfc": nfc, "legal": bool(flags & 4), "no_restart": bool(flags & 8), "xst": xst}, q + 2 + 2 * cch


def read_lists(fc: dict, t: bytes) -> tuple[dict, dict]:
    """({lsid: [深さごとの LVL]}, {ilfo: {"lsid", "over": {深さ: 始まりの数}, "fmt": {深さ: LVL}}})。"""
    lists: dict = {}
    lfos: dict = {}
    at, lcb = fc.get("PlfLst", (0, 0))
    if lcb >= 2:
        n = struct.unpack_from("<H", t, at)[0]
        heads = []
        for k in range(n):
            q = at + 2 + 28 * k
            heads.append((struct.unpack_from("<I", t, q)[0], bool(t[q + 26] & 1)))
        p = at + lcb
        for lsid, simple in heads:
            levels = []
            for _ in range(1 if simple else 9):
                lv, p = _lvl(t, p)
                levels.append(lv)
            lists[lsid] = levels
    at, lcb = fc.get("PlfLfo", (0, 0))
    if lcb >= 4:
        n = struct.unpack_from("<I", t, at)[0]
        p = at + 4 + 16 * n
        for k in range(n):
            q = at + 4 + 16 * k
            lsid = struct.unpack_from("<I", t, q)[0]
            count = t[q + 12]
            over: dict = {}
            fmt: dict = {}
            p += 4  # LFOData.cp
            for _ in range(count):
                start, flags = struct.unpack_from("<iI", t, p)
                p += 8
                ilvl = flags & 0xF
                if flags & 0x10:
                    over[ilvl] = start
                if flags & 0x20:
                    lv, p = _lvl(t, p)
                    fmt[ilvl] = lv
                    over.setdefault(ilvl, lv["start"])
            lfos[k + 1] = {"lsid": lsid, "over": over, "fmt": fmt}
    return lists, lfos


def _digit(n: int, lv: dict) -> str:
    if lv["legal"]:
        return str(n)
    if lv["nfc"] in (BULLET, NONE):
        return ""
    return number_text(n, NFC.get(lv["nfc"], f"番号の形{lv['nfc']}"))


def label_text(levels: list, ilvl: int, counts: list, fmt: dict) -> str:
    lv = fmt.get(ilvl) or levels[ilvl]
    out = []
    for ch in lv["xst"]:
        o = ord(ch)
        if o < 9:
            if o < len(levels):
                out.append(_digit(counts[o] if counts[o] is not None else (fmt.get(o) or levels[o])["start"], fmt.get(o) or levels[o]))
        elif 0xF000 <= o <= 0xF0FF:
            out.append(WINGDINGS.get(o, "•"))  # 記号の書体の印（字としての意味はない）
        else:
            out.append(ch)
    return "".join(out)


class Numberer:
    def __init__(self, lists: dict, lfos: dict):
        self.lists, self.lfos = lists, lfos
        self.counts: dict[int, list] = {}
        self.used_over: set = set()

    def next(self, ilfo: int, ilvl: int) -> str:
        lfo = self.lfos.get(ilfo)
        levels = self.lists.get(lfo["lsid"]) if lfo else None
        if not levels:
            return ""
        ilvl = min(ilvl, len(levels) - 1)
        counts = self.counts.setdefault(lfo["lsid"], [None] * len(levels))
        lv = lfo["fmt"].get(ilvl) or levels[ilvl]
        if (ilfo, ilvl) in self.used_over or ilvl not in lfo["over"]:
            counts[ilvl] = lv["start"] if counts[ilvl] is None else counts[ilvl] + 1
        else:
            self.used_over.add((ilfo, ilvl))
            counts[ilvl] = lfo["over"][ilvl]
        for d in range(ilvl + 1, len(counts)):
            if not (lfo["fmt"].get(d) or levels[d])["no_restart"]:
                counts[d] = None
        return label_text(levels, ilvl, counts, lfo["fmt"])
