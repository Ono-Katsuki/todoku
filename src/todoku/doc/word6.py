"""Word 6.0 / Word 95（.doc、FIB の版 101〜104）を読み、字数（バイトの数）を変えない上書きで書いて、証明する。

形は「Microsoft Word for Windows 6.0 Binary File Format」（corpus/specs/Word6-binary-file-format.txt）で確かめた：
- 複合ファイルの WordDocument（「main stream」）だけに、FIB・文字・書式・表がすべてある（Word 97 の 0Table/1Table・Data はない）
- FIB：0 wIdent（0xA5DC）・2 nFib・6 lid（言語）・0x0A 印（0x0004 fComplex・0x0100 暗号）・0x14 chse（0 は Windows の
  ANSI の文字コード、256 は Macintosh）・0x18 fcMin・0x1C fcMac・0x34 から文字の数（本文・脚注・ヘッダー・マクロ・注釈・
  文末脚注・テキストボックス・ヘッダーのテキストボックス。4 バイトずつ）・0x58 から表の位置と大きさ（4 バイトずつの組）
  ・0x18E cpnBteChp・0x190 cpnBtePap（書式の FKP の数）。仕様書の 10 進の位置には誤りがあるので、16 進の位置を使う
- 文字は 8 ビット：位置（CP）はバイトの位置で、全角の字（2 バイト）は 2 つの CP を使う。文字コードは FIB の lid の言語の
  Windows の ANSI の文字コード（日本語 0x0411 は cp932）
- 高速保存していない文書（fComplex=0）は、CP をそのまま fcMin からのバイトの位置にする。高速保存した文書は CLX の
  ピースの表（PLC：CP の並び・8 バイトの PCD＝印 2・FC 4・prm 2）でたどる。prm は 1 つの sprm か、CLX の grpprl の番号
- 書式の FKP は 512 バイト。文字（CHPX）：FC の並び・1 バイトの位置・CHPX（長さ 1 バイト＋grpprl）。段落（PAPX）：FC の並び・
  7 バイトの BX（最初の 1 バイトが位置）・PAPX（語の数 1 バイト・istd 2 バイト・grpprl）。ビンの表の BTE は 2 バイトのページ番号。
  高速保存していない文書のビンの表は、cpnBte より少なければ続くページで補う
- sprm は 1 バイトの番号と、番号ごとに決まった長さの値（可変のものは長さ 1 バイトが続く）。表の中：24（sprmPFInTable）・
  行の終わり：25（sprmPTtp）・非表示：92（sprmCFVanish）・削除の変更履歴：65・下線：94・特別な字：117・絵の位置：68・
  記号：74・文字コードの違い：73
仕様書にない sprm の番号（日本語版の Word の禁則などの命令かもしれない）は、長さが分からず、その後ろの命令を読めないので、
推測せずに読むのをやめる。フォームの欄の FFDATA の形も仕様書にないので、チェックボックスとドロップダウンの状態は読まず、
「状態は読めません」と示す（書かない）。テキストのフォームの欄は、結果の字だけを読む。

書き込みは、字のバイト数を変えない上書きだけ（一太郎と同じ考え方）：空白・伏せ字・□ の上に、そのバイト数までの値を
文書の文字コードで書く（残りは空白で埋める）。位置の表・書式の表・ピースの表は 1 バイトも変えないので、ずらす必要がない。
"""

from __future__ import annotations

import bisect
import hashlib
import struct
from collections import Counter
from dataclasses import dataclass, field

from ..cfb import CFB
from ..docx.document import Seg, iter_paras
from ..docx.fill import FillError, Patch, _as_bool, _is_text_value
from .book import FC_NAMES, CPRef, DocError, DocFile, Piece, Run
from ..xlsx.audit import AUDIT_BOXES  # 検査で数えない箱の字（Word の仲間と同じ）

WORD6_NFIB = range(101, 105)

# FIB（Word 6）の表の位置と大きさ：名前 → FIB の中の位置（16 進の列で確かめた）
FIB6 = {"StshfOrig": 0x58, "Stshf": 0x60, "PlcffndRef": 0x68, "PlcffndTxt": 0x70, "PlcfandRef": 0x78, "PlcfandTxt": 0x80,
        "PlcfSed": 0x88, "PlcPad": 0x90, "PlcfPhe": 0x98, "SttbfGlsy": 0xA0, "PlcfGlsy": 0xA8, "PlcfHdd": 0xB0,
        "PlcfBteChpx": 0xB8, "PlcfBtePapx": 0xC0, "PlcfSea": 0xC8, "SttbfFfn": 0xD0, "PlcfFldMom": 0xD8, "PlcfFldHdr": 0xE0,
        "PlcfFldFtn": 0xE8, "PlcfFldAtn": 0xF0, "PlcfFldMcr": 0xF8, "SttbfBkmk": 0x100, "PlcfBkf": 0x108, "PlcfBkl": 0x110,
        "Cmds": 0x118, "PlcMcr": 0x120, "SttbfMcr": 0x128, "PrDrvr": 0x130, "PrEnvPort": 0x138, "PrEnvLand": 0x140,
        "Wss": 0x148, "Dop": 0x150, "SttbfAssoc": 0x158, "Clx": 0x160}

_CHS_CP = {128: "cp932", 129: "cp949", 134: "cp936", 136: "cp950", 161: "cp1253", 162: "cp1254", 163: "cp1258", 177: "cp1255", 178: "cp1256", 186: "cp1257",
           204: "cp1251", 222: "cp874", 238: "cp1250"}  # FFN の chs → その文字セットの Windows の文字コード
# lid（言語）→ その言語の Windows の ANSI の文字コード
_LID_CP = {0x11: "cp932", 0x12: "cp949", 0x1E: "cp874", 0x01: "cp1256", 0x0D: "cp1255", 0x08: "cp1253", 0x1F: "cp1254",
           0x19: "cp1251", 0x22: "cp1251", 0x23: "cp1251", 0x02: "cp1251", 0x05: "cp1250", 0x0E: "cp1250", 0x15: "cp1250",
           0x1A: "cp1250", 0x1B: "cp1250", 0x24: "cp1250", 0x18: "cp1250", 0x25: "cp1257", 0x26: "cp1257", 0x27: "cp1257"}
_WESTERN = {0x03, 0x06, 0x07, 0x09, 0x0A, 0x0B, 0x0C, 0x0F, 0x10, 0x13, 0x14, 0x16, 0x1D, 0x2D, 0x36, 0x38, 0x3E, 0x41}

# sprm の値の長さ（仕様書の SPRM DEFINITIONS の表）。-1：長さ 1 バイトが続く。-2：長さ 2 バイトが続き、その値から 1 を引く。
# -3：長さ 2 バイトが続く。23（sprmPChgTabs）は別の決まり
SPRM_LEN: dict[int, int] = {2: 2, 3: -1, 4: 1, 5: 1, 6: 1, 7: 1, 8: 1, 9: 1, 10: 1, 11: 1, 12: -1, 13: 1, 14: 1, 15: -1, 16: 2,
                            17: 2, 18: 2, 19: 2, 20: 4, 21: 2, 22: 2, 24: 1, 25: 1, 26: 2, 27: 2, 28: 2, 29: 1, 30: 2, 31: 2,
                            32: 2, 33: 2, 34: 2, 35: 2, 36: 2, 37: 1, 38: 2, 39: 2, 40: 2, 41: 2, 42: 2, 43: 2, 44: 1, 45: 2,
                            46: 2, 47: 2, 48: 2, 49: 2, 50: 1, 51: 1,
                            65: 1, 66: 1, 67: 1, 68: -1, 69: 2, 70: 4, 71: 1, 72: 2, 73: 3, 74: -1, 75: 1, 80: 2, 81: -1, 82: -1,
                            83: 0, 85: 1, 86: 1, 87: 1, 88: 1, 89: 1, 90: 1, 91: 1, 92: 1, 93: 2, 94: 1, 95: 3, 96: 2, 97: 2,
                            98: 1, 99: 1, 100: 1, 101: 1, 102: 1, 103: -1, 104: 1, 105: -1, 106: -1, 107: 2, 108: -1, 109: 2,
                            110: 2, 117: 1, 118: 1, 119: 1, 120: -1, 121: 2, 122: 2, 123: 2, 124: 2,
                            131: 1, 132: 1, 133: -1, 136: 3, 137: 3, 138: 1, 139: 1, 140: 2, 141: 2, 142: 1, 143: 1, 144: 2,
                            145: 2, 146: 1, 147: 1, 148: 2, 149: 2, 150: 1, 151: 1, 152: 1, 153: 1, 154: 2, 155: 2, 156: 2,
                            157: 2, 158: 1, 159: 1, 160: 2, 161: 2, 162: 1, 164: 2, 165: 2, 166: 2, 167: 2, 168: 2, 169: 2,
                            170: 2, 171: 2,
                            182: 2, 183: 2, 184: 2, 185: 1, 186: 1, 187: 12, 188: -2, 189: 2, 190: -2, 191: -3, 192: 4, 193: 5,
                            194: 4, 195: 2, 196: 4, 197: 2, 198: 2, 199: 5, 200: 4}


def is_word6(blob: bytes) -> bool:
    try:
        c = CFB(blob)
        if not c.has("WordDocument"):
            return False
        ident, nfib = struct.unpack_from("<HH", c.read("WordDocument"), 0)
    except Exception:  # noqa: BLE001
        return False
    return (ident == 0xA5DC or ident == 0xA5EC) and nfib in WORD6_NFIB


def sprms6(grpprl: bytes, where: str = "", pad_ok: bool = False):
    """grpprl の sprm を順に (番号, 値のバイト列)。仕様書にない番号・長さが合わないものは DocError。"""
    i, n = 0, len(grpprl)
    while i < n:
        op = grpprl[i]
        if op == 0 and pad_ok and i == n - 1:
            return  # PAPX の、語の境に合わせる 0 のバイト
        size = SPRM_LEN.get(op)
        if op == 23:  # sprmPChgTabs：長さ 1 バイト。255 のときは、消す数・足す数から数える
            if i + 1 >= n:
                raise DocError(f"{where}書式の命令（sprm 23）が途中で終わっています")
            size = grpprl[i + 1]
            at = i + 2
            if size == 255:
                dels = grpprl[at]
                adds = grpprl[at + 1 + 4 * dels]
                size = 2 + dels * 4 + adds * 3
        elif size is None:
            raise DocError(f"{where}書式の命令（sprm の番号 {op}）は Word 6.0 の仕様書にないため、長さが分からず、"
                           "文書を正しく読めません（読み違えた字や表を出さないよう、読むのをやめました）")
        elif size == -1:
            if i + 1 >= n:
                raise DocError(f"{where}書式の命令（sprm {op}）が途中で終わっています")
            size, at = grpprl[i + 1], i + 2
        elif size in (-2, -3):
            if i + 3 > n:
                raise DocError(f"{where}書式の命令（sprm {op}）が途中で終わっています")
            size = struct.unpack_from("<H", grpprl, i + 1)[0] - (1 if size == -2 else 0)
            at = i + 3
        else:
            at = i + 1
        if size < 0 or at + size > n:
            raise DocError(f"{where}書式の命令（sprm {op}）の長さが、書式の大きさと合いません")
        yield op, grpprl[at:at + size]
        i = at + size


class Word6File(DocFile):
    """Word 6.0/95 の文書。窓・書き込みは Word 97 以降の .doc と同じ形（段落・表・マス）にする。"""

    kind = "doc"

    def __init__(self, blob: bytes, filename: str = ""):  # noqa: C901
        self.blob = blob
        self.filename = filename
        self.cfb = CFB(blob)
        self.embeds = []
        self.word6 = True
        wd = self.cfb.read("WordDocument")
        self.wd = wd
        if len(wd) < 0x192:
            raise DocError("Word 6.0/95 の文書の FIB が短すぎます")
        ident, self.nfib, _prod, self.lid = struct.unpack_from("<HHHH", wd, 0)
        self.flags = struct.unpack_from("<H", wd, 0x0A)[0]
        if self.flags & 0x0100:
            raise DocError("パスワードで暗号化された Word 6.0/95 の文書です")
        chse = struct.unpack_from("<H", wd, 0x14)[0]
        if chse != 0:
            raise DocError(f"Word 6.0/95 の文書の文字は、Windows の ANSI ではない文字コード（chse={chse}。256 は Macintosh）で"
                           "書かれています。この文字コードは確かめていないので、読み違えた字を出さないよう、読むのをやめました")
        prim = self.lid & 0x3FF
        if self.lid == 0x0404 or self.lid == 0x0C04 or self.lid == 0x1404:
            self.enc = "cp950"
        elif self.lid in (0x0804, 0x1004):
            self.enc = "cp936"
        elif prim in _LID_CP:
            self.enc = _LID_CP[prim]
        elif prim in _WESTERN:
            self.enc = "cp1252"
        else:
            raise DocError(f"Word 6.0/95 の文書の言語（lid 0x{self.lid:04X}）から、文字コードを決められません")
        self.complex = bool(self.flags & 0x0004)
        self.fc_min, self.fc_mac = struct.unpack_from("<II", wd, 0x18)
        c = struct.unpack_from("<8i", wd, 0x34)
        self.ccp = {"text": c[0], "ftn": c[1], "hdd": c[2], "mcr": c[3], "atn": c[4], "edn": c[5], "txbx": c[6], "hdrtxbx": c[7]}
        if any(v < 0 for v in c):
            raise DocError("Word 6.0/95 の文書の文字の数が壊れています")
        self.fc = {name: struct.unpack_from("<II", wd, at) for name, at in FIB6.items()}
        for name, (f0, lcb) in self.fc.items():
            if lcb and f0 + lcb > len(wd):
                raise DocError(f"Word 6.0/95 の文書の表（{name}）が、ストリームの外を指しています")
        self.table = wd  # 表はすべて WordDocument の中
        self.table_name = "WordDocument"
        self.font_cps = self._font_table()
        self.data = b""
        self.lw = list(c)
        total = sum(c)
        self.cp_total = total + (1 if any(c[1:]) else 0)
        self.pieces = self._pieces6()
        if self.pieces[-1].cp1 < total:
            raise DocError("ピースの表が、文書の終わりまで届いていません")
        self.grpprls: list[bytes] = getattr(self, "grpprls", [])
        self.chpx = self._fkps6("PlcfBteChpx", papx=False, cpn_at=0x18E)
        self.papx = self._fkps6("PlcfBtePapx", papx=True, cpn_at=0x190)
        self._chpx_fc0 = [r.fc0 for r in self.chpx]
        self._papx_fc0 = [r.fc0 for r in self.papx]
        self._enc_runs = self._font_encodings()
        self.text, self._w = self._text6()
        self.charset_warning = self._charset_note()
        self.protection, self.protection_enforced = self._dop()
        self.has_macros = False
        self.controls = []
        self.bookmarks = []
        self.stories = []
        self._build()

    @property
    def vba(self):
        return None

    # ------------------------------------------------------------------ ピース
    def _pieces6(self) -> list[Piece]:
        wd = self.wd
        if not self.complex:
            n = self.fc_mac - self.fc_min
            if n not in (sum(self.lw), sum(self.lw) + 1) or self.fc_mac > len(wd):
                raise DocError(f"Word 6.0/95 の文書の文字の範囲（fcMin〜fcMac）が、文字の数と合いません（{n} と {sum(self.lw)}）")
            return [Piece(0, n, self.fc_min, True, -1)]
        f0, lcb = self.fc["Clx"]
        if not lcb:
            raise DocError("高速保存された Word 6.0/95 の文書に、ピースの表（CLX）がありません")
        end = f0 + lcb
        i = f0
        self.grpprls = []
        while i < end:
            kind = wd[i]
            if kind == 1:
                cb = struct.unpack_from("<H", wd, i + 1)[0]
                self.grpprls.append(wd[i + 3:i + 3 + cb])
                i += 3 + cb
                continue
            if kind != 2:
                raise DocError(f"ピースの表（CLX）の区切り（{kind}）が壊れています")
            # 仕様書は「2 バイトの大きさ」。Word 97 以降の形（4 バイト）もありうるので、CLX の終わりにちょうど合うほうだけを使う
            for width in (2, 4):
                size = struct.unpack_from("<H" if width == 2 else "<I", wd, i + 1)[0]
                if i + 1 + width + size == end and (size - 4) % 12 == 0 and size >= 16:
                    break
            else:
                raise DocError("ピースの表（CLX）の大きさが、CLX の終わりと合いません")
            base = i + 1 + width
            n = (size - 4) // 12
            cps = struct.unpack_from(f"<{n + 1}I", wd, base)
            out = []
            for k in range(n):
                pp = base + (n + 1) * 4 + k * 8
                fcv, prm = struct.unpack_from("<IH", wd, pp + 2)
                if cps[k + 1] < cps[k] or fcv + (cps[k + 1] - cps[k]) > len(wd):
                    raise DocError("ピースの表のピースが、ストリームの外を指しています")
                pc = Piece(cps[k], cps[k + 1], fcv, True, pp)
                pc.prm = prm
                out.append(pc)
            self.clx_pcdt = i
            # 同じバイトを 2 つのピースが指していないこと（上書きしたバイトが、ほかの所にも出ないように）
            spans = sorted((p.fc, p.fc + p.cp1 - p.cp0) for p in out)
            self.overlap = any(spans[j + 1][0] < spans[j][1] for j in range(len(spans) - 1))
            return out
        raise DocError("ピースの表（CLX）が見つかりません")

    def _prm_sprms(self, piece: Piece) -> list[tuple[int, bytes]]:
        prm = getattr(piece, "prm", 0)
        if not prm:
            return []
        if prm & 1:
            k = prm >> 1
            if k >= len(self.grpprls):
                raise DocError("ピースの書式（prm）が、ない grpprl を指しています")
            return list(sprms6(self.grpprls[k], "ピースの書式の"))
        op, val = (prm >> 1) & 0x7F, prm >> 8
        size = SPRM_LEN.get(op)
        if op == 0:
            return []
        if size is None or size < 0 or size > 1:
            raise DocError(f"ピースの書式（prm）の sprm {op} を、1 バイトの値として読めません")
        return [(op, bytes([val])[:size])]

    # ------------------------------------------------------------------ 書式（FKP）
    def _fkps6(self, name: str, papx: bool, cpn_at: int) -> list[Run]:
        f0, lcb = self.fc[name]
        if not lcb:
            return []
        n = (lcb - 4) // 6
        if 4 * (n + 1) + 2 * n != lcb:
            raise DocError(f"ビンの表（{name}）の大きさが合いません")
        pns = list(struct.unpack_from(f"<{n}H", self.wd, f0 + (n + 1) * 4))
        cpn = struct.unpack_from("<H", self.wd, cpn_at)[0]
        if not self.complex and cpn > len(pns) and pns:
            # 仕様書：高速保存していない文書のビンの表は、足りない分を続くページで補う
            pns += [pns[-1] + k for k in range(1, cpn - len(pns) + 1)]
        out = []
        for pn in pns:
            page = self.wd[pn * 512:(pn + 1) * 512]
            if len(page) < 512:
                raise DocError(f"書式のページ {pn} が、ストリームの外にあります")
            crun = page[511]
            rgfc = struct.unpack_from(f"<{crun + 1}I", page, 0)
            for k in range(crun):
                if papx:
                    boff = page[4 * (crun + 1) + 7 * k] * 2
                else:
                    boff = page[4 * (crun + 1) + k] * 2
                grpprl = b""
                if boff:
                    if papx:
                        cw = page[boff]
                        grpprl = page[boff + 3:boff + 2 * cw] if cw >= 2 else b""
                    else:
                        cb = page[boff]
                        grpprl = page[boff + 1:boff + 1 + cb]
                list(sprms6(grpprl, f"書式のページ {pn} の", pad_ok=papx))  # すべて読めることを先に確かめる
                out.append(Run(rgfc[k], rgfc[k + 1], grpprl))
        out.sort(key=lambda r: r.fc0)
        return out

    def _find_run(self, runs, starts, fc: int) -> bytes:
        k = bisect.bisect_right(starts, fc) - 1
        if 0 <= k < len(runs) and runs[k].fc0 <= fc < runs[k].fc1:
            return runs[k].grpprl
        return b""

    def fc_of(self, cp: int) -> int:
        for p in self.pieces:
            if p.cp0 <= cp < p.cp1:
                return p.fc + (cp - p.cp0)
        return -1

    def para_props(self, cp_end: int) -> dict:
        fc = self.fc_of(cp_end)
        d = {"in_table": False, "ttp": False, "itap": 0, "inner_cell": False, "inner_ttp": False}
        ops = list(sprms6(self._find_run(self.papx, self._papx_fc0, fc), pad_ok=True))
        pc = self.piece_of(cp_end)
        if pc is not None:
            ops += self._prm_sprms(pc)
        for sp, op in ops:
            if sp == 24:
                d["in_table"] = bool(op[0]) if op else False
            elif sp == 25:
                d["ttp"] = bool(op[0]) if op else False
        if d["in_table"]:
            d["itap"] = 1
        return d

    def char_props(self, cp: int) -> dict:
        fc = self.fc_of(cp)
        d = {"hidden": False, "underline": False, "pic": None, "symbol": None, "spec": False, "data": False, "deleted": False,
             "ole": False}
        ops = list(sprms6(self._find_run(self.chpx, self._chpx_fc0, fc)))
        pc = self.piece_of(cp)
        if pc is not None:
            ops += self._prm_sprms(pc)
        strike_, ico_ = False, 0
        for sp, op in ops:
            if sp == 92 and op:
                # 0・1 はその値。128 はスタイルの値、129 はその逆（スタイルの非表示は読まないので、非表示でないとみなす）
                d["hidden"] = op[0] in (1, 129)
            elif sp == 65 and op:
                d["deleted"] = bool(op[0])
            elif sp == 94 and op:
                d["underline"] = bool(op[0])
            elif sp == 87 and op:  # sprmCFStrike。値の読み方は 92 と同じ（スタイルの字の書式は読まない）
                strike_ = op[0] in (1, 129)
            elif sp == 98 and op:  # sprmCIco（Word 6 の色の番号。Word 97 と同じ表）
                ico_ = op[0]
            elif sp == 68 and len(op) >= 4:
                d["pic"] = struct.unpack_from("<I", op, 0)[0]
                d["spec"] = True
            elif sp == 74 and len(op) >= 3:
                d["symbol"] = (struct.unpack_from("<H", op, 0)[0], op[2])
                d["spec"] = True
            elif sp == 117 and op:
                d["spec"] = bool(op[0])
            elif sp == 71 and op:
                d["data"] = bool(op[0])
            elif sp == 75 and op:
                d["ole"] = bool(op[0])
            elif sp == 73 and op and op[0]:
                raise DocError(f"位置 {cp} の字は、文書と違う文字コード（chse）で書かれています。この形は確かめていないので、読むのをやめました")
        from .book import ICO

        d["look"] = (strike_, ICO.get(ico_))
        return d

    # ------------------------------------------------------------------ 文字
    def _font_table(self) -> list[str | None]:
        """書体の表（SttbfFfn。仕様書：書体の名前の順が ftc の番号）：番号 → その書体の文字セットの文字コード（FFN の chs。分からなければ None）。"""
        f0, lcb = self.fc["SttbfFfn"]
        wd = self.wd
        out: list[str | None] = []
        if lcb <= 2:
            return out
        p, end = f0 + 2, f0 + lcb
        self._font_names: list[str] = []
        while p + 6 < end:
            cb = wd[p] + 1
            out.append(_CHS_CP.get(wd[p + 4]))
            self._font_names.append(wd[p + 6:p + cb].split(b"\0")[0].decode("latin-1"))
            p += cb
        return out

    def _font_encodings(self) -> list[tuple[int, int, str]]:
        """字を、書体が文字コードを決めている所：[(CP の始め, 終わり, 文字コード)]。字ごとの書式の sprmCFtc（93。chp.ftc）で書体を指している所だけ。
        文書の文字コードと同じ所は入れない。スタイルから受けた書体は読まない（_charset_note で伝える）。"""
        if not any(c and c != self.enc for c in self.font_cps):
            return []
        out = []
        for pc in self.pieces:
            prm = dict(self._prm_sprms(pc)).get(93)
            n = pc.cp1 - pc.cp0
            fc0, fc1 = pc.fc, pc.fc + n
            k = max(bisect.bisect_right(self._chpx_fc0, fc0) - 1, 0)
            while k < len(self.chpx) and self.chpx[k].fc0 < fc1:
                r = self.chpx[k]
                k += 1
                x0, x1 = max(r.fc0, fc0), min(r.fc1, fc1)
                if x0 >= x1:
                    continue
                op = prm
                if op is None:
                    for sp, val in sprms6(r.grpprl):
                        if sp == 93:
                            op = val
                if op is None or len(op) < 2:
                    continue
                ftc = struct.unpack_from("<H", op, 0)[0]
                cp = self.font_cps[ftc] if ftc < len(self.font_cps) else None
                if cp and cp != self.enc:
                    out.append((pc.cp0 + x0 - fc0, pc.cp0 + x1 - fc0, cp))
        return sorted(out)

    def enc_at(self, cp: int) -> str:
        """この CP の字の文字コード。"""
        for a, b, enc in self._enc_runs:
            if a <= cp < b:
                return enc
        return self.enc

    def _charset_note(self) -> str | None:
        """書体の表に、文書の言語（lid）から決めた文字コードと違うものがあり、半角でない字がそれを指していない所にある：スタイルから受けた書体かも
        しれないので、その事実を伝える（スタイルの書体までは読まない）。どの字もそのような書体を指していれば、読み替えたので伝えない。"""
        others: dict[str, list[str]] = {}
        for i, cp in enumerate(self.font_cps):
            if cp and cp != self.enc:
                others.setdefault(cp, []).append(self._font_names[i])
        if not others:
            return None
        pos = 0
        high = False
        for p in self.pieces:
            for k in range(p.cp1 - p.cp0):
                if self.wd[p.fc + k] >= 0x80 and not any(a <= p.cp0 + k < b for a, b, _e in self._enc_runs):
                    high = True
                    break
            pos += 1
            if high:
                break
        if not high:
            return None
        which = "・".join(f"{cp}（{'、'.join(n[:3])}）" for cp, n in others.items())
        return (f"この文書の字のうち、字の書式で書体を指している所はその書体の文字コードで読み、ほかは {self.enc} として読みましたが、書体の表に別の文字コード {which} があります。"
                "スタイルからその書体を受けた字は別の字に見えている（字が化けている）かもしれません。LibreOffice か Word で .docx に保存し直してもらえば、確かに読めます")

    def _text6(self) -> tuple[str, list[int]]:
        """CP ごとの字（全角の字は最初の CP に置き、2 つ目の CP は「\\x00」）と、字のバイト数（1・2。2 つ目の CP は 0）。"""
        raw = bytearray()
        for p in self.pieces:
            raw += self.wd[p.fc:p.fc + (p.cp1 - p.cp0)]
        runs = self._enc_runs
        next_run = 0
        enc = self.enc
        dbcs = enc in ("cp932", "cp936", "cp949", "cp950")
        out: list[str] = []
        w: list[int] = []
        i, n = 0, len(raw)
        while i < n:
            if runs:
                while next_run < len(runs) and runs[next_run][1] <= i:
                    next_run += 1
                enc = runs[next_run][2] if next_run < len(runs) and runs[next_run][0] <= i else self.enc
                dbcs = enc in ("cp932", "cp936", "cp949", "cp950")
            b = raw[i]
            lead = dbcs and ((0x81 <= b <= 0x9F or 0xE0 <= b <= 0xFC) if enc == "cp932" else 0x81 <= b <= 0xFE)
            if lead:
                if i + 1 >= n or raw[i + 1] < 0x40:
                    raise DocError(f"位置 {i} の全角の字の 2 バイト目がありません（文字コード {enc}）")
                try:
                    ch = bytes(raw[i:i + 2]).decode(enc)
                except UnicodeDecodeError:
                    raise DocError(f"位置 {i} の字が、文字コード {enc} で読めません") from None
                if "" <= ch <= "":
                    raise DocError(f"位置 {i} の字は外字（利用者が作った字）です。どの字か決められないので、読むのをやめました")
                out += [ch, "\x00"]
                w += [2, 0]
                i += 2
                continue
            if b == 0xA0:
                ch = "\xa0"  # 仕様書：改行しない空白は 160
            elif b < 0x80:
                ch = chr(b)
            else:
                try:
                    ch = bytes([b]).decode(enc)
                except UnicodeDecodeError:
                    raise DocError(f"位置 {i} の字（0x{b:02X}）が、文字コード {enc} で読めません") from None
                if "" <= ch <= "" or ch == "\x80":
                    raise DocError(f"位置 {i} の字（0x{b:02X}）は、文字コード {enc} で決まった字がありません")
            out.append(ch)
            w.append(1)
            i += 1
        return "".join(out), w

    def _build(self) -> None:
        c = self.ccp
        at = 0
        self.body = []
        for kind, key in (("body", "text"), ("footnotes", "ftn"), ("header", "hdd"), (None, "mcr"), ("comments", "atn"),
                          ("endnotes", "edn"), ("textbox", "txbx"), ("header-textbox", "hdrtxbx")):
            n = c[key]
            if n > 0 and kind is not None:  # マクロの文字（WordBasic）は文書の中身ではないので、窓に出さない
                blocks = self._blocks(at, at + n)
                if kind == "body":
                    self.body = blocks
                else:
                    self.stories.append((kind, kind, blocks))
            at += n

    def _add_embed(self, storage: str, prog: str, p) -> dict | None:
        return None  # Word 6.0/95 の埋め込み（ObjectPool）は、まだ読まない

    def _form_field(self, f: dict, p):
        from ..docx.document import Control

        instr = f["instr"].replace("\x00", "").strip()
        word = instr.split()[0].upper() if instr.split() else ""
        if word in ("FORMCHECKBOX", "FORMDROPDOWN"):
            # 状態は FFDATA にあるが、Word 6.0 の FFDATA の形は仕様書にない。読み違えた状態を示さない
            what = "チェックボックス" if word == "FORMCHECKBOX" else "ドロップダウン"
            p.segs.append(Seg(f"〔フォームの{what}：状態は読めません（Word 6.0/95 の欄のデータの形が分からないため）〕",
                              "other", CPRef(f["begin"]), None))
            self.unread_controls = getattr(self, "unread_controls", 0) + 1  # 段階は、この数で決める（窓の字は数えない）
            return None
        if word != "FORMTEXT":
            return None
        ctl = Control("text", "formfield", "", None, "doc", p)
        ctl.begin_cp = f["begin"]
        self.controls.append(ctl)
        return ctl

    def width_at(self, cp: int) -> int:
        return self._w[cp] if 0 <= cp < len(self._w) else 0


# ====================================================================== 窓
def para_cps6(doc: Word6File, p) -> list[int | None]:
    """段落の文字（p.text）の 1 字ごとの CP。書き換えられない字（〔…〕などの説明）は None。"""
    out: list[int | None] = []
    for sg in p.segs:
        if sg.hidden:
            continue
        for k in range(len(sg.text)):
            if sg.kind == "t" and sg.node is not None:
                cp = sg.node.s
                # 結合した字の並び：前の字の幅を足していく
                for j in range(k):
                    cp += doc.width_at(cp) or 1
                out.append(cp)
            else:
                out.append(None)
    return out


SPACES = " 　\xa0"


def _room(doc: Word6File, cps: list[int | None]) -> int:
    if any(c is None for c in cps):
        return 0
    return sum(doc.width_at(c) for c in cps)


def capacity(doc: Word6File, it: dict) -> int:
    """その欄に上書きできるバイト数（0 は書けない所）。"""
    c = it.get("control")
    if c is not None:
        cps = [sg.node.s for sg in c.result_segs if sg.kind == "t" and sg.node is not None]
        if not cps:
            return 0
        segs = [sg for sg in c.result_segs if sg.kind == "t" and sg.node is not None and not sg.hidden]
        n = 0
        for sg in segs:
            cp = sg.node.s
            for _ in sg.text:
                n += doc.width_at(cp)
                cp += doc.width_at(cp) or 1
        return n
    p = it.get("para")
    if p is None:
        return 0
    cps = para_cps6(doc, p)
    kind = it.get("kind")
    if kind == "cell":
        t = p.text
        if t and not t.strip(SPACES):
            return _room(doc, cps[:len(t)])
        return 0
    a, b = it.get("span", (0, 0))
    if b <= a or b > len(cps):
        return 0
    if kind == "checkbox":
        return 2 if p.text[a:b] == "□" and _room(doc, cps[a:b]) == 2 else 0
    return _room(doc, cps[a:b])


def build(doc: Word6File) -> dict:
    from ..docx.window import Spots
    from ..docx.window import build as wbuild

    win = wbuild(doc)
    win["unread_controls"] = getattr(doc, "unread_controls", 0)  # 状態を読めないチェックボックス・ドロップダウンの数
    if doc.charset_warning:
        win.setdefault("unread", []).append(doc.charset_warning)
    sp = Spots(doc)
    for f in win["fields"]:
        it = sp.items.get(f["id"])
        room = capacity(doc, it) if it is not None else 0
        if f["kind"] == "checkbox":
            if not room:
                f["readonly"] = True
            continue
        if room:
            # 書けるのはバイトの数まで（全角の字は 2 バイト・半角の字は 1 バイト）。窓の「N 字まで」は全角の字で数える
            f["maxlen"] = max(room // 2, 1) if room >= 2 else room
            f["maxlen_bytes"] = room
        else:
            f["readonly"] = True
    win["word6"] = {"nfib": doc.nfib, "encoding": doc.enc, "fast_saved": doc.complex}
    return win


# ====================================================================== 書き込み
LIMIT = ("Word 6.0/95 の文書は、字のバイト数を変えずに上書きします（字数を変えると、位置の表・書式の表・ピースの表をずらす必要があり、"
         "その書き込みと証明をまだ用意していないため）")


@dataclass
class W6FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    edits: dict = field(default_factory=dict)  # FC（WordDocument の中の位置）→ 書くバイト
    expect: dict = field(default_factory=dict)  # 段落の番号 → 書いた後の段落の文字


def _encode(doc: Word6File, key: str, text: str) -> bytes:
    for ch in text:
        if ord(ch) < 0x20:
            raise FillError(f"{key}: 改行・制御の字は書けません")
    try:
        raw = text.encode(doc.enc)
    except UnicodeEncodeError as e:
        raise FillError(f"{key}: 「{text[e.start:e.end]}」は、この文書の文字コード（{doc.enc}）で書けない字です") from None
    if any(ch != "\xa0" and ch.encode(doc.enc) == b"\xa0" for ch in text):
        raise FillError(f"{key}: この字は文書の中で別の意味（改行しない空白）になるので書けません")
    return raw


def _pad(n: int) -> str:
    """残りのバイトを埋める空白（全角の空白と、端数の半角の空白）。"""
    return "　" * (n // 2) + " " * (n % 2)


def plan(doc: Word6File, values: dict, *, force: bool = False):
    """書き込み → (FC → バイト, 書いた所, 知らせること, 段落 → 書いた後の文字)。書き込みと証明で同じものを使う。"""
    from ..docx.window import Spots, WriteError, resolve_write

    sp = Spots(doc)
    paras = [q for _k, _p, bl in doc.all_blocks() for q in iter_paras(bl)]
    index = {id(q): i for i, q in enumerate(paras)}
    edits: dict[int, int] = {}
    patches: list[Patch] = []
    warnings: list[str] = []
    problems: list[str] = []
    texts: dict[int, str] = {}
    used: dict = {}
    if getattr(doc, "overlap", False):
        raise FillError("この Word 6.0/95 の文書は、同じ字のバイトを 2 か所から使っています（高速保存の結果）。上書きすると別の所の字も"
                        "変わるので、書きません。Word で「高速保存しない」で保存し直した文書なら書けます")
    for key, value in values.items():
        try:
            anchor = None
            if key.startswith("@"):
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    it = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                anchor = dict(value)
                fid, text = it["id"], it["value"]
            else:
                it = sp.items.get(key)
                if it is None:
                    raise FillError(f"{key}: そのような欄はありません{sp.items.hint(key)}")
                fid, text = key, None
            if doc.protection_enforced and not force:
                if doc.protection != "forms" or it.get("control") is None:
                    raise FillError(f"{key}: この文書は保護されていて（{doc.protection}）、ここには書けません")
            c = it.get("control")
            p = it.get("para") or (c.para if c is not None else None)
            kind = it.get("kind")
            cps_all = para_cps6(doc, p)
            if c is not None:
                pos = [i for i, sg in enumerate(p.segs) if any(sg is x for x in c.result_segs)]
                ptext_idx = []
                k = 0
                for i, sg in enumerate(p.segs):
                    if sg.hidden:
                        continue
                    for _ in sg.text:
                        if i in pos and sg.kind == "t":
                            ptext_idx.append(k)
                        k += 1
                if not ptext_idx:
                    raise FillError(f"{key}: このフォームの欄には字がないので書けません。{LIMIT}")
                a, b = ptext_idx[0], ptext_idx[-1] + 1
            elif kind == "checkbox":
                a, b = it["span"]
            elif kind == "cell":
                if p.text.strip(SPACES):
                    raise FillError(f"{key}: このマスにはすでに字があります。{LIMIT}。空白のマスにだけ書けます")
                a, b = 0, len(p.text)
            else:
                a, b = it["span"]
            cps = cps_all[a:b]
            room = _room(doc, cps) if b > a else 0
            if kind == "checkbox" and c is None:
                on = _as_bool(value)
                if p.text[a:b] != "□" or room != 2:
                    raise FillError(f"{key}: チェックの印（□）を書き換えられません")
                new_text = "■" if on else "□"
                patches.append(Patch(fid, "control", on, "doc", key=key))
            else:
                if text is None:
                    text = _is_text_value(value)
                if room == 0:
                    raise FillError(f"{key}: ここには字がないので書けません。{LIMIT}。字のない所（空のマス・段落の終わり）に字を足すことは"
                                    "できません。空白や伏せ字のある所を、字を目印にして上書きしてください")
                raw = _encode(doc, key, text)
                if len(raw) > room:
                    raise FillError(f"{key}: 「{text}」は {len(raw)} バイトです（全角の字は 2 バイト）。{LIMIT}。ここに書けるのは "
                                    f"{room} バイト（全角 {room // 2} 字）までです")
                pad = _pad(room - len(raw))
                new_text = text + pad
                patches.append(Patch(fid, "input", text, "doc", anchor=anchor, key=key))
            if any(doc.enc_at(cp) != doc.enc for cp in cps):
                raise FillError(f"{key}: ここの字は、文書と違う文字コード（{doc.enc_at(next(cp for cp in cps if doc.enc_at(cp) != doc.enc))}）の書体で書かれているので、"
                                "正しい字になるかを確かめられず、書きません。別の所に書くか、.docx に保存し直してください")
            raw_new = new_text.encode(doc.enc)
            if len(raw_new) != room:
                raise FillError(f"{key}: 書くバイトの数が合いません（{len(raw_new)} と {room}）")
            fcs = []
            for cp in cps:
                for j in range(doc.width_at(cp)):
                    fcs.append((cp + j, doc.fc_of(cp + j)))
            # 全角の字が、書式の区切り（文字の書式・ピース）をまたがないこと
            q = 0
            for ch in new_text:
                wch = len(ch.encode(doc.enc))
                if wch == 2:
                    (c0, f0), (c1, f1) = fcs[q], fcs[q + 1]
                    if f1 != f0 + 1 or any(r.fc0 == f1 for r in doc.chpx) or doc.piece_of(c0) is not doc.piece_of(c1):
                        raise FillError(f"{key}: 全角の字が、文字の書式の区切りをまたぐ所に書くことになるので、書けません"
                                        "（半角の字にするか、別の所に書いてください）")
                q += wch
            for (cp, fc), byte in zip(fcs, raw_new):
                if fc < 0:
                    raise FillError(f"{key}: 書く所の位置が分かりません")
                if fc in edits and edits[fc] != byte:
                    raise FillError(f"{key}: ほかの書き込みと同じ字に重なりました")
                if doc.wd[fc] != byte:
                    edits[fc] = byte
            n = index[id(p)]
            base = texts.get(n, p.text)
            texts[n] = base[:a] + new_text + base[b:]
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    return edits, patches, warnings, texts


def apply(doc: Word6File, edits: dict) -> bytes:
    """WordDocument の中の字のバイトだけを書き換えたファイル（ストリームの大きさ・ほかのバイトは変わらない）。"""
    if not edits:
        return doc.blob
    wd = bytearray(doc.wd)
    for fc, b in edits.items():
        wd[fc] = b
    return CFB(doc.blob).replace("WordDocument", bytes(wd))


def fill(doc: Word6File, values: dict, *, force: bool = False) -> W6FillResult:
    from ..embed import is_embed_key

    if any(is_embed_key(k) for k in values):
        raise FillError("Word 6.0/95 の文書の埋め込み（E1!…）には、まだ書けません")
    edits, patches, warnings, texts = plan(doc, values, force=force)
    return W6FillResult(apply(doc, edits), patches, warnings, edits, texts)


# ====================================================================== 証明
@dataclass
class W6Proof:
    original_sha256: str
    new_sha256: str
    streams_total: int = 0
    streams_identical: int = 0
    bytes_written: int = 0
    paragraphs_total: int = 0
    paragraphs_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)
    readback: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors and not self.readback

    def headline(self) -> str:
        return (f"複合ファイルのストリーム {self.streams_identical}/{self.streams_total} がバイト一致。WordDocument で変わったのは、"
                f"上書きした字の {self.bytes_written} バイトだけ。段落 {self.paragraphs_total:,} のうち、変わったのは書いた所のある "
                f"{self.paragraphs_changed} 段落だけ"
                + ("" if self.ok else f"（想定外 {len(self.unexpected) + len(self.lint_errors) + len(self.readback)} 件）"))


def _values_of(written: dict) -> dict:
    out, n = {}, 0
    for fid, w in written.items():
        if isinstance(w, dict) and "@" in w:
            n += 1
            out[f"@{n}"] = w["@"]
        else:
            out[fid] = w
    return out


def prove(original: bytes, new: bytes, written: dict) -> W6Proof:
    """書き込み処理の内部の情報を使わずに、元の文書で書く所を決め直し、前後の 2 つのファイルを読み直して比べる。
    - WordDocument のほかのストリームはバイト一致。WordDocument は大きさが同じで、変わったバイトは書いた所の字のバイトだけ
    - 読み直すと、段落の数・段落の書式（表の中か・行の終わりか）・文字の書式（非表示・下線）が同じで、書いていない段落の字は同じ。
      書いた段落の字は「元の字の、書いた所だけを値（と残りの空白）に替えたもの」"""
    pr = W6Proof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    a = Word6File(original)
    try:
        edits, patches, _w, texts = plan(a, _values_of(written))
    except FillError as e:
        pr.unexpected.append(f"元の文書で、書いた所を決め直せません：{e}")
        return pr
    if {p.id for p in patches} != set(written):
        pr.unexpected.append(f"書いた所が、元の文書で決め直した所と違います（{sorted({p.id for p in patches} ^ set(written))}）")
    ca, cb = CFB(original), CFB(new)
    sa, sb = ca.streams(), cb.streams()
    if set(sa) != set(sb):
        pr.unexpected.append(f"ストリームの構成が変わりました：{sorted(set(sa) ^ set(sb))}")
        return pr
    for name in sa:
        pr.streams_total += 1
        x, y = ca.read(name), cb.read(name)
        if x == y:
            pr.streams_identical += 1
            continue
        if name != "WordDocument":
            pr.unexpected.append(f"ストリーム {name!r} が変わりました（WordDocument のほかは変えないはずです）")
            continue
        if len(x) != len(y):
            pr.unexpected.append(f"WordDocument の大きさが変わりました（{len(x)} → {len(y)}）")
            continue
        diff = [q for q in range(len(x)) if x[q] != y[q]]
        pr.bytes_written = len(diff)
        outside = [q for q in diff if q not in edits]
        if outside:
            pr.unexpected.append(f"WordDocument の {len(outside)} バイトが、書いた所の字の外で変わりました（最初は位置 {outside[0]}）")
        wrong = [q for q, v in edits.items() if y[q] != v]
        if wrong:
            pr.unexpected.append(f"WordDocument の位置 {wrong[0]} などに、書くはずの字が入っていません")
        lo = min(diff) if diff else None
        if lo is not None and lo < 0x200:
            pr.unexpected.append("FIB（WordDocument の頭）が変わりました")
    try:
        b = Word6File(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません：{e}")
        return pr
    if [(p.cp0, p.cp1, p.fc) for p in a.pieces] != [(p.cp0, p.cp1, p.fc) for p in b.pieces] or a.ccp != b.ccp:
        pr.lint_errors.append("書き込み後の文書の、ピースの表か文字の数が変わりました")
    pa = [q for _k, _p, bl in a.all_blocks() for q in iter_paras(bl)]
    pb = [q for _k, _p, bl in b.all_blocks() for q in iter_paras(bl)]
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.readback.append(f"段落の数が変わりました（{len(pa)} → {len(pb)}）")
        return pr
    for n, (x, y) in enumerate(zip(pa, pb)):
        if a.para_props(x.cp1) != b.para_props(y.cp1):
            pr.unexpected.append(f"段落 {n + 1} の書式（表の中か・行の終わりか）が変わっています")
        want = texts.get(n, x.text)
        if n in texts:
            pr.paragraphs_changed += 1
        if y.text != want:
            pr.readback.append(f"段落 {n + 1}：読み直すと「{y.text[:40]}」で、思っていた「{want[:40]}」と違います")
    for cp in range(min(len(a.text), len(b.text))):
        if a.fc_of(cp) in edits:
            pa_, pb_ = a.char_props(cp), b.char_props(cp)
            if (pa_["hidden"], pa_["underline"], pa_["spec"]) != (pb_["hidden"], pb_["underline"], pb_["spec"]):
                pr.unexpected.append(f"位置 {cp} の字の書式が変わっています")
                break
    return pr


# ====================================================================== 検査
def raw_chars(doc: Word6File) -> Counter:
    """窓とは別の読み方：文書の字を CP の順に頭から読み（段落・表に分けない）、フィールドの命令と制御の字を除いて数える。
    マクロの文字（WordBasic）は除く。"""
    c: Counter = Counter()
    start = doc.ccp["text"] + doc.ccp["ftn"] + doc.ccp["hdd"]
    skip = range(start, start + doc.ccp["mcr"])
    depth: list[str] = []
    for cp, ch in enumerate(doc.text[:sum(doc.ccp.values())]):
        if cp in skip:
            continue
        o = ord(ch)
        if o == 0x13:
            depth.append("instr")
            continue
        if o == 0x14 and depth:
            depth[-1] = "result"
            continue
        if o == 0x15 and depth:
            depth.pop()
            continue
        if (depth and depth[-1] == "instr") or o < 0x20 or ch.isspace() or ch in AUDIT_BOXES:
            continue
        pr = doc.char_props(cp)
        if o == 0x28 and pr["symbol"] and pr["spec"]:
            continue  # 記号の字（Word 97 と同じく、窓では記号として出す）
        c[ch.replace("\xa0", " ")] += 1
    return c


def run(blob: bytes, filename: str = ""):
    from ..docx.fill import written_of
    from ..docx.window import markdown_parts
    from ..xlsx.audit import Audit, Check, agent_places_check

    au = Audit(filename)
    try:
        doc = Word6File(blob, filename)
    except DocError as e:
        au.checks.append(Check("開く", False, f"Word 6.0/95 の文書として読めません：{e}"))
        return au
    win = build(doc)
    raw = raw_chars(doc)
    shown: Counter = Counter()
    for st in win["stories"]:
        for t in st["_plain"]:
            for ch in t:
                if not ch.isspace() and ch not in AUDIT_BOXES:
                    shown[ch] += 1
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    au.checks.append(Check("読み落とし（文書の文字）", not missing and not extra,
                           f"本文・表・ヘッダー/フッター・脚注・注釈・テキストボックスの字 {sum(raw.values()):,} 字を、窓とは別の読み方"
                           f"（文字コード {doc.enc} で、CP の順に頭から）で拾い、すべて窓に出ていることを確かめました", missing + extra))
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000, f"Markdown {sum(len(t) for _, t in parts):,} 文字（AIが1回で読める目安 15万文字以内）"))
    targets, blocked = {}, []
    for f in win["fields"]:
        if f.get("readonly"):
            blocked.append(f["id"])
            continue
        room = f.get("maxlen_bytes") or 0
        targets[f["id"]] = True if f["kind"] == "checkbox" else ("１２" if room >= 4 else "1")
    try:
        res = fill(Word6File(blob, filename), targets)
    except FillError as e:
        au.checks.append(Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30]))
        return au
    pr = prove(blob, res.blob, written_of(res.patches))
    au.checks.append(Check("全入力欄への書き込み", pr.ok,
                           f"上書きできる入力欄 {len(targets):,} か所すべてに試しの値を書き込みました。{pr.headline()}",
                           [*pr.unexpected, *pr.lint_errors]))
    au.checks.append(Check("読み戻し", not pr.readback,
                           f"書き込み後の文書を読み直し、段落 {pr.paragraphs_total:,} のすべてで、文字が「元の字の、書いた所だけを"
                           "替えたもの」になっていることを確かめました" if not pr.readback else "書き込み後の文書を読み直すと、思っていた文字と違う段落があります",
                           pr.readback))
    rules_blocked = [f["id"] for f in win["fields"] if f.get("readonly") and not f.get("agent")]
    au.checks.append(Check("書けない所", not rules_blocked,
                           "規則で見つけた入力欄は、すべて空白・伏せ字・□ の上に上書きできます" if not rules_blocked else
                           f"規則で見つけた入力欄のうち {len(rules_blocked)} か所は、字のない所（空のマス・段落の終わり）です。"
                           f"{LIMIT}。そこには書けません（窓では「変更できない欄」）", rules_blocked[:30]))
    ap = agent_places_check(win["fields"])
    if ap is not None:
        au.checks.append(ap)
    return au


__all__ = ["Word6File", "is_word6", "build", "fill", "prove", "run", "plan", "capacity", "FC_NAMES"]
