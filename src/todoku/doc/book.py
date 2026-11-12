"""古い Word（.doc、Word 97〜2003 のバイナリ）を読む（MS-DOC）。外部のライブラリは使わない。

- 複合ファイル（CFB）の WordDocument・1Table/0Table・Data のストリーム
- FIB（文書の目次）→ ピースの表（CLX）で、文字の位置（CP）とストリームの中の位置（FC）を結ぶ
- 段落の書式（PAPX）で表のマス・行の終わりを、文字の書式（CHPX）で非表示・下線・記号・フォームの欄のデータの位置を読む
- フィールド（0x13・0x14・0x15）と、フォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN、FFData）
- しおり（SttbfBkmk・PlcfBkf・PlcfBkl）
読んだ結果は、Word（.docx）と同じ形（段落・表・マス・入力欄）にして、窓を共通にする。
"""

from __future__ import annotations

import bisect
import re
import struct
from dataclasses import dataclass, field

from ..cfb import CFB
from .. import symfont
from ..docx.document import TEXT_TYPES, Cell, Control, Para, Seg, Table, _is_form_instr, _mark_result

FC_NAMES = ["StshfOrig", "Stshf", "PlcffndRef", "PlcffndTxt", "PlcfandRef", "PlcfandTxt", "PlcfSed", "PlcPad", "PlcfPhe",
            "SttbfGlsy", "PlcfGlsy", "PlcfHdd", "PlcfBteChpx", "PlcfBtePapx", "PlcfSea", "SttbfFfn", "PlcfFldMom",
            "PlcfFldHdr", "PlcfFldFtn", "PlcfFldAtn", "PlcfFldMcr", "SttbfBkmk", "PlcfBkf", "PlcfBkl", "Cmds", "Unused1",
            "SttbfMcr", "PrDrvr", "PrEnvPort", "PrEnvLand", "Wss", "Dop", "SttbfAssoc", "Clx", "PlcfPgdFtn",
            "AutosaveSource", "GrpXstAtnOwners", "SttbfAtnBkmk", "Unused2", "Unused3", "PlcSpaMom", "PlcSpaHdr",
            "PlcfAtnBkf", "PlcfAtnBkl", "Pms", "FormFldSttbs", "PlcfendRef", "PlcfendTxt", "PlcfFldEdn", "Unused4",
            "DggInfo", "SttbfRMark", "SttbfCaption", "SttbfAutoCaption", "PlcfWkb", "PlcfSpl", "PlcftxbxTxt",
            "PlcfFldTxbx", "PlcfHdrtxbxTxt", "PlcffldHdrTxbx", "StwUser", "SttbTtmbd", "CookieData", "PgdMotherOldOld",
            "BkdMotherOldOld", "PgdFtnOldOld", "BkdFtnOldOld", "PgdEdnOldOld", "BkdEdnOldOld", "SttbfIntlFld",
            "RouteSlip", "SttbSavedBy", "SttbFnm", "PlfLst", "PlfLfo", "PlcfTxbxBkd", "PlcfTxbxHdrBkd", "DocUndoWord9",
            "RgbUse", "Usp", "Uskf", "PlcupcRgbUse", "PlcupcUsp", "SttbGlsyStyle", "Plgosl", "Plcocx", "PlcfBteLvc",
            "dwLowDateTime_dwHighDateTime", "PlcfLvcPre10", "PlcfAsumy", "PlcfGram", "SttbListNames", "SttbfUssr",
            # Word 2000
            "PlcfTch", "RmdThreading", "Mid", "SttbRgtplc", "MsoEnvelope", "PlcfLad", "RgDofr", "Plcosl", "PlcfCookieOld",
            "PgdMotherOld", "BkdMotherOld", "PgdFtnOld", "BkdFtnOld", "PgdEdnOld", "BkdEdnOld",
            # Word 2002
            "Unused1_2002", "PlcfPgp", "Plcfuim", "PlfguidUim", "AtrdExtra", "Plrsid", "SttbfBkmkFactoid", "PlcfBkfFactoid",
            "Plcfcookie", "PlcfBklFactoid", "FactoidData", "DocUndo", "SttbfBkmkFcc", "PlcfBkfFcc", "PlcfBklFcc",
            "SttbfbkmkBPRepairs", "PlcfbkfBPRepairs", "PlcfbklBPRepairs", "PmsNew", "ODSO", "PlcfpmiOldXP", "PlcfpmiNewXP",
            "PlcfpmiMixedXP", "Unused2_2002", "Plcffactoid", "PlcflvcOldXP", "PlcflvcNewXP", "PlcflvcMixedXP",
            # Word 2003
            "Hplxsdr", "SttbfBkmkSdt", "PlcfBkfSdt", "PlcfBklSdt", "CustomXForm", "SttbfBkmkProt", "PlcfBkfProt",
            "PlcfBklProt", "SttbProtUser", "Unused_2003", "PlcfpmiOld", "PlcfpmiOldInline", "PlcfpmiNew",
            "PlcfpmiNewInline", "PlcflvcOld", "PlcflvcOldInline", "PlcflvcNew", "PlcflvcNewInline", "PgdMother",
            "BkdMother", "AfdMother", "PgdFtn", "BkdFtn", "AfdFtn", "PgdEdn", "BkdEdn", "AfdEdn", "Afd",
            # Word 2007
            "Plcfmthd", "SttbfBkmkMoveFrom", "PlcfBkfMoveFrom", "PlcfBklMoveFrom", "SttbfBkmkMoveTo", "PlcfBkfMoveTo",
            "PlcfBklMoveTo", "Unused1_2007", "Unused2_2007", "Unused3_2007", "SttbfBkmkArto", "PlcfBkfArto",
            "PlcfBklArto", "ArtoData", "Unused4_2007", "Unused5_2007", "Unused6_2007", "OssTheme", "ColorSchemeMapping"]


class DocError(ValueError):
    pass


@dataclass
class Piece:
    cp0: int
    cp1: int
    fc: int  # 文字のストリームの中の位置（バイト）
    compressed: bool
    pcd_pos: int  # CLX の中の、このピースの PCD の位置


@dataclass
class CPRef:
    """段落の文字の1片の、文書の中の位置（.docx の要素と同じように .s で並べられる）。"""

    s: int  # CP


@dataclass
class Run:
    fc0: int
    fc1: int
    grpprl: bytes


def sprms(grpprl: bytes):
    """sprm（書式の命令）を順に：(sprm, 値のバイト列)。"""
    for sprm, at, size in sprm_spans(grpprl):
        yield sprm, grpprl[at:at + size]


def sprm_spans(grpprl: bytes):
    """sprm を順に：(sprm, 値の始まりの位置, 値の大きさ)。値をその場で書き換えるときに使う。"""
    i = 0
    n = len(grpprl)
    while i + 2 <= n:
        sprm = struct.unpack_from("<H", grpprl, i)[0]
        spra = sprm >> 13
        i += 2
        if spra in (0, 1):
            size = 1
        elif spra in (2, 4, 5):
            size = 2
        elif spra == 3:
            size = 4
        elif spra == 7:
            size = 3
        else:
            if i >= n:
                break
            if sprm == 0xD608:  # sprmTDefTable：長さは2バイト
                size = struct.unpack_from("<H", grpprl, i)[0] - 1
                i += 2
            elif sprm == 0xC615:  # sprmPChgTabs
                size = grpprl[i]
                i += 1
                if size == 255:  # 長さは、消す数・足す数から数える（MS-DOC 2.9.182）
                    dels = grpprl[i] if i < n else 0
                    adds = grpprl[i + 1 + 4 * dels] if i + 1 + 4 * dels < n else 0
                    size = 2 + 4 * dels + 3 * adds
            else:
                size = grpprl[i]
                i += 1
        yield sprm, i, size
        i += size


TLC = {1: "dot", 2: "hyphen", 3: "underscore", 4: "heavy", 5: "middleDot"}  # MS-DOC 2.9.301 TabLC（0・7 は線なし）


def lay_tabs(stops: dict, op: bytes, close: bool) -> None:
    """タブ位置を重ねる（位置 twip → TBD）：sprmPChgTabsPapx（close=False：PChgTabsPapxOperand）・sprmPChgTabs
    （close=True：PChgTabsOperand）の値（cb の後ろ）。消す位置（25 twip 以内、close のときは rgdxaClose の幅）を消してから、足す。
    途中で終わる値は、重ねない。"""
    try:
        n = op[0]
        dels = struct.unpack_from(f"<{n}h", op, 1)
        i = 1 + 2 * n
        width = (25,) * n
        if close:
            width = tuple(max(25, w) for w in struct.unpack_from(f"<{n}h", op, i))
            i += 2 * n
        m = op[i]
        adds = struct.unpack_from(f"<{m}h", op, i + 1)
        tbds = op[i + 1 + 2 * m:i + 1 + 3 * m]
        if len(tbds) != m:
            return
    except (struct.error, IndexError):
        return
    for d, w in zip(dels, width):
        for x in [x for x in stops if abs(x - d) <= w]:
            del stops[x]
    for x, t in zip(adds, tbds):
        stops[x] = t


def tab_leaders(stops: dict, top: int = 5) -> tuple:
    """タブ位置のうち、線（リーダー）を引くものの、線の種類（TBD：jc 3 ビット・tlc 3 ビット。jc が 4（縦線）の位置は線なし。MS-DOC 2.9.310）。
    top：その形式の仕様書にある tlc のいちばん大きい値（Word 6.0/95 は 4）。"""
    return tuple(sorted({TLC[t >> 3 & 7] for t in stops.values() if t & 7 != 4 and 1 <= t >> 3 & 7 <= top}))


ICO = {1: "000000", 2: "0000FF", 3: "00FFFF", 4: "00FF00", 5: "FF00FF", 6: "FF0000", 7: "FFFF00", 8: "FFFFFF", 9: "000080", 10: "008080",
       11: "008000", 12: "800080", 13: "800000", 14: "808000", 15: "808080", 16: "C0C0C0"}  # MS-DOC 2.9.119 Ico


def apply_look(look: tuple, ops) -> tuple:
    """取り消し線（sprmCFStrike 0x0837・sprmCFDStrike 0x2A53）と字の色（sprmCIco 0x2A42・sprmCCv 0x6870）を足す。
    取り消し線の値：0・1 はその値、0x80 はスタイルのまま、0x81 はスタイルの逆（MS-DOC 2.9.32 Bool8）。"""
    strike, color = look
    for sp, op in ops:
        if sp in (0x0837, 0x2A53) and op:
            strike = {0: False, 1: True, 0x80: strike, 0x81: not strike}.get(op[0], strike)
        elif sp == 0x2A42 and op:
            color = ICO.get(op[0], "000000") if op[0] else "000000"  # 0 は自動
        elif sp == 0x6870 and len(op) >= 4:
            color = "000000" if op[3] == 0xFF else f"{op[0]:02X}{op[1]:02X}{op[2]:02X}"  # 4 バイト目 0xFF は自動（MS-DOC 2.9.40 COLORREF）
    return strike, color


def _merge_style(under: tuple, over: tuple) -> tuple:
    """段落のスタイルの上に字のスタイル：字のスタイルが決めたものが勝つ（色のない字のスタイルは、段落のスタイルの色のまま）。"""
    return (over[0] or under[0], over[1] if over[1] is not None else under[1])


class DocFile:
    kind = "doc"

    def __new__(cls, blob: bytes = b"", filename: str = "", cfb=None, partial=None):
        # Word 6.0/95（FIB の版 101〜104）の文書は、その形の読み方（word6.py）で読む
        if cls is DocFile and cfb is None:
            from .word6 import Word6File, is_word6

            if is_word6(blob):
                return super().__new__(Word6File)
        return super().__new__(cls)

    def __init__(self, blob: bytes, filename: str = "", cfb=None, partial: dict | None = None):
        """cfb：ほかの文書に埋め込んだ文書（cfb.SubStorage）を読むとき。blob は外側のファイル。
        partial：字の範囲だけを読むとき（lazy）：{"range": (種類, 始め, 終わり), "embeds": 前までの控え, "pictures": 前までの絵の数}。"""
        self.blob = blob
        self.filename = filename
        self.cfb = cfb if cfb is not None else CFB(blob)
        self.partial = partial
        p_ = partial or {}
        self.embeds: list[dict] = [dict(e) for e in p_.get("embeds", ())]  # 埋め込んだ OLE の部品（ObjectPool の入れ物）
        self.pictures: list[dict] = [{"n": i + 1, "stub": True} for i in range(p_.get("pictures", 0))]  # 絵：説明（代わりの字）と飾りの印。段階とは別の注意に使う（docx と同じ形）
        if not self.cfb.has("WordDocument"):
            raise DocError("Word の文書（WordDocument）がありません")
        self.wd = self.cfb.read("WordDocument")
        wd = self.wd
        ident, self.nfib = struct.unpack_from("<HH", wd, 0)
        if ident == 0xA5DC or (ident == 0xA5EC and self.nfib < 105):
            # Word 6.0/95 の版 101〜104 は word6.py で読む。それより古い版（Word 2.0 など）は、形を確かめていない
            raise DocError(f"Word 6.0 より古い Word の文書です（FIB の版 {self.nfib}）。この形は確かめていないので、読み違えた字を"
                           "出さないよう、読むのをやめました")
        if ident != 0xA5EC:
            raise DocError("Word 97〜2003 の文書ではありません")
        self.flags = struct.unpack_from("<H", wd, 0x0A)[0]
        if self.flags & 0x0100:
            raise DocError("パスワードで暗号化された Word の文書です")
        self.table_name = "1Table" if self.flags & 0x0200 else "0Table"
        self.table = self.cfb.read(self.table_name) if self.cfb.has(self.table_name) else b""
        self.data = self.cfb.read("Data") if self.cfb.has("Data") else b""
        csw = struct.unpack_from("<H", wd, 32)[0]
        pos = 34 + csw * 2
        self.cslw = struct.unpack_from("<H", wd, pos)[0]
        self.lw_pos = pos + 2
        self.lw = list(struct.unpack_from(f"<{self.cslw}i", wd, pos + 2))
        pos2 = pos + 2 + self.cslw * 4
        self.cb_fclcb = struct.unpack_from("<H", wd, pos2)[0]
        self.fclcb_pos = pos2 + 2
        vals = struct.unpack_from(f"<{self.cb_fclcb * 2}I", wd, pos2 + 2)
        self.fc: dict[str, tuple[int, int]] = {}
        for k, name in enumerate(FC_NAMES):
            if 2 * k + 1 < len(vals):
                self.fc[name] = (vals[2 * k], vals[2 * k + 1])
        # 文字の数（本文・脚注・ヘッダー・コメント・文末脚注・テキストボックス・ヘッダーのテキストボックス）
        self.ccp = {"text": self.lw[3], "ftn": self.lw[4], "hdd": self.lw[5], "atn": self.lw[7], "edn": self.lw[8],
                    "txbx": self.lw[9], "hdrtxbx": self.lw[10]}
        sk = p_.get("skel")  # 読んである字と表を使う（lazy：字の範囲ごとに読み直さない）
        self.pieces = sk.pieces if sk else self._pieces()
        self.papx = sk.papx if sk else self._fkps("PlcfBtePapx", papx=True)
        self.chpx = sk.chpx if sk else self._fkps("PlcfBteChpx", papx=False)
        self.text = sk.text if sk else (self._text() if not p_.get("no_text") else "")
        self.protection, self.protection_enforced = self._dop()
        self.has_macros = self.cfb.has("Macros/VBA/dir") or any(p.lower().startswith("macros/") for p in self.cfb.streams())
        self.controls: list[Control] = []
        self.bookmarks = sk.bookmarks if sk else self._bookmarks()
        self.stories: list[tuple[str, str, list]] = []
        self._build()

    # ------------------------------------------------------------------ FIB・ピース

    @property
    def vba(self):
        """マクロ（VBA）のプロジェクト：モジュールとソース（todoku.vba.Project）。マクロがなければ None。"""
        if getattr(self, "_vba", None) is None and self.has_macros:
            from ..vba import read_project

            try:
                self._vba = read_project(self._vba_blob())
            except Exception as e:  # noqa: BLE001
                self._vba = e
        v = getattr(self, "_vba", None)
        return v if v is not None and not isinstance(v, Exception) else None

    def _vba_blob(self) -> bytes:
        return self.blob  # Macros/VBA/dir と Macros/PROJECT を、そのまま読む

    def _pieces(self) -> list[Piece]:
        fc, lcb = self.fc["Clx"]
        t = self.table
        i = fc
        while i < min(fc + lcb, len(t)):  # 表の位置と長さの宣言は、ストリームの中まで
            if t[i] == 1:
                cbg = struct.unpack_from("<H", t, i + 1)[0]
                i += 3 + cbg
            elif t[i] == 2:
                plc_lcb = struct.unpack_from("<I", t, i + 1)[0]
                n = (plc_lcb - 4) // 12
                base = i + 5
                cps = struct.unpack_from(f"<{n + 1}I", t, base)
                out = []
                for k in range(n):
                    pp = base + (n + 1) * 4 + k * 8
                    raw = struct.unpack_from("<I", t, pp + 2)[0]
                    comp = bool(raw & 0x40000000)
                    fcv = raw & 0x3FFFFFFF
                    out.append(Piece(cps[k], cps[k + 1], fcv // 2 if comp else fcv, comp, pp))
                self.clx_pcdt = i
                return out
            else:
                break
        raise DocError("ピースの表（CLX）が見つかりません")

    def _text(self) -> str:
        out = []
        for p in self.pieces:
            n = p.cp1 - p.cp0
            if p.compressed:
                out.append(self.wd[p.fc:p.fc + n].decode("cp1252", errors="replace"))
            else:
                out.append(self.wd[p.fc:p.fc + 2 * n].decode("utf-16-le", errors="surrogatepass"))
        return cp_string("".join(out))

    def fc_of(self, cp: int) -> int:
        p = self.piece_of(cp)
        return p.fc + (cp - p.cp0) * (1 if p.compressed else 2) if p is not None else -1

    def piece_of(self, cp: int) -> Piece | None:
        got = self.__dict__.get("_piece_index")
        if got is None or got[0] is not self.pieces:  # 字ごとに全部のピースをなめない（速い保存のピースの多い文書で 2 乗になる）
            starts = [p.cp0 for p in self.pieces]
            got = self.__dict__["_piece_index"] = (self.pieces, starts if starts == sorted(starts) else None)
        starts = got[1]
        if starts is None:  # 順に並んでいない（壊れた）ピースの表
            return next((p for p in self.pieces if p.cp0 <= cp < p.cp1), None)
        i = bisect.bisect_right(starts, cp) - 1
        while i >= 0:  # 同じ位置から始まる空のピースは、前のものを見る
            p = self.pieces[i]
            if p.cp0 <= cp < p.cp1:
                return p
            if p.cp1 > p.cp0:
                return None
            i -= 1
        return None

    # ------------------------------------------------------------------ 絵の説明（代わりの字）
    def _picture_at(self, cp: int, o: int, pic: int | None) -> dict | None:
        """本文の絵の印（0x01：行の中の絵、0x08：浮いた図形）の、説明（MS-ODRAW の wzDescription 0x0381）。
        絵でない図形（テキストボックスなど、BLIP を持たないもの）は None。"""
        if o == 0x01:
            if pic is None or pic + 6 > len(self.data):
                return None
            lcb, cb_header = struct.unpack_from("<IH", self.data, pic)
            mm = struct.unpack_from("<H", self.data, pic + 6)[0] if pic + 8 <= len(self.data) else 0
            start = pic + cb_header
            if mm == 0x66 and start < len(self.data):  # MM_SHAPEFILE：絵の名前が先に来る（MS-DOC 2.9.192）
                start += 1 + self.data[start]
            props = _art_props(self.data, start, min(pic + lcb, len(self.data)))
        else:
            props = self._floating_props().get(self._spid_at(cp))
            if props is None:
                return None
        if 0x0104 not in props and 0x0186 not in props and not props.get("blip"):  # pib・fillBlip：絵を持つ図形だけ
            return None
        alt = props.get(0x0381) or ""
        return {"alt": alt.strip() if isinstance(alt, str) else "", "decorative": False}

    def _spid_at(self, cp: int) -> int | None:
        """浮いた図形の印の CP → 図形の番号（PlcSpaMom の FSPA の spid）。"""
        fc, lcb = self.fc.get("PlcSpaMom", (0, 0))
        if not lcb:
            return None
        n = (lcb - 4) // 30
        cps = struct.unpack_from(f"<{n + 1}i", self.table, fc)
        for k in range(n):
            if cps[k] == cp:
                return struct.unpack_from("<I", self.table, fc + (n + 1) * 4 + 26 * k)[0]
        return None

    def _floating_props(self) -> dict:
        """浮いた図形の性質（図形の番号 → 性質）。DggInfo の OfficeArt の入れ物を一度だけ読む。"""
        got = self.__dict__.get("_spprops")
        if got is None:
            fc, lcb = self.fc.get("DggInfo", (0, 0))
            got = _art_shapes(self.table, fc, fc + lcb) if lcb else {}
            self.__dict__["_spprops"] = got
        return got

    # ------------------------------------------------------------------ 書式（FKP）
    def _fkps(self, name: str, papx: bool) -> list[Run]:
        fc, lcb = self.fc[name]
        if not lcb:
            return []
        n = (lcb - 4) // 8
        t = self.table
        pns = struct.unpack_from(f"<{n}I", t, fc + (n + 1) * 4)
        out = []
        for pn in pns:
            page = self.wd[pn * 512:(pn + 1) * 512]
            if len(page) < 512:
                continue
            crun = page[511]
            rgfc = struct.unpack_from(f"<{crun + 1}I", page, 0)
            for k in range(crun):
                if papx:
                    boff = page[4 * (crun + 1) + 13 * k] * 2
                else:
                    boff = page[4 * (crun + 1) + k] * 2
                grpprl = b""
                if boff:
                    cb = page[boff]
                    if papx:
                        if cb == 0:
                            cb2 = page[boff + 1]
                            grpprl = page[boff + 2:boff + 2 + 2 * cb2]
                        else:
                            grpprl = page[boff + 1:boff + 1 + 2 * cb - 1]
                        if len(grpprl) >= 2:
                            self.__dict__.setdefault("papx_istd", {})[rgfc[k]] = struct.unpack_from("<H", grpprl, 0)[0]
                        grpprl = grpprl[2:]  # istd を除く
                    else:
                        grpprl = page[boff + 1:boff + 1 + cb]
                out.append(Run(rgfc[k], rgfc[k + 1], grpprl))
        out.sort(key=lambda r: r.fc0)
        return out

    def _lookup(self, runs: list[Run], fc: int) -> bytes:
        lo, hi = 0, len(runs) - 1
        while lo <= hi:
            mid = (lo + hi) // 2
            r = runs[mid]
            if fc < r.fc0:
                hi = mid - 1
            elif fc >= r.fc1:
                lo = mid + 1
            else:
                return r.grpprl
        return b""

    def _list_of(self, cp: int, starts: list) -> tuple[int, int]:
        """段落の番号のリスト：(ilfo, ilvl)。段落の書式になければ、スタイルの書式（基のスタイルも）。"""
        fc = self.fc_of(cp)
        ilfo = ilvl = None
        for sp, op in sprms(self._lookup(self.papx, fc)):
            if sp == 0x460B and len(op) >= 2:
                ilfo = struct.unpack_from("<h", op, 0)[0]
            elif sp == 0x260A and op:
                ilvl = op[0]
        if ilfo is None or ilvl is None:
            from .wrap import _styles

            i = bisect.bisect_right(starts, fc) - 1
            istd = self.__dict__.get("papx_istd", {}).get(starts[i]) if i >= 0 else None
            st = _styles(self)
            seen = 0
            while istd in st and seen < 20:
                base, papx, _chpx = st[istd]
                for sp, op in sprms(papx):
                    if sp == 0x460B and len(op) >= 2 and ilfo is None:
                        ilfo = struct.unpack_from("<h", op, 0)[0]
                    elif sp == 0x260A and op and ilvl is None:
                        ilvl = op[0]
                if base == 0x0FFF or base == istd:
                    break
                istd = base
                seen += 1
        return ilfo or 0, ilvl or 0

    def number_labels(self) -> dict:
        """段落の終わりの CP → 段落番号の字。文書の頭から順に数える（本文の一部だけを読むときも、番号が同じになるように）。"""
        got = self.__dict__.get("_number_labels")
        if got is not None:
            return got
        out: dict = {}
        self.__dict__["_number_labels"] = out
        if getattr(self, "word6", False):
            return out  # Word 6.0/95 の段落番号は別の形（読まない）
        from .numbering import Numberer, read_lists

        try:
            lists, lfos = read_lists(self.fc, self.table)
        except (struct.error, ValueError, IndexError) as e:
            self.__dict__["numbering_problem"] = f"段落番号の表を読めないので、段落番号の字を出していません（{e}）"
            return out
        if not lfos:
            return out
        n = Numberer(lists, lfos, lambda ftc: (tuple(self.font_name(f) or "" if f is not None else None for f in ftc),
                                              self._declared(ftc)))
        starts = sorted(self.__dict__.get("papx_istd", {}))
        for m in re.finditer("[\r\x07]", self.text):
            cp = m.start()
            ilfo, ilvl = self._list_of(cp, starts)
            if ilfo > 0:
                pp = self.para_props(cp)
                if not (pp["ttp"] or pp["inner_ttp"]):
                    out[cp] = n.next(ilfo, ilvl)
        return out

    TLC_TOP = 5  # 仕様書にある線の種類（tlc）のいちばん大きい値

    def _lay_grpprl_tabs(self, stops: dict, grpprl: bytes) -> None:
        for sp, at, size in sprm_spans(grpprl):
            if sp == 0xC60D:
                lay_tabs(stops, grpprl[at:at + size], False)
            elif sp == 0xC615 and grpprl[at - 1] != 255:  # cb が 255 のものは、Word が読まない（MS-DOC 2.9.182 の注）
                lay_tabs(stops, grpprl[at:at + size], True)

    def _style_table(self) -> dict:
        """istd → (基のスタイル, 段落の書式の grpprl, 字の書式の grpprl)。"""
        from .wrap import _styles

        return _styles(self)

    def _papx_at(self, fc: int) -> tuple:
        """ファイルの位置 fc の字の段落の、(段落の書式の grpprl, スタイル istd)。"""
        starts = self.__dict__.get("_papx_starts")
        if starts is None:
            starts = self.__dict__["_papx_starts"] = sorted(self.__dict__.get("papx_istd", {}))
        i = bisect.bisect_right(starts, fc) - 1
        return self._lookup(self.papx, fc), (self.__dict__.get("papx_istd", {}).get(starts[i], 0) if i >= 0 else 0)

    def _style_tabs(self, istd, seen: int = 0) -> dict:
        """スタイルのタブ位置（基のスタイルから順に重ねる）。"""
        memo = self.__dict__.setdefault("_style_tabs_memo", {})
        if istd in memo:
            return memo[istd]
        st = self._style_table()
        if istd not in st or seen > 20:
            return {}
        base, papx, _chpx = st[istd]
        stops = dict({} if base == 0x0FFF or base == istd else self._style_tabs(base, seen + 1))
        self._lay_grpprl_tabs(stops, papx or b"")
        memo[istd] = stops
        return stops

    def leaders_of(self, cp_end: int) -> tuple:
        """段落のタブ位置（スタイルの基から順に重ね、段落の書式を重ねる）のうち、線（リーダー）を引くものの、線の種類
        （dot・hyphen・underscore・heavy・middleDot：.docx の w:leader と同じ名前）。"""
        grp, istd = self._papx_at(self.fc_of(cp_end))
        kept = self.__dict__.setdefault("_leaders_memo", {})
        got = kept.get((grp, istd))
        if got is None:
            stops = dict(self._style_tabs(istd))
            self._lay_grpprl_tabs(stops, grp)
            kept[(grp, istd)] = got = tab_leaders(stops, self.TLC_TOP)
        return got

    def para_props(self, cp_end: int) -> dict:
        """段落の終わりの字（0x0D・0x07）の位置から、段落の書式（表の中か・行の終わりか・入れ子の深さ）。"""
        fc = self.fc_of(cp_end)
        d = {"in_table": False, "ttp": False, "itap": 0, "inner_cell": False, "inner_ttp": False}
        for sp, op in sprms(self._lookup(self.papx, fc)):
            if sp == 0x2416:
                d["in_table"] = bool(op[0]) if op else False
            elif sp == 0x2417:
                d["ttp"] = bool(op[0]) if op else False
            elif sp == 0x6649 and len(op) >= 4:
                d["itap"] = struct.unpack_from("<i", op, 0)[0]
            elif sp == 0x244B:
                d["inner_cell"] = bool(op[0]) if op else False
            elif sp == 0x244C:
                d["inner_ttp"] = bool(op[0]) if op else False
        if d["in_table"] and d["itap"] == 0:
            d["itap"] = 1
        return d

    def _add_embed(self, storage: str, prog: str, p) -> dict | None:
        """埋め込んだ OLE の部品を、文書の埋め込みの一覧に足す（入れ物がなければ足さない）。"""
        if not any(n.startswith(storage + "/") for n in self.cfb.streams()):
            return None
        for e in self.embeds:
            if e["storage"] == storage:
                return e
        from ..cfb import SubStorage
        from ..embeds.ole import label_of, stored_prog

        # 部品の種類は、入れ物が自分で名乗る ProgID（\x01CompObj）で決める。EMBED の命令の語は、Word が書いた名前で、
        # ProgID でないことがある（ギリシア語の Word の「Γράφημα」など）
        prog = stored_prog(SubStorage(self.cfb, storage)) or prog
        label = label_of(prog)
        e = {"n": len(self.embeds) + 1, "storage": storage, "prog": prog, "label": label, "para": p}
        self.embeds.append(e)
        return e

    def _add_data_embed(self, pos: int, prog: str, p) -> dict | None:
        """RC4 CryptoAPI で暗号化した文書では、OLE の部品の入れ物は ObjectPool ではなく Data ストリームの pos にある
        （FOBJH の後ろ。MS-DOC 2.9.94・sprmCPicLocation）。そこに入れ物があれば、埋め込みの一覧に足す。"""
        from ..cfb import CFB
        from ..embeds.ole import label_of, stored_prog
        from ..embeds.word import fobjh_storage

        body = fobjh_storage(self.data, pos)
        if body is None:
            return None
        storage = f"Data@{pos}"
        for e in self.embeds:
            if e["storage"] == storage:
                return e
        try:
            prog = stored_prog(CFB(body)) or prog
        except Exception:  # noqa: BLE001  読めない入れ物：命令の語のまま
            pass
        e = {"n": len(self.embeds) + 1, "storage": storage, "prog": prog, "label": label_of(prog), "para": p}
        self.embeds.append(e)
        return e

    def embedded_books(self) -> list[tuple[dict, object]]:
        """埋め込んだ Excel のブック：(埋め込みの情報, ブック)。開けないものは除く。"""
        out = []
        for e in self.embeds:
            wb = self.embedded_book(e)
            if wb is not None:
                out.append((e, wb))
        return out

    def embedded_book(self, e: dict):
        """埋め込み1つの Excel のブック（Excel のブックでない・開けないときは None）。"""
        from ..cfb import SubStorage
        from ..xls.book import XlsWorkbook

        sub = SubStorage(self.cfb, e["storage"])
        names = {n.lower() for n in sub.streams()}
        try:
            if names & {"workbook", "book"}:
                return XlsWorkbook(self.blob, e["storage"], cfb=sub)
            if "package" in names:
                from ..xlsx.workbook import Workbook

                return Workbook(sub.read("Package"), e["storage"])
        except Exception:  # noqa: BLE001
            return None
        return None

    def char_props(self, cp: int) -> dict:
        fc = self.fc_of(cp)
        grp = self._lookup(self.chpx, fc)
        kept = self.__dict__.setdefault("_chpx_props", {})  # 書式の命令の並び → 読んだ結果（同じ書式の字が続くので、字ごとに読み直さない）
        d = kept.get(grp)
        if d is None:
            d = kept[grp] = self._chpx_props_of(grp)
        d = dict(d)
        d["look"] = self._char_look(fc, grp)
        return d

    @staticmethod
    def _chpx_props_of(grp: bytes) -> dict:
        d = {"hidden": False, "underline": False, "pic": None, "symbol": None, "spec": False, "data": False, "deleted": False,
             "inserted": False}
        rm = [None, 0, None, 0]  # 挿入した人の番号・日時、削除した人の番号・日時（sprmCIbstRMark・sprmCDttmRMark・…Del）
        for sp, op in sprms(grp):
            if sp in (0x4804, 0x4863) and len(op) >= 2:
                rm[0 if sp == 0x4804 else 2] = struct.unpack_from("<H", op, 0)[0]
            elif sp in (0x6805, 0x6864) and len(op) >= 4:
                rm[1 if sp == 0x6805 else 3] = struct.unpack_from("<I", op, 0)[0]
            # 命令の番号は MS-DOC の原本（2.6.1）で確かめた：0x083C 非表示（0x0839 は影付きで別のもの）、0x0800 変更履歴の削除、
            # 0x0801 変更履歴の挿入
            if sp == 0x083C:
                d["hidden"] = bool(op[0] & 1) if op else False
            elif sp == 0x0800:
                d["deleted"] = bool(op[0] & 1) if op else False
            elif sp == 0x0801:
                d["inserted"] = bool(op[0] & 1) if op else False
            elif sp == 0x2A3E:
                d["underline"] = bool(op[0]) if op else False
            elif sp == 0x6A03 and len(op) >= 4:
                d["pic"] = struct.unpack_from("<I", op, 0)[0]
            elif sp == 0x6A09 and len(op) >= 4:
                d["symbol"] = struct.unpack_from("<HH", op, 0)
            elif sp == 0x0855:
                d["spec"] = bool(op[0]) if op else False
            elif sp == 0x0806:
                d["data"] = bool(op[0]) if op else False
        if d["deleted"] or d["inserted"]:
            d["rm"] = tuple(rm)
        return d

    def rev_authors(self) -> list[str]:
        """変更履歴・コメントの作者の名前（SttbfRMark。MS-DOC 2.9.290：字が UTF-16 の STTB、余分のデータなし。
        初めの項目は、決まりで「Unknown」＝作者の分からない変更）。読めなければ空。"""
        got = self.__dict__.get("_rmark")
        if got is not None:
            return got
        out: list[str] = []
        self.__dict__["_rmark"] = out
        fc, lcb = self.fc.get("SttbfRMark", (0, 0))
        t = self.table
        if lcb < 6 or fc + lcb > len(t) or struct.unpack_from("<H", t, fc)[0] != 0xFFFF:
            return out
        n, extra = struct.unpack_from("<HH", t, fc + 2)
        i = fc + 6
        for _ in range(n):
            if i + 2 > fc + lcb:
                break
            cch = struct.unpack_from("<H", t, i)[0]
            out.append(t[i + 2:i + 2 + 2 * cch].decode("utf-16-le", errors="replace"))
            i += 2 + 2 * cch + extra
        return out

    def rev_of(self, props: dict) -> tuple | None:
        """字の書式から、変更履歴の事実（docx/revfacts.py の rev）：削除された字は削除の、ほかの挿入された字は挿入の、人と日時。
        人は SttbfRMark の番号（sprmCIbstRMark 0x4804・sprmCIbstRMarkDel 0x4863。0 は「分からない」）、日時は DTTM
        （sprmCDttmRMark 0x6805・sprmCDttmRMarkDel 0x6864）。"""
        if not (props.get("deleted") or props.get("inserted")):
            return None
        from ..docx.revfacts import dttm

        rm = props.get("rm") or (None, 0, None, 0)
        i, when = (rm[2], rm[3]) if props.get("deleted") else (rm[0], rm[1])
        names = self.rev_authors() if i else []
        return ("del" if props.get("deleted") else "ins", names[i] if i and i < len(names) else "", dttm(when))

    # ------------------------------------------------------------------ 書体の表・記号の書体
    def font_table(self) -> list[tuple[str, int]]:
        """書体の表（SttbfFfn。MS-DOC 2.9.286・2.9.83 FFN）：ftc の番号 → (名前, 文字セット chs)。読めなければ空。
        FFN：長さ（1 バイト。その後ろのバイト数）・印 1・太さ 2・chs 1・別名の位置 1・PANOSE 10・FONTSIGNATURE 24・名前（UTF-16）。"""
        got = self.__dict__.get("_ffn")
        if got is not None:
            return got
        out: list[tuple[str, int]] = []
        self.__dict__["_ffn"] = out
        fc, lcb = self.fc.get("SttbfFfn", (0, 0))
        t = self.table
        if lcb < 4 or fc + lcb > len(t):
            return out
        n = struct.unpack_from("<H", t, fc)[0]
        if n == 0xFFFF:
            return out  # 形の違う表（確かめていない）：書体の名前を決めない
        p, end = fc + 4, fc + lcb
        for _ in range(n):
            if p >= end or t[p] < 39 or p + 1 + t[p] > end:
                break
            cb = t[p]
            name = t[p + 40:p + 1 + cb - (cb + 1 - 40) % 2].decode("utf-16-le", "replace").split("\0")[0]
            out.append((name, t[p + 4]))
            p += 1 + cb
        return out

    def font_name(self, ftc: int) -> str | None:
        """書体の番号（ftc）の名前。書体の表にない番号は None。"""
        ft = self.font_table()
        return ft[ftc][0] if 0 <= ftc < len(ft) else None

    def _declared(self, ftcs) -> frozenset:
        """書体の番号 ftcs のうち、書体の表が「記号の書体」（chs が 2：SYMBOL_CHARSET）と宣言した書体の名前（symfont.key の形）。
        宣言は書体の番号ごとに見る：同じ名前の書体に、記号の書体の宣言と、そうでない宣言の 2 つがあることがある。"""
        ft = self.font_table()
        return frozenset(symfont.key(ft[f][0]) for f in ftcs if f is not None and 0 <= f < len(ft) and ft[f][1] == 2)

    def symbol_char(self, ftc: int, code: int) -> str:
        """記号の字（sprmCSymbol：書体 ftc の番号 code）の字。書体の対応表で読む（symfont）。対応表のない記号の書体・
        書体の表にない番号の書体の字は、私用の領域の字（U+F0xx。外字として伝える）。"""
        name = self.font_name(ftc)
        if name is None:
            return chr(0xF000 + code) if code <= 0xFF else chr(code)
        return symfont.decode(name, code, self._declared((ftc,))) or chr(0xF000 + (code & 0xFF))

    def _style_fonts(self, istd, seen: int = 0) -> tuple:
        """スタイル（基のスタイルも）の字の書式の書体の番号 (sprmCRgFtc0：ASCII の字, sprmCRgFtc2：そのほかの字)。決めていなければ None。"""
        memo = self.__dict__.setdefault("_style_fonts_memo", {})
        if istd in memo:
            return memo[istd]
        from .wrap import _styles

        st = _styles(self)
        if istd not in st or seen > 20:
            return (None, None)
        base, _papx, chpx = st[istd]
        got = list((None, None) if base == 0x0FFF or base == istd else self._style_fonts(base, seen + 1))
        for sp, op in sprms(chpx or b""):
            if sp in (0x4A4F, 0x4A51) and len(op) >= 2:
                got[0 if sp == 0x4A4F else 1] = struct.unpack_from("<H", op, 0)[0]
        memo[istd] = got = tuple(got)
        return got

    def char_fonts(self, cp: int) -> tuple | None:
        """位置 cp の字の書体が記号の書体なら (ASCII の字の書体の名前, そのほかの字の書体の名前, 記号の書体と宣言した名前)。
        どちらも記号の書体でなければ None。書体は、段落のスタイル → 字のスタイル（sprmCIstd）→ 字の書式の順で決める
        （文書の既定の書体は見ていない）。記号の書体の宣言（chs が 2）は、書体の番号ごとに見る：同じ名前の書体に、
        記号の書体の宣言と、そうでない宣言の 2 つがあることがある。"""
        fc = self.fc_of(cp)
        return self._fonts_at(fc, self._lookup(self.chpx, fc))

    def _fonts_at(self, fc: int, grp: bytes) -> tuple | None:
        starts = self.__dict__.get("_papx_starts")
        if starts is None:
            starts = self.__dict__["_papx_starts"] = sorted(self.__dict__.get("papx_istd", {}))
        i = bisect.bisect_right(starts, fc) - 1
        istd = self.__dict__.get("papx_istd", {}).get(starts[i]) if i >= 0 else 0
        kept = self.__dict__.setdefault("_char_fonts", {})
        got = kept.get((grp, istd))
        if got is None:
            ftc = list(self._style_fonts(istd))
            for sp, op in sprms(grp):
                if sp == 0x4A30 and len(op) >= 2:
                    ftc = [b if b is not None else a for a, b in zip(ftc, self._style_fonts(struct.unpack_from("<H", op, 0)[0]))]
                elif sp in (0x4A4F, 0x4A51) and len(op) >= 2:
                    ftc[0 if sp == 0x4A4F else 1] = struct.unpack_from("<H", op, 0)[0]
            names = [self.font_name(f) if f is not None else None for f in ftc]
            declared = self._declared(ftc)
            if any(symfont.is_symbol(n, declared) for n in names):
                got = (names[0], names[1], declared)
            else:
                got = False
            kept[(grp, istd)] = got
        return got or None

    def text_char(self, cp: int, ch: str) -> str:
        """本文の字のうち、記号の書体で書いた字（U+0021〜U+00FF と、私用の領域の U+F020〜U+F0FF：Word は、記号の書体で
        書いた字を私用の領域の番号で持つ）を、書体の対応表で読む（symfont）。記号の書体でない書体の字は、そのまま。
        対応表のない記号の書体の字は、私用の領域の字（外字として伝える）。"""
        if getattr(self, "word6", False) or not self.uses_symbol_fonts:
            return ch
        try:
            f = self.char_fonts(cp)
        except (struct.error, IndexError, KeyError):
            return ch
        return symfont.text(ch, f[0], f[1], f[2]) if f else ch

    @property
    def uses_symbol_fonts(self) -> bool:
        """書体の表に、記号の書体（対応表のある書体か、記号の書体と宣言した書体）があるか。なければ、字ごとに書体を引かない。"""
        got = self.__dict__.get("_uses_symbol_fonts")
        if got is None:
            got = self.__dict__["_uses_symbol_fonts"] = any(chs == 2 or symfont.known(n) for n, chs in self.font_table())
        return got

    def _style_look(self, istd, seen: int = 0) -> tuple:
        """スタイル（基のスタイルも）の字の書式の、取り消し線と字の色。"""
        memo = self.__dict__.setdefault("_style_looks", {})
        if istd in memo:
            return memo[istd]
        from .wrap import _styles

        st = _styles(self)
        if istd not in st or seen > 20:
            return (False, None)
        base, _papx, chpx = st[istd]
        got = (False, None) if base == 0x0FFF or base == istd else self._style_look(base, seen + 1)
        memo[istd] = got = apply_look(got, sprms(chpx or b""))
        return got

    def _char_look(self, fc: int, grp: bytes | None = None) -> tuple:
        """字の書式の事実（取り消し線か, 字の色）：段落のスタイル → 字のスタイル（sprmCIstd）→ 字の書式の順（MS-DOC 2.6.1）。
        grp：fc の字の書式の命令の並び（呼ぶ側がもう引いていれば）。"""
        if grp is None:
            grp = self._lookup(self.chpx, fc)
        starts = self.__dict__.get("_papx_starts")
        if starts is None:
            starts = self.__dict__["_papx_starts"] = sorted(self.__dict__.get("papx_istd", {}))
        i = bisect.bisect_right(starts, fc) - 1
        istd = self.__dict__.get("papx_istd", {}).get(starts[i]) if i >= 0 else 0
        kept = self.__dict__.setdefault("_char_looks", {})  # (命令の並び, 段落のスタイル) → 書式の事実
        got = kept.get((grp, istd))
        if got is not None:
            return got
        direct = self.__dict__.setdefault("_chpx_ops", {})
        ops = direct.get(grp)
        if ops is None:
            ops = direct[grp] = list(sprms(grp))
        look = self._style_look(istd)
        for sp, op in ops:
            if sp == 0x4A30 and len(op) >= 2:
                look = _merge_style(look, self._style_look(struct.unpack_from("<H", op, 0)[0]))
        kept[(grp, istd)] = got = apply_look(look, ops)
        return got

    # ------------------------------------------------------------------ 文書の保護
    def _dop(self) -> tuple[str, bool]:
        """文書の保護（DopBase、MS-DOC 2.7.2）：fProtEnabled（7バイト目の bit1）はフォームの欄だけ書ける、
        fLockRev（7バイト目の bit6）は変更履歴としてだけ、fLockAtn（6バイト目の bit4）はコメントだけ。"""
        fc, lcb = self.fc["Dop"]
        t = self.table
        if lcb < 8 or fc + 8 > len(t):
            return "", False
        if t[fc + 7] & 0x02:
            return "forms", True
        if t[fc + 7] & 0x40:
            return "trackedChanges", True
        if t[fc + 6] & 0x10:
            return "comments", True
        return "", False

    @property
    def tracking(self) -> bool:
        """「変更履歴の記録」がオンか（DopBase の fRevMarking：5バイト目の 0x80。tests/test_spec.py で仕様書と照らす。
        Word 6.0/95 の DOP でも同じ位置）。"""
        fc, lcb = self.fc["Dop"]
        t = self.table
        return lcb >= 8 and fc + 8 <= len(t) and bool(t[fc + 5] & 0x80)

    def marks(self) -> dict:
        """書き込んだ欄のしおり（_todoku_…）：名前 → (段落, 始まり, 終わり)。"""
        from ..docx.document import iter_paras
        from ..docx.window import MARK_PREFIX

        out = {}
        want = [(n, a, b) for n, a, b in self.bookmarks if n.startswith(MARK_PREFIX)]
        if not want:
            return out
        paras = [p for _k, _p, blocks in self.all_blocks() for p in iter_paras(blocks)]
        for name, a, b in want:
            for p in paras:
                if p.cp0 <= a <= p.cp1:
                    out[name[len(MARK_PREFIX):].replace("_", ".")] = (p, para_index(p, a), para_index(p, b))
                    break
        return out

    # ------------------------------------------------------------------ しおり
    def _bookmarks(self) -> list[tuple[str, int, int]]:
        fc, lcb = self.fc["SttbfBkmk"]
        if not lcb:
            return []
        t = self.table
        names = []
        i = fc
        ext = struct.unpack_from("<H", t, i)[0]
        if ext != 0xFFFF:
            return []
        cdata = struct.unpack_from("<H", t, i + 2)[0]
        cbextra = struct.unpack_from("<H", t, i + 4)[0]
        i += 6
        for _ in range(cdata):
            cch = struct.unpack_from("<H", t, i)[0]
            names.append(t[i + 2:i + 2 + 2 * cch].decode("utf-16-le", errors="replace"))
            i += 2 + 2 * cch + cbextra
        fcf, lcbf = self.fc["PlcfBkf"]
        fcl, lcbl = self.fc["PlcfBkl"]
        n = (lcbf - 4) // 8 if lcbf else 0
        starts = struct.unpack_from(f"<{n + 1}I", t, fcf) if n else ()
        ibkl = [struct.unpack_from("<H", t, fcf + (n + 1) * 4 + 4 * k)[0] for k in range(n)]
        m = (lcbl - 4) // 4 if lcbl else 0
        ends = struct.unpack_from(f"<{m + 1}I", t, fcl) if m else ()
        out = []
        for k in range(min(n, len(names))):
            e = ends[ibkl[k]] if ibkl[k] < len(ends) else starts[k]
            out.append((names[k], starts[k], e))
        return out

    # ------------------------------------------------------------------ 段落・表・欄
    def _build(self) -> None:
        c = self.ccp
        bounds = []
        at = 0
        for kind, key in (("body", "text"), ("footnotes", "ftn"), ("header", "hdd"), ("comments", "atn"),
                          ("endnotes", "edn"), ("textbox", "txbx"), ("header-textbox", "hdrtxbx")):
            n = c[key]
            if n > 0:
                bounds.append((kind, at, at + n))
            at += n
        self.body = []
        if self.partial is not None:
            kind, a, b = self.partial["range"]
            blocks = self._blocks(a, b) if b > a else []
            if kind == "body":
                self.body = blocks
            else:
                self.stories.append((kind, kind, blocks))
            return
        for kind, a, b in bounds:
            blocks = self._blocks(a, b)
            if kind == "body":
                self.body = blocks
            else:
                self.stories.append((kind, kind, blocks))

    def _blocks(self, a: int, b: int) -> list:
        text = self.text
        paras = []
        start = a
        b = min(b, len(text))  # 字の数の宣言（FIB の ccp）は、実際に読めた字までに切る（宣言を信じて空の段落を作らない）
        for cp in range(a, b):
            ch = text[cp]
            if ch in "\r\x07":
                paras.append((start, cp))
                start = cp + 1
        if start < b:
            paras.append((start, b - 1))
        out: list = []
        table: Table | None = None
        row: list[Cell] = []
        cell_paras: list[Para] = []
        for s, e in paras:
            pp = self.para_props(e)
            para = self._para(s, e)
            if pp["in_table"]:
                if table is None:
                    table = Table(None, "doc")
                    out.append(table)
                if pp["ttp"] or pp["inner_ttp"]:
                    # 行の終わり
                    if cell_paras:
                        row.append(self._cell(table, row, cell_paras))
                        cell_paras = []
                    table.rows.append(row)
                    row = []
                    continue
                cell_paras.append(para)
                if text_at(text, e) == "\x07":
                    row.append(self._cell(table, row, cell_paras))
                    cell_paras = []
                continue
            if table is not None:
                if cell_paras or row:
                    if cell_paras:
                        row.append(self._cell(table, row, cell_paras))
                    if row:
                        table.rows.append(row)
                table, row, cell_paras = None, [], []
            out.append(para)
        if table is not None and (row or cell_paras):
            if cell_paras:
                row.append(self._cell(table, row, cell_paras))
            table.rows.append(row)
        return out

    def _cell(self, table: Table, row: list, paras: list[Para]) -> Cell:
        c = Cell(None, len(table.rows), len(row), table=table)
        c.blocks = list(paras)
        for p in paras:
            p.cell = c
        return c

    def _fields_multi(self) -> list[tuple]:
        """段落・マスをまたぐ、閉じたフィールド（目次・索引など、結果が何段落にもなるもの）：
        (始まりの字, 区切りの字（なければ None）, 終わりの字, 命令の字)。始まりの順。"""
        memo = self.__dict__.get("_fields_multi_memo")
        if memo is not None:
            return memo
        text = self.text
        out, stack = [], []
        # フィールドの印は、同じ部分（本文・脚注・ヘッダー…）の中で組になる：部分の境で、開いたままのものを捨てる
        cuts, at = [], 0
        for key in ("text", "ftn", "hdd", "atn", "edn", "txbx", "hdrtxbx"):
            at += max(0, self.ccp.get(key, 0))
            cuts.append(at)
        k = 0
        for m in re.finditer("[\x13\x14\x15]", text):
            o, cp = ord(m[0]), m.start()
            while k < len(cuts) and cp >= cuts[k]:
                k += 1
                stack.clear()
            if o == 0x13:
                stack.append([cp, None])
            elif o == 0x14 and stack:
                if stack[-1][1] is None:
                    stack[-1][1] = cp
            elif o == 0x15 and stack:
                b, sep = stack.pop()
                if "\r" in text[b:cp] or "\x07" in text[b:cp]:
                    out.append((b, sep, cp, text[b + 1:sep if sep is not None else cp]))
        out.sort()
        self.__dict__["_fields_multi_memo"] = out
        return out

    def _locked_ends(self) -> frozenset:
        """「更新しない」に固定したフィールドの、終わりの印（0x15）の字の位置。フィールドの印の表（PlcfFld…：流れごと。位置は
        その流れの頭から）の、終わりの印の grffld の fLocked（0x10。MS-DOC 2.8.25 Fld・2.9.90 grffldEnd。Word 6.0/95 の
        FLD も同じ形）。表の印が、その位置の字と合わない所は読まない。"""
        memo = self.__dict__.get("_locked_ends_memo")
        if memo is not None:
            return memo
        out, base, text, t = set(), 0, self.text, self.table
        for key, story in (("PlcfFldMom", "text"), ("PlcfFldFtn", "ftn"), ("PlcfFldHdr", "hdd"), ("PlcfFldMcr", "mcr"),
                           ("PlcfFldAtn", "atn"), ("PlcfFldEdn", "edn"), ("PlcfFldTxbx", "txbx"), ("PlcffldHdrTxbx", "hdrtxbx")):
            at, size = self.fc.get(key, (0, 0))
            n = (size - 4) // 6 if size >= 10 else 0
            if n and at + size <= len(t):
                for k, cp in enumerate(struct.unpack_from(f"<{n}I", t, at)):
                    mark, flags = t[at + 4 * (n + 1) + 2 * k] & 0x1F, t[at + 4 * (n + 1) + 2 * k + 1]
                    if mark == 0x15 and flags & 0x10 and base + cp < len(text) and text[base + cp] == "\x15":
                        out.add(base + cp)
            base += max(0, self.ccp.get(story, 0))
        self.__dict__["_locked_ends_memo"] = memo = frozenset(out)
        return memo

    def _para(self, s: int, e: int) -> Para:
        p = Para(None, "doc", [])
        p.cp0, p.cp1 = s, e
        sk = (getattr(self, "partial", None) or {}).get("skel")
        locked = (sk if sk else self)._locked_ends()
        p.label = (sk if sk else self).number_labels().get(e, "")
        text = self.text
        if "\t" in text[s:e]:
            p.leaders = self.leaders_of(e)
        # 前の段落から続くフィールド（閉じると分かっているものだけ）：命令の途中なら命令の字として、結果の途中なら結果の字として読む
        multi = (sk if sk else self)._fields_multi()
        stack: list[dict] = [{"begin": b, "instr": ins, "phase": "result" if sep is not None and sep < s else "instr", "n0": 0,
                              "carried": True} for b, sep, end, ins in multi if b < s <= end] if multi else []
        cp = s
        # 変更履歴の数（承諾も却下もされていない挿入・削除）：印のついた字の、段落の中のひと続きを 1 つと数える
        revs = self.__dict__.setdefault("revisions", {"挿入": 0, "削除": 0})
        rev_prev = (False, False)
        while cp < e:
            ch = text[cp]
            o = ord(ch)
            if ch == CP_PAD:
                cp += 1
                continue
            if o == 0x13:
                stack.append({"begin": cp, "instr": "", "phase": "instr"})
                cp += 1
                continue
            if o == 0x14 and stack and stack[-1].get("carried"):
                stack[-1].update(phase="result", n0=len(p.segs))
                cp += 1
                continue
            if o == 0x14 and stack:
                f = stack[-1]
                f["phase"] = "result"
                f["sep"] = cp
                f["ctl"] = self._form_field(f, p)
                f["n0"] = len(p.segs)
                words = f["instr"].split()
                if words and words[0].upper() == "CONTROL":
                    # ActiveX の部品：入れ物の名前は、区切りの字の sprmCPicLocation の値の前に「_」（MS-DOC 2.1.4）
                    pic = self.char_props(cp)["pic"]
                    if pic is not None:
                        from ..axbin import doc_ref

                        f["ax"] = doc_ref(self, f"ObjectPool/_{pic}", words[1] if len(words) > 1 else "")
                if words and words[0].upper() == "EMBED":
                    # 埋め込んだ OLE の部品：入れ物の名前は、区切りの字の sprmCPicLocation の値の前に「_」（MS-DOC 2.1.4）
                    pic = self.char_props(cp)["pic"]
                    if pic is not None:
                        prog = words[1] if len(words) > 1 else ""
                        f["ole"] = self._add_embed(f"ObjectPool/_{pic}", prog, p) or self._add_data_embed(pic, prog, p)
                cp += 1
                continue
            if o == 0x15 and stack and stack[-1].get("carried"):
                f = stack.pop()
                if f["phase"] == "result" and not _is_form_instr(f["instr"]):
                    _mark_result(p.segs[f["n0"]:], f["instr"], None, "doc", p, cp in locked)
                cp += 1
                continue
            if o == 0x15 and stack:
                f = stack.pop()
                if "ctl" not in f:
                    f["ctl"] = self._form_field(f, p)
                    f["n0"] = len(p.segs)
                ctl = f.get("ctl")
                if ctl is None and len(p.segs) == f["n0"]:
                    self._eq_segs(f["instr"], p, cp)
                if ctl is None and f["phase"] == "result":
                    # 差し込み・計算などのフィールドの結果（MERGEFIELD・DATE・REF・= など）：Word が更新すると作り直す字
                    _mark_result(p.segs[f["n0"]:], f["instr"], None, "doc", p, cp in locked)
                if ctl is None:
                    from ..links import field_target

                    target = field_target(f["instr"])
                    if target:
                        self.__dict__.setdefault("links", []).append(("".join(s.text for s in p.segs[f["n0"]:]), target))
                if ctl is not None:
                    ctl.result_segs = p.segs[f["n0"]:]
                    ctl.end_cp = cp
                    for sg in ctl.result_segs:
                        sg.field = sg.field or ctl
                    if ctl.kind == "checkbox":
                        p.segs.append(Seg("☑" if ctl.checked else "☐", "field", CPRef(cp), None, field=ctl))
                    elif not ctl.result_segs:
                        p.segs.append(Seg("", "field", CPRef(cp), None, field=ctl))
                        ctl.result_segs = [p.segs[-1]]
                cp += 1
                continue
            if stack and stack[-1]["phase"] == "instr":
                stack[-1]["instr"] += ch
                cp += 1
                continue
            props = self.char_props(cp)
            hidden = props["hidden"] or props["deleted"]  # 変更履歴で削除された字は、Word と同じく本文に出さない
            rev = (bool(props.get("inserted")), props["deleted"])
            if rev != rev_prev:
                revs["挿入"] += 1 if rev[0] and not rev_prev[0] else 0
                revs["削除"] += 1 if rev[1] and not rev_prev[1] else 0
                rev_prev = rev
            # 変更履歴で削除された字（窓では、非表示の字と分けて〔変更履歴で削除された字〕と示す。.docx・RTF・OpenDocument と同じ）
            tkind = "del" if props["deleted"] else "t"
            trev = self.rev_of(props) if rev[0] or rev[1] else None
            if o == 0x09:
                p.segs.append(Seg("\t", "tab", CPRef(cp), None, hidden, props["underline"]))
            elif o in (0x0B, 0x0C, 0x0E):
                p.segs.append(Seg("\n", "br", CPRef(cp), None, hidden))
            elif o == 0x1E:
                p.segs.append(Seg("‑", "other", CPRef(cp), None, hidden))
            elif o == 0x1F:
                pass
            elif o in (0x01, 0x08) and props["spec"] and any(x.get("ax") is not None and not x.get("ax_done")
                                                              and x["phase"] == "result" for x in stack):
                # CONTROL のフィールドの結果の絵（控えの絵）の所に、ActiveX の部品を置く
                from ..axfile import _add_word

                fx = next(x for x in reversed(stack) if x.get("ax") is not None and x["phase"] == "result")
                fx["ax_done"] = True
                fx["ax"].pic_cp = cp
                _add_word(self, fx["ax"], None, "doc", p, None, hidden, len(p.segs), CPRef(cp))
            elif o in (0x01, 0x08) and props["spec"]:
                ole = next((x["ole"] for x in reversed(stack) if x.get("ole") and x["phase"] == "result"), None)
                if ole is not None and o == 0x01 and props["pic"] is not None:
                    ole.setdefault("pictures", []).append((cp, props["pic"]))  # Word が表示に使う表の絵（Data の PICF）
                pic = None if ole is not None else self._picture_at(cp, o, props["pic"])
                if pic is not None:
                    pic.update(n=len(self.pictures) + 1, para=p)
                    self.pictures.append(pic)
                    mark = f"〔図: {pic['alt']}〕" if pic["alt"] else f"〔図{pic['n']}（説明なし）〕"
                else:
                    mark = f"〔埋め込み{ole['n']}：{ole['label']}〕" if ole else "〔図〕"
                p.segs.append(Seg(mark, "other", CPRef(cp), None, hidden))
            elif o in (0x02, 0x05) and props["spec"]:
                p.segs.append(Seg("※", "note", CPRef(cp), None, hidden))
            elif o == 0x28 and props["symbol"] and props["spec"]:
                p.segs.append(Seg(self.symbol_char(*props["symbol"]), "del" if props["deleted"] else "sym", CPRef(cp), None, hidden,
                                  props["underline"], look=props.get("look"), rev=trev))
            elif o < 0x20 and o not in (0x09,):
                pass
            else:
                if 0x20 < o <= 0xFF or 0xF020 <= o <= 0xF0FF:
                    ch = self.text_char(cp, ch)  # 記号の書体で書いた字
                # 同じ書式の字はまとめる
                if p.segs and p.segs[-1].kind == tkind and p.segs[-1].hidden == hidden and \
                        p.segs[-1].underline == props["underline"] and p.segs[-1].node.s + len(p.segs[-1].text) == cp \
                        and p.segs[-1].field is None and not stack and p.segs[-1].look == props.get("look") \
                        and p.segs[-1].rev == trev:
                    p.segs[-1].text += ch.replace("\xa0", " ")
                else:
                    p.segs.append(Seg(ch.replace("\xa0", " "), tkind, CPRef(cp), None, hidden, props["underline"], look=props.get("look"),
                                      rev=trev))
            cp += 1
        if stack and multi:
            # 段落の終わりで開いているフィールドのうち、後ろで閉じるもの：ここまでの結果の字に、フィールドの結果の印
            closing = {b: end for b, _sep, end, _ins in multi}
            for f in reversed(stack):
                if f["phase"] == "result" and f.get("ctl") is None and f["begin"] in closing and not _is_form_instr(f["instr"]):
                    _mark_result(p.segs[f["n0"]:], f["instr"], None, "doc", p, closing[f["begin"]] in locked)
        return p

    def _eq_segs(self, instr: str, p: Para, cp: int) -> None:
        """結果の字のない EQ フィールド（ふりがななど）の見える字。"""
        from ..eqfield import eq_text

        got = eq_text(instr)
        if got is None:
            return
        base, ruby = got
        if ruby is None:
            p.segs.append(Seg(base, "other", CPRef(cp), None))
            return
        p.segs.append(Seg(base, "ruby", CPRef(cp), None))
        if ruby:
            p.segs += [Seg("（", "other", CPRef(cp), None), Seg(ruby, "ruby", CPRef(cp), None), Seg("）", "other", CPRef(cp), None)]

    def _form_field(self, f: dict, p: Para) -> Control | None:
        instr = f["instr"].strip()
        word = instr.split()[0].upper() if instr.split() else ""
        if word not in ("FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN"):
            return None
        kind = {"FORMTEXT": "text", "FORMCHECKBOX": "checkbox", "FORMDROPDOWN": "dropdown"}[word]
        ctl = Control(kind, "formfield", "", None, "doc", p)
        ctl.begin_cp = f["begin"]
        # FFData の位置：欄のデータの印（sprmCFData）の付いた字の、絵の位置（sprmCPicLocation）。Word も LibreOffice も、欄の命令の中の
        # 0x01 の字に置く。印のない字の絵の位置は、欄のデータではない（Word の書き出した .doc では、欄の始めの字 0x13 にも
        # 絵の位置 0 が付いていて、そこを読むと、関係のないバイトを欄の名前・状態として読む）。
        end = f["begin"] + 1 + len(f["instr"])
        props = {"pic": None}
        for cp in [f["begin"]] + [f["begin"] + 1 + i for i, ch in enumerate(f["instr"]) if ch == "\x01"] + \
                ([f["sep"]] if f.get("sep") is not None else [end]):
            got = self.char_props(cp)
            if got["data"] and got["pic"] is not None:
                props = got
                break
        if props["pic"] is not None:
            ff = parse_ffdata(self.data, props["pic"])
            if ff:
                ctl.name = ff.get("name", "")
                ctl.help = ff.get("help", "")
                ctl.maxlen = ff.get("cch") or None
                ctl.ffdata_pos = ff.get("_pos")
                ctl.ffdata = ff
                if kind == "checkbox":
                    ctl.checked = ff.get("res", 0) == 1 if ff.get("res", 25) != 25 else bool(ff.get("wDef", 0))
                if kind == "dropdown":
                    ctl.options = ff.get("entries", [])
                    idx = ff.get("res", 0)
                    if idx == 25:
                        idx = ff.get("wDef", 0)
                    ctl.value = ctl.options[idx] if 0 <= idx < len(ctl.options) else None
                ctl.disabled, ctl.calc_on_exit = bool(ff.get("prot")), bool(ff.get("recalc"))
                if kind == "text":
                    ctl.text_type = TEXT_TYPES[ff["type_txt"]] if ff.get("type_txt", 0) < len(TEXT_TYPES) else "regular"
                    ctl.text_format = ff.get("format", "")
                    if ctl.text_type == "calculated":
                        ctl.expression = ff.get("default", "")
                if ff.get("type_txt") == 2:
                    ctl.kind = "date"
                    ctl.date_format = ff.get("format", "")
        self.controls.append(ctl)
        return ctl

    # ------------------------------------------------------------------ .docx と同じ入口
    def all_blocks(self):
        yield "body", "doc", self.body
        for kind, part, blocks in self.stories:
            yield kind, part, blocks


CP_PAD = "\uffff"  # BMP の外の字（𠮷 など）の 2 つ目の CP の場所（字ではない。U+FFFF は文字に使えない番号）
_SURROGATE_PAIR = re.compile("[\ud800-\udbff][\udc00-\udfff]")
_LONE_SURROGATE = re.compile("[\ud800-\udfff]")
_OUTSIDE_BMP = re.compile("[\U00010000-\U0010ffff]")


def cp_string(s: str) -> str:
    """文書の位置（CP。UTF-16 の 1 単位が 1 つ）と、文字列の位置を同じにする：BMP の外の字は、その字と CP_PAD の 2 つにする。
    Python の文字列は 𠮷 を 1 字に数えるので、そのままでは、それより後ろの字の書式（非表示・削除・表のマスの印など）と、
    本文と脚注の境がずれる。ピースの境で分かれた組は 1 字に戻し、組にならないものは U+FFFD にする。"""
    s = _SURROGATE_PAIR.sub(lambda m: m[0].encode("utf-16-le", "surrogatepass").decode("utf-16-le"), s)
    s = _LONE_SURROGATE.sub("\ufffd", s)
    return _OUTSIDE_BMP.sub(lambda m: m[0] + CP_PAD, s)


def units(ch: str) -> int:
    """字の CP の数。"""
    return 2 if ord(ch) > 0xFFFF else 1


def para_index(p: Para, cp: int) -> int:
    """文書の位置（CP）→ 段落の文字（p.text）の中の位置。"""
    k = 0
    for sg in p.segs:
        if sg.hidden:
            continue
        c0 = sg.node.s if sg.node is not None else None
        if c0 is not None and sg.kind == "t" and c0 <= cp <= c0 + len(sg.text):
            return k + (cp - c0)
        if c0 is not None and cp <= c0:
            return k
        k += len(sg.text)
    return k


def text_at(text: str, i: int) -> str:
    return text[i] if 0 <= i < len(text) else ""


def _xstz(b: bytes, i: int) -> tuple[str, int]:
    cch = struct.unpack_from("<H", b, i)[0]
    s = b[i + 2:i + 2 + 2 * cch].decode("utf-16-le", errors="replace")
    return s, i + 2 + 2 * cch + 2


def parse_ffdata(data: bytes, pos: int) -> dict:
    """フォームの欄のデータ（FFData）。Data のストリームの pos から：lcb・cbHeader（68）・見出し・FFData。"""
    try:
        lcb, cbh = struct.unpack_from("<IH", data, pos)
        i = pos + cbh
        ver = struct.unpack_from("<I", data, i)[0]
        if ver != 0xFFFFFFFF:
            return {}
        bits = struct.unpack_from("<H", data, i + 4)[0]
        itype = bits & 0x3
        res = (bits >> 2) & 0x1F
        type_txt = (bits >> 11) & 0x7
        cch, hps = struct.unpack_from("<HH", data, i + 6)
        k = i + 10
        name, k = _xstz(data, k)
        d = {"_pos": i, "type": itype, "res": res, "type_txt": type_txt, "cch": cch, "name": name, "bits_pos": i + 4,
             "prot": bool(bits & 0x0200), "recalc": bool(bits & 0x4000)}  # FFDataBits の fProt・fRecalc（MS-DOC 2.9.78）
        if itype == 0:
            d["default"], k = _xstz(data, k)
        if itype in (1, 2):
            d["wDef"] = struct.unpack_from("<H", data, k)[0]
            k += 2
        d["format"], k = _xstz(data, k)
        d["help"], k = _xstz(data, k)
        d["stat"], k = _xstz(data, k)
        _entry, k = _xstz(data, k)
        _exit, k = _xstz(data, k)
        if itype == 2:
            ext = struct.unpack_from("<H", data, k)[0]
            if ext == 0xFFFF:
                n = struct.unpack_from("<H", data, k + 2)[0]
                cbx = struct.unpack_from("<H", data, k + 4)[0]
                j = k + 6
                ents = []
                for _ in range(n):
                    cc = struct.unpack_from("<H", data, j)[0]
                    ents.append(data[j + 2:j + 2 + 2 * cc].decode("utf-16-le", errors="replace"))
                    j += 2 + 2 * cc + cbx
                d["entries"] = ents
        del lcb
        return d
    except (struct.error, IndexError):
        return {}


def _art_props(data: bytes, a: int, b: int, out: dict | None = None) -> dict:
    """OfficeArt のレコード（a〜b）をたどり、図形の性質（OPT 0xF00B・第3の OPT 0xF122）を集める：性質の番号 → 値。
    複雑な値のうち、名前（0x0380）と説明（0x0381）は字（UTF-16）にする。BLIP の入れ物（0xF007・0xF018〜0xF117）があれば blip。"""
    out = {} if out is None else out
    pos = a
    while pos + 8 <= b:
        verinst, typ, ln = struct.unpack_from("<HHI", data, pos)
        body = pos + 8
        end = min(body + ln, b)
        if (verinst & 0x0F) == 0x0F:
            _art_props(data, body, end, out)
        elif typ in (0xF00B, 0xF122):
            n = verinst >> 4
            props, p = [], body
            for _ in range(n):
                if p + 6 > end:
                    break
                opid, op = struct.unpack_from("<HI", data, p)
                props.append((opid & 0x3FFF, bool(opid & 0x8000), op))
                p += 6
            for pid, complex_, op in props:
                if complex_:
                    raw = data[p:p + op]
                    p += op
                    if pid in (0x0380, 0x0381):
                        out[pid] = raw.decode("utf-16-le", errors="replace").split("\x00", 1)[0]
                else:
                    out[pid] = op
        elif typ == 0xF007 or 0xF018 <= typ <= 0xF117:
            out["blip"] = True
        pos = body + ln
    return out


def _art_shapes(data: bytes, a: int, b: int) -> dict:
    """DggInfo の中の図形の入れ物（0xF004）ごとに、図形の番号（FSP 0xF00A の spid）→ 性質。"""
    shapes: dict = {}

    def walk(x: int, y: int) -> None:
        pos = x
        while pos + 8 <= y:
            verinst, typ, ln = struct.unpack_from("<HHI", data, pos)
            body = pos + 8
            end = min(body + ln, y)
            if typ == 0xF004:
                sp = None
                q = body
                while q + 8 <= end:
                    _vi, t2, l2 = struct.unpack_from("<HHI", data, q)
                    if t2 == 0xF00A and q + 12 <= end:
                        sp = struct.unpack_from("<I", data, q + 8)[0]
                    q += 8 + l2
                if sp is not None:
                    shapes[sp] = _art_props(data, body, end)
            elif (verinst & 0x0F) == 0x0F:
                walk(body, end)
            pos = body + ln

    walk(a, b)
    return shapes
