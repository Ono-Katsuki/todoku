"""古い Word（.doc）に書き込む。

文字の置き換えは、ピースの表（CLX）を分けて、新しい文字をストリームの後ろに足す（Word の「高速保存」と同じ考え方）。
文字の数が変わるので、本文の位置（CP）を持つ表（段落の区切り・フィールド・しおり・図形の位置・脚注の位置など）を、
意味を確かめたものだけずらす。描き直せば作り直される表（スペルチェックや段落の高さの控え）は空にする。
意味の分からない位置の表が本文にあるときは、書き込まずに断る（壊れた文書を作らない）。

新しい文字の書式（CHPX）は、置き換える前の字の書式をそのまま引き継ぐ。書いた所は見えないしおり（_todoku_…）で囲む。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from ..cfb import CFB
from ..docx.fill import FillError, _as_bool, _is_text_value
from ..docx.window import MARK_PREFIX, Spots
from .book import FC_NAMES, DocFile

# 本文の位置（CP）を持つ表と、1つの要素のデータの大きさ（バイト）。どれも MS-DOC の原本（2.8 節の各 PLC と 2.9 節の構造）で確かめた。
#   Plcfbkf=FBKF 4・Plcfbkl=0・Plcfbkfd=FBKFD 6・Plcfbkld=FBKLD 4・Plcbkf=BKF 6・Plcbkl=BKL 4・Plcbkfd=BKFD 10・Plcbkld=BKLD 8
#   文書のすべての部分にまたがる位置を持つもの（スペル・文法・言語の判定・表の控え・クッキー・IME・自動要約・スマートタグ）も、同じ規則でずらす
SHIFT = {"PlcffndRef": 2, "PlcfendRef": 2, "PlcfandRef": 30, "PlcfSed": 12, "PlcfFldMom": 2, "PlcSpaMom": 26,
         "PlcfAtnBkf": 4, "PlcfAtnBkl": 0, "PlcfbkfBPRepairs": 4, "PlcfbklBPRepairs": 0, "PlcfBkfFcc": 6,
         "PlcfBklFcc": 4, "PlcfBkfFactoid": 6, "PlcfBklFactoid": 4, "PlcfBkfSdt": 10, "PlcfBklSdt": 8,
         "PlcfBkfProt": 6, "PlcfBklProt": 4, "PlcfWkb": 12,
         "PlcfSpl": 2, "PlcfGram": 2, "PlcfLad": 2, "PlcfTch": 4, "Plcfcookie": 10, "PlcfCookieOld": 16,
         "Plcfuim": 20, "PlcfAsumy": 4, "Plcffactoid": 2}
# 本文以外の流れ（脚注・ヘッダー・コメント・文末脚注・テキストボックス）の、流れの中の位置を持つ表
SUB = {"ftn": (4, [("PlcffndTxt", 0), ("PlcfFldFtn", 2)]),
       "hdd": (5, [("PlcfHdd", 0), ("PlcfFldHdr", 2), ("PlcSpaHdr", 26)]),
       "atn": (7, [("PlcfandTxt", 0), ("PlcfFldAtn", 2)]),
       "edn": (8, [("PlcfendTxt", 0), ("PlcfFldEdn", 2)]),
       "txbx": (9, [("PlcftxbxTxt", 22), ("PlcfTxbxBkd", 6), ("PlcfFldTxbx", 2)]),
       "hdrtxbx": (10, [("PlcfHdrtxbxTxt", 22), ("PlcfTxbxHdrBkd", 6), ("PlcffldHdrTxbx", 2)])}
ORDER = ["text", "ftn", "hdd", "atn", "edn", "txbx", "hdrtxbx"]
# 原本で「使わない控え（SHOULD NOT be emitted・SHOULD be ignored）」とされているもの：空にする
CACHES = {"PlcfPhe", "PlcfBteLvc", "PlcfLvcPre10", "PlcflvcOld", "PlcflvcOldInline", "PlcflvcNew", "PlcflvcNewInline",
          "PlcflvcOldXP", "PlcflvcNewXP", "PlcflvcMixedXP", "PlcfpmiOld", "PlcfpmiOldInline", "PlcfpmiNew",
          "PlcfpmiNewInline", "PlcfpmiOldXP", "PlcfpmiNewXP", "PlcfpmiMixedXP", "PgdMother", "BkdMother", "AfdMother",
          "PgdFtn", "BkdFtn", "AfdFtn", "PgdEdn", "BkdEdn", "AfdEdn", "Afd", "PgdMotherOld", "BkdMotherOld", "PgdFtnOld",
          "BkdFtnOld", "PgdEdnOld", "BkdEdnOld", "PgdMotherOldOld", "BkdMotherOldOld", "PgdFtnOldOld", "BkdFtnOldOld",
          "PgdEdnOldOld", "BkdEdnOldOld", "PlcfPgdFtn"}
# 本文の位置を持たない（または別の流れの位置を持ち、SUB でずらす）もの：そのままでよい
SAFE = {"StshfOrig", "Stshf", "PlcffndTxt", "PlcfandTxt", "PlcPad", "SttbfGlsy", "PlcfGlsy", "PlcfHdd", "PlcfBteChpx",
        "PlcfBtePapx", "PlcfSea", "SttbfFfn", "PlcfFldHdr", "PlcfFldFtn", "PlcfFldAtn", "PlcfFldMcr", "SttbfBkmk",
        "PlcfBkf", "PlcfBkl", "Cmds", "Unused1", "SttbfMcr", "PrDrvr", "PrEnvPort", "PrEnvLand", "Wss", "Dop",
        "SttbfAssoc", "Clx", "AutosaveSource", "GrpXstAtnOwners", "SttbfAtnBkmk", "Unused2", "Unused3", "PlcSpaHdr",
        "Pms", "FormFldSttbs", "PlcfendTxt", "PlcfFldEdn", "Unused4", "DggInfo", "SttbfRMark", "SttbfCaption",
        "SttbfAutoCaption", "PlcftxbxTxt", "PlcfFldTxbx", "PlcfHdrtxbxTxt", "PlcffldHdrTxbx", "StwUser", "SttbTtmbd",
        "CookieData", "SttbfIntlFld", "RouteSlip", "SttbSavedBy", "SttbFnm", "PlfLst", "PlfLfo", "PlcfTxbxBkd",
        "PlcfTxbxHdrBkd", "DocUndoWord9", "RgbUse", "Usp", "Uskf", "PlcupcRgbUse", "PlcupcUsp", "SttbGlsyStyle",
        "Plgosl", "Plcocx", "dwLowDateTime_dwHighDateTime", "SttbListNames", "SttbfUssr", "RmdThreading", "Mid",
        "SttbRgtplc", "MsoEnvelope", "RgDofr", "Plcosl", "Unused1_2002", "PlfguidUim", "AtrdExtra", "Plrsid",
        "DocUndo", "SttbfBkmkFcc", "SttbfbkmkBPRepairs", "PmsNew", "ODSO", "Unused2_2002", "Hplxsdr", "SttbfBkmkSdt",
        "CustomXForm", "SttbfBkmkProt", "SttbProtUser", "Unused_2003", "Plcfmthd", "SttbfBkmkMoveFrom",
        "PlcfBkfMoveFrom", "PlcfBklMoveFrom", "SttbfBkmkMoveTo", "PlcfBkfMoveTo", "PlcfBklMoveTo", "Unused1_2007",
        "Unused2_2007", "Unused3_2007", "SttbfBkmkArto", "PlcfBkfArto", "PlcfBklArto", "ArtoData", "Unused4_2007",
        "Unused5_2007", "Unused6_2007", "OssTheme", "ColorSchemeMapping",
        # 段落のまとまり（PGPArray）は位置を持たない（段落の書式の ipgpSelf から引く）。スマートタグの名前とデータも位置を持たない
        "PlcfPgp", "SttbfBkmkFactoid", "FactoidData"}


@dataclass
class Edit:
    a: int  # 本文の位置（CP）
    b: int
    text: str
    style_cp: int  # 書式を引き継ぐ字の位置
    mark: str | None = None


@dataclass
class DocFillResult:
    blob: bytes
    patches: list = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class DocWriter:
    def __init__(self, doc: DocFile):
        self.doc = doc
        self.edits: list[Edit] = []
        self.data_patches: list[tuple[int, bytes]] = []  # Data のストリームの、同じ大きさの書き換え

    def starts(self) -> dict[str, int]:
        out, at = {}, 0
        for k in ORDER:
            out[k] = at
            at += self.doc.ccp[k]
        return out

    def story_of(self, cp: int) -> str:
        st = self.starts()
        for k in reversed(ORDER):
            if self.doc.ccp[k] and st[k] <= cp:
                return k
        return "text"

    def add(self, e: Edit) -> None:
        total = sum(self.doc.ccp.values())
        if not (0 <= e.a <= e.b <= total):
            raise FillError("書き込む位置が文書の外です")
        if self.story_of(e.a) != self.story_of(max(e.a, e.b - 1)):
            raise FillError("書き込みが、本文とヘッダーなどの境をまたいでいます")
        if self.story_of(e.a) != "text":
            e.mark = None  # しおりは本文だけに置ける
        for o in self.edits:
            if e.a < o.b and o.a < e.b or (e.a == e.b == o.a == o.b):
                raise FillError("同じ所に2つの書き込みが重なりました")
        self.edits.append(e)

    # ------------------------------------------------------------------ 位置をずらす
    def shift(self, cp: int) -> int:
        """元の位置（CP）→ 書き込み後の位置。挿入（a == b）はその位置の字を後ろへ、置き換えは範囲の後ろを動かす。"""
        d = 0
        for e in self.edits:
            if e.a == e.b:
                if cp >= e.a:
                    d += len(e.text)
            elif cp >= e.b:
                d += len(e.text) - (e.b - e.a)
            elif e.a < cp < e.b:
                return e.a + d + min(cp - e.a, len(e.text))
        return cp + d

    def build(self) -> bytes:
        doc = self.doc
        if not self.edits and not self.data_patches:
            return doc.blob
        self.edits.sort(key=lambda e: (e.a, e.b))
        wd = bytearray(doc.wd)
        tb = bytearray(doc.table)
        fc = dict(doc.fc)
        problems = []
        # 位置の表を調べる：知らない表が本文にあれば断る
        for k, name in enumerate(FC_NAMES):
            if name not in fc or not fc[name][1]:
                continue
            if name in SHIFT or name in CACHES or name in SAFE:
                continue
            problems.append(name)
        if problems:
            raise FillError("この .doc には、書き込みで位置を正しくずらせるか確かめていない表があります（" + "、".join(problems)
                            + "）。壊れた文書を作らないために、書き込みません")
        # 1) 新しい文字をストリームの後ろに足す
        if len(wd) % 2:
            wd += b"\x00"
        texts_fc = []
        for e in self.edits:
            texts_fc.append(len(wd))
            wd += e.text.encode("utf-16-le")
        # 2) ピースの表を作り直す：元の文字の範囲と、足した文字の範囲を、文書の順に並べる
        old = doc.pieces
        t = doc.table

        def rest(pc) -> tuple[bytes, bytes]:
            return t[pc.pcd_pos:pc.pcd_pos + 2], t[pc.pcd_pos + 6:pc.pcd_pos + 8]

        def piece_at(cp: int):
            for pc in old:
                if pc.cp0 <= cp < pc.cp1:
                    return pc
            return old[-1]

        new_pieces: list[list] = []

        def emit_original(c0: int, c1: int) -> None:
            for pc in old:
                s0, e0 = max(c0, pc.cp0), min(c1, pc.cp1)
                if e0 > s0:
                    off = (s0 - pc.cp0) * (1 if pc.compressed else 2)
                    new_pieces.append([s0, e0, pc.fc + off, pc.compressed, *rest(pc)])

        cur = 0
        self.new_starts: list[int] = []
        new_cp = 0
        for k, e in enumerate(self.edits):
            emit_original(cur, e.a)
            new_cp += e.a - cur
            self.new_starts.append(new_cp)
            if e.text:
                pc = piece_at(e.a if e.a < old[-1].cp1 else e.a - 1)
                new_pieces.append([e.a, e.a + len(e.text), texts_fc[k], False, *rest(pc)])
                new_cp += len(e.text)
            cur = e.b
        emit_original(cur, old[-1].cp1)
        # 新しい CP を振り直す（元の CP → 新しい CP の順に並べ、長さから数え直す）
        out_pieces = []
        cp = 0
        for p0, p1, f, comp, flags, prm in new_pieces:
            n = p1 - p0
            out_pieces.append((cp, cp + n, f, comp, flags, prm))
            cp += n
        total_old = old[-1].cp1
        total_new = total_old + sum(len(e.text) - (e.b - e.a) for e in self.edits)
        if cp != total_new:
            raise FillError(f"ピースの表の組み立てが合いません（{cp} ≠ {total_new}）")
        # CLX：元の Prc（書式の命令）はそのまま、Pcdt を作り直して表のストリームの後ろへ
        clx_fc, clx_lcb = fc["Clx"]
        prc = bytes(t[clx_fc:doc.clx_pcdt])
        plc = bytearray()
        for p in out_pieces:
            plc += struct.pack("<I", p[0])
        plc += struct.pack("<I", out_pieces[-1][1])
        for p0, p1, f, comp, flags, prm in out_pieces:
            raw = (f * 2) | 0x40000000 if comp else f
            plc += flags + struct.pack("<I", raw) + prm
        clx = prc + b"\x02" + struct.pack("<I", len(plc)) + bytes(plc)
        fc["Clx"] = (len(tb), len(clx))
        tb += clx
        # 3) 位置の表をずらす（同じ場所で数を書き換える）
        for name, cb in SHIFT.items():
            f0, lcb = fc.get(name, (0, 0))
            if not lcb:
                continue
            n = (lcb - 4) // (4 + cb)
            if 4 * (n + 1) + cb * n != lcb:
                raise FillError(f"{name} の大きさが想定と違います")
            for k in range(n + 1):
                v = struct.unpack_from("<I", tb, f0 + 4 * k)[0]
                # 文書の終わりより後ろの位置（「長さ＋1」の終わりの印など）も、増えた分だけずらす
                struct.pack_into("<I", tb, f0 + 4 * k, self.shift(v))
        # 本文以外の流れの表：流れの中の位置 → 文書の中の位置にしてずらし、流れの中の位置に戻す
        st_old = self.starts()
        for story, (_idx, plcs) in SUB.items():
            if not any(self.story_of(e.a) == story for e in self.edits) and \
                    not any(e.a < st_old[story] for e in self.edits if e.text or e.b > e.a):
                continue
            base_old = st_old[story]
            base_new = self.shift(base_old) if not any(e.a == base_old and e.a == e.b for e in self.edits) else base_old + sum(
                len(e.text) - (e.b - e.a) for e in self.edits if e.a < base_old)
            for name, cb in plcs:
                f0, lcb = fc.get(name, (0, 0))
                if not lcb:
                    continue
                n = (lcb - 4) // (4 + cb)
                for k in range(n + 1):
                    v = struct.unpack_from("<I", tb, f0 + 4 * k)[0]
                    struct.pack_into("<I", tb, f0 + 4 * k, self.shift(base_old + v) - base_new)
        for name in CACHES:
            if name in fc and fc[name][1]:
                fc[name] = (fc[name][0], 0)
        # 4) しおり（元のしおり＋書いた欄のしおり）
        marks = [(n_, self.shift(a), self.shift(b)) for n_, a, b in doc.bookmarks]
        for k, e in enumerate(self.edits):
            if e.mark:
                a = self.new_starts[k]
                marks.append((MARK_PREFIX + e.mark, a, a + len(e.text)))
        if marks:
            sttb, bkf, bkl = _bookmark_tables(doc, marks)
            for name, blob in (("SttbfBkmk", sttb), ("PlcfBkf", bkf), ("PlcfBkl", bkl)):
                fc[name] = (len(tb), len(blob))
                tb += blob
        # 5) 新しい文字の書式（CHPX）：置き換える前の字の書式を引き継ぐ
        runs = []
        for k, e in enumerate(self.edits):
            if not e.text:
                continue
            grpprl = doc._lookup(doc.chpx, doc.fc_of(e.style_cp)) if doc.fc_of(e.style_cp) >= 0 else b""
            runs.append((texts_fc[k], texts_fc[k] + 2 * len(e.text), grpprl))
        if runs:
            while len(wd) % 512:
                wd += b"\x00"
            pages = []
            chunk: list = []
            for r in runs:
                if chunk and not _fits(chunk + [r]):
                    pn = len(wd) // 512
                    wd += _chpx_fkp(chunk)
                    pages.append((chunk[0][0], chunk[-1][1], pn))
                    chunk = []
                chunk.append(r)
            if chunk:
                pn = len(wd) // 512
                wd += _chpx_fkp(chunk)
                pages.append((chunk[0][0], chunk[-1][1], pn))
            f0, lcb = fc["PlcfBteChpx"]
            n = (lcb - 4) // 8
            afc = list(struct.unpack_from(f"<{n + 1}I", t, f0))
            apn = list(struct.unpack_from(f"<{n}I", t, f0 + 4 * (n + 1)))
            for start, end, pn in pages:
                afc.append(end)
                apn.append(pn)
            blob = struct.pack(f"<{len(afc)}I", *afc) + struct.pack(f"<{len(apn)}I", *apn)
            fc["PlcfBteChpx"] = (len(tb), len(blob))
            tb += blob
        # 6) FIB：文字の数・ストリームの長さ・表の位置
        lw = list(doc.lw)
        for story in ORDER:
            idx = 3 if story == "text" else SUB[story][0]
            lw[idx] += sum(len(e.text) - (e.b - e.a) for e in self.edits if self.story_of(e.a) == story)
        lw[0] = len(wd)  # cbMac
        struct.pack_into(f"<{len(lw)}i", wd, doc.lw_pos, *lw)
        vals = list(struct.unpack_from(f"<{doc.cb_fclcb * 2}I", doc.wd, doc.fclcb_pos))
        for k, name in enumerate(FC_NAMES):
            if 2 * k + 1 < len(vals) and name in fc:
                vals[2 * k], vals[2 * k + 1] = fc[name]
        struct.pack_into(f"<{len(vals)}I", wd, doc.fclcb_pos, *vals)
        # 7) 複合ファイルに書き戻す（ほかのストリームはそのまま）
        blob = CFB(doc.blob).replace("WordDocument", bytes(wd))
        blob = CFB(blob).replace(doc.table_name, bytes(tb))
        if self.data_patches:
            data = bytearray(doc.data)
            for pos, b in self.data_patches:
                data[pos:pos + len(b)] = b
            blob = CFB(blob).replace("Data", bytes(data))
        return blob


def _fits(runs) -> bool:
    crun = len(runs)
    used = 4 * (crun + 1) + crun + 1
    seen = set()
    for _a, _b, g in runs:
        if g and g not in seen:
            seen.add(g)
            used += 2 + len(g) + 1
    return used <= 510 and crun <= 0x65


def _chpx_fkp(runs) -> bytes:
    """文字の書式の FKP（512 バイト）：rgfc・rgb・CHPX（後ろから）・crun。"""
    page = bytearray(512)
    crun = len(runs)
    rgfc = [r[0] for r in runs] + [runs[-1][1]]
    struct.pack_into(f"<{crun + 1}I", page, 0, *rgfc)
    top = 511
    offs = []
    cache = {}
    for _a, _b, grpprl in runs:
        if not grpprl:
            offs.append(0)
            continue
        if grpprl in cache:
            offs.append(cache[grpprl])
            continue
        size = 1 + len(grpprl)
        start = top - size
        start -= start % 2
        if start <= 4 * (crun + 1) + crun:
            raise FillError("文字の書式が多すぎて、1つのページに入りません")
        page[start] = len(grpprl)
        page[start + 1:start + 1 + len(grpprl)] = grpprl
        top = start
        offs.append(start // 2)
        cache[grpprl] = start // 2
    for k, o in enumerate(offs):
        page[4 * (crun + 1) + k] = o
    page[511] = crun
    return bytes(page)


def _bookmark_tables(doc: DocFile, marks) -> tuple[bytes, bytes, bytes]:
    """しおりの表（名前の STTB・始まりの PLC・終わりの PLC）を作る。始まりの順に並べる。"""
    t = doc.table
    # 元の FBKF の bkc（表の列の範囲）を保つ
    old_bkc = {}
    f0, lcb = doc.fc["PlcfBkf"]
    if lcb:
        n = (lcb - 4) // 8
        for k, (name, _a, _b) in enumerate(doc.bookmarks):
            if k < n:
                old_bkc[name] = struct.unpack_from("<H", t, f0 + 4 * (n + 1) + 4 * k + 2)[0]
    order = sorted(range(len(marks)), key=lambda i: (marks[i][1], i))
    ends = sorted(range(len(marks)), key=lambda i: (marks[i][2], i))
    end_index = {i: k for k, i in enumerate(ends)}
    sttb = bytearray(struct.pack("<HHH", 0xFFFF, len(marks), 0))
    starts = []
    fbkf = bytearray()
    for i in order:
        name, a, b = marks[i]
        sttb += struct.pack("<H", len(name)) + name.encode("utf-16-le")
        starts.append(a)
        fbkf += struct.pack("<HH", end_index[i], old_bkc.get(name, 0))
    last = doc.ccp["text"] + 1
    bkf = struct.pack(f"<{len(starts) + 1}I", *starts, max(starts + [last])) + bytes(fbkf)
    bkl = struct.pack(f"<{len(ends) + 1}I", *[marks[i][2] for i in ends], max([marks[i][2] for i in ends] + [last]))
    return bytes(sttb), bkf, bkl


# ====================================================================== 入口
def para_cps(p) -> list[int]:
    """段落の文字（p.text）の1字ごとの位置（CP）。"""
    out = []
    for sg in p.segs:
        if sg.hidden:
            continue
        c0 = sg.node.s if sg.node is not None else None
        for k in range(len(sg.text)):
            out.append(c0 + k if c0 is not None and sg.kind == "t" else (c0 if c0 is not None else -1))
    return out


def fill(doc: DocFile, values: dict, *, force: bool = False) -> DocFillResult:
    from ..docx.fill import Patch

    from ..docx.window import WriteError, resolve_write

    sp = Spots(doc)
    w = DocWriter(doc)
    patches = []
    problems = []
    used: dict = {}
    for key, value in values.items():
        it = sp.items.get(key)
        anchor = None
        try:
            if key.startswith("@"):
                # 規則で示していない所：エージェントが字を目印にして指した所
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    it = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                anchor = dict(value)
                key, value = it["id"], it["value"]
            if it is None:
                raise FillError(f"{key}: そのような欄はありません")
            if doc.protection_enforced and not force:
                if doc.protection != "forms" or it.get("control") is None:
                    raise FillError(f"{key}: この文書は保護されていて（{doc.protection}）、ここには書けません")
            c = it.get("control")
            mark = key.replace(".", "_").replace("#", "_")
            if c is not None:
                if c.kind == "checkbox":
                    on = _as_bool(value)
                    _set_ffbits(doc, w, c, res=1 if on else 0)
                    patches.append(Patch(key, "control", on, "doc"))
                    continue
                if c.kind == "dropdown":
                    text = _is_text_value(value)
                    if text not in c.options:
                        raise FillError(f"{key}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
                    _set_ffbits(doc, w, c, res=c.options.index(text))
                    patches.append(Patch(key, "control", text, "doc"))
                    continue
                text = _is_text_value(value)
                if c.maxlen and len(text) > c.maxlen:
                    raise FillError(f"{key}: 「{text}」は {len(text)} 字です（この欄は {c.maxlen} 字まで）")
                cps = [s.node.s for s in c.result_segs if s.node is not None and s.kind == "t"]
                if cps:
                    a = cps[0]
                    b = max(s.node.s + len(s.text) for s in c.result_segs if s.node is not None and s.kind == "t")
                else:
                    a = b = c.end_cp
                w.add(Edit(a, b, text, a if cps else c.end_cp, None))
                patches.append(Patch(key, "input", text, "doc"))
                continue
            if key in sp.marks:
                raise FillError(f"{key}: すでに書き込んだ欄です")
            p = it["para"]
            kind = it["kind"]
            cps = para_cps(p)
            if kind == "cell":
                text = _is_text_value(value)
                if p.text and not p.text.strip(" 　  "):
                    a, b = cps[0], cps[-1] + 1
                else:
                    a = b = p.cp1  # マスの終わりの印（0x07）の前
                w.add(Edit(a, b, text, p.cp1 if a == b else a, mark))
                patches.append(Patch(key, "input", text, "doc"))
            elif kind == "checkbox":
                on = _as_bool(value)
                a0, b0 = it["span"]
                orig = p.text[a0:b0]
                if on:
                    w.add(Edit(cps[a0], cps[a0] + 1, "☑" if orig == "☐" else "■", cps[a0], mark))
                patches.append(Patch(key, "control", on, "doc"))
            elif kind == "inline":
                text = _is_text_value(value)
                a0, b0 = it["span"]
                if a0 == b0:
                    # 挿入：直前の字の後ろ（段落の終わりなら、段落の印の前）。段落の頭は、初めの字の前（書式も初めの字から）
                    if a0 == 0 and cps and cps[0] >= 0:
                        a = b = cps[0]
                    elif a0 == 0 and cps:
                        raise FillError(f"{key}: 書き足す位置が、書き換えられない部分にあります")
                    else:
                        a = b = (cps[a0 - 1] + 1) if a0 > 0 and cps[a0 - 1] >= 0 else p.cp1
                else:
                    a, b = cps[a0], cps[b0 - 1] + 1
                    if b - a != b0 - a0 or -1 in cps[a0:b0]:
                        raise FillError(f"{key}: 空欄が、書き換えられない部分にかかっています")
                w.add(Edit(a, b, text, a, mark))
                patches.append(Patch(key, "input", text, "doc", anchor=anchor))
            else:
                raise FillError(f"{key}: この種類の欄（{kind}）には書けません")
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    return DocFillResult(w.build(), patches)


def _set_ffbits(doc: DocFile, w: DocWriter, c, res: int) -> None:
    pos = getattr(c, "ffdata", {}).get("bits_pos") if getattr(c, "ffdata", None) else None
    if pos is None:
        raise FillError(f"{c.name or 'フォームの欄'}: 欄のデータ（FFData）が見つかりません")
    bits = struct.unpack_from("<H", doc.data, pos)[0]
    bits = (bits & ~(0x1F << 2)) | ((res & 0x1F) << 2)
    w.data_patches.append((pos, struct.pack("<H", bits)))
