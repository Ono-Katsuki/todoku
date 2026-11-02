"""古い Word（.doc、Word 97〜2003 のバイナリ）を読む（MS-DOC）。外部のライブラリは使わない。

- 複合ファイル（CFB）の WordDocument・1Table/0Table・Data のストリーム
- FIB（文書の目次）→ ピースの表（CLX）で、文字の位置（CP）とストリームの中の位置（FC）を結ぶ
- 段落の書式（PAPX）で表のマス・行の終わりを、文字の書式（CHPX）で非表示・下線・記号・フォームの欄のデータの位置を読む
- フィールド（0x13・0x14・0x15）と、フォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN、FFData）
- しおり（SttbfBkmk・PlcfBkf・PlcfBkl）
読んだ結果は、Word（.docx）と同じ形（段落・表・マス・入力欄）にして、窓を共通にする。
"""

from __future__ import annotations

import struct
from dataclasses import dataclass, field

from ..cfb import CFB
from ..docx.document import WINGDINGS, Cell, Control, Para, Seg, Table

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
                if size == 255:
                    size = 0
            else:
                size = grpprl[i]
                i += 1
        yield sprm, grpprl[i:i + size]
        i += size


class DocFile:
    kind = "doc"

    def __init__(self, blob: bytes, filename: str = ""):
        self.blob = blob
        self.filename = filename
        self.cfb = CFB(blob)
        if not self.cfb.has("WordDocument"):
            raise DocError("Word の文書（WordDocument）がありません")
        self.wd = self.cfb.read("WordDocument")
        wd = self.wd
        ident, self.nfib = struct.unpack_from("<HH", wd, 0)
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
        self.pieces = self._pieces()
        self.papx = self._fkps("PlcfBtePapx", papx=True)
        self.chpx = self._fkps("PlcfBteChpx", papx=False)
        self.text = self._text()
        self.protection, self.protection_enforced = self._dop()
        self.has_macros = self.cfb.has("Macros/VBA/dir") or any(p.lower().startswith("macros/") for p in self.cfb.streams())
        self.controls: list[Control] = []
        self.bookmarks = self._bookmarks()
        self.stories: list[tuple[str, str, list]] = []
        self._build()

    # ------------------------------------------------------------------ FIB・ピース
    def _pieces(self) -> list[Piece]:
        fc, lcb = self.fc["Clx"]
        t = self.table
        i = fc
        while i < fc + lcb:
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
                out.append(self.wd[p.fc:p.fc + 2 * n].decode("utf-16-le", errors="replace"))
        return "".join(out)

    def fc_of(self, cp: int) -> int:
        for p in self.pieces:
            if p.cp0 <= cp < p.cp1:
                return p.fc + (cp - p.cp0) * (1 if p.compressed else 2)
        return -1

    def piece_of(self, cp: int) -> Piece | None:
        for p in self.pieces:
            if p.cp0 <= cp < p.cp1:
                return p
        return None

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

    def char_props(self, cp: int) -> dict:
        fc = self.fc_of(cp)
        d = {"hidden": False, "underline": False, "pic": None, "symbol": None, "spec": False, "data": False, "deleted": False}
        for sp, op in sprms(self._lookup(self.chpx, fc)):
            # 命令の番号は MS-DOC の原本（2.6.1）で確かめた：0x083C 非表示（0x0839 は影付きで別のもの）、0x0800 変更履歴の削除
            if sp == 0x083C:
                d["hidden"] = bool(op[0] & 1) if op else False
            elif sp == 0x0800:
                d["deleted"] = bool(op[0] & 1) if op else False
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
        return d

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
        for cp in range(a, b):
            ch = text[cp] if cp < len(text) else "\r"
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

    def _para(self, s: int, e: int) -> Para:
        p = Para(None, "doc", [])
        p.cp0, p.cp1 = s, e
        text = self.text
        stack: list[dict] = []
        cp = s
        while cp < e:
            ch = text[cp]
            o = ord(ch)
            if o == 0x13:
                stack.append({"begin": cp, "instr": "", "phase": "instr"})
                cp += 1
                continue
            if o == 0x14 and stack:
                f = stack[-1]
                f["phase"] = "result"
                f["sep"] = cp
                f["ctl"] = self._form_field(f, p)
                f["n0"] = len(p.segs)
                cp += 1
                continue
            if o == 0x15 and stack:
                f = stack.pop()
                if "ctl" not in f:
                    f["ctl"] = self._form_field(f, p)
                    f["n0"] = len(p.segs)
                ctl = f.get("ctl")
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
            if o == 0x09:
                p.segs.append(Seg("\t", "tab", CPRef(cp), None, hidden, props["underline"]))
            elif o in (0x0B, 0x0C, 0x0E):
                p.segs.append(Seg("\n", "br", CPRef(cp), None, hidden))
            elif o == 0x1E:
                p.segs.append(Seg("‑", "other", CPRef(cp), None, hidden))
            elif o == 0x1F:
                pass
            elif o in (0x01, 0x08) and props["spec"]:
                p.segs.append(Seg("〔図〕", "other", CPRef(cp), None, hidden))
            elif o in (0x02, 0x05) and props["spec"]:
                p.segs.append(Seg("※", "note", CPRef(cp), None, hidden))
            elif o == 0x28 and props["symbol"] and props["spec"]:
                code = props["symbol"][1]
                code2 = code + 0xF000 if code < 0xF000 else code
                p.segs.append(Seg(WINGDINGS.get(code2, "◆"), "sym", CPRef(cp), None, hidden, props["underline"]))
            elif o < 0x20 and o not in (0x09,):
                pass
            else:
                # 同じ書式の字はまとめる
                if p.segs and p.segs[-1].kind == "t" and p.segs[-1].hidden == hidden and \
                        p.segs[-1].underline == props["underline"] and p.segs[-1].node.s + len(p.segs[-1].text) == cp \
                        and p.segs[-1].field is None and not stack:
                    p.segs[-1].text += ch.replace("\xa0", " ")
                else:
                    p.segs.append(Seg(ch.replace("\xa0", " "), "t", CPRef(cp), None, hidden, props["underline"]))
            cp += 1
        return p

    def _form_field(self, f: dict, p: Para) -> Control | None:
        instr = f["instr"].strip()
        word = instr.split()[0].upper() if instr.split() else ""
        if word not in ("FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN"):
            return None
        kind = {"FORMTEXT": "text", "FORMCHECKBOX": "checkbox", "FORMDROPDOWN": "dropdown"}[word]
        ctl = Control(kind, "formfield", "", None, "doc", p)
        ctl.begin_cp = f["begin"]
        props = self.char_props(f["begin"])
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
        d = {"_pos": i, "type": itype, "res": res, "type_txt": type_txt, "cch": cch, "name": name, "bits_pos": i + 4}
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
