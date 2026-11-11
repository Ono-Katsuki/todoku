"""古い Word（.doc）：書いたことで段落の行が折り返すかの見積もり（目安。wrapcheck）。

書いた段落（書いた所の印のしおりのある段落）ごとに、書く前と書いた後の字を同じ幅で並べて、行の数を比べる（docx.wrap と同じ）。
幅（MS-DOC の書式の命令から）：
- 表のマスの段落：その行の終わりの段落の sprmTDefTable の列の境（rgdxaCenter）の差から、マスの左右の余白（sprmTDxaGapHalf の
  2 倍。なければ仮に 108 twip ずつと置き、見積もりの元に「仮に」と書く）と段落の字下げを引いた幅
- ほかの段落：段落のある区切り（PlcfSed の SEPX）の sprmSXaPage・sprmSDxaLeft・sprmSDxaRight（なければ Word の既定
  12240・1800・1800 twip）と段組み（sprmSCcolumns と、段の間 sprmSDxaColumns。なければ MS-DOC の既定 720 twip）から、段落の字下げ（sprmPDxaLeft・sprmPDxaRight・sprmPDxaLeft1）を引いた幅
字の大きさは、字の書式の sprmCHps、なければ字のスタイル（sprmCIstd）・段落のスタイル（STSH の UPX の sprmCHps を
istdBase でたどる）、どれもなければ 10 pt（MS-DOC の既定）。
"""

from __future__ import annotations

import struct

from ..wrapcheck import place_name, say, wraps, written_places
from .book import sprms

_DEFAULT_GAP = 108


def _i16(op: bytes) -> int:
    return struct.unpack_from("<h", op, 0)[0] if len(op) >= 2 else 0


def _u16(op: bytes) -> int:
    return struct.unpack_from("<H", op, 0)[0] if len(op) >= 2 else 0


def _styles(doc) -> dict:
    """STSH：istd → (istdBase, 段落の書式の grpprl, 字の書式の grpprl)。"""
    memo = getattr(doc, "_wrap_styles", None)
    if memo is not None:
        return memo
    out: dict = {}
    fc, lcb = doc.fc.get("Stshf", (0, 0))
    t = doc.table
    try:
        if lcb and fc + lcb <= len(t):
            cb_stshi = _u16(t[fc:fc + 2])
            cstd = _u16(t[fc + 2:fc + 4])
            cb_base = _u16(t[fc + 4:fc + 6])
            i = fc + 2 + cb_stshi
            end = fc + lcb
            for istd in range(cstd):
                if i + 2 > end:
                    break
                cb = _u16(t[i:i + 2])
                std = t[i + 2:i + 2 + cb]
                i += 2 + cb
                if cb < 10:
                    continue
                stk = std[2] & 0x0F
                base = _u16(std[2:4]) >> 4
                cupx = std[4] & 0x0F
                k = cb_base
                if k + 2 > len(std):
                    continue
                cch = _u16(std[k:k + 2])
                k += 2 + 2 * cch + 2
                upx = []
                for _ in range(cupx):
                    if k + 2 > len(std):
                        break
                    n = _u16(std[k:k + 2])
                    upx.append(std[k + 2:k + 2 + n])
                    k += 2 + n + (n & 1)
                papx, chpx = b"", b""
                if stk == 1 and len(upx) >= 2:
                    papx, chpx = upx[0][2:], upx[1]
                elif stk == 2 and upx:
                    chpx = upx[0]
                elif stk == 3 and len(upx) >= 3:
                    papx, chpx = upx[1][2:], upx[2]
                out[istd] = (base, papx, chpx)
    except (struct.error, IndexError):
        out = {}
    doc._wrap_styles = out
    return out


def _style_hps(doc, istd: int) -> int | None:
    st = _styles(doc)
    seen = 0
    while istd in st and seen < 20:
        base, _papx, chpx = st[istd]
        for sp, op in sprms(chpx):
            if sp == 0x4A43:
                return _u16(op)
        if base == 0x0FFF or base == istd:
            break
        istd = base
        seen += 1
    return None


def _papx_istd(doc, fc: int) -> int:
    """段落の書式（FKP の PAPX）の istd（段落のスタイル）。"""
    pfc, lcb = doc.fc["PlcfBtePapx"]
    if not lcb:
        return 0
    n = (lcb - 4) // 8
    t = doc.table
    pns = struct.unpack_from(f"<{n}I", t, pfc + (n + 1) * 4)
    for pn in pns:
        page = doc.wd[pn * 512:(pn + 1) * 512]
        if len(page) < 512:
            continue
        crun = page[511]
        rgfc = struct.unpack_from(f"<{crun + 1}I", page, 0)
        for k in range(crun):
            if rgfc[k] <= fc < rgfc[k + 1]:
                boff = page[4 * (crun + 1) + 13 * k] * 2
                if not boff:
                    return 0
                cb = page[boff]
                at = boff + 2 if cb == 0 else boff + 1
                return _u16(page[at:at + 2])
    return 0


def _sizes(doc, p) -> list[float]:
    """段落の字ごとの大きさ（pt）。"""
    para_istd = _papx_istd(doc, doc.fc_of(p.cp1)) if doc.fc_of(p.cp1) >= 0 else 0
    para_hps = _style_hps(doc, para_istd) or _style_hps(doc, 0) or 20
    out = []
    for sg in p.segs:
        if sg.hidden:
            continue
        c0 = sg.node.s if sg.node is not None else None
        for j in range(len(sg.text)):
            hps = None
            if c0 is not None:
                fc = doc.fc_of(c0 + j)
                if fc >= 0:
                    istd_c = None
                    for sp, op in sprms(doc._lookup(doc.chpx, fc)):
                        if sp == 0x4A43:
                            hps = _u16(op)
                        elif sp == 0x4A30:
                            istd_c = _u16(op)
                    if hps is None and istd_c is not None:
                        hps = _style_hps(doc, istd_c)
            out.append((hps or para_hps) / 2)
    return out


def _papx(doc, cp: int) -> bytes:
    fc = doc.fc_of(cp)
    return doc._lookup(doc.papx, fc) if fc >= 0 else b""


def _cell_width(doc, p, facts: dict | None = None) -> float | None:
    """表のマスの、字を置ける幅（pt。段落の字下げは引く前）。facts に、余白を仮に置いたか（gap_assumed）を入れる。"""
    if p.cell is None:
        return None
    text = doc.text
    for cp in range(p.cp1, min(len(text), p.cp1 + 20000)):
        if text[cp] != "\x07":
            continue
        props = doc.para_props(cp)
        if not (props["ttp"] or props["inner_ttp"]):
            continue
        centers, gap = None, None
        for sp, op in sprms(_papx(doc, cp)):
            if sp == 0xD608 and op:
                n = op[0]
                if len(op) >= 1 + 2 * (n + 1):
                    centers = list(struct.unpack_from(f"<{n + 1}h", op, 1))
            elif sp == 0x9602:
                gap = _i16(op)
        col = p.cell.col
        if centers is None or col + 1 >= len(centers):
            return None
        w = centers[col + 1] - centers[col] - 2 * (_DEFAULT_GAP if gap is None else gap)
        if facts is not None:
            facts["gap_assumed"] = gap is None
        return w / 20
    return None


def _section(doc, cp: int) -> tuple[float, float, float, int, float]:
    """(ページの幅, 左の余白, 右の余白, 段の数, 段の間)（twip）。"""
    xa, left, right, cols, space = 12240, 1800, 1800, 1, 720
    fc, lcb = doc.fc.get("PlcfSed", (0, 0))
    if lcb:
        n = (lcb - 4) // 16
        t = doc.table
        cps = struct.unpack_from(f"<{n + 1}I", t, fc)
        for k in range(n):
            if cps[k] <= cp < cps[k + 1] or k == n - 1:
                sed = t[fc + 4 * (n + 1) + 12 * k: fc + 4 * (n + 1) + 12 * (k + 1)]
                fcs = struct.unpack_from("<i", sed, 2)[0]
                if fcs >= 0 and fcs + 2 <= len(doc.wd):
                    cb = _u16(doc.wd[fcs:fcs + 2])
                    for sp, op in sprms(doc.wd[fcs + 2:fcs + 2 + cb]):
                        if sp == 0xB01F:
                            xa = _u16(op)
                        elif sp == 0xB021:
                            left = _u16(op)
                        elif sp == 0xB022:
                            right = _u16(op)
                        elif sp == 0x500B:
                            cols = _u16(op) + 1
                        elif sp == 0x900C:  # sprmSDxaColumns：段の間の幅（MS-DOC 2.6.4）
                            space = _u16(op)
                break
    return xa, left, right, cols, space


def _ind(doc, p, size: float) -> tuple[float, float, float]:
    left = right = first = 0
    for sp, op in sprms(_papx(doc, p.cp1)):
        if sp in (0x840F, 0x845E):
            left = _i16(op)
        elif sp in (0x840E, 0x845D):
            right = _i16(op)
        elif sp in (0x8411, 0x8460):
            first = _i16(op)
    return left / 20, right / 20, first / 20


def wrap_warnings(before, after, rename: dict | None = None) -> list[str]:
    """書いた段落のうち、書いたことで行が増える見込みのものの警告。before・after は DocFile（書く前・書いた後）。"""
    from ..docx.document import iter_paras

    def paras(doc):
        return [p for _k, _p, blocks in doc.all_blocks() for p in iter_paras(blocks)]

    pa, pb = paras(after), paras(before)
    if len(pa) != len(pb):
        return []
    from ..lazy.blocks import _control_ids

    out = []
    for i, (names, _at) in sorted(written_places(pa, pb, after.marks(), _control_ids(after)).items()):
        a, b = pa[i], pb[i]
        sa, sb = _sizes(after, a), _sizes(before, b)
        size = max(sa) if sa else 10.0
        left, right, first = _ind(after, a, size)
        facts: dict = {}
        width = _cell_width(after, a, facts)
        basis = "表のマスの幅" + (f"（マスの余白は書かれていないので、仮に {_DEFAULT_GAP / 20:g} ポイント）" if facts.get("gap_assumed") else "")
        if width is None:
            if a.cell is not None:
                continue  # 表のマスの幅が読めない
            xa, ml, mr, cols, space = _section(after, a.cp0)
            width = ((xa - ml - mr) - space * (cols - 1)) / cols / 20
            basis = "ページの幅（余白を引いた幅）"
        avail = width - left - right
        got = wraps(b.text, a.text, avail, lambda k, s=sb: s[k] if k < len(s) else size,
                    lambda k, s=sa: s[k] if k < len(s) else size, first)
        if got is None:
            continue
        place = place_name(list(dict.fromkeys((rename or {}).get(n, n) for n in names)))
        out.append(say(place, a.text, got[0], got[1], f"{basis} 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイント", got[2]))
    return out
