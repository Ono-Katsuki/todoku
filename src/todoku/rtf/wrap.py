"""RTF：書いたことで段落の行が折り返すかの見積もり（目安。wrapcheck）。

書いた段落（字が変わった段落：wrapcheck.written_places）ごとに、書く前と書いた後の字を同じ幅で並べて、行の数を比べる（docx.wrap と同じ）。
幅と字の大きさは、書いた後のファイルの制御語を、書いた所まで順に読んで決める（Rich Text Format Specification 1.9.1）：
- 表のマスの段落：マスの左右の端（\\cellx。最初のマスの左は \\trleft）から、マスの左右の余白（\\trpaddl・\\trpaddr。なければ
  \\trgaph）を引いた幅。マスごとの余白（\\clpadl など）は見ない
- ほかの段落：区切りのページの幅（\\pgwsxn。なければ文書の \\paperw、既定 12240 twip）から左右の余白（\\marglsxn・\\margrsxn。
  なければ \\margl・\\margr、既定 1800 twip）を引き、段組み（\\cols・\\colsx）の数で分けた幅
- どちらも、段落の字下げ（\\li・\\ri）を引き、1 行目の字下げ（\\fi）を見る。字の大きさは \\fs（半ポイント。既定 24）
スタイル（\\s）の中身は読まない（RTF は、スタイルの書式を段落ごとに書き出す決まり）。テキストボックスの中の段落は見ない。
"""

from __future__ import annotations

from ..wrapcheck import place_name, say, wraps, written_places
from .lexer import tokens

_PARA = ("li", "ri", "fi")
_SECT = ("pgwsxn", "marglsxn", "margrsxn", "cols", "colsx")
_DOC = {"paperw": 12240, "margl": 1800, "margr": 1800}


def _states(blob: bytes, offsets: list[int]) -> dict[int, dict]:
    """バイトの位置 → その位置での書式（fs・li・ri・fi と、区切り・文書のページの値）。"""
    want = sorted(set(offsets))
    out: dict[int, dict] = {}
    if not want:
        return out
    doc = dict(_DOC)
    sect: dict = {}
    cur = {"fs": 24, "li": 0, "ri": 0, "fi": 0}
    stack: list[dict] = []
    k = 0
    for t in tokens(blob):
        while k < len(want) and t.s >= want[k]:
            out[want[k]] = {**cur, **doc, **sect}
            k += 1
        if k >= len(want):
            break
        if t.kind == "open":
            stack.append(dict(cur))
        elif t.kind == "close":
            if stack:
                cur = stack.pop()
        elif t.kind == "word":
            w, v = t.word, t.param
            if w == "fs" and v:
                cur["fs"] = v
            elif w == "plain":
                cur["fs"] = 24
            elif w == "pard":
                cur["li"] = cur["ri"] = cur["fi"] = 0
            elif w in _PARA and v is not None:
                cur[w] = v
            elif w == "sectd":
                sect = {}
            elif w in _SECT and v is not None:
                sect[w] = v
            elif w in doc and v is not None and len(stack) <= 1:
                doc[w] = v
    for o in want[k:]:
        out[o] = {**cur, **doc, **sect}
    return out


def _offset(p, a: int) -> int | None:
    """段落の字 a 番の、ファイルの中の位置（なければ、段落の初め）。"""
    pos = p.positions()
    if pos:
        si, k = pos[min(a, len(pos) - 1)]
        src = p.segs[si].node
        if src is not None:
            return src.starts[k] if getattr(src, "starts", None) is not None and k < len(src.starts) else src.s
    return getattr(p, "start", None)


def wrap_warnings(before, after) -> list[str]:
    """書いた段落のうち、書いたことで行が増える見込みのものの警告。before・after は RtfDocument（書く前・書いた後）。"""
    from ..docx.document import iter_paras

    def paras(doc):
        return [p for _k, _p, blocks in doc.all_blocks() for p in iter_paras(blocks)]

    pa, pb = paras(after), paras(before)
    if len(pa) != len(pb):
        return []
    from ..lazy.blocks import _control_ids

    todo = []
    for i, (names, k) in sorted(written_places(pa, pb, after.marks(), _control_ids(after)).items()):
        at = _offset(pa[i], k)
        if at is not None:
            todo.append((i, names, at))
    states = _states(after.blob, [at for _i, _n, at in todo])
    out = []
    for i, names, at in todo:
        a, b, st = pa[i], pb[i], states[at]
        size = st["fs"] / 2
        geom = getattr(a.cell, "x_tw", None) if a.cell is not None else None
        if a.cell is not None:
            if not geom:
                continue
            width, basis = (geom[1] - geom[0] - geom[2] - geom[3]) / 20, "表のマスの幅"
        else:
            cols = max(st.get("cols", 1), 1)
            page = st.get("pgwsxn", st["paperw"]) - st.get("marglsxn", st["margl"]) - st.get("margrsxn", st["margr"])
            width = (page - (cols - 1) * st.get("colsx", 720)) / cols / 20
            basis = "ページの幅（余白を引いた幅）" if cols == 1 else f"段組み（{cols} 段）の 1 段の幅"
        avail = width - (st["li"] + st["ri"]) / 20
        if avail <= 0:
            continue
        got = wraps(b.text, a.text, avail, size, size, st["fi"] / 20)
        if got is None:
            continue
        out.append(say(place_name(names), a.text, got[0], got[1],
                       f"{basis} 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイント", got[2]))
    return out
