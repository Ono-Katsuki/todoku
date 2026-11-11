"""書く前の文書のページ数の記録と、行が増えたときに伝える事実（ページのある文書：.docx・.doc・.odt・RTF・HWPX）。

ページ数はレイアウトの結果で、todoku はレイアウトをしない。分かるのは次のことだけなので、その事実だけを伝える：
- 保存したソフトが文書に残したページ数：.docx は docProps/app.xml の Pages（ECMA-376 第 1 部 22.2.2.19）、.doc は
  SummaryInformation の PIDSI_PAGECOUNT（14。MS-OLEPS の property set を読む）、.odt は meta.xml の
  meta:document-statistic の meta:page-count（OpenDocument 1.3 第 3 部 4.3.2.18）、RTF は \info の \nofpages
  （Rich Text Format Specification 1.9.1）。HWPX には、ページ数の記録がない
- 書いたことで行が増える見込み（wrapcheck の見積もり）の段落と行の数

書いた後のページ数は見積もらない。ページの終わりの余りが分からないので、増えるとも増えないとも言わない。
"""

from __future__ import annotations

import re
import struct


def _docx_pages(blob: bytes) -> int | None:
    from .zipraw import RawZip

    z = RawZip(blob)
    if not z.has("docProps/app.xml"):
        return None
    m = re.search(rb"<(?:\w+:)?Pages>\s*(\d+)\s*</(?:\w+:)?Pages>", z.read("docProps/app.xml"))
    return int(m[1]) if m else None


def _doc_pages(blob: bytes) -> int | None:
    from .cfb import CFB

    name = "\x05SummaryInformation"
    cf = CFB(blob)
    if not cf.has(name):
        return None
    d = cf.read(name)
    # PropertySetStream：頭 28 バイトの後に (FMTID 16 バイト, 位置 4 バイト) の並び。最初の組が SummaryInformation（MS-OLEPS 2.21）
    if len(d) < 48 or struct.unpack_from("<I", d, 24)[0] < 1:
        return None
    at = struct.unpack_from("<I", d, 44)[0]
    if at + 8 > len(d):
        return None
    n = struct.unpack_from("<I", d, at + 4)[0]
    for i in range(min(n, (len(d) - at - 8) // 8)):
        pid, off = struct.unpack_from("<II", d, at + 8 + 8 * i)
        if pid != 14:  # PIDSI_PAGECOUNT
            continue
        p = at + off
        if p + 8 > len(d):
            return None
        vt, v = struct.unpack_from("<Ii", d, p)
        return v if vt == 3 and v > 0 else None  # VT_I4
    return None


def _odt_pages(blob: bytes) -> int | None:
    from .zipraw import RawZip

    z = RawZip(blob)
    if not z.has("meta.xml"):
        return None
    m = re.search(rb"<meta:document-statistic\b[^>]*\bmeta:page-count=\"(\d+)\"", z.read("meta.xml"))
    return int(m[1]) if m and int(m[1]) > 0 else None


def _rtf_pages(blob: bytes) -> int | None:
    from .rtf.lexer import tokens

    depth = info = 0
    for t in tokens(blob):
        if t.kind == "open":
            depth += 1
        elif t.kind == "close":
            if info and depth == info:
                return None  # \info を読み終えた
            depth -= 1
        elif t.kind == "word":
            if t.word == "info" and not info:
                info = depth
            elif t.word == "nofpages" and info:
                return t.param if t.param and t.param > 0 else None
            elif t.word in ("pard", "par", "sectd") and not info and depth <= 1:
                return None  # 本文が始まった（\info は本文の前にある）
    return None


_PAGES = {"docx": _docx_pages, "doc": _doc_pages, "odt": _odt_pages, "rtf": _rtf_pages}


def recorded_pages(blob: bytes, k: str) -> int | None:
    """保存したソフトが文書に残したページ数（なければ・読めなければ None）。"""
    try:
        return _PAGES[k](blob) if k in _PAGES else None
    except Exception:  # noqa: BLE001  記録が読めないだけ。書き込みは止めない
        return None


def growth_note(blob: bytes, k: str, wrapped: list[str]) -> tuple[dict, str] | None:
    """行が増える見込みがあるとき：(結果に入れる事実, 伝える言葉)。増えないときは None。"""
    added = sum(getattr(w, "added", 0) for w in wrapped)
    if added <= 0:
        return None
    pages = recorded_pages(blob, k)
    n = sum(1 for w in wrapped if getattr(w, "added", 0))
    facts = {"paragraphs": n, "lines_added": added, "recorded_pages_before": pages}
    where = f"元の文書のページ数の記録は {pages} ページです（保存したソフトが残した数）" if pages else "元の文書にページ数の記録はありません"
    return facts, (f"ページが増えるかもしれません：書いたことで行が増える見込みの段落が {n} か所（合わせて約 {added} 行）あり、{where}。"
                   "ページの終わりにどれだけ余りがあるかは todoku には分からないので、書いた後のページ数は分かりません。"
                   "ページ数が決まっている様式なら、折り返す所の字を短くするか、PDF にしてページ数を確かめてください（ページ数は、開くソフトによって違うことがあります）")
