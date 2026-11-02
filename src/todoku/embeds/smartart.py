"""SmartArt の字を、つながりの順の入れ子の一覧にする（ECMA-376 第1部 21.4 DrawingML Diagrams の dgm:dataModel）。

データの部品（word/diagrams/dataN.xml など）の点（dgm:pt）のうち、字を持つのは種類が node（既定）と asst（補佐）の点
（dgm:t の a:p）。つながり（dgm:cxn、種類 parOf が既定）の srcId → destId が親 → 子、srcOrd が兄弟の順。
いちばん上（種類 doc の点）から順にたどり、たどれない点の字も最後に落とさず出す。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

from .chart import _kid, _kids, _ln, _rich


def smartart_lines(xml: bytes) -> list[str]:
    """字の行（入れ子は先頭の全角の空白の数で示す）。"""
    root = ET.fromstring(xml)
    pts: dict[str, tuple[str, str]] = {}
    order: list[str] = []
    for pt in _kids(_kid(root, "ptLst"), "pt"):
        mid = pt.get("modelId", "")
        typ = pt.get("type", "node")
        t = _kid(pt, "t")
        pts[mid] = (typ, _rich(t).replace("\n", " ／ ") if t is not None else "")
        order.append(mid)
    kids: dict[str, list[tuple[int, str]]] = {}
    for c in _kids(_kid(root, "cxnLst"), "cxn"):
        if c.get("type", "parOf") != "parOf":
            continue
        kids.setdefault(c.get("srcId", ""), []).append((int(c.get("srcOrd", "0") or 0), c.get("destId", "")))
    out: list[str] = []
    seen: set[str] = set()

    def walk(mid: str, depth: int):
        if mid in seen or depth > 50:
            return
        seen.add(mid)
        typ, text = pts.get(mid, ("", ""))
        shown = typ in ("node", "asst")
        if shown:
            out.append("　" * depth + (text or "（空）"))
        for _o, k in sorted(kids.get(mid, [])):
            walk(k, depth + (1 if shown else 0))

    for mid in order:
        if pts[mid][0] == "doc":
            walk(mid, 0)
    for mid in order:
        typ, text = pts[mid]
        if mid not in seen and typ in ("node", "asst") and text:
            out.append(text)
    return out


def smartart_texts(xml: bytes) -> list[str]:
    """（検査用）データの部品の a:t の字を、そのまま全部。"""
    root = ET.fromstring(xml)
    return [t.text for t in root.iter() if _ln(t.tag) == "t" and t.text]
