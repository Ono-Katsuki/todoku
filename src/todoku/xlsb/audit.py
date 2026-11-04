""".xlsb の検査：窓とは別に、シートの部品のレコードを数えて、値のあるセルがすべて窓に出ているかを照らす。"""

from __future__ import annotations

import struct

from ..xlsx.audit import Audit, Check, check_markdown
from .book import C_BLANK, CELLS, ROW, XlsbWorkbook, records


def raw_values(wb) -> dict[str, set[tuple[int, int]]]:
    """シート → 値のあるセル（レコードの行・列をそのまま読む。XlsbWorkbook の読み込みとは別に歩く）。"""
    out: dict[str, set] = {}
    for sh in wb.sheets:
        if sh.state.startswith("non-worksheet"):
            continue
        data = wb.zip.read(sh.path)
        row = -1
        cells = out.setdefault(sh.name, set())
        for rid, p, _ln in records(data):
            if rid == ROW:
                row = struct.unpack_from("<I", data, p)[0]
            elif rid in CELLS and rid != C_BLANK:
                cells.add((row + 1, struct.unpack_from("<I", data, p)[0] + 1))
    return out


def run(blob: bytes, filename: str = "") -> Audit:
    from ..formula.parser import parse_cell
    from ..xlsx import window
    from ..xlsx.calc import verify

    wb = XlsbWorkbook(blob, filename)
    doc = window.build(wb)
    au = Audit(filename)
    raw = raw_values(wb)
    missing = []
    for sh in doc["sheets"]:
        shown = {parse_cell(c["cell"]) for c in sh["cells"] if c.get("value") is not None}
        missing += [f"{sh['name']}!{r},{c}" for r, c in sorted(raw.get(sh["name"], set()) - shown)]
    n = sum(len(v) for v in raw.values())
    au.checks.append(Check("読み落とし（セル）", not missing, f"シートの部品のレコードから、値のあるセル {n:,} 個を数え、すべて窓のセルと照らしました", missing[:20]))
    links = {(c.link_cell[0], c.link_cell[1], c.link_cell[2]) for s in wb.sheets for c in s.controls if c.link_cell}
    au.checks.append(check_markdown(doc, links))
    rep = verify(wb).summary()
    bad = rep["mismatch"]
    au.checks.append(Check("数式の照合", not bad, f"数式を計算し直し、保存された結果と照らしました（違う {bad}・確かめられない {rep['unsupported']}）",
                           agent=bool(rep["unsupported"])))
    if doc.get("unread"):
        au.checks.append(Check("読んでいない部分", True, "読んでいない部分を窓と段階に書きました", doc["unread"], agent=True))
    parts = window.markdown_parts(doc)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000, f"Markdown {sum(len(t) for _, t in parts):,} 文字（AIが1回で読める目安 15万文字以内）"))
    au.checks.append(Check("書き込み", True, "すでに値のあるセルの上書き（レコードの大きさの変わらないもの）だけ書けます。空のセルには書けません"
                           "（.xlsx で保存し直してもらえば書けます）", agent=True))
    return au
