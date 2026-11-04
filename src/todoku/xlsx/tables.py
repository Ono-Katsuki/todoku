"""Excel の表を、見出しの列で行ごとに取り出す（エージェントが表を読むたびに処理を書かなくてよいように）。

- 見出し：エージェントが渡す見出しの字（header）がある行（字が同じセルを先に選ぶ）と、その行の見出しが縦に結合されている範囲。
  その下の行が見出しの続きかどうかは規則では決めず、下の数行をそのまま返す（エージェントが読んで header_rows で決める）。列の名前は、見出しの行を上から下へ「／」でつなぐ（「サービスコード／種類」）。
  結合した見出しの下にいくつもの列があるときは、列の文字を添えて分ける（「算定項目（D）」）。
- 値：Excel の画面に出る字（表示形式どおり）。結合したセルは、範囲のすべてのセルに同じ値を入れる（画面の見た目どおり）。
  値ごとに、元のセルの番地を添える（結合の左上）。
- 見出しの行より下の、字の入った行をすべて返す（表の後ろの注などもそのまま。どこまでが表かはエージェントが決める）。
"""

from __future__ import annotations

import unicodedata

from ..formula.parser import addr, num_to_col
from ..formula.values import ExcelError, is_number
from ..limits import MAX_DECLARED_CELLS


class TableError(ValueError):
    pass


def _norm(s: str) -> str:
    return unicodedata.normalize("NFKC", s).replace(" ", "").replace("　", "").replace("\n", "")


def _grid(wb, sh):
    """(行, 列) → (画面の字, 元のセルの番地)。結合は範囲に広げる。"""
    from .window import display

    g = {}
    rows_there = {r for r, _c in sh.cells}
    cols_there = {c for _r, c in sh.cells}
    for (r, c), cell in sh.cells.items():
        if cell.value is None and not cell.has_formula:
            continue
        m = sh.merged_at(r, c)
        if m and (r, c) != (m[0], m[1]):
            continue  # 結合の左上以外（画面に出ない）
        v = cell.value
        text = v.code if isinstance(v, ExcelError) else display(wb, cell)
        if text is None or (isinstance(text, str) and not text.strip() and not is_number(v)):
            continue
        text = str(text)
        src = addr(r, c)
        if m:
            if (m[2] - m[0] + 1) * (m[3] - m[1] + 1) > MAX_DECLARED_CELLS:
                # 宣言された結合が広すぎる：セルのある行と列だけに広げる（ほかの行・列に、この字のほかの中身はない）
                rs = sorted({m[0]} | {x for x in rows_there if m[0] <= x <= m[2]})
                cs = sorted({m[1]} | {x for x in cols_there if m[1] <= x <= m[3]})
            else:
                rs, cs = range(m[0], m[2] + 1), range(m[1], m[3] + 1)
            for rr in rs:
                for cc in cs:
                    g[(rr, cc)] = (text, src)
        else:
            g[(r, c)] = (text, src)
    return g


def read(wb, header: str, sheet: str | None = None, header_rows: int | None = None) -> dict:
    """{columns: 列の名前, rows: [{_sheet, _row, 列の名前: 値, …}], sources: 値ごとの元のセル, sheets: 読んだシート}。"""
    want = _norm(header)
    if not want:
        raise TableError("header に、表の見出しにある字を指定してください（例「合成」「サービスコード」）")
    sheets = [wb.sheet(sheet)] if sheet else list(wb.sheets)
    if sheet and sheets[0] is None:
        raise TableError(f"シート「{sheet}」がありません")
    columns: list[str] = []
    rows: list[dict] = []
    read_sheets = []
    for sh in sheets:
        if sh is None or getattr(sh, "state", "").startswith("non-worksheet"):
            continue
        g = _grid(wb, sh)
        # 字が見出しと同じセルを先に選ぶ（表題「…サービスコード表」の中の字に引っかからないように）。なければ、字を含むセル
        hits = sorted((r, c) for (r, c), (t, _s) in g.items() if _norm(t) == want) or \
            sorted((r, c) for (r, c), (t, _s) in g.items() if want in _norm(t))
        if not hits:
            continue
        top, hc = hits[0]
        # 見出しの行：規則で決めるのは、構造から確実に分かることだけ（header のある行と、その行の見出しが縦に結合されている範囲）。
        # その下の行が2段目の見出しかどうかは決めない：下の行をそのまま返し、エージェントが読んで header_rows で決める
        if header_rows:
            band = list(range(top, top + header_rows))
            how = f"エージェントの指定（header_rows={header_rows}）"
        else:
            end = top
            for c in {c for (r, c) in g if r == top}:
                m = sh.merged_at(top, c)
                if m and m[0] == top:
                    end = max(end, m[2])
            band = list(range(top, end + 1)) if end - top < 1024 else sorted({r for (r, _c) in g if top <= r <= end})
            how = "見出しの字のある行" + ("（と、縦に結合した見出しの範囲）" if end > top else "")
        cols = sorted({c for (r, c) in g if r in band})
        if not cols:
            continue
        names = {}
        for c in range(cols[0], cols[-1] + 1):
            parts = []
            for r in band:
                t = g.get((r, c))
                text = t[0].strip().replace("\n", "") if t else ""
                if _norm(text) and text not in parts:  # 縦に結合した見出しは、1回だけ
                    parts.append(text)
            if parts:
                names[c] = "／".join(parts)
        # 同じ名前の列（結合した見出しの下の列）：列の文字を添える
        seen: dict[str, list[int]] = {}
        for c, n in names.items():
            seen.setdefault(n, []).append(c)
        for n, cs in seen.items():
            if len(cs) > 1:
                for c in cs:
                    names[c] = f"{n}（{num_to_col(c)}）"
        lo, hi = min(names), max(names)
        below = []
        for r in range(band[-1] + 1, band[-1] + 4):
            cells = {addr(r, c): g[(r, c)][0] for c in range(lo, hi + 1) if (r, c) in g}
            if cells:
                below.append({"row": r, "cells": cells})
        read_sheets.append({"sheet": sh.name, "header_cell": addr(top, hc), "header_rows": [band[0], band[-1]], "header_rule": how,
                            "columns": [names[c] for c in sorted(names)], "rows_below_header": below})
        for n in (names[c] for c in sorted(names)):
            if n not in columns:
                columns.append(n)
        by_row: dict[int, list[int]] = {}  # 字のある行と列だけを見る（遠くのセルが 1 つあっても、行 × 列をなめない）
        for (r, c) in g:
            if r > band[-1] and lo <= c <= hi:
                by_row.setdefault(r, []).append(c)
        for r in sorted(by_row):
            rec = {}
            src = {}
            for c in sorted(by_row[r]):
                if c in names:
                    rec[names[c]] = g[(r, c)][0]
                    src[names[c]] = g[(r, c)][1]
            if rec:
                rows.append({"_sheet": sh.name, "_row": r, **rec, "_cells": src})
    if not read_sheets:
        raise TableError(f"「{header}」のある表が見つかりません")
    return {"columns": columns, "rows": rows, "sheets": read_sheets,
            "check": "見出しの下の行（rows_below_header）が見出しの続き（2段目の見出し・単位の行など）か、データの行かは、"
                     "規則では決めていません。読んで、見出しの続きなら header_rows（見出しの行数）を指定してもう一度呼んでください。"
                     "表の後ろの注の行もそのまま返すので、どこまでが表かも読んで決めてください"}
