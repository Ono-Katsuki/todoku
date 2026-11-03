"""Excel の表を、見出しの列で行ごとに取り出す（エージェントが表を読むたびに処理を書かなくてよいように）。

- 見出し：エージェントが渡す見出しの字（header）がある行から始め、数の入っていない行が続くあいだを見出しの行とする
  （header_rows で行数を決めてもよい）。列の名前は、見出しの行を上から下へ「／」でつなぐ（「サービスコード／種類」）。
  結合した見出しの下にいくつもの列があるときは、列の文字を添えて分ける（「算定項目（D）」）。
- 値：Excel の画面に出る字（表示形式どおり）。結合したセルは、範囲のすべてのセルに同じ値を入れる（画面の見た目どおり）。
  値ごとに、元のセルの番地を添える（結合の左上）。
- 見出しの行より下の、字の入った行をすべて返す（表の後ろの注などもそのまま。どこまでが表かはエージェントが決める）。
"""

from __future__ import annotations

import unicodedata

from ..formula.parser import addr, num_to_col
from ..formula.values import ExcelError, is_number


class TableError(ValueError):
    pass


def _norm(s: str) -> str:
    return unicodedata.normalize("NFKC", s).replace(" ", "").replace("　", "").replace("\n", "")


def _grid(wb, sh):
    """(行, 列) → (画面の字, 元のセルの番地)。結合は範囲に広げる。"""
    from .window import display

    g = {}
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
            for rr in range(m[0], m[2] + 1):
                for cc in range(m[1], m[3] + 1):
                    g[(rr, cc)] = (text, src)
        else:
            g[(r, c)] = (text, src)
    return g


def _is_number_text(s: str) -> bool:
    t = _norm(s).replace(",", "").replace("%", "").replace("¥", "").replace("円", "")
    try:
        float(t)
        return True
    except ValueError:
        return False


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
        hits = sorted((r, c) for (r, c), (t, _s) in g.items() if want in _norm(t))
        if not hits:
            continue
        top = hits[0][0]
        # 見出しの行：header のある行から、数の入っていない行が続くあいだ（header_rows があればその行数）
        band = [top]
        if header_rows:
            band = list(range(top, top + header_rows))
        else:
            r = top + 1
            while r <= top + 3:
                cells = [t for (rr, _c), (t, _s) in g.items() if rr == r]
                if not cells or any(_is_number_text(t) for t in cells):
                    break
                band.append(r)
                r += 1
        cols = sorted({c for (r, c) in g if r in band})
        if not cols:
            continue
        names = {}
        for c in range(cols[0], cols[-1] + 1):
            parts = []
            for r in band:
                t = g.get((r, c))
                if t and _norm(t[0]) and t[0].strip() not in parts:
                    parts.append(t[0].strip().replace("\n", ""))
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
        read_sheets.append({"sheet": sh.name, "header_rows": [band[0], band[-1]], "columns": [names[c] for c in sorted(names)]})
        for n in (names[c] for c in sorted(names)):
            if n not in columns:
                columns.append(n)
        max_row = max((r for (r, _c) in g), default=0)
        for r in range(band[-1] + 1, max_row + 1):
            rec = {}
            src = {}
            for c in range(lo, hi + 1):
                if c in names and (r, c) in g:
                    rec[names[c]] = g[(r, c)][0]
                    src[names[c]] = g[(r, c)][1]
            if rec:
                rows.append({"_sheet": sh.name, "_row": r, **rec, "_cells": src})
    if not read_sheets:
        raise TableError(f"「{header}」のある表が見つかりません")
    return {"columns": columns, "rows": rows, "sheets": read_sheets}
