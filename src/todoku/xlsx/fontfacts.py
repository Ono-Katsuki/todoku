"""字の書式の事実：取り消し線・字の色（シートのふつうの色と違うもの）・字ごとの書式（リッチテキスト）・ふりがな。

取り消した字や赤い字は、値だけを読む道具からは見えない。窓には「事実」としてだけ出し、意味（無効・注意など）は決めない。
字の色の「ふつう」は、シートの値のあるマスでいちばん多い色（自動の色は黒として数える）。どの形式でも Style の字の色と取り消し線、
字ごとの書式・ふりがなは読めた形式の分だけを使う。ふつうの書類（どのマスも同じ色で、取り消し線がない）には、何も足さない。"""

from __future__ import annotations

import collections
import difflib

AUTO = "000000"


def _rgb(c: str | None) -> str:
    return (c or AUTO).upper()


def normal_color(wb, sh) -> str:
    """シートのふつうの字の色：値のあるマスでいちばん多い色（同じ数なら、ブックの既定の字の色に近い順で決まるよう、既定の色を先に数える）。"""
    n: collections.Counter = collections.Counter()
    n[_rgb(wb.style(0).font_color)] += 0
    for cell in sh.cells.values():
        if cell.value not in (None, "") or cell.has_formula:
            n[_rgb(wb.style(cell.style).font_color)] += 1
    return n.most_common(1)[0][0]


def run_props(wb, cell) -> list[tuple[str, bool, str | None]] | None:
    """字ごとの書式 [(字, 取り消し線, 色（セルの書式のままなら None）)]。ない・読めないなら None。"""
    sst = getattr(wb, "cell_sst", {}).get((cell.sheet, cell.row, cell.col))
    runs_of = getattr(wb, "sst_font_runs", None)
    if sst is None or runs_of is None:
        return _xls_runs(wb, cell)
    return runs_of(sst)


def _xls_runs(wb, cell):
    """.xls の字ごとの書式：SST の (字の位置, 字体の番号) と FONT の表から。"""
    rc = getattr(wb, "rich_cells", None)
    if not rc or not isinstance(cell.value, str):
        return None
    i = rc.get((cell.sheet, cell.row, cell.col))
    runs = getattr(wb, "sst_runs", {}).get(i) if i is not None else None
    fonts, strikes = getattr(wb, "_fonts", None), getattr(wb, "_font_strike", None)
    if not runs or fonts is None or strikes is None:
        return None
    t = cell.value
    pts = sorted(runs)
    out = [(t[: pts[0][0]], None, None)] if pts[0][0] > 0 else []
    for k, (ich, ifnt) in enumerate(pts):
        end = pts[k + 1][0] if k + 1 < len(pts) else len(t)
        fnt = ifnt - 1 if ifnt >= 4 else ifnt  # 字体の 4 番は欠番
        st = strikes[fnt] if 0 <= fnt < len(strikes) else None
        col = wb._color(fonts[fnt]) if 0 <= fnt < len(fonts) else None
        out.append((t[ich:end], st, col))
    return out


def marks(wb, cell, normal: str) -> tuple[str, str]:
    """(字の前に置く印, 字の後ろに置く印)。字そのものは変えない。"""
    if cell is None or (cell.value in (None, "") and not cell.has_formula):
        return "", ""
    st = wb.style(cell.style)
    whole_strike = bool(getattr(st, "font_strike", False))
    whole_color = _rgb(st.font_color)
    runs = run_props(wb, cell)
    tail = ""
    if runs:
        shown = [(t, whole_strike if s is None else s, whole_color if c is None else _rgb(c)) for t, s, c in runs if t.strip()]
        if shown and all(x[1] for x in shown):
            whole_strike = True
        if shown and len({x[2] for x in shown}) == 1:
            whole_color = shown[0][2]
        parts = []
        for t, s, c in shown:
            what = []
            if s and not whole_strike:
                what.append("取り消し線")
            if c != whole_color and c != normal:
                what.append(f"字の色 #{c}")
            if what:
                parts.append(f"〔うち{'・'.join(what)}：「{t.replace(chr(10), ' ').strip()}」〕")
        tail = "".join(parts)
    head = []
    if whole_strike:
        head.append("取り消し線")
    if whole_color != normal:
        head.append(f"字の色 #{whole_color}")
    return (f"〔{'・'.join(head)}〕" if head else ""), tail


def reading_parts(base: str, merged: str) -> list[str]:
    """ふりがなを当てた読み（merged）の、元の字（base）と違う所。"""
    out = []
    sm = difflib.SequenceMatcher(None, base, merged, autojunk=False)
    for op, _a, _b, j1, j2 in sm.get_opcodes():
        if op in ("replace", "insert") and merged[j1:j2].strip():
            out.append(merged[j1:j2].strip())
    return out


def _kana(t: str) -> str:
    """ひらがなをかたかなにそろえる。"""
    return "".join(chr(ord(c) + 0x60) if "ぁ" <= c <= "ゖ" else c for c in t)


def furigana(wb, sh) -> list[dict]:
    """ファイルに保存されたふりがな（同じ字と読みは 1 つにまとめる）：[{text, reading, cells, shown}]。"""
    from ..formula.parser import addr

    phon = getattr(wb, "phonetic", None)
    cs = getattr(wb, "cell_sst", None)
    if not phon or not cs:
        return []
    shown = getattr(sh, "phonetic_shown", ()) or ()
    got: dict = {}
    for (r, c), cell in sorted(sh.cells.items()):
        i = cs.get((sh.name, r, c))
        if i is None or i not in phon or not isinstance(cell.value, str):
            continue
        merged = phon.get(i)
        if not merged or _kana(merged) == _kana(cell.value):  # 字と同じ読み（かなの字のかたかな）は、足すことがない
            continue
        k = (cell.value, merged)
        d = got.get(k)
        if d is None:
            d = got[k] = {"text": cell.value, "reading": reading_parts(cell.value, merged), "cells": [], "shown": False}
        d["cells"].append(addr(r, c))
        d["shown"] = d["shown"] or (r, c) in shown
    return [d for d in got.values() if d["reading"]]
