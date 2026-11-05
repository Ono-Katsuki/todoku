"""Excel に書いた後の読み戻し：書いた値と、書いた後のブックのセルを、同じ形にそろえて比べる。

- 日付：書いた字（「2026-11-03」「令和8年11月3日」など）を日付にし、セルのシリアル値を、そのブックの日付の数え方
  （1900 年か 1904 年か：workbookPr date1904・DATEMODE）で日付に戻して比べる
- 数：書いた字を Excel と同じ規則で数にし（全角の数字・桁区切り・%）、セルの数と比べる
本人に伝える値は、セルの表示形式を通した見えている字（「46329」ではなく「2026/11/3」、「3.0」ではなく「3」）。
書いた値から計算し直した数式のセル（金額・合計など）も、見出しと前後の値を添えて伝える。
"""

from __future__ import annotations

from ..formula.dates import serial_to_date
from ..formula.numfmt import is_date_format
from ..formula.parser import addr, parse_cell
from ..formula.values import ExcelError, is_number
from .calc import values_equal


def check(wb_after, written: dict) -> dict[str, tuple[bool, str]]:
    """書いた所の {id: (読み戻しが合うか, 本人に伝える見えている字)}。セルとして読めない id（部品の名前・文字の中の空欄の
    リストなど）は入れない（呼ぶ側の、ほかの比べ方に任せる）。"""
    from .edit import interpret_input, parse_date_text, resolve_target
    from .window import display

    out: dict[str, tuple[bool, str]] = {}
    for key, v in written.items():
        if key.startswith("@") or isinstance(v, (list, tuple, dict)):
            continue
        try:
            sn, r, c = resolve_target(wb_after, key)
            sh = wb_after.sheet(sn)
        except Exception:  # noqa: BLE001  セルでない指定
            continue
        cell = sh.cells.get((r, c))
        got = cell.value if cell is not None else None
        shown = display(wb_after, cell) if cell is not None else ""
        try:
            expected, _notes = interpret_input(wb_after, sn, r, c, v)
        except Exception:  # noqa: BLE001
            continue
        ok = values_equal(expected, got) if not (expected is None and got is None) else True
        if not ok and isinstance(got, str) and isinstance(v, str) and got == v:
            ok = True  # 字として書いたセル（字のあるセルの一部を書き換えた所）：表示形式が日付でも、セルの字は書いた字そのもの
        if cell is not None and cell.schema_error:
            ok = False  # 値が一つに決まらない（<v> が二つなど）。どの値を読んでも、見える値の保証にならない
        fmt = wb_after.style(cell.style).num_fmt if cell is not None else "General"
        if ok and isinstance(v, str) and is_number(got) and not isinstance(got, bool) and is_date_format(fmt):
            # 日付は、セルの値をブックの数え方で日付に戻して、書いた日付と同じかを別に確かめる
            d = parse_date_text(v)
            if d is not None:
                try:
                    ok = serial_to_date(float(got), wb_after.date1904) == d
                except ValueError:
                    ok = False
        out[key] = (ok, shown)
    return out


def recalculated(wb_before, wb_after, cells) -> list[dict]:
    """書いた値から計算し直して、値が変わった数式のセル：{id, label, before, after}（値は表示形式を通した字）。
    cells：計算し直したセルの (シート, 番地, 新しい値) の並び。"""
    from .window import Layout, display

    layouts: dict = {}
    out = []
    for sheet, address, new in cells:
        try:
            r, c = parse_cell(address)
            sh_a = wb_after.sheet(sheet)
            sh_b = wb_before.sheet(sheet)
        except Exception:  # noqa: BLE001
            continue
        cb, ca = sh_b.cells.get((r, c)), sh_a.cells.get((r, c))
        before = display(wb_before, cb) if cb is not None and cb.value is not None else None
        after = display(wb_after, ca) if ca is not None else _plain(new)
        if before == after:
            continue
        lay = layouts.get(sheet)
        if lay is None:
            try:
                lay = layouts[sheet] = Layout(wb_after, sh_a)
            except Exception:  # noqa: BLE001
                lay = layouts[sheet] = False
        label = None
        if lay:
            box = next((b for b in lay.boxes(r) if b.kind == "formula" and b.c1 == c), None)
            if box is not None:
                try:
                    label = lay.describe(box).get("label")
                except Exception:  # noqa: BLE001
                    label = None
        out.append({"id": f"{sheet}!{addr(r, c)}", "label": label or "", "before": before, "after": after})
    return out


def _plain(v) -> str:
    if v is None:
        return ""
    if isinstance(v, ExcelError):
        return v.code
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def say(items: list[dict], limit: int = 12) -> str:
    """本人に伝える1行：「再計算で変わった所：金額（申請!B8）2,000 → 5,000、…」。"""
    parts = []
    for it in items[:limit]:
        name = f"{it['label']}（{it['id']}）" if it["label"] else it["id"]
        before = it["before"] if it["before"] not in (None, "") else "（保存された計算結果なし）"
        parts.append(f"{name}：{before} → {it['after'] if it['after'] != '' else '空'}")
    more = f"ほか {len(items) - limit} か所" if len(items) > limit else ""
    return "再計算で変わった所（書いた値から計算し直したセル）：" + "、".join(parts) + (f"、{more}" if more else "")
