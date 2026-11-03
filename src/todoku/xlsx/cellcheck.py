"""書いたセルの見え方の確かめ（.xlsx・.xls）：書式で見えない・消したように見える字と、セルに収まらず切れて見える字。

- written_format：書いたセルの書式（表示形式・字の色・塗り・取り消し線・条件付き書式・字ごとの書式・非表示の行と列）。
  見えない字になるもの（表示形式「;;;」など字が出ない形・字と塗りが同じ色・書いた値で条件付き書式が字を隠す）は、書かない所の説明に、
  取り消し線・非表示の行などは、伝える書式に入れる
- fit_issues：列の幅・行の高さ・字の大きさ（sheetpic が読む書式）と、字の幅（textwidth：全角は 2、ほかは 1）から、書いた字が
  セルに収まるかを見積もる。隣のセルの字で切れる・結合したセルからはみ出して切れる・折り返した行が行の高さに入らない・数が「####」
  になる所を伝える（見積もりなので、言い方は「見込み」）
- both_controls_checked：同じ行にちょうど2つあるチェックボックスの両方に印をつけた（どちらか1つを選ぶ欄のことが多い）
"""

from __future__ import annotations

import math

from ..formula.parser import addr

_MARGIN = 0.5  # セルの左右の余白（字の数。Excel の列の幅の 5 ピクセル分の余白の見積もり）


def _rgb(c: str | None, default: str) -> str:
    return (c or default).upper()[-6:]


def hidden_by_style(wb, sheet: str, r: int, c: int, value) -> str | None:
    """セルの書式（表示形式・字の色と塗りの色）だけで、value が見えない字になるなら、その説明（字ごとの書式・条件付き書式は見ない）。"""
    from ..formula.numfmt import format_value

    if value is None or value == "":
        return None
    st = wb.style(wb.sheet(sheet).style_of(r, c))
    fmt = st.num_fmt or "General"
    try:
        if not isinstance(value, bool) and format_value(value, fmt, getattr(wb, "date1904", False)) == "":
            return f"表示形式「{fmt}」で見えない"
    except Exception:  # noqa: BLE001
        pass
    font, fill = _rgb(st.font_color, "000000"), _rgb(st.fill_color, "FFFFFF")
    if font == fill:
        return f"字の色と塗りの色が同じ（#{font}）"
    return None


def written_format(wb, plan) -> tuple[list[dict], list[str]]:
    """(伝える書式 [{id, formatting}], 見えない字になるセルの説明)。部品につながったセル（書式で隠すのが決まり）は見ない。"""
    from ..formula.numfmt import format_value
    from .calc import Calc
    from .richtext import describe_props

    found: list[dict] = []
    invisible: list[str] = []
    looks = None
    for cell_, key in plan.direct.items():
        v = plan.targets.get(cell_)
        if v is None or v == "":
            continue
        s_, r_, c_ = cell_
        sh = wb.sheet(s_)
        where = f"{s_}!{addr(r_, c_)}"
        st = wb.style(sh.style_of(r_, c_))
        items: list[str] = []
        fmt = st.num_fmt or "General"
        try:
            shown = format_value(v, fmt, getattr(wb, "date1904", False)) if not isinstance(v, bool) else "x"
        except Exception:  # noqa: BLE001
            shown = None
        if shown == "":
            invisible.append(f"{where}: 表示形式「{fmt}」では、書いた値「{v}」が表示されません")
            items.append(f"表示形式「{fmt}」で見えない")
        font, fill = _rgb(st.font_color, "000000"), _rgb(st.fill_color, "FFFFFF")
        rich = plan.rich.get(cell_)
        # 字ごとの書式のあるセル：書いた字が入る書式（<rPr>・字体の番号）が、その字の書式。書式のない所の字だけがセルの字体
        cell_font = rich is None or any(p in (None, b"") for p in rich[3]) or not rich[3]
        if cell_font and font == fill:
            invisible.append(f"{where}: 字の色と塗りの色が同じ（#{font}）で、書いた字が見えません")
            items.append(f"字の色と塗りの色が同じ（#{font}）")
        elif cell_font and font == "FFFFFF":
            items.append("白い字")
        if cell_font and getattr(st, "font_strike", False):
            items.append("取り消し線")
        if rich is not None:
            for rpr in rich[3]:
                d = _xls_font_props(wb, rpr) if rich[0] == "xls" else describe_props(rpr)
                if "白い字" in d and fill == "FFFFFF":
                    invisible.append(f"{where}: 書いた字が入る所の字（字ごとの書式）が白い字で、見えません")
                items += [f"書いた字の書式：{x}" for x in d]
        if sh.cond_rules_at(r_, c_) if hasattr(sh, "cond_rules_at") else False:
            try:
                if looks is None:
                    from .cf import Looks

                    looks = Looks(wb, Calc(wb, plan.targets))
                lk = looks.look(sh, r_, c_)
                if lk.hidden and font != fill and shown != "":
                    invisible.append(f"{where}: 書いた値で条件付き書式がかかり、字が見えなくなります")
                    items.append("条件付き書式で見えない")
                elif lk.unknown:
                    items.append("条件付き書式（書いた後の見え方を確かめられない規則）")
                elif lk.highlighted:
                    items.append("条件付き書式で色が付く")
            except Exception:  # noqa: BLE001
                items.append("条件付き書式（書いた後の見え方を確かめられない）")
        if sh.is_hidden(r_, c_):
            items.append("非表示の行か列（印刷されません）")
        if items:
            found.append({"id": key if not key.startswith("@") else where, "cell": where, "formatting": list(dict.fromkeys(items))})
    return found, invisible


def _xls_font_props(wb, ifnt) -> list[str]:
    """.xls の字体の番号（字ごとの書式。None はセルの書式の字体）の、取り消し線・字の色。"""
    if ifnt is None:
        return []
    fnt = ifnt - 1 if ifnt >= 4 else ifnt  # 字体の4番は欠番
    out = []
    strikes = getattr(wb, "_font_strike", [])
    if 0 <= fnt < len(strikes) and strikes[fnt]:
        out.append("取り消し線")
    fonts = getattr(wb, "_fonts", [])
    if 0 <= fnt < len(fonts):
        rgb = wb._color(fonts[fnt])
        if rgb == "FFFFFF":
            out.append("白い字")
        elif rgb not in (None, "000000"):
            out.append(f"字の色（#{rgb}）")
    return out


def both_controls_checked(wb, controls: dict) -> list[str]:
    """同じシートの同じ行に、チェックボックスがちょうど2つあり、書いた後に両方に印がつく（今回1つ以上に印をつけた）。"""
    out = []
    by_row: dict = {}
    for (sheet, _sid), (_s, ctl, b) in controls.items():
        if b and getattr(ctl, "anchor", None):
            by_row.setdefault((sheet, ctl.anchor[0]), []).append(ctl)
    for (sheet, row), mine in by_row.items():
        sh = wb.sheet(sheet)
        row_boxes = [c for c in sh.controls if c.kind == "CheckBox" and getattr(c, "anchor", None) and c.anchor[0] == row]
        if len(row_boxes) != 2:
            continue
        on = {c.shape_id for c in mine} | {c.shape_id for c in row_boxes if c.is_checked}
        off = {sid for (s_, sid), (_s, _c, b) in controls.items() if s_ == sheet and not b}
        if {c.shape_id for c in row_boxes} <= (on - off):
            names = "」と「".join((c.text or c.name) if getattr(c, "text", None) else c.name for c in row_boxes)
            out.append(f"{sheet}: 同じ行の2つのチェックボックス「{names}」の両方に印がつきます。どちらか1つを選ぶ欄なら、本人に確かめて"
                       "1つだけにしてください")
    return out


def _formats(wb):
    from ..sheetpic import _xls_formats, _xlsx_formats

    return (_xls_formats(wb) if getattr(wb, "kind", "xlsx") == "xls" else _xlsx_formats(wb))[0]


def _col_w(sh, c: int) -> float:
    if c in sh.hidden_cols:
        return 0.0
    w = sh.col_widths.get(c)
    if w is None:
        w = sh.default_col_width if getattr(sh, "default_col_width", None) else 8.43
    return float(w)


def _row_h(sh, r: int) -> float:
    if r in sh.hidden_rows:
        return 0.0
    return float(sh.row_heights.get(r, sh.default_row_height or 15.0))


def _empty(sh, r: int, c: int, targets: dict) -> bool:
    if (sh.name, r, c) in targets:
        v = targets[(sh.name, r, c)]
        return v is None or v == ""
    if sh.merged_at(r, c) is not None:
        return False
    cell = sh.cells.get((r, c))
    return cell is None or cell.value is None or cell.value == ""


def fit_issues(wb, plan) -> list[dict]:
    """書いた字がセルに収まらず、切れて見える・「####」になる見込みの所：[{id, cell, message}]。"""
    from ..formula.numfmt import format_value
    from ..textwidth import text_cells

    out: list[dict] = []
    if not plan.direct:
        return out
    fmts = _formats(wb)
    base = fmts[0].font.size if fmts and fmts[0].font else 11.0
    for cell_, key in plan.direct.items():
        v = plan.targets.get(cell_)
        if v is None or v == "" or isinstance(v, bool):
            continue
        s_, r_, c_ = cell_
        sh = wb.sheet(s_)
        if sh.is_hidden(r_, c_):
            continue
        where = f"{s_}!{addr(r_, c_)}"
        xf = sh.style_of(r_, c_)
        cf = fmts[xf] if 0 <= xf < len(fmts) else None
        if cf is None or cf.rotation or cf.halign in ("fill", "justify", "distributed", "centerAcross"):
            continue
        st = wb.style(xf)
        try:
            shown = format_value(v, st.num_fmt or "General", getattr(wb, "date1904", False))
        except Exception:  # noqa: BLE001
            continue
        size = cf.font.size if cf.font else base
        scale = size / base if base else 1.0
        m = sh.merged_at(r_, c_)
        r1, c1, r2, c2 = m if m else (r_, c_, r_, c_)
        width = sum(_col_w(sh, c) for c in range(c1, c2 + 1))
        avail = max(width - _MARGIN, 0.5)
        lines = shown.split("\n")
        widths = [text_cells(ln) * scale for ln in lines]
        need = max(widths) if widths else 0.0
        name = key if not key.startswith("@") else where
        if cf.wrap:
            n = sum(max(1, math.ceil(w / avail)) for w in widths)
            height = sum(_row_h(sh, r) for r in range(r1, r2 + 1))
            line_h = size * 1.2
            if n * line_h > height + 1:
                out.append({"id": name, "cell": where, "message":
                            f"{where}: 折り返して {n} 行になる見込みですが、行の高さ（{height:g} ポイント）には約 {max(1, int(height // line_h))} 行"
                            "しか入らず、下の行が切れて見える見込みです（列の幅・行の高さからの見積もり。字を短くするか、本人に確かめてください）"})
            continue
        if len(lines) > 1:
            out.append({"id": name, "cell": where, "message":
                        f"{where}: 書いた字に改行がありますが、このセルは折り返して表示する設定ではないので、Excel では改行が表示されず"
                        "1行に続けて見えます"})
        if need <= avail:
            continue
        if cf.shrink:
            out.append({"id": name, "cell": where, "message":
                        f"{where}: セルの幅に合わせて字が縮んで表示されます（元の大きさの約 {int(100 * avail / need)}%。小さすぎて"
                        "読みにくくならないか、本人に確かめてください）"})
            continue
        if not isinstance(v, str):
            if (st.num_fmt or "General").strip().lower() not in ("general", ""):
                out.append({"id": name, "cell": where, "message":
                            f"{where}: 「{shown}」はセルの幅に入らず、Excel では「####」と表示される見込みです（列の幅からの見積もり）"})
            continue
        if m:
            out.append({"id": name, "cell": where, "message":
                        f"{where}: 書いた字「{shown[:20]}{'…' if len(shown) > 20 else ''}」は、結合したセルの幅に収まらず、はみ出す所が"
                        "切れて見える見込みです（列の幅からの見積もり。字を短くするか、本人に確かめてください）"})
            continue
        h = cf.halign if cf.halign != "general" else "left"
        sides = {"left": [1], "right": [-1], "center": [1, -1]}.get(h, [1])
        over = need - avail
        blocked = None
        for d in sides:
            room = 0.0
            want = over if len(sides) == 1 else over / 2
            c = c_ + d
            while room < want and 1 <= c <= 16384:
                if not _empty(sh, r_, c, plan.targets):
                    blocked = addr(r_, c)
                    break
                room += _col_w(sh, c)
                c += d
            if blocked:
                break
        if blocked:
            out.append({"id": name, "cell": where, "message":
                        f"{where}: 書いた字「{shown[:20]}{'…' if len(shown) > 20 else ''}」はセルの幅に収まらず、隣の {blocked} の字で"
                        "切れて見える見込みです（列の幅からの見積もり。字を短くするか、本人に確かめてください）"})
    return out
