"""書いたセルの見え方の確かめ（.xlsx・.xls）：書式で見えない・消したように見える字と、セルに収まらず切れて見える字。

- written_format：書いたセルの書式（表示形式・字の色・塗り・取り消し線・条件付き書式・字ごとの書式・非表示の行と列）。
  見えない字になるもの（表示形式「;;;」など字が出ない形・字と塗りが同じ色・書いた値で条件付き書式が字を隠す）は、書かない所の説明に、
  取り消し線・非表示の行などは、伝える書式に入れる
- fit_issues：列の幅・行の高さ・字の大きさ（sheetpic が読む書式）と、字の幅（textwidth：全角は 2、ほかは 1）から、書いた字が
  セルに収まるかを見積もる。隣のセルの字で切れる・結合したセルからはみ出して切れる・折り返した行が行の高さに入らない・数が「####」
  になる所を伝える（見積もりなので、言い方は「見込み」）
- both_controls_checked：書いたあと、同じ行のチェックボックスのすべてに印がついている（事実だけを伝える。どれを選ぶ欄かは決めない）

列の幅は ECMA-376 §18.3.1.13（col の width：既定の書体の数字の幅 × 字の数 ＋ 左右の余白 5 ピクセル）でピクセルにし、字の幅
（textwidth：全角は 2、ほかは 1。数字の幅を 1 とした数）と比べる。幅のない列は defaultColWidth、それもなければ baseColWidth
（§18.3.1.81）。行の高さの既定は sheetFormatPr の defaultRowHeight。
"""

from __future__ import annotations

import math

from ..formula.parser import addr

_PAD_PX = 5  # 列の幅の左右の余白（ECMA-376 §18.3.1.13 の「5 pixel padding」：左右 2 ピクセルずつと格子の線 1 ピクセル）


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
        elif cell_font and font != "000000":
            # セルの書式は変えないので、書いた字はこの色になる（元の様式の色。黒にしたいなら様式の作り手か本人が決める）
            items.append(f"字の色（セルの書式の色 #{font}。書いた字もこの色で表示・印刷されます）")
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
    """同じシートの同じ行に、チェックボックスが 2 つ以上あり、書いた後にそのすべてに印がつく（今回 1 つ以上に印をつけた）：その事実。
    どれか 1 つを選ぶ欄かどうかは決めない（伝えるだけ）。"""
    out = []
    by_row: dict = {}
    for (sheet, _sid), (_s, ctl, b) in controls.items():
        if b and getattr(ctl, "anchor", None):
            by_row.setdefault((sheet, ctl.anchor[0]), []).append(ctl)
    for (sheet, row), mine in by_row.items():
        sh = wb.sheet(sheet)
        row_boxes = [c for c in sh.controls if c.kind == "CheckBox" and getattr(c, "anchor", None) and c.anchor[0] == row]
        if len(row_boxes) < 2:
            continue
        on = {c.shape_id for c in mine} | {c.shape_id for c in row_boxes if c.is_checked}
        off = {sid for (s_, sid), (_s, _c, b) in controls.items() if s_ == sheet and not b}
        if {c.shape_id for c in row_boxes} <= (on - off):
            names = "」「".join((c.text or c.name) if getattr(c, "text", None) else c.name for c in row_boxes)
            n = len(row_boxes)
            out.append(f"{sheet}: 同じ行（{row} 行目）の {n} つのチェックボックス「{names}」の{'両方' if n == 2 else 'すべて'}に印がつきます"
                       "（書いたあとの状態）")
    return out


def _formats(wb):
    """ブックの書式の一覧（1 冊に 1 回だけ読む。欄ごとに部品を展開し直すと、書く欄の多い書類で展開の合計の上限に当たる）。"""
    from ..sheetpic import _xls_formats, _xlsx_formats

    got = getattr(wb, "_cellcheck_formats", None)
    if got is None:
        got = (_xls_formats(wb) if getattr(wb, "kind", "xlsx") == "xls" else _xlsx_formats(wb))[0]
        wb._cellcheck_formats = got
    return got


def _mdw(wb) -> int:
    """既定の書体（書式の 0 番）の数字の幅（ピクセル）。"""
    from .circle import mdw_for

    got = getattr(wb, "_cellcheck_mdw", None)
    if got is None:
        fmts = _formats(wb)
        base = fmts[0].font if fmts and fmts[0].font else None
        got = mdw_for(base.face if base else "", base.size if base else 11.0)[0]
        wb._cellcheck_mdw = got
    return got


def _col_w(sh, c: int, mdw: int = 7) -> float:
    """列の、字を書ける幅（数字の幅を 1 とした字の数。左右の余白 _PAD_PX を除く）。"""
    from .circle import width_px

    if c in sh.hidden_cols:
        return 0.0
    px = width_px(sh.col_widths.get(c, sh.default_col_width), mdw, getattr(sh, "base_col_width", 8) or 8)
    return px / mdw


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


def _shown_before(wb, sh, r: int, c: int) -> str:
    """書く前のセルの見えている字（表示形式を通した字）。"""
    from ..formula.numfmt import format_value

    cell = sh.cells.get((r, c))
    if cell is None or cell.value in (None, "") or isinstance(cell.value, bool):
        return ""
    try:
        return format_value(cell.value, wb.style(sh.style_of(r, c)).num_fmt or "General", getattr(wb, "date1904", False))
    except Exception:  # noqa: BLE001
        return str(cell.value)


def _wrap_by_write(wb, sh, rng: tuple, before: str, after: str, size: float) -> str | None:
    """折り返して表示するセルで、書いたことで行が増える見込みなら、その警告（所は「{place}」）。列の幅はピクセルで
    （既定の書体の数字の幅 × 保存された字の数。xlsx.circle と同じ決め方）、セルの左右の余白は約 2 ピクセルずつ。"""
    from ..wrapcheck import say, wraps
    from .circle import mdw_for, width_px

    fmts = _formats(wb)
    base = fmts[0].font if fmts and fmts[0].font else None
    mdw, _known = mdw_for(base.face if base else "", base.size if base else 11.0)
    r1, c1, r2, c2 = rng
    px = sum(0 if c in sh.hidden_cols else width_px(sh.col_widths.get(c, sh.default_col_width), mdw,
                                                   getattr(sh, "base_col_width", 8) or 8) for c in range(c1, c2 + 1))
    avail = max(px - 5, 1) * 0.75
    got = wraps(before, after, avail, size, size)
    if got is None:
        return None
    merged = f"（結合したセル {addr(r1, c1)}:{addr(r2, c2)}）" if (r1, c1) != (r2, c2) else ""
    return say("{place}" + merged, after, got[0], got[1], f"列の幅 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイント", got[2])


def _hashes(v, shown: str, num_fmt: str | None, scale: float, avail: float) -> bool:
    """幅に入らない数が「####」になるか。「標準」以外の形式は、そのまま「####」。「標準」は、小数の桁を減らして入れようと
    するので、整数の部分（符号を含む）も入らないときだけ。"""
    from ..textwidth import text_cells

    if (num_fmt or "General").strip().lower() not in ("general", ""):
        return True
    try:
        f = float(v)
    except (TypeError, ValueError):
        return True
    whole = str(int(round(f))) if abs(f) < 1e11 else shown
    return text_cells(whole) * scale > avail


def fit_issues(wb, plan) -> list[dict]:
    """書いた字がセルに収まらず、切れて見える・「####」になる見込みの所：[{id, cell, message}]。"""
    from ..formula.numfmt import format_value
    from ..textwidth import text_cells

    out: list[dict] = []
    if not plan.direct:
        return out
    fmts = _formats(wb)
    base = fmts[0].font.size if fmts and fmts[0].font else 11.0
    mdw = _mdw(wb)
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
        width = sum(_col_w(sh, c, mdw) for c in range(c1, c2 + 1))
        avail = max(width - _PAD_PX / mdw, 0.5)
        lines = shown.split("\n")
        widths = [text_cells(ln) * scale for ln in lines]
        need = max(widths) if widths else 0.0
        name = key if not key.startswith("@") else where
        if not isinstance(v, str) and not cf.shrink:
            # 数は折り返さない（折り返して表示するセルでも）。幅に入らない数を、Excel は「####」と表示する。
            # 「標準」の形式では小数の桁を減らして入れようとするので、整数の部分（と符号）が入らないときだけ
            if need > avail and _hashes(v, shown, st.num_fmt, scale, avail):
                out.append({"id": name, "cell": where, "message":
                            f"{where}: 数「{shown}」はセルの幅（約 {avail:.1f} 字分）に入らず、Excel では「####」と表示される"
                            "見込みです（数は折り返しません。列の幅からの見積もり。本人に確かめてください）"})
            continue
        if cf.wrap:
            n = sum(max(1, math.ceil(w / avail)) for w in widths)
            height = sum(_row_h(sh, r) for r in range(r1, r2 + 1))
            line_h = size * 1.2
            # 書く前の字（同じセル）：書く前から切れて見える・折り返していたなら、書いたことで変わったのではない
            before = _shown_before(wb, sh, r_, c_)
            n0 = sum(max(1, math.ceil(text_cells(ln) * scale / avail)) for ln in before.split("\n")) if before else 0
            if n * line_h > height + 1:
                out.append({"id": name, "cell": where, "message":
                            f"{where}: 折り返して {n} 行になる見込みですが、行の高さ（{height:g} ポイント）には約 {max(1, int(height // line_h))} 行"
                            "しか入らず、下の行が切れて見える見込みです（列の幅・行の高さからの見積もり。字を短くするか、本人に確かめてください）"
                            + ("（書く前の字でも、同じように切れて見える見込みでした）" if before and n0 * line_h > height + 1 else "")})
                continue
            got = _wrap_by_write(wb, sh, (r1, c1, r2, c2), before, shown, size)
            if got:
                out.append({"id": name, "cell": where, "wrap": True, "message": got.replace("{place}", where)})
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
                room += _col_w(sh, c, mdw)
                c += d
            if blocked:
                break
        if blocked:
            out.append({"id": name, "cell": where, "message":
                        f"{where}: 書いた字「{shown[:20]}{'…' if len(shown) > 20 else ''}」はセルの幅に収まらず、隣の {blocked} の字で"
                        "切れて見える見込みです（列の幅からの見積もり。字を短くするか、本人に確かめてください）"})
    return out
