"""PDFの窓：ページごとに、入力欄の一覧と、上から読む順の文字（入力欄は［p1.3］）を作る。

入力欄は2通り：
- フォームの欄（AcroForm）：PDFにある欄そのもの。名前は欄の完全な名前
- 配置から見つけた所（印刷用のPDF）：空のマス・字の形の空欄・1 字ぶん以上の空き・箱の字・下線の罫線（layout.py の「書く所」）。
  名前は p{ページ}.{頭の字}{左}-{上}（作りの種類と、作りの位置を mm に丸めた数：layout.place_id）。位置だけで決まるので、ほかの所が
  増えても減っても、とどくの版が変わっても、同じ所は同じ名前。書き込むと、同じ名前のフォームの欄をその位置に足す（元のページの描画は
  1バイトも変えない）。前の版のとどくが書いた欄（名前 p1.3・p1.a12）は、その欄とだけ重なる所の名前として、そのまま使う
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from ..xlsx.window import GAP_TOKEN_NOTE, label_text, shown_gap, with_label_keys
from .content import page_content
from .doc import PdfDocument
from .form import FormField, Widget, has_xfa, read_fields
from .layout import OLD_ID, PLACE_ID, Cell, Line, PageLayout, Spot, _overlap

DET = "deterministic"
LAYOUT = "layout"

KIND_JA = {"text": "文字", "comb": "1字ずつ書く欄", "inline": "文字の中の空欄", "checkbox": "チェック", "space": "余白",
           "radio": "ひとつ選ぶ", "combo": "一覧から選ぶ", "list": "一覧から選ぶ", "signature": "署名（電子署名）",
           "button": "ボタン（押すだけ）"}


@dataclass
class _Item:
    x0: float
    y0: float
    x1: float
    y1: float
    cell: Cell | None = None
    line: Line | None = None
    marks: list = field(default_factory=list)  # (順番の x, 印の文字列を返す関数)


def _render_line(ln: Line, spots: list[Spot], name, kept: list | None = None, circled: list | None = None) -> str:
    """行の文字に、入力欄の印（［p1.3］）を入れる。kept があれば、窓に出した元の字（印に置き換えて消えた字を除いたもの）を足す
    （読み落としの検査が、窓に実際に出た字で数えるため）。circled：ページに描かれた○が囲む字 (始め, 終わり, id)：その字の後ろに
    「［○］」（とどくが描いた○は「［p1.c1：○］」）。4 つ目が真のもの（角の丸い四角など、○かもしれない囲み）は「［○？］」。"""
    text = ln.text
    edits = []
    for a, b, ident, *maybe in circled or []:
        edits.append((b, b, "［○？］" if maybe and maybe[0] else f"［{ident}：○］" if ident else "［○］"))
    for sp in sorted(spots, key=lambda x: x.at or (0, 0)):
        a, b = sp.at
        mark = _filled_mark(sp, name)
        if mark is not None:
            edits.append((a, b, mark))
            continue
        q = "？" if sp.confidence == "agent" else ""  # 規則では決められない所
        if sp.cover:
            edits.append((a, b, f"［{name(sp)}{q}＝{sp.cover}］"))  # 伏せ字：元の字を見せる
            continue
        if sp.kind == "checkbox":
            edits.append((a, b, _box_glyph(sp, text[a:b]) + f"［{name(sp)}{q}］"))
        elif sp.kind == "space":
            edits.append((a, b, f" ［{name(sp)}{q}］"))
        else:
            # 空欄の字（下線の字「＿＿＿」など）は消さずに残し、その後ろに印を置く（空白だけの空欄は印に置き換える）
            keep = text[a:b] if text[a:b].strip(" \u3000") else ""
            # 「？」の所で空白だけのものは、行の空白（測った空きから作った字）をそのまま印の中に入れる（「［p1.a1？　］」）
            gap = shown_gap(text[a:b]) if q else ""
            edits.append((a, b, f"{keep}［{name(sp)}{q}{gap}］"))
    edits.sort(key=lambda e: (e[0], e[1]))
    out, at = [], 0
    shown = []
    for a, b, rep in edits:
        if a < at:
            continue
        out.append(text[at:a])
        out.append(rep)
        shown.append(text[at:a])
        if text[a:b].strip(" \u3000") and text[a:b] in rep:
            shown.append(text[a:b])
        at = b
    out.append(text[at:])
    shown.append(text[at:])
    if kept is not None:
        kept.append("".join(shown))
    return "".join(out).strip()


def _box_glyph(sp: Spot, orig: str) -> str:
    """チェックの印の字：印刷された記入済みの印（☑・■・✓）は、元の字のまま出す（■ は箇条の印のこともあるので、☑ に変えない）。"""
    from .layout import CHECKED_GLYPHS

    if orig and orig in CHECKED_GLYPHS:
        return orig
    return "☑" if sp.checked else "☐"


def _filled_mark(sp: Spot, name) -> str | None:
    """配置から見つけた欄に、同じ名前のフォームの欄（書き込んだもの）があれば、その値を出す。"""
    hits = getattr(getattr(sp, "alias_of", sp), "widget", None)
    if not hits:
        return None
    nm = name(sp)
    for f, w in hits:
        if f.name == nm:
            if f.kind == "checkbox":
                return ("☑" if f.checked else "☐") + f"［{nm}］"
            v = f.value_text()
            return f"［{nm}＝{v}］" if v not in (None, "") else f"［{nm}］"
    return None


def _widget_mark(f: FormField, w: Widget) -> str:
    if f.kind == "checkbox":
        return ("☑" if (w.state and w.state != "Off") else "☐") + f"［{f.name}］"
    if f.kind == "radio":
        from .form import name_text

        on = w.state and w.state != "Off"
        opt = name_text(w.on or "")
        if f.options and opt.isdigit() and int(opt) < len(f.options):
            opt = f.options[int(opt)][1] or f.options[int(opt)][0]
        return ("◉" if on else "○") + f"［{f.name}＝{opt}］"
    v = f.value_text()
    if isinstance(v, list):
        v = "、".join(str(x) for x in v)
    return f"［{f.name}＝{v}］" if v not in (None, "") else f"［{f.name}］"


def _unseen_mark(f: FormField, w: Widget) -> str:
    """画面にも紙にも出ない部品・紙に印刷されない部品：窓の欄の印のすぐ後ろに「［name：画面にも紙にも出ない］」と出す。"""
    gone, unprinted = f.visibility([w])
    if gone:
        return f"［{f.name}：画面にも紙にも出ない］"
    if unprinted:
        return f"［{f.name}：紙に印刷されない］"
    return ""


def page_window(doc: PdfDocument, page: dict, number: int, fields: list[FormField], layout: bool = True) -> dict:
    """layout：配置から入力欄を見つけるか（XFA のフォームでは見つけない：XFA の欄だけが、Adobe で表示される欄）。"""
    pc = page_content(doc, page)
    lay = PageLayout(doc, page, pc, number)
    view = lay.view
    ws: list[tuple[FormField, Widget, tuple]] = []
    for f in fields:
        for w in f.widgets:
            if w.page == number - 1:
                ws.append((f, w, view.rect(*w.rect)))
    by_name = {f.name: f for f, _, _ in ws}
    # 配置から見つけた欄のうち、元からあるフォームの欄の中に収まるものは、フォームの欄を使う。重なる・すぐ隣のものは、
    # 消さずに「？」（規則では決められない）にして、その欄の名前を添える（_near_widget・_flag_near_widget）
    spots = []
    # とどくが書き込んだ欄の名前（配置の所・字を目印にした所 w・印影 s・○ c。前の版の番号の名前 p1.3・p1.a12 も）
    tdk = re.compile(rf"p{number}\.(?:[a-z]?\d+|{PLACE_ID})")
    hidden: set[int] = set()
    real_ws = [(f, w, r) for f, w, r in ws if not tdk.fullmatch(f.name)]
    _inherit_ids(lay.spots if layout else [], [(f, r) for f, w, r in ws if tdk.fullmatch(f.name)], number)
    for sp in (lay.spots if layout else []):
        # 配置から見つけた欄に重なる、とどくが書き込んだフォームの欄（その欄の値を、この所の印として出す）
        hits = [(f, w) for f, w, r in ws if tdk.fullmatch(f.name) and _overlap(sp.rect, r) > 0.5]
        sp.widget = hits or None
        if not hits:
            near = _near_widget(sp, real_ws)
            if near == "inside":
                # 本物のフォームの欄（元からある欄）の中に収まる所（欄の外に 1 字も入らない）：書く所はその欄そのもの
                hidden.add(id(sp))
                continue
            if near is not None:
                _flag_near_widget(sp, *near)
        spots.append(sp)
    for sp in spots:
        if sp.alternative_to is not None and id(sp.alternative_to) in hidden:
            sp.alternative_to = None
    used_widgets: set[int] = set()
    items: list[_Item] = []
    cell_items: dict[int, _Item] = {}
    line_spots: dict[int, list[Spot]] = {}
    for sp in spots:
        if isinstance(sp.host, Line):
            line_spots.setdefault(id(sp.host), []).append(sp)
    cells_with = set()
    for sp in spots:
        if isinstance(sp.host, Cell):
            cells_with.add(id(sp.host))
    for c in lay.cells:
        if c.lines or id(c) in cells_with:
            it = _Item(c.x0, c.y0, c.x1, c.y1, cell=c)
            items.append(it)
            cell_items[id(c)] = it
    for ln in lay.free_lines:
        items.append(_Item(ln.x0, ln.y0, ln.x1, ln.y1, line=ln))
    # 行にもマスにも結びつかない候補（見出しのない下線・ページの半分を超える空のマス）：その位置に、印だけを出す
    cell_ids = {id(c) for c in lay.cells}
    for sp in spots:
        if sp.host is None or (isinstance(sp.host, Cell) and id(sp.host) not in cell_ids):
            it = _Item(*sp.rect)
            it.spot = sp
            items.append(it)
    # フォームの欄：マスの中ならそのマスに、外なら単独で
    loose: list[tuple] = []
    # 配置から見つけた欄に書き込んだフォームの欄（名前が p{ページ}.{番号}・p{ページ}.a{番号} で、その欄に重なるもの）は、
    # 読み順の項目にしない：書き込みの前と同じ読み順・同じ番号になるように（値は、その欄の印として出す）
    own = re.compile(rf"p{number}\.(?:{OLD_ID}|{PLACE_ID})")
    written = {id(w) for sp in spots for f, w in (sp.widget or []) if own.fullmatch(f.name)}
    for f, w, r in ws:
        if id(w) in written:
            continue
        cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
        c = lay.index.at(cx, cy)
        if c is not None and id(c) in cell_items:
            cell_items[id(c)].marks.append((cx, f, w))
        elif c is not None:
            it = _Item(c.x0, c.y0, c.x1, c.y1, cell=c)
            items.append(it)
            cell_items[id(c)] = it
            it.marks.append((cx, f, w))
        else:
            loose.append((r, f, w))
    for r, f, w in loose:
        it = _Item(*r)
        it.marks.append(((r[0] + r[2]) / 2, f, w))
        items.append(it)
    # 行にまとめる（上から、行の中は左から）
    items.sort(key=lambda it: (round(it.y0, 1), it.x0))
    rows: list[list[_Item]] = []
    for it in items:
        if rows:
            # 行の帯：いちばん背の低い項目の上半分（縦に長い見出しのマスに引きずられないように）
            mid = min((x.y0 + min(x.y1, x.y0 + 3 * _hint(x))) / 2 for x in rows[-1])
            if it.y0 < mid:
                rows[-1].append(it)
                continue
        rows.append([it])
    names: dict[int, str] = {}

    def name(sp: Spot) -> str:
        # 所の id は位置から決めたもの（layout.place_id）。この所のものでない、元からあるフォームの欄と同じ名前のときだけ、
        # 後ろに「z」を足す（その名前の欄に書いてしまわないように）
        k = id(getattr(sp, "alias_of", sp))
        if k not in names:
            own_names = {f.name for f, _w in (getattr(getattr(sp, "alias_of", sp), "widget", None) or [])}
            cand = sp.id or f"p{number}.g0-0"
            while cand in by_name and cand not in own_names:
                cand += "z"
            names[k] = cand
        return names[k]

    # 欄にしていない空き（字のあるマスの、字の下・上の 1 字ぶん以上の空き）：マスの字の後ろに「［↓余白 n mm］」と出す
    marks_of: dict[int, str] = {}
    emitted_marks = [0]
    for host, arrow, size in (lay.room_marks if layout else []):
        marks_of[id(host)] = marks_of.get(id(host), "") + f"［{arrow}余白 {size / view.mm:.1f} mm］"

    def room_mark(host) -> str:
        m = marks_of.get(id(host))
        if m:
            emitted_marks[0] += 1
        return m or ""

    circ: dict[int, list] = {}  # 行 → ページに描かれた○が囲む字 (始め, 終わり, id)
    for ln_, a_, b_, ident in getattr(lay, "circled", []):
        circ.setdefault(id(ln_), []).append((a_, b_, ident))
    for ln_, a_, b_, ident in getattr(lay, "maybe_circled", []):
        circ.setdefault(id(ln_), []).append((a_, b_, ident, True))
    reading = []
    plain: list[str] = []
    for row in rows:
        if all(_is_vertical(it) for it in row) or _rtl_row(row):
            row.sort(key=lambda it: -it.x1)  # 縦書きの列・右から左へ書く表の行は、右から読む
        else:
            row.sort(key=lambda it: it.x0)
        parts = []
        for it in row:
            if it.line is not None:
                t = _render_line(it.line, line_spots.get(id(it.line), []), name, plain, circ.get(id(it.line))) + room_mark(it.line)
            elif getattr(it, "spot", None) is not None:
                sp = it.spot
                mark = _filled_mark(sp, name)
                t = mark if mark is not None else f"［{name(sp)}{'？' if sp.confidence == 'agent' else ''}］"
            elif it.cell is None:
                t = ""
            else:
                c = it.cell
                pieces = []
                for sp in [sp for sp in spots if sp.host is c]:
                    mark = _filled_mark(sp, name)
                    if mark is not None:
                        pieces.append(mark)
                    elif sp.kind == "checkbox":
                        pieces.append(("☑" if sp.checked else "☐") + f"［{name(sp)}{'？' if sp.confidence == 'agent' else ''}］")
                    else:
                        pieces.append(f"［{name(sp)}{'？' if sp.confidence == 'agent' else ''}］")
                rendered = []
                for ln in c.lines:
                    r_ = _render_line(ln, line_spots.get(id(ln), []), name, plain, circ.get(id(ln))) + room_mark(ln)
                    grp = getattr(ln, "overprint", None)
                    if grp is not None and rendered and rendered[-1][1] == grp:
                        rendered[-1] = (rendered[-1][0] + " ／ " + r_, grp)  # 重ねて描かれた行
                    else:
                        rendered.append((r_, grp))
                pieces.append(_join([r_ for r_, _g in rendered]) + room_mark(c))
                t = " ".join(p for p in pieces if p)
            for _, f, w in sorted(it.marks, key=lambda m: m[0]):
                if f.name in names.values():
                    continue  # 配置から見つけた欄として、値つきで出した
                t = (t + " " if t else "") + _widget_mark(f, w) + _unseen_mark(f, w)
            if t.strip():
                parts.append(t.strip())
        if parts:
            reading.append(" ｜ ".join(parts))
    # 欄の一覧（読み順の番号）
    out_fields = []
    context_of = {}
    for line in reading:
        for m in re.finditer(rf"［(p\d+\.(?:{PLACE_ID}|{OLD_ID})z*)？?[ 　]*[］＝]", line):  # 「？」の所の印は、元の空白を中に入れる
            context_of.setdefault(m[1], line)
    emitted: set[str] = set()
    for sp in spots:
        fid = names.get(id(sp))
        if fid is None:
            fid = name(sp)
        filled = next((f for f, _ in (sp.widget or []) if f.name == fid), None)
        if sp.widget and filled is None:
            continue  # 元からあるフォームの欄と重なる：フォームの欄として出す
        d = {"id": fid, "page": number, "kind": sp.kind, "label": sp.label, "reason": sp.reason,
             "rect": [round(v, 1) for v in sp.rect], "provenance": LAYOUT, "source": "layout",
             "context": context_of.get(fid, ""), "value": True if sp.checked else None, "size": round(sp.size, 2),
             "shape": sp.shape, "facts": dict(sp.facts)}
        if sp.write_size_from:
            d["size_from"] = sp.write_size_from  # 既定の字の大きさの出どころ（layout.default_size）
        if sp.baseline is not None:
            d["baseline"] = round(sp.baseline, 2)  # 行の所：行の字と同じ基線に書く（見た目の y、pt）
        if filled is not None:
            emitted.add(filled.name)
            d["filled"] = True
        if sp.kind == "comb":
            d["boxes"] = len(sp.boxes)
            d["box_rects"] = [[round(v, 2) for v in b] for b in sp.boxes]
            d["cell_ids"] = [name(x) for x in spots if x.alternative_to is sp and x.shape == "cell"]
        if sp.row_label:
            d["row_label"] = sp.row_label
        if sp.column_label:
            d["column_label"] = sp.column_label
        if sp.confidence == "agent":
            d["agent"] = True  # 描かれた空欄でない所（字の間の空きなど）：書く所かはエージェントが窓を読んで決める
            d["why"] = sp.reason
        if sp.kind_candidates:
            d["kind_candidates"] = list(sp.kind_candidates)  # 種類が形だけでは決まらない：書き方は fill の as_kind
        if sp.alternative_to is not None:
            d["alternative_to"] = name(sp.alternative_to)  # 同じ所の別の読み方（両方には書けない）
        if getattr(sp, "alt_widget", None):
            if sp.alternative_to is not None:
                # 別の候補とフォームの欄の両方が、同じ所の別の読み方：フォームの欄を先に（もう一方の候補は理由に残す）
                d["why"] = f"{d['why']}（{d['alternative_to']} とも同じ所の別の読み方）"
            d["alternative_to"] = sp.alt_widget  # 元からあるフォームの欄（同じ所の別の読み方）
        if sp.readonly:
            d["readonly"] = True
            d["why_not"] = sp.readonly
        if sp.cover:
            d["cover"] = True  # 書くときは、元の字（伏せ字）を白い下地の欄で覆う
            d["covered"] = sp.cover
            d["placeholder"] = sp.cover
        if sp.kind != "checkbox":
            d["value"] = None
        if filled is not None:
            d["value"] = filled.checked if filled.kind == "checkbox" else filled.value_text()
        out_fields.append(d)
    for f, w, r in ws:
        if f.name in emitted or any(x.get("id") == f.name for x in out_fields):
            continue
        d = {"id": f.name, "page": number, "kind": f.kind, "label": f.tu or _near_label(lay, r) or f.name.split(".")[-1],
             "name": f.name, "rect": [round(v, 1) for v in r], "provenance": DET, "source": "acroform",
             "value": f.value_text() if f.kind not in ("checkbox",) else f.checked}
        if f.tu:
            d["tooltip"] = f.tu
        near = _near_label(lay, r)
        if near and near != d["label"]:
            d["nearby_text"] = near
        nb = lay._facts_nb(lay._neighbours(r))
        if nb:
            # 欄の四角の左・上・右・下でいちばん近い字（元のまま）と、その離れ（mm）：見出し（/TU がなければ近くの字から選んだもの）を
            # エージェントが確かめるための事実
            d["neighbours"] = nb
        if f.options:
            d["options"] = [o[1] or o[0] for o in f.options]
            d["export_values"] = [o[0] for o in f.options]
            if getattr(f, "bad_opts", None):
                lay.flag("C", "字の読めない選択肢の名前",
                         "一覧から選ぶ欄の選択肢の字が、PDF の文字列の決まり（PDFDocEncoding・UTF-16BE）で読めません"
                         "（Shift_JIS などで書いたもの）。字を推し量らず、元のバイト列を示しました。どれが何の選択肢かは view_page で確かめる",
                         [f"{f.name}：" + "、".join("<" + b.hex().upper() + ">" for b in f.bad_opts[:6])])
        gone, unprinted = f.visibility()
        if gone:
            d["hidden"] = True  # 画面にも紙にも出ない欄（書くときは allow_hidden）
            d["hidden_why"] = gone
        elif unprinted:
            d["not_printed"] = True  # 画面には出るが、紙に印刷されない欄
            d["not_printed_why"] = unprinted
        if f.kind == "radio":
            if f.options:
                d["options"] = [o[1] or o[0] for o in f.options]
            else:
                from .form import name_text, name_unreadable

                d["options"] = [name_text(x.on) for x in f.widgets if x.on]
                bad = [x.on for x in f.widgets if x.on and name_unreadable(x.on)]
                if bad:
                    lay.flag("C", "字の読めない選択肢の名前",
                             "ひとつ選ぶ欄の選択肢の名前が、PDF の名前の決まり（UTF-8）で読めません（Shift_JIS などで作った名前）。"
                             "字を推し量らず、バイト列のまま（#82#CD のように）示しました。どれが何の選択肢かは view_page で確かめ、"
                             "書くときはこの形の名前で選びます", [f"{f.name}：" + "、".join(name_text(b) for b in bad[:6])])
            _label_choices(lay, f, d, ws)
        if f.maxlen:
            d["maxlen"] = f.maxlen
        if f.readonly:
            d["readonly"] = True
        if f.required:
            d["required"] = True
        if f.flags & (1 << 24):
            d["comb"] = True
        if f.ref is not None:
            from . import acrojs

            js = acrojs.actions(doc, doc.resolve(f.ref))
            for k2, src in js.items():
                p_ = acrojs.parse_call(src)
                if k2 in ("F", "K") and p_ and p_[0].startswith("AFDate"):
                    d["format"] = f"日付（{p_[1][0] if p_[1] and isinstance(p_[1][0], str) else '既定の形'}）"
                elif k2 in ("F", "K") and p_ and p_[0].startswith("AFNumber"):
                    d["format"] = "数（桁区切りなどは自動）"
                elif k2 in ("F", "K") and p_ and p_[0].startswith("AFPercent"):
                    d["format"] = "百分率"
                elif k2 in ("F", "K") and p_ and p_[0].startswith("AFSpecial"):
                    d["format"] = ["郵便番号（5桁）", "郵便番号（9桁）", "電話番号", "社会保障番号"][int(p_[1][0]) if p_[1] else 0]
                elif k2 == "C":
                    d["calculated"] = src[:200]
                elif not (p_ and p_[0] in acrojs.KNOWN):
                    d.setdefault("scripts", []).append(k2)
        out_fields.append(d)
    # 配置の所は位置の順（上から、同じ高さは左から）、フォームの欄はその後ろに元の順で
    out_fields.sort(key=lambda d: (0, round(d["rect"][1]), d["rect"][0]) if d.get("source") == "layout" else (1, 0, 0))
    images = len(pc.pictures)
    pictures = _pictures(doc, page, number, pc, lay)
    chars = sum(1 for c in lay.chars if not c.text.isspace())
    annots = []
    for a in doc.resolve(page.get("Annots")) or []:
        ad = doc.resolve(a)
        if not isinstance(ad, dict):
            continue
        st = getattr(ad.get("Subtype"), "name", "")
        if st in ("Widget", "Link", "Popup"):
            continue
        from .objects import decode_text

        txt = decode_text(doc.resolve(ad.get("Contents"))) if ad.get("Contents") is not None else ""
        drawn, unread = _appearance_text(doc, ad)
        if txt or drawn or unread or st in ("FreeText", "Stamp", "Text"):
            d_ = {"kind": st, "text": txt}
            if drawn and drawn.split() != txt.split():
                d_["drawn"] = drawn  # 注釈の見た目（/AP）に描かれた字：ページに見えるのはこの字
            if unread:
                d_["unread"] = unread  # 見た目を読めなかった：ページに見える字は、窓に出ていないかもしれない
                try:
                    r_ = [float(doc.resolve(v)) for v in doc.resolve(ad.get("Rect"))]
                    vx0, vy0, _vx1, _vy1 = lay.view.rect(min(r_[0], r_[2]), min(r_[1], r_[3]), max(r_[0], r_[2]), max(r_[1], r_[3]))
                    where = lay.where(vx0, vy0, txt)
                except (TypeError, ValueError, IndexError):
                    where = f"{number}ページ（位置の分からない注釈）"
                lay.flag("B", "見た目を読めない注釈",
                         "注釈（スタンプ・FreeText など）の見た目（/AP）を読めませんでした。ページに見える字が窓に出ていないかも"
                         "しれないので、その位置は view_page で画像を見て確かめる（中身を推測で埋めない）",
                         [f"{where}（{st}：{unread}）"])
            annots.append(d_)
    marks, inword = unmapped_places(lay)
    circled = [{"text": ln_.text[a_:b_], "id": ident or None, "context": re.sub(r"[ 　]+", " ", ln_.text).strip()[:60],
                "rect": [round(v, 1) for v in ln_.span_rect(a_, b_)]} for ln_, a_, b_, ident in getattr(lay, "circled", [])]
    # 角の丸い四角など（曲線と長さのある直線で閉じた図形）が囲む字：○で選んだ印か、ただの枠か、規則では決められない
    maybe_circled = [{"text": ln_.text[a_:b_], "context": re.sub(r"[ 　]+", " ", ln_.text).strip()[:60],
                      "rect": [round(v, 1) for v in ln_.span_rect(a_, b_)], "agent": True,
                      "why": "曲線と直線で閉じた図形（角の丸い四角など）が囲む字：○で選んだ印か、ただの枠か、規則では決められない"}
                     for ln_, a_, b_, _ident in getattr(lay, "maybe_circled", [])]
    return {"number": number, "width": round(view.width, 1), "height": round(view.height, 1), "rotate": view.rot,
            # UserUnit（ページの座標の 1 単位が何ポイントか）：1 でないページだけ。窓の rect・大きさは、ページの座標の単位
            **({"user_unit": view.unit} if view.unit != 1.0 else {}),
            "fields": out_fields, "reading": reading, "circled": circled, "maybe_circled": maybe_circled, "chars": chars, "unmapped": pc.unmapped, "images": images,
            # 字に直せない字の、行の中での役割（字の並びの形で決める）：行の頭の印（箇条の印）と、それ以外（本文の字）
            "unmapped_marks": marks, "unmapped_text": inword,
            "image_only": chars == 0 and images > 0, "pictures": pictures, "problems": pc.problems, "annotations": annots,
            "_plain": plain, "_dropped": lay.dropped, "_added": lay.added, "read_flags": lay.flags,
            # 窓の作り方で1つにした字：同じ位置に重ねて描いた同じ字（太字に見せる重ね書き）・縦に1字ずつ積んだ字の列
            "merged_glyphs": sum(1 for k, _t in lay.dropped if k == "dup"),
            "unseen_glyphs": sum(1 for k, _t in lay.dropped if k == "size"),  # 0.5pt に満たない字・ページの何倍もある字
            "slivers": len(lay.slivers) if layout else 0,
            "room_marks": emitted_marks[0],  # 欄にしていない空きの印（［↓余白 n mm］）の数
            "stacked_columns": sum(1 for ln in [x for c in lay.cells for x in c.lines] + list(lay.free_lines)
                                   if getattr(ln, "stacked", 0))}


def _where(view, x0: float, y0: float, x1: float, y1: float) -> str:
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    return ("上" if cy < view.height / 3 else "中ほど" if cy < view.height * 2 / 3 else "下") + "の" + \
        ("左" if cx < view.width / 3 else "中央" if cx < view.width * 2 / 3 else "右")


def unmapped_places(lay) -> tuple[list[dict], list[dict]]:
    """字に直せない字（�）の所を、行の中の役割で2つに分ける（字の並びの形だけで決める。数や字の種類では決めない）：
    - 行の頭の印：行の最初の字で、すぐ後ろに空き（空白）があり、その後ろに字の続くもの（「•　項目の文」の •）。
      箇条の印・飾りの形で、行の中の語は1字も欠けていない
    - 本文の字：それ以外（語の中・行の途中・字の続く所）。中身の字が欠けている
    どちらも {行の位置, 近くの字, 数} の一覧（行ごと）。"""
    from .fonts import UNKNOWN

    view = lay.view
    marks, inword = [], []
    for ln in [x for c in lay.cells for x in c.lines] + list(lay.free_lines):
        t = ln.text
        if UNKNOWN not in t:
            continue
        lead = len(t) - len(t.lstrip(" \u3000"))
        rest = t[lead + 1:]
        n = t.count(UNKNOWN)
        is_mark = (t[lead:lead + 1] == UNKNOWN and rest[:1] in (" ", "\u3000") and rest.strip(" \u3000")
                   and UNKNOWN not in rest)
        near = (rest if is_mark else t).strip(" \u3000")[:30]
        item = {"where": _where(view, ln.x0, ln.y0, ln.x1, ln.y1), "rect": [round(v, 1) for v in (ln.x0, ln.y0, ln.x1, ln.y1)],
                "near": near, "count": n}
        (marks if is_mark else inword).append(item)
    return marks, inword


def _appearance_text(doc: PdfDocument, ad: dict) -> tuple[str, str]:
    """注釈の見た目（/AP の /N）に描かれた字を、描いた順に（スタンプ・FreeText の字は、ページの描画の命令の外にある）。
    (字, 読めなかった理由)。見た目のない注釈は ("", "")。"""
    from .content import Interpreter
    from .objects import Stream

    ap = doc.resolve(ad.get("AP"))
    n = doc.resolve(ap.get("N")) if isinstance(ap, dict) else None
    if isinstance(n, dict) and not isinstance(n, Stream):
        st = getattr(doc.resolve(ad.get("AS")), "name", None)
        n = doc.resolve(n.get(st)) if st else None
    if not isinstance(n, Stream):
        return "", ""
    try:
        pc = Interpreter(doc).run_stream(doc.decode(n), doc.resolve(n.dict.get("Resources")) or {})
    except Exception as e:  # noqa: BLE001  読めないことを、黙って「字がない」にしない
        return "", f"{type(e).__name__}: {str(e)[:60]}"
    return "".join(g.text for g in pc.glyphs).strip(), ""


def _pictures(doc: PdfDocument, page: dict, number: int, pc, lay) -> list[dict]:
    """ページの画像：位置・大きさ・説明（代わりの字）。説明はファイルの構造（タグ・印）にあるものだけ。"""
    from .tags import StructTree, describe_pictures

    if not pc.pictures:
        return []
    tree = doc.__dict__.get("_struct_tree")
    if tree is None:
        tree = doc.__dict__["_struct_tree"] = StructTree(doc)
    view = lay.view
    lines = [ln for c in lay.cells for ln in c.lines] + list(lay.free_lines)
    out = []
    for k, (p, d) in enumerate(zip(pc.pictures, describe_pictures(doc, page, pc.pictures, tree)), 1):
        x0, y0, x1, y1 = view.rect(p.x0, p.y0, p.x1, p.y1)
        if x1 <= 0 or y1 <= 0 or x0 >= view.width or y0 >= view.height:
            continue  # ページの外（見えない）
        cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
        where = ("上" if cy < view.height / 3 else "中ほど" if cy < view.height * 2 / 3 else "下") + "の" + \
                ("左" if cx < view.width / 3 else "中央" if cx < view.width * 2 / 3 else "右")
        near = ""
        if lines:
            def dist(ln):
                dx = max(ln.x0 - x1, x0 - ln.x1, 0)
                dy = max(ln.y0 - y1, y0 - ln.y1, 0)
                return dx * dx + dy * dy
            ln = min(lines, key=dist)
            near = ln.text.strip()[:30]
        out.append({"id": f"p{number}.i{k}", "rect": [round(v, 1) for v in (x0, y0, x1, y1)], "where": where,
                    "size_mm": [round((x1 - x0) / view.mm), round((y1 - y0) / view.mm)], "near": near,
                    "alt": d["alt"], "decorative": d["decorative"], "tagged": d["tagged"]})
    return out


def _join(texts: list[str]) -> str:
    out = ""
    for t in texts:
        t = t.strip()
        if not t:
            continue
        if out and not (_cjk(out[-1]) and _cjk(t[0])):
            out += " "
        out += t
    return out


def _cjk(ch: str) -> bool:
    from .layout import no_word_space

    return no_word_space(ch) and ch not in "［］"


def _iou(a, b) -> float:
    x0, y0, x1, y1 = max(a[0], b[0]), max(a[1], b[1]), min(a[2], b[2]), min(a[3], b[3])
    if x1 <= x0 or y1 <= y0:
        return 0.0
    i = (x1 - x0) * (y1 - y0)
    return i / ((a[2] - a[0]) * (a[3] - a[1]) + (b[2] - b[0]) * (b[3] - b[1]) - i)


def _inherit_ids(spots: list[Spot], written: list[tuple], number: int) -> None:
    """とどくが前に書き込んだ欄（名前が配置の所の名前：前の版の番号 p1.3・p1.a12 か、位置の名前）が、ちょうど 1 つの所と
    重なる（重なりが、合わせた広さの半分より多い：両方の四角がほぼ同じ）とき、その所の id を、書いてある欄の名前にする
    （書いた書類を開き直しても、同じ名前で書き直せる）。どの所とも決まらない欄（重なる所が 2 つ以上・ない）は、名前を移さない
    （その欄は、フォームの欄として窓に出る）。字を目印にした所（w）・印影（s）・○（c）は、所の名前にしない。"""
    pat = re.compile(rf"p{number}\.(?:{OLD_ID}|{PLACE_ID})")
    for f, r in written:
        if not pat.fullmatch(f.name):
            continue
        hits = [sp for sp in spots if _iou(sp.rect, r) > 0.5]
        if len(hits) == 1 and hits[0].id != f.name:
            taken = any(sp.id == f.name for sp in spots)
            if not taken:
                hits[0].id = f.name


def _beside_widget(rect, r) -> bool:
    """配置の所 rect が、フォームの欄 r に重なるか、同じ行（上下の半分以上が重なる）で、欄の高さの内に接するか。"""
    h = min(rect[3] - rect[1], r[3] - r[1])
    vov = min(rect[3], r[3]) - max(rect[1], r[1])
    if h <= 0 or vov < 0.5 * h:
        return False
    gap = max(rect[0] - r[2], r[0] - rect[2])  # 左右の離れ（負なら重なる）
    return gap <= (r[3] - r[1])


def _inside_widget(sp: Spot, r) -> bool:
    """配置の所が、フォームの欄 r の中に収まるか：欄からはみ出す幅が、どの辺でもその所の字 1 字に足りない
    （欄の周りに描いた枠の線・余白だけで、欄の外に書く所がない）。"""
    one = max(sp.size, 1.0)
    x0, y0, x1, y1 = sp.rect
    return r[0] - x0 < one and r[1] - y0 < one and x1 - r[2] < one and y1 - r[3] < one


_CHECK_WIDGETS = ("checkbox", "radio")


def _em_overlap(sp: Spot, r) -> float:
    """配置の所と、フォームの欄が、行の向きに重なる長さを、1 字ぶん（1 em）を 1 として測った数（重ならなければ 0）。
    1 em は、その所の字の大きさ。ただし所か欄が 1 字より狭いとき（小さなチェックの箱など）は、狭いほうの幅を 1 とする。
    行と直角の向きには、狭いほうの高さの半分以上が重なるときだけ重なると見る（上下の行にかかるだけの欄は重ならない）。"""
    vertical = isinstance(sp.host, Line) and sp.host.vertical
    a0, a1, b0, b1 = (1, 3, 0, 2) if vertical else (0, 2, 1, 3)  # 行の向き・直角の向き
    along = min(sp.rect[a1], r[a1]) - max(sp.rect[a0], r[a0])
    across = min(sp.rect[b1], r[b1]) - max(sp.rect[b0], r[b0])
    if along <= 0 or across < 0.5 * min(sp.rect[b1] - sp.rect[b0], r[b1] - r[b0]):
        return 0.0
    em = min(max(sp.size, 1.0), sp.rect[a1] - sp.rect[a0], r[a1] - r[a0])
    return along / max(em, 1e-6)


def _near_widget(sp: Spot, real_ws: list):
    """配置の所と、元からあるフォームの欄との関係："inside"（欄の中に収まる：その欄そのもの）、
    (欄, "overlap" か "beside", 離れ pt)（1 字ぶん以上重なる・同じ行ですぐ隣か 1 字に足りないだけ重なる）、None（関係なし）。
    1 字ぶん以上重なる所だけを、欄と同じ所の別の読み方の候補にする（両方に書くと字が重なる）。接する・わずかに重なるだけの所は、
    欄の隣に別に書く所（「電話番号［欄］（　　）」のかっこの中）のことがあるので、重なりとしない。"""
    best = None
    for f, _w, r in real_ws:
        if _inside_widget(sp, r):
            return "inside"
        em = _em_overlap(sp, r)
        if em >= 1.0:
            key = (0, -em)
            cand = (f, "overlap", 0.0)
        elif _overlap(sp.rect, r) > 0 or _beside_widget(sp.rect, r):
            gap = max(sp.rect[0] - r[2], r[0] - sp.rect[2], 0.0)
            key = (1, gap)
            cand = (f, "beside", gap)
        else:
            continue
        if best is None or key < best[0]:
            best = (key, cand)
    return best[1] if best is not None else None


def _flag_near_widget(sp: Spot, f: FormField, how: str, gap: float) -> None:
    """元からあるフォームの欄に重なる・すぐ隣の所：消さずに、その欄の名前を理由に添える。
    - 1 字ぶん以上重なる所（how="overlap"）：「？」（規則では決められない）。同じ種類の読み方（字の所と字の欄・チェックの所と
      チェックの欄）なら、同じ所の別の読み方（alternative_to：どちらか一方に書く。fill は両方に書くと断る）。種類が違えば
      （「□ その他（欄）」の □ と字の欄など）、両方に書くこともあるので alternative_to にはしない。
    - すぐ隣の所（how="beside"）：重ならないので alternative_to にはしない。字の形だけで決まる規則の空欄（かっこの中・下線）は、
      欄の隣にあっても規則の空欄のまま（理由に欄の名前を添えるだけ）。それ以外は「？」（理由に欄の名前）。"""
    same = (sp.kind == "checkbox") == (f.kind in _CHECK_WIDGETS)
    if how == "overlap":
        where = f"フォームの欄 {f.name} に重なる"
        if same:
            why = f"{where}：その欄と同じ所の別の読み方かもしれない（書くならどちらか一方）"
            sp.alt_widget = f.name
        else:
            why = (f"{where}：種類が違う（この所は{KIND_JA.get(sp.kind, sp.kind)}、欄は{KIND_JA.get(f.kind, f.kind)}）。"
                   "その欄とは別に書く所か、規則では決められない")
    else:
        where = f"フォームの欄 {f.name} と同じ行の、すぐ隣（{gap:.0f}pt）"
        if sp.kind == "inline" and sp.confidence != "agent":
            why = f"{where}。重ならず、字の形で決まる空欄なので、その欄とは別の所"
            sp.reason = f"{sp.reason}。{why}" if sp.reason else why
            return
        why = f"{where}：その欄の続きか、その欄とは別に書く所か、規則では決められない（重ならない）"
    sp.confidence = "agent"
    sp.reason = f"{sp.reason}。{why}" if sp.reason else why


def _rtl_row(row: list[_Item]) -> bool:
    """行の字が、すべて右から左へ書く字の行か。"""
    lines = []
    for it in row:
        if it.line is not None:
            lines.append(it.line)
        elif it.cell is not None:
            lines += it.cell.lines
    lines = [ln for ln in lines if not ln.vertical and ln.text.strip()]
    return bool(lines) and all(getattr(ln, "rtl", False) for ln in lines)


def _is_vertical(it: _Item) -> bool:
    if it.line is not None:
        return it.line.vertical
    if it.cell is not None and it.cell.lines:
        return all(ln.vertical for ln in it.cell.lines)
    return False


def _hint(it: _Item) -> float:
    if it.line is not None:
        return it.line.size
    if it.cell is not None and it.cell.lines:
        return max(ln.size for ln in it.cell.lines)
    return max(6.0, (it.y1 - it.y0) / 3)


def _label_within_line(lay: PageLayout, r) -> str:
    """フォームの欄が、同じ高さ（字の高さの半分以上が欄と重なる）の行の、字と字の間の空きに置かれているとき：
    その行の、欄より前の字（左から右へ書く行は欄の左、右から左へ書く行は欄の右）。
    「氏名［欄］年　月　日」の欄は「氏名」の欄で、上の行の見出しではない。欄の前で終わる行（「見出し　［欄］」）も同じ。
    同じ高さに、欄の前の字（200pt の内）がなければ ""。"""
    best = None
    for ln in lay.free_lines:
        if ln.vertical or not ln.text.strip():
            continue
        h = ln.y1 - ln.y0
        if h <= 0 or min(ln.y1, r[3]) - max(ln.y0, r[1]) < 0.5 * h:
            continue
        rtl = bool(getattr(ln, "rtl", False))
        # 見た目の順（左から右へ書く行は左から、右から左へ書く行は右から）に、欄より前で始まる字を取る。欄にかかって
        # 続く語の字（見出しの終わりが欄の下にはみ出した「Given name (…given)」）は、空白か 1 字の半分以上の空きまで足す
        order = sorted(range(len(ln.pos)), key=lambda i: -ln.pos[i][1] if rtl else ln.pos[i][0])
        idx: list[int] = []
        for i in order:
            p = ln.pos[i]
            if (p[1] > r[2]) if rtl else (p[0] < r[0]):
                idx.append(i)
                continue
            q = ln.pos[idx[-1]] if idx else None
            if q is None or ln.text[i].isspace() or ((q[0] - p[1]) if rtl else (p[0] - q[1])) >= 0.5 * ln.size:
                break
            idx.append(i)
        head = "".join(ln.text[i] for i in sorted(idx)).strip()
        if not head:
            continue
        gap = (min(ln.pos[i][0] for i in idx) - r[2]) if rtl else (r[0] - max(ln.pos[i][1] for i in idx))
        if gap > 200:
            continue
        if best is None or gap < best[0]:
            best = (gap, head)
    return label_text(best[1]) if best is not None else ""


def _widget_line(lay: PageLayout, r):
    """フォームの部品と同じ高さ（字の高さの半分以上が部品と重なる）の横書きの行のうち、部品にいちばん近いもの（200pt の内）。"""
    best = None
    for ln in list(lay.free_lines) + [x for c in lay.cells for x in c.lines]:
        if ln.vertical or getattr(ln, "rtl", False) or not ln.text.strip():
            continue
        h = ln.y1 - ln.y0
        if h <= 0 or min(ln.y1, r[3]) - max(ln.y0, r[1]) < 0.5 * h:
            continue
        gap = max(ln.x0 - r[2], r[0] - ln.x1, 0)
        if gap <= 200 and (best is None or gap < best[0]):
            best = (gap, ln)
    return best[1] if best is not None else None


def _run(ln: Line, idx: list[int], toward_right: bool, far: float) -> str:
    """行の字（idx：見た目の順の番号）を、部品の側から外へたどった字の並び。
    字と字の空きが far 以上の所で止める（別の欄・別の段の字）。"""
    seq = idx if toward_right else list(reversed(idx))
    got: list[int] = []
    last = None
    for i in seq:
        if ln.text[i].isspace():
            continue
        p = ln.pos[i]
        if last is not None and ((p[0] - last[1]) if toward_right else (last[0] - p[1])) >= far:
            break
        got.append(i)
        last = p
    return ln.text[min(got):max(got) + 1].strip() if got else ""


def _choice_labels(lay: PageLayout, items: list[tuple[str, tuple]]) -> tuple[str, list[str]] | None:
    """ひとつ選ぶ欄の部品（選択肢の名前, 部品の位置）の並びから、(組の見出し, 選択肢ごとの字)。
    選択肢の字は、同じ行の部品のすぐ隣の字（部品と部品の間の字）。字が部品の左か右かは、行の部品のすべてに字のある側のうち、
    部品と字の空きの合計が小さい側。組の見出しは、最初の選択肢より前の字（同じ行）。なければ、マスの行の見出しか上の行。
    どれかの選択肢の字が決まらないときは None（いままでどおり、最初の部品の近くの字を見出しにする）。"""
    lines = []
    for _o, r in items:
        ln = _widget_line(lay, r)
        if ln is None:
            return None
        lines.append(ln)
    gaps: dict[str, list] = {"before": [], "after": []}
    regions = []
    for k, ((_o, r), ln) in enumerate(zip(items, lines)):
        same = [items[j][1] for j in range(len(items)) if j != k and lines[j] is ln]
        lo = max([x[2] for x in same if x[2] <= r[0] + 1] or [float("-inf")])
        hi = min([x[0] for x in same if x[0] >= r[2] - 1] or [float("inf")])
        order = sorted(range(len(ln.pos)), key=lambda i: ln.pos[i][0])
        left = [i for i in order if ln.pos[i][1] <= r[0] + 1 and ln.pos[i][0] >= lo - 1]
        right = [i for i in order if ln.pos[i][0] >= r[2] - 1 and ln.pos[i][1] <= hi + 1]
        regions.append((left, right, lo > float("-inf"), hi < float("inf")))
        for side, idx, toward in (("before", left, False), ("after", right, True)):
            ink = [i for i in idx if not ln.text[i].isspace()]
            if not ink:
                gaps[side].append(None)
            else:
                gaps[side].append((r[0] - max(ln.pos[i][1] for i in ink)) if not toward else (min(ln.pos[i][0] for i in ink) - r[2]))
    sides = [s for s in ("before", "after") if all(g is not None for g in gaps[s])]
    if not sides:
        return None
    side = min(sides, key=lambda s: sum(gaps[s]))
    labels = []
    group = ""
    for k, (ln, (left, right, bounded_l, bounded_r)) in enumerate(zip(lines, regions)):
        far = 1.5 * ln.size
        if side == "after":
            t = ln.text[min(right):max(right) + 1].strip() if bounded_r else _run(ln, right, True, far)
            if k == 0 and not bounded_l:
                group = _run(ln, left, False, 3 * ln.size)
        elif bounded_l:
            t = ln.text[min(left):max(left) + 1].strip()
        else:
            whole = _run(ln, left, False, 3 * ln.size)
            # 最初の選択肢の字の前に、組の見出しが同じ行にあるとき（「Gender: Male □」）：見出しの終わりの「：」か、
            # 空白で分ける（部品にいちばん近い語が選択肢の字）
            m = re.match(r"(.*[：:])\s*(\S.*)$", whole) or re.match(r"(.*\S)\s+(\S+)$", whole)
            t = m[2] if m else whole
            if k == 0 and m:
                group = m[1]
        labels.append(label_text(t))
    if not all(labels) or len(set(labels)) < len(labels):
        return None
    if not group:
        r0 = items[0][1]
        c = lay.index.at((r0[0] + r0[2]) / 2, (r0[1] + r0[3]) / 2)
        if c is not None:
            group = lay._label_for(c, c)[0]  # マスの中の選択肢：行の見出し（なければ列の見出し）
        else:
            top = min(r[1] for _o, r in items)
            x0, x1 = min(r[0] for _o, r in items), max(r[2] for _o, r in items)
            above = None
            for ln in lay.free_lines:
                if ln.y1 <= top + 1 and top - ln.y1 < 30 and ln.x0 < x1 and ln.x1 > x0 - 20 and ln not in lines:
                    if above is None or ln.y1 > above.y1:
                        above = ln
            group = above.text if above is not None else ""
    return label_text(group), labels


def _option_name(f: FormField, w: Widget) -> str | None:
    """ひとつ選ぶ欄の部品の選択肢の名前（窓の options と同じ字）。"""
    from .form import name_text

    if not w.on:
        return None
    opt = name_text(w.on)
    if f.options and opt.isdigit() and int(opt) < len(f.options):
        opt = f.options[int(opt)][1] or f.options[int(opt)][0]
    return opt


def _label_choices(lay: PageLayout, f: FormField, d: dict, ws: list) -> None:
    """ひとつ選ぶ欄（部品が 2 つ以上）：見出しを組の見出しに、選択肢ごとの字を options_labels に、いまの値を選んだ選択肢の字に。
    最初の部品の隣の字（「Male」）を、欄全体の見出しにしないように。"""
    mine = [(w, r) for f2, w, r in ws if f2 is f]
    if len(mine) < 2 or len(mine) != len(f.widgets):
        return  # ほかのページにも部品がある欄：このページだけでは選択肢の字をそろえられない
    names = [_option_name(f, w) for w, _r in mine]
    if None in names:
        return
    got = _choice_labels(lay, [(n, r) for n, (_w, r) in zip(names, mine)])
    if got is None:
        return
    group, labels = got
    by_name: dict[str, str] = {}
    for n, lab in zip(names, labels):
        if by_name.setdefault(n, lab) != lab:
            return  # 同じ選択肢の部品に、違う字（どちらが選択肢の字か決まらない）
    opts = d.get("options") or []
    if not opts or any(o not in by_name for o in opts):
        return
    d["options_labels"] = [by_name[o] for o in opts]
    if not f.tu:
        d["label"] = group or f.name.split(".")[-1]
    d.pop("nearby_text", None)
    if group and group != d["label"]:
        d["nearby_text"] = group
    v = d.get("value")
    if isinstance(v, str) and v in by_name and by_name[v] != v:
        d["export_value"] = v  # 書類に記録される選択肢の名前（/V）
        d["value"] = by_name[v]  # 本人に伝える値：選んだ選択肢の隣に印刷された字


def _near_label(lay: PageLayout, r) -> str:
    """フォームの欄のそばの文字（同じマスの文字、欄が字の間に置かれた同じ行の前の字、左の行、上の行）。"""
    cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
    c = lay.index.at(cx, cy)
    if c is not None:
        if c.lines:
            return label_text(c.text())
        label, _rl, _cl = lay._label_for(c, c)
        if label:
            return label
    within = _label_within_line(lay, r)
    if within:
        return within
    best = None
    for ln in lay.free_lines:
        if ln.vertical:
            continue
        if ln.x1 <= r[0] + 2 and ln.y0 < r[3] and ln.y1 > r[1] and r[0] - ln.x1 < 200:
            if best is None or ln.x1 > best.x1:
                best = ln
    if best is not None:
        return label_text(best.text)
    above = None
    for ln in lay.free_lines:
        if ln.y1 <= r[1] + 1 and r[1] - ln.y1 < 30 and ln.x0 < r[2] and ln.x1 > r[0] - 20:
            if above is None or ln.y1 > above.y1:
                above = ln
    return label_text(above.text) if above is not None else ""


def unknown_id_message(key: str) -> str:
    """窓にない名前を書き込みに渡されたときに添える知らせ（書き込みの側が使う）：前の版のとどくの番号の名前（p1.3・p1.a12）は、
    位置の名前（p1.t45-112）に変わった。番号から位置を推し量ると、違う所に書くおそれがあるので、読み替えずに断る。"""
    if re.fullmatch(rf"p\d+\.{OLD_ID}", str(key)):
        return (f"{key} は、前の版のとどくの番号の名前です。所の名前は、位置から決める形（p1.t45-112：作りの頭の字と、左・上の mm）に"
                "変わりました。番号からは同じ所と決められないので、書きません。窓（read_window・list_fields）を読み直して、新しい名前で"
                "書いてください（前の版で書き込んだ欄は、元の名前のまま書き直せます）")
    return ""


def build(doc: PdfDocument) -> dict:
    blob = doc.blob
    fields = read_fields(doc)
    out = {
        "todoku": "0.2",
        "kind": "pdf",
        "source": {"file": doc.filename, "sha256": hashlib.sha256(blob).hexdigest(), "bytes": len(blob),
                   "note": "元のファイルが正本。書き込みは、元のバイト列の後ろに足す（増分更新）ので、元の部分は1バイトも変わらない。"},
        "provenance_legend": {DET: "AIを使わない決定的な処理で、ファイルから取り出した",
                              LAYOUT: "配置（罫線・空き・字の並び）から決定的に推定した入力欄。推定なので context と合わせて読む"},
        "notation": {"［p1.t45-112］": "描かれた空欄（空のマス・下線の罫線・下線の字・かっこの中・箱の字）：書く所かは、周りの字（facts）と"
                                       "合わせてエージェントが決める",
                     "［p1.g60-40？］": "字の間の 1 字ぶん以上の空き・点線・ダッシュ・行の前後の空き、空のマスの並び：描かれた空欄では"
                                       "ない所（エージェントが読んで判断する。書くなら id を指定する。書かなければ何も変わらない）",
                     "［p1.g60-40？　］": GAP_TOKEN_NOTE + "（PDF の行の空白は、測った空きから作った字：全角の空白 1 つでおよそ 1 字ぶん）",
                     "［↓余白 12.0 mm］": "欄にしていない、字のあるマスの字の下の空き（↑ は字の上。12.0 mm はその高さ）。"
                                          "書く所と判断したら、字を目印にして {\"@1\": {\"after\": 見出しの字, \"where\": \"below\", "
                                          "\"value\": 書く字, \"page\": ページ}}（上は {\"page\": …, \"box\": [左, 上, 右, 下]（mm）, "
                                          "\"value\": …}）で書く",
                     "p1.t45-112": "所の id：p{ページ}.{作りの頭の字}{左}-{上}（作りの位置を mm に丸めた数。t 空のマス・m 空のマスの並び・"
                                   "g 行の中の空き・b 箱の字・u 下線の罫線）。位置だけで決まるので、ほかの所が増減しても変わらない",
                     "［○］": "直前の字を、ページに描かれた○（線の曲線）が囲んでいる（とどくが描いた○は［p1.c1：○］）",
                     "［○？］": "直前の字を、曲線と直線で閉じた図形（角の丸い四角など）が囲んでいる：○で選んだ印か、ただの枠か、"
                               "規則では決められない（view_page で見て確かめる）",
                     "［名前＝値］": "フォームの欄といまの値", "☐［p1.5］ / ☑": "チェックの印",
                     "○［欄＝選択肢］ / ◉": "ひとつ選ぶ欄の選択肢", " ｜ ": "表のマスの区切り"},
        "form": {"acroform_fields": len(fields), "xfa": has_xfa(doc), "encrypted": doc.decryptor is not None},
        "pages": [],
    }
    xw = _xfa_window(doc)
    for i, p in enumerate(doc.pages()):
        out["pages"].append(page_window(doc, p, i + 1, fields, layout=xw is None or not xw.get("readable")))
    if xw is not None:
        out["xfa"] = xw
        out["form"]["xfa_kind"] = xw.get("kind")
        _merge_xfa(out, xw)
    for pg in out["pages"]:
        with_label_keys(pg["fields"])  # 見出しは字のまま。探すための鍵（字間の空白を詰めた写し）は label_key
    if xw is not None and xw.get("fields"):
        with_label_keys(xw["fields"])
    return out


# ====================================================================== XFA
def _xfa_window(doc: PdfDocument) -> dict | None:
    """XFA の窓（XFA がなければ None。読めなければ、読めない理由だけ）。"""
    from .xfa import XfaError, build_window, read

    try:
        x = read(doc)
    except XfaError as e:
        return {"readable": False, "error": str(e)}
    return build_window(x) if x is not None else None


def _merge_xfa(out: dict, xw: dict) -> None:
    """静的な XFA のフォーム：ページの PDF の欄に、同じ名前の XFA の欄の見出し・決まり・値（Adobe が見せる値）を重ねる。"""
    if not xw.get("readable"):
        return
    by_id = {d["id"]: d for d in xw["fields"]}
    for pg in out["pages"]:
        for d in pg["fields"]:
            x = by_id.get(d["id"])
            if x is None:
                continue
            x["page"] = pg["number"]
            for k, v in x.items():
                if k in ("label",) and not v:
                    continue
                if k == "kind" and d.get("kind") == "radio" and v != "radio":
                    continue
                d[k] = v


# ====================================================================== Markdown
def _md(s) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


def _kind_text(f: dict) -> str:
    k = KIND_JA.get(f["kind"], f["kind"])
    if f["kind"] == "comb":
        k += f"（{f['boxes']} 字）"
    if f.get("options"):
        opts = [str(o) for o in f["options"]]
        if f.get("options_labels"):
            # 選択肢の名前と、その隣に印刷された字が違うとき：「Male（M）」
            opts = [o if lab == o else f"{lab}（{o}）" for o, lab in zip(opts, map(str, f["options_labels"]))]
        k += ": " + "／".join(opts[:15]) + (f"（ほか{len(opts) - 15}件）" if len(opts) > 15 else "")
    if f.get("maxlen"):
        k += f"（{f['maxlen']} 字まで）"
    if f.get("readonly"):
        k += "（読み取り専用）"
    if f.get("hidden"):
        k += f"（画面にも紙にも出ない欄：{f.get('hidden_why', '')}）"
    elif f.get("not_printed"):
        k += f"（画面には出るが、紙に印刷されない欄：{f.get('not_printed_why', '')}）"
    if f.get("required"):
        k += "（必須）"
    if f.get("agent"):
        k = f"規則では決められない：{f.get('why', '')}"
        if f.get("kind_candidates"):
            k += "（種類の候補：" + "／".join(KIND_JA.get(x, x) for x in f["kind_candidates"]) + "）"
        if f.get("alternative_to"):
            k += f"（{f['alternative_to']} と同じ所の別の読み方。両方には書けない）"
        if f.get("readonly"):
            k += f"（書けない：{f.get('why_not', '')}）"
    if f.get("group"):
        k += "（縦に並ぶ組：" + "・".join(f["group"]) + "）"
    if f.get("format"):
        k += f"（書式: {f['format']}）"
    if f.get("calculated"):
        k += "（計算される欄）"
    return k


_SHAPE_JA = {"cell": "空のマス", "cells": "空のマスの並び", "gap": "字間", "lead": "行頭の前", "trail": "行末の後",
             "connector": "下線の字", "leader": "点線", "dash": "ダッシュ", "bracket": "かっこの中", "blank_line": "空白の字だけの行", "placeholder": "伏せ字の形",
             "box": "箱の字", "checked": "記入済みの箱の字", "underline": "下線の罫線"}
_SIDE_JA = {"left": "左", "above": "上", "right": "右", "below": "下"}
_BOUND_JA = {"cell": "マスの縁", "text": "字", "ruling": "縦の罫線", "page": "ページの端"}
_LINE_SHAPES = ("gap", "lead", "trail", "connector", "leader", "dash", "bracket", "blank_line", "placeholder", "box", "checked")


def _q(t: str, n: int = 24) -> str:
    """周りの字（元のまま）。長い字は窓の一覧では n 字で切り、「…」と字の数を添える（全体は facts と読み順に）。"""
    t = re.sub(r"\s+", " ", str(t)).strip()
    return f"「{_md(t)}」" if len(t) <= n else f"「{_md(t[:n])}…」（{len(t)} 字）"


def _nb_text(fa: dict, sides=("left", "above", "right", "below")) -> str:
    return "・".join(f"{_SIDE_JA[k]}{_q(fa[k]['text'])}{fa[k]['mm']}mm" for k in sides if fa.get(k))


def _size_text(f: dict) -> str:
    """書く字の既定の大きさ：行の中の所（行の字の大きさ）は窓に出さない（行の字と同じ）。"""
    if not f.get("size") or f.get("size_from") in (None, "", "行の字の大きさ"):
        return ""
    return f"字 {f['size']:g}pt（{f['size_from'].replace('の大きさ', '')}）"


def _facts_text(f: dict) -> str:
    """配置の所の事実を 1 行に（mm）。"""
    fa = f.get("facts") or {}
    shape = f.get("shape", "")
    parts = []
    head = _SHAPE_JA.get(shape, KIND_JA.get(f["kind"], f["kind"]))
    if shape == "cell" and fa.get("size_mm"):
        head += f" {fa['size_mm'][0]}×{fa['size_mm'][1]}mm"
    elif shape == "cells":
        head += f" {fa.get('cells')} マス（各 {fa['cell_mm'][0]}×{fa['cell_mm'][1]} mm。1 字ずつ書くか、{{\"as\": \"text\"}} で 1 つの欄に）"
        if f.get("cell_ids"):
            head += "：" + "・".join(f"［{x}］" for x in f["cell_ids"])
    elif shape == "underline":
        head += f" {fa.get('length_mm')} mm"
    elif shape in ("box", "checked"):
        head += f"「{fa.get('char', '')}」（{fa.get('char_name', '')}）"
    elif "width_mm" in fa:
        head += f" {fa['width_mm']}mm（{fa.get('width_em')}字）"
        if fa.get("glyph"):
            head += f"・「{fa['glyph']}」{fa.get('count')} 字"
    parts.append(head)
    nb = _nb_text(fa)
    if nb:
        parts.append(nb)
    if shape in _LINE_SHAPES:
        ctx = [x for x in ((f"前{_q(fa['before'], 12)}" if fa.get("before") else ""),
                           (f"後{_q(fa['after'], 12)}" if fa.get("after") else "")) if x]
        if ctx:
            parts.append("・".join(ctx))
    if fa.get("bound"):
        b = fa["bound"]
        ks = [_BOUND_JA.get(k, k) for k in str(b["kind"]).split("・")]
        parts.append(("左右とも" + ks[0] if len(ks) == 2 and ks[0] == ks[1] else "左は" + "・右は".join(ks) if len(ks) == 2 else ks[0])
                     + "まで" + (_q(b["text"], 12) if b.get("text") else ""))
    if fa.get("spaced"):
        parts.append("字間を空けた形の行")
    if fa.get("fill"):
        parts.append(f"塗り {fa['fill']}")
    if fa.get("inner_cells"):
        parts.append(f"中に小さなマス {fa['inner_cells']} 個")
    st = _size_text(f)
    if st:
        parts.append(st)
    return "｜".join(parts)


def _id_xy(fid: str):
    m = re.fullmatch(r"p\d+\.t(\d+)-(\d+)", fid)
    return (int(m[1]), int(m[2])) if m else None


def _grid_blocks(fields: list[dict]) -> tuple[list[tuple[dict, list[str]]], set[str]]:
    """同じ大きさの空のマスが隣り合うまとまり（表の格子・桁の枠）：まとまりごとに、id の作り（x・y の mm）と、列の上の字・行の左の字を
    1 つの行の組にする。列（行）の中で上（左）の字が違うマスは、まとまりに入れず、1 行ずつ示す。返すもの：([(先頭の欄, 行)], まとめた id)。"""
    cand = [f for f in fields if f.get("source") == "layout" and f.get("shape") == "cell" and not f.get("value")
            and not f.get("readonly") and _id_xy(f["id"]) and (f.get("facts") or {}).get("rect_mm")]
    by_id = {f["id"]: f for f in fields}
    groups: dict[tuple, list[dict]] = {}
    for f in cand:
        groups.setdefault(tuple(f["facts"]["size_mm"]), []).append(f)
    out, used = [], set()
    for size, fs in groups.items():
        if len(fs) < 2:
            continue
        # 隣り合い（辺を共にする：0.6 mm の内。マスを作る罫線の許し RULE_TOL 2pt ≒ 0.7 mm と同じほど）でつながるまとまり
        parent = list(range(len(fs)))

        def find(i):
            while parent[i] != i:
                parent[i] = parent[parent[i]]
                i = parent[i]
            return i
        rs = [f["facts"]["rect_mm"] for f in fs]
        for i in range(len(fs)):
            for j in range(i + 1, len(fs)):
                a, b = rs[i], rs[j]
                side = abs(a[1] - b[1]) <= 0.6 and (abs(a[2] - b[0]) <= 0.6 or abs(b[2] - a[0]) <= 0.6)
                vert = abs(a[0] - b[0]) <= 0.6 and (abs(a[3] - b[1]) <= 0.6 or abs(b[3] - a[1]) <= 0.6)
                if side or vert:
                    parent[find(i)] = find(j)
        comps: dict[int, list[dict]] = {}
        for i, f in enumerate(fs):
            comps.setdefault(find(i), []).append(f)
        for comp in comps.values():
            if len(comp) < 2:
                continue
            xs = sorted({_id_xy(f["id"])[0] for f in comp})
            ys = sorted({_id_xy(f["id"])[1] for f in comp})
            top = {x: min((f for f in comp if _id_xy(f["id"])[0] == x), key=lambda f: _id_xy(f["id"])[1]) for x in xs}
            lft = {y: min((f for f in comp if _id_xy(f["id"])[1] == y), key=lambda f: _id_xy(f["id"])[0]) for y in ys}

            def txt(f, side):
                return ((f.get("facts") or {}).get(side) or {}).get("text")
            members = [f for f in comp if txt(f, "above") == txt(top[_id_xy(f["id"])[0]], "above")
                       and txt(f, "left") == txt(lft[_id_xy(f["id"])[1]], "left")]
            if len(members) < 2:
                continue
            ids = {f["id"] for f in members}
            have = {_id_xy(i) for i in ids}
            missing = [f"p{comp[0]['page']}.t{x}-{y}" for y in ys for x in xs if (x, y) not in have]
            first = min(members, key=lambda f: (_id_xy(f["id"])[1], _id_xy(f["id"])[0]))
            pg = first["page"]
            lines = [f"- ［p{pg}.t{{x}}-{{y}}］ 空のマス {len(members)} 個（各 {size[0]}×{size[1]} mm）：横 {len(xs)} × 縦 {len(ys)} に並ぶ。"
                     f"x = {'・'.join(map(str, xs))}、y = {'・'.join(map(str, ys))}（mm）"
                     + (f"。この組み合わせの所はない（別に示すか、マスでない）：{'・'.join(missing)}" if missing else "")]
            cols = [(f"x={x}", txt(top[x], "above"), top[x]["facts"]["above"]["mm"]) for x in xs
                    if top[x]["id"] in ids and txt(top[x], "above")]
            if cols:
                lines.append("  - 列の上の字（いちばん上のマスから）：" + _runs_text(cols))
            rows = [(f"y={y}", txt(lft[y], "left"), lft[y]["facts"]["left"]["mm"]) for y in ys
                    if lft[y]["id"] in ids and txt(lft[y], "left")]
            if rows:
                lines.append("  - 行の左の字（いちばん左のマスから）：" + _runs_text(rows))
            ends = []
            for y in ys:
                row = [f for f in members if _id_xy(f["id"])[1] == y]
                last = max(row, key=lambda f: _id_xy(f["id"])[0])
                if txt(last, "right"):
                    ends.append((f"y={y}", txt(last, "right"), last["facts"]["right"]["mm"]))
            if ends:
                lines.append("  - 行の右の字（いちばん右のマスから）：" + _runs_text(ends))
            bots = []
            for x in xs:
                col = [f for f in members if _id_xy(f["id"])[0] == x]
                low = max(col, key=lambda f: _id_xy(f["id"])[1])
                if txt(low, "below"):
                    bots.append((f"x={x}", txt(low, "below"), low["facts"]["below"]["mm"]))
            if bots:
                lines.append("  - 列の下の字（いちばん下のマスから）：" + _runs_text(bots))
            combs = {f.get("alternative_to") for f in members if f.get("alternative_to")
                     and (by_id.get(f["alternative_to"]) or {}).get("shape") == "cells"}
            if combs:
                every = {f"p{pg}.m{min(_id_xy(f['id'])[0] for f in members if _id_xy(f['id'])[1] == y)}-{y}" for y in ys}
                shown = (f"［p{pg}.m{xs[0]}-{{y}}？］（y は各行の y）" if combs == every and all(c.startswith(f"p{pg}.m{xs[0]}-") for c in combs)
                         else "・".join(f"［{c}？］" for c in sorted(combs, key=lambda c: [int(v) for v in re.findall(r"\d+", c)])))
                lines.append("  - 横に続くマスの並び（1 字ずつ書く。{\"as\": \"text\"} で 1 つの欄に。並びとマスは重なるので、書けるのは"
                             "どちらか一方）：" + shown)
                used.update(combs)
            other = sorted({f["alternative_to"] for f in members if f.get("alternative_to") and f["alternative_to"] not in combs})
            if other:
                lines.append("  - 重なる所（書けるのはどちらか一方）：" + "・".join(other))
            st = _size_text(first)
            if st:
                lines.append(f"  - {st}")
            out.append((first, lines))
            used.update(ids)
    return out, used


def _runs_text(items: list[tuple[str, str, float]]) -> str:
    """列・行ごとの字：同じ字・同じ離れの続く所は「y=35〜110」のようにまとめる。"""
    out: list[list] = []
    for key, t, d in items:
        if out and out[-1][1] == t and out[-1][2] == d:
            out[-1][3] = key
        else:
            out.append([key, t, d, None])
    return "・".join(f"{a}{('〜' + b.split('=')[1]) if b else ''}{_q(t)}{d} mm" for a, t, d, b in out)


def _extras(f: dict) -> str:
    out = ""
    if f.get("alternative_to"):
        out += f"｜{f['alternative_to']} と重なる（書けるのはどちらか一方）"
    if f.get("readonly"):
        out += f"｜書けない：{f.get('why_not', '')}"
    if f.get("why") and "フォームの欄" in f["why"]:
        out += "｜" + f["why"].split("。")[-1]
    v = _value(f) if f["kind"] != "checkbox" or f.get("value") else ""
    if v:
        out += f"｜いまの値 {_md(v)}"
    return out


def _in_line(f: dict) -> str:
    """行の中の所（行ごとにまとめた行の 1 つ）：作り・測った幅・その所の前後の字。"""
    fa = f.get("facts") or {}
    shape = f.get("shape", "")
    q = "？" if f.get("agent") else ""
    head = f"［{f['id']}{q}］"
    if shape in ("box", "checked"):
        t = f"{head}{_SHAPE_JA[shape]}「{fa.get('char', '')}」{fa.get('char_name', '').split(' ')[0]}" + (
            f"・後{_q(fa['after'], 12)}" if fa.get("after") else f"・前{_q(fa['before'], 12)}" if fa.get("before") else "")
    else:
        t = f"{head}{_SHAPE_JA.get(shape, shape)} {fa.get('width_mm')}mm（{fa.get('width_em')}字）"
        if fa.get("glyph"):
            t += f"「{fa['glyph']}」{fa.get('count')} 字"
        if shape in ("lead", "trail") and fa.get("bound"):
            bd = fa["bound"]
            t += f"・{_BOUND_JA.get(bd['kind'], bd['kind'])}まで" + (_q(bd["text"], 12) if bd.get("text") else "")
        elif shape not in ("lead", "trail"):
            before, after = (fa.get("before") or "")[-6:], (fa.get("after") or "")[:6]
            t += f"{_q(before, 6) if before else '（行頭）'}〜{_q(after, 6) if after else '（行末）'}"
        # 行の上下の 1 字ぶん以上の空き（書く四角は行の帯。広く書くなら box で）
        room = fa.get("room_mm") or {}
        em = (f.get("size") or 0) * 25.4 / 72
        roomy = [f"{side}{room[k]}" for k, side in (("above", "↑"), ("below", "↓")) if room.get(k, 0) >= em > 0]
        if roomy:
            t += "・空き" + "".join(roomy) + "mm"
    return t + _extras(f)


def _field_lines(fields: list[dict]) -> list[str]:
    """入力欄の一覧。配置の所は：同じ行の中の所（字の間・行の前後の空き・字の形の空欄・箱の字）は、行ごとに 1 行（行の字は元のまま）。
    空のマス・下線の罫線などは 1 行に 1 か所。同じ大きさの空のマスが隣り合うまとまりは、id の作り（x・y）と列・行の字の組にまとめる
    （まとめた所の事実は、窓の JSON の facts に 1 つずつある。id はどれもそのまま書ける）。"""
    blocks, used = _grid_blocks(fields)
    first_of = {id(f): lines for f, lines in blocks}
    groups: dict[tuple, list[dict]] = {}
    for f in fields:
        fa = f.get("facts") or {}
        if f.get("source") == "layout" and f.get("shape") in _LINE_SHAPES and "line" in fa and f["id"] not in used:
            m = re.search(r"-(\d+)[a-z]*$", f["id"])
            groups.setdefault((f["page"], fa["line"], m[1] if m else ""), []).append(f)
    group_of = {id(f): k for k, fs in groups.items() for f in fs}
    done: set = set()
    out = []
    for f in fields:
        if id(f) in first_of:
            out += first_of[id(f)]
        if f["id"] in used:
            continue
        k = group_of.get(id(f))
        if k is not None:
            if k in done:
                continue
            done.add(k)
            fs = sorted(groups[k], key=lambda x: x["rect"][0])
            spaced = any((x.get("facts") or {}).get("spaced") for x in fs)
            out.append(f"- 行{_q(k[1], 40)}" + ("（字間を空けた形）" if spaced else "") + "：" + "・".join(_in_line(x) for x in fs))
            continue
        q = "？" if f.get("agent") else ""
        if f.get("source") == "layout":
            out.append(f"- ［{f['id']}{q}］ " + _facts_text(f) + _extras(f))
        else:
            v = _value(f)
            nb = _nb_text(f["neighbours"]) if f.get("neighbours") and not f.get("tooltip") else ""
            out.append(f"- ［{f['id']}］ {_md(f['label'])}｜{_md(_kind_text(f))}" + (f"｜周りの字 {nb}" if nb else "")
                       + (f"｜いまの値 {_md(v)}" if v else ""))
    return out


def _value(f: dict) -> str:
    v = f.get("value")
    if f["kind"] == "checkbox":
        return "☑ あり" if v else "☐ なし"
    if v in (None, ""):
        return ""
    return "、".join(map(str, v)) if isinstance(v, list) else str(v)


def _md_head(doc: dict) -> list[str]:
    out = [f"# {doc['source']['file'] or '書類'}", ""]
    out.append(f"- 元ファイルの SHA-256: `{doc['source']['sha256']}`（{doc['source']['bytes']:,} バイト）、PDF、{len(doc['pages'])} ページ")
    out.append("- この文書は読むための窓です。書き込みは `todoku fill` で元のファイルの後ろに足します（元の部分は変えません）。")
    out.append("- 欄として示していない所に書くときは、見出しの字を目印にして、その右（または下）の空いている所に書けます："
               "`{\"@1\": {\"after\": \"見出しの字\", \"value\": \"書く字\", \"page\": 1}}`（下に書くときは "
               "`\"where\": \"below\"`、右から左へ書く字の見出しの左に書くときは `\"where\": \"before\"`。同じ字がいくつもあるときは `occurrence` で何番目かを指定する）。"
               "伏せ字（「○○」）の上に書くときは `{\"@1\": {\"cover\": \"○○\", \"value\": \"書く字\"}}`"
               "（白い下地の欄で覆って書く。元の字は消さない）。")
    out.append("- 印刷された選択肢は欄として並べていません。○で囲むときは、字を指定して circle で書いてください："
               "`{\"@c1\": {\"circle\": \"囲む字\", \"near\": \"近くの見出しの字\", \"page\": 1}}`"
               "（同じ字がいくつもあるときは `occurrence`。詰まった所で楕円がほかの字に触れるときは `\"style\": \"border\"`（四角の囲み線））。")
    out.append("- 記号: " + "、".join(f"{k} は{v}" for k, v in doc["notation"].items()))
    fm = doc["form"]
    if fm["acroform_fields"]:
        out.append(f"- フォームの欄が {fm['acroform_fields']} 個あります（PDFにある欄そのもの）。")
    else:
        out.append("- フォームの欄はありません（印刷用のPDF）。入力欄は、罫線・空き・字の並びから見つけたものです。")
    xw = doc.get("xfa")
    if xw and xw.get("readable"):
        kind = "動的" if xw["kind"] == "dynamic" else "静的"
        out.append(f"- **XFA（Adobe の XML のフォーム）の{kind}なフォームです**。中身と入力欄は「XFA のフォーム」の節にあります"
                   f"（XFA の入力欄 {len(xw['fields'])} か所・書ける欄 {sum(1 for d in xw['fields'] if not d.get('why_not'))}）。")
    elif fm["xfa"]:
        out.append("- **XFA（Adobe の独自のフォーム）を含みますが、読めませんでした**"
                   + (f"（{xw['error']}）" if xw and xw.get("error") else "") + "。")
    out.append("")
    return out


def _md_page(pg: dict) -> list[str]:
    out = [f"## ページ {pg['number']}（{pg['width']:.0f}×{pg['height']:.0f}pt）", ""]
    if pg.get("image_only"):
        out += ["- **このページは画像だけです**（文字の情報がありません）。view_page でページの画像を見て読み、"
                "書く所は画像の目盛り（mm）で枠を決めて、fill_document に {\"@1\": {\"page\": ページ, \"box\": [左, 上, 右, 下], "
                "\"value\": 書く字}} で書きます", ""]
    if pg.get("unmapped"):
        out += [f"- 文字に直せない字が {pg['unmapped']} 字あります（�で示しています）", ""]
        marks = pg.get("unmapped_marks") or []
        text = pg.get("unmapped_text") or []
        if marks:
            out += [f"  - 行の頭の印（箇条の印など。後ろの文の字は欠けていません）：{sum(m['count'] for m in marks)} 字（"
                    + "・".join(f"{m['where']}「{_md(m['near'][:16])}」の前" for m in marks[:8])
                    + ("ほか" if len(marks) > 8 else "") + "）"]
        if text:
            out += [f"  - 本文の中の字：{sum(m['count'] for m in text)} 字（"
                    + "・".join(f"{m['where']}「{_md(m['near'][:16])}」" for m in text[:8])
                    + ("ほか" if len(text) > 8 else "") + "）。この字は窓では読めません。中身を推測で埋めないでください"]
        out.append("")
    if pg.get("merged_glyphs") or pg.get("stacked_columns") or pg.get("unseen_glyphs") or pg.get("slivers"):
        notes = []
        if pg.get("slivers"):
            notes.append(f"幅か高さが 4pt（書き込みのいちばん小さな字）に満たない空のマス {pg['slivers']} 個（枠と枠のすき間など）は、"
                         "どんな字も入らないので入力欄にしていません")
        if pg.get("unseen_glyphs"):
            notes.append(f"見えない大きさ（0.5pt に満たない・ページより何倍も大きい）の字 {pg['unseen_glyphs']} 字は、読み順に入れていません")
        if pg.get("merged_glyphs"):
            notes.append(f"同じ位置に重ねて描いた同じ字 {pg['merged_glyphs']} 字を、1字として読みました（太字に見せる重ね書きなど）")
        if pg.get("stacked_columns"):
            notes.append(f"横書きの字を縦に1字ずつ積んだ列 {pg['stacked_columns']} 個を、上から読む縦の行にしました")
        out += ["- " + "。".join(notes) + "。", ""]
    for fl in pg.get("read_flags") or []:
        out.append(f"- **{fl['title']}**：{fl['detail']}。位置：" + "、".join(_md(w) for w in fl["where"][:12])
                   + (f"（ほか {len(fl['where']) - 12} か所）" if len(fl["where"]) > 12 else ""))
    if pg.get("read_flags"):
        out.append("")
    fields = pg["fields"]
    if fields:
        out += [f"### 入力欄（{len(fields)} か所）", ""]
        if any(f.get("source") == "layout" for f in fields):
            out += ["（配置から見つけた所：作り・大きさ（mm）・周りの字（元のまま）とその離れ（mm）。書く所かどうかは、この事実と"
                    "読み順の字から決める。id の数は所の左・上の位置（mm）。同じ行の中の所（字間・行の前後の空き・箱の字など）は行ごとに 1 行、"
                    "空のマスは 1 か所 1 行（字でも、チェックの印 true でも書ける）。同じ大きさの空のマスが隣り合うまとまりは、"
                    "id の作り（x・y）と列・行の字の組にまとめて示す）", ""]
        out += _field_lines(fields)
        out.append("")
    out += ["### 読み順（上から）", ""]
    for line in pg["reading"]:
        out.append(f"- {line}")
    out.append("")
    pics = [p for p in pg.get("pictures", []) if not p["decorative"]]
    if pics:
        out += [f"### 画像（{len(pics)} 個）", ""]
        for p in pics:
            pos = f"{p['where']}、約 {p['size_mm'][0]}×{p['size_mm'][1]} mm" + (f"、近くの字「{_md(p['near'])}」" if p["near"] else "")
            out.append(f"- {p['id']}（{pos}）：" + (f"説明「{_md(p['alt'])}」" if p["alt"] else
                                                    "**説明（代わりの字）がありません**。中身は字として読めません"))
        out.append("")
    if pg.get("annotations"):
        out += ["### 注釈（ページに貼られたメモなど）", ""]
        for a in pg["annotations"]:
            out.append(f"- {a['kind']}: {_md(a['text']) or '（文字なし）'}"
                       + (f"（見た目に描かれた字：{_md(a['drawn'])}）" if a.get("drawn") else ""))
        out.append("")
    return out


def to_markdown(doc: dict) -> str:
    out = _md_head(doc)
    for pg in doc["pages"]:
        out += _md_page(pg)
    for _t, lines in _xfa_units(doc):
        out += lines
    return "\n".join(out)


BUDGET = 120_000


def _xfa_units(doc: dict) -> list[tuple[str, list[str]]]:
    xw = doc.get("xfa")
    if not xw or not xw.get("readable"):
        return []
    from .xfa import markdown_units

    placed = {f["id"] for pg in doc["pages"] for f in pg["fields"] if f.get("source") == "xfa"}
    return markdown_units(xw, placed)


def _span(titles: list[str]) -> str:
    return titles[0] if len(titles) == 1 else f"{titles[0]}〜{titles[-1]}"


def markdown_parts(doc: dict, budget: int = BUDGET) -> list[tuple[str, str]]:
    whole = to_markdown(doc)
    if len(whole) <= budget:
        return [("全体", whole)]
    parts: list[tuple[str, list[str]]] = []
    cur: list[str] = []
    first = last = None
    for pg in doc["pages"]:
        lines = _md_page(pg)
        size = sum(len(x) + 1 for x in lines)
        if cur and sum(len(x) + 1 for x in cur) + size > budget - 2000:
            parts.append((f"ページ {first}〜{last}" if first != last else f"ページ {first}", cur))
            cur, first = [], None
        if first is None:
            first = pg["number"]
        last = pg["number"]
        cur += lines
    if cur:
        parts.append((f"ページ {first}〜{last}" if first != last else f"ページ {first}", cur))
    # XFA の節：大きな単位（スクリプトなど）は、行の切れ目で分ける（とても長い行は、途中で分ける）
    cur, titles = [], []
    step = budget // 2
    for t, lines in _xfa_units(doc):
        for line in lines:
            for k in range(0, max(len(line), 1), step):
                piece = line[k:k + step]
                if cur and sum(len(x) + 1 for x in cur) + len(piece) > budget - 2000:
                    parts.append((_span(titles), cur))
                    cur, titles = [], []
                if not titles or titles[-1] != t:
                    titles.append(t)
                cur.append(piece)
    if cur:
        parts.append((_span(titles), cur))
    index = _md_head(doc) + ["## この窓の目次", "", f"窓が大きいので、{len(parts)} の部分に分けました。必要な部分を順に読んでください。", ""]
    for i, (title, lines) in enumerate(parts, 1):
        index.append(f"- 部分 {i}: {title}（{sum(len(x) + 1 for x in lines):,} 文字）")
    return [("目次", "\n".join(index))] + [(t, "\n".join(_md_head(doc)[:1] + [""] + lines)) for t, lines in parts]
