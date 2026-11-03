"""PDFの窓：ページごとに、入力欄の一覧と、上から読む順の文字（入力欄は［p1.3］）を作る。

入力欄は2通り：
- フォームの欄（AcroForm）：PDFにある欄そのもの。名前は欄の完全な名前
- 配置から見つけた欄（印刷用のPDF）：空のマス・文字の中の空欄・下線・□ など。名前は p{ページ}.{番号}
  書き込むと、同じ名前のフォームの欄をその位置に足す（元のページの描画は1バイトも変えない）
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from ..xlsx.window import norm_label
from .content import page_content
from .doc import PdfDocument
from .form import FormField, Widget, has_xfa, read_fields
from .layout import Cell, Line, PageLayout, Spot, _overlap

DET = "deterministic"
LAYOUT = "layout"

KIND_JA = {"text": "文字", "comb": "1字ずつ書く欄", "inline": "文字の中の空欄", "checkbox": "チェック", "space": "余白",
           "circle": "○で囲んで選ぶ",
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


def _render_line(ln: Line, spots: list[Spot], name, kept: list | None = None) -> str:
    """行の文字に、入力欄の印（［p1.3］）を入れる。kept があれば、窓に出した元の字（印に置き換えて消えた字を除いたもの）を足す
    （読み落としの検査が、窓に実際に出た字で数えるため）。"""
    text = ln.text
    edits = []
    for sp in sorted(spots, key=lambda x: x.at or (0, 0)):
        if sp.kind == "circle":
            # 選択肢の並びの後ろに「［p1.5：○で囲む］」の印
            last = max(ln.text.index(o) + len(o) if o in ln.text else 0 for o, _r in sp.options)
            chosen = next((f.value_text() for f, _w in (getattr(sp, "widget", None) or []) if f.name == name(sp)), None)
            edits.append((last, last, f"［{name(sp)}{'＝' + chosen if chosen else ''}：○で囲む］"))
            continue
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
            edits.append((a, b, f"{keep}［{name(sp)}{q}］"))
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
    # 配置から見つけた欄のうち、フォームの欄と重なるものは、フォームの欄を使う
    spots = []
    for sp in (lay.spots if layout else []):
        # 配置から見つけた欄に重なるフォームの欄（書き込んだ後の欄、または元からある欄）
        if sp.kind == "circle":
            # ○で囲む欄：選択肢のどれかに重なる欄（書き込みでは、選択肢ごとに囲みの欄を置く）
            hits = [(f, w) for f, w, r in ws if any(_overlap(o, r) > 0.5 for _t, o in sp.options)]
        else:
            hits = [(f, w) for f, w, r in ws if _overlap(sp.rect, r) > 0.5]
        sp.widget = hits or None
        spots.append(sp)
    used_widgets: set[int] = set()
    items: list[_Item] = []
    cell_items: dict[int, _Item] = {}
    line_spots: dict[int, list[Spot]] = {}
    for sp in spots:
        if sp.kind == "circle" and (sp.cell is None or sp.host in sp.cell.lines):
            # 行の中の選択肢（○で囲む）：選択肢の後ろに印を出す
            line_spots.setdefault(id(sp.host), []).append(sp)
            continue
        if sp.kind == "circle":
            continue
        if isinstance(sp.host, Line):
            line_spots.setdefault(id(sp.host), []).append(sp)
            for b in getattr(sp, "also", []):
                line_spots.setdefault(id(b.host), []).append(Spot(sp.kind, sp.rect, at=b.at, host=b.host, id="", checked=sp.checked) if False else _alias(sp, b))
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
        if sp.kind != "circle" and (sp.host is None or (isinstance(sp.host, Cell) and id(sp.host) not in cell_ids)):
            it = _Item(*sp.rect)
            it.spot = sp
            items.append(it)
    # フォームの欄：マスの中ならそのマスに、外なら単独で
    loose: list[tuple] = []
    # 配置から見つけた欄に書き込んだフォームの欄（名前が p{ページ}.{番号}・p{ページ}.a{番号} で、その欄に重なるもの）は、
    # 読み順の項目にしない：書き込みの前と同じ読み順・同じ番号になるように（値は、その欄の印として出す）
    own = re.compile(rf"p{number}\.a?\d+")
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
    counter = [0]
    counter_a = [0]
    names: dict[int, str] = {}

    def name(sp: Spot) -> str:
        # 規則では決められない所は、別の番号（p1.a1）：規則の欄の番号（p1.1）は、そのまま変わらない
        k = id(getattr(sp, "alias_of", sp))
        if k not in names:
            if getattr(sp, "confidence", "high") == "agent":
                counter_a[0] += 1
                names[k] = f"p{number}.a{counter_a[0]}"
            else:
                counter[0] += 1
                names[k] = f"p{number}.{counter[0]}"
        return names[k]

    # 欄にしていない空き（規則の型に合わない、行の右・マスの中の字の下などの空き）：行の後ろに「［→余白 n pt］」と出す。
    # 行の終わりに、もう欄（候補）があるときは出さない
    marks_of: dict[int, str] = {}
    emitted_marks = [0]
    for host, arrow, size in (lay.room_marks if layout else []):
        if isinstance(host, Line) and any(sp.at and sp.at[0] >= len(host.text.rstrip(" \u3000"))
                                          for sp in line_spots.get(id(host), [])):
            continue
        marks_of[id(host)] = f"［{arrow}余白 {size:.0f}pt］"

    def room_mark(host) -> str:
        m = marks_of.get(id(host))
        if m:
            emitted_marks[0] += 1
        return m or ""

    reading = []
    plain: list[str] = []
    for row in rows:
        row.sort(key=lambda it: it.x0)
        parts = []
        for it in row:
            if it.line is not None:
                t = _render_line(it.line, line_spots.get(id(it.line), []), name, plain) + room_mark(it.line)
            elif getattr(it, "spot", None) is not None:
                sp = it.spot
                mark = _filled_mark(sp, name)
                t = mark if mark is not None else f"［{name(sp)}{'？' if sp.confidence == 'agent' else ''}］"
            elif it.cell is None:
                t = ""
            else:
                c = it.cell
                pieces = []
                own = [sp for sp in spots if sp.host is c]
                circles = [sp for sp in spots if sp.kind == "circle" and sp.cell is c and sp.host not in c.lines]
                for sp in own:
                    mark = _filled_mark(sp, name)
                    if mark is not None:
                        pieces.append(mark)
                    elif sp.kind == "checkbox":
                        pieces.append(("☑" if sp.checked else "☐") + f"［{name(sp)}{'？' if sp.confidence == 'agent' else ''}］")
                    else:
                        pieces.append(f"［{name(sp)}{'？' if sp.confidence == 'agent' else ''}］")
                rendered = [_render_line(ln, line_spots.get(id(ln), []), name, plain) + room_mark(ln) for ln in c.lines]
                pieces.append(_join(rendered) + room_mark(c))
                for sp in circles:
                    nm = name(sp)
                    chosen = next((f.value_text() for f, _ in (sp.widget or []) if f.name == nm), None)
                    pieces.insert(0, "（○で囲む［" + nm + (f"＝{chosen}" if chosen else "") + "］: " +
                                  "・".join(o for o, _ in sp.options) + "）")
                t = " ".join(p for p in pieces if p)
            for _, f, w in sorted(it.marks, key=lambda m: m[0]):
                if f.name in names.values():
                    continue  # 配置から見つけた欄として、値つきで出した
                t = (t + " " if t else "") + _widget_mark(f, w)
            if t.strip():
                parts.append(t.strip())
        if parts:
            reading.append(" ｜ ".join(parts))
    # 欄の一覧（読み順の番号）
    out_fields = []
    context_of = {}
    for line in reading:
        for m in re.finditer(r"［(p\d+\.a?\d+)？?[］＝：]", line):
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
             "context": context_of.get(fid, ""), "value": True if sp.checked else None, "size": round(sp.size, 2)}
        if filled is not None:
            emitted.add(filled.name)
            d["filled"] = True
        if sp.kind == "comb":
            d["boxes"] = len(sp.boxes)
            d["box_rects"] = [[round(v, 2) for v in b] for b in sp.boxes]
        if sp.kind == "circle":
            d["options"] = [o for o, _ in sp.options]
            d["option_rects"] = [[o, [round(v, 2) for v in r]] for o, r in sp.options]
        if sp.row_label:
            d["row_label"] = sp.row_label
        if sp.column_label:
            d["column_label"] = sp.column_label
        if sp.confidence == "agent":
            d["agent"] = True  # 規則では決められない所（エージェントが読んで判断する）
            d["why"] = sp.reason
        if sp.kind_candidates:
            d["kind_candidates"] = list(sp.kind_candidates)  # 種類が形だけでは決まらない：書き方は fill の as_kind
        if sp.alternative_to is not None:
            d["alternative_to"] = name(sp.alternative_to)  # 同じ所の別の読み方（両方には書けない）
        if sp.group:
            d["group"] = [name(x) for x in sp.group]  # 縦に並ぶ同じ形の空欄（1つの欄かもしれない組）
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
        if f.options:
            d["options"] = [o[1] or o[0] for o in f.options]
            d["export_values"] = [o[0] for o in f.options]
        if f.kind == "radio":
            if f.options:
                d["options"] = [o[1] or o[0] for o in f.options]
            else:
                from .form import name_text

                d["options"] = [name_text(x.on) for x in f.widgets if x.on]
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
    out_fields.sort(key=lambda d: (0, int(d["id"].split(".")[1])) if re.fullmatch(r"p\d+\.\d+", d["id"]) else (1, 0))
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
        if txt or st in ("FreeText", "Stamp", "Text"):
            annots.append({"kind": st, "text": txt})
    marks, inword = unmapped_places(lay)
    return {"number": number, "width": round(view.width, 1), "height": round(view.height, 1), "rotate": view.rot,
            "fields": out_fields, "reading": reading, "chars": chars, "unmapped": pc.unmapped, "images": images,
            # 字に直せない字の、行の中での役割（字の並びの形で決める）：行の頭の印（箇条の印）と、それ以外（本文の字）
            "unmapped_marks": marks, "unmapped_text": inword,
            "image_only": chars == 0 and images > 0, "pictures": pictures, "problems": pc.problems, "annotations": annots,
            "_plain": plain, "_dropped": lay.dropped,
            # 窓の作り方で1つにした字：同じ位置に重ねて描いた同じ字（太字に見せる重ね書き）・縦に1字ずつ積んだ字の列
            "merged_glyphs": sum(1 for k, _t in lay.dropped if k == "dup"),
            "unseen_glyphs": sum(1 for k, _t in lay.dropped if k == "size"),  # 0.5pt に満たない字・ページの何倍もある字
            "slivers": len(lay.slivers) if layout else 0,
            "room_marks": emitted_marks[0],  # 欄にしていない空きの印（［→余白 n pt］）の数  # 字の入らない細い空のマス（幅か高さが 4pt 未満）
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
                    "size_mm": [round((x1 - x0) * 25.4 / 72), round((y1 - y0) * 25.4 / 72)], "near": near,
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
    return ord(ch) >= 0x2E80 and ch not in "［］"


def _alias(sp: Spot, b: Spot) -> Spot:
    a = Spot(sp.kind, sp.rect, sp.label, host=b.host, at=b.at, checked=sp.checked)
    a.alias_of = sp
    return a


def _hint(it: _Item) -> float:
    if it.line is not None:
        return it.line.size
    if it.cell is not None and it.cell.lines:
        return max(ln.size for ln in it.cell.lines)
    return max(6.0, (it.y1 - it.y0) / 3)


def _near_label(lay: PageLayout, r) -> str:
    """フォームの欄のそばの文字（同じマスの文字、左のマス、左の行、上の行）。"""
    cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
    c = lay.index.at(cx, cy)
    if c is not None:
        if c.lines:
            return norm_label(c.text())
        label, _rl, _cl = lay._label_for(c, c)
        if label:
            return label
    best = None
    for ln in lay.free_lines:
        if ln.vertical:
            continue
        if ln.x1 <= r[0] + 2 and ln.y0 < r[3] and ln.y1 > r[1] and r[0] - ln.x1 < 200:
            if best is None or ln.x1 > best.x1:
                best = ln
    if best is not None:
        return norm_label(best.text)
    above = None
    for ln in lay.free_lines:
        if ln.y1 <= r[1] + 1 and r[1] - ln.y1 < 30 and ln.x0 < r[2] and ln.x1 > r[0] - 20:
            if above is None or ln.y1 > above.y1:
                above = ln
    return norm_label(above.text) if above is not None else ""


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
        "notation": {"［p1.3］": "入力欄（1ページ目の3番目）",
                     "［p1.a1？］": "規則では決められない、書き込む所かもしれない所（エージェントが読んで判断する。書くなら id を指定する。"
                                    "書かなければ何も変わらない）",
                     "［→余白 12pt］": "欄にしていない、行の右の空き（↓ はマスの中の字の下、↑ は上、← は左。12pt はその幅か高さ）。"
                                      "書く所と判断したら、字を目印にして {\"@1\": {\"after\": 見出しの字, \"value\": 書く字, "
                                      "\"page\": ページ}}（下なら \"where\": \"below\"。上・左は {\"page\": …, \"box\": [左, 上, 右, 下]（mm）, "
                                      "\"value\": …}）で書く",
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
        k += ": " + "／".join(opts[:15]) + (f"（ほか{len(opts) - 15}件）" if len(opts) > 15 else "")
    if f.get("maxlen"):
        k += f"（{f['maxlen']} 字まで）"
    if f.get("readonly"):
        k += "（読み取り専用）"
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
               "`\"where\": \"below\"`。同じ字がいくつもあるときは `occurrence` で何番目かを指定する）。"
               "伏せ字（「○○」）の上に書くときは `{\"@1\": {\"cover\": \"○○\", \"value\": \"書く字\"}}`"
               "（白い下地の欄で覆って書く。元の字は消さない）。")
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
    fields = pg["fields"]
    if fields:
        out += [f"### 入力欄（{len(fields)} か所）", "", "| 欄 | 見出し | 種類 | いまの値 |", "|---|---|---|---|"]
        for f in fields:
            out.append(f"| {_md(f['id'])} | {_md(f['label'])} | {_md(_kind_text(f))} | {_md(_value(f))} |")
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
            out.append(f"- {a['kind']}: {_md(a['text']) or '（文字なし）'}")
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
