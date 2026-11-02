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


def _render_line(ln: Line, spots: list[Spot], name) -> str:
    """行の文字に、入力欄の印（［p1.3］）を入れる。"""
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
        if sp.kind == "checkbox":
            edits.append((a, b, ("☑" if sp.checked else "☐") + f"［{name(sp)}］"))
        elif sp.kind == "space":
            edits.append((a, b, f" ［{name(sp)}］"))
        elif a == b == len(text):
            edits.append((a, b, f"［{name(sp)}］"))
        else:
            edits.append((a, b, f"［{name(sp)}］"))
    edits.sort(key=lambda e: (e[0], e[1]))
    out, at = [], 0
    for a, b, rep in edits:
        if a < at:
            continue
        out.append(text[at:a])
        out.append(rep)
        at = b
    out.append(text[at:])
    return "".join(out).strip()


def _filled_mark(sp: Spot, name) -> str | None:
    """配置から見つけた欄に、同じ名前のフォームの欄（書き込んだもの）があれば、その値を出す。"""
    hits = getattr(getattr(sp, "alias_of", sp), "widget", None)
    if not hits:
        return None
    nm = name(sp)
    for f, w in hits:
        if f.name == nm:
            if sp.kind == "checkbox":
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


def page_window(doc: PdfDocument, page: dict, number: int, fields: list[FormField]) -> dict:
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
    for sp in lay.spots:
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
    # フォームの欄：マスの中ならそのマスに、外なら単独で
    loose: list[tuple] = []
    for f, w, r in ws:
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
    names: dict[int, str] = {}

    def name(sp: Spot) -> str:
        k = id(getattr(sp, "alias_of", sp))
        if k not in names:
            counter[0] += 1
            names[k] = f"p{number}.{counter[0]}"
        return names[k]

    reading = []
    plain: list[str] = []
    for row in rows:
        row.sort(key=lambda it: it.x0)
        parts = []
        for it in row:
            if it.line is not None:
                plain.append(it.line.text)
            elif it.cell is not None:
                plain.extend(ln.text for ln in it.cell.lines)
            if it.line is not None:
                t = _render_line(it.line, line_spots.get(id(it.line), []), name)
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
                        pieces.append(("☑" if sp.checked else "☐") + f"［{name(sp)}］")
                    else:
                        pieces.append(f"［{name(sp)}］")
                rendered = [_render_line(ln, line_spots.get(id(ln), []), name) for ln in c.lines]
                pieces.append(_join(rendered))
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
        for m in re.finditer(r"［(p\d+\.\d+)[］＝：]", line):
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
        if sp.confidence != "high":
            d["confidence"] = sp.confidence
        if sp.kind != "checkbox":
            d["value"] = None
        if filled is not None:
            d["value"] = filled.checked if sp.kind == "checkbox" else filled.value_text()
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
    return {"number": number, "width": round(view.width, 1), "height": round(view.height, 1), "rotate": view.rot,
            "fields": out_fields, "reading": reading, "chars": chars, "unmapped": pc.unmapped, "images": images,
            "image_only": chars == 0 and images > 0, "problems": pc.problems, "annotations": annots,
            "_plain": plain, "_dropped": lay.dropped}


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
        "notation": {"［p1.3］": "入力欄（1ページ目の3番目）", "［名前＝値］": "フォームの欄といまの値", "☐［p1.5］ / ☑": "チェックの印",
                     "○［欄＝選択肢］ / ◉": "ひとつ選ぶ欄の選択肢", " ｜ ": "表のマスの区切り"},
        "form": {"acroform_fields": len(fields), "xfa": has_xfa(doc), "encrypted": doc.decryptor is not None},
        "pages": [],
    }
    for i, p in enumerate(doc.pages()):
        out["pages"].append(page_window(doc, p, i + 1, fields))
    return out


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
    if f.get("confidence") == "low":
        k += "（余白：入力欄とは限りません）"
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
    out.append("- 記号: " + "、".join(f"{k} は{v}" for k, v in doc["notation"].items()))
    fm = doc["form"]
    if fm["acroform_fields"]:
        out.append(f"- フォームの欄が {fm['acroform_fields']} 個あります（PDFにある欄そのもの）。")
    else:
        out.append("- フォームの欄はありません（印刷用のPDF）。入力欄は、罫線・空き・字の並びから見つけたものです。")
    if fm["xfa"]:
        out.append("- **XFA（Adobe の独自のフォーム）を含みます**。")
    out.append("")
    return out


def _md_page(pg: dict) -> list[str]:
    out = [f"## ページ {pg['number']}（{pg['width']:.0f}×{pg['height']:.0f}pt）", ""]
    if pg.get("image_only"):
        out += ["- **このページは画像だけです**（文字の情報がありません。画像を見て読む必要があります）", ""]
    if pg.get("unmapped"):
        out += [f"- 文字に直せない字が {pg['unmapped']} 字あります（�で示しています）", ""]
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
    return "\n".join(out)


BUDGET = 120_000


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
    index = _md_head(doc) + ["## この窓の目次", "", f"窓が大きいので、{len(parts)} の部分に分けました。必要な部分を順に読んでください。", ""]
    for i, (title, lines) in enumerate(parts, 1):
        index.append(f"- 部分 {i}: {title}（{sum(len(x) + 1 for x in lines):,} 文字）")
    return [("目次", "\n".join(index))] + [(t, "\n".join(_md_head(doc)[:1] + [""] + lines)) for t, lines in parts]
