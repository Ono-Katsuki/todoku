"""Word の窓：本文を上から（表は行ごとに、マスを ｜ で区切って）、入力欄は［］で示す。

入力欄：
- 旧来のフォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN）とコンテンツコントロール（名前は欄の名前・タグ）
- 表の空のマス（左か上のマスの文字を見出しにする）：T{表}.R{行}.C{列}
- 段落の中の空欄（「令和　　年　　月　　日」、下線を引いた空白、「＿＿＿」）と □：P{段落}.{番号}
書き込んだ欄は、見えないしおり（_todoku_…）で囲み、書いた後も同じ名前で読めるようにする。
"""

from __future__ import annotations

import hashlib
import re

from ..xlsx.window import inline_blanks, norm_label
from .document import Cell, Control, DocxDocument, Para, Table, iter_paras

DET = "deterministic"
LAYOUT = "layout"
BOX = set("□☐")
MARK_PREFIX = "_todoku_"

KIND_JA = {"text": "文字", "inline": "文字の中の空欄", "checkbox": "チェック", "dropdown": "一覧から選ぶ", "combo": "一覧から選ぶ（自由入力も可）",
           "date": "日付", "richtext": "文字（書式つき）", "picture": "画像", "cell": "文字"}


def _underlined_blanks(p: Para) -> list[tuple[int, int]]:
    """下線を引いた空白の並び（「氏名＿＿＿」を下線つきの空白で作ったもの）。"""
    out = []
    pos = 0
    start = None
    for s in p.segs:
        if s.hidden:
            continue
        for ch in s.text:
            is_blank = s.underline and ch in " \u3000\u2002\u2003" and s.kind == "t"
            if is_blank and start is None:
                start = pos
            elif not is_blank and start is not None:
                if pos - start >= 3:
                    out.append((start, pos))
                start = None
            pos += 1
    if start is not None and pos - start >= 3:
        out.append((start, pos))
    return out


def _trailing_blank(text: str) -> list[tuple[int, int]]:
    """短い見出しの後ろに、全角の空白が続いて段落が終わるもの（「住所　　　　　　　」）。"""
    m = re.search(r"[\u3000]{4,}$", text)
    if not m:
        return []
    head = text[:m.start()].strip(" \u3000")
    if not head or len(head) > 20 or head.endswith(("。", "様", "殿", "御中")):
        return []
    return [(m.start(), m.end())]


def para_blanks(p: Para) -> list[tuple[int, int]]:
    from ..pdf.layout import ADDRESS_UNITS, pdf_blanks

    text = p.text
    spans = set(pdf_blanks(text)) | set(_underlined_blanks(p)) | set(_trailing_blank(text))
    # 住所のマス：「　　郡　市」のように住所の単位で終わる段落は、その後ろに町名・番地を書き足す
    core = text.rstrip(" \u3000")
    if p.cell is not None and core and core[-1] in ADDRESS_UNITS and len(core) <= 30 and core == text:
        spans.add((len(text), len(text)))
    kept: list[tuple[int, int]] = []
    for a, b in sorted(spans):
        if kept and a < kept[-1][1]:
            kept[-1] = (kept[-1][0], max(b, kept[-1][1]))
            continue
        kept.append((a, b))
    return kept


class Spots:
    """文書の入力欄の一覧（名前 → 欄の情報）。窓と書き込みで同じものを使う。"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc
        self.items: dict[str, dict] = {}
        self.para_index: dict[int, int] = {}
        self.ranges: dict = {}
        self.marks = doc.marks() if hasattr(doc, "marks") else _bookmarks(doc)
        self._collect()

    def _collect(self) -> None:
        doc = self.doc
        n = 0
        for _kind, _part, blocks in doc.all_blocks():
            for p in iter_paras(blocks):
                n += 1
                self.para_index[id(p)] = n
        # 1) フォームの欄・コンテンツコントロール
        used_names = set()
        for k, c in enumerate(doc.controls, 1):
            name = c.name or f"ctl{k}"
            base = name
            j = 2
            while name in used_names:
                name = f"{base}#{j}"
                j += 1
            used_names.add(name)
            self.items[name] = {"id": name, "kind": c.kind, "source": c.source, "control": c, "para": c.para,
                                "options": c.options or None, "provenance": DET}
        # 2) 表の空のマス
        tno = 0
        for _kind, part, blocks in doc.all_blocks():
            for t in _tables(blocks):
                tno += 1
                for row in t.rows:
                    for c in row:
                        if c.vmerge == "continue" or c.text() or _has_control(c):
                            continue
                        if not c.paras():
                            continue
                        label, rl, cl = _cell_label(t, c)
                        if not label:
                            continue
                        fid = f"T{tno}.R{c.row + 1}.C{c.col + 1}"
                        self.items[fid] = {"id": fid, "kind": "cell", "source": "layout", "cell": c, "para": c.paras()[0],
                                           "label": label, "row_label": rl, "column_label": cl, "provenance": LAYOUT,
                                           "shade": c.shade}
        # 3) 段落の中の空欄と □（空のマスとして欄にした段落は除く）
        cell_paras = {id(it["para"]) for it in self.items.values() if it.get("kind") == "cell"}
        for _kind, part, blocks in doc.all_blocks():
            for p in iter_paras(blocks):
                if id(p) in cell_paras:
                    continue
                if any(s.field is not None for s in p.segs) and all(s.field is not None or not s.text.strip() for s in p.segs):
                    continue
                pi = self.para_index[id(p)]
                text = p.text
                k = 0
                spans = []
                for i, ch in enumerate(text):
                    if ch in BOX and not _in_field(p, i):
                        spans.append((i, i + 1, "checkbox"))
                for a, b in para_blanks(p):
                    if not any(_in_field(p, i) for i in range(a, b)):
                        spans.append((a, b, "inline"))
                for a, b, kind in sorted(spans):
                    k += 1
                    fid = f"P{pi}.{k}"
                    if kind == "checkbox":
                        m = re.search(r"[\u3000 ]{2,}|□|☐|※", text[b:])
                        opt = text[b:b + m.start()] if m else text[b:]
                        label = norm_label(opt.strip(" \u3000")) or norm_label(text[:a].strip(" \u3000"))[-30:]
                    else:
                        label = self._label(p, a, b)
                    self.items[fid] = {"id": fid, "kind": kind, "source": "layout", "para": p, "span": (a, b),
                                       "label": label, "provenance": LAYOUT}
        # 4) 書き込み済みの欄（しおり）：同じ名前の欄が見つからなければ、しおりから作る
        for name, (p, a, b) in self.marks.items():
            if name in self.items:
                continue
            kind = "checkbox" if p.text[a:b] in ("■", "☑", "☒", "□", "☐") else ("cell" if name.startswith("T") else "inline")
            self.items[name] = {"id": name, "kind": kind, "source": "layout", "para": p, "span": (a, b), "filled": True,
                                "label": self._label(p, a, b) if kind == "inline" else "", "provenance": LAYOUT}

    def _label(self, p: Para, a: int, b: int) -> str:
        """段落の中の空欄の見出し：その空欄を［　］にした段落の文字。書き込み済みのほかの欄の値は除く（白紙のときと同じ見出し）。"""
        cuts = [(x, y) for q, x, y in self.marks.values() if q is p and not (x == a and y == b) and y > x]
        out, at = [], 0
        for x, y in sorted(cuts + [(a, b)]):
            if x < at:
                continue
            out.append(p.text[at:x])
            out.append("［　］" if (x, y) == (a, b) else "　")
            at = y
        out.append(p.text[at:])
        return norm_label("".join(out))

    def value(self, item: dict):
        """いまの値（書き込み済みなら、その値）。"""
        name = item["id"]
        if name in self.marks:
            p, a, b = self.marks[name]
            t = p.text[a:b]
            if item["kind"] == "checkbox":
                return t in ("■", "☑", "☒")
            return t
        c: Control | None = item.get("control")
        if c is not None:
            if c.kind == "checkbox":
                return c.checked
            if c.placeholder:
                return None
            t = "".join(s.text for s in c.result_segs)
            if c.kind == "dropdown" and c.source == "formfield":
                return c.value
            return t.strip("\u2002 ") if c.source == "formfield" else t
        if item["kind"] == "checkbox":
            return False
        return None


def _bookmarks(doc: DocxDocument) -> dict[str, tuple[Para, int, int]]:
    """書き込んだ欄のしおり（_todoku_…）：名前 → (段落, 始まり, 終わり)（段落の文字の位置）。"""
    out: dict[str, tuple[Para, int, int]] = {}
    for _kind, part, blocks in doc.all_blocks():
        for p in iter_paras(blocks):
            starts = {}
            pos = 0
            seg_at = {}
            for s in p.segs:
                if s.run is not None:
                    seg_at.setdefault(s.run.s, pos)
                if not s.hidden:
                    pos += len(s.text)
            node = p.node
            # 段落の中の順に、しおりと文字の位置を数える
            cur = 0
            order = []
            for n in node.iter():
                if n.local == "bookmarkStart" and (n.get("name") or "").startswith(MARK_PREFIX):
                    order.append(("s", n.get("id"), n.get("name")[len(MARK_PREFIX):], n.s))
                elif n.local == "bookmarkEnd":
                    order.append(("e", n.get("id"), None, n.s))
            if not order:
                continue
            for s_ in p.segs:
                pass
            # 文字の位置は、各 seg のバイト位置で比べる
            seg_pos = []
            pos = 0
            for s in p.segs:
                if s.hidden:
                    continue
                bpos = s.node.s if s.node is not None else (s.run.s if s.run is not None else 0)
                seg_pos.append((bpos, pos, len(s.text)))
                pos += len(s.text)

            def char_at(byte: int) -> int:
                k = 0
                for bpos, cpos, ln in seg_pos:
                    if bpos < byte:
                        k = cpos + ln
                    else:
                        break
                return k

            for kind, bid, name, byte in order:
                if kind == "s":
                    starts[bid] = (name, char_at(byte))
                elif bid in starts:
                    name, a = starts.pop(bid)
                    out[name.replace("_", ".")] = (p, a, char_at(byte))
            del cur, seg_at
    return out


def _in_field(p: Para, i: int) -> bool:
    k = 0
    for s in p.segs:
        if s.hidden:
            continue
        if k <= i < k + len(s.text):
            return s.field is not None
        k += len(s.text)
    return False


def _has_control(c: Cell) -> bool:
    return any(s.field is not None for p in c.paras() for s in p.segs)


def _tables(blocks):
    for b in blocks:
        if isinstance(b, Table):
            yield b
            for row in b.rows:
                for c in row:
                    yield from _tables(c.blocks)
        elif isinstance(b, Para) and b.boxes:
            yield from _tables(b.boxes)


def _cell_label(t: Table, c: Cell) -> tuple[str, str, str]:
    row = t.rows[c.row]
    left = None
    for x in row:
        if x.col < c.col and x.text():
            if left is None or x.col > left.col:
                left = x
    up = None
    for r in range(c.row - 1, -1, -1):
        for x in t.rows[r]:
            if x.col <= c.col < x.col + x.span and x.text():
                up = x
                break
        if up is not None:
            break
    rl = norm_label(left.text()) if left else ""
    cl = norm_label(up.text()) if up else ""
    adjacent_left = left is not None and left.col + left.span == c.col
    if rl and cl and rl != cl and not (adjacent_left and up is not None and c.row - up.row > 2):
        return f"{rl}／{cl}", rl, cl
    return rl or cl, rl, cl


# ====================================================================== 窓
def editable(doc: DocxDocument, sp: "Spots", item: dict) -> bool:
    """文書の保護のもとで、この欄に書けるか（フォームだけ・読み取り専用＋編集できる範囲）。"""
    if not doc.protection_enforced or not doc.protection:
        return True
    if doc.protection == "forms":
        return item.get("control") is not None
    if doc.protection != "readOnly" or not hasattr(doc, "xml"):
        return False
    ranges = sp.ranges.setdefault("_", {})
    p = item.get("para")
    c = item.get("control")
    part = (c.part if c is not None else (p.part if p is not None else doc.main))
    if part not in ranges:
        xml = doc.xml(part)
        starts, out = {}, []
        for m in re.finditer(rb'<(?:\w+:)?perm(Start|End)\b[^>]*?\bw?:?id="([^"]+)"[^>]*>', xml):
            if m[1] == b"Start":
                starts[m[2]] = m.start()
            elif m[2] in starts:
                out.append((starts.pop(m[2]), m.end()))
        ranges[part] = out
    if c is not None:
        poss = [sg.node.s for sg in c.result_segs if sg.node is not None] or [c.node.s]
    elif "span" in item:
        pos = p.positions()
        a = item["span"][0]
        poss = [p.segs[pos[a][0]].node.s] if a < len(pos) and p.segs[pos[a][0]].node is not None else [p.node.s]
    elif p is not None:
        poss = [sg.node.s for sg in p.segs if sg.node is not None] or [p.node.s]
    else:
        return False
    return any(a <= poss[0] <= b for a, b in ranges[part])


def build(doc: DocxDocument) -> dict:
    sp = Spots(doc)
    by_para: dict[int, list[dict]] = {}
    for it in sp.items.values():
        p = it.get("para")
        if p is not None:
            by_para.setdefault(id(p), []).append(it)
    out = {
        "todoku": "0.2",
        "kind": getattr(doc, "kind", "docx"),
        "source": {"file": doc.filename, "sha256": hashlib.sha256(doc.blob).hexdigest(), "bytes": len(doc.blob),
                   "note": "元のファイルが正本。書き込みは、書いた欄の部分だけを変える差分として適用する。"},
        "notation": {"［T1.R3.C2］": "表の空のマス（1つ目の表の3行目・2列目）", "［P12.1］": "段落の中の空欄（12番目の段落の1つ目）",
                     "［名前＝値］": "入力欄といまの値", "☐［…］ / ☑": "チェック", " ｜ ": "表のマスの区切り"},
        "protection": {"edit": doc.protection, "enforced": doc.protection_enforced},
        "macros": _macros(doc),
        "stories": [],
        "fields": [],
    }
    for kind, part, blocks in doc.all_blocks():
        lines: list[str] = []
        plain: list[str] = []
        _render_blocks(blocks, sp, by_para, lines, plain)
        if lines or kind == "body":
            out["stories"].append({"kind": kind, "part": part, "reading": lines, "_plain": plain})
    contexts = {}
    for st in out["stories"]:
        for line in st["reading"]:
            for m in re.finditer(r"［([^］＝]+)(?:＝[^］]*)?］", line):
                contexts.setdefault(m[1], line)
    for it in sp.items.values():
        c = it.get("control")
        kind = it["kind"]
        label = it.get("label") or ""
        if c is not None:
            label = c.help or c.name or _context_label(it, contexts)
        d = {"id": it["id"], "kind": "text" if kind == "cell" else kind, "label": label, "source": it["source"],
             "provenance": it["provenance"], "value": sp.value(it), "context": contexts.get(it["id"], "")}
        if it.get("options"):
            d["options"] = it["options"]
        if c is not None and c.maxlen:
            d["maxlen"] = c.maxlen
        if c is not None and c.kind == "date" and c.date_format:
            d["format"] = c.date_format
        if it.get("row_label"):
            d["row_label"] = it["row_label"]
        if it.get("column_label"):
            d["column_label"] = it["column_label"]
        if it.get("shade") and it["shade"] not in ("auto", "FFFFFF"):
            d["shaded"] = it["shade"]
        if c is not None and c.lock in ("contentLocked", "sdtContentLocked"):
            d["readonly"] = True
        if not editable(doc, sp, it):
            d["readonly"] = True
            d["protected"] = True  # 文書の保護で、編集できる範囲の外
        out["fields"].append(d)
    return out


def _context_label(it: dict, contexts: dict) -> str:
    line = contexts.get(it["id"], "")
    m = re.search(r"([^｜［］]{1,30})［" + re.escape(it["id"]), line)
    return norm_label(m[1]) if m else ""


def _render_blocks(blocks, sp: Spots, by_para, lines: list[str], plain: list[str]) -> None:
    for b in blocks:
        if isinstance(b, Para):
            t = _render_para(b, sp, by_para)
            if t.strip():
                lines.append(t.strip())
            plain.append(b.raw_text())
            if b.boxes:
                sub: list[str] = []
                _render_blocks(b.boxes, sp, by_para, sub, plain)
                for s in sub:
                    lines.append(f"（テキストボックス）{s}")
        elif isinstance(b, Table):
            for row in b.rows:
                parts = []
                for c in row:
                    if c.vmerge == "continue":
                        continue
                    cid = _cell_field(sp, c)
                    sub: list[str] = []
                    for blk in c.blocks:
                        if isinstance(blk, Para):
                            t = _render_para(blk, sp, by_para)
                            if t.strip():
                                sub.append(t.strip())
                            plain.append(blk.raw_text())
                            if blk.boxes:
                                bx: list[str] = []
                                _render_blocks(blk.boxes, sp, by_para, bx, plain)
                                sub += [f"（テキストボックス）{x}" for x in bx]
                        elif isinstance(blk, Table):
                            inner: list[str] = []
                            _render_blocks([blk], sp, by_para, inner, plain)
                            sub.append("［表］" + " ／ ".join(inner))
                    cell_text = " ".join(sub)
                    if cid is not None and not cell_text:
                        cell_text = f"［{cid}］"
                    parts.append(cell_text)
                if any(x.strip() for x in parts):
                    lines.append(" ｜ ".join(x for x in parts))


def _cell_field(sp: Spots, c: Cell) -> str | None:
    for it in sp.items.values():
        if it.get("cell") is c:
            return it["id"]
    return None


def _render_para(p: Para, sp: Spots, by_para) -> str:
    text = p.text
    items = by_para.get(id(p), [])
    edits = []
    for it in items:
        if it.get("cell") is not None:
            continue
        name = it["id"]
        if name in sp.marks:
            _p, a, b = sp.marks[name]
            v = text[a:b]
            if it["kind"] == "checkbox":
                edits.append((a, b, ("☑" if v in ("■", "☑", "☒") else "☐") + f"［{name}］"))
            else:
                edits.append((a, b, f"［{name}＝{v}］"))
            continue
        if "span" in it:
            a, b = it["span"]
            if it["kind"] == "checkbox":
                edits.append((a, b, f"☐［{name}］"))
            else:
                edits.append((a, b, f"［{name}］"))
            continue
        c: Control | None = it.get("control")
        if c is not None:
            v = sp.value(it)
            # 欄の結果の文字の範囲
            a, b = _ctl_span(p, c)
            if c.kind == "checkbox":
                edits.append((a, b, ("☑" if v else "☐") + f"［{name}］"))
            else:
                edits.append((a, b, f"［{name}＝{v}］" if v not in (None, "") else f"［{name}］"))
    edits.sort()
    out, at = [], 0
    for a, b, rep in edits:
        if a < at:
            continue
        out.append(text[at:a])
        out.append(rep)
        at = b
    out.append(text[at:])
    t = "".join(out)
    hid = p.hidden_text().strip()
    if hid:
        t += f"〔非表示の文字: {hid}〕"
    if p.label:
        t = f"{p.label} {t}" if not p.label.endswith((" ", "\t", "　")) else p.label + t
    return t.replace("\t", " ")


def _ctl_span(p: Para, c: Control) -> tuple[int, int]:
    pos = 0
    a = b = None
    for s in p.segs:
        if s.hidden:
            continue
        if s.field is c:
            if a is None:
                a = pos
            b = pos + len(s.text)
        pos += len(s.text)
    if a is None:
        return len(p.text), len(p.text)
    return a, b


# ====================================================================== Markdown
def _md(s) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


STORY_JA = {"body": "本文", "header": "ヘッダー", "footer": "フッター", "footnotes": "脚注", "endnotes": "文末脚注", "comments": "コメント"}


def _macros(doc):
    from ..macros import macro_info

    return macro_info(getattr(doc, "vba", None), doc.has_macros)


def to_markdown(doc: dict) -> str:
    out = [f"# {doc['source']['file'] or '書類'}", ""]
    out.append(f"- 元ファイルの SHA-256: `{doc['source']['sha256']}`（{doc['source']['bytes']:,} バイト）、Word")
    out.append("- この文書は読むための窓です。書き込みは `todoku fill` で、書いた欄の部分だけを変えます。")
    out.append("- 記号: " + "、".join(f"{k} は{v}" for k, v in doc["notation"].items()))
    pr = doc["protection"]
    if pr["enforced"] and pr["edit"]:
        out.append({"forms": "- **編集の制限：フォームの欄だけ書き込めます**",
                    "readOnly": "- **編集の制限：読み取り専用（編集できる範囲だけ書き込めます）**"}.get(pr["edit"], f"- 編集の制限: {pr['edit']}"))
    from ..macros import summary_line

    line = summary_line(doc["macros"])
    if line:
        out.append("- " + line.replace("ブック", "文書"))
    out.append("")
    if doc["fields"]:
        out += [f"## 入力欄（{len(doc['fields'])} か所）", "", "| 欄 | 見出し | 種類 | いまの値 |", "|---|---|---|---|"]
        for f in doc["fields"]:
            k = KIND_JA.get(f["kind"], f["kind"])
            if f.get("options"):
                k += ": " + "／".join(map(str, f["options"][:15]))
            if f.get("maxlen"):
                k += f"（{f['maxlen']} 字まで）"
            if f.get("format"):
                k += f"（書式 {f['format']}）"
            if f.get("protected"):
                k += "（文書の保護で書けない欄）"
            elif f.get("readonly"):
                k += "（変更できない欄）"
            v = f.get("value")
            vs = ("☑ あり" if v else "☐ なし") if f["kind"] == "checkbox" else ("" if v in (None, "") else str(v))
            out.append(f"| {_md(f['id'])} | {_md(f['label'])} | {_md(k)} | {_md(vs)} |")
        out.append("")
    for st in doc["stories"]:
        out += [f"## {STORY_JA.get(st['kind'], st['kind'])}（{st['part']}）", ""]
        for line in st["reading"]:
            out.append(f"- {line}")
        out.append("")
    from ..macros import markdown as macro_md

    out += macro_md(doc.get("macros"))
    return "\n".join(out)


def markdown_parts(doc: dict, budget: int = 120_000) -> list[tuple[str, str]]:
    whole = to_markdown(doc)
    if len(whole) <= budget:
        return [("全体", whole)]
    lines = whole.split("\n")
    parts = []
    cur: list[str] = []
    for line in lines:
        if cur and sum(len(x) + 1 for x in cur) + len(line) > budget - 2000:
            parts.append(cur)
            cur = []
        cur.append(line)
    if cur:
        parts.append(cur)
    head = lines[:6]
    index = head + ["", "## この窓の目次", "", f"窓が大きいので、{len(parts)} の部分に分けました。", ""]
    index += [f"- 部分 {i}: {sum(len(x) + 1 for x in p):,} 文字" for i, p in enumerate(parts, 1)]
    return [("目次", "\n".join(index))] + [(f"部分 {i}", "\n".join(head[:1] + [""] + p)) for i, p in enumerate(parts, 1)]
