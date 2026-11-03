"""HWPX に書き込む：書いた所の字だけを、本文の節の XML のバイト列への小さな差分として変える（ほかの部品はバイトのまま）。

- 表の空のマス・空の段落：段落の最後の字の並び（hp:run）の中に、hp:t を足す（その並びの文字の形のまま）
- 段落の中の空欄・字を目印にした書き込み：空欄の字を値にする。hp:t を「</hp:t>しおり<hp:t>値</hp:t>しおり<hp:t>」で分ける
  （同じ hp:run の中なので、文字の形はそのまま）。下線を引いた空欄は、Word と同じく下線の長さを保つ
- □：■（☐ は ☑）にする
- 누름틀：fieldBegin の hp:ctrl のすぐ後ろに hp:t で値を書き、2つの印のあいだの字（hp:t）を除く。fieldBegin の dirty を "1"
  にする（한글で書いた欄と同じ。dirty="0" のまま案内の字と同じ字だと、한글は案内の字として扱う）
- フォームの部品：hp:checkBtn・hp:radioBtn の value、hp:comboBox・hp:listBox の selectedValue、hp:edit の hp:text
書いた所（部品の欄を除く）は、見えないしおり（hp:ctrl の hp:bookmark。「_todoku_」＋id と「_todoku_end_」＋id）で囲み、
書いた後も同じ id で読める。字を変えた段落は、한글の行の配置の控え（hp:linesegarray）を除く：控えが残っていると、한글は
控えのとおりに行を分けて描き、字が重なる。控えのない段落は、한글が開くときに配置し直す。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..docx.fill import FillError, _as_bool, _is_text_value, _width
from ..docx.window import Spots, WriteError, box_mark, resolve_write
from ..odf.xml import El, esc_attr, esc_text, set_attr
from .document import END_PREFIX, MARK_PREFIX, HwpxDocument
from .package import L, kid, kids


@dataclass
class Patch:
    id: str
    kind: str
    value: object
    part: str = ""
    anchor: dict | None = None
    key: str = ""


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def written_of(patches) -> dict:
    return {p.id: ({"@": p.anchor, "value": p.value} if p.anchor else p.value) for p in patches}


def _prefix(el: El) -> str:
    uri = el.tag[1:el.tag.index("}")] if el.tag.startswith("{") else ""
    return el.prefix_of(uri) or "hp"


def t_xml(pre: str, value: str) -> bytes:
    """hp:t（中の改行は hp:lineBreak）。空の値は何も書かない。"""
    if value == "":
        return b""
    if "\t" in value:
        raise FillError("値にタブは書けません（HWPX のタブは幅を持つ要素なので、空白で書いてください）")
    parts = re.split(r"\r\n|\n|\r", value)
    inner = f"<{pre}:lineBreak/>".join(esc_text(x) for x in parts)
    return f"<{pre}:t>{inner}</{pre}:t>".encode("utf-8")


# ====================================================================== 書いた字の書式
HEADER = "Contents/header.xml"
# 書いた字に引き継がない文字の形（hh:charPr。KS X 6101）：取り消し線（hh:strikeout の shape）・字の色（textColor が黒でない）・
# 網かけ（shadeColor）。字の大きさ（height）・書体（hh:fontRef）はそのまま（Word の docx.fill.clean_rpr と同じ考え方）
CHAR_NAMES = {"strike": "取り消し線", "color": "字の色", "shade": "網かけ"}


def charpr_bad(el) -> list[str]:
    """文字の形（hh:charPr）のうち、消したように・目立って見えるもの（CHAR_NAMES）。"""
    out = []
    so = kid(el, "strikeout")
    if so is not None and (so.attrs.get("shape") or "NONE").upper() not in ("NONE", ""):
        out.append("strike")
    if (el.attrs.get("textColor") or "#000000").upper() not in ("#000000", "NONE"):
        out.append("color")
    if (el.attrs.get("shadeColor") or "none").upper() not in ("NONE", "#FFFFFF", ""):
        out.append("shade")
    return out


def clean_charpr(xml: bytes, new_id: str) -> bytes:
    """文字の形（hh:charPr のバイト列）の写し：id を new_id に、取り消し線を「なし」、字の色を黒、網かけを「なし」にしたもの。
    ほかのバイトは同じ。証明（hwpx/proof.py）も同じ関数で、足してよい形を作る。"""
    m = re.match(rb"<[^>]*>", xml)
    head, body = m[0], xml[m.end():]
    head = re.sub(rb'(\bid=")[^"]*"', lambda x: x[1] + new_id.encode() + b'"', head, count=1)
    head = re.sub(rb'(\btextColor=")[^"]*"', rb'\g<1>#000000"', head, count=1)
    head = re.sub(rb'(\bshadeColor=")[^"]*"', rb'\g<1>none"', head, count=1)
    body = re.sub(rb'(<(?:[\w.-]+:)?strikeout\b[^>]*?\bshape=")[^"]*"', rb'\g<1>NONE"', body, count=1)
    return head + body


def _run_of(el):
    cur = el
    while cur is not None and L(cur) != "run":
        cur = cur.parent
    return cur


def _place(el: El):
    """要素が hp:t の中か（("t", hp:t)）、hp:run の直下の部品か（("run", その部品)）。"""
    cur = el
    while cur is not None and cur.parent is not None:
        loc = L(cur.parent)
        if loc == "t":
            return "t", cur.parent
        if loc == "run":
            return "run", cur
        cur = cur.parent
    return None, None


class HwpxWriter:
    def __init__(self, doc: HwpxDocument, force: bool = False):
        self.doc = doc
        self.force = force
        self.sp = Spots(doc)
        self.edits: dict[str, list[tuple[int, int, bytes]]] = {}
        self.patches: list[Patch] = []
        self.warnings: list[str] = []
        self.marks_used = set(doc.marks())
        self.relaid: set[int] = set()
        self.clean_ids: dict[str, str] = {}  # 元の文字の形の id → 書いた字の形（取り消し線などを除いた写し）の id
        self._charprs = None

    def charprs(self) -> dict:
        if self._charprs is None:
            self._charprs = {}
            if self.doc.pkg.has(HEADER):
                for el in self.doc.pkg.tree(HEADER).iter():
                    if L(el) == "charPr":
                        self._charprs[el.attrs.get("id", "")] = el
        return self._charprs

    def value_xml(self, item: dict, run, pre: str, value: str, inside_t: bool) -> bytes:
        """書く値の hp:t。その字の並び（hp:run）の文字の形に取り消し線・字の色・網かけがあれば、それを除いた文字の形の字の並びに
        分けて書く（inside_t：hp:t を閉じた所に入れるか。どちらでも、hp:run の直下に置く）。"""
        body = t_xml(pre, value)
        cp = run.attrs.get("charPrIDRef", "0") if run is not None else None
        el = self.charprs().get(cp) if cp is not None else None
        bad = charpr_bad(el) if el is not None else []
        if not bad or not body:
            return body
        if cp not in self.clean_ids:
            nums = [int(k) for k in self.charprs() if k.isdigit()] + [int(v) for v in self.clean_ids.values()]
            self.clean_ids[cp] = str(max(nums, default=-1) + 1)
        self.warnings.append(f"{item['id']}: 書く所に残っていた書式（{'・'.join(CHAR_NAMES[x] for x in bad)}）は、書いた字に引き継がず、"
                             "それを除いた文字の形（字の大きさ・書体はそのまま）で書きました")
        data = self.doc.pkg.xml(self.part_of_run)
        tag = data[run.s:run.se]
        r_open = tag[:-2].rstrip() + b">" if run.se == run.e else tag
        return (f"</{pre}:run><{pre}:run charPrIDRef=\"{self.clean_ids[cp]}\">".encode() + body + f"</{pre}:run>".encode() + r_open)

    def put(self, part: str, s: int, e: int, b: bytes) -> None:
        for a, z, _x in self.edits.get(part, []):
            if (s < z and a < e) or (s == e == a == z):
                raise FillError("同じ所に2つの書き込みが重なりました")
            if s == e and a < s < z:
                raise FillError("同じ所に2つの書き込みが重なりました")
        self.edits.setdefault(part, []).append((s, e, b))

    def relayout(self, p) -> None:
        """字を変えた段落の、行の配置の控え（hp:linesegarray）を除く。"""
        if id(p.el) in self.relaid:
            return
        self.relaid.add(id(p.el))
        ls = kid(p.el, "linesegarray")
        if ls is not None:
            self.put(p.part, ls.s, ls.e, b"")

    def bookmark(self, pre: str, fid: str) -> tuple[bytes, bytes]:
        return (f'<{pre}:ctrl><{pre}:bookmark name="{esc_attr(MARK_PREFIX + fid)}"/></{pre}:ctrl>'.encode("utf-8"),
                f'<{pre}:ctrl><{pre}:bookmark name="{esc_attr(END_PREFIX + fid)}"/></{pre}:ctrl>'.encode("utf-8"))

    def check(self, item: dict, p) -> None:
        if self.force:
            return
        c = p.cell if p is not None else None
        if (p is not None and getattr(p, "protected", False)) or (c is not None and getattr(c, "protected", False)):
            raise FillError(f"{item['id']}: 保護したマス（または保護した所）の中です（書き込めません）")

    # ------------------------------------------------------------------ 段落の字
    def replace_span(self, item: dict, p, a: int, b: int, value: str, keep_width: bool = False) -> None:
        self.check(item, p)
        part = p.part
        pre = _prefix(p.el)
        data = self.doc.pkg.xml(part)
        bm = self.bookmark(pre, item["id"])
        pad = ""
        if keep_width and a < b:
            room = _width(p.text[a:b]) - _width(value)
            pad = " " * room if room > 0 else ""
        self.part_of_run = part

        def body_at(el) -> bytes:
            return bm[0] + self.value_xml(item, _run_of(el), pre, value, True) + bm[1] + t_xml(pre, pad)

        pos = p.positions()
        if a == b:
            if not pos:
                self.para_end(item, p, value)
                return
            i, off = pos[a - 1] if a > 0 else pos[0]
            self._insert(p.segs[i].node, off, a > 0, body_at, pre, data, part, item)
            self.relayout(p)
            return
        if b > len(pos):
            raise FillError(f"{item['id']}: 空欄の位置が見つかりません")
        order: list[int] = []
        ranges: dict[int, list[int]] = {}
        for k in range(a, b):
            i, off = pos[k]
            if i not in ranges:
                order.append(i)
                ranges[i] = []
            ranges[i].append(off)
        first = True
        for i in order:
            sg = p.segs[i]
            src = sg.node
            k0, k1 = min(ranges[i]), max(ranges[i]) + 1
            if src is None or sg.field is not None:
                raise FillError(f"{item['id']}: 空欄が、入力の欄・字の欄（日付など）・図形などの中にかかっています")
            if src.kind == "txt":
                s, e, t_el = src.starts[k0], src.ends[k1 - 1], src.el
            elif src.kind == "elx":
                s, e, t_el = src.s, src.e, src.el.parent
            else:
                raise FillError(f"{item['id']}: 空欄が、書き換えられない部分（{src.kind}）にかかっています")
            if first:
                self.put(part, s, e, f"</{pre}:t>".encode() + body_at(t_el) + data[t_el.s:t_el.se])
            else:
                self.put(part, s, e, b"")
            first = False
        self.relayout(p)

    def _insert(self, src, off: int, after: bool, body_at, pre: str, data: bytes, part: str, item: dict) -> None:
        if src is None:
            raise FillError(f"{item['id']}: 書き足す位置が見つかりません")
        if src.kind == "txt":
            at = src.ends[off] if after else src.starts[off]
            self.put(part, at, at, f"</{pre}:t>".encode() + body_at(src.el) + data[src.el.s:src.el.se])
            return
        where, el = _place(src.el) if src.el is not None else (None, None)
        if where == "t":
            at = src.e if after else src.s
            self.put(part, at, at, f"</{pre}:t>".encode() + body_at(el) + data[el.s:el.se])
        elif where == "run":
            at = el.e if after else el.s
            self.put(part, at, at, body_at(el))
        else:
            raise FillError(f"{item['id']}: 書き足す位置が、書き換えられない部分にあります")

    def para_end(self, item: dict, p, value: str) -> None:
        """段落の終わりに字を足す（空の段落・空のマス）。"""
        self.check(item, p)
        part = p.part
        pre = _prefix(p.el)
        data = self.doc.pkg.xml(part)
        bm = self.bookmark(pre, item["id"])
        self.part_of_run = part
        runs = kids(p.el, "run")
        body = bm[0] + (self.value_xml(item, runs[-1], pre, value, False) if runs else t_xml(pre, value)) + bm[1]
        if runs:
            r = runs[-1]
            if r.se == r.e:
                tag = data[r.s:r.e]
                qname = re.match(rb"<([^\s/>]+)", tag)[1]
                self.put(part, r.s, r.e, tag[:-2].rstrip() + b">" + body + b"</" + qname + b">")
            else:
                self.put(part, r.es, r.es, body)
        else:
            ls = kid(p.el, "linesegarray")
            at = ls.s if ls is not None else p.el.es
            cp0 = self.charprs().get("0")
            if cp0 is not None and charpr_bad(cp0):
                raise FillError(f"{item['id']}: 字の並びのない段落に、既定の文字の形（0）で書くと、"
                                f"{'・'.join(CHAR_NAMES[x] for x in charpr_bad(cp0))}が付きます（書きません）")
            self.put(part, at, at, f'<{pre}:run charPrIDRef="0">'.encode() + body + f"</{pre}:run>".encode())
        self.relayout(p)

    # ------------------------------------------------------------------ 누름틀・フォームの部品
    def control(self, item: dict, value) -> None:
        c = item["control"]
        p = c.para
        if c.lock == "contentLocked" and not self.force:
            raise FillError(f"{item['id']}: 書き換えできない設定の欄・部品です")
        if p is not None:
            self.check(item, p)
        if c.source == "hwpx-clickhere":
            self.click_here(item, c, value)
        elif c.source in ("hwpx-form", "hwpx-radio"):
            self.form(item, c, value)
        else:
            raise FillError(f"{item['id']}: この部品（{c.source}）には書けません")

    def click_here(self, item: dict, c, value) -> None:
        text = _is_text_value(value)
        part = c.part
        p = c.para
        data = self.doc.pkg.xml(part)
        pre = _prefix(p.el)
        if c.end_ctrl is None:
            raise FillError(f"{item['id']}: 누름틀の終わりの印（hp:fieldEnd）が、同じ段落にありません")
        a, z = c.begin_ctrl.e, c.end_ctrl.s
        for run in kids(p.el, "run"):
            if run.e <= a or run.s >= z:
                continue
            for ch in run.els():
                if ch.s < a or ch.e > z:
                    continue
                loc = L(ch)
                if loc == "t":
                    self.put(part, ch.s, ch.e, b"")
                elif loc != "ctrl":
                    raise FillError(f"{item['id']}: 누름틀の中に、字でないもの（{loc or ch.tag}）があります（書き換えると消えるので、書きません）")
        if text:
            self.put(part, a, a, t_xml(pre, text))
        tag = data[c.node.s:c.node.se]
        new = set_attr(tag, "dirty", "1") if text else tag
        if new != tag:
            self.put(part, c.node.s, c.node.se, new)
        self.relayout(p)
        self.patches.append(Patch(item["id"], "input", text, part))

    def form(self, item: dict, c, value) -> None:
        part = c.part
        data = self.doc.pkg.xml(part)
        el = c.node
        kind = getattr(c, "form_kind", "")
        if kind == "checkBtn":
            on = _as_bool(value)
            tag = data[el.s:el.se]
            self.put(part, el.s, el.se, set_attr(tag, "value", "CHECKED" if on else "UNCHECKED"))
            self.patches.append(Patch(item["id"], "control", on, part))
            return
        if kind == "radio":
            text = _is_text_value(value)
            if text not in c.options:
                raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
            k = c.options.index(text)
            for j, r in enumerate(c.radios):
                tag = data[r.s:r.se]
                new = set_attr(tag, "value", "CHECKED" if j == k else "UNCHECKED")
                if new != tag:
                    self.put(part, r.s, r.se, new)
            self.patches.append(Patch(item["id"], "control", text, part))
            return
        if kind in ("comboBox", "listBox"):
            text = _is_text_value(value)
            labels, vals = c.item_labels, c.item_values
            if text in labels:
                sel = vals[labels.index(text)] or text
            elif text in vals and text:
                sel = text
            elif c.kind == "combo":
                sel = text  # 字を書ける組み合わせの欄
            else:
                raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
            tag = data[el.s:el.se]
            self.put(part, el.s, el.se, set_attr(tag, "selectedValue", sel))
            self.patches.append(Patch(item["id"], "control", text, part))
            return
        if kind == "edit":
            text = _is_text_value(value)
            if c.maxlen and len(text) > c.maxlen:
                raise FillError(f"{item['id']}: 「{text}」は {len(text)} 字です（この欄は {c.maxlen} 字まで）")
            if getattr(c, "num_only", False) and text and not text.isdigit():
                raise FillError(f"{item['id']}: この欄は数字だけ書けます（「{text}」）")
            pre = _prefix(el)
            te = kid(el, "text")
            body = esc_text(text).encode("utf-8")
            if te is None:
                fc = kid(el, "formCharPr")
                at = fc.e if fc is not None else (el.es if el.se != el.e else None)
                if at is None:
                    raise FillError(f"{item['id']}: 字を書く場所（hp:text）がありません")
                self.put(part, at, at, f"<{pre}:text>".encode() + body + f"</{pre}:text>".encode())
            elif te.se == te.e:
                tag = data[te.s:te.e]
                qname = re.match(rb"<([^\s/>]+)", tag)[1]
                self.put(part, te.s, te.e, tag[:-2].rstrip() + b">" + body + b"</" + qname + b">")
            else:
                self.put(part, te.se, te.es, body)
            self.patches.append(Patch(item["id"], "input", text, part))
            return
        raise FillError(f"{item['id']}: この部品（{kind}）には書けません")

    def header(self) -> bytes:
        """header.xml：書いた字の文字の形（clean_charpr）を、hh:charProperties の終わりに足し、itemCnt を直す。"""
        return add_charprs(self.doc.pkg.xml(HEADER), self.doc.pkg.tree(HEADER), self.clean_ids)

    # ------------------------------------------------------------------ できあがり
    def build(self) -> bytes:
        changed = {}
        if self.clean_ids:
            changed[HEADER] = self.header()
        for part, eds in self.edits.items():
            data = self.doc.pkg.xml(part)
            out = bytearray()
            at = 0
            for s, e, b in sorted(eds, key=lambda x: (x[0], x[1])):
                out += data[at:s]
                out += b
                at = e
            out += data[at:]
            changed[part] = bytes(out)
        return self.doc.pkg.rebuilt(changed)


def fill(doc: HwpxDocument, values: dict, *, force: bool = False) -> FillResult:
    w = HwpxWriter(doc, force)
    sp = w.sp
    problems = []
    used: dict = {}
    for key, value in values.items():
        item = sp.items.get(key)
        try:
            if isinstance(value, dict) and "circle" in value:
                raise FillError(f"{key}: 字を囲む指定は、fill_document（api.fill）で書いてください（字の囲み線。hwpx/circle.py）")
            if key.startswith("@"):
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    item = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                p = item["para"]
                a, b = item["span"]
                if any(p.segs[i].node is None or p.segs[i].node.kind not in ("txt", "elx") for i in
                       {p.positions()[k][0] for k in range(a, b)}):
                    raise FillError(f"{key}: 目印の字が、書き換えられない部分にかかっています")
                w.replace_span(item, p, a, b, item["value"])
                w.patches.append(Patch(item["id"], "input", item["value"], p.part, anchor=dict(value), key=key))
                continue
            if item is None:
                raise FillError(f"{key}: そのような欄はありません")
            if item.get("control") is not None:
                w.control(item, value)
                continue
            if key in w.marks_used:
                raise FillError(f"{key}: すでに書き込んだ欄です（書き直すには、元の様式に書いてください）")
            kind = item["kind"]
            p = item["para"]
            if kind == "cell":
                text = _is_text_value(value)
                if p.text and not p.text.strip(" 　 "):
                    w.replace_span(item, p, 0, len(p.text), text)
                else:
                    w.para_end(item, p, text)
                w.patches.append(Patch(key, "input", text, p.part))
            elif kind == "checkbox":
                on = _as_bool(value)
                a, b = item["span"]
                orig = p.text[a:b]
                if box_mark(orig, on) != orig:
                    w.replace_span(item, p, a, b, box_mark(orig, on))
                w.patches.append(Patch(key, "control", on, p.part))
            elif kind == "inline":
                text = _is_text_value(value)
                a, b = item["span"]
                pos = p.positions()
                blank = [p.segs[pos[i][0]] for i in range(a, b) if i < len(pos)]
                underlined = bool(blank) and all(sg.underline for sg in blank)
                w.replace_span(item, p, a, b, text, keep_width=underlined and bool(text))
                w.patches.append(Patch(key, "input", text, p.part))
            else:
                raise FillError(f"{key}: この種類の欄（{kind}）には書けません")
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    return FillResult(w.build(), w.patches, w.warnings)


def add_charprs(data: bytes, root, mapping: dict[str, str]) -> bytes:
    """header.xml に、元の文字の形の写し（clean_charpr。mapping：元の id → 新しい id）を足したバイト列。"""
    chars = next((x for x in root.iter() if L(x) == "charProperties"), None)
    if chars is None or chars.se == chars.e:
        raise FillError("文字の形の一覧（hh:charProperties）がないので、取り消し線などを除いた文字の形を足せません")
    by_cp = {x.attrs.get("id"): x for x in kids(chars, "charPr")}
    new = b"".join(clean_charpr(data[by_cp[cp].s:by_cp[cp].e], nid) for cp, nid in mapping.items())
    edits = [(chars.es, chars.es, new)]
    tag = data[chars.s:chars.se]
    m = re.search(rb'\bitemCnt="(\d+)"', tag)
    if m:
        edits.append((chars.s, chars.se, tag[:m.start(1)] + str(int(m[1]) + len(mapping)).encode() + tag[m.end(1):]))
    out = bytearray()
    at = 0
    for s_, e_, b_ in sorted(edits, key=lambda x: (x[0], x[1])):
        out += data[at:s_] + b_
        at = e_
    out += data[at:]
    return bytes(out)


def written_formatting(blob: bytes, name: str = "") -> list[dict]:
    """書いた後の文書を読み直し、書いた所（しおり _todoku_…）の字の文字の形に、取り消し線・字の色・網かけがあれば
    [{id, formatting}]（読み戻しで伝える）。"""
    doc = HwpxDocument(blob, name)
    cps = {el.attrs.get("id", ""): el for el in doc.pkg.tree(HEADER).iter() if L(el) == "charPr"} if doc.pkg.has(HEADER) else {}
    out = []
    for fid, (p, a, b) in doc.marks().items():
        pos = p.positions()
        found: list[str] = []
        for k in range(a, min(b, len(pos))):
            sg = p.segs[pos[k][0]]
            run = sg.run if getattr(sg, "run", None) is not None else (_run_of(sg.node.el) if sg.node is not None and sg.node.el is not None else None)
            el = cps.get(run.attrs.get("charPrIDRef", "0")) if run is not None else None
            for x in (charpr_bad(el) if el is not None else []):
                if CHAR_NAMES[x] not in found:
                    found.append(CHAR_NAMES[x])
        if found:
            out.append({"id": fid, "formatting": found})
    return out
