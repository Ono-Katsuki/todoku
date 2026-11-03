"""OpenDocument の文書（.odt）に書き込む：書いた所の字だけを、元の XML のバイト列への小さな差分として変える。

- 表の空のマス・空の段落：段落（text:p）の中の終わりに字を足す（段落のスタイルのまま）
- 段落の中の空欄：空欄の字を値に置き換える（その字を囲む text:span の中なので、書式もそのまま）。text:s（空白の連なり）は、
  空欄にかかる数だけ減らす。下線を引いた空欄は、Word と同じく下線の長さを保つ
- □：■（☐ は ☑）にする
- 入力の欄：text:placeholder は値に置き換え、text:text-input・text:drop-down は中の字（と選んだもの）を変える。
  Word の旧来の欄（field:fieldmark）は、印のあいだの字・チェック・選んだ番号を変える
- フォームの部品（office:forms）：form:current-value・form:current-state・form:current-selected を変える
書いた所（部品の欄を除く）は見えないしおり（text:bookmark-start/end、名前「_todoku_」＋欄の id）で囲み、書いた後も同じ id で読める。
値の中の空白は text:s、タブは text:tab、改行は text:line-break で書く（ODF の空白の決まりで、字が消えないように）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from ..docx.fill import FillError, _as_bool, _is_text_value, _width
from ..docx.window import Spots, WriteError, resolve_write
from .text import MARK_PREFIX, OdtDocument, Src
from .xml import NS, esc_attr, esc_text, set_attr, top


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


def inline_xml(value: str, tp: str) -> bytes:
    """段落の中に書く字：空白は text:s（1つでも）、タブは text:tab、改行は text:line-break。"""
    out = []
    for tok in re.split(r"( +|\t|\r\n|\n|\r)", value):
        if not tok:
            continue
        if tok[0] == " ":
            out.append(f"<{tp}:s/>" if len(tok) == 1 else f'<{tp}:s {tp}:c="{len(tok)}"/>')
        elif tok == "\t":
            out.append(f"<{tp}:tab/>")
        elif tok in ("\r\n", "\n", "\r"):
            out.append(f"<{tp}:line-break/>")
        else:
            out.append(esc_text(tok))
    return "".join(out).encode("utf-8")


def s_xml(n: int, tp: str) -> bytes:
    if n <= 0:
        return b""
    return (f"<{tp}:s/>" if n == 1 else f'<{tp}:s {tp}:c="{n}"/>').encode()


class OdtWriter:
    def __init__(self, doc: OdtDocument, force: bool = False):
        self.doc = doc
        self.force = force
        self.sp = Spots(doc)
        self.edits: dict[str, list[tuple[int, int, bytes]]] = {}
        self.patches: list[Patch] = []
        self.warnings: list[str] = []
        self.marks_used = set(doc.marks())

    def tp(self, part: str) -> str:
        root = self.doc.pkg.tree(part)
        return top(root).prefix_of(NS["text"]) or "text"

    def put(self, part: str, s: int, e: int, b: bytes) -> None:
        for a, z, _x in self.edits.get(part, []):
            if (s < z and a < e) or (s == e == a == z):
                raise FillError("同じ所に2つの書き込みが重なりました")
            if s == e and a < s < z:
                raise FillError("同じ所に2つの書き込みが重なりました")
        self.edits.setdefault(part, []).append((s, e, b))

    def bookmark(self, part: str, fid: str) -> tuple[bytes, bytes]:
        tp = self.tp(part)
        name = esc_attr(MARK_PREFIX + fid)
        return (f'<{tp}:bookmark-start {tp}:name="{name}"/>'.encode(), f'<{tp}:bookmark-end {tp}:name="{name}"/>'.encode())

    def check(self, item: dict, p) -> None:
        if self.force:
            return
        if getattr(p, "protected", False):
            raise FillError(f"{item['id']}: 保護された節（または表）の中です（書き込めません）")
        if self.doc.protection == "forms" and item.get("control") is None:
            raise FillError(f"{item['id']}: この文書は「フォームの部品だけ書ける」ように保護されています（ここは部品ではありません）")
        c = p.cell if p is not None else None
        if c is not None and getattr(c, "repeated", False):
            raise FillError(f"{item['id']}: 同じ中身を繰り返す表のマス（table:number-…-repeated）なので、1つのマスだけには書けません")

    # ------------------------------------------------------------------ 段落の字
    def replace_span(self, item: dict, p, a: int, b: int, value: str, keep_width: bool = False) -> None:
        self.check(item, p)
        part = p.part
        tp = self.tp(part)
        pos = p.positions()
        bm = self.bookmark(part, item["id"])
        if a == b:
            if not pos:
                self.para_end(item, p, value)
                return
            i, off = pos[a - 1] if a > 0 else pos[0]
            sg = p.segs[i]
            src: Src = sg.node
            after = a > 0
            if src.kind == "txt":
                at = src.ends[off] if after else src.starts[off]
                self.put(part, at, at, bm[0] + inline_xml(value, tp) + bm[1])
            elif src.kind == "s":
                k = off + 1 if after else off
                if 0 < k < src.count:
                    self.put(part, src.s, src.e, s_xml(k, tp) + bm[0] + inline_xml(value, tp) + bm[1] + s_xml(src.count - k, tp))
                else:
                    at = src.e if k >= src.count else src.s
                    self.put(part, at, at, bm[0] + inline_xml(value, tp) + bm[1])
            elif src.kind in ("tab", "br", "el", "ctl", "field", "text"):
                at = src.e if after else src.s
                self.put(part, at, at, bm[0] + inline_xml(value, tp) + bm[1])
            else:
                raise FillError(f"{item['id']}: 書き足す位置が、書き換えられない部分にあります")
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
        pad = ""
        if keep_width:
            room = _width(p.text[a:b]) - _width(value)
            pad = " " * room if room > 0 else ""
        first = True
        for i in order:
            sg = p.segs[i]
            src = sg.node
            k0, k1 = min(ranges[i]), max(ranges[i]) + 1
            new = (bm[0] + inline_xml(value, tp) + bm[1] + inline_xml(pad, tp)) if first else b""
            if src is None or sg.field is not None and sg.field.kind != "field":
                raise FillError(f"{item['id']}: 空欄が、入力の欄の中にかかっています")
            if src.kind == "txt":
                if sg.field is not None:
                    raise FillError(f"{item['id']}: 空欄が、フィールド（日付・ページ番号など）の中にかかっています")
                self.put(part, src.starts[k0], src.ends[k1 - 1], new)
            elif src.kind == "s":
                self.put(part, src.s, src.e, s_xml(k0, tp) + new + s_xml(src.count - k1, tp))
            else:
                raise FillError(f"{item['id']}: 空欄が、書き換えられない部分（{src.kind}）にかかっています")
            first = False

    def para_end(self, item: dict, p, value: str) -> None:
        """段落の終わりに字を足す（空の段落・空のマス）。"""
        self.check(item, p)
        part = p.part
        tp = self.tp(part)
        el = p.el
        data = self.doc.pkg.xml(part)
        bm = self.bookmark(part, item["id"])
        body = bm[0] + inline_xml(value, tp) + bm[1]
        if el.se == el.e:
            tag = data[el.s:el.e]
            qname = re.match(rb"<([^\s/>]+)", tag)[1]
            self.put(part, el.s, el.e, tag[:-2].rstrip() + b">" + body + b"</" + qname + b">")
        else:
            self.put(part, el.es, el.es, body)

    # ------------------------------------------------------------------ 入力の欄・フォームの部品
    def control(self, item: dict, value) -> None:
        c = item["control"]
        part = c.part
        data = self.doc.pkg.xml(part)
        tp = self.tp(part)
        el = c.node
        p = c.para
        if c.lock in ("contentLocked",) and not self.force:
            raise FillError(f"{item['id']}: 読み取り専用・使えない設定の部品です")
        if self.doc.protection == "forms" and c.source not in ("odf-form", "odf-radio") and not self.force:
            raise FillError(f"{item['id']}: この文書は「フォームの部品だけ書ける」ように保護されています")
        if p is not None and getattr(p, "protected", False) and not self.force:
            raise FillError(f"{item['id']}: 保護された節（または表）の中です（書き込めません）")
        src = c.source
        if src == "odf-placeholder":
            text = _is_text_value(value)
            bm = self.bookmark(part, item["id"])
            self.put(part, el.s, el.e, bm[0] + inline_xml(text, tp) + bm[1])
            self.patches.append(Patch(item["id"], "input", text, part))
            return
        if src == "odf-input":
            text = _is_text_value(value)
            self._content(part, el, inline_xml(text, tp), data)
            self.patches.append(Patch(item["id"], "input", text, part))
            return
        if src == "odf-dropdown":
            text = _is_text_value(value)
            if c.options and text not in c.options:
                raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
            for lab in el.findall("text:label"):
                tag = data[lab.s:lab.se]
                new = set_attr(tag, f"{tp}:current-selected", "true" if lab.get("text:value") == text else None)
                if new != tag:
                    self.put(part, lab.s, lab.se, new)
            # 中の字（text:label のあとの字）を置き換える
            labels_end = max([lab.e for lab in el.findall("text:label")], default=el.se)
            if el.se == el.e:
                raise FillError(f"{item['id']}: 選んだものを書く場所がありません")
            self.put(part, labels_end, el.es, inline_xml(text, tp))
            self.patches.append(Patch(item["id"], "control", text, part))
            return
        if src == "odf-fieldmark":
            fp = (el.tag.split("}")[0][1:])
            pre = top(self.doc.pkg.tree(part)).prefix_of(fp) or "field"
            if c.kind == "checkbox":
                on = _as_bool(value)
                self._fm_param(part, el, pre, "Checkbox_Checked", "true" if on else "false", data)
                self.patches.append(Patch(item["id"], "control", on, part))
                return
            if c.kind == "dropdown":
                text = _is_text_value(value)
                if text not in c.options:
                    raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
                self._fm_param(part, el, pre, "Dropdown_Selected", str(c.options.index(text)), data)
                self.patches.append(Patch(item["id"], "control", text, part))
                return
            text = _is_text_value(value)
            end = getattr(c, "fm_end", None)
            if end is None:
                raise FillError(f"{item['id']}: 欄の終わりの印がありません")
            self.put(part, el.e, end.s, inline_xml(text, tp))
            self.patches.append(Patch(item["id"], "input", text, part))
            return
        if src in ("odf-form", "odf-radio"):
            self.form(item, c, value, part, data)
            return
        if src == "odf-cc":
            self.content_control(item, c, value, part, data, tp)
            return
        raise FillError(f"{item['id']}: この部品（{src}）には書けません")

    def content_control(self, item: dict, c, value, part: str, data: bytes, tp: str) -> None:
        """LibreOffice のコンテンツコントロール（loext:content-control）：中の字を値にし、プレースホルダーの印
        （loext:showing-place-holder）を外す。チェックは loext:checked と印の字（loext:checked-state・unchecked-state）、
        日付は loext:date-format を通した字と loext:current-date。"""
        el = c.node
        lp = top(self.doc.pkg.tree(part)).prefix_of(NS["loext"]) or "loext"
        tag = data[el.s:el.se]
        new_tag = set_attr(tag, f"{lp}:showing-place-holder", None) if c.placeholder else tag
        if c.kind == "picture":
            raise FillError(f"{item['id']}: 画像の欄には書けません")
        if c.kind == "checkbox":
            on = _as_bool(value)
            new_tag = set_attr(new_tag, f"{lp}:checked", "true" if on else "false")
            glyph = (el.get("loext:checked-state") or "☒") if on else (el.get("loext:unchecked-state") or "☐")
            self._cc_text(item, c, part, glyph.encode("utf-8") if glyph else b"", tp)
            if new_tag != tag:
                self.put(part, el.s, el.se, new_tag)
            self.patches.append(Patch(item["id"], "control", on, part))
            return
        text = _is_text_value(value)
        if c.kind == "dropdown" and text not in c.options:
            raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
        if c.kind == "date":
            from ..docx.fill import _parse_any_date, format_word_date

            d = _parse_any_date(text)
            fmt = c.date_format or "YYYY-MM-DD"
            # LibreOffice の数の書式の日付（YYYY・MM・DD は大文字でも小文字でも同じ）を、Word の書式の字で読む
            text = format_word_date(d, fmt.replace("Y", "y").replace("D", "d"))
            new_tag = set_attr(new_tag, f"{lp}:current-date", d.strftime("%Y-%m-%dT00:00:00Z"))
        self._cc_text(item, c, part, inline_xml(text, tp), tp)
        if new_tag != tag:
            self.put(part, el.s, el.se, new_tag)
        self.patches.append(Patch(item["id"], "input", text, part))

    def _cc_text(self, item: dict, c, part: str, body: bytes, tp: str) -> None:
        """コンテンツコントロールの中の字を body にする：最初の字の片をすべて body に、ほかの字の片は消す（text:span などの
        書式の要素はそのまま）。中が空なら、終わりのタグの前に入れる。"""
        segs = [s for s in c.result_segs if s.node is not None]
        firsts = [s for s in segs if s.node.kind in ("txt", "s", "tab", "br")]
        if not firsts and c.node.se == c.node.e:
            raise FillError(f"{item['id']}: 中身を入れる所のない形（空の要素）のコンテンツコントロールです")
        if not firsts:
            at = segs[0].node.s if segs else c.node.es
            self.put(part, at, at, body)
            return
        done = False
        for s in firsts:
            src = s.node
            a, z = (src.starts[0], src.ends[-1]) if src.kind == "txt" else (src.s, src.e)
            self.put(part, a, z, b"" if done else body)
            done = True

    def _content(self, part: str, el, body: bytes, data: bytes) -> None:
        if el.se == el.e:
            tag = data[el.s:el.e]
            qname = re.match(rb"<([^\s/>]+)", tag)[1]
            self.put(part, el.s, el.e, tag[:-2].rstrip() + b">" + body + b"</" + qname + b">")
        else:
            self.put(part, el.se, el.es, body)

    def _fm_param(self, part: str, el, pre: str, name: str, value: str, data: bytes) -> None:
        for prm in el.els():
            if prm.local == "param" and prm.get("{" + prm.tag.split("}")[0][1:] + "}name") == name:
                tag = data[prm.s:prm.se]
                self.put(part, prm.s, prm.se, set_attr(tag, f"{pre}:value", value))
                return
        at = el.es if el.es != el.e else None
        if at is None:
            tag = data[el.s:el.e]
            qname = re.match(rb"<([^\s/>]+)", tag)[1]
            self.put(part, el.s, el.e, tag[:-2].rstrip() + b">" + f'<{pre}:param {pre}:name="{name}" {pre}:value="{value}"/>'.encode()
                     + b"</" + qname + b">")
        else:
            self.put(part, at, at, f'<{pre}:param {pre}:name="{name}" {pre}:value="{value}"/>'.encode())

    def form(self, item: dict, c, value, part: str, data: bytes) -> None:
        fp = top(self.doc.pkg.tree(part)).prefix_of(NS["form"]) or "form"
        el = c.node
        kind = getattr(c, "form_kind", "")
        if kind == "checkbox":
            on = _as_bool(value)
            tag = data[el.s:el.se]
            self.put(part, el.s, el.se, set_attr(tag, f"{fp}:current-state", "checked" if on else "unchecked"))
            self.patches.append(Patch(item["id"], "control", on, part))
            return
        if kind == "radio":
            text = _is_text_value(value)
            if text not in c.options:
                raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
            k = c.options.index(text)
            for j, r in enumerate(c.radios):
                tag = data[r.s:r.se]
                new = set_attr(tag, f"{fp}:current-selected", "true" if j == k else None)
                if new != tag:
                    self.put(part, r.s, r.se, new)
            self.patches.append(Patch(item["id"], "control", text, part))
            return
        if kind == "listbox":
            if getattr(c, "multiple", False):
                raise FillError(f"{item['id']}: 複数を選ぶリストには、いまは書けません")
            text = _is_text_value(value)
            if text not in c.options:
                raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
            k = c.options.index(text)
            for j, o in enumerate(el.findall("form:option")):
                tag = data[o.s:o.se]
                new = set_attr(tag, f"{fp}:current-selected", "true" if j == k else None)
                if new != tag:
                    self.put(part, o.s, o.se, new)
            self.patches.append(Patch(item["id"], "control", text, part))
            return
        text = _is_text_value(value)
        if c.kind == "date":
            from ..docx.fill import _parse_any_date

            text = _parse_any_date(text).isoformat()
        if c.maxlen and len(text) > c.maxlen:
            raise FillError(f"{item['id']}: 「{text}」は {len(text)} 字です（この欄は {c.maxlen} 字まで）")
        tag = data[el.s:el.se]
        self.put(part, el.s, el.se, set_attr(tag, f"{fp}:current-value", text))
        self.patches.append(Patch(item["id"], "input", text, part))

    # ------------------------------------------------------------------ できあがり
    def build(self) -> bytes:
        changed = {}
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


def fill(doc: OdtDocument, values: dict, *, force: bool = False) -> FillResult:
    w = OdtWriter(doc, force)
    sp = w.sp
    problems = []
    used: dict = {}
    for key, value in values.items():
        item = sp.items.get(key)
        try:
            if key.startswith("@"):
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    item = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                p = item["para"]
                a, b = item["span"]
                if any(p.segs[i].node is not None and p.segs[i].node.kind not in ("txt", "s") for i in
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
                if p.text and not p.text.strip(" 　  "):
                    w.replace_span(item, p, 0, len(p.text), text)
                else:
                    w.para_end(item, p, text)
                w.patches.append(Patch(key, "input", text, p.part))
            elif kind == "checkbox":
                on = _as_bool(value)
                a, b = item["span"]
                orig = p.text[a:b]
                if on:
                    w.replace_span(item, p, a, b, "☑" if orig == "☐" else "■")
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
