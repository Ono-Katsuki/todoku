"""Word に書き込む：書いた欄の部分だけを、元の XML のバイト列への小さな差分として変える。

- 表の空のマス：最初の段落に、段落記号の書式を引き継いだ文字を足す
- 段落の中の空欄：空白を値に置き換える（その部分だけ文字の要素を分ける）。下線を引いた空白は、下線の長さを保つ
- □：■（☐ は ☑）にする
- 旧来のフォームの欄・コンテンツコントロール：Word と同じ形で値と状態を書く
書いた所は見えないしおり（_todoku_…）で囲み、書いた後も同じ欄の名前で読み戻せるようにする。
文書の保護（フォームだけ・読み取り専用）は守る（編集できる範囲の外には書かない）。
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from .document import DocxDocument, Para, iter_paras
from .window import MARK_PREFIX, Spots
from .xmlpos import Node, escape, prefix_of


class FillError(ValueError):
    pass


@dataclass
class Patch:
    id: str
    kind: str
    value: object
    part: str = ""


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


class Editor:
    def __init__(self, xml: bytes):
        self.xml = xml
        self.edits: list[tuple[int, int, bytes]] = []
        self.pre = prefix_of(xml)

    def put(self, s: int, e: int, b: bytes) -> None:
        for a, z, _ in self.edits:
            if s < z and a < e or (s == e == a):
                if not (s == e and (s <= a or s >= z)):
                    raise FillError("同じ所に2つの書き込みが重なりました")
        self.edits.append((s, e, b))

    def build(self) -> bytes:
        out = bytearray()
        at = 0
        for s, e, b in sorted(self.edits, key=lambda x: (x[0], x[1])):
            out += self.xml[at:s]
            out += b
            at = e
        out += self.xml[at:]
        return bytes(out)


def _t(pre: str, text: str) -> bytes:
    """文字の要素。改行は改行の要素に、空白が端にあるときは xml:space=preserve を付ける。"""
    out = bytearray()
    for i, line in enumerate(text.split("\n")):
        if i:
            out += f"<{pre}:br/>".encode()
        if line:
            out += f'<{pre}:t xml:space="preserve">'.encode() + escape(line) + f"</{pre}:t>".encode()
    return bytes(out)


def _width(s: str) -> int:
    return sum(2 if ord(ch) > 0xFF and ch != " " else 1 for ch in s)


def _is_text_value(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return "" if v is None else str(v)


def _as_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    t = str(v).strip().lower()
    if t in ("1", "true", "yes", "on", "x", "✓", "✔", "☑", "■", "レ", "はい", "あり", "有", "○", "〇"):
        return True
    if t in ("0", "false", "no", "off", "", "いいえ", "なし", "無", "□", "☐"):
        return False
    raise FillError(f"チェックの欄には、はい／いいえ（true / false）で書いてください（「{v}」）")


class Filler:
    def __init__(self, doc: DocxDocument, force: bool = False):
        self.doc = doc
        self.force = force
        self.sp = Spots(doc)
        self.editors: dict[str, Editor] = {}
        self.patches: list[Patch] = []
        self.warnings: list[str] = []
        self.next_bm = 1 + max([int(m) for part in doc.parts for m in re.findall(rb'bookmark(?:Start|End)\b[^>]*?\bw?:?id="(\d+)"', doc.xml(part))] + [0])
        self.run_ops: dict[tuple[str, int], tuple[Node, list]] = {}
        self._perm_ranges()

    def ed(self, part: str) -> Editor:
        if part not in self.editors:
            self.editors[part] = Editor(self.doc.xml(part))
        return self.editors[part]

    # ------------------------------------------------------------------ 保護
    def _perm_ranges(self) -> None:
        self.ranges: dict[str, list[tuple[int, int]]] = {}
        for part in self.doc.parts:
            xml = self.doc.xml(part)
            starts = {}
            out = []
            for m in re.finditer(rb'<(?:\w+:)?perm(Start|End)\b[^>]*?\bw?:?id="([^"]+)"[^>]*>', xml):
                if m[1] == b"Start":
                    starts[m[2]] = m.start()
                elif m[2] in starts:
                    out.append((starts.pop(m[2]), m.end()))
            self.ranges[part] = out

    def check_editable(self, item: dict, part: str, pos: int) -> None:
        doc = self.doc
        if self.force or not doc.protection_enforced or not doc.protection:
            return
        if doc.protection == "forms":
            if item.get("control") is None:
                raise FillError(f"{item['id']}: この文書は「フォームの欄だけ書ける」ように保護されています（ここは欄ではありません）")
            return
        if doc.protection == "readOnly":
            if any(a <= pos <= b for a, b in self.ranges.get(part, [])):
                return
            raise FillError(f"{item['id']}: この文書は読み取り専用で保護されていて、ここは編集できる範囲の外です")
        raise FillError(f"{item['id']}: この文書は保護されていて（{doc.protection}）、書き込めません")

    # ------------------------------------------------------------------ しおり
    def bookmark(self, name: str) -> tuple[bytes, bytes]:
        pre = "w"
        bid = self.next_bm
        self.next_bm += 1
        bm = MARK_PREFIX + name.replace(".", "_").replace("#", "_")
        return (f'<{pre}:bookmarkStart {pre}:id="{bid}" {pre}:name="{bm}"/>'.encode(),
                f'<{pre}:bookmarkEnd {pre}:id="{bid}"/>'.encode())

    # ------------------------------------------------------------------ 段落の文字を置き換える
    def replace_span(self, item: dict, p: Para, a: int, b: int, value: str, keep_width: bool = False) -> None:
        """段落の文字 [a, b) を value にする予定を立てる（同じ文字の要素の中の複数の空欄は、finish でまとめて書く）。"""
        part = p.part
        pos = p.positions()
        if a == b:
            # 挿入（「市」の後ろに書き足すなど）：直前の字の後ろに入れる
            if not pos or a == 0:
                raise FillError(f"{item['id']}: 書き足す位置が見つかりません")
            i, off = pos[a - 1]
            sg = p.segs[i]
            if sg.kind != "t" or sg.node is None or sg.run is None:
                raise FillError(f"{item['id']}: 書き足す位置が、書き換えられない部分にあります")
            self.check_editable(item, part, sg.node.s)
            bm = self.bookmark(item["id"])
            self.run_ops.setdefault((part, sg.run.s), (sg.run, []))[1].append((sg.node, off + 1, off + 1, value, bm, ""))
            return
        if a >= len(pos) or b > len(pos):
            raise FillError(f"{item['id']}: 空欄の位置が見つかりません")
        segs_idx = []
        for k in range(a, b):
            i, _off = pos[k]
            if i not in segs_idx:
                segs_idx.append(i)
        for i in segs_idx:
            sg = p.segs[i]
            if sg.kind not in ("t", "sym") or sg.node is None or sg.run is None:
                raise FillError(f"{item['id']}: 空欄が、書き換えられない部分（{sg.kind}）にかかっています")
        self.check_editable(item, part, p.segs[segs_idx[0]].node.s)
        pad = ""
        if keep_width:
            room = _width(p.text[a:b]) - _width(value)
            pad = " " * room if room > 0 else ""
        bm = self.bookmark(item["id"])
        for n, i in enumerate(segs_idx):
            sg = p.segs[i]
            ks = [off for k in range(a, b) for si, off in [pos[k]] if si == i]
            k0, k1 = min(ks), max(ks) + 1
            self.run_ops.setdefault((part, sg.run.s), (sg.run, []))[1].append(
                (sg.node, k0, k1, value if n == 0 else None, bm if n == 0 else None, pad if n == 0 else ""))

    def finish_runs(self) -> None:
        """文字の要素を分けて、値としおりを入れる（1つの文字の要素ごとに1回で書く）。"""
        for (part, _s), (run, ops) in self.run_ops.items():
            ed = self.ed(part)
            pre = ed.pre
            xml = ed.xml
            rpr = run.find("rPr")
            rpr_b = xml[rpr.s:rpr.e] if rpr is not None else b""
            r_open = xml[run.s:run.se] if run.se != run.e else f"<{pre}:r>".encode()
            r_close = f"</{pre}:r>".encode()
            out = bytearray()
            buf = bytearray()

            def flush():
                nonlocal buf
                if buf:
                    out.extend(r_open + rpr_b + bytes(buf) + r_close)
                    buf = bytearray()

            by_node: dict[int, list] = {}
            for node, k0, k1, value, bm, pad in ops:
                by_node.setdefault(id(node), []).append((k0, k1, value, bm, pad))
            for kid in run.kids:
                if kid.local == "rPr":
                    continue
                todo = sorted(by_node.get(id(kid), []), key=lambda x: x[0])
                if not todo:
                    buf += xml[kid.s:kid.e]
                    continue
                if kid.local == "sym":
                    k0, k1, value, bm, _pad = todo[0]
                    sym = xml[kid.s:kid.e]
                    if value in ("■", "☑", "☒"):
                        sym = re.sub(rb'(char=")[0-9A-Fa-f]+(")', rb"\g<1>F0FE\g<2>", sym)
                    flush()
                    out.extend((bm[0] if bm else b"") + r_open + rpr_b + sym + r_close + (bm[1] if bm else b""))
                    continue
                text = kid.text
                cur = 0
                for k0, k1, value, bm, pad in todo:
                    buf += _t(pre, text[cur:k0])
                    if value is not None:
                        flush()
                        out.extend((bm[0] if bm else b"") + r_open + rpr_b + _t(pre, value) + r_close + (bm[1] if bm else b""))
                        buf += _t(pre, pad)
                    cur = k1
                buf += _t(pre, text[cur:])
            flush()
            ed.put(run.s, run.e, bytes(out))

    # ------------------------------------------------------------------ 表の空のマス
    def fill_cell(self, item: dict, value: str) -> None:
        p: Para = item["para"]
        part = p.part
        ed = self.ed(part)
        pre = ed.pre
        xml = ed.xml
        node = p.node
        self.check_editable(item, part, node.s)
        ppr = node.find("pPr")
        rpr_b = b""
        if ppr is not None:
            mark = ppr.find("rPr")
            if mark is not None:
                inner = b"".join(xml[x.s:x.e] for x in mark.kids if x.local not in ("ins", "del", "moveFrom", "moveTo"))
                rpr_b = f"<{pre}:rPr>".encode() + inner + f"</{pre}:rPr>".encode() if inner else b""
        start_bm, end_bm = self.bookmark(item["id"])
        run = start_bm + f"<{pre}:r>".encode() + rpr_b + _t(pre, value) + f"</{pre}:r>".encode() + end_bm
        if node.se == node.e:
            open_tag = xml[node.s:node.e]
            open_tag = open_tag[:-2].rstrip() + b">"
            ed.put(node.s, node.e, open_tag + run + f"</{pre}:p>".encode())
        else:
            at = ppr.e if ppr is not None else node.se
            ed.put(at, at, run)

    # ------------------------------------------------------------------ フォームの欄・コンテンツコントロール
    def fill_control(self, item: dict, value) -> None:
        c = item["control"]
        part = c.part
        ed = self.ed(part)
        pre = ed.pre
        xml = ed.xml
        inner = [sg.node.s for sg in c.result_segs if sg.node is not None]
        self.check_editable(item, part, inner[0] if inner else c.node.s)
        if c.lock in ("contentLocked", "sdtContentLocked") and not self.force:
            raise FillError(f"{item['id']}: 中身を変えられないように固定された欄です")
        if c.source == "formfield":
            ff = c.node.find("ffData")
            if c.kind == "checkbox":
                on = _as_bool(value)
                cb = ff.find("checkBox") if ff is not None else None
                if cb is None:
                    raise FillError(f"{item['id']}: チェックの欄の定義がありません")
                chk = cb.find("checked")
                v = "1" if on else "0"
                if chk is not None:
                    ed.put(chk.s, chk.e, f'<{pre}:checked {pre}:val="{v}"/>'.encode())
                else:
                    at = cb.es if cb.es != cb.e else None
                    if at is None:
                        ed.put(cb.s, cb.e, xml[cb.s:cb.e][:-2].rstrip() + b">" + f'<{pre}:checked {pre}:val="{v}"/></{pre}:checkBox>'.encode())
                    else:
                        ed.put(at, at, f'<{pre}:checked {pre}:val="{v}"/>'.encode())
                self.patches.append(Patch(item["id"], "control", on, part))
                return
            if c.kind == "dropdown":
                text = _is_text_value(value)
                if text not in c.options:
                    raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
                dd = ff.find("ddList")
                res = dd.find("result")
                i = c.options.index(text)
                if res is not None:
                    ed.put(res.s, res.e, f'<{pre}:result {pre}:val="{i}"/>'.encode())
                else:
                    ed.put(dd.se, dd.se, f'<{pre}:result {pre}:val="{i}"/>'.encode())
                self.patches.append(Patch(item["id"], "control", text, part))
                return
            text = _is_text_value(value)
            if c.kind == "date" and c.date_format:
                text = format_word_date(_parse_any_date(text), c.date_format)
            if c.maxlen and len(text) > c.maxlen:
                raise FillError(f"{item['id']}: 「{text}」は {len(text)} 字です（この欄は {c.maxlen} 字まで）")
            self._replace_result(item, c, text)
            self.patches.append(Patch(item["id"], "input", text, part))
            return
        # コンテンツコントロール
        pr = c.node.find("sdtPr")
        if c.kind == "checkbox":
            on = _as_bool(value)
            box = next((k for k in pr.kids if k.local == "checkbox"), None)
            chk = box.find("checked") if box is not None else None
            if chk is not None:
                tag = xml[chk.s:chk.e]
                tag = re.sub(rb'(:val=")[^"]*(")', rb"\g<1>" + (b"1" if on else b"0") + rb"\g<2>", tag)
                ed.put(chk.s, chk.e, tag)
            state = box.find("checkedState" if on else "uncheckedState") if box is not None else None
            ch = chr(int(state.get("val", "2612" if on else "2610"), 16)) if state is not None else ("☒" if on else "☐")
            segs = [s for s in c.result_segs if s.kind == "t" and s.node is not None]
            if segs:
                n0 = segs[0].node
                ed.put(n0.se, n0.es, escape(ch))
            self.patches.append(Patch(item["id"], "control", on, part))
            return
        text = _is_text_value(value)
        if c.kind == "dropdown" and text not in c.options:
            raise FillError(f"{item['id']}: 「{text}」は選択肢にありません。選べるのは " + "、".join(f"「{x}」" for x in c.options))
        if c.kind == "date":
            d = _parse_any_date(text)
            text = format_word_date(d, c.date_format or "yyyy/MM/dd")
            dnode = next((k for k in pr.kids if k.local == "date"), None)
            if dnode is not None:
                tag = xml[dnode.s:dnode.se]
                full = d.strftime("%Y-%m-%dT00:00:00Z").encode()
                if re.search(rb"fullDate=\"[^\"]*\"", tag):
                    tag = re.sub(rb'(fullDate=")[^"]*(")', rb"\g<1>" + full + rb"\g<2>", tag)
                else:
                    tag = tag[:-1].rstrip(b"/") + f' {pre}:fullDate="'.encode() + full + b'"' + (b"/>" if dnode.se == dnode.e else b">")
                ed.put(dnode.s, dnode.se, tag)
        ph = pr.find("showingPlcHdr")
        if ph is not None:
            ed.put(ph.s, ph.e, b"")
        self._replace_result(item, c, text)
        self.patches.append(Patch(item["id"], "input", text, part))

    def _replace_result(self, item: dict, c, text: str) -> None:
        ed = self.ed(c.part)
        pre = ed.pre
        xml = ed.xml
        ts = [s for s in c.result_segs if s.kind == "t" and s.node is not None]
        if ts:
            first = ts[0]
            n0 = first.node
            ed.put(n0.s, n0.e, f'<{pre}:t xml:space="preserve">'.encode() + escape(text) + f"</{pre}:t>".encode())
            for s in ts[1:]:
                ed.put(s.node.s, s.node.e, f'<{pre}:t xml:space="preserve"></{pre}:t>'.encode())
            # 「例：…」の書式（PlaceholderText）は外す
            for s in ts:
                rpr = s.run.find("rPr") if s.run is not None else None
                st = rpr.find("rStyle") if rpr is not None else None
                if st is not None and "placeholder" in (st.get("val") or "").lower():
                    ed.put(st.s, st.e, b"")
            return
        # 結果の文字がない：結果の終わりの前に、文字を足す
        if c.source == "formfield" and c.end is not None:
            end_run = c.end.parent
            ed.put(end_run.s, end_run.s, f"<{pre}:r>".encode() + _t(pre, text) + f"</{pre}:r>".encode())
            return
        content = c.node.find("sdtContent")
        if content is not None:
            inner_p = next((k for k in content.kids if k.local == "p"), None)
            if inner_p is not None:
                at = inner_p.find("pPr").e if inner_p.find("pPr") is not None else inner_p.se
                ed.put(at, at, f"<{pre}:r>".encode() + _t(pre, text) + f"</{pre}:r>".encode())
            else:
                ed.put(content.se, content.se, f"<{pre}:r>".encode() + _t(pre, text) + f"</{pre}:r>".encode())
            return
        raise FillError(f"{item['id']}: 値を書く場所がありません")


# ====================================================================== 日付
ERAS = [("令和", dt.date(2019, 5, 1)), ("平成", dt.date(1989, 1, 8)), ("昭和", dt.date(1926, 12, 25)), ("大正", dt.date(1912, 7, 30)),
        ("明治", dt.date(1868, 1, 25))]


def _parse_any_date(text: str) -> dt.date:
    t = str(text).strip()
    try:
        from ..xlsx.edit import parse_date_text

        d = parse_date_text(t)
        if d is not None:
            return d.date() if isinstance(d, dt.datetime) else d
    except Exception:  # noqa: BLE001
        pass
    m = re.fullmatch(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?", t)
    if m:
        return dt.date(int(m[1]), int(m[2]), int(m[3]))
    raise FillError(f"「{t}」を日付として読めません")


def format_word_date(d: dt.date, fmt: str) -> str:
    """Word の日付の書式（yyyy/MM/dd、ggge年M月d日、dddd など）。"""
    era, start = next(((n, s) for n, s in ERAS if d >= s), ("", dt.date(1, 1, 1)))
    ey = d.year - start.year + 1

    def rep(m):
        t = m[0]
        return {"yyyy": f"{d.year:04d}", "yy": f"{d.year % 100:02d}", "MMMM": d.strftime("%B"), "MMM": d.strftime("%b"),
                "MM": f"{d.month:02d}", "M": str(d.month), "dddd": "月火水木金土日"[d.weekday()] + "曜日" if "g" in fmt else d.strftime("%A"),
                "ddd": d.strftime("%a"), "dd": f"{d.day:02d}", "d": str(d.day), "ggg": era, "gg": era[:1], "g": era[:1],
                "ee": f"{ey:02d}", "e": "元" if ey == 1 and "ggge年" in fmt else str(ey), "aaaa": "月火水木金土日"[d.weekday()] + "曜日",
                "aaa": "月火水木金土日"[d.weekday()]}[t]

    return re.sub(r"yyyy|yy|MMMM|MMM|MM|M|dddd|ddd|dd|d|ggg|gg|g|ee|e|aaaa|aaa", rep, fmt)


# ====================================================================== 入口
def fill(doc: DocxDocument, values: dict, *, force: bool = False) -> FillResult:
    f = Filler(doc, force)
    problems = []
    for key, value in values.items():
        item = f.sp.items.get(key)
        try:
            if item is None:
                raise FillError(f"{key}: そのような欄はありません")
            if item.get("control") is not None:
                f.fill_control(item, value)
                continue
            if key in f.sp.marks:
                raise FillError(f"{key}: すでに書き込んだ欄です（書き直すには、元の様式に書いてください）")
            kind = item["kind"]
            if kind == "cell":
                text = _is_text_value(value)
                p = item["para"]
                if p.text and not p.text.strip(" \u3000\u2002\u2003"):
                    f.replace_span(item, p, 0, len(p.text), text)
                else:
                    f.fill_cell(item, text)
                f.patches.append(Patch(key, "input", text, item["para"].part))
            elif kind == "checkbox":
                on = _as_bool(value)
                p = item["para"]
                a, b = item["span"]
                orig = p.text[a:b]
                mark = ("☑" if orig == "☐" else "■") if on else orig
                f.replace_span(item, p, a, b, mark)
                f.patches.append(Patch(key, "control", on, p.part))
            elif kind == "inline":
                text = _is_text_value(value)
                p = item["para"]
                a, b = item["span"]
                underlined = all(s.underline for s in p.segs if s.text and not s.hidden and s.text.strip(" 　") == "")
                f.replace_span(item, p, a, b, text, keep_width=underlined and bool(text))
                f.patches.append(Patch(key, "input", text, p.part))
            else:
                raise FillError(f"{key}: この種類の欄（{kind}）には書けません")
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    f.finish_runs()
    z = doc.zip
    from ..zipraw import RawZip

    z2 = RawZip(doc.blob)
    for part, ed in f.editors.items():
        if ed.edits:
            z2.replace(part, ed.build())
    blob = z2.build() if f.editors else doc.blob
    del z
    return FillResult(blob, f.patches, f.warnings)
