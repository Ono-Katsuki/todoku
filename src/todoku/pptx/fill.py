"""PowerPoint（.pptx）に書き込む：書いた段落のあるスライドの部品だけを、書いた字の差分として変える（ほかの部品は、
圧縮したバイトのまま写す：zipraw）。

- 表の空のマス：マスの段落の終わりの印（a:endParaRPr）の前に、字の要素（a:r）を足す。字の書式は、その段落の終わりの印の
  書式（a:endParaRPr は a:rPr と同じ型。ECMA-376 第1部 21.1.2.2.3）を写す。空白だけのマスは、その空白を値に置き換える
- 段落の中の空欄・字を目印にした書き込み：空欄の字（a:t の中）を値に置き換える。字の要素の書式（a:rPr）は、そのまま残す
- □：■（☐ は ☑）にする
PowerPoint の字が作る字（a:fld：スライド番号・日付）・行の区切り（a:br）にかかる所には書かない。改行は書かない。
PowerPoint には、Word のしおりのような見えない印がないので、書いた所は、証明の中で書いた後のファイルを読み直して確かめる。
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from ..docx.document import iter_paras
from ..docx.fill import FillError, _as_bool, _is_text_value, written_of  # 証明に渡す、書いた所 → 値（Word と同じ）
from ..docx.proof import _expected_text
from ..docx.window import Spots, WriteError, box_mark, resolve_write
from ..docx.xmlpos import escape, parse
from ..inputcheck import GLYPH_OFF, GLYPH_ON, checkbox_state
from ..zipraw import RawZip
from .document import PptxDocument, PptxError

_BAD = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


@dataclass
class Patch:
    id: str
    kind: str
    value: object
    part: str = ""
    anchor: dict | None = None
    key: str = ""
    label: str = ""


@dataclass
class FillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def t_chars(xml: bytes, node) -> list[tuple[int, int]]:
    """a:t の字ごとの、XML の中のバイトの範囲（実体参照 &amp;・&#x…; は1字、CR LF は1字）。読み直した字と合わなければ断る。"""
    out: list[tuple[int, int]] = []
    i, end = node.se, node.es
    if node.se == node.e:
        return out
    while i < end:
        c = xml[i]
        if c == 0x3C:  # <：CDATA などは、字の位置を決めずに断る
            raise FillError("字の要素の中に CDATA などがあり、書く位置を決められません")
        if c == 0x26:  # &
            j = xml.index(b";", i)
            out.append((i, j + 1))
            i = j + 1
        elif c == 0x0D:
            j = i + 2 if xml[i + 1:i + 2] == b"\n" else i + 1
            out.append((i, j))
            i = j
        else:
            n = 1 if c < 0x80 else 2 if c < 0xE0 else 3 if c < 0xF0 else 4
            out.append((i, i + n))
            i += n
    if len(out) != len(node.text):
        raise FillError("字の要素の字と、XML の中の位置を結べませんでした")
    return out


# ====================================================================== 書いた字の書式
# 書いた字に引き継がない文字の書式（a:rPr・a:endParaRPr。ECMA-376 第1部 21.1.2.3.9 CT_TextCharacterProperties）：
# 取り消し線（strike）・字の塗り（黒でない色・塗りなし：赤・白・背景と同じ色など）・蛍光ペン（a:highlight）。
# 字の大きさ（sz）・書体（a:latin・a:ea など）はそのまま残す（Word の docx.fill.clean_rpr と同じ考え方）
_STRIKE = re.compile(rb'(\sstrike\s*=\s*")(sngStrike|dblStrike)(")')
_FILL = re.compile(rb"<(?:[\w.-]+:)?(solidFill|gradFill|pattFill|blipFill|grpFill|noFill)\b[^>]*?(?:/>|>.*?</(?:[\w.-]+:)?\1>)", re.S)
_HIGHLIGHT = re.compile(rb"<(?:[\w.-]+:)?highlight\b[^>]*?(?:/>|>.*?</(?:[\w.-]+:)?highlight>)", re.S)
_BLACK = re.compile(rb'^<(?:[\w.-]+:)?solidFill>\s*<(?:[\w.-]+:)?(?:srgbClr\s+val="000000"|schemeClr\s+val="(?:tx1|dk1)"|'
                    rb'sysClr\s+val="windowText"[^>]*?)\s*/>\s*</(?:[\w.-]+:)?solidFill>$')
_FILL_NAMES = {"strike": "取り消し線", "fill": "字の色（黒でない色・塗りなし）", "highlight": "蛍光ペン"}


def clean_rpr(rpr: bytes) -> tuple[bytes, list[str]]:
    """書いた字に付ける字の書式：取り消し線は「なし」（noStrike）に、黒でない字の塗り・蛍光ペンは除いたもの（図形・段落の既定の
    色になる）と、除いた書式の名前。証明（prove）も同じ関数で、認める書式を作る。"""
    names: list[str] = []
    if not rpr:
        return rpr, names
    m = re.match(rb"<[^>]*>", rpr)
    head, body = m[0], rpr[m.end():]
    if _STRIKE.search(head):
        head = _STRIKE.sub(rb"\g<1>noStrike\g<3>", head)
        names.append(_FILL_NAMES["strike"])

    def fill(mm: re.Match) -> bytes:
        if _BLACK.match(mm[0]):
            return mm[0]
        names.append(_FILL_NAMES["fill"])
        return b""

    body = _FILL.sub(fill, body)
    if _HIGHLIGHT.search(body):
        body = _HIGHLIGHT.sub(b"", body)
        names.append(_FILL_NAMES["highlight"])
    return head + body, list(dict.fromkeys(names))


def bad_rpr(rpr: bytes) -> list[str]:
    """字の書式に、消したように・目立って見える（clean_rpr が除く）ものがあれば、その名前。"""
    return clean_rpr(rpr)[1]


def _prefix(xml: bytes, node) -> bytes:
    m = re.match(rb"<([\w.-]+:)?p\b", xml[node.s:node.se])
    return m[1] or b"" if m else b""


class PptxWriter:
    def __init__(self, doc: PptxDocument):
        self.doc = doc
        self.edits: dict[str, list[tuple[int, int, bytes]]] = {}
        self.warnings: list[str] = []

    def note(self, fid: str, names: list[str]) -> None:
        self.warnings.append(f"{fid}: 書く所に残っていた書式（{'・'.join(names)}）は、書いた字に引き継がず、図形・段落の既定の"
                             "書式（字の大きさ・書体はそのまま）で書きました")

    def put_value(self, item: dict, part: str, xml: bytes, sg, s: int, e: int, value: str) -> None:
        """字の要素（sg の a:t）の字 [s, e) を値にする。その字の要素の書式に取り消し線・目立つ色などがあれば、字の要素を分けて、
        その書式を除いた字の要素で書く（前後の字は元の書式のまま。前・後ろに字が残らなければ、その側は分けない）。"""
        body = self.enc(value)
        t = sg.node
        r = t.parent
        rpr = r.find("rPr") if r is not None and r.local == "r" else None
        if rpr is None or not body:
            self.put(part, s, e, body)
            return
        old = xml[rpr.s:rpr.e]
        new, names = clean_rpr(old)
        if not names:
            self.put(part, s, e, body)
            return
        self.note(item["id"], names)
        pre = re.match(rb"<([\w.-]+:)?r\b", xml[r.s:r.se])[1] or b""
        r_open, t_open = xml[r.s:r.se], xml[t.s:t.se]
        close = b"</" + pre + b"t></" + pre + b"r>"
        mid = r_open + new + t_open + body + close
        head_empty = t.se == t.e or s <= t.se
        tail_empty = t.se == t.e or e >= t.es
        if t.se == t.e:
            raise FillError(f"{item['id']}: 中身のない字の要素（<a:t/>）には書けません")
        lo, hi = (r.s if head_empty else s), (r.e if tail_empty else e)
        self.put(part, lo, hi, (b"" if head_empty else close) + mid + (b"" if tail_empty else r_open + old + t_open))

    def put(self, part: str, s: int, e: int, b: bytes) -> None:
        for a, z, _x in self.edits.get(part, []):
            if (s < z and a < e) or (s == e == a == z) or (s == e and a < s < z) or (a == z and s < a < e):
                raise FillError("同じ所に2つの書き込みが重なりました")
        self.edits.setdefault(part, []).append((s, e, b))

    @staticmethod
    def enc(value: str) -> bytes:
        if any(ch in value for ch in "\r\n"):
            raise FillError("改行は書けません（PowerPoint の行の区切りは a:br の要素なので）。1行の字にしてください")
        if _BAD.search(value):
            raise FillError("XML に書けない字（制御の字）が入っています")
        return escape(value)

    def replace_span(self, item: dict, p, a: int, b: int, value: str) -> None:
        xml = self.doc.xml(p.part)
        pos = p.positions()
        if a == b:
            if not pos:
                self.para_end(item, p, value)
                return
            i, off = pos[a - 1] if a > 0 else pos[0]
            sg = p.segs[i]
            if sg.kind != "t" or sg.field is not None or sg.node is None:
                raise FillError(f"{item['id']}: 書き足す位置が、字の要素でない所（{sg.kind}）にあります")
            cm = t_chars(xml, sg.node)
            at = cm[off][1] if a > 0 else cm[off][0]
            self.put_value(item, p.part, xml, sg, at, at, value)
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
            if sg.kind != "t" or sg.node is None:
                raise FillError(f"{item['id']}: 空欄が、字の要素でない所（{sg.kind}）にかかっています")
            if sg.field is not None:
                raise FillError(f"{item['id']}: 空欄が、PowerPoint が作る字（スライド番号・日付など）にかかっています")
            cm = t_chars(xml, sg.node)
            k0, k1 = min(ranges[i]), max(ranges[i]) + 1
            if first:
                self.put_value(item, p.part, xml, sg, cm[k0][0], cm[k1 - 1][1], value)
            else:
                self.put(p.part, cm[k0][0], cm[k1 - 1][1], b"")
            first = False

    def para_end(self, item: dict, p, value: str) -> None:
        """段落の終わりの印（a:endParaRPr）の前に、その印の書式で字の要素を足す。"""
        xml = self.doc.xml(p.part)
        node = p.node
        pre = _prefix(xml, node)
        end = getattr(p, "end_rpr", None)
        rpr = b""
        if end is not None:
            raw = xml[end.s:end.e]
            rpr = re.sub(rb"^<" + re.escape(pre) + rb"endParaRPr\b", b"<" + pre + b"rPr", raw)
            rpr = re.sub(rb"</" + re.escape(pre) + rb"endParaRPr>$", b"</" + pre + b"rPr>", rpr)
            if rpr == raw:
                raise FillError(f"{item['id']}: 段落の終わりの印の書式を写せません")
            rpr, names = clean_rpr(rpr)  # 空のマスに残った取り消し線・目立つ色などは引き継がない
            if names:
                self.note(item["id"], names)
        run = b"<" + pre + b"r>" + rpr + b"<" + pre + b"t>" + self.enc(value) + b"</" + pre + b"t></" + pre + b"r>"
        if node.se == node.e:  # <a:p/>
            self.put(p.part, node.s, node.e, xml[node.s:node.e][:-2].rstrip() + b">" + run + b"</" + pre + b"p>")
        else:
            at = end.s if end is not None else node.es
            self.put(p.part, at, at, run)

    def build(self) -> bytes:
        z = RawZip(self.doc.blob)
        for part, eds in self.edits.items():
            xml = self.doc.xml(part)
            out = bytearray()
            at = 0
            for s, e, b in sorted(eds, key=lambda x: (x[0], x[1])):
                out += xml[at:s] + b
                at = e
            out += xml[at:]
            try:
                parse(bytes(out))
            except Exception as ex:  # noqa: BLE001
                raise FillError(f"{part}：書いた後の XML が正しくありません（{ex}）") from None
            z.replace(part, bytes(out))
        return z.build()


def fill(doc: PptxDocument, values: dict, *, force: bool = False) -> FillResult:
    w = PptxWriter(doc)
    sp = Spots(doc)
    patches: list[Patch] = []
    problems: list[str] = []
    used: dict = {}
    for key, value in values.items():
        try:
            if key.startswith("@"):
                if not isinstance(value, dict):
                    raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
                try:
                    item = resolve_write(sp, value, used)
                except WriteError as e:
                    raise FillError(f"{key}: {e}") from None
                a, b = item["span"]
                w.replace_span(item, item["para"], a, b, item["value"])
                patches.append(Patch(item["id"], "input", item["value"], item["para"].part, dict(value), key, item.get("label", "")))
                continue
            item = sp.items.get(key)
            if item is None:
                raise FillError(f"{key}: そのような欄はありません{sp.items.hint(key)}")
            p = item["para"]
            kind = item["kind"]
            if kind == "cell":
                text = _is_text_value(value)
                if p.text and not p.text.strip(" 　  "):
                    w.replace_span(item, p, 0, len(p.text), text)
                elif p.text:
                    raise FillError(f"{key}: マスに字があります")
                else:
                    w.para_end(item, p, text)
                patches.append(Patch(key, "input", text, p.part, key=key, label=item.get("label", "")))
            elif kind == "checkbox":
                on = _as_bool(value)
                a, b = item["span"]
                if box_mark(p.text[a:b], on) != p.text[a:b]:
                    w.replace_span(item, p, a, b, box_mark(p.text[a:b], on))
                patches.append(Patch(key, "control", on, p.part, key=key, label=item.get("label", "")))
            elif kind == "inline":
                text = _is_text_value(value)
                a, b = item["span"]
                w.replace_span(item, p, a, b, text)
                patches.append(Patch(key, "input", text, p.part, key=key, label=item.get("label", "")))
            else:
                raise FillError(f"{key}: この種類の欄（{kind}）には書けません")
        except FillError as e:
            problems.append(str(e))
    if problems:
        raise FillError("\n".join(problems))
    return FillResult(w.build(), patches, w.warnings)


# ====================================================================== 証明
@dataclass
class PptxProof:
    original_sha256: str
    new_sha256: str
    entries_total: int = 0
    entries_identical: int = 0
    paragraphs_total: int = 0
    paragraphs_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)
    readback: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors and not self.readback

    def headline(self) -> str:
        return (f"ZIP の部品 {self.entries_identical}/{self.entries_total} が圧縮したバイトまで一致。段落 {self.paragraphs_total:,} のうち、"
                f"変わったのは書いた所のある {self.paragraphs_changed} 段落だけ（その段落も、字の要素のほかはバイト一致）"
                + ("" if self.ok else f"（想定外 {len(self.unexpected) + len(self.lint_errors) + len(self.readback)} 件）"))


def _paras(doc: PptxDocument) -> list:
    return [p for _k, _part, bl in doc.all_blocks() for p in iter_paras(bl)]


def _cut(xml: bytes, spans: list[tuple[int, int]]) -> bytes:
    out, at = bytearray(), 0
    for s, e in sorted(spans):
        out += xml[at:s]
        at = e
    return bytes(out + xml[at:])


def _skeleton(xml: bytes, p) -> bytes:
    """段落の XML から、字の要素（a:r）を除いたもの（段落の書式・行の区切り・終わりの印）。中身のない段落（<a:p/>）は、
    書き込みで字の要素を足すときと同じ形（<a:p></a:p>）にそろえる（字の要素を足しただけの段落を、変わったとしない）。"""
    if p.node.se == p.node.e:
        return xml[p.node.s:p.node.e][:-2].rstrip() + b"></" + _prefix(xml, p.node) + b"p>"
    return _cut(xml[p.node.s:p.node.e], [(k.s - p.node.s, k.e - p.node.s) for k in p.node.kids if k.local == "r"])


def _formats(xml: bytes, p) -> set[bytes]:
    out = set()
    for k in p.node.kids:
        if k.local == "r":
            r = k.find("rPr")
            out.add(xml[r.s:r.e] if r is not None else b"")
    return out


def prove(original: bytes, new: bytes, written: dict) -> PptxProof:
    """書き込み処理の結果を使わずに、元の文書で書く所を決め直し、前後の 2 つのファイルを読み直して比べる。
    - ZIP：部品の並びが同じで、書いた段落のあるスライドの部品のほかは、圧縮したバイトまで一致
    - 書いた部品：書いた段落の外のバイトが一致。書いた段落は、字の要素（a:r）のほか（段落の書式・a:br・終わりの印）がバイト一致で、
      字の書式は、元の段落にあったもの（か、終わりの印の書式）だけ
    - 読み直し：段落の数が同じで、書いていない段落の字は同じ。書いた段落の字は「元の字の、書いた所だけを値に替えたもの」"""
    pr = PptxProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    za, zb = RawZip(original), RawZip(new)
    if za.names() != zb.names():
        pr.unexpected.append("ZIP の中の部品の並びが変わっています")
        return pr
    changed = []
    for name in za.names():
        pr.entries_total += 1
        ea, eb = za.get(name), zb.get(name)
        if ea.data == eb.data and za.read(name) == zb.read(name):
            pr.entries_identical += 1
        else:
            changed.append(name)
    a = PptxDocument(original)
    try:
        b = PptxDocument(new)
    except PptxError as e:
        pr.lint_errors.append(f"書き込み後の文書を読めません：{e}")
        return pr
    pa, pb = _paras(a), _paras(b)
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.readback.append(f"段落の数が変わりました（{len(pa)} → {len(pb)}）")
        return pr
    sp = Spots(a)
    index = {id(p): i for i, p in enumerate(pa)}
    targets: dict[int, list] = {}
    used: dict = {}
    for fid, value in written.items():
        if isinstance(value, dict) and "@" in value:
            try:
                it = resolve_write(sp, value["@"], used)
            except WriteError as e:
                pr.unexpected.append(f"書いた所 {fid} の目印が、元の文書で決まりません：{e}")
                continue
            if it["id"] != fid or it["value"] != value["value"]:
                pr.unexpected.append(f"書いた所 {fid} が、目印から決まる所（{it['id']}「{it['value']}」）と違います")
                continue
            value = value["value"]
        else:
            it = sp.items.get(fid)
        if it is None or it.get("para") is None or id(it["para"]) not in index:
            pr.unexpected.append(f"書いた欄 {fid} が、元の文書にありません")
            continue
        targets.setdefault(index[id(it["para"])], []).append((fid, it, value))
    allowed = {pa[i].part for i in targets}
    for part in changed:
        if part not in allowed:
            pr.unexpected.append(f"書いた段落のない部品 {part} が変わっています")
            continue
        xa, xb = a.xml(part), b.xml(part)
        sa = [(pa[i].node.s, pa[i].node.e) for i in targets if pa[i].part == part]
        sb = [(pb[i].node.s, pb[i].node.e) for i in targets if pb[i].part == part]
        if _cut(xa, sa) != _cut(xb, sb):
            pr.unexpected.append(f"{part}：書いた段落の外のバイトが変わっています")
    for i, items in targets.items():
        x, y = pa[i], pb[i]
        pr.paragraphs_changed += 1
        xa, xb = a.xml(x.part), b.xml(y.part)
        if _skeleton(xa, x) != _skeleton(xb, y):
            pr.unexpected.append(f"段落 {i + 1}：字の要素（a:r）のほか（段落の書式・行の区切り・終わりの印）が変わっています")
        fmts = _formats(xa, x) | {b""}
        end = getattr(x, "end_rpr", None)
        if end is not None:
            pre = _prefix(xa, x.node)
            raw = xa[end.s:end.e]
            raw = re.sub(rb"^<" + re.escape(pre) + rb"endParaRPr\b", b"<" + pre + b"rPr", raw)
            fmts.add(re.sub(rb"</" + re.escape(pre) + rb"endParaRPr>$", b"</" + pre + b"rPr>", raw))
        fmts |= {clean_rpr(f)[0] for f in fmts}  # 書いた字には、取り消し線・目立つ色を引き継がない（clean_rpr）
        if not _formats(xb, y) <= fmts:
            pr.unexpected.append(f"段落 {i + 1}：元の段落にない字の書式が使われています")
        want = _expected_text(x, items)
        if want is None or y.text != want:
            pr.readback.append(f"段落 {i + 1}：読み直すと「{y.text[:40]}」で、思っていた「{(want or '')[:40]}」と違います")
    for i, (x, y) in enumerate(zip(pa, pb)):
        if i not in targets and x.text != y.text:
            pr.unexpected.append(f"書いていない段落 {i + 1} の字が変わっています（「{x.text[:30]}」→「{y.text[:30]}」）")
    return pr


def _written_spans(a: PptxDocument, written: dict) -> dict[int, list]:
    """元の文書で書く所を決め直す：段落の番号 → [((初め, 終わり), 書いた字, 欄の id)]。"""
    pa = _paras(a)
    sp = Spots(a)
    index = {id(p): i for i, p in enumerate(pa)}
    used: dict = {}
    per: dict[int, list] = {}
    for fid, value in written.items():
        if isinstance(value, dict) and "@" in value:
            it = resolve_write(sp, value["@"], used)
            value = value["value"]
        else:
            it = sp.items.get(fid)
        p = it["para"]
        if it["kind"] == "cell":
            span, text = ((0, len(p.text)) if p.text else (0, 0)), str(value)
        elif it["kind"] == "checkbox":
            x, y = it["span"]
            on = checkbox_state(value, GLYPH_ON, GLYPH_OFF) is True  # 書き手と同じ読み方（inputcheck）
            span, text = (x, y), box_mark(p.text[x:y], on)
        else:
            span, text = it["span"], str(value)
        per.setdefault(index[id(p)], []).append((span, text, fid))
    return per


def written_formatting(original: bytes, new: bytes, written: dict, docs: tuple | None = None) -> list[dict]:
    """書いた後のファイルを読み直し、書いた字（元の文書で書く所を決め直し、書いた後の段落でのその字の位置）の字の書式に、
    取り消し線・目立つ色・蛍光ペンがあれば [{id, formatting}]（読み戻しで伝える）。docs：もう読んであれば、(元の文書, 書いた後の文書)。"""
    a, b = docs or (PptxDocument(original), PptxDocument(new))
    pb = _paras(b)
    out = []
    for i, ws in _written_spans(a, written).items():
        y = pb[i]
        xml = b.xml(y.part)
        pos = y.positions()
        delta = 0
        for (s0, e0), text, fid in sorted(ws):
            ns = s0 + delta
            delta += len(text) - (e0 - s0)
            found: list[str] = []
            for k in range(ns, min(ns + len(text), len(pos))):
                sg = y.segs[pos[k][0]]
                r = sg.node.parent if sg.node is not None else None
                rpr = r.find("rPr") if r is not None else None
                for n in (bad_rpr(xml[rpr.s:rpr.e]) if rpr is not None else []):
                    if n not in found:
                        found.append(n)
            if found:
                out.append({"id": fid, "formatting": found})
    return out


# ====================================================================== api から
def api_fill(blob: bytes, name: str, values: dict, force: bool = False) -> dict:
    from ..openkinds import OpenFormatError
    from .window import build

    try:
        doc = PptxDocument(blob, name)
        res = fill(doc, values, force=force)
    except (FillError, PptxError) as e:
        raise OpenFormatError(str(e)) from None
    pr = prove(blob, res.blob, written_of(res.patches))
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings),
              "unexpected": (pr.unexpected + pr.readback)[:20], "structure_errors": pr.lint_errors[:20],
              "para_read_back": pr.readback[:30],
              "ids_after_writing": "PowerPoint には書いた所の印を置けません。書いた後のファイルでは、書いた字も印刷された字に数えるので、"
                                   "書いた段落の後ろの空欄の id の字数が増え（P2.c2 に「8」と書くと、P2.c3 だった空欄が P2.c4 になる）、"
                                   "書いたマスは欄の一覧から消えます。続けて書くときは、書いた後のファイルを開き直して id を"
                                   "確かめるか、元のファイルに、まとめて書いてください"}
    from ..openkinds import _formatting, _wrapped

    both: list = []  # 元の文書と書いた後の文書（読み直すのは 1 回だけ。書き手が使った doc は、書く所を決めるのに使った後なので使わない）

    def docs():
        if not both:
            both.extend((PptxDocument(blob, name), PptxDocument(res.blob, name)))
        return both[0], both[1]

    _formatting(result, lambda: written_formatting(blob, res.blob, written_of(res.patches), docs()))

    # 書いたことで段落の行が折り返す見込み（目安）：スライドにはページの増えがないので、行の数だけを伝える
    def estimate():
        from .wrap import wrap_warnings

        a, b = docs()
        spans = _written_spans(a, written_of(res.patches))
        return wrap_warnings(a, b, {i: [fid for _s, _t, fid in ws] for i, ws in spans.items()})

    _wrapped(result, blob, "pptx", estimate, pages=False)
    anchored = {p.key: p.id for p in res.patches if p.anchor}
    if anchored:
        result["anchored"] = anchored
    told = {p.id: p.value for p in res.patches}
    # まだ空いている欄：書いた後のファイルでは id の番号が詰まるので、元のファイルの欄から、書いた所を除いたもの
    rest = [{"id": f["id"], "label": f.get("label", ""), "kind": f.get("kind", ""), "value": f.get("value"),
             **({"agent": True} if f.get("agent") else {})} for f in build(doc)["fields"] if f["id"] not in told]
    return {"blob": res.blob, "result": result, "readback": {}, "labels": {p.id: p.label for p in res.patches},
            "told": told, "rest": rest}
