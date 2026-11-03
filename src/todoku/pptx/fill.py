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
from ..docx.fill import FillError, _as_bool, _is_text_value
from ..docx.proof import _expected_text
from ..docx.window import Spots, WriteError, box_mark, resolve_write
from ..docx.xmlpos import escape, parse
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


def written_of(patches) -> dict:
    return {p.id: ({"@": p.anchor, "value": p.value} if p.anchor else p.value) for p in patches}


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


def _prefix(xml: bytes, node) -> bytes:
    m = re.match(rb"<([\w.-]+:)?p\b", xml[node.s:node.se])
    return m[1] or b"" if m else b""


class PptxWriter:
    def __init__(self, doc: PptxDocument):
        self.doc = doc
        self.edits: dict[str, list[tuple[int, int, bytes]]] = {}

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
            self.put(p.part, at, at, self.enc(value))
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
            self.put(p.part, cm[k0][0], cm[k1 - 1][1], self.enc(value) if first else b"")
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
                raise FillError(f"{key}: そのような欄はありません")
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
    return FillResult(w.build(), patches)


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
    """段落の XML から、字の要素（a:r）を除いたもの（段落の書式・行の区切り・終わりの印）。"""
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
        if not _formats(xb, y) <= fmts:
            pr.unexpected.append(f"段落 {i + 1}：元の段落にない字の書式が使われています")
        want = _expected_text(x, items)
        if want is None or y.text != want:
            pr.readback.append(f"段落 {i + 1}：読み直すと「{y.text[:40]}」で、思っていた「{(want or '')[:40]}」と違います")
    for i, (x, y) in enumerate(zip(pa, pb)):
        if i not in targets and x.text != y.text:
            pr.unexpected.append(f"書いていない段落 {i + 1} の字が変わっています（「{x.text[:30]}」→「{y.text[:30]}」）")
    return pr


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
              "ids_after_writing": "PowerPoint には書いた所の印を置けません。書いた後のファイルでは、書いた段落の残りの空欄の id の番号が"
                                   "前に詰まり、書いたマスは欄の一覧から消えます。続けて書くときは、書いた後のファイルを開き直して id を"
                                   "確かめるか、元のファイルに、まとめて書いてください"}
    anchored = {p.key: p.id for p in res.patches if p.anchor}
    if anchored:
        result["anchored"] = anchored
    told = {p.id: p.value for p in res.patches}
    # まだ空いている欄：書いた後のファイルでは id の番号が詰まるので、元のファイルの欄から、書いた所を除いたもの
    rest = [{"id": f["id"], "label": f.get("label", ""), "kind": f.get("kind", ""), "value": f.get("value"),
             **({"agent": True} if f.get("agent") else {})} for f in build(doc)["fields"] if f["id"] not in told]
    return {"blob": res.blob, "result": result, "readback": {}, "labels": {p.id: p.label for p in res.patches},
            "told": told, "rest": rest}
