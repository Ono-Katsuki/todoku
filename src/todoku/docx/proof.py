"""Word の書き込みの証明：変わったのは、書いた欄のある段落（と、その欄の定義）だけであることを、
元と書き込み後の両方を読み直して確かめる。

- ZIP の中の部品：書き換えた部品以外は、圧縮されたバイト列まで一致
- 書き換えた部品：書いた段落（と、書いたコンテンツコントロールの定義）を除いたバイト列が一致
- 書いた段落：文字が「元の文字の空欄を、書いた値に置き換えたもの」と一致し、文字の書式は元の段落にあったものだけ
- 書き換えた部品は、XML として正しく読める。しおりの番号が重ならない
- 埋め込んだ Excel の表に書いたとき（E1!Sheet1!C4）：ブックの部品は Excel の証明で、表の絵の部品は
  「変わったセルの元の字を消し、新しい字を足しただけ」であることで確かめる（todoku.embed.verify_docx）
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from ..zipraw import RawZip
from .document import DocxDocument, Para, iter_paras
from .window import Spots


@dataclass
class DocxProof:
    original_sha256: str
    new_sha256: str
    entries_total: int = 0
    entries_identical: int = 0
    paragraphs_total: int = 0
    paragraphs_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)
    embed_note: str = ""  # 埋め込んだ表に書いたときの、確かめたことの要約

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        return (f"ZIP の部品 {self.entries_identical}/{self.entries_total} がバイト一致。段落 {self.paragraphs_total:,} のうち、"
                f"変わったのは書いた欄のある {self.paragraphs_changed} 段落だけ"
                + ("" if not self.unexpected else f"（想定外の変更 {len(self.unexpected)} 件）")
                + (f"。{self.embed_note}" if self.embed_note else ""))


def _paras(doc: DocxDocument) -> list[Para]:
    out = []
    for _kind, _part, blocks in doc.all_blocks():
        out += list(iter_paras(blocks))
    return out


def prove(original: bytes, new: bytes, written: dict[str, object]) -> DocxProof:
    from ..embed import is_embed_key, verify_docx

    pr = DocxProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    za, zb = RawZip(original), RawZip(new)
    if za.names() != zb.names():
        pr.unexpected.append("ZIP の中の部品の並びが変わっています")
        return pr
    changed_parts = []
    for name in za.names():
        pr.entries_total += 1
        if za.read(name) == zb.read(name):
            pr.entries_identical += 1
        else:
            changed_parts.append(name)
    try:
        a = DocxDocument(original)
        b = DocxDocument(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません: {e}")
        return pr
    embed_written = {k: v for k, v in written.items() if is_embed_key(k)}
    written = {k: v for k, v in written.items() if k not in embed_written}
    if embed_written:
        # 埋め込んだ表：ブックの部品は Excel の証明で、表の絵の部品は「変わったセルの字だけ」で確かめる
        try:
            allowed, probs, pr.embed_note = verify_docx(a, b, embed_written)
        except Exception as e:  # noqa: BLE001
            allowed, probs = set(), [f"埋め込みの表への書き込みを確かめられません：{e}"]
        pr.unexpected += probs
        changed_parts = [p for p in changed_parts if p not in allowed]
    pa, pb = _paras(a), _paras(b)
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.unexpected.append(f"段落の数が変わっています（{len(pa)} → {len(pb)}）")
        return pr
    sp = Spots(a)
    targets: dict[int, list] = {}
    index = {id(p): i for i, p in enumerate(pa)}
    sdt_spans: dict[str, list[tuple[int, int]]] = {}
    used: dict = {}
    for fid, value in written.items():
        if isinstance(value, dict) and "@" in value:
            # 字を目印にした書き込み：元の文書で、書き込みと同じ決め方で位置を決め直す
            from .window import WriteError, resolve_write

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
        if it is None:
            pr.unexpected.append(f"書いた欄 {fid} が、元の文書にありません")
            continue
        c = it.get("control")
        if c is not None:
            if c.source == "sdt":
                sdt_spans.setdefault(c.part, []).append((c.node.s, c.node.e))
            p = c.para
            if p is not None:
                targets.setdefault(index.get(id(p), -1), []).append((fid, it, value))
            continue
        p = it["para"]
        targets.setdefault(index[id(p)], []).append((fid, it, value))
    # 部品ごとに、書いた段落を除いたバイト列を比べる
    for part in changed_parts:
        if part not in a.parts:
            pr.unexpected.append(f"本文でない部品 {part} が変わっています")
            continue
        xa, xb = a.xml(part), b.xml(part)
        spans_a = [(p.node.s, p.node.e) for i, p in enumerate(pa) if p.part == part and i in targets]
        spans_b = [(p.node.s, p.node.e) for i, p in enumerate(pb) if p.part == part and i in targets]
        spans_a += sdt_spans.get(part, [])
        if sdt_spans.get(part):
            spans_b += _matching_sdt(b, part, len(sdt_spans[part]), sdt_spans[part], a)
        if _cut(xa, spans_a) != _cut(xb, spans_b):
            pr.unexpected.append(f"{part}：書いた欄のある段落の外が変わっています")
        if len(set(re.findall(rb'bookmarkStart\b[^>]*?\bw?:?id="(\d+)"', xb))) != len(re.findall(rb'bookmarkStart\b[^>]*?\bw?:?id="(\d+)"', xb)):
            pr.lint_errors.append(f"{part}：しおりの番号が重なっています")
    # 書いた段落：文字と書式
    for i, items in targets.items():
        if i < 0:
            continue
        p_old, p_new = pa[i], pb[i]
        pr.paragraphs_changed += 1
        want = _expected_text(p_old, items)
        if want is not None and p_new.text != want:
            pr.unexpected.append(f"段落 {i + 1}：書いた後の文字が「{p_new.text[:60]}」で、思っていた「{want[:60]}」と違います")
        old_fmt = {_fmt(p_old, s, a) for s in p_old.segs if s.run is not None}
        old_fmt.add(_mark_fmt(p_old, a))
        old_fmt.add(b"")
        for s in p_new.segs:
            if s.run is not None and _fmt(p_new, s, b) not in old_fmt:
                pr.unexpected.append(f"段落 {i + 1}：元の段落にない文字の書式が使われています")
                break
    for i, (p_old, p_new) in enumerate(zip(pa, pb)):
        if i not in targets and p_old.text != p_new.text:
            pr.unexpected.append(f"書いていない段落 {i + 1} の文字が変わっています")
    return pr


def _matching_sdt(b: DocxDocument, part: str, n: int, spans_a, a: DocxDocument) -> list[tuple[int, int]]:
    """書いたコンテンツコントロールの、書き込み後の文書での位置（元の文書と同じ順番のもの）。"""
    ia = [k for k, c in enumerate(a.controls) if c.part == part and c.source == "sdt" and (c.node.s, c.node.e) in spans_a]
    sb = [c for c in b.controls if c.part == part and c.source == "sdt"]
    return [(sb[k].node.s, sb[k].node.e) for k in ia if k < len(sb)]


def _cut(xml: bytes, spans: list[tuple[int, int]]) -> bytes:
    out = bytearray()
    at = 0
    for s, e in sorted(_merge(spans)):
        out += xml[at:s]
        out += b"\x00"
        at = e
    out += xml[at:]
    return bytes(out)


def _merge(spans):
    out = []
    for s, e in sorted(spans):
        if out and s <= out[-1][1]:
            out[-1] = (out[-1][0], max(out[-1][1], e))
        else:
            out.append((s, e))
    return out


def _fmt(p: Para, s, doc: DocxDocument) -> bytes:
    """文字の書式（rPr）のバイト列。編集の記録の番号（rsid）は除く。"""
    rpr = s.run.find("rPr") if s.run is not None else None
    if rpr is None:
        return b""
    return re.sub(rb'\s+w:rsid\w*="[^"]*"', b"", doc.xml(p.part)[rpr.s:rpr.e])


def _mark_fmt(p: Para, doc: DocxDocument) -> bytes:
    ppr = p.node.find("pPr")
    mark = ppr.find("rPr") if ppr is not None else None
    if mark is None:
        return b""
    xml = doc.xml(p.part)
    inner = b"".join(xml[x.s:x.e] for x in mark.kids if x.local not in ("ins", "del", "moveFrom", "moveTo"))
    return re.sub(rb'\s+w:rsid\w*="[^"]*"', b"", xml[mark.s:mark.se] + inner + b"</w:rPr>") if inner else b""


def _expected_text(p: Para, items) -> str | None:
    text = p.text
    edits = []
    for fid, it, value in items:
        if it.get("control") is not None:
            return None  # 欄の結果の文字（書式を通した値など）は、読み戻しで確かめる
        kind = it["kind"]
        if kind == "cell":
            if text.strip(" 　  "):
                return None
            edits.append((0, len(text), str(value)))
        elif kind == "checkbox":
            a, b = it["span"]
            on = value is True or str(value).lower() in ("true", "1", "yes", "はい", "あり", "☑", "■")
            orig = text[a:b]
            edits.append((a, b, ("☑" if orig == "☐" else "■") if on else orig))
        else:
            a, b = it["span"]
            v = str(value)
            # 書き込みと同じ規則：空欄の字がすべて下線つきのときだけ、幅を保つ（字を目印にした書き込みは、指定のとおり）
            pos = p.positions()
            blank_segs = [p.segs[pos[i][0]] for i in range(a, b) if i < len(pos)]
            underlined = bool(blank_segs) and all(sg.underline for sg in blank_segs) and not it.get("anchored")
            if underlined and v:
                from .fill import _width

                room = _width(text[a:b]) - _width(v)
                v = v + (" " * room if room > 0 else "")
            edits.append((a, b, v))
    out = text
    for a, b, v in sorted(edits, reverse=True):
        out = out[:a] + v + out[b:]
    return out
