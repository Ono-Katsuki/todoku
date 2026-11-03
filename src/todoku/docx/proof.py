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
from ..inputcheck import GLYPH_OFF, GLYPH_ON, checkbox_state
from .window import Spots, box_mark


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
    added = zb.names()[len(za.names()):]  # ActiveX の部品の、共有していた控えの絵から分けた絵（下で確かめる）
    if zb.names()[:len(za.names())] != za.names() or any("/media/" not in n for n in added):
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
    # 結び付けた XML の部品（w:dataBinding の customXml/itemN.xml）：書いた値の所だけが変わり、その値が書いた値として表示されること
    bound = _bound_parts(a, b, za, zb, sp, written, pr)
    changed_parts = [p for p in changed_parts if p not in bound]
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
    # ActiveX の部品（MS-OFORMS）：書いた部品の、書いた性質のバイトと大きさの記録だけが変わり、控えの絵は書いた状態の絵
    ax_parts: set[str] = set()
    if getattr(a, "ax_refs", None):
        from ..axfile import _ax_ref_parts, verify_word_refs

        ids = {id(it["control"].ax): k for k, it in sp.items.items()
               if it.get("control") is not None and it["control"].source == "activex"}
        plain = {k: v for k, v in written.items() if not (isinstance(v, dict) and "@" in v)}
        pr.unexpected += verify_word_refs(a.ax_refs, getattr(b, "ax_refs", []), za, zb, plain, ids)
        for r in list(a.ax_refs) + list(getattr(b, "ax_refs", [])):
            ax_parts |= _ax_ref_parts(r)
        for n in added:
            if n not in ax_parts:
                pr.unexpected.append(f"ActiveX の部品の絵でない部品 {n} が足されました")
        # 絵を分けたときは、関係（.rels）も変わる：分けた絵への関係を足しただけであること（本文の r:id は書いた段落の中）
        from ..axfile import _rels_path, rels_added_only

        for owner in {o for r in a.ax_refs for _k, o, _key, _p in r.pic_refs}:
            rp = _rels_path(owner)
            if rp in changed_parts:
                pr.unexpected += rels_added_only(za.read(rp), zb.read(rp), added)
                ax_parts.add(rp)
    # 部品ごとに、書いた段落を除いたバイト列を比べる
    if "word/numbering.xml" in changed_parts and any(it.get("bullet") is not None for v in targets.values() for _f, it, _x in v):
        if _numbering(a, b, pa, pb, targets, pr):
            pass
        changed_parts = [x for x in changed_parts if x != "word/numbering.xml"]
    for part in changed_parts:
        if part not in a.parts:
            if part not in ax_parts:
                pr.unexpected.append(f"本文でない部品 {part} が変わっています")
            continue
        xa, xb = a.xml(part), b.xml(part)
        spans_a = [(p.node.s, p.node.e) for i, p in enumerate(pa) if p.part == part and i in targets]
        spans_b = [(p.node.s, p.node.e) for i, p in enumerate(pb) if p.part == part and i in targets]
        spans_a += sdt_spans.get(part, [])
        if sdt_spans.get(part):
            spans_b += _matching_sdt(b, part, len(sdt_spans[part]), sdt_spans[part], a)
        # テキストボックスの代わりの中身（mc:Fallback）：書いた段落と同じ順の段落に同じ字を写したか、取り除いたこと
        fa_, fb_ = _fallbacks(a, b, part, spans_a, pr)
        spans_a += fa_
        spans_b += fb_
        if _cut(xa, spans_a) != _cut(xb, spans_b):
            pr.unexpected.append(f"{part}：書いた欄のある段落の外が変わっています")
        pr.lint_errors += _mark_ids(b, part)
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
        # 書いた字には、取り消し線・字の色などを引き継がない（docx.fill.clean_rpr）：その書式を除いた形も、元の段落の書式として認める
        from .fill import clean_rpr

        old_fmt |= {clean_rpr(f)[0] for f in old_fmt}
        # 書いたコンテンツコントロールの字の、プレースホルダーの文字のスタイル（w:rStyle）を外した書式：Word が値を入れるときと
        # 同じ変更なので、書き込みの一部として認める（外してよいのは、そのスタイルの要素だけ）
        from .fill import placeholder_style

        for _fid, it, _v in items:
            c = it.get("control")
            if c is None or c.source != "sdt":
                continue
            for s in c.result_segs:
                st = placeholder_style(a, s.run) if s.run is not None else None
                if st is not None:
                    rpr = s.run.find("rPr")
                    xml = a.xml(p_old.part)
                    old_fmt.add(re.sub(rb'\s+w:rsid\w*="[^"]*"', b"", xml[rpr.s:st.s] + xml[st.e:rpr.e]))
        for s in p_new.segs:
            if s.run is not None and _fmt(p_new, s, b) not in old_fmt:
                pr.unexpected.append(f"段落 {i + 1}：元の段落にない文字の書式が使われています")
                break
    for i, (p_old, p_new) in enumerate(zip(pa, pb)):
        if i not in targets and p_old.text != p_new.text:
            pr.unexpected.append(f"書いていない段落 {i + 1} の文字が変わっています")
    # 段落番号・箇条書きの印：書いた箇条書きの □ の段落だけが、チェックした印に変わる
    bullets = {i for i, items in targets.items() for _f, it, v in items if it.get("bullet") is not None and v is True}
    for i, (p_old, p_new) in enumerate(zip(pa, pb)):
        if i in bullets:
            if p_new.label not in ("■", "☑"):
                pr.unexpected.append(f"段落 {i + 1}：箇条書きの印が「{p_old.label}」から「{p_new.label}」になっています（チェックした印でない）")
        elif p_old.label != p_new.label:
            pr.unexpected.append(f"書いていない段落 {i + 1} の段落番号・箇条書きの印が変わっています（「{p_old.label}」→「{p_new.label}」）")
    return pr


def _bound_parts(a: DocxDocument, b: DocxDocument, za: RawZip, zb: RawZip, sp: Spots, written: dict, pr: DocxProof) -> set[str]:
    """書いたコンテンツコントロールの、結び付けた XML の部品を確かめる：元の部品の値の所（元の文書で XPath をたどり直した所）を、
    書いた後の部品で読んだ値に置き換えたものが、書いた後の部品とバイト一致し、その値を Word が表示する字が書いた値であること。
    確かめた部品の名前を返す（ほかの部品の比べ方から外す）。"""
    from .databind import BindError, binding_of, edit, locate, shown_value

    todo: dict[str, dict] = {}
    for fid, value in written.items():
        if isinstance(value, dict):
            continue
        it = sp.items.get(fid)
        c = it.get("control") if it else None
        if c is None or c.source != "sdt" or binding_of(c) is None:
            continue
        try:
            ta = locate(a, c)
        except BindError as e:
            pr.unexpected.append(f"{fid}: 結び付けた XML の値の所が、元の文書でたどれません（{e}）")
            continue
        todo.setdefault(ta.part, {}).setdefault((ta.s, ta.e), (ta, []))[1].append((fid, value))
    if not todo:
        return set()
    spb = Spots(b)
    allowed = set()
    for part, spots in todo.items():
        exp = za.read(part)
        for (_s, _e), (ta, ws) in sorted(spots.items(), key=lambda x: -x[0][0]):
            got = None
            for fid, value in ws:
                cb = (spb.items.get(fid) or {}).get("control")
                try:
                    tb = locate(b, cb) if cb is not None else None
                    if tb is None or tb.part != part:
                        raise BindError("書いた後の文書に、同じ結び付けがありません")
                    shown = shown_value(cb, tb.value, binding_of(cb))
                except Exception as e:  # noqa: BLE001
                    pr.unexpected.append(f"{fid}: 書いた後の文書で、結び付けた XML の値を読めません（{e}）")
                    exp = None
                    break
                same = (bool(shown) == bool(value)) if isinstance(value, bool) else str(shown) == str(value)
                if not same:
                    pr.unexpected.append(f"{fid}: 結び付けた XML の値「{tb.value}」は、Word で「{shown}」と表示され、書いた「{value}」と違います")
                got = tb.value
            if exp is None:
                break
            exp = edit(exp, ta, got)
        if exp is not None and exp != zb.read(part):
            pr.unexpected.append(f"{part}：結び付けた XML の部品の、書いた値の所のほかが変わっています")
        allowed.add(part)
    return allowed


def _numbering(a: DocxDocument, b: DocxDocument, pa: list, pb: list, targets: dict, pr: DocxProof) -> bool:
    """numbering.xml の差が、書いた箇条書きの □ のための w:num（todoku.docx.bullet）を足しただけか。
    足した w:num は、それぞれ書いた段落の1つだけが使い、元の段落の w:num と同じ抽象の定義を指し、段の定義は元の定義の
    バイト列と、印の字（lvlText）のほかは同じ。"""
    from .bullet import PART, BulletError, checked_lvl, level_of, _root

    xa, xb = a.xml(PART), b.xml(PART)
    try:
        ta, tb = _root(xa), _root(xb)
    except BulletError as e:
        pr.unexpected.append(f"{PART}：{e}")
        return False
    old_ids = {n.get("numId") for n in ta.findall("num")}
    added = [n for n in tb.findall("num") if n.get("numId") not in old_ids]
    if not added or _cut(xb, [(n.s, n.e) for n in added]).replace(b"\x00", b"") != xa:
        pr.unexpected.append(f"{PART}：足した段落番号の定義（w:num）のほかが変わっています")
        return False
    users: dict[str, list[int]] = {}
    for i, p in enumerate(pb):
        npr = p.node.find("pPr").find("numPr") if p.node is not None and p.node.find("pPr") is not None else None
        ni = npr.find("numId") if npr is not None else None
        if ni is not None and ni.get("val") not in old_ids:
            users.setdefault(ni.get("val"), []).append(i)
    ok = True
    for n in added:
        nid = n.get("numId")
        who = users.get(nid, [])
        bl = [it.get("bullet") for k in who for _f, it, _v in targets.get(k, []) if it.get("bullet") is not None]
        if len(who) != 1 or len(bl) != 1:
            pr.unexpected.append(f"{PART}：足した w:num {nid} を使うのが、書いた箇条書きの段落1つではありません")
            ok = False
            continue
        try:
            aid, lvl, num = level_of(xa, bl[0]["num"], bl[0]["ilvl"])
            want_lvl = checked_lvl(xa, lvl)
        except BulletError as e:
            pr.unexpected.append(f"{PART}：{e}")
            ok = False
            continue
        an = n.find("abstractNumId")
        los = [lo for lo in n.findall("lvlOverride") if int(lo.get("ilvl", 0) or 0) == bl[0]["ilvl"]]
        lvl_b = los[0].find("lvl") if len(los) == 1 else None
        if an is None or an.get("val") != aid or lvl_b is None or xb[lvl_b.s:lvl_b.e] != want_lvl:
            pr.unexpected.append(f"{PART}：足した w:num {nid} の定義が、元の段の定義の印の字だけを変えたものではありません")
            ok = False
    return ok


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


def _ptext(node) -> str:
    return "".join(t.text for t in node.iter("t"))


def _fallbacks(a: DocxDocument, b: DocxDocument, part: str, spans_a: list, pr: DocxProof) -> tuple[list, list]:
    """書いた段落のある mc:Choice の mc:Fallback を、書き込みと同じ決め方（docx.fill.fallback_twins）で決め、
    写した段落は、書いた後に Choice の段落と同じ字か、取り除いたものは、書いた後に Fallback がないかを確かめる。
    比べから外す範囲（元・書いた後）を返す。"""
    from .fill import _paras_under, fallback_twins

    ra, rb_ = a.parts[part][1], b.parts[part][1]
    acs_a, acs_b = list(ra.iter("AlternateContent")), list(rb_.iter("AlternateContent"))
    copy, drop = fallback_twins(ra, a.xml(part), [(s_, e_, b"") for s_, e_ in spans_a])
    if not copy and not drop:
        return [], []
    if len(acs_a) != len(acs_b):
        pr.unexpected.append(f"{part}：図形の入れ物（mc:AlternateContent）の数が変わっています")
        return [], []
    out_a, out_b = [], []
    for fp, cp in copy:
        k = next(i for i, ac in enumerate(acs_a) if ac.s <= fp.s and fp.e <= ac.e and ac.find("Fallback") is not None
                 and ac.find("Fallback").s <= fp.s and fp.e <= ac.find("Fallback").e)
        ac_b = acs_b[k]
        ch_b, fb_b = ac_b.find("Choice"), ac_b.find("Fallback")
        ca, fa = _paras_under(acs_a[k].find("Choice")), _paras_under(acs_a[k].find("Fallback"))
        cb, fbs = (_paras_under(ch_b), _paras_under(fb_b)) if ch_b is not None and fb_b is not None else ([], [])
        if len(cb) != len(ca) or len(fbs) != len(fa):
            pr.unexpected.append(f"{part}：テキストボックスの段落の数が変わっています")
            continue
        i = fa.index(fp)
        j = ca.index(cp)
        if _ptext(fbs[i]) != _ptext(cb[j]):
            pr.unexpected.append(f"{part}：テキストボックスの代わりの中身（mc:Fallback）の字「{_ptext(fbs[i])[:40]}」が、"
                                 f"書いた後の表示用の中身の字「{_ptext(cb[j])[:40]}」と違います")
        out_a.append((fp.s, fp.e))
        out_b.append((fbs[i].s, fbs[i].e))
    for fb in drop:
        k = next(i for i, ac in enumerate(acs_a) if ac.find("Fallback") is fb)
        ac_b = acs_b[k]
        if ac_b.find("Fallback") is not None or ac_b.find("Choice") is None:
            pr.unexpected.append(f"{part}：書き写せないテキストボックスの代わりの中身（mc:Fallback）が、取り除かれていません")
            continue
        out_a.append((fb.s, fb.e))
        out_b.append((ac_b.find("Choice").e, ac_b.find("Choice").e))
    return out_a, out_b


_REV = ("ins", "del", "moveTo", "moveFrom")


def _mark_ids(b: DocxDocument, part: str) -> list[str]:
    """書いた所の印（しおり _todoku_…）：変更履歴（w:ins・w:del など）の中にないこと、番号がほかの要素の w:id と重ならないこと。"""
    from .window import MARK_PREFIX

    out = []
    root = b.parts[part][1]
    ids: dict[str, int] = {}
    for n in root.iter():
        v = n.get("id")
        if v is not None and n.local != "docPr" and v.isdigit():
            ids[v] = ids.get(v, 0) + 1
    for n in root.iter("bookmarkStart"):
        if not (n.get("name") or "").startswith(MARK_PREFIX):
            continue
        x = n.parent
        while x is not None and x.local != "p":
            if x.local in _REV:
                out.append(f"{part}：書いた所の印 {n.get('name')} が、変更履歴（{x.local}）の中にあります")
                break
            x = x.parent
        # 同じ番号は、しおりの始まりと終わりの2つだけ
        if ids.get(n.get("id"), 0) > 2:
            out.append(f"{part}：書いた所の印 {n.get('name')} の番号 {n.get('id')} が、ほかの要素（変更履歴・コメントなど）の番号と重なっています")
    return out


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
        if it.get("bullet") is not None:
            continue  # 箇条書きの印：段落の字は変わらない（印は段落番号の定義。下で確かめる）
        if kind == "cell":
            if text.strip(" 　  "):
                return None
            edits.append((0, len(text), str(value)))
        elif kind == "checkbox":
            a, b = it["span"]
            on = checkbox_state(value, GLYPH_ON, GLYPH_OFF) is True  # 書き手と同じ読み方（inputcheck）
            orig = text[a:b]
            edits.append((a, b, box_mark(orig, on)))
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
