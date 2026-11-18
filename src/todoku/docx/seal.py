"""Word（.docx）：本人の印影・署名の画像を、本人が同意したときだけ、字を目印にして置く。

置き方（fill の値）：
- {"@1": {"seal": 画像のパス, "consent": true, "cover": "印", "size": 12}}：その字の上に（画像の中心を、字の中心に）
- {"@1": {"seal": 画像のパス, "consent": true, "after": "代表取締役 山田太郎", "size": 12}}：その字のすぐ後ろに
- "size"（mm）：画像の長い辺の長さ。Word の文書は、欄の高さ・ページの上の位置を保存しないので、省けない
- "near"・"occurrence"：同じ字がいくつもあるときの決め方（○で囲むときと同じ）
- 署名の画像は "seal" の代わりに "signature"
ページと枠（page・box）では置けない（Word の文書には、ページの上の位置がない）。

書き方：目印の字の直前に、字に結び付けた画像（wp:anchor・pic:pic。折り返しなし：ほかの字を動かさない）を入れる。横は
relativeFrom="character"、縦は relativeFrom="line"。位置は、字の大きさと幅からの見積もり（docx/circle.py の楕円と同じ。結果に
「確かめが要る」と添える）。透ける画像（アルファのある PNG）は字の前に、透けない画像（アルファのない PNG・JPEG）は字の後ろに置く
（behindDoc）：どちらも、文書の字は隠れない。
足す部品：画像（word/media/tdkseal番号.png）、その関係（.rels）、要るときだけ、種類の表（[Content_Types].xml）の拡張子。
証明（prove）：元の文書と画像から書き換えを決め直し、書いた後の ZIP が、その書き換えと足した部品のほかは 1 バイトも変わらないことを確かめる。
読み戻し（read_back）：書いた後の文書を開き直し、その名前の画像が、目印の字の直前に、伝えた大きさで、本人の画像のバイト列のままあるか。
"""

from __future__ import annotations

import hashlib
import posixpath
import re

from ..sealimage import KINDS, SealError, SealImage, caution
from . import circle as dc
from .document import DocxDocument

EMU = 12700
MM = 72 / 25.4  # 1 mm の pt
TAG = "TdkSeal"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
EMPTY_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              f'<Relationships xmlns="{PKG_REL}"></Relationships>').encode()
UNCERTAIN = ("Word はページの上の字の位置を保存しないので、置いた画像の位置は、字の大きさと幅から見積もったものです。"
             "ずれていないか、印刷した紙か画面で確かめてください")


def contents_text(key: str) -> str:
    """画像の説明（代わりの字）：画面の読み上げにも、窓にも、何の画像かが分かるように。"""
    return f"本人の{KINDS[key]}の画像（本人の同意を得て置いたもの）"


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


# ---------------------------------------------------------------------- 置く所を決める
def resolve(doc: DocxDocument, spec: dict, key: str, img: SealImage, used: dict, original: DocxDocument | None = None) -> dict:
    """置く所：{id, part, para, start, end, text（目印の字）, mode, w・h・left・top（pt）, behind, label, …}。
    書き込みと証明で、同じ決め方を使う。original：同じ書き込みで値を書く前の文書（あれば、目印の字はその字で決める）。"""
    what = KINDS[key]
    if spec.get("box") is not None or spec.get("page") is not None:
        raise SealError("Word の文書は、ページの上の位置を保存しないので、ページと枠（page・box）では画像を置けません。"
                        '字を目印に（{"cover": "印"} はその字の上、{"after": "山田太郎"} はその字のすぐ後ろ）指定してください')
    if spec.get("where") is not None:
        raise SealError('Word の文書では、目印の字の上（"cover"）か、すぐ後ろ（"after"）にだけ置けます（"where" は使えません）')
    if spec.get("cover") is not None and spec.get("after") is not None:
        raise SealError('"cover"（その字の上）と "after"（その字のすぐ後ろ）は、どちらか1つだけを指定してください')
    mode = "cover" if spec.get("cover") is not None else "after"
    target = spec.get(mode)
    if not isinstance(target, str) or not target.strip():
        raise SealError(f'{what}の画像を置く所は、字の目印（{{"cover": "印"}} はその字の上、{{"after": "山田太郎"}} はその字のすぐ後ろ）で'
                        "指定してください")
    size = spec.get("size")
    if not (isinstance(size, (int, float)) and not isinstance(size, bool) and size > 0):
        raise SealError('"size" に、画像の長い辺の長さ（mm、正の数）を指定してください（Word の文書は、欄の高さを保存しないので、'
                        "大きさを目印の所から決められません。紙に押す印の大きさを本人に確かめてください）")
    sub = {"circle": target, **{k: spec[k] for k in ("near", "occurrence") if k in spec}}
    try:
        p, a, b, _i = dc.find_in_filled(original, doc, sub) if original is not None else dc.find(doc, sub)
    except dc.CircleError as e:
        raise SealError(str(e)) from None
    pos = p.positions()
    for k in range(a, b):
        sg = p.segs[pos[k][0]]
        if sg.kind != "t" or sg.node is None or sg.run is None:
            raise SealError(f"「{target}」が、書き換えられない部分（{sg.kind}）にかかっています")
    gl = dc._glyphs(doc, p)
    mine = [g for g in gl[a:b] if not g.text.isspace()]
    fs = max(g.size for g in mine)
    tw = max(g.x1 for g in mine) - gl[a].x0  # 目印の字の、初めの字の左の端からの幅
    s = size * MM / max(img.w, img.h)
    w, h = img.w * s, img.h * s
    # 横：初めの字の左の端から。縦：行の上から（字の箱は行の上から 0.1 字下にあると見積もる：docx/circle.py と同じ）
    left = (tw - w) / 2 if mode == "cover" else tw + 0.25 * fs
    top = 0.6 * fs - h / 2
    key_ = (p.part, id(p.node))
    for (pk, aa, kind_) in used.get("seals", []):
        if pk == key_ and aa == a and kind_ == mode:
            raise SealError(f"「{target}」の同じ所に、画像を 2 つ置く指定になっています")
    used.setdefault("seals", []).append((key_, a, mode))
    n = len(used["seals"]) + _existing(doc)
    label = spec.get("near") or (p.cell.text()[:30] if p.cell is not None else p.text.strip()[:30])
    wide = dc.east_asian_context(p.text)
    notes = []
    if any(not dc.fixed_width(ch, wide) for ch in p.text[max(0, a - 6):b + 6] if not ch.isspace()):
        notes.append("幅の変わる字（ラテン文字・キリル文字・アラビア文字など、書体で字ごとに幅が違う字）を含む（字の幅は見積もり）")
    if dc._indented(doc, p):
        notes.append("ぶら下げ・1行目の字下げ（w:ind の hanging・firstLine）のある段落で、画像の位置がずれることがあります"
                     "（Word と LibreOffice で、字に結び付けた図形の位置の基準が違う）")
    if p.cell is not None and top < 0:
        notes.append("表のマスの中の字に置いた画像で、行の上にはみ出します：LibreOffice では、はみ出す分だけ、その表の行が下に広がることがあります")
    return {"id": f"S{n}", "part": p.part, "para": p, "start": a, "end": b, "text": p.text[a:b], "style": "seal", "mode": mode,
            "w": round(w, 3), "h": round(h, 3), "left": round(left, 3), "top": round(top, 3), "behind": img.opaque,
            "size": fs, "label": label, "context": re.sub(r"\s+", " ", p.text).strip()[:60], "around": dc.around(p.text, a, b),
            "notes": notes, "what": what, "key": key, "img": img}


def _existing(doc: DocxDocument) -> int:
    n = 0
    for part in doc.parts:
        n += len(re.findall(rb'<(?:\w+:)?docPr\b[^>]*\bname="' + TAG.encode() + rb" ", doc.xml(part)))
    return n


# ---------------------------------------------------------------------- 書く
def drawing_run(pre: str, sd: dict, shape_id: int, rpr: bytes = b"") -> bytes:
    """画像の図形の、文字の要素。rpr：目印の字の書式（w:rPr）。同じ書式にして、行の高さを変えない
    （書式のない文字の要素は、文書の既定の字の大きさになり、小さい字の行を高くすることがある）。"""
    cx, cy = round(sd["w"] * EMU), round(sd["h"] * EMU)
    name = f"{TAG} {_esc(sd['id'])}"
    return (f'<{pre}:r>'.encode() + rpr + (f'<{pre}:drawing>'
            f'<wp:anchor xmlns:wp="{dc.NS["wp"]}" distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            f'relativeHeight="{251659264 + shape_id}" behindDoc="{1 if sd["behind"] else 0}" locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="character"><wp:posOffset>{round(sd["left"] * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="line"><wp:posOffset>{round(sd["top"] * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{shape_id}" name="{name}" descr="{_esc(contents_text(sd["key"]))}"/>'
            f'<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="{dc.NS["a"]}" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            f'<a:graphic xmlns:a="{dc.NS["a"]}"><a:graphicData uri="{PIC}"><pic:pic xmlns:pic="{PIC}">'
            f'<pic:nvPicPr><pic:cNvPr id="{shape_id}" name="{name}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip xmlns:r="{REL}" r:embed="{sd["rid"]}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
            f"</wp:anchor></{pre}:drawing></{pre}:r>").encode("utf-8"))


def rels_path(part: str) -> str:
    d, base = posixpath.split(part)
    return posixpath.join(d, "_rels", base + ".rels")


def _attrs(tag: bytes) -> dict[str, str]:
    return {k.decode(): v.decode("utf-8", "replace") for k, v in re.findall(rb'([\w:.-]+)\s*=\s*"([^"]*)"', tag)}


def _relationships(xml: bytes) -> list[dict[str, str]]:
    return [_attrs(m) for m in re.findall(rb"<(?:\w+:)?Relationship\b[^>]*>", xml)]


def _insert_before_close(xml: bytes, root: str, make) -> bytes:
    """根の要素（root）の閉じる所の直前に、要素を入れる。make(接頭辞) が、入れるバイト列。"""
    m = None
    for m in re.finditer(rb"</((?:\w+:)?)" + root.encode() + rb"\s*>", xml):
        pass
    if m is None:
        raise SealError(f"部品の中に、要素を足す場所（</{root}>）がありません")
    return xml[:m.start()] + make(m[1].decode()) + xml[m.start():]


def _add_rel(xml: bytes, target: str) -> tuple[bytes, str]:
    ids = {r.get("Id") for r in _relationships(xml)}
    n = 1
    while f"rIdTdkSeal{n}" in ids:
        n += 1
    rid = f"rIdTdkSeal{n}"
    return _insert_before_close(xml, "Relationships", lambda px: (
        f'<{px}Relationship Id="{rid}" Type="{REL}/image" Target="{_esc(target)}"/>').encode()), rid


def _add_type(xml: bytes, part: str, img: SealImage) -> bytes:
    """種類の表に、画像の拡張子の種類がなければ足す。同じ拡張子がほかの種類になっていれば、その部品だけの種類（Override）を足す。"""
    ext = posixpath.splitext(part)[1][1:].lower()
    same = [a for a in (_attrs(t) for t in re.findall(rb"<(?:\w+:)?Default\b[^>]*>", xml)) if a.get("Extension", "").lower() == ext]
    if any(a.get("ContentType", "").lower() == img.mime for a in same):
        return xml
    if same:
        return _insert_before_close(xml, "Types", lambda px: f'<{px}Override PartName="/{_esc(part)}" ContentType="{img.mime}"/>'.encode())
    return _insert_before_close(xml, "Types", lambda px: f'<{px}Default Extension="{ext}" ContentType="{img.mime}"/>'.encode())


def plan(doc: DocxDocument, seals: list[dict]) -> tuple[dict, dict[str, bytes], list[tuple[str, bytes, str]]]:
    """(本文の部品の書き換え {部品: [(始め, 終わり, 入れるバイト)]}, 中身を替える部品 {名前: バイト}, 足す部品 [(名前, バイト, ヘッダーを写す部品)])。"""
    z = doc.zip
    names = {n.lower() for n in z.names()}
    media: dict[tuple[str, str], str] = {}
    files: dict[str, bytes] = {}
    added: list[tuple[str, bytes, str]] = []
    new_rels: dict[str, bytes] = {}
    rid_of: dict[tuple[str, str], str] = {}
    types = z.read("[Content_Types].xml")
    for sd in seals:
        img: SealImage = sd["img"]
        d = posixpath.dirname(sd["part"])
        mname = media.get((d, img.sha256))
        if mname is None:
            n = 1
            while posixpath.join(d, "media", f"tdkseal{n}.{img.ext}").lower() in names:
                n += 1
            mname = posixpath.join(d, "media", f"tdkseal{n}.{img.ext}")
            names.add(mname.lower())
            media[(d, img.sha256)] = mname
            added.append((mname, img.data, sd["part"]))
            types = _add_type(types, mname, img)
        rp = rels_path(sd["part"])
        if (rp, mname) not in rid_of:
            cur = files.get(rp, new_rels.get(rp))
            if cur is None:
                cur = z.read(rp) if z.has(rp) else EMPTY_RELS
            cur, rid = _add_rel(cur, posixpath.relpath(mname, d))
            (files if z.has(rp) else new_rels)[rp] = cur
            rid_of[(rp, mname)] = rid
        sd["rid"], sd["media"] = rid_of[(rp, mname)], mname
        sd["draw"] = lambda pre, nid, rpr, sd=sd: drawing_run(pre, sd, nid, rpr)
    if types != z.read("[Content_Types].xml"):
        files["[Content_Types].xml"] = types
    added += [(rp, xml, rp[:rp.index("_rels/")] + posixpath.basename(rp)[:-5]) for rp, xml in new_rels.items()]
    return dc.plan(doc, seals), files, added


def _edited(doc: DocxDocument, part: str, ops_: list) -> bytes:
    data = doc.xml(part)
    for s, e, ins in sorted(ops_, key=lambda x: -x[0]):
        data = data[:s] + ins + data[e:]
    return data


def apply(doc: DocxDocument, seals: list[dict]) -> bytes:
    from ..zipraw import RawZip

    edits, files, added = plan(doc, seals)
    z = RawZip(doc.blob)
    for part, ops_ in edits.items():
        z.replace(part, _edited(doc, part, ops_))
    for n, data in files.items():
        z.replace(n, data)
    for n, data, like in added:
        z.add(n, data, like)
    return z.build()


# ---------------------------------------------------------------------- 証明と読み戻し
def prove(before: bytes, after: bytes, name: str, specs: list[tuple[str, dict, str, SealImage]], original: bytes | None = None) -> list[str]:
    """元の文書と画像から、置く所と書き換えを決め直し、書いた後の ZIP が、その書き換えと足した部品のほかは変わらないことを確かめる。
    specs：[(置いた所の id, 指定, seal か signature, 画像)]。original：値を書く前の文書（目印の字をその字で決めたとき）。"""
    from ..zipraw import RawZip

    doc = DocxDocument(before, name)
    orig = DocxDocument(original, name) if original is not None and original != before else None
    used: dict = {}
    seals, out = [], []
    for fid, spec, key, img in specs:
        try:
            sd = resolve(doc, spec, key, img, used, original=orig)
        except SealError as e:
            return [f"画像 {fid} の所が、元の文書で決まりません：{e}"]
        if sd["id"] != fid:
            out.append(f"画像 {fid} が、決め直した所（{sd['id']}）と違います")
        seals.append(sd)
    edits, files, added = plan(doc, seals)
    za, zb = RawZip(before), RawZip(after)
    new = {n: data for n, data, _like in added}
    if zb.names() != za.names() + [n for n, _d, _l in added]:
        out.append("ZIP のエントリの並びが、画像とその関係を足したほかにも変わっています")
    for n in zb.names():
        if n in edits:
            if zb.read(n) != _edited(doc, n, edits[n]):
                out.append(f"{n}: 画像を置くための書き換えのほかにも変わっています")
        elif n in files:
            if zb.read(n) != files[n]:
                out.append(f"{n}: 画像の関係・種類を足すほかにも変わっています")
        elif n in new:
            if zb.read(n) != new[n]:
                out.append(f"{n}: 足した部品が、本人の画像（とその関係）と一致しません")
        elif za.has(n):
            ea, eb = za.get(n), zb.get(n)
            if (ea.local, ea.data, ea.descriptor) != (eb.local, eb.data, eb.descriptor):
                out.append(f"{n}: 画像を置くのに変わってよい部品ではありません")
        else:
            out.append(f"{n}: 足してよい部品ではありません")
    return out


def _para_of(node):
    """その要素を直に含む段落（テキストボックスの中の段落は、外の段落と取り違えない）。"""
    while node is not None and node.local != "p":
        node = node.parent
    return node


def read_back(blob: bytes, name: str, placed: list[dict], images: dict[str, SealImage]) -> list[dict]:
    """書いた後の文書を開き直し、その名前の画像が、目印の字の直前に、伝えた大きさで、本人の画像のバイト列のままあるかを確かめる。
    images：置いた所の id → 画像（いま読み直したファイル）。"""
    doc = DocxDocument(blob, name)
    out = []
    for pl in placed:
        probs: list[str] = []
        hit = None
        for p in dc._paras(doc):
            for d in p.node.iter("docPr"):
                if d.get("name") == f"{TAG} {pl['id']}" and _para_of(d) is p.node:
                    hit = (p, d)
                    break
            if hit:
                break
        if hit is None:
            out.append({"id": pl["id"], "written": pl["what"], "read_back": "書いた後の文書に、この画像が見つかりません"})
            continue
        p, d = hit
        run = d
        while run is not None and run.local != "r":
            run = run.parent
        xml = doc.xml(p.part)
        raw = xml[run.s:run.e]
        pos = p.positions()
        # 画像の後ろの字（同じ字を目印にした、ほかの画像の印は飛ばす）
        after = [k for k in range(len(p.text)) if p.segs[pos[k][0]].kind == "t" and p.segs[pos[k][0]].run is not None
                 and p.segs[pos[k][0]].run.s > run.s]
        got = "".join(p.text[k] for k in after[:len(pl["text"])])
        if got != pl["text"]:
            probs.append(f"画像の後ろの字が「{got}」（目印の字「{pl['text']}」ではない）")
        ext = re.search(rb'<wp:extent cx="(\d+)" cy="(\d+)"', raw)
        want = [round(v * MM * EMU) for v in pl["size_mm"]]
        if ext is None or any(abs(int(g) - w_) > EMU for g, w_ in zip(ext.groups(), want)):
            probs.append("画像の大きさが、伝えた大きさと違います")
        m = re.search(rb'\br:embed="([^"]*)"', raw)
        rp = rels_path(p.part)
        rel = next((r for r in (_relationships(doc.zip.read(rp)) if doc.zip.has(rp) else [])
                    if m is not None and r.get("Id") == m[1].decode()), None)
        img = images.get(pl["id"])
        if rel is None:
            probs.append("画像への関係が見つかりません")
        else:
            part = posixpath.normpath(posixpath.join(posixpath.dirname(p.part), rel.get("Target", "")))
            data = doc.zip.read(part) if doc.zip.has(part) else None
            if data is None:
                probs.append("画像の部品が見つかりません")
            elif img is None or data != img.data or hashlib.sha256(data).hexdigest() != pl["image_sha256"]:
                probs.append("画像のバイト列が、本人の画像と一致しません")
        if probs:
            out.append({"id": pl["id"], "written": pl["what"], "read_back": "／".join(probs)[:300]})
    return out


def told(sd: dict) -> str:
    """本人に伝える文：何を、どこに置いたか。"""
    where = "上" if sd["mode"] == "cover" else "すぐ後ろ"
    lab = f"「{sd['label']}」の所の" if sd.get("label") and sd["label"] != sd["text"] else ""
    back = "。透けない画像なので、字の後ろに置きました（文書の字は隠れません）" if sd["behind"] else ""
    return (f"{lab}「{sd['text']}」の{where}に、本人の{sd['what']}の画像を置きました"
            f"（横 {sd['w'] / MM:.0f} mm・縦 {sd['h'] / MM:.0f} mm）{back}")


__all__ = ["resolve", "apply", "prove", "read_back", "told", "caution", "UNCERTAIN", "SealError"]
