"""HWPX：本文の字（選択肢の「예」「2. 갱신」など）を囲む。

指定（fill の値）：{"@c1": {"circle": 囲む字, "near": 近くの見出しの字, "occurrence": 何番目か, "style": "border"}}。

書き方は「字の囲み線」（글자 테두리）だけ：囲む字だけを新しい字の並び（hp:run）に分け、その並びの文字の形（charPr）を、元の
文字の形の写しに囲み線（borderFill の上下左右が実線）を付けたものにする。字そのものの書式なので、位置はいつも正しい
（ずれない）。○ではなく□。HWPX は字のページの上の位置を保存しないので、楕円の図形（hp:ellipse）の位置は見積もりになり、
使わない（style "ellipse" は断る）。
- header.xml：borderFill と charPr を1つずつ（字の形ごとに）後ろに足し、itemCnt を直す。ほかのバイトは変えない
- 本文の節：囲む字の hp:t を「</hp:t></hp:run><hp:run charPrIDRef="新しい形"><hp:t>字</hp:t></hp:run><元の hp:run><hp:t>」で分ける
  （字のバイトはそのまま）。字の並びが変わった段落は、行の配置の控え（hp:linesegarray）を除く
"""

from __future__ import annotations

import re

from ..docx.circle import CircleError, find
from ..docx.document import iter_paras
from .document import HwpxDocument
from .package import L, kid, kids

HEADER = "Contents/header.xml"
UNCERTAIN = ("HWPX の字の囲み線（글자 테두리）は、한글で開いて表示を確かめていません（ここでは仕様どおりに書いたことだけを確かめました）。"
             "印刷した紙か画面で確かめてください")


def resolve(doc: HwpxDocument, spec: dict, used: dict) -> dict:
    style = spec.get("style", "border")
    if style != "border":
        raise CircleError('HWPX は字のページの上の位置を保存しないので、○の図形の位置を決められません。style は "border"'
                          "（字の囲み線。○ではなく□で、位置はずれない）にしてください")
    if not doc.pkg.has(HEADER):
        raise CircleError("文字の形の一覧（Contents/header.xml）がないので、囲み線を付けられません")
    p, a, b, pi = find(doc, spec)
    pos = p.positions()
    pieces: list[tuple] = []  # (hp:run, hp:t, 始めのバイト, 終わりのバイト)
    for k in range(a, b):
        sg = p.segs[pos[k][0]]
        src = sg.node
        if sg.kind != "t" or src is None or src.kind != "txt" or sg.field is not None or sg.hidden:
            raise CircleError(f"「{spec['circle']}」が、囲めない部分（{sg.kind}）にかかっています")
        t_el = src.el
        run = t_el.parent
        if L(run) != "run":
            raise CircleError(f"「{spec['circle']}」の字が、字の並び（hp:run）の直下にありません")
        s, e = src.starts[pos[k][1]], src.ends[pos[k][1]]
        if pieces and pieces[-1][1] is t_el and pieces[-1][3] == s:
            pieces[-1] = (run, t_el, pieces[-1][2], e)
        else:
            pieces.append((run, t_el, s, e))
    key = (p.part, id(p.el))
    for pk, aa, bb in used.get("spans", []):
        if pk == key and aa < b and a < bb:
            raise CircleError(f"「{spec['circle']}」は、ほかの囲みと重なります")
    used.setdefault("spans", []).append((key, a, b))
    label = spec.get("near") or (p.cell.text()[:30] if p.cell is not None else p.text.strip()[:30])
    return {"id": f"□{len(used['spans'])}", "part": p.part, "para": p, "pi": pi, "start": a, "end": b, "text": p.text[a:b],
            "style": "border", "pieces": pieces, "label": label, "context": re.sub(r"\s+", " ", p.text).strip()[:60]}


def _header_plan(doc: HwpxDocument, cps: list[str]) -> tuple[bytes, dict[str, str]]:
    """元の文字の形（charPr の id）ごとに、囲み線を付けた写しを足した header.xml と、元の id → 新しい id。"""
    data = doc.pkg.xml(HEADER)
    root = doc.pkg.tree(HEADER)
    top = next(k for k in root.kids if hasattr(k, "tag"))
    chars = next((x for x in top.iter() if L(x) == "charProperties"), None)
    fills = next((x for x in top.iter() if L(x) == "borderFills"), None)
    if chars is None or fills is None or chars.se == chars.e or fills.se == fills.e:
        raise CircleError("文字の形・囲み線の一覧（hh:charProperties・hh:borderFills）がないので、囲み線を付けられません")
    by_cp = {x.attrs.get("id"): x for x in kids(chars, "charPr")}
    by_bf = {x.attrs.get("id"): x for x in kids(fills, "borderFill")}
    nums = lambda d: [int(k) for k in d if k and k.isdigit()]  # noqa: E731
    next_cp, next_bf = max(nums(by_cp), default=-1) + 1, max(nums(by_bf), default=0) + 1
    new_cp, new_bf = [], []
    mapping: dict[str, str] = {}
    for cp in dict.fromkeys(cps):
        el = by_cp.get(cp)
        if el is None:
            raise CircleError(f"文字の形 {cp} が header.xml にありません")
        bf = by_bf.get(el.attrs.get("borderFillIDRef", "")) or next(iter(by_bf.values()), None)
        if bf is None:
            raise CircleError("囲み線の形（hh:borderFill）が1つもないので、写しを作れません")
        bfx = data[bf.s:bf.e]
        bfx = re.sub(rb'(<(?:\w+:)?borderFill\b[^>]*?\bid=")[^"]*"', lambda m: m[1] + str(next_bf).encode() + b'"', bfx, count=1)
        for side in (b"leftBorder", b"rightBorder", b"topBorder", b"bottomBorder"):
            bfx = re.sub(rb"(<(?:\w+:)?" + side + rb'\b[^>]*?\btype=")[^"]*"', lambda m: m[1] + b'SOLID"', bfx, count=1)
            bfx = re.sub(rb"(<(?:\w+:)?" + side + rb'\b[^>]*?\bwidth=")[^"]*"', lambda m: m[1] + b'0.12 mm"', bfx, count=1)
            bfx = re.sub(rb"(<(?:\w+:)?" + side + rb'\b[^>]*?\bcolor=")[^"]*"', lambda m: m[1] + b'#000000"', bfx, count=1)
        cpx = data[el.s:el.e]
        cpx = re.sub(rb'(<(?:\w+:)?charPr\b[^>]*?\bid=")[^"]*"', lambda m: m[1] + str(next_cp).encode() + b'"', cpx, count=1)
        if re.search(rb'<(?:\w+:)?charPr\b[^>]*?\bborderFillIDRef="', cpx):
            cpx = re.sub(rb'(<(?:\w+:)?charPr\b[^>]*?\bborderFillIDRef=")[^"]*"', lambda m: m[1] + str(next_bf).encode() + b'"', cpx, count=1)
        else:
            cpx = re.sub(rb"(<(?:\w+:)?charPr\b)", lambda m: m[1] + b' borderFillIDRef="' + str(next_bf).encode() + b'"', cpx, count=1)
        mapping[cp] = str(next_cp)
        new_cp.append(cpx)
        new_bf.append(bfx)
        next_cp += 1
        next_bf += 1
    edits = [(chars.es, chars.es, b"".join(new_cp)), (fills.es, fills.es, b"".join(new_bf))]
    for el, n in ((chars, len(new_cp)), (fills, len(new_bf))):
        tag = data[el.s:el.se]
        m = re.search(rb'\bitemCnt="(\d+)"', tag)
        if m:
            new = tag[:m.start(1)] + str(int(m[1]) + n).encode() + tag[m.end(1):]
            edits.append((el.s, el.se, new))
    out = bytearray()
    at = 0
    for s, e, b in sorted(edits, key=lambda x: (x[0], x[1])):
        out += data[at:s] + b
        at = e
    out += data[at:]
    return bytes(out), mapping


def apply(doc: HwpxDocument, circles: list[dict]) -> bytes:
    cps = [run.attrs.get("charPrIDRef", "0") for c in circles for run, _t, _s, _e in c["pieces"]]
    header, mapping = _header_plan(doc, cps)
    edits: dict[str, list[tuple[int, int, bytes]]] = {}
    relaid: set[int] = set()
    for c in circles:
        part = c["part"]
        data = doc.pkg.xml(part)
        for run, t_el, s, e in c["pieces"]:
            pre = run.prefix_of(run.tag[1:run.tag.index("}")]) or "hp"
            new_cp = mapping[run.attrs.get("charPrIDRef", "0")]
            body = (f"</{pre}:t></{pre}:run><{pre}:run charPrIDRef=\"{new_cp}\"><{pre}:t>".encode() + data[s:e]
                    + f"</{pre}:t></{pre}:run>".encode() + data[run.s:run.se] + data[t_el.s:t_el.se])
            edits.setdefault(part, []).append((s, e, body))
        p = c["para"]
        ls = kid(p.el, "linesegarray")
        if ls is not None and id(p.el) not in relaid:
            relaid.add(id(p.el))
            edits.setdefault(part, []).append((ls.s, ls.e, b""))
    changed = {HEADER: header}
    for part, eds in edits.items():
        data = doc.pkg.xml(part)
        out = bytearray()
        at = 0
        for s, e, b in sorted(eds, key=lambda x: (x[0], x[1])):
            if s < at:
                raise CircleError("同じ所に2つの囲みが重なりました")
            out += data[at:s] + b
            at = e
        out += data[at:]
        changed[part] = bytes(out)
    return doc.pkg.rebuilt(changed)


def _bordered(doc: HwpxDocument) -> set[str]:
    """囲み線（上下左右が実線）のある文字の形（charPr の id）。"""
    root = doc.pkg.tree(HEADER)
    top = next(k for k in root.kids if hasattr(k, "tag"))
    boxed = set()
    for bf in top.iter():
        if L(bf) == "borderFill":
            sides = [kid(bf, s) for s in ("leftBorder", "rightBorder", "topBorder", "bottomBorder")]
            if all(x is not None and x.attrs.get("type", "NONE") not in ("NONE", "") for x in sides):
                boxed.add(bf.attrs.get("id"))
    return {cp.attrs.get("id") for cp in top.iter() if L(cp) == "charPr" and cp.attrs.get("borderFillIDRef") in boxed}


def prove(before: bytes, after: bytes, name: str, circles: list[dict], specs: list[dict]) -> list[str]:
    """囲んだ後の文書：同じ指定を元の文書で決め直して作ったものとバイト一致、段落の字は変わらない、囲み線の字は指定の字だけ。"""
    bad = []
    a = HwpxDocument(before, name)
    used: dict = {}
    try:
        again = apply(a, [resolve(a, s, used) for s in specs])
    except CircleError as e:
        return [f"囲みの指定が、元の文書で決まりません：{e}"]
    if again != after:
        bad.append("囲んだ後の文書が、元の文書で決め直した書き方と一致しません")
    b = HwpxDocument(after, name)
    ta = [p.text for _k, _p, bl in a.all_blocks() for p in iter_paras(bl)]
    tb = [p.text for _k, _p, bl in b.all_blocks() for p in iter_paras(bl)]
    if ta != tb:
        bad.append("囲んだ後の文書の段落の字が変わっています")
    bad += [f"{x['id']}: {x['read_back']}" for x in read_back(after, name, circles)]
    return bad


def read_back(blob: bytes, name: str, circled: list[dict]) -> list[dict]:
    """書いた後の文書を開き直し、囲み線のある字が、言った字（段落の同じ位置）だけかを確かめる（段落の字と数は変えないので、
    段落の番号と字の位置で照らせる）。"""
    doc = HwpxDocument(blob, name)
    boxed = _bordered(doc)
    paras = [p for _k, _part, blocks in doc.all_blocks() for p in iter_paras(blocks)]
    mine: dict[int, set[int]] = {}
    for c in circled:
        mine.setdefault(c["pi"], set()).update(range(c["start"], c["end"]))
    out = []
    for c in circled:
        p = paras[c["pi"]] if c["pi"] < len(paras) else None
        if p is None or p.text[c["start"]:c["end"]] != c["text"]:
            out.append({"id": c["id"], "written": f"□「{c['text']}」", "read_back": "書いた後の文書で、同じ段落の同じ位置に、その字がありません"})
            continue
        pos = p.positions()
        on = {k for k in range(len(p.text)) if not p.text[k].isspace()
              and p.segs[pos[k][0]].run is not None and p.segs[pos[k][0]].run.attrs.get("charPrIDRef") in boxed}
        want = {k for k in range(c["start"], c["end"]) if not p.text[k].isspace()}
        stray = on - mine[c["pi"]]
        if not want <= on or stray:
            got = "".join(p.text[k] for k in sorted(on))
            out.append({"id": c["id"], "written": f"□「{c['text']}」", "read_back": f"囲み線の字「{got}」"})
    return out
