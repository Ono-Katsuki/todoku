"""Word（.docx）：本文の字（選択肢の「有」「2.更新」など）を○で囲む。

指定（fill の値）：{"@c1": {"circle": 囲む字, "near": 近くの見出しの字, "occurrence": 何番目か, "pad": すき間（pt）,
"style": "ellipse"（既定）か "border"}}。

二つの書き方と、その確かさ（どちらも Word と LibreOffice で同じに表示される形だけを使う）：
- "ellipse"（既定）：囲む字の直前に、字に結び付けた楕円の図形（wp:anchor・wps の ellipse・塗りなし）を入れる。横の位置は
  「その字の位置から」（positionH relativeFrom="character"）、縦は「その行の上から」（positionV relativeFrom="line"）なので、
  字が動けば○も一緒に動く。ただし Word はページの上の字の位置を保存しないので、字の幅（全角は字の大きさ、半角はその半分）
  と行の中の字の高さは見積もり。字間の調整・文字グリッド・均等割り付けがあると、ずれることがある（結果に「確かめが要る」と添える）
- "border"：囲む字だけの文字の要素に、四角い囲み線（w:bdr）を付ける。○ではなく□だが、字そのものの書式なので、位置は
  いつも正しい（ずれない）。○でなくてもよいと本人が確かめたときに使う
1 字だけを囲む「囲い文字」（EQ フィールド）は、LibreOffice での表示が Word と違うので使わない。
"""

from __future__ import annotations

import re

from ..pdf.circle import _pattern, face, inside, shape, touches
from ..textwidth import cells, east_asian_context, fixed_width
from .document import DocxDocument, Para, iter_paras

EMU = 12700
TAG = "TdkCircle"
NS = {"wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"}
UNCERTAIN = ("Word はページの上の字の位置を保存しないので、○（字に結び付けた楕円の図形）の大きさと高さは、字の大きさと幅から"
             "見積もったものです。ずれていないか、印刷した紙か画面で確かめてください")


class CircleError(ValueError):
    pass


class _G:
    __slots__ = ("text", "x0", "y0", "x1", "y1", "size", "vertical")

    def __init__(self, text, x0, y0, x1, y1, size):
        self.text, self.x0, self.y0, self.x1, self.y1, self.size, self.vertical = text, x0, y0, x1, y1, size, False


# ---------------------------------------------------------------------- 字の大きさ
def _styles(doc: DocxDocument) -> dict:
    memo = getattr(doc, "_circle_styles", None)
    if memo is not None:
        return memo
    out = {"default": 10.0, "sz": {}, "based": {}, "para_default": ""}
    root = doc._load("word/styles.xml")
    top = root.kids[0] if root is not None and root.kids else None
    if top is not None:
        dd = top.find("docDefaults")
        rd = dd.find("rPrDefault") if dd is not None else None
        rpr = rd.find("rPr") if rd is not None else None
        sz = rpr.find("sz") if rpr is not None else None
        if sz is not None and str(sz.get("val", "")).isdigit():
            out["default"] = int(sz.get("val")) / 2
        for st in top.findall("style"):
            sid = st.get("styleId")
            b = st.find("basedOn")
            if b is not None:
                out["based"][sid] = b.get("val")
            rp = st.find("rPr")
            s2 = rp.find("sz") if rp is not None else None
            if s2 is not None and str(s2.get("val", "")).isdigit():
                out["sz"][sid] = int(s2.get("val")) / 2
            if st.get("type") == "paragraph" and st.get("default") in ("1", "true"):
                out["para_default"] = sid
    doc._circle_styles = out
    return out


def _style_size(st: dict, sid: str) -> float | None:
    seen = 0
    while sid and seen < 20:
        if sid in st["sz"]:
            return st["sz"][sid]
        sid = st["based"].get(sid, "")
        seen += 1
    return None


def run_size(doc: DocxDocument, p: Para, run) -> float:
    """文字の要素の字の大きさ（pt）：文字の要素の書式 → 文字のスタイル → 段落のスタイル → 既定。"""
    st = _styles(doc)
    rpr = run.find("rPr") if run is not None else None
    if rpr is not None:
        sz = rpr.find("sz")
        if sz is not None and str(sz.get("val", "")).isdigit():
            return int(sz.get("val")) / 2
        rs = rpr.find("rStyle")
        if rs is not None:
            v = _style_size(st, rs.get("val", ""))
            if v:
                return v
    v = _style_size(st, p.style or st["para_default"])
    return v or st["default"]


def width(ch: str, size: float, ambiguous_wide: bool = False) -> float:
    """字の幅の見積もり（textwidth：Unicode の東アジアの幅）：全角（W・F）は字の大きさ、曖昧（A：○・ギリシャ文字・
    キリル文字など）は東アジアの文脈（段落に全角の字がある）なら字の大きさ・ほかは半分、結合の記号は 0、ほかは半分。"""
    return 0.5 * size * cells(ch, ambiguous_wide)


# ---------------------------------------------------------------------- 囲む所
def _paras(doc: DocxDocument):
    for _kind, part, blocks in doc.all_blocks():
        for p in iter_paras(blocks):
            yield p


def _cell_pos(p: Para):
    c = p.cell
    if c is None:
        return None
    return (id(c.table), c.row, c.col)


def find(doc: DocxDocument, spec: dict):
    """(段落, 始め, 終わり, 段落の番号)。"""
    target = spec.get("circle")
    if not isinstance(target, str) or not target.strip():
        raise CircleError('○で囲むときは {"circle": 囲む字, "near": 近くの見出し（省ける）} の形で指定してください')
    pat = _pattern(target)
    paras = list(_paras(doc))
    hits = [(p, m.start(), m.end(), i) for i, p in enumerate(paras) for m in pat.finditer(p.text) if m.end() > m.start()]
    if not hits:
        raise CircleError(f"「{target}」が文書に見つかりません（窓の字のとおりに指定してください）")
    occ = spec.get("occurrence")
    if occ is not None:
        if not isinstance(occ, int) or isinstance(occ, bool) or not 1 <= occ <= len(hits):
            raise CircleError(f"occurrence は 1〜{len(hits)} で指定してください")
        return hits[occ - 1]
    near = spec.get("near")
    if len(hits) > 1 and isinstance(near, str) and near.strip():
        np_ = _pattern(near)
        anchors = [(p, m.start(), i) for i, p in enumerate(paras) for m in np_.finditer(p.text) if m.end() > m.start()]
        if not anchors:
            raise CircleError(f"near の「{near}」が見つかりません")

        def rank(h):
            best = (9, 10 ** 9)
            for ap, a0, ai in anchors:
                cp, ch = _cell_pos(ap), _cell_pos(h[0])
                if ap is h[0]:
                    k = (0, abs(h[1] - a0))
                elif cp and ch and cp[0] == ch[0] and cp[1] == ch[1]:
                    k = (1, abs(cp[2] - ch[2]))  # 表の同じ行
                elif cp and ch and cp[0] == ch[0] and cp[2] == ch[2]:
                    k = (2, abs(cp[1] - ch[1]))  # 表の同じ列
                else:
                    k = (3, abs(ai - h[3]))  # 段落の並びで近いもの
                best = min(best, k)
            return best

        scored = sorted(hits, key=rank)
        if rank(scored[0]) == rank(scored[1]):
            listing = "／".join(f"{k}番目：「{h[0].text.strip()[:24]}」" for k, h in enumerate(hits[:10], 1))
            raise CircleError(f"「{near}」に同じくらい近い「{target}」が 2 か所以上あります。occurrence で決めてください：{listing}")
        return scored[0]
    if len(hits) > 1:
        listing = "／".join(f"{k}番目：「{h[0].text.strip()[:24]}」" for k, h in enumerate(hits[:10], 1))
        raise CircleError(f"「{target}」が {len(hits)} か所にあります。near（近くの見出しの字）か occurrence（何番目か）で"
                          f"1つに決めてください：{listing}")
    return hits[0]


def _glyphs(doc: DocxDocument, p: Para) -> list[_G]:
    """段落の字の箱の見積もり（行の上を 0 とし、横は段落の初めからの pt。折り返しは見ない：囲む字の前後の字だけに使う）。"""
    out = []
    x = 0.0
    pos = p.positions()
    wide = east_asian_context(p.text)
    for k, ch in enumerate(p.text):
        sg = p.segs[pos[k][0]] if k < len(pos) else None
        size = run_size(doc, p, sg.run if sg is not None else None)
        w = width(ch, size, wide) if ch not in "\t\n" else size
        out.append(_G(ch, x, 0.0, x + w, size, size))
        x += w
    return out


def resolve(doc: DocxDocument, spec: dict, used: dict) -> dict:
    """○の所：{id, part, para（段落）, start, end, text, style, ellipse（囲む字の左上からの pt：中心と半径）, size, label}。"""
    style = spec.get("style", "ellipse")
    if style not in ("ellipse", "border"):
        raise CircleError('style は "ellipse"（楕円の図形）か "border"（字の囲み線）です')
    p, a, b, _i = find(doc, spec)
    pos = p.positions()
    for k in range(a, b):
        sg = p.segs[pos[k][0]]
        if sg.kind != "t" or sg.node is None or sg.run is None:
            raise CircleError(f"「{spec['circle']}」が、書き換えられない部分（{sg.kind}）にかかっています")
    gl = _glyphs(doc, p)
    mine = [g for g in gl[a:b] if not g.text.isspace()]
    ids = {id(g) for g in mine}
    # 同じ行の前後の字（前後 6 字まで。折り返しは見ないので、近くだけ）
    others = [g for g in gl[max(0, a - 6):b + 6] if id(g) not in ids and not g.text.isspace()]
    rect = (min(g.x0 for g in mine), 0.0, max(g.x1 for g in mine), max(g.y1 for g in mine))
    size = max(g.size for g in mine)
    pad = spec.get("pad")
    if pad is not None and not (isinstance(pad, (int, float)) and not isinstance(pad, bool) and 0 <= pad <= 3 * size):
        raise CircleError(f"pad は、字と○のすき間（pt、0〜{3 * size:.0f}）で指定してください")
    e = None
    if style == "ellipse":
        e, hit = shape(rect, size, others, None, pad)
        if e is None:
            raise CircleError(f"「{spec['circle']}」を○で囲むと、前後の字" + (f"「{hit.text}」" if hit else "") +
                              "に触れます。pad を小さくするか、style を \"border\"（字の囲み線）にしてください")
        e = (e[0] - rect[0], e[1], e[2], e[3])  # 囲む字の左の端からの位置にする
    key = (p.part, id(p.node))
    for (pk, aa, bb) in used.get("spans", []):
        if pk == key and aa < b and a < bb:
            raise CircleError(f"「{spec['circle']}」は、ほかの○と重なります")
    used.setdefault("spans", []).append((key, a, b))
    k = len(used["spans"]) + _existing(doc)
    label = spec.get("near") or (p.cell.text()[:30] if p.cell is not None else p.text.strip()[:30])
    wide = east_asian_context(p.text)
    near = p.text[max(0, a - 6):b + 6]
    notes = (["幅の変わる字（ラテン文字・キリル文字・アラビア文字など、書体で字ごとに幅が違う字）を含む（字の幅は見積もり）"]
             if any(not fixed_width(ch, wide) for ch in near if not ch.isspace()) else [])
    return {"id": f"○{k}", "part": p.part, "para": p, "start": a, "end": b, "text": p.text[a:b], "style": style,
            "ellipse": [round(v, 3) for v in e] if e else None, "size": size, "label": label,
            "context": re.sub(r"\s+", " ", p.text).strip()[:60], "notes": notes}


def _existing(doc: DocxDocument) -> int:
    n = 0
    for part in doc.parts:
        n += len(re.findall(rb'name="' + TAG.encode() + rb" ", doc.xml(part)))
    return n


# ---------------------------------------------------------------------- 書く
def _ids(doc: DocxDocument) -> int:
    best = 0
    for name in doc.zip.names():
        if name.endswith(".xml") and name.startswith("word/"):
            for m in re.finditer(rb"<(?:\w+:)?docPr\b[^>]*?\bid=\"(\d+)\"", doc.zip.read(name)):
                best = max(best, int(m[1]))
    return best


def drawing_run(pre: str, sd: dict, shape_id: int) -> bytes:
    cx, cy, ra, rb = sd["ellipse"]
    size = sd["size"]
    lw = max(0.6, min(1.2, 0.07 * size))
    # 横：字の位置からの左の端（負の数で、字の左へ）。縦：行の上から（字の箱は行の上から 0.1 字下にあると見積もる）
    left = cx - ra
    top = 0.1 * size + cy - rb
    w, h = 2 * ra, 2 * rb
    desc = f"○で囲んだ字「{sd['text']}」（とどく）"
    esc = lambda s: s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")  # noqa: E731
    return (f'<{pre}:r><{pre}:rPr><{pre}:noProof/></{pre}:rPr><{pre}:drawing>'
            f'<wp:anchor xmlns:wp="{NS["wp"]}" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{251659264 + shape_id}" '
            f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="character"><wp:posOffset>{round(left * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="line"><wp:posOffset>{round(top * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{round(w * EMU)}" cy="{round(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{shape_id}" name="{TAG} {esc(sd["id"])}" descr="{esc(desc)}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic xmlns:a="{NS["a"]}"><a:graphicData uri="{NS["wps"]}"><wps:wsp xmlns:wps="{NS["wps"]}"><wps:cNvSpPr/>'
            f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{round(w * EMU)}" cy="{round(h * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="{round(lw * EMU)}"><a:solidFill>'
            f'<a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>'
            f"</wp:anchor></{pre}:drawing></{pre}:r>").encode("utf-8")


BORDER = '<{p}:bdr {p}:val="single" {p}:sz="6" {p}:space="0" {p}:color="000000"/>'


def plan(doc: DocxDocument, circles: list[dict]) -> dict[str, list[tuple[int, int, bytes]]]:
    """書き換える部品 → [(始め, 終わり, 入れるバイト)]（文字の要素を分けて、○の図形を入れる・囲み線を付ける）。"""
    from .fill import _t
    from .xmlpos import prefix_of

    out: dict[str, list] = {}
    nid = _ids(doc) + 1
    by_run: dict[tuple[str, int], list] = {}
    for sd in circles:
        p = sd["para"]
        pos = p.positions()
        # 囲む字の各文字の要素：(要素の番号, その中の始め, 終わり)
        for k in range(sd["start"], sd["end"]):
            i, off = pos[k]
            sg = p.segs[i]
            lst = by_run.setdefault((sd["part"], sg.run.s), [sg.run, []])[1]
            lst.append((sg.node, off, sd, k == sd["start"]))
    for (part, _s), (run, items) in by_run.items():
        xml = doc.xml(part)
        pre = prefix_of(xml)
        rpr = run.find("rPr")
        rpr_b = xml[rpr.s:rpr.e] if rpr is not None else b""
        r_open = xml[run.s:run.se] if run.se != run.e else f"<{pre}:r>".encode()
        r_close = f"</{pre}:r>".encode()
        border = BORDER.format(p=pre).encode()
        if rpr is not None and rpr.se != rpr.e:
            rpr_bd = xml[rpr.s:rpr.se] + xml[rpr.se:rpr.es] + border + xml[rpr.es:rpr.e]
        elif rpr is not None:
            rpr_bd = xml[rpr.s:rpr.e][:-2].rstrip() + b">" + border + f"</{pre}:rPr>".encode()
        else:
            rpr_bd = f"<{pre}:rPr>".encode() + border + f"</{pre}:rPr>".encode()
        res = bytearray()
        buf = bytearray()
        mode = None  # 囲み線を付ける字の並びの中か

        def flush(rp: bytes):
            nonlocal buf
            if buf:
                res.extend(r_open + rp + bytes(buf) + r_close)
                buf = bytearray()

        by_node: dict[int, dict[int, tuple]] = {}
        for node, off, sd, first in items:
            by_node.setdefault(id(node), {})[off] = (sd, first)
        for kid in run.kids:
            if kid.local == "rPr":
                continue
            marks = by_node.get(id(kid))
            if not marks:
                if mode is not None:
                    flush(rpr_bd)
                    mode = None
                buf += xml[kid.s:kid.e]
                continue
            text = kid.text
            for off, ch in enumerate(text):
                m = marks.get(off)
                want = None if m is None or m[0]["style"] != "border" else "border"
                if m is not None and m[1] and m[0]["style"] == "ellipse":
                    flush(rpr_bd if mode == "border" else rpr_b)
                    mode = None
                    res.extend(drawing_run(pre, m[0], nid))
                    m[0]["shape_id"] = nid
                    nid += 1
                if want != mode:
                    flush(rpr_bd if mode == "border" else rpr_b)
                    mode = want
                buf += _t(pre, ch)
            # 一字ずつ入れた w:t を、同じ書式の並びでまとめる
        flush(rpr_bd if mode == "border" else rpr_b)
        merged = re.sub(rb"</" + pre.encode() + rb":t><" + pre.encode() + rb':t xml:space="preserve">', b"", bytes(res))
        out.setdefault(part, []).append((run.s, run.e, merged))
    return out


def apply(doc: DocxDocument, circles: list[dict]) -> bytes:
    from ..zipraw import RawZip

    edits = plan(doc, circles)
    z = RawZip(doc.blob)
    for part, ops_ in edits.items():
        data = doc.xml(part)
        for s, e, ins in sorted(ops_, key=lambda x: -x[0]):
            data = data[:s] + ins + data[e:]
        z.replace(part, data)
    return z.build()


# ---------------------------------------------------------------------- 証明と読み戻し
def prove(before: bytes, after: bytes, name: str, specs: list[tuple[str, dict]]) -> list[str]:
    """元の文書で○を決め直し、変わった部品が、決め直した書き換え（文字の要素を分けて○を入れる）だけかを確かめる。"""
    from ..zipraw import RawZip

    doc = DocxDocument(before, name)
    used: dict = {}
    circles = []
    out = []
    for fid, spec in specs:
        try:
            sd = resolve(doc, spec, used)
        except CircleError as e:
            return [f"○ {fid} の所が、元の文書で決まりません：{e}"]
        if sd["id"] != fid:
            out.append(f"○ {fid} が、決め直した所（{sd['id']}）と違います")
        circles.append(sd)
    edits = plan(doc, circles)
    za, zb = RawZip(before), RawZip(after)
    if za.names() != zb.names():
        out.append("ZIP のエントリの並びが変わっています")
    for n in zb.names():
        if n in edits:
            want = doc.xml(n)
            for s, e, ins in sorted(edits[n], key=lambda x: -x[0]):
                want = want[:s] + ins + want[e:]
            if zb.read(n) != want:
                out.append(f"{n}: ○を描くための書き換えのほかにも変わっています")
        elif za.has(n):
            ea, eb = za.get(n), zb.get(n)
            if (ea.local, ea.data, ea.descriptor) != (eb.local, eb.data, eb.descriptor):
                out.append(f"{n}: ○を描くのに変わってよい部品ではありません")
    return out


def read_back(blob: bytes, name: str, circled: list[dict]) -> list[dict]:
    """書いた後の文書を開き直し、○の図形（または囲み線）が、言った字の直前（その字だけ）に付いているかを確かめる。"""
    doc = DocxDocument(blob, name)
    out = []
    for c in circled:
        found = None
        for p in _paras(doc):
            xml = doc.xml(p.part)
            seg_text = p.text
            if c.get("style") == "border":
                pos = p.positions()
                runs = [k for k in range(len(seg_text)) if p.segs[pos[k][0]].run is not None and
                        (lambda r: r is not None and r.find("rPr") is not None and r.find("rPr").find("bdr") is not None)(p.segs[pos[k][0]].run)]
                if runs and c["context"] == re.sub(r"\s+", " ", p.text).strip()[:60]:
                    found = "".join(seg_text[k] for k in runs)
                    break
                continue
            for d in p.node.iter("docPr"):
                if d.get("name") == f"{TAG} {c['id']}":
                    # 図形の文字の要素の、すぐ後ろの字
                    run = d
                    while run is not None and run.local != "r":
                        run = run.parent
                    pos = p.positions()
                    after = [k for k in range(len(seg_text)) if p.segs[pos[k][0]].run is not None and p.segs[pos[k][0]].run.s > run.s]
                    k0 = after[0] if after else None
                    found = seg_text[k0:k0 + len(c["text"])] if k0 is not None else ""
                    ext = re.search(rb'<wp:extent cx="(\d+)" cy="(\d+)"', xml[run.s:run.e])
                    if ext:
                        w = int(ext[1]) / EMU
                        wide = east_asian_context(p.text)
                        tw = sum(width(ch, c["size"], wide) for ch in c["text"] if not ch.isspace())
                        if w < tw:
                            found = f"{found}（○の幅 {w:.1f}pt が、字の幅 {tw:.1f}pt より狭い）"
                    break
            if found is not None:
                break
        if found != c["text"]:
            out.append({"id": c["id"], "written": f"○「{c['text']}」",
                        "read_back": f"○の後ろの字「{found}」" if found is not None else "書いた後の文書に、この○が見つかりません"})
    return out
