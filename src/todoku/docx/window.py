"""Word の窓：本文を上から（表は行ごとに、マスを ｜ で区切って）、入力欄は［］で示す。

入力欄：
- 旧来のフォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN）とコンテンツコントロール（名前は欄の名前・タグ）
- 表の空のマス（左か上のマスの文字を見出しにする）：T{表}.R{行}.C{列}
- 段落の中の空欄（「令和　　年　　月　　日」、下線を引いた空白、「＿＿＿」）と □：P{段落}.{番号}
書き込んだ欄は、見えないしおり（_todoku_…）で囲み、書いた後も同じ名前で読めるようにする。
"""

from __future__ import annotations

import hashlib
import itertools
import re

from ..xlsx.window import AGENT_WHY, agent_blanks, inline_blanks, norm_label
from .document import Cell, Control, DocxDocument, Para, Table, iter_paras

DET = "deterministic"
LAYOUT = "layout"
AGENT = "agent"  # 規則では決められない、書き込む所かもしれない所（エージェントが読んで判断する）
BOX = set("□☐")
MARK_PREFIX = "_todoku_"

KIND_JA = {"text": "文字", "inline": "文字の中の空欄", "checkbox": "チェック", "dropdown": "一覧から選ぶ", "combo": "一覧から選ぶ（自由入力も可）",
           "date": "日付", "richtext": "文字（書式つき）", "picture": "画像", "cell": "文字"}


def _underlined_blanks(p: Para) -> list[tuple[int, int]]:
    """下線を引いた空白の並び（「氏名＿＿＿」を下線つきの空白で作ったもの）。"""
    out = []
    pos = 0
    start = None
    for s in p.segs:
        if s.hidden:
            continue
        for ch in s.text:
            is_blank = s.underline and ch in " \u3000\u2002\u2003" and s.kind == "t"
            if is_blank and start is None:
                start = pos
            elif not is_blank and start is not None:
                if pos - start >= 3:
                    out.append((start, pos))
                start = None
            pos += 1
    if start is not None and pos - start >= 3:
        out.append((start, pos))
    return out


def _trailing_blank(text: str) -> list[tuple[int, int]]:
    """短い見出しの後ろに、全角の空白が続いて段落が終わるもの（「住所　　　　　　　」）。"""
    m = re.search(r"[\u3000]{4,}$", text)
    if not m:
        return []
    head = text[:m.start()].strip(" \u3000")
    if not head or len(head) > 20 or head.endswith(("。", "様", "殿", "御中")):
        return []
    start = len(text[:m.start()].rstrip(" \u3000"))  # 「氏　名　  　　　」：見出しの後ろの空白はすべて空欄
    return [(start, m.end())]


def para_blanks(p: Para) -> list[tuple[int, int]]:
    from ..pdf.layout import ADDRESS_UNITS, pdf_blanks

    text = p.text
    spans = set(pdf_blanks(text)) | set(_underlined_blanks(p)) | set(_trailing_blank(text))
    # 住所のマス：「　　郡　市」のように住所の単位で終わる段落は、その後ろに町名・番地を書き足す
    core = text.rstrip(" \u3000")
    if p.cell is not None and core and core[-1] in ADDRESS_UNITS and len(core) <= 30 and core == text:
        spans.add((len(text), len(text)))
    kept: list[tuple[int, int]] = []
    for a, b in sorted(spans):
        if kept and a < kept[-1][1]:
            kept[-1] = (kept[-1][0], max(b, kept[-1][1]))
            continue
        kept.append((a, b))
    return kept


class Spots:
    """文書の入力欄の一覧（名前 → 欄の情報）。窓と書き込みで同じものを使う。"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc
        self.items: dict[str, dict] = {}
        self.para_index: dict[int, int] = {}
        self.ranges: dict = {}
        self.marks = doc.marks() if hasattr(doc, "marks") else _bookmarks(doc)
        self._collect()

    def _collect(self) -> None:
        doc = self.doc
        n = 0
        for _kind, _part, blocks in doc.all_blocks():
            for p in iter_paras(blocks):
                n += 1
                self.para_index[id(p)] = n
        # 1) フォームの欄・コンテンツコントロール
        used_names = set()
        for k, c in enumerate(doc.controls, 1):
            name = c.name or f"ctl{k}"
            base = name
            j = 2
            while name in used_names:
                name = f"{base}#{j}"
                j += 1
            used_names.add(name)
            self.items[name] = {"id": name, "kind": c.kind, "source": c.source, "control": c, "para": c.para,
                                "options": c.options or None, "provenance": DET}
        # 2) 表の空のマス
        tno = 0
        for _kind, part, blocks in doc.all_blocks():
            for t in _tables(blocks):
                tno += 1
                for row in t.rows:
                    for c in row:
                        if c.vmerge == "continue" or c.text() or _has_control(c):
                            continue
                        if not c.paras():
                            continue
                        label, rl, cl = _cell_label(t, c)
                        fid = f"T{tno}.R{c.row + 1}.C{c.col + 1}"
                        if not label:
                            # 見出しの見つからない空のマス：桁のマスのこともあれば、飾りのマスのこともある（エージェントが判断）
                            self.items[fid] = {"id": fid, "kind": "cell", "source": "layout", "cell": c, "para": c.paras()[0],
                                               "label": "", "provenance": AGENT, "why": AGENT_WHY["empty_cell"],
                                               "why_key": "empty_cell", "shade": c.shade}
                            continue
                        self.items[fid] = {"id": fid, "kind": "cell", "source": "layout", "cell": c, "para": c.paras()[0],
                                           "label": label, "row_label": rl, "column_label": cl, "provenance": LAYOUT,
                                           "shade": c.shade}
        # 3) 段落の中の空欄と □（空のマスとして欄にした段落は除く）
        cell_paras = {id(q) for it in self.items.values() if it.get("kind") == "cell" for q in it["cell"].paras()}
        for _kind, part, blocks in doc.all_blocks():
            for p in iter_paras(blocks):
                if id(p) in cell_paras:
                    continue
                if any(s.field is not None for s in p.segs) and all(s.field is not None or not s.text.strip() for s in p.segs):
                    continue
                pi = self.para_index[id(p)]
                text = p.text
                k = 0
                spans = []
                for i, ch in enumerate(text):
                    if ch in BOX and not _in_field(p, i):
                        spans.append((i, i + 1, "checkbox"))
                for a, b in para_blanks(p):
                    if not any(_in_field(p, i) for i in range(a, b)):
                        spans.append((a, b, "inline"))
                # 書き込み済みの欄（しおり）は、その番号を持ったまま（4 で作る）。残りの空欄に、しおりの番号を除いた番号を
                # 位置の順に割り当てる（元の白紙のときと同じ id になる：P16.1 に書いた後も、P16.2・P16.3 はそのまま）
                mark_k: dict[int, tuple[int, int]] = {}
                mark_ka: dict[int, tuple[int, int]] = {}
                for name, (q, x, y) in self.marks.items():
                    mm = re.fullmatch(rf"P{pi}\.(a?)(\d+)", name) if q is p else None
                    if mm:
                        (mark_ka if mm[1] else mark_k)[int(mm[2])] = (x, y)
                written = [(x, y) for q, x, y in self.marks.values() if q is p]
                spans = [t for t in spans if not _overlaps(t[0], t[1], written)]
                free_k = (n for n in itertools.count(1) if n not in mark_k)
                for a, b, kind in sorted(spans):
                    k = next(free_k)
                    fid = f"P{pi}.{k}"
                    if kind == "checkbox":
                        m = re.search(r"[\u3000 ]{2,}|□|☐|※", text[b:])
                        opt = text[b:b + m.start()] if m else text[b:]
                        label = norm_label(opt.strip(" \u3000")) or norm_label(text[:a].strip(" \u3000"))[-30:]
                    else:
                        label = self._label(p, a, b)
                    self.items[fid] = {"id": fid, "kind": kind, "source": "layout", "para": p, "span": (a, b),
                                       "label": label, "provenance": LAYOUT}
                # 規則では決められない所：窓に「？」をつけて示し、エージェントが読んで判断する（id は P段落.a番号）
                if any(s_.field is not None for s_ in p.segs):
                    continue
                if not text.strip(" \u3000") == "" and not "".join(
                        s_.text for s_ in p.segs if not s_.hidden and s_.kind in ("t", "sym", "tab")).strip(" \u3000\t"):
                    continue  # 図・埋め込みの印だけの段落（窓のために作った字だけ）
                free_ka = (n for n in itertools.count(1) if n not in mark_ka)
                found_agent = agent_blanks(text, [(x, y) for x, y, _k in spans] + written, p.cell is not None)
                # 見出しのあるマスの中の空の段落（「〒」の下の行に住所を書くなど）
                c = p.cell
                if c is not None and not text.strip(" \u3000") and not spans and not written:
                    others = [q.text.strip(" \u3000") for q in c.paras() if q is not p and q.text.strip(" \u3000")]
                    before = [q for q in c.paras()[:c.paras().index(p)] if q.text.strip(" \u3000")] if p in c.paras() else []
                    if others and before and all(len(o) <= 16 for o in others):
                        found_agent = [(0, len(text), "next_line")]
                for a, b, why in found_agent:
                    if any(_in_field(p, i) for i in range(a, max(a + 1, b))) or self._beside_cell_field(p, why):
                        continue
                    fid = f"P{pi}.a{next(free_ka)}"
                    it = {"id": fid, "kind": "inline", "source": "layout", "para": p, "span": (a, b),
                          "label": self._label(p, a, b), "provenance": AGENT, "why": AGENT_WHY[why], "why_key": why}
                    if why == "placeholder":
                        it["placeholder"] = text[a:b]
                    self.items[fid] = it
        # 3') マスの中の見出し（「住所」「住所：」「所在地　〒」）の後ろ：すぐ右のマスに規則の欄があれば、値はそちらに書く
        ruled_cells = set()
        for it in self.items.values():
            if it.get("provenance") == AGENT:
                continue
            c_ = it.get("cell") or (it["para"].cell if it.get("para") is not None else None)
            if c_ is not None:
                ruled_cells.add(id(c_))
        for fid in [f for f, it in self.items.items() if it.get("why_key") in ("after_label", "after_colon", "after_mark")]:
            c_ = self.items[fid]["para"].cell
            if c_ is None or c_.table is None:
                continue
            right = [x for x in c_.table.rows[c_.row] if x.col == c_.col + c_.span]
            if right and id(right[0]) in ruled_cells:
                del self.items[fid]
        # 4) 書き込み済みの欄（しおり）：同じ名前の欄が見つからなければ、しおりから作る
        for name, (p, a, b) in self.marks.items():
            if name in self.items:
                continue
            kind = "checkbox" if p.text[a:b] in ("■", "☑", "☒", "□", "☐") else ("cell" if name.startswith("T") else "inline")
            label = self._label(p, a, b) if kind == "inline" else ""
            if kind == "cell" and p.cell is not None and p.cell.table is not None:
                label = _cell_label(p.cell.table, p.cell)[0]  # 書き込み済みのマスも、表の行・列の見出しで示す
            self.items[name] = {"id": name, "kind": kind, "source": "layout", "para": p, "span": (a, b), "filled": True,
                                "label": label, "provenance": LAYOUT}

    def _beside_cell_field(self, p: Para, why: str) -> bool:
        """マスの中の「住所：」のすぐ右のマスが空のマスの欄（そこに書く）か、「人」だけのマスの下に空のマスの欄がある
        （列の見出し）。どちらも、規則で見つけた欄がもうあるので、出さない。"""
        c = p.cell
        if c is None or why not in ("after_colon", "before_unit", "after_label", "after_mark"):
            return False
        for it in self.items.values():
            x = it.get("cell")
            if x is None or it.get("kind") != "cell" or x.table is not c.table:
                continue
            if why in ("after_colon", "after_label", "after_mark") and x.row == c.row and x.col == c.col + c.span \
                    and self.items.get(it["id"], {}).get("provenance") != AGENT:
                return True
            if why == "before_unit" and x.row > c.row and x.col <= c.col < x.col + x.span:
                return True
        return False

    def _label(self, p: Para, a: int, b: int) -> str:
        """段落の中の空欄の見出し：その空欄を［　］にした段落の文字。書き込み済みのほかの欄の値は除く（白紙のときと同じ見出し）。"""
        cuts = [(x, y) for q, x, y in self.marks.values() if q is p and not (x == a and y == b) and y > x]
        out, at = [], 0
        for x, y in sorted(cuts + [(a, b)]):
            if x < at:
                continue
            out.append(p.text[at:x])
            out.append("［　］" if (x, y) == (a, b) else "　")
            at = y
        out.append(p.text[at:])
        return norm_label("".join(out))

    def value(self, item: dict):
        """いまの値（書き込み済みなら、その値）。"""
        name = item["id"]
        if name in self.marks:
            p, a, b = self.marks[name]
            t = p.text[a:b]
            if item["kind"] == "checkbox":
                return t in ("■", "☑", "☒")
            return t
        c: Control | None = item.get("control")
        if c is not None:
            if c.kind == "checkbox":
                return c.checked
            if c.placeholder:
                return None
            t = "".join(s.text for s in c.result_segs)
            if c.kind == "dropdown" and c.source == "formfield":
                return c.value
            return t.strip("\u2002 ") if c.source == "formfield" else t
        if item["kind"] == "checkbox":
            return False
        return None


def _bookmarks(doc: DocxDocument) -> dict[str, tuple[Para, int, int]]:
    """書き込んだ欄のしおり（_todoku_…）：名前 → (段落, 始まり, 終わり)（段落の文字の位置）。"""
    out: dict[str, tuple[Para, int, int]] = {}
    for _kind, part, blocks in doc.all_blocks():
        for p in iter_paras(blocks):
            starts = {}
            pos = 0
            seg_at = {}
            for s in p.segs:
                if s.run is not None:
                    seg_at.setdefault(s.run.s, pos)
                if not s.hidden:
                    pos += len(s.text)
            node = p.node
            # 段落の中の順に、しおりと文字の位置を数える
            cur = 0
            order = []
            for n in node.iter():
                if n.local == "bookmarkStart" and (n.get("name") or "").startswith(MARK_PREFIX):
                    order.append(("s", n.get("id"), n.get("name")[len(MARK_PREFIX):], n.s))
                elif n.local == "bookmarkEnd":
                    order.append(("e", n.get("id"), None, n.s))
            if not order:
                continue
            for s_ in p.segs:
                pass
            # 文字の位置は、各 seg のバイト位置で比べる
            seg_pos = []
            pos = 0
            for s in p.segs:
                if s.hidden:
                    continue
                bpos = s.node.s if s.node is not None else (s.run.s if s.run is not None else 0)
                seg_pos.append((bpos, pos, len(s.text)))
                pos += len(s.text)

            def char_at(byte: int) -> int:
                k = 0
                for bpos, cpos, ln in seg_pos:
                    if bpos < byte:
                        k = cpos + ln
                    else:
                        break
                return k

            for kind, bid, name, byte in order:
                if kind == "s":
                    starts[bid] = (name, char_at(byte))
                elif bid in starts:
                    name, a = starts.pop(bid)
                    out[name.replace("_", ".")] = (p, a, char_at(byte))
            del cur, seg_at
    return out


class WriteError(ValueError):
    pass


def resolve_write(sp: "Spots", spec: dict, used: dict | None = None) -> dict:
    """エージェントが字で指した書き込み → 欄の情報（段落・位置・id「P段落.w番号」・書く字）。書き込みと証明で同じものを使う。

    spec：{"old": 段落の中の元の字, "new": 書き換えた後の字, "near": 近くの欄の id（同じ表の行か、前後 3 段落にしぼる）,
    "occurrence": 何番目か（しぼった中の、文書の順。1 から）}。old と new の前後の同じ字は、元の字と書式のまま残し、
    違う所だけを書き換える。元の字が見つからない・いくつもあって決まらないときは、候補を添えて断る。
    used：同じ書き込みの中で、もう使った番号と範囲（段落の番号 → [(番号, 始め, 終わり)]）。"""
    old, new = spec.get("old"), spec.get("new")
    if not isinstance(old, str) or not old or not isinstance(new, str):
        raise WriteError('字を目印にした書き込みは {"old": 段落の中の元の字, "new": 書き換えた後の字} の形で指定してください')
    if any(ch in old + new for ch in "\r\n\x07\x0b\x0c"):
        raise WriteError("old・new は、1つの段落の中の字にしてください（段落・行の区切りは書けません）")
    if old == new:
        raise WriteError("old と new が同じです")
    paras = [(sp.para_index[id(q)], q) for _k, _part, blocks in sp.doc.all_blocks() for q in iter_paras(blocks)]
    hits = [(n, q, m.start()) for n, q in paras for m in re.finditer(re.escape(old), q.text)]
    near = spec.get("near")
    if near:
        it = sp.items.get(str(near))
        if it is None:
            raise WriteError(f"near に指定した欄 {near} がありません")
        q0 = it.get("para") or (it["control"].para if it.get("control") is not None else None)
        n0 = sp.para_index.get(id(q0)) if q0 is not None else None
        c0 = q0.cell if q0 is not None else None

        def close(n: int, q: Para) -> bool:
            if c0 is not None:
                return q.cell is not None and q.cell.table is c0.table and q.cell.row == c0.row
            return n0 is not None and abs(n - n0) <= 3

        hits = [h for h in hits if close(h[0], h[1])]
    if not hits:
        raise WriteError(f"「{old}」が、{f'欄 {near} の近く' if near else '文書'}に見つかりません（窓の字のとおりに指定してください）")
    occ = spec.get("occurrence")
    if occ is not None:
        if not isinstance(occ, int) or not 1 <= occ <= len(hits):
            raise WriteError(f"occurrence は 1〜{len(hits)} で指定してください")
        hits = [hits[occ - 1]]
    if len(hits) > 1:
        listing = "／".join(f"{k}番目：段落 {n}「{q.text.strip()[:30]}」" for k, (n, q, _i) in enumerate(hits[:10], 1))
        raise WriteError(f"「{old}」が {len(hits)} か所にあります。near（近くの欄の id）か occurrence（何番目か）で、"
                         f"1つに決めてください：{listing}")
    n, q, i = hits[0]
    # 目印は本文の字（字・タブ・記号の要素）だけ：図の説明（〔図: …〕）・脚注の印・改行などの、窓のために作った字は目印にしない
    pos = q.positions()
    kinds = {q.segs[pos[k][0]].kind for k in range(i, min(i + len(old), len(pos)))}
    if not kinds <= {"t", "tab", "sym"}:
        raise WriteError(f"「{old}」は、本文の字でない所（図の説明・脚注の印など）にかかっています。本文の字で指定してください")
    k0 = 0
    while k0 < min(len(old), len(new)) and old[k0] == new[k0]:
        k0 += 1
    k1 = 0
    while k1 < min(len(old), len(new)) - k0 and old[len(old) - 1 - k1] == new[len(new) - 1 - k1]:
        k1 += 1
    a, b = i + k0, i + len(old) - k1
    value = new[k0:len(new) - k1]
    if any(_in_field(q, j) for j in range(max(0, a - (a == b)), max(a + 1, b))):
        raise WriteError("フォームの欄の中・すぐ隣は、その欄の id で書いてください")
    used = used if used is not None else {}
    for _k, x, y in used.get(n, []):
        if _overlaps(a, b, [(x, y)]):
            raise WriteError(f"同じ段落 {n} の、ほかの書き込みと重なります")
    taken = {int(m[1]) for name, (q_, _x, _y) in sp.marks.items() if q_ is q
             for m in [re.fullmatch(rf"P{n}\.w(\d+)", name)] if m}
    taken |= {k for k, _x, _y in used.get(n, [])}
    k = next(x for x in itertools.count(1) if x not in taken)
    used.setdefault(n, []).append((k, a, b))
    return {"id": f"P{n}.w{k}", "kind": "inline", "source": "agent_write", "para": q, "span": (a, b), "value": value,
            "label": sp._label(q, a, b), "provenance": AGENT, "anchored": dict(spec)}


def _overlaps(a: int, b: int, spans) -> bool:
    """[a, b) が spans のどれかと重なるか（長さ 0 の位置は、その位置を含む・端に接する所も重なりとみる）。"""
    return any((a < y and x < b) or (a == b and x <= a <= y) or (x == y and a <= x <= b) for x, y in spans)


def _in_field(p: Para, i: int) -> bool:
    k = 0
    for s in p.segs:
        if s.hidden:
            continue
        if k <= i < k + len(s.text):
            return s.field is not None
        k += len(s.text)
    return False


def _has_control(c: Cell) -> bool:
    return any(s.field is not None for p in c.paras() for s in p.segs)


def _tables(blocks):
    for b in blocks:
        if isinstance(b, Table):
            yield b
            for row in b.rows:
                for c in row:
                    yield from _tables(c.blocks)
        elif isinstance(b, Para) and b.boxes:
            yield from _tables(b.boxes)


def _cell_label(t: Table, c: Cell) -> tuple[str, str, str]:
    row = t.rows[c.row]
    left = None
    for x in row:
        if x.col < c.col and x.text():
            if left is None or x.col > left.col:
                left = x
    up = None
    for r in range(c.row - 1, -1, -1):
        for x in t.rows[r]:
            if x.col <= c.col < x.col + x.span and x.text():
                up = x
                break
        if up is not None:
            break
    rl = norm_label(left.text()) if left else ""
    cl = norm_label(up.text()) if up else ""
    adjacent_left = left is not None and left.col + left.span == c.col
    if rl and cl and rl != cl and not (adjacent_left and up is not None and c.row - up.row > 2):
        return f"{rl}／{cl}", rl, cl
    return rl or cl, rl, cl


# ====================================================================== 窓
def editable(doc: DocxDocument, sp: "Spots", item: dict) -> bool:
    """文書の保護のもとで、この欄に書けるか（フォームだけ・読み取り専用＋編集できる範囲）。"""
    if not doc.protection_enforced or not doc.protection:
        return True
    if doc.protection == "forms":
        return item.get("control") is not None
    if doc.protection != "readOnly" or not hasattr(doc, "xml"):
        return False
    ranges = sp.ranges.setdefault("_", {})
    p = item.get("para")
    c = item.get("control")
    part = (c.part if c is not None else (p.part if p is not None else doc.main))
    if part not in ranges:
        xml = doc.xml(part)
        starts, out = {}, []
        for m in re.finditer(rb'<(?:\w+:)?perm(Start|End)\b[^>]*?\bw?:?id="([^"]+)"[^>]*>', xml):
            if m[1] == b"Start":
                starts[m[2]] = m.start()
            elif m[2] in starts:
                out.append((starts.pop(m[2]), m.end()))
        ranges[part] = out
    if c is not None:
        poss = [sg.node.s for sg in c.result_segs if sg.node is not None] or [c.node.s]
    elif "span" in item:
        pos = p.positions()
        a = item["span"][0]
        poss = [p.segs[pos[a][0]].node.s] if a < len(pos) and p.segs[pos[a][0]].node is not None else [p.node.s]
    elif p is not None:
        poss = [sg.node.s for sg in p.segs if sg.node is not None] or [p.node.s]
    else:
        return False
    return any(a <= poss[0] <= b for a, b in ranges[part])


def build(doc: DocxDocument) -> dict:
    sp = Spots(doc)
    by_para: dict[int, list[dict]] = {}
    for it in sp.items.values():
        p = it.get("para")
        if p is not None:
            by_para.setdefault(id(p), []).append(it)
    out = {
        "todoku": "0.2",
        "kind": getattr(doc, "kind", "docx"),
        "source": {"file": doc.filename, "sha256": hashlib.sha256(doc.blob).hexdigest(), "bytes": len(doc.blob),
                   "note": "元のファイルが正本。書き込みは、書いた欄の部分だけを変える差分として適用する。"},
        "notation": {"［T1.R3.C2］": "表の空のマス（1つ目の表の3行目・2列目）", "［P12.1］": "段落の中の空欄（12番目の段落の1つ目）",
                     "［P12.a1？］": "規則では決められない、書き込む所かもしれない所（エージェントが読んで判断する。書くなら id を指定する。"
                                    "書かなければ何も変わらない。［P12.a1？＝○○］は、伏せ字の○○を書き換える所）",
                     "［名前＝値］": "入力欄といまの値", "☐［…］ / ☑": "チェック", " ｜ ": "表のマスの区切り"},
        "protection": {"edit": doc.protection, "enforced": doc.protection_enforced},
        "macros": _macros(doc),
        "stories": [],
        "fields": [],
    }
    for kind, part, blocks in doc.all_blocks():
        lines: list[str] = []
        plain: list[str] = []
        _render_blocks(blocks, sp, by_para, lines, plain)
        if lines or kind == "body":
            out["stories"].append({"kind": kind, "part": part, "reading": lines, "_plain": plain})
    out["embeds"] = _embeds(doc)
    contexts = {}
    for st in out["stories"]:
        for line in st["reading"]:
            for m in re.finditer(r"［([^］＝？]+)？?(?:＝[^］]*)?］", line):
                contexts.setdefault(m[1], line)
    for it in sp.items.values():
        c = it.get("control")
        kind = it["kind"]
        label = it.get("label") or ""
        if c is not None:
            label = c.help or c.name or _context_label(it, contexts)
        d = {"id": it["id"], "kind": "text" if kind == "cell" else kind, "label": label, "source": it["source"],
             "provenance": it["provenance"], "value": sp.value(it), "context": contexts.get(it["id"], "")}
        if it.get("options"):
            d["options"] = it["options"]
        if c is not None and c.maxlen:
            d["maxlen"] = c.maxlen
        if c is not None and c.kind == "date" and c.date_format:
            d["format"] = c.date_format
        if it.get("provenance") == AGENT:
            d["agent"] = True
            d["why"] = it["why"]
            if it.get("placeholder"):
                d["placeholder"] = it["placeholder"]
        if it.get("row_label"):
            d["row_label"] = it["row_label"]
        if it.get("column_label"):
            d["column_label"] = it["column_label"]
        if it.get("shade") and it["shade"] not in ("auto", "FFFFFF"):
            d["shaded"] = it["shade"]
        if c is not None and c.lock in ("contentLocked", "sdtContentLocked"):
            d["readonly"] = True
        if not editable(doc, sp, it):
            d["readonly"] = True
            d["protected"] = True  # 文書の保護で、編集できる範囲の外
        out["fields"].append(d)
    return out


def _embeds(doc) -> list[dict]:
    """文書に埋め込んだ Excel の表：Excel の窓と同じ読み方で、シートごとに平文にする（本文では〔埋め込み1：…〕の所）。"""
    books = doc.embedded_books() if hasattr(doc, "embedded_books") else []
    out = []
    for e, wb in books:
        from ..xlsx import window as xw

        try:
            xdoc = xw.build(wb)
        except Exception as ex:  # noqa: BLE001
            out.append({"n": e["n"], "label": e["label"], "prog": e["prog"], "storage": e["storage"], "error": str(ex)})
            continue
        out.append({"n": e["n"], "label": e["label"], "prog": e["prog"], "storage": e["storage"],
                    "sheets": [sh["name"] for sh in xdoc["sheets"]],
                    "markdown": [ln for sh in xdoc["sheets"] for ln in xw._md_sheet(sh)]})
    return out


def _context_label(it: dict, contexts: dict) -> str:
    line = contexts.get(it["id"], "")
    m = re.search(r"([^｜［］]{1,30})［" + re.escape(it["id"]), line)
    return norm_label(m[1]) if m else ""


def _render_blocks(blocks, sp: Spots, by_para, lines: list[str], plain: list[str]) -> None:
    for b in blocks:
        if isinstance(b, Para):
            t = _render_para(b, sp, by_para)
            if t.strip():
                lines.append(t.strip())
            plain.append(b.raw_text())
            if b.boxes:
                sub: list[str] = []
                _render_blocks(b.boxes, sp, by_para, sub, plain)
                for s in sub:
                    lines.append(f"（テキストボックス）{s}")
        elif isinstance(b, Table):
            for row in b.rows:
                parts = []
                for c in row:
                    if c.vmerge == "continue":
                        continue
                    cid = _cell_field(sp, c)
                    sub: list[str] = []
                    for blk in c.blocks:
                        if isinstance(blk, Para):
                            t = _render_para(blk, sp, by_para)
                            if t.strip():
                                sub.append(t.strip())
                            plain.append(blk.raw_text())
                            if blk.boxes:
                                bx: list[str] = []
                                _render_blocks(blk.boxes, sp, by_para, bx, plain)
                                sub += [f"（テキストボックス）{x}" for x in bx]
                        elif isinstance(blk, Table):
                            inner: list[str] = []
                            _render_blocks([blk], sp, by_para, inner, plain)
                            sub.append("［表］" + " ／ ".join(inner))
                    cell_text = " ".join(sub)
                    if cid is not None and not cell_text:
                        cell_text = f"［{cid}？］" if sp.items[cid].get("provenance") == AGENT else f"［{cid}］"
                    parts.append(cell_text)
                if any(x.strip() for x in parts):
                    lines.append(" ｜ ".join(x for x in parts))


def _cell_field(sp: Spots, c: Cell) -> str | None:
    for it in sp.items.values():
        if it.get("cell") is c:
            return it["id"]
    return None


def _render_para(p: Para, sp: Spots, by_para) -> str:
    text = p.text
    items = by_para.get(id(p), [])
    edits = []
    for it in items:
        if it.get("cell") is not None:
            continue
        name = it["id"]
        if name in sp.marks:
            _p, a, b = sp.marks[name]
            v = text[a:b]
            if it["kind"] == "checkbox":
                edits.append((a, b, ("☑" if v in ("■", "☑", "☒") else "☐") + f"［{name}］"))
            else:
                edits.append((a, b, f"［{name}＝{v}］"))
            continue
        if "span" in it:
            a, b = it["span"]
            if it["kind"] == "checkbox":
                edits.append((a, b, f"☐［{name}］"))
            elif it.get("provenance") == AGENT:
                edits.append((a, b, f"［{name}？＝{text[a:b]}］" if it.get("placeholder") else f"［{name}？］"))
            else:
                edits.append((a, b, f"［{name}］"))
            continue
        c: Control | None = it.get("control")
        if c is not None:
            v = sp.value(it)
            # 欄の結果の文字の範囲
            a, b = _ctl_span(p, c)
            if c.kind == "checkbox":
                edits.append((a, b, ("☑" if v else "☐") + f"［{name}］"))
            else:
                edits.append((a, b, f"［{name}＝{v}］" if v not in (None, "") else f"［{name}］"))
    edits.sort()
    out, at = [], 0
    for a, b, rep in edits:
        if a < at:
            continue
        out.append(text[at:a])
        out.append(rep)
        at = b
    out.append(text[at:])
    t = "".join(out)
    hid = p.hidden_text().strip()
    if hid:
        t += f"〔非表示の文字: {hid}〕"
    if p.label:
        t = f"{p.label} {t}" if not p.label.endswith((" ", "\t", "　")) else p.label + t
    return t.replace("\t", " ")


def _ctl_span(p: Para, c: Control) -> tuple[int, int]:
    pos = 0
    a = b = None
    for s in p.segs:
        if s.hidden:
            continue
        if s.field is c:
            if a is None:
                a = pos
            b = pos + len(s.text)
        pos += len(s.text)
    if a is None:
        return len(p.text), len(p.text)
    return a, b


# ====================================================================== Markdown
def _md(s) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")


STORY_JA = {"body": "本文", "header": "ヘッダー", "footer": "フッター", "footnotes": "脚注", "endnotes": "文末脚注", "comments": "コメント"}


def _macros(doc):
    from ..macros import macro_info

    return macro_info(getattr(doc, "vba", None), doc.has_macros)


def to_markdown(doc: dict) -> str:
    out = [f"# {doc['source']['file'] or '書類'}", ""]
    out.append(f"- 元ファイルの SHA-256: `{doc['source']['sha256']}`（{doc['source']['bytes']:,} バイト）、Word")
    out.append("- この文書は読むための窓です。書き込みは `todoku fill` で、書いた欄の部分だけを変えます。")
    out.append("- 欄として示していない所に書くときは、字を目印にして書けます："
               "`{\"@1\": {\"old\": \"段落の中の元の字\", \"new\": \"書き換えた後の字\"}}`（old がいくつもあるときは、"
               "`near`（近くの欄の id）か `occurrence`（何番目か）で1つに決める）。変わるのは old と new の違う所だけです。")
    out.append("- 記号: " + "、".join(f"{k} は{v}" for k, v in doc["notation"].items()))
    pr = doc["protection"]
    if pr["enforced"] and pr["edit"]:
        out.append({"forms": "- **編集の制限：フォームの欄だけ書き込めます**",
                    "readOnly": "- **編集の制限：読み取り専用（編集できる範囲だけ書き込めます）**"}.get(pr["edit"], f"- 編集の制限: {pr['edit']}"))
    from ..macros import summary_line

    line = summary_line(doc["macros"])
    if line:
        out.append("- " + line.replace("ブック", "文書"))
    out.append("")
    if doc["fields"]:
        out += [f"## 入力欄（{len(doc['fields'])} か所）", "", "| 欄 | 見出し | 種類 | いまの値 |", "|---|---|---|---|"]
        for f in doc["fields"]:
            k = KIND_JA.get(f["kind"], f["kind"])
            if f.get("options"):
                k += ": " + "／".join(map(str, f["options"][:15]))
            if f.get("maxlen"):
                k += f"（{f['maxlen']} 字まで）"
            if f.get("format"):
                k += f"（書式 {f['format']}）"
            if f.get("agent"):
                k = f"規則では決められない：{f['why']}"
            if f.get("protected"):
                k += "（文書の保護で書けない欄）"
            elif f.get("readonly"):
                k += "（変更できない欄）"
            v = f.get("value")
            vs = ("☑ あり" if v else "☐ なし") if f["kind"] == "checkbox" else ("" if v in (None, "") else str(v))
            out.append(f"| {_md(f['id'])} | {_md(f['label'])} | {_md(k)} | {_md(vs)} |")
        out.append("")
    for st in doc["stories"]:
        out += [f"## {STORY_JA.get(st['kind'], st['kind'])}（{st['part']}）", ""]
        for line in st["reading"]:
            out.append(f"- {line}")
        out.append("")
    for e in doc.get("embeds") or []:
        out += [f"## 埋め込み{e['n']}：{e['label']}（{e['prog']}・{e['storage']}）", ""]
        if e.get("error"):
            out += [f"- この埋め込みは開けませんでした：{e['error']}", ""]
            continue
        out += ["- 本文の〔埋め込み" + str(e["n"]) + "：…〕の所にある表です。Excel の窓と同じ書き方（［C4］はセル）で出しています。", ""]
        out += [ln.replace("## シート", "### シート", 1) if ln.startswith("## シート") else
                (ln.replace("### ", "#### ", 1) if ln.startswith("### ") else ln) for ln in e["markdown"]]
        out.append("")
    from ..macros import markdown as macro_md

    out += macro_md(doc.get("macros"))
    return "\n".join(out)


def markdown_parts(doc: dict, budget: int = 120_000) -> list[tuple[str, str]]:
    whole = to_markdown(doc)
    if len(whole) <= budget:
        return [("全体", whole)]
    lines = whole.split("\n")
    parts = []
    cur: list[str] = []
    for line in lines:
        if cur and sum(len(x) + 1 for x in cur) + len(line) > budget - 2000:
            parts.append(cur)
            cur = []
        cur.append(line)
    if cur:
        parts.append(cur)
    head = lines[:6]
    index = head + ["", "## この窓の目次", "", f"窓が大きいので、{len(parts)} の部分に分けました。", ""]
    index += [f"- 部分 {i}: {sum(len(x) + 1 for x in p):,} 文字" for i, p in enumerate(parts, 1)]
    return [("目次", "\n".join(index))] + [(f"部分 {i}", "\n".join(head[:1] + [""] + p)) for i, p in enumerate(parts, 1)]
