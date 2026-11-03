"""Word の窓：本文を上から（表は行ごとに、マスを ｜ で区切って）、入力欄は［］で示す。

入力欄：
- 旧来のフォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN）とコンテンツコントロール（名前は欄の名前・タグ）
- 表の空のマス（左か上のマスの文字を見出しにする）：T{表}.R{行}.C{列}
- 段落の中の空欄（下線を引いた空白、「＿＿＿」）と □：P{段落}.{番号}
- 字の並びから推定した空欄（「令和　　年　　月　　日」・住所の単位の前後・段落の終わりの空白など）と、チェック済みの印らしい字
  （☑・■・✓）は、同じ P{段落}.{番号} の並びで、規則では決めない所（「？」）として示す（エージェントが判断する）
- 規則の型に合わない、書き込む所かもしれない所：P{段落}.a{番号}（「？」）。すぐ隣の欄の代わりの所は alternative_to で示す
書き込んだ欄は、見えないしおり（_todoku_…）で囲み、書いた後も同じ名前で読めるようにする。
埋め込んだ Excel の表は、Excel の窓と同じ読み方で節に出し、書くときの欄の名前（E1!Sheet1!C4）と、
Word が表示・印刷に使う表の絵を描き直せるか（描き直せなければ、その理由）を添える。
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
BOX = set("□☐")  # 空のチェックの印（規則のチェックボックス）
# 印のついたチェックの印らしい字：チェック済みの欄かもしれないが、■ は箇条の印・飾りのこともあり、✓ は本文の印のこともある。
# 規則では決めず、チェック済み（checked）のチェックボックスの候補として「？」をつけて示す（エージェントが判断する）
CHECKED = set("☑☒■✓✔")
UNCHECK = {"☑": "☐", "☒": "☐", "■": "□", "✓": "□", "✔": "□"}  # チェックを外すときに書く字（印のない同じ形の箱）
MARK_PREFIX = "_todoku_"
BULLET_CHECKED = frozenset("■☑☒")


def is_bullet(item: dict) -> bool:
    """箇条書きの印の □ の欄（P段落.b）。"""
    return item.get("kind") == "checkbox" and str(item.get("id", "")).endswith(".b") and "span" not in item

KIND_JA = {"text": "文字", "inline": "文字の中の空欄", "checkbox": "チェック", "dropdown": "一覧から選ぶ", "combo": "一覧から選ぶ（自由入力も可）",
           "date": "日付", "richtext": "文字（書式つき）", "picture": "画像", "cell": "文字"}

# Word の窓で足した、規則では決められない所の理由（xlsx/window.py の AGENT_WHY に足す。鍵は why_key）
WORD_WHY = {
    **AGENT_WHY,
    "keyword_blank": "空白の前後の字（年・月・日・号・円・「第」・かっこ・「〒」・住所の単位など）から、空欄と推定した所"
                     "（字の並びからの推定で、書類の作りからは決まりません）",
    "address_end": "住所の単位（市・区・町・村・郡・都・道・府・県）で終わるマスの、字の後ろ（町名・番地を書き足す所かもしれません）",
    "trailing_blank": "見出しの後ろに、全角の空白が続いて段落が終わる所（「住所　　　　　」。字の並びからの推定）",
    "trailing_space": "段落の終わりの2つ以上の空白（見出しが長い・「。」「様」「殿」「御中」で終わる・全角の空白が4つ未満など、"
                      "「住所　　　　」の型には合わない所）",
    "short_underline": "下線を引いた短い空白（1〜2字。下線の空欄は3字以上のことが多い）",
    "checked_mark": "印のついたチェックの印らしい字（☑・☒・■・✓。いまはチェック済みとして読む。■ は箇条の印・飾りのこともある）",
    "underscore_in_word": "英数字にはさまれた下線の字（「＿＿」。URL・ファイルの名前の一部のこともある）",
    "next_line_long": "見出しのあるマスの中の、空の段落（マスのほかの字が長く、見出しではなく説明の文のこともある）",
}


def box_mark(orig: str, on: bool) -> str:
    """チェックボックスの字 orig に、チェックを入れる（on）・外すときに書く字（変えないときは orig のまま）。
    □ は ■、☐ は ☑ にする。チェック済みの印（☑・■・✓ など）は、外すときに印のない箱（☐・□）にする。"""
    if orig in CHECKED:
        return orig if on else UNCHECK[orig]
    return ("☑" if orig == "☐" else "■") if on else orig


def _underlined_blanks(p: Para) -> list[tuple[int, int, bool]]:
    """下線を引いた空白の並び（「氏名＿＿＿」を下線つきの空白で作ったもの）：(始め, 終わり, 3字以上か)。
    3字以上は書類の作り（下線の空欄）として規則の欄。1〜2字は、下線つきの見出しの字間のこともあるので、規則では決めない（エージェント）。
    下線つきの字にはさまれた空白（下線を引いた字の中の、語の間の空白）は、空欄の形ではないので出さない。"""
    chars: list[tuple[str, bool]] = []  # (字, 下線つきか)
    blanks: list[bool] = []
    for s in p.segs:
        if s.hidden:
            continue
        for ch in s.text:
            blanks.append(bool(s.underline and ch in " \u3000\u2002\u2003" and s.kind == "t"))
            chars.append((ch, bool(s.underline)))
    out = []
    start = None
    for pos, is_blank in enumerate(blanks + [False]):
        if is_blank and start is None:
            start = pos
        elif not is_blank and start is not None:
            if pos - start >= 3:
                out.append((start, pos, True))
            else:
                inside = start > 0 and pos < len(chars) and chars[start - 1][1] and chars[pos][1] \
                    and not chars[start - 1][0].isspace() and not chars[pos][0].isspace()
                if not inside:
                    out.append((start, pos, False))
            start = None
    return out


def _trailing_blank(text: str) -> list[tuple[int, int, str]]:
    """段落の終わりの空白：(始め, 終わり, 理由の鍵)。どちらも字の並びからの推定なので、エージェントの判断（「？」）。
    - 短い見出しの後ろに、全角の空白が4つ以上続いて段落が終わるもの（「住所　　　　　　　」）：trailing_blank
    - それ以外の、2つ以上の空白で段落が終わるもの（見出しが長い・「様」「殿」で終わる・空白が少ない）：trailing_space"""
    m = re.search(r"[ \u3000]{2,}$", text)
    if not m:
        return []
    head = text[:m.start()].strip(" \u3000")
    if not head:
        return []
    start = len(text[:m.start()].rstrip(" \u3000"))  # 「氏　名　  　　　」：見出しの後ろの空白はすべて空欄
    full = re.search(r"[\u3000]{4,}$", text)
    typical = bool(full) and len(head) <= 20 and not head.endswith(("。", "様", "殿", "御中"))
    return [(start, m.end(), "trailing_blank" if typical else "trailing_space")]


def _blank_why(text: str, blank) -> str | None:
    """pdf_blanks（Excel と同じ規則に、PDF の型を足したもの）が見つけた空欄（pdf.layout.Blank）の理由の鍵。
    None は書類の作りで決まるもの（Blank.structural。PDF と同じ見分け方）：下線の字（＿＿＿・___）の並びと、
    かっこの組の中が空白だけのもの（「（　　）」）。英数字・「/」「.」「-」にはさまれた下線の字（URL・ファイルの名前の
    「__icsFiles」）は、名前の一部のこともあるので規則では決めない（underscore_in_word）。
    ほかは、空白の前後の字（単位・「第」・住所の単位など）から推定したもの（字の並びの推定）で、規則では決めない（keyword_blank）。"""
    from ..pdf.layout import underscore_in_word

    a, b = blank
    if getattr(blank, "structural", False):
        return None
    if b > a and all(ch in "_＿" for ch in text[a:b]) and underscore_in_word(text, a, b):
        return "underscore_in_word"
    return "keyword_blank"


def para_spans(p: Para) -> list[tuple[int, int, str | None]]:
    """段落の中の空欄：(始め, 終わり, 理由の鍵)。理由の鍵が None のものは、書類の作りで決まる空欄（規則の欄）：
    下線を引いた空白の並び（3字以上）と「＿＿＿」。ほかは字の並びからの推定で、理由の鍵（WORD_WHY）をつけてエージェントに示す。
    重なる空欄は1つにまとめる（どれかが作りで決まるものなら、まとめた所も規則の欄）。"""
    from ..pdf.layout import ADDRESS_UNITS, pdf_blanks

    text = p.text
    spans: dict[tuple[int, int], str | None] = {}

    def put(a: int, b: int, why: str | None) -> None:
        if (a, b) in spans and why is not None:
            return  # 同じ所：先に見つけた理由（作りで決まるもの・下線のような作りの手がかり）を残す
        spans[(a, b)] = why

    for blank in pdf_blanks(text):
        a, b = blank
        put(a, b, _blank_why(text, blank))
    for a, b, long_ in _underlined_blanks(p):
        put(a, b, None if long_ else "short_underline")
    for a, b, why in _trailing_blank(text):
        put(a, b, why)
    # 住所のマス：「　　郡　市」のように住所の単位で終わる段落は、その後ろに町名・番地を書き足すことがある
    core = text.rstrip(" \u3000")
    if p.cell is not None and core and core[-1] in ADDRESS_UNITS and core == text:
        put(len(text), len(text), "address_end")
    kept: list[list] = []
    for (a, b), why in sorted(spans.items()):
        if kept and a < kept[-1][1]:
            if b > kept[-1][1]:
                kept[-1][1] = b
            if why is None:
                kept[-1][2] = None
            continue
        kept.append([a, b, why])
    return [(a, b, why) for a, b, why in kept]


def para_blanks(p: Para) -> list[tuple[int, int]]:
    """段落の中の空欄の位置（規則の欄も、字の並びから推定した所も）。"""
    return [(a, b) for a, b, _w in para_spans(p)]


class Spots:
    """文書の入力欄の一覧（名前 → 欄の情報）。窓と書き込みで同じものを使う。"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc
        self.items: dict[str, dict] = {}
        self.para_index: dict[int, int] = {}
        self.ranges: dict = {}
        self.marks = doc.marks() if hasattr(doc, "marks") else _bookmarks(doc)
        # todoku が書いたマス（しおりのあるマス）：ほかの欄の見出しにしない
        self.written_cells = {id(q.cell) for name, (q, _x, _y) in self.marks.items() if q.cell is not None and name.startswith("T")}
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
                        label, rl, cl = _cell_label(t, c, self.written_cells)
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
                bullet_id = f"P{pi}.b"
                bm = self.marks.get(bullet_id)
                if getattr(p, "bullet", None) is not None or (bm is not None and bm[0] is p):
                    # 箇条書きの印の □（段落番号の定義の字）：チェックの欄。書くと、この段落の印だけを ■・☑ にする
                    m = re.search(r"[\u3000 ]{2,}|□|☐|※", text)
                    opt = norm_label((text[:m.start()] if m and m.start() > 0 else text).strip(" \u3000"))[:40]
                    self.items[bullet_id] = {"id": bullet_id, "kind": "checkbox", "source": "layout", "para": p,
                                             "bullet": getattr(p, "bullet", None), "label": opt, "provenance": LAYOUT,
                                             "filled": bm is not None}
                k = 0
                spans = []  # (始め, 終わり, 種類, 理由の鍵（None は規則の欄）)
                for i, ch in enumerate(text):
                    if ch in BOX and not _in_field(p, i):
                        spans.append((i, i + 1, "checkbox", None))
                    elif ch in CHECKED and not _in_field(p, i):
                        spans.append((i, i + 1, "checkbox", "checked_mark"))  # チェック済みの欄かもしれない（■ は箇条の印のことも）
                for a, b, why in para_spans(p):
                    if any(_in_field(p, i) for i in range(a, b)):
                        continue
                    if why in ("trailing_space", "short_underline") and _touches_field(p, a, b):
                        continue  # フォームの欄にすぐ接する空白は、欄の字の飾り（欄はその id で書く）
                    spans.append((a, b, "inline", why))
                # 書き込み済みの欄（しおり）は、その番号を持ったまま（4 で作る）。残りの空欄に、しおりの番号を除いた番号を
                # 位置の順に割り当てる（元の白紙のときと同じ id になる：P16.1 に書いた後も、P16.2・P16.3 はそのまま）。
                # 字の並びから推定した所（理由の鍵のある所）も、同じ番号の並びで P段落.番号 とし、窓に「？」をつけて示す
                mark_k: dict[int, tuple[int, int]] = {}
                mark_ka: dict[int, tuple[int, int]] = {}
                for name, (q, x, y) in self.marks.items():
                    mm = re.fullmatch(rf"P{pi}\.(a?)(\d+)", name) if q is p else None
                    if mm:
                        (mark_ka if mm[1] else mark_k)[int(mm[2])] = (x, y)
                written = [(x, y) for q, x, y in self.marks.values() if q is p]
                spans = [t for t in spans if not _overlaps(t[0], t[1], written)]
                free_k = (n for n in itertools.count(1) if n not in mark_k)
                for a, b, kind, why in sorted(spans, key=lambda t: t[:3]):
                    k = next(free_k)
                    fid = f"P{pi}.{k}"
                    if kind == "checkbox":
                        near = text[b:b + LABEL_WINDOW]  # とても長い段落でも、印の近くだけを見る
                        m = re.search(r"[\u3000 ]{2,}|[□☐☑☒■✓✔]|※", near)
                        opt = near[:m.start()] if m else near
                        label = (norm_label(opt.strip(" \u3000"))
                                 or norm_label(text[max(0, a - LABEL_WINDOW):a].strip(" \u3000"))[-30:])
                    else:
                        label = self._label(p, a, b)
                    it = {"id": fid, "kind": kind, "source": "layout", "para": p, "span": (a, b), "label": label, "provenance": LAYOUT}
                    if why is not None:
                        it.update(provenance=AGENT, why=WORD_WHY[why], why_key=why)
                        if why == "checked_mark":
                            it["checked"] = True
                    self.items[fid] = it
                # 規則では決められない所：窓に「？」をつけて示し、エージェントが読んで判断する（id は P段落.a番号）
                if any(s_.field is not None for s_ in p.segs):
                    continue
                if not text.strip(" \u3000") == "" and not "".join(
                        s_.text for s_ in p.segs if not s_.hidden and s_.kind in ("t", "sym", "tab")).strip(" \u3000\t"):
                    continue  # 図・埋め込みの印だけの段落（窓のために作った字だけ）
                free_ka = (n for n in itertools.count(1) if n not in mark_ka)
                found_agent = agent_blanks(text, [(x, y) for x, y, _k, _w in spans] + written, p.cell is not None)
                # 見出しのあるマスの中の空の段落（「〒」の下の行に住所を書くなど）。マスのほかの字が長いもの（説明の文のこともある）も、
                # 理由を分けて示す（字数では決めない）
                c = p.cell
                if c is not None and not text.strip(" \u3000") and not spans and not written:
                    others = [q.text.strip(" \u3000") for q in c.paras() if q is not p and q.text.strip(" \u3000")]
                    before = [q for q in c.paras()[:c.paras().index(p)] if q.text.strip(" \u3000")] if p in c.paras() else []
                    if others and before:
                        found_agent = [(0, len(text), "next_line" if all(len(o) <= 16 for o in others) else "next_line_long")]
                for a, b, why in found_agent:
                    if any(_in_field(p, i) for i in range(a, max(a + 1, b))):
                        continue
                    fid = f"P{pi}.a{next(free_ka)}"
                    it = {"id": fid, "kind": "inline", "source": "layout", "para": p, "span": (a, b),
                          "label": self._label(p, a, b), "provenance": AGENT, "why": WORD_WHY[why], "why_key": why}
                    if why == "placeholder":
                        it["placeholder"] = text[a:b]
                    alt = self._beside_cell_field(p, why)
                    if alt:
                        it["alternative_to"] = alt
                    self.items[fid] = it
        # 3') マスの中の見出し（「住所」「住所：」「所在地　〒」）の後ろ：すぐ右のマスに欄があれば、値はそちらに書くことが多い。
        # 消さずに残し、どちらに書くかはエージェントが窓を読んで決める（alternative_to：代わりの欄。規則の欄を先に）
        in_cell: dict[int, list[str]] = {}
        for fid_, it in self.items.items():
            c_ = it.get("cell") or (it["para"].cell if it.get("para") is not None else None)
            if c_ is not None:
                in_cell.setdefault(id(c_), []).append(fid_)
        for fid in [f for f, it in self.items.items() if it.get("why_key") in ("after_label", "after_colon", "after_mark")]:
            it = self.items[fid]
            c_ = it["para"].cell
            if it.get("alternative_to") or c_ is None or c_.table is None:
                continue
            right = [x for x in c_.table.rows[c_.row] if x.col == c_.col + c_.span]
            there = in_cell.get(id(right[0]), []) if right else []
            ruled = [x for x in there if self.items[x].get("provenance") != AGENT]
            if ruled or there:
                it["alternative_to"] = (ruled or there)[0]
        # 4) 書き込み済みの欄（しおり）：同じ名前の欄が見つからなければ、しおりから作る
        for name, (p, a, b) in self.marks.items():
            if name in self.items:
                continue
            kind = "checkbox" if p.text[a:b] in tuple(BOX | CHECKED) else ("cell" if name.startswith("T") else "inline")
            label = self._label(p, a, b) if kind == "inline" else ""
            if kind == "cell" and p.cell is not None and p.cell.table is not None:
                label = _cell_label(p.cell.table, p.cell, self.written_cells)[0]  # 書き込み済みのマスも、表の行・列の見出しで示す
            self.items[name] = {"id": name, "kind": kind, "source": "layout", "para": p, "span": (a, b), "filled": True,
                                "label": label, "provenance": LAYOUT}

    def _beside_cell_field(self, p: Para, why: str) -> str | None:
        """マスの中の「住所：」のすぐ右のマスが空のマスの欄（そこに書くことが多い）か、「人」だけのマスの下に空のマスの欄がある
        （列の見出し）：その規則の欄の id。どちらも、候補は消さずに alternative_to（代わりの欄）として示す。"""
        c = p.cell
        if c is None or why not in ("after_colon", "before_unit", "after_label", "after_mark"):
            return None
        for it in self.items.values():
            x = it.get("cell")
            if x is None or it.get("kind") != "cell" or x.table is not c.table:
                continue
            if why in ("after_colon", "after_label", "after_mark") and x.row == c.row and x.col == c.col + c.span \
                    and self.items.get(it["id"], {}).get("provenance") != AGENT:
                return it["id"]
            if why == "before_unit" and x.row > c.row and x.col <= c.col < x.col + x.span:
                return it["id"]
        return None

    def _label(self, p: Para, a: int, b: int) -> str:
        """段落の中の空欄の見出し：その空欄を［　］にした段落の文字。書き込み済みのほかの欄の値は除く（白紙のときと同じ見出し）。
        とても長い段落（改行のない数 MB の .txt など）は、空欄の前後 LABEL_WINDOW 字だけを見出しにする
        （空欄ごとに段落の全部を作り直すと、空欄の数 × 段落の長さの時間がかかる）。"""
        import bisect

        per = self.__dict__.get("_cuts")
        if per is None:  # 書き込み済みの欄の位置：段落ごとに 1 度だけ集めて並べる
            per = self._cuts = {}
            for q, x, y in self.marks.values():
                if y > x:
                    per.setdefault(id(q), []).append((x, y))
            for v in per.values():
                v.sort()
        lo, hi = max(0, a - LABEL_WINDOW), min(len(p.text), b + LABEL_WINDOW)
        mine = per.get(id(p), [])
        near = mine[bisect.bisect_left(mine, (lo, -1)):bisect.bisect_right(mine, (hi, float("inf")))]
        cuts = [(x, y) for x, y in near if not (x == a and y == b) and y <= hi]
        out, at = [], lo
        for x, y in sorted(cuts + [(a, b)]):
            if x < at:
                continue
            out.append(p.text[at:x])
            out.append("［　］" if (x, y) == (a, b) else "　")
            at = y
        out.append(p.text[at:hi])
        return norm_label("".join(out))

    def value(self, item: dict):
        """いまの値（書き込み済みなら、その値）。"""
        name = item["id"]
        if is_bullet(item):
            return item["para"].label in BULLET_CHECKED  # 箇条書きの印
        if name in self.marks:
            p, a, b = self.marks[name]
            t = p.text[a:b]
            if item["kind"] == "checkbox":
                return t in CHECKED
            return t
        c: Control | None = item.get("control")
        if c is not None:
            if c.source == "activex":
                from ..axfile import value_of

                return value_of(c.ax)  # ActiveX の部品：部品の中に保存された値
            if c.kind == "checkbox":
                return c.checked
            if c.placeholder:
                return None
            t = "".join(s.text for s in c.result_segs)
            if c.kind == "dropdown" and c.source == "formfield":
                return c.value
            return t.strip("\u2002 ") if c.source == "formfield" else t
        if item["kind"] == "checkbox":
            return bool(item.get("checked"))  # チェック済みの印らしい字（☑・■ など）の候補は、チェック済みとして読む
        return None


LABEL_WINDOW = 2000  # 空欄の見出しに使う、空欄の前後の字の数（ふつうの段落はまるごと入る長さ）


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
    near_fields = [f for j in range(max(0, a - (a == b)), max(a + 1, b)) for f in [_field_at(q, j)] if f is not None]
    calc = next((f for f in near_fields if f.kind == "field"), None)
    if calc is not None and any(_field_at(q, j) is calc for j in range(a, max(a + 1, b))):
        what = getattr(calc, "instr", "") or calc.name or "フィールド"
        raise WriteError(f"「{old}」は、フィールド（{what}）の結果の字です。Word などで開いてフィールドを更新すると、"
                         "書いた字は消えて元の結果に戻るので、書けません（差し込み印刷の欄なら、差し込みのデータの値になります）。"
                         "様式の作り手に差し込みを解除した様式を求めるか、本人に、印刷した後に手で書いてもらってください")
    if any(f.kind != "field" for f in near_fields):
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


def _touches_field(p: Para, a: int, b: int) -> bool:
    """[a, b) が、フォームの欄の字（長さ 0 の欄も）に接するか重なるか。"""
    k = 0
    for s in p.segs:
        if s.hidden:
            continue
        if s.field is not None and s.field.kind != "field" and k <= b and a <= k + len(s.text):
            return True
        k += len(s.text)
    return False


def _field_at(p: Para, i: int):
    """段落の i 字目の字の、入力欄・フィールドの印（なければ None）。"""
    k = 0
    for s in p.segs:
        if s.hidden:
            continue
        if k <= i < k + len(s.text):
            return s.field
        k += len(s.text)
    return None


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


def _cell_label(t: Table, c: Cell, written: set | None = None) -> tuple[str, str, str]:
    """表の空のマスの見出し：同じ行の左のマス（行の見出し）と、上のマス（列の見出し）。列の見出しは、そのマスの列の幅を
    すべて覆うマスだけ（結合して2列にまたがるマスの上の、1列だけのマスは、その列の見出しではない）。
    todoku が書いたマス（written：書き込みのしおりのあるマス）は、見出しにしない（書いた値は、印刷された見出しではない）。"""
    written = written or set()
    row = t.rows[c.row]
    left = None
    for x in row:
        if x.col < c.col and _printed(x) and id(x) not in written:
            if left is None or x.col > left.col:
                left = x
    up = None
    for r in range(c.row - 1, -1, -1):
        for x in t.rows[r]:
            if x.col <= c.col < x.col + x.span and (x.text() or id(x) in written):
                up = x
                break
        if up is not None:
            break
    if up is not None and (id(up) in written or not _printed(up) or c.col + c.span > up.col + up.span):
        up = None
    rl = norm_label(_printed(left)) if left else ""
    cl = norm_label(_printed(up)) if up else ""
    adjacent_left = left is not None and left.col + left.span == c.col
    if rl and cl and rl != cl and not (adjacent_left and up is not None and c.row - up.row > 2):
        return f"{rl}／{cl}", rl, cl
    return rl or cl, rl, cl


def _printed(c: Cell) -> str:
    """マスの、印刷される字（入力欄の値・フィールドの結果でない字）。"""
    out = []
    for p in c.paras():
        t = "".join(s.text for s in p.segs if not s.hidden and s.field is None).strip()
        if t:
            out.append(t)
    return " ".join(out)


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
                     "［P12.1？］": "段落の中の、字の並び（年・月・日などの字）から推定した空欄。規則では決めないので、エージェントが判断する",
                     "■［P12.1？］": "チェック済みの印らしい字（☑・■・✓ など。■ は箇条の印のこともある）。チェックの欄なら True / False で書く",
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
    # 変更履歴で削除された字の残る段落（窓では〔変更履歴で削除された字〕と示す）
    out["tracked_deletions"] = [f"段落 {sp.para_index[id(q)]}「{q.deleted_text().strip()[:20]}」"
                                for _k, _part, blocks in doc.all_blocks() for q in iter_paras(blocks)
                                if hasattr(q, "deleted_text") and q.deleted_text().strip()]
    out["embeds"] = _embeds(doc)
    out["pictures"] = _pictures(doc)
    from ..embeds.word import objects as _objects

    out["objects"] = _objects(doc, shown={e["n"] for e in out["embeds"]})  # グラフ・SmartArt・Excel の表のほかの埋め込み（読むだけ）
    contexts = {}
    for st in out["stories"]:
        for line in st["reading"]:
            for m in re.finditer(r"［([^］＝？]+)？?(?:＝[^］]*)?］", line):
                contexts.setdefault(m[1], line)
    for it in sp.items.values():
        c = it.get("control")
        kind = it["kind"]
        label = it.get("label") or ""
        printed = False
        if c is not None:
            # 見出しは、文書に印刷される字（欄の前の字・同じ行の左のマス）。なければ、作り手の説明（help）・欄の名前（タグ）
            label = printed_label(sp, it)
            printed = bool(label)
            if c.source == "activex":
                label = label or c.help or c.name  # 見出しは前に書かれた字か、部品の Caption
            else:
                label = label or c.help or c.name
        d = {"id": it["id"], "kind": "text" if kind == "cell" else kind, "label": label, "source": it["source"],
             "provenance": it["provenance"], "value": sp.value(it), "context": contexts.get(it["id"], "")}
        if printed or (c is None and label):
            d["label_printed"] = True  # 見出しは文書に印刷される字（欄の名前・タグではない）
        if c is not None and c.source == "activex":
            from ..axfile import word_describe

            d.update(word_describe(doc, c))
            d["control"] = c.name
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
            if it.get("alternative_to"):
                d["alternative_to"] = it["alternative_to"]  # 同じ値を書く、代わりの欄（どちらか一方に書く）
        if it.get("filled") or (it["id"] in sp.marks and c is None):
            d["filled"] = True  # todoku が書き込んだ所（書き直すには元の様式に書く）
        if it.get("row_label"):
            d["row_label"] = it["row_label"]
        if it.get("column_label"):
            d["column_label"] = it["column_label"]
        if it.get("shade") and it["shade"] not in ("auto", "FFFFFF"):
            d["shaded"] = it["shade"]
        if c is not None and c.lock in ("contentLocked", "sdtContentLocked"):
            d["readonly"] = True
        if c is not None and getattr(c, "unwritable", ""):
            d["readonly"] = True
            d["unwritable"] = True
            d["why"] = c.unwritable  # 書けない理由（欄の状態を入れる所がファイルにないなど）
        if not editable(doc, sp, it):
            d["readonly"] = True
            d["protected"] = True  # 文書の保護で、編集できる範囲の外
        out["fields"].append(d)
    return out


def _pictures(doc) -> list[dict]:
    """文書の絵：説明（代わりの字）があるか、飾りの印があるか。説明のない絵は、本文に〔図N（説明なし）〕と出す。
    where は、その絵のある段落の字（本人に場所を伝えるため）。"""
    out = []
    for pic in getattr(doc, "pictures", []) or []:
        near = re.sub(r"〔[^〕]*〕", "", pic["para"].text).strip()
        out.append({"id": f"図{pic['n']}", "alt": pic["alt"] or None, "decorative": pic["decorative"],
                    "where": f"「{near[:30]}」の段落" if near else "字のない段落"})
    return out


# 埋め込みのブックの Excel の窓：中身（入れ物のバイト列）が同じなら同じ窓になるので、作り直さない（検査では、書き込みの前後や
# 目印の書き込みの後に、同じ文書の窓を何度も作る。埋め込みの大きなブックの窓を毎回作ると、遅すぎる）
_EMBED_WINDOWS: dict[tuple, object] = {}
_EMBED_WINDOWS_MAX = 32


def _embed_key(doc, e: dict) -> tuple | None:
    import hashlib

    h = hashlib.sha256()
    try:
        if getattr(doc, "kind", "") == "doc":
            from ..cfb import SubStorage

            for path, data in sorted(SubStorage(doc.cfb, e["storage"]).contents().items()):
                h.update(path.encode("utf-8") + b"\0" + len(data).to_bytes(8, "little"))
                h.update(data)
        else:
            h.update(doc.zip.read(e["storage"]))
    except Exception:  # noqa: BLE001
        return None
    return (getattr(doc, "kind", "docx"), e["storage"], h.hexdigest())


def _embed_windows(doc) -> list[tuple[dict, object]]:
    """(埋め込み, (Excel の窓, 窓の Markdown の行) か 作れなかった理由の字)。Excel のブックでない埋め込みは除く。"""
    from ..xlsx import window as xw

    if not hasattr(doc, "embedded_book"):
        return []
    out = []
    for e in list(getattr(doc, "embeds", []) or []):
        k = _embed_key(doc, e)
        if k is not None and k in _EMBED_WINDOWS:
            got = _EMBED_WINDOWS[k]
        else:
            wb = doc.embedded_book(e)
            if wb is None:
                got = None  # Excel のブックではない（開けない）
            else:
                try:
                    xdoc = xw.build(wb)
                    got = (xdoc, [ln for sh in xdoc["sheets"] for ln in xw._md_sheet(sh)])
                except Exception as ex:  # noqa: BLE001
                    got = str(ex)
            if k is not None:  # 中身の要約を作れなかったものは、使い回さない
                while len(_EMBED_WINDOWS) >= _EMBED_WINDOWS_MAX:
                    _EMBED_WINDOWS.pop(next(iter(_EMBED_WINDOWS)))
                _EMBED_WINDOWS[k] = got
        if got is not None:
            out.append((e, got))
    return out


def _embeds(doc) -> list[dict]:
    """文書に埋め込んだ Excel の表：Excel の窓と同じ読み方で、シートごとに平文にする（本文では〔埋め込み1：…〕の所）。"""
    from ..embed import window_info

    out = []
    for e, got in _embed_windows(doc):
        if isinstance(got, str):
            out.append({"n": e["n"], "label": e["label"], "prog": e["prog"], "storage": e["storage"], "error": got})
            continue
        xdoc, md = got
        info = window_info(doc, e, xdoc)  # 書くときの欄の名前（E1!Sheet1!C4）と、表の絵を描き直せるか
        out.append({"n": e["n"], "label": e["label"], "prog": e["prog"], "storage": e["storage"],
                    "sheets": [sh["name"] for sh in xdoc["sheets"]], "markdown": list(md), **info})
    return out


def printed_label(sp: "Spots", it: dict) -> str:
    """入力欄（フォームの欄・コンテンツコントロール）の見出し：同じ段落の、欄の前に印刷される字（ほかの欄の値・todoku が
    書いた所は除く。区切り（タブ・2つ以上の空白・表のマスの区切り）より前は含めない）。段落の頭にある欄で、表のマスの中なら、
    同じ行の左のマス（か上のマス）の字。どれもなければ ""。"""
    c: Control = it["control"]
    p = c.para
    if p is None:
        return ""
    cuts = [(x, y) for q, x, y in sp.marks.values() if q is p and y > x]
    a, _b = _ctl_span(p, c)
    out, pos = [], 0
    for s in p.segs:
        if s.hidden:
            continue
        if pos >= a:
            break
        n = len(s.text)
        if s.field is not None and s.field is not c or s.kind in ("note",):
            out.append("\x00")  # ほかの欄の値・脚注の印：見出しに入れない
        else:
            for k, ch in enumerate(s.text[:max(0, a - pos)]):
                out.append("\x00" if any(x <= pos + k < y for x, y in cuts) else ch)
        pos += n
    head = re.split(r"\x00|\t|[ \u3000]{2,}|\n", "".join(out))[-1]
    head = norm_label(head.strip(" \u3000")).rstrip("\uff1a:\u3000 ")
    if len(head) > 30:
        head = head[-30:]
    if head:
        return head
    cell = p.cell
    if cell is not None and cell.table is not None:
        return _cell_label(cell.table, cell, sp.written_cells)[0]
    return ""


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
                edits.append((a, b, ("☑" if v in CHECKED else "☐") + f"［{name}］"))
            else:
                edits.append((a, b, f"［{name}＝{v}］"))
            continue
        if "span" in it:
            a, b = it["span"]
            if it["kind"] == "checkbox" and it.get("provenance") == AGENT:
                edits.append((a, b, f"{text[a:b]}［{name}？］"))  # チェック済みの印らしい字：元の字のまま（■ は箇条の印のことも）
            elif it["kind"] == "checkbox":
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
            if c.source == "activex":
                from ..axfile import word_mark

                edits.append((a, b, word_mark(c, name)))  # ActiveX の部品
                continue
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
    gone = p.deleted_text().strip() if hasattr(p, "deleted_text") else ""
    if gone:
        t += f"〔変更履歴で削除された字（いまの本文ではない）: {gone}〕"
    label = p.label
    bullet = next((it for it in items if is_bullet(it)), None)
    if bullet is not None and label:
        label = ("☑" if label in BULLET_CHECKED else "☐") + f"［{bullet['id']}］"
    if label:
        t = f"{label} {t}" if not label.endswith((" ", "\t", "　")) else label + t
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
    # 同じ窓を使うほかの形式（OpenDocument の文書・RTF）は、kind_ja に形式の名前を入れる
    out.append(f"- 元ファイルの SHA-256: `{doc['source']['sha256']}`（{doc['source']['bytes']:,} バイト）、{doc.get('kind_ja') or 'Word'}")
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
                k = f"規則では決められない：{f['why']}" + (f"（{f['alternative_to']} の代わりの所かもしれません。書くのはどちらか一方）" if f.get("alternative_to") else "")
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
        out += ["- 本文の〔埋め込み" + str(e["n"]) + "：…〕の所にある表です。Excel の窓と同じ書き方（［C4］はセル）で出しています。"]
        ids = [f["id"] for f in e.get("fields") or []]
        if e.get("writable"):
            sheet = (e.get("fields") or [{}])[0].get("sheet") or (e.get("sheets") or ["Sheet1"])[0]
            out.append(f"- **書くときの欄の名前**は、`E{e['n']}!シート!セル` です（例：下の［C4］は `E{e['n']}!{sheet}!C4`）。"
                       "書くと、Word が表示・印刷に使う表の絵も、字が変わったセル（計算結果を含む）の所だけ描き直します。")
            if ids:
                out.append(f"- 入力欄（{len(ids)} か所）: " + "、".join(ids[:200]) + (f" ほか {len(ids) - 200} か所" if len(ids) > 200 else ""))
        else:
            out.append(f"- この表には、まだ書き込めません：{e.get('why') or '理由不明'}")
        out.append("")
        out += [ln.replace("## シート", "### シート", 1) if ln.startswith("## シート") else
                (ln.replace("### ", "#### ", 1) if ln.startswith("### ") else ln) for ln in e["markdown"]]
        out.append("")
    from ..embeds.word import markdown as objects_md

    out += objects_md(doc.get("objects"))
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
