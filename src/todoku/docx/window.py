"""Word の窓：本文を上から（表は行ごとに、マスを ｜ で区切って）、入力欄は［］で示す。

入力欄：
- 旧来のフォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN）とコンテンツコントロール（名前は欄の名前・タグ）
- 表の空のマス（左か上のマスの文字を見出しにする）：T{表}.R{行}.C{列}
- 段落の中の空欄（下線を引いた空白、下線の字「＿＿＿」、かっこの中の空白「（　）」「[ ]」）とチェックの箱の字（☐ □ など）：P{段落}.c{字数}
- 字の形と並びだけで見つけた所（1 字ぶん以上の空白の並び（wide_space）・点線・ダッシュの並び）と、チェック済みの印らしい字（☑・■・✓）、
  規則の型に合わない、書き込む所かもしれない所は、同じ名前の決まりで、規則では決めない所（「？」）として示す
  （字の意味は見ない。エージェントが前後の字から判断する）。すぐ隣の欄の代わりの所は alternative_to で示す
- 名前の字数：段落の頭からその所までにある、空白でも所でもない字の数（_name_places。数え順の番号ではないので、
  ほかの所が増えても減っても、前の所に書いた後も変わらない）。古い数え順の名前（P12.1・P12.a1）は、決まるときだけ受ける（Items）
- 表のマスの見出しは、左・上のマスの字をそのまま（row_label・column_label）と、間のマスの数（row_label_gap・column_label_gap）で示す
書き込んだ欄は、見えないしおり（_todoku_…）で囲み、書いた後も同じ名前で読めるようにする。
埋め込んだ Excel の表は、Excel の窓と同じ読み方で節に出し、書くときの欄の名前（E1!Sheet1!C4）と、
Word が表示・印刷に使う表の絵を描き直せるか（描き直せなければ、その理由）を添える。
"""

from __future__ import annotations

import bisect
import hashlib
import itertools
import re

from .. import chars
from ..xlsx.window import AGENT_WHY, BLANK_RULE_KINDS, _ident, agent_blanks, blank_places, label_tail, label_text, shown_gap, wide_space
from ..xlsx.window import GAP_TOKEN_NOTE, ONE_WIDE_SPACE_WHY, one_wide_space, place_facts
from ..xlsx.window import gap_interior, with_label_keys
from .document import Cell, Control, DocxDocument, Para, Table, iter_paras

DET = "deterministic"
LAYOUT = "layout"
AGENT = "agent"  # 規則では決められない、書き込む所かもしれない所（エージェントが読んで判断する）
BOX = chars.EMPTY_BOXES  # 印のない箱の字（chars.py）。どれも箇条・見出しの印のこともある字なので、規則の欄にせず「？」（理由 bullet_box）で示す
# 印のついたチェックの印らしい字：☑・☒（Ballot Box）はチェック済みの箱、■ は箇条の印・飾りのこともあり、✓ は本文の印のこともある。
# 規則では決めず、チェック済み（checked）のチェックボックスの候補として「？」をつけて示す（エージェントが判断する）
CHECKED = chars.CHECKED_BOXES | chars.MAYBE_CHECKED
UNCHECK = {"☑": "☐", "☒": "☐", "■": "□", "✓": "□", "✔": "□", "▣": "□"}  # チェックを外すときに書く字（印のない同じ形の箱）
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
    "short_underline": "下線を引いた短い空白（1〜2字。下線の空欄は3字以上のことが多い）",
    "checked_mark": "印のついたチェックの印らしい字（☑・☒・■・✓。いまはチェック済みとして読む。■ は箇条の印・飾りのこともある）",
    "underscore_in_word": "英数字にはさまれた下線の字（「＿＿」。URL・ファイルの名前の一部のこともある）",
    "next_line_long": "見出しのあるマスの中の、空の段落（マスのほかの字が長く、見出しではなく説明の文のこともある）",
    "first_line": "字のあるマスの、初めの空の段落（後ろの段落にだけ字がある。その字は見出しのことも、別の欄の見出しのこともある）",
}


def box_mark(orig: str, on: bool) -> str:
    """チェックボックスの字 orig に、チェックを入れる（on）・外すときに書く字（変えないときは orig のまま）。
    ☐ は ☑、ほかの白い箱（□ ❏ ⬜ など）は ■ にする。チェック済みの印（☑・■・✓ など）は、外すときに印のない箱（☐・□）にする。"""
    if orig in CHECKED:
        return orig if on else UNCHECK[orig]
    return ("☑" if orig == "☐" else "■") if on else orig


def _underlined_blanks(p: Para) -> list[tuple[int, int, bool]]:
    """下線を引いた空白の並び（「氏名＿＿＿」を下線つきの空白で作ったもの）：(始め, 終わり, 3字以上か)。
    3字以上は書類の作り（下線の空欄）として規則の欄。1〜2字は、下線つきの見出しの字間のこともあるので、規則では決めない（エージェント）。
    下線つきの字にはさまれた空白（下線を引いた字の中の、語の間の空白）は、空欄の形ではないので出さない。"""
    chars_: list[tuple[str, bool]] = []  # (字, 下線つきか)
    blanks: list[bool] = []
    for s in p.segs:
        if s.hidden:
            continue
        if not s.underline or s.kind != "t":
            blanks.extend([False] * len(s.text))
            chars_.extend((ch, bool(s.underline)) for ch in s.text)
            continue
        for ch in s.text:
            blanks.append(chars.is_space(ch))
            chars_.append((ch, True))
    out = []
    start = None
    for pos, is_blank in enumerate(blanks + [False]):
        if is_blank and start is None:
            start = pos
        elif not is_blank and start is not None:
            if pos - start >= 3:
                out.append((start, pos, True))
            else:
                inside = start > 0 and pos < len(chars_) and chars_[start - 1][1] and chars_[pos][1] \
                    and not chars_[start - 1][0].isspace() and not chars_[pos][0].isspace()
                if not inside:
                    out.append((start, pos, False))
            start = None
    return out


def para_spans(p: Para, text: str | None = None) -> list[tuple[int, int, str | None]]:
    """段落の中の空欄：(始め, 終わり, 理由の鍵)。理由の鍵が None のものは、書類の作りで決まる空欄（規則の欄）：
    下線を引いた空白の並び（3字以上）・下線の字の並び・かっこの組の中が空白だけの所。ほかは字の形と並びだけで見つけた所
    （1 字ぶん以上の空白の並び（wide_space）・点線・ダッシュの並び。xlsx.window.blank_places）で、理由の鍵（WORD_WHY）をつけてエージェントに示す。
    英数字にはさまれた下線の字（URL・ファイルの名前の「__icsFiles」）は、名前の一部のこともあるので規則では決めない。
    重なる空欄は1つにまとめる（どれかが作りで決まるものなら、まとめた所も規則の欄）。text は p.text（呼ぶ側で作ってあれば渡す）。"""
    if text is None:
        text = p.text
    spans: dict[tuple[int, int], str | None] = {}

    def put(a: int, b: int, why: str | None) -> None:
        if (a, b) in spans and why is not None:
            return  # 同じ所：先に見つけた理由（作りで決まるもの・下線のような作りの手がかり）を残す
        spans[(a, b)] = why

    for a, b, long_ in _underlined_blanks(p):  # 下線（書式）は、字の並びより先に見る（同じ所なら、下線の理由を残す）
        put(a, b, None if long_ else "short_underline")
    for a, b, kind in blank_places(text):
        why = None if kind in BLANK_RULE_KINDS else kind
        if kind == "underscore" and 0 < a and b < len(text) and _ident(text[a - 1]) and _ident(text[b]):
            why = "underscore_in_word"
        put(a, b, why)
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


def _field_mask(p: Para) -> bytearray:
    """段落の字ごとに、フォームの欄の字か（1）。段落ごとに 1 回だけ作る（字ごとに区切りをたどり直さない）。"""
    out = bytearray()
    for s in p.segs:
        if s.hidden:
            continue
        out.extend((b"\x01" if s.field is not None else b"\x00") * len(s.text))
    return out


def _option_end(text: str, b: int, stops: list[int]) -> int:
    """チェックの箱の字の後ろの項目の字の終わり：次の箱の字か、1 字ぶん以上の空白の並びの始めまで（stops は段落ごとに作った位置）。"""
    i = bisect.bisect_left(stops, b)
    return stops[i] if i < len(stops) else len(text)


def _option_stops(text: str) -> list[int]:
    """項目の字の区切りになる位置（チェックの箱の字・印のついた箱の字・1 字ぶん以上の空白の並びの始め：wide_space）。段落ごとに 1 回だけ作る。"""
    out = [i for i, ch in enumerate(text) if ch in BOX or ch in CHECKED]
    out += [a for a, b, k in chars.char_runs(text) if k == "space" and wide_space(text[a:b])]
    return sorted(out)


class Items(dict):
    """欄の名前 → 欄の情報。get は、古い数え順の名前（P12.1・P12.a1）も、どの数え方でも同じ所を指すとき（その段落に
    その並びの所が 1 つだけのとき）だけ受ける（legacy：古い名前 → いまの名前）。いくつもあって決まらない古い名前は受けず、
    hint で、いまの名前を示す。"""

    def __init__(self):
        super().__init__()
        self.legacy: dict[str, str] = {}
        self.ambiguous: set[str] = set()

    def get(self, key, default=None):
        if dict.__contains__(self, key):
            return dict.__getitem__(self, key)
        new = self.legacy.get(key) if isinstance(key, str) else None
        return dict.__getitem__(self, new) if new is not None and dict.__contains__(self, new) else default

    def hint(self, key) -> str:
        """欄が見つからないときに添える字：古い名前なら、いまの名前（同じ段落の所）。"""
        m = re.fullmatch(r"(P\d+)\.a?\d+", str(key))
        if not m:
            return ""
        now = [k for k in self if k.startswith(m[1] + ".c")]
        if not now:
            return ""
        why = "いくつもあって決まらない古い番号です" if key in self.ambiguous else "古い数え方の番号です"
        return (f"（{why}。段落の中の所の名前は、その所の前にある字の数で決まります（P段落.c字数）。"
                f"この段落の所：{'・'.join(now[:20])}）")


class Spots:
    """文書の入力欄の一覧（名前 → 欄の情報）。窓と書き込みで同じものを使う。"""

    def __init__(self, doc: DocxDocument):
        self.doc = doc
        self.items: Items = Items()
        self.para_index: dict[int, int] = {}
        self.ranges: dict = {}
        self._texts: dict[int, str] = {}  # 段落の字（id(段落) → 字）。見出しを作るたびに Para.text を作り直さない
        self.marks =doc.marks() if hasattr(doc, "marks") else _bookmarks(doc)
        # todoku が書いたマス（しおりのあるマス）：ほかの欄の見出しにしない
        self.written_cells = {id(q.cell) for name, (q, _x, _y) in self.marks.items() if q.cell is not None and name.startswith("T")}
        self._collect()

    def _collect(self) -> None:
        doc = self.doc
        n = getattr(doc, "para_base", 0)  # 書類の一部だけを読むとき（lazy）：段落の番号は元の書類の番号
        for _kind, _part, blocks in doc.all_blocks():
            for p in iter_paras(blocks):
                n += 1
                self.para_index[id(p)] = n
        # 1) フォームの欄・コンテンツコントロール
        used_names = set(getattr(doc, "control_names", ()))
        fixed = getattr(doc, "control_ids", None) or {}  # 書類の一部だけを読むとき（lazy）：全体の順で決めた名前
        for k, c in enumerate(doc.controls, len(used_names) + 1):
            name = fixed.get(id(c)) or c.name or f"ctl{k}"
            base = name
            j = 2
            while name in used_names:
                name = f"{base}#{j}"
                j += 1
            used_names.add(name)
            self.items[name] = {"id": name, "kind": c.kind, "source": c.source, "control": c, "para": c.para,
                                "options": c.options or None, "provenance": DET}
        # 2) 表の空のマス
        tno = getattr(doc, "table_base", 0)  # 書類の一部だけを読むとき（lazy）：表の番号は元の書類の番号
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
                        nb = cell_neighbours(t, c, self.written_cells)  # 左・上のマスの字（そのまま）と、間のマスの数
                        fid = f"T{tno}.R{c.row + 1}.C{c.col + 1}"
                        if not label:
                            # 見出しの見つからない空のマス：桁のマスのこともあれば、飾りのマスのこともある（エージェントが判断）
                            self.items[fid] = {"id": fid, "kind": "cell", "source": "layout", "cell": c, "para": c.paras()[0],
                                               "label": "", "provenance": AGENT, "why": AGENT_WHY["empty_cell"],
                                               "why_key": "empty_cell", "shade": c.shade, "neighbours": nb}
                            continue
                        self.items[fid] = {"id": fid, "kind": "cell", "source": "layout", "cell": c, "para": c.paras()[0],
                                           "label": label, "provenance": LAYOUT, "shade": c.shade, "neighbours": nb}
        # 3) 段落の中の空欄と □（空のマスとして欄にした段落は除く）
        cell_paras = {id(q) for it in self.items.values() if it.get("kind") == "cell" for q in it["cell"].paras()}
        marks_of: dict[int, list[tuple[str, int, int]]] = {}  # 段落ごとの、書き込み済みの欄（しおり）。1 度だけ集める
        for name, (q, x, y) in self.marks.items():
            marks_of.setdefault(id(q), []).append((name, x, y))
        for _kind, part, blocks in doc.all_blocks():
            for p in iter_paras(blocks):
                if id(p) in cell_paras:
                    continue
                if any(s.field is not None for s in p.segs) and all(s.field is not None or not s.text.strip() for s in p.segs):
                    continue
                pi = self.para_index[id(p)]
                text = p.text
                self._texts[id(p)] = text  # 段落の字は 1 度だけ作る（Para.text は呼ぶたびに区切りをつなぎ直す）
                fmask = _field_mask(p)
                has_field = any(s_.field is not None for s_ in p.segs)  # 長さ 0 の欄（字のない部品）も、欄のある段落
                stops = _option_stops(text)
                bullet_id = f"P{pi}.b"
                bm = self.marks.get(bullet_id)
                if getattr(p, "bullet", None) is not None or (bm is not None and bm[0] is p):
                    # 箇条書きの印の □（段落番号の定義の字）：チェックの欄。書くと、この段落の印だけを ■・☑ にする。
                    # 見出しは、次の箱の字か 1 字ぶん以上の空白の並び（wide_space）までの字（そのまま）
                    end = stops[0] if stops and stops[0] > 0 else len(text)
                    opt = label_text(text[:end])  # 切らない（長い項目も、字のとおり）
                    self.items[bullet_id] = {"id": bullet_id, "kind": "checkbox", "source": "layout", "para": p,
                                             "bullet": getattr(p, "bullet", None), "label": opt, "provenance": LAYOUT,
                                             "filled": bm is not None}
                k = 0
                spans = []  # (始め, 終わり, 種類, 理由の鍵（None は規則の欄）)
                for i, ch in enumerate(text):
                    if ch in BOX and not (has_field and fmask[i]):
                        spans.append((i, i + 1, "checkbox", "bullet_box"))
                    elif ch in CHECKED and not (has_field and fmask[i]):
                        spans.append((i, i + 1, "checkbox", "checked_mark"))  # チェック済みの欄かもしれない（■ は箇条の印のことも）
                for a, b, why in para_spans(p, text):
                    if has_field and any(fmask[a:b]):
                        continue
                    if why in ("gap", "short_underline") and has_field and _touches_field(p, a, b):
                        continue  # フォームの欄にすぐ接する空白は、欄の字の飾り（欄はその id で書く）
                    spans.append((a, b, "inline", why))
                # 書き込み済みの欄（しおり）は、その名前を持ったまま（4 で作る）。残りの所の名前は、すべて集めてから
                # 字の位置で決める（_name_places：P段落.c字数。数え順の番号にしない）
                mine = marks_of.get(id(p), [])
                written = [(x, y) for _n, x, y in mine]
                if written:
                    spans = [t for t in spans if not _overlaps(t[0], t[1], written)]
                placed: list[tuple[dict, str]] = []  # (所, 古い番号の並び："" は P段落.番号、"a" は P段落.a番号)
                for a, b, kind, why in sorted(spans, key=lambda t: t[:3]):
                    if kind == "checkbox":
                        # 見出しは、箱の字の後ろから、次の箱の字か 1 字ぶん以上の空白の並び（wide_space）までの字（そのまま。字の意味では切らない）。
                        # 後ろに字がなければ、前の字
                        # 字は切らない。ただし区切りのないとても長い段落では、前後 LABEL_WINDOW 字まで（作る時間の上限。label_clipped）
                        end = _option_end(text, b, stops)
                        opt = text[b:min(end, b + LABEL_WINDOW)]
                        label = label_text(opt)
                        clipped = end > b + LABEL_WINDOW
                        if not label:
                            # 後ろに字がなければ、前の字：前の箱の字の後ろから（1 字ぶん以上の空きの前は含めない：label_tail）
                            start = max((x + 1 for x in stops if x < a and (text[x] in BOX or text[x] in CHECKED)), default=0)
                            label = label_text(label_tail(text[max(start, a - LABEL_WINDOW):a]).strip(" 　"))
                            clipped = start < a - LABEL_WINDOW
                    else:
                        label = self._label(p, a, b)
                    edges = (False, False)
                    if kind == "inline":
                        # 空白だけの長い空き：書いた字が見出しにくっつかないよう、端の 1 字を残した中に書く（窓の所も中だけ）
                        a, b, edges = gap_interior(text, a, b)
                    it = {"kind": kind, "source": "layout", "para": p, "span": (a, b), "label": label, "provenance": LAYOUT}
                    if kind == "checkbox" and clipped:
                        it["label_clipped"] = True
                    if any(edges):
                        it["kept_edges"] = edges
                    if why is not None:
                        it.update(provenance=AGENT, why=WORD_WHY[why], why_key=why)
                        if one_wide_space(text[a:b]):
                            it["why"] += f"。{ONE_WIDE_SPACE_WHY}"  # いちばん狭い空き：エージェントが順を決める目安
                        if why == "checked_mark":
                            it["checked"] = True
                    placed.append((it, ""))
                # 規則では決められない所：窓に「？」をつけて示し、エージェントが読んで判断する
                # （フォームの欄のある段落・図や埋め込みの印だけの段落（窓のために作った字だけ）は、探さない）
                found_agent = []
                look = not has_field and (text.strip(" \u3000") == "" or "".join(
                    s_.text for s_ in p.segs if not s_.hidden and s_.kind in ("t", "sym", "tab")).strip(" \u3000\t"))
                if look:
                    found_agent = agent_blanks(text, [(x, y) for x, y, _k, _w in spans] + written)
                    # 見出しのあるマスの中の空の段落（「〒」の下の行に住所を書くなど）。マスのほかの字が長いもの（説明の文のこともある）も、
                    # 理由を分けて示す（字数では決めない）
                    c = p.cell
                    if c is not None and not text.strip(" \u3000") and not spans and not written:
                        others = [q.text.strip(" \u3000") for q in c.paras() if q is not p and q.text.strip(" \u3000")]
                        before = [q for q in c.paras()[:c.paras().index(p)] if q.text.strip(" \u3000")] if p in c.paras() else []
                        if others and before:
                            found_agent = [(0, len(text), "next_line" if all(len(o) <= 16 for o in others) else "next_line_long")]
                        elif others:
                            # マスの初めの空の段落（後ろの段落にだけ字がある：「（空）／個人番号」の空の 1 行目に氏名を書くなど）
                            found_agent = [(0, len(text), "first_line")]
                    for a, b, why in found_agent:
                        if has_field and any(fmask[a:max(a + 1, b)]):
                            continue
                        label = self._label(p, a, b)
                        if why in ("first_line", "next_line", "next_line_long") and c is not None and c.table is not None:
                            # マスの中の空の段落：段落の字は空欄だけ（「［　］」）なので、ほかのマスの欄と同じく、表の行・列の見出し
                            label = _cell_label(c.table, c, self.written_cells)[0] or label
                        it = {"kind": "inline", "source": "layout", "para": p, "span": (a, b),
                              "label": label, "provenance": AGENT, "why": WORD_WHY[why], "why_key": why}
                        if why == "placeholder":
                            it["placeholder"] = text[a:b]
                        alt = self._beside_cell_field(p, why)
                        if alt:
                            it["alternative_to"] = alt
                        placed.append((it, "a"))
                self._name_places(p, pi, text, placed, mine)
        # 3') マスの中の見出し（「住所」「住所：」「所在地　〒」）の後ろ：すぐ右のマスに欄があれば、値はそちらに書くことが多い。
        # 消さずに残し、どちらに書くかはエージェントが窓を読んで決める（alternative_to：代わりの欄。規則の欄を先に）
        in_cell: dict[int, list[str]] = {}
        for fid_, it in self.items.items():
            c_ = it.get("cell") or (it["para"].cell if it.get("para") is not None else None)
            if c_ is not None:
                in_cell.setdefault(id(c_), []).append(fid_)
        for fid in [f for f, it in self.items.items() if it.get("why_key") in ("after_label", "after_colon")]:
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

    def _name_places(self, p: Para, pi: int, text: str, placed: list[tuple[dict, str]], mine: list) -> None:
        """段落の中の所に、位置から決まる名前を付ける：P{段落}.c{n}。n は、段落の頭からその所の始めまでにある、
        空白でも所（空欄・箱の字・書き込み済みの所）でもない字の数（＝その所の前にある、印刷された字の数）。
        - 名前は数え順の番号ではないので、ほかの所が増えても減っても（字の並びから推定する所の決まりが変わっても）、
          前にある印刷された字が変わらない限り同じ。書き込んだ所の字は数えないので、前の所に書いた後も同じ
        - 同じ n の所が並ぶとき（字をはさまずに続く空欄）は、2 つ目から P{段落}.c{n}.2、.3 …（位置の順）
        - 書き込み済みの所（しおり）の名前は使わない（その名前は、しおりの所のまま）
        古い名前（P{段落}.{番号}・P{段落}.a{番号}：数え順の番号）は、その段落にその並びの所が 1 つだけのとき（どの数え方でも
        同じ所を指すとき）だけ、legacy として受ける（Items.get）。"""
        cut = getattr(self.doc, "seg_cut", None)  # とても長い行の一部（lazy）：区切りの所は行の頭・終わりではない
        if cut:
            placed = [(it, c_) for it, c_ in placed if not (cut[1] and it["span"][0] >= len(text)) and not (cut[0] and it["span"][1] <= 0)]
        if not placed:
            return
        cover = bytearray(len(text) + 1)
        for it, _cls in placed:
            a, b = it["span"]
            cover[a:b] = b"\x01" * (b - a)
        for _n, x, y in mine:
            cover[x:y] = b"\x01" * (max(0, min(y, len(text)) - x))
        # prefix[i]：text[:i] の中の、空白でも所でもない字の数
        prefix = [0] * (len(text) + 1)
        n = 0
        for i, ch in enumerate(text):
            if not cover[i] and not ch.isspace() and not chars.is_space(ch):
                n += 1
            prefix[i + 1] = n
        taken = {name for name, _x, _y in mine}
        order = sorted(placed, key=lambda t: (t[0]["span"][0], t[0]["span"][1], t[1]))
        next_j: dict[str, int] = {}  # 同じ字数の所の、次に試す番号（同じ字数の所がとても多い段落でも、1 回ずつ進む）
        for it, _cls in order:
            at_abs = getattr(self.doc, "place_abs", None)  # とても長い行（lazy）：id は行の頭からの字の位置で決まる
            base = f"P{pi}.c{prefix[it['span'][0]]}" if at_abs is None else f"P{pi}.x{at_abs + it['span'][0]}"
            j = next_j.get(base, 1)
            fid = base if j == 1 else f"{base}.{j}"
            while fid in taken or fid in self.items:
                j += 1
                fid = f"{base}.{j}"
            next_j[base] = j + 1
            taken.add(fid)
            it["id"] = fid
            self.items[fid] = it
        # 古い名前：数え順の番号（書き込み済みの番号は除いた）。その並びの所が段落に 1 つだけのときだけ受ける
        for cls in ("", "a"):
            same = [it for it, c in sorted(placed, key=lambda t: (t[0]["span"][0], t[0]["span"][1])) if c == cls]
            olds = [name for name, _x, _y in mine if re.fullmatch(rf"P{pi}\.{cls}\d+", name)]
            if len(same) == 1 and not olds:
                self.items.legacy[f"P{pi}.{cls}1"] = same[0]["id"]
            elif same:
                for k in range(1, len(same) + len(olds) + 1):
                    self.items.ambiguous.add(f"P{pi}.{cls}{k}")

    def _beside_cell_field(self, p: Para, why: str) -> str | None:
        """マスの中の字（「住所：」「住所」）の後ろの所で、すぐ右のマスが空のマスの欄（そこに書くことが多い）：その規則の欄の id。
        候補は消さずに alternative_to（代わりの欄）として示す。マスの欄の位置は、1 度だけ索引にする。"""
        c = p.cell
        if c is None or why not in ("after_colon", "after_label"):
            return None
        idx = self.__dict__.get("_cell_fields")
        if idx is None:
            idx = self._cell_fields = {}
            for it in self.items.values():
                x = it.get("cell")
                if x is not None and it.get("kind") == "cell" and it.get("provenance") != AGENT:
                    idx.setdefault((id(x.table), x.row, x.col), it["id"])
        return idx.get((id(c.table), c.row, c.col + c.span))

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
        text = self._texts.get(id(p))
        if text is None:
            text = self._texts[id(p)] = p.text
        lo, hi = max(0, a - LABEL_WINDOW), min(len(text), b + LABEL_WINDOW)
        mine = per.get(id(p), [])
        near = mine[bisect.bisect_left(mine, (lo, -1)):bisect.bisect_right(mine, (hi, float("inf")))]
        cuts = [(x, y) for x, y in near if not (x == a and y == b) and y <= hi]
        out, at = [], lo
        for x, y in sorted(cuts + [(a, b)]):
            if x < at:
                continue
            out.append(text[at:x])
            out.append("［　］" if (x, y) == (a, b) else "　")
            at = y
        out.append(text[at:hi])
        return label_text("".join(out))

    def bound(self, c) -> "tuple | None":
        """結び付けた XML の値（w:dataBinding）：(Word が表示する値, たどれないときの理由, 値の所)。結び付けがなければ None。"""
        if getattr(c, "source", "") != "sdt" or not hasattr(self.doc, "zip"):
            return None
        memo = self.__dict__.setdefault("_bound", {})
        if id(c) not in memo:
            from .databind import binding_of, bound_value

            memo[id(c)] = bound_value(self.doc, c) if binding_of(c) is not None else None
        return memo[id(c)]

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
            b = self.bound(c)
            if b is not None and b[1] is None:
                # 文書の中の XML のデータに結び付いた欄（w:dataBinding）：Word は開くとき、その XML の値で表示し直す
                v = b[0]
                return None if v in ("", None) and c.placeholder and c.kind != "checkbox" else v
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


# 段落の中の所の名前の決まり（Word と同じ窓を使う形式すべて：.doc・.rtf・.odt・.jtd・.hwpx・.txt・.pptx）
PLACE_ID_NOTE = ("段落の中の所（12番目の段落の、前に印刷された字が 5 字ある所）。名前は位置で決まり、数え順の番号ではない：c の後ろの数は、"
                 "段落の頭からその所までにある、空白でも所でもない字の数（字をはさまず続く所は .c5.2・.c5.3）。ほかの所が増えても減っても、"
                 "前の所に書いた後も変わらない")

LABEL_WINDOW = 300  # 空欄の見出しに使う、空欄の前後の字の数（ふつうの段落はまるごと入る長さ。長い段落でも、見出しの長さと作る時間は空欄の数に比例するだけ）


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

    spec：{"old": 段落の中の元の字, "new": 書き換えた後の字, "near": 近くの欄の id, "near_paras": 段落の数,
    "occurrence": 何番目か（しぼった中の、文書の順。1 から）}。
    near でしぼる範囲（距離の決まった数は持たない。広さを変えるのはエージェント）：
    - near_paras を指定したとき：near の欄の段落の前後 near_paras 段落（0 なら同じ段落だけ）
    - 指定しないとき：near の欄が表のマスにあれば同じ表の同じ行（文書の作り）、なければ同じ話（本文・ヘッダー・脚注などの、同じ部分）。
      その中に候補がいくつもあり occurrence もないときは、near の欄の段落、次に near の欄のマスの中の候補に絞る（作りでいちばん近い所）
    old と new の前後の同じ字は、元の字と書式のまま残し、違う所だけを書き換える。元の字が見つからない・いくつもあって決まらないときは、候補を添えて断る。
    used：同じ書き込みの中で、もう使った番号と範囲（段落の番号 → [(番号, 始め, 終わり)]）。"""
    old, new = spec.get("old"), spec.get("new")
    if not isinstance(old, str) or not old or not isinstance(new, str):
        raise WriteError('字を目印にした書き込みは {"old": 段落の中の元の字, "new": 書き換えた後の字} の形で指定してください')
    if any(ch in old + new for ch in "\r\n\x07\x0b\x0c"):
        raise WriteError("old・new は、1つの段落の中の字にしてください（段落・行の区切りは書けません）")
    if old == new:
        raise WriteError("old と new が同じです")
    paras = [(sp.para_index[id(q)], q, (_k, _part)) for _k, _part, blocks in sp.doc.all_blocks() for q in iter_paras(blocks)]
    story = {id(q): st for _n, q, st in paras}
    hits = [(n, q, m.start()) for n, q, _st in paras for m in re.finditer(re.escape(old), q.text)]
    near = spec.get("near")
    span_n = spec.get("near_paras")
    if span_n is not None and (isinstance(span_n, bool) or not isinstance(span_n, int) or span_n < 0):
        raise WriteError("near_paras は 0 以上の整数（near の欄の段落の前後の段落の数）で指定してください")
    if span_n is not None and not near:
        raise WriteError("near_paras は near（近くの欄の id）と一緒に指定してください")
    if near:
        it = sp.items.get(str(near))
        if it is None:
            raise WriteError(f"near に指定した欄 {near} がありません")
        q0 = it.get("para") or (it["control"].para if it.get("control") is not None else None)
        n0 = sp.para_index.get(id(q0)) if q0 is not None else None
        c0 = q0.cell if q0 is not None else None

        def close(n: int, q: Para) -> bool:
            if span_n is not None:
                return n0 is not None and abs(n - n0) <= span_n
            if c0 is not None:
                return q.cell is not None and q.cell.table is c0.table and q.cell.row == c0.row
            return q0 is not None and story.get(id(q)) == story.get(id(q0))

        hits = [h for h in hits if close(h[0], h[1])]
    if not hits:
        raise WriteError(f"「{old}」が、{f'欄 {near} の近く' if near else '文書'}に見つかりません（窓の字のとおりに指定してください）")
    occ = spec.get("occurrence")
    if occ is None and near and span_n is None and len(hits) > 1:
        # near の欄と作りでいちばん近い候補：同じ段落、次に同じマス。それでも 1 つに決まらなければ、広い範囲の候補を示して断る
        for same in (lambda q: q is q0, lambda q: c0 is not None and q.cell is c0):
            mine = [h for h in hits if same(h[1])]
            if mine:
                hits = mine
                break
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


def cell_neighbours(t: Table, c: Cell, written: set | None = None) -> dict:
    """表のマスの、見出しになりうる隣のマス（事実だけ。どれを見出しとして読むかはエージェント）：
    - row_label：同じ行で、左にある、いちばん近い字のあるマスの字（印刷される字のとおり。切らない）と、
      row_label_gap：そのマスとの間にある、字のないマスの列の数（すぐ左なら 0）
    - column_label：上にある、いちばん近い字のあるマス（このマスの列を含むマス）の字と、column_label_gap：間の行の数（すぐ上なら 0）。
      上のマスがこのマスの列の幅をすべては覆わない（結合して 2 列にまたがるマスの上の、1 列だけのマス）ときは、
      column_label_partial（覆う列だけの見出しかもしれない）
    todoku が書いたマス（written：書き込みのしおりのあるマス）の字は、印刷された見出しではないので使わない。"""
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
    if up is not None and (id(up) in written or not _printed(up)):
        up = None
    out: dict = {}
    if left is not None:
        out["row_label"] = _printed(left)
        out["row_label_gap"] = c.col - (left.col + left.span)
    if up is not None:
        out["column_label"] = _printed(up)
        out["column_label_gap"] = c.row - up.row - 1
        if c.col + c.span > up.col + up.span:
            out["column_label_partial"] = True
    return out


def _cell_label(t: Table, c: Cell, written: set | None = None) -> tuple[str, str, str]:
    """表の空のマスの見出し：(見出し, 行の見出し, 列の見出し)。見出しは、いつも同じ決まりで作る：
    行の見出し（左のマス）と列の見出し（上のマス。このマスの列の幅をすべて覆うものだけ）の両方があれば「行／列」、
    片方だけならその字。距離・覆い方で片方を落とすことはしない（距離は cell_neighbours の事実として窓に出す）。"""
    nb = cell_neighbours(t, c, written)
    rl = label_text(nb.get("row_label", ""))
    cl = "" if nb.get("column_label_partial") else label_text(nb.get("column_label", ""))
    if rl and cl and rl != cl:
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
        "notation": {"［T1.R3.C2］": "表の空のマス（1つ目の表の3行目・2列目）",
                     "［P12.c5］": PLACE_ID_NOTE,
                     "［P12.c5？］": "規則では決められない、書き込む所かもしれない所（字の形と並びだけで見つけた所。エージェントが読んで判断する。"
                                    "書くなら id を指定する。書かなければ何も変わらない。［P12.c5？＝○○］は、伏せ字の○○を書き換える所）",
                     "［P12.c5？　］": GAP_TOKEN_NOTE,
                     "■［P12.c5？］": "チェック済みの印らしい字（☑・■・✓ など。■ は箇条の印のこともある）。チェックの欄なら true / false で書く",
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
    from ..embeds import lazy as embed_lazy

    embed_lazy.spend(sum(len(ln) for st in out["stories"] for ln in st["reading"]))
    out["links"] = [{"text": t, "target": u} for t, u in getattr(doc, "links", [])]
    if getattr(doc, "numbering_problem", None):
        out.setdefault("unread", []).append(doc.numbering_problem)
    out["embeds"] = _embeds(doc)
    out["pictures"] = _pictures(doc)
    from ..embeds.word import objects as _objects

    out["objects"] = _objects(doc, shown={e["n"] for e in out["embeds"]})  # グラフ・SmartArt・Excel の表のほかの埋め込み（読むだけ）
    contexts = {}
    for st in out["stories"]:
        for line in st["reading"]:
            for m in re.finditer(r"［([^］＝？]+)？?[ 　]*(?:＝[^］]*)?］", line):  # 「？」の所の印は、元の空白を中に入れる
                contexts.setdefault(m[1], line)
    legacy_of = {new: old for old, new in sp.items.legacy.items()}  # いまの名前 → 受ける古い数え順の名前
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
        if c is None and "span" in it and it.get("para") is not None and it["id"] not in sp.marks:
            # 段落の中の所の事実（所の字・字数・幅の目安・前後の字・字間を空けた並びの字数。xlsx.window.place_facts）
            q = it["para"]
            d["facts"] = place_facts(sp._texts.get(id(q)) or q.text, *it["span"])
        if it["id"] in legacy_of:
            d["legacy_id"] = legacy_of[it["id"]]  # 古い数え順の名前（その段落にこの並びの所が 1 つだけなので、同じ所を指す）
        if c is not None and c.para is not None and c.para.cell is not None and c.para.cell.table is not None:
            d.update(cell_neighbours(c.para.cell.table, c.para.cell, sp.written_cells))  # マスの中の欄：左・上のマスの字（事実）
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
        bnd = sp.bound(c) if c is not None else None
        if bnd is not None:
            # 文書の中の XML のデータに結び付いた欄：窓の値は、Word が開くときに表示する、その XML の値
            from .databind import binding_of

            d["bound_to"] = f"{bnd[2].part} {binding_of(c).xpath}" if bnd[2] is not None else binding_of(c).xpath
            if bnd[1]:
                d["bound_problem"] = bnd[1]
        if it.get("provenance") == AGENT:
            d["agent"] = True
            d["why"] = it["why"]
            if it.get("why_key"):
                d["why_code"] = it["why_key"]  # 理由の符号（字の説明 why は変わることがある。判断・段階はこの符号で）
            if it.get("placeholder"):
                d["placeholder"] = it["placeholder"]
            if it.get("alternative_to"):
                d["alternative_to"] = it["alternative_to"]  # 同じ値を書く、代わりの欄（どちらか一方に書く）
        if it.get("filled") or (it["id"] in sp.marks and c is None):
            d["filled"] = True  # todoku が書き込んだ所（書き直すには元の様式に書く）
        d.update(it.get("neighbours") or {})  # 表のマス：左・上のマスの字（そのまま）と、間のマスの数（row_label_gap・column_label_gap）
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
    with_label_keys(out["fields"])  # 見出しは字のまま。探すための鍵（字間の空白を詰めた写し）は label_key
    return out


def _pictures(doc) -> list[dict]:
    """文書の絵：説明（代わりの字）があるか、飾りの印があるか。説明のない絵は、本文に〔図N（説明なし）〕と出す。
    where は、その絵のある段落の字（本人に場所を伝えるため）。"""
    out = []
    for pic in getattr(doc, "pictures", []) or []:
        near = re.sub(r"〔[^〕]*〕", "", pic["para"].text).strip()
        out.append({"id": f"図{pic['n']}", "alt": pic["alt"] or None, "decorative": pic["decorative"],
                    "where": f"「{near[:30]}」の段落" if near else "字のない段落", "media": pic.get("media", ""),
                    **({"size_mm": pic["size_mm"]} if pic.get("size_mm") else {})})
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
    """(埋め込み, (Excel の窓, 窓の Markdown の行)・作れなかった理由の字・後回しにした概要（dict）)。Excel のブックでない埋め込みは除く。"""
    from .. import limits
    from ..embeds import lazy as embed_lazy
    from ..xlsx import window as xw

    if not hasattr(doc, "embedded_book"):
        return []
    out = []
    embeds = list(getattr(doc, "embeds", []) or [])
    for i, e in enumerate(embeds):
        limits.progress("埋め込みの表を読んでいます", i, len(embeds))
        brief = embed_lazy.later(doc, e)
        if brief is not None:
            out.append((e, brief))  # 窓に入らない：頼まれたときに、頼まれた行だけ読む
            continue
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
            if isinstance(got, tuple):
                embed_lazy.spend(sum(len(ln) for ln in got[1]))
    return out


def _embeds(doc) -> list[dict]:
    """文書に埋め込んだ Excel の表：Excel の窓と同じ読み方で、シートごとに平文にする（本文では〔埋め込み1：…〕の所）。"""
    from ..embed import window_info

    out = []
    for e, got in _embed_windows(doc):
        if isinstance(got, str):
            out.append({"n": e["n"], "label": e["label"], "prog": e["prog"], "storage": e["storage"], "error": got})
            continue
        if isinstance(got, dict):
            out.append(got)
            continue
        xdoc, md = got
        info = window_info(doc, e, xdoc)  # 書くときの欄の名前（E1!Sheet1!C4）と、表の絵を描き直せるか
        out.append({"n": e["n"], "label": e["label"], "prog": e["prog"], "storage": e["storage"],
                    "sheets": [sh["name"] for sh in xdoc["sheets"]], "markdown": list(md), **info})
    return out


def printed_label(sp: "Spots", it: dict) -> str:
    """入力欄（フォームの欄・コンテンツコントロール）の見出し：同じ段落の、欄の前に印刷される字（ほかの欄の値・todoku が
    書いた所は除く。区切り（タブ・1 字ぶん以上の空白の並び（wide_space）・表のマスの区切り）より前は含めない）。段落の頭にある欄で、表のマスの中なら、
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
    # 区切り：ほかの欄・タブ・改行と、1 字ぶん以上の空白の並び（wide_space：窓の空きと同じ 1 em の決まり。欄の直前の空きは
    # 除いて、その前の塊。字間を空けた見出し「氏　名」はまとめる：label_tail）
    head = label_tail(re.split(r"\x00|\t|\n", "".join(out))[-1])
    trim = "".join(ch for ch in set(head) if chars.is_space(ch) or chars.is_label_colon(ch))
    head = label_text(head.strip(" 　")).rstrip(trim)  # 見出しのあとのコロン（: ： ﹕ など）と空白を除く
    if head:
        return head
    cell = p.cell
    if cell is not None and cell.table is not None:
        # 段落の頭の欄：いちばん近い見出しのマスの字をそのまま（同じ行の左のマスがあればそれ、なければ上のマス）。
        # 2 つをつないで 1 つの見出しにはしない（両方の字と間のマスの数は、欄の row_label・column_label に出す）
        nb = cell_neighbours(cell.table, cell, sp.written_cells)
        return label_text(nb.get("row_label") or nb.get("column_label") or "")
    return ""


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
    index = sp.__dict__.get("_by_cell")
    if index is None or index[0] != len(sp.items):  # マスごとに全部の欄をなめない（大きな表で 2 乗になる）
        cells: dict[int, str] = {}
        for it in sp.items.values():
            if it.get("cell") is not None:
                cells.setdefault(id(it["cell"]), it["id"])
        index = sp.__dict__["_by_cell"] = (len(sp.items), cells)
    return index[1].get(id(c))


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
                # 空白だけの所は、元の空白をそのまま印の中に入れる（「［P1.1？　］」：全角の空白 1 つの空き。shown_gap）
                edits.append((a, b, f"［{name}？＝{text[a:b]}］" if it.get("placeholder") else f"［{name}？{shown_gap(text[a:b])}］"))
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
    out += [f"- **読んでいない部分**：{u}" for u in doc.get("unread") or []]  # JSON だけでなく、エージェントの読む窓にも出す
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
    from ..links import markdown as links_md

    out += links_md([(x["text"], x["target"]) for x in doc.get("links") or []])
    for e in doc.get("embeds") or []:
        out += [f"## 埋め込み{e['n']}：{e['label']}（{e['prog']}・{e['storage']}）", ""]
        if e.get("error"):
            out += [f"- この埋め込みは開けませんでした：{e['error']}", ""]
            continue
        if e.get("deferred"):
            from ..embeds.lazy import markdown as later_md

            out += later_md(e)
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
