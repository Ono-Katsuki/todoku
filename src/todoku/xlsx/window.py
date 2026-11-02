"""AIや支援技術が読むための「窓」を作る。元のファイルは正本のまま残る。

窓はJSON（機械向け）とMarkdown（人とAI向け）の2つ。どちらも決定的な処理だけで作り、
すべての項目に「どうやって得たか」（provenance）を付ける。

読み順は、画面に見える行を上から、行の中は左から並べる。結合セルは1マス、入力欄は［］、
計算結果は〈〉、チェックボックスは☐☑で示す。見出しは配置から決めるが、推定であることを明示し、
判断の材料になった行（context）を必ず添える。
"""

from __future__ import annotations

import hashlib
import re
import unicodedata
from dataclasses import dataclass, field

from ..formula.numfmt import format_value, is_date_format
from ..formula.parser import RefNode, addr, iter_refs
from ..formula.values import ExcelError, Unsupported, is_number, to_text
from .workbook import Cell, Control, Sheet, Workbook

DET = "deterministic"
LAYOUT = "layout"  # 配置から決定的に推定した（正しいとは限らないので context と合わせて読む）


def norm_label(s: str) -> str:
    """見出しの空白と改行を1つにまとめ、「事 業 所 番 号」のように字間を空けたものは詰める。文字そのものは変えない。"""
    t = str(s).strip()
    parts = re.split(r"[ \u3000\r\n\t]+", t)
    if len(parts) >= 3 and all(len(p) <= 2 for p in parts):
        return "".join(parts)
    return re.sub(r"[ \u3000\r\n\t]+", " ", t)


def display(wb: Workbook, cell: Cell) -> str:
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, ExcelError):
        return v.code
    fmt = wb.style(cell.style).num_fmt
    try:
        return format_value(v, fmt)
    except (Unsupported, Exception):  # noqa: BLE001
        return to_text(v)


@dataclass
class Anchor:
    r: int
    c1: int
    c2: int
    r2: int
    cell: Cell | None
    text: str


def anchors(sh: Sheet) -> dict[int, list[Anchor]]:
    """行ごとに、表示上の1マス（結合セルは1つ）を左から並べる。"""
    merged_cover: dict[tuple[int, int], tuple[int, int, int, int]] = {}
    for m in sh.merges:
        if (m[2] - m[0] + 1) * (m[3] - m[1] + 1) > 20000:
            continue
        for r in range(m[0], m[2] + 1):
            for c in range(m[1], m[3] + 1):
                merged_cover[(r, c)] = m
    rows: dict[int, list[Anchor]] = {}
    seen = set()
    keys = set(sh.cells) | {(m[0], m[1]) for m in sh.merges}
    for (r, c) in sorted(keys):
        m = merged_cover.get((r, c))
        if m and (r, c) != (m[0], m[1]):
            continue
        if (r, c) in seen:
            continue
        seen.add((r, c))
        cell = sh.cells.get((r, c))
        r2, c2 = (m[2], m[3]) if m else (r, c)
        rows.setdefault(r, []).append(Anchor(r, c, c2, r2, cell, ""))
    return rows


def input_cells(wb: Workbook, sh: Sheet) -> dict[tuple[int, int], list[str]]:
    """入力欄の候補と、そう判断した理由。"""
    out: dict[tuple[int, int], list[str]] = {}

    def add(r, c, why):
        m = sh.merged_at(r, c)
        if m:
            r, c = m[0], m[1]
        cell = sh.cells.get((r, c))
        if cell is not None and cell.has_formula:
            return
        if sh.protected and wb.style(sh.style_of(r, c)).locked:
            return  # 保護されたシートのロックされたセルは、Excelでも書き込めない
        out.setdefault((r, c), [])
        if why not in out[(r, c)]:
            out[(r, c)].append(why)

    if sh.protected:
        for (r, c), cell in sh.cells.items():
            if not wb.style(cell.style).locked and not cell.has_formula:
                add(r, c, "シート保護の中で編集できるセル")
        for m in sh.merges:
            if (m[0], m[1]) not in sh.cells and not wb.style(sh.style_of(m[0], m[1])).locked:
                add(m[0], m[1], "シート保護の中で編集できるセル")
    for v in sh.validations:
        for r1, c1, r2, c2 in v.sqref:
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 5000:
                continue
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    add(r, c, "入力規則があるセル")
    if not sh.protected:
        # 保護も入力規則もない様式：枠で囲まれた空欄・色の付いた空欄で、見出しのそばにあるもの
        for (r, c), why in _boxed_blanks(wb, sh).items():
            add(r, c, why)
    # 数式が参照している定数・空白のセル（計算の入力）
    for cell in sh.cells.values():
        refs = list(iter_refs(cell.ast)) if cell.ast is not None else []
        for ref in refs:
            if not isinstance(ref, RefNode):
                continue
            ref = ref.shifted(cell.dr, cell.dc)
            if ref.sheet and ref.sheet != sh.name:
                continue
            if (ref.r2 - ref.r1 + 1) * (ref.c2 - ref.c1 + 1) > 400:
                continue
            for r in range(ref.r1, ref.r2 + 1):
                for c in range(ref.c1, ref.c2 + 1):
                    t = sh.cells.get((r, c))
                    if t is None or (not t.has_formula and (t.value is None or is_number(t.value))):
                        add(r, c, "数式が参照している値")
    return out


def undecided_cells(wb: Workbook, sh: Sheet, inputs: dict) -> dict[tuple[int, int], str]:
    """入力欄かどうかを、規則では決められないセル：表（列割りが同じ行の連続）の行に入力欄があるのに、
    同じ行の、見出しのある列の空のセルに、入力欄の印（入力規則・枠・ロックの解除）がないもの。
    入力欄とは言い切らず、「エージェントが見出しから判断する所」として示す。値 → その列の見出し。"""
    if sh.state in ("hidden", "veryHidden"):
        return {}  # 非表示のシート（選択肢の元の表など）は、記入する所ではない
    rows = anchors(sh)
    out: dict[tuple[int, int], str] = {}
    for hr, spans, body in detect_tables(wb, sh, rows):
        heads = {a.c1: norm_label(a.cell.value) for a in rows[hr]}
        for br in body:
            byc = {a.c1: a for a in rows.get(br, [])}
            if not any((br, c1) in inputs for c1, _c2 in spans):
                continue
            for c1, _c2 in spans:
                if (br, c1) in inputs or sh.is_hidden(br, c1):
                    continue
                a = byc.get(c1)
                cell = a.cell if a is not None else sh.cells.get((br, c1))
                if cell is not None and (cell.has_formula or cell.value not in (None, "")):
                    continue
                if sh.protected and wb.style(sh.style_of(br, c1)).locked:
                    continue  # 保護されたシートのロックされたセルは、書けない
                out[(br, c1)] = heads.get(c1, "")
    return out


def _boxed_blanks(wb: Workbook, sh: Sheet) -> dict[tuple[int, int], str]:
    """枠線で囲まれた空欄（または色の付いた空欄）で、左か上に見出しの文字があるもの。

    保護も入力規則もない様式では、記入する場所は「枠のある空欄」として描かれている。見出しのない枠は数えない。
    """
    out: dict[tuple[int, int], str] = {}
    rows = anchors(sh)

    def bits(r: int, c: int) -> int:
        return wb.style(sh.style_of(r, c)).borders if r >= 1 and c >= 1 else 0

    def text_of(a: Anchor) -> bool:
        return a.cell is not None and isinstance(a.cell.value, str) and bool(a.cell.value.strip()) and not a.cell.has_formula

    covered: dict[tuple[int, int], Anchor] = {}
    for items in rows.values():
        for a in items:
            for rr in range(a.r, min(a.r2, a.r + 50) + 1):
                covered[(rr, a.c1)] = a
    for r, items in rows.items():
        for k, a in enumerate(items):
            cell = a.cell
            if cell is not None and (cell.has_formula or cell.value not in (None, "")):
                continue
            if sh.is_hidden(a.r, a.c1):
                continue
            left = bits(a.r, a.c1) & 1 or bits(a.r, a.c1 - 1) & 2
            right = bits(a.r, a.c2) & 2 or bits(a.r, a.c2 + 1) & 1
            top = bits(a.r, a.c1) & 4 or bits(a.r - 1, a.c1) & 8
            bottom = bits(a.r2, a.c1) & 8 or bits(a.r2 + 1, a.c1) & 4
            edges = sum(bool(x) for x in (left, right, top, bottom))
            colored = wb.style(sh.style_of(a.r, a.c1)).fill_color not in (None, "FFFFFF")
            if edges < 3 and not (bottom and edges >= 1) and not colored:
                continue
            # 見出し：同じ行の左（あいだに文字のマスをはさまない、空き4列まで）か、上（3行まで）
            labeled = False
            for b in reversed(items[:k]):
                if text_of(b):
                    labeled = a.c1 - b.c2 - 1 <= 4
                    break
                if b.cell is not None and b.cell.value not in (None, ""):
                    break
            if not labeled:
                for rr in range(a.r - 1, max(0, a.r - 4), -1):
                    above = [b for b in rows.get(rr, []) if b.c1 <= a.c2 and b.c2 >= a.c1]
                    above += [x for x in (covered.get((rr, a.c1)),) if x is not None and x not in above]
                    if any(text_of(b) for b in above):
                        labeled = True
                        break
                    if any(b.cell is not None and b.cell.value not in (None, "") for b in above):
                        break
            if labeled:
                out[(a.r, a.c1)] = "色の付いた空欄" if colored and edges < 3 else "枠のある空欄"
    return out


_BLANK_UNITS = ("年間", "か月", "ヶ月", "カ月", "時間", "年度", "年", "月", "日", "号", "円", "人", "名", "件", "回", "時",
                "分", "歳", "才", "枚", "部", "台", "）", ")", "】", "］")
_TOK = re.compile(r"令和|平成|昭和|大正|明治|西暦|[ \u3000]+|[_＿]{3,}|.", re.S)
_ERAS = {"令和", "平成", "昭和", "大正", "明治", "西暦"}


_COUNT_PREFIX = set("第金約計全各満延毎総月年週時")


def _letter(t: str) -> bool:
    """字（漢字・かな・ラテン文字など）1字か、年号。"""
    import unicodedata

    return t in _ERAS or (len(t) == 1 and unicodedata.category(t)[0] == "L")


def inline_blanks(text: str) -> list[tuple[int, int]]:
    """文字の中の、書き込むための空欄（「令和　　年　　月　　日」の空白、「第　　　号」の空白、「＿＿＿」）の位置。

    - 単位（年・月・日・号・円など）や閉じかっこの直前の、2つ以上の空白
    - 「第」「（」「：」の直後の空白、全角の「（」の直前の空白、文の終わりの「）」の後の空白（電話番号の欄など）
    - 「平成　年　月　日」のように、年号と年・月・日が1つの空白でつながったもの
    - 下線（＿＿＿、___）
    字間を空けた見出し（「廃　業　年　月　日」）の空白や、飾りの空白（「北海道知事　　　様」）は空欄にしない。"""
    # 2字を空けた見出し（「氏　　名」「番　　号」）：後ろの字が単位でも、空欄にしない。
    # 「第　　号」「金　　円」「月　　回」のように、前の字が数を受ける字のときだけ空欄
    two = re.fullmatch(r"[ \u3000]*(\S)[ \u3000]+(\S)[ \u3000]*", text)
    if two and _letter(two[1]) and _letter(two[2]) and two[1] not in _COUNT_PREFIX:
        return []
    toks = [(m.start(), m.end(), m[0]) for m in _TOK.finditer(text)]
    # 同じ幅の空白でつながった並び（字間を空けた見出しか、年月日の並び）。
    # 1つの空白ならどんな字でも、2つ以上の空白なら字（漢字・かな・年号）だけの並び（「職　　業　　名」「令和　　年　　月　　日」）
    chain_space: dict[int, bool] = {}
    i = 0
    while i < len(toks):
        j = i
        while j + 2 < len(toks) and toks[j + 1][2].isspace() and not toks[j][2].isspace() and not toks[j + 2][2].isspace() \
                and toks[j + 1][2] == toks[i + 1][2] and \
                (toks[j + 1][2] in (" ", "\u3000") or (_letter(toks[j][2]) and _letter(toks[j + 2][2]))):
            j += 2
        if j - i >= 4:  # 3つ以上の文字が1つの空白でつながっている
            words = [toks[k][2] for k in range(i, j + 1, 2)]
            body = words[1:] if words[0] in _ERAS else words
            date = all(w in ("年", "月", "日") for w in body) and len(body) >= 2
            for k in range(i + 1, j, 2):
                chain_space[k] = date
            i = j + 1
        else:
            i += 1
    out = []
    for k, (a, b, t) in enumerate(toks):
        if t[0] in "_＿":
            out.append((a, b))
            continue
        if not t.isspace():
            continue
        if k in chain_space:
            if chain_space[k]:
                out.append((a, b))
            continue
        after = text[b:b + 2]
        before = text[max(0, a - 1):a]
        jp_before = a > 0 and ord(text[a - 1]) > 0x2E80
        unit_after = any(after.startswith(u) for u in _BLANK_UNITS)
        if b - a == 1:
            # 1つの空白：日本語の文字のあとの半角の空白＋単位（「から 年間」）だけ
            if t == " " and jp_before and unit_after:
                out.append((a, b))
            continue
        opener_before = before in ("第", "（", "：", "－", "〒", "№", "(", ":") or (a == 0 and unit_after)
        paren_after = after[:1] == "（" and jp_before
        tail = b == len(text) and before == "）"
        if unit_after or (opener_before and b < len(text)) or paren_after or tail:
            out.append((a, b))
    return out


def fill_blanks(text: str, values: list) -> str:
    """空欄に、順に値を入れた文字列。"""
    spans = inline_blanks(text)
    if len(values) != len(spans):
        raise ValueError(f"空欄は {len(spans)} か所あります（値は {len(values)} 個）")
    out, at = [], 0
    for (a, b), v in zip(spans, values):
        out.append(text[at:a])
        out.append(to_text(v) if v is not None else text[a:b])
        at = b
    out.append(text[at:])
    return "".join(out)


def detect_tables(wb: Workbook, sh: Sheet, rows: dict[int, list[Anchor]]):
    """見出し行と、同じ列割りで続く行を表として取り出す。"""
    tables = []
    used = set()
    rlist = sorted(rows)
    for i, r in enumerate(rlist):
        if r in used:
            continue
        head = [a for a in rows[r] if a.cell is not None and isinstance(a.cell.value, str) and a.cell.value.strip() and not a.cell.has_formula]
        if len(head) < 3 or len(head) != len(rows[r]):
            continue
        spans = [(a.c1, a.c2) for a in rows[r]]
        body = []
        for r2 in rlist[i + 1:]:
            if r2 != (body[-1] if body else r) + 1:
                break
            sp2 = [(a.c1, a.c2) for a in rows[r2]]
            if not set(sp2) <= set(spans) or len(sp2) < 2:
                break
            body.append(r2)
        if len(body) >= 1:
            tables.append((r, spans, body))
            used.update(body)
    return tables


# ====================================================================== 配置
@dataclass
class Box:
    r: int
    c1: int
    r2: int
    c2: int
    kind: str  # text / input / formula / check
    cell: Cell | None = None
    text: str = ""  # 表示（入力欄・数式は値の表示）
    control: Control | None = None
    outside: bool = False  # 印刷範囲の外（画面だけに出る）
    highlighted: bool = False  # 条件付き書式で色が付いている
    blanks: list = field(default_factory=list)  # 文字の中の空欄（書き込む場所）

    @property
    def ref(self) -> str:
        if self.control is not None:
            ctl = self.control
            if ctl.link_cell:
                return addr(*ctl.link_cell[1:])  # つながったセル（別のシートなら field の id にシート名が入る）
            return ctl.name
        a = addr(self.r, self.c1)
        return a if (self.r2, self.c2) == (self.r, self.c1) else f"{a}:{addr(self.r2, self.c2)}"

    @property
    def cell_ref(self) -> str:
        return addr(self.r, self.c1)


_UNIT_WORDS = ("千円", "万円", "円", "か月", "ヶ月", "カ月", "ケ月", "か所", "ヶ所", "箇所", "日間", "週間", "年間", "時間",
               "世帯", "年度", "年", "月", "日", "人", "名", "件", "回", "分", "秒", "歳", "才", "号", "番", "割", "倍", "時",
               "台", "枚", "部", "冊", "本", "個", "週", "点", "級", "階", "室", "床", "％", "%", "㎡", "km", "kg")
_ERA_WORDS = ("令和", "平成", "昭和", "大正", "明治", "西暦", "〒", "第", "から", "まで", "No.", "№")
_SEP = "-－‐―～~〜・/／（）()「」 　:：,、"
_TOKEN = re.compile("(?:" + "|".join(re.escape(w) for w in sorted(_UNIT_WORDS + _ERA_WORDS, key=len, reverse=True))
                    + "|[" + re.escape(_SEP) + "])+")
_MARKER = re.compile(r"^[\(（]?(?:[①-⑳㉑-㉟㊱-㊿]|[0-9０-９]{1,2}|[イロハニホヘトチリヌ]|[ⅰ-ⅻⅠ-Ⅻ]+|[a-zａ-ｚ])[\)）．.]?$")


def _is_phrase_text(t: str) -> bool:
    t = t.strip()
    return bool(t) and len(t) <= 6 and _TOKEN.fullmatch(t) is not None


def _unit_of(t: str) -> str | None:
    t = t.strip().strip("()（）[]［］")
    return t if t in _UNIT_WORDS else None


class Layout:
    """1枚のシートの、目に見える配置。"""

    def __init__(self, wb: Workbook, sh: Sheet, calc=None):
        from .cf import Looks

        self.wb = wb
        self.sh = sh
        self.looks = Looks(wb, calc)
        self.hidden_by_format: list[tuple[int, int]] = []
        self.inputs = input_cells(wb, sh)
        self.links = {}
        for s in wb.sheets:
            for ctl in s.controls:
                # 部品につながったセルは、部品として示す（同じ欄を、セルとしても二重に出さない）
                if ctl.kind in ("CheckBox", "Radio", "Drop", "List", "Spin", "Scroll") and ctl.link_cell \
                        and ctl.link_cell[0] == sh.name:
                    self.links[(ctl.link_cell[1], ctl.link_cell[2])] = ctl
                if ctl.kind == "GBox" and ctl.fmla_group and s.name == sh.name:
                    from .workbook import _link_cell

                    lc = _link_cell(ctl.fmla_group, s.name)
                    if lc and lc[0] == sh.name:
                        self.links[(lc[1], lc[2])] = ctl
        self.rows: dict[int, list[Box]] = {}
        self.cover: dict[int, list[Box]] = {}  # 上の行から続く結合セル（その行で始まらないもの）
        for r, items in anchors(sh).items():
            for a in items:
                if sh.is_hidden(a.r, a.c1):
                    continue
                cell = a.cell
                outside = not sh.in_print_area(a.r, a.c1)
                if (a.r, a.c1) in self.inputs:
                    if (a.r, a.c1) in self.links:
                        continue  # チェックボックスのつながり先は、チェックボックスとして示す
                    b = Box(a.r, a.c1, a.r2, a.c2, "input", cell, display(wb, cell) if cell else "", outside=outside)
                elif cell is not None and cell.has_formula:
                    b = Box(a.r, a.c1, a.r2, a.c2, "formula", cell, display(wb, cell), outside=outside)
                elif cell is None or cell.value is None or cell.value == "":
                    continue
                else:
                    b = Box(a.r, a.c1, a.r2, a.c2, "text", cell, display(wb, cell), outside=outside)
                    if isinstance(cell.value, str) and not (sh.protected and wb.style(cell.style).locked):
                        b.blanks = inline_blanks(cell.value)
                self.rows.setdefault(a.r, []).append(b)
                for rr in range(a.r + 1, a.r2 + 1):
                    self.cover.setdefault(rr, []).append(b)
        under = set()
        from .controls import CHOICE_KINDS, VALUE_KINDS

        self.control_items: dict[str, list[str]] = {}
        for ctl in sh.controls:
            if ctl.kind not in ("CheckBox", "Radio") + CHOICE_KINDS + VALUE_KINDS or ctl.anchor is None:
                continue
            r = self.control_row(ctl)
            c = ctl.anchor[1]
            if sh.is_hidden(r, c):
                continue
            self.rows.setdefault(r, []).append(Box(r, c, r, c, "check", None, self.control_mark(ctl), ctl))
            (r1, c1), (r2, c2) = ctl.anchor, ctl.anchor_to or ctl.anchor
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    under.add((rr, cc))
        for r, items in self.rows.items():
            # チェックボックスの下に敷かれた空欄（クリックできるようにロックを外してあるもの）は、欄として数えない
            self.rows[r] = [b for b in items if not (b.kind == "input" and not b.text and (b.r, b.c1) in under)]
        if sh.cond_rules:
            for r in list(self.rows):
                keep = []
                for b in self.rows[r]:
                    if b.kind in ("text", "formula", "input") and any(x.covers(b.r, b.c1) for x in sh.cond_rules):
                        lk = self.looks.look(sh, b.r, b.c1)
                        b.highlighted = lk.highlighted
                        if lk.hidden and b.kind != "input" and b.text:
                            # 条件付き書式で、いまは背景と同じ色などになって見えない文字
                            self.hidden_by_format.append((b.r, b.c1))
                            continue
                    keep.append(b)
                self.rows[r] = keep
            hidden = set(self.hidden_by_format)
            for r in list(self.cover):
                self.cover[r] = [b for b in self.cover[r] if (b.r, b.c1) not in hidden]
        for r in self.rows:
            self.rows[r].sort(key=lambda b: (b.c1, b.kind != "check"))

    def control_mark(self, ctl: Control) -> str:
        """読み順に出す部品の状態：☑☐（チェック）・◉○（オプションボタン）・▼（ドロップダウン）・≡（リスト）・⇕（スピン）。"""
        from .controls import items_of

        if ctl.kind == "CheckBox":
            return "☑" if ctl.is_checked else "☐"
        if ctl.kind == "Radio":
            return ("◉" if ctl.checked == "Checked" else "○") + (ctl.text.strip() or "")
        if ctl.kind in ("Drop", "List"):
            items = self.control_items.setdefault(ctl.shape_id, items_of(self.wb, self.sh, ctl))
            if ctl.kind == "List" and (ctl.sel_type or "single") != "single":
                chosen = [items[i - 1] for i in (ctl.multi_sel or []) if 0 < i <= len(items)]
                return "≡" + "、".join(chosen)
            cur = items[ctl.sel - 1] if ctl.sel and 0 < ctl.sel <= len(items) else ""
            return ("▼" if ctl.kind == "Drop" else "≡") + cur
        return f"⇕{ctl.val if ctl.val is not None else ''}"

    # ------------------------------------------------------------------ 位置
    def height(self, r: int) -> float:
        if r in self.sh.hidden_rows:
            return 0.0
        return self.sh.row_heights.get(r, self.sh.default_row_height)

    def control_row(self, ctl: Control) -> int:
        """チェックボックスが見た目で属する行（つながり先が同じ範囲の行ならそれ、なければ縦の中心）。"""
        r1 = ctl.anchor[0]
        r2 = (ctl.anchor_to or ctl.anchor)[0]
        if ctl.link_cell and ctl.link_cell[0] == self.sh.name and r1 <= ctl.link_cell[1] <= r2:
            return ctl.link_cell[1]
        if r1 == r2:
            return r1
        c = ctl.anchor[1]
        best = None
        for r in range(r1, r2 + 1):
            for b in self.rows.get(r, []) + self.cover.get(r, []):
                if b.kind == "text" and b.c1 > c and isinstance(b.cell.value, str) and b.cell.value.strip():
                    key = (b.c1, b.r)
                    if best is None or key < best[0]:
                        best = (key, b.r)
        if best is not None:
            return best[1]
        tops = {r1: 0.0}
        for r in range(r1, r2 + 1):
            tops[r + 1] = tops[r] + self.height(r)
        mid = (ctl.from_off + tops[r2] + ctl.to_off) / 2
        for r in range(r1, r2 + 1):
            if tops[r] <= mid < tops[r + 1]:
                return r
        return r1

    def boxes(self, r: int) -> list[Box]:
        return self.rows.get(r, [])

    def sections(self, r: int, before_col: int) -> list[Box]:
        """その行を含む、縦に結合された見出し（左にあるもの）。結合されていなくても、
        区分の最初の行にだけ書かれ、下の行は同じ列割りで続いている見出しは引き継ぐ。"""
        out = [b for b in self.cover.get(r, []) + self.boxes(r)
               if b.kind == "text" and b.r2 > b.r and b.c2 < before_col and isinstance(b.cell.value, str)]
        texts = [b for b in self.boxes(r) if b.kind == "text" and b.c2 < before_col and isinstance(b.cell.value, str)
                 and b.cell.value.strip() and b.r2 == b.r]
        if len(texts) == 1 and not any(o.c2 < texts[0].c1 for o in out):
            item = texts[0]
            for rr in range(r - 1, max(0, r - 8), -1):
                above = [b for b in self.boxes(rr) if b.kind == "text" and b.c2 < before_col and isinstance(b.cell.value, str)
                         and b.cell.value.strip()]
                same = [b for b in above if (b.c1, b.c2) == (item.c1, item.c2)]
                if not same:
                    break
                lead = [b for b in above if b.c2 < item.c1 and b.r2 == b.r]
                if lead:
                    out.append(lead[-1])
                    break
                if any(b.c2 < item.c1 for b in self.cover.get(rr, [])):
                    break
        return sorted(out, key=lambda b: b.c1)

    # ------------------------------------------------------------------ 見出し
    def phrase(self, r: int, target: Box) -> tuple[list[Box], int]:
        """入力欄の前後に続く「令和［］年［］月［］日」のような、短い語と欄が交互に並ぶところ。"""
        row = [b for b in self.boxes(r) if b.kind != "check" and (not b.outside or target.outside)]
        i = row.index(target)

        def word(b: Box) -> bool:
            return b.kind == "text" and isinstance(b.cell.value, str) and _is_phrase_text(b.cell.value)

        def slot(b: Box) -> bool:
            return b.kind == "input" or (b.kind == "formula" and 0 < len(b.text) <= 6)

        def cell1(b: Box) -> bool:
            return b.kind == "input" and b.c1 == b.c2 and b.r == b.r2

        seq = [target]
        for step in (-1, 1):
            j = i + step
            after_word = False
            prev = target
            while 0 <= j < len(row):
                b = row[j]
                if word(b):
                    after_word = True
                elif slot(b) and (after_word or (cell1(b) and cell1(prev) and abs(b.c1 - prev.c1) == 1)):
                    after_word = False
                else:
                    break
                prev = b
                if step < 0:
                    seq.insert(0, b)
                else:
                    seq.append(b)
                j += step
        # 端の欄は、語でつながっていないので外す
        while len(seq) > 1 and seq[0] is not target and not word(seq[0]):
            seq.pop(0)
        while len(seq) > 1 and seq[-1] is not target and not word(seq[-1]):
            seq.pop()
        return seq, seq.index(target)

    def header(self, target: Box, phrase: list[Box]) -> list[str]:
        """同じ列の上にある見出し（2段まで）。行の中のほかのマスまで覆う見出し（表題など）は使わない。"""
        others = [b for b in self.boxes(target.r) if b not in phrase and b.kind != "check"]
        lo = min(b.c1 for b in phrase) if phrase else target.c1
        hi = max(b.c2 for b in phrase) if phrase else target.c2
        out: list[str] = []
        r = target.r - 1
        budget = 15
        shape = self.shape(target.r)
        span = (target.c1, target.c2)
        while r > 0 and budget > 0 and len(out) < 2:
            if not out and self.shape(r) == shape and shape:
                r -= 1  # 表の本体の行（同じ形）は数えずに上へ
                continue
            budget -= 1
            cand = [b for b in self.boxes(r) + self.cover.get(r, [])
                    if b.c1 <= span[1] and b.c2 >= span[0] and b.r2 >= r]
            if not out and any(b.kind == "input" for b in cand):
                if all((b.c1, b.c2) == (target.c1, target.c2) for b in cand if b.kind in ("input", "formula")):
                    budget += 1  # 同じ列割りの欄が上に続く（表の列）は数えない
                    r -= 1
                    continue
                break  # すぐ上に別の入力欄がある：この欄の見出しは上にはない
            texts = [b for b in cand if b.kind == "text" and isinstance(b.cell.value, str) and b.cell.value.strip()
                     and not _is_phrase_text(b.cell.value)]
            if out and any(b.kind in ("input", "formula") or (b.kind == "text" and not isinstance(b.cell.value, str))
                           for b in cand):
                break  # 親の見出しは、すぐ上に続く見出しだけ（欄や数値をまたがない）
            if not texts:
                r -= 1
                continue
            h = texts[0]
            covered = [o for o in others if o.c1 <= h.c2 and o.c2 >= h.c1 and not (lo <= o.c1 and o.c2 <= hi)
                       and not (o.kind == "formula" and not o.text)
                       and not (o.kind == "text" and isinstance(o.cell.value, str) and _is_phrase_text(o.cell.value))]
            if out:
                # 2段目（親の見出し）は、いくつかの列をまとめていてよい。表題のように広いものは使わない
                if len(covered) > 4 or (h.c2 - h.c1 + 1) > max(12, 3 * (span[1] - span[0] + 1)):
                    break
            elif covered:
                break
            label = norm_label(h.cell.value)
            if label not in out:
                out.insert(0, label)
            span = (h.c1, h.c2)
            r = h.r - 1
        return out

    def describe(self, target: Box, options: list | None = None) -> dict:
        """入力欄（またはチェックボックス）の見出しと、判断の材料。"""
        r = target.r
        row = self.boxes(r)
        phrase, pi = self.phrase(r, target) if target.kind == "input" else ([target], 0)
        first = phrase[0] if phrase else target
        secs = [norm_label(b.cell.value) for b in self.sections(r, first.c1)]
        left_edge = first.c1
        nearest = [b for b in row if b.c2 < left_edge and b.kind != "check"]
        lefts = []
        if nearest:
            b = nearest[-1]
            if (b.kind == "text" and b.r2 == b.r and isinstance(b.cell.value, str) and b.cell.value.strip()
                    and not _is_phrase_text(b.cell.value) and left_edge - b.c2 - 1 <= 4):
                lefts = [b]
        right_edge = phrase[-1].c2 if phrase else target.c2
        rights = [b for b in row if b.kind == "text" and b.c1 > right_edge
                  and isinstance(b.cell.value, str) and b.cell.value.strip()]
        mark = target.kind == "check" or (options is not None and len(options) <= 6
                                          and all(isinstance(o, str) and len(o) <= 3 for o in options))
        item = None
        method = []
        if target.kind == "check":
            carried_right = [b for b in self.cover.get(r, []) if b.kind == "text" and b.c1 > target.c2
                             and isinstance(b.cell.value, str) and b.cell.value.strip()]
            rights = sorted(rights + carried_right, key=lambda b: b.c1)
        if mark and rights:
            # チェック欄・○を付ける欄は、右に書かれた項目が見出し（①などの番号は続く文と合わせる）
            nxt = [b for b in row if b.c1 > target.c2 and b.kind != "check"]
            if (nxt and nxt[0].kind == "text" and not _is_phrase_text(str(nxt[0].cell.value))) or target.kind == "check":
                t = rights[0]
                item = norm_label(t.cell.value)
                if _MARKER.match(item):
                    nxt_text = [b for b in self.boxes(t.r) if b.kind == "text" and b.c1 > t.c2
                                and isinstance(b.cell.value, str) and b.cell.value.strip()]
                    if nxt_text:
                        item = item + norm_label(nxt_text[0].cell.value)
                    elif len(rights) > 1:
                        item = item + norm_label(rights[1].cell.value)
                method.append("右の項目")
        if item is None and lefts:
            item = norm_label(lefts[-1].cell.value)
            method.append("左の見出し")
        if lefts:
            # 同じ行で、見出しのすぐ左に続く見出し（「法人名｜フリガナ」の「法人名」）は区分として読む。
            # あいだに欄がはさまるもの（横に並んだ別の表の見出し）は区分にしない
            more = []
            for b in reversed([b for b in row if b.c2 < lefts[0].c1]):
                if b.kind != "text" or b.r2 != b.r or not isinstance(b.cell.value, str) or not b.cell.value.strip():
                    break
                if not _is_phrase_text(b.cell.value):
                    more.insert(0, b)
            for b in more:
                t = norm_label(b.cell.value)
                if _MARKER.match(t) and item:
                    item = t + item
                elif t not in secs:
                    secs.append(t)
        heads = self.header(target, phrase)
        parts = list(secs)
        sh_ = self.shape(r)
        repeated = bool(sh_) and (self.shape(r - 1) == sh_ or self.shape(r + 1) == sh_)
        if heads and (item is None or repeated):
            # 表の欄（同じ形の行が続く）：行の見出し（左）と列の見出し（上）の両方で呼ぶ。
            # 「見出し｜値」が縦に並ぶだけの行では、上の行の値を見出しと取り違えないよう、上は見ない
            if item is None:
                key = self.row_key(r, first.c1)
                if key and key not in parts:
                    parts.append(key)
            elif item not in parts:
                parts.append(item)
            parts += [h for h in heads if h not in parts]
            method.append("上の見出し")
        elif item:
            parts.append(item)
        phrase_text = None
        unit = None
        n_inputs = sum(b.kind == "input" for b in phrase)
        if len(phrase) > 1 and (n_inputs >= 2 or any(isinstance(b.cell.value, str) and b.cell.value.strip() in _ERA_WORDS
                                                    for b in phrase if b.kind == "text")):
            phrase_text = "".join(
                "［▶］" if b is target else "［　］" if b.kind == "input" else f"〈{b.text}〉" if b.kind == "formula"
                else b.cell.value.strip() for b in phrase)
            singles = [b for b in phrase if b.kind == "input" and b.c1 == b.c2]
            if len(singles) >= 3:
                phrase_text += "（1マスに1文字）"
            method.append("前後の語")
        else:
            after = [b for b in row if b.c1 > target.c2]
            if after and after[0].kind == "text" and isinstance(after[0].cell.value, str):
                unit = _unit_of(after[0].cell.value)
        label = "／".join(dict.fromkeys(p for p in parts if p))
        if phrase_text:
            label = f"{label}：{phrase_text}" if label else phrase_text
        elif unit and label:
            label = f"{label}（{unit}）"
        return {"label": label or None, "label_method": "・".join(method) or None, "unit": unit,
                "context": self.line(r, mark_box=target)}

    def row_key(self, r: int, before_col: int) -> str | None:
        """表の行の名前：左へたどって、欄をとばして最初にある文字の見出し（横に並んだ表でも、その表の行の名前）。"""
        for b in reversed([b for b in self.boxes(r) if b.c2 < before_col and b.kind != "check"]):
            if (b.kind == "text" and isinstance(b.cell.value, str) and len(b.cell.value.strip()) >= 2
                    and not _is_phrase_text(b.cell.value)):
                return norm_label(b.cell.value)
        return None

    def _repeated(self, r: int) -> bool:
        sig = self.signature(r)
        return bool(sig) and (self.signature(r - 1) == sig or self.signature(r + 1) == sig)

    # ------------------------------------------------------------------ 行の文字
    def signature(self, r: int):
        out = []
        for b in self.boxes(r):
            if b.outside:
                continue
            if b.kind == "text":
                out.append(("t", b.c1, b.c2, b.text if not is_number(b.cell.value) else "#"))
            elif b.kind == "formula":
                out.append(("f", b.c1, b.c2, "#" if is_number(b.cell.value) or b.text.strip() in ("", "0") else b.text))
            else:
                out.append((b.kind[0], b.c1, b.c2, None))
        return tuple(out)

    def shape(self, r: int):
        """行の形（マスの位置と種類だけ。文字は見ない）。"""
        return tuple((b.kind, b.c1, b.c2) for b in self.boxes(r) if b.kind != "check")

    def is_blank_row(self, r: int) -> bool:
        return all(not (b.kind == "input" and b.text) and not (b.kind == "check" and b.control.is_checked)
                   for b in self.boxes(r))

    def token(self, b: Box, mark: bool = False) -> str:
        if b.kind == "input":
            inner = b.ref + (f"＝{b.text}" if b.text else "")
            return f"［{'▶' if mark else ''}{inner}］"
        if b.kind == "formula":
            return f"〈{b.cell_ref}＝{b.text}〉"
        if b.kind == "check":
            return f"{b.text}［{'▶' if mark else ''}{b.ref}］"
        if b.blanks:
            # 文字の中の空欄は［　］で示し、どのセルの文字かを添える
            t = b.cell.value
            parts, at = [], 0
            for a, z in b.blanks:
                parts.append(t[at:a])
                parts.append("［▶］" if mark else "［　］")
                at = z
            parts.append(t[at:])
            return f"〔{b.cell_ref}〕" + "".join(parts).replace("\n", " ").strip()
        return b.text.replace("\n", " ").strip()

    def line(self, r: int, mark_box: Box | None = None, include_outside: bool = True) -> str:
        own = [b for b in self.boxes(r) if include_outside or not b.outside]
        first = min((b.c1 for b in own), default=10 ** 6)
        carried = [b for b in self.cover.get(r, []) if b.kind == "text" and b.c2 < first]
        parts = [f"（{norm_label(b.text)}）" for b in sorted(carried, key=lambda b: b.c1)]
        own = [b for b in own if b.kind != "formula" or b.text]
        inside = [self.token(b, b is mark_box) for b in own if not b.outside]
        outside = [self.token(b, b is mark_box) for b in own if b.outside]
        s = " ｜ ".join(p for p in parts + inside if p)
        if outside:
            s += ("　" if s else "") + "〔印刷範囲外〕" + " ｜ ".join(outside)
        return s


# ====================================================================== 入力欄
def _field_kind(wb: Workbook, cell: Cell | None, rules) -> str:
    for v in rules:
        if v.type == "list":
            return "choice"
        if v.type in ("date",):
            return "date"
        if v.type in ("whole", "decimal"):
            return "number"
        if v.type == "time":
            return "time"
    fmt = wb.style(cell.style).num_fmt if cell else "General"
    if fmt == "@":
        return "text"
    if is_date_format(fmt):
        return "date"
    if cell is not None and is_number(cell.value) and not isinstance(cell.value, bool):
        return "number"
    if re.search(r"[0#?]", re.sub(r'"[^"]*"', "", fmt)) and fmt != "General":
        return "number"
    return "text"


KIND_JA = {"text": "文字", "number": "数値", "date": "日付", "time": "時刻", "choice": "選択", "checkbox": "チェックボックス",
           "radio": "オプションボタン（1つ選ぶ）", "select": "ドロップダウン・リスト（1つ選ぶ）",
           "multiselect": "リスト（いくつでも選ぶ）", "spin": "スピンボタン・スクロールバー（数）"}


def _rule_dict(v) -> dict:
    d = {"type": v.type}
    for k, val in (("operator", v.operator), ("formula1", v.formula1), ("formula2", v.formula2),
                   ("prompt_title", v.prompt_title), ("prompt", v.prompt), ("error_title", v.error_title),
                   ("error", v.error)):
        if val:
            d[k] = val
    d["excel_blocks_invalid"] = v.blocks
    return d


class _Options:
    """入力規則の選択肢を、そのセルの位置で計算する（同じ式は1回だけ）。"""

    def __init__(self, wb: Workbook):
        from .calc import Calc

        self.wb = wb
        self.calc = Calc(wb)
        self.cache: dict = {}

    def get(self, sheet: str, v, r: int, c: int) -> tuple[list | None, str | None]:
        from ..formula.parser import parse, unparse
        from .edit import list_options

        f = (v.formula1 or "").strip()
        if not f:
            return None, None
        try:
            key = (id(v), unparse(parse(f.lstrip("=")), r - v.origin[0], c - v.origin[1])) if not f.startswith('"') else (id(v), f)
        except Exception:  # noqa: BLE001
            key = (id(v), r, c)
        if key not in self.cache:
            opts = list_options(self.wb, sheet, f, r, c, v, self.calc)
            dep = None
            if opts is None and "INDIRECT" in f.upper():
                dep = key[1] if isinstance(key[1], str) else f
            self.cache[key] = (opts, dep)
        return self.cache[key]


def fields_for(wb: Workbook, sh: Sheet, lay: Layout, opts: _Options) -> list[dict]:
    out = []
    from .controls import CHOICE_KINDS, VALUE_KINDS, radio_groups

    groups = {id(c): g for g in radio_groups(wb, sh) for c in g.radios}
    done_groups: set[int] = set()
    for r in sorted(lay.rows):
        for b in lay.rows[r]:
            if b.kind == "check" and b.control.kind != "CheckBox":
                f = _control_field(wb, sh, lay, b, groups, done_groups)
                if f is not None:
                    out.append(f)
                continue
            if b.kind == "check":
                ctl = b.control
                d = lay.describe(b)
                f = {"id": f"{sh.name}!{addr(*ctl.link_cell[1:])}" if ctl.link_cell else f"{sh.name}!{ctl.name}",
                     "cell": addr(*ctl.link_cell[1:]) if ctl.link_cell else None, "control": ctl.name,
                     "kind": "checkbox", "value": ctl.is_checked, "at": addr(*ctl.anchor),
                     "write_as": "TRUE（チェックあり）/ FALSE（なし）", **d, "provenance": LAYOUT}
                if ctl.link_cell and ctl.link_cell[0] != sh.name:
                    f["id"] = f"{ctl.link_cell[0]}!{addr(*ctl.link_cell[1:])}"
                out.append(f)
                continue
            if b.kind == "text" and b.blanks:
                t = b.cell.value
                shown = "".join(t[x:y] if k % 2 == 0 else "［　］" for k, (x, y) in enumerate(_split_spans(t, b.blanks)))
                out.append({"id": f"{sh.name}!{b.cell_ref}", "cell": b.cell_ref, "kind": "inline",
                            "value": t, "blanks": len(b.blanks), "label": norm_label(shown),
                            "label_method": "文字の中の空欄", "context": lay.line(r, mark_box=b),
                            "write_as": f"空欄ごとの値のリスト（{len(b.blanks)} 個、例 {['…'] * len(b.blanks)}）か、書き換えたあとの文字全体",
                            "reason": ["文字の中の空欄"], "provenance": LAYOUT})
                continue
            if b.kind != "input":
                continue
            cell = b.cell
            rules = [v for v in sh.validations if v.covers(b.r, b.c1)]
            options = None
            depends = None
            for v in rules:
                if v.type == "list":
                    o, d = opts.get(sh.name, v, b.r, b.c1)
                    if o is not None:
                        options = (options or []) + [x for x in o if x not in (options or [])]
                    depends = depends or d
            d = lay.describe(b, options)
            f = {"id": f"{sh.name}!{b.cell_ref}", "cell": b.cell_ref, "kind": _field_kind(wb, cell, rules),
                 "value": None if cell is None else (cell.value.code if isinstance(cell.value, ExcelError) else cell.value),
                 **d, "reason": lay.inputs[(b.r, b.c1)], "provenance": LAYOUT}
            if b.ref != b.cell_ref:
                f["merged"] = b.ref
            if cell is not None and cell.value not in (None, "") and b.text != to_text(cell.value):
                f["display"] = b.text
            if options is not None:
                f["options"] = options[:500]
                if len(options) > 500:
                    f["options_total"] = len(options)
            if depends:
                f["options_depend_on"] = depends
            if rules:
                f["rules"] = [_rule_dict(v) for v in rules]
            if len([v for v in rules if v.type not in ("any", None)]) > 1:
                f["rules_overlap"] = True  # 入力規則が重なっている（Excelがどれを使うかはファイルから決まらない）
                for d_, v in zip(f["rules"], rules):
                    d_["range"] = v.sqref_text  # その規則のかかる範囲（この欄だけの規則か、広い範囲の規則か）
            if b.outside:
                f["outside_print_area"] = True
            if b.highlighted:
                f["highlighted"] = True  # 条件付き書式で色が付いている（様式が入力を求めている印のことが多い）
            if not d["label"] and not rules:
                f["unlabeled"] = True
            out.append(f)
    return out


def _control_field(wb: Workbook, sh: Sheet, lay: Layout, b: Box, groups: dict, done: set) -> dict | None:
    """チェックボックス以外の、値を書く部品の欄。"""
    from .controls import CHOICE_KINDS, items_of

    ctl = b.control
    d = lay.describe(b)
    if ctl.kind == "Radio":
        g = groups.get(id(ctl))
        if g is None or id(g) in done:
            return None
        done.add(id(g))
        sel = g.selected
        label = (g.box.text.strip() if g.box is not None and g.box.text.strip() else "") or d.get("label") or "選択肢"
        f = {"id": g.id, "cell": addr(*g.link[1:]) if g.link else None, "kind": "radio", "options": g.options,
             "value": g.options[sel - 1] if sel else None, "controls": [x.name for x in g.radios], **d, "label": label,
             "write_as": "選択肢の文字か、何番目か（1から）", "provenance": LAYOUT}
        if not g.consistent:
            f["undecided"] = g.why  # 組の順を規則では決められない：書き込まない
        return f
    if ctl.kind in CHOICE_KINDS:
        items = lay.control_items.get(ctl.shape_id) or items_of(wb, sh, ctl)
        multi = ctl.kind == "List" and (ctl.sel_type or "single") != "single"
        link = ctl.link_cell
        fid = f"{link[0]}!{addr(*link[1:])}" if link and not multi else f"{sh.name}!{ctl.name}"
        value = ([items[i - 1] for i in (ctl.multi_sel or []) if 0 < i <= len(items)] if multi else
                 (items[ctl.sel - 1] if ctl.sel and 0 < ctl.sel <= len(items) else None))
        return {"id": fid, "cell": addr(*link[1:]) if link and not multi else None, "kind": "multiselect" if multi else "select",
                "control": ctl.name, "options": items, "value": value, **d,
                "write_as": "選択肢の文字のリストか、番号（1から）のリスト" if multi else "選択肢の文字か、番号（1から）",
                "provenance": LAYOUT}
    link = ctl.link_cell
    fid = f"{link[0]}!{addr(*link[1:])}" if link else f"{sh.name}!{ctl.name}"
    lo = ctl.vmin if ctl.vmin is not None else 0
    hi = ctl.vmax if ctl.vmax is not None else 100
    return {"id": fid, "cell": addr(*link[1:]) if link else None, "kind": "spin", "control": ctl.name, "min": lo, "max": hi,
            "value": ctl.val, **d, "write_as": f"{lo}〜{hi} の整数", "provenance": LAYOUT}


# ====================================================================== 窓
def build(wb: Workbook) -> dict:
    blob = wb.zip.original
    doc = {
        "todoku": "0.2",
        "kind": "xlsx",
        "source": {
            "file": wb.filename,
            "sha256": hashlib.sha256(blob).hexdigest(),
            "bytes": len(blob),
            "note": "元のファイルが正本。この窓は読むためのもので、書き込みは元のファイルへの差分として適用される。",
        },
        "provenance_legend": {DET: "AIを使わない決定的な処理で、ファイルから取り出した",
                              LAYOUT: "配置から決定的に推定した見出し。推定なので、context（その行の並び）と合わせて読む",
                              "ai": "AIの判断で補った"},
        "notation": {"［L15］": "入力欄", "［L15＝値］": "記入済みの入力欄", "〈AF35＝値〉": "計算結果（数式）",
                     "☐［BA116］ / ☑［BA116］": "チェックボックス（［］はつながっているセル）",
                     "（法人名）": "上の行から続く、縦に結合された見出し", "［▶…］": "いま説明している欄",
                     "〔印刷範囲外〕": "画面には出るが印刷されない"},
        "macros": _macros(wb),
        "names": [
            {"name": d.name, "refers_to": d.text, "scope": d.sheet_scope, "provenance": DET}
            for d in wb.names if not d.name.startswith("_xlnm._FilterDatabase")
        ],
        "sheets": [],
    }
    opts = _Options(wb)
    for sh in wb.sheets:
        if sh.state.startswith("non-worksheet"):
            doc["sheets"].append({"name": sh.name, "state": sh.state, "cells": [], "fields": [], "reading": []})
            continue
        lay = Layout(wb, sh, opts.calc)
        inputs = lay.inputs
        hidden_fmt = set(lay.hidden_by_format)
        cells = []
        for (r, c) in sorted(set(sh.cells) | set(inputs)):
            cell = sh.cells.get((r, c))
            is_input = (r, c) in inputs
            if (cell is None or (cell.value is None and not cell.has_formula)) and not is_input:
                continue
            m = sh.merged_at(r, c)
            item = {"cell": addr(r, c), "provenance": DET}
            if m and (r, c) != (m[0], m[1]):
                if cell is None or (cell.value is None and not cell.has_formula):
                    continue
                # 結合セルの左上以外に残っている値（Excelの画面には出ないが、ファイルにはある）
                item["hidden_by_merge"] = f"{addr(m[0], m[1])}:{addr(m[2], m[3])}"
                is_input = False
            elif m:
                item["merged"] = f"{addr(m[0], m[1])}:{addr(m[2], m[3])}"
            if cell is not None and cell.value is not None:
                v = cell.value
                item["value"] = v.code if isinstance(v, ExcelError) else v
                d = display(wb, cell)
                if d != (str(v) if not is_number(v) else None):
                    item["display"] = d
                if is_number(v) and is_date_format(wb.style(cell.style).num_fmt):
                    item["type"] = "date"
            if cell is not None and cell.formula:
                item["formula"] = "=" + cell.formula
            if is_input:
                item["input"] = True
                item["input_reason"] = inputs[(r, c)]
            if sh.is_hidden(r, c):
                item["hidden"] = True
            elif not sh.in_print_area(r, c):
                item["outside_print_area"] = True
            if (r, c) in hidden_fmt:
                item["hidden_by_format"] = True  # 条件付き書式で、いまは見えない
            rules = [v for v in sh.validations if v.covers(r, c)]
            if rules:
                item["rules"] = [_rule_dict(v) for v in rules]
            if (r, c) in sh.comments:
                item["note"] = sh.comments[(r, c)]
            cells.append(item)
        fields = fields_for(wb, sh, lay, opts)
        undecided = [{"cell": addr(r, c), "column": col, "decide": "agent", "provenance": LAYOUT,
                      "reason": "表の行に入力欄があるが、このセルには入力欄の印（入力規則・枠・ロックの解除）がない。"
                                "書く欄かどうかは、規則では決められない（見出しから判断する）"}
                     for (r, c), col in sorted(undecided_cells(wb, sh, inputs).items())]
        labels = {f["cell"]: f for f in fields if f.get("cell")}
        for item in cells:
            f = labels.get(item["cell"])
            if f is not None and f["kind"] != "checkbox" and f.get("label"):
                item["label"] = f["label"]
        doc["sheets"].append({
            "name": sh.name,
            "state": sh.state,
            "protected": sh.protected,
            "used_range": f"A1:{addr(max(sh.max_row, 1), max(sh.max_col, 1))}",
            "print_area": sh.print_area,
            "hidden_rows": _runs(sorted(sh.hidden_rows)),
            "hidden_cols": _runs(sorted(sh.hidden_cols), cols=True),
            "header_footer": sh.header_footer,
            "fields": fields,
            "undecided": undecided,
            "reading": reading(lay),
            "hidden_by_format": [addr(*k) for k in lay.hidden_by_format],
            "validations": [{"sqref": v.sqref_text, **_rule_dict(v), "provenance": DET} for v in sh.validations],
            "cells": cells,
            "controls": [{"name": c.name, "kind": c.kind, "checked": c.checked or "Unchecked", "link": c.link,
                          "at": addr(*c.anchor) if c.anchor else None, "text": c.text, "provenance": DET}
                         for c in sh.controls],
            "shapes": [{"kind": s.kind, "name": s.name, "text": s.text, "alt": s.alt,
                        "anchor": addr(*s.anchor) if s.anchor else None, "provenance": DET} for s in sh.shapes],
            "notes": [{"cell": addr(*k), "text": t, "provenance": DET} for k, t in sorted(sh.comments.items())],
            "tables": _tables(wb, sh),
        })
    return doc


def _split_spans(text: str, blanks: list[tuple[int, int]]) -> list[tuple[int, int]]:
    """文字を、空欄でないところ・空欄・…の順の位置に分ける。"""
    out, at = [], 0
    for a, z in blanks:
        out.append((at, a))
        out.append((a, z))
        at = z
    out.append((at, len(text)))
    return out


def _runs(xs: list[int], cols: bool = False) -> list[str]:
    from ..formula.parser import num_to_col

    out: list[list[int]] = []
    for x in xs:
        if out and x == out[-1][1] + 1:
            out[-1][1] = x
        else:
            out.append([x, x])
    f = num_to_col if cols else str
    return [f(a) if a == b else f"{f(a)}:{f(b)}" for a, b in out]


def _tables(wb: Workbook, sh: Sheet) -> list[dict]:
    rows = anchors(sh)
    out = []
    for hr, spans, body in detect_tables(wb, sh, rows):
        heads = {a.c1: norm_label(a.cell.value) for a in rows[hr]}
        trs = []
        for br in body:
            byc = {a.c1: a for a in rows[br]}
            trs.append({heads[c1]: (display(wb, byc[c1].cell) if c1 in byc and byc[c1].cell else "") for c1, _ in spans})
        out.append({"header_row": hr, "columns": [heads[c] for c, _ in spans],
                    "rows": trs, "provenance": DET, "method": "列割りが同じ行の連続を表とみなした"})
    return out


def reading(lay: Layout) -> list[dict]:
    """見える行を上から。同じ形の空欄の行（または2〜4行で1組の行）が続くところは1行にまとめる。"""
    out: list[dict] = []
    rows = sorted(lay.rows)
    texts = {r: lay.line(r) for r in rows}
    i = 0
    while i < len(rows):
        r = rows[i]
        if not texts[r]:
            i += 1
            continue
        best = None  # (組の行数, 続く組の数)
        for per in (1, 2, 3, 4):
            group = rows[i:i + per]
            if len(group) < per or group[-1] - group[0] != per - 1:
                break
            sigs = [lay.signature(x) for x in group]
            n = 0
            j = i + per
            while j + per <= len(rows):
                nxt = rows[j:j + per]
                # 上の組から続く結合セル（縦に長い見出しなど）がかかる行は、まとめない
                if (nxt[0] != rows[j - 1] + 1 or nxt[-1] - nxt[0] != per - 1
                        or [lay.signature(x) for x in nxt] != sigs
                        or not all(lay.is_blank_row(x) and all(b.r >= nxt[0] for b in lay.cover.get(x, []))
                                   for x in nxt)):
                    break
                n += 1
                j += per
            if n >= (2 if per == 1 else 1) and (best is None or n * per > best[1] * best[0]):
                best = (per, n)
        if best is None:
            out.append({"row": r, "text": texts[r]})
            i += 1
            continue
        per, n = best
        for x in rows[i:i + per]:
            if texts[x]:
                out.append({"row": x, "text": texts[x]})
        a, b = rows[i + per], rows[i + per * (n + 1) - 1]
        what = f"{r}行目と同じ形の行" if per == 1 else f"{r}〜{rows[i + per - 1]}行目と同じ形の{per}行1組"
        out.append({"rows": f"{a}-{b}", "text": f"{what}が {n} {'行' if per == 1 else '組'}続きます（入力欄はすべて空欄）"})
        i += per * (n + 1)
    return out


# ====================================================================== Markdown
def _macros(wb: Workbook):
    from ..macros import macro_info

    buttons = [{"sheet": sh.name, "control": c.name, "text": c.text, "macro": c.macro}
               for sh in wb.sheets for c in sh.controls if c.macro]
    return macro_info(wb.vba, wb.has_macros, {sh.code_name: sh.name for sh in wb.sheets if sh.code_name}, buttons)


def _md_head(doc: dict) -> list[str]:
    out = [f"# {doc['source']['file'] or '書類'}", ""]
    out.append(f"- 元ファイルの SHA-256: `{doc['source']['sha256']}`（{doc['source']['bytes']:,} バイト）")
    out.append("- この文書は読むための窓です。書き込みは `todoku fill` で元のファイルに差分として適用します。")
    out.append("- 記号: " + "、".join(f"{k} は{v}" for k, v in doc["notation"].items()))
    out.append("- 見出しは配置から推定したものです。迷ったら「読み順」のその行を見てください。")
    from ..macros import summary_line

    line = summary_line(doc["macros"])
    if line:
        out.append("- " + line)
    out.append("")
    return out


def _md_sheet(sh: dict) -> list[str]:
    """1枚のシートの窓（Markdownの行）。"""
    state = sh["state"]
    if state.startswith("non-"):
        return [f"## シート「{sh['name']}」（セルのないシート）", ""]
    if state != "visible":
        n = len(sh["cells"])
        return [f"## シート「{sh['name']}」（非表示）", "",
                f"画面に出ない計算用のシートです（値のあるセル {n:,} 個）。"
                "ここから作られる選択肢は、各入力欄の「選択肢」に展開してあります。すべてのセルはJSONの窓にあります。", ""]
    out = [f"## シート「{sh['name']}」", ""]
    info = []
    if sh.get("print_area"):
        info.append(f"印刷範囲 {sh['print_area'].rsplit('!', 1)[-1].replace('$', '')}")
    if sh.get("hidden_cols"):
        info.append("非表示の列 " + "、".join(sh["hidden_cols"]))
    if sh.get("hidden_rows"):
        info.append("非表示の行 " + "、".join(sh["hidden_rows"]))
    if sh.get("protected"):
        info.append("シート保護あり（入力欄だけ編集できます）")
    if sh.get("hidden_by_format"):
        info.append(f"条件付き書式でいまは見えない文字 {len(sh['hidden_by_format'])} か所（読み順には出していません）")
    if info:
        out += ["- " + "／".join(info), ""]
    fields = sh.get("fields", [])
    main = [f for f in fields if not f.get("unlabeled")]
    if main:
        out += [f"### 入力欄（{len(main)} か所）", "", "| 欄 | 見出し | 種類・選択肢 | いまの値 |", "|---|---|---|---|"]
        for g in _group_fields(main):
            out.append("| " + " | ".join(_md(x) for x in g) + " |")
        out.append("")
    und = sh.get("undecided", [])
    if und:
        by_col: dict[str, list[str]] = {}
        for u in und:
            by_col.setdefault(u["column"] or "（見出しなし）", []).append(u["cell"])
        out += [f"### 入力欄かどうか、規則では決められないセル（{len(und)} か所）", "",
                "表の行に入力欄がありますが、次の列の空のセルには入力欄の印（入力規則・枠・ロックの解除）がありません。"
                "書く欄かどうかは、列の見出しから判断してください（書くときは、ほかの入力欄と同じように書けます）。", ""]
        for col, cells in by_col.items():
            out.append(f"- 列「{_md(col)}」: " + _compact(cells))
        out.append("")
    extra = [f for f in fields if f.get("unlabeled")]
    if extra:
        cells = [f["cell"] for f in extra if f.get("cell")]
        out += [f"- 見出しの見つからない編集できるセル {len(extra)} か所（空欄のままでよい補助のセルのことが多い）: " + _compact(cells), ""]
    out += ["### 読み順（上から）", ""]
    for line in sh.get("reading", []):
        head = f"{line['row']}" if "row" in line else line["rows"].replace("-", "〜")
        out.append(f"- {head}: {line['text']}")
    out.append("")
    notes = sh.get("notes", [])
    shapes = [s for s in sh.get("shapes", []) if s["text"] or s["alt"] or s["kind"] == "picture"]
    if notes or shapes or sh.get("header_footer"):
        out += ["### 図形・メモ・ヘッダー", ""]
        for s in shapes:
            desc = s["text"] or s["alt"] or "（文字も代替テキストもありません）"
            out.append(f"- {s['kind']}（{s['anchor'] or '位置不明'}付近）: {_md(desc)}")
        for n in notes:
            out.append(f"- メモ（{n['cell']}）: {_md(n['text'])}")
        for h in sh.get("header_footer", []):
            out.append(f"- ヘッダー・フッター: {_md(h)}")
        out.append("")
    return out


def to_markdown(doc: dict) -> str:
    from ..macros import markdown as macro_md

    out = _md_head(doc)
    for sh in doc["sheets"]:
        out += _md_sheet(sh)
    out += macro_md(doc.get("macros"))
    return "\n".join(out)


BUDGET = 120_000  # 1つの窓の大きさの目安（文字）。AIが1回で読み切れる量


def markdown_parts(doc: dict, budget: int = BUDGET) -> list[tuple[str, str]]:
    """窓を、AIが1回で読み切れる大きさの部分に分ける。小さければ1つのまま。

    分けるときは、最初の部分を目次にし、シートごと（大きなシートは行の範囲ごと）に分ける。
    どの部分にも、見出しの行（どのシートのどこか）を付ける。
    """
    whole = to_markdown(doc)
    if len(whole) <= budget:
        return [("全体", whole)]
    parts: list[tuple[str, list[str]]] = []
    for sh in doc["sheets"]:
        lines = _md_sheet(sh)
        title = f"シート「{sh['name']}」"
        if sum(len(x) + 1 for x in lines) <= budget:
            parts.append((title, lines))
            continue
        head, chunk, size, k = lines[:2], [], 0, 1
        for line in lines[2:]:
            if size + len(line) + 1 > budget - 2000 and chunk:
                parts.append((f"{title} その{k}", head + [f"（{title}の続き その{k}）", ""] + chunk))
                chunk, size, k = [], 0, k + 1
            chunk.append(line)
            size += len(line) + 1
        if chunk:
            parts.append((f"{title} その{k}" if k > 1 else title, head + ([f"（{title}の続き その{k}）", ""] if k > 1 else []) + chunk))
    from ..macros import markdown as macro_md

    mlines = macro_md(doc.get("macros"))
    chunk, size, k = [], 0, 1
    for line in mlines:
        if size + len(line) + 1 > budget - 2000 and chunk:
            parts.append((f"マクロ その{k}", chunk))
            chunk, size, k = [], 0, k + 1
        chunk.append(line)
        size += len(line) + 1
    if chunk:
        parts.append((f"マクロ その{k}" if k > 1 else "マクロ", chunk))
    index = _md_head(doc) + ["## この窓の目次", "",
                             f"窓が大きいので、{len(parts)} の部分に分けました。必要な部分を順に読んでください。", ""]
    for i, (title, lines) in enumerate(parts, 1):
        sh_name = title.split("」")[0].replace("シート「", "")
        sh = next((x for x in doc["sheets"] if x["name"] == sh_name), None)
        n_fields = len([f for f in (sh or {}).get("fields", []) if not f.get("unlabeled")])
        index.append(f"- 部分 {i}: {title}（{sum(len(x) + 1 for x in lines):,} 文字、入力欄 {n_fields} か所）")
    return [("目次", "\n".join(index))] + [(t, "\n".join(_md_head(doc)[:1] + [""] + lines)) for t, lines in parts]


def _kind_text(f: dict) -> str:
    k = KIND_JA.get(f["kind"], f["kind"])
    if f.get("options") is not None:
        opts = [to_text(o) for o in f["options"]]
        shown = "／".join(opts[:15]) + (f"（ほか{len(opts) - 15}件）" if len(opts) > 15 else "")
        k += f": {shown}"
    elif f.get("options_depend_on"):
        k += f": 選択肢はほかの欄で決まる（{f['options_depend_on']}）"
    for r in f.get("rules", []):
        if r["type"] in ("whole", "decimal", "textLength", "date"):
            name = {"whole": "整数", "decimal": "数値", "textLength": "文字数", "date": "日付"}[r["type"]]
            op = r.get("operator", "between")
            k += f"（{name} {r.get('formula1')}〜{r.get('formula2')}）" if op == "between" else f"（{name} {op} {r.get('formula1')}）"
        if r.get("prompt"):
            k += f"（説明: {r['prompt']}）"
    if f["kind"] == "spin":
        k += f"（{f.get('min')}〜{f.get('max')}）"
    if f.get("undecided"):
        k += f"（規則では決められない：{f['undecided']}）"
    if f.get("unit"):
        k += f"（単位 {f['unit']}）"
    if f.get("highlighted"):
        k += "（色付きの欄）"
    if f.get("rules_overlap"):
        parts = []
        for i, r in enumerate(f.get("rules", []), 1):
            what = f"選択肢 {r.get('formula1')}" if r["type"] == "list" else f"{r['type']} {r.get('operator', '')} {r.get('formula1', '')}"
            rng = r.get("range", "")
            parts.append(f"規則{i}：{what.strip()}（範囲 {rng[:40]}{'…' if len(rng) > 40 else ''}）")
        k += ("（入力規則が重なっています。Excelがどれを使うかはファイルからは決まらないので、規則では決められません。"
              + "／".join(parts) + "）")
    return k


def _value_text(f: dict) -> str:
    if f["kind"] == "checkbox":
        return "☑ あり" if f["value"] else "☐ なし"
    if f["kind"] == "multiselect":
        return "、".join(map(str, f.get("value") or []))
    v = f.get("display", f.get("value"))
    return "" if v is None else to_text(v)


def _group_fields(fields: list[dict]) -> list[tuple[str, str, str, str]]:
    """表の列のように、同じ見出し・同じ種類の欄が縦に続くところ（1〜4行おき）は1行にまとめる。"""
    from ..formula.parser import parse_cell

    out = []
    rest = list(fields)
    while rest:
        f = rest.pop(0)
        if f.get("cell") and f["kind"] != "checkbox":
            r0, c0 = parse_cell(f["cell"])
            same = [g for g in rest if g.get("cell") and g["kind"] == f["kind"] and g.get("label") == f.get("label")
                    and parse_cell(g["cell"])[1] == c0 and parse_cell(g["cell"])[0] > r0]
            run = [f]
            if same:
                step = parse_cell(same[0]["cell"])[0] - r0
                if 1 <= step <= 4:
                    want = r0 + step
                    for g in same:
                        rg = parse_cell(g["cell"])[0]
                        if rg == want:
                            run.append(g)
                            want += step
                        elif rg > want:
                            break
            if len(run) >= 3:
                ids = {id(g) for g in run}
                rest = [g for g in rest if id(g) not in ids]
                filled = sum(1 for g in run if g.get("value") not in (None, ""))
                vals = "空欄" if not filled else f"記入済み {filled}・空欄 {len(run) - filled}"
                outside = sum(1 for g in run if g.get("outside_print_area"))
                note = f"、うち印刷範囲外 {outside} 行" if outside else ""
                step_note = "" if len(run) < 2 or parse_cell(run[1]["cell"])[0] - r0 == 1 else \
                    f"、{parse_cell(run[1]['cell'])[0] - r0} 行おき"
                out.append((f"{run[0]['cell']}〜{run[-1]['cell']}（{len(run)} か所{step_note}{note}）",
                            f.get("label") or "（見出しなし）", _kind_text(f), vals))
                continue
        cell = f.get("cell") or f.get("control")
        if f["kind"] == "checkbox":
            cell = f"{'☑' if f['value'] else '☐'} {f.get('cell') or f.get('control')}"
        if f.get("outside_print_area"):
            cell = f"{cell}（印刷範囲外）"
        out.append((cell, f.get("label") or "（見出しなし）", _kind_text(f), _value_text(f)))
    return out


def _compact(cells: list[str]) -> str:
    if len(cells) <= 12:
        return "、".join(cells)
    return "、".join(cells[:12]) + f" ほか {len(cells) - 12} か所"


def _md(s: str) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")
