"""AIや支援技術が読むための「窓」を作る。元のファイルは正本のまま残る。

窓はJSON（機械向け）とMarkdown（人とAI向け）の2つ。どちらも決定的な処理だけで作り、
すべての項目に「どうやって得たか」（provenance）を付ける。

読み順は、画面に見える行を上から、行の中は左から並べる。結合セルは1マス、入力欄は［］、
計算結果は〈〉、チェックボックスは☐☑で示す。見出しは配置から決めるが、推定であることを明示し、
判断の材料になった行（context）を必ず添える。
"""

from __future__ import annotations

import bisect
import hashlib
import itertools
import re
import unicodedata
from dataclasses import dataclass, field

from ..formula.numfmt import format_value, is_date_format
from ..formula.parser import RefNode, addr, iter_refs
from ..formula.values import ExcelError, Unsupported, is_number, to_text
from .workbook import Cell, Control, Sheet, Workbook

DET = "deterministic"
LAYOUT = "layout"  # 配置から決定的に推定した（正しいとは限らないので context と合わせて読む）
AGENT = "agent"  # 規則では決められない所（欄かどうかを、エージェントが窓を読んで判断する。窓では［…？］）


def norm_label(s: str) -> str:
    """見出しの空白と改行を1つにまとめ、「事 業 所 番 号」のように字間を空けたものは詰める。文字そのものは変えない。"""
    t = str(s).strip()
    parts = re.split(r"[ \u3000\r\n\t]+", t)
    if len(parts) >= 3 and all(len(p) <= 2 for p in parts):
        return "".join(parts)
    return re.sub(r"[ \u3000\r\n\t]+", " ", t)


def display(wb: Workbook, cell: Cell) -> str:
    shown = getattr(wb, "shown_text", None)  # 見える字をファイルに保存している形式（OpenDocument の表計算）は、その字
    if shown is not None:
        t = shown(cell)
        if t is not None:
            return t
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


_MEMO = "_window_memo"  # 窓を作るあいだだけ、シートに置く控え（配置のマス・表）。build の外では使わない


def anchors(sh: Sheet) -> dict[int, list[Anchor]]:
    """行ごとに、表示上の1マス（結合セルは1つ）を左から並べる。

    窓を作るあいだ（build）は、同じシートについて何度も呼ばれる（入力欄・配置・表）ので、1回だけ作って控える。
    窓を作るあいだ、セルと結合は変わらない（控えは build の終わりに捨てる）。"""
    memo = sh.__dict__.get(_MEMO)
    if memo is None:
        return _anchors(sh)
    if "anchors" not in memo:
        memo["anchors"] = _anchors(sh)
    return memo["anchors"]


def _anchors(sh: Sheet) -> dict[int, list[Anchor]]:
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


_SHEET_EDGE = (65536, 256)  # .xls のシートの最後の行・列（.xlsx は 1048576・16384。どちらも、ここより大きい範囲は「列全体・行全体」）


def _used_limits(sh: Sheet) -> tuple[int, int]:
    """シートで使っている範囲（値か書式のあるセル・結合）と印刷範囲を合わせた、いちばん下の行・いちばん右の列。
    列全体・行全体の印刷範囲（$A:$H など）は、列・行を決めるだけで、下・右の端は使っている範囲で決める。"""
    rr, cc = sh.max_row, sh.max_col
    for m in sh.merges:
        rr, cc = max(rr, m[2]), max(cc, m[3])
    for r1, c1, r2, c2 in sh.print_ranges:
        if r2 < _SHEET_EDGE[0]:
            rr = max(rr, r2)
        if c2 < _SHEET_EDGE[1]:
            cc = max(cc, c2)
    return max(rr, 1), max(cc, 1)


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
    lim_r, lim_c = _used_limits(sh)
    for v in sh.validations:
        for r1, c1, r2, c2 in v.sqref:
            # 列全体などの広い範囲は、使っている範囲と印刷範囲（ファイルの作り）までに切る。切った外のセルは、
            # 値も書式もない（Excel の画面では空の所）ので、欄としては出さない（範囲そのものは validations に出す）
            r2, c2 = min(r2, lim_r), min(c2, lim_c)
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    add(r, c, "入力規則があるセル")
    if not sh.protected:
        # 保護も入力規則もない様式：枠で囲まれた空欄・色の付いた空欄で、見出しのそばにあるもの
        for (r, c), why in _boxed_blanks(wb, sh).items():
            add(r, c, why)
    # 数式が参照している定数・空白のセル（計算の入力）。同じセルを何度も見ない（結果は同じ）
    looked: set[tuple[int, int]] = set()
    areas: set[tuple[int, int, int, int]] = set()  # 見た範囲（同じ範囲を指す数式は多い）
    for cell in sh.cells.values():
        refs = list(iter_refs(cell.ast)) if cell.ast is not None else []
        for ref in refs:
            if not isinstance(ref, RefNode):
                continue
            ref = ref.shifted(cell.dr, cell.dc)
            if ref.sheet and ref.sheet != sh.name:
                continue
            # 列全体（A:A）などの広い参照は、使っている範囲と印刷範囲までに切る（その外は、値も書式もないセル）
            box = (ref.r1, ref.c1, min(ref.r2, lim_r), min(ref.c2, lim_c))
            if box in areas:
                continue
            areas.add(box)
            for r in range(box[0], box[2] + 1):
                for c in range(box[1], box[3] + 1):
                    if (r, c) in looked:
                        continue
                    looked.add((r, c))
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
                m = sh.merged_at(br, c1)
                if m is not None and (m[0], m[1]) != (br, c1):
                    continue  # 上や左から続く結合セルの中（見出しなどの一部で、そこだけには書けない）
                out[(br, c1)] = heads.get(c1, "")
    return out


def blank_marks(wb: Workbook, sh: Sheet, a: "Anchor") -> dict | None:
    """空のマスに、ファイルの作りとしてある印：罫線（どの辺か。隣のマスの側に引いた線も数える）・塗りの色・結合。
    印がなければ None。印があっても、入力欄とは決めない（決めるのは input_cells と、エージェント）。"""
    def bits(r: int, c: int) -> int:
        return wb.style(sh.style_of(r, c)).borders if r >= 1 and c >= 1 else 0

    sides = {"左": bits(a.r, a.c1) & 1 or bits(a.r, a.c1 - 1) & 2,
             "右": bits(a.r, a.c2) & 2 or bits(a.r, a.c2 + 1) & 1,
             "上": bits(a.r, a.c1) & 4 or bits(a.r - 1, a.c1) & 8,
             "下": bits(a.r2, a.c1) & 8 or bits(a.r2 + 1, a.c1) & 4}
    edges = [k for k, v in sides.items() if v]
    fill = wb.style(sh.style_of(a.r, a.c1)).fill_color
    fill = fill if fill not in (None, "FFFFFF") else None
    merged = a.c2 > a.c1 or a.r2 > a.r
    # 罫線で囲まれたマス：左右の両方か、上下の両方に線がある（枠）か、下に線がある（下線の上に書く欄）。
    # 左だけ・上だけなどの1辺の線は、隣の表の外枠（表の右・下のマス）なので、印として数えない（根拠には残す）
    boxed = (sides["左"] and sides["右"]) or (sides["上"] and sides["下"]) or sides["下"]
    if not boxed and fill is None and not merged:
        return None
    out: dict = {}
    if edges:
        out["borders"] = edges
    if fill is not None:
        out["fill"] = fill
    if merged:
        out["merged"] = f"{addr(a.r, a.c1)}:{addr(a.r2, a.c2)}"
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
    # 行ごとの、マスの左端（左から並んでいる）と、そこまでのマスの右端のいちばん右。上の行の「列がかかるマス」を、
    # 行のマスを全部見ずに、二分探索で拾う（拾うマスと順は、全部見たときと同じ）
    starts: dict[int, list[int]] = {}
    reach: dict[int, list[int]] = {}
    for rr, items in rows.items():
        starts[rr] = [b.c1 for b in items]
        mx, acc = 0, []
        for b in items:
            mx = max(mx, b.c2)
            acc.append(mx)
        reach[rr] = acc

    def over(rr: int, c1: int, c2: int) -> list[Anchor]:
        """行 rr の、列 c1〜c2 にかかるマス（左から）。"""
        items = rows.get(rr)
        if not items:
            return []
        j = bisect.bisect_right(starts[rr], c2)  # ここより左のマスだけが c1 <= c2
        acc = reach[rr]
        got = []
        while j > 0 and acc[j - 1] >= c1:
            j -= 1
            if items[j].c2 >= c1:
                got.append(items[j])
        got.reverse()
        return got
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
            for j in range(k - 1, -1, -1):
                b = items[j]
                if text_of(b):
                    labeled = a.c1 - b.c2 - 1 <= 4
                    break
                if b.cell is not None and b.cell.value not in (None, ""):
                    break
            if not labeled:
                for rr in range(a.r - 1, max(0, a.r - 4), -1):
                    above = over(rr, a.c1, a.c2)
                    x = covered.get((rr, a.c1))
                    if x is not None and all(x is not b for b in above):  # マスは (行, 列) ごとに1つなので、等しいマスは同じもの
                        above.append(x)
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


def _ident(ch: str) -> bool:
    return ch.isascii() and (ch.isalnum() or ch in ".-/")


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
    chain_space = _chains(toks)
    out = []
    for k, (a, b, t) in enumerate(toks):
        if t[0] in "_＿":
            # 英数字にはさまれた半角の「_」（URL・ファイルの名前・メールアドレスの中の「minji05_00589」）は、名前の一部で空欄ではない
            if t == "_" and 0 < a and b < len(text) and _ident(text[a - 1]) and _ident(text[b]):
                continue
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


def _chains(toks: list[tuple[int, int, str]]) -> dict[int, bool]:
    """同じ幅の空白でつながった並び（字間を空けた見出しか、年月日の並び）の、区切りの空白 → 年月日の並びか。
    1つの空白ならどんな字でも、2つ以上の空白なら字（漢字・かな・年号）だけの並び（「職　　業　　名」「令和　　年　　月　　日」）"""
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
    return chain_space


# ---------------------------------------------------------------- 規則では決められない所（エージェントが判断する）
# 規則の型には合わないが、書き込む所かもしれない所。規則では欄と決めず、窓に「？」をつけて示し、エージェントが読んで判断する。
# 書くと決めたら、ほかの欄と同じように id を指定して書く（証明も同じ）。書かなければ、何も変わらない。
AGENT_WHY = {
    "date_gap": "日付らしい並びの空き（区切りの空白の幅が不揃い・空白が1つだけ）",
    "date_dots": "「・」で区切った日付らしい並びの空き",
    "placeholder": "伏せ字（○○・＊＊＊ など。書き換える所）",
    "after_colon": "「：」で終わる見出しの後ろ",
    "list_item": "箇条の印だけの行の中身",
    "before_unit": "単位だけのマスの、単位の前",
    "date_cell": "「年　月」のように単位で始まるマスの、初めの単位の前",
    "after_label": "見出しだけの短い段落（「住所」「１　農産物名」）の後ろ",
    "after_label_long": "見出しだけの段落の後ろ（1 字か 17 字以上の、長い見出し・表題のこともある段落）",
    "after_label_more": "段落の終わりの後ろ（同じ段落に、ほかの空欄・空きもある）",
    "after_mark": "「〒」「＠」や通貨の記号で終わる段落の、記号の後ろ",
    "gap": "語と語の間の、2つ以上の空白（「議長　　　　氏」「代表者の氏名　　　殿」）",
    "empty_cell": "見出しの見つからない空のマス",
    "next_line": "見出しのあるマスの中の、空の段落",
    # 表計算（Excel・OpenDocument）のマス
    "keyword": "単位・年号などの語の前後の空白（「令和　　年」「第　　号」「金　　円」。語の並びから推定した空欄）",
    "blank_box": "枠・塗り・結合のある空のマスで、入力欄と決める印（入力規則・保護の中のロックの解除・すぐそばの見出し）がないもの",
    "table_row": "表の行に入力欄があるのに、この列のセルには入力欄の印（入力規則・枠・ロックの解除）がないもの",
}
# 単位だけのマスの単位（規則の UNIT_ONLY に、階・坪・棟・戸・室・畳と、㎡ などの単位の記号（CJK 互換用文字の U+3371〜U+33DF）を足す）
_AGENT_UNIT = re.compile(r"[階坪棟戸室畳\u3371-\u33df]|m²|ｍ²|m2|ｍ２")
_PH_RUN = re.compile(r"([○〇●◯＊*×✕△▲◇◆])(?:[ \u3000]?\1)+")
_ELLIPSIS = re.compile(r"[…‥]+")
_LIST_ONLY = re.compile(r"[ \u3000]*(?:[・•●◦]|[①-⑳]|[（(]?[0-9０-９]{1,3}[)）.．]|[（(][ア-ンａ-ｚa-z][)）]|[ア-ン][.．])[ \u3000]*")
_COLON_LABEL = re.compile(r"[ \u3000]*([^。．！？!?：:]{1,24}?)[ \u3000]*[：:]([ \u3000]*)")
_DOTS_ONLY = re.compile(r"[ \u3000]*・[ \u3000]*・[ \u3000]*")
_NEXT_UNIT = {"年": "月", "月": "日", "時": "分"}
_OPTION_NEXT = re.compile(r"[□☐■☑☒○◯〇]|[0-9０-９]{1,2}[.．)）]|[（(][0-9０-９]{1,2}[)）]|[①-⑳]")


def agent_blanks(text: str, found=(), in_cell: bool = False, detail: bool = False) -> list[tuple[int, int, str]]:
    """規則では決められない、書き込む所かもしれない所：(始め, 終わり, 何か（AGENT_WHY の鍵）)。
    found は規則で見つけた空欄（重なる所・すぐ隣の所は出さない）。in_cell は表のマスの中の文字か。"""
    from ..pdf.layout import UNIT_ONLY

    taken = [tuple(x[:2]) for x in found]
    out: list[tuple[int, int, str]] = []

    def add(a: int, b: int, why: str) -> None:
        for x, y in taken + [(o[0], o[1]) for o in out]:
            if (a < y and x < b) or (a == b and x <= a <= y) or (x == y and a <= x <= b):
                return
        out.append((a, b, why))

    stripped = text.strip(" \u3000")
    lead = len(text) - len(text.lstrip(" \u3000"))
    # 伏せ字：同じ記号が2つ以上続く所。続く「……」や、短い区切り（「県」「市」「－」など空白のない3字まで）でつながるものは1つの所
    runs = [(m.start(), m.end()) for m in _PH_RUN.finditer(text)]
    if runs:
        ells = [(m.start(), m.end()) for m in _ELLIPSIS.finditer(text)]
        spans = sorted(runs + [e for e in ells if any(0 <= e[0] - r[1] <= 3 or 0 <= r[0] - e[1] <= 3 for r in runs)])
        merged: list[list[int]] = []
        for a, b in spans:
            gap = text[merged[-1][1]:a] if merged else None
            if merged and gap is not None and len(gap) <= 3 and not any(ch in " \u3000\t" for ch in gap):
                merged[-1][1] = max(merged[-1][1], b)
            else:
                merged.append([a, b])
        for a, b in merged:
            core = text[a:b].replace(" ", "").replace("\u3000", "")
            if b - a >= 8 and len(set(core)) == 1 and text[a:b] == stripped:
                continue  # 同じ記号だけの長い行（区切りの飾り）
            add(a, b, "placeholder")
    # 日付らしい並びの、規則の型に合わない空き（「令和 年　月」の年と月の間・「６年　月」）。字間を空けた見出しの空白は除く
    toks = [(m.start(), m.end(), m[0]) for m in _TOK.finditer(text)]
    chains = _chains(toks)
    for k in range(1, len(toks) - 1):
        a, b, t = toks[k]
        if not t.isspace() or b - a != 1 or chains.get(k) is False:
            continue
        prev, nxt = toks[k - 1][2], toks[k + 1][2]
        if (prev in _ERAS and nxt == "年") or _NEXT_UNIT.get(prev) == nxt:
            add(a, b, "date_gap")
    if _DOTS_ONLY.fullmatch(text):
        d1 = text.index("・")
        d2 = text.index("・", d1 + 1)
        add(0, d1, "date_dots")
        add(d1 + 1, d2, "date_dots")
        add(d2 + 1, len(text), "date_dots")
    # 「：」で終わる短い見出しの後ろ（「住所：」「Facility Name:」）
    m = _COLON_LABEL.fullmatch(text)
    if m and any(unicodedata.category(ch)[0] == "L" for ch in m[1]):
        add(m.start(2), len(text), "after_colon")
    # 箇条の印だけの行（「・」「(1)」）の中身
    if stripped and _LIST_ONLY.fullmatch(text):
        add(len(text.rstrip(" \u3000")), len(text), "list_item")
    # 「〒」「＠」・通貨の記号で終わる段落：記号の後ろに書く（「所在地　〒」「電子メール　　　＠」）
    core_end = len(text.rstrip(" \u3000"))
    if stripped and (stripped[-1] in "〒＠@" or unicodedata.category(stripped[-1]) == "Sc"):
        add(core_end, len(text), "after_mark")
    # 語と語の間の、規則の型に合わない2つ以上の空白（字間を空けた見出し・2字の見出し「氏　　名」は除く）
    if not re.fullmatch(r"[ \u3000]*(\S)[ \u3000]+(\S)[ \u3000]*", text):
        for k in range(1, len(toks) - 1):
            a, b, t = toks[k]
            # 選択肢の区切り（「□新規　　□変更」「１．有　　２．無」）の空白は除く
            if t.isspace() and b - a >= 2 and k not in chains and not _OPTION_NEXT.match(text, b):
                add(a, b, "gap")
    # 見出しだけの段落（「住所」「１　農産物名」）の後ろ：値を見出しの後ろに書き足す書類が多い（表題・節の見出しのこともある）。
    # 字数（以前は 2〜16 字だけ）や、同じ段落にほかの空きがあるか（以前はないときだけ）では落とさず、候補にする。
    # どれに当たるかは detail=True のとき why の鍵で示す（after_label：短い見出しだけ・after_label_long：1 字か 17 字以上・
    # after_label_more：ほかの空き・空欄もある段落）。detail=False（これまでの呼び方）では、どれも after_label
    body = re.sub(r"^(?:[0-9０-９]{1,2}|[（(][0-9０-９]{1,2}[)）]|[①-⑳]|[ア-ン])[ \u3000.．、]+", "", stripped)
    if (body and not re.search(r"[。．!?！？、，,：:]", body)
            and not body.endswith(("殿", "様", "御中")) and any(unicodedata.category(ch)[0] == "L" for ch in body)
            and not re.search(r"[ \u3000]{2,}", body)):
        key = "after_label"
        if detail:
            key = "after_label_more" if (found or out) else "after_label" if 2 <= len(body) <= 16 else "after_label_long"
        add(core_end, len(text), key)
    if in_cell and stripped:
        if UNIT_ONLY.fullmatch(stripped) or _AGENT_UNIT.fullmatch(stripped):
            add(0, lead, "before_unit")  # 「人」「名」「円」「階」「㎡」だけのマス
        elif toks and stripped[0] in _NEXT_UNIT and re.match(r"[年月時][ \u3000]+[月日分]", stripped):
            add(0, lead, "date_cell")  # 「年　月」で始まるマス（数は単位の前に書く）
    return sorted(out)


def _bracketed(text: str, a: int, b: int) -> bool:
    """空白が、かっこの組（「（　　）」「［　］」）の中にそれだけで入っている。字の形だけで決まる空欄。"""
    pairs = {"（": "）", "(": ")", "［": "］", "[": "]", "【": "】", "「": "」"}
    return 0 < a and b < len(text) and pairs.get(text[a - 1]) == text[b] and not text[a:b].strip(" 　")


def cell_blanks(text: str) -> list[tuple[int, int, str | None]]:
    """表計算のマスの文字の中の、書き込む所（の候補）：(始め, 終わり, 何か)。何かが None なら規則で決まる空欄、
    文字なら規則では決められない所（AGENT_WHY の鍵。エージェントが窓を読んで判断する）。

    - 規則（字の形だけで決まる）：下線（＿＿＿、___）・かっこの中の空白だけの所（「（　　）」）
    - 規則では決められない：語（年・月・日・号・円などの単位、年号、「第」「〒」など）の前後の空白（"keyword"）と、
      agent_blanks の見つけ方（語と語の間の空き・伏せ字・「：」の後ろなど）。語の並びから推定した所は、
      空欄とは言い切らず、窓に「？」をつけて示す
    Word・PDF も使う inline_blanks（語から決める空欄）は、そのまま変えない。"""
    found = inline_blanks(text)
    out: list[tuple[int, int, str | None]] = []
    for a, b in found:
        rule = text[a:b][:1] in ("_", "＿") or _bracketed(text, a, b)
        out.append((a, b, None if rule else "keyword"))
    # 見出しだけの短い字（「住所」）の後ろ（after_label）は、表計算のマスでは出さない。表計算では、見出しは1つのマスに書き、
    # 値は別のマスに書く作りで、見出しの後ろまで出すと、すべての見出し・すべての値のマス（表の中の「東京都」など）が候補になり、
    # 窓が読めなくなる。値を書くマスは、入力欄か、印（罫線・塗り・結合）があれば［B5？］として出る。見出しのマスに書き足すときは、
    # 書き換えたあとの文字全体を書く
    out += [x for x in agent_blanks(text, found, True) if x[2] != "after_label"]
    return sorted(out)


def blank_spans(wb: Workbook, sheet: str, r: int, c: int) -> list[tuple[int, int]] | None:
    """窓に出したとおりの、そのマスの文字の中の空欄の位置（規則では決められない所［？］も含む）。窓に出していないマスは None。
    窓は、隣のマスに欄があるかなど配置を見て所を選ぶので、書くときも同じ配置（Layout）から読む（シートごとに1回）。"""
    cache = wb.__dict__.setdefault("_blank_spans", {})
    if sheet not in cache:
        sh = wb.sheet(sheet)
        own = _MEMO not in sh.__dict__
        if own:
            sh.__dict__[_MEMO] = {}
        try:
            lay = Layout(wb, sh)
        finally:
            if own:
                sh.__dict__.pop(_MEMO, None)
        cache[sheet] = {(b.r, b.c1): [(x, y) for x, y, _k in b.blanks]
                        for items in lay.rows.values() for b in items if b.kind == "text" and b.blanks}
    return cache[sheet].get((r, c))


def fill_blanks(text: str, values: list, spans: list[tuple[int, int]] | None = None) -> str:
    """マスの文字の中の空欄に、順に値を入れた文字列。値が None の所は、元の字のまま（書かない）。
    spans は窓に出した空欄（blank_spans）。なければ、文字だけから見つけた所（cell_blanks。規則では決められない所も含む）。"""
    if spans is None:
        spans = [(a, b) for a, b, _k in cell_blanks(text)]
    if len(values) != len(spans):
        raise ValueError(f"空欄は {len(spans)} か所あります（値は {len(values)} 個。規則では決められない所［？］も数えます。"
                         "書かない所は null にしてください）")
    out, at = [], 0
    for (a, b), v in zip(spans, values):
        out.append(text[at:a])
        out.append(to_text(v) if v is not None else text[a:b])
        at = b
    out.append(text[at:])
    return "".join(out)


def detect_tables(wb: Workbook, sh: Sheet, rows: dict[int, list[Anchor]]):
    """見出し行と、同じ列割りで続く行を表として取り出す。窓を作るあいだは、同じ配置について1回だけ調べて控える。"""
    memo = sh.__dict__.get(_MEMO)
    if memo is None:
        return _detect_tables(rows)
    got = memo.get("tables")
    if got is None or got[0] is not rows:
        got = memo["tables"] = (rows, _detect_tables(rows))
    return got[1]


def _detect_tables(rows: dict[int, list[Anchor]]):
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
    kind: str  # text / input / formula / check / agent（規則では欄と決められない空のマス。見出しを決める配置には入れない）
    cell: Cell | None = None
    text: str = ""  # 表示（入力欄・数式は値の表示）
    control: Control | None = None
    outside: bool = False  # 印刷範囲の外（画面だけに出る）
    highlighted: bool = False  # 条件付き書式で色が付いている
    blanks: list = field(default_factory=list)  # 文字の中の空欄（書き込む場所）：(始め, 終わり, None＝規則／AGENT_WHY の鍵)
    why: list = field(default_factory=list)  # kind "agent"（規則では欄と決められない空のマス）：AGENT_WHY の鍵
    evidence: dict = field(default_factory=dict)  # kind "agent"：判断の材料（罫線・塗り・結合・近くの字）

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


# ActiveX の部品（MS-OFORMS）：入力欄として読むもの（文字）と、印で読むもの（チェック・オプション・スピン）
_AX_INPUT = ("TextBox", "ComboBox", "ListBox")
_AX_CHECK = ("CheckBox", "OptionButton", "ToggleButton", "SpinButton", "ScrollBar")


def _ax_kind(ctl: Control) -> str | None:
    ref = getattr(ctl, "ax", None)
    return ref.kind if ref is not None and ref.control is not None else None


def _ax_text(ctl: Control) -> str:
    """ActiveX のテキストボックス・コンボボックス・リストボックスの、いまの文字（改行は ⏎ で示す）。"""
    return ctl.ax.control.value.replace("\r\n", "⏎").replace("\n", "⏎").replace("\r", "⏎")


def _ax_control_info(ctl: Control) -> dict:
    """シートの部品の一覧に添える、ActiveX の部品の中身（見出し・値・組・読めないときの理由）。"""
    ref = getattr(ctl, "ax", None)
    if ref is None:
        return {}
    from ..axfile import value_of

    d: dict = {"activex": ref.label, "persistence": ref.persistence}
    if ref.control is not None:
        d["value"] = value_of(ref)
        if ref.kind == "OptionButton":
            d["group"] = ref.control.get("GroupName") or ""
        if ref.kind in ("SpinButton", "ScrollBar"):
            d["min"], d["max"] = ctl.vmin, ctl.vmax
    elif ref.frame is not None:
        d["frame"] = [{"name": s.name, "kind": s.kind, "caption": s.control.caption if s.control else "",
                       "value": s.control.value if s.control and s.kind in ("TextBox", "ComboBox", "ListBox") else
                       (s.control.checked if s.control and s.kind in ("CheckBox", "OptionButton", "ToggleButton") else None)}
                      for s in ref.frame.sites]
    if ref.error:
        d["unreadable"] = ref.error
    return d


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
                if (ctl.kind in ("CheckBox", "Radio", "Drop", "List", "Spin", "Scroll") or _ax_kind(ctl) in _AX_INPUT + _AX_CHECK) \
                        and ctl.link_cell and ctl.link_cell[0] == sh.name:
                    self.links[(ctl.link_cell[1], ctl.link_cell[2])] = ctl
                if ctl.kind == "GBox" and ctl.fmla_group and s.name == sh.name:
                    from .workbook import _link_cell

                    lc = _link_cell(ctl.fmla_group, s.name)
                    if lc and lc[0] == sh.name:
                        self.links[(lc[1], lc[2])] = ctl
        self.rows: dict[int, list[Box]] = {}
        self.cover: dict[int, list[Box]] = {}  # 上の行から続く結合セル（その行で始まらないもの）
        # 規則では欄と決められない空のマス（kind "agent"）。見出しを決める配置（rows）には入れず、読み順と欄の一覧にだけ出す
        # （入れると、ほかの欄の見出しの探し方が変わる）
        self.agent: dict[int, list[Box]] = {}
        shown_sheet = sh.state not in ("hidden", "veryHidden")  # 非表示のシート（選択肢の元の表など）は、記入する所ではない
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
                    # 空のマス：罫線・塗り・結合（ファイルの作り）があれば、欄かもしれない所としてエージェントに示す。
                    # 保護されたシートのロックされたセル（Excel でも書けない）・部品のつながり先は除く
                    if shown_sheet and (a.r, a.c1) not in self.links and \
                            not (sh.protected and wb.style(sh.style_of(a.r, a.c1)).locked):
                        ev = blank_marks(wb, sh, a)
                        if ev is not None:
                            self.agent.setdefault(a.r, []).append(
                                Box(a.r, a.c1, a.r2, a.c2, "agent", cell, "", outside=outside, why=["blank_box"], evidence=ev))
                    continue
                else:
                    b = Box(a.r, a.c1, a.r2, a.c2, "text", cell, display(wb, cell), outside=outside)
                    if isinstance(cell.value, str) and not (sh.protected and wb.style(cell.style).locked):
                        b.blanks = cell_blanks(cell.value)
                        if not shown_sheet:  # 非表示のシートには、規則では決められない所を出さない（記入する所ではない）
                            b.blanks = [x for x in b.blanks if x[2] is None]
                self.rows.setdefault(a.r, []).append(b)
                for rr in range(a.r + 1, a.r2 + 1):
                    self.cover.setdefault(rr, []).append(b)
        # 入力欄なのに、ファイルにセルの要素がない（値も書式もない）マス：入力規則・数式の参照で決まった欄。
        # 配置のマス（anchors）にないので、ここで入力欄として足す（足さないと、窓の欄の一覧にも読み順にも出ない）
        anchored = {(a.r, a.c1) for items in anchors(sh).values() for a in items}
        for (r, c) in self.inputs:
            if (r, c) in anchored or (r, c) in self.links or sh.is_hidden(r, c):
                continue
            b = Box(r, c, r, c, "input", None, "", outside=not sh.in_print_area(r, c))
            self.rows.setdefault(r, []).append(b)
        under = set()
        from .controls import CHOICE_KINDS, VALUE_KINDS

        self.control_items: dict[str, list[str]] = {}
        for ctl in sh.controls:
            axk = _ax_kind(ctl)
            if getattr(ctl, "ax", None) is not None and axk is None:
                axk = "Frame" if ctl.ax.frame is not None else "?"
            if (ctl.kind not in ("CheckBox", "Radio") + CHOICE_KINDS + VALUE_KINDS and axk not in _AX_INPUT + _AX_CHECK
                    + ("Label", "CommandButton", "Frame", "?")) or ctl.anchor is None:
                continue
            r = self.control_row(ctl)
            c = ctl.anchor[1]
            if sh.is_hidden(r, c):
                continue
            if axk == "Label":
                # ActiveX のラベル：見出しの文字として読む（部品の下のセルは欄から外さない）
                if ctl.text.strip():
                    self.rows.setdefault(r, []).append(Box(r, c, r, c, "text", Cell(sh.name, r, c, value=ctl.text),
                                                           ctl.text, ctl))
                continue
            if axk in ("CommandButton", "Frame", "?"):
                # ボタン・フレーム・読めない ActiveX の部品：読み順の中に、何があるかだけを示す
                from ..axfile import outline

                self.rows.setdefault(r, []).append(Box(r, c, r, c, "button", None, outline(ctl.ax), ctl))
                continue
            if axk in _AX_INPUT:
                # ActiveX のテキストボックス・コンボボックス・リストボックス：入力欄として読む（見出しは左や上から）
                self.rows.setdefault(r, []).append(Box(r, c, r, c, "input", None, _ax_text(ctl), ctl))
            else:
                self.rows.setdefault(r, []).append(Box(r, c, r, c, "check", None, self.control_mark(ctl), ctl))
            (r1, c1), (r2, c2) = ctl.anchor, ctl.anchor_to or ctl.anchor
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    under.add((rr, cc))
        for r, items in self.rows.items():
            # チェックボックスの下に敷かれた空欄（クリックできるようにロックを外してあるもの）は、欄として数えない
            self.rows[r] = [b for b in items if not (b.kind == "input" and b.control is None and not b.text
                                                     and (b.r, b.c1) in under)]
        if sh.cond_rules:
            for r in list(self.rows):
                keep = []
                for b in self.rows[r]:
                    if b.kind in ("text", "formula", "input") and b.control is None and \
                            sh.cond_rules_at(b.r, b.c1):
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
        self._agent_cells(under)
        self._agent_text_spans()

    def _agent_cells(self, under: set) -> None:
        """規則では欄と決められない空のマス（kind "agent"）を整える：部品の下に敷かれたマスを除き、表の行の入力欄の印のない
        マス（undecided_cells）を足し、判断の材料に、近くの字（同じ行の左・同じ列の上）とその離れ方を添える。"""
        sh = self.sh
        for r in list(self.agent):
            self.agent[r] = [b for b in self.agent[r] if (b.r, b.c1) not in under]
        have = {(b.r, b.c1): b for bs in self.agent.values() for b in bs}
        for (r, c), _col in undecided_cells(self.wb, sh, self.inputs).items():
            if (r, c) in under or (r, c) in self.links:
                continue
            b = have.get((r, c))
            if b is None:
                m = sh.merged_at(r, c)
                if m is not None and (m[0], m[1]) != (r, c):
                    continue  # 上や左から続く結合セルの中（見出しなどの一部で、そこだけには書けない）
                r2, c2 = (m[2], m[3]) if m else (r, c)
                b = Box(r, c, r2, c2, "agent", sh.cells.get((r, c)), "", outside=not sh.in_print_area(r, c))
                self.agent.setdefault(r, []).append(b)
                have[(r, c)] = b
            b.why.append("table_row")
        for r in list(self.agent):
            if not self.agent[r]:
                del self.agent[r]
                continue
            self.agent[r].sort(key=lambda b: b.c1)
            for b in self.agent[r]:
                b.evidence.update(self._near_text(b))

    def _near_text(self, b: Box) -> dict:
        """空のマスのいちばん近くの字：同じ行の左（何列はなれているか）と、同じ列の上（何行上か。10行まで見る）。"""
        out: dict = {}

        def txt(x: Box) -> bool:
            return x.kind == "text" and isinstance(x.cell.value, str) and bool(x.cell.value.strip())

        left = [x for x in self.boxes(b.r) + self.cover.get(b.r, []) if x.c2 < b.c1 and x.kind != "check"]
        if left:
            x = max(left, key=lambda x: x.c2)
            if txt(x):
                out["left"] = {"text": norm_label(x.cell.value)[:40], "cols": b.c1 - x.c2 - 1}
        for rr in range(b.r - 1, max(0, b.r - 11), -1):
            above = [x for x in self.boxes(rr) + self.cover.get(rr, []) if x.c1 <= b.c2 and x.c2 >= b.c1 and x.kind != "check"]
            if above:
                x = above[0]
                if txt(x):
                    out["above"] = {"text": norm_label(x.cell.value)[:40], "rows": b.r - x.r2}
                break
        return out

    def _agent_text_spans(self) -> None:
        """文字の中の、規則では決められない所のうち、すぐ隣のマスに欄がある（値はそちらに書く）ものは出さない：
        見出しの後ろ（「住所」「住所：」「〒」）は右のマスが欄・空のマス、単位の前（「人」だけのマス）は左のマスが欄・空のマス。
        """
        right_kinds = ("after_colon", "after_mark")
        left_kinds = ("before_unit", "date_cell")
        for r, items in self.rows.items():
            slots = {(b.c1, "c1"): b for b in items + self.agent.get(r, []) if b.kind in ("input", "agent")}
            slots.update({(b.c2, "c2"): b for b in items + self.agent.get(r, []) if b.kind in ("input", "agent")})
            for b in items:
                if b.kind != "text" or not any(k in right_kinds + left_kinds for _a, _z, k in b.blanks):
                    continue
                keep = []
                for a, z, k in b.blanks:
                    if k in right_kinds and (b.c2 + 1, "c1") in slots:
                        continue
                    if k in left_kinds and (b.c1 - 1, "c2") in slots:
                        continue
                    keep.append((a, z, k))
                b.blanks = keep

    def agent_boxes(self, r: int) -> list[Box]:
        return self.agent.get(r, [])

    def describe_agent(self, b: Box) -> dict:
        """規則では欄と決められない空のマスの、見出しの推定（入力欄と同じ探し方。そのマスを入力欄として置いてみる）。
        置くのは見出しを探すあいだだけで、ほかの欄の見出しは変えない。"""
        row = self.rows.setdefault(b.r, [])
        tmp = Box(b.r, b.c1, b.r2, b.c2, "input", b.cell, "", outside=b.outside)
        row.append(tmp)
        row.sort(key=lambda x: (x.c1, x.kind != "check"))
        try:
            d = self.describe(tmp)
        finally:
            row.remove(tmp)
            if not row:
                del self.rows[b.r]
        d["context"] = self.line(b.r, mark_box=b)
        return d

    def control_mark(self, ctl: Control) -> str:
        """読み順に出す部品の状態：☑☐（チェック）・◉○（オプションボタン）・▼（ドロップダウン）・≡（リスト）・⇕（スピン）。"""
        from .controls import items_of

        if getattr(ctl, "ax", None) is not None:
            from ..axfile import mark

            return mark(ctl.ax)  # ActiveX の部品
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
            if not out:
                # 表の本体の行（同じ形）と、同じ列割りの欄が上に続く行（表の列）は、数えずに上へ。続く所の上まで、まとめて飛ぶ
                r = self._skip_up(r, shape, span)
                if r <= 0:
                    break
            budget -= 1
            cand = self._cand(r, span)
            if not out and any(b.kind == "input" for b in cand):
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
        row = self.boxes(r)
        if r in self.agent:
            row = sorted(row + self.agent[r], key=lambda b: (b.c1, b.kind != "check"))
        for b in row:
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
        """行の形（マスの位置と種類だけ。文字は見ない）。行のマスの並びが同じあいだは、控えを返す。"""
        boxes = self.boxes(r)
        memo = self.__dict__.setdefault("_shapes", {})
        got = memo.get(r)
        if got is not None and got[0] is boxes and got[1] == len(boxes):
            return got[2]
        s = tuple((b.kind, b.c1, b.c2) for b in boxes if b.kind != "check")
        memo[r] = (boxes, len(boxes), s)
        return s

    def _cand(self, r: int, span: tuple[int, int]) -> list:
        """行 r の、列の範囲 span にかかるマス（上から結合で覆うものも含む）。"""
        return [b for b in self.boxes(r) + self.cover.get(r, []) if b.c1 <= span[1] and b.c2 >= span[0] and b.r2 >= r]

    def _skip_up(self, r: int, shape, span: tuple[int, int]) -> int:
        """見出しを探して上へ行くとき、数えずに越える行（target と同じ形の行・同じ列割りの欄だけが続く行）を越えた先の行。
        越える行の続きは、同じ形と列の範囲ごとに控える（行のマスは、窓を作り終えたあとは変わらない）。"""
        memo = self.__dict__.setdefault("_skip", {}).setdefault((shape, span), {})
        path, x = [], r
        while x > 0:
            if x in memo:
                x = memo[x]
                break
            if shape and self.shape(x) == shape:
                path.append(x)
                x -= 1
                continue
            cand = self._cand(x, span)
            if any(b.kind == "input" for b in cand) and \
                    all((b.c1, b.c2) == span for b in cand if b.kind in ("input", "formula")):
                path.append(x)
                x -= 1
                continue
            break
        for p_ in path:
            memo[p_] = x
        return x

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
        if b.kind == "agent":
            return f"［{'▶' if mark else ''}{b.ref}？］"  # 規則では欄と決められない空のマス
        if b.blanks:
            # 文字の中の空欄は［　］で示し、どのセルの文字かを添える。規則では決められない所は［？］
            t = b.cell.value
            parts, at = [], 0
            for a, z, k in b.blanks:
                parts.append(t[at:a])
                if k is None:
                    parts.append("［▶］" if mark else "［　］")
                else:
                    parts.append("［▶？］" if mark else "［？］")
                at = z
            parts.append(t[at:])
            return f"〔{b.cell_ref}〕" + "".join(parts).replace("\n", " ").strip()
        return b.text.replace("\n", " ").strip()

    def line(self, r: int, mark_box: Box | None = None, include_outside: bool = True) -> str:
        own = self.boxes(r)
        if r in self.agent:
            own = sorted(own + self.agent[r], key=lambda b: (b.c1, b.kind != "check"))
        own = [b for b in own if include_outside or not b.outside]
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
           "multiselect": "リスト（いくつでも選ぶ）", "spin": "スピンボタン・スクロールバー（数）",
           "option": "オプションボタン（同じ組から1つ選ぶ）", "combo": "コンボボックス（選ぶか打つ）"}


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
    for r in sorted(set(lay.rows) | set(lay.agent)):
        row = lay.rows.get(r, [])
        if r in lay.agent:
            row = sorted(row + lay.agent[r], key=lambda b: (b.c1, b.kind != "check"))
        for b in row:
            if b.kind == "agent":
                out.append(_agent_field(wb, sh, lay, b))
                continue
            if b.control is not None and getattr(b.control, "ax", None) is not None:
                if b.kind in ("input", "check"):
                    from ..axfile import excel_field

                    out.append(excel_field(wb, sh, lay, b, lay.describe(b)))  # ActiveX の部品の欄
                continue
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
                spans = [(x, y) for x, y, _k in b.blanks]
                shown = "".join(t[x:y] if k % 2 == 0 else ("［　］" if b.blanks[k // 2][2] is None else "［？］")
                                for k, (x, y) in enumerate(_split_spans(t, spans)))
                whys = [k for _x, _y, k in b.blanks if k is not None]
                n = len(b.blanks)
                f = {"id": f"{sh.name}!{b.cell_ref}", "cell": b.cell_ref, "kind": "inline",
                     "value": t, "blanks": n, "label": norm_label(shown),
                     "label_method": "文字の中の空欄", "context": lay.line(r, mark_box=b),
                     "write_as": f"空欄ごとの値のリスト（{n} 個、例 {['…'] * n}。［？］の所も数え、書かない所は null）か、"
                                 "書き換えたあとの文字全体",
                     "spans": [{"n": i, "before": norm_label(t[max(0, x - 8):x]), "after": norm_label(t[y:y + 8]),
                                "rule": k is None, **({"why": AGENT_WHY[k]} if k is not None else {})}
                               for i, (x, y, k) in enumerate(b.blanks, 1)],
                     "reason": ["文字の中の空欄"], "provenance": LAYOUT}
                if whys:
                    # 語の並びから推定した所・規則の型に合わない空き：空欄とは言い切らない（窓では［？］）
                    f["agent"] = True
                    f["why"] = "・".join(dict.fromkeys(AGENT_WHY[k] for k in whys))
                    f["provenance"] = AGENT
                out.append(f)
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


def evidence_text(ev: dict) -> str:
    """判断の材料（blank_marks と近くの字）を、窓に出す短い文に。"""
    parts = []
    if ev.get("borders"):
        parts.append("罫線 " + "".join(ev["borders"]))
    if ev.get("fill"):
        parts.append(f"塗り #{ev['fill']}")
    if ev.get("merged"):
        parts.append(f"結合 {ev['merged']}")
    if ev.get("left"):
        x = ev["left"]
        parts.append(f"左の字「{x['text']}」（{'すぐ左' if not x['cols'] else str(x['cols']) + ' 列あけて左'}）")
    if ev.get("above"):
        x = ev["above"]
        parts.append(f"上の字「{x['text']}」（{x['rows']} 行上）")
    return "・".join(parts) or "印なし"


def _agent_field(wb: Workbook, sh: Sheet, lay: Layout, b: Box) -> dict:
    """規則では欄と決められない空のマスの欄（窓では［B5？］）。エージェントが窓を読んで判断し、書く所ならこの id で書く。"""
    d = lay.describe_agent(b)
    whys = list(dict.fromkeys(b.why))
    f = {"id": f"{sh.name}!{b.cell_ref}", "cell": b.cell_ref, "kind": _field_kind(wb, b.cell, []), "value": None, **d,
         "agent": True, "why": "・".join(AGENT_WHY[k] for k in whys), "why_key": whys,
         "evidence": dict(b.evidence), "evidence_text": evidence_text(b.evidence),
         "reason": [AGENT_WHY[k] for k in whys], "provenance": AGENT}
    if b.ref != b.cell_ref:
        f["merged"] = b.ref
    if b.outside:
        f["outside_print_area"] = True
    return f


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
AGENT_CELLS_RULE = ("規則では入力欄と決められない空のマス。出すのは、ファイルの作りに印のある空のマスだけ"
                    "（罫線で囲まれたもの［左右か上下の両方・下線］・塗り・結合のあるもの、入力欄のある表の行の空のマス）で、印のない空のマスは出さない。"
                    "入力規則のあるマス・保護されたシートでロックを外したマスは入力欄［B5］、保護されたシートのロックされたマス"
                    "（Excel でも書けない）と非表示のシート・行・列は出さない。書く所かは、根拠（罫線・塗り・近くの字）と"
                    "その行を読んで判断し、書くなら入力欄と同じく番地で書く")


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
                              AGENT: "規則では欄と決められない所。欄かどうかは、エージェントが窓（evidence・context）を読んで判断する",
                              "ai": "AIの判断で補った"},
        "notation": {"［L15］": "入力欄", "［L15＝値］": "記入済みの入力欄", "［B5？］": AGENT_CELLS_RULE,
                     "〔B3〕令和［？］年": "文字の中の、規則では決められない空き（語の並びから推定した所など）。"
                                      "書くなら、空欄ごとの値のリストで（書かない所は null）",
                     "〈AF35＝値〉": "計算結果（数式）",
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
        sh.__dict__[_MEMO] = {}
        try:
            doc["sheets"].append(_sheet_doc(wb, sh, opts))
        finally:
            sh.__dict__.pop(_MEMO, None)
    from ..embeds.excel import add_to_window

    add_to_window(wb, doc)  # グラフ（charts）と埋め込み（objects）：読むだけ
    return doc


def _sheet_doc(wb: Workbook, sh: Sheet, opts: "_Options") -> dict:
    """1枚のシートの窓（JSON）。"""
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
    # 規則では入力欄と決められない空のマス（fields にも agent として入れてある）：番地・根拠・推定した見出しの一覧
    undecided = [{"cell": f["cell"], "column": f.get("label") or "", "decide": "agent", "provenance": AGENT,
                  "why": f["why"], "why_key": f["why_key"], "evidence": f["evidence"], "reason": f["why"]}
                 for f in fields if f.get("agent") and f["kind"] != "inline"]
    labels = {f["cell"]: f for f in fields if f.get("cell")}
    for item in cells:
        f = labels.get(item["cell"])
        if f is not None and f["kind"] != "checkbox" and f.get("label"):
            item["label"] = f["label"]
    return {
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
                      "at": addr(*c.anchor) if c.anchor else None, "text": c.text, "provenance": DET,
                      **_ax_control_info(c)}
                     for c in sh.controls],
        "shapes": [{"kind": s.kind, "name": s.name, "text": s.text, "alt": s.alt,
                    "anchor": addr(*s.anchor) if s.anchor else None, "provenance": DET,
                    **({"decorative": True} if getattr(s, "decorative", False) else {})} for s in sh.shapes],
        "notes": [{"cell": addr(*k), "text": t, "provenance": DET} for k, t in sorted(sh.comments.items())],
        "tables": _tables(wb, sh),
    }


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
        # 見出しの行（3つ以上の文字だけのマスの行）と、同じ列割りで続く行を表とみなすのは、配置からの推定（ファイルに
        # 表の印はない）。表かどうか・どの行までが表かは、エージェントが読んで判断する
        out.append({"header_row": hr, "columns": [heads[c] for c, _ in spans],
                    "rows": trs, "provenance": LAYOUT, "decide": "agent",
                    "method": "列割りが同じ行の連続を表とみなした（配置からの推定。ファイルに表の印はない）"})
    return out


def reading(lay: Layout) -> list[dict]:
    """見える行を上から。同じ形の空欄の行（または2〜4行で1組の行）が続くところは1行にまとめる。"""
    out: list[dict] = []
    rows = sorted(set(lay.rows) | set(lay.agent))
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
    from ..embeds.excel import sheet_markdown

    state = sh["state"]
    if state.startswith("non-"):
        what = "グラフのシート" if sh.get("charts") else "セルのないシート"
        return [f"## シート「{sh['name']}」（{what}）", ""] + sheet_markdown(sh)
    if state != "visible":
        n = len(sh["cells"])
        return [f"## シート「{sh['name']}」（非表示）", "",
                f"画面に出ない計算用のシートです（値のあるセル {n:,} 個）。"
                "ここから作られる選択肢は、各入力欄の「選択肢」に展開してあります。すべてのセルはJSONの窓にあります。", ""] + sheet_markdown(sh)
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
    main = [f for f in fields if not f.get("unlabeled") and not f.get("agent")]
    if main:
        out += [f"### 入力欄（{len(main)} か所）", "", "| 欄 | 見出し | 種類・選択肢 | いまの値 |", "|---|---|---|---|"]
        for g in _group_fields(main):
            out.append("| " + " | ".join(_md(x) for x in g) + " |")
        out.append("")
    agent = [f for f in fields if f.get("agent")]
    if agent:
        out += [f"### 規則では決められない所（{len(agent)} か所。読み順では［B5？］・文字の中の［？］）", "",
                "入力欄の印（入力規則・保護の中のロックの解除）がなく、規則では欄と決められない所です。出しているのは、"
                "ファイルの作りに印のある空のマス（罫線で囲まれたもの［左右か上下の両方・下線］・塗り・結合のあるもの、入力欄のある表の行の空のマス）と、文字の中の空きだけです。"
                "書く所かは、根拠とその行（読み順）を読んで判断し、書くなら入力欄と同じく番地で書きます。"
                "近くの字・推定した見出しは、JSON の窓の fields（evidence・label）にあります。", ""]
        texts = [f for f in agent if f["kind"] == "inline"]
        if texts:
            out += ["文字の中の［？］は、空欄ごとの値のリストで書きます（［　］も［？］も数え、書かない所は null）。", "",
                    "| 欄 | 文字（［　］は規則の空欄・［？］は規則では決められない所） | 何か | いまの値 |", "|---|---|---|---|"]
            for f in texts:
                out.append("| " + " | ".join(_md(x) for x in (f["cell"], f.get("label") or "", _kind_text(f), _value_text(f))) + " |")
            out.append("")
        out += _md_agent_cells([f for f in agent if f["kind"] != "inline"])
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
            desc = s["text"] or s["alt"] or ("（飾り）" if s.get("decorative") else
                                              "**説明（代わりの字）がありません**" if s["kind"] == "picture" else "（文字も代替テキストもありません）")
            out.append(f"- {s['kind']}（{s['anchor'] or '位置不明'}付近）: {_md(desc)}")
        for n in notes:
            out.append(f"- メモ（{n['cell']}）: {_md(n['text'])}")
        for h in sh.get("header_footer", []):
            out.append(f"- ヘッダー・フッター: {_md(h)}")
        out.append("")
    return out + sheet_markdown(sh)


def to_markdown(doc: dict) -> str:
    return _to_markdown(doc, [_md_sheet(sh) for sh in doc["sheets"]])


def _to_markdown(doc: dict, sheet_lines: list[list[str]]) -> str:
    from ..macros import markdown as macro_md

    out = _md_head(doc)
    for lines in sheet_lines:
        out += lines
    out += macro_md(doc.get("macros"))
    return "\n".join(out)


BUDGET = 120_000  # 1つの窓の大きさの目安（文字）。AIが1回で読み切れる量


def markdown_parts(doc: dict, budget: int = BUDGET) -> list[tuple[str, str]]:
    """窓を、AIが1回で読み切れる大きさの部分に分ける。小さければ1つのまま。

    分けるときは、最初の部分を目次にし、シートごと（大きなシートは行の範囲ごと）に分ける。
    どの部分にも、見出しの行（どのシートのどこか）を付ける。
    """
    sheet_lines = [_md_sheet(sh) for sh in doc["sheets"]]  # シートの窓は一度だけ作る（全体と部分の両方に使う）
    whole = _to_markdown(doc, sheet_lines)
    if len(whole) <= budget:
        return [("全体", whole)]
    parts: list[tuple[str, list[str]]] = []
    for sh, lines in zip(doc["sheets"], sheet_lines):
        title = f"シート「{sh['name']}」"
        if sum(len(x) + 1 for x in lines) <= budget:
            parts.append((title, lines))
            continue
        chunks = split_lines(lines[2:], budget - 2000, title, lines[:2])
        parts += [(f"{title} その{k}" if len(chunks) > 1 else title, c) for k, c in enumerate(chunks, 1)]
    from ..macros import markdown as macro_md

    mlines = macro_md(doc.get("macros"))
    if mlines:
        chunks = split_lines(mlines, budget - 2000, "マクロ")
        parts += [(f"マクロ その{k}" if len(chunks) > 1 else "マクロ", c) for k, c in enumerate(chunks, 1)]
    index = _md_head(doc) + ["## この窓の目次", "",
                             f"窓が大きいので、{len(parts)} の部分に分けました。必要な部分を順に読んでください。", ""]
    for i, (title, lines) in enumerate(parts, 1):
        sh_name = title.split("」")[0].replace("シート「", "")
        sh = next((x for x in doc["sheets"] if x["name"] == sh_name), None)
        n_fields = len([f for f in (sh or {}).get("fields", []) if not f.get("unlabeled")])
        index.append(f"- 部分 {i}: {title}（{sum(len(x) + 1 for x in lines):,} 文字、入力欄 {n_fields} か所）")
    return [("目次", "\n".join(index))] + [(t, "\n".join(_md_head(doc)[:1] + [""] + lines)) for t, lines in parts]


_TABLE_RULE = re.compile(r"\|(?:\s*:?-+:?\s*\|)+")


def split_lines(lines: list[str], limit: int, title: str, head: list[str] | None = None) -> list[list[str]]:
    """Markdown の行を、limit 文字以内の部分に分ける（決まった手順なので、同じ窓からはいつも同じ分け方になる）。

    要素の中の改行（モジュールのソースなど）も1行ずつに分ける。続きの部分の頭には、どこの続きかの断りと、
    いまの節の見出し（### ・####）を付け、表の途中なら表の見出しの行を、コードの途中なら「```」を付け直して、
    その部分だけで読めるようにする（コードは前の部分の終わりで閉じる）。1行で limit を超える行は、字数で切る。"""
    head = list(head or [])
    flat: list[str] = []
    for x in lines:
        for ln in str(x).split("\n"):
            while len(ln) > limit:
                flat.append(ln[:limit])
                ln = ln[limit:]
            flat.append(ln)
    out: list[list[str]] = []
    cur: list[str] = list(head)
    size = sum(len(x) + 1 for x in cur)
    body = 0  # cur のうち、頭と断りでない行の数
    sec = sub = None  # いまの「### 」「#### 」の見出し
    table: list[str] | None = None  # いまの表の見出しの2行
    fence: str | None = None  # いま開いているコードの「```…」
    code_line = 0  # コードの中の何行目か
    for i, ln in enumerate(flat):
        is_head = ln.startswith("#") and fence is None
        if body and size + len(ln) + 1 > limit:
            if fence is not None:
                cur.append("```")
            out.append(cur)
            k = len(out) + 1
            cur = head + [f"（{title}の続き その{k}）", ""]
            if not is_head:
                cur += [x for x in (sec, sub) if x] + ([""] if sec or sub else [])
                if table is not None and ln.startswith("|"):
                    cur += table
                if fence is not None:
                    cur += [f"（ソースの続き：{code_line + 1} 行目から）", fence]
            size = sum(len(x) + 1 for x in cur)
            body = 0
        cur.append(ln)
        size += len(ln) + 1
        body += 1
        # いまの位置（節・表・コード）を覚える
        if fence is not None:
            if ln.startswith("```"):
                fence = None
            else:
                code_line += 1
            continue
        if ln.startswith("```"):
            fence, code_line = ln, 0
            continue
        if is_head:
            if ln.startswith("#### "):
                sub = ln
            elif ln.startswith("### "):
                sec, sub = ln, None
            else:
                sec = sub = None
            table = None
            continue
        if ln.startswith("|"):
            if table is None and i + 1 < len(flat) and _TABLE_RULE.fullmatch(flat[i + 1].strip()):
                table = [ln, flat[i + 1]]
        else:
            table = None
    if body or not out:
        out.append(cur)
    return out


def _kind_text(f: dict) -> str:
    if f.get("agent"):
        if f["kind"] == "inline":
            return "文字の中の空き：" + "／".join(
                f"{s['n']}「{s['before']}…{s['after']}」" + ("（規則の空欄）" if s["rule"] else f"（{s['why']}）")
                for s in f.get("spans", []))
        return f"{f['why']}（{f.get('evidence_text', '')}）"
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
    if f.get("activex"):
        k += f"（ActiveX の{f['activex']}「{f.get('control')}」" + (f"・組「{f['group']}」" if f.get("group") else "") + "）"
        if f.get("readonly"):
            k += f"（書けない：{f['readonly']}）"
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
    if f["kind"] in ("checkbox", "option") and f.get("activex") and f["value"] is None:
        return "▣ どちらでもない"
    if f["kind"] == "checkbox":
        return "☑ あり" if f["value"] else "☐ なし"
    if f["kind"] == "option":
        return "◉ 選んでいる" if f["value"] else "○ 選んでいない"
    if f["kind"] == "multiselect":
        return "、".join(map(str, f.get("value") or []))
    v = f.get("display", f.get("value"))
    return "" if v is None else to_text(v)


def _group_fields(fields: list[dict]) -> list[tuple[str, str, str, str]]:
    """表の列のように、同じ見出し・同じ種類の欄が縦に続くところ（1〜4行おき）は1行にまとめる。"""
    from ..formula.parser import parse_cell

    # 同じ種類・同じ見出し・同じ列の欄を、元の順のまま組に分けておく（欄ごとに残り全部を見直すと、
    # 欄の多いシートで欄の数の2乗かかる）。組の中で、もう出した欄・まとめた欄は飛ばす
    pos: dict[int, tuple[int, int]] = {}  # 欄の順番 → (行, 列)
    groups: dict[tuple, list[int]] = {}
    for i, g in enumerate(fields):
        if g.get("cell") and g["kind"] != "checkbox":
            pos[i] = parse_cell(g["cell"])
            groups.setdefault((g["kind"], g.get("label"), pos[i][1]), []).append(i)
    head = {k: 0 for k in groups}  # 組の中で、まだ出していない欄の始まり
    used: set[int] = set()  # まとめて出した欄

    def later(key: tuple, i: int, r0: int):
        """組の中で、i より後に残っていて、r0 より下の欄（元の順）。"""
        idx = groups[key]
        while head[key] < len(idx) and idx[head[key]] <= i:
            head[key] += 1
        for j in idx[head[key]:]:
            if j not in used and pos[j][0] > r0:
                yield j

    out = []
    for i, f in enumerate(fields):
        if i in used:
            continue
        if i in pos:
            r0, c0 = pos[i]
            run_idx = [i]
            same = later((f["kind"], f.get("label"), c0), i, r0)
            first = next(same, None)
            if first is not None:
                step = pos[first][0] - r0
                if 1 <= step <= 4:
                    want = r0 + step
                    for j in itertools.chain((first,), same):
                        rg = pos[j][0]
                        if rg == want:
                            run_idx.append(j)
                            want += step
                        elif rg > want:
                            break
            if len(run_idx) >= 3:
                used.update(run_idx)
                run = [fields[j] for j in run_idx]
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


def _md_agent_cells(agent: list[dict]) -> list[str]:
    """規則では欄と決められない空のマスの一覧（Markdown）。1つも落とさずに、短く：同じ行で、同じ印（罫線・塗り・結合の形・
    何か）のマスが続く所は「N2〜X2」にまとめ、同じまとめ方の行が続く所は「5〜40 行」にまとめる。"""
    from ..formula.parser import num_to_col, parse_cell

    if not agent:
        return []
    by_row: dict[int, list[tuple[int, int, tuple]]] = {}
    for f in agent:
        r, c = parse_cell(f["cell"])
        ev = f.get("evidence") or {}
        span = 1
        if ev.get("merged"):
            a, b = ev["merged"].split(":")
            span = parse_cell(b)[1] - parse_cell(a)[1] + 1
        sig = (tuple(ev.get("borders", ())), ev.get("fill"), ev.get("merged") is not None, span,
               tuple(f.get("why_key") or ()), bool(f.get("outside_print_area")))
        by_row.setdefault(r, []).append((c, c + span - 1, sig))
    runs_of: dict[int, tuple] = {}
    for r, cs in by_row.items():
        cs.sort()
        runs: list[list] = []
        for c1, c2, sig in cs:
            if runs and runs[-1][2] == sig and c1 == runs[-1][1] + 1:
                runs[-1][1] = c2
                runs[-1][3] += 1
            else:
                runs.append([c1, c2, sig, 1])
        runs_of[r] = tuple((a, b, s, n) for a, b, s, n in runs)

    def what(sig: tuple) -> str:
        borders, fill, merged, span, whys, outside = sig
        parts = []
        if "table_row" in whys:
            parts.append("表の行")
        if borders:
            parts.append("罫線 " + "".join(borders))
        if fill:
            parts.append(f"塗り #{fill}")
        if merged:
            parts.append(f"結合 {span} 列" if span > 1 else "縦に結合")
        if outside:
            parts.append("印刷範囲外")
        return "・".join(parts) or "印なし"

    out = ["| 行 | 空のマス（印） |", "|---|---|"]
    rows = sorted(runs_of)
    i = 0
    while i < len(rows):
        r = rows[i]
        j = i
        while j + 1 < len(rows) and rows[j + 1] == rows[j] + 1 and runs_of[rows[j + 1]] == runs_of[r]:
            j += 1
        head = str(r) if j == i else f"{r}〜{rows[j]}（同じ形の {j - i + 1} 行）"
        cells = []
        for a, b, sig, n in runs_of[r]:
            at = f"{num_to_col(a)}" if a == b or n == 1 else f"{num_to_col(a)}〜{num_to_col(b)}"
            cells.append(f"{at}（{what(sig)}）")
        out.append(f"| {head} | " + _md("、".join(cells)) + " |")
        i = j + 1
    return out + [""]


def _compact(cells: list[str]) -> str:
    if len(cells) <= 12:
        return "、".join(cells)
    return "、".join(cells[:12]) + f" ほか {len(cells) - 12} か所"


def _md(s: str) -> str:
    return str(s).replace("|", "\\|").replace("\n", " ")
