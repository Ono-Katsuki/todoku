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

from .. import chars
from ..formula.numfmt import format_value, is_date_format
from ..formula.parser import NameNode, RefNode, addr, iter_refs, parse_cell
from ..formula.values import ExcelError, Unsupported, is_number, to_text
from .workbook import Cell, Control, Sheet, Workbook

DET = "deterministic"
LAYOUT = "layout"  # 配置から決定的に推定した（正しいとは限らないので context と合わせて読む）
AGENT = "agent"  # 規則では決められない所（欄かどうかを、エージェントが窓を読んで判断する。窓では［…？］）


def label_text(s) -> str:
    """窓の見出しの字：印刷された字のまま（前後の空白を除き、セルの中の改行だけを空白 1 つにする。字間の空白は詰めない：
    「氏　　名」は「氏　　名」）。探すための、空白を詰めた写しは label_key で別に出す。"""
    return re.sub(r"[ \t]*(?:\r\n|\r|\n)[ \t]*", " ", str(s).strip())


def label_key(label: str) -> str:
    """見出しを探すための鍵（窓の見出しそのものではない）：前後の字がどちらも全角の字（UAX #11 の W・F）の空白の並びは除き
    （「氏　　名」「地 域 区 分」→「氏名」「地域区分」）、ほかの空白の並び（語の区切り）は空白 1 つにする。字数の目安は使わない。"""
    t = str(label).strip()
    return re.sub(r"\s+", lambda m: "" if 0 < m.start() and m.end() < len(t) and chars.is_wide(t[m.start() - 1])
                  and chars.is_wide(t[m.end()]) else " ", t)


def with_label_keys(fields: list[dict]) -> list[dict]:
    """欄に、見出しを探すための鍵（label_key）を添える（鍵が見出しと同じ字なら添えない）。見出しは印刷された字のまま。"""
    for f in fields:
        k = label_key(f["label"]) if f.get("label") else ""
        if k and k != f["label"]:
            f["label_key"] = k
        else:
            f.pop("label_key", None)
    return fields


def display(wb: Workbook, cell: Cell, flag: bool = True) -> str:
    """セルの見えている字（表示形式を通した字）。再現できない表示形式は、推測した字を出さずに、元の値とそう示す
    （flag=False：元の値だけ）。"""
    return shown_text(wb, cell, flag)[0]


def unshown_mark(fmt: str, raw: str) -> str:
    return f"〔表示形式「{fmt}」は再現できません：元の値 {raw}〕"


def shown_text(wb: Workbook, cell: Cell, flag: bool = True) -> tuple[str, str | None]:
    """(見えている字, 再現できなかった理由)。"""
    shown = getattr(wb, "shown_text", None)  # 見える字をファイルに保存している形式（OpenDocument の表計算）は、その字
    if shown is not None:
        t = shown(cell)
        if t is not None:
            return t, None
    v = cell.value
    if v is None:
        return "", None
    if isinstance(v, ExcelError):
        return v.code, None
    fmt = wb.style(cell.style).num_fmt
    try:
        return format_value(v, fmt, getattr(wb, "date1904", False)), None
    except Exception as e:  # noqa: BLE001  再現できない表示形式（Unsupported）・壊れた表示形式
        raw = getattr(e, "shown", None) or to_text(v)  # 地域で見え方が変わる組み込みの形式（LocaleDependent）は ISO 8601 の字
        why = str(e) if isinstance(e, Unsupported) else f"表示形式を読めません（{type(e).__name__}）"
        return (unshown_mark(fmt, raw) if flag else raw), why


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


def _used_limits(sh: Sheet, wb: Workbook | None = None) -> tuple[int, int]:
    """シートで使っている範囲（値か書式のあるセル・結合）と印刷範囲を合わせた、いちばん下の行・いちばん右の列。
    列全体・行全体の印刷範囲（$A:$H など）は、列・行を決めるだけで、下・右の端は使っている範囲で決める。"""
    rr, cc = sh.max_row, sh.max_col
    for m in sh.merges:
        rr, cc = max(rr, m[2]), max(cc, m[3])
    sheet_name = getattr(sh, "name", None)
    for t in (getattr(wb, "tables", None) or {}).values():
        if t.sheet == sheet_name:  # ファイルが宣言するテーブルの範囲も、使っている範囲
            rr, cc = max(rr, t.r2), max(cc, t.c2)
    for r1, c1, r2, c2 in sh.print_ranges:
        if r2 < _SHEET_EDGE[0]:
            rr = max(rr, r2)
        if c2 < _SHEET_EDGE[1]:
            cc = max(cc, c2)
    return max(rr, 1), max(cc, 1)


# 欄の理由の、変わらない鍵（reason_codes）。窓の reason（日本語の文）と同じ順に並べる。採点（grade.py）などは、文でなく鍵で読む
REASON_CODES = {
    "シート保護の中で編集できるセル": "unlocked",  # 保護されたシートで、ロックを外したセル（ECMA-376 §18.8.33 protection の locked）
    "入力規則があるセル": "validation",  # 入力規則（§18.3.1.32 dataValidation）
    "数式が参照している値": "formula_ref",  # 数式が参照している定数・空白のセル
    "見出しのすぐ右の、保護されたシートのロックされたセル": "locked_slot",
    "文字の中の空欄": "inline",
    # 以前の版の理由（いまは出さない）：枠・塗りのある空欄は「？」の所（blank_box）
    "枠のある空欄": "border",
    "色の付いた空欄": "fill",
}


def reason_codes(reasons) -> list[str]:
    """窓の reason（日本語の文）の鍵：REASON_CODES か、規則では決められない所の理由なら AGENT_WHY の鍵。"""
    inv = {v: k for k, v in AGENT_WHY.items()}
    return [REASON_CODES.get(r) or inv.get(r) or "other" for r in (reasons or [])]


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
    lim_r, lim_c = _used_limits(sh, wb)
    for v in sh.validations:
        for r1, c1, r2, c2 in v.sqref:
            # 列全体などの広い範囲は、使っている範囲と印刷範囲（ファイルの作り）までに切る。切った外のセルは、
            # 値も書式もない（Excel の画面では空の所）ので、欄としては出さない（範囲そのものは validations に出す）
            r2, c2 = min(r2, lim_r), min(c2, lim_c)
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    add(r, c, "入力規則があるセル")
    # 保護も入力規則もない様式の、枠で囲まれた空欄・色の付いた空欄は、ここでは入力欄にしない（ファイルが入力欄と宣言していない）。
    # 罫線の辺・塗り・大きさ・近くの字とその離れ方を事実（evidence）として添えて、「？」の所（Layout.agent：blank_box）として示す。
    # 数式が参照している定数・空白のセルも、ここでは入力欄にしない（数式が読むことは、ファイルの入力欄の宣言ではない）。
    # どの数式が読むか（formula_readers）を事実（referenced_by）として添え、空のセルは「？」の所（formula_ref）として示す
    return out


READERS_MAX = 20  # 1 つのセルについて、窓に番地を出す数式の数（窓の大きさを抑えるため。数は n にすべて数える）


def formula_readers(wb: Workbook, sh: Sheet) -> dict[tuple[int, int], dict]:
    """シート sh のセル → それを読む数式（事実）：{"cells": [「シート!番地」…（READERS_MAX まで）], "n": 読む数式の数}。
    ブックのすべてのシートの数式の参照を見る（名前の定義を通した参照も）。列全体（A:A）などの広い参照は、使っている範囲と
    印刷範囲までに切る（その外は、値も書式もないセル）。読まれることは入力欄の宣言ではないので、ここでは欄と決めない。"""
    memo = sh.__dict__.get(_MEMO)
    if memo is not None and "readers" in memo:
        return memo["readers"]
    lim_r, lim_c = _used_limits(sh, wb)
    me = sh.name.lower()
    out: dict[tuple[int, int], dict] = {}

    def boxes(ast, here: str, dr: int, dc: int, depth: int = 0):
        for ref in iter_refs(ast):
            if isinstance(ref, RefNode):
                if (ref.sheet or here).lower() == me:
                    x = ref.shifted(dr, dc)
                    yield min(x.r1, x.r2), min(x.c1, x.c2), max(x.r1, x.r2), max(x.c1, x.c2)
            elif isinstance(ref, NameNode) and depth < 8 and hasattr(wb, "defined_name"):  # 名前の入れ子の深さ（安全のため）
                d = wb.defined_name(ref.name, ref.sheet or here)
                if d is not None and d.ast is not None:
                    yield from boxes(d.ast, d.sheet_scope or here, 0, 0, depth + 1)

    for s_ in wb.sheets:
        for (r0, c0), cell in s_.cells.items():
            if cell.ast is None or not cell.has_formula:
                continue
            who = f"{s_.name}!{addr(r0, c0)}"
            seen: set = set()
            for r1, c1, r2, c2 in boxes(cell.ast, s_.name, cell.dr, cell.dc):
                for r in range(max(r1, 1), min(r2, lim_r) + 1):
                    for c in range(max(c1, 1), min(c2, lim_c) + 1):
                        if (r, c) in seen:
                            continue
                        seen.add((r, c))
                        d = out.setdefault((r, c), {"cells": [], "n": 0})
                        d["n"] += 1
                        if len(d["cells"]) < READERS_MAX:
                            d["cells"].append(who)
    if memo is not None:
        memo["readers"] = out
    return out


def undecided_cells(wb: Workbook, sh: Sheet, inputs: dict) -> dict[tuple[int, int], str]:
    """入力欄かどうかを、規則では決められないセル：表（列割りが同じ行の連続）の行に入力欄があるのに、
    同じ行の、見出しのある列の空のセルに、入力欄の印（入力規則・枠・ロックの解除）がないもの。
    入力欄とは言い切らず、「エージェントが見出しから判断する所」として示す。値 → その列の見出し。"""
    if sh.state in ("hidden", "veryHidden"):
        return {}  # 非表示のシート（選択肢の元の表など）は、記入する所ではない
    rows = anchors(sh)
    out: dict[tuple[int, int], str] = {}
    for t in detect_tables(wb, sh, rows):
        spans, heads = t.spans, t.heads
        for br in t.body:
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


PT_MM = 25.4 / 72  # 1 pt（1/72 インチ）の mm


def _mdw(wb: Workbook) -> tuple[int, str]:
    """列の幅の単位（既定の書体の数字の幅、ピクセル）と、その決め方。ブックごとに 1 度だけ。"""
    got = wb.__dict__.get("_window_mdw")
    if got is None:
        try:
            from .cellcheck import _formats
            from .circle import mdw_for

            fmts = _formats(wb)
            base = fmts[0].font if fmts and fmts[0].font else None
            mdw, known = mdw_for(base.face if base else "", base.size if base else 11.0)
            got = (mdw, "既定の書体の数字の幅 %d ピクセル%s" % (mdw, "" if known else "（書体の名前からの見積もり）"))
        except Exception:  # noqa: BLE001  書式を読めない形式：Excel の既定（Calibri 11pt の数字の幅 7 ピクセル）と仮定
            got = (7, "数字の幅 7 ピクセル（Calibri 11pt）と仮定")
        wb.__dict__["_window_mdw"] = got
    return got


def col_mm(wb: Workbook, sh: Sheet, c: int) -> float:
    """列の幅（mm）。ECMA-376 §18.3.1.13（col の width：字の数 × 数字の幅 ＋ 余白）を 96 dpi のピクセルにして mm に。
    幅のない列は defaultColWidth、それもなければ baseColWidth（§18.3.1.81）。非表示の列は 0。"""
    if c in sh.hidden_cols:
        return 0.0
    from .circle import width_px

    mdw = _mdw(wb)[0]
    w = sh.col_widths.get(c, sh.default_col_width)
    return width_px(w, mdw, getattr(sh, "base_col_width", 8) or 8) * 0.75 * PT_MM


def row_mm(sh: Sheet, r: int) -> float:
    """行の高さ（mm）。row の ht（pt。§18.3.1.73）、なければ sheetFormatPr の defaultRowHeight。非表示の行は 0。"""
    if r in sh.hidden_rows:
        return 0.0
    return float(sh.row_heights.get(r, sh.default_row_height or 15.0)) * PT_MM


def span_mm(wb: Workbook, sh: Sheet, r1: int, c1: int, r2: int, c2: int) -> dict:
    """マスの範囲の大きさ：{"w": 幅 mm, "h": 高さ mm}（小数 1 桁）。"""
    return {"w": round(sum(col_mm(wb, sh, c) for c in range(c1, c2 + 1)), 1),
            "h": round(sum(row_mm(sh, r) for r in range(r1, r2 + 1)), 1)}


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
    out["sides"] = {"left": bool(sides["左"]), "right": bool(sides["右"]), "top": bool(sides["上"]), "bottom": bool(sides["下"])}
    if fill is not None:
        out["fill"] = fill
    if merged:
        out["merged"] = f"{addr(a.r, a.c1)}:{addr(a.r2, a.c2)}"
    out["size_mm"] = span_mm(wb, sh, a.r, a.c1, a.r2, a.c2)
    return out


# ---------------------------------------------------------------- 文字の中の、書き込む所（字の形と並びだけで決める）
# 字の意味（「年」「様」「市」などの語）からは決めない。どの言語の書類にも同じ見つけ方で、決まるのは「どこに・どれだけの
# 空きがあり、その前後にどんな字が印刷されているか」だけ。何を書く所かは、エージェントが前後の字（そのまま示す）から判断する。
# 字の集まりは chars.py（Unicode の資料から作った表）。
#
# - 規則（書類の作りで決まる空欄）：下線の字（一般カテゴリ Pc）の並び・かっこの組（BidiBrackets.txt）の中が空白だけの所
# - 規則では決めない所（「？」）：1 字ぶん以上の空白（一般カテゴリ Zs）の並び（wide_space）・点線（……・......）・
#   ダッシュ（Pd）の並び（点線と同じ長さ：省略の記号「…」1 つ＝3 字以上）・箱の字（□・☐ など）
# 1 字ずつが同じ幅の空白で並ぶ形（「事　　業　　所」「年　　月　　日」）の空白は、別の鍵（spaced）の「？」として示し、
# 並びの字数などの事実（facts）を添える。見出しか書く所かは、エージェントが前後の字から判断する（字数で隠さない）。
#
# 所は Found（(始め, 終わり, 鍵) の組。facts に事実を添える）で返す。組としては 3 つに分けられる（いままでの呼び方のまま）。
# facts の中身（place_facts）：所の字（printed）・字数（len）・幅の目安（width_em：全角＝1 em。chars.width_em）・前後の字
# （before・after：そのまま。CONTEXT_CHARS 字まで）・字が段落の字のすべてか（whole_line）・1 字ずつの並び（chain：字数・字）。
GAP_MIN = 2  # 空白の並びを「空き」とする字数（半角の空白 1 つは語の区切り。全角の空白は 1 つでも空き：wide_space）
SPACED_MIN = 3  # 1 字ずつが同じ空白で並ぶ形の字数の下限（2 字では「同じ幅の空白」が 1 か所しかなく、形の定義にならない）
DASH_MIN = chars.ELLIPSIS_DOTS  # ダッシュの並びを空きとする字数：点線と同じ（省略の記号「…」U+2026 の互換分解の点の数＝3）
_SPACES = "".join(chr(c) for c in range(0x3001) if chr(c).isspace() or chars.is_space(chr(c)))
CONTEXT_CHARS = 40  # facts の前後の字の長さ（窓の大きさを抑えるため。切ったときは before_clipped・after_clipped）


class Found(tuple):
    """所：(始め, 終わり, 鍵) の組。facts に事実（place_facts）を添える。組としては 3 つに分けられ、並べ替えもできる。
    facts は、読むときに初めて作る（所を探すだけの呼び手に、段落の長さぶんの手間をかけない）。"""

    def __new__(cls, a: int, b: int, why, facts: dict | None = None, text: str | None = None, at: tuple | None = None):
        t = super().__new__(cls, (a, b, why))
        t._facts = facts
        t._text = text
        t._at = at or (a, b)
        return t

    @property
    def facts(self) -> dict:
        if self._text is not None:
            extra, self._facts = self._facts or {}, None
            self._facts = {**place_facts(self._text, *self._at), **extra}
            self._text = None
        if self._facts is None:
            self._facts = {}
        return self._facts


def place_facts(text: str, a: int, b: int) -> dict:
    """所 text[a:b] の事実：所の字（printed）・字数（len）・幅の目安（width_em）・前後の字（before・after）・段落の字の
    すべてか（whole_line）・1 字ずつ同じ空白でつないだ並びの中なら、その字数と字（chain）。判断はしない。"""
    sp = "".join(ch for ch in set(text[max(0, a - 1):b + 1]) if chars.is_space(ch))
    core = text[a:b].strip(sp) if sp else text[a:b]
    out: dict = {"printed": text[a:b], "len": b - a, "width_em": chars.width_em(text[a:b]),
                 "before": text[max(0, a - CONTEXT_CHARS):a], "after": text[b:b + CONTEXT_CHARS]}
    if a > CONTEXT_CHARS:
        out["before_clipped"] = True
    if len(text) - b > CONTEXT_CHARS:
        out["after_clipped"] = True
    if b > a and core and text[a:b] == core and not text[:a].strip(_SPACES) and not text[b:].strip(_SPACES):
        out["whole_line"] = True
    if b > a and all(chars.is_space(ch) for ch in text[a:b]):
        cnt, s_, e_ = spaced_chain(text, a, b)
        if cnt >= 2:
            out["chain"] = {"len": cnt, "text": text[s_:e_]}
    return out

BLANK_RULE_KINDS = ("underscore", "bracket")  # 規則の空欄（書類の作りで決まる）


def wide_space(run: str) -> bool:
    """空白の並び run が、字の間に 1 字を書ける空き（1 字ぶん＝1 em 以上の幅）か。字の意味は見ず、空白の字の幅だけで決める：
    - 全角の空白（東アジアの字の幅 UAX #11 が W・F の空白：U+3000 IDEOGRAPHIC SPACE）を含む：U+3000 の幅は、Unicode
      （第 6 章 6.2 の空白の字の表「the width of ideographic (CJK) characters」）と JIS X 4051（全角の空白は 1 字ぶん）で 1 em
      （「第　号」「平成　年　月　日」「郡　市」）
    - 半角の空白などの幅の狭い空白が GAP_MIN 字以上並ぶ（1 つは語の間の空白：欧文で 0.25〜0.33 em。空きにしない）
    字の書類（Excel・Word・テキストなど）では、空白の幅を測れないので、1 em を字で近づける（半角の空白 2 つか、全角の空白 1 つ）。
    PDF の窓の行の空白は、測った空きから作る（1 em 以上の空きだけが、全角の空白か 2 つ以上の半角の空白になる）。PDF は幅を測り、
    字の書類は字で近づけるので、同じ 1 em の決まりだが、まったく同じ結果になるとは限らない（半角の空白 2 つは、書体によっては
    1 em に足りない）。
    段落の頭・終わりの全角の空白 1 つ（字下げ・行末の空白）は、字と字の間の空きではないので、blank_places は空きにしない。"""
    return len(run) >= GAP_MIN or any(chars.is_wide(ch) for ch in run)


def _wide_split(s: str) -> tuple[list[str], list[str]]:
    """s を、1 字ぶん以上の空白の並び（wide_space）で切った字の塊と、その間の空白の並び。"""
    toks, seps, at = [], [], 0
    for a, b, k in chars.char_runs(s):
        if k == "space" and wide_space(s[a:b]):
            toks.append(s[at:a])
            seps.append(s[a:b])
            at = b
    toks.append(s[at:])
    return toks, seps


def label_tail(s: str) -> str:
    """見出しにする、s の終わりの字の塊：1 字ぶん以上の空白の並び（wide_space：窓の空きと同じ 1 em の決まり）で切った最後の塊。
    ただし、1 字ずつを同じ空白でつないだ並び（字間を空けた見出し「氏　名」「事　業　所」）は、まとめて 1 つの塊にする
    （空きで切ると「名」だけになる）。前後の空白は除く。"""
    toks, seps = _wide_split(s.strip(" 　\t\r\n"))
    i = len(toks) - 1
    while i > 0 and len(toks[i]) == 1 and len(toks[i - 1]) == 1 and (i == len(toks) - 1 or seps[i - 1] == seps[i]):
        i -= 1
    return "".join(t + (seps[k] if k < len(seps) else "") for k, t in enumerate(toks[i:], i)).strip()


def label_head(s: str) -> str:
    """見出しにする、s の初めの字の塊（label_tail の逆向き：初めの塊。字間を空けた見出しはまとめる）。"""
    toks, seps = _wide_split(s.strip(" 　\t\r\n"))
    j = 0
    while j < len(toks) - 1 and len(toks[j]) == 1 and len(toks[j + 1]) == 1 and (j == 0 or seps[j - 1] == seps[j]):
        j += 1
    return "".join(t + (seps[k] if k < j else "") for k, t in enumerate(toks[:j + 1])).strip()


def _ident(ch: str) -> bool:
    return ch.isascii() and (ch.isalnum() or ch in ".-/")


def _spaced_heading(text: str) -> bool:
    """字間を空けた見出し（「事　　業　　所」）：前後に空白がなく、1 字ずつが同じ空白の並びで SPACED_MIN 字以上つながる。"""
    runs = [(a, b) for a, b, k in chars.char_runs(text) if k == "space"]
    if not runs or runs[0][0] == 0 or runs[-1][1] == len(text):
        return False
    seps = {text[a:b] for a, b in runs}
    if len(seps) != 1:
        return False
    toks = [text[a:b] for a, b in zip([0] + [y for _x, y in runs], [x for x, _y in runs] + [len(text)])]
    return len(toks) >= SPACED_MIN and all(len(t) == 1 for t in toks)


def _spaced_option(text: str, a: int) -> bool:
    """text[a] の空白が、チェックの箱の字のすぐ後ろの、字間を空けた選択肢の字（「□地　域　移　行　支　援」）の中か：
    前の箱の字から次の箱の字（なければ終わり）までの、前後の空白を除いた並びが、字間を空けた見出しの形。"""
    start = next((i for i in range(a - 1, -1, -1) if text[i] in chars.BOXES or text[i] in chars.MAYBE_CHECKED), None)
    if start is None:
        return False
    end = next((i for i in range(a, len(text)) if text[i] in chars.BOXES or text[i] in chars.MAYBE_CHECKED), len(text))
    seg = text[start + 1:end]
    i, j = 0, len(seg)
    while i < j and chars.is_space(seg[i]):
        i += 1
    while j > i and chars.is_space(seg[j - 1]):
        j -= 1
    return _spaced_heading(seg[i:j])


def spaced_chain(text: str, a: int, b: int) -> tuple[int, int, int]:
    """空白の並び text[a:b] を含む、1 字ずつ（箱の字 ☐ などは数えない）を同じ長さの空白（違いは 1 字まで）でつないだ並び：
    (字数, 始め, 終わり)。空白の前か後ろが 1 字でなければ (0, a, b)。字の意味は見ない（事実を数えるだけ）。
    空白の前後から外へたどるだけで、段落の字をはじめから数え直さない（並びの長さに比例する手間）。"""
    n = len(text)
    sp = chars.is_space

    def tok_left(i: int) -> tuple[int, int] | None:  # i で終わる字の塊
        j = i
        while j > 0 and not sp(text[j - 1]):
            j -= 1
        return (j, i) if j < i else None

    def tok_right(i: int) -> tuple[int, int] | None:  # i で始まる字の塊
        j = i
        while j < n and not sp(text[j]):
            j += 1
        return (i, j) if j > i else None

    def single(t: tuple[int, int]) -> bool:
        return sum(1 for ch in text[t[0]:t[1]] if ch not in chars.BOXES and ch not in chars.MAYBE_CHECKED) == 1

    def boxed(t: tuple[int, int]) -> bool:
        return any(ch in chars.BOXES for ch in text[t[0]:t[1]])

    if not (0 < a < b < n) or sp(text[a - 1]) or sp(text[b]) or not all(sp(ch) for ch in text[a:b]):
        return 0, a, b
    lo, hi = tok_left(a), tok_right(b)
    if lo is None or hi is None or not (single(lo) and single(hi)):
        return 0, a, b
    sep, cnt = b - a, 2
    while True:  # 左へ
        j = lo[0]
        while j > 0 and sp(text[j - 1]):
            j -= 1
        prev = tok_left(j) if j < lo[0] else None
        if prev is None or not single(prev) or abs((lo[0] - j) - sep) > 1 or boxed(lo):
            break
        lo, cnt = prev, cnt + 1
    while True:  # 右へ
        j = hi[1]
        while j < n and sp(text[j]):
            j += 1
        nxt = tok_right(j) if j > hi[1] else None
        if nxt is None or not single(nxt) or abs((j - hi[1]) - sep) > 1 or boxed(nxt):
            break
        hi, cnt = nxt, cnt + 1
    return cnt, lo[0], hi[1]


def chain_head(text: str) -> tuple[int, int] | None:
    """（互換のための名前。いつも None）以前は「年　　月　　日」の形の初めの字の前を、後ろに所が 2 つ以上あるときだけ
    所にしていた（字数による判断）。いまは所にしない：並びの事実は、字の間の所の facts（chain）に出る。"""
    return None


def heading_gap(text: str, a: int, b: int) -> tuple[int, int] | None:
    """（互換のための名前。いつも None）以前は、1 字ずつを同じ空白でつないだ 4 字以上の並びの字の間を「字間を空けた見出し」
    として所にせず隠していた（字数による判断）。いまは隠さない：字の間は鍵 spaced の「？」の所になり、並びの字数は facts の
    chain（spaced_chain）に出る。見出しか書く所かは、エージェントが判断する。"""
    return None


SPACED_MARK = "〔字間〕"  # （互換のための名前）以前、字間を空けた見出しの並びの後ろに付けた印。いまは付けない（heading_gap）
WEAK_WHY = {"narrow": "字 1 字より狭い空き", "decor": "点線・ダッシュ・同じ記号だけの長い行",
            "one_space": "全角の空白 1 つちょうどの空き"}


def weak_space(text: str, a: int, b: int) -> bool:
    """（互換のための名前。いつも False）以前は、全角の空白 1 つちょうどの空きを、並びの字数で「弱い候補」として順を下げていた。
    いまは順を下げない：空きの幅は facts（width_em・printed）と窓の印「？　」に出る。"""
    return False


def weak_line(fields: list[dict]) -> str:
    """窓の欄の一覧の、弱い候補（消さずに順を下げた「？」の所）をまとめた 1 行。読み順の中には、ほかの所と同じ印で出ている。"""
    groups: dict[str, list[str]] = {}
    for f in fields:
        alt = f.get("alternative_to")
        groups.setdefault(f.get("weak") or "", []).append(f["id"] + (f"（{alt} の代わりの所）" if alt else ""))
    body = "／".join(f"{WEAK_WHY.get(k, k)}：" + "・".join(ids) for k, ids in groups.items())
    return (f"- 弱い候補 {len(fields)} か所（規則では決められない所のうち、形と大きさから書く所の見込みが低いもの。消してはいない："
            f"読み順の中に同じ印で出ていて、書くならほかの所と同じように id で書く）：{body}")


def spaced_ends(text: str) -> set[int]:
    """字間を空けた見出しの形の並び（heading_gap）の終わりの位置：窓で、そこに「〔字間〕」の印を付ける。"""
    return {hg[1] for a, b, k in blank_places(text) if k == "spaced" for hg in [heading_gap(text, a, b)] if hg}


def with_spaced_marks(text: str, ends: set[int] | None = None) -> str:
    """text の、字間を空けた見出しの並びの後ろに「〔字間〕」を入れた字（字の間の空白は元のまま）。"""
    ends = spaced_ends(text) if ends is None else ends
    if not ends:
        return text
    out, at = [], 0
    for e in sorted(ends):
        out.append(text[at:e] + SPACED_MARK)
        at = e
    return "".join(out) + text[at:]


SPACED_NOTE = ("直前の字の並びは、1 字ずつを同じ空白でつないだ 4 字以上の並び（字間を空けた見出しの形）。字の間の空白は元のまま"
               "出し、書く所の id にはしていない（書く所と判断したら、字を目印にして書ける）")


def blank_places(text: str) -> list[tuple[int, int, str]]:
    """文字の中の、書き込む所（の候補）：(始め, 終わり, 種類)。種類は BLANK_KINDS の鍵。
    "underscore"・"bracket" は規則の空欄、ほかは規則では決めない所（エージェントが前後の字を読んで判断する）。"""
    out: list[tuple[int, int, str]] = []
    text = str(text)
    spaced = _spaced_heading(text)
    for a, b, kind in chars.char_runs(text):
        run = text[a:b]
        if kind == "connector":
            # 英数字にはさまれた 1 つの「_」（URL・ファイルの名前・メールアドレスの「minji05_00589」）は、名前の一部
            if b - a == 1 and 0 < a and b < len(text) and _ident(text[a - 1]) and _ident(text[b]):
                continue
            out.append(Found(a, b, "underscore", text=text))
        elif kind == "space":
            opener = text[a - 1] if a > 0 else ""
            if opener and b < len(text) and chars.closing(opener) == text[b]:
                out.append(Found(a, b, "bracket", text=text))  # 「（　　）」「[ ]」「( )」：かっこの組の中が空白だけ
            elif wide_space(run) and (b - a >= GAP_MIN or 0 < a and b < len(text)) and b - a < len(text) and not (
                    b < len(text) and (text[b] in chars.BOXES or text[b] in chars.MAYBE_CHECKED)):
                # チェックの箱の字の前の空白（「☐ 新規　　☐ 変更」）は、箱の字の項目どうしの区切り。
                # 1 字ずつが同じ空白で並ぶ形（「事　業　所」と「年　　月　　日」は形が同じ）は、別の鍵（spaced）の「？」で示す
                out.append(Found(a, b, "spaced" if spaced or _spaced_option(text, a) else "gap", text=text))
        elif kind in ("leader", "dash"):
            long_enough = chars.dot_run(text, a, b) if kind == "leader" else b - a >= DASH_MIN
            if not long_enough:
                continue
            # 点線・ダッシュの並び（段落の字のすべてのものは、切り取り線のことも書く線のこともある：facts の whole_line）
            out.append(Found(a, b, kind, text=text))
    return out


def inline_blanks(text: str) -> list[tuple[int, int]]:
    """文字の中の、書き込む所の位置（blank_places の、規則の空欄と規則では決めない所の両方。チェックの箱の字は含めない）。"""
    return [(a, b) for a, b, _k in blank_places(text)]


def blank_token(text: str, a: int, b: int, why, mark: bool = False) -> str:
    """窓で、文字の中の所 text[a:b] の代わりに出す字：規則の空欄は［　］、規則では決めない所は［？］（mark：書く所に ▶）。
    チェックの箱の字（☐ □ など）は、字を消さずに、字のあとに［　］をつける（「☐［　］Male」）。"""
    m = "▶" if mark else ""
    box = text[a:b] if b - a == 1 and text[a:b] in chars.BOXES else ""
    if why is None:
        return f"{box}［{m or '　'}］"
    return f"{box}［{m}？{shown_gap(text[a:b])}］"


def shown_gap(covered: str) -> str:
    """「？」の所の印の中に、その所の元の空白の字をそのまま入れる（「［P1.1？　］」は全角の空白 1 つ、「［P1.1？  ］」は半角の
    空白 2 つの空き）。空きの大きさ（1 字ぶんちょうどか、もっと広いか）を、エージェントが印から読めるようにする。空白だけの所
    （Unicode の Zs）のときだけで、ほかの字（点線・ダッシュなど）の所は ""。"""
    return covered if covered and all(unicodedata.category(ch) == "Zs" for ch in covered) else ""


def one_wide_space(covered: str) -> bool:
    """空きが全角の空白 1 つだけ（1 字ぶんちょうど。字下げ・語の区切りのこともある、いちばん狭い空き）か。"""
    return covered == "　"


GAP_TOKEN_NOTE = ("「？」の所が空白だけの空きなら、その元の空白を印の中にそのまま入れる（「？　」は全角の空白 1 つ、「？　　」は"
                  "2 つ、「？  」は半角の空白 2 つの空き）。空きの大きさを比べる目安")

ONE_WIDE_SPACE_WHY ="空きは 1 つの全角空白（1 字ぶんちょうど。語の区切りのこともある、いちばん狭い空き）"


def space_width(text: str, a: int, b: int) -> float:
    """text[a:b] の幅（全角の字（UAX #11 の W・F）を 1、ほかを 0.5 とした、字の幅のおよその数。窓に示すだけで、判断には使わない）。"""
    return sum(1.0 if chars.is_wide(ch) else 0.5 for ch in text[a:b])


# ---------------------------------------------------------------- 規則では決められない所（エージェントが判断する）
# 規則の型には合わないが、書き込む所かもしれない所。規則では欄と決めず、窓に「？」をつけて示し、エージェントが読んで判断する。
# 書くと決めたら、ほかの欄と同じように id を指定して書く（証明も同じ）。書かなければ、何も変わらない。
AGENT_WHY = {
    "gap": "1 字ぶん以上の空白の並び（全角の空白 1 つ以上か、半角の空白 2 つ以上。字の意味からは決めていません。"
           "空きの幅と前後の字から判断してください）",
    "leader": "点線（……・......）の並び",
    "dash": "ダッシュ（---・－－－）の 3 字以上の並び",
    "decor": "点線・ダッシュ・同じ記号だけの行",
    "spaced": "1 字ずつを同じ空白でつないだ並びの字の間（字間を空けた見出し「事　業　所」の形とも、「年　　月　　日」"
              "「第　　号」の書く所の形とも同じ。並びの字数は facts の chain）",
    "placeholder": "同じ記号の並び（○○・＊＊＊・×× など。伏せ字として書き換える所のことがある）",
    "after_colon": "コロン（:・： など）で終わる字の後ろ",
    "after_label": "文の区切りの字（。・,・、 など）のない段落の後ろ（段落の字数は facts の body_len）",
    "after_label_long": "文の区切りの字のない段落の後ろ（段落の字数は facts の body_len）",
    "after_label_more": "段落の終わりの後ろ（同じ段落に、ほかの空欄・空きもある）",
    "empty_cell": "見出しの見つからない空のマス",
    "next_line": "見出しのあるマスの中の、空の段落",
    # 表計算（Excel・OpenDocument）のマス
    "blank_box": "枠・塗り・結合のある空のマスで、ファイルが入力欄と宣言する印（入力規則・保護の中のロックの解除）がないもの"
                 "（罫線の辺・塗り・大きさ・近くの字とその離れ方は evidence）",
    "formula_ref": "数式が読むセルで、ファイルが入力欄と宣言する印（入力規則・保護の中のロックの解除）がないもの"
                   "（読む数式の番地は evidence の referenced_by。計算の入力か、計算に使う決まった値かは、ファイルからは決まらない）",
    "table_row": "表の行に入力欄があるのに、この列のセルには入力欄の印（入力規則・枠・ロックの解除）がないもの",
    "bullet_box": "印刷された箱の字（□・☐・❏ など。記入の箱か、箇条・見出しの印かは、字からは決まらない。字の Unicode の名前は facts）",
    "checked_box": "印刷された、印のついた箱の字（☑・☒。記入済みの箱か、飾りの字かは、字からは決まらない）",
    "chain_head": "1 字ずつを同じ空白でつないだ並びの、初めの字の前",
}


def _repeat_runs(text: str) -> list[tuple[int, int]]:
    """同じ記号（一般カテゴリ S・「*」「＊」）が 2 つ以上続く所（間に 1 つの空白をはさんでもよい）。
    チェックの箱の字・点線・ダッシュは、それぞれの見つけ方で示すので含めない。"""
    def sym(ch: str) -> bool:
        return (unicodedata.category(ch)[0] == "S" or ch in "*＊") and ch not in chars.BOXES and ch not in chars.MAYBE_CHECKED

    out, i, n = [], 0, len(text)
    while i < n:
        ch = text[i]
        if not sym(ch):
            i += 1
            continue
        j, count = i + 1, 1
        while j < n:
            if text[j] == ch:
                j, count = j + 1, count + 1
            elif chars.is_space(text[j]) and j + 1 < n and text[j + 1] == ch:
                j, count = j + 2, count + 1
            else:
                break
        if count >= 2:
            out.append((i, j))
        i = j
    return out


def _lead_marker(body: str) -> str:
    """段落の頭の番号・記号（「1.」「(2)」「①」「ア」「a)」）と、そのあとの空白を除いた字。字の種類だけで見る。"""
    n = len(body)
    i = 1 if n and chars.closing(body[0]) else 0
    j = i
    while j < n and unicodedata.category(body[j]) in ("Nd", "No", "Nl"):
        j += 1
    circled = j > i and all(unicodedata.category(ch) in ("No", "Nl") for ch in body[i:j])  # 「①」「Ⅳ」：字そのものが番号
    if j == i and j < n and unicodedata.category(body[j])[0] == "L":
        j += 1  # 1 字の字の印（「ア　」「a)」）
    if j == i:
        return body
    sep = j < n and (unicodedata.category(body[j]) == "Pe" or chars.is_terminal_punctuation(body[j]))
    if sep:
        j += 1
    k = j
    while k < n and chars.is_space(body[k]):
        k += 1
    if not (sep or k > j or circled):
        return body  # 番号・字のあとに区切りも空白もない（「100円」「アイス」）：番号ではない
    return body[k:]


def gap_interior(text: str, a: int, b: int, pad: tuple[bool, bool] | None = None) -> tuple[int, int, tuple[bool, bool]]:
    """空白だけの空き text[a:b] に書くとき、書いた字が前後の印刷された字にくっつかないよう（「氏　名山田　花子」）、元の空白を
    端に 1 字ずつ残した中の範囲：(始め, 終わり, (前を残したか, 後ろを残したか))。字数の目安は使わない：
    - 前に印刷された字（空白・開きかっこ Ps・Pi でない字）があれば初めの 1 字を、後ろに印刷された字（空白・閉じかっこ Pe・Pf
      でない字）があれば終わりの 1 字を残す。残したあとに書く所が 1 字ぶん以上残らないときは、残さない（「年　　月」は空き全体）
    - pad（(前, 後ろ)）で、残すかどうかを決めて渡せる（残せない長さのときは残さない）
    空白でない空き（下線の字・点線）は、そのまま返す。どちらを残したかは、所の kept_edges に出る。"""
    seg = text[a:b]
    if not seg or not all(chars.is_space(ch) for ch in seg):
        return a, b, (False, False)
    if pad is None:
        lead = a > 0 and not chars.is_space(text[a - 1]) and unicodedata.category(text[a - 1]) not in ("Ps", "Pi")
        trail = b < len(text) and not chars.is_space(text[b]) and unicodedata.category(text[b]) not in ("Pe", "Pf")
    else:
        lead, trail = bool(pad[0]), bool(pad[1])
    if b - a - lead - trail < 1:
        return a, b, (False, False)
    return a + lead, b - trail, (lead, trail)


def _spaced_pair(core: str, a: int, b: int) -> bool:
    """core 全体が、2 字を空白でつないだだけの見出し（「住　所」「氏　名」「電　話」）で、core[a:b] がその字の間か。
    字間を空けた 2 字の見出しの後ろは、見出しだけの段落と同じく、値を書き足す所のことがある。"""
    cnt, s_, e_ = spaced_chain(core, a, b)
    return cnt == 2 and s_ == 0 and e_ == len(core)


def agent_blanks(text: str, found=(), detail: bool = False) -> list[tuple[int, int, str]]:
    """規則では決められない、書き込む所かもしれない所（blank_places の空き・点線のほか）：(始め, 終わり, 何か（AGENT_WHY の鍵）)。
    found は先に見つけた所（重なる所・すぐ隣の所は出さない）。
    どれも字の形と並び（同じ記号の並び・コロンや文の区切りの字（Unicode の性質）で終わるか）だけで決め、語は見ない。
    所は Found で、facts に事実（place_facts。段落の後ろの所は、段落の本文の字数 body_len と終わりの字 last_char も）を添える。
    detail：段落の後ろの所の鍵を、同じ段落にほかの所もあるとき after_label_more に分ける（事実で分けるだけ。字数では分けない）。"""
    taken = sorted(tuple(x[:2]) for x in found)  # 先に見つけた所（始めの順。重なりを二分探索で調べる：所の数 × 段落の長さにしない）
    reach = list(itertools.accumulate((y for _x, y in taken), max))  # その所までの、終わりの最大
    out: list[tuple[int, int, str]] = []

    def meets(a: int, b: int, x: int, y: int) -> bool:
        return (a < y and x < b) or (a == b and x <= a <= y) or (x == y and a <= x <= b)

    def add(a: int, b: int, why: str, extra: dict | None = None) -> None:
        # 先に見つけた所：始めが b 以下のもののうち、終わりが a 以上のものだけを見る
        i = bisect.bisect_right(taken, (b, float("inf")))
        while i > 0 and reach[i - 1] >= a:
            if meets(a, b, *taken[i - 1]):
                return
            i -= 1
        # ここで足す所は、位置の順に足す（同じ記号の並び → 字の後ろ）ので、最後に足した所とだけ比べればよい
        if out and meets(a, b, out[-1][0], out[-1][1]):
            return
        out.append(Found(a, b, why, extra, text=text))

    core = text.strip("".join(ch for ch in set(text) if chars.is_space(ch)))
    core_end = len(text.rstrip("".join(ch for ch in set(text) if chars.is_space(ch))))
    # 同じ記号の並び（「○○県○○市」は、○○ の並びごとに 1 つの所。あいだの字は facts の before・after にそのまま出る）。
    # 同じ記号だけの行も、ほかと同じ「？」で示す（facts の whole_line）
    for a, b in _repeat_runs(text):
        add(a, b, "placeholder")
    if not core:
        return sorted(out)
    last = core[-1]
    if chars.is_label_colon(last):
        add(core_end, len(text), "after_colon", {"last_char": last})  # 「住所：」「Facility Name:」「Nom :」
        return sorted(out)
    # 文の区切りの字（Unicode の Terminal_Punctuation：。 . , 、 ; ! ? । ؟ など）のない段落の後ろ：見出しだけの段落の後ろに
    # 値を書き足す書類が多い（表題・節の見出しのこともある）。中に 1 字ぶん以上の空白の並び（wide_space：全角の空白 1 つか
    # 半角の空白 2 つ以上）がある段落（空きは別の所として示す。「１　目的」の空きと段落の後ろの両方を出さない）は除く
    body = _lead_marker(core)
    if (body and not any(chars.is_terminal_punctuation(ch) for ch in body)
            and any(unicodedata.category(ch)[0] in ("L", "N") for ch in body)
            and not any(k == "space" and wide_space(core[a:b]) and not _spaced_pair(core, a, b)
                        for a, b, k in chars.char_runs(core))):  # 番号の後ろの空きも数える
        key = "after_label_more" if detail and (found or out) else "after_label"
        add(core_end, len(text), key, {"body_len": len(body), "last_char": last})
    return sorted(out)


def cell_blanks(text: str) -> list[tuple[int, int, str | None]]:
    """表計算のマスの文字の中の、書き込む所（の候補）：(始め, 終わり, 何か)。何かが None なら規則で決まる空欄、
    文字なら規則では決められない所（AGENT_WHY の鍵。エージェントが窓を読んで判断する）。

    - 規則（字の形だけで決まる）：下線の字の並び・かっこの中の空白だけの所（「（　　）」「[ ]」）
    - 規則では決められない：1 字ぶん以上の空白の並び（wide_space）・点線・ダッシュの並び（blank_places）と、agent_blanks の見つけ方
      （同じ記号の並び・コロンで終わる字の後ろ）、箱の字（印のない □ ☐ など：bullet_box、印のついた ☑ ☒：checked_box。
      記入の箱か箇条・見出しの印かは字からは決まらない。Word・PDF と同じ）
    所は Found（facts に事実。箱の字は字の Unicode の性質 glyph も）。"""
    out: list[tuple[int, int, str | None]] = [
        Found(x[0], x[1], None if x[2] in BLANK_RULE_KINDS else x[2], getattr(x, "facts", {})) for x in blank_places(text)]
    for i, ch in enumerate(text):
        if ch in chars.EMPTY_BOXES or ch in chars.CHECKED_BOXES:
            # 箱の字：書くときは、その字を ☑ などに書き換える（どちらの箱かは、エージェントが前後の字から判断する）
            out.append(Found(i, i + 1, "bullet_box" if ch in chars.EMPTY_BOXES else "checked_box",
                             {"glyph": chars.glyph_facts(ch)}, text=text))
    # 文の区切りのない字（「住所」）の後ろ（after_label）は、表計算のマスでは出さない。表計算では、見出しは1つのマスに書き、
    # 値は別のマスに書く作りで、見出しの後ろまで出すと、すべての見出し・すべての値のマス（表の中の「東京都」など）が候補になり、
    # 窓が読めなくなる。値を書くマスは、入力欄か、印（罫線・塗り・結合）があれば［B5？］として出る。見出しのマスに書き足すときは、
    # 書き換えたあとの文字全体を書く
    out += [x for x in agent_blanks(text, [(a, b) for a, b, _k in out]) if x[2] != "after_label"]
    # 空白だけの空き：書いた字が前後の字にくっつかないよう、元の空白を端に 1 字ずつ残した中を所にする（Word の段落と同じ。
    # gap_interior）。元の空き全体は facts の printed、残した端は kept_edges
    res = []
    for x in out:
        a2, b2, edges = gap_interior(text, x[0], x[1])
        facts = dict(getattr(x, "facts", {}))
        if any(edges):
            facts["kept_edges"] = edges
        res.append(Found(a2, b2, x[2], facts))
    return sorted(res)


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


SPAN_SEP = "@"


def span_id(cell_id: str, at: int) -> str:
    """マスの文字の中の所の id：マスの id（「シート!B3」）＋「@」＋所の始めの字の位置（マスの字の、0 から数えた符号位置の番号）。
    位置はマスの字だけで決まるので、ほかの所が増えても減っても変わらない（「何番目か」とは違う）。"""
    return f"{cell_id}{SPAN_SEP}{at}"


def parse_span_id(fid: str) -> tuple[str, int] | None:
    """span_id の逆：(マスの id, 位置)。所の id でなければ None。"""
    head, sep, tail = str(fid).rpartition(SPAN_SEP)
    if not sep or not head or not tail.isdigit():
        return None
    return head, int(tail)


def fill_blanks(text: str, values, spans: list[tuple[int, int]] | None = None) -> str:
    """マスの文字の中の空欄に、順に値を入れた文字列。値が None の所は、元の字のまま（書かない）。
    values はリスト（所の順）か、所の始めの字の位置（span_id の「@」の後ろ。数か「@5」の字）→ 値の辞書（書かない所は省く）。
    spans は窓に出した空欄（blank_spans）。なければ、文字だけから見つけた所（cell_blanks。規則では決められない所も含む）。"""
    if spans is None:
        spans = [(a, b) for a, b, _k in cell_blanks(text)]
    if isinstance(values, dict):
        by_at = {a: k for k, (a, _b) in enumerate(spans)}
        vals = [None] * len(spans)
        for key, v in values.items():
            at = int(str(key).lstrip(SPAN_SEP)) if str(key).lstrip(SPAN_SEP).isdigit() else None
            if at is None or at not in by_at:
                raise ValueError(f"所 {key} はありません（このマスの所の位置は {', '.join(SPAN_SEP + str(a) for a, _b in spans)}）")
            vals[by_at[at]] = v
        values = vals
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


@dataclass
class TableArea:
    """表の範囲（事実）。how："listobject"（ファイルが表と宣言する：テーブル。ECMA-376 §18.5）か "ruled"（罫線で閉じた格子：
    すべてのマスの 4 辺に罫線があり、縦と横の罫線が中で交わる）。top：初めの行（ruled では、見出しかどうかは決めない）。
    heads：列 → 初めの行の字（そのまま。テーブルは見出しの行の字か、テーブルの列の名前）。body：top より下の行。"""
    how: str
    top: int
    spans: list
    body: list
    heads: dict
    name: str = ""


def detect_tables(wb: Workbook, sh: Sheet, rows: dict[int, list[Anchor]]) -> list[TableArea]:
    """ファイルが表と宣言するもの（テーブル）と、罫線で閉じた格子だけを表とする（行・列の数や字の割合の目安は使わない）。
    窓を作るあいだは、同じ配置について1回だけ調べて控える。"""
    memo = sh.__dict__.get(_MEMO)
    if memo is None:
        return _detect_tables(wb, sh, rows)
    got = memo.get("tables")
    if got is None or got[0] is not rows:
        got = memo["tables"] = (rows, _detect_tables(wb, sh, rows))
    return got[1]


def box_sides(wb: Workbook, sh: Sheet, r1: int, c1: int, r2: int, c2: int) -> dict[str, bool]:
    """マスの範囲の 4 辺の罫線（隣のマスの側に引いた線も数える。ECMA-376 §18.8.4 border）。辺のすべての列・行に線があること。"""
    def bits(r: int, c: int) -> int:
        return wb.style(sh.style_of(r, c)).borders if r >= 1 and c >= 1 else 0

    return {"left": all(bits(r, c1) & 1 or bits(r, c1 - 1) & 2 for r in range(r1, r2 + 1)),
            "right": all(bits(r, c2) & 2 or bits(r, c2 + 1) & 1 for r in range(r1, r2 + 1)),
            "top": all(bits(r1, c) & 4 or bits(r1 - 1, c) & 8 for c in range(c1, c2 + 1)),
            "bottom": all(bits(r2, c) & 8 or bits(r2 + 1, c) & 4 for c in range(c1, c2 + 1))}


def _detect_tables(wb: Workbook, sh: Sheet, rows: dict[int, list[Anchor]]) -> list[TableArea]:
    out: list[TableArea] = []
    taken: list[tuple[int, int, int, int]] = []
    for t in (getattr(wb, "tables", None) or {}).values():
        if t.sheet != sh.name:
            continue
        top = t.r1 + t.header_rows - 1 if t.header_rows else t.r1
        heads = {}
        for k, c in enumerate(range(t.c1, t.c2 + 1)):
            cell = sh.cells.get((top, c)) if t.header_rows else None
            v = cell.value if cell is not None and isinstance(cell.value, str) else (t.columns[k] if k < len(t.columns) else "")
            heads[c] = label_text(v or "")
        body = list(range(t.r1 + t.header_rows, t.r2 - t.totals_rows + 1))
        out.append(TableArea("listobject", top, [(c, c) for c in range(t.c1, t.c2 + 1)], body, heads, t.name))
        taken.append((t.r1, t.c1, t.r2, t.c2))

    closed: dict[tuple[int, int, int], bool] = {}

    def is_closed(r: int, c1: int, c2: int) -> bool:
        k = (r, c1, c2)
        if k not in closed:
            closed[k] = all(box_sides(wb, sh, r, c1, r, c2).values())
        return closed[k]

    def split_ok(r: int, c1: int, c2: int) -> bool:
        """行 r で、列 c1〜c2 が 1 つのマス（結合はその範囲ちょうどで、1 行だけ）か。"""
        m = sh.merged_at(r, c1)
        if c1 == c2:
            return m is None
        return m is not None and (m[0], m[1], m[2], m[3]) == (r, c1, r, c2)

    used: set[tuple[int, tuple]] = set()
    for r in sorted(rows):
        run: list[Anchor] = []
        runs: list[list[Anchor]] = []
        for a in sorted(rows[r], key=lambda a: a.c1):
            if a.r2 == a.r and is_closed(r, a.c1, a.c2) and (not run or a.c1 == run[-1].c2 + 1):
                run.append(a)
                continue
            if len(run) >= 2:
                runs.append(run)
            run = [a] if a.r2 == a.r and is_closed(r, a.c1, a.c2) else []
        if len(run) >= 2:
            runs.append(run)
        for run in runs:
            spans = tuple((a.c1, a.c2) for a in run)
            if (r, spans) in used or any(r1 <= r <= r2 and not (spans[-1][1] < c1 or spans[0][0] > c2)
                                         for r1, c1, r2, c2 in taken):
                continue
            body = []
            rr = r + 1
            # 下の行も同じ列割りで、すべてのマスが閉じている（縦と横の罫線が中で交わる：2 列 × 2 行から格子）
            while all(split_ok(rr, c1, c2) and is_closed(rr, c1, c2) for c1, c2 in spans):
                body.append(rr)
                used.add((rr, spans))
                rr += 1
            if body:
                heads = {a.c1: (label_text(a.cell.value) if a.cell is not None and isinstance(a.cell.value, str) else
                                (display(wb, a.cell, flag=False) if a.cell is not None and a.cell.value is not None else ""))
                         for a in run}
                out.append(TableArea("ruled", r, list(spans), body, heads))
    return out


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
    # 画面に出ないマスの字："row" は非表示の行・列、"format" は条件付き書式でいまは見えない字。配置（rows）には入れず、
    # 読み順にだけ印をつけて出す（ファイルにある字で、エージェントは、あることを知っていなければならない）
    hidden: str = ""

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


def _is_marker(t: str) -> bool:
    """項目の番号・記号だけの字（「①」「(2)」「ア.」「Ⅳ」）か。字の種類（数・1 字の字・かっこ・区切り）だけで見て、語は見ない。"""
    t = t.strip()
    return bool(t) and _lead_marker(t) == ""


def _text(b: Box) -> bool:
    """字の書かれたマス（見出しになりうる、印刷される字）。"""
    return b.kind == "text" and isinstance(b.cell.value, str) and bool(b.cell.value.strip())


def _word(b: Box) -> bool:
    """1 行のマスの、字の書かれたマス（入力欄と交互に並ぶ字）。"""
    return _text(b) and b.r2 == b.r


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
        self.readers = formula_readers(wb, sh)  # セル → それを読む数式（事実。欄の宣言ではない）
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
        # 非表示の行・列にある字：画面には出ないが、ファイルにある（エージェントは、あることを知っていなければならない）。
        # 見出しを決める配置（rows）には入れず、読み順に「非表示」の印をつけて出す
        self.hidden_boxes: dict[int, list[Box]] = {}
        shown_sheet = sh.state not in ("hidden", "veryHidden")  # 非表示のシート（選択肢の元の表など）は、記入する所ではない
        for r, items in anchors(sh).items():
            for a in items:
                if sh.is_hidden(a.r, a.c1):
                    cell = a.cell
                    if cell is not None and (cell.has_formula or (cell.value is not None and cell.value != "")):
                        self.hidden_boxes.setdefault(a.r, []).append(
                            Box(a.r, a.c1, a.r2, a.c2, "formula" if cell.has_formula else "text", cell, display(wb, cell),
                                outside=not sh.in_print_area(a.r, a.c1), hidden="row"))
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
                    # 保護されたシートのロックされた空のマスで、見出しのすぐ右にあり、欄の作り（罫線・塗り・結合・日付や数の
                    # 表示形式）のあるもの：見出しの欄として配置に入れ、「保護されたセル」と示す（見出しはロックの有無でなく
                    # 配置で決める。入れないと、見出しがその右の、ロックを外した別のセルについてしまう）
                    if shown_sheet and sh.protected and (a.r, a.c1) not in self.links and cell is not None and \
                            wb.style(sh.style_of(a.r, a.c1)).locked and self._locked_slot(a):
                        b = Box(a.r, a.c1, a.r2, a.c2, "locked", cell, "", outside=outside)
                        self.rows.setdefault(a.r, []).append(b)
                        for rr in range(a.r + 1, a.r2 + 1):
                            self.cover.setdefault(rr, []).append(b)
                        continue
                    # 空のマス：罫線・塗り・結合（ファイルの作り）があれば、欄かもしれない所としてエージェントに示す。
                    # 保護されたシートのロックされたセル（Excel でも書けない）・部品のつながり先は除く
                    if shown_sheet and (a.r, a.c1) not in self.links and \
                            not (sh.protected and wb.style(sh.style_of(a.r, a.c1)).locked):
                        ev = blank_marks(wb, sh, a)
                        read = self.readers.get((a.r, a.c1))
                        if ev is not None or read is not None:
                            ev = dict(ev or {})
                            if read is not None:
                                ev["referenced_by"] = read
                            self.agent.setdefault(a.r, []).append(
                                Box(a.r, a.c1, a.r2, a.c2, "agent", cell, "", outside=outside,
                                    why=(["blank_box"] if len(ev) > ("referenced_by" in ev) else []) +
                                    (["formula_ref"] if read is not None else []), evidence=ev))
                    continue
                elif self._ref_value(a, shown_sheet):
                    # 数式が読む、数の入ったセル（入力欄の宣言はない）：「？」の所（いまの値を添える）。配置（rows）には入れない
                    self.agent.setdefault(a.r, []).append(
                        Box(a.r, a.c1, a.r2, a.c2, "agent", cell, display(wb, cell), outside=outside, why=["formula_ref"],
                            evidence={**(blank_marks(wb, sh, a) or {}), "referenced_by": self.readers[(a.r, a.c1)]}))
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
        # 数式が読む、ファイルにセルの要素がない（値も書式もない）マス：「？」の所（formula_ref）。保護されたシートの
        # ロックされたマス（Excel でも書けない）・非表示のシート・行・列・結合の中は除く
        if shown_sheet:
            for (r, c), read in self.readers.items():
                if (r, c) in anchored or (r, c) in self.links or (r, c) in self.inputs or sh.is_hidden(r, c) or \
                        sh.merged_at(r, c) is not None or (sh.protected and wb.style(sh.style_of(r, c)).locked):
                    continue
                self.agent.setdefault(r, []).append(
                    Box(r, c, r, c, "agent", None, "", outside=not sh.in_print_area(r, c), why=["formula_ref"],
                        evidence={"referenced_by": read}))
        # 結合セルの左上でないマスに残っている字（Excel の画面には出ないが、ファイルにある）：読み順に印をつけて出す
        if sh.merges:
            for (r, c), cell in sh.cells.items():
                if (r, c) in anchored or not (cell.has_formula or (cell.value is not None and cell.value != "")):
                    continue
                m = sh.merged_at(r, c)
                if m and (r, c) != (m[0], m[1]):
                    self.hidden_boxes.setdefault(r, []).append(
                        Box(r, c, r, c, "formula" if cell.has_formula else "text", cell, display(wb, cell),
                            outside=not sh.in_print_area(r, c), hidden="merge"))
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
                            b.hidden = "format"
                            self.hidden_boxes.setdefault(b.r, []).append(b)
                            continue
                    keep.append(b)
                self.rows[r] = keep
            hidden = set(self.hidden_by_format)
            for r in list(self.cover):
                self.cover[r] = [b for b in self.cover[r] if (b.r, b.c1) not in hidden]
        for r in self.rows:
            self.rows[r].sort(key=lambda b: (b.c1, b.kind != "check"))
        # 列ごとの控えを、配置が決まったここで作る（見出しを探すあいだの仮のマスを入れない）
        self._above_box(Box(0, 1, 0, 1, "agent"))
        self._next_row(0, (1, 1))
        self._agent_cells(under)
        self._agent_text_spans()

    def _ref_value(self, a: "Anchor", shown_sheet: bool) -> bool:
        """数式が読む、数の入ったセルで、入力欄の宣言のないもの（「？」の所。書けないセル・部品のつながり先は除く）。"""
        cell = a.cell
        return (shown_sheet and cell is not None and not cell.has_formula and is_number(cell.value)
                and (a.r, a.c1) in self.readers and (a.r, a.c1) not in self.links
                and not (self.sh.protected and self.wb.style(cell.style).locked))

    def _locked_slot(self, a: "Anchor") -> bool:
        """ロックされた空のマスが、見出しの欄の作りか：同じ行の左のいちばん近いマスが見出しの字で（何列はなれていても。
        離れ方（列の数・mm）は欄の near の left に出す）、
        マスに日付・数の表示形式（作り手が値の形を決めた印）があるか、罫線・塗り・結合があって同じ行の右に入力欄がある
        （配置に入れないと、見出しがその入力欄についてしまう所）。"""
        row = self.rows.get(a.r, [])
        left = [b for b in row if b.c2 < a.c1 and b.kind != "check"]
        if not left:
            return False
        b = left[-1]
        if not (b.kind == "text" and b.r2 == b.r and isinstance(b.cell.value, str) and b.cell.value.strip()):
            return False
        fmt = self.wb.style(self.sh.style_of(a.r, a.c1)).num_fmt or "General"
        if fmt not in ("General", "@"):
            return True
        last = self.__dict__.get("_last_input")
        if last is None:
            last = self.__dict__["_last_input"] = {}
            for (r, c) in self.inputs:
                last[r] = max(c, last.get(r, 0))
        return last.get(a.r, 0) > a.c2 and blank_marks(self.wb, self.sh, a) is not None

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
        """空のマスのいちばん近くの字（事実）：同じ行の左（あいだの列の数 cols と mm）と、同じ列の上（行の番号の差 rows と mm）。
        あいだの空の欄（値のない入力欄）はこえて、その数を fields_between に出す（表の格子のマスの、行の見出し・列の見出し）。
        字でない値・数式・値の入った欄があれば、そこで止める。字はそのまま（text）・見出しの字（label）・探す鍵（key）。
        何行・何列はなれていても出す（見出しかどうかはエージェントが判断する）。"""
        out: dict = {}

        def empty(x: Box) -> bool:
            return x.kind == "input" and not x.text and (x.cell is None or x.cell.value in (None, ""))

        def shown(x: Box) -> str:
            if x.kind != "text" or x.cell is None or x.cell.value in (None, ""):
                return ""
            v = x.cell.value
            return v if isinstance(v, str) else display(self.wb, x.cell, flag=False).strip()

        def fact(x: Box, t: str, skipped: int) -> dict:
            d = {"text": t, "label": label_text(t), "key": label_key(t), "cell": x.cell_ref}
            if not isinstance(x.cell.value, str):
                d["number"] = True  # 数の値（表示形式を通した字。時刻の見出しのことも、記入例の数のこともある）
            if skipped:
                d["fields_between"] = skipped
            return d

        left = sorted((x for x in self.boxes(b.r) + self.cover.get(b.r, []) if x.c2 < b.c1 and x.kind != "check"),
                      key=lambda x: x.c2, reverse=True)
        skipped = 0
        for x in left:
            if empty(x):
                skipped += 1
                continue
            t = shown(x)
            if t.strip():
                out["left"] = {**fact(x, t, skipped), "cols": b.c1 - x.c2 - 1,
                               "mm": round(sum(col_mm(self.wb, self.sh, c) for c in range(x.c2 + 1, b.c1)), 1)}
            break
        probe, skipped = b, 0
        while True:
            x = self._above_box(probe)
            if x is None:
                break
            if empty(x):
                skipped += 1
                probe = Box(x.r, b.c1, x.r, b.c2, "probe")
                continue
            t = shown(x)
            if t.strip():
                out["above"] = {**fact(x, t, skipped), "rows": b.r - x.r2,
                                "mm": round(sum(row_mm(self.sh, r) for r in range(x.r2 + 1, b.r)), 1)}
            break
        return out

    def _next_row(self, rr: int, span: tuple[int, int]) -> int:
        """rr から上（rr を含む）で、列の範囲 span にかかるマス（その行で始まるか、上から結合で続くもの）のある、いちばん下の行。
        なければ 0。列ごとの、マスのある行の控え（配置が決まったあとに 1 度だけ作る）を二分探索する。"""
        idx = self.__dict__.get("_rows_by_col")
        if idx is None:
            sets: dict[int, set] = {}
            for rows in (self.rows, self.cover):
                for r_, items in rows.items():
                    for x in items:
                        for c in range(x.c1, min(x.c2, x.c1 + 255) + 1):
                            sets.setdefault(c, set()).add(r_)
            idx = self.__dict__["_rows_by_col"] = {c: sorted(v) for c, v in sets.items()}
        best = 0
        for c in range(span[0], span[1] + 1):
            xs = idx.get(c)
            if xs:
                i = bisect.bisect_right(xs, rr) - 1
                if i >= 0 and xs[i] > best:
                    best = xs[i]
        return best

    def _above_box(self, b: Box) -> Box | None:
        """b の列（c1〜c2）にかかる、b より上でいちばん近い行のマス（左から 1 つ目）。列ごとの控え（行の順）を二分探索する。"""
        idx = self.__dict__.get("_col_index")
        if idx is None:
            idx = {}
            seen = set()
            for rows in (self.rows, self.cover):
                for items in rows.values():
                    for x in items:
                        if id(x) in seen or x.kind == "check":
                            continue
                        seen.add(id(x))
                        for c in range(x.c1, min(x.c2, x.c1 + 255) + 1):
                            idx.setdefault(c, []).append(x)
            for c in idx:
                idx[c].sort(key=lambda x: x.r2)
            self.__dict__["_col_index"] = idx
            self.__dict__["_col_keys"] = {c: [x.r2 for x in xs] for c, xs in idx.items()}
        keys = self.__dict__["_col_keys"]
        best = None
        for c in range(b.c1, b.c2 + 1):
            xs = idx.get(c)
            if not xs:
                continue
            i = bisect.bisect_left(keys[c], b.r) - 1  # 終わりの行が b より上のマスのうち、いちばん下
            if i >= 0:
                x = xs[i]
                if best is None or (x.r2, -x.c1) > (best.r2, -best.c1):
                    best = x
        return best

    def _agent_text_spans(self) -> None:
        """文字の中の、規則では決められない所のうち、コロンで終わる字の後ろ（「住所：」）は、すぐ右のマスが欄・空のマスなら
        出さない（値はそちらに書く。マスの並びだけで決める）。空白の並びは、隣に欄があっても消さずに示す（どちらに書くかは
        エージェントが判断する）。"""
        for r, items in self.rows.items():
            slots = {b.c1 for b in items + self.agent.get(r, []) if b.kind in ("input", "agent")}
            for b in items:
                if b.kind != "text" or not b.blanks:
                    continue
                if b.c2 + 1 in slots:
                    b.blanks = [x for x in b.blanks if x[2] != "after_colon"]

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
        区分の最初の行にだけ書かれ、下の行は同じ列割りで続いている見出しは引き継ぐ（同じ列割りの字の行が続くあいだ。
        行の数では打ち切らない。何行上の字かは、describe が near の section に出す）。"""
        out = [b for b in self.cover.get(r, []) + self.boxes(r)
               if b.kind == "text" and b.r2 > b.r and b.c2 < before_col and isinstance(b.cell.value, str)]
        texts = [b for b in self.boxes(r) if b.kind == "text" and b.c2 < before_col and isinstance(b.cell.value, str)
                 and b.cell.value.strip() and b.r2 == b.r]
        if len(texts) == 1 and not any(o.c2 < texts[0].c1 for o in out):
            item = texts[0]
            for rr in range(r - 1, 0, -1):
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
        """入力欄と、すぐ隣のマス（あいだに空の列をはさまない）の字が交互に並ぶ所（「令和｜［］｜年｜［］｜月｜［］｜日」）。
        字の意味は見ず、マスの並びだけで決める。両端の欄（この欄のほか）は外し、両端が字になるようにする。
        describe は、両端が字で、欄が2つ以上あるもの（字にはさまれた欄の並び）だけを、並びとして見出しに出す。"""
        row = [b for b in self.boxes(r) if b.kind != "check" and (not b.outside or target.outside)]
        i = row.index(target)

        def slot(b: Box) -> bool:
            return b.kind == "input" or (b.kind == "formula" and bool(b.text))

        seq = [target]
        for step in (-1, 1):
            j = i + step
            prev = target
            while 0 <= j < len(row):
                b = row[j]
                if not ((_word(b) and slot(prev)) or (slot(b) and _word(prev))):
                    break
                if step < 0:
                    seq.insert(0, b)
                else:
                    seq.append(b)
                prev = b
                j += step
        while len(seq) > 1 and seq[0] is not target and not _word(seq[0]):
            seq.pop(0)
        while len(seq) > 1 and seq[-1] is not target and not _word(seq[-1]):
            seq.pop()
        return seq, seq.index(target)

    def _list_table(self, r: int, c: int):
        """セル (r, c) を本体（見出しの行・集計の行を除く）に含むテーブル（リストオブジェクト）。なければ None。"""
        for t in (getattr(self.wb, "tables", None) or {}).values():
            if t.sheet == self.sh.name and t.c1 <= c <= t.c2 and t.r1 + t.header_rows <= r <= t.r2 - t.totals_rows:
                return t
        return None

    def header(self, target: Box, phrase: list[Box]) -> list[str]:
        """同じ列の上のいちばん近い見出し（1 段。その上に続く見出しは header_chain の事実として near の header に出す）。"""
        chain = self.header_chain(target, phrase)
        return [label_text(chain[0][0])] if chain else []

    def header_chain(self, target: Box, phrase: list[Box]) -> list[tuple[str, Box | None]]:
        """同じ列の上に続く見出しの並び（近い順。[(印刷された字, マス)]）。何段でも、欄・数式・数のある行に当たるまで
        たどる（段の数・見出しの幅の目安で打ち切らない。表題のような広い字も、そのまま事実として出す）。
        いちばん近い段は、行の中のほかのマスまで覆う字（その列だけの見出しではない）なら、見出しなしとする。
        テーブル（リストオブジェクト）の本体のマスは、テーブルの見出しの行の字（ファイルが見出しと決めている字）だけ。"""
        t = self._list_table(target.r, target.c1)
        if t is not None and t.header_rows:
            hr = t.r1 + t.header_rows - 1
            cell = self.sh.cells.get((hr, target.c1))
            v = cell.value if cell is not None and isinstance(cell.value, str) and cell.value.strip() else None
            if v is None and target.c1 - t.c1 < len(t.columns):
                v = t.columns[target.c1 - t.c1] or None
            if not v:
                return []
            hb = next((x for x in self._cand(hr, (target.c1, target.c1)) if x.kind == "text"), None)
            return [(v, hb)]
        others = [b for b in self.boxes(target.r) if b not in phrase and b.kind != "check"]
        lo = min(b.c1 for b in phrase) if phrase else target.c1
        hi = max(b.c2 for b in phrase) if phrase else target.c2
        out: list[tuple[str, Box | None]] = []
        r = target.r - 1
        shape = self.shape(target.r)
        span = (target.c1, target.c2)
        # 何行はなれていても見る（行の数の目安で打ち切らない。離れ方は near に出す）。列の範囲にマスのない行はとばす
        while r > 0:
            if not out:
                # 表の本体の行（同じ形）と、同じ列割りの欄が上に続く行（表の列）は、数えずに上へ。続く所の上まで、まとめて飛ぶ
                r = self._skip_up(r, shape, span)
                if r <= 0:
                    break
            r = self._next_row(r, span)
            if r <= 0:
                break
            cand = self._cand(r, span)
            if not out and any(b.kind == "input" for b in cand):
                break  # すぐ上に別の入力欄がある：この欄の見出しは上にはない
            texts = [b for b in cand if _text(b)]
            if out and any(b.kind in ("input", "formula") or (b.kind == "text" and not isinstance(b.cell.value, str))
                           for b in cand):
                break  # 上に続く見出しは、欄や数値をまたがない
            if not texts:
                r -= 1
                continue
            h = texts[0]
            if not out:
                covered = [o for o in others if o.c1 <= h.c2 and o.c2 >= h.c1 and not (lo <= o.c1 and o.c2 <= hi)
                           and not (o.kind == "formula" and not o.text)]
                if covered:
                    break
            if all(x is not h for _t, x in out):
                out.append((h.cell.value, h))
            span = (h.c1, h.c2)
            r = h.r - 1
        return out

    def _fact(self, x: Box | None, target: Box, text: str | None = None) -> dict:
        """近くの字の事実：印刷された字のまま（text）・探す鍵（key）・番地・広さ（cols）・target からの離れ方
        （上なら rows と mm、左なら cols_between と mm）。"""
        t = text if text is not None else (x.cell.value if x is not None and x.cell is not None else "")
        d: dict = {"text": t, "label": label_text(t), "key": label_key(t)}
        if x is None:
            return d
        d["cell"] = x.ref if x.control is None else x.cell_ref
        d["cols"] = x.c2 - x.c1 + 1
        if x.r2 < target.r:
            d["rows"] = target.r - x.r2
            d["mm"] = round(sum(row_mm(self.sh, rr) for rr in range(x.r2 + 1, target.r)), 1)
        elif x.c2 < target.c1:
            d["rows"] = 0
            d["cols_between"] = target.c1 - x.c2 - 1
            d["mm"] = round(sum(col_mm(self.wb, self.sh, c) for c in range(x.c2 + 1, target.c1)), 1)
        return d

    def describe(self, target: Box, options: list | None = None) -> dict:
        """入力欄（またはチェックボックス）の見出しと、判断の材料。見出しは、印刷された字をそのまま使う：
        同じ行の前（左。右から左へ表示するシートでは画面の右）のいちばん近い字・同じ列の上の字。字の意味（単位・年号など）では選ばない。"""
        r = target.r
        row = self.boxes(r)
        phrase, pi = self.phrase(r, target) if target.kind == "input" else ([target], 0)
        n_slots = sum(b.kind in ("input", "formula") for b in phrase)
        shown = n_slots >= 2 and _word(phrase[0]) and _word(phrase[-1])  # 字にはさまれた、2 つ以上の欄の並び
        if not shown:
            phrase = [target]
        first = phrase[0]
        sec_boxes = self.sections(r, first.c1)
        secs = [label_text(b.cell.value) for b in sec_boxes]
        left_edge = first.c1
        nearest = [b for b in row if b.c2 < left_edge and b.kind != "check"]
        lefts = []
        if nearest:
            b = nearest[-1]
            if _text(b) and b.r2 == b.r:
                lefts = [b]  # 何列はなれていても、いちばん近い字（離れ方は near に出す）
        right_edge = phrase[-1].c2
        rights = [b for b in row if b.c1 > right_edge and _text(b)]
        # チェックボックス（部品）は、後ろ（右）に書かれた項目が見出し（部品の作り：字は部品の右に置く）。入力規則の選択肢の
        # 数や長さからは、欄の種類（○を付ける欄か）を決めない（選択肢はそのまま options に出し、前後の字は near に出す）
        mark = target.kind == "check"
        item = None
        method = []
        before, after = self.side("左"), self.side("右")
        if target.kind == "check":
            carried_right = [b for b in self.cover.get(r, []) if b.c1 > target.c2 and _text(b)]
            rights = sorted(rights + carried_right, key=lambda b: b.c1)
        if mark and rights:
            # チェック欄・○を付ける欄は、後ろ（右）に書かれた項目が見出し（①などの番号だけのマスは、続く字と合わせる）
            nxt = [b for b in row if b.c1 > target.c2 and b.kind != "check"]
            if (nxt and nxt[0].kind == "text") or target.kind == "check":
                t = rights[0]
                item = label_text(t.cell.value)
                if _is_marker(item):
                    nxt_text = [b for b in self.boxes(t.r) if b.c1 > t.c2 and _text(b)]
                    if nxt_text:
                        item = item + label_text(nxt_text[0].cell.value)
                    elif len(rights) > 1:
                        item = item + label_text(rights[1].cell.value)
                method.append(f"{after}の項目")
        if item is None and lefts:
            item = label_text(lefts[-1].cell.value)
            method.append(f"{before}の見出し")
        if lefts:
            # 同じ行で、見出しのすぐ前に続く見出し（「法人名｜フリガナ」の「法人名」）は区分として読む。
            # あいだに欄がはさまるもの（横に並んだ別の表の見出し）は区分にしない
            more = []
            before_lab = [b for b in row if b.c2 < lefts[0].c1]
            for k in range(len(before_lab) - 1, -1, -1):
                b = before_lab[k]
                if not (_text(b) and b.r2 == b.r):
                    break
                if k > 0 and before_lab[k - 1].kind in ("input", "formula"):
                    break  # すぐ前が欄の字（「［］ヶ月）」の「ヶ月）」）：前の欄の並びの続きで、この欄の区分ではない
                more.insert(0, b)
            for b in more:
                t = label_text(b.cell.value)
                if _is_marker(t) and item:
                    item = t + item
                elif t not in secs:
                    secs.append(t)
        chain = self.header_chain(target, phrase)
        heads = [label_text(chain[0][0])] if chain else []
        parts = list(secs)
        between_text = None  # 前の字が、あいだの別の欄の見出しのとき（見出しにせず、別に示す）
        near = self._near_text(target)
        if heads and (item is None or self._list_table(r, target.c1) is not None):
            # 前（左）に見出しのない欄は、上のいちばん近い見出しで呼ぶ。テーブル（リストオブジェクト：ファイルが表と宣言する）の
            # 本体は、行の見出し（前）と列の見出し（上）の両方で呼ぶ。同じ形の行が続くだけ（宣言のない表）では、上の見出しを
            # 見出しに付けない（上の見出しの並びは near の header に、字のまま・離れ方つきで出す）
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
        elif target.kind != "check":
            # 近くに見出しのない欄（すぐ前が別の欄の、補助のマスなど）：同じ表の続きの中で、前にある字をそのまま示す。
            # あいだにある欄は、method に書く（その字が、この欄の見出しとは限らない）。前に字がなければ（表の格子の、行の見出しの
            # ない行のマスなど）、左と上のいちばん近い字（あいだの空の欄をこえる。near の left・above）のうち近いほう（mm。同じ
            # なら左）。数で表の作りを決めない（字と離れ方は near に出す。見出しかどうかはエージェントが判断する）
            key = self.row_key(r, first.c1, with_between=True)
            if not key and not parts:
                got =[(f_["mm"], k_, f_) for k_, f_ in (("left", near.get("left")), ("above", near.get("above")))
                       if f_ and not f_.get("number")]
                if got:
                    _mm, side, f_ = min(got, key=lambda x: (x[0], x[1] != "left"))
                    parts.append(f_["label"])
                    word = "上" if side == "above" else before
                    method.append(f"{word}の字（あいだに空の欄 {f_['fields_between']}）" if f_.get("fields_between") else f"{word}の字")
            if key:
                text, between = key
                # 補助のセル：数式が参照する値で、もう値の入っているもの（空の 1 桁ずつのマスは、数式が参照しても欄の並び）
                helper = ((target.r, target.c1) in self.readers
                          and target.cell is not None and target.cell.value not in (None, "")) or any(
                    b.kind == "locked" and b.cell_ref in between for b in row)
                if between and helper:
                    # あいだに別の欄がはさまり、この欄が数式の参照する値（補助のセル）か、あいだがロックされた見出しの欄
                    # （B5）：その字はあいだの欄の見出しで、この欄の見出しとは言えない。見出しにせず、前の字として別に示す
                    # （本人に伝えるときに「開始日：5」「氏名：10」と取り違えない）
                    between_text = f"{text}（あいだに欄 {'・'.join(between)}）"
                    method.append(f"{before}の字は、あいだの欄の見出し（この欄の見出しにしない）")
                else:
                    # あいだの欄が、同じ並びの欄（郵便番号を 1 桁ずつ書くマス「〒｜［M24］｜［N24］」）：前の字をそのまま示す
                    parts.append(text)
                    method.append(f"{before}の字（あいだに欄 {'・'.join(between)}）" if between else f"{before}の字")
        phrase_text = None
        if shown:
            phrase_text = "".join(
                "［▶］" if b is target else "［　］" if b.kind == "input" else f"〈{b.text}〉" if b.kind == "formula"
                else b.cell.value.strip() for b in phrase)
            method.append("前後の字")
        label = "／".join(dict.fromkeys(p for p in parts if p))
        section = None
        if phrase_text:
            # 字にはさまれた欄の並び（「第［▶］号」）：見出しは、その並びの印刷された字そのまま。前のマスの見出し・区分・
            # 上の見出し・行の名前は、ほかのマス・行の字なので、見出しの前に付けず section に分けて示す
            # （「個人番号：手帳番号［▶］」のように、別の欄の字を見出しに付けない）
            section = "／".join(p for p in dict.fromkeys(parts) if p) or None
            label = phrase_text
        out = {"label": label or None, "label_method": "・".join(method) or None,
               "context": self.line(r, mark_box=target)}
        if section:
            out["section"] = section  # 区分・上の見出しなど（見出しには付けていない、ほかのマス・行の字）
        if between_text:
            out["text_before"] = between_text  # 前（左）の字と、あいだの欄（見出しにはしていない）
        if not shown and target.kind != "check":
            # すぐ後ろ（右。右から左のシートでは左）のマスの字（「金額｜［］｜円」の「円」）：そのまま示す（単位かどうかは決めない）
            nxt = [b for b in row if b.c1 > target.c2 and b.kind != "check"]
            if nxt and nxt[0].c1 == target.c2 + 1 and _word(nxt[0]):
                out["after"] = label_text(nxt[0].cell.value)
        rt = [b for b in row if b.c1 > target.c2 and b.kind != "check"]
        if rt and _text(rt[0]):
            near["right"] = {"text": rt[0].cell.value, "label": label_text(rt[0].cell.value), "key": label_key(rt[0].cell.value), "cell": rt[0].cell_ref,
                             "cols": rt[0].c1 - target.c2 - 1,
                             "mm": round(sum(col_mm(self.wb, self.sh, c) for c in range(target.c2 + 1, rt[0].c1)), 1)}
        if chain:
            near["header"] = [self._fact(x, target, t) for t, x in chain]  # 上に続く見出しの並び（近い順。何段でも）
        if sec_boxes:
            near["section"] = [self._fact(x, target) for x in sec_boxes]  # 区分の見出し（縦の結合・上から続く行の字）
        if near:
            out["near"] = near  # 近くの字（事実）：左・上・右のいちばん近い字と、その離れ方（マスの数と mm）
        return out

    def side(self, word: str) -> str:
        """窓に書く「左」「右」：右から左へ表示するシート（sheetView の rightToLeft）では、列 A が画面の右にあるので入れ替える。
        配置は列の番号の順（読む順）で調べるので、字の前は列の番号の小さい側（左から右のシートでは左、右から左のシートでは右）。"""
        if not getattr(self.sh, "right_to_left", False):
            return word
        return {"左": "右", "右": "左"}.get(word, word)

    def row_key(self, r: int, before_col: int, with_between: bool = False):
        """表の行の名前：前（左）へたどって、欄をとばして最初にある字の見出し（横に並んだ表でも、その表の行の名前）。
        何も書かれていない列（空の列）か、テーブル（リストオブジェクト）の端まで来たら、別の表なので止める。
        with_between：(字, あいだにとばした欄の番地) を返す。"""
        t = self._list_table(r, before_col)
        stop = t.c1 if t is not None else 0
        boxes = [b for b in self.boxes(r) + self.cover.get(r, []) if b.c2 < before_col and b.kind != "check"]
        edge = before_col
        between: list[str] = []
        for b in sorted(boxes, key=lambda b: b.c2, reverse=True):
            if b.c2 < stop or b.c2 < edge - 1:
                break  # 空の列をはさむ・テーブルの外
            edge = min(edge, b.c1)
            if b.r == r and _text(b) and not _is_marker(b.cell.value):
                text = label_text(b.cell.value)
                return (text, between) if with_between else text
            if b.kind in ("input", "formula", "agent", "locked"):
                between.insert(0, b.cell_ref)
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
                   for b in self.boxes(r)) and not any(b.text for b in self.agent.get(r, []))

    def token(self, b: Box, mark: bool = False) -> str:
        if b.kind == "input":
            inner = b.ref + (f"＝{b.text}" if b.text else "")
            return f"［{'▶' if mark else ''}{inner}］"
        if b.kind == "formula":
            return f"〈{b.cell_ref}＝{b.text}〉"
        if b.kind == "check":
            return f"{b.text}［{'▶' if mark else ''}{b.ref}］"
        if b.kind == "agent":
            inner = b.ref + (f"＝{b.text}" if b.text else "")
            return f"［{'▶' if mark else ''}{inner}？］"  # 規則では欄と決められない所（数式が読む値のあるマスは、その値も）
        if b.kind == "locked":
            return f"［{'▶' if mark else ''}{b.ref}：保護されたセル］"  # 保護されたシートのロックされた欄（書けない）
        if b.blanks:
            # 文字の中の空欄は［　］で示し、どのセルの文字かを添える。規則では決められない所は［？］
            t = b.cell.value
            ends = spaced_ends(t)
            parts, at = [], 0
            for a, z, k in b.blanks:
                parts.append(with_spaced_marks(t[at:a], {e - at for e in ends if at < e <= a}))
                parts.append(blank_token(t, a, z, k, mark))
                at = z
            parts.append(with_spaced_marks(t[at:], {e - at for e in ends if e > at}))
            return f"〔{b.cell_ref}〕" + "".join(parts).replace("\n", " ").strip()
        if b.cell is not None and isinstance(b.cell.value, str) and b.text == b.cell.value:
            return with_spaced_marks(b.text).replace("\n", " ").strip()
        return b.text.replace("\n", " ").strip()

    def line(self, r: int, mark_box: Box | None = None, include_outside: bool = True) -> str:
        own = self.boxes(r)
        if r in self.agent:
            own = sorted(own + self.agent[r], key=lambda b: (b.c1, b.kind != "check"))
        own = [b for b in own if include_outside or not b.outside]
        first = min((b.c1 for b in own), default=10 ** 6)
        carried = [b for b in self.cover.get(r, []) if b.kind == "text" and b.c2 < first]
        parts = [f"（{label_text(b.text)}）" for b in sorted(carried, key=lambda b: b.c1)]
        own = [b for b in own if b.kind != "formula" or b.text]
        inside = [self.token(b, b is mark_box) for b in own if not b.outside]
        outside = [self.token(b, b is mark_box) for b in own if b.outside]
        s = " ｜ ".join(p for p in parts + inside if p)
        if outside:
            s += ("　" if s else "") + "〔印刷範囲外〕" + " ｜ ".join(outside)
        for how, mark in (("row", "〔非表示〕"), ("merge", "〔結合の下に隠れた字〕"), ("format", "〔条件付き書式でいまは見えない〕")):
            hidden = [self.hidden_token(b) for b in sorted(self.hidden_boxes.get(r, []), key=lambda b: b.c1) if b.hidden == how]
            hidden = [t for t in hidden if t]
            if hidden:
                s += ("　" if s else "") + mark + " ｜ ".join(hidden)
        return s

    def hidden_token(self, b: Box) -> str:
        """非表示の行・列のマスの字：番地をつけて示す（画面には出ないので、どこの字かが分かるように）。"""
        t = b.text.replace("\n", " ").strip()
        if b.kind == "formula":
            return f"〈{b.cell_ref}＝{t}〉"
        return f"〔{b.cell_ref}〕{t}" if t else ""

    def content_key(self, r: int) -> tuple[tuple, list[Box]]:
        """行の中身：(マスの並び（位置・種類・印刷範囲・非表示・チェックの状態と、上から続く見出しの字）, 並びの順のマス)。
        読み順で行をまとめるときは、並びが同じで、字の違うマスを1つ残らず示せる行だけをまとめる（reading）。
        signature は数を「#」にまとめるので、それだけで比べると値の違う行も同じとみなし、その値が窓から消える。"""
        row = sorted(list(self.boxes(r)) + self.agent.get(r, []) + self.hidden_boxes.get(r, []),
                     key=lambda b: (b.c1, b.kind, b.hidden))
        struct = tuple((b.kind, b.c1, b.c2, b.outside, b.hidden,
                        bool(b.control is not None and b.control.is_checked) if b.kind == "check" else None) for b in row)
        carried = tuple(sorted((b.c1, b.c2, b.text) for b in self.cover.get(r, []) if b.kind == "text"))
        return (struct, carried), row


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


LOCKED_FIELD = ("保護されたセル：書けない／パスワードが要る（シートの保護でロックされていて、Excel でも書けません。"
                "書くには、保護を外すパスワードが要ります。本人が書く欄か、窓を読んで確かめてください）")


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
                shown = "".join(t[x:y] if k % 2 == 0 else blank_token(t, x, y, b.blanks[k // 2][2])
                                for k, (x, y) in enumerate(_split_spans(t, spans)))
                boxes = [i for i, (x, y, _k) in enumerate(b.blanks, 1) if y - x == 1 and t[x:y] in chars.BOXES]
                whys = [k for _x, _y, k in b.blanks if k is not None]
                n = len(b.blanks)
                f = {"id": f"{sh.name}!{b.cell_ref}", "cell": b.cell_ref, "kind": "inline",
                     "value": t, "blanks": n, "label": label_text(shown),
                     "label_method": "文字の中の空欄", "context": lay.line(r, mark_box=b),
                     "write_as": f"空欄ごとの値のリスト（{n} 個、例 {['…'] * n}。［？］の所も数え、書かない所は null）か、"
                                 "書き換えたあとの文字全体"
                                 + (f"（チェックの箱の字の所（{'・'.join(map(str, boxes))} 番目）は、印をつけるなら ☑ などの字、"
                                    "つけないなら null）" if boxes else ""),
                     # 所ごとの、変わらない id（マスの id ＋「@」＋所の始めの字の位置：span_id）・元の字（空白・下線・点線・
                     # 箱の字）・その幅（全角を 1 とした字の幅のおよその数）と、前後の字（そのまま）・事実（facts）
                     "spans": [{"n": i, "id": span_id(f"{sh.name}!{b.cell_ref}", x), "at": x,
                                "before": t[max(0, x - 8):x], "printed": t[x:y], "after": t[y:y + 8],
                                "width": space_width(t, x, y),
                                "rule": k is None,
                                **({"why": AGENT_WHY[k] + (f"。{ONE_WIDE_SPACE_WHY}" if one_wide_space(t[x:y]) else ""),
                                    "why_key": k} if k is not None else {}),
                                **({"facts": bl.facts} if getattr(bl, "facts", None) else {})}
                               for i, bl in enumerate(b.blanks, 1) for (x, y, k) in [bl]],
                     "reason": ["文字の中の空欄"], "provenance": LAYOUT}
                if whys:
                    # 語の並びから推定した所・規則の型に合わない空き：空欄とは言い切らない（窓では［？］）
                    f["agent"] = True
                    f["why"] = "・".join(dict.fromkeys(AGENT_WHY[k] for k in whys))
                    f["provenance"] = AGENT
                out.append(f)
                continue
            if b.kind == "locked":
                d = lay.describe(b)
                out.append({"id": f"{sh.name}!{b.cell_ref}", "cell": b.cell_ref, "kind": _field_kind(wb, b.cell, []),
                            "value": None, **d, "readonly": True, "locked": True,
                            "protected": LOCKED_FIELD, "reason": ["見出しのすぐ右の、保護されたシートのロックされたセル"],
                            "provenance": LAYOUT})
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
    for f in out:
        # どの欄にも、そのセルを読む数式（事実）。入力欄かどうかは、これでは決めない
        if f.get("cell") and "referenced_by" not in f:
            try:
                rc = parse_cell(f["cell"].split(":")[0])
            except Exception:  # noqa: BLE001
                continue
            rb = lay.readers.get(rc) if f["id"].rsplit("!", 1)[0] == sh.name else None
            if rb is not None:
                f["referenced_by"] = rb
    return out


def _clip(t: str, n: int = 40) -> str:
    """窓の短い文に入れる字（n 字まで。切ったら「…」。もとの字は evidence にそのまま）。"""
    t = str(t)
    return t if len(t) <= n else t[:n] + "…"


def _with_codes(fields: list[dict]) -> list[dict]:
    """欄に、理由の変わらない鍵（reason_codes）を添える（reason の文と同じ順）。"""
    for f in fields:
        if "reason" in f:
            f["reason_codes"] = reason_codes(f["reason"])
    return with_label_keys(fields)  # 探すための、字間の空白を詰めた写し（見出しは印刷された字のまま）


def evidence_text(ev: dict, before: str = "左") -> str:
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
        t = _clip(label_text(x["text"]))
        parts.append(f"{before}の字「{t}」（{'すぐ' + before if not x['cols'] else str(x['cols']) + ' 列' + str(x.get('mm', '')) + ' mm あけて' + before}）")
    if ev.get("above"):
        x = ev["above"]
        t = _clip(label_text(x["text"]))
        parts.append(f"上の字「{t}」（{x['rows']} 行上" + (f"・{x['mm']} mm あけて" if x.get("mm") else "") + "）")
    if ev.get("size_mm"):
        parts.append(f"大きさ {ev['size_mm']['w']}×{ev['size_mm']['h']} mm")
    if ev.get("referenced_by"):
        rb = ev["referenced_by"]
        parts.append(f"数式 {'・'.join(rb['cells'][:3])}{' ほか' if rb['n'] > 3 else ''}（{rb['n']} 個）が読む")
    return "・".join(parts) or "印なし"


def _agent_field(wb: Workbook, sh: Sheet, lay: Layout, b: Box) -> dict:
    """規則では欄と決められない空のマスの欄（窓では［B5？］）。エージェントが窓を読んで判断し、書く所ならこの id で書く。"""
    d = lay.describe_agent(b)
    whys = list(dict.fromkeys(b.why))
    value = None if b.cell is None or b.cell.value in (None, "") else b.cell.value
    f = {"id": f"{sh.name}!{b.cell_ref}", "cell": b.cell_ref, "kind": _field_kind(wb, b.cell, []), "value": value, **d,
         "agent": True, "why": "・".join(AGENT_WHY[k] for k in whys), "why_key": whys,
         "evidence": dict(b.evidence), "evidence_text": evidence_text(b.evidence, lay.side("左")),
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
                    "（罫線で囲まれたもの［左右か上下の両方・下線］・塗り・結合のあるもの、入力欄のある表の行の空のマス）と、"
                    "数式が読むセル（空か数。［B5＝値？］はいまの値。読む数式は evidence の referenced_by）で、印のない空のマスは出さない。"
                    "入力規則のあるマス・保護されたシートでロックを外したマスは入力欄［B5］、保護されたシートのロックされたマス"
                    "（Excel でも書けない。見出しのすぐ右の欄の作りのものだけは［B5：保護されたセル］と出す）と非表示のシート・行・列は出さない。書く所かは、根拠（罫線・塗り・近くの字）と"
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
        "notation": {"［L15］": "入力欄", "［L15＝値］": "記入済みの入力欄", "［B5？］": AGENT_CELLS_RULE, "［？　］": GAP_TOKEN_NOTE,
                     "〔B3〕令和［？］年": "文字の中の、規則では決められない空き（語の並びから推定した所など）。"
                                      "書くなら、空欄ごとの値のリストで（書かない所は null）",
                     "〈AF35＝値〉": "計算結果（数式）",
                     "［B5：保護されたセル］": "見出しの欄だが、シートの保護でロックされていて書けない（パスワードが要る）",
                     "☐［BA116］ / ☑［BA116］": "チェックボックス（［］はつながっているセル）",
                     "（法人名）": "上の行から続く、縦に結合された見出し", "［▶…］": "いま説明している欄",
                     "〔印刷範囲外〕": "画面には出るが印刷されない",
                     "〔非表示〕〔C5〕値": "非表示の行・列にある字（画面には出ないが、ファイルにある）",
                     "〔結合の下に隠れた字〕〔C5〕値": "結合セルの左上でないマスに残っている字（画面には出ないが、ファイルにある）",
                     "〔条件付き書式でいまは見えない〕〔C5〕値": "条件付き書式で、いまは見えない色などになっている字"},
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
            d, why = shown_text(wb, cell)
            if d != (str(v) if not is_number(v) else None):
                item["display"] = d
            if why:
                item["display_unsupported"] = why
            if is_number(v) and is_date_format(wb.style(cell.style).num_fmt):
                item["type"] = "date"
        if cell is not None and cell.formula:
            item["formula"] = "=" + cell.formula
        if is_input:
            item["input"] = True
            item["input_reason"] = inputs[(r, c)]
            item["input_reason_codes"] = reason_codes(inputs[(r, c)])
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
    fields = _with_codes(fields_for(wb, sh, lay, opts))
    # 規則では入力欄と決められない空のマス（fields にも agent として入れてある）：番地・根拠・推定した見出しの一覧
    undecided = [{"cell": f["cell"], "column": f.get("label") or "", "decide": "agent", "provenance": AGENT,
                  "why": f["why"], "why_key": f["why_key"], "evidence": f["evidence"], "reason": f["why"]}
                 for f in fields if f.get("agent") and f["kind"] != "inline"]
    unshown = [{"cell": it["cell"], "format": wb.style(sh.cells[parse_cell(it["cell"])].style).num_fmt,
                "value": it.get("value"), "why": it["display_unsupported"], "provenance": DET}
               for it in cells if it.get("display_unsupported")]
    labels = {f["cell"]: f for f in fields if f.get("cell")}
    for item in cells:
        f = labels.get(item["cell"])
        if f is not None and f["kind"] != "checkbox" and f.get("label"):
            item["label"] = f["label"]
            if f.get("label_key"):
                item["label_key"] = f["label_key"]
    return {
        "name": sh.name,
        "state": sh.state,
        "protected": sh.protected,
        "right_to_left": getattr(sh, "right_to_left", False),
        "used_range": f"A1:{addr(max(sh.max_row, 1), max(sh.max_col, 1))}",
        "print_area": sh.print_area,
        "hidden_rows": _runs(sorted(sh.hidden_rows)),
        "hidden_cols": _runs(sorted(sh.hidden_cols), cols=True),
        "header_footer": sh.header_footer,
        "fields": fields,
        "undecided": undecided,
        "unshown_formats": unshown,  # 再現できない表示形式のセル（窓には元の値と〔表示形式「…」は再現できません〕を出す）
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
    """表（事実）：ファイルが宣言するテーブルと、罫線で閉じた格子。列の名前は初めの行（テーブルは見出しの行）の字のまま。
    同じ名前・空の名前の列は、行の値の鍵に列の記号を添えて分ける（columns は字のまま）。"""
    from ..formula.parser import num_to_col

    rows = anchors(sh)
    out = []
    for t in detect_tables(wb, sh, rows):
        cols = [t.heads.get(c1, "") for c1, _c2 in t.spans]
        keys = [c if c and cols.count(c) == 1 else f"{c}〔{num_to_col(c1)}〕" for c, (c1, _c2) in zip(cols, t.spans)]
        trs = []
        for br in t.body:
            byc = {a.c1: a for a in rows.get(br, [])}
            trs.append({k: (display(wb, byc[c1].cell) if c1 in byc and byc[c1].cell else "") for k, (c1, _c2) in zip(keys, t.spans)})
        last = t.body[-1] if t.body else t.top
        d = {"how": t.how, "range": f"{addr(t.top, t.spans[0][0])}:{addr(last, t.spans[-1][1])}", "header_row": t.top,
             "columns": cols, "rows": trs, "provenance": DET}
        if t.how == "listobject":
            d["name"] = t.name
            d["method"] = "ファイルが宣言するテーブル（リストオブジェクト）。列の名前は見出しの行の字"
        else:
            d["decide"] = "agent"
            d["method"] = ("罫線で閉じた格子（すべてのマスの 4 辺に罫線があり、縦と横の罫線が中で交わる）。columns は初めの行の字の"
                           "まま。初めの行が見出しか・表として読むかは、エージェントが判断する")
        out.append(d)
    return out


def reading(lay: Layout) -> list[dict]:
    """行を上から（非表示の行・列の字は印をつけて）。同じ形の空欄の行（または2〜4行で1組の行）が続くところは1行にまとめる。

    まとめてよいのは、欄のある行で、何も落とさずにまとめられる所だけ：マスの並び（位置・種類・印刷範囲・非表示）が同じで、
    入力欄はすべて空欄、上の組と字が違うマスは1つの列だけ（番号の列など）。その列の字は、まとめの行に番地つきで全部書く。
    字の違う列が2つ以上ある行（データの表など）・欄のない行（注記・データ）は、まとめずに1行ずつ出す。"""
    out: list[dict] = []
    rows = sorted(set(lay.rows) | set(lay.agent) | set(lay.hidden_boxes))
    texts = {r: lay.line(r) for r in rows}
    keys = {}

    def key(x):
        k = keys.get(x)
        if k is None:
            k = keys[x] = lay.content_key(x)
        return k

    def differ(tmpl: int, x: int) -> list[Box] | None:
        """行 x の、組の行 tmpl と字の違うマス（並びが違えば None）。違ってよいのは、文字と計算結果のマスだけ。"""
        (s1, c1), b1 = key(tmpl)
        (s2, c2), b2 = key(x)
        if s1 != s2 or c1 != c2:
            return None
        out_ = []
        for p, q in zip(b1, b2):
            if p.text != q.text:
                if q.kind not in ("text", "formula"):
                    return None
                out_.append(q)
        return out_

    i = 0
    while i < len(rows):
        r = rows[i]
        if not texts[r]:
            i += 1
            continue
        best = None  # (組の行数, 続く組の数, 字の違うマス)
        for per in (1, 2, 3, 4):
            group = rows[i:i + per]
            if len(group) < per or group[-1] - group[0] != per - 1:
                break
            # まとめるのは、欄（入力欄・規則では決められない所・チェックボックス）のある行の組だけ。字だけの行（注記・データ）は
            # 1行ずつ出す
            if not any(b.kind in ("input", "agent", "check") for x in group for b in key(x)[1]):
                continue
            n, cols, diffs = 0, set(), []
            j = i + per
            while j + per <= len(rows):
                nxt = rows[j:j + per]
                # 上の組から続く結合セル（縦に長い見出しなど）がかかる行は、まとめない
                if (nxt[0] != rows[j - 1] + 1 or nxt[-1] - nxt[0] != per - 1
                        or not all(lay.is_blank_row(x) and all(b.r >= nxt[0] for b in lay.cover.get(x, []))
                                   for x in nxt)):
                    break
                got = [differ(t, x) for t, x in zip(group, nxt)]
                if any(d is None for d in got):
                    break
                new_cols = cols | {b.c1 for d in got for b in d}
                if len(new_cols) > 1:
                    break  # 字の違う列が2つ以上：まとめると読みにくい（データの表）。1行ずつ出す
                cols = new_cols
                diffs += [b for d in got for b in d]
                n += 1
                j += per
            if n >= (2 if per == 1 else 1) and (best is None or n * per > best[1] * best[0]):
                best = (per, n, diffs)
        if best is None:
            out.append({"row": r, "text": texts[r]})
            i += 1
            continue
        per, n, diffs = best
        for x in rows[i:i + per]:
            if texts[x]:
                out.append({"row": x, "text": texts[x]})
        a, b = rows[i + per], rows[i + per * (n + 1) - 1]
        what = f"{r}行目と同じ形の行" if per == 1 else f"{r}〜{rows[i + per - 1]}行目と同じ形の{per}行1組"
        item = {"rows": f"{a}-{b}", "same_as": f"{r}-{rows[i + per - 1]}"}
        if diffs:
            from ..formula.parser import num_to_col

            col = num_to_col(diffs[0].c1)
            vals = "、".join(f"{d.cell_ref}＝{d.text.replace(chr(10), ' ').strip()}" for d in diffs)
            item["differs"] = [{"cell": d.cell_ref, "text": d.text} for d in diffs]
            item["text"] = (f"{what}が {n} {'行' if per == 1 else '組'}続きます（入力欄はすべて空欄。"
                            f"字が違うのは {col} 列だけで、ほかの字は同じ：{vals}）")
        else:
            item["text"] = f"{what}が {n} {'行' if per == 1 else '組'}続きます（字もすべて同じで、入力欄はすべて空欄）"
        out.append(item)
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
        # 非表示のシート：画面には出ないが、中身は窓に出す（選択肢の元の表・計算用の値・注意書きがあることを、エージェントが
        # 知っていなければならない）。記入する所ではないので、入力欄の一覧は出さない
        n = sum(1 for c in sh["cells"] if c.get("value") is not None)
        out = [f"## シート「{sh['name']}」（非表示のシート{'（VBA からだけ表示できる）' if state == 'veryHidden' else ''}）", "",
               f"画面に出ないシートです（値のあるセル {n:,} 個）。記入する所ではありません。"
               "ここから作られる選択肢は、各入力欄の「選択肢」に展開してあります。中身は次のとおりです。", ""]
        if sh.get("reading"):
            out += ["### 読み順（上から。非表示のシート）", ""] + _md_reading(sh) + [""]
        return out + sheet_markdown(sh)
    out = [f"## シート「{sh['name']}」", ""]
    info = []
    if sh.get("right_to_left"):
        info.append("右から左へ表示するシート（列 A が右端）。読み順は列 A から（画面の右から左へ）。見出しの「左」「右」は画面で見た向き")
    if sh.get("print_area"):
        info.append(f"印刷範囲 {sh['print_area'].rsplit('!', 1)[-1].replace('$', '')}")
    if sh.get("hidden_cols"):
        info.append("非表示の列 " + "、".join(sh["hidden_cols"]))
    if sh.get("hidden_rows"):
        info.append("非表示の行 " + "、".join(sh["hidden_rows"]))
    if sh.get("hidden_cols") or sh.get("hidden_rows"):
        info.append("非表示の行・列の字は、読み順の行の〔非表示〕のあとに番地つきで出しています")
    if sh.get("protected"):
        info.append("シート保護あり（入力欄だけ編集できます）")
    if sh.get("hidden_by_format"):
        info.append(f"条件付き書式でいまは見えない文字 {len(sh['hidden_by_format'])} か所（読み順の行の"
                    "〔条件付き書式でいまは見えない〕のあとに番地つきで出しています）")
    if sh.get("unshown_formats"):
        u = sh["unshown_formats"]
        info.append(f"表示形式を再現できないセル {len(u)} か所（{_compact([x['cell'] for x in u])}）："
                    "Excel の見える字ではなく、元の値を〔表示形式「…」は再現できません：元の値 …〕と出しています")
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
                "ファイルの作りに印のある空のマス（罫線で囲まれたもの［左右か上下の両方・下線］・塗り・結合のあるもの、入力欄のある表の行の空のマス）と、"
                "数式が読むセル（空か数。［B5＝値？］はいまの値）と、文字の中の空きだけです。"
                "書く所かは、根拠とその行（読み順）を読んで判断し、書くなら入力欄と同じく番地で書きます。"
                "近くの字・推定した見出しは、JSON の窓の fields（evidence・label）にあります。", ""]
        weak = [dict(f, id=f["cell"]) for f in agent if f["kind"] == "inline" and f.get("weak")]
        texts = [f for f in agent if f["kind"] == "inline" and not f.get("weak")]
        if weak:
            out += [weak_line(weak), ""]
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
    out += ["### 読み順（上から）", ""] + _md_reading(sh)
    out.append("")
    notes = sh.get("notes", [])
    shapes = [s for s in sh.get("shapes", []) if s["text"] or s["alt"] or s["kind"] == "picture"]
    if notes or shapes or sh.get("header_footer"):
        out += ["### 図形・メモ・ヘッダー", ""]
        for s in shapes:
            desc = s["text"] or s["alt"] or ("（飾り）" if s.get("decorative") else
                                              "**説明（代わりの字）がありません**" if s["kind"] == "picture" else "（文字も代替テキストもありません）")
            where = f"{s['anchor']}付近" if s["anchor"] else "位置がファイルにない（グループの中の図形など）"
            out.append(f"- {s['kind']}（{where}）: {_md(desc)}")
        for n in notes:
            out.append(f"- メモ（{n['cell']}）: {_md(n['text'])}")
        for h in sh.get("header_footer", []):
            out.append(f"- ヘッダー・フッター: {_md(h)}")
        out.append("")
    return out + sheet_markdown(sh)


def _md_reading(sh: dict) -> list[str]:
    out = []
    for line in sh.get("reading", []):
        head = f"{line['row']}" if "row" in line else line["rows"].replace("-", "〜")
        out.append(f"- {head}: {line['text']}")
    return out


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
    if f.get("locked"):
        k += "（保護されたセル：書けない／パスワードが要る）"
    if f.get("after"):
        k += f"（すぐ後ろのマスの字「{f['after']}」）"
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
        if "formula_ref" in whys:
            parts.append("数式が読む")
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
