"""元のファイルを正本にしたまま、セルに値を書き込む。

シートのXMLを読み直して保存し直すことはしない。書き込むセルと、その値に依存して
計算結果が変わる数式セルの <v> だけを、バイト列の上で差し替える。
それ以外のバイトは1バイトも変えない。
"""

from __future__ import annotations

import hashlib
import math
import re
from dataclasses import dataclass, field
from xml.sax.saxutils import escape

from ..formula.dates import date_to_serial
from ..formula.numfmt import is_date_format
from ..formula.parser import addr, parse, parse_cell
from ..formula.values import ExcelError, Unsupported, compare, is_number, number_to_text, parse_number, to_text
from .calc import Calc, _run_deep, recalc, values_equal
from .cellcheck import both_controls_checked as _both_controls_checked
from .cellcheck import fit_issues, written_format
from .workbook import Control, Workbook


class EditError(ValueError):
    pass


@dataclass
class Patch:
    part: str
    sheet: str
    address: str
    kind: str  # input / macro（セルの値を変えたときに動くマクロが書いた）/ formula / control / workbook
    before: object
    after: object
    span: tuple[int, int] = (0, 0)  # 元のXMLで置き換えた範囲（バイト）
    new_len: int = 0


@dataclass
class EditResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    unverified: list[str] = field(default_factory=list)
    full_calc_on_load: bool = False
    original_sha256: str = ""
    sha256: str = ""
    unchanged_parts: int = 0
    changed_parts: list[str] = field(default_factory=list)
    macro: dict | None = None  # 動かしたマクロ：{"ran": [...], "messages": [...], "files": {...}, "error": ...}
    text_written: dict = field(default_factory=dict)  # 指示の鍵 → (セルの id, 書いた後のセルの字)：字のセルの一部を書き換えた所
    replaced_text: list = field(default_factory=list)  # 印刷された字を、明示なしに置き換えたセル（証明で想定外の変更として示す）
    anchored: dict = field(default_factory=dict)  # 字を目印にした指示の鍵（@…）→ 書いたセルの id
    written_formatting: list = field(default_factory=list)  # 書いたセルの書式（取り消し線・見えない字など）：[{id, formatting}]
    fit: list = field(default_factory=list)  # 書いた字がセルに収まらず切れて見える見込み：[{id, message}]
    exclusive: list = field(default_factory=list)  # 2つのうち1つを選ぶ所の両方に印をつけた（警告の言葉）
    rich_cells: set = field(default_factory=set)  # 字ごとの書式を残して書いたセル (シート, 番地)（証明で、書式が残ったかを確かめる）

    @property
    def allowed_inputs(self) -> set[tuple[str, str]]:
        """書き込みを指示したセルと、それに応じてマクロが書いたセル（証明で、変わってよいセル）。"""
        return {(p.sheet, p.address) for p in self.patches if p.kind in ("input", "macro")}

    def summary(self) -> dict:
        return {
            "inputs": sum(1 for p in self.patches if p.kind == "input"),
            "macro_cells": sum(1 for p in self.patches if p.kind == "macro"),
            "formulas": sum(1 for p in self.patches if p.kind == "formula"),
            "controls": sum(1 for p in self.patches if p.kind == "control"),
            "changed_parts": self.changed_parts,
            "unchanged_parts": self.unchanged_parts,
            "warnings": self.warnings,
            "unverified": self.unverified,
            "full_calc_on_load": self.full_calc_on_load,
            "original_sha256": self.original_sha256,
            "sha256": self.sha256,
        }


# ---------------------------------------------------------------- 入力値の解釈
def resolve_target(wb: Workbook, target: str, *, inner_merge: bool = True) -> tuple[str, int, int]:
    """'シート!B5' や 'B5'（先頭シート）を、結合セルの左上に寄せて返す。
    inner_merge=False（書き込みの指示）：結合したセルの左上でないセル（B9:D9 の C9）は、左上に寄せずに EditError にする
    （Excel でも、結合の中の隠れたセルには入力できない。どのセルのつもりかを確かめてもらう）。"""
    if "!" in target:
        sheet, a = target.rsplit("!", 1)
        sheet = sheet.strip("'")
    else:
        sheet, a = wb.sheets[0].name, target
    try:
        sh = wb.sheet(sheet)
    except KeyError:
        raise EditError(f"シート「{sheet}」がありません")
    a_ = a.replace("$", "").strip()
    if ":" in a_:
        # 結合したセルの範囲（窓に「E23:J23」と出る）：その結合の範囲とちょうど同じときだけ、左上のセルとして受け取る
        try:
            (r, c), (r2, c2) = (parse_cell(x.strip()) for x in a_.split(":", 1))
        except ValueError:
            raise EditError(f"「{target}」はセルの番地ではありません") from None
        m = sh.merged_at(r, c)
        if m is None or (min(r, r2), min(c, c2), max(r, r2), max(c, c2)) != tuple(m):
            raise EditError(f"「{target}」は結合したセルの範囲ではありません（範囲には書けません。1つのセルの番地を指定してください）")
        a = a_.split(":", 1)[0].strip()
    try:
        r, c = parse_cell(a.replace("$", "").strip())
    except ValueError:
        raise EditError(f"「{target}」はセルの番地ではありません")
    m = sh.merged_at(r, c)
    if m is not None and (r, c) != (m[0], m[1]) and not inner_merge:
        rng = f"{addr(m[0], m[1])}:{addr(m[2], m[3])}"
        raise EditError(f"{sh.name}!{addr(r, c)} は結合したセル {rng} の中の、隠れたセルです（Excel でも入力できません）。"
                        f"結合したセルに書くなら {sh.name}!{addr(m[0], m[1])} に書いてください")
    if m is not None and (r, c) != (m[0], m[1]):
        anchor = sh.cells.get((m[0], m[1]))
        if anchor is not None and isinstance(anchor.value, str) and anchor.value.strip() and not anchor.has_formula:
            from .window import input_cells

            if (m[0], m[1]) not in input_cells(wb, sh):
                raise EditError(f"{sh.name}!{a} は見出し「{anchor.value.strip()}」と結合されたセルです（書き込むと見出しが消えます）")
        r, c = m[0], m[1]
    return sh.name, r, c


def find_control(wb: Workbook, target: str) -> Control | None:
    """'シート!Check Box 1' のようにコントロールの名前で指定されたものを探す。"""
    if "!" not in target:
        return None
    sheet, name = target.rsplit("!", 1)
    sheet = sheet.strip("'")
    if not wb.has_sheet(sheet):
        return None
    for ctl in wb.sheet(sheet).controls:
        if ctl.name == name.strip():
            return ctl
    return None


def checkbox_links(wb: Workbook) -> dict[tuple[str, int, int], list[tuple[str, Control]]]:
    """チェックボックスがつながっているセル → (コントロールのあるシート, コントロール)。"""
    out: dict[tuple[str, int, int], list[tuple[str, Control]]] = {}
    for sh in wb.sheets:
        for ctl in sh.controls:
            if ctl.kind == "CheckBox" and ctl.link_cell and wb.has_sheet(ctl.link_cell[0]):
                key = (wb.sheet(ctl.link_cell[0]).name, ctl.link_cell[1], ctl.link_cell[2])
                out.setdefault(key, []).append((sh.name, ctl))
    return out


# Excelが手入力で数値として受け取る全角の文字（数字・符号・小数点・桁区切り・％・円記号）だけを半角にする。
# NFKC は ① を 1 に、㈱ を (株) に変えるなど、文字そのものを変えてしまうので使わない。
_WIDE = str.maketrans("０１２３４５６７８９＋－．，％￥　／：", "0123456789+-.,%¥ /:")
_WIDE_MINUS = str.maketrans({"−": "-", "ー": None})


def half_width(s: str) -> str:
    return s.translate(_WIDE)


_TRUE = {"TRUE", "1", "ON", "CHECKED", "YES", "☑", "✓", "✔", "はい", "あり", "有"}
_FALSE = {"FALSE", "0", "OFF", "UNCHECKED", "NO", "☐", "□", "いいえ", "なし", "無"}


def as_bool(value) -> bool | None:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return bool(value) if value in (0, 1) else None
    t = half_width(str(value)).strip().upper()
    if t in _TRUE:
        return True
    if t in _FALSE or t == "":
        return False
    return None


_WESTERN = re.compile(r"(\d{4})\s*[/\-.年]\s*(\d{1,2})\s*[/\-.月]\s*(\d{1,2})\s*日?")
_WAREKI = re.compile(r"(明治|大正|昭和|平成|令和|[MTSHR])\s*(元|\d{1,2})\s*[年/\-.]\s*(\d{1,2})\s*[月/\-.]\s*(\d{1,2})\s*日?", re.I)
# 年が先頭でない、数字だけの日付（03/11/2026・3.11.26・11-03-2026）：日・月の順が地域で違う
_DAY_MONTH = re.compile(r"\d{1,2}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{1,4}\.?")
_NUMBER = re.compile(r"([+-])?[¥$]?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE]([+-]?\d+))?(%)?")


def parse_date_text(s: str):
    """「2026/4/1」「2026年4月1日」「令和8年4月1日」「R8.4.1」などを日付にする。読めなければ None。"""
    import datetime as dt

    from ..formula.dates import ERAS

    t = half_width(s).strip()
    m = _WESTERN.fullmatch(t)
    try:
        if m:
            return dt.date(int(m[1]), int(m[2]), int(m[3]))
        m = _WAREKI.fullmatch(t)
        if m:
            era = m[1].upper()
            start = next(e[2] for e in ERAS if era in (e[0], e[1]))
            n = 1 if m[2] == "元" else int(m[2])
            return dt.date(start.year + n - 1, int(m[3]), int(m[4]))
    except (ValueError, StopIteration):
        return None
    return None


def _is_percent(fmt: str) -> bool:
    return "%" in re.sub(r'"[^"]*"|\\.', "", fmt)


def interpret_input(wb: Workbook, sheet: str, r: int, c: int, value, *, refuse_ambiguous: bool = True) -> tuple[object, list[str]]:
    """書き込む値を、Excelでそのセルに手で入力したときと同じ規則で、数値・日付・論理値・文字にする。

    日付は、年が先頭の形（ISO の 2026-11-03・2026/11/3・2026年11月3日）と和暦（令和8年11月3日・R8.11.3）を読み、セルの
    表示形式で表示する。日付の表示形式のセルに、年が先頭でない数字の日付（03/11/2026）を書くと、日・月の順が地域で違うので
    推測せず EditError にする（refuse_ambiguous=False：マクロが書いた字などは、文字のまま書き、そのことを返す）。

    セルの書式は変えない。Excelなら書式まで変える入力（日付を標準のセルに入れるなど）は、
    文字のまま書き込み、その旨を返す。
    """
    where = f"{sheet}!{addr(r, c)}"
    cell = wb.sheet(sheet).cells.get((r, c))
    fmt = wb.style(cell.style).num_fmt if cell else "General"
    notes: list[str] = []
    if value is None:
        return None, notes
    if fmt == "@":
        # 文字列の書式のセルは、Excelでも入力したとおりの文字になる
        if isinstance(value, bool):
            return ("TRUE" if value else "FALSE"), notes
        if isinstance(value, (int, float)):
            return number_to_text(float(value)), notes
        s = str(value)
        return (s[1:] if s.startswith("'") else s), notes
    if isinstance(value, bool):
        return value, notes
    if isinstance(value, (int, float)):
        x = float(value)
        if not math.isfinite(x):
            raise EditError(f"{where}: 値 {value!r} は数ではありません（無限大・NaN は書けません）")
        if _general(fmt) and x == int(x) and "E" in _general_shown(x):
            # 12 桁以上の整数（個人番号など）：標準の表示形式では「1.23457E+11」と表示され、番号が読めない
            digits = str(int(value)) if isinstance(value, int) else number_to_text(x)
            return _shown_differs(wb, sheet, r, c, where, digits, x, _general_shown(x), notes)
        if _is_percent(fmt) and abs(x) >= 1:
            notes.append(f"{where}: パーセントの書式のセルに {number_to_text(x)} を書き込みました（{number_to_text(x * 100)}% と表示されます。"
                         f"{number_to_text(x)}% のつもりなら {number_to_text(x / 100)} か「{number_to_text(x)}%」と指定してください）")
        return x, notes
    s = str(value)
    if s.startswith("'"):
        return s[1:], notes
    if s.strip() == "":
        return None, notes
    t = half_width(s).strip()
    if t.upper() in ("TRUE", "FALSE"):
        return t.upper() == "TRUE", notes
    d = parse_date_text(s)
    if d is not None:
        if is_date_format(fmt):
            return float(date_to_serial(d, wb.date1904)), notes
        notes.append(f"{where}: 「{s}」は日付として読めますが、このセルの表示形式は日付ではないので、文字として書き込みました")
        return s, notes
    if is_date_format(fmt) and _DAY_MONTH.fullmatch(t):
        # 03/11/2026 のように年が先頭でない日付：日・月の順（dd/mm）か月・日の順（mm/dd）かは地域で違い、Excel も開く PC の
        # 地域の設定で読み分ける。todoku は推測しない
        msg = (f"{where}: 「{s}」は、日と月のどちらが先かが地域で違う書き方です。日付は ISO の形（2026-11-03 のように年-月-日）で"
               "指定してください（セルの表示形式どおりに表示されます）")
        if refuse_ambiguous:
            raise EditError(msg)
        notes.append(msg + "。文字として書き込みました")
        return s, notes
    m = _NUMBER.fullmatch(t)
    if m:
        digits = re.sub(r"\D", "", m[2])
        if re.fullmatch(r"0\d+", m[2]):
            notes.append(f"{where}: 先頭の 0 を残すため、「{s}」を文字として書き込みました")
            return s, notes
        if len(digits.lstrip("0")) > 15:
            notes.append(f"{where}: 15桁を超える数字はExcelでは末尾が0に変わるため、「{s}」を文字として書き込みました")
            return s, notes
        try:
            x = float(m[2].replace(",", "") + (f"e{m[3]}" if m[3] else ""))
        except (OverflowError, ValueError):
            x = math.inf
        if not math.isfinite(x) or (x == 0 and digits.strip("0")):
            # 1e400 など：Excel で扱える数（±9.99E+307 まで）の外。Excel でも字として受け取る
            notes.append(f"{where}: 「{s}」は Excel で扱える数の大きさ（およそ ±1E+308 から ±1E-307 まで）の外なので、"
                         "文字として書き込みました")
            return s, notes
        if m[1] == "-":
            x = -x
        if m[4]:
            x /= 100
        if _general(fmt) and (m[4] or "¥" in t or "$" in t or "," in m[2]):
            # 「¥1,000」「$100」「50%」「1,000」：Excel は手で入力すると表示形式を通貨・パーセント・桁区切りに変える。
            # todoku はセルの書式を変えないので、標準の表示形式のままでは「1000」「0.5」と、入力した字と違って見える
            return _shown_differs(wb, sheet, r, c, where, s, x, _general_shown(x), notes)
        if _general(fmt) and not m[3] and "E" in _general_shown(x):
            # 12 桁以上の数字（個人番号など）：標準の表示形式では「1.23457E+11」と表示され、番号が読めない
            return _shown_differs(wb, sheet, r, c, where, s, x, _general_shown(x), notes)
        elif _is_percent(fmt):
            # Excelの「パーセント形式での自動入力」：1以上の数は % として受け取る
            if abs(x) >= 1:
                x /= 100
        if t != s.strip() and re.search(r"[０-９]", s):
            notes.append(f"{where}: 全角の数字「{s}」を数値 {number_to_text(x)} として書き込みました（Excelと同じ扱い）")
        return x, notes
    if is_date_format(fmt) and re.search(r"\d", t):
        notes.append(f"{where}: このセルの表示形式は日付・時刻ですが、「{s}」は日付として読めないので、文字として書き込みました"
                     "（日付なら ISO の形 2026-11-03 で指定してください）")
    return s, notes


MAX_CELL_TEXT = 32767  # Excel のセルに入る字の数の上限


def _general(fmt: str) -> bool:
    return (fmt or "General").strip().lower() in ("general", "")


def _general_shown(x: float) -> str:
    from ..formula.numfmt import format_value

    try:
        return format_value(x, "General")
    except Exception:  # noqa: BLE001
        return number_to_text(x)


def _shown_differs(wb, sheet: str, r: int, c: int, where: str, typed: str, x: float, shown: str, notes: list):
    """標準の表示形式のセルに書く数が、書いた字と違う字で表示される（「¥1,000」が「1000」、12 桁の番号が「1.23457E+11」）。
    数式が参照するセル（か、参照するかが決められないとき）と、数の入力規則のあるセルは、計算に使う数として書き、見え方を伝える。
    ほかは、書いた字のまま見えるように字として書き、そのことを伝える（どちらも黙って見え方を変えない）。"""
    if _numeric_use(wb, sheet, r, c):
        notes.append(f"{where}: 「{typed}」を数 {number_to_text(x)} として書き込みました（このセルは数式から参照される・数の入力規則の"
                     f"あるセルなので、数にしました）。セルの表示形式が標準なので「{shown}」と表示されます（Excel で手で入力すると"
                     "表示形式も変わりますが、todoku はセルの書式を変えません）。本人に、印刷される字が「" + shown + "」になることを伝えてください")
        return x, notes
    notes.append(f"{where}: 「{typed}」は、数として書くと標準の表示形式で「{shown}」と表示され、書いた字と違って見えるので、"
                 "書いたとおりの字として書き込みました（計算には使われていないセルです。数として書くなら "
                 f"{number_to_text(x)} と指定してください）")
    return typed, notes


def _numeric_use(wb, sheet: str, r: int, c: int) -> bool:
    """セルの値を、数式（名前・テーブルの参照を含む）か数の入力規則が使うか。決められない（INDIRECT・OFFSET・読めない数式が
    ある）ときも、使うものとする。"""
    sh = wb.sheet(sheet)
    if any(v.covers(r, c) and v.type in ("whole", "decimal", "date", "time") for v in getattr(sh, "validations", ())):
        return True
    refs = _formula_refs(wb)
    if refs is None:
        return True
    return any(s_.lower() == sh.name.lower() and r1 <= r <= r2 and c1 <= c <= c2 for s_, r1, c1, r2, c2 in refs)


def _formula_refs(wb):
    """ブックの数式が参照する範囲 [(シート, r1, c1, r2, c2)]。参照が決められない数式があれば None。ブックごとに一度だけ求める。"""
    cache = getattr(wb, "_todoku_formula_refs", False)
    if cache is not False:
        return cache
    from ..formula.parser import NameNode, RefNode, StructRef, iter_funcs, iter_refs
    from .calc import formula_cells

    out: list = []

    def walk(ast, here: str, dr: int, dc: int, depth: int = 0) -> bool:
        if depth > 8:
            return False
        if any(str(f).upper() in ("INDIRECT", "OFFSET") for f in iter_funcs(ast)):
            return False
        for ref in iter_refs(ast):
            if isinstance(ref, RefNode):
                rr = ref.shifted(dr, dc)
                out.append((rr.sheet or here, min(rr.r1, rr.r2), min(rr.c1, rr.c2), max(rr.r1, rr.r2), max(rr.c1, rr.c2)))
            elif isinstance(ref, NameNode):
                d = wb.defined_name(ref.name, ref.sheet or here) if hasattr(wb, "defined_name") else None
                if d is None:
                    continue  # 名前のない参照（#NAME?）
                if d.ast is None or not walk(d.ast, d.sheet_scope or here, 0, 0, depth + 1):
                    return False
            elif isinstance(ref, StructRef):
                tables = list(getattr(wb, "tables", {}).values())
                hit = [x for x in tables if (x.name.lower() == ref.table.lower() if ref.table else x.sheet == here)]
                if not hit:
                    return False
                out.extend((t.sheet, t.r1, t.c1, t.r2, t.c2) for t in hit)
        return True

    ok = True
    try:
        for cell in formula_cells(wb):
            if cell.array_master is not None:
                continue
            if cell.ast is None or not walk(cell.ast, cell.sheet, cell.dr, cell.dc):
                ok = False
                break
    except Exception:  # noqa: BLE001  確かめられない：使うものとする
        ok = False
    res = out if ok else None
    try:
        wb._todoku_formula_refs = res
    except Exception:  # noqa: BLE001
        pass
    return res


def coerce_input(wb: Workbook, sheet: str, r: int, c: int, value):
    return interpret_input(wb, sheet, r, c, value)[0]


@dataclass
class Issue:
    cell: str
    message: str
    blocking: bool  # Excelでも受け付けない（入力規則のエラーメッセージが「停止」）


def _rule_eval(calc: Calc, sheet: str, v, r: int, c: int, formula: str):
    """入力規則の数式を、そのセルの位置で計算する（相対参照は入力規則の範囲の左上からずらす）。"""
    from ..formula.evaluator import Evaluator

    r0, c0 = v.origin
    ev = Evaluator(calc, sheet, r, c, r - r0, c - c0, array=True)
    return ev, ev.run(parse(formula.lstrip("=")))


def list_options(wb: Workbook, sheet: str, formula1: str, r: int | None = None, c: int | None = None,
                 rule=None, calc: Calc | None = None) -> list | None:
    """入力規則（リスト）の選択肢。範囲・名前・INDIRECT などは、そのセルの位置で計算する。"""
    f = formula1.strip()
    if f.startswith('"') and f.endswith('"'):
        items = [x.strip() for x in f[1:-1].split(",")]
        return [parse_number(x) if parse_number(x) is not None else x for x in items]
    try:
        calc = calc or Calc(wb)
        if rule is not None and r is not None:
            ev, v = _rule_eval(calc, sheet, rule, r, c, f)
        else:
            from ..formula.evaluator import Evaluator

            ev = Evaluator(calc, sheet, r or 1, c or 1, array=True)
            v = ev.run(parse(f.lstrip("=")))
        vals = [x for x in ev.values(v) if x is not None and x != "" and not isinstance(x, ExcelError)]
        return vals or None
    except (Unsupported, Exception):  # noqa: BLE001
        return None


_OPS = {
    "between": lambda x, lo, hi: lo <= x <= hi,
    "notBetween": lambda x, lo, hi: not (lo <= x <= hi),
    "equal": lambda x, lo, hi: x == lo,
    "notEqual": lambda x, lo, hi: x != lo,
    "greaterThan": lambda x, lo, hi: x > lo,
    "lessThan": lambda x, lo, hi: x < lo,
    "greaterThanOrEqual": lambda x, lo, hi: x >= lo,
    "lessThanOrEqual": lambda x, lo, hi: x <= lo,
}
_OP_TEXT = {"between": "{lo}〜{hi}", "notBetween": "{lo}〜{hi} 以外", "equal": "{lo}", "notEqual": "{lo} 以外",
            "greaterThan": "{lo} より大きい", "lessThan": "{lo} より小さい", "greaterThanOrEqual": "{lo} 以上",
            "lessThanOrEqual": "{lo} 以下"}


def _rule_number(calc, sheet, v, r, c, f):
    if f is None:
        return None
    n = parse_number(f)
    if n is not None:
        return n
    try:
        ev, x = _rule_eval(calc, sheet, v, r, c, f)
        x = ev.scalar(x) if hasattr(ev, "scalar") else x
        return float(x) if is_number(x) else None
    except Exception:  # noqa: BLE001
        return None


def rule_issues(wb: Workbook, sheet: str, r: int, c: int, value, calc: Calc | None = None) -> list[Issue]:
    """入力規則に照らした問題。calc に書き込み後の値を入れておくと、ほかの欄に連動する規則も正しく確かめられる。"""
    sh = wb.sheet(sheet)
    where = f"{sheet}!{addr(r, c)}"
    calc = calc or Calc(wb)
    rules = [v for v in sh.validations if v.covers(r, c) and v.type not in ("any", None)]
    if len(rules) > 1:
        per = [_issues_for(wb, sheet, r, c, value, calc, [v]) for v in rules]
        if any(not x for x in per):
            # どれか1つの規則には合う：重なった規則のどれをExcelが使うかは決まらないので、止めずに知らせる
            return [Issue(where, f"（入力規則が {len(rules)} つ重なっています。Excelがどれを使うかはファイルからは決まりません）"
                                 f"{i.message}", False) for x in per for i in x]
        return [i for x in per for i in x]
    return _issues_for(wb, sheet, r, c, value, calc, rules)


def _issues_for(wb: Workbook, sheet: str, r: int, c: int, value, calc: Calc, rules) -> list[Issue]:
    where = f"{sheet}!{addr(r, c)}"
    out: list[Issue] = []
    for v in rules:
        if value is None or value == "":
            continue  # Excelでも、セルを消す操作は入力規則にかからない
        extra = f"（様式の説明: {v.error}）" if v.error else ""
        if v.type == "list" and v.formula1:
            opts = list_options(wb, sheet, v.formula1, r, c, v, calc)
            if opts is None:
                out.append(Issue(where, f"{where}: 選択肢を確かめられませんでした（{v.formula1}）。連動する欄を先に入れると確かめられます", False))
            elif not any(_eq(value, o) or to_text(value) == to_text(o) for o in opts):
                # 文字列の書式のセルに「0」を入れたときも、一覧の 0 と同じものとして受け付ける（見た目の文字で照らす）
                shown = "、".join(f"「{to_text(o)}」" for o in opts[:12]) + ("…" if len(opts) > 12 else "")
                out.append(Issue(where, f"{where}: 「{to_text(value)}」は選択肢にありません。選べるのは {shown}{extra}", v.blocks))
        elif v.type in ("whole", "decimal", "date", "time"):
            if not is_number(value) or isinstance(value, bool):
                kind = {"whole": "整数", "decimal": "数値", "date": "日付", "time": "時刻"}[v.type]
                out.append(Issue(where, f"{where}: {kind}を入力してください{extra}", v.blocks))
                continue
            x = float(value)
            if v.type == "whole" and x != int(x):
                out.append(Issue(where, f"{where}: 整数を入力してください{extra}", v.blocks))
                continue
            lo = _rule_number(calc, sheet, v, r, c, v.formula1)
            hi = _rule_number(calc, sheet, v, r, c, v.formula2)
            op = v.operator or "between"
            if lo is None or (op in ("between", "notBetween") and hi is None):
                continue
            if not _OPS.get(op, lambda *a: True)(x, lo, hi):
                show = number_to_text
                if v.type == "date":
                    def show(n, _b=wb.date1904):  # 日付の規則：シリアル値ではなく、ブックの数え方で日付にして示す
                        from ..formula.numfmt import format_value

                        try:
                            return format_value(n, "yyyy/m/d", _b)
                        except Exception:  # noqa: BLE001
                            return number_to_text(n)
                rng = _OP_TEXT.get(op, op).format(lo=show(lo), hi=show(hi) if hi is not None else "")
                out.append(Issue(where, f"{where}: 値 {show(x)} は入力規則（{rng}）の範囲外です{extra}", v.blocks))
        elif v.type == "textLength":
            lo = _rule_number(calc, sheet, v, r, c, v.formula1)
            hi = _rule_number(calc, sheet, v, r, c, v.formula2)
            op = v.operator or "between"
            n = len(to_text(value))
            if lo is None or (op in ("between", "notBetween") and hi is None):
                continue
            if not _OPS.get(op, lambda *a: True)(n, lo, hi):
                rng = _OP_TEXT.get(op, op).format(lo=number_to_text(lo), hi=number_to_text(hi) if hi is not None else "")
                out.append(Issue(where, f"{where}: 文字数 {n} は入力規則（{rng}文字）に合いません{extra}", v.blocks))
        elif v.type == "custom" and v.formula1:
            try:
                ev, x = _rule_eval(calc, sheet, v, r, c, v.formula1)
                x = ev.scalar(x) if hasattr(ev, "scalar") else x
                ok = bool(x) if isinstance(x, (bool, int, float)) else None
            except Exception:  # noqa: BLE001
                ok = None
            if ok is False:
                out.append(Issue(where, f"{where}: 入力規則（={v.formula1}）を満たしません{extra}", v.blocks))
            elif ok is None:
                out.append(Issue(where, f"{where}: 入力規則（={v.formula1}）を確かめられませんでした", False))
    return out


def check_validation(wb: Workbook, sheet: str, r: int, c: int, value) -> list[str]:
    """入力規則に照らして、問題があれば説明を返す（ほかの欄への書き込みは考えない簡易版）。"""
    return [i.message for i in rule_issues(wb, sheet, r, c, value)]


def _eq(a, b) -> bool:
    try:
        return compare(a, b) == 0
    except ExcelError:
        return False


# ---------------------------------------------------------------- XMLのバイト列を差し替える
def _num_text(x: float) -> str:
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return repr(float(x))


def _row_re(prefix: bytes, r: int) -> re.Pattern:
    p = re.escape(prefix)
    return re.compile(
        rb"<" + p + rb"row\b(?=[^>]*\br=\"" + str(r).encode() + rb"\")[^>]*?(?:/>|>.*?</" + p + rb"row>)",
        re.S,
    )


def _prefix(xml: bytes) -> bytes:
    m = re.search(rb"<(\w+:)?sheetData\b", xml)
    if not m:
        raise EditError("sheetData が見つかりません")
    return m[1] or b""


def _attrs(tag_open: bytes) -> list[tuple[bytes, bytes]]:
    return re.findall(rb'([\w:]+)="([^"]*)"', tag_open)


def _open_tag(prefix: bytes, name: bytes, attrs: list[tuple[bytes, bytes]]) -> bytes:
    return b"<" + prefix + name + b"".join(b" " + k + b'="' + v + b'"' for k, v in attrs)


def _value_xml(prefix: bytes, value) -> tuple[bytes | None, bytes]:
    """(t属性, 中身) を返す。"""
    p = prefix
    if value is None:
        return None, b""
    if isinstance(value, bool):
        return b"b", b"<" + p + b"v>" + (b"1" if value else b"0") + b"</" + p + b"v>"
    if isinstance(value, ExcelError):
        return b"e", b"<" + p + b"v>" + escape(value.code).encode() + b"</" + p + b"v>"
    if is_number(value):
        return None, b"<" + p + b"v>" + _num_text(float(value)).encode() + b"</" + p + b"v>"
    from .richtext import escape_text

    s = escape_text(str(value))  # 「_x0041_」の形の字は「_x005F_x0041_」にして、字のまま残す（ECMA-376 22.9.2.19）
    sp = b' xml:space="preserve"' if s != s.strip() or b"\n" in s else b""
    return b"inlineStr", b"<" + p + b"is><" + p + b"t" + sp + b">" + s + b"</" + p + b"t></" + p + b"is>"


def _with_t(attrs, t: bytes | None):
    out = [(k, v) for k, v in attrs if k != b"t"]
    if t is not None:
        # t は r, s の後ろに置く（順序はExcelの慣習に合わせる）
        idx = max([i + 1 for i, (k, _) in enumerate(out) if k in (b"r", b"s")] or [len(out)])
        out.insert(idx, (b"t", t))
    return out


def _input_cell_xml(prefix: bytes, existing: bytes | None, a: str, value, style: int = 0, rich: bytes | None = None) -> bytes:
    if existing is not None:
        head = re.match(rb"<[\w:]*c\b[^>]*?(?=/?>)", existing)[0]
        attrs = _attrs(head)
    else:
        attrs = [(b"r", a.encode())] + ([(b"s", str(style).encode())] if style else [])
    if rich is not None:
        # 字ごとの書式・ふりがなを残した字（xlsx/richtext.splice の <is> の中身）
        t, inner = b"inlineStr", b"<" + prefix + b"is>" + rich + b"</" + prefix + b"is>"
    else:
        t, inner = _value_xml(prefix, value)
    attrs = _with_t(attrs, t)
    if not inner:
        return _open_tag(prefix, b"c", attrs) + b"/>"
    return _open_tag(prefix, b"c", attrs) + b">" + inner + b"</" + prefix + b"c>"


def _formula_cell_xml(prefix: bytes, existing: bytes, value) -> bytes:
    """数式は残したまま、計算結果 <v> だけを差し替える。"""
    m = re.match(rb"(<[\w:]*c\b[^>]*?)(/?>)", existing)
    attrs = _attrs(m[1])
    if isinstance(value, str):
        from .richtext import escape_text

        t = b"str"
        vxml = b"<" + prefix + b"v>" + escape_text(value) + b"</" + prefix + b"v>"
    else:
        t, vxml = _value_xml(prefix, value)
    body = existing[m.end():]
    if m[2] == b"/>":
        # 値のないセル（配列数式の結果のうち、空だったところ）：中身を作る
        if b"<" + prefix + b"f" in existing:
            raise EditError("数式セルの形式が想定外です")
        body = b"</" + prefix + b"c>"
    body = re.sub(rb"<" + re.escape(prefix) + rb"v>.*?</" + re.escape(prefix) + rb"v>|<" + re.escape(prefix) + rb"v\s*/>", b"", body, count=1, flags=re.S)
    fm = re.search(rb"<" + re.escape(prefix) + rb"f\b[^>]*?(?:/>|>.*?</" + re.escape(prefix) + rb"f>)", body, re.S)
    if fm:
        body = body[:fm.end()] + vxml + body[fm.end():]
    else:
        body = vxml + body
    return _open_tag(prefix, b"c", _with_t(attrs, t)) + b">" + body


def _rc(a: str) -> tuple[int, int]:
    from ..formula.parser import parse_cell

    return parse_cell(a)


def calc_chain_entries(xml: bytes) -> list[tuple[int, str]]:
    """calcChain.xml の記録：(シートの番号 i, セル)。i を省いた記録は、前の記録の i を引き継ぐ（ECMA-376 18.6.1）。"""
    out, cur = [], 0
    for m in re.finditer(rb"<(?:\w+:)?c\b[^>]*?/?>", xml):
        rm = re.search(rb'\br="([A-Z]+\d+)"', m[0])
        im = re.search(rb'\bi="(\d+)"', m[0])
        if im:
            cur = int(im[1])
        if rm:
            out.append((cur, rm[1].decode()))
    return out


def drop_calc_chain(xml: bytes, gone: set[tuple[int, str]]) -> tuple[bytes, int]:
    """calcChain.xml から、gone の記録だけを取り除く。取り除く記録が i（または l：新しい段）を持ち、次の記録がそれを
    省いているときは、次の記録に移す（ほかの記録の意味を変えない）。"""
    out = bytearray()
    last = 0
    cur = 0
    carry_i = None
    carry_l = False
    n = 0
    for m in re.finditer(rb"<(?:\w+:)?c\b[^>]*?/?>", xml):
        tag = m[0]
        rm = re.search(rb'\br="([A-Z]+\d+)"', tag)
        im = re.search(rb'\bi="(\d+)"', tag)
        if im:
            cur = int(im[1])
        if rm and (cur, rm[1].decode()) in gone:
            out += xml[last:m.start()]
            last = m.end()
            n += 1
            if im:
                carry_i = im[1]
            if re.search(rb'\bl="(?:1|true)"', tag):
                carry_l = True
            continue
        if carry_i is not None or carry_l:
            new = tag
            if carry_i is not None and not im:
                new = re.sub(rb'(\br="[A-Z]+\d+")', rb'\1 i="' + carry_i + b'"', new, count=1)
            if carry_l and not re.search(rb'\bl="', new):
                new = re.sub(rb"(\s*/?>)$", rb' l="1"\1', new, count=1)
            out += xml[last:m.start()] + new
            last = m.end()
            carry_i, carry_l = None, False
    out += xml[last:]
    return bytes(out), n


class SheetPatcher:
    def __init__(self, xml: bytes):
        self.xml = xml
        self.prefix = _prefix(xml)
        self.edits: list[tuple[int, int, bytes]] = []
        self.inserts: list[tuple[int, int, int, bytes]] = []  # 行の中への新しいセル：(位置, 行, 列, XML)
        self.row_inserts: dict = {}  # 新しい行：((位置), 行) → {"wrap": (開き, 閉じ), "cells": {列: XML}}
        self._index: dict[str, tuple[int, int]] | None = None

    def find_cell(self, r: int, c: int) -> tuple[int, int] | None:
        if self._index is None:
            # 1回だけ走査して、セル番地→XML上の位置の表を作る（書き換えは build まで反映しないので位置は変わらない）
            p = re.escape(self.prefix)
            pat = re.compile(rb"<" + p + rb"c\b[^>]*?\br=\"([A-Z]+\d+)\"[^>]*?(?:/>|>.*?</" + p + rb"c>)", re.S)
            self._index = {m[1].decode(): (m.start(), m.end()) for m in pat.finditer(self.xml)}
        return self._index.get(addr(r, c))

    def set_input(self, r: int, c: int, value, style: int = 0, rich: bytes | None = None) -> tuple[int, int, int]:
        a = addr(r, c)
        span = self.find_cell(r, c)
        if span:
            new = _input_cell_xml(self.prefix, self.xml[span[0]:span[1]], a, value, rich=rich)
            self.edits.append((span[0], span[1], new))
            return span[0], span[1], len(new)
        # セル要素がないセルは、Excelと同じく行か列の書式を引き継いで作る
        new = _input_cell_xml(self.prefix, None, a, value, style)
        pos = self._insert_pos(r, c)
        if isinstance(pos, tuple):
            # 新しい行（または中身のない <row/>）：同じ行に入るセルは、まとめて1つの行にする（build で列の順に並べる）
            at, wrap = pos
            g = self.row_inserts.setdefault((at, r), {"wrap": wrap, "cells": {}})
            g["cells"][c] = new
            return at[0], at[1], len(new)
        self.inserts.append((pos, r, c, new))
        return pos, pos, len(new)

    def set_formula_value(self, r: int, c: int, value) -> tuple[int, int, int]:
        span = self.find_cell(r, c)
        if not span:
            raise EditError(f"{addr(r, c)} の数式セルが見つかりません")
        new = _formula_cell_xml(self.prefix, self.xml[span[0]:span[1]], value)
        self.edits.append((span[0], span[1], new))
        return span[0], span[1], len(new)

    def _insert_pos(self, r: int, c: int):
        p = self.prefix
        rm = _row_re(p, r).search(self.xml)
        if rm:
            row = self.xml[rm.start():rm.end()]
            if row.endswith(b"/>"):
                # <row r="5" .../> を <row r="5" ...>新しいセル</row> にする
                head = row[:-2].rstrip()
                return (rm.start(), rm.end()), (head + b">", b"</" + p + b"row>")
            for cm in re.finditer(rb"<" + re.escape(p) + rb"c\b[^>]*\br=\"([A-Z]+)(\d+)\"", row):
                from ..formula.parser import col_to_num

                if col_to_num(cm[1].decode()) > c:
                    return rm.start() + cm.start()
            return rm.end() - len(b"</" + p + b"row>")
        # 行がない：行番号の順に差し込む
        sd = re.search(rb"<" + re.escape(p) + rb"sheetData\b[^>]*?(/?)>", self.xml)
        new_row_open = b"<" + p + b'row r="' + str(r).encode() + b'">'
        new_row_close = b"</" + p + b"row>"
        if sd[1] == b"/":
            return (sd.start(), sd.end()), (b"<" + p + b"sheetData>" + new_row_open, new_row_close + b"</" + p + b"sheetData>")
        for rm2 in re.finditer(rb"<" + re.escape(p) + rb"row\b[^>]*\br=\"(\d+)\"", self.xml):
            if int(rm2[1]) > r:
                return (rm2.start(), rm2.start()), (new_row_open, new_row_close)
        end = self.xml.index(b"</" + p + b"sheetData>")
        return (end, end), (new_row_open, new_row_close)

    def build(self) -> bytes:
        # 書き換え（start, end）と、差し込み（同じ位置なら、行の番号・列の番号の順）をまとめて並べる
        edits = [(s_, e_, 0, 0, new) for s_, e_, new in self.edits]
        edits += [(pos, pos, r, c, new) for pos, r, c, new in self.inserts]
        for (at, r), g in self.row_inserts.items():
            cells = b"".join(x for _c, x in sorted(g["cells"].items()))
            edits.append((at[0], at[1], r, 0, g["wrap"][0] + cells + g["wrap"][1]))
        out = bytearray()
        last = 0
        # 同じ位置では、差し込み（start == end）を、そこから始まるセルの書き換えより先に置く（差し込むセルは、
        # そのセルの前に入る。数式の計算結果の書き換えと、その前への新しいセルが重なることがある）
        for s_, e_, _r, _c, new in sorted(edits, key=lambda x: (x[0], x[1] != x[0], x[2], x[3])):
            if s_ < last:
                raise EditError("同じ位置に重なる書き換えがあります")
            out += self.xml[last:s_] + new
            last = e_
        out += self.xml[last:]
        return bytes(out)


# ---------------------------------------------------------------- チェックボックスの状態
def set_ctrlprop_checked(xml: bytes, state: str | None) -> bytes:
    """ctrlProp の checked 属性だけを変える（Excelは objectType の直後に書き、チェックなしのときは書かない）。"""
    m = re.search(rb"<(?:\w+:)?formControlPr\b[^>]*?/?>", xml)
    if not m:
        raise EditError("コントロールの設定（formControlPr）が見つかりません")
    tag = re.sub(rb'\s+checked="[^"]*"', b"", m[0])
    if state:
        tag, n = re.subn(rb'(\bobjectType="[^"]*")', rb'\1 checked="' + state.encode() + b'"', tag, count=1)
        if n == 0:
            tag = re.sub(rb"(\s*/?>)$", b' checked="' + state.encode() + rb'"\1', tag, count=1)
    return xml[:m.start()] + tag + xml[m.end():]


def set_vml_checked(vml: bytes, shape_id: str, state: str | None) -> bytes:
    """VML のチェックボックスの <x:Checked> だけを変える。"""
    sm = None
    for m in re.finditer(rb"<v:shape\b[^>]*>.*?</v:shape>", vml, re.S):
        head = re.match(rb"<v:shape\b[^>]*>", m[0])[0]
        if re.search(rb'\b(?:id|o:spid)="_x0000_s' + shape_id.encode() + rb'"', head):
            sm = m
            break
    if sm is None:
        raise EditError(f"VML に図形 {shape_id} が見つかりません")
    block = sm[0]
    block = re.sub(rb"[ \t]*<x:Checked>[^<]*</x:Checked>[ \t]*\r?\n?", b"", block)
    if state:
        val = b"1" if state == "Checked" else b"2"
        for anchor in (rb"<x:FmlaLink>", rb"<x:NoThreeD", rb"<x:FirstButton", rb"</x:ClientData>"):
            am = re.search(rb"(\r?\n)([ \t]*)" + anchor, block)
            if am:
                ins = am[1] + am[2] + b"<x:Checked>" + val + b"</x:Checked>"
                block = block[:am.start()] + ins + block[am.start():]
                break
        else:
            raise EditError(f"VML の図形 {shape_id} にチェックの状態を書く場所がありません")
    return vml[:sm.start()] + block + vml[sm.end():]


def _run_change_macros(wb, targets: dict, warnings: list, answers, inputs) -> tuple[set, dict | None]:
    """値を変えると動くマクロを、targets の順に（人が1つずつ入力したのと同じように）動かす。
    マクロが書いたセルを targets に足し、そのセルの集まりと、動かした記録を返す。動かせなければ、知らせて何も足さない。"""
    if not getattr(wb, "has_macros", False) or not targets or getattr(wb, "vba", None) is None:
        return set(), None
    from ..macros import change_event_sheets, macro_info

    info = macro_info(wb.vba, True, {sh.code_name: sh.name for sh in wb.sheets if getattr(sh, "code_name", None)})
    ev = change_event_sheets(info)
    if not ev or not (ev.get(None) or {t[0] for t in targets} & {k for k in ev if k}):
        return set(), None
    procs = sorted({p for k, ps in ev.items() if k is None or k in {t[0] for t in targets} for p in ps})
    from ..vbarun import NeedsInput, Unsupported, change_events
    from ..vbarun.parse import VbaSyntaxError

    try:
        run = _run_deep(lambda: change_events(wb, [(s, r, c, v) for (s, r, c), v in targets.items()], answers, inputs))
    except (Unsupported, NeedsInput, VbaSyntaxError) as e:
        warnings.extend(_macro_warnings(wb, targets))
        warnings.append(f"セルの値を変えたときに動くマクロ（{'、'.join(procs)}）を、ここでは再現できませんでした：{e}")
        return set(), {"ran": procs, "reproduced": False, "why": str(e)}
    written = set()
    for (s, r, c), v in run.changed.items():
        if isinstance(v, str):
            v = interpret_input(wb, s, r, c, v, refuse_ambiguous=False)[0]  # マクロが Value に入れた文字も、Excel は入力と同じように読む
        targets[(s, r, c)] = v
        written.add((s, r, c))
    if run.error:
        warnings.append(f"セルの値を変えたときに動くマクロが、Excel でも止まります：{run.error}")
    warnings += [f"マクロのメッセージ：{m}" for m in run.messages]
    warnings += [f"マクロ：{n}" for n in run.notes]
    return written, {"ran": procs, "reproduced": True, "cells": len(written), "messages": run.messages,
                     "files": {k: len(v) for k, v in run.files.items()}, "error": run.error, "volatile": run.volatile}


def _macro_warnings(wb, targets) -> list[str]:
    """セルの値を変えたときに動くマクロ（Worksheet_Change など）のあるシートへの書き込みを知らせる。"""
    if not getattr(wb, "has_macros", False) or not targets:
        return []
    from ..macros import change_event_sheets, macro_info

    project = getattr(wb, "vba", None)
    info = macro_info(project, True, {sh.code_name: sh.name for sh in wb.sheets if getattr(sh, "code_name", None)})
    ev = change_event_sheets(info)
    out = []
    for s in sorted({t[0] for t in targets} & {k for k in ev if k}):
        out.append(f"シート「{s}」には、セルの値を変えたときに動くマクロ（{'、'.join(ev[s])}）があります。"
                   "todoku はマクロを動かさないので、Excelで手で入力したときは、マクロがほかのセルも変えることがあります")
    if ev.get(None):
        out.append(f"どのシートでも、セルの値を変えたときに動くマクロ（{'、'.join(ev[None])}）があります。"
                   "todoku はマクロを動かさないので、Excelで手で入力したときと結果が違うことがあります")
    return out


# ---------------------------------------------------------------- 文字のセルの一部を書き換える（箱の字・字を目印）
def _is_box(ch: str) -> bool:
    from .. import chars

    return ch in chars.BOXES or ch in chars.MAYBE_CHECKED


def _cell_blanks(text: str):
    from .window import cell_blanks

    return cell_blanks(text)


def box_glyph(orig: str, on: bool) -> str:
    """箱の字 orig に印をつける（on）・外すときの字。Word と同じ（☐ は ☑、ほかの白い箱 □ などは ■。外すときは印のない同じ形の箱）。"""
    from ..docx.window import box_mark

    return box_mark(orig, on)


def check_box_in_text(text: str, on: bool, where: str) -> tuple[str, str]:
    """字のあるセルに true/false を書く：字の中のチェックの箱の字が1つだけなら、その字だけを印の字にする（ほかの字は変えない）。
    箱の字がない・いくつもあるときは、どれのことか決められないので書かない（EditError。箱ごとの前後の字を示す）。"""
    at = [i for i, ch in enumerate(text) if _is_box(ch)]
    shown = text.strip()[:60]
    if not at:
        raise EditError(f"{where} は字「{shown}」のセルで、チェックの箱の字（□・☐ など）がありません。true/false を書くと、"
                        "印刷された字が消えるので書いていません。チェックの印にするなら字を目印にした書き込み（{\"@1\": {\"old\": …, "
                        "\"new\": …, \"cell\": …}}）、まるごと書き換えると決めたなら {\"value\": …, \"overwrite_text\": true} で指定してください")
    if len(at) > 1:
        ctx = "／".join(f"{n}「{text[max(0, i - 4):i]}［{text[i]}］{text[i + 1:i + 9]}」" for n, i in enumerate(at, 1))
        raise EditError(f"{where} の字には、チェックの箱の字が {len(at)} つあります（{ctx}）。true/false では、どれに印をつけるか"
                        "決められないので書いていません。空欄ごとの値のリスト（箱の所に true か ☑ などの字、ほかは null）で指定してください")
    i = at[0]
    new = box_glyph(text[i], on)
    if new == text[i]:
        return text, f"{where}: 箱の字「{text[i]}」は、もともと{'印のついた' if on else '印のない'}字です（変えていません）"
    from .. import chars

    bullet = "（この形の箱は箇条書きの印のこともあります。チェックの箱か、窓の前後の字で確かめてください）" \
        if text[i] in chars.BULLET_LIKE_BOXES else ""
    return (text[:i] + new + text[i + 1:],
            f"{where}: 字「{shown}」の中の箱の字「{text[i]}」を「{new}」にしました（ほかの字は元のまま）{bullet}")


def keeps_printed(old: str, new) -> bool:
    """書いたあとの字 new に、元の字 old の（空白でない）字がすべて、同じ順に残っているか（空欄に書き足しただけか）。
    箱の字は、印のついた字・印のない字のどちらに変わってもよい。"""
    if not isinstance(new, str):
        return False
    it = iter(new)
    for ch in old:
        if ch.isspace() or ch == "　" or ch in "_＿":
            continue
        box = _is_box(ch)
        for got in it:
            if got == ch or (box and _is_box(got)):
                break
        else:
            return False
    return True


def anchor_text(wb: Workbook, key: str, spec, *, force: bool = False, allow_formula_overwrite: bool = False,
                pending: dict | None = None):
    """字を目印にした書き込み {"old": 元の字, "new": 書き換えた後の字, "cell"・"sheet"・"occurrence"（任意）}：
    文字のセルの old の字（ちょうど同じ字の並び）を、1か所だけ new にする。(シート, 行, 列, 書き換えた後のセルの字)。
    old が見つからない・いくつもあって決められないときは書かない（EditError。見つかった所を示す）。
    pending：同じ指示の中で、すでに字を書き換えたセルの字（(シート, 行, 列) → 字）。同じセルへの2つ目の old/new は、
    1つ目を書き換えた後の字の上で探す（2つの書き換えが両方残る）。"""
    pending = pending or {}
    if not isinstance(spec, dict) or not isinstance(spec.get("old"), str) or not isinstance(spec.get("new"), str):
        raise EditError(f"{key}: Excel で字を目印にして書くときは {{\"old\": 元の字, \"new\": 書き換えた後の字}} の形で指定してください"
                        "（同じ字がいくつもあれば \"cell\": \"シート!B3\" か \"occurrence\": 何番目か）")
    extra = set(spec) - {"old", "new", "cell", "sheet", "occurrence"}
    if extra:
        raise EditError(f"{key}: 分からない指定があります（{'・'.join(sorted(extra))}）。使えるのは old・new・cell・sheet・occurrence です")
    old, new = spec["old"], spec["new"]
    if old == "":
        raise EditError(f"{key}: old が空です（目印にする元の字を指定してください）")
    if spec.get("cell"):
        sheet, r, c = resolve_target(wb, str(spec["cell"]), inner_merge=False)
        cells = [(sheet, r, c)]
    else:
        sheets = [wb.sheet(str(spec["sheet"]))] if spec.get("sheet") else list(wb.sheets)
        cells = [(sh.name, r, c) for sh in sheets for (r, c) in sorted(sh.cells)]
    hits = []

    def text_of(sheet, r, c):
        cell = wb.sheet(sheet).cells.get((r, c))
        if cell is None or not isinstance(cell.value, str) or cell.has_formula:
            return None
        t_ = pending.get((sheet, r, c), cell.value)
        return t_ if isinstance(t_, str) else None

    for sheet, r, c in cells:
        tx = text_of(sheet, r, c)
        if tx is None:
            continue
        i = tx.find(old)
        while i >= 0:
            hits.append((sheet, r, c, i))
            i = tx.find(old, i + 1)
    if not hits:
        where = f"{spec['cell']} の字" if spec.get("cell") else "文字のセル（数式でないもの）"
        raise EditError(f"{key}: 「{old}」が{where}に見つかりません（字はセルの字のとおり、空白も含めて指定してください）")
    occ = spec.get("occurrence")
    if occ is not None:
        if not isinstance(occ, int) or isinstance(occ, bool) or not 1 <= occ <= len(hits):
            raise EditError(f"{key}: occurrence は 1〜{len(hits)} の数で指定してください（「{old}」は {len(hits)} か所）")
        hit = hits[occ - 1]
    elif len(hits) > 1:
        ctx = "／".join(f"{s}!{addr(r, c)}「{text_of(s, r, c)[max(0, i - 6):i + len(old) + 6].strip()}」"
                       for s, r, c, i in hits[:10])
        raise EditError(f"{key}: 「{old}」が {len(hits)} か所にあり、どれか決められないので書いていません（{ctx}"
                        f"{'ほか' if len(hits) > 10 else ''}）。\"cell\" でセルを、同じセルの中なら \"occurrence\" で何番目かを指定してください")
    else:
        hit = hits[0]
    sheet, r, c, i = hit
    sh = wb.sheet(sheet)
    if sh.protected and wb.style(sh.style_of(r, c)).locked and not force:
        raise EditError(f"{key}: {sheet}!{addr(r, c)} はシート保護でロックされたセルです（Excelでも書き込めません）")
    text = text_of(sheet, r, c)
    return sheet, r, c, text[:i] + new + text[i + len(old):]


# ---------------------------------------------------------------- 書き込みの入口
@dataclass
class InputPlan:
    """書き込む指示を、セルと部品ごとの値にしたもの（ファイルはまだ変えない）。"""
    targets: dict  # (シート, 行, 列) → Excel に入力したのと同じ規則で読んだ値
    controls: dict  # チェックボックス：(シート, 図形の番号) → (シート, 部品, 状態)
    writes: list  # チェックボックス以外のフォームの部品の状態
    ax_writes: list  # ActiveX の部品の性質（MS-OFORMS）
    warnings: list
    text_written: dict = field(default_factory=dict)  # 指示の鍵 → (シート, 行, 列)：文字のセルの一部を書き換えたもの（箱の字・字を目印・空欄）
    replaced_text: list = field(default_factory=list)  # 印刷された字を、明示（overwrite_text・old/new）なしに置き換えたセル
    anchored: dict = field(default_factory=dict)  # 字を目印にした指示の鍵（@…）→ 書いたセルの id
    rich: dict = field(default_factory=dict)  # (シート, 行, 列) → (<is> の中身, 書いた後の字ごとの書式, 外したふりがな)：字ごとの書式を残す
    direct: dict = field(default_factory=dict)  # (シート, 行, 列) → 指示の鍵：セルの番地・字を目印に書いたセル
    exclusive: list = field(default_factory=list)  # 2つのうち1つを選ぶ箱の両方に印をつけた所（警告）


def plan_inputs(wb: Workbook, values: dict[str, object], *, allow_formula_overwrite: bool = False,
                force: bool = False) -> InputPlan:
    """values を、fill と同じ読み方でセルの値と部品の状態にする。入力規則とマクロはまだ見ない。

    セルに書かれることになる値（文字の中の空欄に入れた結果、部品につながったセルの値など）が分かるので、
    ほかのセルの値で決まる選択肢を、書き込み後の状態で計算するのにも使う。
    """
    targets: dict[tuple[str, int, int], object] = {}
    controls: dict[tuple[str, str], tuple[str, Control, bool]] = {}
    warnings: list[str] = []
    links = checkbox_links(wb)
    from . import controls as _ctl

    others = _ctl.controls_linked_to(wb)
    writes: list[tuple[str, Control, dict]] = []  # チェックボックス以外の部品の状態
    from .. import axfile

    ax_links = axfile.excel_links(wb)
    ax_writes: list = []  # ActiveX の部品の性質（MS-OFORMS）

    def take(plan) -> None:
        for key, val_ in plan.cells.items():
            targets[key] = val_
        writes.extend(plan.controls)

    def take_ax(sheet_: str, ctl_, v_, where: str) -> None:
        from ..activex import AxError

        try:
            p_ = axfile.plan_excel(wb, sheet_, ctl_, v_, force=force, prior=ax_writes)
        except AxError as e:
            raise EditError(f"{where}: {e}") from None
        targets.update(p_.cells)
        ax_writes.extend(p_.controls)
        warnings.extend(p_.notes)

    text_written: dict = {}
    replaced_text: list = []
    anchored: dict = {}
    explicit: set = set()  # 字の置き換えを明示したセル（overwrite_text: true・old/new）
    claimed: dict = {}  # (シート, 行, 列) → (指示の鍵, 字を目印にした書き込みか)
    linked_cells: set = set()  # チェックボックスにつながったセル（字の書式の確かめから外す）
    exclusive: list = []

    def claim(cell_: tuple, key_: str, anchor_: bool = False) -> None:
        """同じセルを2つの指示で書かない（B9 と B9:D9、b1 と B1 など）：後の指示が前の指示を黙って消すので、書かない。
        字を目印にした書き込み（old/new）どうしは、同じセルでも、1つ目を書き換えた後の字の上で続けて書き換える。"""
        if cell_ in claimed:
            prev, prev_anchor = claimed[cell_]
            if not (anchor_ and prev_anchor):
                raise EditError(f"「{prev}」と「{key_}」は、同じセル {cell_[0]}!{addr(cell_[1], cell_[2])} を指しています。"
                                "後の指示が前の指示を消してしまうので、書いていません。1つのセルには1つの指示で書いてください")
        claimed[cell_] = (key_, anchor_)

    for k, v in values.items():
        if k.startswith("@"):
            # 字を目印にした書き込み（Word と同じ {"old": 元の字, "new": 書き換えた後の字}）：文字のセルの、old の字を1か所だけ new にする
            sheet, r, c, new_text = anchor_text(wb, k, v, force=force, allow_formula_overwrite=allow_formula_overwrite,
                                                pending={c_: t_ for c_, t_ in targets.items() if c_ in explicit})
            claim((sheet, r, c), k, True)
            targets[(sheet, r, c)] = new_text
            text_written[k] = (sheet, r, c)
            anchored[k] = f"{sheet}!{addr(r, c)}"
            explicit.add((sheet, r, c))
            continue
        overwrite_text = False
        if isinstance(v, dict) and "value" in v and "circle" not in v:
            # {"value": 値, "overwrite_text": true}：印刷された字のあるセルを、まるごと書き換えると明示した
            extra = set(v) - {"value", "overwrite_text"}
            if extra:
                raise EditError(f"{k}: 分からない指定があります（{'・'.join(sorted(extra))}）。使えるのは value と overwrite_text です")
            overwrite_text = v.get("overwrite_text") is True
            v = v["value"]
        ctl = find_control(wb, k)
        if ctl is not None:
            sheet = wb.sheet(k.rsplit("!", 1)[0].strip("'")).name
            if getattr(ctl, "ax", None) is not None:
                take_ax(sheet, ctl, v, k)  # ActiveX の部品：性質と、つながったセル
                continue
            if ctl.kind in ("Radio",) + _ctl.CHOICE_KINDS + _ctl.VALUE_KINDS:
                try:
                    take(_ctl.plan(wb, sheet, ctl, as_bool(v) if ctl.kind == "Radio" and as_bool(v) is not None else v))
                except ValueError as e:
                    raise EditError(f"{k}: {e}")
                continue
            if ctl.kind != "CheckBox":
                raise EditError(f"{k}: {ctl.kind} は、値を書く部品ではありません（ボタン・ラベル・グループボックスなど）")
            b = as_bool(v)
            if b is None:
                raise EditError(f"{k} はチェックボックスです。TRUE（チェックあり）か FALSE（なし）を指定してください")
            controls[(sheet, ctl.shape_id)] = (sheet, ctl, b)
            if ctl.link_cell and wb.has_sheet(ctl.link_cell[0]):
                ls = wb.sheet(ctl.link_cell[0]).name
                claim((ls, ctl.link_cell[1], ctl.link_cell[2]), k)
                linked_cells.add((ls, ctl.link_cell[1], ctl.link_cell[2]))
                targets[(ls, ctl.link_cell[1], ctl.link_cell[2])] = b
                for sh2, other in links.get((ls, ctl.link_cell[1], ctl.link_cell[2]), []):
                    controls[(sh2, other.shape_id)] = (sh2, other, b)
            continue
        sheet, r, c = resolve_target(wb, k, inner_merge=False)
        claim((sheet, r, c), k)
        sh = wb.sheet(sheet)
        cell = sh.cells.get((r, c))
        if cell is not None and cell.has_formula and not allow_formula_overwrite:
            raise EditError(f"{sheet}!{addr(r, c)} は数式のセルです（計算で決まるので、直接は書き込みません）")
        linked = links.get((sheet, r, c))
        other = others.get((sheet, r, c))
        if linked or other is not None or (sheet, r, c) in ax_links:
            linked_cells.add((sheet, r, c))
        if (sheet, r, c) in ax_links:
            # ActiveX の部品につながったセル：部品の値として受け取り、部品とセルを合わせる
            take_ax(*ax_links[(sheet, r, c)], v, f"{sheet}!{addr(r, c)}")
            continue
        if other is not None:
            # オプションボタン・ドロップダウン・スピンなどにつながったセル：部品の選択肢の文字か番号で受け取り、部品も合わせる
            try:
                if other[0] == "radio":
                    take(_ctl.plan_group(wb, other[1], v))
                else:
                    s_, c_ = other[1]
                    take(_ctl.plan(wb, s_, c_, v))
            except ValueError as e:
                raise EditError(f"{sheet}!{addr(r, c)}: {e}")
            continue
        if sh.protected and wb.style(sh.style_of(r, c)).locked and not linked and not force:
            raise EditError(f"{sheet}!{addr(r, c)} はシート保護でロックされたセルです（Excelでも書き込めません）")
        if isinstance(v, (list, tuple)):
            # 文字の中の空欄（「令和　　年　　月　　日」）に、順に値を入れる
            from .window import blank_spans, fill_blanks

            if cell is None or not isinstance(cell.value, str):
                raise EditError(f"{sheet}!{addr(r, c)} は文字の中に空欄のあるセルではありません")
            spans = blank_spans(wb, sheet, r, c)
            vals = list(v)
            sp_ = spans if spans is not None else [(a_, b_) for a_, b_, _k in _cell_blanks(cell.value)]
            if len(sp_) == len(vals):
                # 箱の字の所に true/false：その箱の字を、チェックの印の字にする（「TRUE」の字は書かない）
                vals = [box_glyph(cell.value[a_], x) if isinstance(x, bool) and b_ - a_ == 1 and _is_box(cell.value[a_])
                        else x for (a_, b_), x in zip(sp_, vals)]
            try:
                v = fill_blanks(cell.value, vals, spans)
            except ValueError as e:
                raise EditError(f"{sheet}!{addr(r, c)}: {e}")
            targets[(sheet, r, c)] = v
            text_written[k] = (sheet, r, c)
            exclusive += _both_boxes_marked(cell.value, v, f"{sheet}!{addr(r, c)}")
            continue
        printed_text = (cell is not None and isinstance(cell.value, str) and cell.value.strip() != "" and not cell.has_formula
                        and not linked)
        if isinstance(v, bool) and printed_text and not overwrite_text:
            # 字のあるセルに true/false：セルの字の中のチェックの箱の字（□ など）に印をつける（外す）。字は消さない
            new_text, note = check_box_in_text(cell.value, v, f"{sheet}!{addr(r, c)}")
            exclusive += _both_boxes_marked(cell.value, new_text, f"{sheet}!{addr(r, c)}")
            targets[(sheet, r, c)] = new_text
            text_written[k] = (sheet, r, c)
            warnings.append(note)
            continue
        val, notes = interpret_input(wb, sheet, r, c, v)
        if printed_text and not overwrite_text and not keeps_printed(cell.value, val):
            from .window import input_cells

            if (r, c) in input_cells(wb, sh):
                pass  # 入力欄（ロックの解除・入力規則のあるセル）のいまの値：印刷された字ではなく、書き換える値
            else:
                replaced_text.append(f"{sheet}!{addr(r, c)}: 印刷された字「{cell.value.strip()[:60]}」を、"
                                     f"「{to_text(val)[:60]}」に置き換えました（元の字が残っていません）。"
                                     "字の中の空欄に書くなら空欄ごとの値のリスト、一部を書き換えるなら {\"@1\": {\"old\": …, \"new\": …, "
                                     "\"cell\": …}}、まるごと書き換えると決めたなら {\"value\": …, \"overwrite_text\": true} で指定してください")
        if overwrite_text:
            explicit.add((sheet, r, c))
        if linked:
            b = as_bool(val)
            if b is None:
                names = "、".join(f"「{x[1].name}」" for x in linked)
                raise EditError(f"{sheet}!{addr(r, c)} はチェックボックス{names}につながったセルです。"
                                f"TRUE（チェックあり）か FALSE（なし）を指定してください")
            val = b
            for sh2, other in linked:
                controls[(sh2, other.shape_id)] = (sh2, other, b)
        else:
            warnings += notes
        targets[(sheet, r, c)] = val

    # Excel のセルに入る字は 32,767 字まで（それより長い字は、Excel で開くと切られる）
    for (s_, r_, c_), v_ in targets.items():
        if isinstance(v_, str) and len(v_) > MAX_CELL_TEXT:
            raise EditError(f"{s_}!{addr(r_, c_)}: 字が {len(v_):,} 字あります。Excel のセルに入るのは {MAX_CELL_TEXT:,} 字までなので、"
                            "書いていません（分けて書く欄があるか、書類で確かめてください）")
    # 字の一部を書き換えたセル（空欄・箱の字・old/new）：字ごとの書式（取り消し線・色）とふりがなを残す
    rich: dict = {}
    for key_, cell_ in text_written.items():
        if cell_ in rich or not isinstance(targets.get(cell_), str):
            continue
        s_, r_, c_ = cell_
        if getattr(wb, "kind", "xlsx") == "xls":
            rc = getattr(wb, "rich_cells", {})
            if cell_ not in rc:
                continue
            isst = rc[cell_]
            if isst is None or getattr(wb, "biff", 8) == 5 or getattr(wb, "dual_format", False):
                raise EditError(f"{key_}: {s_}!{addr(r_, c_)} は、字ごとに書式（取り消し線・字の色など）の違う字のセルです。この形の"
                                "古い Excel（.xls）のセルでは、字の一部を書き換えると、その書式が消える（消した字の取り消し線が消えるなど）ので、"
                                "書いていません")
            from .richtext import splice_props, xls_runs, xls_runs_of

            old_text = wb.shared_strings[isst]
            try:
                new_props, written_props = splice_props(old_text, xls_runs(old_text, wb.sst_runs[isst]), targets[cell_])
                runs_ = xls_runs_of(new_props)
            except ValueError as e:
                raise EditError(f"{key_}: {s_}!{addr(r_, c_)}: {e}。書き足す所を、書式の同じ字の間にする"
                                "（old・new で、書き足す字の両側が同じ書式の字になるようにする）か、本人に確かめてください") from None
            rich[cell_] = ("xls", runs_, [], written_props)
            continue
        if not hasattr(wb, "zip") or not hasattr(wb.sheet(s_), "path"):
            continue
        from .richtext import cell_rst, splice

        try:
            rst = cell_rst(wb, s_, r_, c_)
        except (ValueError, KeyError):
            rst = None
        if rst is None or not rst.rich:
            continue
        try:
            rich[cell_] = splice(rst, targets[cell_])
        except ValueError as e:
            raise EditError(f"{key_}: {s_}!{addr(r_, c_)}: {e}。書き足す所を、書式の同じ字の間にする"
                            "（old・new で、書き足す字の両側が同じ書式の字になるようにする）か、本人に確かめてください") from None
        if rich[cell_][2]:
            warnings.append(f"{s_}!{addr(r_, c_)}: 書き換えた字に付いていたふりがな（{'・'.join(rich[cell_][2])} の読み）を外しました")
    direct = {c_: k_ for c_, (k_, _a) in claimed.items() if c_ not in linked_cells}
    return InputPlan(targets, controls, writes, ax_writes, warnings, text_written, replaced_text, anchored, rich, direct, exclusive)


def _both_boxes_marked(old: str, new: str, where: str) -> list[str]:
    """字の中に箱の字がちょうど2つ（「□有　□無」「□はい　□いいえ」）で、書いた後に両方に印がついたら、警告の言葉。
    2つの箱は、どちらか1つを選ぶ欄のことが多い（両方に印をつけると、どちらを選んだか分からない）。"""
    from .. import chars

    marked = chars.CHECKED_BOXES | chars.MAYBE_CHECKED
    a = [ch for ch in old if _is_box(ch)]
    b = [ch for ch in new if _is_box(ch)]
    if len(a) != 2 or len(b) != 2 or not all(ch in marked for ch in b) or all(ch in marked for ch in a):
        return []
    return [f"{where}: 2つの箱の字の両方に印をつけました（「{new.strip()[:40]}」）。どちらか1つを選ぶ欄なら、本人に確かめて"
            "1つだけにしてください"]


def fill(wb: Workbook, values: dict[str, object], *, allow_formula_overwrite: bool = False,
         strict: bool = False, force: bool = False, macro_answers: list | None = None,
         macro_inputs: dict | None = None, by_macro: set | None = None, run_macros: bool = True) -> EditResult:
    """values: {'シート!セル': 値} または {'シート!Check Box 1': True}。

    値は、Excelでそのセルに手で入力したときと同じ規則で受け取る。入力規則のエラーが「停止」のものは、
    Excelと同じく受け付けない（force=True で警告にとどめる）。strict=True なら、どの入力規則の問題も受け付けない。
    書き込んだ値に依存する数式の計算結果と、チェックボックスの表示も更新したファイルを返す。
    """
    plan_ = plan_inputs(wb, values, allow_formula_overwrite=allow_formula_overwrite, force=force)
    targets, controls, writes, ax_writes, warnings = (plan_.targets, plan_.controls, plan_.writes, plan_.ax_writes,
                                                      plan_.warnings)
    text_extra = {"text_written": {k_: (f"{s_}!{addr(r_, c_)}", plan_.targets.get((s_, r_, c_)))
                                    for k_, (s_, r_, c_) in plan_.text_written.items()},
                  "replaced_text": list(plan_.replaced_text), "anchored": dict(plan_.anchored)}
    from .. import axfile
    from . import controls as _ctl

    def check():
        calc = Calc(wb, targets)
        return [i for (sheet, r, c), val in targets.items() for i in rule_issues(wb, sheet, r, c, val, calc)]

    issues = _run_deep(check)
    blocking = [i for i in issues if i.blocking or strict]
    if blocking and not force:
        raise EditError("入力規則に合いません（Excelでも受け付けない値です）:\n" + "\n".join(i.message for i in blocking))
    warnings += [i.message for i in issues]
    # 書いたセルの書式：見えない字（;;;・字と塗りが同じ色・条件付き書式で隠れる）は書かない。取り消し線などは伝える
    fmt_found, invisible = _run_deep(lambda: written_format(wb, plan_))
    if invisible and not force:
        raise EditError("書いた字が見えなくなるセルです（印刷しても見えません）。書いていません：\n" + "\n".join(invisible)
                        + "\n（書式で隠した、計算や部品のためのセルのことが多いので、書く欄か窓で確かめてください）")
    text_extra["written_formatting"] = fmt_found
    text_extra["exclusive"] = list(plan_.exclusive) + _both_controls_checked(wb, controls)
    try:
        text_extra["fit"] = fit_issues(wb, plan_)
    except Exception as e:  # noqa: BLE001  見積もれなかったことを黙らない
        text_extra["fit"] = [{"id": "", "message": f"書いた字がセルに収まるかを見積もれませんでした（{type(e).__name__}）"}]

    # セルの値を変えたときに動くマクロ（Worksheet_Change など）：人が入力したときと同じ順に動かし、マクロが書いたセルも書く
    macro_written, macro_info = (_run_change_macros(wb, targets, warnings, macro_answers, macro_inputs) if run_macros
                                 else (set(), None))

    changed, problems = recalc(wb, targets)
    if getattr(wb, "kind", "xlsx") == "xls":
        res = _fill_xls(wb, targets, changed, problems, controls, warnings, set(by_macro or ()) | macro_written, writes,
                        ax_writes=ax_writes, rich=plan_.rich)
        res.macro = macro_info
        for k_, v_ in text_extra.items():
            setattr(res, k_, v_)
        return res
    patches: list[Patch] = []
    by_sheet: dict[str, SheetPatcher] = {}

    def patcher(sheet: str) -> SheetPatcher:
        if sheet not in by_sheet:
            by_sheet[sheet] = SheetPatcher(wb.zip.read(wb.sheet(sheet).path))
        return by_sheet[sheet]

    by_macro = set(by_macro or ()) | macro_written
    for (sheet, r, c), val in targets.items():
        cell = wb.sheet(sheet).cells.get((r, c))
        before = cell.value if cell else None
        if (cell is not None and values_equal(before, val) and type(before) is type(val) and not cell.has_formula) or \
                (cell is None and val is None):
            continue
        rich_ = plan_.rich.get((sheet, r, c))
        s, e, n = patcher(sheet).set_input(r, c, val, wb.sheet(sheet).style_of(r, c),
                                           rich=rich_[0] if rich_ is not None and isinstance(val, str) else None)
        kind = "macro" if (sheet, r, c) in by_macro else "input"
        patches.append(Patch(wb.sheet(sheet).path, sheet, addr(r, c), kind, before, val, (s, e), n))
    for (sheet, r, c), val in sorted(changed.items()):
        cell = wb.sheet(sheet).cells[(r, c)]
        if (sheet, r, c) in problems:
            continue  # こぼれる範囲が変わる動的配列：Excelで開いたときに計算し直させる
        # 配列数式の一部のセル（結果の配列の、左上以外）も、計算し直した値をそのセルに書く
        s, e, n = patcher(sheet).set_formula_value(r, c, val)
        patches.append(Patch(wb.sheet(sheet).path, sheet, addr(r, c), "formula", cell.value, val, (s, e), n))

    for sheet, p in by_sheet.items():
        wb.zip.replace(wb.sheet(sheet).path, p.build())
    # 数式のセルに値を書いた（マクロが書いた・数式の上書きを許した）：calcChain.xml から、そのセルの記録を取り除く
    # （数式のないセルが残っていると、Excel は開くときに修復を求める）
    gone = {(wb.sheet(p_.sheet).sheet_id, p_.address) for p_ in patches if p_.kind in ("input", "macro")
            and (cc := wb.sheet(p_.sheet).cells.get(_rc(p_.address))) is not None and cc.formula is not None}
    chain = getattr(wb, "calc_chain", None)
    if gone and chain and wb.zip.has(chain):
        new_chain, n_removed = drop_calc_chain(wb.zip.read(chain), gone)
        if n_removed:
            wb.zip.replace(chain, new_chain)
            patches.append(Patch(chain, "", "calcChain", "workbook", None, f"数式を値に置き換えた {n_removed} セルの記録を取り除いた"))

    parts: dict[str, bytes] = {}
    for (sheet, _sid), (_s, ctl, b) in sorted(controls.items(), key=lambda x: (x[0][0], int(x[0][1]) if x[0][1].isdigit() else 0)):
        state = "Checked" if b else None
        if (ctl.checked if ctl.checked == "Checked" else None) == state and ctl.checked != "Mixed":
            continue
        if ctl.prop_part is None and ctl.vml_part is None:
            raise EditError(f"{sheet}!{ctl.name}: チェックの状態を書く場所がありません")
        for part, fn in ((ctl.prop_part, lambda x: set_ctrlprop_checked(x, state)),
                         (ctl.vml_part, lambda x: set_vml_checked(x, ctl.shape_id, state))):
            if part:
                parts[part] = fn(parts.get(part, wb.zip.read(part)))
        patches.append(Patch(ctl.prop_part or ctl.vml_part or "", sheet, ctl.name, "control", ctl.checked, state))
    for sheet, ctl, change in writes:
        if all(getattr(ctl, k_) == (tuple(v_) if k_ == "multi_sel" and v_ is not None else v_) or
               (k_ == "multi_sel" and tuple(ctl.multi_sel or ()) == tuple(v_ or ())) for k_, v_ in change.items()):
            continue  # 状態が変わらない
        if ctl.prop_part is None and ctl.vml_part is None:
            raise EditError(f"{sheet}!{ctl.name}: 部品の状態を書く場所がありません")
        if ctl.prop_part:
            parts[ctl.prop_part] = _ctl.set_ctrlprop(parts.get(ctl.prop_part, wb.zip.read(ctl.prop_part)), change)
        if ctl.vml_part:
            parts[ctl.vml_part] = _ctl.set_vml(parts.get(ctl.vml_part, wb.zip.read(ctl.vml_part)), ctl.shape_id, change)
        patches.append(Patch(ctl.prop_part or ctl.vml_part or "", sheet, ctl.name, "control",
                             {k_: getattr(ctl, k_) for k_ in change}, change))
    if ax_writes:
        # ActiveX の部品：保存の形（activeX*.bin・性質の袋）の性質と、控えの絵（EMF）を書き換える
        from ..activex import AxError

        try:
            ax_parts, ax_done = axfile.apply_excel(wb, ax_writes)
        except (AxError, ValueError) as e:
            raise EditError(str(e)) from None
        parts.update(ax_parts)
        for d in ax_done:
            patches.append(Patch(d["parts"][0], d["sheet"], d["control"], "control", d["before"], {"ax": d["changes"]}))
    for part, data in parts.items():
        wb.zip.replace(part, data)

    unverified = [f"{s}!{addr(r, c)}: {why}" for (s, r, c), why in sorted(problems.items())]
    full_calc = False
    if unverified:
        # 再計算できない数式が影響を受ける：Excelで開いたときに必ず再計算させる
        full_calc = _set_full_calc(wb)
        if full_calc:
            patches.append(Patch(wb.wb_path, "", "calcPr", "workbook", None, "fullCalcOnLoad=1"))

    blob = wb.zip.build()
    res = EditResult(blob=blob, patches=patches, warnings=warnings, unverified=unverified,
                     full_calc_on_load=full_calc, macro=macro_info, **text_extra,
                     rich_cells={(s_, addr(r_, c_)) for (s_, r_, c_) in plan_.rich})
    res.original_sha256 = hashlib.sha256(wb.zip.original).hexdigest()
    res.sha256 = hashlib.sha256(blob).hexdigest()
    res.changed_parts = wb.zip.modified_names()
    res.unchanged_parts = len(wb.zip.entries) - len(res.changed_parts)
    # 次の書き込みに影響しないよう、元に戻しておく
    for e in wb.zip.entries:
        e._replacement = None
    if getattr(wb.zip, "added", None):
        wb.zip.drop_added()  # ActiveX の部品の控えの絵を分けるために足した部品
    return res


def _fill_xls(wb, targets, changed, problems, controls, warnings, by_macro: set | None = None,
              writes: list | None = None, ax_writes: list | None = None, rich: dict | None = None) -> EditResult:
    """古いExcel（.xls）への書き込み：ブックのストリームのレコードを差し替え、複合ファイルのほかの部分は変えない。

    Excel 5.0/95（BIFF5）のブックにも書く。Excel 97 と 95 の両方の形（Workbook と Book）を持つブックは、両方に同じ値を書き、
    Book の数式の計算結果も Book の数式で計算し直して、Workbook と食い違わないことを確かめる。"""
    from ..xls.write import XlsEditor, XlsWriteError

    biff5 = getattr(wb, "biff", 8) == 5
    dual = getattr(wb, "dual_format", False)
    if (biff5 or dual) and problems:
        # BIFF5 には、開いたときに計算し直させる確かな印（BIFF8 の RECALCID）がない。古い計算結果を残さないよう、書かない
        cells = "、".join(f"{s}!{addr(r, c)}" for (s, r, c) in sorted(problems)[:10])
        raise EditError("書いた値から計算される数式のうち、todoku が計算し直せないものがあります（" + cells + "）。"
                        "Excel 5.0/95 の形のブックには、開いたときに計算し直させる確かな印がないので、古い計算結果を残さないよう、書きません")
    if biff5 and (controls or writes or ax_writes):
        raise EditError("Excel 5.0/95 の形のブックの、フォームの部品には書けません（部品のレコードの形を確かめていないため）")
    ed = XlsEditor(wb)
    patches: list[Patch] = []
    try:
        for (sheet, r, c), val in targets.items():
            cell = wb.sheet(sheet).cells.get((r, c))
            before = cell.value if cell else None
            if (cell is not None and values_equal(before, val) and type(before) is type(val)) or (cell is None and val is None):
                continue
            rich_ = (rich or {}).get((sheet, r, c))
            ed.set_cell(sheet, r, c, val, wb.sheet(sheet).style_of(r, c) or None,
                        runs=rich_[1] if rich_ is not None and isinstance(val, str) else None)
            patches.append(Patch(wb.stream_name, sheet, addr(r, c), "macro" if (sheet, r, c) in (by_macro or set()) else "input",
                                 before, val))
        for (sheet, r, c), val in sorted(changed.items()):
            cell = wb.sheet(sheet).cells[(r, c)]
            if cell.array_master is not None:
                if biff5 or dual:
                    raise EditError(f"{sheet}!{addr(r, c)}：配列数式の一部の計算結果が変わります。Excel 5.0/95 の形のブックでは、"
                                    "開いたときに計算し直させる確かな印がないので、書きません")
                problems[(sheet, r, c)] = "配列数式の一部"
                continue
            ed.set_formula_value(sheet, r, c, val)
            patches.append(Patch(wb.stream_name, sheet, addr(r, c), "formula", cell.value, val))
    except XlsWriteError as e:
        raise EditError(str(e)) from None
    for (sheet, _sid), (_s, ctl, b) in sorted(controls.items()):
        if bool(ctl.is_checked) == bool(b):
            continue
        ed.set_checkbox(sheet, ctl.shape_id, bool(b))
        patches.append(Patch(wb.stream_name, sheet, ctl.name, "control", ctl.checked, "Checked" if b else None))
    # チェックボックス以外の部品（オプションボタン・ドロップダウン・リストボックス・スピンボタン・スクロールバー）
    for sheet, ctl, change in writes or ():
        if all(tuple(getattr(ctl, k_) or ()) == tuple(v_ or ()) if k_ == "multi_sel" else getattr(ctl, k_) == v_
               for k_, v_ in change.items()):
            continue  # 状態が変わらない
        if "multi_sel" in change:
            # 複数選択のリストボックスの iSel は「選んだ最初の項目」（MS-XLS 2.5.147。選んでいなければ 0）
            change = {**change, "sel": min(change["multi_sel"]) if change["multi_sel"] else 0}
        try:
            ed.set_control(sheet, ctl.shape_id, change)
        except XlsWriteError as e:
            raise EditError(f"{sheet}!{ctl.name}: {e}") from None
        patches.append(Patch(wb.stream_name, sheet, ctl.name, "control", {k_: getattr(ctl, k_) for k_ in change}, change))
    ax_streams: dict = {}
    if ax_writes:
        # ActiveX の部品：Ctls ストリーム（か MBD の入れ物）の部品の性質と、OBJ の中の部品の位置・大きさ
        from ..activex import AxError
        from ..axbin import apply_xls

        try:
            ax_streams, ax_done = apply_xls(wb, ed, ax_writes)
        except AxError as e:
            raise EditError(str(e)) from None
        for d in ax_done:
            patches.append(Patch(d["parts"][0], d["sheet"], d["control"], "control", d["before"], {"ax": d["changes"]}))
        if ax_done:
            warnings.append("古い Excel（.xls）の ActiveX の部品の控えの絵は描き直していません。代わりに、書いた部品に"
                            "「開いたときに部品を読み込む（fAutoLoad）」「印刷のときに絵を直す（fPrintCalc）」の印を立てました。"
                            "ActiveX を止めた Excel や、プレビューでは、古い絵のまま見えることがあります")
    unverified = [f"{s}!{addr(r, c)}: {why}" for (s, r, c), why in sorted(problems.items())]
    if unverified:
        ed.force_recalc()
    # ブックのストリームは最後の EOF の後ろを読まないので、小さすぎるときは0を足してよい
    blob = wb.cfb.replace(wb.stream_name, ed.build(), pad_ok=True) if ed.changed else wb.zip.original
    changed_parts = [wb.stream_name] if ed.changed else []
    if dual and ed.changed:
        blob = _fill_book(wb, blob, targets, changed, controls, writes)
        changed_parts.append(next(n for n in wb.cfb.streams() if n.lower() == "book"))
        warnings.append("このブックは Excel 97 以降の形（Workbook）と Excel 95 の形（Book）を両方持っています。両方に同じ値を書き、"
                        "Book の数式の計算結果も Book の数式で計算し直して、Workbook と同じになることを確かめました")
    if ax_streams:
        from ..axbin import write_streams

        blob = write_streams(blob, ax_streams)
    res = EditResult(blob=blob, patches=patches, warnings=warnings, unverified=unverified,
                     full_calc_on_load=bool(unverified), rich_cells={(s_, addr(r_, c_)) for (s_, r_, c_) in (rich or {})})
    res.original_sha256 = hashlib.sha256(wb.zip.original).hexdigest()
    res.sha256 = hashlib.sha256(blob).hexdigest()
    res.changed_parts = changed_parts + sorted(ax_streams)
    res.unchanged_parts = len(wb.cfb.streams()) - len(res.changed_parts)
    return res


def _fill_book(wb, blob: bytes, targets, changed, controls, writes) -> bytes:
    """両方の形を持つブックの Excel 95 の形（Book・BIFF5）に、Workbook と同じ値を書く。

    書く前に、書くセルの値が Workbook と同じであること、Book の数式で計算し直した結果が Workbook の結果と同じであることを
    確かめる（違えば、どちらかを読み違えているか、二つの形が初めから食い違っているので、書かない）。"""
    from ..cfb import CFB
    from ..xls.book import XlsWorkbook
    from ..xls.write import XlsEditor, XlsWriteError
    from .calc import recalc

    if controls or writes:
        raise EditError("Excel 97 と 95 の両方の形を持つブックの、フォームの部品には書けません（Book の部品の形を確かめていないため）")
    try:
        w5 = XlsWorkbook(wb.zip.original, wb.filename, stream="Book")
    except Exception as e:  # noqa: BLE001
        raise EditError(f"Excel 95 の形（Book）を読めないので、両方の形を合わせて書けません（{e}）") from None
    if w5.biff != 5 or [s.name for s in w5.sheets] != [s.name for s in wb.sheets]:
        raise EditError("Excel 95 の形（Book）のシートの並びが Workbook と違うので、両方の形を合わせて書けません")
    for (sheet, r, c) in targets:
        a8 = wb.sheet(sheet).cells.get((r, c))
        a5 = w5.sheet(sheet).cells.get((r, c))
        v8 = a8.value if a8 else None
        v5 = a5.value if a5 else None
        if not (values_equal(v8, v5) or (v8 in (None, "") and v5 in (None, ""))) or \
                bool(a8 is not None and a8.has_formula) != bool(a5 is not None and a5.has_formula):
            raise EditError(f"{sheet}!{addr(r, c)}：Workbook（{v8!r}）と Book（{v5!r}）の値が初めから違うので、両方の形を合わせて書けません")
    changed5, problems5 = recalc(w5, targets)
    if problems5:
        cells = "、".join(f"{s}!{addr(r, c)}" for (s, r, c) in sorted(problems5)[:10])
        raise EditError("Excel 95 の形（Book）の数式のうち、書いた値から計算されるものを todoku が計算し直せません（" + cells + "）。"
                        "古い計算結果を残さないよう、書きません")
    for k in set(changed) | set(changed5):
        sheet, r, c = k
        c8, c5 = wb.sheet(sheet).cells.get((r, c)), w5.sheet(sheet).cells.get((r, c))
        after8 = changed[k] if k in changed else (c8.value if c8 else None)
        after5 = changed5[k] if k in changed5 else (c5.value if c5 else None)
        if not values_equal(after8, after5):
            raise EditError(f"{sheet}!{addr(r, c)}：書いたあとの計算結果が、Workbook（{after8!r}）と Book（{after5!r}）で食い違います。"
                            "両方の形を合わせて書けないので、書きません")
    ed = XlsEditor(w5)
    try:
        for (sheet, r, c), val in targets.items():
            cell = w5.sheet(sheet).cells.get((r, c))
            before = cell.value if cell else None
            if (cell is not None and values_equal(before, val) and type(before) is type(val)) or (cell is None and val is None):
                continue
            ed.set_cell(sheet, r, c, val, w5.sheet(sheet).style_of(r, c) or None)
        for (sheet, r, c), val in sorted(changed5.items()):
            if w5.sheet(sheet).cells[(r, c)].array_master is not None:
                raise EditError(f"{sheet}!{addr(r, c)}：Book の配列数式の一部の計算結果が変わるので、書きません")
            ed.set_formula_value(sheet, r, c, val)
    except XlsWriteError as e:
        raise EditError(f"Excel 95 の形（Book）: {e}") from None
    if not ed.changed:
        return blob
    return CFB(blob).replace(w5.stream_name, ed.build(), pad_ok=True)


def _set_full_calc(wb: Workbook) -> bool:
    xml = wb.zip.read(wb.wb_path)
    m = re.search(rb"<(\w+:)?calcPr\b[^>]*?/?>", xml)
    if m:
        tag = m[0]
        if b"fullCalcOnLoad" in tag:
            new = re.sub(rb'fullCalcOnLoad="[^"]*"', b'fullCalcOnLoad="1"', tag)
        else:
            new = tag[:-2] + b' fullCalcOnLoad="1"/>' if tag.endswith(b"/>") else tag[:-1] + b' fullCalcOnLoad="1">'
        xml = xml[:m.start()] + new + xml[m.end():]
    else:
        m2 = re.search(rb"</(\w+:)?workbook>", xml)
        p = m2[1] or b""
        # calcPr は definedNames の後ろ、oleSize などの前に置く必要がある
        anchor = re.search(rb"<" + re.escape(p) + rb"(oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b", xml)
        pos = anchor.start() if anchor else m2.start()
        xml = xml[:pos] + b"<" + p + b'calcPr fullCalcOnLoad="1"/>' + xml[pos:]
    wb.zip.replace(wb.wb_path, xml)
    return True


def bake(wb: Workbook) -> EditResult:
    """計算結果が保存されていない数式セルに、Excelと同じ規則で計算した値を書き込む。

    Excel以外のソフトで作ったファイルは、数式だけがあって計算結果がないことがある。
    そのままではExcel以外のソフトやAIには値が見えない。数式は変えず、<v> だけを足す。
    """
    from .calc import Calc, _run_deep, formula_cells

    def run():
        calc = Calc(wb)
        todo = []
        for cell in formula_cells(wb):
            if cell.array_master is not None or cell.ast is None:
                continue
            v = calc.cell_value(cell.sheet, cell.row, cell.col)
            st = calc.status.get((cell.sheet, cell.row, cell.col), ("ok", ""))[0]
            if st != "ok":
                continue
            if cell.value is None or not (values_equal(cell.value, v) and type(cell.value) is type(v)):
                todo.append((cell, v))
        return todo

    todo = _run_deep(run)
    by_sheet: dict[str, SheetPatcher] = {}
    patches = []
    for cell, v in todo:
        sh = wb.sheet(cell.sheet)
        p = by_sheet.setdefault(sh.name, SheetPatcher(wb.zip.read(sh.path)))
        s, e, n = p.set_formula_value(cell.row, cell.col, v)
        patches.append(Patch(sh.path, sh.name, cell.address, "formula", cell.value, v, (s, e), n))
    for name, p in by_sheet.items():
        wb.zip.replace(wb.sheet(name).path, p.build())
    blob = wb.zip.build()
    res = EditResult(blob=blob, patches=patches)
    res.original_sha256 = hashlib.sha256(wb.zip.original).hexdigest()
    res.sha256 = hashlib.sha256(blob).hexdigest()
    res.changed_parts = wb.zip.modified_names()
    res.unchanged_parts = len(wb.zip.entries) - len(res.changed_parts)
    for e in wb.zip.entries:
        e._replacement = None
    return res
