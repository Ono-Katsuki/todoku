"""元のファイルを正本にしたまま、セルに値を書き込む。

シートのXMLを読み直して保存し直すことはしない。書き込むセルと、その値に依存して
計算結果が変わる数式セルの <v> だけを、バイト列の上で差し替える。
それ以外のバイトは1バイトも変えない。
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from xml.sax.saxutils import escape

from ..formula.dates import date_to_serial
from ..formula.numfmt import is_date_format
from ..formula.parser import addr, parse, parse_cell
from ..formula.values import ExcelError, Unsupported, compare, is_number, number_to_text, parse_number, to_text
from .calc import Calc, _run_deep, recalc, values_equal
from .workbook import Control, Workbook


class EditError(ValueError):
    pass


@dataclass
class Patch:
    part: str
    sheet: str
    address: str
    kind: str  # input / formula / control / workbook
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

    def summary(self) -> dict:
        return {
            "inputs": sum(1 for p in self.patches if p.kind == "input"),
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
def resolve_target(wb: Workbook, target: str) -> tuple[str, int, int]:
    """'シート!B5' や 'B5'（先頭シート）を、結合セルの左上に寄せて返す。"""
    if "!" in target:
        sheet, a = target.rsplit("!", 1)
        sheet = sheet.strip("'")
    else:
        sheet, a = wb.sheets[0].name, target
    try:
        sh = wb.sheet(sheet)
    except KeyError:
        raise EditError(f"シート「{sheet}」がありません")
    try:
        r, c = parse_cell(a.replace("$", "").strip())
    except ValueError:
        raise EditError(f"「{target}」はセルの番地ではありません")
    m = sh.merged_at(r, c)
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


def interpret_input(wb: Workbook, sheet: str, r: int, c: int, value) -> tuple[object, list[str]]:
    """書き込む値を、Excelでそのセルに手で入力したときと同じ規則で、数値・日付・論理値・文字にする。

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
            return float(date_to_serial(d)), notes
        notes.append(f"{where}: 「{s}」は日付として読めますが、このセルの表示形式は日付ではないので、文字として書き込みました")
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
        x = float(m[2].replace(",", "")) * (10 ** int(m[3]) if m[3] else 1)
        if m[1] == "-":
            x = -x
        if m[4]:
            x /= 100
        elif _is_percent(fmt):
            # Excelの「パーセント形式での自動入力」：1以上の数は % として受け取る
            if abs(x) >= 1:
                x /= 100
        if t != s.strip() and re.search(r"[０-９]", s):
            notes.append(f"{where}: 全角の数字「{s}」を数値 {number_to_text(x)} として書き込みました（Excelと同じ扱い）")
        return x, notes
    return s, notes


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
                rng = _OP_TEXT.get(op, op).format(lo=number_to_text(lo), hi=number_to_text(hi) if hi is not None else "")
                out.append(Issue(where, f"{where}: 値 {to_text(value)} は入力規則（{rng}）の範囲外です{extra}", v.blocks))
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
    s = escape(str(value)).encode()
    sp = b' xml:space="preserve"' if s != s.strip() or b"\n" in s else b""
    return b"inlineStr", b"<" + p + b"is><" + p + b"t" + sp + b">" + s + b"</" + p + b"t></" + p + b"is>"


def _with_t(attrs, t: bytes | None):
    out = [(k, v) for k, v in attrs if k != b"t"]
    if t is not None:
        # t は r, s の後ろに置く（順序はExcelの慣習に合わせる）
        idx = max([i + 1 for i, (k, _) in enumerate(out) if k in (b"r", b"s")] or [len(out)])
        out.insert(idx, (b"t", t))
    return out


def _input_cell_xml(prefix: bytes, existing: bytes | None, a: str, value, style: int = 0) -> bytes:
    if existing is not None:
        head = re.match(rb"<[\w:]*c\b[^>]*?(?=/?>)", existing)[0]
        attrs = _attrs(head)
    else:
        attrs = [(b"r", a.encode())] + ([(b"s", str(style).encode())] if style else [])
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
        t = b"str"
        vxml = b"<" + prefix + b"v>" + escape(value).encode() + b"</" + prefix + b"v>"
    else:
        t, vxml = _value_xml(prefix, value)
    body = existing[m.end():]
    if m[2] == b"/>":
        raise EditError("数式セルの形式が想定外です")
    body = re.sub(rb"<" + re.escape(prefix) + rb"v>.*?</" + re.escape(prefix) + rb"v>|<" + re.escape(prefix) + rb"v\s*/>", b"", body, count=1, flags=re.S)
    fm = re.search(rb"<" + re.escape(prefix) + rb"f\b[^>]*?(?:/>|>.*?</" + re.escape(prefix) + rb"f>)", body, re.S)
    if fm:
        body = body[:fm.end()] + vxml + body[fm.end():]
    else:
        body = vxml + body
    return _open_tag(prefix, b"c", _with_t(attrs, t)) + b">" + body


class SheetPatcher:
    def __init__(self, xml: bytes):
        self.xml = xml
        self.prefix = _prefix(xml)
        self.edits: list[tuple[int, int, bytes]] = []
        self._index: dict[str, tuple[int, int]] | None = None

    def find_cell(self, r: int, c: int) -> tuple[int, int] | None:
        if self._index is None:
            # 1回だけ走査して、セル番地→XML上の位置の表を作る（書き換えは build まで反映しないので位置は変わらない）
            p = re.escape(self.prefix)
            pat = re.compile(rb"<" + p + rb"c\b[^>]*?\br=\"([A-Z]+\d+)\"[^>]*?(?:/>|>.*?</" + p + rb"c>)", re.S)
            self._index = {m[1].decode(): (m.start(), m.end()) for m in pat.finditer(self.xml)}
        return self._index.get(addr(r, c))

    def set_input(self, r: int, c: int, value, style: int = 0) -> tuple[int, int, int]:
        a = addr(r, c)
        span = self.find_cell(r, c)
        if span:
            new = _input_cell_xml(self.prefix, self.xml[span[0]:span[1]], a, value)
            self.edits.append((span[0], span[1], new))
            return span[0], span[1], len(new)
        # セル要素がないセルは、Excelと同じく行か列の書式を引き継いで作る
        new = _input_cell_xml(self.prefix, None, a, value, style)
        pos = self._insert_pos(r, c)
        if isinstance(pos, tuple):
            at, wrap = pos
            new = wrap[0] + new + wrap[1]
            self.edits.append((at[0], at[1], new))
            return at[0], at[1], len(new)
        self.edits.append((pos, pos, new))
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
        out = bytearray()
        last = 0
        for s, e, new in sorted(self.edits, key=lambda x: x[0]):
            if s < last:
                raise EditError("同じ位置に重なる書き換えがあります")
            out += self.xml[last:s] + new
            last = e
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


# ---------------------------------------------------------------- 書き込みの入口
def fill(wb: Workbook, values: dict[str, object], *, allow_formula_overwrite: bool = False,
         strict: bool = False, force: bool = False) -> EditResult:
    """values: {'シート!セル': 値} または {'シート!Check Box 1': True}。

    値は、Excelでそのセルに手で入力したときと同じ規則で受け取る。入力規則のエラーが「停止」のものは、
    Excelと同じく受け付けない（force=True で警告にとどめる）。strict=True なら、どの入力規則の問題も受け付けない。
    書き込んだ値に依存する数式の計算結果と、チェックボックスの表示も更新したファイルを返す。
    """
    targets: dict[tuple[str, int, int], object] = {}
    controls: dict[tuple[str, str], tuple[str, Control, bool]] = {}
    warnings: list[str] = []
    links = checkbox_links(wb)
    for k, v in values.items():
        ctl = find_control(wb, k)
        if ctl is not None:
            sheet = wb.sheet(k.rsplit("!", 1)[0].strip("'")).name
            if ctl.kind != "CheckBox":
                raise EditError(f"{k}: {ctl.kind} のコントロールへの書き込みには、まだ対応していません")
            b = as_bool(v)
            if b is None:
                raise EditError(f"{k} はチェックボックスです。TRUE（チェックあり）か FALSE（なし）を指定してください")
            controls[(sheet, ctl.shape_id)] = (sheet, ctl, b)
            if ctl.link_cell and wb.has_sheet(ctl.link_cell[0]):
                ls = wb.sheet(ctl.link_cell[0]).name
                targets[(ls, ctl.link_cell[1], ctl.link_cell[2])] = b
                for sh2, other in links.get((ls, ctl.link_cell[1], ctl.link_cell[2]), []):
                    controls[(sh2, other.shape_id)] = (sh2, other, b)
            continue
        sheet, r, c = resolve_target(wb, k)
        sh = wb.sheet(sheet)
        cell = sh.cells.get((r, c))
        if cell is not None and cell.has_formula and not allow_formula_overwrite:
            raise EditError(f"{sheet}!{addr(r, c)} は数式のセルです（計算で決まるので、直接は書き込みません）")
        linked = links.get((sheet, r, c))
        if sh.protected and wb.style(sh.style_of(r, c)).locked and not linked and not force:
            raise EditError(f"{sheet}!{addr(r, c)} はシート保護でロックされたセルです（Excelでも書き込めません）")
        if isinstance(v, (list, tuple)):
            # 文字の中の空欄（「令和　　年　　月　　日」）に、順に値を入れる
            from .window import fill_blanks

            if cell is None or not isinstance(cell.value, str):
                raise EditError(f"{sheet}!{addr(r, c)} は文字の中に空欄のあるセルではありません")
            try:
                v = fill_blanks(cell.value, list(v))
            except ValueError as e:
                raise EditError(f"{sheet}!{addr(r, c)}: {e}")
            targets[(sheet, r, c)] = v
            continue
        val, notes = interpret_input(wb, sheet, r, c, v)
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

    def check():
        calc = Calc(wb, targets)
        return [i for (sheet, r, c), val in targets.items() for i in rule_issues(wb, sheet, r, c, val, calc)]

    issues = _run_deep(check)
    blocking = [i for i in issues if i.blocking or strict]
    if blocking and not force:
        raise EditError("入力規則に合いません（Excelでも受け付けない値です）:\n" + "\n".join(i.message for i in blocking))
    warnings += [i.message for i in issues]

    changed, problems = recalc(wb, targets)
    if getattr(wb, "kind", "xlsx") == "xls":
        return _fill_xls(wb, targets, changed, problems, controls, warnings)
    patches: list[Patch] = []
    by_sheet: dict[str, SheetPatcher] = {}

    def patcher(sheet: str) -> SheetPatcher:
        if sheet not in by_sheet:
            by_sheet[sheet] = SheetPatcher(wb.zip.read(wb.sheet(sheet).path))
        return by_sheet[sheet]

    for (sheet, r, c), val in targets.items():
        cell = wb.sheet(sheet).cells.get((r, c))
        before = cell.value if cell else None
        if (cell is not None and values_equal(before, val) and type(before) is type(val)) or (cell is None and val is None):
            continue
        s, e, n = patcher(sheet).set_input(r, c, val, wb.sheet(sheet).style_of(r, c))
        patches.append(Patch(wb.sheet(sheet).path, sheet, addr(r, c), "input", before, val, (s, e), n))
    for (sheet, r, c), val in sorted(changed.items()):
        cell = wb.sheet(sheet).cells[(r, c)]
        if cell.array_master is not None:
            problems[(sheet, r, c)] = "配列数式の一部"
            continue
        s, e, n = patcher(sheet).set_formula_value(r, c, val)
        patches.append(Patch(wb.sheet(sheet).path, sheet, addr(r, c), "formula", cell.value, val, (s, e), n))

    for sheet, p in by_sheet.items():
        wb.zip.replace(wb.sheet(sheet).path, p.build())

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
                     full_calc_on_load=full_calc)
    res.original_sha256 = hashlib.sha256(wb.zip.original).hexdigest()
    res.sha256 = hashlib.sha256(blob).hexdigest()
    res.changed_parts = wb.zip.modified_names()
    res.unchanged_parts = len(wb.zip.entries) - len(res.changed_parts)
    # 次の書き込みに影響しないよう、元に戻しておく
    for e in wb.zip.entries:
        e._replacement = None
    return res


def _fill_xls(wb, targets, changed, problems, controls, warnings) -> EditResult:
    """古いExcel（.xls）への書き込み：ブックのストリームのレコードを差し替え、複合ファイルのほかの部分は変えない。"""
    from ..xls.write import XlsEditor

    ed = XlsEditor(wb)
    patches: list[Patch] = []
    for (sheet, r, c), val in targets.items():
        cell = wb.sheet(sheet).cells.get((r, c))
        before = cell.value if cell else None
        if (cell is not None and values_equal(before, val) and type(before) is type(val)) or (cell is None and val is None):
            continue
        ed.set_cell(sheet, r, c, val, wb.sheet(sheet).style_of(r, c) or None)
        patches.append(Patch(wb.stream_name, sheet, addr(r, c), "input", before, val))
    for (sheet, r, c), val in sorted(changed.items()):
        cell = wb.sheet(sheet).cells[(r, c)]
        if cell.array_master is not None:
            problems[(sheet, r, c)] = "配列数式の一部"
            continue
        ed.set_formula_value(sheet, r, c, val)
        patches.append(Patch(wb.stream_name, sheet, addr(r, c), "formula", cell.value, val))
    for (sheet, _sid), (_s, ctl, b) in sorted(controls.items()):
        if bool(ctl.is_checked) == bool(b):
            continue
        ed.set_checkbox(sheet, ctl.shape_id, bool(b))
        patches.append(Patch(wb.stream_name, sheet, ctl.name, "control", ctl.checked, "Checked" if b else None))
    unverified = [f"{s}!{addr(r, c)}: {why}" for (s, r, c), why in sorted(problems.items())]
    if unverified:
        ed.force_recalc()
    # ブックのストリームは最後の EOF の後ろを読まないので、小さすぎるときは0を足してよい
    blob = wb.cfb.replace(wb.stream_name, ed.build(), pad_ok=True) if ed.changed else wb.zip.original
    res = EditResult(blob=blob, patches=patches, warnings=warnings, unverified=unverified,
                     full_calc_on_load=bool(unverified))
    res.original_sha256 = hashlib.sha256(wb.zip.original).hexdigest()
    res.sha256 = hashlib.sha256(blob).hexdigest()
    res.changed_parts = [wb.stream_name] if ed.changed else []
    res.unchanged_parts = len(wb.cfb.streams()) - len(res.changed_parts)
    return res


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
