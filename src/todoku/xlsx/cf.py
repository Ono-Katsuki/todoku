"""条件付き書式で、いま画面に見えている文字と、色で強調されているセルを決める。

様式には「条件を満たすと警告文を背景と同じ色にして隠す」「入力が必要な欄だけ色を付ける」といった
仕掛けが多い。人に見えているとおりにAIへ伝えるため、条件付き書式の式をセルの位置で計算し、
文字の色と塗りの色を決める。計算できない規則（データバー・上位10項目など）は、見え方を変えないものとして扱う。
"""

from __future__ import annotations

from dataclasses import dataclass

from ..formula.parser import parse
from ..formula.values import ExcelError, compare, is_number
from .calc import Calc
from .workbook import Sheet, Workbook

# 式（formula）で判定できる規則
_FORMULA_TYPES = {"expression", "containsText", "notContainsText", "beginsWith", "endsWith", "containsBlanks",
                  "notContainsBlanks", "containsErrors", "notContainsErrors", "timePeriod"}


@dataclass
class Look:
    hidden: bool = False  # 文字が背景と同じ色などで見えない
    highlighted: bool = False  # 条件付き書式で塗りの色が付いている
    unknown: bool = False  # 判定できない規則がかかっている


def _truthy(v) -> bool | None:
    if isinstance(v, ExcelError):
        return False
    if isinstance(v, bool):
        return v
    if is_number(v):
        return float(v) != 0
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        return None
    return False if v is None else None


class Looks:
    def __init__(self, wb: Workbook, calc: Calc | None = None):
        self.wb = wb
        self.calc = calc or Calc(wb)

    def _eval(self, sh: Sheet, rule, r: int, c: int, f: str):
        from ..formula.evaluator import Evaluator

        r0, c0 = rule.origin
        ev = Evaluator(self.calc, sh.name, r, c, r - r0, c - c0)
        return ev.scalar(ev.run(parse(f.lstrip("="))))

    def matches(self, sh: Sheet, rule, r: int, c: int) -> bool | None:
        try:
            if rule.type in _FORMULA_TYPES and rule.formulas:
                return _truthy(self._eval(sh, rule, r, c, rule.formulas[0]))
            if rule.type == "cellIs" and rule.formulas:
                v = self.calc.cell_value(sh.name, r, c)
                a = self._eval(sh, rule, r, c, rule.formulas[0])
                b = self._eval(sh, rule, r, c, rule.formulas[1]) if len(rule.formulas) > 1 else None
                op = rule.operator or "equal"
                if v is None:
                    v = 0.0 if is_number(a) else ""
                x = compare(v, a)
                if op == "between":
                    return x >= 0 and compare(v, b) <= 0
                if op == "notBetween":
                    return not (x >= 0 and compare(v, b) <= 0)
                return {"equal": x == 0, "notEqual": x != 0, "greaterThan": x > 0, "lessThan": x < 0,
                        "greaterThanOrEqual": x >= 0, "lessThanOrEqual": x <= 0}.get(op)
        except Exception:  # noqa: BLE001
            return None
        return None

    def look(self, sh: Sheet, r: int, c: int) -> Look:
        rules = sorted((x for x in sh.cond_rules if x.covers(r, c)), key=lambda x: x.priority)
        base = self.wb.style(sh.style_of(r, c))
        font, fill, fmt = None, None, None
        out = Look()
        for rule in rules:
            if rule.dxf is None or rule.dxf >= len(self.wb.dxfs):
                continue
            hit = self.matches(sh, rule, r, c)
            if hit is None:
                out.unknown = True
                continue
            if not hit:
                continue
            d = self.wb.dxfs[rule.dxf]
            # 優先順位の高い規則の書式が勝つ（項目ごと）
            font = font or d.font_color
            if fill is None and d.fill_color is not None:
                fill = d.fill_color
                out.highlighted = fill not in ("FFFFFF",)
            fmt = fmt or d.num_fmt
            if rule.stop:
                break
        f = font or base.font_color or "000000"
        b = fill or base.fill_color or "FFFFFF"
        num = fmt or base.num_fmt
        out.hidden = f == b or num.strip() == ";;;"
        return out
