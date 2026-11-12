"""OpenDocument の表計算（.ods）に書き込む：書いたセル（と、その値で計算結果が変わる数式のセル）の要素だけを、
content.xml のバイト列への小さな差分として変える。

- 値は、Excel と同じ規則で受け取る（日付・数・割合・真偽・文字。セルの書式で決める。xlsx.edit.interpret_input）。入力規則も
  Excel と同じに確かめる（table:content-validation を Excel の形にしたもの）
- 同じ中身を繰り返す行・セル（table:number-rows-repeated・number-columns-repeated）は、書くセルの前・そのセル・後に分ける
- セルの値（office:value-type・office:value など）と見える字（text:p）を書く。見える字は、そのセルの数の書式で作る
  （書式が保存された見える字と合うと確かめられたときだけ。合わない書式は、書いた値だけを書き、そのことを知らせる）
- 書いた値に連動する数式は、todoku の数式の部品で計算し直し、保存された計算結果（と見える字）を書き換える。
  計算し直せない数式は、計算結果が古いままになることを、結果に書く
"""

from __future__ import annotations

import datetime as dt
import hashlib
import re
from dataclasses import dataclass, field

from ..formula.dates import serial_to_date
from ..formula.numfmt import format_value
from ..formula.parser import addr
from ..formula.values import ExcelError, is_number, number_to_text
from ..xlsx.calc import Calc, _run_deep, recalc, values_equal
from ..xlsx.edit import (EditError, _both_boxes_marked, check_box_in_text, explicit_overwrite, interpret_input, is_printed_text, replaced_printed,
                         resolve_target, rule_issues)
from .sheet import RowEl, OdsWorkbook
from .xml import NS, esc_text, set_attr, top, well_formed

VALUE_ATTRS = ("office:value-type", "office:value", "office:date-value", "office:time-value", "office:boolean-value",
               "office:string-value", "office:currency", "calcext:value-type")


@dataclass
class Patch:
    sheet: str
    address: str
    kind: str  # input / formula
    before: object
    after: object
    key: str = ""


@dataclass
class OdsFillResult:
    blob: bytes
    patches: list[Patch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    stale: list[str] = field(default_factory=list)  # 計算し直せなかった数式（計算結果が古いまま）
    replaced_text: list[str] = field(default_factory=list)  # 印刷された字を、明示なしに置き換えた所（証明で想定外の変更として示す）
    written_formatting: list = field(default_factory=list)  # 書いたセルの書式（取り消し線・白い字など）：[{id, formatting}]


class SheetWriter:
    def __init__(self, wb: OdsWorkbook):
        self.wb = wb
        self.data = wb.data
        root = top(wb.pkg.tree(wb.part))
        self.pre = {k: root.prefix_of(NS[k]) or k for k in ("office", "table", "text", "calcext")}
        self.want: dict[tuple[str, int, int], tuple[object, bool]] = {}  # (シート, 行, 列) → (値, 数式のセルか)

    def set(self, sheet: str, r: int, c: int, value, formula: bool = False) -> None:
        self.want[(sheet, r, c)] = (value, formula)

    # ------------------------------------------------------------------ セルの要素
    def _attr_name(self, qn: str) -> str:
        p, local = qn.split(":")
        return f"{self.pre.get(p, p)}:{local}"

    def cell_xml(self, el, value, formula: bool, style_idx: int) -> bytes:
        data = self.data
        tag = data[el.s:el.se]
        tag = set_attr(tag, self._attr_name("table:number-columns-repeated"), None)
        had_calcext = re.search(rb"\s" + re.escape(self.pre["calcext"].encode()) + rb":value-type=", tag) is not None
        for a in VALUE_ATTRS:
            tag = set_attr(tag, self._attr_name(a), None)
        attrs, shown = self.value_attrs(value, formula, style_idx)
        if not had_calcext:
            attrs = [(k, v) for k, v in attrs if not k.startswith("calcext:")]
        for k, v in attrs:
            tag = set_attr(tag, self._attr_name(k), v)
        if tag.endswith(b"/>"):
            tag = tag[:-2].rstrip() + b">"
        qname = re.match(rb"<([^\s/>]+)", tag)[1]
        paras = b""
        if shown is not None and shown != "":
            tp = self.pre["text"]
            paras = b"".join(f"<{tp}:p>".encode() + _inline(line, tp) + f"</{tp}:p>".encode() for line in shown.split("\n"))
        # 段落でない中身（コメント・図形など）はそのまま、段落だけを書き換える
        kids = b""
        put = False
        if el.se != el.e:
            for k in el.kids:
                is_para = getattr(k, "tag", "") in ("{" + NS["text"] + "}p", "{" + NS["text"] + "}h", "{" + NS["text"] + "}list")
                if is_para:
                    if not put:
                        kids += paras
                        put = True
                    continue
                kids += data[k.s:k.e] if hasattr(k, "kids") else data[k.s:k.e]
        if not put:
            kids += paras
        return tag + kids + b"</" + qname + b">"

    def value_attrs(self, value, formula: bool, style_idx: int) -> tuple[list[tuple[str, str]], str | None]:
        """(値の属性, 見える字)。"""
        wb = self.wb
        st = wb.style(style_idx)
        kind = wb.style_kind[style_idx] if style_idx < len(wb.style_kind) else "number"
        if value is None or value == "":
            return [], None
        if isinstance(value, ExcelError):
            return [("office:value-type", "float"), ("office:value", "0"), ("calcext:value-type", "error")], value.code
        if isinstance(value, bool):
            return ([("office:value-type", "boolean"), ("office:boolean-value", "true" if value else "false"),
                     ("calcext:value-type", "boolean")], self.shown(value, style_idx, "TRUE" if value else "FALSE"))
        if isinstance(value, str):
            attrs = [("office:value-type", "string")]
            if formula:
                attrs.append(("office:string-value", value))
            attrs.append(("calcext:value-type", "string"))
            return attrs, value
        x = float(value)
        if kind == "date":
            d = _serial_date(wb, x)
            if x < 1 and wb.null_date == dt.date(1899, 12, 30):
                # 1900 年 1 月 1 日より前の日付：書式の道具は Excel の数え方（0 は「1900/1/0」、負の数は表示できない）なので、
                # 見える字は作らず、値の字（ISO）を書く
                self.unformatted = getattr(self, "unformatted", 0) + 1
                return [("office:value-type", "date"), ("office:date-value", d), ("calcext:value-type", "date")], d
            return ([("office:value-type", "date"), ("office:date-value", d), ("calcext:value-type", "date")],
                    self.shown(x, style_idx, d))
        if kind == "time":
            return ([("office:value-type", "time"), ("office:time-value", _duration(x)), ("calcext:value-type", "time")],
                    self.shown(x, style_idx, number_to_text(x)))
        typ = kind if kind in ("percentage", "currency") else "float"
        return ([("office:value-type", typ), ("office:value", _num(x)), ("calcext:value-type", typ)],
                self.shown(x, style_idx, number_to_text(x)))

    def shown(self, value, style_idx: int, fallback: str) -> str:
        """見える字：書式が、保存された見える字と合うと確かめられた（か、照らす字のない）書式なら、その書式で作る。"""
        st = self.wb.style(style_idx)
        if getattr(st, "unconvertible", False) or self.wb.format_ok.get(style_idx) is False:
            self.unformatted = getattr(self, "unformatted", 0) + 1
            return fallback
        try:
            return format_value(value, st.num_fmt)
        except Exception:  # noqa: BLE001
            return fallback

    # ------------------------------------------------------------------ 行の要素を作り直す
    def build(self) -> bytes:
        edits: list[tuple[int, int, bytes]] = []
        for sheet, rows in self.wb.rows.items():
            mine = {(r, c): v for (s, r, c), v in self.want.items() if s == sheet}
            if not mine:
                continue
            for rec in rows:
                hit = {(r, c): v for (r, c), v in mine.items() if rec.r0 <= r < rec.r0 + rec.n}
                if hit:
                    edits.append((rec.el.s, rec.el.e, self.row_xml(sheet, rec, hit)))
                    for k in hit:
                        mine.pop(k)
            if mine:
                r, c = next(iter(mine))
                raise EditError(f"{sheet}!{addr(r, c)}：ファイルの中に、このセルの行がありません（表の外です）")
        out = bytearray()
        at = 0
        for s, e, b in sorted(edits):
            out += self.data[at:s]
            out += b
            at = e
        out += self.data[at:]
        return self.wb.pkg.rebuilt({self.wb.part: bytes(out)})

    def row_xml(self, sheet: str, rec: RowEl, hit: dict) -> bytes:
        data = self.data
        el = rec.el
        tag = data[el.s:el.se]
        if el.se == el.e:
            raise EditError(f"{sheet}!{rec.r0} 行目：中身のない行の要素には書けません")
        rows_with = sorted({r for r, _c in hit})
        segs: list[tuple[int, int, bool]] = []  # (始まりの行, 行の数, 書く行か)
        r = rec.r0
        for rr in rows_with:
            if rr > r:
                segs.append((r, rr - r, False))
            segs.append((rr, 1, True))
            r = rr + 1
        if r < rec.r0 + rec.n:
            segs.append((r, rec.r0 + rec.n - r, False))
        rep = self._attr_name("table:number-rows-repeated")
        out = bytearray()
        inner = data[el.se:el.es]
        close = data[el.es:el.e]
        for r0, n, write in segs:
            t = set_attr(tag, rep, str(n) if n > 1 else None)
            if not write:
                out += t + inner + close
                continue
            out += t + self.cells_xml(sheet, rec, r0, {c: v for (rr, c), v in hit.items() if rr == r0}) + close
        return bytes(out)

    def cells_xml(self, sheet: str, rec: RowEl, r: int, cols: dict) -> bytes:
        data = self.data
        el = rec.el
        out = bytearray()
        at = el.se
        rep = self._attr_name("table:number-columns-repeated")
        sh = self.wb.sheet(sheet)
        done = set()
        for ce in rec.cells:
            hits = sorted(c for c in cols if ce.c0 <= c < ce.c0 + ce.n)
            if not hits:
                continue
            if ce.covered:
                raise EditError(f"{sheet}!{addr(r, hits[0])}：結合で隠れたセルには書けません")
            out += data[at:ce.el.s]
            tag = data[ce.el.s:ce.el.se]
            whole = data[ce.el.s:ce.el.e]
            c = ce.c0
            for h in hits:
                if h > c:
                    out += _repeat(whole, tag, rep, h - c, ce.el)
                cell = sh.cells.get((r, h))
                style_idx = cell.style if cell is not None else sh.style_of(r, h)
                value, formula = cols[h]
                out += self.cell_xml(ce.el, value, formula, style_idx)
                done.add(h)
                c = h + 1
            if c < ce.c0 + ce.n:
                out += _repeat(whole, tag, rep, ce.c0 + ce.n - c, ce.el)
            at = ce.el.e
        out += data[at:el.es]
        missing = [c for c in cols if c not in done]
        if missing:
            raise EditError(f"{sheet}!{addr(r, missing[0])}：ファイルの中に、このセルの要素がありません（行の外です）")
        return bytes(out)


def _repeat(whole: bytes, tag: bytes, rep: str, n: int, el) -> bytes:
    """同じ中身のセルを n 回（繰り返しの数だけを変える）。"""
    new_tag = set_attr(tag, rep, str(n) if n > 1 else None)
    return new_tag + whole[len(tag):]


def _inline(line: str, tp: str) -> bytes:
    out = []
    for tok in re.split(r"( +|\t)", line):
        if not tok:
            continue
        if tok[0] == " ":
            out.append(f"<{tp}:s/>" if len(tok) == 1 else f'<{tp}:s {tp}:c="{len(tok)}"/>')
        elif tok == "\t":
            out.append(f"<{tp}:tab/>")
        else:
            out.append(esc_text(tok))
    return "".join(out).encode("utf-8")


def _num(x: float) -> str:
    if x == int(x) and abs(x) < 1e15:
        return str(int(x))
    return repr(x)


def _serial_date(wb: OdsWorkbook, x: float) -> str:
    if wb.null_date == dt.date(1899, 12, 30):
        d = serial_to_date(x, False)
    else:
        d = wb.null_date + dt.timedelta(days=int(x))
    frac = x - int(x)
    if abs(frac) < 1e-9:
        return d.isoformat()
    secs = round(frac * 86400)
    return f"{d.isoformat()}T{secs // 3600:02d}:{secs % 3600 // 60:02d}:{secs % 60:02d}"


def _duration(x: float) -> str:
    secs = round(abs(x) * 86400)
    return f"{'-' if x < 0 else ''}PT{secs // 3600:02d}H{secs % 3600 // 60:02d}M{secs % 60:02d}S"


# ====================================================================== 入口
def fill(wb: OdsWorkbook, values: dict, *, force: bool = False) -> OdsFillResult:
    targets: dict[tuple[str, int, int], object] = {}
    keys: dict[tuple[str, int, int], str] = {}
    warnings: list[str] = []
    problems: list[str] = []
    claimed: dict[tuple[str, int, int], str] = {}
    replaced_text: list[str] = []
    for k, v in values.items():
        try:
            sheet, r, c = resolve_target(wb, k, inner_merge=False)
            v, overwrite_text = explicit_overwrite(k, v)
        except EditError as e:
            problems.append(str(e))
            continue
        if (sheet, r, c) in keys or (sheet, r, c) in claimed:
            # 同じセルを2つの指示で書かない（B9 と B9:D9、b1 と B1）：後の指示が前の指示を黙って消す
            problems.append(f"「{claimed.get((sheet, r, c), keys.get((sheet, r, c)))}」と「{k}」は、同じセル {sheet}!{addr(r, c)} を"
                            "指しています。後の指示が前の指示を消してしまうので、書いていません。1つのセルには1つの指示で書いてください")
            continue
        claimed[(sheet, r, c)] = str(k)
        sh = wb.sheet(sheet)
        cell = sh.cells.get((r, c))
        if cell is not None and cell.has_formula and not force:
            problems.append(f"{sheet}!{addr(r, c)} は数式のセルです（計算で決まるので、直接は書き込みません）")
            continue
        if sh.protected and wb.style(sh.style_of(r, c)).locked and not force:
            problems.append(f"{sheet}!{addr(r, c)} はシートの保護でロックされたセルです（LibreOffice でも書き込めません）")
            continue
        if isinstance(v, (list, tuple)):
            from ..xlsx.window import blank_spans, fill_blanks

            if cell is None or not isinstance(cell.value, str):
                problems.append(f"{sheet}!{addr(r, c)} は文字の中に空欄のあるセルではありません")
                continue
            try:
                v = fill_blanks(cell.value, list(v), blank_spans(wb, sheet, r, c))
            except ValueError as e:
                problems.append(f"{sheet}!{addr(r, c)}: {e}")
                continue
            targets[(sheet, r, c)] = v
            keys[(sheet, r, c)] = str(k)
            continue
        # 印刷された字のあるセル：Excel の書き手と同じ決まり（字を消す書き込みは、明示がなければ証明で示す）
        printed_text = is_printed_text(cell)
        if isinstance(v, bool) and printed_text and not overwrite_text:
            try:
                new_text, note = check_box_in_text(cell.value, v, f"{sheet}!{addr(r, c)}", anchors=False)
            except EditError as e:
                problems.append(str(e))
                continue
            warnings += [note] + _both_boxes_marked(cell.value, new_text, f"{sheet}!{addr(r, c)}")
            targets[(sheet, r, c)] = new_text
            keys[(sheet, r, c)] = str(k)
            continue
        val, notes = interpret_input(wb, sheet, r, c, v)
        if printed_text and not overwrite_text:
            gone = replaced_printed(wb, sh, sheet, r, c, cell, val, anchors=False)
            if gone:
                replaced_text.append(gone)
        warnings += [n.replace("Excelと同じ扱い", "LibreOffice と同じ扱い").replace("Excelでは", "表計算では") for n in notes]
        targets[(sheet, r, c)] = val
        keys[(sheet, r, c)] = str(k)
    if problems:
        raise EditError("\n".join(problems))

    def check():
        calc = Calc(wb, targets)
        return [i for (sheet, r, c), val in targets.items() for i in rule_issues(wb, sheet, r, c, val, calc)]

    issues = _run_deep(check)
    blocking = [i for i in issues if i.blocking]
    if blocking and not force:
        raise EditError("入力規則に合いません（LibreOffice でも受け付けない値です）:\n" + "\n".join(i.message for i in blocking))
    warnings += [i.message for i in issues]
    # 書いたセルの書式：見えない字（字と塗りが同じ色など）は書かない。取り消し線などは伝える（Excel の書き手と同じ）
    from ..xlsx.cellcheck import written_cells_format

    try:
        fmt_found = written_cells_format(wb, targets, keys, force)
    except ValueError as e:
        raise EditError(str(e)) from None
    changed, unsolved = recalc(wb, targets)
    w = SheetWriter(wb)
    patches: list[Patch] = []
    for (sheet, r, c), val in targets.items():
        cell = wb.sheet(sheet).cells.get((r, c))
        before = cell.value if cell is not None else None
        w.set(sheet, r, c, val)
        patches.append(Patch(sheet, addr(r, c), "input", before, val, keys.get((sheet, r, c), "")))
    for (sheet, r, c), val in sorted(changed.items()):
        if (sheet, r, c) in unsolved:
            continue
        cell = wb.sheet(sheet).cells[(r, c)]
        w.set(sheet, r, c, val, formula=True)
        patches.append(Patch(sheet, addr(r, c), "formula", cell.value, val))
    blob = w.build()
    stale = [f"{s}!{addr(r, c)}（{why}）" for (s, r, c), why in sorted(unsolved.items())]
    if stale:
        warnings.append("書いた値に連動する数式のうち、" + "、".join(stale[:10]) + ("ほか" if len(stale) > 10 else "")
                        + " は todoku では計算し直せないので、保存された計算結果は古いままです（表計算ソフトで開いて計算し直すまで、"
                        "古い結果が見えます。本人に伝えてください）")
    if getattr(w, "unformatted", 0):
        warnings.append(f"{w.unformatted} か所は、セルの数の書式を規則で再現できないので、見える字を書いた値のまま書きました"
                        "（値は正しく、表計算ソフトで開くと書式どおりに表示されます）")
    return OdsFillResult(blob, patches, warnings, stale, replaced_text, fmt_found)


# ====================================================================== 証明
@dataclass
class OdsProof:
    original_sha256: str
    new_sha256: str
    entries_total: int = 0
    entries_identical: int = 0
    cells_total: int = 0
    cells_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        return (f"パッケージの部品 {self.entries_identical}/{self.entries_total} がバイト一致。セル {self.cells_total:,} のうち、"
                f"変わったのは書いたセルとその計算結果の {self.cells_changed} か所だけ"
                + ("" if not self.unexpected else f"（想定外の変更 {len(self.unexpected)} 件）"))


def _state(wb: OdsWorkbook, sheet: str, r: int, c: int):
    sh = wb.sheet(sheet)
    cell = sh.cells.get((r, c))
    key = (sheet, r, c)
    if cell is None and sh.__dict__.get("repeat_ranges"):
        cell = wb._virtual_cell(sh, r, c, keep=True)  # 繰り返しの範囲のマス（1 マスずつ作っていない所）
    if cell is None:
        return None
    v = cell.value.code if isinstance(cell.value, ExcelError) else cell.value
    return (v, wb.value_type.get(key), wb.shown.get(key), wb.odf_formula.get(key), wb.style_names[cell.style])


def prove(original: bytes, new: bytes, written: dict, recalculated: dict | None = None) -> OdsProof:
    """written：(シート, 行, 列) → 書いた値。recalculated：計算し直した数式のセル → 値。"""
    from .textproof import package_check

    pr = OdsProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    changed = package_check(pr, original, new, {"content.xml"})
    a = OdsWorkbook(original)
    try:
        b = OdsWorkbook(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の表を読めません：{e}")
        return pr
    for part in changed:
        err = well_formed(b.pkg.xml(b.part))
        if err:
            pr.lint_errors.append(f"{part}：XML として読めません（{err}）")
    recalculated = recalculated or {}
    if [s.name for s in a.sheets] != [s.name for s in b.sheets]:
        pr.unexpected.append("シートの並びが変わっています")
        return pr
    for sa, sb in zip(a.sheets, b.sheets):
        if sorted(sa.merges) != sorted(sb.merges):
            pr.unexpected.append(f"シート「{sa.name}」：セルの結合が変わっています")
        if (sa.hidden_rows, sa.hidden_cols, sa.protected) != (sb.hidden_rows, sb.hidden_cols, sb.protected):
            pr.unexpected.append(f"シート「{sa.name}」：非表示の行・列か、シートの保護が変わっています")
        if [(v.type, v.sqref) for v in sa.validations] != [(v.type, v.sqref) for v in sb.validations]:
            pr.unexpected.append(f"シート「{sa.name}」：入力規則が変わっています")
        keys = set(sa.cells) | set(sb.cells)
        for (r, c) in sorted(keys):
            pr.cells_total += 1
            k = (sa.name, r, c)
            if k in written or k in recalculated:
                pr.cells_changed += 1
                want = written.get(k, recalculated.get(k))
                cb = sb.cells.get((r, c))
                got = cb.value if cb is not None else None
                if not (values_equal(want, got) or (want in (None, "") and got in (None, ""))):
                    pr.unexpected.append(f"{sa.name}!{addr(r, c)}：書いた後の値が {got!r} で、書いた {want!r} と違います")
                if k in recalculated and a.odf_formula.get(k) != b.odf_formula.get(k):
                    pr.unexpected.append(f"{sa.name}!{addr(r, c)}：数式が変わっています")
                if k in written and b.odf_formula.get(k):
                    pr.unexpected.append(f"{sa.name}!{addr(r, c)}：書いたセルに数式があります")
                ca = sa.cells.get((r, c))
                if ca is not None and cb is not None and a.style_names[ca.style] != b.style_names[cb.style]:
                    pr.unexpected.append(f"{sa.name}!{addr(r, c)}：セルの書式が変わっています")
                continue
            if _state(a, sa.name, r, c) != _state(b, sb.name, r, c):
                pr.unexpected.append(f"{sa.name}!{addr(r, c)}：書いていないセルが変わっています")
                if len(pr.unexpected) > 30:
                    return pr
    # バイトの比べ：書いたセルのある行の要素（分けた行は、もとの行の範囲）を除くと、書く前と後が同じ
    touched = set(written) | set(recalculated)
    xa, xb = a.data, b.data
    cut_a, cut_b = [], []
    for name in a.rows:
        rows_t = {r for (s, r, _c) in touched if s == name}
        for rec in a.rows[name]:
            if any(rec.r0 <= r < rec.r0 + rec.n for r in rows_t):
                cut_a.append((rec.el.s, rec.el.e))
                span = [x for x in b.rows[name] if rec.r0 <= x.r0 < rec.r0 + rec.n]
                if span:
                    cut_b.append((span[0].el.s, span[-1].el.e))
    if _cut(xa, cut_a) != _cut(xb, cut_b):
        pr.unexpected.append("content.xml：書いたセルのある行の外が変わっています")
    return pr


def _cut(xml: bytes, spans) -> bytes:
    out = bytearray()
    at = 0
    for s, e in sorted(spans):
        out += xml[at:s] + b"\x00"
        at = e
    out += xml[at:]
    return bytes(out)


# ====================================================================== api から
def api_fill(blob: bytes, name: str, values: dict, force: bool = False) -> dict:
    from ..openkinds import OpenFormatError

    wb = OdsWorkbook(blob, name)
    try:
        res = fill(wb, values, force=force)
    except EditError as e:
        raise OpenFormatError(str(e)) from None
    written = {(p.sheet, *_rc(p.address)): p.after for p in res.patches if p.kind == "input"}
    recalc_ = {(p.sheet, *_rc(p.address)): p.after for p in res.patches if p.kind == "formula"}
    pr = prove(blob, res.blob, written, recalc_)
    result = {"proof_ok": pr.ok, "proof": pr.headline(), "warnings": list(res.warnings), "unexpected": pr.unexpected[:20],
              "structure_errors": pr.lint_errors[:20],
              "recalculated": {f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "formula"}}
    if res.stale:
        result["stale_formulas"] = res.stale
    if res.replaced_text:
        result["replaced_text"] = list(res.replaced_text)  # 書き込みの終わり（api）で、証明の想定外の変更に入れる
    if res.written_formatting:
        from ..xlsx.cellcheck import formatting_notes

        result["written_formatting"] = list(res.written_formatting)
        result["warnings"] += formatting_notes(res.written_formatting)
    # 書いたことでセルの字が折り返す見込み（目安）：証明・読み戻しは字しか見ないので、見た目の変化を伝える
    try:
        wrapped = wrap_warnings(wb, res.patches)
    except Exception as e:  # noqa: BLE001  見積もれなかったことを黙らない
        wrapped = [f"書いたことで行が折り返すかを見積もれませんでした（{type(e).__name__}）"]
    if wrapped:
        result["wrap"] = wrapped
        result["warnings"] += wrapped
    # 折り返さないセルに書いた字が、セルの幅に収まらず切れて見える・数が「###」になる見込み（Excel の書き手と同じ決め方）
    try:
        fit = fit_warnings(wb, res.patches)
    except Exception as e:  # noqa: BLE001  見積もれなかったことを黙らない
        fit = [{"id": "", "message": f"書いた字がセルに収まるかを見積もれませんでした（{type(e).__name__}）"}]
    if fit:
        result["fit"] = fit
        result["warnings"] += [x["message"] for x in fit]
    readback = {p.key or f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "input"}
    return {"blob": res.blob, "result": result, "readback": readback, "labels": {}}


_ODS_PAD = 0.035 * 28.3465  # LibreOffice のセルの左右の余白の既定（0.035cm）の見積もり


def wrap_warnings(wb: OdsWorkbook, patches) -> list[str]:
    """折り返して表示するセル（fo:wrap-option="wrap"）に書いて、行が増える見込み（目安。wrapcheck）なら、その警告。
    幅は列の style:column-width（結合したセルはその範囲）からセルの左右の余白（fo:padding。なければ 0.035cm）を引いた幅、
    字の大きさはセルのスタイルの fo:font-size（なければ 10 pt）。書く前の字は、書く前の見える字。"""
    from ..wrapcheck import say, wraps

    from .sheet import _pt

    st, part = wb.st, wb.part
    out = []
    widths_of: dict[str, list] = {}
    for p in patches:
        if p.kind != "input" or not isinstance(p.after, str):
            continue
        r, c = _rc(p.address)
        sh = wb.sheet(p.sheet)
        name = _style_name(wb, sh, r, c)
        if st.prop(part, "table-cell", name, "style:table-cell-properties", "fo:wrap-option") != "wrap":
            continue
        ws = _col_widths(wb, p.sheet, widths_of)
        m = sh.merged_at(r, c)
        r1, c1, r2, c2 = m if m else (r, c, r, c)
        if c2 > len(ws) or any(w is None for w in ws[c1 - 1:c2]):
            continue
        pad_pt, size, assumed = _pad_size(wb, name)
        avail = sum(ws[c1 - 1:c2]) - 2 * pad_pt
        before = wb.shown.get((p.sheet, r, c)) or ("" if p.before is None else str(p.before))
        got = wraps(before, p.after, avail, size, size)
        if got is None:
            continue
        merged = f"（結合したセル {addr(r1, c1)}:{addr(r2, c2)}）" if m else ""
        out.append(say(f"{p.sheet}!{p.address}{merged}", p.after, got[0], got[1],
                       f"列の幅 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイント"
                       + (f"（{'・'.join(assumed)}）" if assumed else ""), got[2]))
    return out


def _style_name(wb: OdsWorkbook, sh, r: int, c: int) -> str | None:
    """セルのスタイルの名前（セルの要素がなければ、行・列の既定のセルのスタイル：Sheet.style_of）。"""
    k = sh.style_of(r, c)
    return wb.style_names[k] if 0 <= k < len(wb.style_names) else None


def _col_widths(wb: OdsWorkbook, sheet: str, cache: dict) -> list:
    """シートの列の幅（pt。style:column-width。書かれていない列は None）。"""
    from .sheet import _columns, _pt

    if sheet not in cache:
        t = next((t for t in wb.book.findall("table:table") if t.get("table:name") == sheet), None)
        ws: list = []
        for col in _columns(t) if t is not None else []:
            rep = min(int(col.get("table:number-columns-repeated", "1") or 1), 1024)
            w = wb.st.prop(wb.part, "table-column", col.get("table:style-name"), "style:table-column-properties",
                           "style:column-width")
            ws += [_pt(w) if w else None] * rep
        cache[sheet] = ws
    return cache[sheet]


def _pad_size(wb: OdsWorkbook, name: str | None) -> tuple[float, float, list[str]]:
    """セルの左右の余白（pt）・字の大きさ（pt）と、書かれていなくて仮に置いたものの説明。"""
    from .sheet import _pt

    pad = wb.st.prop(wb.part, "table-cell", name, "style:table-cell-properties", "fo:padding")
    pad_pt = _pt(pad) if pad else None
    assumed = []
    if pad_pt is None:
        pad_pt = _ODS_PAD
        assumed.append("左右の余白は LibreOffice の既定 0.035cm と仮定")
    size_v = wb.st.prop(wb.part, "table-cell", name, "style:text-properties", "fo:font-size")
    size = _pt(size_v) if size_v and not size_v.endswith("%") else None
    if size is None:
        size = 10.0
        assumed.append("字の大きさは LibreOffice の既定 10pt と仮定")
    return pad_pt, size, assumed


def fit_warnings(wb: OdsWorkbook, patches) -> list[dict]:
    """折り返さないセルに書いた字が、セルの幅に収まらず切れて見える・数が「###」になる見込みの所：[{id, cell, message}]
    （xlsx.cellcheck.fit_issues と同じ決め方。折り返すセルは wrap_warnings）。幅は列の style:column-width（結合したセルは
    その範囲）からセルの左右の余白を引いた幅、字の幅は、全角を字の大きさ 1 つ分・ほかを半分（wrapcheck と同じ）。
    回した字（style:rotation-angle）・縦に積む字（style:direction="ttb"）・繰り返して埋めるセル（style:repeat-content）・
    両端をそろえるセル（fo:text-align="justify"）・幅の書かれていない列は、見積もらない。寄せ方（fo:text-align）は、
    固定（style:text-align-source="fix"）のときだけ見る。"""
    from ..textwidth import text_cells
    from ..xlsx.cellcheck import _hashes, blocked_by, clipped_message, overflow_sides

    st, part = wb.st, wb.part
    out: list[dict] = []
    widths_of: dict[str, list] = {}
    targets = {(p.sheet, *_rc(p.address)): p.after for p in patches if p.kind == "input"}
    writer = None
    rtl_of: dict[str, bool] = {}
    for p in patches:
        v = p.after
        if p.kind != "input" or v is None or v == "" or isinstance(v, (bool, ExcelError)):
            continue
        r, c = _rc(p.address)
        sh = wb.sheet(p.sheet)
        if sh.is_hidden(r, c):
            continue
        name = _style_name(wb, sh, r, c)
        cellp = lambda a: st.prop(part, "table-cell", name, "style:table-cell-properties", a)  # noqa: E731
        if cellp("fo:wrap-option") == "wrap" or cellp("style:repeat-content") == "true" or cellp("style:direction") == "ttb":
            continue
        if (cellp("style:rotation-angle") or "0").strip() not in ("0", "0deg", "0.0"):
            continue
        fixed = cellp("style:text-align-source") == "fix"
        align = (st.prop(part, "table-cell", name, "style:paragraph-properties", "fo:text-align") or "") if fixed else ""
        if align == "justify":
            continue
        ws = _col_widths(wb, p.sheet, widths_of)
        m = sh.merged_at(r, c)
        r1, c1, r2, c2 = m if m else (r, c, r, c)
        if c2 > len(ws) or any(w is None for w in ws[c1 - 1:c2]):
            continue
        pad_pt, size, assumed = _pad_size(wb, name)
        avail = sum(ws[c1 - 1:c2]) - 2 * pad_pt
        if isinstance(v, str):
            shown = v
        else:
            writer = writer or SheetWriter(wb)
            shown = writer.value_attrs(v, False, sh.style_of(r, c))[1] or ""
        lines = shown.split("\n")
        need = max(text_cells(ln) for ln in lines) * size / 2
        if need <= avail:
            continue
        where = f"{p.sheet}!{p.address}"
        key = p.key if p.key and not p.key.startswith("@") else where
        basis = (f"列の幅 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイントからの見積もり"
                 + (f"。{'・'.join(assumed)}" if assumed else ""))
        cut = f"{shown[:20]}{'…' if len(shown) > 20 else ''}"
        if not isinstance(v, str):
            # 数は折り返さず、隣へもはみ出さない：幅に入らない数を「###」と表示する。「標準」の書式は、小数の桁を減らして
            # 入れようとするので、整数の部分（と符号）も入らないときだけ
            st_ = wb.style(sh.style_of(r, c))
            if _hashes(v, shown, st_.num_fmt, size / 2, avail):
                out.append({"id": key, "cell": where, "message":
                            f"{where}: 数「{cut}」はセルの幅に入らず、「###」と表示される見込みです（数は折り返しません。{basis}。"
                            "本人に確かめてください）"})
            continue
        if cellp("style:shrink-to-fit") == "true":
            out.append({"id": key, "cell": where, "message":
                        f"{where}: セルの幅に合わせて字が縮んで表示されます（元の大きさの約 {int(100 * avail / need)}%。小さすぎて"
                        f"読みにくくならないか、本人に確かめてください。{basis}）"})
            continue
        if m:
            out.append({"id": key, "cell": where, "message":
                        f"{where}: 書いた字「{cut}」は、結合したセルの幅に収まらず、はみ出す所が切れて見える見込みです（{basis}。"
                        "字を短くするか、本人に確かめてください）"})
            continue
        if p.sheet not in rtl_of:
            t = next((t for t in wb.book.findall("table:table") if t.get("table:name") == p.sheet), None)
            mode = st.prop(part, "table", t.get("table:style-name") if t is not None else None, "style:table-properties",
                           "style:writing-mode")
            rtl_of[p.sheet] = (mode or "").startswith("rl")
        # 寄せ方：start・left は左、end・right は右として見る（LibreOffice の Calc は、読む向きによらず start を左に寄せる）。
        # 寄せ方の決まっていないセルは、セルの読む向き（段落かセルの書式の style:writing-mode）が右から左（rl-tb）なら右
        halign = {"start": "left", "left": "left", "end": "right", "right": "right", "center": "center"}.get(align, "general")
        mode = (st.prop(part, "table-cell", name, "style:paragraph-properties", "style:writing-mode")
                or cellp("style:writing-mode") or "")
        sides, sure = overflow_sides(halign, rtl_of[p.sheet], shown, 2 if mode.startswith("rl") else 1 if mode.startswith("lr") else 0)

        def empty_at(col: int) -> bool:
            if (p.sheet, r, col) in targets:
                return targets[(p.sheet, r, col)] in (None, "")
            if sh.merged_at(r, col) is not None:
                return False
            cell = sh.cells.get((r, col))
            return cell is None or cell.value is None or cell.value == ""

        blocked = blocked_by(sides, sure, need - avail, c, lambda col: (ws[col - 1] or 0.0) if col <= len(ws) else 0.0,
                             empty_at, last=len(ws))
        if blocked:
            out.append({"id": key, "cell": where, "message": clipped_message(where, shown, addr(r, blocked), sure)})
    return out


def _rc(a: str) -> tuple[int, int]:
    from ..formula.parser import parse_cell

    return parse_cell(a)


def read_cells(blob: bytes, name: str, keys: dict) -> dict:
    wb = OdsWorkbook(blob, name)
    out = {}
    for k in keys:
        try:
            s, r, c = resolve_target(wb, k)
        except EditError:
            continue
        cell = wb.sheet(s).cells.get((r, c))
        v = cell.value if cell is not None else None
        out[k] = v.code if isinstance(v, ExcelError) else v
    return out


__all__ = ["fill", "prove", "api_fill", "read_cells", "OdsFillResult", "is_number"]
