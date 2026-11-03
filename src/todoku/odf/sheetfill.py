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
from ..xlsx.edit import EditError, interpret_input, resolve_target, rule_issues
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
    for k, v in values.items():
        try:
            sheet, r, c = resolve_target(wb, k, inner_merge=False)
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
        val, notes = interpret_input(wb, sheet, r, c, v)
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
    return OdsFillResult(blob, patches, warnings, stale)


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
    cell = wb.sheet(sheet).cells.get((r, c))
    key = (sheet, r, c)
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
    # 書いたことでセルの字が折り返す見込み（目安）：証明・読み戻しは字しか見ないので、見た目の変化を伝える
    try:
        wrapped = wrap_warnings(wb, res.patches)
    except Exception as e:  # noqa: BLE001  見積もれなかったことを黙らない
        wrapped = [f"書いたことで行が折り返すかを見積もれませんでした（{type(e).__name__}）"]
    if wrapped:
        result["wrap"] = wrapped
        result["warnings"] += wrapped
    readback = {p.key or f"{p.sheet}!{p.address}": p.after for p in res.patches if p.kind == "input"}
    return {"blob": res.blob, "result": result, "readback": readback, "labels": {}}


_ODS_PAD = 0.035 * 28.3465  # LibreOffice のセルの左右の余白の既定（0.035cm）の見積もり


def wrap_warnings(wb: OdsWorkbook, patches) -> list[str]:
    """折り返して表示するセル（fo:wrap-option="wrap"）に書いて、行が増える見込み（目安。wrapcheck）なら、その警告。
    幅は列の style:column-width（結合したセルはその範囲）からセルの左右の余白（fo:padding。なければ 0.035cm）を引いた幅、
    字の大きさはセルのスタイルの fo:font-size（なければ 10 pt）。書く前の字は、書く前の見える字。"""
    from ..wrapcheck import say, wraps
    from .sheet import _columns, _pt

    st, part = wb.st, wb.part
    out = []
    tables = {t.get("table:name"): t for t in wb.book.findall("table:table")}
    widths_of: dict[str, list] = {}
    for p in patches:
        if p.kind != "input" or not isinstance(p.after, str):
            continue
        r, c = _rc(p.address)
        sh = wb.sheet(p.sheet)
        cell = sh.cells.get((r, c))
        name = wb.style_names[cell.style] if cell is not None and 0 <= cell.style < len(wb.style_names) else None
        if st.prop(part, "table-cell", name, "style:table-cell-properties", "fo:wrap-option") != "wrap":
            continue
        if p.sheet not in widths_of:
            ws = []
            for col in _columns(tables.get(p.sheet)) if tables.get(p.sheet) is not None else []:
                rep = min(int(col.get("table:number-columns-repeated", "1") or 1), 1024)
                w = st.prop(part, "table-column", col.get("table:style-name"), "style:table-column-properties",
                            "style:column-width")
                ws += [_pt(w) if w else None] * rep
            widths_of[p.sheet] = ws
        ws = widths_of[p.sheet]
        m = sh.merged_at(r, c)
        r1, c1, r2, c2 = m if m else (r, c, r, c)
        if c2 > len(ws) or any(w is None for w in ws[c1 - 1:c2]):
            continue
        pad = st.prop(part, "table-cell", name, "style:table-cell-properties", "fo:padding")
        pad_pt = _pt(pad) if pad else _ODS_PAD
        avail = sum(ws[c1 - 1:c2]) - 2 * pad_pt
        size_v = st.prop(part, "table-cell", name, "style:text-properties", "fo:font-size")
        size = _pt(size_v) if size_v and not size_v.endswith("%") else 10.0
        before = wb.shown.get((p.sheet, r, c)) or ("" if p.before is None else str(p.before))
        got = wraps(before, p.after, avail, size, size)
        if got is None:
            continue
        merged = f"（結合したセル {addr(r1, c1)}:{addr(r2, c2)}）" if m else ""
        out.append(say(f"{p.sheet}!{p.address}{merged}", p.after, got[0], got[1],
                       f"列の幅 約 {avail:.0f} ポイント・字の大きさ {size:g} ポイント", got[2]))
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
