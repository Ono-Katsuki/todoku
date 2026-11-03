"""エージェントがこの書類を「全て正確に読み書きできるか」を、決定的な処理で確かめる。

読む: ファイルの中の文字と数値を、窓とは別の読み方（XMLを直接）で全部拾い、窓に1つ残らず出ているかを確かめる。
書く: 窓が示したすべての入力欄とチェックボックスに、入力規則に合う試しの値を一度に書き込み、
      変わったのが書いた欄とその計算結果だけであること、Excelが修復なしに開ける構造であること、
      再計算できない数式がないこと、書いた値が窓に正しく読み戻せることを確かめる。
"""

from __future__ import annotations

from ..book import open_book

import datetime as dt
import json
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from ..formula.dates import date_to_serial
from ..formula.parser import addr, parse_cell
from ..formula.values import ExcelError, is_number, number_to_text, to_text
from ..zipraw import RawZip
from . import window
from .calc import Calc, _run_deep, values_equal
from .edit import EditError, fill, list_options, rule_issues
from .proof import prove
from .workbook import Workbook

REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


@dataclass
class Check:
    name: str
    ok: bool
    detail: str
    items: list[str] = field(default_factory=list)  # 問題のあった項目（先頭だけ）
    # 規則では決められないので、窓にそのことを書いて、エージェントの判断に任せた（ok は、そう知らせたうえでの結果）
    agent: bool = False


@dataclass
class Audit:
    file: str
    checks: list[Check] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return all(c.ok for c in self.checks)

    @property
    def needs_agent(self) -> bool:
        return any(c.agent for c in self.checks)

    def to_dict(self) -> dict:
        return {"file": self.file, "ok": self.ok, "needs_agent": self.needs_agent,
                "checks": [{"name": c.name, "ok": c.ok, "agent": c.agent, "detail": c.detail, "items": c.items[:30]}
                           for c in self.checks]}


# ====================================================================== 読む
def _ln(tag) -> str:
    return tag.split("}")[-1] if isinstance(tag, str) else ""


def _attr(el: ET.Element, local: str) -> str | None:
    """名前空間を問わずに属性を読む（Strict Open XML など、名前空間の違う書き方でも拾うため）。"""
    for k, v in el.attrib.items():
        if _ln(k) == local:
            return v
    return None


def _texts(el: ET.Element) -> str:
    """要素の中の文字（t 要素）をつなぐ。ふりがな（rPh）の中は除く。"""
    out = []

    def walk(e, in_ph):
        name = _ln(e.tag)
        if name == "rPh":
            in_ph = True
        if name == "t" and not in_ph:
            out.append(e.text or "")
        for ch in e:
            walk(ch, in_ph)

    walk(el, False)
    return "".join(out)


def raw_texts(blob: bytes) -> dict[str, list[tuple[str, str]]]:
    """窓を作る処理を使わず、XMLから直接すべての文字と数値を拾う。種類 → [(場所, 文字)]。

    要素は名前空間を見ずに名前だけで探す。窓を作る処理が読み落とす書き方（名前空間の違いなど）が
    あれば、ここでは拾えるので、読み落としとして表に出る。
    """
    z = RawZip(blob)
    out: dict[str, list[tuple[str, str]]] = {"セルの文字": [], "セルの数値": [], "図形の文字": [], "メモ": [],
                                             "チェックボックスなどの文字": [], "入力規則の説明": [], "ヘッダー・フッター": [],
                                             "ActiveX の部品の文字": []}
    wb_path = "xl/workbook.xml"
    for t, target in _rels(z, "").values():
        if t == "officeDocument":
            wb_path = target
    wb = ET.fromstring(z.read(wb_path))
    wrels = _rels(z, wb_path)
    sst: list[str] = []
    for t, target in wrels.values():
        if t == "sharedStrings" and z.has(target):
            sst = [_texts(si) for si in ET.fromstring(z.read(target)) if _ln(si.tag) == "si"]
    sheets = [e for e in wb.iter() if _ln(e.tag) == "sheet"]
    for s in sheets:
        name = s.get("name")
        t, path = wrels.get(_attr(s, "id") or "", ("", ""))
        if t != "worksheet" or not z.has(path):
            continue
        root = ET.fromstring(z.read(path))
        for c in root.iter():
            if _ln(c.tag) != "c":
                continue
            ref = c.get("r")
            typ = c.get("t", "n")
            v = next((x for x in c if _ln(x.tag) == "v"), None)
            where = f"{name}!{ref}"
            if typ == "inlineStr":
                is_ = next((x for x in c if _ln(x.tag) == "is"), None)
                text = _texts(is_) if is_ is not None else ""
                if text:
                    out["セルの文字"].append((where, text))
            elif v is None or v.text is None:
                continue
            elif typ == "s":
                i = int(v.text)
                if i < len(sst) and sst[i]:
                    out["セルの文字"].append((where, sst[i]))
            elif typ in ("str", "e"):
                if v.text:
                    out["セルの文字"].append((where, v.text))
            elif typ == "b":
                out["セルの文字"].append((where, "TRUE" if v.text in ("1", "true") else "FALSE"))
            else:
                out["セルの数値"].append((where, v.text))
        for d in root.iter():
            if _ln(d.tag) == "dataValidation":
                for key in ("prompt", "promptTitle", "error", "errorTitle"):
                    if d.get(key):
                        out["入力規則の説明"].append((f"{name}（入力規則）", d.get(key)))
            elif _ln(d.tag) in ("oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"):
                if d.text and re.sub(r"&.", "", d.text).strip():
                    out["ヘッダー・フッター"].append((f"{name}（{_ln(d.tag)}）", d.text))
        for t, target in _rels(z, path).values():
            if not z.has(target):
                continue
            if t == "control":
                # ActiveX の部品（MS-OFORMS）：部品の中身を直接読み、見出しと文字を拾う
                from ..axfile import load_ooxml, texts

                for s in texts(load_ooxml(z, target, "xlsx")):
                    out["ActiveX の部品の文字"].append((f"{name}（ActiveX）", s))
            elif t == "comments":
                for cm in ET.fromstring(z.read(target)).iter():
                    if _ln(cm.tag) == "comment":
                        text = _texts(cm)
                        if text.strip():
                            out["メモ"].append((f"{name}!{cm.get('ref')}", text))
            elif t == "drawing":
                droot = ET.fromstring(z.read(target))
                for sp in droot.iter():
                    if _ln(sp.tag) != "sp":
                        continue
                    if any(_ln(e.tag) == "compatExt" for e in sp.iter()):
                        continue  # フォームコントロールの絵
                    paras = [p for p in sp.iter() if _ln(p.tag) == "p"]
                    text = "\n".join("".join(x.text or "" for x in p.iter() if _ln(x.tag) == "t") for p in paras).strip()
                    if text:
                        out["図形の文字"].append((f"{name}（図形）", text))
            elif t == "vmlDrawing":
                vml = z.read(target).decode("utf-8", errors="replace")
                for m in re.finditer(r"<v:shape\b[^>]*>(.*?)</v:shape>", vml, re.S):
                    if 'ObjectType="Note"' in m[1]:
                        continue  # メモは comments で数える
                    tb = re.search(r"<v:textbox\b[^>]*>(.*?)</v:textbox>", m[1], re.S)
                    if tb:
                        from .workbook import _vml_text

                        text = _vml_text(tb[1])
                        if text:
                            out["チェックボックスなどの文字"].append((f"{name}（コントロール）", text))
    return out


def _rels(z: RawZip, part: str) -> dict[str, tuple[str, str]]:
    import posixpath

    d, base = posixpath.split(part)
    rp = posixpath.join(d, "_rels", base + ".rels")
    if not z.has(rp):
        return {}
    out = {}
    for r in ET.fromstring(z.read(rp)).iter(REL + "Relationship"):
        t = r.get("Target", "")
        full = t if r.get("TargetMode") == "External" else t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
        out[r.get("Id")] = (r.get("Type", "").rsplit("/", 1)[-1], full)
    return out


def check_reading(wb: Workbook, doc: dict) -> list[Check]:
    if getattr(wb, "kind", "xlsx") == "xls":
        from ..xls.proof import raw_texts as raw_xls

        raw = raw_xls(wb.zip.original)
    else:
        raw = raw_texts(wb.zip.original)
    n_raw = len(raw["セルの文字"]) + len(raw["セルの数値"])
    n_win = sum(1 for sh in doc["sheets"] for c in sh["cells"] if c.get("value") is not None)
    head: list[Check] = []
    if n_raw == 0 and n_win == 0:
        why = "" if getattr(wb, "kind", "xlsx") == "xls" else _no_cells(wb.zip.original)
        if getattr(wb, "kind", "xlsx") == "xls" or why:
            return [Check("読み落とし（全体）", False, "セルの値を1つも読めませんでした（対応していない書き方の可能性があります）"
                          + (f"。{why}" if why else ""))]
        # どのシートの XML にも、値のあるセルの要素がない（空のブック）：読み落としではない。ほかの読み落としの検査は続ける
        head.append(Check("読み落とし（全体）", True, "どのシートにも値のあるセルがありません（すべてのシートの XML を直接読み、"
                          "値・字のあるセルの要素がないことを確かめました）"))
    cells = {}
    for sh in doc["sheets"]:
        for c in sh["cells"]:
            cells[f"{sh['name']}!{c['cell']}"] = c
    blobs = {
        "図形の文字": "\n".join(s["text"] + "\n" + s["alt"] for sh in doc["sheets"] for s in sh.get("shapes", [])),
        "メモ": "\n".join(n["text"] for sh in doc["sheets"] for n in sh.get("notes", [])),
        "チェックボックスなどの文字": "\n".join(c["text"] for sh in doc["sheets"] for c in sh.get("controls", [])),
        "ActiveX の部品の文字": "\n".join(
            "\n".join([c.get("text") or "", str(c.get("value") or "")] +
                      [f"{x.get('caption') or ''}\n{x.get('value') or ''}" for x in c.get("frame", [])])
            for sh in doc["sheets"] for c in sh.get("controls", []) if c.get("activex")),
        "入力規則の説明": json.dumps([sh.get("validations") for sh in doc["sheets"]], ensure_ascii=False),
        "ヘッダー・フッター": "\n".join(h for sh in doc["sheets"] for h in sh.get("header_footer", [])),
    }
    out = []
    for kind, items in raw.items():
        missing = []
        for where, text in items:
            if kind in ("セルの文字", "セルの数値"):
                c = cells.get(where)
                got = None if c is None else c.get("value")
                if kind == "セルの数値":
                    ok = c is not None and is_number(got) and values_equal(float(text), float(got))
                else:
                    ok = c is not None and (to_text(got) == text if not isinstance(got, bool) else (("TRUE" if got else "FALSE") == text))
            elif kind == "ヘッダー・フッター":
                from .workbook import _header_text

                ok = _header_text(text) in blobs[kind]
            elif kind == "入力規則の説明":
                ok = json.dumps(text, ensure_ascii=False)[1:-1] in blobs[kind]
            else:
                ok = _squash(text) in _squash(blobs[kind])
            if not ok:
                missing.append(f"{where}: {text[:40]!r}")
        out.append(Check(f"読み落とし（{kind}）", not missing,
                         f"{len(items) - len(missing):,}/{len(items):,} 項目が窓に出ている", missing))
    return head + out


def _no_cells(blob: bytes) -> str:
    """空のブックかを、窓を作る処理を使わずに確かめる。空でない・確かめられないときは、その理由（空なら ""）。
    ブックのすべてのシートが部品につながり、ワークシートの XML のどこにも、値（v）か字（is）を持つセル（c）がないこと。"""
    z = RawZip(blob)
    wb_path = "xl/workbook.xml"
    for t, target in _rels(z, "").values():
        if t == "officeDocument":
            wb_path = target
    if not z.has(wb_path):
        return "ブックの部品が見つかりません"
    wrels = _rels(z, wb_path)
    sheets = [e for e in ET.fromstring(z.read(wb_path)).iter() if _ln(e.tag) == "sheet"]
    if not sheets:
        return "ブックにシートがありません"
    for s in sheets:
        t, path = wrels.get(_attr(s, "id") or "", ("", ""))
        if not z.has(path):
            return f"シート「{s.get('name')}」の部品が見つかりません"
        if t != "worksheet":
            continue  # グラフのシートなど：セルはない
        for c in ET.fromstring(z.read(path)).iter():
            if _ln(c.tag) == "c" and any(_ln(x.tag) in ("v", "is", "f") for x in c):
                return f"シート「{s.get('name')}」に、値か式のあるセルがあります"
    return ""


def _squash(s: str) -> str:
    return re.sub(r"\s+", "", s)


# ====================================================================== 書く
def _sample(f: dict, wb: Workbook, sheet: str, calc: Calc | None):
    """入力欄に書く試しの値（入力規則に合うもの）。"""
    from .edit import _rule_number

    if f.get("activex"):
        from ..axfile import sample_value

        return sample_value(f)  # ActiveX の部品：部品の決まり（文字数・選択肢・範囲・組）に合う値
    if f["kind"] == "checkbox":
        return not bool(f["value"])
    if f["kind"] in ("radio", "select"):
        if f.get("undecided") or not f.get("options"):
            return None
        return next((o for o in f["options"] if o != f.get("value")), f["options"][0])
    if f["kind"] == "multiselect":
        return [o for o in f.get("options", [])[:2] if o not in (f.get("value") or [])] or f.get("options", [])[:1] or None
    if f["kind"] == "spin":
        lo, hi = f.get("min", 0), f.get("max", 100)
        cur = f.get("value")
        return lo if cur != lo else min(hi, lo + 1)
    if f["kind"] == "inline":
        return ["1"] * f["blanks"]
    rules = f.get("rules", [])
    r, c = parse_cell(f["cell"])
    sh = wb.sheet(sheet)
    vs = [v for v in sh.validations if v.covers(r, c)]
    lists = [v for v in vs if v.type == "list"]
    if lists:
        sets = []
        for v in lists:
            opts = list_options(wb, sheet, v.formula1 or "", r, c, v, calc) if calc is not None else f.get("options")
            if not opts:
                return None
            sets.append(opts)
        both = [o for o in sets[0] if all(any(values_equal(o, x) for x in other) for other in sets[1:])]
        opts = both or sets[0]
        cur = f.get("value")
        for o in opts:
            if not values_equal(o, cur):
                return o
        return opts[0]
    for v in vs:
        if v.type in ("whole", "decimal", "date", "time", "textLength"):
            lo = _rule_number(calc or Calc(wb), sheet, v, r, c, v.formula1)
            hi = _rule_number(calc or Calc(wb), sheet, v, r, c, v.formula2)
            op = v.operator or "between"
            x = {"between": lo, "notBetween": (hi or 0) + 1, "equal": lo, "notEqual": (lo or 0) + 1,
                 "greaterThan": (lo or 0) + 1, "lessThan": (lo or 2) - 1, "greaterThanOrEqual": lo,
                 "lessThanOrEqual": lo}.get(op, lo)
            if x is None:
                x = 1
            if v.type == "textLength":
                n = max(1, int(x))
                return "1234567890" * (n // 10) + "1234567890"[: n % 10]
            if v.type == "whole":
                x = int(x)
            return x
    if rules and any(r_["type"] == "custom" for r_ in rules):
        # 独自の入力規則（数式）：いくつかの候補を、その規則で実際に確かめて選ぶ
        from .edit import interpret_input, rule_issues

        for cand in (1, "テスト", "2026/4/1", "1", 10, "A", 0):
            try:
                val = interpret_input(wb, sheet, r, c, cand)[0]
            except EditError:
                continue
            base = dict(calc.overrides) if calc is not None else {}
            base[(sheet, r, c)] = val
            if not rule_issues(wb, sheet, r, c, val, Calc(wb, base)):
                return cand
        return None
    k = f["kind"]
    if k == "date":
        return "2026/4/1"
    if k == "number":
        return 1
    cell = sh.cells.get((r, c))
    if cell is not None and wb.style(cell.style).num_fmt == "@":
        return "テスト"
    return "テスト"


def _planned_cells(wb: Workbook, targets: dict) -> dict:
    """targets を書いたあとのセルの値（fill と同じ読み方。文字の中の空欄・部品につながったセルも含む）。"""
    from .edit import plan_inputs

    try:
        return dict(plan_inputs(wb, targets).targets)
    except EditError:
        # どれかが書けない（その知らせは fill が出す）：書けるものだけを1つずつ読む
        over = {}
        for k, v in targets.items():
            try:
                over.update(plan_inputs(wb, {k: v}).targets)
            except EditError:
                pass
        return over


def _pick_choices(wb: Workbook, later: list, targets: dict, not_applicable: list, rounds: int = 8) -> None:
    """選択肢（入力規則のリスト）の欄に、書き込み後の状態で選べる値を選んで targets に足す。

    選択肢がほかの欄の値で決まる（INDIRECT で県を選ぶと市の一覧が変わる、など）ことがあるので、
    Excel で人が入れるのと同じく、ほかの欄を入れた状態で選択肢を計算する。選んだ値でほかの欄の選択肢が
    変わることもあるので、すべての選んだ値が、すべてを書いた状態の選択肢に入るまで選び直す（決まった回数まで）。
    """
    from .edit import resolve_target

    where: dict[str, tuple[str, int, int]] = {}
    for sheet, f in later:
        try:
            where[f["id"]] = resolve_target(wb, f["id"])
        except EditError:
            pass
    chosen: dict[str, object] = {}
    todo = list(later)
    for _ in range(rounds):
        state = _planned_cells(wb, {**targets, **chosen})

        def pick(todo=todo, state=state):
            calc = Calc(wb, state)
            return [(sheet, f, _sample(f, wb, sheet, calc)) for sheet, f in todo]

        for sheet, f, v in _run_deep(pick):
            if v is None:
                chosen.pop(f["id"], None)
            else:
                chosen[f["id"]] = v
        # すべてを書いた状態で、選んだ値がまだ選択肢に入っているかを確かめる
        state = _planned_cells(wb, {**targets, **chosen})

        def stale(state=state):
            calc = Calc(wb, state)
            out = []
            for sheet, f in later:
                k = where.get(f["id"])
                if f["id"] not in chosen or k is None:
                    continue
                if any(i.blocking for i in rule_issues(wb, k[0], k[1], k[2], state.get(k), calc)):
                    out.append((sheet, f))
            return out

        todo = _run_deep(stale)
        if not todo:
            break
    for sheet, f in later:
        if f["id"] in chosen:
            targets[f["id"]] = chosen[f["id"]]
        else:
            not_applicable.append(f["id"])


def check_writing(wb: Workbook, doc: dict) -> list[Check]:
    fields = [(sh["name"], f) for sh in doc["sheets"] for f in sh.get("fields", [])]
    targets: dict[str, object] = {}
    skipped: list[str] = []
    not_applicable: list[str] = []  # いまの状態では選べるものがない欄（選択肢がほかの欄で決まる・選択肢が空）
    later = []
    for sheet, f in fields:
        if f["kind"] == "choice":
            later.append((sheet, f))  # 選択肢はほかの欄の内容で決まることがあるので、ほかを入れてから選ぶ
            continue
        if f.get("activex") and (f.get("readonly") or f.get("options_unknown")):
            not_applicable.append(f["id"])  # 書けない ActiveX の部品（使えない・固定・選択肢が分からない・複数選択）
            continue
        v = _sample(f, wb, sheet, None)
        if v is None:
            skipped.append(f"{f['id']}: 試しの値を決められませんでした")
            continue
        targets[f["id"]] = v
    if later:
        _pick_choices(wb, later, targets, not_applicable)
    out = []
    try:
        res = fill(wb, targets)
    except EditError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    allowed = res.allowed_inputs
    allowed_c = {(p.sheet, p.address): p.after for p in res.patches if p.kind == "control"}  # 部品 → 書いた状態
    pr = prove(wb.zip.original, res.blob, allowed, allowed_c, {(p.sheet, p.address) for p in res.patches if p.kind == "macro"})
    n_ctl = sum(1 for _, f in fields if f["kind"] in ("checkbox", "radio", "select", "multiselect", "spin"))
    na = (f"（選択肢が空の欄 {len(not_applicable):,} か所は、いまの状態ではExcelでも選べないので除外。"
          "ほかの欄の値で選択肢が決まるものや、選択肢の参照先が壊れているもの）" if not_applicable else "")
    out.append(Check("全入力欄への書き込み", pr.ok and not res.unverified and not skipped,
                     f"入力欄 {len(fields) - n_ctl:,} か所とフォームの部品 {n_ctl} 個のうち {len(targets):,} に試しの値を書き込み{na}。"
                     f"{pr.headline()}",
                     [*pr.unexpected, *pr.lint_errors, *res.unverified, *skipped]))
    # 読み戻し：書き込み後のファイルを読み直して、書いた値とチェックの状態がそのとおりかを確かめる
    after = open_book(res.blob, wb.filename)
    got = {}
    for sh in after.sheets:
        for (r, c), cell in sh.cells.items():
            v = cell.value
            got[f"{sh.name}!{addr(r, c)}"] = v.code if isinstance(v, ExcelError) else v
        for ctl in sh.controls:
            if ctl.kind == "CheckBox":
                key = f"{sh.name}!{addr(*ctl.link_cell[1:])}" if ctl.link_cell else f"{sh.name}!{ctl.name}"
                if not ctl.link_cell:
                    got[key] = ctl.is_checked
    wrong = []
    from .edit import interpret_input, resolve_target

    # フォームの部品の欄（オプションボタン・ドロップダウン・リスト・スピン）は、書き込み後の窓の欄の値で確かめる
    ctl_kinds = ("radio", "select", "multiselect", "spin")
    ctl_fields = {f["id"]: f for _, f in fields if f["kind"] in ctl_kinds or f.get("activex")}
    after_doc = None
    if any(k in ctl_fields for k in targets):
        after_doc = window.build(after)
        after_vals = {f["id"]: f.get("value") for sh in after_doc["sheets"] for f in sh.get("fields", [])}
    for k, v in targets.items():
        if k in ctl_fields:
            g = after_vals.get(k)
            want = v
            if ctl_fields[k]["kind"] == "spin":
                ok = isinstance(g, (int, float)) and float(g) == float(v)
            elif ctl_fields[k]["kind"] == "multiselect":
                ok = list(g or []) == list(v)
            else:
                ok = g == v
            if not ok:
                wrong.append(f"{k}: 書いた値 {want!r} → 窓では {g!r}")
            continue
        try:
            s, r, c = resolve_target(wb, k)
            if isinstance(v, list):
                from .window import blank_spans, fill_blanks

                want = fill_blanks(wb.sheet(s).cells[(r, c)].value, v, blank_spans(wb, s, r, c))
            else:
                want = interpret_input(wb, s, r, c, v)[0]
            key = f"{s}!{addr(r, c)}"
        except EditError:
            want, key = v, k
        g = got.get(key)
        if isinstance(want, bool) or isinstance(g, bool):
            ok = bool(g) == bool(want)
        elif isinstance(want, ExcelError):
            ok = g == want.code
        else:
            ok = values_equal(want, g)
        if not ok:
            wrong.append(f"{k}: 書いた値 {want!r} → 窓では {g!r}")
    out.append(Check("読み戻し", not wrong, f"書いた {len(targets):,} 欄のうち {len(targets) - len(wrong):,} 欄が、書き込み後の窓にそのまま出ている", wrong))
    return out


def run(blob: bytes, filename: str = "", example: bytes | None = None) -> Audit:
    wb = open_book(blob, filename)
    doc = window.build(wb)
    au = Audit(filename)
    au.checks += check_reading(wb, doc)
    from ..embeds.audit import excel_checks

    au.checks += excel_checks(wb, doc)  # グラフ・SmartArt・埋め込みの読み落とし
    parts = window.markdown_parts(doc)
    biggest = max(len(t) for _, t in parts)
    total = sum(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {total:,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                        if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(open_book(blob, filename), doc)
    ap = agent_places_check([f for sh in doc["sheets"] for f in sh.get("fields", [])])
    if ap is not None:
        au.checks.append(ap)  # 規則では決められない所（［B5？］・文字の中の［？］）：書き込みの試しにも含めた
    overlap = [f["id"] for sh in doc["sheets"] for f in sh.get("fields", []) if f.get("rules_overlap")]
    au.checks.append(Check("入力規則の重なり", True,
                           "入力規則の重なったセルはありません" if not overlap else
                           f"入力規則が重なったセルが {len(overlap)} か所あります。Excelがどの規則を使うかはファイルからは決まらず、"
                           "規則では決められません。窓には重なった規則をそれぞれの範囲とともに出し、どれかの規則に合う値は警告つきで書けます"
                           "（どの値にするかはエージェントが見出しから判断。様式の作り手に直してもらう点です）",
                           overlap, agent=bool(overlap)))
    # 書き込めない入力用の部品は、黙って飛ばさず、ここで知らせる
    passive = ("CheckBox", "Button", "Label", "GBox", "Dialog", "Pict", "Shape", "Movie", "Rect", "",
               "ActiveX CommandButton", "ActiveX Label", "ActiveX Image")
    writable = ("Radio", "Drop", "List", "Spin", "Scroll", "ActiveX TextBox", "ActiveX CheckBox", "ActiveX OptionButton",
                "ActiveX ToggleButton", "ActiveX ComboBox", "ActiveX ListBox", "ActiveX SpinButton", "ActiveX ScrollBar")
    other = sorted({f"{sh.name}!{c.name}（{c.kind}）" for sh in wb.sheets for c in sh.controls
                    if c.kind not in passive + writable})
    # ActiveX の部品：中身を読めないもの・フレームの中の部品（読むだけ）も知らせる
    other += sorted({f"{sh.name}!{c.name}（{c.kind}：{c.ax.error}）" for sh in wb.sheets for c in sh.controls
                     if getattr(c, "ax", None) is not None and c.ax.error})
    other += sorted({f"{sh.name}!{c.name}（フレームの中の部品は読むだけです）" for sh in wb.sheets for c in sh.controls
                     if getattr(c, "ax", None) is not None and c.ax.frame is not None
                     and any(s.kind in ("TextBox", "CheckBox", "OptionButton", "ToggleButton", "ComboBox", "ListBox",
                                        "SpinButton", "ScrollBar") for s in c.ax.frame.sites)})
    undecided = [f"{f['id']}: {f['undecided']}" for sh in doc["sheets"] for f in sh.get("fields", []) if f.get("undecided")]
    other += undecided
    au.checks.append(Check("入力用コントロールへの対応", not other,
                           "書き込めない入力用の部品はありません（チェックボックス・オプションボタン・ドロップダウン・リスト・"
                           "スピンボタン・スクロールバーに書き込めます）" if not other
                           else f"書き込みに対応していないコントロールが {len(other)} 個あります", other))
    # マクロ：動作はまだ再現していない（ボタンを押したときの処理など）。あることを必ず知らせる
    buttons = sorted({f"{sh.name}!{c.name}" for sh in wb.sheets for c in sh.controls
                      if c.kind in ("Button", "ActiveX CommandButton")})
    mi = doc.get("macros")
    if not wb.has_macros:
        au.checks.append(Check("マクロ", True, "マクロはありません"))
    elif not isinstance(mi, dict) or not mi.get("readable"):
        au.checks.append(Check("マクロ", False, "マクロを含みますが、マクロの入れ物を読めませんでした（何をするマクロか分かりません）", buttons))
    else:
        n_proc = sum(len(m["procedures"]) for m in mi["modules"])
        on_change = [f"{e['module']}.{e['procedure']}" for e in mi["events"] if e["on_change"]]
        from ..vbarun.parse import VbaSyntaxError, parse_module

        bad = []
        for m in wb.vba.modules:
            try:
                parse_module(m.name, m.source, getattr(wb.vba, "cc_constants", None))
            except VbaSyntaxError as e:
                bad.append(f"{m.name}: {e}")
        au.checks.append(Check("マクロ", not bad,
                               f"マクロのソース（モジュール {len(mi['modules'])} 個・手続き {n_proc} 個）をすべて窓に出し、"
                               + ("構文をすべて読めました。ボタンの処理は run_macro で動かせます（はい・いいえや読むファイルは、"
                                  "勝手に決めずに本人に確かめます）" if not bad else f"{len(bad)} 個のモジュールの構文を読めませんでした")
                               + (f"（マクロを動かすボタン {len(buttons)} 個）" if buttons else "")
                               + (f"。セルの値を変えたときに動くマクロ {len(on_change)} 個は、書き込みのときに同じように動かします"
                                  if on_change else "")
                               + "。ボタンの処理が要るかは、エージェントがソースを読んで判断します",
                               bad + on_change + buttons, agent=True))
    if example is not None:
        from .example import check

        r = check(blob, example)
        au.checks.append(Check("記入例との照合", r.ok, r.headline(),
                               [f"入力欄の見落とし: {x}" for x in r.not_input_fields]
                               + [f"{m[0]}!{m[1]}: Excel={m[3]!r} こちら={m[4]!r}" for m in r.mismatches]
                               + [f"チェック {m[0]}!{m[1]}: Excel={m[2]} こちら={m[3]}" for m in r.control_mismatches]))
    return au


def agent_places_check(fields: list[dict]) -> "Check | None":
    """規則では決められない所（窓の［…？］）：あれば、エージェントの判断に任せた所として数える。"""
    agent = [f for f in fields if f.get("agent")]
    if not agent:
        return None
    kinds: dict[str, int] = {}
    for f in agent:
        kinds[f.get("why", "")] = kinds.get(f.get("why", ""), 0) + 1
    return Check("規則では決められない所", True,
                 f"書き込む所かもしれない所 {len(agent)} か所を、窓に「？」をつけて示しました（"
                 + "・".join(f"{k} {n}" for k, n in kinds.items())
                 + "）。エージェントが判断し、書く所なら id を指定して書きます。全入力欄への書き込みの試しにも含めました",
                 [f["id"] for f in agent[:20]], agent=True)
