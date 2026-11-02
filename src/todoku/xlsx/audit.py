"""エージェントがこの書類を「全て正確に読み書きできるか」を、決定的な処理で確かめる。

読む: ファイルの中の文字と数値を、窓とは別の読み方（XMLを直接）で全部拾い、窓に1つ残らず出ているかを確かめる。
書く: 窓が示したすべての入力欄とチェックボックスに、入力規則に合う試しの値を一度に書き込み、
      変わったのが書いた欄とその計算結果だけであること、Excelが修復なしに開ける構造であること、
      再計算できない数式がないこと、書いた値が窓に正しく読み戻せることを確かめる。
"""

from __future__ import annotations

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

M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"


@dataclass
class Check:
    name: str
    ok: bool
    detail: str
    items: list[str] = field(default_factory=list)  # 問題のあった項目（先頭だけ）


@dataclass
class Audit:
    file: str
    checks: list[Check] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return all(c.ok for c in self.checks)

    def to_dict(self) -> dict:
        return {"file": self.file, "ok": self.ok,
                "checks": [{"name": c.name, "ok": c.ok, "detail": c.detail, "items": c.items[:30]} for c in self.checks]}


# ====================================================================== 読む
def raw_texts(blob: bytes) -> dict[str, list[tuple[str, str]]]:
    """窓を作る処理を使わず、XMLから直接すべての文字と数値を拾う。種類 → [(場所, 文字)]。"""
    z = RawZip(blob)
    out: dict[str, list[tuple[str, str]]] = {"セルの文字": [], "セルの数値": [], "図形の文字": [], "メモ": [],
                                             "チェックボックスなどの文字": [], "入力規則の説明": [], "ヘッダー・フッター": []}
    wb_path = "xl/workbook.xml"
    rels = _rels(z, "")
    for t, target in rels.values():
        if t == "officeDocument":
            wb_path = target
    wb = ET.fromstring(z.read(wb_path))
    wrels = _rels(z, wb_path)
    sst: list[str] = []
    for t, target in wrels.values():
        if t == "sharedStrings" and z.has(target):
            for si in ET.fromstring(z.read(target)).findall(M + "si"):
                sst.append("".join(x.text or "" for x in si.iter(M + "t") if not _in_phonetic(si, x)))
    sheets = wb.find(M + "sheets")
    for s in sheets if sheets is not None else []:
        name = s.get("name")
        rid = s.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        t, path = wrels.get(rid, ("", ""))
        if t != "worksheet" or not z.has(path):
            continue
        root = ET.fromstring(z.read(path))
        for c in root.iter(M + "c"):
            ref = c.get("r")
            typ = c.get("t", "n")
            v = c.find(M + "v")
            where = f"{name}!{ref}"
            if typ == "inlineStr":
                is_ = c.find(M + "is")
                text = "".join(x.text or "" for x in is_.iter(M + "t") if not _in_phonetic(is_, x)) if is_ is not None else ""
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
            if d.tag.endswith("}dataValidation"):
                for key in ("prompt", "promptTitle", "error", "errorTitle"):
                    if d.get(key):
                        out["入力規則の説明"].append((f"{name}（入力規則）", d.get(key)))
        hf = root.find(M + "headerFooter")
        for el in hf if hf is not None else []:
            if el.text and re.sub(r"&.", "", el.text).strip():
                out["ヘッダー・フッター"].append((f"{name}（{el.tag.split('}')[-1]}）", el.text))
        srels = _rels(z, path)
        for t, target in srels.values():
            if not z.has(target):
                continue
            if t == "comments":
                for cm in ET.fromstring(z.read(target)).iter(M + "comment"):
                    text = "".join(x.text or "" for x in cm.iter(M + "t"))
                    if text.strip():
                        out["メモ"].append((f"{name}!{cm.get('ref')}", text))
            elif t == "drawing":
                droot = ET.fromstring(z.read(target))
                for sp in droot.iter():
                    if not sp.tag.endswith("}sp"):
                        continue
                    if any(e.tag.endswith("}compatExt") for e in sp.iter()):
                        continue  # フォームコントロールの絵
                    lines = ["".join(x.text or "" for x in p.iter(A + "t")) for p in sp.iter(A + "p")]
                    text = "\n".join(lines).strip()
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


def _in_phonetic(root: ET.Element, t: ET.Element) -> bool:
    for rph in root.iter(M + "rPh"):
        if any(x is t for x in rph.iter()):
            return True
    return False


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
    raw = raw_texts(wb.zip.original)
    cells = {}
    for sh in doc["sheets"]:
        for c in sh["cells"]:
            cells[f"{sh['name']}!{c['cell']}"] = c
    blobs = {
        "図形の文字": "\n".join(s["text"] + "\n" + s["alt"] for sh in doc["sheets"] for s in sh.get("shapes", [])),
        "メモ": "\n".join(n["text"] for sh in doc["sheets"] for n in sh.get("notes", [])),
        "チェックボックスなどの文字": "\n".join(c["text"] for sh in doc["sheets"] for c in sh.get("controls", [])),
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
    return out


def _squash(s: str) -> str:
    return re.sub(r"\s+", "", s)


# ====================================================================== 書く
def _sample(f: dict, wb: Workbook, sheet: str, calc: Calc | None):
    """入力欄に書く試しの値（入力規則に合うもの）。"""
    from .edit import _rule_number

    if f["kind"] == "checkbox":
        return not bool(f["value"])
    rules = f.get("rules", [])
    r, c = parse_cell(f["cell"])
    sh = wb.sheet(sheet)
    vs = [v for v in sh.validations if v.covers(r, c)]
    for v in vs:
        if v.type == "list":
            opts = list_options(wb, sheet, v.formula1 or "", r, c, v, calc) if calc is not None else f.get("options")
            if not opts:
                return None
            cur = f.get("value")
            for o in opts:
                if not values_equal(o, cur):
                    return o
            return opts[0]
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


def check_writing(wb: Workbook, doc: dict) -> list[Check]:
    fields = [(sh["name"], f) for sh in doc["sheets"] for f in sh.get("fields", [])]
    targets: dict[str, object] = {}
    skipped: list[str] = []
    not_applicable: list[str] = []  # 選択肢がほかの欄で決まり、いまの状態では選べるものがない欄
    later = []
    for sheet, f in fields:
        if f.get("options_depend_on"):
            later.append((sheet, f))
            continue
        v = _sample(f, wb, sheet, None)
        if v is None:
            skipped.append(f"{f['id']}: 試しの値を決められませんでした")
            continue
        targets[f["id"]] = v
    if later:
        # 選択肢がほかの欄で決まるもの（都道府県→市区町村など）は、先に書く値を入れた状態で選ぶ
        from .edit import interpret_input, resolve_target

        over = {}
        for k, v in targets.items():
            try:
                s, r, c = resolve_target(wb, k)
                over[(s, r, c)] = interpret_input(wb, s, r, c, v)[0]
            except EditError:
                pass

        def pick():
            calc = Calc(wb, over)
            return [(f, _sample(f, wb, sheet, calc)) for sheet, f in later]

        for f, v in _run_deep(pick):
            if v is None:
                not_applicable.append(f["id"])
            else:
                targets[f["id"]] = v
    out = []
    try:
        res = fill(wb, targets)
    except EditError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    allowed = {(p.sheet, p.address) for p in res.patches if p.kind == "input"}
    allowed_c = {(p.sheet, p.address) for p in res.patches if p.kind == "control"}
    pr = prove(wb.zip.original, res.blob, allowed, allowed_c)
    n_ctl = sum(1 for _, f in fields if f["kind"] == "checkbox")
    na = (f"（ほかの欄の値しだいで選択肢がなくなる欄 {len(not_applicable):,} か所は、いまの状態ではExcelでも選べないので除外）"
          if not_applicable else "")
    out.append(Check("全入力欄への書き込み", pr.ok and not res.unverified and not skipped,
                     f"入力欄 {len(fields) - n_ctl:,} か所とチェックボックス {n_ctl} 個のうち {len(targets):,} に試しの値を書き込み{na}。"
                     f"{pr.headline()}",
                     [*pr.unexpected, *pr.lint_errors, *res.unverified, *skipped]))
    # 読み戻し：書いた値が、書き込み後のファイルの窓にそのとおり出るか
    after = window.build(Workbook(res.blob, wb.filename))
    got = {}
    for sh in after["sheets"]:
        for c in sh["cells"]:
            got[f"{sh['name']}!{c['cell']}"] = c.get("value")
        for f in sh.get("fields", []):
            if f["kind"] == "checkbox":
                got[f["id"]] = f["value"]
    wrong = []
    from .edit import interpret_input, resolve_target

    for k, v in targets.items():
        try:
            s, r, c = resolve_target(wb, k)
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
    wb = Workbook(blob, filename)
    doc = window.build(wb)
    au = Audit(filename)
    au.checks += check_reading(wb, doc)
    md = window.to_markdown(doc)
    au.checks.append(Check("窓の大きさ", len(md) <= 150_000,
                           f"Markdown {len(md):,} 文字（AIが一度に読める目安 15万文字以内）"))
    au.checks += check_writing(Workbook(blob, filename), doc)
    if example is not None:
        from .example import check

        r = check(blob, example)
        au.checks.append(Check("記入例との照合", r.ok, r.headline(),
                               [f"入力欄の見落とし: {x}" for x in r.not_input_fields]
                               + [f"{m[0]}!{m[1]}: Excel={m[3]!r} こちら={m[4]!r}" for m in r.mismatches]
                               + [f"チェック {m[0]}!{m[1]}: Excel={m[2]} こちら={m[3]}" for m in r.control_mismatches]))
    return au
