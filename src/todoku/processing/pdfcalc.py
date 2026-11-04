"""PDF のフォームの処理：欄の計算（AFSimple_Calculate など）・計算の順（/CO）・書式と確かめのスクリプト。

計算は Acrobat の組み込みの関数（AFSimple_Calculate）だけを、fill と同じ決まりで再現する（acrojs）。独自の JavaScript は実行せず、
そのまま示す（「評価していない」と書く）。
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path

ACTIONS = {"F": "format", "K": "keystroke", "V": "validate", "C": "calculate"}
_SFN = re.compile(r"/\*\*\s*BVCALC(.*?)EVCALC\s*\*\*/", re.S)
_GETFIELD = re.compile(r"""getField\(\s*(?:"((?:[^"\\]|\\.)*)"|'((?:[^'\\]|\\.)*)')\s*\)""")
OPS = {"SUM", "AVG", "PRD", "MIN", "MAX"}
_NOT_RUN = "JavaScript は実行しません。ソースをそのまま示します（評価していない）"


def _doc_scripts(doc) -> list[dict]:
    """ブックを開いたときに動く文書のスクリプト（/Names /JavaScript）と /OpenAction。"""
    from ..pdf.objects import PdfString, Stream, decode_text

    def js_text(act):
        act = doc.resolve(act)
        if not isinstance(act, dict):
            return None
        js = doc.resolve(act.get("JS"))
        if isinstance(js, Stream):
            js = doc.decode(js)
        if isinstance(js, (bytes, PdfString)):
            js = decode_text(js)
        return js if isinstance(js, str) and js.strip() else None

    out = []
    names = doc.resolve(doc.root.get("Names"))
    tree = doc.resolve(names.get("JavaScript")) if isinstance(names, dict) else None

    def walk(node, depth=0):
        node = doc.resolve(node)
        if not isinstance(node, dict) or depth > 16:
            return
        arr = doc.resolve(node.get("Names"))
        if isinstance(arr, list):
            for i in range(0, len(arr) - 1, 2):
                nm = doc.resolve(arr[i])
                t = js_text(arr[i + 1])
                if t:
                    label = decode_text(nm) if isinstance(nm, (bytes, PdfString)) else str(nm)
                    out.append({"kind": "document_javascript", "name": label, "source": t, "evaluated": False, "note": _NOT_RUN})
        for k in doc.resolve(node.get("Kids")) or []:
            walk(k, depth + 1)

    if tree is not None:
        walk(tree)
    t = js_text(doc.root.get("OpenAction"))
    if t:
        out.append({"kind": "document_javascript", "name": "OpenAction", "source": t, "evaluated": False, "note": _NOT_RUN})
    return out


def model(doc) -> dict:
    from ..formula.parser import addr  # noqa: F401  （読み込みの確かめ）
    from ..pdf import acrojs
    from ..pdf.form import read_fields
    from ..pdf.objects import Ref

    fields = read_fields(doc)
    af = doc.resolve(doc.root.get("AcroForm"))
    co = [r.num for r in (doc.resolve(af.get("CO")) or []) if isinstance(r, Ref)] if isinstance(af, dict) else []
    names = {f.name for f in fields}
    rules: list[dict] = []
    other: list[dict] = []
    calc_names: dict[str, dict] = {}
    for f in fields:
        if f.ref is None:
            continue
        js = acrojs.actions(doc, doc.resolve(f.ref))
        for k, src in js.items():
            kind = ACTIONS[k]
            if k == "C":
                rec = _calc_rule(f, src, co, acrojs, names)
                rules.append(rec)
                calc_names[f.name] = rec
            else:
                p = acrojs.parse_call(src)
                rec = {"kind": kind, "field": f.name, "source": src}
                if p:
                    rec["function"] = p[0]
                    rec["arguments"] = p[1]
                rec["evaluated"] = acrojs.known(src)
                if not rec["evaluated"]:
                    rec["note"] = _NOT_RUN
                other.append(rec)
    # 順番：/CO の順。/CO にない計算は後ろ（fill と同じ）
    rules.sort(key=lambda d: (d["order"] if d["order"] is not None else len(co), d["field"]))
    for i, d in enumerate(rules):
        d["id"] = f"C{i + 1}"
    by_field = {d["field"]: d for d in rules}
    read_by: dict[str, list[str]] = {}
    for d in rules:
        for n in d.get("inputs", []):
            read_by.setdefault(n, []).append(d["id"])
    pos = {d["field"]: i for i, d in enumerate(rules)}
    for d in rules:
        late = [n for n in d.get("inputs", []) if n in pos and pos[n] > pos[d["field"]]]
        if late:
            d["reads_before_computed"] = late  # /CO の順では、まだ計算されていない欄を読む（Acrobat はその順で計算する）
    inputs = [{"field": f.name, "kind": f.kind, "value": _value(f), "read_by": read_by[f.name], **({"label": f.tu} if f.tu else {})}
              for f in fields if f.name in read_by and f.name not in by_field]
    outputs = [d["field"] for d in rules if d["field"] not in read_by]
    return {"kind": "pdf", "rules": rules, "order": [d["id"] for d in rules], "calculation_order_declared": bool(co),
            "inputs": inputs, "outputs": outputs, "field_scripts": other, "document_scripts": _doc_scripts(doc),
            "summary": {"fields": len(fields), "rules": len(rules), "evaluated_rules": sum(1 for d in rules if d["evaluated"]),
                        "not_evaluated_rules": sum(1 for d in rules if not d["evaluated"])},
            "xfa": bool(isinstance(af, dict) and af.get("XFA") is not None)}


def _value(f):
    v = f.value_text()
    return v if v == "" or isinstance(v, (str, int, float, list, type(None))) else str(v)


def _calc_rule(f, src: str, co: list, acrojs, names: set) -> dict:
    d: dict = {"kind": "field_calculation", "field": f.name, "source": src, "order": co.index(f.ref.num) if f.ref.num in co else None}
    p = acrojs.parse_call(src)
    if p and p[0] == "AFSimple_Calculate" and len(p[1]) >= 2 and str(p[1][0]).upper() in OPS:
        op = str(p[1][0]).upper()
        ins = acrojs.calc_sources(src)
        d.update(operation=op, inputs=ins, evaluated=True, text=f"{f.name} = {op}({', '.join(ins)})")
        miss = [n for n in ins if n not in names]
        if miss:
            d["missing_inputs"] = miss
        return d
    m = _SFN.search(src)
    if m:
        d.update(operation="simplified_field_notation", expression=m[1].strip(), evaluated=False, text=f"{f.name} = {m[1].strip()}",
                 note="簡易フィールド記法の式です。" + _NOT_RUN)
    else:
        d.update(operation="custom_javascript", evaluated=False, text=f"{f.name} = （独自の JavaScript）", note=_NOT_RUN)
    mentions = [a or b for a, b in _GETFIELD.findall(src)]
    if mentions:
        d["fields_named_in_source"] = list(dict.fromkeys(mentions))  # 字に書かれた欄の名前（実際に読むかは、実行しないと決まらない）
    return d


_TEMPLATE = '''#!/usr/bin/env python3
"""「@@SOURCE@@」のフォームの計算（とどくの extract_processing が書き出した、標準ライブラリだけで動くモジュール）。

calculate({"欄の名前": 値}) で、/CO の順に AFSimple_Calculate を計算して、計算の欄の値を返す。
独自の JavaScript の欄は再現しない（NotReproducible。on_unknown="mark" なら None）。
"""

from __future__ import annotations

import json
import sys
@@STDLIB@@

SOURCE = @@FILENAME@@
VERIFICATION = None  # @@todoku-verification@@
RULES = @@RULES@@
DEFAULTS = @@DEFAULTS@@


@@VENDORED@@


class NotReproducible(Exception):
    """再現できない欄がある。fields：欄 → 理由。"""

    def __init__(self, fields):
        self.fields = fields
        super().__init__("再現できない欄があります：" + "、".join(f"{k}（{v}）" for k, v in fields.items()))


def calculate(inputs=None, on_unknown="raise"):
    values = dict(DEFAULTS)
    values.update(inputs or {})
    out, bad = {}, {}
    for r in RULES:
        if r["js"] is None:
            bad[r["field"]] = r["why"]
            if on_unknown == "mark":
                out[r["field"]] = None
            continue
        v = @@simple_calculate@@(r["js"], lambda n: values.get(n))
        out[r["field"]] = v
        values[r["field"]] = v
    if bad and on_unknown != "mark":
        raise NotReproducible(bad)
    return out


if __name__ == "__main__":
    args = json.load(open(sys.argv[1], encoding="utf-8")) if len(sys.argv) > 1 else {}
    print(json.dumps(calculate(args, on_unknown="mark"), ensure_ascii=False, indent=1))
'''


def generate(doc, m: dict, name: str) -> str:
    from .vendor import ORDER, Vendor

    v = Vendor({"calculate", "RULES", "DEFAULTS", "SOURCE", "VERIFICATION", "NotReproducible", "json", "sys"})
    v.order = ["pdf.acrojs", *ORDER]
    v.load_registry()
    v.registry = {}  # シートの関数は使わない
    v.need("pdf.acrojs", "simple_calculate")
    vendored, final = v.source()
    rules = [{"id": d["id"], "field": d["field"], "js": d["source"] if d["evaluated"] else None,
              "why": None if d["evaluated"] else d.get("note")} for d in m["rules"]]
    from ..pdf.form import read_fields

    defaults = {f.name: _value(f) for f in read_fields(doc)}
    subs = {"@@SOURCE@@": json.dumps(name, ensure_ascii=False)[1:-1], "@@FILENAME@@": repr(name), "@@RULES@@": repr(rules), "@@DEFAULTS@@": repr(defaults),
            "@@VENDORED@@": vendored, "@@STDLIB@@": "\n".join(sorted(v.stdlib)),
            "@@simple_calculate@@": v.global_name(final, "pdf.acrojs", "simple_calculate")}
    return re.sub(r"@@[A-Za-z_]+@@", lambda x: subs.get(x[0], x[0]), _TEMPLATE)


def verify(source: str, target: Path, doc, m: dict) -> dict:
    """モジュールを別のプロセスで動かし、欄に保存された値と、とどくの計算と比べる。"""
    from ..pdf import acrojs
    from ..pdf.form import read_fields

    t0 = time.monotonic()
    fields = {f.name: _value(f) for f in read_fields(doc)}
    with tempfile.TemporaryDirectory(prefix="todoku-verify-") as td:
        mod = Path(td) / "module_under_test.py"
        mod.write_text(source, encoding="utf-8")
        p = subprocess.run([sys.executable, "-I", "-S", str(mod)], cwd=td, capture_output=True, text=True, timeout=60)
    status, mism = {}, []
    if p.returncode != 0:
        return {"ran": False, "error": p.stderr[-2000:], "rule_status": {d["id"]: {"status": "not_proven", "why": ["モジュールが動かなかった"]} for d in m["rules"]},
                "stored": {"compared": 0, "match": 0, "rate": "比べた欄がありません"}, "mismatches": [], "mismatch_count": 0, "seconds": round(time.monotonic() - t0, 2)}
    got = json.loads(p.stdout)
    compared = match = 0
    for d in m["rules"]:
        if not d["evaluated"]:
            status[d["id"]] = {"status": "not_proven", "why": [d.get("note", "評価していない")]}
            continue
        ours = acrojs.simple_calculate(d["source"], lambda n: fields.get(n))
        mod_v = got.get(d["field"])
        if ours is None or mod_v is None or abs(ours - mod_v) > 1e-9 * max(1.0, abs(ours)):
            mism.append({"field": d["field"], "rule": d["id"], "kind": "evaluator", "module": mod_v, "evaluator": ours})
            status[d["id"]] = {"status": "mismatch", "why": ["とどくの計算と違う"]}
            continue
        stored = fields.get(d["field"])
        if stored in (None, ""):
            status[d["id"]] = {"status": "not_proven", "why": ["欄に値が保存されていないので、保存された値と比べられない"]}
            continue
        try:
            sv = acrojs._number(stored)
        except acrojs.JsError:
            sv = None
        compared += 1
        tol = _display_tolerance(stored)
        if sv is not None and abs(sv - mod_v) <= tol:
            match += 1
            status[d["id"]] = {"status": "proven"}
        else:
            mism.append({"field": d["field"], "rule": d["id"], "kind": "stored", "module": mod_v, "stored": stored})
            status[d["id"]] = {"status": "mismatch", "why": [f"保存された値（{stored}）と違う"]}
    rate = "比べた欄がありません" if not compared else ("100%" if match == compared else f"{100 * match / compared:.2f}%")
    proven = sum(1 for s in status.values() if s["status"] == "proven")
    return {"ran": True, "standalone": "別のプロセスで、とどくを読み込めない状態（python -I -S）で動かしました",
            "stored": {"compared": compared, "match": match, "rate": rate}, "mismatches": mism, "mismatch_count": len(mism),
            "rules_proven": proven, "rules_not_proven": len(status) - proven, "rule_status": status, "notes": [],
            "seconds": round(time.monotonic() - t0, 2)}


def _display_tolerance(stored) -> float:
    """保存された字の小数の桁で丸めた差までは同じとみなす（書式で丸めて保存するアプリがある）。"""
    if isinstance(stored, str) and "." in stored:
        d = len(re.sub(r"[^\d]", "", stored.rsplit(".", 1)[-1]))
        return 0.5 * 10 ** -d + 1e-9
    return 1e-9


def extract(blob: bytes, name: str, password: str | None, fmt: str, out_path: str | None, overwrite: bool) -> dict:
    from . import ProcessingError
    from ..pdf.doc import PdfDocument

    doc = PdfDocument(blob, name, password=password) if password else PdfDocument(blob, name)
    res = model(doc)
    if res["xfa"]:
        res["note"] = "XFA のフォームです。XFA の計算のスクリプトは、窓の「スクリプト」の節にそのまま出ます（ここには AcroForm の分だけ）"
    s = res["summary"]
    res["say_to_user"] = (f"PDF のフォームの処理：欄 {s['fields']} 個、計算の規則 {s['rules']} 個（再現する {s['evaluated_rules']}・"
                          f"独自の JavaScript などで再現しない {s['not_evaluated_rules']}）。")
    if fmt == "python":
        if not out_path:
            raise ProcessingError('format="python" には、書き出す先 out_path（.py）が要ります')
        target = Path(out_path).expanduser()
        if target.exists() and not overwrite:
            raise ProcessingError(f"書き出す先「{target.name}」には、すでにファイルがあります（書き出していません）。"
                                  "置き換えてよいと本人が言ったときだけ overwrite に true を入れるか、別の名前を指定してください")
        source = generate(doc, res, name)
        report = verify(source, target, doc, res)
        statuses = report.pop("rule_status")
        for d in res["rules"]:
            if d["id"] in statuses:
                d["verification"] = statuses[d["id"]]
        final = source.replace("\nVERIFICATION = None  # @@todoku-verification@@\n", "\nVERIFICATION = " + repr(
            {k: report.get(k) for k in ("ran", "stored", "rules_proven", "rules_not_proven", "mismatch_count")}) + "\n", 1)
        target.parent.mkdir(parents=True, exist_ok=True)
        tmp = target.with_name(target.name + ".tmp")
        tmp.write_text(final, encoding="utf-8")
        os.replace(tmp, target)
        res["python"] = {"out_path": str(target), "bytes": len(final.encode("utf-8"))}
        res["verification"] = report
    elif out_path:
        target = Path(out_path).expanduser()
        if target.exists() and not overwrite:
            raise ProcessingError(f"書き出す先「{target.name}」には、すでにファイルがあります（書き出していません）")
        target.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        res["out_path"] = str(target)
    return res
