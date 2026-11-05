"""書類の処理を、処理として取り出す（extract_processing）。

データだけを渡して処理をエージェントの推測に任せるのではなく、書類の中の計算の決まりを、そのまま使える形で渡す：
- 規則：同じ形（R1C1 で同じ）の数式を、まとまりごとに 1 つの規則にする（「D3:D100000 = B{row}*C{row}」）。例外（形の違うセル・上書きした定数）も添える
- 構造：入力・パラメータ・参照表（データとして）・名前・出力・規則の順番と循環・分岐・配列数式・外部参照・動く参照・隠れた所・シートのあいだの流れ
- マクロ：手続きのソースと、読み書きするセル（字で決まる所だけ。決まらない所は「分からない」と言う）、いつ動くか
- 再利用：format="python" で、標準ライブラリだけで動く calculate(inputs) のモジュールを書き出し、確かめてから返す
- PDF のフォーム：計算の順と、欄のあいだの規則（独自の JavaScript は計算せず、そのまま示す）

行の多いブックでもセルごとの控えを持たない（scan）。Python の書き出しの確かめだけは、計算した値を別のプロセスの中で一時的に持つ。
"""

from __future__ import annotations

import json
from pathlib import Path


class ProcessingError(ValueError):
    pass


LIMITS = {
    "dependencies": "規則の依存は数式の参照から静的にたどる。INDIRECT・OFFSET・範囲の端が計算で決まる参照は、dynamic_refs に、分からない所とともに示す",
    "outputs": "出力は「ほかの数式に読まれないセル」。条件付き書式・入力規則・マクロ・グラフが読むことは数えない",
    "conditional_formats_and_validations": "条件付き書式と入力規則は、それぞれ規則の種類として並べる（セルの値は変えない）。Python の書き出しには含めない",
    "not_expressible_by_cells": "ピボットテーブル・Power Query・データ接続は、セルの数式では書けない処理として、定義（元の範囲・項目など）だけを示す。計算は再現しない",
}


def _source(path: str, blob: bytes, name: str, k: str, password: str | None):
    """(読み口, 流して読んでいるか)。"""
    from .scan import BookSource, XlsxSource

    if k == "xlsx":
        from ..api import lazy_open
        from ..lazy.xlsxdoc import XlsxLazy

        try:
            lz = lazy_open(path, password, force=True)
        except Exception:  # noqa: BLE001  行の番号とセルの番地が合わない形など：丸ごと読む道へ
            lz = None
        if isinstance(lz, XlsxLazy):
            return XlsxSource(lz), True
    return BookSource(_open_workbook(blob, name, k)), False


def _open_workbook(blob: bytes, name: str, k: str):
    if k in ("ods", "fods"):
        from ..odf.sheet import OdsWorkbook

        return OdsWorkbook(blob, name)
    from ..book import open_book

    return open_book(blob, name)


def model(src, blob: bytes | None, k: str) -> dict:
    """ブックの処理の全体：{rules, inputs, ...}。内部の結果（_rules など）も入れる。"""
    from . import extras, macrostatic, rules as R, scan, structure as S

    sheets, shaper = scan.scan_book(src)
    wb = src.meta
    rules = R.build_blocks(sheets, shaper)
    R.analyse_reads(rules, sheets, wb)
    ix = R.link_rules(rules)
    ord_ = R.order_rules(rules, ix)
    R.exceptions_of(rules, sheets, ix, src)
    st = S.analyse(rules, sheets, src, ix, extras.external_targets(wb))
    labels = st.pop("labels")
    macro = macrostatic.analyse(wb, rules, ix, sheets)
    formulas = sum(s.formulas for s in sheets)
    cells = sum(s.formula_cells for s in sheets)
    out = {
        "kind": k,
        "summary": {"formula_elements": formulas, "formula_cells": cells, "rules": len(rules),
                    "distinct_shapes": len({u.shape.id for u in rules}),
                    "exceptions": sum(u.exception_count for u in rules),
                    "sheets": [{"name": s.name, "state": s.state, "formula_cells": s.formula_cells, "constants": s.constants,
                                "max_row": s.max_row, "max_col": s.max_col} for s in sheets]},
        "rules": [S.rule_json(u, ix, labels, wb) for u in rules],
        "order": ord_["order"],
        "cycles": ord_["cycles"],
        **st,
        "named_ranges": S.names_json(wb, rules, sheets),
        "conditional_formats": extras.conditional_formats(wb),
        "data_validations": extras.data_validations(wb),
        "macros": macro,
        "not_expressible_by_cells": extras.not_cell_processing(blob, wb, k),
        "limits": LIMITS,
    }
    from .consistency import check

    # 説明の字（text・a1・reads）と数式が同じ所を指すかを、説明の字を解析し直して確かめる（確かめた数も結果に残す）
    out["explanation_check"] = check(out["rules"], {u.id: u.origin for u in rules}, S.LISTED)
    book = getattr(src, "wb", None)
    if book is not None:  # 行を少しずつ読む大きなブックでは、数式を一度に読まないので見ない
        from ..xlsx.textorder import text_order_cells

        order = [{"sheet": sh.name, **x} for sh in book.sheets if hasattr(sh, "cells") for x in text_order_cells(book, sh)]
        if order:
            out["text_order_lookups"] = order
    out["_internal"] = {"rules": rules, "sheets": sheets, "ix": ix, "shaper": shaper}
    return out


def _say(res: dict) -> str:
    s = res["summary"]
    lines = [f"処理の取り出し：数式のセル {s['formula_cells']:,} 個 → 規則 {s['rules']:,} 個（形は {s['distinct_shapes']:,} 種類）。"
             f"入力 {len(res.get('inputs', [])):,} か所、パラメータ {len(res.get('parameters', [])):,}、参照表 {len(res.get('lookup_tables', [])):,}。"]
    cyc = [c for c in res.get("cycles", []) if c["kind"] != "ordered"]
    if cyc:
        lines.append(f"循環：{len(cyc)} か所（{'、'.join(c['kind'] for c in cyc[:5])}）。")
    if res.get("external_references"):
        lines.append(f"外部のブックへの参照：{len(res['external_references'])} か所。")
    if res.get("unknown_functions"):
        lines.append(f"とどくが計算できない関数：{'、'.join(res['unknown_functions'])}。")
    if res.get("not_expressible_by_cells"):
        lines.append(f"セルの数式で書けない処理：{len(res['not_expressible_by_cells'])} 個。")
    if res.get("text_order_lookups"):
        lines.append(f"文字の並び順で結果が変わる数式：{len(res['text_order_lookups'])} か所（text_order_lookups。近似一致で文字を探すので、Excel の環境で値が変わることがあります）。")
    ec = res.get("explanation_check") or {}
    if ec.get("failed"):
        lines.append(f"説明の字が数式と合わない規則が {ec['failed']} 個あります（explanation_check.failures）。その規則は text・a1 でなく formula を読んでください。")
    return "\n".join(lines)


def extract(path: str, blob: bytes, name: str, k: str, password: str | None = None, fmt: str = "json",
            out_path: str | None = None, sheet: str | None = None, overwrite: bool = False, budget_chars: int | None = None) -> dict:
    """api.extract_processing の中身。blob は読んだファイルの中身。"""
    if k == "pdf":
        from . import pdfcalc

        return pdfcalc.extract(blob, name, password, fmt, out_path, overwrite)
    src, streaming = _source(path, blob, name, k, password)
    res = model(src, blob, k)
    internal = res.pop("_internal")
    res["streamed"] = streaming
    if fmt == "python":
        if not out_path:
            raise ProcessingError('format="python" には、書き出す先 out_path（.py）が要ります')
        from . import pygen, verify

        target = Path(out_path).expanduser()
        if target.exists() and not overwrite:
            raise ProcessingError(f"書き出す先「{target.name}」には、すでにファイルがあります（書き出していません）。"
                                  "置き換えてよいと本人が言ったときだけ overwrite に true を入れるか、別の名前を指定してください")
        gen = pygen.generate(src, internal, res, name, password_note=bool(password))
        report = verify.run_and_compare(gen, target, src, internal, res, blob, name, k)
        res["python"] = {"out_path": str(target), **gen.facts}
        res["verification"] = report
        statuses = report.pop("rule_status")
        for d in res["rules"]:
            st = statuses.get(d["id"])
            if st:
                d["verification"] = st
    res["say_to_user"] = _say(res)
    if "verification" in res:
        v = res["verification"]
        res["say_to_user"] += (f"\nPython の書き出し：{res['python']['out_path']}。保存された値との一致 {v['stored']['match']:,} / {v['stored']['compared']:,}"
                               f"（{v['stored']['rate']}）。確かめられなかった規則 {v['rules_not_proven']} 個。")
    if sheet is not None:
        res = _only_sheet(res, sheet)
    if fmt == "json" and out_path:
        target = Path(out_path).expanduser()
        if target.exists() and not overwrite:
            raise ProcessingError(f"書き出す先「{target.name}」には、すでにファイルがあります（書き出していません）。"
                                  "置き換えてよいと本人が言ったときだけ overwrite に true を入れるか、別の名前を指定してください")
        target.write_text(json.dumps(res, ensure_ascii=False, indent=1), encoding="utf-8")
        res = _condensed(res, budget_chars, written=str(target))
    elif budget_chars is not None and len(json.dumps(res, ensure_ascii=False)) > budget_chars:
        res = _condensed(res, budget_chars)
    return res


def _only_sheet(res: dict, sheet: str) -> dict:
    res = dict(res)
    res["rules"] = [d for d in res["rules"] if d["sheet"] == sheet]
    ids = {d["id"] for d in res["rules"]}
    res["order"] = [x for x in res["order"] if x in ids]
    for k in ("inputs", "parameters", "lookup_tables", "conditional_formats", "data_validations"):
        res[k] = [d for d in res.get(k, []) if d.get("sheet") == sheet]
    res["only_sheet"] = sheet
    return res


def _condensed(res: dict, budget: int | None, written: str | None = None) -> dict:
    """大きすぎて 1 度に返せないときの要約。全部は out_path（JSON）かシートごと（sheet）に取れる。"""
    budget = budget or 400000
    head = {k: res[k] for k in ("kind", "summary", "cycles", "unknown_functions", "cross_sheet_flow", "limits", "say_to_user", "python", "verification", "explanation_check", "text_order_lookups", "streamed", "only_sheet") if k in res}
    keep = []
    used = len(json.dumps(head, ensure_ascii=False))
    room = budget // 2
    for d in res["rules"]:
        n = len(json.dumps(d, ensure_ascii=False))
        if used + n > room:
            break
        keep.append(d)
        used += n
    head["rules"] = keep
    mac = res.get("macros")
    if mac and mac.get("procedures"):
        head["macros"] = {"present": True, "note": "ソースと読み書きする所は、out_path の JSON か sheet で取れます",
                          "procedures": [{"ref": p["ref"], **({"trigger": p["trigger"]} if p.get("trigger") else {})} for p in mac["procedures"]][:200]}
    head["rules_shown"] = len(keep)
    head["counts"] = {k: len(res.get(k) or []) for k in ("rules", "order", "inputs", "parameters", "lookup_tables", "named_ranges",
                                                          "conditional_formats", "data_validations", "not_expressible_by_cells")}
    if written:
        head["out_path"] = written
        head["note"] = "全部は out_path の JSON にあります（ここには、はじめの規則だけ）"
    else:
        head["note"] = "大きいので、はじめの規則だけを返します。全部は out_path（.json）に書き出すか、sheet でシートごとに取ってください"
    return head
