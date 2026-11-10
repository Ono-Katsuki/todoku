"""コマンドライン：todoku convert / verify / fill / apply / example / audit / bake / diagnose"""

from __future__ import annotations

from .book import open_book

import argparse
import hashlib
import json
import sys
from pathlib import Path

from .formula.values import ExcelError
from .xlsx import window
from .xlsx.calc import verify
from .xlsx.diagnose import diagnose
from .xlsx.edit import EditError, fill
from .xlsx.proof import prove
from . import report


def _json_default(o):
    if isinstance(o, ExcelError):
        return o.code
    if isinstance(o, set):
        return sorted(o)
    return str(o)


def _dump(obj, path: Path) -> None:
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=1, default=_json_default), encoding="utf-8")


def _load(path: str):
    from .book import FormatError, open_book

    p = Path(path)
    try:
        return open_book(p.read_bytes(), p.name)
    except FormatError as e:
        sys.exit(f"{p.name}: {e}")


def _kind(path: str) -> str:
    from .book import kind_of

    return kind_of(Path(path).read_bytes())


def _word_doc(blob: bytes, name: str):
    from .book import kind_of

    if kind_of(blob) == "doc":
        from .doc.book import DocFile

        return DocFile(blob, name)
    from .docx.document import DocxDocument

    return DocxDocument(blob, name)


def _docx_convert(a) -> int:
    from .docx.window import build, markdown_parts

    blob = Path(a.file).read_bytes()
    doc = build(_word_doc(blob, Path(a.file).name))
    out = Path(a.out or ".")
    out.mkdir(parents=True, exist_ok=True)
    stem = Path(a.file).stem
    clean = {**doc, "stories": [{k: v for k, v in st.items() if not k.startswith("_")} for st in doc["stories"]]}
    _dump(clean, out / f"{stem}.window.json")
    parts = markdown_parts(doc)
    (out / f"{stem}.window.md").write_text(parts[0][1], encoding="utf-8")
    for i, (_, text) in enumerate(parts[1:], 1):
        (out / f"{stem}.window.part{i:02d}.md").write_text(text, encoding="utf-8")
    print(f"Word: 入力欄 {len(doc['fields'])} か所" + (f"・保護 {doc['protection']['edit']}" if doc["protection"]["enforced"] else ""))
    print(f"出力: {out}/{stem}.window.md と window.json")
    return 0


def _pdf_convert(a) -> int:
    from .pdf.doc import PdfDocument
    from .pdf.window import build, markdown_parts

    blob = Path(a.file).read_bytes()
    doc = build(PdfDocument(blob, Path(a.file).name))
    out = Path(a.out or ".")
    out.mkdir(parents=True, exist_ok=True)
    stem = Path(a.file).stem
    clean = {**doc, "pages": [{k: v for k, v in pg.items() if not k.startswith("_")} for pg in doc["pages"]]}
    _dump(clean, out / f"{stem}.window.json")
    parts = markdown_parts(doc)
    (out / f"{stem}.window.md").write_text(parts[0][1], encoding="utf-8")
    for i, (_, text) in enumerate(parts[1:], 1):
        (out / f"{stem}.window.part{i:02d}.md").write_text(text, encoding="utf-8")
    n = sum(len(p["fields"]) for p in doc["pages"])
    unmapped = sum(p["unmapped"] for p in doc["pages"])
    print(f"PDF: {len(doc['pages'])} ページ・入力欄 {n} か所・文字に直せない字 {unmapped} 字"
          + ("・フォームの欄あり" if doc["form"]["acroform_fields"] else "・印刷用（入力欄は配置から推定）"))
    print(f"出力: {out}/{stem}.window.md と window.json")
    return 0 if unmapped == 0 else 1


def cmd_convert(a) -> int:
    k = _kind(a.file)
    if k == "pdf":
        return _pdf_convert(a)
    if k in ("docx", "doc"):
        return _docx_convert(a)
    if k == "jtd":
        return _jtd_convert(a)
    wb = _load(a.file)
    out = Path(a.out or ".")
    out.mkdir(parents=True, exist_ok=True)
    stem = Path(a.file).stem
    rebuilt = wb.zip.build()
    rt_ok = rebuilt == wb.zip.original
    rep = verify(wb)
    doc = window.build(wb)
    doc["roundtrip"] = {"sha256_original": hashlib.sha256(wb.zip.original).hexdigest(),
                        "sha256_rebuilt": hashlib.sha256(rebuilt).hexdigest(), "identical": rt_ok}
    doc["verify"] = rep.summary()
    diag = diagnose(wb, rep)
    _dump(doc, out / f"{stem}.window.json")
    parts = window.markdown_parts(doc)
    (out / f"{stem}.window.md").write_text(parts[0][1], encoding="utf-8")
    for i, (_, text) in enumerate(parts[1:], 1):
        # 大きな窓は、AIが1回で読み切れる大きさに分ける（最初のファイルが目次）
        (out / f"{stem}.window.part{i:02d}.md").write_text(text, encoding="utf-8")
    _dump(diag.to_dict(), out / f"{stem}.diagnosis.json")
    (out / f"{stem}.diagnosis.html").write_text(report.render(diag.to_dict(), title=f"診断：{Path(a.file).name}"), encoding="utf-8")
    _dump([r.__dict__ for r in rep.results], out / f"{stem}.verify.json")
    s = rep.summary()
    print(f"往復: {'バイト一致' if rt_ok else '不一致'}（SHA-256 {doc['roundtrip']['sha256_original'][:16]}…）")
    print(f"関数: {s['match']}/{s['total']} 一致（不一致 {s['mismatch']}・再計算不可 {s['unsupported']}・揮発性 {s['volatile']}）")
    sc = diag.score()
    print(f"診断: 到達 {sc['reached_name']}（規則だけでは {sc['rules_reached_name']}）　壁 {sum(1 for f in diag.findings if f.severity == 'blocker')} 件・"
          f"規則では越えられない所 {sum(1 for f in diag.findings if f.severity == 'agent')} 件・注意 {sum(1 for f in diag.findings if f.severity == 'warning')} 件")
    print(f"出力: {out}/{stem}.window.md ほか4ファイル")
    return 0 if rt_ok and s["mismatch"] == 0 else 1


def cmd_verify(a) -> int:
    wb = _load(a.file)
    rebuilt = wb.zip.build()
    rt_ok = rebuilt == wb.zip.original
    rep = verify(wb)
    s = rep.summary()
    print(f"往復: {'バイト一致' if rt_ok else '不一致'}")
    print(f"関数: {s['match']}/{s['total']} 一致（不一致 {s['mismatch']}・再計算不可 {s['unsupported']}・計算のたびに変わる {s['volatile']}）")
    for r in rep.results:
        if r.status not in ("match",) and (a.verbose or r.status == "mismatch"):
            print(f"  [{r.status}] {r.sheet}!{r.address} ={r.formula}  保存値={r.expected!r} 再計算={r.actual!r} {r.reason}")
    return 0 if rt_ok and s["mismatch"] == 0 else 1


def _parse_sets(sets: list[str]) -> dict[str, str]:
    out = {}
    for s in sets or []:
        if "=" not in s:
            sys.exit(f"--set は 'シート!セル=値' の形で指定してください: {s}")
        k, v = s.split("=", 1)
        out[k.strip()] = v
    return out


def _write_fill(path: str, values: dict, out: str, strict: bool = False) -> int:
    """書き込み（どの形式でも、エージェントの道具と同じ道 api.fill）：証明・読み戻し・もとからあった字の確かめを、同じ所で行う。
    結果の全部を「（出力）.proof.json」に書く。終わりの値：0 証明も読み戻しも通った、1 通らなかった（「.unverified」の名前で書き出す）、
    2 書いていない。"""
    from . import api
    from .book import kind_of

    k = kind_of(Path(path).read_bytes())
    excel = k in ("xlsx", "xls")
    if strict and excel:
        # --strict：警告にとどまる入力規則の問題も受け付けない（書く前に、同じ値で確かめる）
        try:
            fill(_load(path), values, strict=True)
        except EditError as e:
            print(f"書き込めません: {e}", file=sys.stderr)
            return 2
    try:
        r = api.fill(path, values, out, overwrite=True)
    except api.ApiError as e:
        print(f"書き込めません: {e}", file=sys.stderr)
        return 2
    if "out_path" not in r:  # 長い仕事にした：もう一度同じ呼び出しをすると、進み具合か結果が返る
        print(r.get("message") or json.dumps(r, ensure_ascii=False, default=_json_default))
        return 1
    ok = bool(r.get("proof_ok") and r.get("read_back_ok"))
    print(r.get("proof") or "")
    for w in r.get("warnings") or []:
        print(f"注意: {w}")
    if r.get("structure_errors"):
        print("構造の問題:", *r["structure_errors"][:20], sep="\n  ")
    if r.get("unexpected"):
        print("想定外の変更:", *r["unexpected"], sep="\n  ")
    if not r.get("read_back_ok"):
        print("読み戻しの違い（書いた後のファイルを読み直した値が、書いた値と違う）:",
              *(json.dumps(d, ensure_ascii=False, default=_json_default) for d in (r.get("read_back_differences") or [])[:20]), sep="\n  ")
    print(f"元 SHA-256: {r.get('original_sha256')}")
    print(f"後 SHA-256: {r.get('new_sha256')}")
    print(r.get("say_to_user") or "")
    out_path = Path(r["out_path"])
    print(f"出力: {out_path}")
    _dump(r, out_path.with_suffix(".proof.json"))
    if excel:
        try:
            wb = open_book(out_path.read_bytes(), out_path.name)
            proof = {"ok": ok, "headline": r.get("proof") or "", "original_sha256": r.get("original_sha256"), "new_sha256": r.get("new_sha256"),
                     "unexpected": list(r.get("unexpected") or []), "cell_changes": _cell_changes(_load(path), r)}
            out_path.with_suffix(".proof.html").write_text(
                report.render(diagnose(wb, verify(wb)).to_dict(), proof=proof, title=f"記入の証明：{out_path.name}"), encoding="utf-8")
        except Exception as e:  # noqa: BLE001  報告のページを作れないだけ（書き込みと証明は終わっている）
            print(f"注意: 証明のページ（.proof.html）を作れませんでした（{type(e).__name__}）")
    return 0 if ok else 1


def _cell_changes(wb, r: dict) -> list[dict]:
    """証明のページの表：書いたセルと、計算し直したセルの、書く前の値 → 後の値。"""
    from .formula.parser import addr
    from .xlsx.edit import resolve_target

    out = []
    written = r.get("written") or {}
    if not isinstance(written, dict):
        written = {x.get("id"): x.get("value") for x in written if isinstance(x, dict)}
    for kind, d in (("input", written), ("formula", r.get("recalculated") or {})):
        for key, after in d.items():
            try:
                s, row, col = resolve_target(wb, key)
            except Exception:  # noqa: BLE001  部品の名前など、セルでない所
                continue
            cell = wb.sheet(s).cells.get((row, col))
            out.append({"sheet": s, "cell": addr(row, col), "kind": kind, "before": cell.value if cell is not None else None, "after": after})
    return out


def cmd_fill(a) -> int:
    values = _parse_sets(a.set)
    if a.json:
        values.update(json.loads(Path(a.json).read_text(encoding="utf-8")))
    if not values:
        sys.exit("書き込む値がありません（--set か --json）")
    return _write_fill(a.file, values, a.out, a.strict)


def cmd_apply(a) -> int:
    """窓（window.json）で書き換えた入力欄の値を、元のファイルに差分として適用する。"""
    doc = json.loads(Path(a.window).read_text(encoding="utf-8"))
    wb = _load(a.original)
    sha = hashlib.sha256(wb.zip.original).hexdigest()
    if sha != doc["source"]["sha256"]:
        sys.exit("この窓は別のファイルから作られています（SHA-256が一致しません）")
    base = window.build(wb)
    before = {(s["name"], c["cell"]): c.get("value") for s in base["sheets"] for c in s["cells"]}
    values = {}
    for s in doc["sheets"]:
        for c in s["cells"]:
            if not c.get("input"):
                continue
            v = c.get("value")
            if before.get((s["name"], c["cell"])) != v:
                values[f"{s['name']}!{c['cell']}"] = v
    if not values:
        Path(a.out).write_bytes(wb.zip.build())
        print("変更はありません。元のファイルとバイト一致のファイルを書き出しました。")
        return 0
    return _write_fill(a.original, values, a.out, a.strict)


def cmd_example(a) -> int:
    """記入例と照らす。Excel：記入例をExcelの答えとして、記入と再計算を確かめる。
    PDF・Word：記入例で書き込まれた所が、白紙で見つけた入力欄にあるか（入力欄の見落とし）を確かめる。"""
    k = _kind(a.blank)
    if k in ("pdf", "docx", "doc"):
        if k == "pdf":
            from .pdf.example import check as pdf_check

            r = pdf_check(Path(a.blank).read_bytes(), Path(a.example).read_bytes())
            print(r.headline())
            for m in r.missed[:40]:
                where = f"（{m.near}の近く）" if m.near else ""
                print(f"  見落としの候補: {m.page}ページ {m.text!r} 位置 {m.rect}{where}")
        else:
            from .docx.example import check as word_check

            r = word_check(Path(a.blank).read_bytes(), Path(a.example).read_bytes())
            print(r.headline())
            for m in r.missed[:40]:
                print(f"  見落としの候補: 段落 {m.para} {m.text!r}（{m.context}）")
        return 0 if r.ok else 1
    from .xlsx.example import check

    r = check(Path(a.blank).read_bytes(), Path(a.example).read_bytes())
    print(r.headline())
    if r.proof is not None:
        print(r.proof.headline())
        print("構造:", "Excelが修復なしに開くための条件を満たしています" if not r.proof.lint_errors else r.proof.lint_errors[:5])
        for u in r.proof.unexpected[:10]:
            print(f"  想定外の変更: {u}")
    print(f"入力欄の見落とし: {len(r.not_input_fields)} か所" + (f"（{', '.join(r.not_input_fields[:10])}）" if r.not_input_fields else ""))
    if r.template_differs:
        print(f"（様式の版の違い {len(r.template_differs)} か所：見出しの文字などが記入例と違うので、そろえてから比べました）")
    if r.structure_differs:
        print(f"（数式が異なり比較から外したセル {len(r.structure_differs)} 個）")
    for m in r.mismatches[:20]:
        print(f"  差: {m[0]}!{m[1]} ={m[2][:90]}  Excel={m[3]!r}  こちら={m[4]!r}")
    for m in r.control_mismatches[:20]:
        print(f"  チェックボックスの差: {m[0]}!{m[1]}  Excel={m[2]}  こちら={m[3]}")
    return 0 if r.ok else 1


def cmd_audit(a) -> int:
    """エージェントがこの書類を全て正確に読み書きできるかを確かめる。"""
    from .xlsx.audit import run

    if not a.example:
        from . import api

        lz = api.lazy_open(a.file)
        if lz is not None:
            return _audit_lazy(a)
    blob = Path(a.file).read_bytes()
    k = _kind(a.file)
    if k == "pdf":
        from .pdf.audit import run as pdf_run

        au = pdf_run(blob, Path(a.file).name)
    elif k == "docx":
        from .docx.audit import run as docx_run

        au = docx_run(blob, Path(a.file).name)
    elif k == "doc":
        from .doc.audit import run as doc_run

        au = doc_run(blob, Path(a.file).name)
    elif k == "jtd":
        from .jtd.audit import run as jtd_run

        au = jtd_run(blob, Path(a.file).name)
    else:
        au = run(blob, Path(a.file).name, Path(a.example).read_bytes() if a.example else None)
    for c in au.checks:
        mark = ("OK・判断はエージェント" if c.agent else "OK") if c.ok else "要確認"
        print(f"[{mark}] {c.name}: {c.detail}")
        for x in c.items[:8]:
            print(f"      - {x}")
        if len(c.items) > 8:
            print(f"      … ほか {len(c.items) - 8} 件")
    print("結論:", "全て正確に読み書きできます" if au.ok else "読み書きに確認が必要な点があります")
    _print_grade(a.file)
    if a.json:
        Path(a.json).write_text(json.dumps(au.to_dict(), ensure_ascii=False, indent=1), encoding="utf-8")
    return 0 if au.ok else 1


def _audit_lazy(a) -> int:
    """大きな書類の検査：分けて読む道で（メモリは書類の大きさに比例しない。長い仕事と同じ手順を前で動かす）。"""
    from . import lazy, limits

    def report(text: str, done: int = 0, total: int = 0) -> None:
        print(f"\r{text} {done:,}/{total:,}" if total else f"\r{text}", end="", file=sys.stderr, flush=True)

    limits.set_progress_reporter(report)
    with limits.call_budget(limits.Budget(total=None)):
        doc = lazy.open_doc(a.file, None, None, force=True)
        state = doc.audit_begin()
        while not state.get("done"):
            state = doc.audit_step(state)
    limits.set_progress_reporter(None)
    print(file=sys.stderr)
    checks = state.get("checks") or []
    ok = all(c.get("ok") for c in checks)
    for c in checks:
        mark = ("OK・判断はエージェント" if c.get("agent") else "OK") if c.get("ok") else "要確認"
        print(f"[{mark}] {c['name']}: {c.get('detail', '')}")
        items = list(c.get("items") or [])
        for x in items[:8]:
            print(f"      - {x}")
        if len(items) > 8:
            print(f"      … ほか {len(items) - 8} 件")
    for x in state.get("still_deferred") or []:
        print(f"[未確認] {x}")
    print("結論:", "全て正確に読み書きできます" if ok and not state.get("still_deferred") else ("確かめた所は正確に読み書きできます（未確認の所があります）" if ok else "読み書きに確認が必要な点があります"))
    if a.json:
        Path(a.json).write_text(json.dumps(state.get("full") or {"checks": checks, "still_deferred": state.get("still_deferred") or []}, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    return 0 if ok else 1


def cmd_bake(a) -> int:
    """計算結果が保存されていない数式に、計算した値を書き込む。"""
    from .xlsx.edit import bake

    wb = _load(a.file)
    res = bake(wb)
    Path(a.out).write_bytes(res.blob)
    pr = prove(wb.zip.original, res.blob, set())
    print(f"計算結果を書き込んだ数式: {len(res.patches)} 個")
    print(pr.headline())
    return 0 if pr.ok else 1


def _print_grade(path: str) -> None:
    from .grade import grade_blob

    g = grade_blob(Path(path).read_bytes(), Path(path).name)
    print(g.headline())
    for label, reasons, plain in (("読む", g.reasons, "SA"), ("書く", g.write_reasons, "A")):
        for r in reasons:
            mark = "　" if r.grade in plain else "▶"
            print(f"  {mark}{label}[{r.grade}] {r.title}：{r.detail}" + (f"（{', '.join(r.where[:5])}）" if r.where else ""))
    for c in g.agent_checks:
        label = "書く" if c["score"] == "write" else "読む"
        print(f"  ？エージェントが確かめること（あれば{label}段階 {c['grade']}）：{c['question']}")


def cmd_diagnose(a) -> int:
    _print_grade(a.file)
    if _kind(a.file) not in ("xlsx", "xls"):
        return 0
    wb = _load(a.file)
    diag = diagnose(wb)
    if a.json:
        print(json.dumps(diag.to_dict(), ensure_ascii=False, indent=1, default=_json_default))
    else:
        for f in diag.findings:
            print(f"[{f.level}.{['届く', '読める', '操作できる', '検証できる'][f.level]}][{f.severity}] {f.title}：{f.detail}")
            print(f"    直し方：{f.fix}")
    if a.html:
        from .grade import grade_blob

        g = grade_blob(Path(a.file).read_bytes(), Path(a.file).name).to_dict()
        Path(a.html).write_text(report.render(diag.to_dict(), title=f"診断：{Path(a.file).name}", grade=g), encoding="utf-8")
    return 0


def cmd_forget(a) -> int:
    from . import jobs

    gone = jobs.purge(everything=True)
    print(f"控えを {len(gone)} 個消しました（{jobs._dir()}）")
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="todoku", description="書類を、内容を一字も変えずにAIと支援技術に届ける")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("convert", help="読むための窓（JSON・Markdown）と診断を作る")
    p.add_argument("file")
    p.add_argument("-o", "--out")
    p.set_defaults(func=cmd_convert)
    p = sub.add_parser("verify", help="往復のバイト一致と、関数の全数照合")
    p.add_argument("file")
    p.add_argument("-v", "--verbose", action="store_true")
    p.set_defaults(func=cmd_verify)
    p = sub.add_parser("fill", help="入力欄に書き込む（変えるのは書き込んだ値と、その計算結果・見た目だけ）")
    p.add_argument("file")
    p.add_argument("--set", action="append", help="'シート!セル=値'、PDF は '欄の名前=値'（何度でも）")
    p.add_argument("--json", help='{"シート!セル": 値} のJSONファイル')
    p.add_argument("-o", "--out", required=True)
    p.add_argument("--strict", action="store_true", help="入力規則に合わない値があれば書き込まない")
    p.set_defaults(func=cmd_fill)
    p = sub.add_parser("apply", help="窓で書き換えた値を、元のファイルに適用する")
    p.add_argument("window")
    p.add_argument("--original", required=True)
    p.add_argument("-o", "--out", required=True)
    p.add_argument("--strict", action="store_true")
    p.set_defaults(func=cmd_apply)
    p = sub.add_parser("example", help="記入例と照らす（Excel：記入と再計算、PDF・Word：入力欄の見落とし）")
    p.add_argument("blank")
    p.add_argument("example")
    p.set_defaults(func=cmd_example)
    p = sub.add_parser("audit", help="エージェントが全て正確に読み書きできるかを確かめる（読み落とし・全欄への書き込み・読み戻し）")
    p.add_argument("file")
    p.add_argument("--example", help="記入例（あれば、Excelの答えとも照らす）")
    p.add_argument("--json", help="結果をJSONで保存する")
    p.set_defaults(func=cmd_audit)
    p = sub.add_parser("bake", help="計算結果が保存されていない数式に、Excelと同じ規則で計算した値を書き込む")
    p.add_argument("file")
    p.add_argument("-o", "--out", required=True)
    p.set_defaults(func=cmd_bake)
    p = sub.add_parser("forget", help="長い仕事の控え（窓・検査の結果。書類の中身が入る）を、動いている仕事のほかはすぐに消す")
    p.set_defaults(func=cmd_forget)
    p = sub.add_parser("diagnose", help="作り手向けの診断")
    p.add_argument("file")
    p.add_argument("--json", action="store_true")
    p.add_argument("--html")
    p.set_defaults(func=cmd_diagnose)
    a = ap.parse_args(argv)
    return a.func(a)


if __name__ == "__main__":
    sys.exit(main())


# ====================================================================== 一太郎（.jtd）
def _jtd_convert(a) -> int:
    """一太郎の窓（Markdown と JSON）を書き出す。"""
    from . import api

    _k, doc, parts = api.window(a.file)
    out = Path(a.out or ".")
    out.mkdir(parents=True, exist_ok=True)
    stem = Path(a.file).stem
    (out / f"{stem}.window.md").write_text("\n\n".join(t for _, t in parts), encoding="utf-8")
    _dump({k: v for k, v in doc.items() if k != "stories"} | {"stories": [{x: y for x, y in st.items() if x != "_plain"}
                                                                          for st in doc["stories"]]},
          out / f"{stem}.window.json")
    print(f"窓：{out / (stem + '.window.md')}（入力欄 {len(doc['fields'])} か所）")
    return 0
