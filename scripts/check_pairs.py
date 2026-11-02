"""集めた「白紙の様式と記入例」の組で、入力欄の見落としを確かめる（scripts/collect_pairs.py で集めた corpus/pairs/）。

- PDF：記入例で書き込まれた所が、白紙で見つけた入力欄の中にあるか（todoku.pdf.example）
- Word：記入例で書き足された所が、白紙で見つけた入力欄にかかっているか（todoku.docx.example）
- Excel：記入例を Excel の答えとして、記入と再計算が一致するか（todoku.xlsx.example）
記入例のページのうち、白紙のどのページとも様式の字の大半が重ならないページ（記載要領など）は照らさない。
そういうページしかない組は、別の書類どうし（組の取り違え）として外す。

使い方: uv run python scripts/check_pairs.py
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "src"))


def main() -> int:
    base = ROOT / "corpus" / "pairs"
    idx = json.loads((base / "index.json").read_text(encoding="utf-8"))
    pairs = [x for x in idx if "blank_file" in x and "example_file" in x]
    only = set(sys.argv[1:])  # 形式をしぼる（例：doc docx）。しぼったときは check.json を書き換えない
    rows = []
    for x in pairs:
        b = (base / x["blank_file"]).read_bytes()
        e = (base / x["example_file"]).read_bytes()
        ext = x["blank_file"].rsplit(".", 1)[-1].lower()
        if only and ext not in only:
            continue
        t = time.time()
        row = {"blank": x["blank_text"][:40], "example": x["example_text"][:40], "kind": ext,
               "blank_file": x["blank_file"], "example_file": x["example_file"]}
        try:
            if ext == "pdf":
                from todoku.pdf.example import check

                r = check(b, e)
                ratio = r.changed_template / max(1, r.blank_chars)
                row.update({"added": r.added, "covered": r.covered, "covered_agent": r.covered_agent,
                            "missed": [m.__dict__ for m in r.missed],
                            "template_ratio": round(ratio, 2), "headline": r.headline(), "other_pages": r.other_pages})
                # 様式が合うページが1つもない組は、別の書類どうし（組の取り違え）。
                # 白紙に文字の情報がない（画像だけのページ：段階 E）組は、照らせないので外す
                row["paired"] = r.blank_chars > 0 and r.added > 0
                if not r.blank_text_chars:
                    row["excluded"] = "白紙に文字の情報がない（画像だけのページ：段階 E）"
                row["same_version"] = r.same_version
                row["page_kept"] = r.page_kept
            elif ext in ("docx", "doc"):
                from todoku.docx.example import check

                r = check(b, e)
                row.update({"added": r.added, "covered": r.covered, "covered_agent": r.covered_agent,
                            "missed": [m.__dict__ for m in r.missed], "drawn": r.drawn, "kept": r.kept,
                            "headline": r.headline(), "paired": r.added > 0, "same_version": r.same_version})
            elif ext in ("xlsx", "xlsm", "xls"):
                from todoku.xlsx.example import check

                r = check(b, e)
                row.update({"headline": r.headline(), "paired": True, "ok": r.ok,
                            "missed": [{"text": f} for f in r.not_input_fields]})
        except Exception as ex:  # noqa: BLE001
            row.update({"error": f"{type(ex).__name__}: {ex}"[:200], "paired": False})
        row["seconds"] = round(time.time() - t, 1)
        rows.append(row)
        mark = (row.get("excluded") or "外した") if not row.get("paired") else ("OK" if not row.get("missed") else "見落とし")
        if row.get("paired") and not row.get("same_version", True):
            mark += "・版違い"
        print(f"[{mark}] {row['kind']} {row['blank'][:24]} ↔ {row['example'][:24]} {row.get('headline', row.get('error', ''))[:160]}", flush=True)
    (base / ("check.json" if not only else "check_" + "_".join(sorted(only)) + ".json")).write_text(
        json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    used = [r for r in rows if r.get("paired")]
    n_e = sum(1 for r in rows if r.get("excluded"))
    n_err = sum(1 for r in rows if r.get("error"))
    print(f"\n組 {len(rows)}・照らした組 {len(used)}（白紙が画像だけで外した組 {n_e}・処理が落ちた組 {n_err}・"
          f"取り違えで外した組 {len(rows) - len(used) - n_e - n_err}）")
    for name, group in (("様式の版が同じ組", [r for r in used if r.get("same_version", True)]),
                        ("様式の版が違う組（参考：位置のずれも見落としに数える）", [r for r in used if not r.get("same_version", True)])):
        added = sum(r.get("added", 0) for r in group)
        covered = sum(r.get("covered", 0) for r in group)
        agent = sum(r.get("covered_agent", 0) for r in group)
        print(f"{name} {len(group)} 組：記入例で書き込まれた所 {added} か所のうち、規則で見つけた入力欄の中 {covered - agent} か所"
              f"（{(covered - agent) / max(1, added):.1%}）、規則では決められない所（エージェントの判断）の中 {agent} か所"
              f"（合わせて {covered / max(1, added):.1%}）")
    kinds = sorted({r["kind"] for r in used})
    for k in kinds:
        g = [r for r in used if r["kind"] == k and r.get("same_version", True)]
        added = sum(r.get("added", 0) for r in g)
        covered = sum(r.get("covered", 0) for r in g)
        agent = sum(r.get("covered_agent", 0) for r in g)
        if added:
            print(f"  {k}（版が同じ {len(g)} 組）：規則 {(covered - agent) / added:.1%}・規則＋エージェント {covered / added:.1%}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
