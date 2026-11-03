"""世界調査の結果（world/results.jsonl）から、ポータルのデータ（world/portal/data.json）を作る。

- 国ごと・層（国 × 形式の群）ごとに、届いた割合・読む段階・書く段階の分布を数え、Wilson の 95% 区間を付ける。
- 国の値は、層の大きさ（目録の登録の数）で重みを付けて合わせる（層別の推定。分散は層ごとの和）。
- 書類ごとの記録は、ページで絞り込めるように短い形で入れる（URL・形式・段階・理由・発行者の分類）。
使い方: uv run python scripts/world_portal.py
"""

from __future__ import annotations

import json
import math
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
W = ROOT / "world"
Z = 1.959964

# 国の番号（地図の ISO 3166-1 の数字）は、母集団の記録（world/population.json）から取る
POP = json.loads((W / "population.json").read_text(encoding="utf-8"))["countries"]
ISO_NUM = {c["iso2"]: c["num"] for c in POP}
READ_OK = set("SAB")  # 正確に読める（素のデータ・規則だけで・エージェントが読めば）
# 様式の見分け（構造の印だけで決める）：作り手が入力欄の印を付けた書類（PDF・XFA のフォームの欄、Word のコンテンツコントロール、
# Excel の入力の印・フォームの部品など）。配置や空白から推定した欄だけの書類は「推定の欄」として別に数える
MARKED = {"XFA のフォームの欄", "フォームの欄", "印のある入力欄", "作り手の印のある入力欄", "データに結びつかない XFA の欄",
          "状態の読めないフォームの欄"}
# 推定の欄：いまの grade.py の書く段階の理由の名前（配置・枠や塗り・文字の中の空欄・数式の入力から推定した欄、規則では決められない所
# （窓の「？」）、重なった入力規則）。古い記録の名前（「解析して決める入力欄」は上の推定の名前に分かれ、「入力欄か決められないセル」は
# 「規則では決められない所」になった）も、取り直す前の記録を読むために残す
INFERRED = {"配置から推定した入力欄", "枠・塗りから推定した入力欄", "文字の中の空欄", "印のない入力欄", "規則では決められない所",
            "重なった入力規則"}
INFERRED_OLD = {"解析して決める入力欄", "入力欄か決められないセル"}


def form_class(r: dict) -> str:
    titles = {t for _g, t in (r.get("read_reasons") or []) + (r.get("write_reasons") or [])}
    if titles & MARKED:
        return "marked"
    if titles & (INFERRED | INFERRED_OLD) or (r.get("fields") or 0) > 0:
        return "inferred"
    return "none"


def wilson(k: int, n: int) -> list:
    if n == 0:
        return [None, None]
    p = k / n
    d = 1 + Z * Z / n
    c = (p + Z * Z / (2 * n)) / d
    h = Z * math.sqrt(p * (1 - p) / n + Z * Z / (4 * n * n)) / d
    return [round(max(0, c - h), 4), round(min(1, c + h), 4)]


def stratified(cells: list[tuple[int, int, int]]) -> dict:
    """cells：(層の大きさ N, 標本 n, 当たり k)。層別の推定と、正規近似の 95% 区間。"""
    cells = [c for c in cells if c[1] > 0]
    tot = sum(c[0] for c in cells)
    if not tot:
        return {"p": None, "ci": [None, None], "n": 0}
    p = sum(N / tot * k / n for N, n, k in cells)
    var = sum((N / tot) ** 2 * (k / n) * (1 - k / n) / max(n - 1, 1) for N, n, k in cells)
    # 区間は、層別の分散から有効な標本の大きさを出して、その大きさのウィルソン区間にする（Korn & Graubard 1998）。
    # 正規近似のままだと、どの層も 0% か 100% のとき分散が 0 になり、区間の幅が 0 になる
    n_all = sum(c[1] for c in cells)
    n_eff = p * (1 - p) / var if var > 0 else n_all
    n_eff = min(n_eff, n_all)
    d = 1 + Z * Z / n_eff
    c = (p + Z * Z / (2 * n_eff)) / d
    h = Z * math.sqrt(p * (1 - p) / n_eff + Z * Z / (4 * n_eff * n_eff)) / d
    return {"p": round(p, 4), "ci": [round(max(0, c - h), 4), round(min(1, c + h), 4)], "n": n_all}


def _ungraded_why(r: dict) -> str:
    e, k = r.get("error") or "", r.get("kind") or ""
    if e == "unknown_kind":
        return {"zip": "ZIP（Office の書類でない）"}.get(k, "対応していない形式")
    return {"password": "パスワード", "encrypted": "暗号化", "timeout": "時間切れ", "crash": "処理の失敗"}.get(e, e or "その他")


def main() -> int:
    frames = json.loads((W / "frames.json").read_text(encoding="utf-8"))
    # 枠のない国：目録はあるが使えない（理由）・公式の国の目録が見つからない。母集団の全部の国を地図に載せる
    unreached = {c["iso2"]: (c["ja"], c["reason"] or "枠をまだ作っていない", c["outcome"]) for c in POP if c["iso2"] not in frames}
    recs = [json.loads(x) for x in (W / "results.jsonl").read_text(encoding="utf-8").splitlines() if x.strip()]
    by_cell = defaultdict(list)
    for r in recs:
        by_cell[(r["country"], r["family"])].append(r)

    countries = {}
    for cc, f in frames.items():
        fams = {}
        reach_c, read_c, write_c, form_write_c = [], [], [], []
        for fam, names in f["strata"].items():
            N = sum(c for _, c in names)
            rs = by_cell.get((cc, fam), [])
            got = [r for r in rs if r.get("reach") == "ok"]
            graded = [r for r in got if r.get("read")]
            forms = [r for r in graded if form_class(r) == "marked"]
            n_ok = sum(1 for r in graded if r["read"] in READ_OK)
            n_wa = sum(1 for r in graded if r["write"] == "A")
            n_fa = sum(1 for r in forms if r["write"] in "AB")
            fams[fam] = {
                "frame": N, "sampled": len(rs), "reached": len(got), "graded": len(graded), "forms": len(forms),
                "reach_ci": wilson(len(got), len(rs)),
                "read": dict(Counter(r["read"] for r in graded)), "write": dict(Counter(r["write"] for r in graded)),
                "read_ok_ci": wilson(n_ok, len(graded)), "write_a_ci": wilson(n_wa, len(graded)),
                "outcomes": dict(Counter(r.get("reach") for r in rs)),
            }
            reach_c.append((N, len(rs), len(got)))
            read_c.append((N, len(graded), n_ok))
            write_c.append((N, len(graded), n_wa))
            form_write_c.append((N, len(forms), n_fa))
        countries[cc] = {"name": f["name"], "iso": ISO_NUM.get(cc), "source": f["source"], "families": fams,
                         "reach": stratified(reach_c), "read_ok": stratified(read_c), "write_a": stratified(write_c),
                         "forms_writable": stratified(form_write_c),
                         "read": dict(Counter(r["read"] for r in recs if r["country"] == cc and r.get("read"))),
                         "write": dict(Counter(r["write"] for r in recs if r["country"] == cc and r.get("write")))}

    docs = []
    for r in recs:
        if r.get("reach") == "frame_empty":
            continue
        docs.append({
            "c": r["country"], "g": r["family"], "k": r.get("kind") or "", "d": r.get("declared") or "",
            "o": r.get("reach"), "r": r.get("read") or "", "w": r.get("write") or "",
            "rr": [t for _g, t in (r.get("read_reasons") or [])][:4], "wr": [t for _g, t in (r.get("write_reasons") or [])][:3],
            "u": r.get("url") or "", "t": str(r.get("dataset") or "")[:140], "p": str(r.get("publisher") or "")[:80],
            "th": (r.get("themes") or [])[:4], "f": r.get("fields") or 0, "fc": form_class(r) if r.get("read") else "", "e": r.get("error") or "",
        })
    reasons = Counter(t for r in recs for _g, t in (r.get("read_reasons") or []))
    wreasons = Counter(t for r in recs for _g, t in (r.get("write_reasons") or []))
    total = {
        "sampled": sum(1 for r in recs if r.get("reach") != "frame_empty"),
        "reached": sum(1 for r in recs if r.get("reach") == "ok"),
        "graded": sum(1 for r in recs if r.get("read")),
        # 届いたが段階を付けられなかった書類（対応していない形式・パスワード・時間切れなど）：分母から除くので、数と理由を出す
        "ungraded": dict(Counter(_ungraded_why(r) for r in recs if r.get("reach") == "ok" and not r.get("read"))),
        "regrade_kept": sum(1 for r in recs if r.get("regrade")),
        "planned": sum(min(250, sum(c for _, c in v)) for x in frames.values() for v in x["strata"].values()),
        "countries": len(frames),
        "outcomes": dict(Counter(r.get("reach") for r in recs)),
        "read": dict(Counter(r["read"] for r in recs if r.get("read"))),
        "write": dict(Counter(r["write"] for r in recs if r.get("write"))),
        "top_read_reasons": reasons.most_common(12), "top_write_reasons": wreasons.most_common(8),
        "form_classes": dict(Counter(form_class(r) for r in recs if r.get("read"))),
        "audited": sum(1 for r in recs if "audit_ok" in r), "audit_ok": sum(1 for r in recs if r.get("audit_ok")),
        "kinds": dict(Counter(r.get("kind") for r in recs if r.get("read"))),
    }
    out = {"total": total, "countries": countries,
           "unreached": {cc: {"name": v[0], "why": v[1], "outcome": v[2], "iso": ISO_NUM.get(cc)} for cc, v in unreached.items()},
           "population": {"size": sum(1 for c in POP if c["un_member"]),
                          "outcomes": dict(Counter(("reachable" if c["iso2"] in frames else c["outcome"]) for c in POP if c["un_member"]))},
           "docs": docs}
    (W / "portal").mkdir(exist_ok=True)
    (W / "portal" / "data.json").write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"書類 {len(docs):,} 件・国 {len(countries)}・段階を付けた {total['graded']:,} 件")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
