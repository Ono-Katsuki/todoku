"""世界調査の結果から、ポータルのデータ（world/portal/data.json）を作る。

入力：world/population.json（母集団と国ごとの結果）・world/domains.json（政府のドメイン）・world/frames.json（枠）・
world/sample.jsonl（標本）・world/results.jsonl（測った結果）。結果は、今の標本と id・digest が同じものだけを使う。

- 層（国 × 形式の群）の中の割合は、選ばれる確率の逆数で重みを付けた比の推定。分散は、索引のブロック（集落）ごとに
  まとめた線形化の推定（同じブロックの書類は似ているため）。
- 国の値は、層の大きさ（枠の推定の数）で重みを付けて合わせる（層別の推定。分散は層ごとの和）。区間は、分散から出した
  有効な標本の大きさ（デザインの自由度で直したもの）に当てた Clopper-Pearson の区間（Korn & Graubard 1998）。
  計算は scripts/eval/common.py の stratified_estimate（最終の評価の数と同じ関数）。
使い方: PYTHONPATH=src .venv/bin/python scripts/world_portal.py
"""

from __future__ import annotations

import json
import sys
import urllib.parse
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(Path(__file__).resolve().parent / "eval"))
from common import stratified_estimate  # noqa: E402

W = ROOT / "world"
FIRST_PASS = 20  # 1回目に測る標本の順番（world_survey.py の --rank）

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


def ratio(rows: list[tuple[float, str, int]]) -> tuple[float | None, float, int, int]:
    """rows：(重み, 集落, 当たり 0/1)。重み付きの割合・集落でまとめた線形化の分散・件数・集落の数。"""
    if not rows:
        return None, 0.0, 0, 0
    sw = sum(w for w, _c, _y in rows)
    p = sum(w * y for w, _c, y in rows) / sw
    by = defaultdict(float)
    for w, c, y in rows:
        by[c] += w * (y - p)
    k = len(by)
    var = (k / (k - 1)) * sum(v * v for v in by.values()) / (sw * sw) if k > 1 else 0.0
    return p, var, len(rows), k


def combine(cells: list[tuple[float, float | None, float, int, int]]) -> dict:
    """cells：(層の大きさ N, 割合 p, 分散 v, 件数 n, 集落の数 k)。層別の推定と Korn & Graubard の区間（Clopper-Pearson）。"""
    e = stratified_estimate(cells)
    if e is None:
        return {"p": None, "ci": [None, None], "n": 0}
    return {"p": round(e["p"], 4), "ci": [round(x, 4) for x in e["ci"]], "n": e["n"]}


def _ungraded_why(r: dict) -> str:
    e, k = r.get("error") or "", r.get("kind") or ""
    if e == "unknown_kind":
        return {"zip": "ZIP（Office の書類でない）"}.get(k, "対応していない形式")
    return {"password": "パスワード", "encrypted": "暗号化", "timeout": "時間切れ", "crash": "処理の失敗", "memory": "メモリの上限"}.get(e, e or "その他")


def _source_line(d: dict) -> str:
    s = [x["suffix"] for x in d["suffixes"] if "excluded" not in x]
    parts = []
    if s:
        parts.append("・".join(s[:6]) + (f" ほか {len(s) - 6}" if len(s) > 6 else ""))
    if d["hosts"]:
        parts.append(f"Wikidata の政府の機関・自治体のサイト {len(d['hosts']):,}")
    return "政府のドメイン：" + "＋".join(parts) if parts else ""


def main() -> int:
    frames = json.loads((W / "frames.json").read_text(encoding="utf-8"))
    crawl = frames["crawl"]
    parts = {"suffix": frames.get("countries", {}), "wikidata": frames.get("wikidata", {})}
    doms = json.loads((W / "domains.json").read_text(encoding="utf-8"))["countries"]
    samples = {s["id"]: s for s in (json.loads(x) for x in (W / "sample.jsonl").read_text(encoding="utf-8").splitlines() if x.strip())}
    recs = []
    if (W / "results.jsonl").exists():
        for x in (W / "results.jsonl").read_text(encoding="utf-8").splitlines():
            try:
                r = json.loads(x)
            except json.JSONDecodeError:
                continue
            s = samples.get(r.get("id"))
            if s and s["digest"] == r.get("digest"):
                recs.append(r)
    by_cell = defaultdict(list)
    for r in recs:
        by_cell[(r["country"], r["family"], r.get("domain_source", "suffix"))].append(r)
    pop = {c["iso2"]: c for c in POP}

    countries, unreached = {}, {}
    for c in POP:
        cc = c["iso2"]
        # 層：国 × 群 × ドメインの出典（重ならない）
        strata = [(fam, src, ff) for src, part in parts.items() for fam, ff in ((part.get(cc) or {}).get("families") or {}).items()]
        if not any(ff["pool"] for _f, _s, ff in strata):
            unreached[cc] = {"name": c["ja"], "why": c.get("reason") or "", "outcome": c["outcome"], "iso": ISO_NUM.get(cc)}
            continue
        fams = {}
        cells = {k: [] for k in ("reach", "read_ok", "write_ok", "forms_writable")}
        for fam, src, ff in strata:
            if not ff["pool"]:
                continue
            N = ff["frame"] or ff["pool"]
            rs = by_cell.get((cc, fam, src), [])
            got = [r for r in rs if r.get("reach") == "ok"]
            graded = [r for r in got if r.get("read")]
            forms = [r for r in graded if form_class(r) == "marked"]
            wt = lambda r: 1.0 / r["pi"]  # noqa: E731
            est = {
                "reach": ratio([(wt(r), str(r["block"]), int(r.get("reach") == "ok")) for r in rs]),
                "read_ok": ratio([(wt(r), str(r["block"]), int(r["read"] in READ_OK)) for r in graded]),
                "write_ok": ratio([(wt(r), str(r["block"]), int(r["write"] in ("A", "B"))) for r in graded]),
                "forms_writable": ratio([(wt(r), str(r["block"]), int(r["write"] in "AB")) for r in forms]),
            }
            for k, tup in est.items():
                cells[k].append((N, *tup))
            g = fams.setdefault(fam, {"frame": 0, "exact": True, "pool": 0, "planned": 0, "sampled": 0, "reached": 0,
                                      "graded": 0, "forms": 0, "live": 0, "truncated": 0, "read": Counter(), "write": Counter(),
                                      "outcomes": Counter()})
            g["frame"] += N
            g["exact"] = g["exact"] and ff["exact"]
            g["pool"] += ff["pool"]
            g["planned"] += min(FIRST_PASS, ff["pool"])
            g["sampled"] += len(rs)
            g["reached"] += len(got)
            g["graded"] += len(graded)
            g["forms"] += len(forms)
            g["live"] += sum(1 for r in got if r.get("source") == "live")
            g["truncated"] += sum(1 for r in rs if r.get("reach") == "truncated")
            g["read"].update(r["read"] for r in graded)
            g["write"].update(r["write"] for r in graded)
            g["outcomes"].update(r.get("reach") for r in rs)
        mine = [r for r in recs if r["country"] == cc]
        countries[cc] = {"name": c["ja"], "iso": ISO_NUM.get(cc), "source": _source_line(doms[cc]), "families": fams,
                         **{k: combine(v) for k, v in cells.items()},
                         "read": dict(Counter(r["read"] for r in mine if r.get("read"))),
                         "write": dict(Counter(r["write"] for r in mine if r.get("write")))}

    docs = []
    for r in recs:
        u = r.get("url") or ""
        sp = urllib.parse.urlsplit(u)
        name = urllib.parse.unquote(sp.path.rsplit("/", 1)[-1]) or sp.netloc
        docs.append({
            "c": r["country"], "g": r["family"], "k": r.get("kind") or "", "d": (r.get("mime_detected") or "").rsplit("/", 1)[-1][:40],
            "o": r.get("reach"), "r": r.get("read") or "", "w": r.get("write") or "",
            "rr": [t for _g, t in (r.get("read_reasons") or [])][:4], "wr": [t for _g, t in (r.get("write_reasons") or [])][:3],
            "u": u, "t": name[:140], "p": sp.netloc[:80], "th": [], "f": r.get("fields") or 0,
            "fc": form_class(r) if r.get("read") else "", "e": r.get("error") or "", "s": r.get("source") or "",
        })
    reasons = Counter(t for r in recs for _g, t in (r.get("read_reasons") or []))
    wreasons = Counter(t for r in recs for _g, t in (r.get("write_reasons") or []))
    total = {
        "crawl": crawl,
        "sampled": len(recs),
        "reached": sum(1 for r in recs if r.get("reach") == "ok"),
        "graded": sum(1 for r in recs if r.get("read")),
        "live": sum(1 for r in recs if r.get("source") == "live"),
        "truncated": sum(1 for r in recs if r.get("cc_truncated")),
        # 届いたが段階を付けられなかった書類（対応していない形式・パスワード・時間切れなど）：分母から除くので、数と理由を出す
        "ungraded": dict(Counter(_ungraded_why(r) for r in recs if r.get("reach") == "ok" and not r.get("read"))),
        "regrade_kept": sum(1 for r in recs if r.get("regrade")),
        "planned": sum(min(FIRST_PASS, v["pool"]) for part in parts.values() for x in part.values()
                       for v in (x.get("families") or {}).values()),
        "countries": len(countries),
        "measured_countries": sum(1 for c in countries.values() if c["read_ok"]["n"]),
        "outcomes": dict(Counter(r.get("reach") for r in recs)),
        "read": dict(Counter(r["read"] for r in recs if r.get("read"))),
        "write": dict(Counter(r["write"] for r in recs if r.get("write"))),
        "top_read_reasons": reasons.most_common(12), "top_write_reasons": wreasons.most_common(8),
        "form_classes": dict(Counter(form_class(r) for r in recs if r.get("read"))),
        "audited": sum(1 for r in recs if "audit_ok" in r), "audit_ok": sum(1 for r in recs if r.get("audit_ok")),
        "kinds": dict(Counter(r.get("kind") for r in recs if r.get("read"))),
    }
    out = {"total": total, "countries": countries, "unreached": unreached,
           "population": {"size": sum(1 for c in POP if c["un_member"]),
                          "outcomes": dict(Counter(pop[cc]["outcome"] for cc in pop if pop[cc]["un_member"]))},
           "docs": docs}
    (W / "portal").mkdir(exist_ok=True)
    (W / "portal" / "data.json").write_text(json.dumps(out, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"書類 {len(docs):,} 件・枠のある国 {len(countries)}・測った国 {total['measured_countries']}・段階を付けた {total['graded']:,} 件")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
