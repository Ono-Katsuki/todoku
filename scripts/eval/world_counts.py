#!/usr/bin/env python3
"""世界地図の数：国ごとと世界全体で、とどくで届く書類の推定の数（枠の大きさ × 割合）と区間。読むと書くは別に。

- 届く：調査の標本を凍結した版で付け直した段階（読む S・A・B／書く A・B）。層（国 × 形式の群 × ドメインの出典）の中は選ばれる確率の逆数の重みの比の推定、
  分散はブロックでまとめた線形化（scripts/world_portal.py と同じ）。国は層の枠の大きさで重み付け。区間は Korn & Graubard (1998)：有効な標本の大きさ n* = p(1-p)/var
  をデザインの自由度で直し（n を超えない）、Clopper-Pearson を当てる（common.py の stratified_estimate。portal と同じ関数）。
- 新しく届く：標本の無作為な一部で、比べる道具ととどくを同じ定義（read_compare.py の exact_read・write_compare.py の write_ok）で動かす。
  new＝とどくは通り、どの比べる道具も通らない書類の割合。lost＝比べる道具のどれかは通り、とどくは通らない割合。正味 ＝ new − lost。枠 × 割合で数にする。
- 数は Common Crawl が集めた範囲の中の推定で、政府の書類の全部ではない（下限）。枠の推定（exact でない層）の誤差は、区間に入っていない。

使い方:
  world_counts.py regrade  --results 写し [--graders 1]                     調査の全部を付け直す（時間がかかる。続きから動く）
  world_counts.py pick     --results 写し --out D [--per-stratum 2] [--seed S]   一部を選ぶ（結果を見ない：id と層だけで決める）
  world_counts.py compare  --out D --baseline-python PY --todoku-python PY [--pkr P]  選んだ書類を Common Crawl から取り直し、比べて、書類を消す
  world_counts.py mock     --out D --read-scores DIR --write-scores DIR       （試運転だけ）計算の確かめ：開発の書類の結果を、選んだ id に無作為に当てる
  world_counts.py estimate --results 写し --out D [--label L]                 数を出す"""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import random
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent))
from common import ROOT, atomic_json, doc_id, load_json, md_table, stratified_estimate  # noqa: E402

W = ROOT / "world"
READ_OK, WRITE_OK = set("SAB"), set("AB")
SEED = 20261105


def read_results(path) -> list[dict]:
    out = []
    for ln in Path(path).read_text(encoding="utf-8").splitlines():
        try:
            out.append(json.loads(ln))
        except json.JSONDecodeError:
            pass
    return out


def stratum(r: dict) -> tuple:
    return (r["country"], r["family"], r.get("domain_source", "suffix"))


def cmd_regrade(a) -> int:
    cmd = [sys.executable, str(HERE.parent / "world_survey.py"), "--regrade", "--jobs", "2", "--graders", str(a.graders), "--results", a.results]
    return subprocess.run(cmd, cwd=ROOT, env={**__import__("os").environ, "PYTHONPATH": str(ROOT / "src")}).returncode


def cmd_pick(a) -> int:
    """届いて段階の付いた書類から、層ごとに per_stratum 件。並びは種と id のハッシュで決める（結果の中身を見ない）。"""
    rows = [r for r in read_results(a.results) if r.get("reach") == "ok" and r.get("id")]
    by = collections.defaultdict(list)
    for r in rows:
        by[stratum(r)].append(r)
    pick = []
    for key in sorted(by):
        rs = sorted(by[key], key=lambda r: hashlib.sha256(f"{a.seed}:{r['id']}".encode()).hexdigest())
        pick += rs[: a.per_stratum]
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    keep = ("id", "country", "family", "domain_source", "pi", "block", "digest", "sha256", "url", "filename", "offset", "length", "mime", "mime_detected", "content_type", "crawl", "timestamp", "rank",
            "block_stratum", "warc_date", "kind")
    atomic_json(out / "subset.json", {"seed": a.seed, "per_stratum": a.per_stratum, "strata": len(by), "docs": [{k: r.get(k) for k in keep} for r in pick]})
    print(f"層 {len(by)}・選んだ書類 {len(pick)}")
    return 0


def cmd_compare(a) -> int:
    import world_survey as ws

    out = Path(a.out)
    sub = load_json(out / "subset.json")
    docs_dir = out / "docs"
    docs_dir.mkdir(exist_ok=True)
    mapping = load_json(out / "docs_map.json") or {}
    # 書類を取り直す。人に断られたら（403 など）、その書類は記録して飛ばす。回り道をしない
    for s in sub["docs"]:
        if s["id"] in mapping:
            continue
        rec, body = ws.obtain({k: v for k, v in s.items() if k != "kind"})
        if body is None:
            mapping[s["id"]] = {"skip": rec.get("reach") or "取れない"}
            atomic_json(out / "docs_map.json", mapping)
            continue
        if s.get("sha256") and hashlib.sha256(body).hexdigest() != s["sha256"]:
            mapping[s["id"]] = {"skip": "中身が変わった"}
            atomic_json(out / "docs_map.json", mapping)
            continue
        suffix = ws.suffix_for({**s, **rec})
        p = docs_dir / f"{s['id']}{suffix}"
        p.write_bytes(body)
        mapping[s["id"]] = {"path": str(p)}
        atomic_json(out / "docs_map.json", mapping)
    files = [m["path"] for m in mapping.values() if m.get("path")]
    lst = out / "files.txt"
    lst.write_text("\n".join(files) + "\n", encoding="utf-8")
    base = [a.baseline_python, a.todoku_python]
    extra = ["--pkr", a.pkr] if a.pkr else []
    r1 = subprocess.run([sys.executable, str(HERE / "read_compare.py"), "--files", str(lst), "--out", str(out / "read"), "--baseline-python", base[0], "--todoku-python", base[1],
                         "--no-excel", "--label", "調査の一部", "--purge", *extra])
    r2 = subprocess.run([sys.executable, str(HERE / "write_compare.py"), "--files", str(lst), "--out", str(out / "write"), "--baseline-python", base[0], "--todoku-python", base[1],
                         "--no-signed", "--label", "調査の一部", "--purge", *extra])
    collect(out, mapping)
    shutil.rmtree(docs_dir, ignore_errors=True)  # 書類は残さない（調査の決まり）
    return r1.returncode or r2.returncode


def collect(out: Path, mapping: dict) -> None:
    scores = {}
    for i, m in mapping.items():
        if not m.get("path"):
            continue
        rd = load_json(out / "read" / "scores" / f"{doc_id(m['path'])}.json") or {}
        wr = load_json(out / "write" / "scores" / f"{doc_id(m['path'])}.json") or {}
        scores[i] = one_score(rd, wr)
    atomic_json(out / "scores.json", {"mock": False, "scores": scores})


def one_score(rd: dict, wr: dict) -> dict:
    t = rd.get("tools") or {}
    d = {"read_todoku": bool((t.get("todoku") or {}).get("exact_read")), "read_base": any(s.get("exact_read") for k, s in t.items() if k != "todoku")}
    d["read_ok"] = bool(t)
    runs = wr.get("runs") or {}
    if runs:  # 書く比べは英数字の値で数える（注釈でも書ける形：比べる道具に有利）
        d["write_todoku"] = bool((runs.get("todoku:ascii") or {}).get("write_ok"))
        d["write_base"] = any(r.get("write_ok") for k, r in runs.items() if k.endswith(":ascii") and not k.startswith("todoku"))
    return d


def cmd_mock(a) -> int:
    """試運転だけ：計算の確かめ。開発の書類の結果を、選んだ id に無作為に当てる。ここから出る「新しく届く」の数に意味はない。"""
    out = Path(a.out)
    sub = load_json(out / "subset.json")
    rd = [load_json(p) for p in sorted(Path(a.read_scores).glob("*.json"))]
    wr = {r["path"]: r for r in (load_json(p) for p in sorted(Path(a.write_scores).glob("*.json")))}
    rng = random.Random(a.seed)
    scores = {}
    for s in sub["docs"]:
        r = rng.choice(rd)
        scores[s["id"]] = one_score(r, wr.get(r["path"], {}))
    atomic_json(out / "scores.json", {"mock": True, "scores": scores})
    return 0


def est(cells: list[tuple]) -> dict:
    """cells：(枠 N, 割合 p, 分散 v, 件数 n, 集落の数 k)。件数のある層だけで重みを付ける（portal と同じ関数）。"""
    e = stratified_estimate(cells)
    if e is None:
        return {"count": None, "ci": [None, None], "p": None, "n": 0, "frame": 0}
    tot, p, (lo, hi) = e["frame"], e["p"], e["ci"]
    return {"count": round(tot * p), "ci": [round(tot * lo), round(tot * hi)], "p": p, "p_ci": e["ci"], "n": e["n"], "n_eff": e["n_eff"], "frame": tot}


def cmd_estimate(a) -> int:
    import world_portal as wp

    out = Path(a.out)
    frames = json.loads((W / "frames.json").read_text(encoding="utf-8"))
    parts = {"suffix": frames.get("countries", {}), "wikidata": frames.get("wikidata", {})}
    samples = {s["id"]: s for s in (json.loads(x) for x in (W / "sample.jsonl").read_text(encoding="utf-8").splitlines() if x.strip())}
    recs = [r for r in read_results(a.results) if samples.get(r.get("id"), {}).get("digest") == r.get("digest")]
    by_cell = collections.defaultdict(list)
    for r in recs:
        by_cell[stratum(r)].append(r)
    sc_file = load_json(out / "scores.json") or {"mock": None, "scores": {}}
    scores = sc_file["scores"]
    by_id = {r["id"]: r for r in recs}
    sub_cell = collections.defaultdict(list)
    for i, s in scores.items():
        if i in by_id:
            sub_cell[stratum(by_id[i])].append((by_id[i], s))
    pop = {c["iso2"]: c for c in wp.POP}
    keys = ("read", "write", "todoku_read", "base_read", "new_read", "lost_read", "todoku_write", "base_write", "new_write", "lost_write")
    world_cells = {k: [] for k in keys}
    countries = {}
    frame_world, frame_unmeasured = 0, 0
    exact_all = True
    for cc in sorted({cc for part in parts.values() for cc in part}):
        cells = {k: [] for k in keys}
        frame = 0
        for src, part in parts.items():
            for fam, ff in ((part.get(cc) or {}).get("families") or {}).items():
                if not ff.get("pool"):
                    continue
                N = ff.get("frame") or ff["pool"]
                frame += N
                exact_all = exact_all and bool(ff.get("exact"))
                graded = [r for r in by_cell.get((cc, fam, src), []) if r.get("reach") == "ok" and r.get("read")]
                wt = lambda r: 1.0 / r["pi"]  # noqa: E731
                est_map = {"read": wp.ratio([(wt(r), str(r["block"]), int(r["read"] in READ_OK)) for r in graded]),
                           "write": wp.ratio([(wt(r), str(r["block"]), int(r["write"] in WRITE_OK)) for r in graded if r.get("write")])}
                ss = sub_cell.get((cc, fam, src), [])
                rd = [(wt(r), str(r["block"]), s) for r, s in ss if s.get("read_ok")]
                wr = [(wt(r), str(r["block"]), s) for r, s in ss if "write_todoku" in s]
                est_map.update({"todoku_read": wp.ratio([(w, b, int(s["read_todoku"])) for w, b, s in rd]),
                                "base_read": wp.ratio([(w, b, int(s["read_base"])) for w, b, s in rd]),
                                "new_read": wp.ratio([(w, b, int(s["read_todoku"] and not s["read_base"])) for w, b, s in rd]),
                                "lost_read": wp.ratio([(w, b, int(s["read_base"] and not s["read_todoku"])) for w, b, s in rd]),
                                "todoku_write": wp.ratio([(w, b, int(s["write_todoku"])) for w, b, s in wr]),
                                "base_write": wp.ratio([(w, b, int(s["write_base"])) for w, b, s in wr]),
                                "new_write": wp.ratio([(w, b, int(s["write_todoku"] and not s["write_base"])) for w, b, s in wr]),
                                "lost_write": wp.ratio([(w, b, int(s["write_base"] and not s["write_todoku"])) for w, b, s in wr])})
                for k, tup in est_map.items():
                    cells[k].append((N, *tup))
                    world_cells[k].append((N, *tup))
        if not frame:
            continue
        frame_world += frame
        c = {k: est(v) for k, v in cells.items()}
        if not c["read"]["n"]:
            frame_unmeasured += frame
        for x, new, lost in (("read", "new_read", "lost_read"), ("write", "new_write", "lost_write")):
            if c[new]["count"] is not None:
                c[f"net_{x}"] = c[new]["count"] - c[lost]["count"]
        countries[cc] = {"name": (pop.get(cc) or {}).get("ja", cc), "iso": (pop.get(cc) or {}).get("num"), "frame": frame, **c}
    world = {k: est(v) for k, v in world_cells.items()}
    for x, new, lost in (("read", "new_read", "lost_read"), ("write", "new_write", "lost_write")):
        if world[new]["count"] is not None:
            world[f"net_{x}"] = world[new]["count"] - world[lost]["count"]
    res = {"label": a.label, "crawl": frames.get("crawl"), "mock_subset": sc_file.get("mock"), "subset_docs": len(scores), "frame_total": frame_world, "frame_unmeasured": frame_unmeasured,
           "all_frames_exact": exact_all, "results_used": len(recs), "world": world, "countries": countries,
           "note": "Common Crawl が集めた範囲の中の推定（政府の書類の全部ではない下限）。枠の推定の誤差は区間に入っていない。"}
    atomic_json(out / "counts.json", res)
    (out / "counts.md").write_text(markdown(res), encoding="utf-8")
    print(markdown(res))
    return 0


def fmt(e: dict | None) -> str:
    if not e or e.get("count") is None:
        return "—"
    return f"{e['count']:,}（{e['ci'][0]:,}〜{e['ci'][1]:,}）"


def markdown(r: dict) -> str:
    w = r["world"]
    L = [f"# 世界地図の数（{r['label']}）", "", f"枠の合計 {r['frame_total']:,} 件（{r['crawl']}）。使った調査の記録 {r['results_used']:,} 件。比べた一部 {r['subset_docs']:,} 件。", "", r["note"], ""]
    if r["mock_subset"]:
        L += ["**「新しく届く」の列は計算の確かめだけ：開発の書類の結果を無作為に当てたもので、数に意味はない。**", ""]
    L += [md_table(["世界全体", "読む", "書く"],
                   [["とどくで届く（段階）", fmt(w["read"]), fmt(w["write"])],
                    ["同じ定義：とどくが通る", fmt(w["todoku_read"]), fmt(w["todoku_write"])],
                    ["同じ定義：比べる道具のどれかが通る", fmt(w["base_read"]), fmt(w["base_write"])],
                    ["新しく届く（とどくだけ通る）", fmt(w["new_read"]), fmt(w["new_write"])],
                    ["とどくだけ通らない", fmt(w["lost_read"]), fmt(w["lost_write"])],
                    ["正味", w.get("net_read", "—"), w.get("net_write", "—")]]), ""]
    top = sorted(r["countries"].items(), key=lambda x: -(x[1]["read"]["count"] or 0))[:25]
    L += ["## 読める書類の多い国（上位 25）", "",
          md_table(["国", "枠", "読む（件・区間）", "書く（件・区間）", "新しく読める", "新しく書ける"],
                   [[c["name"], f"{c['frame']:,}", fmt(c["read"]), fmt(c["write"]), fmt(c["new_read"]), fmt(c["new_write"])] for _k, c in top]), "",
          "全部の国は counts.json（portal の地図の数はここから取れる）。"]
    return "\n".join(L)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["regrade", "pick", "compare", "mock", "estimate"])
    ap.add_argument("--results")
    ap.add_argument("--out", default="")
    ap.add_argument("--graders", type=int, default=1)
    ap.add_argument("--per-stratum", type=int, default=2)
    ap.add_argument("--seed", type=int, default=SEED)
    ap.add_argument("--baseline-python")
    ap.add_argument("--todoku-python")
    ap.add_argument("--pkr", default=None)
    ap.add_argument("--read-scores")
    ap.add_argument("--write-scores")
    ap.add_argument("--label", default="")
    a = ap.parse_args()
    return {"regrade": cmd_regrade, "pick": cmd_pick, "compare": cmd_compare, "mock": cmd_mock, "estimate": cmd_estimate}[a.cmd](a)


if __name__ == "__main__":
    raise SystemExit(main())
