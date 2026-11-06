#!/usr/bin/env python3
"""第 2 段の新しい標本：結果を見る前に決める（docs/再評価の計画.md）。

plan   ：world/frames.json（枠の大きさ）だけから、層ごとの件数とブロックの数を決め、stage2_plan.json に書く（書類の中身も結果も見ない）。リポジトリに入れる。
select ：計画のとおりに Common Crawl の索引を読み、候補を作り、種と digest のハッシュの順に選ぶ（第 2 段の実行のときに動かす。結果を見ない）。

枠と層は公開した方法と同じ（CC-MAIN-2026-43・world/domains.json・国 × 群 × 出典）。違うのは：
- 種（SEED2）と、すでに見た書類を除くこと（exclusion_digests）。
- 層の件数：各層 K_MIN 件 + 残りを枠の大きさに比べて割り当て（v1 の「どの層も 2 件」は、枠の 24% の US の PDF にも 2 件だった）。
- ブロック：国 × 出典ごとに、新しい種で全ブロックから選ぶ。件数に応じて M_MIN〜M_MAX。足りない群があれば、公開した方法と同じく無作為な順に足す（合計 EXTRA_CAP まで）。
候補が POOL_MIN 件以下の層は標本 world/sample.jsonl に全部入っていて、新しい書類がない（公開した枠の 0.39%）。第 2 段では測らない。
π = (m_h / B_h) × (n / 新しい候補の数)。比の推定なので、n の違いは打ち消し合う。"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import math
import os
import random
import sys
from collections import Counter, defaultdict
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
W = ROOT / "world"
PLAN = HERE / "stage2_plan.json"
EXTRA_EXCL = HERE / "stage2_exclude_extra.txt"
CRAWL = "CC-MAIN-2026-43"
SEED2 = 20261107
N_TOTAL = 900
K_MIN = 2
POOL_MIN = 50
M_MIN, M_MAX = 4, 30
EXTRA_CAP = 120
BLOCK_BUDGET = 2600  # 索引のブロックの上限（時間の見積もりの根拠。超えたら足すのをやめ、足りない層は記録する）


def strata_of(frames: dict) -> list[dict]:
    out = []
    for src, part in (("suffix", frames.get("countries", {})), ("wikidata", frames.get("wikidata", {}))):
        for cc, p in part.items():
            for fam, ff in (p.get("families") or {}).items():
                if ff.get("pool"):
                    out.append({"country": cc, "family": fam, "domain_source": src, "pool": ff["pool"], "frame": ff.get("frame") or ff["pool"],
                                "blocks": sum((p.get("blocks") or {}).values())})
    return out


def allocate(strata: list[dict], n_total: int = N_TOTAL, k_min: int = K_MIN, pool_min: int = POOL_MIN) -> list[dict]:
    """各層 k_min 件 + 残りを枠の大きさに比べて（最大剰余法）。層の新しい候補（pool − pool_min）を超えない。件数は枠の大きさだけで決まる。"""
    use = sorted((dict(s) for s in strata if s["pool"] > pool_min), key=lambda s: (s["country"], s["family"], s["domain_source"]))
    for s in use:
        s["cap"] = s["pool"] - pool_min
        s["n"] = min(k_min, s["cap"])
    rest = n_total - sum(s["n"] for s in use)
    while rest > 0:
        tot = sum(s["frame"] for s in use if s["n"] < s["cap"])
        if not tot:
            break
        q = [(s["frame"] / tot * rest, s) for s in use if s["n"] < s["cap"]]
        give = 0
        for x, s in q:
            k = min(int(x), s["cap"] - s["n"])
            s["n"] += k
            give += k
        left = rest - give
        for x, s in sorted(q, key=lambda t: (-(t[0] - int(t[0])), t[1]["country"], t[1]["family"], t[1]["domain_source"]))[:left]:
            if s["n"] < s["cap"]:
                s["n"] += 1
                give += 1
        rest -= give
        if not give:
            break
    return use


def block_plan(strata: list[dict]) -> dict:
    """国 × 出典ごとの最初のブロックの数：その国 × 出典の件数と同じ（M_MIN〜M_MAX。全ブロックを超えない）。"""
    by = defaultdict(int)
    tot = {}
    for s in strata:
        k = f"{s['country']}:{s['domain_source']}"
        by[k] += s["n"]
        tot[k] = s["blocks"]
    return {k: {"docs": v, "blocks_total": tot[k], "first": min(tot[k], max(M_MIN, min(M_MAX, v)))} for k, v in sorted(by.items())}


def cmd_plan(a) -> int:
    frames = json.loads((W / "frames.json").read_text(encoding="utf-8"))
    st = strata_of(frames)
    use = allocate(st)
    frame_all = sum(s["frame"] for s in st)
    frame_use = sum(s["frame"] for s in use)
    bp = block_plan(use)
    doc = {"crawl": CRAWL, "seed": SEED2, "n_total": sum(s["n"] for s in use), "k_min": K_MIN, "pool_min": POOL_MIN, "m_min": M_MIN, "m_max": M_MAX, "extra_cap": EXTRA_CAP, "block_budget": BLOCK_BUDGET,
           "strata_all": len(st), "strata_used": len(use), "frame_all": frame_all, "frame_used": frame_use, "frame_share": round(frame_use / frame_all, 5),
           "countries_used": len({s["country"] for s in use}), "countries_all": len({s["country"] for s in st}),
           "first_blocks": sum(v["first"] for v in bp.values()),
           "allocation": [{k: s[k] for k in ("country", "family", "domain_source", "frame", "pool", "n")} for s in use], "blocks": bp}
    PLAN.write_text(json.dumps(doc, ensure_ascii=False, indent=0) + "\n", encoding="utf-8")
    print({k: v for k, v in doc.items() if k not in ("allocation", "blocks")})
    print("群ごとの件数", Counter({f: sum(s["n"] for s in use if s["family"] == f) for f in ("pdf", "sheet", "text", "csv")}))
    return 0


# ---------------------------------------------------------------------- すでに見た書類
def sha1_b32(b: bytes) -> str:
    return base64.b32encode(hashlib.sha1(b).digest()).decode()


def exclusion_digests(subset: str | None = None, extra_docs: list[str] | None = None) -> set:
    """見た書類の digest の集まり：world/sample.jsonl の全部（rank 50 まで）・world/validation/sample.json・固定版の一部（subset.json）・stage2_exclude_extra.txt。"""
    out = set()
    for ln in (W / "sample.jsonl").read_text(encoding="utf-8").splitlines():
        if ln.strip():
            out.add(json.loads(ln)["digest"])

    def walk(x):
        if isinstance(x, dict):
            if isinstance(x.get("digest"), str):
                out.add(x["digest"])
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)

    vs = W / "validation" / "sample.json"
    if vs.exists():
        walk(json.loads(vs.read_text(encoding="utf-8")))
    if subset and Path(subset).exists():
        walk(json.loads(Path(subset).read_text(encoding="utf-8")))
    if EXTRA_EXCL.exists():
        out |= {x.strip() for x in EXTRA_EXCL.read_text(encoding="utf-8").splitlines() if x.strip() and not x.startswith("#")}
    for d in extra_docs or []:  # ほかの作業が取った書類のフォルダ（中身のハッシュだけを読む。開かない）
        for dp, _dn, fn in os.walk(d):
            for f in fn:
                try:
                    out.add(sha1_b32(Path(dp, f).read_bytes()))
                except OSError:
                    pass
    return out


def _key(seed: int, digest: str) -> str:
    return hashlib.sha256(f"{seed}:{digest}".encode()).hexdigest()


def select_stratum(cands: list[dict], n: int, seed: int, excl: set, m: dict, B: dict, base: dict, tag: str) -> tuple[list[dict], int]:
    """(選んだ行, 新しい候補の数)。cands は 1 つの層の候補（digest で 1 件にまとめたもの）。"""
    pool = [c for c in cands if c["digest"] not in excl and f"url:{c.get('url')}" not in excl]
    pool.sort(key=lambda c: _key(seed, c["digest"]))
    take = pool[:n]
    out = []
    for rank, c in enumerate(take):
        h = c["block_stratum"]
        pi = (m[h] / B[h]) * (len(take) / len(pool))
        out.append({"id": f"S2-{base['country']}-{base['family']}-{tag}{rank:03d}", **base, "rank": rank, "pi": pi, "crawl": CRAWL, **c})
    return out, len(pool)


def cmd_extras(a) -> int:
    """見た書類のうち sample.jsonl にないものを stage2_exclude_extra.txt に：フォルダの書類の中身の SHA-1（開かない）と、*.jsonl の file_url。"""
    lines = set()
    for d in a.exclude_dir:
        for dp, _dn, fn in os.walk(d):
            for f in fn:
                p = Path(dp, f)
                if p.suffix == ".jsonl":
                    for ln in p.read_text(encoding="utf-8", errors="replace").splitlines():
                        try:
                            u = json.loads(ln).get("file_url")
                        except (json.JSONDecodeError, AttributeError):
                            continue
                        if u:
                            lines.add(f"url:{u}")
                else:
                    try:
                        lines.add(sha1_b32(p.read_bytes()))
                    except OSError:
                        pass
    EXTRA_EXCL.write_text("# 第 2 段から除く、見た書類（scripts/eval/v2/stage2_select.py extras が作る）。digest（中身の SHA-1 の base32）か url:\n" + "\n".join(sorted(lines)) + "\n", encoding="utf-8")
    print(f"除く項目 {len(lines)}")
    return 0


def cmd_select(a) -> int:
    sys.path.insert(0, str(HERE.parent.parent))
    import world_frame as wf

    plan = json.loads(PLAN.read_text(encoding="utf-8"))
    out = Path(a.out)
    (out / "cand").mkdir(parents=True, exist_ok=True)
    wf.CACHE = out / "cc"
    doms = json.loads((W / "domains.json").read_text(encoding="utf-8"))["countries"]
    excl = exclusion_digests(a.subset, a.exclude_dir)
    print(f"除く書類 {len(excl):,}", flush=True)
    idx = wf.ClusterIndex()
    want = defaultdict(dict)
    for s in plan["allocation"]:
        want[(s["country"], s["domain_source"])][s["family"]] = s["n"]
    used_blocks = 0
    rows, log = [], {}
    for (cc, src), fam_n in sorted(want.items()):
        cpath = out / "cand" / f"{cc}.{src}.json"
        d = doms[cc]
        own_s = [wf.surt(x["suffix"]) for x in d["suffixes"] if "excluded" not in x] if src == "suffix" else []
        own_h = {wf.surt(h) for h, _ in d["hosts"]} if src == "wikidata" else set()
        saved = json.loads(cpath.read_text(encoding="utf-8")) if cpath.exists() else None
        if saved is None:
            blocks = idx.blocks(own_s + sorted(own_h))
            strata = {h: sorted(i for i, t in blocks.items() if t == h) for h in ("core", "edge")}
            B = {h: len(v) for h, v in strata.items()}
            total = sum(B.values())
            first = min(total, plan["blocks"][f"{cc}:{src}"]["first"])
            m0 = dict(B) if first >= total else {h: (max(1, round(first * B[h] / total)) if B[h] else 0) for h in B}
            rnd = random.Random(f"{plan['seed']}:{cc}:{src}")
            chosen = {h: (rnd.sample(strata[h], min(m0[h], B[h])) if m0[h] < B[h] else list(strata[h])) for h in B}
            rest = [(i, h) for h in B for i in strata[h] if i not in set(chosen[h])]
            random.Random(f"{plan['seed']}:{cc}:{src}:extend").shuffle(rest)
            order = [(i, h) for h in B for i in sorted(chosen[h])]
            n_first = len(order)
            order += rest

            def mine(host_surt: str) -> bool:
                if any(host_surt == p or host_surt.startswith(p + ",") for p in own_s):
                    return True
                parts = host_surt.split(",")
                return any(",".join(parts[:k]) in own_h for k in range(2, len(parts) + 1))

            cand = {f: {} for f in fam_n}
            m = {h: 0 for h in B}
            stop = "exhausted"
            for k, (i, h) in enumerate(order):
                if k >= n_first and all(sum(1 for x in cand[f] if x not in excl) >= fam_n[f] for f in fam_n):
                    stop = "enough"
                    break
                if k >= plan["extra_cap"] and k >= n_first:
                    stop = "cap"
                    break
                if used_blocks >= plan["block_budget"] and k >= n_first:
                    stop = "budget"
                    break
                used_blocks += 1
                m[h] += 1
                for line in idx.read(i):
                    try:
                        key, ts, js = line.split(" ", 2)
                    except ValueError:
                        continue
                    if not mine(wf.line_host(key)):
                        continue
                    r = json.loads(js)
                    if r.get("status") != "200":
                        continue
                    fam, by = wf.family_of(r)
                    if fam not in cand or not r.get("digest") or r["digest"] in cand[fam]:
                        continue
                    cand[fam][r["digest"]] = {"url": r["url"], "timestamp": ts, "mime": r.get("mime"), "mime_detected": r.get("mime-detected"), "digest": r["digest"],
                                              "filename": r["filename"], "offset": int(r["offset"]), "length": int(r["length"]), "block": i, "block_stratum": h,
                                              **({"family_by": by} if by != "mime" else {}), **({"truncated": r["truncated"]} if r.get("truncated") else {})}
            saved = {"B": B, "m": m, "stop": stop, "cand": {f: list(v.values()) for f, v in cand.items()}}
            tmp = cpath.with_suffix(".tmp")
            tmp.write_text(json.dumps(saved, ensure_ascii=False), encoding="utf-8")
            tmp.replace(cpath)
        else:
            used_blocks += sum(saved["m"].values())
        for fam, n in sorted(fam_n.items()):
            base = {"country": cc, "family": fam, "domain_source": src}
            got, pool = select_stratum(saved["cand"].get(fam) or [], n, plan["seed"], excl, saved["m"], saved["B"], base, "" if src == "suffix" else "w")
            rows += got
            log[f"{cc}:{fam}:{src}"] = {"want": n, "got": len(got), "fresh_pool": pool, "stop": saved["stop"], "blocks_read": saved["m"]}
        print(f"[{cc} {src}] ブロック {sum(saved['m'].values())}・{saved['stop']}・累計 {used_blocks}", flush=True)
    # 処理の順：種と id のハッシュ（結果によらない。時間切れで止まっても無作為な一部になる）
    rows.sort(key=lambda r: hashlib.sha256(f"{plan['seed']}:order:{r['id']}".encode()).hexdigest())
    assert not ({r["digest"] for r in rows} & excl) and not ({f"url:{r['url']}" for r in rows} & excl)
    tmp = out / "subset.json.tmp"
    tmp.write_text(json.dumps({"seed": plan["seed"], "plan_sha256": hashlib.sha256(PLAN.read_bytes()).hexdigest(), "blocks_read": used_blocks, "log": log, "docs": rows}, ensure_ascii=False, indent=0), encoding="utf-8")
    tmp.replace(out / "subset.json")
    print(f"選んだ書類 {len(rows)}（計画 {plan['n_total']}）・ブロック {used_blocks}", flush=True)
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["plan", "select", "extras"])
    ap.add_argument("--out", default="")
    ap.add_argument("--subset", default=None, help="固定版の一部（subset.json）")
    ap.add_argument("--exclude-dir", action="append", default=[], help="ほかの作業が取った書類のフォルダ（中身のハッシュで除く）")
    a = ap.parse_args()
    return {"plan": cmd_plan, "select": cmd_select, "extras": cmd_extras}[a.cmd](a)


if __name__ == "__main__":
    raise SystemExit(main())
