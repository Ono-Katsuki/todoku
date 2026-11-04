"""監査用：MANIFEST.jsonl の変異体を 1 つずつ guard で動かし、結果を results.jsonl に足していく（途中から続けられる）。

使い方：fuzz.py 変異体のフォルダ [上限の数]
遅いもの（>10 秒）・大きいもの（>400 MB）・殺されたもの・内部の例外だけを「interesting」に出す。
"""
import json
import os
import shutil
import sys
import tempfile
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from guard import guarded  # noqa: E402


def main():
    d = sys.argv[1]
    limit = int(sys.argv[2]) if len(sys.argv) > 2 else 10**9
    res = os.path.join(d, "results.jsonl")
    done = set()
    if os.path.exists(res):
        with open(res) as f:
            done = {json.loads(line)["file"] for line in f if line.strip()}
    rows = [json.loads(line) for line in open(os.path.join(d, "MANIFEST.jsonl"))]
    t0 = time.time()
    n = 0
    tmp = tempfile.mkdtemp(prefix="memfuzz_", dir=d)
    for r in rows:
        if r["file"] in done:
            continue
        if n >= limit:
            break
        n += 1
        g = guarded(os.path.join(d, r["file"]), outdir=tmp)
        r.update(peak=g["peak"], kills=g["kills"])
        slow = [(x["tool"], x["s"]) for x in g["recs"] if x.get("s", 0) > 10]
        crash = sorted({f"{x['tool']}:{x['err'][:70]}@{x.get('at', '')}" for x in g["recs"] if x.get("crash")})
        r.update(slow=slow, crash=crash, tool_peaks={x["tool"]: x.get("peak", 0) for x in g["recs"] if "tool" in x})
        r["errs"] = sorted({x["err"][:40] for x in g["recs"] if x.get("err")})[:4]
        with open(res, "a") as f:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
        if g["kills"] or slow or g["peak"] > 400:
            print("INTERESTING", json.dumps({k: r[k] for k in ("file", "tag", "peak", "kills", "slow")}, ensure_ascii=False), flush=True)
        for fn in os.listdir(tmp):
            os.unlink(os.path.join(tmp, fn))
        if n % 50 == 0:
            print(f"-- {n} done, {time.time() - t0:.0f}s", flush=True)
    shutil.rmtree(tmp, ignore_errors=True)
    print(f"FINISHED {n} in {time.time() - t0:.0f}s", flush=True)


if __name__ == "__main__":
    main()
