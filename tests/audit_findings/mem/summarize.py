"""監査用：fuzz.py の results.jsonl をまとめる（形式ごとの数・殺されたもの・遅いもの・大きいもの・内部の例外）。

使い方：summarize.py 変異体のフォルダ
"""
import collections
import json
import os
import re
import sys


def main():
    rows = [json.loads(line) for line in open(os.path.join(sys.argv[1], "results.jsonl"))]
    by = collections.Counter(r["fixture"] for r in rows)
    print("cases", len(rows), dict(by))
    kills = collections.Counter()
    slow = collections.Counter()
    big = collections.Counter()
    crash = collections.Counter()
    ex = {}
    for r in rows:
        tag = re.sub(r":\d+$", "", r["tag"])
        tag = re.sub(r"(stream16?:[^:]*):0x[0-9a-f]+", r"\1:*", tag)
        tag = re.sub(r"flip:.*|partflip:.*", "flip", tag)
        if r["kills"]:
            k = (r["fixture"], tag, r["kills"][0]["tool"], r["kills"][0]["why"][:12])
            kills[k] += 1
            ex.setdefault(k, r["file"])
        elif r["slow"]:
            k = (r["fixture"], tag, r["slow"][0][0], max(s for _t, s in r["slow"]))
            slow[k] += 1
            ex.setdefault(k, r["file"])
        elif r["peak"] > 400:
            k = (r["fixture"], tag, r["peak"])
            big[k] += 1
            ex.setdefault(k, r["file"])
        for c in r["crash"]:
            crash[re.sub(r"\d{3,}", "N", c)] += 1
            ex.setdefault(re.sub(r"\d{3,}", "N", c), r["file"])
    print("\n== killed")
    for k, n in kills.most_common():
        print(n, k, ex[k])
    print("\n== slow (>10s)")
    for k, n in slow.most_common():
        print(n, k, ex[k])
    print("\n== peak >400MB")
    for k, n in big.most_common():
        print(n, k, ex[k])
    print("\n== internal exceptions (not ApiError)", sum(crash.values()))
    for k, n in crash.most_common(40):
        print(n, k[:170], ex[k])


if __name__ == "__main__":
    main()
