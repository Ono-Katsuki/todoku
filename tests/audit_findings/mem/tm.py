"""監査用：何がメモリを持っているか（tracemalloc の上位）。小さめの書類で動かす。

使い方：tm.py 書類 道具   （nice 10。guard の見張りはないので、小さな書類だけに使う）
"""
import os
import sys
import tracemalloc

os.environ["TODOKU_ISOLATE"] = "0"
os.nice(10)
from todoku import api  # noqa: E402

path, tool = sys.argv[1], sys.argv[2]
tracemalloc.start(1)
if tool == "audit":
    api.audit(path)
elif tool == "summary":
    api.summary(path)
elif tool == "fill":
    fid = next(f["id"] for f in api.fields(path))
    api.fill(path, {fid: "x"}, path + ".out", overwrite=True)
cur, peak = tracemalloc.get_traced_memory()
print(f"current={cur >> 20}MB peak={peak >> 20}MB")
for st in tracemalloc.take_snapshot().statistics("lineno")[:10]:
    print(f"{st.size >> 20:5d}MB {st.count:9d} {str(st.traceback[0]).split('/src/')[-1][:80]}")
