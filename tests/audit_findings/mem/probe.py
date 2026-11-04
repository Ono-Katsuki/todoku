"""監査用の子プロセス：1 つの書類を api.* で直接（isolate なし）動かす。

使い方：probe.py 書類 道具,道具,... [出力先フォルダ]
道具の前後に 1 行ずつ JSON を出す。親（guard.py）は START を見て時間を数え直し、殺された道具を知る。
"""
import json
import os
import resource
import sys
import time
import traceback

os.environ["TODOKU_ISOLATE"] = "0"
from todoku import api  # noqa: E402

ALL = ["summary", "window", "fields", "grade", "audit", "fill", "view_page", "read_table",
       "calc_model", "calculate", "run_macro"]


def rss_mb():
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1048576  # macOS はバイト


def first_field(path):
    try:
        fl = api.fields(path)
    except Exception:  # noqa: BLE001
        return {}
    for f in fl:
        if not f.get("readonly"):
            return {f["id"]: "x" if f.get("kind") not in ("checkbox", "bool") else True}
    return {}


def main():
    path = sys.argv[1]
    tools = sys.argv[2].split(",") if len(sys.argv) > 2 and sys.argv[2] != "all" else ALL
    outdir = sys.argv[3] if len(sys.argv) > 3 else os.path.dirname(path)
    header = os.environ.get("PROBE_HEADER", "品目")
    cell = os.environ.get("PROBE_CELL", "B9")
    macro = os.environ.get("PROBE_MACRO", "Test")
    page = int(os.environ.get("PROBE_PAGE", "1"))
    dpi = float(os.environ["PROBE_DPI"]) if os.environ.get("PROBE_DPI") else None
    out = os.path.join(outdir, "out_" + os.path.basename(path))
    run = {
        "summary": lambda: api.summary(path),
        "window": lambda: sum(len(p[1]) for p in api.window(path)[2]),
        "fields": lambda: len(api.fields(path)),
        "grade": lambda: api.grade(path).get("headline"),
        "audit": lambda: api.audit(path).get("ok"),
        "fill": lambda: api.fill(path, first_field(path), out, overwrite=True).get("proof_ok"),
        "view_page": lambda: api.view_page(path, page, dpi=dpi).get("width"),
        "read_table": lambda: api.read_table(path, header).get("rows"),
        "calc_model": lambda: len(str(api.calc_model(path))),
        "calculate": lambda: len(str(api.calculate(path, {cell: 1}))),
        "run_macro": lambda: api.run_macro(path, macro, out + ".m", overwrite=True).get("done"),
    }
    for t in tools:
        print(json.dumps({"ev": "START", "tool": t}), flush=True)
        t0 = time.time()
        rec = {"ev": "END", "tool": t}
        try:
            r = run[t]()
            rec["ok"] = True
            rec["r"] = str(r)[:80]
        except api.ApiError as e:
            rec["ok"] = False
            rec["err"] = "ApiError:" + str(e)[:100]
        except MemoryError as e:
            rec["ok"] = False
            rec["err"] = "MemoryError:" + str(e)[:100]
        except ValueError as e:
            rec["ok"] = False
            rec["err"] = type(e).__name__ + ":" + str(e)[:100]
        except RecursionError:
            rec["ok"] = False
            rec["err"] = "RecursionError"
            rec["crash"] = True
        except Exception as e:  # noqa: BLE001
            rec["ok"] = False
            rec["err"] = type(e).__name__ + ":" + str(e)[:100]
            rec["crash"] = True
            tb = traceback.extract_tb(e.__traceback__)
            rec["at"] = f"{tb[-1].filename.split('/src/')[-1]}:{tb[-1].lineno}"
        rec["s"] = round(time.time() - t0, 2)
        rec["rss"] = round(rss_mb())
        print(json.dumps(rec, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
