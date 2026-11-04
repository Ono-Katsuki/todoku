"""監査用：マクロ（VBA）の中の大きな範囲・配列・繰り返し。vbarun.run_macro を別のプロセスで直接動かす（nice 10・1.5 GB・60 秒）。

使い方：vba_cases.py [名前の一部 ...]   （子：vba_cases.py --child 名前）
"""
import fcntl
import json
import os
import resource
import subprocess
import sys
import tempfile
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
LOCK = os.path.join(tempfile.gettempdir(), "todoku_mem_guard.lock")


def sub(body):
    return "Sub Test()\n" + body + "\nEnd Sub\n"


CASES = {
    "for_each_all": sub('Dim c\nFor Each c In Range("A1:XFD1048576")\nNext'),
    "for_each_col": sub('Dim c\nFor Each c In Range("A:A")\nNext'),
    "for_each_cells": sub("Dim c\nFor Each c In Cells\nNext"),
    "value_all": sub('Dim v\nv = Range("A1:XFD1048576").Value'),
    "value_2m": sub('Dim v\nv = Range("A1:B1000000").Value'),
    "set_all": sub('Range("A1:XFD1048576").Value = 1'),
    "set_col": sub('Range("A:A").Value = 1'),
    "formula_col": sub('Range("A:A").Formula = "=1+1"'),
    "clear_all": sub('Range("A1:XFD1048576").Clear'),
    "find_all": sub('Dim f\nSet f = Range("A1:XFD1048576").Find("zz")'),
    "hasformula_all": sub('Dim v\nv = Range("A1:XFD1048576").HasFormula'),
    "currentregion": sub('Dim v\nv = Range("A1").CurrentRegion.Rows.Count'),
    "wsf_sum_all": sub('Dim v\nv = Application.WorksheetFunction.Sum(Range("A1:XFD1048576"))'),
    "wsf_rept": sub('Dim v\nv = Application.WorksheetFunction.Rept("x", 2000000000)'),
    "wsf_seq": sub('Dim v\nv = Application.WorksheetFunction.Sum(Evaluate("SEQUENCE(1048576,16384)"))'),
    "evaluate_all": sub('Dim v\nv = Evaluate("SUM(A:XFD)")'),
    "redim_max": sub("Dim a()\nReDim a(1 To 10000000)\nDim v\nFor Each v In a\nNext"),
    "redim_fill": sub("Dim a()\nReDim a(1 To 10000000)\nDim i\nFor i = 1 To 10000000\na(i) = i\nNext"),
    "redim_many": sub("Dim i\nFor i = 1 To 100000\nDim a()\nReDim a(1 To 10000000)\nNext"),
    "array_copy": sub("Dim a(), b()\nReDim a(1 To 3000000)\nDim i\nFor i = 1 To 50\nb = a\nNext"),
    "array_keep": sub("Dim a()\nDim i\nFor i = 1 To 10000\nReDim Preserve a(1 To i * 1000)\nNext"),
    "string_space": sub("Dim s\ns = Space(2000000000)"),
    "string_double": sub('Dim s\ns = "x"\nDim i\nFor i = 1 To 60\ns = s & s\nNext'),
    "string_many_small": sub('Dim a(1 To 10000000)\nDim i\nFor i = 1 To 10000000\na(i) = "abcdefghijklmnopqrstuvwxyz" & i\nNext'),
    "string_fn": sub('Dim v\nv = String(2000000000, "x")'),
    "replace_amp": sub('Dim s\ns = String(10000000, "a")\ns = Replace(s, "a", String(1000, "b"))'),
    "split_amp": sub('Dim s\ns = String(10000000, ",")\nDim v\nv = Split(s, ",")'),
    "join_amp": sub('Dim a(1 To 10000000)\nDim s\ns = Join(a, String(100, "x"))'),
    "loop_forever": sub("Do\nLoop"),
    "recursion": "Sub Test()\nTest\nEnd Sub\n",
    "cells_write_far": sub("Dim i\nFor i = 1 To 3000000\nCells(i Mod 1000000 + 1, i Mod 16384 + 1).Value = i\nNext"),
    "rows_insert": sub("Dim i\nFor i = 1 To 1000\nRows(1).Insert\nNext"),
    "offset_many": sub('Dim r\nSet r = Range("A1")\nDim i\nFor i = 1 To 5000000\nSet r = r.Offset(0, 0)\nNext'),
}


def child(name):
    sys.path.insert(0, os.path.join(ROOT, "src"))
    sys.path.insert(0, os.path.join(ROOT, "tests"))
    from test_example import make_xlsx

    from todoku.vba import _PROC, Module, Project
    from todoku.vbarun import run_macro
    from todoku.xlsx.workbook import Workbook

    src = CASES[name]
    wb = Workbook(make_xlsx({"入力": ({"A1": 1, "A2": 2}, [])}))
    wb.has_macros = True
    wb._vba = Project(932, "VBAProject", [Module("Module1", "Module1", 0, "procedural", 'Attribute VB_Name = "Module1"\n' + src,
                                                 [m.group(2) for m in _PROC.finditer(src)])])
    wb.sheets[0].code_name = "Sheet1"
    t0 = time.time()
    try:
        run_macro(wb, "Test")
        r = "done"
    except BaseException as e:  # noqa: BLE001
        r = f"{type(e).__name__}: {str(e)[:80]}"
    print(json.dumps({"r": r, "s": round(time.time() - t0, 2), "rss": round(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1048576)}, ensure_ascii=False))


def rss_kb(pid):
    try:
        return int(subprocess.run(["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True).stdout.strip() or 0)
    except ValueError:
        return 0


def main():
    if len(sys.argv) > 2 and sys.argv[1] == "--child":
        return child(sys.argv[2])
    want = sys.argv[1:]
    for name in CASES:
        if want and not any(w in name for w in want):
            continue
        with open(LOCK, "w") as lk:
            fcntl.flock(lk, fcntl.LOCK_EX)
            p = subprocess.Popen(["nice", "-n", "10", sys.executable, os.path.abspath(__file__), "--child", name],
                                 stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, errors="replace",
                                 cwd=os.path.join(ROOT, "tests"))
            t0 = time.time()
            peak = 0
            why = ""
            while p.poll() is None:
                r = rss_kb(p.pid)
                peak = max(peak, r)
                if r > 1.5 * 1048576:
                    why = "RSS>1.5GB"
                elif time.time() - t0 > 60:
                    why = "time>60s"
                if why:
                    p.kill()
                    break
                time.sleep(0.03)
            out = p.communicate()[0]
        print(f"{name}: {time.time() - t0:.1f}s peak={peak // 1024}MB {why or ''} {out.strip()[-180:] if not why else ''}", flush=True)


if __name__ == "__main__":
    main()
