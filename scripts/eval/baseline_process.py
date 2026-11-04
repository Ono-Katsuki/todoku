"""比べる道具でできること（処理の取り出し）：openpyxl は数式の字と保存された値を読める。計算はしない。xlrd は .xls の数式の字を読まない（値だけ）。
stdlib の zipfile で数式を数え、とどくが数えた数と独立に合わせる。使い方（別の venv で python -I）: baseline_process.py <ファイル>"""

import json
import re
import sys
import zipfile

p = sys.argv[1]
o = {}
try:
    if p.lower().endswith((".xlsx", ".xlsm")):
        z = zipfile.ZipFile(p)
        n = 0
        funcs = {}
        for name in z.namelist():
            if re.fullmatch(r"xl/worksheets/[^/]+\.xml", name):
                b = z.read(name)
                n += len(re.findall(rb"<c\b[^>]*>\s*<f\b", b))
                for f in re.findall(rb"<f\b[^>]*>(.*?)</f>", b, flags=re.S):
                    for fn in re.findall(rb"([A-Z][A-Z0-9_.]+)\(", f):
                        k = fn.decode()
                        funcs[k] = funcs.get(k, 0) + 1
        o["xml_formula_cells"] = n
        o["xml_functions"] = sorted(funcs)
        import openpyxl

        wb = openpyxl.load_workbook(p)
        o["openpyxl_formula_text_cells"] = sum(1 for ws in wb.worksheets for row in ws.iter_rows() for c in row if isinstance(c.value, str) and c.value.startswith("="))
        o["openpyxl_can_compute"] = False
    elif p.lower().endswith(".xls"):
        import xlrd

        xlrd.open_workbook(p)
        o["openpyxl_formula_text_cells"] = 0
        o["openpyxl_can_compute"] = False
        o["note"] = "xlrd は .xls の数式の字を読めない（値だけ）"
except Exception as e:  # noqa: BLE001
    o["error"] = f"{type(e).__name__}: {str(e)[:200]}"
print(json.dumps(o, ensure_ascii=False))
