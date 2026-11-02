"""コーパスの全ファイルで、往復のバイト一致と関数の全数照合を行う。"""
import sys, collections, traceback
from pathlib import Path
from todoku.xlsx.workbook import Workbook
from todoku.xlsx.calc import verify
from todoku.zipraw import ZipFormatError

root = Path(sys.argv[1] if len(sys.argv) > 1 else "corpus")
agg = collections.Counter(); reasons = collections.Counter(); mism = []
for p in sorted(root.rglob("*")):
    if p.suffix.lower() not in (".xlsx", ".xlsm", ".bin") and "Xlsdl" not in p.name:
        if p.suffix: continue
    blob = p.read_bytes()
    if blob[:2] != b"PK":
        agg["not_zip(" + ("xls" if blob[:4]==b"\xd0\xcf\x11\xe0" else "other") + ")"] += 1; continue
    try:
        wb = Workbook(blob, str(p))
    except Exception as e:
        agg["load_error"] += 1; print("LOAD", p.name, e); continue
    agg["files"] += 1
    agg["roundtrip_ok" if wb.zip.build() == blob else "roundtrip_NG"] += 1
    try:
        rep = verify(wb)
    except Exception as e:
        agg["verify_crash"] += 1; print("CRASH", p.name); traceback.print_exc(limit=3); continue
    for k, v in rep.summary().items(): agg["f_" + k] += v
    for r in rep.results:
        if r.status == "unsupported": reasons[r.reason[:60]] += 1
        if r.status == "mismatch": mism.append((p.name, r.sheet, r.address, r.formula, r.expected, r.actual))
print(dict(agg))
print("unsupported reasons:", reasons.most_common(15))
print("mismatches:", len(mism))
for m in mism[:30]: print("  ", m)
