"""Re-fill the 46 dev PDFs with the final/ values, under the style given in TODOKU_PDF_STYLE. usage: run_fill.py <style> [--audit]"""
import glob, json, os, sys, time, traceback
BASE = os.environ["TODOKU_PK_DIR"].rstrip("/") + "/../"
R8 = os.environ["TODOKU_PK_FORMS"].rstrip("/") + "/"
style = sys.argv[1]
do_audit = "--audit" in sys.argv
only = [a for a in sys.argv[2:] if not a.startswith("--")]
from todoku import api
from todoku.pdf.doc import PdfDocument
from todoku.pdf.form import read_fields
from todoku.pdf.layout import View
OUT = BASE + f"pk/var/{style}/"
res = {}
forms = sorted(glob.glob(R8 + "redo/out/*/*.pdf"))
for path in forms:
    rel = path.split("/redo/out/")[1]
    if only and rel not in only:
        continue
    vj = R8 + "final/" + rel + ".values.json"
    v = json.load(open(vj))
    out = OUT + rel
    os.makedirs(os.path.dirname(out), exist_ok=True)
    t0 = time.time()
    rec = {"form": v["form"]}
    try:
        r = api.fill(v["form"], v["values"], out, overwrite=True)
        rec.update({k: r.get(k) for k in ("proof_ok", "read_back_ok", "read_back_differences", "unexpected", "structure_errors", "circled", "written_rects", "proof")})
        rec["warnings"] = r.get("warnings")
        rec["out"] = r.get("out_path")
        rec["written"] = r.get("written")
        rec["still_empty_count"] = r.get("still_empty_count")
    except Exception as e:
        rec["refused"] = f"{type(e).__name__}: {str(e)[:1500]}"
    rec["sec"] = round(time.time() - t0, 1)
    if do_audit:
        try:
            from todoku.pdf.audit import run as arun
            t1 = time.time()
            au = arun(open(v["form"], "rb").read(), os.path.basename(v["form"]))
            rec["audit"] = {"ok": getattr(au, "ok", None), "summary": [(c.name, c.ok, str(getattr(c, "detail", ""))[:300]) for c in au.checks if not c.ok]}
            rec["audit_sec"] = round(time.time() - t1, 1)
        except Exception as e:
            rec["audit"] = {"error": f"{type(e).__name__}: {str(e)[:300]}"}
    if style == "field" and "refused" not in rec:
        # the value targets: every field whose widgets are new or whose /V changed, in view coordinates (pt, origin top-left)
        a = PdfDocument(open(v["form"], "rb").read())
        b = PdfDocument(open(rec["out"], "rb").read())
        old = {f.name: f.value_text() for f in read_fields(a)}
        pages = b.pages()
        T = []
        for f in read_fields(b):
            if f.name in old and old[f.name] == f.value_text():
                continue
            for w in f.widgets:
                if not isinstance(w.page, int) or (f.kind in ("checkbox", "radio") and (w.state or "Off") == "Off"):
                    continue
                view = View(b, pages[w.page])
                pts = [view.pt(x, y) for x in (w.rect[0], w.rect[2]) for y in (w.rect[1], w.rect[3])]
                r_ = [min(p[0] for p in pts), min(p[1] for p in pts), max(p[0] for p in pts), max(p[1] for p in pts)]
                wd = b.resolve(w.ref) if w.ref is not None else b.resolve(f.ref)
                da = wd.get("DA") or (b.resolve(f.ref).get("DA") if f.ref is not None else None)
                da = bytes(b.resolve(da)).decode("latin-1") if da is not None else None
                T.append(dict(id=f.name, kind=f.kind, page=w.page + 1, rect=[round(x, 2) for x in r_], text=str(f.value_text()), da=da,
                              source="added" if f.name not in old else "acroform", ff=int(f.flags), rot=int(view.rot),
                              unit=float(view.unit)))
        for c in [c for c in (rec.get("circled") or []) if c.get("box_mm")]:
            mm = 72 / 25.4
            T.append(dict(id=c["id"], kind="circle", page=c["page"], rect=[round(x * mm, 2) for x in c["box_mm"]], text=c["text"], source="circle", da=None, ff=0, rot=0, unit=1.0))
        rec["targets"] = T
    res[rel] = rec
    print(rel, rec.get("proof_ok"), rec.get("read_back_ok"), rec.get("sec"), (rec.get("refused") or "")[:200], flush=True)
name = BASE + f"pk/var/{style}.json" if not only else BASE + f"pk/var/{style}.partial.json"
json.dump(res, open(name, "w"), ensure_ascii=False, indent=0, default=str)
