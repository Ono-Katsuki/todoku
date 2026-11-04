"""確かめ役のために、記入前・記入後のページの絵を作る。描き手は Poppler（必ず）・todoku の view_page（あれば）・PDFKit（pkr があれば）。
Office の書類は LibreOffice で PDF にしてから描く。記入の前後で画素の変わったページだけを残す（多いときは前から max_pages まで）。
使い方（別の venv で python -I）: render_pages.py <元> <記入後> <作業フォルダ> <LibreOffice の設定> <todoku の src または -> <todoku の python または -> <pkr または -> [max_pages]"""

import json
import os
import subprocess
import sys

from PIL import Image, ImageChops

orig, final, work, lo_profile, tsrc, tpy, pkr = sys.argv[1:8]
MAX_PAGES = int(sys.argv[8]) if len(sys.argv) > 8 else 12
DPI = 110
pages_dir = os.path.join(work, "pages")
os.makedirs(pages_dir, exist_ok=True)
notes = []


def run(cmd, timeout=240):
    return subprocess.run([str(c) for c in cmd], capture_output=True, text=True, timeout=timeout)


def as_pdf(path, tag):
    if path.lower().endswith(".pdf"):
        return path
    d = os.path.join(work, "pdf_" + tag)
    os.makedirs(d, exist_ok=True)
    run(["soffice", f"-env:UserInstallation=file://{lo_profile}", "--headless", "--norestore", "--convert-to", "pdf", "--outdir", d, path], 240)
    p = os.path.join(d, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    if not os.path.exists(p):
        raise RuntimeError(f"PDF にできない：{os.path.basename(path)}")
    notes.append("Office の書類は LibreOffice で PDF にしてから描いた")
    return p


def page_count(pdf):
    for ln in run(["pdfinfo", pdf]).stdout.splitlines():
        if ln.startswith("Pages:"):
            return int(ln.split()[1])
    return 0


def poppler(pdf, page, out):
    pre = out[:-4]
    run(["pdftoppm", "-r", DPI, "-f", page, "-l", page, "-png", "-singlefile", pdf, pre])
    return os.path.exists(out)


o = {}
try:
    a, b = as_pdf(orig, "orig"), as_pdf(final, "final")
    n_a, n_b = page_count(a), page_count(b)
    pages = []
    renderers = ["poppler"]
    for pg in range(1, n_b + 1):
        pb = os.path.join(pages_dir, f"poppler_p{pg}.png")
        pa = os.path.join(pages_dir, f"orig_p{pg}.png")
        if not poppler(b, pg, pb):
            continue
        changed = True
        if pg <= n_a and poppler(a, pg, pa):
            ia, ib = Image.open(pa).convert("L"), Image.open(pb).convert("L")
            if ia.size == ib.size:
                changed = ImageChops.difference(ia, ib).point(lambda v: 255 if v > 40 else 0).getbbox() is not None
        if changed and len(pages) < MAX_PAGES:
            pages.append(pg)
        else:
            for f in (pa, pb):
                if os.path.exists(f):
                    os.remove(f)
    if not pages and n_b:
        notes.append("記入の前後で、Poppler の絵の変わったページがない（字が描かれていない可能性）")
        for pg in range(1, min(n_b, 3) + 1):
            poppler(b, pg, os.path.join(pages_dir, f"poppler_p{pg}.png"))
            poppler(a, pg, os.path.join(pages_dir, f"orig_p{pg}.png"))
            pages.append(pg)
    if pkr != "-" and os.path.exists(pkr):
        ok = 0
        for pg in pages:
            run([pkr, "render", b, pg, os.path.join(pages_dir, f"pdfkit_p{pg}.png"), DPI, "pdfkit"])
            ok += os.path.exists(os.path.join(pages_dir, f"pdfkit_p{pg}.png"))
        if ok:
            renderers.append("pdfkit")
    else:
        notes.append("PDFKit の道具（pkr）がないので、PDFKit の絵はない")
    if tsrc != "-":
        code = ("import sys,json,base64\nsys.path.insert(0,sys.argv[1])\nfrom todoku import api\n"
                "for pg in sys.argv[4:]:\n v=api.view_page(sys.argv[2],int(pg))\n open(f'{sys.argv[3]}/view_p{pg}.png','wb').write(base64.b64decode(v['png_base64']))\n")
        r = run([tpy, "-I", "-c", code, tsrc, b, pages_dir, *pages])
        if r.returncode == 0:
            renderers.append("view")
        else:
            notes.append(f"view_page の絵を作れなかった：{r.stderr[-200:]}")
    o = {"pages": pages, "renderers": renderers, "notes": sorted(set(notes)), "pages_total": n_b, "pdf_final": b}
    with open(os.path.join(work, "pages.json"), "w", encoding="utf-8") as f:
        json.dump(o, f, ensure_ascii=False)
except Exception as e:  # noqa: BLE001
    o = {"error": f"{type(e).__name__}: {str(e)[:300]}"}
print(json.dumps(o, ensure_ascii=False))
