"""実物の表示で、書いた字が欠けないかを測る道具（PDFKit・Quartz・Poppler・pdf.js）。手で動かす（試験には入らない）。

使い方（TODOKU_PK_DIR：作業の置き場。TODOKU_PK_FORMS：final/ と redo/out/ のある様式の置き場）:
  swiftc -O pkr.swift -o $TODOKU_PK_DIR/pkr         # PDFKit・Quartz の描画（窓を出さない）
  cp pdfjsb.mjs $TODOKU_PK_DIR/
  （pdf.js は $TODOKU_PK_DIR/../pdfjsdl に npm で pdfjs-dist と playwright-core を入れる。Chrome を見えないまま使う）
  PYTHONPATH=src python run_fill.py field           # 様式に書き、書いた値の位置を var/field.json に
  PYTHONPATH=src python make_controls.py            # 同じ字を幅に余裕のある所に描いた対照（欠けない字の墨の量と幅）
  python3 compare.py field                          # 各表示の画像と元の様式の画像の差から、値ごとに whole／clipped／missing
  python3 text_check.py field                       # PDFKit・Poppler が取り出す字に、値がそのままあるか
  python3 ladder.py                                 # 欄の幅を変えて、PDFKit が欠けなくなる幅を測る
測るたびに、実物の目で切り抜きを確かめること（墨の量は、印刷された字に重なる所で少なく出る）。"""
"""Render each style through each renderer, measure every written value against a no-clip control.
usage: compare.py <style> [<style>...]   (python3 with numpy+PIL; one process at a time)
Result: var/cmp_<style>.json. Crops of flagged values: var/crops/<style>/<renderer>/."""
import json, os, re, shutil, subprocess, sys, time
import numpy as np
from PIL import Image
BASE = os.environ["TODOKU_PK_DIR"].rstrip("/") + "/"
DPI = 200
K = DPI / 72
RENDERERS = ("poppler", "pdfkit", "quartz", "pdfjs", "pdfview")
rd_now = ""
ORIG_STYLE = "_orig"
ctrl_meta = json.load(open(BASE + "var/control_rows.json"))
targets_all = json.load(open(BASE + "var/field.json"))


def render(jobs_by_renderer):
    """jobs: renderer -> list of (pdf, page, png)"""
    for rd, jobs in jobs_by_renderer.items():
        jobs = [j for j in jobs if not os.path.exists(j[2])]
        if not jobs:
            continue
        for j in jobs:
            os.makedirs(os.path.dirname(j[2]), exist_ok=True)
        if rd == "poppler":
            for pdf, pg, png in jobs:
                tmp = png[:-4]
                subprocess.run(["nice", "-n", "10", "pdftoppm", "-cropbox", "-r", str(DPI), "-f", str(pg), "-l", str(pg), "-png", "-singlefile", pdf, tmp], check=False, capture_output=True)
        elif rd == "pdfjs":
            tsv = BASE + "var/_jobs.tsv"
            open(tsv, "w").write("".join(f"{p}\t{g}\t{o}\t{DPI}\n" for p, g, o in jobs))
            r = subprocess.run(["nice", "-n", "10", "node", BASE + "pdfjsb.mjs", tsv], capture_output=True, text=True)
            if r.returncode:
                print("pdfjs err", r.stderr[-500:], flush=True)
        else:
            tsv = BASE + "var/_jobs.tsv"
            open(tsv, "w").write("".join(f"{p}\t{g}\t{o}\t{DPI}\t{rd}\n" for p, g, o in jobs))
            r = subprocess.run(["nice", "-n", "10", BASE + "pkr", "batch", tsv], capture_output=True, text=True)
            if r.returncode:
                print(rd, "err", r.stderr[-500:], flush=True)


def gray(png):
    if not os.path.exists(png):
        return None
    return np.asarray(Image.open(png).convert("L"), dtype=np.uint8)


def control_table(kind, renderer):
    """(text,size) -> (ink, width)"""
    pdf = BASE + f"var/control_{kind}.pdf"
    jobs = [(pdf, p + 1, BASE + f"var/png/control_{kind}/{renderer}/{p + 1}.png") for p in range(ctrl_meta["npages"])]
    render({renderer: jobs})
    out = {}
    imgs = {p + 1: gray(jobs[p][2]) for p in range(ctrl_meta["npages"])}
    for r in ctrl_meta["rows"]:
        im = imgs[r["page"]]
        if im is None:
            continue
        zone = im[int((r["y0"] - 5) * K):int((r["y1"] + 5) * K), :]
        mask = zone < 195
        xs = np.where(mask.any(axis=0))[0]
        out[(r["text"], r["size"], r["h"])] = (int(mask.sum()), float((xs[-1] - xs[0]) / K) if len(xs) else 0.0)
    return out


def measure(new, old, targets):
    """Assign every dark pixel the value added to its nearest target; return per-target ink+bbox (pt)."""
    res = [dict(ink=0, bbox=None) for _ in targets]
    if new is None or old is None or new.shape != old.shape:
        return None
    ys, xs = np.nonzero((old.astype(np.int16) - new) > 60)
    if len(xs) == 0:
        return res
    px, py = xs / K, ys / K
    ds = []
    for t in targets:
        r = t["rect"]
        dx = np.maximum(np.maximum(r[0] - px, px - r[2]), 0)
        dy = np.maximum(np.maximum(r[1] - py, py - r[3]), 0)
        ds.append(np.hypot(dx, dy))
    ds = np.stack(ds)
    near = ds.argmin(axis=0)
    dist = ds.min(axis=0)
    for i in range(len(targets)):
        m = (near == i) & (dist < 60)
        if m.any():
            res[i]["ink"] = int(m.sum())
            res[i]["bbox"] = [float(px[m].min()), float(py[m].min()), float(px[m].max()), float(py[m].max())]
    return res


def judge(t, m, ctrl, ref_ink):
    """whole / clipped / missing / n/a, with the numbers."""
    if m is None:
        return dict(state="norender")
    if t["kind"] in ("circle", "checkbox", "radio"):
        if m["ink"] == 0:
            return dict(state="missing")
        if ref_ink and m["ink"] < 0.3 * ref_ink:
            return dict(state="clipped", ratio=round(m["ink"] / ref_ink, 3))
        return dict(state="whole")
    mt = re.search(r"([\d.]+)\s+Tf", t["da"] or "")
    size = float(mt[1]) if mt else 0.0
    if not t["text"].strip():
        return dict(state="whole")
    if m["ink"] == 0:
        return dict(state="missing")
    key = (t["text"].replace("\n", " "), size, round(t["rect"][3] - t["rect"][1], 1))
    c = ctrl.get(key)
    if not c or size <= 0:
        return dict(state="auto" if ref_ink and m["ink"] > 0 else "whole", ratio=round(m["ink"] / ref_ink, 3) if ref_ink else None)
    ratio = m["ink"] / max(1, c[0])
    w = m["bbox"][2] - m["bbox"][0]
    wr = w / c[1] if c[1] else 1.0
    # lost width beyond rounding (1.5pt or 7%), or (a wrapped, narrower text) lost ink
    if rd_now == "pdfview":
        return dict(state="clipped" if c[1] - w > max(2.0, 0.1 * c[1]) else "whole", ratio=round(ratio, 3), wr=round(wr, 3))
    if c[1] - w > max(1.5, 0.07 * c[1]) and ratio < 0.93 or ratio < 0.70:
        return dict(state="clipped", ratio=round(ratio, 3), wr=round(wr, 3))
    return dict(state="whole", ratio=round(ratio, 3), wr=round(wr, 3))


def crop(png, t, m, dst):
    im = Image.open(png).convert("RGB")
    r = t["rect"]
    box = [int((r[0] - 30) * K), int((r[1] - 12) * K), int((r[2] + 30) * K), int((r[3] + 12) * K)]
    box = [max(0, box[0]), max(0, box[1]), min(im.width, box[2]), min(im.height, box[3])]
    im.crop(box).save(dst)


def run(style):
    t0 = time.time()
    style_res = json.load(open(BASE + f"var/{style}.json"))
    out = {}
    tally = {}
    global rd_now
    for rd in RENDERERS:
        rd_now = rd
        ctrl_field = control_table("field", rd)
        ctrl_content = control_table("content", rd)
        ctrl = ctrl_field if (rd == "pdfkit" and style in ("field", "widen")) or (rd == "pdfview" and style in ("field", "widen")) else ctrl_content
        jobs_new, jobs_old = [], []
        for rel, v in targets_all.items():
            pages = sorted({t["page"] for t in v["targets"]})
            pdf = style_res[rel].get("out") or BASE + f"var/{style}/{rel}"  # a failed proof publishes under a marked name
            slug = rel.replace("/", "__")
            for p in pages:
                jobs_old.append((v["form"], p, BASE + f"var/png/{ORIG_STYLE}/{rd}/{slug}.{p}.png"))
                jobs_new.append((pdf, p, BASE + f"var/png/{style}/{rd}/{slug}.{p}.png"))
        render({rd: jobs_old})
        render({rd: jobs_new})
        print(style, rd, "rendered", round(time.time() - t0), flush=True)
        for rel, v in targets_all.items():
            slug = rel.replace("/", "__")
            for p in sorted({t["page"] for t in v["targets"]}):
                ts = [t for t in v["targets"] if t["page"] == p]
                new_png = BASE + f"var/png/{style}/{rd}/{slug}.{p}.png"
                ms = measure(gray(new_png), gray(BASE + f"var/png/{ORIG_STYLE}/{rd}/{slug}.{p}.png"), ts)
                for i, t in enumerate(ts):
                    m = ms[i] if ms else None
                    ref = None
                    pref = out.get(rel, {}).get(f"{p}:{t['id']}:{i}", {}).get("poppler")
                    if pref:
                        ref = pref.get("ink")
                    j = judge(t, m, ctrl, ref)
                    j.update(ink=(m or {}).get("ink"), bbox=(m or {}).get("bbox"))
                    if m and m["bbox"]:
                        rr = t["rect"]
                        b = m["bbox"]
                        j["spill"] = round(max(rr[0] - b[0], b[2] - rr[2], rr[1] - b[1], b[3] - rr[3], 0), 1)
                    out.setdefault(rel, {}).setdefault(f"{p}:{t['id']}:{i}", dict(kind=t["kind"], text=t["text"], source=t["source"]))[rd] = j
                    if j["state"] in ("clipped", "missing") and rd in ("pdfkit", "poppler", "pdfjs") and os.path.exists(new_png):
                        dd = BASE + f"var/crops/{style}/{rd}/"
                        os.makedirs(dd, exist_ok=True)
                        crop(new_png, t, m, dd + f"{slug}.{p}.{i}.png")
                os.path.exists(new_png) and os.remove(new_png)
    json.dump(out, open(BASE + f"var/cmp_{style}.json", "w"), ensure_ascii=False)
    print(style, "done", round(time.time() - t0), flush=True)


for st in sys.argv[1:]:
    run(st)
