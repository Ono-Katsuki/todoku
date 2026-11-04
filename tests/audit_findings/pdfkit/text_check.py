"""Exact check for page-content text: every written value must appear in the text the renderer itself extracts (PDFKit, poppler)."""
import os
import json, re, subprocess, sys, collections, unicodedata
PK = os.environ["TODOKU_PK_DIR"].rstrip("/") + "/"
style = sys.argv[1]
T = json.load(open(PK + "var/field.json"))
S = json.load(open(PK + f"var/{style}.json"))
def norm(s):
    return re.sub(r"\s", "", unicodedata.normalize("NFKC", s))
miss = collections.Counter(); tot = 0; ex = []
widget_seen = 0
for rel, v in T.items():
    out = S[rel]["out"]
    d = json.loads(subprocess.run([PK + "pkr", "dump", out], capture_output=True, text=True).stdout)
    for t in v["targets"]:
        if t["kind"] != "text" or t["source"] != "added":
            continue
        tot += 1
        pg = t["page"]
        kit = norm(d["pages"][pg - 1]["text"])
        pop = norm(subprocess.run(["pdftotext", "-f", str(pg), "-l", str(pg), out, "-"], capture_output=True).stdout.decode("utf-8", "replace"))
        w = norm(t["text"])
        if w not in kit:
            miss["pdfkit"] += 1; ex.append(("pdfkit", rel, t["id"], t["text"]))
        if w not in pop:
            miss["poppler"] += 1; ex.append(("poppler", rel, t["id"], t["text"]))
    # hidden-widget behaviour in PDFKit
    for p in d["pages"]:
        for a in p["annots"]:
            if a["type"] == "Widget" and a.get("shouldDisplay"):
                widget_seen += 1
print(style, "added text values", tot, "not in extracted text", dict(miss), "PDFKit widgets with shouldDisplay", widget_seen)
for e in ex[:12]:
    print(" ", e)
