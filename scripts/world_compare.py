"""とどくと、エージェントがふだん使う道具との比べ（読む・書く）。scripts/world_validate.py と同じ標本・同じ作業用フォルダを使う。

比べる道具：
  todoku    とどくの窓（Markdown）と fill_document（api.fill）
  markitdown  MarkItDown（Office・PDF を Markdown にする道具。多くのエージェントが使う）
  naive     「エージェントがふだん書く読み方・書き方」：openpyxl（既定の load_workbook）・xlrd・pypdf・python-docx・
            CSV は UTF-8 で開くだけ。書くのは openpyxl（読み込み→セルに代入→保存）・python-docx（マスの text に代入→保存）・
            pypdf（update_page_form_field_values）
基準の読み（正解）は world_validate.py の参照の読み（world_reference.py）で、食い違いは確かめた人が判断したもの
（world/validation/judgments.json）。基準そのものが誤りと判断した書類は、内容の比べから外す。

段階：
  read      1 件ずつ、道具ごとに読んで、内容の再現率・適合率と、特徴（数式・結合・非表示・表示の形式・図形の字・メモ・
            チェックボックス・PDF のフォームの欄・古い形式）を伝えているかを数える。時間も測る
  write     書ける所のある書類に、1 つの値を、道具ごとに書く。書いた後のファイルを元と、ZIP の部品・バイトの単位で比べ、
            何が黙って消えた・変わったか、値が正しい所に入ったかを、とどくでない道具（openpyxl・pypdf など）で確かめる
  summarize compare.json と summary.json の compare の部分を作る

使い方（--out の既定は world/validation、--work は world_validate.py と同じ作業用フォルダ）:
  python scripts/world_compare.py read --work W --todoku-python PY --baseline-python PY --todoku-src SRC
  python scripts/world_compare.py write --work W ... [--extra-results 結果.jsonl --extra-ids ID,ID]
  python scripts/world_compare.py summarize --work W
1 件ごとに別の処理で動かし、2 GB・120 秒を超えたら止めて、その道具の「時間切れ」として記録する。
"""

from __future__ import annotations

import argparse
import collections
import json
import os
import re
import sys
import unicodedata
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import world_validate as wv  # noqa: E402

ROOT = wv.ROOT
TOOLS = ("todoku", "markitdown", "naive")

# ====================================================================== 読む（ふだんの道具）
BASE_READ = r"""
import json, sys, time
tool, p, kind, out = sys.argv[1:5]
t = time.time()
o = {"tool": tool}
txt = None
try:
    if tool == "markitdown":
        from markitdown import MarkItDown
        txt = MarkItDown().convert(p).text_content
    elif kind == "xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(p)  # 既定：数式は数式のまま（保存された計算結果は読まない）
        lines = []
        for ws in wb.worksheets:
            lines.append("# " + ws.title)
            for row in ws.iter_rows(values_only=True):
                vals = [str(v) for v in row if v is not None]
                if vals:
                    lines.append("\t".join(vals))
        txt = "\n".join(lines)
    elif kind == "xls":
        import xlrd
        bk = xlrd.open_workbook(p)
        lines = []
        for sh in bk.sheets():
            lines.append("# " + sh.name)
            for r in range(sh.nrows):
                vals = [str(v) for v in sh.row_values(r) if v not in ("", None)]
                if vals:
                    lines.append("\t".join(vals))
        txt = "\n".join(lines)
    elif kind == "pdf":
        import pypdf
        rd = pypdf.PdfReader(p)
        txt = "\n".join((pg.extract_text() or "") for pg in rd.pages)
    elif kind == "docx":
        import docx
        d = docx.Document(p)
        lines = [para.text for para in d.paragraphs]
        for tb in d.tables:
            for row in tb.rows:
                lines.append("\t".join(c.text for c in row.cells))
        txt = "\n".join(lines)
    elif kind == "csv":
        txt = open(p, encoding="utf-8").read()
    else:
        o["error"] = "unsupported"
except Exception as e:
    o["error"] = f"{type(e).__name__}: {str(e)[:160]}"
if txt is not None:
    open(out, "w", encoding="utf-8").write(txt)
    o["chars"] = len(txt)
o["seconds"] = round(time.time() - t, 2)
print(json.dumps(o, ensure_ascii=False))
"""

# 特徴：書類にあるか（参照の道具で調べる）
FEATURES = r"""
import json, sys, zipfile, io, re, datetime as dt
import xml.etree.ElementTree as ET
p, kind = sys.argv[1:3]
o = {}
def disp(v, fmt):
    # よく使う表示の形式だけ、Excel の見た目を作る（ほかは None：数えない）
    if isinstance(v, bool) or v is None:
        return None
    if isinstance(v, (dt.datetime, dt.date)):
        f = fmt.lower().replace("\\", "").replace('"', "")
        if re.fullmatch(r"yyyy/m+/d+(;@)?|yyyy-mm-dd|yyyy/mm/dd", f):
            sep = "/" if "/" in f else "-"
            mm = "%02d" % v.month if "mm" in f else str(v.month)
            dd = "%02d" % v.day if "dd" in f else str(v.day)
            return f"{v.year}{sep}{mm}{sep}{dd}"
        if re.fullmatch(r"d+\.m+\.yyyy", f):
            return f"{v.day:02d}.{v.month:02d}.{v.year}" if "dd" in f else f"{v.day}.{v.month}.{v.year}"
        if re.fullmatch(r"dd/mm/yyyy", f):
            return f"{v.day:02d}/{v.month:02d}/{v.year}"
        return None
    if not isinstance(v, (int, float)):
        return None
    m = re.fullmatch(r"(#,##)?0(\.0+)?(%)?", fmt.replace("_)", "").replace("_-", "").split(";")[0])
    if not m:
        return None
    dec = len(m.group(2)) - 1 if m.group(2) else 0
    x = v * 100 if m.group(3) else v
    s = f"{x:,.{dec}f}" if m.group(1) else f"{x:.{dec}f}"
    if m.group(3):
        s += "%"
    if s == str(v) or (not m.group(1) and not m.group(3) and dec == 0):
        return None  # 生の値と同じ見た目は数えない
    return s
try:
    if kind == "xlsx":
        import openpyxl
        wf = openpyxl.load_workbook(p)
        wv = openpyxl.load_workbook(p, data_only=True)
        f_s, m_s, h_sh, h_txt, d_s, cm = [], [], [], [], [], []
        for ws in wf.worksheets:
            vs = wv[ws.title]
            hidden_rows = {i for i, d in ws.row_dimensions.items() if d.hidden}
            if ws.sheet_state != "visible":
                h_sh.append(ws.title)
            for row in ws.iter_rows():
                for c in row:
                    if c.value is None:
                        continue
                    v = vs[c.coordinate].value
                    if isinstance(c.value, str) and c.value.startswith("=") and len(f_s) < 40:
                        f_s.append({"cell": f"{ws.title}!{c.coordinate}", "formula": c.value[1:], "cached": None if v is None else str(v)})
                    if (ws.sheet_state != "visible" or c.row in hidden_rows) and isinstance(v, str) and len(v.strip()) >= 4 and len(h_txt) < 20:
                        h_txt.append(v)
                    if c.number_format and c.number_format != "General" and len(d_s) < 40:
                        s = disp(v, c.number_format)
                        if s:
                            d_s.append({"cell": f"{ws.title}!{c.coordinate}", "raw": str(v), "fmt": c.number_format, "shown": s})
                    if c.comment and len(cm) < 20:
                        cm.append(c.comment.text)
            m_s += [str(r) for r in ws.merged_cells.ranges][:20]
        z = zipfile.ZipFile(p)
        tb = []
        A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
        X = "{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}"
        for n in z.namelist():
            if re.match(r"xl/drawings/drawing\d+\.xml$", n):
                for sp in ET.fromstring(z.read(n)).iter(X + "sp"):
                    t = "".join(x.text or "" for x in sp.iter(A + "t"))
                    if len(t.strip()) >= 3:
                        tb.append(t)
        o.update(formulas=f_s, merged=m_s, hidden_sheets=h_sh, hidden_text=h_txt, display=d_s, comments=cm, textboxes=tb[:20])
    elif kind == "docx":
        z = zipfile.ZipFile(p)
        W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
        tb, cm, cbs = [], [], []
        for n in z.namelist():
            if not n.endswith(".xml") or not n.startswith("word/"):
                continue
            root = ET.fromstring(z.read(n))
            for tx in root.iter(W + "txbxContent"):
                t = "".join(x.text or "" for x in tx.iter(W + "t"))
                if len(t.strip()) >= 3:
                    tb.append(t)
            if n == "word/comments.xml":
                for c in root.iter(W + "comment"):
                    cm.append("".join(x.text or "" for x in c.iter(W + "t")))
            for cb in root.iter(W14 + "checkbox"):
                ch = cb.find(W14 + "checked")
                cbs.append(ch is not None and ch.get(W14 + "val") in ("1", "true"))
            for ff in root.iter(W + "checkBox"):
                d = ff.find(W + "default")
                c2 = ff.find(W + "checked")
                cbs.append((c2 is not None and c2.get(W + "val", "1") != "0") or (d is not None and d.get(W + "val") == "1"))
        o.update(textboxes=list(dict.fromkeys(tb))[:20], comments=cm[:20], checkboxes=cbs)
    elif kind == "pdf":
        import pypdf
        rd = pypdf.PdfReader(p)
        fl = rd.get_fields() or {}
        vals, cbs = [], []
        for k, v in fl.items():
            ft, val = v.get("/FT"), v.get("/V")
            if ft == "/Btn":
                cbs.append(str(val) not in ("None", "/Off", ""))
            elif val not in (None, "") and len(str(val)) >= 2:
                vals.append(str(val))
        o.update(form_values=vals[:40], checkboxes=cbs, acroform=len(fl))
except Exception as e:
    o["error"] = f"{type(e).__name__}: {str(e)[:160]}"
print(json.dumps(o, ensure_ascii=False, default=str))
"""


def _n(s: str) -> str:
    return wv.norm(s)


_TODOKU_META = ("### 入力欄", "## 入力欄", "### 規則では決められない所", "## 規則では決められない所", "### 画像", "## この窓の目次")
_TODOKU_NOTE = ("- 非表示", "- シート保護", "- 印刷範囲", "- 文字に直せない字", "- フォームの欄はありません", "- 書き方：", "- 行は「",
                "- **文字コード", "- 名前「")


def todoku_content(md: str) -> str:
    """とどくの窓から、とどくが足した説明（冒頭の説明・入力欄の一覧・規則では決められない所の一覧・画像の一覧・目次・注記の行）を
    除いた、書類の中身の部分（読み順・本文・表・図形の字など）。適合率はこの部分で数える。"""
    out, skip = [], False
    started = False
    for ln in md.splitlines():
        if ln.startswith("#"):
            started = started or ln.startswith("## ")
            skip = ln.startswith(_TODOKU_META)
            continue
        if not started or skip or ln.startswith(_TODOKU_NOTE):
            continue
        out.append(ln)
    return "\n".join(out)


def strip_md(text: str, tool: str) -> str:
    """道具ごとの印（Markdown の表の線・見出しの # ・pandas の NaN と Unnamed）を除き、値だけにする（適合率のため）。"""
    if tool == "todoku":
        return wv.strip_todoku(text)
    t = re.sub(r"^\|?(\s*:?-{3,}:?\s*\|)+\s*$", "", text, flags=re.M)
    t = re.sub(r"^#+\s.*$", "", t, flags=re.M) if tool == "naive" else re.sub(r"^#+\s*", "", t, flags=re.M)
    if tool == "markitdown":
        t = re.sub(r"\bNaN\b|\bUnnamed: \d+\b", "", t)
    return t.replace("|", " ")


def bigram_recall(ref_units: list[str], out: str) -> float | None:
    """2 字の並びの多重集合での再現率：落ちた字・化けた字があると下がる（並びの大きな入れ替えには鈍い）。"""
    a = collections.Counter(u[k:k + 2] for u in ref_units for k in range(len(u) - 1))
    b = collections.Counter()
    for ln in out.splitlines():
        t = wv.norm(ln)
        b.update(t[k:k + 2] for k in range(len(t) - 1))
    tot = sum(a.values())
    return 1 - sum((a - b).values()) / tot if tot else None


def segment_precision(text: str, ref_hay: str) -> float | None:
    """出力の区切り（縦線・タブ・改行）ごとの断片のうち、参照のどこかにある断片の割合（字数で重み）。"""
    tot = hit = 0
    segs = [re.sub(r"\s+", "", re.sub(r"^\s*(?:[-*>]\s+|#+\s+|\d+(?:〜\d+)?:\s)", "", x))
            for x in re.split(r"[|\t\n｜]", unicodedata.normalize("NFKC", text))]
    segs = [x for x in segs if len(x) >= 2]
    if len(segs) > 2000:  # 大きな書類は、種を決めて 2,000 個に間引く
        segs = wv.random.Random(wv.SEED).sample(segs, 2000)
    for s in segs:
        tot += len(s)
        if s in ref_hay:
            hit += len(s)
    return hit / tot if tot else None


def conveys(feat: dict, out: str, tool: str, kind: str) -> dict:
    """書類にある特徴を、その道具の出力が伝えているか（ある書類だけ数える）。"""
    o = _n(out)
    r: dict = {}

    def frac(items, f=lambda x: x):
        items = [x for x in items if len(_n(f(x))) >= 2]
        if not items:
            return None
        return sum(1 for x in items if _n(f(x)) in o) / len(items)

    if feat.get("formulas"):
        r["formula_text"] = frac(feat["formulas"], lambda x: x["formula"])
        cached = [x for x in feat["formulas"] if x["cached"] not in (None, "") and len(x["cached"]) >= 3]
        if cached:
            hits = 0
            for x in cached:
                alts = {x["cached"]}
                try:
                    v = float(x["cached"])
                    for d in (0, 1, 2, 3):
                        alts |= {f"{v:.{d}f}", f"{v:,.{d}f}"}
                        alts |= {a.rstrip("0").rstrip(".") for a in (f"{v:.{d}f}", f"{v:,.{d}f}") if "." in a}
                except ValueError:
                    pass
                hits += any(len(_n(a)) >= 2 and _n(a) in o for a in alts)
            r["formula_value"] = hits / len(cached)
        else:
            r["formula_value"] = None
        if tool == "todoku":  # 窓は計算結果を〈番地＝値〉で示す（数式の字は explain_calculation で別に読む）
            r["formula_marked"] = sum(1 for x in feat["formulas"] if "〈" + x["cell"].split("!")[1] + "＝" in out) / len(feat["formulas"])
        else:
            r["formula_marked"] = r["formula_text"]
    if feat.get("merged"):
        r["merged"] = sum(1 for m in feat["merged"] if m in out) / len(feat["merged"])
    if feat.get("hidden_sheets") or feat.get("hidden_text"):
        marked = ("非表示" in out) if tool == "todoku" else bool(re.search(r"\bhidden\b", out, re.I))
        r["hidden_marked"] = 1.0 if marked else 0.0
        r["hidden_content"] = frac(feat.get("hidden_text") or [])
    if feat.get("display"):
        r["display_format"] = frac(feat["display"], lambda x: x["shown"])
    if feat.get("textboxes"):
        r["textbox"] = frac(feat["textboxes"])
    if feat.get("comments"):
        r["comment"] = frac(feat["comments"])
    if feat.get("checkboxes"):
        n = len(feat["checkboxes"])
        k = len(re.findall(r"[☐☑☒■□]|\[[ xX]\]", out))
        r["checkbox"] = min(1.0, k / n)
    if feat.get("form_values"):
        r["pdf_form_values"] = frac(feat["form_values"])
    if feat.get("acroform"):
        r["pdf_form_fields_listed"] = 1.0 if (tool == "todoku" and "フォームの欄" in out and "［" in out) else 0.0
    if kind in ("xls", "doc", "rtf", "odt", "ods", "jtd"):
        r["legacy_format"] = None  # 内容の再現率で決める（後で）
    return r


def cmd_read(a) -> int:
    work = Path(a.work)
    sample = json.loads((wv.VAL / "sample.json").read_text(encoding="utf-8"))
    log = json.loads((work / "fetch_log.json").read_text(encoding="utf-8"))
    judg = json.loads((wv.VAL / "judgments.json").read_text(encoding="utf-8"))
    cdir = work / "compare"
    cdir.mkdir(exist_ok=True)
    for c in wv.selected(sample, log):
        i = c["id"]
        if a.only and i not in a.only.split(","):
            continue
        outp = cdir / f"{i}.read.json"
        prev = json.loads(outp.read_text(encoding="utf-8")) if outp.exists() else None
        if prev and not (a.force or a.rescore):
            continue
        p = work / "docs" / f"{i}{wv.EXT.get(c['kind'], '')}"
        meas = json.loads((work / "measure" / f"{i}.json").read_text(encoding="utf-8"))
        rec = {"id": i, "kind": c["kind"], "stratum": c["stratum"], "tools": {}}
        feat = wv.last_json(wv.run_watched([a.baseline_python, "-I", "-c", FEATURES, str(p), c["kind"]]))
        rec["features"] = {k: (len(v) if isinstance(v, list) else v) for k, v in feat.items()}
        (cdir / f"{i}.features.json").write_text(json.dumps(feat, ensure_ascii=False, default=str), encoding="utf-8")
        ref_units = [wv.norm(u) for u in json.loads((work / "measure" / f"{i}.ref.txt").read_text(encoding="utf-8"))["units"]] \
            if (work / "measure" / f"{i}.ref.txt").exists() else []
        ref_hay = "".join(ref_units)
        ref_ok = bool(ref_units) and "参照（Latin-1 で読む）のほうが誤り" not in judg.get(i, {}).get("note", "")
        for tool in TOOLS:
            if tool == "todoku":
                md_p = work / "measure" / f"{i}.todoku.md"
                t = meas["todoku"]
                info = {"tool": "todoku", "seconds": t.get("window_seconds", t.get("seconds")), "peak_mb": t.get("peak_mb")}
                if t.get("error"):
                    info["error"] = t["error"]
                out = md_p.read_text(encoding="utf-8") if md_p.exists() and not t.get("error") else None
            else:
                op = cdir / f"{i}.{tool}.txt"
                if a.rescore and prev:  # 道具は動かし直さず、保存した出力で数え直す（時間・メモリは前の記録）
                    info = {k: v for k, v in prev["tools"][tool].items() if k in ("tool", "seconds", "peak_mb", "error", "chars")}
                else:
                    info = wv.last_json(wv.run_watched([a.baseline_python, "-I", "-c", BASE_READ, tool, str(p), c["kind"], str(op)]))
                out = op.read_text(encoding="utf-8") if op.exists() and not info.get("error") else None
            if out is not None:
                s = strip_md(out, tool)
                hay = _n(s)
                if ref_ok:
                    info["char_recall"], _ = wv.char_overlap(ref_hay, hay)
                    info["unit_recall"], _ = wv.unit_recall(ref_units, hay, wv.segments(s))
                    # 適合率：参照の単位は表計算では文字のセルだけなので、表計算では数えない（数の見た目が道具ごとに違う）
                    if c["kind"] not in ("xlsx", "xls", "ods", "zip") and len(ref_hay) < 5_000_000:
                        cs = strip_md(todoku_content(out), tool) if tool == "todoku" else s
                        info["segment_precision"] = segment_precision(cs, ref_hay)
                if ref_ok and len(ref_hay) < 3_000_000:
                    info["bigram_recall"] = bigram_recall(ref_units, s)
                info["conveys"] = conveys(feat, out, tool, c["kind"])
                if "legacy_format" in info["conveys"]:
                    info["conveys"]["legacy_format"] = 1.0 if (info.get("char_recall") or 0) >= 0.9 else 0.0
            else:
                info["conveys"] = {k: 0.0 for k in conveys(feat, "", tool, c["kind"])}
                if "legacy_format" in info["conveys"]:
                    info["conveys"]["legacy_format"] = 0.0
            rec["tools"][tool] = info
        outp.write_text(json.dumps(rec, ensure_ascii=False, default=str), encoding="utf-8")
        print(i, c["kind"], " ".join(f"{t}:{_f(rec['tools'][t].get('char_recall'))}/{_f(rec['tools'][t].get('unit_recall'))}"
                                    f"{'!' + rec['tools'][t]['error'][:30] if rec['tools'][t].get('error') else ''}" for t in TOOLS), flush=True)
    return 0


def _f(x):
    return "-" if x is None else f"{x:.2f}"


# ====================================================================== 書く
TODOKU_WRITE = r"""
import json, sys, time
sys.path.insert(0, sys.argv[1])
from todoku import api
p, out, mode = sys.argv[2], sys.argv[3], sys.argv[4]
o = {}
t = time.time()
try:
    if mode == "pick":
        k, doc, parts = api.window(p)
        fl = api.fields(p)
        names = {}
        if k == "pdf":
            for pg in doc.get("pages", []):
                for f in pg.get("fields", []):
                    names[f["id"]] = {"name": f.get("name"), "source": f.get("source"), "rect": f.get("rect"), "page": pg["number"]}
        cands = [f for f in fl if not f.get("readonly") and f.get("kind") in ("text", "number", "date", None, "")
                 and f.get("value") in (None, "")]
        pref = [f for f in cands if not f.get("agent")] or cands
        # フォームの欄のある PDF：フォームの欄（作り手の印）を選ぶ（記入済みでも上書きする）
        acro = [f for f in fl if names.get(f["id"], {}).get("source") == "acroform" and not f.get("readonly")
                and f.get("kind") in ("text", "number", "date", None, "")]
        if acro:
            pref = acro
        if pref:
            f = pref[0]
            o = {"id": f["id"], "kind": f.get("kind"), "label": (f.get("label") or "")[:60], "agent": bool(f.get("agent")),
                 **names.get(f["id"], {})}
        else:
            o = {"id": None}
        o["kind_doc"] = k
    else:
        fid, val = sys.argv[5], sys.argv[6]
        r = api.fill(p, {fid: val}, out)
        o = {"proof_ok": r.get("proof_ok"), "read_back_ok": r.get("read_back_ok"), "proof": str(r.get("proof"))[:200]}
except Exception as e:
    o["error"] = f"{type(e).__name__}: {str(e)[:200]}"
o["seconds"] = round(time.time() - t, 2)
print(json.dumps(o, ensure_ascii=False, default=str))
"""

BASE_WRITE = r"""
import json, sys, time
tool, p, out, kind, target, val = sys.argv[1:7]
tg = json.loads(target)
o = {}
t = time.time()
try:
    if kind == "xlsx":
        import openpyxl
        wb = openpyxl.load_workbook(p)  # ふだんの書き方：読み込み→代入→保存
        sh, addr = tg["id"].rsplit("!", 1)
        wb[sh][addr.split(":")[0]] = val
        wb.save(out)
    elif kind == "docx":
        import docx, re
        d = docx.Document(p)
        m = re.match(r"T(\d+)\.R(\d+)\.C(\d+)$", tg["id"])
        if not m:
            raise ValueError("python-docx で同じ所を指せない欄（表のマスでない）")
        cell = d.tables[int(m.group(1)) - 1].rows[int(m.group(2)) - 1].cells[int(m.group(3)) - 1]
        o["cell_text_before"] = cell.text[:40]
        cell.text = val
        d.save(out)
    elif kind == "pdf":
        import pypdf
        rd = pypdf.PdfReader(p)
        if not tg.get("name"):
            raise ValueError("フォームの欄がない PDF：pypdf の書き方（update_page_form_field_values）では書けない")
        w = pypdf.PdfWriter(clone_from=rd)
        w.update_page_form_field_values(w.pages[tg.get("page", 1) - 1], {tg["name"]: val}, auto_regenerate=False)
        w.write(out)
    else:
        raise ValueError(f"{kind}：ふだんの道具（openpyxl・python-docx・pypdf）では書けない形式")
except Exception as e:
    o["error"] = f"{type(e).__name__}: {str(e)[:200]}"
o["seconds"] = round(time.time() - t, 2)
print(json.dumps(o, ensure_ascii=False, default=str))
"""

# 書いた後のファイルを元と比べる（とどくでない道具で）
DIFF = r"""
import json, sys, zipfile, re, hashlib, io
orig, new, kind, target, val = sys.argv[1:6]
tg = json.loads(target)
a, b = open(orig, "rb").read(), open(new, "rb").read()
o = {"bytes_before": len(a), "bytes_after": len(b)}
def counts(z):
    c = {}
    names = z.namelist()
    c["parts"] = len(names)
    c["media"] = sum(1 for n in names if "/media/" in n)
    c["vba"] = sum(1 for n in names if n.endswith("vbaProject.bin"))
    c["charts"] = sum(1 for n in names if re.search(r"/charts/chart\d+\.xml$", n))
    c["drawings"] = sum(1 for n in names if re.search(r"/drawings/drawing\d+\.xml$", n))
    c["comments_parts"] = sum(1 for n in names if re.search(r"comments\d*\.xml$", n))
    c["ctrlProps"] = sum(1 for n in names if "/ctrlProps/" in n)
    c["customXml"] = sum(1 for n in names if n.startswith("customXml/"))
    c["printerSettings"] = sum(1 for n in names if "printerSettings" in n)
    c["embeddings"] = sum(1 for n in names if "/embeddings/" in n)
    txt = b"".join(z.read(n) for n in names if n.endswith(".xml") or n.endswith(".rels"))
    for k, pat in {"dataValidation": rb"<(?:\w+:)?dataValidation[ >]", "conditionalFormatting": rb"<(?:\w+:)?conditionalFormatting[ >]",
                   "extLst": rb"<(?:\w+:)?extLst[ >]", "definedName": rb"<(?:\w+:)?definedName[ >]", "sp_shapes": rb"<xdr:sp[ >]",
                   "x14_dataValidation": rb"<x14:dataValidation[ >]", "sdt": rb"<w:sdt[ >]", "txbx": rb"<w:txbxContent[ >]",
                   "fldChar": rb"<w:fldChar[ >]", "sectPr": rb"<w:sectPr[ >]", "mc_AlternateContent": rb"<mc:AlternateContent[ >]",
                   "w14_attrs": rb"w14:paraId", "sheetProtection": rb"<sheetProtection[ >]", "legacyDrawing": rb"<legacyDrawing[ >]",
                   "threadedComments": rb"threadedComment", "pivot": rb"pivotTable", "tableParts": rb"<tablePart[ >]",
                   "cellXfs_xf": rb"<xf[ >]"}.items():
        c[k] = len(re.findall(pat, txt))
    return c
try:
    if kind in ("xlsx", "docx"):
        za, zb = zipfile.ZipFile(io.BytesIO(a)), zipfile.ZipFile(io.BytesIO(b))
        na, nb = set(za.namelist()), set(zb.namelist())
        o["parts_removed"] = sorted(na - nb)[:30]
        o["parts_added"] = sorted(nb - na)[:30]
        o["parts_changed"] = sorted(n for n in na & nb if za.read(n) != zb.read(n))[:40]
        o["parts_same"] = sum(1 for n in na & nb if za.read(n) == zb.read(n))
        ca, cb = counts(za), counts(zb)
        o["lost"] = {k: ca[k] - cb.get(k, 0) for k in ca if ca[k] - cb.get(k, 0) > 0}
        o["gained"] = {k: cb[k] - ca.get(k, 0) for k in cb if cb[k] - ca.get(k, 0) > 0}
        if kind == "xlsx":
            import openpyxl
            sh, addr = tg["id"].rsplit("!", 1)
            got = openpyxl.load_workbook(io.BytesIO(b))[sh][addr.split(":")[0]].value
            o["landed"] = str(got).strip() == val
            o["landed_value"] = str(got)[:40]
        else:
            import docx
            d = docx.Document(io.BytesIO(b))
            m = re.match(r"T(\d+)\.R(\d+)\.C(\d+)$", tg["id"] or "")
            if m:
                cell = d.tables[int(m.group(1)) - 1].rows[int(m.group(2)) - 1].cells[int(m.group(3)) - 1]
                o["landed"] = val in cell.text
            else:
                o["landed"] = val in "\n".join(p.text for p in d.paragraphs)
                o["landed_note"] = "段落の欄：値が本文にあるかだけを確かめた（位置は確かめていない）"
    elif kind == "xls":
        import xlrd
        bk = xlrd.open_workbook(file_contents=b)
        sh, addr = tg["id"].rsplit("!", 1)
        addr = addr.split(":")[0]
        mm = re.match(r"([A-Z]+)(\d+)", addr)
        col = 0
        for ch in mm.group(1):
            col = col * 26 + ord(ch) - 64
        got = bk.sheet_by_name(sh).cell_value(int(mm.group(2)) - 1, col - 1)
        o["landed"] = str(got).strip() == val or (isinstance(got, float) and val.replace(".", "", 1).isdigit() and got == float(val))
        o["prefix_bytes_same"] = None
        o["sheets"] = bk.nsheets
    elif kind == "pdf":
        import pypdf
        o["original_bytes_kept_as_prefix"] = b[:len(a)] == a  # 増分更新なら元のバイト列がそのまま前にある
        ra, rb_ = pypdf.PdfReader(io.BytesIO(a)), pypdf.PdfReader(io.BytesIO(b))
        def info(r):
            root = r.trailer["/Root"]
            af = root.get("/AcroForm")
            af = af.get_object() if af else {}
            fl = r.get_fields() or {}
            return {"pages": len(r.pages), "fields": len(fl), "xfa": af.get("/XFA") is not None if af else False,
                    "sig_fields": sum(1 for v in fl.values() if v.get("/FT") == "/Sig"),
                    "metadata": root.get("/Metadata") is not None, "struct_tree": root.get("/StructTreeRoot") is not None,
                    "outlines": root.get("/Outlines") is not None, "need_appearances": bool(af.get("/NeedAppearances")) if af else False,
                    "objects": len(r.xref_objStm) + sum(len(v) for v in r.xref.values()) if hasattr(r, "xref") else None}
        ia, ib = info(ra), info(rb_)
        o["before"], o["after"] = ia, ib
        o["lost"] = {k: f"{ia[k]} → {ib[k]}" for k in ("xfa", "metadata", "struct_tree", "outlines") if ia[k] and not ib[k]}
        if ia["sig_fields"]:
            o["signature_bytes_preserved"] = o["original_bytes_kept_as_prefix"]
        if tg.get("name"):
            v = (rb_.get_fields() or {}).get(tg["name"], {}).get("/V")
            o["landed"] = str(v) == val
        else:
            # フォームの欄のない PDF：書いた値が、ページの字か、足したフォームの欄（注釈）の値として、選んだ欄の位置にあるか
            for k, v in (rb_.get_fields() or {}).items():
                if str(v.get("/V")) == val:
                    o["landed"] = True
                    o["landed_as"] = "added_form_field"
                    for pg in rb_.pages:
                        for an in pg.get("/Annots") or []:
                            an = an.get_object()
                            av = an.get("/V")
                            if av is None and "/Parent" in an:
                                av = an["/Parent"].get_object().get("/V")
                            if str(av) == val and tg.get("rect"):
                                r = [float(x) for x in an["/Rect"]]
                                cx, cy = (r[0] + r[2]) / 2, (r[1] + r[3]) / 2
                                x0, y0, x1, y1 = tg["rect"]
                                H = float(pg.mediabox.height)  # とどくの rect は上からの座標：PDF の座標（下から）に直す
                                y0, y1 = H - y1, H - y0
                                o["landed_in_rect"] = min(x0, x1) - 3 <= cx <= max(x0, x1) + 3 and min(y0, y1) - 3 <= cy <= max(y0, y1) + 3
            if o.get("landed"):
                raise StopIteration
            import pdfplumber
            with pdfplumber.open(io.BytesIO(b)) as pdf:
                pg = pdf.pages[tg.get("page", 1) - 1]
                words = [w for w in pg.extract_words() if val in w["text"]]
                o["landed"] = bool(words)
                if words and tg.get("rect"):
                    x0, y0, x1, y1 = tg["rect"]
                    w = words[0]
                    H = float(pg.height)
                    # rect は PDF の座標（下から）。値の字の中心が、選んだ欄の四角の中か（少しの余裕）
                    cx, cy = (w["x0"] + w["x1"]) / 2, H - (w["top"] + w["bottom"]) / 2
                    o["landed_in_rect"] = (min(x0, x1) - 3 <= cx <= max(x0, x1) + 30) and (min(y0, y1) - 6 <= cy <= max(y0, y1) + 6)
except StopIteration:
    pass
except Exception as e:
    o["error"] = f"{type(e).__name__}: {str(e)[:200]}"
print(json.dumps(o, ensure_ascii=False, default=str))
"""


def cmd_write(a) -> int:
    work = Path(a.work)
    sample = json.loads((wv.VAL / "sample.json").read_text(encoding="utf-8"))
    log = json.loads((work / "fetch_log.json").read_text(encoding="utf-8"))
    docs = [c for c in wv.selected(sample, log) if c["kind"] in ("xlsx", "xls", "docx", "pdf")]
    if a.extra_ids:  # 標本の外から足す書類（フォームの欄のある PDF など：標本にない種類の書き込みを確かめる）
        recs = {json.loads(x)["id"]: json.loads(x) for x in Path(a.extra_results).read_text(encoding="utf-8").splitlines() if x.strip()}
        for i in a.extra_ids.split(","):
            r = recs[i]
            p = work / "docs" / f"{i}{wv.EXT.get(r['kind'], '')}"
            if not p.exists():
                st, body = wv.get_document(r)
                import hashlib
                if body is None or hashlib.sha256(body).hexdigest() != r.get("sha256"):
                    print("取れない・変わった:", i, st)
                    continue
                p.write_bytes(body)
            docs.append({**r, "stratum": "extra"})
    wdir = work / "write"
    wdir.mkdir(exist_ok=True)
    val_text = "TEST-12345"
    for c in docs:
        i = c["id"]
        if a.only and i not in a.only.split(","):
            continue
        outp = wdir / f"{i}.json"
        if outp.exists() and not a.force:
            continue
        p = work / "docs" / f"{i}{wv.EXT.get(c['kind'], '')}"
        pick = wv.last_json(wv.run_watched([a.todoku_python, "-I", "-c", TODOKU_WRITE, a.todoku_src, str(p), "-", "pick"]))
        rec = {"id": i, "kind": c["kind"], "stratum": c["stratum"], "target": pick, "methods": {}}
        if not pick.get("id"):
            rec["skip"] = pick.get("error") or "書ける欄がない"
            outp.write_text(json.dumps(rec, ensure_ascii=False), encoding="utf-8")
            print(i, "skip", rec["skip"][:60], flush=True)
            continue
        val = "12345" if pick.get("kind") == "number" else val_text
        tgt = json.dumps(pick, ensure_ascii=False)
        for tool in ("todoku", "naive"):
            out = wdir / f"{i}.{tool}{p.suffix}"
            if out.exists():
                out.unlink()
            if tool == "todoku":
                res = wv.last_json(wv.run_watched([a.todoku_python, "-I", "-c", TODOKU_WRITE, a.todoku_src, str(p), str(out), "fill",
                                                   pick["id"], val]))
            else:
                res = wv.last_json(wv.run_watched([a.baseline_python, "-I", "-c", BASE_WRITE, "naive", str(p), str(out), c["kind"],
                                                   tgt, val]))
            if out.exists() and not res.get("error"):
                res["diff"] = wv.last_json(wv.run_watched([a.baseline_python, "-I", "-c", DIFF, str(p), str(out), c["kind"], tgt, val]))
            rec["methods"][tool] = res
        outp.write_text(json.dumps(rec, ensure_ascii=False, default=str), encoding="utf-8")
        print(i, c["kind"], pick.get("id"), " | ".join(
            f"{t}: {'ERR ' + r['error'][:50] if r.get('error') else ('landed' if r.get('diff', {}).get('landed') else 'NOT landed')}"
            f" lost={r.get('diff', {}).get('lost')}" for t, r in rec["methods"].items()), flush=True)
    return 0


# ====================================================================== まとめ
def cmd_summarize(a) -> int:
    work = Path(a.work)
    reads = [json.loads(p.read_text(encoding="utf-8")) for p in sorted((work / "compare").glob("*.read.json"))]
    writes = [json.loads(p.read_text(encoding="utf-8")) for p in sorted((work / "write").glob("*.json"))]
    # 書類の中の字（欄の見出し・書く前のマスの字・読み戻した値）は、リポジトリに残さない（個人の情報が入りうる）
    for w in writes:
        (w.get("target") or {}).pop("label", None)
        for r in (w.get("methods") or {}).values():
            r.pop("cell_text_before", None)
            (r.get("diff") or {}).pop("landed_value", None)
    out = {"read": [], "write": writes}
    agg: dict = {t: collections.defaultdict(list) for t in TOOLS}
    feat: dict = {t: collections.defaultdict(list) for t in TOOLS}
    for r in reads:
        row = {"id": r["id"], "kind": r["kind"], "stratum": r["stratum"], "features": r["features"], "tools": {}}
        for t in TOOLS:
            x = r["tools"][t]
            row["tools"][t] = {k: x.get(k) for k in ("char_recall", "unit_recall", "bigram_recall", "segment_precision", "seconds",
                                                     "peak_mb", "error",
                                                     "conveys")}
            ok = not x.get("error")
            agg[t]["ok"].append(ok)
            agg[t]["seconds"].append(x.get("seconds") or 0)
            for k in ("char_recall", "unit_recall", "bigram_recall", "segment_precision"):
                if x.get(k) is not None:
                    agg[t][k].append(x[k])
                elif k != "segment_precision" and any(r["tools"][u].get(k) is not None for u in TOOLS):
                    agg[t][k].append(0.0)  # 読めなかった（ほかの道具は読めた）：0 と数える
            for k, v in (x.get("conveys") or {}).items():
                if v is not None:
                    feat[t][k].append(v)
        out["read"].append(row)

    def mean(xs):
        return round(sum(xs) / len(xs), 4) if xs else None

    def med(xs):
        xs = sorted(xs)
        return xs[len(xs) // 2] if xs else None

    summ = {"n_read": len(reads), "tools": {}}
    for t in TOOLS:
        g = agg[t]
        summ["tools"][t] = {"opened": sum(g["ok"]), "char_recall_mean": mean(g["char_recall"]), "unit_recall_mean": mean(g["unit_recall"]),
                            "bigram_recall_mean": mean(g["bigram_recall"]),
                            "docs_bigram_recall_lt_0.995": sum(1 for x in g["bigram_recall"] if x < 0.995),
                            "docs_unit_recall_ge_0.99": sum(1 for x in g["unit_recall"] if x >= 0.99),
                            "n_recall": len(g["unit_recall"]),
                            "segment_precision_mean": mean(g["segment_precision"]),
                            "seconds_median": med(g["seconds"]), "seconds_total": round(sum(g["seconds"]), 1),
                            "features": {k: {"docs": len(v), "conveyed_mean": mean(v), "docs_fully": sum(1 for x in v if x >= 0.99)}
                                         for k, v in sorted(feat[t].items())}}
    # 書く
    ws: dict = {}
    for w in writes:
        if w.get("skip"):
            continue
        for t, r in w["methods"].items():
            s = ws.setdefault(f"{w['kind']}:{t}", collections.Counter())
            s["docs"] += 1
            if r.get("error"):
                s["error"] += 1
                continue
            d = r.get("diff", {})
            s["landed"] += bool(d.get("landed"))
            if "landed_in_rect" in d:
                s["landed_in_rect"] += bool(d.get("landed_in_rect"))
            s["lost_something"] += bool(d.get("lost") or d.get("parts_removed"))
            for k in (d.get("lost") or {}):
                s["lost:" + k] += 1
            if d.get("parts_removed"):
                s["parts_removed_docs"] += 1
            if d.get("original_bytes_kept_as_prefix") is not None:
                s["pdf_original_bytes_kept"] += bool(d["original_bytes_kept_as_prefix"])
            if t == "todoku":
                s["proof_ok"] += bool(r.get("proof_ok"))
                s["read_back_ok"] += bool(r.get("read_back_ok"))
            s["parts_changed_total"] += len(d.get("parts_changed") or [])
            s["parts_same_total"] += d.get("parts_same") or 0
    summ["write"] = {k: dict(v) for k, v in sorted(ws.items())}
    summ["write_skipped"] = collections.Counter(w.get("skip", "")[:40] for w in writes if w.get("skip"))
    (wv.VAL / "compare.json").write_text(json.dumps(out, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    old = json.loads((wv.VAL / "summary.json").read_text(encoding="utf-8")) if (wv.VAL / "summary.json").exists() else {}
    old["compare"] = summ
    (wv.VAL / "summary.json").write_text(json.dumps(old, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    print(json.dumps(summ, ensure_ascii=False, indent=1, default=str)[:6000])
    return 0


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    for name in ("read", "write", "summarize"):
        p = sub.add_parser(name)
        p.add_argument("--work", required=True)
        p.add_argument("--out", default=str(wv.VAL))
        p.add_argument("--todoku-python", default=os.environ.get("TODOKU_PYTHON", sys.executable))
        p.add_argument("--baseline-python", default=os.environ.get("BASELINE_PYTHON", sys.executable))
        p.add_argument("--todoku-src", default=str(ROOT / "src"))
        p.add_argument("--force", action="store_true")
        p.add_argument("--rescore", action="store_true", help="read：道具を動かし直さず、保存した出力で数え直す")
        p.add_argument("--only", default="")
        p.add_argument("--extra-results", default="", help="標本の外から足す書類の結果ファイル（write）")
        p.add_argument("--extra-ids", default="", help="足す書類の id（カンマ区切り）")
    a = ap.parse_args(argv)
    wv.VAL = Path(a.out).resolve()
    return {"read": cmd_read, "write": cmd_write, "summarize": cmd_summarize}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
