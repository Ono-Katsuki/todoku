"""世界調査の段階を、とどくとは別の読み方で確かめる（段階の検証）。

とどくの段階（読む S〜F・書く A〜D）は、とどく自身が付けたもの。ここでは、世界調査の結果（world/results.jsonl）から
層別に無作為に選んだ書類を取り直し、確立した別の読み方（openpyxl・xlrd・pdfplumber・pypdf・python-docx・
文字コードを決めて読むだけ・macOS の textutil）で「参照の読み」を作って、とどくの窓と比べる。
食い違った所は、確かめる人（AI エージェント）が両方を見て、どちらが正しいかを決める（参照の道具が正しいとは限らない）。

段階：
  sample   層（群 × 読む段階）ごとに、種を決めて並べ替えた候補の列を作る（world/validation/sample.json）
  fetch    候補を順に取り直す（robots.txt を守り、同じ相手には 2 秒以上の間、1 本ずつ）。SHA-256 が同じものだけを使い、
           変わった・消えたものは記録して、同じ層の次の候補に替える
  measure  1 件ずつ、とどくの窓と参照の読みを作って比べる（別の処理で、2 GB・120 秒を超えたら止める）
  judge    確かめた人の判断（world/validation/judgments.json）と測った値を合わせて、grade_check.json と summary.json を作る

取り直した書類は作業用のフォルダ（--work、リポジトリの外）だけに置き、リポジトリには URL・SHA-256・数・短い説明だけを残す。
書類は信頼しないデータとして扱う：マクロは動かさず、読む Python は -I で動かす。

書類の取得：結果の行に Common Crawl の記録の場所（filename・offset・length、または warc_filename・warc_record_offset・
warc_record_length）があれば、https://data.commoncrawl.org/<filename> を範囲指定（Range）で取り、WARC の記録から本文を出す。
なければ元の URL から取る（どちらも robots.txt を確かめ、同じ相手には 3 秒の間、1 本ずつ）。

使い方（--out の既定は world/validation）:
  python scripts/world_validate.py sample --results <結果.jsonl> [--out DIR] [--seed N] [--cap N]
  python scripts/world_validate.py fetch --work /path/to/scratch [--out DIR]
  python scripts/world_validate.py measure --work /path/to/scratch --todoku-python PY --baseline-python PY
  python scripts/world_validate.py judge --work /path/to/scratch
"""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import math
import os
import random
import re
import subprocess
import sys
import time
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VAL = ROOT / "world" / "validation"  # --out で替えられる（main で置き換える）
SEED = 20261104
CAP = 8  # 層ごとの件数の上限（小さな層は全部）
FAMILIES = ("pdf", "sheet", "text", "csv")
GRADES = "SABCDEF"
EXT = {"csv": ".csv", "zip": ".zip", "xlsx": ".xlsx", "xls": ".xls", "ods": ".ods", "pdf": ".pdf", "docx": ".docx",
       "doc": ".doc", "odt": ".odt", "rtf": ".rtf", "jtd": ".jtd"}
MEM_LIMIT_MB = 2048
TIME_LIMIT_S = 120
DELAY = 3.0  # 同じ相手への間（調査の 2 秒より長く：調査が同時に動いていても、相手への負担を増やしすぎない）


# ====================================================================== 見張り（1 件ごとに別の処理）
def _rss_mb(pid: int) -> float:
    try:
        out = subprocess.run(["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True, timeout=5).stdout.strip()
        return int(out) / 1024 if out else 0.0
    except (ValueError, subprocess.SubprocessError):
        return 0.0


def run_watched(cmd: list[str], limit_s: float = TIME_LIMIT_S, mem_mb: float = MEM_LIMIT_MB, env: dict | None = None,
                cwd: str | None = None) -> dict:
    """優先度を下げて動かし、メモリと時間を見張る。超えたらその処理だけ止める。"""
    t = time.time()
    proc = subprocess.Popen(["nice", "-n", "10", *cmd], stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env, cwd=cwd)
    why, peak, out_b, err_b = None, 0.0, b"", b""
    while True:
        try:
            out_b, err_b = proc.communicate(timeout=0.5)
            break
        except subprocess.TimeoutExpired:
            peak = max(peak, _rss_mb(proc.pid))
            if peak > mem_mb:
                why = "memory"
            elif time.time() - t > limit_s:
                why = "timeout"
            if why:
                proc.kill()
                proc.communicate()
                break
    return {"why": why, "peak_mb": round(peak), "seconds": round(time.time() - t, 2), "code": proc.returncode,
            "stdout": out_b.decode("utf-8", "replace"), "stderr": err_b.decode("utf-8", "replace")[-1500:]}


def last_json(r: dict) -> dict:
    if r["why"]:
        return {"error": r["why"], "peak_mb": r["peak_mb"], "seconds": r["seconds"]}
    lines = r["stdout"].strip().splitlines()
    try:
        out = json.loads(lines[-1]) if lines else {"error": "crash", "stderr": r["stderr"][-600:]}
    except json.JSONDecodeError:
        out = {"error": "crash", "stderr": r["stderr"][-600:]}
    out.setdefault("seconds", r["seconds"])
    out["peak_mb"] = r["peak_mb"]
    return out


# ====================================================================== 標本
def cmd_sample(a) -> int:
    src = Path(a.results)
    raw = src.read_bytes()
    recs = [json.loads(x) for x in raw.decode("utf-8").splitlines() if x.strip()]
    graded = [r for r in recs if r.get("reach") == "ok" and r.get("read") in list(GRADES)]
    strata: dict[str, list] = collections.defaultdict(list)
    for r in graded:
        strata[f"{r['family']}:{r['read']}"].append(r)
    out = {"seed": a.seed, "cap": a.cap, "frame": {"file": str(src.name), "sha256": hashlib.sha256(raw).hexdigest(),
                                                   "records": len(recs), "graded": len(graded)},
           "design": "層 = 調査の群（pdf/sheet/text/csv）× 調査の読む段階。層ごとに種 f'{seed}:{層}' で候補を並べ替え、"
                     "前から取り直して、SHA-256 が調査のときと同じものを上限（cap）まで使う。重み = 層の書類数 ÷ 使えた数。",
           "absent_strata": [f"{fam}:{g}" for fam in FAMILIES for g in GRADES if f"{fam}:{g}" not in strata],
           "strata": {}}
    for key in sorted(strata):
        rows = sorted(strata[key], key=lambda r: r["id"])
        random.Random(f"{a.seed}:{key}").shuffle(rows)
        out["strata"][key] = {"N": len(rows), "candidates": [
            {k: r.get(k) for k in ("id", "country", "family", "kind", "url", "catalog_url", "resolve", "sha256", "bytes",
                                   "read", "write", "read_reasons", "write_reasons", "fields", "writable_fields", "pages", "chars",
                                   *WARC_KEYS)}
            for r in rows[: max(a.cap * 4, a.cap + 12)]]}
    VAL.mkdir(parents=True, exist_ok=True)
    (VAL / "sample.json").write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"層 {len(out['strata'])}・ない層 {out['absent_strata']}")
    for k, v in out["strata"].items():
        print(f"  {k}: N={v['N']}")
    return 0


# ====================================================================== 取り直し
# 調査と同じ名乗り（User-Agent）と、RFC 9309 の robots.txt の読み（urllib.robotparser）。調査の道具が替わっても動くよう、ここに持つ
UA = "todoku-research/0.2 (accessibility research on public documents; polite crawler; https://github.com/Ono-Katsuki/todoku)"
MAX_BYTES = 25 * 1024 * 1024
CC_BASE = "https://data.commoncrawl.org/"
# Common Crawl の記録の場所：結果の行にこの欄があれば、元の URL ではなく WARC の記録を範囲指定で取る
WARC_KEYS = ("filename", "offset", "length", "warc_filename", "warc_record_offset", "warc_record_length",
             "warc_offset", "warc_length")
_last: dict[str, float] = {}
_robots: dict[str, object] = {}


def _safe_url(url: str) -> str:
    import urllib.parse

    return urllib.parse.quote(url.strip(), safe=":/?#[]@!$&'()*+,;=%~")


def _wait(host: str) -> None:
    at = _last.get(host, 0) + DELAY
    now = time.time()
    if at > now:
        time.sleep(at - now)
    _last[host] = time.time()


def _get(url: str, headers: dict | None = None, limit: int = MAX_BYTES) -> dict:
    import urllib.error
    import urllib.parse
    import urllib.request

    url = _safe_url(url)
    _wait(urllib.parse.urlparse(url).netloc)
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*", **(headers or {})})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            body = r.read(limit + 1)
            return {"status": r.status, "body": body, "too_large": len(body) > limit}
    except urllib.error.HTTPError as e:
        return {"status": e.code, "body": b""}
    except Exception as e:  # noqa: BLE001
        return {"status": 0, "error": f"{type(e).__name__}: {str(e)[:120]}", "body": b""}


def allowed(url: str) -> bool:
    """robots.txt（RFC 9309）が、この名乗りでの取得を許すか。robots.txt がない（4xx）・取れないときは、禁止の指定はない。"""
    import urllib.parse
    import urllib.robotparser

    url = _safe_url(url)
    pu = urllib.parse.urlparse(url)
    base = f"{pu.scheme}://{pu.netloc}"
    if base not in _robots:
        r = _get(base + "/robots.txt", limit=500_000)
        rp = None
        if r.get("status") == 200:
            rp = urllib.robotparser.RobotFileParser()
            rp.parse(r["body"].decode("utf-8", "replace").splitlines())
        _robots[base] = rp
    rp = _robots[base]
    return True if rp is None else rp.can_fetch(UA, url)


def warc_where(c: dict) -> tuple[str, int, int] | None:
    fn = c.get("warc_filename") or c.get("filename")
    off = next((c.get(k) for k in ("warc_record_offset", "warc_offset", "offset") if c.get(k) is not None), None)
    ln = next((c.get(k) for k in ("warc_record_length", "warc_length", "length") if c.get(k) is not None), None)
    if fn and off is not None and ln is not None:
        return str(fn), int(off), int(ln)
    return None


def warc_payload(raw: bytes) -> tuple[int, bytes]:
    """WARC の記録（gzip の1メンバー）から、HTTP の応答の状態と本文を取り出す（転送の符号化 chunked は解く）。"""
    import gzip

    rec = gzip.decompress(raw)
    head, _, rest = rec.partition(b"\r\n\r\n")  # WARC の見出し
    http_head, _, body = rest.partition(b"\r\n\r\n")  # HTTP の見出し
    lines = http_head.split(b"\r\n")
    try:
        status = int(lines[0].split()[1])
    except (IndexError, ValueError):
        status = 0
    hdr = {ln.split(b":", 1)[0].strip().lower(): ln.split(b":", 1)[1].strip() for ln in lines[1:] if b":" in ln}
    m = re.search(rb"Content-Length:\s*(\d+)", head, re.I)
    if m:  # WARC の記録の長さ（HTTP の見出しを含む）で、本文の終わりを決める（末尾の \r\n\r\n を除く）
        total = int(m.group(1))
        body = rest[:total][len(http_head) + 4:]
    if b"chunked" in hdr.get(b"transfer-encoding", b"").lower():
        out, i = b"", 0
        while True:
            j = body.find(b"\r\n", i)
            if j < 0:
                break
            n = int(body[i:j].split(b";")[0] or b"0", 16)
            if n == 0:
                break
            out += body[j + 2:j + 2 + n]
            i = j + 2 + n + 2
        body = out
    return status, body


def get_document(c: dict) -> tuple[str, bytes | None]:
    """(状態, 中身)。WARC の場所があれば Common Crawl から範囲指定で、なければ元の URL から（robots.txt を確かめて）。"""
    w = warc_where(c)
    if w:
        fn, off, ln = w
        url = CC_BASE + fn.lstrip("/")
        if not allowed(url):
            return "robots", None
        r = _get(url, headers={"Range": f"bytes={off}-{off + ln - 1}"}, limit=ln + 1024)
        if r.get("status") not in (200, 206):
            return (f"http_{r.get('status')}" if r.get("status") else "error"), None
        try:
            st, body = warc_payload(r["body"])
        except Exception as e:  # noqa: BLE001
            return f"warc_error:{type(e).__name__}", None
        return ("ok" if st == 200 else f"http_{st}"), body
    url = c["url"]
    if not allowed(url):
        return "robots", None
    r = _get(url)
    if r.get("status") != 200 and c.get("resolve") == "datagovsg" and c.get("catalog_url") and allowed(c["catalog_url"]):
        try:  # 目録が問い合わせの答えで URL を示す型（data.gov.sg）：一時的な URL が切れていれば、問い合わせ直す
            u = ((json.loads(_get(c["catalog_url"])["body"]) or {}).get("data") or {}).get("url")
        except (json.JSONDecodeError, KeyError, TypeError):
            u = None
        if str(u or "").startswith("https://") and allowed(u):
            r = _get(u)
    if r.get("status") != 200:
        return (f"http_{r.get('status')}" if r.get("status") else "error"), None
    if r.get("too_large"):
        return "too_large", None
    return "ok", r["body"]


def cmd_fetch(a) -> int:
    work = Path(a.work)
    docs = work / "docs"
    docs.mkdir(parents=True, exist_ok=True)
    sample = json.loads((VAL / "sample.json").read_text(encoding="utf-8"))
    log_path = work / "fetch_log.json"
    log = json.loads(log_path.read_text(encoding="utf-8")) if log_path.exists() else {}
    for key, st in sample["strata"].items():
        used = sum(1 for c in st["candidates"] if log.get(c["id"], {}).get("status") == "same")
        for c in st["candidates"]:
            if used >= sample["cap"]:
                break
            if c["id"] in log:
                continue
            rec = {"t": time.strftime("%Y-%m-%dT%H:%M:%S"), "via": "warc" if warc_where(c) else "url"}
            status, body = get_document(c)
            rec["status"] = status
            if body is not None:
                h = hashlib.sha256(body).hexdigest()
                rec["sha256_now"] = h
                if not c.get("sha256") or h == c["sha256"]:
                    rec["status"] = "same"
                    (docs / f"{c['id']}{EXT.get(c.get('kind') or '', '')}").write_bytes(body)
                    used += 1
                else:
                    rec["status"] = "changed"
            log[c["id"]] = rec
            log_path.write_text(json.dumps(log, ensure_ascii=False, indent=1), encoding="utf-8")
            print(f"[{rec['status']}] {key} {c['id']} {c['url'][:90]}", flush=True)
        print(f"== {key}: {used}/{min(sample['cap'], st['N'])}", flush=True)
    return 0


def selected(sample: dict, log: dict) -> list[dict]:
    out = []
    for key, st in sample["strata"].items():
        n = 0
        for c in st["candidates"]:
            if n >= sample["cap"]:
                break
            if log.get(c["id"], {}).get("status") == "same":
                out.append({**c, "stratum": key})
                n += 1
    return out


# ====================================================================== 比べる（正規化）
def norm(s: str) -> str:
    """NFKC にして、空白（改行・タブ・全角の空白を含む）をすべて除く。"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", s or ""))


def char_overlap(ref: str, got: str) -> tuple[float | None, float | None]:
    """字の多重集合での再現率・適合率（並びは見ない）。"""
    a, b = collections.Counter(ref), collections.Counter(got)
    inter = sum((a & b).values())
    return (inter / sum(a.values()) if a else None), (inter / sum(b.values()) if b else None)


def unit_recall(units: list[str], hay: str, segs: set | None = None, max_units: int = 4000,
                max_scan: int = 150) -> tuple[float | None, list[str]]:
    """参照の単位（セルの値・行）が、並びを保ったまま（部分文字列として）出ている割合（字数で重み）。見つからない例も返す。
    大きな書類は、単位を種を決めて max_units 個に間引き、区切り（｜・|・改行）で切った断片の集合で先に引き当て、
    当たらないものだけを、max_scan 個まで部分文字列として探す（それより先は、断片の集合だけで判定）。"""
    units = [u for u in units if u]
    if len(hay) >= 2_000_000:  # 大きな出力：単位を 300 個に間引き、すべて部分文字列として探す（区切りの違いに左右されない）
        max_units, max_scan = 300, 300
    sampled = len(units) > max_units
    if sampled:
        units = random.Random(SEED).sample(units, max_units)
    tot = hit = scanned = 0
    miss = []
    for u in units:
        tot += len(u)
        ok = segs is not None and u in segs
        if not ok and (scanned < max_scan or len(hay) < 2_000_000):  # 小さな出力は全部探す
            scanned += 1
            ok = u in hay
        if ok:
            hit += len(u)
        elif len(miss) < 30:
            miss.append(u[:80])
    return (hit / tot if tot else None), miss


def segments(text: str) -> set:
    """出力を区切り（表の縦線・改行・タブ）で切り、記号を除いた断片の集合。"""
    t = unicodedata.normalize("NFKC", text)
    out = set()
    for seg in re.split(r"[|\n\t]", t):
        seg = re.sub(r"^\s*(?:[-*>]\s+|#+\s+)", "", seg)
        seg = re.sub(r"\s+", "", seg)
        seg = re.sub(r"^\d+(?:〜\d+)?:", "", seg)
        if seg:
            out.add(seg)
            m = re.match(r"^[［〈\[<]?[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?[＝=](.*?)[］〉\]>]?$", seg)
            if m:
                out.add(m.group(1))
    return out


_ID = (r"(?:▶)?(?:p\d+\.[\w.]+|[A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?(?:〜[A-Z]{1,3}\d+)?|T\d+(?:\.[A-Z]?\d+)*|P\d+(?:\.\w+)*|E\d+[^］〉]*"
       r"|[^［］!]{1,40}![A-Z]{1,3}\d+(?::[A-Z]{1,3}\d+)?|[→↓↑←][^］]*)？?")


def strip_todoku(md: str) -> str:
    """とどくの窓の記号（欄の名前・番地・余白の印・区切りの縦線）を除き、値だけを残す（参照と比べるため）。
    書類の字の「［　］」などは、欄の名前の形でなければ残す。冒頭の説明（SHA-256・記号の説明）も除く。"""
    out = []
    for ln in md.splitlines():
        if ln.startswith(("- 元ファイルの SHA-256", "- この文書は読むための窓", "- 記号:", "- 欄として示していない所", "- 見出しは配置から")):
            continue
        out.append(ln)
    t = "\n".join(out)
    t = re.sub(r"［" + _ID + r"＝(.*?)］", r"\1", t)
    t = re.sub(r"〈" + _ID + r"＝(.*?)〉", r"\1", t)
    t = re.sub(r"［" + _ID + r"］", "", t)
    t = re.sub(r"^\s*-\s+\d+(?:〜\d+)?:\s", "", t, flags=re.M)
    return t.replace("｜", " ").replace("↵", "\n")  # ↵ はセルの中の改行の印


# ====================================================================== 測る
TODOKU_WORKER = r"""
import json, sys, time, hashlib
sys.path.insert(0, sys.argv[1])
from todoku import api
p, out_md = sys.argv[2], sys.argv[3]
t = time.time()
o = {}
try:
    k, doc, parts = api.window(p)
    md = "\n\n".join(b for _, b in parts)
    if k == "zip":  # ZIP の窓は中のファイルの一覧：中のファイル（前から 50 個まで）の窓を続ける（参照も同じ 50 個）
        o["members"] = []
        for mem in doc.get("members", [])[:50]:
            if not mem.get("readable"):
                o["members"].append({"name": mem["name"][-80:], "readable": False})
                continue
            try:
                k2, d2, p2 = api.window(p + "!/" + mem["name"])
                md += "\n\n" + "\n\n".join(b for _, b in p2)
                o["members"].append({"name": mem["name"][-80:], "kind": k2, "readable": True})
            except Exception as e:
                o["members"].append({"name": mem["name"][-80:], "error": f"{type(e).__name__}: {e}"[:120]})
    open(out_md, "w", encoding="utf-8").write(md)
    o["kind"] = k
    o["window_seconds"] = round(time.time() - t, 2)
    o["parts"] = len(parts)
    o["window_chars"] = len(md)
    try:
        g = api.grade(p)
        o["read"] = g["grade"]; o["write"] = g["write"]["grade"]
        o["read_reasons"] = [(r["grade"], r["title"], r.get("where", [])[:6]) for r in g["reasons"]]
        o["write_reasons"] = [(r["grade"], r["title"]) for r in g["write"]["reasons"]]
    except Exception as e:
        o["grade_error"] = f"{type(e).__name__}: {e}"[:200]
    try:
        fl = api.fields(p)
        o["fields"] = [{"id": f["id"], "kind": f.get("kind"), "agent": bool(f.get("agent")), "readonly": bool(f.get("readonly")),
                        "has_value": f.get("value") not in (None, "", False)} for f in fl][:3000]
        o["n_fields"] = len(fl)
    except Exception as e:
        o["fields_error"] = f"{type(e).__name__}: {e}"[:200]
    def sheet_info(doc):
        sh_out = []
        for sh in doc.get("sheets", []):
            cells = {}
            formulas = {}
            for c in sh.get("cells", []):
                cells[c["cell"]] = c.get("value")
                if c.get("formula"):
                    formulas[c["cell"]] = c["formula"]
            sh_out.append({"name": sh["name"], "state": sh.get("state"), "hidden_rows": len(sh.get("hidden_rows") or []),
                           "hidden_cols": len(sh.get("hidden_cols") or []), "n_cells": len(cells),
                           "merged": sorted({c.get("merged") for c in sh.get("cells", []) if c.get("merged")}),
                           "n_validations": len(sh.get("validations") or []), "protected": sh.get("protected"),
                           "n_shapes": len(sh.get("shapes") or []), "n_notes": len(sh.get("notes") or []),
                           "n_undecided": len(sh.get("undecided") or []), "cells": cells, "formulas": formulas,
                           "cell_keys": sorted({kk for c in sh.get("cells", []) for kk in c})})
        return sh_out
    if k in ("xlsx", "xls", "ods"):
        o["sheets"] = sheet_info(doc)
        o["macros"] = bool(doc.get("macros"))
    if k == "pdf":
        o["form"] = doc.get("form")
        o["pages"] = [{"n": pg["number"], "image_only": bool(pg.get("image_only")), "unmapped": pg.get("unmapped") or 0,
                       "n_fields": len(pg.get("fields") or []),
                       "sources": dict(__import__("collections").Counter(f.get("source", "?") for f in pg.get("fields") or [])),
                       "n_agent": sum(1 for f in pg.get("fields") or [] if f.get("agent"))} for pg in doc.get("pages", [])]
    if k in ("docx", "doc", "odt", "rtf", "jtd"):
        o["n_doc_fields"] = len(doc.get("fields") or [])
        o["field_sources"] = dict(__import__("collections").Counter(f.get("source", "?") for f in doc.get("fields") or []))
        o["macros"] = bool(doc.get("macros"))
        o["n_pictures"] = len(doc.get("pictures") or [])
except Exception as e:
    o["error"] = f"{type(e).__name__}: {e}"[:300]
o["seconds"] = round(time.time() - t, 2)
print(json.dumps(o, ensure_ascii=False, default=str))
"""


def todoku_measure(a, path: Path, out_md: Path) -> dict:
    r = run_watched([a.todoku_python, "-I", "-c", TODOKU_WORKER, a.todoku_src, str(path), str(out_md)])
    return last_json(r)


def ref_measure(a, path: Path, kind: str, out_txt: Path) -> dict:
    extra = []
    tu = None
    if kind in ("doc", "rtf"):  # 古い Word・RTF：macOS の textutil で字にする（とどくとは別の読み方）
        tu = out_txt.with_suffix(".textutil.txt")
        r0 = run_watched(["textutil", "-convert", "txt", "-encoding", "UTF-8", "-output", str(tu), str(path)])
        if r0["why"] or not tu.exists():
            return {"error": f"textutil:{r0['why'] or 'failed'}", "seconds": r0["seconds"]}
        extra = ["--textutil", str(tu)]
    cmd = [a.baseline_python, "-I", str(ROOT / "scripts" / "world_reference.py"), str(path), kind, str(out_txt), *extra]
    out = last_json(run_watched(cmd))
    if kind == "pdf" and out.get("error") in ("timeout", "memory"):
        out2 = last_json(run_watched(cmd + ["--pdf-fast"]))  # pdfplumber が重すぎる：pypdf の字で
        out2["pdfplumber_error"] = out["error"]
        out = out2
    return out


def cmd_measure(a) -> int:
    work = Path(a.work)
    sample = json.loads((VAL / "sample.json").read_text(encoding="utf-8"))
    log = json.loads((work / "fetch_log.json").read_text(encoding="utf-8"))
    mdir = work / "measure"
    mdir.mkdir(exist_ok=True)
    for c in selected(sample, log):
        out = mdir / f"{c['id']}.json"
        if (a.only and c["id"] not in a.only.split(",")) or (out.exists() and not a.force and not a.only):
            continue
        p = work / "docs" / f"{c['id']}{EXT.get(c['kind'], '')}"
        tod = todoku_measure(a, p, mdir / f"{c['id']}.todoku.md")
        ref = ref_measure(a, p, c["kind"], mdir / f"{c['id']}.ref.txt")
        m = compare(c, tod, ref, mdir)
        out.write_text(json.dumps({"id": c["id"], "stratum": c["stratum"], "todoku": tod, "ref": ref, "metrics": m},
                                  ensure_ascii=False, default=str), encoding="utf-8")
        print(f"{c['id']:16} {c['stratum']:8} {c['kind']:5} recorded {c['read']}/{c['write']} now {tod.get('read')}/{tod.get('write')} "
              f"char_r={_f(m.get('char_recall'))} unit_r={_f(m.get('unit_recall'))} cell={_f(m.get('cell_agree'))} "
              f"{tod.get('error', '')}{ref.get('error', '')}", flush=True)
    return 0


def cmd_recompare(a) -> int:
    """保存した出力から、比べた値だけを計算し直す（処理は動かさない）。"""
    work = Path(a.work)
    sample = json.loads((VAL / "sample.json").read_text(encoding="utf-8"))
    log = json.loads((work / "fetch_log.json").read_text(encoding="utf-8"))
    mdir = work / "measure"
    for c in selected(sample, log):
        out = mdir / f"{c['id']}.json"
        if not out.exists():
            continue
        d = json.loads(out.read_text(encoding="utf-8"))
        d["metrics"] = compare(c, d["todoku"], d["ref"], mdir)
        out.write_text(json.dumps(d, ensure_ascii=False, default=str), encoding="utf-8")
        m = d["metrics"]
        print(f"{c['id']:16} {c['stratum']:8} {c['kind']:5} rec {c['read']}/{c['write']} now {d['todoku'].get('read')}/{d['todoku'].get('write')} "
              f"char_r={_f(m.get('char_recall'))} unit_r={_f(m.get('unit_recall'))} prec={_f(m.get('char_precision_raw'))} "
              f"cell={_f(m.get('cell_agree'))} {d['todoku'].get('error', '')}{d['ref'].get('error', '')}", flush=True)
    return 0


def _f(x):
    return "-" if x is None else f"{x:.3f}"


def compare(c: dict, tod: dict, ref: dict, mdir: Path) -> dict:
    m: dict = {}
    md_p, ref_p = mdir / f"{c['id']}.todoku.md", mdir / f"{c['id']}.ref.txt"
    if tod.get("error") or ref.get("error") or not md_p.exists() or not ref_p.exists():
        return {"error": tod.get("error") or ref.get("error") or "missing"}
    md = strip_todoku(md_p.read_text(encoding="utf-8"))
    hay = norm(md)
    units = [norm(u) for u in json.loads(ref_p.read_text(encoding="utf-8"))["units"]]
    refall = "".join(units)
    m["ref_chars"] = len(refall)
    m["char_recall"], _ = char_overlap(refall, hay)
    m["unit_recall"], m["missing_units"] = unit_recall(units, hay, segments(md))
    m["units_sampled"] = len([u for u in units if u]) > 4000
    m["char_precision_raw"] = char_overlap(refall, hay)[1]
    # 表計算：番地ごとの値（参照の openpyxl/xlrd の値と、窓の値）
    if ref.get("sheets") and tod.get("sheets"):
        m.update(cell_compare(ref["sheets"], tod["sheets"]))
        m["compressed"] = compressed_rows(md_p.read_text(encoding="utf-8"), ref["sheets"])
    if ref.get("pages") is not None and tod.get("pages") is not None:
        ref = {**ref, "pdf": ref}
        rp = ref["pdf"]["pages"]
        m["pdf_pages"] = {"ref": len(rp), "todoku": len(tod["pages"])}
        m["image_only_pages"] = {"ref_no_text_with_image": [p["n"] for p in rp if p["chars"] == 0 and p["images"] > 0],
                                 "ref_no_text": [p["n"] for p in rp if p["chars"] == 0],
                                 "todoku": [p["n"] for p in tod["pages"] if p["image_only"]]}
        m["cid_chars_ref"] = sum(p.get("cid", 0) for p in rp)
        m["unmapped_todoku"] = sum(p["unmapped"] for p in tod["pages"])
        m["acroform"] = {"ref_fields": ref["pdf"].get("acroform_fields"), "ref_widgets": ref["pdf"].get("widgets"),
                         "todoku_form": sum(p["sources"].get("form", 0) for p in tod["pages"]),
                         "todoku_layout": sum(p["sources"].get("layout", 0) for p in tod["pages"]),
                         "todoku_agent": sum(p["n_agent"] for p in tod["pages"])}
    return m


_COMP = re.compile(r"^- (\d+)〜(\d+): (\d+)(?:〜\d+)?行目と同じ形の")


def compressed_rows(md: str, rsheets: list[dict]) -> dict:
    """窓が「○行目と同じ形の行が続きます」とまとめた行に、参照では値のあるセルがいくつあるか（まとめで消えた中身）。"""
    vals = {s["name"]: s.get("values") or {} for s in rsheets}
    rows_by_sheet: dict[str, set] = {}
    cur = None
    for ln in md.splitlines():
        h = re.match(r"^## シート「(.*)」", ln)
        if h:
            cur = h.group(1)
            continue
        m = _COMP.match(ln)
        if m and cur is not None:
            rows_by_sheet.setdefault(cur, set()).update(range(int(m.group(1)), int(m.group(2)) + 1))
    lost = rows = 0
    examples = []
    for sh, rws in rows_by_sheet.items():
        rows += len(rws)
        for addr, v in vals.get(sh, {}).items():
            r = int(re.sub(r"^[A-Z]+", "", addr))
            if r in rws and v not in (None, "", []):
                lost += 1
                if len(examples) < 5:
                    examples.append(f"{sh}!{addr}")
    return {"rows": rows, "nonempty_cells_in_compressed_rows": lost, "examples": examples}


def _num(v):
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(",", ""))
    except (TypeError, ValueError):
        return None


def same_value(rv, tv) -> bool:
    if isinstance(rv, list):
        return any(same_value(x, tv) for x in rv)
    if rv in (None, "") and tv in (None, ""):
        return True
    a, b = _num(rv), _num(tv)
    if a is not None and b is not None:
        return math.isclose(a, b, rel_tol=1e-9, abs_tol=1e-9)
    na, nb = norm(str(rv)), norm(str(tv if tv is not None else ""))
    return na == nb


def cell_compare(rsheets: list[dict], tsheets: list[dict]) -> dict:
    tmap = {s["name"]: s for s in tsheets}
    tot = agree = absent = 0
    diffs = []
    hidden_ref = [s["name"] for s in rsheets if s.get("state") not in (None, "visible")]
    hidden_tod = [s["name"] for s in tsheets if s.get("state") not in (None, "visible")]
    missing_sheets = [s["name"] for s in rsheets if s["name"] not in tmap]
    f_tot = f_marked = 0
    merged_ref = merged_tod = 0
    for s in rsheets:
        t = tmap.get(s["name"])
        merged_ref += len(s.get("merged") or [])
        if t:
            merged_tod += len(t.get("merged") or [])
        for addr, rv in (s.get("values") or {}).items():
            if rv in (None, "", []):
                continue
            tot += 1
            tv = (t or {}).get("cells", {}).get(addr) if t else None
            if tv is None and t and addr not in t.get("cells", {}):
                absent += 1
            if same_value(rv, tv):
                agree += 1
            elif len(diffs) < 25:
                diffs.append({"sheet": s["name"], "cell": addr, "ref": str(rv)[:60], "todoku": None if tv is None else str(tv)[:60]})
        for addr in (s.get("formulas") or {}):
            f_tot += 1
            if t and (addr in (t.get("formulas") or {})):
                f_marked += 1
    return {"cells_ref": tot, "cell_agree": (agree / tot if tot else None), "cells_absent": absent, "cell_diffs": diffs,
            "hidden_sheets": {"ref": hidden_ref, "todoku": hidden_tod}, "missing_sheets": missing_sheets,
            "formulas_ref": f_tot, "formulas_in_todoku_json": f_marked, "merged": {"ref": merged_ref, "todoku": merged_tod}}


# ====================================================================== 判断を合わせる
def wilson(k: float, n: float, z: float = 1.96) -> tuple[float | None, float | None]:
    if n <= 0:
        return None, None
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    h = z * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n)) / d
    return max(0.0, c - h), min(1.0, c + h)


def weighted_rate(rows: list[dict], key: str, n_boot: int = 4000, seed: int = SEED) -> dict:
    """層別の重み付きの割合と、層の中で取り直す bootstrap の 95% 区間。"""
    by: dict[str, list] = collections.defaultdict(list)
    for r in rows:
        by[r["stratum"]].append(r)
    Ns = {s: v[0]["N"] for s, v in by.items()}
    tot = sum(Ns.values())

    def est(groups):
        return sum(Ns[s] * (sum(1 for r in g if r[key]) / len(g)) for s, g in groups.items()) / tot

    p = est(by)
    # 層の中が全部同じ（0/8 など）だと、取り直しの区間は幅 0 になる。層ごとに Agresti–Coull（x+2)/(n+4) で分散を見積もった
    # 区間も出す（こちらを主に使う）：分散 = Σ W_h² p̃_h(1-p̃_h)/(n_h+4)
    var = 0.0
    adj = 0.0
    for s_, g in by.items():
        x = sum(1 for r in g if r[key])
        pt = (x + 2) / (len(g) + 4)
        wh = Ns[s_] / tot
        adj += wh * pt
        var += wh * wh * pt * (1 - pt) / (len(g) + 4)
    half = 1.96 * math.sqrt(var)
    ac = [max(0.0, p - half), min(1.0, p + half)]  # 中心は重み付きの割合、幅は Agresti–Coull の分散（控えめ）
    rnd = random.Random(seed)
    boots = sorted(est({s: [rnd.choice(g) for _ in g] for s, g in by.items()}) for _ in range(n_boot))
    return {"rate": p, "ci95": ac, "ci95_method": "estimate ± 1.96·SE（層ごとの分散を Agresti–Coull の p̃ で見積もる）", "ci95_bootstrap": [boots[int(0.025 * n_boot)], boots[int(0.975 * n_boot) - 1]],
            "n": len(rows), "population": tot}


def cmd_judge(a) -> int:
    work = Path(a.work)
    sample = json.loads((VAL / "sample.json").read_text(encoding="utf-8"))
    log = json.loads((work / "fetch_log.json").read_text(encoding="utf-8"))
    judg = json.loads((VAL / "judgments.json").read_text(encoding="utf-8"))
    rows = []
    for c in selected(sample, log):
        mp = work / "measure" / f"{c['id']}.json"
        m = json.loads(mp.read_text(encoding="utf-8")) if mp.exists() else {}
        j = judg.get(c["id"])
        if not j:
            print("判断がない:", c["id"])
            continue
        met = m.get("metrics", {})
        tod = m.get("todoku", {})
        rows.append({
            "id": c["id"], "url": c["url"], "sha256": c["sha256"], "stratum": c["stratum"], "kind": c["kind"],
            "N": sample["strata"][c["stratum"]]["N"],
            "todoku_read": c["read"], "todoku_write": c["write"], "todoku_reasons": c["read_reasons"],
            "regraded_now": [tod.get("read"), tod.get("write")],
            "judged_read": j["read"], "judged_write": j.get("write"),
            "read_agree": j["read"] == c["read"],
            "write_agree": (j.get("write") == c["write"]) if j.get("write") else None,
            "claim_true": j.get("claim_true"), "window_faithful": j.get("window_faithful"),
            "cause": j.get("cause", ""), "note": j.get("note", ""), "bug": j.get("bug"),
            "measures": {k: met.get(k) for k in ("char_recall", "unit_recall", "char_precision_raw", "cell_agree", "cells_ref",
                                                 "cells_absent", "formulas_ref", "merged", "hidden_sheets", "image_only_pages",
                                                 "acroform", "unmapped_todoku", "cid_chars_ref", "ref_chars")},
        })
    for r in rows:
        r["strict_agree"] = bool(r["read_agree"] and r["window_faithful"] is not False)
        tj, tn = GRADES.index(r["judged_read"]), GRADES.index(r["todoku_read"])
        r["within_one"] = abs(tj - tn) <= 1
        r["todoku_more_severe"] = tn > tj
        r["todoku_more_lenient"] = tn < tj
        now = r["regraded_now"][0]
        r["now_agree"] = now == r["judged_read"]
        r["now_same_as_recorded"] = now == r["todoku_read"]
        r["claim_ok"] = bool(r["claim_true"])
    (VAL / "grade_check.json").write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    summ = {"seed": sample["seed"], "cap": sample["cap"], "n": len(rows), "frame": sample["frame"],
            "absent_strata": sample["absent_strata"]}
    summ["read_agreement"] = weighted_rate(rows, "read_agree")
    summ["read_agreement_strict"] = weighted_rate(rows, "strict_agree")
    summ["read_within_one_grade"] = weighted_rate(rows, "within_one")
    summ["todoku_more_severe"] = weighted_rate(rows, "todoku_more_severe")
    summ["todoku_more_lenient"] = weighted_rate(rows, "todoku_more_lenient")
    # 調べたときのコード（写し）で付け直した段階：調査に記録された段階と同じか・確かめた段階と合うか（重みは同じ層の重み）
    summ["regrade_same_as_recorded"] = weighted_rate(rows, "now_same_as_recorded")
    summ["read_agreement_current_code"] = weighted_rate(rows, "now_agree")
    nc: dict = collections.defaultdict(lambda: collections.Counter())
    for r in rows:
        nc[r["regraded_now"][0] or "error"][r["judged_read"]] += 1
    summ["confusion_current_code"] = {g: dict(v) for g, v in sorted(nc.items())}
    summ["claim_true"] = weighted_rate(rows, "claim_ok")
    wf = [dict(r, wf=r["window_faithful"] is not False) for r in rows]
    summ["window_faithful"] = weighted_rate(wf, "wf")
    per = {}
    for g in GRADES:
        rr = [r for r in rows if r["todoku_read"] == g]
        if rr:
            k = sum(r["read_agree"] for r in rr)
            kc = sum(bool(r["claim_true"]) for r in rr)
            per[g] = {"n": len(rr), "agree": k, "rate": k / len(rr), "wilson95": wilson(k, len(rr)),
                      "claim_true": kc, "claim_wilson95": wilson(kc, len(rr))}
    summ["per_grade"] = per
    conf: dict = collections.defaultdict(lambda: collections.Counter())
    for r in rows:
        conf[r["todoku_read"]][r["judged_read"]] += 1
    summ["confusion"] = {g: dict(conf[g]) for g in GRADES if g in conf}
    ww = [r for r in rows if r["write_agree"] is not None]
    if ww:
        k = sum(r["write_agree"] for r in ww)
        summ["write_agreement_unweighted"] = {"n": len(ww), "agree": k, "rate": k / len(ww), "wilson95": wilson(k, len(ww))}
        wc: dict = collections.defaultdict(lambda: collections.Counter())
        for r in ww:
            wc[r["todoku_write"]][r["judged_write"]] += 1
        summ["write_confusion"] = {g: dict(v) for g, v in sorted(wc.items())}
    summ["disagreements"] = [{"id": r["id"], "todoku": r["todoku_read"], "judged": r["judged_read"], "cause": r["cause"]}
                             for r in rows if not r["read_agree"]]
    summ["bugs"] = [{"id": r["id"], "url": r["url"], "bug": r["bug"]} for r in rows if r.get("bug")]
    old = json.loads((VAL / "summary.json").read_text(encoding="utf-8")) if (VAL / "summary.json").exists() else {}
    old["grade_check"] = summ
    (VAL / "summary.json").write_text(json.dumps(old, ensure_ascii=False, indent=1), encoding="utf-8")
    print(json.dumps({k: (round(summ[k]["rate"], 3), [round(x, 3) for x in summ[k]["ci95"]]) for k in (
        "read_agreement", "read_agreement_strict", "read_within_one_grade", "todoku_more_severe", "todoku_more_lenient",
        "claim_true", "window_faithful", "regrade_same_as_recorded", "read_agreement_current_code")},
                     ensure_ascii=False, indent=1))
    print(json.dumps(summ["confusion"], ensure_ascii=False))
    return 0


def main(argv=None) -> int:
    global VAL
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("sample")
    s.add_argument("--results", required=True, help="調査の結果（1行に1件の JSON）")
    s.add_argument("--out", default=str(VAL), help="出力のフォルダ（sample.json・grade_check.json・summary.json）")
    s.add_argument("--seed", type=int, default=SEED)
    s.add_argument("--cap", type=int, default=CAP)
    for name in ("fetch", "measure", "recompare", "judge"):
        p = sub.add_parser(name)
        p.add_argument("--work", required=True, help="書類を置く作業用のフォルダ（リポジトリの外）")
        p.add_argument("--out", default=str(VAL), help="出力のフォルダ（sample.json がある所）")
        p.add_argument("--todoku-python", default=os.environ.get("TODOKU_PYTHON", sys.executable))
        p.add_argument("--baseline-python", default=os.environ.get("BASELINE_PYTHON", sys.executable))
        p.add_argument("--todoku-src", default=str(ROOT / "src"),
                       help="調べるとどくの src（調査の段階を付けたときのコードの写し。木のハッシュを記録する）")
        p.add_argument("--force", action="store_true")
        p.add_argument("--only", default="", help="この id だけ（カンマ区切り）")
    a = ap.parse_args(argv)
    VAL = Path(a.out).resolve()
    return {"sample": cmd_sample, "fetch": cmd_fetch, "measure": cmd_measure, "recompare": cmd_recompare, "judge": cmd_judge}[a.cmd](a)


if __name__ == "__main__":
    sys.exit(main())
