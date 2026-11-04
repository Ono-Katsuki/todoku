"""世界調査：各国の政府のサイトにある書類を、Common Crawl の収集から層別に無作為に選び、AI エージェントにとって
「読めるか（読む段階 S〜F）」「書けるか（書く段階 A〜D）」を、決定的な処理で測る。設計は docs/世界調査の設計.md。

- 枠と標本は scripts/world_frame.py が作る（world/domains.json・world/frames.json・world/sample.jsonl）。
  層は国 × 形式の群（pdf・sheet・text・csv）。標本は層ごとに順番（rank）が付いていて、--rank 未満を測る
  （1回目 20、2回目 50。選び直さないので、2回目は足すだけ）。
- 書類は Common Crawl の保存庫（data.commoncrawl.org）から、索引の位置（WARC のファイル・offset・length）で
  その記録だけを範囲指定で取る。政府のサーバーには負担をかけない。保存の中身の SHA-1 が索引の digest と合うことを確かめる。
- Common Crawl は大きな書類を途中で切って保存する（WARC-Truncated）。切れた書類は、そのままでは段階を付けない。
  その URL の robots.txt（RFC 9309）が許すときだけ、政府のサイトから取り直す（同じホストには 2 秒以上の間・一度に一つ）。
  許されない・取れないときは truncated と記録する。どこから取ったか（cc・live）を記録する。
- 礼儀：Common Crawl への要求は一度に一つ・0.5 秒以上の間・混んでいたら待って取り直す。名乗る User-Agent を付ける。
  CAPTCHA・ログインは破らず、そう記録する。
- 書類は段階を付けたら消す（確かめ直し用の 1% だけを corpus/world/ に残す。公開しない）。公開するのは URL・SHA-256・結果だけ。
- 途中で止めても、もう一度動かせば続きから（world/results.jsonl にある標本は飛ばす）。

使い方: PYTHONPATH=src .venv/bin/python scripts/world_survey.py [--rank 20] [--countries JP,DE] [--jobs 4] [--graders 2]
       道具を直したあと、同じ標本で段階を付け直す：... scripts/world_survey.py --regrade
"""

from __future__ import annotations

import argparse
import base64
import gzip
import hashlib
import http.client
import json
import re
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "world"
KEEP = ROOT / "corpus" / "world"
TMP = ROOT / "corpus" / "world_tmp"
UA = "todoku-research/0.2 (accessibility research on public documents; polite crawler; https://github.com/Ono-Katsuki/todoku)"
CC_DATA = "https://data.commoncrawl.org/"
DELAY_DOC = 2.0  # 政府のサイト（切れた書類の取り直しだけ）：同じホストに 2 秒以上
DELAY_CC = 0.5  # Common Crawl：0.5 秒以上
MAX_BYTES = 25 * 1024 * 1024
KEEP_RATE = 0.01  # 確かめ直し用に残す割合（SHA-256 で決める）
AUDIT_RATE = 0.10  # 道具の正しさを確かめる（全欄に書いて証明する）部分標本の割合（SHA-256 で決める）


# ====================================================================== 礼儀
_last: dict[str, float] = {}
_host_locks: dict[str, threading.Lock] = {}
_lock = threading.Lock()
_robots: dict[str, "Robots"] = {}
_cc = threading.Lock()  # Common Crawl へは一度に一つ


def polite_wait(host: str, delay: float) -> None:
    with _lock:
        now = time.time()
        at = max(now, _last.get(host, 0) + delay)
        _last[host] = at
    if at > now:
        time.sleep(at - now)


class Robots:
    """robots.txt を RFC 9309 のとおりに読む。標準ライブラリの urllib.robotparser は、User-agent の行のあとの空行で組を閉じ、
    最初に当たった行で決め（最も長く当たる行でなく）、* や $ を解さないため、禁止されている経路を許可と答えることがある。"""

    TOKEN = UA.split("/", 1)[0].lower()  # 製品の名前（todoku-research）

    def __init__(self, text: str | None, status: int):
        self.status = status
        self.rules: list[tuple[bool, str]] = []
        self.delay = 0.0
        if text is None:
            return
        groups, cur, in_ua = [], None, False
        for line in text.splitlines():
            line = line.split("#", 1)[0].strip()
            if ":" not in line:
                continue  # 空行は組を閉じない（RFC 9309 2.2）
            k, v = (x.strip() for x in line.split(":", 1))
            k = k.lower()
            if k == "user-agent":
                if not in_ua:
                    cur = {"ua": [], "rules": [], "delay": None}
                    groups.append(cur)
                cur["ua"].append(v.lower())
                in_ua = True
                continue
            in_ua = False
            if cur is None:
                continue
            if k in ("allow", "disallow") and v:
                cur["rules"].append((k == "allow", v))
            elif k == "crawl-delay":
                try:
                    cur["delay"] = float(v)
                except ValueError:
                    pass
        mine = [g for g in groups if any(u != "*" and self.TOKEN.startswith(u) for u in g["ua"])]
        use = mine or [g for g in groups if "*" in g["ua"]]  # 名指しの組があればそれだけ、なければ * の組（同じ名の組はまとめる）
        for g in use:
            self.rules += g["rules"]
            self.delay = max(self.delay, g["delay"] or 0)

    @staticmethod
    def _match(rule: str, path: str) -> bool:
        pat = re.escape(rule).replace(r"\*", ".*")
        if pat.endswith(r"\$"):
            pat = pat[:-2] + "$"
        return re.match(pat, path) is not None

    def can_fetch(self, url: str) -> bool:
        if self.status >= 500 or self.status == 0:
            return False  # robots.txt に届かない（5xx・つながらない）：すべて禁止とみなす（RFC 9309 2.3.1.4）
        p = urllib.parse.urlparse(url)
        path = (p.path or "/") + (f"?{p.query}" if p.query else "")
        hits = [(len(r), allow) for allow, r in self.rules if self._match(r, path)]
        return True if not hits else max(hits)[1]  # 最も長く当たる行。同じ長さなら Allow（RFC 9309 2.2.2）


def allowed(url: str) -> bool:
    url = safe_url(url)
    p = urllib.parse.urlparse(url)
    base = f"{p.scheme}://{p.netloc}"
    with _lock:
        hl = _host_locks.setdefault(base, threading.Lock())
    with hl:
        if base not in _robots:
            polite_wait(p.netloc, DELAY_DOC)
            req = urllib.request.Request(base + "/robots.txt", headers={"User-Agent": UA})
            try:
                with urllib.request.urlopen(req, timeout=20) as r:
                    _robots[base] = Robots(r.read(500_000).decode("utf-8", "replace"), r.status)
            except urllib.error.HTTPError as e:
                _robots[base] = Robots(None, e.code)  # 4xx：禁止の指定はない。5xx：すべて禁止とみなす
            except Exception:  # noqa: BLE001
                _robots[base] = Robots(None, 0)
    return _robots[base].can_fetch(url)


def crawl_delay(url: str) -> float:
    """robots.txt の Crawl-delay（あれば）。間はこれと決めた間の長い方をとる。"""
    p = urllib.parse.urlparse(safe_url(url))
    rb = _robots.get(f"{p.scheme}://{p.netloc}")
    return rb.delay if rb else 0.0


def safe_url(url: str) -> str:
    """空白・制御文字・ASCII でない字を含む URL を、意味を変えずに % の形にする（すでに % の形の所はそのまま）。"""
    return urllib.parse.quote(url.strip(), safe=":/?#[]@!$&'()*+,;=%~")


def fetch(url: str, delay: float, limit: int = MAX_BYTES, timeout: int = 60) -> dict:
    url = safe_url(url)
    p = urllib.parse.urlparse(url)
    with _lock:
        hl = _host_locks.setdefault(f"get:{p.netloc}", threading.Lock())
    with hl:  # 同じホストには一度に一つ
        polite_wait(p.netloc, max(delay, crawl_delay(url)))
        req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
        try:
            with urllib.request.urlopen(req, timeout=timeout) as r:
                body = r.read(limit + 1)
                return {"status": r.status, "final_url": r.geturl(), "type": r.headers.get("Content-Type", ""),
                        "body": body, "too_large": len(body) > limit}
        except urllib.error.HTTPError as e:
            try:
                body = e.read(200_000)
            except Exception:  # noqa: BLE001
                body = b""
            return {"status": e.code, "final_url": url, "type": (e.headers.get("Content-Type", "") if e.headers else ""), "body": body}
        except Exception as e:  # noqa: BLE001
            return {"status": 0, "error": f"{type(e).__name__}: {str(e)[:160]}", "body": b""}


# ====================================================================== Common Crawl の保存庫から取る
_cc_conn: list = [None]


def cc_get(path: str, offset: int, length: int) -> bytes:
    """WARC の記録を一つだけ範囲指定で取る。一度に一つ・間をあける・混んでいたら待って取り直す（間を倍に延ばす）。
    つながりは一つを使い回す（keep-alive）。"""
    wait = 2.0
    for attempt in range(8):
        with _cc:
            polite_wait("data.commoncrawl.org", DELAY_CC)
            st, raw, err = 0, b"", ""
            try:
                if _cc_conn[0] is None:
                    _cc_conn[0] = http.client.HTTPSConnection("data.commoncrawl.org", timeout=120)
                c = _cc_conn[0]
                c.request("GET", "/" + path, headers={"User-Agent": UA, "Range": f"bytes={offset}-{offset + length - 1}"})
                r = c.getresponse()
                st, raw = r.status, r.read()
                err = f"HTTP {st}"
            except (http.client.HTTPException, OSError) as e:
                err = type(e).__name__
                if _cc_conn[0] is not None:
                    _cc_conn[0].close()
                _cc_conn[0] = None
        if st == 206 and len(raw) == length:
            return raw
        if st not in (0, 206, 429, 500, 502, 503, 504):
            raise RuntimeError(err)
        if attempt == 7:
            raise RuntimeError(f"{err}（取り直しても取れない）")
        print(f"  （Common Crawl が混んでいる：{err}。{wait:.0f} 秒待つ）", flush=True)
        time.sleep(wait)
        wait = min(wait * 2, 300)
    raise RuntimeError("unreachable")


def parse_warc(raw: bytes) -> dict:
    """gzip の1メンバーの WARC 応答の記録 → WARC の見出し・HTTP の見出し・中身。"""
    rec = gzip.decompress(raw)
    head, rest = rec.split(b"\r\n\r\n", 1)
    wh = {}
    for line in head.decode("utf-8", "replace").split("\r\n")[1:]:
        if ":" in line:
            k, v = line.split(":", 1)
            wh[k.strip().lower()] = v.strip()
    block = rest[: int(wh.get("content-length", len(rest)))]
    hh, _, payload = block.partition(b"\r\n\r\n")
    http = {}
    lines = hh.decode("latin-1").split("\r\n")
    for line in lines[1:]:
        if ":" in line:
            k, v = line.split(":", 1)
            http[k.strip().lower()] = v.strip()
    return {"warc": wh, "status_line": lines[0] if lines else "", "http": http, "payload": payload}


def sha1_b32(b: bytes) -> str:
    return base64.b32encode(hashlib.sha1(b).digest()).decode()


CAPTCHA = (b"captcha", b"cf-chl", b"challenge-platform", b"Just a moment", b"Attention Required", b"hcaptcha", b"recaptcha")


def live_copy(url: str) -> tuple[dict, bytes | None]:
    """切れた書類を、政府のサイトから取り直す（robots.txt が許すときだけ）。"""
    if not allowed(url):
        return {"live": "robots"}, None
    r = fetch(url, DELAY_DOC)
    st, body = r.get("status", 0), r.get("body", b"")
    if st == 0:
        return {"live": "error", "live_error": r.get("error", "")}, None
    if st in (401, 403):
        return {"live": "captcha" if any(m in body[:200_000] for m in CAPTCHA) else f"http_{st}"}, None
    if st >= 400:
        return {"live": f"http_{st}"}, None
    if r.get("too_large"):
        return {"live": "too_large"}, None
    head = body[:2048].lstrip().lower()
    if head.startswith((b"<!doctype html", b"<html")) or b"<html" in head[:512]:
        if any(m in body[:200_000] for m in CAPTCHA):
            return {"live": "captcha"}, None
        if b'type="password"' in body or any(w in r.get("final_url", "").lower() for w in ("login", "signin", "auth")):
            return {"live": "login"}, None
        return {"live": "html"}, None
    if not body:
        return {"live": "empty"}, None
    return {"live": "ok", "live_content_type": r.get("type", "")}, body


def obtain(s: dict) -> tuple[dict, bytes | None]:
    """標本の書類の中身。reach：ok（cc か live）・truncated・cc_error・digest_mismatch など。"""
    try:
        w = parse_warc(cc_get(s["filename"], s["offset"], s["length"]))
    except Exception as e:  # noqa: BLE001
        return {"reach": "cc_error", "error": f"{type(e).__name__}: {str(e)[:160]}"}, None
    info = {"warc_date": w["warc"].get("warc-date"), "content_type": w["http"].get("content-type", "")}
    trunc = w["warc"].get("warc-truncated")
    if trunc:
        info["cc_truncated"] = trunc
        info["cc_bytes"] = len(w["payload"])
        live, body = live_copy(s["url"])
        info.update(live)
        if body is None:
            return {**info, "reach": "truncated"}, None
        info["source"] = "live"
        info["same_as_crawl"] = sha1_b32(body) == s["digest"]  # 収集のあとで書き換わったか（切れる前の中身の digest）
        return {**info, "reach": "ok"}, body
    body = w["payload"]
    if sha1_b32(body) != s["digest"]:
        return {**info, "reach": "digest_mismatch"}, None
    if not body:
        return {**info, "reach": "empty"}, None
    return {**info, "reach": "ok", "source": "cc"}, body


# ====================================================================== 段階
GRADE_CODE = r"""
import json, sys
from pathlib import Path
from todoku.book import kind_of
from todoku import api
p, do_audit = sys.argv[1], sys.argv[2] == "1"
b = open(p, "rb").read()
out = {}
try:
    k = kind_of(b, Path(p).name)
except Exception as e:
    k = None
out["kind"] = k
if k not in api.KIND_JA:
    out["error"] = "encrypted" if k == "encrypted" else "unknown_kind"
    print(json.dumps(out)); sys.exit()
try:
    g = api.grade(p)
except api.PasswordNeeded:
    out["error"] = "password"; print(json.dumps(out)); sys.exit()
out["read"] = g["grade"]; out["write"] = g["write"]["grade"]
out["read_reasons"] = [(r["grade"], r["title"]) for r in g["reasons"]]
out["write_reasons"] = [(r["grade"], r["title"]) for r in g["write"]["reasons"]]
out["agent_checks"] = len(g.get("agent_checks") or [])
out["alerts"] = [a.get("title") for a in g.get("alerts") or []]
try:
    s = api.summary(p)
    out.update(fields=s["fields"], writable_fields=s["writable_fields"], chars=s["window_chars"], pages=s.get("pages"))
except Exception as e:
    out["summary_error"] = type(e).__name__
if do_audit:
    try:
        au = api.audit(p)
        out["audit_ok"] = au["ok"]
        out["audit_failed"] = [c["name"] for c in au["checks"] if not c["ok"]]
    except Exception as e:
        out["audit_ok"] = False; out["audit_failed"] = ["crash:" + type(e).__name__]
print(json.dumps(out, ensure_ascii=False, default=str))
"""

_graders: threading.Semaphore | None = None


MEM_LIMIT_MB = 4096  # 段階付けの1本が使ってよいメモリ。超えたら、その書類だけ止める（パソコン全体を巻き込まない）


def _rss_mb(pid: int) -> float:
    """その処理が使っている実メモリ（MB）。ps の RSS（KB）から。"""
    try:
        out = subprocess.run(["ps", "-o", "rss=", "-p", str(pid)], capture_output=True, text=True, timeout=5).stdout.strip()
        return int(out) / 1024 if out else 0.0
    except (ValueError, subprocess.SubprocessError):
        return 0.0


def grade_file(path: Path, do_audit: bool, timeout: int = 300) -> dict:
    assert _graders is not None
    with _graders:
        t = time.time()
        limit = timeout * (3 if do_audit else 1)
        # 優先度を下げて動かし、メモリと時間を見張る
        proc = subprocess.Popen(["nice", "-n", "10", sys.executable, "-c", GRADE_CODE, str(path), "1" if do_audit else "0"],
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, cwd=ROOT)
        why = None
        peak = 0.0
        out_b, err_b = b"", b""
        while True:
            try:
                out_b, err_b = proc.communicate(timeout=1)
                break
            except subprocess.TimeoutExpired:
                peak = max(peak, _rss_mb(proc.pid))
                if peak > MEM_LIMIT_MB:
                    why = "memory"
                elif time.time() - t > limit:
                    why = "timeout"
                if why:
                    proc.kill()
                    proc.communicate()
                    break
        if why:
            out = {"error": why, **({"peak_mb": round(peak)} if why == "memory" else {})}
        else:
            lines = out_b.decode("utf-8", "replace").strip().splitlines()
            try:
                out = json.loads(lines[-1]) if lines else {"error": "crash", "stderr": err_b.decode("utf-8", "replace")[-400:]}
            except json.JSONDecodeError:
                out = {"error": "crash"}
        out["seconds"] = round(time.time() - t, 2)
        return out


# 書類の名前の拡張子：とどくは CSV・TSV・.txt を中身だけでは決めず、名前で見分けるため、URL の拡張子・中身で見分けた形式から決める
KNOWN_SUFFIX = {".pdf", ".xlsx", ".xlsm", ".xls", ".ods", ".docx", ".doc", ".odt", ".rtf", ".csv", ".tsv", ".txt"}
MIME_SUFFIX = {
    "application/pdf": ".pdf", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-excel": ".xls", "application/vnd.ms-excel.sheet.macroenabled.12": ".xlsm",
    "application/vnd.oasis.opendocument.spreadsheet": ".ods", "application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.oasis.opendocument.text": ".odt", "application/rtf": ".rtf", "text/rtf": ".rtf", "text/csv": ".csv",
}


def suffix_for(s: dict) -> str:
    ext = Path(urllib.parse.unquote(urllib.parse.urlsplit(s["url"]).path)).suffix.lower()
    md = (s.get("mime_detected") or "").lower()
    if ext in KNOWN_SUFFIX and MIME_SUFFIX.get(md, ext)[:4] == ext[:4]:
        return ext  # 名前と中身が合う（.xlsm・.tsv なども、そのまま）
    return MIME_SUFFIX.get(md, ext if ext in KNOWN_SUFFIX else "")


def _rate(h: str, salt: str) -> float:
    return int(hashlib.sha256((salt + h).encode()).hexdigest()[:8], 16) / 0xFFFFFFFF


def run_sample(s: dict) -> dict:
    rec = dict(s)
    r, body = obtain(s)
    rec.update(r)
    if body is None:
        return rec
    h = hashlib.sha256(body).hexdigest()
    rec.update(sha256=h, bytes=len(body))
    # 標本ごとの名前（同じ書類が別の標本で同時に選ばれても、ほかの標本の段階付けの途中で消さない）
    path = TMP / f"{h[:20]}-{s['id']}{suffix_for(s)}"
    path.write_bytes(body)
    try:
        rec.update(grade_file(path, _rate(h, "audit") < AUDIT_RATE))
    finally:
        if _rate(h, "keep") < KEEP_RATE:
            KEEP.mkdir(parents=True, exist_ok=True)
            path.replace(KEEP / f"{h[:20]}{path.suffix}")
        else:
            path.unlink(missing_ok=True)
    return rec


# ====================================================================== 全体
GRADE_KEYS = ("sha256", "bytes", "kind", "read", "write", "read_reasons", "write_reasons", "agent_checks",
              "alerts", "fields", "writable_fields", "chars", "pages", "seconds", "summary_error", "audit_ok", "audit_failed",
              "error", "changed", "peak_mb", "stderr")
OBTAIN_KEYS = ("reach", "warc_date", "content_type", "cc_truncated", "cc_bytes", "live", "live_error", "live_content_type",
               "source", "same_as_crawl")


def _read_results(path: Path) -> list[dict]:
    out = []
    if path.exists():
        for line in path.read_text(encoding="utf-8").splitlines():
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                pass  # 止めたときに途中まで書かれた行
    return out


def regrade(jobs: int, only_ungraded: bool = False, only_grades: str = "", results: Path | None = None) -> int:
    """同じ標本（world/results.jsonl）を取り直して、段階だけを付け直す（道具を直したあと）。
    届いた書類だけを取り直す（届かなかった標本は、そのまま）。中身が前と変わっていれば changed に記録する。
    only_ungraded：届いたが段階を付けられなかった書類（対応していなかった形式など）だけ。
    only_grades：読む段階がこれ（「E,F」など）の書類だけ（only_ungraded と合わせると、どちらか）。"""
    res_path = results or OUT / "results.jsonl"
    old = _read_results(res_path)
    new_path = res_path.with_name("results.regrade.jsonl")
    done = {r["id"]: r for r in _read_results(new_path) if "id" in r}
    lock = threading.Lock()

    def one(rec: dict):
        if rec["id"] in done or rec.get("reach") != "ok":
            return
        if only_ungraded or only_grades:
            pick = (only_ungraded and not rec.get("read")) or (rec.get("read") in only_grades.split(",") if only_grades else False)
            if not pick:
                return
        base = {k: v for k, v in rec.items() if k not in GRADE_KEYS and k not in OBTAIN_KEYS}
        before = rec.get("sha256")
        r = run_sample(base)
        if r.get("reach") != "ok":
            # 取り直せなかった（一時の失敗など）：届いたという前の記録を残し、段階は前の道具のものと印を付ける
            r = dict(rec, regrade=f"取り直せない（{r.get('reach')}）")
        elif before and r.get("sha256") != before:
            r["changed"] = True
        with lock:
            with new_path.open("a", encoding="utf-8") as f:
                f.write(json.dumps(r, ensure_ascii=False, default=str) + "\n")
            done[r["id"]] = r
        print(f"[{r.get('reach')}] {r['id']} {r.get('kind') or ''} 読む {rec.get('read', '-')}→{r.get('read', '-')} "
              f"書く {rec.get('write', '-')}→{r.get('write', '-')}", flush=True)

    with ThreadPoolExecutor(jobs) as ex:
        list(ex.map(one, old))
    merged = [done.get(r["id"], r) for r in old]  # 前の順で、付け直した記録に置き換える
    res_path.write_text("".join(json.dumps(r, ensure_ascii=False, default=str) + "\n" for r in merged), encoding="utf-8")
    new_path.unlink(missing_ok=True)
    print("終わり")
    return 0


def main(argv=None) -> int:
    global _graders
    ap = argparse.ArgumentParser()
    ap.add_argument("--rank", type=int, default=20, help="層ごとに、標本の順番がこれ未満のものを測る（1回目 20・2回目 50）")
    ap.add_argument("--countries", default="", help="既定：world/sample.jsonl のすべての国")
    ap.add_argument("--jobs", type=int, default=4)
    ap.add_argument("--graders", type=int, default=2)
    ap.add_argument("--regrade", action="store_true", help="同じ標本を取り直して、段階だけを付け直す")
    ap.add_argument("--only-ungraded", action="store_true", help="--regrade で、段階を付けられなかった書類だけ")
    ap.add_argument("--only-grades", default="", help="--regrade で、読む段階がこれ（E,F など）の書類だけ")
    ap.add_argument("--results", default="", help="--regrade で、結果のファイル（既定：world/results.jsonl。最終の評価では作業ツリーを汚さないよう外の写しを渡す）")
    a = ap.parse_args(argv)
    _graders = threading.Semaphore(a.graders)
    TMP.mkdir(parents=True, exist_ok=True)
    if a.regrade:
        return regrade(a.jobs, a.only_ungraded, a.only_grades, Path(a.results) if a.results else None)
    want = {c for c in a.countries.split(",") if c}
    samples = [json.loads(x) for x in (OUT / "sample.jsonl").read_text(encoding="utf-8").splitlines() if x.strip()]
    samples = [s for s in samples if s["rank"] < a.rank and (not want or s["country"] in want)]
    res_path = OUT / "results.jsonl"
    # 済みの標本：id と digest が同じもの（枠を作り直して同じ id に別の書類が入ったときは、測り直す）
    done = {(r["id"], r.get("digest")) for r in _read_results(res_path) if "id" in r}
    todo = [s for s in samples if (s["id"], s["digest"]) not in done]
    # 国を混ぜて進める（止めたときに、どの国もそれなりに進んでいるように）：順番 → 国 → 群
    todo.sort(key=lambda s: (s["rank"], s["country"], s["family"]))
    print(f"標本 {len(samples):,}・済み {len(samples) - len(todo):,}・これから {len(todo):,}", flush=True)
    out_lock = threading.Lock()
    t0, n = time.time(), [0]

    def one(s: dict):
        rec = run_sample(s)
        with out_lock:
            with res_path.open("a", encoding="utf-8") as f:
                f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
            n[0] += 1
            eta = (time.time() - t0) / n[0] * (len(todo) - n[0]) / 3600
        print(f"[{rec.get('reach')}{'・live' if rec.get('source') == 'live' else ''}] {s['id']} {rec.get('kind') or ''} "
              f"読む {rec.get('read', '-')} 書く {rec.get('write', '-')} {rec.get('error', '')} "
              f"（{n[0]}/{len(todo)}・残り約 {eta:.1f} 時間）", flush=True)

    with ThreadPoolExecutor(a.jobs) as ex:
        list(ex.map(one, todo))
    print("終わり")
    return 0


if __name__ == "__main__":
    sys.exit(main())
