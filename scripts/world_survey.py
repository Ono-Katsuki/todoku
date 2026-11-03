"""世界調査：各国の政府の公式のオープンデータ目録から、公開書類を層別に無作為に選び、AI エージェントにとって
「届くか」「読めるか（読む段階 S〜F）」「書けるか（書く段階 A〜D）」を、決定的な処理で測る。設計は docs/世界調査の設計.md。

- 層：国 × 形式の群（PDF／表計算／文書／CSV）。層ごとに --quota 件（既定 250）。
- 層の中：目録の検索の何番目かを無作為に決めて登録を1つ取り、その登録の中の、その群の書類を1つ無作為に選ぶ（二段の抽出）。
  選ばれる重みを記録する（層の登録の数 × 登録の中の書類の数 ÷ 標本の数）。検索の深さに上限がある目録は、目録の中の小さな目録ごとに先に選ぶ。
- 礼儀：robots.txt を守り、同じ相手には間をあける（書類の取得 2 秒・目録の API 1 秒）。CAPTCHA・ログインは破らず、そう記録する。
- 書類は段階を付けたら消す（確かめ直し用の 1% だけを corpus/world/ に残す。公開しない）。公開するのは URL・SHA-256・結果だけ（world/）。
- 途中で止めても、もう一度動かせば続きから（world/results.jsonl にある標本は飛ばす）。

使い方: uv run python scripts/world_survey.py [--quota 250] [--countries JP,GB,...] [--jobs 24] [--graders 8]
       道具を直したあと、同じ標本で段階を付け直す：uv run python scripts/world_survey.py --regrade
"""

from __future__ import annotations

import argparse
import hashlib
import json
import random
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import urllib.robotparser
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "world"
KEEP = ROOT / "corpus" / "world"
UA = "todoku-research/0.2 (accessibility research on public documents; polite crawler; https://github.com/Ono-Katsuki/todoku)"
DELAY_DOC = 2.0
DELAY_API = 1.0
MAX_BYTES = 25 * 1024 * 1024
KEEP_RATE = 0.01  # 確かめ直し用に残す割合（SHA-256 で決める）
AUDIT_RATE = 0.10  # 道具の正しさを確かめる（全欄に書いて証明する）部分標本の割合（SHA-256 で決める）
EU_WINDOW = 10_000  # EU の統合目録の検索で取れる深さ

FAMILIES = {
    "pdf": ["PDF"],
    "sheet": ["XLSX", "XLS", "ODS"],
    "text": ["DOCX", "DOC", "ODT", "RTF"],
    "csv": ["CSV"],
}
FAMILY_OF = {f: fam for fam, fs in FAMILIES.items() for f in fs}
MIME = {
    "application/pdf": "PDF", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "XLSX",
    "application/vnd.ms-excel": "XLS", "application/msword": "DOC",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "DOCX",
    "application/vnd.oasis.opendocument.spreadsheet": "ODS", "application/vnd.oasis.opendocument.text": "ODT",
    "application/rtf": "RTF", "text/rtf": "RTF", "text/csv": "CSV",
}

# 国の目録（CKAN）：robots.txt で package_search に届くことを確かめたもの（2026-11-03）
CKAN = {
    "JP": ("日本", "https://data.e-gov.go.jp/data"),
    "GB": ("英国", "https://data.gov.uk"),
    "CA": ("カナダ", "https://open.canada.ca/data/en"),
    "NL": ("オランダ", "https://data.overheid.nl/data"),
    "GR": ("ギリシャ", "https://data.gov.gr"),
    "UA": ("ウクライナ", "https://data.gov.ua"),
    "HR": ("クロアチア", "https://data.gov.hr/ckan"),
    "LV": ("ラトビア", "https://data.gov.lv/dati"),
    "IL": ("イスラエル", "https://data.gov.il"),
}
# EU の統合目録（data.europa.eu）で国ごとに選ぶ国（国の目録に届かない国）
EU_HUB = "https://data.europa.eu/api/hub/search/search"
EU = {
    "DE": "ドイツ", "AT": "オーストリア", "IT": "イタリア", "BE": "ベルギー", "PT": "ポルトガル", "FR": "フランス",
    "SI": "スロベニア", "CH": "スイス", "ES": "スペイン", "IE": "アイルランド", "PL": "ポーランド", "SK": "スロバキア",
    "MD": "モルドバ", "SE": "スウェーデン", "FI": "フィンランド", "RO": "ルーマニア", "HU": "ハンガリー", "EE": "エストニア",
    "DK": "デンマーク", "ME": "モンテネグロ", "MT": "マルタ", "NO": "ノルウェー", "LT": "リトアニア", "CZ": "チェコ",
    "AL": "アルバニア", "IS": "アイスランド",
}
# 目録に届かなかった国（地図に「届かない」として示す。理由は probe の結果）
UNREACHED = {
    "AU": ("オーストラリア", "robots.txt で目録の API が禁止"), "NZ": ("ニュージーランド", "robots.txt で目録の API が禁止"),
    "TW": ("台湾", "robots.txt で目録の API が禁止"), "SG": ("シンガポール", "robots.txt で目録の API が禁止"),
    "TH": ("タイ", "robots.txt で目録の API が禁止"), "MX": ("メキシコ", "robots.txt で目録の API が禁止"),
    "BR": ("ブラジル", "robots.txt で目録の API が禁止"), "AR": ("アルゼンチン", "robots.txt で目録の API が禁止"),
    "CL": ("チリ", "robots.txt で目録の API が禁止"), "UY": ("ウルグアイ", "robots.txt で目録の API が禁止"),
    "EC": ("エクアドル", "robots.txt で目録の API が禁止"), "BG": ("ブルガリア", "robots.txt で目録の API が禁止"),
}


# ====================================================================== 礼儀
_last: dict[str, float] = {}
_host_locks: dict[str, threading.Lock] = {}
_lock = threading.Lock()
_robots: dict[str, urllib.robotparser.RobotFileParser | None] = {}


def polite_wait(host: str, delay: float) -> None:
    with _lock:
        now = time.time()
        at = max(now, _last.get(host, 0) + delay)
        _last[host] = at
    if at > now:
        time.sleep(at - now)


def allowed(url: str) -> bool:
    url = safe_url(url)
    p = urllib.parse.urlparse(url)
    base = f"{p.scheme}://{p.netloc}"
    with _lock:
        hl = _host_locks.setdefault(base, threading.Lock())
    with hl:
        if base not in _robots:
            rp = urllib.robotparser.RobotFileParser()
            rp.set_url(base + "/robots.txt")
            try:
                polite_wait(p.netloc, DELAY_DOC)
                req = urllib.request.Request(base + "/robots.txt", headers={"User-Agent": UA})
                with urllib.request.urlopen(req, timeout=20) as r:
                    rp.parse(r.read(500_000).decode("utf-8", "replace").splitlines())
            except urllib.error.HTTPError as e:
                rp = None if e.code >= 400 else rp  # robots.txt がない（4xx/5xx）：禁止の指定はない
            except Exception:  # noqa: BLE001
                rp = None
            _robots[base] = rp
    rp = _robots[base]
    return True if rp is None else rp.can_fetch(UA, url)


def safe_url(url: str) -> str:
    """空白・制御文字・ASCII でない字を含む URL を、意味を変えずに % の形にする（すでに % の形の所はそのまま）。"""
    return urllib.parse.quote(url.strip(), safe=":/?#[]@!$&'()*+,;=%~")


def fetch(url: str, delay: float, limit: int = MAX_BYTES, timeout: int = 60) -> dict:
    url = safe_url(url)
    p = urllib.parse.urlparse(url)
    polite_wait(p.netloc, delay)
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


def get_json(url: str):
    r = fetch(url, DELAY_API, limit=20_000_000, timeout=90)
    if r.get("status") != 200:
        return None
    try:
        return json.loads(r["body"])
    except json.JSONDecodeError:
        return None


def norm_format(s) -> str:
    s = str(s or "").strip()
    if s.lower() in MIME:
        return MIME[s.lower()]
    return s.upper().rsplit("/", 1)[-1].lstrip(".").replace("Х", "X")  # キリル文字の Х（хlsx）も


# ====================================================================== 標本の枠
class CkanFrame:
    def __init__(self, cc: str, base: str):
        self.cc, self.base = cc, base

    def _search(self, **params):
        q = urllib.parse.urlencode({k: (json.dumps(v) if isinstance(v, (list, dict)) else v) for k, v in params.items()})
        d = get_json(f"{self.base}/api/3/action/package_search?{q}")
        return (d or {}).get("result") if (d or {}).get("success", True) else None

    def strata(self) -> dict:
        """群 → [(目録での形式の名前, 登録の数)]。"""
        r = self._search(rows=0, **{"facet.field": ["res_format"], "facet.limit": 1000})
        if not r:
            return {}
        out: dict[str, list] = {}
        for it in r.get("search_facets", {}).get("res_format", {}).get("items", []):
            fam = FAMILY_OF.get(norm_format(it.get("name")))
            if fam:
                out.setdefault(fam, []).append((it["name"], int(it.get("count", 0))))
        return out

    def draw(self, fam: str, names: list, rnd: random.Random) -> dict | None:
        total = sum(c for _, c in names)
        nm, cnt = _pick_weighted(names, rnd)
        off = rnd.randrange(cnt)
        r = self._search(rows=1, start=off, fq=f'res_format:"{nm}"')
        if not r or not r.get("results"):
            return None
        pkg = r["results"][0]
        res = [x for x in pkg.get("resources", []) if FAMILY_OF.get(norm_format(x.get("format"))) == fam
               and str(x.get("url") or "").startswith(("http://", "https://"))]
        if not res:
            return {"empty": True}
        x = rnd.choice(res)
        return {"url": x["url"], "declared": norm_format(x.get("format")), "dataset": pkg.get("title") or pkg.get("name"),
                "dataset_id": pkg.get("id"), "publisher": (pkg.get("organization") or {}).get("title", ""),
                # 発行者が付けた分類（目録のグループ・タグ）。キーワードで当てたものではない
                "themes": [g.get("title") or g.get("name") for g in pkg.get("groups") or []],
                "tags": [t.get("name") for t in pkg.get("tags") or []][:30],
                "frame_datasets": total, "in_dataset": len(res)}


class EuFrame:
    def __init__(self, cc: str):
        self.cc = cc

    def _search(self, facets: dict, limit: int = 0, page: int = 0):
        q = urllib.parse.urlencode({"filter": "dataset", "limit": limit, "page": page, "facets": json.dumps(facets)})
        d = get_json(f"{EU_HUB}?{q}")
        return (d or {}).get("result")

    @staticmethod
    def _facet(r, fid) -> list:
        return [(i["id"], int(i["count"])) for f in (r or {}).get("facets", []) if f["id"] == fid for i in f["items"]]

    def strata(self) -> dict:
        r = self._search({"country": [self.cc.lower()]})
        out: dict[str, list] = {}
        for name, cnt in self._facet(r, "format"):
            fam = FAMILY_OF.get(norm_format(name))
            if fam and cnt:
                out.setdefault(fam, []).append((name, cnt))
        return out

    def draw(self, fam: str, names: list, rnd: random.Random) -> dict | None:
        total = sum(c for _, c in names)
        nm, cnt = _pick_weighted(names, rnd)
        facets = {"country": [self.cc.lower()], "format": [nm]}
        if cnt > EU_WINDOW:
            # 深い所が取れない：目録の中の小さな目録（catalog）を、登録の数に比べた確率で先に選ぶ
            cats = self._facet(self._search(facets), "catalog")
            if not cats:
                return None
            cat, cnt = _pick_weighted(cats, rnd)
            facets["catalog"] = [cat]
        off = rnd.randrange(min(cnt, EU_WINDOW))
        r = self._search(facets, limit=1, page=off)
        if not r or not r.get("results"):
            return None
        ds = r["results"][0]
        res = []
        for d in ds.get("distributions") or []:
            f = norm_format((d.get("format") or {}).get("id"))
            urls = (d.get("download_url") or []) + (d.get("access_url") or [])
            url = next((u for u in urls if str(u).startswith(("http://", "https://"))), None)
            if FAMILY_OF.get(f) == fam and url:
                res.append((url, f))
        if not res:
            return {"empty": True}
        url, f = rnd.choice(res)
        title = ds.get("title") or {}
        pub = ds.get("publisher") or {}
        return {"url": url, "declared": f, "dataset": (title.get("en") or next(iter(title.values()), "")) if isinstance(title, dict) else title,
                "dataset_id": ds.get("id"), "publisher": pub.get("name", "") if isinstance(pub, dict) else str(pub),
                "catalog": (ds.get("catalog") or {}).get("id"),
                # 発行者が付けた分類（DCAT-AP の data-theme：SOCI・HEAL・GOVE など 13 種）とキーワード
                "themes": [c.get("id") for c in ds.get("categories") or [] if isinstance(c, dict)],
                "tags": [k.get("label") or k.get("id") for k in ds.get("keywords") or [] if isinstance(k, dict)][:30],
                "frame_datasets": total, "in_dataset": len(res),
                "window_capped": cnt > EU_WINDOW}


def _pick_weighted(items: list, rnd: random.Random):
    total = sum(c for _, c in items)
    x = rnd.randrange(total)
    for nm, c in items:
        if x < c:
            return nm, c
        x -= c
    return items[-1]


# ====================================================================== 取得と段階
CAPTCHA = (b"captcha", b"cf-chl", b"challenge-platform", b"Just a moment", b"Attention Required", b"hcaptcha", b"recaptcha")


def reach(url: str) -> tuple[dict, bytes | None]:
    if not allowed(url):
        return {"reach": "robots"}, None
    r = fetch(url, DELAY_DOC)
    st, body = r.get("status", 0), r.get("body", b"")
    if st == 0:
        return {"reach": "error", "error": r.get("error", "")}, None
    if st in (401, 403):
        return {"reach": "captcha" if any(m in body[:200_000] for m in CAPTCHA) else f"http_{st}"}, None
    if st >= 400:
        return {"reach": f"http_{st}"}, None
    if r.get("too_large"):
        return {"reach": "too_large"}, None
    head = body[:2048].lstrip().lower()
    if head.startswith((b"<!doctype html", b"<html")) or b"<html" in head[:512]:
        if any(m in body[:200_000] for m in CAPTCHA):
            return {"reach": "captcha"}, None
        if b'type="password"' in body or any(w in r.get("final_url", "").lower() for w in ("login", "signin", "auth")):
            return {"reach": "login"}, None
        return {"reach": "html"}, None
    if not body:
        return {"reach": "empty"}, None
    return {"reach": "ok", "content_type": r.get("type", "")}, body


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


def grade_file(path: Path, do_audit: bool, timeout: int = 300) -> dict:
    assert _graders is not None
    with _graders:
        t = time.time()
        try:
            r = subprocess.run([sys.executable, "-c", GRADE_CODE, str(path), "1" if do_audit else "0"],
                               capture_output=True, timeout=timeout * (3 if do_audit else 1), cwd=ROOT)
            lines = r.stdout.decode("utf-8", "replace").strip().splitlines()
            out = json.loads(lines[-1]) if lines else {"error": "crash", "stderr": r.stderr.decode("utf-8", "replace")[-400:]}
        except subprocess.TimeoutExpired:
            out = {"error": "timeout"}
        except json.JSONDecodeError:
            out = {"error": "crash"}
        out["seconds"] = round(time.time() - t, 2)
        return out


EXT = {"PDF": ".pdf", "XLSX": ".xlsx", "XLS": ".xls", "ODS": ".ods", "DOCX": ".docx", "DOC": ".doc", "ODT": ".odt",
       "RTF": ".rtf", "CSV": ".csv"}


def _rate(h: str, salt: str) -> float:
    return int(hashlib.sha256((salt + h).encode()).hexdigest()[:8], 16) / 0xFFFFFFFF


def run_sample(s: dict, tmp: Path) -> dict:
    rec = dict(s)
    r, body = reach(s["url"])
    rec.update(r)
    if body is None:
        return rec
    h = hashlib.sha256(body).hexdigest()
    rec.update(sha256=h, bytes=len(body))
    # 標本ごとの名前（同じ書類が別の標本で同時に選ばれても、ほかの標本の段階付けの途中で消さない）
    path = tmp / f"{h[:20]}-{s['id']}{EXT.get(s.get('declared', ''), '')}"
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
def build_frames(countries: list[str]) -> dict:
    path = OUT / "frames.json"
    frames = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    todo = [cc for cc in countries if cc not in frames]

    def one(cc):
        fr = CkanFrame(cc, CKAN[cc][1]) if cc in CKAN else EuFrame(cc)
        st = fr.strata()
        return cc, {"name": CKAN[cc][0] if cc in CKAN else EU[cc], "source": CKAN[cc][1] if cc in CKAN else EU_HUB,
                    "strata": st, "built": time.strftime("%Y-%m-%d")}

    with ThreadPoolExecutor(8) as ex:
        for cc, f in ex.map(one, todo):
            frames[cc] = f
            print(f"[枠] {cc} {f['name']}: " + "・".join(f"{fam} {sum(c for _, c in v):,}" for fam, v in f["strata"].items()), flush=True)
    path.write_text(json.dumps(frames, ensure_ascii=False, indent=1), encoding="utf-8")
    return frames


GRADE_KEYS = ("content_type", "sha256", "bytes", "kind", "read", "write", "read_reasons", "write_reasons", "agent_checks",
              "alerts", "fields", "writable_fields", "chars", "pages", "seconds", "summary_error", "audit_ok", "audit_failed",
              "error", "changed")


def regrade(jobs: int, only_ungraded: bool = False, only_grades: str = "") -> int:
    """同じ標本（world/results.jsonl の URL）を取り直して、段階だけを付け直す（道具を直したあと）。
    届いた書類だけを取り直す（届かなかった標本は、そのまま）。中身が前と変わっていれば changed に記録する。
    only_ungraded：届いたが段階を付けられなかった書類（対応していなかった形式など）だけ。
    only_grades：読む段階がこれ（「E,F」など）の書類だけ（only_ungraded と合わせると、どちらか）。"""
    res_path = OUT / "results.jsonl"
    old = [json.loads(x) for x in res_path.read_text(encoding="utf-8").splitlines() if x.strip()]
    tmp = KEEP.parent / "world_tmp"
    tmp.mkdir(parents=True, exist_ok=True)
    new_path = OUT / "results.regrade.jsonl"
    done = {}
    if new_path.exists():
        for line in new_path.read_text(encoding="utf-8").splitlines():
            try:
                r = json.loads(line)
                done[r["id"]] = r
            except (json.JSONDecodeError, KeyError):
                pass
    lock = threading.Lock()

    def one(rec: dict):
        if rec["id"] in done or rec.get("reach") != "ok":
            return
        if only_ungraded or only_grades:
            pick = (only_ungraded and not rec.get("read")) or (rec.get("read") in only_grades.split(",") if only_grades else False)
            if not pick:
                return
        base = {k: v for k, v in rec.items() if k not in GRADE_KEYS}
        before = rec.get("sha256")
        r = run_sample({k: v for k, v in base.items() if k != "reach"}, tmp)
        if r.get("reach") != "ok":
            # 取り直せなかった（一時の 403・時間切れなど）：届いたという前の記録を残し、段階は前の道具のものと印を付ける
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
    # 前の順で、付け直した記録に置き換える
    merged = [done.get(r["id"], r) for r in old]
    res_path.write_text("".join(json.dumps(r, ensure_ascii=False, default=str) + "\n" for r in merged), encoding="utf-8")
    new_path.unlink()
    print("終わり")
    return 0


def main(argv=None) -> int:
    global _graders
    ap = argparse.ArgumentParser()
    ap.add_argument("--quota", type=int, default=250)
    ap.add_argument("--countries", default=",".join(list(CKAN) + list(EU)))
    ap.add_argument("--seed", type=int, default=20261103)
    ap.add_argument("--jobs", type=int, default=24)
    ap.add_argument("--graders", type=int, default=8)
    ap.add_argument("--regrade", action="store_true", help="同じ標本を取り直して、段階だけを付け直す")
    ap.add_argument("--only-ungraded", action="store_true", help="--regrade で、段階を付けられなかった書類だけ")
    ap.add_argument("--only-grades", default="", help="--regrade で、読む段階がこれ（E,F など）の書類だけ")
    a = ap.parse_args(argv)
    OUT.mkdir(parents=True, exist_ok=True)
    _graders = threading.Semaphore(a.graders)
    if a.regrade:
        return regrade(a.jobs, a.only_ungraded, a.only_grades)
    countries = [c for c in a.countries.split(",") if c]
    frames = build_frames(countries)
    (OUT / "unreached.json").write_text(json.dumps(UNREACHED, ensure_ascii=False, indent=1), encoding="utf-8")
    res_path = OUT / "results.jsonl"
    done = set()
    if res_path.exists():
        for line in res_path.read_text(encoding="utf-8").splitlines():
            try:
                done.add(json.loads(line)["id"])
            except (json.JSONDecodeError, KeyError):
                pass
    tmp = KEEP.parent / "world_tmp"
    tmp.mkdir(parents=True, exist_ok=True)
    out_lock = threading.Lock()

    def cell(cc: str, fam: str):
        names = frames[cc]["strata"].get(fam) or []
        if not names:
            return
        fr = CkanFrame(cc, CKAN[cc][1]) if cc in CKAN else EuFrame(cc)
        n = min(a.quota, sum(c for _, c in names))
        for i in range(n):
            sid = f"{cc}-{fam}-{i:04d}"
            if sid in done:
                continue
            rnd = random.Random(f"{a.seed}:{sid}")  # 標本ごとの種：続きからでも、同じ標本になる
            d = None
            for _try in range(3):
                d = fr.draw(fam, names, rnd)
                if d and not d.get("empty"):
                    break
            if not d or d.get("empty"):
                rec = {"id": sid, "country": cc, "family": fam, "reach": "frame_empty"}
            else:
                rec = run_sample({"id": sid, "country": cc, "family": fam, "quota": n, **d}, tmp)
            with out_lock:
                with res_path.open("a", encoding="utf-8") as f:
                    f.write(json.dumps(rec, ensure_ascii=False, default=str) + "\n")
            print(f"[{rec.get('reach')}] {sid} {rec.get('kind') or ''} 読む {rec.get('read', '-')} 書く {rec.get('write', '-')} "
                  f"{rec.get('error', '')} {str(rec.get('url', ''))[:80]}", flush=True)

    cells = [(cc, fam) for cc in countries if cc in frames for fam in FAMILIES]
    with ThreadPoolExecutor(a.jobs) as ex:
        list(ex.map(lambda x: cell(*x), cells))
    print("終わり")
    return 0


if __name__ == "__main__":
    sys.exit(main())
