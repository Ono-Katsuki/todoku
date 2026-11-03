"""世界調査：各国の政府の公式のオープンデータ目録から、公開書類を層別に無作為に選び、AI エージェントにとって
「届くか」「読めるか（読む段階 S〜F）」「書けるか（書く段階 A〜D）」を、決定的な処理で測る。設計は docs/世界調査の設計.md。

- 層：国 × 形式の群（PDF／表計算／文書／CSV）。層ごとに --quota 件（既定 250）。
- 層の中：目録の検索の何番目かを無作為に決めて登録を1つ取り、その登録の中の、その群の書類を1つ無作為に選ぶ（二段の抽出）。
  選ばれる重みを記録する（層の登録の数 × 登録の中の書類の数 ÷ 標本の数）。検索の深さに上限がある目録は、目録の中の小さな目録ごとに先に選ぶ。
- 母集団：国連の加盟 193 か国。国ごとの結果（届く・目録はあるが使えない・見つからない）と根拠は world/population.json。
  届く国だけ、目録の型（CKAN・Socrata・udata・DCAT の data.json・data.gov.sg）に合う読み取り口で枠を作る。
- 礼儀：robots.txt を RFC 9309 のとおりに守り（目録の API も）、同じ相手には間をあける（書類の取得 2 秒・目録の API 1 秒、
  Crawl-delay があればその長い方）。CAPTCHA・ログインは破らず、そう記録する。
- 書類は段階を付けたら消す（確かめ直し用の 1% だけを corpus/world/ に残す。公開しない）。公開するのは URL・SHA-256・結果だけ（world/）。
- 途中で止めても、もう一度動かせば続きから（world/results.jsonl にある標本は飛ばす）。

使い方: uv run python scripts/world_survey.py [--quota 250] [--countries JP,GB,...] [--jobs 24] [--graders 8]
       標本の枠だけを作る（書類は取らない）：uv run python scripts/world_survey.py --frames-only --countries CO
       道具を直したあと、同じ標本で段階を付け直す：uv run python scripts/world_survey.py --regrade
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import random
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

EU_HUB = "https://data.europa.eu/api/hub/search/search"
# 母集団：国連の加盟 193 か国（と、参考の地域）。国ごとに、公式の目録を探した結果（届く・目録はあるが使えない・見つからない）と
# その根拠を world/population.json に記録する。届く国だけ、目録の型（platform）に合う読み取り口で標本の枠を作る
POPULATION = OUT / "population.json"
CATALOG = OUT / "catalog"  # 一覧を取る型の目録（Socrata・DCAT の data.json・data.gov.sg）の登録の一覧（枠を作り直せるように残す）


def load_population() -> list[dict]:
    return json.loads(POPULATION.read_text(encoding="utf-8"))["countries"]


def reachable() -> dict:
    return {r["iso2"]: r for r in load_population() if r["outcome"] == "reachable"}


# ====================================================================== 礼儀
_last: dict[str, float] = {}
_host_locks: dict[str, threading.Lock] = {}
_lock = threading.Lock()
_robots: dict[str, "Robots"] = {}


def polite_wait(host: str, delay: float) -> None:
    with _lock:
        now = time.time()
        at = max(now, _last.get(host, 0) + delay)
        _last[host] = at
    if at > now:
        time.sleep(at - now)


class Robots:
    """robots.txt を RFC 9309 のとおりに読む。標準ライブラリの urllib.robotparser は、User-agent の行のあとの空行で組を閉じ、
    最初に当たった行で決め（最も長く当たる行でなく）、* や $ を解さないため、禁止されている経路を許可と答えることがある
    （例：data.gov.gr の「User-agent: *」＋空行＋「Disallow: /api/」、data.gov.il の「Allow: /」のあとの「Disallow: /api/」）。"""

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


def get_json(url: str):
    if not allowed(url):  # 目録の API も robots.txt に従う
        return None
    r = fetch(url, DELAY_API, limit=50_000_000, timeout=90)
    for wait in (10, 30):  # つながらない・5xx：間をあけて2回まで取り直す
        if r.get("status") and r["status"] < 500:
            break
        time.sleep(wait)
        r = fetch(url, DELAY_API, limit=50_000_000, timeout=90)
    if not 200 <= (r.get("status") or 0) < 300:  # data.gov.sg の poll-download は 201
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


class UdataFrame:
    """udata（data.gouv.fr の型）：検索の format で絞り、何番目かを無作為に決めて登録を1つ取る（CKAN と同じ二段の抽出・同じ数え方）。"""

    def __init__(self, cc: str, api: str):
        self.cc, self.api = cc, api.rstrip("/")

    def _list(self, **q):
        return get_json(f"{self.api}/datasets/?{urllib.parse.urlencode(q)}")

    def strata(self) -> dict:
        out: dict[str, list] = {}
        for fam, fs in FAMILIES.items():
            for f in fs:
                d = self._list(format=f.lower(), page_size=1)
                if d is None:
                    raise RuntimeError("目録の検索が取れない")
                if d.get("total"):
                    out.setdefault(fam, []).append((f.lower(), int(d["total"])))
        return out

    def draw(self, fam: str, names: list, rnd: random.Random) -> dict | None:
        total = sum(c for _, c in names)
        nm, cnt = _pick_weighted(names, rnd)
        off = rnd.randrange(cnt)
        d = self._list(format=nm, page_size=1, page=off + 1)
        if not d or not d.get("data"):
            return None
        x = d["data"][0]
        res = [r for r in x.get("resources") or [] if FAMILY_OF.get(norm_format(r.get("format") or r.get("mime"))) == fam
               and str(r.get("url") or "").startswith(("http://", "https://"))]
        if not res:
            return {"empty": True}
        r = rnd.choice(res)
        return {"url": r["url"], "declared": norm_format(r.get("format") or r.get("mime")), "dataset": x.get("title"),
                "dataset_id": x.get("id"), "publisher": (x.get("organization") or {}).get("name", ""), "themes": [],
                "tags": (x.get("tags") or [])[:30], "frame_datasets": total, "in_dataset": len(res)}


class ListFrame:
    """検索で形式を絞れない目録（Socrata・DCAT の data.json・data.gov.sg）：登録の一覧を一度すべて取り（ページを順にめくる）、
    world/catalog/<国>.json.gz に残して、そこから選ぶ。層の大きさは、その群の書類を1つ以上持つ登録の数（CKAN と同じ数え方）。
    層の中は、登録を等しい確率で1つ選び、その登録の中の、その群の書類を1つ選ぶ（CKAN と同じ二段の抽出）。"""

    def __init__(self, cc: str, platform: str, api: str):
        self.cc, self.platform, self.api = cc, platform, api
        self._items = None

    # ---- 目録の型ごとの、登録の一覧の取り方。登録 = {id, title, publisher, themes, tags, res: [[url, 形式, 書き出し?]]}
    def _socrata(self):
        dom = urllib.parse.urlparse(self.api).netloc
        scroll = None
        while True:
            # scroll_id があると id の順に返る（最初は空）。前のページの最後の id から続ける
            q = {"domains": dom, "search_context": dom, "only": "dataset,file", "limit": 1000, "scroll_id": scroll or ""}
            d = get_json(f"https://{dom}/api/catalog/v1?{urllib.parse.urlencode(q)}")
            if d is None:  # 途中のページが取れない：欠けた一覧で枠を作らない
                raise RuntimeError("目録の一覧のページが取れない")
            rs = d.get("results") or []
            if not rs:
                return
            for x in rs:
                r, c = x.get("resource") or {}, x.get("classification") or {}
                if r.get("type") == "dataset":  # 表のデータ：目録が CSV に書き出す（発行者のファイルそのものではない）
                    res = [[f"https://{dom}/api/views/{r['id']}/rows.csv?accessType=DOWNLOAD", "CSV", True]]
                elif r.get("blob_mime_type"):  # 発行者が載せたファイル
                    res = [[f"https://{dom}/download/{r['id']}/{urllib.parse.quote(r['blob_mime_type'], safe='')}",
                            norm_format(r["blob_mime_type"]), False]]
                else:
                    continue
                yield {"id": r.get("id"), "title": r.get("name"), "publisher": r.get("attribution") or "",
                       "themes": [c["domain_category"]] if c.get("domain_category") else [],
                       "tags": (c.get("domain_tags") or c.get("tags") or [])[:30], "res": res}
            scroll = rs[-1]["resource"]["id"]

    def _dcat_json(self):
        d = get_json(self.api)
        if d is None:
            raise RuntimeError("目録の一覧が取れない")
        for x in d.get("dataset") or []:
            res = []
            for r in x.get("distribution") or []:
                f = norm_format(r.get("format") or r.get("mediaType"))
                if FAMILY_OF.get(f) is None and r.get("mediaType"):
                    f = norm_format(r.get("mediaType"))
                res.append([r.get("downloadURL") or r.get("accessURL"), f, False])
            pub = x.get("publisher") or {}
            yield {"id": x.get("identifier"), "title": x.get("title"), "publisher": pub.get("name", "") if isinstance(pub, dict) else str(pub),
                   "themes": x.get("theme") or [], "tags": (x.get("keyword") or [])[:30], "res": res}

    def _datagovsg(self):
        """シンガポール data.gov.sg の公開 API（v2）。登録ごとに形式が1つ。ファイルは poll-download が返す URL から取る。"""
        page, pages = 1, 1
        while page <= pages:
            d = get_json(f"{self.api.rstrip('/')}/v2/public/api/datasets?page={page}")
            if d is None or d.get("code") != 0:
                raise RuntimeError("目録の一覧のページが取れない")
            pages = int(d["data"].get("pages") or 1)
            for x in d["data"].get("datasets") or []:
                did = x.get("datasetId")
                yield {"id": did, "title": x.get("name"), "publisher": x.get("managedByAgencyName") or "", "themes": [], "tags": [],
                       "res": [[f"https://api-open.data.gov.sg/v1/public/api/datasets/{did}/poll-download", norm_format(x.get("format")), False]],
                       "resolve": "datagovsg"}
            page += 1

    def items(self) -> list:
        if self._items is not None:
            return self._items
        path = CATALOG / f"{self.cc}.json.gz"
        if path.exists():
            self._items = json.loads(gzip.decompress(path.read_bytes()))["items"]
            return self._items
        seen, out, listed = set(), [], 0
        for it in getattr(self, "_" + self.platform)():
            listed += 1
            it["res"] = [r for r in it["res"] if FAMILY_OF.get(r[1]) and str(r[0] or "").startswith(("http://", "https://"))]
            if it["res"] and it["id"] not in seen:
                seen.add(it["id"])
                out.append(it)
        out.sort(key=lambda it: str(it["id"]))  # 順を決めておく（同じ種なら同じ標本）
        CATALOG.mkdir(parents=True, exist_ok=True)
        path.write_bytes(gzip.compress(json.dumps({"source": self.api, "platform": self.platform, "built": time.strftime("%Y-%m-%d"),
                                                   "listed": listed, "items": out}, ensure_ascii=False).encode("utf-8"), mtime=0))
        self._items = out
        return out

    def strata(self) -> dict:
        out: dict[str, list] = {}
        for fam in FAMILIES:
            n = sum(1 for it in self.items() if any(FAMILY_OF[r[1]] == fam for r in it["res"]))
            if n:
                out[fam] = [(fam, n)]
        return out

    def draw(self, fam: str, names: list, rnd: random.Random) -> dict | None:
        pool = [it for it in self.items() if any(FAMILY_OF[r[1]] == fam for r in it["res"])]
        if not pool:
            return None
        it = pool[rnd.randrange(len(pool))]
        res = [r for r in it["res"] if FAMILY_OF[r[1]] == fam]
        url, f, export = rnd.choice(res)
        return {"url": url, **({"resolve": it["resolve"]} if it.get("resolve") else {}), "declared": f, "dataset": it.get("title"), "dataset_id": it.get("id"), "publisher": it.get("publisher", ""),
                "themes": it.get("themes") or [], "tags": it.get("tags") or [], "frame_datasets": len(pool), "in_dataset": len(res),
                **({"export": True} if export else {})}  # export：目録が表のデータから書き出したファイル


def make_frame(cc: str, rec: dict):
    pf = rec["platform"]
    if pf == "ckan":
        return CkanFrame(cc, rec["api"])
    if pf == "eu_hub":
        return EuFrame(cc)
    if pf == "udata":
        return UdataFrame(cc, rec["api"])
    return ListFrame(cc, pf, rec["api"])


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

EXT = {"PDF": ".pdf", "XLSX": ".xlsx", "XLS": ".xls", "ODS": ".ods", "DOCX": ".docx", "DOC": ".doc", "ODT": ".odt",
       "RTF": ".rtf", "CSV": ".csv"}
# 書類の名前の拡張子：目録の宣言（declared）がないか中身と違うことがあるので、URL の拡張子・Content-Type も見る
# （とどくは CSV・TSV・JSON・.txt を中身だけでは決めず、名前で見分けるため）
KNOWN_SUFFIX = {".pdf", ".xlsx", ".xlsm", ".xls", ".ods", ".docx", ".doc", ".odt", ".rtf", ".csv", ".tsv", ".txt",
                ".json", ".jsonl", ".pptx", ".zip", ".jtd"}
CTYPE_SUFFIX = {"text/csv": ".csv", "text/tab-separated-values": ".tsv", "text/plain": ".txt", "application/json": ".json",
                "application/pdf": ".pdf", "application/zip": ".zip"}


def suffix_for(s: dict, url: str, content_type: str | None) -> str:
    from urllib.parse import unquote, urlsplit

    ext = Path(unquote(urlsplit(url).path)).suffix.lower()
    if ext in KNOWN_SUFFIX:
        return ext
    ct = (content_type or "").split(";")[0].strip().lower()
    if ct in CTYPE_SUFFIX:
        return CTYPE_SUFFIX[ct]
    return EXT.get(s.get("declared", ""), "")


def _rate(h: str, salt: str) -> float:
    return int(hashlib.sha256((salt + h).encode()).hexdigest()[:8], 16) / 0xFFFFFFFF


def resolve_url(s: dict) -> str | None:
    """目録が書類の URL を直に示さず、問い合わせの答えで示す型（data.gov.sg の poll-download）。"""
    if s.get("resolve") == "datagovsg":
        d = get_json(s["url"])
        u = ((d or {}).get("data") or {}).get("url")
        return u if str(u or "").startswith("https://") else None
    return s["url"]


def run_sample(s: dict, tmp: Path) -> dict:
    rec = dict(s)
    if s.get("resolve"):
        u = resolve_url(s)
        if not u:
            rec["reach"] = "resolve_failed"
            return rec
        rec["url"], rec["catalog_url"] = u, s["url"]
    r, body = reach(rec["url"])
    rec.update(r)
    if body is None:
        return rec
    h = hashlib.sha256(body).hexdigest()
    rec.update(sha256=h, bytes=len(body))
    # 標本ごとの名前（同じ書類が別の標本で同時に選ばれても、ほかの標本の段階付けの途中で消さない）
    path = tmp / f"{h[:20]}-{s['id']}{suffix_for(s, rec['url'], rec.get('content_type'))}"
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
    pop = reachable()
    todo = [cc for cc in countries if cc not in frames and cc in pop]

    def one(cc):
        rec = pop[cc]
        try:
            st = make_frame(cc, rec).strata()
        except RuntimeError as e:  # 一覧が欠ける：この国の枠は作らない（次に動かしたとき作り直す）
            print(f"[枠] {cc} 作れない：{e}", flush=True)
            return cc, None
        return cc, {"name": rec["ja"], "source": rec["api"], "platform": rec["platform"], "strata": st,
                    "built": time.strftime("%Y-%m-%d")}

    with ThreadPoolExecutor(8) as ex:  # 国ごとに相手が違う（EU の統合目録は同じ相手なので、間は polite_wait が守る）
        for cc, f in ex.map(one, todo):
            if f is None:
                continue
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
    ap.add_argument("--countries", default="", help="既定：population.json で届く国のすべて")
    ap.add_argument("--frames-only", action="store_true", help="標本の枠だけを作る（書類は取らない）")
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
    countries = [c for c in a.countries.split(",") if c] or list(reachable())
    frames = build_frames(countries)
    if a.frames_only:
        return 0
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
    pop = reachable()

    def cell(cc: str, fam: str):
        names = frames[cc]["strata"].get(fam) or []
        if not names:
            return
        fr = make_frame(cc, pop[cc])
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
