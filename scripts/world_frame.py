"""世界調査の標本の枠を作る：各国の政府のサイトの一覧（ドメイン）と、Common Crawl の収集の中の、その国の政府の書類の枠と標本。
設計は docs/世界調査の設計.md。書類そのものは取らない（取るのは scripts/world_survey.py）。

1. domains：国ごとの政府のドメインの一覧を world/domains.json に作る。
   (a) 国別のトップレベルドメイン（ccTLD）の、政府用の第2レベル（gov.br・go.jp・gouv.fr・gob.mx など）。
       根拠は Public Suffix List（ICANN の節。TLD ごとに、登録機関の規則のページが書いてある）。意味の違う同じ綴り
       （go.it はゴリツィア県の記号、kg.kr は幼稚園）は除き、理由を記録する。PSL にない国の政府のドメイン
       （米国の .gov・ドイツの bund.de・スイスの admin.ch など）は、根拠の URL を付けて足す。
   (b) Wikidata：国（P17）があり、公式のサイト（P856）を持つ、政府の組織（Q2659904 の下位の類）・政府（Q7188）・
       行政区画（Q56061。市町村の公式のサイトは、その自治体のサイト）の項目。教育・病院・博物館・図書館・企業・放送・
       軍・研究所・在外公館の類は除く。SNS・ブログ・ホスティング（PSL の私的な節）・国際機関・2か国以上に現れるホストは除く。
2. frames：Common Crawl の一つの収集（CRAWL）の索引から、国 × 形式の群（pdf・sheet・text・csv）の枠と標本を作る。
   索引の全体（数百 GB）は取らない。索引の索引（cluster.idx、約 100 MB）で国のドメインに当たる索引のブロック
   （1ブロック 3,000 行）を決め、ブロックを無作為に選んで、そのブロックだけを範囲指定で取る（集落抽出の第1段）。
   ブロックが BLOCKS_MAX 以下の国は、すべてのブロックを読む（枠を数え尽くす）。ブロックは、ドメインの範囲の内側に
   まるごと入るもの（core）と、端のもの（edge）に分けて層にし、数に比べて割り当てる。
   選んだブロックの中の、その国のドメインの、状態 200・中身で見分けた形式（mime-detected）がその群の行を、
   同じ中身（digest）を1件にまとめて候補にする。候補ごとに乱数の鍵（種と digest のハッシュ）を付け、鍵の小さい順に
   SAMPLE_MAX 件までを標本の順番（rank）として world/sample.jsonl に書く。1回目は rank < 20、2回目で 50 まで足せる
   （選び直さない）。選ばれる確率 π = (m_h / B_h) × (n / 候補の数) を記録する。

使い方: PYTHONPATH=src .venv/bin/python scripts/world_frame.py domains
       PYTHONPATH=src .venv/bin/python scripts/world_frame.py frames [--countries JP,DE] [--blocks 30]
"""

from __future__ import annotations

import argparse
import bisect
import gzip
import http.client
import hashlib
import json
import random
import re
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "world"
CACHE = ROOT / "corpus" / "cc"  # 取ったものの控え（公開しない。.gitignore の corpus/）
CRAWL = "CC-MAIN-2026-43"  # 2026 年 10 月の収集（2026-10-05〜10-18）。https://index.commoncrawl.org/collinfo.json の最新
CC_DATA = "https://data.commoncrawl.org/"
WDQS = "https://query.wikidata.org/sparql"
PSL_URL = "https://publicsuffix.org/list/public_suffix_list.dat"
UA = "todoku-research/0.2 (accessibility research on public documents; polite crawler; https://github.com/Ono-Katsuki/todoku)"
SEED = 20261103
BLOCKS_MAX = 30  # 国ごとに読む索引のブロックの数の上限（それ以下の国は全部を読む）
SAMPLE_MAX = 50  # 層ごとに標本の順番を付けておく数（1回目 20・2回目 50）
PAUSE = 0.5  # 同じ相手への要求の間（秒）。一度に一つだけ

FAMILY_MIME = {
    "application/pdf": "pdf",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "sheet",
    "application/vnd.ms-excel": "sheet",
    "application/vnd.ms-excel.sheet.macroenabled.12": "sheet",
    "application/vnd.oasis.opendocument.spreadsheet": "sheet",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "text",
    "application/msword": "text",
    "application/vnd.oasis.opendocument.text": "text",
    "application/rtf": "text",
    "text/rtf": "text",
    "text/csv": "csv",
}
FAMILIES = ("pdf", "sheet", "text", "csv")

# ====================================================================== 通信（一度に一つ・間をあける・混んでいたら待って取り直す）
_net = threading.Lock()
_last: dict[str, float] = {}
_conns: dict[str, http.client.HTTPSConnection] = {}


def keepalive_get(url: str, headers: dict, timeout: int = 120) -> tuple[int, bytes, dict]:
    """同じ相手には一つのつながり（keep-alive）を使い回す（つなぎ直しの手間を相手にもかけない）。"""
    u = urllib.parse.urlsplit(url)
    for again in (0, 1):
        c = _conns.get(u.netloc)
        if c is None:
            c = _conns[u.netloc] = http.client.HTTPSConnection(u.netloc, timeout=timeout)
        try:
            c.request("GET", u.path + (f"?{u.query}" if u.query else ""), headers={"User-Agent": UA, **headers})
            r = c.getresponse()
            return r.status, r.read(), dict(r.getheaders())
        except (http.client.HTTPException, OSError):
            c.close()
            _conns.pop(u.netloc, None)
            if again:
                raise
    raise RuntimeError("unreachable")


def http_get(url: str, headers: dict | None = None, timeout: int = 120, data: bytes | None = None,
             retry_codes: tuple = (429, 500, 502, 503, 504)) -> bytes:
    host = urllib.parse.urlparse(url).netloc
    wait = 2.0
    for attempt in range(8):
        with _net:
            gap = _last.get(host, 0) + PAUSE - time.time()
            if gap > 0:
                time.sleep(gap)
            req = urllib.request.Request(url, data=data, headers={"User-Agent": UA, **(headers or {})})
            err = None
            try:
                with urllib.request.urlopen(req, timeout=timeout) as r:
                    body = r.read()
                _last[host] = time.time()
                return body
            except urllib.error.HTTPError as e:
                _last[host] = time.time()
                if e.code not in retry_codes:
                    raise
                err = e
                retry = e.headers.get("Retry-After") if e.headers else None
                if retry and retry.isdigit():
                    wait = max(wait, float(retry))
            except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as e:
                _last[host] = time.time()
                err = e
        if attempt == 7:
            raise err  # type: ignore[misc]
        print(f"  （{host} が混んでいる：{type(err).__name__} {getattr(err, 'code', '')}。{wait:.0f} 秒待つ）", flush=True)
        time.sleep(wait)
        wait = min(wait * 2, 300)
    raise RuntimeError("unreachable")


# ====================================================================== (a) 政府用の第2レベル（PSL）と、足すドメイン
GOV_LABELS = {"gov", "gob", "gouv", "go", "govt", "gv", "gub", "gc", "lg", "nic", "gok", "gop", "gos"}
# 同じ綴りでも政府の意味でないもの・意味を確かめられないもの（除く）
PSL_NOT_GOV = {
    "go.it": "イタリアの県の記号（ゴリツィア）",
    "kg.kr": "韓国の幼稚園の区分",
    "lg.ua": "ウクライナの地名（ルハーンシク）",
    "nic.za": "登録機関の名前",
    "nic.tj": "登録機関の名前",
    "go.tj": "政府の区分か確かめられない（gov.tj を使う）",
    "go.ci": "政府の区分か確かめられない（gouv.ci を使う）",
    "go.gov.br": "gov.br に含まれる",
    "gov.nc.tr": "北キプロス（国連の加盟国でない）",
}
TLD_COUNTRY = {"uk": "GB"}  # ccTLD と ISO 3166-1 の綴りが違うもの
# PSL にない、国の政府のドメイン（根拠つき）
EXTRA = {
    "US": [("gov", "https://get.gov/domains/requirements/", ".gov は米国の政府（連邦・州・地方・部族）だけが登録できる TLD")],
    "DE": [("bund.de", "https://www.bund.de/", "連邦の行政のドメイン（bund.de の下に連邦の機関のサイト）")],
    "CH": [("admin.ch", "https://www.admin.ch/", "連邦の行政のドメイン（admin.ch の下に連邦の機関のサイト）")],
    "BE": [("fgov.be", "https://www.belgium.be/", "連邦の公務のドメイン（fgov.be の下に連邦の機関のサイト）"),
           ("belgium.be", "https://www.belgium.be/", "連邦政府の案内のサイト")],
    "CA": [("canada.ca", "https://www.canada.ca/", "カナダ政府の共通のサイト")],
    "KR": [("korea.kr", "https://www.korea.kr/", "韓国政府の政策の案内のサイト（korea.kr の下に省庁のサイト）")],
}


def load_psl() -> tuple[list[tuple[str, str]], set[str]]:
    """ICANN の節の（suffix, その TLD の登録機関の URL）と、私的な節の suffix の集合。"""
    path = CACHE / "public_suffix_list.dat"
    if not path.exists():
        CACHE.mkdir(parents=True, exist_ok=True)
        path.write_bytes(http_get(PSL_URL))
    icann_txt, private_txt = path.read_text(encoding="utf-8").split("// ===END ICANN DOMAINS===")
    out, src = [], ""
    for line in icann_txt.splitlines():
        line = line.strip()
        if line.startswith("// ") and " : " in line and len(line.split(" : ")[0][3:].strip()) <= 3:
            src = line.split(" : ", 1)[1].strip()
            continue
        if line and not line.startswith("//"):
            out.append((line, src))
    private = {ln.strip().lstrip("*.!") for ln in private_txt.splitlines() if ln.strip() and not ln.startswith("//")}
    return out, private


def gov_suffixes(pop: list[dict]) -> dict[str, list[dict]]:
    psl, _ = load_psl()
    iso = {c["iso2"] for c in pop}
    out: dict[str, list[dict]] = defaultdict(list)
    for suffix, src in psl:
        p = suffix.split(".")
        if len(p) < 2 or p[0] not in GOV_LABELS or len(p[-1]) != 2:
            continue
        cc = TLD_COUNTRY.get(p[-1], p[-1].upper())
        if cc not in iso:
            continue
        if suffix in PSL_NOT_GOV:
            out[cc].append({"suffix": suffix, "source": "psl", "cite": src, "excluded": PSL_NOT_GOV[suffix]})
            continue
        out[cc].append({"suffix": suffix, "source": "psl", "cite": src,
                        "basis": f"Public Suffix List（ICANN の節）の .{p[-1]} の第2レベル。登録機関：{src}"})
    for cc, items in EXTRA.items():
        for suffix, cite, basis in items:
            out[cc].append({"suffix": suffix, "source": "curated", "cite": cite, "basis": basis})
    return out


# ====================================================================== (b) Wikidata
# 類（P31 で直に付いているもの）。下位の類をたどる問い合わせ（P279*）は国ごとでは時間切れになるので、決めた類の一覧を
# 類ごとに世界全体で問い合わせ（ページに分けて）、国（P17）で振り分ける。一覧の外の類（例：ドイツの市町村の細かい類）は入らない
WD_CLASSES = {
    "Q192350": "ministry", "Q327333": "government agency", "Q2659904": "government organization", "Q7188": "government",
    "Q1802419": "state government", "Q294414": "public office", "Q3505887": "city council",
    "Q3788231": "municipal government in Canada", "Q15284": "municipality", "Q515": "city", "Q1093829": "city in the United States",
    "Q484170": "commune of France", "Q747074": "comune of Italy", "Q2074737": "municipality of Spain",
    "Q3184121": "municipality of Brazil", "Q5153359": "municipality of the Czech Republic", "Q667509": "municipality of Austria",
    "Q6784672": "municipality of Slovakia", "Q562061": "market municipality", "Q24764": "municipality of the Philippines",
    "Q27676428": "municipality", "Q493522": "municipality of Belgium", "Q104841013": "hromada",
    "Q2039348": "municipality of the Netherlands", "Q856076": "municipality of Finland", "Q57058": "municipality of Croatia",
    "Q755707": "municipality of Norway", "Q1349648": "municipality of Greece", "Q127448": "municipality of Sweden",
    "Q1840161": "commune of Chile", "Q748149": "prefecture-level city of China", "Q3191695": "regency of Indonesia",
}
WD_PAGE = 20000


def wikidata_items() -> dict[str, tuple[str, set]]:
    """項目 → (国の項目, 公式のサイトの集合)。時間切れの類は飛ばして記録する（world/domains.json の sources.wikidata_skipped）。"""
    items: dict[str, tuple[str, set]] = {}
    for cls, label in WD_CLASSES.items():
        off = 0
        while True:
            name = f"class-{cls}-{off}"
            try:
                rows = _cached(name, lambda: sparql(
                    f"SELECT ?item ?site ?cty WHERE {{ ?item wdt:P31 wd:{cls} . ?item wdt:P856 ?site . ?item wdt:P17 ?cty . }} "
                    f"ORDER BY ?item LIMIT {WD_PAGE} OFFSET {off}"))
            except urllib.error.HTTPError as e:
                WD_SKIPPED[cls] = f"{label}：時間切れ（{e.code}）"
                print(f"  [Wikidata] {cls} {label} を飛ばす：{e}", flush=True)
                break
            for r in rows:
                it = _q(r["item"])
                cty, sites = items.setdefault(it, (_q(r["cty"]), set()))
                if _q(r["cty"]) != cty:
                    items[it] = ("", sites)  # 国が二つ以上：振り分けない
                sites.add(r["site"])
            print(f"  [Wikidata] {cls} {label}：{off + len(rows):,} 行", flush=True)
            if len(rows) < WD_PAGE:
                break
            off += WD_PAGE
    return items


WD_SKIPPED: dict[str, str] = {}
# SNS・動画・ブログ・共有のサイト（PSL の私的な節に入らないもの）
GENERIC = {
    "facebook.com", "fb.com", "twitter.com", "x.com", "instagram.com", "youtube.com", "youtu.be", "linkedin.com",
    "google.com", "goo.gl", "tiktok.com", "vk.com", "ok.ru", "t.me", "telegram.me", "medium.com", "wikipedia.org",
    "wikimedia.org", "wikidata.org", "archive.org", "flickr.com", "issuu.com", "scribd.com", "bit.ly", "linktr.ee",
    "about.me", "yahoo.com", "weibo.com", "qq.com", "baidu.com", "naver.com", "daum.net", "livejournal.com",
    "wordpress.com", "wix.com", "weebly.com", "jimdo.com", "tumblr.com", "europa.eu", "un.org", "who.int",
    "google.co.jp", "google.de", "line.me", "whatsapp.com", "pinterest.com",
}


def sparql(query: str) -> list[dict]:
    body = urllib.parse.urlencode({"query": query, "format": "json"}).encode()
    raw = http_get(WDQS, headers={"Accept": "application/sparql-results+json",
                                  "Content-Type": "application/x-www-form-urlencoded"}, data=body, timeout=70,
                   retry_codes=(429, 502, 503))  # 500・504 は問い合わせの時間切れ：取り直さず、分けて問い合わせる
    return [{k: v["value"] for k, v in b.items()} for b in json.loads(raw)["results"]["bindings"]]


def _cached(name: str, fn):
    path = CACHE / "wd" / f"{name}.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    v = fn()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(v, ensure_ascii=False), encoding="utf-8")
    return v


def _q(url: str) -> str:
    return url.rsplit("/", 1)[-1]


# 国の記号（P297）を持たないが、項目の国（P17）に使われる項目：オランダ（王国 Q29999 の中の国。市町村の P17 はこちら）
WD_COUNTRY_ALIAS = {"Q55": "NL"}


def country_items() -> dict[str, str]:
    """ISO 3166-1 の2文字 → 国の Wikidata の項目（今ある国・地域。なくなった国（P576）は除く）。"""
    rows = _cached("countries", lambda: sparql(
        "SELECT ?c ?iso WHERE { ?c wdt:P297 ?iso . FILTER NOT EXISTS { ?c wdt:P576 ?end } }"))
    out: dict[str, list[str]] = defaultdict(list)
    for r in rows:
        out[r["iso"].upper()].append(_q(r["c"]))
    # 同じ記号に項目が二つ以上あるとき：小さい番号（最初に作られた、その国そのものの項目）
    return {k: sorted(v, key=lambda q: int(q[1:]))[0] for k, v in out.items()}


def host_of(url: str) -> str | None:
    try:
        h = (urllib.parse.urlsplit(url.strip()).hostname or "").lower().rstrip(".")
    except ValueError:
        return None
    if not h or "." not in h or h.replace(".", "").isdigit() or not re.fullmatch(r"[a-z0-9.\-\u0080-\uffff]+", h):
        return None
    try:
        h = h.encode("idna").decode("ascii")
    except UnicodeError:
        return None
    return h[4:] if h.startswith("www.") else h


def under(host: str, suffix: str) -> bool:
    return host == suffix or host.endswith("." + suffix)


def build_domains(use_wikidata: bool = True) -> int:
    pop = json.loads((OUT / "population.json").read_text(encoding="utf-8"))["countries"]
    suffixes = gov_suffixes(pop)
    _, private = load_psl()
    cq = country_items() if use_wikidata else {}
    # 国ごとのホスト（Wikidata）
    host_c: dict[str, set] = defaultdict(set)
    host_item: dict[tuple, str] = {}
    stats = {}
    if use_wikidata:
        q2iso = {q: iso for iso, q in cq.items()}
        q2iso.update(WD_COUNTRY_ALIAS)
        for it, (cty, sites) in wikidata_items().items():
            iso2 = q2iso.get(cty)
            if not iso2:
                continue
            for site in sorted(sites):
                h = host_of(site)
                if h:
                    host_c[h].add(iso2)
                    host_item.setdefault((h, iso2), it)
                    stats.setdefault(iso2, {"items": set()})["items"].add(it)
        stats = {k: {"items": len(v["items"])} for k, v in stats.items()}
    all_suffix = [(s["suffix"], cc) for cc, ss in suffixes.items() for s in ss if "excluded" not in s]
    kept: dict[str, list] = defaultdict(list)
    drops: dict[str, Counter] = defaultdict(Counter)
    for h, ccs in host_c.items():
        parts = h.split(".")
        parents = {".".join(parts[k:]) for k in range(1, len(parts))}
        gov_of = {cc for s, cc in all_suffix if under(h, s)}
        if gov_of:
            why = "covered" if gov_of & ccs else "other_country"  # 政府用の第2レベルに入っている／ほかの国の第2レベルの下
        elif parents & private or h in private:
            why = "hosting"  # ホスティング・ブログ（PSL の私的な節）
        elif ".".join(parts[-2:]) in GENERIC or h in GENERIC or parents & GENERIC or h.endswith(".int"):
            why = "generic"  # SNS・共有のサイト・国際機関
        elif len(ccs) > 1:
            why = "shared"  # 2か国以上の項目に現れる
        else:
            why = None
        for cc in ccs:
            if why is None:
                kept[cc].append([h, host_item[(h, cc)]])
            else:
                drops[cc][why] += 1
    out = {}
    for c in pop:
        iso2 = c["iso2"]
        out[iso2] = {"name": c["ja"], "wikidata_country": cq.get(iso2), "suffixes": suffixes.get(iso2, []),
                     "hosts": sorted(kept.get(iso2, [])),
                     "wikidata": {**stats.get(iso2, {}), "hosts": len(kept.get(iso2, [])), "dropped": dict(drops.get(iso2, {}))}}
    doc = {"built": time.strftime("%Y-%m-%d"),
           "sources": {"psl": PSL_URL, "wikidata": WDQS,
                       "wikidata_classes": WD_CLASSES if use_wikidata else {}, "wikidata_skipped": WD_SKIPPED,
                       "gov_labels": sorted(GOV_LABELS), "psl_not_gov": PSL_NOT_GOV},
           "countries": out}
    (OUT / "domains.json").write_text(json.dumps(doc, ensure_ascii=False, indent=0), encoding="utf-8")
    n_s = sum(1 for v in out.values() if any("excluded" not in s for s in v["suffixes"]))
    n_h = sum(1 for v in out.values() if v["hosts"])
    print(f"国：政府用の第2レベルあり {n_s}・Wikidata のホストあり {n_h}・どちらか {sum(1 for v in out.values() if v['hosts'] or any('excluded' not in s for s in v['suffixes']))}")
    return 0


# ====================================================================== 2. 枠と標本（Common Crawl の索引）
def surt(host: str) -> str:
    return ",".join(reversed(host.split(".")))


class ClusterIndex:
    def __init__(self):
        path = CACHE / f"{CRAWL}.cluster.idx"
        if not path.exists():
            CACHE.mkdir(parents=True, exist_ok=True)
            print("索引の索引（cluster.idx）を取る", flush=True)
            path.write_bytes(http_get(f"{CC_DATA}cc-index/collections/{CRAWL}/indexes/cluster.idx", timeout=600))
        self.keys, self.loc = [], []
        with path.open(encoding="utf-8") as f:
            for line in f:
                k, fn, off, ln, _n = line.rstrip("\n").split("\t")
                self.keys.append(k.split(" ", 1)[0])
                self.loc.append((fn, int(off), int(ln)))

    def blocks(self, prefixes: list[str]) -> dict[int, str]:
        """接頭辞（SURT）の範囲に当たるブロック → core（まるごと範囲の内側）／edge。"""
        out: dict[int, str] = {}
        n = len(self.keys)
        for p in prefixes:
            for s, e in ((p + ")", p + ",￿"), (p + ":", p + ":￿")):
                lo = max(bisect.bisect_right(self.keys, s) - 1, 0)
                hi = bisect.bisect_right(self.keys, e) - 1
                for i in range(lo, hi + 1):
                    nxt = self.keys[i + 1] if i + 1 < n else "￿"
                    if nxt <= s:
                        continue
                    core = self.keys[i] >= s and nxt <= e
                    if core:
                        out[i] = "core"
                    else:
                        out.setdefault(i, "edge")
        return out

    def read(self, i: int) -> list[str]:
        fn, off, ln = self.loc[i]
        raw = cc_range(f"{CC_DATA}cc-index/collections/{CRAWL}/indexes/{fn}", off, ln)
        if len(raw) != ln:
            raise RuntimeError(f"ブロックの長さが違う {fn} {off}: {len(raw)} != {ln}")
        return gzip.decompress(raw).decode("utf-8", "replace").splitlines()


def cc_range(url: str, off: int, ln: int) -> bytes:
    """Common Crawl の保存庫から範囲指定で取る。一度に一つ・PAUSE 秒の間・混んでいたら待って、間を倍に延ばして取り直す。"""
    wait = 2.0
    for attempt in range(8):
        with _net:
            gap = _last.get("cc", 0) + PAUSE - time.time()
            if gap > 0:
                time.sleep(gap)
            try:
                st, body, hd = keepalive_get(url, {"Range": f"bytes={off}-{off + ln - 1}"})
                err = f"HTTP {st}"
            except (http.client.HTTPException, OSError) as e:
                st, body, hd, err = 0, b"", {}, type(e).__name__
            _last["cc"] = time.time()
        if st == 206 and len(body) == ln:
            return body
        if st not in (0, 206, 429, 500, 502, 503, 504):
            raise RuntimeError(f"{url} {off}: {err}")
        if attempt == 7:
            raise RuntimeError(f"{url} {off}: {err}（取り直しても取れない）")
        print(f"  （Common Crawl が混んでいる：{err}。{wait:.0f} 秒待つ）", flush=True)
        time.sleep(wait)
        wait = min(wait * 2, 300)
    raise RuntimeError("unreachable")


def line_host(key: str) -> str:
    h = key.split(")", 1)[0]
    return h.split(":", 1)[0]


# 中身で見分けた形式が大まかなもの（Tika が OLE・OOXML・ZIP までしか決められない。多くは途中で切れた保存）：
# URL の拡張子で群を決め、family_by: extension と記録する
GENERIC_MIME = {"application/x-tika-msoffice", "application/x-tika-ooxml", "application/octet-stream", "application/zip",
                "text/plain"}
EXT_FAMILY = {".pdf": "pdf", ".xls": "sheet", ".xlsx": "sheet", ".xlsm": "sheet", ".ods": "sheet", ".doc": "text",
              ".docx": "text", ".odt": "text", ".rtf": "text", ".csv": "csv"}
FRAME_VERSION = 2  # 群の決め方を変えたら上げる（作り直す）


def family_of(rec: dict) -> tuple[str | None, str]:
    md = (rec.get("mime-detected") or "").split(";")[0].strip().lower()
    if md in FAMILY_MIME:
        return FAMILY_MIME[md], "mime"
    if md in GENERIC_MIME:
        path = urllib.parse.unquote(urllib.parse.urlsplit(rec.get("url") or "").path).lower()
        ext = path[path.rfind("."):] if "." in path.rsplit("/", 1)[-1] else ""
        fam = EXT_FAMILY.get(ext)
        if fam and not (md == "text/plain" and fam != "csv"):
            return fam, "extension"
    return None, ""


def _key(digest: str) -> str:
    return hashlib.sha256(f"{SEED}:{digest}".encode()).hexdigest()


def build_frames(countries: list[str], blocks_max: int, source: str = "suffix") -> int:
    """source：suffix（政府用の区分と足したドメイン）か wikidata（Wikidata のホストで、区分の下にないもの）。
    二つは重ならない（Wikidata のホストのうち区分の下のものは除いてある）ので、国 × 群 × 出典を別の層にする。
    Wikidata の層を足しても、区分の層の標本は変わらない。"""
    part = "countries" if source == "suffix" else "wikidata"
    doms = json.loads((OUT / "domains.json").read_text(encoding="utf-8"))["countries"]
    fpath, spath = OUT / "frames.json", OUT / "sample.jsonl"
    frames = json.loads(fpath.read_text(encoding="utf-8")) if fpath.exists() else {"crawl": CRAWL, "countries": {}}
    frames.setdefault(part, {})
    samples: dict[tuple, list] = defaultdict(list)
    if spath.exists():
        for line in spath.read_text(encoding="utf-8").splitlines():
            r = json.loads(line)
            samples[(r["country"], r.get("domain_source", "suffix"))].append(line)
    idx = ClusterIndex()
    print(f"索引のブロック {len(idx.keys):,}", flush=True)
    todo = countries or list(doms)
    for cc in todo:
        d = doms[cc]
        own_s = [surt(s["suffix"]) for s in d["suffixes"] if "excluded" not in s] if source == "suffix" else []
        own_h = {surt(h) for h, _ in d["hosts"]} if source == "wikidata" else set()
        prefixes = own_s + sorted(own_h)
        old = frames[part].get(cc)
        # 同じドメインの一覧・同じ上限で作ってあれば作り直さない
        if (old and old.get("blocks_max") == blocks_max and old.get("prefixes") == len(prefixes)
                and old.get("version") == FRAME_VERSION):
            continue

        def mine(host_surt: str) -> bool:
            if any(host_surt == p or host_surt.startswith(p + ",") for p in own_s):
                return True
            parts = host_surt.split(",")
            return any(",".join(parts[:k]) in own_h for k in range(2, len(parts) + 1))

        blocks = idx.blocks(prefixes)
        strata = {h: sorted(i for i, t in blocks.items() if t == h) for h in ("core", "edge")}
        B = {h: len(v) for h, v in strata.items()}
        total = sum(B.values())
        rec = {"name": d["name"], "prefixes": len(prefixes), "blocks": B, "blocks_max": blocks_max, "version": FRAME_VERSION}
        if total == 0:
            rec.update(families={}, read={"core": 0, "edge": 0}, exact=True)
            frames[part][cc] = rec
            print(f"[枠] {cc} {d['name']}：ドメイン {len(prefixes)}・収集の中に索引のブロックがない", flush=True)
            _save(frames, samples, fpath, spath)
            continue
        # 割り当て：数に比べて（上限 blocks_max。層が空でなければ1つ以上）
        if total <= blocks_max:
            m = dict(B)
        else:
            m = {h: (max(1, round(blocks_max * B[h] / total)) if B[h] else 0) for h in B}
        rnd = random.Random(f"{SEED}:{cc}" + ("" if source == "suffix" else f":{source}"))
        chosen = {h: (rnd.sample(strata[h], m[h]) if m[h] < B[h] else list(strata[h])) for h in B}
        cand: dict[str, dict] = {f: {} for f in FAMILIES}
        lines_in = Counter()
        mimes = Counter()
        for h in ("core", "edge"):
            for i in sorted(chosen[h]):
                for line in idx.read(i):
                    try:
                        key, ts, js = line.split(" ", 2)
                    except ValueError:
                        continue
                    if not mine(line_host(key)):
                        continue
                    lines_in[h] += 1
                    r = json.loads(js)
                    if r.get("status") != "200":
                        continue
                    fam, by = family_of(r)
                    md = (r.get("mime-detected") or "").lower()
                    if md.startswith("application/") and not md.endswith(("html", "xml", "json", "javascript")):
                        mimes[md] += 1
                    if not fam or not r.get("digest"):
                        continue
                    if r["digest"] in cand[fam]:
                        continue
                    cand[fam][r["digest"]] = {"url": r["url"], "timestamp": ts, "mime": r.get("mime"),
                                              "mime_detected": r.get("mime-detected"), "digest": r["digest"],
                                              "filename": r["filename"], "offset": int(r["offset"]),
                                              "length": int(r["length"]), "block": i, "block_stratum": h,
                                              **({"family_by": by} if by != "mime" else {}),
                                              **({"truncated": r["truncated"]} if r.get("truncated") else {})}
        fams = {}
        out_lines = []
        for fam in FAMILIES:
            pool = list(cand[fam].values())
            by_h = Counter(c["block_stratum"] for c in pool)
            est = sum(B[h] / m[h] * by_h[h] for h in B if m[h])
            exact = all(m[h] == B[h] for h in B)
            pool.sort(key=lambda c: _key(c["digest"]))
            fams[fam] = {"frame": round(est), "pool": len(pool), "exact": exact,
                         "pool_by_block_stratum": dict(by_h),
                         "truncated_in_pool": sum(1 for c in pool if c.get("truncated"))}
            for rank, c in enumerate(pool[:SAMPLE_MAX]):
                n_take = min(SAMPLE_MAX, len(pool))
                pi = (m[c["block_stratum"]] / B[c["block_stratum"]]) * (n_take / len(pool))
                sid = f"{cc}-{fam}-{rank:03d}" if source == "suffix" else f"{cc}-{fam}-w{rank:03d}"
                out_lines.append(json.dumps({"id": sid, "country": cc, "family": fam, "rank": rank, "domain_source": source,
                                             "pi": pi, "crawl": CRAWL, **c}, ensure_ascii=False))
        rec.update(read=m, exact=all(m[h] == B[h] for h in B), lines_in_frame=dict(lines_in), families=fams,
                   other_document_mimes=dict(mimes.most_common(8)), built=time.strftime("%Y-%m-%d"))
        frames[part][cc] = rec
        samples[(cc, source)] = out_lines
        print(f"[枠] {cc} {d['name']}：ブロック {B['core']}+{B['edge']}（読んだ {m['core']}+{m['edge']}）・"
              + "・".join(f"{f} {fams[f]['frame']:,}（候補 {fams[f]['pool']}）" for f in FAMILIES), flush=True)
        _save(frames, samples, fpath, spath)
    return 0


RARE = ("sheet", "text")  # 申請の様式が多い群。1回目のブロックでは候補が少ない
RARE_TARGET = 20  # 層ごとにこの数の候補が集まるまで、ブロックを足す
EXTRA_MAX = 90  # 国 × 出典ごとに足すブロックの上限（1回目の 30 と合わせて 120 まで）


def extend_rare(blocks_max: int, countries: list[str], source: str) -> int:
    """表計算・文書の層の候補が RARE_TARGET 未満の国で、まだ読んでいないブロックを、別の種で無作為な順に並べ、
    候補がそろうか、ブロックが尽きるか、EXTRA_MAX に届くまで、順に読む（1回目と合わせて、層の中の単純無作為抽出になる）。
    足したブロックは表計算・文書の候補にだけ使う（PDF・CSV の標本は 1回目のブロックのまま変えない）。
    枠の推定と選ばれる確率は、群ごとに読んだブロックの数（blocks_read）で出す。止め方が観測に依るので、推定はわずかに偏りうる。"""
    part = "countries" if source == "suffix" else "wikidata"
    doms = json.loads((OUT / "domains.json").read_text(encoding="utf-8"))["countries"]
    fpath, spath = OUT / "frames.json", OUT / "sample.jsonl"
    frames = json.loads(fpath.read_text(encoding="utf-8"))
    samples: dict[tuple, list] = defaultdict(list)
    for line in spath.read_text(encoding="utf-8").splitlines():
        r = json.loads(line)
        samples[(r["country"], r.get("domain_source", "suffix"))].append(line)
    idx = ClusterIndex()
    for cc in countries or list(frames[part]):
        rec = frames[part].get(cc)
        fams = (rec or {}).get("families") or {}
        if not fams or rec.get("extend") is not None:
            continue
        need = [f for f in RARE if fams[f]["pool"] < RARE_TARGET]
        d = doms[cc]
        own_s = [surt(x["suffix"]) for x in d["suffixes"] if "excluded" not in x] if source == "suffix" else []
        own_h = {surt(h) for h, _ in d["hosts"]} if source == "wikidata" else set()
        blocks = idx.blocks(own_s + sorted(own_h))
        strata = {h: sorted(i for i, t in blocks.items() if t == h) for h in ("core", "edge")}
        B = {h: len(v) for h, v in strata.items()}
        m = rec["read"]
        # 1回目に読んだブロック（build_frames と同じ種で、同じ選び方）
        rnd = random.Random(f"{SEED}:{cc}" + ("" if source == "suffix" else f":{source}"))
        chosen = {h: (rnd.sample(strata[h], m[h]) if m[h] < B[h] else list(strata[h])) for h in B}
        first = set(chosen["core"]) | set(chosen["edge"])
        if any(json.loads(x)["block"] not in first for x in samples[(cc, source)]):
            raise RuntimeError(f"{cc} {source}：1回目のブロックを作り直せない（枠を作り直してから足す）")
        rest = [(i, h) for h in B for i in strata[h] if i not in first]
        random.Random(f"{SEED}:{cc}:{source}:extend").shuffle(rest)
        if not need or not rest:
            rec["extend"] = {"read": {"core": 0, "edge": 0}, "stop": "enough" if not need else "exhausted"}
            _save(frames, samples, fpath, spath)
            continue

        def mine(host_surt: str) -> bool:
            if any(host_surt == p or host_surt.startswith(p + ",") for p in own_s):
                return True
            parts = host_surt.split(",")
            return any(",".join(parts[:k]) in own_h for k in range(2, len(parts) + 1))

        old = {f: [json.loads(x) for x in samples[(cc, source)] if json.loads(x)["family"] == f] for f in RARE}
        cand = {f: {c["digest"]: c for c in old[f]} for f in RARE}  # 候補が 50 未満の層は、標本の行が候補のすべて
        extra = Counter()
        stop = "exhausted"
        for i, h in rest:
            if all(len(cand[f]) >= RARE_TARGET for f in RARE):
                stop = "enough"
                break
            if sum(extra.values()) >= EXTRA_MAX:
                stop = "cap"
                break
            extra[h] += 1
            for line in idx.read(i):
                try:
                    key, ts, js = line.split(" ", 2)
                except ValueError:
                    continue
                if not mine(line_host(key)):
                    continue
                r = json.loads(js)
                if r.get("status") != "200":
                    continue
                fam, by = family_of(r)
                if fam not in RARE or not r.get("digest") or r["digest"] in cand[fam]:
                    continue
                cand[fam][r["digest"]] = {"url": r["url"], "timestamp": ts, "mime": r.get("mime"),
                                          "mime_detected": r.get("mime-detected"), "digest": r["digest"],
                                          "filename": r["filename"], "offset": int(r["offset"]),
                                          "length": int(r["length"]), "block": i, "block_stratum": h,
                                          **({"family_by": by} if by != "mime" else {}),
                                          **({"truncated": r["truncated"]} if r.get("truncated") else {})}
        m2 = {h: m[h] + extra[h] for h in B}
        lines = [x for x in samples[(cc, source)] if json.loads(x)["family"] not in RARE]
        for fam in RARE:
            pool = [{k: v for k, v in c.items() if k not in ("id", "country", "family", "rank", "domain_source", "pi", "crawl")}
                    for c in cand[fam].values()]
            by_h = Counter(c["block_stratum"] for c in pool)
            pool.sort(key=lambda c: _key(c["digest"]))
            fams[fam].update(frame=round(sum(B[h] / m2[h] * by_h[h] for h in B if m2[h])), pool=len(pool),
                             exact=all(m2[h] == B[h] for h in B), pool_by_block_stratum=dict(by_h), blocks_read=m2,
                             truncated_in_pool=sum(1 for c in pool if c.get("truncated")))
            n_take = min(SAMPLE_MAX, len(pool))
            for rank, c in enumerate(pool[:SAMPLE_MAX]):
                pi = (m2[c["block_stratum"]] / B[c["block_stratum"]]) * (n_take / len(pool))
                sid = f"{cc}-{fam}-{rank:03d}" if source == "suffix" else f"{cc}-{fam}-w{rank:03d}"
                lines.append(json.dumps({"id": sid, "country": cc, "family": fam, "rank": rank, "domain_source": source,
                                         "pi": pi, "crawl": CRAWL, **c}, ensure_ascii=False))
        samples[(cc, source)] = lines
        rec["extend"] = {"read": dict(extra), "stop": stop}
        print(f"[足す] {cc} {source}：ブロック +{sum(extra.values())}（{stop}）・"
              + "・".join(f"{f} 候補 {fams[f]['pool']}（推定 {fams[f]['frame']:,}）" for f in RARE), flush=True)
        _save(frames, samples, fpath, spath)
    return 0


def _save(frames, samples, fpath, spath):
    fpath.write_text(json.dumps(frames, ensure_ascii=False, indent=1), encoding="utf-8")
    tmp = spath.with_suffix(".tmp")
    tmp.write_text("".join(line + "\n" for k in sorted(samples) for line in samples[k]), encoding="utf-8")
    tmp.replace(spath)


OUTCOMES = {
    "reachable": "政府のサイトの書類が収集にあり、標本の枠を作れる",
    "no_documents": "政府のサイトは収集にあるが、PDF・表計算・文書・CSV の書類が見つからない",
    "not_in_crawl": "政府のドメインは分かるが、この収集に入っていない",
    "no_domains": "出典（Public Suffix List・足したドメイン・Wikidata）に政府のドメインが見つからない",
}


def update_population() -> None:
    """world/population.json の国ごとの結果を、この調査の枠から書き直す（目録の調査の項目は消す）。"""
    path = OUT / "population.json"
    doc = json.loads(path.read_text(encoding="utf-8"))
    doms = json.loads((OUT / "domains.json").read_text(encoding="utf-8"))["countries"]
    fr = json.loads((OUT / "frames.json").read_text(encoding="utf-8"))
    base = ("iso2", "iso3", "num", "ja", "en", "region", "un_member")
    out = []
    for c in doc["countries"]:
        cc = c["iso2"]
        d, f = doms.get(cc, {}), fr["countries"].get(cc, {})
        fw = fr.get("wikidata", {}).get(cc, {})
        suf = [x["suffix"] for x in d.get("suffixes", []) if "excluded" not in x]
        hosts = len(d.get("hosts", []))
        fams, pools = Counter(), Counter()
        for part in (f, fw):
            for k, v in (part.get("families") or {}).items():
                fams[k] += v["frame"]
                pools[k] += v["pool"]
        blocks_read = sum((f.get("read") or {}).values()) + sum((fw.get("read") or {}).values())
        if any(pools.values()):
            outcome, reason = "reachable", ""
        elif sum((f.get("blocks") or {}).values()) + sum((fw.get("blocks") or {}).values()):
            outcome = "no_documents"
            reason = f"政府のサイトは収集にあるが、書類が見つからない（索引のブロック {blocks_read} を読んだ）"
        elif suf or hosts:
            outcome, reason = "not_in_crawl", "政府のドメインは分かるが、この収集に入っていない"
        else:
            outcome, reason = "no_domains", "出典に政府のドメインが見つからない"
        out.append({**{k: c[k] for k in base}, "outcome": outcome, "reason": reason,
                    "domains": {"suffixes": suf, "wikidata_hosts": hosts},
                    "frame": {"crawl": fr["crawl"], "families": dict(fams), "pool": dict(pools),
                              "suffix": {k: f.get(k) for k in ("blocks", "read", "exact")} if f else None,
                              "wikidata": {k: fw.get(k) for k in ("blocks", "read", "exact")} if fw else None}})
    doc = {"population": doc["population"], "checked_on": time.strftime("%Y-%m-%d"), "crawl": fr["crawl"],
           "outcomes": OUTCOMES, "countries": out}
    path.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    n = Counter(c["outcome"] for c in out if c["un_member"])
    print("国連の加盟国：" + "・".join(f"{k} {v}" for k, v in n.items()))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("stage", choices=["wikidata", "domains", "frames", "extend", "population"])
    ap.add_argument("--countries", default="")
    ap.add_argument("--blocks", type=int, default=BLOCKS_MAX)
    ap.add_argument("--no-wikidata", action="store_true", help="domains：政府用の第2レベルと足すドメインだけ（Wikidata を使わない）")
    ap.add_argument("--source", choices=["suffix", "wikidata"], default="suffix", help="frames：どの出典のドメインの層を作るか")
    a = ap.parse_args(argv)
    OUT.mkdir(exist_ok=True)
    if a.stage == "wikidata":  # 問い合わせて控えに置くだけ（domains がこの控えを使う）
        items = wikidata_items()
        print(f"項目 {len(items):,}・飛ばした類 {WD_SKIPPED}")
        return 0
    if a.stage == "domains":
        return build_domains(not a.no_wikidata)
    if a.stage == "extend":
        rc = extend_rare(a.blocks, [c for c in a.countries.split(",") if c], a.source)
        update_population()
        return rc
    if a.stage == "population":
        update_population()
        return 0
    rc = build_frames([c for c in a.countries.split(",") if c], a.blocks, a.source)
    update_population()
    return rc


if __name__ == "__main__":
    sys.exit(main())
