"""世界地図のための調査：各国の政府のオープンデータ目録（CKAN）から、公開書類（PDF・Excel・Word）を無作為に抜き出し、
AI エージェントにとってどこまで「届く」かを4段階で採点する。

段階（その書類で、どこまで進めたか）
  0 届かない   取得できない（robots.txt で禁止・403・ログイン・CAPTCHA・別のページ・時間切れ）
  1 届く       取得できたが、AI が読めない（暗号化・画像だけ・文字に直せない字・壊れたファイル）
  2 読める     読めるが、入力欄に正しく書き込めない（または書き込みの証明・読み戻しが合わない）
  3 操作できる 読めて、すべての入力欄に書けて、読み戻せる（入力欄のない書類は、読めればここまで）
  4 検証できる 書き込みの証明に加え、数式がすべて計算し直せて一致し、マクロがない
（コンセプトの「届く／読める／操作できる／検証できる」に、届かない を足したもの）

礼儀：robots.txt を守り、同じ相手には 2 秒以上の間隔を空ける。CAPTCHA やボット対策は破らず「止められた」と記録する。
取得した書類は corpus/world/（公開しない）に置き、公開するのは採点と取得元の URL だけ（world/results.json）。

使い方: uv run python scripts/world_survey.py --per-format 8 [--portals JP,GB]
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
FILES = ROOT / "corpus" / "world"
UA = "todoku-research/0.1 (accessibility research on public forms; polite crawler)"
DELAY = 2.0
MAX_BYTES = 25 * 1024 * 1024

PORTALS = {
    "JP": ("日本", "https://data.e-gov.go.jp/data"),
    "GB": ("英国", "https://data.gov.uk"),
    "CA": ("カナダ", "https://open.canada.ca/data/en"),
    "AU": ("オーストラリア", "https://data.gov.au/data"),
    "IE": ("アイルランド", "https://data.gov.ie"),
    "CH": ("スイス", "https://ckan.opendata.swiss"),
    "DE": ("ドイツ", "https://ckan.govdata.de"),
    "NL": ("オランダ", "https://data.overheid.nl/data"),
    "IT": ("イタリア", "https://www.dati.gov.it/opendata"),
    "UY": ("ウルグアイ", "https://catalogodatos.gub.uy"),
    "CL": ("チリ", "https://datos.gob.cl"),
    # 目録の API に届かなかった国も、記録として残す
    "US": ("米国", "https://catalog.data.gov"),
    "MX": ("メキシコ", "https://www.datos.gob.mx"),
    "ES": ("スペイン", "https://datos.gob.es"),
    "AR": ("アルゼンチン", "https://datos.gob.ar"),
    "BR": ("ブラジル", "https://dados.gov.br/dados"),
}
FORMATS = ["PDF", "XLSX", "XLS", "DOCX", "DOC"]

_last: dict[str, float] = {}
_lock = threading.Lock()
_robots: dict[str, urllib.robotparser.RobotFileParser | None] = {}


def polite_wait(host: str) -> None:
    with _lock:
        now = time.time()
        wait = _last.get(host, 0) + DELAY - now
        _last[host] = max(now, _last.get(host, 0) + DELAY)
    if wait > 0:
        time.sleep(wait)


def allowed(url: str) -> bool:
    p = urllib.parse.urlparse(url)
    base = f"{p.scheme}://{p.netloc}"
    with _lock:
        known = base in _robots
    if not known:
        rp = urllib.robotparser.RobotFileParser()
        rp.set_url(base + "/robots.txt")
        try:
            polite_wait(p.netloc)
            rp.read()
        except Exception:  # noqa: BLE001
            rp = None
        with _lock:
            _robots[base] = rp
    rp = _robots[base]
    return True if rp is None else rp.can_fetch(UA, url)


def fetch(url: str, limit: int = MAX_BYTES, timeout: int = 40) -> dict:
    p = urllib.parse.urlparse(url)
    polite_wait(p.netloc)
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
        return {"status": e.code, "final_url": url, "type": e.headers.get("Content-Type", "") if e.headers else "", "body": body}
    except Exception as e:  # noqa: BLE001
        return {"status": 0, "error": f"{type(e).__name__}: {str(e)[:120]}", "body": b""}


def api(base: str, action: str, **params) -> dict | None:
    q = urllib.parse.urlencode({k: (json.dumps(v) if isinstance(v, (list, dict)) else v) for k, v in params.items()})
    r = fetch(f"{base}/api/3/action/{action}?{q}", limit=5_000_000, timeout=40)
    if r.get("status") != 200:
        return {"_error": r.get("error") or f"HTTP {r.get('status')}"}
    try:
        d = json.loads(r["body"])
        return d.get("result") if d.get("success", True) else {"_error": "success=false"}
    except json.JSONDecodeError:
        return {"_error": "JSON でない応答"}


# ====================================================================== 目録
def catalog(cc: str, base: str, per_format: int, rnd: random.Random) -> dict:
    out = {"code": cc, "portal": base, "formats": {}, "samples": []}
    head = api(base, "package_search", rows=0, **{"facet.field": ["res_format"], "facet.limit": 200})
    if not head or "_error" in head:
        out["catalog_error"] = (head or {}).get("_error", "応答なし")
        return out
    out["datasets"] = head.get("count", 0)
    items = head.get("search_facets", {}).get("res_format", {}).get("items", [])
    names: dict[str, list[tuple[str, int]]] = {}
    for it in items:
        nm = str(it.get("name", ""))
        up = nm.upper().rsplit("/", 1)[-1]
        for f in FORMATS:
            if up == f or up == f"." + f:
                names.setdefault(f, []).append((nm, int(it.get("count", 0))))
    for f in FORMATS:
        lst = names.get(f, [])
        out["formats"][f] = sum(c for _, c in lst)
        if not lst:
            continue
        nm, cnt = max(lst, key=lambda x: x[1])
        offsets = rnd.sample(range(cnt), min(per_format, cnt))
        for off in offsets:
            r = api(base, "package_search", rows=1, start=off, fq=f'res_format:"{nm}"')
            if not r or "_error" in r or not r.get("results"):
                continue
            pkg = r["results"][0]
            res = [x for x in pkg.get("resources", []) if str(x.get("format", "")).upper().rsplit("/", 1)[-1].lstrip(".") == f]
            if not res:
                continue
            x = rnd.choice(res)
            url = x.get("url") or ""
            if not url.startswith(("http://", "https://")):
                continue
            out["samples"].append({"format": f, "url": url, "dataset": pkg.get("title") or pkg.get("name"),
                                   "publisher": (pkg.get("organization") or {}).get("title", "")})
    return out


# ====================================================================== 取得と採点
CAPTCHA = (b"captcha", b"cf-chl", b"challenge-platform", b"Just a moment", b"Attention Required", b"hcaptcha", b"recaptcha")


def reach(sample: dict) -> tuple[dict, bytes | None]:
    url = sample["url"]
    if not allowed(url):
        return {"reach": "robots", "level": 0, "why": "robots.txt で取得が禁止されています"}, None
    r = fetch(url)
    st = r.get("status", 0)
    body = r.get("body", b"")
    if st == 0:
        return {"reach": "error", "level": 0, "why": r.get("error", "取得できませんでした")}, None
    if st in (401, 403):
        if any(m in body[:200_000] for m in CAPTCHA):
            return {"reach": "captcha", "level": 0, "why": f"CAPTCHA・ボット対策で止められました（HTTP {st}）"}, None
        return {"reach": f"http_{st}", "level": 0, "why": f"アクセスを断られました（HTTP {st}）"}, None
    if st >= 400:
        return {"reach": f"http_{st}", "level": 0, "why": f"HTTP {st}"}, None
    if r.get("too_large"):
        return {"reach": "too_large", "level": 0, "why": f"{MAX_BYTES // 1024 // 1024}MB を超えるので、取得を見送りました"}, None
    head = body[:2048].lstrip().lower()
    if head.startswith((b"<!doctype html", b"<html")) or b"<html" in head[:512]:
        if any(m in body[:200_000] for m in CAPTCHA):
            return {"reach": "captcha", "level": 0, "why": "CAPTCHA・ボット対策で止められました"}, None
        if b'type="password"' in body or any(w in r.get("final_url", "").lower() for w in ("login", "signin", "auth")):
            return {"reach": "login", "level": 0, "why": "ログインが必要です"}, None
        return {"reach": "html", "level": 0, "why": "書類ではなく、ウェブページが返りました"}, None
    return {"reach": "ok"}, body


def score(path: Path, timeout: int = 180) -> dict:
    """別のプロセスで、とどくの検査をかける（時間がかかりすぎる書類で全体が止まらないように）。"""
    code = (
        "import json,sys\n"
        "from todoku.book import kind_of\n"
        "from todoku import api\n"
        "p=sys.argv[1]\n"
        "b=open(p,'rb').read()\n"
        "k=kind_of(b)\n"
        "out={'kind':k}\n"
        "if k not in api.KIND_JA:\n"
        "    out['error']='encrypted' if k=='encrypted' else 'unknown'\n"
        "    print(json.dumps(out));sys.exit()\n"
        "au=api.audit(p)\n"
        "out['checks']={c['name']:c['ok'] for c in au['checks']}\n"
        "out['details']={c['name']:c['detail'][:200] for c in au['checks'] if not c['ok']}\n"
        "try:\n"
        "    s=api.summary(p); out['fields']=s['fields']\n"
        "except Exception as e:\n"
        "    out['fields']=None\n"
        "if k in ('xlsx','xls'):\n"
        "    from todoku.book import open_book\n"
        "    from todoku.xlsx.calc import verify\n"
        "    wb=open_book(b,'x'); v=verify(wb).summary(); out['formulas']=v; out['macros']=wb.has_macros\n"
        "print(json.dumps(out,ensure_ascii=False,default=str))\n"
    )
    try:
        r = subprocess.run([sys.executable, "-c", code, str(path)], capture_output=True, timeout=timeout, cwd=ROOT)
        lines = r.stdout.decode("utf-8", "replace").strip().splitlines()
        if not lines:
            return {"error": "crash", "stderr": r.stderr.decode("utf-8", "replace")[-300:]}
        return json.loads(lines[-1])
    except subprocess.TimeoutExpired:
        return {"error": "timeout"}


def ladder(sc: dict) -> tuple[int, str]:
    if sc.get("error") == "encrypted":
        return 1, "パスワードで暗号化されています"
    if sc.get("error") in ("unknown",):
        return 1, "申請書類の形式（Excel・PDF・Word）として読めませんでした"
    if sc.get("error"):
        return 1, f"処理できませんでした（{sc['error']}）"
    ch = sc.get("checks", {})
    read_names = [n for n in ch if n.startswith("読み落とし") or n in ("文字に直せない字", "画像だけのページ", "窓の大きさ")]
    if not all(ch[n] for n in read_names):
        bad = [n for n in read_names if not ch[n]]
        return 1, "読めない：" + "・".join(bad)
    write_names = [n for n in ("全入力欄への書き込み", "読み戻し") if n in ch]
    if not all(ch[n] for n in write_names):
        return 2, "書き込めない欄、または読み戻せない欄があります"
    f = sc.get("formulas") or {}
    if f.get("mismatch") or f.get("unsupported") or sc.get("macros") or ch.get("マクロ") is False:
        why = []
        if f.get("mismatch") or f.get("unsupported"):
            why.append("計算し直せない数式")
        if sc.get("macros") or ch.get("マクロ") is False:
            why.append("マクロ")
        return 3, "検証できない：" + "・".join(why)
    other = [n for n, ok in ch.items() if not ok and n not in read_names + write_names]
    if other:
        return 3, "確認が必要：" + "・".join(other)
    return 4, ""


def run_sample(cc: str, s: dict) -> dict:
    r, body = reach(s)
    rec = {"country": cc, **{k: s[k] for k in ("format", "url", "dataset", "publisher")}, **r}
    if body is None:
        return rec
    h = hashlib.sha256(body).hexdigest()
    rec["sha256"] = h
    rec["bytes"] = len(body)
    ext = "." + s["format"].lower()
    d = FILES / cc
    d.mkdir(parents=True, exist_ok=True)
    path = d / f"{h[:16]}{ext}"
    path.write_bytes(body)
    sc = score(path)
    lv, why = ladder(sc)
    rec.update({"level": lv, "why": why, "kind": sc.get("kind"), "fields": sc.get("fields"),
                "checks": sc.get("checks"), "formulas": sc.get("formulas")})
    return rec


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--per-format", type=int, default=8)
    ap.add_argument("--portals", default=",".join(PORTALS))
    ap.add_argument("--seed", type=int, default=2026)
    ap.add_argument("--jobs", type=int, default=6)
    a = ap.parse_args(argv)
    OUT.mkdir(parents=True, exist_ok=True)
    rnd = random.Random(a.seed)
    cats = {}
    for cc in a.portals.split(","):
        name, base = PORTALS[cc]
        c = catalog(cc, base, a.per_format, random.Random(rnd.random()))
        c["name"] = name
        cats[cc] = c
        print(f"[目録] {cc} {name}: " + (f"届かない（{c['catalog_error']}）" if "catalog_error" in c else
                                       f"データセット {c.get('datasets', 0):,} 件・抜き出し {len(c['samples'])} 件 {c['formats']}"), flush=True)
    (OUT / "catalogs.json").write_text(json.dumps(cats, ensure_ascii=False, indent=1), encoding="utf-8")
    jobs = [(cc, s) for cc, c in cats.items() for s in c.get("samples", [])]
    results = []
    with ThreadPoolExecutor(a.jobs) as ex:
        for rec in ex.map(lambda x: run_sample(*x), jobs):
            results.append(rec)
            print(f"[{rec['level']}] {rec['country']} {rec['format']} {rec.get('reach')} {rec.get('why', '')[:60]} {rec['url'][:90]}", flush=True)
            (OUT / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"終わり：{len(results)} 件")
    return 0


if __name__ == "__main__":
    sys.exit(main())
