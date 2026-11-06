#!/usr/bin/env python3
"""未見の日本の様式（無作為でない別の集まり。世界の推定に足さない）。規則は docs/再評価の計画.md の 10 節。

check：候補の掲載ページのうち、まだ見ていないホストで robots.txt が許すものを、sources/jp_forms_pages.tsv に書く。掲載ページの HTML だけを読む。様式は取らない。
fetch：一覧のページから様式を取る（タグ eval-freeze-2 のあとだけ）。ページごとにリンクを SHA-256(種:URL) の順に並べ、ホストごとに 12 件まで・合計 60 件まで。
       robots.txt の禁止・CAPTCHA・ログイン・401・403・429 が 1 度でもあれば、そのホストをやめる。見た書類の digest は除く。中身は開かない（SHA-256 と先頭の形だけ）。
       記録は 出力/sources_jp_forms.tsv（sources/ と同じ列）。
使い方: jp_forms.py check | jp_forms.py fetch --out D"""

from __future__ import annotations

import argparse
import base64
import hashlib
import json
import re
import sys
import time
import urllib.parse as U
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(ROOT / "scripts"))
PAGES = ROOT / "sources" / "jp_forms_pages.tsv"
SEED = 20261107
PER_HOST, TOTAL = 12, 60
EXTS = {".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pdf"}
MAGIC = {".xls": b"\xd0\xcf\x11\xe0", ".doc": b"\xd0\xcf\x11\xe0", ".pdf": b"%PDF"}
REFUSED = re.compile(r"^(robots|captcha|login|http_401|http_403|http_429)$")
CANDIDATES = [  # 検索で見つけた掲載ページ（2026-11-06。様式は開いていない）
    ("https://www.pref.fukushima.lg.jp/sec/21035a/sinnsei.html", "福島県"),
    ("https://www.city.sakai.lg.jp/kenko/fukushikaigo/shogaifukushi/sougou_shienhou/syougaisa-bisuyousiki.html", "堺市"),
    ("https://www.city.akita.lg.jp/kurashi/shogaifukushi/1006068/1011789.html", "秋田市"),
    ("https://www.city.kitakyushu.lg.jp/contents/924_11163.html", "北九州市"),
    ("https://www.pref.osaka.lg.jp/o070080/toshimiryoku/osakathemuseum/kyoso_program/hojoyoko.html", "大阪府"),
    ("https://www.city.saitama.lg.jp/002/003/003/001/002/p003093.html", "さいたま市"),
    ("https://www.city.minoh.lg.jp/kaigo/ji/osirase/cm_sinseisho/cm_daikoushinsei.html", "箕面市"),
    ("https://www.city.fukuoka.lg.jp/fukushi/kaigohoken/shisei/5-010307_2.html", "福岡市"),
    ("https://www.city.kashiwa.lg.jp/kourei/shiseijoho/forms/kojin/kaigo/index.html", "柏市"),
    ("https://www.city.yokohama.lg.jp/isogo/kenko-iryo-fukushi/fukushi_kaigo/koreisha_kaigo/kaigo-hoken/zigyousyamukeannai.html", "横浜市"),
]
_URL = re.compile(rb"https?://([A-Za-z0-9.-]+)")


def host_of(url: str) -> str:
    return U.urlsplit(url).netloc.lower()


def seen_hosts(extra_dirs: list[str]) -> set:
    """これまでに見たホスト：sources/・tests/sources.json・world/sample.jsonl と、渡したフォルダの .json・.jsonl・.md の中の URL。"""
    out = set()
    files = [*(ROOT / "sources").glob("*"), ROOT / "tests" / "sources.json", ROOT / "world" / "sample.jsonl"]
    for d in extra_dirs:
        files += [p for p in Path(d).rglob("*") if p.suffix in (".json", ".jsonl", ".md", ".txt", ".tsv") and p.is_file()]
    for f in files:
        if f == PAGES or not f.is_file():
            continue
        try:
            out |= {m.decode("ascii", "replace").lower() for m in _URL.findall(f.read_bytes())}
        except OSError:
            pass
    return out


def links_of(page: str, html: bytes, ctype: str) -> list[str]:
    import collect_forms as cf

    text = html.decode("utf-8", "replace")
    m = re.search(r'charset=["\']?([\w-]+)', ctype + text[:2000], re.I)
    if m and m[1].lower() not in ("utf-8", "utf8"):
        try:
            text = html.decode(m[1], "replace")
        except LookupError:
            pass
    lp = cf.Links()
    lp.feed(text)
    return sorted({U.urljoin(page, h) for h, _t in lp.links if Path(U.urlsplit(U.urljoin(page, h)).path).suffix.lower() in EXTS})


def cmd_check(a) -> int:
    import world_survey as ws

    seen = seen_hosts(a.seen_dir)
    rows = []
    for page, pub in CANDIDATES:
        host = host_of(page)
        if host in seen:
            print(f"除く（見たホスト）{host}")
            continue
        if not ws.allowed(page):
            print(f"除く（robots）{host}")
            continue
        r = ws.fetch(page, 2.0)
        if r.get("status") != 200:
            print(f"除く（HTTP {r.get('status')}）{host}")
            continue
        links = links_of(page, r["body"], r.get("type", ""))
        print(f"入れる {host} 様式のリンク {len(links)}")
        if links:
            rows.append((page, pub, host, str(len(links))))
    PAGES.write_text("# 未見の日本の様式の掲載ページ（様式を取る前に固定。docs/再評価の計画.md の 10 節）\n# 掲載ページ\t発行者\tホスト\t確かめたときのリンクの数\n" + "".join("\t".join(x) + "\n" for x in rows), encoding="utf-8")
    print(f"ページ {len(rows)}")
    return 0


def cmd_fetch(a) -> int:
    import stage2_select as SS
    import world_survey as ws

    out = Path(a.out)
    (out / "docs").mkdir(parents=True, exist_ok=True)
    excl = SS.exclusion_digests(a.subset)
    log = []
    files = []
    for ln in PAGES.read_text(encoding="utf-8").splitlines():
        if ln.startswith("#") or not ln.strip():
            continue
        page, pub, host, _n = ln.split("\t")
        if len(files) >= TOTAL:
            break
        if not ws.allowed(page):
            log.append((page, "", "robots のページ"))
            continue
        r = ws.fetch(page, 2.0)
        if r.get("status") != 200:
            log.append((page, "", f"HTTP {r.get('status')}"))
            continue
        urls = sorted(links_of(page, r["body"], r.get("type", "")), key=lambda u: hashlib.sha256(f"{SEED}:{u}".encode()).hexdigest())
        took = 0
        for u in urls:
            if took >= PER_HOST or len(files) >= TOTAL:
                break
            st, body = ws.live_copy(u)
            why = st.get("live")
            if body is None:
                log.append((page, u, why))
                if REFUSED.match(str(why)):  # 断られたら、そのホストをやめる
                    break
                continue
            ext = Path(U.urlsplit(u).path).suffix.lower()
            if not body.startswith(MAGIC.get(ext, b"PK")):
                log.append((page, u, "形式が違う"))
                continue
            if base64.b32encode(hashlib.sha1(body).digest()).decode() in excl:
                log.append((page, u, "見た書類"))
                continue
            h = hashlib.sha256(body).hexdigest()
            p = out / "docs" / f"{h[:16]}{ext}"
            p.write_bytes(body)
            files.append({"path": str(p), "url": u, "page": page, "publisher": pub, "sha256": h, "bytes": len(body), "retrieved": time.strftime("%F")})
            took += 1
    (out / "files.txt").write_text("".join(f["path"] + "\n" for f in files), encoding="utf-8")
    (out / "fetch_log.json").write_text(json.dumps(log, ensure_ascii=False, indent=0), encoding="utf-8")
    (out / "sources_jp_forms.tsv").write_text("path\turl\tpage\tpublisher\tsha256\tretrieved\n" + "".join(f"{Path(f['path']).name}\t{f['url']}\t{f['page']}\t{f['publisher']}\t{f['sha256']}\t{f['retrieved']}\n" for f in files), encoding="utf-8")
    print(f"様式 {len(files)} 件・飛ばした {len(log)} 件")
    return 0 if files else 3


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["check", "fetch"])
    ap.add_argument("--out")
    ap.add_argument("--subset")
    ap.add_argument("--seen-dir", action="append", default=[])
    a = ap.parse_args()
    return {"check": cmd_check, "fetch": cmd_fetch}[a.cmd](a)


if __name__ == "__main__":
    raise SystemExit(main())
