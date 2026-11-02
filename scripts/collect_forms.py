"""申請様式の掲載ページから、様式のファイルを礼儀正しく集める（汎用性の検査用）。

- robots.txt に従う。止められたページは取らず「止められた」と記録する。
- 1件ごとに間隔をあける。CAPTCHAやボット対策は破らない。
- 集めた様式は再配布しない（corpus/ は .gitignore）。取得元のURLとSHA-256を記録する。
- 様式の形式（Excel・Word・PDFなど）の内訳も数える。どの形式に対応すべきかの根拠にする。

使い方: uv run python scripts/collect_forms.py [--per-page 8] [--out corpus/wide]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import time
import urllib.parse
import urllib.request
import urllib.robotparser
from datetime import date
from html.parser import HTMLParser
from pathlib import Path

UA = "todoku-research/0.1 (accessibility research on public forms; polite crawler)"
ROOT = Path(__file__).resolve().parent.parent
EXT = re.compile(r"\.(xlsx|xlsm|xls|docx|doc|pdf|ods|odt|zip)$", re.I)
EXCEL = {"xlsx", "xlsm"}
DELAY = 2.0


class Links(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links: list[tuple[str, str]] = []
        self._href = None
        self._text: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            self._href = dict(attrs).get("href")
            self._text = []

    def handle_data(self, data):
        if self._href is not None:
            self._text.append(data)

    def handle_endtag(self, tag):
        if tag == "a" and self._href:
            self.links.append((self._href, "".join(self._text).strip()))
            self._href = None


_robots: dict[str, urllib.robotparser.RobotFileParser | None] = {}


def allowed(url: str) -> bool:
    p = urllib.parse.urlsplit(url)
    base = f"{p.scheme}://{p.netloc}"
    if base not in _robots:
        rp = urllib.robotparser.RobotFileParser()
        try:
            req = urllib.request.Request(base + "/robots.txt", headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=20) as r:
                rp.parse(r.read().decode("utf-8", errors="replace").splitlines())
        except Exception:  # noqa: BLE001
            rp = None  # robots.txt がない・読めない：一般的な扱いに従い、許可とみなす
        _robots[base] = rp
    rp = _robots[base]
    return True if rp is None else rp.can_fetch(UA, url)


def get(url: str, limit: int = 20_000_000) -> tuple[bytes, str]:
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept-Language": "ja,en;q=0.8"})
    with urllib.request.urlopen(req, timeout=60) as r:
        data = r.read(limit + 1)
        if len(data) > limit:
            raise ValueError("大きすぎる")
        return data, r.headers.get("Content-Type", "")


def ext_of(url: str) -> str | None:
    path = urllib.parse.urlsplit(url).path
    m = EXT.search(path)
    return m[1].lower() if m else None


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--per-page", type=int, default=8)
    ap.add_argument("--out", default=str(ROOT / "corpus" / "wide"))
    a = ap.parse_args(argv)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    sources = json.loads((ROOT / "tests" / "sources.json").read_text(encoding="utf-8"))["sources"]
    index_path = out / "index.json"
    index = json.loads(index_path.read_text(encoding="utf-8")) if index_path.exists() else {"pages": [], "files": []}
    done_pages = {p["page"] for p in index["pages"]}
    have = {f["url"] for f in index["files"]}
    for src in sources:
        page = src["page"]
        if page in done_pages:
            continue
        rec = {**src, "status": "", "formats": {}, "retrieved": date.today().isoformat()}
        print(f"- {src['publisher']} {src['domain']}: {page}", flush=True)
        try:
            if not allowed(page):
                rec["status"] = "止められた（robots.txt）"
                index["pages"].append(rec)
                continue
            if ext_of(page) in EXCEL:
                links = [(page, "")]
            else:
                html, ctype = get(page)
                time.sleep(DELAY)
                text = html.decode("utf-8", errors="replace")
                m = re.search(r'charset=["\']?([\w-]+)', ctype + text[:2000], re.I)
                if m and m[1].lower() not in ("utf-8", "utf8"):
                    try:
                        text = html.decode(m[1], errors="replace")
                    except LookupError:
                        pass
                lp = Links()
                lp.feed(text)
                links = [(urllib.parse.urljoin(page, h), t) for h, t in lp.links]
            seen = set()
            files = []
            for url, label in links:
                e = ext_of(url)
                if not e or url in seen:
                    continue
                seen.add(url)
                rec["formats"][e] = rec["formats"].get(e, 0) + 1
                if e in EXCEL:
                    files.append((url, label))
            rec["status"] = "ok"
            for url, label in files[: a.per_page]:
                if url in have:
                    continue
                if not allowed(url):
                    index["files"].append({"url": url, "page": page, "status": "止められた（robots.txt）"})
                    continue
                try:
                    data, _ = get(url)
                except Exception as e:  # noqa: BLE001
                    index["files"].append({"url": url, "page": page, "status": f"取得できない（{type(e).__name__}）"})
                    time.sleep(DELAY)
                    continue
                time.sleep(DELAY)
                if data[:2] != b"PK":
                    index["files"].append({"url": url, "page": page, "status": "Excelではなかった"})
                    continue
                host = urllib.parse.urlsplit(url).netloc
                name = Path(urllib.parse.unquote(urllib.parse.urlsplit(url).path)).name
                dest = out / host / name
                dest.parent.mkdir(parents=True, exist_ok=True)
                if dest.exists() and dest.read_bytes() != data:
                    dest = dest.with_name(f"{dest.stem}_{hashlib.sha256(url.encode()).hexdigest()[:6]}{dest.suffix}")
                dest.write_bytes(data)
                index["files"].append({"url": url, "page": page, "publisher": src["publisher"], "country": src["country"],
                                       "domain": src["domain"], "label": label, "path": str(dest.relative_to(out)),
                                       "sha256": hashlib.sha256(data).hexdigest(), "bytes": len(data), "status": "ok",
                                       "retrieved": date.today().isoformat()})
                have.add(url)
                print(f"    取得: {name}（{len(data):,} バイト）", flush=True)
        except Exception as e:  # noqa: BLE001
            rec["status"] = f"取得できない（{type(e).__name__}: {e}）"[:200]
        index["pages"].append(rec)
        index_path.write_text(json.dumps(index, ensure_ascii=False, indent=1), encoding="utf-8")
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=1), encoding="utf-8")
    fmts: dict[str, int] = {}
    for p in index["pages"]:
        for k, v in p.get("formats", {}).items():
            fmts[k] = fmts.get(k, 0) + v
    ok = [f for f in index["files"] if f.get("status") == "ok"]
    print(f"ページ {len(index['pages'])}・取得した様式 {len(ok)} 件。掲載されていた様式の形式: {fmts}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
