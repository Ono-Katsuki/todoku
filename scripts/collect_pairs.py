"""申請様式の掲載ページから、「記入例」とその元の白紙の様式の組を、礼儀正しく集める（入力欄の見落としの検査用）。

- 記入例：リンクの文字かファイル名に「記入例・記載例・作成例・記入見本・見本」などがあるもの
- 元の様式：同じページの同じ形式のリンクで、記入例の語を除いたリンクの文字がいちばん似ているもの（すぐ前のリンクを優先）
- robots.txt に従い、1件ごとに間隔をあける。集めたファイルは再配布しない（corpus/pairs/ は公開しない）。

使い方: uv run python scripts/collect_pairs.py [--extra-pages 追加のページの一覧.json]
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import re
import sys
import urllib.parse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import time  # noqa: E402

from collect_forms import Links, allowed  # noqa: E402
from collect_forms import get as _get  # noqa: E402

DELAY = 2.0


def get(url: str, limit: int = 20_000_000):
    """取得するたびに間隔をあける。"""
    try:
        return _get(url, limit)
    finally:
        time.sleep(DELAY)

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "corpus" / "pairs"
DOC = re.compile(r"\.(xlsx|xlsm|xls|docx|doc|pdf)$", re.I)
EXAMPLE_TEXT = re.compile(r"記入例|記載例|作成例|記入見本|記載見本|見本|入力例|例）|（例|\(例\)")
EXAMPLE_FILE = re.compile(r"(kisai|kinyu|kinyuu|kinyuurei|kisairei|sakuseirei|mihon|example|sample)", re.I)
STRIP = re.compile(r"記入例|記載例|作成例|記入見本|記載見本|見本|入力例|[（(]?例[）)]?|【|】|\[|\]|（|）|\(|\)|PDF|Word|Excel|ワード|エクセル|\d+KB|\d+MB|[\s　:：・-]")


def is_example(text: str, url: str) -> bool:
    name = Path(urllib.parse.urlparse(url).path).name
    return bool(EXAMPLE_TEXT.search(text or "") or EXAMPLE_FILE.search(name))


def key(text: str, url: str) -> str:
    name = Path(urllib.parse.urlparse(url).path).stem
    t = STRIP.sub("", text or "")
    n = EXAMPLE_FILE.sub("", re.sub(r"[_\-]", "", name))
    return t or n


def pairs_on(page: str) -> list[dict]:
    body, _ = get(page, limit=5_000_000)
    p = Links()
    p.feed(body.decode("utf-8", errors="replace"))
    links = []
    for href, text in p.links:
        url = urllib.parse.urljoin(page, href)
        if DOC.search(urllib.parse.urlparse(url).path):
            links.append((url, " ".join((text or "").split())))
    out = []
    for i, (url, text) in enumerate(links):
        if not is_example(text, url):
            continue
        ext = Path(urllib.parse.urlparse(url).path).suffix.lower()
        k = key(text, url)
        best, best_score = None, 0.0
        for j, (u2, t2) in enumerate(links):
            if j == i or is_example(t2, u2) or Path(urllib.parse.urlparse(u2).path).suffix.lower() != ext:
                continue
            s = difflib.SequenceMatcher(None, k, key(t2, u2)).ratio()
            if j == i - 1:
                s += 0.25  # すぐ前のリンクは元の様式であることが多い
            if s > best_score:
                best, best_score = (u2, t2), s
        if best is not None and best_score >= 0.55:
            out.append({"example": url, "example_text": text, "blank": best[0], "blank_text": best[1],
                        "score": round(best_score, 2), "page": page})
    return out


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--sources", default=str(ROOT / "tests" / "sources.json"))
    ap.add_argument("--extra-pages", default=None)
    ap.add_argument("--max", type=int, default=200)
    a = ap.parse_args(argv)
    pages = [s["page"] for s in json.loads(Path(a.sources).read_text(encoding="utf-8"))["sources"]]
    if a.extra_pages:
        pages += json.loads(Path(a.extra_pages).read_text(encoding="utf-8"))
    OUT.mkdir(parents=True, exist_ok=True)
    index = []
    seen = set()
    for page in pages:
        if not page.endswith((".html", ".htm", "/")) and DOC.search(page):
            continue
        if not allowed(page):
            print(f"- 止められた（robots.txt）: {page}")
            index.append({"page": page, "blocked": "robots"})
            continue
        try:
            found = pairs_on(page)
        except Exception as e:  # noqa: BLE001
            print(f"- 取れなかった: {page} {type(e).__name__}")
            index.append({"page": page, "error": f"{type(e).__name__}: {e}"[:200]})
            continue
        print(f"- {page}: 記入例の組 {len(found)}")
        for pr in found:
            if (pr["blank"], pr["example"]) in seen or len([x for x in index if "blank_file" in x]) >= a.max:
                continue
            seen.add((pr["blank"], pr["example"]))
            host = urllib.parse.urlparse(pr["example"]).netloc
            d = OUT / host
            d.mkdir(parents=True, exist_ok=True)
            ok = True
            for role in ("blank", "example"):
                url = pr[role]
                if not allowed(url):
                    pr[f"{role}_blocked"] = "robots"
                    ok = False
                    break
                try:
                    body, _ctype = get(url)
                except Exception as e:  # noqa: BLE001
                    pr[f"{role}_error"] = f"{type(e).__name__}"
                    ok = False
                    break
                h = hashlib.sha256(body).hexdigest()
                ext = Path(urllib.parse.urlparse(url).path).suffix.lower()
                name = f"{h[:12]}_{role}{ext}"
                (d / name).write_bytes(body)
                pr[f"{role}_file"] = f"{host}/{name}"
                pr[f"{role}_sha256"] = h
            if ok:
                print(f"    組: {pr['blank_text'][:30]} ↔ {pr['example_text'][:30]}")
            index.append(pr)
        (OUT / "index.json").write_text(json.dumps(index, ensure_ascii=False, indent=1), encoding="utf-8")
    n = len([x for x in index if "blank_file" in x and "example_file" in x])
    print(f"終わり：記入例の組 {n} 組")
    return 0


if __name__ == "__main__":
    sys.exit(main())
