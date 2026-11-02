"""CKANの目録から公開Excelを少数取得し、手元の検証用コーパスにする。

robots.txt を守り、アクセスの間隔を空ける。取得した書類は corpus/ に置き、公開しない（.gitignore）。
使い方: uv run python scripts/fetch_sample.py --portal https://data.e-gov.go.jp/data --n 30
"""

from __future__ import annotations

import argparse
import hashlib
import json
import random
import time
import urllib.parse
import urllib.request
import urllib.robotparser
from pathlib import Path

UA = "todoku-research/0.1 (+https://github.com/; accessibility research, polite)"
ROOT = Path(__file__).resolve().parent.parent / "corpus"


def get(url: str, timeout: int = 30) -> tuple[int, bytes]:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status, r.read(20 * 1024 * 1024)
    except urllib.error.HTTPError as e:
        return e.code, b""


_robots: dict[str, urllib.robotparser.RobotFileParser | None] = {}


def allowed(url: str) -> bool:
    p = urllib.parse.urlparse(url)
    base = f"{p.scheme}://{p.netloc}"
    if base not in _robots:
        rp = urllib.robotparser.RobotFileParser()
        rp.set_url(base + "/robots.txt")
        try:
            rp.read()
            _robots[base] = rp
        except Exception:  # noqa: BLE001
            _robots[base] = None
    rp = _robots[base]
    return True if rp is None else rp.can_fetch(UA, url)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--portal", default="https://data.e-gov.go.jp/data")
    ap.add_argument("--fmt", default="XLSX")
    ap.add_argument("--n", type=int, default=30)
    ap.add_argument("--seed", type=int, default=1)
    a = ap.parse_args()
    out = ROOT / urllib.parse.urlparse(a.portal).netloc
    out.mkdir(parents=True, exist_ok=True)
    api = f"{a.portal}/api/3/action/resource_search?query=format:{a.fmt}&limit=1"
    total = json.loads(get(api)[1])["result"]["count"]
    rnd = random.Random(a.seed)
    offsets = sorted(rnd.sample(range(total), min(a.n, total)))
    log = []
    for off in offsets:
        meta = json.loads(get(f"{a.portal}/api/3/action/resource_search?query=format:{a.fmt}&limit=1&offset={off}")[1])
        res = meta["result"]["results"][0]
        url = res.get("download_url") or res["url"]
        entry = {"url": url, "name": res.get("name"), "id": res["id"]}
        if not allowed(url):
            entry["status"] = "robots"
        else:
            code, body = get(url)
            entry["http"] = code
            if code == 200 and body:
                h = hashlib.sha256(body).hexdigest()[:16]
                ext = Path(urllib.parse.urlparse(url).path).suffix.lower() or ".bin"
                (out / f"{h}{ext}").write_bytes(body)
                entry["file"] = f"{h}{ext}"
                entry["status"] = "ok"
            else:
                entry["status"] = f"http_{code}"
        log.append(entry)
        print(entry.get("status"), url)
        time.sleep(1.5)
    (out / "index.json").write_text(json.dumps(log, ensure_ascii=False, indent=1))


if __name__ == "__main__":
    main()
