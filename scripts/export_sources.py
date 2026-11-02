"""集めた書類の出どころ（URL・掲載ページ・発行者・SHA-256）を、corpus/ の記録から sources/ に書き出す。

書類そのものはリポジトリに入れない（corpus/ は .gitignore）。ここに書き出すのは、同じものをもう一度取りに行けるための記録だけ。
使い方: uv run python scripts/export_sources.py
"""

from __future__ import annotations

import json
import shutil
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CORPUS = ROOT / "corpus"
OUT = ROOT / "sources"
COLS = ("path", "url", "page", "publisher", "sha256", "retrieved")


def _rows(items) -> list[dict]:
    if isinstance(items, dict) and isinstance(items.get("files"), list):
        items = items["files"]  # 集めた様式の記録：{pages: 掲載ページ, files: ファイル}
    if isinstance(items, dict):
        items = [dict(v, path=k) if isinstance(v, dict) else {"path": k} for k, v in items.items()]
    out = []
    for it in items:
        if not isinstance(it, dict) or it.get("status") not in (None, "ok", 200):
            continue
        url = it.get("url") or it.get("example") or ""
        if not url:
            continue
        out.append({"path": it.get("path") or it.get("file") or "", "url": url, "page": it.get("page", ""),
                    "publisher": it.get("publisher", ""), "sha256": it.get("sha256", ""),
                    "retrieved": it.get("retrieved", "")})
    return out


def main() -> int:
    OUT.mkdir(exist_ok=True)
    written = []
    for d in sorted(p for p in CORPUS.iterdir() if p.is_dir()):
        for sub in [d] + sorted(p for p in d.iterdir() if p.is_dir()):
            src = sub / "SOURCES.txt"
            name = sub.relative_to(CORPUS).as_posix().replace("/", "__")
            if src.is_file():
                shutil.copyfile(src, OUT / f"{name}.txt")
                written.append(f"{name}.txt")
                continue
            for idx in ("index.json", "manifest.json"):
                f = sub / idx
                if f.is_file() and sub == d:
                    rows = _rows(json.loads(f.read_text(encoding="utf-8")))
                    if rows:
                        lines = ["\t".join(COLS)] + ["\t".join(str(r[c]) for c in COLS) for r in rows]
                        (OUT / f"{name}.tsv").write_text("\n".join(lines) + "\n", encoding="utf-8")
                        written.append(f"{name}.tsv（{len(rows)} 件）")
                    break
    print("書き出しました：\n  " + "\n  ".join(written))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
