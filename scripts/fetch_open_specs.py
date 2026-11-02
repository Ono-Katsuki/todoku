"""CSV・JSON・RTF・OpenDocument の公式の仕様書を取得する（corpus/specs/、リポジトリには入れない・公開しない）。

とどくの CSV・JSON・RTF・OpenDocument の読み書きは、この仕様書から作った：
- RFC 4180（CSV）・RFC 8259（JSON）・RFC 6901（JSON Pointer）・JSON Lines の慣習
- Microsoft の Rich Text Format（RTF）Specification 1.9.1
- OASIS OpenDocument 1.3（第2部 パッケージ・第3部 スキーマ・第4部 数式）
使い方: uv run python scripts/fetch_open_specs.py
"""

from __future__ import annotations

import sys
import time
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "corpus" / "specs"
UA = "todoku-research/0.1 (accessibility research on public forms; polite crawler)"
SPECS = {
    "RFC4180.txt": "https://www.rfc-editor.org/rfc/rfc4180.txt",
    "RFC8259.txt": "https://www.rfc-editor.org/rfc/rfc8259.txt",
    "RFC6901.txt": "https://www.rfc-editor.org/rfc/rfc6901.txt",
    "MSFT-RTF.pdf": "https://interoperability.blob.core.windows.net/files/Archive_References/%5bMSFT-RTF%5d.pdf",
    "ODF13-part2-packages.html": "https://docs.oasis-open.org/office/OpenDocument/v1.3/os/part2-packages/OpenDocument-v1.3-os-part2-packages.html",
    "ODF13-part3-schema.html": "https://docs.oasis-open.org/office/OpenDocument/v1.3/os/part3-schema/OpenDocument-v1.3-os-part3-schema.html",
    "ODF13-part4-formula.html": "https://docs.oasis-open.org/office/OpenDocument/v1.3/os/part4-formula/OpenDocument-v1.3-os-part4-formula.html",
}


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    for name, url in SPECS.items():
        path = OUT / name
        if path.exists():
            continue
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        try:
            with urllib.request.urlopen(req, timeout=180) as r:
                path.write_bytes(r.read())
            print(f"{name}: {path.stat().st_size:,} バイト")
        except Exception as e:  # noqa: BLE001
            print(f"{name}: 取得できませんでした（{e}）", file=sys.stderr)
        time.sleep(2)
    return 0


if __name__ == "__main__":
    sys.exit(main())
