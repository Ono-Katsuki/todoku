"""ファイル形式の公式の仕様書（Microsoft の公開仕様・ISO 32000-1）を取得し、とどく自身で文字に直す。

仕様書は著作物なので、リポジトリには入れない（corpus/specs/、公開しない）。tests/test_spec.py が、
とどくの中の表（文書の目次の並び・位置の表の要素の大きさ・書式の命令の番号・保護のビット）を原本と照らす。
使い方: uv run python scripts/fetch_specs.py
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
    "MS-DOC": "https://officeprotocoldoc.z19.web.core.windows.net/files/MS-DOC/%5bMS-DOC%5d.pdf",
    "MS-XLS": "https://officeprotocoldoc.z19.web.core.windows.net/files/MS-XLS/%5bMS-XLS%5d.pdf",
    "MS-OFFCRYPTO": "https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OFFCRYPTO/%5bMS-OFFCRYPTO%5d.pdf",
    "MS-OVBA": "https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OVBA/%5bMS-OVBA%5d.pdf",
    # Excel のフォームの部品（ctrlProps の formControlPr）と、VML の x:ClientData（Office の拡張）
    "MS-XLSX": "https://officeprotocoldoc.z19.web.core.windows.net/files/MS-XLSX/%5bMS-XLSX%5d.pdf",
    "MS-OI29500": "https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OI29500/%5bMS-OI29500%5d.pdf",
    # 文書に埋め込んだ表（OLE の部品）と、Word が表示・印刷に使う表の絵（.doc の Data の中の OfficeArt、EMF・WMF）。
    # Windows の仕様は置き場所が別
    "MS-ODRAW": "https://officeprotocoldoc.z19.web.core.windows.net/files/MS-ODRAW/%5bMS-ODRAW%5d.pdf",
    "MS-OLEDS": "https://winprotocoldoc.z19.web.core.windows.net/MS-OLEDS/%5bMS-OLEDS%5d.pdf",
    "MS-EMF": "https://winprotocoldoc.z19.web.core.windows.net/MS-EMF/%5bMS-EMF%5d.pdf",
    "MS-WMF": "https://winprotocoldoc.z19.web.core.windows.net/MS-WMF/%5bMS-WMF%5d.pdf",
    "MS-CFB": "https://winprotocoldoc.z19.web.core.windows.net/MS-CFB/%5bMS-CFB%5d.pdf",
    # ActiveX の部品（Microsoft Forms 2.0：テキストボックス・チェックボックス・コンボボックスなど）の保存の形
    "MS-OFORMS": "https://officeprotocoldoc.z19.web.core.windows.net/files/MS-OFORMS/%5bMS-OFORMS%5d.pdf",
    "PDF32000_2008": "https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/PDF32000_2008.pdf",
    # XFA 3.3（PDF Association の norm-refs の写し。pdfa.org は直接の取得を断るので、Internet Archive の写しから）
    "XFA-3_3": "https://web.archive.org/web/20220817145529id_/https://www.pdfa.org/norm-refs/XFA-3_3.pdf",
}


def to_text(pdf: Path, txt: Path) -> None:
    sys.path.insert(0, str(ROOT / "src"))
    from todoku.pdf.content import page_content
    from todoku.pdf.doc import PdfDocument
    from todoku.pdf.layout import PageLayout

    d = PdfDocument(pdf.read_bytes(), pdf.name)
    out = []
    for i, p in enumerate(d.pages()):
        lay = PageLayout(d, p, page_content(d, p), i + 1)
        lines = sorted(lay.free_lines + [ln for c in lay.cells for ln in c.lines], key=lambda ln: (round(ln.y0), ln.x0))
        out.append(f"=== page {i + 1}\n" + "\n".join(ln.text for ln in lines))
    txt.write_text("\n".join(out), encoding="utf-8")


def main() -> int:
    OUT.mkdir(parents=True, exist_ok=True)
    for name, url in SPECS.items():
        pdf = OUT / f"{name}.pdf"
        if not pdf.exists() or pdf.read_bytes()[:5] != b"%PDF-":  # 前に取得に失敗したもの（404 のページ）は取り直す
            req = urllib.request.Request(url, headers={"User-Agent": UA})
            with urllib.request.urlopen(req, timeout=120) as r:
                pdf.write_bytes(r.read())
            time.sleep(2)
        txt = pdf.with_suffix(".txt")
        if name.startswith("MS-") and not txt.exists():
            t = time.time()
            to_text(pdf, txt)
            print(f"{name}: {pdf.stat().st_size:,} バイト → 文字 {txt.stat().st_size:,} バイト（{time.time() - t:.0f} 秒）")
    return 0


if __name__ == "__main__":
    sys.exit(main())
