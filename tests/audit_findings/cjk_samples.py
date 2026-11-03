"""縦書き・ルビ・異体字セレクタ・外字・全角半角の見本（平らな ODT）。LibreOffice で PDF・docx にして読む。

使い方: python cjk_samples.py OUTDIR  → OUTDIR/vertical.fodt（縦書きのページ）と OUTDIR/cjk.fodt（横書き）
"""

from __future__ import annotations

import sys
from pathlib import Path

HEAD = """<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
<office:font-face-decls/>
<office:styles>
 <style:default-style style:family="paragraph"><style:text-properties style:font-name-asian="Hiragino Mincho ProN"
   style:font-family-asian="'Hiragino Mincho ProN'" fo:font-size="14pt" style:font-size-asian="14pt"/></style:default-style>
</office:styles>
<office:automatic-styles>
 <style:page-layout style:name="PL"><style:page-layout-properties fo:page-width="210mm" fo:page-height="297mm"
   style:writing-mode="{mode}" fo:margin="20mm"/></style:page-layout>
</office:automatic-styles>
<office:master-styles><style:master-page style:name="Standard" style:page-layout-name="PL"/></office:master-styles>
<office:body><office:text>
"""
TAIL = "</office:text></office:body></office:document>\n"

VERTICAL = [
    "一、申請者の氏名を記入してください。",
    "二、住所は令和七年十一月三日現在のものとします。",
    "三、電話番号（０３－１２３４－５６７８）",
]

CJK = [
    # ルビ：「葛飾」に「かつしか」
    '<text:p>ruby | <text:ruby><text:ruby-base>葛飾</text:ruby-base><text:ruby-text>かつしか</text:ruby-text></text:ruby>区役所</text:p>',
    "<text:p>ivs | 葛\U000E0100飾区 辻\U000E0101 邊\U000E0103</text:p>",
    "<text:p>pua | 髙橋 山太郎</text:p>",
    "<text:p>width | ＡＢＣ１２３ ABC123 ｱｲｳ アイウ ㈱ ①②</text:p>",
]


def main(out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    body = "\n".join(f"<text:p>{t}</text:p>" for t in VERTICAL)
    (out / "vertical.fodt").write_text(HEAD.replace("{mode}", "tb-rl") + body + TAIL, encoding="utf-8")
    (out / "cjk.fodt").write_text(HEAD.replace("{mode}", "lr-tb") + "\n".join(CJK) + TAIL, encoding="utf-8")
    print(out / "vertical.fodt", out / "cjk.fodt")


if __name__ == "__main__":
    main(Path(sys.argv[1]))
