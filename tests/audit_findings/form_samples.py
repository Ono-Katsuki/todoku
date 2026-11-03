"""書き込みの試し用の様式：見出しと空のマスの表（平らな ODT → LibreOffice で PDF・docx・odt に）と、
同じ形の xlsx（openpyxl）。見出しは英語（どの字の値を書くかは write_check.py が決める）。

使い方: python form_samples.py OUTDIR
"""

from __future__ import annotations

import sys
from pathlib import Path

LABELS = ["Arabic name", "Hebrew name", "Hindi name", "Bengali name", "Tamil name", "Thai name", "Khmer name",
          "Myanmar name", "Russian name", "Greek name", "Vietnamese name", "Chinese name", "Korean name", "Japanese name",
          "Georgian name", "Amharic name"]

FODT = """<?xml version="1.0" encoding="UTF-8"?>
<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 office:version="1.3" office:mimetype="application/vnd.oasis.opendocument.text">
<office:automatic-styles>
 <style:style style:name="T" style:family="table"><style:table-properties style:width="17cm" table:align="left"/></style:style>
 <style:style style:name="C" style:family="table-cell"><style:table-cell-properties fo:border="0.5pt solid #000000" fo:padding="0.1cm"/></style:style>
</office:automatic-styles>
<office:body><office:text>
<text:p>Application form</text:p>
<table:table table:name="F" table:style-name="T">
<table:table-column table:number-columns-repeated="2"/>
{rows}
</table:table>
</office:text></office:body></office:document>
"""


def main(out: Path) -> None:
    out.mkdir(parents=True, exist_ok=True)
    rows = "\n".join(f'<table:table-row><table:table-cell table:style-name="C"><text:p>{lab}</text:p></table:table-cell>'
                     f'<table:table-cell table:style-name="C"><text:p/></table:table-cell></table:table-row>' for lab in LABELS)
    (out / "form.fodt").write_text(FODT.replace("{rows}", rows), encoding="utf-8")
    import openpyxl
    from openpyxl.styles import Border, Side

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "form"
    thin = Side(style="thin")
    ws["A1"] = "Application form"
    for i, lab in enumerate(LABELS, 2):
        ws.cell(i, 1, lab)
        ws.cell(i, 2).border = Border(top=thin, bottom=thin, left=thin, right=thin)
    ws.column_dimensions["B"].width = 40
    wb.save(out / "form.xlsx")
    print(out / "form.fodt", out / "form.xlsx")


if __name__ == "__main__":
    main(Path(sys.argv[1]))
