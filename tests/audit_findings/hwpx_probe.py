"""最小の HWPX（OWPML の zip：mimetype が application/hwp+zip、本文は Contents/section0.xml）を作り、
todoku がどう扱うか（形式の判定・窓・段階）を確かめる。

使い方: PYTHONPATH=src python hwpx_probe.py OUT.hwpx
"""

from __future__ import annotations

import sys
import zipfile

from todoku import api
from todoku.book import kind_of

SECTION = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
 <hp:p id="0" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>신청인 성명:</hp:t></hp:run></hp:p>
 <hp:p id="1" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>주소:</hp:t></hp:run></hp:p>
</hs:sec>
"""
HPF = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<opf:package xmlns:opf="http://www.idpf.org/2007/opf/"><opf:manifest>
<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/></opf:manifest>
<opf:spine><opf:itemref idref="section0"/></opf:spine></opf:package>
"""


def main(out):
    with zipfile.ZipFile(out, "w") as z:
        z.writestr(zipfile.ZipInfo("mimetype"), "application/hwp+zip")
        z.writestr("version.xml", '<?xml version="1.0"?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" major="5" minor="1"/>')
        z.writestr("Contents/content.hpf", HPF)
        z.writestr("Contents/section0.xml", SECTION)
        z.writestr("META-INF/manifest.xml", '<?xml version="1.0"?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>')
    blob = open(out, "rb").read()
    print("kind_of:", kind_of(blob, out))
    try:
        k, doc, parts = api.window(out)
        print("window kind", k, "\n".join(md for _t, md in parts)[:600])
    except Exception as e:  # noqa: BLE001
        print("window error:", type(e).__name__, e)
    try:
        g = api.grade(out)
        print("grade:", g.get("grade"), g.get("name"), [r.get("detail") for r in g.get("reasons", [])][:2])
    except Exception as e:  # noqa: BLE001
        print("grade error:", type(e).__name__, e)


if __name__ == "__main__":
    main(sys.argv[1])
