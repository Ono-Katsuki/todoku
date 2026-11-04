"""OpenOffice.org 1.x の文書（.sxw）・表計算（.sxc）を読む（読むだけ）。

入れ物と XML の作りは OpenDocument の元になった形。名前空間の URI と本文の包み方が違うだけなので、content.xml・styles.xml を
OpenDocument の形に写して（namespace の置き換え・office:text で包む・値の属性の接頭辞）、OpenDocument の読み方で読む。
書き込みはしない（写したものへの書き込みを、元のファイルに戻して証明する方法がない）。
"""

from __future__ import annotations

import io
import re
import zipfile

from ..zipraw import RawZip

MIMES = {"application/vnd.sun.xml.writer": "sxw", "application/vnd.sun.xml.writer.template": "sxw",
         "application/vnd.sun.xml.writer.global": "sxw",
         "application/vnd.sun.xml.calc": "sxc", "application/vnd.sun.xml.calc.template": "sxc"}
KIND_JA = {"sxw": "OpenOffice.org 1.x の文書（.sxw。読むだけ）", "sxc": "OpenOffice.org 1.x の表計算（.sxc。読むだけ）"}
WRITE = "OpenOffice.org 1.x の形には書きません。LibreOffice で .odt・.ods に保存し直してもらえば書けます"
_NS = {"http://openoffice.org/2000/office": "urn:oasis:names:tc:opendocument:xmlns:office:1.0",
       "http://openoffice.org/2000/text": "urn:oasis:names:tc:opendocument:xmlns:text:1.0",
       "http://openoffice.org/2000/style": "urn:oasis:names:tc:opendocument:xmlns:style:1.0",
       "http://openoffice.org/2000/table": "urn:oasis:names:tc:opendocument:xmlns:table:1.0",
       "http://openoffice.org/2000/drawing": "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0",
       "http://openoffice.org/2000/datastyle": "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0",
       "http://openoffice.org/2000/form": "urn:oasis:names:tc:opendocument:xmlns:form:1.0",
       "http://openoffice.org/2000/dr3d": "urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0",
       "http://openoffice.org/2000/chart": "urn:oasis:names:tc:opendocument:xmlns:chart:1.0",
       "http://openoffice.org/2000/script": "urn:oasis:names:tc:opendocument:xmlns:script:1.0",
       "http://openoffice.org/2000/meta": "urn:oasis:names:tc:opendocument:xmlns:meta:1.0",
       "http://www.w3.org/1999/XSL/Format": "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0",
       "http://www.w3.org/2000/svg": "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"}


class SxoError(ValueError):
    pass


def kind_of_package(z: RawZip, names: set) -> str | None:
    if "mimetype" not in names:
        return None
    try:
        return MIMES.get(z.read("mimetype").decode("ascii", "replace").strip())
    except Exception:  # noqa: BLE001
        return None


def _translate(xml: bytes, wrap: str | None) -> bytes:
    s = xml.decode("utf-8", "replace")
    s = re.sub(r"<!DOCTYPE[^>]*>", "", s, count=1)
    for a, b in _NS.items():
        s = s.replace(f'"{a}"', f'"{b}"')
    for attr in ("value-type", "value", "date-value", "time-value", "boolean-value", "string-value", "currency"):
        s = re.sub(rf"\btable:{attr}=", f"office:{attr}=", s)
    if wrap:
        s = re.sub(r"<office:body\b([^>]*)>", rf"<office:body\1><office:{wrap}>", s, count=1)
        s = s.replace("</office:body>", f"</office:{wrap}></office:body>", 1)
    return s.encode("utf-8")


def as_odf(blob: bytes) -> tuple[bytes, str]:
    """(OpenDocument の形に写したパッケージ, sxw か sxc)。"""
    try:
        z = RawZip(blob)
    except Exception as e:  # noqa: BLE001
        raise SxoError(f"ZIP として読めません（{e}）") from None
    names = set(z.names())
    k = kind_of_package(z, names)
    if k is None or "content.xml" not in names:
        raise SxoError("OpenOffice.org 1.x の文書として読めません")
    buf = io.BytesIO()
    mime = "application/vnd.oasis.opendocument.text" if k == "sxw" else "application/vnd.oasis.opendocument.spreadsheet"
    with zipfile.ZipFile(buf, "w") as out:
        out.writestr("mimetype", mime)
        out.writestr("content.xml", _translate(z.read("content.xml"), "text" if k == "sxw" else "spreadsheet"))
        if "styles.xml" in names:
            out.writestr("styles.xml", _translate(z.read("styles.xml"), None))
    return buf.getvalue(), k


def raw_text(blob: bytes) -> str:
    """窓とは別に：元の content.xml の text:p・text:h の中の字を、タグを除いて拾う（検査用）。"""
    import html

    x = RawZip(blob).read("content.xml").decode("utf-8", "replace")
    out = []
    for m in re.finditer(r"(?s)<text:(p|h)\b[^>]*?(?<!/)>(.*?)</text:\1>", x):
        out.append(html.unescape(re.sub(r"<[^>]*>", "", m[2])))
    return "".join(out)
