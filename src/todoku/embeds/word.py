"""Word の文書（.docx・.doc）の中のグラフ・SmartArt・埋め込みを集めて、窓の形にする（読むだけ）。

- グラフ（.docx の word/charts/chartN.xml。本文の〔グラフ1〕の所）と SmartArt（word/diagrams/dataN.xml。〔SmartArt1〕の所）
- 埋め込み（〔埋め込み1：…〕の所）：Excel の表として窓に出すもの（embedded_books が開けるもの。docx.window._embeds が出す）
  のほかを、ole.describe で読む。.doc の ObjectPool の入れ物、.docx の word/embeddings の部品、RC4 CryptoAPI で暗号化した .doc の
  Data ストリームの中の入れ物（FOBJH の後ろ、MS-DOC 2.9.94）
"""

from __future__ import annotations

import struct
import zlib

from ..cfb import CFB, SubStorage
from .chart import chart_dict, chart_markdown, read_chart
from .ole import describe, object_markdown

MAGIC_CFB = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"


def fobjh_storage(data: bytes, pos: int) -> bytes | None:
    """Data ストリームの pos の FOBJH（cbHeader 2・フラグ 2（bit0 fCompressed）・cbObj 4）と、その後ろの OLE の入れ物
    （複合ファイル。fCompressed なら RFC 1950 の zlib で縮めたもの）。形が合わなければ None。"""
    if pos < 0 or pos + 8 > len(data):
        return None
    cb, flags, cb_obj = struct.unpack_from("<hHi", data, pos)
    if cb != 8 or cb_obj <= 8 or pos + cb_obj > len(data):
        return None
    body = data[pos + 8:pos + cb_obj]
    if flags & 1:
        try:
            body = zlib.decompress(body)
        except zlib.error:
            return None
    return body if body[:8] == MAGIC_CFB else None


def _source(doc, e: dict):
    st = e["storage"]
    if getattr(doc, "kind", "") == "docx":
        return doc.zip.read(st)
    if st.startswith("Data@"):
        return fobjh_storage(doc.data, int(st[5:]))
    return SubStorage(doc.cfb, st)


def objects(doc, depth: int = 0, shown: set[int] | None = None) -> list[dict]:
    """文書のグラフ・SmartArt・（Excel の表として出すもののほかの）埋め込み。文書の中の順（種類ごと）。
    shown：窓の「埋め込み」の節に Excel の表として出した埋め込みの番号（なければ embedded_books で調べる）。"""
    out: list[dict] = []
    pkg = getattr(doc, "zip", None)
    for c in getattr(doc, "charts", []):
        try:
            from .chart import workbook_resolver  # noqa: F401

            ch = read_chart(pkg.read(c["part"]), c["part"], pkg)
        except Exception as ex:  # noqa: BLE001
            from .chart import Chart

            ch = Chart(part=c["part"], error=str(ex))
        ch.alt = c.get("alt", "")
        out.append({"kind": "chart", **chart_dict(ch, c["n"])})
    for g in getattr(doc, "diagrams", []):
        from .smartart import smartart_lines

        try:
            lines = smartart_lines(pkg.read(g["part"]))
            err = ""
        except Exception as ex:  # noqa: BLE001
            lines, err = [], str(ex)
        out.append({"kind": "smartart", "n": g["n"], "part": g["part"], "lines": lines, **({"error": err} if err else {})})
    books = set(shown or ())
    if shown is None and depth == 0 and hasattr(doc, "embedded_books"):
        try:
            books = {e["n"] for e, _wb in doc.embedded_books()}
        except Exception:  # noqa: BLE001
            books = set()
    for e in getattr(doc, "embeds", []):
        if e["n"] in books:
            continue  # Excel の表として、窓の「埋め込み」の節に出る
        try:
            src = _source(doc, e)
        except Exception as ex:  # noqa: BLE001
            src, err = None, str(ex)
        if src is None:
            d = {"kind": "other", "label": e.get("label", "埋め込み"), "class": e.get("prog", ""), "markdown": [],
                 "error": "入れ物を見つけられませんでした"}
        else:
            d = describe(src, e["storage"], e.get("prog", ""), e.get("alt", ""), depth)
        d.update({"n": e["n"], "storage": e["storage"], "prog": e.get("prog", "")})
        out.append({**d, "kind_of": d["kind"], "kind": "embed"})
    return out


def item_markdown(o: dict, level: str = "##") -> list[str]:
    if o["kind"] == "chart":
        return chart_markdown(o, level)
    if o["kind"] == "smartart":
        out = [f"{level} SmartArt{o['n']}（{o['part']}）", "", "- 本文の〔SmartArt" + str(o["n"]) + "〕の所にある図表の字です（入れ子は字下げ）。"]
        if o.get("error"):
            out.append(f"- 読めませんでした：{o['error']}")
        out.append("")
        out += [f"- {x}" for x in o["lines"]]
        return out + [""]
    return object_markdown(o, level)


def markdown(items: list[dict] | None, level: str = "##") -> list[str]:
    out: list[str] = []
    for o in items or []:
        out += item_markdown(o, level)
    return out


del CFB
