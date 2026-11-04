"""数式のセルの外にある処理：条件付き書式・入力規則（それぞれ規則の種類として）、ピボットテーブル・Power Query・データ接続
（セルの数式では書けない処理として、その定義を）。"""

from __future__ import annotations

import base64
import io
import re
import struct
import xml.etree.ElementTree as ET
import zipfile

from .rules import rect_text


def conditional_formats(wb) -> list[dict]:
    out = []
    for s in wb.sheets:
        for rule in getattr(s, "cond_rules", ()):
            d = {"kind": "conditional_format", "sheet": s.name,
                 "range": " ".join(rect_text(*x) for x in rule.sqref[:32]), "type": rule.type, "priority": rule.priority}
            if rule.operator:
                d["operator"] = rule.operator
            if rule.formulas:
                d["formulas"] = list(rule.formulas)
            if rule.stop:
                d["stop_if_true"] = True
            dxf = _dxf(wb, rule.dxf)
            if dxf:
                d["format"] = dxf
            out.append(d)
    return out


def _dxf(wb, i) -> dict | None:
    xs = getattr(wb, "dxfs", None)
    if i is None or not xs or not 0 <= i < len(xs):
        return None
    x = xs[i]
    d = {k: getattr(x, k) for k in ("font_color", "fill_color", "num_fmt") if getattr(x, k, None)}
    return d or None


def data_validations(wb) -> list[dict]:
    out = []
    for s in wb.sheets:
        for v in getattr(s, "validations", ()):
            d = {"kind": "data_validation", "sheet": s.name, "range": " ".join(rect_text(*x) for x in v.sqref[:32]) or v.sqref_text,
                 "type": v.type, "blocks": v.blocks, "allow_blank": v.allow_blank}
            for k in ("operator", "formula1", "formula2", "prompt", "error"):
                x = getattr(v, k, None)
                if x:
                    d[k] = x
            out.append(d)
    return out


# ---------------------------------------------------------------- ピボット・Power Query・接続
_M = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_NOT_FORMULA = "セルの数式では書けない処理。定義だけを示し、計算は再現しない"


def not_cell_processing(blob: bytes | None, wb, kind: str) -> list[dict]:
    """ピボットテーブル・Power Query・データ接続・クエリテーブル。"""
    if kind in ("ods", "fods"):
        return _ods_pilots(blob, kind)
    if kind != "xlsx" and kind != "xlsb":
        return _binary_book_notes(blob, kind)
    z = getattr(wb, "zip", None)
    if z is None:
        return []
    names = list(z.names())
    out: list[dict] = []
    if kind == "xlsb":
        return _xlsb_notes(names)
    caches = _pivot_caches(z, names)
    sheet_of = _sheet_of_part(wb, z)
    for n in sorted(names):
        if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.xml", n):
            out.append(_pivot(z, n, caches, sheet_of))
    for n in sorted(names):
        low = n.lower()
        if low.startswith("customxml/") and low.endswith(".xml") and "/_rels/" not in low:
            pq = _power_query(z.read(n), n)
            if pq:
                out.append(pq)
    if z.has("xl/connections.xml"):
        out += _connections(z.read("xl/connections.xml"))
    for n in sorted(names):
        if re.fullmatch(r"xl/queryTables/queryTable\d+\.xml", n):
            root = ET.fromstring(z.read(n))
            t = next((t for t in getattr(wb, "tables", {}).values() if t.table_type == "queryTable"), None)
            d = {"kind": "query_table", "part": n, "name": root.get("name"), "connection_id": root.get("connectionId"),
                 "expressible_by_cells": False, "note": _NOT_FORMULA}
            if t is not None:
                d["output"] = {"sheet": t.sheet, "range": rect_text(t.r1, t.c1, t.r2, t.c2)}
            out.append(d)
    return out


def _sheet_of_part(wb, z) -> dict[str, str]:
    """ピボットの部品 → 置かれたシート。"""
    out = {}
    for s in wb.sheets:
        path = getattr(s, "path", "")
        if not path:
            continue
        try:
            rels = wb.rels(path)
        except Exception:  # noqa: BLE001
            continue
        for typ, target in rels.values():
            if typ == "pivotTable":
                out[target] = s.name
    return out


def _pivot_caches(z, names) -> dict[str, dict]:
    out = {}
    for n in names:
        if not re.fullmatch(r"xl/pivotCache/pivotCacheDefinition\d+\.xml", n):
            continue
        root = ET.fromstring(z.read(n))
        d: dict = {"fields": [f.get("name") for f in root.iter(_M + "cacheField")]}
        src = root.find(_M + "cacheSource")
        if src is not None:
            d["source_type"] = src.get("type")
            ws = src.find(_M + "worksheetSource")
            if ws is not None:
                d["source"] = {k: ws.get(k) for k in ("sheet", "ref", "name") if ws.get(k)}
            if src.get("connectionId"):
                d["connection_id"] = src.get("connectionId")
        if root.get("refreshedDate"):
            d["refreshed_date"] = root.get("refreshedDate")
        out[n] = d
    return out


def _pivot(z, part: str, caches: dict, sheet_of: dict) -> dict:
    import posixpath

    root = ET.fromstring(z.read(part))
    cache = {}
    d, base = posixpath.split(part)
    rp = posixpath.join(d, "_rels", base + ".rels")
    if z.has(rp):
        for rel in ET.fromstring(z.read(rp)):
            if rel.get("Type", "").endswith("/pivotCacheDefinition"):
                cache = caches.get(posixpath.normpath(posixpath.join(d, rel.get("Target", ""))), {})
    fields = cache.get("fields", [])

    def field_name(i) -> str:
        i = int(i)
        return fields[i] if 0 <= i < len(fields) else f"#{i}"

    def axis(tag: str, child: str, attr: str = "x") -> list[str]:
        el = root.find(_M + tag)
        out = []
        for x in el if el is not None else []:
            if x.tag == _M + child and x.get(attr) is not None and int(x.get(attr)) >= 0:
                out.append(field_name(x.get(attr)))
        return out

    loc = root.find(_M + "location")
    rec = {"kind": "pivot_table", "name": root.get("name"), "sheet": sheet_of.get(part),
           "location": loc.get("ref") if loc is not None else None,
           "source": cache.get("source"), "source_type": cache.get("source_type"),
           "fields": fields, "row_fields": axis("rowFields", "field"), "column_fields": axis("colFields", "field"),
           "page_fields": axis("pageFields", "pageField", "fld"),
           "data_fields": [{"name": x.get("name"), "field": field_name(x.get("fld", 0)), "function": x.get("subtotal", "sum")}
                           for x in (root.find(_M + "dataFields") if root.find(_M + "dataFields") is not None else [])],
           "expressible_by_cells": False, "note": _NOT_FORMULA}
    return {k: v for k, v in rec.items() if v not in (None, [])}


def _power_query(data: bytes, part: str) -> dict | None:
    """DataMashup（MS-QDEFF）の中の M のソース。"""
    if b"D\x00a\x00t\x00a\x00M\x00a\x00s\x00h\x00u\x00p" not in data and b"DataMashup" not in data:
        return None
    try:
        text = data.decode("utf-16") if data[:2] in (b"\xff\xfe", b"\xfe\xff") else data.decode("utf-8", "replace")
        root = ET.fromstring(text)
    except Exception:  # noqa: BLE001
        return {"kind": "power_query", "part": part, "expressible_by_cells": False, "note": "定義（DataMashup）を読めませんでした"}
    if not root.tag.endswith("DataMashup"):
        return None
    d: dict = {"kind": "power_query", "part": part, "expressible_by_cells": False, "note": _NOT_FORMULA}
    try:
        raw = base64.b64decode(root.text or "")
        _ver, size = struct.unpack_from("<II", raw, 0)
        with zipfile.ZipFile(io.BytesIO(raw[8:8 + size])) as zf:
            d["m_source"] = {n: zf.read(n).decode("utf-8", "replace") for n in zf.namelist() if n.lower().endswith(".m")}
    except Exception as e:  # noqa: BLE001
        d["note"] += f"。M のソースは読めませんでした（{type(e).__name__}）"
    return d


def _connections(data: bytes) -> list[dict]:
    out = []
    root = ET.fromstring(data)
    for c in root.findall(_M + "connection"):
        d = {"kind": "data_connection", "id": c.get("id"), "name": c.get("name"), "type": c.get("type"),
             "expressible_by_cells": False, "note": _NOT_FORMULA}
        db = c.find(_M + "dbPr")
        if db is not None:
            d["command"] = db.get("command")
            # 接続の字（パスワードが入っていることがある）は写さない
            d["connection_text_withheld"] = bool(db.get("connection"))
        web = c.find(_M + "webPr")
        if web is not None:
            d["url"] = web.get("url")
        out.append({k: v for k, v in d.items() if v not in (None, "")})
    return out


def _xlsb_notes(names: list[str]) -> list[dict]:
    out = []
    for n in sorted(names):
        if re.fullmatch(r"xl/pivotTables/pivotTable\d+\.bin", n):
            out.append({"kind": "pivot_table", "part": n, "expressible_by_cells": False,
                        "note": "ピボットテーブルがあります。.xlsb の中のピボットの定義は読んでいません（置いてある値だけを読みます）"})
    return out


def _binary_book_notes(blob: bytes | None, kind: str) -> list[dict]:
    if kind != "xls" or not blob:
        return []
    try:
        return _xls_pivots(blob)
    except Exception as e:  # noqa: BLE001  定義を読めないことは、処理の取り出しを止める理由にならない。ただし黙らない
        return [{"kind": "pivot_table", "expressible_by_cells": False, "note": f"ピボットの定義を読めませんでした（{type(e).__name__}）"}]


def _xls_pivots(blob: bytes) -> list[dict]:
    """.xls のピボットの定義（xls.pivot。[MS-XLS] から）。"""
    from ..cfb import CFB
    from ..xls.pivot import views

    c = CFB(blob)
    names = {n.lower(): n for n in c.streams()}
    key = names.get("workbook") or names.get("book")
    if key is None:
        return []
    out = []
    for v in views(c, c.read(key)):
        r1, c1, r2, c2 = v.pop("location")
        v.pop("counts", None)
        src = v.get("source")
        if src and "ref" in src:
            a_, b_, c_, d_ = src["ref"]
            src["ref"] = rect_text(a_ + 1, b_ + 1, c_ + 1, d_ + 1)
        rec = {"kind": "pivot_table", **v, "location": rect_text(r1 + 1, c1 + 1, r2 + 1, c2 + 1),
               "expressible_by_cells": False, "note": _NOT_FORMULA}
        out.append({k: x for k, x in rec.items() if x not in (None, [], "")})
    return out


def _ods_pilots(blob: bytes | None, kind: str) -> list[dict]:
    """.ods のデータパイロット（table:data-pilot-table）。"""
    out = []
    if not blob:
        return out
    data = blob
    if kind == "ods":
        with zipfile.ZipFile(io.BytesIO(blob)) as zf:
            data = zf.read("content.xml")
    t = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
    root = ET.fromstring(data)
    for p in root.iter(f"{{{t}}}data-pilot-table"):
        d = {"kind": "pivot_table", "name": p.get(f"{{{t}}}name"), "location": p.get(f"{{{t}}}target-range-address"),
             "expressible_by_cells": False, "note": _NOT_FORMULA}
        src = p.find(f"{{{t}}}source-cell-range")
        if src is not None:
            d["source"] = {"ref": src.get(f"{{{t}}}cell-range-address")}
        d["fields"] = [{"name": f.get(f"{{{t}}}source-field-name"), "orientation": f.get(f"{{{t}}}orientation"),
                        "function": f.get(f"{{{t}}}function")} for f in p.findall(f"{{{t}}}data-pilot-field")]
        out.append(d)
    return out


def external_targets(wb) -> dict[str, str]:
    """[n] → 外部ブックの場所（書かれたまま）。"""
    out: dict[str, str] = {}
    z = getattr(wb, "zip", None)
    if z is None or not hasattr(wb, "rels"):
        return out
    try:
        from ..xlsx.workbook import M, R

        wb_path = next((t for typ, t in wb.rels("").values() if typ == "officeDocument"), "xl/workbook.xml")
        root = wb.xml(wb_path)
        er = root.find(M + "externalReferences")
        rels = wb.rels(wb_path)
        for i, e in enumerate(er if er is not None else [], 1):
            part = rels.get(e.get(R + "id"), ("", ""))[1]
            for typ, target in wb.rels(part).values():
                if typ in ("externalLinkPath", "xlPathMissing", "xlExternalLinkPath"):
                    out[str(i)] = target
    except Exception:  # noqa: BLE001
        pass
    return out
