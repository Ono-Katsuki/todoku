""".xls のピボットテーブルの定義（公開仕様 [MS-XLS] v20250819 から実装）。計算は再現しない。読むのは定義だけ。

- ブック：SXStreamID（2.4.303）・SXVS（2.4.317）と、その後の DConRef（2.4.86）・DConName（2.4.85）・DConBin（2.4.83）
  （PIVOTCACHEDEFINITION = SXStreamID SXVS [SXSRC] [SXADDLCACHE]）
- ピボットのキャッシュ：_SX_DB_CUR の、idStm を 4 桁の 16 進にした名前のストリーム（2.1.7.12）の SXFDB（2.4.283）が項目の名前
- シート：SxView（2.4.313）に続く Sxvd（2.4.309）・SxIvd（2.4.292。行、次に列）・SXPI（2.4.298）・SXDI（2.4.278）
  （PIVOTCORE = SxView *PIVOTVD *2PIVOTIVD [PIVOTPI] *SXDI *PIVOTLI PIVOTEX）
- ピボットの項目 i はキャッシュの項目 i に対応する（2.2.5.4.3）。SxIvd の -2 は値の項目（2.5.258）
"""

from __future__ import annotations

import struct

from .book import records

RT_BOF, RT_EOF, RT_CONTINUE, RT_BOUNDSHEET = 0x0809, 0x000A, 0x003C, 0x0085
RT_SXVIEW, RT_SXVD, RT_SXIVD, RT_SXPI, RT_SXDI = 176, 177, 180, 182, 197
RT_SXFDB, RT_SXSTREAMID, RT_SXVS = 199, 213, 227
RT_DCONREF, RT_DCONNAME, RT_DCONBIN = 81, 82, 437

_SOURCE = {1: "worksheet", 2: "external", 4: "consolidation", 16: "scenario"}
# SXDI の iiftab（2.4.278）。名前は .xlsx の dataField の subtotal と同じ
_FUNC = ["sum", "count", "average", "max", "min", "product", "countNums", "stdDev", "stdDevp", "var", "varp"]
_DATA_FIELD = -2  # SxIvd の値：行・列にある「値」の項目


def _no_cch(b: bytes, at: int, cch: int) -> tuple[str, int]:
    """XLUnicodeStringNoCch：頭の 1 バイトの下位 1 ビットが 16 ビットの字か。(字, 次の位置)。"""
    if at >= len(b):
        return "", at
    wide = b[at] & 1
    at += 1
    n = cch * (2 if wide else 1)
    raw = b[at:at + n]
    return (raw.decode("utf-16-le", "replace") if wide else raw.decode("latin-1")), at + n


def _ustr(b: bytes, at: int) -> str:
    """XLUnicodeString：字の数（2 バイト）と XLUnicodeStringNoCch。"""
    if at + 2 > len(b):
        return ""
    return _no_cch(b, at + 2, struct.unpack_from("<H", b, at)[0])[0]


def _dcon_file(st: str) -> dict:
    """DConFile（2.5.69）：self-reference = 0x0002 シート名、ほかはほかのファイル（名前のまま）。"""
    if st.startswith("\x02"):
        return {"sheet": st[1:]}
    return {"file": st.lstrip("\x01")} if st else {}


def _cache_sources(recs: list) -> list[dict]:
    """ブックの SXStreamID の順に：{id, source_type, source}。SxView の iCache はこの番号。"""
    out: list[dict] = []
    for _pos, t, b in recs:
        if t == RT_SXSTREAMID and len(b) >= 2:
            out.append({"id": struct.unpack_from("<H", b, 0)[0]})
        elif not out:
            continue
        elif t == RT_SXVS and len(b) >= 2 and "source_type" not in out[-1]:
            v = struct.unpack_from("<H", b, 0)[0]
            out[-1]["source_type"] = _SOURCE.get(v, str(v))
        elif t == RT_DCONREF and len(b) >= 8 and "source" not in out[-1]:
            r1, r2, c1, c2, cch = struct.unpack_from("<HHBBH", b, 0)
            st = _no_cch(b, 8, cch)[0] if cch else ""
            out[-1]["source"] = {"ref": (r1, c1, r2, c2), **_dcon_file(st)}
        elif t == RT_DCONNAME and len(b) >= 2 and "source" not in out[-1]:
            cch = struct.unpack_from("<H", b, 0)[0]
            name, at = _no_cch(b, 2, cch)
            cf = struct.unpack_from("<H", b, at)[0] if at + 2 <= len(b) else 0
            out[-1]["source"] = {"name": name, **(_dcon_file(_no_cch(b, at + 2, cf)[0]) if cf else {})}
        elif t == RT_DCONBIN and len(b) >= 6 and "source" not in out[-1]:
            out[-1]["source"] = {"builtin_name": b[0]}
    return out


def _cache_fields(cfb, stream_id: int) -> list[str]:
    name = f"_SX_DB_CUR/{stream_id:04X}"
    if not cfb.has(name):
        return []
    # SXFDB の頭の 14 バイトの後が stFieldName（XLUnicodeString）
    return [_ustr(b, 14) for _p, t, b in records(cfb.read(name)) if t == RT_SXFDB]


def _sheet_names(recs: list) -> list[str]:
    """BoundSheet8（2.4.28）：ストリームの位置の順に、シートの名前。"""
    got = []
    for _pos, t, b in recs:
        if t == RT_BOUNDSHEET and len(b) >= 8:
            got.append((struct.unpack_from("<I", b, 0)[0], _no_cch(b, 7, b[6])[0]))
    return [n for _o, n in sorted(got)]


def _with_continue(recs: list, i: int) -> bytes:
    """i 番目の記録の中身と、続く Continue の中身（PIVOTIVD = SxIvd *Continue。PIVOTPI も同じ）。"""
    data = recs[i][2]
    j = i + 1
    while j < len(recs) and recs[j][1] == RT_CONTINUE:
        data += recs[j][2]
        j += 1
    return data


def _view(recs: list, i: int, sheet: str | None, caches: list[dict], cfb, memo: dict) -> dict:
    b = recs[i][2]
    r1, r2, c1, c2 = struct.unpack_from("<HHHH", b, 0)
    i_cache = struct.unpack_from("<h", b, 14)[0]
    c_dim, c_rw, c_col, c_pg, c_data = struct.unpack_from("<hHHHh", b, 20)
    cch_table = struct.unpack_from("<H", b, 40)[0]
    name = _no_cch(b, 44, cch_table)[0] if cch_table not in (0, 0xFFFF) else ""
    cache = caches[i_cache] if 0 <= i_cache < len(caches) else {}
    if cache and cache["id"] not in memo:
        memo[cache["id"]] = _cache_fields(cfb, cache["id"])
    fields = memo.get(cache.get("id"), [])
    shown: list[str] = []  # ピボットの項目ごとの名前：項目の表示名か、キャッシュの項目の名前

    def field(k: int) -> str:
        if 0 <= k < len(shown) and shown[k]:
            return shown[k]
        return fields[k] if 0 <= k < len(fields) else f"#{k}"

    ivd: list[list[int]] = []
    page: list[int] = []
    data_items: list[dict] = []
    j = i + 1
    while j < len(recs):
        t, d = recs[j][1], recs[j][2]
        if t in (RT_SXVIEW, RT_EOF, RT_BOF):
            break
        if t == RT_SXVD and len(d) >= 10:
            cch = struct.unpack_from("<H", d, 8)[0]
            shown.append(_no_cch(d, 10, cch)[0] if cch != 0xFFFF else "")
        elif t == RT_SXIVD:
            whole = _with_continue(recs, j)
            ivd.append(list(struct.unpack_from(f"<{len(whole) // 2}h", whole, 0)))
        elif t == RT_SXPI:
            whole = _with_continue(recs, j)
            page = [struct.unpack_from("<h", whole, k)[0] for k in range(0, len(whole) - 5, 6)]  # SXPI_Item の頭が isxvd
        elif t == RT_SXDI and len(d) >= 14:
            fld, func = struct.unpack_from("<hh", d, 0)
            cch = struct.unpack_from("<H", d, 12)[0]
            data_items.append({"name": _no_cch(d, 14, cch)[0] if cch != 0xFFFF else "", "field": fld, "function": func})
        j += 1
    rows = ivd[0] if c_rw and ivd else []
    cols = (ivd[1] if c_rw and len(ivd) > 1 else ivd[0] if not c_rw and ivd else []) if c_col else []

    def axis(items: list[int]) -> list[str]:
        return ["（値）" if k == _DATA_FIELD else field(k) for k in items]

    return {"name": name, "sheet": sheet, "location": (r1, c1, r2, c2), "source": cache.get("source"), "source_type": cache.get("source_type"),
            "fields": [field(k) for k in range(max(c_dim, len(shown)))] or fields, "row_fields": axis(rows), "column_fields": axis(cols),
            "page_fields": axis(page),
            "data_fields": [{"name": x["name"], "field": field(x["field"]),
                             "function": _FUNC[x["function"]] if 0 <= x["function"] < len(_FUNC) else str(x["function"])} for x in data_items],
            "counts": {"fields": c_dim, "rows": c_rw, "columns": c_col, "pages": c_pg, "data": c_data}}


def views(cfb, stream: bytes) -> list[dict]:
    """ブックのすべてのピボットの定義。"""
    recs = list(records(stream))
    names = _sheet_names(recs)
    caches = _cache_sources(recs)
    memo: dict = {}
    out = []
    depth = 0
    substream = -1  # 0 がブック、1 からがシート（シートの中のグラフは、入れ子の BOF）
    for i, (_pos, t, _b) in enumerate(recs):
        if t == RT_BOF:
            depth += 1
            if depth == 1:
                substream += 1
        elif t == RT_EOF:
            depth = max(depth - 1, 0)
        elif t == RT_SXVIEW and len(recs[i][2]) >= 44:
            sheet = names[substream - 1] if 0 < substream <= len(names) else None
            out.append(_view(recs, i, sheet, caches, cfb, memo))
    return out
