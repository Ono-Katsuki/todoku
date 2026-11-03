"""古い Excel（.xls、BIFF8）のグラフを読む（MS-XLS 2.1.7.20.1 Chart Sheet Substream、2.2.3 Charts）。

グラフは、BOF（種類 0x0020）から EOF までのレコードの並び（グラフの部分）。グラフのシートでは、そのシートの部分そのもの。
シートに置いたグラフでは、シートの部分の中の、グラフの OBJ（ot＝5）のすぐ後ろに入れ子で入っている。

- 系列（Series 0x1003、Begin … End）：4つの AI（BRAI 0x1051。id 0 名前・1 値・2 項目・3 バブルの大きさ）と、名前を直接
  書いたときの SeriesText（0x100D）。BRAI の rt が 2 なら、formula（ChartParsedFormula）がセルの範囲を指す
- 値の控え（SERIESDATA：SIIndex 0x1065 のあとの Number・Label・BoolErr・Blank。行＝点の番号、列＝系列の番号）。
  グラフと同じシートのセルを指すデータには控えがない（2.2.3.2）ので、そのときは指しているセルを読む
- 字（Text 0x1025、Begin … End）：中の AI と SeriesText が字、ObjectLink（0x1027）が何の字か（1 題名・2 数値軸・
  3 項目軸・4 データラベル・7 系列軸・0x0C 表示単位）
- グラフの種類：ChartFormat（0x1014）の中の Bar・Line・Pie・Area・Scatter・Radar・RadarArea・Surf・BopPop
  （系列は SerToCrt 0x1045 で、どのグループのものか。SerParent 0x104A がある系列は、近似曲線か誤差範囲）
"""

from __future__ import annotations

import struct

from .chart import Chart, Series, number_text, workbook_resolver

BOF, EOF, CONTINUE = 0x0809, 0x000A, 0x003C
SERIES, SERIESTEXT, BRAI, TEXT, OBJECTLINK, BEGIN, END = 0x1003, 0x100D, 0x1051, 0x1025, 0x1027, 0x1033, 0x1034
CHARTFORMAT, SERTOCRT, SERPARENT, SIINDEX, AXIS, CHART3D, LEGEND = 0x1014, 0x1045, 0x104A, 0x1065, 0x101D, 0x103A, 0x1015
NUMBER, LABEL, BOOLERR, BLANK, LABELSST = 0x0203, 0x0204, 0x0205, 0x0201, 0x00FD
BAR, LINE, PIE, AREA, SCATTER, RADAR, SURF, RADARAREA, BOPPOP = 0x1017, 0x1018, 0x1019, 0x101A, 0x101B, 0x103E, 0x103F, 0x1040, 0x1061
_ERRORS = {0x00: "#NULL!", 0x07: "#DIV/0!", 0x0F: "#VALUE!", 0x17: "#REF!", 0x1D: "#NAME?", 0x24: "#NUM!", 0x2A: "#N/A"}


def substream(data: bytes, start: int) -> list[tuple[int, int, bytes]]:
    """start の BOF から、対応する EOF までのレコード：(位置, 種類, 中身)。入れ子の BOF … EOF も含む。"""
    out = []
    depth = 0
    pos = start
    n = len(data)
    while pos + 4 <= n:
        t, ln = struct.unpack_from("<HH", data, pos)
        out.append((pos, t, data[pos + 4:pos + 4 + ln]))
        pos += 4 + ln
        if t == BOF:
            depth += 1
        elif t == EOF:
            depth -= 1
            if depth <= 0:
                break
    return out


def _short_string(b: bytes, pos: int) -> str:
    """ShortXLUnicodeString（文字数 u8・fHighByte・文字）。"""
    if pos + 2 > len(b):
        return ""
    cch, flags = b[pos], b[pos + 1]
    raw = b[pos + 2:pos + 2 + cch * (2 if flags & 1 else 1)]
    return raw.decode("utf-16-le", errors="replace") if flags & 1 else raw.decode("latin-1")


def _kind(t: int, b: bytes, three_d: bool) -> str:
    flags = struct.unpack_from("<H", b, len(b) - 2)[0] if len(b) >= 2 else 0
    if t == BAR:
        flags = struct.unpack_from("<H", b, 4)[0] if len(b) >= 6 else 0
        name = "横棒グラフ" if flags & 0x01 else "縦棒グラフ"
        name = ("100% 積み上げ" if flags & 0x04 else "積み上げ" if flags & 0x02 else "") + name
    elif t == LINE:
        name = ("100% 積み上げ" if flags & 0x02 else "積み上げ" if flags & 0x01 else "") + "折れ線グラフ"
    elif t == PIE:
        donut = struct.unpack_from("<H", b, 2)[0] if len(b) >= 4 else 0
        name = "ドーナツグラフ" if donut else "円グラフ"
    elif t == AREA:
        name = ("100% 積み上げ" if flags & 0x02 else "積み上げ" if flags & 0x01 else "") + "面グラフ"
    elif t == SCATTER:
        f = struct.unpack_from("<H", b, 4)[0] if len(b) >= 6 else 0
        name = "バブルチャート" if f & 0x01 else "散布図"
    elif t in (RADAR, RADARAREA):
        name = "レーダーチャート" + ("（塗りつぶし）" if t == RADARAREA else "")
    elif t == SURF:
        name = "表面グラフ"
    else:
        name = "補助グラフ付き円グラフ"
    return ("3-D " if three_d else "") + name


class _Label:
    def __init__(self):
        self.text = ""
        self.ref: bytes | None = None
        self.link: tuple[int, int, int] | None = None


def read_biff_chart(wb, data: bytes, start: int, part: str = "", sheet: str | None = None) -> Chart:
    """wb（XlsWorkbook）の、data の start の位置から始まるグラフの部分を読む。sheet はグラフを置いたシート
    （シートの名前のない参照の先）。"""
    ch = Chart(part=part)
    try:
        _read(wb, substream(data, start), ch, sheet)
    except (struct.error, IndexError, ValueError) as e:
        ch.error = f"グラフのレコードが壊れています（{e}）"
    return ch


def _read(wb, recs: list[tuple[int, int, bytes]], ch: Chart, sheet: str | None) -> None:
    resolve = workbook_resolver(wb, sheet)
    stack: list[int] = []  # Begin で開いたレコードの種類
    prev = None
    series: list[dict] = []  # Series レコードの順（近似曲線・誤差範囲も含む。値の控えの列の番号）
    cur_ser: dict | None = None
    groups: dict[int, str] = {}  # グラフのグループの番号（icrt）→ 種類
    cur_group: int | None = None
    labels: list[_Label] = []
    cur_label: _Label | None = None
    three_d = any(t == CHART3D for _p, t, _b in recs)
    cache: dict[tuple[int, int, int], str] = {}  # (種類 1 値・2 項目・3 大きさ, 系列の列, 点) → 字
    si = 0
    depth = 0
    for _pos, t, b in recs:
        if t == BOF:
            depth += 1
            continue
        if t == EOF:
            depth -= 1
            continue
        if depth != 1:
            continue  # グラフの中の、さらに入れ子の部分（ふつうはない）
        if t == BEGIN:
            stack.append(prev if prev is not None else 0)
            prev = None
            continue
        if t == END:
            opened = stack.pop() if stack else None
            if opened == SERIES:
                cur_ser = None
            elif opened == TEXT and cur_label is not None:
                labels.append(cur_label)
                cur_label = None
            elif opened == CHARTFORMAT:
                cur_group = None
            prev = None
            continue
        prev = t
        if t == SERIES:
            sdt_x, _sdt_y, n_x, n_y = struct.unpack_from("<HHHH", b, 0)
            cur_ser = {"ai": {}, "text": None, "crt": 0, "parent": False, "n_x": n_x, "n_y": n_y, "text_x": sdt_x == 3}
            series.append(cur_ser)
        elif t == TEXT:
            cur_label = _Label()
        elif t == BRAI:
            aid, rt = b[0], b[1]
            flags, ifmt = struct.unpack_from("<HH", b, 2)
            cce = struct.unpack_from("<H", b, 6)[0]
            rgce = b[8:8 + cce]
            if cur_label is not None and aid == 0:
                cur_label.ref = rgce if rt == 2 and rgce else None
            elif cur_ser is not None:
                cur_ser["ai"][aid] = (rt, rgce, ifmt if flags & 1 else None)
        elif t == SERIESTEXT:
            txt = _short_string(b, 2)
            if cur_label is not None:
                cur_label.text = txt
            elif cur_ser is not None:
                cur_ser["text"] = txt
        elif t == OBJECTLINK and cur_label is not None:
            cur_label.link = struct.unpack_from("<HHH", b, 0)
        elif t == SERTOCRT and cur_ser is not None:
            cur_ser["crt"] = struct.unpack_from("<H", b, 0)[0]
        elif t == SERPARENT and cur_ser is not None:
            cur_ser["parent"] = True  # 近似曲線か誤差範囲
        elif t == CHARTFORMAT:
            cur_group = struct.unpack_from("<H", b, 18)[0] if len(b) >= 20 else len(groups)
        elif t in (BAR, LINE, PIE, AREA, SCATTER, RADAR, SURF, RADARAREA, BOPPOP) and cur_group is not None:
            groups[cur_group] = _kind(t, b, three_d)
        elif t == SIINDEX:
            si = struct.unpack_from("<H", b, 0)[0]
        elif t in (NUMBER, LABEL, BOOLERR, BLANK, LABELSST) and si and not stack:
            rw, col, ixfe = struct.unpack_from("<HHH", b, 0)
            if t == NUMBER:
                v = struct.unpack_from("<d", b, 6)[0]
                cache[(si, col, rw)] = number_text(repr(v), wb.style(ixfe).num_fmt, getattr(wb, "date1904", False), ch.notes)
            elif t == LABEL:
                from ..xls.book import xlstring

                cache[(si, col, rw)] = xlstring(b, 6)[0]
            elif t == LABELSST:
                isst = struct.unpack_from("<I", b, 6)[0]
                cache[(si, col, rw)] = wb.shared_strings[isst] if isst < len(wb.shared_strings) else ""
            elif t == BOOLERR:
                v, is_err = b[6], b[7]
                # 真偽は 0 か 1 だけ。グラフの控えでは、Excel はエラー（#N/A の 0x2A など）を fError＝0 のまま書くことがあるので、
                # 0・1 でない値はエラーの番号として読む
                cache[(si, col, rw)] = _ERRORS.get(v, "#N/A") if is_err or v > 1 else ("TRUE" if v else "FALSE")
            else:
                cache[(si, col, rw)] = ""
    for k in groups.values():
        if k not in ch.kinds:
            ch.kinds.append(k)
    xy_groups = {g for g, k in groups.items() if k in ("散布図", "バブルチャート", "3-D 散布図")}
    data_series = []
    for col, s in enumerate(series):
        if s["parent"]:
            if s["text"]:
                ch.notes.append(f"近似曲線・誤差範囲：{s['text']}")
            continue
        ser = Series("", kind=groups.get(s["crt"], ""), xy=s["crt"] in xy_groups)
        rt, rgce, _f = s["ai"].get(0, (0, b"", None))
        if rt == 1 and s["text"] is not None:
            ser.name = s["text"]
        elif rt == 2 and rgce:
            ref = _decode(wb, rgce)
            if ref:
                ser.refs["name"] = ref
                got = resolve(ref)
                ser.name = " ".join(x for x in (got or []) if x)
            if not ser.name and s["text"]:
                ser.name = s["text"]  # 参照の後ろの SeriesText は、参照先の字の控え
        for aid, key, si_ in ((1, "vals", 1), (2, "cats", 2), (3, "sizes", 3)):
            rt, rgce, ifmt = s["ai"].get(aid, (0, b"", None))
            got = {rw: v for (k, c, rw), v in cache.items() if k == si_ and c == col}
            if rt == 2 and rgce:
                ref = _decode(wb, rgce)
                if ref:
                    ser.refs[{"vals": "val", "cats": "cat", "sizes": "size"}[key]] = ref
                    if not got:
                        vals = resolve(ref)
                        if vals is not None:
                            got = dict(enumerate(vals))
            if ifmt is not None and got:
                fmt = getattr(wb, "_formats", {}).get(ifmt)
                if fmt:
                    got = {i: _reformat(v, fmt, getattr(wb, "date1904", False), ser.notes) for i, v in got.items()}
            setattr(ser, key, got)
        ser.n_cats = s["n_x"]
        ser.n_vals = s["n_y"]
        data_series.append((col, ser))
    ch.series = [s for _c, s in data_series]
    for k, s in enumerate(ch.series):
        if not s.name:
            s.name = "（名前なし）" if "name" in s.refs else f"系列{k + 1}"
    col_to_ser = {col: s for col, s in data_series}
    for lab in labels:
        text = lab.text
        if not text and lab.ref:
            ref = _decode(wb, lab.ref)
            text = " ".join(x for x in (resolve(ref) or []) if x) if ref else ""
        obj = lab.link[0] if lab.link else 0
        if obj == 1:
            ch.title = text or (ch.series[0].name if len(ch.series) == 1 else "")
        elif not text:
            continue
        elif obj in (2, 3, 7):
            which = {2: "縦軸（数値）", 3: "横軸（項目）", 7: "系列軸"}[obj]
            if ch.series and all(s.xy for s in ch.series):
                which = {2: "縦軸（Y）", 3: "横軸（X）", 7: "系列軸"}[obj]
            ch.axes.append((which, text))
        elif obj == 4:
            s = col_to_ser.get(lab.link[1])
            if s is not None:
                s.labels[-1 if lab.link[2] == 0xFFFF else lab.link[2]] = text
            else:
                ch.texts.append(f"データラベル：{text}")
        elif obj == 0x0C:
            ch.texts.append(f"表示単位：{text}")
        else:
            ch.texts.append(f"グラフの中の文字：{text}")


def _decode(wb, rgce: bytes) -> str:
    try:
        return wb.decoder.decode(rgce)
    except Exception:  # noqa: BLE001
        return ""


def _reformat(v: str, fmt: str, date1904: bool = False, notes: list | None = None) -> str:
    """date1904：ブックの日付の数え方（BIFF の DATEMODE）。notes：地域で見え方が変わる表示形式の注を足す所。"""
    try:
        x = float(v.replace(",", ""))
    except ValueError:
        return v
    return number_text(repr(x), fmt, date1904, notes)


# ====================================================================== シートの中のグラフ
def sheet_charts(wb, sheet) -> list[Chart]:
    """シート（XlsWorkbook のシート）に置いたグラフ、またはグラフのシートそのもののグラフ。"""
    from ..formula.parser import addr
    from ..xls.book import _drawing_anchors

    data = wb.data
    try:
        lb = int(str(sheet.path).lstrip("#"))
    except ValueError:
        return []
    if sheet.state.startswith("non-worksheet:chartsheet"):
        return [read_biff_chart(wb, data, lb, f"#{lb}")]
    recs = substream(data, lb)
    draw: list[bytes] = []
    objs: list[int | None] = []  # OBJ の ot（順に）
    starts: dict[int, int] = {}  # OBJ の番号 → グラフの部分の始まりの位置
    depth = 0
    for i, (pos, t, b) in enumerate(recs):
        if t == BOF:
            depth += 1
            if depth == 2 and objs and objs[-1] == 0x05:
                starts[len(objs) - 1] = pos
            continue
        if t == EOF:
            depth -= 1
            continue
        if depth != 1:
            continue
        if t == 0x00EC:  # MSODRAWING（続く CONTINUE も）
            draw.append(b)
            j = i + 1
            while j < len(recs) and recs[j][1] == CONTINUE and recs[j - 1][1] in (0x00EC, CONTINUE):
                draw.append(recs[j][2])
                j += 1
        elif t == 0x005D:  # OBJ：最初の部品 FtCmo の ot
            ot = None
            if len(b) >= 6 and struct.unpack_from("<H", b, 0)[0] == 0x15:
                ot = struct.unpack_from("<H", b, 4)[0]
            objs.append(ot)
    anchors = _drawing_anchors(b"".join(draw))
    out = []
    for k, start in sorted(starts.items()):
        ch = read_biff_chart(wb, data, start, f"#{start}", sheet.name)
        a = anchors[k] if k < len(anchors) else None
        if a:
            if a.get("r1"):
                ch.where = f"{addr(a['r1'], a['c1'])}〜{addr(a['r2'], a['c2'])}"
            ch.name = a.get("name", "")
            ch.alt = a.get("alt", "")
        out.append(ch)
    return out
