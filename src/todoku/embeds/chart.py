"""グラフを、題名つきのデータの表として読む（読むだけ）。

グラフの部品（Excel の xl/charts/chartN.xml、Word の word/charts/chartN.xml、PowerPoint の ppt/charts/chartN.xml）は
DrawingML のグラフ（ECMA-376 第1部 21.2、c:chartSpace）。系列（c:ser）ごとに、名前（c:tx）・項目（c:cat。散布図は c:xVal）・
値（c:val。散布図は c:yVal）・バブルの大きさ（c:bubbleSize）を持ち、それぞれ元のセルへの参照（c:f）と、グラフが持っている
値の控え（c:strCache・c:numCache・c:multiLvlStrCache、または参照のない値 c:strLit・c:numLit）がある。
窓には控えの値を、控えの表示形式（formatCode）で見えている字にして出す（Excel・Word が画面に出す値と同じ）。
控えがないとき（作ったソフトが控えを書かなかったとき）は、resolve（参照 → 見えている字の並び）で元のセルを読む。

題名（c:title）・軸の題名・表示単位のラベル・データラベルの独自の字（c:dLbl の c:tx、セルの範囲から作ったラベル c15:datalabelsRange）・
近似曲線の名前・グラフの中に置いた図形の字（c:userShapes → cdr: の図形）も読む。新しい種類のグラフ（Office 2016 からの
じょうごグラフ・ウォーターフォールなど、cx:chartSpace）は、データ（cx:chartData の cx:strDim・cx:numDim）と系列の名前を読む。
"""

from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

from ..formula.numfmt import format_value

# グラフの種類（c:plotArea の子の要素の名前 → 日本語の名前）
KIND_JA = {
    "barChart": "棒グラフ", "bar3DChart": "3-D 棒グラフ", "lineChart": "折れ線グラフ", "line3DChart": "3-D 折れ線グラフ",
    "pieChart": "円グラフ", "pie3DChart": "3-D 円グラフ", "doughnutChart": "ドーナツグラフ", "ofPieChart": "補助グラフ付き円グラフ",
    "areaChart": "面グラフ", "area3DChart": "3-D 面グラフ", "scatterChart": "散布図", "bubbleChart": "バブルチャート",
    "radarChart": "レーダーチャート", "stockChart": "株価チャート", "surfaceChart": "等高線グラフ", "surface3DChart": "3-D 表面グラフ",
}
# 新しい種類のグラフ（cx:series の layoutId）
CX_KIND_JA = {"waterfall": "ウォーターフォール図", "funnel": "じょうごグラフ", "treemap": "ツリーマップ", "sunburst": "サンバースト",
              "boxWhisker": "箱ひげ図", "clusteredColumn": "ヒストグラム", "paretoLine": "パレート図", "regionMap": "マップ"}


@dataclass
class Series:
    """1つの系列。値などは、点の番号（0から）→ 見えている字。"""

    name: str
    kind: str = ""  # この系列のグラフの種類（複合グラフでは系列ごとに違う）
    cats: dict[int, str] = field(default_factory=dict)  # 項目（散布図・バブルチャートでは X の値）
    vals: dict[int, str] = field(default_factory=dict)  # 値（散布図では Y の値）
    sizes: dict[int, str] = field(default_factory=dict)  # バブルの大きさ
    n_cats: int = 0  # 控えに書かれた点の数（値のない点も数える）
    n_vals: int = 0
    labels: dict[int, str] = field(default_factory=dict)  # データラベルの独自の字（点の番号 → 字。-1 は系列全体）
    refs: dict[str, str] = field(default_factory=dict)  # 元のセルへの参照（name・cat・val・size）
    xy: bool = False  # 散布図・バブルチャート（項目ではなく X の値）
    notes: list[str] = field(default_factory=list)  # 近似曲線の名前など


@dataclass
class Chart:
    title: str = ""
    kinds: list[str] = field(default_factory=list)
    series: list[Series] = field(default_factory=list)
    axes: list[tuple[str, str]] = field(default_factory=list)  # (どの軸か, 題名)
    texts: list[str] = field(default_factory=list)  # グラフの中の図形の字・表示単位など、ほかの字
    notes: list[str] = field(default_factory=list)  # 読み手に伝えること（データの元・読めなかったこと）
    part: str = ""  # 部品の名前、または .xls の中の位置
    name: str = ""  # 図形としての名前（「グラフ 1」など）
    alt: str = ""  # 代替テキスト
    where: str = ""  # 置いてある所（「O6〜Z65」など）
    error: str = ""


# ====================================================================== XML の小さな道具（名前空間を問わず、要素の名前の後ろで引く）
def _ln(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _kid(el: ET.Element | None, name: str) -> ET.Element | None:
    if el is None:
        return None
    for k in el:
        if _ln(k.tag) == name:
            return k
    return None


def _kids(el: ET.Element | None, name: str) -> list[ET.Element]:
    return [k for k in el if _ln(k.tag) == name] if el is not None else []


def _val(el: ET.Element | None, name: str, default: str | None = None) -> str | None:
    k = _kid(el, name)
    if k is None:
        return default
    for a, v in k.attrib.items():
        if _ln(a) == "val":
            return v
    return default


def _attr(el: ET.Element | None, name: str) -> str | None:
    if el is None:
        return None
    for a, v in el.attrib.items():
        if _ln(a) == name:
            return v
    return None


def _rich(el: ET.Element | None) -> str:
    """書式つきの字（a:p の a:r・a:fld の a:t）。段落は改行でつなぐ。"""
    if el is None:
        return ""
    paras = [p for p in el.iter() if _ln(p.tag) == "p"]
    if not paras:
        return "".join(t.text or "" for t in el.iter() if _ln(t.tag) == "t")
    out = []
    for p in paras:
        out.append("".join(t.text or "" for t in p.iter() if _ln(t.tag) == "t"))
    return "\n".join(out).strip("\n")


def _parse(xml: bytes) -> ET.Element:
    return ET.fromstring(xml)


# ====================================================================== 値の控え
def number_text(v: str, code: str | None, date1904: bool = False) -> str:
    """控えの数値（文字で書かれた値）を、表示形式で見えている字にする。形式を再現できなければ、書かれたまま。"""
    try:
        x = float(v)
    except (TypeError, ValueError):
        return v or ""
    code = code or "General"
    try:
        from ..formula.numfmt import is_date_format

        if date1904 and is_date_format(code):
            x += 1462  # 1904 年の日付の数え方（ECMA-376 第1部 21.2.2.38 date1904）
        return format_value(x, code)
    except Exception:  # noqa: BLE001
        return v


def _cache(el: ET.Element | None, date1904: bool) -> tuple[dict[int, str], int, str | None]:
    """データの元（c:tx・c:cat・c:val など）から、控えの値：(点の番号 → 見えている字, 点の数, 参照)。"""
    if el is None:
        return {}, 0, None
    for kind in ("strRef", "numRef", "multiLvlStrRef"):
        ref = _kid(el, kind)
        if ref is not None:
            f = _kid(ref, "f")
            cache = _kid(ref, {"strRef": "strCache", "numRef": "numCache", "multiLvlStrRef": "multiLvlStrCache"}[kind])
            vals, n = _cache_values(cache, kind == "numRef", date1904)
            return vals, n, (f.text or "") if f is not None else ""
    for kind in ("strLit", "numLit"):
        lit = _kid(el, kind)
        if lit is not None:
            vals, n = _cache_values(lit, kind == "numLit", date1904)
            return vals, n, None
    v = _kid(el, "v")  # 系列の名前を直接書いたもの（c:tx/c:v）
    if v is not None:
        return {0: v.text or ""}, 1, None
    return {}, 0, None


def _cache_values(cache: ET.Element | None, numeric: bool, date1904: bool) -> tuple[dict[int, str], int]:
    if cache is None:
        return {}, 0
    n = int(_val(cache, "ptCount", "0") or 0)
    lvls = _kids(cache, "lvl")
    if lvls:
        # 複数の段の項目（c:multiLvlStrCache）：最初の段がいちばん内側。点ごとに、外の段から順に「／」でつなぐ
        levels = []
        for lv in lvls:
            pts = sorted(((int(_attr(p, "idx") or 0), _text(_kid(p, "v"))) for p in _kids(lv, "pt")))
            levels.append(pts)
        out = {}
        for i in range(n):
            parts = []
            for k, pts in enumerate(reversed(levels)):
                lab = ""
                if k == len(levels) - 1:  # いちばん内側の段は、その点の見出しだけ
                    lab = next((t for idx, t in pts if idx == i), "")
                else:
                    for idx, t in pts:  # 外の段の見出しは、次の見出しの点の前まで続く
                        if idx <= i:
                            lab = t
                        else:
                            break
                if lab:
                    parts.append(lab)
            out[i] = " ／ ".join(parts)
        return out, n
    fmt = _kid(cache, "formatCode")
    code = fmt.text if fmt is not None else None
    out = {}
    for p in _kids(cache, "pt"):
        idx = int(_attr(p, "idx") or 0)
        raw = _text(_kid(p, "v"))
        out[idx] = number_text(raw, _attr(p, "formatCode") or code, date1904) if numeric else raw
    return out, max(n, (max(out) + 1) if out else 0)


def _text(el: ET.Element | None) -> str:
    return (el.text or "") if el is not None else ""


# ====================================================================== DrawingML のグラフ（c:chartSpace）
def read_chart(xml: bytes, part: str = "", package=None, resolve=None) -> Chart:
    """グラフの部品を読む。package は部品を読むもの（.has・.read。グラフの中の図形の部品を読むのに使う）、
    resolve は控えがないときに参照（「Sheet1!$B$2:$B$5」）を見えている字の並びにするもの（なければ None）。"""
    ch = Chart(part=part)
    try:
        root = _parse(xml)
    except ET.ParseError as e:
        ch.error = f"グラフの部品を XML として読めません（{e}）"
        return ch
    if _ln(root.tag) == "chartSpace" and root.tag.startswith("{http://schemas.microsoft.com/office/drawing/2014/chartex"):
        return _read_chartex(root, ch)
    date1904 = _val(root, "date1904", "0") in ("1", "true")
    chart = _kid(root, "chart")
    if chart is None:
        ch.error = "グラフの部品に、グラフ（c:chart）がありません"
        return ch
    plot = _kid(chart, "plotArea")
    sers: list[tuple[int, Series]] = []
    for grp in list(plot) if plot is not None else []:
        name = _ln(grp.tag)
        if name not in KIND_JA:
            continue
        kind = _kind_name(grp, name)
        if kind not in ch.kinds:
            ch.kinds.append(kind)
        for s in _kids(grp, "ser"):
            order = int(_val(s, "order", _val(s, "idx", "0")) or 0)
            sers.append((order, _series(s, kind, name in ("scatterChart", "bubbleChart"), date1904, resolve)))
    # フィルターで非表示にした系列（Office 2013 からの拡張 c15:filteredBarSeries など）：グラフには出ないが、落とさず出す
    for ext in (plot.iter() if plot is not None else []):
        name = _ln(ext.tag)
        if not (name.startswith("filtered") and name.endswith("Series")):
            continue
        base = name[8:-6]
        key = base[:1].lower() + base[1:] + "Chart"  # filteredBarSeries → barChart
        for s in _kids(ext, "ser"):
            ser = _series(s, KIND_JA.get(key, "グラフ"), key in ("scatterChart", "bubbleChart"), date1904, resolve)
            ser.notes.append("フィルターで非表示にした系列です（グラフには出ていません）")
            sers.append((10_000 + int(_val(s, "order", "0") or 0), ser))
    sers.sort(key=lambda x: x[0])
    ch.series = [s for _o, s in sers]
    for k, s in enumerate(ch.series):
        if not s.name:
            # 名前の参照はあるが空のセル → 凡例も空。名前（c:tx）がない → Excel は「系列1」のように番号で出す
            s.name = "（名前なし）" if "name" in s.refs else f"系列{k + 1}"
    title = _kid(chart, "title")
    if title is not None:
        ch.title = _title_text(title, date1904)
        if not ch.title and len(ch.series) == 1:
            ch.title = ch.series[0].name  # 題名の字がなく系列が1つなら、Excel は系列の名前を題名に出す
    for ax in list(plot) if plot is not None else []:
        if _ln(ax.tag) not in ("catAx", "valAx", "dateAx", "serAx"):
            continue
        which = _axis_name(ax, ch)
        t = _kid(ax, "title")
        if t is not None:
            txt = _title_text(t, date1904)
            if txt:
                ch.axes.append((which, txt))
        du = _kid(ax, "dispUnits")
        if du is not None:
            lbl = _kid(du, "dispUnitsLbl")
            txt = _title_text(lbl, date1904) if lbl is not None else ""
            unit = _val(du, "builtInUnit")
            if txt or unit:
                ch.texts.append(f"{which}の表示単位：{txt or unit}")
    if plot is not None:
        dt = _kid(plot, "dTable")
        if dt is not None:
            ch.notes.append("グラフの下にデータテーブル（この表と同じ値）が表示されています")
    ps = _kid(root, "pivotSource")
    if ps is not None:
        ch.notes.append(f"ピボットグラフ（元のピボットテーブル：{_text(_kid(ps, 'name'))}）")
    if package is not None and part:
        ch.texts += _user_shapes(root, part, package)
    return ch


def _kind_name(grp: ET.Element, name: str) -> str:
    base = KIND_JA[name]
    if name in ("barChart", "bar3DChart"):
        base = ("横棒グラフ" if _val(grp, "barDir") == "bar" else "縦棒グラフ") + ("（3-D）" if name == "bar3DChart" else "")
    if name == "ofPieChart":
        base = "補助円グラフ付き円グラフ" if _val(grp, "ofPieType") == "pie" else "補助縦棒付き円グラフ"
    g = _val(grp, "grouping")
    if g == "stacked":
        base = "積み上げ" + base
    elif g == "percentStacked":
        base = "100% 積み上げ" + base
    return base


def _title_text(title: ET.Element | None, date1904: bool) -> str:
    if title is None:
        return ""
    tx = _kid(title, "tx")
    if tx is None:
        return ""
    rich = _kid(tx, "rich")
    if rich is not None:
        return _rich(rich)
    vals, _n, _f = _cache(tx, date1904)
    return " ".join(v for _i, v in sorted(vals.items()) if v)


def _series(s: ET.Element, kind: str, xy: bool, date1904: bool, resolve) -> Series:
    ser = Series("", kind=kind, xy=xy)
    names, _n, ref = _cache(_kid(s, "tx"), date1904)
    ser.name = " ".join(v for _i, v in sorted(names.items()) if v)
    if not ser.name:  # 名前の参照先がフィルターで隠れているときの、控えの名前（c15:filteredSeriesTitle）
        ft = next((e for e in s.iter() if _ln(e.tag) == "filteredSeriesTitle"), None)
        if ft is not None:
            ser.name = " ".join(v.text or "" for v in ft.iter() if _ln(v.tag) == "v" and v.text)
    if not ser.name and ref and resolve is not None:
        ser.name = " ".join(t for t in (resolve(ref) or []) if t)
    if ref:
        ser.refs["name"] = ref
    for key, tags in (("cats", ("cat", "xVal")), ("vals", ("val", "yVal")), ("sizes", ("bubbleSize",))):
        el = next((_kid(s, t) for t in tags if _kid(s, t) is not None), None)
        if el is None:
            continue
        vals, n, ref = _cache(el, date1904)
        if ref:
            ser.refs[{"cats": "cat", "vals": "val", "sizes": "size"}[key]] = ref
            has_cache = any(_ln(x.tag) in ("strCache", "numCache", "multiLvlStrCache") for r_ in el for x in r_)
            if not has_cache and resolve is not None:
                got = resolve(ref)
                if got is not None:
                    vals, n = {i: t for i, t in enumerate(got)}, len(got)
                    ser.notes.append(f"{'項目' if key == 'cats' else '値'}は控えがないので、参照先のセル（{ref}）から読みました")
        setattr(ser, key, vals)
        if key == "cats":
            ser.n_cats = n
        elif key == "vals":
            ser.n_vals = n
    lbls = _kid(s, "dLbls")
    for d in _kids(lbls, "dLbl"):
        tx = _kid(d, "tx")
        if tx is None:
            continue
        idx = int(_val(d, "idx", "-1") or -1)
        rich = _kid(tx, "rich")
        txt = _rich(rich) if rich is not None else " ".join(v for _i, v in sorted(_cache(tx, date1904)[0].items()))
        if txt:
            ser.labels[idx] = txt
    for ext in s.iter():
        if _ln(ext.tag) == "datalabelsRange":  # セルの範囲の字をデータラベルにしたもの（Office 2013 からの拡張）
            cache = _kid(ext, "dlblRangeCache")
            vals, _n = _cache_values(cache, False, date1904)
            for i, t in vals.items():
                if t and i not in ser.labels:
                    ser.labels[i] = t
    for tl in _kids(s, "trendline"):
        nm = _text(_kid(tl, "name"))
        lbl = _kid(tl, "trendlineLbl")
        lt = _title_text(lbl, date1904) if lbl is not None else ""
        if nm or lt:
            ser.notes.append("近似曲線：" + "／".join(x for x in (nm, lt) if x))
    return ser


def _axis_name(ax: ET.Element, ch: Chart) -> str:
    pos = _val(ax, "axPos", "")
    side = {"b": "横軸", "t": "横軸（上）", "l": "縦軸", "r": "縦軸（右）"}.get(pos or "", "軸")
    kind = {"catAx": "項目", "dateAx": "日付", "valAx": "数値", "serAx": "系列"}[_ln(ax.tag)]
    hidden = "・非表示" if _val(ax, "delete", "0") in ("1", "true") else ""
    return f"{side}（{kind}{hidden}）"


def _user_shapes(root: ET.Element, part: str, package) -> list[str]:
    """グラフの中に置いた図形の字（c:userShapes → グラフの図形の部品 cdr:）。"""
    us = _kid(root, "userShapes")
    rid = _attr(us, "id") if us is not None else None
    if not rid:
        return []
    target = rels_of(package, part).get(rid)
    if not target or not package.has(target):
        return []
    try:
        d = _parse(package.read(target))
    except ET.ParseError:
        return []
    out = []
    for sp in d.iter():
        if _ln(sp.tag) == "txBody":
            t = _rich(sp)
            if t.strip():
                out.append(f"グラフの中の文字：{t}")
    return out


def rels_of(package, part: str) -> dict[str, str]:
    """部品の関係（_rels/….rels）：Id → 部品の名前（外部のものは除く）。"""
    d, base = posixpath.split(part)
    rp = posixpath.join(d, "_rels", base + ".rels")
    if not package.has(rp):
        return {}
    out = {}
    try:
        root = _parse(package.read(rp))
    except ET.ParseError:
        return {}
    for r in root:
        if r.get("TargetMode") == "External":
            continue
        t = r.get("Target", "")
        full = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
        out[r.get("Id", "")] = full
    return out


def rel_types(package, part: str) -> dict[str, tuple[str, str]]:
    """部品の関係：Id → (種類の最後の語, 部品の名前)。"""
    d, base = posixpath.split(part)
    rp = posixpath.join(d, "_rels", base + ".rels")
    if not package.has(rp):
        return {}
    out = {}
    try:
        root = _parse(package.read(rp))
    except ET.ParseError:
        return {}
    for r in root:
        t = r.get("Target", "")
        if r.get("TargetMode") == "External":
            full = t
        else:
            full = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(d, t))
        out[r.get("Id", "")] = (r.get("Type", "").rsplit("/", 1)[-1], full)
    return out


# ====================================================================== 新しい種類のグラフ（cx:chartSpace）
def _read_chartex(root: ET.Element, ch: Chart) -> Chart:
    data: dict[str, dict[str, tuple[dict[int, str], str]]] = {}
    cd = _kid(root, "chartData")
    for d in _kids(cd, "data"):
        did = _attr(d, "id") or "0"
        dims = {}
        for dim in d:
            if _ln(dim.tag) not in ("strDim", "numDim"):
                continue
            typ = _attr(dim, "type") or ""
            f = _kid(dim, "f")
            lvls = _kids(dim, "lvl")
            vals: dict[int, str] = {}
            if lvls:
                code = _attr(lvls[0], "formatCode")
                # 段が複数あるとき（ツリーマップ・サンバースト）：最初の段がいちばん内側。外の段から「／」でつなぐ
                per = []
                for lv in lvls:
                    per.append({int(_attr(p, "idx") or 0): (p.text or "") for p in _kids(lv, "pt")})
                n = max((max(x) + 1 for x in per if x), default=0)
                for i in range(n):
                    parts = [x.get(i, "") for x in reversed(per)]
                    if _ln(dim.tag) == "numDim":
                        vals[i] = number_text(per[0].get(i, ""), code) if i in per[0] else ""
                    else:
                        vals[i] = " ／ ".join(p for p in parts if p)
            dims[typ] = (vals, (f.text or "") if f is not None else "")
        data[did] = dims
    chart = _kid(root, "chart")
    title = _kid(chart, "title")
    if title is not None:
        ch.title = " ".join(t.text or "" for t in title.iter() if _ln(t.tag) in ("v", "t") and t.text)
    plot = _kid(_kid(chart, "plotArea"), "plotAreaRegion")
    for k, s in enumerate(_kids(plot, "series")):
        kind = CX_KIND_JA.get(_attr(s, "layoutId") or "", _attr(s, "layoutId") or "グラフ")
        if kind not in ch.kinds:
            ch.kinds.append(kind)
        tx = _kid(s, "tx")
        name = " ".join(t.text or "" for t in tx.iter() if _ln(t.tag) == "v" and t.text) if tx is not None else ""
        ser = Series(name or f"系列{k + 1}", kind=kind)
        did = _val(s, "dataId", "0") or "0"
        dims = data.get(did, {})
        for typ, (vals, f) in dims.items():
            if typ == "cat":
                ser.cats, ser.n_cats = vals, len(vals)
                if f:
                    ser.refs["cat"] = f
            elif typ in ("val", "size", "x", "y"):
                if not ser.vals:
                    ser.vals, ser.n_vals = vals, len(vals)
                    if f:
                        ser.refs["val"] = f
        ch.series.append(ser)
    for ax in (_kids(_kid(chart, "plotArea"), "axis") if chart is not None else []):
        t = _kid(ax, "title")
        if t is not None:
            txt = " ".join(x.text or "" for x in t.iter() if _ln(x.tag) in ("v", "t") and x.text)
            if txt:
                ch.axes.append(("軸", txt))
    return ch


# ====================================================================== 窓（JSON の形と Markdown の表）
def _md(s) -> str:
    return str(s).replace("|", "\\|").replace("\r", " ").replace("\n", " ")


def _count(s: Series) -> int:
    """系列の点の数（控えに書かれた数と、値のある最後の点の、大きいほう）。"""
    return max(s.n_cats, s.n_vals, (max(s.cats) + 1) if s.cats else 0, (max(s.vals) + 1) if s.vals else 0,
               (max(s.sizes) + 1) if s.sizes else 0)


def chart_dict(ch: Chart, n: int) -> dict:
    """窓の JSON の形（Markdown もこの形から作る）。値の並びは、点の番号の順（値のない点は空の字）。"""
    sers = []
    for s in ch.series:
        k = _count(s)
        d = {"name": s.name, "kind": s.kind, "categories": [s.cats.get(i, "") for i in range(k)] if s.cats else [],
             "values": [s.vals.get(i, "") for i in range(k)]}
        if s.xy:
            d["xy"] = True
        if s.sizes:
            d["sizes"] = [s.sizes.get(i, "") for i in range(k)]
        if s.labels:
            d["labels"] = {str(i): t for i, t in sorted(s.labels.items())}
        if s.refs:
            d["refs"] = dict(s.refs)
        if s.notes:
            d["notes"] = list(s.notes)
        sers.append(d)
    out = {"n": n, "title": ch.title, "kinds": list(ch.kinds), "part": ch.part, "series": sers,
           "axes": [{"axis": a, "title": t} for a, t in ch.axes], "texts": list(ch.texts), "notes": list(ch.notes)}
    for key in ("name", "alt", "where", "error"):
        if getattr(ch, key):
            out[key] = getattr(ch, key)
    return out


def table_rows(d: dict) -> list[tuple[str, list[list[str]]]]:
    """グラフの値の表：[(表の題, 行の並び（1行目は見出し）)]。系列の項目がそろっていれば1つの表、違えば系列ごと。
    項目を持たない系列は、項目軸の見出し（項目を持つ最初の系列の項目）を使う（Excel と同じ）。"""
    out = []
    plain = [s for s in d["series"] if not s.get("xy")]
    base = next((s["categories"] for s in plain if s["categories"]), [])
    groups: list[tuple[tuple, list[dict]]] = []
    for s in plain:
        n = max(len(s["values"]), len(s["categories"]), len(s.get("sizes") or []))
        cats = s["categories"] or (base if len(base) >= n else [])
        key = tuple(cats) if cats else ("#", n)
        for k, g in groups:
            if k == key:
                g.append(s)
                break
        else:
            groups.append((key, [s]))
    for key, g in groups:
        cats = list(key) if key[:1] != ("#",) else []
        n = max(max(len(s["values"]) for s in g), len(cats))
        rows = [["項目"] + [s["name"] for s in g]]
        for i in range(n):
            label = cats[i] if i < len(cats) else str(i + 1)
            rows.append([label] + [s["values"][i] if i < len(s["values"]) else "" for s in g])
        out.append(("" if len(groups) == 1 else "・".join(s["name"] for s in g), rows))
    for s in d["series"]:
        if not s.get("xy"):
            continue
        n = max(len(s["values"]), len(s["categories"]))
        rows = [["X", "Y"] + (["大きさ"] if s.get("sizes") else [])]
        for i in range(n):
            row = [s["categories"][i] if i < len(s["categories"]) else str(i + 1), s["values"][i] if i < len(s["values"]) else ""]
            if s.get("sizes"):
                row.append(s["sizes"][i] if i < len(s["sizes"]) else "")
            rows.append(row)
        out.append((s["name"], rows))
    return out


def chart_markdown(d: dict, level: str = "###") -> list[str]:
    """グラフ（chart_dict の形）を、題名と値の表にした Markdown の行。"""
    kinds = "・".join(d.get("kinds") or []) or "グラフ"
    title = d.get("title") or "（題名なし）"
    place = d.get("where", "")
    out = [f"{level} グラフ{d['n']}：{_md(title)}（{kinds}{'、' + place if place else ''}）", ""]
    if d.get("error"):
        return out + [f"- このグラフは読めませんでした：{d['error']}", ""]
    info = []
    if d.get("name"):
        info.append(f"名前「{_md(d['name'])}」")
    if d.get("alt"):
        info.append(f"代替テキスト「{_md(d['alt'])}」")
    if info:
        out.append("- " + "、".join(info))
    for a in d.get("axes") or []:
        out.append(f"- {a['axis']}の題名：{_md(a['title'])}")
    for t in d.get("texts") or []:
        out.append(f"- {_md(t)}")
    for s in d["series"]:
        if s.get("labels"):
            parts = []
            for i, t in s["labels"].items():
                i = int(i)
                head = "系列全体" if i < 0 else (s["categories"][i] if i < len(s["categories"]) and s["categories"][i] else f"{i + 1}番目")
                parts.append(f"{_md(head)}「{_md(t)}」")
            out.append(f"- データラベル（{_md(s['name'])}）：" + "、".join(parts))
        for t in s.get("notes") or []:
            out.append(f"- {_md(s['name'])}：{_md(t)}")
    for t in d.get("notes") or []:
        out.append(f"- {_md(t)}")
    if not d["series"]:
        out.append("- 系列（データ）がありません")
    if out[-1] != "":
        out.append("")
    for sub, rows in table_rows(d):
        if sub:
            out += [f"系列「{_md(sub)}」", ""]
        out.append("| " + " | ".join(_md(x).strip() for x in rows[0]) + " |")
        out.append("|" + "---|" * len(rows[0]))
        for r in rows[1:]:
            out.append("| " + " | ".join(_md(x).strip() for x in r) + " |")
        out.append("")
    return out


_REF = re.compile(r"^(?:'((?:[^']|'')+)'|([^!'\s(),]+))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)$")


def split_ref(text: str) -> list[tuple[str | None, int, int, int, int]]:
    """参照の字（「'Sheet 1'!$B$2:$B$5」「(Sheet1!$A$1,Sheet1!$C$1)」）を、(シート, 行1, 列1, 行2, 列2) の並びにする。"""
    from ..formula.parser import parse_cell

    t = text.strip()
    if t.startswith("(") and t.endswith(")"):
        t = t[1:-1]
    out = []
    for part in _split_top(t):
        part = part.strip()
        m = _REF.match(part)
        if m:
            sheet = (m[1].replace("''", "'") if m[1] is not None else m[2])
            rng = m[3].replace("$", "")
        elif re.fullmatch(r"\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?", part):
            sheet, rng = None, part.replace("$", "")
        else:
            continue
        a, _, b = rng.partition(":")
        r1, c1 = parse_cell(a)
        r2, c2 = parse_cell(b) if b else (r1, c1)
        out.append((sheet, min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)))
    return out


def _split_top(t: str) -> list[str]:
    out, depth, q, cur = [], 0, False, []
    for ch in t:
        if ch == "'":
            q = not q
        elif not q and ch == "(":
            depth += 1
        elif not q and ch == ")":
            depth -= 1
        if ch == "," and not q and depth == 0:
            out.append("".join(cur))
            cur = []
            continue
        cur.append(ch)
    out.append("".join(cur))
    return out


def workbook_resolver(wb, default_sheet: str | None = None):
    """ブックのセルを読む resolve：参照 → 見えている字の並び（範囲の左上から行ごと）。読めない参照は None。
    シートの名前のない参照は default_sheet（なければ最初のシート）のセル。"""
    from ..xlsx.window import display

    def resolve(ref: str):
        try:
            areas = split_ref(ref)
        except Exception:  # noqa: BLE001
            return None
        if not areas:
            return None
        out = []
        for sheet, r1, c1, r2, c2 in areas:
            try:
                sh = wb.sheet(sheet or default_sheet) if (sheet or default_sheet) else wb.sheets[0]
            except (KeyError, IndexError):
                return None
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 100_000:
                return None
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    cell = sh.cells.get((r, c))
                    try:
                        out.append(display(wb, cell) if cell is not None else "")
                    except Exception:  # noqa: BLE001
                        out.append(str(cell.value) if cell is not None and cell.value is not None else "")
        return out

    return resolve
