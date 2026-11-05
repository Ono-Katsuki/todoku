"""PDFの欄に付いた Acrobat の標準のスクリプト（AFDate_FormatEx・AFNumber_Format・AFSimple_Calculate など）を、
JavaScript を動かさずに、同じ規則で決定的に行う。

- 書式（F）と入力の制限（K）：書く値を、その欄の書式どおりの文字にする。書式に合わない値は、Acrobat と同じく受け付けない
- 値の確認（V）：AFRange_Validate の範囲
- 計算（C）：AFSimple_Calculate（SUM・AVG・PRD・MIN・MAX）。フォームの計算の順（/CO）に計算し直す
これら以外の独自のスクリプトは動かさず、検査で「独自のスクリプト」として知らせる。
"""

from __future__ import annotations

import datetime as dt
import re

from .objects import PdfString, Stream, decode_text

MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
OLD_DATE_FORMATS = ["m/d", "m/d/yy", "mm/dd/yy", "mm/yy", "d-mmm", "d-mmm-yy", "dd-mmm-yy", "yy-mm-dd", "mmm-yy",
                    "mmmm-yy", "mmm d, yyyy", "mmmm d, yyyy", "m/d/yy h:MM tt", "m/d/yy HH:MM"]


class JsError(ValueError):
    pass


def actions(doc, node) -> dict[str, str]:
    """欄の追加の動作（/AA）のスクリプト：F 書式・K 入力・V 確認・C 計算。"""
    out = {}
    aa = doc.resolve(node.get("AA")) if isinstance(node, dict) else None
    if not isinstance(aa, dict):
        return out
    for k in ("F", "K", "V", "C"):
        act = doc.resolve(aa.get(k))
        if not isinstance(act, dict):
            continue
        js = doc.resolve(act.get("JS"))
        if isinstance(js, Stream):
            js = doc.decode(js)
        if isinstance(js, (bytes, PdfString)):
            js = decode_text(js)
        if isinstance(js, str) and js.strip():
            out[k] = js.strip()
    return out


_CALL = re.compile(r"^\s*(AF\w+)\s*\((.*)\)\s*;?\s*$", re.S)


_JS_ESC = {"n": "\n", "t": "\t", "r": "\r", "b": "\b", "f": "\f", "v": "\v", "0": "\0"}


def js_unescape(s: str) -> str:
    """JavaScript の文字列の中のエスケープ（\\u00a5・\\x41・\\n・\\"・行の続き など。ECMA-262 の文字列リテラル）を字にする。"""
    def rep(m):
        e = m[1]
        if e[0] == "u":
            return chr(int(e[2:-1] if e[1] == "{" else e[1:], 16))
        if e[0] == "x":
            return chr(int(e[1:], 16))
        if e in ("\n", "\r\n", "\r", "\u2028", "\u2029"):
            return ""  # 行の続き
        return _JS_ESC.get(e, e)

    return re.sub(r"\\(u\{[0-9A-Fa-f]+\}|u[0-9A-Fa-f]{4}|x[0-9A-Fa-f]{2}|\r\n|[\s\S])", rep, s)


def calc_sources(js: str) -> list[str]:
    """計算のスクリプト（AFSimple_Calculate）の、元になる欄の名前。読めなければ空。"""
    p = parse_call(js)
    if p is None or p[0] != "AFSimple_Calculate" or len(p[1]) < 2:
        return []
    names = p[1][1]
    return list(names) if isinstance(names, list) else [x.strip() for x in str(names).split(",") if x.strip()]


def _args(s: str) -> list:
    out = []
    for m in re.finditer(r'"((?:[^"\\]|\\.)*)"|\'((?:[^\'\\]|\\.)*)\'|new\s+Array\s*\(([^)]*)\)|(-?\d+(?:\.\d+)?)|(true|false)', s):
        if m[1] is not None or m[2] is not None:
            out.append(js_unescape(m[1] if m[1] is not None else m[2]))
        elif m[3] is not None:
            out.append([js_unescape(a if a is not None else b)
                        for a, b in re.findall(r'"((?:[^"\\]|\\.)*)"|\'((?:[^\'\\]|\\.)*)\'', m[3])])
        elif m[4] is not None:
            out.append(float(m[4]) if "." in m[4] else int(m[4]))
        else:
            out.append(m[5] == "true")
    return out


def parse_call(js: str):
    m = _CALL.match(js)
    if not m:
        return None
    return m[1], _args(m[2])


SPECIAL_NAMES = ("郵便番号（5桁）", "郵便番号（9桁）", "電話番号", "社会保障番号")


def special_kind(args) -> int | None:
    """AFSpecial_* の引数が、決まった形の番号（0〜3）なら、その番号。引数がなければ 0。
    字の形（'+65 9999 9999' のような字の並び）やほかの番号なら None：形を決めつけない。"""
    if not args:
        return 0
    a = args[0]
    if isinstance(a, str):
        a = a.strip()
        if not a.isascii() or not a.isdigit():
            return None
    try:
        k = int(a)
    except (TypeError, ValueError):
        return None
    return k if k in range(len(SPECIAL_NAMES)) and float(args[0]) == k else None


def known(js: str) -> bool:
    p = parse_call(js)
    return p is not None and p[0] in KNOWN


KNOWN = {"AFDate_FormatEx", "AFDate_KeystrokeEx", "AFDate_Format", "AFDate_Keystroke", "AFTime_Format", "AFTime_Keystroke",
         "AFNumber_Format", "AFNumber_Keystroke", "AFPercent_Format", "AFPercent_Keystroke", "AFSpecial_Format",
         "AFSpecial_Keystroke", "AFSpecial_KeystrokeEx", "AFRange_Validate", "AFSimple_Calculate"}


# ====================================================================== 日付
def _parse_date(value: str, fmt: str | None) -> dt.datetime:
    try:
        return _parse_date_(value, fmt)
    except (ValueError, OverflowError) as e:
        if isinstance(e, JsError):
            raise
        raise JsError(f"「{str(value).strip()}」は、ありえない日付です（月は 1〜12、日はその月の日数まで）") from None


def _parse_date_(value: str, fmt: str | None) -> dt.datetime:
    v = str(value).strip()
    if not v:
        raise JsError("日付が空です")
    # 和暦（令和8年4月1日）
    try:
        from ..xlsx.edit import parse_date_text

        d = parse_date_text(v)
        if d is not None:
            if isinstance(d, dt.datetime):
                return d
            return dt.datetime(d.year, d.month, d.day)
    except Exception:  # noqa: BLE001
        pass
    m = re.fullmatch(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:[ T](\d{1,2}):(\d{2}))?", v)
    if m:
        return dt.datetime(int(m[1]), int(m[2]), int(m[3]), int(m[4] or 0), int(m[5] or 0))
    if fmt:
        d = _parse_with(v, fmt)
        if d is not None:
            return d
    # 月の名前（2 November 2026、Nov 2, 2026）
    for pat, order in ((r"(\d{1,2})[ -]([A-Za-z]{3,})[ -,]*(\d{2,4})", "dmy"), (r"([A-Za-z]{3,})[ -](\d{1,2}),?[ -](\d{2,4})", "mdy")):
        m = re.fullmatch(pat, v)
        if m:
            a, b, c = m.groups()
            mon_s, day_s = (b, a) if order == "dmy" else (a, b)
            mon = next((i + 1 for i, n in enumerate(MONTHS) if n.lower().startswith(mon_s.lower()[:3])), None)
            if mon:
                y = int(c)
                return dt.datetime(y + 2000 if y < 50 else (y + 1900 if y < 100 else y), mon, int(day_s))
    raise JsError(f"「{v}」を日付として読めません")


def _parse_with(v: str, fmt: str) -> dt.datetime | None:
    """書式（dd-mmm-yyyy など）の並びで、数字と月の名前を読む。"""
    order = re.findall(r"yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|HH|H|hh|h|MM|M|ss|s|tt", fmt)
    parts = re.findall(r"\d+|[A-Za-z]+", v)
    vals = {}
    k = 0
    for tok in order:
        if tok in ("ddd", "dddd"):
            if k < len(parts) and parts[k].isalpha():
                k += 1
            continue
        if k >= len(parts):
            break
        p = parts[k]
        k += 1
        if tok in ("mmm", "mmmm"):
            mon = next((i + 1 for i, n in enumerate(MONTHS) if n.lower().startswith(p.lower()[:3])), None)
            if mon is None:
                return None
            vals["m"] = mon
        elif tok in ("m", "mm") and p.isdigit():
            vals["m"] = int(p)
        elif tok in ("d", "dd") and p.isdigit():
            vals["d"] = int(p)
        elif tok in ("yy", "yyyy") and p.isdigit():
            y = int(p)
            vals["y"] = y + 2000 if y < 50 else (y + 1900 if y < 100 else y)
        elif tok in ("H", "HH", "h", "hh") and p.isdigit():
            vals["H"] = int(p)
        elif tok in ("M", "MM") and p.isdigit():
            vals["M"] = int(p)
        elif tok == "tt" and p.lower() in ("pm", "p") and "H" in vals and vals["H"] < 12:
            vals["H"] += 12
    if "m" not in vals or ("d" not in vals and "y" not in vals):
        return None
    try:
        return dt.datetime(vals.get("y", dt.date.today().year), vals["m"], vals.get("d", 1), vals.get("H", 0), vals.get("M", 0))
    except ValueError:
        return None


def format_date(d: dt.datetime, fmt: str) -> str:
    def rep(m):
        t = m[0]
        return {
            "yyyy": f"{d.year:04d}", "yy": f"{d.year % 100:02d}", "mmmm": MONTHS[d.month - 1], "mmm": MONTHS[d.month - 1][:3],
            "mm": f"{d.month:02d}", "m": str(d.month), "dddd": DAYS[d.weekday()], "ddd": DAYS[d.weekday()][:3],
            "dd": f"{d.day:02d}", "d": str(d.day), "HH": f"{d.hour:02d}", "H": str(d.hour),
            "hh": f"{(d.hour % 12) or 12:02d}", "h": str((d.hour % 12) or 12), "MM": f"{d.minute:02d}", "M": str(d.minute),
            "ss": f"{d.second:02d}", "s": str(d.second), "tt": "pm" if d.hour >= 12 else "am",
        }[t]

    return re.sub(r"yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|HH|H|hh|h|MM|M|ss|s|tt", rep, fmt)


# ====================================================================== 数
def _number(value) -> float:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    t = str(value).strip()
    t = t.translate(str.maketrans("０１２３４５６７８９．，－＋", "0123456789.,-+"))
    neg = t.startswith("(") and t.endswith(")")
    t = re.sub(r"[^\d.\-+eE]", "", t.replace(",", ""))
    if not t:
        raise JsError(f"「{value}」は数ではありません")
    try:
        x = float(t)
    except ValueError:
        raise JsError(f"「{value}」は数ではありません") from None
    return -x if neg else x


def format_number(x: float, ndec: int, sep: int, neg: int, curr: str = "", prepend: bool = True) -> str:
    s = f"{abs(x):,.{int(ndec)}f}"
    if sep == 1:
        s = s.replace(",", "")
    elif sep == 2:
        s = s.replace(",", "_").replace(".", ",").replace("_", ".")
    elif sep == 3:
        s = s.replace(",", "").replace(".", ",")
    elif sep == 4:
        s = s.replace(",", "'")
    if curr:
        s = curr + s if prepend else s + curr
    if x < 0:
        s = f"({s})" if neg in (2, 3) else "-" + s
    return s


# ====================================================================== 書く値を決める
def commit(js: dict[str, str], value) -> str:
    """書式・入力の制限・値の確認のスクリプトに従って、見た目に出る字を決める。合わなければ JsError。"""
    return commit_pair(js, value)[1]


def _plain(x: float) -> str:
    if float(x).is_integer():
        return str(int(x))
    s = repr(float(x))
    if "e" in s or "E" in s:
        from decimal import Decimal

        s = format(Decimal(s), "f")
    return s


def commit_pair(js: dict[str, str], value) -> tuple[str, str]:
    """(欄の値 /V, 見た目に出る字)。Acrobat と同じく、書式（F）は見た目だけを整え、値はそのまま持つ：
    数・百分率・郵便番号などは、値は数字だけの形（1234.5・0.25・1234567）、見た目は書式どおり（¥1,234.50・25.00%）。
    日付・時刻は、書式（F）がその値を読み直して見た目を作るので、値も書式の形（yyyy/mm/dd など）にする。"""
    text = "" if value is None else str(value)
    if text == "":
        return text, text
    f = parse_call(js["F"]) if "F" in js else None
    k = parse_call(js["K"]) if "K" in js else None
    fn = (f or k or (None, []))[0]
    args = (f or k or (None, []))[1]
    out = text
    raw = text
    x = None
    if fn in ("AFDate_FormatEx", "AFDate_KeystrokeEx", "AFDate_Format", "AFDate_Keystroke"):
        fmt = args[0] if args and isinstance(args[0], str) else (OLD_DATE_FORMATS[int(args[0])] if args else "mm/dd/yyyy")
        d = _parse_date(text, fmt)
        out = raw = format_date(d, fmt)
    elif fn in ("AFTime_Format", "AFTime_Keystroke"):
        fmt = ["HH:MM", "h:MM tt", "HH:MM:ss", "h:MM:ss tt"][int(args[0]) if args else 0]
        m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([ap]m)?", text.strip(), re.I)
        if not m:
            raise JsError(f"「{text}」を時刻として読めません")
        h = int(m[1]) + (12 if (m[4] or "").lower() == "pm" and int(m[1]) < 12 else 0)
        try:
            out = raw = format_date(dt.datetime(2000, 1, 1, h, int(m[2]), int(m[3] or 0)), fmt)
        except ValueError:
            raise JsError(f"「{text}」は、ありえない時刻です") from None
    elif fn in ("AFNumber_Format", "AFNumber_Keystroke"):
        x = _number(text)
        nd, sep, neg = (list(args) + [2, 0, 0])[:3]
        curr = args[4] if len(args) > 4 and isinstance(args[4], str) else ""
        pre = args[5] if len(args) > 5 else True
        out = format_number(x, int(nd), int(sep), int(neg), curr, bool(pre))
        raw = _plain(x)
    elif fn in ("AFPercent_Format", "AFPercent_Keystroke"):
        x = _number(text.rstrip("%％"))
        if "%" in text or "％" in text:
            x /= 100.0
        nd, sep = (list(args) + [2, 0])[:2]
        out = format_number(x * 100, int(nd), int(sep), 0) + "%"
        raw = _plain(x)
    elif fn in ("AFSpecial_Format", "AFSpecial_Keystroke") and special_kind(args) is not None:
        kind = special_kind(args)
        digits = re.sub(r"\D", "", text.translate(str.maketrans("０１２３４５６７８９", "0123456789")))
        need = {0: 5, 1: 9, 2: 10, 3: 9}.get(kind)
        if kind == 2 and len(digits) == 7:
            need = 7
        if need is None or len(digits) != need:
            raise JsError(f"「{text}」は、この欄の形（{SPECIAL_NAMES[kind]}）に合いません")
        out = {0: digits, 1: f"{digits[:5]}-{digits[5:]}", 3: f"{digits[:3]}-{digits[3:5]}-{digits[5:]}",
               2: (f"({digits[:3]}) {digits[3:6]}-{digits[6:]}" if len(digits) == 10 else f"{digits[:3]}-{digits[3:]}")}[kind]
        raw = digits
    if "V" in js:
        v = parse_call(js["V"])
        if v and v[0] == "AFRange_Validate":
            a = list(v[1]) + [False, 0, False, 0]
            xv = x if x is not None else _number(out)
            if a[0] and xv < a[1]:
                raise JsError(f"{xv:g} は、{a[1]} 以上である必要があります")
            if a[2] and xv > a[3]:
                raise JsError(f"{xv:g} は、{a[3]} 以下である必要があります")
    return raw, out


def simple_calculate(js: str, get) -> float | None:
    """AFSimple_Calculate("SUM", new Array("a", "b"))。get(名前) は欄の値（数でなければ 0）。"""
    p = parse_call(js)
    if p is None or p[0] != "AFSimple_Calculate" or len(p[1]) < 2:
        return None
    op = str(p[1][0]).upper()
    names = p[1][1] if isinstance(p[1][1], list) else [x.strip() for x in str(p[1][1]).split(",")]
    vals = []
    for n in names:
        v = get(n)
        try:
            vals.append(_number(v) if v not in (None, "") else 0.0)
        except JsError:
            vals.append(0.0)
    if not vals:
        return 0.0
    if op == "SUM":
        return sum(vals)
    if op == "AVG":
        return sum(vals) / len(vals)
    if op == "PRD":
        r = 1.0
        for v in vals:
            r *= v
        return r
    if op == "MIN":
        return min(vals)
    if op == "MAX":
        return max(vals)
    return None
