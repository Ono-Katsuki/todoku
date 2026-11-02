"""XFA の書式（picture clause、仕様 XFA 3.3 26章）：日付・文字・数の書式で、値を整える・読み取る・合うかを確かめる。

扱う記号：
- date{…}：D DD J JJJ M MM MMM MMMM E EEE EEEE YY YYYY と区切りの字
- text{…}：9（数字）A（英字）X（何でも1字）O・0（英数字）t（タブ）
- num{…}：9（数字）z Z（数字か、なし）, .（区切り・小数点）$ s S ( ) % と区切りの字
- null{}・zero{}、決まった書式（date.short{}・num.integer{} など。localeSet の定義、なければ en_US）、| で区切った候補
決められない書式（時刻・暦の記号など）は None を返し、確かめずに、そのことを知らせる。
"""

from __future__ import annotations

import datetime as _dt
import re

EN_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November",
             "December"]
EN_ABBR = [m[:3] for m in EN_MONTHS]
EN_DAYS = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
EN_DATE = {"short": "M/D/YY", "medium": "MMM D, YYYY", "long": "MMMM D, YYYY", "full": "EEEE, MMMM D, YYYY",
           "default": "MMM D, YYYY"}
EN_NUM = {"numeric": "z,zz9.zzz", "integer": "z,zzz,zzz,zz9", "decimal": "z,zz9.zzz", "currency": "$z,zz9.99|($z,zz9.99)",
          "percent": "z,zz9%"}


class Locale:
    """書式に使う、地域の決まり（月の名前・決まった書式）。localeSet の locale 要素から。"""

    def __init__(self, el=None):
        self.months, self.abbr, self.days = list(EN_MONTHS), list(EN_ABBR), list(EN_DAYS)
        self.date = dict(EN_DATE)
        self.num = dict(EN_NUM)
        self.currency = "$"
        if el is None:
            return
        for e in el.iter():
            if e.tag == "monthNames":
                names = [m.text() for m in e.elements]
                if len(names) == 12:
                    if e.get("abbr") == "1":
                        self.abbr = names
                    else:
                        self.months = names
            elif e.tag == "dayNames" and e.get("abbr") != "1":
                names = [m.text() for m in e.elements]
                if len(names) == 7:
                    self.days = names
            elif e.tag == "datePattern" and e.get("name"):
                self.date[e.get("name")] = e.text()
            elif e.tag == "numberPattern" and e.get("name"):
                self.num[e.get("name")] = e.text()
            elif e.tag == "currencySymbol" and e.get("name") == "symbol":
                self.currency = e.text() or "$"
        self.date.setdefault("default", self.date.get("medium", EN_DATE["medium"]))
        if "numeric" in self.num:
            self.num.setdefault("decimal", self.num["numeric"])
            self.num.setdefault("integer", self.num["numeric"].split(".")[0])


def alternatives(pic: str, hint: str, loc: Locale) -> list[tuple[str, str]] | None:
    """書式を候補の並び [(種類, 中身)] にする。読めない形なら None。hint は値の種類（date・num・text）。"""
    out = []
    for alt in _split_bar(pic.strip()):
        alt = alt.strip()
        if not alt:
            continue
        m = re.fullmatch(r"(date|time|datetime|num|text|null|zero)(?:\.([a-z]+))?(?:\(([^)]*)\))?\{(.*)\}", alt, re.S)
        if m:
            cat, sub, body = m[1], m[2], m[4]
            if sub:
                table = loc.date if cat == "date" else loc.num if cat == "num" else None
                if table is None or sub not in table:
                    return None
                for b in _split_bar(table[sub]):
                    out.append((cat, b))
                continue
            if "{" in body:
                return None  # いくつもの書式をつないだもの
            out.append((cat, body))
        elif "{" in alt:
            return None
        else:
            out.append((hint, alt))
    return out


def _split_bar(s: str) -> list[str]:
    out, cur, q, depth = [], "", False, 0
    for ch in s:
        if ch == "'":
            q = not q
        elif not q and ch == "{":
            depth += 1
        elif not q and ch == "}":
            depth -= 1
        if ch == "|" and not q and depth == 0:
            out.append(cur)
            cur = ""
        else:
            cur += ch
    out.append(cur)
    return out


def _tokens(body: str, symbols: str) -> list[tuple[str, str]]:
    """(種類, 字)：sym は書式の記号の並び（同じ字の続き）、lit は字そのもの。"""
    out = []
    i = 0
    while i < len(body):
        ch = body[i]
        if ch == "'":
            j = i + 1
            lit = ""
            while j < len(body):
                if body[j] == "'":
                    if j + 1 < len(body) and body[j + 1] == "'":
                        lit += "'"
                        j += 2
                        continue
                    break
                lit += body[j]
                j += 1
            out.append(("lit", lit))
            i = j + 1
        elif ch in symbols:
            j = i
            while j < len(body) and body[j] == ch:
                j += 1
            out.append(("sym", body[i:j]))
            i = j
        else:
            out.append(("lit", ch))
            i += 1
    return out


# ====================================================================== 日付
def _date_fmt(d: _dt.date, body: str, loc: Locale) -> str | None:
    out = []
    for kind, s in _tokens(body, "DJMEYeGw"):
        if kind == "lit":
            out.append(s)
            continue
        if s == "D":
            out.append(str(d.day))
        elif s == "DD":
            out.append(f"{d.day:02d}")
        elif s == "M":
            out.append(str(d.month))
        elif s == "MM":
            out.append(f"{d.month:02d}")
        elif s == "MMM":
            out.append(loc.abbr[d.month - 1])
        elif s == "MMMM":
            out.append(loc.months[d.month - 1])
        elif s == "YY":
            out.append(f"{d.year % 100:02d}")
        elif s == "YYYY":
            out.append(f"{d.year:04d}")
        elif s in ("EEE", "EEEE"):
            n = loc.days[(d.weekday() + 1) % 7]
            out.append(n if s == "EEEE" else n[:3])
        elif s == "J":
            out.append(str(d.timetuple().tm_yday))
        elif s == "JJJ":
            out.append(f"{d.timetuple().tm_yday:03d}")
        else:
            return None
    return "".join(out)


def _date_parse(text: str, body: str, loc: Locale) -> _dt.date | None:
    rx, groups = "", []
    for kind, s in _tokens(body, "DJMEYeGw"):
        if kind == "lit":
            rx += re.escape(s)
        elif s in ("D", "M"):
            rx += r"(\d{1,2})"
            groups.append(s)
        elif s in ("DD", "MM", "YY"):
            rx += r"(\d{2})"
            groups.append(s)
        elif s == "YYYY":
            rx += r"(\d{4})"
            groups.append(s)
        elif s in ("MMM", "MMMM"):
            names = loc.abbr if s == "MMM" else loc.months
            rx += "(" + "|".join(re.escape(n) for n in names) + ")"
            groups.append(s)
        elif s in ("EEE", "EEEE"):
            rx += r"[^\d\s,]+"
        else:
            return None
    m = re.fullmatch(rx, text.strip(), re.I)
    if not m:
        return None
    y = mo = da = None
    for g, v in zip(groups, m.groups()):
        if g in ("D", "DD"):
            da = int(v)
        elif g in ("M", "MM"):
            mo = int(v)
        elif g == "YYYY":
            y = int(v)
        elif g == "YY":
            y = 2000 + int(v) if int(v) < 50 else 1900 + int(v)  # 仕様：世紀の区切りは既定で 50
        elif g in ("MMM", "MMMM"):
            names = [n.lower() for n in (loc.abbr if g == "MMM" else loc.months)]
            mo = names.index(v.lower()) + 1
    if None in (y, mo, da):
        return None
    try:
        return _dt.date(y, mo, da)
    except ValueError:
        return None


def canonical_date(text: str) -> _dt.date | None:
    """日付の決まった形（仕様 24章：YYYY-MM-DD・YYYYMMDD）。"""
    m = re.fullmatch(r"(\d{4})-?(\d{2})-?(\d{2})", text.strip())
    if not m:
        return None
    try:
        return _dt.date(int(m[1]), int(m[2]), int(m[3]))
    except ValueError:
        return None


# ====================================================================== 文字
_TEXT = {"9": str.isdigit, "A": str.isalpha, "X": lambda c: True, "O": str.isalnum, "0": str.isalnum, "t": lambda c: c == "\t"}


def _text_slots(body: str) -> list[tuple[str, str]]:
    out = []
    for kind, s in _tokens(body, "9AXO0t"):
        if kind == "lit":
            out.extend(("lit", c) for c in s)
        else:
            out.extend(("sym", c) for c in s)
    return out


def _text_fmt(v: str, body: str) -> str | None:
    slots = _text_slots(body)
    syms = [s for k, s in slots if k == "sym"]
    if len(v) != len(syms) or not all(_TEXT[s](c) for s, c in zip(syms, v)):
        return None
    it = iter(v)
    return "".join(next(it) if k == "sym" else s for k, s in slots)


def _text_parse(text: str, body: str) -> str | None:
    slots = _text_slots(body)
    if len(text) != len(slots):
        return None
    out = []
    for (k, s), c in zip(slots, text):
        if k == "lit":
            if c != s:
                return None
        elif not _TEXT[s](c):
            return None
        else:
            out.append(c)
    return "".join(out)


# ====================================================================== 数
_NUM_SYM = "9zZsSE$%(),."


def _num_parts(body: str):
    """数の書式から（整数部の桁の記号, 小数部の桁の記号, 区切りを使うか, 前後の字）を取り出す。読めなければ None。"""
    toks = _tokens(body, "9zZ")
    lead, frac, seen_dot, group = "", "", False, False
    for kind, s in toks:
        if kind == "sym":
            if seen_dot:
                frac += s
            else:
                lead += s
        elif s == "." and not seen_dot and (lead or frac or True):
            seen_dot = True
        elif s == "," and not seen_dot:
            group = True
        elif s in "Ee":
            return None
    return lead, frac, group


def _num_fmt(v: str, body: str, loc: Locale) -> str | None:
    p = _num_parts(body)
    if p is None:
        return None
    lead, frac, group = p
    try:
        x = float(v)
    except ValueError:
        return None
    neg = x < 0
    if neg and not re.search(r"[sS(-]", body):
        return None
    x = abs(x)
    s = f"{x:.{len(frac)}f}"
    ip, _, fp = s.partition(".")
    if frac and set(frac) <= {"z", "Z"}:
        fp = fp.rstrip("0")
    if len(ip.lstrip("0") or "0") > len(lead) and lead:
        return None
    need = sum(1 for c in lead if c == "9")
    ip = ip.lstrip("0")
    if len(ip) < need:
        ip = "0" * (need - len(ip)) + ip
    if not ip and not need and lead:
        ip = "0" if not fp else ""
    if group and ip:
        ip = f"{int(ip):,}"
    num = ip + ("." + fp if fp else "")
    pre = "".join(s for k, s in _tokens(body.split(".")[0], "9zZ,") if k == "lit" and s not in ",")
    post = "".join(s for k, s in _tokens(body.split(".", 1)[1], "9zZ,") if k == "lit") if "." in body else ""
    pre = pre.replace("$", loc.currency).replace("s", "-" if neg else "").replace("S", "-" if neg else " ")
    post = post.replace("s", "").replace("S", "")
    if neg and "(" in pre:
        return pre + num + post
    if "(" in pre and not neg:
        pre = pre.replace("(", "")
        post = post.replace(")", "")
    return ("-" if neg and "s" not in body and "S" not in body and "(" not in body else "") + pre + num + post


def _canon_num(t: str) -> str:
    """数の決まった形：整数部の前の 0 を除く（"01" → "1"）。"""
    t = t.lstrip("+")
    ip, dot, fp = t.partition(".")
    fp = fp.rstrip("0")
    ip = ip.lstrip("0") or "0"
    return ip + (dot + fp if fp else "")


def canonical_number(text: str) -> bool:
    return re.fullmatch(r"[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?", text.strip()) is not None


# ====================================================================== 入口
def format_value(value: str, pic: str, hint: str, loc: Locale) -> str | None:
    """決まった形の値を、書式で整えた字にする。整えられない・書式を読めないときは None。"""
    alts = alternatives(pic, hint, loc)
    if alts is None:
        return None
    for cat, body in alts:
        r = _one_format(value, cat, body, loc)
        if r is not None:
            return r
    return None


def _one_format(value: str, cat: str, body: str, loc: Locale) -> str | None:
    if cat == "null":
        return "" if value == "" else None
    if cat == "zero":
        return body if canonical_number(value) and float(value) == 0 and body else None
    if value == "":
        return None
    if cat == "date":
        d = canonical_date(value)
        return _date_fmt(d, body, loc) if d else None
    if cat == "text":
        return _text_fmt(value, body)
    if cat == "num":
        return _num_fmt(value, body, loc) if canonical_number(value) else None
    return None


def parse_value(text: str, pic: str, hint: str, loc: Locale) -> str | None:
    """書式どおりに書いた字を、決まった形の値に読む（日付は YYYY-MM-DD、文字は記号の所の字、数は区切りを除いた数）。"""
    alts = alternatives(pic, hint, loc)
    if alts is None:
        return None
    for cat, body in alts:
        if cat == "date":
            d = _date_parse(text, body, loc)
            if d:
                return d.isoformat()
        elif cat == "text":
            r = _text_parse(text, body)
            if r is not None:
                return r
        elif cat == "num":
            t = text.strip().replace(",", "").replace(loc.currency, "").replace("$", "")
            neg = t.startswith("(") and t.endswith(")")
            t = t.strip("()").strip()
            if canonical_number(t) and _num_fmt(("-" if neg else "") + t, body, loc) is not None:
                return ("-" if neg else "") + _canon_num(t)
        elif cat == "null" and text == "":
            return ""
    return None


def conforms(value: str, pic: str, hint: str, loc: Locale) -> bool | None:
    """値（決まった形）が書式に合うか（合えば True）。書式を読めないときは None。"""
    alts = alternatives(pic, hint, loc)
    if alts is None:
        return None
    return any(_one_format(value, cat, body, loc) is not None for cat, body in alts)
