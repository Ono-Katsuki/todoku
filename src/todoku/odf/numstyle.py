"""OpenDocument の数の書式（number:*-style、ODF 1.3 第3部 16.29）を、Excel の表示形式の書き方にする。

使い道は2つ：(1) 書式の種類（日付・時刻・割合・通貨・文字・数）で、書く値の受け取り方を決める、(2) 書いた値の見える字
（text:p）を作る。(2) は、同じ書式のセルに保存されている見える字と照らして、同じ字になる書式だけを使う
（照らせない・合わない書式は、書いた値の見える字を作らずに、そのことを知らせる）。
"""

from __future__ import annotations

from .xml import El, q

_N = "{urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0}"


def _lit(s: str) -> str:
    """Excel の表示形式の中の、そのままの字（"…" で囲む。引用符は使えないので \\ で）。"""
    if not s:
        return ""
    if s in ("-", "/", ":", ".", ",", " ", "(", ")", "+", "$"):
        return s if s not in (".", ",") else '"' + s + '"'
    return '"' + s.replace('"', '\\"') + '"'


def _text(el: El, data: bytes) -> str:
    from .package import text_of

    return text_of(el, data)


def _number(el: El, percent: bool = False) -> str:
    dp = int(el.get("number:decimal-places", "0") or 0)
    mind = el.get("number:min-decimal-places")
    mind = int(mind) if mind is not None and mind.isdigit() else (
        int(el.get("{urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0}min-decimal-places", dp) or dp))
    mi = int(el.get("number:min-integer-digits", "1") or 1)
    grouping = el.get("number:grouping") == "true"
    if grouping:
        intpart = "#,##0" if mi == 1 else "#,###" if mi == 0 else "#,##" + "0" * mi
    else:
        intpart = "0" * mi if mi > 0 else "#"
    frac = ""
    if dp:
        frac = "." + "0" * min(mind, dp) + "#" * (dp - min(mind, dp))
    return intpart + frac


_CALENDARS = {"gengou", "ROC", "buddhist", "hijri"}


def _locale_prefix(el: El, cal: str) -> str | None:
    """日付・時刻の書式の言語（number:language・number:country、ODF 1.3 第3部 19.351・19.342）と暦を、Excel の表示形式の
    [$-…]（BCP 47 の名前）・B2 にする。言語の指定がなければ何も付けない（文書の既定の言語）。"""
    lang = (el.get("number:language") or "").strip()
    country = (el.get("number:country") or "").strip()
    tag = f"{lang}-{country}" if lang and country else lang
    if cal == "ROC":
        if tag and tag.lower() not in ("zh-tw", "zh-hant"):
            return None  # 民国の暦を、台湾の中国語でない言語で書く書式：Excel の表示形式で同じ字になるかを確かめていない
        return "[$-zh-TW]"
    if cal == "hijri":
        return ("[$-" + tag + "]" if tag else "") + "B2"
    return "[$-" + tag + "]" if tag else ""


def section(el: El, data: bytes) -> tuple[str, str]:
    """1つの書式（style:map を除く）→ (Excel の書式の1節, 種類)。"""
    loc = el.local
    kind = {"number-style": "number", "currency-style": "currency", "percentage-style": "percentage",
            "date-style": "date", "time-style": "time", "boolean-style": "boolean", "text-style": "text"}.get(loc, "number")
    out = []
    color = None
    # 暦（number:calendar、ODF 1.3 第3部 19.341：gregorian・gengou・ROC・hanja_yoil・hanja・hijri・jewish・buddhist など）。
    # 年月日の要素ごとに書けるが、1つの書式に違う暦が混じるものは作らない。Excel の表示形式で同じ字になる暦だけを使う：
    # gengou → 元号（g・e）、ROC → 民国の年（[$-zh-TW] の e）、buddhist → 仏暦（b）、hijri → ヒジュラ暦（B2）。ほかは作らない
    # （グレゴリオ暦の年を黙って出さない）
    cals = {ch.get("number:calendar") for ch in el.els() if ch.tag.startswith(_N) and ch.get("number:calendar")}
    cals.discard("gregorian")
    if len(cals) > 1 or (cals and next(iter(cals)) not in _CALENDARS):
        return "", kind
    cal = next(iter(cals), "")
    for ch in el.els():
        t = ch.tag
        if t == q("style:text-properties"):
            color = ch.get("fo:color")
            continue
        if not t.startswith(_N):
            continue
        n = ch.local
        long_ = ch.get("number:style") == "long"
        gengou = cal == "gengou"
        if n == "number":
            out.append(_number(ch, kind == "percentage"))
        elif n == "scientific-number":
            dp = int(ch.get("number:decimal-places", "0") or 0)
            ex = int(ch.get("number:min-exponent-digits", "2") or 2)
            out.append("0" + ("." + "0" * dp if dp else "") + "E+" + "0" * ex)
        elif n == "fraction":
            dn = int(ch.get("number:min-denominator-digits", "1") or 1)
            nn = int(ch.get("number:min-numerator-digits", "1") or 1)
            mi = int(ch.get("number:min-integer-digits", "0") or 0)
            dv = ch.get("number:denominator-value")
            out.append(("#" if mi == 0 else "0" * mi) + " " + "?" * nn + "/" + (dv if dv else "?" * dn))
        elif n == "text":
            out.append(_lit(_text(ch, data)))
        elif n == "currency-symbol":
            out.append(_lit(_text(ch, data)))
        elif n == "text-content":
            out.append("@")
        elif n == "boolean":
            out.append("TRUE/FALSE")
        elif n == "year":
            if gengou or cal == "ROC":
                out.append("e")
            elif cal == "buddhist":
                out.append("bbbb" if long_ else "bb")
            else:
                out.append("yyyy" if long_ else "yy")
        elif n == "era":
            if cal not in ("", "gengou"):
                return "", kind  # 民国・仏暦・ヒジュラ暦の元号の名前は、Excel の表示形式で同じ字にできるかを確かめていない
            out.append("ggg" if long_ else "g")
        elif n == "month":
            if ch.get("number:textual") == "true":
                out.append("mmmm" if long_ else "mmm")
            else:
                out.append("mm" if long_ else "m")
        elif n == "day":
            out.append("dd" if long_ else "d")
        elif n == "day-of-week":
            out.append("aaaa" if long_ else "aaa")
        elif n == "hours":
            trunc = el.get("number:truncate-on-overflow") == "false"
            h = "hh" if long_ else "h"
            out.append(f"[{h}]" if trunc else h)
        elif n == "minutes":
            out.append("mm" if long_ else "m")
        elif n == "seconds":
            dp = int(ch.get("number:decimal-places", "0") or 0)
            out.append(("ss" if long_ else "s") + ("." + "0" * dp if dp else ""))
        elif n == "am-pm":
            out.append("AM/PM")
        elif n in ("week-of-year", "quarter"):
            return "", kind  # Excel の書き方にない部分：作れない
        elif n == "embedded-text":
            return "", kind
    fmt = "".join(out)
    if fmt and kind in ("date", "time"):
        pre = _locale_prefix(el, cal)
        if pre is None:
            return "", kind
        fmt = pre + fmt
    if kind == "percentage" and "%" not in fmt:
        fmt += "%"
    if color and fmt:
        name = {"#ff0000": "[Red]", "#0000ff": "[Blue]", "#000000": "[Black]"}.get(color.lower())
        if name:
            fmt = name + fmt
    return fmt, kind


def excel_format(el: El | None, data: bytes, lookup) -> tuple[str, str]:
    """number:*-style → (Excel の表示形式, 種類)。style:map（正の数・負の数など）は、Excel の書式の節にする。
    作れない書式は ("", 種類)。lookup(名前) は、ほかの数の書式（style:apply-style-name）を引く。"""
    if el is None:
        return "General", "number"
    main, kind = section(el, data)
    maps = [m for m in el.els() if m.tag == q("style:map")]
    if not maps:
        return (main or ""), kind
    secs = {}
    for m in maps:
        cond = (m.get("style:condition") or "").replace(" ", "")
        target = lookup(m.get("style:apply-style-name"))
        if target is None:
            return "", kind
        fmt, _k = section(target, data)
        if not fmt:
            return "", kind
        secs[cond] = fmt
    if set(secs) == {"value()>=0"}:
        return secs["value()>=0"] + ";" + main, kind
    if set(secs) == {"value()>0", "value()<0"}:
        return secs["value()>0"] + ";" + secs["value()<0"] + ";" + main, kind
    if set(secs) == {"value()>0"}:
        return secs["value()>0"] + ";" + main, kind
    return "", kind
