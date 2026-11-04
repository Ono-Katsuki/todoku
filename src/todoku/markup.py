"""HTML・XML（.html・.htm・.xhtml・.xml）：字だけを読む（読むだけ。書かない）。外部のライブラリは使わない。

- 文字コード：.txt と同じ決め方（datafile/textenc.py：BOM・UTF・多バイト・1バイトの順に、仕様の形で決める）
- 字：タグの間の字（要素の中身）を、Python の html.parser で取り出す。タグ・属性・注釈（<!-- -->）・スクリプトと
  スタイル（<script>・<style>）の中身は読まない。文字参照（&amp; など）は字にする。外部の実体（DTD）は読まない
- 行：ブロックの要素（p・div・li・tr・見出し・br など）と、元の改行で区切り、行の中の空白の並びは1つにする。
  XML の要素の終わりも区切りにする（データの値が1行ずつ並ぶ）
- 見た目（表の形・字の大きさ・リンクの先）は読まないので、段階は読む B（「HTML/XML は字だけを読む」）。
  書き込みはしない（タグのある字への書き込みを証明する方法を作っていないため）
"""

from __future__ import annotations

import re
from html.parser import HTMLParser

KIND_JA = "HTML・XML（字だけを読む）"
EXTENSIONS = (".html", ".htm", ".xhtml", ".xml")
_SKIP = {"script", "style", "template", "noscript", "head"}
_BLOCK = {"address", "article", "aside", "blockquote", "br", "caption", "dd", "div", "dl", "dt", "fieldset", "figcaption",
          "figure", "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "li", "main", "nav", "ol", "p",
          "pre", "section", "table", "tbody", "td", "tfoot", "th", "thead", "title", "tr", "ul"}
NOTE = ("HTML/XML は字だけを読みます：タグの間の字を取り出し、タグ・属性・注釈・スクリプトとスタイルの中身・見た目（表の形・"
        "字の大きさ・リンクの先）は読みません")


class MarkupError(ValueError):
    pass


class _Text(HTMLParser):
    def __init__(self, xml: bool):
        super().__init__(convert_charrefs=True)
        self.xml = xml
        self.lines: list[str] = [""]
        self.skip = 0
        self.root: str | None = None

    def _break(self) -> None:
        if self.lines[-1].strip():
            self.lines.append("")

    def handle_starttag(self, tag, attrs):
        if self.root is None:
            self.root = tag
        if not self.xml and tag in _SKIP:
            self.skip += 1
        elif self.xml or tag in _BLOCK:
            self._break()

    def handle_startendtag(self, tag, attrs):
        if self.root is None:
            self.root = tag
        if self.xml or tag in _BLOCK:
            self._break()

    def handle_endtag(self, tag):
        if not self.xml and tag in _SKIP:
            self.skip = max(0, self.skip - 1)
        elif self.xml or tag in _BLOCK:
            self._break()

    def handle_data(self, data):
        if self.skip:
            return
        parts = re.split(r"\r\n|\r|\n", data)
        for k, part in enumerate(parts):
            if k:
                self._break()
            self.lines[-1] += part

    def unknown_decl(self, data):  # <![CDATA[ … ]]>：中身は字
        if data.upper().startswith("CDATA[") and not self.skip:
            self.handle_data(data[6:])


def build(blob: bytes, name: str, encoding: str | None = None) -> dict:
    """窓の JSON：{kind, format, lines, fields（いつも空：書く所は示さない）, note}。"""
    from .datafile.textenc import decode

    try:
        dec = decode(blob, encoding)
    except UnicodeDecodeError as e:
        raise MarkupError(f"文字コードを読めません（{e.reason}）") from None
    if dec.binary:
        raise MarkupError("文字のファイルではありません（どの文字コードとしても読めず、NUL のバイトがあります）")
    text = dec.text
    head = text.lstrip("﻿ \t\r\n")[:200].lower()
    xml = name.lower().endswith(".xml") or (head.startswith("<?xml") and "<html" not in text[:2000].lower())
    p = _Text(xml)
    try:
        p.feed(text)
        p.close()
    except Exception as e:  # noqa: BLE001  html.parser は寛容だが、読めなかったことは黙らずに伝える
        raise MarkupError(f"{'XML' if xml else 'HTML'} として字を取り出せません（{type(e).__name__}）") from None
    lines = [re.sub(r"[ \t　 ]+", " ", ln).strip() for ln in p.lines]
    lines = [ln for ln in lines if ln]
    return {"kind": "markup", "file": name, "format": {"markup": "XML" if xml else "HTML", "encoding": dec.encoding,
                                                       "encoding_ja": dec.label, "root": p.root, "lines": len(lines),
                                                       **dec.info(), "encoding_notes": dec.note_lines()},
            "lines": lines, "fields": [], "note": NOTE}


def markdown_parts(win: dict) -> list[tuple[str, str]]:
    f = win["format"]
    out = [f"# {win['file']}（{f['markup']}・{f['encoding_ja']}）", "", f"> {win['note']}。書き込みはしません。", ""]
    for n in f.get("encoding_notes") or []:
        out.append(f"> {n}")
    out += ["", "## 字（タグの間の字。上から）", ""]
    out += [f"{k}｜{ln}" for k, ln in enumerate(win["lines"], 1)]
    return [("本文", "\n".join(out) + "\n")]


def grade(blob: bytes, name: str, encoding: str | None = None):
    from .grade import Reason, _final, _locked

    try:
        win = build(blob, name, encoding)
    except MarkupError as e:
        return _locked("読めない", str(e), "読めないので、書けません")
    reasons = [Reason("B", "HTML/XML は字だけを読む", f"{NOTE}。字は {win['format']['lines']} 行です")]
    write =[Reason("D", "HTML/XML には書かない", "とどくは HTML・XML に書き込みません（タグのある字への書き込みを証明する方法を"
                    "作っていません）。書く必要があれば、本人に、Word・Excel・PDF など書ける形式の様式があるかを確かめてください")]
    return _final(reasons, write)


def audit(blob: bytes, name: str):
    """検査：窓とは別の読み方（正規表現でタグ・注釈・スクリプトを除いた字）で取り出した字の並びが、窓の字の並びと同じか
    （空白は比べない）。書き込みはしないので、書く検査はない。"""
    from .xlsx.audit import Audit, Check

    au = Audit(name)
    try:
        win = build(blob, name)
    except MarkupError as e:
        au.checks.append(Check("読む", False, str(e)))
        return au
    from .datafile.textenc import decode

    text = decode(blob).text
    if win["format"]["markup"] == "HTML":
        text = re.sub(r"(?is)<(script|style|template|noscript|head)\b.*?</\1\s*>", " ", text)
    text = re.sub(r"(?s)<!--.*?-->", " ", text)
    text = re.sub(r"(?s)<!\[CDATA\[(.*?)\]\]>", r"\1", text)
    text = re.sub(r"(?s)<[^>]*>", " ", text)
    import html as _html

    other = re.sub(r"\s+", "", _html.unescape(text))
    mine = re.sub(r"\s+", "", "".join(win["lines"]))
    ok = other == mine
    au.checks.append(Check("読み落とし（字）", ok, "窓の字が、別の読み方（タグを除いた字）と一致しました" if ok else
                           f"窓の字（{len(mine)} 字）が、別の読み方（{len(other)} 字）と違います"))
    au.checks.append(Check("書き込み", True, "HTML・XML には書き込みません（読むだけの形式）"))
    return au
