"""文字だけのファイル（.txt）を読む。外部のライブラリは使わない。

- 文字コード：CSV と同じ決め方（datafile/textenc.py：BOM・UTF・多バイト・1バイトの順に、仕様の形で決める。決められないときは
  候補を窓に示し、どれでも読めなければ「決められない」と書いて Latin-1 の字で出す）
- 行：改行（CRLF・LF・CR）で区切る。行の終わりの改行は行ごとに覚える（書いても、改行の字は変えない）
- 読んだ結果は、Word（.docx）と同じ形（1行を1段落）にして、窓・入力欄の推定（「令和　年　月　日」・「？」の所）・
  字を目印にした書き込みを共通にする。入力欄の id の「P12.1」は、12行目の1つ目の空欄
- 字の位置（文字列の中の番号）とファイルのバイトの位置は、textenc の表で結ぶ（書くときは、書いた所だけをバイトで置き換える）
"""

from __future__ import annotations

from dataclasses import dataclass

from ..datafile.textenc import Decoded, decode
from ..docx.document import Para, Seg

NEWLINE_JA = {"\r\n": "CRLF", "\n": "LF", "\r": "CR", "": "なし"}


class TxtError(ValueError):
    pass


@dataclass
class Line(Para):
    """1行（Word の段落と同じ形）。a・b は行の中身の字の範囲（ファイル全体の文字列の中の番号。改行は含まない）。"""

    a: int = 0
    b: int = 0
    terminator: str = ""  # 行の終わりの改行（最後の行で改行がなければ ""）
    number: int = 0  # 何行目か（1 から）


def split_lines(text: str) -> list[tuple[int, int, str]]:
    """(行の中身の始め, 終わり, 改行)。改行は CRLF・LF・CR。最後の改行の後ろに字がなければ、そこに行は作らない。"""
    out: list[tuple[int, int, str]] = []
    i, n = 0, len(text)
    while i < n:
        j = i
        while j < n and text[j] not in "\r\n":
            j += 1
        term = "\r\n" if text.startswith("\r\n", j) else (text[j] if j < n else "")
        out.append((i, j, term))
        i = j + len(term)
    return out


def _segs(text: str, a: int) -> list[Seg]:
    """行の字を、タブとそれ以外の並びに分ける（Word の w:t・w:tab と同じ分け方）。node は、その並びの初めの字の番号。"""
    segs: list[Seg] = []
    k = 0
    while k < len(text):
        tab = text[k] == "\t"
        m = k
        while m < len(text) and (text[m] == "\t") == tab:
            m += 1
        segs.append(Seg(text[k:m], "tab" if tab else "t", a + k))
        k = m
    return segs


class TxtDocument:
    """文字だけのファイル。Word の文書と同じ問い合わせ（all_blocks・controls・marks など）に答える。"""

    kind = "txt"

    def __init__(self, blob: bytes, filename: str = "", encoding: str | None = None):
        self.blob = blob
        self.filename = filename
        try:
            self.dec: Decoded = decode(blob, encoding)
        except UnicodeDecodeError as e:
            raise TxtError(f"文字コードを読めません（{e.reason}）") from None
        self.text = self.dec.text
        if self.dec.binary:
            raise TxtError("文字のファイルではありません（どの文字コードとしても読めず、NUL のバイトがあります）")
        self.lines: list[Line] = []
        for k, (a, b, term) in enumerate(split_lines(self.text), 1):
            self.lines.append(Line(None, "txt", _segs(self.text[a:b], a), a=a, b=b, terminator=term, number=k))
        self.body = list(self.lines)
        terms = [ln.terminator for ln in self.lines if ln.terminator]
        self.newline = max(("\r\n", "\n", "\r"), key=lambda t: (terms.count(t), t == "\r\n")) if terms else ""
        self.mixed_newlines = len(set(terms)) > 1
        self.ends_with_newline = bool(self.lines) and bool(self.lines[-1].terminator)
        # Word の文書と同じ問い合わせに答えるための、ないもの
        self.controls: list = []
        self.protection = ""
        self.protection_enforced = False
        self.has_macros = False
        self.vba = None
        self.pictures: list = []
        self.unread: list[str] = []

    def all_blocks(self):
        yield "body", "txt", self.body

    def marks(self) -> dict:
        """書き込んだ欄の印：文字だけのファイルには、見えない印を置けない（書いた所は、証明で読み直して確かめる）。"""
        return {}

    def format_info(self) -> dict:
        return {"encoding": self.dec.encoding, "encoding_ja": self.dec.label, "bom": bool(self.dec.bom),
                "ascii_only": self.dec.ascii_only, "newline": NEWLINE_JA[self.newline], "mixed_newlines": self.mixed_newlines,
                "ends_with_newline": self.ends_with_newline, "lines": len(self.lines),
                **self.dec.info(), "encoding_notes": self.dec.note_lines()}

    def line_bytes(self, ln: Line) -> tuple[int, int, int]:
        """(行の中身の始め, 終わり, 改行の終わり) のバイトの位置。"""
        x, y = self.dec.span(self.blob, ln.a, ln.b)
        z = self.dec.span(self.blob, ln.b, ln.b + len(ln.terminator))[1]
        return x, y, z
