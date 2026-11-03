"""文字のファイルの文字コード：仕様で決まる手がかりだけで決め、決まらないときは候補を示す（同じファイルなら、いつも同じ答え）。

決め方（上から順に。どの段も、規則として書けることだけを使う。ファイルの集まりに合わせて決めた数はない）：

1. BOM（UTF-8・UTF-32・UTF-16）：あれば、その文字コード。
2. NUL（0x00）のバイトがあれば、BOM のない UTF-32・UTF-16 を試す（ASCII の字が 0x00 と組になる形）。字の表のとおりに
   読めて「きれい」（下の定義）で、制御文字がなく、行の区切りのあるものが1つだけなら、それ。2つ以上なら、候補として示す。
   NUL があって、UTF-8 の多バイトの並びもない（ASCII の字と制御文字だけ）・どれでも読めないものは、文字のファイルではない
   （binary）。UTF-8 として多バイトの並びまで厳密に読めるものは、値の中の NUL は字（制御文字）として読む。
3. ASCII の字だけなら UTF-8（ファイルからは決まらない。ascii_only）。
4. UTF-8 として厳密に読めれば UTF-8（UTF-8 は、先頭のバイトと続くバイトの形が決まっていて、ほかの文字コードの字の並びが
   たまたま全体で合うことは、仕様の形からほぼない。HTML の仕様の決め方と同じ）。
5. 多バイトの文字コード（Shift_JIS（CP932）・EUC-JP・Big5（CP950・HKSCS）・GBK（CP936）・UHC（CP949））を、それぞれ厳密に
   試す。これらは、先頭のバイトと続くバイトの範囲が決まっているので、全体が合うことが手がかりになる。きれいに読めたものが
   1つなら、それ。2つ以上で、読んだ字が違うなら、**規則では決められない**：いちばん前のもの（todoku が主に読む日本の書類の
   文字コードを前に置いた、決まった順）で窓を作り、ほかを candidates に示す。どれが正しいかは、エージェントが字を読んで決める。
6. 1バイトの文字コード（Windows の 1250〜1258・874・KOI8・ISO-8859 の各部・Mac Roman）：ほぼすべてのバイトに字を当てるので、
   読めることは手がかりにならない。多バイトのどれでも読めないときだけ試し、きれいに読めたものが2つ以上（読んだ字が違う）なら、
   候補として示す（規則では決められない）。
   NUL のバイトがあるのに 1・2・4 の UTF として読めないものは、5・6 を試さない（UTF でない文字のファイルに NUL はない）。
7. どれでも読めなければ、**決められない**（encoding "latin-1"、basis "none"）：バイトを1字ずつ Latin-1 の字にして窓を作る
   （バイトの位置を保つため）。字が正しいとは限らないことを、窓と段階に必ず書く（黙って Latin-1 にしない）。

「きれい」：厳密に読めて（errors="strict"）、読んだ字に、割り当てのない字（Unicode の Cn）・サロゲート（Cs）・C1 の制御文字
（U+0080〜U+009F）がないこと。C0 の制御文字（NUL・SUB など）は、ASCII と同じバイトの文字コードではどれでも同じ字になるので、
文字コードを選ぶ手がかりにはならない（数だけを controls に数えて、窓に書く）。

書き込みはバイトの位置で行うので、字の位置（文字列の中の番号）とバイトの位置を結ぶ。Shift_JIS（CP932）には、
同じ字に2つのバイト列があるもの（NEC 選定 IBM 拡張文字と IBM 拡張文字など）があるので、ファイル全体を文字列にして
書き直すことはしない（書いた所だけを、そのファイルの文字コードで作ったバイト列に置き換える）。
"""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass, field

# BOM：UTF-32 の LE は UTF-16 の LE の BOM で始まるので、先に見る
BOMS = [(b"\xef\xbb\xbf", "utf-8"), (b"\xff\xfe\x00\x00", "utf-32-le"), (b"\x00\x00\xfe\xff", "utf-32-be"),
        (b"\xff\xfe", "utf-16-le"), (b"\xfe\xff", "utf-16-be")]
WIDE = ("utf-32-le", "utf-32-be", "utf-16-le", "utf-16-be")  # ASCII の字も1バイトでない文字コード
# 多バイトの文字コード（試す順。前のものほど、決められないときに窓を作るのに使う）
MULTI = ("cp932", "euc_jp", "cp950", "big5hkscs", "gbk", "cp949")
# 1バイトの文字コード
SINGLE = ("cp1252", "cp1250", "cp1251", "cp1253", "cp1254", "cp1255", "cp1256", "cp1257", "cp1258", "cp874",
          "koi8_r", "koi8_u", "iso8859_2", "iso8859_3", "iso8859_4", "iso8859_5", "iso8859_6", "iso8859_7", "iso8859_8",
          "iso8859_9", "iso8859_10", "iso8859_11", "iso8859_13", "iso8859_14", "iso8859_15", "iso8859_16", "mac_roman")
TRY = ("utf-8",) + MULTI  # 互換のため（多バイトまでの順）
NAME_JA = {"utf-8": "UTF-8", "cp932": "Shift_JIS（CP932）", "euc_jp": "EUC-JP", "cp950": "Big5（CP950）",
           "big5hkscs": "Big5-HKSCS", "gbk": "GBK（CP936）", "cp949": "EUC-KR／UHC（CP949）",
           "utf-16-le": "UTF-16（リトルエンディアン）", "utf-16-be": "UTF-16（ビッグエンディアン）",
           "utf-32-le": "UTF-32（リトルエンディアン）", "utf-32-be": "UTF-32（ビッグエンディアン）",
           "cp1252": "Windows-1252（西欧）", "cp1250": "Windows-1250（中欧）", "cp1251": "Windows-1251（キリル）",
           "cp1253": "Windows-1253（ギリシャ）", "cp1254": "Windows-1254（トルコ）", "cp1255": "Windows-1255（ヘブライ）",
           "cp1256": "Windows-1256（アラビア）", "cp1257": "Windows-1257（バルト）", "cp1258": "Windows-1258（ベトナム）",
           "cp874": "Windows-874（タイ）", "koi8_r": "KOI8-R", "koi8_u": "KOI8-U", "mac_roman": "Mac Roman",
           "latin-1": "Latin-1（文字コードを決められなかった）"}
for _k in range(2, 17):
    NAME_JA.setdefault(f"iso8859_{_k}", f"ISO-8859-{_k}")
BASIS_JA = {"bom": "BOM", "nul": "NUL のバイトの並び（BOM のない UTF-16・UTF-32）", "ascii": "ASCII の字だけ",
            "utf-8": "UTF-8 として厳密に読める", "multibyte": "多バイトの文字コードのうち、これだけが厳密に読める",
            "singlebyte": "1バイトの文字コードのうち、これだけがきれいに読める",
            "order": "規則では決められない（いくつもの文字コードで読める。決まった順の前のもので読んだ）",
            "none": "どの文字コードとしても読めない", "given": "指定された文字コード"}
WHITESPACE_CC = set("\t\n\v\f\r")


class EncodingError(ValueError):
    pass


def clean(text: str) -> bool:
    """読んだ字に、割り当てのない字（Cn）・サロゲート（Cs）・C1 の制御文字がないか（文字コードの選び方の「きれい」）。"""
    for ch in set(text):
        if "\x80" <= ch <= "\x9f":
            return False
        if ch > "\x7f" and unicodedata.category(ch) in ("Cn", "Cs"):
            return False
    return True


def controls(text: str) -> int:
    """C0 の制御文字（タブ・改行などの空白を除く。NUL・SUB など）と DEL の数。"""
    return sum(1 for ch in text if (ch < " " and ch not in WHITESPACE_CC) or ch == "\x7f")


@dataclass
class Decoded:
    encoding: str
    bom: bytes
    text: str
    ascii_only: bool  # ASCII の字だけ（文字コードは UTF-8 とみなしたが、ファイルからは決まらない）
    basis: str = ""  # どう決めたか（BASIS_JA の鍵）
    # 規則では決められないとき、ほかに読める文字コード：[{encoding, label, sample}]（sample は、読んだ字が違う最初の行）
    candidates: list = field(default_factory=list)
    controls: int = 0  # 制御文字（NUL など）の数

    @property
    def label(self) -> str:
        return NAME_JA.get(self.encoding, self.encoding) + ("・BOM あり" if self.bom else "・BOM なし")

    @property
    def binary(self) -> bool:
        """文字のファイルではない：NUL のバイトがあり、文字のファイルである手がかり（BOM・UTF-16/32 の形・UTF-8 の多バイトの
        並び）がない（どの文字コードとしても読めないか、ASCII の字と制御文字だけ）。POSIX の「文字のファイル」には NUL がない。"""
        return self.basis in ("none", "ascii") and "\x00" in self.text

    @property
    def undetermined(self) -> bool:
        """文字コードを規則で決められなかった（候補がいくつもある・どれでも読めない）。字は、エージェントが確かめる。"""
        return bool(self.candidates) or self.basis == "none"

    def info(self) -> dict:
        """窓の format に入れる、文字コードの決め方。"""
        out = {"encoding_basis": self.basis, "encoding_basis_ja": BASIS_JA.get(self.basis, self.basis)}
        if self.candidates:
            out["encoding_candidates"] = list(self.candidates)
        if self.controls:
            out["control_chars"] = self.controls
        return out

    def note_lines(self) -> list[str]:
        """Markdown の窓の注意（文字コードを決められないとき・制御文字があるとき）。"""
        out = []
        if self.basis == "none":
            out.append("- **文字コードを決められませんでした**：BOM がなく、UTF-8・UTF-16・多バイト（Shift_JIS・EUC-JP・Big5・"
                       "GBK・EUC-KR）・1バイトの文字コードのどれとしても、きれいに読めません。バイトを1字ずつ Latin-1 の字にして"
                       "出しています（字が正しいとは限りません。字化けしていれば、そう伝えてください）。")
        elif self.candidates:
            out.append(f"- **文字コードを規則では決められません**：{NAME_JA.get(self.encoding, self.encoding)} で読みましたが、"
                       "次の文字コードでも、仕様どおりにきれいに読めます。字の違う最初の行を並べます。意味の通る方を選んでください"
                       "（この窓の字が化けていれば、ほかの文字コードが正しい）。")
            out.append(f"  - {NAME_JA.get(self.encoding, self.encoding)}（この窓）：{self.sample_of(self.text, self.candidates[0])}")
            for c in self.candidates:
                out.append(f"  - {c['label']}：{c['sample']}")
        if self.controls:
            out.append(f"- 制御文字（NUL など、タブ・改行以外）が {self.controls:,} 個あります。字としてそのまま読みました"
                       "（窓では見えないことがあります）。")
        return out

    @staticmethod
    def sample_of(text: str, cand: dict) -> str:
        n = cand.get("line", 0)
        lines = text.splitlines()
        return _clip(lines[n]) if n < len(lines) else ""

    def byte_of(self, i: int) -> int:
        """文字列の i 番目の字の始まりの、ファイルの中のバイトの位置（字ごとのバイト数を足し上げた表を、1回だけ作る）。"""
        pre = getattr(self, "_prefix", None)
        if pre is None:
            sizes: dict[str, int] = {}
            pre = [0] * (len(self.text) + 1)
            at = 0
            for k, ch in enumerate(self.text):
                if ch < "\x80" and self.encoding not in WIDE:
                    at += 1
                else:
                    n = sizes.get(ch)
                    if n is None:
                        n = sizes[ch] = len(ch.encode(self.encoding))
                    at += n
                pre[k + 1] = at
            self._prefix = pre
        return len(self.bom) + pre[i]

    def span(self, blob: bytes, a: int, b: int) -> tuple[int, int]:
        """字の範囲 [a, b) → バイトの範囲。読み直して同じ字になることを確かめる。"""
        x, y = self.byte_of(a), self.byte_of(b)
        if blob[x:y].decode(self.encoding) != self.text[a:b]:
            raise EncodingError("字の位置とバイトの位置を結べませんでした")
        return x, y

    def encode(self, s: str) -> bytes:
        try:
            return s.encode(self.encoding)
        except UnicodeEncodeError as e:
            bad = s[e.start:e.end]
            raise EncodingError(f"「{bad}」は、このファイルの文字コード（{NAME_JA.get(self.encoding, self.encoding)}）では書けない字です") from None


def _clip(s: str, n: int = 120) -> str:
    """見本の行を、窓に出す長さにする（表示だけ。決め方には使わない）。"""
    s = s.strip()
    return s if len(s) <= n else s[:n] + "…"


def _strict(blob: bytes, enc: str) -> str | None:
    try:
        return blob.decode(enc)
    except UnicodeDecodeError:
        return None


def _first_diff_line(a: str, b: str) -> int:
    la, lb = a.splitlines(), b.splitlines()
    for k, (x, y) in enumerate(zip(la, lb)):
        if x != y:
            return k
    return min(len(la), len(lb))


def _choose(blob: bytes, encs: tuple, basis_one: str, text_of=None) -> Decoded | None:
    """encs のうち、きれいに読めるもの。1つならそれ（basis_one）、読んだ字の違うものが2つ以上なら、前のもので読み、
    ほかを候補に。どれも読めなければ None。"""
    got: list[tuple[str, str]] = []
    for enc in encs:
        t = _strict(blob, enc) if text_of is None else text_of(enc)
        if t is not None and clean(t) and all(t != u for _, u in got):
            got.append((enc, t))
    if not got:
        return None
    enc, text = got[0]
    cands = []
    for e2, t2 in got[1:]:
        k = _first_diff_line(text, t2)
        lines = t2.splitlines()
        cands.append({"encoding": e2, "label": NAME_JA.get(e2, e2), "line": k,
                      "sample": _clip(lines[k]) if k < len(lines) else ""})
    return Decoded(enc, b"", text, False, basis_one if not cands else "order", cands, controls(text))


def decode(blob: bytes, encoding: str | None = None) -> Decoded:
    """バイト列 → 字（決め方はこのモジュールの説明のとおり）。encoding を渡せば、その文字コードで厳密に読む
    （エージェントが候補から選んだとき）。"""
    if encoding:
        enc = encoding.lower().replace("_", "-") if encoding.lower().startswith("utf") else encoding.lower()
        bom = next((b for b, e in BOMS if e == enc and blob.startswith(b)), b"")
        text = blob[len(bom):].decode(enc, errors="strict")
        return Decoded(enc, bom, text, False, "given", [], controls(text))
    for bom, enc in BOMS:
        if blob.startswith(bom):
            text = _strict(blob[len(bom):], enc)
            if text is None and enc == "utf-32-le":
                continue  # FF FE 00 00 は、UTF-16 の LE の BOM と NUL のこともある
            if text is None:
                raise UnicodeDecodeError(enc, blob, len(bom), len(blob), "BOM の文字コードとして読めません")
            return Decoded(enc, bom, text, False, "bom", [], controls(text))
    if b"\x00" in blob:
        got = _choose(blob, WIDE, "nul")
        if got is not None and got.controls == 0 and ("\n" in got.text or "\r" in got.text):
            # NUL のバイトが字の上位（下位）のバイトとして説明でき、制御文字なく、行の区切りのある文字として読めた
            # （BOM のない UTF-16・UTF-32）。行の区切りがなければ、文字のファイルとしての手がかりがないので、選ばない
            return got
    if blob.isascii():
        text = blob.decode("ascii")
        return Decoded("utf-8", b"", text, True, "ascii", [], controls(text))
    text = _strict(blob, "utf-8")
    if text is not None:
        return Decoded("utf-8", b"", text, False, "utf-8", [], controls(text))
    if b"\x00" not in blob:
        # NUL のバイトは、UTF-8・UTF-16・UTF-32 でなければ、文字のファイルにはない（POSIX の「文字のファイル」の定義）。
        # NUL があるのに UTF の形で読めないものは、多バイト・1バイトの文字コードを試さない（バイナリか、壊れたファイル）
        for encs, basis in ((MULTI, "multibyte"), (SINGLE, "singlebyte")):
            got = _choose(blob, encs, basis)
            if got is not None:
                return got
    text = blob.decode("latin-1")
    return Decoded("latin-1", b"", text, False, "none", [], controls(text))
