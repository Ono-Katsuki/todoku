"""文字のファイルの文字コード：仕様で決まる手がかりだけで決め、決まらないときは候補を示す（同じファイルなら、いつも同じ答え）。

決め方（上から順に。どの段も、規則として書けることだけを使う。ファイルの集まりに合わせて決めた数はない）：

1. BOM（UTF-8・UTF-32・UTF-16）：あれば、その文字コード。
2. NUL（0x00）のバイトがあれば、BOM のない UTF-32・UTF-16 を試す（ASCII の字が 0x00 と組になる形）。字の表のとおりに
   読めて「きれい」（下の定義）で、制御文字がなく、行の区切りのあるものが1つだけなら、それ。2つ以上なら、候補として示す。
   NUL があって、UTF-8 の多バイトの並びもない（ASCII の字と制御文字だけ）・どれでも読めないものは、文字のファイルではない
   （binary）。UTF-8 として多バイトの並びまで厳密に読めるものは、値の中の NUL は字（制御文字）として読む。
3. ASCII の字だけなら UTF-8（ファイルからは決まらない。ascii_only）。
3'. ASCII のバイトだけで、ESC（0x1B）があるとき：ISO/IEC 2022 の切り替えの印（ESC $ B など。文字集合を、ファイルの中で指す
   印）かもしれない。ISO-2022-JP（-1・-2・半角カナつき・-3・-2004）と ISO-2022-KR を、それぞれ厳密に試す。読めて、きれいで、
   ESC の数がバイトより減り（切り替えの印として読めた ESC が 1 つはある）、SO・SI（0x0E・0x0F：ISO/IEC 2022 の、もう
   1 つの切り替えの印）が字として残らないものが、その読み方。読んだ字が同じなら、順の前の名前（basis "escape"）。
   読んだ字の違うものが 2 つ以上なら、候補として示す。どれでも読めないときだけ、Windows の ISO-2022-JP（CP50220・
   CP50221・CP50222：2 バイトの字を CP932 の表で読む。NEC 特殊文字（13 区の ①・㈱ など）・NEC 選定 IBM 拡張文字（89〜92 区
   の 髙・﨑 など）・SO/SI の半角カナ。iso2022ms.py）を、同じ決まりで試す。それでも読めなければ、3 のまま（ESC は
   制御文字として数える）。
   切り替えの印でない ESC（端末の色の印など）は、制御文字として残り、数える。HZ・UTF-7 は、ふつうの ASCII の字
   （~{ や +）が印なので、規則では決めない（エージェントが encoding で選ぶ）。
4. UTF-8 として厳密に読めれば UTF-8（UTF-8 は、先頭のバイトと続くバイトの形が決まっていて、ほかの文字コードの字の並びが
   たまたま全体で合うことは、仕様の形からほぼない。HTML の仕様の決め方と同じ）。
4'. ファイルの中の、文字コードの宣言：XML の宣言（ファイルの頭の <?xml … encoding="…"?>。XML 1.0 4.3.3）、HTML の meta
   （頭の 1024 バイトの中の <meta charset> か <meta http-equiv="Content-Type" content="…; charset=…">。HTML の仕様の
   prescan と同じ範囲）、RTF の \ansicpg（頭の {\rtf1 の後の、ANSI のコードページ）。宣言があり、その文字コードで厳密に
   きれいに読めれば、それ（basis "declared"）。宣言と中身が合わない（読めない）ときは、宣言を使わず、下の段で決める。
5. BOM も宣言もなく、UTF-8 でもない 8 ビットの字のファイルは、**バイトからは文字コードが決まらない**。読める文字コードを
   すべて試し、きれいに読めたものが 1 つだけなら、それ。読んだ字の違うものが 2 つ以上なら、決まった順のいちばん前のもので
   窓を作り、ほかを candidates に示す。どれが正しいかは、エージェントが字を読んで決める（encoding で渡す）。
   - 多バイトの文字コード（Shift_JIS（CP932）・EUC-JP・Big5（CP950・HKSCS）・GBK（CP936）・UHC（CP949））を、それぞれ厳密に
     試す。先頭のバイトと続くバイトの範囲が決まっているので、合わないものは落ちる。
   - 1バイトの文字コード（Windows の 1250〜1258・874・KOI8・ISO-8859 の各部・Mac Roman）も、いつも試す。ほぼすべてのバイトに
     字を当てるので、多バイトで読めたファイルも 1 バイトとして読める（ドイツ語の Windows-1252 の「Müller」は GBK の
     「M黮ler」としても、タイ語の Windows-874 は Shift_JIS の半角カナとしても、厳密に読める）。多バイトで読めたことだけでは
     決めない。
   - 順は、多バイト（上の順）・1バイト（上の順）。順は決め方ではない（決められないときに、どれで窓を作るかだけ）。
6. 1バイトの文字コードの読み方は、ファイルにある 0x80 以上のバイトの種類だけで決まる（1 バイトが 1 字）。バイトの種類を
   1 回だけ集め、文字コードごとの表で、きれいに読めるか・読み方が同じかを見る（ファイルを文字コードの数だけ読み直さない）。
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

from array import array

import codecs
import functools
import re
import unicodedata
from dataclasses import dataclass, field

from . import iso2022ms  # Windows の ISO-2022-JP（CP50220・CP50221・CP50222）の codec を登録する

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
           "iso2022_jp": "ISO-2022-JP", "iso2022_jp_1": "ISO-2022-JP-1", "iso2022_jp_2": "ISO-2022-JP-2",
           "iso2022_jp_ext": "ISO-2022-JP（半角カナつき）", "iso2022_jp_3": "ISO-2022-JP-3",
           "iso2022_jp_2004": "ISO-2022-JP-2004", "iso2022_kr": "ISO-2022-KR", "hz": "HZ-GB-2312", "utf-7": "UTF-7",
           "iso2022_jp_ms": "ISO-2022-JP（Windows の CP50221：NEC 特殊文字・IBM 拡張文字つき）",
           "latin-1": "Latin-1（文字コードを決められなかった）"}
# ISO/IEC 2022 の切り替えの印で文字集合を指す文字コード（試す順。字の少ないものが前）
ISO2022 = ("iso2022_jp", "iso2022_jp_1", "iso2022_jp_2", "iso2022_jp_ext", "iso2022_jp_3", "iso2022_jp_2004", "iso2022_kr")
# 上のどれでも読めないときだけ試す：Windows の ISO-2022-JP（2 バイトの字を CP932 の表で読む）
ISO2022_MS = iso2022ms.NAME
for _k in range(2, 17):
    NAME_JA.setdefault(f"iso8859_{_k}", f"ISO-8859-{_k}")
BASIS_JA = {"bom": "BOM", "nul": "NUL のバイトの並び（BOM のない UTF-16・UTF-32）", "ascii": "ASCII の字だけ",
            "utf-8": "UTF-8 として厳密に読める", "multibyte": "多バイトの文字コードのうち、これだけが厳密に読める（1バイトの文字コードでは、きれいに読めない）",
            "singlebyte": "1バイトの文字コードのうち、これだけがきれいに読める（多バイトの文字コードでは読めない）",
            "order": "規則では決められない（いくつもの文字コードで読める。決まった順の前のもので読んだ）",
            "none": "どの文字コードとしても読めない",
            "declared": "ファイルの中の宣言（XML の宣言・HTML の meta charset・RTF の \\ansicpg）",
            "escape": "ISO/IEC 2022 の切り替えの印（ESC で始まる、文字集合を指す印）のとおりに、厳密に読める",
            "given": "エージェントが選んで指定した文字コード（encoding）"}
# 文字コードの名前（宣言・エージェントの指定）→ このモジュールの名前。WHATWG Encoding の対応に合わせる
# （Shift_JIS は Windows-31J（CP932）、EUC-KR は Windows-949、GB2312 は GBK、ISO-8859-1 は Windows-1252 として読む）
ALIASES = {"shift_jis": "cp932", "shift-jis": "cp932", "sjis": "cp932", "x-sjis": "cp932", "windows-31j": "cp932",
           "ms932": "cp932", "ms_kanji": "cp932", "csshiftjis": "cp932", "cp932": "cp932",
           "euc-jp": "euc_jp", "eucjp": "euc_jp", "x-euc-jp": "euc_jp", "euc_jp": "euc_jp",
           "gb2312": "gbk", "gbk": "gbk", "x-gbk": "gbk", "cp936": "gbk", "windows-936": "gbk", "gb18030": "gb18030",
           "big5": "cp950", "cp950": "cp950", "big5-hkscs": "big5hkscs", "big5hkscs": "big5hkscs",
           "euc-kr": "cp949", "ks_c_5601-1987": "cp949", "windows-949": "cp949", "cp949": "cp949", "uhc": "cp949",
           "iso-8859-1": "cp1252", "latin1": "cp1252", "us-ascii": "cp1252", "ascii": "utf-8", "utf8": "utf-8",
           "utf-8": "utf-8", "utf-8-sig": "utf-8", "utf-16le": "utf-16-le", "utf-16be": "utf-16-be",
           "utf-32le": "utf-32-le", "utf-32be": "utf-32-be", "mac": "mac_roman", "macintosh": "mac_roman"}
for _k in range(1250, 1259):
    ALIASES[f"windows-{_k}"] = ALIASES[f"cp{_k}"] = f"cp{_k}"
ALIASES["windows-874"] = ALIASES["cp874"] = "cp874"
for _k in (50220, 50221, 50222):  # Windows のコードページの名前（Python に、この名前の codec はない）
    ALIASES[f"cp{_k}"] = ALIASES[f"windows-{_k}"] = ALIASES[f"ms{_k}"] = ISO2022_MS
ALIASES["iso-2022-jp-ms"] = ALIASES[ISO2022_MS] = ISO2022_MS
ALIASES.update({v.lower(): k for k, v in NAME_JA.items() if k != "latin-1"})  # 日本語の名前（窓の候補の label）でも指せる
WHITESPACE_CC = set("\t\n\v\f\r")


class EncodingError(ValueError):
    pass


_CLEAN: set[str] = set()  # きれいと確かめた字（同じ字を、文字コードの候補ごとに確かめ直さない）


def clean(text: str) -> bool:
    """読んだ字に、割り当てのない字（Cn）・サロゲート（Cs）・C1 の制御文字がないか（文字コードの選び方の「きれい」）。"""
    if text.isascii():
        return True
    new = set(text).difference(_CLEAN)
    for ch in new:
        if "\x80" <= ch <= "\x9f":
            return False
        if ch > "\x7f" and unicodedata.category(ch) in ("Cn", "Cs"):
            return False
    if len(_CLEAN) < 200_000:
        _CLEAN.update(new)
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
    # 規則では決められないとき、ほかに読める文字コード：[{encoding, label, line, sample}]（sample は、読んだ字が違う最初の所。
    # line は、その行の番号。own_sample は、同じ所を窓の文字コードで読んだ字：窓の注意にだけ使い、結果には入れない）
    candidates: list = field(default_factory=list)
    controls: int = 0  # 制御文字（NUL など）の数
    declared_by: str = ""  # 宣言で決めたとき、どこの宣言か（XML の宣言・HTML の meta charset・RTF の \\ansicpg）

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
        if self.basis == "given":
            out["encoding_chosen_by"] = "agent"  # エージェントが候補から選んで指定した（規則で決めたのではない）
        if self.declared_by:
            out["encoding_declared_by"] = self.declared_by
        if self.candidates:
            out["encoding_candidates"] = [{k: v for k, v in c.items() if k != "own_sample"} for c in self.candidates]
        if self.controls:
            out["control_chars"] = self.controls
        return out

    def note_lines(self) -> list[str]:
        """Markdown の窓の注意（文字コードを決められないとき・制御文字があるとき）。"""
        out = []
        if self.basis == "given":
            out.append(f"- 文字コードは、エージェントが選んで指定した {NAME_JA.get(self.encoding, self.encoding)} で読みました"
                       "（規則では決められなかったもの。このファイルを厳密に読めることは確かめました）。")
        if self.basis == "none":
            out.append("- **文字コードを決められませんでした**：BOM がなく、UTF-8・UTF-16・多バイト（Shift_JIS・EUC-JP・Big5・"
                       "GBK・EUC-KR）・1バイトの文字コードのどれとしても、きれいに読めません。バイトを1字ずつ Latin-1 の字にして"
                       "出しています（字が正しいとは限りません。字化けしていれば、そう伝えてください）。")
        elif self.candidates:
            out.append(f"- **文字コードを規則では決められません**：{NAME_JA.get(self.encoding, self.encoding)} で読みましたが、"
                       "次の文字コードでも、仕様どおりにきれいに読めます。字の違う最初の所を並べます。意味の通る方を選んでください"
                       "（この窓の字が化けていれば、ほかの文字コードが正しい）。選んだら、その名前（encoding の値）を、"
                       "どの道具にも encoding で渡してください（窓・欄・書き込み・段階が、その文字コードで読み直されます）。")
            out.append(f"  - {NAME_JA.get(self.encoding, self.encoding)}（この窓。encoding: \"{self.encoding}\"）："
                       f"{self.sample_of(self.text, self.candidates[0])}")
            for c in self.candidates:
                out.append(f"  - {c['label']}（encoding: \"{c['encoding']}\"）：{c['sample']}")
        if self.controls:
            out.append(f"- 制御文字（NUL など、タブ・改行以外）が {self.controls:,} 個あります。字としてそのまま読みました"
                       "（窓では見えないことがあります）。")
        return out

    @staticmethod
    def sample_of(text: str, cand: dict) -> str:
        """窓の文字コードで読んだ、候補と字の違う最初の所（見本）。"""
        if "own_sample" in cand:
            return cand["own_sample"]
        n = cand.get("line", 0)
        lines = text.splitlines()
        return _clip(lines[n]) if n < len(lines) else ""

    def shifts(self) -> "Shifts":
        """切り替えの印で字の読み方が変わる文字コードの、字とバイトの結びつき（1 回だけ作る）。"""
        tab = getattr(self, "_shift_tab", None)
        if tab is None:
            raw = getattr(self, "_raw", None)
            if raw is None:
                raise EncodingError("字の位置とバイトの位置を結べませんでした")
            tab = self._shift_tab = Shifts(raw[len(self.bom):], self.encoding, self.text)
        return tab

    def byte_of(self, i: int) -> int:
        """文字列の i 番目の字の始まりの、ファイルの中のバイトの位置。字ごとのバイト数を足し上げた表を 1 回だけ作るが、
        1024 字ごとの控えだけを持つ（字ごとに持つと、1 億字の行で何 GB にもなる）。ASCII の字だけなら表は要らない。
        切り替えの印のある文字コードは、その字の前の、書く所の境目（Shifts.at）。"""
        if stateful(self.encoding):
            return len(self.bom) + self.shifts().at(i)[0]
        if self.encoding not in WIDE and self.text.isascii():
            return len(self.bom) + i
        pre = getattr(self, "_prefix", None)
        if pre is None:
            pre = array("Q", [0])
            self._sizes: dict[str, int] = {}
            at = 0
            for k, ch in enumerate(self.text, 1):
                at += self._char_size(ch)
                if not k % 1024:
                    pre.append(at)
            self._prefix = pre
        q, r = divmod(i, 1024)
        at = pre[q]
        if r:
            for ch in self.text[q * 1024:i]:
                at += self._char_size(ch)
        return len(self.bom) + at

    def _char_size(self, ch: str) -> int:
        if ch < "\x80" and self.encoding not in WIDE:
            return 1
        n = self._sizes.get(ch)
        if n is None:
            n = self._sizes[ch] = len(ch.encode(self.encoding))
        return n

    def span(self, blob: bytes, a: int, b: int) -> tuple[int, int]:
        """字の範囲 [a, b) → バイトの範囲。読み直して同じ字になることを確かめる。"""
        if stateful(self.encoding):
            x, y, _sx, _sy = self.shifts().span(a, b)
            return len(self.bom) + x, len(self.bom) + y
        x, y = self.byte_of(a), self.byte_of(b)
        if blob[x:y].decode(self.encoding) != self.text[a:b]:
            raise EncodingError("字の位置とバイトの位置を結べませんでした")
        return x, y

    def patch(self, blob: bytes, a: int, b: int, s: str) -> tuple[int, int, bytes]:
        """字の範囲 [a, b) を s に置き換えるための (バイトの始め, 終わり, 置くバイト列)。切り替えの印のある文字コードは、
        置くバイト列の前後に、その所の状態に合わせる印を足す（書いた後ろの字の読み方が変わらないことを、デコーダで確かめる）。"""
        if not stateful(self.encoding):
            x, y = self.span(blob, a, b)
            return x, y, self.encode(s)
        x, y, new = self.shifts().plan([(a, b, s)], self.encode)[0]
        return len(self.bom) + x, len(self.bom) + y, new

    def pieces(self, blob: bytes, edits: list[tuple[int, int, str]]) -> list[tuple[int, int, bytes]]:
        """字の範囲の置き換え [(a, b, s)] → 前から順の [(バイトの始め, 終わり, 置くバイト列)]。重なる書き込みは断る。
        切り替えの印のある文字コードは、書き込みを前から順に、前の書き込みの後ろの状態から作る（Shifts.plan）。"""
        edits = sorted(edits, key=lambda e: (e[0], e[1]))
        for (_a1, b1, _s1), (a2, _b2, _s2) in zip(edits, edits[1:]):
            if a2 < b1:
                raise EncodingError("書く所が重なっています")
        if stateful(self.encoding):
            n = len(self.bom)
            return [(n + x, n + y, new) for x, y, new in self.shifts().plan(edits, self.encode)]
        return [(*self.span(blob, a, b), self.encode(v)) for a, b, v in edits]

    def rewrite(self, blob: bytes, edits: list[tuple[int, int, str]]) -> bytes:
        """字の範囲の置き換え [(a, b, s)] を全部した後の、ファイルのバイト列。書く所のほかのバイトは変えない。
        切り替えの印のある文字コードは、できたバイト列の全体を厳密に読んで、字の上での置き換えと同じ字になることを
        確かめる（ならなければ、書かずに断る）。"""
        edits = sorted(edits, key=lambda e: (e[0], e[1]))
        out = bytearray()
        at = 0
        for x, y, new in self.pieces(blob, edits):
            out += blob[at:x]
            out += new
            at = y
        out += blob[at:]
        if stateful(self.encoding):
            want, k = [], 0
            for a, b, v in edits:
                want += [self.text[k:a], v]
                k = b
            want.append(self.text[k:])
            if _strict(bytes(out[len(self.bom):]), self.encoding) != "".join(want):
                raise EncodingError(_CANNOT.format(NAME_JA.get(self.encoding, self.encoding)))
        return bytes(out)

    def encode(self, s: str) -> bytes:
        try:
            return s.encode(self.encoding)
        except UnicodeEncodeError as e:
            bad = s[e.start:e.end]
            raise EncodingError(f"「{bad}」は、このファイルの文字コード（{NAME_JA.get(self.encoding, self.encoding)}）では書けない字です"
                                + self._windows_hint(bad)) from None

    def _windows_hint(self, bad: str) -> str:
        """規格の ISO-2022-JP の仲間として読んだファイルに、Windows の ISO-2022-JP（CP50221）にだけある字を書こうとしたとき：
        ファイルが CP50221 としても読めるなら、そのことと、読める字が同じかを添える（どちらで読むかは、エージェントが決める）。"""
        raw = getattr(self, "_raw", None)
        if raw is None or not self.encoding.startswith("iso2022_jp") or self.encoding == ISO2022_MS:
            return ""
        try:
            bad.encode(ISO2022_MS)
        except UnicodeEncodeError:
            return ""
        t = _strict(raw, ISO2022_MS)
        if t is None:
            return ""
        same = "いまの字は、どちらで読んでも同じ" if t == self.text else "いまの字のうち、読み方の変わる字があります（〜 と ～ など）"
        return (f"（Windows の ISO-2022-JP（CP50221）にはある字です。このファイルは CP50221 としても読めます：{same}。"
                "Windows の決まりのファイルとして書くなら、encoding に cp50221 を渡します）")


@functools.lru_cache(maxsize=None)
def stateful(enc: str) -> bool:
    """切り替えの印で字の読み方が変わる文字コードか（ISO-2022・HZ・UTF-7）：同じバイトが、前の印によって別の字になるので、
    バイトの並びだけからは、字の境目・区切り・引用符を探せない。"""
    try:
        name = codecs.lookup(enc).name
    except LookupError:
        return False
    return name.startswith("iso2022") or name in ("hz", "utf-7")


# 切り替えのない状態に戻す印（ISO-2022-JP は ESC ( B：RFC 1468。ISO-2022-KR は SI：RFC 1557。HZ は ~}：RFC 1843。
# UTF-7 は「-」：RFC 2152）。どれを使うかは、デコーダで確かめて決める
RESETS = (b"\x1b(B", b"\x0f", b"~}", b"-")
_CANNOT = ("この所には、ファイルの文字コード（{}）では書けません"
           "（この文字コードの切り替えの印の決まりでは、ここに書くと、後ろの字の読み方を元と同じに保てません）。"
           "ここに書くには、ファイルを UTF-8 などの文字コードに直したものに書きます（直したことは本人に伝える）")


class Shifts:
    """切り替えの印で字の読み方が変わる文字コードの、字の位置とバイトの位置の結びつき。

    バイトを 1 つずつ増分のデコーダに渡し、字が出た所・切り替えの印が終わった所と、そこでのデコーダの状態を覚える。
    書くときは、置くバイト列を、書く所の前の状態から読んで、書く字になり、後ろの状態が元と同じになることを、デコーダで
    確かめる（後ろのバイトは変えないので、状態が同じなら、後ろの字の読み方も同じ）。
    - 後ろの状態を元と同じにできない所（ISO-2022-KR の、行の頭の印 ESC $ ) C より前に韓国語の字を書く：印を書いた後は、
      集合を指した状態を元に戻す印がない）は、後ろの元のバイトを、新しい状態と元の状態の 2 つのデコーダで並べて読み、
      同じ字が出て、2 つの状態が同じになる（合流する）か、ファイルの終わりまで同じ字なら、書ける（meets）。
    - いくつかのバイトがまとまって字になる所（UTF-7 の base64 のひと続き：字の境目がバイトの境目にない）の中・すぐ後ろに
      書くときは、書く範囲を、そのまとまりの字の全部に広げ、広げた分は同じ字で書き直す（widen）。
    - いくつもの書き込みは、前から順に作る：前の書き込みの後ろで状態が元と違う間（合流するまで）にある書き込みは、
      その違う状態から読んで確かめる（plan）。"""

    def __init__(self, blob: bytes, enc: str, text: str):
        self.blob, self.enc, self.text = blob, enc, text
        dec = codecs.getincrementaldecoder(enc)("strict")
        self.init = dec.getstate()[1]
        code = "I" if len(blob) < 1 << 32 else "Q"
        self.start = array(code)  # 字ごとの、その字のバイトの始まり（前の印の後ろ）
        self.end = array(code)  # 字ごとの、その字のバイトの終わり
        self.after = array("I")  # 字ごとの、その字の後ろの状態（states の番号）
        self.states = [self.init]
        self.marks: list[tuple[int, int, int, int]] = []  # 切り替えの印：(始まり, 終わり, 後ろの状態の番号, 次の字の番号)
        self.glued: set[int] = set()  # 前の字と同じバイト列から出た字（その前には書けない）
        self._plain: dict[int, bool] = {}
        seen = {self.init: 0}
        at = k = 0
        try:
            for p in range(len(blob)):
                out = dec.decode(blob[p:p + 1])
                pending, state = dec.getstate()
                if not out and pending:
                    continue  # 字・印の途中
                n = seen.get(state)
                if n is None:
                    n = seen[state] = len(self.states)
                    self.states.append(state)
                if out:
                    if text[k:k + len(out)] != out:
                        raise EncodingError("字の位置とバイトの位置を結べませんでした")
                    for j in range(len(out)):
                        self.start.append(p + 1 if j else at)
                        self.end.append(p + 1)
                        self.after.append(n)
                        if j:
                            self.glued.add(k + j)
                    k += len(out)
                else:
                    self.marks.append((at, p + 1, n, k))
                at = p + 1
            tail = dec.decode(b"", True)
        except UnicodeDecodeError:
            raise EncodingError("字の位置とバイトの位置を結べませんでした") from None
        if tail or k != len(text) or at != len(blob):
            raise EncodingError("字の位置とバイトの位置を結べませんでした")
        self.mark_of: dict[int, tuple[int, int]] = {}  # 字の番号 → その字の前の印（marks の範囲）
        for j, m in enumerate(self.marks):
            a, _b = self.mark_of.get(m[3], (j, j))
            self.mark_of[m[3]] = (a, j + 1)

    def gap(self, i: int) -> list[tuple[int, int]]:
        """i 番目の字の前の、字のない所の位置と、そこでの状態：[(バイトの位置, 状態)]（前の字の終わりから、印ごとに）。"""
        out = [(self.end[i - 1], self.states[self.after[i - 1]]) if i else (0, self.init)]
        a, b = self.mark_of.get(i, (0, 0))
        out += [(m[1], self.states[m[2]]) for m in self.marks[a:b]]
        return out

    def plain(self, state: int) -> bool:
        """ASCII の字を、そのまま ASCII の字として読む状態か（デコーダで確かめる。状態ごとに 1 回）。"""
        got = self._plain.get(state)
        if got is None:
            got = self._plain[state] = self.reads(b"a", state, "a", state)
        return got

    def at(self, i: int) -> tuple[int, int]:
        """i 番目の字の前の、書く所の境目：(バイトの位置, そこでの状態)。字の前の印のうち、ASCII の字をそのまま読む
        状態になるいちばん後ろの所（なければ、字のバイトの始まり）。境目は字ごとに 1 つなので、隣り合う書き込みは重ならない。"""
        g = self.gap(i)
        return next((x for x in reversed(g) if self.plain(x[1])), g[-1])

    def span(self, a: int, b: int) -> tuple[int, int, int, int]:
        if not 0 <= a <= b <= len(self.start):
            raise EncodingError("字の位置とバイトの位置を結べませんでした")
        if a in self.glued or b in self.glued:
            raise EncodingError(f"この所は、ファイルの文字コード（{NAME_JA.get(self.enc, self.enc)}）では、前の字と同じバイトの"
                                "並びから出た字の途中なので、ここだけを書き換えられません")
        (x, sx), (y, sy) = self.at(a), self.at(b)
        return x, y, sx, sy

    def recent(self, pos: int, most: int = 6) -> list[bytes]:
        """pos より前の切り替えの印を、後ろから 1 つ・2 つ…（most まで）つないだもの：書いた後ろの状態を、元と同じに戻す印の候補。"""
        from bisect import bisect_right

        j = bisect_right(self.marks, pos, key=lambda m: m[1])
        out, acc = [], b""
        for m in reversed(self.marks[max(0, j - most):j]):
            acc = self.blob[m[0]:m[1]] + acc
            out.append(acc)
        return out

    def reads(self, new: bytes, sx: int, s: str, sy: int) -> bool:
        """new を、状態 sx から読むと s になり、読んだ後の状態が sy か。"""
        dec = codecs.getincrementaldecoder(self.enc)("strict")
        try:
            dec.setstate((b"", sx))
            return dec.decode(new) == s and dec.getstate() == (b"", sy)
        except (UnicodeDecodeError, TypeError, ValueError, OverflowError):
            return False

    def widen(self, a: int, b: int) -> tuple[int, int]:
        """前の字と同じバイトの並びから出た字の途中にかかる範囲を、その並びの字の全部に広げる。"""
        while a in self.glued:
            a -= 1
        while b in self.glued:
            b += 1
        return a, b

    def _after(self, new: bytes, sx: int, s: str) -> int | None:
        """new を、状態 sx から読むと s になるとき、読んだ後の状態（字・印の途中で終わるなら None）。"""
        dec = codecs.getincrementaldecoder(self.enc)("strict")
        try:
            dec.setstate((b"", sx))
            if dec.decode(new) != s:
                return None
            pending, state = dec.getstate()
        except (UnicodeDecodeError, TypeError, ValueError, OverflowError):
            return None
        return None if pending else state

    def meets(self, y: int, t: int, sy: int) -> int | None:
        """y から後ろの元のバイトを、状態 t と、元の状態 sy の 2 つのデコーダで並べて読む。同じ字が出て、2 つの状態が同じに
        なった所（合流した所）のバイトの位置。ファイルの終わりまで同じ字なら、終わりの位置。字が違えば None。"""
        d1 = codecs.getincrementaldecoder(self.enc)("strict")
        d2 = codecs.getincrementaldecoder(self.enc)("strict")
        try:
            d1.setstate((b"", t))
            d2.setstate((b"", sy))
            pos, step, o1, o2 = y, 64, "", ""
            while pos < len(self.blob):
                chunk = self.blob[pos:pos + step]
                o1 += d1.decode(chunk)
                o2 += d2.decode(chunk)
                k = min(len(o1), len(o2))
                if o1[:k] != o2[:k]:
                    return None
                o1, o2 = o1[k:], o2[k:]
                pos += len(chunk)
                if not o1 and not o2 and d1.getstate() == d2.getstate():
                    return pos
                step = min(step * 2, 1 << 16)
            if o1 + d1.decode(b"", True) != o2 + d2.decode(b"", True):
                return None
        except (UnicodeDecodeError, TypeError, ValueError, OverflowError):
            return None
        return len(self.blob)

    def fit(self, y: int, sx: int, sy: int, s: str, core: bytes) -> tuple[bytes, tuple[int, int] | None]:
        """状態 sx の所に置いて s と読めるバイト列と、後ろの状態が元（sy）と違うときの (新しい状態, 合流する位置)。
        core は s を切り替えのない状態から書いたバイト列（終わりも切り替えのない状態）。core の前後の印を省いたもの・
        切り替えのない状態に戻す印を前に足したもの・後ろの状態に合わせる印を後ろに足したものを、短い順に、デコーダで
        確かめて選ぶ：後ろの状態が元と同じになるものが先。なければ、エンコーダが書くとおりに、切り替えのない状態に戻して
        終わるバイト列（後ろの印を省かず、足さないもの）のうち、後ろの元のバイトが同じ字に読めるもの（meets）。終わりの印を
        省いたものは、ここでは使わない（Python のデコーダは、ISO-2022-KR の SO を改行で戻すが、ほかの読み手は戻さない）。"""
        heads, tails = [0], [len(core)]
        if s:
            own = Shifts(core, self.enc, s)
            heads = [p for p, _st in own.gap(0)]
            tails = [p for p, _st in own.gap(len(s))]
        bodies = dict.fromkeys(core[p:q] for p in heads for q in tails if p <= q)
        pres = (b"",) + RESETS
        posts = [b""] + self.recent(y)
        cands = sorted(dict.fromkeys(pre + body + post for body in bodies for pre in pres for post in posts), key=len)
        closed = {pre + core[p:] for p in heads for pre in pres}
        after = []
        for new in cands:
            t = self._after(new, sx, s)
            if t == sy:
                return new, None
            after.append((new, t))
        for new, t in after:
            if t is not None and new in closed:
                c = self.meets(y, t, sy)
                if c is not None:
                    return new, (t, c)
        raise EncodingError(_CANNOT.format(NAME_JA.get(self.enc, self.enc)))

    def plan(self, edits: list[tuple[int, int, str]], encode) -> list[tuple[int, int, bytes]]:
        """字の範囲の置き換え [(a, b, s)]（前から順。重ならない）→ [(バイトの始め, 終わり, 置くバイト列)]。
        広げた範囲（widen）が重なる書き込みは、1 つにまとめる。"""
        groups: list[list] = []  # [広げた始め, 広げた終わり, [(a, b, s)]]
        for a, b, s in edits:
            if not 0 <= a <= b <= len(self.start):
                raise EncodingError("字の位置とバイトの位置を結べませんでした")
            wa, wb = self.widen(a, b)
            if groups and wa < groups[-1][1]:
                groups[-1][1] = max(groups[-1][1], wb)
                groups[-1][2].append((a, b, s))
            else:
                groups.append([wa, wb, [(a, b, s)]])
        out: list[tuple[int, int, bytes]] = []
        drift = None  # 前の書き込みの後ろで、状態が元と違う間：(バイトの位置, 字の番号, そこでの状態, 合流する位置)
        for wa, wb, parts in groups:
            s, k = "", wa
            for a, b, v in parts:
                s += self.text[k:a] + v
                k = b
            s += self.text[k:wb]
            x, y, sx, sy = self.span(wa, wb)
            if drift is not None and x < drift[3]:
                got = self._after(self.blob[drift[0]:x], drift[2], self.text[drift[1]:wa])
                if got is None:
                    raise EncodingError(_CANNOT.format(NAME_JA.get(self.enc, self.enc)))
                sx = got
            new, moved = self.fit(y, sx, sy, s, encode(s))
            out.append((x, y, new))
            drift = None if moved is None else (y, wb, moved[0], moved[1])  # 状態が元と同じに戻れば、後ろは元のまま
        return out


def read_as(raw: bytes, enc: str) -> str | None:
    """raw を、名前で指された文字コード enc で厳密に読んだ字（読めなければ None）。enc が規格の ISO-2022-JP の仲間で、規格の
    表では読めないときは、Windows の ISO-2022-JP（CP50221）で読む：Windows のメールは、ISO-2022-JP の名前で、規格にない
    NEC 特殊文字・IBM 拡張文字（①・㈱・髙 など）を書く。"""
    t = _strict(raw, enc)
    if t is None and enc.startswith("iso2022_jp") and enc != ISO2022_MS:
        t = _strict(raw, ISO2022_MS)
    return t


def kept_encoding(a: Decoded, b: Decoded, new: bytes) -> bool:
    """書いた後のファイル（b・new）が、元（a）と同じ文字コードのままか。ASCII の字だけになったファイルには、文字コードの
    手がかりが残らない：元の文字コードで厳密に読んで同じ字になれば、同じ文字コードのまま。"""
    if (a.encoding, a.bom) == (b.encoding, b.bom):
        return True
    if a.bom != b.bom or not b.ascii_only:
        return False
    return _strict(new[len(b.bom):], a.encoding) == b.text


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
    got = _readings(blob, encs, text_of)
    if not got:
        return None
    enc, text = got[0]
    cands = _line_candidates(text, got[1:])
    return Decoded(enc, b"", text, False, basis_one if not cands else "order", cands, controls(text))


def _readings(blob: bytes, encs: tuple, text_of=None) -> list[tuple[str, str]]:
    """encs のうち、厳密にきれいに読めるもの：[(文字コード, 字)]。読んだ字が前のものと同じなら入れない。"""
    got: list[tuple[str, str]] = []
    for enc in encs:
        t = _strict(blob, enc) if text_of is None else text_of(enc)
        if t is not None and clean(t) and all(t != u for _, u in got):
            got.append((enc, t))
    return got


def _line_candidates(text: str, others: list[tuple[str, str]]) -> list[dict]:
    cands = []
    for e2, t2 in others:
        k = _first_diff_line(text, t2)
        lines = t2.splitlines()
        cands.append({"encoding": e2, "label": NAME_JA.get(e2, e2), "line": k,
                      "sample": _clip(lines[k]) if k < len(lines) else ""})
    return cands


# ---------------------------------------------------------------------- 1バイトの文字コード
_ASCII = bytes(range(128))
_HIGH = re.compile(rb"[\x80-\xff]")
# str.splitlines と同じ行の区切り（どれも 0x40 より小さいので、多バイトの文字コードの 2 バイト目には現れない）
_SEP = re.compile(rb"\r\n|[\n\r\x0b\x0c\x1c-\x1e]")
SNIP = 48  # 候補の見本の字数（表示だけ。決め方には使わない）
_TABLES: dict[str, tuple] = {}


def _table(enc: str) -> tuple:
    """1バイトの文字コードの、0x80〜0xFF の読み方（字。割り当てがない・きれいでないバイトは None）。"""
    t = _TABLES.get(enc)
    if t is None:
        row = []
        for b in range(128, 256):
            try:
                ch = bytes([b]).decode(enc)
            except UnicodeDecodeError:
                ch = None
            row.append(ch if ch is not None and clean(ch) else None)
        t = _TABLES[enc] = tuple(row)
    return t


def high_bytes(chunks) -> list[int]:
    """ファイルにある 0x80 以上のバイトの種類（小さい順）。見つけた種類を消しながら進むので、ファイルを 1 回たどるだけ。"""
    gone, present = _ASCII, set()
    for c in chunks:
        rest = c.translate(None, gone)
        while rest:
            new = set(rest[:4096])
            present |= new
            gone += bytes(new)
            rest = rest.translate(None, gone)
    return sorted(present)


def single_readings(high: list[int]) -> list[tuple[str, tuple]]:
    """ファイルにある 0x80 以上のバイトの種類 → きれいに読める 1 バイトの文字コードと、そのバイトの読み方
    [(文字コード, 読み方)]。読み方が前のものと同じ（＝ファイルの字が同じ）なら入れない。"""
    out: list[tuple[str, tuple]] = []
    seen = set()
    for enc in SINGLE:
        tab = _table(enc)
        key = tuple(tab[b - 128] for b in high)
        if None in key or key in seen:
            continue
        seen.add(key)
        out.append((enc, key))
    return out


def differing(high: list[int], key1: tuple, key2: tuple) -> "re.Pattern":
    """2 つの 1 バイトの読み方で、字の違うバイトを探す式。"""
    return re.compile(b"[" + b"".join(re.escape(bytes([b])) for b, x, y in zip(high, key1, key2) if x != y) + b"]")


def first_offset(chunks, pat) -> int:
    """pat に合う最初のバイトの位置（なければ -1）。chunks は、ファイルの頭から順のバイト列。"""
    pos = 0
    for c in chunks:
        m = pat.search(c)
        if m:
            return pos + m.start()
        pos += len(c)
    return -1


def breaks_before(chunks, off: int) -> int:
    """off より前にある行の区切りの数（＝ off のある行の番号。0 から。str.splitlines と同じ数え方）。"""
    n, pos, cr = 0, 0, False
    for c in chunks:
        if pos >= off:
            break
        if pos + len(c) > off:
            c = c[:off - pos]
        n += len(_SEP.findall(c))
        if cr and c[:1] == b"\n":
            n -= 1  # 前の塊の終わりの CR と、この塊の頭の LF は 1 つの区切り
        cr = c.endswith(b"\r")
        pos += len(c)
    return n


def snippet(read, size: int, off: int, enc: str) -> str:
    """off（字の違う最初のバイト）のある行を enc で読んだ見本：違う所の少し前から SNIP 字まで（前後を切ったら「…」）。
    read(a, b) はファイルのバイトを返す。off より前の、その行のバイトは、くらべる読み方のどれでも同じ字（ASCII か、同じ
    読み方の 1 バイト）で、1 バイトが 1 字なので、行の途中から読んでよい。"""
    a0, b0 = max(0, off - 13), min(size, off + 4 * SNIP)
    piece = read(a0, b0)
    rel = off - a0
    a, cut_head = (1, True) if a0 else (0, False)  # 頭の 1 バイトは、行の区切りかどうかを見るだけ
    for m in _SEP.finditer(piece, 0, rel):
        a, cut_head = m.end(), False
    m = _SEP.search(piece, rel)
    b = m.start() if m else len(piece)
    s = piece[a:b].decode(enc, "ignore")  # 終わりで切れた多バイトの字は捨てる
    out = s[:SNIP].rstrip()
    return ("…" if cut_head else "") + out + ("…" if len(s) > SNIP or (m is None and b0 < size) else "")


def _legacy(blob: bytes) -> Decoded | None:
    """BOM も宣言もない 8 ビットの字のファイル：多バイト・1バイトの読める文字コードをすべて並べる（説明の 5・6）。"""
    multi = _readings(blob, MULTI)
    high = high_bytes((blob,))
    single = single_readings(high)
    if not multi and not single:
        return None
    lines: dict[int, int] = {}

    def cand(e2: str, off: int, own: str) -> dict:
        if off not in lines:
            lines[off] = breaks_before((blob,), off)
        return {"encoding": e2, "label": NAME_JA.get(e2, e2), "line": lines[off],
                "sample": _snip(blob, off, e2),
                "own_sample": _snip(blob, off, own)}

    if multi:
        enc, text = multi[0]
        cands = _line_candidates(text, multi[1:])
        # 1バイトの読み方が、窓（多バイト）の読み方と違う最初の所は、最初の 0x80 以上のバイト
        off = first_offset((blob,), _HIGH)
        cands += [cand(e2, off, enc) for e2, _key in single]
        basis = "multibyte"
    else:
        enc, key = single[0]
        text = blob.decode(enc)
        cands = [cand(e2, first_offset((blob,), differing(high, key, k2)), enc) for e2, k2 in single[1:]]
        basis = "singlebyte"
    return Decoded(enc, b"", text, False, basis if not cands else "order", cands, controls(text))


def _snip(blob: bytes, off: int, enc: str) -> str:
    return snippet(lambda a, b: blob[a:b], len(blob), off, enc)


def canonical(label: str) -> str | None:
    """文字コードの名前 → このモジュールの名前（Python の codec の名前）。知らない名前は None。"""
    import codecs

    x = (label or "").strip().lower()
    if not x:
        return None
    if x in ALIASES:
        return ALIASES[x]
    try:
        name = codecs.lookup(x).name
    except LookupError:
        return None
    return ALIASES.get(name, name)


_XML_DECL = re.compile(rb"^<\?xml\s[^>]*?\bencoding\s*=\s*[\"']([A-Za-z][A-Za-z0-9._:-]*)[\"']")
_META = re.compile(rb"<meta\b[^>]*?\bcharset\s*=\s*[\"']?\s*([A-Za-z][A-Za-z0-9._:-]*)", re.I)
_RTF_CPG = re.compile(rb"^\{\\rtf1?\b[^{}]*?\\ansicpg(\d+)")


def declared(blob: bytes) -> tuple[str, str] | None:
    """ファイルの中の、文字コードの宣言：(宣言の名前, どこの宣言か)。なければ None（このモジュールの説明の 4'）。"""
    m = _XML_DECL.match(blob[:512])
    if m:
        return m[1].decode("ascii"), "XML の宣言"
    m = _RTF_CPG.match(blob[:1024])
    if m:
        return f"cp{int(m[1])}", "RTF の \\ansicpg"
    m = _META.search(blob[:1024])
    if m:
        return m[1].decode("ascii"), "HTML の meta charset"
    return None


def decode(blob: bytes, encoding: str | None = None) -> Decoded:
    """バイト列 → 字（決め方はこのモジュールの説明のとおり）。encoding を渡せば、その文字コードで厳密に読む
    （エージェントが候補から選んだとき）。"""
    if encoding:
        return _given(blob, encoding)
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
        if b"\x1b" in blob:
            got = _escaped(blob)
            if got is not None:
                return got
        text = blob.decode("ascii")
        return Decoded("utf-8", b"", text, True, "ascii", [], controls(text))
    text = _strict(blob, "utf-8")
    if text is not None:
        return Decoded("utf-8", b"", text, False, "utf-8", [], controls(text))
    if b"\x00" not in blob:
        dec = declared(blob)
        enc = canonical(dec[0]) if dec else None
        if enc and enc not in WIDE:
            t = _strict(blob, enc)
            if t is not None and clean(t):
                # ファイルの中の宣言のとおりに、厳密にきれいに読める：宣言で決まる（仕様の決め方。候補はない）
                return Decoded(enc, b"", t, False, "declared", [], controls(t), dec[1])
    if b"\x00" not in blob:
        # NUL のバイトは、UTF-8・UTF-16・UTF-32 でなければ、文字のファイルにはない（POSIX の「文字のファイル」の定義）。
        # NUL があるのに UTF の形で読めないものは、多バイト・1バイトの文字コードを試さない（バイナリか、壊れたファイル）
        got = _legacy(blob)
        if got is not None:
            return got
    text = blob.decode("latin-1")
    return Decoded("latin-1", b"", text, False, "none", [], controls(text))


def _escaped(blob: bytes) -> Decoded | None:
    """ASCII のバイトだけで ESC のあるファイルを、ISO/IEC 2022 の切り替えの印のとおりに読む（決め方の 3'）。"""
    n = blob.count(b"\x1b")
    got = []
    for encs in (ISO2022, (ISO2022_MS,)):
        got = [(e, t) for e, t in _readings(blob, encs) if t.count("\x1b") < n and "\x0e" not in t and "\x0f" not in t]
        if got:
            break
    if not got:
        return None
    enc, text = got[0]
    cands = _line_candidates(text, got[1:])
    d = Decoded(enc, b"", text, False, "escape" if not cands else "order", cands, controls(text))
    d._raw = blob
    return d


def _given(blob: bytes, encoding: str) -> Decoded:
    """エージェントが選んで指定した文字コードで、厳密に読む（読めなければ EncodingError。きれいかは問わない：選んだのは
    エージェント）。BOM は、指定した文字コードの BOM なら外して覚える（書くときに元のまま残す）。"""
    enc = canonical(encoding)
    if enc is None:
        raise EncodingError(f"文字コード「{encoding}」を知りません（例：UTF-8・Shift_JIS・EUC-JP・GBK・Big5・EUC-KR・Windows-1252）")
    if enc in ("utf-16", "utf-32"):
        # 向き（エンディアン）は BOM で決める。BOM がなければ決められないので、向きまで指定してもらう
        hit = next(((b, e) for b, e in BOMS if e.startswith(enc) and blob.startswith(b)), None)
        if hit is None:
            raise EncodingError(f"{encoding}：BOM がないので、向きが決まりません。{enc}-le か {enc}-be を指定してください")
        enc = hit[1]
    bom = next((b for b, e in BOMS if e == enc and blob.startswith(b)), b"")
    try:
        text = blob[len(bom):].decode(enc, errors="strict")
    except UnicodeDecodeError as e:
        raise EncodingError(f"指定された文字コード（{NAME_JA.get(enc, enc)}）では、このファイルを厳密に読めません"
                            f"（{e.start} バイト目：{e.reason}）") from None
    except LookupError:
        raise EncodingError(f"文字コード「{encoding}」を知りません") from None
    d = Decoded(enc, bom, text, False, "given", [], controls(text))
    if stateful(enc):
        d._raw = blob
    return d
