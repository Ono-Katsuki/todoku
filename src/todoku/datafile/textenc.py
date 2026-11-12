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

import re
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
            "utf-8": "UTF-8 として厳密に読める", "multibyte": "多バイトの文字コードのうち、これだけが厳密に読める（1バイトの文字コードでは、きれいに読めない）",
            "singlebyte": "1バイトの文字コードのうち、これだけがきれいに読める（多バイトの文字コードでは読めない）",
            "order": "規則では決められない（いくつもの文字コードで読める。決まった順の前のもので読んだ）",
            "none": "どの文字コードとしても読めない",
            "declared": "ファイルの中の宣言（XML の宣言・HTML の meta charset・RTF の \\ansicpg）",
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

    def byte_of(self, i: int) -> int:
        """文字列の i 番目の字の始まりの、ファイルの中のバイトの位置。字ごとのバイト数を足し上げた表を 1 回だけ作るが、
        1024 字ごとの控えだけを持つ（字ごとに持つと、1 億字の行で何 GB にもなる）。ASCII の字だけなら表は要らない。"""
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
    return Decoded(enc, bom, text, False, "given", [], controls(text))
