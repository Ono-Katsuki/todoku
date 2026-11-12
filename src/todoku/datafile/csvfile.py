"""CSV・TSV を読み書きする（RFC 4180。TSV は区切りがタブのもの）。

読む：文字コード（BOM・UTF-8・Shift_JIS（CP932）・EUC-JP の順に決める）・改行（CRLF・LF・CR）・区切り・引用符の付け方を、
      ファイルから決まった規則で読み取り、マスごとに「元の字の範囲（引用符も含む）」と値を持つ。
書く：書いたマスのバイトだけを、そのファイルの文字コード・引用符の付け方で作った字に置き換える。ほかのマスの字・引用符・
      改行・BOM はそのまま。新しい行は、ファイルの終わりに、多く使われている改行で足す。
証明：書く前と後を読み直し、書いたマスの値だけが変わり、ほかのマスの元の字（引用符を含む）と改行が同じことを確かめる。
"""

from __future__ import annotations

import contextlib
import hashlib
import re
from contextvars import ContextVar
from dataclasses import dataclass, field

from ..limits import budget
from .textenc import Decoded, EncodingError, decode

NEWLINE_JA = {"\r\n": "CRLF", "\n": "LF", "\r": "CR", "": "なし"}
DELIM_JA = {",": "「,」（カンマ）", "\t": "タブ", ";": "「;」（セミコロン）", "|": "「|」（縦線）"}


def delim_ja(d: str) -> str:
    return DELIM_JA.get(d, f"「{d}」")


class CsvError(ValueError):
    pass


@dataclass
class Field:
    a: int  # 元の字の範囲（引用符も含む）
    b: int
    value: str
    quoted: bool


@dataclass
class Record:
    a: int  # 行の中身の範囲（改行は含まない）
    b: int
    terminator: str  # 行の終わりの改行（最後の行で改行がなければ ""）
    fields: list[Field] = field(default_factory=list)

    @property
    def values(self) -> list[str]:
        return [f.value for f in self.fields]


def parse(text: str, delim: str) -> tuple[list[Record], list[str]]:
    """RFC 4180 の読み方。引用符は、マスの初めにあるときだけ意味を持つ（途中の引用符は字のまま）。
    閉じていない引用符・閉じた引用符のあとの字は、字のまま読み、problems に書く。"""
    out: list[Record] = []
    problems: list[str] = []
    n = len(text)
    i = 0
    while i < n:
        rec = Record(i, i, "")
        while True:
            a = i
            if i < n and text[i] == '"':
                j = i + 1
                buf = []
                closed = False
                while j < n:
                    k = text.find('"', j)
                    if k < 0:
                        break
                    if k + 1 < n and text[k + 1] == '"':
                        buf.append(text[j:k + 1])
                        j = k + 2
                        continue
                    buf.append(text[j:k])
                    j = k + 1
                    closed = True
                    break
                if not closed:
                    buf.append(text[j:])
                    j = n
                    problems.append(f"{len(out) + 1} 行目：閉じていない引用符があります（ファイルの終わりまでを1つのマスとして読みました）")
                value = "".join(buf)
                # 閉じた引用符のあとに、区切りでも改行でもない字が続く：字のまま、同じマスに続けて読む
                m = j
                while m < n and text[m] not in (delim, "\r", "\n"):
                    m += 1
                if m > j:
                    problems.append(f"{len(out) + 1} 行目：引用符を閉じたあとに字があります（「{text[j:m][:20]}」を同じマスの字として読みました）")
                    value += text[j:m]
                i = m
                rec.fields.append(Field(a, i, value, True))
            else:
                m = i
                while m < n and text[m] not in (delim, "\r", "\n"):
                    m += 1
                rec.fields.append(Field(a, m, text[a:m], False))
                i = m
            if i < n and text[i] == delim:
                i += 1
                continue
            break
        rec.b = i
        if text.startswith("\r\n", i):
            rec.terminator = "\r\n"
        elif i < n and text[i] in "\r\n":
            rec.terminator = text[i]
        i += len(rec.terminator)
        out.append(rec)
    return out, problems


SEPS = (",", "\t", ";")  # 区切りとして見る字（どれも、多バイトの文字コードの 2 バイト目には現れないバイト）
_SEP_LINE = re.compile(r"sep=([,;\t])(?:\r\n|\n|\r|$)")


def delimiters(text: str, kind: str) -> tuple[str, list[tuple[str, int]], bool]:
    """区切り：(区切り, ほかにも「そろっている」区切り [(区切り, 列の数)], 1行目の sep= で決めたか)。
    - .tsv はタブ。
    - 1行目が「sep=,」「sep=;」「sep=（タブ）」（Excel の、区切りの宣言の行）なら、その区切り。
    - .csv は、行ごとに、引用符の外の「,」・タブ・「;」の数を数える（頭の 200,000 字・字のある 200 行）。どの行にも同じ数
      （1つ以上）あることを「そろっている」と呼ぶ。
      1. カンマがそろっていれば、カンマ。
      2. そうでなければ、タブ・「;」のうち、そろっているもの（タブが先）。カンマが値の中にあるファイル（小数点がカンマの
         「Meier;1,50;Bonn」など）は、カンマの数が行によって違い、「;」の数はどの行も同じ。
      3. どれもそろっていなければ、カンマ（行によって列の数の違う CSV・1列だけのファイル）。
    - 選んだもののほかにも、そろっている区切りがあれば、決めた区切りが正しいとは限らない（見出しのない「1,5;2,5」は、
      カンマでも「;」でも同じ数に分かれる）。規則では決めず、ほかの区切りと、それで分けた列の数を返す。
    - 3 つの字がどの行にもない（窓は 1 列になる）ときは、ほかの ASCII の記号（引用符のほか）も同じように数え、そろっている
      ものを、ほかの区切りとして返す（「|」で区切ったファイル）。値の中の記号（日付の「-」）もそろうことがあるので、決めない。"""
    if kind == "tsv":
        return "\t", [], False
    m = _SEP_LINE.match(text)
    if m:
        return m[1], [], True
    lines = _sample(text)
    even = {d: n for d in SEPS if (n := _even(lines, d))}
    delim = next((d for d in SEPS if d in even), ",")
    if not even and lines and not any(d in ln for ln in lines for d in SEPS):
        even = {d: n for d in OTHER_SEPS if (n := _even(lines, d))}
    return delim, [(d, n) for d, n in even.items() if d != delim], False


OTHER_SEPS = tuple(c for c in map(chr, range(0x21, 0x7F)) if not c.isalnum() and c not in '",;')


def _sample(text: str) -> list[str]:
    """区切りを数える行：頭の 200,000 字の、引用符で囲んだ所を除いた、字のある 200 行。"""
    return [ln for ln in re.split(r"\r\n|\n|\r", re.sub(r'"(?:[^"]|"")*"', "", text[:200_000])) if ln.strip()][:200]


def _even(lines: list[str], d: str) -> int:
    """d がどの行にも同じ数（1つ以上）あれば、それで分けた列の数。そうでなければ 0。"""
    counts = {ln.count(d) for ln in lines}
    return counts.pop() + 1 if len(counts) == 1 and counts != {0} else 0


def _delimiter(text: str, kind: str) -> str:
    return delimiters(text, kind)[0]


class DelimiterError(CsvError):
    """エージェントが選んだ区切りを使えない。"""


# ---------------------------------------------------------------------- エージェントが選んだ区切り
# 区切りは、規則で決まらないことがある（どの行も同じ数に分かれる区切りが 2 つある・「|」などの、規則が数えない字で区切ったファイル）。
# エージェントが窓を読んで選んだ区切りを、1 回の呼び出しの間だけ置く（呼び出しが終われば戻す）。
_CHOSEN: ContextVar[str | None] = ContextVar("todoku_csv_delimiter", default=None)


def chosen_delimiter() -> str | None:
    return _CHOSEN.get()


def as_delimiter(value) -> str:
    """エージェントが渡した区切り → 区切りの字。ASCII の記号 1 字か、タブ（"\\t"・"tab" とも書ける）。
    英数字・空白・引用符・改行は、値の字・引用符・行の終わりと区別がつかないので、選べない。"""
    if isinstance(value, str) and value.lower() in ("\\t", "tab"):
        value = "\t"
    if (not isinstance(value, str) or len(value) != 1 or value.isalnum() or value in '" \r\n'
            or not (value == "\t" or 0x21 <= ord(value) <= 0x7E)):
        raise DelimiterError("区切り（delimiter）は、ASCII の記号 1 字かタブで指定してください（例：\",\"・\";\"・\"\\t\"・\"|\"。"
                             "英数字・空白・引用符・改行は選べません）")
    return value


@contextlib.contextmanager
def choosing(delimiter):
    """この中で読む CSV・TSV を、エージェントが選んだ区切りで読む（None なら何もしない）。"""
    if delimiter is None:
        yield
        return
    token = _CHOSEN.set(as_delimiter(delimiter))
    try:
        yield
    finally:
        _CHOSEN.reset(token)


def settle(rule: tuple[str, list, bool], pick: str | None, text: str) -> tuple[str, list, bool, str | None, int]:
    """(区切り, ほかにもそろう区切り, 1行目の sep= があるか, 規則で読むときの区切り（選んでいなければ None）, 選んだ区切りで
    そろう列の数（そろわなければ 0））。pick：エージェントが選んだ区切り。選んだら、その区切りで読む（規則の答えは、窓に残す）。
    そろうかどうかは、規則と同じ数え方（頭の 200 行・引用符の外。区切りの宣言の行は除く）。"""
    if pick is None:
        return rule[0], rule[1], rule[2], None, 0
    lines = _sample(text)
    return pick, [], rule[2], rule[0], _even(lines[1:] if rule[2] else lines, pick)


def delimiter_info(delim: str, also: list, declared: bool, rule: str | None = None, even: int = 0) -> dict:
    """窓の format に入れる、区切りの決め方（ほかの区切りでもそろうとき・宣言の行があるとき・エージェントが選んだときだけ、鍵を足す）。"""
    out = {"delimiter": delim, "delimiter_ja": delim_ja(delim)}
    if declared:
        out["delimiter_declared_by"] = "1行目の sep="
    if also:
        out["delimiter_also"] = [{"delimiter": d, "delimiter_ja": delim_ja(d), "columns": n} for d, n in also]
        if any(d not in SEPS for d, _n in also):
            out["delimiter_absent"] = True  # 「,」・タブ・「;」がどの行にもない（窓は 1 列）
    if rule is not None:
        out["delimiter_chosen_by"] = "エージェント"
        if rule != delim:
            out["delimiter_by_rule"] = rule
        if even:
            out["delimiter_even_columns"] = even
    return out


def delimiter_notes(fm: dict) -> list[str]:
    """Markdown の窓の注意（区切りを宣言の行で決めたとき・ほかの区切りでもそろうとき）。"""
    out = []
    if fm.get("delimiter_chosen_by"):
        rule = fm.get("delimiter_by_rule")
        was = "" if rule is None else f"規則で読むと、区切りは {delim_ja(rule)} です" + ("（1行目の宣言）。" if fm.get("delimiter_declared_by") else "。")
        n = fm.get("delimiter_even_columns")
        fact = (f"頭の 200 行までは、引用符の外で数えて、どの行も {n} 列に分かれます。" if n else
                "行によって列の数が違うか、区切りのない行があります（引用符の外で数えました）。")
        out.append(f"- 区切りは、エージェントが選んで指定した {fm['delimiter_ja']} で読みました（delimiter）。{fact}{was}"
                   "同じファイルには、どの道具にも同じ delimiter を渡してください。")
        if fm.get("delimiter_declared_by"):
            out.append("- 1行目（R1）は、区切りの宣言の行（Excel の「sep=」）です（R1 は列の見出しではありません。見出しは R2 から読んでください）。")
    elif fm.get("delimiter_declared_by"):
        out.append(f"- 1行目（R1）は、区切りの宣言の行（Excel の「sep=」）です。区切りは、宣言のとおり {fm['delimiter_ja']} で読みました"
                   "（R1 は列の見出しではありません。見出しは R2 から読んでください）。")
    also = fm.get("delimiter_also")
    if also:
        others = "・".join(f"{a['delimiter_ja']}（{a['columns']} 列）" for a in also)
        how = "そのときは、その区切りの字を delimiter で渡して読み直し、同じファイルには、どの道具にも同じ値を渡してください（書き込みは、読んだ分かれ方のマスに書きます）。窓の読み方が正しいときは、いまの区切りの字を渡すと、エージェントが確かめた読み方として残ります。"
        if fm.get("delimiter_absent"):
            out.append(f"- **区切りを規則では決められません**：「,」・タブ・「;」がどの行にもないので、窓は 1 列です。{others} なら、"
                       "どの行も同じ数の列に分かれます（引用符の外で数えました。頭の 200 行まで）。その記号が値の一部（日付の「-」など）なら、"
                       f"窓のとおり 1 列です。値を区切っている字なら、区切りはそちらです。{how}")
        else:
            out.append(f"- **区切りを規則では決められません**：{fm['delimiter_ja']} で読みましたが、{others} でも、どの行も同じ数の列に"
                       "分かれます（引用符の外で数えました）。窓の列が意味の通らない分かれ方（数の途中で切れているなど）なら、"
                       f"区切りはそちらです。{how}")
    return out


class CsvFile:
    """CSV・TSV のファイル。rows は値の表（1行目も含む。1行目を見出しとして扱うのは窓の側）。"""

    def __init__(self, blob: bytes, filename: str = "", kind: str = "csv", encoding: str | None = None):
        self.blob = blob
        self.filename = filename
        self.kind = kind
        try:
            self.dec: Decoded = decode(blob, encoding)
        except UnicodeDecodeError as e:
            raise CsvError(f"文字コードを読めません（{e.reason}）") from None
        self.text = self.dec.text
        # エージェントが選んだ区切り（choosing）があれば、その区切り。なければ、規則
        self.delim, self.delim_also, self.delim_declared, self.delim_rule, self.delim_even = settle(
            delimiters(self.text, kind), chosen_delimiter(), self.text)
        # 圧縮の爆弾（書庫の中の CSV）：バイトの上限の内でも、小さなマスが何千万もあれば、読むだけでメモリを使い切る。作る前に数える
        budget().count_part("csv", self.text.count(self.delim) + self.text.count("\n") + self.text.count("\r"), "CSV のマス")
        self.records, self.problems = parse(self.text, self.delim)
        terms = [r.terminator for r in self.records if r.terminator]
        self.newline = max(("\r\n", "\n", "\r"), key=lambda t: (terms.count(t), t == "\r\n")) if terms else "\r\n"
        self.mixed_newlines = len(set(terms)) > 1
        self.ends_with_newline = bool(self.records) and bool(self.records[-1].terminator)
        filled = [f for r in self.records for f in r.fields if f.value != ""]
        empty = [f for r in self.records for f in r.fields if f.value == ""]
        self.quote_all = bool(filled) and all(f.quoted for f in filled)  # 字のあるマスが、すべて引用符で囲まれている
        self.quote_empty = self.quote_all and bool(empty) and all(f.quoted for f in empty)
        self.width = max((len(r.fields) for r in self.records), default=0)
        # 字のない行（改行だけの行。ファイルの終わりの空の行も）：RFC 4180 の文法では「空のマスが1つの行」として読む
        self.empty_lines = [i for i, r in enumerate(self.records, 1) if r.a == r.b]

    # ------------------------------------------------------------------ 問い合わせ
    @property
    def rows(self) -> list[list[str]]:
        return [r.values for r in self.records]

    @property
    def header(self) -> list[str]:
        return self.records[0].values if self.records else []

    def sha256(self) -> str:
        return hashlib.sha256(self.blob).hexdigest()

    def format_info(self) -> dict:
        q = "すべてのマスを引用符で囲む" if self.quote_all else "必要なマスだけ引用符で囲む"
        return {"encoding": self.dec.encoding, "encoding_ja": self.dec.label, "bom": bool(self.dec.bom),
                "ascii_only": self.dec.ascii_only, "newline": NEWLINE_JA[self.newline], "mixed_newlines": self.mixed_newlines,
                "ends_with_newline": self.ends_with_newline,
                **delimiter_info(self.delim, self.delim_also, self.delim_declared, self.delim_rule, self.delim_even),
                "quoting": q, "empty_lines": self.empty_lines,
                **self.dec.info(), "encoding_notes": self.dec.note_lines()}

    # ------------------------------------------------------------------ id
    def resolve(self, key: str) -> tuple[int, int]:
        """id → (行, 列)（どちらも 1 から）。R3.C2・R3C2・R3.氏名・R3.「氏名」・行3.列2・行3.列「氏名」。"""
        k = str(key).strip()
        m = re.fullmatch(r"(?:R|行)\s*(\d+)\s*[.．]?\s*(?:C|列)?\s*(.+)", k, re.I)
        if not m:
            raise CsvError(f"「{key}」はマスの id ではありません（R行.C列 の形。例 R3.C2・R3.氏名・行3.列「氏名」）")
        r = int(m[1])
        col = m[2].strip()
        if re.fullmatch(r"\d+", col):
            c = int(col)
        else:
            name = col.strip("「」\"' ")
            hits = [i + 1 for i, h in enumerate(self.header) if h.strip() == name]
            if not hits:
                raise CsvError(f"「{name}」という見出しの列がありません（見出し：" + "、".join(f"「{h}」" for h in self.header[:30]) + "）")
            if len(hits) > 1:
                raise CsvError(f"「{name}」という見出しの列が {len(hits)} つあります（C{hits[0]}・C{hits[1]} など）。列の番号で指定してください")
            c = hits[0]
        if r < 1 or c < 1:
            raise CsvError(f"「{key}」：行と列は 1 から数えます")
        return r, c

    def value(self, r: int, c: int) -> str | None:
        if r > len(self.records):
            return None
        fs = self.records[r - 1].fields
        return fs[c - 1].value if c <= len(fs) else None


# ====================================================================== 書く
def quote(value: str, f: CsvFile, was_quoted: bool, empty_ok: bool = True) -> str:
    """書く字：区切り・引用符・改行を含むときは必ず引用符で囲む（中の引用符は2つ重ねる）。元のマスが囲まれていたとき・
    ファイルがすべてのマスを囲むときも囲む。"""
    need = any(ch in value for ch in (f.delim, '"', "\r", "\n"))
    if f.kind == "tsv" and not was_quoted and not f.quote_all and ("\t" in value):
        need = True
    if need or was_quoted or (f.quote_all and (value != "" or f.quote_empty)):
        return '"' + value.replace('"', '""') + '"'
    return value


def to_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


@dataclass
class CsvPatch:
    id: str  # R行.C列
    row: int
    col: int
    value: str
    key: str = ""  # 指定したときの名前


@dataclass
class CsvFillResult:
    blob: bytes
    patches: list[CsvPatch] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    appended_rows: int = 0
    formula_like: list[dict] = field(default_factory=list)  # 表計算で数式として読まれうる値（formula_like を見る）


# 表計算（Excel・LibreOffice）が、CSV のマスを数式として読む先頭の字（OWASP の CSV Injection）：= + - @ タブ CR
FORMULA_START = ("=", "+", "-", "@", "\t", "\r")
_PLAIN_NEGATIVE = __import__("re").compile(r"-\d+(?:[.,]\d+)?")


def formula_like(text: str) -> bool:
    """値が、表計算で開くと数式として読まれうる形か（先頭が FORMULA_START の字。ただの負の数「-5」は除く）。"""
    return text[:1] in FORMULA_START and not _PLAIN_NEGATIVE.fullmatch(text)


def formula_note(key: str, text: str) -> str:
    shown = text.replace("\t", "（タブ）").replace("\r", "（CR）")
    head = {"\t": "タブ", "\r": "CR"}.get(text[:1], text[:1])
    return (f"{key}: 「{shown}」は「{head}」で始まるので、Excel・LibreOffice で開くと"
            f"数式として計算されることがあります（CSV インジェクション）。値は変えずに、そのまま書きました。字のまま見せるなら、先頭に"
            f"「'」を付けた「'{shown}」で書き直せます（読み込む仕組みによっては「'」も字として残るので、どちらにするか本人に確かめてください）")


def fill(f: CsvFile, values: dict) -> CsvFillResult:
    """values：{id: 値}。今ある行のマスは、そのマスの字だけを置き換える。行の右に足りないマスは、区切りを足して書く。
    今ある行の次から（R{行数+1}〜）に書くと、新しい行をファイルの終わりに足す（行を飛ばしては足さない）。"""
    want: dict[tuple[int, int], tuple[str, str]] = {}
    problems: list[str] = []
    for key, v in values.items():
        try:
            r, c = f.resolve(key)
        except CsvError as e:
            problems.append(str(e))
            continue
        if (r, c) in want:
            problems.append(f"{key}: 同じマス（R{r}.C{c}）に2回書こうとしています")
            continue
        text = to_text(v)
        if f.kind == "tsv" and ("\t" in text) and not f.quote_all and not any(x.quoted for rec in f.records for x in rec.fields):
            problems.append(f"{key}: タブ区切りのファイルなので、値にタブは書けません")
            continue
        want[(r, c)] = (text, str(key))
    n = len(f.records)
    new_rows = sorted({r for r, _c in want if r > n})
    if new_rows and new_rows != list(range(n + 1, n + 1 + len(new_rows))):
        problems.append(f"新しい行は、今ある {n} 行の次（R{n + 1}）から続けて足します（行を飛ばせません：" + "、".join(f"R{r}" for r in new_rows) + "）")
    if problems:
        raise CsvError("\n".join(problems))
    edits: list[tuple[int, int, bytes]] = []
    patches: list[CsvPatch] = []
    warnings: list[str] = []
    try:
        for r in range(1, n + 1):
            cols = sorted(c for (rr, c) in want if rr == r)
            if not cols:
                continue
            rec = f.records[r - 1]
            tail: list[str] = []
            for c in cols:
                text, key = want[(r, c)]
                if c <= len(rec.fields):
                    fd = rec.fields[c - 1]
                    x, y = f.dec.span(f.blob, fd.a, fd.b)
                    edits.append((x, y, f.dec.encode(quote(text, f, fd.quoted))))
                else:
                    tail.append(str(c))
                patches.append(CsvPatch(f"R{r}.C{c}", r, c, text, key))
            if tail:
                # 行の右に足りないマス：区切りを足して、行の終わり（改行の前）に書く
                last = len(rec.fields)
                hi = max(int(c) for c in tail)
                cells = []
                for c in range(last + 1, hi + 1):
                    cells.append(quote(want[(r, c)][0], f, False) if (r, c) in want else
                                 ('""' if f.quote_empty else ""))
                at = f.dec.byte_of(rec.b)
                edits.append((at, at, f.dec.encode(f.delim + f.delim.join(cells))))
                warnings.append(f"R{r}：もとの行は {last} マスなので、区切りを足して C{hi} まで書きました")
        if new_rows:
            width = max(f.width, max(c for (_r, c) in want))
            lines = []
            for r in new_rows:
                cells = []
                for c in range(1, width + 1):
                    if (r, c) in want:
                        cells.append(quote(want[(r, c)][0], f, False))
                        patches.append(CsvPatch(f"R{r}.C{c}", r, c, want[(r, c)][0], want[(r, c)][1]))
                    else:
                        cells.append('""' if f.quote_empty else "")
                lines.append(f.delim.join(cells))
            add = (f.newline if f.records and not f.ends_with_newline else "") + f.newline.join(lines) + \
                (f.newline if f.ends_with_newline or not f.records else "")
            edits.append((len(f.blob), len(f.blob), f.dec.encode(add)))
            if f.width and width > f.width:
                warnings.append(f"新しい行は {width} マスです（もとの表は {f.width} 列）")
    except EncodingError as e:
        raise CsvError(str(e)) from None
    if f.dec.ascii_only and any(not p.value.isascii() for p in patches):
        warnings.append("もとのファイルは ASCII の字だけで文字コードが決まらないので、UTF-8（BOM なし）で書きました")
    out = bytearray()
    at = 0
    for x, y, b in sorted(edits, key=lambda e: (e[0], e[1])):
        out += f.blob[at:x]
        out += b
        at = y
    out += f.blob[at:]
    patches.sort(key=lambda p: (p.row, p.col))
    risky = [{"id": p.key, "value": p.value, "safe_form": "'" + p.value} for p in patches if formula_like(p.value)]
    warnings += [formula_note(x["id"], x["value"]) for x in risky]
    return CsvFillResult(bytes(out), patches, warnings, len(new_rows), risky)


# ====================================================================== 証明
@dataclass
class CsvProof:
    original_sha256: str
    new_sha256: str
    cells_total: int = 0
    cells_changed: int = 0
    rows_added: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        add = f"・足した行 {self.rows_added}" if self.rows_added else ""
        return (f"マス {self.cells_total:,} のうち、変わったのは書いた {self.cells_changed} マスだけ{add}。"
                "ほかのマスの字・引用符・改行・文字コードはそのまま"
                + ("" if not self.unexpected else f"（想定外の変更 {len(self.unexpected)} 件）"))


def prove(original: bytes, new: bytes, written: dict[tuple[int, int], str], filename: str = "", kind: str = "csv",
          encoding: str | None = None) -> CsvProof:
    """written：(行, 列) → 書いた値。書く前と後を、それぞれ読み直して比べる（encoding：エージェントが選んだ文字コード。
    前と後の両方を、その文字コードで読む）。"""
    pr = CsvProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    a = CsvFile(original, filename, kind, encoding)
    try:
        b = CsvFile(new, filename, kind, encoding)
    except CsvError as e:
        pr.lint_errors.append(f"書いた後のファイルを読めません：{e}")
        return pr
    if (a.dec.encoding, a.dec.bom) != (b.dec.encoding, b.dec.bom) and not a.dec.ascii_only:
        pr.unexpected.append(f"文字コードが変わっています（{a.dec.label} → {b.dec.label}）")
    if a.delim != b.delim:
        pr.unexpected.append("区切りの字が変わっています")
    pr.lint_errors += [p for p in b.problems if p not in a.problems]
    n = len(a.records)
    pr.rows_added = len(b.records) - n
    if pr.rows_added < 0 or (pr.rows_added and not all(r > n for r, _c in written if r > n)):
        pr.unexpected.append(f"行の数が変わっています（{n} → {len(b.records)}）")
        return pr
    want_new = sorted({r for r, _c in written if r > n})
    if want_new != list(range(n + 1, n + 1 + pr.rows_added)):
        pr.unexpected.append(f"足した行の数が、書いた行と合いません（{pr.rows_added} 行）")
    for i, (ra, rb) in enumerate(zip(a.records, b.records), 1):
        last = i == n
        if ra.terminator != rb.terminator and not (last and not ra.terminator and pr.rows_added):
            pr.unexpected.append(f"{i} 行目の改行が変わっています（{NEWLINE_JA[ra.terminator]} → {NEWLINE_JA[rb.terminator]}）")
        for c in range(1, max(len(ra.fields), len(rb.fields)) + 1):
            pr.cells_total += 1
            fa = ra.fields[c - 1] if c <= len(ra.fields) else None
            fb = rb.fields[c - 1] if c <= len(rb.fields) else None
            if (i, c) in written:
                pr.cells_changed += 1
                if fb is None or fb.value != written[(i, c)]:
                    pr.unexpected.append(f"R{i}.C{c}：書いた後の値が「{fb.value if fb else ''}」で、書いた「{written[(i, c)]}」と違います")
                continue
            if fa is None:
                if fb is not None and fb.value != "":
                    pr.unexpected.append(f"R{i}.C{c}：書いていないマスに字が入っています")
                continue
            if fb is None or a.text[fa.a:fa.b] != b.text[fb.a:fb.b]:
                pr.unexpected.append(f"R{i}.C{c}：書いていないマスの字が変わっています")
    for i in range(n + 1, len(b.records) + 1):
        rb = b.records[i - 1]
        for c, fb in enumerate(rb.fields, 1):
            pr.cells_total += 1
            if (i, c) in written:
                pr.cells_changed += 1
                if fb.value != written[(i, c)]:
                    pr.unexpected.append(f"R{i}.C{c}：足した行の値が「{fb.value}」で、書いた「{written[(i, c)]}」と違います")
            elif fb.value != "":
                pr.unexpected.append(f"R{i}.C{c}：足した行の、書いていないマスに字が入っています")
    for (r, c) in written:
        if r > len(b.records):
            pr.unexpected.append(f"R{r}.C{c}：書いた行がありません")
    # バイトの比べ：書いた行より前は、書く前とバイト一致（BOM・文字コードの違いも、ここで分かる）
    first = min((r for r, _c in written), default=n + 1)
    if first > 1:
        cut = a.dec.byte_of(a.records[first - 2].b + len(a.records[first - 2].terminator)) if first - 2 < n else len(original)
        if original[:cut] != new[:cut]:
            pr.unexpected.append("書いた行より前の部分のバイトが変わっています")
    # 書いた行より後ろも、書く前とバイト一致（行を足したときは、元の最後の行の改行が変わりうるので、足した行の証明に任せる）
    last = max((r for r, _c in written if r <= n), default=0)
    if last and last < n and not pr.rows_added:
        ra_, rb_ = a.records[last - 1], b.records[last - 1]
        ca_ = a.dec.byte_of(ra_.b + len(ra_.terminator))
        cb_ = b.dec.byte_of(rb_.b + len(rb_.terminator))
        if original[ca_:] != new[cb_:]:
            pr.unexpected.append("書いた行より後ろの部分のバイトが変わっています")
    return pr
