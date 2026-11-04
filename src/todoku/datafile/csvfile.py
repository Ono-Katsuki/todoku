"""CSV・TSV を読み書きする（RFC 4180。TSV は区切りがタブのもの）。

読む：文字コード（BOM・UTF-8・Shift_JIS（CP932）・EUC-JP の順に決める）・改行（CRLF・LF・CR）・区切り・引用符の付け方を、
      ファイルから決まった規則で読み取り、マスごとに「元の字の範囲（引用符も含む）」と値を持つ。
書く：書いたマスのバイトだけを、そのファイルの文字コード・引用符の付け方で作った字に置き換える。ほかのマスの字・引用符・
      改行・BOM はそのまま。新しい行は、ファイルの終わりに、多く使われている改行で足す。
証明：書く前と後を読み直し、書いたマスの値だけが変わり、ほかのマスの元の字（引用符を含む）と改行が同じことを確かめる。
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from ..limits import budget
from .textenc import Decoded, EncodingError, decode

NEWLINE_JA = {"\r\n": "CRLF", "\n": "LF", "\r": "CR", "": "なし"}
DELIM_JA = {",": "「,」（カンマ）", "\t": "タブ", ";": "「;」（セミコロン）"}


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


def _delimiter(text: str, kind: str) -> str:
    """区切り：.tsv はタブ、.csv はカンマ。ただし .csv で、どの行にもカンマがなく、どの行にも同じ数のタブ（または「;」）が
    あるものは、それを区切りとする（引用符の中は数えない）。"""
    if kind == "tsv":
        return "\t"
    lines = [ln for ln in re.split(r"\r\n|\n|\r", re.sub(r'"(?:[^"]|"")*"', "", text[:200_000])) if ln.strip()][:200]
    if not lines or any("," in ln for ln in lines):
        return ","
    for d in ("\t", ";"):
        counts = {ln.count(d) for ln in lines}
        if len(counts) == 1 and counts != {0}:
            return d
    return ","


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
        self.delim = _delimiter(self.text, kind)
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
                "ends_with_newline": self.ends_with_newline, "delimiter": self.delim,
                "delimiter_ja": DELIM_JA.get(self.delim, repr(self.delim)), "quoting": q, "empty_lines": self.empty_lines,
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
