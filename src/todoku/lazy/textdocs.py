"""CSV・TSV・JSON Lines・.txt を分けて読み書きする。メモリは、頼んだ範囲に比例する（ファイルの大きさには比例しない）。

- 読む：記録（行）の位置の索引（textstream）から範囲のバイトだけを読み、datafile・txt の読み手でその範囲を読む
- 欄：頼んだ範囲で、空のマス・欄を見つける
- 書く：書く記録のバイトだけを置き換えて、ファイルを写す（stream_patch）
- 証明：書いた記録を比べ（書いたマスのほかは字も引用符も同じ）、書き換えなかった範囲のハッシュを、書いたファイルを読み直して比べる
- 読み戻し：書いたファイルの、書いた記録の位置を読み直す（窓の読み手と、別の読み手の両方で）
- とても長い行：字の範囲で分けて出し、分けたことを印で示す。.txt の id は、そのような行では字の位置で決まる P{行}.x{字の位置}
"""

from __future__ import annotations

import bisect
import codecs
import csv as _csv
import io
import json
import re
from pathlib import Path

from .. import limits
from ..datafile.csvfile import (
    NEWLINE_JA,
    CsvError,
    CsvFile,
    CsvPatch,
    chosen_delimiter,
    delimiter_info,
    delimiter_notes,
    delimiters,
    settle,
    formula_like,
    formula_note,
    parse,
    quote,
    to_text,
)
from ..datafile.textenc import Decoded, EncodingError, NAME_JA, stateful
from ..datafile.window import cell_text, row_line
from .base import ByteRanges, LazyDoc, Page, RangeError, comma, cut_note, parse_range, stream_patch, temp_beside, verify_ranges
from .textstream import (
    STRIDE,
    CsvStats,
    Source,
    TextIndex,
    detect,
    head_text,
    index_csv,
    index_lines,
    recode,
    second_byte_leads,
    span_hash,
)


_csv.field_size_limit(2 ** 31 - 1)  # 別の読み方の確かめで、長いマスを「壊れている」としない（既定は 131072 字）


class TextDoc(LazyDoc):
    streams = True

    def __init__(self, path: str, name: str, kind: str, encoding: str | None):
        self.path = path
        self.name = name
        self.kind = kind
        self._src = Source(path)
        self.file_bytes = self._src.size
        self.est_chars = self._src.size  # 字の数はバイト数を超えない
        self._encoding = encoding
        self._dec = None
        self._idx: TextIndex | None = None
        self.told = None  # 切り替えの印のある文字コードのファイル：元のファイルの読み方（読むのは、UTF-8 に直した写し）

    @property
    def src(self):
        """読むバイトの元。文字コードを決めてから渡す（切り替えの印のある文字コードは、UTF-8 に直した写しに替わる）。"""
        self.dec  # noqa: B018
        return self._src

    @property
    def dec(self):
        if self._dec is None:
            d = detect(self._src, self._encoding)
            if stateful(d.encoding):
                # 同じバイトが、前の印によって別の字になる：行・区切り・引用符をバイトのまま探せないので、字を UTF-8 に直した
                # 写しの上で読む。書き込み・検査・段階は、元のファイルの全体を読む道でする（streams を False に）
                was = self._src
                self._src, d2 = recode(was, d)
                was.close()
                self.told, d = d, d2
                self.streams = False
            self._dec = d
        return self._dec

    @property
    def recoded(self) -> bool:
        self.dec  # noqa: B018
        return self.told is not None

    @property
    def shown(self):
        """伝えるときの、ファイルの文字コード（写しを読むときも、元のファイルの読み方）。"""
        d = self.dec
        return self.told or d

    def _audit_whole(self, state: dict) -> dict:
        from .. import api

        limits.progress("書類の全体を検査しています")
        au = api.audit_whole(self.path, None, self._encoding)
        state["checks"] = au.get("checks", [])
        state["full"] = au
        state["done"] = True
        return state

    @property
    def start(self) -> int:
        return len(self.dec.bom)

    def text_of(self, a: int, b: int) -> str:
        data = self.src.read(a, b)
        return data.decode(self.dec.encoding)

    def _slice_dec(self, text: str) -> Decoded:
        return Decoded(self.dec.encoding, b"", text, False, "given")

    def source_line(self) -> str:
        line = f"- 元ファイル {self.name}（{self.file_bytes:,} バイト）。SHA-256 は書き込みの結果と検査で示します（ここでは計算しません）"
        if self.recoded:
            line += (f"\n- このファイルの文字コード（{self.told.label}）は、切り替えの印（ESC など）で、同じバイトの読み方が変わります。"
                     "範囲で読むときは、字を UTF-8 に直した写しを読みます（字は同じ。行のバイト数は、写しでの数）。"
                     "書き込みと検査は、元のファイルの全体を読んでします。")
        return line

    def _guide(self) -> list[str]:
        return ["- 窓の全部を一度には返しません（読む量を分けるだけで、読めない所はありません）。"
                "read_window に範囲を渡すと、その範囲を返します。cursor をたどると、最後まで読めます。"]


# ====================================================================== CSV・TSV
class CsvDoc(TextDoc):
    def __init__(self, path: str, name: str, kind: str, encoding: str | None):
        super().__init__(path, name, kind, encoding)
        self._st: CsvStats | None = None
        self._head: list[str] | None = None
        self._delim = None
        self.delim_also: list = []
        self.delim_declared = False
        self.delim_rule = None
        self.delim_even = 0
        self._pick = chosen_delimiter()  # エージェントが選んだ区切り（この書類を作った呼び出しのもの）

    @property
    def delim(self) -> str:
        if self._delim is None:
            head = head_text(self.src, self.start, self.dec.encoding)
            delim, self.delim_also, self.delim_declared, self.delim_rule, self.delim_even = settle(
                delimiters(head, self.kind), self._pick, head)
            second_byte_leads(self.dec.encoding, delim)  # バイトの並びから探せない区切りは、ここで断る
            self._delim = delim
        return self._delim

    @property
    def idx(self) -> TextIndex:
        if self._idx is None:
            self._idx, self._st = index_csv(self.src, self.start, self.dec.encoding, self.delim)
        return self._idx

    @property
    def st(self) -> CsvStats:
        self.idx  # noqa: B018
        return self._st

    # csvfile の quote・resolve が使う問い合わせ
    @property
    def quote_all(self) -> bool:
        return self.st.quote_all

    @property
    def quote_empty(self) -> bool:
        return self.st.quote_empty

    @property
    def width(self) -> int:
        return self.st.width

    @property
    def header(self) -> list[str]:
        if self._head is None:
            self._head = self.record_values(1) if self.idx.count else []
        return self._head

    resolve = CsvFile.resolve

    def records(self, first: int, last: int):
        """[(行の番号, Record, 記録のテキストの Decoded, 範囲の頭のバイト位置)]：first〜last（1 から）。"""
        a, b, base = self.idx.bucket(first, last)
        text = self.text_of(a, b)
        recs, _p = parse(text, self.delim)
        dec = self._slice_dec(text)
        return [(base + i, r, dec, a) for i, r in enumerate(recs) if first <= base + i <= last]

    def record_values(self, r: int) -> list[str]:
        got = self.records(r, r)
        return got[0][1].values if got else []

    def _label(self, c: int) -> str:
        head = self.header
        return head[c - 1].strip() if c <= len(head) and head[c - 1].strip() else f"C{c}"

    # ------------------------------------------------------------------ 概要
    def format_info(self) -> dict:
        idx, st = self.idx, self.st
        q = "すべてのマスを引用符で囲む" if st.quote_all else "必要なマスだけ引用符で囲む"
        return {"encoding": self.shown.encoding, "encoding_ja": self.shown.label, "bom": bool(self.shown.bom),
                "ascii_only": self.shown.ascii_only, "newline": NEWLINE_JA[idx.newline or "\r\n"], "mixed_newlines": idx.mixed_newlines,
                "ends_with_newline": idx.ends_with_newline,
                **delimiter_info(self.delim, self.delim_also, self.delim_declared, self.delim_rule, self.delim_even),
                "quoting": q,
                "empty_lines_count": st.empty_lines, "empty_lines_first": st.first_empty_lines,
                **self.shown.info(), "encoding_notes": self.shown.note_lines()}

    def overview(self) -> dict:
        return {"format": self.format_info(), "rows": self.idx.count, "columns": self.width, "header": self.header,
                "problems": self.st.first_problems, "problem_count": self.st.problem_rows, "ranges": "rows（行の範囲）",
                "bytes": self.file_bytes}

    def overview_text(self) -> str:
        fm = self.format_info()
        n = self.idx.count
        kind = "TSV（タブ区切り）" if self.kind == "tsv" else "CSV"
        out = [f"# {self.name}", "", self.source_line(),
               f"- {kind}・{n:,} 行・{self.width} 列。文字コード {fm['encoding_ja']}／改行 {fm['newline']}"
               + ("（行によって改行が違います）" if fm["mixed_newlines"] else "")
               + f"／区切り {fm['delimiter_ja']}／引用符 {fm['quoting']}"
               + ("" if fm["ends_with_newline"] else "／最後の行に改行なし"),
               f"- 大きいので、ここでは全体を出していません。read_window に rows: \"2-500\" のように行の範囲を渡して読みます"
               f"（cursor をたどれば R1〜R{n} のすべてを読めます）。list_fields も rows で、その範囲の空のマスを出します。",
               f"- マスの id は R行.C列（R2.C3・R2.氏名）。どのマスにも id で書けます。R{n + 1}.C1 に書くと、行を足します。"
               "書き込みは、書いた行のバイトだけを変え、ほかのバイトはそのままファイルを写します。",
               "", "## 列", ""]
        out += [f"- C{i}：{h}" for i, h in enumerate(self.header, 1)]
        if self.width > len(self.header):
            out.append(f"- C{len(self.header) + 1}〜C{self.width}：1 行目より右にある列（見出しなし）")
        if self.st.empty_lines:
            out.append(f"- 字のない行（改行だけ）：{self.st.empty_lines:,} 行（はじめの方：" + "、".join(f"R{r}" for r in self.st.first_empty_lines) + "）")
        if fm.get("ascii_only"):
            out.append("- このファイルは ASCII の字だけなので、文字コードは決まりません（日本語を書くと UTF-8 で書きます）。")
        out += fm.get("encoding_notes") or []
        out += delimiter_notes(fm)
        for p in self.st.first_problems[:20]:
            out.append(f"- **書き方の問題**：{p}")
        if self.st.problem_rows > 20:
            out.append(f"- **書き方の問題**はほかに {self.st.problem_rows - 20:,} か所（窓の範囲を読むと、その行の問題を示します）")
        return "\n".join(out) + "\n"

    # ------------------------------------------------------------------ 読む
    def read(self, rng: dict, max_chars: int) -> Page:
        n = self.idx.count
        a, b, skip = _row_args(rng, n)
        width = max(self.width, 1)
        head = ["| 行 | " + " | ".join(f"C{c}" for c in range(1, width + 1)) + " |", "|---" * (width + 1) + "|"]
        if a > 1 and n:
            head.append(row_line("R1（見出し）", self.header, width))
        body: list[str] = []
        size = sum(len(x) + 1 for x in head)
        r = a
        stop = None
        cut = None
        while r <= b:
            last = min(b, r + STRIDE - 1)
            for rn, rec, _dec, _at in self.records(r, last):
                line = row_line(f"R{rn}" + ("（見出し）" if rn == 1 else ""), rec.values, width)
                if len(line) - skip > max_chars:
                    if body:
                        stop = rn
                        break
                    piece = line[skip:skip + max_chars]
                    body.append(piece + cut_note(skip, skip + max_chars, len(line)))
                    cut = (rn, skip + max_chars)
                    stop = rn
                    break
                if skip:
                    line = line[skip:] + f"〔ここまでで R{rn} は全部〕"
                    skip = 0
                if body and size + len(line) + 1 > max_chars:
                    stop = rn
                    break
                body.append(line)
                size += len(line) + 1
            if stop is not None:
                break
            r = last + 1
        lines = list(head) + body
        end_of_file = stop is None and b == n
        if end_of_file:
            add = "| R" + str(n + 1) + "（足す行） | " + " | ".join(f"［R{n + 1}.C{c}？］" for c in range(1, width + 1)) + " |"
            lines.append(add)
        if stop is None:
            return Page("\n".join(lines), f"R{a}〜R{b}（全 {comma(n)} 行）", None, "", body)
        done_to = stop - 1 if cut is None else stop
        nxt: dict = {"cursor": f"rows:{cut[0]}:{cut[1]}" + (f":{b}" if b != n else "") if cut else f"rows:{stop}" + (f":0:{b}" if b != n else "")}
        cov = f"R{a}〜R{done_to}" + ("（の途中まで）" if cut else "") + f"（全 {comma(n)} 行）"
        return Page("\n".join(lines), cov, nxt, f"{comma(b - stop + 1)} 行", body)

    # ------------------------------------------------------------------ 欄
    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        n = self.idx.count
        a, b, _skip = _row_args(rng, n, default_from=2 if n > 1 else 1)
        out: list[dict] = []
        width = self.width
        r = a
        blanks = 0
        looked = 0
        stop = None
        while r <= b and stop is None:
            last = min(b, r + STRIDE - 1)
            for rn, rec, _d, _at in self.records(r, last):
                vals = rec.values
                looked += 1
                row_ctx = None
                for c in range(1, width + 1):
                    v = vals[c - 1] if c <= len(vals) else None
                    if v in (None, ""):
                        blanks += 1
                    elif only_empty:
                        continue
                    if row_ctx is None:
                        row_ctx = row_line(f"R{rn}", [_short_value(x)[0] for x in vals], width)
                    name = self._label(c)
                    shown, whole = _short_value(v)
                    out.append({"id": f"R{rn}.C{c}", "kind": "text", "label": f"{name}（{rn}行目）", "value": shown, "row": rn, "col": c,
                                "column": name, "provenance": "deterministic", "context": row_ctx,
                                **({"value_chars": whole, "value_cut": f"値は {whole:,} 字あり、頭の {FIELD_VALUE_CHARS:,} 字だけを示しています。全部は read_window で読めます"}
                                   if whole else {})})
                if len(out) >= limit:
                    stop = rn + 1
                    break
            r = last + 1
        res = {"fields": out, "total_rows": n, "total_columns": width, "looked_rows": f"R{a}〜R{(stop - 1) if stop else b}",
               "blank_cells_in_range": blanks, "next": None}
        if stop and stop <= b:
            res["next"] = {"rows": f"{stop}-{b}"}
            res["remaining"] = f"{comma(b - stop + 1)} 行"
        elif b == n:
            for c in range(1, max(width, 1) + 1):
                res["fields"].append({"id": f"R{n + 1}.C{c}", "kind": "text", "label": f"{self._label(c)}（足す行 {n + 1}行目）", "value": None,
                                      "row": n + 1, "col": c, "column": self._label(c), "agent": True, "provenance": "deterministic",
                                      "why": "新しい行（足すかどうかは、エージェントが決める。書けばファイルの終わりに行を足す）"})
        return res

    # ------------------------------------------------------------------ 書く
    def fill(self, values: dict, out_path: str, force: bool, overwrite: bool, said: dict | None) -> dict:
        want: dict[tuple[int, int], tuple[str, str]] = {}
        problems: list[str] = []
        for key, v in values.items():
            try:
                r, c = self.resolve(key)
            except CsvError as e:
                problems.append(str(e))
                continue
            if (r, c) in want:
                problems.append(f"{key}: 同じマス（R{r}.C{c}）に2回書こうとしています")
                continue
            text = to_text(v)
            if self.kind == "tsv" and "\t" in text and not self.st.quote_all and not self.st.quoted_any:
                problems.append(f"{key}: タブ区切りのファイルなので、値にタブは書けません")
                continue
            want[(r, c)] = (text, str(key))
        n = self.idx.count
        new_rows = sorted({r for r, _c in want if r > n})
        if new_rows and new_rows != list(range(n + 1, n + 1 + len(new_rows))):
            problems.append(f"新しい行は、今ある {n} 行の次（R{n + 1}）から続けて足します（行を飛ばせません：" + "、".join(f"R{r}" for r in new_rows) + "）")
        if problems:
            raise CsvError("\n".join(problems))
        edits: list[tuple[int, int, bytes]] = []
        patches: list[CsvPatch] = []
        warnings: list[str] = []
        touched: dict[int, tuple] = {}  # 行 → (元の始めのバイト, 元の終わりのバイト（改行を含む）)
        before: dict[int, tuple[list, str]] = {}
        was: dict[str, str | None] = {}  # 書いたマスの、書く前の値（もとからあった字が残っているかを見る）
        try:
            for r in sorted({r for r, _c in want if r <= n}):
                rn, rec, dec, at = self.records(r, r)[0]
                xa = at + dec.byte_of(rec.a)
                xz = at + dec.byte_of(rec.b + len(rec.terminator))
                touched[r] = (xa, xz)
                before[r] = ([dec.text[f_.a:f_.b] for f_ in rec.fields], rec.terminator)
                tail = []
                for c in sorted(c for (rr, c) in want if rr == r):
                    text, key = want[(r, c)]
                    if c <= len(rec.fields):
                        fd = rec.fields[c - 1]
                        was[key] = fd.value
                        edits.append((at + dec.byte_of(fd.a), at + dec.byte_of(fd.b), dec.encode(quote(text, self, fd.quoted))))
                    else:
                        tail.append(c)
                    patches.append(CsvPatch(f"R{r}.C{c}", r, c, text, key))
                if tail:
                    last = len(rec.fields)
                    hi = max(tail)
                    cells = [quote(want[(r, c)][0], self, False) if (r, c) in want else ('""' if self.quote_empty else "")
                             for c in range(last + 1, hi + 1)]
                    pos = at + dec.byte_of(rec.b)
                    edits.append((pos, pos, dec.encode(self.delim + self.delim.join(cells))))
                    warnings.append(f"R{r}：もとの行は {last} マスなので、区切りを足して C{hi} まで書きました")
            if new_rows:
                width = max(self.width, max(c for (_r, c) in want))
                lines = []
                for r in new_rows:
                    cells = []
                    for c in range(1, width + 1):
                        if (r, c) in want:
                            cells.append(quote(want[(r, c)][0], self, False))
                            patches.append(CsvPatch(f"R{r}.C{c}", r, c, want[(r, c)][0], want[(r, c)][1]))
                        else:
                            cells.append('""' if self.quote_empty else "")
                    lines.append(self.delim.join(cells))
                nl = self.idx.newline or "\r\n"
                ewn = self.idx.ends_with_newline
                add = (nl if n and not ewn else "") + nl.join(lines) + (nl if ewn or not n else "")
                d0 = self._slice_dec("")
                edits.append((self.src.size, self.src.size, d0.encode(add)))
                if self.width and width > self.width:
                    warnings.append(f"新しい行は {width} マスです（もとの表は {self.width} 列）")
        except EncodingError as e:
            raise CsvError(str(e)) from None
        if self.dec.ascii_only and any(not p.value.isascii() for p in patches):
            warnings.append("もとのファイルは ASCII の字だけで文字コードが決まらないので、UTF-8（BOM なし）で書きました")
        patches.sort(key=lambda p: (p.row, p.col))
        risky = [{"id": p.key, "value": p.value, "safe_form": "'" + p.value} for p in patches if formula_like(p.value)]
        warnings += [formula_note(x["id"], x["value"]) for x in risky]
        dst = Path(out_path).expanduser().resolve()
        tmp = temp_beside(dst)
        try:
            ranges, h_old, h_new = stream_patch(self.src, edits, tmp)
            unexpected, back = self._prove_csv(tmp, ranges, edits, patches, touched, before, want, new_rows)
            result = {"kind": self.kind, "out_path": str(dst), "original_sha256": h_old, "new_sha256": h_new,
                      "proof_ok": not unexpected, "unexpected": unexpected[:20], "structure_errors": [], "warnings": warnings,
                      "proof": (f"行 {comma(n)} のうち、書いた行 {len(touched)} 行（足した行 {len(new_rows)} 行）のほかは 1 バイトも変わっていません"
                                f"（変えていない {len(ranges.items)} の範囲・{comma(ranges.total)} バイトのハッシュを、書いた後のファイルを読み直して比べました）。"
                                "書いた行は、書いたマスのほかの字・引用符・改行が同じ"
                                + ("" if not unexpected else f"（想定外の変更 {len(unexpected)} 件）")),
                      "read_back_by": "offset", "read_back_ok": not back, "read_back_differences": back[:30],
                      "cells": {p.key: p.id for p in patches if p.key != p.id}, "streamed": True}
            if new_rows:
                result["appended_rows"] = len(new_rows)
            if risky:
                result["formula_like"] = risky
            _finish_fill(result, {p.key: p.value for p in patches}, {p.key: f"{self._label(p.col)}（{p.row}行目）" for p in patches}, said,
                         tmp, dst, self.path, overwrite, f"list_fields に rows を渡すと、その範囲の空のマスが分かります（全 {comma(n)} 行）", before=was)
            return result
        finally:
            tmp.unlink(missing_ok=True)

    def _prove_csv(self, tmp: Path, ranges: ByteRanges, edits, patches, touched, before, want, new_rows):
        """書いた行だけを比べ、書き換えなかった範囲はハッシュで比べる。読み戻しは、書いたファイルの位置を読み直す。"""
        unexpected = verify_ranges(tmp, ranges)
        back: list[dict] = []
        new_src = Source(str(tmp))
        enc = self.dec.encoding
        marks = sorted(edits, key=lambda e: (e[0], e[1]))
        k_m = 0
        shift = 0
        try:
            for r, (xa, xz) in sorted(touched.items()):
                while k_m < len(marks) and marks[k_m][1] <= xa and marks[k_m][0] < xa:
                    shift += len(marks[k_m][2]) - (marks[k_m][1] - marks[k_m][0])
                    k_m += 1
                inner = 0
                k2 = k_m
                while k2 < len(marks) and marks[k2][1] <= xz and marks[k2][0] < xz + 1 and marks[k2][0] != self.src.size:
                    inner += len(marks[k2][2]) - (marks[k2][1] - marks[k2][0])
                    k2 += 1
                k_m = k2
                a2, b2 = xa + shift, xz + shift + inner
                shift += inner
                # 行の中の編集で行の終わりの位置が動く（行の終わりに足した区切りも、行の中）
                text = new_src.read(a2, b2).decode(enc)
                recs, probs = parse(text, self.delim)
                if len(recs) != 1:
                    unexpected.append(f"{r} 行目：書いた後の行が {len(recs)} 行に読めます")
                    continue
                rec = recs[0]
                old_cells, old_term = before[r]
                if rec.terminator != old_term:
                    unexpected.append(f"{r} 行目の改行が変わっています（{NEWLINE_JA[old_term]} → {NEWLINE_JA[rec.terminator]}）")
                if probs:
                    unexpected.append(f"{r} 行目：書いた後の行に、書き方の問題があります：{probs[0]}")
                for c in range(1, max(len(old_cells), len(rec.fields)) + 1):
                    fb = rec.fields[c - 1] if c <= len(rec.fields) else None
                    if (r, c) in want:
                        if fb is None or fb.value != want[(r, c)][0]:
                            back.append({"id": f"R{r}.C{c}", "written": want[(r, c)][0], "read_back": fb.value if fb else None})
                        continue
                    if c <= len(old_cells):
                        if fb is None or text[fb.a:fb.b] != old_cells[c - 1]:
                            unexpected.append(f"R{r}.C{c}：書いていないマスの字が変わっています")
                    elif fb is not None and fb.value != "":
                        unexpected.append(f"R{r}.C{c}：書いていないマスに字が入っています")
                # 別の読み手（Python の標準の csv）でも、書いた値のとおりに読めるか
                rows = list(_csv.reader(io.StringIO(text, newline=""), delimiter=self.delim))
                if rows and rec.fields and not rec.a == rec.b:
                    for (rr, c), (val, _k) in want.items():
                        if rr == r and not (c <= len(rows[0]) and rows[0][c - 1] == val):
                            back.append({"id": f"R{r}.C{c}", "written": val, "read_back": rows[0][c - 1] if c <= len(rows[0]) else None,
                                         "why": "別の読み方（csv）での読み戻し"})
            if new_rows:
                added = self.src.size + sum(len(new) - (y - x) for x, y, new in marks)
                tail = sum(len(new) for x, y, new in marks if x == y == self.src.size)
                text = new_src.read(added - tail, added).decode(enc)
                recs, _p = parse(text, self.delim)
                recs = [q for q in recs if q.a != q.b or q.terminator == ""] if text.startswith(("\r", "\n")) else recs
                if len(recs) != len(new_rows):
                    unexpected.append(f"足した行の数が合いません（{len(recs)} 行・書いた行 {len(new_rows)} 行）")
                else:
                    for k, r in enumerate(new_rows):
                        for (rr, c), (val, _kk) in want.items():
                            got = recs[k].values
                            if rr == r and not (c <= len(got) and got[c - 1] == val):
                                back.append({"id": f"R{r}.C{c}", "written": val, "read_back": got[c - 1] if c <= len(got) else None})
        finally:
            new_src.close()
        return unexpected, back

    # ------------------------------------------------------------------ 段階
    def grade(self) -> dict:
        from ..grade import Grade, Reason, _final, _locked
        from ..openkinds import KIND_JA, delimiter_reason, encoding_reason

        direct = [Reason("A", "データそのもの", "行と列のデータに、そのまま書けます")]
        if self.dec.binary:
            return _locked("読めない", "文字のファイルではありません（どの文字コードとしても読めず、NUL のバイトがあります）", "読めないので、書けません")
        reasons = [Reason("S", "データそのもの", f"表のデータ（{KIND_JA[self.kind]}・{self.dec.label}・{self.idx.count:,} 行）で、変換しなくても機械が読めます"
                          "（ファイルの全部の字と行の区切りを調べました）")]
        if self.st.first_problems:
            reasons.append(Reason("S", "書き方の問題", "RFC 4180 に合わない書き方があります（どう読んだかは窓に書いてあります）", self.st.first_problems[:5]))
        enc, wenc = encoding_reason(self.dec)
        sep, wsep = delimiter_reason(self.format_info())
        more = [r for r in (enc, sep) if r is not None]
        g = (_final(reasons + more, [w for w in (wenc, wsep) if w is not None] + direct, base="S") if more
             else Grade("S", reasons, write="A", write_reasons=direct))
        g.deferred = []  # 文字コード・行の区切り・マスの数はファイルの全部で調べたので、残した確かめはない
        return g

    # ------------------------------------------------------------------ 検査
    def audit_begin(self) -> dict:
        return {"next_row": 1, "cells": 0, "chars": 0, "problems": [], "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        """1 区切り（STRIDE×256 行）：別の読み手（Python の csv）で、窓の読み手と同じ値かを確かめる。全部の行が終わったら書く試し。"""
        if self.recoded:
            return self._audit_whole(state)
        n = self.idx.count
        r = state["next_row"]
        if r <= n:
            last = min(n, r + STRIDE * 256 - 1)
            a, b, base = self.idx.bucket(r, last)
            text = self.text_of(a, b)
            mine = [x.values for x in parse(text, self.delim)[0]]
            other = list(_csv.reader(io.StringIO(text, newline=""), delimiter=self.delim, strict=False))
            # csv の読み手は、字のない行を [] にする
            k = 0
            for i, vals in enumerate(mine):
                rn = base + i
                oth = other[k] if k < len(other) else None
                k += 1
                if vals == [""] and oth == []:
                    continue
                if rn >= r and vals != oth and len(state["problems"]) < 30:
                    state["problems"].append(f"R{rn}：窓の読み手と別の読み手で値が違います")
            state["cells"] += sum(len(v) for v in mine)
            state["chars"] += len(text)
            state["next_row"] = base + len(mine)
            limits.progress("行を別の読み手で確かめています", state["next_row"], n)
            return state
        state["checks"] = _text_checks("読み落とし（表の字）", state, f"{n:,} 行・マス {state['cells']:,} か所のすべてを、窓とは別の読み方（Python の標準の csv の読み手）で読み、同じ値であることを確かめました")
        # 書く試し：いちばん後ろの行と頭に近い行で、引用符・カンマ・改行を含む値を書き、証明と読み戻しを確かめる
        state["checks"].append(self._write_trial())
        state["done"] = True
        return state

    def _write_trial(self) -> dict:
        import tempfile

        n = self.idx.count
        sample = {}
        for k, r in enumerate(sorted({min(2, n), max(n // 2, 1), n})):
            if r >= 1:
                sample[f"R{r}.C1"] = 'カンマ,と"引用符"' if k == 0 else "改行\nあり" if k == 1 else "テスト"
        try:
            for v in sample.values():
                v.encode(self.dec.encoding)
        except UnicodeEncodeError:
            sample = {k: "test" for k in sample}
        with tempfile.TemporaryDirectory(prefix="todoku-audit-") as td:
            try:
                out = str(Path(td) / f"t{Path(self.name).suffix}")
                res = self.fill(sample, out, False, True, None)
                ok = bool(res.get("proof_ok") and res.get("read_back_ok"))
                return {"name": "書き込みの試し", "ok": ok, "agent": False,
                        "detail": f"頭・真ん中・終わりの行（{'、'.join(sample)}）に試しの値を書き、{res.get('proof', '')}。すべてのマスは同じ書き方で書きます"
                                  "（マスの数が多いので、すべてのマスに試しを書くことはしていません）",
                        "items": (res.get("unexpected") or []) + [str(x) for x in res.get("read_back_differences") or []]}
            except (CsvError, ValueError) as e:
                return {"name": "書き込みの試し", "ok": False, "agent": False, "detail": f"書けませんでした：{e}", "items": []}


FIELD_VALUE_CHARS = 2000  # 欄の一覧に入れる値の字数。とても長いマスで、一覧が何百 MB にならない（切ったときは、切ったと示す）


def _short_value(v) -> tuple:
    """(一覧に出す値, 切ったときの全体の字数。切らなければ 0)。"""
    if isinstance(v, str) and len(v) > FIELD_VALUE_CHARS:
        return v[:FIELD_VALUE_CHARS], len(v)
    return v, 0


def _text_checks(name: str, state: dict, detail: str) -> list[dict]:
    return [{"name": name, "ok": not state["problems"], "agent": False, "detail": detail, "items": list(state["problems"])}]


def _row_args(rng: dict, n: int, default_from: int = 1) -> tuple[int, int, int]:
    """頼まれた行の範囲：(始め, 終わり, 長い行の中の字の位置)。cursor は「rows:行」か「rows:行:字の位置[:終わり]」。"""
    cursor = rng.get("cursor")
    if cursor:
        m = re.fullmatch(r"rows:(\d+)(?::(\d+))?(?::(\d+))?", str(cursor))
        if not m:
            raise RangeError(f"cursor「{cursor}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
        a = int(m[1])
        b = int(m[3]) if m[3] else n
        return a, min(b, n), int(m[2] or 0)
    spec = rng.get("rows") if rng.get("rows") is not None else rng.get("records")
    a, b = parse_range(spec, "行")
    a = default_from if a is None else a
    b = n if b is None else min(b, n)
    if a < 1:
        raise RangeError("行は 1 から数えます")
    if a > max(n, 1):
        raise RangeError(f"{a} 行目はありません（全 {n:,} 行）")
    return a, b, 0


def _finish_fill(result: dict, written: dict, labels: dict, said: dict | None, tmp: Path, dst: Path, src_path: str, overwrite: bool, rest_note: str,
                 before: dict | None = None, number_changed: dict | None = None) -> None:
    """大きな書類の書き込みの終わり（どの形式でも）：ふつうの道の終わりと同じことを伝える。
    before：書いた所の、書く前の字や値（id → 値。空の所は None）。written（書いた後のファイルを読み直した字）と比べて、もとからあった字が
    元のままでない所を replaced_existing に入れる（api._replaced_values：ふつうの道と同じ決め方）。手で行う所（hand_only）も入れる。"""
    from .. import api

    kind = result.get("kind")
    gone = api._replaced_values(before or {}, written, written, number_changed, lambda x: api._writer_bool(kind, x), api._is_sheet(kind))
    if gone:
        result["replaced_existing"] = gone
    api._add_hand_only(result)
    told = api._to_tell([], written, result, labels, set(), {}, said or {})
    result.update(told)
    result["empty_fields_note"] = rest_note
    result["say_to_user"] = (result.get("say_to_user") or "") + f"\n（大きな書類なので、まだ空いている欄は全部を数えていません。{rest_note}）"
    api._publish(tmp, dst, Path(src_path).expanduser().resolve(), overwrite, result)


# ====================================================================== 行の読み取り（JSON Lines・.txt）
class LineDoc(TextDoc):
    @property
    def idx(self) -> TextIndex:
        if self._idx is None:
            self._idx = index_lines(self.src, self.start, self.dec.encoding)
        return self._idx

    def line_spans(self, first: int, last: int) -> list[tuple[int, int, int, str]]:
        """[(行の番号, 中身の始め, 中身の終わり, 改行)]（バイト）。行の中身は読まない（長い行でもメモリは一定）。"""
        idx = self.idx
        if not idx.count or first > idx.count:
            return []
        last = min(last, idx.count)
        a, b, base = idx.bucket(first, last)
        enc = self.dec.encoding
        out = []
        unit = 2 if enc.startswith("utf-16") else 4 if enc.startswith("utf-32") else 1
        from .textstream import Window, tokens

        win = Window(self.src, a, unit)
        win.origin = self.start
        pat = re.compile(b"|".join(re.escape(x) for x in tokens(enc, "\r\n", "\n", "\r")))
        pos = a
        n = base
        while n <= last and pos < b:
            hit = win.search(pat, pos) if pos < self.src.size else None
            if hit is None or hit[0] >= b:
                if n >= first:
                    out.append((n, pos, b, ""))
                break
            x, y, t = hit
            if n >= first:
                out.append((n, pos, x, t.decode(enc)))
            pos = y
            n += 1
        return out


# ====================================================================== JSON Lines
class JsonlDoc(LineDoc):
    def overview(self) -> dict:
        return {"format": {"encoding": self.shown.encoding, "encoding_ja": self.shown.label, "bom": bool(self.shown.bom), **self.shown.info(),
                           "encoding_notes": self.shown.note_lines()},
                "records": self.idx.count, "ranges": "records（行の範囲）", "bytes": self.file_bytes}

    def overview_text(self) -> str:
        n = self.idx.count
        return "\n".join([f"# {self.name}", "", self.source_line(),
                          f"- JSON Lines・{n:,} 行。文字コード {self.shown.label}",
                          f"- 大きいので、ここでは全体を出していません。read_window に records: \"1-500\" のように行の範囲を渡して読みます"
                          f"（cursor をたどれば L1〜L{n} のすべてを読めます）。list_fields も records で、その範囲の値を出します。",
                          "- 値の id は「L行/名前」（L3/name。L3 は 3 行目の全体）。書き込みは、書いた値の字だけを変え、ほかのバイトはそのままファイルを写します。",
                          *self.shown.note_lines()]) + "\n"

    def _parse_line(self, n: int, data: bytes):
        from ..datafile.jsonfile import JsonError, JsonFile

        f = JsonFile(data, self.name, "jsonl", self.dec.encoding)
        for d in f.docs:
            d.line = n
        return f, JsonError

    def _shown(self, n: int, data: bytes) -> tuple[str, list]:
        """1 行の窓の字と (id, 値)。JSON として読めない行は、字のまま出して、読めないことを示す。"""
        from ..datafile.jsonfile import JsonError
        from ..datafile.window import literal_text

        text = data.decode(self.dec.encoding)
        if not text.strip():
            return "", []
        try:
            f, _ = self._parse_line(n, data)
        except JsonError as e:
            return f"- `L{n}`〔JSON として読めません：{e}〕 {text}", []
        leaves = [(i, nd, literal_text(nd, f.text)) for i, nd, _d in f.leaves()]
        return f"- `L{n}` " + " ／ ".join(f"`{i.split('/', 1)[1] if '/' in i else i}` ＝ {lit}" for i, _nd, lit in leaves), leaves

    def read(self, rng: dict, max_chars: int) -> Page:
        n = self.idx.count
        a, b, skip = _rec_args(rng, n)
        body: list[str] = []
        size = 0
        r = a
        stop = None
        cut = None
        while r <= b and stop is None:
            last = min(b, r + STRIDE - 1)
            for ln, x, y, _t in self.line_spans(r, last):
                line, _ = self._shown(ln, self.src.read(x, y))
                if len(line) - skip > max_chars:
                    if body:
                        stop = ln
                        break
                    body.append(line[skip:skip + max_chars] + cut_note(skip, skip + max_chars, len(line)))
                    cut = (ln, skip + max_chars)
                    stop = ln
                    break
                if skip:
                    line = line[skip:]
                    skip = 0
                if body and size + len(line) + 1 > max_chars:
                    stop = ln
                    break
                if line:
                    body.append(line)
                    size += len(line) + 1
            r = last + 1
        if stop is None:
            return Page("\n".join(body), f"L{a}〜L{b}（全 {comma(n)} 行）", None, "", body)
        return Page("\n".join(body), f"L{a}〜L{stop - (0 if cut else 1)}（全 {comma(n)} 行）",
                    {"cursor": f"records:{cut[0]}:{cut[1]}" + (f":{b}" if b != n else "") if cut else f"records:{stop}" + (f":0:{b}" if b != n else "")},
                    f"{comma(b - stop + 1)} 行", body)

    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        from ..datafile.jsonfile import unesc_token

        n = self.idx.count
        a, b, _s = _rec_args(rng, n)
        out: list[dict] = []
        stop = None
        r = a
        blanks = 0
        while r <= b and stop is None:
            last = min(b, r + STRIDE - 1)
            for ln, x, y, _t in self.line_spans(r, last):
                _line, leaves = self._shown(ln, self.src.read(x, y))
                for i, nd, lit in leaves:
                    if nd.value in (None, ""):
                        blanks += 1
                    elif only_empty:
                        continue
                    label = unesc_token(i.rsplit("/", 1)[-1]) if "/" in i else i
                    kind = {"true": "checkbox", "false": "checkbox", "number": "number"}.get(nd.kind, "text")
                    out.append({"id": i, "kind": kind, "label": label, "value": nd.value, "json_type": nd.kind, "provenance": "deterministic",
                                "context": f"{i} ＝ {lit}"})
                if len(out) >= limit:
                    stop = ln + 1
                    break
            r = last + 1
        res = {"fields": out, "total_records": n, "looked_rows": f"L{a}〜L{(stop - 1) if stop else b}", "blank_values_in_range": blanks, "next": None}
        if stop and stop <= b:
            res["next"] = {"records": f"{stop}-{b}"}
            res["remaining"] = f"{comma(b - stop + 1)} 行"
        return res

    def fill(self, values: dict, out_path: str, force: bool, overwrite: bool, said: dict | None) -> dict:
        from ..datafile.jsonfile import JsonError, JsonFile, _plain, fill as jfill, prove as jprove

        by_line: dict[int, dict] = {}
        problems = []
        for key, v in values.items():
            m = re.fullmatch(r"[Ll](\d+)(/.*)?", str(key).strip())
            if not m:
                problems.append(f"「{key}」は値の id ではありません（L行/名前 の形。例 L3/name）")
                continue
            by_line.setdefault(int(m[1]), {})[key] = (m[2] or "", v)
        n = self.idx.count
        edits = []
        patches = []
        notes: list[str] = []
        spans = {}
        was: dict = {}  # 書いた値の、書く前の値（もとからあった値が残っているかを見る）
        for ln, items in sorted(by_line.items()):
            got = self.line_spans(ln, ln)
            if not got or ln > n:
                problems.append(f"{ln} 行目に JSON の値がありません（全 {n:,} 行）")
                continue
            _n, x, y, _t = got[0]
            data = self.src.read(x, y)
            try:
                f = JsonFile(data, self.name, "jsonl", self.dec.encoding)
                for key, (ptr, _v) in items.items():
                    try:
                        node = f.find("L1" + ptr)[0]
                    except ValueError:
                        continue  # まだない値（足す所）
                    if node is not None:
                        was[key] = _plain(node)
                res = jfill(f, {"L1" + p: val for p, val in items.values()})
            except JsonError as e:
                problems.append(f"{ln} 行目：{e}")
                continue
            pr = jprove(data, res.blob, {p.id: p.value for p in res.patches}, self.name, "jsonl", self.dec.encoding)
            if not pr.ok:
                problems += [f"{ln} 行目：{u}" for u in pr.unexpected + pr.lint_errors]
            rev = {"L1" + ptr: key for key, (ptr, _v) in items.items()}
            for p in res.patches:
                patches.append({"key": rev.get(p.key, p.key), "id": f"L{ln}" + p.id[2:], "value": p.value, "text": p.text, "added": p.added})
            notes += res.warnings
            edits.append((x, y, res.blob))
            spans[ln] = (x, y, res.blob)
        if problems:
            raise JsonError("\n".join(problems))
        dst = Path(out_path).expanduser().resolve()
        tmp = temp_beside(dst)
        try:
            ranges, h_old, h_new = stream_patch(self.src, edits, tmp)
            unexpected = verify_ranges(tmp, ranges)
            back = []
            new_src = Source(str(tmp))
            try:
                shift = 0
                for ln, (x, y, new) in sorted(spans.items()):
                    data = new_src.read(x + shift, x + shift + len(new))
                    if data != new:
                        unexpected.append(f"{ln} 行目：書いた後のファイルの字が、書いた行と違います")
                    f = JsonFile(data, self.name, "jsonl", self.dec.encoding)
                    for p in patches:
                        m = re.fullmatch(r"L(\d+)(/.*)?", p["id"])
                        if int(m[1]) != ln:
                            continue
                        from ..datafile.jsonfile import _plain

                        node = f.find("L1" + (m[2] or ""))[0]
                        got = _plain(node) if node is not None else None
                        if got != p["value"]:
                            back.append({"id": p["id"], "written": p["value"], "read_back": got})
                        # 別の読み手（Python の json）でも同じか
                        obj = json.loads(data.decode(self.dec.encoding))
                        for t in (m[2] or "").split("/")[1:]:
                            t = t.replace("~1", "/").replace("~0", "~")
                            obj = obj[int(t)] if isinstance(obj, list) else obj[t]
                        if obj != p["value"]:
                            back.append({"id": p["id"], "written": p["value"], "read_back": obj, "why": "別の読み方（json）での読み戻し"})
                    shift += len(new) - (y - x)
            finally:
                new_src.close()
            result = {"kind": "jsonl", "out_path": str(dst), "original_sha256": h_old, "new_sha256": h_new, "proof_ok": not unexpected,
                      "unexpected": unexpected[:20], "structure_errors": [], "warnings": notes, "streamed": True,
                      "proof": (f"全 {comma(n)} 行のうち、書いた {len(spans)} 行のほかは 1 バイトも変わっていません（変えていない {len(ranges.items)} の範囲・"
                                f"{comma(ranges.total)} バイトのハッシュを、書いた後のファイルを読み直して比べました）。書いた行は、書いた値の外の字（空白を含む）が同じ"),
                      "read_back_by": "offset", "read_back_ok": not back, "read_back_differences": back[:30],
                      "written_as": {p["key"]: p["text"] for p in patches}}
            added = {p["key"]: p["id"] for p in patches if p["added"]}
            if added:
                result["added"] = added
            _finish_fill(result, {p["key"]: p["value"] for p in patches}, {p["key"]: p["id"].rsplit("/", 1)[-1] for p in patches}, said, tmp, dst,
                         self.path, overwrite, f"list_fields に records を渡すと、その範囲の値が分かります（全 {comma(n)} 行）", before=was)
            return result
        finally:
            tmp.unlink(missing_ok=True)

    def grade(self) -> dict:
        from ..grade import Grade, Reason, _final
        from ..openkinds import encoding_reason

        direct = [Reason("A", "データそのもの", "名前と値のデータに、そのまま書けます")]
        reasons = [Reason("S", "データそのもの", f"構造のあるデータ（JSON Lines・{self.dec.label}・{self.idx.count:,} 行）で、変換しなくても機械が読めます")]
        enc, wenc = encoding_reason(self.dec)
        g = _final(reasons + [enc], [wenc] + direct, base="S") if enc is not None else Grade("S", reasons, write="A", write_reasons=direct)
        g.deferred = ["未確認：全行が JSON として読めるか・同じ名前が 2 つある所がないか（audit_document で全行を確かめます）"]
        return g

    def audit_begin(self) -> dict:
        return {"next_row": 1, "values": 0, "problems": [], "dups": [], "done": False, "checks": []}

    def audit_step(self, state: dict) -> dict:
        if self.recoded:
            return self._audit_whole(state)
        from ..datafile.jsonfile import JsonError, JsonFile

        n = self.idx.count
        r = state["next_row"]
        if r <= n:
            last = min(n, r + STRIDE * 32 - 1)
            for ln, x, y, _t in self.line_spans(r, last):
                data = self.src.read(x, y)
                if not data.strip():
                    continue
                try:
                    f = JsonFile(data, self.name, "jsonl", self.dec.encoding)
                    other = json.loads(data.decode(self.dec.encoding))
                    if len(list(f.leaves())) != _count_leaves(other) and not f.duplicates and len(state["problems"]) < 30:
                        state["problems"].append(f"L{ln}：値の数が別の読み方（Python の json）と違います")
                    state["values"] += sum(1 for _ in f.leaves())
                    if f.duplicates and len(state["dups"]) < 30:
                        state["dups"] += [f"L{ln}{d[2:]}" for d in f.duplicates]
                except (JsonError, ValueError) as e:
                    if len(state["problems"]) < 30:
                        state["problems"].append(f"L{ln}：JSON として読めません（{e}）")
            state["next_row"] = last + 1
            limits.progress("行を確かめています", state["next_row"], n)
            return state
        state["checks"] = _text_checks("読み落とし（JSON Lines の値）", state, f"{n:,} 行・値 {state['values']:,} 個のすべてを、窓とは別の読み方（Python の json）で読み直しました")
        if state["dups"]:
            state["checks"].append({"name": "同じ名前", "ok": False, "agent": False, "detail": "オブジェクトに同じ名前が 2 つ以上ある所があります", "items": state["dups"]})
        state["done"] = True
        return state


def _count_leaves(o) -> int:
    if isinstance(o, dict):
        return sum(_count_leaves(v) for v in o.values())
    if isinstance(o, list):
        return sum(_count_leaves(v) for v in o)
    return 1


def _rec_args(rng: dict, n: int) -> tuple[int, int, int]:
    cursor = rng.get("cursor")
    if cursor:
        m = re.fullmatch(r"records:(\d+)(?::(\d+))?(?::(\d+))?", str(cursor))
        if not m:
            raise RangeError(f"cursor「{cursor}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
        return int(m[1]), min(int(m[3]) if m[3] else n, n), int(m[2] or 0)
    spec = next((rng[k] for k in ("records", "rows", "paragraphs") if rng.get(k) is not None), None)
    a, b = parse_range(spec, "行")
    a = 1 if a is None else a
    b = n if b is None else min(b, n)
    if a > max(n, 1):
        raise RangeError(f"{a} 行目はありません（全 {n:,} 行）")
    return a, b, 0


# ====================================================================== .txt
SEG_CHARS = 50_000  # とても長い行を分ける単位（字）。id は字の位置で決まるので、この単位は id には関わらない
SCAN_SEGS = 40  # list_fields が 1 回に調べる区切りの数（200 万字）。残りは next の cursor で


class HugeLine:
    """とても長い行：字の位置 ↔ バイトの位置の控え（およそ 1 MB ごと）。字の範囲をデコードするのに、行の頭から読み直さない。"""

    def __init__(self, doc: "TxtLazy", x: int, y: int):
        self.doc = doc
        self.x, self.y = x, y
        enc = doc.dec.encoding
        dec = codecs.getincrementaldecoder(enc)("strict")
        self.marks: list[tuple[int, int]] = [(0, x)]  # (字の位置, バイトの位置)。デコーダが途中の字を持っていない所だけ
        chars = 0
        pos = x
        for c in doc.src.chunks(x, y):
            t = dec.decode(c)
            chars += len(t)
            pos += len(c)
            pending = dec.getstate()[0]
            if not pending:
                self.marks.append((chars, pos))
        chars += len(dec.decode(b"", True))
        self.chars = chars
        self._keys = [m[0] for m in self.marks]

    def text(self, a: int, b: int) -> str:
        """字の範囲 [a, b)。"""
        i = bisect.bisect_right(self._keys, a) - 1
        ca, pos = self.marks[i]
        dec = codecs.getincrementaldecoder(self.doc.dec.encoding)("strict")
        out: list[str] = []
        have = ca
        for c in self.doc.src.chunks(pos, self.y):
            t = dec.decode(c)
            out.append(t)
            have += len(t)
            if have >= b:
                break
        s = "".join(out)
        return s[a - ca:b - ca]

    def byte_of(self, a: int) -> int:
        i = bisect.bisect_right(self._keys, a) - 1
        ca, pos = self.marks[i]
        if ca == a:
            return pos
        return pos + len(self.text(ca, a).encode(self.doc.dec.encoding))


class TxtLazy(LineDoc):
    def __init__(self, path: str, name: str, kind: str, encoding: str | None):
        super().__init__(path, name, kind, encoding)
        self._huge: dict[int, HugeLine] = {}
        from ..book import _text_kind, content_note

        # 形式を中身で決めたファイル（book.content_kind）：窓・段階にそのことを残す
        self.kind_note = content_note("txt", name) if _text_kind(self.src.read(0, 4096), name) is None else ""

    def _binary_check(self) -> None:
        if self.dec.binary:
            from ..txt.document import TxtError

            raise TxtError("文字のファイルではありません（どの文字コードとしても読めず、NUL のバイトがあります）")

    def huge(self, n: int, x: int, y: int) -> HugeLine | None:
        """この行が、1 回に出せないほど長いなら、その控え。"""
        if y - x <= SEG_CHARS:
            return None
        h = self._huge.get(n)
        if h is None:
            h = HugeLine(self, x, y)
            if len(self._huge) >= 8:
                self._huge.pop(next(iter(self._huge)))
            self._huge[n] = h
        return h if h.chars > SEG_CHARS else None

    # ------------------------------------------------------------------ 範囲の文書
    def _slice_doc(self, n: int, data: bytes, char_base: int | None = None):
        """行の範囲（や、長い行の一部）を、id が元のファイルの行の番号になる TxtDocument にする。"""
        from ..txt.document import TxtDocument

        doc = TxtDocument(data, self.name, self.dec.encoding)
        doc.para_base = n - 1
        for k, ln in enumerate(doc.lines):
            ln.number = n + k
        if char_base is not None:
            doc.place_abs = char_base
        return doc

    def _window_lines(self, doc) -> tuple[list[str], list[dict]]:
        from ..txt.window import build, line_text

        win = build(doc)
        return [line_text(x["n"], x["shown"]) for x in win["lines"]], win["fields"]

    def _seg_bounds(self, h: HugeLine, a: int) -> tuple[int, int]:
        """字の位置 a の区切りから始まる、1 つの区切りの範囲。空白の並びの途中では切らない（空欄が 2 つに分かれないように）。"""
        b = min(a + SEG_CHARS, h.chars)
        if b >= h.chars:
            return a, h.chars
        # 区切りの字（Python の isspace でない字）まで進める
        step = h.text(b, min(b + 4096, h.chars))
        for i, ch in enumerate(step):
            if not ch.isspace() and ch not in "＿_□☐":
                return a, b + i
        return a, b

    def _grid(self, h: HugeLine) -> list[int]:
        """区切りの始まりの並び：同じ行ならいつも同じ（id の決まり方に関わる）。"""
        g = getattr(h, "grid", None)
        if g is None:
            g = [0]
            a = 0
            while True:
                _a, b = self._seg_bounds(h, a)
                if b >= h.chars:
                    break
                g.append(b)
                a = b
            h.grid = g
        return g

    def _segment_doc(self, n: int, h: HugeLine, k: int):
        g = self._grid(h)
        a = g[k]
        b = g[k + 1] if k + 1 < len(g) else h.chars
        doc = self._slice_doc(n, h.text(a, b).encode(self.dec.encoding), a)
        doc.seg_cut = (a > 0, b < h.chars)
        return doc, a, b

    # ------------------------------------------------------------------ 概要
    def overview(self) -> dict:
        idx = self.idx
        return {"format": {"encoding": self.shown.encoding, "encoding_ja": self.shown.label, "bom": bool(self.shown.bom), "ascii_only": self.shown.ascii_only,
                           "newline": NEWLINE_JA[idx.newline], "mixed_newlines": idx.mixed_newlines, "ends_with_newline": idx.ends_with_newline,
                           "lines": idx.count, **self.shown.info(), "encoding_notes": self.shown.note_lines(),
                           **({"kind_by": "content", "kind_note": self.kind_note} if self.kind_note else {})},
                "lines": idx.count, "longest_line_bytes": idx.longest, "ranges": "paragraphs（行の範囲）", "bytes": self.file_bytes}

    def overview_text(self) -> str:
        idx = self.idx
        out = [f"# {self.name}", "", self.source_line(),
               f"- テキスト・{idx.count:,} 行（いちばん長い行は {idx.longest:,} バイト）。文字コード {self.shown.label}／改行 {NEWLINE_JA[idx.newline]}",
               f"- 大きいので、ここでは全体を出していません。read_window に paragraphs: \"1-500\" のように行の範囲を渡して読みます"
               f"（cursor をたどれば、すべての行のすべての字を読めます。{SEG_CHARS:,} 字を超える行は、字の範囲で分けて出し、分けたことを印で示します）。",
               "- list_fields も paragraphs で、その範囲の欄を出します。欄の id は窓と同じ P行.c字数。"
               f"{SEG_CHARS:,} 字を超える行では P行.x字の位置（行の頭から数えた字の位置。ほかの所を書いても変わりません）。",
               "- 書き込みは、書いた所のバイトだけを変え、ほかのバイトはそのままファイルを写します。字を目印にして書くときは "
               "`{\"@1\": {\"old\": \"元の字\", \"new\": \"書き換えた後の字\"}}`（old がファイルにいくつもあるときは occurrence か near で 1 つに決める）。",
               *self.shown.note_lines(), *([f"- {self.kind_note}。"] if self.kind_note else [])]
        return "\n".join(out) + "\n"

    # ------------------------------------------------------------------ 読む
    def read(self, rng: dict, max_chars: int) -> Page:
        n = self.idx.count
        a, b, skip = _para_args(rng, n)
        body: list[str] = []
        size = 0
        r = a
        stop: tuple[int, int] | None = None  # 続き：(行, 字の位置)
        while r <= b and stop is None:
            last = min(b, r + STRIDE - 1)
            spans = self.line_spans(r, last)
            if not spans:
                break
            # 長い行の前までの、ふつうの行をまとめて読む
            k = 0
            while k < len(spans) and stop is None:
                ln, x, y, _t = spans[k]
                h = self.huge(ln, x, y)
                if h is not None:
                    g = self._grid(h)
                    seg = bisect.bisect_right(g, skip) - 1
                    while seg < len(g):
                        doc, s0, s1 = self._segment_doc(ln, h, seg)
                        shown, _f = self._window_lines(doc)
                        text = shown[0] if shown else f"{ln}｜"
                        prefix = f"{ln}｜"
                        inner = text[len(prefix):]
                        offset = skip - s0 if skip > s0 else 0
                        if offset >= len(inner):
                            seg += 1
                            skip = 0
                            continue
                        if len(inner) - offset > max_chars:
                            piece = inner[offset:offset + max_chars]
                            body.append(f"{ln}｜〔長い行：{s0 + offset + 1:,} 字目から〕{piece}")
                            stop = (ln, s0 + offset + max_chars)
                            break
                        if body and size + len(inner) - offset > max_chars:
                            stop = (ln, s0 + offset)
                            break
                        tail = f"{ln}｜〔長い行を分けて出しています：{s0 + offset + 1:,}〜{s1:,} 字目／全 {h.chars:,} 字〕{inner[offset:]}"
                        body.append(tail)
                        size += len(tail) + 1
                        skip = 0
                        seg += 1
                    if stop is not None:
                        break
                    k += 1
                    continue
                j = k
                while j < len(spans) and spans[j][2] - spans[j][1] <= SEG_CHARS and spans[j][2] - spans[k][1] <= max(SEG_CHARS, max_chars * 2):
                    j += 1
                j = max(j, k + 1)
                first, x0, _y0, _t0 = spans[k]
                z = spans[j][1] if j < len(spans) else (spans[j - 1][2] + _term_len(spans[j - 1][3], self.dec.encoding))
                doc = self._slice_doc(first, self.src.read(x0, z))
                shown, _f = self._window_lines(doc)
                emitted = 0
                for line in shown:
                    if body and size + len(line) + 1 > max_chars:
                        stop = (first + emitted, 0)
                        break
                    if len(line) > max_chars:  # 窓の印を入れて長くなった行
                        body.append(line[skip:skip + max_chars] + cut_note(skip, skip + max_chars, len(line)))
                        if skip + max_chars < len(line):
                            stop = (first + emitted, skip + max_chars)
                            break
                        skip = 0
                    else:
                        body.append(line)
                    size += len(line) + 1
                    emitted += 1
                k = j
            r = last + 1
        if stop is None:
            return Page("\n".join(body), f"{a}〜{b} 行目（全 {comma(n)} 行）", None, "", body)
        ln, ch = stop
        nxt = {"cursor": f"lines:{ln}:{ch}" + (f":{b}" if b != n else "")}
        cov = f"{a}〜{ln - (0 if ch else 1)} 行目" + (f"（{ln} 行目は {ch:,} 字まで）" if ch else "") + f"（全 {comma(n)} 行）"
        return Page("\n".join(body), cov, nxt, f"{comma(b - ln + 1)} 行" + ("（途中の行を含む）" if ch else ""), body)

    # ------------------------------------------------------------------ 欄
    def _spots_of(self, doc):
        from ..docx.window import Spots

        return Spots(doc)

    def fields(self, rng: dict, only_empty: bool, limit: int) -> dict:
        n = self.idx.count
        a, b, skip = _para_args(rng, n)
        out: list[dict] = []
        r = a
        stop = None
        while r <= b and stop is None:
            last = min(b, r + STRIDE - 1)
            for ln, x, y, _t in self.line_spans(r, last):
                h = self.huge(ln, x, y)
                from ..api import _field

                # 長い行は区切りごとに、1 回の呼び出しで SCAN_SEGS 区切りまで（欄のない長い行で、時間の上限に当たらない。続きは cursor で）
                segs = (range(max(bisect.bisect_right(self._grid(h), skip) - 1, 0), len(self._grid(h))) if h is not None else (None,))
                for done, seg in enumerate(segs, 1):
                    doc = self._segment_doc(ln, h, seg)[0] if h is not None else self._slice_doc(ln, self.src.read(x, y))
                    _lines, fl = self._window_lines(doc)
                    for f in fl:
                        d = _field(f)
                        if only_empty and d.get("value") not in (None, "", False):
                            continue
                        out.append(d)
                    if seg is not None and seg + 1 < len(self._grid(h)) and (len(out) >= limit or done >= SCAN_SEGS):
                        stop = (ln, self._grid(h)[seg + 1])
                        break
                if stop:
                    break
                skip = 0
                if len(out) >= limit:
                    stop = (ln + 1, 0)
                    break
            r = last + 1
        res = {"fields": out, "total_lines": n, "next": None}
        if stop and stop[0] <= b:
            res["next"] = {"cursor": f"lines:{stop[0]}:{stop[1]}" + (f":{b}" if b != n else "")}
            res["remaining"] = f"{comma(b - stop[0] + 1)} 行" + (f"（{stop[0]} 行目は {stop[1]:,} 字の後ろから）" if stop[1] else "")
        return res

    # ------------------------------------------------------------------ 書く
    def _resolve(self, key: str, spec, used: dict) -> tuple:
        """書く所：(行の番号, 行の文書, その文書の欄, 文書の頭のバイト位置)。"""
        from ..docx.fill import FillError
        from ..docx.window import WriteError, resolve_write

        if key.startswith("@"):
            if not isinstance(spec, dict):
                raise FillError(f'{key}: 字を目印にした書き込みは {{"old": …, "new": …}} の形で指定してください')
            return self._resolve_anchor(key, spec, used)
        m = re.fullmatch(r"P(\d+)\.x(\d+)(?:\.\d+)?", key)
        if m:
            ln, ch = int(m[1]), int(m[2])
            got = self.line_spans(ln, ln)
            h = self.huge(ln, got[0][1], got[0][2]) if got else None
            if h is None:
                raise FillError(f"{key}: そのような欄はありません（{ln} 行目は、字の位置の id を使う長い行ではありません）")
            g = self._grid(h)
            seg = bisect.bisect_right(g, ch) - 1
            doc, s0, _s1 = self._segment_doc(ln, h, seg)
            sp = self._spots_of(doc)
            it = sp.items.get(key)
            if it is None:
                raise FillError(f"{key}: そのような欄はありません")
            return ln, doc, it, h.byte_of(s0)
        m = re.match(r"P(\d+)\b", key)
        if not m:
            raise FillError(f"{key}: そのような欄はありません（P行.c字数 の形。窓の［ ］の中の id を使ってください）")
        ln = int(m[1])
        got = self.line_spans(ln, ln)
        if not got:
            raise FillError(f"{key}: {ln} 行目はありません（全 {self.idx.count:,} 行）")
        _n, x, y, _t = got[0]
        if self.huge(ln, x, y) is not None:
            raise FillError(f"{key}: {ln} 行目は長いので、字の位置の id（P{ln}.x…）を使います。read_window か list_fields で、その所の id を確かめてください")
        doc = self._slice_doc(ln, self.src.read(x, y))
        sp = self._spots_of(doc)
        it = sp.items.get(key)
        if it is None:
            raise FillError(f"{key}: そのような欄はありません{sp.items.hint(key)}")
        return ln, doc, it, x

    def _find_all(self, needle: bytes, a: int, b: int, cap: int = 1000) -> tuple[list[int], int]:
        """バイトの並びの位置（はじめの cap 個）と、全部の数。枠をまたいでも探す。"""
        found: list[int] = []
        count = 0
        overlap = len(needle) - 1
        pos = a
        prev = b""
        unit = 2 if self.dec.encoding.startswith("utf-16") else 4 if self.dec.encoding.startswith("utf-32") else 1
        for c in self.src.chunks(a, b):
            buf = prev + c
            base = pos - len(prev)
            i = 0
            while True:
                i = buf.find(needle, i)
                if i < 0:
                    break
                if unit == 1 or (base + i - self.start) % unit == 0:
                    count += 1
                    if len(found) < cap:
                        found.append(base + i)
                i += 1
            prev = buf[-overlap:] if overlap else b""
            # 次の枠の頭に重ねる分（すでに見つけた所を数え直さないよう、needle より 1 バイト短く）
            pos += len(c)
        return found, count

    def _line_of(self, pos: int) -> tuple[int, int, int]:
        """バイト位置 pos を含む行：(行の番号, 中身の始め, 中身の終わり)。"""
        idx = self.idx
        i = bisect.bisect_right(idx.starts, pos) - 1
        first = i * STRIDE + 1
        for ln, x, y, t in self.line_spans(first, min(first + STRIDE - 1, idx.count)):
            if pos <= y:
                return ln, x, y
        raise ValueError("行を決められません")

    def _resolve_anchor(self, key: str, spec: dict, used: dict) -> tuple:
        from ..docx.fill import FillError
        from ..docx.window import WriteError, resolve_write

        old = spec.get("old")
        if not isinstance(old, str) or not old:
            raise FillError(f"{key}: old（元の字）を指定してください")
        needle = old.encode(self.dec.encoding)
        a, b = self.start, self.src.size
        near = spec.get("near")
        if near:
            m = re.match(r"P(\d+)\b", str(near))
            if not m:
                raise FillError(f"{key}: near は、近くの欄の id（P行.c字数）で指定してください")
            got = self.line_spans(int(m[1]), int(m[1]))
            if not got:
                raise FillError(f"{key}: near の {m[1]} 行目はありません")
            a, b = got[0][1], got[0][2]
        found, count = self._find_all(needle, a, b)
        if not count:
            raise FillError(f"{key}: 「{old[:40]}」が{'その行' if near else 'ファイル'}に見つかりません")
        occ = spec.get("occurrence")
        if occ is None and count > 1:
            where = "、".join(str(self._line_of(p)[0]) for p in found[:20])
            raise FillError(f"{key}: 「{old[:40]}」が {count:,} か所あります（行：{where}{'ほか' if count > 20 else ''}）。occurrence（何番目か）か near（近くの欄の id）で 1 つに決めてください")
        if occ is not None and not 1 <= int(occ) <= len(found):
            raise FillError(f"{key}: occurrence は 1〜{len(found)} です（全 {count:,} か所）")
        pos = found[(int(occ) if occ else 1) - 1]
        ln, x, y = self._line_of(pos)
        h = self.huge(ln, x, y)
        if h is not None:
            # 長い行：字の前後だけの文書にして書く所を決める（文書の頭のバイト位置を覚える）
            ch = len(self.src.read(x, pos).decode(self.dec.encoding)) if pos - x < 1 << 26 else None
            if ch is None:
                raise FillError(f"{key}: 行の頭から遠すぎて、字の位置を決められません")
            g = self._grid(h)
            seg = bisect.bisect_right(g, ch) - 1
            doc, s0, s1 = self._segment_doc(ln, h, seg)
            base = h.byte_of(s0)
            if ch + len(old) > s1:  # 区切りをまたぐ字：その字だけの文書にする
                doc = self._slice_doc(ln, h.text(ch, ch + len(old)).encode(self.dec.encoding), ch)
                base = pos
        else:
            doc = self._slice_doc(ln, self.src.read(x, y))
            base = x
        # 文書の中で何番目かを決める（文書の頭から pos までに、同じ字がいくつあるか）
        inner = dict(spec)
        prefix = self.src.read(base, pos).decode(self.dec.encoding)
        inner["occurrence"] = prefix.count(old) + 1
        inner.pop("near", None)
        sp = self._spots_of(doc)
        try:
            it = resolve_write(sp, inner, used)
        except WriteError as e:
            raise FillError(f"{key}: {e}") from None
        return ln, doc, it, base

    def fill(self, values: dict, out_path: str, force: bool, overwrite: bool, said: dict | None) -> dict:
        from ..docx.fill import LINE_BREAKS, FillError, _as_bool, _is_text_value
        from ..txt.fill import _edit_of

        edits: list[tuple[int, int, bytes]] = []
        told: dict = {}
        labels: dict = {}
        anchored: dict = {}
        problems: list[str] = []
        checks: list[tuple] = []  # 読み戻し：(id, 元の字の範囲, 書いた字)
        used: dict = {}
        self._binary_check()
        for key, value in values.items():
            try:
                ln, doc, it, base = self._resolve(str(key), value, used)
                if str(key).startswith("@"):
                    text = it["value"]
                    fid = it["id"]
                    anchored[key] = fid
                else:
                    fid = str(key)
                    text = _as_bool(value) if it["kind"] == "checkbox" else _is_text_value(value)
                if isinstance(text, str) and any(ch in text for ch in LINE_BREAKS):
                    raise FillError(f"{key}: 改行・行を分ける字（CR・LF・U+2028・U+0085・改ページの字など）は書けません（文字だけのファイルでは、"
                                    "読み手によって行が増えて、ほかの行の id が変わります）。1行ずつ、それぞれの行の欄に書いてください")
                e = _edit_of(it, text)
                told[fid] = text
                labels[fid] = it.get("label", "")
                if e is None:
                    continue
                a, b, v = e
                p = it["para"]
                x, y = doc.dec.span(doc.blob, p.a + a, p.a + b)
                try:
                    enc = doc.dec.encode(v)
                except EncodingError as ex:
                    raise FillError(f"{key}: {ex}") from None
                edits.append((base + x, base + y, enc))
                checks.append((fid, base + x, base + y, v, ln))
            except FillError as e:
                problems.append(str(e))
        edits.sort(key=lambda t: (t[0], t[1]))
        for (a1, b1, _x), (a2, _b2, _y) in zip(edits, edits[1:]):
            if a2 < b1:
                problems.append("同じ所に2つの書き込みが重なりました")
                break
        if problems:
            raise FillError("\n".join(problems))
        warnings = []
        if self.dec.ascii_only and any(not e[2].isascii() for e in edits):
            warnings.append("もとのファイルは ASCII の字だけで文字コードが決まらないので、UTF-8（BOM なし）で書きました")
        dst = Path(out_path).expanduser().resolve()
        tmp = temp_beside(dst)
        try:
            ranges, h_old, h_new = stream_patch(self.src, edits, tmp)
            unexpected = verify_ranges(tmp, ranges)
            back = []
            # 読み戻し：書いた所のバイトを、書いたファイルの位置から読み直す（前の編集の分だけ位置がずれる）
            new_src = Source(str(tmp))
            shift = 0
            prev = sorted(checks, key=lambda t: t[1])
            try:
                for fid, x, y, v, ln in prev:
                    enc = v.encode(self.dec.encoding)
                    got = new_src.read(x + shift, x + shift + len(enc))
                    if got != enc:
                        back.append({"id": fid, "written": v, "read_back": got.decode(self.dec.encoding, "replace")})
                    shift += len(enc) - (y - x)
            finally:
                new_src.close()
            n = self.idx.count
            result = {"kind": "txt", "out_path": str(dst), "original_sha256": h_old, "new_sha256": h_new, "proof_ok": not unexpected,
                      "unexpected": unexpected[:20], "structure_errors": [], "warnings": warnings, "streamed": True,
                      "proof": (f"書いた {len(edits)} か所のバイトのほかは 1 バイトも変わっていません（変えていない {len(ranges.items)} の範囲・{comma(ranges.total)} バイトの"
                                "ハッシュを、書いた後のファイルを読み直して比べました）"),
                      "read_back_by": "offset", "read_back_ok": not back, "read_back_differences": back[:30],
                      "ids_after_writing": "文字だけのファイルには、書いた所の印を置けません。書いた後のファイルでは、書いた字も印刷された字に数えるので、"
                                           "同じ行の後ろの空欄の id の字数が増えます。続けて書くときは、書いた後のファイルを開き直して id を確かめるか、元のファイルに、まとめて書いてください"}
            if anchored:
                result["anchored"] = anchored
            _finish_fill(result, told, labels, said, tmp, dst, self.path, overwrite, f"list_fields に paragraphs を渡すと、その範囲の欄が分かります（全 {comma(n)} 行）")
            return result
        finally:
            tmp.unlink(missing_ok=True)

    # ------------------------------------------------------------------ 段階
    def grade(self) -> dict:
        from ..grade import Reason, _final, _locked
        from ..openkinds import encoding_reason

        if self.dec.binary:
            return _locked("読めない", "文字のファイルではありません（どの文字コードとしても読めず、NUL のバイトがあります）", "読めないので、書けません")
        write = [Reason("B", "欄を範囲ごとに見つける", "大きなテキストなので、入力欄は頼まれた行の範囲ごとに見つけます。欄のない所は字を目印にして書けます")]
        reasons: list = []
        if self.kind_note:
            reasons.append(Reason("B", "形式を中身から決めた", self.kind_note))
            write.append(Reason("B", "形式を中身から決めた", "書く前に、エージェントが、このファイルが文字のファイルとして書いてよいものかを"
                                "字を読んで確かめます（書くのは、書いた所の字のバイトだけ）"))
        enc, wenc = encoding_reason(self.dec, "A")
        g = _final(reasons + ([enc] if enc else []), write + ([wenc] if wenc else []))
        g.deferred = ["未確認：全部の行の、入力欄の洗い出し（audit_document で全部の行を確かめます）"]
        return g

    # ------------------------------------------------------------------ 検査
    def audit_begin(self) -> dict:
        return {"next_row": 1, "lines": 0, "fields": 0, "chars": 0, "problems": [], "done": False, "checks": [], "seg": 0}

    def audit_step(self, state: dict) -> dict:
        if self.recoded:
            return self._audit_whole(state)
        from ..txt.document import split_lines

        n = self.idx.count
        r = state["next_row"]
        if r <= n:
            last = min(n, r + STRIDE * 8 - 1)
            spans = self.line_spans(r, last)
            for ln, x, y, t in spans:
                h = self.huge(ln, x, y)
                if h is not None:
                    g = self._grid(h)
                    for k in range(state["seg"], len(g)):
                        doc, s0, s1 = self._segment_doc(ln, h, k)
                        state["fields"] += len(self._spots_of(doc).items)
                        state["chars"] += len(doc.text)
                        state["seg"] = k + 1
                        state["next_row"] = ln
                        limits.progress(f"{ln} 行目（{s1:,}／{h.chars:,} 字）を確かめています", ln, n)
                        return state
                    state["seg"] = 0
                    if state["chars"] < 0:
                        pass
                    state["next_row"] = ln + 1
                    state["lines"] += 1
                    return state
                data = self.src.read(x, y)
                text = data.decode(self.dec.encoding)
                if len(split_lines(text)) > 1 and len(state["problems"]) < 30:
                    state["problems"].append(f"{ln} 行目：行の区切りの読み方が、別の読み方と違います")
                state["chars"] += len(text)
                state["lines"] += 1
            if spans:
                first = spans[0][0]
                data = self.src.read(spans[0][1], spans[-1][2])
                doc = self._slice_doc(first, data)
                state["fields"] += len(self._spots_of(doc).items)
            state["next_row"] = last + 1
            limits.progress("行を確かめています", state["next_row"], n)
            return state
        state["checks"] = _text_checks("読み落とし（行の字）", state,
                                       f"全 {n:,} 行・{state['chars']:,} 字をデコードし、行の区切りを別の読み方と照らしました。見つけた欄 {state['fields']:,} か所")
        state["checks"].append(self._txt_write_trial())
        state["done"] = True
        return state

    def _txt_write_trial(self) -> dict:
        import tempfile

        from ..docx.fill import FillError

        fl = self.fields({"paragraphs": "1-"}, True, 3).get("fields", [])
        wr = [f for f in fl if f.get("kind") in ("inline", "text")][:2]
        if not wr:
            return {"name": "書き込みの試し", "ok": True, "agent": False, "detail": "はじめの所に、書ける欄が見つかりませんでした（試しませんでした）", "items": []}
        with tempfile.TemporaryDirectory(prefix="todoku-audit-") as td:
            try:
                res = self.fill({f["id"]: "テスト" for f in wr}, str(Path(td) / f"t{Path(self.name).suffix}"), True, True, None)
            except (FillError, ValueError) as e:
                return {"name": "書き込みの試し", "ok": False, "agent": False, "detail": f"書けませんでした：{e}", "items": []}
        return {"name": "書き込みの試し", "ok": bool(res.get("proof_ok") and res.get("read_back_ok")), "agent": False,
                "detail": f"{len(wr)} か所に試しの値を書き、{res.get('proof')}", "items": res.get("unexpected") or []}


def _term_len(t: str, enc: str) -> int:
    return len(t.encode(enc)) if t else 0


def _para_args(rng: dict, n: int) -> tuple[int, int, int]:
    cursor = rng.get("cursor")
    if cursor:
        m = re.fullmatch(r"lines:(\d+)(?::(\d+))?(?::(\d+))?", str(cursor))
        if not m:
            raise RangeError(f"cursor「{cursor}」はこの書類の続きの印ではありません（前の結果の next をそのまま渡してください）")
        return int(m[1]), min(int(m[3]) if m[3] else n, n), int(m[2] or 0)
    spec = next((rng[k] for k in ("paragraphs", "rows", "records") if rng.get(k) is not None), None)
    a, b = parse_range(spec, "行")
    a = 1 if a is None else a
    b = n if b is None else min(b, n)
    if a > max(n, 1):
        raise RangeError(f"{a} 行目はありません（全 {n:,} 行）")
    return a, b, 0
