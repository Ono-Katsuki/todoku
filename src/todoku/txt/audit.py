"""文字だけのファイル（.txt）：エージェントが「全て正確に読み書きできるか」を、決まった処理で確かめる。

読む: 窓とは別の読み方（BOM を codecs の印で見分け、窓が示した文字コードで Python の codecs が厳しく読み、io の
      決まりの改行（CRLF・LF・CR）で行に分ける）で、行の数・行ごとの字と改行が窓と同じこと、窓の Markdown に
      すべての行が「行番号｜字」でそのまま出ていること（空欄の印に置き換えた所のほか、1字残らず）を確かめる。
書く: 窓が示したすべての入力欄に試しの値を一度に書き込み、証明（書いた所の字のバイトのほかは1バイトも違わない）と
      読み戻し（書いた後のファイルを読み直した行の字）を確かめる。字を目印にした書き込みも試す。
"""

from __future__ import annotations

import codecs
import io
from collections import Counter

from ..docx.audit import _sample
from ..docx.fill import FillError
from ..docx.window import _in_field
from ..xlsx.audit import Audit, Check, agent_places_check
from .document import TxtDocument, TxtError
from .fill import fill, prove, written_of
from .window import build, line_text, mark, markdown_parts

_BOMS = ((codecs.BOM_UTF8, "utf-8"), (codecs.BOM_UTF16_LE, "utf-16-le"), (codecs.BOM_UTF16_BE, "utf-16-be"))


def independent_lines(blob: bytes, encoding: str) -> tuple[bytes, list[tuple[str, str]]]:
    """(BOM, [(行の字, 改行)])：窓とは別の読み方。encoding は窓が示した文字コード（読めなければ UnicodeDecodeError）。"""
    bom = next((b for b, enc in _BOMS if blob.startswith(b) and enc == encoding), b"")
    text = codecs.decode(blob[len(bom):], encoding, "strict")
    out = []
    for raw in io.StringIO(text, newline="").readlines():  # newline="" ：改行の字を変えずに、CRLF・LF・CR で区切る
        body = raw.rstrip("\r\n")
        out.append((body, raw[len(body):]))
    return bom, out


def check_reading(doc: TxtDocument, win: dict, md: str) -> Check:
    enc = win["format"]["encoding"]
    problems: list[str] = []
    try:
        bom, other = independent_lines(doc.blob, enc)
    except UnicodeDecodeError as e:
        return Check("読み落とし（行の字）", False, f"窓が示した文字コード（{enc}）では、ファイルを読めません：{e.reason}")
    if bool(bom) != win["format"]["bom"]:
        problems.append("BOM のあるなしが、窓と違います")
    mine = win["lines"]
    if len(mine) != len(other):
        problems.append(f"行の数が違います（窓 {len(mine)}・別の読み方 {len(other)}）")
    shown_lines = set(md.split("\n"))
    covered: Counter = Counter()
    for x, (text, term) in zip(mine, other):
        if x["text"] != text:
            problems.append(f"{x['n']} 行目：窓「{x['text'][:30]}」・別の読み方「{text[:30]}」")
        elif x["terminator"] != term:
            problems.append(f"{x['n']} 行目の改行が違います")
        # 窓の Markdown の行：別の読み方の字の、空欄の所だけを印に置き換えたもの
        out, at = [], 0
        for a, b, fid, kind in sorted(tuple(s) for s in x["spans"]):
            if a < at or b > len(text):
                problems.append(f"{x['n']} 行目：空欄の範囲が正しくありません（{fid}）")
                break
            out += [text[at:a], mark(fid, kind, text[a:b])]
            if kind != "placeholder":
                covered.update(ch for ch in text[a:b] if not ch.isspace())
            at = b
        out.append(text[at:])
        if line_text(x["n"], "".join(out)) not in shown_lines:
            problems.append(f"{x['n']} 行目が、窓の Markdown にそのままの形で出ていません")
        if len(problems) > 30:
            break
    total = sum(1 for _t, _e in other for ch in _t if not ch.isspace())
    note = ""
    if covered:
        note = "。空欄の印に置き換えた字（空白のほか）：" + "・".join(f"「{ch}」{n} 字" for ch, n in covered.most_common(8))
    return Check("読み落とし（行の字）", not problems,
                 f"{len(other):,} 行・字 {total:,} 字を、窓とは別の読み方（Python の codecs で {enc} として厳しく読み、"
                 "CRLF・LF・CR で行に分ける）で読み、すべての行の字と改行が窓と同じで、窓の Markdown にそのまま出ていることを確かめました"
                 + note, problems)


def _writable(doc: TxtDocument, value, ascii_value: str):
    """試しの値：ファイルの文字コードで書けない字なら、ASCII の字にする（Latin-1 として読んだファイルなど）。"""
    if not isinstance(value, str):
        return value
    try:
        value.encode(doc.dec.encoding)
        return value
    except UnicodeEncodeError:
        return ascii_value


def check_writing(blob: bytes, filename: str, win: dict, encoding: str | None = None) -> list[Check]:
    doc = TxtDocument(blob, filename, encoding)
    targets = {f["id"]: _writable(doc, _sample(f), "12") for f in win["fields"]}
    try:
        res = fill(TxtDocument(blob, filename, encoding), targets)
    except FillError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    pr = prove(blob, res.blob, written_of(res.patches), encoding)
    out = [Check("全入力欄への書き込み", not pr.unexpected and not pr.lint_errors,
                 f"入力欄 {len(targets):,} か所すべてに試しの値を書き込みました。{pr.headline()}", [*pr.unexpected, *pr.lint_errors])]
    out.append(Check("読み戻し", not pr.readback,
                     f"書き込み後のファイルを読み直し、{pr.lines_total:,} 行のすべてで、字が「元の字の、書いた所だけを替えたもの」"
                     "になっていることを確かめました" if not pr.readback else "書き込み後のファイルを読み直すと、思っていた字と違う行があります",
                     pr.readback[:30]))
    return out


def anchored_check(blob: bytes, filename: str, encoding: str | None = None) -> Check:
    """字を目印にした書き込み：ファイルの中に1つしかない4字の後ろに試しの字を書き、証明と読み戻し（その行の字）を確かめる。"""
    doc = TxtDocument(blob, filename, encoding)
    full = "\n".join(ln.text for ln in doc.lines)
    probe = _writable(doc, "（試し）", "(test)")
    tried, last = 0, ""
    for ln in doc.lines:
        t = ln.text
        for i in range(0, len(t) - 3):
            frag = t[i:i + 4]
            if frag != frag.strip(" 　\t") or full.count(frag) != 1 or any(_in_field(ln, j) for j in range(i, i + 5)):
                continue
            tried += 1
            try:
                res = fill(TxtDocument(blob, filename, encoding), {"@1": {"old": frag, "new": frag + probe}})
            except FillError as e:
                last = str(e)
                if tried >= 5:
                    break
                continue
            pr = prove(blob, res.blob, written_of(res.patches), encoding)
            return Check("字を目印にした書き込み", pr.ok,
                         f"規則で示していない所への書き込み（{ln.number} 行目の「{frag}」の後ろに「{probe}」）を {res.patches[0].id} として書き、"
                         f"{pr.headline()}。読み戻し：{'その行の字が思っていたとおり' if not pr.readback else '思っていた字と違う'}",
                         [*pr.unexpected, *pr.lint_errors, *pr.readback])
        if tried >= 5:
            break
    if not tried:
        return Check("字を目印にした書き込み", True, "ファイルの中に1つしかない字がないので、試していません（同じ字の所は near・occurrence で指す）")
    return Check("字を目印にした書き込み", False, f"字を目印にした書き込みを {tried} か所で試し、書けませんでした", [last])


def run(blob: bytes, filename: str = "", encoding: str | None = None) -> Audit:
    au = Audit(filename)
    try:
        doc = TxtDocument(blob, filename, encoding)
    except TxtError as e:
        au.checks.append(Check("開く", False, f"文字のファイルとして読めません：{e}"))
        return au
    win = build(doc)
    parts = markdown_parts(win)
    md = "\n".join(t for _n, t in parts)
    au.checks.append(check_reading(doc, win, md))
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {sum(len(t) for _, t in parts):,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                                             if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(blob, filename, win, encoding)
    ap = agent_places_check(win["fields"])
    if ap is not None:
        au.checks.append(ap)
    au.checks.append(anchored_check(blob, filename, encoding))
    return au
