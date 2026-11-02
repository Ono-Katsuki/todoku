"""RTF：エージェントが「全て正確に読み書きできるか」を、決まった処理で確かめる。

読む: 窓とは別の読み方（段落や表を作らずに、字句を正規表現で拾い、見えない宛先を飛ばして、\\'hh・\\uN を文字にするだけの
      読み手）で、本文・表・ヘッダー/フッター・脚注・コメント・図形の字・フィールドの結果の字を数え、窓に1字残らず出ているかを確かめる。
書く: 窓が示したすべての入力欄に試しの値を一度に書き込み、証明（書いた段落の外のバイトは1バイトも違わない）と読み戻しを確かめる。
"""

from __future__ import annotations

import codecs
import re
from collections import Counter

from ..docx.audit import _sample, anchored_check, macro_check
from ..docx.fill import FillError
from ..xlsx.audit import Audit, Check, agent_places_check
from .document import CHARSET_CP, SKIP, RtfDocument
from .fill import fill, prove, written_of
from .window import build, markdown_parts

BOXES = set("□☐☑☒■")
_TOK = re.compile(rb"\\([a-zA-Z]{1,32})(-?\d{1,10})? ?|\\'([0-9a-fA-F]{2})|\\(.)|([{}])|([\r\n]+)|([^\\{}\r\n]+)", re.S)
_CHARS = {"emdash": "—", "endash": "–", "bullet": "•", "lquote": "‘", "rquote": "’", "ldblquote": "“", "rdblquote": "”"}
# 窓の段落の字に入れない宛先：書式・情報など（SKIP）に、フィールドの命令・段落番号・しおり・フォームの欄の定義を足す
_HIDE = set(SKIP) | {"fldinst", "listtext", "pntext", "bkmkstart", "bkmkend", "formfield"}


def independent_text(blob: bytes) -> Counter:
    """もう1つの読み方：字句を拾い、見えない宛先を飛ばし、字を数える（空白と □☐☑☒■ は数えない）。"""
    head = blob[:200_000]
    m = re.search(rb"\\ansicpg(\d+)", head)
    ansi = f"cp{int(m[1])}" if m else "cp1252"
    try:
        codecs.lookup(ansi)
    except LookupError:
        ansi = "cp1252"
    fonts = {}
    for fm in re.finditer(rb"\{\\f(\d+)", blob[:400_000]):
        seg = blob[fm.end():fm.end() + 300]
        stop = min([x for x in (seg.find(b";"), seg.find(b"}")) if x >= 0] or [len(seg)])
        cm = re.search(rb"\\fcharset(\d+)", seg[:stop])
        if cm:
            fonts.setdefault(int(fm[1]), int(cm[1]))
    deff = re.search(rb"\\deff(\d+)", head)
    out = Counter()
    stack = []
    hide, uc, font, star, first = False, 1, int(deff[1]) if deff else None, False, True
    skip = 0
    high = 0
    buf = bytearray()
    cp_buf = ansi

    def cp_of(f):
        cs = fonts.get(f)
        if cs is None or cs == 1:
            return ansi
        if cs == 2:
            return "symbol"
        return CHARSET_CP.get(cs, ansi)

    def flush():
        nonlocal buf
        if buf and not hide:
            if cp_buf == "symbol":
                text = "◆" * len(buf)
            else:
                text = bytes(buf).decode(cp_buf, "replace")
            for ch in text:
                if not ch.isspace() and ch not in BOXES:
                    out[ch] += 1
        buf = bytearray()

    for t in _TOK.finditer(blob):
        word, param, hexv, sym, brace, _nl, text = t.groups()
        if skip > 0 and brace is None and _nl is None:
            if text is not None:
                n = len(text)
                if n <= skip:
                    skip -= n
                    continue
                text = text[skip:]
                skip = 0
            else:
                skip -= 1
                continue
        if brace is not None:
            flush()
            skip = 0
            if brace == b"{":
                stack.append((hide, uc, font))
                star, first = False, True
            else:
                if stack:
                    hide, uc, font = stack.pop()
                first = False
            continue
        if _nl is not None:
            continue
        if hexv is not None:
            cp = cp_of(font)
            if buf and cp != cp_buf:
                flush()
            cp_buf = cp
            buf.append(int(hexv, 16))
            first = False
            continue
        flush()
        if sym is not None:
            s = sym.decode("latin-1")
            if s == "*":
                star = True
                continue
            first = False
            if s in "\\{}" and not hide:
                out[s] += 1
            elif s == "_" and not hide:
                out["‑"] += 1
            continue
        if word is not None:
            w = word.decode("ascii")
            if first and (star or w in _HIDE) and w not in ("shpinst", "do"):
                hide = True  # 見えない宛先（図形の定義 \\shpinst と古い図形 \\do は、中の字の箱が見える）
            first = False
            if w == "u" and param is not None:
                n = int(param)
                if n < 0:
                    n += 65536
                ch = None
                if 0xD800 <= n < 0xDC00:
                    high = n  # サロゲートの組の前半：次の \\uN と合わせる
                elif 0xDC00 <= n < 0xE000:
                    ch = chr(0x10000 + ((high - 0xD800) << 10) + (n - 0xDC00)) if high else None
                    high = 0
                else:
                    ch = chr(n)
                if ch is not None and not hide and not ch.isspace() and ch not in BOXES:
                    out[ch] += 1
                skip = uc
            elif w == "uc" and param is not None:
                uc = int(param)
            elif w == "f" and param is not None:
                font = int(param)
            elif w in _CHARS and not hide:
                out[_CHARS[w]] += 1
            continue
        if text is not None:
            first = False
            if any(x >= 0x80 for x in text):
                cp = cp_of(font)
                if buf and cp != cp_buf:
                    flush()
                cp_buf = cp
                buf.extend(text)
                flush()
                continue
            if not hide:
                for ch in text.decode("latin-1"):
                    if not ch.isspace() and ch not in BOXES:
                        out[ch] += 1
    flush()
    return out


def _count(texts) -> Counter:
    c = Counter()
    for t in texts:
        for ch in t:
            if not ch.isspace() and ch not in BOXES:
                c[ch] += 1
    return c


def check_reading(doc: RtfDocument, win: dict) -> list[Check]:
    raw = independent_text(doc.blob)
    shown = Counter()
    for st in win["stories"]:
        shown += _count(st["_plain"])
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    return [Check("読み落とし（文書の文字）", not missing and not extra,
                  f"本文・表・ヘッダー/フッター・脚注・コメント・図形の字 {sum(raw.values()):,} 字を、窓とは別の読み方"
                  "（段落や表を作らずに字句を拾うだけの読み手）で数え、すべて窓に出ていることを確かめました", missing + extra)]


def check_writing(blob: bytes, filename: str, win: dict) -> list[Check]:
    targets = {}
    protected = []
    for f in win["fields"]:
        if f.get("readonly"):
            protected.append(f["id"])
            continue
        targets[f["id"]] = _sample(f)
    try:
        res = fill(RtfDocument(blob, filename), targets)
    except FillError as e:
        return [Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30])]
    pr = prove(blob, res.blob, written_of(res.patches))
    note = f"（保護で書けない欄 {len(protected)} か所は除外）" if protected else ""
    out = [Check("全入力欄への書き込み", pr.ok, f"入力欄 {len(win['fields']):,} か所のうち {len(targets):,} に試しの値を書き込み{note}。{pr.headline()}",
                 [*pr.unexpected, *pr.lint_errors])]
    after = build(RtfDocument(res.blob, filename))
    got = {f["id"]: f.get("value") for f in after["fields"]}
    written = {p.id: p.value for p in res.patches}
    wrong = []
    for k, v in targets.items():
        g = got.get(k)
        w = written.get(k, v)
        ok = bool(g) == v if isinstance(v, bool) else str(g) == str(w)
        if not ok:
            wrong.append(f"{k}: 書いた値 {w!r} → 窓では {g!r}")
    out.append(Check("読み戻し", not wrong, f"書いた {len(targets):,} 欄のうち {len(targets) - len(wrong):,} 欄が、書き込み後の窓にそのまま出ている", wrong))
    return out


def run(blob: bytes, filename: str = "") -> Audit:
    doc = RtfDocument(blob, filename)
    win = build(doc)
    au = Audit(filename)
    au.checks += check_reading(doc, win)
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000,
                           f"Markdown {sum(len(t) for _, t in parts):,} 文字" + (f"を {len(parts)} 部分に分け、いちばん大きい部分も {biggest:,} 文字"
                                                                             if len(parts) > 1 else "") + "（AIが1回で読める目安 15万文字以内）"))
    au.checks += check_writing(blob, filename, win)
    ap = agent_places_check(win["fields"])
    if ap is not None:
        au.checks.append(ap)
    if not (doc.protection_enforced and doc.protection):
        au.checks.append(anchored_check(blob, filename, RtfDocument, fill, prove))
    au.checks.append(macro_check(doc, win))
    return au
