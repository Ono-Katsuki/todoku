"""古い Word（.doc）：エージェントが「全て正確に読み書きできるか」を確かめる（読み落とし・全欄への書き込み・証明・読み戻し）。

証明：書き込み後の文書を読み直し、
- WordDocument・表・Data 以外のストリームはバイト一致
- 段落の数が同じで、書いた欄のない段落の文字はすべて同じ。書いた欄のある段落の文字は「空欄を値に置き換えたもの」
- 段落の書式（PAPX）は、すべての段落で同じ
- ピースの表が文書の全体をすき間なく覆い、位置の表（段落の区切り・フィールド・しおり など）が並び順を保つ
"""

from __future__ import annotations

import hashlib
import struct
from collections import Counter
from dataclasses import dataclass, field

from ..cfb import CFB
from ..docx.document import WINGDINGS, iter_paras
from ..docx.window import Spots, build, markdown_parts
from ..xlsx.audit import Audit, Check
from .book import DocFile
from .fill import SHIFT, FillError, fill

BOXES = set("□☐☑☒■")


def raw_chars(doc: DocFile) -> list[str]:
    """ピースの表から直接（段落や表に分けずに）文字を拾う。フィールドの命令の文字と、制御の字は除く。"""
    out = []
    depth = []
    text = doc.text
    for cp, ch in enumerate(text):
        o = ord(ch)
        if o == 0x13:
            depth.append("instr")
            continue
        if o == 0x14 and depth:
            depth[-1] = "result"
            continue
        if o == 0x15 and depth:
            depth.pop()
            continue
        if depth and depth[-1] == "instr":
            continue
        if o < 0x20:
            continue
        if o == 0x28:
            pr = doc.char_props(cp)
            if pr["symbol"] and pr["spec"]:
                code = pr["symbol"][1]
                code2 = code + 0xF000 if code < 0xF000 else code
                out.append(WINGDINGS.get(code2, "◆"))
                continue
        out.append(ch)
    return out


def _count(texts) -> Counter:
    c = Counter()
    for t in texts:
        for ch in t:
            if not ch.isspace() and ch not in BOXES:
                c[ch] += 1
    return c


@dataclass
class DocProof:
    original_sha256: str
    new_sha256: str
    entries_total: int = 0
    entries_identical: int = 0
    paragraphs_total: int = 0
    paragraphs_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        return (f"複合ファイルのストリーム {self.entries_identical}/{self.entries_total} がバイト一致。段落 {self.paragraphs_total:,} のうち、"
                f"変わったのは書いた欄のある {self.paragraphs_changed} 段落だけ"
                + ("" if not self.unexpected else f"（想定外の変更 {len(self.unexpected)} 件）"))


def lint(doc: DocFile) -> list[str]:
    errs = []
    total = sum(doc.ccp.values())
    cp = 0
    for p in doc.pieces:
        if p.cp0 != cp:
            errs.append(f"ピースの表にすき間があります（{cp} → {p.cp0}）")
        cp = p.cp1
    if doc.pieces and doc.pieces[-1].cp1 < total:
        errs.append(f"ピースの表が文書の終わり（{total}）まで届いていません（{doc.pieces[-1].cp1}）")
    for name, cb in SHIFT.items():
        f0, lcb = doc.fc.get(name, (0, 0))
        if not lcb or name in ("Plcfuim", "Plcfcookie", "PlcfCookieOld"):
            continue  # 並んでいなくてよい表・最後の位置が定義されない表
        n = (lcb - 4) // (4 + cb)
        vals = struct.unpack_from(f"<{n + 1}I", doc.table, f0)
        if any(b < a for a, b in zip(vals, vals[1:])):
            errs.append(f"{name} の位置の並びが崩れています")
    if doc.lw[0] != len(doc.wd):
        errs.append("WordDocument の長さ（cbMac）が合いません")
    return errs


def prove(original: bytes, new: bytes, written: dict) -> DocProof:
    pr = DocProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    ca, cb_ = CFB(original), CFB(new)
    sa, sb = ca.streams(), cb_.streams()
    a = DocFile(original)
    allowed = {"WordDocument", a.table_name, "Data"}
    for name in sa:
        pr.entries_total += 1
        if name in sb and ca.read(name) == cb_.read(name):
            pr.entries_identical += 1
        elif name not in allowed:
            pr.unexpected.append(f"ストリーム {name} が変わっています")
    try:
        b = DocFile(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません: {e}")
        return pr
    pr.lint_errors += lint(b)
    pa = [p for _k, _p, bl in a.all_blocks() for p in iter_paras(bl)]
    pb = [p for _k, _p, bl in b.all_blocks() for p in iter_paras(bl)]
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.unexpected.append(f"段落の数が変わっています（{len(pa)} → {len(pb)}）")
        return pr
    sp = Spots(a)
    targets: dict[int, list] = {}
    index = {id(p): i for i, p in enumerate(pa)}
    for fid, value in written.items():
        it = sp.items.get(fid)
        if it is None:
            continue
        p = it.get("para") or (it.get("control").para if it.get("control") is not None else None)
        if p is not None and id(p) in index:
            targets.setdefault(index[id(p)], []).append((fid, it, value))
    for i, (x, y) in enumerate(zip(pa, pb)):
        if a.para_props(x.cp1) != b.para_props(y.cp1):
            pr.unexpected.append(f"段落 {i + 1} の書式（表の中か・行の終わりか）が変わっています")
        if i in targets:
            pr.paragraphs_changed += 1
            want = _expected(x, targets[i])
            if want is not None and y.text != want:
                pr.unexpected.append(f"段落 {i + 1}：書いた後の文字が「{y.text[:50]}」で、思っていた「{want[:50]}」と違います")
        elif x.text != y.text:
            pr.unexpected.append(f"書いていない段落 {i + 1} の文字が変わっています（「{x.text[:30]}」→「{y.text[:30]}」）")
    return pr


def _expected(p, items) -> str | None:
    text = p.text
    edits = []
    for fid, it, value in items:
        if it.get("control") is not None:
            return None
        kind = it["kind"]
        if kind == "cell":
            if text.strip(" 　  "):
                return None
            edits.append((0, len(text), str(value)))
        elif kind == "checkbox":
            a, b = it["span"]
            on = value is True
            orig = text[a:b]
            edits.append((a, b, ("☑" if orig == "☐" else "■") if on else orig))
        else:
            a, b = it["span"]
            edits.append((a, b, str(value)))
    out = text
    for a, b, v in sorted(edits, reverse=True):
        out = out[:a] + v + out[b:]
    return out


def run(blob: bytes, filename: str = "") -> Audit:
    doc = DocFile(blob, filename)
    win = build(doc)
    au = Audit(filename)
    raw = _count(raw_chars(doc))
    shown = Counter()
    for st in win["stories"]:
        shown += _count(st["_plain"])
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    au.checks.append(Check("読み落とし（文書の文字）", not missing and not extra,
                           f"本文・表・ヘッダー/フッター・脚注・コメント・テキストボックスの字 {sum(raw.values()):,} 字を、"
                           "窓とは別の読み方（ピースの表から直接）で拾い、すべて窓に出ていることを確かめました", missing + extra))
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000, f"Markdown {sum(len(t) for _, t in parts):,} 文字（AIが1回で読める目安 15万文字以内）"))
    from ..docx.audit import macro_check, _sample

    targets, protected = {}, []
    for f in win["fields"]:
        if f.get("readonly"):
            protected.append(f["id"])
            continue
        targets[f["id"]] = _sample(f)
    try:
        res = fill(DocFile(blob, filename), targets)
    except FillError as e:
        au.checks.append(Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30]))
        au.checks.append(macro_check(doc, win))
        return au
    pr = prove(blob, res.blob, {p.id: p.value for p in res.patches})
    note = f"（文書の保護で書けない欄 {len(protected)} か所は除外）" if protected else ""
    au.checks.append(Check("全入力欄への書き込み", pr.ok, f"入力欄 {len(win['fields']):,} か所のうち {len(targets):,} に試しの値を書き込み{note}。{pr.headline()}",
                           [*pr.unexpected, *pr.lint_errors]))
    after = build(DocFile(res.blob, filename))
    got = {f["id"]: f.get("value") for f in after["fields"]}
    written = {p.id: p.value for p in res.patches}
    wrong = []
    in_text = 0
    for k, v in targets.items():
        g = got.get(k)
        w = written.get(k, v)
        if k not in got and pr.ok:
            in_text += 1  # ヘッダーなど本文以外の流れ：書いた値は段落の文字の一部（証明で、段落の文字が期待どおりと確かめた）
            continue
        ok = (bool(g) == v) if isinstance(v, bool) else str(g) == str(w)
        if not ok:
            wrong.append(f"{k}: 書いた値 {w!r} → 窓では {g!r}")
    note2 = f"（うち {in_text} 欄はヘッダーなど本文以外の欄で、段落の文字として読み戻しました）" if in_text else ""
    au.checks.append(Check("読み戻し", not wrong, f"書いた {len(targets):,} 欄のうち {len(targets) - len(wrong):,} 欄が、書き込み後の窓にそのまま出ている{note2}", wrong))
    au.checks.append(macro_check(doc, win))
    return au
