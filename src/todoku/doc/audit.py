"""古い Word（.doc）：エージェントが「全て正確に読み書きできるか」を確かめる（読み落とし・全欄への書き込み・証明・読み戻し）。

証明：書き込み後の文書を読み直し、
- WordDocument・表・Data 以外のストリームはバイト一致
- 段落の数が同じで、書いた欄のない段落の文字はすべて同じ。書いた欄のある段落の文字は「空欄を値に置き換えたもの」
- 段落の書式（PAPX）は、すべての段落で同じ
- ピースの表が文書の全体をすき間なく覆い、位置の表（段落の区切り・フィールド・しおり など）が並び順を保つ
- 埋め込んだ Excel の表に書いたとき（E1!Sheet1!C4）：ObjectPool の中で変わってよいのはブックと印刷用の絵だけ。
  ブックは Excel の証明、絵は「変わったセルの字だけ」、表示用の絵は Data の後ろに足したものと、その位置を指す
  sprmCPicLocation の値だけが変わったことで確かめる（todoku.embed.verify_doc）
"""

from __future__ import annotations

import hashlib
import struct
from collections import Counter
from dataclasses import dataclass, field

from ..cfb import CFB
from ..docx.document import WINGDINGS, iter_paras
from ..docx.window import Spots, box_mark, build, markdown_parts
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
    embed_note: str = ""  # 埋め込んだ表に書いたときの、確かめたことの要約
    word_document_kept: bool = False  # WordDocument の元のバイトが、FIB の位置の記録のほかは同じ位置にそのまま残った

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        return (f"複合ファイルのストリーム {self.entries_identical}/{self.entries_total} がバイト一致"
                + ("（本文のストリームも、元のバイトは位置と長さの記録のほかそのまま残り、書いた字は後ろに足しただけ）"
                   if self.word_document_kept else "") + f"。段落 {self.paragraphs_total:,} のうち、"
                f"変わったのは書いた欄のある {self.paragraphs_changed} 段落だけ"
                + ("" if not self.unexpected else f"（想定外の変更 {len(self.unexpected)} 件）")
                + (f"。{self.embed_note}" if self.embed_note else ""))


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
    # cbMac：WordDocument のうち意味のあるバイトの数。その後ろは無視される（MS-DOC 2.5.4 FibRgLw97）ので、ストリームより短くてよい
    if doc.lw[0] > len(doc.wd):
        errs.append("WordDocument の長さ（cbMac）が、ストリームの長さを超えています")
    return errs


from ..docx.fill import written_of  # noqa: E402


def _word_document_outside(a: DocFile, b: DocFile) -> str | None:
    """WordDocument のバイトの比べ：元のバイト列は、書き込み後もそのまま同じ位置に残る（新しい字・書式は後ろに足す）。
    変わってよいのは、FIB の位置と長さの記録（FibRgLw97・FibRgFcLcb）と、埋め込んだ表の表示用の絵の位置を指す
    sprmCPicLocation の値（4 バイト。どこを指すかは埋め込みの証明で確かめる）だけ。合えば None、違えば最初に違う所。"""
    x, y = a.wd, b.wd
    if len(y) < len(x):
        return f"ストリームが短くなりました（{len(x):,} → {len(y):,} バイト）"
    fib = (a.lw_pos, a.fclcb_pos + a.cb_fclcb * 8)
    if (b.lw_pos, b.fclcb_pos, b.cb_fclcb) != (a.lw_pos, a.fclcb_pos, a.cb_fclcb):
        return "FIB の作り（記録の数）が変わりました"
    i = 0
    n = len(x)
    while i < n:
        if x[i] == y[i] or fib[0] <= i < fib[1]:
            i += 1
            continue
        # sprmCPicLocation（0x6A03）の値の 4 バイトの中なら、その 4 バイトを飛ばす
        s0 = next((s_ for s_ in range(max(2, i - 3), i + 1) if x[s_ - 2:s_] == b"\x03\x6a" and y[s_ - 2:s_] == b"\x03\x6a"), None)
        if s0 is None:
            return f"元のバイトが {i:,} バイト目で変わっています（FIB の位置の記録・絵の位置のほか）"
        i = s0 + 4
    return None


def prove(original: bytes, new: bytes, written: dict) -> DocProof:
    from ..embed import is_embed_key, verify_doc
    from .word6 import is_word6

    if is_word6(original):
        from .word6 import prove as prove6

        return prove6(original, new, written)
    pr = DocProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    ca, cb_ = CFB(original), CFB(new)
    sa, sb = ca.streams(), cb_.streams()
    a = DocFile(original)
    allowed = {"WordDocument", a.table_name, "Data"}
    try:
        b = DocFile(new)
    except Exception as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません: {e}")
        return pr
    if getattr(a, "ax_refs", None):
        # ActiveX の部品（ObjectPool の入れ物）：書いた部品の、書いた性質だけが変わっていること
        from ..axbin import prove_doc
        from ..cfbwrite import check as cfb_check

        sp0 = Spots(a)
        ids = {id(it["control"].ax): k for k, it in sp0.items.items()
               if it.get("control") is not None and it["control"].source == "activex"}
        plain = {k: v for k, v in written.items() if not (isinstance(v, dict) and "@" in v) and not is_embed_key(k)}
        ax_bad, ax_seen = prove_doc(a, b, plain, ids)
        pr.unexpected += ax_bad
        allowed |= ax_seen
        pr.lint_errors += cfb_check(new)
    embed_written = {k: v for k, v in written.items() if is_embed_key(k)}
    written = {k: v for k, v in written.items() if k not in embed_written}
    if embed_written:
        # 埋め込んだ表：ブックは Excel の証明で、表の絵は「変わったセルの字だけ」で、Data と文字の書式は絵の移動だけで確かめる
        try:
            more, probs, pr.embed_note = verify_doc(a, b, embed_written)
        except Exception as e:  # noqa: BLE001
            more, probs = set(), [f"埋め込みの表への書き込みを確かめられません：{e}"]
        allowed |= more
        pr.unexpected += probs
    for name in sa:
        pr.entries_total += 1
        if name in sb and ca.read(name) == cb_.read(name):
            pr.entries_identical += 1
        elif name not in allowed:
            pr.unexpected.append(f"ストリーム {name} が変わっています")
    where = _word_document_outside(a, b)
    if where is None:
        pr.word_document_kept = True
    else:
        pr.unexpected.append(f"WordDocument：{where}")
    pr.lint_errors += lint(b)
    if len(b.wd) != len(a.wd) and b.lw[0] != len(b.wd):
        pr.lint_errors.append("WordDocument を書き足したのに、意味のあるバイトの数（cbMac）がストリームの長さと合いません")
    pa = [p for _k, _p, bl in a.all_blocks() for p in iter_paras(bl)]
    pb = [p for _k, _p, bl in b.all_blocks() for p in iter_paras(bl)]
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.unexpected.append(f"段落の数が変わっています（{len(pa)} → {len(pb)}）")
        return pr
    sp = Spots(a)
    targets: dict[int, list] = {}
    index = {id(p): i for i, p in enumerate(pa)}
    used: dict = {}
    for fid, value in written.items():
        if isinstance(value, dict) and "@" in value:
            # 字を目印にした書き込み：元の文書で、書き込みと同じ決め方で位置を決め直す
            from ..docx.window import WriteError, resolve_write

            try:
                it = resolve_write(sp, value["@"], used)
            except WriteError as e:
                pr.unexpected.append(f"書いた所 {fid} の目印が、元の文書で決まりません：{e}")
                continue
            if it["id"] != fid or it["value"] != value["value"]:
                pr.unexpected.append(f"書いた所 {fid} が、目印から決まる所（{it['id']}「{it['value']}」）と違います")
                continue
            value = value["value"]
        else:
            it = sp.items.get(fid)
        if it is None:
            pr.unexpected.append(f"書いた欄 {fid} が、元の文書にありません")
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
            edits.append((a, b, box_mark(orig, on)))
        else:
            a, b = it["span"]
            edits.append((a, b, str(value)))
    out = text
    for a, b, v in sorted(edits, reverse=True):
        out = out[:a] + v + out[b:]
    return out


def run(blob: bytes, filename: str = "") -> Audit:
    from .word6 import is_word6

    if is_word6(blob):
        from .word6 import run as run6

        return run6(blob, filename)
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
    from ..axfile import reading_check

    axc = reading_check(getattr(doc, "ax_refs", []), win)  # ActiveX の部品の見出しと文字
    if axc is not None:
        au.checks.append(axc)
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000, f"Markdown {sum(len(t) for _, t in parts):,} 文字（AIが1回で読める目安 15万文字以内）"))
    from ..docx.audit import macro_check, _sample

    targets, protected = {}, []
    for f in win["fields"]:
        if f.get("readonly") or f.get("options_unknown"):
            protected.append(f["id"])
            continue
        targets[f["id"]] = _sample(f)
    try:
        res = fill(DocFile(blob, filename), targets)
    except FillError as e:
        au.checks.append(Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30]))
        au.checks.append(macro_check(doc, win))
        return au
    pr = prove(blob, res.blob, written_of(res.patches))
    note = f"（文書の保護で書けない欄 {len(protected)} か所は除外）" if protected else ""
    au.checks.append(Check("全入力欄への書き込み", pr.ok, f"入力欄 {len(win['fields']):,} か所のうち {len(targets):,} に試しの値を書き込み{note}。{pr.headline()}",
                           [*pr.unexpected, *pr.lint_errors]))
    after = build(DocFile(res.blob, filename))
    got = {f["id"]: f.get("value") for f in after["fields"]}
    written = {p.id: p.value for p in res.patches}
    # 本文の外（ヘッダー・テキストボックスなど）の欄：.doc はしおりを本文にしか置けないので、書いた後の窓の id は欄に結びつかない
    from ..docx.window import Spots

    main_end = doc.ccp["text"]
    outside = {fid for fid, it in Spots(doc).items.items()
               if it.get("para") is not None and getattr(it["para"], "cp0", 0) >= main_end}
    wrong = []
    in_text = 0
    for k, v in targets.items():
        g = got.get(k)
        w = written.get(k, v)
        if (k not in got or k in outside) and pr.ok:
            in_text += 1  # ヘッダーなど本文以外の流れ：書いた値は段落の文字の一部（証明で、段落の文字が期待どおりと確かめた）
            continue
        ok = (bool(g) == v) if isinstance(v, bool) else str(g) == str(w)
        if not ok:
            wrong.append(f"{k}: 書いた値 {w!r} → 窓では {g!r}")
    note2 = f"（うち {in_text} 欄はヘッダーなど本文以外の欄で、段落の文字として読み戻しました）" if in_text else ""
    au.checks.append(Check("読み戻し", not wrong, f"書いた {len(targets):,} 欄のうち {len(targets) - len(wrong):,} 欄が、書き込み後の窓にそのまま出ている{note2}", wrong))
    from ..docx.audit import embeds_check
    from ..xlsx.audit import agent_places_check

    ec = embeds_check(doc, win, DocFile, fill, prove)
    if ec is not None:
        au.checks.append(ec)
    from ..embeds.audit import word_checks

    au.checks += word_checks(doc, win)  # 埋め込みの読み落とし
    ap = agent_places_check(win["fields"])
    if ap is not None:
        au.checks.append(ap)
    if not (doc.protection_enforced and doc.protection):
        from ..docx.audit import anchored_check

        au.checks.append(anchored_check(blob, filename, DocFile, fill, prove, main_end=doc.ccp["text"]))
    au.checks.append(macro_check(doc, win))
    return au
