"""一太郎（.jtd）：エージェントが「全て正確に読み書きできるか」を確かめる（読み落とし・ストリームの一覧・書き込みの証明・読み戻し）。

読み落とし：窓とは別の読み方で、すべてのストリームの、目録にあるすべての文字のセグメントを、段落・表・枠・組に分けずに
頭から読み（記録と組の印は、その長さで飛ばす）、字を数えて、窓に1字残らず出ているかを確かめる。さらに、複合ファイルの
すべてのストリームが、意味を確かめたもの（文字・書式・位置・図形・文書の情報など）であることを確かめる。知らないストリームが
あれば、字があるかを確かめていないので、そう知らせる。

証明：書き込み後のファイルを、書き込み処理の内部の情報を使わずに読み直して比べる。
同じ字数の上書き：
- 複合ファイルの全体で、変わったバイトは、上書きした字の 2 バイトずつだけ（ほかのストリーム・目録・セクタの表は1バイトも同じ）
- 上書きした所には、書くはずの字が入っている
- 書き込み後の文書を読み直すと、記録・組の形がすべて同じで、どの段落の文字も「元の文字の、上書きした字だけを替えたもの」
字数を変える書き込み・圧縮した文書：proof.py（直すと決めた構造のほかはバイトが同じで、直した構造がそれぞれ思ったとおりか）
"""

from __future__ import annotations

import hashlib
from collections import Counter
from dataclasses import dataclass, field

from ..cfb import CFB
from ..docx.document import iter_paras
from ..docx.fill import FillError, written_of
from ..xlsx.audit import Audit, Check
from .book import SSMG, JtdError, JtdFile, para_units, ssmg, text_units
from .fill import fill, plan
from .window import build, markdown_parts

BOXES = set("□☐☑☒■")
# 意味を確かめたストリーム（名前 → 中身）。字を持つのは「文字」のものだけで、窓に出す
KNOWN = {
    "DocumentText": "本文の文字", "LayoutBoxText": "レイアウト枠の文字", "Header": "柱（ヘッダー・フッター）の文字",
    "PageLayoutStyleHeader": "ページスタイルの柱の文字", "Footnote": "脚注の文字",
    "LayoutBox": "レイアウト枠の位置・枠線", "LayoutBoxTextPositionTables": "レイアウト枠の文字の位置の表",
    "DocumentTextPositionTables": "本文の文字の位置の表", "Frame": "枠の一覧（位置・種類）", "FrameName": "枠の名前の表",
    "HyperLinkFrame": "リンクの枠の表", "Figure": "図形の一覧", "Font": "フォントの名前", "TextLayoutStyle": "文字・段落のスタイル",
    "PageLayoutStyle": "ページのスタイル", "DocumentViewStyles": "表示の設定", "DocumentEditStyles": "編集の設定",
    "DocumentPeripheral": "文書の周辺の設定", "DocumentPeripheralThree": "文書の周辺の設定", "LineMark": "行の印",
    "PageMark": "ページの印", "PaperMark": "用紙の印", "MarkTag": "印の名札", "ReferenceInfo": "参照の情報",
    "RelatedDocuments": "関連する文書の情報", "ThinkingTemplate": "テンプレートの情報", "ItemizationData": "箇条書きの設定",
    "NumberingLink": "段落番号のつながり", "CorrectAdditionalData": "校正の情報（校正した人の名前）",
    "ObjectSheets/SheetInfo": "シートの名前（窓の見出しに出す）", "JSCompDocument": "圧縮して保存した文書（解いて読む）",
    "OutlineTag": "アウトラインの印（16 バイトの数）", "AutoTextInfo": "自動の文字の位置の表（Ssmg。文字のセグメントはない）",
    "\x04JSRV_SummaryInformation": "文書の情報（作成者・日時など）", "\x05SummaryInformation": "文書のプロパティ",
    "\x05DocumentSummaryInformation": "文書のプロパティ",
}
KNOWN_PREFIX = {"FigureData/": "図形の描画", "DocumentMacro/": "マクロの入れ物"}
TEXT_STREAMS = {"DocumentText", "LayoutBoxText", "Header", "PageLayoutStyleHeader", "Footnote"}


def raw_chars(doc: JtdFile) -> Counter:
    """窓とは別の読み方：すべてのストリームの、目録にあるすべての文字のセグメントを、段落・表・枠・組に分けずに頭から読み、
    記録と組の印（0x1C・0x1E から、その長さの分）を飛ばした字を数える（空白と □ などのチェックの印は数えない）。"""
    c: Counter = Counter()
    for name in doc.cfb.streams():
        data = doc.cfb.read(name)
        if data[:8] != SSMG:
            continue
        for sg in ssmg(data, name):
            u = text_units(sg)
            if u is None:
                continue
            i = 0
            while i < len(u):
                x = u[i]
                if x == 0x1C and i + 2 < len(u):
                    i += max(u[i + 2], 1)
                    continue
                if x == 0x1E and i + 1 < len(u):
                    i += max(u[i + 1], 1)
                    continue
                ch = chr(x)
                if x >= 0x20 and not ch.isspace() and ch not in BOXES:
                    c[ch] += 1
                i += 1
    return c


def _count(texts) -> Counter:
    c: Counter = Counter()
    for t in texts:
        for ch in t:
            if not ch.isspace() and ch not in BOXES:
                c[ch] += 1
    return c


def inventory(doc: JtdFile) -> tuple[list[str], list[str]]:
    """(意味を確かめたストリームの説明, 知らないストリーム)。"""
    known, unknown = [], []
    for name in doc.cfb.streams():
        base = name.split("/")[-1]
        desc = KNOWN.get(name) or next((v for k, v in KNOWN_PREFIX.items() if name.startswith(k)), None)
        if desc is None and base == "\x04JSRV_SegmentInformation":
            desc = "ストリームの一覧"
        if desc is None:
            unknown.append(name)
        else:
            known.append(f"{name.replace(chr(4), '').replace(chr(5), '')}：{desc}")
    return known, unknown


def check_reading(doc: JtdFile, win: dict) -> list[Check]:
    raw = raw_chars(doc)
    shown: Counter = Counter()
    for st in win["stories"]:
        shown += _count(st["_plain"])
    missing = [f"「{ch}」が {n} 字、窓に出ていません" for ch, n in (raw - shown).items()]
    extra = [f"「{ch}」が {n} 字、元にないのに窓に出ています" for ch, n in (shown - raw).items()]
    n_seg = len(doc.texts)
    # 字のストリームの中の、字の並び（TextV.01）でないセグメント：形を確かめていないので読めていない（字があるかもしれない）
    odd = []
    for name in TEXT_STREAMS:
        if doc.cfb.has(name):
            for sg in ssmg(doc.cfb.read(name), name):
                if sg.data[:8] != b"TextV.01":
                    odd.append(f"{name}#{sg.sid}（{sg.data[:8]!r}、{len(sg.data):,} バイト）は、形を確かめていないセグメントなので"
                               "読んでいません（ここに字があれば、窓に出ていません）")
    out = [Check("読み落とし（文書の文字）", not missing and not extra and not odd,
                 f"本文・表・レイアウト枠・柱・脚注の字 {sum(raw.values()):,} 字（文字のセグメント {n_seg} 個）を、窓とは別の読み方"
                 "（すべての文字のセグメントを、段落・表・枠に分けずに頭から読む）で拾い、すべて窓に出ていることを確かめました"
                 + (f"。ただし、字のストリームに、形を確かめていないセグメントが {len(odd)} 個あり、読めていません" if odd else ""),
                 missing + extra + odd)]
    known, unknown = inventory(doc)
    out.append(Check("ストリームの一覧", not unknown,
                     f"複合ファイルのストリーム {len(known) + len(unknown)} 個のうち {len(known)} 個は、意味を確かめたもの"
                     "（字を持つのは本文・レイアウト枠・柱・脚注の文字のストリームだけで、それはすべて窓に出しています）"
                     + (f"。知らないストリーム {len(unknown)} 個は、字があるかを確かめていません" if unknown else ""),
                     [f"知らないストリーム：{n!r}" for n in unknown]))
    return out


# ====================================================================== 証明
@dataclass
class JtdProof:
    original_sha256: str
    new_sha256: str
    streams_total: int = 0
    streams_identical: int = 0
    chars_written: int = 0
    bytes_changed: int = 0
    paragraphs_total: int = 0
    paragraphs_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)
    readback: list[str] = field(default_factory=list)  # 読み直した段落が、思っていた文字と違う所
    structures: list[str] = field(default_factory=list)  # 字数を変える書き込み・縮め直しで直した構造（証明で確かめたもの）
    chars_delta: int = 0
    general: bool = False  # 字数を変える書き込み・圧縮した文書の証明（proof.py）

    @property
    def ok(self) -> bool:
        return not self.unexpected and not self.lint_errors and not self.readback

    def headline(self) -> str:
        if self.general:
            return (f"複合ファイルのストリーム {self.streams_identical}/{self.streams_total} がバイト一致。直したのは、"
                    + "／".join(self.structures) + f" だけ（字数 {self.chars_delta:+d}）。段落 {self.paragraphs_total:,} のうち、"
                    f"変わったのは書いた所のある {self.paragraphs_changed} 段落だけで、ほかの段落・字・書式・バイトは元のまま"
                    + ("" if self.ok else f"（想定外 {len(self.unexpected) + len(self.lint_errors) + len(self.readback)} 件）"))
        return (f"複合ファイルのストリーム {self.streams_identical}/{self.streams_total} がバイト一致。ファイル全体で変わったのは、"
                f"上書きした {self.chars_written} 字の {self.bytes_changed} バイトだけ。段落 {self.paragraphs_total:,} のうち、"
                f"変わったのは書いた所のある {self.paragraphs_changed} 段落だけ"
                + ("" if self.ok else f"（想定外 {len(self.unexpected) + len(self.lint_errors) + len(self.readback)} 件）"))


def _file_offsets(cfb: CFB, name: str):
    """ストリームの中の位置 → ファイルの中の位置、を返す関数。"""
    e = cfb._find(name)
    if e.size >= cfb.cutoff:
        chain = cfb.chain(e.start)
        return lambda p: (chain[p // cfb.ssize] + 1) * cfb.ssize + p % cfb.ssize
    minis = []
    s = e.start
    while s < len(cfb.minifat) and len(minis) * cfb.msize < e.size:
        minis.append(s)
        s = cfb.minifat[s]
    root = cfb.chain(cfb.entries[0].start)

    def at(p: int) -> int:
        q = minis[p // cfb.msize] * cfb.msize + p % cfb.msize
        return (root[q // cfb.ssize] + 1) * cfb.ssize + q % cfb.ssize

    return at


def _values_of(written: dict) -> dict:
    """証明に渡された「書いた所 → 値」から、書き込みの指定を作り直す（字を目印にした書き込みは、その目印で）。"""
    out, n = {}, 0
    for fid, w in written.items():
        if isinstance(w, dict) and "@" in w:
            n += 1
            out[f"@{n}"] = w["@"]
        else:
            out[fid] = w
    return out


def prove(original: bytes, new: bytes, written: dict) -> JtdProof:
    pr = JtdProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest())
    a = JtdFile(original)
    try:
        edits, ops, patches, _w = plan(a, _values_of(written))
    except FillError as e:
        pr.unexpected.append(f"元の文書で、書いた所を決め直せません：{e}")
        return pr
    got_ids = {p.id for p in patches}
    if got_ids != set(written):
        pr.unexpected.append(f"書いた所が、元の文書で決め直した所と違います（{sorted(set(written) ^ got_ids)}）")
    if ops or a.compressed:
        from .proof import prove_general

        pr.general = True
        prove_general(pr, original, new, edits, ops)
        return pr
    pr.chars_written = len(edits)
    # 1) ストリーム：変わってよいのは、上書きした字のあるストリームの、その字の 2 バイトずつだけ
    ca, cb = CFB(original), CFB(new)
    sa, sb = ca.streams(), cb.streams()
    if set(sa) != set(sb):
        pr.unexpected.append(f"ストリームの構成が変わりました：{sorted(set(sa) ^ set(sb))}")
        return pr
    want: dict[str, dict[int, int]] = {}  # ストリーム → {位置: 書くはずのバイト}
    for (k, i), ch in edits.items():
        st = a.texts[k]
        off = st.unit_offset(i)
        b2 = ord(ch).to_bytes(2, "big")
        want.setdefault(st.stream, {}).update({off: b2[0], off + 1: b2[1]})
    allowed_file: set[int] = set()
    for name in sa:
        pr.streams_total += 1
        x, y = ca.read(name), cb.read(name)
        if x == y and name not in want:
            pr.streams_identical += 1
            continue
        if len(x) != len(y):
            pr.unexpected.append(f"ストリーム {name!r} の大きさが変わりました（{len(x)} → {len(y)}）")
            continue
        w = want.get(name, {})
        for p in range(len(x)):
            if x[p] != y[p] and p not in w:
                pr.unexpected.append(f"ストリーム {name!r} の位置 {p} のバイトが、上書きした字の外で変わりました")
                break
        for p, v in w.items():
            if y[p] != v:
                pr.unexpected.append(f"ストリーム {name!r} の位置 {p} に、書くはずの字が入っていません")
                break
        fo = _file_offsets(cb, name)
        allowed_file |= {fo(p) for p in w}
    # 2) ファイルの全体：変わったバイトは、上書きした字のバイトだけ（目録・セクタの表・ほかのストリームはそのまま）
    if len(original) != len(new):
        pr.unexpected.append(f"ファイルの大きさが変わりました（{len(original):,} → {len(new):,} バイト）")
    else:
        diff = [p for p in range(len(original)) if original[p] != new[p]]
        pr.bytes_changed = len(diff)
        outside = [p for p in diff if p not in allowed_file]
        if outside:
            pr.unexpected.append(f"上書きした字のほかに、ファイルの {len(outside)} バイトが変わっています（最初は位置 {outside[0]}）")
    # 3) 読み直し：記録・組の形が同じで、どの段落も「元の文字の、上書きした字だけを替えたもの」
    try:
        b = JtdFile(new)
    except (JtdError, Exception) as e:  # noqa: BLE001
        pr.lint_errors.append(f"書き込み後の文書を読めません：{e}")
        return pr
    if len(a.texts) != len(b.texts) or any(len(x.units) != len(y.units) for x, y in zip(a.texts, b.texts)):
        pr.lint_errors.append("書き込み後の文書の、文字のセグメントの数・長さが変わりました")
        return pr
    for x, y in zip(a.texts, b.texts):
        for i, (u, v) in enumerate(zip(x.units, y.units)):
            if u != v and ((x.k, i) not in edits or ord(edits[(x.k, i)]) != v):
                pr.lint_errors.append(f"{x.stream}#{x.sid} の単位 {i} が、上書きした字の外で変わりました")
                break
    pa = [p for _k, _p, bl in a.all_blocks() for p in iter_paras(bl)]
    pb = [p for _k, _p, bl in b.all_blocks() for p in iter_paras(bl)]
    pr.paragraphs_total = len(pa)
    if len(pa) != len(pb):
        pr.readback.append(f"段落の数が変わりました（{len(pa)} → {len(pb)}）")
        return pr
    for n, (p, q) in enumerate(zip(pa, pb), 1):
        us = para_units(a, p)
        expect = "".join(edits.get(u, ch) if u is not None else ch for u, ch in zip(us, p.text))
        if any(u in edits for u in us if u is not None):
            pr.paragraphs_changed += 1
        if q.text != expect:
            pr.readback.append(f"段落 {n}：読み直すと「{q.text[:40]}」で、思っていた「{expect[:40]}」と違います")
    return pr


# ====================================================================== 検査
def _sample(f: dict):
    """試しの値：上書きだけの欄はその字数まで、字数を変えて書ける欄（maxlen なし）は、空白より長くなりうる値。"""
    if f["kind"] == "checkbox":
        return True
    room = f.get("maxlen")
    if room is None:
        return "１２３" if f["kind"] == "inline" else "テスト記入"
    return ("１２" if f["kind"] == "inline" else "テスト")[:room] or "1"


def anchored_check(blob: bytes, filename: str, grow: bool = False) -> Check:
    """字を目印にした書き込み：文書の中に1つしかない「字＋全角の空白 2 字」を探し、空白 2 字を「試し」で上書きして、
    証明と読み直しを確かめる。grow=True のときは、空白 2 字を 4 字の「試しの値」に替える（字数を変える書き込み）。"""
    name = "字数を変える、字を目印にした書き込み" if grow else "字を目印にした書き込み"
    word = "試しの値" if grow else "試し"
    doc = JtdFile(blob, filename)
    paras = [q for _k, _p, bl in doc.all_blocks() for q in iter_paras(bl)]
    full = "\n".join(q.text for q in paras)
    for q in paras:
        t = q.text
        us = para_units(doc, q)
        for i in range(1, len(t) - 2):
            frag = t[i - 1:i + 2]
            if t[i:i + 2] != "　　" or t[i - 1] in " 　" or full.count(frag) != 1:
                continue
            if any(u is None for u in us[i - 1:i + 2]):
                continue
            if grow and doc.texts[us[i][0]].stream != "DocumentText":
                continue
            spec = {"old": frag, "new": frag[0] + word}
            try:
                res = fill(JtdFile(blob, filename), {"@1": spec})
            except FillError as e:
                return Check(name, False, f"「{frag}」の空白を「{word}」に替えられませんでした", [str(e)])
            pr = prove(blob, res.blob, written_of(res.patches))
            after = "\n".join(p.text for _k, _p, bl in JtdFile(res.blob, filename).all_blocks() for p in iter_paras(bl))
            ok = pr.ok and (frag[0] + word) in after
            return Check(name, ok,
                         f"規則で示していない所への書き込み（「{frag}」の空白 2 字を「{word}」に）を {res.patches[0].id} として書き、"
                         f"{pr.headline()}。読み直すと「{frag[0]}{word}」が{'出ています' if ok else '出ていません'}",
                         [*pr.unexpected, *pr.lint_errors, *pr.readback])
    return Check(name, True, "文書の中に1つしかない「字＋全角の空白 2 字」がないので、試していません")


def run(blob: bytes, filename: str = "") -> Audit:
    au = Audit(filename)
    try:
        doc = JtdFile(blob, filename)
    except Exception as e:  # noqa: BLE001
        au.checks.append(Check("開く", False, f"一太郎の文書として読めません：{e}"))
        return au
    win = build(doc)
    au.checks += check_reading(doc, win)
    parts = markdown_parts(win)
    biggest = max(len(t) for _, t in parts)
    au.checks.append(Check("窓の大きさ", biggest <= 150_000, f"Markdown {sum(len(t) for _, t in parts):,} 文字（AIが1回で読める目安 15万文字以内）"))
    targets, blocked = {}, []
    for f in win["fields"]:
        if f.get("readonly"):
            blocked.append(f["id"])
            continue
        targets[f["id"]] = _sample(f)
    rules_blocked = [f["id"] for f in win["fields"] if f.get("readonly") and not f.get("agent")]
    try:
        res = fill(JtdFile(blob, filename), targets)
    except FillError as e:
        au.checks.append(Check("全入力欄への書き込み", False, f"{len(targets)} 欄に書き込めませんでした", str(e).splitlines()[:30]))
        return au
    pr = prove(blob, res.blob, written_of(res.patches))
    grown = len(res.ops)
    au.checks.append(Check("全入力欄への書き込み", pr.ok,
                           f"書ける入力欄 {len(targets):,} か所すべてに試しの値を書き込みました"
                           + (f"（うち {grown} か所は、空のマスに足す・空白より長い値に替える、字数を変える書き込み）" if grown else "")
                           + ("。圧縮して保存した文書なので、中の文書に書いて -lh5- で縮め直しました" if doc.compressed else "")
                           + f"。{pr.headline()}",
                           [*pr.unexpected, *pr.lint_errors]))
    au.checks.append(Check("読み戻し", not pr.readback,
                           f"書き込み後の文書を読み直し、段落 {pr.paragraphs_total:,} のすべてで、文字が「元の文字に、書いた所だけを"
                           "当てたもの」になっていることを確かめました" if not pr.readback else "書き込み後の文書を読み直すと、思っていた文字と違う段落があります",
                           pr.readback))
    whys = sorted({f.get("why_fixed", "") for f in win["fields"] if f.get("readonly") and not f.get("agent")} - {""})
    au.checks.append(Check("書けない所", not rules_blocked,
                           ("規則で見つけた入力欄は、すべて書けます（空白・伏せ字・□ の上書きか、字数を変える書き込み）" if not rules_blocked else
                            f"規則で見つけた入力欄のうち {len(rules_blocked)} か所は、書けません（窓では「変更できない欄」）。理由："
                            + "／".join(whys[:5]))
                           + (f"。規則では決められない所のうち {len(blocked) - len(rules_blocked)} か所も、書けません"
                              if len(blocked) > len(rules_blocked) else ""),
                           rules_blocked[:30]))
    from ..xlsx.audit import agent_places_check

    ap = agent_places_check(win["fields"])
    if ap is not None:
        au.checks.append(ap)
    au.checks.append(anchored_check(blob, filename))
    from .edit import policy

    if policy(doc).ok:
        au.checks.append(anchored_check(blob, filename, grow=True))
    return au
