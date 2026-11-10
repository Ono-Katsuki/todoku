"""配置から見つけた所に書いた字を、ページの描画の後ろに足す（増分の保存）。どのビューアーでも同じ字形が、そのまま見える。

なぜ欄の見た目（/AP）だけにしないか（実物で測った。tests/audit_findings/pdfkit/ に測り方）：
- Apple の PDFKit（Mac のプレビュー・iPhone などの標準の表示）は、文字の欄の /AP を使わず、/V から描き直す。字の幅に余白を足した幅が
  ないと語の切れ目で折り返し、2 行目は欄の外なので描かれない（昭和60 → 昭和、H20 → H2）。字の大きさも欄の高さで変わる
- Quartz（CGContext.drawPDFPage）は注釈をどれも描かない
  ページの描画なら、どちらも、Poppler・pdf.js も、同じように描く

書き方：
- 欄の字の見た目（フォームの XObject）を、ページの描画の後ろのストリームから `Do` で描く（○と同じ、前後に足すストリーム）。
  元のストリームは 1 バイトも変えない。ページの /Resources には、足した XObject の名前（TdkW…）だけを足す
- 足す順は、ページの読む順（上から、同じ行は左から）。後ろに足すので、描画の順で読む読み手（Chromium・pdf.js の字の層）では、
  値は印刷された字の後ろにまとめて読まれる。元のストリームを変えずにこれを見出しの隣に動かす方法はない（タグのある PDF は structadd）
- 値の控えは、同じ名前の欄（/F の Hidden、値 /V と見た目 /AP、ページを /P で指す）を AcroForm の /Fields にだけ置き、ページの /Annots には
  入れない。隠した欄でも注釈に入れると、Mac のプレビューが読み上げの木に出して、同じ値を 2 度読む（実物で測った）。
  ビューアーは注釈の一覧にない欄を描かないが、値は機械で読め、窓・読み戻し・直して書き直すときに使う。/TdkDrawn が値を描いた XObject を指す
- 読むときは、値と描いた字が一致する欄の XObject だけを、印刷された字として読まない（verified_refs）。一致しないものは、
  書類の字として読む（書類の作り手が、この印で字を読み手から隠すことはできない）
"""

from __future__ import annotations

import re
from typing import NamedTuple

from .objects import PdfString, Ref, Stream, decode_text, dumps, text_string

KEY = "TdkDrawn"  # 欄の部品の辞書：その値を描いた、ページの描画の XObject
LABEL = "TdkLabel"  # 欄の辞書：窓の見出しそのまま（空きの位置の印つき。窓で読むときの見出し）。/TU は読み手のために整えた字
PREFIX = "TdkW"
ANNOT_HIDDEN = 2


class Item(NamedTuple):
    """ページに描く値：欄の名前・描く XObject・欄の左下と右上（ページの座標）。"""
    name: str
    ref: Ref
    x: float
    y: float
    x1: float
    y1: float


def reading_order(entries: list, box) -> list:
    """読む順（上から、同じ行は左から）。box(項目) → 見た目の箱。行：先頭の箱と上下が重なる箱。決まった順に並べる（証明も同じ順に）。"""
    rows: list[list] = []
    for e in sorted(entries, key=lambda e: (box(e)[1], box(e)[0])):
        b = box(e)
        if rows and b[1] < box(rows[-1][0])[3]:
            rows[-1].append(e)
        else:
            rows.append([e])
    return [e for row in rows for e in sorted(row, key=lambda e: box(e)[0])]


def view_box(view, it: Item) -> tuple[float, float, float, float]:
    return tuple(view.rect(it.x, it.y, it.x1, it.y1))


def draw_ops(items: list[Item], taken, mcids: dict | None = None) -> tuple[bytes, dict[str, Ref]]:
    """items：描く順の値。taken：ページの /XObject にすでにある名前。mcids：欄の名前 → 印の番号（タグのある PDF）。
    返す：(足すストリームの命令, 足す /XObject の項目)。証明が同じ決め方で照らすので、決まった順に、決まった形で。"""
    out = bytearray()
    names: dict[str, Ref] = {}
    used = set(taken)
    n = 0
    for it in items:
        n += 1
        key = f"{PREFIX}{n}"
        while key in used:
            n += 1
            key = f"{PREFIX}{n}"
        used.add(key)
        names[key] = it.ref
        out += one_op(it, key, (mcids or {}).get(it.name))
    return bytes(out), names


def one_op(it: Item, key: str, mcid: int | None) -> bytes:
    extra = b" /MCID " + str(mcid).encode() if mcid is not None else b""
    return (b"/Span <</TdkField " + dumps(text_string(it.name)) + extra + b">> BDC\nq\n"
            + f"1 0 0 1 {it.x:.3f} {it.y:.3f} cm\n/{key} Do\nQ\nEMC\n".encode())


SAID_PREFIX = "TdkS"
SAID_KEY = "TdkSaid"  # フォームの辞書：○・チェックの印が示す字を透明に描くフォーム（読み手に伝えるための字）
_SAID_OPEN = b"q\n3 Tr\n"


def said_ops(text_ops_of, text: str, w: float, h: float) -> bytes:
    """印が示す字を透明（文字の描き方 3：描かない）に描く命令。枠の中に収まる大きさで（隣の字の所まで伸ばさない）。text_ops_of は fill.text_ops。"""
    s = max(1.0, min(h * 0.8, w / max(1, len(text))))
    return _SAID_OPEN + text_ops_of([text], s, 0.0, max(0.0, (h - s) / 2), s) + b"Q\n"


def is_said_form(doc, st) -> bool:
    """とどくが足した、透明な字だけのフォームか：初めに文字の描き方を 3（描かない）にし、ほかの描き方にも、図形にも、別の部品にもしない。
    これだけを印刷された字として読まない。見える字は、この印があっても書類の字として読む（書類の作り手が字を読み手から隠すことはできない）。"""
    if not isinstance(st, Stream) or st.dict.get(SAID_KEY) is not True:
        return False
    try:
        data = doc.decode(st)
    except Exception:  # noqa: BLE001
        return False
    if not data.startswith(_SAID_OPEN):
        return False
    from .content import tokens

    n_tr = 0
    for kind, val in tokens(data):
        if kind == "o":
            if val == "Tr":
                n_tr += 1
            elif val in ("f", "F", "f*", "S", "s", "B", "B*", "b", "b*", "Do", "sh", "EI", "BI", "gs"):
                return False
    return n_tr == 1


def page_ops(view, depth: tuple[int, int], circles: list[dict], items: list[Item], taken, first_mcid: int | None,
             describe: bool = True, said=None) -> tuple[bytes, dict[str, Ref], list[dict]]:
    """後ろに足すストリーム全部：閉じていない内容の印と q を戻し（depth は circle.open_depth の答え）、○と値を読む順に描く。書き込みと証明が同じものを作る。
    said(sd, 資源の名前) → ○の示す字のフォームの参照か None。describe=False：○を飾りのまま描く（署名のある文書）。
    返す：(命令, 足す /XObject, 番号をつけた範囲 [{id, mcid, rect, kind}])。"""
    from . import circle

    entries: list[tuple] = [("c", sd) for sd in circles] + [("v", it) for it in items]

    def box(e):
        return tuple(e[1]["rect"]) if e[0] == "c" else view_box(view, e[1])

    if describe:
        entries = reading_order(entries, box)  # 署名のある文書の○は、これまでと同じ順・同じ形のまま
    out = bytearray(circle.closing(depth))
    names: dict[str, Ref] = {}
    used = set(taken)
    seq: list[dict] = []
    counters = {PREFIX: 0, SAID_PREFIX: 0}

    def new_key(prefix: str) -> str:
        while True:
            counters[prefix] += 1
            key = f"{prefix}{counters[prefix]}"
            if key not in used:
                used.add(key)
                return key

    mcid = first_mcid
    for e in entries:
        m = mcid
        if e[0] == "c":
            m = m if describe else None
            key = None
            if describe and said is not None and circle.description(e[1]):
                key = new_key(SAID_PREFIX)
                ref = said(e[1], key)
                if ref is None:
                    used.discard(key)
                    counters[SAID_PREFIX] -= 1
                    key = None
                else:
                    names[key] = ref
            out += circle.one(view, e[1], m, key, describe)
            seq_id, kind = e[1]["id"], "circle"
        else:
            key = new_key(PREFIX)
            names[key] = e[1].ref
            out += one_op(e[1], key, m)
            seq_id, kind = e[1].name, "value"
        if m is not None:
            seq.append({"id": seq_id, "mcid": m, "rect": box(e), "kind": kind})
            mcid += 1
    return bytes(out), names, seq


def widget_drawn(doc, wd) -> Ref | None:
    """部品の辞書の、描いた XObject（なければ None）。"""
    v = wd.get(KEY) if isinstance(wd, dict) else None
    return v if isinstance(v, Ref) else None


def origin(doc, wd) -> tuple[float, float]:
    r = [float(doc.resolve(v)) for v in doc.resolve(wd.get("Rect"))]
    return min(r[0], r[2]), min(r[1], r[3])


def records(doc) -> list[tuple[Ref, dict]]:
    """値を描いた欄（値の控え）の、/Fields の順の (部品の参照, 部品)。とどくは /Fields のいちばん上に足す。文書ごとに一度だけ読む。"""
    cached = doc.__dict__.get("_tdk_records")
    if cached is not None:
        return cached
    out: list[tuple[Ref, dict]] = []
    try:
        af = doc.resolve(doc.root.get("AcroForm"))
        for r in (doc.resolve(af.get("Fields")) or []) if isinstance(af, dict) else []:
            wd = doc.resolve(r)
            if isinstance(r, Ref) and isinstance(wd, dict) and widget_drawn(doc, wd) is not None:
                out.append((r, wd))
    except Exception:  # noqa: BLE001  読めないフォームには、控えがないものとして読む（印刷された字として読む）
        out = []
    doc.__dict__["_tdk_records"] = out
    return out


def _on_page(wd: dict, page: dict) -> bool:
    p = wd.get("P")
    return isinstance(p, Ref) and page.get("_ref") == p


def item_of(doc, wd: dict) -> Item:
    t = doc.resolve(wd.get("T"))
    r = [float(doc.resolve(v)) for v in doc.resolve(wd.get("Rect"))]
    return Item(decode_text(t) if isinstance(t, (bytes, PdfString)) else "", widget_drawn(doc, wd),
                min(r[0], r[2]), min(r[1], r[3]), max(r[0], r[2]), max(r[1], r[3]))


def items_of(doc, page: dict, nums: set[int] | None = None) -> list[Item]:
    """値を描いた欄（/Fields の順）。nums があれば、その番号の欄だけ。"""
    return [item_of(doc, wd) for r, wd in records(doc) if _on_page(wd, page) and (nums is None or r.num in nums)]


def value_problem(doc, wd) -> str | None:
    """欄の値と、ページに描いた字が合うか。合わなければ理由。文字の欄は /V と描いた字、チェックは「入」のときに状態の見た目と同じ印。"""
    from .glyphs import appearance_text

    ref = widget_drawn(doc, wd)
    st = doc.resolve(ref) if ref is not None else None
    if not isinstance(st, Stream):
        return "描いた見た目（フォームの XObject）がありません"
    ft = getattr(doc.resolve(wd.get("FT")), "name", "")
    if ft == "Btn":
        state = getattr(doc.resolve(wd.get("AS")), "name", "Off")
        ap = doc.resolve(wd.get("AP"))
        n = doc.resolve(ap.get("N")) if isinstance(ap, dict) else None
        want = doc.resolve(n.get(state)) if isinstance(n, dict) else None
        if not isinstance(want, Stream):
            return f"状態 {state} の見た目がありません"
        if doc.decode(want) != doc.decode(st):
            return f"ページに描いた印が、状態 {state} の見た目と違います"
        return None
    v = doc.resolve(wd.get("V"))
    want = decode_text(v) if isinstance(v, (bytes, PdfString)) else ""
    got, probs = appearance_text(doc, st)
    if probs:
        return "描いた字：" + "／".join(probs)[:200]
    if re.sub(r"\s", "", got) != re.sub(r"\s", "", want):
        return f"値は「{want}」なのに、ページに描いた字は「{got}」です"
    return None


def verified_refs(doc, page: dict) -> set[int]:
    """印刷された字として読まない XObject の番号：このページの、値と描いた字が一致する欄が指すもの。"""
    out: set[int] = set()
    if not isinstance(page, dict):
        return out
    for _r, wd in records(doc):
        try:
            if _on_page(wd, page) and value_problem(doc, wd) is None:
                out.add(widget_drawn(doc, wd).num)
        except Exception:  # noqa: BLE001  読めない欄は、一致を確かめられない：書類の字として読む
            continue
    try:
        res = doc.resolve(page.get("Resources"))
        xo = doc.resolve(res.get("XObject")) if isinstance(res, dict) else None
        for k, v in (xo.items() if isinstance(xo, dict) else ()):
            if k.startswith(SAID_PREFIX) and isinstance(v, Ref) and is_said_form(doc, doc.resolve(v)):
                out.add(v.num)
    except Exception:  # noqa: BLE001
        pass
    return out


def may_draw(doc, sigs=None) -> bool:
    """書いた字をページの描画に足してよいか（文書の事実で決める）：
    - 電子署名のある PDF：署名のあとにページの描画を変えると、署名の確かめで、記入より重い変更と扱われることがある
      （Acrobat の扱いは、ここでは確かめていない）。フォームの欄として足す
    - 権限の設定がページの描画の変更を許していない（暗号の辞書の P のビット 4）：描画は変えず、欄として足して知らせる"""
    if sigs is None:
        from .signatures import read as read_signatures

        sigs = read_signatures(doc)
    if sigs.signed or sigs.docmdp is not None:
        return False
    dec = doc.decryptor
    if dec is not None and any(op == "modify" for op, _what in dec.restrictions()):
        return False
    return True


def clean_label(label: str | None) -> str:
    """欄の説明（/TU）にする字：窓が空きの位置を示す印（［　］）を除き、空白の並びを 1 つの空白に。字は足さない。"""
    s = re.sub(r"［[\s　]*］", " ", label or "")
    return re.sub(r"[\s　]+", " ", s).strip()


_SPAN = re.compile(rb"/Span <</(TdkField|TdkCircle) (\((?:[^()\\]|\\.)*\)|<[0-9A-Fa-f]*>)(.*?)>> BDC(.*?)EMC", re.S)


def spans(doc, page: dict) -> list[dict]:
    """ページの描画にある、とどくが足した印の範囲：{kind, id, mcid, said}。said：透明な字のフォームが描く字。"""
    from .glyphs import appearance_text
    from .objects import Lexer, parse_object

    res = doc.resolve(page.get("Resources"))
    xo = doc.resolve(res.get("XObject")) if isinstance(res, dict) else None
    xo = xo if isinstance(xo, dict) else {}
    out = []
    for m in _SPAN.finditer(doc.page_content(page)):
        d = {"kind": m[1].decode(), "id": decode_text(parse_object(Lexer(m[2], 0))), "mcid": None, "said": None}
        n = re.search(rb"/MCID (\d+)", m[3])
        if n:
            d["mcid"] = int(n[1])
        k = re.search(rb"/(" + SAID_PREFIX.encode() + rb"\d+) Do", m[4])
        if k:
            st = doc.resolve(xo.get(k[1].decode()))
            if is_said_form(doc, st):
                d["said"] = appearance_text(doc, st)[0]
        out.append(d)
    return out


def read_back(doc, patches: list) -> list[dict]:
    """書いた後の文書で、ページに足した印の範囲を読み直す：値を描いた欄・○のそれぞれに範囲があり、○の透明な字は囲んだ字を言い、
    番号のある範囲は構造の木につながり、値を描いた欄はページの注釈に入っていない。"""
    from . import circle, structadd

    out: list[dict] = []
    pages = doc.pages()
    rec = {decode_text(doc.resolve(wd.get("T"))): (r, wd) for r, wd in records(doc)}
    per_page: dict[int, list] = {}
    for p in patches:
        if p.page and ((getattr(p, "kind", "") == "circle") or p.id in rec):
            per_page.setdefault(p.page, []).append(p)
    for n, ps in per_page.items():
        page = pages[n - 1]
        by_id: dict[str, list] = {}
        for s in spans(doc, page):
            by_id.setdefault(s["id"], []).append(s)
        annots = {a.num for a in (doc.resolve(page.get("Annots")) or []) if isinstance(a, Ref)}
        numbered = []
        for p in ps:
            mine = by_id.get(p.id, [])
            if p.kind == "circle":
                if not mine:
                    continue  # 署名のある文書の○は飾りの印のまま（circle.read_back が確かめる）
                want = circle.description(p.circle)
                if want != mine[-1]["said"] and p.circle.get("said_failed") is not True:
                    out.append({"id": p.id, "written": want, "read_back": f"○の字：{mine[-1]['said']}"})
            else:
                if not mine:
                    out.append({"id": p.id, "written": p.value, "read_back": "ページの描画に、この値の範囲がありません"})
                    continue
                if rec[p.id][0].num in annots:
                    out.append({"id": p.id, "written": p.value, "read_back": "値の控えがページの注釈に入っています（読み手が 2 度読みます）"})
            if mine[-1]["mcid"] is not None:
                numbered.append({"id": p.id, "mcid": mine[-1]["mcid"]})
        for why in structadd.read_back(doc, n, page, numbered):
            out.append({"id": why.split("：")[0], "written": "文書の構造の木に足した要素", "read_back": why})
    return out
