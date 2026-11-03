"""PDFの書き込みの証明：元のバイト列がそのまま先頭に残り、変わったのは書いた欄とその見た目だけであることを確かめる。

- 元のファイルのバイト列が、新しいファイルの先頭とバイト一致する（増分更新）
- 新しい相互参照の各項目が、正しい位置の「n g obj」を指す（修復なしに開ける）
- 書き換えたオブジェクトは、許した変更（ページの注釈の一覧に足す・フォームの欄の値と見た目・欄の一覧に足す・書体を足す）だけ
- 足したオブジェクトは、欄・見た目・書体だけで、どれも書き換えたオブジェクトからたどれる
- 本人の印影・署名の画像を置いたとき（本人が同意したときだけ）：足してよいのは、スタンプの注釈・その見た目・画像（と透け）だけ。
  どれも、元の文書と本人の画像から決め直した置き方（seal.verify）と同じで、画像は本人の画像と一致する
  （JPEG はファイルのバイト列、PNG は開いた画素とアルファ）。確かめていないスタンプ・画像が足されていれば、想定外とする
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from .doc import PdfDocument
from .objects import Lexer, Name, PdfString, Ref, Stream


@dataclass
class PdfProof:
    original_sha256: str
    new_sha256: str
    prefix_ok: bool
    objects_added: int = 0
    objects_changed: int = 0
    unexpected: list[str] = field(default_factory=list)
    lint_errors: list[str] = field(default_factory=list)
    changed_fields: list[str] = field(default_factory=list)
    original_bytes: int = 0
    new_bytes: int = 0
    xfa: dict = field(default_factory=dict)  # XFA の確かめの結果（バイト一致の部品の数・データの変わった所の数）
    xfa_keys: set = field(default_factory=set)  # XFA の確かめで認めた、AcroForm の辞書の変化（datasets を足した /XFA）
    stamps: list = field(default_factory=list)  # 確かめた、置いた画像（印影・署名）の名前
    page_contents: dict = field(default_factory=dict)  # ○を描いたページ → (前に足したストリーム, 後ろに足したストリーム) の番号
    circles: list = field(default_factory=list)  # 確かめた○の名前

    @property
    def ok(self) -> bool:
        return self.prefix_ok and not self.unexpected and not self.lint_errors

    def headline(self) -> str:
        if not self.prefix_ok:
            return "元のバイト列が保たれていません"
        xfa = ""
        if self.xfa:
            xfa = (f"XFA：template・config などの部品 {self.xfa.get('identical_packets', 0)} 個はバイト一致、"
                   f"datasets で変わったのは {self.xfa.get('data_changes', 0)} か所。")
        return xfa + (f"元のファイル {self.original_bytes:,} バイトは先頭からバイト一致のまま、後ろに {self.new_bytes - self.original_bytes:,} バイトを足しました"
                f"（足したオブジェクト {self.objects_added}・書き換えたオブジェクト {self.objects_changed}）。"
                + (("変わったのは書いた欄とその見た目" + ("、置いた画像（本人の画像と一致）" if self.stamps else "") + (f"、○ {len(self.circles)} つ（ページの描画の後ろに足した楕円）" if self.circles else "") + "だけです")
                   if not self.unexpected else f"想定外の変更 {len(self.unexpected)} 件"))

    # 書き込みの証明を、ほかの形式と同じ形で使えるように
    @property
    def entries_total(self) -> int:
        return 1

    @property
    def entries_identical(self) -> int:
        return 1 if self.prefix_ok else 0


ALLOWED_FIELD_KEYS = {"V", "AS", "AP", "DA"}


def _norm(o):
    """比べるための形（参照・名前・文字列を、そのまま比べられる値に）。"""
    if isinstance(o, Stream):
        return ("stream", _norm(o.dict), o.raw)
    if isinstance(o, dict):
        return {k: _norm(v) for k, v in o.items() if k != "Length"}
    if isinstance(o, list):
        return [_norm(x) for x in o]
    if isinstance(o, PdfString):
        return ("s", bytes(o))
    return o


def lint(blob: bytes, password: str | bytes | None = None) -> list[str]:
    """相互参照の各項目が正しいオブジェクトを指しているか（ビューアーが修復なしに開けるか）。"""
    errs = []
    try:
        doc = PdfDocument(blob, password=password)
    except Exception as e:  # noqa: BLE001
        return [f"開けません: {e}"]
    if doc.repaired:
        errs.append("相互参照が壊れていて、修復が必要です")
        return errs
    for num, ent in doc.xref.items():
        if ent[0] != "o":
            continue
        off, gen = ent[1], ent[2]
        m = re.match(rb"\s*(\d+)\s+(\d+)\s+obj", blob[off:off + 40])
        if not m or int(m[1]) != num or int(m[2]) != gen:
            errs.append(f"オブジェクト {num} の位置（{off}）が、そのオブジェクトを指していません")
    if "Root" not in doc.trailer:
        errs.append("文書の入口（Root）がありません")
    try:
        doc.pages()
    except Exception as e:  # noqa: BLE001
        errs.append(f"ページをたどれません: {e}")
    return errs


def prove(original: bytes, new: bytes, allowed_fields: set[str] | None = None,
          anchored: dict[str, dict] | None = None, password: str | bytes | None = None) -> PdfProof:
    """anchored：見出しの字を目印にして書いた欄 → 目印（書いた順）。元の文書で同じ決め方で位置を決め直し、足した欄の位置と照らす。
    password：開くためのパスワードがかかったPDFのとき、そのパスワード（どちらのファイルも、それで開いて比べる）。"""
    pr = PdfProof(hashlib.sha256(original).hexdigest(), hashlib.sha256(new).hexdigest(),
                  new[:len(original)] == original, original_bytes=len(original), new_bytes=len(new))
    if not pr.prefix_ok:
        pr.unexpected.append("元のバイト列が、新しいファイルの先頭と一致しません")
        return pr
    pr.lint_errors = [e for e in lint(new, password) if e not in lint(original, password)]
    a = PdfDocument(original, password=password)
    a.verify_xref()  # 元の相互参照が壊れていれば、頭から読み直した位置で比べる（書き込みも同じ位置を使う）
    b = PdfDocument(new, password=password)
    if a.decryptor is not None or b.decryptor is not None:
        _check_encryption(a, b, pr)
    tail_start = len(original)
    added, changed = [], []
    for num, ent in b.xref.items():
        if ent[0] == "o" and ent[1] >= tail_start:
            if num in a.xref and a.xref[num][0] != "f":
                changed.append(num)
            else:
                added.append(num)
        elif ent[0] == "s":
            # 圧縮されたオブジェクト：元の文書と同じ置き場所であること
            if a.xref.get(num) != ent:
                changed.append(num)
        elif num in a.xref and a.xref[num] != ent:
            pr.unexpected.append(f"オブジェクト {num} の置き場所が変わっています")
    pr.objects_added = len(added)
    pr.objects_changed = len(changed)
    from .xfaproof import check as check_xfa

    xfa_done = check_xfa(a, b, changed, added, allowed_fields, pr)
    reach: set[int] = set()
    for num in changed:
        old, cur = a.get(num), b.get(num)
        if num in xfa_done:
            _collect(b, cur, reach, set(added))
            continue
        what = _check_change(a, b, num, old, cur, pr)
        if what:
            pr.changed_fields.append(what)
        _collect(b, cur, reach, set(added))
    # 足したオブジェクト：書き換えたものからたどれて、欄・見た目・書体だけ（置いた画像は、下で1つずつ確かめる）
    seal_objs = set()
    for num in added:
        o = b.get(num)
        if num not in reach and not _is_xref(o):
            pr.unexpected.append(f"足したオブジェクト {num} が、どこからも使われていません")
        kind = "xfa" if num in xfa_done else _kind(o)
        if kind is None:
            pr.unexpected.append(f"足したオブジェクト {num} が、欄・見た目・書体のどれでもありません")
        if kind in ("stamp", "image", "content"):
            seal_objs.add(num)
        if kind == "field":
            t = o.get("T")
            if isinstance(t, (bytes, PdfString)):
                from .objects import decode_text

                pr.changed_fields.append(decode_text(t))
    if allowed_fields is not None:
        for name in pr.changed_fields:
            if name and name not in allowed_fields and not any(name.startswith(x + ".") for x in allowed_fields):
                pr.unexpected.append(f"書くと言っていない欄 {name} が変わっています")
    verified: set[int] = set()
    if anchored:
        verified = _check_anchored(a, b, anchored, pr)
    if pr.page_contents and not anchored:
        verified = _check_circles(a, b, [], pr)
    for num in sorted(seal_objs - verified):
        what = "ページの描画" if _kind(b.get(num)) == "content" else "スタンプの注釈か画像"
        pr.unexpected.append(f"足したオブジェクト {num}（{what}）は、置くと言った画像・○として確かめられません")
    return pr


def _check_anchored(a: PdfDocument, b: PdfDocument, anchored: dict[str, dict], pr: PdfProof) -> set[int]:
    """返す：確かめた、置いた画像のオブジェクトの番号（注釈・見た目・画像・透け）。"""
    from ..sealimage import SealError, decode, seal_key
    from . import seal
    from .anchor import AnchorError, resolve_write
    from .form import read_fields
    from .layout import View

    names = {f.name for f in read_fields(a)}
    new_fields = {f.name: f for f in read_fields(b)}
    used: dict = {}
    layouts: dict = {}
    verified: set[int] = set()
    stamp_names = None
    circles = []
    for fid, spec in anchored.items():
        if isinstance(spec, dict) and "circle" in spec:
            circles.append((fid, spec))
            continue
        try:
            skey = seal_key(spec)
        except SealError as e:
            pr.unexpected.append(f"置いた所 {fid}：{e}")
            continue
        if skey is not None:
            # 置いた画像：同意の印・画像のファイルを読み直し、元の文書で置き方を決め直して、足した注釈と画像を照らす
            if stamp_names is None:
                stamp_names = seal.stamp_names(a)
            try:
                if spec.get("consent") is not True:
                    raise SealError("本人の同意（consent: true）の印がありません")
                from pathlib import Path

                img = decode(Path(str(spec.get(skey))).expanduser().read_bytes())
                sd = seal.resolve(a, spec, skey, img, used, layouts, names | stamp_names)
            except (SealError, AnchorError, OSError) as e:
                pr.unexpected.append(f"置いた画像 {fid} の置き方が、元の文書で決まりません：{e}")
                continue
            if sd["id"] != fid:
                pr.unexpected.append(f"置いた画像 {fid} が、決め直した所（{sd['id']}）と違います")
                continue
            probs, nums = seal.verify(b, sd["page"], sd, skey, img)
            pr.unexpected += probs
            if not probs:
                verified |= nums
                pr.stamps.append(fid)
            continue
        try:
            sd = resolve_write(a, spec, used, layouts, names)
        except AnchorError as e:
            pr.unexpected.append(f"書いた所 {fid} の目印が、元の文書で決まりません：{e}")
            continue
        if sd["id"] != fid:
            pr.unexpected.append(f"書いた所 {fid} が、目印から決まる所（{sd['id']}）と違います")
            continue
        fld = new_fields.get(fid)
        ws = [w for w in (fld.widgets if fld else []) if w.page == sd["page"] - 1]
        if not ws:
            pr.unexpected.append(f"書いた所 {fid} の欄が、{sd['page']} ページにありません")
            continue
        view = View(a, a.pages()[sd["page"] - 1])
        want = [round(v, 3) for v in view.back_rect(*sd["rect"])]
        got = [round(float(v), 3) for v in ws[0].rect]
        if any(abs(x - y) > 0.01 for x, y in zip(sorted(want[0::2]) + sorted(want[1::2]), sorted(got[0::2]) + sorted(got[1::2]))):
            pr.unexpected.append(f"書いた所 {fid} の位置 {got} が、目印から決まる位置 {want} と違います")
    if circles or pr.page_contents:
        verified |= _check_circles(a, b, circles, pr, layouts)
    return verified


def _check_circles(a: PdfDocument, b: PdfDocument, circles: list, pr: PdfProof, layouts: dict | None = None) -> set[int]:
    """○：元の文書で、同じ決め方で囲む所を決め直し、足したストリームの中身が、そこから描くものとバイト一致かを確かめる。"""
    from . import circle
    from .anchor import AnchorError
    from .layout import View

    layouts = layouts if layouts is not None else {}
    used: dict = {}
    by_page: dict[int, list] = {}
    for fid, spec in circles:
        try:
            sd = circle.resolve(a, spec, layouts, used)
        except AnchorError as e:
            pr.unexpected.append(f"○ {fid} の所が、元の文書で決まりません：{e}")
            continue
        if sd["id"] != fid:
            pr.unexpected.append(f"○ {fid} が、決め直した所（{sd['id']}）と違います")
            continue
        by_page.setdefault(sd["page"], []).append(sd)
    verified: set[int] = set()
    pages = a.pages()
    nums = {pg["_ref"].num: n for n, pg in enumerate(pages, 1) if pg.get("_ref") is not None}
    for pnum, (pre, post) in pr.page_contents.items():
        n = nums.get(pnum)
        want = by_page.pop(n, None) if n is not None else None
        if not want:
            pr.unexpected.append(f"ページ（オブジェクト {pnum}）の描画が変わっていますが、○を描くと言っていません")
            continue
        page = pages[n - 1]
        expect = circle.ops(View(a, page), want, circle.open_depth(a.page_content(page)))
        got_pre, got_post = b.get(pre), b.get(post)
        if not (isinstance(got_pre, Stream) and b.decode(got_pre) == circle.PREFIX):
            pr.unexpected.append(f"{n} ページの描画の前に足したストリームが、「q」だけではありません")
            continue
        if not (isinstance(got_post, Stream) and b.decode(got_post) == expect):
            pr.unexpected.append(f"{n} ページの描画の後ろに足したストリームが、決め直した○の描き方と違います")
            continue
        verified |= {pre, post}
        pr.circles += [sd["id"] for sd in want]
    for n in by_page:
        pr.unexpected.append(f"{n} ページに○を描くと言いましたが、描画が変わっていません")
    return verified


def _check_encryption(a: PdfDocument, b: PdfDocument, pr: PdfProof) -> None:
    """暗号化された文書：暗号の設定（暗号の辞書・文書の鍵・鍵の元になるファイルの ID の1つめ）が、元のファイルと同じであること。"""
    if a.decryptor is None or b.decryptor is None:
        pr.unexpected.append("暗号のあり・なしが、元のファイルと変わっています")
        return
    if _norm(a.resolve(a.trailer.get("Encrypt"))) != _norm(b.resolve(b.trailer.get("Encrypt"))) \
            or a.decryptor.key != b.decryptor.key:
        pr.unexpected.append("暗号の設定（暗号の辞書・文書の鍵）が、元のファイルと変わっています")
    ia, ib = a.resolve(a.trailer.get("ID")), b.resolve(b.trailer.get("ID"))
    if _norm(ia[0] if isinstance(ia, list) and ia else None) != _norm(ib[0] if isinstance(ib, list) and ib else None):
        pr.unexpected.append("ファイルの ID（暗号の鍵の元）が、元のファイルと変わっています")


def _is_xref(o) -> bool:
    return isinstance(o, Stream) and getattr(o.dict.get("Type"), "name", "") == "XRef"


def _kind(o) -> str | None:
    if _is_xref(o):
        return "xref"
    if isinstance(o, Stream):
        st = getattr(o.dict.get("Subtype"), "name", "")
        if not st and "Type" not in o.dict:
            return "content"  # ページの描画のストリーム（○を描いたとき。中身は○の確かめで照らす）
        return "appearance" if st == "Form" else "image" if st == "Image" else None
    if isinstance(o, dict):
        t = getattr(o.get("Type"), "name", "")
        st = getattr(o.get("Subtype"), "name", "")
        if st == "Stamp" and t == "Annot":
            return "stamp"
        if st == "Widget" or "FT" in o or ("T" in o and "Kids" in o):
            return "field"
        if t in ("Font", "FontDescriptor"):
            return "font"
        if "Fields" in o:
            return "acroform"
        return None
    if isinstance(o, list):
        return "array" if all(isinstance(x, Ref) for x in o) else None
    return None


def _collect(doc, o, reach: set[int], allowed: set[int], depth: int = 0) -> None:
    if depth > 30:
        return
    if isinstance(o, Ref):
        if o.num in allowed and o.num not in reach:
            reach.add(o.num)
            _collect(doc, doc.get(o.num), reach, allowed, depth + 1)
        return
    if isinstance(o, Stream):
        _collect(doc, o.dict, reach, allowed, depth + 1)
    elif isinstance(o, dict):
        for k, v in o.items():
            if k in ("P", "Parent"):
                continue
            _collect(doc, v, reach, allowed, depth + 1)
    elif isinstance(o, list):
        for v in o:
            _collect(doc, v, reach, allowed, depth + 1)


def _check_change(a, b, num, old, cur, pr: PdfProof) -> str | None:
    """書き換えたオブジェクトの差が、許した変更だけかを確かめる。書き換えた欄の名前（あれば）を返す。"""
    if isinstance(old, list) and isinstance(cur, list):
        if _norm(cur[:len(old)]) != _norm(old):
            pr.unexpected.append(f"オブジェクト {num}（一覧）：元の項目が変わっています（足すだけのはず）")
        return None
    if not isinstance(old, dict) or not isinstance(cur, dict):
        pr.unexpected.append(f"オブジェクト {num} が、辞書でない形で書き換えられています")
        return None
    keys = set(old) | set(cur)
    diff = {k for k in keys if _norm(old.get(k)) != _norm(cur.get(k))}
    t = getattr(old.get("Type"), "name", "")
    if t == "Page" or ("Contents" in old and "MediaBox" in old) or (t == "" and "Annots" in diff and "Contents" in old):
        bad = diff - {"Annots", "Contents"}
        if bad:
            pr.unexpected.append(f"ページ（オブジェクト {num}）の {sorted(bad)} が変わっています")
        if "Contents" in diff:
            # ○を描いたページ：元の描画の前に1つ、後ろに1つのストリームを足しただけ（中身は、○の確かめで照らす）
            oc = old.get("Contents")
            if isinstance(oc, list):
                items = oc
            elif isinstance(oc, Ref) and isinstance(a.resolve(oc), list):
                items = a.resolve(oc)
            else:
                items = [] if oc is None else [oc]
            cc = cur.get("Contents")
            if not (isinstance(cc, list) and len(cc) == len(items) + 2 and _norm(cc[1:-1]) == _norm(items)
                    and isinstance(cc[0], Ref) and isinstance(cc[-1], Ref)):
                pr.unexpected.append(f"ページ（オブジェクト {num}）の描画の一覧で、元の描画が変わっています（前後に足すだけのはず）")
            else:
                pr.page_contents[num] = (cc[0].num, cc[-1].num)
        if "Annots" in diff:
            oa, ca = old.get("Annots"), cur.get("Annots")
            if isinstance(oa, list) and (not isinstance(ca, list) or _norm(ca[:len(oa)]) != _norm(oa)):
                pr.unexpected.append(f"ページ（オブジェクト {num}）の注釈の一覧で、元の注釈が変わっています")
        return None
    if t == "Catalog":
        bad = diff - {"AcroForm"}
        if bad:
            pr.unexpected.append(f"文書の入口（Catalog）の {sorted(bad)} が変わっています")
        return None
    if "Fields" in old or "Fields" in cur:
        bad = diff - {"Fields", "DR", "DA"} - pr.xfa_keys
        if bad:
            pr.unexpected.append(f"フォーム（AcroForm）の {sorted(bad)} が変わっています")
        of, cf = old.get("Fields"), cur.get("Fields")
        if isinstance(of, list) and (not isinstance(cf, list) or _norm(cf[:len(of)]) != _norm(of)):
            pr.unexpected.append("フォームの欄の一覧で、元の欄が変わっています")
        return None
    if "FT" in old or getattr(old.get("Subtype"), "name", "") == "Widget" or "Kids" in old and "T" in old:
        bad = diff - ALLOWED_FIELD_KEYS
        if bad:
            pr.unexpected.append(f"欄（オブジェクト {num}）の {sorted(bad)} が変わっています")
        from .objects import decode_text

        name = decode_text(old.get("T")) if isinstance(old.get("T"), (bytes, PdfString)) else None
        if name is None and isinstance(old.get("Parent"), Ref):
            par = a.get(old["Parent"].num)
            if isinstance(par, dict) and isinstance(par.get("T"), (bytes, PdfString)):
                name = decode_text(par["T"])
        return _full_name(a, num, old) or name
    if "Font" in old or t in ("Font",):
        bad = {k for k in diff if k != "Font"}
        if bad:
            pr.unexpected.append(f"資源の辞書（オブジェクト {num}）の {sorted(bad)} が変わっています")
        of, cf = old.get("Font"), cur.get("Font")
        if isinstance(of, dict) and isinstance(cf, dict):
            for k, v in of.items():
                if _norm(cf.get(k)) != _norm(v):
                    pr.unexpected.append(f"書体 {k} が変わっています（足すだけのはず）")
        return None
    # 書体の一覧そのもの（DR の Font が別オブジェクト）
    if all(isinstance(v, Ref) for v in old.values()) and all(isinstance(v, Ref) for v in cur.values()):
        for k, v in old.items():
            if _norm(cur.get(k)) != _norm(v):
                pr.unexpected.append(f"オブジェクト {num} の項目 {k} が変わっています（足すだけのはず）")
        return None
    pr.unexpected.append(f"オブジェクト {num} の {sorted(diff)} が変わっています")
    return None


def _full_name(doc, num, o) -> str | None:
    from .objects import decode_text

    parts = []
    cur = o
    seen = set()
    while isinstance(cur, dict) and len(parts) < 32:
        t = cur.get("T")
        if isinstance(t, (bytes, PdfString)):
            parts.append(decode_text(t))
        p = cur.get("Parent")
        if not isinstance(p, Ref) or p.num in seen:
            break
        seen.add(p.num)
        cur = doc.get(p.num)
    return ".".join(reversed(parts)) if parts else None
