"""PDFの書き込みの証明：元のバイト列がそのまま先頭に残り、変わったのは書いた欄とその見た目だけであることを確かめる。

- 元のファイルのバイト列が、新しいファイルの先頭とバイト一致する（増分更新）
- 新しい相互参照の各項目が、正しい位置の「n g obj」を指す（修復なしに開ける）
- 書き換えたオブジェクトは、許した変更（ページの注釈の一覧に足す・フォームの欄の値と見た目・欄の一覧に足す・書体を足す）だけ
- 足したオブジェクトは、欄・見た目・書体だけで、どれも書き換えたオブジェクトからたどれる
- ページの描画に足した字（pagedraw）：足したストリームが、足した欄から決め直した描き方とバイト一致で、値と描いた字が合い、
  ページの資源の差は足した XObject の項目だけ。元の描画のストリームは一覧にそのまま残る（○と同じ確かめ）。値の控えの欄は隠してあり、注釈に入っていない
- 文書の構造の木に足した要素（structadd）：足した要素は、足したストリームの番号ごとに 1 つ、その番号だけを持つ /Span で、親の /K に入り、
  ParentTree の一覧のその番号が指す。元の要素の /K・一覧の項目は 1 つも変わらず、並びの間に足しただけ
- 本人の印影・署名の画像を置いたとき（本人が同意したときだけ）：足してよいのは、スタンプの注釈・その見た目・画像（と透け）だけ。
  どれも、元の文書と本人の画像から決め直した置き方（seal.verify）と同じで、画像は本人の画像と一致する
  （JPEG はファイルのバイト列、PNG は開いた画素とアルファ）。確かめていないスタンプ・画像が足されていれば、想定外とする
"""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field

from .doc import PdfDocument
from .objects import Name, PdfString, Ref, Stream


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
    struct_expect: dict = field(default_factory=dict)  # (ページのオブジェクトの番号, 番号) → 足したストリームから決まる範囲（structadd）
    drawn: list = field(default_factory=list)  # ページの描画に描いたと確かめた値の欄の名前
    added_nums: set = field(default_factory=set)  # 足したオブジェクトの番号

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
                + (("変わったのは書いた欄とその見た目" + ("、置いた画像（本人の画像と一致）" if self.stamps else "") + (f"、○ {len(self.circles)} つ（ページの描画の後ろに足した楕円）" if self.circles else "") + (f"、書いた字 {len(self.drawn)} か所（元の描画はそのまま、その後ろに足した描画）" if self.drawn else "") + (f"、文書の構造の木に足した要素 {len(self.struct_expect)}" if self.struct_expect else "") + "だけです")
                   if not self.unexpected else f"想定外の変更 {len(self.unexpected)} 件"))

    # 書き込みの証明を、ほかの形式と同じ形で使えるように
    @property
    def entries_total(self) -> int:
        return 1

    @property
    def entries_identical(self) -> int:
        return 1 if self.prefix_ok else 0


ALLOWED_FIELD_KEYS = {"V", "AS", "AP", "DA", "I", "TI"}  # I・TI：リストの選んだ番号・一番上に見せる行


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
    from ..lazy.filebytes import same_prefix, sha256_hex

    pr = PdfProof(sha256_hex(original), sha256_hex(new), same_prefix(new, original), original_bytes=len(original), new_bytes=len(new))
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
    pr.added_nums = set(added)
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
    _check_states(b, set(changed) | set(added), pr)
    _check_ap_geometry(b, set(changed) | set(added), set(added), pr)
    # 足したオブジェクト：書き換えたものからたどれて、欄・見た目・書体だけ（置いた画像は、下で1つずつ確かめる）
    seal_objs = set()
    # 足した書体の、字の対応（ToUnicode の CMap）のストリーム
    tounicode = {o["ToUnicode"].num for o in (b.get(n) for n in added)
                 if isinstance(o, dict) and _kind(o) == "font" and isinstance(o.get("ToUnicode"), Ref)}
    for num in added:
        o = b.get(num)
        if num not in reach and not _is_xref(o):
            pr.unexpected.append(f"足したオブジェクト {num} が、どこからも使われていません")
        kind = "xfa" if num in xfa_done else ("font" if num in tounicode and isinstance(o, Stream)
                                              and b"begincmap" in b.decode(o) else _kind(o))
        if kind == "appearance":
            _check_glyphs(b, num, o, pr)
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
    _check_drawn_placed(b, pr)
    _check_struct(a, b, pr)
    for num in sorted(seal_objs - verified):
        what = "ページの描画" if _kind(b.get(num)) == "content" else "スタンプの注釈か画像"
        pr.unexpected.append(f"足したオブジェクト {num}（{what}）は、置くと言った画像・○として確かめられません")
    return pr


def _check_states(b: PdfDocument, touched: set[int], pr: PdfProof) -> None:
    """書き換えた・足した欄の、値と見た目の状態がつじつまが合うか（書いた後の文書だけで決まること）：
    - チェック・ラジオ：入（/AS が Off でない）の部品の状態は、どれも値（/V）と同じ名前。値と同じ「入」の名前を持つ部品は入
      （ISO 32000-1 12.7.4.2.3・12.7.4.2.4）
    - リスト：選んだ番号（/I）があれば、小さい順で、その選択肢の書き出しの値が値（/V）と同じ（12.7.4.4）"""
    from .form import read_fields
    from .objects import decode_text

    for f in read_fields(b):
        nums = {w.ref.num for w in f.widgets if w.ref is not None} | ({f.ref.num} if f.ref is not None else set())
        if not nums & touched:
            continue
        if f.ftype == "Btn" and f.kind in ("checkbox", "radio"):
            v = f.value
            v = v.name if isinstance(v, Name) else (decode_text(v) if isinstance(v, (bytes, PdfString)) else None)
            v = v or "Off"
            for w in f.widgets:
                st = w.state or "Off"
                if st != "Off" and st != v:
                    pr.unexpected.append(f"欄 {f.name}：値（/V）は {v} なのに、部品の見た目の状態（/AS）が {st} の入になっています")
                elif st == "Off" and v != "Off" and v in w.states:
                    pr.unexpected.append(f"欄 {f.name}：値（/V）は {v} なのに、その状態の部品が切（Off）のままです")
        if f.ftype == "Ch" and f.ref is not None:
            node = b.resolve(f.ref)
            idx = b.resolve(node.get("I")) if isinstance(node, dict) else None
            if idx is None:
                continue
            idx = [b.resolve(x) for x in idx] if isinstance(idx, list) else idx
            if not isinstance(idx, list) or not all(isinstance(i, int) and 0 <= i < len(f.options) for i in idx) \
                    or idx != sorted(set(idx)):
                pr.unexpected.append(f"欄 {f.name}：選んだ番号（/I）{idx} が、選択肢の番号の小さい順の並びではありません")
                continue
            vals = f.value_text()
            vals = vals if isinstance(vals, list) else ([] if vals in (None, "") else [vals])
            if [f.options[i][0] for i in idx] != [str(x) for x in vals]:
                pr.unexpected.append(f"欄 {f.name}：選んだ番号（/I）{idx} の選択肢と、値（/V）{vals} が合いません")


def _check_ap_geometry(b: PdfDocument, touched: set[int], added: set[int], pr: PdfProof) -> None:
    """書いた・足した欄と注釈の、足した見た目（/AP の N）が、欄の四角（/Rect）にそのままの大きさ・決まった向きで描かれるか。

    ビューアーは見た目のフォームを、BBox を Matrix で写した四角が Rect に重なるように、縦と横を別々の倍率で伸び縮みさせて描く
    （ISO 32000-1 12.5.5 Algorithm 8.1）。BBox を Matrix で写した四角の大きさが Rect と違えば、字がつぶれたり点になったりする
    ので、想定外とする。Matrix の回す向きは、欄なら /MK の R（反時計回りの角度、表 189）、スタンプ（置いた画像）ならページの
    /Rotate を打ち消す向き（見た目の向きで描く）と同じであること（ほかの向きでは、字や画像が横倒し・逆さに見える）。"""
    import math

    from .fill import norm_rot

    page_rot = {}
    for pg in b.pages():
        if pg.get("_ref") is not None:
            page_rot[pg["_ref"].num] = norm_rot(b.resolve(pg.get("Rotate")) or 0)

    def nums_of(v):
        try:
            return [float(b.resolve(x)) for x in b.resolve(v)]
        except (TypeError, ValueError):
            return None

    for num in sorted(touched):
        o = b.get(num)
        if not isinstance(o, dict) or "Rect" not in o:
            continue
        apd = b.resolve(o.get("AP"))
        n = apd.get("N") if isinstance(apd, dict) else None
        refs = []
        if isinstance(n, Ref):
            nd = b.resolve(n)
            if isinstance(nd, Stream):
                refs = [("", n)]
            elif isinstance(nd, dict):
                n = nd
        if isinstance(n, dict) and not isinstance(n, Stream):
            refs = [(k, v) for k, v in n.items() if isinstance(v, Ref)]
        refs = [(k, r) for k, r in refs if r.num in added]
        if not refs:
            continue
        sub = getattr(o.get("Subtype"), "name", "")
        if sub == "Stamp":
            p = o.get("P")
            want = page_rot.get(p.num, 0) if isinstance(p, Ref) else 0
            what = f"置いた画像（オブジェクト {num}）"
        else:
            mk = b.resolve(o.get("MK"))
            want = norm_rot(b.resolve(mk.get("R"))) if isinstance(mk, dict) else 0
            from .objects import decode_text

            t = o.get("T")
            what = f"欄 {decode_text(t)}" if isinstance(t, (bytes, PdfString)) else f"欄（オブジェクト {num}）"
        rect = nums_of(o.get("Rect"))
        if not rect or len(rect) != 4:
            pr.unexpected.append(f"{what}：位置（/Rect）を読めません")
            continue
        rw, rh = abs(rect[2] - rect[0]), abs(rect[3] - rect[1])
        for state, ref in refs:
            st = b.resolve(ref)
            label = f"{what}の見た目（オブジェクト {ref.num}{'・状態 ' + state if state else ''}）"
            if not isinstance(st, Stream):
                pr.unexpected.append(f"{label}が、フォームのストリームではありません")
                continue
            bbox = nums_of(st.dict.get("BBox"))
            m = nums_of(st.dict.get("Matrix")) if st.dict.get("Matrix") is not None else [1.0, 0.0, 0.0, 1.0, 0.0, 0.0]
            if not bbox or len(bbox) != 4 or not m or len(m) != 6:
                pr.unexpected.append(f"{label}：BBox か Matrix を読めません")
                continue
            a_, b_, c_, d_ = m[:4]
            ang = math.radians(want)
            co, si = round(math.cos(ang)), round(math.sin(ang))
            if max(abs(a_ - co), abs(b_ - si), abs(c_ + si), abs(d_ - co)) > 1e-6:
                pr.unexpected.append(f"{label}：Matrix {[round(x, 3) for x in m[:4]]} が、{want} 度回す向き"
                                     f"（{'/MK の R' if sub != 'Stamp' else 'ページの /Rotate を打ち消す向き'}）ではありません"
                                     "（字や画像が横倒し・逆さ・ゆがんで見えます）")
                continue
            pts = [(a_ * x + c_ * y + m[4], b_ * x + d_ * y + m[5]) for x in (bbox[0], bbox[2]) for y in (bbox[1], bbox[3])]
            tw = max(p[0] for p in pts) - min(p[0] for p in pts)
            th = max(p[1] for p in pts) - min(p[1] for p in pts)
            tol_w, tol_h = 0.01 + 1e-4 * rw, 0.01 + 1e-4 * rh
            if tw <= 0 or th <= 0 or abs(tw - rw) > tol_w or abs(th - rh) > tol_h:
                sx = rw / tw if tw > 0 else float("inf")
                sy = rh / th if th > 0 else float("inf")
                pr.unexpected.append(f"{label}：BBox を Matrix で写した大きさ {tw:.2f}×{th:.2f} が、欄の大きさ {rw:.2f}×{rh:.2f} と"
                                     f"違います。ビューアーは横 {sx:.2f} 倍・縦 {sy:.2f} 倍に伸び縮みさせて描くので、字がつぶれて"
                                     "見えます（ISO 32000-1 12.5.5 Algorithm 8.1）")


def _check_glyphs(b: PdfDocument, num: int, o: Stream, pr: PdfProof) -> None:
    """足した見た目のストリームが描く字が、どれも、その書体で描ける字形か（/V だけでなく、見た目そのものを確かめる）。"""
    from .glyphs import appearance_text

    _text, probs = appearance_text(b, o)
    for p in probs:
        pr.unexpected.append(f"見た目（オブジェクト {num}）：{p}")


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
    from . import pagedraw, structadd
    from .glyphs import appearance_text

    new_pages = b.pages()
    for pnum, (pre, post) in pr.page_contents.items():
        n = nums.get(pnum)
        want = by_page.pop(n, None) if n is not None else None
        items = pagedraw.items_of(b, new_pages[n - 1], pr.added_nums) if n is not None else []
        if not want and not items:
            pr.unexpected.append(f"ページ（オブジェクト {pnum}）の描画が変わっていますが、○を描くと言っていません")
            continue
        want = want or []
        page = pages[n - 1]
        old_res = a.resolve(page.get("Resources")) or {}
        old_xo = a.resolve(old_res.get("XObject")) or {}
        may = pagedraw.may_draw(a)
        first = structadd.first_mcid(a, page)[0] if may else None
        new_res_ = b.resolve(new_pages[n - 1].get("Resources")) or {}
        new_xo_ = b.resolve(new_res_.get("XObject")) or {}

        def said(sd, key, new_xo_=new_xo_):
            # ○の透明な字：足した書体で、囲んだ字を言うフォームだけ。なければ、説明なしの形で照らす（描ける書体がなかったとき）
            ref = new_xo_.get(key)
            if not isinstance(ref, Ref) or ref.num not in pr.added_nums:
                return None
            st = b.resolve(ref)
            if not pagedraw.is_said_form(b, st) or appearance_text(b, st)[0] != circle.description(sd):
                pr.unexpected.append(f"○ {sd['id']}：足した透明な字が、囲んだ字と違います")
            return ref

        expect, names, seq = pagedraw.page_ops(View(a, page), circle.open_depth(a.page_content(page)), want, items, old_xo, first, may, said)
        for e in seq:
            pr.struct_expect[(pnum, e["mcid"])] = e
        if not _resources_only_added(a, b, old_res, b.resolve(new_pages[n - 1].get("Resources")) or {}, names):
            pr.unexpected.append(f"{n} ページの資源（/Resources）が、足した描画の分のほかにも変わっています")
            continue
        got_pre, got_post = b.get(pre), b.get(post)
        if not (isinstance(got_pre, Stream) and b.decode(got_pre) == circle.PREFIX):
            pr.unexpected.append(f"{n} ページの描画の前に足したストリームが、「q」だけではありません")
            continue
        if not (isinstance(got_post, Stream) and b.decode(got_post) == expect):
            pr.unexpected.append(f"{n} ページの描画の後ろに足したストリームが、決め直した○の描き方と違います")
            continue
        verified |= {pre, post}
        pr.circles += [sd["id"] for sd in want]
        pr.drawn += [it[0] for it in items]
    for n in by_page:
        pr.unexpected.append(f"{n} ページに○を描くと言いましたが、描画が変わっていません")
    return verified


def _resources_only_added(a, b, old: dict, new: dict, names: dict) -> bool:
    """ページの資源の差が、足した描画の XObject の項目だけか。"""
    if {k: _norm(v) for k, v in old.items() if k != "XObject"} != {k: _norm(v) for k, v in new.items() if k != "XObject"}:
        return False
    oxo, nxo = a.resolve(old.get("XObject")) or {}, b.resolve(new.get("XObject")) or {}
    want = dict(oxo)
    want.update(names)
    return _norm(want) == _norm(nxo)


def _check_drawn_placed(b: PdfDocument, pr: PdfProof) -> None:
    """足した「値をページに描いた欄」が、どれも実際にページに描かれ、値と描いた字が合い、隠す印だけで隠してあるか。"""
    from . import pagedraw

    done = set(pr.drawn)
    for num in sorted(pr.added_nums):
        o = b.get(num)
        if not isinstance(o, dict) or pagedraw.widget_drawn(b, o) is None:
            continue
        t = o.get("T")
        name = decode_field_name(t)
        why = pagedraw.value_problem(b, o)
        if why:
            pr.unexpected.append(f"欄 {name}：{why}")
        if b.resolve(o.get("F")) != pagedraw.ANNOT_HIDDEN:
            pr.unexpected.append(f"欄 {name}：値の控えの欄が、隠す印だけになっていません")
        if any(isinstance(x, Ref) and x.num == num for pg in b.pages() for x in (b.resolve(pg.get("Annots")) or [])):
            pr.unexpected.append(f"欄 {name}：値の控えの欄が、ページの注釈に入っています")
        if name not in done:
            pr.unexpected.append(f"欄 {name}：値を描いたはずのページの描画に、描かれていません")


def _check_struct(a: PdfDocument, b: PdfDocument, pr: PdfProof) -> None:
    """足した構造の要素：足したストリームの番号ごとに 1 つ、その番号だけを持ち、親の /K に入り、ParentTree の一覧のその番号が指す。
    Alt・ActualText は持たない。"""
    from . import structadd

    tree = structadd.Tree(b)
    seen: set = set()
    for num in sorted(pr.added_nums):
        o = b.get(num)
        if not (isinstance(o, dict) and getattr(o.get("Type"), "name", "") == "StructElem"):
            continue
        s = getattr(o.get("S"), "name", "")
        if s == "Document":
            old = structadd.tree_of(a)
            if not (isinstance(o.get("P"), Ref) and old.ref is not None and o["P"].num == old.ref.num):
                pr.unexpected.append(f"足した文書の要素（オブジェクト {num}）が、木の根の下にありません")
            continue
        k = b.resolve(o.get("K"))
        pg = o.get("Pg")
        e = pr.struct_expect.get((pg.num, k)) if isinstance(pg, Ref) and isinstance(k, int) else None
        if s != "Span" or e is None:
            pr.unexpected.append(f"足した構造の要素（オブジェクト {num}）が、足した描画の番号の範囲として確かめられません")
            continue
        seen.add((pg.num, k))
        if o.get("ActualText") is not None:
            pr.unexpected.append(f"足した構造の要素（オブジェクト {num}）に ActualText があります")
        if o.get("Alt") is not None:
            pr.unexpected.append(f"足した構造の要素（オブジェクト {num}）に Alt があります")
        page = b.get(pg.num)
        key = b.resolve(page.get("StructParents")) if isinstance(page, dict) else None
        arr = b.resolve(tree.entries.get(key)) if isinstance(key, int) else None
        if not (isinstance(arr, list) and 0 <= k < len(arr) and isinstance(arr[k], Ref) and arr[k].num == num):
            pr.unexpected.append(f"足した構造の要素（オブジェクト {num}）を、ページの一覧（ParentTree）の番号 {k} が指していません")
        par = b.resolve(o.get("P"))
        kids = b.resolve(par.get("K")) if isinstance(par, dict) else None
        kids = kids if isinstance(kids, list) else [par.get("K") if isinstance(par, dict) else None]
        if not any(isinstance(x, Ref) and x.num == num for x in kids):
            pr.unexpected.append(f"足した構造の要素（オブジェクト {num}）が、親の要素の /K に入っていません")
        if not tree.chain(o.get("P") if isinstance(o.get("P"), Ref) else None) and getattr(o.get("P"), "num", None) != (tree.ref.num if tree.ref else None):
            pr.unexpected.append(f"足した構造の要素（オブジェクト {num}）の親から、木の根へたどれません")
    for k in sorted(set(pr.struct_expect) - seen):
        pr.unexpected.append(f"足した描画の番号 {k[1]}（{pr.struct_expect[k]['id']}）の要素が、構造の木にありません")


def _only_inserted(old: list, cur: list, added: set[int]) -> bool:
    """cur が、old の項目をそのままの順に持ち、ほかは足したオブジェクトへの参照（か、一覧の埋め草の null）だけか。"""
    rest = [x for x in cur if not (isinstance(x, Ref) and x.num in added) and x is not None]
    return _norm(rest) == _norm([x for x in old if x is not None])


def _check_struct_change(a, b, num, old: dict, cur: dict, diff: set, pr: PdfProof) -> bool:
    """構造の木の元のオブジェクトの書き換え：足すだけか。構造の木のオブジェクトなら True。"""
    t = getattr(old.get("Type"), "name", "")
    if t == "StructTreeRoot":
        bad = diff - {"ParentTree", "ParentTreeNextKey", "K"}
        if bad:
            pr.unexpected.append(f"文書の構造の木の根の {sorted(bad)} が変わっています")
        if "K" in diff and not _kids_only_inserted(a, b, old.get("K"), cur.get("K"), pr.added_nums):
            pr.unexpected.append("文書の構造の木の根の要素が変わっています（足すだけのはず）")
        if "ParentTree" in diff and isinstance(old.get("ParentTree"), dict):
            _check_nums(a, b, old["ParentTree"], cur.get("ParentTree"), pr)
        elif "ParentTree" in diff and old.get("ParentTree") is not None:
            pr.unexpected.append("文書の構造の木の一覧（ParentTree）が差し替えられています")
        return True
    if t == "StructElem" or ("S" in old and "P" in old):
        if diff - {"K"}:
            pr.unexpected.append(f"構造の要素（オブジェクト {num}）の {sorted(diff - {'K'})} が変わっています")
        elif not _kids_only_inserted(a, b, old.get("K"), cur.get("K"), pr.added_nums):
            pr.unexpected.append(f"構造の要素（オブジェクト {num}）の中身が変わっています（足すだけのはず）")
        return True
    if "Nums" in old or "Limits" in old:
        if diff - {"Nums", "Limits"}:
            pr.unexpected.append(f"構造の一覧（オブジェクト {num}）の {sorted(diff - {'Nums', 'Limits'})} が変わっています")
        _check_nums(a, b, old, cur, pr)
        return True
    return False


def _kids_only_inserted(a, b, old, cur, added: set[int]) -> bool:
    ol = a.resolve(old) if isinstance(old, Ref) else old
    if isinstance(old, Ref) and isinstance(ol, list) and isinstance(cur, Ref) and cur.num == old.num:
        return True  # /K が別のオブジェクトの一覧：その一覧の書き換えとして照らす（_check_change の一覧の所）
    ol = old if not isinstance(ol, list) else ol
    ol = ol if isinstance(ol, list) else ([] if ol is None else [ol])
    cl = cur if isinstance(cur, list) else ([] if cur is None else [cur])
    return _only_inserted(ol, cl, added)


def _check_nums(a, b, old: dict, cur, pr: PdfProof) -> None:
    """ParentTree の葉：元の鍵の値はそのまま（葉の中に直接ある一覧は、足すだけ）。鍵は足すだけ。"""
    if not isinstance(cur, dict):
        pr.unexpected.append("文書の構造の木の一覧（ParentTree）の形が変わっています")
        return
    on = a.resolve(old.get("Nums")) or []
    cn = b.resolve(cur.get("Nums")) or []
    cm = {cn[i]: cn[i + 1] for i in range(0, len(cn) - 1, 2)}
    for i in range(0, len(on) - 1, 2):
        k, v = on[i], on[i + 1]
        w = cm.get(k)
        if isinstance(v, list) and isinstance(w, list):
            if not _only_inserted(v, w, pr.added_nums):
                pr.unexpected.append(f"文書の構造の木の一覧の鍵 {k} の項目が変わっています（足すだけのはず）")
        elif _norm(w) != _norm(v):
            pr.unexpected.append(f"文書の構造の木の一覧の鍵 {k} が変わっています（足すだけのはず）")
    for k in cm:
        if not any(on[i] == k for i in range(0, len(on) - 1, 2)) and not (isinstance(cm[k], Ref) and cm[k].num in pr.added_nums):
            pr.unexpected.append(f"文書の構造の木の一覧の鍵 {k} が、足した一覧を指していません")


def decode_field_name(t) -> str:
    from .objects import decode_text

    return decode_text(t) if isinstance(t, (bytes, PdfString)) else ""


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
        if "Length1" in o.dict:
            return "font"  # 埋め込んだ書体の中身（FontFile2）
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
        if t == "StructElem" or "Nums" in o:
            return "struct"  # 文書の構造の木に足したもの（_check_struct）
        if "Fields" in o:
            return "acroform"
        return None
    if isinstance(o, list):
        return "array" if all(isinstance(x, Ref) or x is None for x in o) else None
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
    if isinstance(old, Stream) and isinstance(cur, Stream):
        # とどくがページに描いた値を、直して書き直した（描く XObject の中身だけ）
        if num in _drawn_nums(b):
            _check_glyphs(b, num, cur, pr)
            return None
        pr.unexpected.append(f"オブジェクト {num}（ストリーム）が書き換えられています")
        return None
    if isinstance(old, list) and isinstance(cur, list):
        if _norm(cur[:len(old)]) != _norm(old) and not _only_inserted(old, cur, pr.added_nums):
            pr.unexpected.append(f"オブジェクト {num}（一覧）：元の項目が変わっています（足すだけのはず）")
        return None
    if not isinstance(old, dict) or not isinstance(cur, dict):
        pr.unexpected.append(f"オブジェクト {num} が、辞書でない形で書き換えられています")
        return None
    keys = set(old) | set(cur)
    diff = {k for k in keys if _norm(old.get(k)) != _norm(cur.get(k))}
    t = getattr(old.get("Type"), "name", "")
    if t == "Page" or ("Contents" in old and "MediaBox" in old) or (t == "" and "Annots" in diff and "Contents" in old):
        bad = diff - {"Annots", "Contents", "Resources"}
        if "StructParents" in bad and old.get("StructParents") is None and isinstance(cur.get("StructParents"), int):
            bad = bad - {"StructParents"}  # 文書の構造の木に足した字のある、番号のなかったページ
        if "Resources" in diff and "Contents" not in diff:
            bad = bad | {"Resources"}  # 資源は、足した描画の分だけ（中身は _check_circles で照らす）
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
    if _check_struct_change(a, b, num, old, cur, diff, pr):
        return None
    if t == "Catalog":
        bad = diff - {"AcroForm"}
        if bad:
            pr.unexpected.append(f"文書の入口（Catalog）の {sorted(bad)} が変わっています")
        return None
    if "Fields" in old or "Fields" in cur:
        bad = diff - {"Fields", "DR", "DA"} - pr.xfa_keys
        if "NeedAppearances" in bad and a.resolve(old.get("NeedAppearances")) is True and b.resolve(cur.get("NeedAppearances")) is False:
            bad = bad - {"NeedAppearances"}  # とどくが描いた見た目のまま見せる（fill.clear_need_appearances）
        if bad:
            pr.unexpected.append(f"フォーム（AcroForm）の {sorted(bad)} が変わっています")
        of, cf = old.get("Fields"), cur.get("Fields")
        if isinstance(of, list) and (not isinstance(cf, list) or _norm(cf[:len(of)]) != _norm(of)):
            pr.unexpected.append("フォームの欄の一覧で、元の欄が変わっています")
        return None
    if "FT" in old or getattr(old.get("Subtype"), "name", "") == "Widget" or "Kids" in old and "T" in old:
        bad = diff - ALLOWED_FIELD_KEYS
        if "Rect" in bad and _rect_only_widened(a, old, cur):
            bad = bad - {"Rect"}  # 見えない四角を横に広げただけ（fill._widen_for_redraw）
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


def _rect_only_widened(a: PdfDocument, old: dict, cur: dict) -> bool:
    """欄の四角の違いが、枠も下地もない部品を横に広げただけか（上下は同じで、元の四角を含む）。"""
    ro, rc = a.resolve(old.get("Rect")), cur.get("Rect")
    if not (isinstance(ro, list) and isinstance(rc, list) and len(ro) == 4 and len(rc) == 4):
        return False
    if not all(isinstance(a.resolve(v), (int, float)) for v in ro) or not all(isinstance(v, (int, float)) for v in rc):
        return False
    ro = [a.resolve(v) for v in ro]
    mk = a.resolve(old.get("MK"))
    if isinstance(mk, dict) and any(a.resolve(mk.get(k)) for k in ("BC", "BG", "R")):
        return False
    return (ro[1] == rc[1] and ro[3] == rc[3] and min(rc[0], rc[2]) <= min(ro[0], ro[2]) + 1e-3
            and max(rc[0], rc[2]) >= max(ro[0], ro[2]) - 1e-3)


def _drawn_nums(b: PdfDocument) -> set[int]:
    from . import pagedraw
    from .form import read_fields

    out = set()
    for f in read_fields(b):
        for w in f.widgets:
            wd = b.resolve(w.ref) if w.ref is not None else (b.resolve(f.ref) if f.ref is not None else None)
            ref = pagedraw.widget_drawn(b, wd)
            if ref is not None:
                out.add(ref.num)
    return out


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
