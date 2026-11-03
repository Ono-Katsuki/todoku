"""XFA のフォームに書く：datasets（データ）の、書いた欄に結ばれた所だけを、元のバイト列の小さな差し替えで書き換える。

- 欄に結ばれたデータがあれば、その中身（または属性の値）だけを書き換える
- なければ、結び方（直接の一致・dataRef）が決める所に作る（群の最後に足す。仕様 4章「Re-Normalization」と同じ置き方）。
  同じ名前の欄の何番目かに合わせて、前の空の値も作る
- データがまったくない（空のまま結んだ）フォームでは、くり返しの数が変わらないように、
  いまの Form DOM の名前のある subform の群（空）も作る
- datasets の部品がなければ、新しく作って /XFA の配列の postamble の前に入れる
- 書く前に、欄の決まり（文字数・1字ずつのマス・書式・選択肢・チェックの入切の値・整数や桁）を確かめ、合わなければ断る
"""

from __future__ import annotations

import re
import zlib
from dataclasses import dataclass, field

from .objects import Name, Ref, Stream
from .xfa import XFA_DATA, El, Slot, XfaForm, XField, hint_of
from .xfapicture import canonical_date, canonical_number, conforms, format_value, parse_value


class XfaFillError(ValueError):
    pass




def _text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def normalize(x: XfaForm, f: XField, value, force: bool = False) -> tuple[str | list[str], object, list[str]]:
    """書く値を、データに入れる値にする：(データの値, 窓に出る値, 知らせること)。合わなければ XfaFillError。"""
    notes: list[str] = []
    if not f.writable:
        raise XfaFillError(f"{f.id}: {f.why_not}")
    if f.readonly and not force:
        raise XfaFillError(f"{f.id}: 読み取り専用の欄です")
    if f.calculated in ("error",) and not force:
        raise XfaFillError(f"{f.id}: フォームのスクリプトが計算して決める欄です（Adobe で開くと計算し直され、上書きできません）")
    if f.calculated and f.calculated not in ("error",):
        notes.append(f"{f.id}: スクリプトで計算される欄です。Adobe で開くと、計算し直した値に変わることがあります")
    if f.hidden:
        notes.append(f"{f.id}: 表示されない欄に書きました（スクリプトが表示を切り替える欄かもしれません）")
    loc = x.locale_of(f.f)[1]
    if f.kind == "checkbox":
        # true / false と、XFA の <items> が決めた「入」「切」の値（f.on・f.off）だけ（言葉の意味では決めない）
        from ..inputcheck import InputError, checkbox_value

        try:
            on = checkbox_value(value, tuple(x for x in (f.on,) if x != ""), tuple(x for x in (f.off,) if x != ""), f.id)
        except InputError as e:
            raise XfaFillError(str(e)) from None
        return (f.on if on else f.off), on, notes
    if f.kind == "radio":
        t = _text(value)
        if t == "":
            return "", None, notes
        if t in f.options:
            return t, t, notes
        hits = [o for o, lab in zip(f.options, f.option_labels) if lab == t or lab.strip() == t.strip()]
        if len(hits) == 1:
            return hits[0], hits[0], notes
        raise XfaFillError(f"{f.id}: 「{t}」は選択肢にありません。選べるのは " +
                           "、".join(f"「{o}」（{lab}）" if lab != o else f"「{o}」" for o, lab in zip(f.options, f.option_labels)))
    if f.kind in ("combo", "list"):
        vals = value if isinstance(value, (list, tuple)) else ([v for v in _text(value).split("\n")] if f.multi else [_text(value)])
        out = []
        for v in vals:
            v = _text(v)
            if v == "" and not f.multi:
                out.append("")
            elif v in f.exports:
                out.append(v)
            elif v in f.options:
                out.append(f.exports[f.options.index(v)])
            elif f.text_entry:
                out.append(v)
            else:
                shown = "、".join(f"「{o}」" for o in f.options[:30]) + ("ほか" if len(f.options) > 30 else "")
                raise XfaFillError(f"{f.id}: 「{v}」は選択肢にありません。選べるのは {shown}")
        if f.multi:
            return out, out or None, notes
        return out[0], out[0] or None, notes
    t = _text(value)
    if f.kind == "date":
        if t and canonical_date(t) is None:
            canon = None
            for k in ("ui", "format", "validate"):
                if f.pictures.get(k):
                    canon = parse_value(t, f.pictures[k], "date", loc)
                    if canon:
                        break
            if canon is None:
                raise XfaFillError(f"{f.id}: 日付「{t}」を読めません。YYYY-MM-DD（例 2026-04-01）で書いてください")
            notes.append(f"{f.id}: 「{t}」を、日付の決まった形 {canon} にして書きました")
            t = canon
        elif t:
            t = canonical_date(t).isoformat()
    elif f.kind in ("time", "datetime"):
        if t and not re.fullmatch(r"\d{2}:?\d{2}(:?\d{2}(\.\d{3})?)?([zZ]|[+-]\d{2}(:?\d{2})?)?" if f.kind == "time" else
                                  r"\d{4}-?\d{2}-?\d{2}T\d{2}:?\d{2}(:?\d{2}(\.\d{3})?)?([zZ]|[+-]\d{2}(:?\d{2})?)?", t):
            raise XfaFillError(f"{f.id}: 「{t}」は{'時刻（HH:MM:SS）' if f.kind == 'time' else '日時（YYYY-MM-DDTHH:MM:SS）'}の形ではありません")
    elif f.kind == "number" or f.value_type in ("integer", "decimal", "float"):
        if t and not canonical_number(t):
            canon = None
            for k in ("ui", "format"):
                if f.pictures.get(k):
                    canon = parse_value(t, f.pictures[k], "num", loc)
                    if canon:
                        break
            if canon is None:
                raise XfaFillError(f"{f.id}: 「{t}」は数ではありません。1234.5 のように、区切りなしで書いてください")
            notes.append(f"{f.id}: 「{t}」を、数の決まった形 {canon} にして書きました")
            t = canon
        if t:
            t = t.strip()
            ip, _, fp = t.lstrip("+-").partition(".")
            if f.value_type == "integer" and (fp.strip("0") or "e" in t.lower()):
                raise XfaFillError(f"{f.id}: 整数の欄です（「{t}」）")
            if f.lead is not None and f.lead >= 0 and len(ip.lstrip("0")) > f.lead:
                raise XfaFillError(f"{f.id}: 整数部は {f.lead} 桁までです（「{t}」）")
            if f.value_type == "decimal" and f.frac is not None and f.frac >= 0 and len(fp.rstrip("0")) > f.frac:
                raise XfaFillError(f"{f.id}: 小数部は {f.frac} 桁までです（「{t}」）")
    else:
        if any(ord(c) < 32 and c not in "\t\n" for c in t):
            raise XfaFillError(f"{f.id}: 書けない制御文字が入っています")
        if f.maxchars and len(t) > f.maxchars:
            raise XfaFillError(f"{f.id}: 「{t}」は {len(t)} 字です（この欄は {f.maxchars} 字まで）")
        if f.comb and len(t) > f.comb:
            raise XfaFillError(f"{f.id}: 「{t}」は {len(t)} 字です（この欄は1字ずつ {f.comb} マス）")
        pic = f.pictures.get("validate")
        if t and pic and f.format_test != "disabled" and conforms(t, pic, "text", loc) is False:
            # 書式どおりに書いた字なら、決まった形にして確かめ直す
            canon = parse_value(t, pic, "text", loc)
            if canon is not None and conforms(canon, pic, "text", loc):
                notes.append(f"{f.id}: 「{t}」を、書式 {pic} の決まった形「{canon}」にして書きました")
                t = canon
    pic = f.pictures.get("validate")
    if t and pic and f.format_test != "disabled":
        ok = conforms(t, pic, hint_of(f), loc)
        if ok is False and not force:
            level = "（Adobe では" + ("書けません" if f.format_test == "error" else "注意が出ます") + "）"
            raise XfaFillError(f"{f.id}: 「{t}」は、この欄の確かめの書式 {pic} に合いません{level}")
        if ok is None:
            notes.append(f"{f.id}: 確かめの書式 {pic} は、規則では確かめられません（書いた後に確かめてください）")
    if t == "" and f.required:
        notes.append(f"{f.id}: 必須の欄を空にしました")
    data = t
    bp = f.pictures.get("bind")
    if t and bp:
        fv = format_value(t, bp, hint_of(f), loc)
        if fv is None:
            raise XfaFillError(f"{f.id}: データの書式 {bp} に、規則では整えられません（書きません）")
        data = fv
        back = parse_value(fv, bp, hint_of(f), loc)
        if back is not None:
            t = back  # 窓には、データを決まった形に読み直した値が出る
    return data, (t or None), notes


# ====================================================================== データの書き換え
@dataclass
class New:
    """新しく作るデータの要素。"""
    name: str
    text: str | None = None
    kids: list = field(default_factory=list)


def _esc(s: str) -> str:
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\r", "&#13;")


def _qname(buf: bytes, el: El) -> bytes:
    m = re.match(rb"<([^\s/>]+)", buf[el.s:el.se])
    return m[1]


class DataEdit:
    """datasets の書き換えの計画。終わったら build() で、新しい部品のバイト列を作る。"""

    def __init__(self, x: XfaForm):
        self.x = x
        self.buf = x.xdp.combined
        self.values: dict[int, tuple[El, str]] = {}  # 中身を書き換える要素
        self.attrs: dict[tuple[int, str], tuple[El, str]] = {}  # 属性の値
        self.content: dict[int, tuple[El, list[New]]] = {}  # 中身をまるごと（いくつも選ぶ一覧）
        self.append: dict[int, tuple[El, list[New]]] = {}  # 最後に足す要素
        self.children: dict[int, dict[str, list]] = {}  # 置き場所 → 名前 → 要素（元のもの・新しいもの）
        self.new_data: New | None = None  # xfa:data がないときに作るもの
        self.new_packet = False
        self.written: list[tuple[list, str]] = []  # (道筋, 値)：証明の参考
        ds = x.xdp.parts.get("datasets")
        self.ds = ds
        if ds is None:
            self.new_packet = True

    # ------------------------------------------------------------------ 置き場所
    def _data_parent(self):
        if self.x.data.data_el is not None:
            return self.x.data.data_el
        if self.new_data is None:
            self.new_data = New("xfa:data")
        return self.new_data

    def _kids(self, parent) -> dict[str, list]:
        k = id(parent)
        if k not in self.children:
            names: dict[str, list] = {}
            if isinstance(parent, El):
                for c in parent.elements:
                    names.setdefault(c.tag, []).append(c)
            self.children[k] = names
        return self.children[k]

    def _add(self, parent, new: New) -> None:
        self._kids(parent).setdefault(new.name, []).append(new)
        if isinstance(parent, El):
            self.append.setdefault(id(parent), (parent, []))[1].append(new)
        else:
            parent.kids.append(new)

    def ensure(self, slot: Slot):
        """置き場所の要素（元の El か、新しく作る New）。"""
        if slot.node is not None:
            return slot.node.el
        parent = self.ensure(slot.parent) if slot.parent is not None else self._data_parent()
        if isinstance(parent, El) and parent.ns == XFA_DATA and parent.tag == "data" and slot.parent is not None:
            pass
        same = self._kids(parent).get(slot.name, [])
        while len(same) < slot.index:
            self._add(parent, New(slot.name))
            same = self._kids(parent)[slot.name]
        if len(same) > slot.index:
            hit = same[slot.index]
            if isinstance(hit, El) and slot.kind == "value":
                raise XfaFillError(f"データの {slot.name}[{slot.index}] は、ほかの欄に結ばれているか、形が違います（書きません）")
            return hit
        n = New(slot.name)
        self._add(parent, n)
        return n

    # ------------------------------------------------------------------ 値
    def set_value(self, slot: Slot, text: str) -> None:
        node = slot.node
        if node is not None and node.attr is not None:
            self.attrs[(id(node.el), node.attr)] = (node.el, text)
            self.written.append((node.path(), text))
            return
        target = self.ensure(slot)
        if isinstance(target, New):
            target.text = text
        else:
            if any(isinstance(k, El) for k in target.kids):
                raise XfaFillError(f"データの {target.tag} は中に要素を持つ値（書式つきの文字など）なので、書きません")
            self.values[id(target)] = (target, text)
        self.written.append((slot.path(), text))

    def set_list(self, slot: Slot, values: list[str]) -> None:
        target = self.ensure(slot)
        name = "value"
        if isinstance(target, El) and target.elements:
            name = target.elements[0].tag
        news = [New(name, v) for v in values]
        if isinstance(target, New):
            target.kids = news
        else:
            self.content[id(target)] = (target, news)
        self.written.append((slot.path(), "\n".join(values)))

    def skeleton(self) -> None:
        """データのないフォーム：いまの Form DOM の名前のある subform の群を作る（くり返しの数を変えないため）。"""
        from .xfa import _walk

        for f in _walk(self.x.merge.root):
            if f.kind == "subform" and f.slot is not None and not f.transparent and f.slot.node is None:
                try:
                    self.ensure(f.slot)
                except XfaFillError:
                    pass

    # ------------------------------------------------------------------ バイト列
    def _enc(self, s: str) -> bytes:
        enc = self.x.xdp.encoding
        if enc.replace("-", "") in ("utf8", ""):
            return s.encode("utf-8")
        return s.encode(enc, errors="xmlcharrefreplace")

    def _new_xml(self, n: New) -> str:
        inner = (_esc(n.text) if n.text else "") + "".join(self._new_xml(k) for k in n.kids)
        return f"<{n.name}>{inner}</{n.name}>" if inner else f"<{n.name}/>"

    def build(self) -> tuple[int, int, bytes]:
        """書き換えた後のバイト列：(元の部品の中の始まり, 終わり, 新しいバイト列)。部品の位置は つないだ XDP の中。"""
        buf = self.buf
        edits: list[tuple[int, int, bytes]] = []
        for el, text in self.values.values():
            edits.append(self._value_edit(el, text))
        for el, news in self.content.values():
            body = "".join(self._new_xml(n) for n in news)
            if el.se == el.e:
                edits.append((el.s, el.e, buf[el.s:el.se - 2].rstrip() + b">" + self._enc(body) + b"</" + _qname(buf, el) + b">"))
            else:
                edits.append((el.se, el.es, self._enc(body)))
        for (_k, attr), (el, text) in self.attrs.items():
            edits.append(self._attr_edit(el, attr, text))
        for el, news in self.append.values():
            body = self._enc("".join(self._new_xml(n) for n in news))
            if el.se == el.e:
                edits.append((el.s, el.e, buf[el.s:el.se - 2].rstrip() + b">" + body + b"</" + _qname(buf, el) + b">"))
            else:
                edits.append((el.es, el.es, body))
        if self.new_data is not None and self.ds is not None:
            pre = _qname(buf, self.ds).split(b":")[0] if b":" in _qname(buf, self.ds) else None
            nd = New((pre.decode() + ":data") if pre else "xfa:data", kids=self.new_data.kids)
            xml = self._new_xml(nd)
            if not pre:
                xml = xml.replace("<xfa:data", f'<xfa:data xmlns:xfa="{XFA_DATA}"', 1)
            el = self.ds
            body = self._enc(xml)
            if el.se == el.e:
                edits.append((el.s, el.e, buf[el.s:el.se - 2].rstrip() + b">" + body + b"</" + _qname(buf, el) + b">"))
            else:
                edits.append((el.es, el.es, body))
        edits.sort(key=lambda e: (e[0], e[1]))
        for a, b in zip(edits, edits[1:]):
            if b[0] < a[1]:
                raise XfaFillError("データの書き換えが重なりました（書きません）")
        out = bytearray(buf)
        for s, e, rep in reversed(edits):
            out[s:e] = rep
        return 0, len(buf), bytes(out)

    def _value_edit(self, el: El, text: str) -> tuple[int, int, bytes]:
        buf = self.buf
        start = buf[el.s:el.se]
        # xsi:nil="true" があれば、値が空として読まれるので外す
        start2 = re.sub(rb'\s+[A-Za-z_][\w.-]*:nil\s*=\s*(["\'])true\1', b"", start) if b"nil" in start else start
        if el.se == el.e:
            head = start2[:-2].rstrip() + b">"
            return el.s, el.e, head + self._enc(_esc(text)) + b"</" + _qname(buf, el) + b">"
        if start2 != start:
            return el.s, el.es, start2 + self._enc(_esc(text))
        return el.se, el.es, self._enc(_esc(text))

    def _attr_edit(self, el: El, attr: str, text: str) -> tuple[int, int, bytes]:
        buf = self.buf
        start = buf[el.s:el.se]
        local = attr.split("}")[-1].encode()
        m = re.search(rb'(\s(?:[\w.-]+:)?' + re.escape(local) + rb'\s*=\s*)(["\'])(.*?)\2', start, re.S)
        if not m:
            raise XfaFillError(f"属性 {attr} が見つかりません")
        val = self._enc(_esc(text).replace('"', "&quot;"))
        new = start[:m.start(3)] + val + start[m.end(3):]
        return el.s, el.se, new


def write(x: XfaForm, values: dict[str, object], force: bool = False):
    """欄の id → 値。(新しい部品のバイト列, 書いた欄 → (データの値, 窓の値), 知らせること)。"""
    edit = DataEdit(x)
    done: dict[str, tuple[object, object]] = {}
    notes: list[str] = []
    problems: list[str] = []
    seen: dict[tuple, tuple[str, object]] = {}
    for key, value in values.items():
        f = x.by_id.get(key)
        if f is None:
            continue
        try:
            data, shown, n = normalize(x, f, value, force)
            notes += n
            if f.f.slot is not None:
                p = tuple(f.f.slot.path())
                if p in seen and seen[p][1] != data:
                    raise XfaFillError(f"{key}: {seen[p][0]} と同じデータを共有する欄なのに、違う値を書こうとしています"
                                       "（どちらか一方だけ書いてください）")
                seen[p] = (key, data)
            if f.kind == "radio":
                if f.f.slot is not None:
                    edit.set_value(f.f.slot, data)
                else:
                    for m, on in zip(f.members, f.options):
                        if m.slot is not None:
                            off = ""
                            items = m.t.find("items")
                            if items is not None and len(items.kids) > 1:
                                off = items.kids[1].text()
                            edit.set_value(m.slot, on if (data == on and data != "") else off)
            elif f.multi:
                edit.set_list(f.f.slot, data)
            else:
                edit.set_value(f.f.slot, data)
            done[key] = (data, shown)
        except XfaFillError as e:
            problems.append(str(e))
    if problems:
        raise XfaFillError("\n".join(problems))
    if done and x.merge.empty:
        edit.skeleton()
    _s, _e, new = edit.build()
    return new, done, notes, edit


def packet_bytes(x: XfaForm, combined: bytes, edit: DataEdit) -> tuple[str, bytes]:
    """書き換えた XDP から、書き直す部品のバイト列を取り出す。("datasets"|"xdp"|"new", バイト列)。"""
    xdp = x.xdp
    if xdp.mode == "stream":
        return "xdp", combined
    delta = len(combined) - len(xdp.combined)
    if edit.new_packet:
        return "new", b""
    ds = edit.ds
    pk = xdp.packet_of(ds)
    return "datasets", combined[pk.start:pk.end + delta]


def new_datasets(x: XfaForm, edit: DataEdit) -> bytes:
    """datasets の部品がないとき：新しく作る部品。"""
    data = New("xfa:data", kids=edit.new_data.kids if edit.new_data is not None else [])
    xml = f'<xfa:datasets xmlns:xfa="{XFA_DATA}">' + edit._new_xml(data) + "</xfa:datasets>\n"
    return edit._enc(xml)


def stream_for(doc, ref: Ref | None, data: bytes) -> Stream:
    """部品のストリーム：元の辞書（Type など）を保ち、圧縮は FlateDecode にする（暗号のフィルターは保つ）。"""
    old = doc.resolve(ref) if ref is not None else None
    d = dict(old.dict) if isinstance(old, Stream) else {}
    filters = d.get("Filter")
    fl = filters if isinstance(filters, list) else ([filters] if filters is not None else [])
    parms = d.get("DecodeParms")
    pl = parms if isinstance(parms, list) else [parms] * len(fl)
    keep_f, keep_p = [], []
    for fi, pa in zip(fl, pl + [None] * (len(fl) - len(pl))):
        if getattr(doc.resolve(fi), "name", "") == "Crypt":
            keep_f.append(fi)
            keep_p.append(pa)
    d.pop("DecodeParms", None)
    d["Filter"] = keep_f + [Name("FlateDecode")] if keep_f else Name("FlateDecode")
    if keep_f:
        d["DecodeParms"] = keep_p + [None]
    d.pop("Length", None)
    return Stream(d, zlib.compress(data, 6))


