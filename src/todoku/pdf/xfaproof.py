"""XFA の書き込みの証明：datasets を元のファイルと書いた後のファイルで読み直し、変わったのが書いた欄に結ばれたデータだけで、
template・config などほかの部品はバイト一致であることを確かめる。

- 部品（packet）：名前と並びが元のまま（datasets がなかったときは、datasets を1つ足しただけ）。datasets のほかはバイト一致
- datasets：xfa:data の外（一覧のデータなど）は元のまま。xfa:data の中で、値が変わった・足された所は、
  書くと言った欄に、元のファイルで結び方から決まる置き場所だけ。消えた所はない。
  足した群は、その置き場所の上の群か、データのないフォームで Form DOM の subform の群（空）だけ
データの読み直しは、書き込みとは別の読み手（xml.etree）で行う。
"""

from __future__ import annotations

import xml.etree.ElementTree as ET

from .objects import Ref

XFA_DATA = "http://www.xfa.org/schema/xfa-data/1.0/"


def _et(data: bytes, decl: str = "") -> ET.Element:
    from .xfa import repair_xml

    fixed, _ = repair_xml(data)
    try:
        return ET.fromstring(fixed)
    except ET.ParseError:
        wrapped = f"<todoku-wrap {decl}>".encode() + fixed + b"</todoku-wrap>"
        return ET.fromstring(wrapped)[0]


def _ln(tag) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _data_map(data_el: ET.Element | None) -> dict[tuple, tuple[str, dict, bool]]:
    """xfa:data の中の要素：道筋 → (中の字, 属性, 子の要素があるか)。"""
    out: dict[tuple, tuple[str, dict, bool]] = {}
    if data_el is None:
        return out

    def walk(e: ET.Element, path: tuple):
        counts: dict[str, int] = {}
        for c in e:
            if not isinstance(c.tag, str):
                continue
            n = _ln(c.tag)
            i = counts.get(n, 0)
            counts[n] = i + 1
            p = path + ((n, i),)
            if c.tag.startswith("{http://www.w3.org/1999/xhtml}"):
                continue
            attrs = {k: v for k, v in c.attrib.items() if not k.startswith("{http://www.w3.org/2000/xmlns/")}
            text = "".join(t for t in [c.text or ""] + [k.tail or "" for k in c])
            out[p] = (text, attrs, len(c) > 0)
            walk(c, p)
    walk(data_el, ())
    return out


def _split(ds: ET.Element | None):
    if ds is None:
        return None, []
    data, others = None, []
    for c in ds:
        if c.tag == "{" + XFA_DATA + "}data" and data is None:
            data = c
        else:
            others.append(c)
    return data, others


def check(a, b, changed: list[int], added: list[int], allowed: set[str] | None, pr) -> set[int]:
    """XFA の部品の変化を確かめる。確かめたオブジェクトの番号を返す（ほかの確かめ方では見ない）。"""
    from .form import has_xfa
    from .xfa import Xdp, XfaError, read

    if not has_xfa(a):
        return set()
    try:
        xa, xb = Xdp(a), Xdp(b)
    except XfaError as e:
        pr.unexpected.append(f"XFA の部品を読めません：{e}")
        return set()
    handled: set[int] = set()
    na, nb = [p.name for p in xa.packets], [p.name for p in xb.packets]
    ds_a = next((p for p in xa.packets if p.name == "datasets"), None)
    ds_b = next((p for p in xb.packets if p.name == "datasets"), None)
    if xa.mode != xb.mode:
        pr.unexpected.append("XFA の形（配列か、1つのストリームか）が変わっています")
        return handled
    if xa.mode == "array":
        exp = list(na)
        if ds_a is None and ds_b is not None:
            exp.insert(len(exp) - 1, "datasets")
        if nb != exp:
            pr.unexpected.append(f"XFA の部品の並びが変わっています（{na} → {nb}）")
        refs_a = [p.ref for p in xa.packets]
        refs_b = [p.ref for p in xb.packets if not (ds_a is None and p is ds_b)]
        if refs_a != refs_b:
            pr.unexpected.append("XFA の部品のオブジェクトが入れ替わっています")
        same = 0
        for pa, pb in zip([p for p in xa.packets if p.name != "datasets"], [p for p in xb.packets if p.name != "datasets"]):
            if pa.data != pb.data:
                pr.unexpected.append(f"XFA の部品 {pa.name} が変わっています（書き換えてよいのは datasets だけ）")
            else:
                same += 1
        pr.xfa["identical_packets"] = same
        if ds_b is not None and ds_b.ref is not None:
            handled.add(ds_b.ref.num)
        if xb.array_ref is not None and isinstance(xb.array_ref, Ref):
            handled.add(xb.array_ref.num)
        pr.xfa_keys = {"XFA"} if ds_a is None else set()
        before = ds_a.data if ds_a else None
        after = ds_b.data if ds_b else None
    else:
        ea, eb = xa.parts.get("datasets"), xb.parts.get("datasets")
        ra = (ea.s, ea.e) if ea is not None else None
        rb = (eb.s, eb.e) if eb is not None else None
        ca, cb = xa.combined, xb.combined
        if ra and rb:
            ok = ca[:ra[0]] == cb[:rb[0]] and ca[ra[1]:] == cb[rb[1]:]
        elif rb:
            ok = ca == cb[:rb[0]] + cb[rb[1]:]
        else:
            ok = ca == cb
        if not ok:
            pr.unexpected.append("XFA のストリームの、datasets のほかの所が変わっています")
        if xb.stream_ref is not None:
            handled.add(xb.stream_ref.num)
        before = ca[ra[0]:ra[1]] if ra else None
        after = cb[rb[0]:rb[1]] if rb else None
        pr.xfa["identical_packets"] = len(xa.xdp.elements) - 1
    if before == after:
        pr.xfa["data_changes"] = 0
        return handled
    decl = f'xmlns:xfa="{XFA_DATA}"'
    try:
        ta = _et(before, decl) if before else None
        tb = _et(after, decl) if after else None
    except ET.ParseError as e:
        pr.unexpected.append(f"書いた後の datasets を XML として読めません：{e}")
        return handled
    da, oa = _split(ta)
    db, ob = _split(tb)
    if [ET.tostring(x) for x in oa] != [ET.tostring(x) for x in ob]:
        pr.unexpected.append("datasets の中の、xfa:data の外（一覧のデータなど）が変わっています")
    ma, mb = _data_map(da), _data_map(db)
    # 書くと言った欄の置き場所（元のファイルで、結び方から決まる所）
    try:
        x = read(a)
    except XfaError as e:
        pr.unexpected.append(f"元のファイルの XFA を読めません：{e}")
        return handled
    ok_values: dict[tuple, str] = {}
    ok_groups: set[tuple] = set()
    for fid in (allowed or set()):
        f = x.by_id.get(fid)
        if f is None:
            continue
        slots = [f.f.slot] if f.f.slot is not None else [m.slot for m in f.members if m.slot is not None]
        for s in slots:
            p = tuple(s.path())
            ok_values[p] = fid
            for i in range(1, len(p)):
                ok_groups.add(p[:i])
            if f.multi:
                ok_groups.add(p)
    skeleton: set[tuple] = set()
    if x.merge.empty:
        from .xfa import _walk

        for fn in _walk(x.merge.root):
            if fn.kind == "subform" and fn.slot is not None and not fn.transparent:
                skeleton.add(tuple(fn.slot.path()))
    changed_ids: set[str] = set()
    n = 0
    for p, (text, attrs, _kids) in mb.items():
        old = ma.get(p)
        if old is not None:
            if old[0] == text and old[1] == attrs:
                continue
            n += 1
            attr_changes = {k for k in set(old[1]) | set(attrs) if old[1].get(k) != attrs.get(k)}
            attr_ok = all(p + (("@" + k.rsplit("}", 1)[-1], 0),) in ok_values or k.rsplit("}", 1)[-1] == "nil"
                          for k in attr_changes)
            fid = ok_values.get(p) or (ok_values.get(p[:-1]) if p[:-1] in ok_groups and len(p) > 1 else None)
            if (old[0] != text and fid is None and p not in ok_groups) or not attr_ok:
                pr.unexpected.append(f"書くと言っていないデータ {_ps(p)} が変わっています")
            elif fid:
                changed_ids.add(fid)
            for k in attr_changes:
                hit = ok_values.get(p + (("@" + k.rsplit("}", 1)[-1], 0),))
                if hit:
                    changed_ids.add(hit)
            continue
        n += 1
        if p in ok_values:
            changed_ids.add(ok_values[p])
            continue
        if p in ok_groups or p in skeleton or (p[:-1] in ok_groups and len(p) > 1 and any(
                q[:-1] == p[:-1] and q[-1][0] == p[-1][0] and q[-1][1] > p[-1][1] for q in ok_values)) or \
                (len(p) > 1 and p[:-1] in ok_values and x.by_id.get(ok_values[p[:-1]]) is not None
                 and x.by_id[ok_values[p[:-1]]].multi):
            if p not in ok_groups and p not in skeleton and text.strip() and not (len(p) > 1 and p[:-1] in ok_values):
                pr.unexpected.append(f"足した空のはずのデータ {_ps(p)} に、値が入っています")
            continue
        pr.unexpected.append(f"書くと言っていないデータ {_ps(p)} が足されています")
    for p in ma:
        if p not in mb:
            pr.unexpected.append(f"データ {_ps(p)} が消えています")
    pr.xfa["data_changes"] = n
    pr.changed_fields.extend(sorted(changed_ids))
    if allowed is not None:
        for fid in changed_ids:
            if fid not in allowed:
                pr.unexpected.append(f"書くと言っていない欄 {fid} のデータが変わっています")
    return handled


def _ps(p: tuple) -> str:
    return ".".join(f"{n}[{i}]" for n, i in p) or "（xfa:data）"
