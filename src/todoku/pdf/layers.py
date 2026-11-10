"""PDF のレイヤー（optional content。ISO 32000-1 8.11）：既定の設定で、画面に見えるか・印刷に出るか。

描く側（render/paint.py）と読む側（content.py・form.py・window.py）が、同じ決まりを使う：
- 既定の設定（OCProperties の D）：BaseState・ON・OFF
- 使い道ごとの切り替え（D の AS）：Event が View / Print のものについて、Category に並ぶ使い道のうち View・Print・Export を、
  グループの Usage の ViewState・PrintState・ExportState で読む（1 つでも OFF なら OFF、そうでなく ON があれば ON）
- まとめの印（OCMD）：VE（And・Or・Not）があればそれ、なければ P（AllOn・AnyOn・AnyOff・AllOff）

決めていないこと（ビューアーの使い方で変わる）：Category の Zoom・Language・User・PageElement、D 以外の設定（Configs）、
ビューアーで人が切り替えた後の状態。
"""

from __future__ import annotations

from dataclasses import dataclass

from .objects import PdfString, decode_text

_STATE_KEY = {"View": "ViewState", "Print": "PrintState", "Export": "ExportState"}


@dataclass(frozen=True)
class State:
    view: bool  # 既定の設定で、画面に見えるか
    print: bool  # 既定の設定で、印刷に出るか
    name: str  # レイヤーの名前（OCMD は、まとめたレイヤーの名前を「・」でつなぐ）

    def both(self, inner: "State") -> "State":
        """このレイヤーの中の、もう 1 つのレイヤー：どちらも見えるときだけ見える。"""
        return State(self.view and inner.view, self.print and inner.print,
                     self.name if inner.name in ("", self.name) else f"{self.name}・{inner.name}" if self.name else inner.name)

    def how(self) -> str:
        """画面と印刷に出るかを、字で。"""
        if not self.view:
            return "画面に出ない・印刷には出る" if self.print else "画面にも印刷にも出ない"
        return "画面にも印刷にも出る" if self.print else "画面には出る・印刷には出ない"


def _key(ref):
    return getattr(ref, "num", id(ref))


def _name(o) -> str:
    return getattr(o, "name", "") if o is not None else ""


class Layers:
    def __init__(self, doc):
        self.doc = doc
        self.any = False  # 既定の設定を持つ書類か（持たなければ、内容の印はどれも見える）
        self._base_off = False
        self._on: set = set()
        self._off: set = set()
        self._event: dict[str, dict] = {"View": {}, "Print": {}}  # 使い道ごとの切り替え：グループ → 見えるか
        self._states: dict = {}
        try:
            ocp = doc.resolve(doc.root.get("OCProperties"))
        except Exception:  # noqa: BLE001  壊れた設定は、設定がないものとして読む（どの内容も見える）
            return
        if not isinstance(ocp, dict):
            return
        self.any = True
        try:
            d = doc.resolve(ocp.get("D"))
            d = d if isinstance(d, dict) else {}
            self._base_off = _name(doc.resolve(d.get("BaseState"))) == "OFF"
            self._on = {_key(x) for x in self._list(d.get("ON"))}
            self._off = {_key(x) for x in self._list(d.get("OFF"))}
            for app in self._list(d.get("AS")):
                app = doc.resolve(app)
                if not isinstance(app, dict):
                    continue
                event = _name(doc.resolve(app.get("Event")))
                if event not in self._event:
                    continue
                cats = [_name(doc.resolve(c)) for c in self._list(app.get("Category"))]
                for g in self._list(app.get("OCGs")):
                    st = self._usage(g, cats)
                    if st is not None:
                        # 同じグループが 2 つの切り替えに入っているとき：1 つでも OFF なら OFF
                        self._event[event][_key(g)] = st and self._event[event].get(_key(g), True)
        except (TypeError, ValueError, AttributeError, KeyError, IndexError, RecursionError):
            pass  # 読めた所までの設定で決める

    def _list(self, o) -> list:
        o = self.doc.resolve(o)
        return o if isinstance(o, list) else []

    def _usage(self, g, cats: list[str]) -> bool | None:
        """グループの Usage が、この使い道で勧める状態（決めていなければ None）。"""
        doc = self.doc
        gd = doc.resolve(g)
        usage = doc.resolve(gd.get("Usage")) if isinstance(gd, dict) else None
        if not isinstance(usage, dict):
            return None
        said = []
        for c in cats:
            u = doc.resolve(usage.get(c)) if c in _STATE_KEY else None
            st = _name(doc.resolve(u.get(_STATE_KEY[c]))) if isinstance(u, dict) else ""
            if st in ("ON", "OFF"):
                said.append(st == "ON")
        return all(said) if said else None

    # ------------------------------------------------------------------ グループ・まとめの印
    def _group(self, ref, event: str) -> bool:
        k = _key(ref)
        got = self._event[event].get(k)
        if got is not None:
            return got
        return k not in self._off and not (self._base_off and k not in self._on)

    def _ve(self, e, event: str, depth: int = 0):
        """見えるかの式（VE）：[/And …]・[/Or …]・[/Not x]。読めない式は None（P で決める）。"""
        if depth > 32:
            return None
        v = self.doc.resolve(e)
        if isinstance(v, dict):
            return self._group(e, event)
        if not isinstance(v, list) or not v:
            return None
        op = _name(self.doc.resolve(v[0]))
        args = [self._ve(x, event, depth + 1) for x in v[1:]]
        if not args or any(a is None for a in args):
            return None
        if op == "And":
            return all(args)
        if op == "Or":
            return any(args)
        if op == "Not" and len(args) == 1:
            return not args[0]
        return None

    def _members(self, d: dict) -> list:
        """まとめの印（OCMD）のグループ（VE の中のものも）。名前を伝えるため。"""
        doc = self.doc
        gs = doc.resolve(d.get("OCGs"))
        out = list(gs) if isinstance(gs, list) else [d.get("OCGs")] if gs is not None else []

        def walk(e, depth=0):
            v = doc.resolve(e)
            if isinstance(v, dict):
                if all(_key(e) != _key(o) for o in out):
                    out.append(e)
            elif isinstance(v, list) and depth < 32:
                for x in v[1:]:
                    walk(x, depth + 1)

        if d.get("VE") is not None:
            walk(d.get("VE"))
        return [g for g in out if isinstance(doc.resolve(g), dict)]

    def _visible(self, oc, d: dict, event: str) -> bool:
        if _name(self.doc.resolve(d.get("Type"))) != "OCMD":
            return self._group(oc, event)
        if d.get("VE") is not None:
            got = self._ve(d.get("VE"), event)
            if got is not None:
                return got
        gs = self.doc.resolve(d.get("OCGs"))
        lst = gs if isinstance(gs, list) else [d.get("OCGs")] if gs is not None else []
        states = [self._group(g, event) for g in lst if isinstance(self.doc.resolve(g), dict)]
        if not states:
            return True  # グループのないまとめの印は、見え方を変えない
        p = _name(self.doc.resolve(d.get("P"))) or "AnyOn"
        return {"AllOn": all(states), "AnyOn": any(states), "AnyOff": not all(states), "AllOff": not any(states)}.get(p, any(states))

    def _label(self, d: dict) -> str:
        n = self.doc.resolve(d.get("Name"))
        return decode_text(n).strip() if isinstance(n, (bytes, PdfString)) else ""

    def state(self, oc) -> State | None:
        """内容の印（/OC：OCG・OCMD）の、既定の設定での見え方。印がない・設定のない書類は None（ふつうに見える）。"""
        if oc is None or not self.any:
            return None
        k = _key(oc)
        if k in self._states:
            return self._states[k]
        try:
            d = self.doc.resolve(oc)
            if not isinstance(d, dict):
                return None
            if _name(self.doc.resolve(d.get("Type"))) == "OCMD":
                names = [self._label(self.doc.resolve(g)) for g in self._members(d)]
                name = "・".join(dict.fromkeys(n for n in names if n))
            else:
                name = self._label(d)
            st = State(self._visible(oc, d, "View"), self._visible(oc, d, "Print"), name)
        except (TypeError, ValueError, AttributeError, KeyError, IndexError, RecursionError):
            st = None  # 読めない印は、見え方を変えない（ビューアーは、読めない印の内容を見せる）
        if hasattr(oc, "num"):
            self._states[k] = st
        return st

    def special(self) -> list[State]:
        """書類のレイヤーのうち、既定の設定で、画面に見えない・印刷に出ないもの（書類に並ぶ順）。"""
        if not self.any:
            return []
        try:
            groups = self._list(self.doc.resolve(self.doc.root.get("OCProperties")).get("OCGs"))
        except Exception:  # noqa: BLE001
            return []
        out = []
        for g in groups:
            st = self.state(g)
            if st is not None and not (st.view and st.print):
                out.append(st)
        return out


def layers(doc) -> Layers:
    """書類のレイヤーの設定（書類ごとに 1 度だけ読む）。"""
    held = getattr(doc, "__dict__", None)
    if held is None:
        return Layers(doc)
    got = held.get("_layers")
    if got is None:
        got = held["_layers"] = Layers(doc)
    return got
