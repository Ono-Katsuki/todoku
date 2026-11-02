"""OpenDocument のマクロ（LibreOffice Basic など）：パッケージの Basic/ の部品のソースと、動くきっかけ（office:event-listeners）。

todoku はマクロを動かさない。ソースをそのまま窓に出し、いつ動くか（文書を開いたとき・部品を押したとき・値を変えたとき）を
きっかけの名前（script:event-name）から決まった表で読む。窓の「マクロ」の節は、Excel・Word と同じ形にする。
"""

from __future__ import annotations

import re

from .package import OdfPackage
from .xml import El, parse, top

# きっかけの名前（ODF 1.3 第3部 10.3 の event-name）→ いつ動くか
WHEN = {"dom:load": "文書を開いたとき", "office:new": "新しい文書を作ったとき", "office:create-document": "新しい文書を作ったとき",
        "office:load-finished": "文書を開いたとき", "dom:unload": "文書を閉じるとき", "office:prepare-unload": "文書を閉じる前",
        "office:save": "保存するとき", "office:save-as": "名前を付けて保存するとき", "office:print": "印刷するとき",
        "form:performaction": "部品を押したとき", "dom:click": "部品を押したとき", "form:change": "部品の値を変えたとき",
        "dom:change": "値を変えたとき", "form:textchange": "部品の字を変えたとき", "form:itemstatechange": "部品の選んだものを変えたとき",
        "dom:mouseup": "部品を押したとき", "office:content-changed": "中身を変えたとき"}
ON_CHANGE = {"form:change", "dom:change", "form:textchange", "form:itemstatechange", "office:content-changed"}
_PROC = re.compile(r"^[ \t]*(?:(?:Private|Public|Static)[ \t]+)?(?:Sub|Function)[ \t]+([A-Za-z_぀-鿿][\w぀-鿿]*)", re.I | re.M)


def macro_info(pkg: OdfPackage, content_root: El | None = None) -> dict | bool:
    """窓の「マクロ」の節（macros.markdown が読む形）。マクロがなければ False。"""
    modules = []
    if not pkg.flat:
        for name in pkg.names:
            if name.startswith("Basic/") and name.endswith(".xml") and not name.endswith(("script-lc.xml", "script-lb.xml",
                                                                                           "dialog-lc.xml", "dialog-lb.xml")):
                data = pkg.zip.read(name)
                try:
                    root = top(parse(data))
                except Exception:  # noqa: BLE001
                    modules.append({"name": name, "kind": "LibreOffice Basic", "sheet": None, "procedures": [], "lines": 0,
                                    "source": "", "unreadable": True})
                    continue
                if root.local != "module":
                    continue
                src = "".join(k.value(data) for k in root.kids if not isinstance(k, El))
                lib = name.split("/")[1] if name.count("/") >= 2 else ""
                mod = root.get("{http://openoffice.org/2000/script}name") or name.rsplit("/", 1)[-1][:-4]
                modules.append({"name": f"{lib}.{mod}" if lib else mod, "kind": "LibreOffice Basic", "sheet": None,
                                "procedures": _PROC.findall(src), "lines": len(src.splitlines()), "source": src})
            elif name.startswith("Scripts/") and not name.endswith("/") and not name.endswith(".xml"):
                data = pkg.zip.read(name)
                src = data.decode("utf-8", "replace")
                modules.append({"name": name, "kind": "スクリプト（" + name.rsplit(".", 1)[-1] + "）", "sheet": None,
                                "procedures": [], "lines": len(src.splitlines()), "source": src})
    events = []
    if content_root is not None:
        for el in content_root.iter("script:event-listener"):
            ev = el.get("script:event-name", "") or ""
            href = el.get("xlink:href", "") or el.get("script:macro-name", "") or ""
            m = re.search(r"vnd\.sun\.star\.script:([^?]+)", href)
            target = m[1] if m else href
            mod, _, proc = target.rpartition(".")
            owner = el.parent.parent if el.parent is not None else None
            who = owner.get("form:name") if owner is not None and owner.get("form:name") else ""
            when = WHEN.get(ev, ev)
            if who:
                when = f"{when}（部品「{who}」）"
            events.append({"module": mod or target, "sheet": None, "procedure": proc or target, "when": when,
                           "on_change": ev in ON_CHANGE})
    if not modules and not events:
        return False
    return {"present": True, "readable": not any(m.get("unreadable") for m in modules), "project": "LibreOffice Basic",
            "locked_for_viewing": False, "modules": [m for m in modules if m["procedures"] or m["lines"]], "events": events,
            "buttons": [],
            "note": "todoku はマクロを動かしません。マクロの処理（部品を押したときの処理・自動で動く処理）が要るかどうかは、"
                    "ソースを読んで判断してください。書き込みは、マクロを動かさずに、書いた所の値だけを変えます"}
