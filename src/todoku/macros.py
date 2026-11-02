"""窓の「マクロ」の節：マクロ（VBA）のモジュール・自動で動く処理・ボタン・全部のソース（Excel・Word 共通）。

todoku はマクロを動かさない（マクロの処理を規則で再現することはしない）。かわりに、
- いつ動くか（ブックを開いたとき・セルの値を変えたとき・ボタンを押したとき）を、手続きの名前から決定的に読み取り、
- ソースをそのまま窓に出して、何をするマクロかは、読んだエージェントが判断できるようにする。
セルの値を変えたときに動くマクロがある所に書き込むと、Excelで手で入力したときと結果が違うことがあるので、必ず知らせる。
"""

from __future__ import annotations

from .vba import KINDS, Project, visible_source

# 手続きの名前（小文字）→ いつ動くか
EVENTS = {
    "workbook_open": "ブックを開いたとき", "auto_open": "ブックを開いたとき", "autoopen": "文書を開いたとき",
    "document_open": "文書を開いたとき", "autoexec": "Word を起動したとき", "document_new": "新しい文書を作ったとき",
    "autonew": "新しい文書を作ったとき",
    "workbook_beforeclose": "ブックを閉じる前", "auto_close": "ブックを閉じる前", "autoclose": "文書を閉じる前",
    "document_close": "文書を閉じる前",
    "workbook_beforesave": "保存する前", "workbook_aftersave": "保存した後", "workbook_beforeprint": "印刷する前",
    "workbook_sheetchange": "どれかのシートのセルの値を変えたとき",
    "workbook_sheetcalculate": "どれかのシートが計算し直されたとき",
    "workbook_sheetselectionchange": "どれかのシートで、選ぶセルを変えたとき",
    "workbook_sheetactivate": "シートを切り替えたとき", "workbook_newsheet": "シートを足したとき",
    "worksheet_change": "このシートのセルの値を変えたとき",
    "worksheet_calculate": "このシートが計算し直されたとき",
    "worksheet_selectionchange": "このシートで、選ぶセルを変えたとき",
    "worksheet_activate": "このシートを開いたとき", "worksheet_deactivate": "このシートから離れたとき",
    "worksheet_beforedoubleclick": "セルをダブルクリックしたとき", "worksheet_beforerightclick": "セルを右クリックしたとき",
    "worksheet_followhyperlink": "シートの中のリンクを押したとき",
}
# 値を変えると動く（書き込みの結果が、Excelで手で入力したときと変わりうる）
CHANGE_EVENTS = {"workbook_sheetchange", "worksheet_change", "worksheet_calculate", "workbook_sheetcalculate"}


def _macro_name(ref: str) -> str:
    """ボタンに書かれたマクロの参照（「[0]!Module3.印刷」「'ブック.xlsm'!印刷」）から、手続きの名前を取り出す。"""
    return ref.rsplit("!", 1)[-1].strip("'")


def macro_info(project: Project | None, present: bool, sheets_by_code: dict[str, str] | None = None,
               buttons: list[dict] | None = None) -> dict | bool:
    if not present:
        return False
    if project is None:
        return {"present": True, "readable": False,
                "note": "マクロの入れ物（vbaProject.bin）を読めませんでした。マクロが何をするかは分かりません"}
    sheets_by_code = sheets_by_code or {}
    modules, events = [], []
    for m in project.modules:
        sheet = sheets_by_code.get(m.name)
        src = visible_source(m)
        modules.append({"name": m.name, "kind": KINDS.get(m.kind, m.kind), "sheet": sheet, "procedures": m.procedures,
                        "lines": len(src.splitlines()) if src else 0, "source": src})
        for proc in m.procedures:
            key = proc.lower()
            when = EVENTS.get(key)
            if when is None and m.kind in ("document", "form") and key.endswith("_click"):
                when = f"ボタン「{proc[:-6]}」を押したとき"
            if when:
                events.append({"module": m.name, "sheet": sheet, "procedure": proc, "when": when,
                               "on_change": key in CHANGE_EVENTS})
    return {"present": True, "readable": True, "project": project.name, "locked_for_viewing": project.locked,
            "modules": [x for x in modules if x["procedures"] or x["lines"]], "events": events,
            "buttons": [dict(b, macro=_macro_name(b["macro"])) for b in (buttons or [])],
            "note": "todoku はマクロを動かしません。マクロの処理（ボタンの処理・自動で動く処理）が要るかどうかは、"
                    "ソースを読んで判断してください。書き込みは、マクロを動かさずに、セルの値と計算結果だけを変えます"}


def change_event_sheets(info: dict | bool) -> dict[str | None, list[str]]:
    """値を変えると動くマクロ：シートの名前（ブック全体なら None）→ 手続き。"""
    out: dict[str | None, list[str]] = {}
    if not isinstance(info, dict):
        return out
    for e in info.get("events", []):
        if e["on_change"]:
            out.setdefault(e["sheet"] if e["procedure"].lower().startswith("worksheet_") else None, []).append(
                f"{e['module']}.{e['procedure']}")
    return out


def summary_line(info: dict | bool) -> str | None:
    if not isinstance(info, dict):
        return None
    if not info.get("readable"):
        return "**マクロを含みます**（マクロの入れ物を読めませんでした。何をするマクロかは分かりません）"
    n_mod = len(info["modules"])
    n_proc = sum(len(m["procedures"]) for m in info["modules"])
    s = f"**マクロを含みます**：モジュール {n_mod} 個・手続き {n_proc} 個。todoku はマクロを動かしません（下の「マクロ」の節に、全部のソースがあります）"
    if any(e["on_change"] for e in info["events"]):
        s += "。**セルの値を変えたときに動くマクロがあります**（Excelで手で入力したときは、マクロがほかのセルも変えることがあります）"
    return s


def markdown(info: dict | bool) -> list[str]:
    if not isinstance(info, dict) or not info.get("readable"):
        return []
    out = ["## マクロ（VBA）", "", f"- {info['note']}"]
    if info.get("locked_for_viewing"):
        out.append("- このマクロは、Excel の画面ではパスワードがないと見られない設定です（ファイルの中のソースは、そのまま読めます）")
    out.append("")
    if info["events"]:
        out += ["### 自動で動く処理・ボタンで動く処理", ""]
        for e in info["events"]:
            where = f"（シート「{e['sheet']}」）" if e["sheet"] else ""
            mark = "　← **値を変えると動く**" if e["on_change"] else ""
            out.append(f"- {e['when']}: {e['module']}{where}.{e['procedure']}{mark}")
        out.append("")
    if info["buttons"]:
        out += ["### マクロを割り当てたボタン", ""]
        for b in info["buttons"]:
            label = f"「{b['text']}」" if b.get("text") else ""
            out.append(f"- シート「{b['sheet']}」の {b['control']}{label}: {b['macro']}")
        out.append("")
    out += ["### モジュールとソース", ""]
    for m in info["modules"]:
        where = f"（シート「{m['sheet']}」）" if m["sheet"] else ""
        out.append(f"#### {m['name']}{where}：{m['kind']}、手続き {len(m['procedures'])} 個、{m['lines']} 行")
        out.append("")
        if m["procedures"]:
            out.append("手続き: " + "、".join(m["procedures"]))
            out.append("")
        if m["source"]:
            out += ["```vba", m["source"], "```", ""]
    return out
