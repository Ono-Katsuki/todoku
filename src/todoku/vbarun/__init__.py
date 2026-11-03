"""マクロ（VBA）を動かす仕組み：VBA の文法の解析（parse）・実行（interp）・組み込みの関数（builtins）・
マクロから見える Excel（excel）。

使い方：
- run_macro(wb, "Module1.集計", answers=[...], inputs={...}) … ボタンのマクロを動かし、マクロが変えたセルを返す
- change_events(wb, [(シート, 行, 列, 値), …]) … 人がセルに値を入れたときと同じように、Worksheet_Change などを動かす
仕組みで動かせない命令は Unsupported、人の答えやファイルが要るときは NeedsInput を投げる（勝手に決めない）。
"""

from __future__ import annotations

from dataclasses import dataclass, field

from .interp import NeedsInput, Unsupported, VbaError

__all__ = ["MacroRun", "NeedsInput", "Unsupported", "VbaError", "change_events", "load", "run_macro"]


@dataclass
class MacroRun:
    changed: dict = field(default_factory=dict)  # マクロが変えたセル：(シート, 行, 列) → 値（None は空にした）
    messages: list[str] = field(default_factory=list)  # MsgBox で出した文
    files: dict[str, bytes] = field(default_factory=dict)  # マクロが書いたファイル
    printed: list[str] = field(default_factory=list)  # 印刷したシート
    volatile: bool = False  # 今の日付・時刻を使った（動かすたびに結果が変わる）
    error: str | None = None  # Excel でも出る実行時エラー（マクロはそこで止まる。止まるまでに変えたセルは changed）
    notes: list[str] = field(default_factory=list)  # 結果について伝えること（地域で見え方が変わる字をマクロに渡した、など）


def load(wb, overrides: dict | None = None, answers=None, inputs=None, now=None):
    """ブックのマクロを読み、動かす準備をする：(Runtime, Host)。"""
    from .excel import Host
    from .interp import Runtime
    from .parse import parse_module

    project = getattr(wb, "vba", None)
    if project is None:
        raise Unsupported("マクロの入れ物を読めません")
    consts = getattr(project, "cc_constants", None)
    modules = [parse_module(m.name, m.source, consts) for m in project.modules if m.source]
    host = Host(wb, overrides, now)
    rt = Runtime(modules, host, answers, inputs)
    return rt, host


def _result(rt, host) -> MacroRun:
    return MacroRun(dict(host.changed), list(rt.messages), {k: bytes(v) for k, v in rt.files_out.items()},
                    list(host.printed), host.volatile, notes=list(host.notes))


def run_macro(wb, name: str, answers=None, inputs=None, now=None, overrides: dict | None = None) -> MacroRun:
    rt, host = load(wb, overrides, answers, inputs, now)
    try:
        rt.run(name)
    except VbaError as e:
        res = _result(rt, host)
        where = getattr(e, "where", None)
        at = f"{where[0]} の {where[1]} 行目で" if where else ""
        res.error = f"{at}実行時エラー '{e.number}': {e.description}（Excel で動かしても、ここでエラーになって止まります）"
        return res
    return _result(rt, host)


def change_events(wb, entries: list[tuple[str, int, int, object]], answers=None, inputs=None, now=None) -> MacroRun:
    """entries の順に、人がセルに値を入れたのと同じように書き、そのたびに Change イベントを動かす。"""
    from .excel import Range

    rt, host = load(wb, None, answers, inputs, now)
    errors = []
    for sheet, r, c, v in entries:
        host.set(sheet, r, c, v, by_macro=False)
        try:
            host.fire_change(Range(host, host.sheet(sheet).name, r, c))
        except VbaError as e:
            # Excel でも、イベントの処理はここで止まり（エラーの画面が出る）、入れた値は残る。次の入力のイベントは動く
            where = getattr(e, "where", None)
            at = f"{where[0]} の {where[1]} 行目で" if where else ""
            errors.append(f"{sheet}!{chr(64 + c) if c <= 26 else c}{r} に入れたとき、{at}実行時エラー '{e.number}': {e.description}")
            host.depth = 0
    res = _result(rt, host)
    if errors:
        res.error = "／".join(errors) + "（Excel でも、ここでエラーの画面が出て、そのイベントの処理は止まります）"
    return res
