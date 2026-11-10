"""マクロが「選ぶ・切り替える・印刷する・画面の設定・知らせる」だけか（セルや本文の値を変えないか）を、VBA の文として読んで決める。

行の字ではなく構文木で決めるので、1 行に「:」で並べた文・行の終わりのコメント・字（文字列）の中の語で、答えが変わらない。
決められないもの（読めない文・知らない名前・型の分からない変数への代入・条件付きコンパイル）は、どれも False：
マクロの処理は、エージェントがソースを読んで判断する（段階 C）。
"""

from __future__ import annotations

from . import parse as P

# 文として呼ぶと、選ぶ・切り替える・印刷するだけのもの
_ACTIONS = {"select", "activate", "printout", "printpreview"}
# 画面・動かし方の設定（Application の下）
_APP_SETTINGS = {"screenupdating", "displayalerts", "moveafterreturndirection", "statusbar", "cursor", "enableevents"}
# 式の中で使える、読むだけ・たどるだけのオブジェクトの名前（Excel・Word）。同じ名前で値を変えるメソッドのないもの
_READS = {
    "application", "parent", "activesheet", "activeworkbook", "activewindow", "activecell", "activechart", "activedocument",
    "activeprinter", "selection", "selectedsheets", "thisworkbook", "thisdocument", "workbooks", "windows", "sheets", "worksheets",
    "charts", "range", "cells", "rows", "columns", "item", "offset", "resize", "end", "entirerow", "entirecolumn", "currentregion",
    "usedrange", "value", "value2", "text", "formula", "name", "codename", "count", "row", "column", "address", "index", "visible",
    "caption", "path", "fullname", "version", "username", "pagesetup", "printarea", "areas", "names", "number", "description",
}
# 値を返すが、ほかの処理を動かす・オブジェクトを作る関数（式の中でも使えない）
_UNSAFE_FUNCS = {"callbyname", "createobject", "run", "application.run"}
# 数・字・日付の型：この型と宣言した変数への代入は、変数を変えるだけ
# （オブジェクトを入れられる変数への Let の代入は、そのオブジェクトの既定のプロパティ（セルの値など）に書きうる）
_SCALAR = {"integer", "long", "longlong", "longptr", "single", "double", "currency", "string", "boolean", "date", "byte"}


def _pure_funcs() -> set[str]:
    """式の中で呼べる、組み込みの関数（動かす道具が知っている VBA の関数）と、たどるだけの Excel の関数。"""
    from .builtins import FUNCS

    return (set(FUNCS) - _UNSAFE_FUNCS) | {"sheets", "worksheets", "range", "cells", "rows", "columns", "charts", "workbooks",
                                           "windows", "names", "intersect", "union", "rgb"}


def _scalar(v: P.DimVar) -> bool:
    return v.bounds is None and not v.new and (v.type or "").lower() in _SCALAR


def _walk(body):
    for st in body or []:
        yield st
        if isinstance(st, P.If):
            for _c, b in st.branches:
                yield from _walk(b)
            yield from _walk(st.else_body)
        elif isinstance(st, P.Select):
            for _c, b in st.cases:
                yield from _walk(b)
        elif isinstance(st, (P.For, P.ForEach, P.Do, P.With)):
            yield from _walk(st.body)


class _Judge:
    def __init__(self, mod: P.Module):
        self.mod = mod
        self.funcs = _pure_funcs()
        self.procs = {p.name for p in mod.procs.values() if p.kind in ("sub", "function")}  # 同じモジュールの手続き（その中身も確かめる）
        self.scalars: set[str] = set()
        self.arrays: set[str] = set()

    def proc(self, p: P.Proc) -> bool:
        self.scalars = {v.name for d in self.mod.decls for v in d.vars if _scalar(v)}
        self.arrays = {v.name for d in self.mod.decls for v in d.vars if v.bounds is not None}
        for prm in p.params:  # 引数の型は分からない：同じ名前のモジュールの変数を隠す
            self.scalars.discard(prm.name)
            self.arrays.discard(prm.name)
        for st in _walk(p.body):
            if isinstance(st, P.Dim):
                for v in st.vars:
                    self.scalars.discard(v.name)
                    self.arrays.discard(v.name)
                    if v.bounds is not None:
                        self.arrays.add(v.name)
                    elif _scalar(v):
                        self.scalars.add(v.name)
        return self.block(p.body)

    def pure(self, e) -> bool:
        """式が、読むだけか。"""
        if e is None or isinstance(e, (P.Lit, P.DateLit, P.Name)):
            return True
        if isinstance(e, P.Unary):
            return self.pure(e.operand)
        if isinstance(e, P.Binary):
            return self.pure(e.left) and self.pure(e.right)
        if isinstance(e, P.Member):
            return e.name in _READS and not e.bang and self.pure(e.obj)
        if isinstance(e, P.Call):
            if not all(self.pure(a.value) for a in e.args):
                return False
            t = e.target
            if isinstance(t, P.Name):
                return t.name in self.funcs or t.name in self.procs or t.name in self.arrays
            if isinstance(t, P.Member):
                return t.name in _READS and not t.bang and self.pure(t.obj)
        return False

    def block(self, body) -> bool:
        return all(self.stmt(st) for st in body or [])

    def stmt(self, st) -> bool:
        if isinstance(st, P.Dim):
            return all(self.pure(v.init) and all(self.pure(lo) and self.pure(hi) for lo, hi in (v.bounds or [])) for v in st.vars)
        if isinstance(st, (P.OnError, P.Exit, P.GoTo, P.Label, P.Resume, P.End)):
            return True
        if isinstance(st, P.If):
            return all(self.pure(c) and self.block(b) for c, b in st.branches) and self.block(st.else_body)
        if isinstance(st, P.With):
            return self.pure(st.obj) and self.block(st.body)
        if isinstance(st, P.For):
            return (isinstance(st.var, P.Name) and st.var.name in self.scalars and self.pure(st.start) and self.pure(st.end)
                    and self.pure(st.step) and self.block(st.body))
        if isinstance(st, P.ForEach):
            return isinstance(st.var, P.Name) and self.pure(st.coll) and self.block(st.body)
        if isinstance(st, P.Do):
            return all(self.pure(c[1]) for c in (st.pre, st.post) if c) and self.block(st.body)
        if isinstance(st, P.Select):
            return self.pure(st.expr) and all(self._case(conds) and self.block(b) for conds, b in st.cases)
        if isinstance(st, P.Assign):
            t = st.target
            if isinstance(t, P.Name):
                # Set は変数に参照を入れるだけ。Let は、数・字の型と宣言した変数だけ
                return (st.set or t.name in self.scalars) and self.pure(st.value)
            if isinstance(t, P.Member) and isinstance(t.obj, P.Name) and t.obj.name == "application" and t.name in _APP_SETTINGS:
                return self.pure(st.value)
            return False
        if isinstance(st, P.CallStmt):
            e, args = st.expr, []
            if isinstance(e, P.Call):
                e, args = e.target, e.args
            if not all(self.pure(a.value) for a in args):
                return False
            if isinstance(e, P.Member):
                return e.name in _ACTIONS and not e.bang and self.pure(e.obj)
            if isinstance(e, P.Name):
                return e.name == "msgbox" or e.name in self.procs
        return False

    def _case(self, conds) -> bool:
        return all(self.pure(x) for c in conds or [] for x in c if isinstance(x, P.Node))


def display_only(source: str) -> bool:
    """モジュールの手続きの文が、すべて「選ぶ・切り替える・印刷する・画面の設定・知らせる」だけなら True。"""
    src = source.replace("\r\n", "\n").replace("\r", "\n")
    try:
        if P.preprocess(src) != src:
            return False  # 条件付きコンパイル：使わない側の文を読んでいない
        mod = P.parse_module("m", src)
    except Exception:  # noqa: BLE001  読めない文のあるマクロは、決めない
        return False
    j = _Judge(mod)
    return all(j.proc(p) for p in mod.procs.values())
