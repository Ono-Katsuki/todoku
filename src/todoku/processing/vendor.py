"""計算の決まりを、とどくの計算エンジンのソースから、使う分だけ切り出す（同じ決まりの写しを手で書かない）。

formula/ の関数は、関数の丸め・日付・エラーの決まりをそのまま持っている。使う関数からたどって要る定義だけを集め、
1 つの Python のソースにする。作ったソースは標準ライブラリにしか頼らない。
"""

from __future__ import annotations

import ast
import builtins
import importlib
import warnings
from dataclasses import dataclass
from pathlib import Path


class VendorError(RuntimeError):
    pass


PKG = "todoku"
# 上から順に定義する（下の方の定数が上の定義を使う）
ORDER = ["formula.values", "formula.dates", "formula.parser", "formula.locale_fmt", "formula.numfmt", "formula.evaluator",
         "formula.functions", "formula.arrays"]
BUILTIN = set(dir(builtins))


@dataclass
class _Mod:
    name: str
    tree: ast.Module
    defs: dict[str, ast.stmt]  # 上の階の名前 → 定義する文
    froms: dict[str, tuple[str, str]]  # 名前 → (モジュール, 元の名前)
    aliases: dict[str, str]  # 名前 → モジュール（from . import arrays as _arrays）
    stdlib: dict[str, str]  # 名前 → import の文


def _resolve(here: str, level: int, module: str | None) -> str:
    parts = here.split(".")
    base = parts[:len(parts) - level]
    return ".".join(base + ([module] if module else []))


def _load(name: str) -> _Mod:
    mod = importlib.import_module(f"{PKG}.{name}")
    src = Path(mod.__file__).read_text(encoding="utf-8")
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", SyntaxWarning)
        warnings.simplefilter("ignore", DeprecationWarning)
        tree = ast.parse(src)
    m = _Mod(name, tree, {}, {}, {}, {})
    for st in tree.body:
        if isinstance(st, ast.ImportFrom):
            if st.module == "__future__":
                continue
            if st.level:
                target = _resolve(name, st.level, st.module)
                for a in st.names:
                    if st.module is None:
                        m.aliases[a.asname or a.name] = f"{target}.{a.name}" if target else a.name
                    else:
                        m.froms[a.asname or a.name] = (target, a.name)
            else:
                for a in st.names:
                    m.stdlib[a.asname or a.name] = f"from {st.module} import {a.name}" + (f" as {a.asname}" if a.asname else "")
        elif isinstance(st, ast.Import):
            for a in st.names:
                m.stdlib[a.asname or a.name] = f"import {a.name}" + (f" as {a.asname}" if a.asname else "")
        elif isinstance(st, (ast.FunctionDef, ast.ClassDef, ast.AsyncFunctionDef)):
            m.defs[st.name] = st
        elif isinstance(st, ast.Assign):
            for t in st.targets:
                for n in ([t] if isinstance(t, ast.Name) else t.elts if isinstance(t, ast.Tuple) else []):
                    if isinstance(n, ast.Name):
                        m.defs[n.id] = st
        elif isinstance(st, ast.AnnAssign) and isinstance(st.target, ast.Name):
            m.defs[st.target.id] = st
    return m


class Vendor:
    def __init__(self, reserved: set[str]):
        self.mods: dict[str, _Mod] = {}
        self.need_stmts: dict[str, list] = {}
        self.picked: dict[tuple[str, str], ast.stmt] = {}
        self.literals: dict[tuple[str, str], object] = {}
        self.stdlib: set[str] = set()
        self.reserved = set(reserved)
        self.functions: set[str] = set()
        self._todo: list[tuple[str, str]] = []
        self._walked: set[int] = set()
        self.registry: dict[str, tuple[str, object]] = {}
        self.order = list(ORDER)

    def mod(self, name: str) -> _Mod:
        got = self.mods.get(name)
        if got is None:
            got = self.mods[name] = _load(name)
        return got

    # ------------------------------------------------------------ たどる
    def need(self, mod: str, name: str) -> None:
        self._todo.append((mod, name))
        self._drain()

    def _drain(self) -> None:
        while self._todo:
            mod, name = self._todo.pop()
            self._need1(mod, name)

    def _need1(self, mod: str, name: str) -> None:
        k = (mod, name)
        if k in self.picked or k in self.literals:
            return
        if mod == "limits":  # 上限の定数は、値のまま書く
            v = getattr(importlib.import_module(f"{PKG}.limits"), name)
            if not isinstance(v, (int, float)):
                raise VendorError(f"limits.{name} は数ではありません")
            self.literals[k] = v
            return
        m = self.mod(mod)
        if name in m.froms:
            t, orig = m.froms[name]
            self._todo.append((t, orig))
            return
        if name in m.stdlib:
            self.stdlib.add(m.stdlib[name])
            return
        st = m.defs.get(name)
        if st is None:
            if name in BUILTIN:
                return
            raise VendorError(f"{mod} の {name} が見つかりません")
        self.picked[k] = st
        self._walk(m, st)

    def _walk(self, m: _Mod, st: ast.stmt) -> None:
        if id(st) in self._walked:
            return
        self._walked.add(id(st))
        for n in ast.walk(st):
            if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load):
                nm = n.id
                if nm in m.aliases:
                    continue
                if nm in m.defs or nm in m.froms or nm in m.stdlib:
                    self._todo.append((m.name, nm))
            elif isinstance(n, ast.Attribute) and isinstance(n.value, ast.Name) and n.value.id in m.aliases:
                self._todo.append((m.aliases[n.value.id], n.attr))
            elif isinstance(n, ast.ImportFrom):
                if n.level:
                    t = _resolve(m.name, n.level, n.module)
                    for a in n.names:
                        self._todo.append((t, a.name))
            elif isinstance(n, ast.Constant) and isinstance(n.value, str) and n.value in self.registry:
                self.want_function(n.value)  # EAGER[name] のように字で引く関数（SUBTOTAL など）

    def load_registry(self) -> None:
        """とどくに登録されている関数：名前 → (表, モジュール, 定義の名前, 工場に渡す字か None)。ソースの登録の書き方から読む。"""
        from ..formula import functions as F

        live = set(F.EAGER) | set(F.LAZY)
        found: dict[str, tuple] = {}
        for mod in ("formula.functions", "formula.arrays"):
            m = self.mod(mod)
            for st in m.tree.body:
                if isinstance(st, ast.FunctionDef):
                    for d in st.decorator_list:
                        if isinstance(d, ast.Call) and isinstance(d.func, ast.Name) and d.func.id == "fn":
                            lazy = any(k.arg == "lazy" and getattr(k.value, "value", False) for k in d.keywords)
                            for nm in d.args[0].value.split():
                                found[nm] = ("LAZY" if lazy else "EAGER", mod, st.name, None)
                elif isinstance(st, ast.Assign) and len(st.targets) == 1 and isinstance(st.targets[0], ast.Subscript):
                    t = st.targets[0]
                    if isinstance(t.value, ast.Name) and t.value.id in ("EAGER", "LAZY") and isinstance(st.value, ast.Name):
                        found[t.slice.value] = (t.value.id, mod, st.value.id, None)
                elif isinstance(st, ast.For) and isinstance(st.iter, ast.Tuple) and len(st.body) == 1 and isinstance(st.body[0], ast.Assign):
                    a = st.body[0]
                    t = a.targets[0]
                    if isinstance(t, ast.Subscript) and isinstance(t.value, ast.Name) and t.value.id in ("EAGER", "LAZY") and isinstance(a.value, ast.Call):
                        for x in st.iter.elts:
                            found[x.value] = (t.value.id, mod, a.value.func.id, x.value)
        missing = live - set(found)
        if missing:
            raise VendorError(f"登録の書き方を読めない関数: {sorted(missing)}")
        self.registry = {k: v for k, v in found.items() if k in live}

    def want_function(self, name: str) -> None:
        """シートの関数を使う（登録も出す）。登録されていなければ何もしない（計算のときに Unsupported）。"""
        got = self.registry.get(name)
        if got is None or name in self.functions:
            return
        self.functions.add(name)
        _table, mod, top, _arg = got
        self._todo.append((mod, top))

    def registrations(self, final: dict) -> list[str]:
        out = []
        for f in sorted(self.functions):
            table, mod, top, arg = self.registry[f]
            tb = self.global_name(final, "formula.functions", table)
            target = self.global_name(final, mod, top)
            out.append(f"{tb}[{f!r}] = {target}({arg!r})" if arg is not None else f"{tb}[{f!r}] = {target}")
        return out

    # ------------------------------------------------------------ 書き出す
    def source(self, extra: list[tuple[str, ast.stmt]] | None = None) -> tuple[str, dict[tuple[str, str], str]]:
        """(ソース, 名前の対応)。同じ名前が 2 つのモジュールにあれば、後のものに印を付ける。"""
        self._drain()
        final: dict[tuple[str, str], str] = {}
        taken = set(self.reserved)
        for mod in ["limits", *self.order]:
            for k in [k for k in (*self.literals, *self.picked) if k[0] == mod]:
                nm = k[1]
                if nm in taken:
                    nm = f"{mod.rsplit('.', 1)[-1]}__{k[1]}"
                final[k] = nm
                taken.add(nm)
        extra_mods = sorted({k[0] for k in self.picked} - set(self.order) - {"limits"})
        if extra_mods:
            raise VendorError(f"想定していないモジュール: {extra_mods}")
        lines: list[str] = []
        for k, v in sorted(self.literals.items()):
            lines.append(f"{final[k]} = {v!r}")
        done: set[int] = set()
        for mod in self.order:
            m = self.mods.get(mod)
            if m is None:
                continue
            for st in m.tree.body:
                names = [n for (mm, n), s in self.picked.items() if mm == mod and s is st]
                if not names or id(st) in done:
                    continue
                done.add(id(st))
                lines.append(ast.unparse(_Rewrite(m, final, self).rewrite(st)))
        return "\n".join(lines), final

    def global_name(self, final: dict, mod: str, name: str) -> str:
        m = self.mods.get(mod)
        if m is not None and name in m.froms:
            return self.global_name(final, *m.froms[name])
        return final.get((mod, name), name)


class _Rewrite(ast.NodeTransformer):
    """1 つの文を、1 つのソースに入れる形にする：`@fn` と注釈と説明の字を除き、関数の中の import を除き、名前を付け替える。"""

    def __init__(self, m: _Mod, final: dict, v: Vendor):
        self.m = m
        self.final = final
        self.v = v
        self.scopes: list[set[str]] = []
        self.local_from: list[dict[str, tuple[str, str]]] = []

    def rewrite(self, st: ast.stmt) -> ast.stmt:
        import copy

        return self.visit(copy.deepcopy(st))

    def _map(self, name: str) -> str:
        for s, fr in zip(reversed(self.scopes), reversed(self.local_from)):
            if name in fr:
                return self.v.global_name(self.final, *fr[name])
            if name in s:
                return name
        m = self.m
        if name in m.froms:
            return self.v.global_name(self.final, *m.froms[name])
        if (m.name, name) in self.final:
            return self.final[(m.name, name)]
        return name

    def visit_Name(self, n):
        n.id = self._map(n.id)
        return n

    def visit_Attribute(self, n):
        if isinstance(n.value, ast.Name) and n.value.id in self.m.aliases and not any(n.value.id in s for s in self.scopes):
            nm = self.v.global_name(self.final, self.m.aliases[n.value.id], n.attr)
            return ast.copy_location(ast.Name(nm, n.ctx), n)
        return self.generic_visit(n)

    def _enter(self, node) -> None:
        local, fr = _locals(node, self.m)
        self.scopes.append(local)
        self.local_from.append(fr)

    def _leave(self) -> None:
        self.scopes.pop()
        self.local_from.pop()

    def visit_FunctionDef(self, n):
        n.decorator_list = [d for d in n.decorator_list if not (isinstance(d, ast.Call) and isinstance(d.func, ast.Name) and d.func.id == "fn")]
        if self.scopes:
            pass
        else:
            n.name = self._map(n.name)
        n.returns = None
        self._enter(n)
        for a in [*n.args.posonlyargs, *n.args.args, *n.args.kwonlyargs, n.args.vararg, n.args.kwarg]:
            if a is not None:
                a.annotation = None
        n.args.defaults = [self.visit(d) for d in n.args.defaults]
        n.args.kw_defaults = [self.visit(d) if d is not None else None for d in n.args.kw_defaults]
        body = list(n.body)
        if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant) and isinstance(body[0].value.value, str):
            body = body[1:]
        out = []
        for st in body:
            got = self.visit(st)
            if got is None:
                continue
            out.extend(got) if isinstance(got, list) else out.append(got)
        n.body = out or [ast.Pass()]
        n.decorator_list = [self.visit(d) for d in n.decorator_list]
        self._leave()
        return n

    def visit_Lambda(self, n):
        self._enter(n)
        n.body = self.visit(n.body)
        n.args.defaults = [self.visit(d) for d in n.args.defaults]
        self._leave()
        return n

    def _comp(self, n):
        self._enter(n)
        r = self.generic_visit(n)
        self._leave()
        return r

    visit_ListComp = visit_SetComp = visit_GeneratorExp = visit_DictComp = _comp

    def visit_ClassDef(self, n):
        n.name = self._map(n.name) if not self.scopes else n.name
        body = list(n.body)
        if body and isinstance(body[0], ast.Expr) and isinstance(body[0].value, ast.Constant) and isinstance(body[0].value.value, str):
            body = body[1:]
        n.bases = [self.visit(b) for b in n.bases]
        n.decorator_list = [self.visit(d) for d in n.decorator_list]
        # クラスの中の名前は、そのクラスの属性（書き換えない）。メソッドの中の大域の名前だけ書き換える
        out = []
        for st in body:
            if isinstance(st, (ast.FunctionDef, ast.AsyncFunctionDef)):
                scopes, fr = self.scopes, self.local_from
                self.scopes, self.local_from = [], []
                name = st.name
                got = self.visit_FunctionDef(st)
                got.name = name
                self.scopes, self.local_from = scopes, fr
                out.append(got)
            elif isinstance(st, ast.AnnAssign):
                if st.value is not None:
                    st.value = self.visit(st.value)
                out.append(st)  # 注釈は、データクラスの欄として要る（字のまま）
            else:
                out.append(self.visit(st))
        n.body = out or [ast.Pass()]
        return n

    def visit_ImportFrom(self, n):
        if n.level and self.scopes:
            return None
        return n

    def visit_Import(self, n):
        return n


def _locals(node, m: _Mod) -> tuple[set[str], dict[str, tuple[str, str]]]:
    """関数（内包表記）の中で決まる名前と、関数の中の from . import の名前。"""
    out: set[str] = set()
    fr: dict[str, tuple[str, str]] = {}
    if isinstance(node, (ast.ListComp, ast.SetComp, ast.GeneratorExp, ast.DictComp)):
        for g in node.generators:
            for n in ast.walk(g.target):
                if isinstance(n, ast.Name):
                    out.add(n.id)
        return out, fr
    args = node.args
    for a in [*args.posonlyargs, *args.args, *args.kwonlyargs, args.vararg, args.kwarg]:
        if a is not None:
            out.add(a.arg)
    if isinstance(node, ast.Lambda):
        return out, fr

    def walk(st):
        for n in ast.iter_child_nodes(st):
            if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                out.add(n.name)
                continue
            if isinstance(n, (ast.Lambda, ast.ListComp, ast.SetComp, ast.GeneratorExp, ast.DictComp)):
                continue
            if isinstance(n, ast.Name) and isinstance(n.ctx, (ast.Store, ast.Del)):
                out.add(n.id)
            elif isinstance(n, ast.ExceptHandler) and n.name:
                out.add(n.name)
            elif isinstance(n, ast.Import):
                for a in n.names:
                    out.add((a.asname or a.name).split(".")[0])
            elif isinstance(n, ast.ImportFrom):
                if n.level:
                    t = _resolve(m.name, n.level, n.module)
                    for a in n.names:
                        fr[a.asname or a.name] = (t, a.name)
                else:
                    for a in n.names:
                        out.add(a.asname or a.name)
            walk(n)

    for st in node.body:
        if isinstance(st, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            out.add(st.name)
            continue
        if isinstance(st, ast.Global):
            continue
        walk(ast.Module([st], []))
    return out, fr


def method_source(rel_module: str, cls: str, method: str, v: Vendor, final: dict) -> str:
    """とどくのクラスのメソッドのソース（1 つのソースに入れる形）。"""
    m = v.mod(rel_module)
    for st in m.tree.body:
        if isinstance(st, ast.ClassDef) and st.name == cls:
            for x in st.body:
                if isinstance(x, ast.FunctionDef) and x.name == method:
                    r = _Rewrite(m, final, v)
                    r.scopes, r.local_from = [set()], [{}]
                    import copy

                    x2 = r.visit_FunctionDef(copy.deepcopy(x))
                    return ast.unparse(x2)
    raise VendorError(f"{rel_module}.{cls}.{method} が見つかりません")


def need_method(rel_module: str, cls: str, method: str, v: Vendor) -> None:
    m = v.mod(rel_module)
    for st in m.tree.body:
        if isinstance(st, ast.ClassDef) and st.name == cls:
            for x in st.body:
                if isinstance(x, ast.FunctionDef) and x.name == method:
                    v._walk(m, x)
                    v._drain()
                    return
    raise VendorError(f"{rel_module}.{cls}.{method} が見つかりません")
