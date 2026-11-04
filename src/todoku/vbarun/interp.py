"""VBA の実行：変数・配列・手続きの呼び出し（参照渡し）・制御の文・エラー処理・組み込みの関数。

Excel の部品（Range・Worksheet・Application）は host（excel.py）が持ち、ここは言語の規則だけを持つ。
決めて動かせないこと（MsgBox の答え・InputBox・ファイルの読み込み）は、勝手に決めずに NeedsInput を投げて、
答えやファイルを渡してもらう。動かせない命令は Unsupported（その行を添えて）を投げる。
"""

from __future__ import annotations

import datetime as dt
import math
import re
from dataclasses import dataclass, field

from . import parse as P
from .parse import EMPTY, MISSING, NOTHING, NULL

# ====================================================================== 例外
class VbaError(Exception):
    """VBA の実行時エラー（Err.Number / Err.Description）。On Error で受け止められる。"""

    def __init__(self, number: int, description: str = ""):
        super().__init__(f"実行時エラー {number}: {description}")
        self.number = number
        self.description = description or _ERR_TEXT.get(number, "")


class Unsupported(Exception):
    """この仕組みでは再現しない命令（規則では再現できない）。"""

    def __init__(self, what: str, line: int = 0, module: str = ""):
        where = f"{module} の {line} 行目" if module and line else (f"{line} 行目" if line else "")
        super().__init__(f"{where}：{what}" if where else what)
        self.what = what
        self.line = line
        self.module = module


class NeedsInput(Exception):
    """人の答え（MsgBox・InputBox）や、外のファイルが要る。kind と問いを返す。"""

    def __init__(self, kind: str, prompt: str, line: int = 0):
        super().__init__(f"{kind}: {prompt}")
        self.kind = kind
        self.prompt = prompt
        self.line = line


class _ExitProc(Exception):
    pass


class _ExitLoop(Exception):
    def __init__(self, what: str):
        self.what = what


class _GoTo(Exception):
    def __init__(self, label: str):
        self.label = label


class _EndProgram(Exception):
    pass


_ERR_TEXT = {5: "プロシージャの呼び出し、または引数が不正です", 6: "オーバーフローしました", 9: "インデックスが有効範囲にありません",
             11: "0 で除算しました", 13: "型が一致しません", 91: "オブジェクト変数または With ブロック変数が設定されていません",
             424: "オブジェクトが必要です", 438: "オブジェクトは、このプロパティまたはメソッドをサポートしていません",
             1004: "アプリケーション定義またはオブジェクト定義のエラーです"}


# ====================================================================== 値
class VDate(float):
    """日付（Excel と同じシリアル値。1899/12/30 が 0）。"""


@dataclass
class VArray:
    lbounds: list[int]
    ubounds: list[int]
    data: dict = field(default_factory=dict)  # 添字のタプル → 値

    def __post_init__(self):
        from ..limits import MAX_VBA_ARRAY

        n = 1
        for lo, hi in zip(self.lbounds, self.ubounds):
            n *= max(0, hi - lo + 1)
            if n > MAX_VBA_ARRAY:
                raise Unsupported(f"とても大きな配列（要素が {MAX_VBA_ARRAY:,} 個を超える）")

    def check(self, idx: tuple) -> tuple:
        if len(idx) != len(self.lbounds):
            raise VbaError(9)
        for i, lo, hi in zip(idx, self.lbounds, self.ubounds):
            if not (lo <= i <= hi):
                raise VbaError(9)
        return idx

    def get(self, idx: tuple):
        return self.data.get(self.check(idx), EMPTY)

    def set(self, idx: tuple, v) -> None:
        self.data[self.check(idx)] = v

    def flat(self) -> list:
        """要素を、VBA の For Each の順（最初の添字がいちばん速く変わる）で。"""
        import itertools

        ranges = [range(lo, hi + 1) for lo, hi in zip(self.lbounds, self.ubounds)]
        return [self.data.get(tuple(reversed(k)), EMPTY) for k in itertools.product(*reversed(ranges))]


class Ref:
    """変数の入れ物（参照渡しに使う）。"""

    __slots__ = ("value", "is_array_decl", "type")

    def __init__(self, value=EMPTY, type_: str | None = None):
        self.value = value
        self.is_array_decl = False
        self.type = type_


class Collection:
    def __init__(self):
        self.items: list = []
        self.keys: dict[str, int] = {}


class Dictionary:
    def __init__(self):
        self.d: dict = {}
        self.compare_text = False

    def key(self, k):
        if isinstance(k, str) and self.compare_text:
            return k.lower()
        return k


@dataclass
class UserType:
    name: str
    fields: dict


def serial_from(d: dt.datetime) -> VDate:
    base = dt.datetime(1899, 12, 30)
    delta = d - base
    return VDate(delta.days + delta.seconds / 86400)


def to_datetime(v) -> dt.datetime:
    x = to_double(v)
    base = dt.datetime(1899, 12, 30)
    days = math.floor(x)
    secs = round((x - days) * 86400)
    return base + dt.timedelta(days=days, seconds=secs)


# ---------------------------------------------------------------------- 変換
def is_numeric_str(s: str) -> bool:
    try:
        float(s.replace(",", "").strip())
        return s.strip() != ""
    except ValueError:
        return False


def to_double(v) -> float:
    if v is EMPTY or v is None:
        return 0.0
    if hasattr(v, "vba_default") and not isinstance(v, (VArray,)):
        return to_double(v.vba_default())
    if isinstance(v, bool):
        return -1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        t = v.strip().replace(",", "")
        if t == "":
            raise VbaError(13)
        if t.lower().startswith("&h"):
            return float(int(t[2:], 16))
        try:
            return float(t)
        except ValueError:
            d = parse_date_text(v)
            if d is not None:
                return float(d)
            raise VbaError(13)
    if v is NULL:
        raise VbaError(94, "Null の使い方が不正です")
    from ..formula.values import ExcelError

    if isinstance(v, ExcelError):
        raise VbaError(13)
    raise VbaError(13)


def bankers(x: float) -> int:
    """VBA の整数への丸め（銀行型：0.5 は偶数へ）。"""
    return int(round(x))


def to_long(v) -> int:
    x = to_double(v)
    if not -2147483648 <= x <= 2147483647:
        raise VbaError(6)
    return bankers(x)


def to_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    if hasattr(v, "vba_default") and not isinstance(v, (VArray,)):
        return to_bool(v.vba_default())
    if v is EMPTY:
        return False
    if isinstance(v, str):
        t = v.strip().lower()
        if t in ("true", "#true#"):
            return True
        if t in ("false", "#false#"):
            return False
    return to_double(v) != 0


def to_str(v) -> str:
    if v is EMPTY or v is None:
        return ""
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, VDate):
        return format_date_default(v)
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        r = repr(v)
        if "e" in r:
            m, e = r.split("e")
            return f"{m}E{'+' if not e.startswith('-') else ''}{e}"
        return f"{v:.15g}"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, str):
        return v
    if v is NULL:
        raise VbaError(94, "Null の使い方が不正です")
    from ..formula.values import ExcelError

    if isinstance(v, ExcelError):
        return f"Error {_CVERR.get(v.code, 2015)}"
    if hasattr(v, "vba_default"):
        return to_str(v.vba_default())
    raise VbaError(13)


_CVERR = {"#NULL!": 2000, "#DIV/0!": 2007, "#VALUE!": 2015, "#REF!": 2023, "#NAME?": 2029, "#NUM!": 2036, "#N/A": 2042}


def format_date_default(v) -> str:
    d = to_datetime(v)
    if float(v) == int(float(v)):
        return f"{d.year}/{d.month:02d}/{d.day:02d}"
    if float(v) < 1:
        return d.strftime("%H:%M:%S")
    return f"{d.year}/{d.month:02d}/{d.day:02d} {d.strftime('%H:%M:%S')}"


def parse_date_text(s: str) -> VDate | None:
    t = s.strip()
    m = re.fullmatch(r"(\d{4})[/\-年](\d{1,2})[/\-月](\d{1,2})日?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?", t)
    if m:
        y, mo, d = int(m[1]), int(m[2]), int(m[3])
        hh, mm, ss = int(m[4] or 0), int(m[5] or 0), int(m[6] or 0)
        try:
            return serial_from(dt.datetime(y, mo, d, hh, mm, ss))
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/(\d{4})", t)
    if m:
        try:
            return serial_from(dt.datetime(int(m[3]), int(m[1]), int(m[2])))
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{1,2}):(\d{2})(?::(\d{2}))?", t)
    if m:
        return VDate((int(m[1]) * 3600 + int(m[2]) * 60 + int(m[3] or 0)) / 86400)
    return None


def num_result(x: float):
    """計算の結果：整数で表せるときは整数（VBA の Integer・Long と Double の区別を、表示が変わらないように近づける）。"""
    if isinstance(x, float) and x == int(x) and abs(x) < 2**31:
        return int(x)
    return x


def compare(a, b, text: bool = False) -> int:
    if a is NULL or b is NULL:
        return 0
    if a is EMPTY:
        a = "" if isinstance(b, str) else 0
    if b is EMPTY:
        b = "" if isinstance(a, str) else 0
    if isinstance(a, str) and isinstance(b, str):
        x, y = (a.lower(), b.lower()) if text else (a, b)
        return (x > y) - (x < y)
    if isinstance(a, str) and not isinstance(b, str):
        if not is_numeric_str(a):
            return 1  # 数より文字列が大きい
        a = to_double(a)
    if isinstance(b, str) and not isinstance(a, str):
        if not is_numeric_str(b):
            return -1
        b = to_double(b)
    x, y = to_double(a), to_double(b)
    return (x > y) - (x < y)


def like(s: str, pat: str) -> bool:
    out = []
    i = 0
    while i < len(pat):
        ch = pat[i]
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        elif ch == "#":
            out.append(r"\d")
        elif ch == "[":
            j = pat.index("]", i)
            body = pat[i + 1:j]
            if body.startswith("!"):
                body = "^" + body[1:]
            out.append("[" + body + "]")
            i = j
        else:
            out.append(re.escape(ch))
        i += 1
    return re.fullmatch("".join(out), s, re.S) is not None


# ====================================================================== 実行
@dataclass
class Frame:
    proc: P.Proc | None
    module: str
    locals: dict[str, Ref] = field(default_factory=dict)
    with_stack: list = field(default_factory=list)
    on_error: str = "none"  # none / resume_next / goto
    error_label: str = ""
    resume_at: int | None = None  # On Error GoTo のあと Resume Next で戻る場所（いちばん外の文の番号）
    in_handler: bool = False


class Runtime:
    def __init__(self, modules: list[P.Module], host, answers: list | None = None, inputs: dict | None = None):
        self.modules = {m.name.lower(): m for m in modules}
        self.host = host
        host.rt = self
        self.answers = list(answers or [])  # MsgBox・InputBox の答え（順に使う）
        self.inputs = dict(inputs or {})  # 読むファイル：パス → 中身（bytes）
        self.module_vars: dict[str, dict[str, Ref]] = {}
        self.statics: dict[str, dict[str, Ref]] = {}
        self.err = ErrObject()
        self.messages: list[str] = []  # MsgBox で出した文
        self.files_out: dict[str, bytearray] = {}
        self.open_files: dict[int, dict] = {}
        self.steps = 0
        self.max_steps = 5_000_000
        self.big_chars = 0  # この実行で作った大きな文字列（1 MB 以上）の字の数の合計（limits.MAX_VBA_BYTES まで）
        self.frames: list[Frame] = []
        for m in modules:
            vars_ = {}
            for d in m.decls:
                for v in d.vars:
                    vars_[v.name] = Ref(self._initial(v, m.name, const=d.const))
            self.module_vars[m.name.lower()] = vars_

    # ---------------------------------------------------------------- 入口
    def find_proc(self, name: str, module: str | None = None) -> P.Proc | None:
        name = name.lower()
        if "." in name:
            mod, _, pn = name.rpartition(".")
            m = self.modules.get(mod)
            return m.procs.get(pn) if m else None
        if module and module.lower() in self.modules and name in self.modules[module.lower()].procs:
            return self.modules[module.lower()].procs[name]
        for m in self.modules.values():
            p = m.procs.get(name)
            if p is not None and not p.private:
                return p
        return None

    def run(self, name: str, args: list | None = None, module: str | None = None):
        p = self.find_proc(name, module)
        if p is None:
            raise Unsupported(f"マクロ「{name}」が見つかりません")
        try:
            return self.call_proc(p, [Ref(a) for a in (args or [])])
        except _EndProgram:
            return EMPTY

    # ---------------------------------------------------------------- 手続き
    def call_proc(self, p: P.Proc, args: list[Ref], named: dict | None = None):
        if len(self.frames) > 200:
            raise VbaError(28, "スタック領域が不足しています")
        fr = Frame(p, p.module)
        named = named or {}
        pos = 0
        for prm in p.params:
            if prm.paramarray:
                rest = [a.value for a in args[pos:]]
                fr.locals[prm.name] = Ref(VArray([0], [len(rest) - 1], {(i,): v for i, v in enumerate(rest)}))
                pos = len(args)
                continue
            if prm.name in named:
                a = named[prm.name]
            elif pos < len(args):
                a = args[pos]
                pos += 1
            else:
                a = None
            if a is None or a.value is MISSING:
                if not prm.optional:
                    raise VbaError(449, f"引数「{prm.name}」が省略されています")
                v = self.eval(prm.default, fr) if prm.default is not None else MISSING
                fr.locals[prm.name] = Ref(v)
            elif prm.byval:
                fr.locals[prm.name] = Ref(_copy(a.value))
            else:
                fr.locals[prm.name] = a  # 参照渡し
        if p.kind in ("function", "get"):
            fr.locals["__ret__"] = Ref(EMPTY)
        self.frames.append(fr)
        try:
            self.exec_body(p.body, fr, top=True)
        except _ExitProc:
            pass
        finally:
            self.frames.pop()
        return fr.locals["__ret__"].value if "__ret__" in fr.locals else EMPTY

    def _initial(self, v: P.DimVar, module: str, const: bool = False, fr: Frame | None = None):
        if const and v.init is not None:
            return self.eval(v.init, fr or Frame(None, module))
        if v.bounds is not None:
            if not v.bounds:
                return VArray([], [])  # 動的配列（ReDim まで空）
            fr = fr or Frame(None, module)
            base = self.modules.get(module.lower(), P.Module(module)).option_base
            lbs, ubs = [], []
            for lo, hi in v.bounds:
                lbs.append(to_long(self.eval(lo, fr)) if lo is not None else base)
                ubs.append(to_long(self.eval(hi, fr)))
            return VArray(lbs, ubs)
        if v.new:
            return self.host.new_object(v.type or "")
        t = (v.type or "variant").lower()
        if t in ("integer", "long", "longlong", "byte", "single", "double", "currency", "decimal"):
            return 0
        if t == "string":
            return ""
        if t == "boolean":
            return False
        if t == "date":
            return VDate(0)
        m = self.modules.get(module.lower())
        for mm in self.modules.values():
            if t in mm.types:
                return UserType(t, {f: EMPTY for f in mm.types[t]})
        if t in ("object", "range", "worksheet", "workbook", "collection", "dictionary") or (m and t not in ("variant",)):
            return NOTHING if t != "variant" else EMPTY
        return EMPTY

    # ---------------------------------------------------------------- 文
    def exec_body(self, body: list[P.Stmt], fr: Frame, top: bool = False) -> None:
        if not top:
            for st in body:
                self.exec_stmt(st, fr)
            return
        # いちばん外の文の並び：GoTo とエラーの行き先（ラベル）を扱う
        labels = {st.name: i for i, st in enumerate(body) if isinstance(st, P.Label)}
        i = 0
        while i < len(body):
            st = body[i]
            try:
                self.exec_stmt(st, fr)
                i += 1
            except _GoTo as g:
                if g.label not in labels:
                    raise Unsupported(f"GoTo の行き先「{g.label}」が、手続きのいちばん外にありません", st.line, fr.module)
                i = labels[g.label]
            except VbaError as e:
                self.err.set(e)
                if fr.on_error == "resume_next":
                    i += 1
                    continue
                if fr.on_error == "goto" and not fr.in_handler:
                    if fr.error_label not in labels:
                        raise Unsupported(f"On Error GoTo の行き先「{fr.error_label}」が、手続きのいちばん外にありません",
                                          st.line, fr.module)
                    fr.resume_at = i + 1
                    fr.in_handler = True
                    i = labels[fr.error_label]
                    continue
                raise
            except _ResumeSignal as r:
                fr.in_handler = False
                self.err.clear()
                if r.mode == "next":
                    i = fr.resume_at if fr.resume_at is not None else i + 1
                elif r.mode == "label":
                    i = labels[r.label]
                else:
                    i = (fr.resume_at - 1) if fr.resume_at is not None else i

    def make_string(self, n: int, what: str = "文字列") -> None:
        """長さ n の文字列を作る前に：1 つの文字列の長さ（limits.MAX_VBA_STRING）と、この実行で作る大きな文字列の合計
        （limits.MAX_VBA_BYTES）の上限を確かめる（超えれば Unsupported：Excel でもメモリが足りなくなる大きさ）。"""
        from ..limits import MAX_VBA_BYTES, MAX_VBA_STRING

        if n > MAX_VBA_STRING:
            raise Unsupported(f"とても長い{what}（{n:,} 字。{MAX_VBA_STRING:,} 字まで再現します）")
        if n >= 1 << 20:
            self.big_chars += n
            if self.big_chars > MAX_VBA_BYTES:
                raise Unsupported(f"大きな文字列をたくさん作るマクロ（合計 {MAX_VBA_BYTES:,} 字まで再現します）")

    def tick(self, st: P.Stmt, fr: Frame) -> None:
        self.steps += 1
        if self.steps > self.max_steps:
            raise Unsupported("繰り返しが多すぎます（終わらないマクロかもしれません）", st.line, fr.module)

    def exec_stmt(self, st: P.Stmt, fr: Frame) -> None:
        self.tick(st, fr)
        try:
            self._exec(st, fr)
        except VbaError as e:
            if not getattr(e, "where", None):
                e.where = (fr.module, st.line)  # はじめにエラーになった文（モジュールと行）
            raise
        except (_ExitProc, _ExitLoop, _GoTo, _EndProgram, NeedsInput, Unsupported, _ResumeSignal):
            raise
        except RecursionError:
            raise VbaError(28, "スタック領域が不足しています")

    def _exec(self, st: P.Stmt, fr: Frame) -> None:
        if isinstance(st, P.Assign):
            v = self.eval(st.value, fr)
            if not st.set and _is_object(v) and not isinstance(v, (VArray, UserType)) and hasattr(v, "vba_default"):
                v = v.vba_default()
            self.assign(st.target, v, fr, set_=st.set)
            return
        if isinstance(st, P.CallStmt):
            self.eval(st.expr, fr, statement=True)
            return
        if isinstance(st, P.If):
            for cond, body in st.branches:
                if to_bool(self.value(cond, fr)):
                    self.exec_body(body, fr)
                    return
            if st.else_body is not None:
                self.exec_body(st.else_body, fr)
            return
        if isinstance(st, P.For):
            self.exec_for(st, fr)
            return
        if isinstance(st, P.ForEach):
            coll = self.eval(st.coll, fr)
            items = self.host.iterate(coll) if not isinstance(coll, (VArray, Collection, Dictionary)) else _iter_items(coll)
            try:
                for item in items:
                    if not st.body:
                        self.tick(st, fr)
                    self.assign(st.var, item, fr, set_=_is_object(item))
                    try:
                        self.exec_body(st.body, fr)
                    except _ExitLoop as e:
                        if e.what == "for":
                            return
                        raise
            except _ExitLoop as e:
                if e.what != "for":
                    raise
            return
        if isinstance(st, P.Do):
            self.exec_do(st, fr)
            return
        if isinstance(st, P.Select):
            v = self.value(st.expr, fr)
            text = self._text_compare(fr)
            for conds, body in st.cases:
                if conds is None or any(self._case_match(v, c, fr, text) for c in conds):
                    self.exec_body(body, fr)
                    return
            return
        if isinstance(st, P.With):
            obj = self.eval(st.obj, fr)
            fr.with_stack.append(obj)
            try:
                self.exec_body(st.body, fr)
            finally:
                fr.with_stack.pop()
            return
        if isinstance(st, P.Dim):
            for v in st.vars:
                key = v.name
                if st.static:
                    store = self.statics.setdefault(f"{fr.module}.{fr.proc.name if fr.proc else ''}", {})
                    if key not in store:
                        store[key] = Ref(self._initial(v, fr.module, fr=fr))
                    fr.locals[key] = store[key]
                    continue
                fr.locals[key] = Ref(self._initial(v, fr.module, const=st.const, fr=fr), v.type)
            return
        if isinstance(st, P.ReDim):
            for target, bounds in st.vars:
                ref = self.lookup_ref(target, fr, create=True)
                base = self.modules.get(fr.module.lower(), P.Module(fr.module)).option_base
                lbs = [to_long(self.value(lo, fr)) if lo is not None else base for lo, _hi in bounds]
                ubs = [to_long(self.value(hi, fr)) for _lo, hi in bounds]
                new = VArray(lbs, ubs)
                old = ref.value
                if st.preserve and isinstance(old, VArray):
                    for k, val in old.data.items():
                        if all(lo <= i <= hi for i, lo, hi in zip(k, lbs, ubs)):
                            new.data[k] = val
                ref.value = new
            return
        if isinstance(st, P.Exit):
            if st.what in ("sub", "function", "property"):
                raise _ExitProc()
            raise _ExitLoop(st.what)
        if isinstance(st, P.GoTo):
            raise _GoTo(st.label)
        if isinstance(st, P.Label):
            return
        if isinstance(st, P.OnError):
            if st.mode == "resume_next":
                fr.on_error = "resume_next"
            elif st.mode == "goto0":
                fr.on_error = "none"
                fr.in_handler = False
            else:
                fr.on_error, fr.error_label = "goto", st.label
            self.err.clear()
            return
        if isinstance(st, P.Resume):
            raise _ResumeSignal(st.mode, st.label)
        if isinstance(st, P.End):
            raise _EndProgram()
        if isinstance(st, P.Erase):
            for n in st.names:
                ref = self.lookup_ref(n, fr)
                if isinstance(ref.value, VArray):
                    if ref.value.lbounds and ref.is_array_decl:
                        ref.value.data.clear()
                    else:
                        ref.value = VArray([], [])
            return
        if isinstance(st, P.FileStmt):
            self.file_stmt(st, fr)
            return
        raise Unsupported(f"この文（{type(st).__name__}）は再現しません", st.line, fr.module)

    def _text_compare(self, fr: Frame) -> bool:
        m = self.modules.get(fr.module.lower())
        return bool(m and m.option_compare_text)

    def _case_match(self, v, c, fr, text) -> bool:
        if c[0] == "eq":
            return compare(v, self.value(c[1], fr), text) == 0
        if c[0] == "to":
            return compare(v, self.value(c[1], fr), text) >= 0 and compare(v, self.value(c[2], fr), text) <= 0
        op, rhs = c[1], self.value(c[2], fr)
        r = compare(v, rhs, text)
        return {"=": r == 0, "<>": r != 0, "<": r < 0, ">": r > 0, "<=": r <= 0, ">=": r >= 0}[op]

    def exec_for(self, st: P.For, fr: Frame) -> None:
        start = self.value(st.start, fr)
        end = self.value(st.end, fr)
        step = self.value(st.step, fr) if st.step is not None else 1
        x = start
        if isinstance(start, str):
            x = to_double(start)
        endv, stepv = to_double(end), to_double(step)
        self.assign(st.var, num_result(to_double(x)), fr)
        while True:
            if not st.body:  # 中身のない繰り返しも数える（文のある繰り返しは文を数える）
                self.tick(st, fr)
            cur = to_double(self.value(st.var, fr))
            if (stepv >= 0 and cur > endv) or (stepv < 0 and cur < endv):
                return
            try:
                self.exec_body(st.body, fr)
            except _ExitLoop as e:
                if e.what == "for":
                    return
                raise
            cur = to_double(self.value(st.var, fr)) + stepv
            self.assign(st.var, num_result(cur), fr)

    def exec_do(self, st: P.Do, fr: Frame) -> None:
        def ok(cond) -> bool:
            if cond is None:
                return True
            w, e = cond
            b = to_bool(self.value(e, fr)) if self.value(e, fr) is not NULL else False
            return b if w == "while" else not b

        while True:
            if not st.body:  # 中身のない繰り返しも数える（文のある繰り返しは文を数える）
                self.tick(st, fr)
            if st.pre is not None and not ok(st.pre):
                return
            try:
                self.exec_body(st.body, fr)
            except _ExitLoop as e:
                if e.what in ("do", "while"):
                    return
                raise
            if st.post is not None and not ok(st.post):
                return

    # ---------------------------------------------------------------- 代入
    def assign(self, target: P.Node, v, fr: Frame, set_: bool = False) -> None:
        if isinstance(target, P.Name):
            ref = self.lookup_ref(target, fr, create=True)
            if ref is None:
                # Range("A1") の既定のメンバーなど、host の名前への代入
                obj = self.host.global_object(target.name)
                if obj is not None and hasattr(obj, "vba_set"):
                    obj.vba_set("value", v)
                    return
                raise VbaError(424)
            ref.value = _coerce(v, ref.type) if not set_ else v
            return
        if isinstance(target, P.Call):
            # 配列の要素、または Range("A1") = 値（既定のメンバー）
            base = target.target
            ref = self.lookup_ref(base, fr) if isinstance(base, P.Name) else None
            if ref is not None and isinstance(ref.value, VArray):
                idx = tuple(to_long(self.value(a.value, fr)) for a in target.args)
                ref.value.set(idx, v)
                return
            if ref is not None and isinstance(ref.value, (Collection, Dictionary)):
                key = self.value(target.args[0].value, fr)
                _coll_set(ref.value, key, v)
                return
            obj = self.eval(target, fr)
            if hasattr(obj, "vba_set"):
                obj.vba_set("value", v)
                return
            if isinstance(base, P.Member):
                parent = self.eval(base.obj, fr) if base.obj is not None else fr.with_stack[-1]
                args = [self.value(a.value, fr) for a in target.args]
                if hasattr(parent, "vba_set_indexed"):
                    parent.vba_set_indexed(base.name, args, v)
                    return
            raise Unsupported("この代入は再現しません", target.line, fr.module)
        if isinstance(target, P.Member):
            obj = self.eval(target.obj, fr) if target.obj is not None else (fr.with_stack[-1] if fr.with_stack else None)
            if obj is None:
                raise VbaError(91)
            if isinstance(obj, UserType):
                obj.fields[target.name] = v
                return
            if isinstance(obj, Dictionary) and target.bang:
                obj.d[obj.key(target.raw)] = v
                return
            if hasattr(obj, "vba_set"):
                obj.vba_set(target.name, v)
                return
            if obj is NOTHING:
                raise VbaError(91)
            raise VbaError(438)
        raise Unsupported("この代入は再現しません", target.line, fr.module)

    def lookup_ref(self, node: P.Node, fr: Frame, create: bool = False) -> Ref | None:
        if not isinstance(node, P.Name):
            return None
        n = node.name
        if n in fr.locals:
            return fr.locals[n]
        if fr.proc is not None and n == fr.proc.name and "__ret__" in fr.locals:
            return fr.locals["__ret__"]
        mv = self.module_vars.get(fr.module.lower(), {})
        if n in mv:
            return mv[n]
        for name, vars_ in self.module_vars.items():
            if n in vars_ and name != fr.module.lower():
                return vars_[n]
        if create and not self._is_known_name(n, fr):
            r = Ref(EMPTY)
            fr.locals[n] = r  # Option Explicit のないモジュールの、宣言なしの変数
            return r
        return None

    def _is_known_name(self, n: str, fr: Frame) -> bool:
        return self.find_proc(n, fr.module) is not None or self.host.global_object(n) is not None

    # ---------------------------------------------------------------- 式
    def value(self, node, fr: Frame):
        v = self.eval(node, fr)
        if hasattr(v, "vba_default") and not isinstance(v, (VArray, Collection, Dictionary, UserType)):
            return v.vba_default()
        return v

    def eval(self, node, fr: Frame, statement: bool = False):
        if node is None:
            return MISSING
        if isinstance(node, P.Lit):
            return node.value
        if isinstance(node, P.DateLit):
            d = parse_date_text(node.text)
            if d is None:
                raise Unsupported(f"日付「#{node.text}#」を読めません", node.line, fr.module)
            return d
        if isinstance(node, P.Name):
            return self.eval_name(node, fr, statement)
        if isinstance(node, P.Member):
            return self.eval_member(node, fr, [], statement)
        if isinstance(node, P.Call):
            return self.eval_call(node, fr, statement)
        if isinstance(node, P.Unary):
            v = self.value(node.operand, fr)
            if node.op == "-":
                if v is NULL:
                    return NULL
                return num_result(-to_double(v)) if not isinstance(v, VDate) else VDate(-float(v))
            if node.op == "+":
                return v
            if node.op == "not":
                if v is NULL:
                    return NULL
                if isinstance(v, bool):
                    return not v
                return ~to_long(v)
        if isinstance(node, P.Binary):
            return self.eval_binary(node, fr)
        if isinstance(node, P.New):
            return self.host.new_object(node.cls)
        raise Unsupported(f"この式（{type(node).__name__}）は再現しません", getattr(node, "line", 0), fr.module)

    def eval_binary(self, node: P.Binary, fr: Frame):
        op = node.op
        if op == "is":
            a, b = self.eval(node.left, fr), self.eval(node.right, fr)
            return a is b or (a is NOTHING and b is NOTHING) or (hasattr(a, "same") and a.same(b))
        a = self.value(node.left, fr)
        b = self.value(node.right, fr)
        if op == "&":
            if a is NULL and b is NULL:
                return NULL
            sa, sb = (to_str(a) if a is not NULL else ""), (to_str(b) if b is not NULL else "")
            self.make_string(len(sa) + len(sb))
            return sa + sb
        if a is NULL or b is NULL:
            return NULL
        text = self._text_compare(fr)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            r = compare(a, b, text)
            return {"=": r == 0, "<>": r != 0, "<": r < 0, ">": r > 0, "<=": r <= 0, ">=": r >= 0}[op]
        if op == "like":
            return like(to_str(a), to_str(b)) if not text else like(to_str(a).lower(), to_str(b).lower())
        if op in ("and", "or", "xor", "eqv", "imp", "andalso", "orelse"):
            if isinstance(a, bool) and isinstance(b, bool):
                return {"and": a and b, "andalso": a and b, "or": a or b, "orelse": a or b, "xor": a != b,
                        "eqv": a == b, "imp": (not a) or b}[op]
            x, y = to_long(a), to_long(b)
            return {"and": x & y, "andalso": x & y, "or": x | y, "orelse": x | y, "xor": x ^ y, "eqv": ~(x ^ y),
                    "imp": (~x) | y}[op]
        if op == "+":
            if isinstance(a, str) and isinstance(b, str):
                return a + b
            if a is EMPTY and isinstance(b, str):
                return b
            if b is EMPTY and isinstance(a, str):
                return a
            if isinstance(a, VDate) or isinstance(b, VDate):
                return VDate(to_double(a) + to_double(b))
            return num_result(to_double(a) + to_double(b))
        if op == "-":
            if isinstance(a, VDate) and not isinstance(b, VDate):
                return VDate(to_double(a) - to_double(b))
            return num_result(to_double(a) - to_double(b))
        if op == "*":
            return num_result(to_double(a) * to_double(b))
        if op == "/":
            y = to_double(b)
            if y == 0:
                raise VbaError(11)
            return to_double(a) / y
        if op == "\\":
            y = to_long(b)
            if y == 0:
                raise VbaError(11)
            x = to_long(a)
            q = abs(x) // abs(y)
            return q if (x >= 0) == (y >= 0) else -q
        if op == "mod":
            y = to_long(b)
            if y == 0:
                raise VbaError(11)
            x = to_long(a)
            return int(math.fmod(x, y))
        if op == "^":
            return float(to_double(a) ** to_double(b))
        raise Unsupported(f"演算「{op}」", node.line, fr.module)

    def eval_name(self, node: P.Name, fr: Frame, statement: bool = False):
        n = node.name
        ref = self.lookup_ref(node, fr)
        if ref is not None:
            return ref.value
        if n in ("me",):
            return self.host.me(fr.module)
        if n == "err":
            return self.err
        c = _constant(n, self)
        if c is not MISSING:
            return c
        p = self.find_proc(n, fr.module)
        if p is not None:
            return self.call_proc(p, [])
        from . import builtins

        if n in builtins.NOARG:
            return builtins.call(self, n, [], fr, node)
        obj = self.host.global_object(n)
        if obj is not None:
            return obj
        if n in builtins.FUNCS:
            return builtins.call(self, n, [], fr, node)
        raise Unsupported(f"名前「{node.raw}」が分かりません", node.line, fr.module)

    def eval_member(self, node: P.Member, fr: Frame, args: list, statement: bool = False, named: dict | None = None):
        if isinstance(node.obj, P.Name) and node.obj.name in self.modules and self.lookup_ref(node.obj, fr) is None \
                and not self.host.global_object(node.obj.name):
            # 「Call Module1.Proc」：モジュールの名前で手続きを呼ぶ（変数・モジュール変数も同じ形）
            m = self.modules[node.obj.name]
            p = m.procs.get(node.name) or m.procs.get(f"get:{node.name}")
            if p is not None:
                return self.call_proc(p, [Ref(a) for a in args], {k: Ref(v) for k, v in (named or {}).items()})
            if node.name in self.module_vars.get(node.obj.name, {}):
                return self.module_vars[node.obj.name][node.name].value
            c = m.enums.get(node.name)
            if c is not None:
                return c
        if node.obj is None:
            if not fr.with_stack:
                raise VbaError(91)
            obj = fr.with_stack[-1]
        else:
            obj = self.eval(node.obj, fr)
        if obj is NOTHING or obj is EMPTY:
            if node.obj is not None and isinstance(node.obj, P.Name) and node.obj.name == "debug":
                return EMPTY
            raise VbaError(91)
        if isinstance(obj, UserType):
            return obj.fields.get(node.name, EMPTY)
        if isinstance(obj, ErrObject):
            return obj.member(node.name, args)
        if isinstance(obj, (Collection, Dictionary)):
            return _coll_member(self, obj, node, args)
        if isinstance(obj, _Debug):
            return EMPTY
        if hasattr(obj, "vba_get"):
            return obj.vba_get(node.name, args, named or {})
        raise Unsupported(f"「{node.raw}」を持たない値です", node.line, fr.module)

    def eval_call(self, node: P.Call, fr: Frame, statement: bool = False):
        target = node.target
        # 引数：変数そのものは参照（ByRef）で渡せるようにしておく
        if isinstance(target, P.Name):
            n = target.name
            ref = self.lookup_ref(target, fr)
            if ref is not None:
                v = ref.value
                if isinstance(v, VArray):
                    idx = tuple(to_long(self.value(a.value, fr)) for a in node.args)
                    return v.get(idx)
                if isinstance(v, (Collection, Dictionary)):
                    return _coll_item(v, self.value(node.args[0].value, fr))
                if hasattr(v, "vba_get"):
                    return v.vba_get("item", [self.value(a.value, fr) for a in node.args], {})
                if not node.args:
                    return v
            p = self.find_proc(n, fr.module)
            if p is not None:
                refs, named = self.arg_refs(node.args, fr, p)
                return self.call_proc(p, refs, named)
            if n == "__typeof__":
                obj = self.eval(node.args[0].value, fr)
                return self.host.type_of(obj, node.args[1].value.value)
            from . import builtins

            if n in builtins.FUNCS:
                return builtins.call(self, n, node.args, fr, node)
            obj = self.host.global_object(n)
            if obj is not None and hasattr(obj, "vba_call"):
                # オブジェクト（Range など）はオブジェクトのまま渡す。値が要る所では、受け取る側で値に直す
                return obj.vba_call([self.eval(a.value, fr) if a.value is not None else MISSING for a in node.args
                                     if a.name is None],
                                    {a.name: self.eval(a.value, fr) for a in node.args if a.name})
            raise Unsupported(f"手続き・関数「{target.raw}」が分かりません", node.line, fr.module)
        if isinstance(target, P.Member):
            args = [self.eval(a.value, fr) if a.value is not None else MISSING for a in node.args if a.name is None]
            named = {a.name: self.eval(a.value, fr) for a in node.args if a.name}
            # Module1.Proc の形の呼び出し
            if isinstance(target.obj, P.Name) and target.obj.name in self.modules and self.lookup_ref(target.obj, fr) is None:
                p = self.modules[target.obj.name].procs.get(target.name)
                if p is not None:
                    refs, nm = self.arg_refs(node.args, fr, p)
                    return self.call_proc(p, refs, nm)
            args = [a.vba_default() if hasattr(a, "vba_default") and _wants_value(a) else a for a in args]
            return self.eval_member(target, fr, args, statement, named)
        if isinstance(target, P.Call):
            # Range("A1:B2")(1, 2) のような、結果をさらに添字で引くもの
            obj = self.eval(target, fr)
            args = [self.value(a.value, fr) for a in node.args]
            if isinstance(obj, VArray):
                return obj.get(tuple(to_long(x) for x in args))
            if hasattr(obj, "vba_get"):
                return obj.vba_get("item", args, {})
        raise Unsupported("この呼び出しは再現しません", node.line, fr.module)

    def arg_refs(self, args: list[P.Arg], fr: Frame, p: P.Proc) -> tuple[list[Ref], dict]:
        refs, named = [], {}
        for a in args:
            if a.value is None:
                r = Ref(MISSING)
            elif isinstance(a.value, P.Name) and not a.byval_paren and self.lookup_ref(a.value, fr) is not None:
                r = self.lookup_ref(a.value, fr)
            else:
                v = self.eval(a.value, fr)
                r = Ref(v)
            if a.name:
                named[a.name] = r
            else:
                refs.append(r)
        return refs, named

    # ---------------------------------------------------------------- ファイル
    def file_stmt(self, st: P.FileStmt, fr: Frame) -> None:
        if st.kind == "open":
            path, mode, num = st.args
            p = to_str(self.value(path, fr))
            n = to_long(self.value(num, fr))
            if mode in ("output", "append", "binary") and mode != "binary":
                buf = self.files_out.setdefault(p, bytearray())
                if mode == "output":
                    buf.clear()
                self.open_files[n] = {"path": p, "mode": mode, "buf": buf}
                return
            if mode == "input":
                if p not in self.inputs:
                    raise NeedsInput("file", f"マクロがファイル「{p}」を読もうとしています。中身を渡してください", st.line)
                data = self.inputs[p]
                text = data.decode("cp932", errors="replace") if isinstance(data, bytes) else str(data)
                self.open_files[n] = {"path": p, "mode": "input", "lines": text.replace("\r\n", "\n").split("\n"),
                                      "pos": 0}
                return
            raise Unsupported(f"ファイルを「{mode}」で開く", st.line, fr.module)
        if st.kind == "close":
            nums = [to_long(self.value(a, fr)) for a in st.args] or list(self.open_files)
            for n in nums:
                self.open_files.pop(n, None)
            return
        if st.kind == "debugprint":
            # イミディエイト ウィンドウは、書類にも人にも出ない。式は、Excel と同じく計算する（呼んだ関数が動き、エラーも出る）
            for it in st.args[1]:
                if isinstance(it, tuple):
                    if it[1] is not None:
                        self.value(it[1], fr)
                elif it not in (";", ","):
                    self.value(it, fr)
            return
        if st.kind in ("print", "write"):
            num, items = st.args
            f = self.open_files.get(to_long(self.value(num, fr)))
            if f is None or f["mode"] not in ("output", "append"):
                raise VbaError(52, "ファイル名または番号が不正です")
            parts, newline = [], True
            for k, it in enumerate(items):
                if isinstance(it, tuple):
                    # Spc(n)：n 個の空白。Tab(n)：n 桁目へ（過ぎていれば次の行）。Tab だけなら次の区切り（14 桁ごと）
                    col = len("".join(parts).rsplit("\n", 1)[-1])
                    if it[0] == "spc":
                        parts.append(" " * max(0, to_long(self.value(it[1], fr))))
                    elif it[1] is None:
                        parts.append(" " * (14 - col % 14))
                    else:
                        n = max(1, to_long(self.value(it[1], fr)))
                        parts.append(" " * (n - 1 - col) if col < n else "\r\n" + " " * (n - 1))
                    continue
                if it in (";", ","):
                    if k == len(items) - 1:
                        newline = False
                    if it == "," and st.kind == "print":
                        parts.append("\t")
                    continue
                v = self.value(it, fr)
                if st.kind == "write":
                    s = f'"{to_str(v)}"' if isinstance(v, str) else (
                        f"#{to_str(v).upper()}#" if isinstance(v, bool) else to_str(v))
                    parts.append(s)
                else:
                    parts.append(to_str(v) if not isinstance(v, (int, float)) or isinstance(v, bool) else
                                 (" " + to_str(v) if v >= 0 else to_str(v)) + " ")
            text = ",".join(parts) if st.kind == "write" else "".join(parts)
            f["buf"] += (text + ("\r\n" if newline else "")).encode("cp932", errors="replace")
            return
        if st.kind in ("lineinput", "input"):
            num = st.args[0]
            f = self.open_files.get(to_long(self.value(num, fr)))
            if f is None or f["mode"] != "input":
                raise VbaError(52, "ファイル名または番号が不正です")
            if f["pos"] >= len(f["lines"]):
                raise VbaError(62, "ファイルにこれ以上データがありません")
            line = f["lines"][f["pos"]]
            f["pos"] += 1
            if st.kind == "lineinput":
                self.assign(st.args[1], line, fr)
                return
            values = _split_csv(line)
            for target, v in zip(st.args[1], values):
                self.assign(target, v, fr)
            return
        raise Unsupported(f"ファイルの文（{st.kind}）", st.line, fr.module)

    def eof(self, n: int) -> bool:
        f = self.open_files.get(n)
        if f is None or f["mode"] != "input":
            raise VbaError(52)
        rest = f["lines"][f["pos"]:]
        return not rest or (len(rest) == 1 and rest[0] == "")


class _ResumeSignal(Exception):
    def __init__(self, mode: str, label: str):
        self.mode = mode
        self.label = label


class _Debug:
    pass


class ErrObject:
    def __init__(self):
        self.clear()

    def clear(self):
        self.number = 0
        self.description = ""
        self.source = ""

    def set(self, e: VbaError):
        self.number = e.number
        self.description = e.description

    def member(self, name: str, args: list):
        if name == "number":
            return self.number
        if name == "description":
            return self.description
        if name == "source":
            return self.source
        if name == "clear":
            self.clear()
            return EMPTY
        if name == "raise":
            raise VbaError(to_long(args[0]), to_str(args[2]) if len(args) > 2 and args[2] is not MISSING else "")
        raise VbaError(438)

    def vba_default(self):
        return self.number


def _split_csv(line: str) -> list:
    out, cur, q = [], [], False
    i = 0
    while i < len(line):
        ch = line[i]
        if q:
            if ch == '"':
                if i + 1 < len(line) and line[i + 1] == '"':
                    cur.append('"')
                    i += 2
                    continue
                q = False
            else:
                cur.append(ch)
        elif ch == '"':
            q = True
        elif ch == ",":
            out.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


def _copy(v):
    if isinstance(v, VArray):
        return VArray(list(v.lbounds), list(v.ubounds), dict(v.data))
    if isinstance(v, UserType):
        return UserType(v.name, dict(v.fields))
    return v


def _coerce(v, t: str | None):
    if t is None or v is EMPTY:
        return v
    t = t.lower()
    try:
        if t in ("integer", "long", "byte", "longlong"):
            return to_long(v)
        if t in ("double", "single", "currency", "decimal"):
            return to_double(v)
        if t == "string":
            return to_str(v)
        if t == "boolean":
            return to_bool(v)
        if t == "date":
            return VDate(to_double(v))
    except VbaError:
        raise
    return v


def _is_object(v) -> bool:
    return not isinstance(v, (int, float, str, bool, VArray)) and v not in (EMPTY, NULL, MISSING) and v is not None


def _wants_value(v) -> bool:
    return False


def _iter_items(coll) -> list:
    if isinstance(coll, VArray):
        return coll.flat()
    if isinstance(coll, Collection):
        return list(coll.items)
    if isinstance(coll, Dictionary):
        return list(coll.d.keys())
    return []


def _coll_item(c, key):
    if isinstance(c, Collection):
        if isinstance(key, str):
            if key.lower() not in c.keys:
                raise VbaError(5)
            return c.items[c.keys[key.lower()]]
        i = to_long(key)
        if not 1 <= i <= len(c.items):
            raise VbaError(9)
        return c.items[i - 1]
    k = c.key(key)
    if k not in c.d:
        c.d[k] = EMPTY  # Dictionary は、ない鍵を引くと空の値で足す
    return c.d[k]


def _coll_set(c, key, v):
    if isinstance(c, Dictionary):
        c.d[c.key(key)] = v
        return
    raise VbaError(438)


def _coll_member(rt: Runtime, obj, node: P.Member, args: list):
    name = node.name
    if isinstance(obj, Collection):
        if name == "add":
            item = args[0]
            key = args[1] if len(args) > 1 and args[1] is not MISSING else None
            if key is not None:
                if to_str(key).lower() in obj.keys:
                    raise VbaError(457, "このキーは既にこのコレクションの要素に割り当てられています")
                obj.keys[to_str(key).lower()] = len(obj.items)
            obj.items.append(item)
            return EMPTY
        if name == "count":
            return len(obj.items)
        if name == "item":
            return _coll_item(obj, args[0])
        if name == "remove":
            k = args[0]
            i = obj.keys.pop(to_str(k).lower()) if isinstance(k, str) else to_long(k) - 1
            obj.items.pop(i)
            obj.keys = {kk: (j if j < i else j - 1) for kk, j in obj.keys.items()}
            return EMPTY
    if isinstance(obj, Dictionary):
        if node.bang:
            return obj.d.get(obj.key(node.raw), EMPTY)
        if name == "add":
            k = obj.key(args[0])
            if k in obj.d:
                raise VbaError(457, "このキーは既にこのコレクションの要素に割り当てられています")
            obj.d[k] = args[1]
            return EMPTY
        if name == "exists":
            return obj.key(args[0]) in obj.d
        if name == "count":
            return len(obj.d)
        if name == "item":
            return _coll_item(obj, args[0])
        if name == "keys":
            ks = list(obj.d.keys())
            return VArray([0], [len(ks) - 1], {(i,): k for i, k in enumerate(ks)})
        if name == "items":
            vs = list(obj.d.values())
            return VArray([0], [len(vs) - 1], {(i,): v for i, v in enumerate(vs)})
        if name == "remove":
            obj.d.pop(obj.key(args[0]), None)
            return EMPTY
        if name == "removeall":
            obj.d.clear()
            return EMPTY
        if name == "comparemode":
            return 1 if obj.compare_text else 0
    raise VbaError(438)


def _constant(n: str, rt: Runtime):
    from .builtins import CONSTANTS

    if n in CONSTANTS:
        return CONSTANTS[n]
    for m in rt.modules.values():
        if n in m.enums:
            return m.enums[n]
    return MISSING
