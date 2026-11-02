"""VBA のソースを、字句に分け、構文木にする（MS-VBAL の文法の、申請書類のマクロで使われる範囲）。

字句：識別子（日本語も）・数（10進・&H・&O・型の印）・文字列（"" は " ）・日付（#…#）・演算子・コメント（' と Rem）・
行の継続（行末の " _"）・文の区切り（:）。大文字と小文字は区別しない（キーワード・名前とも小文字にして比べる）。
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


class VbaSyntaxError(Exception):
    def __init__(self, msg: str, line: int = 0):
        super().__init__(f"{line} 行目: {msg}" if line else msg)
        self.line = line


# ====================================================================== 字句
@dataclass
class Tok:
    kind: str  # id / num / str / date / op / nl / eof
    value: object
    line: int
    space_before: bool = False  # 前に空白があるか（「a (1)」と「a(1)」の区別に使う）

    def is_op(self, *ops: str) -> bool:
        return self.kind == "op" and self.value in ops

    def is_kw(self, *kws: str) -> bool:
        return self.kind == "id" and self.value.lower() in kws


_NUM = re.compile(r"(?:&[Hh][0-9A-Fa-f]+|&[Oo]?[0-7]+|(?:\d+\.?\d*|\.\d+)(?:[eEdD][+-]?\d+)?)[%&!#@^]?")
_ID = re.compile(r"[^\W\d][\w]*[$%&!#@]?|\[[^\]\n]+\]", re.UNICODE)
_OPS = ["<=", ">=", "<>", ":=", "=", "<", ">", "+", "-", "*", "/", "\\", "^", "&", ".", ",", "(", ")", "!", ";", ":", "#"]


def tokenize(src: str) -> list[Tok]:
    src = src.replace("\r\n", "\n").replace("\r", "\n")
    out: list[Tok] = []
    i, line, n = 0, 1, len(src)
    while i < n:
        ch = src[i]
        sp = i > 0 and src[i - 1] in " \t"
        if ch in " \t":
            i += 1
            continue
        if ch == "_" and (i + 1 == n or src[i + 1] == "\n") and sp:
            i += 2  # 行の継続
            line += 1
            continue
        if ch == "\n":
            out.append(Tok("nl", "\n", line))
            line += 1
            i += 1
            continue
        if ch == "'":
            while i < n and src[i] != "\n":
                # コメントも、行末の " _" で続く
                if src[i] == "_" and i + 1 < n and src[i + 1] == "\n" and src[i - 1] in " \t":
                    i += 2
                    line += 1
                    continue
                i += 1
            continue
        if ch == '"':
            j = i + 1
            buf = []
            while j < n:
                if src[j] == '"':
                    if j + 1 < n and src[j + 1] == '"':
                        buf.append('"')
                        j += 2
                        continue
                    break
                if src[j] == "\n":
                    raise VbaSyntaxError("文字列が閉じていません", line)
                buf.append(src[j])
                j += 1
            out.append(Tok("str", "".join(buf), line, sp))
            i = j + 1
            continue
        if ch == "#" and (not out or out[-1].kind in ("op", "nl")) and not (out and out[-1].is_op(")")):
            m = re.match(r"#([^#\n]+)#", src[i:])
            if m and re.search(r"\d", m.group(1)) and not re.match(r"\s*\d+\s*,", m.group(1)):
                out.append(Tok("date", m.group(1).strip(), line, sp))
                i += m.end()
                continue
        if ch.isdigit() or (ch == "." and i + 1 < n and src[i + 1].isdigit() and not (out and out[-1].kind == "id")) or \
                (ch == "&" and i + 1 < n and src[i + 1] in "HhOo0123456789" and not (out and out[-1].kind in ("id", "num", "str"))):
            m = _NUM.match(src, i)
            if m:
                out.append(Tok("num", _number(m.group(0)), line, sp))
                i = m.end()
                continue
        m = _ID.match(src, i)
        if m:
            word = m.group(0)
            if word.lower() == "rem" and (not out or out[-1].kind == "nl" or out[-1].is_op(":")):
                while i < n and src[i] != "\n":
                    i += 1
                continue
            out.append(Tok("id", word.strip("[]"), line, sp))
            i = m.end()
            continue
        for op in _OPS:
            if src.startswith(op, i):
                out.append(Tok("op", op, line, sp))
                i += len(op)
                break
        else:
            raise VbaSyntaxError(f"読めない字「{ch}」", line)
    out.append(Tok("nl", "\n", line))
    out.append(Tok("eof", None, line))
    return out


def _number(s: str):
    t = s.rstrip("%&!#@^")
    suffix = s[len(t):]
    if t[:2].lower() == "&h":
        v = int(t[2:], 16)
        if suffix in ("", "%") and v > 0x7FFF and v <= 0xFFFF:
            v -= 0x10000  # &HFFFF は Integer の -1
        elif suffix == "&" and v > 0x7FFFFFFF:
            v -= 0x100000000
        return v
    if t[:1] == "&":
        return int(t.lstrip("&Oo") or "0", 8)
    t = t.replace("d", "e").replace("D", "e")
    if suffix in ("!", "#", "@") or "." in t or "e" in t.lower():
        return float(t)
    return int(t)


# ====================================================================== 構文木
@dataclass
class Node:
    line: int = 0


@dataclass
class Lit(Node):
    value: object = None


@dataclass
class DateLit(Node):
    text: str = ""


@dataclass
class Name(Node):
    name: str = ""  # 小文字にしたもの
    raw: str = ""


@dataclass
class Member(Node):
    obj: Node | None = None  # None なら With の対象
    name: str = ""
    raw: str = ""
    bang: bool = False  # obj!name


@dataclass
class Arg:
    value: Node | None  # None は省略
    name: str | None = None  # 名前つきの引数（name:=value）
    byval_paren: bool = False  # 「(x)」のように、かっこで包んだ（値渡しになる）


@dataclass
class Call(Node):
    target: Node | None = None
    args: list[Arg] = field(default_factory=list)


@dataclass
class Unary(Node):
    op: str = ""
    operand: Node | None = None


@dataclass
class Binary(Node):
    op: str = ""
    left: Node | None = None
    right: Node | None = None


@dataclass
class New(Node):
    cls: str = ""


# 文
@dataclass
class Stmt(Node):
    pass


@dataclass
class DimVar:
    name: str
    bounds: list[tuple[Node | None, Node]] | None = None  # 配列の境界（下限, 上限）。() は動的配列（空のリスト）
    type: str | None = None
    new: bool = False
    init: Node | None = None  # Const の値


@dataclass
class Dim(Stmt):
    vars: list[DimVar] = field(default_factory=list)
    const: bool = False
    static: bool = False


@dataclass
class ReDim(Stmt):
    preserve: bool = False
    vars: list[tuple[Node, list[tuple[Node | None, Node]]]] = field(default_factory=list)


@dataclass
class Assign(Stmt):
    target: Node | None = None
    value: Node | None = None
    set: bool = False


@dataclass
class CallStmt(Stmt):
    expr: Node | None = None


@dataclass
class If(Stmt):
    branches: list[tuple[Node, list[Stmt]]] = field(default_factory=list)
    else_body: list[Stmt] | None = None


@dataclass
class For(Stmt):
    var: Node | None = None
    start: Node | None = None
    end: Node | None = None
    step: Node | None = None
    body: list[Stmt] = field(default_factory=list)


@dataclass
class ForEach(Stmt):
    var: Node | None = None
    coll: Node | None = None
    body: list[Stmt] = field(default_factory=list)


@dataclass
class Do(Stmt):
    pre: tuple[str, Node] | None = None  # ("while"|"until", 条件)
    post: tuple[str, Node] | None = None
    body: list[Stmt] = field(default_factory=list)


@dataclass
class Select(Stmt):
    expr: Node | None = None
    cases: list[tuple[list[tuple], list[Stmt]]] = field(default_factory=list)  # [(条件の並び, 文)]、条件 None は Case Else


@dataclass
class With(Stmt):
    obj: Node | None = None
    body: list[Stmt] = field(default_factory=list)


@dataclass
class Exit(Stmt):
    what: str = ""  # sub / function / property / for / do


@dataclass
class GoTo(Stmt):
    label: str = ""


@dataclass
class Label(Stmt):
    name: str = ""


@dataclass
class OnError(Stmt):
    mode: str = ""  # resume_next / goto0 / goto
    label: str = ""


@dataclass
class Resume(Stmt):
    mode: str = ""  # "" / next / label
    label: str = ""


@dataclass
class End(Stmt):
    pass


@dataclass
class FileStmt(Stmt):
    kind: str = ""  # open / close / print / write / input / lineinput
    args: list = field(default_factory=list)


@dataclass
class Erase(Stmt):
    names: list[Node] = field(default_factory=list)


@dataclass
class Param:
    name: str
    byval: bool = False
    optional: bool = False
    default: Node | None = None
    paramarray: bool = False
    is_array: bool = False


@dataclass
class Proc:
    name: str  # 小文字
    raw: str
    kind: str  # sub / function / get / let / set
    params: list[Param]
    body: list[Stmt]
    private: bool = False
    static: bool = False
    module: str = ""
    line: int = 0


@dataclass
class Module:
    name: str
    procs: dict[str, Proc] = field(default_factory=dict)  # 名前（小文字）→ 手続き（Property は「get:名前」など）
    decls: list[Dim] = field(default_factory=list)
    option_base: int = 0
    option_compare_text: bool = False
    types: dict[str, list[str]] = field(default_factory=dict)
    enums: dict[str, int] = field(default_factory=dict)


# ====================================================================== 構文の解析
_END_BLOCK = {"end", "else", "elseif", "next", "loop", "wend", "case"}


class Parser:
    def __init__(self, toks: list[Tok], module: str):
        self.t = toks
        self.i = 0
        self.module = module

    # ---------------------------------------------------------------- 基本
    @property
    def tok(self) -> Tok:
        return self.t[self.i]

    def peek(self, k: int = 1) -> Tok:
        return self.t[min(self.i + k, len(self.t) - 1)]

    def next(self) -> Tok:
        tk = self.t[self.i]
        self.i += 1
        return tk

    def expect_op(self, op: str) -> Tok:
        if not self.tok.is_op(op):
            raise VbaSyntaxError(f"「{op}」があるはずの所に「{self.tok.value}」", self.tok.line)
        return self.next()

    def expect_kw(self, *kws: str) -> Tok:
        if not self.tok.is_kw(*kws):
            raise VbaSyntaxError(f"「{'／'.join(kws)}」があるはずの所に「{self.tok.value}」", self.tok.line)
        return self.next()

    def at_end_of_stmt(self) -> bool:
        return self.tok.kind in ("nl", "eof") or self.tok.is_op(":") or self.tok.is_kw("else")

    def skip_nl(self) -> None:
        while self.tok.kind == "nl" or self.tok.is_op(":"):
            self.next()

    def ident(self) -> Tok:
        if self.tok.kind != "id":
            raise VbaSyntaxError(f"名前があるはずの所に「{self.tok.value}」", self.tok.line)
        return self.next()

    # ---------------------------------------------------------------- モジュール
    def module_(self) -> Module:
        m = Module(self.module)
        while self.tok.kind != "eof":
            self.skip_nl()
            if self.tok.kind == "eof":
                break
            tk = self.tok
            low = tk.value.lower() if tk.kind == "id" else ""
            if low == "attribute":
                self.skip_line()
                continue
            if low == "option":
                self.next()
                w = self.ident().value.lower()
                if w == "base":
                    m.option_base = int(self.next().value)
                elif w == "compare":
                    m.option_compare_text = self.ident().value.lower() == "text"
                self.skip_line()
                continue
            mods = []
            while self.tok.is_kw("public", "private", "global", "friend", "static"):
                mods.append(self.next().value.lower())
            if self.tok.is_kw("declare"):
                self.skip_line()
                continue
            if self.tok.is_kw("sub", "function", "property"):
                p = self.proc(mods)
                key = p.name if p.kind in ("sub", "function") else f"{p.kind}:{p.name}"
                m.procs[key] = p
                continue
            if self.tok.is_kw("type"):
                self.next()
                tname = self.ident().value.lower()
                fields = []
                self.skip_nl()
                while not self.tok.is_kw("end"):
                    fields.append(self.ident().value.lower())
                    self.skip_line()
                    self.skip_nl()
                self.next()
                self.expect_kw("type")
                m.types[tname] = fields
                continue
            if self.tok.is_kw("enum"):
                self.next()
                self.ident()
                self.skip_nl()
                val = -1
                while not self.tok.is_kw("end"):
                    name = self.ident().value.lower()
                    if self.tok.is_op("="):
                        self.next()
                        e = self.expr()
                        val = int(_const_value(e, m.enums))
                    else:
                        val += 1
                    m.enums[name] = val
                    self.skip_line()
                    self.skip_nl()
                self.next()
                self.expect_kw("enum")
                continue
            if self.tok.is_kw("implements") or (self.tok.is_kw("event") and self.peek().kind == "id"):
                self.skip_line()  # インターフェースの実装・イベントの宣言（動きには関わらない）
                continue
            if self.tok.is_kw("dim", "const") or mods:
                const = False
                if self.tok.is_kw("dim"):
                    self.next()
                elif self.tok.is_kw("const"):
                    self.next()
                    const = True
                m.decls.append(self.dim_list(const=const, line=tk.line))
                continue
            raise VbaSyntaxError(f"モジュールの中に、手続きの外の文「{tk.value}」があります", tk.line)
        return m

    def skip_line(self) -> None:
        while self.tok.kind not in ("nl", "eof"):
            self.next()

    def proc(self, mods: list[str]) -> Proc:
        line = self.tok.line
        kind = self.next().value.lower()
        if kind == "property":
            kind = self.ident().value.lower()  # get / let / set
        name_tok = self.ident()
        params: list[Param] = []
        if self.tok.is_op("("):
            self.next()
            while not self.tok.is_op(")"):
                params.append(self.param())
                if self.tok.is_op(","):
                    self.next()
            self.expect_op(")")
        if self.tok.is_kw("as"):
            self.next()
            self.type_name()
            if self.tok.is_op("(") and self.peek().is_op(")"):
                self.next()  # 「As String()」：配列を返す（5.3.1.1 の function-type）
                self.next()
        end_kw = {"sub": "sub", "function": "function"}.get(kind, "property")
        body = self.block({"end"})
        self.expect_kw("end")
        self.expect_kw(end_kw)
        return Proc(name_tok.value.lower(), name_tok.value, kind, params, body, private="private" in mods,
                    static="static" in mods, module=self.module, line=line)

    def param(self) -> Param:
        p = Param("")
        while self.tok.is_kw("optional", "byval", "byref", "paramarray"):
            w = self.next().value.lower()
            if w == "optional":
                p.optional = True
            elif w == "byval":
                p.byval = True
            elif w == "paramarray":
                p.paramarray = True
        p.name = self.ident().value.lower().rstrip("$%&!#@")
        if self.tok.is_op("("):
            self.next()
            self.expect_op(")")
            p.is_array = True
        if self.tok.is_kw("as"):
            self.next()
            self.type_name()
        if self.tok.is_op("="):
            self.next()
            p.default = self.expr()
        return p

    def type_name(self) -> str:
        if self.tok.is_kw("new"):
            self.next()
        name = self.ident().value
        while self.tok.is_op("."):
            self.next()
            name += "." + self.ident().value
        if self.tok.is_op("*"):  # String * 10
            self.next()
            self.next()
        return name.lower()

    def dim_list(self, const: bool = False, static: bool = False, line: int = 0) -> Dim:
        d = Dim(line=line, const=const, static=static)
        while True:
            if not const and self.tok.is_kw("withevents") and self.peek().kind == "id":
                self.next()  # 「Public WithEvents TextBox As MSForms.TextBox」：イベントを受け取る変数（5.2.3.1）
            tk = self.ident()
            v = DimVar(tk.value.lower().rstrip("$%&!#@"))
            if self.tok.is_op("("):
                self.next()
                v.bounds = []
                while not self.tok.is_op(")"):
                    lo = None
                    hi = self.expr()
                    if self.tok.is_kw("to"):
                        self.next()
                        lo, hi = hi, self.expr()
                    v.bounds.append((lo, hi))
                    if self.tok.is_op(","):
                        self.next()
                self.expect_op(")")
            if self.tok.is_kw("as"):
                self.next()
                if self.tok.is_kw("new"):
                    v.new = True
                v.type = self.type_name()
            if self.tok.is_op("="):
                self.next()
                v.init = self.expr()
            d.vars.append(v)
            if self.tok.is_op(","):
                self.next()
                continue
            break
        return d

    # ---------------------------------------------------------------- 文の並び
    def block(self, until: set[str]) -> list[Stmt]:
        """until のキーワード（Else・Next・Loop・Case・End If など）の手前までの文。
        「End」は、後ろに Sub・Function・If・Select・With などが続くときだけ、ブロックの終わり（ひとつだけの End は文）。"""
        body: list[Stmt] = []
        while True:
            self.skip_nl()
            tk = self.tok
            if tk.kind == "eof":
                raise VbaSyntaxError(f"「{'／'.join(sorted(until))}」がないまま終わりました", tk.line)
            if tk.kind == "id":
                low = tk.value.lower()
                if low in until and (low != "end" or self.peek().is_kw("sub", "function", "property", "if", "select",
                                                                         "with", "type", "enum")):
                    return body
            body.extend(self.statement())

    def statement(self) -> list[Stmt]:
        tk = self.tok
        line = tk.line
        # 行の名前（ラベル）・行番号
        if tk.kind == "num" and isinstance(tk.value, int) and (self.i == 0 or self.t[self.i - 1].kind == "nl"):
            self.next()
            return [Label(line=line, name=str(tk.value))]
        if tk.kind == "id" and self.peek().is_op(":") and not self.peek(2).is_op("=") and \
                (self.i == 0 or self.t[self.i - 1].kind == "nl") and tk.value.lower() not in _KEYWORDS:
            self.next()
            self.next()
            return [Label(line=line, name=tk.value.lower())]
        low = tk.value.lower() if tk.kind == "id" else ""
        if low == "attribute":
            self.skip_line()  # 手続きの設定（VB_ProcData など。エディタが付け、動きには関わらない）
            return []
        if low in ("dim", "static"):
            self.next()
            return [self.dim_list(static=low == "static", line=line)]
        if low == "const":
            self.next()
            return [self.dim_list(const=True, line=line)]
        if low in ("private", "public") and self.peek().is_kw("const", "dim"):
            self.next()
            return self.statement()
        if low == "redim":
            self.next()
            r = ReDim(line=line)
            if self.tok.is_kw("preserve"):
                self.next()
                r.preserve = True
            while True:
                # 名前（With の中の .名前 や、obj.名前 も）と、「(下限 To 上限, …)」
                target = self.primary()
                while self.tok.is_op(".") and self.peek().kind == "id":
                    self.next()
                    nm = self.ident()
                    target = Member(line=line, obj=target, name=nm.value.lower(), raw=nm.value)
                self.expect_op("(")
                bounds = []
                while not self.tok.is_op(")"):
                    lo = None
                    hi = self.expr()
                    if self.tok.is_kw("to"):
                        self.next()
                        lo, hi = hi, self.expr()
                    bounds.append((lo, hi))
                    if self.tok.is_op(","):
                        self.next()
                self.expect_op(")")
                if self.tok.is_kw("as"):
                    self.next()
                    self.type_name()
                r.vars.append((target, bounds))
                if self.tok.is_op(","):
                    self.next()
                    continue
                break
            return [r]
        if low == "set":
            self.next()
            target = self.postfix(self.primary(), allow_call_args=False)
            self.expect_op("=")
            return [Assign(line=line, target=target, value=self.expr(), set=True)]
        if low == "let":
            self.next()
            return self.statement()
        if low == "call":
            self.next()
            e = self.postfix(self.primary(), allow_call_args=False)
            return [CallStmt(line=line, expr=e)]
        if low == "if":
            return [self.if_()]
        if low == "for":
            return [self.for_()]
        if low == "do":
            return [self.do_()]
        if low == "while":
            self.next()
            cond = self.expr()
            body = self.block({"wend"})
            self.next()
            return [Do(line=line, pre=("while", cond), body=body)]
        if low == "select":
            return [self.select_()]
        if low == "with":
            self.next()
            obj = self.expr()
            body = self.block({"end"})
            self.expect_kw("end")
            self.expect_kw("with")
            return [With(line=line, obj=obj, body=body)]
        if low == "exit":
            self.next()
            return [Exit(line=line, what=self.ident().value.lower())]
        if low == "goto":
            self.next()
            return [GoTo(line=line, label=str(self.next().value).lower())]
        if low == "on" and self.peek().is_kw("error"):
            self.next()
            self.next()
            if self.tok.is_kw("resume"):
                self.next()
                self.expect_kw("next")
                return [OnError(line=line, mode="resume_next")]
            self.expect_kw("goto")
            lab = self.next()
            if lab.kind == "num" and lab.value == 0:
                return [OnError(line=line, mode="goto0")]
            if lab.is_op("-"):
                self.next()
                return [OnError(line=line, mode="goto0")]
            return [OnError(line=line, mode="goto", label=str(lab.value).lower())]
        if low == "resume":
            self.next()
            if self.at_end_of_stmt():
                return [Resume(line=line)]
            if self.tok.is_kw("next"):
                self.next()
                return [Resume(line=line, mode="next")]
            return [Resume(line=line, mode="label", label=str(self.next().value).lower())]
        if low in ("end",) and (self.peek().kind == "nl" or self.peek().is_op(":") or self.peek().kind == "eof"):
            self.next()
            return [End(line=line)]
        if low == "stop":
            self.next()
            return []
        if low == "erase":
            self.next()
            names = [self.expr()]
            while self.tok.is_op(","):
                self.next()
                names.append(self.expr())
            return [Erase(line=line, names=names)]
        if low == "open" and not self.peek().is_op("(", "."):
            return [self.open_()]
        if low == "close" and not self.peek().is_op("(", "."):
            self.next()
            args = []
            while not self.at_end_of_stmt():
                if self.tok.is_op("#", ","):
                    self.next()
                    continue
                args.append(self.expr())
            return [FileStmt(line=line, kind="close", args=args)]
        if low == "debug" and self.peek().is_op(".") and self.peek(2).is_kw("print"):
            # Debug.Print：イミディエイト ウィンドウに出す。式は Print # と同じ並び（「;」「,」・Spc・Tab）
            self.next()
            self.next()
            self.next()
            return [FileStmt(line=line, kind="debugprint", args=[None, self.output_list()])]
        if low in ("print", "write") and self.peek().is_op("#"):
            return [self.print_(low)]
        if low == "input" and self.peek().is_op("#"):
            self.next()
            self.next()
            num = self.expr()
            self.expect_op(",")
            targets = [self.postfix(self.primary(), allow_call_args=False)]
            while self.tok.is_op(","):
                self.next()
                targets.append(self.postfix(self.primary(), allow_call_args=False))
            return [FileStmt(line=line, kind="input", args=[num, targets])]
        if low == "line" and self.peek().is_kw("input"):
            self.next()
            self.next()
            self.expect_op("#")
            num = self.expr()
            self.expect_op(",")
            target = self.postfix(self.primary(), allow_call_args=False)
            return [FileStmt(line=line, kind="lineinput", args=[num, target])]
        if low in ("randomize", "doevents"):
            self.next()
            self.skip_to_stmt_end()
            return []
        # 代入か、呼び出し
        target = self.postfix(self.primary(), allow_call_args=False)
        if self.tok.is_op("="):
            self.next()
            return [Assign(line=line, target=target, value=self.expr())]
        # 呼び出し：「Foo a, b」「obj.Method a, b」「Foo(a)」
        if not self.at_end_of_stmt():
            args = self.arg_list(until_end=True)
            if isinstance(target, Call) and len(target.args) == 1 and not target.args[0].name and args:
                # 「Foo (a), b」：最初の引数がかっこで包まれていた
                first = target.args[0]
                first.byval_paren = True
                return [CallStmt(line=line, expr=Call(line=line, target=target.target, args=[first] + args))]
            return [CallStmt(line=line, expr=Call(line=line, target=target, args=args))]
        return [CallStmt(line=line, expr=target)]

    def skip_to_stmt_end(self) -> None:
        while not self.at_end_of_stmt():
            self.next()

    def arg_list(self, until_end: bool = False) -> list[Arg]:
        args: list[Arg] = []
        while True:
            if self.tok.is_op(",") or (until_end and self.at_end_of_stmt()) or (not until_end and self.tok.is_op(")")):
                args.append(Arg(None))
            else:
                name = None
                if self.tok.kind == "id" and self.peek().is_op(":="):
                    name = self.next().value.lower()
                    self.next()
                args.append(Arg(self.expr(), name))
            if self.tok.is_op(","):
                self.next()
                continue
            break
        if args and args[-1].value is None and args[-1].name is None and len(args) == 1:
            return []
        return args

    def if_(self) -> If:
        line = self.tok.line
        self.next()
        cond = self.expr()
        self.expect_kw("then")
        if self.tok.kind not in ("nl", "eof"):
            # 1行の If（Then の後ろに文が続く）
            node = If(line=line)
            body = self.inline_stmts()
            node.branches.append((cond, body))
            if self.tok.is_kw("else"):
                self.next()
                node.else_body = self.inline_stmts()
            return node
        node = If(line=line)
        body = self.block({"elseif", "else", "end"})
        node.branches.append((cond, body))
        while True:
            if self.tok.is_kw("elseif"):
                self.next()
                c = self.expr()
                self.expect_kw("then")
                node.branches.append((c, self.block({"elseif", "else", "end"})))
                continue
            if self.tok.is_kw("else"):
                self.next()
                node.else_body = self.block({"end"})
                continue
            self.expect_kw("end")
            self.expect_kw("if")
            return node

    def inline_stmts(self) -> list[Stmt]:
        out: list[Stmt] = []
        while self.tok.kind not in ("nl", "eof") and not self.tok.is_kw("else"):
            if self.tok.is_op(":"):
                self.next()
                continue
            out.extend(self.statement())
        return out

    def for_(self) -> Stmt:
        line = self.tok.line
        self.next()
        if self.tok.is_kw("each"):
            self.next()
            var = self.primary()
            self.expect_kw("in")
            coll = self.expr()
            body = self.block({"next"})
            self.next()
            if self.tok.kind == "id" and not self.at_end_of_stmt():
                self.next()
            return ForEach(line=line, var=var, coll=coll, body=body)
        var = self.postfix(self.primary(), allow_call_args=False)
        self.expect_op("=")
        start = self.expr()
        self.expect_kw("to")
        end = self.expr()
        step = None
        if self.tok.is_kw("step"):
            self.next()
            step = self.expr()
        body = self.block({"next"})
        self.next()
        # 「Next i」「Next i, j」（2つ目以降は、外側の For を閉じる）
        extra = 0
        if self.tok.kind == "id":
            self.next()
            while self.tok.is_op(","):
                self.next()
                self.ident()
                extra += 1
        node = For(line=line, var=var, start=start, end=end, step=step, body=body)
        node.closes_outer = extra  # type: ignore[attr-defined]
        return node

    def do_(self) -> Do:
        line = self.tok.line
        self.next()
        d = Do(line=line)
        if self.tok.is_kw("while", "until"):
            w = self.next().value.lower()
            d.pre = (w, self.expr())
        d.body = self.block({"loop"})
        self.next()
        if self.tok.is_kw("while", "until"):
            w = self.next().value.lower()
            d.post = (w, self.expr())
        return d

    def select_(self) -> Select:
        line = self.tok.line
        self.next()
        self.expect_kw("case")
        s = Select(line=line, expr=self.expr())
        self.skip_nl()
        while not self.tok.is_kw("end"):
            self.expect_kw("case")
            conds: list[tuple] | None = []
            if self.tok.is_kw("else"):
                self.next()
                conds = None
            else:
                while True:
                    if self.tok.is_kw("is"):
                        self.next()
                        op = self.next().value
                        conds.append(("is", op, self.expr()))
                    else:
                        a = self.expr()
                        if self.tok.is_kw("to"):
                            self.next()
                            conds.append(("to", a, self.expr()))
                        else:
                            conds.append(("eq", a))
                    if self.tok.is_op(","):
                        self.next()
                        continue
                    break
            body = self.block({"case", "end"})
            s.cases.append((conds, body))
        self.expect_kw("end")
        self.expect_kw("select")
        return s

    def open_(self) -> FileStmt:
        line = self.tok.line
        self.next()
        path = self.expr()
        mode = "random"
        if self.tok.is_kw("for"):
            self.next()
            mode = self.ident().value.lower()
        while self.tok.is_kw("access", "read", "write", "shared", "lock"):
            self.next()
        self.expect_kw("as")
        if self.tok.is_op("#"):
            self.next()
        num = self.expr()
        if self.tok.is_kw("len"):
            self.next()
            self.expect_op("=")
            self.expr()
        return FileStmt(line=line, kind="open", args=[path, mode, num])

    def print_(self, kind: str) -> FileStmt:
        line = self.tok.line
        self.next()
        self.expect_op("#")
        num = self.expr()
        items: list = []
        if self.tok.is_op(","):
            self.next()
            items = self.output_list()
        return FileStmt(line=line, kind=kind, args=[num, items])

    def output_list(self) -> list:
        """Print #・Write #・Debug.Print の出力の並び（MS-VBAL 5.4.5.8 output-list）。
        要素は式か、区切りの「;」「,」か、("spc", 式)・("tab", 式または None)。"""
        items: list = []
        while not self.at_end_of_stmt():
            if self.tok.is_op(";", ","):
                items.append(self.next().value)
                continue
            if self.tok.is_kw("spc", "tab") and (self.peek().is_op("(") or self.tok.is_kw("tab")):
                w = self.next().value.lower()
                arg = None
                if self.tok.is_op("("):
                    self.next()
                    arg = self.expr()
                    self.expect_op(")")
                elif w == "spc":
                    raise VbaSyntaxError("「Spc」の後ろに「(数)」がありません", self.tok.line)
                items.append((w, arg))
                continue
            items.append(self.expr())
        return items

    # ---------------------------------------------------------------- 式
    def expr(self) -> Node:
        return self.binary(0)

    _LEVELS = [("imp",), ("eqv",), ("xor",), ("or", "orelse"), ("and", "andalso"), ("not",),
               ("=", "<>", "<", ">", "<=", ">=", "like", "is"), ("&",), ("+", "-"), ("mod",), ("\\",), ("*", "/"),
               ("neg",), ("^",)]

    def binary(self, level: int) -> Node:
        if level >= len(self._LEVELS):
            return self.postfix(self.primary())
        ops = self._LEVELS[level]
        if ops == ("not",):
            if self.tok.is_kw("not"):
                line = self.next().line
                return Unary(line=line, op="not", operand=self.binary(level))
            return self.binary(level + 1)
        if ops == ("neg",):
            if self.tok.is_op("-", "+"):
                tk = self.next()
                return Unary(line=tk.line, op=tk.value, operand=self.binary(level))
            return self.binary(level + 1)
        left = self.binary(level + 1)
        while True:
            tk = self.tok
            op = None
            if tk.kind == "op" and tk.value in ops:
                op = tk.value
            elif tk.kind == "id" and tk.value.lower() in ops:
                op = tk.value.lower()
            if op is None:
                return left
            self.next()
            if op == "^":
                right = self.binary(level + 1) if not self.tok.is_op("-") else self.binary(level - 1)
            else:
                right = self.binary(level + 1)
            left = Binary(line=tk.line, op=op, left=left, right=right)

    def primary(self) -> Node:
        tk = self.tok
        if tk.kind == "num":
            self.next()
            return Lit(line=tk.line, value=tk.value)
        if tk.kind == "str":
            self.next()
            return Lit(line=tk.line, value=tk.value)
        if tk.kind == "date":
            self.next()
            return DateLit(line=tk.line, text=tk.value)
        if tk.is_op("("):
            self.next()
            e = self.expr()
            self.expect_op(")")
            return e
        if tk.is_op(".", "!"):
            # With の対象のメンバー
            self.next()
            name = self.ident()
            return Member(line=tk.line, obj=None, name=name.value.lower(), raw=name.value, bang=tk.value == "!")
        if tk.kind == "id":
            low = tk.value.lower()
            if low == "new":
                self.next()
                return New(line=tk.line, cls=self.type_name())
            if low in ("true", "false"):
                self.next()
                return Lit(line=tk.line, value=low == "true")
            if low == "nothing":
                self.next()
                return Lit(line=tk.line, value=NOTHING)
            if low == "empty":
                self.next()
                return Lit(line=tk.line, value=EMPTY)
            if low == "null":
                self.next()
                return Lit(line=tk.line, value=NULL)
            if low in ("typeof",):
                self.next()
                obj = self.postfix(self.primary())
                self.expect_kw("is")
                cls = self.type_name()
                return Call(line=tk.line, target=Name(line=tk.line, name="__typeof__", raw="TypeOf"),
                            args=[Arg(obj), Arg(Lit(line=tk.line, value=cls))])
            self.next()
            return Name(line=tk.line, name=low.rstrip("$%&!#@"), raw=tk.value)
        raise VbaSyntaxError(f"式があるはずの所に「{tk.value}」", tk.line)

    def postfix(self, node: Node, allow_call_args: bool = True) -> Node:
        while True:
            tk = self.tok
            if tk.is_op("("):
                self.next()
                args = self.arg_list() if not self.tok.is_op(")") else []
                self.expect_op(")")
                node = Call(line=tk.line, target=node, args=args)
                continue
            if tk.is_op(".") and self.peek().kind == "id":
                self.next()
                name = self.ident()
                node = Member(line=tk.line, obj=node, name=name.value.lower(), raw=name.value)
                continue
            if tk.is_op("!") and self.peek().kind == "id":
                self.next()
                name = self.ident()
                node = Member(line=tk.line, obj=node, name=name.value.lower(), raw=name.value, bang=True)
                continue
            return node


class _Sentinel:
    def __init__(self, name: str):
        self.name = name

    def __repr__(self) -> str:
        return self.name

    def __bool__(self) -> bool:
        return False


EMPTY = _Sentinel("Empty")
NULL = _Sentinel("Null")
NOTHING = _Sentinel("Nothing")
MISSING = _Sentinel("Missing")

_KEYWORDS = {"if", "then", "else", "elseif", "end", "for", "next", "do", "loop", "while", "wend", "until", "select", "case",
             "with", "dim", "set", "let", "call", "exit", "goto", "on", "resume", "sub", "function", "property", "to",
             "step", "each", "in", "and", "or", "not", "xor", "mod", "is", "like", "new", "as", "byval", "byref",
             "optional", "const", "redim", "preserve", "true", "false", "nothing", "empty", "null"}


def _const_value(e: Node, enums: dict) -> int:
    if isinstance(e, Lit):
        return int(e.value)
    if isinstance(e, Unary) and e.op == "-":
        return -_const_value(e.operand, enums)
    if isinstance(e, Name) and e.name in enums:
        return enums[e.name]
    if isinstance(e, Binary):
        a, b = _const_value(e.left, enums), _const_value(e.right, enums)
        return {"+": a + b, "-": a - b, "*": a * b, "or": a | b, "and": a & b}.get(e.op, 0)
    return 0


# ====================================================================== 条件付きコンパイル（MS-VBAL 3.4）
# 実装が決める定数（3.4.1）。Windows の 64 ビット版 Office（いまの Microsoft 365 の標準の入れ方）と同じにする。
# Win32 は 64 ビット版でも真（32 ビット以上の Windows で真）
PREDEFINED = {"vba6": True, "vba7": True, "win16": False, "win32": True, "win64": True, "mac": False}

_DIRECTIVE = re.compile(r"[ \t]*#[ \t]*(const|if|elseif|else|end[ \t]*if)(?![\w])(.*)", re.I | re.S)


def _continues(line: str) -> bool:
    """行末の「 _」で、次の行に続く（コメントの行も続く）。"""
    t = line.rstrip(" \t")
    return t.endswith("_") and len(t) >= 2 and t[-2] in " \t"


def project_constants(text: str) -> dict:
    """プロジェクトの条件付きコンパイルの引数（MS-OVBA 2.3.4.2.1.12 PROJECTCONSTANTS、「名前 = 数 : 名前 = 数」）。"""
    out = {}
    for part in (text or "").split(":"):
        k, eq, v = part.partition("=")
        k, v = k.strip(), v.strip()
        if not eq or not re.fullmatch(r"[^\W\d]\w*", k):
            continue
        try:
            out[k.lower()] = int(v)
        except ValueError:
            continue
    return out


def _cc_value(e: Node, env: dict):
    """条件付きコンパイルの式（cc-expression）の値。決めていない名前は Empty（0 と同じ）。"""
    if isinstance(e, Lit):
        v = e.value
        return 0 if isinstance(v, _Sentinel) else v
    if isinstance(e, Name):
        return env.get(e.name, 0)
    if isinstance(e, Unary):
        v = _cc_value(e.operand, env)
        if e.op == "not":
            return (not v) if isinstance(v, bool) else ~_cc_int(v)
        return -_cc_num(v) if e.op == "-" else _cc_num(v)
    if isinstance(e, Binary):
        a, b = _cc_value(e.left, env), _cc_value(e.right, env)
        op = e.op
        if op in ("and", "or", "xor", "eqv", "imp", "andalso", "orelse"):
            if isinstance(a, bool) and isinstance(b, bool):
                return {"and": a and b, "andalso": a and b, "or": a or b, "orelse": a or b, "xor": a != b,
                        "eqv": a == b, "imp": (not a) or b}[op]
            x, y = _cc_int(a), _cc_int(b)
            return {"and": x & y, "andalso": x & y, "or": x | y, "orelse": x | y, "xor": x ^ y,
                    "eqv": ~(x ^ y), "imp": (~x) | y}[op]
        if op == "&":
            return _cc_text(a) + _cc_text(b)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            if isinstance(a, str) and isinstance(b, str):
                x, y = a, b
            else:
                x, y = _cc_num(a), _cc_num(b)
            return {"=": x == y, "<>": x != y, "<": x < y, ">": x > y, "<=": x <= y, ">=": x >= y}[op]
        x, y = _cc_num(a), _cc_num(b)
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "^":
            return x ** y
        if op in ("/", "\\", "mod"):
            if y == 0:
                raise VbaSyntaxError("条件付きコンパイルの式で、0 で割っています", e.line)
            if op == "/":
                return x / y
            q = abs(_cc_int(x)) // abs(_cc_int(y))
            if op == "\\":
                return q if (_cc_int(x) >= 0) == (_cc_int(y) >= 0) else -q
            r = abs(_cc_int(x)) % abs(_cc_int(y))
            return r if _cc_int(x) >= 0 else -r
    raise VbaSyntaxError("条件付きコンパイルの式に、定数でないものがあります", e.line)


def _cc_num(v):
    if isinstance(v, bool):
        return -1 if v else 0
    if isinstance(v, (int, float)):
        return v
    try:
        return float(v)
    except (TypeError, ValueError):
        raise VbaSyntaxError(f"条件付きコンパイルの式で、数にできない値「{v}」") from None


def _cc_int(v) -> int:
    x = _cc_num(v)
    return int(round(x)) if isinstance(x, float) else x


def _cc_text(v) -> str:
    if isinstance(v, bool):
        return "True" if v else "False"
    return str(v)


def _cc_truth(v) -> bool:
    return v if isinstance(v, bool) else _cc_num(v) != 0


def preprocess(source: str, constants: dict | None = None) -> str:
    """条件付きコンパイル（#Const・#If・#ElseIf・#Else・#End If）を解き、使わない行を空の行にしたソース。

    行の番号を変えないため、指示の行と、使わない側の行は、空の行にする。定数は、実装が決めるもの（PREDEFINED）、
    プロジェクトの引数（constants）、モジュールの #Const の順に上書きする。"""
    src = source.replace("\r\n", "\n").replace("\r", "\n")
    if "#" not in src:
        return src
    lines = src.split("\n")
    env = dict(PREDEFINED)
    env.update({k.lower(): v for k, v in (constants or {}).items()})
    out: list[str] = []
    stack: list[list] = []  # [外側が使われているか, どれかの枝を使ったか, #If の行]
    active = True
    i = 0
    while i < len(lines):
        j = i
        while _continues(lines[j]) and j + 1 < len(lines):
            j += 1
        phys = lines[i:j + 1]
        line_no = i + 1
        logical = " ".join(x.rstrip(" \t")[:-1] if k < len(phys) - 1 else x for k, x in enumerate(phys))
        m = _DIRECTIVE.match(logical)
        i = j + 1
        if not m:
            out.extend(phys if active else [""] * len(phys))
            continue
        out.extend([""] * len(phys))
        word = re.sub(r"[ \t]+", "", m.group(1)).lower()
        toks = tokenize(m.group(2))
        for t in toks:
            t.line = line_no
        ps = Parser(toks, "")

        def cond() -> bool:
            e = ps.expr()
            ps.expect_kw("then")
            if ps.tok.kind not in ("nl", "eof"):
                raise VbaSyntaxError(f"「#{m.group(1)} … Then」の後ろに「{ps.tok.value}」", line_no)
            return _cc_truth(_cc_value(e, env))

        if word == "const":
            name = ps.ident().value.lower().rstrip("$%&!#@")
            ps.expect_op("=")
            v = _cc_value(ps.expr(), env)
            if active:
                env[name] = v
        elif word == "if":
            took = active and cond()
            stack.append([active, took, line_no])
            active = took
        elif word == "elseif":
            if not stack:
                raise VbaSyntaxError("「#If」のない「#ElseIf」があります", line_no)
            outer, taken, _ = stack[-1]
            ok = outer and not taken and cond()
            stack[-1][1] = taken or ok
            active = ok
        elif word == "else":
            if not stack:
                raise VbaSyntaxError("「#If」のない「#Else」があります", line_no)
            outer, taken, _ = stack[-1]
            active = outer and not taken
            stack[-1][1] = True
        else:  # end if
            if not stack:
                raise VbaSyntaxError("「#If」のない「#End If」があります", line_no)
            active = stack.pop()[0]
    if stack:
        raise VbaSyntaxError("「#If」が「#End If」で閉じていません", stack[-1][2])
    return "\n".join(out)


def parse_module(name: str, source: str, constants: dict | None = None) -> Module:
    """モジュールのソースを構文木にする。constants はプロジェクトの条件付きコンパイルの引数。"""
    return Parser(tokenize(preprocess(source, constants)), name).module_()
