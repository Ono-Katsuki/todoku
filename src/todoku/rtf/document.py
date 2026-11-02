"""RTF を読む（Rich Text Format Specification 1.9.1）。外部のライブラリは使わない。

- 文字：\\'hh はフォントの文字セット（\\fcharset）の文字コード（なければ \\ansicpg）で読み、2バイトの字（Shift_JIS など）は
  続く \\'hh と合わせて1字にする。\\uN は Unicode の字で、そのあとの代わりの字（\\ucN の数）は読まない
- 段落（\\par）・表（\\trowd・\\cellx・\\cell・\\row、入れ子の \\itap・\\nestcell・\\nestrow、縦の結合 \\clvmgf・\\clvmrg、
  横の結合 \\clmgf・\\clmrg）
- フィールド（\\field・\\fldinst・\\fldrslt）と、旧来のフォームの欄（FORMTEXT・FORMCHECKBOX・FORMDROPDOWN、\\formfield）
- 図形の字（\\shp の \\shptxt、古い \\do の \\dptxbxtext）・ヘッダー/フッター・脚注・コメント（\\annotation）・しおり（\\bkmkstart）・
  段落番号（\\listtext・\\pntext）・隠し文字（\\v）・下線（\\ul…）
読んだ結果は、Word（.docx）と同じ形（段落・表・マス・入力欄）にして、窓・入力欄の推定・字を目印にした書き込みを共通にする。
段落の字は、1字ごとにファイルの中のバイトの範囲を覚える（書き込みを小さな差分にするため）。
"""

from __future__ import annotations

import codecs
import re
from dataclasses import dataclass, field

from ..docx.document import WINGDINGS, Cell, Control, Para, Seg, Table, iter_paras
from .lexer import Tok, tokens

MARK_PREFIX = "_todoku_"

CHARSET_CP = {0: "cp1252", 77: "mac_roman", 128: "cp932", 129: "cp949", 130: "johab", 134: "gbk", 136: "big5",
              161: "cp1253", 162: "cp1254", 163: "cp1258", 177: "cp1255", 178: "cp1256", 186: "cp1257", 204: "cp1251",
              222: "cp874", 238: "cp1250", 255: "cp437"}
SPECIAL = {"emdash": "—", "endash": "–", "emspace": "\u2003", "enspace": "\u2002", "qmspace": "\u2005", "bullet": "•",
           "lquote": "‘", "rquote": "’", "ldblquote": "“", "rdblquote": "”"}
INVISIBLE = {"zwj", "zwnj", "ltrmark", "rtlmark", "zwbo", "zwnbo", "softline", "softcol", "softpage"}
# 中の字を出さない宛先（書式・情報・絵・フィールドの命令など）
SKIP = {"fonttbl", "colortbl", "stylesheet", "info", "pict", "listtable", "listoverridetable", "revtbl", "rsidtbl",
        "generator", "xmlnstbl", "latentstyles", "datastore", "themedata", "colorschememapping", "objdata", "nonshppict",
        "shprslt", "pnseclvl", "pn", "ftnsep", "ftnsepc", "aftnsep", "aftnsepc", "ftncn", "aftncn", "pgdsctbl", "mmathPr",
        "atnid", "atnauthor", "atntime", "atnicn", "atnref", "atndate", "atrfstart", "atrfend", "template", "userprops",
        "docvar", "password", "passwordhash", "xform", "defchp", "defpap", "listpicture", "factoidname", "mhtmltag",
        "htmltag", "panose", "falt", "fname", "objclass", "objname", "objtime", "oleclsid", "bliptag", "blipuid", "picprop",
        "wgrffmtfilter", "tc", "tcn", "xe", "pntxta", "pntxtb", "background", "sp", "sn", "sv",
        "fchars", "lchars", "ffname", "ffdeftext", "ffl", "ffhelptext", "ffstattext", "ffentrymcr", "ffexitmcr",
        "keycode", "bkmkpub", "category", "comment", "doccomm", "hlinkbase", "operator", "subject", "title", "author",
        "manager", "company", "keywords", "revtim", "creatim", "printim", "buptim", "propname", "staticval", "leveltext",
        "levelnumbers", "listname", "liststylename", "objalias", "objsect", "datafield", "fontemb", "fontfile", "file",
        "filetbl", "protusertbl", "pxe", "rxe", "txe", "ud", "upr", "mmath", "oleclsid", "atnparent", "atndate"}
STORY = {"header": "header", "headerl": "header", "headerr": "header", "headerf": "header", "footer": "footer",
         "footerl": "footer", "footerr": "footer", "footerf": "footer"}


class RtfError(ValueError):
    pass


@dataclass
class RSrc:
    """段落の字の1片が、RTF のどこから来たか（Seg.node に入れる）。"""

    kind: str  # txt（字）/ tab / br / el（字の印・図の説明）/ field（フィールドの表示）/ ctl（フォームの欄の印）
    s: int
    e: int
    starts: list[int] | None = None
    ends: list[int] | None = None
    uc: int = 1  # この字を読んだときの \\uc（書き足す字の、代わりの字の数）


@dataclass
class _St:
    """グループの状態（「{」で複写し、「}」で戻す）。"""

    dest: str | None = None  # 宛先
    skip: bool = False
    uc: int = 1
    font: int | None = None
    hidden: bool = False
    ul: bool = False
    intbl: bool = False
    itap: int = 0
    first: bool = True  # グループの最初の字句か（宛先の制御語を見分ける）
    star: bool = False  # \* のあと


@dataclass
class _Field:
    instr: list = field(default_factory=list)
    ctl: Control | None = None
    kind: str = ""
    ff: dict = field(default_factory=dict)
    n0: int = 0
    para: Para | None = None


class _Cell:
    def __init__(self):
        self.blocks: list = []


class _Level:
    """表の入れ子の1段（いま作っている表・行・マス）。"""

    def __init__(self):
        self.table: Table | None = None
        self.rows: list[list] = []  # [(マスの中身 _Cell, ...)]
        self.row: list[_Cell] = []
        self.cell = _Cell()
        self.defs: list[list[dict]] = []  # 行ごとのマスの定義（\cellx・結合）
        self.cur_def: list[dict] = []


class _Story:
    def __init__(self, kind: str, start: int):
        self.kind = kind
        self.blocks: list = []
        self.levels: list[_Level] = []
        self.para: Para | None = None
        self.last_end = start


def nearest(groups: dict, depth: int):
    """いまの深さで集めている宛先（しおりの名前・段落番号など）：いまの深さ以下の、いちばん深いもの。"""
    keys = [k for k in groups if k <= depth]
    return groups[max(keys)] if keys else None


class RtfDocument:
    kind = "rtf"

    def __init__(self, blob: bytes, filename: str = ""):
        self.blob = blob
        self.filename = filename
        if not blob.lstrip(b"\xef\xbb\xbf \t\r\n").startswith(b"{\\rtf"):
            raise RtfError("RTF ではありません（{\\rtf で始まっていません）")
        self.controls: list[Control] = []
        self.marks_raw: list[tuple[str, str, Para, int]] = []
        self.protection = ""
        self.protection_enforced = False
        self.has_macros = False
        self.vba = None
        self.ansicpg = "cp1252"
        self.fonts: dict[int, tuple[int | None, str]] = {}  # フォント → (文字セット, 名前)
        self.deff: int | None = None
        self.stories: list[tuple[str, str, list]] = []
        self.unread: list[str] = []
        self._fonts_pass()
        head = blob[:200_000]
        if re.search(rb"\\formprot(?![a-z])", head):
            self.protection, self.protection_enforced = "forms", True
        elif re.search(rb"\\readprot(?![a-z])", head):
            self.protection, self.protection_enforced = "readOnly", True
        self._parse()

    # ------------------------------------------------------------------ フォントの表
    def _fonts_pass(self) -> None:
        b = self.blob
        m = re.search(rb"\\ansicpg(\d+)", b[:4000])
        if m:
            self.ansicpg = f"cp{int(m[1])}"
            try:
                codecs.lookup(self.ansicpg)
            except LookupError:
                self.ansicpg = "cp1252"
        m = re.search(rb"\\deff(\d+)", b[:4000])
        if m:
            self.deff = int(m[1])
        i = b.find(b"{\\fonttbl")
        if i < 0:
            return
        depth = 0
        j = i
        cur = None
        name = bytearray()
        charset = None
        for t in tokens(b[i:]):
            if t.kind == "open":
                depth += 1
            elif t.kind == "close":
                depth -= 1
                if depth == 1 and cur is not None:
                    self.fonts[cur] = (charset, self._decode_name(bytes(name), charset))
                    cur, name, charset = None, bytearray(), None
                if depth == 0:
                    break
            elif t.kind == "word":
                if t.word == "f" and t.param is not None and depth >= 1 and cur is None:
                    cur = t.param
                elif t.word == "fcharset" and t.param is not None:
                    charset = t.param
            elif t.kind == "hex" and depth == 2:
                name += t.data
            elif t.kind == "text" and depth == 2:
                name += t.data
        del j

    def _decode_name(self, raw: bytes, charset: int | None) -> str:
        cp = CHARSET_CP.get(charset or 0, self.ansicpg) if charset not in (None, 1) else self.ansicpg
        try:
            s = raw.decode(cp)
        except (UnicodeDecodeError, LookupError):
            s = raw.decode("latin-1")
        return s.split(";")[0].strip()

    def codepage(self, font: int | None) -> str:
        f = self.fonts.get(font if font is not None else self.deff if self.deff is not None else -1)
        if f is None:
            return self.ansicpg
        cs, name = f
        if cs == 2:
            return "symbol:" + name.lower()
        if cs is None or cs == 1:
            return self.ansicpg
        return CHARSET_CP.get(cs, self.ansicpg)

    # ------------------------------------------------------------------ 読む
    def _parse(self) -> None:
        b = self.blob
        stack: list[_St] = []
        st = _St()
        story = _Story("body", 0)
        self.body_story = story
        stories: list[_Story] = [story]
        story_stack: list[tuple[int, _Story]] = []  # (グループの深さ, 外の流れ)
        fields: list[tuple[int, _Field]] = []  # (グループの深さ, フィールド)
        pending: list[tuple[int, int, int, str]] = []  # \'hh：(バイト, 始め, 終わり, 文字コード)
        uskip = 0  # \uN のあとに読まない字の残り
        last_u: list | None = None  # 直前の \uN の字（読まない字の範囲を足す）
        special_groups: dict[int, dict] = {}  # 深さ → 宛先の、集めている情報（しおり・段落番号・フォームの欄など）
        self._box_targets: list = []
        depth = 0

        def para() -> Para:
            if story.para is None:
                p = Para(None, "rtf", [])
                p.start = story.last_end
                p.mark = None
                p.protected = False
                story.para = p
            return story.para

        def flush_bytes() -> None:
            nonlocal pending
            if not pending:
                return
            run = pending
            pending = []
            cp = run[0][3]
            if cp.startswith("symbol:"):
                for v, s, e, _cp in run:
                    ch = WINGDINGS.get(0xF000 + v, chr(v) if v < 0x80 else "◆") if "wingding" in cp or "symbol" in cp else chr(v)
                    add_char(ch, s, e)
                return
            try:
                dec = codecs.getincrementaldecoder(cp)(errors="replace")
            except LookupError:
                dec = codecs.getincrementaldecoder("cp1252")(errors="replace")
            first = None
            for v, s, e, _cp in run:
                if first is None:
                    first = s
                out = dec.decode(bytes([v]))
                if out:
                    for ch in out:
                        add_char(ch, first, e)
                    first = None
            tail = dec.decode(b"", final=True)
            if tail and first is not None:
                add_char(tail, first, run[-1][2])

        def add_char(ch: str, s: int, e: int, kind: str = "t") -> None:
            if st.skip or st.dest in SKIP:
                return
            dest = st.dest
            if dest in ("listtext", "pntext"):
                info = nearest(special_groups, depth)
                if info is not None:
                    info["text"] += ch
                return
            if dest == "fldinst":
                if fields:
                    fields[-1][1].instr.append(ch)
                return
            if dest in ("bkmkstart", "bkmkend", "ffname", "ffl", "ffdeftext"):
                return
            if ch == "\xa0":
                ch = " "
            p = para()
            fctl = None
            for _d, fd in reversed(fields):
                if fd.kind == "result":
                    fctl = fd.ctl if fd.ctl is not None else fd.dummy
                    break
            segs = p.segs
            if kind == "t" and segs and segs[-1].kind == "t" and segs[-1].node is not None and segs[-1].node.kind == "txt" \
                    and segs[-1].hidden == st.hidden and segs[-1].underline == st.ul and segs[-1].field is fctl \
                    and getattr(segs[-1], "_open", False) and segs[-1].node.e == s_join[0]:
                sg = segs[-1]
                sg.text += ch
                sg.node.starts.append(s)
                sg.node.ends.append(e)
                sg.node.e = e
            else:
                if segs and getattr(segs[-1], "_open", False):
                    segs[-1]._open = False
                sg = Seg(ch, kind, RSrc("txt", s, e, [s], [e], st.uc), None, st.hidden, st.ul, fctl)
                sg._open = True
                segs.append(sg)
            s_join[0] = e

        s_join = [-1]  # 字をまとめる：直前の字の終わり（あいだに制御語があれば、まとめない）

        def end_para(tok: Tok, how: str) -> None:
            """段落の終わり（how：par・cell・row・nestcell・nestrow・sect・end）。"""
            nonlocal story
            flush_bytes()
            p = story.para
            if p is None and how in ("par", "sect", "cell", "nestcell"):
                p = para()
            if p is not None:
                for sg in p.segs:
                    sg._open = False
                p.mark = (tok.s, tok.e) if tok is not None else None
                p.end = tok.e if tok is not None else (p.segs[-1].node.e if p.segs and p.segs[-1].node else p.start)
                p.uc = st.uc
                story.last_end = p.end
                story.para = None
            level = st.itap if st.itap > 0 else (1 if st.intbl else 0)
            if how in ("cell", "nestcell"):
                level = max(level, 1 if how == "cell" else 2)
            self._place(story, p, level, how)

        for t in tokens(b):
            k = t.kind
            # \uN のあとの、代わりの字を読まない
            if uskip > 0 and k in ("text", "hex", "sym", "word") and not (k == "word" and t.word in ("par",)):
                if k == "text":
                    n = len(t.data)
                    if n <= uskip:
                        uskip -= n
                        if last_u is not None:
                            last_u[1] = t.e
                        continue
                    cut = t.s + uskip
                    if last_u is not None:
                        last_u[1] = cut
                    t = Tok("text", cut, t.e, data=t.data[uskip:])
                    uskip = 0
                else:
                    uskip -= 1
                    if last_u is not None:
                        last_u[1] = t.e
                    continue
            if last_u is not None and (uskip == 0 or k in ("open", "close")):
                ch, end_ = last_u[0], last_u[1]
                last_u = None
                add_char(ch, last_u_start[0], end_)
                uskip = 0
            if k == "open":
                flush_bytes()
                stack.append(st)
                st = _St(st.dest, st.skip, st.uc, st.font, st.hidden, st.ul, st.intbl, st.itap, True, False)
                depth += 1
                s_join[0] = -1
                continue
            if k == "close":
                flush_bytes()
                s_join[0] = -1
                if st.dest == "fldrslt" and (not stack or stack[-1].dest != "fldrslt") and fields:
                    fields[-1][1].result_close = t.s  # フィールドの結果のグループの終わり（空の欄に書く所）
                info = special_groups.pop(depth, None)
                if info is not None:
                    self._close_special(info, story, fields, t, depth)
                while fields and fields[-1][0] == depth:
                    _d, fd = fields.pop()
                    self._close_field(fd, story, t)
                if story_stack and story_stack[-1][0] == depth:
                    _d, outer = story_stack.pop()
                    end_para(None, "end")
                    self._close_levels(story)
                    inner = story
                    story = outer
                    self._attach(inner, story)
                if not stack:
                    break
                st = stack.pop()
                depth -= 1
                continue
            if k == "nl":
                continue
            if k == "bin":
                continue
            if k == "text":
                flush_bytes()
                st.first = False
                data = t.data
                if st.dest in ("bkmkstart", "bkmkend", "ffname", "ffl", "ffdeftext", "ffhelptext", "ffstattext"):
                    info = nearest(special_groups, depth)
                    if info is not None:
                        info.setdefault("raw", bytearray()).extend(data)
                        info["cp"] = self.codepage(st.font)
                    continue
                if any(x >= 0x80 for x in data):
                    cp = self.codepage(st.font)
                    for i, x in enumerate(data):
                        if x >= 0x80:
                            pending.append((x, t.s + i, t.s + i + 1, cp))
                        else:
                            flush_bytes()
                            add_char(chr(x), t.s + i, t.s + i + 1)
                    continue
                for i, x in enumerate(data):
                    add_char(chr(x), t.s + i, t.s + i + 1)
                continue
            if k == "hex":
                st.first = False
                if st.dest in ("bkmkstart", "bkmkend", "ffname", "ffl", "ffdeftext", "ffhelptext", "ffstattext"):
                    info = nearest(special_groups, depth)
                    if info is not None:
                        info.setdefault("raw", bytearray()).extend(t.data)
                        info["cp"] = self.codepage(st.font)
                    continue
                cp = self.codepage(st.font)
                if pending and pending[-1][3] != cp:
                    flush_bytes()
                pending.append((t.param, t.s, t.e, cp))
                continue
            flush_bytes()
            if k == "sym":
                w = t.word
                if w == "*":
                    st.star = True  # 「\\*」のあとの制御語は、知らなければ飛ばす宛先（グループの最初の字句のまま）
                    continue
                st.first = False
                if w in ("\\", "{", "}"):
                    add_char(w, t.s, t.e)
                elif w == "~":
                    add_char(" ", t.s, t.e)
                elif w == "_":
                    add_char("‑", t.s, t.e)
                elif w == "-":
                    pass  # 任意のハイフン（見えない）
                continue
            # 制御語
            w = t.word
            first = st.first
            st.first = False
            s_join[0] = -1
            if first and (st.star or w in SKIP or w in STORY or w in ("fldinst", "fldrslt", "field", "footnote", "annotation",
                                                                      "shptxt", "dptxbxtext", "shpinst", "shp", "do",
                                                                      "listtext", "pntext", "bkmkstart", "bkmkend",
                                                                      "formfield", "result", "object", "nesttableprops")):
                if st.skip and w not in ("ffname", "ffl", "ffdeftext", "ffhelptext", "ffstattext"):
                    st.dest = w if st.star else st.dest  # 飛ばしている宛先の中（\\shprslt の中の \\dptxbxtext など）は、そのまま飛ばす
                    continue
                known = self._start_dest(w, t, st, stack, story, fields, special_groups, depth)
                if known == "story":
                    story_stack.append((depth, story))
                    kind = STORY.get(w) or ("comments" if w == "annotation" else "footnotes" if w == "footnote" else "box")
                    new = _Story(kind, t.e)
                    new.anchor = story.para if kind == "box" else None
                    if kind == "box":
                        new.anchor = para()
                    if w == "footnote":
                        p = para()
                        p.segs.append(Seg("※", "note", RSrc("el", t.s, t.e), None, st.hidden))
                    stories.append(new)
                    story = new
                    st.dest = None
                    st.skip = False
                    st.intbl = False
                    st.itap = 0
                    continue
                if known == "skip" or (st.star and known is None):
                    st.dest = w
                    st.skip = True
                    continue
                if known == "keep":
                    continue
            if st.skip:
                if st.dest == "formfield" or (fields and st.dest in ("formfield",)):
                    pass
                if st.dest in ("formfield",):
                    self._ff_word(fields, t)
                continue
            if w == "u" and t.param is not None:
                code = t.param + 65536 if t.param < 0 else t.param
                ch = chr(code) if not 0xD800 <= code < 0xE000 else ""
                if 0xDC00 <= code < 0xE000 and story.para is not None and story.para.segs and \
                        0xD800 <= ord(story.para.segs[-1].text[-1:] or "a") < 0xDC00:
                    # サロゲートの組：前の字と合わせる
                    prev = story.para.segs[-1]
                    hi = ord(prev.text[-1])
                    prev.text = prev.text[:-1] + chr(0x10000 + ((hi - 0xD800) << 10) + (code - 0xDC00))
                    prev.node.ends[-1] = t.e
                    prev.node.e = t.e
                    uskip = st.uc
                    continue
                if 0xD800 <= code < 0xDC00:
                    ch = chr(code)
                last_u = [ch, t.e]
                last_u_start = [t.s]
                uskip = st.uc
                if uskip == 0:
                    add_char(ch, t.s, t.e)
                    last_u = None
                continue
            if w == "uc" and t.param is not None:
                st.uc = t.param
            elif w == "f" and t.param is not None:
                st.font = t.param
            elif w == "plain":
                st.font, st.hidden, st.ul = self.deff, False, False
            elif w == "v":
                st.hidden = t.param != 0
            elif w.startswith("ul") and w not in ("ulc",):
                st.ul = not (w == "ulnone" or t.param == 0)
            elif w == "pard":
                st.intbl, st.itap = False, 0
            elif w == "intbl":
                st.intbl = True
            elif w == "itap" and t.param is not None:
                st.itap = t.param
            elif w in ("par", "sect"):
                end_para(t, w)
            elif w in ("cell", "nestcell"):
                end_para(t, w)
            elif w in ("row", "nestrow"):
                self._end_row(story, 2 if w == "nestrow" else 1, t)
            elif w == "trowd":
                self._row_def(story, st, "trowd", None)
            elif w in ("cellx", "clvmgf", "clvmrg", "clmgf", "clmrg", "clcbpat"):
                self._row_def(story, st, w, t.param)
            elif w == "tab":
                para().segs.append(Seg("\t", "tab", RSrc("tab", t.s, t.e), None, st.hidden, st.ul, None))
            elif w in ("line", "page", "column"):
                para().segs.append(Seg("\n", "br", RSrc("br", t.s, t.e), None, st.hidden, False, None))
            elif w in SPECIAL:
                add_char(SPECIAL[w], t.s, t.e)
            elif w == "chftn":
                pass
            elif w == "protect":
                pass
        flush_bytes()
        if last_u is not None:
            add_char(last_u[0], last_u_start[0], last_u[1])
        # 終わりの段落（\par のない最後の段落）と、閉じていない表
        for s_ in stories:
            if s_.para is not None and s_.para.segs:
                p = s_.para
                p.end = p.segs[-1].node.e if p.segs[-1].node is not None else p.start
                self._place(s_, p, 0, "end")
                s_.para = None
            self._close_levels(s_)
        self.body = self.body_story.blocks
        for s_ in stories[1:]:
            if s_.kind in ("header", "footer", "footnotes", "comments") and s_.blocks:
                self.stories.append((s_.kind, "rtf", s_.blocks))
        # 隠し文字だけの段落も、そのまま（窓では〔非表示の文字〕）

    def _start_dest(self, w: str, t: Tok, st: _St, stack, story, fields, special: dict, depth: int) -> str | None:
        if w in STORY or w in ("footnote", "annotation", "shptxt", "dptxbxtext"):
            return "story"
        if w == "field":
            fd = _Field()
            fd.dummy = Control("field", "rtf-field", "", None, "rtf")
            fd.start = t.s
            fields.append((depth, fd))
            st.dest = "field"
            return "keep"
        if w == "fldinst":
            st.dest = "fldinst"
            if fields:
                fields[-1][1].kind = "instr"
            return "keep"
        if w == "fldrslt":
            st.dest = "fldrslt"
            if fields:
                fd = fields[-1][1]
                fd.kind = "result"
                self._open_result(fd, story, t)
            return "keep"
        if w in ("listtext", "pntext"):
            st.dest = w
            special[depth] = {"kind": "label", "text": ""}
            return "keep"
        if w in ("bkmkstart", "bkmkend"):
            st.dest = w
            special[depth] = {"kind": w, "text": ""}
            return "keep"
        if w == "formfield":
            st.dest = "formfield"
            st.skip = True
            if fields:
                fields[-1][1].ff["span"] = (t.s, None)
                fields[-1][1].ff["depth"] = depth
            return "keep"
        if w in ("ffname", "ffl", "ffdeftext", "ffhelptext", "ffstattext"):
            st.dest = w
            special[depth] = {"kind": w, "text": ""}
            return "keep"
        if w in ("shp", "shpinst", "do", "object", "nesttableprops"):
            st.dest = w
            return "keep"
        if w == "result":
            st.dest = None
            return "keep"
        if w in SKIP:
            return "skip"
        return None

    def _ff_word(self, fields, t: Tok) -> None:
        if not fields:
            return
        ff = fields[-1][1].ff
        if t.word in ("fftype", "ffres", "ffdefres", "ffmaxlen", "fftypetxt"):
            ff[t.word] = t.param if t.param is not None else 0
            ff[t.word + "_pos"] = (t.s, t.e)

    def _close_special(self, info: dict, story, fields, t: Tok, depth: int) -> None:
        text = info.get("text", "")
        if "raw" in info:
            try:
                text += bytes(info["raw"]).decode(info.get("cp", self.ansicpg), "replace")
            except LookupError:
                text += bytes(info["raw"]).decode("latin-1")
        kind = info["kind"]
        if kind == "label":
            p = story.para
            if p is None:
                p = Para(None, "rtf", [])
                p.start = story.last_end
                p.mark = None
                p.protected = False
                story.para = p
            p.label = text.replace("\t", " ").strip()
        elif kind in ("bkmkstart", "bkmkend"):
            p = story.para
            if p is None:
                p = Para(None, "rtf", [])
                p.start = story.last_end
                p.mark = None
                p.protected = False
                story.para = p
            pos = len(p.text)
            self.marks_raw.append((text.strip(), "start" if kind == "bkmkstart" else "end", p, pos))
        elif kind in ("ffname", "ffl", "ffdeftext", "ffhelptext", "ffstattext") and fields:
            ff = fields[-1][1].ff
            if kind == "ffl":
                ff.setdefault("ffl", []).append(text)
            else:
                ff[kind] = text

    def _open_result(self, fd: _Field, story, t: Tok) -> None:
        instr = "".join(fd.instr).strip()
        word = instr.split()[0].upper() if instr.split() else ""
        p = story.para
        if p is None:
            p = Para(None, "rtf", [])
            p.start = story.last_end
            p.mark = None
            p.protected = False
            story.para = p
        fd.para = p
        fd.n0 = len(p.segs)
        ff = fd.ff
        if word in ("FORMTEXT", "FORMCHECKBOX", "FORMDROPDOWN"):
            kind = {"FORMTEXT": "text", "FORMCHECKBOX": "checkbox", "FORMDROPDOWN": "dropdown"}[word]
            ctl = Control(kind, "formfield", ff.get("ffname", ""), None, "rtf", p, help=ff.get("ffhelptext", "") or "")
            ctl.rtf = True
            ctl.field_s = getattr(fd, "start", t.s)
            if ff.get("ffmaxlen"):
                ctl.maxlen = ff["ffmaxlen"]
            res = ff.get("ffres", 25)
            if res == 25:
                res = ff.get("ffdefres", 0)
            if kind == "checkbox":
                ctl.checked = res == 1
            if kind == "dropdown":
                ctl.options = list(ff.get("ffl", []))
                ctl.value = ctl.options[res] if 0 <= res < len(ctl.options) else None
            if ff.get("fftypetxt") == 2:
                ctl.kind = "date"
            ctl.ff = dict(ff)
            ctl.result_open = t.e
            fd.ctl = ctl
            self.controls.append(ctl)

    def _close_field(self, fd: _Field, story, t: Tok) -> None:
        ctl = fd.ctl
        if ctl is None:
            return
        p = fd.para
        if p is None:
            return
        segs = p.segs[fd.n0:]
        ctl.result_segs = [s for s in segs if s.field is ctl]
        ctl.result_close = getattr(fd, "result_close", None) or t.s
        if ctl.kind == "checkbox":
            sg = Seg("☑" if ctl.checked else "☐", "field", RSrc("ctl", t.s, t.s), None, False, False, ctl)
            p.segs.append(sg)
            ctl.result_segs = [sg]
        elif ctl.kind == "dropdown" and not "".join(s.text for s in ctl.result_segs).strip():
            sg = Seg(ctl.value or "", "field", RSrc("ctl", t.s, t.s), None, False, False, ctl)
            p.segs.append(sg)
            ctl.result_segs = [sg]
        elif not ctl.result_segs:
            sg = Seg("", "field", RSrc("ctl", t.s, t.s), None, False, False, ctl)
            p.segs.append(sg)
            ctl.result_segs = [sg]

    # ------------------------------------------------------------------ 表
    def _row_def(self, story: _Story, st: _St, w: str, param) -> None:
        level = st.itap if st.itap > 0 else 1
        while len(story.levels) < level:
            story.levels.append(_Level())
        lv = story.levels[level - 1]
        if w == "trowd":
            lv.cur_def = []
            lv.pend = {}
            return
        pend = getattr(lv, "pend", None)
        if pend is None:
            pend = lv.pend = {}
        if w == "cellx":
            lv.cur_def.append({"x": param or 0, **pend})
            lv.pend = {}
        else:
            pend[w] = param if param is not None else True

    def _place(self, story: _Story, p: Para | None, level: int, how: str) -> None:
        """段落を置く：表の外なら本文に（開いている表は閉じる）、表の中なら、その段のいまのマスに。
        深い段の表（入れ子の表）が開いていれば、先に閉じて、このマスの中身にする。\\cell・\\nestcell なら、マスを閉じる。"""
        if how == "cell":
            level = 1
        elif how == "nestcell":
            level = max(level, 2)
        if level <= 0:
            self._close_levels(story)
            if p is not None:
                story.blocks.append(p)
            return
        while len(story.levels) < level:
            story.levels.append(_Level())
        lv = story.levels[level - 1]
        while len(story.levels) > level:
            inner = self._close_level(story, attach=False)
            if inner is not None:
                lv.cell.blocks.append(inner)
        if p is not None:
            lv.cell.blocks.append(p)
        if how in ("cell", "nestcell"):
            lv.row.append(lv.cell)
            lv.cell = _Cell()

    def _end_row(self, story: _Story, level: int, t: Tok) -> None:
        while len(story.levels) < level:
            story.levels.append(_Level())
        while len(story.levels) > level:
            inner = self._close_level(story, attach=False)
            if inner is not None:
                story.levels[level - 1].cell.blocks.append(inner)
        lv = story.levels[level - 1]
        if lv.cell.blocks:
            lv.row.append(lv.cell)
            lv.cell = _Cell()
        lv.rows.append(lv.row)
        lv.defs.append(list(lv.cur_def))
        lv.row = []
        story.last_end = t.e
        if story.para is not None and not story.para.segs:
            story.para.start = t.e

    def _close_levels(self, story: _Story) -> None:
        while story.levels:
            self._close_level(story)

    def _close_level(self, story: _Story, attach: bool = True) -> Table | None:
        lv = story.levels.pop()
        if lv.row or lv.cell.blocks:
            if lv.cell.blocks:
                lv.row.append(lv.cell)
            lv.rows.append(lv.row)
            lv.defs.append(list(lv.cur_def))
        if not lv.rows:
            return None
        t = self._make_table(lv)
        if attach:
            if story.levels:
                story.levels[-1].cell.blocks.append(t)
            else:
                story.blocks.append(t)
        return t

    def _make_table(self, lv: _Level) -> Table:
        """行ごとのマスの右の端（\\cellx）を合わせた格子で、マスの列と幅を決める。横の結合（\\clmrg）は前のマスに含める。"""
        t = Table(None, "rtf")
        grid = sorted({d["x"] for defs in lv.defs for d in defs if "x" in d})
        for r, (cells, defs) in enumerate(zip(lv.rows, lv.defs)):
            row: list[Cell] = []
            prev_x = None
            seq = 0
            for k, cc in enumerate(cells):
                d = defs[k] if k < len(defs) else {}
                x = d.get("x")
                if d.get("clmrg") and row:
                    last = row[-1]
                    last.blocks += cc.blocks
                    for p in cc.blocks:
                        if isinstance(p, Para):
                            p.cell = last
                    if x in grid:
                        last.span = max(last.span, grid.index(x) - last.col + 1)
                    prev_x = x
                    continue
                if x is not None and x in grid:
                    c0 = 0 if prev_x is None or prev_x not in grid else grid.index(prev_x) + 1
                    c1 = grid.index(x)
                    col, span = c0, max(1, c1 - c0 + 1)
                else:
                    col, span = (row[-1].col + row[-1].span) if row else seq, 1
                vm = "continue" if d.get("clvmrg") else ("restart" if d.get("clvmgf") else "")
                c = Cell(None, r, col, span, vm, table=t)
                c.blocks = cc.blocks
                for p in c.blocks:
                    if isinstance(p, Para):
                        p.cell = c
                row.append(c)
                prev_x = x
                seq = col + span
            t.rows.append(row)
        return t

    def _attach(self, inner: _Story, outer: _Story) -> None:
        """テキストボックス・図形の字は、置いた段落の枠（boxes）にする。"""
        if inner.kind == "box":
            anchor = getattr(inner, "anchor", None)
            if anchor is None:
                return
            anchor.boxes += inner.blocks

    # ------------------------------------------------------------------ Word と同じ問い合わせ
    def all_blocks(self):
        yield "body", "rtf", self.body
        for kind, part, blocks in self.stories:
            yield kind, part, blocks

    def marks(self) -> dict[str, tuple[Para, int, int]]:
        starts: dict[str, tuple[Para, int]] = {}
        out: dict[str, tuple[Para, int, int]] = {}
        for name, which, p, pos in self.marks_raw:
            if not name.startswith(MARK_PREFIX):
                continue
            key = name[len(MARK_PREFIX):].replace("_", ".")
            if which == "start":
                starts[key] = (p, pos)
            elif key in starts and starts[key][0] is p:
                out[key] = (p, starts[key][1], pos)
        return out


def paragraphs(doc: RtfDocument) -> list[Para]:
    return [p for _k, _p, bl in doc.all_blocks() for p in iter_paras(bl)]
