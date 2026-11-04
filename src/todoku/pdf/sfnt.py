"""埋め込まれたフォントのプログラムから、文字の対応を読む（TrueType の cmap・post、CFF の charset・encoding、Type1 の Encoding）。

ToUnicode を持たないフォントの文字を決めるために使う。形が壊れていても止まらず、読めた分だけ返す。
"""

from __future__ import annotations

import re
import struct

from ..limits import budget
from .ranges import RangeMap


# ====================================================================== TrueType / OpenType
def tables(data: bytes) -> dict[str, bytes]:
    out: dict[str, bytes] = {}
    if len(data) < 12:
        return out
    try:
        n = struct.unpack_from(">H", data, 4)[0]
        room = 2 * len(data)  # 本物の表は重ならない：宣言の 16 バイトが、同じ大きな場所を何度も写させないように
        for i in range(n):
            tag, _chk, off, ln = struct.unpack_from(">4sIII", data, 12 + 16 * i)
            piece = data[off:off + min(ln, room)]
            room -= len(piece)
            if room < 0:
                break
            out[tag.decode("latin-1")] = piece
    except struct.error:
        pass
    return out


def cmap_tables(t: bytes) -> dict[tuple[int, int], dict[int, int]]:
    """cmap の表から、(platform, encoding) → {文字コード: グリフ番号}。同じ位置を指す記録は 1 度だけ読む。読めない部分表は除く。"""
    out: dict[tuple[int, int], dict[int, int]] = {}
    if len(t) < 4:
        return out
    read: dict[int, dict[int, int]] = {}
    try:
        n = struct.unpack_from(">H", t, 2)[0]
        for i in range(n):
            pid, eid, off = struct.unpack_from(">HHI", t, 4 + 8 * i)
            if off not in read:
                try:
                    read[off] = _subtable(t, off)
                except (struct.error, IndexError):
                    read[off] = {}
            if read[off]:
                out.setdefault((pid, eid), read[off])
    except struct.error:
        pass
    return out


def cmap_subtables(data: bytes) -> dict[tuple[int, int], dict[int, int]]:
    """(platform, encoding) → {文字コード: グリフ番号}。"""
    return cmap_tables(tables(data).get("cmap") or b"")


def _format4(t: bytes, off: int) -> RangeMap:
    """形式 4：区切りのまま持ち、引くときに値を求める（区切りが重なっても、広げた数は番号の数 65,536 まで）。"""
    segx2 = struct.unpack_from(">H", t, off + 6)[0]
    seg = segx2 // 2
    ends = struct.unpack_from(f">{seg}H", t, off + 14)
    starts = struct.unpack_from(f">{seg}H", t, off + 16 + segx2)
    deltas = struct.unpack_from(f">{seg}h", t, off + 16 + 2 * segx2)
    roff_pos = off + 16 + 3 * segx2
    roffs = struct.unpack_from(f">{seg}H", t, roff_pos)

    def value(i: int, c: int) -> int | None:
        s, d, ro = starts[i], deltas[i], roffs[i]
        if ro == 0:
            return ((c + d) & 0xFFFF) or None
        p = roff_pos + 2 * i + ro + 2 * (c - s)
        if p + 2 > len(t):
            return None
        g = struct.unpack_from(">H", t, p)[0]
        return (((g + d) & 0xFFFF) or None) if g else None

    return RangeMap([(starts[i], ends[i], i) for i in range(seg) if starts[i] != 0xFFFF], value)


def _format12_value(v: tuple[int, int], c: int) -> int:
    return v[0] + c - v[1]


def _subtable(t: bytes, off: int) -> dict[int, int]:
    fmt = struct.unpack_from(">H", t, off)[0]
    m: dict[int, int] = {}
    if fmt == 0:
        for c in range(256):
            g = t[off + 6 + c]
            if g:
                m[c] = g
    elif fmt == 4:
        return _format4(t, off)
    elif fmt == 6:
        first, count = struct.unpack_from(">HH", t, off + 6)
        for i, g in enumerate(struct.unpack_from(f">{count}H", t, off + 10)):
            if g:
                m[first + i] = g
    elif fmt == 12:
        # 区切り 12 バイトが、広げると 110 万字になる：区切りのまま持つ。Unicode の外の番号（0x10FFFF より上）は字にならない
        n = min(struct.unpack_from(">I", t, off + 12)[0], (len(t) - off - 16) // 12)  # 宣言でなく、実際にある数
        groups = []
        for i in range(n):
            s, e, g0 = struct.unpack_from(">III", t, off + 16 + 12 * i)
            if e - s <= 0x10FFFF and s <= 0x10FFFF:
                groups.append((s, min(e, 0x10FFFF), (g0, s)))
        return RangeMap(groups, _format12_value)
    return m


def _glyph_codes(m: dict[int, int]) -> dict[int, int]:
    """cmap を逆に引く：グリフ番号 → それを指す、いちばん小さな文字コード（Unicode の字になるものだけ）。"""
    if isinstance(m, RangeMap) and m.value is _format12_value:
        # 区切りのまま逆にする：後に書いたものが勝つので、大きな文字コードから順に（結果は、グリフの番号の数 65,536 まで）
        rev = []
        for lo, hi, (g0, s) in reversed(m.segments()):
            for a, b in ((max(lo, 0xE000), hi), (max(lo, 1), min(hi, 0xD7FF))):
                g_a = g0 + a - s
                if a <= b and g_a <= 0xFFFF:
                    rev.append((g_a, min(g_a + b - a, 0xFFFF), (a, g_a)))
        return dict(RangeMap(rev, _inverse_value).items())
    out: dict[int, int] = {}
    for code, g in sorted(m.items()):
        if g not in out and 0 < code <= 0x10FFFF and not 0xD800 <= code <= 0xDFFF:
            out[g] = code
    return out


def _inverse_value(v: tuple[int, int], g: int) -> int:
    return v[0] + g - v[1]


def unicode_by_glyph(data: bytes) -> dict[int, str]:
    """グリフ番号 → 文字。Unicode の cmap を逆に引き、なければ post のグリフ名から決める。"""
    from .fonts import glyph_to_unicode

    out: dict[int, str] = {}
    subs = cmap_subtables(data)
    for key in ((3, 10), (0, 6), (0, 4), (3, 1), (0, 3), (0, 2), (0, 1), (0, 0)):
        m = subs.get(key)
        if not m:
            continue
        for g, code in _glyph_codes(m).items():
            out.setdefault(g, chr(code))
    for g, name in post_names(data).items():
        if g not in out:
            u = glyph_to_unicode(name)
            if u:
                out[g] = u
    # cmap にない字形（縦書き用の形・合字・色の重ね）：置き換え（GSUB）と色の重ね（COLR）から、もとの字に結びつける
    try:
        subs_map = gsub_sources(data)
    except (struct.error, IndexError):
        subs_map = {}
    try:
        layers = colr_layers(data)
    except (struct.error, IndexError):
        layers = {}
    for _round in range(4):
        grew = False
        for g, srcs in subs_map.items():
            if g in out:
                continue
            us = [out.get(x) for x in srcs]
            if all(us):
                out[g] = "".join(us)
                grew = True
        if not grew:
            break
    for g, base in layers.items():
        if g not in out and base in out:
            out[g] = out[base]
    return out


def _coverage(t: bytes, off: int) -> list[int]:
    fmt = struct.unpack_from(">H", t, off)[0]
    if fmt == 1:
        n = struct.unpack_from(">H", t, off + 2)[0]
        return list(struct.unpack_from(f">{n}H", t, off + 4))
    if fmt == 2:
        n = struct.unpack_from(">H", t, off + 2)[0]
        out = []
        last = -1  # 範囲は順に並び、重ならない（表 Coverage の決まり）：重なった所は数えない（小さな表が 65,536 字を何度も広げないように）
        for i in range(n):
            s, e, _k = struct.unpack_from(">HHH", t, off + 4 + 6 * i)
            s = max(s, last + 1)
            if s <= e:
                budget().count(e - s + 1, "GSUB の字形の範囲")
                out.extend(range(s, e + 1))
                last = e
        return out
    return []


def gsub_sources(data: bytes) -> dict[int, list[int]]:
    """字形の置き換え（GSUB）を逆にたどる：置き換えた後の字形 → もとの字形の並び（1字、または合字の部品）。"""
    t = tables(data).get("GSUB")
    out: dict[int, list[int]] = {}
    if not t or len(t) < 10:
        return out
    lookup_list = struct.unpack_from(">H", t, 8)[0]
    n = struct.unpack_from(">H", t, lookup_list)[0]
    for i in range(n):
        lo = lookup_list + struct.unpack_from(">H", t, lookup_list + 2 + 2 * i)[0]
        ltype, _flag, nsub = struct.unpack_from(">HHH", t, lo)
        for j in range(nsub):
            so = lo + struct.unpack_from(">H", t, lo + 6 + 2 * j)[0]
            typ = ltype
            if typ == 7:
                _fmt, typ, ext = struct.unpack_from(">HHI", t, so)
                so = so + ext
            fmt = struct.unpack_from(">H", t, so)[0]
            if typ == 1:
                cov = _coverage(t, so + struct.unpack_from(">H", t, so + 2)[0])
                if fmt == 1:
                    delta = struct.unpack_from(">h", t, so + 4)[0]
                    for g in cov:
                        out.setdefault((g + delta) & 0xFFFF, [g])
                elif fmt == 2:
                    cnt = struct.unpack_from(">H", t, so + 4)[0]
                    subs = struct.unpack_from(f">{cnt}H", t, so + 6)
                    for g, s2 in zip(cov, subs):
                        out.setdefault(s2, [g])
            elif typ == 3 and fmt == 1:
                cov = _coverage(t, so + struct.unpack_from(">H", t, so + 2)[0])
                cnt = struct.unpack_from(">H", t, so + 4)[0]
                for k in range(min(cnt, len(cov))):
                    ao = so + struct.unpack_from(">H", t, so + 6 + 2 * k)[0]
                    na = struct.unpack_from(">H", t, ao)[0]
                    budget().count(na, "GSUB の置き換え")
                    for a in struct.unpack_from(f">{na}H", t, ao + 2):
                        out.setdefault(a, [cov[k]])
            elif typ == 4 and fmt == 1:
                cov = _coverage(t, so + struct.unpack_from(">H", t, so + 2)[0])
                cnt = struct.unpack_from(">H", t, so + 4)[0]
                for k in range(min(cnt, len(cov))):
                    lso = so + struct.unpack_from(">H", t, so + 6 + 2 * k)[0]
                    nl = struct.unpack_from(">H", t, lso)[0]
                    budget().count(nl, "GSUB の合字")
                    for q in range(nl):
                        lgo = lso + struct.unpack_from(">H", t, lso + 2 + 2 * q)[0]
                        lig, comp = struct.unpack_from(">HH", t, lgo)
                        budget().count(comp, "GSUB の合字の部品")
                        rest = list(struct.unpack_from(f">{comp - 1}H", t, lgo + 4)) if comp > 1 else []
                        out.setdefault(lig, [cov[k]] + rest)
    return out


def colr_layers(data: bytes) -> dict[int, int]:
    """色の付いた字（COLR 版0）：重ねる層の字形 → もとの字形。"""
    t = tables(data).get("COLR")
    out: dict[int, int] = {}
    if not t or len(t) < 14:
        return out
    _ver, nbase, boff, loff, nlayer = struct.unpack_from(">HHIIH", t, 0)
    layers = [struct.unpack_from(">HH", t, loff + 4 * i)[0] for i in range(nlayer)]
    for i in range(nbase):
        g, first, num = struct.unpack_from(">HHH", t, boff + 6 * i)
        budget().count(min(num, len(layers)), "色の字の層")
        for k in range(first, min(first + num, len(layers))):
            out.setdefault(layers[k], g)
    return out


def symbol_cmap(data: bytes) -> dict[int, int]:
    """記号のフォント（cmap 3,0）の 文字コード → グリフ番号。0xF000 を足した番号でも引ける。"""
    subs = cmap_subtables(data)
    m = subs.get((3, 0)) or subs.get((1, 0)) or {}
    out: dict[int, int] = {}
    for c, g in m.items():
        out[c & 0xFF] = g
    return out


def post_names(data: bytes) -> dict[int, str]:
    t = tables(data).get("post")
    out: dict[int, str] = {}
    if not t or len(t) < 34:
        return out
    try:
        ver = struct.unpack_from(">I", t, 0)[0]
        if ver == 0x00020000:
            n = struct.unpack_from(">H", t, 32)[0]
            idx = struct.unpack_from(f">{n}H", t, 34)
            p = 34 + 2 * n
            names = []
            while p < len(t):
                ln = t[p]
                names.append(t[p + 1:p + 1 + ln].decode("latin-1"))
                p += 1 + ln
            for g, k in enumerate(idx):
                if k >= 258 and k - 258 < len(names):
                    out[g] = names[k - 258]
                elif k < 258:
                    out[g] = MAC_GLYPHS[k] if k < len(MAC_GLYPHS) else ""
    except (struct.error, IndexError):
        pass
    return {g: n for g, n in out.items() if n}


MAC_GLYPHS = (
    ".notdef .null nonmarkingreturn space exclam quotedbl numbersign dollar percent ampersand quotesingle parenleft "
    "parenright asterisk plus comma hyphen period slash zero one two three four five six seven eight nine colon semicolon "
    "less equal greater question at A B C D E F G H I J K L M N O P Q R S T U V W X Y Z bracketleft backslash bracketright "
    "asciicircum underscore grave a b c d e f g h i j k l m n o p q r s t u v w x y z braceleft bar braceright asciitilde "
    "Adieresis Aring Ccedilla Eacute Ntilde Odieresis Udieresis aacute agrave acircumflex adieresis atilde aring ccedilla "
    "eacute egrave ecircumflex edieresis iacute igrave icircumflex idieresis ntilde oacute ograve ocircumflex odieresis "
    "otilde uacute ugrave ucircumflex udieresis dagger degree cent sterling section bullet paragraph germandbls registered "
    "copyright trademark acute dieresis notequal AE Oslash infinity plusminus lessequal greaterequal yen mu partialdiff "
    "summation product pi integral ordfeminine ordmasculine Omega ae oslash questiondown exclamdown logicalnot radical "
    "florin approxequal Delta guillemotleft guillemotright ellipsis nonbreakingspace Agrave Atilde Otilde OE oe endash "
    "emdash quotedblleft quotedblright quoteleft quoteright divide lozenge ydieresis Ydieresis fraction currency "
    "guilsinglleft guilsinglright fi fl daggerdbl periodcentered quotesinglbase quotedblbase perthousand Acircumflex "
    "Ecircumflex Aacute Edieresis Egrave Iacute Icircumflex Idieresis Igrave Oacute Ocircumflex apple Ograve Uacute "
    "Ucircumflex Ugrave dotlessi circumflex tilde macron breve dotaccent ring cedilla hungarumlaut ogonek caron Lslash "
    "lslash Scaron scaron Zcaron zcaron brokenbar Eth eth Yacute yacute Thorn thorn minus multiply onesuperior twosuperior "
    "threesuperior onehalf onequarter threequarters franc Gbreve gbreve Idotaccent Scedilla scedilla Cacute cacute Ccaron "
    "ccaron dcroat").split()


# ====================================================================== CFF
def _cff_index(d: bytes, p: int) -> tuple[list[bytes], int]:
    count = struct.unpack_from(">H", d, p)[0]
    if count == 0:
        return [], p + 2
    osz = d[p + 2]
    offs = []
    q = p + 3
    for _ in range(count + 1):
        offs.append(int.from_bytes(d[q:q + osz], "big"))
        q += osz
    base = q - 1
    # 本物の索引は並びの順で重ならない：宣言の位置が、同じ大きな場所を何度も写させないように
    if sum(max(0, offs[i + 1] - offs[i]) for i in range(count)) > len(d):
        raise struct.error("CFF の索引の位置が重なっています")
    items = [d[base + offs[i]:base + offs[i + 1]] for i in range(count)]
    return items, base + offs[-1]


def _cff_dict(b: bytes) -> dict[int, list]:
    out: dict[int, list] = {}
    ops: list = []
    i = 0
    while i < len(b):
        v = b[i]
        if v <= 21:
            if v == 12:
                key = 1200 + b[i + 1]
                i += 2
            else:
                key = v
                i += 1
            out[key] = ops
            ops = []
        elif v == 28:
            ops.append(struct.unpack_from(">h", b, i + 1)[0])
            i += 3
        elif v == 29:
            ops.append(struct.unpack_from(">i", b, i + 1)[0])
            i += 5
        elif v == 30:
            s = ""
            i += 1
            done = False
            while i < len(b) and not done:
                for nib in (b[i] >> 4, b[i] & 15):
                    if nib == 15:
                        done = True
                        break
                    s += "0123456789.EE?-"[nib] if nib != 12 else "E-"
                i += 1
            try:
                ops.append(float(s.replace("?", "")))
            except ValueError:
                ops.append(0.0)
        elif 32 <= v <= 246:
            ops.append(v - 139)
            i += 1
        elif 247 <= v <= 250:
            ops.append((v - 247) * 256 + b[i + 1] + 108)
            i += 2
        elif 251 <= v <= 254:
            ops.append(-(v - 251) * 256 - b[i + 1] - 108)
            i += 2
        else:
            i += 1
    return out


STD_STRINGS_HEAD = MAC_GLYPHS  # 使わない（CFF の標準文字列は下で定義）


def cff_glyph_info(data: bytes) -> tuple[dict[int, str], dict[int, int], bool]:
    """CFF：(グリフ番号 → グリフ名, 文字コード → グリフ番号（組み込みの encoding）, CIDで引くフォントか)。"""
    names: dict[int, str] = {}
    enc: dict[int, int] = {}
    try:
        hdr = data[2]
        p = hdr
        _names, p = _cff_index(data, p)
        tops, p = _cff_index(data, p)
        strings, p = _cff_index(data, p)
        top = _cff_dict(tops[0]) if tops else {}
        is_cid = 1230 in top
        cs_off = top.get(17, [0])[0]
        charstrings, _ = _cff_index(data, cs_off) if cs_off else ([], 0)
        n = len(charstrings)

        def sid_name(sid: int) -> str:
            if sid < len(CFF_STD):
                return CFF_STD[sid]
            k = sid - len(CFF_STD)
            return strings[k].decode("latin-1") if k < len(strings) else ""

        cs_charset = top.get(15, [0])[0]
        sids: list[int] = [0]
        if cs_charset > 2:
            fmt = data[cs_charset]
            q = cs_charset + 1
            if fmt == 0:
                for _ in range(n - 1):
                    sids.append(struct.unpack_from(">H", data, q)[0])
                    q += 2
            else:
                while len(sids) < n:
                    first = struct.unpack_from(">H", data, q)[0]
                    if fmt == 1:
                        left = data[q + 2]
                        q += 3
                    else:
                        left = struct.unpack_from(">H", data, q + 2)[0]
                        q += 4
                    for k in range(left + 1):
                        sids.append(first + k)
        elif cs_charset == 0:
            sids = list(range(n))
        if not is_cid:
            for g, sid in enumerate(sids[:n]):
                names[g] = sid_name(sid)
            enc_off = top.get(16, [0])[0]
            if enc_off > 1:
                fmt = data[enc_off] & 0x7F
                q = enc_off + 1
                if fmt == 0:
                    cnt = data[q]
                    for g in range(1, cnt + 1):
                        enc[data[q + g]] = g
                elif fmt == 1:
                    nr = data[q]
                    g = 1
                    q += 1
                    for _ in range(nr):
                        first, left = data[q], data[q + 1]
                        for k in range(left + 1):
                            enc[first + k] = g
                            g += 1
                        q += 2
            else:
                # 標準の encoding：グリフ名から引く
                from .fonts import STANDARD_ENCODING

                by_name = {nm: g for g, nm in names.items()}
                for code, nm in enumerate(STANDARD_ENCODING):
                    if nm and nm in by_name:
                        enc[code] = by_name[nm]
        return names, enc, is_cid
    except (struct.error, IndexError):
        return names, enc, False


def type1_encoding(data: bytes) -> dict[int, str]:
    """Type1 フォントの平文の部分にある組み込みの Encoding（dup 65 /A put）。"""
    head = data[:max(0, data.find(b"eexec"))] if b"eexec" in data else data[:20000]
    out: dict[int, str] = {}
    if b"/Encoding StandardEncoding" in head:
        from .fonts import STANDARD_ENCODING

        return {i: n for i, n in enumerate(STANDARD_ENCODING) if n}
    for m in re.finditer(rb"dup\s+(\d+)\s*/([^\s/\[\]{}()<>%]+)\s+put", head):
        out[int(m[1])] = m[2].decode("latin-1")
    return out


CFF_STD = (
    ".notdef space exclam quotedbl numbersign dollar percent ampersand quoteright parenleft parenright asterisk plus comma "
    "hyphen period slash zero one two three four five six seven eight nine colon semicolon less equal greater question at "
    "A B C D E F G H I J K L M N O P Q R S T U V W X Y Z bracketleft backslash bracketright asciicircum underscore quoteleft "
    "a b c d e f g h i j k l m n o p q r s t u v w x y z braceleft bar braceright asciitilde exclamdown cent sterling "
    "fraction yen florin section currency quotesingle quotedblleft guillemotleft guilsinglleft guilsinglright fi fl endash "
    "dagger daggerdbl periodcentered paragraph bullet quotesinglbase quotedblbase quotedblright guillemotright ellipsis "
    "perthousand questiondown grave acute circumflex tilde macron breve dotaccent dieresis ring cedilla hungarumlaut ogonek "
    "caron emdash AE ordfeminine Lslash Oslash OE ordmasculine ae dotlessi lslash oslash oe germandbls onesuperior "
    "logicalnot mu trademark Eth onehalf plusminus Thorn onequarter divide brokenbar degree thorn threequarters twosuperior "
    "registered minus eth multiply threesuperior copyright Aacute Acircumflex Adieresis Agrave Aring Atilde Ccedilla Eacute "
    "Ecircumflex Edieresis Egrave Iacute Icircumflex Idieresis Igrave Ntilde Oacute Ocircumflex Odieresis Ograve Otilde "
    "Scaron Uacute Ucircumflex Udieresis Ugrave Yacute Ydieresis Zcaron aacute acircumflex adieresis agrave aring atilde "
    "ccedilla eacute ecircumflex edieresis egrave iacute icircumflex idieresis igrave ntilde oacute ocircumflex odieresis "
    "ograve otilde scaron uacute ucircumflex udieresis ugrave yacute ydieresis zcaron exclamsmall Hungarumlautsmall "
    "dollaroldstyle dollarsuperior ampersandsmall Acutesmall parenleftsuperior parenrightsuperior twodotenleader "
    "onedotenleader zerooldstyle oneoldstyle twooldstyle threeoldstyle fouroldstyle fiveoldstyle sixoldstyle sevenoldstyle "
    "eightoldstyle nineoldstyle commasuperior threequartersemdash periodsuperior questionsmall asuperior bsuperior "
    "centsuperior dsuperior esuperior isuperior lsuperior msuperior nsuperior osuperior rsuperior ssuperior tsuperior ff ffi "
    "ffl parenleftinferior parenrightinferior Circumflexsmall hyphensuperior Gravesmall Asmall Bsmall Csmall Dsmall Esmall "
    "Fsmall Gsmall Hsmall Ismall Jsmall Ksmall Lsmall Msmall Nsmall Osmall Psmall Qsmall Rsmall Ssmall Tsmall Usmall Vsmall "
    "Wsmall Xsmall Ysmall Zsmall colonmonetary onefitted rupiah Tildesmall exclamdownsmall centoldstyle Lslashsmall "
    "Scaronsmall Zcaronsmall Dieresissmall Brevesmall Caronsmall Dotaccentsmall Macronsmall figuredash hypheninferior "
    "Ogoneksmall Ringsmall Cedillasmall questiondownsmall oneeighth threeeighths fiveeighths seveneighths onethird "
    "twothirds zerosuperior foursuperior fivesuperior sixsuperior sevensuperior eightsuperior ninesuperior zeroinferior "
    "oneinferior twoinferior threeinferior fourinferior fiveinferior sixinferior seveninferior eightinferior nineinferior "
    "centinferior dollarinferior periodinferior commainferior Agravesmall Aacutesmall Acircumflexsmall Atildesmall "
    "Adieresissmall Aringsmall AEsmall Ccedillasmall Egravesmall Eacutesmall Ecircumflexsmall Edieresissmall Igravesmall "
    "Iacutesmall Icircumflexsmall Idieresissmall Ethsmall Ntildesmall Ogravesmall Oacutesmall Ocircumflexsmall Otildesmall "
    "Odieresissmall OEsmall Oslashsmall Ugravesmall Uacutesmall Ucircumflexsmall Udieresissmall Yacutesmall Thornsmall "
    "Ydieresissmall 001.000 001.001 001.002 001.003 Black Bold Book Light Medium Regular Roman Semibold").split()
