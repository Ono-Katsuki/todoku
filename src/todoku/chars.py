"""字の種類（どの言語の書類にも同じに使う、Unicode の資料で決まる字の集まり）。

窓の「書き込む所」は、字の意味（「年」「様」などの語）からは決めず、字の形・並び（空白の並び・下線の字の並び・
点線・かっこの組・チェックの箱の字）だけで示す。ここに、その字の集まりを、出典つきで1か所にまとめる。

- 空白：一般カテゴリ Zs（U+0020・U+00A0・U+2000〜U+200A・U+3000 など）
- 下線の字：一般カテゴリ Pc（connector punctuation：_ ‿ ⁀ ⁔ ︳ ︴ ﹍ ﹎ ﹏ ＿）
- ダッシュ：一般カテゴリ Pd（dash punctuation：- ‐ – — ― － など）
- 点線（リーダー）：下の LEADERS（名前に DOT LEADER・ELLIPSIS・MIDDLE DOT を含む句読点と FULL STOP）
- かっこの組：BidiBrackets.txt（Bidi_Paired_Bracket）
- チェックの箱の字：Ballot Box（U+2610〜U+2612）・Geometric Shapes の白い四角（U+25A1・U+25A2・U+25FB・U+25FD）・
  Dingbats の影つきの四角（U+274F〜U+2752）・U+2B1C
- 文の終わりの字：PropList.txt の Sentence_Terminal（。 . ! ? । ۔ ؟ ። ։ など）
- 見出しのコロン：一般カテゴリ Po で、名前に語 COLON を含む字（: ： ﹕ ︓ ፥ ᠄ ܃ など）。
  U+061B ؛（ARABIC SEMICOLON）はセミコロン、U+0589 ։（ARMENIAN FULL STOP）は文の終わりの字で、どちらもコロンではない

表の部分は scripts/make_chars.py が Unicode の資料から作る（版は Python の unicodedata と同じ。UCD_VERSION）。
"""

from __future__ import annotations

import bisect
import unicodedata


def is_space(ch: str) -> bool:
    """空白（一般カテゴリ Zs）。"""
    return unicodedata.category(ch) == "Zs"


def is_connector(ch: str) -> bool:
    """下線の字（一般カテゴリ Pc）。"""
    return unicodedata.category(ch) == "Pc"


def is_dash(ch: str) -> bool:
    """ダッシュ（一般カテゴリ Pd）。"""
    return unicodedata.category(ch) == "Pd"


# 点線（リーダー）に使う字。Unicode の名前（unicodedata.name）で選んだもの：
#   U+002E FULL STOP・U+FF0E FULLWIDTH FULL STOP・U+00B7 MIDDLE DOT・U+2024 ONE DOT LEADER・U+2025 TWO DOT LEADER・
#   U+2026 HORIZONTAL ELLIPSIS・U+22EF MIDLINE HORIZONTAL ELLIPSIS・U+2E31 WORD SEPARATOR MIDDLE DOT・
#   U+30FB KATAKANA MIDDLE DOT・U+FF65 HALFWIDTH KATAKANA MIDDLE DOT・U+1801 MONGOLIAN ELLIPSIS・
#   U+FE19 PRESENTATION FORM FOR VERTICAL HORIZONTAL ELLIPSIS・U+FE30 PRESENTATION FORM FOR VERTICAL TWO DOT LEADER
LEADERS = frozenset(".．·․‥…⋯⸱・･᠁︙︰")
# 1 字で 2 つ以上の点を表す字（「…」「‥」）。1 字だけでも点の並び（2 字以上の点線と同じ）として数える
MULTI_DOT = frozenset("‥…⋯᠁︙︰")

# チェックの箱の字（印のないもの）：U+2610 BALLOT BOX・U+25A1 WHITE SQUARE・U+25A2 WHITE SQUARE WITH ROUNDED CORNERS・
# U+25FB WHITE MEDIUM SQUARE・U+25FD WHITE MEDIUM SMALL SQUARE・U+274F〜U+2752（影つきの白い四角）・
# U+2B1C WHITE LARGE SQUARE
EMPTY_BOXES = frozenset("☐□▢◻◽❏❐❑❒⬜")
# 印のついた箱：U+2611 BALLOT BOX WITH CHECK・U+2612 BALLOT BOX WITH X（Ballot Box の残りの 2 字。チェック済みと決まる字）
CHECKED_BOXES = frozenset("☑☒")
BOXES = EMPTY_BOXES | CHECKED_BOXES
# 印のついた箱かもしれない字（U+25A0 BLACK SQUARE は箇条の印・飾りのことも、U+2713/U+2714 CHECK MARK は本文の印のこともある）
MAYBE_CHECKED = frozenset("■✓✔")


def is_wide(ch: str) -> bool:
    """東アジアの字の幅（UAX #11）が W（全角）か F（全角の互換用の字）。"""
    return unicodedata.east_asian_width(ch) in ("W", "F")


def closing(ch: str) -> str | None:
    """開きかっこ ch と対になる閉じかっこ（BidiBrackets.txt）。かっこでなければ None。"""
    return BRACKET_PAIRS.get(ch)


def _in(ranges, starts, ch: str) -> bool:
    c = ord(ch)
    i = bisect.bisect_right(starts, c) - 1
    return i >= 0 and ranges[i][0] <= c <= ranges[i][1]


def is_sentence_terminal(ch: str) -> bool:
    """文の終わりの字（PropList.txt の Sentence_Terminal：。 . ! ? । ۔ ؟ ። ։ など）。"""
    return _in(SENTENCE_TERMINAL_RANGES, _TERM_STARTS, ch)


def is_terminal_punctuation(ch: str) -> bool:
    """文・句の区切りの字（PropList.txt の Terminal_Punctuation：文の終わりの字に、読点・コロン・セミコロンを足したもの）。"""
    return _in(TERMINAL_PUNCTUATION_RANGES, _TP_STARTS, ch)


def is_label_colon(ch: str) -> bool:
    """見出しのあとのコロン（一般カテゴリ Po で、名前に語 COLON を含む字）。"""
    return ch in LABEL_COLONS


def char_runs(text: str) -> list[tuple[int, int, str]]:
    """同じ種類の字の並び：(始め, 終わり, 種類)。種類は "space"（Zs）・"connector"（Pc）・"leader"（点線の字）・
    "dash"（Pd）。ほかの字は出さない。"""
    out: list[tuple[int, int, str]] = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        kind = ("space" if is_space(ch) else "connector" if is_connector(ch) else "leader" if ch in LEADERS
                else "dash" if is_dash(ch) else None)
        if kind is None:
            i += 1
            continue
        j = i + 1
        while j < n:
            c2 = text[j]
            k2 = ("space" if is_space(c2) else "connector" if is_connector(c2) else "leader" if c2 in LEADERS
                  else "dash" if is_dash(c2) else None)
            if k2 != kind:
                break
            j += 1
        out.append((i, j, kind))
        i = j
    return out


def dot_run(text: str, a: int, b: int) -> bool:
    """text[a:b]（点線の字の並び）が、点線と読める長さか：3 字以上か、「…」「‥」を含む 2 字以上。"""
    n = b - a
    return n >= 3 or (n >= 2 and any(ch in MULTI_DOT for ch in text[a:b]))


# ---- ここから scripts/make_chars.py が作る（手で直さない）
UCD_VERSION = "14.0.0"  # unicodedata.unidata_version と同じ版で作った

# BidiBrackets.txt：開きかっこ → 閉じかっこ（Bidi_Paired_Bracket_Type = Open）
BRACKET_PAIRS: dict[str, str] = {
    "\u0028": "\u0029",  # LEFT PARENTHESIS / RIGHT PARENTHESIS
    "\u005b": "\u005d",  # LEFT SQUARE BRACKET / RIGHT SQUARE BRACKET
    "\u007b": "\u007d",  # LEFT CURLY BRACKET / RIGHT CURLY BRACKET
    "\u0f3a": "\u0f3b",  # TIBETAN MARK GUG RTAGS GYON / TIBETAN MARK GUG RTAGS GYAS
    "\u0f3c": "\u0f3d",  # TIBETAN MARK ANG KHANG GYON / TIBETAN MARK ANG KHANG GYAS
    "\u169b": "\u169c",  # OGHAM FEATHER MARK / OGHAM REVERSED FEATHER MARK
    "\u2045": "\u2046",  # LEFT SQUARE BRACKET WITH QUILL / RIGHT SQUARE BRACKET WITH QUILL
    "\u207d": "\u207e",  # SUPERSCRIPT LEFT PARENTHESIS / SUPERSCRIPT RIGHT PARENTHESIS
    "\u208d": "\u208e",  # SUBSCRIPT LEFT PARENTHESIS / SUBSCRIPT RIGHT PARENTHESIS
    "\u2308": "\u2309",  # LEFT CEILING / RIGHT CEILING
    "\u230a": "\u230b",  # LEFT FLOOR / RIGHT FLOOR
    "\u2329": "\u232a",  # LEFT-POINTING ANGLE BRACKET / RIGHT-POINTING ANGLE BRACKET
    "\u2768": "\u2769",  # MEDIUM LEFT PARENTHESIS ORNAMENT / MEDIUM RIGHT PARENTHESIS ORNAMENT
    "\u276a": "\u276b",  # MEDIUM FLATTENED LEFT PARENTHESIS ORNAMENT / MEDIUM FLATTENED RIGHT PARENTHESIS ORNAMENT
    "\u276c": "\u276d",  # MEDIUM LEFT-POINTING ANGLE BRACKET ORNAMENT / MEDIUM RIGHT-POINTING ANGLE BRACKET ORNAMENT
    "\u276e": "\u276f",  # HEAVY LEFT-POINTING ANGLE QUOTATION MARK ORNAMENT / HEAVY RIGHT-POINTING ANGLE QUOTATION MARK ORNAMENT
    "\u2770": "\u2771",  # HEAVY LEFT-POINTING ANGLE BRACKET ORNAMENT / HEAVY RIGHT-POINTING ANGLE BRACKET ORNAMENT
    "\u2772": "\u2773",  # LIGHT LEFT TORTOISE SHELL BRACKET ORNAMENT / LIGHT RIGHT TORTOISE SHELL BRACKET ORNAMENT
    "\u2774": "\u2775",  # MEDIUM LEFT CURLY BRACKET ORNAMENT / MEDIUM RIGHT CURLY BRACKET ORNAMENT
    "\u27c5": "\u27c6",  # LEFT S-SHAPED BAG DELIMITER / RIGHT S-SHAPED BAG DELIMITER
    "\u27e6": "\u27e7",  # MATHEMATICAL LEFT WHITE SQUARE BRACKET / MATHEMATICAL RIGHT WHITE SQUARE BRACKET
    "\u27e8": "\u27e9",  # MATHEMATICAL LEFT ANGLE BRACKET / MATHEMATICAL RIGHT ANGLE BRACKET
    "\u27ea": "\u27eb",  # MATHEMATICAL LEFT DOUBLE ANGLE BRACKET / MATHEMATICAL RIGHT DOUBLE ANGLE BRACKET
    "\u27ec": "\u27ed",  # MATHEMATICAL LEFT WHITE TORTOISE SHELL BRACKET / MATHEMATICAL RIGHT WHITE TORTOISE SHELL BRACKET
    "\u27ee": "\u27ef",  # MATHEMATICAL LEFT FLATTENED PARENTHESIS / MATHEMATICAL RIGHT FLATTENED PARENTHESIS
    "\u2983": "\u2984",  # LEFT WHITE CURLY BRACKET / RIGHT WHITE CURLY BRACKET
    "\u2985": "\u2986",  # LEFT WHITE PARENTHESIS / RIGHT WHITE PARENTHESIS
    "\u2987": "\u2988",  # Z NOTATION LEFT IMAGE BRACKET / Z NOTATION RIGHT IMAGE BRACKET
    "\u2989": "\u298a",  # Z NOTATION LEFT BINDING BRACKET / Z NOTATION RIGHT BINDING BRACKET
    "\u298b": "\u298c",  # LEFT SQUARE BRACKET WITH UNDERBAR / RIGHT SQUARE BRACKET WITH UNDERBAR
    "\u298d": "\u2990",  # LEFT SQUARE BRACKET WITH TICK IN TOP CORNER / RIGHT SQUARE BRACKET WITH TICK IN TOP CORNER
    "\u298f": "\u298e",  # LEFT SQUARE BRACKET WITH TICK IN BOTTOM CORNER / RIGHT SQUARE BRACKET WITH TICK IN BOTTOM CORNER
    "\u2991": "\u2992",  # LEFT ANGLE BRACKET WITH DOT / RIGHT ANGLE BRACKET WITH DOT
    "\u2993": "\u2994",  # LEFT ARC LESS-THAN BRACKET / RIGHT ARC GREATER-THAN BRACKET
    "\u2995": "\u2996",  # DOUBLE LEFT ARC GREATER-THAN BRACKET / DOUBLE RIGHT ARC LESS-THAN BRACKET
    "\u2997": "\u2998",  # LEFT BLACK TORTOISE SHELL BRACKET / RIGHT BLACK TORTOISE SHELL BRACKET
    "\u29d8": "\u29d9",  # LEFT WIGGLY FENCE / RIGHT WIGGLY FENCE
    "\u29da": "\u29db",  # LEFT DOUBLE WIGGLY FENCE / RIGHT DOUBLE WIGGLY FENCE
    "\u29fc": "\u29fd",  # LEFT-POINTING CURVED ANGLE BRACKET / RIGHT-POINTING CURVED ANGLE BRACKET
    "\u2e22": "\u2e23",  # TOP LEFT HALF BRACKET / TOP RIGHT HALF BRACKET
    "\u2e24": "\u2e25",  # BOTTOM LEFT HALF BRACKET / BOTTOM RIGHT HALF BRACKET
    "\u2e26": "\u2e27",  # LEFT SIDEWAYS U BRACKET / RIGHT SIDEWAYS U BRACKET
    "\u2e28": "\u2e29",  # LEFT DOUBLE PARENTHESIS / RIGHT DOUBLE PARENTHESIS
    "\u2e55": "\u2e56",  # LEFT SQUARE BRACKET WITH STROKE / RIGHT SQUARE BRACKET WITH STROKE
    "\u2e57": "\u2e58",  # LEFT SQUARE BRACKET WITH DOUBLE STROKE / RIGHT SQUARE BRACKET WITH DOUBLE STROKE
    "\u2e59": "\u2e5a",  # TOP HALF LEFT PARENTHESIS / TOP HALF RIGHT PARENTHESIS
    "\u2e5b": "\u2e5c",  # BOTTOM HALF LEFT PARENTHESIS / BOTTOM HALF RIGHT PARENTHESIS
    "\u3008": "\u3009",  # LEFT ANGLE BRACKET / RIGHT ANGLE BRACKET
    "\u300a": "\u300b",  # LEFT DOUBLE ANGLE BRACKET / RIGHT DOUBLE ANGLE BRACKET
    "\u300c": "\u300d",  # LEFT CORNER BRACKET / RIGHT CORNER BRACKET
    "\u300e": "\u300f",  # LEFT WHITE CORNER BRACKET / RIGHT WHITE CORNER BRACKET
    "\u3010": "\u3011",  # LEFT BLACK LENTICULAR BRACKET / RIGHT BLACK LENTICULAR BRACKET
    "\u3014": "\u3015",  # LEFT TORTOISE SHELL BRACKET / RIGHT TORTOISE SHELL BRACKET
    "\u3016": "\u3017",  # LEFT WHITE LENTICULAR BRACKET / RIGHT WHITE LENTICULAR BRACKET
    "\u3018": "\u3019",  # LEFT WHITE TORTOISE SHELL BRACKET / RIGHT WHITE TORTOISE SHELL BRACKET
    "\u301a": "\u301b",  # LEFT WHITE SQUARE BRACKET / RIGHT WHITE SQUARE BRACKET
    "\ufe59": "\ufe5a",  # SMALL LEFT PARENTHESIS / SMALL RIGHT PARENTHESIS
    "\ufe5b": "\ufe5c",  # SMALL LEFT CURLY BRACKET / SMALL RIGHT CURLY BRACKET
    "\ufe5d": "\ufe5e",  # SMALL LEFT TORTOISE SHELL BRACKET / SMALL RIGHT TORTOISE SHELL BRACKET
    "\uff08": "\uff09",  # FULLWIDTH LEFT PARENTHESIS / FULLWIDTH RIGHT PARENTHESIS
    "\uff3b": "\uff3d",  # FULLWIDTH LEFT SQUARE BRACKET / FULLWIDTH RIGHT SQUARE BRACKET
    "\uff5b": "\uff5d",  # FULLWIDTH LEFT CURLY BRACKET / FULLWIDTH RIGHT CURLY BRACKET
    "\uff5f": "\uff60",  # FULLWIDTH LEFT WHITE PARENTHESIS / FULLWIDTH RIGHT WHITE PARENTHESIS
    "\uff62": "\uff63",  # HALFWIDTH LEFT CORNER BRACKET / HALFWIDTH RIGHT CORNER BRACKET
}

# PropList.txt：Sentence_Terminal の範囲（両端を含む）
SENTENCE_TERMINAL_RANGES: tuple[tuple[int, int], ...] = (
    (0x0021, 0x0021),
    (0x002E, 0x002E),
    (0x003F, 0x003F),
    (0x0589, 0x0589),
    (0x061D, 0x061F),
    (0x06D4, 0x06D4),
    (0x0700, 0x0702),
    (0x07F9, 0x07F9),
    (0x0837, 0x0837),
    (0x0839, 0x0839),
    (0x083D, 0x083E),
    (0x0964, 0x0965),
    (0x104A, 0x104B),
    (0x1362, 0x1362),
    (0x1367, 0x1368),
    (0x166E, 0x166E),
    (0x1735, 0x1736),
    (0x1803, 0x1803),
    (0x1809, 0x1809),
    (0x1944, 0x1945),
    (0x1AA8, 0x1AAB),
    (0x1B5A, 0x1B5B),
    (0x1B5E, 0x1B5F),
    (0x1B7D, 0x1B7E),
    (0x1C3B, 0x1C3C),
    (0x1C7E, 0x1C7F),
    (0x203C, 0x203D),
    (0x2047, 0x2049),
    (0x2E2E, 0x2E2E),
    (0x2E3C, 0x2E3C),
    (0x2E53, 0x2E54),
    (0x3002, 0x3002),
    (0xA4FF, 0xA4FF),
    (0xA60E, 0xA60F),
    (0xA6F3, 0xA6F3),
    (0xA6F7, 0xA6F7),
    (0xA876, 0xA877),
    (0xA8CE, 0xA8CF),
    (0xA92F, 0xA92F),
    (0xA9C8, 0xA9C9),
    (0xAA5D, 0xAA5F),
    (0xAAF0, 0xAAF1),
    (0xABEB, 0xABEB),
    (0xFE52, 0xFE52),
    (0xFE56, 0xFE57),
    (0xFF01, 0xFF01),
    (0xFF0E, 0xFF0E),
    (0xFF1F, 0xFF1F),
    (0xFF61, 0xFF61),
    (0x10A56, 0x10A57),
    (0x10F55, 0x10F59),
    (0x10F86, 0x10F89),
    (0x11047, 0x11048),
    (0x110BE, 0x110C1),
    (0x11141, 0x11143),
    (0x111C5, 0x111C6),
    (0x111CD, 0x111CD),
    (0x111DE, 0x111DF),
    (0x11238, 0x11239),
    (0x1123B, 0x1123C),
    (0x112A9, 0x112A9),
    (0x1144B, 0x1144C),
    (0x115C2, 0x115C3),
    (0x115C9, 0x115D7),
    (0x11641, 0x11642),
    (0x1173C, 0x1173E),
    (0x11944, 0x11944),
    (0x11946, 0x11946),
    (0x11A42, 0x11A43),
    (0x11A9B, 0x11A9C),
    (0x11C41, 0x11C42),
    (0x11EF7, 0x11EF8),
    (0x16A6E, 0x16A6F),
    (0x16AF5, 0x16AF5),
    (0x16B37, 0x16B38),
    (0x16B44, 0x16B44),
    (0x16E98, 0x16E98),
    (0x1BC9F, 0x1BC9F),
    (0x1DA88, 0x1DA88),
)

# PropList.txt：Terminal_Punctuation の範囲（文・句の区切りの字。読点・コロン・セミコロンも含む）
TERMINAL_PUNCTUATION_RANGES: tuple[tuple[int, int], ...] = (
    (0x0021, 0x0021),
    (0x002C, 0x002C),
    (0x002E, 0x002E),
    (0x003A, 0x003B),
    (0x003F, 0x003F),
    (0x037E, 0x037E),
    (0x0387, 0x0387),
    (0x0589, 0x0589),
    (0x05C3, 0x05C3),
    (0x060C, 0x060C),
    (0x061B, 0x061B),
    (0x061D, 0x061F),
    (0x06D4, 0x06D4),
    (0x0700, 0x070A),
    (0x070C, 0x070C),
    (0x07F8, 0x07F9),
    (0x0830, 0x083E),
    (0x085E, 0x085E),
    (0x0964, 0x0965),
    (0x0E5A, 0x0E5B),
    (0x0F08, 0x0F08),
    (0x0F0D, 0x0F12),
    (0x104A, 0x104B),
    (0x1361, 0x1368),
    (0x166E, 0x166E),
    (0x16EB, 0x16ED),
    (0x1735, 0x1736),
    (0x17D4, 0x17D6),
    (0x17DA, 0x17DA),
    (0x1802, 0x1805),
    (0x1808, 0x1809),
    (0x1944, 0x1945),
    (0x1AA8, 0x1AAB),
    (0x1B5A, 0x1B5B),
    (0x1B5D, 0x1B5F),
    (0x1B7D, 0x1B7E),
    (0x1C3B, 0x1C3F),
    (0x1C7E, 0x1C7F),
    (0x203C, 0x203D),
    (0x2047, 0x2049),
    (0x2E2E, 0x2E2E),
    (0x2E3C, 0x2E3C),
    (0x2E41, 0x2E41),
    (0x2E4C, 0x2E4C),
    (0x2E4E, 0x2E4F),
    (0x2E53, 0x2E54),
    (0x3001, 0x3002),
    (0xA4FE, 0xA4FF),
    (0xA60D, 0xA60F),
    (0xA6F3, 0xA6F7),
    (0xA876, 0xA877),
    (0xA8CE, 0xA8CF),
    (0xA92F, 0xA92F),
    (0xA9C7, 0xA9C9),
    (0xAA5D, 0xAA5F),
    (0xAADF, 0xAADF),
    (0xAAF0, 0xAAF1),
    (0xABEB, 0xABEB),
    (0xFE50, 0xFE52),
    (0xFE54, 0xFE57),
    (0xFF01, 0xFF01),
    (0xFF0C, 0xFF0C),
    (0xFF0E, 0xFF0E),
    (0xFF1A, 0xFF1B),
    (0xFF1F, 0xFF1F),
    (0xFF61, 0xFF61),
    (0xFF64, 0xFF64),
    (0x1039F, 0x1039F),
    (0x103D0, 0x103D0),
    (0x10857, 0x10857),
    (0x1091F, 0x1091F),
    (0x10A56, 0x10A57),
    (0x10AF0, 0x10AF5),
    (0x10B3A, 0x10B3F),
    (0x10B99, 0x10B9C),
    (0x10F55, 0x10F59),
    (0x10F86, 0x10F89),
    (0x11047, 0x1104D),
    (0x110BE, 0x110C1),
    (0x11141, 0x11143),
    (0x111C5, 0x111C6),
    (0x111CD, 0x111CD),
    (0x111DE, 0x111DF),
    (0x11238, 0x1123C),
    (0x112A9, 0x112A9),
    (0x1144B, 0x1144D),
    (0x1145A, 0x1145B),
    (0x115C2, 0x115C5),
    (0x115C9, 0x115D7),
    (0x11641, 0x11642),
    (0x1173C, 0x1173E),
    (0x11944, 0x11944),
    (0x11946, 0x11946),
    (0x11A42, 0x11A43),
    (0x11A9B, 0x11A9C),
    (0x11AA1, 0x11AA2),
    (0x11C41, 0x11C43),
    (0x11C71, 0x11C71),
    (0x11EF7, 0x11EF8),
    (0x12470, 0x12474),
    (0x16A6E, 0x16A6F),
    (0x16AF5, 0x16AF5),
    (0x16B37, 0x16B39),
    (0x16B44, 0x16B44),
    (0x16E97, 0x16E98),
    (0x1BC9F, 0x1BC9F),
    (0x1DA87, 0x1DA8A),
)

# 一般カテゴリ Po で、名前に語「COLON」を含む字（SEMICOLON は含めない）
LABEL_COLONS = frozenset(
    "\u003a\u0703\u0704\u0705\u0706\u0707\u0708\u0709\u1365\u1366\u1804\ua6f4\ufe13\ufe55\uff1a\U00012471\U00012472\U0001da8a"
)
# ---- ここまで scripts/make_chars.py が作る

_TERM_STARTS = [a for a, _b in SENTENCE_TERMINAL_RANGES]
_TP_STARTS = [a for a, _b in TERMINAL_PUNCTUATION_RANGES]
