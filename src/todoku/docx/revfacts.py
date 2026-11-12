"""変更履歴の事実（まだ承諾も却下もされていない挿入・削除を、だれが・いつ）：窓に事実としてだけ出す印。

字の片（Seg）の rev は (種類 "ins"・"del", 変更した人, 日時)。人と日時は、ファイルに書いてある値のまま（日時は、時刻の帯を直さない）。
削除された字は本文に入れず、段落の後ろに〔変更履歴で削除された字（…）: 字〕。挿入された字は本文の字として読み、段落の字がみな同じ
挿入なら段落の前に〔変更履歴で挿入された段落（…）〕、一部だけなら段落の後ろに〔うち変更履歴で挿入された字（…）：「字」〕。
段落の区切り（段落の終わりの印）だけの挿入・削除は段落の後ろに、表の行の挿入・削除は行の頭の段落の前に示す（.docx）。"""

from __future__ import annotations


def dttm(v: int) -> str:
    """Word の DTTM（MS-DOC 2.9.65。RTF の \\revdttm も同じ）：分 6 ビット・時 5・日 5・月 4・年 9（1900 から）・曜日 3。
    日か月が 0（日時なし）と、暦にない値は ""。時刻の帯は書いていないので、書いてある数のまま「年-月-日 時:分」にする。"""
    v &= 0xFFFFFFFF
    mi, hr, day, mon, yr = v & 0x3F, (v >> 6) & 0x1F, (v >> 11) & 0x1F, (v >> 16) & 0xF, (v >> 20) & 0x1FF
    if mi > 59 or hr > 23 or not 1 <= day <= 31 or not 1 <= mon <= 12:  # 日・月が 0 の DTTM は、決まりで「ない」
        return ""
    return f"{1900 + yr:04d}-{mon:02d}-{day:02d} {hr:02d}:{mi:02d}"


def who(rev) -> str:
    """「変更した人：…・日時：…」（ファイルに書いていないものは出さない）。"""
    if rev is None:
        return ""
    a, d = (rev[1] or "").strip(), (rev[2] or "").strip()
    return "・".join(x for x in (f"変更した人：{a}" if a else "", f"日時：{d}" if d else "") if x)


def inserted_text(p) -> str:
    """段落の、変更履歴で挿入された字（本文の字として読んでいる字）。"""
    return "".join(s.text for s in p.segs if not s.hidden and s.kind != "del" and _ins(s))


def tracked_items(numbered) -> tuple[list[str], list[str], list[str]]:
    """（段落の番号, 段落）の並びから、変更履歴で削除された字の残る段落、挿入された字のある段落（番号と、字の頭）、
    挿入・削除された段落の区切り・表の行（窓に印を出すもの）の一覧。"""
    gone, new, marks = [], [], []
    for n, q in numbered:
        if hasattr(q, "deleted_text") and q.deleted_text().strip():
            gone.append(f"段落 {n}「{q.deleted_text().strip()[:20].replace(chr(10), '⏎')}」")
        t = inserted_text(q)
        if t:
            new.append(f"段落 {n}「{t.strip()[:20].replace(chr(10), ' ')}」")
        row = getattr(q, "row_rev", None)
        if row is not None:
            marks.append(f"段落 {n}：{_DONE[row[0]]}表の行")
        if break_mark(q):
            marks.append(f"段落 {n}：{_DONE[q.mark_rev[0]]}段落の区切り")
    return gone, new, marks


_DONE = {"ins": "挿入された", "del": "削除された"}
_STATE = {"ins": "まだ承諾も却下もされていない", "del": "いまの本文ではない"}


def row_mark(p) -> str:
    """段落の前に置く印：この段落が頭にある表の行の、変更履歴での挿入・削除（w:trPr の中の印）。"""
    rev = getattr(p, "row_rev", None)
    if rev is None:
        return ""
    w = who(rev)
    return f"〔変更履歴で{_DONE[rev[0]]}表の行（{_STATE[rev[0]]}" + (f"。{w}" if w else "") + "）〕"


def break_mark(p) -> str:
    """段落の後ろに置く印：段落の区切り（段落の終わりの印）の、変更履歴での挿入・削除。段落の字がみな同じ変更（段落ごとの挿入・削除）なら、
    字の印が伝えているので置かない。"""
    rev = getattr(p, "mark_rev", None)
    if rev is None:
        return ""
    if rev[0] == "ins" and inserted_marks(p)[0]:
        return ""
    if rev[0] == "del" and not p.text.strip() and p.deleted_text().strip():
        return ""
    w = who(rev)
    if rev[0] == "ins":
        return "〔変更履歴で挿入された段落の区切り（まだ承諾も却下もされていない" + (f"。{w}" if w else "") + "）〕"
    return "〔変更履歴で削除された段落の区切り（いまの本文では、次の段落に続く" + (f"。{w}" if w else "") + "）〕"


def _ins(s) -> bool:
    rev = getattr(s, "rev", None)
    return rev is not None and rev[0] == "ins"


def deleted_mark(p) -> str:
    """段落の後ろに置く、削除された字の印。変更した人・日時の違う削除は、分けて示す。"""
    groups: list[list] = []
    for s in p.segs:
        if s.kind != "del":
            continue
        w = who(getattr(s, "rev", None))
        if groups and groups[-1][1] == w:
            groups[-1][0] += s.text
        else:
            groups.append([s.text, w])
    if len({w for _t, w in groups}) == 1:
        groups = [["".join(t for t, _w in groups), groups[0][1]]]
    out = []
    for t, w in groups:
        t = t.strip()
        if t:
            # 段落をまたいで削除された所（OpenDocument）は、消えた段落の切れ目を ⏎ で示す
            out.append("〔変更履歴で削除された字（いまの本文ではない" + ("。⏎ は段落の切れ目" if "\n" in t else "")
                       + (f"。{w}" if w else "") + f"）: {t.replace(chr(10), '⏎')}〕")
    return "".join(out)


def inserted_marks(p) -> tuple[str, str]:
    """(段落の前に置く印, 段落の後ろに置く印)：変更履歴で挿入された字（まだ承諾も却下もされていない）。"""
    runs: list[list] = []
    for s in p.segs:
        if s.hidden or s.kind == "del" or not s.text:
            continue
        ins = _ins(s)
        w = who(s.rev) if ins else ""
        if runs and runs[-1][2] == ins and runs[-1][1] == w:
            runs[-1][0] += s.text
        else:
            runs.append([s.text, w, ins])
    got = [r for r in runs if r[2]]
    if not got:
        return "", ""
    shown = [r for r in runs if r[0].strip()]
    if shown and all(r[2] for r in shown) and len({r[1] for r in shown}) == 1:
        w = shown[0][1]
        return "〔変更履歴で挿入された段落（まだ承諾も却下もされていない" + (f"。{w}" if w else "") + "）〕", ""
    return "", "".join("〔うち変更履歴で挿入された字（まだ承諾も却下もされていない" + (f"。{w}" if w else "")
                       + f"）：「{t.replace(chr(10), ' ').strip() or t.replace(chr(10), ' ')}」〕" for t, w, _i in got)
