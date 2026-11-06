#!/usr/bin/env python3
"""読む比べ（採点 v2。scripts/eval/read_compare.py の写しに、確かめた直しだけを当てたもの。とどくにも比べる道具にも同じに効く）。
v1 からの違い（理由と試験は docs/再評価の計画.md・tests/test_eval_v2.py）：
  R1 exact_read：2 字の並びの再現率 0.0 を「なし」と同じに通していた（(bigram or 1)）。0.0 は落ちる。
  R2 環境で字が変わる組み込みの日付（truth.py の T5）は、日付の並びのどれか（年月日・日月年・月日年、区切り / - . 、年月日の字）を画面の字として認める。
  R3 PDF：文字認識の字数に比べて字の層の字数が半分より少ない書類（先頭 3 ページ）は、絵の中の字を読めていないので、どの道具も exact_read にしない（image_text）。
     字の層だけで決めた値も exact_text_layer に残す（感度の比べ）。
  R4 書類ごとに別の処理で採点し、時間の上限を超えたら「未採点」（unscored）と記録する。大量の失敗で止まる。失敗は scores に書かない（次の回でやり直す）。
  R6 感度の比べだけ（exact_value）：表計算で元の値を正しく出した道具も通す数を別に残す。「新しく届く」が画面の書式の差だけでないかを見るため。
  R7 ひとつ選ぶ欄の印 ◉（選んだ）・○（選んでいない）も状態の印に数える。比べる道具の ☑☐ と同じ扱い。
  R8 窓の復号：マスの区切りは「 ｜ 」（空白つき）だけを消す。書類の字の ｜ は残す。［　］（規則の空欄）は中の空白に戻す。
     フォームの欄の印（☐［Check14］など）は字の比べの前に除く（書類の字でなく、語の中に入ると 2 字の並びが切れる）。位置で決まる所の前の ☐ は印刷の字なので残す。
  R9 Word の字の色：文書でいちばん多い色を除いた項目を color_not_normal として別に数える（定義は変えない。黒か文書のふつうの色かは決まっていない）。
  R5 とどくの src を --todoku-src で渡す（道具の版と採点の版を分ける）。Office は動かさない。

読む比べ：ファイルの一覧を、比べる道具ととどくで読み、字の再現率だけでなく、特徴ごとに数える。

特徴（正解は truth.py・作りは probes.py。どちらもとどくを使わない）：
  displayed    画面の字（日付・和暦・数の形式）。セルごとに、画面の字が値の境目つきで出力にあるか
  formula      数式の値（再計算した値、または再計算と同じ保存された値）が出力にあるか
  stale_mut    保存された値をわざと変えた写し：出力が、変えた値を印なしで出すか（古い値の見分け。作った試験で、とどくの得意な形）
  merged_rows  縦の結合の値が、結合の 2 行目以降の行まで届くか（その行の字と同じ行に出るか。窓がまとめた行は言明のとおりに広げて見る）
  hidden       非表示のシート・行・列の字に印があるか（hidden_leak は印なしで見える字として出した数）
  strike/color/furigana  取り消し線・字の色の印、ふりがなの字
  shapes       図形・テキストボックスの字
  checkboxes   チェックボックスの状態の印（☑ と ☐ の数）
  form_values  PDF のフォームの欄の値
  scan_flag    字のないページがあることを伝えるか
PDF の字の正解は pdftotext -layout（主）・Vision の文字認識（あれば）・とどく自身の窓（副。とどくに有利。計画の指定なので残す）。

正確に読めた（exact_read。world_counts.py の「同じ定義」もこれ）：開けて、字がある書類で、字と 2 字の並びの再現率が 0.995 以上（前の比べで使った
値を、ここで調整せずに使う）、画面の字と数式の値も 0.995 以上。

使い方: read_compare.py --files 一覧 --out 出力 --baseline-python PY --todoku-python PY --todoku-src SRC [--label dev] [--pkr PATH] [--purge] [--meta JSON] [--doc-limit-s 900]
途中で止めても、もう一度動かせば続きから（scores/<id>.json がある書類は飛ばす）。"""

from __future__ import annotations

import argparse
import collections
import datetime
import hashlib
import io
import json
import random
import re
import shutil
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import truth as T  # noqa: E402
from common import SCORER_VERSION, atomic_json, doc_id, drive_docs, guard_path, kind_of, last_json, load_json, md_table, norm, norm_sp, pct, read_list, run_watched  # noqa: E402

EXACT = 0.995
IMAGE_TEXT_RATIO = 0.5  # 字の層の字数が文字認識の字数のこの割合より少なければ、字の多くは絵の中（R3）
TOOLS = {"xlsx": ("markitdown", "openpyxl"), "xls": ("markitdown", "xlrd"), "docx": ("markitdown", "python-docx"), "doc": ("markitdown", "python-docx"),
         "pdf": ("markitdown", "pdfplumber", "pypdf"), "csv": ("markitdown", "charset"), "txt": ("markitdown", "charset"),
         "odt": ("markitdown",), "ods": ("markitdown",), "rtf": ("markitdown",)}
HID = re.compile(r"非表示|hidden", re.I)
STRIKE = re.compile(r"取消線|取り消し線|消し線|~~|<s>|<del>|strike", re.I)
COLOR = re.compile(r"色[:：]|字の色|赤字|赤い字|color|#[0-9a-f]{6}", re.I)
CHECK_ON = re.compile(r"☑|☒|✓|✔|◉|\[x\]", re.I)
CHECK_OFF = re.compile(r"☐|○|\[ \]")
SCAN = re.compile(r"画像|image|スキャン|scan", re.I)
STALE = re.compile(r"保存された値と合わない|古い値|stale|文字の並び順で結果が変わる", re.I)
VOLATILE = re.compile(r"\b(TODAY|NOW|RAND|RANDBETWEEN|OFFSET|INDIRECT|CELL|INFO)\s*\(", re.I)
_TODOKU_META = ("### 入力欄", "## 入力欄", "### 規則では決められない所", "## 規則では決められない所", "### 画像", "## この窓の目次", "### ふりがな")
_TODOKU_NOTE = ("- 非表示", "- シート保護", "- 印刷範囲", "- 文字に直せない字", "- フォームの欄はありません", "- 書き方：", "- 行は「", "- **文字コード", "- 名前「", "- 文字の並び順で")


def todoku_content(md: str) -> str:
    """とどくの窓から、とどくが足した説明（冒頭・入力欄の一覧など）を除いた中身。再現率を説明の字で水増ししないため（world_compare.py と同じ切り方）。"""
    out, skip, started = [], False, False
    for ln in md.splitlines():
        if ln.startswith("#"):
            started = started or ln.startswith("## ")
            skip = ln.startswith(_TODOKU_META)
            # シートの見出しは本文に入れない（シート名の字で再現率が上がらないように）
            continue
        if skip and ln.startswith("|") and ln.count("|") >= 5 and not ln.startswith(("|---", "| 欄 |")):
            # 「規則では決められない所」の表の最後の列（いまの値）：読み順で「［？］」や「同じ形の行」にまとめた字は、ここに元のまま出る
            out.append(ln.rstrip().rstrip("|").rsplit("|", 1)[-1])
            continue
        if not started or skip or ln.startswith(_TODOKU_NOTE):
            continue
        out.append(ln)
    return "\n".join(out)


_ID = r"(?:▶)?[^［］〈〉〔〕＝？\s]*?(?:\d|\})"
_COLLAPSED = re.compile(r"^- (\d+)〜(\d+): (\d+)(?:〜(\d+))?行目と同じ形の(?:\d+行1組|行)が \d+ (?:行|組)続きます（(.*)$")
_ROW_LINE = re.compile(r"^- (\d+): (.*)$")


def expand_collapsed(md: str) -> str:
    """窓が読み順で「同じ形の行が続く」とまとめた所を、窓の言明のとおりに行へ広げる。言明が違っていれば（字の違う行があれば）、
    その字は正解と合わないので、再現率に落ち度として出る。広げないときの数も char_literal に残す。"""
    out, rows, last = [], {}, None
    for ln in md.splitlines():
        if ln.startswith("#"):
            rows = {} if ln.startswith("## ") else rows
            last = None
        m = _ROW_LINE.match(ln)
        c = _COLLAPSED.match(ln)
        if m:
            rows[int(m[1])] = last = int(m[1])
            rows[last] = m[2]
        elif c:
            a, b, r0, r1 = int(c[1]), int(c[2]), int(c[3]), int(c[4] or c[3])
            per = r1 - r0 + 1
            for x in range(a, b + 1):
                t = rows.get(r0 + (x - a) % per)
                if t:
                    out.append(f"- {x}: {t}")
            vals = re.findall(r"[A-Z]+\d+＝([^、）]*)", c[5])  # 字の違う列の値
            out.extend(f"- {v}" for v in vals)
            last = None
            continue
        elif last is not None and ln.strip() and not ln.startswith("- "):
            rows[last] += " " + ln  # マスの字の改行で続く行
        elif ln.startswith("- "):
            last = None
        out.append(ln)
    return "\n".join(out)


_FONT_MARK = re.compile(r"〔(?:取り消し線|字の色 #[0-9A-Fa-f]{6}|・)+〕")
_FONT_PART = re.compile(r"〔うち[^〕]*〕")


def drop_font_marks(t: str) -> str:
    """字の書式の印（窓の説明の字）を、字の再現率の比べから除く。印の確かめは Hay で別に見る。"""
    return _FONT_PART.sub("", _FONT_MARK.sub("", t))


_CTRL_MARK = re.compile(r"[☐☑☒○◉](?=\s*［(?!P\d+\.c|p\d+\.[a-z]))")


def drop_ctrl_marks(md: str) -> str:
    """フォームの欄の印を、字の比べの前に除く（R8）。状態の印の数え（checks）は別に元の字で見る。"""
    return _CTRL_MARK.sub("", md)


def clean_todoku(md: str) -> str:
    """窓の記号（欄の名前・番地・「？」の空き・余白の印）を除いて、書類の字に戻す。「［B5？　］」は元の字（空白・点線など）、「［B5＝値］」は値、「〔A2〕」は消す。"""
    t = re.sub(r"［(?:" + _ID + r")?？([^］＝]*)］", r"\1", md)
    t = re.sub(r"［[↓↑←→]余白[^］]*］", " ", t)
    t = re.sub(r"［" + _ID + r"？?＝(.*?)？?］", r"\1", t)
    t = re.sub(r"〈" + _ID + r"＝(.*?)〉", r"\1", t)
    t = re.sub(r"［" + _ID + r"？?(?:：[^］]*)?］", " ", t)
    t = re.sub(r"〔" + _ID + r"〕", "", t)
    t = re.sub(r"^\s*-\s+(?:\d+(?:〜\d+)?:\s)?", "", t, flags=re.M)
    t = re.sub(r"［([\s　]*)］", r"\1", t)  # 規則の空欄
    t = re.sub(r"(?:^|(?<=\s))｜(?=\s|$)", " ", t, flags=re.M)  # マスの区切りだけ。書類の字の ｜ は残す
    return t.replace("↵", " ")


def recalls(units: list[str], out: str) -> dict:
    u = [norm(x) for x in units]
    ref = collections.Counter("".join(u))
    got = collections.Counter(norm(out))
    tot = sum(ref.values())
    # 2 字の並び：正解は単位の中だけ（表のマス・段落・PDF の字のかたまり）。出力は改行をまたいで数える（行の分け方の違いを、読み落としに数えない）
    a = collections.Counter(x[k:k + 2] for x in u for k in range(len(x) - 1))
    t = norm(out)
    b = collections.Counter(t[k:k + 2] for k in range(len(t) - 1))
    btot = sum(a.values())
    return {"char": round(1 - sum((ref - got).values()) / tot, 5) if tot else None,
            "bigram": round(1 - sum((a - b).values()) / btot, 5) if btot else None, "ref_chars": tot}


_DATE_NUM = re.compile(r"(?<![0-9.,])(\d{1,4})[-/.](\d{1,2})[-/.](\d{1,4})(?![0-9])")
_DATE_JA = re.compile(r"(?<![0-9.,])(\d{4})年(\d{1,2})月(\d{1,2})日")


def serial_date(raw) -> tuple[int, int, int] | None:
    """ファイルの日付の数（1900 年の系）→ (年の下 2 桁, 月, 日)。"""
    try:
        v = float(raw)
    except (TypeError, ValueError):
        return None
    if not 1 <= v < 2_958_466:
        return None
    x = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
    return x.year % 100, x.month, x.day


class Hay:
    """出力を比べやすくしたもの。"""

    def __init__(self, text: str):
        self.lines = [norm_sp(x) for x in text.splitlines() if x.strip()]
        self.flat = "\n".join(self.lines)
        self.nospace = norm(text)
        self._n = {}
        self._dates = None

    def date_count(self, y2: int, m: int, d: int) -> int:
        """その日付を、どの環境の並びでもよいとして数える（R2）。"""
        if self._dates is None:
            c = collections.Counter()
            for mt in _DATE_NUM.finditer(self.flat):
                a, b, e = mt[1], int(mt[2]), mt[3]
                keys = set()
                if len(a) == 4:  # 年月日
                    keys.add((int(a) % 100, b, int(e)))
                else:
                    keys.update({(int(e) % 100, int(a), b), (int(e) % 100, b, int(a))})  # 月日年・日月年
                    if len(a) == 2 and len(e) <= 2:
                        keys.add((int(a), b, int(e)))  # 年の下 2 桁が先
                for k in keys:
                    c[k] += 1
            for mt in _DATE_JA.finditer(self.flat):
                c[(int(mt[1]) % 100, int(mt[2]), int(mt[3]))] += 1
            self._dates = c
        return self._dates[(y2, m, d)]

    def count(self, tok: str) -> int:
        """値の境目つきの出現の数（「12」が「123」の中に当たらないように）。"""
        if tok in self._n:
            return self._n[tok]
        pat = re.compile(r"(?<![0-9.,])" + re.escape(tok) + r"(?![0-9])")
        n = len(pat.findall(self.flat))
        self._n[tok] = n
        return n

    def lines_with(self, tok: str) -> list[str]:
        return [ln for ln in self.lines if tok in ln]


def canon_num(raw) -> str | None:
    try:
        v = float(raw)
    except (TypeError, ValueError):
        return None
    return str(int(v)) if v.is_integer() and abs(v) < 1e15 else str(v)


def sheet_maps(t: dict):
    """(シートの作り, 画面の字, 再計算の字) を順に。"""
    disp = {s["name"]: s for s in t.get("sheets", [])}
    for sp in (t.get("probes") or {}).get("sheets", []):
        d = disp.get(sp["name"]) or {"display": {}, "fresh": {}}
        yield sp, d["display"], d.get("fresh") or d["display"]


def rc(addr: str) -> tuple[int, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", addr)
    return int(m[2]), T.col_num(m[1])


def expand(ref: str):
    a, _, b = ref.partition(":")
    r0, c0 = rc(a)
    r1, c1 = rc(b or a)
    return r0, c0, r1, c1


def score_sheet(t: dict, hay: Hay, raw_text: str) -> dict:
    cat: dict = {}
    disp_cnt, disp_kind = collections.Counter(), {}
    loc_cnt, loc_kind, loc_disp = collections.Counter(), {}, {}
    fo_n = fo_ok = st_n = st_ok = 0
    merged_n = merged_ok = 0
    hid_n = hid_ok = leak = 0
    sh_n = sh_ok = 0
    strike_n = strike_ok = color_n = color_ok = ph_n = ph_ok = 0
    for sp, disp, fresh in sheet_maps(t):
        cells = sp.get("cells", {})
        state = sp.get("state", "visible")
        if state != "visible":
            sh_n += 1
            nm = norm_sp(sp["name"])
            if any(nm in ln and HID.search(ln) for ln in hay.lines) or re.search(r"非表示[^\n]*" + re.escape(nm), hay.flat):
                sh_ok += 1
        hr = set(sp.get("hidden_rows") or [])
        hc = set(sp.get("hidden_cols") or [])
        visible_vals = set()
        for addr, c in cells.items():
            r, col = rc(addr)
            if r not in hr and col not in hc:
                visible_vals.add(norm_sp(disp.get(addr) or c.get("raw") or ""))
        by_row = collections.defaultdict(dict)
        for addr, txt in disp.items():
            r, col = rc(addr)
            by_row[r][col] = txt
        for addr, c in cells.items():
            d = norm_sp(disp.get(addr, ""))
            ymd = serial_date(c.get("raw")) if T.is_locale_date(c) else None
            if c.get("fk") and d and not re.fullmatch(r"#+", d):
                if ymd:
                    loc_cnt[ymd] += 1
                    loc_kind.setdefault(ymd, c["fk"])
                    loc_disp.setdefault(ymd, d)
                else:
                    disp_cnt[d] += 1
                    disp_kind.setdefault(d, c["fk"])
            if "f" in c:
                fv = norm_sp(fresh.get(addr, disp.get(addr, "")))
                if fv and not re.fullmatch(r"#+", fv):
                    fo_n += 1
                    rawc = canon_num(c.get("raw"))
                    same = norm_sp(disp.get(addr, "")) == fv
                    if hay.count(fv) or (ymd and hay.date_count(*ymd)) or (same and rawc and hay.count(rawc)) or (same and c.get("raw") and hay.count(norm_sp(str(c["raw"])))):
                        fo_ok += 1
                    if not same and not VOLATILE.search(c.get("f") or ""):
                        st_n += 1
                        if hay.count(fv) or STALE.search(raw_text):
                            st_ok += 1
            val = norm_sp(disp.get(addr) or c.get("raw") or "")
            if not val:
                continue
            r, col = rc(addr)
            if (r in hr or col in hc) and state == "visible" and len(norm(val)) >= 2:
                hid_n += 1
                ls = hay.lines_with(val)
                if any(HID.search(ln) for ln in ls):
                    hid_ok += 1
                elif ls and val not in visible_vals:
                    leak += 1
            if c.get("strike"):
                strike_n += 1
                strike_ok += any(STRIKE.search(ln) for ln in hay.lines_with(val))
            if c.get("color"):
                color_n += 1
                color_ok += any(COLOR.search(ln) for ln in hay.lines_with(val))
            if c.get("ph") and norm(c["ph"]) not in norm(val):
                ph_n += 1
                # ふりがなのかたまりごと（probes が \x01 で区切る）に出力にあるか
                ph_ok += all(norm(x) in hay.nospace for x in c["ph"].split("\x01") if norm(x))
        checked = 0
        for ref in sp.get("merges") or []:
            if checked >= 300:
                break
            r0, c0, r1, c1 = expand(ref)
            if r1 == r0:
                continue
            tl = norm_sp(disp.get(f"{T.col_name(c0)}{r0}", ""))
            if len(norm(tl)) < 1:
                continue
            for r in range(r0 + 1, min(r1, r0 + 200) + 1):
                others = [norm_sp(v) for cc, v in by_row.get(r, {}).items() if not c0 <= cc <= c1 and len(norm(v)) >= 2][:3]
                if not others:
                    continue
                checked += 1
                merged_n += 1
                ls = [ln for ln in hay.lines if all(o in ln for o in others)]
                merged_ok += any(tl in ln for ln in ls)
    pr = t.get("probes") or {}
    disp_tot = collections.Counter()
    disp_got = collections.Counter()
    items = list(disp_cnt.items())
    if len(items) > 4000:
        random.Random(20261105).shuffle(items)
        items = items[:4000]
    for d, n in items:
        k = disp_kind[d]
        disp_tot[k] += n
        disp_got[k] += min(n, hay.count(d))
    for ymd, n in loc_cnt.items():
        k = loc_kind[ymd]
        disp_tot[k] += n
        disp_got[k] += min(n, max(hay.date_count(*ymd), hay.count(loc_disp[ymd])))
    for k in disp_tot:
        cat[f"displayed_{k}"] = {"n": disp_tot[k], "ok": disp_got[k]}
    cat["displayed"] = {"n": sum(disp_tot.values()), "ok": sum(disp_got.values())}
    cat["formula"] = {"n": fo_n, "ok": fo_ok}
    cat["stale_natural"] = {"n": st_n, "ok": st_ok}
    cat["merged_rows"] = {"n": merged_n, "ok": merged_ok}
    cat["hidden_sheets"] = {"n": sh_n, "ok": sh_ok}
    cat["hidden_cells"] = {"n": hid_n, "ok": hid_ok, "leak": leak}
    cat["strike"] = {"n": strike_n, "ok": strike_ok}
    cat["color"] = {"n": color_n, "ok": color_ok}
    cat["furigana"] = {"n": ph_n, "ok": ph_ok}
    shapes = pr.get("shapes") or []
    cat["shapes"] = {"n": len(shapes), "ok": sum(1 for s in shapes if norm(s["text"]) in hay.nospace)}
    cat["checkboxes"] = checks(pr.get("checkboxes") or [], raw_text)
    return cat


def checks(items: list, raw_text: str) -> dict:
    on = sum(1 for c in items if c.get("checked"))
    off = len(items) - on
    return {"n": len(items), "ok": min(on, len(CHECK_ON.findall(raw_text))) + min(off, len(CHECK_OFF.findall(raw_text)))}


def score_docx(t: dict, hay: Hay, raw_text: str) -> dict:
    pr = t.get("probes") or {}

    def lined(items, rx):
        return {"n": len(items), "ok": sum(1 for x in items if any(rx.search(ln) for ln in hay.lines_with(norm_sp(x))))}

    return {"shapes": {"n": len(pr.get("textboxes") or []), "ok": sum(1 for s in pr.get("textboxes") or [] if norm(s) in hay.nospace)},
            "checkboxes": checks(pr.get("checkboxes") or [], raw_text),
            "strike": lined(pr.get("strike") or [], STRIKE), "color": lined(pr.get("color") or [], COLOR),
            "color_not_normal": lined([x for x, v in zip(pr.get("color") or [], pr.get("color_vals") or []) if v != pr.get("color_normal")], COLOR),
            "hidden_cells": {**lined(pr.get("hidden") or [], HID), "leak": 0},
            "furigana": {"n": len(pr.get("ruby") or []), "ok": sum(1 for s in pr.get("ruby") or [] if norm(s) in hay.nospace)}}


def score_pdf(t: dict, hay: Hay, raw_text: str) -> dict:
    pr = t.get("probes") or {}
    empty = [c for c in (t.get("pages_text_chars") or [])[: pr.get("pages") or 0] if c == 0]
    return {"form_values": {"n": len(pr.get("fields") or []), "ok": sum(1 for f in pr.get("fields") or [] if norm(f["value"]) in hay.nospace)},
            "checkboxes": checks(pr.get("checkboxes") or [], raw_text),
            "scan_flag": {"n": 1 if empty else 0, "ok": 1 if empty and SCAN.search(raw_text) else 0}}


_PDF_SPLIT = re.compile(r"[\s\u200e\u200f\u202a-\u202e\u2066-\u2069]+")


def _recall_units(t: dict) -> list:
    """字の比べの単位。PDF の pdftotext -layout は、離れた列の字を空白 1 つでつなぐ（縦書きの「受給資格」と隣の「所持」が「受 所」になる）ので、
    空白をまたぐ 2 字の並びは、本当に並んだ字とは限らない。どの道具にも同じに、空白で切って数える。"""
    units = t.get("units") or []
    # v2：向きの制御の字でも切る。pdftotext が向きの違う並びの境目に足す印で、その両側が書類の中で続いているとは限らない（並びの中の順はこれまでどおり数える）
    return [w for u in units for w in _PDF_SPLIT.split(u) if w] if t.get("kind") == "pdf" else units


def value_units(t: dict) -> list:
    """感度の比べ（R6）：表計算の、書式を当てる前の値の単位。数は str(float) の形（Python のライブラリの出し方）、日付の書式のマスは除く（値の形が道具で違う）。"""
    out = []
    for sp, disp, _fresh in sheet_maps(t):
        if sp.get("state", "visible") != "visible":
            continue
        cells = sp.get("cells", {})
        for addr, v in disp.items():
            c = cells.get(addr) or {}
            if c.get("fk") in ("date", "era"):
                continue
            if c.get("fk") == "number":
                out.append(canon_num(c.get("raw")) or v)
            else:
                out.append(v)
    return out


def image_text(t: dict) -> bool:
    """PDF の字の多くが絵の中か（R3）。字の層は字を描く命令だけを表す（ISO 32000）。文字認識の字数は、字を読み違えても数としては使える（言語によらない）。"""
    oc = t.get("ocr_units")
    if t.get("kind") != "pdf" or oc is None:
        return False
    seen = len(norm("".join(oc)))
    return seen > 0 and sum((t.get("pages_text_chars") or [])[:T.OCR_PAGES]) < IMAGE_TEXT_RATIO * seen


def score_tool(t: dict, out_text: str, tool: str, res: dict) -> dict:
    literal = None
    chars_out = len(out_text)
    if tool == "todoku":
        literal = recalls(_recall_units(t), drop_font_marks(clean_todoku(drop_ctrl_marks(todoku_content(out_text)))))
        out_text = expand_collapsed(out_text)
        content = drop_font_marks(clean_todoku(drop_ctrl_marks(todoku_content(out_text))))
        out_text = clean_todoku(out_text)
    else:
        content = out_text
    r = recalls(_recall_units(t), content)
    s = {"tool": tool, "opened": "error" not in res, "error": res.get("error"), "seconds": res.get("seconds"), **r, "chars_out": chars_out}
    if literal is not None:
        s["char_literal"], s["bigram_literal"] = literal["char"], literal["bigram"]
    if not s["opened"]:
        return s
    hay = Hay(out_text)
    k = t["kind"]
    if k in ("xlsx", "xls"):
        s["cat"] = score_sheet(t, hay, out_text)
    elif k in ("docx", "doc"):
        s["cat"] = score_docx(t, hay, out_text)
    elif k == "pdf":
        s["cat"] = score_pdf(t, hay, out_text)
        if t.get("ocr_units"):
            s["ocr"] = recalls(t["ocr_units"], content)
    else:
        s["cat"] = {}
    s["exact_text_layer"] = exact_read(s)
    s["image_text"] = image_text(t)
    s["exact_read"] = s["exact_text_layer"] and not s["image_text"]
    # R6（公開の定義には入れない）：表計算で、書式を当てた字でなく元の値を正しく出した道具も通すとどうなるか
    s["exact_value"] = s["exact_read"]
    if k in ("xlsx", "xls") and not s["exact_read"]:
        v = recalls(value_units(t), content)
        s["exact_value"] = bool(s.get("ref_chars")) and (v["char"] or 0) >= EXACT and (v["bigram"] is None or v["bigram"] >= EXACT)
    return s


def exact_read(s: dict) -> bool:
    if not s.get("opened") or not s.get("ref_chars"):
        return False  # 字のない書類（スキャン）は、どの道具にとっても「読めた」にしない
    # 2 字の並びがない（1 字だけの書類）ときの None だけを通す。0.0 は落ちる
    if (s.get("char") or 0) < EXACT or (s.get("bigram") is not None and s["bigram"] < EXACT):
        return False
    for key in ("displayed", "formula"):
        c = (s.get("cat") or {}).get(key)
        if c and c["n"] and c["ok"] / c["n"] < EXACT:
            return False
    return True


# ---------------------------------------------------------------------- 保存された値をわざと変えた写し
def mutate_cached(src: Path, dst: Path, did: str, max_cells: int = 10) -> list[dict]:
    """数式のセルの保存された値（整数のもの）を、別の値に変える。[{part, addr, old, new}]。"""
    rng = random.Random(hashlib.sha256(did.encode()).hexdigest())
    z = zipfile.ZipFile(src)
    pat = re.compile(rb'(<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>\s*<f\b[^>]*?(?:/>|>.*?</f>)\s*<v>)(-?\d+(?:\.0+)?)(</v>)', re.S)
    cand = []
    parts = {}
    for n in z.namelist():
        if re.fullmatch(r"xl/worksheets/[^/]+\.xml", n):
            b = z.read(n)
            parts[n] = b
            for m in pat.finditer(b):
                if b' t="' not in m[1].split(b">")[0] or b' t="n"' in m[1].split(b">")[0]:
                    cand.append((n, m[2].decode(), m[3].decode()))
    if not cand:
        return []
    rng.shuffle(cand)
    pick = cand[:max_cells]
    out = []
    want = collections.defaultdict(dict)
    for n, a, old in pick:
        new = int(float(old)) * 3 + 7919
        want[n][a] = new
        out.append({"part": n, "addr": a, "old": old, "new": str(new)})
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            data = z.read(it.filename)
            if it.filename in want:
                w = want[it.filename]
                data = pat.sub(lambda m: m[1] + str(w[m[2].decode()]).encode() + m[4] if m[2].decode() in w else m[0], data)
            zo.writestr(it, data)
    dst.write_bytes(buf.getvalue())
    return out


def score_mutated(muts: list[dict], text: str, res: dict) -> dict:
    hay = Hay(text)
    wrong = 0
    for m in muts:
        tok = norm_sp(m["new"])
        tok2 = f"{int(m['new']):,}"
        if hay.count(tok) or hay.count(tok2):
            wrong += 1
    flagged = bool(STALE.search(text)) or any(STALE.search(r) for r in res.get("grade_reasons") or [])
    return {"n": len(muts), "wrong_shown": wrong, "flagged": flagged, "silent": wrong > 0 and not flagged, "opened": "error" not in res}


# ---------------------------------------------------------------------- 実行
def run_tool(tool: str, path: Path, out: Path, a, limit_s: float = 300) -> dict:
    if out.with_suffix(".json").exists():
        return load_json(out.with_suffix(".json"))
    if tool == "todoku":
        r = run_watched([a.todoku_python, "-I", HERE / "todoku_read.py", a.todoku_src, path, out], limit_s=limit_s)
    else:
        r = run_watched([a.baseline_python, "-I", HERE / "baseline_read.py", tool, path, out], limit_s=limit_s)
    res = last_json(r)
    if r.get("reason"):
        res = {"tool": tool, "error": r["reason"]}
    atomic_json(out.with_suffix(".json"), res)
    return res


def one(path: str, a, lo: T.LibreOffice, base: Path) -> dict:
    guard_path(path)
    did = doc_id(path)
    k = kind_of(path)
    tpath = base / "truth" / f"{did}.json"
    t = load_json(tpath)
    if t is None:
        meta = (load_json(a.meta) or {}).get(path, {}) if a.meta else {}
        t = T.build(path, k, did, base, a.baseline_python, lo, a.pkr, meta.get("content_type"))
        T.save(tpath, t)
    src = base / "work" / did / ("doc" + Path(path).suffix.lower())
    if not src.exists():  # 正解が控えにあり、作業のファイルは消してある（続きから動かしたとき）
        src.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(path, src)
    outs = base / "outputs" / did
    shutil.rmtree(outs, ignore_errors=True)  # 前の回の途中の出力（失敗の記録を含む）を使わない。書類の単位でやり直す
    outs.mkdir(parents=True, exist_ok=True)
    rec = {"id": did, "path": path, "kind": k, "scorer": SCORER_VERSION, "engine": t.get("engine"), "notes": t.get("notes"), "probes_error": (t.get("probes") or {}).get("error"),
           "encoding_source": t.get("encoding_source"), "tools": {}}
    texts = {}
    for tool in ("todoku", *TOOLS.get(k, ())):
        o = outs / f"{tool}.txt"
        res = run_tool(tool, src, o, a)
        txt = o.read_text(encoding="utf-8", errors="replace") if o.exists() and "error" not in res else ""
        texts[tool] = txt
        rec["tools"][tool] = score_tool(t, txt, tool, res)
        if tool == "todoku":
            rec["tools"][tool]["read_grade"] = res.get("read_grade")
            rec["tools"][tool]["write_grade"] = res.get("write_grade")
    if k == "pdf" and texts.get("todoku"):
        ref = [ln for ln in clean_todoku(todoku_content(texts["todoku"])).splitlines() if ln.strip()]
        for tool, s in rec["tools"].items():
            if tool != "todoku" and s.get("opened"):
                s["vs_todoku_window"] = recalls(ref, texts[tool])
    if k == "xlsx":
        mut = outs / "stale.xlsx"
        muts = load_json(outs / "stale.json")
        if muts is None:
            muts = mutate_cached(src, mut, did)
            atomic_json(outs / "stale.json", muts)
        if muts:
            rec["stale_mutated"] = {}
            for tool in ("todoku", *TOOLS[k]):
                o = outs / f"{tool}.stale.txt"
                res = run_tool(tool, mut, o, a)
                txt = o.read_text(encoding="utf-8", errors="replace") if o.exists() and "error" not in res else ""
                rec["stale_mutated"][tool] = score_mutated(muts, clean_todoku(txt) if tool == "todoku" else txt, res)
    return rec


def _purge_doc(base: Path, did: str) -> None:
    """書類ごとの作業のファイルだけを消す。採点の記録を書いたあとで呼ぶ。元の書類（--files のパス）は消さない。"""
    shutil.rmtree(base / "outputs" / did, ignore_errors=True)
    shutil.rmtree(base / "work" / did, ignore_errors=True)


CATS = ("displayed", "displayed_date", "displayed_era", "displayed_number", "formula", "stale_natural", "merged_rows", "hidden_sheets", "hidden_cells",
        "strike", "color", "color_not_normal", "furigana", "shapes", "checkboxes", "form_values", "scan_flag")


def summarize(recs: list[dict]) -> dict:
    out: dict = {"docs": len(recs), "by_kind": {}, "engines": dict(collections.Counter(r.get("engine") for r in recs)),
                 "probes_errors": sum(1 for r in recs if r.get("probes_error"))}
    for k in sorted({r["kind"] for r in recs}):
        rs = [r for r in recs if r["kind"] == k]
        tools = sorted({t for r in rs for t in r["tools"]}, key=lambda x: (x != "todoku", x))
        kd: dict = {"docs": len(rs), "tools": {}}
        for tool in tools:
            ss = [r["tools"][tool] for r in rs if tool in r["tools"]]
            opened = [s for s in ss if s.get("opened")]
            with_ref = [s for s in ss if s.get("ref_chars")]
            mean = lambda key: round(sum((s.get(key) if s.get("opened") else 0) or 0 for s in with_ref) / len(with_ref), 4) if with_ref else None  # noqa: E731
            d = {"docs": len(ss), "opened": len(opened), "char": mean("char"), "bigram": mean("bigram"), "exact_read": sum(1 for s in ss if s.get("exact_read")),
                 "docs_with_text": len(with_ref), "seconds_total": round(sum(s.get("seconds") or 0 for s in ss), 1), "cat": {},
                 "chars_out_total": sum(s.get("chars_out") or 0 for s in ss)}
            if any("char_literal" in s for s in ss):
                d["char_literal"] = mean("char_literal")
            for c in CATS:
                ns = [s["cat"][c] for s in opened if c in (s.get("cat") or {}) and s["cat"][c]["n"]]
                na = sum(1 for s in ss if c in (s.get("cat") or {}) and s["cat"][c]["n"])
                if ns:
                    n, ok = sum(x["n"] for x in ns), sum(x["ok"] for x in ns)
                    d["cat"][c] = {"n": n, "ok": ok, "docs": na, "docs_all_ok": sum(1 for x in ns if x["ok"] == x["n"]), "leak": sum(x.get("leak", 0) for x in ns)}
            vs = [s["vs_todoku_window"] for s in opened if s.get("vs_todoku_window") and s["vs_todoku_window"].get("char") is not None]
            if vs:
                d["vs_todoku_window_char"] = round(sum(v["char"] for v in vs) / len(vs), 4)
                d["vs_todoku_window_bigram"] = round(sum(v["bigram"] or 0 for v in vs) / len(vs), 4)
            oc = [s["ocr"] for s in opened if s.get("ocr") and s["ocr"].get("char") is not None]
            if oc:
                d["ocr_char"] = round(sum(v["char"] for v in oc) / len(oc), 4)
                d["ocr_bigram"] = round(sum(v["bigram"] or 0 for v in oc) / len(oc), 4)
                d["ocr_docs"] = len(oc)
            sm = [r["stale_mutated"][tool] for r in rs if tool in (r.get("stale_mutated") or {})]
            if sm:
                d["stale_mutated"] = {"docs": len(sm), "cells": sum(x["n"] for x in sm), "wrong_shown": sum(x["wrong_shown"] for x in sm),
                                      "flagged_docs": sum(1 for x in sm if x["flagged"]), "silent_docs": sum(1 for x in sm if x["silent"])}
            kd["tools"][tool] = d
        out["by_kind"][k] = kd
    return out


CAT_JA = {"displayed": "画面の字（全部）", "displayed_date": "　日付", "displayed_era": "　和暦", "displayed_number": "　数の形式", "formula": "数式の値",
          "stale_natural": "再計算と違う保存値（自然）", "merged_rows": "結合が各行に届く", "hidden_sheets": "非表示のシートの印", "hidden_cells": "非表示の字の印",
          "strike": "取り消し線", "color": "字の色", "color_not_normal": "字の色（文書でいちばん多い色を除く）", "furigana": "ふりがな", "shapes": "図形・テキストボックスの字", "checkboxes": "チェックボックスの状態",
          "form_values": "PDF の欄の値", "scan_flag": "字のないページを伝える"}


def markdown(s: dict, label: str) -> str:
    L = [f"# 読む比べ（{label}）", "", f"書類 {s['docs']} 件。正解の出どころ：{s['engines']}。作りを読めなかった書類 {s['probes_errors']} 件。", ""]
    for k, kd in s["by_kind"].items():
        tools = list(kd["tools"])
        L += [f"## {k}（{kd['docs']} 件）", ""]
        rows = [["開けた"] + [f"{kd['tools'][t]['opened']}/{kd['tools'][t]['docs']}" for t in tools],
                ["字の再現率（字のある書類の平均。開けないと 0）"] + [kd["tools"][t]["char"] for t in tools],
                ["字の再現率（窓が「同じ形の行」とまとめた所を広げない）"] + [kd["tools"][t].get("char_literal", "—") for t in tools],
                ["2 字の並びの再現率"] + [kd["tools"][t]["bigram"] for t in tools],
                ["正確に読めた（定義は上の説明）"] + [f"{kd['tools'][t]['exact_read']}/{kd['tools'][t]['docs_with_text']}" for t in tools]]
        if any("vs_todoku_window_char" in kd["tools"][t] for t in tools):
            rows.append(["字（とどくの窓を正解に：とどくに有利）"] + [kd["tools"][t].get("vs_todoku_window_char", "—") for t in tools])
        if any("ocr_char" in kd["tools"][t] for t in tools):
            rows.append(["字（Vision の文字認識を正解に。先頭 3 ページ）"] + [kd["tools"][t].get("ocr_char", "—") for t in tools])
        for c in CATS:
            if any(c in kd["tools"][t]["cat"] for t in tools):
                row = [CAT_JA[c]]
                for t in tools:
                    x = kd["tools"][t]["cat"].get(c)
                    row.append(pct(x["ok"], x["n"]) + (f"・印なし {x['leak']}" if x and x.get("leak") else "") if x else "—")
                rows.append(row)
        if any("stale_mutated" in kd["tools"][t] for t in tools):
            row = ["変えた保存値：印なしで出した書類"]
            for t in tools:
                x = kd["tools"][t].get("stale_mutated")
                row.append(f"{x['silent_docs']}/{x['docs']}（印あり {x['flagged_docs']}）" if x else "—")
            rows.append(row)
        L += [md_table(["", *tools], rows), ""]
    return "\n".join(L)


def unscored_record(base: Path, path: str, reason: str) -> dict:
    """採点できなかった書類：道具の開けた結果だけを残す。ここから exact_read は真にならない。"""
    tools = {}
    for f in (base / "outputs" / doc_id(path)).glob("*.json"):
        d = load_json(f) or {}
        if "tool" in d and ".stale." not in f.name:
            tools[d["tool"]] = {"tool": d["tool"], "opened": "error" not in d, "error": d.get("error"), "seconds": d.get("seconds"), "exact_read": False, "unscored": True}
    return {"id": doc_id(path), "path": path, "kind": kind_of(path), "scorer": SCORER_VERSION, "unscored": reason, "tools": tools}


def bad_read(rec: dict) -> str | None:
    """大量の失敗の見張り：とどくが開けない、または比べる道具のある形式で、比べる道具がどれも開けない。"""
    t = rec.get("tools") or {}
    if rec.get("unscored"):
        return f"未採点：{rec['unscored']}"
    if not (t.get("todoku") or {}).get("opened"):
        return f"とどくが開けない：{(t.get('todoku') or {}).get('error')}"
    if rec.get("kind") in ("pdf", "xlsx", "docx", "csv") and not any(s.get("opened") for k, s in t.items() if k != "todoku"):
        return "比べる道具がどれも開けない"
    return None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--files", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--baseline-python", required=True)
    ap.add_argument("--todoku-python", required=True)
    ap.add_argument("--todoku-src", required=True)
    ap.add_argument("--label", default="")
    ap.add_argument("--pkr", default=None)
    ap.add_argument("--meta", default=None, help="{パス: {content_type}}。HTTP の charset を文字のファイルの正解に使う")
    ap.add_argument("--doc-limit-s", type=float, default=900)
    ap.add_argument("--purge", action="store_true", help="書類ごとの作業のファイルを、数えたら消す（調査の書類）。元の書類は消さない")
    ap.add_argument("--one", default=None, help="（内部）1 つの書類だけを採点して scores に書く")
    a = ap.parse_args()
    base = Path(a.out)
    (base / "scores").mkdir(parents=True, exist_ok=True)
    if a.one:
        lo = T.LibreOffice(base)
        rec = one(a.one, a, lo, base)
        atomic_json(base / "scores" / f"{doc_id(a.one)}.json", rec)
        if a.purge:
            _purge_doc(base, doc_id(a.one))
        return 0
    files = [p for p in read_list(a.files) if kind_of(p)]
    T.LibreOffice(base)  # 設定を先に作る（書類の処理が同時に作らない）
    child = ["--files", a.files, "--out", a.out, "--baseline-python", a.baseline_python, "--todoku-python", a.todoku_python, "--todoku-src", a.todoku_src, "--doc-limit-s", str(a.doc_limit_s)]
    child += [x for k, v in (("--pkr", a.pkr), ("--meta", a.meta)) if v for x in (k, v)] + (["--purge"] if a.purge else [])
    def after(p, rec):
        if a.purge:
            _purge_doc(base, doc_id(p))
    def show(i, n, p, rec):
        print(f"[{i}/{n}] {Path(p).name} {rec.get('engine')} {rec.get('unscored') or ''} " + " ".join(f"{t}:{(s.get('char') if s.get('opened') else 'NG')}" for t, s in rec["tools"].items()), flush=True)
    rc, notes = drive_docs(Path(__file__).resolve(), files, base, child, a.doc_limit_s, lambda p, why: unscored_record(base, p, why), bad_read, after, show)
    if rc:
        print(notes, flush=True)
        return rc
    recs = [load_json(base / "scores" / f"{doc_id(p)}.json") for p in files if (base / "scores" / f"{doc_id(p)}.json").exists()]
    recs = [r for r in recs if r]
    s = summarize(recs)
    s["scorer"] = SCORER_VERSION
    s["unscored"] = [{"path": r["path"], "why": r["unscored"]} for r in recs if r.get("unscored")]
    atomic_json(base / "summary.json", s)
    (base / "summary.md").write_text(markdown(s, a.label), encoding="utf-8")
    print(markdown(s, a.label))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
