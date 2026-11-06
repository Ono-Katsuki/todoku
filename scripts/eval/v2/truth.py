"""採点 v2 の正解。とどくのコードを使わない。scripts/eval/truth.py の写しに、確かめた直しだけを当てたもの（一覧は下）。Office は動かさない（世界の書類は v1 でも LibreOffice）。

- 表計算の「画面の字」：LibreOffice の CSV の書き出し。
- PDF：Poppler の pdftotext -layout（主）。Vision の文字認識（Poppler で描いたページの絵から：字の層によらない。pkr があるときだけ）。
- Word：XML を直に読む（段落）。.doc は LibreOffice で変える。
- CSV・文字：下の decode_text と csv_cells。

v1 からの違い（理由と試験は docs/再評価の計画.md・tests/test_eval_v2.py）：
  T1 CSV の正解はマスの字。区切りの字と引用符は書き方であって中身でない（RFC 4180）。v1 は生の行を正解にしたので、表に読んだ道具だけが落ちた。
  T2 文字コードを charset_normalizer に決めさせない（比べる道具 charset と同じものなので、合うのが当たり前で、誤りも正解になる）。
     BOM → UTF-8 として正しいか → HTTP の charset の順に決める。どれでも決まらないときは、文字コードによらない ASCII の字だけを正解にする。
  T3 LibreOffice の CSV の第 9 の字句（Save cell contents as shown）を true に。v1 は false で、「画面の字」が元の値だった（#,##0 の 51583.1377 が 51583.1377 のまま）。
  T4 LibreOffice の環境を en-US に。ECMA-376 の組み込みの書式の表は en-US で、ja-JP では世界の書類の d-mmm-yy が 15-1月-24、AM/PM が 午後 になる。
  T5 環境で変わる組み込みの日付の書式（番号 14・22・27〜36・50〜58）のマスは、字の比べ（units）から除き、画面の字の比べで日付の並びのどれかを認める（read_compare.py）。"""

from __future__ import annotations

import codecs
import csv
import io
import json
import re
import shutil
from pathlib import Path

from common import atomic_json, last_json, lo_profile, load_json, norm, run_watched

OCR_PAGES = 3
HERE = Path(__file__).resolve().parent
# Save cell contents as shown（第 9 の字句）を true に。ほかは v1 と同じ
LO_CSV_FILTER = "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,true,false,false,-1"
LOCALE_DATE_IDS = frozenset({14, 22, *range(27, 37), *range(50, 59)})  # ECMA-376 Part 1 18.8.30：環境で字が変わる組み込みの番号
LOCALE_FMT_NAMES = frozenset(f"id{i}" for i in LOCALE_DATE_IDS)
CSV_DELIMS = (",", ";", "\t", "|")


def col_name(c: int) -> str:
    s = ""
    while c:
        c, m = divmod(c - 1, 26)
        s = chr(65 + m) + s
    return s


def col_num(s: str) -> int:
    n = 0
    for ch in s:
        n = n * 26 + ord(ch) - 64
    return n


# ---------------------------------------------------------------------- 文字のファイル
_DECLARED = re.compile(r"charset\s*=\s*\"?([A-Za-z0-9._:-]+)", re.I)
_WHATWG = {"iso-8859-1": "cp1252", "iso8859-1": "cp1252", "latin1": "cp1252", "latin-1": "cp1252", "us-ascii": "cp1252", "ascii": "cp1252"}


def decode_text(b: bytes, content_type: str | None = None) -> tuple[str, str]:
    """(字, どう決めたか)。bom・utf8・declared・unverified。unverified のときの字は ASCII の字だけ（それ以外は U+FFFD にする：どの文字コードが正しいか、規則で決まらない）。
    declared の iso-8859-1・ascii は WHATWG の Encoding Standard のとおり windows-1252 として読み、それで読めなければ iso-8859-1 にする。"""
    for bom, enc in ((codecs.BOM_UTF8, "utf-8"), (codecs.BOM_UTF32_LE, "utf-32-le"), (codecs.BOM_UTF32_BE, "utf-32-be"), (codecs.BOM_UTF16_LE, "utf-16-le"), (codecs.BOM_UTF16_BE, "utf-16-be")):
        if b.startswith(bom):
            try:
                return b[len(bom):].decode(enc), "bom"
            except UnicodeDecodeError:
                break
    try:
        return b.decode("utf-8"), "utf8"
    except UnicodeDecodeError:
        pass
    m = _DECLARED.search(content_type or "")
    if m:
        name = m[1].lower()
        try:
            codecs.lookup(name)
        except LookupError:
            name = ""
        if name and name not in ("utf-8", "utf8"):
            for enc in (_WHATWG.get(name, name), name):
                try:
                    return b.decode(enc), "declared"
                except (UnicodeDecodeError, LookupError):
                    pass
    return b.decode("ascii", "replace"), "unverified"


def csv_delimiter(text: str, suffix: str = ".csv") -> str:
    """規則で決める：.tsv はタブ。ほかは、候補ごとに csv.reader で読み、2 列以上でいちばん多い列の数の行が、空でない行に占める割合の高いもの。
    同じ割合なら、その行が持つ区切りの数の多いもの。それも同じならコンマ（RFC 4180）を先に。どの候補でも 2 列にならなければコンマ（1 列の表になる）。"""
    if suffix.lower() == ".tsv":
        return "\t"
    lines = text.splitlines()[:2000]
    best, score = ",", (0.0, 0)
    for d in CSV_DELIMS:
        counts: dict[int, int] = {}
        rows = 0
        try:
            for row in csv.reader(io.StringIO("\n".join(lines)), delimiter=d):
                if row:
                    rows += 1
                    if len(row) > 1:
                        counts[len(row)] = counts.get(len(row), 0) + 1
        except csv.Error:
            continue
        if not counts:
            continue
        cols, k = max(counts.items(), key=lambda x: (x[1], x[0]))
        s = (round(k / rows, 6), (cols - 1) * k)
        if s > score:
            best, score = d, s
    return best


def csv_cells(text: str, suffix: str = ".csv") -> tuple[list[str], str]:
    """(空でないマスの字, 区切りの字)。"""
    d = csv_delimiter(text, suffix)
    out = []
    csv.field_size_limit(10_000_000)
    try:
        for row in csv.reader(io.StringIO(text, newline=""), delimiter=d):
            out += [c for c in row if c.strip()]
    except csv.Error:  # CSV として読めない行：字のまま（区切りを除けない）
        return [ln for ln in text.splitlines() if ln.strip()], d
    return out, d


# ---------------------------------------------------------------------- LibreOffice
class LibreOffice:
    """自分だけの設定のフォルダで動かす（本人の LibreOffice と混ざらない）。"""

    def __init__(self, base: Path):
        self.base = Path(base)
        self.plain = self.base / "lo_profile_plain"
        self.recalc = self.base / "lo_profile_recalc"
        lo_profile(self.plain, locale="en-US")
        lo_profile(self.recalc, recalc=True, locale="en-US")

    def convert(self, src: Path, to: str, outdir: Path, recalc: bool = False, limit_s: float = 180) -> dict:
        outdir.mkdir(parents=True, exist_ok=True)
        prof = self.recalc if recalc else self.plain
        return run_watched(["soffice", f"-env:UserInstallation=file://{prof}", "--headless", "--norestore", "--convert-to", to, "--outdir", outdir, src],
                           limit_s=limit_s, mem_mb=3000)

    def sheets_csv(self, src: Path, names: list[str], work: Path, recalc: bool) -> dict:
        """シートごとの画面の字（表示の形式どおり）：{シート名: {番地: 字}}。"""
        d = work / ("csv_recalc" if recalc else "csv_plain")
        if d.exists():
            shutil.rmtree(d)
        self.convert(src, LO_CSV_FILTER, d, recalc=recalc)
        out = {}
        files = sorted(d.glob("*.csv"))
        for f in files:
            sn = f.stem[len(src.stem) + 1:] if f.stem.startswith(src.stem + "-") else f.stem
            cells = {}
            with f.open(encoding="utf-8", newline="", errors="replace") as fh:
                for r, row in enumerate(csv.reader(fh), 1):
                    for c, v in enumerate(row, 1):
                        if v.strip():
                            cells[f"{col_name(c)}{r}"] = v
            out[sn] = cells
        return out


def _work(base: Path, did: str) -> Path:
    w = base / "work" / did
    w.mkdir(parents=True, exist_ok=True)
    return w


def probes(py: str, path, kind: str, out: Path) -> dict | None:
    r = run_watched([py, "-I", HERE / "probes.py", path, kind, out], limit_s=300)
    if last_json(r).get("ok"):
        return load_json(out)
    return {"error": last_json(r).get("error")}


def is_locale_date(c: dict) -> bool:
    """環境で字が変わる組み込みの日付の書式のマス（作りの probes の記録から）。"""
    return c.get("fmt") in LOCALE_FMT_NAMES


def build(path: str, kind: str, did: str, base: Path, py: str, lo: LibreOffice, pkr: str | None, content_type: str | None = None) -> dict:
    """1 つの書類の正解。"""
    w = _work(base, did)
    src = w / ("doc" + Path(path).suffix.lower())
    if not src.exists():
        shutil.copy(path, src)
    t: dict = {"path": path, "kind": kind, "units": [], "notes": [], "scorer": "v2"}
    if kind in ("xlsx", "xls"):
        pr_src = src
        if kind == "xls":
            d = w / "as_xlsx"
            lo.convert(src, "xlsx", d)
            pr_src = d / "doc.xlsx"
            t["notes"].append("作り（結合・非表示・字の書式・図形）は LibreOffice で .xlsx に変えたものから読んだ")
        t["probes"] = probes(py, pr_src, "xlsx", w / "probes.json") if pr_src.exists() else {"error": "変えられない"}
        names = [s["name"] for s in (t["probes"] or {}).get("sheets", [])]
        t["engine"] = "libreoffice"
        plain = lo.sheets_csv(src, names, w, recalc=False)
        has_f = any("f" in c for s in (t["probes"] or {}).get("sheets", []) for c in s.get("cells", {}).values())
        fresh = lo.sheets_csv(src, names, w, recalc=True) if has_f else plain
        t["sheets"] = [{"name": n, "display": plain.get(n, {}), "fresh": fresh.get(n, plain.get(n, {}))} for n in (names or list(plain))]
        # 字の比べの正解は、見えるシートの字（非表示のシートは画面にない。印は別の項目で数える）。非表示の行・列は残す。環境で字が変わる日付は除く（T5）
        sheets_pr = {s["name"]: s for s in (t["probes"] or {}).get("sheets", [])}
        hidden = {n for n, s in sheets_pr.items() if s.get("state") != "visible"}
        excl = 0
        for s in t["sheets"]:
            if s["name"] in hidden:
                continue
            cells = sheets_pr.get(s["name"], {}).get("cells", {})
            for addr, v in s["display"].items():
                if is_locale_date(cells.get(addr) or {}):
                    excl += 1
                else:
                    t["units"].append(v)
        t["locale_dates_excluded"] = excl
    elif kind in ("docx", "doc"):
        pr_src = src
        if kind == "doc":
            d = w / "as_docx"
            lo.convert(src, "docx", d)
            pr_src = d / "doc.docx"
        t["probes"] = probes(py, pr_src, "docx", w / "probes.json") if pr_src.exists() else {"error": "変えられない"}
        t["engine"] = "xml" if kind == "docx" else "libreoffice"
        pr = t["probes"] or {}
        t["units"] = list(pr.get("paragraphs", [])) + list(pr.get("textboxes", []))
    elif kind == "pdf":
        t["probes"] = probes(py, src, "pdf", w / "probes.json")
        t["engine"] = "poppler"
        r = run_watched(["pdftotext", "-layout", "-enc", "UTF-8", src, "-"], limit_s=180)
        txt = r["stdout"] if r["returncode"] == 0 else ""
        t["pages_text_chars"] = [len(norm(pg)) for pg in txt.split("\f")]
        # 字のかたまりに分ける：-layout は同じ高さの別の列を 1 行に並べるので、列をまたぐ 2 字の並びを正解にしない
        t["units"] = [u for ln in txt.splitlines() for u in re.split(r"[ 　]{2,}", ln) if u.strip()]
        if pkr:
            ocr = _ocr(src, w, pkr, (t["probes"] or {}).get("pages") or 1)
            if ocr is not None:
                t["ocr_units"] = ocr
    elif kind in ("odt", "rtf"):
        d = w / "txt"
        lo.convert(src, "txt:Text (encoded):UTF8", d)
        f = d / "doc.txt"
        t["engine"] = "libreoffice"
        t["units"] = [ln for ln in f.read_text(encoding="utf-8", errors="replace").splitlines() if ln.strip()] if f.exists() else []
    elif kind == "ods":
        plain = lo.sheets_csv(src, [], w, recalc=False)
        t["engine"] = "libreoffice"
        t["sheets"] = [{"name": n, "display": c, "fresh": c} for n, c in plain.items()]
        for c in plain.values():
            t["units"] += list(c.values())
    else:  # csv・txt
        text, how = decode_text(src.read_bytes(), content_type)
        t["engine"] = "decode_text"
        t["encoding_source"] = how
        if how == "unverified":
            t["notes"].append("文字コードが規則で決まらない：ASCII の字だけを正解にした")
        if kind == "csv":
            cells, delim = csv_cells(text, Path(path).suffix)
            t["csv_delimiter"] = delim
            t["units"] = cells
        else:
            t["units"] = [ln for ln in text.splitlines() if ln.strip()]
        if how == "unverified":
            # 分からない字の所で単位を切る（その前後の字は、書類の中で並んでいない）
            t["units"] = [x for u in t["units"] for x in u.split("�") if x.strip()]
    return t


def _ocr(src: Path, w: Path, pkr: str, pages: int) -> list[str] | None:
    d = w / "ocr"
    d.mkdir(exist_ok=True)
    pngs = []
    for pg in range(1, min(pages, OCR_PAGES) + 1):
        png = d / f"p{pg}.png"
        if not png.exists():
            run_watched(["pdftoppm", "-r", "150", "-f", pg, "-l", pg, "-png", "-singlefile", src, d / f"p{pg}"], limit_s=120)
        if png.exists():
            pngs.append(png)
    if not pngs:
        return None
    r = run_watched([pkr, "ocr", *pngs], limit_s=300)
    out = []
    for ln in r["stdout"].splitlines():
        try:
            out += json.loads(ln).get("text", [])
        except json.JSONDecodeError:
            pass
    for f in pngs:
        f.unlink(missing_ok=True)
    return out if r["returncode"] == 0 else None


def save(path, t) -> None:
    atomic_json(path, t)
