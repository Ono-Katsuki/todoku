"""正解を作る。とどくのコードを使わない。
- 表計算の「画面の字」：本物の Excel（外から動かす。ファイルの名前をちょうど指定し、開いたものだけを閉じる）。動かせないときは LibreOffice。
- PDF：Poppler の pdftotext -layout（主）。Vision の文字認識（Poppler で描いたページの絵から：字の層によらない。pkr があるときだけ）。
- Word：XML を直に読む（段落）。.doc は LibreOffice で変える。
Excel は確認の画面が出たら押さない：その時点で Excel を使うのをやめ、残りは LibreOffice にして、そのことを記録する。"""

from __future__ import annotations

import csv
import hashlib
import json
import os
import re
import shutil
import subprocess
import time
from pathlib import Path

from common import Heavy, atomic_json, last_json, load_json, norm, run_watched

CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/todoku_eval"
EXCEL_APP = "Microsoft Excel"
MAX_EXCEL_CELLS = 150_000
OCR_PAGES = 3
HERE = Path(__file__).resolve().parent


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


class ExcelBlocked(Exception):
    pass


def parse_ss(text: str):
    """osascript -ss の出力（AppleScript のリストの字）を Python の値に。"""
    i, n = 0, len(text)

    def ws():
        nonlocal i
        while i < n and text[i] in " \t\r\n,":
            i += 1

    def val():
        nonlocal i
        ws()
        if i >= n:
            return ""
        ch = text[i]
        if ch == "{":
            i += 1
            out = []
            while True:
                ws()
                if i >= n:
                    return out
                if text[i] == "}":
                    i += 1
                    return out
                out.append(val())
        if ch == '"':
            i += 1
            buf = []
            while i < n:
                c = text[i]
                if c == "\\" and i + 1 < n:
                    nx = text[i + 1]
                    buf.append({"n": "\n", "t": "\t", "r": "\r"}.get(nx, nx))
                    i += 2
                elif c == '"':
                    i += 1
                    break
                else:
                    buf.append(c)
                    i += 1
            return "".join(buf)
        j = i
        while j < n and text[j] not in ",}":
            j += 1
        tok = text[i:j].strip()
        i = j
        return tok

    return val()


class Excel:
    """本物の Excel。ファイルは Excel の入れ物のフォルダに、この評価だけの名前で写して開く（ほかのブックと名前がぶつからない）。"""

    def __init__(self, enabled: bool = True):
        self.enabled = enabled
        self.notice = None
        self.was_running = subprocess.run(["pgrep", "-x", EXCEL_APP], capture_output=True).returncode == 0
        self.used = False

    def _osa(self, script: str, timeout: float, retry: bool = True) -> str:
        try:
            r = subprocess.run(["nice", "-n", "10", "osascript", "-ss", "-e", f'tell application "{EXCEL_APP}"\n{script}\nend tell'],
                               capture_output=True, text=True, timeout=timeout)
        except subprocess.TimeoutExpired:
            raise ExcelBlocked(f"応答がない（{timeout:.0f} 秒）") from None
        if r.returncode:
            if "-609" in r.stderr and retry:  # 立ち上げ・終わりの途中で接続が切れた：少し待って一度だけやり直す
                time.sleep(5)
                return self._osa(script, timeout, retry=False)
            raise RuntimeError(r.stderr.strip()[:300])
        return r.stdout.rstrip("\n")

    def _disable(self, why: str) -> None:
        self.enabled = False
        self.notice = f"Excel を途中で使うのをやめた（{why}）。残りの正解は LibreOffice。確認の画面が残っていれば、本人が閉じること"

    def read(self, path: str) -> dict | None:
        """{sheets: [{name, display, fresh}]}。使えないときは None。"""
        if not self.enabled or Path(path).suffix.lower() not in (".xlsx", ".xls"):
            return None
        CONTAINER.mkdir(parents=True, exist_ok=True)
        name = f"todokueval_{hashlib.sha256(str(path).encode()).hexdigest()[:8]}_{os.getpid()}{Path(path).suffix.lower()}"
        dst = CONTAINER / name
        shutil.copy(path, dst)
        wb = f'workbook "{name}"'
        opened = False
        try:
            with Heavy():
                self.used = True
                self._osa(f'open workbook workbook file name "{dst}" update links do not update links read only true', 90)
                opened = True
                names = parse_ss(self._osa(f"get name of worksheets of {wb}", 30))
                names = names if isinstance(names, list) else [names]
                sheets = []
                for i, sn in enumerate(names, 1):
                    sheets.append({"name": sn, "display": self._cells(wb, i), "fresh": {}})
                self._osa("calculate full", 120)
                for i, s in enumerate(sheets, 1):
                    s["fresh"] = self._cells(wb, i)
                self._osa(f'close workbook "{name}" saving no', 30)
                opened = False
            return {"sheets": sheets}
        except ExcelBlocked as e:
            self._disable(f"{Path(path).name}：{e}")
            return None
        except (RuntimeError, ValueError, IndexError) as e:
            # 開けない・大きすぎる：この書類だけ LibreOffice に回す
            if opened:
                try:
                    self._osa(f'close workbook "{name}" saving no', 30)
                    opened = False
                except (RuntimeError, ExcelBlocked) as e2:
                    self._disable(f"閉じられない：{e2}")
            return {"error": str(e)[:200]}
        finally:
            if not opened:
                for f in (dst, CONTAINER / f"~${name}"):
                    try:
                        f.unlink()
                    except OSError:
                        pass

    def _cells(self, wb: str, i: int) -> dict:
        addr = parse_ss(self._osa(f"get get address (used range of worksheet {i} of {wb})", 30))
        m = re.fullmatch(r"\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?", addr)
        if not m:
            raise ValueError(f"使っている範囲が読めない：{addr}")
        c0, r0 = col_num(m[1]), int(m[2])
        c1, r1 = col_num(m[3] or m[1]), int(m[4] or m[2])
        if (c1 - c0 + 1) * (r1 - r0 + 1) > MAX_EXCEL_CELLS:
            raise ValueError("範囲が大きい")
        raw = parse_ss(self._osa(f"get string value of (used range of worksheet {i} of {wb})", 240))
        rows = raw if isinstance(raw, list) else [[raw]]
        if rows and not isinstance(rows[0], list):
            rows = [[v] for v in rows] if c0 == c1 else [rows]
        out = {}
        for ri, row in enumerate(rows):
            for ci, v in enumerate(row):
                if isinstance(v, str) and v.strip():
                    out[f"{col_name(c0 + ci)}{r0 + ri}"] = v
        return out

    def finish(self) -> None:
        """自分が立ち上げた Excel で、開いているブックがなければ終える（ほかのものは閉じない）。"""
        if self.used and not self.was_running and self.notice is None:
            try:
                self._osa("if (count of workbooks) = 0 then quit", 30)
            except (RuntimeError, ExcelBlocked):
                pass
        try:
            if CONTAINER.exists() and not any(CONTAINER.iterdir()):
                CONTAINER.rmdir()
        except OSError:
            pass


class LibreOffice:
    """自分だけの設定のフォルダで動かす（本人の LibreOffice と混ざらない）。"""

    def __init__(self, base: Path):
        self.base = Path(base)
        self.plain = self.base / "lo_profile_plain"
        self.recalc = self.base / "lo_profile_recalc"
        # 日本語の環境の表示にする（既定の日付の形が変わる）。_recalc は読み込むときに必ず計算し直す
        locale = '<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>ja-JP</value></prop></item>\n'
        recalc = ('<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
                  '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n')
        for prof, body in ((self.plain, locale), (self.recalc, locale + recalc)):
            p = prof / "user"
            if not (p / "registrymodifications.xcu").exists():
                p.mkdir(parents=True, exist_ok=True)
                (p / "registrymodifications.xcu").write_text(
                    '<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
                    'xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
                    + body + "</oor:items>\n", encoding="utf-8")

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
        self.convert(src, 'csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1', d, recalc=recalc)
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


def build(path: str, kind: str, did: str, base: Path, py: str, excel: Excel, lo: LibreOffice, pkr: str | None) -> dict:
    """1 つの書類の正解。"""
    w = _work(base, did)
    src = w / ("doc" + Path(path).suffix.lower())
    if not src.exists():
        shutil.copy(path, src)
    t: dict = {"path": path, "kind": kind, "units": [], "notes": []}
    if kind in ("xlsx", "xls"):
        pr_src = src
        if kind == "xls":
            d = w / "as_xlsx"
            lo.convert(src, "xlsx", d)
            pr_src = d / "doc.xlsx"
            t["notes"].append("作り（結合・非表示・字の書式・図形）は LibreOffice で .xlsx に変えたものから読んだ")
        t["probes"] = probes(py, pr_src, "xlsx", w / "probes.json") if pr_src.exists() else {"error": "変えられない"}
        names = [s["name"] for s in (t["probes"] or {}).get("sheets", [])]
        # 作りを読めない書類（パスワードなど）は Excel に渡さない：確認の画面を出さないため
        got = excel.read(str(src)) if excel is not None and not (t["probes"] or {}).get("error") else None
        if got and got.get("sheets"):
            t["engine"] = "excel"
            t["sheets"] = got["sheets"]
        else:
            if got and got.get("error"):
                t["notes"].append(f"Excel で読めない：{got['error']}")
            t["engine"] = "libreoffice"
            plain = lo.sheets_csv(src, names, w, recalc=False)
            has_f = any("f" in c for s in (t["probes"] or {}).get("sheets", []) for c in s.get("cells", {}).values())
            fresh = lo.sheets_csv(src, names, w, recalc=True) if has_f else plain
            t["sheets"] = [{"name": n, "display": plain.get(n, {}), "fresh": fresh.get(n, plain.get(n, {}))} for n in (names or list(plain))]
        # 字の比べの正解は、見えるシートの字（非表示のシートは画面にない。印は別の項目で数える）。非表示の行・列は残す
        hidden = {s["name"] for s in (t["probes"] or {}).get("sheets", []) if s.get("state") != "visible"}
        for s in t["sheets"]:
            if s["name"] not in hidden:
                t["units"] += list(s["display"].values())
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
    else:  # csv・txt：文字コードを決めて読むだけ（比べる道具 charset と同じなので、ここの比べは参考）
        out = w / "charset.txt"
        run_watched([py, "-I", HERE / "baseline_read.py", "charset", src, out], limit_s=120)
        t["engine"] = "charset_normalizer"
        t["units"] = [ln for ln in out.read_text(encoding="utf-8").splitlines() if ln.strip()] if out.exists() else []
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
