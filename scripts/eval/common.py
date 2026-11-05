"""最終評価の共通の部品：ファイルの一覧・見張りつきの実行（nice・メモリと時間の上限・重い処理は一度に一つ）・記録・区間。
プロジェクトの依存を増やさない（標準ライブラリだけ）。比べる道具は別の venv（setup_baselines.sh）で、別の処理として動かす。"""

from __future__ import annotations

import fcntl
import hashlib
import json
import math
import os
import re
import shutil
import signal
import subprocess
import sys
import time
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
EVAL_DIR = Path(__file__).resolve().parent
Z = 1.959964
HELDOUT_NAMES = ("forms_tama", "forms_world")
KINDS = {".xlsx": "xlsx", ".xlsm": "xlsx", ".xls": "xls", ".docx": "docx", ".doc": "doc", ".pdf": "pdf"}
MORE_KINDS = {".csv": "csv", ".tsv": "csv", ".txt": "txt", ".rtf": "rtf", ".odt": "odt", ".ods": "ods"}


# ---------------------------------------------------------------------- 開発用の守り：試すときは、とっておいた書類に決して触らない
def dev_only() -> bool:
    return os.environ.get("EVAL_DEV_ONLY") == "1"


def guard_path(p) -> None:
    """EVAL_DEV_ONLY=1（試運転）のあいだ、とっておいた書類のパスは開く前に止める。"""
    if dev_only() and any(n in str(p) for n in HELDOUT_NAMES):
        raise SystemExit(f"試運転では、とっておいた書類を開きません：{p}")


# ---------------------------------------------------------------------- ファイルの一覧
def kind_of(p) -> str | None:
    s = Path(p).suffix.lower()
    return KINDS.get(s) or MORE_KINDS.get(s)


def doc_id(p) -> str:
    """結果のファイルの名前：読める短い名前 + パスのハッシュ（同じ名前の様式があるため）。"""
    p = str(p)
    stem = re.sub(r"[^0-9A-Za-z_.-]+", "_", Path(p).name)[-40:]
    return f"{hashlib.sha256(p.encode('utf-8', 'surrogateescape')).hexdigest()[:10]}_{stem}"


def walk_files(root, kinds: dict | None = None) -> list[str]:
    guard_path(root)
    want = kinds or KINDS
    out = []
    for dp, dn, fn in os.walk(root):
        dn.sort()
        for f in sorted(fn):
            q = os.path.join(dp, f)
            if not f.startswith(("~$", ".")) and Path(f).suffix.lower() in want:
                out.append(q)
    return out


def read_list(path) -> list[str]:
    out = [ln.strip() for ln in Path(path).read_text(encoding="utf-8").splitlines() if ln.strip() and not ln.startswith("#")]
    for p in out:
        guard_path(p)
    return out


# ---------------------------------------------------------------------- 実行
class Heavy:
    """重い処理は一度に一つ（この PC は一度落ちた）。同じ錠のファイルを使う処理のあいだで、順番に動く。"""

    def __init__(self, path=None):
        self.path = path or os.environ.get("EVAL_LOCK") or str(Path(os.environ.get("TMPDIR", "/tmp")) / "todoku_eval.lock")
        self.f = None

    def __enter__(self):
        self.f = open(self.path, "w")
        fcntl.flock(self.f, fcntl.LOCK_EX)
        return self

    def __exit__(self, *a):
        fcntl.flock(self.f, fcntl.LOCK_UN)
        self.f.close()


def nice_prefix() -> list[str]:
    """優先度を nice 10 にそろえる（すでに下げてあれば、それ以上下げない）。"""
    cur = os.nice(0)
    return ["nice", "-n", str(10 - cur)] if cur < 10 else []


def _group_rss_mb(pgid: int) -> float:
    try:
        out = subprocess.run(["ps", "-A", "-o", "pgid=,rss="], capture_output=True, text=True, timeout=10).stdout
    except (subprocess.SubprocessError, OSError):
        return 0.0
    tot = 0
    for ln in out.splitlines():
        a = ln.split()
        if len(a) == 2 and a[0] == str(pgid):
            tot += int(a[1])
    return tot / 1024


def run_watched(cmd: list, limit_s: float = 300, mem_mb: float = 2048, env: dict | None = None, cwd=None, heavy: bool = True,
                stdout_limit: int = 20_000_000) -> dict:
    """nice -n 10 で動かし、メモリ（処理のまとまりの合計）と時間を見張る。超えたら止めて、その理由を返す。"""
    e = dict(os.environ)
    if env:
        e.update(env)

    def body():
        t = time.time()
        p = subprocess.Popen([*nice_prefix(), *[str(c) for c in cmd]], stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=e, cwd=cwd,
                             start_new_session=True)
        why, peak = None, 0.0
        o, er = b"", b""
        while True:
            try:
                o, er = p.communicate(timeout=1)
                break
            except subprocess.TimeoutExpired:
                peak = max(peak, _group_rss_mb(p.pid))
                if peak > mem_mb:
                    why = "memory"
                elif time.time() - t > limit_s:
                    why = "timeout"
                if why:
                    try:
                        os.killpg(p.pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                    o, er = p.communicate()
                    break
        return {"returncode": p.returncode, "stdout": o[:stdout_limit].decode("utf-8", "replace"), "stderr": er[-2000:].decode("utf-8", "replace"),
                "reason": why, "seconds": round(time.time() - t, 2), "peak_mb": round(peak)}

    if heavy:
        with Heavy():
            return body()
    return body()


def last_json(res: dict) -> dict:
    lines = (res.get("stdout") or "").strip().splitlines()
    for ln in reversed(lines):
        try:
            d = json.loads(ln)
            if isinstance(d, dict):
                return d
        except json.JSONDecodeError:
            continue
    return {"error": res.get("reason") or f"出力なし（rc={res.get('returncode')}）：{(res.get('stderr') or '')[-200:]}"}


def atomic_json(path, obj) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_name(path.name + ".tmp")
    tmp.write_text(json.dumps(obj, ensure_ascii=False, indent=1, default=str), encoding="utf-8")
    tmp.replace(path)


def load_json(path, default=None):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default


# ---------------------------------------------------------------------- LibreOffice の設定
USER_LO_FONTS = Path.home() / "Library/Application Support/LibreOffice/4/user/fonts"


def lo_profile(path, recalc: bool = False) -> str:
    """LibreOffice の、この評価だけの設定を作る（本人の設定と混ざらない）。日本語の環境の表示にする。recalc は読み込むときに必ず計算し直す。
    この Mac の LibreOffice はシステムの書体を使わず、日本語の字は本人の設定の user/fonts（Noto Sans JP）から描く。新しい設定にはそれがなく、字が OpenSymbol に
    落ちて絵に出ない（字は PDF にあるのに描かれない）ので、書体を写す。ないときは止める：日本語の出ない絵で確かめても意味がないため。"""
    p = Path(path) / "user"
    f = p / "registrymodifications.xcu"
    if f.exists():
        return str(path)
    fonts = sorted(USER_LO_FONTS.glob("*.[to]tf")) if USER_LO_FONTS.is_dir() else []
    if not fonts:
        raise SystemExit(f"LibreOffice の日本語の書体が見つかりません（{USER_LO_FONTS}）。Noto Sans JP を入れてから動かしてください")
    (p / "fonts").mkdir(parents=True, exist_ok=True)
    for x in fonts:
        shutil.copy(x, p / "fonts" / x.name)
    items = ['<item oor:path="/org.openoffice.Setup/L10N"><prop oor:name="ooSetupSystemLocale" oor:op="fuse"><value>ja-JP</value></prop></item>']
    if recalc:
        items += ['<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>',
                  '<item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>']
    f.write_text('<?xml version="1.0" encoding="UTF-8"?>\n<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
                 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n' + "\n".join(items) + "\n</oor:items>\n", encoding="utf-8")
    return str(path)


# ---------------------------------------------------------------------- 字の比べ
_BOXES = str.maketrans({"☐": "□", "❏": "□", "▢": "□"})  # 同じに見える箱の字（道具によって字の番号が変わる）は同じに数える


def norm(s: str) -> str:
    """NFKC にして、空白を除く（字の並びの比べ）。"""
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", s or "").translate(_BOXES))


def norm_sp(s: str) -> str:
    """NFKC にして、空白の並びを 1 つにする（値の境目を残す比べ）。"""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", s or "")).strip()


# ---------------------------------------------------------------------- 区間
def wilson(k: float, n: float) -> list:
    if not n:
        return [None, None]
    p = k / n
    d = 1 + Z * Z / n
    c = (p + Z * Z / (2 * n)) / d
    h = Z * math.sqrt(max(p * (1 - p) / n + Z * Z / (4 * n * n), 0)) / d
    return [round(max(0.0, c - h), 4), round(min(1.0, c + h), 4)]


def _betacf(a: float, b: float, x: float) -> float:
    def safe(v: float) -> float:
        return v if abs(v) > 1e-300 else 1e-300

    qab, qap, qam = a + b, a + 1, a - 1
    c, d = 1.0, 1 / safe(1 - qab * x / qap)
    h = d
    for m in range(1, 500):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 / safe(1 + aa * d)
        c = safe(1 + aa / c)
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 / safe(1 + aa * d)
        c = safe(1 + aa / c)
        de = d * c
        h *= de
        if abs(de - 1) < 1e-12:
            break
    return h


def betainc(a: float, b: float, x: float) -> float:
    """正則化された不完全ベータ関数 I_x(a, b)。"""
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    lb = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x)
    if x < (a + 1) / (a + b + 2):
        return math.exp(lb) * _betacf(a, b, x) / a
    return 1 - math.exp(lb) * _betacf(b, a, 1 - x) / b


def beta_ppf(q: float, a: float, b: float) -> float:
    lo, hi = 0.0, 1.0
    for _ in range(80):
        mid = (lo + hi) / 2
        if betainc(a, b, mid) < q:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def clopper_pearson(x: float, n: float, alpha: float = 0.05) -> list:
    """x は小数でもよい（有効な標本の大きさ × 割合）。"""
    if n <= 0:
        return [None, None]
    lo = 0.0 if x <= 0 else beta_ppf(alpha / 2, x, n - x + 1)
    hi = 1.0 if x >= n else beta_ppf(1 - alpha / 2, x + 1, n - x)
    return [round(lo, 5), round(hi, 5)]


def t_ppf(q: float, df: float) -> float:
    """t 分布の分位点（q > 0.5）。分布関数は不完全ベータ関数から。"""
    lo, hi = 0.0, 1e7
    for _ in range(200):
        mid = (lo + hi) / 2
        cdf = 1 - 0.5 * betainc(df / 2, 0.5, df / (df + mid * mid))
        if cdf < q:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def korn_graubard(p: float, var: float, n: float, alpha: float = 0.05, df: float | None = None) -> dict:
    """Korn & Graubard (1998)：分散から有効な標本の大きさ n* = p(1-p)/var を出し、デザインの自由度 df があれば
    n* × (t[n-1] / t[df])^2 に直して（n を超えない）、Clopper-Pearson の区間を当てる。分散が 0（どの層も 0% か 100%）のときは n* = n。"""
    n_eff = p * (1 - p) / var if var > 0 and 0 < p < 1 else n
    if var > 0 and 0 < p < 1 and df is not None and df >= 1 and n > 1:
        n_eff *= (t_ppf(1 - alpha / 2, n - 1) / t_ppf(1 - alpha / 2, df)) ** 2
    n_eff = min(n_eff, n)
    ci = clopper_pearson(p * n_eff, n_eff, alpha)
    return {"p": round(p, 5), "ci": ci, "n": n, "n_eff": round(n_eff, 1), **({"df": df} if df is not None else {})}


def stratified_estimate(cells: list[tuple], alpha: float = 0.05) -> dict | None:
    """層別の推定と Korn & Graubard の区間（世界のポータル・最終の評価の数で同じもの）。
    cells：(層の大きさ N, 割合 p, 分散 v, 件数 n, 集落の数 k)。件数のない層は除く。df ＝ 集落の数 − 層の数。層がなければ None。"""
    cells = [c for c in cells if c[3] > 0 and c[0] > 0 and c[1] is not None]
    tot = sum(c[0] for c in cells)
    if not tot:
        return None
    p = sum(c[0] / tot * c[1] for c in cells)
    var = sum((c[0] / tot) ** 2 * c[2] for c in cells)
    df = sum(c[4] - 1 for c in cells)
    return {**korn_graubard(p, var, sum(c[3] for c in cells), alpha, df), "frame": tot}


# ---------------------------------------------------------------------- 表
def md_table(header: list, rows: list) -> str:
    out = ["| " + " | ".join(str(h) for h in header) + " |", "|" + "---|" * len(header)]
    out += ["| " + " | ".join("" if c is None else str(c) for c in r) + " |" for r in rows]
    return "\n".join(out)


def pct(k, n, d: int = 1) -> str:
    return f"{k}/{n}（{100 * k / n:.{d}f}%）" if n else "—"


def eprint(*a):
    print(*a, file=sys.stderr, flush=True)
