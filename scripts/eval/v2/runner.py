#!/usr/bin/env python3
"""再評価の実行（採点 v2）。凍結のコードを変えず、この中だけで動く。命令は run_freeze2.sh から呼ぶ。

  tool      --tag T --out D               タグの src を D/tool/src に取り出す（作業ツリーと別。道具の版を固定する）
  preflight --out D                       試しの小さな書類で、道具・比べる道具・採点が通しで動くことを確かめる。通らなければ止まる
  fetch     --out D --subset S [--docs-from DIR]  書類を D/docs に取る（sha256 を確かめる。断られた書類は記録して、その書類をやめる）
  score     --out D [--budget-hours H]    読む・書くの比べを 50 件ずつ。止めても続きから。空きが 3 GB 未満で止まる
  verify    --out D                       全部に採点の記録があるか、失敗が多くないかを確かめる。通れば VERIFIED を書く
  estimate  --out D --subset S --label L  数を出す（D/publish_candidate に、地図とポータルが読む形で。本物のファイルには書かない）
  cleanup   --out D                       VERIFIED と数があるときだけ、書類と作業のファイルを消す

守ること：書類を消すのは cleanup だけ。結果の置き場は --out の下だけ（リポジトリ・動画の world_counts.json・ポータルの data.json には書けない）。"""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import os
import shutil
import signal
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import ROOT, SCORER_VERSION, atomic_json, doc_id, eprint, kind_of, load_json, md_table, stratified_estimate  # noqa: E402

W = ROOT / "world"
MIN_FREE_GB = 3.0
CHUNK = 50
EXIT_DISK, EXIT_ABORT, EXIT_PREFLIGHT, EXIT_BUSY, EXIT_BUDGET = 75, 78, 77, 76, 74
OTHER_HEAVY = ("rescore.py", "truth2.py", "rewrite.py", "world_survey.py", "world_counts.py compare", "eval/read_compare.py", "eval/write_compare.py", "fill_runner.py")
WRITE_KINDS = {"xlsx", "docx", "pdf", "xls", "doc"}
NO_BASELINE_KINDS = {"doc", "odt", "ods", "rtf"}  # 比べる道具（Python のライブラリ）に読む手段がない形式。「新しく届く」の感度の比べに使う


# ---------------------------------------------------------------------- 守り
def real_targets() -> list[Path]:
    home_film = os.environ.get("TODOKU_FILM_COUNTS")
    out = [W / "portal" / "data.json", W / "results.jsonl", W / "sample.jsonl"]
    if home_film:
        out.append(Path(home_film))
    return [p.resolve() for p in out]


def guard_out(out: Path) -> Path:
    """結果の置き場はリポジトリの外で、film の下でもない。"""
    o = out.resolve()
    if str(o).startswith(str(ROOT.resolve()) + os.sep) or o == ROOT.resolve():
        raise SystemExit(f"結果の置き場はリポジトリの外にしてください：{o}")
    if any(p.name == "film" for p in [o, *o.parents]):
        raise SystemExit(f"動画の作業フォルダの下には書きません：{o}")
    return o


def free_gb(path: Path) -> float:
    return shutil.disk_usage(path).free / 1024**3


def others_running() -> list[str]:
    try:
        txt = subprocess.run(["ps", "-Ao", "pid,command"], capture_output=True, text=True, timeout=20).stdout
    except (subprocess.SubprocessError, OSError):
        return []
    me = str(os.getpid())
    return [ln.strip()[:150] for ln in txt.splitlines() if any(k in ln for k in OTHER_HEAVY) and "ps -Ao" not in ln and not ln.strip().startswith(me) and "scripts/eval/v2/" not in ln]


def env_lock(out: Path) -> dict:
    e = dict(os.environ)
    e["EVAL_LOCK"] = str(out / ".heavy.lock")
    return e


def pythons(a) -> tuple[str, str]:
    bpy = a.baseline_python or os.environ.get("TODOKU_BASELINE_PY") or ""
    tpy = a.todoku_python or os.environ.get("TODOKU_PY") or str(ROOT / ".venv/bin/python")
    if not Path(bpy).exists():
        raise SystemExit(f"比べる道具の python がありません：{bpy!r}（--baseline-python か TODOKU_BASELINE_PY）")
    if not Path(tpy).exists():
        raise SystemExit(f"とどくの python がありません：{tpy!r}（--todoku-python か TODOKU_PY）")
    return bpy, tpy


# ---------------------------------------------------------------------- 道具
def cmd_tool(a) -> int:
    out = guard_out(Path(a.out))
    dst = out / "tool"
    sha = subprocess.run(["git", "-C", str(ROOT), "rev-parse", f"{a.tag}^{{commit}}"], capture_output=True, text=True).stdout.strip()
    if not sha:
        print(f"タグがありません：{a.tag}", file=sys.stderr)
        return 2
    prov = load_json(dst / "provenance.json")
    if prov and prov.get("commit") == sha and (dst / "src").exists():
        print(f"道具はもうある：{a.tag} {sha}")
        return 0
    if dst.exists():
        shutil.rmtree(dst)
    dst.mkdir(parents=True)
    arch = subprocess.Popen(["git", "-C", str(ROOT), "archive", a.tag, "src"], stdout=subprocess.PIPE)
    tar = subprocess.run(["tar", "-x", "-C", str(dst)], stdin=arch.stdout)
    arch.wait()
    if arch.returncode or tar.returncode:
        print("取り出せない", file=sys.stderr)
        return 2
    scorer = subprocess.run(["git", "-C", str(ROOT), "rev-parse", "HEAD"], capture_output=True, text=True).stdout.strip()
    dirty = subprocess.run(["git", "-C", str(ROOT), "status", "--porcelain", "scripts/eval/v2"], capture_output=True, text=True).stdout.strip()
    atomic_json(dst / "provenance.json", {"tag": a.tag, "commit": sha, "scorer": SCORER_VERSION, "scorer_commit": scorer, "scorer_dirty": bool(dirty), "at": time.strftime("%F %T")})
    print(f"道具：{a.tag} {sha}（採点の commit {scorer}{'・作業中の変更あり' if dirty else ''}）")
    return 0


# ---------------------------------------------------------------------- 試し（canary）
def cmd_preflight(a) -> int:
    out = guard_out(Path(a.out))
    bpy, tpy = pythons(a)
    src = out / "tool" / "src"
    problems: list[str] = []
    if os.getpriority(os.PRIO_PROCESS, 0) >= 20:
        problems.append("nice が 20：v1 の PermissionError の大量の失敗の条件。nice を重ねないこと")
    if free_gb(out if out.exists() else out.parent) < MIN_FREE_GB:
        problems.append(f"空きが {MIN_FREE_GB} GB 未満")
    for tool in ("soffice", "pdftotext", "pdftoppm"):
        if not shutil.which(tool):
            problems.append(f"{tool} がない")
    if not a.pkr or not Path(a.pkr).exists():
        problems.append(f"文字認識の道具 pkr がない：{a.pkr}（PDF の確かめと R3 に要る）")
    if not src.exists():
        problems.append(f"道具がない：{src}（先に tool）")
    if problems:
        print("通らない：" + "；".join(problems), file=sys.stderr)
        return EXIT_PREFLIGHT
    d = out / "canary"
    if d.exists():
        shutil.rmtree(d)
    d.mkdir(parents=True)
    env = env_lock(out)
    r = subprocess.run([bpy, "-I", str(HERE / "canary_make.py"), str(d / "docs")], capture_output=True, text=True, env=env, timeout=300)
    if r.returncode:
        print("試しの書類を作れない：" + r.stderr[-500:], file=sys.stderr)
        return EXIT_PREFLIGHT
    files = sorted(str(p) for p in (d / "docs").iterdir() if kind_of(p))
    lst = d / "files.txt"
    lst.write_text("\n".join(files) + "\n", encoding="utf-8")
    common = ["--files", str(lst), "--baseline-python", bpy, "--todoku-python", tpy, "--todoku-src", str(src), "--pkr", a.pkr, "--doc-limit-s", "300"]
    r1 = subprocess.run([sys.executable, "-I", str(HERE / "read_compare.py"), "--out", str(d / "read"), *common], capture_output=True, text=True, env=env)
    r2 = subprocess.run([sys.executable, "-I", str(HERE / "write_compare.py"), "--out", str(d / "write"), "--no-signed", *common], capture_output=True, text=True, env=env)
    if r1.returncode or r2.returncode:
        print(f"比べが通らない：rc {r1.returncode}/{r2.returncode}\n{r1.stderr[-600:]}\n{r2.stderr[-600:]}", file=sys.stderr)
        return EXIT_PREFLIGHT
    for p in files:
        k = kind_of(p)
        rd = load_json(d / "read" / "scores" / f"{doc_id(p)}.json") or {}
        t = rd.get("tools") or {}
        td = t.get("todoku") or {}
        if not td.get("exact_read"):
            problems.append(f"{Path(p).name}：とどくが正確に読めない（字 {td.get('char')}・2 字 {td.get('bigram')}・{td.get('error')}）。窓の印の変わりで採点の復号が合わないかもしれない")
        if not any(s.get("exact_read") for n, s in t.items() if n != "todoku"):
            problems.append(f"{Path(p).name}：比べる道具のどれも正確に読めない")
        if k in WRITE_KINDS:
            wr = load_json(d / "write" / "scores" / f"{doc_id(p)}.json") or {}
            runs = wr.get("runs") or {}
            if not (runs.get("todoku:ascii") or {}).get("write_ok"):
                problems.append(f"{Path(p).name}：とどくが書けない：{(runs.get('todoku:ascii') or {}).get('error') or wr.get('skip')}")
            if not any(v.get("write_ok") for n, v in runs.items() if n.endswith(":ascii") and not n.startswith("todoku")):
                problems.append(f"{Path(p).name}：比べる道具が書けない")
    if problems:
        print("通らない：\n  " + "\n  ".join(problems), file=sys.stderr)
        return EXIT_PREFLIGHT
    shutil.rmtree(d, ignore_errors=True)
    atomic_json(out / "PREFLIGHT_OK", {"at": time.strftime("%F %T"), "files": [Path(p).name for p in files], "nice": os.getpriority(os.PRIO_PROCESS, 0)})
    print(f"試し：通った（{len(files)} 件）")
    return 0


# ---------------------------------------------------------------------- 書類を取る
def cmd_fetch(a) -> int:
    from common import FailGuard, MassFailure

    sys.path.insert(0, str(ROOT / "scripts"))
    import world_survey as ws

    out = guard_out(Path(a.out))
    if a.mock_limit is None and others_running() and not a.ignore_busy:
        print("ほかの重い処理が動いている：\n  " + "\n  ".join(others_running()), file=sys.stderr)
        return EXIT_BUSY
    sub = load_json(a.subset)
    if not sub or not sub.get("docs"):
        print(f"一部の一覧がない：{a.subset}", file=sys.stderr)
        return 2
    docs = sub["docs"][: a.mock_limit] if a.mock_limit else sub["docs"]
    if a.only:
        want = {x for x in a.only.split(",") if x}
        docs = [s for s in sub["docs"] if s["id"] in want]
    dd = out / "docs"
    dd.mkdir(parents=True, exist_ok=True)
    mp_path, meta_path = out / "docs_map.json", out / "meta.json"
    mapping = load_json(mp_path) or {}
    meta = load_json(meta_path) or {}
    guard = FailGuard()
    reused = fetched = 0
    from_dir = Path(a.docs_from) if a.docs_from else None
    for s in docs:
        if s["id"] in mapping:
            continue
        if free_gb(dd) < MIN_FREE_GB:
            print(f"空きが {MIN_FREE_GB} GB 未満。止める（もう一度動かすと続きから）", file=sys.stderr)
            return EXIT_DISK
        rec, body = {}, None
        if from_dir and s.get("sha256"):  # すでに手元にある写し（同じ sha256 のときだけ。読むだけ）
            for c in from_dir.glob(f"{s['id']}.*"):
                b = c.read_bytes()
                if hashlib.sha256(b).hexdigest() == s["sha256"]:
                    body, rec = b, {"reach": "ok", "source": "local", "content_type": s.get("content_type")}
                    reused += 1
                    break
        if body is None:
            rec, body = ws.obtain({k: v for k, v in s.items() if k != "kind"})
            fetched += 1
        try:
            if body is None:
                why = rec.get("reach") or "取れない"
                if why == "cc_error":  # 一時の失敗かもしれない：記録せず、次の回で取り直す。多ければ止める
                    guard.add(True, str(rec.get("error")))
                    continue
                # truncated（robots・ログイン・CAPTCHA・HTTP の断り）・digest_mismatch・empty：その書類はやめる。回り道をしない
                mapping[s["id"]] = {"skip": why, "live": rec.get("live")}
                guard.add(False)
            else:
                h = hashlib.sha256(body).hexdigest()
                if s.get("sha256") and h != s["sha256"]:
                    mapping[s["id"]] = {"skip": "中身が変わった"}
                else:
                    p = dd / f"{s['id']}{ws.suffix_for({**s, **rec})}"
                    p.write_bytes(body)
                    mapping[s["id"]] = {"path": str(p), "sha256": h, "source": rec.get("source"), "live": rec.get("live"), "same_as_crawl": rec.get("same_as_crawl")}
                    meta[str(p)] = {"content_type": rec.get("live_content_type") or rec.get("content_type") or s.get("content_type")}
                guard.add(False)
        except MassFailure as e:
            atomic_json(mp_path, mapping)
            atomic_json(meta_path, meta)
            (out / "ABORTED").write_text(f"取るのを止めた：{e}\n", encoding="utf-8")
            print(f"取るのを止めた：{e}", file=sys.stderr)
            return EXIT_ABORT
        atomic_json(mp_path, mapping)
        atomic_json(meta_path, meta)
    left = sum(1 for s in docs if s["id"] not in mapping)
    print(f"書類 {len(docs)}・取った {fetched}・手元の写し {reused}・やめた {sum(1 for m in mapping.values() if m.get('skip'))}・残り {left}")
    return 0 if not left else 3


# ---------------------------------------------------------------------- 比べ
def _run_child(cmd: list, out: Path, log: Path) -> int:
    """空きを見張りながら動かす。空きが足りなければ止めて EXIT_DISK。"""
    with log.open("a", encoding="utf-8") as f:
        p = subprocess.Popen(cmd, stdout=f, stderr=subprocess.STDOUT, env=env_lock(out), start_new_session=True)
        while True:
            try:
                return p.wait(timeout=15)
            except subprocess.TimeoutExpired:
                if free_gb(out) < MIN_FREE_GB:
                    try:
                        os.killpg(p.pid, signal.SIGTERM)
                    except (ProcessLookupError, PermissionError):
                        pass
                    p.wait()
                    return EXIT_DISK


def cmd_score(a) -> int:
    out = guard_out(Path(a.out))
    bpy, tpy = pythons(a)
    src = out / "tool" / "src"
    if not (out / "PREFLIGHT_OK").exists():
        print("先に preflight が要ります", file=sys.stderr)
        return EXIT_PREFLIGHT
    if others_running() and not a.ignore_busy:
        print("ほかの重い処理が動いている：\n  " + "\n  ".join(others_running()), file=sys.stderr)
        return EXIT_BUSY
    mapping = load_json(out / "docs_map.json") or {}
    files = [m["path"] for m in mapping.values() if m.get("path") and kind_of(m["path"])]
    if not files:
        print("書類がない", file=sys.stderr)
        return 2
    lst_dir = out / "lists"
    lst_dir.mkdir(exist_ok=True)
    logs = out / "logs"
    logs.mkdir(exist_ok=True)
    common = ["--baseline-python", bpy, "--todoku-python", tpy, "--todoku-src", str(src), "--doc-limit-s", str(a.doc_limit_s), "--purge"]
    if a.pkr:
        common += ["--pkr", a.pkr]
    t0 = time.time()
    (out / "read").mkdir(exist_ok=True)
    (out / "write").mkdir(exist_ok=True)
    done_write = lambda p: (out / "write" / "scores" / f"{doc_id(p)}.json").exists() or kind_of(p) not in WRITE_KINDS  # noqa: E731
    done_read = lambda p: (out / "read" / "scores" / f"{doc_id(p)}.json").exists()  # noqa: E731
    todo = [p for p in files if not (done_read(p) and done_write(p))]
    # 処理の順は subset の順（結果によらない）。時間の上限で止まっても、残りは無作為な一部
    for i in range(0, len(todo), CHUNK):
        if a.budget_hours and (time.time() - t0) / 3600 >= a.budget_hours:
            print(f"時間の上限 {a.budget_hours} 時間：{i}/{len(todo)} 件で止める。ここまでで数を出せる")
            atomic_json(out / "BUDGET_STOP", {"at": time.strftime("%F %T"), "done_this_run": i, "todo_this_run": len(todo)})
            return EXIT_BUDGET
        chunk = todo[i: i + CHUNK]
        lst = lst_dir / f"chunk_{i:05d}.txt"
        lst.write_text("\n".join(chunk) + "\n", encoding="utf-8")
        args = [*common, "--files", str(lst)]
        rc = _run_child([sys.executable, "-I", str(HERE / "read_compare.py"), "--out", str(out / "read"), "--meta", str(out / "meta.json"), *args], out, logs / "read.log")
        if rc == 0:
            rc = _run_child([sys.executable, "-I", str(HERE / "write_compare.py"), "--out", str(out / "write"), "--no-signed", *args], out, logs / "write.log")
        if rc:
            print(f"比べが止まった（rc={rc}）。もう一度動かすと続きから。記録：{logs}", file=sys.stderr)
            return rc
        print(f"[{min(i + CHUNK, len(todo))}/{len(todo)}] {round((time.time() - t0) / 60)} 分", flush=True)
    return 0


# ---------------------------------------------------------------------- 確かめ
def verify_state(out: Path) -> tuple[list[str], dict]:
    mapping = load_json(out / "docs_map.json") or {}
    probs: list[str] = []
    if (out / "ABORTED").exists() or (out / "read" / "ABORTED").exists() or (out / "write" / "ABORTED").exists():
        probs.append("大量の失敗で止まった印（ABORTED）がある")
    files = {i: m["path"] for i, m in mapping.items() if m.get("path") and kind_of(m["path"])}
    st = collections.Counter(docs=len(files), skipped=sum(1 for m in mapping.values() if m.get("skip")))
    opened = collections.Counter()
    for i, p in files.items():
        rd = load_json(out / "read" / "scores" / f"{doc_id(p)}.json")
        if rd is None:
            st["read_missing"] += 1
        else:
            st["read_scored"] += 1
            if rd.get("unscored"):
                st["read_unscored"] += 1
            t = rd.get("tools") or {}
            opened["todoku"] += bool((t.get("todoku") or {}).get("opened"))
            if rd.get("kind") in ("pdf", "xlsx", "docx", "csv"):
                opened["base_n"] += 1
                opened["base"] += any(s.get("opened") for k, s in t.items() if k != "todoku")
            st["read_exact_todoku"] += bool((t.get("todoku") or {}).get("exact_read"))
        if kind_of(p) in WRITE_KINDS:
            wr = load_json(out / "write" / "scores" / f"{doc_id(p)}.json")
            if wr is None:
                st["write_missing"] += 1
            else:
                st["write_scored"] += 1
                st["write_unscored"] += bool(wr.get("unscored"))
                st["write_runs"] += bool(wr.get("runs"))
                st["write_ok_todoku"] += bool((wr.get("runs") or {}).get("todoku:ascii", {}).get("write_ok"))
    n = len(files)
    if not n:
        probs.append("書類がない")
    if st["read_missing"] or st["write_missing"]:
        probs.append(f"採点の記録がない：読む {st['read_missing']}・書く {st['write_missing']}")
    if n and st["read_unscored"] / n > 0.05:
        probs.append(f"読むの未採点が多い：{st['read_unscored']}/{n}")
    if st["read_scored"] and opened["todoku"] / st["read_scored"] < 0.9:
        probs.append(f"とどくが開けた割合が低い：{opened['todoku']}/{st['read_scored']}")
    if opened["base_n"] and opened["base"] / opened["base_n"] < 0.8:
        probs.append(f"比べる道具が開けた割合が低い：{opened['base']}/{opened['base_n']}")
    if st["read_scored"] and not st["read_exact_todoku"]:
        probs.append("正確に読めた書類が 0 件：数え方か道具がおかしい")
    if st["write_scored"] and not st["write_ok_todoku"]:
        probs.append("書けた書類が 0 件：数え方か道具がおかしい")
    return probs, dict(st)


def cmd_verify(a) -> int:
    out = guard_out(Path(a.out))
    probs, st = verify_state(out)
    print(st)
    if probs:
        print("確かめが通らない：\n  " + "\n  ".join(probs), file=sys.stderr)
        (out / "VERIFIED").unlink(missing_ok=True)
        return 1
    atomic_json(out / "VERIFIED", {"at": time.strftime("%F %T"), **st})
    print("確かめ：通った")
    return 0


# ---------------------------------------------------------------------- 数
def one_score(rd: dict, wr: dict) -> dict:
    """1 つの書類の結果（v1 の world_counts.one_score と同じ意味。読めない・未採点は read_ok が偽）。"""
    t = rd.get("tools") or {}
    d = {"read_todoku": bool((t.get("todoku") or {}).get("exact_read")), "read_base": any(s.get("exact_read") for k, s in t.items() if k != "todoku")}
    d["read_ok"] = bool(t) and not rd.get("unscored")
    d["read_todoku_text_layer"] = bool((t.get("todoku") or {}).get("exact_text_layer", d["read_todoku"]))
    d["read_base_text_layer"] = any(s.get("exact_text_layer", s.get("exact_read")) for k, s in t.items() if k != "todoku")
    d["read_todoku_value"] = bool((t.get("todoku") or {}).get("exact_value", d["read_todoku"]))
    d["read_base_value"] = any(s.get("exact_value", s.get("exact_read")) for k, s in t.items() if k != "todoku")
    d["kind"] = rd.get("kind") or wr.get("kind")
    d["encoding_source"] = rd.get("encoding_source")
    d["read_grade"] = (t.get("todoku") or {}).get("read_grade")
    d["write_grade"] = (t.get("todoku") or {}).get("write_grade")
    d["write_unscored"] = bool(wr.get("unscored"))
    runs = wr.get("runs") or {}
    # 前の定義に近い数（文字の中の空欄を選ばない・断られて選び直さない）。比べる道具も同じ書類だけ
    tgt = wr.get("target") or {}
    prev_scope = bool(runs) and not tgt.get("inline") and not wr.get("pick_refused")
    if runs:  # 書く比べは英数字の値で数える（注釈でも書ける形：比べる道具に有利）
        d["write_todoku"] = bool((runs.get("todoku:ascii") or {}).get("write_ok"))
        d["write_base"] = any(r.get("write_ok") for k, r in runs.items() if k.endswith(":ascii") and not k.startswith("todoku"))
        d["write_todoku_prev"] = prev_scope and d["write_todoku"]
        d["write_base_prev"] = prev_scope and d["write_base"]
        d["write_unverified"] = bool((runs.get("todoku:ascii") or {}).get("unverified"))
        # 感度（W5 を入れない）：消えたのが数式の保存された値だけなら通す（Excel で開けば計算し直される）
        d["write_base_no_w5"] = d["write_base"] or any(
            (r.get("check") or {}).get("landed") and set((r.get("check") or {}).get("lost") or {}) == {"formula_cached_values"} and not (r.get("check") or {}).get("error")
            for k, r in runs.items() if k.endswith(":ascii") and not k.startswith("todoku"))
    return d


def ratio(rows: list[tuple]) -> tuple:
    """world_portal.ratio と同じ：(重み, 集落, 当たり 0/1) → (割合, 線形化の分散, 件数, 集落の数)。"""
    if not rows:
        return None, 0.0, 0, 0
    sw = sum(w for w, _c, _y in rows)
    p = sum(w * y for w, _c, y in rows) / sw
    by = collections.defaultdict(float)
    for w, c, y in rows:
        by[c] += w * (y - p)
    k = len(by)
    var = (k / (k - 1)) * sum(v * v for v in by.values()) / (sw * sw) if k > 1 else 0.0
    return p, var, len(rows), k


def est(cells: list[tuple]) -> dict:
    e = stratified_estimate(cells)
    if e is None:
        return {"count": None, "ci": [None, None], "p": None, "n": 0, "frame": 0}
    tot, p, (lo, hi) = e["frame"], e["p"], e["ci"]
    return {"count": round(tot * p), "ci": [round(tot * lo), round(tot * hi)], "p": p, "p_ci": e["ci"], "n": e["n"], "n_eff": e["n_eff"], "frame": tot}


READ_YS = {"read_todoku": lambda s: s["read_todoku"], "read_base": lambda s: s["read_base"], "read_new": lambda s: s["read_todoku"] and not s["read_base"], "read_lost": lambda s: s["read_base"] and not s["read_todoku"]}
WRITE_YS = {"write_todoku": lambda s: s.get("write_todoku"), "write_base": lambda s: s.get("write_base"), "write_new": lambda s: s.get("write_todoku") and not s.get("write_base"), "write_lost": lambda s: s.get("write_base") and not s.get("write_todoku")}
SENS = {"read_todoku_text_layer": (lambda s: s.get("read_todoku_text_layer"), "R3 を入れない（字の層だけで決める）"),
        "read_base_text_layer": (lambda s: s.get("read_base_text_layer"), "R3 を入れない（比べる道具）"),
        "read_new_text_layer": (lambda s: s.get("read_todoku_text_layer") and not s.get("read_base_text_layer"), "R3 を入れない（新しく届く）"),
        "read_base_value": (lambda s: s.get("read_base_value"), "表計算で元の値を正しく出した道具も通す（比べる道具）"),
        "read_new_value": (lambda s: s.get("read_todoku_value") and not s.get("read_base_value"), "新しく届く：表計算で元の値を正しく出した比べる道具も通す（画面の書式の差を当たりに数えない）"),
        "read_lost_value": (lambda s: s.get("read_base_value") and not s.get("read_todoku_value"), "とどくだけ通らない：同じく元の値も通す"),
        "read_new_with_baseline_format": (lambda s: s["read_todoku"] and not s["read_base"] and s.get("kind") not in NO_BASELINE_KINDS, "新しく届く：比べる道具に読む手段のある形式だけ（.doc・.odt・.ods・.rtf を当たりに数えない）"),
        "write_new_with_baseline_format": (lambda s: s.get("write_todoku") and not s.get("write_base") and s.get("kind") in ("xlsx", "docx", "pdf"), "新しく届く（書く）：比べる道具に書く手段のある形式だけ（.xls・.doc を当たりに数えない）"),
        "write_base_no_w5": (lambda s: s.get("write_base_no_w5"), "比べる道具：数式の保存された値が消えただけなら書けたに数える（W5 を入れない）"),
        "write_new_no_w5": (lambda s: s.get("write_todoku") and not s.get("write_base_no_w5"), "新しく届く（書く）：W5 を入れない"),
        "write_todoku_prev": (lambda s: s.get("write_todoku_prev"), "前の定義に近い：文字の中の空欄を選ばない・断られた所を選び直さない（とどく）"),
        "write_base_prev": (lambda s: s.get("write_base_prev"), "前の定義に近い（比べる道具）"),
        "read_grade_sab": (lambda s: s.get("read_grade") in ("S", "A", "B"), "とどくの読む段階が S・A・B"),
        "write_grade_ab": (lambda s: s.get("write_grade") in ("A", "B"), "とどくの書く段階が A・B")}


def estimate(subset_docs: list[dict], scores: dict, frames: dict, exclude_unscored: bool = True) -> dict:
    """層 × 重み 1/π の比の推定（凍結の world_measured.py と同じ）。書くは、書く比べの対象でない書類・書ける欄のない書類を否に数える（公開した定義）。
    exclude_unscored：未採点の書類を分母から除く（否に数えない・当たりにも数えない）。"""
    parts = {"suffix": frames.get("countries", {}), "wikidata": frames.get("wikidata", {})}
    by = {d["id"]: d for d in subset_docs}
    cell = collections.defaultdict(list)
    for i, s in scores.items():
        if i in by:
            d = by[i]
            cell[(d["country"], d["family"], d.get("domain_source", "suffix"))].append((d, s))

    def run(y, read: bool):
        world, ctry = [], {}
        for cc in sorted({c for p in parts.values() for c in p}):
            cells = []
            for src, p in parts.items():
                for fam, ff in ((p.get(cc) or {}).get("families") or {}).items():
                    if not ff.get("pool"):
                        continue
                    N = ff.get("frame") or ff["pool"]
                    rows = cell.get((cc, fam, src), [])
                    if exclude_unscored:
                        rows = [(d, s) for d, s in rows if (s.get("read_ok") if read else not s.get("write_unscored"))]
                    cells.append((N, *ratio([(1 / d["pi"], str(d["block"]), int(bool(y(s)))) for d, s in rows])))
            world += cells
            e = est(cells)
            if e["count"] is not None:
                ctry[cc] = e
        return est(world), ctry

    res = {"world": {}, "countries": {}, "sensitivity": {}}
    for ys, read in ((READ_YS, True), (WRITE_YS, False)):
        for name, y in ys.items():
            w, c = run(y, read)
            res["world"][name] = w
            for cc, e in c.items():
                res["countries"].setdefault(cc, {})[name] = e
    for name, (y, note) in SENS.items():
        w, _c = run(y, name.startswith("read"))
        res["sensitivity"][name] = {**w, "note": note}
    return res


def portal_fragment(res: dict, version: str) -> dict:
    """ポータルの data.json の final と countries[cc].counts の形。"""
    w = res["world"]

    def nc(k):
        return {"n": w[k]["count"], "ci": w[k]["ci"]}

    final = {"version": version, "read": {"reach": nc("read_todoku"), "new": nc("read_new"), "base": nc("read_base"), "lost": nc("read_lost")},
             "write": {"reach": nc("write_todoku"), "new": nc("write_new"), "base": nc("write_base"), "lost": nc("write_lost")}}
    countries = {cc: {"counts": {"read": {"reach": c["read_todoku"]["count"], "new": c["read_new"]["count"]}, "write": {"reach": c["write_todoku"]["count"], "new": c["write_new"]["count"]}}} for cc, c in res["countries"].items()}
    return {"final": final, "countries": countries}


def film_counts(res: dict, status: str) -> dict:
    """動画の地図の world_counts.json の形（status・countries_surveyed・total・countries）。公開した定義：とどくで正確に書ける文書（欄なし＝否）。"""
    ctry = {cc: c["write_todoku"]["count"] for cc, c in sorted(res["countries"].items())}
    return {"status": status, "countries_surveyed": len(ctry), "total": sum(ctry.values()), "basis": "とどくで正確に書ける文書（欄なし=否）", "world_point": res["world"]["write_todoku"]["count"], "countries": ctry}


def markdown(res: dict, label: str, meta: dict) -> str:
    def fmt(e):
        return "—" if not e or e.get("count") is None else f"{e['count']:,}（{e['ci'][0]:,}〜{e['ci'][1]:,}）n={e['n']}"

    w = res["world"]
    L = [f"# 再評価の数（{label}）", "", f"採点 {SCORER_VERSION}。道具 {meta.get('tool')}。比べた書類 {meta.get('docs')} 件。枠 {meta.get('frame_total'):,} 件のうち測った層 {meta.get('frame_covered'):,} 件。未採点 読む {meta.get('read_unscored')}・書く {meta.get('write_unscored')}。", ""]
    if meta.get("mock"):
        L += ["**試運転（mock）：数に意味はない。公開しない。**", ""]
    L += [md_table(["世界全体", "読む", "書く"], [["正確に（とどく）", fmt(w["read_todoku"]), fmt(w["write_todoku"])], ["比べる道具のどれか", fmt(w["read_base"]), fmt(w["write_base"])],
                                            ["新しく届く", fmt(w["read_new"]), fmt(w["write_new"])], ["とどくだけ通らない", fmt(w["read_lost"]), fmt(w["write_lost"])]]), "", "## 感度の比べ", ""]
    L += [md_table(["", "数", "説明"], [[k, fmt(v), v["note"]] for k, v in res["sensitivity"].items()]), ""]
    return "\n".join(L)


def cmd_estimate(a) -> int:
    out = guard_out(Path(a.out))
    sub = load_json(a.subset)
    mapping = load_json(out / "docs_map.json") or {}
    frames = json.loads((W / "frames.json").read_text(encoding="utf-8"))
    scores = {}
    miss = 0
    for i, m in mapping.items():
        if not m.get("path"):
            continue
        rd = load_json(out / "read" / "scores" / f"{doc_id(m['path'])}.json")
        wr = load_json(out / "write" / "scores" / f"{doc_id(m['path'])}.json") or {}
        if rd is None:
            miss += 1
            continue
        scores[i] = one_score(rd, wr)
    verified = (out / "VERIFIED").exists()
    if not verified and not a.mock:
        print("VERIFIED がないので数を出しません（確かめの前の数を、本物と取り違えないため）", file=sys.stderr)
        return 1
    res = estimate(sub["docs"], scores, frames)
    prov = load_json(out / "tool" / "provenance.json") or {}
    fr_all = sum((ff.get("frame") or ff["pool"]) for part in (frames.get("countries", {}), frames.get("wikidata", {})) for p in part.values() for ff in (p.get("families") or {}).values() if ff.get("pool"))
    meta = {"label": a.label, "mock": bool(a.mock), "scorer": SCORER_VERSION, "tool": prov.get("tag"), "tool_commit": prov.get("commit"), "scorer_commit": prov.get("scorer_commit"),
            "docs": len(scores), "not_scored_yet": miss, "frame_total": fr_all, "frame_covered": res["world"]["read_todoku"]["frame"],
            "read_unscored": sum(1 for s in scores.values() if not s.get("read_ok")), "write_unscored": sum(1 for s in scores.values() if s.get("write_unscored")),
            "budget_stop": (out / "BUDGET_STOP").exists(), "at": time.strftime("%F %T")}
    atomic_json(out / "scores.json", {"mock": bool(a.mock), "scores": scores})
    atomic_json(out / "counts_v2.json", {**meta, **res})
    (out / "counts_v2.md").write_text(markdown(res, a.label, meta), encoding="utf-8")
    pub = out / "publish_candidate"
    pub.mkdir(exist_ok=True)
    version = "mock（公開しない）" if a.mock else (prov.get("tag") or "?")
    atomic_json(pub / "portal_fragment.json", {"_label": version, "mock": bool(a.mock), **portal_fragment(res, version)})
    fc = film_counts(res, "provisional" if a.mock else "final")
    if a.mock:
        fc["mock"] = True
    atomic_json(pub / "film_world_counts.json", fc)
    print(markdown(res, a.label, meta))
    print(f"書いた：{pub}（本物のファイルには書かない）")
    return 0


def cmd_cleanup(a) -> int:
    out = guard_out(Path(a.out))
    if not (out / "VERIFIED").exists() or not (out / "counts_v2.json").exists():
        print("VERIFIED と counts_v2.json がそろうまで、書類を消しません", file=sys.stderr)
        return 1
    for d in ("docs", "read/work", "read/outputs", "write/work", "canary"):
        shutil.rmtree(out / d, ignore_errors=True)
    print("書類と作業のファイルを消した（採点の記録・正解・数は残す）")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["tool", "preflight", "fetch", "score", "verify", "estimate", "cleanup"])
    ap.add_argument("--out", required=True)
    ap.add_argument("--tag")
    ap.add_argument("--subset")
    ap.add_argument("--docs-from")
    ap.add_argument("--only", default="", help="（試運転）この id だけ")
    ap.add_argument("--mock-limit", type=int, default=None)
    ap.add_argument("--mock", action="store_true", help="試運転の印を付ける（公開しない）")
    ap.add_argument("--label", default="")
    ap.add_argument("--baseline-python")
    ap.add_argument("--todoku-python")
    ap.add_argument("--pkr")
    ap.add_argument("--doc-limit-s", type=float, default=900)
    ap.add_argument("--budget-hours", type=float, default=0)
    ap.add_argument("--ignore-busy", action="store_true", help="ほかの重い処理が動いていても動かす（試運転の小さな回だけ）")
    a = ap.parse_args()
    Path(a.out).mkdir(parents=True, exist_ok=True)
    return {"tool": cmd_tool, "preflight": cmd_preflight, "fetch": cmd_fetch, "score": cmd_score, "verify": cmd_verify, "estimate": cmd_estimate, "cleanup": cmd_cleanup}[a.cmd](a)


if __name__ == "__main__":
    raise SystemExit(main())
