#!/usr/bin/env python3
"""見ないで記入する：様式ごとに、新しいエージェント（claude -p）に空の作業フォルダを渡して記入させ、別のエージェントが絵で確かめる。

「見ない」の守り（round 8 の欠点：前の値のファイルを指していた）：
- 作業フォルダには様式の写しだけ。cwd もそこ。設定・CLAUDE.md・記憶・セッションの履歴を読まない（--setting-sources ""・自動の記憶を切る）
- 道具は Read・Write（作業フォルダの中だけ）と todoku の MCP だけ。Bash・Web・サブエージェントはない
- MCP も作業フォルダの外のパスを断る（fill_mcp.py）
- 終わったあと、やりとりの記録を調べ、作業フォルダの外に触れた呼び出しがあれば blind_ok=false にする（数えた結果に出す）
- 申請者の情報は固定で、様式ごとに作らない（前の回の値を使わない）。ほかの値は「申請者に聞く」として空ける

数える：記入済み（エージェントの分類＋最後の書き込みの proof_ok・read_back_ok＋確かめ役の verdict ok）、人の手が残ること（印刷・押印・自署・窓口など）、
証明・読み戻し（記録から機械で）、おかしいまま残った所、誤って断った数。

使い方: fill_runner.py run --set dev --root 様式のフォルダ --out 出力 --baseline-python PY --todoku-python PY [--agents 4] [--limit N] [--model opus] [--pkr PATH] [--seed S]
       fill_runner.py summarize --out 出力 [--set dev]
途中で止めても、もう一度動かせば続きから。使用量の上限・認証などで断られたら、そこで止める（回り道をしない）。"""

from __future__ import annotations

import argparse
import collections
import json
import os
import random
import re
import shutil
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from common import ROOT, atomic_json, doc_id, guard_path, lo_profile, load_json, md_table, pct, run_watched, walk_files, wilson  # noqa: E402

FILL_KINDS = {".pdf": 1, ".xlsx": 1, ".xlsm": 1, ".xls": 1, ".docx": 1, ".doc": 1, ".jtd": 1, ".rtf": 1, ".odt": 1, ".ods": 1}
APPLICANT = {
    "ja": """氏名 山田 花子（ヤマダ ハナコ）、昭和60年4月1日生（1985-04-01、41歳）、女
住所 {PLACE}〇〇一丁目2番3号、電話 03-0000-0000
身体障害者手帳 1級（肢体不自由）第000000号、障害支援区分 6
申請日 令和8年11月10日
振込先 架空銀行 本店 普通 0000000 名義 ヤマダ ハナコ
希望するサービス 重度訪問介護
申請者本人が対象者です。""",
    "world": """Name: Hanako Yamada (given name Hanako, family name Yamada), born 1 April 1985, female
Address: 1-2-3 Example Street, in the jurisdiction this form belongs to; phone +00 000 000 0000
Physical disability: wheelchair user, highest support level; disability certificate no. 000000
Date of application: 10 November 2026
Bank: Example Bank, main branch, account 0000000, holder Hanako Yamada
Requested service: personal care assistance
The applicant is the person concerned. Answer in the language of the form.""",
}
STOP_WORDS = re.compile(r"rate.?limit|usage limit|quota|credit balance|authentication|unauthorized|forbidden|not logged in|overloaded_error", re.I)
_stop = threading.Event()
_stop_why = [None]


def forms_of(root: str, limit: int | None, seed: int) -> list[str]:
    guard_path(root)
    fs = [p for p in walk_files(root, FILL_KINDS)]
    if limit:
        random.Random(seed).shuffle(fs)
        fs = sorted(fs[:limit])
    return fs


def claude_cmd(prompt: str, sandbox: Path, model: str, mcp: dict | None, system_file: Path | None, tools: str, allowed: list[str]) -> list:
    cmd = ["claude", "-p", prompt, "--model", model, "--setting-sources", "", "--strict-mcp-config", "--no-session-persistence",
           "--output-format", "stream-json", "--verbose", "--permission-mode", "dontAsk", "--tools", tools, "--allowedTools", *allowed]
    if mcp is not None:
        cmd += ["--mcp-config", json.dumps(mcp, ensure_ascii=False)]
    if system_file:
        cmd += ["--append-system-prompt-file", str(system_file)]
    return cmd


def run_agent(cmd: list, cwd: Path, transcript: Path, limit_s: float) -> dict:
    env = {"CLAUDE_CODE_DISABLE_AUTO_MEMORY": "1", "DISABLE_AUTOUPDATER": "1"}
    r = run_watched(cmd, limit_s=limit_s, mem_mb=4000, env=env, cwd=cwd, heavy=False, stdout_limit=200_000_000)
    transcript.write_text(r["stdout"], encoding="utf-8")
    end = {}
    for ln in reversed(r["stdout"].splitlines()):
        try:
            d = json.loads(ln)
        except json.JSONDecodeError:
            continue
        if d.get("type") == "result":
            end = d
            break
    bad = end.get("is_error") or (not end and r["returncode"])
    if bad and STOP_WORDS.search((str(end.get("result")) + str(end.get("api_error_status")) + r.get("stderr", ""))[:3000]):
        _stop.set()
        _stop_why[0] = (str(end.get("result")) or r.get("stderr", ""))[:300]
    return {"returncode": r["returncode"], "reason": r["reason"], "seconds": r["seconds"], "peak_mb": r["peak_mb"], "is_error": bool(end.get("is_error")),
            "cost_usd": end.get("total_cost_usd"), "num_turns": end.get("num_turns"), "end_text": str(end.get("result"))[:300]}


def blind_check(transcript: Path, allowed_root: Path, allowed_tools: set[str]) -> dict:
    """やりとりの記録を調べる：作業フォルダの外のパスに触れた道具の呼び出し・許していない道具・読まれた記憶。"""
    root = os.path.realpath(allowed_root)
    outside, tools = [], collections.Counter()
    init = {}
    for ln in transcript.read_text(encoding="utf-8", errors="replace").splitlines():
        try:
            d = json.loads(ln)
        except json.JSONDecodeError:
            continue
        if d.get("type") == "system" and d.get("subtype") == "init":
            init = {"tools": d.get("tools"), "mcp_servers": d.get("mcp_servers"), "memory_paths": d.get("memory_paths"), "skills": d.get("skills"), "model": d.get("model"), "cwd": d.get("cwd")}
        if d.get("type") != "assistant":
            continue
        for c in (d.get("message") or {}).get("content") or []:
            if c.get("type") != "tool_use":
                continue
            tools[c["name"]] += 1
            inp = c.get("input") or {}
            for key in ("file_path", "path", "out_path", "notebook_path"):
                v = inp.get(key)
                if isinstance(v, str):
                    p = os.path.realpath(v.split("!/")[0] if os.path.isabs(v) else os.path.join(root, v))
                    if p != root and not p.startswith(root + os.sep):
                        outside.append(f"{c['name']}:{v}")
    not_allowed = sorted(t for t in tools if t not in allowed_tools and not t.startswith("mcp__todoku__"))
    leaks = [str(x)[:120] for x in (init.get("memory_paths") or {}).values()] if isinstance(init.get("memory_paths"), dict) else []
    return {"blind_ok": not outside and not not_allowed, "outside": outside[:10], "not_allowed_tools": not_allowed, "tools": dict(tools), "init": init, "memory_paths": leaks}


def read_log(path: Path) -> list[dict]:
    if not path.exists():
        return []
    out = []
    for ln in path.read_text(encoding="utf-8").splitlines():
        try:
            out.append(json.loads(ln))
        except json.JSONDecodeError:
            pass
    return out


def fill_one(form: str, a, base: Path, schema: dict, applicant: str) -> dict:
    did = doc_id(form)
    d = base / did
    rec_path = d / "state.json"
    state = load_json(rec_path) or {}
    sandbox = d / "sandbox"
    (sandbox / "form").mkdir(parents=True, exist_ok=True)
    (sandbox / "out").mkdir(exist_ok=True)
    name = Path(form).name
    src = sandbox / "form" / name
    if not src.exists():
        shutil.copy(form, src)
    out_path = sandbox / "out" / name
    record_path = sandbox / "record.json"
    log = d / "tool_log.jsonl"
    state.update({"id": did, "form": form, "file": name})

    if not state.get("fill_done") and not _stop.is_set():
        prompt = (HERE / "fill_prompt.md").read_text(encoding="utf-8").format(FORM_PATH=src, OUT_PATH=out_path, SANDBOX=sandbox, RECORD_PATH=record_path, APPLICANT=applicant,
                                                                               SCHEMA=json.dumps(schema["record"], ensure_ascii=False, indent=1))
        mcp = {"mcpServers": {"todoku": {"command": a.todoku_python, "args": ["-I", str(HERE / "fill_mcp.py"), str(ROOT / "src"), str(sandbox), str(log), str(a.lo_profile)],
                                         "env": {"EVAL_LOCK": os.environ.get("EVAL_LOCK", "")}}}}
        allowed = [f"Read(//{str(sandbox).lstrip('/')}/**)", f"Edit(//{str(sandbox).lstrip('/')}/**)", "mcp__todoku__*"]
        cmd = claude_cmd(prompt, sandbox, a.model, mcp, ROOT / "skills" / "shinsei" / "SKILL.md", "Read,Write", allowed)
        tr = d / "fill_transcript.jsonl"
        for attempt in range(2):
            res = run_agent(cmd, sandbox, tr, a.fill_limit)
            state["fill_run"] = res
            if record_path.exists() or _stop.is_set():
                break
        state["blind"] = blind_check(tr, sandbox, {"Read", "Write"})
        mcp_up = [s.get("status") for s in ((state["blind"].get("init") or {}).get("mcp_servers") or [])]
        if "connected" not in mcp_up:
            # 囲いの失敗で、道具の失敗ではない：数えない。ここで止める（直してから続ける）
            _stop.set()
            _stop_why[0] = f"todoku の MCP サーバーにつながらなかった（{mcp_up}）：{name}"
            record_path.unlink(missing_ok=True)
            state["fill_done"] = False
            atomic_json(rec_path, state)
            return state
        state["fill_done"] = record_path.exists() or (res.get("reason") == "timeout") or res.get("returncode") == 0
        atomic_json(rec_path, state)
    # 確かめ役
    if state.get("fill_done") and not state.get("review_done") and not _stop.is_set():
        review_dir = d / "review"
        review_dir.mkdir(exist_ok=True)
        review_path = review_dir / "review.json"
        if out_path.exists():
            if not (review_dir / "pages.json").exists():
                r = run_watched([a.baseline_python, "-I", HERE / "render_pages.py", src, out_path, review_dir, a.lo_profile, ROOT / "src", a.todoku_python, a.pkr or "-"], limit_s=900, mem_mb=3000)
                state["render_error"] = (r["stdout"][-300:] + r["stderr"][-200:]) if not (review_dir / "pages.json").exists() else None
            calls = read_log(log)
            fills = [c for c in calls if c.get("tool") == "fill_document"]
            written = [c for c in fills if not c.get("is_error") and str(c.get("out_path")) == str(out_path)]
            last = written[-1] if written else {}
            atomic_json(review_dir / "written.json", {"values": last.get("values"), "say_to_user": last.get("say_to_user"), "warnings": last.get("warnings"), "written_rects": last.get("written_rects")})
            if (review_dir / "pages.json").exists():
                prompt = (HERE / "review_prompt.md").read_text(encoding="utf-8").format(REVIEW_DIR=review_dir, REVIEW_PATH=review_path, APPLICANT=applicant,
                                                                                         SCHEMA=json.dumps(schema["review"], ensure_ascii=False, indent=1))
                allowed = [f"Read(//{str(review_dir).lstrip('/')}/**)", f"Edit(//{str(review_dir).lstrip('/')}/**)"]
                cmd = claude_cmd(prompt, review_dir, a.model, None, None, "Read,Write", allowed)
                tr = d / "review_transcript.jsonl"
                state["review_run"] = run_agent(cmd, review_dir, tr, a.review_limit)
                state["review_blind"] = blind_check(tr, review_dir, {"Read", "Write"})
        else:
            state["no_output"] = True
        state["review_done"] = review_path.exists() or not out_path.exists() or (state.get("review_run") or {}).get("reason") == "timeout"
        atomic_json(rec_path, state)
    atomic_json(rec_path, state)
    print(f"{name}: 記入 {state.get('fill_done')} 確かめ {state.get('review_done')} 見ない {(state.get('blind') or {}).get('blind_ok')} "
          f"{(state.get('fill_run') or {}).get('seconds')}s", flush=True)
    return state


def form_record(d: Path) -> dict | None:
    st = load_json(d / "state.json")
    if not st:
        return None
    rec = load_json(d / "sandbox" / "record.json")
    rev = load_json(d / "review" / "review.json")
    calls = read_log(d / "tool_log.jsonl")
    out_path = str(d / "sandbox" / "out" / st["file"])
    fills = [c for c in calls if c.get("tool") == "fill_document"]
    good = [c for c in fills if not c.get("is_error") and str(c.get("out_path")) == out_path]
    last = good[-1] if good else {}
    refused = [c for c in fills if c.get("is_error")]
    audits = [c for c in calls if c.get("tool") == "audit_document"]
    r = {"id": st["id"], "file": st["file"], "form": st["form"], "agent_class": (rec or {}).get("classification"), "record_ok": rec is not None, "review_ok": rev is not None,
         "output": Path(out_path).exists(), "proof_ok": last.get("proof_ok"), "read_back_ok": last.get("read_back_ok"), "fills": len(fills), "fills_ok": len(good),
         "tool_refusals": len(refused), "audited": bool(audits), "denied": sum(1 for c in calls if c.get("denied_path")),
         "blind_ok": (st.get("blind") or {}).get("blind_ok"), "review_blind_ok": (st.get("review_blind") or {}).get("blind_ok"),
         "seconds": (st.get("fill_run") or {}).get("seconds"), "review_seconds": (st.get("review_run") or {}).get("seconds"),
         "cost": ((st.get("fill_run") or {}).get("cost_usd") or 0) + ((st.get("review_run") or {}).get("cost_usd") or 0),
         "tool_seconds": round(sum(c.get("seconds") or 0 for c in calls), 1), "no_output": bool(st.get("no_output")), "render_error": st.get("render_error"),
         "verdict": (rev or {}).get("verdict"), "wrong_left": list((rec or {}).get("wrong_left") or []) , "review_wrong": (rev or {}).get("wrong_left") or [],
         "missed": (rev or {}).get("missed_answerable") or [], "renderer_disagreements": (rev or {}).get("renderer_disagreements") or [],
         "refusals": (rec or {}).get("refusals") or [], "remains": ((rec or {}).get("remains_for_people") or []) + ((rev or {}).get("remains_for_people") or []),
         "ask_applicant": (rec or {}).get("ask_applicant") or [], "wrong_fixed": (rec or {}).get("wrong_fixed") or [], "cannot_judge": (rev or {}).get("cannot_judge") or []}
    r["out_of_scope"] = r["agent_class"] == "対象外"
    r["filled"] = bool(r["agent_class"] == "記入済み" and r["proof_ok"] and r["read_back_ok"] and r["verdict"] == "ok" and not r["wrong_left"])
    r["wrong_refusals"] = sum(1 for x in r["refusals"] if isinstance(x, dict) and x.get("correct") is False)
    return r


def remain_types(r: dict) -> set:
    return {(x.get("type") if isinstance(x, dict) else "その他") or "その他" for x in r["remains"]}


def summarize_set(base: Path) -> dict:
    recs = [x for x in (form_record(d) for d in sorted(base.iterdir()) if d.is_dir()) if x] if base.exists() else []
    scope = [r for r in recs if not r["out_of_scope"]]
    filled = sum(1 for r in scope if r["filled"])
    types = collections.Counter()
    for r in scope:
        types.update(remain_types(r))
    done = [r for r in scope if r["output"]]
    s = {"forms": len(recs), "out_of_scope": len(recs) - len(scope), "in_scope": len(scope), "filled": filled, "filled_ci": wilson(filled, len(scope)),
         "agent_said_filled": sum(1 for r in scope if r["agent_class"] == "記入済み"), "with_output": len(done),
         "proof_ok": sum(1 for r in done if r["proof_ok"]), "read_back_ok": sum(1 for r in done if r["read_back_ok"]), "audited": sum(1 for r in done if r["audited"]),
         "reviewed": sum(1 for r in scope if r["review_ok"]), "review_ok": sum(1 for r in scope if r["verdict"] == "ok"),
         "wrong_left_items_review": sum(len(r["review_wrong"]) for r in scope), "wrong_left_forms_review": sum(1 for r in scope if r["review_wrong"]),
         "wrong_left_items_agent": sum(len(r["wrong_left"]) for r in scope), "missed_answerable": sum(len(r["missed"]) for r in scope),
         "renderer_disagreements": sum(len(r["renderer_disagreements"]) for r in scope), "cannot_judge": sum(len(r["cannot_judge"]) for r in scope),
         "tool_refusals": sum(r["tool_refusals"] for r in scope), "refusals_recorded": sum(len(r["refusals"]) for r in scope), "wrong_refusals": sum(r["wrong_refusals"] for r in scope),
         "remains_types": dict(types), "forms_with_remains": sum(1 for r in scope if r["remains"]), "ask_applicant_items": sum(len(r["ask_applicant"]) for r in scope),
         "blind_ok": sum(1 for r in recs if r["blind_ok"] and r["review_blind_ok"] is not False), "blind_checked": sum(1 for r in recs if r["blind_ok"] is not None),
         "denied_paths": sum(r["denied"] for r in recs), "record_missing": sum(1 for r in recs if not r["record_ok"]), "no_output": sum(1 for r in scope if r["no_output"]),
         "fill_minutes_mean": round(sum(r["seconds"] or 0 for r in recs) / 60 / max(len(recs), 1), 1), "review_minutes_mean": round(sum(r["review_seconds"] or 0 for r in recs) / 60 / max(len(recs), 1), 1),
         "tool_seconds_mean": round(sum(r["tool_seconds"] for r in recs) / max(len(recs), 1), 1), "cost_usd_total": round(sum(r["cost"] for r in recs), 2),
         "not_filled": [{"file": r["file"], "agent": r["agent_class"], "verdict": r["verdict"], "proof": r["proof_ok"], "read_back": r["read_back_ok"],
                         "review_wrong": [str(x.get("what") if isinstance(x, dict) else x)[:140] for x in r["review_wrong"]][:3],
                         "missed": [str(x.get("place") if isinstance(x, dict) else x)[:80] for x in r["missed"]][:3]} for r in scope if not r["filled"]][:60]}
    return s


def markdown(all_sets: dict) -> str:
    names = list(all_sets)
    def row(label, f):
        return [label] + [f(all_sets[n]) for n in names]
    rows = [row("様式（対象外を除く）", lambda s: f"{s['in_scope']}（対象外 {s['out_of_scope']}）"),
            row("記入済み（エージェント＋証明＋確かめ役）", lambda s: f"{pct(s['filled'], s['in_scope'])} 95% {s['filled_ci']}"),
            row("エージェントが記入済みと言った", lambda s: s["agent_said_filled"]),
            row("証明が通った（出力のある様式の）", lambda s: pct(s["proof_ok"], s["with_output"])),
            row("読み戻しが通った", lambda s: pct(s["read_back_ok"], s["with_output"])),
            row("確かめ役が見た／ok", lambda s: f"{s['reviewed']}／{s['review_ok']}"),
            row("おかしいまま残った（確かめ役：所／様式）", lambda s: f"{s['wrong_left_items_review']}／{s['wrong_left_forms_review']}"),
            row("おかしいまま残った（記入者の申告）", lambda s: s["wrong_left_items_agent"]),
            row("答えられるのに空いていた（確かめ役）", lambda s: s["missed_answerable"]),
            row("描き手で見え方が違う", lambda s: s["renderer_disagreements"]),
            row("道具が断った回数／誤って断った", lambda s: f"{s['tool_refusals']}／{s['wrong_refusals']}"),
            row("人の手が残る（種類ごとの様式の数）", lambda s: "・".join(f"{k} {v}" for k, v in sorted(s["remains_types"].items(), key=lambda x: -x[1])) or "なし"),
            row("申請者に聞くこと（所）", lambda s: s["ask_applicant_items"]),
            row("見ないで行えた（記録を調べた）", lambda s: f"{s['blind_ok']}/{s['blind_checked']}"),
            row("1 様式の時間（記入／確かめ。分）", lambda s: f"{s['fill_minutes_mean']}／{s['review_minutes_mean']}"),
            row("費用の合計（USD）", lambda s: s["cost_usd_total"])]
    return "# 見ないでの記入\n\n" + md_table(["", *names], rows) + "\n"


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["run", "summarize"])
    ap.add_argument("--set", default="dev")
    ap.add_argument("--root")
    ap.add_argument("--out", required=True)
    ap.add_argument("--baseline-python")
    ap.add_argument("--todoku-python")
    ap.add_argument("--agents", type=int, default=4)
    ap.add_argument("--limit", type=int)
    ap.add_argument("--seed", type=int, default=20261105)
    ap.add_argument("--model", default="opus")
    ap.add_argument("--pkr", default=None)
    ap.add_argument("--fill-limit", type=float, default=45 * 60)
    ap.add_argument("--review-limit", type=float, default=20 * 60)
    a = ap.parse_args()
    out = Path(a.out)
    if a.cmd == "run":
        base = out / a.set
        base.mkdir(parents=True, exist_ok=True)
        a.lo_profile = out / "lo_profile"
        lo_profile(a.lo_profile)
        schema = json.loads((HERE / "fill_schema.json").read_text(encoding="utf-8"))
        forms = forms_of(a.root, a.limit, a.seed)
        print(f"{a.set}：様式 {len(forms)} 件", flush=True)
        applicant = APPLICANT["world"] if a.set == "world" else APPLICANT["ja"]

        def run(form):
            if _stop.is_set():
                return
            place = "東京都" if a.set != "world" else ""
            fill_one(form, a, base, schema, applicant.replace("{PLACE}", place))

        with ThreadPoolExecutor(max(1, a.agents)) as ex:
            list(ex.map(run, forms))
        if _stop.is_set():
            print(f"断られたので止めた：{_stop_why[0]}\nもう一度動かせば続きから。", flush=True)
            return 3
    sets = sorted(p.name for p in out.iterdir() if p.is_dir() and p.name != "lo_profile") if out.exists() else []
    sets = sorted(sets, key=lambda n: (n != "dev", n))
    allsets = {n: summarize_set(out / n) for n in sets}
    atomic_json(out / "summary.json", allsets)
    (out / "summary.md").write_text(markdown(allsets), encoding="utf-8")
    print(markdown(allsets))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
