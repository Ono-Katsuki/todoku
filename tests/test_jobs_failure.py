"""長い仕事が落ちたとき：優先度を下げられないことでは落ちない。異常終了を「動いている」と返し続けず、失敗として返す。"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

import pytest

from todoku import api, jobs, mcp_server

SRC = str(Path(jobs.__file__).resolve().parent.parent)


@pytest.fixture
def jd(tmp_path, monkeypatch):
    d = tmp_path / "jobs"
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(d))
    jobs._purged.clear()
    return d


def _csv(tmp_path) -> Path:
    p = tmp_path / "a.csv"
    p.write_text("名前,値\n" + "".join(f"利用者{i},{i}\n" for i in range(50)), encoding="utf-8")
    return p


def _broken_python(tmp_path, monkeypatch, text: str = "RuntimeError: だめでした") -> None:
    """仕事のプロセスが、すぐに異常終了するようにする（本物のプロセスの終わり方で試す）。"""
    fake = tmp_path / "broken-python"
    fake.write_text(f"#!/bin/sh\ncat > /dev/null\necho 'Traceback' >&2\necho '{text}' >&2\nexit 3\n", encoding="utf-8")
    fake.chmod(0o755)
    monkeypatch.setattr(jobs.sys, "executable", str(fake))


def _wait_dead(rec: dict) -> None:
    t0 = time.time()
    while jobs._alive(rec.get("pid")) and time.time() - t0 < 20:
        time.sleep(0.02)


def test_job_finishes_when_priority_cannot_be_lowered(jd, tmp_path):
    """優先度がすでに低いプロセスから（macOS では os.nice が PermissionError）：検査は最後まで進む。"""
    p = _csv(tmp_path)
    script = tmp_path / "run.py"
    script.write_text(
        "import json, sys, time\n"
        "from todoku import api\n"
        "au = api.audit(sys.argv[1], job=True)\n"
        "t0 = time.time()\n"
        "while au['job']['running'] and time.time() - t0 < 60:\n"
        "    time.sleep(0.1)\n"
        "    au = api.audit(sys.argv[1], job=True)\n"
        "print(json.dumps({'finished': au['job']['finished'], 'ok': au['ok']}))\n", encoding="utf-8")
    env = {**os.environ, "PYTHONPATH": SRC, "TODOKU_JOBS_DIR": str(jd)}
    r = subprocess.run(["nice", "-n", "15", sys.executable, str(script), str(p)], env=env, capture_output=True, text=True, timeout=120)
    assert r.returncode == 0, r.stderr
    assert json.loads(r.stdout.strip().splitlines()[-1]) == {"finished": True, "ok": True}


def test_lowering_priority_never_raises(monkeypatch):
    def refuse(_n):
        raise PermissionError(1, "Operation not permitted")

    monkeypatch.setattr(jobs.os, "nice", refuse)
    jobs._lower_priority()


def test_crashing_job_is_reported_as_failed_and_not_started_forever(jd, tmp_path, monkeypatch):
    p = _csv(tmp_path)
    _broken_python(tmp_path, monkeypatch)
    pids = []
    au = None
    for _ in range(jobs.MAX_TRIES + 3):
        au = api.audit(str(p), job=True)
        rec = jobs.status(str(p))
        if rec.get("pid") not in pids:
            pids.append(rec["pid"])
        if not au["job"]["running"]:
            break
        _wait_dead(rec)
    assert au["job"]["running"] is False and au["job"]["finished"] is False
    assert au["job"]["failed"] is True and au["ok"] is False
    assert "RuntimeError: だめでした" in au["job"]["stopped"]
    assert len(pids) == jobs.MAX_TRIES  # 同じ失敗で、立ち上げ続けない
    again = api.audit(str(p), job=True)
    assert again["job"]["failed"] is True and jobs.status(str(p))["pid"] == pids[-1]


def test_failed_job_starts_again_when_asked(jd, tmp_path, monkeypatch):
    p = _csv(tmp_path)
    with monkeypatch.context() as m:
        _broken_python(tmp_path, m)
        for _ in range(jobs.MAX_TRIES + 1):
            au = api.audit(str(p), job=True)
            _wait_dead(jobs.status(str(p)))
        assert au["job"]["failed"]
    au = api.audit(str(p), job=True, restart=True)
    t0 = time.time()
    while au["job"]["running"] and time.time() - t0 < 60:
        time.sleep(0.1)
        au = api.audit(str(p), job=True)
    assert au["job"]["finished"] and not au["job"].get("failed") and au["ok"]


def test_stopped_job_that_made_progress_keeps_going(jd, tmp_path, monkeypatch):
    """進んでいる仕事が途中で止まる（PC を閉じた・メモリの上限）のは失敗にしない：続きから進む。"""
    p = _csv(tmp_path)
    _broken_python(tmp_path, monkeypatch)
    key = jobs.job_key(str(p), None)
    for steps in range(1, jobs.MAX_TRIES * 2):
        rec = jobs.load(key) or {"key": key, "path": str(p), "encoding": None}
        rec["steps"] = steps  # 前の立ち上げの間に進んだ
        jobs.save(key, rec)
        rec = jobs.start(str(p))
        assert not rec.get("failed")
        _wait_dead(rec)


def test_job_reports_the_exception_it_caught(jd, tmp_path, monkeypatch):
    """仕事が自分で控えに書いた理由が、立ち上げ直しで消えない。"""
    p = _csv(tmp_path)
    _broken_python(tmp_path, monkeypatch)
    key = jobs.job_key(str(p), None)
    rec = None
    for _ in range(jobs.MAX_TRIES):
        rec = jobs.load(key) or {"key": key, "path": str(p), "encoding": None, "steps": 0}
        rec["stopped"] = "途中で問題が起きました（ValueError）：読めません"
        rec.pop("pid", None)
        jobs.save(key, rec)
        rec = jobs.start(str(p))
        if rec.get("failed"):
            break
        _wait_dead(rec)
    assert rec["failed"] and "ValueError" in rec["stopped"]


def test_mcp_says_the_audit_failed(jd, tmp_path, monkeypatch):
    p = _csv(tmp_path)
    _broken_python(tmp_path, monkeypatch)
    out = None
    for _ in range(jobs.MAX_TRIES + 1):
        out = mcp_server.call_tool("audit_document", {"path": str(p), "job": True})
        _wait_dead(jobs.status(str(p)))
    text = out["content"][0]["text"]
    assert "検査は最後まで進みませんでした" in text and "restart" in text
    assert "もう一度 audit_document を呼ぶと、進み具合" not in text


def test_exit_of_the_job_is_seen_from_another_process(jd, tmp_path, monkeypatch):
    """立ち上げたプロセスが生きたまま仕事が終わっても、ゾンビとして残らず、ほかのプロセスから「動いている」と見えない。"""
    p = _csv(tmp_path)
    _broken_python(tmp_path, monkeypatch)
    pid = jobs.start(str(p))["pid"]
    t0 = time.time()
    while time.time() - t0 < 20:
        try:
            os.kill(pid, 0)
        except OSError:
            return
        time.sleep(0.02)
    pytest.fail("仕事のプロセスがゾンビとして残っています")


def test_window_job_gives_up(jd, tmp_path, monkeypatch):
    p = _csv(tmp_path)
    _broken_python(tmp_path, monkeypatch)
    rec = None
    for _ in range(jobs.MAX_TRIES + 1):
        rec = jobs.start_window(str(p), None)
        if rec.get("stopped"):
            break
        _wait_dead(rec)
    assert "RuntimeError: だめでした" in rec["stopped"]
    assert jobs.window_failed(str(p), None) == rec["stopped"]


def test_calc_job_gives_up_without_progress(jd, tmp_path, monkeypatch):
    p = _csv(tmp_path)
    _broken_python(tmp_path, monkeypatch)
    rec = None
    for _ in range(jobs.MAX_TRIES + 1):
        rec = jobs.start_calc(str(p), None, {}, ["A1"], {})
        if rec.get("failed"):
            break
        _wait_dead(rec)
    assert rec["failed"] and "RuntimeError" in rec["stopped"] and not jobs.calc_running(rec)
