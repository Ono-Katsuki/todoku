"""長い仕事の控え（書類の中身が入る）の消し方と、段階が終わった検査を使うこと。"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from todoku import api, jobs


@pytest.fixture
def jd(tmp_path, monkeypatch):
    d = tmp_path / "jobs"
    monkeypatch.setenv("TODOKU_JOBS_DIR", str(d))
    jobs._purged.clear()
    return d


def _old(p: Path, hours: float) -> None:
    t = time.time() - hours * 3600
    os.utime(p, (t, t))


def test_old_records_of_every_kind_are_purged_and_new_ones_stay(jd, tmp_path, monkeypatch):
    src = tmp_path / "a.txt"
    src.write_text("x")
    d = jobs._dir()
    for name in ("win-1.window", "win-1.json", "calc-1.json", "2.json", "anchor-3.json", "4.5.tmp"):
        (d / name).write_text('{"path": "%s", "finished": 1}' % src)
        _old(d / name, 25)
    (d / "new.json").write_text('{"path": "%s", "finished": 1}' % src)
    gone = jobs.purge()
    assert sorted(gone) == sorted(["win-1.window", "win-1.json", "calc-1.json", "2.json", "anchor-3.json", "4.5.tmp"])
    assert [p.name for p in d.iterdir()] == ["new.json"]


def test_ttl_comes_from_the_environment(jd, tmp_path, monkeypatch):
    src = tmp_path / "a.txt"
    src.write_text("x")
    d = jobs._dir()
    (d / "a.json").write_text('{"path": "%s", "finished": 1}' % src)
    _old(d / "a.json", 3)
    assert jobs.purge() == []  # 既定は 24 時間
    monkeypatch.setenv("TODOKU_JOBS_TTL_HOURS", "2")
    assert jobs.purge() == ["a.json"]
    monkeypatch.setenv("TODOKU_JOBS_TTL_HOURS", "どれくらいか")
    assert jobs.ttl_seconds() == jobs.DEFAULT_TTL_HOURS * 3600


def test_running_job_is_never_purged(jd, tmp_path):
    src = tmp_path / "a.txt"
    src.write_text("x")
    d = jobs._dir()
    (d / "r.json").write_text('{"path": "%s", "pid": %d}' % (src, os.getpid()))
    _old(d / "r.json", 100)
    assert jobs.purge() == []
    assert jobs.purge(everything=True) == []
    (d / "r.json").write_text('{"path": "%s", "pid": %d, "finished": 1}' % (src, os.getpid()))
    assert jobs.purge(everything=True) == ["r.json"]


def test_record_of_a_deleted_file_is_purged_at_once(jd, tmp_path):
    d = jobs._dir()
    (d / "win-g.json").write_text('{"path": "%s", "finished": 1}' % (tmp_path / "gone.xlsx"))
    (d / "win-g.window").write_bytes(b"content")
    assert sorted(jobs.purge()) == ["win-g.json", "win-g.window"]


def test_zero_ttl_deletes_a_window_when_read(jd, tmp_path, monkeypatch):
    import pickle

    src = tmp_path / "a.txt"
    src.write_text("x")
    wf = jobs._window_file(jobs._window_key(str(src), None))
    wf.write_bytes(pickle.dumps(("sha", ("txt", {}, []))))
    assert jobs.window_result(str(src), None, "sha") == ("txt", {}, [])  # 残す設定：残る
    assert wf.exists()
    monkeypatch.setenv("TODOKU_JOBS_TTL_HOURS", "0")
    assert jobs.window_result(str(src), None, "sha") == ("txt", {}, [])
    assert not wf.exists()


def test_zero_ttl_finished_record_goes_after_it_was_returned(jd, tmp_path, monkeypatch):
    monkeypatch.setenv("TODOKU_JOBS_TTL_HOURS", "0")
    src = tmp_path / "a.txt"
    src.write_text("x")
    key = jobs.job_key(str(src), None)
    jobs.save(key, {"key": key, "path": str(src), "finished": 1.0, "state": {"full": {"checks": []}}})
    assert jobs.purge() == []  # 読む前は消さない（読まれなければ 1 時間で）
    assert api._job_view(jobs.load(key), str(src), None, None)["job"]["finished"]
    assert jobs.load(key) is None


def test_cli_forget(jd, tmp_path, capsys):
    from todoku import cli

    src = tmp_path / "a.txt"
    src.write_text("x")
    (jobs._dir() / "x.window").write_bytes(b"content")
    assert cli.main(["forget"]) == 0
    assert "1 個消しました" in capsys.readouterr().out
    assert not (jobs._dir() / "x.window").exists()


# ------------------------------------------------------------------ 段階は、終わった検査を使う
def _embedded(tmp_path, monkeypatch):
    import test_embed_lazy as t
    from todoku import limits

    p = tmp_path / "e.docx"
    p.write_bytes(t._docx(t._big_book(), t._emf()))
    monkeypatch.setattr(limits, "CONTEXT_BUDGET", 2000)
    api._WINDOWS.clear()
    return p


def test_grade_takes_a_finished_audit_for_deferred_embeds(jd, tmp_path, monkeypatch):
    from todoku.embeds.lazy import whole
    from todoku.grade import DEFERRED_EMBEDS

    p = _embedded(tmp_path, monkeypatch)
    g = api.grade(str(p))
    assert any(r["title"] == DEFERRED_EMBEDS for r in g["write"]["reasons"])
    key = jobs.job_key(str(p), None)
    with whole():  # 検査の仕事の中と同じ：後回しにしない
        full = api.grade(str(p))
    assert not any(r["title"] == DEFERRED_EMBEDS for r in full["write"]["reasons"])
    jobs.save(key, {"key": key, "path": str(p), "finished": 1.0, "state": {"full": {"grade": full}}})
    g = api.grade(str(p))
    assert not any(r["title"] == DEFERRED_EMBEDS for r in g["write"]["reasons"]) and g["write"]["grade"] == full["write"]["grade"]
    api._WINDOWS.clear()


def test_grade_keeps_the_deferred_reason_while_the_audit_is_not_finished(jd, tmp_path, monkeypatch):
    from todoku.grade import DEFERRED_EMBEDS

    p = _embedded(tmp_path, monkeypatch)
    key = jobs.job_key(str(p), None)
    jobs.save(key, {"key": key, "path": str(p), "state": {}})
    assert any(r["title"] == DEFERRED_EMBEDS for r in api.grade(str(p))["write"]["reasons"])
    api._WINDOWS.clear()
