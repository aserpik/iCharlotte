"""Tests for webcompanion.job_manager."""
import textwrap
import time

import pytest

from webcompanion import jobs as J
from webcompanion import task_defs as T
from webcompanion.job_manager import JobManager
from webcompanion.jobs import JobStore, new_job
from webcompanion.task_defs import TaskDef


def _wait(cond, timeout=20.0):
    deadline = time.time() + timeout
    while not cond() and time.time() < deadline:
        time.sleep(0.05)
    assert cond(), "condition not met in time"


@pytest.fixture
def fake_task(tmp_path, monkeypatch):
    """Register task 'fake' whose script is a stub we control via env-free args.

    The stub script reads its positional arg (a .txt 'mode file') whose first
    line selects behavior: ok | fail | await.
    """
    script = tmp_path / "fake_script.py"
    script.write_text(textwrap.dedent("""
        import sys
        from pathlib import Path
        arg = sys.argv[-1]
        if sys.argv[1].startswith("--phase=resume"):
            print("PROGRESS:90:phase2")
            print("OUTPUT:" + str(Path(arg).with_name("phase2.docx")))
            sys.exit(0)
        mode = Path(arg).read_text(encoding="utf-8").strip()
        print("PROGRESS:10:working")
        if mode == "fail":
            print("boom")
            sys.exit(2)
        if mode == "await":
            session = Path(arg).with_name("session.json")
            session.write_text("{}", encoding="utf-8")
            print("AWAITING_INPUT:" + str(session))
            sys.exit(0)
        print("OUTPUT:" + str(Path(arg).with_name("result.docx")))
        sys.exit(0)
    """), encoding="utf-8")

    spec = TaskDef(task_id="fake", title="Fake", glyph="F",
                   script_name="UNUSED", description="test task",
                   two_phase=True, phase2_flag="--phase=resume")
    monkeypatch.setitem(T.TASKS, "fake", spec)
    monkeypatch.setattr(T, "script_path", lambda name: str(script))
    # job_manager imported build_* from task_defs; patch there too
    import webcompanion.job_manager as jm
    monkeypatch.setattr(
        jm, "build_phase1_argv",
        lambda task, p: [str(script), *task.phase1_args, p])
    monkeypatch.setattr(
        jm, "build_phase2_argv",
        lambda task, s: [str(script), task.phase2_flag, s])
    return script


def _submit(manager, tmp_path, mode, name="mode.txt"):
    mode_file = tmp_path / name
    mode_file.write_text(mode, encoding="utf-8")
    job = new_job("fake", str(tmp_path), "9999", [str(mode_file)])
    return manager.submit(job)


def test_success_with_explicit_output(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    job = _submit(manager, tmp_path, "ok")
    _wait(lambda: manager.store.get(job.id).state == J.DONE)
    final = manager.store.get(job.id)
    assert final.output_path.endswith("result.docx")
    assert final.progress == 100


def test_failure_marks_failed(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    job = _submit(manager, tmp_path, "fail")
    _wait(lambda: manager.store.get(job.id).state == J.FAILED)
    final = manager.store.get(job.id)
    assert "code 2" in final.error
    assert any("boom" in ln for ln in final.log)


def test_awaiting_then_resume(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    job = _submit(manager, tmp_path, "await")
    _wait(lambda: manager.store.get(job.id).state == J.AWAITING_INPUT)
    mid = manager.store.get(job.id)
    assert mid.session_path.endswith("session.json")
    manager.resume(job.id)
    _wait(lambda: manager.store.get(job.id).state == J.DONE)
    assert manager.store.get(job.id).output_path.endswith("phase2.docx")


def test_per_case_serialization(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    a = _submit(manager, tmp_path, "ok", "a.txt")
    b = _submit(manager, tmp_path, "ok", "b.txt")  # same case_path → queued
    # b must not run while a runs; both eventually done
    _wait(lambda: manager.store.get(a.id).state == J.DONE)
    _wait(lambda: manager.store.get(b.id).state == J.DONE)


def test_two_phase_rejects_multiple_files(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    f1 = tmp_path / "x.txt"; f1.write_text("ok", encoding="utf-8")
    f2 = tmp_path / "y.txt"; f2.write_text("ok", encoding="utf-8")
    with pytest.raises(ValueError):
        manager.submit(new_job("fake", str(tmp_path), "9999",
                               [str(f1), str(f2)]))


def test_cancel_queued_job(tmp_path, fake_task):
    manager = JobManager(JobStore(tmp_path / "jobs.json"), max_concurrent=1)
    a = _submit(manager, tmp_path, "await", "a.txt")   # occupies the slot
    _wait(lambda: manager.store.get(a.id).state == J.AWAITING_INPUT)
    b = _submit(manager, tmp_path, "ok", "b.txt")
    # a is awaiting (not RUNNING) so b may start; cancel a instead
    manager.cancel(a.id)
    assert manager.store.get(a.id).state == J.CANCELLED
    _wait(lambda: manager.store.get(b.id).state == J.DONE)
