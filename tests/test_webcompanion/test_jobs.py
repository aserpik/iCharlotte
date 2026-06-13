"""Tests for webcompanion.jobs — Job model + JobStore persistence."""
import json

from webcompanion import jobs as J
from webcompanion.jobs import Job, JobStore, new_job


def _mk(task_id="summarize_documents"):
    return new_job(task_id, r"E:\cases\1234", "1234", [r"E:\cases\1234\doc.pdf"])


def test_new_job_defaults():
    job = _mk()
    assert job.state == J.QUEUED and job.progress == 0 and len(job.id) == 12


def test_log_capped():
    job = _mk()
    for i in range(250):
        job.add_log(f"line {i}")
    assert len(job.log) == 200 and job.log[-1] == "line 249"


def test_store_roundtrip(tmp_path):
    path = tmp_path / "jobs.json"
    store = JobStore(path)
    job = _mk()
    job.state = J.DONE
    store.add(job)
    store2 = JobStore(path)
    loaded = store2.get(job.id)
    assert loaded is not None and loaded.state == J.DONE
    assert loaded.files == job.files


def test_active_jobs_marked_interrupted_on_load(tmp_path):
    path = tmp_path / "jobs.json"
    store = JobStore(path)
    running, queued, awaiting = _mk(), _mk(), _mk()
    running.state = J.RUNNING
    queued.state = J.QUEUED
    awaiting.state = J.AWAITING_INPUT
    awaiting.session_path = r"C:\logs\s.json"
    for j in (running, queued, awaiting):
        store.add(j)
    store2 = JobStore(path)
    assert store2.get(running.id).state == J.INTERRUPTED
    assert store2.get(queued.id).state == J.INTERRUPTED
    # awaiting survives restarts — session file is on disk, phase 2 can run
    assert store2.get(awaiting.id).state == J.AWAITING_INPUT


def test_all_sorted_newest_first(tmp_path):
    store = JobStore(tmp_path / "jobs.json")
    a, b = _mk(), _mk()
    a.created_at, b.created_at = 100.0, 200.0
    store.add(a)
    store.add(b)
    assert [j.id for j in store.all()] == [b.id, a.id]


def test_corrupt_file_tolerated(tmp_path):
    path = tmp_path / "jobs.json"
    path.write_text("{not json", encoding="utf-8")
    store = JobStore(path)  # must not raise
    assert store.all() == []
