"""Endpoint tests for the web companion server."""
import textwrap

import pytest
from fastapi.testclient import TestClient

from webcompanion import cases as cases_mod
from webcompanion import jobs as J
from webcompanion import task_defs as T
from webcompanion.job_manager import JobManager
from webcompanion.jobs import JobStore
from webcompanion.server import create_app


@pytest.fixture
def case_dir(tmp_path):
    root = tmp_path / "case_9999"
    (root / "DISCOVERY").mkdir(parents=True)
    (root / "DISCOVERY" / "resp.pdf").write_bytes(b"x")
    (root / "doc.pdf").write_bytes(b"x")
    return root


@pytest.fixture
def client(tmp_path, case_dir, monkeypatch):
    fake_case = {"file_number": "9999", "plaintiff_last_name": "Smith",
                 "case_path": str(case_dir)}
    monkeypatch.setattr(cases_mod, "list_cases", lambda q="": [fake_case])
    monkeypatch.setattr(cases_mod, "get_case",
                        lambda fn: fake_case if fn == "9999" else None)
    manager = JobManager(JobStore(tmp_path / "jobs.json"))
    app = create_app(manager)
    c = TestClient(app)
    c.manager = manager
    return c


def test_home_lists_cases(client):
    r = client.get("/")
    assert r.status_code == 200 and "9999" in r.text and "Smith" in r.text


def test_case_page_shows_task_cards(client):
    r = client.get("/case/9999")
    assert r.status_code == 200
    assert "Summarize Documents" in r.text and "Depo Prep" in r.text


def test_case_page_404(client):
    assert client.get("/case/0000").status_code == 404


def test_picker_lists_dirs_and_files(client):
    r = client.get("/case/9999/task/summarize_documents")
    assert r.status_code == 200
    assert "DISCOVERY" in r.text and "doc.pdf" in r.text


def test_picker_rejects_traversal(client):
    r = client.get("/case/9999/task/summarize_documents",
                   params={"path": "../.."})
    assert r.status_code == 400


def test_start_requires_files(client):
    r = client.post("/case/9999/task/summarize_documents/start",
                    data={}, follow_redirects=False)
    assert r.status_code == 400


def test_start_submits_job_and_redirects(client, monkeypatch):
    # Don't actually launch a subprocess.
    submitted = {}
    monkeypatch.setattr(client.manager, "submit",
                        lambda job: submitted.setdefault("job", job) or job)
    r = client.post("/case/9999/task/summarize_documents/start",
                    data={"files": ["doc.pdf"]}, follow_redirects=False)
    assert r.status_code == 303
    job = submitted["job"]
    assert job.task_id == "summarize_documents"
    assert job.files[0].endswith("doc.pdf")
    assert r.headers["location"] == f"/job/{job.id}"
