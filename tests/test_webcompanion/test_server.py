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


def _make_job(client, state=J.RUNNING, **kw):
    from webcompanion.jobs import new_job
    job = new_job("summarize_documents", "E:/case", "9999", ["E:/case/d.pdf"])
    job.state = state
    for k, v in kw.items():
        setattr(job, k, v)
    client.manager.store.add(job)
    return job


def test_job_page_renders(client):
    job = _make_job(client, progress=40)
    r = client.get(f"/job/{job.id}")
    assert r.status_code == 200 and "Summarize Documents" in r.text


def test_job_page_404(client):
    assert client.get("/job/nope").status_code == 404


def test_job_state_api(client):
    job = _make_job(client, progress=55)
    job.add_log("hello")
    r = client.get(f"/api/job/{job.id}")
    body = r.json()
    assert body["state"] == "running" and body["progress"] == 55
    assert body["log"][-1] == "hello" and body["has_output"] is False


def test_cancel_route(client):
    job = _make_job(client, state=J.QUEUED)
    r = client.post(f"/job/{job.id}/cancel", follow_redirects=False)
    assert r.status_code == 303
    assert client.manager.store.get(job.id).state == J.CANCELLED


def test_output_download(client, tmp_path):
    docx = tmp_path / "result.docx"
    docx.write_bytes(b"PK fake docx")
    job = _make_job(client, state=J.DONE, output_path=str(docx))
    r = client.get(f"/job/{job.id}/output")
    assert r.status_code == 200
    assert r.headers["content-type"].startswith(
        "application/vnd.openxmlformats")


def test_output_missing_file_404(client):
    job = _make_job(client, state=J.DONE, output_path="E:/nope/gone.docx")
    assert client.get(f"/job/{job.id}/output").status_code == 404


def test_depo_prep_start_shows_settings(client):
    r = client.post("/case/9999/task/depo_prep/start",
                    data={"files": ["doc.pdf"]})
    assert r.status_code == 200
    assert "deponent_name" in r.text and "Lock-down" in r.text


def test_depo_prep_submit_builds_config(client, monkeypatch):
    import json
    from pathlib import Path
    submitted = {}
    monkeypatch.setattr(client.manager, "submit",
                        lambda job: submitted.setdefault("job", job) or job)
    r = client.post("/case/9999/task/depo_prep/submit", data={
        "files": "doc.pdf",
        "deponent_name": "Dr. Jones",
        "deponent_role": "Treating physician",
        "style": "expert",
        "free_text_notes": "Focus on causation.",
        "flag_strategic": "on",
    }, follow_redirects=False)
    assert r.status_code == 303
    job = submitted["job"]
    assert job.task_id == "depo_prep" and len(job.files) == 1
    cfg = json.loads(Path(job.files[0]).read_text(encoding="utf-8"))
    assert cfg["deponent_name"] == "Dr. Jones"
    assert cfg["style"] == "expert"
    assert cfg["per_topic_flags"]["strategic_note"] is True
    assert cfg["per_topic_flags"]["source_facts"] is False
    assert cfg["deponent_sources"][0].endswith("doc.pdf")
    assert cfg["context_sources"] == []


import json as _json


def _awaiting_job(client, tmp_path, task_id, session_data, topics=None):
    from webcompanion.jobs import new_job
    session = tmp_path / "session.json"
    session.write_text(_json.dumps(session_data), encoding="utf-8")
    if topics is not None:
        (tmp_path / "topics.json").write_text(
            _json.dumps({"topics": topics}), encoding="utf-8")
    job = new_job(task_id, "E:/case", "9999", ["E:/case/d.pdf"])
    job.state = J.AWAITING_INPUT
    job.session_path = str(session)
    client.manager.store.add(job)
    return job


def test_awaiting_deposition_form(client, tmp_path):
    job = _awaiting_job(client, tmp_path, "summarize_depositions", {
        "deponent_name": "Dr. Jones", "deponent_type": "expert",
        "deposition_date": "2026-01-15",
        "topics": [{"title": "Background"}, {"title": "Treatment"}],
    })
    r = client.get(f"/job/{job.id}/awaiting")
    assert r.status_code == 200
    assert "Dr. Jones" in r.text and "Background" in r.text


def test_awaiting_deposition_resume(client, tmp_path, monkeypatch):
    job = _awaiting_job(client, tmp_path, "summarize_depositions", {
        "topics": [{"title": "Background"}],
    })
    applied, resumed = {}, []
    monkeypatch.setattr(
        "webcompanion.server.T.apply_deposition_user_config",
        lambda sp, cfg: applied.update(cfg))
    monkeypatch.setattr(client.manager, "resume",
                        lambda jid: resumed.append(jid))
    r = client.post(f"/job/{job.id}/resume", data={
        "topic": ["Background"], "added_topics": "Damages\nPrognosis",
        "bullets": "7", "deponent_label": "Dr. Jones",
        "audience": "pro_defense", "tone": "recitation",
        "cross_check": "on",
    }, follow_redirects=False)
    assert r.status_code == 303 and resumed == [job.id]
    assert applied["selected_topics"] == ["Background"]
    assert applied["added_topics"] == ["Damages", "Prognosis"]
    assert applied["bullets_per_topic"] == 7
    assert applied["cross_check_enabled"] is True
    assert applied["audience"] == "pro_defense"


def test_awaiting_deposition_requires_topic(client, tmp_path):
    job = _awaiting_job(client, tmp_path, "summarize_depositions",
                        {"topics": []})
    r = client.post(f"/job/{job.id}/resume", data={"added_topics": ""})
    assert r.status_code == 400


def test_awaiting_med_chron_form_and_resume(client, tmp_path, monkeypatch):
    job = _awaiting_job(client, tmp_path, "med_chron_analysis", {
        "provider_name": "Kaiser",
        "catalog": [{"id": "gaps", "label": "Treatment gaps"},
                    {"id": "billing", "label": "Billing analysis"}],
    })
    r = client.get(f"/job/{job.id}/awaiting")
    assert "Kaiser" in r.text and "Treatment gaps" in r.text

    applied, resumed = {}, []
    monkeypatch.setattr(
        "webcompanion.server.T.apply_med_chron_user_config",
        lambda sp, cfg: applied.update(cfg))
    monkeypatch.setattr(client.manager, "resume",
                        lambda jid: resumed.append(jid))
    r = client.post(f"/job/{job.id}/resume", data={
        "analysis": ["gaps"],
        "custom_label_1": "IME prep", "custom_instruction_1": "Flag inconsistencies",
    }, follow_redirects=False)
    assert r.status_code == 303 and resumed == [job.id]
    assert applied["selected_catalog_ids"] == ["gaps"]
    assert applied["custom_analyses"] == [
        {"label": "IME prep", "instruction": "Flag inconsistencies",
         "context_files": []}]


def test_awaiting_depo_prep_form_and_resume(client, tmp_path, monkeypatch):
    topics = [{"id": "t01", "title": "Background", "strategic_note": "note",
               "relevant_digest_refs": ["d1"], "default_checked": True,
               "lawyer_added": False}]
    job = _awaiting_job(client, tmp_path, "depo_prep", {}, topics=topics)
    r = client.get(f"/job/{job.id}/awaiting")
    assert "Background" in r.text

    resumed = []
    monkeypatch.setattr(client.manager, "resume",
                        lambda jid: resumed.append(jid))
    r = client.post(f"/job/{job.id}/resume", data={
        "keep_0": "on", "title_0": "Background (edited)", "note_0": "note",
        "new_title_1": "Damages", "new_note_1": "",
    }, follow_redirects=False)
    assert r.status_code == 303 and resumed == [job.id]
    written = _json.loads(
        (tmp_path / "topics.json").read_text(encoding="utf-8"))["topics"]
    assert written[0]["title"] == "Background (edited)"
    assert written[0]["relevant_digest_refs"] == ["d1"]
    assert written[1]["title"] == "Damages" and written[1]["lawyer_added"] is True


def test_awaiting_on_non_awaiting_job_404(client):
    job = _make_job(client, state=J.RUNNING)
    assert client.get(f"/job/{job.id}/awaiting").status_code == 404
