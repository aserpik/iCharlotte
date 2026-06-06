import os
from types import SimpleNamespace

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtWidgets import QLineEdit, QPlainTextEdit

from icharlotte_core.ui.wizard.pages import case_intake_docket_page as mod
from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
    CaseIntakeDocketOutputPage,
    CaseIntakeSettingsPage,
    CaseMetadataReviewPage,
    REVIEW_FIELDS,
    TASK_PAGE_DOCKET_STATUS,
    TASK_PAGE_OUTPUT,
    TASK_PAGE_REVIEW,
    build_output_summary,
    build_case_intake_docket_tab,
    find_complaint_candidate,
    find_latest_docket_pdf,
    load_case_metadata,
    normalize_review_value,
    save_reviewed_metadata,
)


class FakeManager:
    def __init__(self, values=None):
        self.values = dict(values or {})
        self.saved = []

    def get_value(self, file_number, key):
        return self.values.get(key)

    def save_variable(self, file_number, key, value, source="agent", auto_tag=True, extra_tags=None):
        self.saved.append({
            "file_number": file_number,
            "key": key,
            "value": value,
            "source": source,
            "auto_tag": auto_tag,
            "extra_tags": list(extra_tags or []),
        })
        self.values[key] = value


class FakeCaseAgentWorker(QObject):
    status = Signal(str)
    progress = Signal(int)
    finished = Signal(str)
    failed = Signal(str)
    cancelled = Signal()

    instances = []

    def __init__(self, script_name, case_path, file_number, *args, **kwargs):
        super().__init__(kwargs.get("parent"))
        self.script_name = script_name
        self.case_path = case_path
        self.file_number = file_number
        self.recent_lines = []
        self.started = False
        FakeCaseAgentWorker.instances.append(self)

    def start(self):
        self.started = True

    def cancel(self):
        self.cancelled.emit()


def _spec():
    return SimpleNamespace(
        task_id="case_intake_docket",
        title="Case Intake & Docket",
        icon_glyph="",
        description="Run complaint intake and docket.",
    )


def _tab(qtbot, monkeypatch, tmp_path):
    FakeCaseAgentWorker.instances = []
    monkeypatch.setattr(mod, "CaseAgentWorker", FakeCaseAgentWorker)
    tab = build_case_intake_docket_tab(_spec(), str(tmp_path), "1234.001", None)
    qtbot.addWidget(tab)
    return tab


def test_normalize_review_value_splits_list_fields():
    assert normalize_review_value("plaintiffs", "Alice\nBob") == ["Alice", "Bob"]
    assert normalize_review_value("causes_of_action", "Negligence; Battery") == [
        "Negligence",
        "Battery",
    ]
    assert normalize_review_value("case_number", " 23STCV00123 ") == "23STCV00123"


def test_normalize_review_value_drops_placeholder_values():
    assert normalize_review_value("case_number", "None") == ""
    assert normalize_review_value("venue_county", " n/a ") == ""
    assert normalize_review_value("plaintiffs", "Alice\nnull\nN/A") == ["Alice"]


def test_load_case_metadata_reads_review_fields():
    manager = FakeManager({
        "case_number": "23STCV00123",
        "venue_county": "Los Angeles",
        "plaintiffs": ["Alice", "Bob"],
    })

    metadata = load_case_metadata("1234.001", manager=manager)

    assert metadata["case_number"] == "23STCV00123"
    assert metadata["venue_county"] == "Los Angeles"
    assert metadata["plaintiffs"] == ["Alice", "Bob"]
    assert set(REVIEW_FIELDS) <= set(metadata)


def test_save_reviewed_metadata_uses_meta_data_tags():
    manager = FakeManager()

    save_reviewed_metadata(
        "1234.001",
        {"case_number": "23STCV00123", "plaintiffs": "Alice\nBob"},
        manager=manager,
    )

    assert manager.saved[0]["key"] == "case_number"
    assert manager.saved[0]["source"] == "wizard_case_intake"
    assert manager.saved[0]["extra_tags"] == ["Meta Data"]
    assert manager.saved[1]["key"] == "plaintiffs"
    assert manager.saved[1]["value"] == ["Alice", "Bob"]


def test_find_latest_docket_pdf_returns_newest(tmp_path):
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)
    older = out_dir / "Docket_2026.01.01.pdf"
    newer = out_dir / "Docket_2026.02.01.pdf"
    older.write_bytes(b"old")
    newer.write_bytes(b"new")
    os.utime(older, (100, 100))
    os.utime(newer, (200, 200))

    assert find_latest_docket_pdf(str(tmp_path)) == str(newer)


def test_find_complaint_candidate_prefers_pleadings(tmp_path):
    pleadings = tmp_path / "PLEADINGS"
    pleadings.mkdir()
    complaint = pleadings / "Complaint.pdf"
    complaint.write_bytes(b"%PDF")

    assert find_complaint_candidate(str(tmp_path)) == str(complaint)


def test_find_complaint_candidate_prefers_complaint_over_newer_summons_and_service(tmp_path):
    pleadings = tmp_path / "PLEADINGS"
    pleadings.mkdir()
    complaint = pleadings / "Complaint.pdf"
    summons = pleadings / "Summons.pdf"
    proof = pleadings / "Proof of Service Summons.pdf"
    service = pleadings / "Service of Summons.pdf"
    complaint.write_bytes(b"%PDF")
    summons.write_bytes(b"%PDF")
    proof.write_bytes(b"%PDF")
    service.write_bytes(b"%PDF")
    os.utime(complaint, (100, 100))
    os.utime(summons, (200, 200))
    os.utime(proof, (300, 300))
    os.utime(service, (400, 400))

    assert find_complaint_candidate(str(tmp_path)) == str(complaint)


def test_find_complaint_candidate_prefers_amended_score_over_newer_generic(tmp_path):
    pleadings = tmp_path / "PLEADINGS"
    pleadings.mkdir()
    generic = pleadings / "Complaint.pdf"
    amended = pleadings / "Second Amended Complaint.pdf"
    generic.write_bytes(b"%PDF")
    amended.write_bytes(b"%PDF")
    os.utime(amended, (100, 100))
    os.utime(generic, (200, 200))

    assert find_complaint_candidate(str(tmp_path)) == str(amended)


def test_build_output_summary_marks_no_docket_pdf_as_partial(tmp_path):
    variables_docx = tmp_path / "NOTES" / "AI OUTPUT" / "variables.docx"
    variables_docx.parent.mkdir(parents=True)
    variables_docx.write_bytes(b"docx")
    manager = FakeManager({
        "trial_date": "2026-09-01",
        "other_hearings": "CMC on 2026-07-01",
        "procedural_history": "Complaint filed.",
    })

    summary = build_output_summary(
        str(tmp_path),
        "1234.001",
        manager=manager,
        master_db=None,
        recent_lines=["Venue 'ventura' is not supported. Skipping download."],
        success=True,
    )

    assert summary["success"] is True
    assert summary["state"] == "partial"
    assert "no docket pdf" in summary["warning"].lower()
    assert summary["docket_pdf"] == ""
    assert summary["variables_docx"] == str(variables_docx)
    assert "No docket PDF was found" in summary["status"]
    assert summary["trial_date"] == "2026-09-01"


def test_build_output_summary_marks_success_when_docket_pdf_exists(tmp_path):
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)
    docket_pdf = out_dir / "Docket_2026.02.01.pdf"
    docket_pdf.write_bytes(b"%PDF")

    summary = build_output_summary(
        str(tmp_path),
        "1234.001",
        manager=FakeManager(),
        master_db=None,
        success=True,
    )

    assert summary["success"] is True
    assert summary["state"] == "success"
    assert summary["warning"] == ""
    assert summary["docket_pdf"] == str(docket_pdf)


def test_build_output_summary_marks_failed_when_process_fails(tmp_path):
    summary = build_output_summary(
        str(tmp_path),
        "1234.001",
        manager=FakeManager(),
        master_db=None,
        recent_lines=["scraper traceback"],
        success=False,
    )

    assert summary["success"] is False
    assert summary["state"] == "failed"
    assert "failed" in summary["warning"]
    assert summary["recent_lines"] == ["scraper traceback"]


def test_build_output_summary_failure_does_not_expose_stale_docket_pdf(tmp_path):
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)
    stale_docket = out_dir / "Docket_2026.02.01.pdf"
    variables_docx = out_dir / "variables.docx"
    stale_docket.write_bytes(b"%PDF")
    variables_docx.write_bytes(b"docx")

    summary = build_output_summary(
        str(tmp_path),
        "1234.001",
        manager=FakeManager(),
        master_db=None,
        recent_lines=["current run failed"],
        success=False,
    )

    assert summary["success"] is False
    assert summary["state"] == "failed"
    assert summary["docket_pdf"] == ""
    assert summary["variables_docx"] == str(variables_docx)
    assert "failed" in summary["status"].lower()


def test_review_page_round_trips_metadata_and_complaint_file(qtbot):
    page = CaseMetadataReviewPage()
    qtbot.addWidget(page)
    complaint = r"C:\cases\pleadings\Complaint.pdf"

    page.load_metadata(
        {
            "case_number": "23STCV00123",
            "venue_county": "Los Angeles",
            "plaintiffs": ["Alice Smith", "Bob Smith"],
            "defendants": "Acme Corp\nDriver Doe",
            "causes_of_action": "Negligence; Premises Liability",
        },
        complaint_file=complaint,
    )

    assert set(REVIEW_FIELDS) == set(page._field_widgets)
    assert isinstance(page._field_widgets["case_number"], QLineEdit)
    assert isinstance(page._field_widgets["plaintiffs"], QPlainTextEdit)

    payload = page.to_dict()
    assert payload["case_number"] == "23STCV00123"
    assert payload["venue_county"] == "Los Angeles"
    assert payload["plaintiffs"] == ["Alice Smith", "Bob Smith"]
    assert payload["defendants"] == ["Acme Corp", "Driver Doe"]
    assert payload["causes_of_action"] == ["Negligence", "Premises Liability"]
    assert payload["complaint_file"] == complaint

    restored = CaseMetadataReviewPage()
    qtbot.addWidget(restored)
    restored.from_dict(payload)

    assert restored.to_dict() == payload


def test_review_page_requires_case_number_and_venue_county(qtbot):
    page = CaseMetadataReviewPage()
    qtbot.addWidget(page)

    assert page.run_docket_btn.isEnabled() is False

    page._field_widgets["case_number"].setText("23STCV00123")
    assert page.run_docket_btn.isEnabled() is False

    page._field_widgets["venue_county"].setText("Los Angeles")
    assert page.run_docket_btn.isEnabled() is True

    page._field_widgets["case_number"].clear()
    assert page.run_docket_btn.isEnabled() is False


def test_review_page_placeholders_do_not_enable_run_docket(qtbot):
    page = CaseMetadataReviewPage()
    qtbot.addWidget(page)

    page._field_widgets["case_number"].setText("None")
    page._field_widgets["venue_county"].setText("n/a")

    assert page.run_docket_btn.isEnabled() is False


def test_run_docket_click_emits_normalized_metadata(qtbot):
    page = CaseMetadataReviewPage()
    qtbot.addWidget(page)
    page.load_metadata(
        {
            "case_number": " 23STCV00123 ",
            "venue_county": " Los Angeles ",
            "plaintiffs": " Alice Smith \n Bob Smith ",
            "defendants": "Acme Corp",
            "causes_of_action": "Negligence; Battery",
        },
        complaint_file="Complaint.pdf",
    )

    with qtbot.waitSignal(page.run_docket_requested, timeout=500) as blocker:
        qtbot.mouseClick(page.run_docket_btn, Qt.MouseButton.LeftButton)

    payload = blocker.args[0]
    assert payload["case_number"] == "23STCV00123"
    assert payload["venue_county"] == "Los Angeles"
    assert payload["plaintiffs"] == ["Alice Smith", "Bob Smith"]
    assert payload["defendants"] == ["Acme Corp"]
    assert payload["causes_of_action"] == ["Negligence", "Battery"]
    assert payload["complaint_file"] == "Complaint.pdf"


def test_output_page_shows_summary_output_path_warning_and_recent_lines(qtbot):
    page = CaseIntakeDocketOutputPage()
    qtbot.addWidget(page)
    summary = {
        "state": "partial",
        "status": "Docket finished with no PDF.",
        "warning": "No docket PDF was found.",
        "docket_pdf": "",
        "variables_docx": r"C:\cases\NOTES\AI OUTPUT\variables.docx",
        "trial_date": "2026-09-01",
        "other_hearings": "CMC on 2026-07-01",
        "procedural_history": "Complaint filed.",
        "recent_lines": ["Venue unsupported.", "Skipped docket download."],
    }

    page.show_summary(summary)

    assert page.output_path == summary["variables_docx"]
    assert page.summary == summary
    text = page.summary_view.toPlainText()
    assert "partial" in text
    assert summary["variables_docx"] in text
    assert "No docket PDF was found." in text
    assert "Venue unsupported." in text
    assert "Skipped docket download." in text

    summary["status"] = "mutated"
    assert page.summary["status"] == "Docket finished with no PDF."


def test_output_page_load_output_routes_variables_docx_to_variables(qtbot, tmp_path):
    page = CaseIntakeDocketOutputPage()
    qtbot.addWidget(page)
    variables_docx = tmp_path / "NOTES" / "AI OUTPUT" / "variables.docx"
    variables_docx.parent.mkdir(parents=True)
    variables_docx.write_bytes(b"docx")

    page.show_summary({
        "state": "partial",
        "status": "No docket PDF was found.",
        "docket_pdf": "",
        "variables_docx": "",
    })
    page.load_output(str(variables_docx))

    summary = page.summary
    text = page.summary_view.toPlainText()
    assert summary["variables_docx"] == str(variables_docx)
    assert summary.get("docket_pdf", "") != str(variables_docx)
    assert page.output_path == str(variables_docx)
    assert f"Docket PDF: (not found)" in text
    assert f"Variables: {variables_docx}" in text


def test_settings_page_disables_run_without_file_number_and_emits_when_present(qtbot):
    empty_page = CaseIntakeSettingsPage("")
    qtbot.addWidget(empty_page)
    assert empty_page.run_complaint_btn.isEnabled() is False
    assert empty_page.to_dict() == {}
    empty_page.from_dict({"ignored": True})

    page = CaseIntakeSettingsPage("1234.001")
    qtbot.addWidget(page)
    assert page.run_complaint_btn.isEnabled() is True

    with qtbot.waitSignal(page.run_complaint_requested, timeout=500):
        qtbot.mouseClick(page.run_complaint_btn, Qt.MouseButton.LeftButton)


def test_task_tab_runs_complaint_then_loads_review_metadata(qtbot, monkeypatch, tmp_path):
    complaint = str(tmp_path / "PLEADINGS" / "Complaint.pdf")
    metadata = {
        "case_number": "23STCV00123",
        "venue_county": "Los Angeles",
        "plaintiffs": ["Alice Smith"],
    }
    monkeypatch.setattr(mod, "load_case_metadata", lambda file_number: dict(metadata))
    monkeypatch.setattr(mod, "find_complaint_candidate", lambda case_path: complaint)
    tab = _tab(qtbot, monkeypatch, tmp_path)

    tab.settings_page.run_complaint_requested.emit()

    assert len(FakeCaseAgentWorker.instances) == 1
    worker = FakeCaseAgentWorker.instances[0]
    assert worker.script_name == "complaint.py"
    assert worker.case_path == str(tmp_path)
    assert worker.file_number == "1234.001"
    assert worker.started is True

    worker.finished.emit("")

    assert tab.currentIndex() == TASK_PAGE_REVIEW
    payload = tab.review_page.to_dict()
    assert payload["case_number"] == "23STCV00123"
    assert payload["venue_county"] == "Los Angeles"
    assert payload["plaintiffs"] == ["Alice Smith"]
    assert payload["complaint_file"] == complaint
    assert tab._worker is None


def test_task_tab_saves_reviewed_metadata_and_starts_docket(qtbot, monkeypatch, tmp_path):
    saved = []
    monkeypatch.setattr(
        mod,
        "save_reviewed_metadata",
        lambda file_number, metadata: saved.append((file_number, dict(metadata))),
    )
    tab = _tab(qtbot, monkeypatch, tmp_path)
    tab.load_review_state({
        "case_number": "23STCV00123",
        "venue_county": "Los Angeles",
        "plaintiffs": ["Alice Smith"],
        "complaint_file": "Complaint.pdf",
    })

    tab.review_page.run_docket_requested.emit(tab.review_page.to_dict())

    assert saved
    assert saved[0][0] == "1234.001"
    assert saved[0][1]["case_number"] == "23STCV00123"
    assert saved[0][1]["venue_county"] == "Los Angeles"
    assert saved[0][1]["complaint_file"] == "Complaint.pdf"
    assert len(FakeCaseAgentWorker.instances) == 1
    worker = FakeCaseAgentWorker.instances[0]
    assert worker.script_name == "docket.py"
    assert worker.started is True
    assert tab.currentIndex() == TASK_PAGE_DOCKET_STATUS


def test_task_tab_docket_finish_shows_output_and_emits_recent_task(
    qtbot, monkeypatch, tmp_path
):
    summary = {
        "success": True,
        "state": "success",
        "status": "Docket PDF ready.",
        "warning": "",
        "docket_pdf": str(tmp_path / "NOTES" / "AI OUTPUT" / "Docket_2026.pdf"),
        "variables_docx": str(tmp_path / "NOTES" / "AI OUTPUT" / "variables.docx"),
        "trial_date": "2026-09-01",
        "other_hearings": "",
        "procedural_history": "",
        "recent_lines": ["done"],
    }

    def fake_summary(case_path, file_number, manager=None, master_db=None, recent_lines=None, success=True):
        assert case_path == str(tmp_path)
        assert file_number == "1234.001"
        assert master_db == "master-db"
        assert recent_lines == ["done"]
        assert success is True
        return dict(summary)

    monkeypatch.setattr(mod, "save_reviewed_metadata", lambda file_number, metadata: None)
    monkeypatch.setattr(mod, "MasterCaseDatabase", lambda: "master-db")
    monkeypatch.setattr(mod, "build_output_summary", fake_summary)
    tab = _tab(qtbot, monkeypatch, tmp_path)
    tab.load_review_state({"case_number": "23STCV00123", "venue_county": "Los Angeles"})
    tab.review_page.run_docket_requested.emit(tab.review_page.to_dict())
    worker = FakeCaseAgentWorker.instances[0]
    worker.recent_lines = ["done"]

    with qtbot.waitSignal(tab.task_completed, timeout=500) as blocker:
        worker.finished.emit("")

    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert tab.output_page.summary == summary
    entry = blocker.args[0]
    assert entry["task_id"] == "case_intake_docket"
    assert entry["settings"] == {}
    assert entry["summary"] == summary
    assert entry["output_path"] == summary["docket_pdf"]
    assert entry["output_paths"] == [summary["docket_pdf"], summary["variables_docx"]]
    assert entry["completed_at"]
    assert tab._worker is None


def test_task_tab_docket_failure_shows_failed_summary_without_stale_docket(
    qtbot, monkeypatch, tmp_path
):
    stale_docket = str(tmp_path / "NOTES" / "AI OUTPUT" / "Docket_stale.pdf")
    summary = {
        "success": False,
        "state": "failed",
        "status": "Docket processing failed. Review the final log lines below.",
        "warning": "Docket processing failed. Review the final log lines below.",
        "docket_pdf": "",
        "variables_docx": str(tmp_path / "NOTES" / "AI OUTPUT" / "variables.docx"),
        "trial_date": "",
        "other_hearings": "",
        "procedural_history": "",
        "recent_lines": ["traceback"],
    }

    def fake_summary(case_path, file_number, manager=None, master_db=None, recent_lines=None, success=True):
        assert success is False
        assert recent_lines == ["traceback"]
        return dict(summary)

    monkeypatch.setattr(mod, "save_reviewed_metadata", lambda file_number, metadata: None)
    monkeypatch.setattr(mod, "MasterCaseDatabase", lambda: "master-db")
    monkeypatch.setattr(mod, "build_output_summary", fake_summary)
    tab = _tab(qtbot, monkeypatch, tmp_path)
    tab.load_review_state({"case_number": "23STCV00123", "venue_county": "Los Angeles"})
    tab.review_page.run_docket_requested.emit(tab.review_page.to_dict())
    worker = FakeCaseAgentWorker.instances[0]
    worker.recent_lines = ["traceback"]

    with qtbot.waitSignal(tab.task_completed, timeout=500) as blocker:
        worker.failed.emit("script failed")

    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert tab.output_page.summary["state"] == "failed"
    assert tab.output_page.summary["docket_pdf"] == ""
    entry = blocker.args[0]
    assert entry["summary"]["state"] == "failed"
    assert entry["output_path"] == summary["variables_docx"]
    assert stale_docket not in entry["output_paths"]
    assert summary["docket_pdf"] not in entry["output_paths"]


def test_task_tab_restore_loads_review_and_output_state(qtbot, monkeypatch, tmp_path):
    tab = _tab(qtbot, monkeypatch, tmp_path)
    metadata = {
        "case_number": "23STCV00123",
        "venue_county": "Los Angeles",
        "plaintiffs": ["Alice Smith"],
        "complaint_file": "Complaint.pdf",
    }

    tab.load_review_state(metadata)

    assert tab.currentIndex() == TASK_PAGE_REVIEW
    assert tab.review_page.to_dict()["case_number"] == "23STCV00123"
    assert tab.review_page.to_dict()["complaint_file"] == "Complaint.pdf"

    summary = {
        "success": False,
        "state": "failed",
        "status": "Docket failed.",
        "warning": "Docket failed.",
        "docket_pdf": "",
        "variables_docx": "variables.docx",
        "recent_lines": ["failed"],
    }
    updated_metadata = dict(metadata, venue_county="Orange")

    tab.load_output_summary(summary, metadata=updated_metadata)

    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert tab.output_page.summary == summary
    assert tab.review_page.to_dict()["venue_county"] == "Orange"
