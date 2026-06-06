import os

from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
    REVIEW_FIELDS,
    build_output_summary,
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


def test_normalize_review_value_splits_list_fields():
    assert normalize_review_value("plaintiffs", "Alice\nBob") == ["Alice", "Bob"]
    assert normalize_review_value("causes_of_action", "Negligence; Battery") == [
        "Negligence",
        "Battery",
    ]
    assert normalize_review_value("case_number", " 23STCV00123 ") == "23STCV00123"


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
