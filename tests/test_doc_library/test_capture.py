from icharlotte_core.doc_library.capture import (
    AUTO_CAPTURE_TASK_IDS, capture_from_task_entry)
from icharlotte_core.doc_library.library import DocumentLibrary


def test_allow_list_contents():
    assert "summarize_documents" in AUTO_CAPTURE_TASK_IDS
    assert "summarize_discovery" in AUTO_CAPTURE_TASK_IDS
    assert "summarize_depositions" in AUTO_CAPTURE_TASK_IDS
    assert "medical_records" in AUTO_CAPTURE_TASK_IDS
    assert "med_chron_analysis" in AUTO_CAPTURE_TASK_IDS
    assert "med_record_extractor" in AUTO_CAPTURE_TASK_IDS
    assert "oppose_motion" not in AUTO_CAPTURE_TASK_IDS
    assert "chat" not in AUTO_CAPTURE_TASK_IDS


def test_non_allowed_task_is_skipped(tmp_path):
    entry = {"task_id": "oppose_motion", "files": [], "settings": {}}
    assert capture_from_task_entry(str(tmp_path), entry) is None
    assert DocumentLibrary(str(tmp_path)).list_entries() == []


def test_allowed_task_captures(tmp_path):
    src = tmp_path / "depo.pdf"
    src.write_bytes(b"bytes")
    from icharlotte_core.doc_library.extract import Extracted
    entry = {"task_id": "summarize_depositions", "files": [str(src)],
             "settings": {"party": "Plaintiff"}}
    result = capture_from_task_entry(
        str(tmp_path), entry,
        extractor=lambda p: Extracted("T", 1, "pdf_native", None))
    assert result is not None
    assert DocumentLibrary(str(tmp_path)).list_entries()[0].label \
        == "Plaintiff's Deposition Transcript"


def test_no_files_is_skipped(tmp_path):
    entry = {"task_id": "summarize_documents", "files": [], "settings": {}}
    assert capture_from_task_entry(str(tmp_path), entry) is None


def test_deposition_label_uses_deponent_name_from_output(tmp_path):
    # No session match for this tmp transcript, so it falls back to parsing the
    # "Deposition of <name>.docx" output filename, title-cased for display.
    from icharlotte_core.doc_library.extract import Extracted
    src = tmp_path / "Santos Depo.pdf"
    src.write_bytes(b"bytes")
    out = str(tmp_path / "Deposition of JOE SMITH.docx")
    entry = {"task_id": "summarize_depositions", "files": [str(src)],
             "settings": {}, "output_path": out, "output_paths": [out]}
    capture_from_task_entry(
        str(tmp_path), entry,
        extractor=lambda p: Extracted("T", 1, "pdf_native", None))
    labels = [e.label for e in DocumentLibrary(str(tmp_path)).list_entries()]
    assert "Deposition of Joe Smith" in labels
