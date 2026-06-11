import json
import os
import time

from icharlotte_core.ui.wizard.summary_outputs import (
    discover_summary_outputs,
    task_id_for_summary_action,
)


def _touch(path, *, offset=0):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(b"placeholder docx")
    stamp = time.time() + offset
    os.utime(path, (stamp, stamp))
    return str(path)


def _write_state(case_root, entries):
    state_dir = case_root / "NOTES" / "AI OUTPUT" / ".icharlotte"
    state_dir.mkdir(parents=True, exist_ok=True)
    (state_dir / "wizard_state.json").write_text(
        json.dumps({"version": 1, "open_tabs": [], "recent_tasks": entries}),
        encoding="utf-8",
    )


def test_task_id_for_summary_action_maps_task_browser_cards():
    assert task_id_for_summary_action("open_summarize_documents_outputs") == "summarize_documents"
    assert task_id_for_summary_action("open_summarize_discovery_outputs") == "summarize_discovery"
    assert task_id_for_summary_action("open_summarize_depositions_outputs") == "summarize_depositions"
    assert task_id_for_summary_action("open_medical_records_outputs") == "medical_records"
    assert task_id_for_summary_action("open_separate_index") is None


def test_discovery_combines_wizard_history_and_legacy_task_patterns(tmp_path):
    upper_dir = tmp_path / "NOTES" / "AI OUTPUT"
    mixed_dir = tmp_path / "NOTES" / "AI Output"

    legacy_doc = _touch(upper_dir / "AI_OUTPUT.docx", offset=1)
    legacy_discovery = _touch(upper_dir / "Discovery_Responses_Smith.docx", offset=2)
    legacy_depo = _touch(upper_dir / "Deposition of Jane Roe.docx", offset=3)
    legacy_medical = _touch(upper_dir / "Med_Record_Ortho Clinic.docx", offset=4)

    wizard_doc = _touch(mixed_dir / "AI_OUTPUT - motion.docx", offset=5)
    wizard_depo = _touch(mixed_dir / "AI_OUTPUT - jane transcript.docx", offset=6)
    wizard_medical = _touch(mixed_dir / "AI_OUTPUT - records batch.docx", offset=7)

    _write_state(
        tmp_path,
        [
            {
                "task_id": "summarize_documents",
                "output_paths": [os.path.relpath(wizard_doc, tmp_path)],
            },
            {
                "task_id": "summarize_depositions",
                "output_path": wizard_depo,
            },
            {
                "task_id": "medical_records",
                "output_path": wizard_medical,
            },
        ],
    )

    document_names = [
        os.path.basename(entry.path)
        for entry in discover_summary_outputs(str(tmp_path), "summarize_documents")
    ]
    assert document_names == ["AI_OUTPUT - motion.docx", "AI_OUTPUT.docx"]

    discovery_names = [
        os.path.basename(entry.path)
        for entry in discover_summary_outputs(str(tmp_path), "summarize_discovery")
    ]
    assert discovery_names == ["Discovery_Responses_Smith.docx"]

    deposition_names = [
        os.path.basename(entry.path)
        for entry in discover_summary_outputs(str(tmp_path), "summarize_depositions")
    ]
    assert deposition_names == [
        "AI_OUTPUT - jane transcript.docx",
        "Deposition of Jane Roe.docx",
    ]

    medical_names = [
        os.path.basename(entry.path)
        for entry in discover_summary_outputs(str(tmp_path), "medical_records")
    ]
    assert medical_names == [
        "AI_OUTPUT - records batch.docx",
        "Med_Record_Ortho Clinic.docx",
    ]

    assert legacy_doc in [entry.path for entry in discover_summary_outputs(str(tmp_path), "summarize_documents")]
    assert legacy_discovery not in [entry.path for entry in discover_summary_outputs(str(tmp_path), "summarize_documents")]
    assert legacy_depo not in [entry.path for entry in discover_summary_outputs(str(tmp_path), "summarize_documents")]
    assert legacy_medical not in [entry.path for entry in discover_summary_outputs(str(tmp_path), "summarize_documents")]


def test_discovery_dedupes_wizard_and_legacy_matches(tmp_path):
    output_path = _touch(tmp_path / "NOTES" / "AI OUTPUT" / "Discovery_Responses_Smith.docx")
    _write_state(
        tmp_path,
        [{"task_id": "summarize_discovery", "output_paths": [output_path]}],
    )

    outputs = discover_summary_outputs(str(tmp_path), "summarize_discovery")

    assert [entry.path for entry in outputs] == [output_path]
    assert outputs[0].source == "Wizard"
