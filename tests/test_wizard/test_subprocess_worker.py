"""Unit tests for SubprocessWorker output detection.

The worker no longer relies on set-difference of file paths; instead it
captures {path: mtime} before each run and treats any file whose mtime
advanced (or is new) as the produced output. This covers the case where
agents like summarize.py append in-place to a single AI_OUTPUT.docx —
which previously caused a false-failure on the 2nd file in a batch.
"""
import os
import time

import pytest

pytest.importorskip("pytestqt")

from PySide6.QtCore import QProcess

from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker


def _make_worker(case_path, files):
    """Build a worker without launching anything."""
    return SubprocessWorker(
        script_name="summarize.py",
        case_path=str(case_path),
        file_number="0000.000",
        files=list(files),
        settings={},
    )


def test_scan_outputs_returns_path_mtime_dict(tmp_path):
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    a = out_dir / "A.docx"
    b = out_dir / "B.docx"
    a.write_text("x")
    b.write_text("y")
    other = out_dir / "ignore.txt"
    other.write_text("z")

    w = _make_worker(tmp_path, [str(tmp_path / "input.pdf")])
    scan = w._scan_outputs()
    assert set(scan.keys()) == {str(a), str(b)}
    assert all(isinstance(v, float) for v in scan.values())


def test_detects_appended_output_on_second_run(tmp_path, qtbot):
    """Regression: 2nd file in batch must not be flagged as failure even
    when the agent appends to a pre-existing AI_OUTPUT.docx."""
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    shared = out_dir / "AI_OUTPUT.docx"
    shared.write_text("run-1 content")

    w = _make_worker(tmp_path, [str(tmp_path / "doc1.pdf"), str(tmp_path / "doc2.pdf")])

    # Simulate: snapshot BEFORE run-2 captures the file's existing mtime.
    w._pre_existing_outputs = w._scan_outputs()
    original_mtime = w._pre_existing_outputs[str(shared)]

    # Agent appends — bump mtime forward.
    time.sleep(0.05)
    shared.write_text("run-1 content\nrun-2 appended content")
    new_mtime = os.path.getmtime(shared)
    assert new_mtime > original_mtime

    # Now run the detection branch directly.
    failed_signals = []
    w.failed.connect(lambda msg: failed_signals.append(msg))
    w._on_process_finished(exit_code=0, exit_status=QProcess.ExitStatus.NormalExit)

    # Must NOT have emitted failed; must have recorded the appended file.
    assert failed_signals == []
    assert w._newest_output == str(shared)


def test_emits_failed_when_no_file_changed(tmp_path, qtbot):
    """If nothing in AI Output/ changed at all, surface the failure."""
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    shared = out_dir / "AI_OUTPUT.docx"
    shared.write_text("untouched")

    w = _make_worker(tmp_path, [str(tmp_path / "doc1.pdf")])
    w._pre_existing_outputs = w._scan_outputs()

    failed_signals = []
    w.failed.connect(lambda msg: failed_signals.append(msg))
    w._on_process_finished(exit_code=0, exit_status=QProcess.ExitStatus.NormalExit)

    assert len(failed_signals) == 1
    assert "no .docx" in failed_signals[0].lower()


def test_detects_new_output_file(tmp_path, qtbot):
    """A freshly created .docx (new path) is detected as the output."""
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)

    w = _make_worker(tmp_path, [str(tmp_path / "doc1.pdf")])
    w._pre_existing_outputs = w._scan_outputs()  # empty snapshot

    # Agent creates a new file.
    new_file = out_dir / "Discovery_Responses_Smith.docx"
    new_file.write_text("freshly written")

    failed_signals = []
    w.failed.connect(lambda msg: failed_signals.append(msg))
    w._on_process_finished(exit_code=0, exit_status=QProcess.ExitStatus.NormalExit)

    assert failed_signals == []
    assert w._newest_output == str(new_file)


def test_explicit_output_line_honored_for_nested_path(tmp_path, qtbot):
    """Agents that write into a session subfolder under NOTES/AI Output declare
    their file via an OUTPUT:<path> line. The worker must honor it even though
    the non-recursive scan would not find a nested file (regression: Depo Prep
    reported 'no .docx created' despite producing one)."""
    out_dir = tmp_path / "NOTES" / "AI Output" / "Depo Prep - Jane - 2026-05-27 2119"
    out_dir.mkdir(parents=True)
    nested = out_dir / "outline.docx"
    nested.write_text("generated")

    w = _make_worker(tmp_path, [str(tmp_path / "config.json")])
    w._pre_existing_outputs = w._scan_outputs()  # top-level scan misses the nested file

    w._handle_line(f"OUTPUT:{nested}")

    failed_signals = []
    finished_signals = []
    w.failed.connect(lambda m: failed_signals.append(m))
    w.finished.connect(lambda p: finished_signals.append(p))
    w._on_process_finished(exit_code=0, exit_status=QProcess.ExitStatus.NormalExit)

    assert failed_signals == []
    assert finished_signals == [str(nested)]
    assert w._newest_output == str(nested)


def test_explicit_output_missing_file_fails_clearly(tmp_path, qtbot):
    """If the agent declares an OUTPUT path that doesn't exist, surface a clear
    failure (not a silent success)."""
    w = _make_worker(tmp_path, [str(tmp_path / "config.json")])
    w._handle_line(r"OUTPUT:Z:\does\not\exist\outline.docx")

    failed_signals = []
    w.failed.connect(lambda m: failed_signals.append(m))
    w._on_process_finished(exit_code=0, exit_status=QProcess.ExitStatus.NormalExit)

    assert len(failed_signals) == 1
    assert "not found" in failed_signals[0].lower()


def test_output_line_suppressed_from_status_log(tmp_path, qtbot):
    """The OUTPUT: line is captured, not echoed to the status log."""
    w = _make_worker(tmp_path, [str(tmp_path / "config.json")])
    status_lines = []
    w.status.connect(lambda s: status_lines.append(s))
    w._handle_line("OUTPUT:C:/x/outline.docx")
    assert w._explicit_output == "C:/x/outline.docx"
    assert status_lines == []


def test_phase1_args_are_prepended_to_invocation(qtbot, tmp_path):
    """SubprocessWorker with phase1_args=['--phase=prep'] invokes the script
    with that flag before the file path."""
    from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker

    file_path = str(tmp_path / "in.docx")
    open(file_path, "w").close()

    captured = {}
    def fake_launch(self, extra_argv):
        captured["argv"] = list(extra_argv)
    SubprocessWorker._launch_process = fake_launch  # type: ignore

    w = SubprocessWorker(
        script_name="med_chron.py",
        case_path=str(tmp_path),
        file_number="",
        files=[file_path],
        settings={},
        phase1_args=["--phase=prep"],
        phase2_flag="--phase=run",
    )
    w.start()

    # Expect: [script_path, "--phase=prep", file_path]
    assert captured["argv"][1] == "--phase=prep"
    assert captured["argv"][2] == file_path


def test_phase2_flag_is_used_in_resume_with_config(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker

    captured = {}
    def fake_launch(self, extra_argv):
        captured["argv"] = list(extra_argv)
    SubprocessWorker._launch_process = fake_launch  # type: ignore

    w = SubprocessWorker(
        script_name="med_chron.py",
        case_path=str(tmp_path),
        file_number="",
        files=[str(tmp_path / "in.docx")],
        settings={},
        phase1_args=["--phase=prep"],
        phase2_flag="--phase=run",
    )
    w.resume_with_config(str(tmp_path / "session.json"))

    assert "--phase=run" in captured["argv"]
    assert str(tmp_path / "session.json") in captured["argv"]


def test_default_phase_flags_preserve_deposition_behavior(qtbot, tmp_path):
    """Default constructor args must still produce --phase=summary so the
    existing deposition flow keeps working."""
    from icharlotte_core.ui.wizard.runners.subprocess_worker import SubprocessWorker

    captured = {}
    def fake_launch(self, extra_argv):
        captured["argv"] = list(extra_argv)
    SubprocessWorker._launch_process = fake_launch  # type: ignore

    w = SubprocessWorker(
        script_name="summarize_deposition.py",
        case_path=str(tmp_path),
        file_number="",
        files=[str(tmp_path / "depo.pdf")],
        settings={},
    )
    w.resume_with_config(str(tmp_path / "session.json"))
    assert "--phase=summary" in captured["argv"]
