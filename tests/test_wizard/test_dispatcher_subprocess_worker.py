"""Unit tests for DispatcherSubprocessWorker.

The dispatcher worker is used for scripts that have their own internal
multi-file dispatcher (summarize_discovery.py). It launches a single
subprocess passing all input files positionally and reads
``worker_*.txt`` files from ``--report-dir`` to discover the outputs the
script produced. The party-grouping + consolidation passes happen
inside the script itself; this worker just plumbs progress through and
collects outputs at the end.

We don't spawn real QProcesses here — we exercise the pure-Python
plumbing: constructor validation, argv assembly, report-dir parsing
(dedup, missing files, ordering), and finalize-success / -failure
emission paths.
"""
import os
import time

import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.runners.dispatcher_subprocess_worker import (
    DispatcherSubprocessWorker,
)


def _make_worker(case_path, files, script_name="summarize_discovery.py"):
    return DispatcherSubprocessWorker(
        script_name=script_name,
        case_path=str(case_path),
        file_number="0000.000",
        files=list(files),
        settings={},
    )


# ---------- Construction ----------


def test_refuses_empty_file_list(tmp_path):
    with pytest.raises(ValueError, match="at least one"):
        _make_worker(tmp_path, [])


def test_script_path_points_at_scripts_dir(tmp_path):
    w = _make_worker(tmp_path, ["a.pdf"])
    path = w._script_path()
    assert path.endswith(os.path.join("Scripts", "summarize_discovery.py"))


# ---------- Argv assembly ----------


def test_argv_passes_all_files_positionally_with_report_dir(tmp_path):
    """All input files should be passed as positional args, with a single
    ``--report-dir`` at the end. No ``--output_path``: that flag would
    disable the script's party-grouping by forcing one output per input.
    """
    files = ["/case/a.pdf", "/case/b.pdf", "/case/c.pdf"]
    w = _make_worker(tmp_path, files)
    argv = w._build_argv(report_dir=str(tmp_path / "reports"))

    # First arg is "-u" for unbuffered stdout; second is the script path.
    assert argv[0] == "-u"
    assert argv[1].endswith("summarize_discovery.py")
    # All files appear in order, BEFORE --report-dir.
    for i, f in enumerate(files):
        assert argv[2 + i] == f
    # --report-dir <path> follows the files.
    assert argv[2 + len(files)] == "--report-dir"
    assert argv[2 + len(files) + 1] == str(tmp_path / "reports")
    # Critically: no --output_path anywhere.
    assert "--output_path" not in argv


# ---------- Report-dir parsing ----------


def _write_worker_report(report_dir, path: str, suffix: str = "1") -> None:
    """Simulate what summarize_discovery.py's process_document writes."""
    os.makedirs(report_dir, exist_ok=True)
    with open(os.path.join(report_dir, f"worker_{suffix}.txt"), "w") as f:
        f.write(path)


def test_collect_output_paths_empty_when_no_report_dir(tmp_path):
    w = _make_worker(tmp_path, ["a.pdf"])
    w._report_dir = None
    assert w._collect_output_paths() == []


def test_collect_output_paths_empty_when_report_dir_missing(tmp_path):
    w = _make_worker(tmp_path, ["a.pdf"])
    w._report_dir = str(tmp_path / "does_not_exist")
    assert w._collect_output_paths() == []


def test_collect_output_paths_dedupes_same_party(tmp_path):
    """Two children writing for the same party share one output file.
    They each write a worker_*.txt pointing at that same path. We should
    return the path once.
    """
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    party_file = out_dir / "Discovery_Responses_PlaintiffSmith.docx"
    party_file.write_text("x")

    report_dir = tmp_path / "reports"
    _write_worker_report(str(report_dir), str(party_file), suffix="form")
    _write_worker_report(str(report_dir), str(party_file), suffix="special")

    w = _make_worker(tmp_path, ["a.pdf", "b.pdf"])
    w._report_dir = str(report_dir)
    paths = w._collect_output_paths()
    assert paths == [str(party_file)]


def test_collect_output_paths_skips_missing_files(tmp_path):
    """If a worker reported a path that no longer exists, drop it
    silently — don't surface a broken link to the OutputPage.
    """
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    real = out_dir / "Discovery_Responses_Real.docx"
    real.write_text("x")
    bogus = out_dir / "Discovery_Responses_Missing.docx"  # never created

    report_dir = tmp_path / "reports"
    _write_worker_report(str(report_dir), str(real), suffix="1")
    _write_worker_report(str(report_dir), str(bogus), suffix="2")

    w = _make_worker(tmp_path, ["a.pdf", "b.pdf"])
    w._report_dir = str(report_dir)
    paths = w._collect_output_paths()
    assert paths == [str(real)]


def test_collect_output_paths_returns_multiple_unique_outputs(tmp_path):
    """Two different parties → two unique outputs."""
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    p1 = out_dir / "Discovery_Responses_PlaintiffOne.docx"
    p2 = out_dir / "Discovery_Responses_PlaintiffTwo.docx"
    p1.write_text("x")
    time.sleep(0.05)
    p2.write_text("y")

    report_dir = tmp_path / "reports"
    _write_worker_report(str(report_dir), str(p1), suffix="1a")
    _write_worker_report(str(report_dir), str(p1), suffix="1b")
    _write_worker_report(str(report_dir), str(p2), suffix="2a")
    _write_worker_report(str(report_dir), str(p2), suffix="2b")

    w = _make_worker(tmp_path, ["a.pdf", "b.pdf", "c.pdf", "d.pdf"])
    w._report_dir = str(report_dir)
    paths = w._collect_output_paths()
    # Sorted by mtime (oldest first); p1 was written first.
    assert paths == [str(p1), str(p2)]


# ---------- Finalize behavior ----------


def test_finalize_success_emits_file_completed_per_unique_output(tmp_path, qtbot):
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    p1 = out_dir / "Discovery_Responses_One.docx"
    p2 = out_dir / "Discovery_Responses_Two.docx"
    p1.write_text("x")
    time.sleep(0.05)
    p2.write_text("y")

    report_dir = tmp_path / "reports"
    _write_worker_report(str(report_dir), str(p1), suffix="1")
    _write_worker_report(str(report_dir), str(p2), suffix="2")

    w = _make_worker(tmp_path, ["a.pdf", "b.pdf"])
    w._report_dir = str(report_dir)

    completed = []
    finished = []
    w.file_completed.connect(lambda p: completed.append(p))
    w.finished.connect(lambda p: finished.append(p))

    w._finalize_success()

    assert completed == [str(p1), str(p2)]
    # finished() reports the newest output for the OutputPage's initial view.
    assert finished == [str(p2)]
    # And output_paths exposes all of them.
    assert w.output_paths == [str(p1), str(p2)]


def test_finalize_success_with_no_outputs_emits_failed(tmp_path, qtbot):
    """Process exited 0 but the script produced no party files."""
    w = _make_worker(tmp_path, ["a.pdf"])
    w._report_dir = str(tmp_path / "empty_reports")
    os.makedirs(w._report_dir, exist_ok=True)

    failed = []
    finished = []
    w.failed.connect(lambda msg: failed.append(msg))
    w.finished.connect(lambda p: finished.append(p))

    w._finalize_success()
    assert finished == []
    assert len(failed) == 1
    assert "no output" in failed[0].lower()


def test_finalize_success_cleans_up_report_dir(tmp_path, qtbot):
    """After a successful run, the temp report-dir is wiped."""
    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    p1 = out_dir / "Discovery_Responses_One.docx"
    p1.write_text("x")

    report_dir = tmp_path / "reports"
    _write_worker_report(str(report_dir), str(p1), suffix="1")
    assert os.path.isdir(str(report_dir))

    w = _make_worker(tmp_path, ["a.pdf"])
    w._report_dir = str(report_dir)
    w._finalize_success()

    assert not os.path.isdir(str(report_dir)), "report-dir should be cleaned up"
    assert w._report_dir is None


def test_cancel_before_start_emits_cancelled(tmp_path, qtbot):
    w = _make_worker(tmp_path, ["a.pdf"])
    cancelled = []
    w.cancelled.connect(lambda: cancelled.append(True))
    w.cancel()
    assert cancelled == [True]
