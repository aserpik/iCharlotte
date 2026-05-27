"""Unit tests for ParallelSubprocessWorker.

We don't spawn real QProcesses here — instead we exercise the pure-Python
plumbing (path construction, refusal of two-phase agents, finalization
logic with mock state) and one black-box behavior via a tiny stub.
"""
import os

import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.runners.parallel_subprocess_worker import (
    ParallelSubprocessWorker,
)


def _make_worker(case_path, files, script_name="summarize.py", max_concurrent=3):
    return ParallelSubprocessWorker(
        script_name=script_name,
        case_path=str(case_path),
        file_number="0000.000",
        files=list(files),
        settings={},
        max_concurrent=max_concurrent,
    )


def test_refuses_two_phase_script(tmp_path):
    with pytest.raises(ValueError, match="two-phase"):
        _make_worker(tmp_path, ["a.pdf", "b.pdf"], script_name="summarize_deposition.py")


def test_refuses_empty_file_list(tmp_path):
    with pytest.raises(ValueError, match="at least one"):
        _make_worker(tmp_path, [])


def test_refuses_bad_concurrency(tmp_path):
    with pytest.raises(ValueError, match="max_concurrent"):
        _make_worker(tmp_path, ["a.pdf"], max_concurrent=0)


def test_unique_output_path_uses_input_stem(tmp_path):
    w = _make_worker(tmp_path, ["doc1.pdf", "doc2.pdf"])
    p = w._unique_output_path_for("/some/dir/doc1.pdf")
    assert os.path.basename(p) == "AI_OUTPUT - doc1.docx"
    assert "NOTES" in p and "AI Output" in p


def test_unique_output_path_disambiguates_when_in_flight_targets_collide(tmp_path):
    """If a prior file in this batch already targets the same output, append (2)."""
    w = _make_worker(tmp_path, ["a.pdf", "b.pdf"])
    # Pretend an earlier scheduled file already claimed the AI_OUTPUT - same.docx path.
    w._output_paths.append(
        os.path.join(str(tmp_path), "NOTES", "AI Output", "AI_OUTPUT - same.docx")
    )
    p = w._unique_output_path_for("/elsewhere/same.pdf")
    assert os.path.basename(p) == "AI_OUTPUT - same (2).docx"


def test_unique_output_path_sanitizes_unsafe_chars(tmp_path):
    w = _make_worker(tmp_path, ["a.pdf"])
    p = w._unique_output_path_for('/dir/foo:bar*baz?.pdf')
    # ':' '*' '?' should be stripped.
    assert "AI_OUTPUT - foobarbaz" in os.path.basename(p)


def test_finalize_emits_failed_when_no_outputs_produced(tmp_path, qtbot):
    w = _make_worker(tmp_path, ["a.pdf", "b.pdf"])
    # Mark both as completed but no outputs collected.
    w._completed_files = {"a.pdf", "b.pdf"}
    failed = []
    w.failed.connect(lambda msg: failed.append(msg))
    w._finalize()
    assert len(failed) == 1
    assert "no output" in failed[0].lower()


def test_finalize_emits_failed_when_any_child_failed(tmp_path, qtbot):
    w = _make_worker(tmp_path, ["a.pdf", "b.pdf"])
    w._completed_files = {"a.pdf"}
    w._failed_files = {"b.pdf"}
    w._first_failure_msg = "[b.pdf] agent exited with code 1"
    # Pretend "a.pdf" produced a docx
    out = tmp_path / "NOTES" / "AI Output" / "AI_OUTPUT - a.docx"
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text("x")
    w._output_paths.append(str(out))

    failed = []
    finished = []
    w.failed.connect(lambda msg: failed.append(msg))
    w.finished.connect(lambda p: finished.append(p))

    w._finalize()
    assert finished == []
    assert len(failed) == 1
    assert "1 of 2" in failed[0]
    assert "b.pdf" in failed[0]


def test_finalize_emits_finished_with_newest_output(tmp_path, qtbot):
    import time

    out_dir = tmp_path / "NOTES" / "AI Output"
    out_dir.mkdir(parents=True)
    older = out_dir / "AI_OUTPUT - a.docx"
    newer = out_dir / "AI_OUTPUT - b.docx"
    older.write_text("x")
    time.sleep(0.05)
    newer.write_text("y")

    w = _make_worker(tmp_path, ["a.pdf", "b.pdf"])
    w._completed_files = {"a.pdf", "b.pdf"}
    w._output_paths = [str(older), str(newer)]

    finished = []
    w.finished.connect(lambda p: finished.append(p))
    w._finalize()
    assert finished == [str(newer)]


def test_aggregate_progress_averages_in_flight(tmp_path, qtbot):
    w = _make_worker(tmp_path, ["a.pdf", "b.pdf", "c.pdf", "d.pdf"])
    # b is fully done, c is at 50%, a at 20%, d untouched.
    w._completed_files = {"b.pdf"}
    w._raw_pct = {"a.pdf": 20, "b.pdf": 0, "c.pdf": 50, "d.pdf": 0}

    progress_values = []
    w.progress.connect(lambda v: progress_values.append(v))
    w._emit_aggregate_progress()
    # (20 + 100 + 50 + 0) / 4 = 42
    assert progress_values == [42]
