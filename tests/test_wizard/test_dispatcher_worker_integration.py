"""End-to-end test: DispatcherSubprocessWorker runs a stub script that
mimics summarize_discovery.py's dispatcher mode — takes multiple files
positionally, accepts ``--report-dir``, writes ``worker_*.txt`` reports
pointing at party-grouped output .docx files, and emits PROGRESS lines.
"""
import os
import textwrap

import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.runners.dispatcher_subprocess_worker import (
    DispatcherSubprocessWorker,
)


# Mimics summarize_discovery.py's dispatcher path: receives N files, a
# --report-dir, groups inputs whose filename starts with "p1_" into one
# output and "p2_" into another, writes a worker_*.txt per input pointing
# at the chosen output, and overwrites each output once at the end (to
# simulate the consolidation pass).
STUB_DISPATCHER = textwrap.dedent('''
    import os, sys, random

    args = sys.argv[1:]
    report_dir = None
    positional = []
    i = 0
    while i < len(args):
        if args[i] == "--report-dir" and i + 1 < len(args):
            report_dir = args[i + 1]
            i += 2
        else:
            positional.append(args[i])
            i += 1

    assert report_dir is not None, "stub requires --report-dir"
    assert len(positional) > 1, "stub dispatcher requires multi-file input"

    # Outputs go alongside the inputs for simplicity (real script uses
    # NOTES/AI Output/; we just need a directory the test can inspect).
    out_dir = os.path.dirname(positional[0])
    os.makedirs(out_dir, exist_ok=True)

    party_outputs = set()
    print("PROGRESS:5:Dispatcher mode: " + str(len(positional)) + " files", flush=True)

    for n, p in enumerate(positional, 1):
        stem = os.path.basename(p).lower()
        if stem.startswith("p1_"):
            party = "PartyOne"
        elif stem.startswith("p2_"):
            party = "PartyTwo"
        else:
            party = "Other"
        out_path = os.path.join(out_dir, "Discovery_Responses_" + party + ".docx")
        # Create or append (we just write — test cares about path identity, not content).
        with open(out_path, "ab") as f:
            f.write(b"entry-for-" + os.path.basename(p).encode())
        # Write a worker_*.txt report identical to what process_document does.
        rname = "worker_" + str(os.getpid()) + "_" + str(random.randint(0, 99999)) + "_" + str(n) + ".txt"
        with open(os.path.join(report_dir, rname), "w") as rf:
            rf.write(out_path)
        party_outputs.add(out_path)
        pct = 5 + int(80 * n / len(positional))
        print("PROGRESS:" + str(pct) + ":Processed " + os.path.basename(p), flush=True)

    # Simulate consolidation pass (overwrites each unique output).
    for op in sorted(party_outputs):
        with open(op, "wb") as f:
            f.write(b"consolidated " + os.path.basename(op).encode())
    print("PROGRESS:100:Done", flush=True)
    sys.exit(0)
''')


@pytest.fixture
def stub_dispatcher_script(tmp_path, monkeypatch):
    scripts_dir = tmp_path / "Scripts"
    scripts_dir.mkdir()
    script_path = scripts_dir / "stub_dispatcher.py"
    script_path.write_text(STUB_DISPATCHER, encoding="utf-8")

    def _fake_script_path(self):
        return str(script_path)

    monkeypatch.setattr(
        DispatcherSubprocessWorker, "_script_path", _fake_script_path
    )
    return script_path


def test_dispatcher_worker_groups_files_by_party(tmp_path, stub_dispatcher_script, qtbot):
    """Four files (2 parties × 2 discovery types) → 2 unique outputs in the
    OutputPage, mirroring advanced-mode behavior. file_completed fires
    once per unique party file, not once per input."""
    case_path = tmp_path / "case"
    case_path.mkdir()
    inputs = []
    for name in (
        "p1_form_rogs.pdf",
        "p1_special_rogs.pdf",
        "p2_form_rogs.pdf",
        "p2_special_rogs.pdf",
    ):
        p = case_path / name
        p.write_text("x")
        inputs.append(str(p))

    w = DispatcherSubprocessWorker(
        script_name="stub_dispatcher.py",
        case_path=str(case_path),
        file_number="0000.000",
        files=inputs,
        settings={},
    )

    finished = []
    failed = []
    progress = []
    file_completed = []
    w.finished.connect(lambda p: finished.append(p))
    w.failed.connect(lambda m: failed.append(m))
    w.progress.connect(lambda v: progress.append(v))
    w.file_completed.connect(lambda p: file_completed.append(p))

    with qtbot.waitSignal(w.finished, timeout=15000):
        w.start()

    assert failed == [], f"unexpected failure: {failed}"
    assert finished, "worker did not emit finished"

    # 4 inputs but only 2 unique party outputs → file_completed twice.
    assert len(file_completed) == 2, file_completed
    completed_basenames = sorted(os.path.basename(p) for p in file_completed)
    assert completed_basenames == [
        "Discovery_Responses_PartyOne.docx",
        "Discovery_Responses_PartyTwo.docx",
    ]

    # Both outputs exist on disk and are post-consolidation.
    produced = sorted(p.name for p in case_path.glob("Discovery_Responses_*.docx"))
    assert produced == [
        "Discovery_Responses_PartyOne.docx",
        "Discovery_Responses_PartyTwo.docx",
    ]

    # output_paths exposes both, ordered by mtime (oldest → newest).
    assert sorted(os.path.basename(p) for p in w.output_paths) == completed_basenames

    # finished() carries one of the two (the newest).
    assert os.path.basename(finished[0]) in completed_basenames

    # Progress reached 100.
    assert progress[-1] == 100


def test_dispatcher_worker_failure_when_script_exits_nonzero(tmp_path, monkeypatch, qtbot):
    """Non-zero exit code is surfaced as failed(), not finished()."""
    scripts_dir = tmp_path / "Scripts"
    scripts_dir.mkdir()
    script_path = scripts_dir / "always_fails.py"
    script_path.write_text(
        "import sys\nprint('boom', flush=True)\nsys.exit(2)\n", encoding="utf-8"
    )

    def _fake_script_path(self):
        return str(script_path)

    monkeypatch.setattr(
        DispatcherSubprocessWorker, "_script_path", _fake_script_path
    )

    case_path = tmp_path / "case"
    case_path.mkdir()
    inputs = [str(case_path / "a.pdf"), str(case_path / "b.pdf")]
    for p in inputs:
        with open(p, "w") as f:
            f.write("x")

    w = DispatcherSubprocessWorker(
        script_name="always_fails.py",
        case_path=str(case_path),
        file_number="0000.000",
        files=inputs,
        settings={},
    )
    finished = []
    failed = []
    w.finished.connect(lambda p: finished.append(p))
    w.failed.connect(lambda m: failed.append(m))

    with qtbot.waitSignal(w.failed, timeout=10000):
        w.start()

    assert finished == []
    assert len(failed) == 1
    assert "code 2" in failed[0]
