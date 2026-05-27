"""End-to-end test: ParallelSubprocessWorker runs a tiny stub script
that mimics a wizard agent (writes a .docx, prints PROGRESS lines,
honors --output_path). Verifies the full QProcess wiring.

The stub stands in for Scripts/summarize.py so we don't need a real
LLM call. To swap the script root, we monkey-patch the _script_path
method.
"""
import os
import textwrap

import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.ui.wizard.runners.parallel_subprocess_worker import (
    ParallelSubprocessWorker,
)


STUB_SCRIPT = textwrap.dedent('''
    # Stub agent: takes <input_path> [--output_path <path>] and writes a fake docx.
    import os, sys

    args = sys.argv[1:]
    output_path = None
    positional = []
    i = 0
    while i < len(args):
        if args[i] == "--output_path" and i + 1 < len(args):
            output_path = args[i + 1]
            i += 2
        else:
            positional.append(args[i])
            i += 1

    input_path = positional[0]
    print("PROGRESS:25:Starting work on " + os.path.basename(input_path), flush=True)
    print("PROGRESS:50:Halfway", flush=True)

    if not output_path:
        print("ERROR:missing --output_path", flush=True)
        sys.exit(1)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "wb") as f:
        f.write(b"stub-docx")

    print("PROGRESS:100:Done", flush=True)
    sys.exit(0)
''')


@pytest.fixture
def stub_script(tmp_path, monkeypatch):
    """Drop a stub agent on disk and wire the worker to invoke it."""
    scripts_dir = tmp_path / "Scripts"
    scripts_dir.mkdir()
    script_path = scripts_dir / "stub_agent.py"
    script_path.write_text(STUB_SCRIPT, encoding="utf-8")

    # Force the worker to find OUR stub.
    def _fake_script_path(self):
        return str(script_path)

    monkeypatch.setattr(
        ParallelSubprocessWorker, "_script_path", _fake_script_path
    )
    return script_path


def test_parallel_worker_runs_stub_on_multiple_files(tmp_path, stub_script, qtbot):
    case_path = tmp_path / "case"
    case_path.mkdir()
    # Create three fake input "documents".
    inputs = []
    for name in ("doc_a.pdf", "doc_b.pdf", "doc_c.pdf"):
        p = case_path / name
        p.write_text("x")
        inputs.append(str(p))

    w = ParallelSubprocessWorker(
        script_name="stub_agent.py",
        case_path=str(case_path),
        file_number="0000.000",
        files=inputs,
        settings={},
        max_concurrent=2,
    )

    finished = []
    failed = []
    progress = []
    w.finished.connect(lambda p: finished.append(p))
    w.failed.connect(lambda m: failed.append(m))
    w.progress.connect(lambda v: progress.append(v))

    file_completed = []
    w.file_completed.connect(lambda p: file_completed.append(p))

    with qtbot.waitSignal(w.finished, timeout=15000):
        w.start()

    assert failed == [], f"unexpected failure: {failed}"
    assert finished, "worker did not emit finished"
    # Each successful child should have emitted file_completed exactly once.
    assert len(file_completed) == 3
    assert sorted(os.path.basename(p) for p in file_completed) == [
        "AI_OUTPUT - doc_a.docx",
        "AI_OUTPUT - doc_b.docx",
        "AI_OUTPUT - doc_c.docx",
    ]
    out_dir = case_path / "NOTES" / "AI Output"
    produced = sorted(p.name for p in out_dir.glob("*.docx"))
    assert produced == [
        "AI_OUTPUT - doc_a.docx",
        "AI_OUTPUT - doc_b.docx",
        "AI_OUTPUT - doc_c.docx",
    ]
    # Progress climbs from <100 to 100.
    assert progress, "no progress emitted"
    assert progress[-1] == 100
    # Reported output path is one of the three.
    assert os.path.basename(finished[0]).startswith("AI_OUTPUT - doc_")
