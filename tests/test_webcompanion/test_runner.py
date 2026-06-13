"""Tests for webcompanion.runner — subprocess driver."""
import textwrap
import time

from webcompanion.runner import ScriptRunner


def _write_script(tmp_path, body):
    p = tmp_path / "fake_agent.py"
    p.write_text(textwrap.dedent(body), encoding="utf-8")
    return str(p)


def _run_collect(argv, timeout=15.0):
    events, exits = [], []
    r = ScriptRunner(argv, on_event=events.append, on_exit=exits.append)
    r.start()
    deadline = time.time() + timeout
    while not exits and time.time() < deadline:
        time.sleep(0.05)
    assert exits, "script did not exit in time"
    return events, exits[0]


def test_events_and_clean_exit(tmp_path):
    script = _write_script(tmp_path, """
        print("PROGRESS:10:starting")
        print("hello status")
        print("OUTPUT:E:/out/result.docx")
    """)
    events, rc = _run_collect([script])
    assert rc == 0
    kinds = [e.kind for e in events]
    assert "progress" in kinds and "status" in kinds and "output" in kinds
    out = next(e for e in events if e.kind == "output")
    assert out.path == "E:/out/result.docx"


def test_nonzero_exit(tmp_path):
    script = _write_script(tmp_path, """
        import sys
        print("about to fail")
        sys.exit(3)
    """)
    events, rc = _run_collect([script])
    assert rc == 3
    assert any(e.message == "about to fail" for e in events)


def test_cancel_terminates(tmp_path):
    script = _write_script(tmp_path, """
        import time
        print("PROGRESS:1")
        time.sleep(60)
    """)
    exits = []
    r = ScriptRunner([script], on_event=lambda e: None, on_exit=exits.append)
    r.start()
    time.sleep(1.0)
    r.cancel()
    deadline = time.time() + 10
    while not exits and time.time() < deadline:
        time.sleep(0.05)
    assert exits and exits[0] != 0
