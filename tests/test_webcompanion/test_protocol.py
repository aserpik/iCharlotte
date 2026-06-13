"""Tests for webcompanion.protocol — wizard stdout line parsing."""
from webcompanion.protocol import parse_line


def test_progress_plain():
    p = parse_line("PROGRESS: 42")
    assert p.kind == "progress" and p.pct == 42 and p.message == ""


def test_progress_with_message():
    p = parse_line("PROGRESS:7:Extracting text")
    assert p.kind == "progress" and p.pct == 7 and p.message == "Extracting text"


def test_progress_clamped():
    assert parse_line("PROGRESS:150").pct == 100


def test_awaiting_input():
    p = parse_line(r"AWAITING_INPUT:C:\logs\depo_sessions\abc.json")
    assert p.kind == "awaiting" and p.path == r"C:\logs\depo_sessions\abc.json"


def test_output():
    p = parse_line(r"OUTPUT:E:\case\NOTES\AI Output\summary.docx")
    assert p.kind == "output" and p.path.endswith("summary.docx")


def test_status_fallthrough():
    p = parse_line("Loading model...")
    assert p.kind == "status" and p.message == "Loading model..."
