"""Regression test for the SKIP_KEYWORDS filter in summarize_discovery.py.

Repro: user selected 4 discovery PDFs in the wizard. Two contained "RFA" in
the filename. The single-file branch of ``main()`` did
``if any(kw in input_filename_lower for kw in SKIP_KEYWORDS): sys.exit(0)``
*before* checking whether the caller supplied --output_path. The wizard's
parallel worker treats exit-0-with-no-output as a failure and emits a hidden
error, so the user just saw "only 2 of 4 outputs" with no explanation.

The wizard parallel worker passes ``--output_path`` per file. That flag is a
clear signal: "the caller selected this file deliberately; don't second-guess
the filename." When it's present, the filter must be bypassed.

These tests exercise ``main()`` end-to-end by patching the heavy lifting
(``process_document``, ``DocumentRegistry``, ``AgentLogger``) and watching
which branch ``main()`` takes.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path
from unittest.mock import patch

import pytest

ROOT = Path(__file__).resolve().parents[2]
SCRIPTS = ROOT / "Scripts"
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

import Scripts.summarize_discovery as sd  # noqa: E402


@pytest.fixture
def fake_pdf(tmp_path):
    p = tmp_path / "George Cline - Responses to RFAs, Set One.pdf"
    p.write_bytes(b"%PDF-1.4\n%fake")
    return str(p)


def _run_main_with_argv(argv: list[str]):
    """Invoke summarize_discovery.main() with the given argv, capturing sys.exit
    and the call to process_document. Returns ``(exit_code, process_document_call)``
    where ``process_document_call`` is the (args, kwargs) tuple, or ``None`` if
    process_document was never called.
    """
    seen = {"called": None}

    def _fake_process_document(*args, **kwargs):
        seen["called"] = (args, kwargs)
        return True

    with patch.object(sd, "process_document", _fake_process_document), \
         patch.object(sys, "argv", ["summarize_discovery.py", *argv]):
        try:
            sd.main()
            exit_code = 0
        except SystemExit as e:
            exit_code = e.code if e.code is not None else 0

    return exit_code, seen["called"]


def test_skip_keyword_in_filename_skips_without_output_path(fake_pdf):
    """Baseline: in single-file mode without --output_path, the filter still
    fires (the agent is still a directory-walking CLI in legacy contexts)."""
    exit_code, called = _run_main_with_argv([fake_pdf])
    assert exit_code == 0
    assert called is None, "process_document should not have run when filter matched"


def test_output_path_override_bypasses_skip_filter(fake_pdf):
    """The fix: when --output_path is supplied (wizard parallel-worker contract),
    the SKIP_KEYWORDS check must NOT short-circuit. process_document must run.
    """
    output_path = str(Path(fake_pdf).parent / "AI_OUTPUT - George Cline RFAs.docx")
    exit_code, called = _run_main_with_argv(
        [fake_pdf, "--output_path", output_path]
    )
    assert called is not None, (
        "process_document should have been invoked despite the filename containing 'RFA' — "
        "--output_path means the caller explicitly chose this file"
    )
    args, kwargs = called
    assert kwargs.get("output_path_override") == output_path
    assert exit_code == 0


def test_output_path_override_bypass_works_for_all_skip_keywords(tmp_path):
    """Every keyword in SKIP_KEYWORDS must be bypassed when --output_path is set."""
    for keyword in sd.SKIP_KEYWORDS:
        pdf = tmp_path / f"some {keyword} document.pdf"
        pdf.write_bytes(b"%PDF-1.4\n%fake")
        out = tmp_path / f"AI_OUTPUT - {keyword}.docx"
        exit_code, called = _run_main_with_argv(
            [str(pdf), "--output_path", str(out)]
        )
        assert called is not None, (
            f"keyword {keyword!r} should not block processing when --output_path is set"
        )
