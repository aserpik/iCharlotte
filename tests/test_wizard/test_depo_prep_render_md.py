from pathlib import Path

import pytest

from Scripts.depo_prep_lib.render_md import render_outline_md


def test_render_md_includes_questions_and_headers(tmp_path):
    payload = {
        "deponent_name": "Jane Doe", "deponent_role": "Plaintiff",
        "topics": [
            {"topic_id": "t01", "title": "Pre-existing conditions",
             "strategic_note": "Establish baseline.",
             "questions": [
                 {"n": 1, "text": "Before 2024", "purpose": "Baseline",
                  "source_facts": ["RFA #7 denied"]},
             ]}
        ],
        "coverage_gaps": ["Gap 1"],
    }
    out = tmp_path / "outline.md"
    render_outline_md(outline=payload, output_path=out)
    md = out.read_text(encoding="utf-8")
    assert "# Depo Prep Outline — Jane Doe" in md
    assert "## Pre-existing conditions" in md
    assert "_Strategic: Establish baseline._" in md
    assert "1." in md and "Before 2024" in md
    assert "Purpose" in md
    assert "RFA #7 denied" in md
    assert "## Coverage notes from the AI" in md
    assert "Gap 1" in md
