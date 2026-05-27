"""Tests for the statute verifier — fetch + cache + verdict mapping."""

from __future__ import annotations

import json
import os
import tempfile
from unittest.mock import MagicMock

import pytest

from icharlotte_core.legal_research.models import StatuteResult
from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.statute_verifier import StatuteVerifier


@pytest.fixture
def tmp_cache_dir():
    with tempfile.TemporaryDirectory() as tmp:
        yield tmp


def make_citation(law_code="CCP", section_num="2024.020", proposition="The deadline is 30 days."):
    return Citation(
        kind="statute",
        raw_text=f"Code Civ. Proc., § {section_num}",
        normalized=f"{law_code} {section_num}",
        proposition=proposition,
        body_offset=0,
        law_code=law_code,
        section_num=section_num,
    )


def test_fetch_hits_leginfo_on_cache_miss(tmp_cache_dir):
    leginfo = MagicMock()
    leginfo.get_section.return_value = StatuteResult(
        code="CCP",
        section="2024.020",
        title="Code of Civil Procedure",
        text="Discovery shall be completed on or before...",
        url="https://leginfo.legislature.ca.gov/...",
    )
    llm = MagicMock(return_value='{"verdict": "SUPPORTED", "evidence": "Discovery shall be completed...", "note": "Direct."}')

    v = StatuteVerifier(leginfo_client=leginfo, llm_callback=llm, cache_dir=tmp_cache_dir)
    cv = v.verify(make_citation())

    leginfo.get_section.assert_called_once_with("CCP", "2024.020")
    assert cv.verdict == "SUPPORTED"
    # Cache file written
    cache_file = os.path.join(tmp_cache_dir, "CCP_2024.020.json")
    assert os.path.exists(cache_file)


def test_cache_hit_skips_leginfo(tmp_cache_dir):
    cache_file = os.path.join(tmp_cache_dir, "CCP_2024.020.json")
    with open(cache_file, "w", encoding="utf-8") as f:
        json.dump(
            {
                "code": "CCP",
                "section": "2024.020",
                "title": "Code of Civil Procedure",
                "text": "Cached statute text about 30-day deadlines.",
                "url": "https://leginfo.legislature.ca.gov/...",
            },
            f,
        )

    leginfo = MagicMock()
    llm = MagicMock(return_value='{"verdict": "SUPPORTED", "evidence": "Cached statute text", "note": "ok"}')

    v = StatuteVerifier(leginfo_client=leginfo, llm_callback=llm, cache_dir=tmp_cache_dir)
    v.verify(make_citation())

    leginfo.get_section.assert_not_called()
    llm.assert_called_once()


def test_not_found_short_circuits_llm(tmp_cache_dir):
    leginfo = MagicMock()
    leginfo.get_section.return_value = None  # leginfo 404
    llm = MagicMock()

    v = StatuteVerifier(leginfo_client=leginfo, llm_callback=llm, cache_dir=tmp_cache_dir)
    cv = v.verify(make_citation(section_num="9999.999"))

    assert cv.verdict == "NOT_FOUND"
    llm.assert_not_called()
    assert "leginfo" in cv.note.lower()


def test_llm_returns_partial_verdict_propagates(tmp_cache_dir):
    leginfo = MagicMock()
    leginfo.get_section.return_value = StatuteResult(
        code="CCP",
        section="2024.020",
        title="Code of Civil Procedure",
        text="Statute text.",
        url="https://leginfo.legislature.ca.gov/...",
    )
    llm = MagicMock(return_value='{"verdict": "PARTIAL", "evidence": "Statute text.", "note": "Partial support."}')

    v = StatuteVerifier(leginfo_client=leginfo, llm_callback=llm, cache_dir=tmp_cache_dir)
    cv = v.verify(make_citation())

    assert cv.verdict == "PARTIAL"
    assert cv.evidence == "Statute text."
    assert cv.note == "Partial support."


def test_invalid_llm_json_falls_back_to_unverified(tmp_cache_dir):
    leginfo = MagicMock()
    leginfo.get_section.return_value = StatuteResult(
        code="CCP",
        section="2024.020",
        title="x",
        text="t",
        url="u",
    )
    llm = MagicMock(return_value="not json")

    v = StatuteVerifier(leginfo_client=leginfo, llm_callback=llm, cache_dir=tmp_cache_dir)
    cv = v.verify(make_citation())

    assert cv.verdict == "UNVERIFIED"
    assert "could not parse" in cv.note.lower() or "invalid" in cv.note.lower()


from icharlotte_core.opposition.statute_verifier import _parse_verdict_response


@pytest.mark.parametrize(
    "raw, expected_verdict",
    [
        ('{"verdict": "SUPPORTED", "evidence": "x", "note": "y"}', "SUPPORTED"),
        ('```json\n{"verdict": "PARTIAL", "evidence": "x", "note": "y"}\n```', "PARTIAL"),
        ('  ```\n{"verdict": "NOT_SUPPORTED", "evidence": "", "note": ""}\n```  ', "NOT_SUPPORTED"),
        ("not json at all", ""),
        ("", ""),
        ('{"verdict": "supported", "evidence": "x", "note": "y"}', "SUPPORTED"),  # lowercase upper-cased
    ],
)
def test_parse_verdict_response_variants(raw, expected_verdict):
    verdict, _evidence, _note = _parse_verdict_response(raw)
    assert verdict == expected_verdict
