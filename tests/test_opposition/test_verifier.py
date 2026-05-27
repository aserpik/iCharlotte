"""Tests for the citation verifier orchestrator."""

from __future__ import annotations

import tempfile
from unittest.mock import MagicMock

import pytest

from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.models import CitationVerification
from icharlotte_core.opposition.verifier import OppositionVerifier


@pytest.fixture
def tmp_cache_root():
    with tempfile.TemporaryDirectory() as tmp:
        yield tmp


def _supported(citation_text, kind="case", **extra):
    cv = CitationVerification(
        citation_text=citation_text,
        kind=kind,
        verdict="SUPPORTED",
        evidence="evidence",
        note="ok",
        **extra,
    )
    return cv


def test_dispatches_case_to_case_verifier(tmp_cache_root):
    case_v = MagicMock()
    case_v.verify.return_value = _supported("Smith v. Jones (2010) 50 Cal.4th 100")
    statute_v = MagicMock()

    v = OppositionVerifier(
        case_verifier=case_v,
        statute_verifier=statute_v,
        max_workers=1,
    )
    citations = [
        Citation(kind="case", raw_text="Smith v. Jones (2010) 50 Cal.4th 100", normalized="Smith 50 Cal.4th 100")
    ]
    results = v.verify_all(citations)

    case_v.verify.assert_called_once()
    statute_v.verify.assert_not_called()
    assert len(results) == 1
    assert results[0].verdict == "SUPPORTED"


def test_dispatches_statute_to_statute_verifier(tmp_cache_root):
    case_v = MagicMock()
    statute_v = MagicMock()
    statute_v.verify.return_value = _supported("CCP § 2024.020", kind="statute")

    v = OppositionVerifier(case_verifier=case_v, statute_verifier=statute_v, max_workers=1)
    cites = [Citation(kind="statute", raw_text="CCP § 2024.020", normalized="CCP 2024.020")]
    v.verify_all(cites)

    statute_v.verify.assert_called_once()
    case_v.verify.assert_not_called()


def test_rules_of_court_are_unverified(tmp_cache_root):
    case_v = MagicMock()
    statute_v = MagicMock()

    v = OppositionVerifier(case_verifier=case_v, statute_verifier=statute_v, max_workers=1)
    cites = [Citation(kind="rule", raw_text="California Rules of Court, rule 3.1345", normalized="CRC rule 3.1345")]
    results = v.verify_all(cites)

    assert results[0].verdict == "UNVERIFIED"
    case_v.verify.assert_not_called()
    statute_v.verify.assert_not_called()


def test_duplicate_cites_verified_once(tmp_cache_root):
    case_v = MagicMock()
    case_v.verify.return_value = _supported("Smith v. Jones (2010) 50 Cal.4th 100")
    statute_v = MagicMock()

    v = OppositionVerifier(case_verifier=case_v, statute_verifier=statute_v, max_workers=1)
    cite = Citation(kind="case", raw_text="Smith v. Jones (2010) 50 Cal.4th 100", normalized="Smith 50 Cal.4th 100")
    results = v.verify_all([cite, cite])

    assert case_v.verify.call_count == 1
    assert len(results) == 2
    assert all(r.verdict == "SUPPORTED" for r in results)


def test_progress_callback_invoked_per_citation(tmp_cache_root):
    case_v = MagicMock()
    case_v.verify.return_value = _supported("Smith v. Jones (2010) 50 Cal.4th 100")
    statute_v = MagicMock()
    statute_v.verify.return_value = _supported("CCP § 2024.020", kind="statute")

    progress_msgs: list[str] = []

    v = OppositionVerifier(case_verifier=case_v, statute_verifier=statute_v, max_workers=1)
    v.verify_all(
        [
            Citation(kind="case", raw_text="Smith v. Jones (2010) 50 Cal.4th 100", normalized="Smith"),
            Citation(kind="statute", raw_text="CCP § 2024.020", normalized="CCP 2024.020"),
        ],
        on_progress=progress_msgs.append,
    )

    assert any("Smith" in m for m in progress_msgs)
    assert any("CCP" in m for m in progress_msgs)


def test_unknown_kind_is_unverified(tmp_cache_root):
    case_v = MagicMock()
    statute_v = MagicMock()
    v = OppositionVerifier(case_verifier=case_v, statute_verifier=statute_v, max_workers=1)
    results = v.verify_all([Citation(kind="unknown", raw_text="???", normalized="???")])
    assert results[0].verdict == "UNVERIFIED"


def test_build_opposition_verifier_uses_project_cache_paths():
    from icharlotte_core.opposition.verifier import build_opposition_verifier

    # Pass dummy callbacks; we only check the verifier object has the right shape.
    def llm(_sys, _user):
        return ""

    v = build_opposition_verifier(courtlistener_token="X", llm_callback=llm)
    assert isinstance(v, OppositionVerifier)
    # Cache dirs should be under Scripts/prompts/oppose_motion/.cache
    assert "oppose_motion" in v.case.cache_dir.replace("\\", "/").lower()
    assert "oppose_motion" in v.statute.cache_dir.replace("\\", "/").lower()
