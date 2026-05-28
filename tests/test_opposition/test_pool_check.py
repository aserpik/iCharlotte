"""Tests for the deterministic pool-membership check."""

from __future__ import annotations

from icharlotte_core.opposition.citation_parser import Citation
from icharlotte_core.opposition.models import RetrievedAuthority
from icharlotte_core.opposition.verifier import pool_membership_check


def _pool():
    return [RetrievedAuthority(cluster_id="1", case_name="A v. B", citation="226 Cal.App.4th 401")]


def test_in_pool_case_passes_through():
    cites = [Citation(kind="case", raw_text="*A v. B* (2014) 226 Cal.App.4th 401",
                      normalized="A v. B 226 Cal.App.4th 401", reporter_citation="226 Cal.App.4th 401")]
    to_verify, off_pool = pool_membership_check(cites, _pool())
    assert len(to_verify) == 1
    assert off_pool == []


def test_off_pool_case_is_flagged_not_found():
    cites = [Citation(kind="case", raw_text="*Ghost v. Phantom* (2019) 9 Cal.5th 9",
                      normalized="Ghost v. Phantom 9 Cal.5th 9", reporter_citation="9 Cal.5th 9")]
    to_verify, off_pool = pool_membership_check(cites, _pool())
    assert to_verify == []
    assert len(off_pool) == 1
    assert off_pool[0].verdict == "NOT_FOUND"
    assert "pool" in off_pool[0].note.lower()


def test_statutes_always_pass_through():
    cites = [Citation(kind="statute", raw_text="Code Civ. Proc., § 2024.020",
                      normalized="CCP 2024.020", law_code="CCP", section_num="2024.020")]
    to_verify, off_pool = pool_membership_check(cites, _pool())
    assert len(to_verify) == 1
    assert off_pool == []


def test_empty_pool_passes_cases_through():
    # When no authorities were retrieved at all, do not flag everything; let the
    # network verifier handle it (grounding simply did not run).
    cites = [Citation(kind="case", raw_text="*A v. B* (2014) 226 Cal.App.4th 401",
                      normalized="A v. B 226 Cal.App.4th 401", reporter_citation="226 Cal.App.4th 401")]
    to_verify, off_pool = pool_membership_check(cites, [])
    assert len(to_verify) == 1
    assert off_pool == []
