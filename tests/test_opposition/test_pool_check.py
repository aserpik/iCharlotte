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


def test_pincite_pool_variant_still_matches():
    # Pool citation carries a pincite; the brief cites the clean reporter cite.
    # They should still be treated as in-pool.
    pool = [RetrievedAuthority(cluster_id="1", case_name="A v. B",
                               citation="226 Cal.App.4th 401, 415")]
    cites = [Citation(kind="case", raw_text="*A v. B* (2014) 226 Cal.App.4th 401",
                      normalized="A v. B 226 Cal.App.4th 401", reporter_citation="226 Cal.App.4th 401")]
    to_verify, off_pool = pool_membership_check(cites, pool)
    assert len(to_verify) == 1
    assert off_pool == []


def test_digit_overlap_does_not_false_match():
    # A fabricated cite whose normalized form is a bare substring of a real pool
    # cite (e.g. "0cal.app.4th100" inside "100cal.app.4th1000") must NOT be
    # treated as in-pool — the old bidirectional-substring test wrongly matched it.
    pool = [RetrievedAuthority(cluster_id="1", case_name="Real v. Case",
                               citation="100 Cal.App.4th 1000")]
    cites = [Citation(kind="case", raw_text="*Ghost v. Phantom* (2019) 0 Cal.App.4th 100",
                      normalized="Ghost v. Phantom 0 Cal.App.4th 100", reporter_citation="0 Cal.App.4th 100")]
    to_verify, off_pool = pool_membership_check(cites, pool)
    assert to_verify == []
    assert len(off_pool) == 1
    assert off_pool[0].verdict == "NOT_FOUND"


def test_enrich_with_pool_signals_copies_count_and_year():
    from icharlotte_core.opposition.models import CitationVerification
    from icharlotte_core.opposition.verifier import enrich_with_pool_signals

    verifications = [CitationVerification(citation_text="*A v. B* (2014) 226 Cal.App.4th 401",
                                          kind="case", verdict="SUPPORTED",
                                          normalized_citation="A v. B 226 Cal.App.4th 401")]
    pool = [RetrievedAuthority(cluster_id="1", case_name="A v. B", citation="226 Cal.App.4th 401",
                               citation_count=37, latest_citing_year="2021")]
    enrich_with_pool_signals(verifications, pool)
    assert verifications[0].citation_count == 37
    assert verifications[0].latest_citing_year == "2021"


def test_enrich_ignores_digit_overlap_mismatch():
    from icharlotte_core.opposition.models import CitationVerification
    from icharlotte_core.opposition.verifier import enrich_with_pool_signals

    verifications = [CitationVerification(citation_text="*Ghost v. Phantom* (2019) 0 Cal.App.4th 100",
                                          kind="case", verdict="NOT_FOUND",
                                          normalized_citation="Ghost v. Phantom 0 Cal.App.4th 100")]
    pool = [RetrievedAuthority(cluster_id="1", case_name="Real v. Case", citation="100 Cal.App.4th 1000",
                               citation_count=50, latest_citing_year="2020")]
    enrich_with_pool_signals(verifications, pool)
    assert verifications[0].citation_count is None
    assert verifications[0].latest_citing_year == ""
