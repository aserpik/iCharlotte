# tests/test_firm_briefs/test_index.py
import os
import concurrent.futures
import numpy as np
import pytest

from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite


def _vec(seed: int) -> np.ndarray:
    rng = np.random.default_rng(seed)
    v = rng.standard_normal(384).astype(np.float32)
    return v / np.linalg.norm(v)


@pytest.fixture
def idx(tmp_path):
    db = str(tmp_path / "fb.db")
    vec = str(tmp_path / "profiles.f16")
    index = FirmBriefIndex(db_path=db, vectors_path=vec)
    index.create_schema()
    return index


def test_upsert_and_has_current(idx):
    cites = [HarvestedCite(case_name="Townsend v. Superior Court",
                           reporter_citation="61 Cal.App.4th 1431", year="1998",
                           norm_cite="61cal.app.4th1431",
                           proposition="meet and confer is required",
                           quoted_passage="reasonable and good faith effort")]
    bid = idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel",
                           side="opposition", heading="MEET AND CONFER",
                           profile="compel meet and confer", profile_vec=_vec(1),
                           char_len=5000, ocr_ratio=0.1, cites=cites)
    assert bid > 0
    assert idx.has_current("p1.pdf", "h1") is True
    assert idx.has_current("p1.pdf", "DIFFERENT") is False


def test_authority_candidates_keyword(idx):
    idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel",
                     side="opposition", heading="", profile="x", profile_vec=_vec(1),
                     char_len=10, ocr_ratio=0.0,
                     cites=[HarvestedCite(case_name="Townsend v. Superior Court",
                                          reporter_citation="61 Cal.App.4th 1431",
                                          year="1998", norm_cite="61cal.app.4th1431",
                                          proposition="a party must meet and confer in good faith",
                                          quoted_passage="good faith")])
    hits = idx.authority_candidates("meet and confer good faith", motion_type="compel", limit=5)
    assert any("Townsend" in h["case_name"] for h in hits)
    # type filter excludes other types
    assert idx.authority_candidates("meet and confer", motion_type="msj", limit=5) == []


def test_upsert_replaces_on_rehash(idx):
    idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel", side="moving",
                     heading="", profile="x", profile_vec=_vec(1), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1cal.5th1",
                                          proposition="p1")])
    idx.upsert_brief(path="p1.pdf", content_hash="h2", motion_type="compel", side="moving",
                     heading="", profile="x", profile_vec=_vec(2), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="2 Cal.5th 2", norm_cite="2cal.5th2",
                                          proposition="p2")])
    assert idx.has_current("p1.pdf", "h2")
    hits = idx.authority_candidates("p2", motion_type="compel", limit=5)
    assert any(h["norm_cite"] == "2cal.5th2" for h in hits)
    assert idx.authority_candidates("p1", motion_type="compel", limit=5) == []  # old cite gone


def test_mark_stale_excludes(idx):
    idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel", side="moving",
                     heading="", profile="x", profile_vec=_vec(1), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1cal.5th1",
                                          proposition="stale point")])
    idx.mark_stale("p1.pdf")
    assert idx.authority_candidates("stale point", motion_type="compel", limit=5) == []


def test_thread_local_connections(idx):
    idx.upsert_brief(path="p1.pdf", content_hash="h1", motion_type="compel", side="moving",
                     heading="", profile="x", profile_vec=_vec(1), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1cal.5th1",
                                          proposition="threaded meet and confer")])

    def _q(_):
        return idx.authority_candidates("threaded meet and confer", motion_type="compel", limit=5)

    with concurrent.futures.ThreadPoolExecutor(max_workers=4) as pool:
        results = list(pool.map(_q, range(8)))
    assert all(len(r) >= 1 for r in results)  # no swallowed cross-thread errors
