# tests/test_firm_briefs/test_authority_semantic.py
import numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite

def _v(i):
    v = np.zeros(384, np.float32); v[i] = 1.0; return v

def test_semantic_rerank_orders_paraphrase_first(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    # two cites that BOTH keyword-match "discovery", but only one is semantically the query
    idx.upsert_brief(path="p.pdf", content_hash="h", motion_type="compel", side="opposition",
                     heading="", profile="p", profile_vec=_v(0), char_len=1, ocr_ratio=0.0,
                     cites=[
                         HarvestedCite(case_name="A", reporter_citation="1 Cal.5th 1", norm_cite="1",
                                       proposition="forensic discovery imaging is permitted on a showing"),
                         HarvestedCite(case_name="B", reporter_citation="2 Cal.5th 2", norm_cite="2",
                                       proposition="discovery is generally broad"),
                     ],
                     prop_vecs=[_v(5), _v(9)])  # A->dim5, B->dim9
    # query vector aligned with A (dim5)
    hits = idx.authority_candidates("discovery", motion_type="compel", limit=5, query_vec=_v(5))
    assert hits[0]["case_name"] == "A"   # semantic rerank floats A above B

def test_fts_only_fallback_when_no_query_vec(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    idx.upsert_brief(path="p.pdf", content_hash="h", motion_type="compel", side="opposition",
                     heading="", profile="p", profile_vec=_v(0), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(case_name="A", reporter_citation="1 Cal.5th 1", norm_cite="1",
                                          proposition="meet and confer required")])
    hits = idx.authority_candidates("meet and confer", motion_type="compel", limit=5)
    assert any(h["case_name"] == "A" for h in hits)   # works without query_vec
