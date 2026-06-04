import os, numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite

def _v(): v=np.ones(384,np.float32); return v/np.linalg.norm(v)

def test_compact_drops_stale_vector_rows(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    for p in ("a.pdf", "b.pdf"):
        idx.upsert_brief(path=p, content_hash="h", motion_type="compel", side="moving",
                         heading="", profile="p", profile_vec=_v(), char_len=1, ocr_ratio=0.0,
                         cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1", proposition="x")])
    idx.mark_stale("a.pdf")            # a becomes stale -> its vector row is dead
    rows_before = idx.load_vectors().shape[0]
    idx.compact()
    rows_after = idx.load_vectors().shape[0]
    assert rows_after < rows_before    # dead row reclaimed
    # surviving brief still queryable + vec aligned
    hits = idx.style_candidates(_v(), motion_type="compel", side="moving", k=5)
    assert [h["path"] for h in hits] == ["b.pdf"]
