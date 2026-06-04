import numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite
from clean_firm_index_names import clean_names

def _vec():
    v = np.ones(384, dtype=np.float32); return v/np.linalg.norm(v)

def test_backfill_cleans_existing(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    idx.upsert_brief(path="p.pdf", content_hash="h", motion_type="compel", side="moving",
                     heading="", profile="p", profile_vec=_vec(), char_len=1, ocr_ratio=0.0,
                     cites=[HarvestedCite(case_name="See Townsend v. Superior Court",
                                          reporter_citation="61 Cal.App.4th 1431",
                                          norm_cite="61cal.app.4th1431", proposition="x")])
    changed = clean_names(idx)
    assert changed == 1
    con = idx._conn()
    name = con.execute("SELECT case_name FROM citations").fetchone()[0]
    assert name == "Townsend v. Superior Court"
    assert clean_names(idx) == 0   # idempotent
