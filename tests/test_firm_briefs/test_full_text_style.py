# tests/test_firm_briefs/test_full_text_style.py
import numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite
from icharlotte_core.firm_briefs import style

def _vec(): v = np.ones(384, np.float32); return v/np.linalg.norm(v)

def test_index_stores_and_returns_full_text(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    idx.upsert_brief(path="p.pdf", content_hash="h", motion_type="compel", side="opposition",
                     heading="", profile="p", profile_vec=_vec(), char_len=10, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1", proposition="x")],
                     full_text="ARGUMENT\nThe motion fails because the meet and confer was inadequate.")
    assert "meet and confer" in idx.get_full_text("p.pdf")

def test_style_uses_stored_full_text_no_extraction(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path/"fb.db"), vectors_path=str(tmp_path/"v.f16"))
    idx.create_schema()
    idx.upsert_brief(path="img.pdf", content_hash="h", motion_type="compel", side="opposition",
                     heading="", profile="meet confer", profile_vec=_vec(), char_len=10, ocr_ratio=0.0,
                     cites=[HarvestedCite(reporter_citation="1 Cal.5th 1", norm_cite="1", proposition="x")],
                     full_text="ARGUMENT\n" + "stored opposition body " * 20)
    class Emb:
        dim=384
        def encode(self, t): return np.ones((len(t),384), np.float32)
    from types import SimpleNamespace
    m = SimpleNamespace(relief_requested="compel", principal_arguments=["meet and confer"])
    # extract_fn raises -> proves it used stored full_text, not extraction
    def boom(p): raise AssertionError("should not extract")
    out = style.select_exemplars("compel", "opposition", m, index=idx, embedder=Emb(),
                                 extract_fn=boom, cache_dir=str(tmp_path), max_chars=200)
    assert len(out) == 1 and out[0].startswith("ARGUMENT")
