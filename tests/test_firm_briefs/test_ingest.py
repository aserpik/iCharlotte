# tests/test_firm_briefs/test_ingest.py
import os
import numpy as np
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.embedding import get_embedder
from icharlotte_core.firm_briefs.ingest import ingest_root

ROOT_NAME = "5800_AMTRUST_Pleadings_PDFs"
SAMPLE = ("I. PLAINTIFF FAILED TO MEET AND CONFER\n"
          "A party must meet and confer in good faith. "
          "Townsend v. Superior Court (1998) 61 Cal.App.4th 1431, 1438.\n")


def _make_lib(tmp_path):
    root = tmp_path / ROOT_NAME / "Oppositions" / "Motion to Compel"
    root.mkdir(parents=True)
    f = root / "008 - Rosas__opp.pdf"
    f.write_text("placeholder")  # content irrelevant; extract_fn is injected
    return str(tmp_path / ROOT_NAME), str(f)


def test_ingest_indexes_one_brief(tmp_path):
    root, fpath = _make_lib(tmp_path)
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"),
                         vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    emb = get_embedder(fake=True)
    res = ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    assert res["added"] == 1
    hits = idx.authority_candidates("meet and confer good faith", motion_type="compel", limit=5)
    assert any("Townsend" in h["case_name"] for h in hits)


def test_ingest_is_incremental(tmp_path):
    root, fpath = _make_lib(tmp_path)
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"),
                         vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    emb = get_embedder(fake=True)
    ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    res2 = ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    assert res2["added"] == 0 and res2["skipped"] == 1  # unchanged → skipped


def test_ingest_marks_removed_stale(tmp_path):
    root, fpath = _make_lib(tmp_path)
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"),
                         vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    emb = get_embedder(fake=True)
    ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    os.remove(fpath)
    res = ingest_root(root, idx, emb, extract_fn=lambda p: SAMPLE)
    assert res["staled"] == 1
    assert idx.authority_candidates("meet and confer", motion_type="compel", limit=5) == []
