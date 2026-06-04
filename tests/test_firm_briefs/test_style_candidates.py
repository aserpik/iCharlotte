import numpy as np
import pytest
from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite


def _unit(*vals):
    v = np.zeros(384, dtype=np.float32)
    for i, x in vals:
        v[i] = x
    n = np.linalg.norm(v) or 1.0
    return v / n


@pytest.fixture
def idx(tmp_path):
    ix = FirmBriefIndex(db_path=str(tmp_path / "fb.db"), vectors_path=str(tmp_path / "v.f16"))
    ix.create_schema()
    return ix


def _add(idx, path, side, vec, char_len=5000, ocr=0.0):
    idx.upsert_brief(path=path, content_hash="h", motion_type="compel", side=side,
                     heading="", profile="p", profile_vec=vec, char_len=char_len,
                     ocr_ratio=ocr, cites=[HarvestedCite(reporter_citation="1 Cal.5th 1",
                                                         norm_cite="1cal.5th1", proposition="x")])


def test_picks_most_similar_same_side(idx):
    _add(idx, "a.pdf", "opposition", _unit((0, 1.0)))
    _add(idx, "b.pdf", "opposition", _unit((1, 1.0)))
    _add(idx, "c.pdf", "opposition", _unit((2, 1.0)))
    res = idx.style_candidates(_unit((1, 1.0)), motion_type="compel", side="opposition", k=2)
    assert res[0]["path"] == "b.pdf"            # exact match ranks first
    assert len(res) == 2


def test_side_filter_excludes_other_side(idx):
    _add(idx, "moving.pdf", "moving", _unit((1, 1.0)))
    _add(idx, "opp.pdf", "opposition", _unit((1, 1.0)))
    res = idx.style_candidates(_unit((1, 1.0)), motion_type="compel", side="opposition", k=5)
    assert [r["path"] for r in res] == ["opp.pdf"]


def test_quality_penalty_downranks_tiny_noisy(idx):
    # Both equally similar, but 'tiny' is short + high OCR noise -> should rank lower.
    _add(idx, "good.pdf", "opposition", _unit((1, 1.0)), char_len=8000, ocr=0.0)
    _add(idx, "tiny.pdf", "opposition", _unit((1, 1.0)), char_len=300, ocr=0.4)
    res = idx.style_candidates(_unit((1, 1.0)), motion_type="compel", side="opposition", k=2)
    assert res[0]["path"] == "good.pdf"


def test_dedupes_versioned_copies(idx):
    _add(idx, "Brief.pdf", "opposition", _unit((1, 1.0)))
    _add(idx, "Brief_1.pdf", "opposition", _unit((1, 1.0)))
    res = idx.style_candidates(_unit((1, 1.0)), motion_type="compel", side="opposition", k=5)
    assert len(res) == 1   # _1 versioned duplicate collapsed
