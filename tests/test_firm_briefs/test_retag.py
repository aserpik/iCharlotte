from icharlotte_core.firm_briefs.index import FirmBriefIndex
from icharlotte_core.firm_briefs.citation_harvest import HarvestedCite
from retag_firm_index import retag_other
import numpy as np


def _vec():
    v = np.ones(384, dtype=np.float32)
    return v / np.linalg.norm(v)


def _add(idx, path, mtype, side):
    idx.upsert_brief(path=path, content_hash="h", motion_type=mtype, side=side,
                     heading="", profile="p", profile_vec=_vec(), char_len=10,
                     ocr_ratio=0.0, cites=[HarvestedCite(reporter_citation="1 Cal.5th 1",
                                                         norm_cite="1cal.5th1", proposition="x")])


def test_retag_reclassifies_other(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"), vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    _add(idx, r"C:\lib\Motions - Other\072 - X__Motion for Leave to Conduct IME.pdf", "other", "moving")
    _add(idx, r"C:\lib\Motions - Other\072 - X__Motion to Tax Costs.pdf", "other", "moving")
    _add(idx, r"C:\lib\Motion - Compel\008__opp.pdf", "compel", "moving")  # not 'other' -> untouched
    changed = retag_other(idx)
    assert changed == 1  # only the IME row reclassifies; Tax Costs stays other
    con = idx._conn()
    types = dict(con.execute("SELECT motion_type, COUNT(*) FROM briefs GROUP BY motion_type").fetchall())
    assert types.get("ime") == 1
    assert types.get("compel") == 1
    assert types.get("other") == 1  # Tax Costs remains


def test_retag_idempotent(tmp_path):
    idx = FirmBriefIndex(db_path=str(tmp_path / "fb.db"), vectors_path=str(tmp_path / "v.f16"))
    idx.create_schema()
    _add(idx, r"C:\lib\Motions - Other\x__Motion for Reconsideration.pdf", "other", "moving")
    assert retag_other(idx) == 1
    assert retag_other(idx) == 0  # nothing left to change
