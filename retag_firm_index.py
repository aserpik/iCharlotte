"""Re-tag existing 'other' firm-index rows by normalizing their filename.

In-place UPDATE only: no re-extraction, no re-embedding (profile vectors and
citations are unchanged). Idempotent. Usage:
    python retag_firm_index.py
"""
import os
import sys

sys.path.insert(0, r"C:\geminiterminal2")

from icharlotte_core.firm_briefs.motion_taxonomy import normalize_motion_type


def _type_from_path(path: str) -> str:
    name = os.path.basename(path).split("__", 1)[-1]
    return normalize_motion_type(os.path.splitext(name)[0])


def retag_other(index) -> int:
    """Reclassify rows currently tagged 'other' using the filename. Returns count changed."""
    con = index._conn()
    rows = con.execute("SELECT id, path FROM briefs WHERE motion_type='other'").fetchall()
    changed = 0
    for r in rows:
        new_type = _type_from_path(r["path"])
        if new_type and new_type != "other":
            con.execute("UPDATE briefs SET motion_type=? WHERE id=?", (new_type, r["id"]))
            changed += 1
    con.commit()
    return changed


def main() -> int:
    from icharlotte_core.firm_briefs import factory
    from icharlotte_core.firm_briefs.index import FirmBriefIndex
    if not factory.index_available():
        print("No firm index built; nothing to re-tag.")
        return 1
    db, vec = factory.index_paths()
    idx = FirmBriefIndex(db_path=db, vectors_path=vec)
    idx.create_schema()
    before = dict(idx._conn().execute(
        "SELECT motion_type, COUNT(*) FROM briefs WHERE status='ok' GROUP BY motion_type").fetchall())
    n = retag_other(idx)
    after = dict(idx._conn().execute(
        "SELECT motion_type, COUNT(*) FROM briefs WHERE status='ok' GROUP BY motion_type").fetchall())
    print(f"Re-tagged {n} 'other' rows.")
    print("before:", before)
    print("after: ", after)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
