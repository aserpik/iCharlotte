"""One-off: clean existing citations.case_name in the firm index (strip signal
words / collapse whitespace). In-place UPDATE; idempotent."""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from icharlotte_core.firm_briefs.citation_harvest import clean_case_name


def clean_names(index) -> int:
    con = index._conn()
    rows = con.execute("SELECT id, case_name FROM citations").fetchall()
    changed = 0
    for r in rows:
        cleaned = clean_case_name(r["case_name"] or "")
        if cleaned != (r["case_name"] or ""):
            con.execute("UPDATE citations SET case_name=? WHERE id=?", (cleaned, r["id"]))
            changed += 1
    con.commit()
    return changed


def main() -> int:
    from icharlotte_core.firm_briefs import factory
    from icharlotte_core.firm_briefs.index import FirmBriefIndex
    if not factory.index_available():
        print("No index; nothing to clean."); return 1
    db, vec = factory.index_paths()
    idx = FirmBriefIndex(db_path=db, vectors_path=vec); idx.create_schema()
    print(f"Cleaned {clean_names(idx)} case names."); return 0


if __name__ == "__main__":
    raise SystemExit(main())
