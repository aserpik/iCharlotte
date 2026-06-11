import sys, os
sys.path.insert(0, r"C:\geminiterminal2")
from icharlotte_core.firm_briefs import factory
from icharlotte_core.firm_briefs.index import FirmBriefIndex
db, vec = factory.index_paths()
idx = FirmBriefIndex(db_path=db, vectors_path=vec); idx.create_schema()
con = idx._conn()
rows = con.execute("SELECT path, side FROM briefs WHERE status='ok' AND motion_type='other' ORDER BY path").fetchall()
print(f"{len(rows)} 'other' briefs:")
for r in rows:
    name = os.path.basename(r["path"]).split("__", 1)[-1]
    print(f"  [{r['side']:<10}] {name[:80]}")
