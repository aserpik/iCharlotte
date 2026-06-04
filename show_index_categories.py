import sys
sys.path.insert(0, r"C:\geminiterminal2")
from collections import defaultdict
from icharlotte_core.firm_briefs import factory
from icharlotte_core.firm_briefs.index import FirmBriefIndex

db, vec = factory.index_paths()
idx = FirmBriefIndex(db_path=db, vectors_path=vec)
idx.create_schema()
con = idx._conn()
rows = con.execute(
    "SELECT motion_type, side, COUNT(*) c, "
    "SUM((SELECT COUNT(*) FROM citations WHERE brief_id=briefs.id)) cites "
    "FROM briefs WHERE status='ok' GROUP BY motion_type, side"
).fetchall()

bytype = defaultdict(dict)
for r in rows:
    bytype[r["motion_type"]][r["side"]] = (r["c"], r["cites"] or 0)

order = sorted(bytype, key=lambda t: -sum(v[0] for v in bytype[t].values()))
print(f"{'MOTION TYPE':<20}{'moving':>16}{'opposition':>16}{'reply':>12}")
print("-" * 64)
for t in order:
    d = bytype[t]
    def cell(s):
        return f"{d[s][0]}br/{d[s][1]}c" if s in d else "-"
    print(f"{t:<20}{cell('moving'):>16}{cell('opposition'):>16}{cell('reply'):>12}")

tot = con.execute("SELECT COUNT(*) FROM briefs WHERE status='ok'").fetchone()[0]
tc = con.execute("SELECT COUNT(*) FROM citations").fetchone()[0]
print("-" * 64)
print(f"TOTAL: {tot} briefs, {tc} citations, {len(bytype)} motion types")
