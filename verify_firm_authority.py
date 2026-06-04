"""End-to-end check: firm index -> resolve cites against local corpus -> verified authority."""
import sys
sys.path.insert(0, r"C:\geminiterminal2")

from icharlotte_core.firm_briefs import factory
from icharlotte_core.firm_briefs.provider import FirmAuthorityProvider
from icharlotte_core.ui.wizard.pages.oppose_motion_page import _make_local_corpus

idx = factory.make_index()
print("firm index available:", idx is not None, "| stats:", idx.stats() if idx else None)
corpus = _make_local_corpus()
print("local corpus available:", corpus is not None)

prov = FirmAuthorityProvider(idx, corpus, cl_client=None)
tests = [
    ("a party must meet and confer in good faith before moving to compel further discovery", "compel", "opposition"),
    ("opposing summary judgment requires a triable issue of material fact", "msj", "opposition"),
    ("demurrer should be overruled with leave to amend", "demurrer", "opposition"),
]
for prop, mt, side in tests:
    cands = prov.candidates_for(prop, motion_type=mt, side=side, limit=4)
    print(f"\n[{mt}/{side}] {len(cands)} firm candidate(s):")
    for c in cands:
        name = c["case_name"][:40]
        print(f"  {c['source']}/{c['verification']:<15} {name:<42} {c['citation']:<20} text_chars={len(c.get('text',''))}")
