"""CLI: python -m icharlotte_core.firm_briefs --root <path> [--root ...]"""
from __future__ import annotations

import argparse
import os

from icharlotte_core import config
from .factory import index_paths, DATA_DIR
from .index import FirmBriefIndex
from .embedding import get_embedder
from .ingest import ingest_root


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Ingest firm brief libraries.")
    ap.add_argument("--root", action="append", default=[], help="library root (repeatable)")
    ap.add_argument("--fake-embed", action="store_true", help="deterministic embedder (tests)")
    args = ap.parse_args(argv)

    roots = args.root or config.FIRM_BRIEFS_ROOTS
    if not roots:
        print("No roots given and config.FIRM_BRIEFS_ROOTS is empty.")
        return 2
    os.makedirs(DATA_DIR, exist_ok=True)
    db, vec = index_paths()
    index = FirmBriefIndex(db_path=db, vectors_path=vec)
    index.create_schema()
    embedder = get_embedder(fake=args.fake_embed)
    totals = {"added": 0, "updated": 0, "skipped": 0, "failed": 0, "staled": 0}
    for root in roots:
        print(f"Ingesting {root} ...")
        res = ingest_root(root, index, embedder, on_progress=lambda m: print(m))
        for k in totals:
            totals[k] += res.get(k, 0)
        print(f"  {res}")
    print(f"DONE {totals}  stats={index.stats()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
