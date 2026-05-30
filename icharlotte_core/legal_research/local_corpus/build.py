"""Build/refresh the local CA case-law corpus from bulk data.

CLI: python -m icharlotte_core.legal_research.local_corpus.build --source {cap|cl|all}

CAP volumes are downloaded from static.case.law; CL bulk is streamed from S3 and
filtered to CA + post-cutoff. Both feed the same DB + vectors.f16 via CorpusIndexer.
"""
from __future__ import annotations

import argparse
import logging
import os
from typing import Any

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.authority_signals import build_signals
from icharlotte_core.legal_research.local_corpus.embedder import Embedder, OnnxEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.loaders import cap_loader, cl_bulk_loader

logger = logging.getLogger(__name__)

CAP_CUTOFF_DATE = "2018-01-01"   # overlap buffer; CAP wins dedup so CL only fills the gap

# CA reporter series on static.case.law and their volume counts (see spec).
CAP_REPORTERS = {
    "cal": 219, "cal-2d": 71, "cal-3d": 54, "cal-4th": 63, "cal-5th": 1,
    "cal-app": 140, "cal-app-2d": 276, "cal-app-3d": 235, "cal-app-4th": 248,
    "cal-app-5th": 11, "cal-rptr-3d": 56, "cal-unrep": 7,
}


def _default_paths() -> tuple[str, str]:
    from icharlotte_core.config import CASELAW_DATA_DIR
    return (os.path.join(CASELAW_DATA_DIR, "corpus.db"),
            os.path.join(CASELAW_DATA_DIR, "vectors.f16"))


def build_from_cap_zips(zip_paths: list[str], *, db_path: str, vectors_path: str,
                        embedder: Embedder) -> dict[str, Any]:
    con = schema.connect(db_path)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vectors_path, embedder=embedder)
    n_cases = 0
    for zp in zip_paths:
        with open(zp, "rb") as f:
            data = f.read()
        for case, passages in cap_loader.iter_cases_from_zip(data):
            if idx.add(case, passages):
                n_cases += 1
    idx.finalize()
    build_signals(con)
    con.close()
    return {"cases": n_cases}


def build_from_cl_streams(*, courts_stream, clusters_stream, opinions_stream,
                          db_path: str, vectors_path: str, embedder: Embedder,
                          cutoff_date: str = CAP_CUTOFF_DATE) -> dict[str, Any]:
    con = schema.connect(db_path)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vectors_path, embedder=embedder)
    n_cases = 0
    for case, passages in cl_bulk_loader.iter_recent_ca_cases(
        courts_stream=courts_stream, clusters_stream=clusters_stream,
        opinions_stream=opinions_stream, cutoff_date=cutoff_date,
    ):
        if idx.add(case, passages):
            n_cases += 1
    idx.finalize()
    build_signals(con)
    con.close()
    return {"cases": n_cases}


def _download_cap_volumes(scratch_dir: str) -> list[str]:  # pragma: no cover - network
    """Download every CA reporter volume ZIP to scratch_dir; skip existing."""
    import urllib.request
    os.makedirs(scratch_dir, exist_ok=True)
    paths: list[str] = []
    for rep, count in CAP_REPORTERS.items():
        for vol in range(1, count + 1):
            dest = os.path.join(scratch_dir, f"{rep}-{vol}.zip")
            if not os.path.exists(dest):
                url = f"https://static.case.law/{rep}/{vol}.zip"
                try:
                    req = urllib.request.Request(url, headers={"User-Agent": "iCharlotte/1.0"})
                    with urllib.request.urlopen(req, timeout=60) as r, open(dest, "wb") as out:
                        out.write(r.read())
                except Exception:
                    logger.warning("CAP download failed: %s", url, exc_info=True)
                    continue
            paths.append(dest)
    return paths


def main() -> None:  # pragma: no cover - CLI wrapper
    ap = argparse.ArgumentParser(description="Build the local CA case-law corpus")
    ap.add_argument("--source", choices=["cap", "cl", "all"], default="all")
    ap.add_argument("--data-dir", default=None)
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO)

    if args.data_dir:
        db_path = os.path.join(args.data_dir, "corpus.db")
        vectors_path = os.path.join(args.data_dir, "vectors.f16")
    else:
        db_path, vectors_path = _default_paths()
    embedder = OnnxEmbedder()

    if args.source in ("cap", "all"):
        scratch = os.path.join(os.path.dirname(db_path), "_cap_scratch")
        zips = _download_cap_volumes(scratch)
        summary = build_from_cap_zips(zips, db_path=db_path, vectors_path=vectors_path, embedder=embedder)
        logger.info("CAP ingest: %s cases", summary["cases"])
    if args.source in ("cl", "all"):
        logger.info("CL bulk ingest: stream CourtListener bulk CSVs into "
                    "build_from_cl_streams (see README for the exact stream wiring).")


if __name__ == "__main__":  # pragma: no cover
    main()
