"""Build/refresh the local CA case-law corpus from bulk data.

CLI: python -m icharlotte_core.legal_research.local_corpus.build --source {cap|cl|all}

CAP volumes are downloaded from static.case.law; CL bulk is streamed from S3 and
filtered to CA + post-cutoff. Both feed the same DB + vectors.f16 via CorpusIndexer.
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import os
import time
from typing import Any

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.authority_signals import build_signals
from icharlotte_core.legal_research.local_corpus.embedder import Embedder, OnnxEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.loaders import (
    cap_loader,
    cl_bulk_loader,
    parenthetical_loader,
)

logger = logging.getLogger(__name__)

CAP_CUTOFF_DATE = "2018-01-01"   # overlap buffer; CAP wins dedup so CL only fills the gap

# CA reporter series on static.case.law and their volume counts (see spec).
CAP_REPORTERS = {
    "cal": 219, "cal-2d": 71, "cal-3d": 54, "cal-4th": 63, "cal-5th": 1,
    "cal-app": 140, "cal-app-2d": 276, "cal-app-3d": 235, "cal-app-4th": 248,
    "cal-app-5th": 11, "cal-rptr-3d": 56, "cal-unrep": 7,
}


def _install_metadata_helpers() -> None:
    if hasattr(schema, "get_meta") and hasattr(schema, "set_meta"):
        return

    def _ensure_meta_table(con) -> None:
        con.execute(
            "CREATE TABLE IF NOT EXISTS corpus_metadata ("
            "key TEXT PRIMARY KEY, "
            "value TEXT NOT NULL)"
        )

    def get_meta(con) -> dict[str, str]:
        _ensure_meta_table(con)
        return {
            str(row["key"]): str(row["value"])
            for row in con.execute("SELECT key, value FROM corpus_metadata")
        }

    def set_meta(con, **values: str) -> None:
        _ensure_meta_table(con)
        for key, value in values.items():
            con.execute(
                "INSERT OR REPLACE INTO corpus_metadata (key, value) VALUES (?, ?)",
                (str(key), "" if value is None else str(value)),
            )
        con.commit()

    schema.get_meta = get_meta
    schema.set_meta = set_meta


_install_metadata_helpers()


def _write_corpus_metadata(con, *, vectors_path: str, cl_snapshot_date: str = "") -> None:
    existing = schema.get_meta(con)
    counts = {
        str(row["source"] or ""): int(row["n"])
        for row in con.execute("SELECT source, COUNT(*) AS n FROM cases GROUP BY source")
    }
    max_date = con.execute(
        "SELECT MAX(decision_date) FROM cases WHERE decision_date IS NOT NULL AND decision_date<>''"
    ).fetchone()[0] or ""
    case_count = con.execute("SELECT COUNT(*) FROM cases").fetchone()[0]
    passage_count = con.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
    schema.set_meta(
        con,
        built_at=time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
        source_counts=json.dumps(counts, sort_keys=True),
        max_decision_date=max_date,
        case_count=str(case_count),
        passage_count=str(passage_count),
        vector_bytes=str(os.path.getsize(vectors_path) if os.path.exists(vectors_path) else 0),
        cl_snapshot_date=cl_snapshot_date or existing.get("cl_snapshot_date", ""),
    )


def _default_paths() -> tuple[str, str]:
    from icharlotte_core.config import CASELAW_DATA_DIR
    return (os.path.join(CASELAW_DATA_DIR, "corpus.db"),
            os.path.join(CASELAW_DATA_DIR, "vectors.f16"))


def _clear_temp(*paths: str) -> None:
    for p in paths:
        for suffix in ("", "-wal", "-shm"):
            try:
                os.remove(p + suffix)
            except OSError:
                pass


def _finish_build(con, *, db_tmp: str, vec_tmp: str, db_path: str, vectors_path: str) -> None:
    """Checkpoint, close, and atomically move temp build outputs into place.

    Building into temp paths keeps a half-built corpus invisible to the wizard
    (which requires BOTH final files) and leaves clean state on a crash.
    """
    build_signals(con)
    try:
        con.execute("PRAGMA wal_checkpoint(TRUNCATE)")
    except Exception:
        logger.warning("wal_checkpoint failed", exc_info=True)
    con.close()
    os.replace(vec_tmp, vectors_path)   # vectors first...
    os.replace(db_tmp, db_path)         # ...db last (its presence flips _corpus_available)
    _clear_temp(db_tmp)


def build_from_cap_zips(zip_paths: list[str], *, db_path: str, vectors_path: str,
                        embedder: Embedder, embed: bool = True,
                        embed_year_cutoff: int | None = None,
                        resume: bool | None = None) -> dict[str, Any]:
    """Build (or resume) the corpus from CAP volume ZIPs.

    Volume-checkpointed: each volume is committed atomically, so a crash/reboot
    mid-build resumes from the last completed volume. ``resume`` defaults to
    auto (resume iff a prior ``.building`` DB exists). ``embed_year_cutoff``
    limits semantic embedding to cases from that year onward (older cases stay
    keyword-searchable).
    """
    db_tmp, vec_tmp = db_path + ".building", vectors_path + ".building"
    if resume is None:
        resume = os.path.exists(db_tmp)
    if not resume:
        _clear_temp(db_tmp, vec_tmp)
    con = schema.connect(db_tmp)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec_tmp, embedder=embedder, embed=embed,
                        embed_year_cutoff=embed_year_cutoff, resume=resume)
    if resume:
        logger.info("CAP ingest: RESUMING — %d volumes already done", len(idx._done_volumes))
    n_cases = 0
    total = len(zip_paths)
    for i, zp in enumerate(zip_paths, 1):
        name = os.path.basename(zp)
        if idx.is_volume_done(name):
            continue
        try:
            with open(zp, "rb") as f:
                data = f.read()
            for case, passages in cap_loader.iter_cases_from_zip(data):
                if idx.add(case, passages):
                    n_cases += 1
            idx.commit_volume(name)
        except Exception:
            logger.warning("CAP ingest failed for %s; rolling back volume", zp, exc_info=True)
            idx.abort_volume()
        if i % 25 == 0 or i == total:
            logger.info("CAP ingest: %d/%d volumes done, %d new cases this run", i, total, n_cases)
    logger.info("CAP ingest: finalizing (good-law signals + atomic publish)...")
    idx.finalize()
    _finish_build(con, db_tmp=db_tmp, vec_tmp=vec_tmp, db_path=db_path, vectors_path=vectors_path)
    logger.info("CAP ingest: DONE — %d new cases this run", n_cases)
    return {"cases": n_cases}


def build_from_cl_streams(*, courts_stream, clusters_stream, opinions_stream,
                          db_path: str, vectors_path: str, embedder: Embedder,
                          embed: bool = True,
                          cutoff_date: str = CAP_CUTOFF_DATE) -> dict[str, Any]:
    db_tmp, vec_tmp = db_path + ".building", vectors_path + ".building"
    _clear_temp(db_tmp, vec_tmp)
    con = schema.connect(db_tmp)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec_tmp, embedder=embedder, embed=embed)
    n_cases = 0
    for case, passages in cl_bulk_loader.iter_recent_ca_cases(
        courts_stream=courts_stream, clusters_stream=clusters_stream,
        opinions_stream=opinions_stream, cutoff_date=cutoff_date,
    ):
        if idx.add(case, passages):
            n_cases += 1
    idx.finalize()
    _finish_build(con, db_tmp=db_tmp, vec_tmp=vec_tmp, db_path=db_path, vectors_path=vectors_path)
    return {"cases": n_cases}


def append_cl_to_corpus(*, citations_stream, clusters_stream, opinions_stream,
                        db_path: str, vectors_path: str, embedder: Embedder,
                        cutoff_date: str = "2017-01-01", published_only: bool = True,
                        embed: bool = False) -> dict[str, Any]:
    """Append recent CA cases from CourtListener bulk into an EXISTING published
    corpus, IN PLACE.

    Additive only — never modifies/deletes existing CAP rows (build_signals only
    updates the soft citation_count). CL cases dedup against existing cases by
    citation. With ``embed=False`` (keyword-only) each CL passage gets a zero
    placeholder vector, preserving vec_row<->passage alignment with the existing
    embedded corpus; semantic can be backfilled later. Periodic commits make the
    long ~50 GB stream durable + re-runnable (a re-run skips already-added cases).
    """
    import os
    if not (os.path.exists(db_path) and os.path.exists(vectors_path)):
        raise FileNotFoundError(
            "append_cl_to_corpus requires an existing published corpus "
            f"({db_path} + {vectors_path}); build CAP first.")
    con = schema.connect(db_path)
    schema.create_schema(con)  # idempotent (ensures ingested_volumes etc. exist)
    # embed=False -> high cutoff so every CL case gets a zero placeholder vector;
    # embed=True -> embed all CL cases (None cutoff). Either way indexer.embed=True
    # so a vector row is written per passage (real or zero), keeping alignment.
    cutoff_year = None if embed else 9999
    idx = CorpusIndexer(con, vectors_path=vectors_path, embedder=embedder,
                        embed=True, embed_year_cutoff=cutoff_year, resume=True)
    pre_cases = con.execute("SELECT COUNT(*) FROM cases").fetchone()[0]
    n = 0
    for case, passages in cl_bulk_loader.iter_recent_ca_cases(
        citations_stream=citations_stream, clusters_stream=clusters_stream,
        opinions_stream=opinions_stream, cutoff_date=cutoff_date,
        published_only=published_only,
    ):
        if idx.add(case, passages):
            n += 1
            if n % 1000 == 0:
                idx.commit_volume(f"cl-batch-{n}")   # durable checkpoint (fsync + commit)
                logger.info("CL append: %d new CA cases added", n)
    idx.finalize()
    build_signals(con)
    con.commit()
    con.close()
    logger.info("CL append: DONE — %d new cases (corpus now %d cases)", n, pre_cases + n)
    return {"added": n}


def append_parentheticals_to_corpus(
    *,
    parentheticals_stream,
    opinions_stream,
    citations_stream,
    db_path: str,
    vectors_path: str,
    embedder: Embedder,
    snapshot_date: str,
    min_score: float = 0.5,
    max_per_case: int = 25,
    embed: bool = False,
    refresh_opinion_map: bool = False,
) -> dict[str, Any]:
    """Append CL parentheticals to existing local corpus cases, in place."""
    if not (os.path.exists(db_path) and os.path.exists(vectors_path)):
        raise FileNotFoundError(
            "append_parentheticals_to_corpus requires an existing published corpus "
            f"({db_path} + {vectors_path}); build CAP/CL cases first."
        )
    con = schema.connect(db_path)
    schema.create_schema(con)
    opinion_map = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=opinions_stream,
        snapshot_date=snapshot_date,
        refresh=refresh_opinion_map,
    )
    cluster_case_map = parenthetical_loader.build_cluster_case_map(
        con,
        citations_stream=citations_stream,
    )
    passages = parenthetical_loader.iter_parenthetical_passages(
        parentheticals_stream=parentheticals_stream,
        opinion_cluster_map=opinion_map,
        cluster_case_map=cluster_case_map,
        min_score=min_score,
        max_per_case=max_per_case,
    )

    idx = CorpusIndexer(
        con,
        vectors_path=vectors_path,
        embedder=embedder,
        embed=True,
        resume=True,
    )
    added = idx.add_passages(passages, embed=embed)
    idx.finalize()
    schema.set_meta(
        con,
        parentheticals_snapshot_date=snapshot_date,
        parentheticals_count=str(
            con.execute(
                "SELECT COUNT(*) FROM passages WHERE passage_type='parenthetical'"
            ).fetchone()[0]
        ),
        parentheticals_min_score=str(float(min_score)),
        parentheticals_max_per_case=str(int(max_per_case)),
    )
    _write_corpus_metadata(con, vectors_path=vectors_path)
    con.commit()
    con.close()
    logger.info("CL parentheticals: DONE - %d new passages", added)
    return {"added": added}


def _download_cap_volumes(scratch_dir: str, reporters: dict | None = None) -> list[str]:
    """Download every CA reporter volume ZIP to scratch_dir; skip existing.

    Idempotent: an already-downloaded volume is reused, so a re-run resumes.
    """
    import urllib.request
    reporters = reporters if reporters is not None else CAP_REPORTERS
    os.makedirs(scratch_dir, exist_ok=True)
    total = sum(reporters.values())
    paths: list[str] = []
    done = 0
    for rep, count in reporters.items():
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
                    done += 1
                    continue
            paths.append(dest)
            done += 1
            if done % 50 == 0 or done == total:
                logger.info("CAP download: %d/%d volumes", done, total)
    return paths


CL_BULK_BASE = "https://com-courtlistener-storage.s3-us-west-2.amazonaws.com/bulk-data"
CL_BULK_DATE = "2026-03-31"   # newest quarterly snapshot (browse the bucket for current)


def _default_opener(url: str, headers: dict, timeout: int):  # pragma: no cover - network
    import urllib.request
    req = urllib.request.Request(url, headers=headers)
    return urllib.request.urlopen(req, timeout=timeout)


class _ResumableHTTPReader(io.RawIOBase):
    """A read-only byte stream over an HTTP GET that transparently reconnects
    with a ``Range`` request when the connection drops, so a 50 GB transfer over
    hours survives transient drops. Presents a gap-free stream to bz2."""

    def __init__(self, url, *, opener=_default_opener, max_retries: int = 12,
                 timeout: int = 180, backoff: float = 2.0):
        self.url = url
        self._opener = opener
        self.max_retries = max_retries
        self.timeout = timeout
        self.backoff = backoff
        self.offset = 0          # absolute bytes delivered so far
        self.total = None        # full Content-Length (from the offset=0 response)
        self._resp = None
        self._open()

    def _open(self):
        headers = {}
        if self.offset:
            headers["Range"] = f"bytes={self.offset}-"
        self._resp = self._opener(self.url, headers, self.timeout)
        if self.offset == 0:
            try:
                cl = self._resp.headers.get("Content-Length")
                self.total = int(cl) if cl else None
            except Exception:
                self.total = None
        elif self.offset:
            # Defensive: a 200 (not 206) on a Range request would restart from 0
            # and corrupt the stream. Detect and refuse.
            status = getattr(self._resp, "status", None) or getattr(self._resp, "code", None)
            if status not in (206, None):
                raise IOError(f"Range request not honored (HTTP {status}); cannot resume safely")

    def readable(self):
        return True

    def read(self, size=-1):
        if size is None or size < 0:
            chunks = []
            while True:
                b = self.read(1 << 20)
                if not b:
                    break
                chunks.append(b)
            return b"".join(chunks)
        retries = 0
        while True:
            try:
                data = self._resp.read(size)
            except Exception:
                data = None
            if data:
                self.offset += len(data)
                return data
            if data == b"":
                # Clean end-of-response. If we know the total and haven't reached
                # it, the connection ended early — reconnect and continue.
                if self.total is None or self.offset >= self.total:
                    return b""
            # else: data is None (read raised) -> reconnect.
            if retries >= self.max_retries:
                raise IOError(f"resumable read failed at byte {self.offset} after {retries} retries")
            retries += 1
            wait = min(self.backoff * (2 ** (retries - 1)), 30.0) if self.backoff else 0.0
            if wait:
                time.sleep(wait)
            logger.warning("CL stream: connection dropped at byte %d; reconnecting (retry %d/%d)",
                           self.offset, retries, self.max_retries)
            try:
                self._open()
            except Exception:
                continue  # retry the reconnect

    def close(self):
        try:
            if self._resp is not None:
                self._resp.close()
        except Exception:
            pass
        super().close()


def _stream_cl_bulk(name: str, date: str):  # pragma: no cover - network
    """Open a CL bulk .csv.bz2 from S3 as a resumable, decompressed text stream."""
    import bz2
    url = f"{CL_BULK_BASE}/{name}-{date}.csv.bz2"
    logger.info("CL stream: opening %s", url)
    reader = _ResumableHTTPReader(url)
    return io.TextIOWrapper(bz2.BZ2File(reader), encoding="utf-8", errors="replace")


def run_cl_append(*, db_path: str, vectors_path: str, embedder: Embedder,
                  date: str = CL_BULK_DATE, cutoff_date: str = "2017-01-01",
                  embed: bool = False) -> dict[str, Any]:  # pragma: no cover - network
    """Stream the three CL bulk files from S3 and append recent CA cases."""
    return append_cl_to_corpus(
        citations_stream=_stream_cl_bulk("citations", date),
        clusters_stream=_stream_cl_bulk("opinion-clusters", date),
        opinions_stream=_stream_cl_bulk("opinions", date),   # ~50 GB, streamed
        db_path=db_path, vectors_path=vectors_path, embedder=embedder,
        cutoff_date=cutoff_date, embed=embed,
    )


def run_parenthetical_append(
    *,
    db_path: str,
    vectors_path: str,
    embedder: Embedder,
    date: str = CL_BULK_DATE,
    min_score: float = 0.5,
    max_per_case: int = 25,
    embed: bool = False,
    refresh_opinion_map: bool = False,
) -> dict[str, Any]:  # pragma: no cover - network
    return append_parentheticals_to_corpus(
        parentheticals_stream=_stream_cl_bulk("parentheticals", date),
        opinions_stream=(
            _stream_cl_bulk("opinions", date)
            if refresh_opinion_map
            else None
        ),
        citations_stream=_stream_cl_bulk("citations", date),
        db_path=db_path,
        vectors_path=vectors_path,
        embedder=embedder,
        snapshot_date=date,
        min_score=min_score,
        max_per_case=max_per_case,
        embed=embed,
        refresh_opinion_map=refresh_opinion_map,
    )


def main() -> None:  # pragma: no cover - CLI wrapper
    ap = argparse.ArgumentParser(
        description="Build the local CA case-law corpus. Re-run after an "
        "interruption to RESUME automatically (volume-checkpointed).")
    ap.add_argument("--source", choices=["cap", "cl", "parentheticals", "all"], default="all")
    ap.add_argument("--data-dir", default=None)
    ap.add_argument("--fts-only", action="store_true",
                    help="Skip semantic embedding (BM25 keyword search only; builds in minutes).")
    ap.add_argument("--embed-since", type=int, default=None, metavar="YEAR",
                    help="Only embed cases decided in YEAR or later; older cases stay "
                         "keyword-searchable with a zero placeholder vector (cuts build time).")
    ap.add_argument("--cl-date", default=CL_BULK_DATE,
                    help="CourtListener bulk snapshot date (YYYY-MM-DD) for --source cl.")
    ap.add_argument("--cl-cutoff", default="2017-01-01",
                    help="Only ingest CL cases decided on/after this date (fills the gap above CAP).")
    ap.add_argument("--cl-embed", action="store_true",
                    help="Semantically embed the CL recent cases too (~hours). Default: keyword-only.")
    ap.add_argument("--parentheticals-min-score", type=float, default=0.5)
    ap.add_argument("--parentheticals-max-per-case", type=int, default=25)
    ap.add_argument("--embed-parentheticals", action="store_true")
    ap.add_argument("--refresh-opinion-map", action="store_true")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

    if args.data_dir:
        db_path = os.path.join(args.data_dir, "corpus.db")
        vectors_path = os.path.join(args.data_dir, "vectors.f16")
    else:
        db_path, vectors_path = _default_paths()
    embed = not args.fts_only
    embedder = OnnxEmbedder()
    if args.fts_only:
        mode = "FTS5 keyword-only (no embedding)"
    elif args.embed_since:
        mode = f"keyword for all + semantic embedding for cases >= {args.embed_since}"
    else:
        mode = "keyword + semantic embedding (all years)"
    logger.info("Build mode: %s", mode)

    if args.source in ("cap", "all"):
        scratch = os.path.join(os.path.dirname(db_path), "_cap_scratch")
        zips = _download_cap_volumes(scratch)
        summary = build_from_cap_zips(zips, db_path=db_path, vectors_path=vectors_path,
                                      embedder=embedder, embed=embed,
                                      embed_year_cutoff=args.embed_since)
        logger.info("CAP ingest: %s cases", summary["cases"])
    if args.source in ("cl", "all"):
        logger.info("CL append: streaming bulk snapshot %s, cutoff %s, embed=%s",
                    args.cl_date, args.cl_cutoff, args.cl_embed)
        summary = run_cl_append(
            db_path=db_path, vectors_path=vectors_path, embedder=embedder,
            date=args.cl_date, cutoff_date=args.cl_cutoff, embed=args.cl_embed,
        )
        logger.info("CL append: %s new CA cases", summary["added"])
    if args.source == "parentheticals":
        logger.info(
            "CL parentheticals: streaming bulk snapshot %s, min_score=%s, "
            "max_per_case=%s, embed=%s, refresh_opinion_map=%s",
            args.cl_date,
            args.parentheticals_min_score,
            args.parentheticals_max_per_case,
            args.embed_parentheticals,
            args.refresh_opinion_map,
        )
        summary = run_parenthetical_append(
            db_path=db_path,
            vectors_path=vectors_path,
            embedder=embedder,
            date=args.cl_date,
            min_score=args.parentheticals_min_score,
            max_per_case=args.parentheticals_max_per_case,
            embed=args.embed_parentheticals,
            refresh_opinion_map=args.refresh_opinion_map,
        )
        logger.info("CL parentheticals: %s new passages", summary["added"])


if __name__ == "__main__":  # pragma: no cover
    main()
