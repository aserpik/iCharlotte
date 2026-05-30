"""CourtListener bulk CSV stream-filter -> normalized recent CA CaseRecords.

CL bulk is full-corpus, single-format CSV. We never store the 50 GB opinions
file: we stream it, keep only rows whose cluster is CA + post-cutoff, and
discard the rest. Callers pass decompressed text streams (build.py wraps bz2).
"""
from __future__ import annotations

import csv
import logging
import re
from typing import Iterator, TextIO

from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.textproc import chunk_passages, normalize_text

logger = logging.getLogger(__name__)

# Court ids on CourtListener that are California state courts.
CA_COURT_IDS = {
    "cal", "calctapp", "calag", "calapp", "calsuperct",
    "calapp1st", "calapp2nd", "calapp3rd", "calapp4th", "calapp5th", "calapp6th",
}

# Opinion text columns in priority order (mirror courtlistener.py field priority).
_TEXT_COLS = ("plain_text", "html_with_citations", "html", "html_columbia",
              "html_lawbox", "xml_harvard")


def _ca_court_ids_from_courts(courts_stream: TextIO) -> set[str]:
    found: set[str] = set()
    for row in csv.DictReader(courts_stream):
        cid = (row.get("id") or "").strip()
        if cid in CA_COURT_IDS:
            found.add(cid)
    # Always include the known set even if the courts file is sparse.
    return found or set(CA_COURT_IDS)


def _recent_ca_clusters(clusters_stream: TextIO, ca_courts: set[str], cutoff: str) -> dict[str, dict]:
    keep: dict[str, dict] = {}
    for row in csv.DictReader(clusters_stream):
        court = (row.get("court_id") or "").strip()
        date = (row.get("date_filed") or "").strip()
        if court in ca_courts and date >= cutoff:
            cid = (row.get("id") or "").strip()
            if cid:
                keep[cid] = row
    return keep


def _strip_html(s: str) -> str:
    return re.sub(r"<[^>]+>", "", s or "")


def iter_recent_ca_cases(
    *,
    courts_stream: TextIO,
    clusters_stream: TextIO,
    opinions_stream: TextIO,
    cutoff_date: str,
) -> Iterator[tuple[CaseRecord, list[PassageRecord]]]:
    ca_courts = _ca_court_ids_from_courts(courts_stream)
    clusters = _recent_ca_clusters(clusters_stream, ca_courts, cutoff_date)
    if not clusters:
        return

    for row in csv.DictReader(opinions_stream):
        cid = (row.get("cluster_id") or "").strip()
        meta = clusters.get(cid)
        if not meta:
            continue
        text = ""
        for col in _TEXT_COLS:
            raw = row.get(col) or ""
            if raw:
                text = normalize_text(_strip_html(raw))
                if text:
                    break
        if not text:
            continue
        case_uid = f"cl:{cid}"
        date = (meta.get("date_filed") or "").strip()
        rec = CaseRecord(
            case_uid=case_uid, source="cl",
            name=meta.get("case_name", ""), name_abbreviation=meta.get("case_name", ""),
            citation=(meta.get("citation") or "").strip(),
            court=(meta.get("court_id") or "").strip(),
            decision_date=date, year=(date[:4] if len(date) >= 4 else ""),
            url=f"https://www.courtlistener.com/opinion/{cid}/",
            full_text=text,
        )
        passages = [
            PassageRecord(passage_uid=f"{case_uid}#{i}", case_uid=case_uid,
                          ordinal=i, text=chunk, page_label="")
            for i, chunk in enumerate(chunk_passages(text))
        ]
        yield rec, passages
