"""CourtListener bulk CSV stream-filter -> normalized recent CA CaseRecords.

CL bulk is full-corpus CSV across several files. We identify California cases by
their CITATION REPORTER (the `citations` file: reporter starting "Cal."), which
avoids needing the 5 GB dockets file just to resolve a court. We then join:

    citations (cluster_id -> CA reporter cite)   ~121 MB
    opinion-clusters (cluster_id -> date_filed, case_name, citation_count,
                      precedential_status)         ~2.3 GB
    opinions (cluster_id -> opinion text)          ~50 GB  (streamed, never stored)

filter to CA + published + decided after `cutoff_date`, and yield CaseRecord +
PassageRecords. Callers pass decompressed text streams (build wraps bz2).
"""
from __future__ import annotations

import csv
import logging
import re
import sys
from typing import Iterator, TextIO

from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord
from icharlotte_core.legal_research.local_corpus.textproc import chunk_passages, normalize_text

logger = logging.getLogger(__name__)

# Opinion text columns in priority order (mirror courtlistener.py field priority).
_TEXT_COLS = ("plain_text", "html_with_citations", "html", "html_columbia",
              "html_lawbox", "xml_harvard")

# CL precedential_status values we treat as "published".
_PUBLISHED = {"Published"}

# Big CSV fields (opinion html/text) can exceed the default csv field-size limit.
try:
    csv.field_size_limit(sys.maxsize)
except OverflowError:  # pragma: no cover - platform-dependent
    csv.field_size_limit(2**31 - 1)


def _is_ca_reporter(reporter: str) -> bool:
    # CA official + Cal. Rptr. all start with "Cal." (Cal., Cal. 2d..5th,
    # Cal. App. ..., Cal. Rptr. ..., Cal. Unrep., Cal. App. Supp.).
    return (reporter or "").strip().startswith("Cal.")


def _prefer_official(cites: list[tuple[str, str]]) -> str:
    """From [(reporter, 'vol rep page'), ...] prefer a non-Rptr official Cal. cite."""
    if not cites:
        return ""
    official = [c for rep, c in cites if not rep.startswith("Cal. Rptr")]
    return (official[0] if official else cites[0][1])


def _ca_clusters_from_citations(citations_stream: TextIO) -> dict[str, str]:
    """cluster_id -> preferred CA reporter citation, for CA cases only."""
    by_cluster: dict[str, list[tuple[str, str]]] = {}
    for row in csv.DictReader(citations_stream):
        reporter = (row.get("reporter") or "").strip()
        if not _is_ca_reporter(reporter):
            continue
        cid = (row.get("cluster_id") or "").strip()
        vol = (row.get("volume") or "").strip()
        page = (row.get("page") or "").strip()
        if cid and vol and page:
            by_cluster.setdefault(cid, []).append((reporter, f"{vol} {reporter} {page}"))
    return {cid: _prefer_official(cites) for cid, cites in by_cluster.items()}


def _recent_published(clusters_stream: TextIO, ca_cites: dict[str, str],
                      cutoff: str, published_only: bool) -> dict[str, dict]:
    keep: dict[str, dict] = {}
    for row in csv.DictReader(clusters_stream):
        cid = (row.get("id") or "").strip()
        if cid not in ca_cites:
            continue
        date = (row.get("date_filed") or "").strip()
        if date < cutoff:
            continue
        if published_only and (row.get("precedential_status") or "").strip() not in _PUBLISHED:
            continue
        keep[cid] = row
    return keep


def _strip_html(s: str) -> str:
    return re.sub(r"<[^>]+>", "", s or "")


def iter_recent_ca_cases(
    *,
    citations_stream: TextIO,
    clusters_stream: TextIO,
    opinions_stream: TextIO,
    cutoff_date: str,
    published_only: bool = True,
) -> Iterator[tuple[CaseRecord, list[PassageRecord]]]:
    ca_cites = _ca_clusters_from_citations(citations_stream)
    if not ca_cites:
        return
    clusters = _recent_published(clusters_stream, ca_cites, cutoff_date, published_only)
    if not clusters:
        return

    # The opinions file has one row per opinion; a cluster may have several
    # (lead + concurrence/dissent). Accumulate text per cluster across rows.
    text_by_cluster: dict[str, str] = {}
    for row in csv.DictReader(opinions_stream):
        cid = (row.get("cluster_id") or "").strip()
        if cid not in clusters:
            continue
        for col in _TEXT_COLS:
            raw = row.get(col) or ""
            if raw:
                t = normalize_text(_strip_html(raw))
                if t:
                    text_by_cluster[cid] = (text_by_cluster.get(cid, "") + "\n\n" + t).strip()
                    break

    for cid, meta in clusters.items():
        text = text_by_cluster.get(cid, "")
        if not text:
            continue
        case_uid = f"cl:{cid}"
        date = (meta.get("date_filed") or "").strip()
        try:
            cc = int(meta.get("citation_count") or 0)
        except (TypeError, ValueError):
            cc = None
        name = meta.get("case_name") or meta.get("case_name_short") or ""
        status = (meta.get("precedential_status") or "").strip()
        rec = CaseRecord(
            case_uid=case_uid, source="cl",
            name=name, name_abbreviation=meta.get("case_name_short", "") or name,
            citation=ca_cites.get(cid, ""),
            court="", decision_date=date, year=(date[:4] if len(date) >= 4 else ""),
            url=f"https://www.courtlistener.com/opinion/{cid}/",
            full_text=text, citation_count=cc,
            published_status=status,
            citable=status in _PUBLISHED,
        )
        passages = [
            PassageRecord(passage_uid=f"{case_uid}#{i}", case_uid=case_uid,
                          ordinal=i, text=chunk, page_label="")
            for i, chunk in enumerate(chunk_passages(text))
        ]
        yield rec, passages
