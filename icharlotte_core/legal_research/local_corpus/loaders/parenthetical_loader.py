"""CourtListener parenthetical bulk loader.

Parentheticals describe one opinion from another opinion. This loader attaches
them to cases already present in the local corpus and emits provenance-tagged
PassageRecord rows; it never mutates case full_text.
"""
from __future__ import annotations

import csv
import json
import re
import sys
from collections import defaultdict
from typing import Iterator, TextIO

from icharlotte_core.legal_research.local_corpus.models import PassageRecord
from icharlotte_core.legal_research.local_corpus.textproc import normalize_text

try:
    csv.field_size_limit(sys.maxsize)
except OverflowError:  # pragma: no cover - platform-dependent
    csv.field_size_limit(2**31 - 1)


_PARENTHETICAL_ORDINAL_BASE = 1_000_000


def _norm_citation(citation: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", (citation or "").lower())


def _is_ca_reporter(reporter: str) -> bool:
    return (reporter or "").strip().startswith("Cal.")


def _citation_from_row(row: dict) -> str:
    reporter = (row.get("reporter") or "").strip()
    volume = (row.get("volume") or "").strip()
    page = (row.get("page") or "").strip()
    if not (reporter and volume and page):
        return ""
    return f"{volume} {reporter} {page}"


def _local_citation_index(con) -> dict[str, str]:
    out: dict[str, str] = {}
    rows = con.execute(
        "SELECT case_uid, citation, parallel_citations FROM cases"
    ).fetchall()
    for row in rows:
        citations = [row["citation"] or ""]
        try:
            citations.extend(json.loads(row["parallel_citations"] or "[]"))
        except (TypeError, ValueError):
            pass
        for citation in citations:
            norm = _norm_citation(citation)
            if norm and norm not in out:
                out[norm] = row["case_uid"]
    return out


def load_opinion_cluster_map(
    con,
    *,
    opinions_stream: TextIO | None,
    snapshot_date: str,
    refresh: bool = False,
) -> dict[str, str]:
    if not refresh:
        cached = {
            str(row["opinion_id"]): str(row["cluster_id"])
            for row in con.execute(
                "SELECT opinion_id, cluster_id FROM courtlistener_opinion_map "
                "WHERE snapshot_date=?",
                (snapshot_date,),
            )
        }
        if cached:
            return cached
    if opinions_stream is None:
        raise ValueError("opinions_stream is required when no cached opinion map exists")

    if refresh:
        con.execute(
            "DELETE FROM courtlistener_opinion_map WHERE snapshot_date=?",
            (snapshot_date,),
        )

    out: dict[str, str] = {}
    for row in csv.DictReader(opinions_stream):
        opinion_id = (row.get("id") or "").strip()
        cluster_id = (row.get("cluster_id") or "").strip()
        if not opinion_id or not cluster_id:
            continue
        out[opinion_id] = cluster_id
        con.execute(
            "INSERT OR REPLACE INTO courtlistener_opinion_map "
            "(opinion_id, cluster_id, snapshot_date) VALUES (?,?,?)",
            (opinion_id, cluster_id, snapshot_date),
        )
    con.commit()
    return out


def build_cluster_case_map(con, *, citations_stream: TextIO | None) -> dict[str, str]:
    out: dict[str, str] = {}
    for row in con.execute("SELECT case_uid FROM cases WHERE case_uid LIKE 'cl:%'"):
        uid = str(row["case_uid"])
        out[uid.split(":", 1)[1]] = uid

    if citations_stream is None:
        return out

    by_citation = _local_citation_index(con)
    for row in csv.DictReader(citations_stream):
        reporter = (row.get("reporter") or "").strip()
        if not _is_ca_reporter(reporter):
            continue
        cluster_id = (row.get("cluster_id") or "").strip()
        citation = _citation_from_row(row)
        case_uid = by_citation.get(_norm_citation(citation))
        if cluster_id and case_uid and cluster_id not in out:
            out[cluster_id] = case_uid
    return out


def _float_score(value: str) -> float:
    try:
        return float(value or 0.0)
    except (TypeError, ValueError):
        return 0.0


def iter_parenthetical_passages(
    *,
    parentheticals_stream: TextIO,
    opinion_cluster_map: dict[str, str],
    cluster_case_map: dict[str, str],
    min_score: float = 0.5,
    max_per_case: int = 25,
) -> Iterator[PassageRecord]:
    max_per_case = max(1, int(max_per_case))
    min_score = float(min_score)
    buckets: dict[str, list[tuple[float, str, PassageRecord]]] = defaultdict(list)

    for row in csv.DictReader(parentheticals_stream):
        parenthetical_id = (row.get("id") or "").strip()
        text = normalize_text(row.get("text") or "")
        score = _float_score(row.get("score") or "")
        described_opinion_id = (row.get("described_opinion_id") or "").strip()
        describing_opinion_id = (row.get("describing_opinion_id") or "").strip()
        described_cluster_id = opinion_cluster_map.get(described_opinion_id, "")
        case_uid = cluster_case_map.get(described_cluster_id, "")
        if not (parenthetical_id and text and case_uid):
            continue
        if score < min_score:
            continue
        describing_cluster_id = opinion_cluster_map.get(describing_opinion_id, "")
        passage = PassageRecord(
            passage_uid=f"{case_uid}#parenthetical:{parenthetical_id}",
            case_uid=case_uid,
            ordinal=_PARENTHETICAL_ORDINAL_BASE,
            text=text,
            passage_type="parenthetical",
            source="courtlistener_parenthetical",
            parenthetical_id=parenthetical_id,
            parenthetical_score=score,
            described_opinion_id=described_opinion_id,
            describing_opinion_id=describing_opinion_id,
            describing_cluster_id=describing_cluster_id,
        )
        buckets[case_uid].append((score, parenthetical_id, passage))

    for case_uid in sorted(buckets):
        selected = sorted(
            buckets[case_uid],
            key=lambda item: (-item[0], item[1]),
        )[:max_per_case]
        for offset, (_score, _pid, passage) in enumerate(selected):
            passage.ordinal = _PARENTHETICAL_ORDINAL_BASE + offset
            yield passage
