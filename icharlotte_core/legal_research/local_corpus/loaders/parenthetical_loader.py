"""CourtListener parenthetical bulk loader.

Parentheticals describe one opinion from another opinion. This loader attaches
them to cases already present in the local corpus and emits provenance-tagged
PassageRecord rows; it never mutates case full_text.
"""
from __future__ import annotations

import csv
import json
import logging
import re
import sys
from collections import defaultdict
from collections.abc import Iterator, Mapping
from typing import TextIO

from icharlotte_core.legal_research.local_corpus.models import PassageRecord
from icharlotte_core.legal_research.local_corpus.textproc import normalize_text

try:
    csv.field_size_limit(sys.maxsize)
except OverflowError:  # pragma: no cover - platform-dependent
    csv.field_size_limit(2**31 - 1)


_PARENTHETICAL_ORDINAL_BASE = 1_000_000
_OPINION_MAP_BATCH_SIZE = 100_000
logger = logging.getLogger(__name__)


def _cl_dict_reader(stream: TextIO) -> csv.DictReader:
    # CourtListener bulk CSV uses backslash-escaped quotes in large text/html
    # columns. The default csv dialect treats those quotes as structural and
    # misaligns rows, which makes opinion_id -> cluster_id maps garbage.
    return csv.DictReader(stream, escapechar="\\")


def _meta_key(snapshot_date: str) -> str:
    return f"opinion_map_complete:{snapshot_date}"


def _get_meta(con, key: str) -> str:
    con.execute(
        "CREATE TABLE IF NOT EXISTS corpus_meta ("
        "key TEXT PRIMARY KEY, value TEXT NOT NULL)"
    )
    row = con.execute("SELECT value FROM corpus_meta WHERE key=?", (key,)).fetchone()
    return str(row["value"]) if row else ""


def _set_meta(con, key: str, value: str) -> None:
    con.execute(
        "CREATE TABLE IF NOT EXISTS corpus_meta ("
        "key TEXT PRIMARY KEY, value TEXT NOT NULL)"
    )
    con.execute(
        "INSERT OR REPLACE INTO corpus_meta (key, value) VALUES (?, ?)",
        (key, value),
    )


class OpinionClusterLookup(Mapping[str, str]):
    """DB-backed opinion_id -> cluster_id lookup for one snapshot."""

    def __init__(self, con, snapshot_date: str) -> None:
        self._con = con
        self._snapshot_date = snapshot_date

    def __getitem__(self, opinion_id: str) -> str:
        row = self._con.execute(
            "SELECT cluster_id FROM courtlistener_opinion_map "
            "WHERE opinion_id=? AND snapshot_date=?",
            (str(opinion_id), self._snapshot_date),
        ).fetchone()
        if row is None:
            raise KeyError(opinion_id)
        return str(row["cluster_id"])

    def __iter__(self) -> Iterator[str]:
        for row in self._con.execute(
            "SELECT opinion_id FROM courtlistener_opinion_map "
            "WHERE snapshot_date=? ORDER BY opinion_id",
            (self._snapshot_date,),
        ):
            yield str(row["opinion_id"])

    def __len__(self) -> int:
        row = self._con.execute(
            "SELECT COUNT(*) FROM courtlistener_opinion_map WHERE snapshot_date=?",
            (self._snapshot_date,),
        ).fetchone()
        return int(row[0])

    def get(self, opinion_id: str, default: str = "") -> str:
        row = self._con.execute(
            "SELECT cluster_id FROM courtlistener_opinion_map "
            "WHERE opinion_id=? AND snapshot_date=?",
            (str(opinion_id), self._snapshot_date),
        ).fetchone()
        if row is None:
            return default
        return str(row["cluster_id"])

    def __eq__(self, other: object) -> bool:
        if isinstance(other, Mapping):
            return dict(self.items()) == dict(other.items())
        return NotImplemented


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
        "SELECT case_uid, citation, parallel_citations FROM cases "
        "ORDER BY CASE WHEN source='cap' THEN 0 ELSE 1 END, case_uid"
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
) -> Mapping[str, str]:
    complete_key = _meta_key(snapshot_date)
    if not refresh:
        cached = con.execute(
            "SELECT 1 FROM courtlistener_opinion_map WHERE snapshot_date=? LIMIT 1",
            (snapshot_date,),
        ).fetchone()
        if cached and _get_meta(con, complete_key) != "0":
            return OpinionClusterLookup(con, snapshot_date)
    if opinions_stream is None:
        raise ValueError("opinions_stream is required when no cached opinion map exists")

    con.execute(
        "DELETE FROM courtlistener_opinion_map WHERE snapshot_date=?",
        (snapshot_date,),
    )
    _set_meta(con, complete_key, "0")
    con.commit()

    pending: list[tuple[str, str, str]] = []
    seen = 0
    inserted = 0
    for row in _cl_dict_reader(opinions_stream):
        seen += 1
        opinion_id = (row.get("id") or "").strip()
        cluster_id = (row.get("cluster_id") or "").strip()
        if not opinion_id or not cluster_id:
            continue
        pending.append((opinion_id, cluster_id, snapshot_date))
        if len(pending) >= _OPINION_MAP_BATCH_SIZE:
            con.executemany(
                "INSERT OR REPLACE INTO courtlistener_opinion_map "
                "(opinion_id, cluster_id, snapshot_date) VALUES (?,?,?)",
                pending,
            )
            inserted += len(pending)
            pending.clear()
            con.commit()
            logger.info(
                "CL parentheticals: cached %d opinion map rows after %d opinion rows",
                inserted,
                seen,
            )
    if pending:
        con.executemany(
            "INSERT OR REPLACE INTO courtlistener_opinion_map "
            "(opinion_id, cluster_id, snapshot_date) VALUES (?,?,?)",
            pending,
        )
        inserted += len(pending)
    _set_meta(con, complete_key, "1")
    con.commit()
    logger.info("CL parentheticals: cached %d opinion map rows total", inserted)
    return OpinionClusterLookup(con, snapshot_date)


def build_cluster_case_map(con, *, citations_stream: TextIO | None) -> dict[str, str]:
    out: dict[str, str] = {}
    for row in con.execute("SELECT case_uid FROM cases WHERE case_uid LIKE 'cl:%'"):
        uid = str(row["case_uid"])
        out[uid.split(":", 1)[1]] = uid

    if citations_stream is None:
        return out

    by_citation = _local_citation_index(con)
    for row in _cl_dict_reader(citations_stream):
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


def _id_sort_key(value: str) -> tuple[int, int, str]:
    value = value or ""
    if value.isdigit():
        return (0, int(value), value)
    return (1, 0, value)


def _parenthetical_sort_key(item: tuple[float, str, PassageRecord]):
    return (-item[0], _id_sort_key(item[1]))


def iter_parenthetical_passages(
    *,
    parentheticals_stream: TextIO,
    opinion_cluster_map: Mapping[str, str],
    cluster_case_map: Mapping[str, str],
    min_score: float = 0.5,
    max_per_case: int = 25,
) -> Iterator[PassageRecord]:
    max_per_case = max(1, int(max_per_case))
    min_score = float(min_score)
    buckets: dict[str, list[tuple[float, str, PassageRecord]]] = defaultdict(list)

    for row in _cl_dict_reader(parentheticals_stream):
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
        bucket = buckets[case_uid]
        bucket.append((score, parenthetical_id, passage))
        if len(bucket) > max_per_case:
            bucket.sort(key=_parenthetical_sort_key)
            del bucket[max_per_case:]

    for case_uid in sorted(buckets):
        selected = sorted(buckets[case_uid], key=_parenthetical_sort_key)
        for offset, (_score, _pid, passage) in enumerate(selected):
            passage.ordinal = _PARENTHETICAL_ORDINAL_BASE + offset
            yield passage
