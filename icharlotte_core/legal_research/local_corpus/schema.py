"""SQLite schema + connection helper for the local case-law corpus."""
from __future__ import annotations

import os
import sqlite3

_DDL = """
CREATE TABLE IF NOT EXISTS cases (
    case_uid            TEXT PRIMARY KEY,
    source              TEXT NOT NULL,
    name                TEXT,
    name_abbreviation   TEXT,
    citation            TEXT,
    parallel_citations  TEXT,
    court               TEXT,
    decision_date       TEXT,
    year                TEXT,
    docket_number       TEXT,
    url                 TEXT,
    full_text           TEXT,
    citation_count      INTEGER,
    latest_citing_year  TEXT,
    cites_to            TEXT
);
CREATE INDEX IF NOT EXISTS idx_cases_citation ON cases(citation);

CREATE TABLE IF NOT EXISTS passages (
    passage_uid  TEXT PRIMARY KEY,
    case_uid     TEXT NOT NULL,
    ordinal      INTEGER NOT NULL,
    text         TEXT NOT NULL,
    page_label   TEXT,
    vec_row      INTEGER
);
CREATE INDEX IF NOT EXISTS idx_passages_case ON passages(case_uid);
CREATE INDEX IF NOT EXISTS idx_passages_vec  ON passages(vec_row);

CREATE TABLE IF NOT EXISTS citation_edges (
    from_case_uid  TEXT NOT NULL,
    to_citation    TEXT NOT NULL,
    weight         INTEGER DEFAULT 1
);
CREATE INDEX IF NOT EXISTS idx_edges_to ON citation_edges(to_citation);

CREATE VIRTUAL TABLE IF NOT EXISTS passages_fts USING fts5(
    text,
    content=''        -- external-content-less; we store text here directly
);
"""


def connect(db_path: str) -> sqlite3.Connection:
    """Open (creating parent dirs) a corpus DB with row dict access."""
    if db_path != ":memory:":
        os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
    con = sqlite3.connect(db_path)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    return con


def create_schema(con: sqlite3.Connection) -> None:
    con.executescript(_DDL)
    con.commit()
