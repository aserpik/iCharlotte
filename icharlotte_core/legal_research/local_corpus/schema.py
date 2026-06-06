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
    passage_uid            TEXT PRIMARY KEY,
    case_uid               TEXT NOT NULL,
    ordinal                INTEGER NOT NULL,
    text                   TEXT NOT NULL,
    page_label             TEXT,
    vec_row                INTEGER,
    passage_type           TEXT DEFAULT 'opinion',
    source                 TEXT DEFAULT '',
    parenthetical_id       TEXT DEFAULT '',
    parenthetical_score    REAL,
    described_opinion_id   TEXT DEFAULT '',
    describing_opinion_id  TEXT DEFAULT '',
    describing_cluster_id  TEXT DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_passages_case ON passages(case_uid);
CREATE INDEX IF NOT EXISTS idx_passages_vec  ON passages(vec_row);
CREATE INDEX IF NOT EXISTS idx_passages_type ON passages(passage_type);
CREATE INDEX IF NOT EXISTS idx_passages_parenthetical ON passages(parenthetical_id);

CREATE TABLE IF NOT EXISTS courtlistener_opinion_map (
    opinion_id     TEXT PRIMARY KEY,
    cluster_id     TEXT NOT NULL,
    snapshot_date  TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_cl_opinion_map_cluster ON courtlistener_opinion_map(cluster_id);
CREATE INDEX IF NOT EXISTS idx_cl_opinion_map_snapshot ON courtlistener_opinion_map(snapshot_date);

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

-- Volumes fully ingested + committed, for resumable builds. A volume is the
-- checkpoint unit: on restart, already-listed volumes are skipped.
CREATE TABLE IF NOT EXISTS ingested_volumes (
    name TEXT PRIMARY KEY
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


def ensure_runtime_schema(con: sqlite3.Connection) -> None:
    """Ensure additive runtime schema changes exist on an existing corpus DB."""
    con.executescript(_DDL)
    tables = {
        r[0]
        for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type IN ('table','view')"
        ).fetchall()
    }

    if "passages" in tables:
        passage_columns = {
            r[1] for r in con.execute("PRAGMA table_info(passages)").fetchall()
        }
        passage_additions = {
            "passage_type": "TEXT DEFAULT 'opinion'",
            "source": "TEXT DEFAULT ''",
            "parenthetical_id": "TEXT DEFAULT ''",
            "parenthetical_score": "REAL",
            "described_opinion_id": "TEXT DEFAULT ''",
            "describing_opinion_id": "TEXT DEFAULT ''",
            "describing_cluster_id": "TEXT DEFAULT ''",
        }
        for name, ddl in passage_additions.items():
            if name not in passage_columns:
                con.execute(f"ALTER TABLE passages ADD COLUMN {name} {ddl}")
        con.execute("CREATE INDEX IF NOT EXISTS idx_passages_type ON passages(passage_type)")
        con.execute(
            "CREATE INDEX IF NOT EXISTS idx_passages_parenthetical ON passages(parenthetical_id)"
        )

    con.execute(
        "CREATE TABLE IF NOT EXISTS courtlistener_opinion_map ("
        "opinion_id TEXT PRIMARY KEY, "
        "cluster_id TEXT NOT NULL, "
        "snapshot_date TEXT NOT NULL)"
    )
    con.execute(
        "CREATE INDEX IF NOT EXISTS idx_cl_opinion_map_cluster "
        "ON courtlistener_opinion_map(cluster_id)"
    )
    con.execute(
        "CREATE INDEX IF NOT EXISTS idx_cl_opinion_map_snapshot "
        "ON courtlistener_opinion_map(snapshot_date)"
    )
    con.commit()
