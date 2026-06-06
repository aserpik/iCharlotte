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
    opinion_id     TEXT NOT NULL,
    cluster_id     TEXT NOT NULL,
    snapshot_date  TEXT NOT NULL,
    PRIMARY KEY (opinion_id, snapshot_date)
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

CREATE TABLE IF NOT EXISTS corpus_meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
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
    ensure_runtime_schema(con)


def get_meta(con: sqlite3.Connection) -> dict[str, str]:
    con.execute(
        "CREATE TABLE IF NOT EXISTS corpus_meta ("
        "key TEXT PRIMARY KEY, "
        "value TEXT NOT NULL)"
    )
    return {
        str(row[0]): str(row[1])
        for row in con.execute("SELECT key, value FROM corpus_meta").fetchall()
    }


def set_meta(con: sqlite3.Connection, **items) -> None:
    con.execute(
        "CREATE TABLE IF NOT EXISTS corpus_meta ("
        "key TEXT PRIMARY KEY, "
        "value TEXT NOT NULL)"
    )
    for key, value in items.items():
        con.execute(
            "INSERT OR REPLACE INTO corpus_meta (key, value) VALUES (?, ?)",
            (str(key), "" if value is None else str(value)),
        )
    con.commit()


def _ensure_opinion_map_schema(con: sqlite3.Connection) -> None:
    con.execute(
        "CREATE TABLE IF NOT EXISTS courtlistener_opinion_map ("
        "opinion_id TEXT NOT NULL, "
        "cluster_id TEXT NOT NULL, "
        "snapshot_date TEXT NOT NULL, "
        "PRIMARY KEY (opinion_id, snapshot_date))"
    )

    pk_columns = [
        row[1]
        for row in sorted(
            con.execute("PRAGMA table_info(courtlistener_opinion_map)").fetchall(),
            key=lambda r: r[5],
        )
        if row[5]
    ]
    if pk_columns != ["opinion_id", "snapshot_date"]:
        con.execute("ALTER TABLE courtlistener_opinion_map RENAME TO courtlistener_opinion_map_old")
        old_columns = {
            row[1]
            for row in con.execute("PRAGMA table_info(courtlistener_opinion_map_old)").fetchall()
        }
        con.execute(
            "CREATE TABLE courtlistener_opinion_map ("
            "opinion_id TEXT NOT NULL, "
            "cluster_id TEXT NOT NULL, "
            "snapshot_date TEXT NOT NULL, "
            "PRIMARY KEY (opinion_id, snapshot_date))"
        )
        if "snapshot_date" in old_columns:
            con.execute(
                "INSERT OR REPLACE INTO courtlistener_opinion_map "
                "(opinion_id, cluster_id, snapshot_date) "
                "SELECT opinion_id, cluster_id, snapshot_date "
                "FROM courtlistener_opinion_map_old "
                "WHERE opinion_id IS NOT NULL AND opinion_id != '' "
                "AND snapshot_date IS NOT NULL AND snapshot_date != ''"
            )
        con.execute("DROP TABLE courtlistener_opinion_map_old")

    con.execute(
        "CREATE INDEX IF NOT EXISTS idx_cl_opinion_map_cluster "
        "ON courtlistener_opinion_map(cluster_id)"
    )
    con.execute(
        "CREATE INDEX IF NOT EXISTS idx_cl_opinion_map_snapshot "
        "ON courtlistener_opinion_map(snapshot_date)"
    )


def ensure_runtime_schema(con: sqlite3.Connection) -> None:
    """Ensure additive runtime schema changes exist on an existing corpus DB."""
    con.executescript(
        """
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

        CREATE TABLE IF NOT EXISTS citation_edges (
            from_case_uid  TEXT NOT NULL,
            to_citation    TEXT NOT NULL,
            weight         INTEGER DEFAULT 1
        );
        CREATE INDEX IF NOT EXISTS idx_edges_to ON citation_edges(to_citation);

        CREATE VIRTUAL TABLE IF NOT EXISTS passages_fts USING fts5(
            text,
            content=''
        );

        CREATE TABLE IF NOT EXISTS ingested_volumes (
            name TEXT PRIMARY KEY
        );

        CREATE TABLE IF NOT EXISTS corpus_meta (
            key   TEXT PRIMARY KEY,
            value TEXT NOT NULL
        );
        """
    )

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
    con.execute("CREATE INDEX IF NOT EXISTS idx_passages_case ON passages(case_uid)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_passages_vec ON passages(vec_row)")
    con.execute("CREATE INDEX IF NOT EXISTS idx_passages_type ON passages(passage_type)")
    con.execute(
        "CREATE INDEX IF NOT EXISTS idx_passages_parenthetical ON passages(parenthetical_id)"
    )

    _ensure_opinion_map_schema(con)
    con.commit()
