# CourtListener Parentheticals Corpus Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add CourtListener bulk parentheticals to the existing local California case-law corpus as high-signal, provenance-tagged passages attached only to cases already present locally.

**Architecture:** Extend the existing `icharlotte_core/legal_research/local_corpus` package. Parentheticals become tagged `passages` rows with provenance columns, are appended under existing cases through a passage-only indexer path, and participate in corpus search without entering `cases.full_text` or citation verification text.

**Tech Stack:** Python 3.12, SQLite/FTS5, bz2 streamed CourtListener bulk CSV, existing `CorpusIndexer`, existing `FakeEmbedder`/`OnnxEmbedder`, pytest.

---

## Scope

Implement the approved design in `docs/superpowers/specs/2026-06-06-courtlistener-parentheticals-corpus-design.md`.

This plan does not add a new legal research source, does not rebuild CAP/CL case text, and does not change Word/document generation. It only extends the local corpus storage, ingest, and search behavior.

The worktree may contain unrelated dirty files. Every commit command in this plan stages only the files named in that task.

## File Structure

- Modify `icharlotte_core/legal_research/local_corpus/schema.py`
  - Add parenthetical passage columns and the `courtlistener_opinion_map` cache table.
  - Keep migrations additive in `ensure_runtime_schema`.
- Modify `icharlotte_core/legal_research/local_corpus/models.py`
  - Add optional parenthetical/provenance fields to `PassageRecord`.
- Modify `icharlotte_core/legal_research/local_corpus/indexer.py`
  - Persist parenthetical fields.
  - Add a passage-only append method for existing cases.
- Create `icharlotte_core/legal_research/local_corpus/loaders/parenthetical_loader.py`
  - Parse CourtListener opinion, citation, and parenthetical CSV streams.
  - Build opinion-id and cluster-to-local-case maps.
  - Emit capped, score-filtered parenthetical `PassageRecord` rows.
- Modify `icharlotte_core/legal_research/local_corpus/build.py`
  - Add `append_parentheticals_to_corpus`, network stream wiring, metadata, and CLI options.
- Modify `icharlotte_core/legal_research/local_corpus/corpus.py`
  - Include parenthetical passage matches in ranking and snippet selection.
  - Keep `get_opinion_text` unchanged.
- Modify `icharlotte_core/legal_research/models.py`
  - Add optional snippet provenance fields to `CaseResult`.
- Modify `icharlotte_core/legal_research/local_corpus/README.md`
  - Document parenthetical ingest, defaults, and verification boundary.
- Create `tests/test_legal_research/test_local_corpus/test_parentheticals_schema.py`
- Create `tests/test_legal_research/test_local_corpus/test_parenthetical_loader.py`
- Create `tests/test_legal_research/test_local_corpus/test_parenthetical_append.py`
- Modify `tests/test_legal_research/test_local_corpus/test_corpus.py`
- Modify `tests/test_opposition/test_local_case_verifier.py`

## Task 1: Schema, Models, And Passage-Only Indexing

**Files:**
- Modify: `icharlotte_core/legal_research/local_corpus/schema.py`
- Modify: `icharlotte_core/legal_research/local_corpus/models.py`
- Modify: `icharlotte_core/legal_research/local_corpus/indexer.py`
- Create: `tests/test_legal_research/test_local_corpus/test_parentheticals_schema.py`

- [ ] **Step 1: Write failing schema/model/indexer tests**

Create `tests/test_legal_research/test_local_corpus/test_parentheticals_schema.py`:

```python
import sqlite3

import numpy as np

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord


def test_schema_creates_parenthetical_passage_columns_and_opinion_map():
    con = sqlite3.connect(":memory:")
    schema.create_schema(con)

    passage_columns = {
        row[1] for row in con.execute("PRAGMA table_info(passages)").fetchall()
    }
    assert {
        "passage_type",
        "source",
        "parenthetical_id",
        "parenthetical_score",
        "described_opinion_id",
        "describing_opinion_id",
        "describing_cluster_id",
    }.issubset(passage_columns)

    tables = {
        row[0]
        for row in con.execute(
            "SELECT name FROM sqlite_master WHERE type IN ('table','view')"
        ).fetchall()
    }
    assert "courtlistener_opinion_map" in tables


def test_passage_record_accepts_parenthetical_metadata():
    passage = PassageRecord(
        passage_uid="cap:1#parenthetical:900",
        case_uid="cap:1",
        ordinal=1_000_000,
        text="describing the case as adopting a burden-shifting standard",
        passage_type="parenthetical",
        source="courtlistener_parenthetical",
        parenthetical_id="900",
        parenthetical_score=0.88,
        described_opinion_id="10",
        describing_opinion_id="20",
        describing_cluster_id="200",
    )

    assert passage.passage_type == "parenthetical"
    assert passage.parenthetical_id == "900"
    assert passage.parenthetical_score == 0.88
    assert passage.describing_cluster_id == "200"


def test_indexer_appends_parenthetical_passage_to_existing_case(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    con = schema.connect(db)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:1",
            source="cap",
            name="Aguilar v. Atlantic Richfield Co.",
            citation="25 Cal. 4th 826",
            decision_date="2001-06-14",
            year="2001",
            full_text="primary opinion text",
        ),
        [
            PassageRecord(
                passage_uid="cap:1#0",
                case_uid="cap:1",
                ordinal=0,
                text="primary opinion text",
            )
        ],
    )
    idx.finalize()
    con.close()

    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    added = idx.add_passages(
        [
            PassageRecord(
                passage_uid="cap:1#parenthetical:900",
                case_uid="cap:1",
                ordinal=1_000_000,
                text="describing Aguilar as setting the summary judgment burden",
                passage_type="parenthetical",
                source="courtlistener_parenthetical",
                parenthetical_id="900",
                parenthetical_score=0.91,
                described_opinion_id="10",
                describing_opinion_id="20",
                describing_cluster_id="200",
            )
        ],
        embed=False,
    )
    idx.finalize()

    assert added == 1
    row = con.execute(
        "SELECT * FROM passages WHERE passage_uid=?",
        ("cap:1#parenthetical:900",),
    ).fetchone()
    assert row["passage_type"] == "parenthetical"
    assert row["parenthetical_id"] == "900"
    assert row["parenthetical_score"] == 0.91
    assert row["described_opinion_id"] == "10"
    assert row["describing_opinion_id"] == "20"
    assert row["describing_cluster_id"] == "200"

    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 16)
    assert arr.shape[0] == 2
    assert np.allclose(arr[int(row["vec_row"])], 0.0)


def test_indexer_skips_duplicate_parenthetical_passage_uid(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    con = schema.connect(db)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(case_uid="cap:1", source="cap", citation="25 Cal. 4th 826"),
        [PassageRecord(passage_uid="cap:1#0", case_uid="cap:1", ordinal=0, text="x")],
    )
    idx.finalize()
    con.close()

    passage = PassageRecord(
        passage_uid="cap:1#parenthetical:900",
        case_uid="cap:1",
        ordinal=1_000_000,
        text="summary judgment burden",
        passage_type="parenthetical",
        parenthetical_id="900",
    )
    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    assert idx.add_passages([passage], embed=False) == 1
    idx.commit_volume("parentheticals-batch-1")
    assert idx.add_passages([passage], embed=False) == 0
    idx.finalize()

    assert con.execute(
        "SELECT COUNT(*) FROM passages WHERE parenthetical_id='900'"
    ).fetchone()[0] == 1
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_parentheticals_schema.py -q
```

Expected: FAIL because `PassageRecord` does not accept `passage_type` and `CorpusIndexer` has no `add_passages`.

- [ ] **Step 3: Add additive schema columns and cache table**

In `icharlotte_core/legal_research/local_corpus/schema.py`, update `_DDL` so `passages` has the new columns:

```python
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
```

Extend `ensure_runtime_schema` with this exact additive migration block after the `cases` column migration:

```python
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
```

- [ ] **Step 4: Add parenthetical fields to PassageRecord**

In `icharlotte_core/legal_research/local_corpus/models.py`, replace `PassageRecord` with:

```python
@dataclass
class PassageRecord:
    passage_uid: str          # f"{case_uid}#{ordinal}" or f"{case_uid}#parenthetical:{id}"
    case_uid: str
    ordinal: int
    text: str
    page_label: str = ""      # reporter page this passage starts on (pin-cite)
    vec_row: int | None = None  # row index into vectors.f16 (set by indexer)
    passage_type: str = "opinion"
    source: str = ""
    parenthetical_id: str = ""
    parenthetical_score: float | None = None
    described_opinion_id: str = ""
    describing_opinion_id: str = ""
    describing_cluster_id: str = ""
```

- [ ] **Step 5: Persist parenthetical fields and add passage-only append**

In `icharlotte_core/legal_research/local_corpus/indexer.py`, add this method to `CorpusIndexer`:

```python
    def add_passages(self, passages: Iterable[PassageRecord], *, embed: bool = True) -> int:
        """Buffer passages for cases that already exist in the corpus.

        Returns the number of newly buffered passages. Existing passage_uid rows
        are skipped, which makes parenthetical snapshot re-runs idempotent.
        """
        added = 0
        for p in passages:
            exists = self.con.execute(
                "SELECT 1 FROM passages WHERE passage_uid=?",
                (p.passage_uid,),
            ).fetchone()
            if exists:
                continue
            case_exists = self.con.execute(
                "SELECT 1 FROM cases WHERE case_uid=?",
                (p.case_uid,),
            ).fetchone()
            if not case_exists:
                continue
            self._pending.append((p, bool(embed)))
            added += 1
            if len(self._pending) >= _BATCH:
                self._flush()
        return added
```

In `_flush`, replace the passage insert with:

```python
            self.con.execute(
                "INSERT OR REPLACE INTO passages ("
                "passage_uid, case_uid, ordinal, text, page_label, vec_row, "
                "passage_type, source, parenthetical_id, parenthetical_score, "
                "described_opinion_id, describing_opinion_id, describing_cluster_id"
                ") VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                (
                    p.passage_uid,
                    p.case_uid,
                    p.ordinal,
                    p.text,
                    p.page_label,
                    vec_row,
                    p.passage_type,
                    p.source,
                    p.parenthetical_id,
                    p.parenthetical_score,
                    p.described_opinion_id,
                    p.describing_opinion_id,
                    p.describing_cluster_id,
                ),
            )
```

- [ ] **Step 6: Run tests to verify Task 1 passes**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_parentheticals_schema.py -q
```

Expected: PASS.

- [ ] **Step 7: Commit Task 1**

Run:

```powershell
git add -- icharlotte_core/legal_research/local_corpus/schema.py icharlotte_core/legal_research/local_corpus/models.py icharlotte_core/legal_research/local_corpus/indexer.py tests/test_legal_research/test_local_corpus/test_parentheticals_schema.py
git diff --cached --check
git commit -m "feat(corpus): store parenthetical passage metadata"
```

Expected: commit succeeds.

## Task 2: Parenthetical Loader And Local Case Mapping

**Files:**
- Create: `icharlotte_core/legal_research/local_corpus/loaders/parenthetical_loader.py`
- Create: `tests/test_legal_research/test_local_corpus/test_parenthetical_loader.py`

- [ ] **Step 1: Write failing loader tests**

Create `tests/test_legal_research/test_local_corpus/test_parenthetical_loader.py`:

```python
import csv
import io

from icharlotte_core.legal_research.local_corpus import schema
from icharlotte_core.legal_research.local_corpus.loaders import parenthetical_loader
from icharlotte_core.legal_research.local_corpus.models import CaseRecord


def _csv(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    for row in rows:
        writer.writerow(row)
    buf.seek(0)
    return buf


def _insert_case(con, case):
    row = case.to_row()
    con.execute(
        "INSERT INTO cases (%s) VALUES (%s)"
        % (",".join(row.keys()), ",".join(["?"] * len(row))),
        list(row.values()),
    )
    con.commit()


def test_load_opinion_cluster_map_caches_snapshot_rows(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)
    opinions = _csv([
        {"id": "10", "cluster_id": "100", "plain_text": "ignored"},
        {"id": "20", "cluster_id": "200", "plain_text": "ignored"},
        {"id": "", "cluster_id": "300", "plain_text": "ignored"},
    ])

    out = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=opinions,
        snapshot_date="2026-03-31",
        refresh=True,
    )

    assert out == {"10": "100", "20": "200"}
    rows = con.execute(
        "SELECT opinion_id, cluster_id, snapshot_date FROM courtlistener_opinion_map "
        "ORDER BY opinion_id"
    ).fetchall()
    assert [tuple(row) for row in rows] == [
        ("10", "100", "2026-03-31"),
        ("20", "200", "2026-03-31"),
    ]


def test_load_opinion_cluster_map_reuses_cache_without_stream(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)
    con.execute(
        "INSERT INTO courtlistener_opinion_map VALUES (?,?,?)",
        ("10", "100", "2026-03-31"),
    )
    con.commit()

    out = parenthetical_loader.load_opinion_cluster_map(
        con,
        opinions_stream=None,
        snapshot_date="2026-03-31",
        refresh=False,
    )

    assert out == {"10": "100"}


def test_build_cluster_case_map_prefers_direct_cl_case_then_cap_citation(tmp_path):
    con = schema.connect(str(tmp_path / "c.db"))
    schema.create_schema(con)
    _insert_case(con, CaseRecord(case_uid="cl:100", source="cl", citation="15 Cal. 5th 1"))
    _insert_case(
        con,
        CaseRecord(
            case_uid="cap:aguilar",
            source="cap",
            citation="25 Cal. 4th 826",
            parallel_citations=["107 Cal. Rptr. 2d 841"],
        ),
    )
    citations = _csv([
        {"id": "1", "volume": "15", "reporter": "Cal. 5th", "page": "1", "type": "8", "cluster_id": "100"},
        {"id": "2", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
        {"id": "3", "volume": "999", "reporter": "N.Y.3d", "page": "1", "type": "8", "cluster_id": "400"},
    ])

    out = parenthetical_loader.build_cluster_case_map(con, citations_stream=citations)

    assert out["100"] == "cl:100"
    assert out["300"] == "cap:aguilar"
    assert "400" not in out


def test_iter_parenthetical_passages_filters_scores_and_attaches_to_local_cases():
    parentheticals = _csv([
        {
            "id": "900",
            "text": "describing Aguilar as assigning the summary judgment burden",
            "score": "0.95",
            "described_opinion_id": "10",
            "describing_opinion_id": "20",
            "group_id": "5",
        },
        {
            "id": "901",
            "text": "low quality description",
            "score": "0.10",
            "described_opinion_id": "10",
            "describing_opinion_id": "21",
            "group_id": "5",
        },
        {
            "id": "902",
            "text": "describes a missing local case",
            "score": "0.99",
            "described_opinion_id": "99",
            "describing_opinion_id": "21",
            "group_id": "5",
        },
    ])

    rows = list(
        parenthetical_loader.iter_parenthetical_passages(
            parentheticals_stream=parentheticals,
            opinion_cluster_map={"10": "300", "20": "200", "21": "201", "99": "999"},
            cluster_case_map={"300": "cap:aguilar"},
            min_score=0.5,
            max_per_case=25,
        )
    )

    assert len(rows) == 1
    passage = rows[0]
    assert passage.case_uid == "cap:aguilar"
    assert passage.passage_uid == "cap:aguilar#parenthetical:900"
    assert passage.passage_type == "parenthetical"
    assert passage.source == "courtlistener_parenthetical"
    assert passage.parenthetical_id == "900"
    assert passage.parenthetical_score == 0.95
    assert passage.described_opinion_id == "10"
    assert passage.describing_opinion_id == "20"
    assert passage.describing_cluster_id == "200"
    assert "summary judgment burden" in passage.text


def test_iter_parenthetical_passages_keeps_top_scored_rows_per_case():
    parentheticals = _csv([
        {"id": "1", "text": "score one", "score": "0.1", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "2", "text": "score nine", "score": "0.9", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
        {"id": "3", "text": "score five", "score": "0.5", "described_opinion_id": "10", "describing_opinion_id": "20", "group_id": ""},
    ])

    rows = list(
        parenthetical_loader.iter_parenthetical_passages(
            parentheticals_stream=parentheticals,
            opinion_cluster_map={"10": "300", "20": "200"},
            cluster_case_map={"300": "cap:aguilar"},
            min_score=0.0,
            max_per_case=2,
        )
    )

    assert [p.parenthetical_id for p in rows] == ["2", "3"]
    assert [p.ordinal for p in rows] == [1_000_000, 1_000_001]
```

- [ ] **Step 2: Run tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_parenthetical_loader.py -q
```

Expected: FAIL with `ImportError` for `parenthetical_loader`.

- [ ] **Step 3: Implement parenthetical_loader**

Create `icharlotte_core/legal_research/local_corpus/loaders/parenthetical_loader.py`:

```python
"""CourtListener parenthetical bulk loader.

Parentheticals describe one opinion from another opinion. This loader attaches
them to cases already present in the local corpus and emits provenance-tagged
PassageRecord rows; it never mutates case full_text.
"""
from __future__ import annotations

import csv
import heapq
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
        item = (score, parenthetical_id, passage)
        bucket = buckets[case_uid]
        if len(bucket) < max_per_case:
            heapq.heappush(bucket, item)
        else:
            heapq.heappushpop(bucket, item)

    for case_uid in sorted(buckets):
        selected = sorted(buckets[case_uid], key=lambda item: (-item[0], item[1]))
        for offset, (_score, _pid, passage) in enumerate(selected):
            passage.ordinal = _PARENTHETICAL_ORDINAL_BASE + offset
            yield passage
```

- [ ] **Step 4: Run loader tests to verify they pass**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_parenthetical_loader.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit Task 2**

Run:

```powershell
git add -- icharlotte_core/legal_research/local_corpus/loaders/parenthetical_loader.py tests/test_legal_research/test_local_corpus/test_parenthetical_loader.py
git diff --cached --check
git commit -m "feat(corpus): parse CourtListener parentheticals"
```

Expected: commit succeeds.

## Task 3: Corpus Append Function And CLI Wiring

**Files:**
- Modify: `icharlotte_core/legal_research/local_corpus/build.py`
- Create: `tests/test_legal_research/test_local_corpus/test_parenthetical_append.py`

- [ ] **Step 1: Write failing append tests**

Create `tests/test_legal_research/test_local_corpus/test_parenthetical_append.py`:

```python
import csv
import io

import numpy as np

from icharlotte_core.legal_research.local_corpus import build, schema
from icharlotte_core.legal_research.local_corpus.embedder import FakeEmbedder
from icharlotte_core.legal_research.local_corpus.indexer import CorpusIndexer
from icharlotte_core.legal_research.local_corpus.models import CaseRecord, PassageRecord


def _csv(rows):
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=list(rows[0].keys()))
    writer.writeheader()
    for row in rows:
        writer.writerow(row)
    buf.seek(0)
    return buf


def _seed_corpus(db, vec, emb):
    con = schema.connect(db)
    schema.create_schema(con)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:aguilar",
            source="cap",
            name="Aguilar v. Atlantic Richfield Co.",
            citation="25 Cal. 4th 826",
            decision_date="2001-06-14",
            year="2001",
            full_text="primary Aguilar opinion text",
        ),
        [
            PassageRecord(
                passage_uid="cap:aguilar#0",
                case_uid="cap:aguilar",
                ordinal=0,
                text="primary Aguilar opinion text",
            )
        ],
    )
    idx.finalize()
    con.close()


def test_append_parentheticals_adds_rows_and_metadata(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    _seed_corpus(db, vec, emb)
    con = schema.connect(db)
    schema.set_meta(con, cl_snapshot_date="2026-03-31")
    con.close()

    opinions = _csv([
        {"id": "10", "cluster_id": "300", "plain_text": ""},
        {"id": "20", "cluster_id": "200", "plain_text": ""},
    ])
    citations = _csv([
        {"id": "1", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
    ])
    parentheticals = _csv([
        {
            "id": "900",
            "text": "describing Aguilar as allocating the summary judgment burden",
            "score": "0.95",
            "described_opinion_id": "10",
            "describing_opinion_id": "20",
            "group_id": "",
        }
    ])

    summary = build.append_parentheticals_to_corpus(
        parentheticals_stream=parentheticals,
        opinions_stream=opinions,
        citations_stream=citations,
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        snapshot_date="2026-03-31",
        min_score=0.5,
        max_per_case=25,
        embed=False,
        refresh_opinion_map=True,
    )

    assert summary == {"added": 1}
    con = schema.connect(db)
    row = con.execute(
        "SELECT * FROM passages WHERE parenthetical_id='900'"
    ).fetchone()
    assert row["case_uid"] == "cap:aguilar"
    assert row["passage_type"] == "parenthetical"
    assert "summary judgment burden" in row["text"]
    meta = schema.get_meta(con)
    assert meta["parentheticals_snapshot_date"] == "2026-03-31"
    assert meta["parentheticals_count"] == "1"
    assert meta["parentheticals_min_score"] == "0.5"
    assert meta["parentheticals_max_per_case"] == "25"
    assert meta["cl_snapshot_date"] == "2026-03-31"
    arr = np.memmap(vec, dtype=np.float16, mode="r").reshape(-1, 16)
    assert arr.shape[0] == con.execute("SELECT COUNT(*) FROM passages").fetchone()[0]
    assert np.allclose(arr[int(row["vec_row"])], 0.0)


def test_append_parentheticals_is_idempotent(tmp_path):
    db = str(tmp_path / "corpus.db")
    vec = str(tmp_path / "vectors.f16")
    emb = FakeEmbedder(dim=16)
    _seed_corpus(db, vec, emb)
    opinions = _csv([{"id": "10", "cluster_id": "300", "plain_text": ""}])
    citations = _csv([
        {"id": "1", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
    ])
    parentheticals = [
        {
            "id": "900",
            "text": "describing Aguilar as allocating the summary judgment burden",
            "score": "0.95",
            "described_opinion_id": "10",
            "describing_opinion_id": "10",
            "group_id": "",
        }
    ]

    first = build.append_parentheticals_to_corpus(
        parentheticals_stream=_csv(parentheticals),
        opinions_stream=opinions,
        citations_stream=citations,
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        snapshot_date="2026-03-31",
        refresh_opinion_map=True,
    )
    second = build.append_parentheticals_to_corpus(
        parentheticals_stream=_csv(parentheticals),
        opinions_stream=None,
        citations_stream=_csv([
            {"id": "1", "volume": "25", "reporter": "Cal. 4th", "page": "826", "type": "8", "cluster_id": "300"},
        ]),
        db_path=db,
        vectors_path=vec,
        embedder=emb,
        snapshot_date="2026-03-31",
        refresh_opinion_map=False,
    )

    assert first == {"added": 1}
    assert second == {"added": 0}
    con = schema.connect(db)
    assert con.execute(
        "SELECT COUNT(*) FROM passages WHERE parenthetical_id='900'"
    ).fetchone()[0] == 1
```

- [ ] **Step 2: Run append tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_parenthetical_append.py -q
```

Expected: FAIL because `append_parentheticals_to_corpus` is not defined.

- [ ] **Step 3: Preserve existing CL case metadata**

In `icharlotte_core/legal_research/local_corpus/build.py`, update `_write_corpus_metadata` so calls that do not pass a CL case snapshot do not erase an existing value:

```python
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
```

- [ ] **Step 4: Add append function and metadata**

In `icharlotte_core/legal_research/local_corpus/build.py`, update the loader import:

```python
from icharlotte_core.legal_research.local_corpus.loaders import (
    cap_loader,
    cl_bulk_loader,
    parenthetical_loader,
)
```

Add this function after `append_cl_to_corpus`:

```python
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
```

- [ ] **Step 5: Add network wrapper and CLI options**

In `build.py`, add:

```python
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
```

Change the source choices line in `main()`:

```python
    ap.add_argument("--source", choices=["cap", "cl", "parentheticals", "all"], default="all")
```

Add CLI args after `--cl-embed`:

```python
    ap.add_argument("--parentheticals-min-score", type=float, default=0.5)
    ap.add_argument("--parentheticals-max-per-case", type=int, default=25)
    ap.add_argument("--embed-parentheticals", action="store_true")
    ap.add_argument("--refresh-opinion-map", action="store_true")
```

Add this execution block after the CL block:

```python
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
```

Do not run parenthetical ingest automatically for `--source all`; that would unexpectedly require the large opinion-id map stream on first setup.

- [ ] **Step 6: Run append tests**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_parenthetical_append.py -q
```

Expected: PASS.

- [ ] **Step 7: Commit Task 3**

Run:

```powershell
git add -- icharlotte_core/legal_research/local_corpus/build.py tests/test_legal_research/test_local_corpus/test_parenthetical_append.py
git diff --cached --check
git commit -m "feat(corpus): append CourtListener parentheticals"
```

Expected: commit succeeds.

## Task 4: Search Provenance And Verification Boundary

**Files:**
- Modify: `icharlotte_core/legal_research/models.py`
- Modify: `icharlotte_core/legal_research/local_corpus/corpus.py`
- Modify: `tests/test_legal_research/test_local_corpus/test_corpus.py`
- Modify: `tests/test_opposition/test_local_case_verifier.py`

- [ ] **Step 1: Write failing search and verifier-boundary tests**

Append these tests to `tests/test_legal_research/test_local_corpus/test_corpus.py`:

```python
def test_search_finds_case_through_parenthetical_passage(tmp_path):
    db = str(tmp_path / "c.db")
    vec = str(tmp_path / "v.f16")
    con = schema.connect(db)
    schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:aguilar",
            source="cap",
            name="Aguilar v. Atlantic Richfield Co.",
            citation="25 Cal. 4th 826",
            court="Cal.",
            decision_date="2001-06-14",
            year="2001",
            full_text="The opinion discusses asbestos and procedure.",
        ),
        [
            PassageRecord(
                passage_uid="cap:aguilar#0",
                case_uid="cap:aguilar",
                ordinal=0,
                text="The opinion discusses asbestos and procedure.",
            )
        ],
    )
    idx.finalize()
    con.close()

    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    idx.add_passages(
        [
            PassageRecord(
                passage_uid="cap:aguilar#parenthetical:900",
                case_uid="cap:aguilar",
                ordinal=1_000_000,
                text="describing Aguilar as allocating the summary judgment burden",
                passage_type="parenthetical",
                source="courtlistener_parenthetical",
                parenthetical_id="900",
                parenthetical_score=0.95,
            )
        ],
        embed=False,
    )
    idx.finalize()
    con.close()

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)
    results = corpus.search_opinions("summary judgment burden", semantic=False, max_results=5)

    assert results
    assert results[0].cluster_id == "cap:aguilar"
    assert "summary judgment burden" in results[0].snippet
    assert results[0].snippet_source == "parenthetical"
    assert results[0].snippet_parenthetical_id == "900"


def test_get_opinion_text_excludes_parenthetical_passages(tmp_path):
    db = str(tmp_path / "c.db")
    vec = str(tmp_path / "v.f16")
    con = schema.connect(db)
    schema.create_schema(con)
    emb = FakeEmbedder(dim=64)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb)
    idx.add(
        CaseRecord(
            case_uid="cap:aguilar",
            source="cap",
            citation="25 Cal. 4th 826",
            full_text="primary opinion text only",
        ),
        [PassageRecord(passage_uid="cap:aguilar#0", case_uid="cap:aguilar", ordinal=0, text="primary opinion text only")],
    )
    idx.finalize()
    con.close()
    con = schema.connect(db)
    idx = CorpusIndexer(con, vectors_path=vec, embedder=emb, resume=True)
    idx.add_passages(
        [
            PassageRecord(
                passage_uid="cap:aguilar#parenthetical:900",
                case_uid="cap:aguilar",
                ordinal=1_000_000,
                text="external parenthetical text",
                passage_type="parenthetical",
                parenthetical_id="900",
            )
        ],
        embed=False,
    )
    idx.finalize()
    con.close()

    corpus = LocalCaseCorpus(db_path=db, vectors_path=vec, embedder=emb)

    assert corpus.get_opinion_text("cap:aguilar") == "primary opinion text only"
    assert "external parenthetical" not in corpus.get_opinion_text("cap:aguilar")
```

Append this test to `tests/test_opposition/test_local_case_verifier.py`:

```python
def test_verifier_uses_full_text_not_parenthetical_context(monkeypatch):
    import icharlotte_core.opposition.local_case_verifier as mod
    monkeypatch.setattr(mod, "get_prompt", lambda *_a, **_k: "{authority_text}")
    captured = {}

    def llm(_sys, user):
        captured["user"] = user
        return json.dumps({
            "verdict": "NOT_SUPPORTED",
            "evidence": "",
            "note": "parenthetical text is not in the opinion",
        })

    found = {
        "case_uid": "cap:1",
        "full_text": "primary opinion text only",
        "name": "Aguilar v. Atlantic Richfield Co.",
        "url": "u",
        "court": "Cal.",
        "decision_date": "2001-06-14",
    }
    verifier = LocalCaseVerifier(corpus=_FakeCorpus(found), llm_callback=llm)
    citation = Citation(
        kind="case",
        raw_text="25 Cal. 4th 826",
        normalized="25 Cal. 4th 826",
        reporter_citation="25 Cal. 4th 826",
        proposition="summary judgment burden",
    )

    cv = verifier.verify(citation)

    assert "primary opinion text only" in captured["user"]
    assert "summary judgment burden from parenthetical" not in captured["user"]
    assert cv.verdict == "NOT_SUPPORTED"
```

- [ ] **Step 2: Run focused tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_corpus.py::test_search_finds_case_through_parenthetical_passage tests/test_legal_research/test_local_corpus/test_corpus.py::test_get_opinion_text_excludes_parenthetical_passages tests/test_opposition/test_local_case_verifier.py::test_verifier_uses_full_text_not_parenthetical_context -q
```

Expected: FAIL because `CaseResult` has no `snippet_source` field and search snippet selection does not prefer the matching parenthetical passage.

- [ ] **Step 3: Add snippet provenance fields to CaseResult**

In `icharlotte_core/legal_research/models.py`, add fields to `CaseResult`:

```python
    snippet_source: str = ""
    snippet_parenthetical_id: str = ""
```

Update `CaseResult.to_dict()` with:

```python
            "snippet_source": self.snippet_source,
            "snippet_parenthetical_id": self.snippet_parenthetical_id,
```

- [ ] **Step 4: Make LocalCaseCorpus rank and display parenthetical matches**

In `icharlotte_core/legal_research/local_corpus/corpus.py`, add a parenthetical ranking arm:

```python
    def _parenthetical_case_ranking(self, query: str, limit: int) -> list[str]:
        con = self._conn()
        rows = con.execute(
            "SELECT p.case_uid AS uid, bm25(passages_fts) AS score "
            "FROM passages_fts JOIN passages p ON p.vec_row = passages_fts.rowid - 1 "
            "WHERE passages_fts MATCH ? AND p.passage_type='parenthetical' "
            "ORDER BY score LIMIT ?",
            (_fts_query(query), limit),
        ).fetchall()
        seen, order = set(), []
        for row in rows:
            if row["uid"] not in seen:
                seen.add(row["uid"])
                order.append(row["uid"])
        return order
```

In `search_opinions`, change the rankings setup to:

```python
        parentheticals = self._parenthetical_case_ranking(query, _CANDIDATES)
        rankings = [metadata, bm25]
        if parentheticals:
            rankings.append(parentheticals)
```

Add best-snippet helper:

```python
    def _best_passage_for_query(self, case_uid: str, query: str):
        con = self._conn()
        try:
            row = con.execute(
                "SELECT p.text, p.passage_type, p.parenthetical_id "
                "FROM passages_fts JOIN passages p ON p.vec_row = passages_fts.rowid - 1 "
                "WHERE p.case_uid=? AND passages_fts MATCH ? "
                "ORDER BY bm25(passages_fts) LIMIT 1",
                (case_uid, _fts_query(query)),
            ).fetchone()
            if row:
                return row
        except sqlite3.Error:
            logger.debug("best passage lookup failed for %s", case_uid, exc_info=True)
        return con.execute(
            "SELECT text, passage_type, parenthetical_id FROM passages "
            "WHERE case_uid=? AND passage_type='opinion' ORDER BY ordinal LIMIT 1",
            (case_uid,),
        ).fetchone()
```

In `_case_result`, initialize snippet provenance before the `if c:` block:

```python
        snippet = ""
        display_name = ""
        snippet_source = ""
        snippet_parenthetical_id = ""
```

Then replace the first-passage snippet query inside the `if c:` block with:

```python
            p = self._best_passage_for_query(case_uid, query)
            snippet = (p["text"][:400] if p else (c["full_text"] or "")[:400])
            snippet_source = (p["passage_type"] if p else "") or ""
            snippet_parenthetical_id = (p["parenthetical_id"] if p else "") or ""
```

Update the returned `CaseResult`:

```python
            relevance_score=relevance_score,
            snippet_source=snippet_source,
            snippet_parenthetical_id=snippet_parenthetical_id,
```

- [ ] **Step 5: Run focused tests**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_corpus.py::test_search_finds_case_through_parenthetical_passage tests/test_legal_research/test_local_corpus/test_corpus.py::test_get_opinion_text_excludes_parenthetical_passages tests/test_opposition/test_local_case_verifier.py::test_verifier_uses_full_text_not_parenthetical_context -q
```

Expected: PASS.

- [ ] **Step 6: Run existing model tests**

Run:

```powershell
python -m pytest tests/test_legal_research/test_models.py -q
```

Expected: PASS.

- [ ] **Step 7: Commit Task 4**

Run:

```powershell
git add -- icharlotte_core/legal_research/models.py icharlotte_core/legal_research/local_corpus/corpus.py tests/test_legal_research/test_local_corpus/test_corpus.py tests/test_opposition/test_local_case_verifier.py
git diff --cached --check
git commit -m "feat(corpus): surface parenthetical matches in search"
```

Expected: commit succeeds.

## Task 5: README, Compile, And Focused Regression Run

**Files:**
- Modify: `icharlotte_core/legal_research/local_corpus/README.md`

- [ ] **Step 1: Update README with parenthetical ingest**

In `icharlotte_core/legal_research/local_corpus/README.md`, add this section after "CourtListener bulk stream wiring":

````markdown
### CourtListener parenthetical ingest

Parentheticals are secondary descriptions written by one opinion about another.
They are indexed as `passage_type='parenthetical'` rows under cases already in
the local CA corpus. They improve retrieval recall, but they are not appended to
`cases.full_text` and are not treated as quoted support from the described case.

```powershell
# First run for a snapshot normally needs the opinion-id map.
python -m icharlotte_core.legal_research.local_corpus.build `
  --source parentheticals `
  --cl-date 2026-03-31 `
  --refresh-opinion-map

# Later re-runs of the same snapshot can reuse the cached opinion map.
python -m icharlotte_core.legal_research.local_corpus.build `
  --source parentheticals `
  --cl-date 2026-03-31
```

Defaults:

- `--parentheticals-min-score 0.5`
- `--parentheticals-max-per-case 25`
- Parenthetical vectors are keyword-only by default. Use
  `--embed-parentheticals` to generate semantic vectors during ingest.

The parenthetical file is much smaller than the opinions file, but the first
map-building run may stream the matching opinions snapshot to cache
`opinion_id -> cluster_id`. The opinions file is streamed and not stored.
````

- [ ] **Step 2: Run parenthetical test suite**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus/test_parentheticals_schema.py tests/test_legal_research/test_local_corpus/test_parenthetical_loader.py tests/test_legal_research/test_local_corpus/test_parenthetical_append.py -q
```

Expected: PASS.

- [ ] **Step 3: Run local corpus regression tests**

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus -q
```

Expected: PASS. If this broad local-corpus folder run fails due an environment-only integration dependency, rerun all non-integration local-corpus files individually and record the exact failing dependency in the final summary.

- [ ] **Step 4: Run verifier boundary test**

Run:

```powershell
python -m pytest tests/test_opposition/test_local_case_verifier.py -q
```

Expected: PASS.

- [ ] **Step 5: Compile changed modules**

Run:

```powershell
python -m py_compile icharlotte_core/legal_research/models.py icharlotte_core/legal_research/local_corpus/schema.py icharlotte_core/legal_research/local_corpus/models.py icharlotte_core/legal_research/local_corpus/indexer.py icharlotte_core/legal_research/local_corpus/corpus.py icharlotte_core/legal_research/local_corpus/build.py icharlotte_core/legal_research/local_corpus/loaders/parenthetical_loader.py
```

Expected: command exits with code 0.

- [ ] **Step 6: Commit Task 5**

Run:

```powershell
git add -- icharlotte_core/legal_research/local_corpus/README.md
git diff --cached --check
git commit -m "docs(corpus): document parenthetical ingest"
```

Expected: commit succeeds.

## Final Verification

Run:

```powershell
python -m pytest tests/test_legal_research/test_local_corpus tests/test_opposition/test_local_case_verifier.py -q
python -m py_compile icharlotte_core/legal_research/models.py icharlotte_core/legal_research/local_corpus/schema.py icharlotte_core/legal_research/local_corpus/models.py icharlotte_core/legal_research/local_corpus/indexer.py icharlotte_core/legal_research/local_corpus/corpus.py icharlotte_core/legal_research/local_corpus/build.py icharlotte_core/legal_research/local_corpus/loaders/parenthetical_loader.py
git status --short
```

Expected:

- Local corpus tests pass.
- Local case verifier tests pass.
- `py_compile` exits with code 0.
- `git status --short` shows only unrelated pre-existing dirty files or no changes.

## Implementation Notes

- Use the existing `FakeEmbedder` for tests. Do not require a real model download.
- Do not run network parenthetical ingest during tests.
- Do not stage unrelated dirty files.
- Keep `cases.full_text` unchanged for every parenthetical operation.
- Keep parenthetical ingest explicit through `--source parentheticals`; do not run it as part of `--source all`.
