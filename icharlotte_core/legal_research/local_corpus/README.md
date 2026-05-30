# Local CA Case-Law Corpus

Offline retrieval + verification of California case law, built **entirely from
bulk data** so the Oppose-a-Motion pipeline no longer depends on the
rate-limited CourtListener live API (5 requests/minute on search + clusters).

Design spec: `docs/superpowers/specs/2026-05-29-ca-caselaw-local-corpus-design.md`
Implementation plan: `docs/superpowers/plans/2026-05-29-ca-caselaw-local-corpus.md`

## What it is

- **Two bulk sources, one corpus:** Harvard Caselaw Access Project (CAP) for the
  ≈1850–2020 backbone (~2.4 GB of per-volume ZIPs from `static.case.law`), plus a
  stream-filtered slice of CourtListener bulk for 2020→present (the gap CAP froze
  at). Both normalize into one SQLite DB.
- **Hybrid retrieval:** FTS5 BM25 + local ONNX semantic rerank (fastembed BGE-small),
  fused by Reciprocal Rank Fusion. Vectors live in a `vectors.f16` memmap sidecar.
- **Drop-in:** `LocalCaseCorpus` mirrors `CourtListenerClient` (`search_opinions`,
  `get_opinion_text`, `get_authority_signals`), so `argument_research` uses it
  unchanged. `LocalCaseVerifier` verifies cites against local text — no network.

## Storage

- Default dir: `CASELAW_DATA_DIR` (config.py) = `.gemini/caselaw/` (gitignored).
  **Relocatable** — set the `CASELAW_DATA_DIR` env var to a roomier drive.
- Files: `corpus.db` (cases + passages + FTS5) and `vectors.f16` (passage vectors).

## Building

```bash
# CAP backbone (downloads ~1,381 CA reporter volume ZIPs, ingests, embeds):
python -m icharlotte_core.legal_research.local_corpus.build --source cap

# Everything (CAP + CourtListener 2020+ gap):
python -m icharlotte_core.legal_research.local_corpus.build --source all

# Custom location:
python -m icharlotte_core.legal_research.local_corpus.build --source all --data-dir D:\caselaw
```

The build is **idempotent/restartable** — already-downloaded CAP volumes are skipped.

### CourtListener bulk stream wiring

CL bulk files are full-corpus, single-format CSVs on S3. We never store the ~50 GB
opinions file — we stream-decompress and keep only CA + post-cutoff rows. Wire the
three streams into `build_from_cl_streams`:

```python
import bz2, urllib.request, io
from icharlotte_core.legal_research.local_corpus import build
from icharlotte_core.legal_research.local_corpus.embedder import OnnxEmbedder

BASE = "https://com-courtlistener-storage.s3-us-west-2.amazonaws.com/bulk-data"
DATE = "2026-03-31"  # newest quarterly snapshot (browse the bucket for current date)

def _stream(name):
    url = f"{BASE}/{name}-{DATE}.csv.bz2"
    raw = urllib.request.urlopen(url, timeout=300)
    return io.TextIOWrapper(bz2.BZ2File(raw), encoding="utf-8")

build.build_from_cl_streams(
    courts_stream=_stream("courts"),
    clusters_stream=_stream("opinion-clusters"),
    opinions_stream=_stream("opinions"),   # ~50 GB — streamed, never stored
    db_path=..., vectors_path=..., embedder=OnnxEmbedder(),
)
```

> The opinions stream is a one-time ~52 GB network transfer. A mid-stream failure
> re-streams (true byte-offset resume over bz2 is infeasible); re-running is safe
> (idempotent — dedup by citation, CAP wins overlap).

## Quarterly refresh

CAP is frozen, so `--source cap` re-fetches nothing. To pick up new CA appellate
law, re-run `--source cl` after each CourtListener quarterly snapshot (last day of
Mar/Jun/Sep/Dec) by bumping `DATE` above.

## Swapping the embedder

`Embedder` is a Protocol (`embedder.py`). To use a stronger model, implement
`encode(texts) -> (n, dim) float32 unit-norm` and pass it to `LocalCaseCorpus(...,
embedder=YourEmbedder())` and the build functions. Only the vectors need
regenerating; schema/index/search are untouched.

## Known limits

- **No true good-law check.** `get_authority_signals` gives a soft staleness hint
  (inbound citation count + latest citing year), **not** Shepard's/KeyCite. The
  manual attorney good-law step remains.
- **Recency gap.** CAP ends ~2020; CL bulk is quarterly (≈2 months stale). The
  freshest 0–3 months of CA appellate law is absent until the next rebuild.
- **CA-only.** No federal, sister-state, or statutory authority in the pool.
- **Semantic quality < hosted.** BGE-small is good but below CourtListener's
  fine-tuned legal model; mitigated over time via the swappable embedder.
- **Recent slice = lowest-quality text.** CL-bulk 2020+ text is multi-source
  (OCR/HTML/syllabus-only) vs. CAP's canonical reporter text.
- **TODO (scale):** `LocalCaseCorpus.lookup_by_citation` does a normalized linear
  scan for v1 simplicity. For very large corpora, add a normalized-citation index
  column and query it directly.
