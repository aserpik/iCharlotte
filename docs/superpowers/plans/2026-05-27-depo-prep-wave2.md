# Depo Prep — Wave 2: Phase 1 Pipeline

> Sub-plan of [2026-05-27-depo-prep.md](2026-05-27-depo-prep.md). Wave 1 must be complete and green.

Goal of this wave: build the Phase 1 pipeline (per-source digest + topic clustering) and wire it into `Scripts/depo_prep.py --phase=analyze`. After this wave, you can run Phase 1 from the command line end-to-end.

---

### Task 4: Per-source digest module

**Files:**
- Create: `Scripts/depo_prep_lib/source_digest.py`
- Create: `tests/test_wizard/test_depo_prep_source_digest.py`

This module wraps a single `LLMCaller.call` per source with file-hash caching. Concurrency is handled by the orchestrator, NOT this module — keep it single-source so it's trivially testable.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_source_digest.py
import json
from pathlib import Path
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.source_digest import (
    digest_single_source,
    DigestResult,
)


def _mock_llm_returning(payload: dict):
    caller = MagicMock()
    caller.call = MagicMock(return_value=json.dumps(payload))
    return caller


def _good_payload(source_id="x.pdf"):
    return {
        "source_id": source_id,
        "source_kind": "medical_records",
        "deponent_statements": [],
        "factual_anchors": [{"fact": "x", "location": "p.1", "topic_tags": ["x"]}],
        "inconsistencies": [],
        "summary": "test",
    }


def test_digest_single_source_writes_json(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"some content")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("...extracted text...", encoding="utf-8")

    caller = _mock_llm_returning(_good_payload(source_id="test.pdf"))

    result = digest_single_source(
        source_path=src,
        extracted_text_path=digests_dir / "raw" / "test.txt",
        digests_dir=digests_dir,
        llm_caller=caller,
        deponent_name="Jane Doe",
        deponent_role="Plaintiff",
    )
    assert isinstance(result, DigestResult)
    assert result.from_cache is False
    out = digests_dir / "test.json"
    assert out.exists()
    data = json.loads(out.read_text(encoding="utf-8"))
    assert data["source_id"] == "test.pdf"


def test_digest_single_source_uses_cache_on_same_hash(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"same bytes")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("...", encoding="utf-8")

    caller = _mock_llm_returning(_good_payload(source_id="test.pdf"))

    # First call — LLM is invoked.
    r1 = digest_single_source(
        source_path=src,
        extracted_text_path=digests_dir / "raw" / "test.txt",
        digests_dir=digests_dir, llm_caller=caller,
        deponent_name="Jane", deponent_role="P",
    )
    assert r1.from_cache is False
    assert caller.call.call_count == 1

    # Second call with same file — should hit cache, no LLM call.
    r2 = digest_single_source(
        source_path=src,
        extracted_text_path=digests_dir / "raw" / "test.txt",
        digests_dir=digests_dir, llm_caller=caller,
        deponent_name="Jane", deponent_role="P",
    )
    assert r2.from_cache is True
    assert caller.call.call_count == 1  # unchanged


def test_digest_single_source_invalidates_when_file_changes(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"original")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("v1", encoding="utf-8")

    caller = _mock_llm_returning(_good_payload(source_id="test.pdf"))

    digest_single_source(source_path=src, extracted_text_path=digests_dir/"raw"/"test.txt",
                         digests_dir=digests_dir, llm_caller=caller,
                         deponent_name="Jane", deponent_role="P")

    # Change file contents → hash changes → cache invalidates.
    src.write_bytes(b"changed")
    (digests_dir / "raw" / "test.txt").write_text("v2", encoding="utf-8")

    r2 = digest_single_source(source_path=src, extracted_text_path=digests_dir/"raw"/"test.txt",
                              digests_dir=digests_dir, llm_caller=caller,
                              deponent_name="Jane", deponent_role="P")
    assert r2.from_cache is False
    assert caller.call.call_count == 2


def test_digest_single_source_strips_json_fences(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"x")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("...", encoding="utf-8")

    fenced = "```json\n" + json.dumps(_good_payload()) + "\n```"
    caller = MagicMock()
    caller.call = MagicMock(return_value=fenced)

    result = digest_single_source(source_path=src, extracted_text_path=digests_dir/"raw"/"test.txt",
                                  digests_dir=digests_dir, llm_caller=caller,
                                  deponent_name="J", deponent_role="P")
    assert result.from_cache is False
    # The JSON inside the fence is parsed; the digest file is written.
    assert (digests_dir / "x.pdf.json").exists() or any(p.suffix == ".json" for p in digests_dir.glob("*.json"))


def test_digest_single_source_raises_on_invalid_json(tmp_path):
    src = tmp_path / "test.pdf"
    src.write_bytes(b"x")
    digests_dir = tmp_path / "digests"
    digests_dir.mkdir()
    (digests_dir / "raw").mkdir()
    (digests_dir / "raw" / "test.txt").write_text("...", encoding="utf-8")

    caller = MagicMock()
    caller.call = MagicMock(return_value="not json at all")

    with pytest.raises(ValueError, match="JSON"):
        digest_single_source(source_path=src, extracted_text_path=digests_dir/"raw"/"test.txt",
                             digests_dir=digests_dir, llm_caller=caller,
                             deponent_name="J", deponent_role="P")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_source_digest.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/source_digest.py
"""Stage 1.2 — per-source structured digest with file-hash caching."""
from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Union

from .prompts import build_per_source_digest_prompt
from .schemas import validate_source_digest_dict
from .session_io import file_sha256, write_json, read_json


_FENCE_RE = re.compile(r"^```(?:json)?\s*\n(.*)\n```\s*$", re.DOTALL)


@dataclass
class DigestResult:
    digest_path: Path
    digest_data: dict
    from_cache: bool


def _strip_fences(s: str) -> str:
    s = (s or "").strip()
    m = _FENCE_RE.match(s)
    return m.group(1).strip() if m else s


def _parse_llm_json(raw: str) -> dict:
    text = _strip_fences(raw)
    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"LLM returned invalid JSON: {e}") from e


def _cache_hash_path(digest_path: Path) -> Path:
    return digest_path.with_suffix(digest_path.suffix + ".sha256")


def digest_single_source(
    *,
    source_path: Union[str, Path],
    extracted_text_path: Union[str, Path],
    digests_dir: Union[str, Path],
    llm_caller,
    deponent_name: str,
    deponent_role: str,
) -> DigestResult:
    """Produce (or load from cache) the structured digest for one source file.

    Cache key = sha256 of source_path. Side-by-side .sha256 file holds the hash
    of the source the digest was produced from; if the source's current hash
    matches, we reuse the digest.
    """
    source_path = Path(source_path)
    digests_dir = Path(digests_dir)
    digests_dir.mkdir(parents=True, exist_ok=True)

    digest_path = digests_dir / f"{source_path.name}.json"
    hash_path = _cache_hash_path(digest_path)

    current_hash = file_sha256(source_path)

    if digest_path.exists() and hash_path.exists():
        cached_hash = hash_path.read_text(encoding="utf-8").strip()
        if cached_hash == current_hash:
            return DigestResult(
                digest_path=digest_path,
                digest_data=read_json(digest_path),
                from_cache=True,
            )

    extracted_text = Path(extracted_text_path).read_text(encoding="utf-8", errors="replace")

    prompt, text_payload = build_per_source_digest_prompt(
        deponent_name=deponent_name,
        deponent_role=deponent_role,
        source_text=extracted_text,
        source_filename=source_path.name,
    )

    raw = llm_caller.call(
        prompt=prompt,
        text=text_payload,
        task_type="extraction",
        agent_id="DepoPrep",
        pass_name="source_digest",
    )

    data = _parse_llm_json(raw)
    # Force source_id to match the actual filename in case the LLM hallucinated.
    data["source_id"] = source_path.name
    validate_source_digest_dict(data)

    write_json(digest_path, data)
    hash_path.write_text(current_hash, encoding="utf-8")

    return DigestResult(digest_path=digest_path, digest_data=data, from_cache=False)
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_source_digest.py -v`
Expected: 5 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/source_digest.py tests/test_wizard/test_depo_prep_source_digest.py
git commit -m "feat(depo_prep): source_digest — per-source extraction with hash caching"
```

---

### Task 5: Topic clustering module

**Files:**
- Create: `Scripts/depo_prep_lib/topics.py`
- Create: `tests/test_wizard/test_depo_prep_topics.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_topics.py
import json
from unittest.mock import MagicMock

import pytest

from Scripts.depo_prep_lib.topics import cluster_topics, TopicsResult


def _payload(topics):
    return json.dumps({"topics": topics})


def _topic(i, **kw):
    base = {
        "id": f"t{i:02d}",
        "title": f"Topic {i}",
        "strategic_note": "note",
        "relevant_digest_refs": [],
        "default_checked": True,
    }
    base.update(kw)
    return base


def test_cluster_topics_returns_topics_list():
    caller = MagicMock()
    caller.call = MagicMock(return_value=_payload([_topic(i) for i in range(1, 11)]))

    result = cluster_topics(
        digests=[{"source_id": "x.pdf", "summary": "s"}],
        llm_caller=caller,
        deponent_name="Jane", deponent_role="P",
        style="discovery", free_text_notes="",
    )
    assert isinstance(result, TopicsResult)
    assert len(result.topics) == 10
    assert result.warning is None


def test_cluster_topics_warns_below_3():
    caller = MagicMock()
    caller.call = MagicMock(return_value=_payload([_topic(1), _topic(2)]))

    result = cluster_topics(
        digests=[{"source_id": "x.pdf", "summary": "s"}],
        llm_caller=caller,
        deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
    )
    assert result.warning is not None
    assert "thin" in result.warning.lower() or "few" in result.warning.lower()


def test_cluster_topics_truncates_above_20():
    caller = MagicMock()
    caller.call = MagicMock(return_value=_payload([_topic(i) for i in range(1, 30)]))

    result = cluster_topics(
        digests=[{"source_id": "x.pdf", "summary": "s"}],
        llm_caller=caller,
        deponent_name="J", deponent_role="P",
        style="discovery", free_text_notes="",
    )
    assert len(result.topics) == 20
    assert result.warning is not None
    assert "truncat" in result.warning.lower()


def test_cluster_topics_raises_on_bad_payload():
    caller = MagicMock()
    caller.call = MagicMock(return_value="not json")

    with pytest.raises(ValueError):
        cluster_topics(digests=[], llm_caller=caller,
                       deponent_name="J", deponent_role="P",
                       style="discovery", free_text_notes="")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_topics.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/topics.py
"""Stage 1.3 — cluster per-source digests into 8–15 topics."""
from __future__ import annotations

import json
from dataclasses import dataclass
from typing import List, Optional

from .prompts import build_topic_clustering_prompt
from .schemas import validate_topics_dict, Topic
from .source_digest import _parse_llm_json  # reuse fence-stripping JSON parser


_MAX_TOPICS = 20
_MIN_TOPICS_WARN = 3


@dataclass
class TopicsResult:
    topics: List[dict]   # plain dicts (matching Topic schema), ready to serialize
    warning: Optional[str]


def _digests_to_summary_text(digests: List[dict]) -> str:
    """Render the digest list as a single text payload for the clustering prompt."""
    blocks = []
    for d in digests:
        blocks.append(f"=== DIGEST: {d.get('source_id', 'unknown')} ===\n"
                      + json.dumps(d, indent=2, ensure_ascii=False))
    return "\n\n".join(blocks)


def cluster_topics(
    *,
    digests: List[dict],
    llm_caller,
    deponent_name: str,
    deponent_role: str,
    style: str,
    free_text_notes: str,
) -> TopicsResult:
    prompt, text_payload = build_topic_clustering_prompt(
        deponent_name=deponent_name, deponent_role=deponent_role,
        style=style, free_text_notes=free_text_notes,
        digests_summary_text=_digests_to_summary_text(digests),
    )

    raw = llm_caller.call(
        prompt=prompt, text=text_payload,
        task_type="general", agent_id="DepoPrep", pass_name="topic_clustering",
    )

    data = _parse_llm_json(raw)
    validate_topics_dict(data)

    topics = list(data["topics"])
    warning = None
    if len(topics) > _MAX_TOPICS:
        topics = topics[:_MAX_TOPICS]
        warning = f"LLM produced more than {_MAX_TOPICS} topics; truncated."
    if len(topics) < _MIN_TOPICS_WARN:
        warning = ("Source material appears thin — only "
                   f"{len(topics)} topic(s) emerged. Consider adding more sources "
                   "or detail in your strategy notes.")

    # Normalize each topic through the Topic dataclass to fill defaults.
    normalized = [Topic.from_dict(t).to_dict() for t in topics]

    return TopicsResult(topics=normalized, warning=warning)
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_topics.py -v`
Expected: 4 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/topics.py tests/test_wizard/test_depo_prep_topics.py
git commit -m "feat(depo_prep): topics — cluster digests into 8-15 topics with bounds checks"
```

---

### Task 6: Phase 1 orchestrator in `depo_prep.py`

**Files:**
- Create: `Scripts/depo_prep.py`
- Create: `tests/test_wizard/test_depo_prep_phase1_orchestrator.py`

`depo_prep.py` is a thin CLI: parses args, extracts source text, fans out per-source digests with `ThreadPoolExecutor(max_workers=4)`, calls `cluster_topics`, writes `session.json` + `topics.json`, prints `AWAITING_INPUT:<session_path>`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_phase1_orchestrator.py
"""Phase 1 orchestrator — exercised at the function level (no subprocess)."""
import json
from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest


@pytest.fixture
def fake_case_root(tmp_path):
    root = tmp_path / "Smith v. Jones"
    (root / "RECORDS").mkdir(parents=True)
    (root / "PLEADINGS").mkdir()
    return root


def _make_text_source(path, content):
    path.write_text(content, encoding="utf-8")
    return path


def _mock_llm():
    """LLM mock that returns digest payload for extraction, topics payload for general."""
    caller = MagicMock()
    def call(prompt, text, task_type="general", agent_id=None, pass_name=None, **kwargs):
        if pass_name == "source_digest":
            return json.dumps({
                "source_id": "echo",
                "source_kind": "other",
                "deponent_statements": [],
                "factual_anchors": [{"fact": "x", "location": "p.1", "topic_tags": ["t"]}],
                "inconsistencies": [],
                "summary": "ok",
            })
        if pass_name == "topic_clustering":
            return json.dumps({"topics": [
                {"id": f"t{i:02d}", "title": f"Topic {i}", "strategic_note": "n",
                 "relevant_digest_refs": [], "default_checked": True}
                for i in range(1, 11)
            ]})
        return ""
    caller.call.side_effect = call
    return caller


def test_phase1_writes_session_and_topics(fake_case_root, tmp_path, capsys):
    from Scripts.depo_prep_lib import phase1

    src1 = _make_text_source(fake_case_root / "RECORDS" / "med.txt", "medical records text")
    src2 = _make_text_source(fake_case_root / "PLEADINGS" / "complaint.txt", "complaint text")

    config = {
        "deponent_name": "Jane Doe",
        "deponent_role": "Plaintiff",
        "deponent_sources": [str(src1)],
        "context_sources": [str(src2)],
        "style": "discovery",
        "free_text_notes": "Focus on causation.",
        "per_topic_flags": {
            "strategic_note": True, "source_facts": True,
            "impeachment_hook": False, "objection_alts": False,
        },
        "case_root": str(fake_case_root),
    }

    session_path = phase1.run_phase1(
        config=config,
        llm_caller=_mock_llm(),
        progress=lambda n, msg=None: None,
    )

    assert Path(session_path).exists()
    session_dir = Path(session_path).parent
    assert (session_dir / "topics.json").exists()
    topics = json.loads((session_dir / "topics.json").read_text(encoding="utf-8"))
    assert len(topics["topics"]) == 10

    digests_dir = session_dir / "digests"
    assert (digests_dir / "med.txt.json").exists()
    assert (digests_dir / "complaint.txt.json").exists()


def test_phase1_passes_extraction_task_type_to_llm(fake_case_root):
    from Scripts.depo_prep_lib import phase1

    src = _make_text_source(fake_case_root / "RECORDS" / "med.txt", "text")
    config = {
        "deponent_name": "J", "deponent_role": "P",
        "deponent_sources": [str(src)], "context_sources": [],
        "style": "discovery", "free_text_notes": "",
        "per_topic_flags": {"strategic_note": False, "source_facts": False,
                             "impeachment_hook": False, "objection_alts": False},
        "case_root": str(fake_case_root),
    }

    caller = _mock_llm()
    phase1.run_phase1(config=config, llm_caller=caller, progress=lambda *a, **kw: None)

    # At least one call must have been for source_digest with task_type="extraction".
    digest_calls = [c for c in caller.call.call_args_list
                    if c.kwargs.get("pass_name") == "source_digest"]
    assert digest_calls
    assert all(c.kwargs.get("task_type") == "extraction" for c in digest_calls)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_phase1_orchestrator.py -v`
Expected: ImportError on `Scripts.depo_prep_lib.phase1`.

- [ ] **Step 3: Write minimal implementation**

First, create the in-library orchestrator (testable, no subprocess):

```python
# Scripts/depo_prep_lib/phase1.py
"""Phase 1 orchestrator — runs in-process; depo_prep.py calls run_phase1()."""
from __future__ import annotations

import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
from typing import Callable, List

from .session_io import compute_session_paths, write_json
from .source_digest import digest_single_source
from .topics import cluster_topics


_TEXT_EXTS = {".txt", ".md"}


def _extract_text_to_raw(source_path: Path, raw_dir: Path, logger=None) -> Path:
    """Extract text from source_path into raw_dir/<source>.txt and return that path.

    Uses icharlotte_core.document_processor for PDFs/DOCX; plain text is copied.
    """
    raw_dir.mkdir(parents=True, exist_ok=True)
    out = raw_dir / f"{source_path.name}.txt"
    ext = source_path.suffix.lower()
    if ext in _TEXT_EXTS:
        out.write_text(source_path.read_text(encoding="utf-8", errors="replace"),
                       encoding="utf-8")
        return out
    if ext == ".pdf":
        from icharlotte_core.document_processor import DocumentProcessor, OCRConfig
        processor = DocumentProcessor(ocr_config=OCRConfig(adaptive=True), logger=logger)
        result = processor.extract_with_dynamic_ocr(str(source_path))
        text = result.text if result.success else ""
        out.write_text(text, encoding="utf-8")
        return out
    if ext == ".docx":
        from icharlotte_core.document_processor import extract_docx_text
        text = extract_docx_text(str(source_path))
        out.write_text(text or "", encoding="utf-8")
        return out
    # Unsupported types: write empty text and let the LLM see an empty source.
    out.write_text("", encoding="utf-8")
    return out


def run_phase1(*, config: dict, llm_caller, progress: Callable[[int, str], None]) -> str:
    """Execute Phase 1: ingest → per-source digest → topic clustering → persist.

    Returns the absolute path to session.json.
    """
    deponent_name = config["deponent_name"]
    deponent_role = config.get("deponent_role", "")
    style = config.get("style", "discovery")
    free_text = config.get("free_text_notes", "")
    case_root = config["case_root"]

    all_sources = list(config.get("deponent_sources", [])) + list(config.get("context_sources", []))
    if not all_sources:
        raise ValueError("Phase 1 requires at least one source file")

    paths = compute_session_paths(
        case_root=case_root, deponent_name=deponent_name,
        when_iso=datetime.now().isoformat(timespec="minutes"),
    )
    paths.session_dir.mkdir(parents=True, exist_ok=True)
    paths.digests_dir.mkdir(parents=True, exist_ok=True)
    paths.raw_dir.mkdir(parents=True, exist_ok=True)

    progress(5, "Extracting source text…")
    extracted_map: dict = {}
    for i, src_str in enumerate(all_sources, 1):
        src = Path(src_str)
        extracted_map[src] = _extract_text_to_raw(src, paths.raw_dir)
        progress(5 + int(15 * i / len(all_sources)), f"Extracted {src.name}")

    progress(25, "Building per-source digests…")
    digests: List[dict] = []

    def _one(src_path: Path):
        return digest_single_source(
            source_path=src_path,
            extracted_text_path=extracted_map[src_path],
            digests_dir=paths.digests_dir,
            llm_caller=llm_caller,
            deponent_name=deponent_name,
            deponent_role=deponent_role,
        )

    with ThreadPoolExecutor(max_workers=4) as pool:
        futures = {pool.submit(_one, Path(s)): s for s in all_sources}
        done = 0
        for fut in as_completed(futures):
            done += 1
            try:
                result = fut.result()
                digests.append(result.digest_data)
            except Exception as e:
                digests.append({
                    "source_id": Path(futures[fut]).name,
                    "source_kind": "other",
                    "deponent_statements": [], "factual_anchors": [],
                    "inconsistencies": [],
                    "summary": f"DIGEST FAILED: {e}",
                })
            progress(25 + int(50 * done / len(all_sources)),
                     f"Digested {done}/{len(all_sources)}")

    progress(80, "Clustering topics…")
    topics_result = cluster_topics(
        digests=digests, llm_caller=llm_caller,
        deponent_name=deponent_name, deponent_role=deponent_role,
        style=style, free_text_notes=free_text,
    )

    write_json(paths.topics_json, {
        "topics": topics_result.topics,
        "warning": topics_result.warning,
    })

    write_json(paths.session_json, {
        "version": 1,
        "phase": "awaiting_input",
        "deponent_name": deponent_name,
        "deponent_role": deponent_role,
        "style": style,
        "free_text_notes": free_text,
        "per_topic_flags": config.get("per_topic_flags", {}),
        "case_root": case_root,
        "deponent_sources": list(config.get("deponent_sources", [])),
        "context_sources": list(config.get("context_sources", [])),
        "digests_index": [d.get("source_id") for d in digests],
        "topics_warning": topics_result.warning,
    })

    progress(95, f"Discovered {len(topics_result.topics)} topics")
    return str(paths.session_json)
```

Now the CLI entry point. **Critical**: `sys.path.insert(0, project_root)` MUST come before any `icharlotte_core` import.

```python
# Scripts/depo_prep.py
"""Depo Prep CLI agent.

Two phases:
  --phase=analyze --config=<path>    Reads config.json, runs Phase 1, emits
                                     AWAITING_INPUT:<session.json path>.
  --phase=generate --session=<path>  Reads session.json + topics.json (mutated
                                     by the wizard during topic editing), runs
                                     Phase 2, writes outline.docx + outline.md.

Prints PROGRESS:<int>:<msg> lines for the wizard's status page.
"""
from __future__ import annotations

import argparse
import os
import sys
import traceback

# MUST come BEFORE any icharlotte_core import.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _progress(n: int, msg: str = "") -> None:
    if msg:
        print(f"PROGRESS:{n}:{msg}", flush=True)
    else:
        print(f"PROGRESS:{n}", flush=True)


def _make_llm_caller():
    from icharlotte_core.llm_config import LLMCaller
    return LLMCaller()


def _cmd_analyze(config_path: str) -> int:
    from Scripts.depo_prep_lib import phase1
    from Scripts.depo_prep_lib.session_io import read_json

    config = read_json(config_path)
    try:
        session_json_path = phase1.run_phase1(
            config=config, llm_caller=_make_llm_caller(), progress=_progress,
        )
    except Exception as e:
        print(f"ERROR: Phase 1 failed: {e}", flush=True)
        traceback.print_exc()
        return 1

    print(f"AWAITING_INPUT:{session_json_path}", flush=True)
    return 0


def _cmd_generate(session_path: str) -> int:
    # Implemented in Wave 3 — for now, a clear placeholder.
    print("ERROR: --phase=generate not yet implemented (Wave 3)", flush=True)
    return 2


def main():
    parser = argparse.ArgumentParser(description="Depo Prep agent")
    parser.add_argument("--phase", required=True, choices=("analyze", "generate"))
    parser.add_argument("--config", default=None,
                        help="Path to config.json (required for --phase=analyze)")
    parser.add_argument("--session", default=None,
                        help="Path to session.json (required for --phase=generate)")
    args = parser.parse_args()

    if args.phase == "analyze":
        if not args.config:
            print("ERROR: --config is required for --phase=analyze", flush=True)
            return 2
        return _cmd_analyze(args.config)
    if args.phase == "generate":
        if not args.session:
            print("ERROR: --session is required for --phase=generate", flush=True)
            return 2
        return _cmd_generate(args.session)
    return 2


if __name__ == "__main__":
    sys.exit(main())
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_phase1_orchestrator.py -v`
Expected: 2 passed.

- [ ] **Step 5: Manual smoke test from CLI**

Create a temp config file and a tiny text source:

```powershell
$tmp = New-TemporaryFile
$case = New-Item -ItemType Directory -Path (Join-Path $env:TEMP "depo_prep_smoke_$(Get-Random)")
New-Item -ItemType Directory -Path (Join-Path $case "RECORDS") | Out-Null
Set-Content -Path (Join-Path $case "RECORDS\med.txt") -Value "Patient reported prior back pain in 2019."
$config = @{
  deponent_name="Jane Doe"; deponent_role="Plaintiff";
  deponent_sources=@((Join-Path $case "RECORDS\med.txt"));
  context_sources=@();
  style="discovery"; free_text_notes="Focus on causation.";
  per_topic_flags=@{strategic_note=$true; source_facts=$true; impeachment_hook=$false; objection_alts=$false};
  case_root=$case.FullName
} | ConvertTo-Json -Depth 5
Set-Content -Path $tmp -Value $config
python Scripts\depo_prep.py --phase=analyze --config=$tmp
```

Expected: PROGRESS lines, then `AWAITING_INPUT:<full path to session.json>`. The session folder under `<case>\NOTES\AI Output\Depo Prep - Jane Doe - ...\` should contain `session.json`, `topics.json`, and `digests/`. This will actually call the LLM, so requires API keys in `.env`.

- [ ] **Step 6: Commit**

```bash
git add Scripts/depo_prep.py Scripts/depo_prep_lib/phase1.py tests/test_wizard/test_depo_prep_phase1_orchestrator.py
git commit -m "feat(depo_prep): Phase 1 orchestrator — extract → digest → cluster → AWAITING_INPUT"
```

---

### Wave 2 self-check

- [ ] All Wave 2 tests pass: `python -m pytest tests/test_wizard/test_depo_prep_source_digest.py tests/test_wizard/test_depo_prep_topics.py tests/test_wizard/test_depo_prep_phase1_orchestrator.py -v` → green
- [ ] Manual CLI smoke (Task 6 Step 5) produced session.json + topics.json with ≥3 topics
- [ ] `Scripts/depo_prep.py --phase=generate` returns the placeholder error (correct for now)
- [ ] No regressions: `python -m pytest tests/test_wizard/ -v` still green
- [ ] Three new commits on the branch

Proceed to Wave 3.
