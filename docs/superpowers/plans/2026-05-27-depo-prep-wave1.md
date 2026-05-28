# Depo Prep — Wave 1: Foundation

> Sub-plan of [2026-05-27-depo-prep.md](2026-05-27-depo-prep.md). Execute Wave 1 fully before starting Wave 2.

Goal of this wave: stand up the data layer (schemas, session I/O) and the prompt library so Phase 1 and Phase 2 modules in Waves 2–3 have stable contracts to import.

---

### Task 1: Package scaffold + schemas

**Files:**
- Create: `Scripts/depo_prep_lib/__init__.py` (empty)
- Create: `Scripts/depo_prep_lib/schemas.py`
- Create: `tests/test_wizard/test_depo_prep_schemas.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_schemas.py
"""Schema validation tests for depo_prep_lib.schemas."""
import pytest

from Scripts.depo_prep_lib.schemas import (
    DeponentStatement,
    FactualAnchor,
    Inconsistency,
    SourceDigest,
    Topic,
    Question,
    TopicQuestions,
    validate_source_digest_dict,
    validate_topics_dict,
)


def test_source_digest_roundtrip():
    digest = SourceDigest(
        source_id="med_records.pdf",
        source_kind="medical_records",
        deponent_statements=[
            DeponentStatement(text="I had no prior pain.", location="p.47:18-22", context="Direct exam")
        ],
        factual_anchors=[
            FactualAnchor(fact="MRI 2024-09-12 showed 4mm protrusion", location="p.12 Bates DEF-00154",
                          topic_tags=["injury", "imaging"])
        ],
        inconsistencies=[
            Inconsistency(claim_a="RFA #7: pain immediate", claim_a_source="this file, RFA #7",
                          claim_b="ER triage: no acute pain", claim_b_source="med_records p.3",
                          topic_tags=["credibility"])
        ],
        summary="Med records show chronic LBP prior to accident.",
    )
    d = digest.to_dict()
    assert d["source_id"] == "med_records.pdf"
    assert d["deponent_statements"][0]["location"] == "p.47:18-22"
    # Round-trip
    digest2 = SourceDigest.from_dict(d)
    assert digest2.summary == digest.summary


def test_validate_source_digest_rejects_missing_field():
    bad = {"source_id": "x", "source_kind": "other"}  # missing required lists
    with pytest.raises(ValueError, match="missing"):
        validate_source_digest_dict(bad)


def test_topic_default_checked_true():
    t = Topic(id="t01", title="Pain timeline", strategic_note="Establish baseline",
              relevant_digest_refs=["a.pdf#factual_anchors[0]"])
    assert t.default_checked is True
    assert t.lawyer_added is False


def test_question_optional_fields_default_none():
    q = Question(n=1, text="When did pain begin?")
    assert q.purpose is None
    assert q.source_facts is None
    assert q.impeachment_hook is None
    assert q.objection_alts is None


def test_topic_questions_roundtrip():
    tq = TopicQuestions(
        topic_id="t01",
        questions=[Question(n=1, text="Q1", purpose="P1")],
    )
    d = tq.to_dict()
    assert d["questions"][0]["purpose"] == "P1"
    tq2 = TopicQuestions.from_dict(d)
    assert tq2.questions[0].text == "Q1"


def test_validate_topics_dict_accepts_minimal_topic():
    payload = {"topics": [{"id": "t01", "title": "X", "strategic_note": "Y",
                            "relevant_digest_refs": []}]}
    validate_topics_dict(payload)  # should not raise


def test_validate_topics_dict_rejects_non_list_topics():
    with pytest.raises(ValueError):
        validate_topics_dict({"topics": "nope"})
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_schemas.py -v`
Expected: ImportError on `Scripts.depo_prep_lib.schemas`.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/__init__.py
# (empty package marker)
```

```python
# Scripts/depo_prep_lib/schemas.py
"""Dataclasses + JSON validators for Depo Prep session artifacts."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import List, Optional


@dataclass
class DeponentStatement:
    text: str
    location: str
    context: str = ""

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "DeponentStatement":
        return cls(text=d["text"], location=d["location"], context=d.get("context", ""))


@dataclass
class FactualAnchor:
    fact: str
    location: str
    topic_tags: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "FactualAnchor":
        return cls(fact=d["fact"], location=d["location"], topic_tags=list(d.get("topic_tags", [])))


@dataclass
class Inconsistency:
    claim_a: str
    claim_a_source: str
    claim_b: str
    claim_b_source: str
    topic_tags: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Inconsistency":
        return cls(
            claim_a=d["claim_a"], claim_a_source=d["claim_a_source"],
            claim_b=d["claim_b"], claim_b_source=d["claim_b_source"],
            topic_tags=list(d.get("topic_tags", [])),
        )


@dataclass
class SourceDigest:
    source_id: str
    source_kind: str
    deponent_statements: List[DeponentStatement] = field(default_factory=list)
    factual_anchors: List[FactualAnchor] = field(default_factory=list)
    inconsistencies: List[Inconsistency] = field(default_factory=list)
    summary: str = ""

    def to_dict(self) -> dict:
        return {
            "source_id": self.source_id,
            "source_kind": self.source_kind,
            "deponent_statements": [s.to_dict() for s in self.deponent_statements],
            "factual_anchors": [a.to_dict() for a in self.factual_anchors],
            "inconsistencies": [i.to_dict() for i in self.inconsistencies],
            "summary": self.summary,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "SourceDigest":
        return cls(
            source_id=d["source_id"],
            source_kind=d["source_kind"],
            deponent_statements=[DeponentStatement.from_dict(s) for s in d.get("deponent_statements", [])],
            factual_anchors=[FactualAnchor.from_dict(a) for a in d.get("factual_anchors", [])],
            inconsistencies=[Inconsistency.from_dict(i) for i in d.get("inconsistencies", [])],
            summary=d.get("summary", ""),
        )


@dataclass
class Topic:
    id: str
    title: str
    strategic_note: str
    relevant_digest_refs: List[str] = field(default_factory=list)
    default_checked: bool = True
    lawyer_added: bool = False

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: dict) -> "Topic":
        return cls(
            id=d["id"], title=d["title"], strategic_note=d.get("strategic_note", ""),
            relevant_digest_refs=list(d.get("relevant_digest_refs", [])),
            default_checked=bool(d.get("default_checked", True)),
            lawyer_added=bool(d.get("lawyer_added", False)),
        )


@dataclass
class Question:
    n: int
    text: str
    purpose: Optional[str] = None
    source_facts: Optional[List[str]] = None
    impeachment_hook: Optional[str] = None
    objection_alts: Optional[List[str]] = None

    def to_dict(self) -> dict:
        d = {"n": self.n, "text": self.text}
        if self.purpose is not None: d["purpose"] = self.purpose
        if self.source_facts is not None: d["source_facts"] = list(self.source_facts)
        if self.impeachment_hook is not None: d["impeachment_hook"] = self.impeachment_hook
        if self.objection_alts is not None: d["objection_alts"] = list(self.objection_alts)
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "Question":
        return cls(
            n=int(d["n"]), text=d["text"],
            purpose=d.get("purpose"),
            source_facts=list(d["source_facts"]) if "source_facts" in d else None,
            impeachment_hook=d.get("impeachment_hook"),
            objection_alts=list(d["objection_alts"]) if "objection_alts" in d else None,
        )


@dataclass
class TopicQuestions:
    topic_id: str
    questions: List[Question] = field(default_factory=list)
    error: Optional[str] = None

    def to_dict(self) -> dict:
        d = {"topic_id": self.topic_id, "questions": [q.to_dict() for q in self.questions]}
        if self.error is not None: d["error"] = self.error
        return d

    @classmethod
    def from_dict(cls, d: dict) -> "TopicQuestions":
        return cls(
            topic_id=d["topic_id"],
            questions=[Question.from_dict(q) for q in d.get("questions", [])],
            error=d.get("error"),
        )


_DIGEST_REQUIRED = ("source_id", "source_kind", "deponent_statements", "factual_anchors", "inconsistencies")


def validate_source_digest_dict(d: dict) -> None:
    """Raise ValueError if d is not a valid digest dict."""
    if not isinstance(d, dict):
        raise ValueError("source digest must be a dict")
    missing = [k for k in _DIGEST_REQUIRED if k not in d]
    if missing:
        raise ValueError(f"source digest missing keys: {missing}")
    for list_key in ("deponent_statements", "factual_anchors", "inconsistencies"):
        if not isinstance(d[list_key], list):
            raise ValueError(f"{list_key} must be a list")


def validate_topics_dict(d: dict) -> None:
    """Raise ValueError if d is not a valid topics dict."""
    if not isinstance(d, dict) or "topics" not in d:
        raise ValueError("topics payload must have 'topics' key")
    if not isinstance(d["topics"], list):
        raise ValueError("topics must be a list")
    for t in d["topics"]:
        for k in ("id", "title"):
            if k not in t:
                raise ValueError(f"topic missing '{k}': {t}")
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest tests/test_wizard/test_depo_prep_schemas.py -v`
Expected: 7 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/__init__.py Scripts/depo_prep_lib/schemas.py tests/test_wizard/test_depo_prep_schemas.py
git commit -m "feat(depo_prep): schemas — dataclasses + validators for digest/topics/questions"
```

---

### Task 2: Session I/O helpers

**Files:**
- Create: `Scripts/depo_prep_lib/session_io.py`
- Create: `tests/test_wizard/test_depo_prep_session_io.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_session_io.py
"""Session folder layout + JSON read/write tests."""
import json
import os
from pathlib import Path

import pytest

from Scripts.depo_prep_lib.session_io import (
    SessionPaths,
    build_session_folder_name,
    compute_session_paths,
    write_json,
    read_json,
    file_sha256,
)


def test_build_session_folder_name_includes_deponent_and_date():
    name = build_session_folder_name("Jane Doe", when_iso="2026-05-27T14:32:00")
    assert "Depo Prep" in name
    assert "Jane Doe" in name
    assert "2026-05-27" in name
    # No illegal Windows filename chars
    for bad in '\\/*?:"<>|':
        assert bad not in name


def test_build_session_folder_name_sanitizes_deponent():
    name = build_session_folder_name('Joe "Slick" O\'Malley/Jr', when_iso="2026-05-27T00:00:00")
    for bad in '\\/*?:"<>|':
        assert bad not in name
    assert "Slick" in name  # kept words, dropped quotes


def test_compute_session_paths_creates_expected_subpaths(tmp_path):
    case_root = tmp_path / "Smith v. Jones"
    case_root.mkdir()
    paths = compute_session_paths(
        case_root=str(case_root),
        deponent_name="Jane Doe",
        when_iso="2026-05-27T14:32:00",
    )
    assert isinstance(paths, SessionPaths)
    assert "NOTES" in str(paths.session_dir)
    assert "AI Output" in str(paths.session_dir)
    assert paths.session_json.name == "session.json"
    assert paths.topics_json.name == "topics.json"
    assert paths.digests_dir.name == "digests"
    assert paths.outline_docx.name == "outline.docx"
    assert paths.outline_md.name == "outline.md"


def test_write_and_read_json_roundtrip(tmp_path):
    p = tmp_path / "x.json"
    write_json(p, {"a": 1, "b": ["x", "y"]})
    assert read_json(p) == {"a": 1, "b": ["x", "y"]}


def test_write_json_creates_parent_dirs(tmp_path):
    p = tmp_path / "deep" / "nested" / "x.json"
    write_json(p, {"ok": True})
    assert p.exists()
    assert read_json(p)["ok"] is True


def test_file_sha256_is_stable(tmp_path):
    f = tmp_path / "x.txt"
    f.write_bytes(b"hello world")
    h1 = file_sha256(f)
    h2 = file_sha256(f)
    assert h1 == h2
    assert len(h1) == 64  # hex SHA-256
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_session_io.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/session_io.py
"""Session folder layout + JSON helpers for Depo Prep."""
from __future__ import annotations

import hashlib
import json
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Union


# Windows-illegal filename characters (matches summarize_deposition.py convention).
_UNSAFE_FILENAME_CHARS = re.compile(r'[\\/*?:"<>|]')


@dataclass(frozen=True)
class SessionPaths:
    session_dir: Path
    session_json: Path
    topics_json: Path
    digests_dir: Path
    raw_dir: Path
    outline_docx: Path
    outline_md: Path
    trace_log: Path


def build_session_folder_name(deponent_name: str, when_iso: str | None = None) -> str:
    """Return a Windows-safe folder name like 'Depo Prep - Jane Doe - 2026-05-27 1432'."""
    when_iso = when_iso or datetime.now().isoformat(timespec="minutes")
    # Strip illegal chars; collapse whitespace.
    safe_name = _UNSAFE_FILENAME_CHARS.sub("", deponent_name or "Unknown").strip()
    safe_name = re.sub(r"\s+", " ", safe_name) or "Unknown"
    # "2026-05-27T14:32" → "2026-05-27 1432"
    when = when_iso.replace("T", " ").replace(":", "")
    return f"Depo Prep - {safe_name} - {when}"


def compute_session_paths(case_root: str, deponent_name: str, when_iso: str | None = None) -> SessionPaths:
    """Compute the canonical layout for a Depo Prep session.

    Layout: {case_root}/NOTES/AI Output/{folder_name}/
              session.json
              topics.json
              digests/
                raw/<source>.txt
                <source>.json
              outline.docx
              outline.md
              trace.log
    """
    folder = build_session_folder_name(deponent_name, when_iso=when_iso)
    session_dir = Path(case_root) / "NOTES" / "AI Output" / folder
    digests = session_dir / "digests"
    return SessionPaths(
        session_dir=session_dir,
        session_json=session_dir / "session.json",
        topics_json=session_dir / "topics.json",
        digests_dir=digests,
        raw_dir=digests / "raw",
        outline_docx=session_dir / "outline.docx",
        outline_md=session_dir / "outline.md",
        trace_log=session_dir / "trace.log",
    )


def write_json(path: Union[str, Path], data) -> None:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def read_json(path: Union[str, Path]):
    with Path(path).open("r", encoding="utf-8") as f:
        return json.load(f)


def file_sha256(path: Union[str, Path]) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_session_io.py -v`
Expected: 6 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/session_io.py tests/test_wizard/test_depo_prep_session_io.py
git commit -m "feat(depo_prep): session_io — session folder layout + JSON/sha256 helpers"
```

---

### Task 3: Prompt library

**Files:**
- Create: `Scripts/depo_prep_lib/prompts.py`
- Create: `tests/test_wizard/test_depo_prep_prompts.py`

The prompts go in code (not external `.txt` files) because they're small, style-parameterized, and easier to unit-test inline. They each return a `(system_or_instructions: str, text_payload: str)` tuple shaped for `LLMCaller.call(prompt=..., text=...)`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_wizard/test_depo_prep_prompts.py
import pytest

from Scripts.depo_prep_lib.prompts import (
    build_per_source_digest_prompt,
    build_topic_clustering_prompt,
    build_per_topic_questions_prompt,
    build_dedup_prompt,
    build_polish_prompt,
    STYLE_DIRECTIVES,
)


def test_style_directives_has_all_four():
    assert set(STYLE_DIRECTIVES.keys()) == {"discovery", "lockdown", "expert", "friendly"}
    for v in STYLE_DIRECTIVES.values():
        assert isinstance(v, str) and len(v) > 50  # non-trivial directive


def test_per_source_digest_prompt_mentions_deponent_and_kind_hint():
    prompt, text_payload = build_per_source_digest_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff", source_text="...transcript...",
        source_filename="jane_doe_depo.pdf",
    )
    assert "Jane Doe" in prompt
    assert "Plaintiff" in prompt
    assert "deponent_statements" in prompt  # field name guidance
    assert "factual_anchors" in prompt
    assert "inconsistencies" in prompt
    assert "JSON" in prompt
    assert text_payload == "...transcript..."


def test_topic_clustering_prompt_includes_style_and_count_range():
    prompt, text_payload = build_topic_clustering_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff",
        style="lockdown",
        free_text_notes="Focus on causation and prior injuries.",
        digests_summary_text="...digests...",
    )
    assert "8" in prompt and "15" in prompt  # 8-15 topic range
    assert "lockdown" in prompt.lower() or "lock-down" in prompt.lower() or STYLE_DIRECTIVES["lockdown"][:30] in prompt
    assert "causation" in prompt  # free text injected
    assert text_payload == "...digests..."


def test_per_topic_questions_prompt_conditionally_includes_field_instructions():
    # All flags off → only basic question text
    prompt, _ = build_per_topic_questions_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff", style="discovery",
        topic_title="Pre-existing conditions",
        strategic_note="Establish chronic LBP",
        digest_excerpts_text="...digest excerpts...",
        free_text_notes="",
        include_strategic_note=False,
        include_source_facts=False,
        include_impeachment_hook=False,
        include_objection_alts=False,
    )
    assert "purpose" not in prompt.lower()
    assert "source_facts" not in prompt.lower()
    assert "impeachment" not in prompt.lower()
    assert "objection" not in prompt.lower()

    # All flags on → all field instructions present
    prompt2, _ = build_per_topic_questions_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff", style="discovery",
        topic_title="Pre-existing conditions",
        strategic_note="Establish chronic LBP",
        digest_excerpts_text="...digest excerpts...",
        free_text_notes="",
        include_strategic_note=True,
        include_source_facts=True,
        include_impeachment_hook=True,
        include_objection_alts=True,
    )
    assert "purpose" in prompt2.lower()
    assert "source_facts" in prompt2
    assert "impeachment_hook" in prompt2 or "impeachment" in prompt2.lower()
    assert "objection_alts" in prompt2 or "objection" in prompt2.lower()


def test_per_topic_questions_truncates_free_text_above_2000_chars():
    long_notes = "x" * 5000
    prompt, _ = build_per_topic_questions_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff", style="discovery",
        topic_title="t", strategic_note="s", digest_excerpts_text="d",
        free_text_notes=long_notes,
        include_strategic_note=True, include_source_facts=False,
        include_impeachment_hook=False, include_objection_alts=False,
    )
    # The truncation cap is 2000 chars; we should see fewer than 2050 'x's plus a marker.
    assert prompt.count("x") < 2100
    assert "truncated" in prompt.lower()


def test_dedup_prompt_takes_topic_outputs_summary():
    prompt, text = build_dedup_prompt(topic_outputs_summary="Topic 1: 4 Qs\nTopic 2: 5 Qs",
                                       digest_summary="med_records: causation, gaps")
    assert "duplicates" in prompt.lower()
    assert "coverage" in prompt.lower()
    assert "Topic 1" in text or "Topic 1" in prompt


def test_polish_prompt_forbids_substantive_changes():
    prompt, text = build_polish_prompt(outline_text="full outline here")
    # The polish prompt must explicitly forbid adding/dropping/changing questions.
    assert "no new" in prompt.lower() or "do not add" in prompt.lower()
    assert "drop" in prompt.lower() or "remove" in prompt.lower()
    assert "phrasing" in prompt.lower() or "transitions" in prompt.lower()
    assert text == "full outline here"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_wizard/test_depo_prep_prompts.py -v`
Expected: ImportError.

- [ ] **Step 3: Write minimal implementation**

```python
# Scripts/depo_prep_lib/prompts.py
"""Prompt builders for Depo Prep. Each builder returns (prompt, text_payload)."""
from __future__ import annotations

from typing import Tuple


_FREE_TEXT_CAP = 2000


STYLE_DIRECTIVES = {
    "discovery": (
        "Style: DISCOVERY / FACT-GATHERING. Use open-ended questions designed to "
        "develop the witness's complete account. Prefer 'Tell me about…', "
        "'What happened next?', 'Walk me through…' phrasings. Do not lead. The goal "
        "is to uncover testimony, not to box the witness in."
    ),
    "lockdown": (
        "Style: LOCK-DOWN / LEADING. Use short, closed, leading questions designed "
        "to extract specific admissions for use at trial or MSJ. Prefer "
        "'Isn't it true that…', 'You agree that…', 'You did X, correct?' phrasings. "
        "Each question should produce a yes/no answer or a precise concession. "
        "Keep questions short — never compound."
    ),
    "expert": (
        "Style: EXPERT CHALLENGE (Daubert-style). Probe methodology, qualifications, "
        "scope of opinions, materials reviewed, and the reliability/general acceptance "
        "of the underlying methods. Look for ipse dixit reasoning, gaps in the analysis, "
        "and reliance on unreliable foundations. Tie every opinion to specific support."
    ),
    "friendly": (
        "Style: FRIENDLY (own-client prep). Use clear, organized questions that allow "
        "the witness to tell their story cleanly. Anticipate weaknesses and address "
        "them head-on with rehabilitative questions. Avoid jargon. Build credibility."
    ),
}


def _clip(text: str, cap: int = _FREE_TEXT_CAP) -> str:
    if not text:
        return ""
    if len(text) <= cap:
        return text
    return text[:cap] + f"\n[...truncated at {cap} chars]"


def build_per_source_digest_prompt(
    *, deponent_name: str, deponent_role: str, source_text: str, source_filename: str,
) -> Tuple[str, str]:
    prompt = f"""You are an extraction agent for a deposition-prep tool.

The deponent we are preparing to depose is: {deponent_name} ({deponent_role}).

The following source document is named: {source_filename}.

Your job: read the document and produce a STRUCTURED JSON DIGEST of facts and quotes
relevant to deposing this witness. Output **JSON ONLY**, no commentary.

Schema (exact field names required):

{{
  "source_id": "{source_filename}",
  "source_kind": "medical_records | deposition_transcript | discovery_response | pleading | other",
  "deponent_statements": [
    {{ "text": "<verbatim quote from the witness>",
       "location": "<page/line citation if available>",
       "context": "<who was questioning, what segment>" }}
  ],
  "factual_anchors": [
    {{ "fact": "<short factual claim found in the doc, e.g. 'MRI 2024-09-12 showed 4mm protrusion'>",
       "location": "<page/Bates citation>",
       "topic_tags": ["<short free-form tags for clustering, e.g. 'injury', 'causation'>"] }}
  ],
  "inconsistencies": [
    {{ "claim_a": "...", "claim_a_source": "...",
       "claim_b": "...", "claim_b_source": "...",
       "topic_tags": ["credibility", "..."] }}
  ],
  "summary": "<2-3 sentence summary of what this source contributes to the depo prep>"
}}

Rules:
- Quote verbatim where possible. Do not paraphrase witness statements.
- Use citations the document itself contains (page numbers, Bates, page:line).
- If a list has no entries, return an empty list (never null).
- Output JSON only — no markdown fences, no preamble.
"""
    return prompt, source_text


def build_topic_clustering_prompt(
    *, deponent_name: str, deponent_role: str, style: str,
    free_text_notes: str, digests_summary_text: str,
) -> Tuple[str, str]:
    style_block = STYLE_DIRECTIVES.get(style, STYLE_DIRECTIVES["discovery"])
    notes_block = _clip(free_text_notes, _FREE_TEXT_CAP)
    prompt = f"""You are designing the topic structure for a deposition outline.

Deponent: {deponent_name} ({deponent_role}).

{style_block}

Lawyer's strategy notes:
{notes_block or '(none provided)'}

You will be given a concatenated set of per-source digests (JSON) describing the
case material. Your job: cluster the facts/quotes/inconsistencies into 8–15 TOPICS
that organize the deposition outline.

Each topic must include:
- "id": stable short id ("t01", "t02", ...)
- "title": 3–8 word topic name in title case
- "strategic_note": 1–3 sentences naming what the lawyer is trying to ESTABLISH,
   UNDERMINE, or LOCK DOWN under this topic, anchored to the lawyer's strategy notes
   when relevant
- "relevant_digest_refs": list of strings in the form
   "<source_id>#<schema_field>[<index>]" pointing to the digest entries this topic
   draws on. Example: "med_records.pdf#factual_anchors[2]".
- "default_checked": true unless the topic is genuinely speculative

Output JSON ONLY (no fences, no commentary):
{{ "topics": [ {{...}}, {{...}} ] }}

Produce between 8 and 15 topics. Aim for clean, non-overlapping coverage.
"""
    return prompt, digests_summary_text


def build_per_topic_questions_prompt(
    *, deponent_name: str, deponent_role: str, style: str,
    topic_title: str, strategic_note: str, digest_excerpts_text: str,
    free_text_notes: str,
    include_strategic_note: bool, include_source_facts: bool,
    include_impeachment_hook: bool, include_objection_alts: bool,
) -> Tuple[str, str]:
    style_block = STYLE_DIRECTIVES.get(style, STYLE_DIRECTIVES["discovery"])
    notes_block = _clip(free_text_notes, _FREE_TEXT_CAP)

    field_instructions = []
    if include_strategic_note:
        field_instructions.append(
            '  "purpose": "<one-sentence statement of what this question is trying to '
            'establish, lock in, undermine, or develop>",'
        )
    if include_source_facts:
        field_instructions.append(
            '  "source_facts": [ "<bullet pointing to the specific document/page/quote '
            'that justifies this question>", ... ],'
        )
    if include_impeachment_hook:
        field_instructions.append(
            '  "impeachment_hook": "<if the witness denies / equivocates, the exact '
            'prior statement or document to confront them with>",'
        )
    if include_objection_alts:
        field_instructions.append(
            '  "objection_alts": [ "<cleaner rephrasing if opposing counsel objects '
            'vague/compound/asked-and-answered>", ... ],'
        )

    optional_fields_block = "\n".join(field_instructions) if field_instructions else "  (only 'n' and 'text' fields)"

    prompt = f"""You are drafting deposition questions for one topic of an outline.

Deponent: {deponent_name} ({deponent_role}).

{style_block}

Topic: {topic_title}
Strategic note (what we're trying to establish): {strategic_note}

Lawyer's overall strategy notes:
{notes_block or '(none provided)'}

You will be given digest excerpts (verbatim quotes, factual anchors, inconsistencies)
relevant to this topic. Draft 5–10 questions that probe this topic, grounded in
those source facts.

Output JSON ONLY (no fences, no commentary):
{{
  "topic_id": "<echoed>",
  "questions": [
    {{
      "n": 1,
      "text": "<the question itself; never compound>",
{optional_fields_block}
    }},
    ...
  ]
}}

Rules:
- Never invent facts not present in the digest excerpts. Every factual claim in a
  question must be traceable to the excerpts.
- Each question is a single, clean inquiry — no compound questions.
- Number sequentially starting at 1.
- If you cannot generate meaningful questions for this topic (e.g., no relevant
  source facts), return an empty questions list.
"""
    return prompt, digest_excerpts_text


def build_dedup_prompt(*, topic_outputs_summary: str, digest_summary: str) -> Tuple[str, str]:
    prompt = """You are auditing a deposition outline for duplicates and coverage gaps.

You will be given:
- A summary of every topic's questions (numbered as "<topic_id>.q<n>: <text>").
- A summary of the source-digest facts available.

Identify:
1. Duplicate questions across topics — same substantive ask, possibly different phrasing.
   For each pair, choose which to KEEP and which to DROP.
2. Coverage gaps — important facts from the digest that no question addresses.

Output JSON ONLY:
{
  "duplicates": [
    { "keep": "<topic_id>.q<n>", "drop": "<topic_id>.q<n>", "reason": "<short>" }
  ],
  "coverage_gaps": [ "<one line, ending with a topic suggestion>" ],
  "renumber_after_dedup": true
}

Be conservative. Only flag duplicates that are substantively the same.
"""
    return prompt, f"=== TOPIC OUTPUTS ===\n{topic_outputs_summary}\n\n=== DIGEST SUMMARY ===\n{digest_summary}"


def build_polish_prompt(*, outline_text: str) -> Tuple[str, str]:
    prompt = """You are doing a final phrasing pass on a deposition outline.

ABSOLUTE RULES:
- Do not add any new questions.
- Do not drop or remove any questions.
- Do not change any factual content in a question — only phrasing.
- Do not change strategic notes substantively.
- Do not renumber questions.

Allowed changes:
- Tighten redundant phrasing.
- Add brief topic-to-topic transitions in strategic notes only.
- Normalize question phrasing consistency within a topic.
- Fix obvious typos.

Return the polished outline in the SAME structure as input, JSON ONLY.
"""
    return prompt, outline_text
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest tests/test_wizard/test_depo_prep_prompts.py -v`
Expected: 7 passed.

- [ ] **Step 5: Commit**

```bash
git add Scripts/depo_prep_lib/prompts.py tests/test_wizard/test_depo_prep_prompts.py
git commit -m "feat(depo_prep): prompt library — style-aware builders for all stages"
```

---

### Wave 1 self-check

Before starting Wave 2, verify:

- [ ] `python -m pytest tests/test_wizard/test_depo_prep_schemas.py tests/test_wizard/test_depo_prep_session_io.py tests/test_wizard/test_depo_prep_prompts.py -v` → all green
- [ ] `Scripts/depo_prep_lib/__init__.py`, `schemas.py`, `session_io.py`, `prompts.py` all exist
- [ ] No new imports in non-test code outside `Scripts/depo_prep_lib/`
- [ ] Three commits on the branch, one per task

Proceed to Wave 2.
