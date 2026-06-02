# Document Text Library Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Persist the source text iCharlotte extracts when it runs tasks into a per-case library, and let the Chat tab query that text via a checkbox list without re-uploading or re-extracting.

**Architecture:** A greenfield `icharlotte_core/doc_library/` package owns extraction + storage (per-case, under the case's `.icharlotte/doc_library/` folder, content-hashed blobs with ref-counted dedup). Tasks auto-populate it by listening to the existing `TaskTab.task_completed` signal (no worker changes). The Chat tab gains a "Saved Documents" tree that loads cached blob text straight into the existing `read_files_content()` pipeline.

**Tech Stack:** Python 3.13, PySide6 (the app is PySide6, not PyQt6), `python-docx`, PyMuPDF (`fitz`), existing `DocumentProcessor` / `extract_docx_text` / `TokenCounter`. Tests: `pytest`, `pytest.importorskip("pytestqt")` for UI.

**Spec:** `docs/superpowers/specs/2026-06-01-document-text-library-design.md`

---

## File Structure

**New package — `icharlotte_core/doc_library/`:**
- `__init__.py` — public exports.
- `models.py` — `MemberFile`, `LibraryEntry` dataclasses with `to_dict`/`from_dict`.
- `extract.py` — `Extracted` dataclass + `extract_any(path)` dispatcher (reuses `DocumentProcessor` / `extract_docx_text` / `.doc` COM).
- `labels.py` — `auto_label(task_type, source_paths, metadata, existing_labels)`.
- `library.py` — `DocumentLibrary` (paths, atomic load/save, `list_entries`, `add_entry`, `rename_entry`, `reset_label`, `delete_entry`, `get_member_text`).
- `capture.py` — `AUTO_CAPTURE_TASK_IDS` allow-list + `capture_from_task_entry(case_root, entry)`.

**Modified:**
- `icharlotte_core/ui/wizard/task_tab.py` — connect a library capture to `task_completed` (or wire it at the container that owns `case_root`; see Task 8).
- `icharlotte_core/ui/tabs.py` — `ChatTab`: "Saved Documents" tree, manual add, library text in `read_files_content`/`send_message`, budget warning, selection memory.

**Tests — `tests/test_doc_library/`:**
- `test_models.py`, `test_extract.py`, `test_labels.py`, `test_library.py`, `test_capture.py`, `test_chat_library_integration.py`.

---

## Phase 1 — Backend package

### Task 1: Data models

**Files:**
- Create: `icharlotte_core/doc_library/__init__.py`
- Create: `icharlotte_core/doc_library/models.py`
- Test: `tests/test_doc_library/__init__.py`, `tests/test_doc_library/test_models.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_doc_library/test_models.py
from icharlotte_core.doc_library.models import MemberFile, LibraryEntry


def test_memberfile_roundtrips_through_dict():
    m = MemberFile(
        source_path=r"Z:\case\Depo Vol 1.pdf",
        source_name="Depo Vol 1.pdf",
        blob="abc123.txt",
        char_count=100,
        est_tokens=25,
        extract_method="pdf_native",
        error=None,
    )
    assert MemberFile.from_dict(m.to_dict()) == m


def test_entry_roundtrips_and_defaults_members():
    e = LibraryEntry(
        id="id1",
        label="Plaintiff's Deposition Transcript",
        auto_label="Plaintiff's Deposition Transcript",
        task_type="summarize_depositions",
        created_at="2026-06-02T10:00:00",
        members=[MemberFile("p", "p", "h.txt")],
    )
    d = e.to_dict()
    assert d["members"][0]["blob"] == "h.txt"
    assert LibraryEntry.from_dict(d) == e


def test_entry_from_dict_tolerates_missing_members():
    e = LibraryEntry.from_dict({
        "id": "x", "label": "L", "auto_label": "L",
        "task_type": "manual", "created_at": "t",
    })
    assert e.members == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_models.py -v`
Expected: FAIL — `ModuleNotFoundError: icharlotte_core.doc_library`.

- [ ] **Step 3: Write minimal implementation**

```python
# icharlotte_core/doc_library/models.py
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class MemberFile:
    """One source document inside a library entry."""
    source_path: str
    source_name: str
    blob: Optional[str]          # "<sha1>.txt" in blobs/, or None if extraction failed
    char_count: int = 0
    est_tokens: int = 0
    extract_method: str = ""
    error: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "source_path": self.source_path,
            "source_name": self.source_name,
            "blob": self.blob,
            "char_count": self.char_count,
            "est_tokens": self.est_tokens,
            "extract_method": self.extract_method,
            "error": self.error,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "MemberFile":
        return cls(
            source_path=d.get("source_path", ""),
            source_name=d.get("source_name", ""),
            blob=d.get("blob"),
            char_count=int(d.get("char_count", 0)),
            est_tokens=int(d.get("est_tokens", 0)),
            extract_method=d.get("extract_method", ""),
            error=d.get("error"),
        )


@dataclass
class LibraryEntry:
    """One task run (or manual add); the expandable unit in the UI."""
    id: str
    label: str
    auto_label: str
    task_type: str
    created_at: str
    members: list = field(default_factory=list)  # list[MemberFile]

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "label": self.label,
            "auto_label": self.auto_label,
            "task_type": self.task_type,
            "created_at": self.created_at,
            "members": [m.to_dict() for m in self.members],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "LibraryEntry":
        return cls(
            id=d.get("id", ""),
            label=d.get("label", ""),
            auto_label=d.get("auto_label", d.get("label", "")),
            task_type=d.get("task_type", ""),
            created_at=d.get("created_at", ""),
            members=[MemberFile.from_dict(m) for m in d.get("members", [])],
        )
```

```python
# icharlotte_core/doc_library/__init__.py
from .models import MemberFile, LibraryEntry

__all__ = ["MemberFile", "LibraryEntry"]
```

```python
# tests/test_doc_library/__init__.py
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_models.py -v`
Expected: PASS (3 passed).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/doc_library/__init__.py icharlotte_core/doc_library/models.py tests/test_doc_library/__init__.py tests/test_doc_library/test_models.py
git commit -m "feat(doc_library): MemberFile + LibraryEntry models"
```

---

### Task 2: `extract_any` dispatcher

**Files:**
- Create: `icharlotte_core/doc_library/extract.py`
- Test: `tests/test_doc_library/test_extract.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_doc_library/test_extract.py
from icharlotte_core.doc_library.extract import extract_any, Extracted


def test_extract_txt(tmp_path):
    p = tmp_path / "note.txt"
    p.write_text("hello world", encoding="utf-8")
    out = extract_any(str(p))
    assert isinstance(out, Extracted)
    assert "hello world" in out.text
    assert out.error is None


def test_extract_docx(tmp_path):
    from docx import Document
    p = tmp_path / "d.docx"
    doc = Document()
    doc.add_paragraph("First para")
    doc.save(str(p))
    out = extract_any(str(p))
    assert "First para" in out.text
    assert out.extract_method == "docx"
    assert out.error is None


def test_extract_missing_file_returns_error():
    out = extract_any(r"Z:\nope\missing.pdf")
    assert out.text == ""
    assert out.error  # non-empty error string
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_extract.py -v`
Expected: FAIL — `ModuleNotFoundError` / `ImportError: extract_any`.

- [ ] **Step 3: Write minimal implementation**

```python
# icharlotte_core/doc_library/extract.py
"""Single-call text extraction reused by the document library.

Dispatches by extension, reusing the app's existing extractors so the library
never re-implements OCR/Word logic:
  - .docx -> extract_docx_text (tables + headers, document order)
  - .doc  -> Word COM read (read-only attach; never Quit / never set Visible)
  - everything else (.pdf/.txt/.msg/...) -> DocumentProcessor.extract_text
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from typing import Optional


@dataclass
class Extracted:
    text: str
    page_count: int
    extract_method: str
    error: Optional[str]


def extract_any(path: str) -> Extracted:
    if not os.path.exists(path):
        return Extracted("", 0, "failed", f"File not found: {path}")

    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".docx":
            from ..document_processor import extract_docx_text
            text = extract_docx_text(path)
            return Extracted(text, 0, "docx", None)
        if ext == ".doc":
            from ..ui.tabs import ChatTab  # reuse the read-only COM helper
            text = ChatTab._extract_doc_text(path)
            err = None
            if text.strip().startswith("[Error reading .doc"):
                err = text.strip()
            return Extracted(text, 0, "doc_com", err)
        # PDF / TXT / MSG / other -> DocumentProcessor
        from ..document_processor import DocumentProcessor, ExtractionMethod
        result = DocumentProcessor().extract_text(path, ocr_enabled=True)
        err = result.error
        if result.extraction_method == ExtractionMethod.FAILED and not err:
            err = "extraction failed"
        return Extracted(
            text=result.text or "",
            page_count=result.page_count or 0,
            extract_method=result.extraction_method.value,
            error=err,
        )
    except Exception as e:  # never raise into a task-completion handler
        return Extracted("", 0, "failed", f"{type(e).__name__}: {e}")
```

> Note: importing `ChatTab` for `_extract_doc_text` is a static-method reuse — no Qt instance is created. If the executor finds this import too heavy (it pulls the UI module), copy `_extract_doc_text` verbatim into `extract.py` instead; it has no `self` dependencies.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_extract.py -v`
Expected: PASS (3 passed). `.doc`/`.msg`/OCR paths are exercised in production, not unit tests (COM/Tesseract unavailable in CI).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/doc_library/extract.py tests/test_doc_library/test_extract.py
git commit -m "feat(doc_library): extract_any dispatcher over existing extractors"
```

---

### Task 3: Auto-labeling

**Files:**
- Create: `icharlotte_core/doc_library/labels.py`
- Test: `tests/test_doc_library/test_labels.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_doc_library/test_labels.py
from icharlotte_core.doc_library.labels import auto_label


def test_deposition_with_party():
    lbl = auto_label("summarize_depositions", [r"Z:\x\depo.pdf"],
                     {"party": "Plaintiff"}, existing_labels=[])
    assert lbl == "Plaintiff's Deposition Transcript"


def test_deposition_without_party_drops_possessive():
    lbl = auto_label("summarize_depositions", [r"Z:\x\depo.pdf"],
                     {}, existing_labels=[])
    assert lbl == "Deposition Transcript"


def test_discovery_with_party():
    lbl = auto_label("summarize_discovery", [r"Z:\x\rfp.pdf"],
                     {"party": "Defendant"}, existing_labels=[])
    assert lbl == "Defendant's Discovery Responses"


def test_medical_with_name():
    lbl = auto_label("medical_records", [r"Z:\x\recs.pdf"],
                     {"name": "Brier Buchalter"}, existing_labels=[])
    assert lbl == "Medical Records — Brier Buchalter"


def test_summarize_documents_uses_clean_filename():
    lbl = auto_label("summarize_documents",
                     [r"Z:\x\TRAFFIC_COLLISION_REPORT.pdf"],
                     {}, existing_labels=[])
    assert lbl == "Traffic Collision Report"


def test_multi_file_summarize_documents_suffix():
    lbl = auto_label("summarize_documents",
                     [r"Z:\x\a.pdf", r"Z:\x\b.pdf", r"Z:\x\c.pdf"],
                     {}, existing_labels=[])
    assert lbl == "A +2 more"


def test_collision_appends_numeric_suffix():
    existing = ["Deposition Transcript"]
    lbl = auto_label("summarize_depositions", [r"Z:\x\d.pdf"], {}, existing)
    assert lbl == "Deposition Transcript (2)"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_labels.py -v`
Expected: FAIL — `ImportError: auto_label`.

- [ ] **Step 3: Write minimal implementation**

```python
# icharlotte_core/doc_library/labels.py
"""User-friendly, list-friendly auto labels for library entries.

Pattern per task type, filled from best-effort metadata (party/role/name the
task already collected). Falls back to a cleaned filename. Never surfaces a raw
cryptic filename. Collisions get a numeric suffix ("(2)"), matching the
existing wizard instance_naming convention.
"""
from __future__ import annotations

import os
import re

_MED_TASKS = {"medical_records", "med_chron_analysis", "med_record_extractor"}


def _clean_filename(path: str) -> str:
    stem = os.path.splitext(os.path.basename(path))[0]
    stem = re.sub(r"[_\-]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip()
    return stem.title() if stem else "Document"


def _filename_label(source_paths: list) -> str:
    if not source_paths:
        return "Document"
    base = _clean_filename(source_paths[0])
    extra = len(source_paths) - 1
    return f"{base} +{extra} more" if extra > 0 else base


def _base_label(task_type: str, source_paths: list, metadata: dict) -> str:
    party = (metadata.get("party") or "").strip()
    name = (metadata.get("name") or "").strip()
    if task_type == "summarize_depositions":
        return f"{party}'s Deposition Transcript" if party else "Deposition Transcript"
    if task_type == "summarize_discovery":
        return f"{party}'s Discovery Responses" if party else "Discovery Responses"
    if task_type in _MED_TASKS:
        return f"Medical Records — {name}" if name else "Medical Records"
    # summarize_documents, manual, or anything else -> filename
    return _filename_label(source_paths)


def auto_label(task_type: str, source_paths: list, metadata: dict,
               existing_labels: list) -> str:
    metadata = metadata or {}
    base = _base_label(task_type, source_paths, metadata)
    if base not in set(existing_labels):
        return base
    n = 2
    while f"{base} ({n})" in set(existing_labels):
        n += 1
    return f"{base} ({n})"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_labels.py -v`
Expected: PASS (7 passed).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/doc_library/labels.py tests/test_doc_library/test_labels.py
git commit -m "feat(doc_library): friendly auto-labels with collision suffix"
```

---

### Task 4: `DocumentLibrary` — paths, atomic store, `list_entries`

**Files:**
- Create: `icharlotte_core/doc_library/library.py`
- Modify: `icharlotte_core/doc_library/__init__.py`
- Test: `tests/test_doc_library/test_library.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_doc_library/test_library.py
import os
from icharlotte_core.doc_library.library import DocumentLibrary


def test_folder_path_under_icharlotte(tmp_path):
    lib = DocumentLibrary(str(tmp_path))
    assert lib.folder == os.path.join(
        str(tmp_path), "NOTES", "AI OUTPUT", ".icharlotte", "doc_library")


def test_empty_library_lists_nothing(tmp_path):
    assert DocumentLibrary(str(tmp_path)).list_entries() == []


def test_save_then_reload_persists_entries(tmp_path):
    from icharlotte_core.doc_library.models import LibraryEntry, MemberFile
    lib = DocumentLibrary(str(tmp_path))
    entry = LibraryEntry("id1", "L", "L", "manual", "t",
                         [MemberFile("p", "p", "h.txt", 10, 3, "docx")])
    lib._save_entries([entry])
    reloaded = DocumentLibrary(str(tmp_path)).list_entries()
    assert len(reloaded) == 1
    assert reloaded[0].id == "id1"
    assert reloaded[0].members[0].blob == "h.txt"


def test_corrupt_index_is_recovered_not_raised(tmp_path):
    lib = DocumentLibrary(str(tmp_path))
    os.makedirs(lib.folder, exist_ok=True)
    with open(lib.index_path, "w", encoding="utf-8") as f:
        f.write("{ this is not valid json")
    assert lib.list_entries() == []  # backed up + reinitialized
    assert os.path.exists(lib.index_path + ".corrupt")
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_library.py -v`
Expected: FAIL — `ImportError: DocumentLibrary`.

- [ ] **Step 3: Write minimal implementation**

```python
# icharlotte_core/doc_library/library.py
"""Per-case persistent library of extracted source text.

Stored on the case drive, alongside wizard_state.json:
    <case_root>/NOTES/AI OUTPUT/.icharlotte/doc_library/
        index.json          # catalog of LibraryEntry
        blobs/<sha1>.txt     # extracted text, deduped by file content hash

Atomic writes via tmp + os.replace (same pattern as WizardStatePersistence).
"""
from __future__ import annotations

import hashlib
import json
import os
import uuid
from datetime import datetime
from typing import Callable, Optional

from .models import LibraryEntry, MemberFile

SCHEMA_VERSION = 1


class DocumentLibrary:
    def __init__(self, case_root: str):
        self.case_root = case_root

    # ---- paths ----
    @property
    def folder(self) -> str:
        return os.path.join(self.case_root, "NOTES", "AI OUTPUT",
                            ".icharlotte", "doc_library")

    @property
    def index_path(self) -> str:
        return os.path.join(self.folder, "index.json")

    @property
    def blobs_dir(self) -> str:
        return os.path.join(self.folder, "blobs")

    # ---- index load/save ----
    def _load(self) -> dict:
        if not os.path.isfile(self.index_path):
            return {"version": SCHEMA_VERSION, "entries": []}
        try:
            with open(self.index_path, "r", encoding="utf-8") as f:
                raw = json.load(f)
            if not isinstance(raw, dict) or not isinstance(raw.get("entries"), list):
                raise ValueError("bad index shape")
            return raw
        except (json.JSONDecodeError, OSError, ValueError):
            try:
                os.replace(self.index_path, self.index_path + ".corrupt")
            except OSError:
                pass
            return {"version": SCHEMA_VERSION, "entries": []}

    def _save(self, data: dict) -> None:
        os.makedirs(self.folder, exist_ok=True)
        tmp = self.index_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        os.replace(tmp, self.index_path)

    def _save_entries(self, entries: list) -> None:
        self._save({"version": SCHEMA_VERSION,
                    "entries": [e.to_dict() for e in entries]})

    # ---- read ----
    def list_entries(self) -> list:
        return [LibraryEntry.from_dict(d) for d in self._load().get("entries", [])]
```

Update exports:

```python
# icharlotte_core/doc_library/__init__.py
from .models import MemberFile, LibraryEntry
from .extract import extract_any, Extracted
from .labels import auto_label
from .library import DocumentLibrary

__all__ = [
    "MemberFile", "LibraryEntry", "Extracted",
    "extract_any", "auto_label", "DocumentLibrary",
]
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_library.py -v`
Expected: PASS (4 passed).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/doc_library/library.py icharlotte_core/doc_library/__init__.py tests/test_doc_library/test_library.py
git commit -m "feat(doc_library): DocumentLibrary store with atomic index + recovery"
```

---

### Task 5: `add_entry` with content-hash dedup + idempotent replace

**Files:**
- Modify: `icharlotte_core/doc_library/library.py`
- Test: `tests/test_doc_library/test_library.py`

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_doc_library/test_library.py
from icharlotte_core.doc_library.extract import Extracted


def _fake_extractor(text="DEPO TEXT"):
    return lambda path: Extracted(text=text, page_count=2,
                                  extract_method="pdf_native", error=None)


def test_add_entry_creates_blob_and_entry(tmp_path):
    src = tmp_path / "depo.pdf"
    src.write_bytes(b"%PDF-1.4 fake bytes")
    lib = DocumentLibrary(str(tmp_path))
    entry = lib.add_entry("summarize_depositions", [str(src)],
                          {"party": "Plaintiff"}, extractor=_fake_extractor())
    assert entry.label == "Plaintiff's Deposition Transcript"
    assert entry.members[0].char_count == len("DEPO TEXT")
    assert entry.members[0].est_tokens > 0
    blob = os.path.join(lib.blobs_dir, entry.members[0].blob)
    assert os.path.isfile(blob)
    assert lib.get_member_text(entry.members[0].blob) == "DEPO TEXT"
    assert len(lib.list_entries()) == 1


def test_same_file_in_two_entries_extracts_once(tmp_path):
    src = tmp_path / "depo.pdf"
    src.write_bytes(b"same bytes")
    calls = {"n": 0}

    def counting_extractor(path):
        calls["n"] += 1
        return Extracted("X", 1, "pdf_native", None)

    lib = DocumentLibrary(str(tmp_path))
    lib.add_entry("manual", [str(src)], {}, extractor=counting_extractor)
    lib.add_entry("summarize_documents", [str(src)], {}, extractor=counting_extractor)
    assert calls["n"] == 1  # second add reused the blob


def test_rerun_same_task_same_files_replaces_entry(tmp_path):
    src = tmp_path / "depo.pdf"
    src.write_bytes(b"bytes")
    lib = DocumentLibrary(str(tmp_path))
    lib.add_entry("summarize_depositions", [str(src)], {"party": "Plaintiff"},
                  extractor=_fake_extractor())
    lib.add_entry("summarize_depositions", [str(src)], {"party": "Plaintiff"},
                  extractor=_fake_extractor())
    assert len(lib.list_entries()) == 1  # replaced, not duplicated


def test_extraction_failure_records_error_member_no_blob(tmp_path):
    src = tmp_path / "bad.pdf"
    src.write_bytes(b"bytes")
    lib = DocumentLibrary(str(tmp_path))
    fail = lambda path: Extracted("", 0, "failed", "boom")
    entry = lib.add_entry("manual", [str(src)], {}, extractor=fail)
    assert entry.members[0].error == "boom"
    assert entry.members[0].blob is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_library.py -k add_entry -v`
Expected: FAIL — `AttributeError: 'DocumentLibrary' object has no attribute 'add_entry'`.

- [ ] **Step 3: Write minimal implementation**

Add to `DocumentLibrary` in `library.py` (and the import for tokens):

```python
# at top of library.py, alongside other imports
from .extract import extract_any as _default_extractor
from ..chat.token_counter import TokenCounter
```

```python
    # ---- hashing / blobs ----
    @staticmethod
    def _hash_file(path: str) -> str:
        h = hashlib.sha1()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()

    def get_member_text(self, blob: Optional[str]) -> str:
        if not blob:
            return ""
        path = os.path.join(self.blobs_dir, blob)
        if not os.path.isfile(path):
            return ""
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            return f.read()

    # ---- write ----
    def add_entry(self, task_type: str, source_paths: list,
                  metadata: Optional[dict] = None,
                  extractor: Callable = _default_extractor) -> LibraryEntry:
        os.makedirs(self.blobs_dir, exist_ok=True)
        entries = self.list_entries()

        # Idempotency: drop any prior entry with the same task_type + file set.
        target_set = {os.path.normcase(os.path.abspath(p)) for p in source_paths}
        kept = []
        for e in entries:
            e_set = {os.path.normcase(os.path.abspath(m.source_path)) for m in e.members}
            if not (e.task_type == task_type and e_set == target_set):
                kept.append(e)
        entries = kept

        members = []
        for path in source_paths:
            name = os.path.basename(path)
            try:
                digest = self._hash_file(path)
            except OSError as ex:
                members.append(MemberFile(path, name, None, 0, 0, "failed", str(ex)))
                continue
            blob_name = f"{digest}.txt"
            blob_path = os.path.join(self.blobs_dir, blob_name)
            if os.path.isfile(blob_path):
                text = self.get_member_text(blob_name)
                members.append(MemberFile(
                    path, name, blob_name, len(text),
                    TokenCounter.estimate_tokens(text), "cached", None))
                continue
            result = extractor(path)
            if result.error or not result.text:
                members.append(MemberFile(
                    path, name, None, 0, 0, result.extract_method,
                    result.error or "empty extraction"))
                continue
            tmp = blob_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                f.write(result.text)
            os.replace(tmp, blob_path)
            members.append(MemberFile(
                path, name, blob_name, len(result.text),
                TokenCounter.estimate_tokens(result.text),
                result.extract_method, None))

        existing_labels = [e.label for e in entries]
        label = auto_label(task_type, source_paths, metadata or {}, existing_labels)
        entry = LibraryEntry(
            id=uuid.uuid4().hex,
            label=label,
            auto_label=label,
            task_type=task_type,
            created_at=datetime.now().isoformat(timespec="seconds"),
            members=members,
        )
        entries.append(entry)
        self._save_entries(entries)
        return entry
```

Add the `auto_label` import at top of `library.py`:

```python
from .labels import auto_label
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_library.py -v`
Expected: PASS (all add_entry tests + Task 4 tests green).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/doc_library/library.py tests/test_doc_library/test_library.py
git commit -m "feat(doc_library): add_entry with content-hash dedup + idempotent replace"
```

---

### Task 6: rename / reset / delete (ref-counted blob GC)

**Files:**
- Modify: `icharlotte_core/doc_library/library.py`
- Test: `tests/test_doc_library/test_library.py`

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_doc_library/test_library.py
def test_rename_and_reset(tmp_path):
    src = tmp_path / "d.pdf"
    src.write_bytes(b"b")
    lib = DocumentLibrary(str(tmp_path))
    e = lib.add_entry("summarize_depositions", [str(src)], {"party": "Plaintiff"},
                      extractor=_fake_extractor())
    lib.rename_entry(e.id, "My Custom Name")
    assert lib.list_entries()[0].label == "My Custom Name"
    lib.reset_label(e.id)
    assert lib.list_entries()[0].label == "Plaintiff's Deposition Transcript"


def test_delete_entry_gcs_unreferenced_blob(tmp_path):
    src = tmp_path / "d.pdf"
    src.write_bytes(b"b")
    lib = DocumentLibrary(str(tmp_path))
    e = lib.add_entry("manual", [str(src)], {}, extractor=_fake_extractor())
    blob_path = os.path.join(lib.blobs_dir, e.members[0].blob)
    assert os.path.isfile(blob_path)
    lib.delete_entry(e.id)
    assert lib.list_entries() == []
    assert not os.path.isfile(blob_path)  # GC'd: no other entry referenced it


def test_delete_keeps_blob_referenced_by_another_entry(tmp_path):
    src = tmp_path / "d.pdf"
    src.write_bytes(b"b")
    lib = DocumentLibrary(str(tmp_path))
    e1 = lib.add_entry("manual", [str(src)], {}, extractor=_fake_extractor())
    e2 = lib.add_entry("summarize_documents", [str(src)], {}, extractor=_fake_extractor())
    blob_path = os.path.join(lib.blobs_dir, e1.members[0].blob)
    lib.delete_entry(e1.id)
    assert os.path.isfile(blob_path)  # still referenced by e2
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_library.py -k "rename or delete" -v`
Expected: FAIL — `AttributeError: rename_entry`.

- [ ] **Step 3: Write minimal implementation**

Add to `DocumentLibrary`:

```python
    def _find(self, entries: list, entry_id: str):
        for e in entries:
            if e.id == entry_id:
                return e
        return None

    def rename_entry(self, entry_id: str, new_label: str) -> None:
        entries = self.list_entries()
        e = self._find(entries, entry_id)
        if e is None:
            return
        e.label = new_label
        self._save_entries(entries)

    def reset_label(self, entry_id: str) -> None:
        entries = self.list_entries()
        e = self._find(entries, entry_id)
        if e is None:
            return
        e.label = e.auto_label
        self._save_entries(entries)

    def delete_entry(self, entry_id: str) -> None:
        entries = self.list_entries()
        victim = self._find(entries, entry_id)
        if victim is None:
            return
        remaining = [e for e in entries if e.id != entry_id]
        still_referenced = {
            m.blob for e in remaining for m in e.members if m.blob
        }
        for m in victim.members:
            if m.blob and m.blob not in still_referenced:
                try:
                    os.remove(os.path.join(self.blobs_dir, m.blob))
                except OSError:
                    pass
        self._save_entries(remaining)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_library.py -v`
Expected: PASS (all green).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/doc_library/library.py tests/test_doc_library/test_library.py
git commit -m "feat(doc_library): rename/reset/delete with ref-counted blob GC"
```

---

### Task 7: Task-capture allow-list + entry adapter

**Files:**
- Create: `icharlotte_core/doc_library/capture.py`
- Test: `tests/test_doc_library/test_capture.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_doc_library/test_capture.py
from icharlotte_core.doc_library.capture import (
    AUTO_CAPTURE_TASK_IDS, capture_from_task_entry)
from icharlotte_core.doc_library.library import DocumentLibrary


def test_allow_list_contents():
    assert "summarize_documents" in AUTO_CAPTURE_TASK_IDS
    assert "summarize_discovery" in AUTO_CAPTURE_TASK_IDS
    assert "summarize_depositions" in AUTO_CAPTURE_TASK_IDS
    assert "medical_records" in AUTO_CAPTURE_TASK_IDS
    assert "med_chron_analysis" in AUTO_CAPTURE_TASK_IDS
    assert "med_record_extractor" in AUTO_CAPTURE_TASK_IDS
    assert "oppose_motion" not in AUTO_CAPTURE_TASK_IDS
    assert "chat" not in AUTO_CAPTURE_TASK_IDS


def test_non_allowed_task_is_skipped(tmp_path):
    entry = {"task_id": "oppose_motion", "files": [], "settings": {}}
    assert capture_from_task_entry(str(tmp_path), entry) is None
    assert DocumentLibrary(str(tmp_path)).list_entries() == []


def test_allowed_task_captures(tmp_path):
    src = tmp_path / "depo.pdf"
    src.write_bytes(b"bytes")
    from icharlotte_core.doc_library.extract import Extracted
    entry = {"task_id": "summarize_depositions", "files": [str(src)],
             "settings": {"party": "Plaintiff"}}
    result = capture_from_task_entry(
        str(tmp_path), entry,
        extractor=lambda p: Extracted("T", 1, "pdf_native", None))
    assert result is not None
    assert DocumentLibrary(str(tmp_path)).list_entries()[0].label \
        == "Plaintiff's Deposition Transcript"


def test_no_files_is_skipped(tmp_path):
    entry = {"task_id": "summarize_documents", "files": [], "settings": {}}
    assert capture_from_task_entry(str(tmp_path), entry) is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_capture.py -v`
Expected: FAIL — `ModuleNotFoundError: ...capture`.

- [ ] **Step 3: Write minimal implementation**

```python
# icharlotte_core/doc_library/capture.py
"""Bridge from a finished wizard task to the document library.

Connected to TaskTab.task_completed (whose payload already carries task_id,
files, and settings). Best-effort: never raises into the UI.
"""
from __future__ import annotations

import logging
from typing import Callable, Optional

from .extract import extract_any
from .library import DocumentLibrary

logger = logging.getLogger(__name__)

# Source-document-producing tasks (ids from icharlotte_core/ui/wizard/registry.py).
AUTO_CAPTURE_TASK_IDS = {
    "summarize_documents",
    "summarize_discovery",
    "summarize_depositions",
    "medical_records",
    "med_chron_analysis",
    "med_record_extractor",
}


def _metadata_from_settings(settings: dict) -> dict:
    """Pull labeling hints from a task's settings dict, defensively.

    Settings schemas vary per task; we read a few common keys and let
    auto_label fall back to the filename when absent.
    """
    settings = settings or {}
    party = (settings.get("party") or settings.get("audience_party")
             or settings.get("role") or "")
    name = (settings.get("name") or settings.get("deponent")
            or settings.get("patient") or settings.get("client_name") or "")
    return {"party": party, "name": name}


def capture_from_task_entry(case_root: str, entry: dict,
                            extractor: Callable = extract_any) -> Optional[object]:
    if not case_root:
        return None
    task_id = entry.get("task_id")
    if task_id not in AUTO_CAPTURE_TASK_IDS:
        return None
    files = [f for f in (entry.get("files") or []) if f]
    if not files:
        return None
    try:
        lib = DocumentLibrary(case_root)
        return lib.add_entry(task_id, files,
                             _metadata_from_settings(entry.get("settings", {})),
                             extractor=extractor)
    except Exception:  # never break task completion
        logger.exception("doc_library capture failed for task %s", task_id)
        return None
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_capture.py -v`
Expected: PASS (4 passed).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/doc_library/capture.py tests/test_doc_library/test_capture.py
git commit -m "feat(doc_library): task-completion capture bridge + allow-list"
```

---

## Phase 2 — Wizard auto-capture hook

### Task 8: Wire `task_completed` → library capture

**Context:** `TaskTab._on_worker_finished` (`icharlotte_core/ui/wizard/task_tab.py:356`) already emits `task_completed` with `{"task_id", "files", "settings", ...}`. Find where `task_completed` is consumed (the container/main window that records recent tasks). That consumer knows the case root. Add the capture call there so no worker code changes.

**Files:**
- Modify: the `task_completed` consumer — locate via:
  `grep -rn "task_completed" icharlotte_core/`
  (likely `icharlotte_core/ui/wizard/` container or `iCharlotte.py`). The consumer that already has the case path is the target.
- Test: `tests/test_doc_library/test_capture.py` (logic already covered by Task 7; this task is wiring).

- [ ] **Step 1: Locate the consumer and the case root**

Run: `grep -rn "task_completed" icharlotte_core/ iCharlotte.py`
Identify the slot connected to `task_completed`. Confirm how that object gets the case folder path (the wizard already persists under `<case_root>/NOTES/AI OUTPUT/.icharlotte`, so `case_root` is available where `WizardStatePersistence` is constructed — reuse that exact value).

- [ ] **Step 2: Add the capture call in that slot**

In the existing `task_completed` handler (the one that appends to recent tasks), after the current logic add:

```python
from icharlotte_core.doc_library.capture import capture_from_task_entry
# `case_root` here MUST be the same path used to build WizardStatePersistence.
capture_from_task_entry(case_root, entry)
```

Run extraction off the UI thread if the handler is on the GUI thread and large OCR is possible: wrap in a `QThreadPool`/`QRunnable` or a short-lived `QThread`. Minimal version (acceptable for v1, since dedup means most captures are hash-only on re-runs):

```python
from PySide6.QtCore import QThreadPool, QRunnable

class _CaptureJob(QRunnable):
    def __init__(self, case_root, entry):
        super().__init__()
        self._case_root, self._entry = case_root, entry
    def run(self):
        capture_from_task_entry(self._case_root, self._entry)

QThreadPool.globalInstance().start(_CaptureJob(case_root, entry))
```

- [ ] **Step 3: Manual verification**

Run iCharlotte from `C:\geminiterminal2` (the main checkout, not a worktree — the running app ignores worktree edits). Load a case, run **Summarize Documents** on a small PDF. Then check:

Run: `python -c "from icharlotte_core.doc_library.library import DocumentLibrary; import json; print([e.label for e in DocumentLibrary(r'<CASE_ROOT>').list_entries()])"`
Expected: the new entry label is listed, and `<CASE_ROOT>/NOTES/AI OUTPUT/.icharlotte/doc_library/index.json` exists.

- [ ] **Step 4: Commit**

```bash
git add icharlotte_core/ui/wizard/   # whichever file held the consumer
git commit -m "feat(wizard): capture finished-task source text into the document library"
```

---

## Phase 3 — Chat tab integration

### Task 9: "Saved Documents" tree in the chat sidebar

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (sidebar build ~369; `load_case` ~507)
- Test: `tests/test_doc_library/test_chat_library_integration.py`

**Design:** a `QTreeWidget` (`self.library_tree`) inserted after the file-button row (line 369). Top-level items = entries (checkable, `ItemIsEditable` for inline rename — **never** `setItemWidget`, per the drag-reorder gotcha); children = members (checkable). Populated by `self._refresh_library_tree()`, called from `load_case`. The case root comes from `getattr(self.window(), "case_path", None)` (ChatTab already reads `main_win.case_path`).

- [ ] **Step 1: Write the failing test (headless logic via pytestqt)**

```python
# tests/test_doc_library/test_chat_library_integration.py
import os
import pytest
pytest.importorskip("pytestqt")
from PySide6.QtWidgets import QApplication  # noqa: E402

from icharlotte_core.doc_library.library import DocumentLibrary  # noqa: E402
from icharlotte_core.doc_library.extract import Extracted  # noqa: E402


@pytest.fixture(scope="module")
def app():
    return QApplication.instance() or QApplication([])


def _seed_library(case_root):
    src = os.path.join(case_root, "depo.pdf")
    with open(src, "wb") as f:
        f.write(b"bytes")
    lib = DocumentLibrary(case_root)
    lib.add_entry("summarize_depositions", [src], {"party": "Plaintiff"},
                  extractor=lambda p: Extracted("DEPO BODY TEXT", 1, "pdf_native", None))
    return lib


def test_tree_populates_from_library(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)   # injected accessor (see Step 3)
    tab._refresh_library_tree()
    assert tab.library_tree.topLevelItemCount() == 1
    top = tab.library_tree.topLevelItem(0)
    assert top.text(0) == "Plaintiff's Deposition Transcript"
    assert top.childCount() == 1  # one member
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -v`
Expected: FAIL — `AttributeError: 'ChatTab' object has no attribute 'library_tree'`.

- [ ] **Step 3: Implement the tree + refresh**

In `ChatTab.__init__`, after line 369 (`settings_layout.addLayout(file_btn_layout)`), insert:

```python
        # --- Saved Documents (document text library) ---
        from PySide6.QtWidgets import QTreeWidget, QTreeWidgetItem  # local import
        self._QTreeWidgetItem = QTreeWidgetItem
        settings_layout.addWidget(QLabel("Saved Documents:"))
        self.library_tree = QTreeWidget()
        self.library_tree.setHeaderHidden(True)
        self.library_tree.setMinimumHeight(60)
        self.library_tree.setMaximumHeight(300)
        settings_layout.addWidget(self.library_tree)

        lib_btn_layout = QHBoxLayout()
        add_lib_btn = QPushButton("Add to Library…")
        add_lib_btn.clicked.connect(self.add_to_library)
        lib_all_btn = QPushButton("All")
        lib_all_btn.clicked.connect(lambda: self._set_all_library_checks(True))
        lib_none_btn = QPushButton("None")
        lib_none_btn.clicked.connect(lambda: self._set_all_library_checks(False))
        lib_refresh_btn = QPushButton("Refresh")
        lib_refresh_btn.clicked.connect(self._refresh_library_tree)
        for b in (add_lib_btn, lib_all_btn, lib_none_btn, lib_refresh_btn):
            lib_btn_layout.addWidget(b)
        settings_layout.addLayout(lib_btn_layout)
        self.library_selected_label = QLabel("Selected: 0 docs · ~0 tokens")
        settings_layout.addWidget(self.library_selected_label)
```

Add these methods to `ChatTab`:

```python
    def _library(self):
        """Return a DocumentLibrary for the current case, or None."""
        from icharlotte_core.doc_library.library import DocumentLibrary
        root = getattr(self, "_case_root_for_library", None) \
            or getattr(self.window(), "case_path", None)
        return DocumentLibrary(root) if root else None

    def _refresh_library_tree(self):
        from PySide6.QtCore import Qt
        self.library_tree.clear()
        lib = self._library()
        if lib is None:
            self._update_library_selected_label()
            return
        Item = self._QTreeWidgetItem
        try:
            entries = lib.list_entries()
        except Exception:
            entries = []
        for e in entries:
            top = Item([e.label])
            top.setData(0, Qt.ItemDataRole.UserRole, {"kind": "entry", "id": e.id})
            top.setFlags(top.flags() | Qt.ItemFlag.ItemIsUserCheckable
                         | Qt.ItemFlag.ItemIsEditable
                         | Qt.ItemFlag.ItemIsAutoTristate)
            top.setCheckState(0, Qt.CheckState.Unchecked)
            for m in e.members:
                child = Item([m.source_name + (" [extract failed]" if m.error else "")])
                child.setData(0, Qt.ItemDataRole.UserRole,
                              {"kind": "member", "blob": m.blob,
                               "name": m.source_name, "tokens": m.est_tokens})
                child.setFlags(child.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                child.setCheckState(0, Qt.CheckState.Unchecked)
                top.addChild(child)
            self.library_tree.addTopLevelItem(top)
        self.library_tree.expandAll()
        self.library_tree.itemChanged.connect(self._on_library_item_changed)
        self._update_library_selected_label()

    def _set_all_library_checks(self, checked: bool):
        from PySide6.QtCore import Qt
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for i in range(self.library_tree.topLevelItemCount()):
            self.library_tree.topLevelItem(i).setCheckState(0, state)

    def _on_library_item_changed(self, item, column):
        self._update_library_selected_label()

    def _iter_checked_library_members(self):
        from PySide6.QtCore import Qt
        for i in range(self.library_tree.topLevelItemCount()):
            top = self.library_tree.topLevelItem(i)
            for j in range(top.childCount()):
                child = top.child(j)
                if child.checkState(0) == Qt.CheckState.Checked:
                    yield child.data(0, Qt.ItemDataRole.UserRole)

    def _update_library_selected_label(self):
        members = list(self._iter_checked_library_members())
        toks = sum(int(m.get("tokens", 0)) for m in members)
        self.library_selected_label.setText(
            f"Selected: {len(members)} docs · ~{toks} tokens")
```

> If `ChatTab()` can't be constructed bare in tests (needs args), the executor should match the existing test pattern in `tests/` for instantiating ChatTab and inject `_case_root_for_library`.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/tabs.py tests/test_doc_library/test_chat_library_integration.py
git commit -m "feat(chat): Saved Documents tree populated from the document library"
```

---

### Task 10: Manual "Add to Library…"

**Files:**
- Modify: `icharlotte_core/ui/tabs.py`
- Test: `tests/test_doc_library/test_chat_library_integration.py`

- [ ] **Step 1: Write the failing test**

```python
# append to test_chat_library_integration.py
def test_add_to_library_adds_entry(app, tmp_path, monkeypatch):
    from icharlotte_core.ui.tabs import ChatTab
    src = tmp_path / "Traffic Collision Report.pdf"
    src.write_bytes(b"bytes")
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    # Stub the file dialog and the extractor used by manual add.
    monkeypatch.setattr(tab, "_pick_files_for_library", lambda: [str(src)])
    from icharlotte_core.doc_library import extract as ex
    monkeypatch.setattr(
        "icharlotte_core.doc_library.library._default_extractor",
        lambda p: ex.Extracted("ACCIDENT TEXT", 1, "pdf_native", None))
    tab.add_to_library()
    labels = [e.label for e in tab._library().list_entries()]
    assert "Traffic Collision Report" in labels
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -k add_to_library -v`
Expected: FAIL — `AttributeError: add_to_library`.

- [ ] **Step 3: Implement manual add**

Add to `ChatTab`:

```python
    def _pick_files_for_library(self):
        from PySide6.QtWidgets import QFileDialog
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Add documents to library", "",
            "Documents (*.pdf *.docx *.doc *.txt *.msg);;All files (*.*)")
        return list(paths)

    def add_to_library(self):
        lib = self._library()
        if lib is None:
            from PySide6.QtWidgets import QMessageBox
            QMessageBox.information(self, "No Case Loaded",
                                    "Load a case before adding to its library.")
            return
        paths = self._pick_files_for_library()
        if not paths:
            return
        # Extraction may be slow (OCR) -> background thread.
        from PySide6.QtCore import QThread, Signal, QObject

        class _Worker(QObject):
            done = Signal()
            def __init__(self, lib, paths):
                super().__init__()
                self._lib, self._paths = lib, paths
            def run(self):
                try:
                    self._lib.add_entry("manual", self._paths, {})
                finally:
                    self.done.emit()

        self._lib_thread = QThread()
        self._lib_worker = _Worker(lib, paths)
        self._lib_worker.moveToThread(self._lib_thread)
        self._lib_thread.started.connect(self._lib_worker.run)
        self._lib_worker.done.connect(self._lib_thread.quit)
        self._lib_worker.done.connect(self._refresh_library_tree)
        self._lib_thread.start()
```

> The test calls `add_to_library` synchronously; for deterministic CI, the executor may special-case "run inline when no event loop is spinning" or have the test wait on the thread via `qtbot.waitUntil`. Prefer `qtbot.waitUntil(lambda: tab._library().list_entries())`.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -k add_to_library -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/tabs.py tests/test_doc_library/test_chat_library_integration.py
git commit -m "feat(chat): manual Add to Library with background extraction"
```

---

### Task 11: Feed checked library text into the LLM

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (`read_files_content` ~1235; `send_message` ~1422)
- Test: `tests/test_doc_library/test_chat_library_integration.py`

**Design:** a new `ChatTab.read_library_content()` returns the same `"--- FILE: name ---\n<text>"` framing for checked members, loading blob text via `lib.get_member_text(blob)` (no extraction). `send_message` concatenates it onto `file_content`.

- [ ] **Step 1: Write the failing test**

```python
# append to test_chat_library_integration.py
def test_read_library_content_includes_checked_text(app, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.tabs import ChatTab
    _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    tab.library_tree.topLevelItem(0).setCheckState(0, Qt.CheckState.Checked)
    out = tab.read_library_content()
    assert "DEPO BODY TEXT" in out
    assert "--- FILE:" in out


def test_read_library_content_empty_when_nothing_checked(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    assert tab.read_library_content() == ""
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -k read_library_content -v`
Expected: FAIL — `AttributeError: read_library_content`.

- [ ] **Step 3: Implement**

Add to `ChatTab`:

```python
    def read_library_content(self) -> str:
        lib = self._library()
        if lib is None:
            return ""
        content = ""
        for m in self._iter_checked_library_members():
            blob = m.get("blob")
            if not blob:
                continue
            text = lib.get_member_text(blob)
            if not text:
                content += f"\n--- FILE: {m.get('name','document')} ---\n[saved text unavailable]\n"
                continue
            content += f"\n--- FILE: {m.get('name','document')} ---\n{text}\n"
        return content
```

In `send_message`, change line 1422 from:

```python
        file_content = self.read_files_content()
```
to:
```python
        file_content = self.read_files_content() + self.read_library_content()
```

Also update the checked-count guard (line 1404-1410) so a library-only send isn't blocked:

```python
        lib_checked = len(list(self._iter_checked_library_members()))
        if not user_text and checked_count == 0 and lib_checked == 0:
            return
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -v`
Expected: PASS (all).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/tabs.py tests/test_doc_library/test_chat_library_integration.py
git commit -m "feat(chat): include checked library text in the LLM request"
```

---

### Task 12: Pre-send context-budget warning

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (`send_message`, after `file_content` is built ~1422)
- Test: `tests/test_doc_library/test_chat_library_integration.py`

**Design:** a pure helper `ChatTab._library_budget_warning(file_content, history_tokens)` returns an optional warning string when `est(file_content) + history + reserve` exceeds the model's context limit (via `TokenCounter.get_context_limit`). `send_message` shows a `QMessageBox` and bails if the user cancels. The helper is unit-tested without a dialog.

- [ ] **Step 1: Write the failing test**

```python
# append to test_chat_library_integration.py
def test_budget_warning_fires_when_over_limit(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    # Force a tiny limit by selecting a model with small window via stub.
    tab._context_limit_for_test = 100
    huge = "x" * 100_000  # ~25k tokens
    warn = tab._library_budget_warning(huge, history_tokens=0)
    assert warn is not None
    assert "context" in warn.lower()


def test_budget_warning_silent_when_under_limit(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._context_limit_for_test = 1_000_000
    assert tab._library_budget_warning("small text", history_tokens=0) is None
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -k budget -v`
Expected: FAIL — `AttributeError: _library_budget_warning`.

- [ ] **Step 3: Implement**

Add to `ChatTab`:

```python
    def _context_limit(self) -> int:
        forced = getattr(self, "_context_limit_for_test", None)
        if forced is not None:
            return forced
        from icharlotte_core.chat.token_counter import TokenCounter
        return TokenCounter.get_context_limit(
            self.model_combo.currentText(), self.provider_combo.currentText())

    def _library_budget_warning(self, file_content: str, history_tokens: int) -> str:
        from icharlotte_core.chat.token_counter import TokenCounter
        reserve = 16384
        used = (TokenCounter.estimate_tokens(file_content) + history_tokens + reserve)
        limit = self._context_limit()
        if used <= limit:
            return None
        return (f"The selected documents (~{TokenCounter.format_token_count(used)} "
                f"tokens) exceed this model's context window "
                f"(~{TokenCounter.format_token_count(limit)}). The request may be "
                f"truncated or rejected. Deselect some Saved Documents, or proceed "
                f"anyway?")
```

In `send_message`, immediately after `file_content = ...` (Task 11 line), add:

```python
        warn = self._library_budget_warning(file_content, history_tokens=0)
        if warn:
            from PySide6.QtWidgets import QMessageBox
            reply = QMessageBox.warning(
                self, "Context Budget", warn,
                QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.Cancel)
            if reply == QMessageBox.StandardButton.Cancel:
                self.send_btn.setEnabled(True)
                self.stop_btn.setEnabled(False)
                return
```

> `history_tokens=0` keeps v1 simple; a later refinement can sum the conversation history tokens. Stated here so it isn't mistaken for an oversight.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -k budget -v`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/tabs.py tests/test_doc_library/test_chat_library_integration.py
git commit -m "feat(chat): pre-send context-budget warning for saved-document selections"
```

---

### Task 13: Inline rename + selection memory

**Files:**
- Modify: `icharlotte_core/ui/tabs.py` (`_on_library_item_changed`; `load_case` ~507)
- Test: `tests/test_doc_library/test_chat_library_integration.py`

**Design:** (a) editing a top-level item's text calls `lib.rename_entry`; (b) checked entry ids persist in the chat settings block and restore on `load_case`.

- [ ] **Step 1: Write the failing test**

```python
# append to test_chat_library_integration.py
def test_inline_rename_persists(app, tmp_path):
    from PySide6.QtCore import Qt
    from icharlotte_core.ui.tabs import ChatTab
    _seed_library(str(tmp_path))
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    top = tab.library_tree.topLevelItem(0)
    top.setText(0, "Renamed Depo")
    tab._on_library_item_changed(top, 0)
    labels = [e.label for e in tab._library().list_entries()]
    assert "Renamed Depo" in labels


def test_selection_roundtrip(app, tmp_path):
    from icharlotte_core.ui.tabs import ChatTab
    lib = _seed_library(str(tmp_path))
    entry_id = lib.list_entries()[0].id
    tab = ChatTab()
    tab._case_root_for_library = str(tmp_path)
    tab._refresh_library_tree()
    from PySide6.QtCore import Qt
    tab.library_tree.topLevelItem(0).setCheckState(0, Qt.CheckState.Checked)
    saved = tab._collect_checked_entry_ids()
    assert entry_id in saved
    # New tab restores from the same id list
    tab2 = ChatTab()
    tab2._case_root_for_library = str(tmp_path)
    tab2._refresh_library_tree()
    tab2._restore_checked_entry_ids(saved)
    assert tab2.library_tree.topLevelItem(0).checkState(0) == Qt.CheckState.Checked
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -k "rename or roundtrip" -v`
Expected: FAIL — `AttributeError: _collect_checked_entry_ids`.

- [ ] **Step 3: Implement**

Replace `_on_library_item_changed` and add helpers:

```python
    def _on_library_item_changed(self, item, column):
        from PySide6.QtCore import Qt
        data = item.data(0, Qt.ItemDataRole.UserRole) or {}
        # Inline rename of an entry (top-level item, text changed).
        if data.get("kind") == "entry" and column == 0:
            lib = self._library()
            new_label = item.text(0).strip()
            if lib and new_label:
                try:
                    lib.rename_entry(data["id"], new_label)
                except Exception:
                    pass
        self._update_library_selected_label()
        self._persist_library_selection()

    def _collect_checked_entry_ids(self):
        from PySide6.QtCore import Qt
        ids = []
        for i in range(self.library_tree.topLevelItemCount()):
            top = self.library_tree.topLevelItem(i)
            if top.checkState(0) in (Qt.CheckState.Checked,
                                     Qt.CheckState.PartiallyChecked):
                data = top.data(0, Qt.ItemDataRole.UserRole) or {}
                if data.get("id"):
                    ids.append(data["id"])
        return ids

    def _restore_checked_entry_ids(self, ids):
        from PySide6.QtCore import Qt
        wanted = set(ids or [])
        for i in range(self.library_tree.topLevelItemCount()):
            top = self.library_tree.topLevelItem(i)
            data = top.data(0, Qt.ItemDataRole.UserRole) or {}
            if data.get("id") in wanted:
                top.setCheckState(0, Qt.CheckState.Checked)

    def _persist_library_selection(self):
        if not getattr(self, "persistence", None):
            return
        try:
            self.persistence.set_setting(
                "library_selected_ids", self._collect_checked_entry_ids())
        except Exception:
            pass
```

In `load_case`, after the library tree is refreshed, add:

```python
        self._refresh_library_tree()
        if getattr(self, "persistence", None):
            try:
                saved = self.persistence.get_setting("library_selected_ids", [])
            except Exception:
                saved = []
            self._restore_checked_entry_ids(saved)
```

> If `ChatPersistence` lacks `get_setting`/`set_setting`, the executor adds thin wrappers over its existing settings dict (the chat JSON already has a `"settings"` block — see `icharlotte_core/chat/persistence.py`). Match the existing accessor naming there.

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_doc_library/test_chat_library_integration.py -v`
Expected: PASS (all).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/ui/tabs.py icharlotte_core/chat/persistence.py tests/test_doc_library/test_chat_library_integration.py
git commit -m "feat(chat): inline rename + persistent saved-document selection"
```

---

### Task 14: Full-suite regression + manual smoke

- [ ] **Step 1: Run the doc_library suite**

Run: `python -m pytest tests/test_doc_library/ -v`
Expected: all green.

- [ ] **Step 2: Run the broader chat/wizard suites for regressions**

Run: `python -m pytest tests/ -k "chat or wizard or tabs" -q`
Expected: no new failures (note any pre-existing failures recorded in memory, e.g. the flaky `use_all_text_check` test, are not regressions).

- [ ] **Step 3: Manual smoke (from `C:\geminiterminal2`, the running checkout)**

1. Load a case, run **Summarize Deposition** on a small transcript → confirm a "…Deposition Transcript" entry appears under **Saved Documents** in Chat (Refresh if needed).
2. Use **Add to Library…** to add a traffic-collision PDF → entry appears named from the filename.
3. Check the deposition entry, ask "what did the plaintiff testify about the intersection?" → answer is grounded in the cached text with **no re-upload**.
4. Rename an entry inline; reopen the case → rename and selection persisted.
5. Check a very large transcript with a small-window model → budget warning appears.

- [ ] **Step 4: Commit (if any fixes were needed)**

```bash
git add -A
git commit -m "test(doc_library): full-suite regression pass + smoke fixes"
```

---

## Self-Review notes (author)

- **Spec coverage:** raw-text storage (Task 5/11), per-run-expandable entries (Task 9 tree), auto-labels + rename/reset (Tasks 3/6/13), tasks+manual population (Tasks 7/8/10), case-traveling storage (Task 4), dedup/refcount (Tasks 5/6), budget warning (Task 12), selection memory (Task 13), error handling (Tasks 2/5 error members; Task 4 index recovery), tests throughout. ✅
- **Known executor confirmations (not placeholders, but real lookups):** Task 8 must locate the `task_completed` consumer that owns `case_root`; Task 9/13 must match the project's existing ChatTab test-construction pattern and `ChatPersistence` setting accessors. These are explicitly flagged with exact grep commands and fallbacks.
- **Type consistency:** `Extracted` fields (`text/page_count/extract_method/error`) used identically in Tasks 2/5/7/9/10; `MemberFile`/`LibraryEntry` shapes consistent across Tasks 1/4/5/6; `DocumentLibrary` method names (`add_entry/list_entries/get_member_text/rename_entry/reset_label/delete_entry`) stable across all callers.
