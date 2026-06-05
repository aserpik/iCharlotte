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
        if name:
            return f"Deposition of {name}"
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
