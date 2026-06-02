"""Bridge from a finished wizard task to the document library.

Connected to TaskTab.task_completed (whose payload already carries task_id,
files, and settings). Best-effort: never raises into the UI.
"""
from __future__ import annotations

import logging
from typing import Callable, Optional

from .extract import extract_any
from .library import DocumentLibrary
from .models import LibraryEntry

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
                            extractor: Callable = extract_any) -> Optional[LibraryEntry]:
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
        logger.exception("doc_library capture failed for task %s (%d files)", task_id, len(files))
        return None
