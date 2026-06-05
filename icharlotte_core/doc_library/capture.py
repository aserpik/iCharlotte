"""Bridge from a finished wizard task to the document library.

Connected to TaskTab.task_completed (whose payload already carries task_id,
files, and settings). Best-effort: never raises into the UI.
"""
from __future__ import annotations

import logging
import os
import re
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


def _titlecase_caption(name: str) -> str:
    """All-caps deposition captions (e.g. 'JOE SMITH') -> 'Joe Smith' for display.
    Mixed-case names are left untouched."""
    name = (name or "").strip()
    return name.title() if name and name == name.upper() else name


def _deponent_name_from_sessions(transcript_path: str) -> str:
    """Authoritative deponent name from the kept Phase-1 deposition session JSON,
    matched by input transcript path (most recently written session wins)."""
    try:
        import glob
        import json
        from ..deposition.session_manager import SESSION_DIR
        target = os.path.normcase(os.path.abspath(transcript_path))
    except Exception:
        return ""
    best = None  # (mtime, name)
    try:
        session_files = glob.glob(os.path.join(str(SESSION_DIR), "*.json"))
    except Exception:
        return ""
    for sp in session_files:
        try:
            with open(sp, "r", encoding="utf-8") as fh:
                data = json.load(fh)
            ip = data.get("input_path")
            nm = (data.get("deponent_name") or "").strip()
            if not (ip and nm):
                continue
            if os.path.normcase(os.path.abspath(ip)) != target:
                continue
            mt = os.path.getmtime(sp)
        except Exception:
            continue
        if best is None or mt > best[0]:
            best = (mt, nm)
    return best[1] if best else ""


_DEPO_OUTPUT_RE = re.compile(r"^Deposition of (.+?)(?:\s+v\.?\d+)?\.docx$", re.IGNORECASE)


def _deponent_name_from_outputs(entry: dict) -> str:
    """Fallback: parse the deponent name from a 'Deposition of X.docx' output."""
    outs = list(entry.get("output_paths") or [])
    if entry.get("output_path"):
        outs.append(entry["output_path"])
    for p in outs:
        m = _DEPO_OUTPUT_RE.match(os.path.basename(p or ""))
        if m:
            return m.group(1).strip()
    return ""


def _deponent_name_for_entry(entry: dict) -> str:
    """Best-effort deponent name for a finished summarize_depositions task.

    Primary source is the Phase-1 session JSON (authoritative; kept on disk);
    falls back to the 'Deposition of X.docx' output filename. Returns "" if
    neither yields a name.
    """
    files = [f for f in (entry.get("files") or []) if f]
    name = _deponent_name_from_sessions(files[0]) if files else ""
    if not name:
        name = _deponent_name_from_outputs(entry)
    return _titlecase_caption(name)


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
        metadata = _metadata_from_settings(entry.get("settings", {}))
        # Depositions: the settings dict carries no deponent name, so derive it
        # (session JSON / output filename) → label becomes "Deposition of <name>".
        if task_id == "summarize_depositions" and not metadata.get("name"):
            deponent = _deponent_name_for_entry(entry)
            if deponent:
                metadata["name"] = deponent
        return lib.add_entry(task_id, files, metadata, extractor=extractor)
    except Exception:  # never break task completion
        logger.exception("doc_library capture failed for task %s (%d files)", task_id, len(files))
        return None
