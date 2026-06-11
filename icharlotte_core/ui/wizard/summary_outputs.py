"""Discovery helpers for task-specific summary browser tabs."""
from __future__ import annotations

from dataclasses import dataclass
import os
from typing import Iterable

from .persistence import WizardStatePersistence


SUMMARY_BROWSER_ACTIONS: dict[str, str] = {
    "open_summarize_documents_outputs": "summarize_documents",
    "open_summarize_discovery_outputs": "summarize_discovery",
    "open_summarize_depositions_outputs": "summarize_depositions",
    "open_medical_records_outputs": "medical_records",
}

SUMMARY_BROWSER_TITLES: dict[str, str] = {
    "summarize_documents": "Summary Browser - Documents",
    "summarize_discovery": "Summary Browser - Discovery",
    "summarize_depositions": "Summary Browser - Depositions",
    "medical_records": "Summary Browser - Medical Records",
}


@dataclass(frozen=True)
class SummaryOutputEntry:
    path: str
    task_id: str
    source: str
    mtime: float

    @property
    def name(self) -> str:
        return os.path.basename(self.path)


def task_id_for_summary_action(action_id: str) -> str | None:
    return SUMMARY_BROWSER_ACTIONS.get(action_id)


def summary_browser_title(task_id: str) -> str:
    return SUMMARY_BROWSER_TITLES.get(task_id, "Summary Browser")


def discover_summary_outputs(case_path: str, task_id: str) -> list[SummaryOutputEntry]:
    """Return task-specific summary outputs, newest first.

    Wizard history supplies exact task provenance. Legacy/manual Advanced Mode
    outputs are included through filename conventions in the case AI OUTPUT
    folder.
    """
    if task_id not in SUMMARY_BROWSER_TITLES or not case_path:
        return []

    discovered: dict[str, SummaryOutputEntry] = {}
    wizard_owners = _wizard_output_owner_map(case_path)

    def add(path: str, source: str) -> None:
        if not path:
            return
        abs_path = _resolve_output_path(case_path, path)
        if not abs_path or not os.path.isfile(abs_path):
            return
        if not abs_path.lower().endswith(".docx"):
            return
        try:
            mtime = os.path.getmtime(abs_path)
        except OSError:
            return
        key = os.path.normcase(os.path.abspath(abs_path))
        existing = discovered.get(key)
        if existing is not None and existing.source == "Wizard":
            return
        discovered[key] = SummaryOutputEntry(
            path=os.path.normpath(abs_path),
            task_id=task_id,
            source=source,
            mtime=mtime,
        )

    for path in _wizard_output_paths(case_path, task_id):
        add(path, "Wizard")

    for output_dir in _ai_output_dirs(case_path):
        try:
            names = os.listdir(output_dir)
        except OSError:
            continue
        for name in names:
            path = os.path.join(output_dir, name)
            if not os.path.isfile(path):
                continue
            if _matches_legacy_task_pattern(task_id, name):
                key = os.path.normcase(os.path.abspath(path))
                owner = wizard_owners.get(key)
                if owner is not None and owner != task_id:
                    continue
                add(path, "Legacy")

    return sorted(discovered.values(), key=lambda entry: entry.mtime, reverse=True)


def _wizard_output_paths(case_path: str, task_id: str) -> Iterable[str]:
    return [
        path
        for owner_task_id, path in _iter_wizard_output_paths(case_path)
        if owner_task_id == task_id
    ]


def _wizard_output_owner_map(case_path: str) -> dict[str, str]:
    owners: dict[str, str] = {}
    for task_id, path in _iter_wizard_output_paths(case_path):
        abs_path = _resolve_output_path(case_path, path)
        if not abs_path:
            continue
        owners[os.path.normcase(os.path.abspath(abs_path))] = task_id
    return owners


def _iter_wizard_output_paths(case_path: str) -> Iterable[tuple[str, str]]:
    try:
        persistence = WizardStatePersistence(case_path)
        entries = persistence.get_recent_tasks() + persistence.get_open_tabs()
    except Exception:
        return []

    paths: list[tuple[str, str]] = []
    for entry in entries:
        if not isinstance(entry, dict):
            continue
        task_id = entry.get("task_id")
        if task_id not in SUMMARY_BROWSER_TITLES:
            continue
        output_paths = entry.get("output_paths")
        if isinstance(output_paths, list):
            paths.extend((task_id, str(path)) for path in output_paths if path)
        output_path = entry.get("output_path")
        if output_path:
            paths.append((task_id, str(output_path)))
    return paths


def _resolve_output_path(case_path: str, path: str) -> str:
    if os.path.isabs(path):
        return path
    return os.path.join(case_path, path)


def _ai_output_dirs(case_path: str) -> list[str]:
    candidates = [
        os.path.join(case_path, "NOTES", "AI OUTPUT"),
        os.path.join(case_path, "NOTES", "AI Output"),
    ]
    seen: set[str] = set()
    dirs: list[str] = []
    for path in candidates:
        key = os.path.normcase(os.path.abspath(path))
        if key in seen or not os.path.isdir(path):
            continue
        seen.add(key)
        dirs.append(path)
    return dirs


def _matches_legacy_task_pattern(task_id: str, filename: str) -> bool:
    lower = filename.lower()
    if not lower.endswith(".docx"):
        return False
    if task_id == "summarize_discovery":
        return lower.startswith("discovery_responses_")
    if task_id == "summarize_depositions":
        return (
            lower.startswith("deposition of ")
            or lower.startswith("deposition_summaries")
            or (lower.endswith(" - summary.docx") and ("depo" in lower or "deposition" in lower))
        )
    if task_id == "summarize_documents":
        return lower.startswith("ai_output")
    if task_id == "medical_records":
        return lower.startswith("med_record_")
    return False
