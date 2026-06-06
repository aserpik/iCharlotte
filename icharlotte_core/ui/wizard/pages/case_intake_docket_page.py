"""Case Intake & Docket wizard task."""
from __future__ import annotations

import glob
import os
import re
from typing import Any


REVIEW_FIELDS = [
    "case_number",
    "venue_county",
    "case_name",
    "filing_date",
    "plaintiffs",
    "defendants",
    "client_name",
    "client_email",
    "plaintiff_counsel",
    "causes_of_action",
]

LIST_FIELDS = {"plaintiffs", "defendants", "causes_of_action"}


def _case_manager():
    from Scripts.case_data_manager import CaseDataManager

    return CaseDataManager()


def normalize_review_value(key: str, value: Any) -> Any:
    if key in LIST_FIELDS:
        if isinstance(value, list):
            return [str(v).strip() for v in value if str(v).strip()]
        text = "" if value is None else str(value)
        parts = re.split(r"[\n;]+", text)
        return [p.strip() for p in parts if p.strip()]
    if value is None:
        return ""
    return str(value).strip()


def load_case_metadata(file_number: str, manager=None) -> dict[str, Any]:
    mgr = manager or _case_manager()
    metadata: dict[str, Any] = {}
    for key in REVIEW_FIELDS:
        metadata[key] = normalize_review_value(key, mgr.get_value(file_number, key))
    return metadata


def save_reviewed_metadata(file_number: str, metadata: dict[str, Any], manager=None) -> None:
    mgr = manager or _case_manager()
    for key in REVIEW_FIELDS:
        if key not in metadata:
            continue
        value = normalize_review_value(key, metadata.get(key, ""))
        mgr.save_variable(
            file_number,
            key,
            value,
            source="wizard_case_intake",
            extra_tags=["Meta Data"],
        )


def find_latest_docket_pdf(case_path: str) -> str:
    out_dir = os.path.join(case_path, "NOTES", "AI OUTPUT")
    candidates = glob.glob(os.path.join(out_dir, "Docket_*.pdf"))
    candidates = [p for p in candidates if os.path.isfile(p)]
    if not candidates:
        return ""
    return max(candidates, key=os.path.getmtime)


def find_variables_docx(case_path: str) -> str:
    path = os.path.join(case_path, "NOTES", "AI OUTPUT", "variables.docx")
    return path if os.path.isfile(path) else ""


def find_complaint_candidate(case_path: str) -> str:
    folders = ["PLEADINGS", "PLEADING"]
    patterns = ["*complaint*.pdf", "*complaint*.docx", "*s&c*.pdf", "*summons*.pdf"]
    found: list[str] = []
    for folder in folders:
        base = os.path.join(case_path, folder)
        if not os.path.isdir(base):
            continue
        for pattern in patterns:
            found.extend(glob.glob(os.path.join(base, "**", pattern), recursive=True))
    found = [p for p in found if os.path.isfile(p)]
    if not found:
        return ""
    return max(found, key=os.path.getmtime)


def build_output_summary(
    case_path: str,
    file_number: str,
    manager=None,
    master_db=None,
    recent_lines: list[str] | None = None,
    success: bool = True,
) -> dict[str, Any]:
    mgr = manager or _case_manager()
    docket_pdf = find_latest_docket_pdf(case_path)
    variables_docx = find_variables_docx(case_path)
    case_row = None
    if master_db is not None:
        try:
            case_row = master_db.get_case(file_number)
        except Exception:
            case_row = None
    trial_date = ""
    if case_row:
        trial_date = case_row.get("trial_date") or ""
    if not trial_date:
        trial_date = mgr.get_value(file_number, "trial_date") or ""
    other_hearings = mgr.get_value(file_number, "other_hearings") or ""
    procedural_history = mgr.get_value(file_number, "procedural_history") or ""
    lines = list(recent_lines or [])

    if not success:
        status = "Docket processing failed. Review the final log lines below."
    elif docket_pdf:
        status = "Docket processing finished and a docket PDF was found."
    else:
        status = (
            "Docket processing finished. No docket PDF was found. "
            "The venue may be unsupported or the scraper may have skipped the download."
        )

    return {
        "success": bool(success),
        "status": status,
        "docket_pdf": docket_pdf,
        "variables_docx": variables_docx,
        "trial_date": str(trial_date or ""),
        "other_hearings": str(other_hearings or ""),
        "procedural_history": str(procedural_history or ""),
        "recent_lines": lines,
    }
