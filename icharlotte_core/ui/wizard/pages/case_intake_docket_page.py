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

COMPLAINT_USEFUL_TERMS = [
    "complaint",
    "s&c",
    "fac",
    "sac",
    "tac",
    "3ac",
    "4ac",
    "amended",
    "cmpl",
    "cmp",
    "pleading",
    "pld",
    "summons",
]

COMPLAINT_REJECT_TERMS = [
    "motion",
    "response",
    "demurrer",
    "answer",
    "reply",
    "cross",
    "notice",
    "proof",
    "pos",
    "svc",
    "service",
]

COMPLAINT_EXTENSIONS = (".pdf", ".docx")


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


def _has_name_term(name: str, term: str) -> bool:
    if term in {"fac", "sac", "tac", "3ac", "4ac", "cmpl", "cmp", "pld", "pos", "svc"}:
        return re.search(rf"(?<![a-z0-9]){re.escape(term)}(?![a-z0-9])", name) is not None
    return term in name


def _complaint_candidate_score(name: str) -> int:
    if _has_name_term(name, "4ac") or "4th amended" in name or "fourth amended" in name:
        return 70
    if (
        _has_name_term(name, "tac")
        or _has_name_term(name, "3ac")
        or "3rd amended" in name
        or "third amended" in name
    ):
        return 60
    if _has_name_term(name, "sac") or "2nd amended" in name or "second amended" in name:
        return 50
    if _has_name_term(name, "fac") or "1st amended" in name or "first amended" in name:
        return 40
    if "amended" in name:
        return 30
    if (
        "complaint" in name
        or "s&c" in name
        or _has_name_term(name, "cmpl")
        or _has_name_term(name, "cmp")
        or "pleading" in name
        or _has_name_term(name, "pld")
    ):
        return 20
    if "summons" in name:
        return 10
    return 0


def find_complaint_candidate(case_path: str) -> str:
    folders = ["PLEADINGS", "PLEADING"]
    candidates: list[tuple[int, float, str]] = []
    for folder in folders:
        base = os.path.join(case_path, folder)
        if not os.path.isdir(base):
            continue
        for root, dirs, files in os.walk(base):
            dirs.sort()
            for filename in sorted(files):
                if not filename.lower().endswith(COMPLAINT_EXTENSIONS):
                    continue
                path = os.path.join(root, filename)
                if not os.path.isfile(path):
                    continue
                name = filename.lower()
                if not any(_has_name_term(name, term) for term in COMPLAINT_USEFUL_TERMS):
                    continue
                if any(_has_name_term(name, term) for term in COMPLAINT_REJECT_TERMS):
                    continue
                score = _complaint_candidate_score(name)
                if score <= 0:
                    continue
                candidates.append((score, os.path.getmtime(path), path))
    if not candidates:
        return ""
    return max(candidates, key=lambda candidate: (candidate[0], candidate[1]))[2]


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

    warning = ""
    if not success:
        state = "failed"
        warning = "Docket processing failed. Review the final log lines below."
        status = "Docket processing failed. Review the final log lines below."
    elif docket_pdf:
        state = "success"
        status = "Docket processing finished and a docket PDF was found."
    else:
        state = "partial"
        warning = (
            "No docket PDF was found. "
            "The venue may be unsupported or the scraper may have skipped the download."
        )
        status = (
            "Docket processing finished. No docket PDF was found. "
            "The venue may be unsupported or the scraper may have skipped the download."
        )

    return {
        "success": bool(success),
        "state": state,
        "warning": warning,
        "status": status,
        "docket_pdf": docket_pdf,
        "variables_docx": variables_docx,
        "trial_date": str(trial_date or ""),
        "other_hearings": str(other_hearings or ""),
        "procedural_history": str(procedural_history or ""),
        "recent_lines": lines,
    }
