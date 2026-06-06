"""Case Intake & Docket wizard task."""
from __future__ import annotations

import glob
import os
import re
from typing import Any

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.ui.wizard import theme


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
PLACEHOLDER_VALUES = {"none", "null", "n/a", "na"}


def _review_label(key: str) -> str:
    return key.replace("_", " ").title()


def _case_manager():
    from Scripts.case_data_manager import CaseDataManager

    return CaseDataManager()


def normalize_review_value(key: str, value: Any) -> Any:
    if key in LIST_FIELDS:
        if isinstance(value, list):
            return [
                text
                for v in value
                if (text := str(v).strip()) and text.lower() not in PLACEHOLDER_VALUES
            ]
        text = "" if value is None else str(value)
        parts = re.split(r"[\n;]+", text)
        return [
            part
            for p in parts
            if (part := p.strip()) and part.lower() not in PLACEHOLDER_VALUES
        ]
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in PLACEHOLDER_VALUES:
        return ""
    return text


class CaseIntakeSettingsPage(QWidget):
    """Initial Case Intake page that runs the complaint agent."""

    run_complaint_requested = Signal()

    def __init__(self, file_number: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.file_number = (file_number or "").strip()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.page_title("Case Intake & Docket"))
        layout.addWidget(theme.helper_text(
            "Run the complaint intake agent, review the case metadata, then run docket."
        ))

        display_file_number = self.file_number or "No file number selected"
        self.file_number_label = QLabel(f"File number: {display_file_number}")
        layout.addWidget(self.file_number_label)
        layout.addStretch(1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.run_complaint_btn = theme.primary_button("Run Complaint Intake")
        self.run_complaint_btn.setEnabled(bool(self.file_number))
        self.run_complaint_btn.clicked.connect(self.run_complaint_requested.emit)
        btn_row.addWidget(self.run_complaint_btn)
        layout.addLayout(btn_row)

    def to_dict(self) -> dict:
        return {}

    def from_dict(self, data: dict | None) -> None:
        return None


class CaseMetadataReviewPage(QWidget):
    """Review complaint metadata before running the docket agent."""

    run_docket_requested = Signal(dict)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._complaint_file = ""
        self._field_widgets: dict[str, QLineEdit | QPlainTextEdit] = {}

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.page_title("Review Case Metadata"))
        layout.addWidget(theme.helper_text(
            "Confirm the extracted case details before the docket agent runs."
        ))

        self.complaint_label = theme.caption("Complaint: Not selected")
        layout.addWidget(self.complaint_label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        form_host = QWidget()
        form = QFormLayout(form_host)
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(theme.SPACE_MD)

        for key in REVIEW_FIELDS:
            if key in LIST_FIELDS:
                widget = QPlainTextEdit()
                widget.setFixedHeight(72)
                widget.textChanged.connect(self._update_run_docket_enabled)
            else:
                widget = QLineEdit()
                widget.textChanged.connect(self._update_run_docket_enabled)
            widget.setObjectName(f"case_intake_{key}")
            self._field_widgets[key] = widget
            form.addRow(_review_label(key), widget)

        scroll.setWidget(form_host)
        layout.addWidget(scroll, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.run_docket_btn = theme.primary_button("Run Docket")
        self.run_docket_btn.clicked.connect(self._emit_run_docket_requested)
        btn_row.addWidget(self.run_docket_btn)
        layout.addLayout(btn_row)

        self._update_run_docket_enabled()

    def load_metadata(self, metadata: dict[str, Any] | None, complaint_file: str = "") -> None:
        metadata = metadata or {}
        self._complaint_file = str(complaint_file or "").strip()
        self._update_complaint_label()
        for key in REVIEW_FIELDS:
            self._set_field_value(key, metadata.get(key, ""))
        self._update_run_docket_enabled()

    def to_dict(self) -> dict[str, Any]:
        data = {
            key: normalize_review_value(key, self._field_text(key))
            for key in REVIEW_FIELDS
        }
        data["complaint_file"] = self._complaint_file
        return data

    def from_dict(self, data: dict | None) -> None:
        data = data or {}
        complaint_file = str(data.get("complaint_file", "") or "")
        self.load_metadata(data, complaint_file=complaint_file)

    def _set_field_value(self, key: str, value: Any) -> None:
        widget = self._field_widgets[key]
        normalized = normalize_review_value(key, value)
        if key in LIST_FIELDS:
            text = "\n".join(normalized) if isinstance(normalized, list) else str(normalized)
            widget.setPlainText(text)
        else:
            widget.setText(str(normalized or ""))

    def _field_text(self, key: str) -> str:
        widget = self._field_widgets[key]
        if isinstance(widget, QPlainTextEdit):
            return widget.toPlainText()
        return widget.text()

    def _update_complaint_label(self) -> None:
        if not self._complaint_file:
            self.complaint_label.setText("Complaint: Not selected")
            self.complaint_label.setToolTip("")
            return
        display = os.path.basename(self._complaint_file) or self._complaint_file
        self.complaint_label.setText(f"Complaint: {display}")
        self.complaint_label.setToolTip(self._complaint_file)

    def _update_run_docket_enabled(self, *_args) -> None:
        if not hasattr(self, "run_docket_btn"):
            return
        has_case_number = bool(normalize_review_value("case_number", self._field_text("case_number")))
        has_venue = bool(normalize_review_value("venue_county", self._field_text("venue_county")))
        self.run_docket_btn.setEnabled(has_case_number and has_venue)

    def _emit_run_docket_requested(self) -> None:
        self.run_docket_requested.emit(self.to_dict())


class CaseIntakeDocketOutputPage(QWidget):
    """Readable docket run summary."""

    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._summary: dict[str, Any] = {}
        self._output_path = ""

        layout = QVBoxLayout(self)
        layout.setContentsMargins(
            theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL
        )
        layout.setSpacing(theme.SPACE_MD)

        layout.addWidget(theme.page_title("Docket Summary"))
        self.summary_view = QPlainTextEdit()
        self.summary_view.setReadOnly(True)
        layout.addWidget(self.summary_view, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        self.rerun_btn = theme.secondary_button("Re-run")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        btn_row.addWidget(self.rerun_btn)
        self.edit_settings_btn = theme.secondary_button("Edit Settings & Re-run")
        self.edit_settings_btn.clicked.connect(self.edit_settings_requested.emit)
        btn_row.addWidget(self.edit_settings_btn)
        layout.addLayout(btn_row)

    @property
    def summary(self) -> dict[str, Any]:
        return self._copy_summary(self._summary)

    @property
    def output_path(self) -> str:
        return str(self._output_path or "")

    def show_summary(self, summary: dict[str, Any] | None) -> None:
        self._summary = self._copy_summary(summary or {})
        self._output_path = str(
            self._summary.get("docket_pdf")
            or self._summary.get("variables_docx")
            or ""
        )
        self.summary_view.setPlainText(self._render_summary(self._summary))

    def load_output(self, output_path: str) -> None:
        summary = self.summary
        path = str(output_path or "").strip()
        lower_path = path.lower()
        if lower_path.endswith(".pdf"):
            summary["docket_pdf"] = path
        elif lower_path.endswith(".docx") or os.path.basename(lower_path) == "variables.docx":
            summary["variables_docx"] = path
        self.show_summary(summary)

    @staticmethod
    def _copy_summary(summary: dict[str, Any]) -> dict[str, Any]:
        copied = dict(summary)
        if isinstance(copied.get("recent_lines"), list):
            copied["recent_lines"] = list(copied["recent_lines"])
        return copied

    @staticmethod
    def _clean(value: Any) -> str:
        return str(value or "").strip()

    def _render_summary(self, summary: dict[str, Any]) -> str:
        state = self._clean(summary.get("state"))
        lines = [
            f"Status: {self._clean(summary.get('status')) or 'No status available.'}",
        ]
        if state:
            lines.append(f"State: {state}")

        lines.extend([
            f"Docket PDF: {self._clean(summary.get('docket_pdf')) or '(not found)'}",
            f"Variables: {self._clean(summary.get('variables_docx')) or '(not found)'}",
            f"Trial date: {self._clean(summary.get('trial_date')) or '(not available)'}",
            f"Other hearings: {self._clean(summary.get('other_hearings')) or '(not available)'}",
            (
                "Procedural history: "
                f"{self._clean(summary.get('procedural_history')) or '(not available)'}"
            ),
        ])

        warning = self._clean(summary.get("warning"))
        if warning:
            lines.extend(["", f"Warning: {warning}"])

        recent_lines = summary.get("recent_lines") or []
        if recent_lines:
            lines.extend(["", "Final log lines:"])
            lines.extend(str(line) for line in recent_lines)
        return "\n".join(lines)


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
    docket_pdf = "" if not success else find_latest_docket_pdf(case_path)
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
