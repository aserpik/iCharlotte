# Case Intake & Docket Wizard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a Wizard Mode task that runs complaint intake, requires review of extracted case metadata, then runs docket processing from the reviewed values.

**Architecture:** Add a file-number subprocess worker for case-level scripts, a custom `CaseIntakeDocketTaskTab` with intake/status/review/output pages, and registry/routing/main-window integration. Keep `Scripts/complaint.py` and `Scripts/docket.py` as the execution source of truth in v1; the wizard wraps them and persists reviewed metadata through `Scripts.case_data_manager.CaseDataManager`.

**Tech Stack:** Python, PySide6, QProcess, pytest, pytest-qt, existing iCharlotte Wizard scaffold, `Scripts.case_data_manager.CaseDataManager`, `icharlotte_core.master_db.MasterCaseDatabase`.

---

## File Structure

- Create `icharlotte_core/ui/wizard/runners/case_agent_worker.py`
  - Runs `Scripts/<script_name> <file_number> --headless` via `QProcess`.
  - Emits `status`, `finished`, `failed`, and `cancelled`.
  - Stores recent stdout lines for output summaries.

- Create `icharlotte_core/ui/wizard/pages/case_intake_docket_page.py`
  - Contains pure helper functions for metadata load/save and output summary.
  - Defines `CaseIntakeSettingsPage`, `CaseMetadataReviewPage`, `CaseIntakeDocketOutputPage`, and `CaseIntakeDocketTaskTab`.
  - Defines `build_case_intake_docket_tab(spec, case_path, file_number, parent)`.

- Modify `icharlotte_core/ui/wizard/registry.py`
  - Add `case_intake_docket` task card in the General category.

- Modify `icharlotte_core/ui/wizard/task_routing.py`
  - Route `case_intake_docket` to `build_case_intake_docket_tab`.

- Modify `icharlotte_core/ui/wizard/in_process_task_tab.py`
  - Add a builder shim that imports the custom page builder.

- Modify `iCharlotte.py`
  - Restore and reopen `case_intake_docket` tabs with their saved review/output state rather than treating them as ephemeral builder-only tasks.

- Create `tests/test_wizard/test_case_agent_worker.py`
  - Unit tests for command construction and stdout parsing.

- Create `tests/test_wizard/test_case_intake_docket_page.py`
  - Unit/widget tests for metadata helpers, review persistence, output summary, and task tab orchestration.

- Modify `tests/test_wizard/test_registry.py`
  - Add the new task ID to the expected registry set.

- Modify `tests/test_wizard/test_task_categories.py`
  - Update General category count and expected assignment.

- Modify `tests/test_wizard/test_task_routing.py`
  - Assert the task uses a custom builder and skips the generic file picker.

- Create `tests/test_wizard/test_case_intake_docket_restore.py`
  - Regression tests for restore/reopen behavior through `MainWindow`.

---

## Task 1: Add Case-Level Agent Worker

**Files:**
- Create: `tests/test_wizard/test_case_agent_worker.py`
- Create: `icharlotte_core/ui/wizard/runners/case_agent_worker.py`

- [ ] **Step 1: Write the failing worker tests**

Create `tests/test_wizard/test_case_agent_worker.py`:

```python
import os

from icharlotte_core.ui.wizard.runners.case_agent_worker import CaseAgentWorker


def test_case_agent_worker_builds_file_number_command(monkeypatch):
    worker = CaseAgentWorker(
        script_name="complaint.py",
        case_path=r"C:\cases\1234",
        file_number="1234.001",
    )
    monkeypatch.setattr(worker, "_script_path", lambda: r"C:\repo\Scripts\complaint.py")

    assert worker.command_argv() == [
        r"C:\repo\Scripts\complaint.py",
        "1234.001",
        "--headless",
    ]


def test_case_agent_worker_preserves_extra_flags_after_headless(monkeypatch):
    worker = CaseAgentWorker(
        script_name="docket.py",
        case_path=r"C:\cases\1234",
        file_number="1234.001",
        extra_flags=["--headful"],
    )
    monkeypatch.setattr(worker, "_script_path", lambda: r"C:\repo\Scripts\docket.py")

    assert worker.command_argv() == [
        r"C:\repo\Scripts\docket.py",
        "1234.001",
        "--headless",
        "--headful",
    ]


def test_case_agent_worker_keeps_recent_lines_bounded():
    worker = CaseAgentWorker(
        script_name="docket.py",
        case_path=r"C:\cases\1234",
        file_number="1234.001",
        recent_line_limit=3,
    )

    for line in ["one", "two", "three", "four"]:
        worker._handle_line(line)

    assert worker.recent_lines == ["two", "three", "four"]


def test_case_agent_worker_script_path_points_to_scripts_folder():
    worker = CaseAgentWorker(
        script_name="docket.py",
        case_path=r"C:\cases\1234",
        file_number="1234.001",
    )

    path = worker._script_path()

    assert path.endswith(os.path.join("Scripts", "docket.py"))
```

- [ ] **Step 2: Run the worker tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_agent_worker.py -q
```

Expected: fails with `ModuleNotFoundError: No module named 'icharlotte_core.ui.wizard.runners.case_agent_worker'`.

- [ ] **Step 3: Implement `CaseAgentWorker`**

Create `icharlotte_core/ui/wizard/runners/case_agent_worker.py`:

```python
"""CaseAgentWorker - runs case-level Scripts/*.py agents by file number."""
from __future__ import annotations

import os
import sys
from collections import deque
from typing import Iterable

from PySide6.QtCore import QProcess, QTimer

from .base import BaseWorker


class CaseAgentWorker(BaseWorker):
    """Run an existing case-level agent script with the loaded file number.

    Unlike SubprocessWorker, this worker does not receive document paths and
    does not scan NOTES/AI Output for produced docx files. Complaint and docket
    update multiple case-level artifacts, so callers summarize current case
    state after the process exits.
    """

    def __init__(
        self,
        script_name: str,
        case_path: str,
        file_number: str,
        extra_flags: Iterable[str] | None = None,
        recent_line_limit: int = 80,
        parent=None,
    ):
        super().__init__(
            case_path=case_path,
            file_number=file_number,
            files=[],
            settings={},
            parent=parent,
        )
        self._script_name = script_name
        self._extra_flags = list(extra_flags or [])
        self._process: QProcess | None = None
        self._stdout_buf = bytearray()
        self._recent_lines = deque(maxlen=max(1, int(recent_line_limit)))

    @property
    def recent_lines(self) -> list[str]:
        return list(self._recent_lines)

    def command_argv(self) -> list[str]:
        return [self._script_path(), self.file_number, "--headless", *self._extra_flags]

    def _script_path(self) -> str:
        here = os.path.abspath(__file__)
        repo_root = here
        for _ in range(5):
            repo_root = os.path.dirname(repo_root)
        return os.path.join(repo_root, "Scripts", self._script_name)

    def start(self) -> None:
        argv = ["-u", *self.command_argv()]
        self.status.emit(
            "Running: python "
            + " ".join(
                os.path.basename(a) if os.sep in a or "/" in a else a
                for a in argv
            )
        )
        self._process = QProcess(self)
        self._process.setProcessChannelMode(QProcess.ProcessChannelMode.MergedChannels)
        self._process.readyReadStandardOutput.connect(self._drain_stdout)
        self._process.finished.connect(self._on_process_finished)
        self._process.errorOccurred.connect(self._on_process_error)
        self._process.start(sys.executable, argv)

    def cancel(self) -> None:
        super().cancel()
        if self._process is None:
            self.cancelled.emit()
            return
        self.status.emit("Cancellation requested.")
        self._process.terminate()
        QTimer.singleShot(2000, self._hard_kill_if_running)

    def _hard_kill_if_running(self) -> None:
        if (
            self._process is not None
            and self._process.state() != QProcess.ProcessState.NotRunning
        ):
            self.status.emit("Forcing kill.")
            self._process.kill()

    def _drain_stdout(self) -> None:
        if self._process is None:
            return
        data = bytes(self._process.readAllStandardOutput())
        if not data:
            return
        self._stdout_buf.extend(data)
        while b"\n" in self._stdout_buf:
            line_bytes, _, rest = self._stdout_buf.partition(b"\n")
            self._stdout_buf = bytearray(rest)
            line = line_bytes.decode("utf-8", errors="replace").rstrip("\r")
            self._handle_line(line)

    def _handle_line(self, line: str) -> None:
        text = (line or "").strip()
        if not text:
            return
        self._recent_lines.append(text)
        self.status.emit(text)

    def _on_process_finished(self, exit_code: int, exit_status: QProcess.ExitStatus) -> None:
        self._drain_stdout()
        if self._stdout_buf:
            self._handle_line(self._stdout_buf.decode("utf-8", errors="replace"))
            self._stdout_buf.clear()

        if self.is_cancel_requested:
            self.cancelled.emit()
            return
        if exit_code != 0:
            self.failed.emit(f"Agent exited with code {exit_code}.")
            return
        self.finished.emit("")

    def _on_process_error(self, err: QProcess.ProcessError) -> None:
        if self.is_cancel_requested:
            return
        self.failed.emit(f"Process error: {err}")
```

- [ ] **Step 4: Run the worker tests to verify they pass**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_agent_worker.py -q
```

Expected: all tests in `test_case_agent_worker.py` pass.

- [ ] **Step 5: Commit the worker**

```powershell
git add tests/test_wizard/test_case_agent_worker.py icharlotte_core/ui/wizard/runners/case_agent_worker.py
git commit -m "feat(wizard): add case agent subprocess worker"
```

---

## Task 2: Add Metadata And Output Summary Helpers

**Files:**
- Create: `tests/test_wizard/test_case_intake_docket_page.py`
- Create: `icharlotte_core/ui/wizard/pages/case_intake_docket_page.py`

- [ ] **Step 1: Write failing helper tests**

Create `tests/test_wizard/test_case_intake_docket_page.py`:

```python
import os
import time
from pathlib import Path

from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
    REVIEW_FIELDS,
    build_output_summary,
    find_complaint_candidate,
    find_latest_docket_pdf,
    load_case_metadata,
    normalize_review_value,
    save_reviewed_metadata,
)


class FakeManager:
    def __init__(self, values=None):
        self.values = dict(values or {})
        self.saved = []

    def get_value(self, file_number, key):
        return self.values.get(key)

    def save_variable(self, file_number, key, value, source="agent", auto_tag=True, extra_tags=None):
        self.saved.append({
            "file_number": file_number,
            "key": key,
            "value": value,
            "source": source,
            "auto_tag": auto_tag,
            "extra_tags": list(extra_tags or []),
        })
        self.values[key] = value


def test_normalize_review_value_splits_list_fields():
    assert normalize_review_value("plaintiffs", "Alice\nBob") == ["Alice", "Bob"]
    assert normalize_review_value("causes_of_action", "Negligence; Battery") == [
        "Negligence",
        "Battery",
    ]
    assert normalize_review_value("case_number", " 23STCV00123 ") == "23STCV00123"


def test_load_case_metadata_reads_review_fields():
    manager = FakeManager({
        "case_number": "23STCV00123",
        "venue_county": "Los Angeles",
        "plaintiffs": ["Alice", "Bob"],
    })

    metadata = load_case_metadata("1234.001", manager=manager)

    assert metadata["case_number"] == "23STCV00123"
    assert metadata["venue_county"] == "Los Angeles"
    assert metadata["plaintiffs"] == ["Alice", "Bob"]
    assert set(REVIEW_FIELDS) <= set(metadata)


def test_save_reviewed_metadata_uses_meta_data_tags():
    manager = FakeManager()

    save_reviewed_metadata(
        "1234.001",
        {"case_number": "23STCV00123", "plaintiffs": "Alice\nBob"},
        manager=manager,
    )

    assert manager.saved[0]["key"] == "case_number"
    assert manager.saved[0]["source"] == "wizard_case_intake"
    assert manager.saved[0]["extra_tags"] == ["Meta Data"]
    assert manager.saved[1]["key"] == "plaintiffs"
    assert manager.saved[1]["value"] == ["Alice", "Bob"]


def test_find_latest_docket_pdf_returns_newest(tmp_path):
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)
    older = out_dir / "Docket_2026.01.01.pdf"
    newer = out_dir / "Docket_2026.02.01.pdf"
    older.write_bytes(b"old")
    time.sleep(0.01)
    newer.write_bytes(b"new")

    assert find_latest_docket_pdf(str(tmp_path)) == str(newer)


def test_find_complaint_candidate_prefers_pleadings(tmp_path):
    pleadings = tmp_path / "PLEADINGS"
    pleadings.mkdir()
    complaint = pleadings / "Complaint.pdf"
    complaint.write_bytes(b"%PDF")

    assert find_complaint_candidate(str(tmp_path)) == str(complaint)


def test_build_output_summary_marks_no_docket_pdf_as_partial(tmp_path):
    variables_docx = tmp_path / "NOTES" / "AI OUTPUT" / "variables.docx"
    variables_docx.parent.mkdir(parents=True)
    variables_docx.write_bytes(b"docx")
    manager = FakeManager({
        "trial_date": "2026-09-01",
        "other_hearings": "CMC on 2026-07-01",
        "procedural_history": "Complaint filed.",
    })

    summary = build_output_summary(
        str(tmp_path),
        "1234.001",
        manager=manager,
        master_db=None,
        recent_lines=["Venue 'ventura' is not supported. Skipping download."],
        success=True,
    )

    assert summary["success"] is True
    assert summary["docket_pdf"] == ""
    assert summary["variables_docx"] == str(variables_docx)
    assert "No docket PDF was found" in summary["status"]
    assert summary["trial_date"] == "2026-09-01"
```

- [ ] **Step 2: Run helper tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_page.py -q
```

Expected: fails with `ModuleNotFoundError` for `case_intake_docket_page`.

- [ ] **Step 3: Implement helper functions and constants**

Create `icharlotte_core/ui/wizard/pages/case_intake_docket_page.py` with this initial helper-only content:

```python
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
            "Docket processing finished, but no docket PDF was found. "
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
```

- [ ] **Step 4: Run helper tests to verify they pass**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_page.py -q
```

Expected: all helper tests pass.

- [ ] **Step 5: Commit helpers**

```powershell
git add tests/test_wizard/test_case_intake_docket_page.py icharlotte_core/ui/wizard/pages/case_intake_docket_page.py
git commit -m "feat(wizard): add case intake docket helpers"
```

---

## Task 3: Add Review And Output Widgets

**Files:**
- Modify: `tests/test_wizard/test_case_intake_docket_page.py`
- Modify: `icharlotte_core/ui/wizard/pages/case_intake_docket_page.py`

- [ ] **Step 1: Add failing widget tests**

Append to `tests/test_wizard/test_case_intake_docket_page.py`:

```python
import pytest

pytest.importorskip("pytestqt")


def test_review_page_round_trips_metadata(qtbot):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
        CaseMetadataReviewPage,
    )

    page = CaseMetadataReviewPage()
    qtbot.addWidget(page)
    page.load_metadata({
        "case_number": "23STCV00123",
        "venue_county": "Los Angeles",
        "plaintiffs": ["Alice", "Bob"],
        "defendants": ["Acme"],
        "causes_of_action": ["Negligence"],
    }, complaint_file=r"C:\case\PLEADINGS\Complaint.pdf")

    data = page.to_dict()

    assert data["case_number"] == "23STCV00123"
    assert data["venue_county"] == "Los Angeles"
    assert data["plaintiffs"] == ["Alice", "Bob"]
    assert data["defendants"] == ["Acme"]
    assert data["causes_of_action"] == ["Negligence"]
    assert data["complaint_file"] == r"C:\case\PLEADINGS\Complaint.pdf"


def test_review_page_requires_case_number_and_venue(qtbot):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
        CaseMetadataReviewPage,
    )

    page = CaseMetadataReviewPage()
    qtbot.addWidget(page)
    page.load_metadata({"case_number": "", "venue_county": ""})

    assert page.run_docket_btn.isEnabled() is False

    page._field_widgets["case_number"].setText("23STCV00123")
    page._field_widgets["venue_county"].setText("Los Angeles")

    assert page.run_docket_btn.isEnabled() is True


def test_output_page_shows_summary_and_exposes_output_path(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
        CaseIntakeDocketOutputPage,
    )

    docket = tmp_path / "NOTES" / "AI OUTPUT" / "Docket_2026.06.06.pdf"
    docket.parent.mkdir(parents=True)
    docket.write_bytes(b"%PDF")
    page = CaseIntakeDocketOutputPage()
    qtbot.addWidget(page)

    page.show_summary({
        "success": True,
        "status": "Docket processing finished and a docket PDF was found.",
        "docket_pdf": str(docket),
        "variables_docx": "",
        "trial_date": "2026-09-01",
        "other_hearings": "CMC on 2026-07-01",
        "procedural_history": "Complaint filed.",
        "recent_lines": ["Agent finished"],
    })

    assert page.output_path == str(docket)
    assert "2026-09-01" in page.summary_view.toPlainText()
    assert "Agent finished" in page.summary_view.toPlainText()
```

- [ ] **Step 2: Run widget tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_page.py -q
```

Expected: fails because `CaseMetadataReviewPage` and `CaseIntakeDocketOutputPage` are not defined.

- [ ] **Step 3: Implement the review and output widgets**

Append these imports near the top of `case_intake_docket_page.py`:

```python
from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPlainTextEdit,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.ui.wizard import theme
```

Append these classes below `build_output_summary`:

```python
FIELD_LABELS = {
    "case_number": "Case number",
    "venue_county": "Venue county",
    "case_name": "Case name",
    "filing_date": "Filing date",
    "plaintiffs": "Plaintiffs",
    "defendants": "Defendants",
    "client_name": "Client name",
    "client_email": "Client email",
    "plaintiff_counsel": "Plaintiff counsel",
    "causes_of_action": "Causes of action",
}


class CaseIntakeSettingsPage(QWidget):
    run_complaint_requested = Signal()

    def __init__(self, file_number: str, parent: QWidget | None = None):
        super().__init__(parent)
        self.file_number = file_number
        outer = QVBoxLayout(self)
        outer.setContentsMargins(theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL)
        outer.setSpacing(theme.SPACE_LG)
        outer.addWidget(theme.page_title("Run complaint intake"))
        outer.addWidget(theme.helper_text(
            "This runs the Complaint Agent for the loaded case, then opens a review page before any docket search starts."
        ))
        self.file_label = QLabel(f"File number: {file_number or '(none)'}")
        self.file_label.setStyleSheet(f"font-weight: 600; color: {theme.TEXT};")
        outer.addWidget(self.file_label)
        outer.addStretch(1)
        row = QHBoxLayout()
        row.addStretch()
        self.run_btn = theme.primary_button("Run Complaint Intake")
        self.run_btn.setEnabled(bool(file_number))
        self.run_btn.clicked.connect(self.run_complaint_requested.emit)
        row.addWidget(self.run_btn)
        outer.addLayout(row)

    def to_dict(self) -> dict:
        return {}

    def from_dict(self, data: dict) -> None:
        return


class CaseMetadataReviewPage(QWidget):
    run_docket_requested = Signal(dict)

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._field_widgets: dict[str, QLineEdit | QPlainTextEdit] = {}
        self._complaint_file = ""
        outer = QVBoxLayout(self)
        outer.setContentsMargins(theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL)
        outer.setSpacing(theme.SPACE_MD)
        outer.addWidget(theme.page_title("Review case metadata"))
        outer.addWidget(theme.helper_text(
            "Confirm the case number and venue before running the Docket Agent. Edit any extracted values that look wrong."
        ))
        self.complaint_label = QLabel("Complaint file: -")
        self.complaint_label.setWordWrap(True)
        self.complaint_label.setStyleSheet(f"color: {theme.TEXT_MUTED};")
        outer.addWidget(self.complaint_label)

        form_host = QWidget()
        form = QFormLayout(form_host)
        form.setLabelAlignment(form.labelAlignment())
        for key in REVIEW_FIELDS:
            if key in LIST_FIELDS:
                widget = QPlainTextEdit()
                widget.setMinimumHeight(58)
                widget.textChanged.connect(self._update_run_enabled)
            else:
                widget = QLineEdit()
                widget.textChanged.connect(self._update_run_enabled)
            self._field_widgets[key] = widget
            form.addRow(FIELD_LABELS[key], widget)
        outer.addWidget(form_host, 1)

        self.error_label = theme.error_text("")
        outer.addWidget(self.error_label)

        row = QHBoxLayout()
        row.addStretch()
        self.run_docket_btn = theme.primary_button("Run Docket")
        self.run_docket_btn.clicked.connect(self._emit_run_docket)
        row.addWidget(self.run_docket_btn)
        outer.addLayout(row)
        self._update_run_enabled()

    def load_metadata(self, metadata: dict[str, Any], complaint_file: str = "") -> None:
        self._complaint_file = complaint_file or metadata.get("complaint_file", "") or ""
        self.complaint_label.setText(f"Complaint file: {self._complaint_file or '-'}")
        for key, widget in self._field_widgets.items():
            value = metadata.get(key, "")
            if key in LIST_FIELDS:
                if isinstance(value, list):
                    text = "\n".join(str(v) for v in value)
                else:
                    text = str(value or "")
                widget.setPlainText(text)
            else:
                widget.setText(str(value or ""))
        self._update_run_enabled()

    def to_dict(self) -> dict[str, Any]:
        data: dict[str, Any] = {}
        for key, widget in self._field_widgets.items():
            if isinstance(widget, QPlainTextEdit):
                raw = widget.toPlainText()
            else:
                raw = widget.text()
            data[key] = normalize_review_value(key, raw)
        data["complaint_file"] = self._complaint_file
        return data

    def from_dict(self, data: dict) -> None:
        self.load_metadata(data or {}, complaint_file=(data or {}).get("complaint_file", ""))

    def _update_run_enabled(self) -> None:
        if not hasattr(self, "run_docket_btn"):
            return
        data = self.to_dict()
        missing = []
        if not data.get("case_number"):
            missing.append("case number")
        if not data.get("venue_county"):
            missing.append("venue county")
        self.run_docket_btn.setEnabled(not missing)
        self.error_label.setText(
            f"Required before docket: {', '.join(missing)}" if missing else ""
        )

    def _emit_run_docket(self) -> None:
        self.run_docket_requested.emit(self.to_dict())


class CaseIntakeDocketOutputPage(QWidget):
    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self._summary: dict[str, Any] = {}
        self._output_path = ""
        outer = QVBoxLayout(self)
        outer.setContentsMargins(theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL, theme.SPACE_XL)
        outer.setSpacing(theme.SPACE_MD)
        outer.addWidget(theme.page_title("Case intake and docket complete"))
        self.summary_view = QPlainTextEdit()
        self.summary_view.setReadOnly(True)
        self.summary_view.setStyleSheet(
            f"font-family: {theme.MONO}; font-size: {theme.FONT_BODY}px;"
        )
        outer.addWidget(self.summary_view, 1)
        row = QHBoxLayout()
        row.addStretch()
        self.rerun_btn = theme.secondary_button("Run Again")
        self.rerun_btn.clicked.connect(self.rerun_requested.emit)
        row.addWidget(self.rerun_btn)
        self.edit_settings_btn = theme.secondary_button("Review Metadata")
        self.edit_settings_btn.clicked.connect(self.edit_settings_requested.emit)
        row.addWidget(self.edit_settings_btn)
        outer.addLayout(row)

    @property
    def output_path(self) -> str:
        return self._output_path

    @property
    def summary(self) -> dict[str, Any]:
        return dict(self._summary)

    def show_summary(self, summary: dict[str, Any]) -> None:
        self._summary = dict(summary or {})
        self._output_path = self._summary.get("docket_pdf") or self._summary.get("variables_docx") or ""
        lines = [
            self._summary.get("status", ""),
            "",
            f"Docket PDF: {self._summary.get('docket_pdf') or '-'}",
            f"Variables: {self._summary.get('variables_docx') or '-'}",
            f"Trial date: {self._summary.get('trial_date') or '-'}",
            f"Other hearings: {self._summary.get('other_hearings') or '-'}",
            "",
            "Procedural history:",
            self._summary.get("procedural_history") or "-",
        ]
        recent = self._summary.get("recent_lines") or []
        if recent:
            lines.extend(["", "Final log lines:", *[str(x) for x in recent]])
        self.summary_view.setPlainText("\n".join(lines))

    def load_output(self, output_path: str) -> None:
        summary = dict(self._summary)
        if output_path:
            summary["docket_pdf"] = output_path
        self.show_summary(summary)
```

- [ ] **Step 4: Run widget tests to verify they pass**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_page.py -q
```

Expected: all tests in `test_case_intake_docket_page.py` pass.

- [ ] **Step 5: Commit widgets**

```powershell
git add tests/test_wizard/test_case_intake_docket_page.py icharlotte_core/ui/wizard/pages/case_intake_docket_page.py
git commit -m "feat(wizard): add case intake review widgets"
```

---

## Task 4: Add Task Tab Orchestration

**Files:**
- Modify: `tests/test_wizard/test_case_intake_docket_page.py`
- Modify: `icharlotte_core/ui/wizard/pages/case_intake_docket_page.py`

- [ ] **Step 1: Add failing task-tab tests**

Append to `tests/test_wizard/test_case_intake_docket_page.py`:

```python
from PySide6.QtCore import QObject, Signal


class FakeCaseAgentWorker(QObject):
    status = Signal(str)
    progress = Signal(int)
    finished = Signal(str)
    failed = Signal(str)
    cancelled = Signal()

    instances = []

    def __init__(self, script_name, case_path, file_number, parent=None, **kwargs):
        super().__init__(parent)
        self.script_name = script_name
        self.case_path = case_path
        self.file_number = file_number
        self.recent_lines = []
        self.started = False
        FakeCaseAgentWorker.instances.append(self)

    def start(self):
        self.started = True

    def cancel(self):
        self.cancelled.emit()


def test_task_tab_starts_complaint_then_loads_review(qtbot, monkeypatch, tmp_path):
    import icharlotte_core.ui.wizard.pages.case_intake_docket_page as page_mod
    from icharlotte_core.ui.wizard.registry import TaskSpec

    FakeCaseAgentWorker.instances = []
    monkeypatch.setattr(page_mod, "CaseAgentWorker", FakeCaseAgentWorker)
    monkeypatch.setattr(
        page_mod,
        "load_case_metadata",
        lambda file_number: {"case_number": "23STCV00123", "venue_county": "Los Angeles"},
    )
    monkeypatch.setattr(page_mod, "find_complaint_candidate", lambda case_path: "Complaint.pdf")

    spec = TaskSpec(
        task_id="case_intake_docket",
        title="Case Intake & Docket",
        description="d",
        icon_glyph="D",
        script_name="",
    )
    tab = page_mod.CaseIntakeDocketTaskTab(spec, str(tmp_path), "1234.001")
    qtbot.addWidget(tab)

    tab.settings_page.run_complaint_requested.emit()

    assert FakeCaseAgentWorker.instances[0].script_name == "complaint.py"
    assert FakeCaseAgentWorker.instances[0].started is True

    FakeCaseAgentWorker.instances[0].finished.emit("")

    assert tab.currentIndex() == page_mod.TASK_PAGE_REVIEW
    assert tab.review_page.to_dict()["case_number"] == "23STCV00123"


def test_task_tab_saves_review_then_runs_docket(qtbot, monkeypatch, tmp_path):
    import icharlotte_core.ui.wizard.pages.case_intake_docket_page as page_mod
    from icharlotte_core.ui.wizard.registry import TaskSpec

    FakeCaseAgentWorker.instances = []
    saved = []
    monkeypatch.setattr(page_mod, "CaseAgentWorker", FakeCaseAgentWorker)
    monkeypatch.setattr(page_mod, "save_reviewed_metadata", lambda file_number, metadata: saved.append((file_number, metadata)))

    spec = TaskSpec(
        task_id="case_intake_docket",
        title="Case Intake & Docket",
        description="d",
        icon_glyph="D",
        script_name="",
    )
    tab = page_mod.CaseIntakeDocketTaskTab(spec, str(tmp_path), "1234.001")
    qtbot.addWidget(tab)

    tab.review_page.load_metadata({"case_number": "23STCV00123", "venue_county": "Los Angeles"})
    tab.review_page.run_docket_requested.emit(tab.review_page.to_dict())

    assert saved[0][0] == "1234.001"
    assert saved[0][1]["case_number"] == "23STCV00123"
    assert FakeCaseAgentWorker.instances[0].script_name == "docket.py"
    assert tab.currentIndex() == page_mod.TASK_PAGE_DOCKET_STATUS


def test_task_tab_emits_recent_task_entry_on_docket_finish(qtbot, monkeypatch, tmp_path):
    import icharlotte_core.ui.wizard.pages.case_intake_docket_page as page_mod
    from icharlotte_core.ui.wizard.registry import TaskSpec

    FakeCaseAgentWorker.instances = []
    monkeypatch.setattr(page_mod, "CaseAgentWorker", FakeCaseAgentWorker)
    monkeypatch.setattr(page_mod, "save_reviewed_metadata", lambda file_number, metadata: None)
    monkeypatch.setattr(
        page_mod,
        "build_output_summary",
        lambda *a, **k: {
            "success": True,
            "status": "done",
            "docket_pdf": r"C:\case\NOTES\AI OUTPUT\Docket_2026.06.06.pdf",
            "variables_docx": "",
            "trial_date": "",
            "other_hearings": "",
            "procedural_history": "",
            "recent_lines": [],
        },
    )

    spec = TaskSpec(
        task_id="case_intake_docket",
        title="Case Intake & Docket",
        description="d",
        icon_glyph="D",
        script_name="",
    )
    tab = page_mod.CaseIntakeDocketTaskTab(spec, str(tmp_path), "1234.001")
    qtbot.addWidget(tab)
    entries = []
    tab.task_completed.connect(entries.append)

    tab.review_page.load_metadata({"case_number": "23STCV00123", "venue_county": "Los Angeles"})
    tab.review_page.run_docket_requested.emit(tab.review_page.to_dict())
    FakeCaseAgentWorker.instances[0].finished.emit("")

    assert tab.currentIndex() == page_mod.TASK_PAGE_OUTPUT
    assert entries[0]["task_id"] == "case_intake_docket"
    assert entries[0]["output_path"].endswith("Docket_2026.06.06.pdf")
    assert entries[0]["settings"]["case_number"] == "23STCV00123"
```

- [ ] **Step 2: Run task-tab tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_page.py -q
```

Expected: fails because `CaseIntakeDocketTaskTab` and page constants are not defined.

- [ ] **Step 3: Implement the custom task tab**

Append these imports near the existing imports in `case_intake_docket_page.py`:

```python
from datetime import datetime

from PySide6.QtCore import Signal

from icharlotte_core.master_db import MasterCaseDatabase
from icharlotte_core.ui.wizard.pages.status_page import StatusPage
from icharlotte_core.ui.wizard.runners.case_agent_worker import CaseAgentWorker
from icharlotte_core.ui.wizard.task_scaffold import WizardTaskContainer
```

Append the task tab below the widget classes:

```python
TASK_PAGE_SETTINGS = 0
TASK_PAGE_COMPLAINT_STATUS = 1
TASK_PAGE_REVIEW = 2
TASK_PAGE_DOCKET_STATUS = 3
TASK_PAGE_OUTPUT = 4


class CaseIntakeDocketTaskTab(WizardTaskContainer):
    task_completed = Signal(dict)

    def __init__(
        self,
        spec,
        case_path: str,
        file_number: str,
        parent: QWidget | None = None,
    ):
        super().__init__(
            spec,
            steps=["Intake", "Complaint", "Review", "Docket", "Output"],
            parent=parent,
        )
        self._case_path = case_path
        self._file_number = file_number
        self._worker = None
        self._last_settings: dict[str, Any] = {}
        self._last_summary: dict[str, Any] = {}

        self.settings_page = CaseIntakeSettingsPage(file_number=file_number)
        self.complaint_status_page = StatusPage()
        self.review_page = CaseMetadataReviewPage()
        self.docket_status_page = StatusPage()
        self.output_page = CaseIntakeDocketOutputPage()

        self.addWidget(self.settings_page)
        self.addWidget(self.complaint_status_page)
        self.addWidget(self.review_page)
        self.addWidget(self.docket_status_page)
        self.addWidget(self.output_page)

        self.settings_page.run_complaint_requested.connect(self._start_complaint)
        self.review_page.run_docket_requested.connect(self._start_docket)
        self.complaint_status_page.cancel_requested.connect(self._on_cancel)
        self.docket_status_page.cancel_requested.connect(self._on_cancel)
        self.output_page.rerun_requested.connect(self._restart_flow)
        self.output_page.edit_settings_requested.connect(
            lambda: self.setCurrentIndex(TASK_PAGE_REVIEW)
        )

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list[str]:
        return []

    def load_review_state(self, metadata: dict[str, Any]) -> None:
        self._last_settings = dict(metadata or {})
        self.review_page.from_dict(self._last_settings)
        self.setCurrentIndex(TASK_PAGE_REVIEW)

    def load_output_summary(self, summary: dict[str, Any], metadata: dict[str, Any] | None = None) -> None:
        if metadata:
            self._last_settings = dict(metadata)
            self.review_page.from_dict(self._last_settings)
        self._last_summary = dict(summary or {})
        self.output_page.show_summary(self._last_summary)
        self.setCurrentIndex(TASK_PAGE_OUTPUT)

    def _start_complaint(self) -> None:
        self.complaint_status_page.reset()
        self.complaint_status_page.progress_bar.setRange(0, 0)
        self.complaint_status_page.on_status("Running Complaint Agent.")
        self.setCurrentIndex(TASK_PAGE_COMPLAINT_STATUS)
        worker = CaseAgentWorker(
            script_name="complaint.py",
            case_path=self._case_path,
            file_number=self._file_number,
            parent=self,
        )
        worker.status.connect(self.complaint_status_page.on_status)
        worker.finished.connect(lambda _unused: self._on_complaint_finished(worker))
        worker.failed.connect(lambda err: self._on_worker_failed(self.complaint_status_page, err))
        worker.cancelled.connect(self._on_worker_cancelled)
        self._worker = worker
        worker.start()

    def _on_complaint_finished(self, worker) -> None:
        self._worker = None
        metadata = load_case_metadata(self._file_number)
        complaint_file = find_complaint_candidate(self._case_path)
        metadata["complaint_file"] = complaint_file
        self._last_settings = dict(metadata)
        self.review_page.load_metadata(metadata, complaint_file=complaint_file)
        self.setCurrentIndex(TASK_PAGE_REVIEW)

    def _start_docket(self, metadata: dict[str, Any]) -> None:
        self._last_settings = dict(metadata or {})
        save_reviewed_metadata(self._file_number, self._last_settings)
        self.docket_status_page.reset()
        self.docket_status_page.progress_bar.setRange(0, 0)
        self.docket_status_page.on_status("Running Docket Agent.")
        self.setCurrentIndex(TASK_PAGE_DOCKET_STATUS)
        worker = CaseAgentWorker(
            script_name="docket.py",
            case_path=self._case_path,
            file_number=self._file_number,
            parent=self,
        )
        worker.status.connect(self.docket_status_page.on_status)
        worker.finished.connect(lambda _unused: self._on_docket_finished(worker, success=True))
        worker.failed.connect(lambda err: self._on_docket_failed(worker, err))
        worker.cancelled.connect(self._on_worker_cancelled)
        self._worker = worker
        worker.start()

    def _on_docket_failed(self, worker, err: str) -> None:
        self.docket_status_page.on_status(f"FAILED: {err}")
        self._on_docket_finished(worker, success=False)

    def _on_docket_finished(self, worker, success: bool) -> None:
        self._worker = None
        summary = build_output_summary(
            self._case_path,
            self._file_number,
            master_db=MasterCaseDatabase(),
            recent_lines=getattr(worker, "recent_lines", []),
            success=success,
        )
        self._last_summary = dict(summary)
        self.output_page.show_summary(summary)
        self.setCurrentIndex(TASK_PAGE_OUTPUT)
        entry = {
            "task_id": self._spec.task_id,
            "title": self._spec.title,
            "files": [],
            "settings": dict(self._last_settings),
            "summary": dict(summary),
            "output_path": summary.get("docket_pdf") or summary.get("variables_docx") or "",
            "output_paths": [
                p for p in [summary.get("docket_pdf"), summary.get("variables_docx")] if p
            ],
            "completed_at": datetime.now().isoformat(timespec="seconds"),
        }
        self.task_completed.emit(entry)

    def _on_worker_failed(self, status_page: StatusPage, err: str) -> None:
        self._worker = None
        status_page.on_status(f"FAILED: {err}")
        status_page.cancel_btn.setText("Back")
        status_page.cancel_btn.setEnabled(True)
        try:
            status_page.cancel_btn.clicked.disconnect()
        except RuntimeError:
            pass
        status_page.cancel_btn.clicked.connect(lambda: self.setCurrentIndex(TASK_PAGE_SETTINGS))

    def _on_cancel(self) -> None:
        if self._worker is not None:
            self._worker.cancel()
        else:
            self.setCurrentIndex(TASK_PAGE_SETTINGS)

    def _on_worker_cancelled(self) -> None:
        self._worker = None
        self.setCurrentIndex(TASK_PAGE_SETTINGS)

    def _restart_flow(self) -> None:
        self._worker = None
        self._last_summary = {}
        self.setCurrentIndex(TASK_PAGE_SETTINGS)


def build_case_intake_docket_tab(
    spec,
    case_path: str,
    file_number: str,
    parent: QWidget | None = None,
) -> CaseIntakeDocketTaskTab:
    return CaseIntakeDocketTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        parent=parent,
    )
```

- [ ] **Step 4: Run task-tab tests to verify they pass**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_page.py -q
```

Expected: all tests in `test_case_intake_docket_page.py` pass.

- [ ] **Step 5: Commit task orchestration**

```powershell
git add tests/test_wizard/test_case_intake_docket_page.py icharlotte_core/ui/wizard/pages/case_intake_docket_page.py
git commit -m "feat(wizard): orchestrate case intake docket flow"
```

---

## Task 5: Register And Route The Wizard Task

**Files:**
- Modify: `tests/test_wizard/test_registry.py`
- Modify: `tests/test_wizard/test_task_categories.py`
- Modify: `tests/test_wizard/test_task_routing.py`
- Modify: `icharlotte_core/ui/wizard/registry.py`
- Modify: `icharlotte_core/ui/wizard/task_routing.py`
- Modify: `icharlotte_core/ui/wizard/in_process_task_tab.py`

- [ ] **Step 1: Add failing registry/routing assertions**

Update `tests/test_wizard/test_registry.py::test_initial_tasks_registered` expected IDs to include `case_intake_docket`:

```python
def test_initial_tasks_registered():
    ids = {t.task_id for t in list_tasks()}
    assert ids == {
        "summarize_documents",
        "summarize_discovery",
        "summarize_depositions",
        "depo_prep",
        "medical_records",
        "med_chron_analysis",
        "med_record_extractor",
        "separate",
        "subpoena_tracker",
        "respond_to_discovery",
        "oppose_motion",
        "generate_motion",
        "mediation_brief",
        "case_intake_docket",
        "chat",
    }
```

Append to `tests/test_wizard/test_task_routing.py`:

```python
    def test_case_intake_docket_uses_custom_builder_without_file_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("case_intake_docket"),
            "build_case_intake_docket_tab",
        )
        self.assertTrue(is_in_process_task("case_intake_docket"))
        self.assertFalse(requires_initial_file_picker("case_intake_docket"))
```

Update `tests/test_wizard/test_task_categories.py::test_expected_category_assignments` General section:

```python
    # General
    assert TASK_REGISTRY["chat"].category == "General"
    assert TASK_REGISTRY["separate"].category == "General"
    assert TASK_REGISTRY["case_intake_docket"].category == "General"
```

Update `tests/test_wizard/test_task_categories.py::test_empty_query_returns_all_tasks_grouped_in_category_order` General count:

```python
    assert len(grouped["General"]) == 3
```

- [ ] **Step 2: Run registry/routing tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_wizard/test_registry.py tests/test_wizard/test_task_categories.py tests/test_wizard/test_task_routing.py -q
```

Expected: fails because `case_intake_docket` is not registered or routed.

- [ ] **Step 3: Register the task**

In `icharlotte_core/ui/wizard/registry.py`, add this `TaskSpec` before `"chat"`:

```python
    "case_intake_docket": TaskSpec(
        task_id="case_intake_docket",
        title="Case Intake & Docket",
        description="Extract complaint metadata, review case details, then download and process the court docket.",
        icon_glyph="\U0001F5C2",
        script_name="",
        category="General",
        keywords=[
            "complaint", "docket", "case number", "venue",
            "intake", "hearing", "trial",
        ],
        default_folders=[],
    ),
```

In `icharlotte_core/ui/wizard/task_routing.py`, add the builder mapping:

```python
_IN_PROCESS_TASK_BUILDERS = {
    "subpoena_tracker": "build_subpoena_tab",
    "med_record_extractor": "build_med_extractor_tab",
    "respond_to_discovery": "build_respond_to_discovery_tab",
    "oppose_motion": "build_oppose_motion_tab",
    "separate": "build_separate_tab",
    "generate_motion": "build_generate_motion_tab",
    "mediation_brief": "build_mediation_brief_tab",
    "case_intake_docket": "build_case_intake_docket_tab",
}
```

In `icharlotte_core/ui/wizard/in_process_task_tab.py`, append this builder:

```python
def build_case_intake_docket_tab(
    spec,
    case_path: str,
    file_number: str,
    parent: QWidget | None,
):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
        build_case_intake_docket_tab as _build,
    )

    return _build(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        parent=parent,
    )
```

- [ ] **Step 4: Run registry/routing tests to verify they pass**

Run:

```powershell
python -m pytest tests/test_wizard/test_registry.py tests/test_wizard/test_task_categories.py tests/test_wizard/test_task_routing.py -q
```

Expected: all listed tests pass.

- [ ] **Step 5: Commit registration and routing**

```powershell
git add tests/test_wizard/test_registry.py tests/test_wizard/test_task_categories.py tests/test_wizard/test_task_routing.py icharlotte_core/ui/wizard/registry.py icharlotte_core/ui/wizard/task_routing.py icharlotte_core/ui/wizard/in_process_task_tab.py
git commit -m "feat(wizard): register case intake docket task"
```

---

## Task 6: Add Restore And Recent-Task Reopen Support

**Files:**
- Create: `tests/test_wizard/test_case_intake_docket_restore.py`
- Modify: `iCharlotte.py`

- [ ] **Step 1: Write failing restore/reopen tests**

Create `tests/test_wizard/test_case_intake_docket_restore.py`:

```python
"""Restore/reopen tests for Case Intake & Docket wizard tabs."""
import os
import sys

import pytest

pytest.importorskip("pytestqt")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PySide6.QtWidgets import QTabWidget, QWidget


class _Stub(QWidget):
    def __init__(self, case_path, file_number="1234.001"):
        super().__init__()
        self.tabs = QTabWidget()
        self.case_path = case_path
        self.file_number = file_number

    def _on_task_completed(self, entry):
        pass

    def _hide_fixed_close_buttons(self):
        pass


def _bind(stub, method_name):
    import iCharlotte as ich

    method = getattr(ich.MainWindow, method_name)
    setattr(stub, method_name, method.__get__(stub, type(stub)))


def _write_docket(tmp_path):
    out_dir = tmp_path / "NOTES" / "AI OUTPUT"
    out_dir.mkdir(parents=True)
    docket = out_dir / "Docket_2026.06.06.pdf"
    docket.write_bytes(b"%PDF")
    return docket


def test_restore_reloads_case_intake_docket_output(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import TASK_PAGE_OUTPUT
    from icharlotte_core.ui.wizard.persistence import WizardStatePersistence

    case_path = str(tmp_path)
    docket = _write_docket(tmp_path)
    p = WizardStatePersistence(case_path)
    p.set_open_tabs([{
        "task_id": "case_intake_docket",
        "instance_suffix": "",
        "files": [],
        "settings": {
            "case_number": "23STCV00123",
            "venue_county": "Los Angeles",
        },
        "summary": {
            "success": True,
            "status": "done",
            "docket_pdf": os.path.relpath(str(docket), case_path),
            "variables_docx": "",
            "trial_date": "2026-09-01",
            "other_hearings": "",
            "procedural_history": "",
            "recent_lines": [],
        },
        "page": "output",
        "output_path": os.path.relpath(str(docket), case_path),
    }])
    p.save()

    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    _bind(stub, "_restore_task_tabs_for_case")
    stub._restore_task_tabs_for_case()

    assert stub.tabs.count() == 1
    tab = stub.tabs.widget(0)
    assert type(tab).__name__ == "CaseIntakeDocketTaskTab"
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert os.path.basename(tab.output_page.output_path) == os.path.basename(str(docket))


def test_reopen_recent_reloads_case_intake_docket_output(qtbot, tmp_path):
    from icharlotte_core.ui.wizard.pages.case_intake_docket_page import TASK_PAGE_OUTPUT

    case_path = str(tmp_path)
    docket = _write_docket(tmp_path)
    stub = _Stub(case_path)
    qtbot.addWidget(stub)
    _bind(stub, "_on_reopen_recent_task")
    stub._on_reopen_recent_task({
        "task_id": "case_intake_docket",
        "instance_suffix": "",
        "files": [],
        "settings": {
            "case_number": "23STCV00123",
            "venue_county": "Los Angeles",
        },
        "summary": {
            "success": True,
            "status": "done",
            "docket_pdf": os.path.relpath(str(docket), case_path),
            "variables_docx": "",
            "trial_date": "2026-09-01",
            "other_hearings": "",
            "procedural_history": "",
            "recent_lines": [],
        },
        "output_path": os.path.relpath(str(docket), case_path),
    })

    assert stub.tabs.count() == 1
    tab = stub.tabs.widget(0)
    assert type(tab).__name__ == "CaseIntakeDocketTaskTab"
    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert os.path.basename(tab.output_page.output_path) == os.path.basename(str(docket))
```

- [ ] **Step 2: Run restore tests to verify they fail**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_restore.py -q
```

Expected: fails because `MainWindow` treats the new custom builder as an ephemeral in-process tab and does not restore output state.

- [ ] **Step 3: Update `iCharlotte.py` custom restore/reopen branches**

In `iCharlotte.py`, update both builder-name exclusion tuples in `_on_reopen_recent_task` and `_restore_task_tabs_for_case` to include `build_case_intake_docket_tab`:

```python
if builder_name and builder_name not in (
    "build_oppose_motion_tab",
    "build_mediation_brief_tab",
    "build_generate_motion_tab",
    "build_case_intake_docket_tab",
):
```

In `_on_reopen_recent_task`, add this branch after the mediation/generate branches and before the generic `TaskTab` branch:

```python
        elif get_in_process_task_builder_name(task_id) == "build_case_intake_docket_tab":
            from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
                CaseIntakeDocketTaskTab,
                TASK_PAGE_OUTPUT,
                TASK_PAGE_REVIEW,
            )

            settings = dict(entry.get("settings") or {})
            task_tab = CaseIntakeDocketTaskTab(
                spec=spec,
                case_path=self.case_path,
                file_number=self.file_number,
                parent=self,
            )
            task_tab.load_review_state(settings)
            summary = dict(entry.get("summary") or {})
            if summary:
                for key in ("docket_pdf", "variables_docx"):
                    value = summary.get(key)
                    if value and self.case_path and not os.path.isabs(value):
                        summary[key] = os.path.join(self.case_path, value)
                task_tab.load_output_summary(summary, metadata=settings)
            output_page = TASK_PAGE_OUTPUT
            settings_page = TASK_PAGE_REVIEW
```

In `_restore_task_tabs_for_case`, add this branch after the mediation/generate branches and before the generic `TaskTab` branch:

```python
            elif get_in_process_task_builder_name(task_id) == "build_case_intake_docket_tab":
                from icharlotte_core.ui.wizard.pages.case_intake_docket_page import (
                    CaseIntakeDocketTaskTab,
                    TASK_PAGE_OUTPUT,
                    TASK_PAGE_REVIEW,
                )

                tab = CaseIntakeDocketTaskTab(
                    spec=spec,
                    case_path=self.case_path,
                    file_number=self.file_number,
                    parent=self,
                )
                tab.load_review_state(settings_dict)
                summary = dict(entry.get("summary") or {})
                if summary:
                    for key in ("docket_pdf", "variables_docx"):
                        value = summary.get(key)
                        if value and self.case_path and not os.path.isabs(value):
                            summary[key] = os.path.join(self.case_path, value)
                    tab.load_output_summary(summary, metadata=settings_dict)
                output_page = TASK_PAGE_OUTPUT
                settings_page = TASK_PAGE_REVIEW
```

Also update `_snapshot_open_task_tabs` so it stores `summary` when a task output page exposes it. In the `snapshots.append` block, include:

```python
                "summary": getattr(tab.output_page, "summary", {}),
```

The resulting snapshot dict should contain:

```python
            snapshots.append({
                "task_id": tab.spec.task_id,
                "instance_suffix": tab.property("wizard_instance_suffix") or "",
                "files": files_rel,
                "settings": tab.settings_page.to_dict()
                    if tab.spec.task_id != "case_intake_docket"
                    else getattr(tab, "_last_settings", {}),
                "summary": getattr(tab.output_page, "summary", {}),
                "page": page,
                "output_path": output_path_rel,
            })
```

- [ ] **Step 4: Run restore tests to verify they pass**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_intake_docket_restore.py -q
```

Expected: both tests pass.

- [ ] **Step 5: Run restart persistence regression tests**

Run:

```powershell
python -m pytest tests/test_wizard/test_restart_persistence.py tests/test_wizard/test_generate_motion_restore.py tests/test_wizard/test_mediation_brief_restore.py -q
```

Expected: all listed tests pass; no regression in existing restore behavior.

- [ ] **Step 6: Commit restore/reopen support**

```powershell
git add tests/test_wizard/test_case_intake_docket_restore.py iCharlotte.py
git commit -m "feat(wizard): restore case intake docket tabs"
```

---

## Task 7: Verification And Smoke Checks

**Files:**
- No new files.
- Verify the files changed in Tasks 1-6.

- [ ] **Step 1: Run focused wizard tests**

Run:

```powershell
python -m pytest tests/test_wizard/test_case_agent_worker.py tests/test_wizard/test_case_intake_docket_page.py tests/test_wizard/test_case_intake_docket_restore.py tests/test_wizard/test_registry.py tests/test_wizard/test_task_categories.py tests/test_wizard/test_task_routing.py -q
```

Expected: all listed tests pass.

- [ ] **Step 2: Run related restore and launcher tests**

Run:

```powershell
python -m pytest tests/test_wizard/test_wizard_tab.py tests/test_wizard/test_wizard_tab_search.py tests/test_wizard/test_restart_persistence.py tests/test_wizard/test_generate_motion_restore.py tests/test_wizard/test_mediation_brief_restore.py -q
```

Expected: all listed tests pass.

- [ ] **Step 3: Compile changed Python modules**

Run:

```powershell
python -m py_compile icharlotte_core/ui/wizard/runners/case_agent_worker.py icharlotte_core/ui/wizard/pages/case_intake_docket_page.py icharlotte_core/ui/wizard/registry.py icharlotte_core/ui/wizard/task_routing.py icharlotte_core/ui/wizard/in_process_task_tab.py iCharlotte.py
```

Expected: command exits with status 0 and prints no syntax errors.

- [ ] **Step 4: Check staged diff hygiene before final commit**

Run:

```powershell
git diff --check
git status --short
```

Expected: `git diff --check` prints no whitespace errors. `git status --short` shows only the files intentionally changed by this feature plus pre-existing unrelated dirty files that must not be staged or reverted.

- [ ] **Step 5: Commit final verification-only adjustments if any were needed**

If Task 7 required fixes, commit only those feature files:

```powershell
git add icharlotte_core/ui/wizard/runners/case_agent_worker.py icharlotte_core/ui/wizard/pages/case_intake_docket_page.py icharlotte_core/ui/wizard/registry.py icharlotte_core/ui/wizard/task_routing.py icharlotte_core/ui/wizard/in_process_task_tab.py iCharlotte.py tests/test_wizard/test_case_agent_worker.py tests/test_wizard/test_case_intake_docket_page.py tests/test_wizard/test_case_intake_docket_restore.py tests/test_wizard/test_registry.py tests/test_wizard/test_task_categories.py tests/test_wizard/test_task_routing.py
git commit -m "test(wizard): verify case intake docket task"
```

Expected: commit succeeds only if fixes were needed. If no fixes were needed after Task 6, skip this commit.

---

## Completion Criteria

- Wizard launcher shows `Case Intake & Docket` under General.
- Selecting the task opens the custom intake tab without a file picker.
- Complaint step runs `Scripts/complaint.py <file_number> --headless`.
- Docket cannot run until review page has at least case number and venue county.
- Review edits are saved through `CaseDataManager.save_variable(file_number, key, value, source="wizard_case_intake", extra_tags=["Meta Data"])` before docket starts.
- Docket step runs `Scripts/docket.py <file_number> --headless`.
- Output page clearly distinguishes docket PDF found versus no docket PDF found.
- Recent task and restart restore output/review state.
- Focused tests and `py_compile` commands pass.
