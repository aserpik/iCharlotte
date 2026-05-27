"""Wizard page for responding to incoming discovery."""
from __future__ import annotations

import os
from dataclasses import asdict

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QCheckBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QRadioButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.discovery.response_generation_engine import (
    _apply_proposal_to_review_state,
    _build_pending_review_state,
)
from icharlotte_core.discovery.response_context_index import (
    build_context_chunks,
)
from icharlotte_core.discovery.discovery_parse_worker import DiscoveryParseWorker
from icharlotte_core.discovery.proposal_coordinator import ProposalCoordinator
from icharlotte_core.llm_config import LLMConfig
from icharlotte_core.discovery.form_interrogatory_selection import (
    complete_selected_form_interrogatories,
    extract_selected_form_interrogatory_numbers,
)
from icharlotte_core.discovery.response_parser import (
    ParsedDiscovery,
    ParsedRequest,
)
from icharlotte_core.discovery.response_review_state import (
    RequestReview,
    ReviewState,
    insert_quick_objection,
    remove_quick_objection,
)
from icharlotte_core.discovery.response_rule_library import (
    RuleCategory,
    RuleMode,
    ResponseRule,
    built_in_rules_for,
    get_quick_objection_rules,
    load_global_rules,
    save_global_rules,
)
from icharlotte_core.discovery.response_rules import ResponseRules
from icharlotte_core.discovery.response_type_detector import normalize_discovery_type
from icharlotte_core.discovery._io import (
    read_document_text,
)


def parsed_discovery_to_dict(parsed: ParsedDiscovery | None) -> dict | None:
    return asdict(parsed) if parsed is not None else None


def parsed_discovery_from_dict(data: dict | None) -> ParsedDiscovery | None:
    if not data:
        return None
    raw = dict(data)
    requests = [
        ParsedRequest(**item)
        for item in raw.pop("requests", [])
        if isinstance(item, dict)
    ]
    return ParsedDiscovery(**raw, requests=requests)


def context_file_start_dir(discovery_file: str, case_root: str = "") -> str:
    """Return the default folder for context-file selection."""
    case_status = _find_status_dir(case_root)
    if case_status:
        return case_status

    parent = os.path.dirname(discovery_file or "")
    if parent:
        parent_status = _find_status_dir(parent)
        if parent_status:
            return parent_status
        if os.path.isdir(parent):
            return parent
    return case_root or ""


def _find_status_dir(parent: str) -> str:
    if not parent or not os.path.isdir(parent):
        return ""
    try:
        entries = os.listdir(parent)
    except OSError:
        return ""
    match = next((entry for entry in entries if entry.lower() == "status"), None)
    if not match:
        return ""
    candidate = os.path.join(parent, match)
    return candidate if os.path.isdir(candidate) else ""


def response_save_default_dir(case_root: str) -> str:
    """Return the user-facing default save folder for final responses."""
    discovery_dir = _find_child_dir(case_root, "DISCOVERY")
    if discovery_dir:
        responses_dir = _find_child_dir(discovery_dir, "RESPONSES")
        return responses_dir or os.path.join(discovery_dir, "RESPONSES")
    return os.path.join(case_root or "", "DISCOVERY", "RESPONSES")


def suggested_response_filename(parsed: ParsedDiscovery | None) -> str:
    """Return the recommended response filename for a parsed discovery set."""
    if parsed is None:
        return "Discovery Responses.docx"
    from icharlotte_core.discovery.response_assembler import build_response_filename

    return build_response_filename(
        abbreviation=_response_abbreviation(parsed.responding_party),
        disc_type=parsed.discovery_type,
        set_number=parsed.set_number,
    )


def preview_response_output_path(case_root: str, parsed: ParsedDiscovery | None) -> str:
    """Return an internal preview path; final user save happens via Save As."""
    preview_dir = os.path.join(
        case_root or "",
        ".icharlotte",
        "wizard_previews",
        "respond_to_discovery",
    )
    return os.path.join(preview_dir, suggested_response_filename(parsed))


def _find_child_dir(parent: str, child_name: str) -> str:
    if not parent or not os.path.isdir(parent):
        return ""
    try:
        entries = os.listdir(parent)
    except OSError:
        return ""
    target = child_name.lower()
    match = next((entry for entry in entries if entry.lower() == target), None)
    if not match:
        return ""
    candidate = os.path.join(parent, match)
    return candidate if os.path.isdir(candidate) else ""


REFER_TO_DOCUMENT_RESPONSE = (
    "Responding Party further elects to exercise its right pursuant to Code of "
    "Civil Procedure section 2030.230 to specify writings from which the answer "
    "may be derived or ascertained, and identifies and refers to the documents "
    "produced concurrently herewith.."
)

WILL_PRODUCE_RESPONSE = (
    "Responding Party will comply with this request and produce all "
    "non-privileged documents in Responding Party's possession, custody and "
    "control that Responding Party understands to be responsive to this Request. "
    "Responding Party identifies and refers to the documents produced "
    "concurrently herewith."
)

WONT_PRODUCE_RESPONSE = (
    "Upon a diligent search and reasonable inquiry made in an effort to locate "
    "the item(s) requested, Responding Party is unable to comply with this "
    "request at this time because the documents responsive to this request, if "
    "they exist, are not in the possession, custody or control of Responding Party."
)

CANT_ADMIT_DENY_RESPONSE = (
    "After a reasonable inquiry concerning the matter in this request, the "
    "information known or readily obtainable to Responding Party is insufficient "
    "to enable Responding Party to admit the matter."
)


class RespondDiscoverySettingsPage(QWidget):
    """Rules, context selection, and review UI for the discovery response task."""

    run_requested = Signal(dict)

    def __init__(
        self,
        case_root: str,
        file_number: str,
        discovery_file: str,
        detected_type: str,
        parsed_discovery: ParsedDiscovery | None = None,
        review_state: ReviewState | None = None,
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.case_root = case_root
        self.file_number = file_number
        self.discovery_file = discovery_file
        self.detected_type = normalize_discovery_type(detected_type) or "SI"
        self.fi_mode = "fixed" if self.detected_type == "FI" else "custom"
        self.context_files: list[str] = []
        self.parsed_discovery = parsed_discovery
        self.review_state = review_state
        self._rules: list[ResponseRule] = []
        self._rule_checks: dict[str, QCheckBox] = {}
        self._quick_checks: dict[str, QCheckBox] = {}
        self._quick_response_buttons: dict[str, QPushButton] = {}
        self._parse_worker: DiscoveryParseWorker | None = None
        self._coordinator: ProposalCoordinator | None = None
        self._context_chunks: list = []
        self._loaded_response_rules = None
        self._current_review_index = 0

        self._build_ui()
        if review_state is not None:
            self._show_review()

    # ------------------------------------------------------------------
    # Rules screen
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 24, 24, 24)
        outer.setSpacing(12)

        self.rules_widget = QWidget()
        rules_outer = QVBoxLayout(self.rules_widget)
        rules_outer.setContentsMargins(0, 0, 0, 0)
        rules_outer.setSpacing(12)
        outer.addWidget(self.rules_widget, 1)

        header = QLabel(
            f"Responding to: {os.path.basename(self.discovery_file)} "
            f"({self.detected_type})"
        )
        header.setStyleSheet("font-size: 14px; font-weight: 600;")
        rules_outer.addWidget(header)

        if self.detected_type == "FI":
            fi_group = QGroupBox("Form Interrogatory Rule Mode")
            fi_layout = QVBoxLayout(fi_group)
            self.rb_fi_fixed = QRadioButton("Use Fixed Objections/Responses")
            self.rb_fi_custom = QRadioButton("Use Custom Objections/Responses")
            self.rb_fi_fixed.toggled.connect(
                lambda checked: checked and self.set_fi_mode("fixed")
            )
            self.rb_fi_custom.toggled.connect(
                lambda checked: checked and self.set_fi_mode("custom")
            )
            self._sync_fi_mode_controls()
            fi_layout.addWidget(self.rb_fi_fixed)
            fi_layout.addWidget(self.rb_fi_custom)
            rules_outer.addWidget(fi_group)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.rule_container = QWidget()
        self.rule_layout = QVBoxLayout(self.rule_container)
        self.rule_layout.setContentsMargins(0, 0, 0, 0)
        self.rule_layout.setSpacing(10)
        self.scroll.setWidget(self.rule_container)
        rules_outer.addWidget(self.scroll, 1)

        btn_row = QHBoxLayout()
        self.add_rule_btn = QPushButton("Add Rule")
        self.add_rule_btn.clicked.connect(self._on_add_rule)
        btn_row.addWidget(self.add_rule_btn)
        btn_row.addStretch()
        self.next_btn = QPushButton("Next: Context Files")
        self.next_btn.clicked.connect(self._on_select_context_files)
        btn_row.addWidget(self.next_btn)
        rules_outer.addLayout(btn_row)

        self.status_label = QLabel("")
        self.status_label.setWordWrap(True)
        rules_outer.addWidget(self.status_label)

        self.review_widget = self._build_review_widget()
        self.review_widget.hide()
        outer.addWidget(self.review_widget, 1)
        self._rebuild_rules()

    def set_fi_mode(self, mode: str) -> None:
        self.fi_mode = "fixed" if mode == "fixed" else "custom"
        self._sync_fi_mode_controls()
        self._rebuild_rules()

    def _sync_fi_mode_controls(self) -> None:
        fixed_btn = getattr(self, "rb_fi_fixed", None)
        custom_btn = getattr(self, "rb_fi_custom", None)
        if not fixed_btn or not custom_btn:
            return
        fixed_btn.blockSignals(True)
        custom_btn.blockSignals(True)
        fixed_btn.setChecked(self.fi_mode == "fixed")
        custom_btn.setChecked(self.fi_mode != "fixed")
        fixed_btn.blockSignals(False)
        custom_btn.blockSignals(False)

    def _clear_rules(self) -> None:
        while self.rule_layout.count():
            item = self.rule_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._rule_checks.clear()

    def _rebuild_rules(self) -> None:
        self._clear_rules()
        self._rules = self._rules_for_current_type()
        if not self._rules:
            label = QLabel(
                "Fixed FI objections/responses will be used. Other applicable "
                "form interrogatories will be drafted from the context file(s)."
            )
            label.setWordWrap(True)
            self.rule_layout.addWidget(label)
            return

        for heading, category in (
            ("Objection Rules", RuleCategory.OBJECTION),
            ("Substantive Response Rules", RuleCategory.SUBSTANTIVE),
        ):
            category_rules = [r for r in self._rules if r.category == category]
            if not category_rules:
                continue
            group = QGroupBox(heading)
            group_layout = QVBoxLayout(group)
            for rule in category_rules:
                cb = QCheckBox(rule.name)
                cb.setChecked(rule.enabled_by_default)
                cb.setToolTip(rule.description)
                group_layout.addWidget(cb)

                desc = QLabel(rule.description)
                desc.setWordWrap(True)
                desc.setStyleSheet("color: #666; margin-left: 20px;")
                group_layout.addWidget(desc)
                self._rule_checks[rule.id] = cb
            self.rule_layout.addWidget(group)
        self.rule_layout.addStretch(1)

    def _on_add_rule(self) -> None:
        name, ok = QInputDialog.getText(self, "Add Rule", "Rule name:")
        if not ok or not name.strip():
            return
        category_label, ok = QInputDialog.getItem(
            self,
            "Add Rule",
            "Rule section:",
            ["Objection Rules", "Substantive Response Rules"],
            0,
            False,
        )
        if not ok:
            return
        description, ok = QInputDialog.getMultiLineText(
            self, "Add Rule", "What should this rule do?"
        )
        if not ok:
            return
        output_text, ok = QInputDialog.getMultiLineText(
            self, "Add Rule", "Exact objection or instruction text:"
        )
        if not ok:
            return
        category = (
            RuleCategory.SUBSTANTIVE
            if category_label == "Substantive Response Rules"
            else RuleCategory.OBJECTION
        )
        applies_key = "FI_CUSTOM" if self.detected_type == "FI" else self.detected_type
        custom = ResponseRule(
            id=f"custom_{len(self._rules) + 1}_{abs(hash(name))}",
            name=name.strip(),
            category=category,
            mode=RuleMode.INSTRUCTION if category == RuleCategory.SUBSTANTIVE else RuleMode.CONDITIONAL,
            applies_to=(applies_key,),
            description=description.strip(),
            output_text=output_text.strip(),
            enabled_by_default=True,
        )
        self._rules.append(custom)
        self._rebuild_custom_rule(custom)
        save_choice = QMessageBox.question(
            self,
            "Save Rule",
            "Save this rule globally for future cases/sessions?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if save_choice == QMessageBox.StandardButton.Yes:
            path = _global_rules_path()
            existing = load_global_rules(path)
            save_global_rules(
                path,
                existing + [ResponseRule.from_dict({**custom.to_dict(), "is_global": True})],
            )

    def _rebuild_custom_rule(self, rule: ResponseRule) -> None:
        group = QGroupBox(
            "Substantive Response Rules"
            if rule.category == RuleCategory.SUBSTANTIVE
            else "Objection Rules"
        )
        group_layout = QVBoxLayout(group)
        cb = QCheckBox(rule.name)
        cb.setChecked(True)
        cb.setToolTip(rule.description)
        group_layout.addWidget(cb)
        desc = QLabel(rule.description)
        desc.setWordWrap(True)
        desc.setStyleSheet("color: #666; margin-left: 20px;")
        group_layout.addWidget(desc)
        self._rule_checks[rule.id] = cb
        self.rule_layout.insertWidget(max(0, self.rule_layout.count() - 1), group)

    def visible_rule_count(self) -> int:
        return len(self._rules)

    def visible_rule_labels(self) -> list[str]:
        return [rule.name for rule in self._rules]

    def has_substantive_rules(self) -> bool:
        return any(rule.category == RuleCategory.SUBSTANTIVE for rule in self._rules)

    def selected_rules(self) -> list[ResponseRule]:
        checks = self._rule_checks
        return [r for r in self._rules if checks.get(r.id) and checks[r.id].isChecked()]

    def selected_rule_ids(self) -> list[str]:
        return [rule.id for rule in self.selected_rules()]

    def _rules_for_current_type(self) -> list[ResponseRule]:
        rules = built_in_rules_for(self.detected_type, fi_mode=self.fi_mode)
        if self.detected_type == "FI" and self.fi_mode == "fixed":
            return rules
        applies_key = "FI_CUSTOM" if self.detected_type == "FI" else self.detected_type
        global_rules = [
            rule
            for rule in load_global_rules(_global_rules_path())
            if applies_key in rule.applies_to or "ALL" in rule.applies_to
        ]
        return rules + global_rules

    # ------------------------------------------------------------------
    # Context and proposal generation
    # ------------------------------------------------------------------

    def _on_select_context_files(self) -> None:
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Select context file(s)",
            context_file_start_dir(self.discovery_file, self.case_root),
            "Context files (*.pdf *.docx *.txt);;All files (*.*)",
        )
        self.context_files = list(paths or [])
        self._generate_proposals()

    def _generate_proposals(self) -> None:
        self.next_btn.setEnabled(False)
        self.status_label.setText("Parsing discovery and drafting proposed responses...")
        if self.parsed_discovery is not None:
            self._kick_off_drafting(self.parsed_discovery)
            return

        self._parse_worker = DiscoveryParseWorker(
            discovery_file=self.discovery_file,
            detected_type=self.detected_type,
            parent=self,
        )
        self._parse_worker.parse_finished.connect(self._on_parse_finished)
        self._parse_worker.start()

    def _on_parse_finished(self, success: bool, payload: object) -> None:
        self._parse_worker = None
        if not success:
            self.next_btn.setEnabled(True)
            self.status_label.setText(str(payload))
            return
        self.parsed_discovery = payload  # ParsedDiscovery
        self._kick_off_drafting(self.parsed_discovery)

    def _kick_off_drafting(self, parsed) -> None:
        self.next_btn.setEnabled(True)
        response_rules = load_respond_response_rules(self.file_number)
        self._loaded_response_rules = response_rules
        context_text_by_path = {
            path: read_document_text(path)
            for path in self.context_files
            if os.path.isfile(path)
        }
        self._context_chunks = build_context_chunks(context_text_by_path)
        self.review_state = _build_pending_review_state(
            parsed, response_rules, self.fi_mode,
        )
        self._open_review_screen_with_pending(parsed)

        if self._coordinator is not None:
            self._coordinator.cancel()
        max_concurrent = LLMConfig().discovery_response_max_concurrent()
        self._coordinator = ProposalCoordinator(
            max_concurrent=max_concurrent, parent=self,
        )
        self._coordinator.proposal_ready.connect(self._on_proposal_ready)
        self._coordinator.progress.connect(self._on_coordinator_progress)
        self._coordinator.all_done.connect(self._on_coordinator_all_done)
        self._coordinator.start(
            parsed=parsed,
            selected_rules=self.selected_rules(),
            context_chunks=self._context_chunks,
            response_rules=response_rules,
            fi_mode=self.fi_mode,
        )

    def _open_review_screen_with_pending(self, parsed) -> None:
        if self.review_state is None:
            response_rules = self._loaded_response_rules or load_respond_response_rules(
                self.file_number,
            )
            self.review_state = _build_pending_review_state(
                parsed, response_rules, self.fi_mode,
            )
        self._show_review()

    def _on_proposal_ready(self, req_number: str, proposal) -> None:
        if not self.review_state or not self.parsed_discovery:
            return
        target_index = next(
            (i for i, r in enumerate(self.review_state.requests)
             if r.number == req_number),
            -1,
        )
        if target_index < 0:
            return
        previous = self.review_state.requests[target_index]
        user_edited = (
            (previous.proposed_substantive_response or "").strip() != ""
            or (previous.proposed_objections or "").strip() != ""
        )
        if user_edited:
            # Stash for the conflict banner; do not overwrite.
            previous.pending_replacement = proposal
            previous.is_pending = False
            self._refresh_review_row(target_index)
            return

        response_rules = self._loaded_response_rules or load_respond_response_rules(
            self.file_number,
        )
        _apply_proposal_to_review_state(
            review_state=self.review_state,
            req_number=req_number,
            proposal=proposal,
            parsed=self.parsed_discovery,
            selected_rules=self.selected_rules(),
            response_rules=response_rules,
            fi_mode=self.fi_mode,
        )
        self._refresh_review_row(target_index)

    def _on_coordinator_progress(self, completed: int, total: int) -> None:
        if hasattr(self, "review_status_label") and self.review_status_label:
            self.review_status_label.setText(f"Generated {completed} / {total}")

    def _on_coordinator_all_done(self) -> None:
        self._refresh_finalize_button()

    def _refresh_review_row(self, target_index: int) -> None:
        if target_index == self._current_review_index:
            self._load_current_review()
        self._refresh_finalize_button()

    def _refresh_finalize_button(self) -> None:
        if not hasattr(self, "finalize_btn"):
            return
        any_pending = self.review_state and any(
            r.is_pending for r in self.review_state.requests
        )
        self.finalize_btn.setEnabled(not any_pending)

    def _on_regenerate_clicked(self) -> None:
        review = self._current_review()
        if not review or not self.parsed_discovery or not self._coordinator:
            return
        instruction = self.regenerate_instruction_edit.text().strip()
        response_rules = self._loaded_response_rules or load_respond_response_rules(
            self.file_number,
        )
        queued = self._coordinator.regenerate(
            request_number=review.number,
            parsed=self.parsed_discovery,
            selected_rules=self.selected_rules(),
            context_chunks=self._context_chunks,
            response_rules=response_rules,
            fi_mode=self.fi_mode,
            override_instruction=instruction,
        )
        if not queued:
            return
        review.is_pending = True
        review.review_reason = "Regenerating..."
        review.needs_review = False
        review.approved = False
        review.pending_replacement = None
        self.regenerate_instruction_edit.clear()
        self._load_current_review()
        self._refresh_finalize_button()

    def _view_pending_replacement(self) -> None:
        review = self._current_review()
        if not review or not review.pending_replacement:
            return
        QMessageBox.information(
            self,
            f"New draft for Request No. {review.number}",
            review.pending_replacement.proposed_substantive_response
            or "(no substantive text in draft)",
        )

    def _apply_pending_replacement(self) -> None:
        review = self._current_review()
        if not review or not review.pending_replacement:
            return
        response_rules = self._loaded_response_rules or load_respond_response_rules(
            self.file_number,
        )
        _apply_proposal_to_review_state(
            review_state=self.review_state,
            req_number=review.number,
            proposal=review.pending_replacement,
            parsed=self.parsed_discovery,
            selected_rules=self.selected_rules(),
            response_rules=response_rules,
            fi_mode=self.fi_mode,
        )
        # _apply_proposal_to_review_state replaces the row; re-fetch.
        new_review = self._current_review()
        if new_review:
            new_review.pending_replacement = None
        self._load_current_review()

    def _discard_pending_replacement(self) -> None:
        review = self._current_review()
        if not review:
            return
        review.pending_replacement = None
        self._load_current_review()

    # ------------------------------------------------------------------
    # Review screen
    # ------------------------------------------------------------------

    def _build_review_widget(self) -> QWidget:
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)

        self.review_title = QLabel("Review Discovery Responses")
        self.review_title.setStyleSheet("font-size: 14px; font-weight: 600;")
        layout.addWidget(self.review_title)

        self.request_label = QLabel("")
        self.request_label.setWordWrap(True)
        self.request_label.setStyleSheet("font-weight: 600;")
        layout.addWidget(self.request_label)

        self.request_text = QLabel("")
        self.request_text.setWordWrap(True)
        self.request_text.setTextInteractionFlags(Qt.TextSelectableByMouse)
        layout.addWidget(self.request_text)

        self.review_warning_label = QLabel("")
        self.review_warning_label.setWordWrap(True)
        self.review_warning_label.setStyleSheet("color: #8a5a00; font-weight: 600;")
        self.review_warning_label.hide()
        layout.addWidget(self.review_warning_label)

        layout.addWidget(QLabel("Objections"))
        self.objection_edit = QPlainTextEdit()
        self.objection_edit.setMinimumHeight(115)
        layout.addWidget(self.objection_edit)

        layout.addWidget(QLabel("Substantive Response"))
        self.response_edit = QPlainTextEdit()
        self.response_edit.setMinimumHeight(115)
        layout.addWidget(self.response_edit)

        quick_row = QHBoxLayout()

        quick_group = QGroupBox("Quick Objections")
        quick_layout = QVBoxLayout(quick_group)
        for quick_rule in get_quick_objection_rules():
            cb = QCheckBox(quick_rule.name)
            cb.toggled.connect(
                lambda checked, rule=quick_rule: self._on_quick_rule_toggled(rule, checked)
            )
            quick_layout.addWidget(cb)
            self._quick_checks[quick_rule.id] = cb
        quick_row.addWidget(quick_group, 1)

        quick_response_group = QGroupBox("Quick Responses")
        quick_response_layout = QVBoxLayout(quick_response_group)
        for label, text in self._quick_response_options():
            btn = QPushButton(label)
            btn.clicked.connect(
                lambda _checked=False, response_text=text: self._insert_quick_response(response_text)
            )
            quick_response_layout.addWidget(btn)
            self._quick_response_buttons[label] = btn
        quick_response_layout.addStretch(1)
        quick_row.addWidget(quick_response_group, 1)
        layout.addLayout(quick_row)

        # Conflict banner — shown when a new draft arrives for a row the
        # user has already edited.
        self.conflict_banner = QWidget()
        banner_layout = QHBoxLayout(self.conflict_banner)
        banner_layout.setContentsMargins(8, 6, 8, 6)
        self.conflict_banner.setStyleSheet(
            "background-color: #fff7d6; border: 1px solid #d4b85a; border-radius: 4px;"
        )
        banner_label = QLabel("New draft available — your edits are preserved.")
        banner_label.setStyleSheet("font-weight: 600;")
        banner_layout.addWidget(banner_label, 1)
        self.conflict_view_btn = QPushButton("View")
        self.conflict_view_btn.clicked.connect(self._view_pending_replacement)
        self.conflict_apply_btn = QPushButton("Apply")
        self.conflict_apply_btn.clicked.connect(self._apply_pending_replacement)
        self.conflict_discard_btn = QPushButton("Discard")
        self.conflict_discard_btn.clicked.connect(self._discard_pending_replacement)
        for btn in (self.conflict_view_btn, self.conflict_apply_btn, self.conflict_discard_btn):
            banner_layout.addWidget(btn)
        self.conflict_banner.hide()
        layout.addWidget(self.conflict_banner)

        # Regenerate row.
        regenerate_row = QHBoxLayout()
        self.regenerate_instruction_edit = QLineEdit()
        self.regenerate_instruction_edit.setPlaceholderText(
            "Optional instructions for regeneration (e.g., lean harder on privilege)"
        )
        regenerate_row.addWidget(self.regenerate_instruction_edit, 1)
        self.regenerate_btn = QPushButton("Regenerate")
        self.regenerate_btn.clicked.connect(self._on_regenerate_clicked)
        regenerate_row.addWidget(self.regenerate_btn)
        layout.addLayout(regenerate_row)

        self.review_status_label = QLabel("")
        self.review_status_label.setStyleSheet("color: #666; font-style: italic;")
        layout.addWidget(self.review_status_label)

        nav = QHBoxLayout()
        self.prev_btn = QPushButton("Previous")
        self.prev_btn.clicked.connect(lambda: self._move_review(-1))
        nav.addWidget(self.prev_btn)
        nav.addStretch()
        self.next_review_btn = QPushButton("Next")
        self.next_review_btn.clicked.connect(lambda: self._move_review(1))
        nav.addWidget(self.next_review_btn)
        self.finalize_btn = QPushButton("Finalize Responses")
        self.finalize_btn.clicked.connect(self._emit_run_requested)
        nav.addWidget(self.finalize_btn)
        layout.addLayout(nav)
        return widget

    def _show_review(self) -> None:
        self.rules_widget.hide()
        self.review_widget.show()
        self._current_review_index = 0
        self._load_current_review()
        self._refresh_finalize_button()

    def _current_review(self) -> RequestReview | None:
        if not self.review_state or not self.review_state.requests:
            return None
        return self.review_state.requests[self._current_review_index]

    def _save_current_review_edits(self) -> None:
        review = self._current_review()
        if not review:
            return
        review.proposed_objections = self.objection_edit.toPlainText().strip()
        review.proposed_substantive_response = self.response_edit.toPlainText().strip()

    def _load_current_review(self) -> None:
        review = self._current_review()
        count = len(self.review_state.requests) if self.review_state else 0
        if not review:
            self.request_label.setText("No parsed requests.")
            return
        icon = self._status_icon_for(review)
        self.review_title.setText(
            f"Review Discovery Responses ({self._current_review_index + 1} of {count})"
        )
        self.request_label.setText(f"{icon} Request No. {review.number}".strip())
        self.request_text.setText(review.request_text)

        is_pending = bool(review.is_pending)
        self.objection_edit.setPlainText(
            "" if is_pending else review.proposed_objections
        )
        self.response_edit.setPlainText(
            "" if is_pending else review.proposed_substantive_response
        )
        self.objection_edit.setReadOnly(is_pending)
        self.response_edit.setReadOnly(is_pending)
        self.objection_edit.setPlaceholderText("Generating..." if is_pending else "")
        self.response_edit.setPlaceholderText("Generating..." if is_pending else "")

        review_reason = (review.review_reason or "").strip()
        if is_pending:
            self.review_warning_label.setText("Generating...")
            self.review_warning_label.show()
        elif review.needs_review:
            warning_text = (
                f"Needs review: {review_reason}" if review_reason else "Needs review."
            )
            self.review_warning_label.setText(warning_text)
            self.review_warning_label.show()
        else:
            self.review_warning_label.setText("")
            self.review_warning_label.hide()

        for rule_id, cb in self._quick_checks.items():
            cb.blockSignals(True)
            cb.setChecked(rule_id in review.selected_quick_objection_ids)
            cb.blockSignals(False)
        self.prev_btn.setEnabled(self._current_review_index > 0)
        self.next_review_btn.setEnabled(self._current_review_index < count - 1)

        has_pending_replacement = review.pending_replacement is not None
        if hasattr(self, "conflict_banner"):
            self.conflict_banner.setVisible(has_pending_replacement)

    def _status_icon_for(self, review) -> str:
        if review.is_pending:
            return "⏳"  # hourglass U+23F3
        if review.approved:
            return "✓"  # check U+2713
        if review.needs_review:
            return "⚠"  # warning U+26A0
        # "Edited" = non-default content but not yet approved.
        has_content = bool(
            (review.proposed_objections or "").strip()
            or (review.proposed_substantive_response or "").strip()
        )
        if has_content:
            return "✏"  # pencil U+270F
        return ""

    def _move_review(self, delta: int) -> None:
        if not self.review_state:
            return
        self._save_current_review_edits()
        if delta > 0:
            review = self._current_review()
            if review:
                review.approved = True
        self._current_review_index = max(
            0,
            min(len(self.review_state.requests) - 1, self._current_review_index + delta),
        )
        self._load_current_review()

    def _on_quick_rule_toggled(self, rule: ResponseRule, checked: bool) -> None:
        review = self._current_review()
        if not review:
            return
        review.proposed_objections = self.objection_edit.toPlainText().strip()
        term = ""
        if "{term}" in rule.output_text:
            parsed_req = self._parsed_request_for_review(review)
            if parsed_req and parsed_req.defined_terms_used:
                term = parsed_req.defined_terms_used[0]
            else:
                term, ok = QInputDialog.getText(self, "Undefined Term", "Term:")
                if not ok or not term.strip():
                    self._quick_checks[rule.id].blockSignals(True)
                    self._quick_checks[rule.id].setChecked(False)
                    self._quick_checks[rule.id].blockSignals(False)
                    return
        if checked:
            insert_quick_objection(review, rule, term=term.strip() or None)
        else:
            remove_quick_objection(review, rule, term=term.strip() or None)
        self.objection_edit.setPlainText(review.proposed_objections)

    def quick_response_labels(self) -> list[str]:
        return list(self._quick_response_buttons)

    def quick_response_button(self, label: str) -> QPushButton:
        return self._quick_response_buttons[label]

    def _quick_response_options(self) -> list[tuple[str, str]]:
        dtype = (self.detected_type or "").upper()
        if dtype in {"FI", "SI"}:
            return [("Refer to Document", REFER_TO_DOCUMENT_RESPONSE)]
        if dtype == "RPD":
            return [
                ("Will produce", WILL_PRODUCE_RESPONSE),
                ("Wont produce", WONT_PRODUCE_RESPONSE),
            ]
        if dtype == "RFA":
            return [
                ("Admit", "Admit."),
                ("Deny", "Deny."),
                ("Cant Admit/Deny", CANT_ADMIT_DENY_RESPONSE),
            ]
        return []

    def _insert_quick_response(self, response_text: str) -> None:
        review = self._current_review()
        if review:
            review.proposed_substantive_response = response_text
        self.response_edit.setPlainText(response_text)

    def _parsed_request_for_review(self, review: RequestReview) -> ParsedRequest | None:
        if not self.parsed_discovery:
            return None
        for req in self.parsed_discovery.requests:
            if req.number == review.number:
                return req
        return None

    # ------------------------------------------------------------------
    # Persistence and final handoff
    # ------------------------------------------------------------------

    def to_dict(self) -> dict:
        self._save_current_review_edits()
        return {
            "discovery_file": self.discovery_file,
            "detected_type": self.detected_type,
            "fi_mode": self.fi_mode,
            "context_files": list(self.context_files),
            "selected_rule_ids": self.selected_rule_ids(),
            "parsed_discovery": parsed_discovery_to_dict(self.parsed_discovery),
            "review_state": self.review_state.to_dict() if self.review_state else None,
            "save_default_dir": response_save_default_dir(self.case_root),
            "suggested_filename": suggested_response_filename(self.parsed_discovery),
        }

    def from_dict(self, data: dict) -> None:
        if not data:
            return
        self.fi_mode = "fixed" if data.get("fi_mode", self.fi_mode) == "fixed" else "custom"
        self.context_files = list(data.get("context_files", []))
        self.parsed_discovery = parsed_discovery_from_dict(data.get("parsed_discovery"))
        if data.get("review_state"):
            self.review_state = ReviewState.from_dict(data["review_state"])
        self._sync_fi_mode_controls()
        self._rebuild_rules()
        selected = set(data.get("selected_rule_ids", []))
        if selected:
            for rule_id, cb in self._rule_checks.items():
                cb.setChecked(rule_id in selected)
        if self.review_state:
            self._show_review()

    def _emit_run_requested(self) -> None:
        self._save_current_review_edits()
        review = self._current_review()
        if review:
            review.approved = True
        if not self.review_state or not self.review_state.requests:
            QMessageBox.warning(self, "No responses", "Generate responses before finalizing.")
            return
        if not self.review_state.all_approved():
            QMessageBox.warning(
                self,
                "Approval required",
                "Every discovery request must be approved before final assembly.",
            )
            return
        self.run_requested.emit(self.to_dict())
class RespondDiscoveryWorker(QThread):
    progress = Signal(str)
    warning = Signal(str)
    finished_result = Signal(bool, str)

    def __init__(self, case_path: str, file_number: str, settings: dict, parent=None):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.settings = dict(settings or {})

    def validate_settings(self) -> tuple[bool, str]:
        review_state = ReviewState.from_dict(self.settings.get("review_state") or {})
        if not review_state.requests:
            return False, "No reviewed discovery responses are available."
        if not review_state.all_approved():
            return False, "Every discovery request must be approved before final assembly."
        if not self.settings.get("parsed_discovery"):
            return False, "Parsed discovery data is missing."
        return True, ""

    def run(self) -> None:
        ok, message = self.validate_settings()
        if not ok:
            self.finished_result.emit(False, message)
            return
        try:
            from icharlotte_core.discovery.assembler import DiscoveryAssembler
            from icharlotte_core.discovery.response_assembler import ResponseAssembler
            from icharlotte_core.word_validator import (
                validate_discovery_response_docx,
            )

            parsed = parsed_discovery_from_dict(self.settings.get("parsed_discovery"))
            review_state = ReviewState.from_dict(self.settings.get("review_state") or {})
            rules = load_respond_response_rules(self.file_number)
            response_text = review_state.to_plain_text(parsed, rules)

            caption_path = DiscoveryAssembler.find_caption_page(self.case_path)
            if not caption_path:
                self.finished_result.emit(False, "No caption page found in case folder.")
                return

            output_path = preview_response_output_path(self.case_path, parsed)

            self.progress.emit("Assembling Word response...")
            assembler = ResponseAssembler(caption_path)
            assembler.assemble(
                parsed=parsed,
                response_text=response_text,
                rules=rules,
                output_path=output_path,
            )

            self.progress.emit("Validating Word response...")
            validation = validate_discovery_response_docx(output_path)
            if getattr(validation, "has_errors", False):
                self.finished_result.emit(
                    False,
                    "Word validation failed. Review the generated response before use.",
                )
                return

            self.finished_result.emit(True, output_path)
        except Exception as exc:
            self.finished_result.emit(False, str(exc))


def _response_abbreviation(responding_party: str) -> str:
    words = (responding_party or "").replace(",", "").split()
    skip = {"defendant", "plaintiff", "cross-defendant", "cross-complainant"}
    return next((word for word in words if word.lower() not in skip), "Def")


def _global_rules_path() -> str:
    project_root = os.path.dirname(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
    )
    return os.path.join(project_root, "config", "response_wizard_rules.json")


def load_respond_response_rules(file_number: str) -> ResponseRules:
    """Load the same per-case response rules used by the advanced Respond tab."""
    if not file_number:
        return ResponseRules()
    try:
        from Scripts.case_data_manager import CaseDataManager

        rules_data = CaseDataManager().get_value(file_number, "respond_rules")
        if isinstance(rules_data, dict):
            return ResponseRules.from_dict(rules_data)
    except Exception:
        pass
    return ResponseRules()


def _normalize_and_filter_parsed_discovery(
    parsed: ParsedDiscovery,
    detected_type: str,
    discovery_file: str,
) -> ParsedDiscovery:
    """Canonicalize discovery type and keep only checked flattened FROG items."""
    normalized_detected = normalize_discovery_type(detected_type)
    parsed.discovery_type = normalized_detected or normalize_discovery_type(
        parsed.discovery_type
    )
    if parsed.discovery_type != "FI":
        return parsed
    selected_numbers = extract_selected_form_interrogatory_numbers(discovery_file)
    return complete_selected_form_interrogatories(
        parsed,
        discovery_file,
        selected_numbers,
    )


