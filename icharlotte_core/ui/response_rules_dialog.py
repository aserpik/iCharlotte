"""
Response Rules Dialog — 3-tab editor for discovery response configuration.

Tab 1 (Strategy): Radio buttons for aggressiveness, checkboxes for always-include,
                   radio buttons for SI style, RFA posture, RPD posture.
Tab 2 (Boilerplate): Editable QTextEdit fields for each boilerplate block,
                      with per-field Reset buttons.
Tab 3 (Custom Instructions): Large QTextEdit for free-form instructions.
"""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QGroupBox, QRadioButton, QCheckBox, QTextEdit, QPushButton,
    QLabel, QScrollArea, QDialogButtonBox, QSplitter,
)

from icharlotte_core.discovery.response_rules import ResponseRules


class ResponseRulesDialog(QDialog):
    """Three-tab dialog for editing ResponseRules strategy and boilerplate."""

    def __init__(self, rules: ResponseRules, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Response Rules")
        self.setMinimumSize(720, 600)
        self._defaults = ResponseRules()

        # Main layout
        main_layout = QVBoxLayout(self)

        # Tab widget
        self.tab_widget = QTabWidget()
        main_layout.addWidget(self.tab_widget)

        # Build tabs
        self.tab_widget.addTab(self._build_strategy_tab(), "Strategy")
        self.tab_widget.addTab(self._build_boilerplate_tab(), "Boilerplate")
        self.tab_widget.addTab(self._build_custom_tab(), "Custom Instructions")

        # Bottom button row: Reset All + OK/Cancel
        bottom_layout = QHBoxLayout()
        self._reset_btn = QPushButton("Reset All to Defaults")
        self._reset_btn.clicked.connect(self._on_reset_defaults)
        bottom_layout.addWidget(self._reset_btn)
        bottom_layout.addStretch()

        self._button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        self._button_box.accepted.connect(self.accept)
        self._button_box.rejected.connect(self.reject)
        bottom_layout.addWidget(self._button_box)
        main_layout.addLayout(bottom_layout)

        # Populate UI from provided rules
        self._load_rules(rules)

    # ------------------------------------------------------------------
    # Tab builders
    # ------------------------------------------------------------------

    def _build_strategy_tab(self) -> QWidget:
        """Build Tab 1: Strategy controls."""
        widget = QWidget()
        outer = QVBoxLayout(widget)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        outer.addWidget(scroll)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(10)
        scroll.setWidget(container)

        # --- Objection Aggressiveness ---
        agg_group = QGroupBox("Objection Aggressiveness")
        agg_layout = QVBoxLayout(agg_group)
        self.rb_aggressive = QRadioButton("Aggressive — include all possible objections")
        self.rb_moderate_obj = QRadioButton("Moderate — include reasonably applicable objections")
        self.rb_conservative = QRadioButton("Conservative — include only clearly applicable objections")
        agg_layout.addWidget(self.rb_aggressive)
        agg_layout.addWidget(self.rb_moderate_obj)
        agg_layout.addWidget(self.rb_conservative)
        layout.addWidget(agg_group)

        # --- Always Include ---
        always_group = QGroupBox("Always Include")
        always_layout = QVBoxLayout(always_group)
        self.cb_privacy = QCheckBox("Privacy objection")
        self.cb_privilege = QCheckBox("Attorney-client / work-product privilege")
        self.cb_burden = QCheckBox("Burden / oppression objection")
        always_layout.addWidget(self.cb_privacy)
        always_layout.addWidget(self.cb_privilege)
        always_layout.addWidget(self.cb_burden)
        layout.addWidget(always_group)

        # --- Auto-Detection ---
        auto_group = QGroupBox("Auto-Detection")
        auto_layout = QVBoxLayout(auto_group)
        self.cb_compound = QCheckBox("Flag compound questions")
        self.cb_broad_defs = QCheckBox("Flag broad definitions")
        auto_layout.addWidget(self.cb_compound)
        auto_layout.addWidget(self.cb_broad_defs)
        layout.addWidget(auto_group)

        # --- SI Response Style ---
        si_group = QGroupBox("SI Response Style")
        si_layout = QVBoxLayout(si_group)
        self.rb_si_minimal = QRadioButton("Minimal — narrow, vague when possible")
        self.rb_si_moderate = QRadioButton("Moderate — standard responsive")
        self.rb_si_detailed = QRadioButton("Detailed — full responsive")
        si_layout.addWidget(self.rb_si_minimal)
        si_layout.addWidget(self.rb_si_moderate)
        si_layout.addWidget(self.rb_si_detailed)
        layout.addWidget(si_group)

        # --- RFA Default Posture ---
        rfa_group = QGroupBox("RFA Default Posture")
        rfa_layout = QVBoxLayout(rfa_group)
        self.rb_rfa_cautious = QRadioButton("Cautious — lean toward Deny/Insufficient")
        self.rb_rfa_balanced = QRadioButton("Balanced — best judgment")
        self.rb_rfa_cooperative = QRadioButton("Cooperative — admit undisputed facts")
        rfa_layout.addWidget(self.rb_rfa_cautious)
        rfa_layout.addWidget(self.rb_rfa_balanced)
        rfa_layout.addWidget(self.rb_rfa_cooperative)
        layout.addWidget(rfa_group)

        # --- RPD Default Posture ---
        rpd_group = QGroupBox("RPD Default Posture")
        rpd_layout = QVBoxLayout(rpd_group)
        self.rb_rpd_unable = QRadioButton("Unable to Comply (default)")
        self.rb_rpd_comply = QRadioButton("Will Comply")
        self.rb_rpd_context = QRadioButton("Context-Dependent")
        rpd_layout.addWidget(self.rb_rpd_unable)
        rpd_layout.addWidget(self.rb_rpd_comply)
        rpd_layout.addWidget(self.rb_rpd_context)
        layout.addWidget(rpd_group)

        layout.addStretch()
        return widget

    def _build_boilerplate_tab(self) -> QWidget:
        """Build Tab 2: Boilerplate text editors with per-field reset buttons."""
        widget = QWidget()
        outer = QVBoxLayout(widget)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        outer.addWidget(scroll)

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setSpacing(8)
        scroll.setWidget(container)

        # Define (attribute_name, label, default_getter) tuples
        fields_spec = [
            ("waiver_edit",       "Waiver Language",               lambda d: d.waiver_language),
            ("reservation_edit",  "Reservation Clause",            lambda d: d.reservation_clause),
            ("prelim_fi_edit",    "Preliminary Statement (FI)",     lambda d: d.preliminary_statement_fi),
            ("prelim_si_edit",    "Preliminary Statement (SI)",     lambda d: d.preliminary_statement_si),
            ("prelim_rfa_edit",   "Preliminary Statement (RFA)",    lambda d: d.preliminary_statement_rfa),
            ("prelim_rpd_edit",   "Preliminary Statement (RPD)",    lambda d: d.preliminary_statement_rpd),
            ("intro_fi_edit",     "Introduction Template (FI)",     lambda d: d.intro_template_fi),
            ("intro_si_edit",     "Introduction Template (SI)",     lambda d: d.intro_template_si),
            ("intro_rfa_edit",    "Introduction Template (RFA)",    lambda d: d.intro_template_rfa),
            ("intro_rpd_edit",    "Introduction Template (RPD)",    lambda d: d.intro_template_rpd),
            ("gen_obj_rfa_edit",  "General Objections (RFA)",       lambda d: d.general_objections_rfa),
            ("gen_obj_rpd_edit",  "General Objections (RPD)",       lambda d: d.general_objections_rpd),
            ("verification_edit", "Verification Template",          lambda d: d.verification_template),
        ]

        for attr, label, default_fn in fields_spec:
            group, edit = self._make_boilerplate_field(label, default_fn)
            setattr(self, attr, edit)
            layout.addWidget(group)

        layout.addStretch()
        return widget

    def _make_boilerplate_field(self, label: str, default_fn) -> tuple:
        """Return (QGroupBox, QTextEdit) for a single boilerplate field."""
        group = QGroupBox(label)
        g_layout = QVBoxLayout(group)

        edit = QTextEdit()
        edit.setMaximumHeight(120)
        g_layout.addWidget(edit)

        reset_btn = QPushButton("Reset to Default")
        reset_btn.setMaximumWidth(150)

        def _reset(checked=False, _edit=edit, _fn=default_fn):
            _edit.setPlainText(_fn(self._defaults))

        reset_btn.clicked.connect(_reset)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_row.addWidget(reset_btn)
        g_layout.addLayout(btn_row)

        return group, edit

    def _build_custom_tab(self) -> QWidget:
        """Build Tab 3: Free-form custom instructions."""
        widget = QWidget()
        layout = QVBoxLayout(widget)

        lbl = QLabel(
            "Custom instructions are appended to every LLM prompt when drafting responses. "
            "Use this to specify jurisdiction-specific rules, client preferences, or style guidelines."
        )
        lbl.setWordWrap(True)
        layout.addWidget(lbl)

        self.custom_instructions_edit = QTextEdit()
        self.custom_instructions_edit.setPlaceholderText(
            "e.g. Always object to any request asking for mental health records. "
            "Use passive voice in all denials. Client prefers abbreviated preliminary statements."
        )
        layout.addWidget(self.custom_instructions_edit)

        return widget

    # ------------------------------------------------------------------
    # Load / save
    # ------------------------------------------------------------------

    def _load_rules(self, rules: ResponseRules) -> None:
        """Populate all UI widgets from a ResponseRules instance."""
        # Objection aggressiveness
        self.rb_aggressive.setChecked(rules.objection_aggressiveness == "aggressive")
        self.rb_moderate_obj.setChecked(rules.objection_aggressiveness == "moderate")
        self.rb_conservative.setChecked(rules.objection_aggressiveness == "conservative")

        # Always-include checkboxes
        self.cb_privacy.setChecked(rules.always_include_privacy_objection)
        self.cb_privilege.setChecked(rules.always_include_privilege_objection)
        self.cb_burden.setChecked(rules.always_include_burden_objection)

        # Auto-detection checkboxes
        self.cb_compound.setChecked(rules.auto_flag_compound)
        self.cb_broad_defs.setChecked(rules.auto_flag_broad_definitions)

        # SI response style
        self.rb_si_minimal.setChecked(rules.si_response_style == "minimal")
        self.rb_si_moderate.setChecked(rules.si_response_style == "moderate")
        self.rb_si_detailed.setChecked(rules.si_response_style == "detailed")

        # RFA posture — model values: cautious, cooperative, deny_all; 'balanced' is dialog-only alias
        self.rb_rfa_cautious.setChecked(rules.rfa_default_posture in ("cautious", "deny_all"))
        self.rb_rfa_balanced.setChecked(rules.rfa_default_posture == "balanced")
        self.rb_rfa_cooperative.setChecked(rules.rfa_default_posture == "cooperative")

        # RPD posture — model values: context_dependent, produce_all, withhold_pending_protective
        # Dialog adds 'unable_to_comply' and 'will_comply' as aliases
        self.rb_rpd_unable.setChecked(
            rules.rpd_default_posture in ("context_dependent", "unable_to_comply")
        )
        self.rb_rpd_comply.setChecked(
            rules.rpd_default_posture in ("produce_all", "will_comply")
        )
        self.rb_rpd_context.setChecked(
            rules.rpd_default_posture == "withhold_pending_protective"
        )

        # Boilerplate fields
        self.waiver_edit.setPlainText(rules.waiver_language)
        self.reservation_edit.setPlainText(rules.reservation_clause)
        self.prelim_fi_edit.setPlainText(rules.preliminary_statement_fi)
        self.prelim_si_edit.setPlainText(rules.preliminary_statement_si)
        self.prelim_rfa_edit.setPlainText(rules.preliminary_statement_rfa)
        self.prelim_rpd_edit.setPlainText(rules.preliminary_statement_rpd)
        self.intro_fi_edit.setPlainText(rules.intro_template_fi)
        self.intro_si_edit.setPlainText(rules.intro_template_si)
        self.intro_rfa_edit.setPlainText(rules.intro_template_rfa)
        self.intro_rpd_edit.setPlainText(rules.intro_template_rpd)
        self.gen_obj_rfa_edit.setPlainText(rules.general_objections_rfa)
        self.gen_obj_rpd_edit.setPlainText(rules.general_objections_rpd)
        self.verification_edit.setPlainText(rules.verification_template)

        # Custom instructions
        self.custom_instructions_edit.setPlainText(rules.custom_instructions)

    def get_rules(self) -> ResponseRules:
        """Read all UI state and return a new ResponseRules instance."""
        # Objection aggressiveness
        if self.rb_aggressive.isChecked():
            aggressiveness = "aggressive"
        elif self.rb_moderate_obj.isChecked():
            aggressiveness = "moderate"
        else:
            aggressiveness = "conservative"

        # SI style
        if self.rb_si_minimal.isChecked():
            si_style = "minimal"
        elif self.rb_si_moderate.isChecked():
            si_style = "moderate"
        else:
            si_style = "detailed"

        # RFA posture
        if self.rb_rfa_cooperative.isChecked():
            rfa_posture = "cooperative"
        elif self.rb_rfa_balanced.isChecked():
            rfa_posture = "balanced"
        else:
            rfa_posture = "cautious"

        # RPD posture
        if self.rb_rpd_comply.isChecked():
            rpd_posture = "will_comply"
        elif self.rb_rpd_context.isChecked():
            rpd_posture = "withhold_pending_protective"
        else:
            rpd_posture = "unable_to_comply"

        return ResponseRules(
            # Strategy
            objection_aggressiveness=aggressiveness,
            always_include_privacy_objection=self.cb_privacy.isChecked(),
            always_include_privilege_objection=self.cb_privilege.isChecked(),
            always_include_burden_objection=self.cb_burden.isChecked(),
            auto_flag_compound=self.cb_compound.isChecked(),
            auto_flag_broad_definitions=self.cb_broad_defs.isChecked(),
            si_response_style=si_style,
            rfa_default_posture=rfa_posture,
            rpd_default_posture=rpd_posture,
            # Boilerplate
            waiver_language=self.waiver_edit.toPlainText(),
            reservation_clause=self.reservation_edit.toPlainText(),
            preliminary_statement_fi=self.prelim_fi_edit.toPlainText(),
            preliminary_statement_si=self.prelim_si_edit.toPlainText(),
            preliminary_statement_rfa=self.prelim_rfa_edit.toPlainText(),
            preliminary_statement_rpd=self.prelim_rpd_edit.toPlainText(),
            intro_template_fi=self.intro_fi_edit.toPlainText(),
            intro_template_si=self.intro_si_edit.toPlainText(),
            intro_template_rfa=self.intro_rfa_edit.toPlainText(),
            intro_template_rpd=self.intro_rpd_edit.toPlainText(),
            general_objections_rfa=self.gen_obj_rfa_edit.toPlainText(),
            general_objections_rpd=self.gen_obj_rpd_edit.toPlainText(),
            verification_template=self.verification_edit.toPlainText(),
            # Custom
            custom_instructions=self.custom_instructions_edit.toPlainText(),
        )

    def _on_reset_defaults(self) -> None:
        """Reset all fields to factory defaults."""
        self._load_rules(ResponseRules())
