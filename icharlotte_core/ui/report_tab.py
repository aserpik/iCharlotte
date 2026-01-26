import os
import sys
import json
import shutil
import time
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
    QListWidget, QSplitter, QFrame, QFileDialog, QMessageBox,
    QDialog, QFormLayout, QLineEdit, QComboBox, QGroupBox,
    QStackedWidget, QListWidgetItem, QAbstractItemView, QMenu,
    QCheckBox, QPlainTextEdit, QTableWidget, QTableWidgetItem, QHeaderView,
    QWidgetAction
)
from PySide6.QtCore import Qt, QUrl, QTimer, QSize, Signal
from PySide6.QtGui import QAction, QIcon, QDragEnterEvent, QDropEvent
from PySide6.QtWebEngineWidgets import QWebEngineView
from PySide6.QtWebEngineCore import QWebEnginePage

from icharlotte_core.config import SCRIPTS_DIR, GEMINI_DATA_DIR, TEMP_DIR
from icharlotte_core.ui.widgets import StatusWidget, AgentRunner
from icharlotte_core.ui.logs_tab import LogManager

PRESETS = {
    "main_heading": {
        "name": "Format Main Headings",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": "^(FACTUAL BACKGROUND|PROCEDURAL HISTORY|LEGAL ANALYSIS|INTRODUCTION|CONCLUSION|[A-Z][A-Z ]{3,})$",
            "case_sensitive": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "font_name": "Times New Roman",
                "font_size": 12,
                "font_bold": True,
                "alignment": "left",
                "space_after": 0
            }
        }
    },
    "subheading_a": {
        "name": "Level A Heading (A., B., C.)",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": "^[A-Z]\\.\\s+.*",
            "case_sensitive": True,
            "is_list": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "font_name": "Times New Roman",
                "font_size": 12,
                "font_bold": True,
                "font_italic": False,
                "left_indent": 1.0,
                "first_line_indent": -0.5,
                "dynamic_properties": {
                    "Range.Font.Underline": 1
                }
            }
        }
    },
    "subheading_1": {
        "name": "Level 1 Heading (1., 2., 3.)",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": "^\\d+\\.\\s+.*",
            "case_sensitive": True,
            "is_list": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "font_name": "Times New Roman",
                "font_size": 12,
                "font_bold": True,
                "font_italic": False,
                "left_indent": 1.5,
                "first_line_indent": -0.5,
                "dynamic_properties": {
                    "Range.Font.Underline": 1
                }
            }
        }
    },
    "bullet": {
        "name": "Format Bullet Points",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": ".*",
            "is_list": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "left_indent": 1.0,
                "first_line_indent": -0.5,
                "space_after": 6
            }
        }
    },
    "narrative": {
        "name": "Format Narrative Text",
        "enabled": True,
        "trigger": {
            "scope": "paragraph",
            "match_type": "regex",
            "pattern": "^(?!FACTUAL|PROCEDURAL|LEGAL|INTRODUCTION|CONCLUSION|[A-Z]\\.|\\d+\\.|[\\u2022\\-o]).{10,}",
            "case_sensitive": True
        },
        "action": {
            "type": "format",
            "formatting": {
                "font_name": "Times New Roman",
                "font_size": 12,
                "alignment": "left",
                "first_line_indent": 0.5,
                "space_after": 0
            }
        }
    }
}

class RuleDialog(QDialog):
    def __init__(self, rule=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Rule Editor")
        self.resize(550, 650)
        self.rule = rule or {}
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QVBoxLayout(self)

        # --- Top Toolbar for AI/Automation ---
        auto_layout = QHBoxLayout()
        
        preset_btn = QPushButton("📋 Load Preset")
        preset_btn.setStyleSheet("background-color: #009688; color: white; font-weight: bold;")
        preset_btn.clicked.connect(self.apply_preset)
        auto_layout.addWidget(preset_btn)
        
        ai_build_btn = QPushButton("✨ AI Rule Builder")
        ai_build_btn.setStyleSheet("background-color: #673AB7; color: white; font-weight: bold;")
        ai_build_btn.clicked.connect(self.ai_build_rule)
        
        self.ai_model_combo = QComboBox()
        self.ai_model_combo.addItems([
            "models/gemini-3-flash-preview", 
            "models/gemini-3-pro-preview", 
            "models/gemini-2.5-flash",
            "models/gemini-1.5-pro",
            "models/gemini-1.5-flash"
        ])
        self.ai_model_combo.setToolTip("Model for AI Rule Builder and Pattern Generator")
        
        word_sync_btn = QPushButton("📝 Get from Word Selection")
        word_sync_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold;")
        word_sync_btn.clicked.connect(self.get_from_word)
        
        auto_layout.addWidget(ai_build_btn)
        auto_layout.addWidget(self.ai_model_combo)
        auto_layout.addSpacing(10)
        auto_layout.addWidget(word_sync_btn)
        layout.addLayout(auto_layout)

        # --- General Info ---
        form = QFormLayout()
        self.name_edit = QLineEdit()
        form.addRow("Rule Name:", self.name_edit)
        layout.addLayout(form)

        # --- Trigger Section ---
        trigger_group = QGroupBox("Trigger (Where to apply?)")
        trigger_layout = QFormLayout(trigger_group)
        
        # Natural Language Trigger Input
        nl_trig_layout = QVBoxLayout()
        nl_trig_row = QHBoxLayout()
        self.nl_trig_input = QLineEdit()
        self.nl_trig_input.setPlaceholderText("Describe matching criteria (e.g. 'Bold paragraphs starting with Note')")
        self.nl_trig_btn = QPushButton("✨ Auto-Fill Trigger")
        self.nl_trig_btn.clicked.connect(self.generate_trigger_config)
        nl_trig_row.addWidget(self.nl_trig_input)
        nl_trig_row.addWidget(self.nl_trig_btn)
        nl_trig_layout.addLayout(nl_trig_row)
        
        # Trigger Property Table (Hidden by default or shown?)
        self.trig_prop_label = QLabel("Property Filters (Word Object Model):")
        self.trig_prop_table = QTableWidget()
        self.trig_prop_table.setColumnCount(2)
        self.trig_prop_table.setHorizontalHeaderLabels(["Property Path", "Value"])
        self.trig_prop_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.trig_prop_table.setFixedHeight(100)
        
        trig_prop_btn_row = QHBoxLayout()
        add_tp_btn = QPushButton("+")
        add_tp_btn.setFixedWidth(30)
        add_tp_btn.clicked.connect(lambda: self.trig_prop_table.insertRow(self.trig_prop_table.rowCount()))
        del_tp_btn = QPushButton("-")
        del_tp_btn.setFixedWidth(30)
        del_tp_btn.clicked.connect(lambda: self.trig_prop_table.removeRow(self.trig_prop_table.currentRow()))
        trig_prop_btn_row.addWidget(self.trig_prop_label)
        trig_prop_btn_row.addStretch()
        trig_prop_btn_row.addWidget(add_tp_btn)
        trig_prop_btn_row.addWidget(del_tp_btn)
        
        nl_trig_layout.addLayout(trig_prop_btn_row)
        nl_trig_layout.addWidget(self.trig_prop_table)
        
        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        nl_trig_layout.addWidget(line)
        
        trigger_layout.addRow(nl_trig_layout)
        
        self.scope_combo = QComboBox()
        self.scope_combo.addItem("Paragraph (e.g., Bullet points)", "paragraph")
        self.scope_combo.addItem("Entire Document (Global Search)", "all_text")
        trigger_layout.addRow("Scope:", self.scope_combo)
        
        self.match_type_combo = QComboBox()
        self.match_type_combo.addItem("Contains Text", "contains")
        self.match_type_combo.addItem("Starts With", "starts_with")
        self.match_type_combo.addItem("Wildcard Pattern (*)", "wildcard")
        self.match_type_combo.addItem("Regex (Advanced)", "regex")
        trigger_layout.addRow("Match Condition:", self.match_type_combo)
        
        pattern_row = QHBoxLayout()
        self.pattern_edit = QLineEdit()
        self.pattern_edit.setPlaceholderText("Text to find...")
        pattern_row.addWidget(self.pattern_edit)
        
        ai_pattern_btn = QPushButton("AI")
        ai_pattern_btn.setToolTip("Generate pattern from examples")
        ai_pattern_btn.setFixedWidth(40)
        ai_pattern_btn.clicked.connect(self.ai_generate_pattern)
        pattern_row.addWidget(ai_pattern_btn)
        trigger_layout.addRow("Pattern:", pattern_row)
        
        # New Options Layout
        opts_layout = QHBoxLayout()
        self.chk_whole_word = QCheckBox("Whole Word Only")
        self.chk_case_sensitive = QCheckBox("Case Sensitive")
        self.chk_is_list = QCheckBox("Is List Item?")
        self.chk_is_list.setToolTip("Matches only if the paragraph is an automatic list item (bullet/number).")
        
        opts_layout.addWidget(self.chk_whole_word)
        opts_layout.addWidget(self.chk_case_sensitive)
        opts_layout.addWidget(self.chk_is_list)
        trigger_layout.addRow("", opts_layout)
        
        layout.addWidget(trigger_group)

        # --- Action Section ---
        action_group = QGroupBox("Action (What to do?)")
        action_layout = QVBoxLayout(action_group)
        
        form_act = QFormLayout()
        self.action_type_combo = QComboBox()
        self.action_type_combo.addItems(["replace", "cycle", "format", "format_advanced"])
        self.action_type_combo.setItemText(3, "Advanced / AI (Natural Language)")
        self.action_type_combo.currentIndexChanged.connect(self.on_action_changed)
        form_act.addRow("Type:", self.action_type_combo)
        action_layout.addLayout(form_act)
        
        self.action_stack = QStackedWidget()
        
        # Page 0: Replace
        page_replace = QWidget()
        pr_layout = QFormLayout(page_replace)
        self.replace_edit = QLineEdit()
        pr_layout.addRow("Replace With:", self.replace_edit)
        self.action_stack.addWidget(page_replace)
        
        # Page 1: Cycle
        page_cycle = QWidget()
        pc_layout = QVBoxLayout(page_cycle)
        pc_layout.addWidget(QLabel("Variations (One per line):"))
        
        self.variations_list = QListWidget()
        pc_layout.addWidget(self.variations_list)
        
        btn_row = QHBoxLayout()
        add_var_btn = QPushButton("Add")
        add_var_btn.clicked.connect(self.add_variation)
        rem_var_btn = QPushButton("Remove")
        rem_var_btn.clicked.connect(self.remove_variation)
        btn_row.addWidget(add_var_btn)
        btn_row.addWidget(rem_var_btn)
        pc_layout.addLayout(btn_row)
        
        self.action_stack.addWidget(page_cycle)

        # Page 2: Format
        page_format = QWidget()
        pf_layout = QVBoxLayout(page_format)
        
        # Indent & Spacing Scroll Area or just GroupBoxes
        ind_grp = QGroupBox("Indentation & Spacing")
        ind_form = QFormLayout(ind_grp)
        self.left_ind_edit = QLineEdit()
        self.left_ind_edit.setPlaceholderText("e.g. 0.5 or 0.5 in")
        ind_form.addRow("Left Indent:", self.left_ind_edit)
        
        self.first_line_ind_edit = QLineEdit()
        self.first_line_ind_edit.setPlaceholderText("e.g. 0.5 or 0.5 in")
        ind_form.addRow("First Line Indent:", self.first_line_ind_edit)
        
        self.space_after_edit = QLineEdit()
        self.space_after_edit.setPlaceholderText("e.g. 12 or 1 line")
        ind_form.addRow("Space After:", self.space_after_edit)
        
        self.alignment_combo = QComboBox()
        self.alignment_combo.addItems(["", "Left", "Center", "Right", "Justify"])
        ind_form.addRow("Alignment:", self.alignment_combo)
        pf_layout.addWidget(ind_grp)
        
        font_grp = QGroupBox("Font & Style")
        font_form = QFormLayout(font_grp)
        self.font_name_edit = QLineEdit()
        self.font_name_edit.setPlaceholderText("e.g. Times New Roman")
        font_form.addRow("Font Name:", self.font_name_edit)
        
        self.font_size_edit = QLineEdit()
        self.font_size_edit.setPlaceholderText("e.g. 12")
        font_form.addRow("Font Size:", self.font_size_edit)
        
        font_opts = QHBoxLayout()
        self.chk_bold = QCheckBox("Bold")
        self.chk_italic = QCheckBox("Italic")
        font_opts.addWidget(self.chk_bold)
        font_opts.addWidget(self.chk_italic)
        font_form.addRow("Effects:", font_opts)
        
        self.style_edit = QLineEdit()
        self.style_edit.setPlaceholderText("e.g. Heading 1")
        font_form.addRow("Word Style:", self.style_edit)
        
        pf_layout.addWidget(font_grp)
        pf_layout.addStretch()
        self.action_stack.addWidget(page_format)
        
        # Page 3: Format Advanced (Natural Language / Dynamic)
        page_adv = QWidget()
        pa_layout = QVBoxLayout(page_adv)
        
        pa_layout.addWidget(QLabel("Describe formatting (e.g., 'Small caps and keep with next'):"))
        
        adv_input_row = QHBoxLayout()
        self.adv_input = QLineEdit()
        self.adv_input.setPlaceholderText("Type instruction here...")
        self.adv_gen_btn = QPushButton("Generate Properties")
        self.adv_gen_btn.clicked.connect(self.generate_dynamic_properties)
        adv_input_row.addWidget(self.adv_input)
        adv_input_row.addWidget(self.adv_gen_btn)
        pa_layout.addLayout(adv_input_row)
        
        pa_layout.addWidget(QLabel("Generated Properties (Word Object Model):"))
        self.prop_table = QTableWidget()
        self.prop_table.setColumnCount(2)
        self.prop_table.setHorizontalHeaderLabels(["Property Path (e.g. Range.Font.Bold)", "Value"])
        self.prop_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        pa_layout.addWidget(self.prop_table)
        
        btn_row_adv = QHBoxLayout()
        add_prop_btn = QPushButton("+ Row")
        add_prop_btn.clicked.connect(lambda: self.prop_table.insertRow(self.prop_table.rowCount()))
        del_prop_btn = QPushButton("- Row")
        del_prop_btn.clicked.connect(lambda: self.prop_table.removeRow(self.prop_table.currentRow()))
        btn_row_adv.addWidget(add_prop_btn)
        btn_row_adv.addWidget(del_prop_btn)
        btn_row_adv.addStretch()
        pa_layout.addLayout(btn_row_adv)
        
        self.action_stack.addWidget(page_adv)
        
        action_layout.addWidget(self.action_stack)
        layout.addWidget(action_group)

        # --- Buttons ---
        btn_box = QHBoxLayout()
        save_btn = QPushButton("Save Rule")
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        btn_box.addStretch()
        btn_box.addWidget(cancel_btn)
        btn_box.addWidget(save_btn)
        layout.addLayout(btn_box)

    def generate_trigger_config(self):
        desc = self.nl_trig_input.text().strip()
        if not desc:
             QMessageBox.warning(self, "Input Required", "Please enter a description first.")
             return
             
        model = self.ai_model_combo.currentText()
        script_path = os.path.join(SCRIPTS_DIR, "rule_builder.py")
        
        self.nl_trig_btn.setText("Generating...")
        self.nl_trig_btn.setEnabled(False)
        from PySide6.QtWidgets import QApplication
        QApplication.processEvents()
        
        import subprocess
        try:
            res = subprocess.check_output([sys.executable, script_path, "--describe", desc, "--model", model])
            data = json.loads(res.decode())
            
            if "error" in data:
                QMessageBox.critical(self, "AI Error", f"AI failed: {data['error']}")
            else:
                trig = data.get('trigger', {})
                
                # Update Text Fields
                self.pattern_edit.setText(trig.get('pattern', ''))
                
                scope = trig.get('scope', 'paragraph')
                idx = self.scope_combo.findData(scope)
                if idx >= 0: self.scope_combo.setCurrentIndex(idx)
                
                mtype = trig.get('match_type', 'contains')
                idx = self.match_type_combo.findData(mtype)
                if idx >= 0: self.match_type_combo.setCurrentIndex(idx)
                
                self.chk_whole_word.setChecked(trig.get('whole_word', False))
                self.chk_case_sensitive.setChecked(trig.get('case_sensitive', False))
                self.chk_is_list.setChecked(trig.get('is_list', False))
                
                # Update Property Table
                props = trig.get('property_match', {})
                self.trig_prop_table.setRowCount(0)
                for key, val in props.items():
                    row = self.trig_prop_table.rowCount()
                    self.trig_prop_table.insertRow(row)
                    self.trig_prop_table.setItem(row, 0, QTableWidgetItem(str(key)))
                    self.trig_prop_table.setItem(row, 1, QTableWidgetItem(str(val)))
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to call AI Builder: {e}")
        finally:
            self.nl_trig_btn.setText("✨ Auto-Fill Trigger")
            self.nl_trig_btn.setEnabled(True)

    def generate_dynamic_properties(self):
        desc = self.adv_input.text().strip()
        if not desc:
             QMessageBox.warning(self, "Input Required", "Please enter a description first.")
             return
             
        model = self.ai_model_combo.currentText()
        script_path = os.path.join(SCRIPTS_DIR, "rule_builder.py")
        
        # Show loading...
        self.adv_gen_btn.setText("Generating...")
        self.adv_gen_btn.setEnabled(False)
        from PySide6.QtWidgets import QApplication
        QApplication.processEvents()
        
        import subprocess
        try:
            res = subprocess.check_output([sys.executable, script_path, "--describe", desc, "--model", model])
            data = json.loads(res.decode())
            
            if "error" in data:
                QMessageBox.critical(self, "AI Error", f"AI failed: {data['error']}")
            else:
                # Extract dynamic properties
                formatting = data.get('action', {}).get('formatting', {})
                dyn_props = formatting.get('dynamic_properties', {})
                
                # Populate Table
                self.prop_table.setRowCount(0)
                for key, val in dyn_props.items():
                    row = self.prop_table.rowCount()
                    self.prop_table.insertRow(row)
                    self.prop_table.setItem(row, 0, QTableWidgetItem(str(key)))
                    self.prop_table.setItem(row, 1, QTableWidgetItem(str(val)))
                    
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to call AI Builder: {e}")
        finally:
            self.adv_gen_btn.setText("Generate Properties")
            self.adv_gen_btn.setEnabled(True)

    def apply_preset(self):
        menu = QMenu(self)
        
        # Helper to create action
        def add_preset_action(key, name):
            action = QAction(name, self)
            action.triggered.connect(lambda: self._load_preset_data(key))
            menu.addAction(action)

        add_preset_action("main_heading", "Main Heading (e.g. FACTUAL BACKGROUND)")
        add_preset_action("subheading_a", "Subheading A (e.g. A. )")
        add_preset_action("subheading_1", "Subheading 1 (e.g. 1. )")
        add_preset_action("bullet", "Bullet Points (Auto-List or Char)")
        add_preset_action("narrative", "Narrative Text")
        
        menu.exec(self.mapToGlobal(self.sender().pos()))

    def _load_preset_data(self, key):
        if key in PRESETS:
            # Deep copy to avoid modifying the global preset
            import copy
            self.rule = copy.deepcopy(PRESETS[key])
            self.load_data()
            QMessageBox.information(self, "Preset Loaded", f"Loaded preset: {self.rule['name']}")

    def ai_build_rule(self):
        from PySide6.QtWidgets import QInputDialog
        desc, ok = QInputDialog.getMultiLineText(self, "AI Rule Builder", 
                                                 "Describe the rule you want to create (e.g., 'Make paragraphs starting with Note bold and blue'):")
        if not ok or not desc.strip(): return
        
        script_path = os.path.join(SCRIPTS_DIR, "rule_builder.py")
        model = self.ai_model_combo.currentText()
        import subprocess
        try:
            res = subprocess.check_output([sys.executable, script_path, "--describe", desc, "--model", model])
            data = json.loads(res.decode())
            if "error" in data:
                QMessageBox.critical(self, "AI Error", f"AI failed: {data['error']}")
                return
            
            # Populate fields
            self.rule = data
            self.load_data()
            QMessageBox.information(self, "AI Success", "Rule data generated and loaded!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to call AI Builder: {e}")

    def get_from_word(self):
        # Access the ReportTab instance (parent of this dialog)
        if not self.parent() or not hasattr(self.parent(), 'web_view'):
            QMessageBox.warning(self, "Error", "Cannot access document preview.")
            return

        js_script = """
        (function() {
            var sel = window.getSelection();
            if (sel.rangeCount < 1) return JSON.stringify({error: "No selection found"});
            
            var node = sel.anchorNode;
            if (!node) return JSON.stringify({error: "No element selected"});
            if (node.nodeType === 3) node = node.parentElement; // Text node -> Element
            
            var computed = window.getComputedStyle(node);
            
            // Attempt to guess Word Style from class attribute
            // Word HTML filtered often puts style name in class or nothing
            var styleName = "";
            if (node.className) {
                // e.g. "MsoNormal", "Heading1"
                styleName = node.className.replace(/^Mso/, "");
            }

            return JSON.stringify({
                font_name: computed.fontFamily.split(",")[0].replace(/['"]/g, ""),
                font_size: computed.fontSize,
                font_weight: computed.fontWeight,
                font_style: computed.fontStyle,
                text_align: computed.textAlign,
                margin_left: computed.marginLeft,
                margin_bottom: computed.marginBottom,
                style: styleName
            });
        })();
        """
        
        self.parent().web_view.page().runJavaScript(js_script, self.handle_preview_selection)

    def handle_preview_selection(self, result):
        if not result:
            QMessageBox.warning(self, "Error", "Failed to retrieve selection from preview.")
            return
            
        try:
            data = json.loads(result)
        except:
            QMessageBox.warning(self, "Error", "Invalid data from preview.")
            return

        if "error" in data:
            QMessageBox.warning(self, "Selection Error", data["error"])
            return

        # Parse units
        def parse_px_to_pt(px_str):
            if not px_str: return 0.0
            val = float("".join(c for c in px_str if c.isdigit() or c == '.'))
            # 1px = 0.75pt
            return round(val * 0.75, 1)

        def parse_px_to_in(px_str):
            if not px_str: return 0.0
            val = float("".join(c for c in px_str if c.isdigit() or c == '.'))
            # 96px = 1in
            return round(val / 96.0, 2)

        # Update UI
        self.action_type_combo.setCurrentText("format")
        
        # Indent (px -> in)
        self.left_ind_edit.setText(str(parse_px_to_in(data.get('margin_left'))))
        
        # Space After (px -> pt)
        self.space_after_edit.setText(str(parse_px_to_pt(data.get('margin_bottom'))))
        
        # Alignment
        align_map = {"left": "Left", "center": "Center", "right": "Right", "justify": "Justify"}
        # textAlign might be "start", "end"
        align_val = data.get('text_align', 'left')
        if align_val == 'start': align_val = 'left'
        if align_val == 'end': align_val = 'right'
        
        self.alignment_combo.setCurrentText(align_map.get(align_val, 'Left'))
        
        # Font Name
        self.font_name_edit.setText(data.get('font_name', ''))
        
        # Font Size (px -> pt)
        self.font_size_edit.setText(str(parse_px_to_pt(data.get('font_size'))))
        
        # Bold/Italic
        fw = data.get('font_weight', 'normal')
        is_bold = (fw == 'bold') or (str(fw).isdigit() and int(fw) >= 700)
        self.chk_bold.setChecked(is_bold)
        
        self.chk_italic.setChecked(data.get('font_style') == 'italic')
        
        # Style
        self.style_edit.setText(data.get('style', ''))

        QMessageBox.information(self, "Success", "Formatting extracted from Preview selection!")

    def ai_generate_pattern(self):
        from PySide6.QtWidgets import QInputDialog
        examples, ok = QInputDialog.getMultiLineText(self, "AI Pattern Generator", 
                                                    "Paste examples of text you want this rule to match (one per line):")
        if not ok or not examples.strip(): return
        
        ex_list = [e.strip() for e in examples.split('\n') if e.strip()]
        script_path = os.path.join(SCRIPTS_DIR, "rule_builder.py")
        model = self.ai_model_combo.currentText()
        import subprocess
        try:
            cmd = [sys.executable, script_path, "--describe" if False else "--examples"] + ex_list + ["--model", model]
            res = subprocess.check_output(cmd)
            data = json.loads(res.decode())
            
            idx = self.match_type_combo.findData(data.get('match_type', 'contains'))
            if idx >= 0: self.match_type_combo.setCurrentIndex(idx)
            self.pattern_edit.setText(data.get('pattern', ''))
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to generate pattern: {e}")

    def _parse_unit(self, val, target_unit="pt"):
        if not val: return 0
        s = str(val).lower().strip()
        
        # Remove units and parse number
        num_str = "".join(c for c in s if c.isdigit() or c == '.')
        try:
            num = float(num_str)
        except:
            return 0
            
        # Detect source unit
        source_unit = "in" if ("in" in s or "inch" in s) else "line" if "line" in s else "pt"
        
        # Convert to points first
        points = num
        if source_unit == "in":
            points = num * 72.0
        elif source_unit == "line":
            points = num * 12.0
            
        # Convert to target unit
        if target_unit == "in":
            return round(points / 72.0, 3)
        return round(points, 2)

    def on_action_changed(self):
        idx = self.action_type_combo.currentIndex()
        self.action_stack.setCurrentIndex(idx)

    def add_variation(self):
        # Simple input dialog logic inline
        # For now, just add a placeholder or simple dialog
        # Let's use QInputDialog equivalent or a simple QLineEdit
        # We'll just add an editable item
        item = QListWidgetItem("New Variation")
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
        self.variations_list.addItem(item)
        self.variations_list.editItem(item)

    def remove_variation(self):
        row = self.variations_list.currentRow()
        if row >= 0:
            self.variations_list.takeItem(row)

    def load_data(self):
        if not self.rule:
            return
            
        self.name_edit.setText(self.rule.get('name', ''))
        
        trig = self.rule.get('trigger', {})
        
        # Set Scope (by data, defaulting to paragraph)
        scope_val = trig.get('scope', 'paragraph')
        idx = self.scope_combo.findData(scope_val)
        if idx >= 0: self.scope_combo.setCurrentIndex(idx)
        
        # Set Match Type
        match_val = trig.get('match_type', 'contains')
        idx = self.match_type_combo.findData(match_val)
        if idx >= 0: self.match_type_combo.setCurrentIndex(idx)

        self.pattern_edit.setText(trig.get('pattern', ''))
        self.chk_whole_word.setChecked(trig.get('whole_word', False))
        self.chk_case_sensitive.setChecked(trig.get('case_sensitive', False))
        self.chk_is_list.setChecked(trig.get('is_list', False))
        
        # Load Property Matches
        props = trig.get('property_match', {})
        self.trig_prop_table.setRowCount(0)
        for k, v in props.items():
            row = self.trig_prop_table.rowCount()
            self.trig_prop_table.insertRow(row)
            self.trig_prop_table.setItem(row, 0, QTableWidgetItem(str(k)))
            self.trig_prop_table.setItem(row, 1, QTableWidgetItem(str(v)))
        
        act = self.rule.get('action', {})
        atype = act.get('type', 'replace')
        
        # Handle Format Advanced detection
        fmt = act.get('formatting', {})
        if atype == 'format' and 'dynamic_properties' in fmt:
            atype = 'format_advanced'
            
        idx = self.action_type_combo.findText("Advanced / AI (Natural Language)") if atype == 'format_advanced' else self.action_type_combo.findText(atype)
        if idx >= 0: self.action_type_combo.setCurrentIndex(idx)
        else: self.action_type_combo.setCurrentText(atype)
        
        if atype == 'replace':
            self.replace_edit.setText(act.get('replacement', ''))
        elif atype == 'cycle':
            vars = act.get('variations', [])
            for v in vars:
                item = QListWidgetItem(v)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable)
                self.variations_list.addItem(item)
        elif atype == 'format':
            self.left_ind_edit.setText(str(fmt.get('left_indent', '')))
            self.first_line_ind_edit.setText(str(fmt.get('first_line_indent', '')))
            self.space_after_edit.setText(str(fmt.get('space_after', '')))
            self.alignment_combo.setCurrentText(fmt.get('alignment', '').capitalize())
            self.font_name_edit.setText(fmt.get('font_name', ''))
            self.font_size_edit.setText(str(fmt.get('font_size', '')))
            self.chk_bold.setChecked(fmt.get('font_bold', False))
            self.chk_italic.setChecked(fmt.get('font_italic', False))
            self.style_edit.setText(fmt.get('style', ''))
        elif atype == 'format_advanced':
            dyn = fmt.get('dynamic_properties', {})
            self.prop_table.setRowCount(0)
            for k, v in dyn.items():
                row = self.prop_table.rowCount()
                self.prop_table.insertRow(row)
                self.prop_table.setItem(row, 0, QTableWidgetItem(str(k)))
                self.prop_table.setItem(row, 1, QTableWidgetItem(str(v)))

    def get_rule_data(self):
        vars_list = []
        for i in range(self.variations_list.count()):
            vars_list.append(self.variations_list.item(i).text())

        # Formatting data
        formatting = {}
        if self.left_ind_edit.text():
            formatting['left_indent'] = self._parse_unit(self.left_ind_edit.text(), "in")
            
        if self.first_line_ind_edit.text():
            formatting['first_line_indent'] = self._parse_unit(self.first_line_ind_edit.text(), "in")
            
        if self.space_after_edit.text():
            formatting['space_after'] = self._parse_unit(self.space_after_edit.text(), "pt")
            
        if self.alignment_combo.currentText():
            formatting['alignment'] = self.alignment_combo.currentText().lower()
        if self.font_name_edit.text():
            formatting['font_name'] = self.font_name_edit.text()
        if self.font_size_edit.text():
            try: formatting['font_size'] = float(self.font_size_edit.text())
            except: pass
        if self.chk_bold.isChecked(): formatting['font_bold'] = True
        if self.chk_italic.isChecked(): formatting['font_italic'] = True
        if self.style_edit.text(): formatting['style'] = self.style_edit.text()

        action_type = self.action_type_combo.currentText()
        action_data = {"type": action_type}
        
        if action_type == "replace":
            action_data["replacement"] = self.replace_edit.text()
        elif action_type == "cycle":
            action_data["variations"] = vars_list
        elif action_type == "format":
            action_data["formatting"] = formatting
        elif action_type == "Advanced / AI (Natural Language)":
            # For advanced, we also default type to 'format' but include dynamic_properties
            action_data["type"] = "format"
            dyn_props = {}
            for i in range(self.prop_table.rowCount()):
                key = self.prop_table.item(i, 0).text() if self.prop_table.item(i, 0) else ""
                val_str = self.prop_table.item(i, 1).text() if self.prop_table.item(i, 1) else ""
                
                if key:
                    # Try to parse value (bool/int/float)
                    val = val_str
                    if val_str.lower() == 'true': val = True
                    elif val_str.lower() == 'false': val = False
                    else:
                        try:
                            if '.' in val_str: val = float(val_str)
                            else: val = int(val_str)
                        except:
                            pass # Keep as string
                    dyn_props[key] = val
            formatting['dynamic_properties'] = dyn_props
            action_data["formatting"] = formatting

        trigger_data = {
            "scope": self.scope_combo.currentData(),
            "match_type": self.match_type_combo.currentData(),
            "pattern": self.pattern_edit.text(),
            "whole_word": self.chk_whole_word.isChecked(),
            "case_sensitive": self.chk_case_sensitive.isChecked()
        }
        
        # Collect Property Matches
        prop_match = {}
        for i in range(self.trig_prop_table.rowCount()):
            key = self.trig_prop_table.item(i, 0).text() if self.trig_prop_table.item(i, 0) else ""
            val_str = self.trig_prop_table.item(i, 1).text() if self.trig_prop_table.item(i, 1) else ""
            if key:
                # Type conversion (same as dynamic props)
                val = val_str
                if val_str.lower() == 'true': val = True
                elif val_str.lower() == 'false': val = False
                else:
                    try:
                        if '.' in val_str: val = float(val_str)
                        else: val = int(val_str)
                    except: pass
                prop_match[key] = val
                
        if prop_match:
            trigger_data['property_match'] = prop_match
        
        if self.chk_is_list.isChecked():
            trigger_data["is_list"] = True

        return {
            "name": self.name_edit.text(),
            "enabled": True,
            "trigger": trigger_data,
            "action": action_data
        }

DEFAULT_PROMPTS = {
    "narrative": (
        "You are a professional legal report assistant. Rewrite the following text to be a cohesive, "
        "professional narrative. If there are bullet points, convert them into well-structured paragraphs. "
        "Maintain all factual details and use a formal, objective tone. Do not add information not present "
        "in the source unless it is necessary for flow. Output ONLY the rewritten text."
    ),
    "improve": (
        "You are a professional legal report assistant. Improve the following text for clarity, "
        "professionalism, and flow. Maintain the formal legal tone. Output ONLY the improved text."
    )
}

class AIToolDialog(QDialog):
    def __init__(self, tool=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("AI Tool Editor")
        self.resize(500, 400)
        self.tool = tool or {}
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        form = QFormLayout()
        
        self.name_edit = QLineEdit()
        form.addRow("Tool Name:", self.name_edit)
        
        self.mode_combo = QComboBox()
        self.mode_combo.addItems(["narrative", "improve"])
        self.mode_combo.currentTextChanged.connect(self.on_mode_changed)
        form.addRow("Mode:", self.mode_combo)
        
        self.model_combo = QComboBox()
        # Populating with some common models, ideally fetch dynamically like ChatTab
        self.model_combo.addItems([
            "models/gemini-3-flash-preview", 
            "models/gemini-3-pro-preview", 
            "models/gemini-2.5-flash",
            "models/gemini-1.5-pro",
            "models/gemini-1.5-flash"
        ])
        form.addRow("Model:", self.model_combo)
        
        layout.addLayout(form)
        
        layout.addWidget(QLabel("Custom Prompt (Optional - overrides default mode prompt):"))
        self.prompt_edit = QPlainTextEdit()
        layout.addWidget(self.prompt_edit)
        
        btn_box = QHBoxLayout()
        reset_btn = QPushButton("Reset to Default")
        reset_btn.clicked.connect(self.reset_to_default)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        
        btn_box.addWidget(reset_btn)
        btn_box.addStretch()
        btn_box.addWidget(cancel_btn)
        btn_box.addWidget(save_btn)
        layout.addLayout(btn_box)

    def on_mode_changed(self, mode):
        # If the prompt is currently a default or empty, update it
        current = self.prompt_edit.toPlainText().strip()
        if not current or current in DEFAULT_PROMPTS.values():
            self.prompt_edit.setPlainText(DEFAULT_PROMPTS.get(mode, ""))

    def reset_to_default(self):
        mode = self.mode_combo.currentText()
        self.prompt_edit.setPlainText(DEFAULT_PROMPTS.get(mode, ""))

    def load_data(self):
        if not self.tool:
            idx = self.model_combo.findText("models/gemini-3-flash-preview")
            if idx != -1: self.model_combo.setCurrentIndex(idx)
            return
        
        self.name_edit.setText(self.tool.get('name', ''))
        
        mode = self.tool.get('mode', 'improve')
        self.mode_combo.setCurrentText(mode)
        
        idx = self.model_combo.findText(self.tool.get('model', ''))
        if idx != -1: self.model_combo.setCurrentIndex(idx)
        
        prompt = self.tool.get('prompt', '')
        if not prompt:
            prompt = DEFAULT_PROMPTS.get(mode, "")
            
        self.prompt_edit.setPlainText(prompt)

    def get_tool_data(self):
        return {
            "name": self.name_edit.text(),
            "mode": self.mode_combo.currentText(),
            "model": self.model_combo.currentText(),
            "prompt": self.prompt_edit.toPlainText()
        }

class CustomWebPage(QWebEnginePage):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.report_tab = parent.report_tab if hasattr(parent, 'report_tab') else None

    def acceptNavigationRequest(self, url, _type, isMainFrame):
        if url.scheme() == "icharlotte":
            if self.report_tab:
                self.report_tab.on_custom_url(url)
            return False
        return super().acceptNavigationRequest(url, _type, isMainFrame)

class DroppableWebView(QWebEngineView):
    def __init__(self, parent=None, on_drop_callback=None):
        super().__init__(parent)
        self.report_tab = parent
        self.on_drop_callback = on_drop_callback
        self.setAcceptDrops(True)
        
        # Set custom page to intercept commands
        self.setPage(CustomWebPage(self))
        
        # Configure settings to allow editing and scripts
        settings = self.settings()
        settings.setAttribute(settings.WebAttribute.JavascriptEnabled, True)
        settings.setAttribute(settings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
        settings.setAttribute(settings.WebAttribute.LocalContentCanAccessFileUrls, True)
        settings.setAttribute(settings.WebAttribute.JavascriptCanAccessClipboard, True)
        settings.setAttribute(settings.WebAttribute.AllowRunningInsecureContent, True)

    def createWindow(self, type):
        # Prevent new windows from opening (like when clicking links)
        return self

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile():
                path = urls[0].toLocalFile()
                if path.lower().endswith(('.docx', '.doc')):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls and urls[0].isLocalFile():
            path = urls[0].toLocalFile()
            if self.on_drop_callback:
                self.on_drop_callback(path)
            event.acceptProposedAction()

    def contextMenuEvent(self, event):
        if not self.page().hasSelection():
            super().contextMenuEvent(event)
            return

        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { background-color: white; border: 1px solid #ccc; padding: 5px; }
            QMenu::item { padding: 5px 25px 5px 20px; border-radius: 3px; }
            QMenu::item:selected { background-color: #2196F3; color: white; }
        """)

        # 1. AI Tools Section
        title_action = QAction("✨ AI TOOLS", self)
        title_action.setEnabled(False)
        menu.addAction(title_action)
        menu.addSeparator()

        if self.report_tab:
            for i, tool in enumerate(self.report_tab.ai_tools):
                act = QAction(tool.get('name', 'AI Tool'), self)
                # Capture current index i
                act.triggered.connect(lambda checked, idx=i: self.report_tab.rewrite_selected(idx))
                menu.addAction(act)

        menu.addSeparator()
        
        # 2. Custom Prompt Input
        input_container = QWidget()
        input_layout = QVBoxLayout(input_container)
        input_layout.setContentsMargins(5, 5, 5, 5)
        
        prompt_input = QLineEdit()
        prompt_input.setPlaceholderText("Custom AI instruction...")
        prompt_input.setFixedWidth(200)
        
        def handle_custom_prompt():
            text = prompt_input.text().strip()
            if text:
                menu.close()
                self.report_tab.rewrite_selected(-1, custom_prompt=text)

        prompt_input.returnPressed.connect(handle_custom_prompt)
        input_layout.addWidget(QLabel("<b>Custom Prompt:</b>"))
        input_layout.addWidget(prompt_input)
        
        input_action = QWidgetAction(self)
        input_action.setDefaultWidget(input_container)
        menu.addAction(input_action)

        menu.exec(event.globalPos())


class DroppableListWidget(QListWidget):
    fileDropped = Signal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)
        # self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection) # Optional, keeping default

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        files = [u.toLocalFile() for u in urls if u.isLocalFile()]
        if files:
            self.fileDropped.emit(files)
            event.acceptProposedAction()
        else:
            event.ignore()

class ReportTab(QWidget):
    def __init__(self, main_window=None, parent=None):
        super().__init__(parent)
        self.main_window = main_window
        self.setAcceptDrops(True) # Enable Drag and Drop
        self.current_doc_path = None
        self.rules = []
        self.ai_tools = []
        self.agent_runner = None
        self.config_path = os.path.join(GEMINI_DATA_DIR, "report_rules.json")
        self.ai_tools_path = os.path.join(GEMINI_DATA_DIR, "report_ai_tools.json")
        self.state_path = os.path.join(GEMINI_DATA_DIR, "report_state.json")
        self.context_state_path = os.path.join(GEMINI_DATA_DIR, "report_context.json")

        self._post_save_callback = None
        self._silent_save = False

        # Use case-specific state if available
        if self.main_window and hasattr(self.main_window, 'file_number'):
            self.state_path = os.path.join(GEMINI_DATA_DIR, f"{self.main_window.file_number}_report_state.json")
            self.context_state_path = os.path.join(GEMINI_DATA_DIR, f"{self.main_window.file_number}_report_context.json")
        
        self.setup_ui()
        self.load_rules()
        self.load_ai_tools()
        self.load_state()
        self.load_context_state()

    def load_context_state(self):
        if os.path.exists(self.context_state_path):
            try:
                self.context_list.blockSignals(True)
                with open(self.context_state_path, 'r') as f:
                    data = json.load(f)
                    for entry in data:
                        path = entry.get('path')
                        checked = entry.get('checked', True)
                        if path and os.path.exists(path):
                            item = QListWidgetItem(os.path.basename(path))
                            item.setData(Qt.ItemDataRole.UserRole, path)
                            item.setToolTip(path)
                            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                            item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
                            self.context_list.addItem(item)
                self.context_list.blockSignals(False)
            except Exception as e:
                print(f"Error loading context state: {e}")
                self.context_list.blockSignals(False)

    def save_context_state(self):
        data = []
        for i in range(self.context_list.count()):
            item = self.context_list.item(i)
            data.append({
                'path': item.data(Qt.ItemDataRole.UserRole),
                'checked': item.checkState() == Qt.CheckState.Checked
            })
        try:
            with open(self.context_state_path, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            print(f"Error saving context state: {e}")

    def load_state(self):
        if os.path.exists(self.state_path):
            try:
                with open(self.state_path, 'r') as f:
                    state = json.load(f)
                    last_path = state.get('last_doc_path')
                    if last_path and os.path.exists(last_path):
                        self.load_document_path(last_path)
            except Exception as e:
                print(f"Error loading report state: {e}")

    def save_state(self):
        try:
            with open(self.state_path, 'w') as f:
                json.dump({'last_doc_path': self.current_doc_path}, f)
        except Exception as e:
            print(f"Error saving report state: {e}")

    def reset_state(self):
        """Clears the report tab state for a new case."""
        # Update paths for the new case
        if self.main_window and hasattr(self.main_window, 'file_number'):
            self.state_path = os.path.join(GEMINI_DATA_DIR, f"{self.main_window.file_number}_report_state.json")
            self.context_state_path = os.path.join(GEMINI_DATA_DIR, f"{self.main_window.file_number}_report_context.json")

        self.current_doc_path = None
        self.file_label.setText("No file loaded")
        self.web_view.setHtml("<div style='text-align:center; padding-top:50px; color:#666;'>Drag & Drop a Word document here...</div>")
        self.process_btn.setEnabled(False)
        self.save_doc_btn.setEnabled(False)
        self.context_list.clear()
        self.ai_output_list.clear()
        
        # Load state for the new case (if any)
        self.load_state()
        self.load_context_state()

    def rewrite_selected(self, tool_index, custom_prompt=None):
        if not self.current_doc_path:
            QMessageBox.warning(self, "No Document", "Please load a document first.")
            return
            
        # Trigger silent save, then run execution
        self.save_document(
            silent=True, 
            callback=lambda: self._execute_rewrite(tool_index, custom_prompt)
        )

    def _execute_rewrite(self, tool_index, custom_prompt):
        # Fix: Ensure tool is defined for custom prompts (-1)
        if tool_index >= 0 and tool_index < len(self.ai_tools):
            tool = self.ai_tools[tool_index]
        else:
            tool = {"name": "Custom", "mode": "improve", "model": "models/gemini-3-flash-preview"}
        
        # 1. Collect Context from Sidebar files (Only if CHECKED)
        sidebar_context = ""
        for i in range(self.context_list.count()):
            item = self.context_list.item(i)
            if item.checkState() != Qt.CheckState.Checked:
                continue
                
            path = item.data(Qt.ItemDataRole.UserRole)
            if path and os.path.exists(path):
                try:
                    ext = os.path.splitext(path)[1].lower()
                    if ext == '.pdf':
                        import fitz
                        doc = fitz.open(path)
                        text = ""
                        for page in doc: text += page.get_text()
                        sidebar_context += f"\\n--- REFERENCE FILE: {os.path.basename(path)} ---\\n{text}\\n"
                    elif ext in ['.docx', '.doc']:
                        import win32com.client as win32
                        word = win32.GetActiveObject("Word.Application")
                        d = word.Documents.Open(path, ReadOnly=True, Visible=False)
                        sidebar_context += f"\\n--- REFERENCE FILE: {os.path.basename(path)} ---\\n{d.Content.Text}\\n"
                        d.Close()
                    elif ext == '.txt':
                        with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                            sidebar_context += f"\\n--- REFERENCE FILE: {os.path.basename(path)} ---\\n{f.read()}\\n"
                except Exception as e:
                    print(f"Error reading context file {path}: {e}")

        # Use a small delay to ensure the context menu is closed and focus returns to body
        QTimer.singleShot(100, lambda: self._run_capture_js(tool, custom_prompt, sidebar_context))

    def _run_capture_js(self, tool, custom_prompt, sidebar_context):
        # Capture selection AND full text for context
        # We prepend sidebar context to the fullText
        js = f"""
        (function() {{
            try {{
                var sel = window.getSelection().toString();
                var full = document.body.innerText || document.documentElement.innerText || "";
                var sidebar = `{sidebar_context}`;
                return JSON.stringify({{
                    selection: sel,
                    fullText: sidebar + "\\n\\n--- MAIN DOCUMENT ---\\n" + full
                }});
            }} catch (e) {{
                return JSON.stringify({{error: e.toString()}});
            }}
        }})();
        """
        self.web_view.page().runJavaScript(js, lambda data: self.handle_ai_rewrite_selection(data, tool, custom_prompt))

    def setup_ui(self):
        # ... (setup_ui code is above)
        pass

    # Add drag and drop support to sidebar
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if os.path.isfile(path):
                    item = QListWidgetItem(os.path.basename(path))
                    item.setData(Qt.ItemDataRole.UserRole, path)
                    item.setToolTip(path)
                    self.context_list.addItem(item)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)

    def handle_ai_rewrite_selection(self, json_data, tool, custom_prompt=None):
        logger = LogManager()
        logger.add_log("Report Tab", f"Capture result: {str(json_data)[:200]}...")

        if not json_data:
            QMessageBox.critical(self, "Error", "Failed to capture text from preview (No data returned from JS).")
            return

        try:
            data = json.loads(json_data)
        except:
            QMessageBox.critical(self, "Error", f"Failed to parse text from preview: {json_data}")
            return

        if "error" in data:
            QMessageBox.critical(self, "Error", f"JS Error during capture: {data['error']}")
            return

        selection = data.get('selection', '')
        full_text = data.get('fullText', '')

        # Fallback: if browser failed to get full text, try reading the preview file from disk
        if not full_text:
            try:
                html_path = os.path.join(TEMP_DIR, "preview.mhtml")
                if os.path.exists(html_path):
                    with open(html_path, 'r', encoding='utf-8', errors='ignore') as f:
                        full_text = f.read()
                    logger.add_log("Report Tab", "Using disk-based preview for context fallback.")
            except: pass

        if not selection or len(selection.strip()) < 5:
            QMessageBox.warning(self, "Invalid Selection", "Please select a longer portion of text.")
            return

        # 1. Save selection and context to temp files
        temp_text_path = os.path.join(TEMP_DIR, "ai_rewrite_selection.txt")
        temp_context_path = os.path.join(TEMP_DIR, "ai_rewrite_context.txt")
        
        try:
            with open(temp_text_path, 'w', encoding='utf-8') as f:
                f.write(selection)
            with open(temp_context_path, 'w', encoding='utf-8') as f:
                f.write(full_text)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save temporary text: {e}")
            return

        # 2. Run ai_rewrite.py
        script_path = os.path.join(SCRIPTS_DIR, "ai_rewrite.py")
        
        mode = tool.get('mode', 'improve')
        model = tool.get('model', 'models/gemini-3-flash-preview')
        
        cmd_args = [
            script_path, 
            self.current_doc_path, 
            temp_text_path, 
            "--mode", mode, 
            "--model", model,
            "--context_file", temp_context_path
        ]
        
        # Priority: custom_prompt (from menu input) -> tool.get('prompt') (from preset tool)
        final_prompt = custom_prompt if custom_prompt else tool.get('prompt', '')
        if final_prompt:
            temp_prompt_path = os.path.join(TEMP_DIR, "ai_custom_prompt.txt")
            with open(temp_prompt_path, 'w', encoding='utf-8') as f:
                f.write(final_prompt)
            cmd_args.extend(["--prompt_file", temp_prompt_path])

        if self.main_window:
            # self.main_window.tabs.setCurrentIndex(1) # Optional: Switch to status
            self.rewrite_runner = self.main_window.add_status_task(
                f"AI: {tool.get('name', 'Rewrite')}", 
                f"Processing with document context...",
                sys.executable,
                cmd_args
            )
            
            logger = LogManager()
            logger.add_log("Report Tab", f"Starting AI Rewrite with Global Context for {self.current_doc_path}")
            self.rewrite_runner.log_update.connect(lambda msg: logger.add_log("Report Tab", msg.strip()))
            self.rewrite_runner.finished.connect(self.on_rewrite_finished)
        else:
            self.rewrite_runner = AgentRunner(sys.executable, cmd_args, None)
            self.rewrite_runner.finished.connect(self.on_rewrite_finished)
            self.rewrite_runner.start()


    def on_rewrite_finished(self, success):
        logger = LogManager()
        if success:
            logger.add_log("Report Tab", "AI Rewrite completed successfully.")
            self.update_preview()
        else:
            logger.add_log("Report Tab", "AI Rewrite failed.")
            QMessageBox.critical(self, "Error", "AI Rewrite failed. Check logs.")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if urls and urls[0].isLocalFile():
                path = urls[0].toLocalFile()
                if path.lower().endswith(('.docx', '.doc')):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls and urls[0].isLocalFile():
            path = urls[0].toLocalFile()
            self.load_document_path(path)
            event.acceptProposedAction()

    def add_context_files(self, files):
        added_any = False
        for path in files:
            if os.path.isfile(path):
                # Check duplicates
                exists = False
                for i in range(self.context_list.count()):
                    item = self.context_list.item(i)
                    if item.data(Qt.ItemDataRole.UserRole) == path:
                        exists = True
                        break
                
                if not exists:
                    item = QListWidgetItem(os.path.basename(path))
                    item.setData(Qt.ItemDataRole.UserRole, path)
                    item.setToolTip(path)
                    item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    item.setCheckState(Qt.CheckState.Checked)
                    self.context_list.addItem(item)
                    added_any = True
        
        if added_any:
            self.save_context_state()

    def load_document_path(self, path):
        if os.path.exists(path):
            self.current_doc_path = path
            self.file_label.setText(os.path.basename(path))
            self.process_btn.setEnabled(True)
            self.update_preview()
            self.save_state()
            self.refresh_ai_outputs()

    def refresh_ai_outputs(self):
        self.ai_output_list.clear()
        
        # Determine case path from main window
        case_path = None
        if self.main_window and hasattr(self.main_window, 'case_path'):
            case_path = self.main_window.case_path
        elif self.current_doc_path:
            # Fallback: try to deduce from current doc
            # This is less reliable than main_window.case_path
            path_parts = self.current_doc_path.split(os.sep)
            # Most cases follow a structure where we can find the root
            # But let's stick to main_window if possible.
            pass
            
        if not case_path:
            return

        ai_output_dir = os.path.join(case_path, "NOTES", "AI OUTPUT")
        if os.path.exists(ai_output_dir):
            try:
                files = os.listdir(ai_output_dir)
                # Sort by modification time, newest first
                files.sort(key=lambda x: os.path.getmtime(os.path.join(ai_output_dir, x)), reverse=True)
                
                for f in files:
                    full_path = os.path.join(ai_output_dir, f)
                    if os.path.isfile(full_path):
                        item = QListWidgetItem(f)
                        item.setData(Qt.ItemDataRole.UserRole, full_path)
                        item.setToolTip(full_path)
                        self.ai_output_list.addItem(item)
            except Exception as e:
                print(f"Error refreshing AI outputs: {e}")

    def open_ai_output(self, item):
        path = item.data(Qt.ItemDataRole.UserRole)
        if path and os.path.exists(path):
            try:
                os.startfile(path)
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Could not open file: {e}")

    def setup_ui(self):
        self.setAcceptDrops(True)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        splitter = QSplitter(Qt.Orientation.Horizontal)
        
        # --- Left Panel ---
        left_panel = QFrame()
        left_layout = QVBoxLayout(left_panel)
        
        # Rule Manager Section
        left_layout.addWidget(QLabel("<b>Rule Manager</b>"))
        self.rule_list_widget = QListWidget()
        self.rule_list_widget.setAlternatingRowColors(True)
        self.rule_list_widget.itemDoubleClicked.connect(self.edit_rule)
        self.rule_list_widget.itemChanged.connect(self.on_rule_item_changed)
        left_layout.addWidget(self.rule_list_widget, 2)
        
        btn_grid = QHBoxLayout()
        add_btn = QPushButton("Add")
        add_btn.clicked.connect(self.add_rule)
        edit_btn = QPushButton("Edit")
        edit_btn.clicked.connect(self.edit_rule)
        del_btn = QPushButton("Delete")
        del_btn.clicked.connect(self.delete_rule)
        btn_grid.addWidget(add_btn)
        btn_grid.addWidget(edit_btn)
        btn_grid.addWidget(del_btn)
        left_layout.addLayout(btn_grid)
        
        move_grid = QHBoxLayout()
        up_btn = QPushButton("Up")
        up_btn.clicked.connect(lambda: self.move_rule(-1))
        down_btn = QPushButton("Down")
        down_btn.clicked.connect(lambda: self.move_rule(1))
        move_grid.addWidget(up_btn)
        move_grid.addWidget(down_btn)
        left_layout.addLayout(move_grid)
        
        left_layout.addSpacing(10)
        
        # AI Tool Editor Section
        left_layout.addWidget(QLabel("<b>AI Tool Editor</b>"))
        self.ai_tool_list = QListWidget()
        self.ai_tool_list.setAlternatingRowColors(True)
        self.ai_tool_list.itemDoubleClicked.connect(self.edit_ai_tool)
        left_layout.addWidget(self.ai_tool_list, 1)
        
        ai_btn_grid = QHBoxLayout()
        add_ai_btn = QPushButton("Add Tool")
        add_ai_btn.clicked.connect(self.add_ai_tool)
        edit_ai_btn = QPushButton("Edit")
        edit_ai_btn.clicked.connect(self.edit_ai_tool)
        del_ai_btn = QPushButton("Delete")
        del_ai_btn.clicked.connect(self.delete_ai_tool)
        ai_btn_grid.addWidget(add_ai_btn)
        ai_btn_grid.addWidget(edit_ai_btn)
        ai_btn_grid.addWidget(del_ai_btn)
        left_layout.addLayout(ai_btn_grid)

        left_layout.addSpacing(10)
        
        # AI Output Section
        left_layout.addWidget(QLabel("<b>AI Output Documents</b>"))
        self.ai_output_list = QListWidget()
        self.ai_output_list.setAlternatingRowColors(True)
        self.ai_output_list.itemDoubleClicked.connect(self.open_ai_output)
        left_layout.addWidget(self.ai_output_list, 1)
        
        refresh_output_btn = QPushButton("🔄 Refresh AI Outputs")
        refresh_output_btn.clicked.connect(self.refresh_ai_outputs)
        left_layout.addWidget(refresh_output_btn)

        left_layout.addSpacing(20)
        left_layout.addWidget(QLabel("<b>Document Actions</b>"))
        
        # Doc Controls
        self.file_label = QLabel("No file loaded")
        self.file_label.setWordWrap(True)
        left_layout.addWidget(self.file_label)
        
        load_btn = QPushButton("Load Document")
        load_btn.clicked.connect(self.load_document)
        left_layout.addWidget(load_btn)
        
        self.process_btn = QPushButton("Run Rules & Process")
        self.process_btn.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold; padding: 10px;")
        self.process_btn.clicked.connect(self.process_document)
        self.process_btn.setEnabled(False)
        left_layout.addWidget(self.process_btn)
        
        left_layout.addStretch()
        
        splitter.addWidget(left_panel)
        
        # --- Right Panel (Preview) ---
        right_panel = QFrame()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create a Container for the Preview Area (Toolbar + WebView)
        preview_area = QWidget()
        preview_area_layout = QVBoxLayout(preview_area)
        preview_area_layout.setContentsMargins(0, 5, 0, 0)

        # Header Row (Centered to match the 8.5in document)
        header_container = QHBoxLayout()
        header_container.setContentsMargins(10, 0, 10, 0)
        
        # Inner layout to group Label and Toolbar and center them like the page
        center_row = QHBoxLayout()
        center_row.addStretch()
        
        # The actual content group (Matches document width roughly)
        content_group = QWidget()
        content_group.setFixedWidth(816) # 8.5 inches at 96 DPI
        cg_layout = QHBoxLayout(content_group)
        cg_layout.setContentsMargins(0, 0, 0, 0)
        
        preview_label = QLabel("<b>Document Preview</b>")
        preview_label.setStyleSheet("font-size: 13px; color: #333;")
        cg_layout.addWidget(preview_label)
        cg_layout.addSpacing(15)

        # Formatting Toolbar
        self.format_toolbar = QFrame()
        self.format_toolbar.setObjectName("formatToolbar")
        self.format_toolbar.setStyleSheet("""
            #formatToolbar { background: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; padding: 4px; }
            QPushButton { border: 1px solid transparent; border-radius: 3px; background: transparent; padding: 4px; font-weight: normal; }
            QPushButton:hover { background: #e9ecef; border: 1px solid #adb5bd; }
            QComboBox { background: white; border: 1px solid #ccc; border-radius: 3px; padding: 2px; }
        """)
        ft_layout = QVBoxLayout(self.format_toolbar)
        ft_layout.setContentsMargins(4, 4, 4, 4)
        ft_layout.setSpacing(6)

        # Row 1: Styles, Font, Size, and History
        row1 = QHBoxLayout()
        row1.setSpacing(4)
        self.style_gallery = QComboBox()
        self.style_gallery.addItems(["Normal", "Heading 1", "Heading 2", "Heading 3", "Title", "Subtitle"])
        self.style_gallery.setFixedWidth(110)
        self.style_gallery.currentTextChanged.connect(lambda s: self.exec_format_cmd("formatBlock", s))
        row1.addWidget(self.style_gallery)
        row1.addSpacing(5)
        self.font_family_combo = QComboBox()
        self.font_family_combo.addItems(["Times New Roman", "Arial", "Calibri", "Courier New", "Georgia"])
        self.font_family_combo.setCurrentText("Times New Roman")
        self.font_family_combo.setFixedWidth(130)
        self.font_family_combo.currentTextChanged.connect(lambda f: self.exec_format_cmd("fontName", f))
        row1.addWidget(self.font_family_combo)
        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems(["8", "9", "10", "11", "12", "14", "16", "18", "20", "24", "26", "28", "36", "48", "72"])
        self.font_size_combo.setCurrentText("12")
        self.font_size_combo.setFixedWidth(55)
        self.font_size_combo.currentTextChanged.connect(lambda sz: self.exec_format_cmd("fontSize", sz))
        row1.addWidget(self.font_size_combo)
        row1.addSpacing(10)
        
        def create_format_btn(text, tooltip, cmd, value=None):
            btn = QPushButton(text)
            btn.setToolTip(tooltip)
            btn.setFixedWidth(28); btn.setFixedHeight(28)
            btn.setStyleSheet("font-weight: bold; font-family: 'Segoe UI', sans-serif; font-size: 13px;")
            btn.clicked.connect(lambda: self.exec_format_cmd(cmd, value))
            return btn

        row1.addWidget(create_format_btn("↶", "Undo", "undo"))
        row1.addWidget(create_format_btn("↷", "Redo", "redo"))
        row1.addStretch()
        ft_layout.addLayout(row1)

        # Row 2: Basic & Extended Formatting
        row2 = QHBoxLayout()
        row2.setSpacing(2)
        row2.addWidget(create_format_btn("B", "Bold", "bold"))
        row2.addWidget(create_format_btn("I", "Italic", "italic"))
        row2.addWidget(create_format_btn("U", "Underline", "underline"))
        row2.addWidget(create_format_btn("S", "Strikethrough", "strikeThrough"))
        sub_btn = create_format_btn("x₂", "Subscript", "subscript"); sub_btn.setStyleSheet("font-size: 10px;")
        row2.addWidget(sub_btn)
        sup_btn = create_format_btn("x²", "Superscript", "superscript"); sup_btn.setStyleSheet("font-size: 10px;")
        row2.addWidget(sup_btn)
        row2.addSpacing(8)
        highlight_btn = create_format_btn("H", "Highlight", "hiliteColor")
        highlight_btn.setStyleSheet("background-color: yellow; border: 1px solid #ccc;")
        row2.addWidget(highlight_btn)
        row2.addSpacing(10)
        row2.addWidget(create_format_btn("⫷", "Left", "justifyLeft"))
        row2.addWidget(create_format_btn("⫸", "Center", "justifyCenter"))
        row2.addWidget(create_format_btn("⫶", "Right", "justifyRight"))
        row2.addSpacing(10)
        row2.addWidget(create_format_btn("•", "Bullets", "insertUnorderedList"))
        row2.addWidget(create_format_btn("1.", "Numbered List", "insertOrderedList"))
        
        row2.addSpacing(10)
        
        # Spacing & Table Menus
        self.spacing_menu_btn = QPushButton("↕ Spacing")
        self.spacing_menu_btn.setToolTip("Line & Paragraph Spacing")
        self.spacing_menu_btn.setStyleSheet("padding: 2px 8px; border: 1px solid #ccc;")
        self.spacing_menu_btn.clicked.connect(self.show_spacing_menu)
        row2.addWidget(self.spacing_menu_btn)
        
        self.table_menu_btn = QPushButton("⊞ Table")
        self.table_menu_btn.setToolTip("Insert or Modify Tables")
        self.table_menu_btn.setStyleSheet("padding: 2px 8px; border: 1px solid #ccc;")
        self.table_menu_btn.clicked.connect(self.show_table_menu)
        row2.addWidget(self.table_menu_btn)
        
        row2.addStretch()
        ft_layout.addLayout(row2)
        
        cg_layout.addWidget(self.format_toolbar)
        cg_layout.addStretch()
        
        center_row.addWidget(content_group)
        center_row.addStretch()
        
        self.save_doc_btn = QPushButton("💾 Save Changes")
        self.save_doc_btn.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold; padding: 5px 12px;")
        self.save_doc_btn.clicked.connect(self.save_document)
        self.save_doc_btn.setEnabled(False)
        
        header_container.addLayout(center_row)
        header_container.addWidget(self.save_doc_btn)
        preview_area_layout.addLayout(header_container)

        # --- Find Bar ---
        self.find_bar = QFrame()
        self.find_bar.setStyleSheet("background: #f1f3f4; border-bottom: 1px solid #ccc;")
        fb_layout = QHBoxLayout(self.find_bar)
        fb_layout.setContentsMargins(10, 5, 10, 5)
        
        self.find_input = QLineEdit()
        self.find_input.setPlaceholderText("Find...")
        self.find_input.textChanged.connect(self.run_find)
        
        self.replace_input = QLineEdit()
        self.replace_input.setPlaceholderText("Replace with...")
        
        replace_btn = QPushButton("Replace")
        replace_btn.clicked.connect(self.run_replace)
        
        replace_all_btn = QPushButton("Replace All")
        replace_all_btn.clicked.connect(lambda: self.run_replace(all=True))
        
        close_find_btn = QPushButton("×")
        close_find_btn.setFixedWidth(20)
        close_find_btn.clicked.connect(lambda: self.find_bar.hide())
        
        fb_layout.addWidget(QLabel("🔍"))
        fb_layout.addWidget(self.find_input)
        fb_layout.addWidget(QLabel("➔"))
        fb_layout.addWidget(self.replace_input)
        fb_layout.addWidget(replace_btn)
        fb_layout.addWidget(replace_all_btn)
        fb_layout.addWidget(close_find_btn)
        self.find_bar.hide()
        preview_area_layout.addWidget(self.find_bar)
        
        self.web_view = DroppableWebView(self, on_drop_callback=self.load_document_path)
        self.web_view.setHtml("<div style='text-align:center; padding-top:50px; color:#666;'>Drag & Drop a Word document here...</div>")
        preview_area_layout.addWidget(self.web_view, 1)

        # --- Preview & Sidebar Splitter ---
        preview_splitter = QSplitter(Qt.Orientation.Horizontal)
        preview_splitter.addWidget(preview_area)
        
        # Sidebar
        self.sidebar = QFrame()
        self.sidebar.setFixedWidth(250)
        self.sidebar.setStyleSheet("background: #f8f9fa; border-left: 1px solid #ccc;")
        sidebar_layout = QVBoxLayout(self.sidebar)
        
        sidebar_layout.addWidget(QLabel("<b>Reference Context</b>"))
        sidebar_layout.addWidget(QLabel("<small>Drop files here to use as AI context</small>"))
        
        self.context_list = DroppableListWidget(self)
        self.context_list.setAlternatingRowColors(True)
        self.context_list.fileDropped.connect(self.add_context_files)
        self.context_list.itemChanged.connect(lambda item: self.save_context_state())
        sidebar_layout.addWidget(self.context_list)
        
        self.clear_context_btn = QPushButton("Clear Context")
        self.clear_context_btn.clicked.connect(lambda: self.context_list.clear())
        sidebar_layout.addWidget(self.clear_context_btn)
        
        preview_splitter.addWidget(self.sidebar)
        right_layout.addWidget(preview_splitter, 1)

        splitter.addWidget(right_panel)
        splitter.setSizes([350, 850]) # Moved after adding both panels
        
        layout.addWidget(splitter)
        
    def run_find(self):
        query = self.find_input.text()
        # Use browser's built-in find functionality (case-insensitive by default)
        # Note: True = forward, False = caseSensitive, False = backwards
        self.web_view.findText(query)

    def run_replace(self, all=False):
        find_text = self.find_input.text()
        replace_text = self.replace_input.text()
        if not find_text: return
        
        # JS logic for replace / replace all
        # We use a custom search/replace to avoid breaking the DOM too much
        js = f"""
        (function() {{
            var find = "{find_text}";
            var repl = "{replace_text}";
            var all = {str(all).lower()};
            
            // Simple approach: search for text nodes
            function walk(node) {{
                var child, next;
                switch (node.nodeType) {{
                    case 1:  // Element
                    case 9:  // Document
                    case 11: // Document fragment
                        child = node.firstChild;
                        while (child) {{
                            next = child.nextSibling;
                            walk(child);
                            child = next;
                        }}
                        break;
                    case 3: // Text node
                        var val = node.nodeValue;
                        if (val.toLowerCase().includes(find.toLowerCase())) {{
                            var regex = new RegExp(find, all ? "gi" : "i");
                            node.nodeValue = val.replace(regex, repl);
                            if (!all) return true; // Stop after first match
                        }}
                        break;
                }}
                return false;
            }}
            walk(document.body);
        }})();
        """
        self.web_view.page().runJavaScript(js)

    def show_spacing_menu(self):
        menu = QMenu(self)
        menu.setStyleSheet("QMenu { background: white; border: 1px solid #ccc; } QMenu::item:selected { background: #e3f2fd; color: black; }")
        
        spacing_options = [("1.0", "1.0"), ("1.15", "1.15"), ("1.5", "1.5"), ("2.0", "2.0"), ("2.5", "2.5"), ("3.0", "3.0")]
        for label, val in spacing_options:
            act = QAction(label, self)
            act.triggered.connect(lambda checked, v=val: self.exec_format_cmd("lineSpacing", v))
            menu.addAction(act)
        
        menu.addSeparator()
        
        add_before = QAction("Add Space Before Paragraph", self)
        add_before.triggered.connect(lambda: self.exec_format_cmd("paraSpacing", "before"))
        menu.addAction(add_before)
        
        add_after = QAction("Add Space After Paragraph", self)
        add_after.triggered.connect(lambda: self.exec_format_cmd("paraSpacing", "after"))
        menu.addAction(add_after)
        
        rem_all = QAction("Remove All Spacing", self)
        rem_all.triggered.connect(lambda: self.exec_format_cmd("paraSpacing", "remove"))
        menu.addAction(rem_all)
        
        menu.exec(self.spacing_menu_btn.mapToGlobal(self.spacing_menu_btn.rect().bottomLeft()))

    def show_table_menu(self):
        menu = QMenu(self)
        menu.setStyleSheet("QMenu { background: white; border: 1px solid #ccc; } QMenu::item:selected { background: #e3f2fd; color: black; }")
        
        # Grid insert logic simplified for menu
        insert_group = QMenu("Insert Table", self)
        for r in range(1, 6):
            for c in range(1, 6):
                act = QAction(f"{r}x{c}", self)
                act.triggered.connect(lambda checked, rows=r, cols=c: self.exec_format_cmd("insertTable", {"rows": rows, "cols": cols}))
                insert_group.addAction(act)
        menu.addMenu(insert_group)
        
        menu.addSeparator()
        
        menu.addAction("Insert Row Above", lambda: self.exec_format_cmd("tableRow", "above"))
        menu.addAction("Insert Row Below", lambda: self.exec_format_cmd("tableRow", "below"))
        menu.addSeparator()
        menu.addAction("Insert Column Left", lambda: self.exec_format_cmd("tableCol", "left"))
        menu.addAction("Insert Column Right", lambda: self.exec_format_cmd("tableCol", "right"))
        menu.addSeparator()
        menu.addAction("Delete Table", lambda: self.exec_format_cmd("tableDelete", None))
        
        menu.exec(self.table_menu_btn.mapToGlobal(self.table_menu_btn.rect().bottomLeft()))

    def exec_format_cmd(self, cmd, value=None):
        if not self.current_doc_path: return
        
        js = ""
        if cmd == "fontSize":
            # Map pts to 1-7 for execCommand fallback or use CSS in JS
            size_map = {"8": "1", "10": "2", "12": "3", "14": "4", "18": "5", "24": "6", "36": "7"}
            js = f"document.execCommand('fontSize', false, '{size_map.get(value, '3')}');"
        elif cmd == "fontName":
            js = f"document.execCommand('fontName', false, '{value}');"
        elif cmd == "formatBlock":
            block_map = {
                "Normal": "P",
                "Heading 1": "H1",
                "Heading 2": "H2",
                "Heading 3": "H3",
                "Title": "H1",
                "Subtitle": "H2"
            }
            tag = block_map.get(value, "P")
            js = f"document.execCommand('formatBlock', false, '<{tag}>');"
        elif cmd == "hiliteColor":
            js = "document.execCommand('hiliteColor', false, (window.getComputedStyle(window.getSelection().anchorNode.parentElement).backgroundColor === 'rgb(255, 255, 0)') ? 'transparent' : 'yellow');"
        elif cmd == "lineSpacing":
            js = f"""
            (function() {{
                var sel = window.getSelection();
                if (sel.rangeCount > 0) {{
                    var node = sel.anchorNode;
                    if (node.nodeType === 3) node = node.parentElement;
                    var para = node.closest('p, div, li, h1, h2, h3') || document.body;
                    para.style.lineHeight = '{value}';
                }}
            }})();
            """
        elif cmd == "paraSpacing":
            js = f"""
            (function() {{
                var sel = window.getSelection();
                if (sel.rangeCount > 0) {{
                    var node = sel.anchorNode;
                    if (node.nodeType === 3) node = node.parentElement;
                    var para = node.closest('p, div, li, h1, h2, h3') || document.body;
                    if ('{value}' === 'before') para.style.marginTop = '12pt';
                    else if ('{value}' === 'after') para.style.marginBottom = '12pt';
                    else {{ para.style.marginTop = '0'; para.style.marginBottom = '0'; }}
                }}
            }})();
            """
        elif cmd == "insertTable":
            rows = value['rows']
            cols = value['cols']
            table_html = "<table border='1' style='border-collapse: collapse; width: 100%;'>"
            for _ in range(rows):
                table_html += "<tr>"
                for _ in range(cols):
                    table_html += "<td style='padding: 5px; min-width: 50px; border: 1px solid #ccc;'>&nbsp;</td>"
                table_html += "</tr>"
            table_html += "</table><p>&nbsp;</p>"
            js = f"document.execCommand('insertHTML', false, `{table_html}`);"
        elif cmd.startswith("table"):
            # Simple table manipulation
            sub = cmd.replace("table", "").lower()
            js = f"""
            (function() {{
                var sel = window.getSelection();
                var td = sel.anchorNode.parentElement.closest('td');
                if (!td) return;
                var tr = td.parentElement;
                var table = tr.closest('table');
                if ('{sub}' === 'delete') {{ table.remove(); return; }}
                
                if ('{value}' === 'above' || '{value}' === 'below') {{
                    var newRow = table.insertRow('{value}' === 'above' ? tr.rowIndex : tr.rowIndex + 1);
                    for (var i = 0; i < tr.cells.length; i++) {{
                        var newCell = newRow.insertCell(i);
                        newCell.style.padding = '5px'; newCell.style.border = '1px solid #ccc'; newCell.innerHTML = '&nbsp;';
                    }}
                }} else if ('{value}' === 'left' || '{value}' === 'right') {{
                    var idx = td.cellIndex + ('{value}' === 'right' ? 1 : 0);
                    Array.from(table.rows).forEach(row => {{
                        var newCell = row.insertCell(idx);
                        newCell.style.padding = '5px'; newCell.style.border = '1px solid #ccc'; newCell.innerHTML = '&nbsp;';
                    }});
                }}
            }})();
            """
        else:
            js = f"document.execCommand('{cmd}', false, null);"
            
        self.web_view.page().runJavaScript(js)

    def save_document(self, callback=None, silent=False):
        if not self.current_doc_path: return
        
        # Safety: If called via signal (clicked), callback might be a boolean.
        if callback is not None and not callable(callback):
            callback = None

        self._post_save_callback = callback
        self._silent_save = silent

        # 1. Get edited HTML from WebView
        # We want the full document to preserve Word's structure as much as possible
        self.web_view.page().runJavaScript("document.documentElement.outerHTML", self.perform_save_to_word)

    def perform_save_to_word(self, html_content):
        if not html_content:
            QMessageBox.critical(self, "Error", "Failed to retrieve edited content from preview.")
            self._post_save_callback = None
            self._silent_save = False
            return

        # 2. Save edited MHTML to a temp file
        temp_html_path = os.path.join(TEMP_DIR, "edited_preview.mhtml")
        try:
            with open(temp_html_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save temporary preview: {e}")
            self._post_save_callback = None
            self._silent_save = False
            return

        # 3. Run rule_engine.py --save-preview
        script_path = os.path.join(SCRIPTS_DIR, "rule_engine.py")
        
        self.save_doc_btn.setEnabled(False)
        self.save_doc_btn.setText("Saving...")
        
        if self.main_window:
            self.save_runner = self.main_window.add_status_task(
                "Syncing Edits", 
                f"Saving changes to {os.path.basename(self.current_doc_path)}",
                sys.executable,
                [script_path, "--save-preview", temp_html_path, self.current_doc_path]
            )
            
            logger = LogManager()
            if not self._silent_save:
                logger.add_log("Report Tab", f"Saving edited preview back to {self.current_doc_path}")
            
            self.save_runner.log_update.connect(
                lambda msg: logger.add_log("Report Tab", f"[Save] {msg.strip()}")
            )
            self.save_runner.finished.connect(self.on_save_finished)
        else:
            # Fallback
            self.save_runner = AgentRunner(sys.executable, [script_path, "--save-html", temp_html_path, self.current_doc_path], None)
            self.save_runner.finished.connect(self.on_save_finished)
            self.save_runner.start()

    def on_save_finished(self, success):
        self.save_doc_btn.setEnabled(True)
        self.save_doc_btn.setText("💾 Save Changes")
        
        if success:
            if not self._silent_save:
                QMessageBox.information(self, "Success", "Changes saved to the Word document successfully.")
            
            # Execute callback if one exists
            if self._post_save_callback:
                cb = self._post_save_callback
                self._post_save_callback = None
                self._silent_save = False
                cb()
            else:
                self._silent_save = False
                
            # We don't necessarily need to refresh preview since the user just edited it,
            # but it ensures the HTML on disk matches the Word doc exactly.
            # self.update_preview() 
        else:
            QMessageBox.critical(self, "Error", "Failed to save changes to the Word document. Check logs.")
            self._post_save_callback = None
            self._silent_save = False

    def load_rules(self):
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r') as f:
                    self.rules = json.load(f)
            except:
                self.rules = []
        else:
            # Default Example Rules (Disabled by default)
            self.rules = [
                {
                    "name": "Fix 'the plaintiff'",
                    "enabled": False,
                    "trigger": {"scope": "paragraph", "match_type": "contains", "pattern": "the plaintiff"},
                    "action": {"type": "replace", "replacement": "Plaintiff"}
                },
                {
                    "name": "Variety for 'Plaintiff testified'",
                    "enabled": False,
                    "trigger": {"scope": "paragraph", "match_type": "starts_with", "pattern": "Plaintiff testified"},
                    "action": {
                        "type": "cycle", 
                        "variations": ["Plaintiff stated", "According to Plaintiff,", "During testimony, Plaintiff noted"]
                    }
                }
            ]
        self.refresh_rule_list()

    def save_rules(self):
        try:
            with open(self.config_path, 'w') as f:
                json.dump(self.rules, f, indent=2)
        except Exception as e:
            print(f"Error saving rules: {e}")

    def load_ai_tools(self):
        if os.path.exists(self.ai_tools_path):
            try:
                with open(self.ai_tools_path, 'r') as f:
                    self.ai_tools = json.load(f)
            except:
                self.ai_tools = []
        else:
            # Default Tools
            self.ai_tools = [
                {"name": "Convert to Narrative", "mode": "narrative", "prompt": "", "model": "models/gemini-3-flash-preview"},
                {"name": "Improve Text", "mode": "improve", "prompt": "", "model": "models/gemini-3-flash-preview"}
            ]
            self.save_ai_tools()
        self.refresh_ai_tool_list()

    def save_ai_tools(self):
        try:
            with open(self.ai_tools_path, 'w') as f:
                json.dump(self.ai_tools, f, indent=2)
        except Exception as e:
            print(f"Error saving AI tools: {e}")

    def refresh_rule_list(self):
        self.rule_list_widget.blockSignals(True)
        self.rule_list_widget.clear()
        for r in self.rules:
            item = QListWidgetItem(r.get('name', 'Unnamed Rule'))
            item.setCheckState(Qt.CheckState.Checked if r.get('enabled', True) else Qt.CheckState.Unchecked)
            self.rule_list_widget.addItem(item)
        self.rule_list_widget.blockSignals(False)

    def on_rule_item_changed(self, item):
        row = self.rule_list_widget.row(item)
        if row >= 0 and row < len(self.rules):
            self.rules[row]['enabled'] = (item.checkState() == Qt.CheckState.Checked)
            self.save_rules()

    def refresh_ai_tool_list(self):
        self.ai_tool_list.clear()
        for t in self.ai_tools:
            self.ai_tool_list.addItem(t.get('name', 'Unnamed Tool'))

    def add_rule(self):
        dialog = RuleDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.rules.append(dialog.get_rule_data())
            self.save_rules()
            self.refresh_rule_list()

    def edit_rule(self):
        row = self.rule_list_widget.currentRow()
        if row < 0: return
        
        dialog = RuleDialog(self.rules[row], parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.rules[row] = dialog.get_rule_data()
            self.save_rules()
            self.refresh_rule_list()

    def delete_rule(self):
        row = self.rule_list_widget.currentRow()
        if row >= 0:
            self.rules.pop(row)
            self.save_rules()
            self.refresh_rule_list()

    def move_rule(self, direction):
        row = self.rule_list_widget.currentRow()
        if row < 0: return
        
        new_row = row + direction
        if 0 <= new_row < len(self.rules):
            self.rules[row], self.rules[new_row] = self.rules[new_row], self.rules[row]
            self.save_rules()
            self.refresh_rule_list()
            self.rule_list_widget.setCurrentRow(new_row)

    def add_ai_tool(self):
        dialog = AIToolDialog(parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.ai_tools.append(dialog.get_tool_data())
            self.save_ai_tools()
            self.refresh_ai_tool_list()

    def edit_ai_tool(self):
        row = self.ai_tool_list.currentRow()
        if row < 0: return
        
        dialog = AIToolDialog(self.ai_tools[row], parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.ai_tools[row] = dialog.get_tool_data()
            self.save_ai_tools()
            self.refresh_ai_tool_list()

    def delete_ai_tool(self):
        row = self.ai_tool_list.currentRow()
        if row >= 0:
            self.ai_tools.pop(row)
            self.save_ai_tools()
            self.refresh_ai_tool_list()

    def load_document(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Document", "", "Word Documents (*.docx *.doc)")
        if path:
            self.load_document_path(path)

    def update_preview(self):
        if not self.current_doc_path: return
        
        # Temp HTML path
        html_path = os.path.join(TEMP_DIR, "preview.html")
        script_path = os.path.join(SCRIPTS_DIR, "rule_engine.py")
        
        self.web_view.setHtml("<div style='text-align:center; padding-top:50px; font-family:sans-serif;'><h3>Generating High-Fidelity Preview...</h3><p>Please wait, capturing images and formatting.</p></div>")
        self.process_btn.setEnabled(False)
        
        # Use AgentRunner with output capture for errors
        logger = LogManager()
        logger.add_log("Report Tab", f"Generating Preview for {os.path.basename(self.current_doc_path)}...")
        
        self.preview_runner = AgentRunner(sys.executable, [script_path, "--preview", self.current_doc_path, html_path], None)
        self.preview_runner.log_update.connect(lambda msg: logger.add_log("Report Tab", f"[Preview] {msg.strip()}"))
        self.preview_runner.finished.connect(lambda s: self.on_preview_ready(s, html_path))
        self.preview_runner.start()

    def on_preview_ready(self, success, html_path):
        if success and os.path.exists(html_path):
            import time
            url = QUrl.fromLocalFile(html_path)
            # Append timestamp to query to force reload
            url.setQuery(f"t={int(time.time())}")
            
            # Ensure we don't have multiple connections
            try: self.web_view.loadFinished.disconnect(self.make_editable)
            except: pass
            
            self.web_view.loadFinished.connect(self.make_editable)
            self.web_view.setUrl(url)
            self.save_doc_btn.setEnabled(True)
            self.process_btn.setEnabled(True)
        else:
            self.web_view.setHtml("<div style='color:red; text-align:center;'>Preview Generation Failed. Check logs.</div>")
            self.save_doc_btn.setEnabled(False)
            self.process_btn.setEnabled(False)

    def make_editable(self, ok):
        if ok:
            # Inject CSS/JS to make body editable, add Ruler, and improve appearance
            js = """
            (function() {
                document.body.contentEditable = 'true';
                
                // 1. Basic Document Styling
                document.body.style.fontFamily = "'Times New Roman', serif";
                document.body.style.fontSize = "12pt";
                document.body.style.lineHeight = "1.15";
                document.body.style.backgroundColor = "white";
                document.body.style.padding = "1in";
                document.body.style.maxWidth = "8.5in";
                document.body.style.margin = "60px auto 40px auto"; // Increased top margin for better ruler
                document.body.style.boxShadow = "0 0 15px rgba(0,0,0,0.2)";
                document.body.style.minHeight = "10.5in";
                document.body.style.position = "relative";
                document.body.style.tabSize = "4";
                
                document.documentElement.style.backgroundColor = "#e9ecef";
                
                var style = document.createElement('style');
                style.innerHTML = `
                    * { outline: none !important; }
                    
                    /* Headings Mapping */
                    h1 { font-size: 16pt; font-weight: bold; margin-top: 12pt; margin-bottom: 6pt; }
                    h2 { font-size: 14pt; font-weight: bold; margin-top: 10pt; margin-bottom: 4pt; }
                    h3 { font-size: 12pt; font-weight: bold; margin-top: 8pt; margin-bottom: 4pt; }
                    
                    /* Tables */
                    table { border-collapse: collapse; width: 100%; margin: 10pt 0; }
                    /* Fix: Don't force borders on cells, as it reveals invisible layout tables (dashes on blank lines) */
                    td, th { padding: 4pt; min-width: 20pt; }

                    #word-ruler {
                        position: fixed; top: 0; left: 50%; transform: translateX(-50%);
                        width: 8.5in; height: 35px; background: white;
                        border-bottom: 1px solid #adb5bd; z-index: 10000;
                        display: flex; align-items: flex-end; user-select: none;
                        font-family: 'Segoe UI', Arial, sans-serif; font-size: 9px; color: #495057;
                        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
                    }
                    .ruler-bg-margin {
                        position: absolute; top: 0; bottom: 0; background: #e9ecef; width: 1in; z-index: 1;
                    }
                    .ruler-tick { position: absolute; bottom: 0; width: 1px; height: 6px; background: #adb5bd; z-index: 5; }
                    .ruler-num { position: absolute; bottom: 8px; width: 20px; text-align: center; transform: translateX(-50%); z-index: 5; font-weight: 600; }
                    
                    .marker {
                        position: absolute; width: 0; height: 0; 
                        cursor: ew-resize; z-index: 10001;
                    }
                    #first-line-indent {
                        border-left: 5px solid transparent; border-right: 5px solid transparent;
                        border-top: 8px solid #212529; top: 5px;
                    }
                    #hanging-indent {
                        border-left: 5px solid transparent; border-right: 5px solid transparent;
                        border-bottom: 8px solid #212529; bottom: 12px;
                    }
                    #left-base { 
                        position: absolute; width: 10px; height: 8px; 
                        background: #212529; bottom: 3px; cursor: ew-resize; 
                        z-index: 10001; transform: translateX(-50%);
                    }
                    #right-indent {
                        border-left: 5px solid transparent; border-right: 5px solid transparent;
                        border-bottom: 8px solid #212529; bottom: 3px; transform: translateX(50%);
                    }
                    
                    /* Visual page guide */
                    .margin-guide {
                        position: absolute; top: 0; bottom: 0; width: 1px;
                        border-left: 1px dashed rgba(33, 150, 243, 0.4);
                        pointer-events: none; z-index: 9999;
                    }

                    /* List Indentation Styling */
                    ul, ol { margin-left: 0.5in; padding-left: 0; }
                    li { margin-bottom: 4pt; }
                `;
                document.head.appendChild(style);

                // 2. Create Ruler Elements
                if (!document.getElementById('word-ruler')) {
                    var ruler = document.createElement('div');
                    ruler.id = 'word-ruler';
                    
                    // Left and Right Margin Backgrounds
                    var leftMg = document.createElement('div');
                    leftMg.className = 'ruler-bg-margin'; leftMg.style.left = '0';
                    var rightMg = document.createElement('div');
                    rightMg.className = 'ruler-bg-margin'; rightMg.style.right = '0';
                    ruler.appendChild(leftMg);
                    ruler.appendChild(rightMg);

                    // Add Ticks (Quarter-Inch)
                    for (var i = 0; i <= 34; i++) {
                        var posInches = i * 0.25;
                        var leftPos = posInches + 'in';
                        
                        if (i % 4 === 0) {
                            var inchNum = i;
                            if (i > 0 && i < 8) {
                                var num = document.createElement('div');
                                num.className = 'ruler-num';
                                num.innerText = inchNum;
                                num.style.left = leftPos;
                                ruler.appendChild(num);
                            }
                            var tick = document.createElement('div');
                            tick.className = 'ruler-tick';
                            tick.style.height = '10px';
                            tick.style.left = leftPos;
                            ruler.appendChild(tick);
                        } else {
                            var tick = document.createElement('div');
                            tick.className = 'ruler-tick';
                            tick.style.height = (i % 2 === 0) ? '6px' : '4px';
                            tick.style.left = leftPos;
                            ruler.appendChild(tick);
                        }
                    }

                    var firstLine = document.createElement('div');
                    firstLine.id = 'first-line-indent'; firstLine.className = 'marker';
                    
                    var hangingIndent = document.createElement('div');
                    hangingIndent.id = 'hanging-indent'; hangingIndent.className = 'marker';
                    
                    var leftBase = document.createElement('div');
                    leftBase.id = 'left-base';
                    
                    var rightIndent = document.createElement('div');
                    rightIndent.id = 'right-indent'; rightIndent.className = 'marker';
                    
                    ruler.appendChild(firstLine);
                    ruler.appendChild(hangingIndent);
                    ruler.appendChild(leftBase);
                    ruler.appendChild(rightIndent);
                    document.body.parentElement.appendChild(ruler);
                }

                // 3. Ruler Logic
                var ppi = 96; // Pixels per inch
                var snap = 12; // 0.125in snap

                function updateMarkersFromPara() {
                    var sel = window.getSelection();
                    if (sel.rangeCount > 0) {
                        var node = sel.anchorNode;
                        if (node.nodeType === 3) node = node.parentElement;
                        var activePara = node.closest('p, div, h1, h2, h3, li, td') || document.body;
                        
                        var bodyRect = document.body.getBoundingClientRect();
                        var paraRect = activePara.getBoundingClientRect();
                        
                        var style = window.getComputedStyle(activePara);
                        var textIndent = parseFloat(style.textIndent || 0);
                        
                        // Relative to the left edge of the 8.5in page (the body)
                        var leftPos = paraRect.left - bodyRect.left;
                        var rightPos = paraRect.right - bodyRect.left;
                        
                        document.getElementById('first-line-indent').style.left = (leftPos + textIndent) + 'px';
                        document.getElementById('hanging-indent').style.left = leftPos + 'px';
                        document.getElementById('left-base').style.left = leftPos + 'px';
                        document.getElementById('right-indent').style.left = rightPos + 'px';
                    }
                }

                document.addEventListener('selectionchange', updateMarkersFromPara);

                // Dragging logic
                var dragging = null;
                document.addEventListener('mousedown', function(e) {
                    if (e.target.classList.contains('marker') || e.target.id === 'left-base') {
                        dragging = e.target.id;
                        e.preventDefault();
                    }
                });

                document.addEventListener('mousemove', function(e) {
                    if (!dragging) return;
                    
                    var sel = window.getSelection();
                    var node = sel.anchorNode;
                    if (node.nodeType === 3) node = node.parentElement;
                    var activePara = node.closest('p, div, h1, h2, h3, li') || document.body;
                    
                    var rulerRect = document.getElementById('word-ruler').getBoundingClientRect();
                    var pageMargin = 1 * ppi;
                    var xRel = e.clientX - rulerRect.left - pageMargin;
                    var snappedX = Math.round(xRel / snap) * snap;
                    
                    var style = window.getComputedStyle(activePara);
                    var marginLeft = parseFloat(style.marginLeft || 0);
                    var textIndent = parseFloat(style.textIndent || 0);
                    
                    if (dragging === 'first-line-indent') {
                        activePara.style.textIndent = (snappedX - marginLeft) + 'px';
                    } else if (dragging === 'hanging-indent') {
                        // Hanging indent moves margin but keeps first line in place
                        var firstLinePos = marginLeft + textIndent;
                        activePara.style.marginLeft = snappedX + 'px';
                        activePara.style.textIndent = (firstLinePos - snappedX) + 'px';
                    } else if (dragging === 'left-base') {
                        // Square moves both
                        activePara.style.marginLeft = snappedX + 'px';
                    } else if (dragging === 'right-indent') {
                        var rightEdge = 7.5 * ppi;
                        var snappedRight = Math.round((rightEdge - (e.clientX - rulerRect.left)) / snap) * snap;
                        activePara.style.marginRight = snappedRight + 'px';
                    }
                    
                    updateMarkersFromPara();
                });

                document.addEventListener('mouseup', function() { dragging = null; });

                // 4. Keyboard Improvements
                document.body.addEventListener('keydown', function(e) {
                    if (e.key === 'Tab') {
                        e.preventDefault();
                        // Insert a visual tab (atomic span) that is deleted in one go
                        // Using contenteditable=false makes it behave like a single object
                        document.execCommand('insertHTML', false, '<span contenteditable="false" style="display:inline-block; width: 36px;">&nbsp;</span><span>&nbsp;</span>');
                    }
                });

                // 5. Hide bullets on empty lines and Normalize Spacing
                function enforceLayout() {
                    var paras = document.querySelectorAll('p');
                    
                    paras.forEach(function(p) {
                        // Detect Content
                        // Remove zero-width, normal space, and NBSP (\u00A0)
                        var rawText = p.innerText.replace(/[\u200B-\u200D\uFEFF\u00A0\s]/g, ''); 
                        var isEmpty = (rawText === "");
                        
                        // Detect Word Bullet (Dash)
                        var bullet = p.querySelector('span[style*="mso-list:Ignore"], span[style*="mso-list: Ignore"]');
                        if (bullet) {
                            var bulletText = bullet.innerText.replace(/[\u200B-\u200D\uFEFF\u00A0\s]/g, '');
                            // If paragraph text is same as bullet text, it's just a bullet
                            if (rawText === bulletText) {
                                isEmpty = true;
                            }
                        }
                        
                        // Fix Empty/Dash-Only Lines
                        if (isEmpty) {
                            // 1. Remove artifacts (dash, underline)
                            p.innerHTML = '&nbsp;'; 
                            
                            // 2. Force standard single-line height
                            p.style.minHeight = '14pt';
                            p.style.lineHeight = '14pt';
                            
                            // 3. Reset margins for the spacer itself so it doesn't double-up
                            p.style.margin = '0';
                            p.style.padding = '0';
                            p.style.textDecoration = 'none';
                            p.style.border = 'none';
                        }
                        // Note: We DO NOT touch non-empty paragraphs. 
                        // They keep their CSS/Word styles.
                    });
                }
                
                enforceLayout();
                // Re-run on changes
                document.addEventListener('input', enforceLayout);

                console.log("Improved Word-style Ruler and Formatting initialized.");
            })();
            """
            self.web_view.page().runJavaScript(js)

    def on_custom_url(self, url):
        if url.toString() == "icharlotte://show-find":
            self.find_bar.show()
            self.find_input.setFocus()
            self.find_input.selectAll()

    def process_document(self):
        if not self.current_doc_path: return
        
        # 1. Save current active rules to a temp json file
        temp_rules_path = os.path.join(TEMP_DIR, "active_rules.json")
        try:
            with open(temp_rules_path, 'w') as f:
                json.dump(self.rules, f)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save rules: {e}")
            return
            
        # 2. Run rule_engine.py --apply
        script_path = os.path.join(SCRIPTS_DIR, "rule_engine.py")
        
        # Clear preview to release any file locks? 
        # WebEngineView shouldn't lock the .docx (it locks the .html preview), but let's be safe.
        self.web_view.setHtml("<div style='text-align:center; padding-top:50px; color:blue;'>Processing Rules... Please Wait.</div>")
        self.process_btn.setEnabled(False)
        
        if self.main_window:
            self.main_window.tabs.setCurrentIndex(1) # Switch to Status Tab (Index 1) automatically? 
            # Or just add the task. Let's just add the task so user can switch if they want.
            # User complained "nothing in status tab", so putting it there is key.
            # Maybe switching to it is helpful but intrusive? 
            # "it applies the changes to quickly for me to really look at teh status tab"
            # If we switch, they see it.
            
            # Use add_status_task
            self.process_runner = self.main_window.add_status_task(
                "Rule Engine", 
                f"Applying to {os.path.basename(self.current_doc_path)}",
                sys.executable,
                [script_path, "--apply", self.current_doc_path, temp_rules_path]
            )
            
            # --- LOGGING INTEGRATION ---
            logger = LogManager()
            logger.add_log("Report Tab", f"Starting Rule Engine for {self.current_doc_path}")
            self.process_runner.log_update.connect(
                lambda msg: logger.add_log("Report Tab", msg.strip())
            )
            # ---------------------------

            self.process_runner.finished.connect(self.on_process_finished)
            
        else:
            # Fallback
            self.process_runner = AgentRunner(sys.executable, [script_path, "--apply", self.current_doc_path, temp_rules_path], None)
            
            # --- LOGGING INTEGRATION ---
            logger = LogManager()
            logger.add_log("Report Tab", f"Starting Rule Engine (Fallback) for {self.current_doc_path}")
            self.process_runner.log_update.connect(
                lambda msg: logger.add_log("Report Tab", msg.strip())
            )
            # ---------------------------
            
            self.process_runner.finished.connect(self.on_process_finished)
            self.process_runner.start()

    def on_process_finished(self, success):
        logger = LogManager()
        if success:
            logger.add_log("Report Tab", "Rule Engine completed successfully.")
        else:
            logger.add_log("Report Tab", "Rule Engine failed.")
            
        self.process_btn.setEnabled(True)
        if success:
            QMessageBox.information(self, "Success", "Rules applied successfully!")
            self.update_preview() # Refresh the view
        else:
            QMessageBox.critical(self, "Error", "Failed to apply rules.")
            self.web_view.setHtml("<div style='color:red; text-align:center;'>Processing Failed</div>")
