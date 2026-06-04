"""Offscreen smoke test: PromptTemplateDialog writes to the global store. Run directly:
   QT_QPA_PLATFORM=offscreen python tests/smoke_global_prompts_dialog.py
"""
import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
import sys
import tempfile

# Ensure the project root is on sys.path when run as a script.
_project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _project_root not in sys.path:
    sys.path.insert(0, _project_root)

from PySide6.QtWidgets import QApplication, QMessageBox

app = QApplication.instance() or QApplication(sys.argv)

# Silence modal popups so the headless run does not block.
QMessageBox.information = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Ok)
QMessageBox.warning = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Ok)
QMessageBox.question = staticmethod(lambda *a, **k: QMessageBox.StandardButton.Yes)

# Point the global singleton at a temp store.
import icharlotte_core.chat.global_prompts as gp
tmp = tempfile.mkdtemp()
gp._global_store = gp.GlobalQuickPromptStore(
    path=os.path.join(tmp, "g.json"),
    case_data_dir=os.path.join(tmp, "case_data"),
    auto_migrate=False,
)

from icharlotte_core.ui.chat_dialogs import PromptTemplateDialog

dlg = PromptTemplateDialog(persistence=None)
dlg.add_prompt()
dlg.name_edit.setText("Smoke Template")
dlg.prompt_edit.setPlainText("Smoke prompt body")
dlg.category_edit.setCurrentText("Custom")
dlg.save_prompt()

stored = gp.get_global_quick_prompt_store().get_quick_prompts()
assert any(p.name == "Smoke Template" and p.prompt == "Smoke prompt body" for p in stored), stored

# Survives a fresh store instance (cross-session).
fresh = gp.GlobalQuickPromptStore(path=gp._global_store.path,
                                  case_data_dir=gp._global_store.case_data_dir,
                                  auto_migrate=False)
assert any(p.name == "Smoke Template" for p in fresh.get_quick_prompts())
print("UI SMOKE OK")
