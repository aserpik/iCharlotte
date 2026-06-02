"""Wizard Mode "Separate Documents" task.

Settings (pick sensitivity, Analyze) -> Status (analyzing) -> Workbench
(embedded SeparatorWorkbench for review + split/merge). Backed by
SeparateAnalysisWorker, which runs Scripts/separate.py --headless to produce
the document map (and the Word index, as a side effect, exactly like Advanced
Mode).
"""
import json
import os
import re
import subprocess
import sys

from PySide6.QtCore import QThread, Signal

from icharlotte_core.config import SCRIPTS_DIR


class SeparateAnalysisWorker(QThread):
    progress = Signal(str)
    finished_analysis = Signal(bool, object)  # (success, list[dict] | error str)

    def __init__(self, pdf_path: str, sensitivity: int, parent=None):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.sensitivity = sensitivity

    @staticmethod
    def _parse_docs(stdout: str):
        match = re.search(r"JSON_MAP:\s*(.+)", stdout)
        if not match:
            return None
        json_path = match.group(1).strip()
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                docs = json.load(f)
        except Exception:
            return None
        try:
            os.remove(json_path)
        except OSError:
            pass
        return docs

    def run(self):
        try:
            script_path = os.path.join(SCRIPTS_DIR, "separate.py")
            self.progress.emit(f"Analyzing {os.path.basename(self.pdf_path)}...")
            proc = subprocess.run(
                [sys.executable, script_path, "--headless",
                 "--sensitivity", str(self.sensitivity), self.pdf_path],
                capture_output=True, text=True, encoding="utf-8", errors="replace",
            )
            for line in (proc.stdout or "").splitlines():
                if line.strip():
                    self.progress.emit(line.strip())
            if proc.returncode != 0:
                self.finished_analysis.emit(
                    False, f"Analysis failed (exit {proc.returncode}). {proc.stderr[-500:]}")
                return
            docs = self._parse_docs(proc.stdout or "")
            if docs is None:
                self.finished_analysis.emit(
                    False, "Analysis completed but no document map was produced.")
                return
            self.finished_analysis.emit(True, docs)
        except Exception as e:
            self.finished_analysis.emit(False, str(e))
