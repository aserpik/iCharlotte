"""
Office Hotkey Integration - Global hotkey to trigger LLM processing on Word/Outlook selections.

Press Win+V to show a popup with prompt options when Word or Outlook compose window is active.
The selected/entered prompt is sent to the LLM along with any selected text, and the result
replaces the selection (or inserts at cursor if nothing selected).

Supports:
- Microsoft Word documents
- Microsoft Outlook email compose windows (new email, reply, forward)

The popup automatically detects which application is active and shows context-appropriate
prompts (document prompts for Word, email prompts for Outlook).
"""

import os
import json
import threading
import logging
import uuid
from dataclasses import dataclass, field
from typing import Optional, Callable, Dict, Any

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QLineEdit, QPushButton, QMessageBox, QApplication,
    QTextEdit, QFrame, QSizePolicy, QCheckBox, QSpinBox,
    QGroupBox, QGridLayout, QFontComboBox, QDoubleSpinBox,
    QTabWidget, QWidget, QScrollArea, QFileDialog
)
from PySide6.QtCore import Qt, Signal, QObject, QTimer, QThread
from PySide6.QtGui import QFont, QKeySequence, QShortcut
import re

# Import MasterCaseDatabase for case variable insertion
try:
    from .master_db import MasterCaseDatabase
    HAS_MASTER_DB = True
except ImportError:
    HAS_MASTER_DB = False

try:
    import keyboard
    HAS_KEYBOARD = True
except ImportError:
    HAS_KEYBOARD = False

# Windows RegisterHotKey API - more reliable than keyboard library hooks
import ctypes
import ctypes.wintypes as wintypes
import threading
_WM_HOTKEY = 0x0312
_MOD_WIN = 0x0008
_MOD_NOREPEAT = 0x4000
_VK_V = 0x56
_HOTKEY_ID = 0xBFFF  # Arbitrary ID in valid range

try:
    import win32com.client
    import pythoncom
    import win32gui
    import win32process
    import subprocess
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

logger = logging.getLogger("icharlotte.word_hotkey")

# Set up file logging for word_hotkey so COM errors are captured
def _setup_word_hotkey_logger():
    import datetime
    from logging.handlers import RotatingFileHandler
    _log_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'logs')
    os.makedirs(_log_dir, exist_ok=True)
    log_file = os.path.join(_log_dir, f"icharlotte_{datetime.datetime.now().strftime('%Y-%m-%d')}.log")
    handler = RotatingFileHandler(log_file, maxBytes=10*1024*1024, backupCount=7, encoding='utf-8')
    handler.setLevel(logging.DEBUG)
    handler.setFormatter(logging.Formatter(
        '%(asctime)s [%(name)s] %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    ))
    logger.addHandler(handler)
    logger.setLevel(logging.DEBUG)

_setup_word_hotkey_logger()

try:
    from .document_processor import DocumentProcessor
    HAS_DOC_PROCESSOR = True
except ImportError:
    HAS_DOC_PROCESSOR = False

# Application context types
APP_CONTEXT_UNKNOWN = "unknown"
APP_CONTEXT_WORD = "word"
APP_CONTEXT_OUTLOOK = "outlook"

# Window class names
WORD_WINDOW_CLASS = "OpusApp"
OUTLOOK_INSPECTOR_CLASS = "rctrl_renwnd32"

# Default email prompts
DEFAULT_OUTLOOK_PROMPTS = [
    {"name": "Make Professional", "prompt": "Rewrite this email in a professional, courteous tone while maintaining the original meaning and intent."},
    {"name": "Shorten Email", "prompt": "Condense this email to be more concise while preserving all key points and action items."},
    {"name": "Fix Grammar", "prompt": "Fix any grammar, spelling, or punctuation errors in this email."},
    {"name": "Soften Tone", "prompt": "Rewrite this email with a softer, more diplomatic tone while keeping the message clear."},
    {"name": "Add Clarity", "prompt": "Improve clarity and structure of this email. Ensure requests and next steps are clearly stated."},
]

# Email-specific system prompt
EMAIL_SYSTEM_PROMPT = (
    "You are a helpful email writing assistant. Follow the user's instructions precisely. "
    "Output ONLY the processed body text - never include Subject lines, To/From/CC headers, "
    "greetings, signatures, or any email metadata. Just output the revised text content directly. "
    "Do not add any preamble or explanation."
)


def detect_active_app_context() -> tuple:
    """
    Detect whether Word or Outlook compose window is active.
    Returns: (context_type, inspector_object_or_None)
    """
    if not HAS_WIN32:
        return APP_CONTEXT_UNKNOWN, None

    try:
        # Get the foreground window
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return APP_CONTEXT_UNKNOWN, None

        class_name = win32gui.GetClassName(hwnd)

        # Check for Word
        if class_name == WORD_WINDOW_CLASS:
            return APP_CONTEXT_WORD, None

        # Check for Outlook Inspector (compose/reply window)
        if class_name == OUTLOOK_INSPECTOR_CLASS:
            # Verify it's a compose window via COM
            outlook = None
            inspector = None
            word_editor = None
            try:
                pythoncom.CoInitialize()
                outlook = win32com.client.GetActiveObject("Outlook.Application")
                inspector = outlook.ActiveInspector()
                if inspector:
                    # Check if the inspector has a WordEditor (compose mode)
                    try:
                        word_editor = inspector.WordEditor
                        if word_editor:
                            del word_editor; word_editor = None
                            del outlook; outlook = None
                            return APP_CONTEXT_OUTLOOK, inspector
                    except:
                        pass
            except Exception as e:
                logger.debug(f"Outlook COM check failed: {e}")
            finally:
                # Release COM objects to prevent GC crashes
                try:
                    if word_editor is not None:
                        del word_editor
                    if inspector is not None:
                        del inspector
                    if outlook is not None:
                        del outlook
                except Exception:
                    pass

        return APP_CONTEXT_UNKNOWN, None
    except Exception as e:
        logger.error(f"COM error detecting app context: {type(e).__name__}: {e}", exc_info=True)
        return APP_CONTEXT_UNKNOWN, None


def detect_case_from_document() -> tuple:
    """
    Detect case info from the currently open Word document's file path.
    Returns: (case_dict_or_None, file_number_or_None, error_message_or_None)

    Parses the document path for file number pattern (####.###) and looks up
    the case in MasterCaseDatabase.
    """
    if not HAS_WIN32:
        return None, None, "Win32 not available"

    if not HAS_MASTER_DB:
        return None, None, "Database not available"

    doc = None
    word = None
    try:
        pythoncom.CoInitialize()

        # Try to connect to Word
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except:
            try:
                word = win32com.client.GetObject(None, "Word.Application")
            except:
                pass

        if not word:
            return None, None, "Word not running"

        doc = word.ActiveDocument
        if not doc:
            del word; word = None
            return None, None, "No active document"

        # Get the document's full path
        try:
            doc_path = doc.FullName
        except:
            del doc; doc = None
            del word; word = None
            return None, None, "Document not saved"

        # Release COM objects immediately — we only need the string path
        del doc; doc = None
        del word; word = None

        if not doc_path or doc_path.startswith("Document"):
            return None, None, "Document not saved"

        # Try multiple patterns to extract file number from path
        file_number = None

        # Pattern 1: Direct match for ####.### format (e.g., "1280.001 - Name")
        direct_pattern = r'(\d{4}\.\d{3})'
        direct_matches = re.findall(direct_pattern, doc_path)
        if direct_matches:
            file_number = direct_matches[0]

        # Pattern 2: Split folder structure (e.g., "3800- NATIONWIDE\167 - Nikitina")
        # Carrier folder: 4 digits followed by optional space and hyphen
        # Case folder: 1-3 digits followed by space and hyphen
        if not file_number:
            # Look for carrier folder pattern: ####- or #### -
            carrier_pattern = r'[\\/](\d{4})\s*-\s*[^\\/]+'
            # Look for case subfolder pattern: ### - (1-3 digits)
            case_pattern = r'[\\/](\d{1,3})\s*-\s*[^\\/]+'

            carrier_matches = re.findall(carrier_pattern, doc_path)
            case_matches = re.findall(case_pattern, doc_path)

            if carrier_matches and case_matches:
                # Use the last carrier and case match (most specific to the document)
                carrier = carrier_matches[-1]
                case_num = case_matches[-1].zfill(3)  # Pad to 3 digits
                file_number = f"{carrier}.{case_num}"

        if not file_number:
            # Show truncated path for debugging
            short_path = doc_path if len(doc_path) <= 60 else "..." + doc_path[-57:]
            return None, None, f"No case number in: {short_path}"

        # Look up in database
        try:
            db = MasterCaseDatabase()
            case = db.get_case(file_number)
            if case:
                return case, file_number, None
            else:
                return None, file_number, f"Case {file_number} not in database"
        except Exception as e:
            return None, file_number, f"Database error: {str(e)}"

    except Exception as e:
        logger.error(f"COM error in detect_case_from_document: {type(e).__name__}: {e}", exc_info=True)
        return None, None, f"Error: {str(e)}"
    finally:
        # Ensure COM objects are released to prevent GC crashes
        try:
            if doc is not None:
                del doc
            if word is not None:
                del word
        except Exception:
            pass


_last_zombie_check = 0.0  # timestamp of last kill_zombie run

def kill_zombie_word_processes():
    """Kill Word processes that don't have visible windows (zombies)."""
    global _last_zombie_check
    if not HAS_WIN32:
        return

    # Skip if checked recently (avoid blocking main thread with subprocess)
    import time as _time
    now = _time.time()
    if now - _last_zombie_check < 30.0:
        return
    _last_zombie_check = now

    try:
        # Find PIDs of visible Word windows
        visible_pids = set()

        def enum_callback(hwnd, results):
            if win32gui.IsWindowVisible(hwnd):
                class_name = win32gui.GetClassName(hwnd)
                if class_name == "OpusApp":
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    results.add(pid)
            return True

        win32gui.EnumWindows(enum_callback, visible_pids)

        if not visible_pids:
            # No visible Word windows, don't kill anything
            return

        # Find all Word processes
        result = subprocess.run(
            ['tasklist', '/FI', 'IMAGENAME eq WINWORD.EXE', '/FO', 'CSV'],
            capture_output=True, text=True, timeout=5
        )

        all_pids = set()
        lines = result.stdout.strip().split('\n')
        for line in lines[1:]:  # Skip header
            parts = line.replace('"', '').split(',')
            if len(parts) >= 2:
                try:
                    all_pids.add(int(parts[1]))
                except:
                    pass

        # Kill zombie processes (those without visible windows)
        zombie_pids = all_pids - visible_pids
        for pid in zombie_pids:
            try:
                subprocess.run(
                    ['taskkill', '/PID', str(pid), '/F'],
                    capture_output=True, timeout=5
                )
                print(f"Killed zombie Word process: PID {pid}")
            except Exception as e:
                print(f"Failed to kill PID {pid}: {e}")

    except Exception as e:
        print(f"Error killing zombie processes: {e}")

from icharlotte_core.config import GEMINI_DATA_DIR
from icharlotte_core.redline_config import load_redline_settings, save_redline_settings


# Format options
FORMAT_PLAIN = "Plain Text"
FORMAT_MATCH = "Match Selection"
FORMAT_MARKDOWN = "Parse Markdown"
FORMAT_DEFAULT = "Document Default"
FORMAT_CUSTOM = "Custom..."

# Available LLM models for selection
# Format: (display_name, provider, model_id)
AVAILABLE_MODELS = [
    ("Gemini 2.5 Flash (Fast)", "Gemini", "gemini-2.5-flash"),
    ("Gemini 2.5 Pro", "Gemini", "gemini-2.5-pro"),
    ("Gemini 2.0 Flash", "Gemini", "models/gemini-2.0-flash"),
    ("Gemini 3.1 Pro Preview", "Gemini", "gemini-3.1-pro-preview"),
    ("Gemini 3.1 Flash Lite Preview", "Gemini", "gemini-3.1-flash-lite-preview"),
    ("Gemini 3 Pro Preview", "Gemini", "gemini-3-pro-preview"),
    ("Gemini 3 Flash Preview", "Gemini", "gemini-3-flash-preview"),
    ("Claude Sonnet 4", "Claude", "claude-sonnet-4-20250514"),
    ("Claude Haiku 4 (Fast)", "Claude", "claude-haiku-4-20250514"),
    ("GPT-4o", "OpenAI", "gpt-4o"),
    ("GPT-4o Mini (Fast)", "OpenAI", "gpt-4o-mini"),
]

DEFAULT_MODEL_INDEX = 0  # Gemini 2.5 Flash


# ════════════════════════════════════════════════════════════════════════════
# Parallel Task Infrastructure — TaskData, bookmarks, TaskManager, StatusBar
# ════════════════════════════════════════════════════════════════════════════

@dataclass
class TaskData:
    """All state needed to insert an LLM result into Word after async completion."""
    task_id: str = field(default_factory=lambda: str(uuid.uuid4())[:8])
    bookmark_name: str = ""
    document_name: str = ""
    document_com: Any = None
    has_selection: bool = False
    original_text: str = ""
    original_text_raw: str = ""
    captured_format: Optional[Dict] = None
    format_type: str = FORMAT_PLAIN
    redline_mode_active: bool = False
    redline_settings: Optional[Dict] = None
    research_result: Any = None
    do_legal_research: bool = False  # Whether to run legal research in worker thread
    legal_research_engine: Any = None  # LegalResearchEngine instance (if research requested)
    legal_research_model: Optional[tuple] = None  # (provider, model_id) for research LLM calls
    prompt_preview: str = ""
    provider: str = ""
    model_id: str = ""
    system_prompt: str = ""
    full_prompt: str = ""
    status: str = "pending"       # pending -> running -> inserting -> done -> error
    result_text: str = ""
    error_msg: str = ""


def _create_task_bookmark(doc_com, range_start: int, range_end: int, task_id: str) -> str:
    """Create a temporary Word bookmark at the given range.

    Bookmarks auto-adjust their positions when text is inserted/deleted
    elsewhere in the document, solving the range displacement problem
    for parallel tasks.
    """
    bookmark_name = f"_ICHARLOTTE_{task_id}"
    try:
        # Clamp to document bounds to avoid "Value out of range" if the
        # document was edited between capture time and bookmark creation.
        doc_end = doc_com.Content.End
        range_start = min(range_start, doc_end - 1)
        range_end = min(range_end, doc_end - 1)
        if range_start < 0:
            range_start = 0
        if range_end < range_start:
            range_end = range_start
        target_range = doc_com.Range(range_start, range_end)
        doc_com.Bookmarks.Add(Name=bookmark_name, Range=target_range)
        print(f"Created bookmark '{bookmark_name}' at {range_start}-{range_end}")
        return bookmark_name
    except Exception as e:
        print(f"Failed to create bookmark: {e}")
        return ""


def _get_bookmark_range(doc_com, bookmark_name: str):
    """Retrieve the current range of a bookmark. Returns (start, end) or None."""
    try:
        if doc_com.Bookmarks.Exists(bookmark_name):
            bm = doc_com.Bookmarks(bookmark_name)
            return bm.Range.Start, bm.Range.End
        else:
            print(f"Bookmark '{bookmark_name}' no longer exists")
            return None
    except Exception as e:
        print(f"Failed to get bookmark range: {e}")
        return None


def _delete_task_bookmark(doc_com, bookmark_name: str):
    """Delete a temporary bookmark after insertion is complete."""
    try:
        if doc_com.Bookmarks.Exists(bookmark_name):
            doc_com.Bookmarks(bookmark_name).Delete()
            print(f"Deleted bookmark '{bookmark_name}'")
    except Exception as e:
        print(f"Failed to delete bookmark '{bookmark_name}': {e}")


class CustomFormatDialog(QDialog):
    """Dialog for specifying custom text formatting."""

    def __init__(self, parent=None, current_settings=None):
        super().__init__(parent)
        self.setWindowTitle("Custom Format Settings")
        self.setMinimumWidth(350)
        self.settings = current_settings or {}

        self.setStyleSheet("""
            QDialog {
                background-color: #1e1e2e;
            }
            QLabel {
                color: #cdd6f4;
                font-size: 12px;
            }
            QGroupBox {
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 5px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px;
            }
            QComboBox, QSpinBox, QDoubleSpinBox, QFontComboBox {
                background-color: #313244;
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 5px;
                padding: 5px;
            }
            QCheckBox {
                color: #cdd6f4;
            }
            QCheckBox::indicator {
                width: 16px;
                height: 16px;
            }
            QPushButton {
                background-color: #6c63ff;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7c73ff;
            }
            QPushButton#cancelBtn {
                background-color: #45475a;
            }
        """)

        self.setup_ui()
        self.load_settings()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Font Group
        font_group = QGroupBox("Font")
        font_layout = QGridLayout(font_group)

        font_layout.addWidget(QLabel("Font:"), 0, 0)
        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(QFont("Times New Roman"))
        font_layout.addWidget(self.font_combo, 0, 1)

        font_layout.addWidget(QLabel("Size:"), 1, 0)
        self.size_spin = QSpinBox()
        self.size_spin.setRange(6, 72)
        self.size_spin.setValue(12)
        font_layout.addWidget(self.size_spin, 1, 1)

        layout.addWidget(font_group)

        # Style Group
        style_group = QGroupBox("Style")
        style_layout = QHBoxLayout(style_group)

        self.bold_check = QCheckBox("Bold")
        self.italic_check = QCheckBox("Italic")
        self.underline_check = QCheckBox("Underline")

        style_layout.addWidget(self.bold_check)
        style_layout.addWidget(self.italic_check)
        style_layout.addWidget(self.underline_check)
        style_layout.addStretch()

        layout.addWidget(style_group)

        # Paragraph Group
        para_group = QGroupBox("Paragraph")
        para_layout = QGridLayout(para_group)

        para_layout.addWidget(QLabel("First Line Indent:"), 0, 0)
        self.first_indent_spin = QDoubleSpinBox()
        self.first_indent_spin.setRange(0, 2)
        self.first_indent_spin.setSingleStep(0.25)
        self.first_indent_spin.setValue(0)
        self.first_indent_spin.setSuffix(" in")
        para_layout.addWidget(self.first_indent_spin, 0, 1)

        para_layout.addWidget(QLabel("Left Indent:"), 1, 0)
        self.left_indent_spin = QDoubleSpinBox()
        self.left_indent_spin.setRange(0, 4)
        self.left_indent_spin.setSingleStep(0.25)
        self.left_indent_spin.setValue(0)
        self.left_indent_spin.setSuffix(" in")
        para_layout.addWidget(self.left_indent_spin, 1, 1)

        para_layout.addWidget(QLabel("Line Spacing:"), 2, 0)
        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems(["Single", "1.5 Lines", "Double"])
        para_layout.addWidget(self.line_spacing_combo, 2, 1)

        layout.addWidget(para_group)

        # Buttons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        cancel_btn = QPushButton("Cancel")
        cancel_btn.setObjectName("cancelBtn")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)

        ok_btn = QPushButton("OK")
        ok_btn.clicked.connect(self.accept)
        btn_layout.addWidget(ok_btn)

        layout.addLayout(btn_layout)

    def load_settings(self):
        """Load settings from dictionary."""
        if not self.settings:
            return

        if 'font_name' in self.settings:
            self.font_combo.setCurrentFont(QFont(self.settings['font_name']))
        if 'font_size' in self.settings:
            self.size_spin.setValue(self.settings['font_size'])
        if 'bold' in self.settings:
            self.bold_check.setChecked(self.settings['bold'])
        if 'italic' in self.settings:
            self.italic_check.setChecked(self.settings['italic'])
        if 'underline' in self.settings:
            self.underline_check.setChecked(self.settings['underline'])
        if 'first_indent' in self.settings:
            self.first_indent_spin.setValue(self.settings['first_indent'])
        if 'left_indent' in self.settings:
            self.left_indent_spin.setValue(self.settings['left_indent'])
        if 'line_spacing' in self.settings:
            idx = self.line_spacing_combo.findText(self.settings['line_spacing'])
            if idx >= 0:
                self.line_spacing_combo.setCurrentIndex(idx)

    def get_settings(self) -> dict:
        """Get the format settings as a dictionary."""
        return {
            'font_name': self.font_combo.currentFont().family(),
            'font_size': self.size_spin.value(),
            'bold': self.bold_check.isChecked(),
            'italic': self.italic_check.isChecked(),
            'underline': self.underline_check.isChecked(),
            'first_indent': self.first_indent_spin.value(),
            'left_indent': self.left_indent_spin.value(),
            'line_spacing': self.line_spacing_combo.currentText()
        }


def apply_flat_diff_redline(doc_com, range_start, range_end, original_text,
                            revised_text, auto_enable_track_changes=True):
    """Apply tracked changes using flat diff mapped to Word paragraphs.

    Standalone reusable function extracted from the AI Assistant's redline engine.
    Can be called from both the AI Assistant and the ReportReviewer.

    Strategy: flatten both texts (strip paragraph breaks), diff them
    character-by-character, then map each change back to the correct
    Word paragraph. This never touches paragraph marks (\\r), so
    formatting is always preserved regardless of LLM paragraph behavior.

    Args:
        doc_com: win32com Word Document object
        range_start: Start character position of the target range
        range_end: End character position of the target range
        original_text: Original text before LLM processing
        revised_text: LLM-revised text
        auto_enable_track_changes: Whether to auto-enable Track Changes

    Returns:
        dict with keys: success (bool), total_changes (int),
        equal_chars (int), deleted_chars (int), inserted_chars (int),
        pre_content_para_count (int), orig_flat_length (int)
    """
    result = {
        'success': False, 'total_changes': 0,
        'equal_chars': 0, 'deleted_chars': 0, 'inserted_chars': 0,
        'pre_content_para_count': 0, 'orig_flat_length': 0,
    }

    try:
        from diff_match_patch import diff_match_patch
        import bisect
        import re
        import time as _time

        if not doc_com:
            print("No document available")
            return result

        # Auto-enable Track Changes if needed
        if auto_enable_track_changes:
            try:
                if not doc_com.TrackRevisions:
                    print("Auto-enabling Track Changes for redlining")
                    doc_com.TrackRevisions = True
            except Exception as e:
                print(f"Could not enable Track Changes: {e}")

        # Reconstruct range
        try:
            range_obj = doc_com.Range(range_start, range_end)
        except Exception as e:
            print(f"Could not reconstruct range: {e}")
            return result

        # --- 1. Build Word paragraph map ---
        # \x01 = inline shape (image/snippet), \x07 = cell mark
        INLINE_SHAPE_CHAR = '\x01'
        word_paras = []
        try:
            para_collection = range_obj.Paragraphs
            para_count = para_collection.Count
            for i in range(1, para_count + 1):
                wp = para_collection(i)
                text = wp.Range.Text
                content = text[:-1] if text.endswith('\r') else text
                # Track positions of inline shape chars within this paragraph
                # so we can exclude them from the flat diff text but preserve
                # their Word positions for correct offset mapping.
                shape_offsets = [j for j, ch in enumerate(content)
                                 if ch == INLINE_SHAPE_CHAR]
                clean_content = content.replace(INLINE_SHAPE_CHAR, '')
                word_paras.append({
                    'start': wp.Range.Start,
                    'end': wp.Range.End,
                    'text': text,
                    'content': content,
                    'clean_content': clean_content,
                    'shape_offsets': shape_offsets,
                })
                if shape_offsets:
                    print(f"  Para [{i-1}]: {len(shape_offsets)} inline shape(s) detected — preserved")
        except Exception as e:
            print(f"Could not enumerate Word paragraphs: {e}")
            return result

        print(f"Word paragraphs: {len(word_paras)}")

        # --- 2. Build flat original text with boundary spaces ---
        # Use clean_content (no \x01 inline shape chars) so the diff
        # never sees or tries to delete inline shapes.
        orig_flat = ''
        flat_para_starts = []
        boundary_positions = set()

        for i, wp in enumerate(word_paras):
            if i > 0 and wp['clean_content']:
                prev_had_content = any(word_paras[j]['clean_content'] for j in range(i))
                if prev_had_content:
                    boundary_positions.add(len(orig_flat))
                    orig_flat += ' '
            flat_para_starts.append(len(orig_flat))
            orig_flat += wp['clean_content']
        flat_para_starts.append(len(orig_flat))  # sentinel

        print(f"Original flat: {len(orig_flat)} chars, "
              f"{len(boundary_positions)} boundary spaces")

        # --- 3. Flatten revised text ---
        rev_flat = revised_text.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ')
        rev_flat = re.sub(r'  +', ' ', rev_flat).strip()

        print(f"Revised flat: {len(rev_flat)} chars")

        # --- 4. Diff the flat texts ---
        dmp = diff_match_patch()
        diffs = dmp.diff_main(orig_flat, rev_flat)
        dmp.diff_cleanupSemantic(diffs)

        eq = sum(len(t) for op, t in diffs if op == 0)
        de = sum(len(t) for op, t in diffs if op == -1)
        ins = sum(len(t) for op, t in diffs if op == 1)
        pct = eq / len(orig_flat) * 100 if orig_flat else 0
        print(f"Flat diff: {eq} equal, {de} del, {ins} ins ({pct:.1f}% preserved)")

        # --- 5. Build operations, filtering boundary artifacts ---
        operations = []
        pos = 0
        for op, text in diffs:
            if op == 0:
                pos += len(text)
            elif op == -1:
                all_boundary = all(
                    pos + i in boundary_positions for i in range(len(text))
                )
                if not all_boundary:
                    i = pos
                    real_end = pos + len(text)
                    while i < real_end:
                        if i in boundary_positions:
                            i += 1
                            continue
                        seg_start = i
                        while i < real_end and i not in boundary_positions:
                            i += 1
                        operations.append((
                            'delete', seg_start, i, orig_flat[seg_start:i]
                        ))
                pos += len(text)
            elif op == 1:
                if text == ' ' and pos in boundary_positions:
                    pass  # skip boundary space insert
                else:
                    operations.append(('insert', pos, text))

        print(f"Operations to apply: {len(operations)}")

        # --- 6. Map flat positions to Word and apply in reverse ---
        _time.sleep(0.2)

        import pywintypes as _pywintypes

        def _com_op(func, retries=3, delay=2.0):
            """Retry a COM operation on RPC_E_CALL_REJECTED."""
            for attempt in range(retries):
                try:
                    return func()
                except _pywintypes.com_error as e:
                    if e.hresult == -2147418111 and attempt < retries - 1:
                        print(f"    COM busy, retrying in {delay}s "
                              f"(attempt {attempt + 1})...")
                        _time.sleep(delay)
                    else:
                        raise

        total_changes = 0

        def flat_to_word_pos(flat_pos):
            para_idx = bisect.bisect_right(flat_para_starts, flat_pos) - 1
            para_idx = max(0, min(para_idx, len(word_paras) - 1))
            offset = flat_pos - flat_para_starts[para_idx]
            # Adjust offset to account for inline shape chars (\x01) that
            # were stripped from clean_content but still occupy Word positions.
            shape_offsets = word_paras[para_idx].get('shape_offsets', [])
            if shape_offsets:
                adjusted = offset
                for so in shape_offsets:
                    if so <= adjusted:
                        adjusted += 1
                return word_paras[para_idx]['start'] + adjusted
            return word_paras[para_idx]['start'] + offset

        for operation in reversed(operations):
            if operation[0] == 'delete':
                _, flat_start, flat_end, del_text = operation
                start_para = bisect.bisect_right(
                    flat_para_starts, flat_start) - 1
                end_para = bisect.bisect_right(
                    flat_para_starts, flat_end - 1) - 1

                if start_para == end_para:
                    word_start = flat_to_word_pos(flat_start)
                    word_end = flat_to_word_pos(flat_end)
                    # Never delete past the paragraph mark
                    para_content_end = word_paras[start_para]['end'] - 1
                    if word_end > para_content_end:
                        word_end = para_content_end
                    if word_start >= word_end:
                        continue
                    # Skip trailing whitespace at paragraph end
                    if word_end == para_content_end and del_text.strip() == '':
                        continue
                    dr = _com_op(lambda ws=word_start, we=word_end: doc_com.Range(ws, we))
                    dr_text = _com_op(lambda: dr.Text[:50])
                    print(f"  DEL in para [{start_para}]: "
                          f"{repr(dr_text)}")
                    _com_op(lambda: dr.Delete())
                    total_changes += 1
                else:
                    for pidx in reversed(range(start_para, end_para + 1)):
                        p_start = flat_para_starts[pidx]
                        p_end = flat_para_starts[pidx + 1]
                        clip_start = max(flat_start, p_start)
                        clip_end = min(flat_end, p_end)
                        if clip_start >= clip_end:
                            continue
                        ws = word_paras[pidx]['start'] + (clip_start - p_start)
                        we = word_paras[pidx]['start'] + (clip_end - p_start)
                        para_content_end = word_paras[pidx]['end'] - 1
                        if we > para_content_end:
                            we = para_content_end
                        if ws >= we:
                            continue
                        dr = _com_op(lambda _ws=ws, _we=we: doc_com.Range(_ws, _we))
                        dr_text = _com_op(lambda: dr.Text[:50])
                        print(f"  DEL in para [{pidx}]: "
                              f"{repr(dr_text)}")
                        _com_op(lambda: dr.Delete())
                        total_changes += 1

            elif operation[0] == 'insert':
                _, flat_pos, ins_text = operation
                word_pos = flat_to_word_pos(flat_pos)
                ir = _com_op(lambda wp=word_pos: doc_com.Range(wp, wp))
                _com_op(lambda: ir.InsertAfter(ins_text))
                para_idx = bisect.bisect_right(
                    flat_para_starts, flat_pos) - 1
                print(f"  INS in para [{para_idx}]: "
                      f"{repr(ins_text[:50])}")
                total_changes += 1

        print(f"Applied {total_changes} redline changes")

        pre_content_paras = sum(
            1 for wp in word_paras if wp['content'].strip()
        )

        result.update({
            'success': True,
            'total_changes': total_changes,
            'equal_chars': eq,
            'deleted_chars': de,
            'inserted_chars': ins,
            'pre_content_para_count': pre_content_paras,
            'orig_flat_length': len(orig_flat),
        })
        return result

    except ImportError as e:
        print(f"diff-match-patch not available: {e}")
        return result
    except Exception as e:
        print(f"Unexpected error in apply_flat_diff_redline: {e}")
        import traceback
        traceback.print_exc()
        return result


class HotkeySignals(QObject):
    """Signals for cross-thread communication."""
    show_popup = Signal()


class LLMWorkerThread(QThread):
    """Worker thread for running LLM calls without blocking the UI."""
    finished = Signal(str, str)  # (result, error_message)

    def __init__(self, llm_func, prompt: str, parent=None):
        super().__init__(parent)
        self.llm_func = llm_func
        self.prompt = prompt
        self._cancelled = False

    def run(self):
        """Execute the LLM call in a separate thread."""
        if self._cancelled:
            self.finished.emit("", "cancelled")
            return

        try:
            result = self.llm_func(self.prompt)
            if self._cancelled:
                self.finished.emit("", "cancelled")
            else:
                self.finished.emit(result or "", "")
        except Exception as e:
            if not self._cancelled:
                self.finished.emit("", str(e))

    def cancel(self):
        """Mark the thread as cancelled."""
        self._cancelled = True


# ════════════════════════════════════════════════════════════════════════════
# Parallel Task Execution — TaskLLMWorkerThread, TaskManager, TaskStatusBar
# ════════════════════════════════════════════════════════════════════════════

class TaskLLMWorkerThread(QThread):
    """Worker thread for a single parallel task's LLM call."""
    finished = Signal(str, str, str)  # (task_id, result, error)
    status_update = Signal(str, str)  # (task_id, status_message)

    def __init__(self, task_data: TaskData, parent=None):
        super().__init__(parent)
        self.task_data = task_data
        self._cancelled = False

    def _emit_status(self, message: str):
        """Emit a status update for the status bar."""
        self.status_update.emit(self.task_data.task_id, message)

    def run(self):
        if self._cancelled:
            self.finished.emit(self.task_data.task_id, "", "cancelled")
            return
        try:
            # ── Legal research (runs in worker thread to keep main thread free) ──
            if self.task_data.do_legal_research and self.task_data.legal_research_engine:
                self._run_legal_research()
                if self._cancelled:
                    self.finished.emit(self.task_data.task_id, "", "cancelled")
                    return

            if self.task_data.do_legal_research:
                self._emit_status("Research 4/4 · Generating response with citations...")
            else:
                self._emit_status("Generating response...")
            from icharlotte_core.llm import LLMHandler
            settings = {
                'temperature': 0.7,
                'top_p': 0.95,
                'max_tokens': -1,
                'stream': False,
                'thinking_level': 'None',
            }
            result = LLMHandler.generate(
                provider=self.task_data.provider,
                model=self.task_data.model_id,
                system_prompt=self.task_data.system_prompt,
                user_prompt=self.task_data.full_prompt,
                file_contents="",
                settings=settings,
            )
            if self._cancelled:
                self.finished.emit(self.task_data.task_id, "", "cancelled")
            else:
                self.finished.emit(self.task_data.task_id, (result or "").strip(), "")
        except Exception as e:
            if not self._cancelled:
                self.finished.emit(self.task_data.task_id, "", str(e))

    def _run_legal_research(self):
        """Run legal research and augment the task's full_prompt. Runs in worker thread."""
        try:
            engine = self.task_data.legal_research_engine
            provider, model_id = self.task_data.legal_research_model or (self.task_data.provider, self.task_data.model_id)

            from icharlotte_core.legal_research.prompts import (
                QUERY_EXTRACTION_PROMPT,
                CITATION_INSTRUCTION,
                RESEARCH_FRAMING_INSTRUCTION,
            )
            from icharlotte_core.llm import LLMHandler

            def _llm_for_research(system_prompt, user_prompt):
                settings = {
                    'temperature': 0.3,
                    'top_p': 0.95,
                    'max_tokens': -1,
                    'stream': False,
                    'thinking_level': 'None',
                }
                return LLMHandler.generate(
                    provider=provider, model=model_id,
                    system_prompt=system_prompt,
                    user_prompt=user_prompt,
                    file_contents="", settings=settings,
                )

            # Step 1/4: Extract legal questions
            self._emit_status("Research 1/4 · Extracting legal questions...")
            research_query = _llm_for_research(
                QUERY_EXTRACTION_PROMPT, self.task_data.full_prompt[:8000]
            )
            if not research_query or len(research_query.strip()) < 20:
                research_query = self.task_data.full_prompt[:2000]

            if self._cancelled:
                return

            # Step 2/4: Searching legal databases
            self._emit_status("Research 2/4 · Searching legal databases...")
            research_result = engine.research(
                query=research_query,
                llm_callback=_llm_for_research,
                status_callback=lambda msg: self._emit_status(f"Research 2/4 · {msg}"),
                original_prompt=self.task_data.full_prompt[:8000],
            )

            if self._cancelled:
                return

            # Step 3/4: Building authority block
            self._emit_status("Research 3/4 · Compiling citations...")
            authority_block = research_result.format_authority_block()
            self.task_data.research_result = research_result

            if authority_block and research_result:
                # Override system prompt to legal drafting persona
                from icharlotte_core.legal_research.prompts import build_augmented_system_prompt
                memo = research_result.memo or ""
                self.task_data.system_prompt = build_augmented_system_prompt(
                    self.task_data.system_prompt, authority_block, research_memo=memo
                )
                # Also inject into user prompt for models that weight user content more
                if memo:
                    self.task_data.full_prompt = (
                        f"{RESEARCH_FRAMING_INSTRUCTION}\n\n"
                        f"{self.task_data.full_prompt}\n\n"
                        f"[LEGAL RESEARCH MEMO]\n{memo}\n[/LEGAL RESEARCH MEMO]\n\n"
                        f"{authority_block}\n\n"
                        f"{CITATION_INSTRUCTION}"
                    )
                else:
                    self.task_data.full_prompt = (
                        f"{self.task_data.full_prompt}\n\n"
                        f"{authority_block}\n\n"
                        f"{CITATION_INSTRUCTION}"
                    )

            # Step 4/4 will be "Generating response..." emitted in run()

        except Exception as e:
            print(f"[TaskLLMWorkerThread] Legal research failed: {e}")
            self.task_data.research_result = None

    def cancel(self):
        self._cancelled = True


class _InsertionProxy:
    """Lightweight adapter to reuse WordLLMPopup's text insertion methods.

    Rather than duplicating ~500 lines of insertion/formatting code,
    we set the attributes the methods access and bind the methods
    from WordLLMPopup's class.
    """
    def __init__(self, task_data: TaskData):
        self._captured_format = task_data.captured_format
        self._original_document = task_data.document_com
        self.custom_format_settings = task_data.captured_format or {}
        self.redline_settings = task_data.redline_settings or {}

    # These methods will be bound after WordLLMPopup is defined (see bottom of this block)


class TaskManagerSignals(QObject):
    """Signals for TaskManager → TaskStatusBar communication."""
    task_added = Signal(str)                # task_id
    task_status_changed = Signal(str, str)  # task_id, new_status
    task_description_changed = Signal(str, str)  # task_id, description
    task_completed = Signal(str)            # task_id
    task_error = Signal(str, str)           # task_id, error_msg
    all_tasks_done = Signal()


class TaskManager(QObject):
    """Singleton managing parallel LLM tasks and sequential Word insertions.

    - Multiple LLM calls run in parallel (each in its own QThread)
    - Word COM insertions happen sequentially on the main thread
    - Bookmarks track ranges so positions auto-adjust across parallel edits
    """

    _instance = None

    @classmethod
    def instance(cls) -> 'TaskManager':
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        super().__init__()
        self.signals = TaskManagerSignals()
        self._tasks: Dict[str, TaskData] = {}
        self._workers: Dict[str, TaskLLMWorkerThread] = {}
        self._insertion_queue: list = []
        self._inserting = False

    def submit_task(self, task_data: TaskData) -> str:
        """Submit a new task for parallel execution. Returns task_id."""
        task_data.status = "running"
        self._tasks[task_data.task_id] = task_data

        worker = TaskLLMWorkerThread(task_data)
        worker.finished.connect(self._on_llm_finished)
        worker.status_update.connect(self._on_worker_status_update)
        self._workers[task_data.task_id] = worker

        self.signals.task_added.emit(task_data.task_id)
        self.signals.task_status_changed.emit(task_data.task_id, "running")

        worker.start()
        print(f"[TaskManager] Task {task_data.task_id} submitted: '{task_data.prompt_preview}'")
        return task_data.task_id

    def _on_worker_status_update(self, task_id: str, message: str):
        """Forward worker status updates to the status bar."""
        self.signals.task_description_changed.emit(task_id, message)

    def cancel_task(self, task_id: str):
        """Cancel a running task."""
        if task_id in self._workers:
            self._workers[task_id].cancel()
        if task_id in self._tasks:
            td = self._tasks[task_id]
            td.status = "cancelled"
            if td.bookmark_name and td.document_com:
                _delete_task_bookmark(td.document_com, td.bookmark_name)
            self.signals.task_status_changed.emit(task_id, "cancelled")
            self._check_all_done()

    def get_task(self, task_id: str) -> Optional[TaskData]:
        return self._tasks.get(task_id)

    def get_active_tasks(self) -> list:
        return [td for td in self._tasks.values()
                if td.status in ("pending", "running", "inserting")]

    def _on_llm_finished(self, task_id: str, result: str, error: str):
        """Handle LLM completion — queue for insertion (runs on main thread)."""
        if task_id not in self._tasks:
            return

        td = self._tasks[task_id]
        self._workers.pop(task_id, None)

        if error:
            td.status = "error"
            td.error_msg = error
            if error != "cancelled":
                self.signals.task_error.emit(task_id, error)
            if td.bookmark_name and td.document_com:
                _delete_task_bookmark(td.document_com, td.bookmark_name)
            self._check_all_done()
            return

        if not result:
            td.status = "error"
            td.error_msg = "LLM returned empty response"
            self.signals.task_error.emit(task_id, td.error_msg)
            if td.bookmark_name and td.document_com:
                _delete_task_bookmark(td.document_com, td.bookmark_name)
            self._check_all_done()
            return

        # Apply deterministic citation cross-check on the final LLM output
        # if legal research was performed (non-LLM safety net)
        if td.research_result and td.research_result.cases:
            try:
                from icharlotte_core.legal_research.engine import LegalResearchEngine
                known_names = td.research_result.get_known_case_names()
                result = LegalResearchEngine._deterministic_citation_check(
                    result, known_names
                )
            except Exception as e:
                print(f"[TaskManager] Deterministic citation check failed: {e}")

        td.result_text = result
        td.status = "inserting"
        self._insertion_queue.append(task_id)
        self.signals.task_status_changed.emit(task_id, "inserting")

        # Defer insertion so the event loop can process UI events first
        # (prevents freezing the second popup if user opened one)
        QTimer.singleShot(0, self._process_insertion_queue)

    def _process_insertion_queue(self):
        """Process the next task in the insertion queue (sequential, main thread)."""
        if self._inserting or not self._insertion_queue:
            return

        self._inserting = True
        task_id = self._insertion_queue.pop(0)
        td = self._tasks.get(task_id)
        if not td:
            self._inserting = False
            return

        # Process pending events (e.g. show_popup signal) before blocking on COM
        QApplication.processEvents()

        try:
            self._insert_result(td)
            td.status = "done"
            self.signals.task_completed.emit(task_id)
        except Exception as e:
            print(f"[TaskManager] Insertion failed for {task_id}: {e}")
            import traceback
            traceback.print_exc()
            td.status = "error"
            td.error_msg = str(e)
            # Fallback: copy to clipboard
            try:
                QApplication.clipboard().setText(td.result_text)
                print(f"[TaskManager] Result copied to clipboard for task {task_id}")
            except Exception:
                pass
            self.signals.task_error.emit(task_id, f"Insertion failed: {e}\nResult copied to clipboard.")
        finally:
            # Clean up bookmark
            if td and td.bookmark_name and td.document_com:
                _delete_task_bookmark(td.document_com, td.bookmark_name)
            self._inserting = False
            self._check_all_done()
            # Process next in queue
            if self._insertion_queue:
                QTimer.singleShot(100, self._process_insertion_queue)

    def _insert_result(self, td: TaskData):
        """Insert LLM result into Word at the bookmark location."""
        # Verify document is still open and get its Application object.
        # IMPORTANT: We derive `word` from `doc.Application` rather than
        # GetActiveObject so that word and doc are always in the same
        # Word process.  When multiple Word instances are running (common
        # on Windows when documents are opened from different sources),
        # GetActiveObject returns whichever instance is foreground, which
        # may be a DIFFERENT process than the one owning the target doc.
        # That mismatch causes word.Selection to point at the wrong
        # document, so the insertion silently goes nowhere.
        doc = td.document_com
        word = None

        if doc:
            try:
                _ = doc.Name  # test if COM ref is alive
                word = doc.Application
                print(f"[TaskManager] Using doc.Application for '{doc.Name}'")
            except Exception as e:
                print(f"[TaskManager] Stored doc COM ref stale: {e}")
                doc = None

        # Fallback: find document by name via GetActiveObject
        if not doc and td.document_name:
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                try:
                    word = win32com.client.GetObject(Class="Word.Application")
                except Exception:
                    raise Exception("Could not connect to Word")
            if word:
                try:
                    for i in range(1, word.Documents.Count + 1):
                        candidate = word.Documents(i)
                        if candidate.Name == td.document_name:
                            doc = candidate
                            word = doc.Application  # ensure same process
                            print(f"[TaskManager] Found doc by name: '{td.document_name}'")
                            break
                except Exception:
                    pass

        if not doc:
            raise Exception(f"Document '{td.document_name or 'unknown'}' is no longer open")
        if not word:
            raise Exception("Could not connect to Word")

        # Activate document
        doc.Activate()

        # Resolve bookmark to current range
        if td.bookmark_name:
            bm_range = _get_bookmark_range(doc, td.bookmark_name)
            if bm_range is None:
                # Bookmark lost — fall back to current selection
                print(f"[TaskManager] Bookmark lost, falling back to current selection")
                selection = word.Selection
                range_start = selection.Range.Start
                range_end = selection.Range.End
            else:
                range_start, range_end = bm_range
        else:
            # No bookmark (e.g., no selection was made) — insert at current cursor
            print(f"[TaskManager] No bookmark — inserting at current cursor position")
            selection = word.Selection
            range_start = selection.Range.Start
            range_end = selection.Range.Start  # Insertion point

        # Lock Word UI during insertion
        ui_locked = False
        try:
            word.ScreenUpdating = False
            word.Interactive = False
            ui_locked = True
        except Exception:
            pass

        try:
            # Clamp to document bounds to avoid "Value out of range"
            doc_end = doc.Content.End
            range_start = min(range_start, doc_end - 1)
            range_end = min(range_end, doc_end - 1)
            if range_start < 0:
                range_start = 0
            if range_end < range_start:
                range_end = range_start

            # Select the bookmark range
            if td.has_selection:
                target_range = doc.Range(range_start, range_end)
            else:
                target_range = doc.Range(range_start, range_start)
            target_range.Select()
            selection = word.Selection

            if td.redline_mode_active:
                r = apply_flat_diff_redline(
                    doc_com=doc,
                    range_start=range_start,
                    range_end=range_end,
                    original_text=td.original_text_raw or td.original_text,
                    revised_text=td.result_text,
                    auto_enable_track_changes=(
                        td.redline_settings.get("auto_enable_track_changes", True)
                        if td.redline_settings else True
                    ),
                )
                if not r['success']:
                    # Fallback to plain replacement
                    self._do_plain_insert(word, selection, td)
            else:
                self._do_plain_insert(word, selection, td)

            print(f"[TaskManager] Successfully inserted result for task {td.task_id}")
        finally:
            if ui_locked:
                try:
                    word.Interactive = True
                    word.ScreenUpdating = True
                except Exception:
                    try:
                        fresh = win32com.client.GetActiveObject("Word.Application")
                        fresh.Interactive = True
                        fresh.ScreenUpdating = True
                    except Exception:
                        print("WARNING: Could not unlock Word UI")

            # Bring Word to the foreground showing the correct document.
            # After doc.Activate(), the document is active within Word but
            # the window may be behind other apps if the user switched away.
            try:
                if doc:
                    doc.Activate()
                if word and HAS_WIN32:
                    import win32con
                    hwnd = getattr(word, 'hwnd', 0) or 0
                    if hwnd:
                        if win32gui.IsIconic(hwnd):
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.SetForegroundWindow(hwnd)
                        print(f"[TaskManager] Brought Word to foreground (hwnd={hwnd})")
            except Exception as e:
                print(f"[TaskManager] Could not bring Word to foreground: {e}")


    def _do_plain_insert(self, word, selection, td: TaskData):
        """Insert text using the appropriate format mode via _InsertionProxy."""
        proxy = _InsertionProxy(td)
        proxy._set_word_text_internal(
            td.result_text, td.format_type,
            word=word, selection=selection
        )

    def _check_all_done(self):
        """Emit all_tasks_done if no active tasks remain."""
        if not self.get_active_tasks():
            self.signals.all_tasks_done.emit()


class TaskStatusBar(QWidget):
    """Floating always-on-top widget showing active AI tasks.

    Appears in bottom-right of screen. Each task shows:
    - Spinner/status icon
    - Document name + prompt preview
    - Cancel button

    Auto-hides 3 seconds after all tasks complete.
    """

    _instance = None

    @classmethod
    def instance(cls) -> 'TaskStatusBar':
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        super().__init__(None)
        self.setWindowFlags(
            Qt.WindowType.Tool |
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self._task_rows: Dict[str, QWidget] = {}
        self._spinner_timers: Dict[str, QTimer] = {}
        self._auto_hide_timer = QTimer(self)
        self._auto_hide_timer.setSingleShot(True)
        self._auto_hide_timer.timeout.connect(self.hide)

        self._setup_ui()
        self._connect_signals()

    def _setup_ui(self):
        self.setMinimumWidth(320)
        self.setMaximumWidth(420)

        self._container = QFrame(self)
        self._container.setStyleSheet("""
            QFrame#taskStatusContainer {
                background-color: #1e1e2e;
                border: 1px solid #6c63ff;
                border-radius: 8px;
            }
        """)
        self._container.setObjectName("taskStatusContainer")

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.addWidget(self._container)

        self._layout = QVBoxLayout(self._container)
        self._layout.setContentsMargins(10, 8, 10, 8)
        self._layout.setSpacing(4)

        header = QLabel("AI Tasks")
        header.setStyleSheet("color: #6c63ff; font-weight: bold; font-size: 11px; background: transparent;")
        self._layout.addWidget(header)

    def _connect_signals(self):
        tm = TaskManager.instance()
        tm.signals.task_added.connect(self._on_task_added)
        tm.signals.task_status_changed.connect(self._on_task_status_changed)
        tm.signals.task_description_changed.connect(self._on_task_description_changed)
        tm.signals.task_completed.connect(self._on_task_completed)
        tm.signals.task_error.connect(self._on_task_error)
        tm.signals.all_tasks_done.connect(self._on_all_done)

    def show_preparing(self, prep_id: str, prompt_preview: str):
        """Show a 'preparing' row immediately before the task is submitted.

        Called right after the popup hides so the user sees instant feedback.
        The row is replaced when the real task is submitted via _on_task_added.
        """
        self._auto_hide_timer.stop()

        row = QFrame()
        row.setStyleSheet("QFrame { background: transparent; }")
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 2, 0, 2)
        row_layout.setSpacing(6)

        spinner = QLabel("⠋")
        spinner.setStyleSheet("color: #f9e2af; font-size: 12px; background: transparent;")
        spinner.setFixedWidth(16)
        spinner.setObjectName("spinner")
        row_layout.addWidget(spinner)

        spin_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        spin_idx = [0]
        timer = QTimer()
        timer.setInterval(100)
        timer.timeout.connect(lambda: (
            spinner.setText(spin_chars[spin_idx[0] % len(spin_chars)]),
            spin_idx.__setitem__(0, spin_idx[0] + 1)
        ))
        timer.start()
        self._spinner_timers[prep_id] = timer

        prompt_short = prompt_preview[:35] + "..." if len(prompt_preview) > 35 else prompt_preview
        desc_label = QLabel(f"Preparing · {prompt_short}")
        desc_label.setStyleSheet("color: #cdd6f4; font-size: 11px; background: transparent;")
        row_layout.addWidget(desc_label, stretch=1)

        self._task_rows[prep_id] = row
        self._layout.addWidget(row)
        self._reposition()
        self.show()
        self.raise_()

    def upgrade_preparing(self, prep_id: str, task_id: str):
        """Replace a 'preparing' row with the real task_id after submission."""
        if prep_id in self._task_rows:
            row = self._task_rows.pop(prep_id)
            timer = self._spinner_timers.pop(prep_id, None)
            if timer:
                self._spinner_timers[task_id] = timer
            self._task_rows[task_id] = row

            # Update description from TaskManager
            tm = TaskManager.instance()
            td = tm.get_task(task_id)
            if td:
                for child in row.findChildren(QLabel):
                    if child.objectName() != "spinner":
                        doc_short = td.document_name[:18] if td.document_name else "Word"
                        prompt_short = td.prompt_preview[:35] + "..." if len(td.prompt_preview) > 35 else td.prompt_preview
                        child.setText(f"{doc_short} · {prompt_short}")
                        child.setToolTip(f"Document: {td.document_name}\nPrompt: {td.prompt_preview}")
                        break

            # Add cancel button if not present
            cancel_btn = row.findChild(QPushButton)
            if not cancel_btn:
                cancel_btn = QPushButton("✕")
                cancel_btn.setFixedSize(18, 18)
                cancel_btn.setStyleSheet("""
                    QPushButton {
                        background: #45475a; color: #cdd6f4;
                        border-radius: 9px; font-size: 9px;
                        border: none; padding: 0px;
                    }
                    QPushButton:hover { background: #f38ba8; color: white; }
                """)
                cancel_btn.setToolTip("Cancel task")
                cancel_btn.clicked.connect(lambda checked=False, tid=task_id: TaskManager.instance().cancel_task(tid))
                row.layout().addWidget(cancel_btn)

    def _on_task_added(self, task_id: str):
        self._auto_hide_timer.stop()

        # If there's already a preparing row for this task, skip creating a new one
        if task_id in self._task_rows:
            return

        tm = TaskManager.instance()
        td = tm.get_task(task_id)
        if not td:
            return

        row = QFrame()
        row.setStyleSheet("QFrame { background: transparent; }")
        row_layout = QHBoxLayout(row)
        row_layout.setContentsMargins(0, 2, 0, 2)
        row_layout.setSpacing(6)

        # Animated spinner
        spinner = QLabel("⠋")
        spinner.setStyleSheet("color: #f9e2af; font-size: 12px; background: transparent;")
        spinner.setFixedWidth(16)
        spinner.setObjectName("spinner")
        row_layout.addWidget(spinner)

        # Start spinner animation
        spin_chars = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]
        spin_idx = [0]
        timer = QTimer()
        timer.setInterval(100)
        timer.timeout.connect(lambda: (
            spinner.setText(spin_chars[spin_idx[0] % len(spin_chars)]),
            spin_idx.__setitem__(0, spin_idx[0] + 1)
        ))
        timer.start()
        self._spinner_timers[task_id] = timer

        # Task description
        doc_short = td.document_name[:18] if td.document_name else "Word"
        prompt_short = td.prompt_preview[:35] + "..." if len(td.prompt_preview) > 35 else td.prompt_preview
        desc_label = QLabel(f"{doc_short} · {prompt_short}")
        desc_label.setStyleSheet("color: #cdd6f4; font-size: 11px; background: transparent;")
        desc_label.setToolTip(f"Document: {td.document_name}\nPrompt: {td.prompt_preview}")
        row_layout.addWidget(desc_label, stretch=1)

        # Cancel button
        cancel_btn = QPushButton("✕")
        cancel_btn.setFixedSize(18, 18)
        cancel_btn.setStyleSheet("""
            QPushButton {
                background: #45475a; color: #cdd6f4;
                border-radius: 9px; font-size: 9px;
                border: none; padding: 0px;
            }
            QPushButton:hover { background: #f38ba8; color: white; }
        """)
        cancel_btn.setToolTip("Cancel task")
        cancel_btn.clicked.connect(lambda checked=False, tid=task_id: TaskManager.instance().cancel_task(tid))
        row_layout.addWidget(cancel_btn)

        self._task_rows[task_id] = row
        self._layout.addWidget(row)

        self._reposition()
        self.show()
        self.raise_()

    def _on_task_status_changed(self, task_id: str, status: str):
        row = self._task_rows.get(task_id)
        if not row:
            return
        spinner = row.findChild(QLabel, "spinner")
        if not spinner:
            return

        if status == "inserting":
            # Stop animation, show static icon
            timer = self._spinner_timers.pop(task_id, None)
            if timer:
                timer.stop()
            spinner.setText("▸")
            spinner.setStyleSheet("color: #89b4fa; font-size: 12px; background: transparent;")
            # Update description to show inserting state
            for child in row.findChildren(QLabel):
                if child is not spinner and child.objectName() != "spinner":
                    child.setText("Inserting into document...")
                    break
        elif status == "cancelled":
            timer = self._spinner_timers.pop(task_id, None)
            if timer:
                timer.stop()
            spinner.setText("—")
            spinner.setStyleSheet("color: #6c7086; font-size: 12px; background: transparent;")
            QTimer.singleShot(2000, lambda: self._remove_task_row(task_id))

    def _on_task_description_changed(self, task_id: str, description: str):
        """Update the description label for a running task."""
        row = self._task_rows.get(task_id)
        if not row:
            return
        spinner = row.findChild(QLabel, "spinner")
        for child in row.findChildren(QLabel):
            if child is not spinner and child.objectName() != "spinner":
                child.setText(description)
                child.setToolTip(description)
                break

    def _on_task_completed(self, task_id: str):
        timer = self._spinner_timers.pop(task_id, None)
        if timer:
            timer.stop()

        row = self._task_rows.get(task_id)
        if not row:
            return
        spinner = row.findChild(QLabel, "spinner")
        if spinner:
            spinner.setText("✓")
            spinner.setStyleSheet("color: #a6e3a1; font-size: 12px; font-weight: bold; background: transparent;")
        for child in row.findChildren(QLabel):
            if child is not spinner and child.objectName() != "spinner":
                child.setText("Done — inserted into document")
                child.setStyleSheet("color: #a6e3a1; font-size: 11px; background: transparent;")
                break
        QTimer.singleShot(2000, lambda: self._remove_task_row(task_id))

    def _on_task_error(self, task_id: str, error: str):
        timer = self._spinner_timers.pop(task_id, None)
        if timer:
            timer.stop()

        row = self._task_rows.get(task_id)
        if not row:
            return
        spinner = row.findChild(QLabel, "spinner")
        if spinner:
            spinner.setText("✗")
            spinner.setStyleSheet("color: #f38ba8; font-size: 12px; font-weight: bold; background: transparent;")
            spinner.setToolTip(error)
        # Update description label to show error
        for child in row.findChildren(QLabel):
            if child is not spinner and child.objectName() != "spinner":
                short_err = error[:60] + "..." if len(error) > 60 else error
                child.setText(f"Error: {short_err}")
                child.setStyleSheet("color: #f38ba8; font-size: 11px; background: transparent;")
                break
        QTimer.singleShot(8000, lambda: self._remove_task_row(task_id))

    def _remove_task_row(self, task_id: str):
        timer = self._spinner_timers.pop(task_id, None)
        if timer:
            timer.stop()
        row = self._task_rows.pop(task_id, None)
        if row:
            self._layout.removeWidget(row)
            row.deleteLater()
            self._reposition()

    def _on_all_done(self):
        self._auto_hide_timer.start(3000)

    def _reposition(self):
        screen = QApplication.primaryScreen()
        if screen:
            geo = screen.availableGeometry()
            self.adjustSize()
            x = geo.right() - self.width() - 20
            y = geo.bottom() - self.height() - 20
            self.move(x, y)


class ReviewWorkerThread(QThread):
    """Worker thread for running the report review without blocking the UI."""
    progress = Signal(str)  # progress message
    finished = Signal(object, int, str)  # (structural_result, total_changes, error)

    def __init__(self, reviewer, selection_range=None, include_case_data=False,
                 extra_doc_paths=None, doc_name=None, parent=None):
        super().__init__(parent)
        self.reviewer = reviewer
        self.selection_range = selection_range
        self.include_case_data = include_case_data
        self.extra_doc_paths = extra_doc_paths
        self.doc_name = doc_name  # Document name to re-acquire in worker thread
        self._cancelled = False

    def run(self):
        """Execute the review in a separate thread."""
        import pythoncom
        pythoncom.CoInitialize()
        try:
            if self._cancelled:
                self.finished.emit(None, 0, "cancelled")
                return

            # Re-acquire Word COM object in this thread (COM objects can't cross threads)
            # Use gencache (early-binding) to avoid _LazyAddAttr_ fatal crashes (0x8001010e)
            try:
                word = win32com.client.gencache.EnsureDispatch("Word.Application")
            except Exception:
                word = win32com.client.Dispatch("Word.Application")
            if self.doc_name:
                # Find the document by name
                doc = None
                for i in range(1, word.Documents.Count + 1):
                    if word.Documents(i).Name == self.doc_name:
                        doc = word.Documents(i)
                        break
                if not doc:
                    doc = word.ActiveDocument
            else:
                doc = word.ActiveDocument

            if not doc:
                self.finished.emit(None, 0, "Could not re-acquire Word document in worker thread")
                return

            self.reviewer.doc = doc

            structural, total_changes = self.reviewer.run(
                selection_range=self.selection_range,
                include_case_data=self.include_case_data,
                extra_doc_paths=self.extra_doc_paths,
                progress_callback=self._on_progress,
            )
            if self._cancelled:
                self.finished.emit(None, 0, "cancelled")
            else:
                self.finished.emit(structural, total_changes, "")
        except Exception as e:
            if not self._cancelled:
                import traceback
                tb = traceback.format_exc()
                print(f"[ReviewWorker ERROR] {tb}")
                self.finished.emit(None, 0, str(e))
        finally:
            pythoncom.CoUninitialize()

    def _on_progress(self, msg):
        if not self._cancelled:
            self.progress.emit(msg)

    def cancel(self):
        self._cancelled = True


# Accepted file extensions for attachments
ATTACHMENT_EXTENSIONS = (".doc", ".docx", ".pdf", ".msg")


class FileExtractorThread(QThread):
    """Worker thread for extracting text from attached files."""
    extracted = Signal(str, str, int)  # (file_path, text, char_count)
    error = Signal(str, str)  # (file_path, error_message)

    def __init__(self, file_path: str, parent=None):
        super().__init__(parent)
        self.file_path = file_path

    def run(self):
        try:
            ext = os.path.splitext(self.file_path)[1].lower()

            if ext == ".doc" and HAS_WIN32:
                # Legacy .doc format — use Word COM
                pythoncom.CoInitialize()
                try:
                    word = win32com.client.Dispatch("Word.Application")
                    doc = word.Documents.Open(self.file_path, ReadOnly=True, Visible=False)
                    text = doc.Content.Text or ""
                    doc.Close(False)
                    self.extracted.emit(self.file_path, text, len(text))
                finally:
                    pythoncom.CoUninitialize()
            elif HAS_DOC_PROCESSOR:
                processor = DocumentProcessor()
                result = processor.extract_text(self.file_path)
                if result.success:
                    self.extracted.emit(self.file_path, result.text, len(result.text))
                else:
                    self.error.emit(self.file_path, result.error or "Extraction failed")
            else:
                self.error.emit(self.file_path, "DocumentProcessor not available")
        except Exception as e:
            self.error.emit(self.file_path, str(e))


class AttachmentArea(QFrame):
    """Drop zone and file chip area for attaching context files."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._attachments = {}  # file_path -> extracted_text
        self._chips = {}  # file_path -> chip QFrame
        self._extractors = []  # keep refs to running threads
        self._setup_ui()
        self.setAcceptDrops(True)

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 4, 0, 0)
        layout.setSpacing(4)

        # Drop zone / browse row
        drop_row = QHBoxLayout()
        drop_row.setSpacing(6)

        self._drop_label = QLabel("Drop files here or click to attach")
        self._drop_label.setStyleSheet(
            "color: #a6adc8; font-size: 10px; font-style: italic; padding: 4px;"
        )
        drop_row.addWidget(self._drop_label, 1)

        browse_btn = QPushButton("\U0001f4ce")  # 📎
        browse_btn.setFixedSize(28, 28)
        browse_btn.setToolTip("Attach files (.doc, .docx, .pdf, .msg)")
        browse_btn.setStyleSheet(
            "QPushButton { background: #313244; border: 1px solid #45475a; "
            "border-radius: 4px; font-size: 14px; }"
            "QPushButton:hover { background: #45475a; border-color: #6c63ff; }"
        )
        browse_btn.clicked.connect(self._browse_files)
        drop_row.addWidget(browse_btn)

        layout.addLayout(drop_row)

        # Chips container
        self._chips_widget = QWidget()
        self._chips_layout = QHBoxLayout(self._chips_widget)
        self._chips_layout.setContentsMargins(0, 0, 0, 0)
        self._chips_layout.setSpacing(4)
        self._chips_layout.addStretch()
        self._chips_widget.setVisible(False)
        layout.addWidget(self._chips_widget)

        # Summary label
        self._summary_label = QLabel("")
        self._summary_label.setStyleSheet("color: #a6adc8; font-size: 10px; font-style: italic;")
        self._summary_label.setVisible(False)
        layout.addWidget(self._summary_label)

        # Style the frame itself with dashed border
        self.setStyleSheet(
            "AttachmentArea { border: 1px dashed #45475a; border-radius: 4px; padding: 2px; }"
        )

    def _browse_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self, "Attach Files",
            "",
            "Documents (*.doc *.docx *.pdf *.msg);;All Files (*)"
        )
        if paths:
            self.add_files(paths)

    def add_files(self, paths):
        for path in paths:
            if path in self._attachments or path in self._chips:
                continue  # already added
            ext = os.path.splitext(path)[1].lower()
            if ext not in ATTACHMENT_EXTENSIONS:
                continue
            self._add_chip(path)
            self._start_extraction(path)

    def _add_chip(self, file_path):
        chip = QFrame()
        chip.setStyleSheet(
            "QFrame { background: #313244; border: 1px solid #45475a; "
            "border-radius: 4px; padding: 2px 6px; }"
        )
        chip_layout = QHBoxLayout(chip)
        chip_layout.setContentsMargins(4, 2, 4, 2)
        chip_layout.setSpacing(4)

        name = os.path.basename(file_path)
        name_label = QLabel(name if len(name) <= 30 else name[:27] + "...")
        name_label.setToolTip(file_path)
        name_label.setStyleSheet("color: #cdd6f4; font-size: 10px; border: none;")
        chip_layout.addWidget(name_label)

        status_label = QLabel("...")
        status_label.setStyleSheet("color: #a6adc8; font-size: 10px; border: none;")
        status_label.setObjectName("status")
        chip_layout.addWidget(status_label)

        remove_btn = QPushButton("\u00d7")  # ×
        remove_btn.setFixedSize(16, 16)
        remove_btn.setStyleSheet(
            "QPushButton { background: transparent; color: #a6adc8; border: none; "
            "font-size: 12px; font-weight: bold; }"
            "QPushButton:hover { color: #f38ba8; }"
        )
        remove_btn.clicked.connect(lambda checked, p=file_path: self._remove_file(p))
        chip_layout.addWidget(remove_btn)

        self._chips[file_path] = chip
        # Insert before the stretch
        self._chips_layout.insertWidget(self._chips_layout.count() - 1, chip)
        self._chips_widget.setVisible(True)

    def _start_extraction(self, file_path):
        thread = FileExtractorThread(file_path, self)
        thread.extracted.connect(self._on_extracted)
        thread.error.connect(self._on_extract_error)
        thread.finished.connect(lambda t=thread: self._extractors.remove(t) if t in self._extractors else None)
        self._extractors.append(thread)
        thread.start()

    def _on_extracted(self, file_path, text, char_count):
        self._attachments[file_path] = text
        chip = self._chips.get(file_path)
        if chip:
            status = chip.findChild(QLabel, "status")
            if status:
                if char_count > 1000:
                    status.setText(f"{char_count // 1000}k chars")
                else:
                    status.setText(f"{char_count} chars")
                status.setStyleSheet("color: #a6e3a1; font-size: 10px; border: none;")
        self._update_summary()

    def _on_extract_error(self, file_path, error_msg):
        self._attachments[file_path] = ""
        chip = self._chips.get(file_path)
        if chip:
            status = chip.findChild(QLabel, "status")
            if status:
                status.setText("error")
                status.setToolTip(error_msg)
                status.setStyleSheet("color: #f38ba8; font-size: 10px; border: none;")
        self._update_summary()

    def _remove_file(self, file_path):
        self._attachments.pop(file_path, None)
        chip = self._chips.pop(file_path, None)
        if chip:
            self._chips_layout.removeWidget(chip)
            chip.deleteLater()
        if not self._chips:
            self._chips_widget.setVisible(False)
        self._update_summary()

    def _update_summary(self):
        files_with_text = {p: t for p, t in self._attachments.items() if t}
        if not files_with_text:
            self._summary_label.setVisible(False)
            return
        total_chars = sum(len(t) for t in files_with_text.values())
        n = len(files_with_text)
        if total_chars > 1000:
            self._summary_label.setText(f"{n} file{'s' if n != 1 else ''}, ~{total_chars // 1000}k chars")
        else:
            self._summary_label.setText(f"{n} file{'s' if n != 1 else ''}, {total_chars} chars")
        self._summary_label.setVisible(True)

    def get_attachment_context(self):
        """Return formatted context blocks for all extracted attachments."""
        blocks = []
        for file_path, text in self._attachments.items():
            if text:
                name = os.path.basename(file_path)
                blocks.append(f"\n\n=== ATTACHED FILE: {name} ===\n{text}")
        return "".join(blocks)

    def clear(self):
        for path in list(self._chips.keys()):
            self._remove_file(path)

    # Drag and drop support
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if path and os.path.splitext(path)[1].lower() in ATTACHMENT_EXTENSIONS:
                    event.acceptProposedAction()
                    self.setStyleSheet(
                        "AttachmentArea { border: 2px dashed #6c63ff; border-radius: 4px; "
                        "padding: 2px; background: rgba(108, 99, 255, 0.05); }"
                    )
                    return
        event.ignore()

    def dragLeaveEvent(self, event):
        self.setStyleSheet(
            "AttachmentArea { border: 1px dashed #45475a; border-radius: 4px; padding: 2px; }"
        )

    def dropEvent(self, event):
        self.setStyleSheet(
            "AttachmentArea { border: 1px dashed #45475a; border-radius: 4px; padding: 2px; }"
        )
        paths = []
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path and os.path.isfile(path) and os.path.splitext(path)[1].lower() in ATTACHMENT_EXTENSIONS:
                paths.append(path)
        if paths:
            event.acceptProposedAction()
            self.add_files(paths)
        else:
            event.ignore()


class WordLLMPopup(QDialog):
    """Popup dialog for LLM processing of Word content."""

    def __init__(self, parent=None, llm_callback: Optional[Callable] = None, cursor_pos=None):
        super().__init__(parent)
        self.llm_callback = llm_callback
        self.cursor_pos = cursor_pos  # Position to show popup at (QPoint or None)
        self.prompts_path = os.path.join(GEMINI_DATA_DIR, "word_llm_prompts.json")
        self.outlook_prompts_path = os.path.join(GEMINI_DATA_DIR, "outlook_llm_prompts.json")
        self.format_settings_path = os.path.join(GEMINI_DATA_DIR, "word_format_settings.json")
        self.model_settings_path = os.path.join(GEMINI_DATA_DIR, "word_model_settings.json")
        self.prompts = []  # Word prompts
        self.outlook_prompts = []  # Outlook/email prompts
        self.custom_format_settings = {}  # Custom format settings
        self.selected_model_index = DEFAULT_MODEL_INDEX  # Selected model index
        self._word_app = None  # Stored Word COM reference during execution
        self._captured_format = None  # Captured format from selection

        # Load redline settings
        self.redline_settings = load_redline_settings(GEMINI_DATA_DIR)

        # App context for Word vs Outlook
        self.app_context = APP_CONTEXT_WORD  # Default context
        self.active_inspector = None  # Outlook Inspector reference when in email mode

        # Worker thread for LLM calls
        self._worker_thread = None
        self._review_thread = None  # Worker thread for report review
        self._pending_format_type = None  # Format type for pending task
        self._is_outlook_task = False  # Whether pending task is for Outlook

        # Case detection for variable insertion
        self._detected_case = None  # Case dict from database
        self._detected_file_number = None  # File number extracted from path
        self._case_detection_error = None  # Error message if detection failed

        # Drag/move support for independent positioning
        self._drag_pos = None  # Starting position for drag
        self._is_dragging = False

        # Redline state
        self._original_text = None  # Original text before LLM processing
        self._original_document = None  # Document object where selection was made
        self._original_document_name = None  # Document name for fallback lookup
        self._original_range_start = None  # Range start position
        self._original_range_end = None  # Range end position
        self._original_has_selection = False  # Whether original capture had a selection
        self._redline_mode_active = False  # Whether current operation uses redline
        self._word_ui_locked = False  # Whether Word UI is locked (Interactive=False)

        self.setWindowTitle("AI Assistant")
        self.setWindowFlags(
            Qt.WindowType.Dialog |
            Qt.WindowType.FramelessWindowHint
        )
        self.setMinimumWidth(450)
        self.setStyleSheet("""
            QDialog {
                background-color: #1e1e2e;
                border: 2px solid #6c63ff;
                border-radius: 10px;
            }
            QLabel {
                color: #cdd6f4;
                font-size: 12px;
            }
            QComboBox {
                background-color: #313244;
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 5px;
                padding: 8px;
                font-size: 12px;
            }
            QComboBox:hover {
                border-color: #6c63ff;
            }
            QComboBox::drop-down {
                border: none;
                width: 30px;
            }
            QComboBox QAbstractItemView {
                background-color: #313244;
                color: #cdd6f4;
                selection-background-color: #6c63ff;
            }
            QLineEdit, QTextEdit {
                background-color: #313244;
                color: #cdd6f4;
                border: 1px solid #45475a;
                border-radius: 5px;
                padding: 8px;
                font-size: 12px;
            }
            QLineEdit:focus, QTextEdit:focus {
                border-color: #6c63ff;
            }
            QPushButton {
                background-color: #6c63ff;
                color: white;
                border: none;
                border-radius: 5px;
                padding: 8px 16px;
                font-size: 12px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #7c73ff;
            }
            QPushButton:pressed {
                background-color: #5c53ef;
            }
            QPushButton#cancelBtn {
                background-color: #45475a;
            }
            QPushButton#cancelBtn:hover {
                background-color: #585b70;
            }
            QPushButton#deleteBtn {
                background-color: #f38ba8;
            }
            QPushButton#deleteBtn:hover {
                background-color: #f5a0b8;
            }
            QPushButton#reviewBtn {
                background-color: #89b4fa;
                color: #1e1e2e;
                font-weight: bold;
            }
            QPushButton#reviewBtn:hover {
                background-color: #a6c8ff;
            }
        """)

        self.setup_ui()
        self.load_prompts()
        self._load_format_settings()
        self._load_model_settings()
        self._update_format_preview()
        self._update_redline_checkbox_visibility()  # Set initial visibility based on context

        # Close on Escape
        self.shortcut_escape = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        self.shortcut_escape.activated.connect(self.close)

        # Execute on Enter (when in custom prompt field)
        self.shortcut_enter = QShortcut(QKeySequence(Qt.Key.Key_Return), self)
        self.shortcut_enter.activated.connect(self.execute)

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Header row with drag handle and title
        header_widget = QWidget()
        header_widget.setCursor(Qt.CursorShape.SizeAllCursor)
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 0)
        header_layout.setSpacing(8)

        # Drag handle indicator (grip dots)
        drag_handle = QLabel("\u2630")  # Unicode hamburger menu / grip icon
        drag_handle.setStyleSheet("font-size: 14px; color: #6c7086;")
        drag_handle.setToolTip("Drag to move window")
        header_layout.addWidget(drag_handle)

        # Title (dynamic based on context)
        self.title_label = QLabel("AI Assistant for Word")
        self.title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #6c63ff;")
        self.title_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(self.title_label, 1)  # Stretch to fill

        # Spacer to balance the drag handle
        spacer_label = QLabel("")
        spacer_label.setFixedWidth(20)
        header_layout.addWidget(spacer_label)

        self.header_widget = header_widget
        layout.addWidget(header_widget)

        # Case info header (shows detected case)
        self.case_info_label = QLabel("")
        self.case_info_label.setStyleSheet("color: #a6e3a1; font-size: 11px; padding: 4px;")
        self.case_info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.case_info_label.setWordWrap(True)
        layout.addWidget(self.case_info_label)

        # Tab widget for AI Prompt vs Case Variables
        self.tab_widget = QTabWidget()
        self.tab_widget.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #45475a;
                border-radius: 5px;
                background-color: #1e1e2e;
            }
            QTabBar::tab {
                background-color: #313244;
                color: #cdd6f4;
                padding: 8px 16px;
                border: 1px solid #45475a;
                border-bottom: none;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #6c63ff;
                color: white;
            }
            QTabBar::tab:hover:!selected {
                background-color: #45475a;
            }
        """)
        layout.addWidget(self.tab_widget)

        # Create AI Prompt tab
        self._setup_ai_prompt_tab()

        # Create Case Variables tab
        self._setup_case_variables_tab()

        # Status label (shared across tabs)
        self.status_label = QLabel("")
        self.status_label.setStyleSheet("color: #a6adc8; font-style: italic;")
        layout.addWidget(self.status_label)

        # Buttons (shared across tabs)
        btn_row = QHBoxLayout()
        btn_row.addStretch()

        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.setObjectName("cancelBtn")
        self.cancel_btn.clicked.connect(self._on_cancel_clicked)
        btn_row.addWidget(self.cancel_btn)

        self.execute_btn = QPushButton("Execute (Enter)")
        self.execute_btn.clicked.connect(self.execute)
        btn_row.addWidget(self.execute_btn)

        self.review_btn = QPushButton("Review Report")
        self.review_btn.setObjectName("reviewBtn")
        self.review_btn.setToolTip(
            "Review a junior associate's draft report against\n"
            "your voice, style, and formatting standards.\n"
            "Applies corrections as Track Changes."
        )
        self.review_btn.clicked.connect(self._on_review_report_clicked)
        self.review_btn.setVisible(False)  # shown only in Word context
        btn_row.addWidget(self.review_btn)

        layout.addLayout(btn_row)

    def _setup_ai_prompt_tab(self):
        """Setup the AI Prompt tab with existing prompt/format/model controls."""
        ai_tab = QWidget()
        ai_layout = QVBoxLayout(ai_tab)
        ai_layout.setContentsMargins(10, 10, 10, 10)
        ai_layout.setSpacing(8)

        # Saved Prompts Section
        prompt_label = QLabel("Saved Prompts:")
        ai_layout.addWidget(prompt_label)

        prompt_row = QHBoxLayout()
        self.prompt_combo = QComboBox()
        self.prompt_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.prompt_combo.currentIndexChanged.connect(self.on_prompt_selected)
        prompt_row.addWidget(self.prompt_combo)

        delete_btn = QPushButton("Delete")
        delete_btn.setObjectName("deleteBtn")
        delete_btn.setFixedWidth(70)
        delete_btn.clicked.connect(self.delete_prompt)
        prompt_row.addWidget(delete_btn)

        ai_layout.addLayout(prompt_row)

        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("background-color: #45475a;")
        ai_layout.addWidget(line)

        # Custom Prompt Section
        custom_label = QLabel("Custom Prompt (or type to add new):")
        ai_layout.addWidget(custom_label)

        self.custom_input = QTextEdit()
        self.custom_input.setPlaceholderText("Enter your prompt here...\nExample: 'Convert this to formal legal language' or 'Summarize in bullet points'")
        self.custom_input.setMaximumHeight(80)
        ai_layout.addWidget(self.custom_input)

        # Save prompt checkbox row
        save_row = QHBoxLayout()
        self.save_name_input = QLineEdit()
        self.save_name_input.setPlaceholderText("Name for saving (optional)")
        self.save_name_input.setFixedWidth(200)
        save_row.addWidget(self.save_name_input)

        save_btn = QPushButton("Save Prompt")
        save_btn.setFixedWidth(100)
        save_btn.clicked.connect(self.save_prompt)
        save_row.addWidget(save_btn)
        save_row.addStretch()

        ai_layout.addLayout(save_row)

        # Separator
        line2 = QFrame()
        line2.setFrameShape(QFrame.Shape.HLine)
        line2.setStyleSheet("background-color: #45475a;")
        ai_layout.addWidget(line2)

        # Format Section
        format_label = QLabel("Output Format:")
        ai_layout.addWidget(format_label)

        format_row = QHBoxLayout()
        self.format_combo = QComboBox()
        self.format_combo.addItems([
            FORMAT_MARKDOWN,
            FORMAT_PLAIN,
            FORMAT_MATCH,
            FORMAT_DEFAULT,
            FORMAT_CUSTOM
        ])
        self.format_combo.currentTextChanged.connect(self._on_format_changed)
        self.format_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        format_row.addWidget(self.format_combo)

        self.format_settings_btn = QPushButton("Settings")
        self.format_settings_btn.setFixedWidth(70)
        self.format_settings_btn.clicked.connect(self._open_format_settings)
        self.format_settings_btn.setEnabled(False)  # Only enabled for Custom
        format_row.addWidget(self.format_settings_btn)

        ai_layout.addLayout(format_row)

        # Format preview label
        self.format_preview = QLabel("")
        self.format_preview.setStyleSheet("color: #a6adc8; font-size: 10px; font-style: italic;")
        self.format_preview.setWordWrap(True)
        ai_layout.addWidget(self.format_preview)

        # Model Selection Section
        model_label = QLabel("AI Model:")
        ai_layout.addWidget(model_label)

        self.model_combo = QComboBox()
        for display_name, provider, model_id in AVAILABLE_MODELS:
            self.model_combo.addItem(f"{display_name}", (provider, model_id))
        self.model_combo.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.model_combo.currentIndexChanged.connect(self._on_model_changed)
        ai_layout.addWidget(self.model_combo)

        # File attachments area
        self.attachment_area = AttachmentArea(self)
        ai_layout.addWidget(self.attachment_area)

        # Use All Text checkbox
        self.use_all_text_check = QCheckBox("Use all document text as context (still replaces only selection)")
        self.use_all_text_check.setStyleSheet("color: #cdd6f4; font-size: 11px;")
        self.use_all_text_check.setToolTip(
            "When checked, sends the entire document to the LLM for context,\n"
            "but only replaces the selected text with the output."
        )
        self.use_all_text_check.setChecked(self.redline_settings.get("use_all_text_default", False))
        self.use_all_text_check.stateChanged.connect(self._save_use_all_text_preference)
        ai_layout.addWidget(self.use_all_text_check)

        # Redline mode checkbox (only visible for Word context)
        self.redline_checkbox = QCheckBox("✏️ Use Redline Mode (Track Changes)")
        self.redline_checkbox.setStyleSheet("color: #cdd6f4; font-size: 11px;")
        self.redline_checkbox.setToolTip(
            "Instead of replacing text, insert AI suggestions as Track Changes "
            "that you can accept/reject in Word"
        )
        self.redline_checkbox.setChecked(self.redline_settings.get("redline_mode_default", False))
        self.redline_checkbox.stateChanged.connect(self._save_redline_preference)
        ai_layout.addWidget(self.redline_checkbox)

        # Legal Research checkbox (only visible for Word context)
        self.legal_research_checkbox = QCheckBox("📚 Perform Legal Research")
        self.legal_research_checkbox.setStyleSheet("color: #cdd6f4; font-size: 11px;")
        self.legal_research_checkbox.setToolTip(
            "Search California case law and statutes, then inject verified\n"
            "legal citations into the AI response. Uses CourtListener API\n"
            "and CA Legislative Info."
        )
        self.legal_research_checkbox.setChecked(
            self.redline_settings.get("legal_research_default", False)
        )
        self.legal_research_checkbox.stateChanged.connect(self._save_legal_research_preference)
        ai_layout.addWidget(self.legal_research_checkbox)

        ai_layout.addStretch()
        self.tab_widget.addTab(ai_tab, "AI Prompt")

    def _setup_case_variables_tab(self):
        """Setup the Case Variables tab for inserting case info into documents."""
        case_tab = QWidget()
        case_layout = QVBoxLayout(case_tab)
        case_layout.setContentsMargins(10, 10, 10, 10)
        case_layout.setSpacing(8)

        # Scroll area for variables list
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("""
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                background-color: #313244;
                width: 10px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background-color: #45475a;
                border-radius: 5px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #6c63ff;
            }
        """)

        # Container for variable rows
        self.variables_container = QWidget()
        self.variables_layout = QVBoxLayout(self.variables_container)
        self.variables_layout.setContentsMargins(0, 0, 0, 0)
        self.variables_layout.setSpacing(6)

        # Placeholder - will be populated when case is detected
        self.no_case_label = QLabel("No case detected.\nOpen a document from a case folder to see variables.")
        self.no_case_label.setStyleSheet("color: #a6adc8; font-style: italic; padding: 20px;")
        self.no_case_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.no_case_label.setWordWrap(True)
        self.variables_layout.addWidget(self.no_case_label)

        self.variables_layout.addStretch()
        scroll.setWidget(self.variables_container)
        case_layout.addWidget(scroll)

        self.tab_widget.addTab(case_tab, "Case Variables")

    def _populate_case_variables(self):
        """Populate the Case Variables tab with detected case info."""
        try:
            self._populate_case_variables_impl()
        except Exception as e:
            # Log but don't crash — this is a non-critical UI update
            try:
                from icharlotte_core.app_crash_handler import log_error
                log_error(f"_populate_case_variables failed: {e}")
            except Exception:
                pass

    def _populate_case_variables_impl(self):
        """Internal implementation of case variable population."""
        # Clear existing variable widgets (except stretch)
        while self.variables_layout.count() > 0:
            item = self.variables_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if not self._detected_case:
            # Show no case message
            self.no_case_label = QLabel(
                self._case_detection_error or "No case detected.\nOpen a document from a case folder to see variables."
            )
            self.no_case_label.setStyleSheet("color: #a6adc8; font-style: italic; padding: 20px;")
            self.no_case_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.no_case_label.setWordWrap(True)
            self.variables_layout.addWidget(self.no_case_label)
            self.variables_layout.addStretch()
            return

        # Define the case variables to display
        case_variables = [
            ("File Number", self._detected_case.get("file_number", "")),
            ("Plaintiff Name", self._detected_case.get("plaintiff_last_name", "")),
            ("Next Hearing", self._detected_case.get("next_hearing_date", "")),
            ("Trial Date", self._detected_case.get("trial_date", "")),
            ("Assigned Attorney", self._detected_case.get("assigned_attorney", "")),
            ("Case Summary", self._detected_case.get("case_summary", "")),
            ("Last Report", self._detected_case.get("last_report_text", "")),
        ]

        for label_text, value in case_variables:
            if not value:
                continue  # Skip empty values

            row_widget = QWidget()
            row_layout = QHBoxLayout(row_widget)
            row_layout.setContentsMargins(0, 0, 0, 0)
            row_layout.setSpacing(8)

            # Label
            label = QLabel(f"{label_text}:")
            label.setFixedWidth(120)
            label.setStyleSheet("color: #89b4fa; font-weight: bold;")
            row_layout.addWidget(label)

            # Value preview (truncated for long text)
            display_value = value if len(value) <= 50 else value[:47] + "..."
            value_label = QLabel(display_value)
            value_label.setStyleSheet("color: #ffffff; font-size: 12px;")
            value_label.setToolTip(value if len(value) > 50 else "")
            value_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            row_layout.addWidget(value_label)

            # Insert button
            insert_btn = QPushButton("Insert")
            insert_btn.setFixedWidth(60)
            insert_btn.setStyleSheet("""
                QPushButton {
                    background-color: #45475a;
                    padding: 4px 8px;
                    font-size: 11px;
                }
                QPushButton:hover {
                    background-color: #6c63ff;
                }
            """)
            # Store the value in the button for the click handler
            insert_btn.setProperty("insert_value", value)
            insert_btn.clicked.connect(lambda checked, v=value: self._insert_case_variable(v))
            row_layout.addWidget(insert_btn)

            self.variables_layout.addWidget(row_widget)

        self.variables_layout.addStretch()

    def _insert_case_variable(self, value: str):
        """Insert a case variable value into Word at the current cursor position."""
        if not value:
            return

        try:
            pythoncom.CoInitialize()

            # Get Word application
            word = None
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except:
                try:
                    word = win32com.client.GetObject(None, "Word.Application")
                except:
                    pass

            if not word:
                self.status_label.setText("Error: Word not running")
                self.status_label.setStyleSheet("color: #f38ba8; font-style: italic;")
                return

            selection = word.Selection
            if not selection:
                self.status_label.setText("Error: No selection available")
                self.status_label.setStyleSheet("color: #f38ba8; font-style: italic;")
                return

            # Insert the text at cursor
            selection.TypeText(value)

            self.status_label.setText("Inserted!")
            self.status_label.setStyleSheet("color: #a6e3a1; font-style: italic;")

            # Auto-close after successful insert
            QTimer.singleShot(500, self.close)

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)[:50]}")
            self.status_label.setStyleSheet("color: #f38ba8; font-style: italic;")

    def _detect_and_update_case(self):
        """Detect case from document and update UI."""
        if self.app_context != APP_CONTEXT_WORD:
            # Case detection only applies to Word documents
            self._detected_case = None
            self._detected_file_number = None
            self._case_detection_error = "Case variables only available in Word"
            self.case_info_label.setText("")
            return

        try:
            self._detected_case, self._detected_file_number, self._case_detection_error = detect_case_from_document()
        except Exception as e:
            logger.error(f"COM error in _detect_and_update_case: {type(e).__name__}: {e}", exc_info=True)
            self._detected_case = None
            self._detected_file_number = None
            self._case_detection_error = f"COM error: {e}"

        if self._detected_case:
            plaintiff = self._detected_case.get("plaintiff_last_name", "Unknown")
            file_num = self._detected_file_number or ""
            self.case_info_label.setText(f"Case: {plaintiff} ({file_num})")
            self.case_info_label.setStyleSheet("color: #a6e3a1; font-size: 11px; padding: 4px;")
        elif self._detected_file_number:
            self.case_info_label.setText(f"Case {self._detected_file_number} not in database")
            self.case_info_label.setStyleSheet("color: #fab387; font-size: 11px; padding: 4px;")
        else:
            self.case_info_label.setText("")

        # Populate the variables tab
        try:
            self._populate_case_variables()
        except Exception as e:
            logger.error(f"COM error in _populate_case_variables: {type(e).__name__}: {e}", exc_info=True)

    def load_prompts(self):
        """Load saved prompts from file (both Word and Outlook prompts)."""
        # Load Word prompts
        self.prompts = []
        if os.path.exists(self.prompts_path):
            try:
                with open(self.prompts_path, 'r', encoding='utf-8') as f:
                    self.prompts = json.load(f)
            except:
                pass

        # Add default Word prompts if none exist
        if not self.prompts:
            self.prompts = [
                {"name": "Improve Writing", "prompt": "Improve this text for clarity and professionalism while maintaining the original meaning."},
                {"name": "Convert to Narrative", "prompt": "Convert this text into a cohesive narrative paragraph format. Maintain all facts and details."},
                {"name": "Formal Legal Tone", "prompt": "Rewrite this in formal legal language appropriate for court documents."},
                {"name": "Summarize", "prompt": "Summarize this text concisely while keeping key points."},
                {"name": "Fix Grammar", "prompt": "Fix any grammar, spelling, or punctuation errors in this text."},
            ]
            self.save_prompts_to_file()

        # Load Outlook prompts
        self.outlook_prompts = []
        if os.path.exists(self.outlook_prompts_path):
            try:
                with open(self.outlook_prompts_path, 'r', encoding='utf-8') as f:
                    self.outlook_prompts = json.load(f)
            except:
                pass

        # Add default Outlook prompts if none exist
        if not self.outlook_prompts:
            self.outlook_prompts = list(DEFAULT_OUTLOOK_PROMPTS)
            self._save_outlook_prompts()

        self.refresh_combo()

    def refresh_combo(self):
        """Refresh the combo box with prompts for current context."""
        self.prompt_combo.clear()
        self.prompt_combo.addItem("-- Select a saved prompt --", None)

        # Get prompts for current context
        if self.app_context == APP_CONTEXT_OUTLOOK:
            prompts = self.outlook_prompts
        else:
            prompts = self.prompts

        for p in prompts:
            self.prompt_combo.addItem(p["name"], p["prompt"])

    def save_prompts_to_file(self):
        """Save Word prompts to file."""
        try:
            os.makedirs(os.path.dirname(self.prompts_path), exist_ok=True)
            with open(self.prompts_path, 'w', encoding='utf-8') as f:
                json.dump(self.prompts, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving prompts: {e}")

    def _save_outlook_prompts(self):
        """Save Outlook prompts to file."""
        try:
            os.makedirs(os.path.dirname(self.outlook_prompts_path), exist_ok=True)
            with open(self.outlook_prompts_path, 'w', encoding='utf-8') as f:
                json.dump(self.outlook_prompts, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error saving Outlook prompts: {e}")

    def on_prompt_selected(self, index):
        """When a saved prompt is selected, populate the custom input."""
        prompt_text = self.prompt_combo.currentData()
        if prompt_text:
            self.custom_input.setPlainText(prompt_text)

    def save_prompt(self):
        """Save the current custom prompt to the appropriate list based on context."""
        name = self.save_name_input.text().strip()
        prompt = self.custom_input.toPlainText().strip()

        if not name:
            QMessageBox.warning(self, "Name Required", "Please enter a name for the prompt.")
            return
        if not prompt:
            QMessageBox.warning(self, "Prompt Required", "Please enter a prompt to save.")
            return

        # Get the appropriate prompts list and save function based on context
        if self.app_context == APP_CONTEXT_OUTLOOK:
            prompts_list = self.outlook_prompts
            save_func = self._save_outlook_prompts
            context_name = "email"
        else:
            prompts_list = self.prompts
            save_func = self.save_prompts_to_file
            context_name = "document"

        # Check for duplicate name
        for p in prompts_list:
            if p["name"].lower() == name.lower():
                reply = QMessageBox.question(
                    self, "Overwrite?",
                    f"A prompt named '{name}' already exists. Overwrite?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.Yes:
                    p["prompt"] = prompt
                    save_func()
                    self.refresh_combo()
                    self.status_label.setText(f"Updated '{name}' ({context_name})")
                return

        prompts_list.append({"name": name, "prompt": prompt})
        save_func()
        self.refresh_combo()
        self.save_name_input.clear()
        self.status_label.setText(f"Saved '{name}' ({context_name})")

    def delete_prompt(self):
        """Delete the selected prompt from the appropriate list based on context."""
        index = self.prompt_combo.currentIndex()
        if index <= 0:
            QMessageBox.warning(self, "Select Prompt", "Please select a prompt to delete.")
            return

        name = self.prompt_combo.currentText()
        reply = QMessageBox.question(
            self, "Delete Prompt?",
            f"Delete '{name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            # Delete from appropriate list based on context
            if self.app_context == APP_CONTEXT_OUTLOOK:
                self.outlook_prompts = [p for p in self.outlook_prompts if p["name"] != name]
                self._save_outlook_prompts()
            else:
                self.prompts = [p for p in self.prompts if p["name"] != name]
                self.save_prompts_to_file()
            self.refresh_combo()
            self.custom_input.clear()
            self.status_label.setText(f"Deleted '{name}'")

    def _update_redline_checkbox_visibility(self):
        """Show/hide redline checkbox, legal research checkbox, and review button based on app context."""
        is_word = self.app_context == APP_CONTEXT_WORD
        self.redline_checkbox.setVisible(is_word)
        self.legal_research_checkbox.setVisible(is_word)
        self.review_btn.setVisible(is_word)

    def _save_redline_preference(self):
        """Save the current redline checkbox state to settings."""
        self.redline_settings["redline_mode_default"] = self.redline_checkbox.isChecked()
        save_redline_settings(GEMINI_DATA_DIR, self.redline_settings)

    def _save_use_all_text_preference(self):
        """Save the use-all-text checkbox state to settings."""
        self.redline_settings["use_all_text_default"] = self.use_all_text_check.isChecked()
        save_redline_settings(GEMINI_DATA_DIR, self.redline_settings)

    def _save_legal_research_preference(self, state):
        """Persist legal research checkbox state."""
        self.redline_settings["legal_research_default"] = bool(state)
        self._save_redline_settings()

    def _save_redline_settings(self):
        """Helper to persist redline settings to disk."""
        save_redline_settings(GEMINI_DATA_DIR, self.redline_settings)

    def _get_legal_research_engine(self):
        """Lazy-initialize the legal research engine."""
        if not hasattr(self, '_legal_research_engine') or self._legal_research_engine is None:
            import os
            from icharlotte_core.legal_research.engine import LegalResearchEngine
            token = os.environ.get("COURTLISTENER_API_TOKEN", "")
            if not token:
                return None
            self._legal_research_engine = LegalResearchEngine(courtlistener_token=token)
        return self._legal_research_engine

    def set_app_context(self, context: str, inspector=None):
        """Set the application context and update UI accordingly."""
        self.app_context = context
        self.active_inspector = inspector

        try:
            if context == APP_CONTEXT_OUTLOOK:
                self.title_label.setText("AI Assistant for Outlook")
                self.setWindowTitle("AI Assistant - Email")
                # Hide case variables tab for Outlook
                self.tab_widget.setTabVisible(1, False)
            else:
                self.title_label.setText("AI Assistant for Word")
                self.setWindowTitle("AI Assistant - Document")
                # Show case variables tab for Word
                self.tab_widget.setTabVisible(1, True)
                # Detect case from document path
                self._detect_and_update_case()

            # Update redline checkbox visibility based on context
            self._update_redline_checkbox_visibility()

            self.refresh_combo()  # Refresh prompts for context
        except Exception as e:
            logger.error(f"Error in set_app_context({context}): {type(e).__name__}: {e}", exc_info=True)

    def _on_format_changed(self, format_type: str):
        """Handle format dropdown change."""
        self.format_settings_btn.setEnabled(format_type == FORMAT_CUSTOM)

        # Pre-capture format when Match Selection is chosen so user can see preview
        if format_type == FORMAT_MATCH:
            self._pre_capture_format()

        self._update_format_preview()

    def _pre_capture_format(self):
        """Pre-capture format from current selection to show in preview."""
        try:
            if self.app_context == APP_CONTEXT_OUTLOOK:
                self._capture_outlook_selection_format()
            else:
                word = self._get_word_app()
                if word:
                    self._capture_selection_format(word)
        except Exception as e:
            print(f"Pre-capture format error: {e}")
            self._captured_format = None

    def _update_format_preview(self):
        """Update the format preview label with enhanced formatting info."""
        format_type = self.format_combo.currentText()

        if format_type == FORMAT_PLAIN:
            self.format_preview.setText("No formatting will be applied")
        elif format_type == FORMAT_MATCH:
            # Show captured format details if available
            if self._captured_format:
                parts = []

                # Show paragraph style if captured
                style_name = self._captured_format.get('style_name')
                if style_name:
                    parts.append(f"Style: {style_name}")

                # Show font info
                font_name = self._captured_format.get('font_name')
                font_size = self._captured_format.get('font_size')
                if font_name and font_size:
                    parts.append(f"{font_name} {font_size}pt")

                # Show table context
                if self._captured_format.get('in_table'):
                    parts.append("[Table cell]")

                # Show track changes
                if self._captured_format.get('track_changes_enabled'):
                    parts.append("[Track Changes]")

                if parts:
                    self.format_preview.setText(" | ".join(parts))
                else:
                    self.format_preview.setText("Formatting will match the selected text")
            else:
                self.format_preview.setText("Formatting will match the selected text (captured on execute)")
        elif format_type == FORMAT_MARKDOWN:
            self.format_preview.setText("Markdown: # headings, **bold**, *italic*, - bullets, `code`")
        elif format_type == FORMAT_DEFAULT:
            self.format_preview.setText("Uses Word's default paragraph style")
        elif format_type == FORMAT_CUSTOM:
            if self.custom_format_settings:
                s = self.custom_format_settings
                preview = f"{s.get('font_name', 'Default')}, {s.get('font_size', 12)}pt"
                styles = []
                if s.get('bold'):
                    styles.append('Bold')
                if s.get('italic'):
                    styles.append('Italic')
                if s.get('underline'):
                    styles.append('Underline')
                if styles:
                    preview += f", {'/'.join(styles)}"
                self.format_preview.setText(preview)
            else:
                self.format_preview.setText("Click Settings to configure custom format")

    def _open_format_settings(self):
        """Open the custom format settings dialog."""
        dialog = CustomFormatDialog(self, self.custom_format_settings)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.custom_format_settings = dialog.get_settings()
            self._save_format_settings()
            self._update_format_preview()

    def _load_format_settings(self):
        """Load custom format settings from file."""
        if os.path.exists(self.format_settings_path):
            try:
                with open(self.format_settings_path, 'r', encoding='utf-8') as f:
                    self.custom_format_settings = json.load(f)
            except:
                pass

    def _save_format_settings(self):
        """Save custom format settings to file."""
        try:
            os.makedirs(os.path.dirname(self.format_settings_path), exist_ok=True)
            with open(self.format_settings_path, 'w', encoding='utf-8') as f:
                json.dump(self.custom_format_settings, f, indent=2)
        except Exception as e:
            print(f"Error saving format settings: {e}")

    def _load_model_settings(self):
        """Load model selection from file."""
        if os.path.exists(self.model_settings_path):
            try:
                with open(self.model_settings_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.selected_model_index = data.get('model_index', DEFAULT_MODEL_INDEX)
                    # Validate index is within range
                    if self.selected_model_index >= len(AVAILABLE_MODELS):
                        self.selected_model_index = DEFAULT_MODEL_INDEX
            except:
                pass

        # Set the combo box to the saved selection
        if hasattr(self, 'model_combo'):
            self.model_combo.setCurrentIndex(self.selected_model_index)

    def _save_model_settings(self):
        """Save model selection to file."""
        try:
            os.makedirs(os.path.dirname(self.model_settings_path), exist_ok=True)
            with open(self.model_settings_path, 'w', encoding='utf-8') as f:
                json.dump({'model_index': self.selected_model_index}, f, indent=2)
        except Exception as e:
            print(f"Error saving model settings: {e}")

    def _on_model_changed(self, index: int):
        """Handle model dropdown change."""
        self.selected_model_index = index
        self._save_model_settings()
        # Get model info for display
        if index < len(AVAILABLE_MODELS):
            display_name, provider, model_id = AVAILABLE_MODELS[index]
            print(f"Model changed to: {display_name} ({provider})")

    def _get_selected_model(self) -> tuple:
        """Get the currently selected model (provider, model_id)."""
        if self.selected_model_index < len(AVAILABLE_MODELS):
            _, provider, model_id = AVAILABLE_MODELS[self.selected_model_index]
            return provider, model_id
        # Fallback to default
        _, provider, model_id = AVAILABLE_MODELS[DEFAULT_MODEL_INDEX]
        return provider, model_id

    def _capture_selection_format(self, word):
        """Capture the formatting of the current selection in Word.

        Enhanced to capture:
        - Paragraph styles (Heading 1, Normal, Body Text, etc.)
        - Table cell context and formatting
        - Track Changes state
        - Inline font formatting
        """
        try:
            selection = word.Selection
            font = selection.Font
            para = selection.ParagraphFormat

            # Word constants for underline
            # wdUnderlineNone = 0, wdUnderlineSingle = 1, etc.
            underline_val = font.Underline
            has_underline = underline_val != 0 and underline_val != 9999999  # 9999999 = mixed

            self._captured_format = {
                'font_name': font.Name if font.Name != '' else None,
                'font_size': font.Size if font.Size != 9999999 else 12,
                'bold': font.Bold == -1,  # -1 = True in Word
                'italic': font.Italic == -1,
                'underline': has_underline,
                'first_indent': para.FirstLineIndent / 72 if para.FirstLineIndent != 9999999 else 0,
                'left_indent': para.LeftIndent / 72 if para.LeftIndent != 9999999 else 0,
            }

            # Capture paragraph style name
            try:
                style = selection.Style
                if style:
                    style_name = style.NameLocal  # Localized name (e.g., "Heading 1")
                    self._captured_format['style_name'] = style_name
                    print(f"Captured paragraph style: {style_name}")
            except Exception as e:
                print(f"Could not capture style: {e}")
                self._captured_format['style_name'] = None

            # Detect if selection is inside a table cell
            try:
                # wdWithInTable = 12
                if selection.Information(12):  # wdWithInTable
                    self._captured_format['in_table'] = True

                    # Capture table cell info
                    try:
                        cell = selection.Cells(1) if selection.Cells.Count > 0 else None
                        if cell:
                            # Capture cell shading (background color)
                            shading = cell.Shading
                            if shading:
                                self._captured_format['cell_shading_color'] = shading.BackgroundPatternColor
                                self._captured_format['cell_shading_pattern'] = shading.Texture

                            # Capture cell vertical alignment
                            self._captured_format['cell_vertical_align'] = cell.VerticalAlignment

                            # Capture cell width if available
                            try:
                                self._captured_format['cell_width'] = cell.Width
                            except:
                                pass

                            print(f"Captured table cell formatting")
                    except Exception as e:
                        print(f"Could not capture cell details: {e}")
                else:
                    self._captured_format['in_table'] = False
            except Exception as e:
                print(f"Could not detect table context: {e}")
                self._captured_format['in_table'] = False

            # Capture Track Changes state
            try:
                doc = word.ActiveDocument
                if doc:
                    # TrackRevisions property indicates if Track Changes is enabled
                    self._captured_format['track_changes_enabled'] = doc.TrackRevisions
                    if doc.TrackRevisions:
                        print("Track Changes is enabled - will preserve revision tracking")
            except Exception as e:
                print(f"Could not check Track Changes: {e}")
                self._captured_format['track_changes_enabled'] = False

            # Capture line spacing
            try:
                # wdLineSpaceSingle = 0, wdLineSpace1pt5 = 1, wdLineSpaceDouble = 2
                # wdLineSpaceAtLeast = 3, wdLineSpaceExactly = 4, wdLineSpaceMultiple = 5
                line_spacing_rule = para.LineSpacingRule
                self._captured_format['line_spacing_rule'] = line_spacing_rule
                if line_spacing_rule in [3, 4, 5]:  # wdLineSpaceAtLeast, Exactly, Multiple
                    self._captured_format['line_spacing'] = para.LineSpacing
            except:
                pass

            # Capture font color
            try:
                font_color = font.Color
                if font_color != 9999999:  # Not mixed
                    self._captured_format['font_color'] = font_color
            except:
                pass

            print(f"Captured format: {self._captured_format}")
        except Exception as e:
            print(f"Error capturing format: {e}")
            self._captured_format = None

    def _on_cancel_clicked(self):
        """Handle cancel button click - either close dialog or cancel running task."""
        if self._worker_thread and self._worker_thread.isRunning():
            # Cancel the running LLM task
            self._worker_thread.cancel()
            self.status_label.setText("Cancelling...")
            self.cancel_btn.setEnabled(False)
            QTimer.singleShot(200, self._cleanup_and_close)
        elif self._review_thread and self._review_thread.isRunning():
            # Cancel the running review task
            self._review_thread.cancel()
            self.status_label.setText("Cancelling review...")
            self.cancel_btn.setEnabled(False)
            QTimer.singleShot(500, self._cleanup_and_close)
        else:
            # No task running, just close
            self.close()

    def _cleanup_and_close(self):
        """Clean up worker threads and close dialog."""
        if self._worker_thread:
            if self._worker_thread.isRunning():
                self._worker_thread.terminate()
                self._worker_thread.wait(500)
            self._worker_thread = None
        if self._review_thread:
            if self._review_thread.isRunning():
                self._review_thread.terminate()
                self._review_thread.wait(500)
            self._review_thread = None
        self.close()

    def execute(self):
        """Execute the LLM processing."""
        prompt = self.custom_input.toPlainText().strip()
        if not prompt:
            QMessageBox.warning(self, "Prompt Required", "Please enter or select a prompt.")
            return

        # Capture all UI state before hiding popup
        self._exec_format_type = self.format_combo.currentText()
        self._exec_use_all_text = self.use_all_text_check.isChecked()
        self._exec_legal_research = self.legal_research_checkbox.isChecked()
        self._exec_redline_checked = self.redline_checkbox.isChecked()

        # Generate a prep_id so we can show the status bar immediately
        import uuid
        self._prep_id = str(uuid.uuid4())[:8]

        # Hide popup immediately so user can keep working
        self.hide()

        # Show status bar with "Preparing..." row right away
        tsb = TaskStatusBar.instance()
        tsb.show_preparing(self._prep_id, prompt[:40])
        QApplication.processEvents()  # Force paint so status bar renders immediately

        # Small delay to allow UI to update before _do_execute
        QTimer.singleShot(50, lambda: self._do_execute(prompt))

    def _do_execute(self, prompt: str):
        """Actually execute the LLM call - dispatch to Word or Outlook handler."""
        try:
            print(f"[DEBUG] _do_execute entered. app_context={self.app_context}")
            # Dispatch based on context
            if self.app_context == APP_CONTEXT_OUTLOOK:
                self.show()  # Re-show for Outlook (uses old worker thread path)
                self._do_execute_outlook(prompt)
                return

            # Default: Word context
            # Check if Word is running with a document open
            QApplication.processEvents()  # Keep UI responsive
            word = self._get_word_app()
            QApplication.processEvents()  # Keep UI responsive

            if not word:
                prep_id = getattr(self, '_prep_id', None)
                if prep_id:
                    tsb = TaskStatusBar.instance()
                    tsb._remove_task_row(prep_id)
                    tsb._on_all_done()
                    self._prep_id = None
                QMessageBox.warning(None, "Word Not Found",
                    "Microsoft Word is not running or no document is open.\n\nPlease open a Word document first, then try again.")
                self.close()
                return

            # Use pre-captured UI state (captured in execute() before popup was hidden)
            format_type = getattr(self, '_exec_format_type', FORMAT_PLAIN)
            use_all_text = getattr(self, '_exec_use_all_text', False)

            # Capture format if "Match Selection" is chosen
            if format_type == FORMAT_MATCH:
                self._capture_selection_format(word)
                QApplication.processEvents()  # Keep UI responsive

            # Get Word selection
            selected_text, has_selection = self._get_word_selection_internal(word)
            QApplication.processEvents()  # Keep UI responsive

            # When redline mode is active, prepend instruction to preserve unchanged text
            redline_prefix = ""
            if self._redline_mode_active:
                redline_prefix = (
                    "IMPORTANT — REDLINE MODE IS ACTIVE: Your output will be compared "
                    "word-by-word against the original text to generate Track Changes. "
                    "You MUST follow these rules:\n"
                    "1. Preserve any text that does not need to change EXACTLY as-is — "
                    "same wording, same sentence structure, same word order.\n"
                    "2. PRESERVE THE EXACT PARAGRAPH STRUCTURE — keep the same number of "
                    "paragraphs and blank lines between them. Each paragraph must start "
                    "and end at the same boundaries as the original. Do NOT merge paragraphs, "
                    "split paragraphs, or remove blank lines.\n"
                    "3. Do NOT rephrase, reorganize, or rewrite portions that are already correct.\n"
                    "4. Only modify the specific words or sentences that need to change.\n"
                    "5. If a sentence is fine as-is, copy it verbatim.\n\n"
                )
                prompt = redline_prefix + prompt

            if use_all_text and has_selection and selected_text:
                # Get all document text for context
                all_text = self._get_all_document_text()
                if all_text:
                    # Build prompt with full document context but mark selected text
                    full_prompt = (
                        f"{prompt}\n\n"
                        f"=== FULL DOCUMENT (for context) ===\n{all_text}\n\n"
                        f"=== SELECTED TEXT TO PROCESS ===\n{selected_text}\n\n"
                        f"Process only the SELECTED TEXT above according to the instructions. "
                        f"Use the full document for context but output only the processed version of the selected text."
                    )
                    self.status_label.setText("Processing with full document context...")
                else:
                    # Fallback to just selected text
                    full_prompt = f"{prompt}\n\nText to process:\n{selected_text}"
                    self.status_label.setText("Processing selected text...")
            elif has_selection and selected_text:
                # Build full prompt with selected text only
                full_prompt = f"{prompt}\n\nText to process:\n{selected_text}"
                self.status_label.setText("Processing selected text...")
            elif use_all_text:
                # No selection but use all text - process entire document
                all_text = self._get_all_document_text()
                if all_text:
                    full_prompt = f"{prompt}\n\nText to process:\n{all_text}"
                    self.status_label.setText("Processing entire document...")
                else:
                    full_prompt = prompt
                    self.status_label.setText("Processing prompt...")
            else:
                # No selection, just use the prompt
                full_prompt = prompt
                self.status_label.setText("Processing prompt...")

            # Append any attached file context
            attachment_context = self.attachment_area.get_attachment_context()
            if attachment_context:
                full_prompt += attachment_context

            QApplication.processEvents()

            # Legal Research: defer to worker thread (no longer blocks main thread)
            _do_legal_research = getattr(self, '_exec_legal_research', False)
            print(f"[DEBUG] legal_research = {_do_legal_research}")
            _legal_research_engine = None
            _legal_research_model = None
            if _do_legal_research:
                _legal_research_engine = self._get_legal_research_engine()
                if _legal_research_engine:
                    _legal_research_model = self._get_selected_model()
                    print("[DEBUG] Legal research will run in worker thread")
                else:
                    _do_legal_research = False
                    print("[LegalResearch] No COURTLISTENER_API_TOKEN in .env — skipping research")
            else:
                print("[DEBUG] Legal research checkbox NOT checked — skipping")
            self._pending_research_result = None

            print("[DEBUG] Past legal research block, about to build TaskData...")
            # ── Build TaskData and submit to TaskManager ──
            print("[DEBUG] Reached TaskManager path — building TaskData...")
            provider, model_id = self._get_selected_model()
            print(f"[DEBUG] Selected model: {provider}/{model_id}")

            # Build system prompt (same logic as _default_llm_call)
            if self._redline_mode_active:
                system_prompt = (
                    "You are a helpful writing assistant operating in REDLINE MODE. "
                    "Your output will be diffed against the original text to produce "
                    "Track Changes in Microsoft Word. You MUST preserve unchanged text "
                    "EXACTLY as written — same words, same order, same punctuation. "
                    "Only modify the specific parts that need changing per the user's "
                    "instructions. Output only the requested text without any preamble, "
                    "explanation, or markdown formatting unless specifically asked."
                )
            else:
                system_prompt = (
                    "You are a helpful writing assistant for Microsoft Word documents. "
                    "Follow the user's instructions precisely. Output only the requested "
                    "text without any preamble or explanation.\n\n"
                    "FORMAT RULES (strict defaults — override ONLY if the user explicitly "
                    "asks for headings, bullet points, numbered lists, or other structure):\n"
                    "- Always write in plain narrative prose. No headings (#, ##, ###), "
                    "no bullet points, no numbered lists, no markdown formatting.\n"
                    "- Use **bold** sparingly and only when the user asks for emphasis.\n"
                    "- Do not use code blocks or HTML.\n\n"
                    "CONTINUATION RULE: When the user asks you to finish a sentence, "
                    "fill in a blank, or continue from a specific point, output ONLY "
                    "the new text that comes after the existing content. Do NOT repeat "
                    "any part of the preceding sentence, paragraph, or context that was "
                    "provided to you. Start exactly where the existing text leaves off."
                )

            task_data = TaskData(
                document_name=self._original_document_name or "",
                document_com=self._original_document,
                has_selection=self._original_has_selection,
                original_text=getattr(self, '_original_text', "") or "",
                original_text_raw=getattr(self, '_original_text_raw', "") or "",
                captured_format=self._captured_format,
                format_type=format_type,
                redline_mode_active=self._redline_mode_active,
                redline_settings=self.redline_settings.copy() if self.redline_settings else {},
                research_result=None,
                do_legal_research=_do_legal_research,
                legal_research_engine=_legal_research_engine,
                legal_research_model=_legal_research_model,
                prompt_preview=prompt[:40],
                provider=provider,
                model_id=model_id,
                system_prompt=system_prompt,
                full_prompt=full_prompt,
            )

            # Create bookmark for range tracking (auto-adjusts on concurrent edits)
            if self._original_range_start is not None and self._original_document:
                bm_name = _create_task_bookmark(
                    self._original_document,
                    self._original_range_start,
                    self._original_range_end if self._original_range_end is not None else self._original_range_start,
                    task_data.task_id,
                )
                task_data.bookmark_name = bm_name

            # Upgrade the "preparing" row BEFORE submitting (submit emits task_added
            # which would create a duplicate row if prep row isn't renamed first)
            prep_id = getattr(self, '_prep_id', None)
            if prep_id:
                TaskStatusBar.instance().upgrade_preparing(prep_id, task_data.task_id)
                self._prep_id = None

            # Submit to TaskManager for parallel execution
            print("[DEBUG] About to submit to TaskManager...")
            TaskManager.instance().submit_task(task_data)
            print("[DEBUG] TaskManager.submit_task() returned OK. Task submitted successfully.")

            # Close popup for good (already hidden since execute())
            self.close()

        except Exception as e:
            import traceback
            print(f"[DEBUG] _do_execute EXCEPTION: {e}")
            traceback.print_exc()
            logger.error(f"_do_execute failed: {type(e).__name__}: {e}", exc_info=True)
            # Clean up the "preparing" row in status bar
            prep_id = getattr(self, '_prep_id', None)
            if prep_id:
                tsb = TaskStatusBar.instance()
                tsb._remove_task_row(prep_id)
                tsb._on_all_done()
                self._prep_id = None
            # Popup is already hidden — show standalone error dialog and close
            QMessageBox.critical(None, "AI Assistant Error", f"Failed to process:\n{e}")
            self.close()

    def _on_llm_result(self, result: str, error: str):
        """Handle the result from the LLM worker thread."""
        self._worker_thread = None

        # Check if cancelled
        if error == "cancelled":
            self.status_label.setText("Cancelled")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            self.cancel_btn.setEnabled(True)
            return

        # Check for error
        if error:
            self.status_label.setText(f"Error: {error[:50]}")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            self.cancel_btn.setEnabled(True)
            QMessageBox.critical(self, "Error", f"Failed to process:\n{error}")
            return

        # Check for empty result
        if not result:
            self.status_label.setText("No response from LLM")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            self.cancel_btn.setEnabled(True)
            return

        # Insert result into the appropriate application
        try:
            if self._is_outlook_task:
                self._set_outlook_text(result, self._pending_format_type)
            else:
                # Restore original document context before applying
                # (handles case where user switched documents while LLM ran)
                # NOTE: _restore_original_context() also locks Word UI
                # (Interactive=False) to prevent cursor movement during
                # insertion. We MUST unlock in the finally block below.
                word = None
                try:
                    word, selection = self._restore_original_context()
                except Exception as ctx_err:
                    print(f"Could not restore original context: {ctx_err}")
                    QApplication.clipboard().setText(result)
                    QMessageBox.warning(
                        self, "Document Changed",
                        f"Could not return to original document:\n{ctx_err}\n\n"
                        "The result has been copied to your clipboard instead."
                    )
                    self.status_label.setText("Result copied to clipboard")
                    QTimer.singleShot(1500, self.close)
                    return

                try:
                    # Check if redline mode is active
                    if self._redline_mode_active:
                        # Redline mode - apply changes as Track Changes
                        try:
                            # Apply redlines using diff-match-patch
                            # Use raw (un-stripped) text so positions match Word range
                            original_for_diff = getattr(self, '_original_text_raw', self._original_text)
                            success = self._apply_redlines(
                                word,
                                selection,
                                original_for_diff,
                                result
                            )

                            if success:
                                self.status_label.setText("✓ Redlines applied!")
                                self.status_label.setStyleSheet("color: #a6e3a1; font-style: italic;")
                            else:
                                # Fallback to replace mode
                                if self.redline_settings.get("redline_fallback_notify", True):
                                    self.status_label.setText("⚠ Applied as replacement (redline unavailable)")
                                    self.status_label.setStyleSheet("color: #f9e2af; font-style: italic;")
                                self._set_word_text_internal(result, self._pending_format_type,
                                                            word=word, selection=selection)

                        except Exception as redline_error:
                            print(f"Redline application failed: {redline_error}")
                            # Fallback to replace mode on error
                            if self.redline_settings.get("redline_fallback_notify", True):
                                self.status_label.setText("⚠ Applied as replacement (redline error)")
                                self.status_label.setStyleSheet("color: #f9e2af; font-style: italic;")
                            self._set_word_text_internal(result, self._pending_format_type,
                                                        word=word, selection=selection)
                    else:
                        # Replace mode
                        self._set_word_text_internal(result, self._pending_format_type,
                                                    word=word, selection=selection)
                finally:
                    # ALWAYS unlock Word UI after insertion completes
                    if word:
                        self._unlock_word_ui(word)

            # Show legal research verification summary if available
            if hasattr(self, '_pending_research_result') and self._pending_research_result:
                rr = self._pending_research_result
                if rr.verification:
                    pass_count = sum(1 for v in rr.verification if v.status == "PASS")
                    fixed_count = sum(1 for v in rr.verification if v.status == "FIXED")
                    flagged_count = sum(1 for v in rr.verification if v.status == "FLAGGED")
                    summary_parts = []
                    if pass_count:
                        summary_parts.append(f"✓ {pass_count} verified")
                    if fixed_count:
                        summary_parts.append(f"⚠ {fixed_count} corrected")
                    if flagged_count:
                        summary_parts.append(f"✗ {flagged_count} flagged")
                    if summary_parts:
                        self.status_label.setText(f"Citations: {', '.join(summary_parts)}")
                        QTimer.singleShot(3000, self.close)
                    else:
                        self.status_label.setText("Done!")
                        QTimer.singleShot(500, self.close)
                else:
                    self.status_label.setText("Done!")
                    QTimer.singleShot(500, self.close)
                self._pending_research_result = None
            else:
                self.status_label.setText("Done!")
                QTimer.singleShot(500, self.close)
        except Exception as e:
            self.status_label.setText(f"Error inserting: {str(e)[:50]}")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            self.cancel_btn.setEnabled(True)
            QMessageBox.critical(self, "Error", f"Failed to insert result:\n{e}")

    # ────── Report Review ──────

    def _on_review_report_clicked(self):
        """Handle the Review Report button click."""
        try:
            word = self._get_word_app()
            if not word:
                QMessageBox.warning(self, "Word Not Found",
                    "Microsoft Word is not running or no document is open.\n"
                    "Please open a Word document first, then try again.")
                return

            doc = word.ActiveDocument
            if not doc:
                QMessageBox.warning(self, "No Document",
                    "No Word document is open.")
                return

            # Detect case from document path
            case_detected = False
            case_path = None
            file_number = None
            try:
                case_info, file_num, error = detect_case_from_document()
                if file_num:
                    file_number = file_num
                    case_detected = True
                    # Resolve case folder path from doc path
                    try:
                        doc_path = doc.FullName
                        # Walk up to find the case folder (contains file number pattern)
                        parts = doc_path.replace('/', '\\').split('\\')
                        for i, part in enumerate(parts):
                            if re.search(r'\d{4}\.\d{3}', part) or re.search(r'^\d{1,3}\s*-', part):
                                case_path = '\\'.join(parts[:i+1])
                                break
                        if not case_path:
                            # Try parent of NOTES folder
                            for i, part in enumerate(parts):
                                if part.upper() in ('NOTES', 'AI OUTPUT'):
                                    case_path = '\\'.join(parts[:i])
                                    break
                    except:
                        pass
            except:
                pass

            # Check for selection
            has_selection = False
            selection_range = None
            try:
                sel = word.Selection
                if sel and sel.Type != 1:  # wdSelectionIP = 1
                    sel_text = sel.Text
                    if sel_text and len(sel_text.strip()) > 50:
                        has_selection = True
                        selection_range = (sel.Range.Start, sel.Range.End)
            except:
                pass

            # Show the review dialog
            from icharlotte_core.ui.report_review_dialog import ReportReviewDialog
            dialog = ReportReviewDialog(
                has_selection=has_selection,
                case_detected=case_detected,
                parent=self,
            )
            if dialog.exec() != dialog.DialogCode.Accepted:
                return

            options = dialog.get_options()

            # Determine selection range
            sel_range = None
            if options["scope"] == "selection" and selection_range:
                sel_range = selection_range

            # Create reviewer
            from icharlotte_core.report_reviewer import ReportReviewer
            reviewer = ReportReviewer(doc_com=doc, case_path=case_path)

            # Disable buttons, show progress
            self.execute_btn.setEnabled(False)
            self.review_btn.setEnabled(False)
            self.cancel_btn.setText("Cancel Review")
            self.status_label.setText("Starting report review...")
            self.status_label.setStyleSheet("color: #a6adc8; font-style: italic;")
            QApplication.processEvents()

            # Capture doc name for re-acquisition in worker thread
            try:
                doc_name = doc.Name
            except Exception:
                doc_name = None

            # Start worker thread
            self._review_thread = ReviewWorkerThread(
                reviewer=reviewer,
                selection_range=sel_range,
                include_case_data=options["include_case_data"],
                extra_doc_paths=options["extra_doc_paths"] or None,
                doc_name=doc_name,
                parent=self,
            )
            self._review_thread.progress.connect(self._on_review_progress)
            self._review_thread.finished.connect(self._on_review_finished)
            self._review_thread.start()

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)[:60]}")
            self.status_label.setStyleSheet("color: #f38ba8; font-style: italic;")
            self.execute_btn.setEnabled(True)
            self.review_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            import traceback
            traceback.print_exc()

    def _on_review_progress(self, msg: str):
        """Update status label with review progress."""
        self.status_label.setText(msg)
        QApplication.processEvents()

    def _on_review_finished(self, structural_result, total_changes: int, error: str):
        """Handle review completion."""
        self._review_thread = None
        self.execute_btn.setEnabled(True)
        self.review_btn.setEnabled(True)
        self.cancel_btn.setText("Cancel")

        if error == "cancelled":
            self.status_label.setText("Review cancelled")
            self.status_label.setStyleSheet("color: #f9e2af; font-style: italic;")
            return

        if error:
            self.status_label.setText(f"Review error: {error[:60]}")
            self.status_label.setStyleSheet("color: #f38ba8; font-style: italic;")
            QMessageBox.critical(self, "Review Error", f"Report review failed:\n{error}")
            return

        # Success
        if total_changes > 0:
            self.status_label.setText(
                f"Review complete — {total_changes} tracked changes applied"
            )
            self.status_label.setStyleSheet("color: #a6e3a1; font-style: italic;")
        else:
            self.status_label.setText("Review complete — no changes needed")
            self.status_label.setStyleSheet("color: #a6e3a1; font-style: italic;")

    def _get_word_app(self):
        """Get a connection to the running Word application with active document."""
        word = None

        # Kill any zombie Word processes first
        kill_zombie_word_processes()

        # Try to connect to Word
        try:
            word = win32com.client.GetActiveObject("Word.Application")
        except:
            try:
                word = win32com.client.GetObject(Class="Word.Application")
            except:
                try:
                    word = win32com.client.Dispatch("Word.Application")
                except:
                    print("Could not connect to Word at all")
                    return None

        if not word:
            print("No Word connection")
            return None

        # Check for regular documents first
        try:
            if word.Documents.Count > 0:
                print(f"Connected to Word - {word.Documents.Count} docs open")
                return word
        except:
            pass

        # Check for Protected View windows
        try:
            pv_count = word.ProtectedViewWindows.Count
            if pv_count > 0:
                print(f"Found {pv_count} Protected View window(s)")
                # Try to activate/edit the first Protected View window
                # This will convert it to a regular document
                try:
                    pv_window = word.ProtectedViewWindows(1)
                    print(f"Protected View document: {pv_window.Document.Name}")
                    # Note: We can still work with Selection even in Protected View
                    return word
                except Exception as e:
                    print(f"Could not access Protected View window: {e}")
        except Exception as e:
            print(f"Error checking Protected View: {e}")

        # Check if Word has any windows at all via Windows API
        try:
            word_windows = []
            def enum_callback(hwnd, results):
                if win32gui.IsWindowVisible(hwnd):
                    class_name = win32gui.GetClassName(hwnd)
                    if class_name == "OpusApp":
                        title = win32gui.GetWindowText(hwnd)
                        results.append((hwnd, title))
                return True
            win32gui.EnumWindows(enum_callback, word_windows)

            if word_windows:
                print(f"Found {len(word_windows)} Word window(s) via Windows API")
                for hwnd, title in word_windows:
                    print(f"   - {title}")
                # Word has windows, so let's try to use Selection anyway
                # Sometimes Documents.Count is 0 but Selection still works
                return word
        except Exception as e:
            print(f"Window enumeration failed: {e}")

        print("Word running but no accessible documents")
        return None

    def _restore_original_context(self) -> tuple:
        """Activate the original document, lock Word UI, and restore selection.

        Call this before applying LLM results to ensure changes go to the
        correct document and location, even if the user switched documents
        or moved the cursor while the LLM was running.

        IMPORTANT: This method locks Word's UI (Interactive=False) to prevent
        the user from moving the cursor during insertion. The caller MUST
        call _unlock_word_ui(word) in a finally block after insertion is done.

        Returns:
            (word_app, selection) tuple

        Raises:
            Exception if the original document is no longer open
        """
        word = self._get_word_app()
        if not word:
            raise Exception("Could not connect to Word")

        doc = None

        # Step 1: Try to activate the stored document COM object
        if self._original_document:
            try:
                _ = self._original_document.Name  # test if COM ref is alive
                self._original_document.Activate()
                doc = self._original_document
                print(f"Activated original document via COM reference: {doc.Name}")
            except Exception as e:
                print(f"Stored document COM reference stale: {e}")
                self._original_document = None

        # Step 2: Fallback — find document by name
        if doc is None and self._original_document_name:
            try:
                for i in range(1, word.Documents.Count + 1):
                    candidate = word.Documents(i)
                    if candidate.Name == self._original_document_name:
                        candidate.Activate()
                        doc = candidate
                        self._original_document = doc
                        print(f"Found and activated document by name: {doc.Name}")
                        break
            except Exception as e:
                print(f"Could not find document by name '{self._original_document_name}': {e}")

        # Step 3: Document not found — it was closed
        if doc is None:
            raise Exception(
                f"Original document '{self._original_document_name or 'unknown'}' "
                f"is no longer open"
            )

        # Step 4: Lock Word UI BEFORE selecting the range.
        # This prevents the user from moving the cursor between Select()
        # and the subsequent TypeText() calls. COM cross-process calls pump
        # Windows messages, so without this lock, user clicks/keystrokes
        # processed during COM calls could move the cursor mid-insertion.
        try:
            word.ScreenUpdating = False
            word.Interactive = False
            self._word_ui_locked = True
            print("Locked Word UI for insertion")
        except Exception as e:
            print(f"Could not lock Word UI: {e}")
            self._word_ui_locked = False

        # Step 5: Restore the selection to the original range
        if self._original_range_start is not None and self._original_range_end is not None:
            try:
                if self._original_has_selection:
                    target_range = doc.Range(
                        self._original_range_start,
                        self._original_range_end
                    )
                else:
                    # Insertion point — place cursor at original position
                    target_range = doc.Range(
                        self._original_range_start,
                        self._original_range_start
                    )
                target_range.Select()
                print(f"Restored selection: {self._original_range_start}-{self._original_range_end}")
            except Exception as e:
                print(f"Could not restore original range (document may have been edited): {e}")

        selection = word.Selection
        if not selection:
            self._unlock_word_ui(word)
            raise Exception("Could not access Word selection after restoring context")

        return word, selection

    def _unlock_word_ui(self, word):
        """Re-enable Word UI after insertion is complete.

        Must be called in a finally block after _restore_original_context().
        """
        if not getattr(self, '_word_ui_locked', False):
            return
        try:
            word.Interactive = True
            word.ScreenUpdating = True
            self._word_ui_locked = False
            print("Unlocked Word UI")
        except Exception as e:
            print(f"Could not unlock Word UI: {e}")
            # Last resort: try to get a fresh Word connection and unlock
            try:
                fresh_word = win32com.client.GetActiveObject("Word.Application")
                fresh_word.Interactive = True
                fresh_word.ScreenUpdating = True
                self._word_ui_locked = False
                print("Unlocked Word UI via fresh connection")
            except Exception:
                print("WARNING: Could not unlock Word UI - user may need to restart Word")

    def _capture_initial_document(self):
        """Capture the active document and selection at Win+V time.

        Called from _show_popup BEFORE the popup steals focus, so
        word.ActiveDocument still refers to the document the user was
        working in when they pressed Win+V.  The captured reference is
        later used by _get_word_selection_internal to ensure the result
        goes back to the correct document even if the user switches to
        a different document while typing their prompt.
        """
        try:
            word = self._get_word_app()
            if not word:
                return

            doc = word.ActiveDocument
            if doc:
                self._original_document = doc
                self._original_document_name = doc.Name
                print(f"[Win+V] Pre-captured document at hotkey time: {self._original_document_name}")

                # Also pre-capture the selection range so we know where
                # the cursor/selection was when the hotkey was pressed.
                selection = word.Selection
                if selection:
                    self._original_range_start = selection.Range.Start
                    self._original_range_end = selection.Range.End
                    text = selection.Text.strip() if selection.Text else ""
                    self._original_has_selection = bool(text)
                    self._original_text = text
                    self._original_text_raw = selection.Text if selection.Text else ""
                    print(f"[Win+V] Pre-captured selection: range={self._original_range_start}-{self._original_range_end}, has_selection={self._original_has_selection}")
        except Exception as e:
            print(f"[Win+V] Could not pre-capture document: {e}")

    def _get_word_selection_internal(self, word=None) -> tuple:
        """Get selected text from Word. Returns (text, has_selection)."""
        try:
            if word is None:
                word = self._get_word_app()

            if not word:
                print("Could not connect to Word")
                return "", False

            # If we pre-captured the document at Win+V time, activate it now
            # so that word.Selection refers to the correct document's selection.
            # This is the key fix for the "output goes to wrong document" bug:
            # the user may have switched to a different document while typing
            # their prompt, so word.ActiveDocument would be wrong.
            if self._original_document is not None:
                try:
                    _ = self._original_document.Name  # test if COM ref alive
                    self._original_document.Activate()
                    print(f"Activated pre-captured document: {self._original_document_name}")
                except Exception as e:
                    print(f"Pre-captured document COM reference stale: {e}")
                    self._original_document = None
                    self._original_document_name = None

            # Try to get selection - this might work even if Documents.Count is 0
            selection = None
            try:
                selection = word.Selection
            except Exception as e:
                print(f"Could not get Selection object: {e}")

            if not selection:
                print("Selection object is None")
                return "", False

            # Capture the document for later use (if not already pre-captured)
            if self._original_document is None:
                try:
                    self._original_document = word.ActiveDocument
                    self._original_document_name = self._original_document.Name if self._original_document else None
                    print(f"Captured document (fallback): {self._original_document_name or 'None'}")
                except Exception as e:
                    print(f"Could not capture document: {e}")
                    self._original_document = None
                    self._original_document_name = None

            text = ""
            raw_text_for_redline = ""
            try:
                raw_text = selection.Text
                raw_text_for_redline = raw_text if raw_text else ""
                text = raw_text.strip() if raw_text else ""
                if text:
                    print(f"Got selection text: '{text[:50]}...' ({len(text)} chars)")
                    print(f"Raw text repr (first 100): {repr(raw_text_for_redline[:100])}")
                else:
                    print("No text selected (cursor at insertion point)")
            except Exception as e:
                print(f"Error getting selection text: {e}")
                return "", False

            # Always capture range coordinates so we can restore context later
            # (even if user switches documents while LLM is running)
            try:
                self._original_range_start = selection.Range.Start
                self._original_range_end = selection.Range.End
                print(f"Captured range: Start={self._original_range_start}, End={self._original_range_end}")
            except Exception as e:
                print(f"Could not capture range coordinates: {e}")
                self._original_range_start = None
                self._original_range_end = None

            # Store original text and range for potential redlining
            if getattr(self, '_exec_redline_checked', self.redline_checkbox.isChecked()):
                # Validate redline prerequisites
                is_valid, error_msg = self._validate_redline_prerequisites(text)
                if not is_valid:
                    # Disable checkbox and show error in tooltip
                    self.redline_checkbox.setEnabled(False)
                    self.redline_checkbox.setToolTip(error_msg)
                    self._redline_mode_active = False
                    print(f"Redline validation failed: {error_msg}")
                else:
                    # Validation passed - proceed with redline setup
                    # Re-enable checkbox if it was disabled and restore tooltip
                    self.redline_checkbox.setEnabled(True)
                    self.redline_checkbox.setToolTip(
                        "Instead of replacing text, insert AI suggestions as Track Changes "
                        "that you can accept/reject in Word"
                    )
                    # Store RAW text (not stripped) for redline diffing —
                    # positions must match the Word range exactly
                    self._original_text_raw = raw_text_for_redline
                    self._original_text = text
                    try:
                        # Store the document and range for later use
                        self._original_document = word.ActiveDocument
                        self._original_range_start = selection.Range.Start
                        self._original_range_end = selection.Range.End
                        self._redline_mode_active = True
                        print(f"Captured document and range for redlining: Start={self._original_range_start}, End={self._original_range_end}")
                    except Exception as e:
                        print(f"Could not capture range coordinates: {e}")
                        self._redline_mode_active = False
            else:
                self._redline_mode_active = False

            # wdSelectionIP = 1 (insertion point, no selection)
            has_selection = False
            try:
                sel_type = selection.Type
                has_selection = sel_type != 1 and len(text) > 0
                print(f"Selection type: {sel_type}, has_selection: {has_selection}")
            except:
                has_selection = len(text) > 0

            self._original_has_selection = has_selection
            return text, has_selection
        except Exception as e:
            print(f"Could not get Word selection: {e}")
            return "", False

    def _get_all_document_text(self) -> str:
        """Get all text from the active Word document."""
        try:
            word = self._get_word_app()
            if not word:
                return ""

            # Try to get the active document's content
            try:
                doc = word.ActiveDocument
                if doc:
                    # Get all text from the document
                    content = doc.Content.Text
                    return content.strip() if content else ""
            except Exception as e:
                print(f"Error getting document content: {e}")

            return ""
        except Exception as e:
            print(f"Could not get all document text: {e}")
            return ""

    def _validate_redline_prerequisites(self, selection_text: str) -> tuple[bool, str]:
        """Check if redline mode can be used.

        Args:
            selection_text: Currently selected text

        Returns:
            (is_valid, error_message) tuple
        """
        # Check if text is selected
        if not selection_text or len(selection_text.strip()) == 0:
            return False, "Select text first to use redline mode"

        # Check text length limit
        max_length = self.redline_settings.get("max_redline_text_length", 50000)
        if len(selection_text) > max_length:
            return False, f"Selection too large for redline ({len(selection_text)} > {max_length} chars)"

        return True, ""

    def _apply_redlines(self, word_app, selection, original_text: str, revised_text: str) -> bool:
        """Apply redlines using flat diff mapped back to Word paragraphs.

        Delegates to the standalone apply_flat_diff_redline() function,
        then runs validation.

        Args:
            word_app: Word COM application object
            selection: Word Selection object
            original_text: Original text before LLM processing
            revised_text: LLM-revised text

        Returns:
            True if redlines applied successfully, False if fallback needed
        """
        try:
            # Get document object
            try:
                doc = self._original_document if self._original_document else word_app.ActiveDocument
                if not doc:
                    print("No document available")
                    return False
            except Exception as e:
                print(f"Could not access Word document: {e}")
                return False

            auto_tc = self.redline_settings.get("auto_enable_track_changes", True)

            # Delegate to standalone function
            r = apply_flat_diff_redline(
                doc_com=doc,
                range_start=self._original_range_start,
                range_end=self._original_range_end,
                original_text=original_text,
                revised_text=revised_text,
                auto_enable_track_changes=auto_tc,
            )

            if not r['success']:
                return False

            # Post-application validation
            import time as _time
            _time.sleep(0.5)
            try:
                from icharlotte_core.word_validator import validate_redline
                val_result = validate_redline(
                    doc_com=doc,
                    range_start=self._original_range_start,
                    range_end=self._original_range_end,
                    pre_content_para_count=r['pre_content_para_count'],
                    orig_flat_length=r['orig_flat_length'],
                    deleted_chars=r['deleted_chars'],
                    inserted_chars=r['inserted_chars'],
                )
                val_result.print_summary()
            except Exception as ve:
                print(f"Validation error: {ve}")

            return True

        except Exception as e:
            print(f"Unexpected error in _apply_redlines: {e}")
            import traceback
            traceback.print_exc()
            return False

    def _set_word_text_internal(self, text: str, format_type: str = FORMAT_PLAIN,
                               word=None, selection=None):
        """Set text in Word with formatting.

        Args:
            text: Text to insert
            format_type: Formatting mode constant
            word: Pre-resolved Word COM app (optional, reconnects if None)
            selection: Pre-resolved Selection (optional, gets from word if None)
        """
        try:
            if not word:
                word = self._get_word_app()
            if not word:
                raise Exception("Could not connect to Word")

            if not selection:
                selection = word.Selection
            if not selection:
                raise Exception("Could not access Word selection")

            # Use stored document if available, otherwise fall back to ActiveDocument
            doc = self._original_document if self._original_document else word.ActiveDocument

            # Check if Track Changes is enabled and preserve it for all format modes
            track_changes_active = False
            try:
                track_changes_active = doc.TrackRevisions if doc else False
                if track_changes_active:
                    print("Track Changes is active - changes will be tracked")
            except:
                pass

            print(f"[DEBUG] _set_word_text_internal: format_type='{format_type}', text[:200]='{text[:200]}'")

            if format_type == FORMAT_PLAIN:
                # Insert text, but still handle bullet points properly
                self._insert_with_bullets(word, selection, text)

            elif format_type == FORMAT_MATCH:
                # Apply captured format (includes Track Changes handling)
                self._insert_with_format(selection, text, self._captured_format)

            elif format_type == FORMAT_MARKDOWN:
                # Parse markdown and apply formatting
                self._insert_with_markdown(word, selection, text)

            elif format_type == FORMAT_DEFAULT:
                # Reset to Normal style, then insert
                selection.Style = doc.Styles("Normal")
                selection.TypeText(text)

            elif format_type == FORMAT_CUSTOM:
                # Apply custom format settings with Track Changes awareness
                custom_settings = dict(self.custom_format_settings) if self.custom_format_settings else {}
                custom_settings['track_changes_enabled'] = track_changes_active
                self._insert_with_format(selection, text, custom_settings)

            else:
                # Fallback to plain text
                selection.TypeText(text)

            print(f"Successfully inserted {len(text)} characters into Word with format: {format_type}")

        except Exception as e:
            print(f"_set_word_text_internal error: {e}")
            raise

    def _insert_with_format(self, selection, text: str, format_settings: dict):
        """Insert text with specific formatting.

        Enhanced to apply:
        - Paragraph styles (Heading 1, Normal, Body Text, etc.)
        - Table cell formatting when inside tables
        - Track Changes support (applies changes as tracked edits)
        - Inline font formatting
        """
        if not format_settings:
            selection.TypeText(text)
            return

        try:
            word = selection.Application
            doc = self._original_document if self._original_document else word.ActiveDocument

            # Handle Track Changes - ensure revisions are tracked if originally enabled
            track_changes_was_enabled = format_settings.get('track_changes_enabled', False)
            original_track_state = None
            if track_changes_was_enabled:
                try:
                    original_track_state = doc.TrackRevisions
                    # Ensure Track Changes is on so our edits are tracked
                    if not doc.TrackRevisions:
                        doc.TrackRevisions = True
                        print("Temporarily enabled Track Changes for this edit")
                except Exception as e:
                    print(f"Could not set Track Changes: {e}")

            # Apply paragraph style first (if captured)
            style_name = format_settings.get('style_name')
            if style_name:
                try:
                    # Try to apply the style by name
                    selection.Style = doc.Styles(style_name)
                    print(f"Applied paragraph style: {style_name}")
                except Exception as e:
                    print(f"Could not apply style '{style_name}': {e}")
                    # Style might not exist - continue with other formatting

            font = selection.Font
            para = selection.ParagraphFormat

            # Apply font settings
            if format_settings.get('font_name'):
                font.Name = format_settings['font_name']
            if format_settings.get('font_size'):
                font.Size = format_settings['font_size']

            # Apply font color if captured
            if 'font_color' in format_settings:
                try:
                    font.Color = format_settings['font_color']
                except:
                    pass

            # Apply style (True = -1 in Word COM, False = 0)
            font.Bold = -1 if format_settings.get('bold') else 0
            font.Italic = -1 if format_settings.get('italic') else 0

            # Underline: wdUnderlineSingle = 1, wdUnderlineNone = 0
            font.Underline = 1 if format_settings.get('underline') else 0

            # Apply paragraph settings (convert inches to points: 1 inch = 72 points)
            if 'first_indent' in format_settings:
                para.FirstLineIndent = format_settings['first_indent'] * 72
            if 'left_indent' in format_settings:
                para.LeftIndent = format_settings['left_indent'] * 72

            # Line spacing - handle both simple string values and captured rule values
            line_spacing = format_settings.get('line_spacing', None)
            line_spacing_rule = format_settings.get('line_spacing_rule', None)

            if line_spacing_rule is not None:
                try:
                    para.LineSpacingRule = line_spacing_rule
                    # If it's a specific spacing value, set that too
                    if line_spacing_rule in [3, 4, 5] and 'line_spacing' in format_settings:
                        para.LineSpacing = format_settings['line_spacing']
                except:
                    pass
            elif line_spacing:
                if line_spacing == 'Single':
                    para.LineSpacingRule = 0  # wdLineSpaceSingle
                elif line_spacing == '1.5 Lines':
                    para.LineSpacingRule = 1  # wdLineSpace1pt5
                elif line_spacing == 'Double':
                    para.LineSpacingRule = 2  # wdLineSpaceDouble

            # Insert the text
            selection.TypeText(text)

            # Handle table cell formatting after text insertion
            if format_settings.get('in_table'):
                try:
                    # Re-check if we're still in a table after insertion
                    if selection.Information(12):  # wdWithInTable
                        cell = selection.Cells(1) if selection.Cells.Count > 0 else None
                        if cell:
                            # Apply cell shading if it was captured
                            if 'cell_shading_color' in format_settings:
                                cell.Shading.BackgroundPatternColor = format_settings['cell_shading_color']
                            if 'cell_shading_pattern' in format_settings:
                                cell.Shading.Texture = format_settings['cell_shading_pattern']

                            # Apply vertical alignment
                            if 'cell_vertical_align' in format_settings:
                                cell.VerticalAlignment = format_settings['cell_vertical_align']

                            print("Applied table cell formatting")
                except Exception as e:
                    print(f"Could not apply table cell formatting: {e}")

            # Restore original Track Changes state if we modified it
            if original_track_state is not None and not original_track_state:
                try:
                    doc.TrackRevisions = False
                    print("Restored Track Changes to original state (off)")
                except:
                    pass

        except Exception as e:
            print(f"Error applying format: {e}")
            # Fallback to plain text
            selection.TypeText(text)

    def _is_in_table(self, selection) -> bool:
        """Check if the current selection is inside a table cell."""
        try:
            return selection.Information(12)  # wdWithInTable = 12
        except:
            return False

    def _insert_with_bullets(self, word, selection, text: str):
        """Insert text with proper Word bullet formatting (no inline markdown).

        Enhanced to handle table cell context - bullets work differently in tables.
        """
        # Check if we're in a table - bullet handling may differ
        in_table = self._is_in_table(selection)
        if in_table:
            print("Inserting bullets within table cell")

        # First pass: group lines into items (bullet items include their continuations)
        items = self._parse_bullet_items(text)

        in_bullet_list = False
        for i, item in enumerate(items):
            if item['is_bullet']:
                if not in_bullet_list:
                    # First bullet - apply formatting and set up list template
                    in_bullet_list = True

                    # Apply Word's bullet formatting (enables auto-continuation on Enter)
                    try:
                        selection.Range.ListFormat.ApplyBulletDefault()

                        # Modify the ListTemplate level settings for correct indentation
                        # This is what actually controls bullet/text positioning in Word lists
                        # In tables, we use smaller indents to fit within cells
                        list_fmt = selection.Range.ListFormat
                        if list_fmt.ListTemplate and list_fmt.ListLevelNumber > 0:
                            level = list_fmt.ListTemplate.ListLevels(list_fmt.ListLevelNumber)
                            if in_table:
                                # Smaller indents for table cells
                                level.NumberPosition = 18  # 0.25 inch
                                level.TextPosition = 36    # 0.5 inch
                                level.TabPosition = 36     # 0.5 inch
                            else:
                                level.NumberPosition = 36  # 0.5 inch - where bullet sits
                                level.TextPosition = 54    # 0.75 inch - where text starts
                                level.TabPosition = 54     # 0.75 inch - tab stop
                    except Exception as e:
                        print(f"Could not modify list template: {e}")

                # For subsequent bullets, Word's list continuation handles formatting
                selection.TypeText(item['text'])

                if i < len(items) - 1:
                    selection.TypeParagraph()
            else:
                # Non-bullet item
                if in_bullet_list:
                    in_bullet_list = False
                    try:
                        selection.Range.ListFormat.RemoveNumbers()
                        para = selection.ParagraphFormat
                        para.LeftIndent = 0
                        para.FirstLineIndent = 0
                    except Exception as e:
                        print(f"Could not remove bullet formatting: {e}")

                if item['text']:
                    selection.TypeText(item['text'])

                if i < len(items) - 1:
                    selection.TypeParagraph()

    def _parse_markdown_items(self, text: str) -> list:
        """Parse markdown text into structured items: headings, bullets, and body text.

        Returns list of dicts with keys:
            type: 'heading', 'bullet', 'body'
            text: content (markdown markers stripped)
            level: heading level (1-3) for headings only
        """
        lines = text.split('\n')
        bullet_pattern = re.compile(r'^[\s]*[-*•]\s+(.*)$')
        numbered_pattern = re.compile(r'^[\s]*\d+[.)]\s+(.*)$')
        heading_pattern = re.compile(r'^(#{1,3})\s+(.+)$')

        items = []
        current_bullet = None

        for line in lines:
            heading_match = heading_pattern.match(line)
            bullet_match = bullet_pattern.match(line)
            numbered_match = numbered_pattern.match(line)

            if heading_match:
                # Flush pending bullet
                if current_bullet is not None:
                    items.append({'type': 'bullet', 'text': current_bullet})
                    current_bullet = None
                level = len(heading_match.group(1))
                items.append({'type': 'heading', 'text': heading_match.group(2).strip(), 'level': level})

            elif bullet_match or numbered_match:
                # Save previous bullet if exists
                if current_bullet is not None:
                    items.append({'type': 'bullet', 'text': current_bullet})

                # Start new bullet
                current_bullet = bullet_match.group(1) if bullet_match else numbered_match.group(1)

            elif current_bullet is not None and line.strip():
                # Continuation of current bullet - append with space
                current_bullet += " " + line.strip()

            else:
                # Non-bullet line or empty line
                if current_bullet is not None:
                    items.append({'type': 'bullet', 'text': current_bullet})
                    current_bullet = None

                # Add body item (even if empty, to preserve paragraph breaks)
                if line.strip() or (items and items[-1].get('type') == 'bullet'):
                    items.append({'type': 'body', 'text': line.strip()})

        # Don't forget the last bullet
        if current_bullet is not None:
            items.append({'type': 'bullet', 'text': current_bullet})

        return items

    def _parse_bullet_items(self, text: str) -> list:
        """Parse text into items, combining bullet lines with their continuations.

        Legacy wrapper - converts new format to old format for _insert_with_bullets.
        """
        md_items = self._parse_markdown_items(text)
        items = []
        for item in md_items:
            if item['type'] == 'bullet':
                items.append({'is_bullet': True, 'text': item['text']})
            else:
                items.append({'is_bullet': False, 'text': item['text']})
        return items

    def _insert_with_markdown(self, word, selection, text: str):
        """Parse markdown and insert with Word formatting.

        Handles headings, bullet lists, inline bold/italic/code, and body text.
        Heading formatting follows report generator conventions:
            # Heading  → Bold, ALL CAPS (section header)
            ## Heading → Bold + Underline (L1 subheading)
            ### Heading → Bold (L2 subheading)
        """
        # Check if we're in a table
        in_table = self._is_in_table(selection)
        if in_table:
            print("Inserting markdown within table cell")

        items = self._parse_markdown_items(text)

        in_bullet_list = False
        for i, item in enumerate(items):
            item_type = item['type']

            # --- End bullet list if switching to non-bullet ---
            if item_type != 'bullet' and in_bullet_list:
                in_bullet_list = False
                try:
                    selection.Range.ListFormat.RemoveNumbers()
                    para = selection.ParagraphFormat
                    para.LeftIndent = 0
                    para.FirstLineIndent = 0
                except Exception as e:
                    print(f"Could not remove bullet formatting: {e}")

            if item_type == 'heading':
                level = item.get('level', 1)
                heading_text = item['text']

                # Reset paragraph formatting for heading
                try:
                    para = selection.ParagraphFormat
                    para.LeftIndent = 0
                    para.FirstLineIndent = 0
                except:
                    pass

                font = selection.Font
                font.Bold = -1  # All headings are bold
                font.Italic = 0

                if level == 1:
                    # Section header: bold, all caps
                    font.AllCaps = -1
                    font.Underline = 0
                    heading_text = heading_text.upper()
                elif level == 2:
                    # L1 subheading: bold + underline
                    font.AllCaps = 0
                    font.Underline = 1
                else:
                    # L2 subheading: bold only
                    font.AllCaps = 0
                    font.Underline = 0

                # Strip any inline markdown from heading text
                clean_text = re.sub(r'\*\*(.+?)\*\*', r'\1', heading_text)
                clean_text = re.sub(r'__(.+?)__', r'\1', clean_text)
                selection.TypeText(clean_text)

                # Reset font after heading
                font.Bold = 0
                font.AllCaps = 0
                font.Underline = 0

                if i < len(items) - 1:
                    selection.TypeParagraph()

            elif item_type == 'bullet':
                if not in_bullet_list:
                    in_bullet_list = True
                    try:
                        selection.Range.ListFormat.ApplyBulletDefault()
                        list_fmt = selection.Range.ListFormat
                        if list_fmt.ListTemplate and list_fmt.ListLevelNumber > 0:
                            level = list_fmt.ListTemplate.ListLevels(list_fmt.ListLevelNumber)
                            if in_table:
                                level.NumberPosition = 18
                                level.TextPosition = 36
                                level.TabPosition = 36
                            else:
                                level.NumberPosition = 36
                                level.TextPosition = 54
                                level.TabPosition = 54
                    except Exception as e:
                        print(f"Could not modify list template: {e}")

                self._insert_formatted_text(selection, item['text'])

                if i < len(items) - 1:
                    selection.TypeParagraph()

            else:
                # Body text
                if item['text']:
                    self._insert_formatted_text(selection, item['text'])

                if i < len(items) - 1:
                    selection.TypeParagraph()

    def _insert_formatted_text(self, selection, text: str):
        """Insert text with inline markdown formatting (bold, italic, etc.)."""
        segments = self._parse_markdown_segments(text)

        for segment in segments:
            seg_text = segment['text']
            if not seg_text:
                continue

            # Set formatting based on segment type
            font = selection.Font

            # Reset to defaults first
            font.Bold = 0
            font.Italic = 0
            font.Underline = 0

            if segment.get('bold'):
                font.Bold = -1
            if segment.get('italic'):
                font.Italic = -1
            if segment.get('underline'):
                font.Underline = 1
            if segment.get('code'):
                # Use monospace font for code
                font.Name = "Consolas"

            selection.TypeText(seg_text)

            # Reset font name if we changed it for code
            if segment.get('code'):
                font.Name = "Times New Roman"

    def _parse_markdown_segments(self, text: str) -> list:
        """Parse markdown text into segments with formatting info."""
        segments = []
        current_pos = 0

        # Combined pattern for markdown formatting
        # Order matters: check ** before *, __ before _
        pattern = r'(\*\*(.+?)\*\*|__(.+?)__|(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)|(?<!_)_(?!_)(.+?)(?<!_)_(?!_)|~~(.+?)~~|`([^`]+)`)'

        for match in re.finditer(pattern, text):
            start, end = match.span()

            # Add plain text before this match
            if start > current_pos:
                segments.append({
                    'text': text[current_pos:start],
                    'bold': False, 'italic': False, 'underline': False, 'code': False
                })

            # Determine which group matched
            groups = match.groups()
            if groups[1]:  # **bold**
                segments.append({'text': groups[1], 'bold': True, 'italic': False, 'underline': False, 'code': False})
            elif groups[2]:  # __bold__
                segments.append({'text': groups[2], 'bold': True, 'italic': False, 'underline': False, 'code': False})
            elif groups[3]:  # *italic*
                segments.append({'text': groups[3], 'bold': False, 'italic': True, 'underline': False, 'code': False})
            elif groups[4]:  # _underline_ (we'll use single underscore for underline)
                segments.append({'text': groups[4], 'bold': False, 'italic': False, 'underline': True, 'code': False})
            elif groups[5]:  # ~~strikethrough~~ - we'll treat as italic for now
                segments.append({'text': groups[5], 'bold': False, 'italic': True, 'underline': False, 'code': False})
            elif groups[6]:  # `code`
                segments.append({'text': groups[6], 'bold': False, 'italic': False, 'underline': False, 'code': True})

            current_pos = end

        # Add remaining plain text
        if current_pos < len(text):
            segments.append({
                'text': text[current_pos:],
                'bold': False, 'italic': False, 'underline': False, 'code': False
            })

        return segments if segments else [{'text': text, 'bold': False, 'italic': False, 'underline': False, 'code': False}]

    def _is_word_running(self) -> bool:
        """Check if Word is running."""
        if not HAS_WIN32:
            return False
        try:
            pythoncom.CoInitialize()
            word = win32com.client.GetActiveObject("Word.Application")
            return word is not None
        except:
            return False

    def _get_word_selection(self) -> tuple:
        """Get selected text from Word. Returns (text, has_selection)."""
        if not HAS_WIN32:
            return "", False

        try:
            pythoncom.CoInitialize()
            try:
                word = win32com.client.GetActiveObject("Word.Application")
            except Exception:
                # Word not running
                return "", False

            if not word:
                return "", False

            selection = word.Selection
            if not selection:
                return "", False

            text = ""
            try:
                text = selection.Text.strip() if selection.Text else ""
            except:
                pass

            # Check if there's actually a selection (not just cursor position)
            # wdSelectionIP = 1 (insertion point, no selection)
            has_selection = False
            try:
                has_selection = selection.Type != 1
            except:
                has_selection = len(text) > 0

            return text, has_selection
        except Exception as e:
            print(f"Could not get Word selection: {e}")
            return "", False

    def _set_word_text(self, text: str, replace: bool = True):
        """Set text in Word (replace selection or insert at cursor)."""
        if not HAS_WIN32:
            raise Exception("win32com not available")

        try:
            pythoncom.CoInitialize()
            word = win32com.client.GetActiveObject("Word.Application")

            if not word:
                raise Exception("Could not connect to Word")

            selection = word.Selection
            if not selection:
                raise Exception("Could not get Word selection")

            # TypeText replaces selection if there is one, or inserts at cursor
            selection.TypeText(text)

        except Exception as e:
            print(f"Could not set Word text: {e}")
            raise

    def _default_llm_call(self, prompt: str) -> str:
        """Default LLM call using the app's LLM infrastructure."""
        try:
            from icharlotte_core.llm import LLMHandler
            from icharlotte_core.config import API_KEYS

            # Get the selected model
            provider, model_id = self._get_selected_model()
            print(f"Using model: {provider} / {model_id}")

            # Check API key is available
            api_key = API_KEYS.get(provider)
            if not api_key:
                raise ValueError(f"API key for {provider} not configured. Check your .env file.")
            print(f"API key for {provider}: {'*' * 8}...{api_key[-4:] if len(api_key) > 4 else '****'}")

            # Use LLMHandler.generate with proper parameters
            settings = {
                'temperature': 0.7,
                'top_p': 0.95,
                'max_tokens': -1,
                'stream': False,
                'thinking_level': 'None'
            }

            # Use redline-aware system prompt when redline mode is active
            if self._redline_mode_active:
                system_prompt = (
                    "You are a helpful writing assistant operating in REDLINE MODE. "
                    "Your output will be diffed against the original text to produce "
                    "Track Changes in Microsoft Word. You MUST preserve unchanged text "
                    "EXACTLY as written — same words, same order, same punctuation. "
                    "Only modify the specific parts that need changing per the user's "
                    "instructions. Output only the requested text without any preamble, "
                    "explanation, or markdown formatting unless specifically asked."
                )
            else:
                system_prompt = (
                    "You are a helpful writing assistant for Microsoft Word documents. "
                    "Follow the user's instructions precisely. Output only the requested "
                    "text without any preamble or explanation.\n\n"
                    "FORMAT RULES (strict defaults — override ONLY if the user explicitly "
                    "asks for headings, bullet points, numbered lists, or other structure):\n"
                    "- Always write in plain narrative prose. No headings (#, ##, ###), "
                    "no bullet points, no numbered lists, no markdown formatting.\n"
                    "- Use **bold** sparingly and only when the user asks for emphasis.\n"
                    "- Do not use code blocks or HTML.\n\n"
                    "CONTINUATION RULE: When the user asks you to finish a sentence, "
                    "fill in a blank, or continue from a specific point, output ONLY "
                    "the new text that comes after the existing content. Do NOT repeat "
                    "any part of the preceding sentence, paragraph, or context that was "
                    "provided to you. Start exactly where the existing text leaves off."
                )

            print(f"Calling LLMHandler.generate...")
            result = LLMHandler.generate(
                provider=provider,
                model=model_id,
                system_prompt=system_prompt,
                user_prompt=prompt,
                file_contents="",
                settings=settings
            )
            print(f"LLM result type: {type(result)}, length: {len(result) if result else 0}")

            if not result:
                raise ValueError("LLM returned empty response")

            return result.strip()
        except Exception as e:
            import traceback
            print(f"LLM call failed: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise

    # ========== Outlook-specific methods ==========

    def _do_execute_outlook(self, prompt: str):
        """Execute LLM processing for Outlook email context."""
        try:
            # Verify inspector is still valid
            if not self.active_inspector:
                self.status_label.setText("Outlook compose window not found")
                self.execute_btn.setEnabled(True)
                self.cancel_btn.setText("Cancel")
                QMessageBox.warning(self, "Outlook Not Found",
                    "No active Outlook compose window found.\n\nPlease open an email compose window first, then try again.")
                return

            # Capture format if "Match Selection" is chosen
            format_type = self.format_combo.currentText()
            if format_type == FORMAT_MATCH:
                self._capture_outlook_selection_format()

            # Get selection from Outlook
            selected_text, has_selection = self._get_outlook_selection()

            if has_selection and selected_text:
                full_prompt = f"{prompt}\n\nEmail text to process:\n{selected_text}"
                self.status_label.setText("Processing selected email text...")
            else:
                full_prompt = prompt
                self.status_label.setText("Processing prompt...")

            # Append any attached file context
            attachment_context = self.attachment_area.get_attachment_context()
            if attachment_context:
                full_prompt += attachment_context

            QApplication.processEvents()

            # Store format type for callback
            self._pending_format_type = format_type
            self._is_outlook_task = True

            # Start worker thread for LLM call (using email-specific function)
            self._worker_thread = LLMWorkerThread(self._call_llm_for_email, full_prompt, self)
            # Store redline state on thread for result handler (not used for Outlook)
            self._worker_thread.redline_mode = False  # Redline not supported for Outlook
            self._worker_thread.finished.connect(self._on_llm_result)
            self._worker_thread.start()

        except Exception as e:
            self.status_label.setText(f"Error: {str(e)[:50]}")
            self.execute_btn.setEnabled(True)
            self.cancel_btn.setText("Cancel")
            QMessageBox.critical(self, "Error", f"Failed to process:\n{e}")

    def _get_outlook_selection(self) -> tuple:
        """Get selected text from Outlook compose window. Returns (text, has_selection)."""
        if not self.active_inspector:
            return "", False

        try:
            word_editor = self.active_inspector.WordEditor
            if not word_editor:
                print("WordEditor not available in Outlook Inspector")
                return "", False

            selection = word_editor.Application.Selection
            if not selection:
                return "", False

            text = ""
            try:
                raw_text = selection.Text
                text = raw_text.strip() if raw_text else ""
                if text:
                    print(f"Got Outlook selection text: '{text[:50]}...' ({len(text)} chars)")
                else:
                    print("No text selected in Outlook (cursor at insertion point)")
            except Exception as e:
                print(f"Error getting Outlook selection text: {e}")
                return "", False

            # wdSelectionIP = 1 (insertion point, no selection)
            has_selection = False
            try:
                sel_type = selection.Type
                has_selection = sel_type != 1 and len(text) > 0
                print(f"Outlook selection type: {sel_type}, has_selection: {has_selection}")
            except:
                has_selection = len(text) > 0

            return text, has_selection
        except Exception as e:
            print(f"Could not get Outlook selection: {e}")
            return "", False

    def _set_outlook_text(self, text: str, format_type: str = FORMAT_PLAIN):
        """Insert text into Outlook compose window with formatting."""
        if not self.active_inspector:
            raise Exception("No active Outlook inspector")

        try:
            word_editor = self.active_inspector.WordEditor
            if not word_editor:
                raise Exception("Could not access Outlook WordEditor")

            selection = word_editor.Application.Selection
            if not selection:
                raise Exception("Could not get selection in Outlook")

            # Use same formatting logic as Word since WordEditor provides full Word DOM access
            if format_type == FORMAT_PLAIN:
                self._insert_with_bullets(word_editor.Application, selection, text)

            elif format_type == FORMAT_MATCH:
                self._insert_with_format(selection, text, self._captured_format)

            elif format_type == FORMAT_MARKDOWN:
                self._insert_with_markdown(word_editor.Application, selection, text)

            elif format_type == FORMAT_DEFAULT:
                # For email, just insert plain text
                selection.TypeText(text)

            elif format_type == FORMAT_CUSTOM:
                self._insert_with_format(selection, text, self.custom_format_settings)

            else:
                selection.TypeText(text)

            print(f"Successfully inserted {len(text)} characters into Outlook with format: {format_type}")

        except Exception as e:
            print(f"_set_outlook_text error: {e}")
            raise

    def _capture_outlook_selection_format(self):
        """Capture the formatting of the current selection in Outlook compose window.

        Enhanced to capture paragraph styles and table context in email composition.
        """
        if not self.active_inspector:
            self._captured_format = None
            return

        try:
            word_editor = self.active_inspector.WordEditor
            if not word_editor:
                self._captured_format = None
                return

            selection = word_editor.Application.Selection
            font = selection.Font
            para = selection.ParagraphFormat

            underline_val = font.Underline
            has_underline = underline_val != 0 and underline_val != 9999999

            self._captured_format = {
                'font_name': font.Name if font.Name != '' else None,
                'font_size': font.Size if font.Size != 9999999 else 11,  # Default 11 for email
                'bold': font.Bold == -1,
                'italic': font.Italic == -1,
                'underline': has_underline,
                'first_indent': para.FirstLineIndent / 72 if para.FirstLineIndent != 9999999 else 0,
                'left_indent': para.LeftIndent / 72 if para.LeftIndent != 9999999 else 0,
            }

            # Capture paragraph style name
            try:
                style = selection.Style
                if style:
                    style_name = style.NameLocal
                    self._captured_format['style_name'] = style_name
                    print(f"Captured Outlook paragraph style: {style_name}")
            except Exception as e:
                print(f"Could not capture Outlook style: {e}")
                self._captured_format['style_name'] = None

            # Detect if selection is inside a table (emails can have tables)
            try:
                if selection.Information(12):  # wdWithInTable
                    self._captured_format['in_table'] = True
                    try:
                        cell = selection.Cells(1) if selection.Cells.Count > 0 else None
                        if cell:
                            shading = cell.Shading
                            if shading:
                                self._captured_format['cell_shading_color'] = shading.BackgroundPatternColor
                                self._captured_format['cell_shading_pattern'] = shading.Texture
                            self._captured_format['cell_vertical_align'] = cell.VerticalAlignment
                            print("Captured Outlook table cell formatting")
                    except:
                        pass
                else:
                    self._captured_format['in_table'] = False
            except:
                self._captured_format['in_table'] = False

            # Capture font color
            try:
                font_color = font.Color
                if font_color != 9999999:
                    self._captured_format['font_color'] = font_color
            except:
                pass

            # Capture line spacing
            try:
                line_spacing_rule = para.LineSpacingRule
                self._captured_format['line_spacing_rule'] = line_spacing_rule
                if line_spacing_rule in [3, 4, 5]:
                    self._captured_format['line_spacing'] = para.LineSpacing
            except:
                pass

            # No Track Changes in Outlook email composition
            self._captured_format['track_changes_enabled'] = False

            print(f"Captured Outlook format: {self._captured_format}")
        except Exception as e:
            print(f"Error capturing Outlook format: {e}")
            self._captured_format = None

    def _call_llm_for_email(self, prompt: str) -> str:
        """LLM call with email-specific system prompt."""
        try:
            from icharlotte_core.llm import LLMHandler
            from icharlotte_core.config import API_KEYS

            provider, model_id = self._get_selected_model()
            print(f"Using model for email: {provider} / {model_id}")

            # Check API key is available
            api_key = API_KEYS.get(provider)
            if not api_key:
                raise ValueError(f"API key for {provider} not configured. Check your .env file.")

            settings = {
                'temperature': 0.7,
                'top_p': 0.95,
                'max_tokens': -1,
                'stream': False,
                'thinking_level': 'None'
            }

            print(f"Calling LLMHandler.generate for email...")
            result = LLMHandler.generate(
                provider=provider,
                model=model_id,
                system_prompt=EMAIL_SYSTEM_PROMPT,
                user_prompt=prompt,
                file_contents="",
                settings=settings
            )
            print(f"Email LLM result type: {type(result)}, length: {len(result) if result else 0}")

            if not result:
                raise ValueError("LLM returned empty response")

            return result.strip()
        except Exception as e:
            import traceback
            print(f"Email LLM call failed: {e}")
            print(f"Traceback: {traceback.format_exc()}")
            raise

    # ========== End Outlook-specific methods ==========

    def showEvent(self, event):
        """Position the dialog when shown."""
        super().showEvent(event)

        if self.cursor_pos:
            # Position at cursor location
            x = self.cursor_pos.x()
            y = self.cursor_pos.y()

            # Get the screen that contains the cursor
            screen = QApplication.screenAt(self.cursor_pos)
            if screen:
                screen_geo = screen.availableGeometry()
            else:
                screen_geo = QApplication.primaryScreen().availableGeometry()

            # Ensure the popup stays within screen bounds
            # Adjust if popup would go off the right edge
            if x + self.width() > screen_geo.right():
                x = screen_geo.right() - self.width()
            # Adjust if popup would go off the bottom edge
            if y + self.height() > screen_geo.bottom():
                y = screen_geo.bottom() - self.height()
            # Ensure not off left or top edge
            if x < screen_geo.left():
                x = screen_geo.left()
            if y < screen_geo.top():
                y = screen_geo.top()

            self.move(x, y)
        else:
            # Fallback: center on screen
            screen = QApplication.primaryScreen().geometry()
            self.move(
                (screen.width() - self.width()) // 2,
                (screen.height() - self.height()) // 3
            )

        # Focus the custom input
        self.custom_input.setFocus()

    def closeEvent(self, event):
        """Clean up worker thread and ensure Word UI is unlocked."""
        if self._worker_thread and self._worker_thread.isRunning():
            self._worker_thread.cancel()
            self._worker_thread.terminate()
            self._worker_thread.wait(500)
            self._worker_thread = None
        # Safety: ensure Word UI is never left locked (use direct COM, skip zombie check)
        if getattr(self, '_word_ui_locked', False):
            try:
                word = win32com.client.GetActiveObject("Word.Application")
                if word:
                    self._unlock_word_ui(word)
            except Exception:
                pass
        super().closeEvent(event)

    def keyPressEvent(self, event):
        """Handle key presses."""
        if event.key() == Qt.Key.Key_Escape:
            self._on_cancel_clicked()
        else:
            super().keyPressEvent(event)

    def mousePressEvent(self, event):
        """Start dragging if clicking on header area."""
        if event.button() == Qt.MouseButton.LeftButton:
            # Check if click is in the header/title area (top 50 pixels)
            if event.position().y() <= 50:
                self._is_dragging = True
                self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        """Move window while dragging."""
        if self._is_dragging and self._drag_pos is not None:
            new_pos = event.globalPosition().toPoint() - self._drag_pos
            self.move(new_pos)
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        """Stop dragging."""
        if event.button() == Qt.MouseButton.LeftButton and self._is_dragging:
            self._is_dragging = False
            self._drag_pos = None
            event.accept()
            return
        super().mouseReleaseEvent(event)


# Bind _InsertionProxy methods from WordLLMPopup (must happen after class is defined)
_InsertionProxy._set_word_text_internal = WordLLMPopup._set_word_text_internal
_InsertionProxy._insert_with_format = WordLLMPopup._insert_with_format
_InsertionProxy._insert_with_markdown = WordLLMPopup._insert_with_markdown
_InsertionProxy._insert_with_bullets = WordLLMPopup._insert_with_bullets
_InsertionProxy._insert_formatted_text = WordLLMPopup._insert_formatted_text
_InsertionProxy._parse_markdown_items = WordLLMPopup._parse_markdown_items
_InsertionProxy._parse_bullet_items = WordLLMPopup._parse_bullet_items
_InsertionProxy._parse_markdown_segments = WordLLMPopup._parse_markdown_segments
_InsertionProxy._is_in_table = WordLLMPopup._is_in_table


class WordHotkeyManager:
    """Manages global hotkey registration and popup display."""

    def __init__(self, main_window=None):
        self.main_window = main_window
        self.popup = None
        self.signals = HotkeySignals()
        self.signals.show_popup.connect(self._show_popup)
        self._hotkey_registered = False
        self._hotkey_thread = None
        self._hotkey_thread_id = None  # Win32 thread ID for posting quit message
        self._stop_event = threading.Event()
        self._watchdog_timer = None

    def start(self):
        """Start listening for the global hotkey using Windows RegisterHotKey API."""
        # Check if already registered AND thread is still alive
        if self._hotkey_registered:
            if self._hotkey_thread is not None and self._hotkey_thread.is_alive():
                return True
            # Thread died — reset and re-register
            print("Win+V hotkey thread died, re-registering...")
            self._hotkey_registered = False

        self._stop_event.clear()
        self._hotkey_thread = threading.Thread(target=self._hotkey_message_loop, daemon=True)
        self._hotkey_thread.start()

        # Wait briefly for registration to complete
        import time
        time.sleep(0.1)

        # Start watchdog timer to auto-restart if thread dies
        self._start_watchdog()

        return self._hotkey_registered

    def _hotkey_message_loop(self):
        """Dedicated thread running a Windows message pump for RegisterHotKey."""
        try:
            # Store this thread's ID so we can post WM_QUIT to it later
            self._hotkey_thread_id = ctypes.windll.kernel32.GetCurrentThreadId()

            # Register the hotkey on THIS thread
            result = ctypes.windll.user32.RegisterHotKey(
                None, _HOTKEY_ID, _MOD_WIN | _MOD_NOREPEAT, _VK_V
            )
            if not result:
                err = ctypes.GetLastError()
                print(f"RegisterHotKey failed (error {err}) - Win+V may be in use by another app")
                # Fall back to keyboard library
                self._fallback_to_keyboard_lib()
                return

            self._hotkey_registered = True
            print("Win+V hotkey registered via RegisterHotKey API")

            # Message pump - blocks until WM_QUIT or stop event
            msg = wintypes.MSG()
            while not self._stop_event.is_set():
                # Use GetMessage (blocks until a message arrives)
                ret = ctypes.windll.user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
                if ret <= 0:  # WM_QUIT or error
                    break
                if msg.message == _WM_HOTKEY and msg.wParam == _HOTKEY_ID:
                    import time as _t
                    logger.debug(f"Win+V hotkey detected at {_t.strftime('%H:%M:%S')}, emitting show_popup signal")
                    self.signals.show_popup.emit()

            # Cleanup
            ctypes.windll.user32.UnregisterHotKey(None, _HOTKEY_ID)
            self._hotkey_registered = False
            print("Win+V hotkey unregistered")
        except Exception as e:
            print(f"Win+V hotkey thread crashed: {e}")
            try:
                ctypes.windll.user32.UnregisterHotKey(None, _HOTKEY_ID)
            except Exception:
                pass
            self._hotkey_registered = False

    def _start_watchdog(self):
        """Start a periodic timer that checks if the hotkey thread is alive."""
        if self._watchdog_timer is not None:
            return
        from PySide6.QtCore import QTimer
        self._watchdog_timer = QTimer()
        self._watchdog_timer.setInterval(30000)  # Check every 30 seconds
        self._watchdog_timer.timeout.connect(self._watchdog_check)
        self._watchdog_timer.start()

    def _watchdog_check(self):
        """Auto-restart the hotkey thread if it died."""
        if self._stop_event.is_set():
            return
        if self._hotkey_thread is None or not self._hotkey_thread.is_alive():
            if not self._hotkey_registered:
                print("Win+V watchdog: hotkey thread died, restarting...")
                self._hotkey_thread = None
                self._hotkey_thread_id = None
                self._stop_event.clear()
                self._hotkey_thread = threading.Thread(target=self._hotkey_message_loop, daemon=True)
                self._hotkey_thread.start()

    def _fallback_to_keyboard_lib(self):
        """Fall back to keyboard library if RegisterHotKey fails."""
        if not HAS_KEYBOARD:
            print("keyboard module not available - hotkey disabled")
            return
        try:
            self._keyboard_handle = keyboard.add_hotkey('win+v', lambda: self.signals.show_popup.emit(), suppress=True)
            self._hotkey_registered = True
            print("Win+V hotkey registered via keyboard library (fallback)")
        except Exception as e:
            print(f"Keyboard library fallback also failed: {e}")

    def stop(self):
        """Stop listening for the global hotkey."""
        if self._watchdog_timer is not None:
            self._watchdog_timer.stop()
            self._watchdog_timer = None

        self._stop_event.set()

        # Post WM_QUIT to the hotkey thread to unblock GetMessage
        if self._hotkey_thread_id is not None:
            ctypes.windll.user32.PostThreadMessageW(
                self._hotkey_thread_id, 0x0012, 0, 0  # WM_QUIT
            )
            self._hotkey_thread_id = None

        # Clean up keyboard library fallback if used
        if hasattr(self, '_keyboard_handle') and HAS_KEYBOARD:
            try:
                keyboard.remove_hotkey(self._keyboard_handle)
            except:
                pass

        if self._hotkey_thread is not None:
            self._hotkey_thread.join(timeout=2)
            self._hotkey_thread = None

        self._hotkey_registered = False

    def _show_popup(self):
        """Show the popup dialog (on main thread)."""
        try:
            logger.debug(f"_show_popup called. popup={self.popup}, visible={self.popup.isVisible() if self.popup else 'N/A'}")

            if self.popup is None or not self.popup.isVisible():
                # Detect active application context BEFORE creating popup
                app_context, inspector = detect_active_app_context()
                logger.debug(f"_show_popup: app_context={app_context}")

                # Only show if Word or Outlook compose is active
                if app_context == APP_CONTEXT_UNKNOWN:
                    # Log what the foreground window actually is for debugging
                    try:
                        hwnd = win32gui.GetForegroundWindow()
                        cls = win32gui.GetClassName(hwnd) if hwnd else "None"
                        title = win32gui.GetWindowText(hwnd)[:60] if hwnd else "None"
                        logger.debug(f"Win+V pressed but neither Word nor Outlook compose is active (fg={cls}, title={title})")
                    except Exception:
                        logger.debug("Win+V pressed but neither Word nor Outlook compose is active")
                    return

                # Capture cursor position on main thread
                from PySide6.QtGui import QCursor
                cursor_pos = QCursor.pos()

                logger.debug("_show_popup: creating new WordLLMPopup...")
                self.popup = WordLLMPopup(
                    parent=None,  # No parent so it's a top-level window
                    llm_callback=self._get_llm_callback(),
                    cursor_pos=cursor_pos
                )
                logger.debug("_show_popup: popup created, setting context and showing")

                # Set the detected context before showing
                self.popup.set_app_context(app_context, inspector)

                # Capture the active document NOW (at Win+V time) before the
                # popup steals focus.  This prevents the bug where switching
                # to another Word document while typing a prompt causes the
                # result to be inserted into the wrong document.
                if app_context == APP_CONTEXT_WORD:
                    self.popup._capture_initial_document()

                self.popup.show()

                # Force focus to the popup window (Windows focus stealing prevention workaround)
                self._force_focus(self.popup)
            else:
                logger.debug("_show_popup: popup already visible, skipping")
        except Exception as e:
            logger.error(f"Error in _show_popup: {type(e).__name__}: {e}", exc_info=True)

    def _force_focus(self, window):
        """Force focus to the window using Windows API."""
        try:
            # First, use Qt methods
            window.raise_()
            window.activateWindow()
            QApplication.setActiveWindow(window)

            # Then use Windows API for guaranteed focus
            if HAS_WIN32:
                import ctypes
                hwnd = int(window.winId())

                # Allow our process to set foreground window
                ctypes.windll.user32.AllowSetForegroundWindow(-1)  # ASFW_ANY

                # Attach to the foreground thread to steal focus
                foreground_hwnd = ctypes.windll.user32.GetForegroundWindow()
                foreground_thread = ctypes.windll.user32.GetWindowThreadProcessId(foreground_hwnd, None)
                current_thread = ctypes.windll.kernel32.GetCurrentThreadId()

                if foreground_thread != current_thread:
                    ctypes.windll.user32.AttachThreadInput(foreground_thread, current_thread, True)

                ctypes.windll.user32.SetForegroundWindow(hwnd)
                ctypes.windll.user32.BringWindowToTop(hwnd)
                ctypes.windll.user32.SetFocus(hwnd)

                if foreground_thread != current_thread:
                    ctypes.windll.user32.AttachThreadInput(foreground_thread, current_thread, False)

            # Schedule focus to the input field after a brief delay
            QTimer.singleShot(50, lambda: self._focus_input(window))

        except Exception as e:
            print(f"Force focus error: {e}")
            # Fallback
            window.activateWindow()
            window.raise_()

    def _focus_input(self, window):
        """Set focus to the custom input field."""
        try:
            if window and hasattr(window, 'custom_input'):
                window.custom_input.setFocus()
                window.custom_input.activateWindow()
        except Exception as e:
            print(f"Focus input error: {e}")

    def _get_llm_callback(self):
        """Get the LLM callback function."""
        # Could customize this based on main_window settings
        return None  # Use default LLM call in popup


# Global instance
_hotkey_manager: Optional[WordHotkeyManager] = None


def init_word_hotkey(main_window=None) -> bool:
    """Initialize the Word hotkey manager."""
    global _hotkey_manager

    if _hotkey_manager is not None:
        # Manager exists — but ensure hotkey is actually alive
        if _hotkey_manager._hotkey_registered and \
           _hotkey_manager._hotkey_thread is not None and \
           _hotkey_manager._hotkey_thread.is_alive():
            return True
        # Thread died or registration lost — restart
        print("Win+V hotkey lost, restarting...")
        return _hotkey_manager.start()

    _hotkey_manager = WordHotkeyManager(main_window)
    return _hotkey_manager.start()


def stop_word_hotkey():
    """Stop the Word hotkey manager."""
    global _hotkey_manager

    if _hotkey_manager is not None:
        _hotkey_manager.stop()
        _hotkey_manager = None
