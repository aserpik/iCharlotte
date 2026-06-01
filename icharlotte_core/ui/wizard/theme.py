"""Wizard design system — central tokens, widget builders, and a scoped stylesheet.

Everything visual in Wizard Mode should pull from here so the tasks look and
behave consistently. Nothing in this module imports other wizard modules, so it
is safe to import from anywhere.

Usage:
    from icharlotte_core.ui.wizard import theme
    self.setStyleSheet(theme.wizard_stylesheet())     # on a wizard root widget
    btn = theme.primary_button("Continue")
    title = theme.page_title("Summarize Documents")

The stylesheet is applied to wizard *roots* (the landing page and each task
container), so it cascades to their children only — it does NOT touch the
classic-mode tabs.
"""
from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel, QPushButton, QWidget


# --------------------------------------------------------------------------- #
# Color tokens
# --------------------------------------------------------------------------- #

# Brand / primary action
PRIMARY = "#1976D2"
PRIMARY_HOVER = "#1565C0"
PRIMARY_PRESSED = "#0D47A1"
PRIMARY_SUBTLE = "#E8F1FB"   # tinted background (selected rows, icon tiles)

# Semantic
SUCCESS = "#1E8E3E"
SUCCESS_BG = "#E6F4EA"
WARNING = "#E65100"
WARNING_TEXT = "#8A5A00"
WARNING_BG = "#FFF7E6"
ERROR = "#D32F2F"
ERROR_HOVER = "#B71C1C"
ERROR_BG = "#FDECEA"
INFO = PRIMARY

# Neutrals — text
TEXT = "#1A1A1A"        # headings / emphasis
TEXT_BODY = "#333333"   # default body
TEXT_MUTED = "#666666"  # secondary / hints
TEXT_FAINT = "#909090"  # placeholders / disabled

# Neutrals — surfaces
BORDER = "#D8DCE0"
BORDER_LIGHT = "#ECEFF1"
BG = "#FFFFFF"
BG_SUBTLE = "#F7F9FC"
BG_CARD = "#FFFFFF"

# Fonts
MONO = "Consolas, 'Courier New', monospace"


# --------------------------------------------------------------------------- #
# Spacing / radius / type scale (px)
# --------------------------------------------------------------------------- #

SPACE_XS = 4
SPACE_SM = 8
SPACE_MD = 12
SPACE_LG = 16
SPACE_XL = 24
SPACE_XXL = 32

RADIUS_SM = 4
RADIUS_MD = 6
RADIUS_LG = 10

FONT_H1 = 22
FONT_H2 = 16
FONT_H3 = 13
FONT_BODY = 12
FONT_CAPTION = 11


# --------------------------------------------------------------------------- #
# Label builders
# --------------------------------------------------------------------------- #

def page_title(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setStyleSheet(f"font-size: {FONT_H2}px; font-weight: 600; color: {TEXT};")
    return lbl


def section_header(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setStyleSheet(f"font-size: {FONT_H3}px; font-weight: 600; color: {TEXT};")
    return lbl


def helper_text(text: str) -> QLabel:
    """A muted, word-wrapped one-liner that explains what to do on a page."""
    lbl = QLabel(text)
    lbl.setWordWrap(True)
    lbl.setStyleSheet(f"font-size: {FONT_BODY}px; color: {TEXT_MUTED};")
    return lbl


def caption(text: str) -> QLabel:
    lbl = QLabel(text)
    lbl.setStyleSheet(f"font-size: {FONT_CAPTION}px; color: {TEXT_MUTED};")
    return lbl


def error_text(text: str = "") -> QLabel:
    lbl = QLabel(text)
    lbl.setWordWrap(True)
    lbl.setStyleSheet(f"font-size: {FONT_CAPTION}px; color: {ERROR};")
    return lbl


# --------------------------------------------------------------------------- #
# Button builders / restylers
# --------------------------------------------------------------------------- #

def _btn_qss(bg: str, fg: str, hover: str, pressed: str, bold: bool = True) -> str:
    weight = "600" if bold else "500"
    return (
        f"QPushButton {{ background-color: {bg}; color: {fg}; font-weight: {weight};"
        f" border: none; padding: 7px 20px; border-radius: {RADIUS_SM}px; }}"
        f"QPushButton:hover {{ background-color: {hover}; }}"
        f"QPushButton:pressed {{ background-color: {pressed}; }}"
        f"QPushButton:disabled {{ background-color: #BFD4EA; color: #EAF2FB; }}"
    )


def style_primary_button(btn: QPushButton) -> QPushButton:
    btn.setStyleSheet(_btn_qss(PRIMARY, "#FFFFFF", PRIMARY_HOVER, PRIMARY_PRESSED))
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


def style_danger_button(btn: QPushButton) -> QPushButton:
    btn.setStyleSheet(_btn_qss(ERROR, "#FFFFFF", ERROR_HOVER, ERROR_HOVER))
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


def style_secondary_button(btn: QPushButton) -> QPushButton:
    """A quiet, outlined button (default look provided by the stylesheet)."""
    btn.setStyleSheet(
        f"QPushButton {{ background-color: #FFFFFF; color: {TEXT_BODY};"
        f" border: 1px solid {BORDER}; padding: 6px 16px; border-radius: {RADIUS_SM}px; }}"
        f"QPushButton:hover {{ background-color: {BG_SUBTLE}; border-color: #BCC4CC; }}"
        f"QPushButton:pressed {{ background-color: #EEF1F4; }}"
        f"QPushButton:disabled {{ color: {TEXT_FAINT}; background-color: #F5F6F8;"
        f" border-color: {BORDER_LIGHT}; }}"
    )
    btn.setCursor(Qt.CursorShape.PointingHandCursor)
    return btn


def primary_button(text: str) -> QPushButton:
    return style_primary_button(QPushButton(text))


def secondary_button(text: str) -> QPushButton:
    return style_secondary_button(QPushButton(text))


def danger_button(text: str) -> QPushButton:
    return style_danger_button(QPushButton(text))


# --------------------------------------------------------------------------- #
# Scoped stylesheet
# --------------------------------------------------------------------------- #

def wizard_stylesheet() -> str:
    """QSS applied to wizard root widgets so all default child controls share a
    consistent look. Accent buttons created via the builders above carry their
    own inline styling and intentionally override these defaults."""
    return f"""
    QLabel {{ color: {TEXT_BODY}; }}

    QPushButton {{
        background-color: #FFFFFF;
        color: {TEXT_BODY};
        border: 1px solid {BORDER};
        border-radius: {RADIUS_SM}px;
        padding: 6px 14px;
    }}
    QPushButton:hover {{ background-color: {BG_SUBTLE}; border-color: #BCC4CC; }}
    QPushButton:pressed {{ background-color: #EEF1F4; }}
    QPushButton:disabled {{ color: {TEXT_FAINT}; background-color: #F5F6F8; border-color: {BORDER_LIGHT}; }}

    QLineEdit, QPlainTextEdit, QTextEdit, QComboBox, QSpinBox, QListWidget {{
        border: 1px solid {BORDER};
        border-radius: {RADIUS_SM}px;
        background-color: #FFFFFF;
        selection-background-color: {PRIMARY};
        selection-color: #FFFFFF;
    }}
    QLineEdit, QComboBox, QSpinBox {{ padding: 5px 8px; }}
    QPlainTextEdit, QTextEdit {{ padding: 6px 8px; }}
    QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus,
    QComboBox:focus, QSpinBox:focus, QListWidget:focus {{
        border: 1px solid {PRIMARY};
    }}

    QListWidget::item {{ padding: 4px 6px; border-radius: 3px; }}
    QListWidget::item:selected {{ background-color: {PRIMARY_SUBTLE}; color: {TEXT}; }}

    QProgressBar {{
        border: 1px solid {BORDER};
        border-radius: {RADIUS_SM}px;
        background-color: {BG_SUBTLE};
        text-align: center;
        height: 16px;
    }}
    QProgressBar::chunk {{ background-color: {PRIMARY}; border-radius: 3px; }}

    QGroupBox {{
        border: 1px solid {BORDER};
        border-radius: {RADIUS_MD}px;
        margin-top: 10px;
        font-weight: 600;
        color: {TEXT};
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        subcontrol-position: top left;
        left: 10px;
        padding: 0 4px;
    }}

    QScrollArea {{ border: none; }}
    QSplitter::handle {{ background-color: {BORDER_LIGHT}; }}
    QSplitter::handle:hover {{ background-color: #C4CCD4; }}
    QSplitter::handle:pressed {{ background-color: #AAB4BD; }}

    QToolTip {{
        background-color: #2A2A2A;
        color: #FFFFFF;
        border: none;
        padding: 4px 7px;
    }}
    """
