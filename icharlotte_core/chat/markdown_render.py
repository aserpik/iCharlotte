"""Markdown -> HTML rendering for the chat UI.

Qt's QTextDocument renders only a CSS subset. Pipe tables produced by
python-markdown lose their borders unless we add HTML attributes; bare
`<pre>`/`<blockquote>` look unstyled until we attach a stylesheet via
`QTextDocument.setDefaultStyleSheet(CHAT_MARKDOWN_CSS)`.
"""
from __future__ import annotations

import re

import markdown


_MD_EXTENSIONS = ["fenced_code", "tables", "nl2br", "sane_lists"]

_TABLE_OPEN_RE = re.compile(r"<table(\s[^>]*)?>")


CHAT_MARKDOWN_CSS = """
table { border-collapse: collapse; margin: 6px 0; }
th, td { border: 1px solid #888; padding: 4px 8px; vertical-align: top; }
th { background-color: #eaeaea; font-weight: bold; text-align: left; }
pre { background-color: #f4f4f4; padding: 8px; border-radius: 4px;
      font-family: Consolas, 'Courier New', monospace; font-size: 12px; }
code { background-color: #f4f4f4; padding: 1px 4px; border-radius: 3px;
       font-family: Consolas, 'Courier New', monospace; }
blockquote { border-left: 3px solid #999; margin-left: 8px;
             padding-left: 8px; color: #555; }
h1 { font-size: 18px; margin: 10px 0 4px 0; }
h2 { font-size: 16px; margin: 8px 0 4px 0; }
h3 { font-size: 14px; margin: 6px 0 4px 0; }
ul, ol { margin: 4px 0 4px 20px; }
hr { border: 0; border-top: 1px solid #ccc; margin: 8px 0; }
"""


def render_markdown(text: str | None) -> str:
    """Render chat-message markdown to Qt-friendly HTML.

    - Pipe tables get `border`/`cellpadding`/`cellspacing` attributes so Qt
      draws real gridlines (QTextDocument ignores `border-collapse` on its own).
    - GFM-style single-newline line breaks via `nl2br`.
    - Empty/None input returns an empty string.
    """
    if not text:
        return ""
    try:
        html = markdown.markdown(text, extensions=_MD_EXTENSIONS)
    except Exception:
        return text.replace("\n", "<br>")

    def _attrs(match: re.Match) -> str:
        existing = match.group(1) or ""
        return f'<table border="1" cellspacing="0" cellpadding="4"{existing}>'

    return _TABLE_OPEN_RE.sub(_attrs, html)
