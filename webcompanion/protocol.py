"""Parsing for the wizard agent stdout protocol.

Same line grammar as icharlotte_core/ui/wizard/runners/subprocess_worker.py:
  PROGRESS:<int>[:<message>]   AWAITING_INPUT:<path>   OUTPUT:<path>
Anything else is a plain status line.
"""
import re
from dataclasses import dataclass
from typing import Optional

_PROGRESS_RE = re.compile(r"^PROGRESS:\s*(\d+)\s*(?::(.*))?$")
_AWAITING_RE = re.compile(r"^AWAITING_INPUT:(.+)$")
_OUTPUT_RE = re.compile(r"^OUTPUT:(.+)$")


@dataclass(frozen=True)
class ParsedLine:
    kind: str  # 'progress' | 'awaiting' | 'output' | 'status'
    pct: Optional[int] = None
    message: str = ""
    path: str = ""


def parse_line(line: str) -> ParsedLine:
    m = _PROGRESS_RE.match(line)
    if m:
        pct = max(0, min(100, int(m.group(1))))
        return ParsedLine(kind="progress", pct=pct, message=(m.group(2) or "").strip())
    m = _AWAITING_RE.match(line)
    if m:
        return ParsedLine(kind="awaiting", path=m.group(1).strip())
    m = _OUTPUT_RE.match(line)
    if m:
        return ParsedLine(kind="output", path=m.group(1).strip())
    return ParsedLine(kind="status", message=line)
