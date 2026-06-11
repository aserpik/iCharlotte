"""URL helpers for Harvard Caselaw Access Project bulk records."""
from __future__ import annotations

import re
from urllib.parse import quote_plus, urlparse


_CAL_UNREP_RE = re.compile(
    r"\b(?P<volume>\d+)\s+Cal\.?\s+Unrep\.?\s+(?P<page>\d+)\b",
    re.I,
)
_CAL_REPORTER_RE = re.compile(
    r"\b(?P<volume>\d+)\s+Cal\.?\s*"
    r"(?P<app>App\.?\s*)?"
    r"(?P<rptr>Rptr\.?\s*)?"
    r"(?P<series>2d|3d|4th|5th)?"
    r"\s+(?P<page>\d+)\b",
    re.I,
)
_BROKEN_STATIC_ID_RE = re.compile(r"^/cap:\d+/?$", re.I)


def is_broken_cap_static_id_url(url: str) -> bool:
    parsed = urlparse(str(url or "").strip())
    return (
        parsed.scheme.lower() in {"http", "https"}
        and parsed.netloc.lower() == "static.case.law"
        and bool(_BROKEN_STATIC_ID_RE.match(parsed.path or ""))
    )


def cap_static_html_url(citation: str, *, file_name: str = "") -> str:
    parsed = _parse_cal_reporter_citation(citation)
    if not parsed:
        return ""
    reporter, volume, page = parsed
    html_file = _normalize_file_name(file_name) or f"{int(page):04d}-01.html"
    return f"https://static.case.law/{reporter}/{volume}/html/{html_file}"


def cap_citation_search_url(citation: str) -> str:
    value = str(citation or "").strip()
    if not value:
        return ""
    return f"https://case.law/caselaw/?citation={quote_plus(value)}"


def repair_broken_cap_url(url: str, citation: str) -> str:
    value = str(url or "").strip()
    if value and not is_broken_cap_static_id_url(value):
        return value
    return cap_static_html_url(citation) or cap_citation_search_url(citation)


def _parse_cal_reporter_citation(citation: str) -> tuple[str, str, str] | None:
    text = str(citation or "")
    unrep = _CAL_UNREP_RE.search(text)
    if unrep:
        return "cal-unrep", unrep.group("volume"), unrep.group("page")

    match = _CAL_REPORTER_RE.search(text)
    if not match:
        return None

    series = (match.group("series") or "").lower()
    suffix = f"-{series}" if series else ""
    if match.group("rptr"):
        reporter = f"cal-rptr{suffix}"
    elif match.group("app"):
        reporter = f"cal-app{suffix}"
    else:
        reporter = f"cal{suffix}"
    return reporter, match.group("volume"), match.group("page")


def _normalize_file_name(file_name: str) -> str:
    value = str(file_name or "").strip().strip("/")
    if not value:
        return ""
    if not value.lower().endswith(".html"):
        value = f"{value}.html"
    return value
