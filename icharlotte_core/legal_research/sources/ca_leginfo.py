"""California Legislative Information scraper client.

Fetches statute text from https://leginfo.legislature.ca.gov by code and
section number.  Parses natural-language code references (e.g. "Civil Code
section 1714") into the internal code abbreviation + section used by the site.
"""
import logging
import re
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlencode

import requests
from bs4 import BeautifulSoup

from icharlotte_core.legal_research.models import StatuteResult

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Code alias / title dictionaries
# ---------------------------------------------------------------------------

CODE_TITLES: Dict[str, str] = {
    "BPC": "Business and Professions Code",
    "CCP": "Code of Civil Procedure",
    "CIV": "Civil Code",
    "COM": "Commercial Code",
    "CORP": "Corporations Code",
    "EDC": "Education Code",
    "ELEC": "Elections Code",
    "EVID": "Evidence Code",
    "FAM": "Family Code",
    "FIN": "Financial Code",
    "FGC": "Fish and Game Code",
    "FCA": "Food and Agricultural Code",
    "GOV": "Government Code",
    "HNC": "Harbors and Navigation Code",
    "HSC": "Health and Safety Code",
    "INS": "Insurance Code",
    "LAB": "Labor Code",
    "MVC": "Military and Veterans Code",
    "PEN": "Penal Code",
    "PROB": "Probate Code",
    "PCC": "Public Contract Code",
    "PRC": "Public Resources Code",
    "PUC": "Public Utilities Code",
    "RTC": "Revenue and Taxation Code",
    "SHC": "Streets and Highways Code",
    "UIC": "Unemployment Insurance Code",
    "VEH": "Vehicle Code",
    "WAT": "Water Code",
    "WIC": "Welfare and Institutions Code",
}

# Maps lowercase full names AND common abbreviations to internal code keys.
CODE_ALIASES: Dict[str, str] = {
    # Full names (lowercase)
    "business and professions code": "BPC",
    "code of civil procedure": "CCP",
    "civil code": "CIV",
    "commercial code": "COM",
    "corporations code": "CORP",
    "education code": "EDC",
    "elections code": "ELEC",
    "evidence code": "EVID",
    "family code": "FAM",
    "financial code": "FIN",
    "fish and game code": "FGC",
    "food and agricultural code": "FCA",
    "government code": "GOV",
    "harbors and navigation code": "HNC",
    "health and safety code": "HSC",
    "insurance code": "INS",
    "labor code": "LAB",
    "military and veterans code": "MVC",
    "penal code": "PEN",
    "probate code": "PROB",
    "public contract code": "PCC",
    "public resources code": "PRC",
    "public utilities code": "PUC",
    "revenue and taxation code": "RTC",
    "streets and highways code": "SHC",
    "unemployment insurance code": "UIC",
    "vehicle code": "VEH",
    "water code": "WAT",
    "welfare and institutions code": "WIC",
    # Common abbreviations (lowercase)
    "bus. & prof. code": "BPC",
    "bpc": "BPC",
    "ccp": "CCP",
    "code civ. proc.": "CCP",
    "civ. code": "CIV",
    "civ": "CIV",
    "com. code": "COM",
    "corp. code": "CORP",
    "educ. code": "EDC",
    "elec. code": "ELEC",
    "evid. code": "EVID",
    "fam. code": "FAM",
    "fin. code": "FIN",
    "fish & game code": "FGC",
    "food & agric. code": "FCA",
    "gov. code": "GOV",
    "gov": "GOV",
    "harb. & nav. code": "HNC",
    "health & safety code": "HSC",
    "hsc": "HSC",
    "ins. code": "INS",
    "lab. code": "LAB",
    "mil. & vet. code": "MVC",
    "pen. code": "PEN",
    "pen": "PEN",
    "prob. code": "PROB",
    "pub. contract code": "PCC",
    "pub. resources code": "PRC",
    "pub. util. code": "PUC",
    "rev. & tax. code": "RTC",
    "sts. & hy. code": "SHC",
    "unemp. ins. code": "UIC",
    "veh. code": "VEH",
    "veh": "VEH",
    "water code": "WAT",
    "welf. & inst. code": "WIC",
}

# Minimum character length for scraped statute text to be considered valid.
_MIN_TEXT_LENGTH = 20


class CALegInfoClient:
    """Scraper client for California Legislative Information website."""

    BASE_URL = (
        "https://leginfo.legislature.ca.gov/faces/codes_displaySection.xhtml"
    )

    # Pre-compiled regex: matches "<code name> (section|sec\.|s\.|section-symbol) <number>"
    # The code-name portion is greedy so multi-word names are captured.
    _REF_PATTERN = re.compile(
        r"(.+?)\s+(?:section|sec\.?|s\.|\u00a7)\s*(\d+\w*)",
        re.IGNORECASE,
    )

    def __init__(self, timeout: int = 15):
        self.timeout = timeout

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def parse_code_reference(self, text: str) -> Tuple[Optional[str], Optional[str]]:
        """Parse a natural-language code reference into (code_key, section).

        Examples:
            "Civil Code section 1714"           -> ("CIV", "1714")
            "Code of Civil Procedure section 437c" -> ("CCP", "437c")
            "Civ. Code section-symbol 1714"          -> ("CIV", "1714")

        Returns (None, None) if the reference cannot be parsed.
        """
        m = self._REF_PATTERN.search(text)
        if not m:
            return None, None

        raw_code = m.group(1).strip().lower()
        section = m.group(2).strip()

        # Try progressively shorter suffixes of raw_code to handle leading
        # noise (e.g. "pursuant to the Civil Code").
        words = raw_code.split()
        for start in range(len(words)):
            candidate = " ".join(words[start:])
            if candidate in CODE_ALIASES:
                return CODE_ALIASES[candidate], section

        return None, None

    def get_section(
        self, code: str, section: str
    ) -> Optional[StatuteResult]:
        """Fetch a single code section from leginfo.legislature.ca.gov.

        Args:
            code: Internal code key (e.g. "CIV", "CCP").
            section: Section number (e.g. "1714", "437c").

        Returns:
            A StatuteResult on success, or None if the section was not found
            or an error occurred.
        """
        params = {"sectionNum": section, "lawCode": code}
        url = f"{self.BASE_URL}?{urlencode(params)}"

        try:
            resp = requests.get(url, timeout=self.timeout)
            resp.raise_for_status()
        except Exception:
            logger.warning("Failed to fetch %s \u00a7 %s: %s", code, section, url)
            return None

        soup = BeautifulSoup(resp.text, "html.parser")
        content_div = soup.find("div", id="codeLawSectionNoContent")
        if content_div is None:
            return None

        text = content_div.get_text(separator="\n", strip=True)
        if len(text) < _MIN_TEXT_LENGTH:
            return None

        title = CODE_TITLES.get(code, code)
        return StatuteResult(
            code=code,
            section=section,
            title=title,
            text=text,
            url=url,
        )

    def search_code(
        self, code: str, section: str
    ) -> List[StatuteResult]:
        """Search for a code section. Returns a list with 0 or 1 results.

        This is a thin wrapper around :meth:`get_section` that returns the
        result in list form for consistency with other search interfaces.
        """
        result = self.get_section(code, section)
        if result is not None:
            return [result]
        return []
