# Legal Research Agent Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add a legal research agent that queries CourtListener, CA Leginfo, and CA Courts for real case law and statutes, then injects verified citations into LLM responses in both ChatTab and Win+V Word popup.

**Architecture:** Intercept-and-augment — when the "Perform Legal Research" checkbox is checked, intercept the prompt before the LLM call, run a 4-phase pipeline (plan → search → draft → verify), and inject `[LEGAL AUTHORITY]` context so the LLM cites only real sources. The engine is pure Python (no Qt), wrapped in QThread workers for each UI surface.

**Tech Stack:** Python `requests` for API calls, `beautifulsoup4` for HTML scraping, existing `LLMHandler` for LLM calls, `CaseDataManager` for caching, PyQt6 for UI checkboxes.

---

## Task 1: Data Models

**Files:**
- Create: `icharlotte_core/legal_research/__init__.py`
- Create: `icharlotte_core/legal_research/models.py`
- Test: `tests/test_legal_research/test_models.py`
- Create: `tests/test_legal_research/__init__.py`

**Step 1: Write the failing test**

```python
# tests/test_legal_research/__init__.py
# (empty)

# tests/test_legal_research/test_models.py
import unittest
from icharlotte_core.legal_research.models import (
    CaseResult, StatuteResult, ResearchResult, VerificationStatus
)

class TestCaseResult(unittest.TestCase):
    def test_case_result_creation(self):
        case = CaseResult(
            name="Rowland v. Christian",
            citation="69 Cal.2d 108",
            date="1968-08-08",
            court="Supreme Court of California",
            snippet="A person who maintains premises is liable...",
            url="https://www.courtlistener.com/opinion/123/",
            cluster_id=123
        )
        self.assertEqual(case.name, "Rowland v. Christian")
        self.assertEqual(case.citation, "69 Cal.2d 108")

    def test_case_result_formatted_citation(self):
        case = CaseResult(
            name="Rowland v. Christian",
            citation="69 Cal.2d 108",
            date="1968-08-08",
            court="Supreme Court of California",
            snippet="",
            url="",
            cluster_id=123
        )
        self.assertIn("Rowland v. Christian", case.formatted_citation)
        self.assertIn("69 Cal.2d 108", case.formatted_citation)
        self.assertIn("1968", case.formatted_citation)

class TestStatuteResult(unittest.TestCase):
    def test_statute_result_creation(self):
        statute = StatuteResult(
            code="CIV",
            section="1714",
            title="Civil Code",
            text="Everyone is responsible for an injury...",
            url="https://leginfo.legislature.ca.gov/..."
        )
        self.assertEqual(statute.code, "CIV")
        self.assertEqual(statute.section, "1714")

    def test_statute_formatted_citation(self):
        statute = StatuteResult(
            code="CIV", section="1714", title="Civil Code",
            text="", url=""
        )
        self.assertIn("Civ. Code", statute.formatted_citation)
        self.assertIn("1714", statute.formatted_citation)

class TestResearchResult(unittest.TestCase):
    def test_empty_result(self):
        result = ResearchResult(query="test", cases=[], statutes=[], memo="")
        self.assertEqual(len(result.cases), 0)
        self.assertEqual(len(result.statutes), 0)

    def test_authority_block_formatting(self):
        case = CaseResult(
            name="Test v. Case", citation="1 Cal.2d 1",
            date="2020-01-01", court="Supreme Court of California",
            snippet="Holding text here", url="", cluster_id=1
        )
        statute = StatuteResult(
            code="CIV", section="1714", title="Civil Code",
            text="Section text here", url=""
        )
        result = ResearchResult(
            query="test", cases=[case], statutes=[statute], memo=""
        )
        block = result.format_authority_block()
        self.assertIn("[LEGAL AUTHORITY]", block)
        self.assertIn("Test v. Case", block)
        self.assertIn("Civ. Code", block)

class TestVerificationStatus(unittest.TestCase):
    def test_pass_status(self):
        v = VerificationStatus(citation="Test v. Case", status="PASS", detail="")
        self.assertEqual(v.status, "PASS")

    def test_fixed_status(self):
        v = VerificationStatus(
            citation="Test v. Case", status="FIXED",
            detail="Wrong page, corrected to p.412",
            original="p.410", corrected="p.412"
        )
        self.assertEqual(v.status, "FIXED")

if __name__ == "__main__":
    unittest.main()
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_legal_research/test_models.py -v`
Expected: FAIL with "ModuleNotFoundError: No module named 'icharlotte_core.legal_research'"

**Step 3: Write minimal implementation**

```python
# icharlotte_core/legal_research/__init__.py
"""Legal research agent for California case law and statutes."""

# icharlotte_core/legal_research/models.py
from dataclasses import dataclass, field
from typing import List, Optional

# Map internal code abbreviations to display names
CODE_DISPLAY_NAMES = {
    "CIV": "Civ. Code",
    "CCP": "Code Civ. Proc.",
    "PEN": "Pen. Code",
    "VEH": "Veh. Code",
    "FAM": "Fam. Code",
    "PROB": "Prob. Code",
    "GOV": "Gov. Code",
    "BPC": "Bus. & Prof. Code",
    "LAB": "Lab. Code",
    "INS": "Ins. Code",
    "CORP": "Corp. Code",
    "EVID": "Evid. Code",
    "HSC": "Health & Saf. Code",
    "WIC": "Welf. & Inst. Code",
    "EDC": "Ed. Code",
    "COM": "Com. Code",
    "FIN": "Fin. Code",
    "CONS": "Cal. Const.",
}


@dataclass
class CaseResult:
    """A case law search result."""
    name: str                    # e.g. "Rowland v. Christian"
    citation: str                # e.g. "69 Cal.2d 108"
    date: str                    # ISO date "1968-08-08"
    court: str                   # e.g. "Supreme Court of California"
    snippet: str                 # Key excerpt from opinion
    url: str                     # CourtListener URL
    cluster_id: Optional[int] = None  # For citation graph lookups
    negative_treatment: Optional[str] = None  # e.g. "Overruled by..."
    relevance_score: float = 0.0

    @property
    def formatted_citation(self) -> str:
        year = self.date[:4] if self.date else ""
        return f"{self.name} ({year}) {self.citation}"


@dataclass
class StatuteResult:
    """A statute/code search result."""
    code: str         # Internal code abbreviation: "CIV", "CCP", etc.
    section: str      # Section number: "1714"
    title: str        # Full code name: "Civil Code"
    text: str         # Section text
    url: str          # leginfo URL

    @property
    def formatted_citation(self) -> str:
        display = CODE_DISPLAY_NAMES.get(self.code, self.title)
        return f"{display}, \u00a7 {self.section}"


@dataclass
class VerificationStatus:
    """Verification result for a single citation."""
    citation: str
    status: str           # "PASS", "FIXED", "FLAGGED"
    detail: str = ""
    original: str = ""    # Original text (for FIXED)
    corrected: str = ""   # Corrected text (for FIXED)


@dataclass
class ResearchResult:
    """Complete research result from the engine."""
    query: str
    cases: List[CaseResult] = field(default_factory=list)
    statutes: List[StatuteResult] = field(default_factory=list)
    memo: str = ""
    verification: List[VerificationStatus] = field(default_factory=list)

    def format_authority_block(self) -> str:
        """Format results as an [LEGAL AUTHORITY] block for LLM injection."""
        lines = ["[LEGAL AUTHORITY]", ""]
        if self.cases:
            lines.append("CASE LAW:")
            for c in self.cases:
                lines.append(f"  - {c.formatted_citation}")
                if c.snippet:
                    # Truncate long snippets
                    snip = c.snippet[:500] + "..." if len(c.snippet) > 500 else c.snippet
                    lines.append(f"    Holding/Excerpt: {snip}")
                if c.negative_treatment:
                    lines.append(f"    WARNING: {c.negative_treatment}")
                lines.append("")
        if self.statutes:
            lines.append("STATUTES:")
            for s in self.statutes:
                lines.append(f"  - {s.formatted_citation}")
                if s.text:
                    snip = s.text[:500] + "..." if len(s.text) > 500 else s.text
                    lines.append(f"    Text: {snip}")
                lines.append("")
        if not self.cases and not self.statutes:
            lines.append("No relevant legal authority found for this query.")
        lines.append("[END LEGAL AUTHORITY]")
        return "\n".join(lines)

    def to_dict(self) -> dict:
        """Serialize for CaseDataManager storage."""
        return {
            "query": self.query,
            "cases": [
                {"name": c.name, "citation": c.citation, "date": c.date,
                 "court": c.court, "snippet": c.snippet, "url": c.url,
                 "cluster_id": c.cluster_id, "negative_treatment": c.negative_treatment}
                for c in self.cases
            ],
            "statutes": [
                {"code": s.code, "section": s.section, "title": s.title,
                 "text": s.text, "url": s.url}
                for s in self.statutes
            ],
            "memo": self.memo,
            "verification": [
                {"citation": v.citation, "status": v.status, "detail": v.detail}
                for v in self.verification
            ],
        }
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_legal_research/test_models.py -v`
Expected: All 7 tests PASS

**Step 5: Commit**

```bash
git add icharlotte_core/legal_research/__init__.py icharlotte_core/legal_research/models.py tests/test_legal_research/__init__.py tests/test_legal_research/test_models.py
git commit -m "feat(legal-research): add data models for case law, statutes, and research results"
```

---

## Task 2: CourtListener API Client

**Files:**
- Create: `icharlotte_core/legal_research/sources/__init__.py`
- Create: `icharlotte_core/legal_research/sources/courtlistener.py`
- Test: `tests/test_legal_research/test_courtlistener.py`

**Step 1: Write the failing test**

```python
# tests/test_legal_research/test_courtlistener.py
import unittest
from unittest.mock import patch, MagicMock
from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
from icharlotte_core.legal_research.models import CaseResult


class TestCourtListenerClient(unittest.TestCase):
    def setUp(self):
        self.client = CourtListenerClient(token="test-token-123")

    def test_init_stores_token(self):
        self.assertEqual(self.client.token, "test-token-123")

    def test_headers_include_auth(self):
        headers = self.client._headers()
        self.assertEqual(headers["Authorization"], "Token test-token-123")

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_opinions_returns_case_results(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "count": 1,
            "results": [
                {
                    "caseName": "Rowland v. Christian",
                    "citation": ["69 Cal.2d 108"],
                    "dateFiled": "1968-08-08",
                    "court": "Supreme Court of California",
                    "snippet": "A person who maintains premises...",
                    "cluster_id": 123,
                    "absolute_url": "/opinion/123/rowland-v-christian/",
                }
            ],
        }
        mock_get.return_value = mock_response

        results = self.client.search_opinions("premises liability duty of care")
        self.assertEqual(len(results), 1)
        self.assertIsInstance(results[0], CaseResult)
        self.assertEqual(results[0].name, "Rowland v. Christian")
        self.assertEqual(results[0].citation, "69 Cal.2d 108")
        self.assertEqual(results[0].cluster_id, 123)

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_opinions_handles_empty_results(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"count": 0, "results": []}
        mock_get.return_value = mock_response

        results = self.client.search_opinions("nonexistent legal topic xyz")
        self.assertEqual(len(results), 0)

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_opinions_handles_api_error(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 429
        mock_response.raise_for_status.side_effect = Exception("Rate limited")
        mock_get.return_value = mock_response

        results = self.client.search_opinions("test query")
        self.assertEqual(len(results), 0)

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_search_filters_by_california(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {"count": 0, "results": []}
        mock_get.return_value = mock_response

        self.client.search_opinions("test", jurisdiction="cal")
        call_args = mock_get.call_args
        self.assertIn("court", call_args[1].get("params", {}))

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_citing_cases(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.json.return_value = {
            "count": 1,
            "results": [
                {
                    "caseName": "Later v. Case",
                    "citation": ["100 Cal.App.4th 200"],
                    "dateFiled": "2002-03-15",
                    "court": "Court of Appeal",
                    "snippet": "Following Rowland...",
                    "cluster_id": 456,
                    "absolute_url": "/opinion/456/later-v-case/",
                }
            ],
        }
        mock_get.return_value = mock_response

        results = self.client.get_citing_cases(cluster_id=123)
        self.assertEqual(len(results), 1)
        self.assertEqual(results[0].name, "Later v. Case")


if __name__ == "__main__":
    unittest.main()
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_legal_research/test_courtlistener.py -v`
Expected: FAIL with "ModuleNotFoundError"

**Step 3: Write minimal implementation**

```python
# icharlotte_core/legal_research/sources/__init__.py
"""Legal research data source clients."""

# icharlotte_core/legal_research/sources/courtlistener.py
"""CourtListener REST API v4 client for California case law."""
import requests
from typing import List, Optional
from ..models import CaseResult

# California court identifiers for CourtListener
CA_COURTS = "cal,calctapp,calag,calapp1st,calapp2nd,calapp3rd,calapp4th,calapp5th,calapp6th"


class CourtListenerClient:
    """Client for the CourtListener REST API (v4)."""
    BASE_URL = "https://www.courtlistener.com/api/rest/v4"

    def __init__(self, token: str):
        self.token = token

    def _headers(self) -> dict:
        return {
            "Authorization": f"Token {self.token}",
            "Content-Type": "application/json",
        }

    def search_opinions(
        self, query: str, jurisdiction: str = "cal", max_results: int = 15
    ) -> List[CaseResult]:
        """Search for opinions matching query in California courts."""
        try:
            params = {
                "q": query,
                "type": "o",  # opinions
                "court": CA_COURTS if jurisdiction == "cal" else jurisdiction,
                "order_by": "score desc",
                "page_size": max_results,
            }
            resp = requests.get(
                f"{self.BASE_URL}/search/",
                headers=self._headers(),
                params=params,
                timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
            return [self._parse_result(r) for r in data.get("results", [])]
        except Exception as e:
            print(f"[CourtListener] Search error: {e}")
            return []

    def get_citing_cases(
        self, cluster_id: int, max_results: int = 10
    ) -> List[CaseResult]:
        """Get cases that cite the given cluster (for cite-checking)."""
        try:
            params = {
                "q": f"cites:({cluster_id})",
                "type": "o",
                "order_by": "dateFiled desc",
                "page_size": max_results,
            }
            resp = requests.get(
                f"{self.BASE_URL}/search/",
                headers=self._headers(),
                params=params,
                timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
            return [self._parse_result(r) for r in data.get("results", [])]
        except Exception as e:
            print(f"[CourtListener] Citing cases error: {e}")
            return []

    def get_opinion_text(self, cluster_id: int) -> Optional[str]:
        """Get the full opinion text for a cluster."""
        try:
            resp = requests.get(
                f"{self.BASE_URL}/clusters/{cluster_id}/",
                headers=self._headers(),
                timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
            # Get the first opinion's plain text
            sub_opinions = data.get("sub_opinions", [])
            if sub_opinions:
                op_url = sub_opinions[0].get("resource_uri", "")
                if op_url:
                    op_resp = requests.get(
                        op_url, headers=self._headers(), timeout=30
                    )
                    op_resp.raise_for_status()
                    op_data = op_resp.json()
                    return op_data.get("plain_text") or op_data.get("html", "")
            return None
        except Exception as e:
            print(f"[CourtListener] Opinion text error: {e}")
            return None

    def _parse_result(self, r: dict) -> CaseResult:
        """Parse a CourtListener search result into a CaseResult."""
        citations = r.get("citation", [])
        citation_str = citations[0] if citations else ""
        return CaseResult(
            name=r.get("caseName", "Unknown"),
            citation=citation_str,
            date=r.get("dateFiled", ""),
            court=r.get("court", ""),
            snippet=r.get("snippet", ""),
            url=f"https://www.courtlistener.com{r.get('absolute_url', '')}",
            cluster_id=r.get("cluster_id"),
        )
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_legal_research/test_courtlistener.py -v`
Expected: All 7 tests PASS

**Step 5: Commit**

```bash
git add icharlotte_core/legal_research/sources/__init__.py icharlotte_core/legal_research/sources/courtlistener.py tests/test_legal_research/test_courtlistener.py
git commit -m "feat(legal-research): add CourtListener API client with search and cite-checking"
```

---

## Task 3: CA Legislative Info Client

**Files:**
- Create: `icharlotte_core/legal_research/sources/ca_leginfo.py`
- Test: `tests/test_legal_research/test_ca_leginfo.py`

**Step 1: Write the failing test**

```python
# tests/test_legal_research/test_ca_leginfo.py
import unittest
from unittest.mock import patch, MagicMock
from icharlotte_core.legal_research.sources.ca_leginfo import CALegInfoClient
from icharlotte_core.legal_research.models import StatuteResult


class TestCALegInfoClient(unittest.TestCase):
    def setUp(self):
        self.client = CALegInfoClient()

    def test_parse_code_reference_civil_code(self):
        code, section = self.client.parse_code_reference("Civil Code section 1714")
        self.assertEqual(code, "CIV")
        self.assertEqual(section, "1714")

    def test_parse_code_reference_ccp(self):
        code, section = self.client.parse_code_reference("Code of Civil Procedure section 437c")
        self.assertEqual(code, "CCP")
        self.assertEqual(section, "437c")

    def test_parse_code_reference_abbreviated(self):
        code, section = self.client.parse_code_reference("Civ. Code § 1714")
        self.assertEqual(code, "CIV")
        self.assertEqual(section, "1714")

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_get_section_returns_statute(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = '''<html><body>
        <div id="codeLawSectionNoContent">
            <p>Everyone is responsible, not only for the result of his or her willful acts,
            but also for an injury occasioned to another by his or her want of ordinary care.</p>
        </div>
        </body></html>'''
        mock_get.return_value = mock_response

        result = self.client.get_section("CIV", "1714")
        self.assertIsInstance(result, StatuteResult)
        self.assertEqual(result.code, "CIV")
        self.assertEqual(result.section, "1714")
        self.assertIn("responsible", result.text)

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_get_section_handles_not_found(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = "<html><body>Section not found</body></html>"
        mock_get.return_value = mock_response

        result = self.client.get_section("CIV", "99999999")
        self.assertIsNone(result)

    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    def test_search_code_returns_results(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = '''<html><body>
        <div id="codeLawSectionNoContent">
            <p>Test section text</p>
        </div>
        </body></html>'''
        mock_get.return_value = mock_response

        results = self.client.search_code("CIV", "1714")
        self.assertIsInstance(results, list)


if __name__ == "__main__":
    unittest.main()
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_legal_research/test_ca_leginfo.py -v`
Expected: FAIL with "ModuleNotFoundError"

**Step 3: Write minimal implementation**

```python
# icharlotte_core/legal_research/sources/ca_leginfo.py
"""California Legislative Information scraper for statutes and codes."""
import re
import requests
from typing import List, Optional, Tuple
from bs4 import BeautifulSoup
from ..models import StatuteResult

# Map various code references to leginfo URL codes
CODE_ALIASES = {
    # Full names
    "civil code": "CIV",
    "code of civil procedure": "CCP",
    "penal code": "PEN",
    "vehicle code": "VEH",
    "family code": "FAM",
    "probate code": "PROB",
    "government code": "GOV",
    "business and professions code": "BPC",
    "labor code": "LAB",
    "insurance code": "INS",
    "corporations code": "CORP",
    "evidence code": "EVID",
    "health and safety code": "HSC",
    "welfare and institutions code": "WIC",
    "education code": "EDC",
    "commercial code": "COM",
    "financial code": "FIN",
    "constitution": "CONS",
    # Abbreviations
    "civ. code": "CIV",
    "civ code": "CIV",
    "ccp": "CCP",
    "code civ. proc.": "CCP",
    "pen. code": "PEN",
    "veh. code": "VEH",
    "fam. code": "FAM",
    "prob. code": "PROB",
    "gov. code": "GOV",
    "gov code": "GOV",
    "bus. & prof. code": "BPC",
    "lab. code": "LAB",
    "ins. code": "INS",
    "corp. code": "CORP",
    "evid. code": "EVID",
    "health & saf. code": "HSC",
    "welf. & inst. code": "WIC",
    "ed. code": "EDC",
    "com. code": "COM",
    "fin. code": "FIN",
}

# Full display names for codes
CODE_TITLES = {
    "CIV": "Civil Code",
    "CCP": "Code of Civil Procedure",
    "PEN": "Penal Code",
    "VEH": "Vehicle Code",
    "FAM": "Family Code",
    "PROB": "Probate Code",
    "GOV": "Government Code",
    "BPC": "Business and Professions Code",
    "LAB": "Labor Code",
    "INS": "Insurance Code",
    "CORP": "Corporations Code",
    "EVID": "Evidence Code",
    "HSC": "Health and Safety Code",
    "WIC": "Welfare and Institutions Code",
    "EDC": "Education Code",
    "COM": "Commercial Code",
    "FIN": "Financial Code",
    "CONS": "California Constitution",
}


class CALegInfoClient:
    """Scraper for California Legislative Information (leginfo.legislature.ca.gov)."""
    BASE_URL = "https://leginfo.legislature.ca.gov/faces/codes_displaySection.xhtml"

    def parse_code_reference(self, text: str) -> Tuple[Optional[str], Optional[str]]:
        """Parse a natural language code reference into (code, section).

        Examples:
            "Civil Code section 1714" -> ("CIV", "1714")
            "Civ. Code § 1714" -> ("CIV", "1714")
            "CCP 437c" -> ("CCP", "437c")
        """
        text_lower = text.lower().strip()

        # Try matching "Code Name [section|§|sec.] number"
        for alias, code in sorted(CODE_ALIASES.items(), key=lambda x: -len(x[0])):
            if alias in text_lower:
                # Extract section number after the code name
                after_code = text_lower[text_lower.index(alias) + len(alias):]
                section_match = re.search(r'(?:section|§|sec\.?)?\s*(\d+[\w.]*)', after_code)
                if section_match:
                    return code, section_match.group(1)

        # Try matching just an abbreviation like "CCP 437c"
        abbrev_match = re.match(r'([A-Z]{2,4})\s+(?:section|§|sec\.?)?\s*(\d+[\w.]*)', text.strip())
        if abbrev_match:
            code_upper = abbrev_match.group(1).upper()
            if code_upper in CODE_TITLES:
                return code_upper, abbrev_match.group(2)

        return None, None

    def get_section(self, code: str, section: str) -> Optional[StatuteResult]:
        """Fetch a specific code section from leginfo."""
        try:
            params = {
                "sectionNum": section,
                "lawCode": code,
            }
            resp = requests.get(self.BASE_URL, params=params, timeout=30)
            resp.raise_for_status()

            soup = BeautifulSoup(resp.text, "html.parser")
            content_div = soup.find("div", {"id": "codeLawSectionNoContent"})
            if not content_div:
                return None

            text = content_div.get_text(separator="\n", strip=True)
            if not text or len(text) < 10:
                return None

            title = CODE_TITLES.get(code, code)
            url = f"{self.BASE_URL}?sectionNum={section}&lawCode={code}"
            return StatuteResult(
                code=code, section=section, title=title, text=text, url=url
            )
        except Exception as e:
            print(f"[CALegInfo] Error fetching {code} § {section}: {e}")
            return None

    def search_code(self, code: str, section: str) -> List[StatuteResult]:
        """Search for a code section. Returns a list (0 or 1 results)."""
        result = self.get_section(code, section)
        return [result] if result else []
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_legal_research/test_ca_leginfo.py -v`
Expected: All 6 tests PASS

**Step 5: Commit**

```bash
git add icharlotte_core/legal_research/sources/ca_leginfo.py tests/test_legal_research/test_ca_leginfo.py
git commit -m "feat(legal-research): add CA Legislative Info scraper for statutes"
```

---

## Task 4: CA Courts Recent Opinions Client

**Files:**
- Create: `icharlotte_core/legal_research/sources/ca_courts.py`
- Test: `tests/test_legal_research/test_ca_courts.py`

**Step 1: Write the failing test**

```python
# tests/test_legal_research/test_ca_courts.py
import unittest
from unittest.mock import patch, MagicMock
from icharlotte_core.legal_research.sources.ca_courts import CACourtsClient
from icharlotte_core.legal_research.models import CaseResult


class TestCACourtsClient(unittest.TestCase):
    def setUp(self):
        self.client = CACourtsClient()

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_search_recent_opinions(self, mock_get):
        mock_response = MagicMock()
        mock_response.status_code = 200
        mock_response.text = '''<html><body>
        <table class="opinion-list">
            <tr>
                <td><a href="/opinions/documents/S123456.PDF">Smith v. Jones</a></td>
                <td>S123456</td>
                <td>03/01/2026</td>
                <td>Supreme Court</td>
            </tr>
        </table>
        </body></html>'''
        mock_get.return_value = mock_response

        results = self.client.search_recent("premises liability")
        # May or may not match depending on content — just verify it doesn't crash
        self.assertIsInstance(results, list)

    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_handles_connection_error(self, mock_get):
        mock_get.side_effect = Exception("Connection refused")
        results = self.client.search_recent("test")
        self.assertEqual(results, [])


if __name__ == "__main__":
    unittest.main()
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_legal_research/test_ca_courts.py -v`
Expected: FAIL with "ModuleNotFoundError"

**Step 3: Write minimal implementation**

```python
# icharlotte_core/legal_research/sources/ca_courts.py
"""California Courts official website scraper for recent opinions."""
import re
import requests
from typing import List
from bs4 import BeautifulSoup
from ..models import CaseResult


class CACourtsClient:
    """Scraper for courts.ca.gov recent published/unpublished opinions."""
    OPINIONS_URL = "https://www.courts.ca.gov/opinions.htm"

    def search_recent(self, query: str, max_results: int = 10) -> List[CaseResult]:
        """Search recent CA court opinions.

        Note: courts.ca.gov doesn't have a search API, so we fetch the
        recent opinions list and do keyword matching. For deep search,
        CourtListener is the primary source.
        """
        try:
            resp = requests.get(self.OPINIONS_URL, timeout=30)
            resp.raise_for_status()
            soup = BeautifulSoup(resp.text, "html.parser")

            results = []
            query_terms = query.lower().split()

            # Look for opinion links in the page
            for link in soup.find_all("a", href=True):
                href = link.get("href", "")
                text = link.get_text(strip=True)

                # Match PDF opinion links
                if not (href.endswith(".PDF") or href.endswith(".pdf")):
                    continue
                if not text:
                    continue

                # Simple keyword match against case name
                text_lower = text.lower()
                if any(term in text_lower for term in query_terms):
                    # Try to extract date from nearby elements
                    parent_row = link.find_parent("tr")
                    date_str = ""
                    court_str = "California Court"
                    if parent_row:
                        cells = parent_row.find_all("td")
                        for cell in cells:
                            cell_text = cell.get_text(strip=True)
                            date_match = re.search(r'\d{1,2}/\d{1,2}/\d{4}', cell_text)
                            if date_match:
                                date_str = date_match.group(0)

                    full_url = href if href.startswith("http") else f"https://www.courts.ca.gov{href}"
                    results.append(CaseResult(
                        name=text,
                        citation="",  # Not always available from the listing
                        date=date_str,
                        court=court_str,
                        snippet="",
                        url=full_url,
                    ))

                    if len(results) >= max_results:
                        break

            return results
        except Exception as e:
            print(f"[CACourts] Error: {e}")
            return []
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_legal_research/test_ca_courts.py -v`
Expected: All 2 tests PASS

**Step 5: Commit**

```bash
git add icharlotte_core/legal_research/sources/ca_courts.py tests/test_legal_research/test_ca_courts.py
git commit -m "feat(legal-research): add CA Courts recent opinions scraper"
```

---

## Task 5: LLM Prompts for Query Planning, Synthesis, and Verification

**Files:**
- Create: `icharlotte_core/legal_research/prompts.py`
- Test: `tests/test_legal_research/test_prompts.py`

**Step 1: Write the failing test**

```python
# tests/test_legal_research/test_prompts.py
import unittest
from icharlotte_core.legal_research.prompts import (
    QUERY_PLANNING_PROMPT,
    SYNTHESIS_PROMPT,
    VERIFICATION_PROMPT,
    CITATION_INSTRUCTION,
    build_augmented_system_prompt,
)


class TestPrompts(unittest.TestCase):
    def test_query_planning_prompt_exists(self):
        self.assertIn("search terms", QUERY_PLANNING_PROMPT.lower())
        self.assertIn("JSON", QUERY_PLANNING_PROMPT)

    def test_synthesis_prompt_exists(self):
        self.assertIn("cite", SYNTHESIS_PROMPT.lower())

    def test_verification_prompt_exists(self):
        self.assertIn("PASS", VERIFICATION_PROMPT)
        self.assertIn("FIXED", VERIFICATION_PROMPT)
        self.assertIn("FLAGGED", VERIFICATION_PROMPT)

    def test_citation_instruction_anti_hallucination(self):
        self.assertIn("MUST ONLY cite", CITATION_INSTRUCTION)
        self.assertIn("Do NOT fabricate", CITATION_INSTRUCTION)

    def test_build_augmented_system_prompt(self):
        base = "You are a helpful assistant."
        authority = "[LEGAL AUTHORITY]\nCASE LAW:\n  - Test v. Case\n[END LEGAL AUTHORITY]"
        result = build_augmented_system_prompt(base, authority)
        self.assertIn(base, result)
        self.assertIn(authority, result)
        self.assertIn("MUST ONLY cite", result)


if __name__ == "__main__":
    unittest.main()
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_legal_research/test_prompts.py -v`
Expected: FAIL with "ModuleNotFoundError"

**Step 3: Write minimal implementation**

```python
# icharlotte_core/legal_research/prompts.py
"""LLM prompts for the legal research pipeline."""

QUERY_PLANNING_PROMPT = """\
You are a legal research assistant specializing in California law. Given a user's legal question or document context, extract structured search terms for legal databases.

Output valid JSON with this exact structure:
{
    "case_queries": ["search query 1", "search query 2"],
    "statute_queries": ["Civil Code section 1714", "CCP section 437c"],
    "legal_topics": ["premises liability", "duty of care"]
}

Rules:
- case_queries: 2-5 natural language search terms for case law databases (use legal terminology)
- statute_queries: specific California code sections referenced or likely relevant (use format "Code Name section NUMBER")
- legal_topics: 1-3 broad legal topics for categorization
- If the user mentions specific cases or statutes, include them
- Focus on California state law unless federal law is clearly needed
- Be specific — "premises liability dog bite" is better than "personal injury"
"""

SYNTHESIS_PROMPT = """\
You are a legal research assistant. Synthesize the research results below into properly cited legal analysis.

Rules:
- Cite cases in California format: Case Name (Year) Volume Reporter Page (e.g., Rowland v. Christian (1968) 69 Cal.2d 108)
- Cite statutes as: Code Name, § Section (e.g., Civ. Code, § 1714, subd. (a))
- Only cite sources that appear in the research results — do NOT invent citations
- Note the holding or relevant text for each case cited
- Flag any case with negative treatment (overruled, distinguished)
- If authority is insufficient, state "Additional research may be needed on [topic]"
"""

VERIFICATION_PROMPT = """\
You are a legal citation verification specialist. Review the draft text below and verify EVERY citation against the provided source data.

For each citation in the draft, check:
1. EXISTENCE: Does this case/statute appear in the [LEGAL AUTHORITY] source data?
2. ACCURACY: Is the case name, citation (volume, reporter, page), and year correct?
3. SUPPORT: Does the holding/text of the cited authority ACTUALLY support the proposition it's cited for?
4. TREATMENT: Is there any negative treatment noted (overruled, distinguished, superseded)?

Output valid JSON with this exact structure:
{
    "verifications": [
        {
            "citation": "Rowland v. Christian (1968) 69 Cal.2d 108",
            "status": "PASS",
            "detail": "Citation accurate, holding supports duty of care proposition"
        },
        {
            "citation": "Smith v. Jones (2020) 50 Cal.App.5th 300",
            "status": "FIXED",
            "detail": "Wrong page number",
            "original": "50 Cal.App.5th 300",
            "corrected": "50 Cal.App.5th 305"
        },
        {
            "citation": "Doe v. Roe (2015) 240 Cal.App.4th 100",
            "status": "FLAGGED",
            "detail": "Holding addresses breach of contract, not premises liability — does not support the proposition cited for"
        }
    ],
    "corrected_text": "The full draft text with all FIXED citations corrected and FLAGGED citations marked with [UNVERIFIED]"
}

CRITICAL: Be thorough. A wrong citation in a legal document is malpractice-level serious.
"""

CITATION_INSTRUCTION = (
    "You MUST ONLY cite cases and statutes from the [LEGAL AUTHORITY] section below. "
    "Do NOT fabricate or hallucinate any citations. If the provided authority is insufficient, "
    "explicitly state that additional research is needed rather than inventing citations."
)


def build_augmented_system_prompt(base_system_prompt: str, authority_block: str) -> str:
    """Combine a base system prompt with legal authority and citation instructions."""
    return (
        f"{base_system_prompt}\n\n"
        f"{CITATION_INSTRUCTION}\n\n"
        f"{authority_block}"
    )
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_legal_research/test_prompts.py -v`
Expected: All 5 tests PASS

**Step 5: Commit**

```bash
git add icharlotte_core/legal_research/prompts.py tests/test_legal_research/test_prompts.py
git commit -m "feat(legal-research): add LLM prompts for query planning, synthesis, and verification"
```

---

## Task 6: Legal Research Engine (Core Pipeline)

**Files:**
- Create: `icharlotte_core/legal_research/engine.py`
- Test: `tests/test_legal_research/test_engine.py`

**Step 1: Write the failing test**

```python
# tests/test_legal_research/test_engine.py
import json
import unittest
from unittest.mock import MagicMock, patch
from icharlotte_core.legal_research.engine import LegalResearchEngine
from icharlotte_core.legal_research.models import CaseResult, StatuteResult, ResearchResult


class TestLegalResearchEngine(unittest.TestCase):
    def setUp(self):
        self.engine = LegalResearchEngine(courtlistener_token="test-token")

    @patch.object(LegalResearchEngine, '_search_sources')
    @patch.object(LegalResearchEngine, '_plan_queries')
    def test_research_returns_result(self, mock_plan, mock_search):
        mock_plan.return_value = {
            "case_queries": ["premises liability duty of care"],
            "statute_queries": ["Civil Code section 1714"],
            "legal_topics": ["premises liability"],
        }
        mock_search.return_value = (
            [CaseResult(
                name="Rowland v. Christian", citation="69 Cal.2d 108",
                date="1968-08-08", court="Supreme Court of California",
                snippet="Duty of care", url="http://test", cluster_id=1
            )],
            [StatuteResult(
                code="CIV", section="1714", title="Civil Code",
                text="Everyone is responsible...", url="http://test"
            )],
        )

        def mock_llm(system, user):
            return "Analysis with citations."

        result = self.engine.research("premises liability", llm_callback=mock_llm)
        self.assertIsInstance(result, ResearchResult)
        self.assertEqual(len(result.cases), 1)
        self.assertEqual(len(result.statutes), 1)
        self.assertEqual(result.query, "premises liability")

    def test_research_with_no_results(self):
        # Mock all source clients to return empty
        self.engine.cl_client = MagicMock()
        self.engine.cl_client.search_opinions.return_value = []
        self.engine.leginfo_client = MagicMock()
        self.engine.leginfo_client.get_section.return_value = None
        self.engine.leginfo_client.parse_code_reference.return_value = (None, None)
        self.engine.courts_client = MagicMock()
        self.engine.courts_client.search_recent.return_value = []

        def mock_llm(system, user):
            if "search terms" in system.lower():
                return json.dumps({
                    "case_queries": ["nonexistent topic"],
                    "statute_queries": [],
                    "legal_topics": ["nothing"],
                })
            return "No relevant authority found."

        result = self.engine.research("nonexistent topic", llm_callback=mock_llm)
        self.assertIsInstance(result, ResearchResult)
        self.assertEqual(len(result.cases), 0)

    def test_cache_key_generation(self):
        key1 = self.engine._cache_key("premises liability dog bite")
        key2 = self.engine._cache_key("premises liability dog bite")
        key3 = self.engine._cache_key("different query")
        self.assertEqual(key1, key2)
        self.assertNotEqual(key1, key3)


if __name__ == "__main__":
    unittest.main()
```

**Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_legal_research/test_engine.py -v`
Expected: FAIL with "ModuleNotFoundError"

**Step 3: Write minimal implementation**

```python
# icharlotte_core/legal_research/engine.py
"""Legal Research Engine — core pipeline for searching and verifying legal authority."""
import hashlib
import json
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Callable, Dict, List, Optional, Tuple

from .models import CaseResult, ResearchResult, StatuteResult, VerificationStatus
from .prompts import (
    QUERY_PLANNING_PROMPT,
    SYNTHESIS_PROMPT,
    VERIFICATION_PROMPT,
)
from .sources.courtlistener import CourtListenerClient
from .sources.ca_leginfo import CALegInfoClient
from .sources.ca_courts import CACourtsClient

# Type alias: llm_callback(system_prompt, user_prompt) -> str
LLMCallback = Callable[[str, str], str]


class LegalResearchEngine:
    """Stateless research engine — no Qt dependency.

    Pipeline:
    1. Query planning (LLM extracts search terms from natural language)
    2. Parallel source search (CourtListener, CA Leginfo, CA Courts)
    3. Synthesis (LLM produces memo with citations)
    4. Verification (LLM cross-checks every citation against source data)
    """

    def __init__(self, courtlistener_token: str):
        self.cl_client = CourtListenerClient(token=courtlistener_token)
        self.leginfo_client = CALegInfoClient()
        self.courts_client = CACourtsClient()

    def research(
        self,
        query: str,
        llm_callback: LLMCallback,
        status_callback: Optional[Callable[[str], None]] = None,
    ) -> ResearchResult:
        """Run the full research pipeline.

        Args:
            query: Natural language legal question or document context
            llm_callback: Function(system_prompt, user_prompt) -> str for LLM calls
            status_callback: Optional function(status_text) for progress updates
        """
        def status(msg: str):
            if status_callback:
                status_callback(msg)

        # Phase 1: Query Planning
        status("Analyzing legal question...")
        search_plan = self._plan_queries(query, llm_callback)

        # Phase 2: Parallel Source Search
        status("Searching legal databases...")
        cases, statutes = self._search_sources(search_plan)

        # Phase 3: Synthesis
        status("Synthesizing legal analysis...")
        result = ResearchResult(query=query, cases=cases, statutes=statutes)
        authority_block = result.format_authority_block()
        memo = llm_callback(
            SYNTHESIS_PROMPT,
            f"User's legal question: {query}\n\n{authority_block}"
        )
        result.memo = memo

        # Phase 4: Verification
        if cases or statutes:
            status("Verifying citations...")
            result.verification = self._verify_citations(
                memo, authority_block, llm_callback
            )
            # Apply corrections from verification
            corrected = self._apply_corrections(result.verification, memo)
            if corrected:
                result.memo = corrected

        status("Research complete")
        return result

    def _plan_queries(self, query: str, llm_callback: LLMCallback) -> Dict:
        """Use LLM to extract structured search terms from natural language."""
        raw = llm_callback(QUERY_PLANNING_PROMPT, query)
        try:
            # Strip markdown code fences if present
            cleaned = raw.strip()
            if cleaned.startswith("```"):
                cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
                if cleaned.endswith("```"):
                    cleaned = cleaned[:-3]
                cleaned = cleaned.strip()
            return json.loads(cleaned)
        except (json.JSONDecodeError, ValueError):
            # Fallback: use the raw query as a single search term
            return {
                "case_queries": [query],
                "statute_queries": [],
                "legal_topics": [],
            }

    def _search_sources(
        self, plan: Dict
    ) -> Tuple[List[CaseResult], List[StatuteResult]]:
        """Search all sources in parallel based on the query plan."""
        cases: List[CaseResult] = []
        statutes: List[StatuteResult] = []

        with ThreadPoolExecutor(max_workers=5) as pool:
            futures = {}

            # CourtListener searches
            for q in plan.get("case_queries", []):
                f = pool.submit(self.cl_client.search_opinions, q)
                futures[f] = ("case", q)

            # CA Courts recent opinions
            for q in plan.get("case_queries", [])[:2]:  # Limit to first 2 queries
                f = pool.submit(self.courts_client.search_recent, q)
                futures[f] = ("recent_case", q)

            # Statute lookups
            for ref in plan.get("statute_queries", []):
                code, section = self.leginfo_client.parse_code_reference(ref)
                if code and section:
                    f = pool.submit(self.leginfo_client.get_section, code, section)
                    futures[f] = ("statute", ref)

            # Collect results
            for future in as_completed(futures):
                source_type, _ = futures[future]
                try:
                    result = future.result()
                    if source_type in ("case", "recent_case"):
                        if isinstance(result, list):
                            cases.extend(result)
                    elif source_type == "statute":
                        if result is not None:
                            statutes.append(result)
                except Exception as e:
                    print(f"[LegalResearch] Source error: {e}")

        # Deduplicate cases by name+citation
        seen = set()
        unique_cases = []
        for c in cases:
            key = (c.name.lower(), c.citation.lower())
            if key not in seen:
                seen.add(key)
                unique_cases.append(c)

        return unique_cases, statutes

    def _verify_citations(
        self,
        draft: str,
        authority_block: str,
        llm_callback: LLMCallback,
    ) -> List[VerificationStatus]:
        """Verify every citation in the draft against source data."""
        prompt = (
            f"DRAFT TEXT:\n{draft}\n\n"
            f"SOURCE DATA:\n{authority_block}\n\n"
            f"Verify every citation in the draft."
        )
        raw = llm_callback(VERIFICATION_PROMPT, prompt)
        try:
            cleaned = raw.strip()
            if cleaned.startswith("```"):
                cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
                if cleaned.endswith("```"):
                    cleaned = cleaned[:-3]
                cleaned = cleaned.strip()
            data = json.loads(cleaned)
            verifications = []
            for v in data.get("verifications", []):
                verifications.append(VerificationStatus(
                    citation=v.get("citation", ""),
                    status=v.get("status", "FLAGGED"),
                    detail=v.get("detail", ""),
                    original=v.get("original", ""),
                    corrected=v.get("corrected", ""),
                ))
            return verifications
        except (json.JSONDecodeError, ValueError):
            return []

    def _apply_corrections(
        self, verifications: List[VerificationStatus], draft: str
    ) -> Optional[str]:
        """Apply FIXED corrections and mark FLAGGED citations in the draft."""
        corrected = draft
        had_changes = False
        for v in verifications:
            if v.status == "FIXED" and v.original and v.corrected:
                if v.original in corrected:
                    corrected = corrected.replace(v.original, v.corrected, 1)
                    had_changes = True
            elif v.status == "FLAGGED" and v.citation:
                # Mark flagged citations
                if v.citation in corrected:
                    corrected = corrected.replace(
                        v.citation, f"{v.citation} [UNVERIFIED]", 1
                    )
                    had_changes = True
        return corrected if had_changes else None

    @staticmethod
    def _cache_key(query: str) -> str:
        """Generate a cache key for a query."""
        normalized = query.strip().lower()
        return f"legal_research_{hashlib.md5(normalized.encode()).hexdigest()[:12]}"
```

**Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_legal_research/test_engine.py -v`
Expected: All 3 tests PASS

**Step 5: Commit**

```bash
git add icharlotte_core/legal_research/engine.py tests/test_legal_research/test_engine.py
git commit -m "feat(legal-research): add core research engine with 4-phase pipeline"
```

---

## Task 7: Win+V Word Popup Integration

**Files:**
- Modify: `icharlotte_core/word_hotkey.py:1577-1579` (add checkbox after redline checkbox)
- Modify: `icharlotte_core/word_hotkey.py:2373-2390` (intercept before LLM call)
- Modify: `icharlotte_core/word_hotkey.py:1940` (checkbox visibility)

**Step 1: Add the "Perform Legal Research" checkbox**

In `icharlotte_core/word_hotkey.py`, after the redline checkbox (line 1577), add:

```python
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
```

**Step 2: Add preference persistence method**

After `_save_redline_preference` method, add:

```python
    def _save_legal_research_preference(self, state):
        """Persist legal research checkbox state."""
        self.redline_settings["legal_research_default"] = bool(state)
        self._save_redline_settings()
```

**Step 3: Update checkbox visibility**

In `_update_redline_checkbox_visibility` (line 1940), add:

```python
        self.legal_research_checkbox.setVisible(is_word)
```

**Step 4: Add the research engine import and initialization**

At the top of `_do_execute` (line 2289), the method needs access to the engine. Add a lazy-init helper to the class:

```python
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
```

**Step 5: Intercept in _do_execute before LLM call**

In `_do_execute`, after the `full_prompt` is assembled and attachment context is appended (after line 2376), but BEFORE the LLMWorkerThread is created (line 2385), insert:

```python
            # Legal Research: if checkbox is checked, run research and augment prompt
            if self.legal_research_checkbox.isChecked():
                engine = self._get_legal_research_engine()
                if engine:
                    self.status_label.setText("Researching legal authority...")
                    QApplication.processEvents()

                    from icharlotte_core.legal_research.prompts import build_augmented_system_prompt

                    def _llm_for_research(system_prompt, user_prompt):
                        """Synchronous LLM call for research sub-steps."""
                        from icharlotte_core.llm import LLMHandler
                        provider, model_id = self._get_selected_model()
                        settings = {
                            'temperature': 0.3,  # Lower temp for structured output
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

                    try:
                        research_result = engine.research(
                            query=full_prompt,
                            llm_callback=_llm_for_research,
                            status_callback=lambda msg: (
                                self.status_label.setText(msg),
                                QApplication.processEvents(),
                            ),
                        )
                        # Augment the system prompt with legal authority
                        authority_block = research_result.format_authority_block()
                        # Store for result display
                        self._pending_research_result = research_result
                    except Exception as e:
                        print(f"[LegalResearch] Engine error: {e}")
                        self._pending_research_result = None
                        authority_block = ""

                    if authority_block:
                        full_prompt = f"{full_prompt}\n\n{authority_block}\n\nIMPORTANT: You MUST ONLY cite cases and statutes from the [LEGAL AUTHORITY] section above. Do NOT fabricate or hallucinate any citations."

                    self.status_label.setText("Generating response with legal citations...")
                    QApplication.processEvents()
                else:
                    self._pending_research_result = None
                    print("[LegalResearch] No COURTLISTENER_API_TOKEN in .env — skipping research")
            else:
                self._pending_research_result = None
```

**Step 6: Show verification report in _on_llm_result**

In `_on_llm_result` (line 2398), after the result is processed successfully, add a verification summary to the status:

```python
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
                self.status_label.setText(f"Citations: {', '.join(summary_parts)}")
            self._pending_research_result = None
```

**Step 7: Add COURTLISTENER_API_TOKEN to .env**

Add to `.env`:
```
COURTLISTENER_API_TOKEN=your_courtlistener_token_here
```

**Step 8: Manual test**

1. Set `COURTLISTENER_API_TOKEN` in `.env` with your real token
2. Run `python iCharlotte.py`
3. Open a Word document, press Win+V
4. Check the "Perform Legal Research" checkbox
5. Type: "Draft a summary of California premises liability law, cite to legal authority"
6. Click Execute
7. Verify: status shows research phases, response includes real citations, verification summary appears
8. Uncheck the checkbox, try again — should work normally without research

**Step 9: Commit**

```bash
git add icharlotte_core/word_hotkey.py .env
git commit -m "feat(legal-research): integrate research engine into Win+V Word popup"
```

---

## Task 8: ChatTab Integration

**Files:**
- Modify: `icharlotte_core/ui/tabs.py:302-304` (add checkbox in toolbar)
- Modify: `icharlotte_core/ui/tabs.py:1038-1123` (intercept send_message)
- Modify: `icharlotte_core/ui/tabs.py:1138-1186` (show sources in finalize_response)

**Step 1: Add checkbox to ChatTab toolbar**

In `icharlotte_core/ui/tabs.py`, in the toolbar area after the Templates button (after line 302), add:

```python
        # Legal Research checkbox
        self.legal_research_check = QCheckBox("Legal Research")
        self.legal_research_check.setStyleSheet("font-size: 11px;")
        self.legal_research_check.setToolTip(
            "Search CA case law and statutes, inject verified citations into response"
        )
        toolbar_layout.addWidget(self.legal_research_check)
```

**Step 2: Add research engine lazy initializer to ChatTab**

Add method to ChatTab class:

```python
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
```

**Step 3: Modify send_message to intercept when checkbox is checked**

In `send_message()`, after `file_content = self.read_files_content()` (around line 1071) and before the LLMWorker is created (line 1110), add the research intercept:

```python
        # Legal Research: if checked, run research before LLM call
        research_result = None
        if self.legal_research_check.isChecked():
            engine = self._get_legal_research_engine()
            if engine:
                self.chat_history.append("<i>🔍 Researching legal authority...</i>")
                QApplication.processEvents()

                from icharlotte_core.legal_research.prompts import build_augmented_system_prompt

                research_query = user_text
                if file_content:
                    research_query += "\n\nContext:\n" + file_content[:2000]

                def _llm_for_research(system_prompt, user_prompt):
                    from icharlotte_core.llm import LLMHandler
                    return LLMHandler.generate(
                        provider=self.provider_combo.currentText(),
                        model=self.model_combo.currentText(),
                        system_prompt=system_prompt,
                        user_prompt=user_prompt,
                        file_contents="",
                        settings={**self.settings, 'stream': False, 'temperature': 0.3},
                    )

                try:
                    research_result = engine.research(
                        query=research_query,
                        llm_callback=_llm_for_research,
                        status_callback=lambda msg: (
                            self.chat_history.append(f"<i>  {msg}</i>"),
                            QApplication.processEvents(),
                        ),
                    )
                except Exception as e:
                    self.chat_history.append(f"<font color='orange'>Research error: {e}</font>")

        # Store for use in finalize_response
        self._pending_research = research_result
```

Then, when creating the LLMWorker, modify the `system_prompt` if research was done:

```python
        # Build system prompt (augment with legal authority if research was done)
        effective_system_prompt = self.system_prompt
        if research_result:
            from icharlotte_core.legal_research.prompts import build_augmented_system_prompt
            authority = research_result.format_authority_block()
            effective_system_prompt = build_augmented_system_prompt(
                self.system_prompt, authority
            )
```

And use `effective_system_prompt` instead of `self.system_prompt` in the LLMWorker constructor.

**Step 4: Show collapsible sources in finalize_response**

In `finalize_response()` (line 1146), after the HTML is inserted (line 1164), add:

```python
        # Show legal research sources if available
        if hasattr(self, '_pending_research') and self._pending_research:
            rr = self._pending_research
            sources_html = "<details><summary><b>📚 Legal Sources Found</b> "
            if rr.verification:
                pass_count = sum(1 for v in rr.verification if v.status == "PASS")
                fixed_count = sum(1 for v in rr.verification if v.status == "FIXED")
                flagged_count = sum(1 for v in rr.verification if v.status == "FLAGGED")
                parts = []
                if pass_count: parts.append(f"✓{pass_count}")
                if fixed_count: parts.append(f"⚠{fixed_count}")
                if flagged_count: parts.append(f"✗{flagged_count}")
                sources_html += f"({', '.join(parts)})"
            sources_html += "</summary><ul>"
            for c in rr.cases:
                sources_html += f"<li><b>{c.formatted_citation}</b>"
                if c.url:
                    sources_html += f' — <a href="{c.url}">View</a>'
                sources_html += "</li>"
            for s in rr.statutes:
                sources_html += f"<li><b>{s.formatted_citation}</b>"
                if s.url:
                    sources_html += f' — <a href="{s.url}">View</a>'
                sources_html += "</li>"
            sources_html += "</ul></details>"
            self.chat_history.append(sources_html)
            self._pending_research = None
```

**Note:** QTextBrowser may not support `<details>` HTML5 tag. If it doesn't render correctly, use a simpler format with a separator line and list. Test and adjust.

**Step 5: Manual test**

1. Run `python iCharlotte.py`
2. Open the Chat tab, load a case
3. Check the "Legal Research" checkbox in the toolbar
4. Type: "What is the standard for premises liability in California? Cite case law."
5. Verify: research status messages appear, response includes real citations, sources section shows below

**Step 6: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(legal-research): integrate research engine into ChatTab with sources display"
```

---

## Task 9: Caching via CaseDataManager

**Files:**
- Modify: `icharlotte_core/legal_research/engine.py` (add cache check/save)

**Step 1: Add caching to the engine's research method**

Modify `engine.py` to accept an optional `file_number` and `data_manager`:

```python
    def research(
        self,
        query: str,
        llm_callback: LLMCallback,
        status_callback: Optional[Callable[[str], None]] = None,
        file_number: Optional[str] = None,
        data_manager=None,
    ) -> ResearchResult:
        """Run the full research pipeline with optional caching."""
        def status(msg: str):
            if status_callback:
                status_callback(msg)

        # Check cache
        cache_key = self._cache_key(query)
        if file_number and data_manager:
            cached = data_manager.get_value(file_number, cache_key)
            if cached:
                status("Using cached research results")
                return self._from_cache(cached)

        # ... existing pipeline code ...

        # Save to cache
        if file_number and data_manager:
            data_manager.save_variable(
                file_number, cache_key, result.to_dict(),
                source="legal_research_agent", auto_tag=True,
            )

        return result

    @staticmethod
    def _from_cache(data: dict) -> ResearchResult:
        """Reconstruct ResearchResult from cached dict."""
        cases = [
            CaseResult(**c) for c in data.get("cases", [])
        ]
        statutes = [
            StatuteResult(**s) for s in data.get("statutes", [])
        ]
        verification = [
            VerificationStatus(
                citation=v.get("citation", ""),
                status=v.get("status", ""),
                detail=v.get("detail", ""),
            )
            for v in data.get("verification", [])
        ]
        return ResearchResult(
            query=data.get("query", ""),
            cases=cases,
            statutes=statutes,
            memo=data.get("memo", ""),
            verification=verification,
        )
```

**Step 2: Pass file_number and data_manager from integration points**

In the Win+V popup integration (Task 7), pass the active file number if available.
In the ChatTab integration (Task 8), pass `self.case_data_manager` and current file number.

**Step 3: Commit**

```bash
git add icharlotte_core/legal_research/engine.py
git commit -m "feat(legal-research): add per-case caching via CaseDataManager"
```

---

## Task 10: End-to-End Integration Test

**Files:**
- Create: `tests/test_legal_research/test_integration.py`

**Step 1: Write integration test with mocked API responses**

```python
# tests/test_legal_research/test_integration.py
"""End-to-end test of the legal research pipeline with mocked external calls."""
import json
import unittest
from unittest.mock import patch, MagicMock
from icharlotte_core.legal_research.engine import LegalResearchEngine
from icharlotte_core.legal_research.models import ResearchResult


class TestLegalResearchIntegration(unittest.TestCase):
    """Full pipeline test with mocked API + LLM calls."""

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    @patch("icharlotte_core.legal_research.sources.ca_leginfo.requests.get")
    @patch("icharlotte_core.legal_research.sources.ca_courts.requests.get")
    def test_full_pipeline(self, mock_courts_get, mock_leginfo_get, mock_cl_get):
        # Mock CourtListener response
        cl_response = MagicMock()
        cl_response.status_code = 200
        cl_response.json.return_value = {
            "count": 1,
            "results": [{
                "caseName": "Rowland v. Christian",
                "citation": ["69 Cal.2d 108"],
                "dateFiled": "1968-08-08",
                "court": "Supreme Court of California",
                "snippet": "Owner of premises owes duty of care to all persons",
                "cluster_id": 123,
                "absolute_url": "/opinion/123/rowland-v-christian/",
            }],
        }

        # Mock CA Leginfo response
        leginfo_response = MagicMock()
        leginfo_response.status_code = 200
        leginfo_response.text = '''<html><body>
        <div id="codeLawSectionNoContent">
            <p>Everyone is responsible, not only for the result of his or her willful acts,
            but also for an injury occasioned to another by his or her want of ordinary care
            or skill in the management of his or her property or person.</p>
        </div>
        </body></html>'''

        # Mock CA Courts response (empty - no recent matches)
        courts_response = MagicMock()
        courts_response.status_code = 200
        courts_response.text = "<html><body>No recent opinions</body></html>"

        # Route mocks by URL
        def route_get(url, **kwargs):
            if "courtlistener.com" in url:
                return cl_response
            elif "leginfo.legislature" in url:
                return leginfo_response
            else:
                return courts_response

        mock_cl_get.side_effect = route_get
        mock_leginfo_get.side_effect = route_get
        mock_courts_get.side_effect = route_get

        # Mock LLM calls
        call_count = [0]
        def mock_llm(system_prompt, user_prompt):
            call_count[0] += 1
            if call_count[0] == 1:
                # Query planning
                return json.dumps({
                    "case_queries": ["premises liability duty of care California"],
                    "statute_queries": ["Civil Code section 1714"],
                    "legal_topics": ["premises liability"],
                })
            elif call_count[0] == 2:
                # Synthesis
                return (
                    "Under California law, a property owner owes a duty of care to all "
                    "persons on the premises. (Rowland v. Christian (1968) 69 Cal.2d 108.) "
                    "This duty is codified in Civ. Code, § 1714, subd. (a)."
                )
            elif call_count[0] == 3:
                # Verification
                return json.dumps({
                    "verifications": [
                        {
                            "citation": "Rowland v. Christian (1968) 69 Cal.2d 108",
                            "status": "PASS",
                            "detail": "Citation accurate, holding supports duty of care"
                        }
                    ],
                    "corrected_text": None,
                })
            return ""

        engine = LegalResearchEngine(courtlistener_token="test-token")
        result = engine.research(
            query="What duty does a property owner owe to visitors?",
            llm_callback=mock_llm,
        )

        # Verify pipeline output
        self.assertIsInstance(result, ResearchResult)
        self.assertGreaterEqual(len(result.cases), 1)
        self.assertEqual(result.cases[0].name, "Rowland v. Christian")
        self.assertIn("Rowland", result.memo)
        self.assertGreaterEqual(len(result.verification), 1)
        self.assertEqual(result.verification[0].status, "PASS")

        # Verify authority block formatting
        block = result.format_authority_block()
        self.assertIn("[LEGAL AUTHORITY]", block)
        self.assertIn("Rowland v. Christian", block)


if __name__ == "__main__":
    unittest.main()
```

**Step 2: Run test**

Run: `python -m pytest tests/test_legal_research/ -v`
Expected: All tests across all files PASS

**Step 3: Commit**

```bash
git add tests/test_legal_research/test_integration.py
git commit -m "test(legal-research): add end-to-end integration test with mocked sources"
```

---

## Task 11: Live API Smoke Test

**Files:** None — manual testing only

**Step 1: Verify CourtListener API token works**

Run from the project root:
```bash
python -c "
import os
from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
token = os.environ.get('COURTLISTENER_API_TOKEN', '')
print(f'Token: {token[:8]}...' if token else 'NO TOKEN SET')
client = CourtListenerClient(token=token)
results = client.search_opinions('premises liability', max_results=3)
for r in results:
    print(f'  {r.formatted_citation}')
    print(f'    {r.snippet[:100]}...')
print(f'Total: {len(results)} results')
"
```
Expected: 1-3 California case results with real citations

**Step 2: Verify CA Leginfo works**

```bash
python -c "
from icharlotte_core.legal_research.sources.ca_leginfo import CALegInfoClient
client = CALegInfoClient()
result = client.get_section('CIV', '1714')
if result:
    print(f'{result.formatted_citation}')
    print(f'{result.text[:200]}...')
else:
    print('FAILED: No result for CIV 1714')
"
```
Expected: Civil Code § 1714 text about responsibility for willful acts

**Step 3: Full pipeline smoke test**

```bash
python -c "
import os
from icharlotte_core.legal_research.engine import LegalResearchEngine
from icharlotte_core.llm import LLMHandler
from icharlotte_core.config import API_KEYS

engine = LegalResearchEngine(courtlistener_token=os.environ.get('COURTLISTENER_API_TOKEN', ''))

def llm_call(system_prompt, user_prompt):
    return LLMHandler.generate(
        provider='Gemini', model='gemini-2.0-flash',
        system_prompt=system_prompt, user_prompt=user_prompt,
        file_contents='', settings={'temperature': 0.3, 'stream': False, 'thinking_level': 'None'}
    )

result = engine.research(
    'What is the standard for premises liability in California?',
    llm_callback=llm_call,
    status_callback=print,
)
print('\\n=== CASES ===')
for c in result.cases[:5]:
    print(f'  {c.formatted_citation}')
print('\\n=== STATUTES ===')
for s in result.statutes:
    print(f'  {s.formatted_citation}')
print('\\n=== VERIFICATION ===')
for v in result.verification:
    print(f'  [{v.status}] {v.citation}: {v.detail}')
print('\\n=== MEMO ===')
print(result.memo[:500])
"
```
Expected: Real citations, real statute text, verification results

**Step 4: Commit (final)**

```bash
git add -A
git commit -m "feat(legal-research): complete legal research agent with CourtListener, CA Leginfo, and verification pipeline"
```

---

## Summary

| Task | Component | Files Created/Modified |
|------|-----------|-----------------------|
| 1 | Data Models | `models.py`, `__init__.py`, test |
| 2 | CourtListener Client | `courtlistener.py`, test |
| 3 | CA Leginfo Client | `ca_leginfo.py`, test |
| 4 | CA Courts Client | `ca_courts.py`, test |
| 5 | LLM Prompts | `prompts.py`, test |
| 6 | Research Engine | `engine.py`, test |
| 7 | Win+V Integration | `word_hotkey.py` modified |
| 8 | ChatTab Integration | `tabs.py` modified |
| 9 | Caching | `engine.py` modified |
| 10 | Integration Test | test file |
| 11 | Live Smoke Test | manual verification |
