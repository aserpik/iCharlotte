"""Data models for the legal research agent."""
from dataclasses import dataclass, field
from typing import List, Optional, Dict, Any


# Mapping of California code abbreviations to display names
CODE_DISPLAY_NAMES: Dict[str, str] = {
    "BPC": "Bus. & Prof. Code",
    "CCP": "Code Civ. Proc.",
    "CIV": "Civ. Code",
    "COM": "Com. Code",
    "CORP": "Corp. Code",
    "EDC": "Educ. Code",
    "ELEC": "Elec. Code",
    "EVID": "Evid. Code",
    "FAM": "Fam. Code",
    "FIN": "Fin. Code",
    "FGC": "Fish & Game Code",
    "FCA": "Food & Agric. Code",
    "GOV": "Gov. Code",
    "HNC": "Harb. & Nav. Code",
    "HSC": "Health & Safety Code",
    "INS": "Ins. Code",
    "LAB": "Lab. Code",
    "MVC": "Mil. & Vet. Code",
    "PEN": "Pen. Code",
    "PROB": "Prob. Code",
    "PCC": "Pub. Contract Code",
    "PRC": "Pub. Resources Code",
    "PUC": "Pub. Util. Code",
    "RTC": "Rev. & Tax. Code",
    "SHC": "Sts. & Hy. Code",
    "UIC": "Unemp. Ins. Code",
    "VEH": "Veh. Code",
    "WAT": "Water Code",
    "WIC": "Welf. & Inst. Code",
}


@dataclass
class CaseResult:
    """A case law search result."""
    name: str
    citation: str
    date: str
    court: str
    snippet: str
    url: str
    cluster_id: Optional[int] = None
    negative_treatment: Optional[str] = None
    relevance_score: float = 0.0
    snippet_source: str = ""
    snippet_parenthetical_id: str = ""

    @property
    def formatted_citation(self) -> str:
        """Return citation in standard format: Name (Year) Citation."""
        year = self.date[:4] if len(self.date) >= 4 else self.date
        return f"{self.name} ({year}) {self.citation}"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "citation": self.citation,
            "date": self.date,
            "court": self.court,
            "snippet": self.snippet,
            "url": self.url,
            "cluster_id": self.cluster_id,
            "negative_treatment": self.negative_treatment,
            "relevance_score": self.relevance_score,
            "snippet_source": self.snippet_source,
            "snippet_parenthetical_id": self.snippet_parenthetical_id,
        }


@dataclass
class StatuteResult:
    """A statute search result."""
    code: str
    section: str
    title: str
    text: str
    url: str

    @property
    def formatted_citation(self) -> str:
        """Return citation in standard format: Display Name, section symbol Section."""
        display = CODE_DISPLAY_NAMES.get(self.code, self.code)
        return f"{display}, \u00a7 {self.section}"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "code": self.code,
            "section": self.section,
            "title": self.title,
            "text": self.text,
            "url": self.url,
        }


@dataclass
class VerificationStatus:
    """Verification result for a single citation."""
    citation: str
    status: str  # "PASS", "FIXED", "FLAGGED"
    detail: str
    original: Optional[str] = None
    corrected: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "citation": self.citation,
            "status": self.status,
            "detail": self.detail,
            "original": self.original,
            "corrected": self.corrected,
        }


@dataclass
class ResearchResult:
    """Complete result from a legal research query."""
    query: str
    cases: List[CaseResult] = field(default_factory=list)
    statutes: List[StatuteResult] = field(default_factory=list)
    memo: str = ""
    verification: List[VerificationStatus] = field(default_factory=list)
    requested_cases: List[str] = field(default_factory=list)
    unfound_cases: List[str] = field(default_factory=list)

    def format_authority_block(self) -> str:
        """Format results as a [LEGAL AUTHORITY] block for insertion into documents.

        Includes case holdings/snippets and statute text so the LLM can
        cite specific propositions from the authorities.
        """
        lines = ["[LEGAL AUTHORITY]"]

        if self.cases:
            lines.append("")
            lines.append("Case Law:")
            for case in self.cases:
                entry = f"  - {case.formatted_citation}"
                if case.negative_treatment:
                    entry += f" [NOTE: {case.negative_treatment}]"
                lines.append(entry)
                # Include holding/snippet so LLM can cite specific propositions
                if case.snippet:
                    snippet_text = case.snippet.strip()
                    if snippet_text:
                        if case.snippet_source == "parenthetical":
                            lines.append(f"    Secondary parenthetical: {snippet_text}")
                        else:
                            lines.append(f"    Holding: {snippet_text}")

        if self.statutes:
            lines.append("")
            lines.append("Statutes:")
            for statute in self.statutes:
                lines.append(f"  - {statute.formatted_citation}")
                # Include statute text so LLM can cite specific language
                if statute.text:
                    stat_text = statute.text[:4000].strip()
                    if stat_text:
                        lines.append(f"    Text: {stat_text}")

        if self.unfound_cases:
            lines.append("")
            lines.append("UNFOUND CASES (DO NOT CITE):")
            for case_name in self.unfound_cases:
                lines.append(
                    f"  *** WARNING: '{case_name}' was NOT FOUND in any legal "
                    f"database. Do NOT cite this case. Do NOT fabricate a "
                    f"citation for it. If the user asked you to cite this case, "
                    f"you MUST state that it could not be verified. ***"
                )

        if self.verification:
            flagged = [v for v in self.verification if v.status == "FLAGGED"]
            if flagged:
                lines.append("")
                lines.append("Verification Warnings:")
                for v in flagged:
                    lines.append(f"  - {v.citation}: {v.detail}")

        lines.append("")
        lines.append("[/LEGAL AUTHORITY]")
        return "\n".join(lines)

    def get_known_case_names(self) -> List[str]:
        """Return a list of all case names from the research results."""
        return [c.name for c in self.cases]

    def to_dict(self) -> Dict[str, Any]:
        """Return a JSON-serializable dictionary."""
        return {
            "query": self.query,
            "cases": [c.to_dict() for c in self.cases],
            "statutes": [s.to_dict() for s in self.statutes],
            "memo": self.memo,
            "verification": [v.to_dict() for v in self.verification],
            "requested_cases": self.requested_cases,
            "unfound_cases": self.unfound_cases,
        }
