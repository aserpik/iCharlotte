"""
Phase 1: Discovery Response Parser.

Parses incoming discovery PDFs into structured ParsedDiscovery/ParsedRequest
objects. Uses LLM for extraction and rule-based post-processing for compound
detection and definition term extraction.
"""
import json
import re
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class ParsedRequest:
    """A single parsed discovery request."""
    number: str
    text: str
    definitions: List[str] = field(default_factory=list)
    is_compound: bool = False
    defined_terms_used: List[str] = field(default_factory=list)


@dataclass
class ParsedDiscovery:
    """Structured result of parsing an incoming discovery document."""
    discovery_type: str       # "FI", "SI", "RFA", "RPD"
    propounding_party: str
    responding_party: str
    set_number: int
    set_word: str             # "ONE", "TWO", etc.
    case_number: str
    requests: List[ParsedRequest] = field(default_factory=list)


# Discovery type detection
_TYPE_PATTERNS = {
    "FI": re.compile(r"FORM\s+INTERROGATOR", re.IGNORECASE),
    "SI": re.compile(r"SPECIAL\s+INTERROGATOR", re.IGNORECASE),
    "RFA": re.compile(r"REQUEST\s+FOR\s+ADMISSION", re.IGNORECASE),
    "RPD": re.compile(r"REQUEST\s+FOR\s+PRODUCTION", re.IGNORECASE),
}

def detect_discovery_type(text: str) -> Optional[str]:
    for dtype, pattern in _TYPE_PATTERNS.items():
        if pattern.search(text):
            return dtype
    return None


# Compound question detection
_COMPOUND_PATTERN = re.compile(
    r'\b(state|identify|describe|list|explain|set forth)\b.*?\bAND\b.*?'
    r'\b(state|identify|describe|list|explain|set forth)\b',
    re.IGNORECASE | re.DOTALL,
)

def detect_compound(text: str) -> bool:
    return bool(_COMPOUND_PATTERN.search(text))


# Defined term extraction — ALL-CAPS words 3+ chars, excluding stopwords
_CAPS_STOPWORDS = frozenset({
    "A", "AN", "AND", "ARE", "AS", "AT", "BE", "BY", "DO", "FOR", "FROM",
    "HAS", "HIS", "HER", "IF", "IN", "IS", "IT", "NO", "NOT", "OF", "ON",
    "OR", "SO", "THE", "TO", "WAS", "SET", "ONE", "TWO", "THREE",
    "INTERROGATORY", "INTERROGATORIES", "REQUEST", "ADMISSION", "PRODUCTION",
    "RESPONSE", "SPECIAL", "FORM", "ADMIT", "DENY", "STATE", "DESCRIBE",
    "IDENTIFY", "ALL", "EACH", "ANY", "THAT", "THIS", "WHICH", "WHAT",
    "WHEN", "WHERE", "WHO", "HOW", "DOES", "DID", "WERE",
})

def extract_defined_terms(text: str) -> List[str]:
    words = re.findall(r'\b([A-Z]{3,})\b', text)
    seen = set()
    result = []
    for w in words:
        if w not in _CAPS_STOPWORDS and w not in seen:
            seen.add(w)
            result.append(w)
    return result


# LLM prompt for Phase 1 parsing
def build_parse_prompt(document_text: str) -> str:
    return f"""You are a legal document parser. Extract structured data from this discovery document.

Return a JSON object with these fields:
- "discovery_type": the discovery type — one of "FI" (Form Interrogatories), "SI" (Special Interrogatories), "RFA" (Requests for Admission), "RPD" (Requests for Production)
- "propounding_party": the full name of the propounding party (e.g., "Plaintiff JOHN DOE")
- "set_number": integer (1, 2, etc.)
- "case_number": the case number (e.g., "23STCV12345")
- "requests": array of objects, each with:
  - "number": string (e.g., "1.1" for form interrogatories, "1" for others)
  - "text": the full text of the request/interrogatory
  - "definitions": array of any inline definition footnotes associated with this request

Return ONLY the JSON object, no other text. If a field cannot be determined, use null.

DOCUMENT TEXT:
{document_text}"""


# LLM response parsing
_SET_WORDS = {
    1: "ONE", 2: "TWO", 3: "THREE", 4: "FOUR", 5: "FIVE",
    6: "SIX", 7: "SEVEN", 8: "EIGHT", 9: "NINE", 10: "TEN",
}

def parse_llm_response(llm_json: str, our_client_name: str) -> ParsedDiscovery:
    text = llm_json.strip()
    if text.startswith("```"):
        text = re.sub(r'^```(?:json)?\s*', '', text)
        text = re.sub(r'\s*```$', '', text)

    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise ValueError(f"Failed to parse LLM response as JSON: {e}") from e

    set_number = int(data.get("set_number", 1) or 1)
    set_word = _SET_WORDS.get(set_number, str(set_number))

    requests = []
    for req_data in data.get("requests", []):
        req_text = req_data.get("text", "")
        req = ParsedRequest(
            number=str(req_data.get("number", "")),
            text=req_text,
            definitions=req_data.get("definitions", []) or [],
            is_compound=detect_compound(req_text),
            defined_terms_used=extract_defined_terms(req_text),
        )
        requests.append(req)

    return ParsedDiscovery(
        discovery_type=data.get("discovery_type", ""),
        propounding_party=data.get("propounding_party", ""),
        responding_party=our_client_name,
        set_number=set_number,
        set_word=set_word,
        case_number=data.get("case_number", ""),
        requests=requests,
    )
