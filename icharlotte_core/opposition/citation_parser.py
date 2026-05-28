"""Parse case and statute citations out of a drafted opposition body.

The parser identifies citation kinds (case / statute / rule / unknown), extracts
the surrounding 1-2 sentences as the brief's proposition, and computes a stable
``normalized`` form for use as a cache key.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field


@dataclass
class Citation:
    kind: str = "unknown"
    raw_text: str = ""
    normalized: str = ""
    proposition: str = ""
    body_offset: int = 0

    # Case-specific.
    case_name: str = ""
    reporter_citation: str = ""
    year: str = ""

    # Statute-specific.
    law_code: str = ""
    section_num: str = ""


# ---------------------------------------------------------------------------
# Case-cite regex
# ---------------------------------------------------------------------------

# California reporter tokens. Order matters: longer / more specific first
# (Cal.App.Supp. must come before Cal.App. or the suffix is swallowed).
_REPORTER_PATTERN = (
    r"Cal\.\s*App\.\s*Supp\.\s*(?:2d|3d|4th|5th)?"
    r"|Cal\.\s*App\.\s*(?:2d|3d|4th|5th|6th)?"
    r"|Cal\.\s*Rptr\.\s*(?:2d|3d)?"
    r"|Cal\.\s*(?:2d|3d|4th|5th|6th)?"
    r"|P\.\s*(?:2d|3d)"
)

# Case name: Two capitalized phrases separated by " v. " or " vs. ". May be
# wrapped in *...* or _..._ italic markers (one OR two characters - bold
# uses **...** or __...__ ). We capture the inner name without markers.
# Allows hyphens, apostrophes, and dots inside individual words. Word
# tokens may be joined by:
#   - whitespace alone ("Smith Jones")
#   - comma + whitespace ("Sinaiko, Inc.")
#   - ampersand + whitespace ("Goldberg & Bagula")
#   - combinations ("Hecht, Solberg, Robinson, Goldberg & Bagula")
# "et al." accepted as a special token in the word sequence so "Smith, et
# al. v. Jones" matches. Up to 10 word tokens per side so long law-firm
# names match.
_WORD = r"[A-Z][A-Za-z0-9'.\-]*"
_ET_AL = r"et\s+al\."
_WORD_OR_ETAL = rf"(?:{_WORD}|{_ET_AL})"
_WORD_SEP = r"(?:\s*[,&]\s*|\s+)"
_CASE_NAME_FRAGMENT = (
    r"(?:[\*_]{1,2})?"                            # optional italic/bold open
    r"(" + _WORD +                                 # first word
    r"(?:\s+(?:de|del|la|of|the|von|van))?"       # optional connector
    r"(?:" + _WORD_SEP + _WORD_OR_ETAL + r"){0,10}"   # up to 10 more words
    r"\s+vs?\.\s+"                                # required " v. " or " vs. "
    + _WORD +
    r"(?:" + _WORD_SEP + _WORD_OR_ETAL + r"){0,10}"   # up to 10 more words
    r")"
    r"(?:[\*_]{1,2})?"                            # optional italic/bold close
)

# Single-party citations: "In re ...", "Estate of ...", "Ex parte ...",
# "Matter of ...". These have no "v." separator.
_SINGLE_PARTY_PREFIX = r"(?:In\s+re|Estate\s+of|Ex\s+parte|Matter\s+of)"
_SINGLE_PARTY_NAME_FRAGMENT = (
    r"(?:[\*_]{1,2})?"
    r"(" + _SINGLE_PARTY_PREFIX + r"\s+(?:Marriage\s+of\s+)?"
    + _WORD + r"(?:" + _WORD_SEP + _WORD_OR_ETAL + r"){0,10}"
    r")"
    r"(?:[\*_]{1,2})?"
)

_YEAR = r"(\d{4})"
_VOL = r"(\d+)"
_PAGE = r"(\d+)"

# Optional pincite: ", 415" / ", 415-17" optionally followed by ", fn. 3"
# or ", n. 3". The negative lookahead `(?!\s+(?:Cal|P)\.)` prevents the
# pincite from greedily consuming what is actually the first volume of
# a parallel citation (e.g. "100, 105 Cal.Rptr.3d 50" - the ", 105"
# belongs to the parallel reporter, not the pincite).
# The `(?!\d)` after the digit run prevents the engine from backtracking
# the greedy `\d+` to a shorter prefix in order to satisfy the negative-
# reporter lookahead. Without it, "100, 105 Cal.Rptr.3d 50" would match
# pincite=", 10" (chopping "105" in half) just to keep "5 Cal." outside
# the lookahead window.
_PINCITE = (
    r"(?:\s*,\s*\d+(?:-\d+)?(?!\d)(?!\s+(?:Cal|P)\.)"
    r"(?:\s*,\s*(?:fn|n)\.\s*\d+)?)?"
)

# Parallel-cite chunk: ", 105 Cal.Rptr.3d 50" optionally with its own
# pincite. Used as a post-match extension for raw_text. Same guard as
# _PINCITE: prevent the inner pincite from eating what is actually the
# volume number of the NEXT parallel reporter.
_PARALLEL_CITE_RE = re.compile(
    rf"\s*,\s*\d+\s+(?:{_REPORTER_PATTERN})\s+\d+"
    rf"(?:\s*,\s*\d+(?:-\d+)?(?!\d)(?!\s+(?:Cal|P)\.))?"
)

_CASE_CITE_RE = re.compile(
    rf"{_CASE_NAME_FRAGMENT}\s*\({_YEAR}\)\s+{_VOL}\s+({_REPORTER_PATTERN})\s+{_PAGE}{_PINCITE}",
)

_SINGLE_PARTY_CITE_RE = re.compile(
    rf"{_SINGLE_PARTY_NAME_FRAGMENT}\s*\({_YEAR}\)\s+{_VOL}\s+({_REPORTER_PATTERN})\s+{_PAGE}{_PINCITE}",
)


def _strip_italic_markers(s: str) -> str:
    return s.strip().strip("*_").strip()


def _normalize_case(case_name: str, vol: str, reporter: str, page: str) -> str:
    name = _strip_italic_markers(case_name)
    return f"{name} {vol} {reporter} {page}".strip()


# ---------------------------------------------------------------------------
# Statute cites
# ---------------------------------------------------------------------------

# Map normalized citation prefixes to leginfo lawCode values.
_CODE_ALIASES: dict[str, str] = {
    "code of civil procedure": "CCP",
    "code civ. proc.": "CCP",
    "code civ proc": "CCP",
    "ccp": "CCP",
    "evidence code": "EVID",
    "evid. code": "EVID",
    "evid code": "EVID",
    "civil code": "CIV",
    "civ. code": "CIV",
    "civ code": "CIV",
    "penal code": "PEN",
    "pen. code": "PEN",
    "government code": "GOV",
    "gov. code": "GOV",
    "business and professions code": "BPC",
    "bus. & prof. code": "BPC",
    "b&p code": "BPC",
    "health and safety code": "HSC",
    "health & saf. code": "HSC",
    "labor code": "LAB",
    "lab. code": "LAB",
    "vehicle code": "VEH",
    "veh. code": "VEH",
    "family code": "FAM",
    "fam. code": "FAM",
    "probate code": "PROB",
    "prob. code": "PROB",
}

# Build a single alternation for the code-name prefix, longest match first.
_CODE_PREFIX_ALT = "|".join(
    sorted((re.escape(k) for k in _CODE_ALIASES.keys()), key=len, reverse=True)
)

# Section token. `§+` covers `§` and `§§` (and any rarer repeats);
# `sections?` covers both "section" and "sections" (plural form used
# when listing multiple sections). Token captured so caller can tell
# whether the citation is multi-section.
_SECTION_TOKEN = r"(§+|sections?|sec\.?|s\.)"

# Optional subdivision suffix that should be preserved in raw_text but
# not used as the cache key: ", subd. (a)" / ", subdiv. (a)(1)".
# Note: `subd(?:iv)?` matches both "subd" and "subdiv". Earlier mistake
# was `subdiv?` which means "subdi" with an optional trailing "v" -- that
# pattern requires 5+ characters and never matches the abbreviated form.
_STATUTE_SUBDIV = r"(?:\s*,?\s*subd(?:iv)?\.\s*\([a-zA-Z0-9]+\)(?:\([a-zA-Z0-9]+\))?)?"

_STATUTE_CITE_RE = re.compile(
    rf"({_CODE_PREFIX_ALT})\s*,?\s*{_SECTION_TOKEN}\s*(\d+[\w.]*){_STATUTE_SUBDIV}",
    re.IGNORECASE,
)

# Reversed form: "section 2024.020 of the Code of Civil Procedure" or
# "sections 2024.020, 2024.050 and 2031.030 of the Code of Civil Procedure".
# The middle group captures the entire comma/and-separated section list;
# extract_citations parses individual sections out of it.
_STATUTE_CITE_REVERSED_RE = re.compile(
    rf"{_SECTION_TOKEN}\s*(\d+[\w.]*(?:\s*(?:,|and)\s*\d+[\w.]*)*)\s+of\s+(?:the\s+)?({_CODE_PREFIX_ALT})",
    re.IGNORECASE,
)

# Extract individual section numbers from a "2024.020, 2024.050 and 2031.030"
# string captured by the reversed-form regex.
_REVERSED_SECTION_SPLIT_RE = re.compile(r"\d+[\w.]*")

# After a multi-section match like "Code Civ. Proc., §§ 2024.020, 2024.050",
# this regex captures each additional ", <section>" suffix so we can emit
# one Citation per section sharing the same law_code.
_ADDITIONAL_SECTION_RE = re.compile(r"\s*,\s*(\d+[\w.]*)")


def _normalize_statute(code_prefix: str, section_num: str) -> tuple[str, str]:
    law_code = _CODE_ALIASES.get(code_prefix.strip().lower(), "")
    return law_code, section_num.strip().rstrip(".")


def _is_multi_section_token(token: str) -> bool:
    """True if the section token signals a list of sections (§§, sections)."""
    t = token.strip().lower()
    return t.startswith("§§") or t == "sections"


# ---------------------------------------------------------------------------
# Rules of Court
# ---------------------------------------------------------------------------

_RULE_CITE_RE = re.compile(
    r"(?:(?:California|Cal\.)\s+Rules?\s+of\s+Court|CRC),?\s+rule\s+(\d+\.\d+(?:\.\d+)?)",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------------
# Proposition extraction (sentence window)
# ---------------------------------------------------------------------------

# Sentence boundary: period/question/exclamation followed by whitespace or end.
# Tries to avoid splitting on "Inc." / "v." / "§" abbreviations by requiring a
# trailing space and an uppercase or end-of-string after.  Negative lookbehinds
# guard against splitting after legal abbreviations like " v.", " Inc.", " App.".
_SENTENCE_END_RE = re.compile(
    r"(?<![\s][A-Za-z]\.)(?<=[.!?])(?=\s+[A-Z*_(])"
    r"|(?<![\s][A-Za-z]\.)(?<=[.!?])\s*$"
)


def _sentence_spans(text: str) -> list[tuple[int, int]]:
    """Return [(start, end), ...] of sentence-like spans in text."""
    if not text:
        return []
    spans: list[tuple[int, int]] = []
    start = 0
    for m in _SENTENCE_END_RE.finditer(text):
        end = m.start()
        if end > start:
            spans.append((start, end + 1))
        start = m.end()
    if start < len(text):
        spans.append((start, len(text)))
    return spans


def _proposition_for_offset(body_text: str, offset: int) -> str:
    """Return the containing sentence + up to 2 sentences of prior context."""
    spans = _sentence_spans(body_text)
    if not spans:
        return ""
    containing_idx = None
    for i, (start, end) in enumerate(spans):
        if start <= offset < end:
            containing_idx = i
            break
    if containing_idx is None:
        return ""
    prior_idx = max(0, containing_idx - 2)
    start = spans[prior_idx][0]
    end = spans[containing_idx][1]
    return body_text[start:end].strip()


def _normalize_body_text(body_text: str) -> str:
    """Defensive cleanup of LLM body text before parsing citations.

    - Convert HTML entities the LLM occasionally emits (``&amp;`` -> ``&``)
      so the case-name regex sees the actual ampersand.
    - Collapse non-breaking spaces (U+00A0) to regular spaces.
    - Normalize Unicode dashes (en/em dash) and smart quotes to ASCII so
      sentence/word regexes match consistently.
    """
    if not body_text:
        return ""
    text = body_text.replace("&amp;", "&")
    text = text.replace(" ", " ")          # non-breaking space
    text = text.replace("–", "-")          # en dash
    text = text.replace("—", "-")          # em dash
    text = text.replace("‘", "'")          # left single quote
    text = text.replace("’", "'")          # right single quote / apostrophe
    text = text.replace("“", '"')          # left double quote
    text = text.replace("”", '"')          # right double quote
    return text


def _strip_markers(s: str) -> str:
    """Strip italic/bold markdown markers from a citation's raw text."""
    return s.replace("*", "").replace("_", "")


def _emit_case_cite(
    citations: list[Citation],
    *,
    body_text: str,
    match: re.Match,
    case_name_raw: str,
    year: str,
    vol: str,
    reporter: str,
    page: str,
) -> None:
    """Append a case Citation, extending raw_text past any parallel cites."""
    end = match.end()
    while True:
        pm = _PARALLEL_CITE_RE.match(body_text, end)
        if not pm:
            break
        end = pm.end()
    raw_text = _strip_markers(body_text[match.start():end])
    case_name = _strip_italic_markers(case_name_raw)
    citations.append(
        Citation(
            kind="case",
            raw_text=raw_text,
            normalized=_normalize_case(case_name, vol, reporter, page),
            proposition=_proposition_for_offset(body_text, match.start()),
            body_offset=match.start(),
            case_name=case_name,
            year=year,
            reporter_citation=f"{vol} {reporter} {page}",
        )
    )


def extract_citations(body_text: str) -> list[Citation]:
    """Extract case + statute + rule citations from a draft body."""
    citations: list[Citation] = []
    if not body_text:
        return citations
    body_text = _normalize_body_text(body_text)

    # Case cites (with "v." or "vs.")
    for m in _CASE_CITE_RE.finditer(body_text):
        case_name_raw, year, vol, reporter, page = m.group(1, 2, 3, 4, 5)
        _emit_case_cite(
            citations,
            body_text=body_text,
            match=m,
            case_name_raw=case_name_raw,
            year=year,
            vol=vol,
            reporter=reporter,
            page=page,
        )

    # Single-party cites (In re X / Estate of X / Ex parte X / Matter of X)
    for m in _SINGLE_PARTY_CITE_RE.finditer(body_text):
        name_raw, year, vol, reporter, page = m.group(1, 2, 3, 4, 5)
        _emit_case_cite(
            citations,
            body_text=body_text,
            match=m,
            case_name_raw=name_raw,
            year=year,
            vol=vol,
            reporter=reporter,
            page=page,
        )

    # Statute cites (forward form: "Code Civ. Proc., section N")
    for m in _STATUTE_CITE_RE.finditer(body_text):
        prefix, token, section = m.group(1, 2, 3)
        law_code, section_num = _normalize_statute(prefix, section)
        if not law_code:
            continue
        citations.append(
            Citation(
                kind="statute",
                raw_text=m.group(0),
                normalized=f"{law_code} {section_num}",
                proposition=_proposition_for_offset(body_text, m.start()),
                body_offset=m.start(),
                law_code=law_code,
                section_num=section_num,
            )
        )
        # Multi-section list ("section sign sign 2024.020, 2024.050"):
        # walk forward past comma-separated section numbers and emit one
        # Citation per additional section, sharing the same law_code.
        if _is_multi_section_token(token):
            cursor = m.end()
            while True:
                em = _ADDITIONAL_SECTION_RE.match(body_text, cursor)
                if not em:
                    break
                extra = em.group(1).rstrip(".")
                citations.append(
                    Citation(
                        kind="statute",
                        raw_text=em.group(0).lstrip(", "),
                        normalized=f"{law_code} {extra}",
                        proposition=_proposition_for_offset(body_text, em.start(1)),
                        body_offset=em.start(1),
                        law_code=law_code,
                        section_num=extra,
                    )
                )
                cursor = em.end()

    # Statute cites (reversed form: "section N of the X Code" or
    # "sections N, M and O of the X Code")
    for m in _STATUTE_CITE_REVERSED_RE.finditer(body_text):
        sections_blob, prefix = m.group(2), m.group(3)
        law_code = _CODE_ALIASES.get(prefix.strip().lower(), "")
        if not law_code:
            continue
        # Split the captured section-list blob into individual sections.
        blob_offset = m.start(2)
        for sm in _REVERSED_SECTION_SPLIT_RE.finditer(sections_blob):
            section_num = sm.group(0).rstrip(".")
            citations.append(
                Citation(
                    kind="statute",
                    raw_text=m.group(0),
                    normalized=f"{law_code} {section_num}",
                    proposition=_proposition_for_offset(body_text, m.start()),
                    body_offset=blob_offset + sm.start(),
                    law_code=law_code,
                    section_num=section_num,
                )
            )

    # Rules of Court
    for m in _RULE_CITE_RE.finditer(body_text):
        rule_num = m.group(1)
        citations.append(
            Citation(
                kind="rule",
                raw_text=m.group(0),
                normalized=f"CRC rule {rule_num}",
                proposition=_proposition_for_offset(body_text, m.start()),
                body_offset=m.start(),
            )
        )

    # Deduplicate by (normalized, body_offset) - guards against the rare
    # case where two regexes match overlapping spans.
    seen: set[tuple[str, int]] = set()
    deduped: list[Citation] = []
    for c in citations:
        key = (c.normalized, c.body_offset)
        if key in seen:
            continue
        seen.add(key)
        deduped.append(c)

    deduped.sort(key=lambda c: c.body_offset)
    return deduped
