# Mediation Brief Generator Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Mediation Brief" template to the Chat tab that generates a comprehensive defense-side mediation brief from selected documents, outputs it as a formatted Word document, and supports conversational refinement.

**Architecture:** A standalone `icharlotte_core/mediation_brief.py` module owns the entire pipeline — planning pass, sequential section generation (Introduction last), Word assembly using the case's caption template, and LLM-driven conversational refinement. The ChatTab gets minimal changes: a new menu item, a confirmation dialog, and hooks to launch and display generation. Per-section prompts are stored in `Scripts/prompts/mediation_brief/` and editable via the Prompt Engineering Workbench.

**Tech Stack:** Python 3.x, PyQt6 (QThread for background generation), python-docx + lxml for Word assembly, PyMuPDF (fitz) for sample PDF extraction, LLMCaller for multi-provider LLM calls with fallback.

**Spec:** `docs/superpowers/specs/2026-04-10-mediation-brief-design.md`

---

## File Structure

| File | Responsibility |
|------|----------------|
| **Create:** `icharlotte_core/mediation_brief.py` | `MediationBriefGenerator` (pipeline orchestration, style caching, section generation, Word assembly, refinement state), `MediationBriefWorker(QThread)` (background execution with signals) |
| **Create:** `Scripts/prompts/mediation_brief/planning_current.txt` | Planning pass prompt — extract facts, arguments, quotes, dates from documents |
| **Create:** `Scripts/prompts/mediation_brief/statement_of_facts_current.txt` | Statement of Facts section prompt |
| **Create:** `Scripts/prompts/mediation_brief/procedural_status_current.txt` | Procedural Status section prompt |
| **Create:** `Scripts/prompts/mediation_brief/liability_current.txt` | Liability section prompt |
| **Create:** `Scripts/prompts/mediation_brief/damages_current.txt` | Damages section prompt |
| **Create:** `Scripts/prompts/mediation_brief/settlement_position_current.txt` | Settlement Position section prompt |
| **Create:** `Scripts/prompts/mediation_brief/conclusion_current.txt` | Conclusion section prompt |
| **Create:** `Scripts/prompts/mediation_brief/introduction_current.txt` | Introduction section prompt (sees all other sections) |
| **Create:** `Scripts/prompts/mediation_brief/routing_current.txt` | Refinement routing prompt — identify which sections to regenerate |
| **Create:** `tests/test_mediation_brief.py` | Unit tests for generator, caption handling, text parsing, Word formatting |
| **Modify:** `icharlotte_core/ui/tabs.py` | Add "Mediation Brief" to `update_template_menu()`, confirmation dialog, generation hooks, refinement mode routing in `send_message()` |
| **Modify:** `icharlotte_core/ui/dialogs.py` | Add `"mediation_brief": "agent_mediation_brief"` to `WORKBENCH_TO_AGENT_ID` |
| **Modify:** `icharlotte_core/chat/models.py` | Add `builtin_mediation_brief` to `BUILTIN_PROMPTS` list |
| **Modify:** `Scripts/prompts/registry.json` | Register all 9 mediation_brief prompt passes |
| **Modify:** `config/llm_preferences.json` | Register `agent_mediation_brief` agent config |

---

### Task 1: Create Prompt Files and Register in Workbench

**Files:**
- Create: `Scripts/prompts/mediation_brief/planning_current.txt`
- Create: `Scripts/prompts/mediation_brief/statement_of_facts_current.txt`
- Create: `Scripts/prompts/mediation_brief/procedural_status_current.txt`
- Create: `Scripts/prompts/mediation_brief/liability_current.txt`
- Create: `Scripts/prompts/mediation_brief/damages_current.txt`
- Create: `Scripts/prompts/mediation_brief/settlement_position_current.txt`
- Create: `Scripts/prompts/mediation_brief/conclusion_current.txt`
- Create: `Scripts/prompts/mediation_brief/introduction_current.txt`
- Create: `Scripts/prompts/mediation_brief/routing_current.txt`
- Modify: `Scripts/prompts/registry.json`
- Modify: `icharlotte_core/ui/dialogs.py` (~line 398)
- Modify: `icharlotte_core/chat/models.py` (~line 185)
- Modify: `config/llm_preferences.json`

- [ ] **Step 1: Create the prompt directory**

```bash
mkdir -p Scripts/prompts/mediation_brief
```

- [ ] **Step 2: Create planning pass prompt**

Write `Scripts/prompts/mediation_brief/planning_current.txt`:

```text
You are a senior defense litigation attorney preparing for mediation. Analyze the following case documents and extract all key information needed to draft a comprehensive mediation brief.

Extract the following in a structured format:

CASE INFORMATION:
- Plaintiff name(s)
- Defendant name(s)
- Case number
- Court and jurisdiction
- Judge
- Incident date
- Trial date
- Deposition dates (with deponent names)

KEY FACTS:
List all material facts chronologically. Include dates, locations, witnesses, and specific details.

LIABILITY ARGUMENTS:
For each distinct argument that challenges plaintiff's ability to establish liability:
- Argument title
- Supporting facts
- Applicable law (if apparent from documents)
- Supporting deposition testimony (with page:line citations)

DAMAGES ARGUMENTS:
For each distinct argument that challenges plaintiff's claimed damages:
- Argument title
- Supporting facts
- Medical evidence or lack thereof
- Inconsistencies in plaintiff's claims
- Supporting deposition testimony (with page:line citations)

COMPARATIVE FAULT:
If applicable, list all arguments for plaintiff's comparative fault with supporting facts.

DEPOSITION QUOTES:
For each significant deposition quote:
- Deponent last name
- Quote text (verbatim, cleaned of transcript artifacts like line numbers and dashes)
- Page number and line number
- Which argument it supports

SETTLEMENT HISTORY:
- Policy limits
- Prior demands (with dates and amounts)
- Prior offers (with dates and amounts)

OTHER FAVORABLE FACTS:
Any additional facts favorable to the defense not captured above.
```

- [ ] **Step 3: Create Statement of Facts prompt**

Write `Scripts/prompts/mediation_brief/statement_of_facts_current.txt`:

```text
You are a senior defense litigation attorney drafting the STATEMENT OF FACTS section of a confidential mediation brief. Write a detailed, comprehensive, and persuasive factual narrative.

Requirements:
- Present facts chronologically
- Include all material facts necessary for the reader to understand the case and the arguments in the rest of the brief
- Emphasize facts favorable to the defense while remaining factually accurate
- Where deposition testimony is available, include verbatim quotes (cleaned of transcript artifacts) with citations in this format: (LastName Depo Trns., at p. PageNum:LineNum.)
- Write in a professional, authoritative tone
- Be detailed and thorough — length is not a concern
- Do not include section headings — the heading will be added separately
- Do not use placeholder text or brackets — write around any missing information naturally
```

- [ ] **Step 4: Create Procedural Status prompt**

Write `Scripts/prompts/mediation_brief/procedural_status_current.txt`:

```text
You are a senior defense litigation attorney drafting the PROCEDURAL STATUS section of a confidential mediation brief.

Requirements:
- Identify the trial date
- Identify the dates that party depositions were taken and the names of deponents
- Note any other significant procedural events (motions filed, discovery status, etc.) if apparent from the documents
- Write in a professional, concise tone
- Do not include section headings — the heading will be added separately
- Do not use placeholder text or brackets — write around any missing information naturally
```

- [ ] **Step 5: Create Liability prompt**

Write `Scripts/prompts/mediation_brief/liability_current.txt`:

```text
You are a senior defense litigation attorney drafting the LIABILITY section of a confidential mediation brief. This section must be detailed, comprehensive, persuasive, and written to convince the mediator and opposing counsel that plaintiff cannot establish liability at trial.

Requirements:
- Start with a 1-2 sentence introduction paragraph
- For each key liability argument, write a separate subsection with:
  - A descriptive title in Title Case (the subsection letter and formatting will be added separately — just write the title text)
  - A brief statement of the applicable law
  - A detailed, persuasive application of the law to the case facts
  - Explanation of why plaintiff will not be able to establish this element of liability at trial
- Include verbatim deposition quotes where available, cleaned of transcript artifacts, with citations: (LastName Depo Trns., at p. PageNum:LineNum.)
- Be aggressive and persuasive while remaining professional
- Length is not a concern — be thorough
- Do not use placeholder text or brackets — write around any missing information naturally

Format each subsection as:
SUBSECTION: [Title Text]
[Content paragraphs]
```

- [ ] **Step 6: Create Damages prompt**

Write `Scripts/prompts/mediation_brief/damages_current.txt`:

```text
You are a senior defense litigation attorney drafting the DAMAGES section of a confidential mediation brief. This section must be detailed, comprehensive, persuasive, and written to convince the mediator and opposing counsel that plaintiff cannot recover the claimed damages at trial.

Requirements:
- Start with a 1-2 sentence introduction paragraph
- For each key damages argument, write a separate subsection with:
  - A descriptive title in Title Case (the subsection letter and formatting will be added separately — just write the title text)
  - A detailed, persuasive argument for why plaintiff will not be able to recover the claimed damages
  - Reference specific case facts, medical records, inconsistencies, and gaps in plaintiff's evidence
- Include verbatim deposition quotes where available, cleaned of transcript artifacts, with citations: (LastName Depo Trns., at p. PageNum:LineNum.)
- Address specific categories of damages (medical specials, general damages, future damages, lost wages, etc.) as applicable
- Be aggressive and persuasive while remaining professional
- Length is not a concern — be thorough
- Do not use placeholder text or brackets — write around any missing information naturally

Format each subsection as:
SUBSECTION: [Title Text]
[Content paragraphs]
```

- [ ] **Step 7: Create Settlement Position prompt**

Write `Scripts/prompts/mediation_brief/settlement_position_current.txt`:

```text
You are a senior defense litigation attorney drafting the SETTLEMENT POSITION section of a confidential mediation brief.

Requirements:
- Summarize the applicable policy limits
- Summarize all prior settlement demands and offers, with dates and amounts where available
- Note the current posture of settlement negotiations
- Write in a professional, measured tone
- Do not include section headings — the heading will be added separately
- Do not use placeholder text or brackets — write around any missing information naturally
```

- [ ] **Step 8: Create Conclusion prompt**

Write `Scripts/prompts/mediation_brief/conclusion_current.txt`:

```text
You are a senior defense litigation attorney drafting the CONCLUSION section of a confidential mediation brief.

Requirements:
- Write 1-2 paragraphs summarizing the defense position
- Explain why the defense will prevail at trial
- Summarize the key weaknesses in plaintiff's case (both liability and damages)
- End with a tone that is firm but expresses willingness to engage in good-faith mediation
- Do not include section headings — the heading will be added separately
- Do not use placeholder text or brackets — write around any missing information naturally
```

- [ ] **Step 9: Create Introduction prompt**

Write `Scripts/prompts/mediation_brief/introduction_current.txt`:

```text
You are a senior defense litigation attorney drafting the INTRODUCTION section of a confidential mediation brief. This section is generated AFTER all other sections and must accurately summarize the arguments made throughout the brief.

Requirements:
- Paragraph 1: A one-paragraph summary of the basic facts of the case
- Paragraph 2: State whether we dispute liability or concede liability for purposes of mediation. Summarize ALL key arguments challenging liability (or for comparative fault if conceding liability). This must accurately reflect the arguments in the Liability section.
- Paragraph 3: State that we challenge the nature and scope of Plaintiff's claimed injuries and damages. Summarize ALL key arguments challenging damages. This must accurately reflect the arguments in the Damages section.
- Paragraph 4 (if needed): Summarize any other favorable defense arguments not covered in paragraphs 2 or 3
- Final sentence: One sentence indicating we come to this mediation in good faith but plaintiff must recognize the significant problems in their case
- The introduction must be persuasive and comprehensive, serving as a preview of all arguments in the brief
- Do not include section headings — the heading will be added separately
- Do not use placeholder text or brackets — write around any missing information naturally
```

- [ ] **Step 10: Create Routing prompt**

Write `Scripts/prompts/mediation_brief/routing_current.txt`:

```text
You are a routing assistant for a mediation brief refinement system. The user has generated a mediation brief with these sections:

1. introduction
2. statement_of_facts
3. procedural_status
4. liability
5. damages
6. settlement_position
7. conclusion

Based on the user's message, determine which section(s) need to be regenerated. Respond with ONLY a comma-separated list of section names from the list above, or "none" if the message is not about refining the brief.

Examples:
- "Make the damages section stronger" → damages
- "Add a comparative fault argument" → liability
- "Change the trial date to March 2027" → procedural_status
- "What's the weather like?" → none
- "Rewrite the intro and conclusion" → introduction,conclusion
- "Add more detail about the medical records" → damages,statement_of_facts
```

- [ ] **Step 11: Register prompts in registry.json**

Read `Scripts/prompts/registry.json`, then add 9 new entries under the `"prompts"` key. Each entry follows this pattern (repeat for all 9 passes):

```json
"mediation_brief:planning": {
    "agent": "mediation_brief",
    "pass_name": "planning",
    "current_version": "v1",
    "versions": [
        {
            "version": "v1",
            "created": "<current ISO timestamp>",
            "description": "Initial mediation brief planning pass",
            "author": "",
            "is_current": true,
            "performance_score": 0.0,
            "usage_count": 0
        }
    ]
}
```

The 9 pass names are: `planning`, `statement_of_facts`, `procedural_status`, `liability`, `damages`, `settlement_position`, `conclusion`, `introduction`, `routing`.

- [ ] **Step 12: Add Workbench mapping**

In `icharlotte_core/ui/dialogs.py`, find the `WORKBENCH_TO_AGENT_ID` dict (~line 388) and add:

```python
"mediation_brief": "agent_mediation_brief",
```

after the existing `"chat": "func_chat"` entry.

- [ ] **Step 13: Add builtin prompt entry**

In `icharlotte_core/chat/models.py`, add to the `BUILTIN_PROMPTS` list (~line 185):

```python
QuickPrompt(
    id='builtin_mediation_brief',
    name='Mediation Brief',
    prompt='',  # Not used — triggers special generation flow
    category='Generation',
    is_builtin=True
),
```

- [ ] **Step 14: Register agent in LLM config**

In `config/llm_preferences.json`, add an `agent_mediation_brief` entry following the existing agent pattern. Set high max_tokens (-1 for unlimited) and a 300-second timeout. Use the project's default model sequence.

- [ ] **Step 15: Commit**

```bash
git add Scripts/prompts/mediation_brief/ icharlotte_core/ui/dialogs.py icharlotte_core/chat/models.py Scripts/prompts/registry.json config/llm_preferences.json
git commit -m "feat(mediation-brief): create prompt files and register in workbench"
```

---

### Task 2: Sample Brief Style Extraction and Caching

**Files:**
- Create: `icharlotte_core/mediation_brief.py` (first portion — style extraction only)
- Test: `tests/test_mediation_brief.py` (first tests)

- [ ] **Step 1: Write tests for style extraction**

Create `tests/test_mediation_brief.py`:

```python
"""Tests for mediation brief generator."""
import json
import os
import tempfile
import unittest
from unittest.mock import patch, MagicMock


class TestStyleExtraction(unittest.TestCase):
    """Test sample brief style extraction and caching."""

    def test_extract_sections_from_text(self):
        """Sections are extracted by roman numeral heading pattern."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        sample_text = (
            "I.     INTRODUCTION\n"
            "This is the introduction paragraph.\n\n"
            "II.     STATEMENT OF FACTS\n"
            "These are the facts of the case.\n"
            "More facts here.\n\n"
            "III.     LIABILITY\n"
            "Liability arguments here.\n"
        )
        gen = MediationBriefGenerator.__new__(MediationBriefGenerator)
        sections = gen._extract_sections_from_text(sample_text)

        self.assertIn("introduction", sections)
        self.assertIn("statement_of_facts", sections)
        self.assertIn("liability", sections)
        self.assertIn("introduction paragraph", sections["introduction"])
        self.assertIn("facts of the case", sections["statement_of_facts"])

    def test_extract_sections_handles_missing_sections(self):
        """Missing sections return empty strings, not errors."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        sample_text = (
            "I.     INTRODUCTION\n"
            "Intro text.\n\n"
            "VII.     CONCLUSION\n"
            "Conclusion text.\n"
        )
        gen = MediationBriefGenerator.__new__(MediationBriefGenerator)
        sections = gen._extract_sections_from_text(sample_text)

        self.assertIn("introduction", sections)
        self.assertIn("conclusion", sections)
        # Missing sections should not be in dict
        self.assertNotIn("liability", sections)

    def test_cache_structure(self):
        """Cache JSON has source_hashes and sections keys."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator.__new__(MediationBriefGenerator)
        gen._sample_dir = "C:\\AI\\Mediation Briefs"

        # Mock the PDF reading
        with patch.object(gen, '_read_sample_pdfs') as mock_read:
            mock_read.return_value = {
                "hashes": {"sample1.pdf": "abc123"},
                "sections": {
                    "introduction": ["Intro text from sample 1"],
                    "liability": ["Liability text from sample 1"],
                }
            }
            with tempfile.TemporaryDirectory() as tmpdir:
                cache_path = os.path.join(tmpdir, "style_cache.json")
                gen._cache_path = cache_path
                gen._save_style_cache(mock_read.return_value)

                with open(cache_path, 'r') as f:
                    cached = json.load(f)

                self.assertIn("source_hashes", cached)
                self.assertIn("sections", cached)
                self.assertEqual(cached["source_hashes"]["sample1.pdf"], "abc123")


if __name__ == '__main__':
    unittest.main()
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_mediation_brief.py -v
```

Expected: FAIL — `icharlotte_core.mediation_brief` does not exist yet.

- [ ] **Step 3: Implement style extraction in mediation_brief.py**

Create `icharlotte_core/mediation_brief.py` with the style extraction portion:

```python
"""Mediation Brief Generator.

Generates comprehensive defense-side mediation briefs from case documents.
Supports sequential section generation, Word document assembly from caption
templates, and conversational refinement.
"""

import hashlib
import json
import logging
import os
import re
from typing import Dict, List, Optional, Tuple

log = logging.getLogger(__name__)

# Sample briefs location
SAMPLE_BRIEFS_DIR = r"C:\AI\Mediation Briefs"

# Section names in document order
SECTION_ORDER = [
    "introduction",
    "statement_of_facts",
    "procedural_status",
    "liability",
    "damages",
    "settlement_position",
    "conclusion",
]

# Generation order (Introduction last so it can reference all other sections)
GENERATION_ORDER = [
    "statement_of_facts",
    "procedural_status",
    "liability",
    "damages",
    "settlement_position",
    "conclusion",
    "introduction",
]

# Map section names to roman numeral headings
SECTION_HEADINGS = {
    "introduction": ("I", "INTRODUCTION"),
    "statement_of_facts": ("II", "STATEMENT OF FACTS"),
    "procedural_status": ("III", "PROCEDURAL STATUS"),
    "liability": ("IV", "LIABILITY"),
    "damages": ("V", "DAMAGES"),
    "settlement_position": ("VI", "SETTLEMENT POSITION"),
    "conclusion": ("VII", "CONCLUSION"),
}

# Heading patterns for extracting sections from sample text
_HEADING_PATTERN = re.compile(
    r'^([IVX]+)\.\s+(INTRODUCTION|STATEMENT OF FACTS|FACTUAL BACKGROUND|'
    r'PROCEDURAL STATUS|PROCEDURAL HISTORY|LIABILITY|DAMAGES|'
    r'SETTLEMENT POSITION|SETTLEMENT|CONCLUSION)\s*$',
    re.MULTILINE | re.IGNORECASE
)

# Map variant heading names to canonical section names
_HEADING_TO_SECTION = {
    "INTRODUCTION": "introduction",
    "STATEMENT OF FACTS": "statement_of_facts",
    "FACTUAL BACKGROUND": "statement_of_facts",
    "PROCEDURAL STATUS": "procedural_status",
    "PROCEDURAL HISTORY": "procedural_status",
    "LIABILITY": "liability",
    "DAMAGES": "damages",
    "SETTLEMENT POSITION": "settlement_position",
    "SETTLEMENT": "settlement_position",
    "CONCLUSION": "conclusion",
}

SCRIPTS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "Scripts")
PROMPTS_DIR = os.path.join(SCRIPTS_DIR, "prompts", "mediation_brief")
CACHE_PATH = os.path.join(PROMPTS_DIR, "style_cache.json")


class MediationBriefGenerator:
    """Orchestrates mediation brief generation, assembly, and refinement."""

    def __init__(self):
        self._sample_dir = SAMPLE_BRIEFS_DIR
        self._cache_path = CACHE_PATH
        self._style_cache: Optional[Dict] = None

        # State for refinement mode
        self.sections: Dict[str, str] = {}
        self.planning_output: str = ""
        self.document_content: str = ""
        self.caption_template_path: Optional[str] = None
        self.is_active = False  # True after generation completes

    # ── Style extraction & caching ──────────────────────────────

    def _extract_sections_from_text(self, text: str) -> Dict[str, str]:
        """Extract sections from sample brief text by roman numeral headings.

        Returns dict mapping canonical section name to section body text.
        """
        matches = list(_HEADING_PATTERN.finditer(text))
        if not matches:
            return {}

        sections = {}
        for i, match in enumerate(matches):
            heading_name = match.group(2).upper().strip()
            section_key = _HEADING_TO_SECTION.get(heading_name)
            if not section_key:
                continue

            start = match.end()
            end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
            body = text[start:end].strip()
            sections[section_key] = body

        return sections

    def _hash_file(self, path: str) -> str:
        """Return MD5 hex digest of a file."""
        h = hashlib.md5()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()

    def _read_sample_pdfs(self) -> Dict:
        """Read sample PDFs and extract sections from each.

        Returns dict with 'hashes' and 'sections' keys.
        """
        import fitz  # PyMuPDF

        if not os.path.isdir(self._sample_dir):
            log.warning("Sample briefs directory not found: %s", self._sample_dir)
            return {"hashes": {}, "sections": {}}

        pdf_files = [f for f in os.listdir(self._sample_dir)
                     if f.lower().endswith(".pdf")]
        if not pdf_files:
            log.warning("No PDF files found in %s", self._sample_dir)
            return {"hashes": {}, "sections": {}}

        hashes = {}
        all_sections: Dict[str, List[str]] = {}

        for fname in pdf_files:
            fpath = os.path.join(self._sample_dir, fname)
            hashes[fname] = self._hash_file(fpath)

            try:
                doc = fitz.open(fpath)
                full_text = ""
                for page in doc:
                    full_text += page.get_text() + "\n"
                doc.close()
            except Exception as e:
                log.warning("Failed to read sample PDF %s: %s", fname, e)
                continue

            sections = self._extract_sections_from_text(full_text)
            for section_name, body in sections.items():
                if not body:
                    continue
                if section_name not in all_sections:
                    all_sections[section_name] = []
                all_sections[section_name].append(body)

        # Keep best 2 per section (longest = most complete)
        for key in all_sections:
            excerpts = sorted(all_sections[key], key=len, reverse=True)
            all_sections[key] = excerpts[:2]

        return {"hashes": hashes, "sections": all_sections}

    def _save_style_cache(self, data: Dict):
        """Save extracted style data to cache file."""
        os.makedirs(os.path.dirname(self._cache_path), exist_ok=True)
        with open(self._cache_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        log.info("Style cache saved to %s", self._cache_path)

    def _load_style_cache(self) -> Optional[Dict]:
        """Load style cache if valid (hashes match current files)."""
        if not os.path.exists(self._cache_path):
            return None

        try:
            with open(self._cache_path, "r", encoding="utf-8") as f:
                cached = json.load(f)
        except (json.JSONDecodeError, OSError):
            return None

        # Validate hashes
        if not os.path.isdir(self._sample_dir):
            return cached  # Can't validate, use what we have

        current_files = [f for f in os.listdir(self._sample_dir)
                         if f.lower().endswith(".pdf")]
        cached_hashes = cached.get("source_hashes", {})

        for fname in current_files:
            fpath = os.path.join(self._sample_dir, fname)
            current_hash = self._hash_file(fpath)
            if cached_hashes.get(fname) != current_hash:
                log.info("Cache invalidated: %s has changed", fname)
                return None

        return cached

    def get_style_excerpts(self) -> Dict[str, List[str]]:
        """Get cached style excerpts, extracting from samples if needed.

        Returns dict mapping section name to list of example excerpts.
        """
        if self._style_cache is not None:
            return self._style_cache.get("sections", {})

        # Try loading from disk
        cached = self._load_style_cache()
        if cached:
            self._style_cache = cached
            return cached.get("sections", {})

        # Extract from samples and cache
        log.info("Extracting style from sample briefs...")
        data = self._read_sample_pdfs()
        if data["sections"]:
            self._save_style_cache(data)
        self._style_cache = data
        return data.get("sections", {})
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_mediation_brief.py -v
```

Expected: All 3 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief.py tests/test_mediation_brief.py
git commit -m "feat(mediation-brief): style extraction and caching from sample briefs"
```

---

### Task 3: Caption Template Handling

**Files:**
- Modify: `icharlotte_core/mediation_brief.py`
- Modify: `tests/test_mediation_brief.py`

- [ ] **Step 1: Write tests for caption template processing**

Add to `tests/test_mediation_brief.py`:

```python
from docx import Document
from docx.shared import Pt
from copy import deepcopy


class TestCaptionHandling(unittest.TestCase):
    """Test caption template finding and processing."""

    def _make_caption_doc(self, tmpdir, include_sig_block=False):
        """Helper: create a minimal caption template .docx."""
        doc = Document()
        doc.add_paragraph("LAW FIRM NAME")
        doc.add_paragraph("CAPTION PAGE")
        doc.add_paragraph("Some caption content")
        if include_sig_block:
            doc.add_paragraph("")
            doc.add_paragraph("DATED: April 10, 2026")
            doc.add_paragraph("")
            doc.add_paragraph("By: ____________________")
            doc.add_paragraph("John Smith, Esq.")
            doc.add_paragraph("State Bar No. 123456")
        path = os.path.join(tmpdir, "case_caption.docx")
        doc.save(path)
        return path

    def test_find_caption_in_folder(self):
        """Finds .docx with 'caption' in name."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            self._make_caption_doc(tmpdir)
            # Also create a non-caption file
            doc2 = Document()
            doc2.add_paragraph("Not a caption")
            doc2.save(os.path.join(tmpdir, "other_doc.docx"))

            result = gen.find_caption_template(tmpdir)
            self.assertIsNotNone(result)
            self.assertIn("caption", os.path.basename(result).lower())

    def test_find_caption_returns_none_when_missing(self):
        """Returns None when no caption file exists."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            doc = Document()
            doc.add_paragraph("Not a caption")
            doc.save(os.path.join(tmpdir, "other.docx"))

            result = gen.find_caption_template(tmpdir)
            self.assertIsNone(result)

    def test_replace_caption_page_text(self):
        """CAPTION PAGE is replaced with brief title in body."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            caption_path = self._make_caption_doc(tmpdir)
            output_path = os.path.join(tmpdir, "output.docx")
            gen.prepare_caption_template(caption_path, output_path)

            doc = Document(output_path)
            all_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertNotIn("CAPTION PAGE", all_text)
            self.assertIn("DEFENDANT'S CONFIDENTIAL MEDIATION BRIEF", all_text)

    def test_signature_block_detection(self):
        """Signature block paragraphs are detected and extracted."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            caption_path = self._make_caption_doc(tmpdir, include_sig_block=True)
            output_path = os.path.join(tmpdir, "output.docx")
            sig_paras = gen.prepare_caption_template(caption_path, output_path)

            # Signature block should have been detected
            self.assertTrue(len(sig_paras) > 0)
            sig_text = " ".join(p.text for p in sig_paras)
            self.assertIn("DATED", sig_text)

    def test_no_signature_block(self):
        """No signature block returns empty list."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        gen = MediationBriefGenerator()
        with tempfile.TemporaryDirectory() as tmpdir:
            caption_path = self._make_caption_doc(tmpdir, include_sig_block=False)
            output_path = os.path.join(tmpdir, "output.docx")
            sig_paras = gen.prepare_caption_template(caption_path, output_path)

            self.assertEqual(len(sig_paras), 0)
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_mediation_brief.py::TestCaptionHandling -v
```

Expected: FAIL — `find_caption_template` and `prepare_caption_template` don't exist.

- [ ] **Step 3: Implement caption template handling**

Add these methods to `MediationBriefGenerator` in `icharlotte_core/mediation_brief.py`:

```python
    # ── Caption template handling ───────────────────────────────

    def find_caption_template(self, folder: str) -> Optional[str]:
        """Find a .docx file with 'caption' in the name in the given folder.

        Returns the path to the first match, or None if not found.
        """
        if not os.path.isdir(folder):
            return None

        for fname in os.listdir(folder):
            if fname.lower().endswith(".docx") and "caption" in fname.lower():
                return os.path.join(folder, fname)
        return None

    def prepare_caption_template(
        self, caption_path: str, output_path: str
    ) -> list:
        """Process caption template: replace CAPTION PAGE, extract signature block.

        Args:
            caption_path: Path to the original caption .docx
            output_path: Path to save the processed copy

        Returns:
            List of signature block Paragraph objects (may be empty).
            These paragraphs are removed from the document and should be
            appended after the brief content.
        """
        import shutil
        from docx import Document
        from lxml import etree

        # Work on a copy
        shutil.copy2(caption_path, output_path)
        doc = Document(output_path)

        # Replace CAPTION PAGE in body (XML iteration for nested tables)
        self._replace_caption_page_body(doc)

        # Replace CAPTION PAGE in footers
        self._replace_caption_page_footers(doc)

        # Detect and extract signature block
        sig_paragraphs = self._extract_signature_block(doc)

        doc.save(output_path)
        return sig_paragraphs

    def _replace_caption_page_body(self, doc):
        """Replace 'CAPTION PAGE' text in document body using XML iteration.

        Searches all w:t elements including those in nested tables.
        Replaces with "DEFENDANT'S CONFIDENTIAL MEDIATION BRIEF" —
        bold, all caps, 'CONFIDENTIAL' underlined.
        """
        from lxml import etree
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        W_T = f"{{{W_NS}}}t"
        W_P = f"{{{W_NS}}}p"
        W_R = f"{{{W_NS}}}r"

        for t_elem in doc.element.body.iter(W_T):
            if t_elem.text and "CAPTION PAGE" in t_elem.text.upper():
                # Find parent paragraph
                para_elem = t_elem.getparent()
                while para_elem is not None and para_elem.tag != W_P:
                    para_elem = para_elem.getparent()
                if para_elem is None:
                    continue

                # Clear all runs
                for run_elem in list(para_elem.iter(W_R)):
                    para_elem.remove(run_elem)

                # Get existing paragraph's font info for sizing
                # Build three runs: "DEFENDANT'S " (bold) + "CONFIDENTIAL" (bold+underline) + " MEDIATION BRIEF" (bold)
                parts = [
                    ("DEFENDANT'S ", False),
                    ("CONFIDENTIAL", True),
                    (" MEDIATION BRIEF", False),
                ]
                for text, underline in parts:
                    run = OxmlElement("w:r")
                    rPr = OxmlElement("w:rPr")
                    rPr.append(OxmlElement("w:b"))
                    if underline:
                        u = OxmlElement("w:u")
                        u.set(qn("w:val"), "single")
                        rPr.append(u)
                    run.append(rPr)
                    t = OxmlElement("w:t")
                    t.text = text
                    t.set(qn("xml:space"), "preserve")
                    run.append(t)
                    para_elem.append(run)
                return  # Only replace first occurrence in body

    def _replace_caption_page_footers(self, doc):
        """Replace 'CAPTION PAGE' in all document footers."""
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        for section in doc.sections:
            for footer in [section.footer, section.even_page_footer,
                           section.first_page_footer]:
                if footer is None or not footer.is_linked_to_previous:
                    pass  # Process it
                if footer is None:
                    continue
                for para in footer.paragraphs:
                    if "CAPTION PAGE" in para.text.upper():
                        # Clear and rebuild
                        for run in para.runs:
                            run.text = ""
                        # Remove existing runs at XML level
                        W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                        W_R = f"{{{W_NS}}}r"
                        for run_elem in list(para._element.iter(W_R)):
                            para._element.remove(run_elem)

                        parts = [
                            ("DEFENDANT'S ", False),
                            ("CONFIDENTIAL", True),
                            (" MEDIATION BRIEF", False),
                        ]
                        for text, underline in parts:
                            run = OxmlElement("w:r")
                            rPr = OxmlElement("w:rPr")
                            rPr.append(OxmlElement("w:b"))
                            if underline:
                                u = OxmlElement("w:u")
                                u.set(qn("w:val"), "single")
                                rPr.append(u)
                            run.append(rPr)
                            t = OxmlElement("w:t")
                            t.text = text
                            t.set(qn("xml:space"), "preserve")
                            run.append(t)
                            para._element.append(run)

    def _extract_signature_block(self, doc) -> list:
        """Detect and extract signature block paragraphs from end of document.

        Scans backwards from the last paragraph looking for signature indicators
        (By:, DATED:, State Bar No., etc.). Extracts all paragraphs from the
        first indicator found to the end.

        Returns list of extracted Paragraph objects. These are removed from the doc.
        """
        sig_indicators = re.compile(
            r'^\s*(By:|DATED:|Respectfully submitted|State Bar No\.|'
            r'Attorney[s]? for|Counsel for)',
            re.IGNORECASE
        )

        paragraphs = list(doc.paragraphs)
        if not paragraphs:
            return []

        # Scan backwards to find the start of signature block
        sig_start = None
        # Look in last 15 paragraphs max
        search_range = min(15, len(paragraphs))
        for i in range(len(paragraphs) - 1, len(paragraphs) - search_range - 1, -1):
            if i < 0:
                break
            text = paragraphs[i].text.strip()
            if sig_indicators.match(text):
                sig_start = i

        if sig_start is None:
            return []

        # Collect signature paragraphs (store their XML elements)
        sig_elements = []
        for i in range(sig_start, len(paragraphs)):
            sig_elements.append(paragraphs[i])

        # Remove from document body (reverse order to preserve indices)
        for para in reversed(sig_elements):
            para._element.getparent().remove(para._element)

        return sig_elements
```

Also add these imports at the top of the file:

```python
import shutil
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_mediation_brief.py::TestCaptionHandling -v
```

Expected: All 5 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief.py tests/test_mediation_brief.py
git commit -m "feat(mediation-brief): caption template finding, processing, signature extraction"
```

---

### Task 4: Section Generation Pipeline

**Files:**
- Modify: `icharlotte_core/mediation_brief.py`
- Modify: `tests/test_mediation_brief.py`

- [ ] **Step 1: Write tests for section generation**

Add to `tests/test_mediation_brief.py`:

```python
class TestSectionGeneration(unittest.TestCase):
    """Test LLM-based section generation pipeline."""

    def test_build_section_prompt_includes_planning_output(self):
        """Section prompt includes the planning pass extraction."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        gen.planning_output = "KEY FACTS:\n- Accident on Jan 1, 2025"
        gen.sections = {}
        gen.document_content = "Document text here"
        gen._style_cache = {"sections": {}}

        prompt = gen._build_section_prompt("statement_of_facts")
        self.assertIn("Accident on Jan 1, 2025", prompt)

    def test_build_section_prompt_includes_prior_sections(self):
        """Later sections see earlier section text."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        gen.planning_output = "Planning data"
        gen.sections = {
            "statement_of_facts": "The plaintiff was injured on Main St.",
            "procedural_status": "Trial is set for June 2027.",
        }
        gen.document_content = "Doc text"
        gen._style_cache = {"sections": {}}

        prompt = gen._build_section_prompt("liability")
        self.assertIn("injured on Main St", prompt)
        self.assertIn("Trial is set for June 2027", prompt)

    def test_build_introduction_prompt_includes_all_sections(self):
        """Introduction (generated last) sees all other sections."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        gen.planning_output = "Planning data"
        gen.sections = {
            "statement_of_facts": "Facts text",
            "procedural_status": "Status text",
            "liability": "Liability text",
            "damages": "Damages text",
            "settlement_position": "Settlement text",
            "conclusion": "Conclusion text",
        }
        gen.document_content = "Doc text"
        gen._style_cache = {"sections": {}}

        prompt = gen._build_section_prompt("introduction")
        self.assertIn("Liability text", prompt)
        self.assertIn("Damages text", prompt)
        self.assertIn("Conclusion text", prompt)

    def test_generation_order(self):
        """Introduction is generated last."""
        from icharlotte_core.mediation_brief import GENERATION_ORDER
        self.assertEqual(GENERATION_ORDER[-1], "introduction")
        self.assertEqual(len(GENERATION_ORDER), 7)
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_mediation_brief.py::TestSectionGeneration -v
```

Expected: FAIL — `_build_section_prompt` does not exist.

- [ ] **Step 3: Implement section generation**

Add these methods to `MediationBriefGenerator` in `icharlotte_core/mediation_brief.py`:

```python
    # ── Hard-coded style guide ──────────────────────────────────

    STYLE_GUIDE = """
STYLE AND TONE GUIDE:
- Write as a senior defense litigation attorney addressing a mediator
- Tone: professional, authoritative, persuasive, and firm
- Use active voice and strong declarative statements
- Avoid hedging or weak qualifiers ("perhaps", "might", "it could be argued")
- Present defense arguments as the clear and logical reading of the evidence
- When discussing plaintiff's position, highlight inconsistencies and weaknesses
- Use specific facts, dates, and evidence — avoid vague generalities
- When including deposition quotes, clean them of transcript artifacts (line numbers, dashes, extra characters) but keep them verbatim
- Citation format for deposition quotes: (LastName Depo Trns., at p. PageNum:LineNum.)
- Do not use placeholder text like [TBD] or [INSERT] — write around missing information naturally
- Be thorough and detailed — length is not a concern
"""

    FORMATTING_RULES = """
FORMATTING RULES:
- Do NOT include level-one section headings (roman numerals) — they are added by the system
- For the LIABILITY and DAMAGES sections: mark each subsection with "SUBSECTION: Title Text" on its own line, followed by the content paragraphs. The system will convert these to properly formatted level-two headings.
- For deposition quotes: put the quote on its own paragraph, preceded by a blank line. Follow with the citation on the same paragraph: (LastName Depo Trns., at p. PageNum:LineNum.)
- Write in plain text. Do not use markdown formatting (no **, ##, etc.)
"""

    # ── Prompt construction ─────────────────────────────────────

    def _get_section_prompt(self, section_name: str) -> str:
        """Load the editable main prompt for a section from the Workbench."""
        from icharlotte_core.prompt_manager import get_prompt_manager
        pm = get_prompt_manager()
        prompt = pm.get_prompt("mediation_brief", section_name)
        if prompt:
            return prompt
        # Fallback: read directly from file
        path = os.path.join(PROMPTS_DIR, f"{section_name}_current.txt")
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return f.read()
        log.warning("No prompt found for section: %s", section_name)
        return ""

    def _build_system_prompt(self, section_name: str) -> str:
        """Build the hard-coded system prompt for a section."""
        return (
            "You are a senior defense litigation attorney drafting a section of a "
            "confidential mediation brief. Follow the style guide and formatting rules exactly.\n\n"
            + self.STYLE_GUIDE + "\n"
            + self.FORMATTING_RULES
        )

    def _build_section_prompt(self, section_name: str, refinement_instruction: str = "") -> str:
        """Build the full prompt for generating a section.

        Combines: main prompt + planning output + style excerpts +
        previously generated sections + document content.
        """
        parts = []

        # 1. Main prompt from Workbench
        main_prompt = self._get_section_prompt(section_name)
        if main_prompt:
            parts.append(main_prompt)

        # 2. Refinement instruction (if refining)
        if refinement_instruction:
            parts.append(f"\nADDITIONAL INSTRUCTION FROM USER:\n{refinement_instruction}")

        # 3. Style excerpts from sample briefs
        excerpts = self.get_style_excerpts()
        section_excerpts = excerpts.get(section_name, [])
        if section_excerpts:
            parts.append("\n--- EXAMPLE FROM SAMPLE BRIEF ---")
            # Include first excerpt (truncate at 10000 chars)
            parts.append(section_excerpts[0][:10000])
            parts.append("--- END EXAMPLE ---")

        # 4. Planning pass output
        if self.planning_output:
            parts.append(f"\n--- EXTRACTED CASE INFORMATION ---\n{self.planning_output}")

        # 5. Previously generated sections
        prior_sections = []
        for name in GENERATION_ORDER:
            if name == section_name:
                break
            if name in self.sections and self.sections[name]:
                roman, title = SECTION_HEADINGS[name]
                prior_sections.append(f"{roman}. {title}\n{self.sections[name]}")

        # For introduction (last), include ALL other sections
        if section_name == "introduction":
            prior_sections = []
            for name in SECTION_ORDER:
                if name == "introduction":
                    continue
                if name in self.sections and self.sections[name]:
                    roman, title = SECTION_HEADINGS[name]
                    prior_sections.append(f"{roman}. {title}\n{self.sections[name]}")

        if prior_sections:
            parts.append("\n--- PREVIOUSLY GENERATED SECTIONS ---")
            parts.append("\n\n".join(prior_sections))
            parts.append("--- END PREVIOUSLY GENERATED SECTIONS ---")

        return "\n\n".join(parts)

    def generate_section(self, section_name: str, refinement_instruction: str = "") -> str:
        """Generate a single section using LLMCaller.

        Args:
            section_name: One of the SECTION_ORDER names
            refinement_instruction: Optional additional instruction for refinement

        Returns:
            Generated section text.
        """
        from icharlotte_core.llm_config import LLMCaller

        caller = LLMCaller()
        system = self._build_system_prompt(section_name)
        prompt = self._build_section_prompt(section_name, refinement_instruction)

        result = caller.call(
            prompt=prompt,
            text=self.document_content,
            agent_id="agent_mediation_brief",
        )
        return result or ""

    def run_planning_pass(self) -> str:
        """Run the planning pass to extract structured case information.

        Returns:
            Planning output text with extracted facts, arguments, quotes, dates.
        """
        from icharlotte_core.llm_config import LLMCaller

        caller = LLMCaller()
        prompt = self._get_section_prompt("planning")
        system = (
            "You are a senior defense litigation attorney analyzing case documents "
            "for a mediation brief. Extract all relevant information as instructed."
        )

        result = caller.call(
            prompt=prompt,
            text=self.document_content,
            agent_id="agent_mediation_brief",
        )
        self.planning_output = result or ""
        return self.planning_output
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_mediation_brief.py::TestSectionGeneration -v
```

Expected: All 4 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief.py tests/test_mediation_brief.py
git commit -m "feat(mediation-brief): section generation pipeline with prompt construction"
```

---

### Task 5: Word Document Assembly

**Files:**
- Modify: `icharlotte_core/mediation_brief.py`
- Modify: `tests/test_mediation_brief.py`

- [ ] **Step 1: Write tests for text parsing and Word formatting**

Add to `tests/test_mediation_brief.py`:

```python
class TestTextParsing(unittest.TestCase):
    """Test parsing LLM output into structured elements."""

    def test_parse_body_paragraphs(self):
        """Plain text is parsed as body paragraphs."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        text = "First paragraph of the section.\n\nSecond paragraph here."
        elements = gen._parse_section_text(text, "statement_of_facts")
        body_elements = [e for e in elements if e["type"] == "body"]
        self.assertEqual(len(body_elements), 2)
        self.assertIn("First paragraph", body_elements[0]["text"])

    def test_parse_subsection_headings(self):
        """SUBSECTION: markers are parsed as level-two headings."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        text = (
            "Introduction paragraph.\n\n"
            "SUBSECTION: Defendant Owed No Duty Of Care\n"
            "Content under first subsection.\n\n"
            "SUBSECTION: Plaintiff Was Comparatively At Fault\n"
            "Content under second subsection.\n"
        )
        elements = gen._parse_section_text(text, "liability")
        headings = [e for e in elements if e["type"] == "l2_heading"]
        self.assertEqual(len(headings), 2)
        self.assertEqual(headings[0]["text"], "Defendant Owed No Duty Of Care")
        self.assertEqual(headings[1]["text"], "Plaintiff Was Comparatively At Fault")

    def test_parse_deposition_quotes(self):
        """Lines with depo citation pattern are parsed as quotes."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        text = (
            "The plaintiff testified as follows:\n\n"
            "I was not paying attention to the road at the time of the accident. "
            "(Smith Depo Trns., at p. 45:12.)\n\n"
            "This admission is significant."
        )
        elements = gen._parse_section_text(text, "liability")
        quotes = [e for e in elements if e["type"] == "depo_quote"]
        self.assertEqual(len(quotes), 1)
        self.assertIn("not paying attention", quotes[0]["text"])


class TestWordAssembly(unittest.TestCase):
    """Test Word document assembly."""

    def test_assemble_creates_docx(self):
        """Assembly produces a valid .docx file."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        gen.sections = {
            "introduction": "This is the introduction.",
            "statement_of_facts": "These are the facts.",
            "procedural_status": "Trial is set for June 2027.",
            "liability": "SUBSECTION: No Duty\nDefendant owed no duty.",
            "damages": "SUBSECTION: No Causation\nNo evidence of causation.",
            "settlement_position": "Policy limits are $1M.",
            "conclusion": "We will prevail at trial.",
        }

        with tempfile.TemporaryDirectory() as tmpdir:
            # Create a minimal caption template
            caption = Document()
            caption.add_paragraph("LAW FIRM")
            caption.add_paragraph("CAPTION PAGE")
            caption_path = os.path.join(tmpdir, "caption.docx")
            caption.save(caption_path)

            output_path = os.path.join(tmpdir, "brief.docx")
            gen.assemble_document(caption_path, output_path)

            self.assertTrue(os.path.exists(output_path))
            doc = Document(output_path)
            all_text = "\n".join(p.text for p in doc.paragraphs)
            self.assertIn("INTRODUCTION", all_text)
            self.assertIn("STATEMENT OF FACTS", all_text)
            self.assertIn("introduction", all_text.lower())
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_mediation_brief.py::TestTextParsing tests/test_mediation_brief.py::TestWordAssembly -v
```

Expected: FAIL — `_parse_section_text` and `assemble_document` don't exist.

- [ ] **Step 3: Implement text parsing**

Add to `MediationBriefGenerator` in `icharlotte_core/mediation_brief.py`:

```python
    # ── Text parsing ────────────────────────────────────────────

    # Pattern for SUBSECTION: Title
    _SUBSECTION_RE = re.compile(r'^SUBSECTION:\s*(.+)$', re.MULTILINE)

    # Pattern for deposition citations
    _DEPO_CITE_RE = re.compile(
        r'\([A-Z][a-z]+ Depo Trns\., at p\. \d+:\d+\.\)'
    )

    def _parse_section_text(self, text: str, section_name: str) -> List[Dict]:
        """Parse LLM output text into structured elements.

        Returns list of dicts with keys:
            type: 'body', 'l2_heading', 'depo_quote'
            text: the content
        """
        elements = []
        # Split into paragraphs
        paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]

        for para in paragraphs:
            # Check for subsection heading
            sub_match = self._SUBSECTION_RE.match(para)
            if sub_match:
                elements.append({
                    "type": "l2_heading",
                    "text": sub_match.group(1).strip(),
                })
                # If there's content after the heading line in the same paragraph
                remaining = para[sub_match.end():].strip()
                if remaining:
                    elements.append({"type": "body", "text": remaining})
                continue

            # Check for deposition quote (paragraph containing a depo citation)
            if self._DEPO_CITE_RE.search(para):
                elements.append({"type": "depo_quote", "text": para})
                continue

            # Regular body paragraph
            elements.append({"type": "body", "text": para})

        return elements
```

- [ ] **Step 4: Implement Word document assembly**

Add to `MediationBriefGenerator` in `icharlotte_core/mediation_brief.py`:

```python
    # ── Word document assembly ──────────────────────────────────

    def assemble_document(
        self,
        caption_path: str,
        output_path: str,
        signature_paragraphs: Optional[list] = None,
    ):
        """Assemble the final Word document from caption template and sections.

        Args:
            caption_path: Path to the original caption template
            output_path: Path to save the assembled document
            signature_paragraphs: Optional extracted signature block elements
        """
        from docx import Document
        from docx.shared import Inches, Pt, Emu
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        from docx.enum.text import WD_ALIGN_PARAGRAPH

        # Prepare the caption template (returns sig paragraphs if found)
        sig_paras = self.prepare_caption_template(caption_path, output_path)
        if signature_paragraphs is not None:
            sig_paras = signature_paragraphs

        doc = Document(output_path)

        # Add page break after caption
        doc.add_page_break()

        # Insert sections in document order
        for section_name in SECTION_ORDER:
            section_text = self.sections.get(section_name, "")
            if not section_text:
                continue

            roman, title = SECTION_HEADINGS[section_name]

            # Add level-one heading
            self._add_l1_heading(doc, roman, title)

            # Parse and insert section content
            elements = self._parse_section_text(section_text, section_name)
            letter_counter = 0
            for elem in elements:
                if elem["type"] == "l2_heading":
                    letter = chr(ord("A") + letter_counter)
                    self._add_l2_heading(doc, letter, elem["text"])
                    letter_counter += 1
                elif elem["type"] == "depo_quote":
                    self._add_depo_quote(doc, elem["text"])
                else:
                    self._add_body_paragraph(doc, elem["text"])

        # Append signature block if present
        if sig_paras:
            # Add spacing before signature
            doc.add_paragraph()
            for sig_para in sig_paras:
                # Re-add the XML elements to the document body
                doc.element.body.append(sig_para._element)

        doc.save(output_path)
        log.info("Mediation brief assembled at %s", output_path)

        # Validate per project conventions
        try:
            from icharlotte_core.word_validator import validate_report
            result = validate_report(output_path)
            result.print_summary()
        except Exception as e:
            log.warning("Validation skipped: %s", e)

    def _add_l1_heading(self, doc, roman: str, title: str):
        """Add a Level 1 heading: 'I.     INTRODUCTION' format.

        Roman numeral + period bold, tab, title bold + underlined, all caps.
        0.5 inch hanging indent.
        """
        from docx.shared import Inches, Pt
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(12)
        para.paragraph_format.space_after = Pt(6)

        # Set hanging indent: first line at 0, left indent at 0.5"
        pPr = para._element.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(int(Inches(0.5))))
        ind.set(qn("w:hanging"), str(int(Inches(0.5))))
        pPr.append(ind)

        # Add tab stop at 0.5"
        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "left")
        tab.set(qn("w:pos"), str(int(Inches(0.5))))
        tabs.append(tab)
        pPr.append(tabs)

        # Run 1: roman numeral + period (bold only)
        run1 = para.add_run(f"{roman}.")
        run1.bold = True

        # Tab character
        run_tab = para.add_run("\t")

        # Run 2: title (bold + underlined)
        run2 = para.add_run(title)
        run2.bold = True
        run2.underline = True

    def _add_l2_heading(self, doc, letter: str, title: str):
        """Add a Level 2 heading: 'A.     Title Text' format.

        Letter + period bold, tab, title bold + underlined, title case.
        0.5 inch hanging indent.
        """
        from docx.shared import Inches, Pt
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        para = doc.add_paragraph()
        para.paragraph_format.space_before = Pt(10)
        para.paragraph_format.space_after = Pt(4)

        pPr = para._element.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), str(int(Inches(0.5))))
        ind.set(qn("w:hanging"), str(int(Inches(0.5))))
        pPr.append(ind)

        tabs = OxmlElement("w:tabs")
        tab = OxmlElement("w:tab")
        tab.set(qn("w:val"), "left")
        tab.set(qn("w:pos"), str(int(Inches(0.5))))
        tabs.append(tab)
        pPr.append(tabs)

        # Run 1: letter + period (bold only)
        run1 = para.add_run(f"{letter}.")
        run1.bold = True

        # Tab
        para.add_run("\t")

        # Run 2: title (bold + underlined)
        run2 = para.add_run(title)
        run2.bold = True
        run2.underline = True

    def _add_body_paragraph(self, doc, text: str):
        """Add a normal body paragraph."""
        from docx.shared import Pt

        para = doc.add_paragraph(text)
        para.paragraph_format.space_after = Pt(6)

    def _add_depo_quote(self, doc, text: str):
        """Add a deposition quote paragraph with 0.5 inch left indent."""
        from docx.shared import Inches, Pt

        para = doc.add_paragraph(text)
        para.paragraph_format.left_indent = Inches(0.5)
        para.paragraph_format.space_before = Pt(6)
        para.paragraph_format.space_after = Pt(6)
```

- [ ] **Step 5: Run tests to verify they pass**

```bash
python -m pytest tests/test_mediation_brief.py::TestTextParsing tests/test_mediation_brief.py::TestWordAssembly -v
```

Expected: All 4 tests PASS.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/mediation_brief.py tests/test_mediation_brief.py
git commit -m "feat(mediation-brief): text parsing and Word document assembly"
```

---

### Task 6: Background Worker and Full Pipeline

**Files:**
- Modify: `icharlotte_core/mediation_brief.py`
- Modify: `tests/test_mediation_brief.py`

- [ ] **Step 1: Write tests for the full pipeline**

Add to `tests/test_mediation_brief.py`:

```python
class TestPipeline(unittest.TestCase):
    """Test the full generation pipeline orchestration."""

    def test_generate_all_sections_order(self):
        """Sections are generated in correct order with Introduction last."""
        from icharlotte_core.mediation_brief import (
            MediationBriefGenerator, GENERATION_ORDER
        )

        gen = MediationBriefGenerator()
        gen.document_content = "Test document content"
        gen._style_cache = {"sections": {}}

        call_order = []
        def mock_generate(section_name, refinement_instruction=""):
            call_order.append(section_name)
            return f"Generated {section_name}"

        with patch.object(gen, 'generate_section', side_effect=mock_generate):
            with patch.object(gen, 'run_planning_pass', return_value="Planning output"):
                gen.generate_all_sections()

        self.assertEqual(call_order, GENERATION_ORDER)
        self.assertEqual(gen.sections["introduction"], "Generated introduction")
        self.assertEqual(gen.sections["conclusion"], "Generated conclusion")

    def test_pipeline_sets_active_flag(self):
        """After generation completes, is_active is True."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        gen.document_content = "Test doc"
        gen._style_cache = {"sections": {}}

        with patch.object(gen, 'generate_section', return_value="text"):
            with patch.object(gen, 'run_planning_pass', return_value="plan"):
                gen.generate_all_sections()

        self.assertTrue(gen.is_active)
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_mediation_brief.py::TestPipeline -v
```

Expected: FAIL — `generate_all_sections` does not exist.

- [ ] **Step 3: Implement pipeline and worker**

Add to `MediationBriefGenerator` in `icharlotte_core/mediation_brief.py`:

```python
    # ── Full pipeline ───────────────────────────────────────────

    def generate_all_sections(
        self,
        progress_callback=None,
    ):
        """Generate all sections sequentially.

        Args:
            progress_callback: Optional callable(section_name, index, total)
                called before each section generation.
        """
        total = len(GENERATION_ORDER) + 1  # +1 for planning pass

        # Step 0: Planning pass
        if progress_callback:
            progress_callback("planning", 0, total)
        self.run_planning_pass()

        # Steps 1-7: Generate sections
        for i, section_name in enumerate(GENERATION_ORDER):
            if progress_callback:
                progress_callback(section_name, i + 1, total)
            self.sections[section_name] = self.generate_section(section_name)

        self.is_active = True

    def reset(self):
        """Clear all state for a fresh generation."""
        self.sections = {}
        self.planning_output = ""
        self.document_content = ""
        self.caption_template_path = None
        self.is_active = False
```

Now add the `MediationBriefWorker` class at module level (after `MediationBriefGenerator`):

```python
try:
    from PySide6.QtCore import QThread, Signal
except ImportError:
    from PyQt6.QtCore import QThread
    from PyQt6.QtCore import pyqtSignal as Signal


class MediationBriefWorker(QThread):
    """Background worker for mediation brief generation.

    Signals:
        section_started(str, int, int): section_name, index, total
        section_complete(str, str): section_name, section_text
        all_complete(dict): {section_name: text} for all sections
        error(str): error message
    """
    section_started = Signal(str, int, int)
    section_complete = Signal(str, str)
    all_complete = Signal(dict)
    error = Signal(str)

    def __init__(self, generator: MediationBriefGenerator, parent=None):
        super().__init__(parent)
        self.generator = generator
        self._stop_requested = False

    def request_stop(self):
        self._stop_requested = True

    def run(self):
        try:
            total = len(GENERATION_ORDER) + 1

            # Planning pass
            self.section_started.emit("planning", 0, total)
            self.generator.run_planning_pass()

            # Generate sections
            for i, section_name in enumerate(GENERATION_ORDER):
                if self._stop_requested:
                    return
                self.section_started.emit(section_name, i + 1, total)
                text = self.generator.generate_section(section_name)
                self.generator.sections[section_name] = text
                self.section_complete.emit(section_name, text)

            self.generator.is_active = True
            self.all_complete.emit(dict(self.generator.sections))
        except Exception as e:
            log.exception("Mediation brief generation failed")
            self.error.emit(str(e))
```

Move the Qt imports to the top of the file inside a try/except block:

```python
try:
    from PySide6.QtCore import QThread, Signal
except ImportError:
    from PyQt6.QtCore import QThread
    from PyQt6.QtCore import pyqtSignal as Signal
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_mediation_brief.py::TestPipeline -v
```

Expected: All 2 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief.py tests/test_mediation_brief.py
git commit -m "feat(mediation-brief): full pipeline orchestration and background worker"
```

---

### Task 7: Conversational Refinement

**Files:**
- Modify: `icharlotte_core/mediation_brief.py`
- Modify: `tests/test_mediation_brief.py`

- [ ] **Step 1: Write tests for refinement routing and regeneration**

Add to `tests/test_mediation_brief.py`:

```python
class TestRefinement(unittest.TestCase):
    """Test conversational refinement."""

    def test_route_sections_parses_comma_list(self):
        """Routing response with comma-separated sections is parsed correctly."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        result = gen._parse_routing_response("damages,liability")
        self.assertEqual(result, ["damages", "liability"])

    def test_route_sections_handles_none(self):
        """'none' routing response returns empty list."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        result = gen._parse_routing_response("none")
        self.assertEqual(result, [])

    def test_route_sections_filters_invalid(self):
        """Invalid section names are filtered out."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        result = gen._parse_routing_response("damages,invalid_section,liability")
        self.assertEqual(result, ["damages", "liability"])

    def test_refine_regenerates_introduction_when_liability_changes(self):
        """Refining liability also triggers introduction regeneration."""
        from icharlotte_core.mediation_brief import MediationBriefGenerator

        gen = MediationBriefGenerator()
        gen.sections = {
            "introduction": "Old intro",
            "statement_of_facts": "Facts",
            "procedural_status": "Status",
            "liability": "Old liability",
            "damages": "Old damages",
            "settlement_position": "Settlement",
            "conclusion": "Conclusion",
        }
        gen.planning_output = "Planning"
        gen.document_content = "Doc"
        gen._style_cache = {"sections": {}}
        gen.is_active = True

        call_log = []
        def mock_generate(section_name, refinement_instruction=""):
            call_log.append(section_name)
            return f"New {section_name}"

        with patch.object(gen, 'generate_section', side_effect=mock_generate):
            gen.refine_sections(["liability"], "Make it stronger")

        self.assertIn("liability", call_log)
        self.assertIn("introduction", call_log)
        self.assertEqual(gen.sections["liability"], "New liability")
        self.assertEqual(gen.sections["introduction"], "New introduction")
```

- [ ] **Step 2: Run tests to verify they fail**

```bash
python -m pytest tests/test_mediation_brief.py::TestRefinement -v
```

Expected: FAIL — `_parse_routing_response` and `refine_sections` don't exist.

- [ ] **Step 3: Implement refinement**

Add to `MediationBriefGenerator` in `icharlotte_core/mediation_brief.py`:

```python
    # ── Conversational refinement ───────────────────────────────

    # Sections whose changes should trigger Introduction regeneration
    _INTRO_TRIGGERS = {"liability", "damages", "statement_of_facts", "conclusion"}

    def _parse_routing_response(self, response: str) -> List[str]:
        """Parse the routing LLM response into a list of valid section names.

        Args:
            response: Raw LLM response (e.g., "damages,liability" or "none")

        Returns:
            List of valid section names, or empty list if "none".
        """
        cleaned = response.strip().lower()
        if cleaned == "none":
            return []

        valid_sections = set(SECTION_ORDER)
        names = [n.strip() for n in cleaned.split(",")]
        return [n for n in names if n in valid_sections]

    def route_refinement(self, user_message: str) -> List[str]:
        """Determine which sections to regenerate based on user's message.

        Uses LLM routing call with the routing prompt.

        Args:
            user_message: The user's refinement instruction

        Returns:
            List of section names to regenerate, or empty list if not a brief refinement.
        """
        from icharlotte_core.llm_config import LLMCaller

        caller = LLMCaller()
        routing_prompt = self._get_section_prompt("routing")
        full_prompt = f"{routing_prompt}\n\nUser message: {user_message}"

        response = caller.call(
            prompt=full_prompt,
            text="",
            agent_id="agent_mediation_brief",
        )
        if not response:
            return []

        return self._parse_routing_response(response)

    def refine_sections(
        self,
        section_names: List[str],
        instruction: str,
        progress_callback=None,
    ) -> List[str]:
        """Regenerate specified sections with the user's refinement instruction.

        If any of the regenerated sections are in _INTRO_TRIGGERS,
        the Introduction is also regenerated.

        Args:
            section_names: Sections to regenerate
            instruction: User's refinement instruction
            progress_callback: Optional callable(section_name, index, total)

        Returns:
            List of all section names that were actually regenerated.
        """
        regenerated = list(section_names)

        # Check if Introduction needs regeneration too
        needs_intro = any(s in self._INTRO_TRIGGERS for s in section_names)
        if needs_intro and "introduction" not in regenerated:
            regenerated.append("introduction")

        # Ensure introduction is last if it's being regenerated
        if "introduction" in regenerated:
            regenerated.remove("introduction")
            regenerated.append("introduction")

        total = len(regenerated)
        for i, section_name in enumerate(regenerated):
            if progress_callback:
                progress_callback(section_name, i, total)

            # Only apply instruction to the sections the user asked about
            instr = instruction if section_name in section_names else ""
            self.sections[section_name] = self.generate_section(
                section_name, refinement_instruction=instr
            )

        return regenerated
```

- [ ] **Step 4: Run tests to verify they pass**

```bash
python -m pytest tests/test_mediation_brief.py::TestRefinement -v
```

Expected: All 4 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/mediation_brief.py tests/test_mediation_brief.py
git commit -m "feat(mediation-brief): conversational refinement with LLM routing"
```

---

### Task 8: ChatTab Integration — Menu, Confirmation Dialog, and Generation Hooks

**Files:**
- Modify: `icharlotte_core/ui/tabs.py`

This is the largest integration task. It wires the mediation brief generator into the ChatTab UI.

- [ ] **Step 1: Add import to tabs.py**

At the top of `icharlotte_core/ui/tabs.py`, add with the other icharlotte_core imports:

```python
from icharlotte_core.mediation_brief import MediationBriefGenerator, MediationBriefWorker
```

- [ ] **Step 2: Add state variables to ChatTab.__init__()**

In the `ChatTab.__init__()` method (~line 237, after `self.worker = None`), add:

```python
        # Mediation brief state
        self.med_brief_generator = None
        self.med_brief_worker = None
```

- [ ] **Step 3: Modify update_template_menu() to add Mediation Brief**

In `update_template_menu()` (~line 1720), after the built-in prompts loop but before the custom prompts section, add a separator and the Mediation Brief action. Find this code block:

```python
    # Built-in prompts
    for prompt in BUILTIN_PROMPTS:
        action = QAction(prompt.name, self)
        action.triggered.connect(lambda checked, p=prompt: self.insert_template(p.prompt))
        menu.addAction(action)
```

After it, add:

```python
    # Mediation Brief (special — triggers generation, not text insert)
    menu.addSeparator()
    med_brief_action = QAction("Mediation Brief", self)
    med_brief_action.triggered.connect(self._on_mediation_brief_selected)
    menu.addAction(med_brief_action)
```

- [ ] **Step 4: Add the confirmation dialog method**

Add this method to the `ChatTab` class:

```python
    def _on_mediation_brief_selected(self):
        """Handle Mediation Brief template selection — show confirmation dialog."""
        from PySide6.QtWidgets import QMessageBox

        # Check for selected documents
        checked_files = []
        for i in range(self.file_list.count()):
            item = self.file_list.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                path = item.data(Qt.ItemDataRole.UserRole)
                if path:
                    checked_files.append(os.path.basename(path))

        if not checked_files:
            QMessageBox.warning(
                self, "No Documents",
                "Please select documents in the file list before generating a mediation brief."
            )
            return

        # Build confirmation message
        case_name = self.file_number or "Unknown Case"
        file_list_text = "\n".join(f"  - {f}" for f in checked_files[:15])
        if len(checked_files) > 15:
            file_list_text += f"\n  ... and {len(checked_files) - 15} more"

        msg = (
            f"Generate a Mediation Brief for case {case_name}?\n\n"
            f"Documents ({len(checked_files)}):\n{file_list_text}\n\n"
            "This will generate a comprehensive brief section-by-section. "
            "The process may take several minutes."
        )

        reply = QMessageBox.question(
            self, "Generate Mediation Brief", msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self._start_mediation_brief_generation()
```

- [ ] **Step 5: Add the generation start method**

Add this method to the `ChatTab` class:

```python
    def _start_mediation_brief_generation(self):
        """Start the mediation brief generation pipeline."""
        from PySide6.QtWidgets import QFileDialog, QMessageBox

        # Find caption template
        main_win = self.window()
        case_path = getattr(main_win, 'case_path', None)
        parent_folder = os.path.dirname(case_path) if case_path else None

        generator = MediationBriefGenerator()

        caption_path = None
        if parent_folder:
            caption_path = generator.find_caption_template(parent_folder)

        if not caption_path:
            search_loc = parent_folder or ""
            caption_path, _ = QFileDialog.getOpenFileName(
                self,
                f"No caption template found in {search_loc}. Select one:",
                search_loc,
                "Word Documents (*.docx)"
            )
            if not caption_path:
                return  # User cancelled

        generator.caption_template_path = caption_path

        # Read document content
        generator.document_content = self.read_files_content()
        if not generator.document_content.strip():
            QMessageBox.warning(
                self, "No Content",
                "Could not read any content from the selected documents."
            )
            return

        # Cache style excerpts (first run extracts from samples)
        generator.get_style_excerpts()

        # Store generator reference
        self.med_brief_generator = generator

        # Display start message in chat
        self.chat_history.append("<b>Mediation Brief Generator</b>")
        self.chat_history.append("<i>Starting generation...</i>")
        self.chat_history.append("")

        # Disable send button during generation
        self.send_btn.setEnabled(False)

        # Start background worker
        self.med_brief_worker = MediationBriefWorker(generator, parent=self)
        self.med_brief_worker.section_started.connect(self._on_brief_section_started)
        self.med_brief_worker.section_complete.connect(self._on_brief_section_complete)
        self.med_brief_worker.all_complete.connect(self._on_brief_all_complete)
        self.med_brief_worker.error.connect(self._on_brief_error)
        self.med_brief_worker.start()
```

- [ ] **Step 6: Add signal handler methods**

Add these methods to `ChatTab`:

```python
    def _on_brief_section_started(self, section_name: str, index: int, total: int):
        """Display progress when a section starts generating."""
        from icharlotte_core.mediation_brief import SECTION_HEADINGS
        if section_name == "planning":
            self.chat_history.append("<i>Analyzing documents...</i>")
        else:
            heading = SECTION_HEADINGS.get(section_name, ("", section_name.upper()))
            display = f"{heading[0]}. {heading[1]}" if heading[0] else section_name
            self.chat_history.append(f"<i>Generating {display} ({index} of {total - 1})...</i>")

    def _on_brief_section_complete(self, section_name: str, text: str):
        """Display completed section text in chat."""
        import markdown
        from icharlotte_core.mediation_brief import SECTION_HEADINGS

        heading = SECTION_HEADINGS.get(section_name)
        if heading:
            self.chat_history.append(f"<br><b>{heading[0]}.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;{heading[1]}</b>")

        try:
            html = markdown.markdown(text, extensions=['fenced_code', 'tables'])
        except Exception:
            html = text.replace('\n', '<br>')
        self.chat_history.append(html)
        self.chat_history.append("<hr>")
        self.chat_history.ensureCursorVisible()

    def _on_brief_all_complete(self, sections: dict):
        """Handle completion of all sections — assemble document and save."""
        import tempfile
        from PySide6.QtWidgets import QFileDialog

        self.send_btn.setEnabled(True)
        self.chat_history.append("<b>All sections generated. Assembling document...</b>")

        gen = self.med_brief_generator

        # Determine default save location
        main_win = self.window()
        case_path = getattr(main_win, 'case_path', None)
        default_dir = os.path.dirname(case_path) if case_path else ""
        default_name = os.path.join(default_dir, "Defendant's Confidential Mediation Brief.docx")

        # Assemble to temp first, then Save As
        with tempfile.TemporaryDirectory() as tmpdir:
            temp_output = os.path.join(tmpdir, "mediation_brief.docx")
            try:
                gen.assemble_document(gen.caption_template_path, temp_output)
            except Exception as e:
                self.chat_history.append(f"<b style='color:red'>Assembly error: {e}</b>")
                return

            # Save As dialog
            save_path, _ = QFileDialog.getSaveFileName(
                self,
                "Save Mediation Brief",
                default_name,
                "Word Documents (*.docx)"
            )
            if save_path:
                shutil.copy2(temp_output, save_path)
                self.chat_history.append(f"<b>Mediation brief saved to:</b> {save_path}")
            else:
                self.chat_history.append("<i>Save cancelled. You can refine sections and save later.</i>")

        self.chat_history.append("")
        self.chat_history.append("<i>You can now refine the brief by typing instructions "
                                 "(e.g., 'make the Damages section more aggressive'). "
                                 "Or send a normal message to exit brief mode.</i>")

    def _on_brief_error(self, error_msg: str):
        """Handle generation error."""
        self.send_btn.setEnabled(True)
        self.chat_history.append(f"<b style='color:red'>Generation error: {error_msg}</b>")
```

Add `import shutil` at the top of `tabs.py` if not already present.

- [ ] **Step 7: Add refinement routing to send_message()**

In `send_message()` (~line 1212), add a check at the very beginning of the method (before any other processing) to intercept messages when in refinement mode:

```python
    def send_message(self):
        # --- Mediation brief refinement mode ---
        if (self.med_brief_generator and self.med_brief_generator.is_active
                and not self.med_brief_worker or not self.med_brief_worker.isRunning()):
            user_text = self.chat_input.toPlainText().strip()
            if user_text:
                sections = self.med_brief_generator.route_refinement(user_text)
                if sections:
                    self._start_brief_refinement(user_text, sections)
                    return
        # --- End mediation brief check ---

        # ... existing send_message code continues ...
```

Add the refinement handler:

```python
    def _start_brief_refinement(self, instruction: str, section_names: list):
        """Regenerate specified sections with user's instruction."""
        from icharlotte_core.mediation_brief import SECTION_HEADINGS

        self.chat_input.clear()
        self.chat_history.append(f"<b>You:</b> {instruction}")
        self.chat_history.append("")

        section_display = ", ".join(
            SECTION_HEADINGS.get(s, ("", s))[1] for s in section_names
        )
        self.chat_history.append(
            f"<i>Regenerating: {section_display}...</i>"
        )

        self.send_btn.setEnabled(False)

        # Run refinement in background using MediationBriefWorker pattern
        gen = self.med_brief_generator

        # Reuse MediationBriefWorker but override run() for refinement
        class _RefinementWorker(QThread):
            """One-off worker for section refinement."""
            section_done = Signal(str, str)
            all_done = Signal(list)
            error = Signal(str)

            def __init__(self, gen, sections, instruction, parent=None):
                super().__init__(parent)
                self.gen = gen
                self._sections = sections
                self._instruction = instruction

            def run(self):
                try:
                    regenerated = self.gen.refine_sections(
                        self._sections, self._instruction,
                    )
                    for name in regenerated:
                        self.section_done.emit(name, self.gen.sections[name])
                    self.all_done.emit(regenerated)
                except Exception as e:
                    self.error.emit(str(e))

        worker = _RefinementWorker(gen, section_names, instruction, parent=self)
        worker.section_done.connect(self._on_brief_section_complete)
        worker.all_done.connect(lambda regenerated: self._on_refinement_complete(regenerated))
        worker.error.connect(self._on_brief_error)
        self.med_brief_worker = worker
        worker.start()

    def _on_refinement_complete(self, regenerated: list):
        """Handle refinement completion — reassemble and offer save."""
        self._on_brief_all_complete(self.med_brief_generator.sections)
```

- [ ] **Step 8: Clear refinement state on case switch**

In the `load_case()` method (~line 482), add after the existing cleanup:

```python
        # Clear mediation brief state
        self.med_brief_generator = None
        if self.med_brief_worker and self.med_brief_worker.isRunning():
            self.med_brief_worker.request_stop()
        self.med_brief_worker = None
```

- [ ] **Step 9: Test the integration manually**

Run the app and verify:
1. "Mediation Brief" appears in the Templates dropdown
2. Selecting it shows the confirmation dialog with document list
3. Cancel dismisses without action
4. Generate starts the pipeline (verify with a small test doc)

```bash
python iCharlotte.py
```

- [ ] **Step 10: Commit**

```bash
git add icharlotte_core/ui/tabs.py
git commit -m "feat(mediation-brief): ChatTab integration with menu, confirmation, and refinement"
```

---

### Task 9: End-to-End Testing and Polish

**Files:**
- Modify: `icharlotte_core/mediation_brief.py` (if fixes needed)
- Modify: `icharlotte_core/ui/tabs.py` (if fixes needed)
- Modify: `tests/test_mediation_brief.py`

- [ ] **Step 1: Run the full test suite**

```bash
python -m pytest tests/test_mediation_brief.py -v
```

Expected: All tests pass.

- [ ] **Step 2: Run the existing test suite to verify no regressions**

```bash
python -m pytest tests/ -v --timeout=60
```

Expected: No new failures.

- [ ] **Step 3: Manual end-to-end test**

1. Open iCharlotte with a case that has a caption template in the parent folder
2. Select several case documents in the Chat tab file list
3. Click Templates -> Mediation Brief
4. Verify confirmation dialog shows correct info
5. Click Generate and watch sections stream to chat
6. Verify Save As dialog opens and defaults to case parent folder
7. Save the document and open it in Word
8. Verify:
   - Caption page has "DEFENDANT'S CONFIDENTIAL MEDIATION BRIEF" (bold, CONFIDENTIAL underlined)
   - Line numbers are present
   - Level 1 headings have correct format (roman numeral, bold, underlined title, 0.5" indent)
   - Level 2 headings in Liability/Damages have correct format (letter, bold, underlined, restart at A.)
   - Deposition quotes are indented 0.5"
   - Signature block appears at the end (if present in original caption)
   - Footer has the replaced text
9. Test refinement: type "Make the Damages section stronger" and verify only Damages + Introduction regenerate
10. Test exit: type a non-brief message and verify normal chat behavior

```bash
python iCharlotte.py
```

- [ ] **Step 4: Fix any issues found during testing**

Address any formatting, Word assembly, or UI issues discovered during manual testing. Common areas to check:
- Caption templates with nested table layouts (XML iteration must work)
- Long section outputs (streaming display scrolling)
- Save As dialog when caption template is from a network drive
- Signature block detection edge cases

- [ ] **Step 5: Commit fixes**

```bash
git add -A
git commit -m "fix(mediation-brief): polish and end-to-end test fixes"
```

- [ ] **Step 6: Run all tests one final time**

```bash
python -m pytest tests/ -v --timeout=60
```

Expected: All tests pass.

- [ ] **Step 7: Final commit**

```bash
git add -A
git commit -m "feat(mediation-brief): complete mediation brief generator with refinement"
```
