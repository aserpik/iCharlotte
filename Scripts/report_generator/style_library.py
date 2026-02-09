"""
Style Library - Parses example reports to build a style reference library.

One-time setup: Extracts text section-by-section from the example reports,
selects the best examples per section type, and distills a style guide using LLM.

Also supports updating the library with new reports over time.

Usage:
    python -m Scripts.report_generator.style_library build [--reports-dir PATH]
    python -m Scripts.report_generator.style_library update <new_report.docx>
"""

import os
import sys
import json
import re
import logging
from typing import Dict, List, Optional, Tuple
from docx import Document

logger = logging.getLogger(__name__)

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DEFAULT_REPORTS_DIR = os.path.join(PROJECT_ROOT, "autodownloadreports")
STYLE_GUIDE_PATH = os.path.join(PROJECT_ROOT, "config", "report_style_guide.json")
STYLE_EXAMPLES_DIR = os.path.join(PROJECT_ROOT, "config", "report_style_examples")

# Section headings to look for (must match template_extractor.py)
REPORT_SECTIONS = [
    "FACTUAL BACKGROUND",
    "PROCEDURAL STATUS",
    "INVESTIGATION",
    "DISCOVERY",
    "MEDICAL RECORD REVIEW",
    "EVALUATION OF LIABILITY",
    "EVALUATION OF EXPOSURE",
    "SETTLEMENT",
    "SETTLEMENT STATUS",
    "FURTHER CASE HANDLING",
]

# Normalize section names (e.g., SETTLEMENT and SETTLEMENT STATUS are the same)
SECTION_NORMALIZE = {
    "SETTLEMENT STATUS": "SETTLEMENT",
}

MAX_EXAMPLES_PER_SECTION = 5


def extract_sections_from_docx(docx_path: str) -> Dict[str, str]:
    """
    Extract text content for each major section from a report .docx file.
    Returns dict mapping section name -> text content.
    """
    try:
        doc = Document(docx_path)
    except Exception as e:
        logger.warning(f"Could not open {docx_path}: {e}")
        return {}

    sections = {}
    current_section = None
    current_text = []

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            if current_section:
                current_text.append("")  # preserve paragraph breaks
            continue

        # Check if this is a section heading
        upper = text.upper()
        matched_section = None
        for heading in REPORT_SECTIONS:
            if upper == heading:
                matched_section = heading
                break

        if matched_section:
            # Save previous section
            if current_section and current_text:
                section_key = SECTION_NORMALIZE.get(current_section, current_section)
                sections[section_key] = "\n".join(current_text).strip()

            current_section = matched_section
            current_text = []
        elif current_section:
            current_text.append(text)

    # Save last section
    if current_section and current_text:
        section_key = SECTION_NORMALIZE.get(current_section, current_section)
        sections[section_key] = "\n".join(current_text).strip()

    return sections


def build_style_library(reports_dir: str = None) -> Dict:
    """
    Parse all reports in the directory and build the style reference library.
    Returns the full library structure.
    """
    if reports_dir is None:
        reports_dir = DEFAULT_REPORTS_DIR

    if not os.path.exists(reports_dir):
        raise FileNotFoundError(f"Reports directory not found: {reports_dir}")

    # Find all .docx files
    docx_files = [
        os.path.join(reports_dir, f)
        for f in os.listdir(reports_dir)
        if f.endswith('.docx') and not f.startswith('~$')
    ]

    if not docx_files:
        raise ValueError(f"No .docx files found in {reports_dir}")

    print(f"Parsing {len(docx_files)} reports...")

    # Collect all sections from all reports
    all_sections: Dict[str, List[Tuple[str, str]]] = {}  # section -> [(filename, text)]

    for filepath in docx_files:
        filename = os.path.basename(filepath)
        print(f"  Processing: {filename}")
        sections = extract_sections_from_docx(filepath)

        for section_name, text in sections.items():
            if not text or len(text) < 50:  # Skip very short sections
                continue
            if section_name not in all_sections:
                all_sections[section_name] = []
            all_sections[section_name].append((filename, text))

    # Select best examples per section (longest/most detailed, up to MAX_EXAMPLES)
    library = {}
    for section_name, examples in all_sections.items():
        # Sort by length (longer = more detailed), take top N
        examples.sort(key=lambda x: len(x[1]), reverse=True)
        selected = examples[:MAX_EXAMPLES_PER_SECTION]

        library[section_name] = {
            "total_examples_found": len(examples),
            "selected_count": len(selected),
            "examples": [
                {"source": src, "text": text, "char_count": len(text)}
                for src, text in selected
            ]
        }

    # Save examples
    os.makedirs(STYLE_EXAMPLES_DIR, exist_ok=True)
    for section_name, data in library.items():
        safe_name = section_name.lower().replace(" ", "_")
        filepath = os.path.join(STYLE_EXAMPLES_DIR, f"{safe_name}.json")
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        print(f"  Saved {data['selected_count']} examples for {section_name}")

    # Build summary
    summary = {
        "reports_parsed": len(docx_files),
        "sections": {
            name: {
                "total_found": data["total_examples_found"],
                "examples_stored": data["selected_count"]
            }
            for name, data in library.items()
        }
    }

    return summary


def distill_style_guide(llm_caller=None) -> Dict:
    """
    Use LLM to analyze the stored examples and distill a concise style guide.
    If llm_caller is None, generates a rule-based guide from pattern analysis.
    """
    # Load all examples
    if not os.path.exists(STYLE_EXAMPLES_DIR):
        raise FileNotFoundError("Style examples not found. Run 'build' first.")

    all_examples = {}
    for filename in os.listdir(STYLE_EXAMPLES_DIR):
        if filename.endswith('.json'):
            section_name = filename.replace('.json', '').replace('_', ' ').upper()
            filepath = os.path.join(STYLE_EXAMPLES_DIR, filename)
            with open(filepath, 'r', encoding='utf-8') as f:
                all_examples[section_name] = json.load(f)

    if llm_caller:
        return _distill_with_llm(all_examples, llm_caller)
    else:
        return _distill_rule_based(all_examples)


def _distill_rule_based(all_examples: Dict) -> Dict:
    """Generate a style guide using pattern analysis (no LLM required)."""
    guide = {
        "general": {
            "tone": "Formal, professional, objective — characteristic of legal defense litigation reports",
            "perspective": "Defense counsel perspective, written for insurance carrier audience",
            "voice": "Third person, passive constructions common ('Plaintiff alleges...', 'It is anticipated that...')",
            "hedging_phrases": [
                "We believe", "It is anticipated that", "It appears that",
                "Based on our review", "In our assessment", "It is our understanding"
            ],
            "formatting": {
                "font": "Times New Roman 12pt",
                "margins": "1 inch all sides",
                "section_headings": "ALL CAPS, bold, underlined",
                "sub_headings": "Title Case, bold, underlined (List Paragraph style)",
                "body_indent": "0.5 inch first line indent",
                "paragraph_spacing": "12pt after"
            }
        },
        "sections": {}
    }

    # Analyze each section's patterns
    for section_name, data in all_examples.items():
        examples = data.get("examples", [])
        if not examples:
            continue

        section_guide = _analyze_section_patterns(section_name, examples)
        guide["sections"][section_name] = section_guide

    # Save
    os.makedirs(os.path.dirname(STYLE_GUIDE_PATH), exist_ok=True)
    with open(STYLE_GUIDE_PATH, 'w', encoding='utf-8') as f:
        json.dump(guide, f, indent=2, ensure_ascii=False)

    print(f"Style guide saved to: {STYLE_GUIDE_PATH}")
    return guide


def _distill_with_llm(all_examples: Dict, llm_caller) -> Dict:
    """Use LLM to distill a comprehensive style guide from examples."""
    # Build a prompt with representative examples from each section
    section_summaries = []
    for section_name, data in all_examples.items():
        examples = data.get("examples", [])
        if not examples:
            continue
        # Use the first (longest) example, truncated to ~2000 chars
        sample = examples[0]["text"][:2000]
        section_summaries.append(f"### {section_name}\n{sample}\n")

    prompt = f"""Analyze these litigation report sections written by a defense attorney for insurance carriers.
Distill a concise style guide covering:

1. **Tone & Voice**: How formal? Active vs passive? Person?
2. **Common Phrases**: Recurring legal phrases, hedging language, transition words
3. **Structure Patterns**: How each section type is typically organized (sub-headings, paragraph flow)
4. **Detail Level**: How specific/granular is the analysis in each section?
5. **Uncertainty Handling**: How does the author handle sections with limited information?
6. **Section-Specific Conventions**: Any unique formatting or structural patterns per section

Here are representative examples from each section type:

{"".join(section_summaries)}

Return the style guide as a JSON object with 'general' and 'sections' keys."""

    try:
        response = llm_caller.generate(prompt, task_type="general")
        # Try to parse JSON from response
        json_match = re.search(r'\{[\s\S]*\}', response)
        if json_match:
            guide = json.loads(json_match.group())
        else:
            logger.warning("LLM did not return valid JSON, falling back to rule-based")
            return _distill_rule_based(all_examples)
    except Exception as e:
        logger.warning(f"LLM distillation failed: {e}, falling back to rule-based")
        return _distill_rule_based(all_examples)

    # Save
    os.makedirs(os.path.dirname(STYLE_GUIDE_PATH), exist_ok=True)
    with open(STYLE_GUIDE_PATH, 'w', encoding='utf-8') as f:
        json.dump(guide, f, indent=2, ensure_ascii=False)

    print(f"Style guide saved to: {STYLE_GUIDE_PATH}")
    return guide


def _analyze_section_patterns(section_name: str, examples: List[Dict]) -> Dict:
    """Analyze patterns in a set of section examples."""
    texts = [ex["text"] for ex in examples]

    # Find common sub-heading patterns
    sub_headings = set()
    for text in texts:
        # Look for lines that are likely sub-headings (short, often followed by content)
        for line in text.split('\n'):
            line = line.strip()
            if not line:
                continue
            # Sub-headings are typically short (< 80 chars) and may end with colon
            # or be in a lettered format (A. Something)
            if len(line) < 80 and (
                re.match(r'^[A-Z]\.\s+', line) or
                re.match(r'^\d+\.\s+', line) or
                line.endswith(':')
            ):
                sub_headings.add(line)

    # Estimate average paragraph length
    all_paras = []
    for text in texts:
        paras = [p.strip() for p in text.split('\n\n') if p.strip()]
        all_paras.extend(paras)
    avg_para_len = sum(len(p) for p in all_paras) / max(len(all_paras), 1)

    return {
        "typical_length_chars": sum(len(t) for t in texts) // len(texts),
        "avg_paragraph_length": int(avg_para_len),
        "common_sub_headings": sorted(list(sub_headings))[:15],
        "example_count": len(examples),
    }


def get_style_guide() -> Dict:
    """Load the style guide from disk."""
    if not os.path.exists(STYLE_GUIDE_PATH):
        return {}
    with open(STYLE_GUIDE_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)


def get_section_examples(section_name: str, max_examples: int = 3) -> List[str]:
    """Load example texts for a specific section type."""
    safe_name = section_name.lower().replace(" ", "_")
    filepath = os.path.join(STYLE_EXAMPLES_DIR, f"{safe_name}.json")

    if not os.path.exists(filepath):
        # Try normalized name
        normalized = SECTION_NORMALIZE.get(section_name.upper(), section_name.upper())
        safe_name = normalized.lower().replace(" ", "_")
        filepath = os.path.join(STYLE_EXAMPLES_DIR, f"{safe_name}.json")

    if not os.path.exists(filepath):
        return []

    with open(filepath, 'r', encoding='utf-8') as f:
        data = json.load(f)

    examples = data.get("examples", [])
    return [ex["text"] for ex in examples[:max_examples]]


def update_library(new_report_path: str):
    """Add a new report to the style library."""
    if not os.path.exists(new_report_path):
        raise FileNotFoundError(f"Report not found: {new_report_path}")

    sections = extract_sections_from_docx(new_report_path)
    filename = os.path.basename(new_report_path)

    updated_sections = []
    for section_name, text in sections.items():
        if not text or len(text) < 50:
            continue

        safe_name = section_name.lower().replace(" ", "_")
        filepath = os.path.join(STYLE_EXAMPLES_DIR, f"{safe_name}.json")

        if os.path.exists(filepath):
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
        else:
            data = {"total_examples_found": 0, "selected_count": 0, "examples": []}

        # Check if this source is already in the library
        existing_sources = {ex["source"] for ex in data["examples"]}
        if filename in existing_sources:
            continue

        # Add new example
        data["examples"].append({
            "source": filename,
            "text": text,
            "char_count": len(text)
        })
        data["total_examples_found"] += 1

        # Re-sort by length and trim to max
        data["examples"].sort(key=lambda x: x["char_count"], reverse=True)
        data["examples"] = data["examples"][:MAX_EXAMPLES_PER_SECTION]
        data["selected_count"] = len(data["examples"])

        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        updated_sections.append(section_name)

    if updated_sections:
        print(f"Updated sections from {filename}: {', '.join(updated_sections)}")
    else:
        print(f"No new sections to add from {filename}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python -m Scripts.report_generator.style_library build [--reports-dir PATH]")
        print("  python -m Scripts.report_generator.style_library update <report.docx>")
        sys.exit(1)

    command = sys.argv[1]

    if command == "build":
        reports_dir = None
        if "--reports-dir" in sys.argv:
            idx = sys.argv.index("--reports-dir")
            reports_dir = sys.argv[idx + 1]

        summary = build_style_library(reports_dir)
        print("\n--- Style Library Summary ---")
        print(f"Reports parsed: {summary['reports_parsed']}")
        for name, info in summary["sections"].items():
            print(f"  {name}: {info['total_found']} found, {info['examples_stored']} stored")

        # Distill style guide (rule-based for now, LLM later)
        print("\nDistilling style guide...")
        guide = distill_style_guide()
        print("Done!")

    elif command == "update":
        if len(sys.argv) < 3:
            print("Usage: python -m Scripts.report_generator.style_library update <report.docx>")
            sys.exit(1)
        update_library(sys.argv[2])

    else:
        print(f"Unknown command: {command}")
        sys.exit(1)
