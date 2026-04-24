"""
Refine Stage (Stage 2) - Per-section LLM synthesis of raw agent outputs into report prose.

Takes raw agent output (docket text, interrogatory responses, complaint extractions,
medical record summaries, etc.) and SYNTHESIZES it into polished report-ready prose
that matches the user's writing voice, using the style guide and example sections.

The LLM's job is to write report sections from scratch using the raw data as its source
of facts — not to copy-paste or reformat the raw input.

Usage:
    from Scripts.report_generator.refine import refine_sections
    refined = refine_sections(gathered_data)
"""

import logging
from typing import Dict, List, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

logger = logging.getLogger(__name__)

# Sections that bypass LLM refinement and use raw gathered content directly.
# These sections already contain well-structured prose from their source agents.
PASSTHROUGH_SECTIONS = {
    "FACTUAL BACKGROUND",
    "PROCEDURAL STATUS",
    "DISCOVERY",
    "INVESTIGATION",
}

# Placeholder text for sections without data (FSR reports)
PLACEHOLDER_TEMPLATES = {
    "FACTUAL BACKGROUND": (
        "A detailed factual background will be provided upon receipt and review "
        "of the relevant pleadings and documents."
    ),
    "PROCEDURAL STATUS": (
        "The procedural status will be updated as the case progresses through "
        "the court system."
    ),
    "INVESTIGATION": (
        "An investigation is currently underway. The results of the investigation "
        "will be reported in a subsequent status report."
    ),
    "DISCOVERY": (
        "Discovery responses have not yet been received. This section will be "
        "updated in a subsequent status report once responses are obtained."
    ),
    "EXPERTS": (
        "No expert opinions have been obtained at this time. This section will be "
        "updated in a subsequent status report once expert evaluations are completed."
    ),
    "MEDICAL RECORD REVIEW": (
        "Medical records have not yet been obtained and reviewed. This section "
        "will be updated in a subsequent status report once records are received "
        "and analyzed."
    ),
    "EVALUATION OF LIABILITY": (
        "A full evaluation of liability will be provided once the investigation "
        "and discovery process is further along."
    ),
    "EVALUATION OF EXPOSURE": (
        "An evaluation of exposure will be provided once sufficient information "
        "has been gathered through discovery and investigation."
    ),
    "SETTLEMENT STATUS": (
        "No settlement demand has been made at this time. We will advise once "
        "there is a development in settlement discussions."
    ),
    "FURTHER CASE HANDLING": (
        "We will continue to monitor the progress of this case and provide updates "
        "as developments occur. We recommend the following actions at this time:"
    ),
}


def refine_sections(gathered_data: Dict, llm_caller=None,
                    max_workers: int = 3) -> Dict[str, str]:
    """
    Refine all included sections from raw agent output into report-ready prose.

    Args:
        gathered_data: Output from gather_case_data()
        llm_caller: LLMCaller instance (if None, creates one)
        max_workers: Max parallel LLM calls for section refinement

    Returns:
        Dict mapping section name -> refined prose text
    """
    if llm_caller is None:
        llm_caller = _get_llm_caller()

    sections = gathered_data["sections"]
    included = gathered_data["included_sections"]
    report_type = gathered_data["report_type"]
    prior_sections = gathered_data.get("prior_sections", {})
    metadata = gathered_data.get("metadata", {})

    # Load style guide and examples
    style_guide = _load_style_guide()
    section_examples = _load_section_examples()

    refined = {}

    # Build refinement tasks
    tasks = []
    for section_name in included:
        raw_content = sections.get(section_name)
        if not raw_content:
            refined[section_name] = PLACEHOLDER_TEMPLATES.get(
                section_name, "This section will be updated in a subsequent report."
            )
            continue

        # Passthrough sections use raw content directly (no LLM rewriting)
        if section_name in PASSTHROUGH_SECTIONS:
            refined[section_name] = raw_content
            logger.info(f"  {section_name}: passthrough ({len(raw_content)} chars)")
            continue

        prior = prior_sections.get(section_name, "")
        examples = section_examples.get(section_name, [])

        tasks.append((section_name, raw_content, report_type, prior, examples,
                       style_guide, metadata))

    if not tasks:
        return refined

    # Refine sections (parallel for 3+, sequential otherwise)
    if len(tasks) <= 2 or llm_caller is None:
        for args in tasks:
            name, result = _refine_single_section(*args, llm_caller=llm_caller)
            refined[name] = result
    else:
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for args in tasks:
                future = executor.submit(
                    _refine_single_section, *args, llm_caller=llm_caller
                )
                futures[future] = args[0]

            for future in as_completed(futures):
                section_name = futures[future]
                try:
                    name, result = future.result()
                    refined[name] = result
                except Exception as e:
                    logger.error(f"Error refining {section_name}: {e}")
                    refined[section_name] = sections.get(section_name, "")

    # Log content length changes for monitoring (no fallback — LLM synthesis is expected
    # to produce different lengths than raw input since it's synthesizing, not copying)
    for section_name in refined:
        raw = sections.get(section_name, "")
        ref = refined[section_name]
        if raw:
            ratio = len(ref) / len(raw) if len(raw) > 0 else 0
            logger.info(
                f"  {section_name}: {len(raw)} raw -> {len(ref)} refined ({ratio*100:.0f}%)"
            )

    logger.info(f"Refined {len(refined)} sections")
    return refined


def _refine_single_section(
    section_name: str,
    raw_content: str,
    report_type: str,
    prior_section: str,
    examples: List[str],
    style_guide: Dict,
    metadata: Dict,
    llm_caller=None,
) -> tuple:
    """Refine a single section using LLM. Returns (section_name, refined_text)."""
    logger.info(f"Refining: {section_name} ({len(raw_content)} chars)")

    prompt = _build_refinement_prompt(
        section_name, raw_content, report_type, prior_section,
        examples, style_guide, metadata
    )

    if llm_caller is None:
        logger.warning(f"No LLM caller, returning raw content for {section_name}")
        return (section_name, raw_content)

    try:
        result = llm_caller.call(
            prompt=prompt,
            text=raw_content,
            task_type="summary",
            agent_id="agent_report_refine",
        )
        if result:
            return (section_name, result.strip())
        else:
            logger.warning(f"LLM returned empty for {section_name}, using raw content")
            return (section_name, raw_content)
    except Exception as e:
        logger.error(f"LLM refinement failed for {section_name}: {e}")
        return (section_name, raw_content)


def _build_refinement_prompt(
    section_name: str,
    raw_content: str,
    report_type: str,
    prior_section: str,
    examples: List[str],
    style_guide: Dict,
    metadata: Dict,
) -> str:
    """Build the LLM prompt for synthesizing a section from raw data."""

    general_style = style_guide.get("general", {})
    hedging = general_style.get("hedging_phrases", [
        "We believe", "It is anticipated that", "It appears that",
        "Based on our review", "In our assessment",
    ])
    section_stats = style_guide.get("sections", {}).get(section_name, {})
    typical_length = section_stats.get("typical_length_chars", 3000)

    # Build example block — generous length so the LLM truly learns the voice
    example_block = ""
    if examples:
        for i, ex in enumerate(examples[:2], 1):
            truncated = ex[:10000] if len(ex) > 10000 else ex
            example_block += (
                f"\n=== EXAMPLE {i} — STUDY THIS VOICE AND STRUCTURE CAREFULLY ===\n"
                f"{truncated}\n"
                f"=== END EXAMPLE {i} ===\n"
            )

    # Report type context
    if report_type == "FSR":
        type_instruction = (
            "This is an Initial Litigation Plan / Full Status Report (ILP/FSR). "
            "Be comprehensive and thorough."
        )
    else:
        type_instruction = (
            "This is a Status Report update. Focus on new developments since the "
            "prior report, but still provide full context."
        )

    # Prior report context
    prior_block = ""
    if prior_section:
        prior_truncated = prior_section[:15000] if len(prior_section) > 15000 else prior_section
        prior_block = (
            "\nPRIOR REPORT VERSION (match this voice and structure):\n"
            f"{prior_truncated}\n"
        )

    # Section-specific instructions
    section_instructions = _get_section_specific_instructions(section_name)

    prompt = f"""You are writing the "{section_name}" section of a defense litigation report for an insurance carrier. You will be given raw extracted data (from court dockets, complaint summaries, interrogatory responses, medical records, etc.) that you must SYNTHESIZE into polished report prose.

CRITICAL: You are NOT copy-pasting or reformatting the raw input. You are WRITING a professional report section from scratch, using the raw data only as your source of facts. Your output must read EXACTLY like the examples below — same voice, same sentence structures, same level of professional polish.

REPORT TYPE: {type_instruction}

=== VOICE AND STYLE (match the examples precisely) ===
1. Write from defense counsel's perspective, addressing the report to the insurance carrier.
2. Use these hedging phrases naturally throughout: {', '.join(hedging)}.
3. Formal, professional tone. Refer to parties as "Plaintiff" and "Defendant" (not first names).
4. Use phrases like "We will argue...", "Plaintiff will contend...", "It is anticipated that...", "Here, ..." when analyzing legal issues.
5. Write in complete, polished paragraphs. No bullet points or numbered lists in body text.

=== CONTENT ===
1. Extract ALL significant facts, dates, names, dollar amounts, and legal points from the raw data.
2. Organize them into coherent narrative prose following the structure shown in the examples.
3. Do NOT include raw data formatting (no "FROG No. 1.1:", no "SROG No. 4:", no interrogatory numbers). Weave responses into flowing narrative paragraphs organized by topic.
4. Do NOT include source file markers, file references, or "--- Source:" separators.
5. Write thoroughly. Typical length for this section: ~{typical_length} characters.

=== FORMAT ===
1. Sub-headings go on their own line, separated by blank lines from body text.
2. Separate body paragraphs with blank lines.
3. Do NOT write the main section heading (don't write "{section_name}" at the top).
4. No markdown (no **, no ##, no *). Plain prose only. Sub-headings are plain text lines.

{section_instructions}
{example_block}
{prior_block}
=== RAW DATA TO SYNTHESIZE INTO REPORT PROSE ===
"""

    return prompt.strip()


def _get_section_specific_instructions(section_name: str) -> str:
    """Section-specific instructions for structure and sub-heading conventions."""
    instructions = {
        "FACTUAL BACKGROUND": (
            "=== SECTION STRUCTURE ===\n"
            "- Dense narrative prose, typically 2-5 paragraphs. NO sub-headings.\n"
            "- Open with the type of matter and key parties.\n"
            "- Include date of incident, location, nature of claims, causes of action.\n"
            "- End with a statement of the damages sought.\n"
            "- Use past tense throughout."
        ),
        "PROCEDURAL STATUS": (
            "=== SECTION STRUCTURE ===\n"
            "- Use these sub-headings in order (include only those applicable):\n"
            "  Government Tort Claim, Complaint (or Pleadings), Responsive Pleading,\n"
            "  Venue, Representation, [any Motions sub-headings], Future Hearings.\n"
            "- Under each sub-heading, write narrative prose with ALL dates.\n"
            "- For Complaint: describe each cause of action and factual basis.\n"
            "- For Venue: name the court, department, judge, and why venue is proper.\n"
            "- For Representation: identify counsel for each party.\n"
            "- End with Future Hearings. Close with: 'We confirm that we have\n"
            "  calendared [this date/these dates] and will appear at the same to\n"
            "  protect [Defendant]'s interests.'"
        ),
        "INVESTIGATION": (
            "=== SECTION STRUCTURE ===\n"
            "- Open with: 'Below is a summary of the investigation we have completed to date:'\n"
            "- Create a SEPARATE sub-heading for EACH source document provided in the raw data.\n"
            "  The raw data contains '--- Source: [filename] ---' markers between documents.\n"
            "  Each document MUST get its own sub-heading and summary.\n"
            "- Name each sub-heading after the document type or description, e.g.:\n"
            "  'Text Messages Between Defendant and Contractor',\n"
            "  'Text Messages Between Defendant and Plaintiff',\n"
            "  'Property Inspection Report (January 8, 2019)',\n"
            "  'Commercial Lease Agreement', 'Contractor Invoices and Estimates',\n"
            "  'Email Correspondence — Roof Repairs', etc.\n"
            "- Do NOT group or blend information from different documents together.\n"
            "  Each sub-heading should summarize only what is contained in that\n"
            "  single source document. The reader must be able to identify which\n"
            "  document each piece of information came from.\n"
            "- Under each sub-heading, provide a detailed narrative summary of\n"
            "  that document's contents, including ALL key facts, dates, names,\n"
            "  and dollar amounts.\n"
            "- If a document contains related litigation materials (lawsuits,\n"
            "  court filings), give it its own sub-heading as well."
        ),
        "DISCOVERY": (
            "=== SECTION STRUCTURE ===\n"
            "- 'Written Discovery' is a TOP-LEVEL sub-heading.\n"
            "- Individual party responses are NESTED sub-headings under Written Discovery\n"
            "  (e.g., 'Plaintiff Mimi Dudash's Responses' is nested under 'Written Discovery').\n"
            "- Do NOT prefix sub-headings with letters or numbers — the renderer adds those.\n"
            "- For interrogatory responses: DO NOT list FROG/SROG numbers.\n"
            "  Instead, synthesize responses into narrative paragraphs organized\n"
            "  by topic (background, incident, injuries, witnesses, damages).\n"
            "- Use intro phrases like 'Below is a summary of Plaintiff's\n"
            "  responses to written discovery:'\n"
            "- Include ALL dollar amounts, dates, and factual claims.\n"
            "- ONLY include discovery responses and deposition testimony summaries.\n"
            "- DO NOT include 'Related Litigation Documents' — those belong in\n"
            "  the Investigation section.\n"
            "- End with Outstanding Discovery sub-heading if applicable."
        ),
        "EXPERTS": (
            "=== SECTION STRUCTURE ===\n"
            "- Use expert name or examination type as sub-headings.\n"
            "- Under each sub-heading, summarize the expert's findings,\n"
            "  opinions, and conclusions.\n"
            "- Include dates of examination, expert credentials if available.\n"
            "- Note any preliminary opinions and their basis.\n"
            "- If IME reports, include diagnoses, causation opinions, and\n"
            "  future treatment recommendations."
        ),
        "MEDICAL RECORD REVIEW": (
            "=== SECTION STRUCTURE ===\n"
            "- Use treatment provider names as sub-headings.\n"
            "- Under each provider, chronological narrative of visits with dates.\n"
            "- Include diagnoses, treatments, prognosis, and billing amounts.\n"
            "- End with total medical specials if available."
        ),
        "EVALUATION OF LIABILITY": (
            "=== SECTION STRUCTURE ===\n"
            "- Open with a brief intro identifying the causes of action.\n"
            "- Each CAUSE OF ACTION is a top-level sub-heading\n"
            "  (e.g., 'Continuing Nuisance', 'Continuing Trespass to Land', 'Negligence').\n"
            "- Each ELEMENT or defense argument under a cause of action is a NESTED\n"
            "  sub-heading (e.g., 'Causation', 'Consent', 'Ownership and Creation of Condition',\n"
            "  'Intent and Statute of Limitations Defense').\n"
            "- Do NOT prefix sub-headings with letters or numbers — the renderer adds those.\n"
            "- EVERY element sub-heading MUST be on its own line, separated by blank lines.\n"
            "- Under each element, analyze using case-specific facts.\n"
            "- Present Plaintiff's likely arguments, then defense arguments.\n"
            "- Cite legal authority (CACI instructions, case law) as shown in examples.\n"
            "- End with 'Summary' sub-heading providing overall liability assessment."
        ),
        "EVALUATION OF EXPOSURE": (
            "=== SECTION STRUCTURE ===\n"
            "- Use sub-headings for damage categories: Injuries, Medical Damages,\n"
            "  General Damages, Lost Wages, etc.\n"
            "- Reference specific findings from Medical Record Review and Liability.\n"
            "- Provide exposure ranges where possible.\n"
            "- Account for comparative fault and defense strengths."
        ),
        "SETTLEMENT STATUS": (
            "=== SECTION STRUCTURE ===\n"
            "- Brief narrative, typically 1-3 paragraphs.\n"
            "- Note demands, offers, mediation activity.\n"
            "- Provide settlement recommendation."
        ),
        "FURTHER CASE HANDLING": (
            "=== SECTION STRUCTURE ===\n"
            "- Use a numbered list format for recommended next steps.\n"
            "- Each item should be a specific, actionable recommendation.\n"
            "- Reference specific outstanding items from earlier sections\n"
            "  (e.g., 'propound supplemental discovery', 'depose the Plaintiff',\n"
            "  'retain an expert', 'file a motion for summary judgment').\n"
            "- Note upcoming deadlines and hearings with specific dates.\n"
            "- Include budgetary considerations where relevant.\n"
            "- End with a forward-looking recommendation or strategic note.\n"
            "- Typical items include: obtain/review discovery responses,\n"
            "  take/defend depositions, retain experts, file dispositive motions,\n"
            "  attend mediation, prepare for trial."
        ),
    }
    return instructions.get(section_name, "")


def _load_style_guide() -> Dict:
    """Load the distilled style guide."""
    try:
        from Scripts.report_generator.style_library import get_style_guide
        guide = get_style_guide()
        if guide:
            logger.info(f"Loaded style guide with {len(guide.get('sections', {}))} sections")
        return guide
    except Exception as e:
        logger.warning(f"Could not load style guide: {e}")
        return {}


def _load_section_examples() -> Dict[str, List[str]]:
    """Load example sections for each section type."""
    examples = {}
    try:
        from Scripts.report_generator.style_library import get_section_examples, REPORT_SECTIONS
        for section in REPORT_SECTIONS:
            exs = get_section_examples(section, max_examples=2)
            if exs:
                examples[section] = exs
                logger.info(f"Loaded {len(exs)} examples for {section}")
    except Exception as e:
        logger.warning(f"Could not load section examples: {e}")
    return examples


def _get_llm_caller():
    """Create an LLMCaller instance."""
    try:
        from icharlotte_core.llm_config import LLMCaller
        return LLMCaller()
    except ImportError:
        logger.warning("LLMCaller not available")
        return None
