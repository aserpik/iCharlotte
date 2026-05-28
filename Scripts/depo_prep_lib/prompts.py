"""Prompt builders for Depo Prep. Each builder returns (prompt, text_payload).

Prompts are workbench-editable: the default body of each pass lives in a
module-level template with ``{placeholder}`` tokens. At call time the builder
loads the current version from the PromptManager store
(``Scripts/prompts/depo_prep/<pass>_current.txt``) — edited via the Prompt
Workbench — falling back to the hardcoded default, then substitutes the
placeholders. Templates are plain strings (NOT f-strings), so the literal JSON
braces in the schema examples need no escaping.
"""
from __future__ import annotations

from typing import Tuple


_FREE_TEXT_CAP = 2000


STYLE_DIRECTIVES = {
    "discovery": (
        "Style: DISCOVERY / FACT-GATHERING. Use open-ended questions designed to "
        "develop the witness's complete account. Prefer 'Tell me about...', "
        "'What happened next?', 'Walk me through...' phrasings. Do not lead. The goal "
        "is to uncover testimony, not to box the witness in."
    ),
    "lockdown": (
        "Style: LOCK-DOWN / LEADING. Use short, closed, leading questions designed "
        "to extract specific admissions for use at trial or MSJ. Prefer "
        "'Isn't it true that...', 'You agree that...', 'You did X, correct?' phrasings. "
        "Each question should produce a yes/no answer or a precise concession. "
        "Keep questions short - never compound."
    ),
    "expert": (
        "Style: EXPERT CHALLENGE (Daubert-style). Probe methodology, qualifications, "
        "scope of opinions, materials reviewed, and the reliability/general acceptance "
        "of the underlying methods. Look for ipse dixit reasoning, gaps in the analysis, "
        "and reliance on unreliable foundations. Tie every opinion to specific support."
    ),
    "friendly": (
        "Style: FRIENDLY (own-client prep). Use clear, organized questions that allow "
        "the witness to tell their story cleanly. Anticipate weaknesses and address "
        "them head-on with rehabilitative questions. Avoid jargon. Build credibility."
    ),
}


# =============================================================================
# Default prompt templates (workbench-editable). Plain strings — NOT f-strings.
# Placeholders are substituted by the builders below via str.replace().
# =============================================================================

_SOURCE_DIGEST_TEMPLATE = """You are an extraction agent for a deposition-prep tool.

The deponent we are preparing to depose is: {deponent_name} ({deponent_role}).

The following source document is named: {source_filename}.

Your job: read the document and produce a STRUCTURED JSON DIGEST of facts and quotes
relevant to deposing this witness. Output **JSON ONLY**, no commentary.

Schema (exact field names required):

{
  "source_id": "{source_filename}",
  "source_kind": "medical_records | deposition_transcript | discovery_response | pleading | other",
  "deponent_statements": [
    { "text": "<verbatim quote from the witness>",
      "location": "<page/line citation if available>",
      "context": "<who was questioning, what segment>" }
  ],
  "factual_anchors": [
    { "fact": "<short factual claim found in the doc, e.g. 'MRI 2024-09-12 showed 4mm protrusion'>",
      "location": "<page/Bates citation>",
      "topic_tags": ["<short free-form tags for clustering, e.g. 'injury', 'causation'>"] }
  ],
  "inconsistencies": [
    { "claim_a": "...", "claim_a_source": "...",
      "claim_b": "...", "claim_b_source": "...",
      "topic_tags": ["credibility", "..."] }
  ],
  "summary": "<2-3 sentence summary of what this source contributes to the depo prep>"
}

Rules:
- Quote verbatim where possible. Do not paraphrase witness statements.
- Use citations the document itself contains (page numbers, Bates, page:line).
- If a list has no entries, return an empty list (never null).
- Output JSON only - no markdown fences, no preamble.
"""


_TOPIC_CLUSTERING_TEMPLATE = """You are designing the topic structure for a deposition outline.

Deponent: {deponent_name} ({deponent_role}).

{style_directive}

Lawyer's strategy notes:
{strategy_notes}

You will be given a concatenated set of per-source digests (JSON) describing the
case material. Your job: cluster the facts/quotes/inconsistencies into 8-15 TOPICS
that organize the deposition outline.

Each topic must include:
- "id": stable short id ("t01", "t02", ...)
- "title": 3-8 word topic name in title case
- "strategic_note": 1-3 sentences naming what the lawyer is trying to ESTABLISH,
   UNDERMINE, or LOCK DOWN under this topic, anchored to the lawyer's strategy notes
   when relevant
- "relevant_digest_refs": list of strings in the form
   "<source_id>#<schema_field>[<index>]" pointing to the digest entries this topic
   draws on. Example: "med_records.pdf#factual_anchors[2]".
- "default_checked": true unless the topic is genuinely speculative

Output JSON ONLY (no fences, no commentary):
{ "topics": [ {...}, {...} ] }

Produce between 8 and 15 topics. Aim for clean, non-overlapping coverage.
"""


_TOPIC_QUESTIONS_TEMPLATE = """You are drafting deposition questions for one topic of an outline.

Deponent: {deponent_name} ({deponent_role}).

{style_directive}

Topic: {topic_title}
Strategic note (what we're trying to establish): {strategic_note}

Lawyer's overall strategy notes:
{strategy_notes}

You will be given digest excerpts (verbatim quotes, factual anchors, inconsistencies)
relevant to this topic. Draft 5-10 questions that probe this topic, grounded in
those source facts.

Output JSON ONLY (no fences, no commentary):
{
  "topic_id": "<echoed>",
  "questions": [
    {
      "n": 1,
      "text": "<the question itself; never compound>",
{optional_fields_block}
    },
    ...
  ]
}

Rules:
- Never invent facts not present in the digest excerpts. Every factual claim in a
  question must be traceable to the excerpts.
- Each question is a single, clean inquiry - no compound questions.
- Number sequentially starting at 1.
- If you cannot generate meaningful questions for this topic (e.g., no relevant
  source facts), return an empty questions list.
"""


_DEDUP_TEMPLATE = """You are auditing a deposition outline for duplicates and coverage gaps.

You will be given:
- A summary of every topic's questions (numbered as "<topic_id>.q<n>: <text>").
- A summary of the source-digest facts available.

Identify:
1. Duplicate questions across topics - same substantive ask, possibly different phrasing.
   For each pair, choose which to KEEP and which to DROP.
2. Coverage gaps - important facts from the digest that no question addresses.

Output JSON ONLY:
{
  "duplicates": [
    { "keep": "<topic_id>.q<n>", "drop": "<topic_id>.q<n>", "reason": "<short>" }
  ],
  "coverage_gaps": [ "<one line, ending with a topic suggestion>" ],
  "renumber_after_dedup": true
}

Be conservative. Only flag duplicates that are substantively the same.
"""


_POLISH_TEMPLATE = """You are doing a final phrasing pass on a deposition outline.

ABSOLUTE RULES:
- Do not add any new questions.
- Do not drop or remove any questions.
- Do not change any factual content in a question - only phrasing.
- Do not change strategic notes substantively.
- Do not renumber questions.

Allowed changes:
- Tighten redundant phrasing.
- Add brief topic-to-topic transitions in strategic notes only.
- Normalize question phrasing consistency within a topic.
- Fix obvious typos.

Return the polished outline in the SAME structure as input, JSON ONLY.
"""


# Workbench registration: pass_name -> default template. Imported by
# icharlotte_core.prompt_manager.seed_pipeline_prompts so these show up as
# editable prompts under the "depo_prep" agent in the Prompt Workbench.
DEPO_PREP_PROMPT_DEFAULTS = {
    "source_digest": _SOURCE_DIGEST_TEMPLATE,
    "topic_clustering": _TOPIC_CLUSTERING_TEMPLATE,
    "topic_questions": _TOPIC_QUESTIONS_TEMPLATE,
    "dedup": _DEDUP_TEMPLATE,
    "polish": _POLISH_TEMPLATE,
}

DEPO_PREP_PROMPT_DESCRIPTIONS = {
    "source_digest": "Phase 1: per-source structured digest extraction",
    "topic_clustering": "Phase 1: cluster per-source digests into 8-15 topics",
    "topic_questions": "Phase 2: per-topic question generation (the optional-field block is injected by code based on the per-topic checkboxes)",
    "dedup": "Phase 2: cross-topic dedup + coverage-gap check",
    "polish": "Phase 2: phrasing-only polish pass",
}


def _safe(value: str) -> str:
    """Strip brace characters from substituted values so a user-supplied string
    can't introduce a placeholder token that a later replacement would catch."""
    return (value or "").replace("{", "").replace("}", "")


def _load_template(pass_name: str) -> str:
    """Return the current workbench-edited template for ``pass_name``, falling
    back to the hardcoded default if the store is unavailable/unseeded."""
    default = DEPO_PREP_PROMPT_DEFAULTS[pass_name]
    try:
        from icharlotte_core.prompt_manager import get_prompt
        loaded = get_prompt("depo_prep", pass_name)
        return loaded if loaded else default
    except Exception:
        return default


def _render(pass_name: str, substitutions: dict) -> str:
    template = _load_template(pass_name)
    for key, value in substitutions.items():
        template = template.replace("{" + key + "}", _safe(value))
    return template


def _clip(text: str, cap: int = _FREE_TEXT_CAP) -> str:
    if not text:
        return ""
    if len(text) <= cap:
        return text
    return text[:cap] + f"\n[...truncated at {cap} chars]"


def build_per_source_digest_prompt(
    *, deponent_name: str, deponent_role: str, source_text: str, source_filename: str,
) -> Tuple[str, str]:
    prompt = _render("source_digest", {
        "deponent_name": deponent_name,
        "deponent_role": deponent_role,
        "source_filename": source_filename,
    })
    return prompt, source_text


def build_topic_clustering_prompt(
    *, deponent_name: str, deponent_role: str, style: str,
    free_text_notes: str, digests_summary_text: str,
) -> Tuple[str, str]:
    notes_block = _clip(free_text_notes, _FREE_TEXT_CAP) or "(none provided)"
    prompt = _render("topic_clustering", {
        "deponent_name": deponent_name,
        "deponent_role": deponent_role,
        "style_directive": STYLE_DIRECTIVES.get(style, STYLE_DIRECTIVES["discovery"]),
        "strategy_notes": notes_block,
    })
    return prompt, digests_summary_text


def build_per_topic_questions_prompt(
    *, deponent_name: str, deponent_role: str, style: str,
    topic_title: str, strategic_note: str, digest_excerpts_text: str,
    free_text_notes: str,
    include_strategic_note: bool, include_source_facts: bool,
    include_impeachment_hook: bool, include_objection_alts: bool,
) -> Tuple[str, str]:
    notes_block = _clip(free_text_notes, _FREE_TEXT_CAP) or "(none provided)"

    field_instructions = []
    if include_strategic_note:
        field_instructions.append(
            '  "purpose": "<one-sentence statement of what this question is trying to '
            'establish, lock in, undermine, or develop>",'
        )
    if include_source_facts:
        field_instructions.append(
            '  "source_facts": [ "<bullet pointing to the specific document/page/quote '
            'that justifies this question>", ... ],'
        )
    if include_impeachment_hook:
        field_instructions.append(
            '  "impeachment_hook": "<if the witness denies / equivocates, the exact '
            'prior statement or document to confront them with>",'
        )
    if include_objection_alts:
        field_instructions.append(
            '  "objection_alts": [ "<cleaner rephrasing if opposing counsel objects '
            'vague/compound/asked-and-answered>", ... ],'
        )

    if field_instructions:
        # Strip trailing comma from the last entry so the assembled JSON schema is valid.
        field_instructions[-1] = field_instructions[-1].rstrip(",")
        optional_fields_block = "\n".join(field_instructions)
    else:
        optional_fields_block = "  (only 'n' and 'text' fields)"

    prompt = _render("topic_questions", {
        "deponent_name": deponent_name,
        "deponent_role": deponent_role,
        "style_directive": STYLE_DIRECTIVES.get(style, STYLE_DIRECTIVES["discovery"]),
        "topic_title": topic_title,
        "strategic_note": strategic_note,
        "strategy_notes": notes_block,
        "optional_fields_block": optional_fields_block,
    })
    return prompt, digest_excerpts_text


def build_dedup_prompt(*, topic_outputs_summary: str, digest_summary: str) -> Tuple[str, str]:
    prompt = _load_template("dedup")
    return prompt, f"=== TOPIC OUTPUTS ===\n{topic_outputs_summary}\n\n=== DIGEST SUMMARY ===\n{digest_summary}"


def build_polish_prompt(*, outline_text: str) -> Tuple[str, str]:
    prompt = _load_template("polish")
    return prompt, outline_text
