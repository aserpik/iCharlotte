"""Prompt builders for Depo Prep. Each builder returns (prompt, text_payload)."""
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


def _clip(text: str, cap: int = _FREE_TEXT_CAP) -> str:
    if not text:
        return ""
    if len(text) <= cap:
        return text
    return text[:cap] + f"\n[...truncated at {cap} chars]"


def build_per_source_digest_prompt(
    *, deponent_name: str, deponent_role: str, source_text: str, source_filename: str,
) -> Tuple[str, str]:
    prompt = f"""You are an extraction agent for a deposition-prep tool.

The deponent we are preparing to depose is: {deponent_name} ({deponent_role}).

The following source document is named: {source_filename}.

Your job: read the document and produce a STRUCTURED JSON DIGEST of facts and quotes
relevant to deposing this witness. Output **JSON ONLY**, no commentary.

Schema (exact field names required):

{{
  "source_id": "{source_filename}",
  "source_kind": "medical_records | deposition_transcript | discovery_response | pleading | other",
  "deponent_statements": [
    {{ "text": "<verbatim quote from the witness>",
       "location": "<page/line citation if available>",
       "context": "<who was questioning, what segment>" }}
  ],
  "factual_anchors": [
    {{ "fact": "<short factual claim found in the doc, e.g. 'MRI 2024-09-12 showed 4mm protrusion'>",
       "location": "<page/Bates citation>",
       "topic_tags": ["<short free-form tags for clustering, e.g. 'injury', 'causation'>"] }}
  ],
  "inconsistencies": [
    {{ "claim_a": "...", "claim_a_source": "...",
       "claim_b": "...", "claim_b_source": "...",
       "topic_tags": ["credibility", "..."] }}
  ],
  "summary": "<2-3 sentence summary of what this source contributes to the depo prep>"
}}

Rules:
- Quote verbatim where possible. Do not paraphrase witness statements.
- Use citations the document itself contains (page numbers, Bates, page:line).
- If a list has no entries, return an empty list (never null).
- Output JSON only - no markdown fences, no preamble.
"""
    return prompt, source_text


def build_topic_clustering_prompt(
    *, deponent_name: str, deponent_role: str, style: str,
    free_text_notes: str, digests_summary_text: str,
) -> Tuple[str, str]:
    style_block = STYLE_DIRECTIVES.get(style, STYLE_DIRECTIVES["discovery"])
    notes_block = _clip(free_text_notes, _FREE_TEXT_CAP)
    prompt = f"""You are designing the topic structure for a deposition outline.

Deponent: {deponent_name} ({deponent_role}).

{style_block}

Lawyer's strategy notes:
{notes_block or '(none provided)'}

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
{{ "topics": [ {{...}}, {{...}} ] }}

Produce between 8 and 15 topics. Aim for clean, non-overlapping coverage.
"""
    return prompt, digests_summary_text


def build_per_topic_questions_prompt(
    *, deponent_name: str, deponent_role: str, style: str,
    topic_title: str, strategic_note: str, digest_excerpts_text: str,
    free_text_notes: str,
    include_strategic_note: bool, include_source_facts: bool,
    include_impeachment_hook: bool, include_objection_alts: bool,
) -> Tuple[str, str]:
    style_block = STYLE_DIRECTIVES.get(style, STYLE_DIRECTIVES["discovery"])
    notes_block = _clip(free_text_notes, _FREE_TEXT_CAP)

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

    prompt = f"""You are drafting deposition questions for one topic of an outline.

Deponent: {deponent_name} ({deponent_role}).

{style_block}

Topic: {topic_title}
Strategic note (what we're trying to establish): {strategic_note}

Lawyer's overall strategy notes:
{notes_block or '(none provided)'}

You will be given digest excerpts (verbatim quotes, factual anchors, inconsistencies)
relevant to this topic. Draft 5-10 questions that probe this topic, grounded in
those source facts.

Output JSON ONLY (no fences, no commentary):
{{
  "topic_id": "<echoed>",
  "questions": [
    {{
      "n": 1,
      "text": "<the question itself; never compound>",
{optional_fields_block}
    }},
    ...
  ]
}}

Rules:
- Never invent facts not present in the digest excerpts. Every factual claim in a
  question must be traceable to the excerpts.
- Each question is a single, clean inquiry - no compound questions.
- Number sequentially starting at 1.
- If you cannot generate meaningful questions for this topic (e.g., no relevant
  source facts), return an empty questions list.
"""
    return prompt, digest_excerpts_text


def build_dedup_prompt(*, topic_outputs_summary: str, digest_summary: str) -> Tuple[str, str]:
    prompt = """You are auditing a deposition outline for duplicates and coverage gaps.

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
    return prompt, f"=== TOPIC OUTPUTS ===\n{topic_outputs_summary}\n\n=== DIGEST SUMMARY ===\n{digest_summary}"


def build_polish_prompt(*, outline_text: str) -> Tuple[str, str]:
    prompt = """You are doing a final phrasing pass on a deposition outline.

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
    return prompt, outline_text
