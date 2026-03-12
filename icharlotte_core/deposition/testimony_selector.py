"""
Testimony Selector — LLM-based relevance selection.

Sends the full Q/A index to an LLM and receives back the IDs of exchanges
relevant to the user's prompt. The LLM never produces the output text;
it only selects which exchanges are relevant.
"""

import json
import re
import logging
import math
from typing import List, Optional

from .models import TranscriptIndex, QAExchange, ExtractionResult

logger = logging.getLogger(__name__)

# Maximum exchanges per LLM chunk (to stay within context limits)
MAX_EXCHANGES_PER_CHUNK = 500

SYSTEM_PROMPT_BASE = """You are a legal assistant specializing in deposition transcript analysis.

You will be given a list of question-and-answer exchanges from a deposition transcript, each with a unique ID number. You will also receive a user request describing what testimony they want extracted.

{relevance_guidance}

IMPORTANT RULES:
1. Return ONLY a JSON array of the relevant exchange ID numbers.
2. Do NOT reproduce or modify any testimony text.
3. Do NOT explain your reasoning or add commentary.
4. Include brief transitional exchanges (e.g., "Let me ask you about X") only if they introduce a relevant block of testimony.

Example response format:
[1, 2, 3, 7, 12, 15, 16, 17, 23]

Return ONLY the JSON array, nothing else."""

RELEVANCE_GUIDANCE = {
    1: (
        "Your task: Be MAXIMALLY inclusive. Include ANY exchange where the requested topic "
        "is mentioned, referenced, or even implied — no matter how briefly or tangentially. "
        "Also include exchanges that discuss related body parts, activities, limitations, "
        "treatments, or life impacts that COULD be connected to the topic, even indirectly. "
        "Include surrounding context exchanges that help tell the full story. "
        "When in doubt, ALWAYS include the exchange. Err heavily on the side of inclusion. "
        "It is far better to include too much than to miss something potentially relevant."
    ),
    2: (
        "Your task: Be broadly inclusive. Include testimony that discusses the topic or "
        "provides useful context for understanding it. Include exchanges where the topic is "
        "mentioned even briefly as part of a longer answer about other subjects. Include "
        "testimony about related impacts, limitations, and life changes that may be connected "
        "to the topic. When in doubt, include the exchange."
    ),
    3: (
        "Your task: Include testimony that clearly relates to the user's request. "
        "Some surrounding context is fine, but skip testimony that only tangentially "
        "touches the subject. If the topic is mentioned as a minor aside within testimony "
        "about a different subject, you may skip it."
    ),
    4: (
        "Your task: Include only testimony that directly addresses the user's request. "
        "Skip background and tangential mentions unless they are clearly important to "
        "understanding the topic. Do NOT include testimony about unrelated subjects even "
        "if the topic gets a brief mention. Focus on exchanges where the topic is a "
        "primary subject of the question or answer."
    ),
    5: (
        "Your task: Be EXTREMELY selective. Include ONLY exchanges where the question or "
        "answer is SPECIFICALLY and PRIMARILY about the exact topic requested. Exclude "
        "exchanges where the topic is merely mentioned in passing, listed among other "
        "injuries/topics, or discussed as secondary context. If an exchange covers multiple "
        "subjects and the requested topic is not the main focus, EXCLUDE it. "
        "The bar for inclusion should be very high — when in doubt, EXCLUDE."
    ),
}

# Default level matches roughly the old behavior
DEFAULT_RELEVANCE_LEVEL = 4


def _build_system_prompt(relevance_level: int = DEFAULT_RELEVANCE_LEVEL) -> str:
    """Build the system prompt with the appropriate relevance guidance."""
    level = max(1, min(5, relevance_level))
    return SYSTEM_PROMPT_BASE.format(relevance_guidance=RELEVANCE_GUIDANCE[level])


class TestimonySelector:
    """
    Selects relevant Q/A exchanges from a transcript index using an LLM.

    Usage:
        selector = TestimonySelector()
        result = selector.select(index, "testimony about back pain before the accident")
    """

    def __init__(self, caller=None):
        """
        Initialize the selector.

        Args:
            caller: LLMCaller instance. If None, creates one.
        """
        if caller is None:
            from icharlotte_core.llm_config import LLMCaller
            caller = LLMCaller()
        self.caller = caller

    def select(self, index: TranscriptIndex, prompt: str, context: str = "",
               on_chunk_done=None, relevance_level: int = DEFAULT_RELEVANCE_LEVEL) -> ExtractionResult:
        """
        Select relevant exchanges from the index based on the user's prompt.

        Args:
            index: Parsed transcript index with all Q/A exchanges.
            prompt: User's extraction prompt (topic, category, question, etc.)
            on_chunk_done: Optional callback(current_chunk, total_chunks) for progress.

        Returns:
            ExtractionResult with selected IDs grouped into consecutive runs.
        """
        if not index.exchanges:
            logger.warning("No exchanges in index — nothing to select")
            return ExtractionResult(prompt=prompt)

        # Chunk if needed
        chunks = self._chunk_exchanges(index.exchanges)
        all_selected_ids = []

        for chunk_num, chunk in enumerate(chunks, 1):
            logger.info(f"Processing chunk {chunk_num}/{len(chunks)} ({len(chunk)} exchanges)")
            selected = self._select_chunk(chunk, prompt, index.deponent.full_name, context, relevance_level)
            all_selected_ids.extend(selected)
            logger.info(f"Chunk {chunk_num}: selected {len(selected)} exchanges")
            if on_chunk_done:
                on_chunk_done(chunk_num, len(chunks))

        # Build result with consecutive grouping
        result = ExtractionResult(prompt=prompt)
        result.group_consecutive(all_selected_ids)

        logger.info(
            f"Total selected: {len(result.selected_ids)} exchanges "
            f"in {len(result.groups)} consecutive groups"
        )
        return result

    def _chunk_exchanges(self, exchanges: List[QAExchange]) -> List[List[QAExchange]]:
        """Split exchanges into chunks if they exceed context limits."""
        if len(exchanges) <= MAX_EXCHANGES_PER_CHUNK:
            return [exchanges]

        num_chunks = math.ceil(len(exchanges) / MAX_EXCHANGES_PER_CHUNK)
        chunk_size = math.ceil(len(exchanges) / num_chunks)
        return [exchanges[i:i + chunk_size] for i in range(0, len(exchanges), chunk_size)]

    def _select_chunk(
        self, exchanges: List[QAExchange], prompt: str, deponent_name: str,
        context: str = "", relevance_level: int = DEFAULT_RELEVANCE_LEVEL
    ) -> List[int]:
        """Run LLM selection on a single chunk of exchanges."""
        # Format exchanges for LLM
        formatted = self._format_exchanges(exchanges)

        # Build prompt (instructions) and text (exchanges) separately
        # so LLMCaller formats them correctly for each provider
        system_prompt = _build_system_prompt(relevance_level)
        context_block = f"\n\nCASE CONTEXT:\n{context}" if context else ""
        instructions = (
            f"{system_prompt}\n\n"
            f"DEPONENT: {deponent_name}"
            f"{context_block}\n\n"
            f"USER'S REQUEST:\n{prompt}"
        )

        # Call LLM — exchanges go as "text" (document content)
        response = self.caller.call(
            prompt=instructions,
            text=formatted,
            agent_id="agent_depo_extract",
            task_type="extraction",
        )

        if not response:
            logger.error("LLM returned no response for testimony selection")
            return []

        logger.info(f"LLM response (first 500 chars): {response[:500]}")

        # Parse response — expect a JSON array of IDs
        parsed = self._parse_response(response, exchanges)
        logger.info(f"Parsed {len(parsed)} IDs from response")
        return parsed

    def _format_exchanges(self, exchanges: List[QAExchange]) -> str:
        """Format exchanges for LLM input."""
        lines = []
        for ex in exchanges:
            lines.append(
                f"[ID: {ex.id}] Page {ex.page_start}, Lines {ex.line_start}-"
                f"{ex.page_end}:{ex.line_end}"
            )
            lines.append(f"Q. {ex.question}")
            lines.append(f"A. {ex.answer}")
            lines.append("")
        return "\n".join(lines)

    def _parse_response(self, response: str, exchanges: List[QAExchange]) -> List[int]:
        """Parse LLM response to extract selected IDs."""
        # Try to find a JSON array in the response
        # The LLM might include extra text despite instructions
        response = response.strip()

        # Try direct JSON parse first
        try:
            ids = json.loads(response)
            if isinstance(ids, list):
                return self._validate_ids(ids, exchanges)
        except json.JSONDecodeError:
            pass

        # Try to extract JSON array from within text
        match = re.search(r'\[[\d,\s]+\]', response)
        if match:
            try:
                ids = json.loads(match.group())
                if isinstance(ids, list):
                    return self._validate_ids(ids, exchanges)
            except json.JSONDecodeError:
                pass

        # Fallback: extract all numbers from response
        logger.warning("Could not parse JSON array from LLM response, extracting numbers")
        numbers = re.findall(r'\b(\d+)\b', response)
        ids = [int(n) for n in numbers]
        return self._validate_ids(ids, exchanges)

    def _validate_ids(self, ids: List[int], exchanges: List[QAExchange]) -> List[int]:
        """Validate that IDs exist in the exchange list."""
        valid_ids = {ex.id for ex in exchanges}
        validated = [int(i) for i in ids if int(i) in valid_ids]
        invalid = [i for i in ids if int(i) not in valid_ids]
        if invalid:
            logger.warning(f"LLM returned {len(invalid)} invalid IDs: {invalid[:10]}...")
        return validated
