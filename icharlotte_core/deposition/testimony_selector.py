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

SYSTEM_PROMPT = """You are a legal assistant specializing in deposition transcript analysis.

You will be given a list of question-and-answer exchanges from a deposition transcript, each with a unique ID number. You will also receive a user request describing what testimony they want extracted.

Your task: Identify the exchanges that are DIRECTLY relevant to the user's request. Be precise — only include testimony that substantively addresses the requested topic. Do NOT include exchanges that merely happen to be nearby or that discuss unrelated subjects.

IMPORTANT RULES:
1. Return ONLY a JSON array of the relevant exchange ID numbers.
2. Do NOT reproduce or modify any testimony text.
3. Do NOT explain your reasoning or add commentary.
4. Only include exchanges where the question or answer directly discusses the requested topic.
5. Do NOT include testimony about unrelated topics (e.g., if the user asks about the accident, do not include testimony about employment history, salary, education, or other background unless it directly relates to the accident).
6. Include brief transitional exchanges (e.g., "Let me ask you about X") only if they introduce a directly relevant block of testimony.

Example response format:
[1, 2, 3, 7, 12, 15, 16, 17, 23]

Return ONLY the JSON array, nothing else."""


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

    def select(self, index: TranscriptIndex, prompt: str, on_chunk_done=None) -> ExtractionResult:
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
            selected = self._select_chunk(chunk, prompt, index.deponent.full_name)
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
        self, exchanges: List[QAExchange], prompt: str, deponent_name: str
    ) -> List[int]:
        """Run LLM selection on a single chunk of exchanges."""
        # Format exchanges for LLM
        formatted = self._format_exchanges(exchanges)

        # Build prompt (instructions) and text (exchanges) separately
        # so LLMCaller formats them correctly for each provider
        instructions = (
            f"{SYSTEM_PROMPT}\n\n"
            f"DEPONENT: {deponent_name}\n\n"
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

        # Parse response — expect a JSON array of IDs
        return self._parse_response(response, exchanges)

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
