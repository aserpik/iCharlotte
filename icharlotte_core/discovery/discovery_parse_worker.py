"""One-shot QThread that parses an incoming discovery document via LLM."""
from __future__ import annotations

from PySide6.QtCore import QThread, Signal

from icharlotte_core.discovery.response_parser import (
    ParsedDiscovery,
    build_parse_prompt,
    parse_llm_response,
)
from icharlotte_core.discovery._io import (
    normalize_and_filter_parsed_discovery,
    read_document_text,
)
from icharlotte_core.llm_config import call_llm


class DiscoveryParseWorker(QThread):
    """Reads the discovery file, runs the parse LLM, normalizes, emits once."""

    # parse_finished(success: bool, parsed_or_error_message: ParsedDiscovery | str)
    parse_finished = Signal(bool, object)

    def __init__(
        self,
        discovery_file: str,
        detected_type: str,
        parent=None,
    ):
        super().__init__(parent)
        self.discovery_file = discovery_file
        self.detected_type = detected_type

    def run(self) -> None:
        try:
            discovery_text = read_document_text(self.discovery_file)
            if not discovery_text.strip():
                self.parse_finished.emit(False, "Could not read text from the discovery file.")
                return

            raw = call_llm(
                build_parse_prompt(discovery_text),
                "",
                task_type="extraction",
                agent_id="agent_sum_disc",
            )
            if not raw:
                self.parse_finished.emit(False, "The parser did not return a response.")
                return

            parsed = parse_llm_response(raw)
            parsed = normalize_and_filter_parsed_discovery(
                parsed, self.detected_type, self.discovery_file,
            )
            self.parse_finished.emit(True, parsed)
        except Exception as exc:
            self.parse_finished.emit(False, str(exc))
