"""
Main orchestrator for the discovery generation pipeline.

Coordinates template loading, variable substitution, set tracking,
and LLM prompt building to produce DiscoverySet objects ready for
assembly into Word documents.
"""
import logging
import os
from typing import Dict, List, Optional

from .models import (
    CustomStyle,
    DiscoveryMode,
    DiscoveryRequest,
    DiscoverySet,
    DiscoveryType,
    Party,
    SetTrackerResult,
    number_to_word,
)
from .templates import (
    TemplateLoader,
    extract_requests_from_text,
    substitute_variables,
    _STANDARD_TYPE_FOLDERS,
    _TEMPLATE_FILENAMES,
)
from .set_tracker import SetTracker

logger = logging.getLogger(__name__)


class DiscoveryEngine:
    """Orchestrate the discovery generation pipeline.

    Parameters
    ----------
    discovery_dir : str
        Path to the root discovery directory containing template folders.
    """

    def __init__(self, discovery_dir: str):
        self.discovery_dir = discovery_dir
        self.loader = TemplateLoader(discovery_dir)

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def generate_standard(
        self,
        standard_type: str,
        discovery_types: List[DiscoveryType],
        directed_to: Party,
        propounding_party: Party,
        variables: Optional[Dict[str, str]] = None,
    ) -> List[DiscoverySet]:
        """Generate standard discovery sets from templates.

        For each discovery type: load requests from the template, load
        definitions and instructions, apply variable substitution if
        a variables dict is provided.

        Parameters
        ----------
        standard_type : str
            Key for the standard template folder (e.g. ``"negligence"``).
        discovery_types : list of DiscoveryType
            Types of discovery to generate (SI, RPD, RFA).
        directed_to : Party
            The party the discovery is directed to.
        propounding_party : Party
            The propounding party.
        variables : dict or None
            Variable substitution map (e.g. party names).

        Returns
        -------
        list of DiscoverySet
        """
        results: List[DiscoverySet] = []

        # Load definitions once (shared across types)
        definitions_block = self._load_definitions(variables)

        for disc_type in discovery_types:
            # Load requests from the standard template
            requests = self.loader.load_standard_requests(standard_type, disc_type)

            # Load instructions from the template file
            template_path = self._get_template_path(standard_type, disc_type)
            instructions_block = ""
            if template_path and os.path.isfile(template_path):
                try:
                    instructions_block = self.loader.load_instructions(
                        disc_type, template_path
                    )
                except Exception:
                    logger.warning(
                        "Could not load instructions for %s from %s",
                        disc_type.abbreviation, template_path,
                    )

            # Apply variable substitution to instructions
            if variables and instructions_block:
                instructions_block = substitute_variables(instructions_block, variables)

            ds = DiscoverySet(
                discovery_type=disc_type,
                set_number=1,
                directed_to=directed_to,
                propounding_party=propounding_party,
                requests=requests,
                definitions_block=definitions_block,
                instructions_block=instructions_block,
                previous_count=0,
            )
            results.append(ds)

        return results

    def generate_custom(
        self,
        custom_style: CustomStyle,
        standard_type: str,
        discovery_types: List[DiscoveryType],
        directed_to: Party,
        propounding_party: Party,
        user_instructions: str,
        context_text: str = "",
        variables: Optional[Dict[str, str]] = None,
    ) -> List[DiscoverySet]:
        """Generate custom discovery sets.

        For ``STANDARD_PLUS_CUSTOM``: loads standard requests first; the
        LLM-generated custom requests will be numbered starting after them.

        For ``CUSTOM_ONLY`` and ``MODIFIED_STANDARD``: returns DiscoverySets
        with empty requests (LLM fills them later).

        Parameters
        ----------
        custom_style : CustomStyle
            The custom generation style.
        standard_type : str
            Key for the standard template folder.
        discovery_types : list of DiscoveryType
            Types of discovery to generate.
        directed_to : Party
            The party the discovery is directed to.
        propounding_party : Party
            The propounding party.
        user_instructions : str
            User's instructions for custom generation.
        context_text : str
            Additional context documents text.
        variables : dict or None
            Variable substitution map.

        Returns
        -------
        list of DiscoverySet
        """
        results: List[DiscoverySet] = []

        # Load definitions once (shared across types)
        definitions_block = self._load_definitions(variables)

        for disc_type in discovery_types:
            # Load instructions from the template file
            template_path = self._get_template_path(standard_type, disc_type)
            instructions_block = ""
            if template_path and os.path.isfile(template_path):
                try:
                    instructions_block = self.loader.load_instructions(
                        disc_type, template_path
                    )
                except Exception:
                    logger.warning(
                        "Could not load instructions for %s", disc_type.abbreviation
                    )

            if variables and instructions_block:
                instructions_block = substitute_variables(instructions_block, variables)

            if custom_style == CustomStyle.STANDARD_PLUS_CUSTOM:
                # Load standard requests first
                try:
                    requests = self.loader.load_standard_requests(
                        standard_type, disc_type
                    )
                except Exception:
                    logger.warning(
                        "Could not load standard requests for %s",
                        disc_type.abbreviation,
                    )
                    requests = []
            else:
                # CUSTOM_ONLY and MODIFIED_STANDARD: empty requests (LLM fills)
                requests = []

            ds = DiscoverySet(
                discovery_type=disc_type,
                set_number=1,
                directed_to=directed_to,
                propounding_party=propounding_party,
                requests=requests,
                definitions_block=definitions_block,
                instructions_block=instructions_block,
                previous_count=0,
            )
            results.append(ds)

        return results

    def prepare_additional(
        self,
        case_path: str,
        discovery_types: List[DiscoveryType],
        directed_to: Party,
        propounding_party: Party,
    ) -> List[DiscoverySet]:
        """Prepare additional discovery set shells using SetTracker.

        Uses SetTracker to resolve the previous set (next set number,
        last request number, definitions) for each discovery type.
        Returns DiscoverySet shells with empty requests (LLM fills them
        later).

        Parameters
        ----------
        case_path : str
            Root path of the case folder.
        discovery_types : list of DiscoveryType
            Types of discovery to generate.
        directed_to : Party
            The party the discovery is directed to.
        propounding_party : Party
            The propounding party.

        Returns
        -------
        list of DiscoverySet
        """
        tracker = SetTracker(case_path)
        results: List[DiscoverySet] = []

        for disc_type in discovery_types:
            tracker_result = tracker.resolve_previous_set(
                disc_type, directed_to, self.discovery_dir
            )

            ds = DiscoverySet(
                discovery_type=disc_type,
                set_number=tracker_result.next_set_number,
                directed_to=directed_to,
                propounding_party=propounding_party,
                requests=[],
                definitions_block=tracker_result.previous_definitions,
                instructions_block=tracker_result.previous_instructions,
                previous_count=tracker_result.last_request_number,
            )
            results.append(ds)

        return results

    def build_llm_prompt(
        self,
        discovery_type: DiscoveryType,
        user_instructions: str,
        custom_style: CustomStyle = CustomStyle.CUSTOM_ONLY,
        context_text: str = "",
        start_number: int = 1,
        standard_requests_text: str = "",
    ) -> str:
        """Build a prompt instructing the LLM to generate discovery requests.

        The prompt includes the discovery type name, header template format,
        start number, user instructions, and optional context documents.

        Parameters
        ----------
        discovery_type : DiscoveryType
            The type of discovery (SI, RPD, RFA).
        user_instructions : str
            User's instructions for what to generate.
        custom_style : CustomStyle
            The custom generation style.
        context_text : str
            Additional context documents text.
        start_number : int
            The number for the first generated request.
        standard_requests_text : str
            For MODIFIED_STANDARD: the standard requests text to modify.
            For STANDARD_PLUS_CUSTOM: the standard requests for reference.

        Returns
        -------
        str
            The constructed LLM prompt.
        """
        header_template = discovery_type.request_header_template
        header_example = header_template.format(number=start_number)
        section_heading = discovery_type.section_heading

        parts: List[str] = []

        # -- Preamble: role and task description --
        parts.append(
            "You are a California litigation attorney drafting written discovery. "
            f"Generate {section_heading} formatted as numbered requests."
        )
        parts.append("")

        # -- Format instructions --
        parts.append("FORMAT REQUIREMENTS:")
        parts.append(f"- Each request MUST start with a header: {header_example}")
        parts.append(
            f"- Number requests sequentially starting at NO. {start_number}"
        )
        parts.append("- Output ONLY the numbered requests, no other text")
        parts.append("- Separate each request with a blank line")
        parts.append("")

        # -- Style-specific instructions --
        if custom_style == CustomStyle.MODIFIED_STANDARD:
            parts.append("TASK: Modify the following standard discovery requests "
                         "according to the user's instructions below.")
            parts.append("")
            parts.append("STANDARD REQUESTS TO MODIFY:")
            parts.append(standard_requests_text)
            parts.append("")

        elif custom_style == CustomStyle.STANDARD_PLUS_CUSTOM:
            parts.append(
                f"NOTE: A standard set of requests (numbered 1 through "
                f"{start_number - 1}) has already been generated. "
                f"Generate ADDITIONAL custom requests starting at "
                f"NO. {start_number}."
            )
            parts.append("")
            if standard_requests_text:
                parts.append("EXISTING STANDARD REQUESTS (for reference, do not repeat):")
                parts.append(standard_requests_text)
                parts.append("")

        elif custom_style == CustomStyle.CUSTOM_ONLY:
            parts.append(
                "TASK: Generate original discovery requests based on the "
                "user's instructions and any context provided."
            )
            parts.append("")

        # -- User instructions --
        parts.append("USER INSTRUCTIONS:")
        parts.append(user_instructions)
        parts.append("")

        # -- Context documents --
        if context_text:
            parts.append("CONTEXT DOCUMENTS:")
            parts.append(context_text)
            parts.append("")

        return "\n".join(parts)

    def parse_llm_response(
        self, response: str, discovery_type: DiscoveryType
    ) -> List[DiscoveryRequest]:
        """Parse an LLM response into a list of DiscoveryRequest objects.

        Delegates to ``extract_requests_from_text()``.

        Parameters
        ----------
        response : str
            The raw LLM response text.
        discovery_type : DiscoveryType
            The type of discovery for header pattern matching.

        Returns
        -------
        list of DiscoveryRequest
        """
        return extract_requests_from_text(response, discovery_type)

    # ------------------------------------------------------------------
    # Private helpers
    # ------------------------------------------------------------------

    def _load_definitions(
        self, variables: Optional[Dict[str, str]] = None
    ) -> str:
        """Load defined terms and apply variable substitution."""
        try:
            definitions = self.loader.load_defined_terms()
        except FileNotFoundError:
            logger.warning("Defined terms file not found")
            return ""

        if variables:
            definitions = substitute_variables(definitions, variables)

        return definitions

    def _get_template_path(
        self, standard_type: str, disc_type: DiscoveryType
    ) -> Optional[str]:
        """Build the path to a template file.

        Returns None if the standard type or discovery type filename
        is not configured.
        """
        folder = _STANDARD_TYPE_FOLDERS.get(standard_type)
        if not folder:
            return None
        filename = _TEMPLATE_FILENAMES.get(disc_type.abbreviation)
        if not filename:
            return None
        return os.path.join(self.discovery_dir, folder, filename)
