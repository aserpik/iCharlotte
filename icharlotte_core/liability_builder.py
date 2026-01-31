"""
Liability Evaluation Builder

Utilities for generating liability evaluations using the JSON schema + Jinja2 template approach.

Usage:
    from icharlotte_core.liability_builder import LiabilityBuilder

    builder = LiabilityBuilder()

    # Generate prompt for LLM
    prompt = builder.build_prompt(documents_text, liability_type="Premises Liability")

    # Parse LLM response and render to HTML
    html = builder.parse_and_render(llm_json_response)
"""

import os
import json
import re
from typing import Optional, Dict, Any, List
from jinja2 import Environment, FileSystemLoader, select_autoescape

from .liability_schema import (
    LiabilityEvaluation,
    LiabilityType,
    get_llm_prompt_schema,
    parse_liability_json
)


class LiabilityBuilder:
    """Builder for generating liability evaluations using LLM + templates."""

    TEMPLATES_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'Templates')
    TEMPLATE_NAME = 'LiabilityEvaluation.jinja2.html'

    def __init__(self):
        """Initialize the builder with Jinja2 environment."""
        self.env = Environment(
            loader=FileSystemLoader(self.TEMPLATES_DIR),
            autoescape=select_autoescape(['html', 'xml'])
        )
        self.template = self.env.get_template(self.TEMPLATE_NAME)

    def build_prompt(
        self,
        documents_text: str,
        liability_type: str = "Premises Liability",
        our_client: Optional[str] = None,
        case_name: Optional[str] = None,
        additional_context: Optional[str] = None
    ) -> str:
        """
        Build the LLM prompt for generating a liability evaluation.

        Args:
            documents_text: Combined text from all relevant documents
            liability_type: Type of liability case
            our_client: Name of the defendant we represent
            case_name: Case name for reference
            additional_context: Any additional instructions or context

        Returns:
            Complete prompt string for the LLM
        """
        schema = get_llm_prompt_schema()

        prompt = f'''You are a legal analyst preparing an EVALUATION OF LIABILITY section for a litigation report.

Analyze the following documents and generate a comprehensive liability evaluation.

**Case Information:**
- Liability Type: {liability_type}
{f"- Our Client (Defendant): {our_client}" if our_client else ""}
{f"- Case Name: {case_name}" if case_name else ""}

**Instructions:**

1. Read and analyze all provided documents carefully.

2. Identify the causes of action alleged against our client.

3. For each element of liability (Duty/Control, Breach, Causation), provide:
   - The legal standard that applies
   - Arguments we will make in defense of our client
   - Arguments the plaintiff will likely make
   - Key evidence supporting each side
   - Your analysis of how this element is likely to be resolved

4. If applicable, analyze:
   - Comparative fault / open and obvious condition defenses
   - Risk transfer among co-defendants

5. Provide a summary with your overall assessment.

6. Use proper legal citations (CACI jury instructions, case law) where appropriate.

7. Write in a professional legal tone, suitable for a report to an insurance carrier.

8. Be balanced - acknowledge both strengths and weaknesses of our position.

{f"**Additional Context:** {additional_context}" if additional_context else ""}

**Output Format:**

{schema}

**IMPORTANT:**
- Output ONLY valid JSON, no markdown code blocks or other formatting
- Ensure all string values are properly escaped
- The "analysis" field in each section should be a complete narrative paragraph, not bullet points
- Reference specific testimony, documents, and evidence by name

**Documents to Analyze:**

{documents_text}

**Generate the JSON output now:**'''

        return prompt

    def parse_and_render(self, json_response: str) -> str:
        """
        Parse LLM JSON response and render to HTML.

        Args:
            json_response: JSON string from LLM

        Returns:
            Rendered HTML string

        Raises:
            ValueError: If JSON is invalid or doesn't match schema
        """
        # Clean up common LLM output issues
        cleaned = self._clean_json_response(json_response)

        # Parse and validate
        evaluation = parse_liability_json(cleaned)

        # Convert to dict for template
        data = evaluation.model_dump()

        # Render template
        html = self.template.render(data=data)

        return html

    def render_from_dict(self, data: Dict[str, Any]) -> str:
        """
        Render HTML from a dictionary (already parsed/validated).

        Args:
            data: Dictionary matching LiabilityEvaluation schema

        Returns:
            Rendered HTML string
        """
        return self.template.render(data=data)

    def validate_json(self, json_response: str) -> tuple[bool, Optional[str], Optional[Dict]]:
        """
        Validate JSON response against schema.

        Args:
            json_response: JSON string to validate

        Returns:
            Tuple of (is_valid, error_message, parsed_data)
        """
        try:
            cleaned = self._clean_json_response(json_response)
            evaluation = parse_liability_json(cleaned)
            return True, None, evaluation.model_dump()
        except json.JSONDecodeError as e:
            return False, f"Invalid JSON: {str(e)}", None
        except Exception as e:
            return False, f"Schema validation error: {str(e)}", None

    def _clean_json_response(self, response: str) -> str:
        """
        Clean up common LLM JSON output issues.

        - Remove markdown code blocks
        - Fix trailing commas
        - Handle escaped characters
        """
        text = response.strip()

        # Remove markdown code blocks
        if text.startswith('```'):
            # Find the end of the first line (```json or ```)
            first_newline = text.find('\n')
            if first_newline != -1:
                text = text[first_newline + 1:]

            # Remove trailing ```
            if text.endswith('```'):
                text = text[:-3]

            text = text.strip()

        # Remove any leading/trailing non-JSON content
        # Find the first { and last }
        start = text.find('{')
        end = text.rfind('}')
        if start != -1 and end != -1 and end > start:
            text = text[start:end + 1]

        # Fix trailing commas before } or ]
        text = re.sub(r',\s*}', '}', text)
        text = re.sub(r',\s*]', ']', text)

        return text

    def get_example_output(self) -> Dict[str, Any]:
        """
        Return an example output structure for testing/documentation.
        """
        return {
            "case_name": "Haydel v. Pacific Painting",
            "our_client": "Pacific Painting",
            "liability_type": "Premises Liability",
            "causes_of_action": ["General Negligence", "Premises Liability"],

            "legal_standard_overview": "In order to prevail on these claims, Plaintiff must demonstrate that Pacific Painting: (1) owned, leased, occupied, or controlled the property or created a dangerous condition thereon; (2) was negligent in the use or maintenance of the property or the execution of its work; (3) that Plaintiff was harmed; and (4) that Pacific Painting's negligence was a substantial factor in causing Plaintiff's harm. (CACI No. 1000, 1001.)",

            "duty_and_control": {
                "legal_standard": "Plaintiff must establish that Pacific Painting owned, leased, occupied, or controlled the area where the incident occurred, or created the dangerous condition. '[P]roperty owners are liable for injuries on land they own, possess, or control.' (Alcaraz v. Vece, 14 Cal. 4th 1149 (1997).) However, 'a defendant need not own, possess and control property in order to be held liable; control alone is sufficient.' (Id.)",
                "our_arguments": [
                    "Pacific Painting did not own or control the premises generally, as that duty lies with the property owners.",
                    "Pacific Painting maintains that it was not present at the property on the date of the incident due to rain.",
                    "Kwang Hae Chong testified that he and his crew left the site on March 4, 2022, and did not return on the day of the loss."
                ],
                "opposing_arguments": [
                    "Photographs from the date of the incident showing ladders suggest Pacific Painting's equipment was present.",
                    "As the painting subcontractor, Pacific Painting exercised control over their work area."
                ],
                "key_evidence": [
                    {
                        "description": "Photographs showing ladders at the scene on incident date",
                        "source": "Valerie Smith production",
                        "favorable_to": "plaintiff"
                    },
                    {
                        "description": "Chong testimony that crew left site on March 4, 2022",
                        "source": "Deposition of Kwang Hae Chong",
                        "favorable_to": "defense"
                    }
                ],
                "citations": [
                    {
                        "citation": "Alcaraz v. Vece, 14 Cal. 4th 1149 (1997)",
                        "relevance": "Control alone is sufficient for premises liability"
                    }
                ],
                "analysis": "We will argue that Pacific Painting did not own or control the premises generally, as that duty lies with the property owners. While a contractor has a duty to maintain their immediate work area in a safe condition, Pacific Painting maintains that it was not present at the property on the date of the incident, March 5, 2022, due to rain. Kwang Hae Chong testified that he and his crew left the site on March 4, 2022, and did not return on the day of the loss. However, Valerie Smith's recent production of photographs from the date of the incident depicting ladders leaning against the building significantly weakens this defense. If these photographs are authenticated, they suggest that Pacific Painting's equipment was present, and potentially that work was being performed or the site was not cleared, thereby establishing the requisite control over the premises.",
                "assessment": "Favorable for Plaintiff"
            },

            "breach_of_duty": {
                "analysis": "Plaintiff will argue that Pacific Painting, as the painting subcontractor, placed the plastic sheeting on the stairs in the course of their work and failed to remove it or warn of its slippery nature, thereby creating the hazard that caused the fall. On the other hand, we will argue that Pacific Painting did not create the dangerous condition in question. Mr. Chong testified that Pacific Painting has a strict policy of removing all protective coverings and debris at the conclusion of each workday. He stated he was personally on-site on March 4, 2022, and confirmed that all plastic was removed before the crew departed. The testimony of non-party witness, Valerie Smith, previously supported this defense. Ms. Smith testified that the painting crew had not been at the location where the incident happened for approximately 3 days prior to the incident. However, the new photographs produced by Ms. Smith directly contradict both Pacific Painting's testimony as well as Ms. Smith's own testimony. If Pacific Painting's equipment was present, it becomes difficult to argue that the plastic sheeting, which Mr. Chong admitted is similar to what he uses, was placed by a third party.",
                "our_arguments": [
                    "Pacific Painting has strict policy of removing all coverings at end of each workday",
                    "Mr. Chong confirmed all plastic was removed before crew departed on March 4"
                ],
                "opposing_arguments": [
                    "Photographs show Pacific Painting equipment present on incident date",
                    "Chong admitted plastic sheeting is similar to what Pacific Painting uses"
                ],
                "key_evidence": [],
                "citations": [],
                "assessment": "Favorable for Plaintiff"
            },

            "causation": {
                "analysis": "Plaintiff must also establish that Pacific Painting's negligence was a substantial factor in causing his harm. Based on the investigation to date, significant questions exist regarding the legitimacy of the incident itself. If the incident was staged, there is no negligence and no liability. Additionally, if the plastic was placed by the Landlord's handyman on the morning of the incident as suspected by Ms. Smith, Pacific Painting's actions cannot be a substantial factor in Plaintiff's harm. However, this argument is now tenuous given the photographic evidence placing Pacific Painting's ladders at the scene. If the jury finds we left the ladders and plastic, the chain of causation remains with Pacific Painting.",
                "our_arguments": [
                    "Significant questions about legitimacy of incident - may have been staged",
                    "Plastic may have been placed by Landlord's handyman, not Pacific Painting"
                ],
                "opposing_arguments": [
                    "Photographic evidence places Pacific Painting equipment at scene",
                    "If plastic was Pacific Painting's, causation chain is established"
                ],
                "key_evidence": [],
                "citations": [],
                "assessment": "Uncertain"
            },

            "comparative_fault": {
                "analysis": "Even if Plaintiff can establish that Pacific Painting was responsible for the plastic, we will argue that the condition was open and obvious. Plaintiff testified that he saw the plastic covering when he ascended the stairs immediately prior to the incident. He acknowledged stepping on it on his way up. If a danger is so obvious that a person could reasonably be expected to see it, the condition itself serves as a warning. (CACI No. 1004.) We will argue that, despite being aware of the rain and the presence of the plastic, Plaintiff failed to exercise reasonable care for his own safety by proceeding down the stairs, particularly given his testimony that he may have been distracted by his phone at the time of the fall. This will serve to shift a significant percentage of liability, if any is found, to Plaintiff under the doctrine of comparative negligence.",
                "our_arguments": [
                    "Condition was open and obvious - Plaintiff saw plastic on way up",
                    "Plaintiff admitted being distracted by phone during fall",
                    "Plaintiff failed to exercise reasonable care despite awareness of conditions"
                ],
                "opposing_arguments": [
                    "Wet plastic became more slippery than when Plaintiff first encountered it",
                    "No warning signs or barriers were present"
                ],
                "key_evidence": [],
                "citations": [
                    {
                        "citation": "CACI No. 1004",
                        "relevance": "Open and obvious danger serves as its own warning"
                    }
                ],
                "assessment": "Favorable for Defense"
            },

            "risk_transfer": {
                "plaintiff_comparative_fault": "We anticipate that Plaintiff will be assigned some comparative fault based on his awareness of the plastic and his admitted distraction with his phone.",
                "co_defendants": [
                    {
                        "name": "Nina Jahanbin",
                        "role": "property owner",
                        "basis_for_liability": "based on her non-delegable duty as the property owner as well as witness testimony indicating that Jahanbin had actual notice of the hazard prior to the fall.",
                        "estimated_fault_percentage": "15-25%"
                    },
                    {
                        "name": "Young Moon Painting",
                        "role": "general contractor",
                        "basis_for_liability": "as they are the party that secured the prime contract with the property owner and stand in the shoes of Pacific Painting as the general contractor, facing liability for negligent hiring and supervision.",
                        "estimated_fault_percentage": "10-20%"
                    }
                ],
                "our_client_exposure": "As the party contracted to perform the physical painting work, Pacific Painting will likely be allocated the largest proportion of fault, as they were the party responsible for the site work.",
                "analysis": "In addition to Plaintiff's comparative fault, we anticipate that a jury will apportion liability among all of the named defendants, as each bears a distinct share of responsibility for the incident."
            },

            "summary": "Based on the testimony and evidence developed to date, we believe that Plaintiff has a viable pathway to establish liability against Pacific Painting. While prior evidence suggested that Pacific Painting was not present on the day of the incident and had cleared the site the day prior, the newly produced photographs from the date of the incident, if authenticated, have significant implications to the liability analysis. These photographs suggest that Pacific Painting's equipment was present on the day of the incident and contradict Mr. Chong and Ms. Smith's prior testimony, damaging their credibility. While we will still pursue the 'open and obvious' defense and question the legitimacy of the fall, Plaintiff will likely rely on these photographs to establish that Pacific Painting's crew, or at least their equipment, was present at the incident location at the time of the incident.",

            "overall_assessment": "Favorable for Plaintiff"
        }


# Convenience function for quick rendering
def render_liability_evaluation(json_data: str | Dict[str, Any]) -> str:
    """
    Quick helper to render a liability evaluation.

    Args:
        json_data: Either a JSON string or dict matching the schema

    Returns:
        Rendered HTML
    """
    builder = LiabilityBuilder()

    if isinstance(json_data, str):
        return builder.parse_and_render(json_data)
    else:
        return builder.render_from_dict(json_data)
