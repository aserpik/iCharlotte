"""
Liability Evaluation JSON Schema

Pydantic models for structured LLM output that feeds into the
LiabilityEvaluation.jinja2.html template.

Based on the structure from Carrier001.docx EVALUATION OF LIABILITY section.
"""

from typing import Optional, List
from pydantic import BaseModel, Field
from enum import Enum


class LiabilityType(str, Enum):
    """Types of liability cases."""
    AUTO = "Auto"
    PREMISES_LIABILITY = "Premises Liability"
    WRONGFUL_DEATH = "Wrongful Death"
    NUISANCE = "Nuisance"
    HABITABILITY = "Habitability"
    DANGEROUS_CONDITION = "Dangerous Condition of Public Property"
    GENERAL_NEGLIGENCE = "General Negligence"
    PRODUCTS_LIABILITY = "Products Liability"
    MEDICAL_MALPRACTICE = "Medical Malpractice"
    OTHER = "Other"


class LiabilityStrength(str, Enum):
    """Assessment of liability strength."""
    STRONG_FOR_DEFENSE = "Strong for Defense"
    FAVORABLE_FOR_DEFENSE = "Favorable for Defense"
    UNCERTAIN = "Uncertain"
    FAVORABLE_FOR_PLAINTIFF = "Favorable for Plaintiff"
    STRONG_FOR_PLAINTIFF = "Strong for Plaintiff"


class EvidenceItem(BaseModel):
    """A piece of evidence referenced in the analysis."""
    description: str = Field(..., description="Brief description of the evidence")
    source: Optional[str] = Field(None, description="Source document or testimony")
    favorable_to: Optional[str] = Field(None, description="Which party this evidence favors: 'defense', 'plaintiff', or 'neutral'")


class LegalCitation(BaseModel):
    """A legal citation (case or statute)."""
    citation: str = Field(..., description="The citation text, e.g., 'CACI No. 1000' or 'Alcaraz v. Vece, 14 Cal. 4th 1149 (1997)'")
    relevance: Optional[str] = Field(None, description="Brief note on why this citation is relevant")


class AnalysisSection(BaseModel):
    """A section of liability analysis with arguments from both sides."""

    legal_standard: Optional[str] = Field(
        None,
        description="The legal standard or elements that apply to this section"
    )

    our_arguments: List[str] = Field(
        default_factory=list,
        description="Arguments we will make in favor of our client (defense)"
    )

    opposing_arguments: List[str] = Field(
        default_factory=list,
        description="Arguments plaintiff will make against our client"
    )

    key_evidence: List[EvidenceItem] = Field(
        default_factory=list,
        description="Key evidence relevant to this section"
    )

    citations: List[LegalCitation] = Field(
        default_factory=list,
        description="Legal citations supporting the analysis"
    )

    analysis: str = Field(
        ...,
        description="Narrative analysis tying together the arguments and evidence"
    )

    assessment: Optional[LiabilityStrength] = Field(
        None,
        description="Assessment of how this element favors defense vs plaintiff"
    )


class CoDefendant(BaseModel):
    """A co-defendant in the case for risk transfer analysis."""
    name: str = Field(..., description="Name of the co-defendant")
    role: Optional[str] = Field(None, description="Role in the incident (e.g., 'property owner', 'general contractor')")
    basis_for_liability: Optional[str] = Field(None, description="Why this party may bear liability")
    estimated_fault_percentage: Optional[str] = Field(None, description="Estimated percentage of fault, e.g., '20-30%'")


class RiskTransferSection(BaseModel):
    """Analysis of liability apportionment among multiple defendants."""

    plaintiff_comparative_fault: Optional[str] = Field(
        None,
        description="Analysis of plaintiff's own negligence and estimated fault percentage"
    )

    co_defendants: List[CoDefendant] = Field(
        default_factory=list,
        description="Other defendants and their share of liability"
    )

    our_client_exposure: Optional[str] = Field(
        None,
        description="Analysis of our client's likely share of liability"
    )

    analysis: str = Field(
        ...,
        description="Narrative analysis of how liability will be apportioned"
    )


class LiabilityEvaluation(BaseModel):
    """
    Complete liability evaluation structure.

    This is the root model that the LLM should output as JSON.
    It feeds directly into the LiabilityEvaluation.jinja2.html template.
    """

    # Case identification
    case_name: Optional[str] = Field(None, description="Case name, e.g., 'Haydel v. Pacific Painting'")
    our_client: Optional[str] = Field(None, description="Name of the defendant we represent")
    liability_type: LiabilityType = Field(..., description="Type of liability case")

    # Causes of action
    causes_of_action: List[str] = Field(
        default_factory=list,
        description="List of causes of action alleged against our client, e.g., ['General Negligence', 'Premises Liability']"
    )

    # Legal standard / Elements overview
    legal_standard_overview: str = Field(
        ...,
        description="Opening paragraph explaining the legal elements plaintiff must prove. Include CACI or other jury instruction references."
    )

    # Main analysis sections
    duty_and_control: AnalysisSection = Field(
        ...,
        description="Analysis of whether defendant owed a duty or had control over the premises/situation"
    )

    breach_of_duty: AnalysisSection = Field(
        ...,
        description="Analysis of whether defendant breached their duty of care"
    )

    causation: AnalysisSection = Field(
        ...,
        description="Analysis of whether defendant's actions were a substantial factor in causing harm"
    )

    comparative_fault: Optional[AnalysisSection] = Field(
        None,
        description="Analysis of plaintiff's own negligence and open/obvious conditions. Include if applicable."
    )

    risk_transfer: Optional[RiskTransferSection] = Field(
        None,
        description="Analysis of liability apportionment among multiple defendants. Include if there are co-defendants."
    )

    # Summary
    summary: str = Field(
        ...,
        description="Concluding paragraph summarizing the overall liability assessment and key factors"
    )

    overall_assessment: LiabilityStrength = Field(
        ...,
        description="Overall assessment of liability exposure"
    )

    class Config:
        use_enum_values = True
        json_schema_extra = {
            "example": {
                "case_name": "Haydel v. Pacific Painting",
                "our_client": "Pacific Painting",
                "liability_type": "Premises Liability",
                "causes_of_action": ["General Negligence", "Premises Liability"],
                "legal_standard_overview": "Plaintiff's Complaint alleges causes of action for General Negligence and Premises Liability against Pacific Painting. In order to prevail on these claims, Plaintiff must demonstrate that Pacific Painting: (1) owned, leased, occupied, or controlled the property or created a dangerous condition thereon; (2) was negligent in the use or maintenance of the property or the execution of its work; (3) that Plaintiff was harmed; and (4) that Pacific Painting's negligence was a substantial factor in causing Plaintiff's harm. (CACI No. 1000, 1001.)",
                "summary": "Based on the testimony and evidence developed to date, we believe that Plaintiff has a viable pathway to establish liability against Pacific Painting.",
                "overall_assessment": "Favorable for Plaintiff"
            }
        }


def get_json_schema() -> dict:
    """Return the JSON schema for LLM structured output."""
    return LiabilityEvaluation.model_json_schema()


def parse_liability_json(json_str: str) -> LiabilityEvaluation:
    """Parse JSON string into LiabilityEvaluation model with validation."""
    import json
    data = json.loads(json_str)
    return LiabilityEvaluation.model_validate(data)


def get_llm_prompt_schema() -> str:
    """
    Return a simplified schema description for including in LLM prompts.
    This is more readable than the full JSON schema.
    """
    return '''
You must output a JSON object with the following structure:

{
  "case_name": "string - Case name",
  "our_client": "string - Name of defendant we represent",
  "liability_type": "string - One of: Auto, Premises Liability, Wrongful Death, Nuisance, Habitability, Dangerous Condition of Public Property, General Negligence, Products Liability, Medical Malpractice, Other",
  "causes_of_action": ["string array - Causes of action alleged"],

  "legal_standard_overview": "string - Opening paragraph with legal elements plaintiff must prove, including CACI citations",

  "duty_and_control": {
    "legal_standard": "string (optional) - Legal standard for duty/control",
    "our_arguments": ["string array - Arguments in favor of defense"],
    "opposing_arguments": ["string array - Arguments plaintiff will make"],
    "key_evidence": [
      {
        "description": "string - Description of evidence",
        "source": "string (optional) - Source document",
        "favorable_to": "string (optional) - 'defense', 'plaintiff', or 'neutral'"
      }
    ],
    "citations": [
      {
        "citation": "string - Legal citation",
        "relevance": "string (optional) - Why relevant"
      }
    ],
    "analysis": "string - Narrative analysis paragraph",
    "assessment": "string (optional) - One of: Strong for Defense, Favorable for Defense, Uncertain, Favorable for Plaintiff, Strong for Plaintiff"
  },

  "breach_of_duty": { /* same structure as duty_and_control */ },

  "causation": { /* same structure as duty_and_control */ },

  "comparative_fault": { /* same structure as duty_and_control, optional - include if plaintiff has contributory negligence */ },

  "risk_transfer": {
    "plaintiff_comparative_fault": "string (optional) - Plaintiff's own negligence analysis",
    "co_defendants": [
      {
        "name": "string - Co-defendant name",
        "role": "string (optional) - Role in incident",
        "basis_for_liability": "string (optional) - Why they may be liable",
        "estimated_fault_percentage": "string (optional) - e.g., '20-30%'"
      }
    ],
    "our_client_exposure": "string (optional) - Our client's likely share",
    "analysis": "string - Narrative analysis of liability apportionment"
  },

  "summary": "string - Concluding paragraph summarizing liability assessment",

  "overall_assessment": "string - One of: Strong for Defense, Favorable for Defense, Uncertain, Favorable for Plaintiff, Strong for Plaintiff"
}
'''
