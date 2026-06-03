"""Per-type configuration for the Generate Motion task.

Each ``MotionTypeConfig`` carries the type-specific knowledge that makes a
generated motion correct: the legal standard to ground the Legal Standard
section, a default section plan, the labels of attachments emitted as
placeholders, and prompt guidance for the target-document analyzer.

Three California civil motion types ship fully configured (compel, demurrer,
strike). Any other ``type_id`` falls back to the generic config.
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional


@dataclass(frozen=True)
class MotionTypeConfig:
    type_id: str
    display_name: str
    target_doc_guidance: str
    legal_standard_hint: str
    section_plan: List[str] = field(default_factory=list)
    placeholder_attachments: List[str] = field(default_factory=list)
    analyzer_prompt: str = ""
    grounds_prompt: str = ""


# Standard memorandum spine shared by every motion type.
_BASE_SECTIONS = [
    "Introduction",
    "Statement of Facts",
    "Legal Standard",
    "Argument",
    "Conclusion",
]


MOTION_TYPE_CONFIGS: Dict[str, MotionTypeConfig] = {
    "compel": MotionTypeConfig(
        type_id="compel",
        display_name="Motion to Compel Further Responses",
        target_doc_guidance=(
            "Add the discovery requests and the served responses at issue "
            "(e.g. the propounded interrogatories/RFPs and the responding "
            "party's verified responses)."
        ),
        legal_standard_hint=(
            "A motion to compel further responses to interrogatories is "
            "governed by Code of Civil Procedure section 2030.300; a motion "
            "to compel further responses to requests for production is "
            "governed by Code of Civil Procedure section 2031.310. The motion "
            "must be accompanied by a meet-and-confer declaration (CCP "
            "2016.040) and a separate statement (Cal. Rules of Court, rule "
            "3.1345)."
        ),
        section_plan=_BASE_SECTIONS,
        placeholder_attachments=[
            "Meet and Confer Declaration",
            "Separate Statement (Cal. Rules of Court, rule 3.1345)",
        ],
        analyzer_prompt=(
            "Identify the specific discovery requests and the served responses "
            "that are deficient, and explain why each response is inadequate "
            "(boilerplate objections, evasive or incomplete answers, "
            "unsupported privilege claims)."
        ),
        grounds_prompt=(
            "Propose the grounds and the relief sought for a motion to compel "
            "further responses: which requests are at issue and why further "
            "responses are warranted."
        ),
    ),
    "demurrer": MotionTypeConfig(
        type_id="demurrer",
        display_name="Demurrer",
        target_doc_guidance=(
            "Add the complaint (or cross-complaint) being challenged."
        ),
        legal_standard_hint=(
            "A demurrer tests the legal sufficiency of a pleading. Under Code "
            "of Civil Procedure section 430.10, subdivision (e), a demurrer "
            "lies where the pleading does not state facts sufficient to "
            "constitute a cause of action. A demurrer must be accompanied by a "
            "meet-and-confer declaration (CCP 430.41)."
        ),
        section_plan=_BASE_SECTIONS,
        placeholder_attachments=[
            "Meet and Confer Declaration (CCP 430.41)",
        ],
        analyzer_prompt=(
            "Identify each cause of action in the pleading that fails to state "
            "facts sufficient to constitute a cause of action, and explain the "
            "missing or defective element for each."
        ),
        grounds_prompt=(
            "Propose the grounds for a demurrer: which causes of action are "
            "subject to demurrer and on what statutory basis."
        ),
    ),
    "strike": MotionTypeConfig(
        type_id="strike",
        display_name="Motion to Strike",
        target_doc_guidance=(
            "Add the complaint (or cross-complaint) containing the matter to "
            "be stricken (e.g. punitive-damages allegations)."
        ),
        legal_standard_hint=(
            "A motion to strike is governed by Code of Civil Procedure "
            "sections 435 and 436, which permit the court to strike out any "
            "irrelevant, false, or improper matter, or any part of a pleading "
            "not filed in conformity with the law. A motion to strike must be "
            "accompanied by a meet-and-confer declaration (CCP 435.5)."
        ),
        section_plan=_BASE_SECTIONS,
        placeholder_attachments=[
            "Meet and Confer Declaration (CCP 435.5)",
        ],
        analyzer_prompt=(
            "Identify the specific allegations or portions of the pleading "
            "that are irrelevant, false, or improper and should be stricken "
            "(e.g. unsupported punitive-damages or attorney-fee allegations), "
            "and explain why each is improper."
        ),
        grounds_prompt=(
            "Propose the grounds for a motion to strike: which portions of the "
            "pleading should be stricken and why."
        ),
    ),
    "generic": MotionTypeConfig(
        type_id="generic",
        display_name="Motion",
        target_doc_guidance=(
            "Add any documents relevant to the motion you want to bring."
        ),
        legal_standard_hint="",
        section_plan=_BASE_SECTIONS,
        placeholder_attachments=[],
        analyzer_prompt=(
            "Identify the grounds for the motion the user wants to bring and "
            "the relief sought, based on the supplied documents."
        ),
        grounds_prompt=(
            "Propose the grounds and the relief sought for the motion, based "
            "on the supplied documents and the user's description."
        ),
    ),
}


def get_motion_config(type_id: Optional[str]) -> MotionTypeConfig:
    """Return the config for ``type_id``; unknown/empty ids fall back to generic."""
    if not type_id:
        return MOTION_TYPE_CONFIGS["generic"]
    return MOTION_TYPE_CONFIGS.get(type_id, MOTION_TYPE_CONFIGS["generic"])
