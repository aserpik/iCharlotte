"""Per-type configuration for the Generate Motion task.

Each ``MotionTypeConfig`` carries the type-specific knowledge that makes a
generated motion correct: the legal standard to ground the Legal Standard
section, a default section plan, the labels of attachments emitted as
placeholders, and prompt guidance for the target-document analyzer.

Three California civil motion types ship fully configured (compel, demurrer,
strike). Any other ``type_id`` falls back to the generic config.
"""
import os
from dataclasses import asdict, dataclass, field
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

    def to_dict(self) -> Dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Optional[Dict]) -> "MotionTypeConfig":
        data = data or {}
        return cls(
            type_id=data.get("type_id", ""),
            display_name=data.get("display_name", ""),
            target_doc_guidance=data.get("target_doc_guidance", ""),
            legal_standard_hint=data.get("legal_standard_hint", ""),
            section_plan=list(data.get("section_plan", []) or []),
            placeholder_attachments=list(data.get("placeholder_attachments", []) or []),
            analyzer_prompt=data.get("analyzer_prompt", ""),
            grounds_prompt=data.get("grounds_prompt", ""),
        )


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
    "msj": MotionTypeConfig(
        type_id="msj", display_name="Motion for Summary Judgment/Adjudication",
        target_doc_guidance="Add the separate statement of undisputed material facts and the supporting evidence (declarations, discovery excerpts) the motion relies on.",
        legal_standard_hint="A motion for summary judgment/adjudication is governed by Code of Civil Procedure section 437c. Summary judgment is proper only where there is no triable issue of material fact and the moving party is entitled to judgment as a matter of law (CCP 437c(c)); a defendant meets its burden by showing an element cannot be established or a complete defense exists (CCP 437c(p)(2)).",
        section_plan=_BASE_SECTIONS),
    "ex_parte": MotionTypeConfig(
        type_id="ex_parte", display_name="Ex Parte Application",
        target_doc_guidance="Add the supporting declaration showing irreparable harm/urgency and the notice given to opposing counsel (Cal. Rules of Court, rule 3.1204).",
        legal_standard_hint="Ex parte relief is governed by California Rules of Court, rules 3.1200-3.1207. The applicant must make an affirmative factual showing of irreparable harm, immediate danger, or other statutory basis for ex parte relief, and must give notice by 10:00 a.m. the court day before.",
        section_plan=_BASE_SECTIONS),
    "ime": MotionTypeConfig(
        type_id="ime", display_name="Motion for Leave to Conduct IME",
        target_doc_guidance="Add the showing of good cause and the proposed examination's time, place, manner, conditions, scope, and examiner (CCP 2032.310, 2032.320).",
        legal_standard_hint="A motion for a physical or mental examination is governed by Code of Civil Procedure sections 2032.310 and 2032.320. The motion must specify the time, place, manner, conditions, scope, and nature of the examination and the examiner, and must be supported by a showing of good cause.",
        section_plan=_BASE_SECTIONS),
    "gfs": MotionTypeConfig(
        type_id="gfs", display_name="Motion for Determination of Good Faith Settlement",
        target_doc_guidance="Add the settlement terms and the facts bearing on the Tech-Bilt factors (settlor's proportionate liability, amount paid, allocation).",
        legal_standard_hint="A determination of good faith settlement is governed by Code of Civil Procedure section 877.6 and evaluated under the factors in Tech-Bilt, Inc. v. Woodward-Clyde & Associates (1985) 38 Cal.3d 488. A good faith determination bars other defendants' claims for equitable contribution or indemnity.",
        section_plan=_BASE_SECTIONS),
    "dismiss": MotionTypeConfig(
        type_id="dismiss", display_name="Motion to Dismiss",
        target_doc_guidance="Add the procedural basis for dismissal (e.g., failure to prosecute, forum non conveniens) and supporting facts.",
        legal_standard_hint="Grounds for dismissal include the discretionary and mandatory dismissal statutes (Code of Civil Procedure sections 583.130-583.430) and forum non conveniens (CCP 410.30).",
        section_plan=_BASE_SECTIONS),
    "leave": MotionTypeConfig(
        type_id="leave", display_name="Motion for Leave (Amend/File)",
        target_doc_guidance="Add the proposed amended pleading or cross-complaint and a declaration explaining the amendment and its timing (Cal. Rules of Court, rule 3.1324).",
        legal_standard_hint="Leave to amend is governed by Code of Civil Procedure sections 473(a) and 576 and is liberally granted; the motion must comply with California Rules of Court, rule 3.1324 (copy of the proposed amendment and a supporting declaration). Leave to file a cross-complaint is governed by CCP 426.50 / 428.50.",
        section_plan=_BASE_SECTIONS),
    "consolidate": MotionTypeConfig(
        type_id="consolidate", display_name="Motion to Consolidate",
        target_doc_guidance="Add the case captions/numbers to be consolidated and the common questions of law or fact.",
        legal_standard_hint="Consolidation is governed by Code of Civil Procedure section 1048(a): the court may consolidate actions involving a common question of law or fact to avoid unnecessary cost or delay.",
        section_plan=_BASE_SECTIONS),
    "quash": MotionTypeConfig(
        type_id="quash", display_name="Motion to Quash",
        target_doc_guidance="Add the summons/subpoena at issue and the facts showing lack of jurisdiction or an improper/overbroad subpoena.",
        legal_standard_hint="A motion to quash service of summons for lack of personal jurisdiction is governed by Code of Civil Procedure section 418.10. A motion to quash a subpoena is governed by CCP 1987.1.",
        section_plan=_BASE_SECTIONS),
    "sanctions": MotionTypeConfig(
        type_id="sanctions", display_name="Motion for Sanctions",
        target_doc_guidance="Add the conduct at issue and the prior orders/meet-and-confer establishing the basis for sanctions.",
        legal_standard_hint="Discovery sanctions are governed by Code of Civil Procedure sections 2023.010-2023.030; monetary, issue, evidence, and terminating sanctions escalate with the abuse. CCP 128.5/128.7 govern sanctions for bad-faith actions or tactics.",
        section_plan=_BASE_SECTIONS),
    "continue_trial": MotionTypeConfig(
        type_id="continue_trial", display_name="Motion to Continue Trial",
        target_doc_guidance="Add the declaration showing good cause for the continuance and the current trial/related dates (Cal. Rules of Court, rule 3.1332).",
        legal_standard_hint="Trial continuances are disfavored and granted only on an affirmative showing of good cause under California Rules of Court, rule 3.1332. Trial preference is governed by Code of Civil Procedure section 36.",
        section_plan=_BASE_SECTIONS),
    "protective_order": MotionTypeConfig(
        type_id="protective_order", display_name="Motion for Protective Order",
        target_doc_guidance="Add the discovery at issue and the facts showing annoyance, oppression, or undue burden justifying protection.",
        legal_standard_hint="Protective orders are governed by the method-specific statutes (e.g., Code of Civil Procedure sections 2025.420 for depositions, 2030.090 for interrogatories, 2031.060 for inspection demands); the court may make any order that justice requires to protect a party from unwarranted annoyance, oppression, or undue burden and expense.",
        section_plan=_BASE_SECTIONS),
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


# The hardcoded built-ins are the seed for the editable registry; the registry
# (Scripts/prompts/generate_motion/motion_types.json) is the source of truth at
# runtime once it exists. `BUILTIN_SEED` always lives in code as the fallback
# and the target of "Restore Defaults".
BUILTIN_SEED: Dict[str, MotionTypeConfig] = dict(MOTION_TYPE_CONFIGS)


def motion_types_path() -> str:
    """Canonical path of the editable motion-types registry JSON.

    Resolved relative to the project root (this file lives at
    <root>/icharlotte_core/motion_generation/config.py). Patchable in tests.
    """
    root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(root, "Scripts", "prompts", "generate_motion", "motion_types.json")


_registry_singleton = None


def _registry():
    global _registry_singleton
    if _registry_singleton is None:
        from .types_registry import MotionTypeRegistry

        _registry_singleton = MotionTypeRegistry.load(motion_types_path())
    return _registry_singleton


def reload_motion_types() -> None:
    """Drop the cached registry so the next access reloads from disk/seed.

    Called after the Workbench saves edits so running code sees the new types.
    """
    global _registry_singleton
    _registry_singleton = None


def get_motion_config(type_id: Optional[str]) -> MotionTypeConfig:
    """Return the config for ``type_id`` from the editable registry; unknown or
    empty ids fall back to the generic type."""
    return _registry().get(type_id)


def list_motion_types() -> List[MotionTypeConfig]:
    """All motion types from the editable registry, in registry order."""
    return _registry().list_types()
