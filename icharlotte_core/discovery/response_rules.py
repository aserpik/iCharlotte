"""ResponseRules dataclass — strategy settings and boilerplate defaults for discovery responses."""

import json
import os
from dataclasses import dataclass, field, fields, asdict
from typing import Any, Dict


# ---------------------------------------------------------------------------
# Module-level boilerplate constants
# ---------------------------------------------------------------------------

_DEFAULT_WAIVER = (
    "Subject to and without waiving the foregoing objections, Responding Party responds as follows:"
)

_DEFAULT_RESERVATION = (
    "Discovery and investigation are ongoing and Responding Party reserves the right to amend, "
    "modify and/or supplement this response in the future in the event that additional documents, "
    "facts and/or information are discovered, or their relevance becomes apparent."
)

# Superseded firm-standard reservation clauses. Older per-case ``respond_rules``
# persisted the prior wording (the clause appended after every substantive
# response); when such data is reloaded it is upgraded to _DEFAULT_RESERVATION so
# generated responses always carry the current firm-standard language. Genuine
# per-case customizations (anything not in this set) are left untouched.
_SUPERSEDED_RESERVATIONS = frozenset({
    "Discovery and investigation are ongoing and Responding Party reserves the right to amend, "
    "modify and/or supplement this response as additional facts and further information is "
    "obtained, new analyses are made, and legal research is completed.",
})

_DEFAULT_PRELIMINARY_FI = (
    "These responses are made solely for the purpose of this action. Each response is subject to all "
    "appropriate objections, including competency, relevancy, materiality, propriety and admissibility, "
    "which would require the exclusion of any response set forth herein if the question were asked of, "
    "or any response were made by, a witness present and testifying in court. Additionally, each response "
    "is subject to all objections listed in the responses to the Interrogatories, which shall be "
    "incorporated herein by reference. All such objections are reserved and may be interposed at the "
    "time of trial.\n\n"
    "This Responding Party has not completed its investigation of the facts relating to this action, "
    "has not yet completed its discovery in this action, and has not yet completed preparation for trial. "
    "Consequently, the following responses are given without prejudice to this Party's right to allege "
    "and/or produce evidence of any subsequently-discovered facts or circumstances.\n\n"
    "Except for facts explicitly admitted herein, no admission of any nature is to be implied or inferred. "
    "The fact that any interrogatory herein has been answered should not be taken as an admission, or a "
    "confusion of the existence of any facts set forth or assumed by such interrogatory or that such "
    "response constitutes any fact thus set forth or assumed. All responses are given on the basis of a "
    "good faith effort to locate the requested information.\n\n"
    "This Party relies on well-established California authority to the effect that interrogatories cannot "
    "be unilaterally designated as continuing in nature, and serves notice that we will not voluntarily "
    "provide further responses to these interrogatories if additional information is acquired by us after "
    "these responses are served. Notwithstanding the above, this Responding Party reserves the right to "
    "change any and all responses herein as additional facts and further information is obtained, new "
    "analyses are made, and legal research is completed. The information contained herein is given in a "
    "good faith effort to supply as much factual material as is presently known by Responding Party, but "
    "should in no way prejudice this Responding Party's right to make new contentions or provide "
    "additional facts or additional information derived from further discovery, investigation, research "
    "and/or legal analysis. This preliminary statement shall apply to each and every response given "
    "herein, and shall be incorporated by reference as though fully set forth in all of the interrogatory "
    "responses appearing on the following pages."
)

_DEFAULT_PRELIMINARY_SI = (
    "These responses are made solely for the purpose of this action. Each response is subject to all "
    "appropriate objections, including, but not limited to, objections concerning competency, relevancy, "
    "materiality, propriety, and admissibility, which would require the exclusion of any statement "
    "contained herein if the interrogatories were asked of, or any statement contained herein were made "
    "by, a witness present and testifying in a court. All such objections and grounds therefore are "
    "reserved and may be interposed at the time of trial.\n\n"
    "This Responding Party has not completed its investigation of the facts relating to this action, "
    "has not yet completed discovery in this action, and has not yet completed preparation for trial. "
    "The following answers are therefore given without prejudice to this party's right to allege and/or "
    "produce evidence of any subsequently discovered facts or circumstances.\n\n"
    "Except for facts explicitly admitted herein, no admission of any nature is to be implied or inferred. "
    "The fact that any interrogatory herein has been answered should not be taken as an admission, or "
    "confession of the existence of any facts set forth or assumed by such interrogatory or that such "
    "response constitutes evidence of any fact thus set forth or assumed. All responses are given on the "
    "basis of a good faith effort to locate the requested information.\n\n"
    "This party relies on well-established California authority to the effect that interrogatories cannot "
    "unilaterally be designated continuing in nature and serve notice that we will not voluntarily provide "
    "further responses to these interrogatories if additional information is acquired by us after these "
    "responses are served.\n\n"
    "Notwithstanding the above, this Responding Party reserves the right to change any and all responses "
    "herein as additional facts and further information is obtained, new analyses are made, and legal "
    "research is completed. The information contained herein is given in a good faith effort to supply as "
    "much factual material as is presently known by Responding Party, but should in no way prejudice this "
    "Responding Party's right to make new contentions or provide additional facts or additional information "
    "derived from further discovery, investigation, research and/or legal analysis.\n\n"
    "This preliminary statement shall apply to each and every response given herein, and shall be "
    "incorporated by reference as though fully set forth in all of the demand responses appearing on the "
    "following pages."
)

_DEFAULT_PRELIMINARY_RFA = (
    "These responses are made solely for the purpose of this action. Each response is subject to all "
    "appropriate objections, including competency, relevancy, materiality, propriety and admissibility, "
    "which would require the exclusion of any response set forth herein if the question were asked of, "
    "or any response were made by, a witness present and testifying in court. All such objections are "
    "reserved and may be interposed at the time of trial.\n\n"
    "This Responding Party has not completed its investigation of the facts relating to this action, "
    "has not yet completed its discovery in this action, and has not yet completed preparation for trial. "
    "Consequently, the following responses are given without prejudice to this party's right to allege "
    "and/or produce evidence of any subsequently-discovered facts or circumstances.\n\n"
    "Except for facts explicitly admitted herein, no admission of any nature is to be implied or inferred. "
    "The fact that any demand or request herein has been answered should not be taken as an admission, or "
    "a confusion of the existence of any facts set forth or assumed by such demand or that such response "
    "constitutes any fact thus set forth or assumed. All responses are given on the basis of a good faith "
    "effort to locate the requested information.\n\n"
    "This party relies on well-established California authority to the effect that demands and requests "
    "cannot be unilaterally designated as continuing in nature, and serves notice that we will not "
    "voluntarily provide further responses if additional information is acquired by us after these "
    "responses are served.\n\n"
    "Notwithstanding the above, this Responding Party reserves the right to change any and all responses "
    "herein as additional facts and further information is obtained, new analyses are made, and legal "
    "research is completed.\n\n"
    "The information contained herein is given in a good faith effort to supply as much factual material "
    "as is presently known by Responding Party, but should in no way prejudice this responding party's "
    "right to make new contentions or provide additional facts or additional information derived from "
    "further discovery, investigation, research and/or legal analysis.\n\n"
    "This preliminary statement shall apply to each and every response given herein, and shall be "
    "incorporated by reference as though fully set forth in all of the demand responses appearing on the "
    "following pages."
)

_DEFAULT_PRELIMINARY_RPD = (
    'The responses set forth below represent the present knowledge of {responding_party} (hereinafter, '
    '"Responding Party") based on discovery, investigation, and case preparation to date. Responding '
    "Party has made reasonable efforts to respond to the Requests, as it understands and interprets each "
    "Request, and the contents hereof are based on the information obtained from these efforts. Responding "
    "Party will make a reasonable effort to gather information responsive to each Request as it understands "
    "and interprets each Request. If Propounding Party subsequently asserts a different interpretation, "
    "Responding Party reserves the right to supplement its objections and/or responses. Responding Party's "
    "investigation into this matter is, and will continue to be, ongoing. Responding Party may locate "
    "additional responsive information or documents at a later date, and it may assert appropriate "
    "objections to the use of the information or documents identified herein. Responding Party, therefore, "
    "expressly reserves the right to modify or supplement these responses and to rely on additional "
    "responsive information or documents, whether located in the course of its continuing investigation "
    "or in the course of discovery, at all future hearings and at trial, and the right to object on "
    "appropriate grounds to the use of any information or documents produced in response to the Requests."
)

_DEFAULT_INTRO_FI = (
    '{responding_party} ("{responding_short}" or "Responding Party") hereby responds and objects to the '
    "{set_word_title} Set of Form Interrogatories served by {propounding_party} "
    '("{propounding_short}" or "Propounding Party"), as follows:'
)

_DEFAULT_INTRO_SI = (
    '{responding_party} ("{responding_short}" or "Responding Party") hereby responds and objects to the '
    "{set_word_title} Set of Special Interrogatories served by {propounding_party} "
    '("{propounding_short}" or "Propounding Party"), as follows:'
)

_DEFAULT_INTRO_RFA = (
    '{responding_party} ("{responding_short}" or "Responding Party") hereby responds and objects to the '
    "{set_word_title} Set of Requests for Admission served by {propounding_party} "
    '("{propounding_short}" or "Propounding Party"), as follows:'
)

_DEFAULT_INTRO_RPD = (
    '{responding_party} ("{responding_short}" or "Responding Party") hereby responds and objects to the '
    "{set_word_title} Request for Production of Documents served by {propounding_party} "
    '("{propounding_short}" or "Propounding Party"), as follows:'
)

_DEFAULT_GENERAL_OBJECTIONS_RFA = (
    "Responding Party objects to each and every Request for Admission to the extent it seeks information "
    "protected by the attorney-client privilege, the attorney work product doctrine, or any other "
    "applicable privilege or protection. Responding Party will not admit, deny, or respond to any Request "
    "that calls for privileged information without an express reservation of all such privileges.\n\n"
    "Responding Party objects to each and every Request to the extent it is vague, ambiguous, overbroad, "
    "or unduly burdensome, and to the extent it calls for information that is neither relevant to the "
    "subject matter of this action nor reasonably calculated to lead to the discovery of admissible "
    "evidence. Each response is made based upon Responding Party's reasonable interpretation of the "
    "Request as stated.\n\n"
    "Responding Party's investigation of this matter and discovery are ongoing and incomplete. Responding "
    "Party reserves the right to amend or supplement its responses as additional information becomes "
    "available. Nothing herein shall be deemed a waiver of any applicable privilege, immunity, or "
    "protection, and Responding Party expressly preserves all such protections."
)

_DEFAULT_GENERAL_OBJECTIONS_RPD = (
    "1. Responding Party objects to the Requests to the extent they seek information or documents "
    "constituting or containing trade secrets or proprietary, confidential, or commercially sensitive "
    "business information, the disclosure of which would be contrary to applicable law or court order.\n\n"
    "2. Responding Party objects to the Requests to the extent they seek information or documents "
    "protected by the attorney-client privilege, the attorney work product doctrine, the joint defense "
    "privilege, the common interest privilege, or any other applicable privilege or immunity.\n\n"
    "3. Responding Party objects to the Requests to the extent they call for legal conclusions or "
    "characterizations.\n\n"
    "4. Responding Party objects to the Requests to the extent they seek information or documents that "
    "are subject to confidentiality agreements or protective orders.\n\n"
    "5. Responding Party objects to the Requests to the extent they are unduly burdensome and "
    "oppressive, and/or the burden or expense of the proposed discovery outweighs its likely benefit.\n\n"
    "6. Responding Party objects to the Requests to the extent they seek information or documents not "
    "relevant to the subject matter of this action and not reasonably calculated to lead to the "
    "discovery of admissible evidence.\n\n"
    "7. Responding Party objects to the Requests to the extent they seek information or documents beyond "
    "the permissible scope of discovery under the California Code of Civil Procedure.\n\n"
    "8. Responding Party objects to the Requests to the extent they are vague, ambiguous, or undefined, "
    "and as such fail to describe the requested documents with reasonable particularity.\n\n"
    "9. Nothing in these responses shall be deemed or construed as an admission by Responding Party of "
    "the truth or accuracy of any allegation in the Requests or of the existence of any documents "
    "described therein.\n\n"
    "10. Responding Party objects to the Requests to the extent they seek information subject to the "
    "right of privacy under the California Constitution, Article I, Section 1, or any other applicable "
    "privacy protection, whether statutory or common law.\n\n"
    "11. Responding Party objects to the Requests to the extent they use definitions or instructions "
    "that are inconsistent with the California Code of Civil Procedure or that impose obligations "
    "beyond those required by law."
)

_DEFAULT_VERIFICATION = (
    "I, {verifier_name}, declare:\n\n"
    "I am a party to this action. I have read the foregoing {document_title} and know the contents "
    "thereof. The same is true of my own knowledge, except as to those matters which are therein stated "
    "on my information and belief, and as to those matters, I believe it to be true.\n\n"
    "I declare under penalty of perjury under the laws of the State of California that the foregoing "
    "is true and correct.\n\n"
    "Executed on _______________, at _________________, California.\n\n"
    "___________________________________\n"
    "{verifier_name}"
)

_DEFAULT_FI_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it is compound, calls for "
    "speculation, and is vague, ambiguous, uncertain and overbroad. Responding Party objects to this "
    "request to the extent this interrogatory violates the attorney-client or work-product privilege."
)

_DEFAULT_FI_7_1_OBJECTIONS = (
    "Responding Party objects on the basis that this request is meant for Plaintiff and not a "
    "defendant as Responding Party is not asserting any claims. Responding Party further objects "
    "on the grounds that this request is not relevant or reasonably calculated to lead to "
    "discoverable evidence. Responding Party also objects on the basis that this request is vague, "
    "ambiguous, and overbroad."
)

_DEFAULT_FI_12_1_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it calls for speculation, "
    "and is compound, vague, ambiguous, uncertain and overbroad. Responding Party specifically "
    "objects to this Interrogatory on the grounds that the terms \"witnessed,\" \"knowledge,\" "
    "and \"statement\" are vague and ambiguous. Responding Party further objects to this "
    "Interrogatory to the extent that it improperly violates the attorney-client privilege and/or "
    "attorney work product doctrines. Responding Party further objects to this Interrogatory on "
    "the grounds that the information sought is equally available to Propounding Party."
)

_DEFAULT_FI_12_2_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it is compound, calls for "
    "speculation, and is vague, ambiguous, uncertain and overbroad. Responding Party specifically "
    "objects to this Interrogatory on the grounds that the term \"interviewed\" is vague and "
    "ambiguous. Responding Party objects to this request to the extent this interrogatory violates "
    "the attorney-client or work-product privilege."
)

_DEFAULT_FI_12_3_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it is compound, calls for "
    "speculation, and is vague, ambiguous, uncertain and overbroad. Responding Party specifically "
    "objects to this Interrogatory on the grounds that the term \"statement\" is vague and ambiguous. "
    "Responding Party objects to this request to the extent this interrogatory violates the "
    "attorney-client or work-product privilege."
)

_DEFAULT_FI_12_6_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it is compound, calls for "
    "speculation, and is vague, ambiguous, uncertain and overbroad. Responding Party objects to "
    "this request to the extent this interrogatory violates the attorney-client or work-product "
    "privilege."
)

_DEFAULT_FI_14_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it is compound, calls for "
    "speculation, and is vague, ambiguous, uncertain and overbroad. Responding Party objects to "
    "this request to the extent this interrogatory violates the attorney-client or work-product "
    "privilege. Responding Party objects to this request to the extent this interrogatory calls "
    "for a legal conclusion and/or expert opinion."
)

_DEFAULT_FI_15_1_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it is vague and ambiguous "
    "as to the term \"material.\" Responding Party further objects to the extent this "
    "Interrogatory invades the attorney-client privilege and work product doctrine. Responding "
    "Party further objects to this Interrogatory on the grounds that it calls for an expert "
    "opinion and a legal conclusion, and seeks the legal reasoning and theories of Responding "
    "Party's contentions. Responding Party is not required to prepare the Propounding Party's "
    "case. Discovery is continuing and Responding Party reserves the right to amend this response "
    "upon discovery of additional facts and information."
)

_DEFAULT_FI_16_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it calls for speculation "
    "and is vague, ambiguous, uncertain, and overbroad. Responding Party further objects to this "
    "Interrogatory on the grounds that it calls for an expert opinion and a legal conclusion, and "
    "seeks the legal reasoning and theories of Responding Party's contentions; Responding Party "
    "is not required to prepare Propounding Party's case. Responding Party further objects to "
    "this Interrogatory on the grounds that it seeks premature disclosure of expert opinion and "
    "violates the attorney work-product privilege. Moreover, pursuant to instruction 2(d) to the "
    "official form interrogatories, the interrogatories in section 16.0 should not be used until "
    "the defendant has had a reasonable opportunity to conduct an investigation or discovery into "
    "plaintiff's damages. At this time, Responding Party has not yet had an opportunity to depose "
    "the Plaintiff, conduct an IME, or obtain all of the Plaintiff's medical records. As such, "
    "Responding Party is not in a position to fully answer this Interrogatory at this time."
)

_DEFAULT_FI_20_OBJECTIONS = (
    "Responding Party objects to this Interrogatory on the grounds that it is compound, calls for "
    "speculation, and is vague, ambiguous, uncertain and overbroad."
)

_DEFAULT_FI_OBJECTIONS_BY_NUMBER = {
    "7.1": _DEFAULT_FI_7_1_OBJECTIONS,
    "12.1": _DEFAULT_FI_12_1_OBJECTIONS,
    "12.2": _DEFAULT_FI_12_2_OBJECTIONS,
    "12.3": _DEFAULT_FI_12_3_OBJECTIONS,
    "12.5": _DEFAULT_FI_OBJECTIONS,
    "12.6": _DEFAULT_FI_12_6_OBJECTIONS,
    "12.7": _DEFAULT_FI_12_6_OBJECTIONS,
    "13.1": _DEFAULT_FI_12_6_OBJECTIONS,
    "13.2": _DEFAULT_FI_12_6_OBJECTIONS,
    "14.1": _DEFAULT_FI_14_OBJECTIONS,
    "14.2": _DEFAULT_FI_14_OBJECTIONS,
    "15.1": _DEFAULT_FI_15_1_OBJECTIONS,
    "16.*": _DEFAULT_FI_16_OBJECTIONS,
    "20.3": _DEFAULT_FI_20_OBJECTIONS,
    "20.4": _DEFAULT_FI_20_OBJECTIONS,
    "20.5": _DEFAULT_FI_20_OBJECTIONS,
    "20.6": _DEFAULT_FI_20_OBJECTIONS,
    "20.7": _DEFAULT_FI_20_OBJECTIONS,
    "20.8": _DEFAULT_FI_20_OBJECTIONS,
    "20.9": _DEFAULT_FI_20_OBJECTIONS,
    "20.10": _DEFAULT_FI_20_OBJECTIONS,
    "20.11": _DEFAULT_FI_20_OBJECTIONS,
}

_MANDATORY_FI_OBJECTIONS_BY_NUMBER = {
    key: _DEFAULT_FI_OBJECTIONS_BY_NUMBER[key]
    for key in (
        "12.1",
        "12.2",
        "12.3",
        "12.6",
        "12.7",
        "13.1",
        "13.2",
        "14.1",
        "14.2",
        "15.1",
        "16.*",
    )
}

_DEFAULT_FI_1_1_RESPONSE = (
    "Responding Party and its attorneys of record, {firm_name}, {firm_address}; {firm_phone}."
)

_DEFAULT_FI_15_1_RESPONSE = (
    "A general denial is interposed as a matter of right based in part on California Code of "
    "Civil Procedure § 431.30. As to affirmative defenses, this interrogatory is premature at "
    "this time."
)

_DEFAULT_FI_16_RESPONSE = (
    "Pursuant to instruction 2(d) to the official form interrogatories, the interrogatories in "
    "section 16.0 should not be used until the defendant has had a reasonable opportunity to conduct "
    "an investigation or discovery into plaintiff's injuries and damages. At this time, responding "
    "party has yet to have an opportunity to depose the Plaintiffs, obtain an IME and obtain all of "
    "the Plaintiffs' medical records. As such, Responding Party is not in a position to fully answer "
    "this interrogatory at this time."
)

_DEFAULT_FI_3_7_RESPONSE = (
    "No, other than the customary licenses necessary to operate a business."
)

_DEFAULT_FI_7_1_RESPONSE = (
    "Not Applicable. Responding Party is not making a claim for damages in this action."
)

_DEFAULT_FI_7_RESPONSE = "Not Applicable."

_DEFAULT_FI_RESPONSES_BY_NUMBER = {
    "3.7": _DEFAULT_FI_3_7_RESPONSE,
    "7.1": _DEFAULT_FI_7_1_RESPONSE,
    "7.2": _DEFAULT_FI_7_RESPONSE,
    "7.3": _DEFAULT_FI_7_RESPONSE,
    "15.1": _DEFAULT_FI_15_1_RESPONSE,
    "16.*": "",
}

_MANDATORY_FI_RESPONSES_BY_NUMBER = dict(_DEFAULT_FI_RESPONSES_BY_NUMBER)


# ---------------------------------------------------------------------------
# ResponseRules dataclass
# ---------------------------------------------------------------------------

@dataclass
class ResponseRules:
    """
    Strategy settings and boilerplate defaults for generating discovery responses.

    Instances are serializable to/from JSON and support partial initialization
    so that user-saved preferences override only the fields they specify.
    """

    # ------------------------------------------------------------------
    # Objection strategy
    # ------------------------------------------------------------------
    objection_aggressiveness: str = "aggressive"
    """One of: 'aggressive', 'moderate', 'conservative'."""

    always_include_privacy_objection: bool = True
    always_include_privilege_objection: bool = True
    always_include_burden_objection: bool = True
    auto_flag_compound: bool = True
    auto_flag_broad_definitions: bool = True

    # ------------------------------------------------------------------
    # Response strategy
    # ------------------------------------------------------------------
    si_response_style: str = "minimal"
    """One of: 'minimal', 'narrative', 'detailed'."""

    rfa_default_posture: str = "cautious"
    """One of: 'cautious', 'cooperative', 'deny_all'."""

    rpd_default_posture: str = "context_dependent"
    """One of: 'context_dependent', 'produce_all', 'withhold_pending_protective'."""

    fi_17_1_auto_refresh: bool = False
    """If True, automatically regenerate FI 17.1 when underlying responses change."""

    # ------------------------------------------------------------------
    # Boilerplate text
    # ------------------------------------------------------------------
    waiver_language: str = field(default_factory=lambda: _DEFAULT_WAIVER)
    reservation_clause: str = field(default_factory=lambda: _DEFAULT_RESERVATION)

    preliminary_statement_fi: str = field(default_factory=lambda: _DEFAULT_PRELIMINARY_FI)
    preliminary_statement_si: str = field(default_factory=lambda: _DEFAULT_PRELIMINARY_SI)
    preliminary_statement_rfa: str = field(default_factory=lambda: _DEFAULT_PRELIMINARY_RFA)
    preliminary_statement_rpd: str = field(default_factory=lambda: _DEFAULT_PRELIMINARY_RPD)

    intro_template_fi: str = field(default_factory=lambda: _DEFAULT_INTRO_FI)
    intro_template_si: str = field(default_factory=lambda: _DEFAULT_INTRO_SI)
    intro_template_rfa: str = field(default_factory=lambda: _DEFAULT_INTRO_RFA)
    intro_template_rpd: str = field(default_factory=lambda: _DEFAULT_INTRO_RPD)

    general_objections_rfa: str = field(default_factory=lambda: _DEFAULT_GENERAL_OBJECTIONS_RFA)
    general_objections_rpd: str = field(default_factory=lambda: _DEFAULT_GENERAL_OBJECTIONS_RPD)

    verification_template: str = field(default_factory=lambda: _DEFAULT_VERIFICATION)

    # ------------------------------------------------------------------
    # Fixed FI responses
    # ------------------------------------------------------------------
    fi_objections: str = field(default_factory=lambda: _DEFAULT_FI_OBJECTIONS)
    fi_objections_by_number: Dict[str, str] = field(
        default_factory=lambda: dict(_DEFAULT_FI_OBJECTIONS_BY_NUMBER)
    )
    fi_responses_by_number: Dict[str, str] = field(
        default_factory=lambda: dict(_DEFAULT_FI_RESPONSES_BY_NUMBER)
    )
    fi_1_1_response: str = field(default_factory=lambda: _DEFAULT_FI_1_1_RESPONSE)
    fi_15_1_response: str = field(default_factory=lambda: _DEFAULT_FI_15_1_RESPONSE)
    fi_16_response: str = field(default_factory=lambda: _DEFAULT_FI_16_RESPONSE)

    # ------------------------------------------------------------------
    # User-defined extensions
    # ------------------------------------------------------------------
    custom_instructions: str = ""
    """Free-text instructions appended to LLM prompts when drafting responses."""

    # ------------------------------------------------------------------
    # Serialization helpers
    # ------------------------------------------------------------------

    def to_dict(self) -> Dict[str, Any]:
        """Return all fields as a plain dictionary (JSON-serializable)."""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "ResponseRules":
        """
        Create a ResponseRules from a (possibly partial) dictionary.

        Fields not present in *data* keep their dataclass defaults.
        Unknown keys in *data* are silently ignored.
        """
        defaults = cls()
        valid_keys = {f.name for f in fields(cls)}
        merged = defaults.to_dict()
        for key, value in data.items():
            if key in valid_keys:
                merged[key] = value
        merged["fi_objections_by_number"] = dict(
            merged.get("fi_objections_by_number") or {}
        )
        merged["fi_objections_by_number"].update(_MANDATORY_FI_OBJECTIONS_BY_NUMBER)
        merged["fi_responses_by_number"] = dict(
            merged.get("fi_responses_by_number") or {}
        )
        merged["fi_responses_by_number"].update(_MANDATORY_FI_RESPONSES_BY_NUMBER)
        # Upgrade superseded firm-standard reservation wording persisted in older
        # case data to the current default. Genuine customizations are preserved.
        if (merged.get("reservation_clause") or "").strip() in _SUPERSEDED_RESERVATIONS:
            merged["reservation_clause"] = _DEFAULT_RESERVATION
        return cls(**merged)

    def save_to_json(self, path: str) -> None:
        """Serialize this instance to a JSON file at *path*."""
        os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(self.to_dict(), fh, indent=2, ensure_ascii=False)

    @classmethod
    def load_from_json(cls, path: str) -> "ResponseRules":
        """
        Load a ResponseRules from a JSON file.

        Returns default ResponseRules if the file does not exist.
        """
        if not os.path.exists(path):
            return cls()
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
        return cls.from_dict(data)
