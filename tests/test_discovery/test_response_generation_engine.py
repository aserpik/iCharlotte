import unittest

from icharlotte_core.discovery.response_generation_engine import (
    ConditionalRuleDecision,
    DraftCallbacks,
    build_structured_proposal_prompt,
    generate_review_state,
    parse_structured_proposal_response,
)
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_rule_library import (
    RuleCategory,
    built_in_rules_for,
)
from icharlotte_core.discovery.response_rules import ResponseRules


def _parsed(discovery_type="SI"):
    return ParsedDiscovery(
        discovery_type=discovery_type,
        propounding_party="Plaintiff Smith",
        responding_party="Defendant Jones",
        set_number=1,
        set_word="ONE",
        case_number="123",
        requests=[
            ParsedRequest(number="1", text="Identify all witnesses."),
            ParsedRequest(
                number="2",
                text="State all facts regarding INCIDENT.",
                defined_terms_used=["INCIDENT"],
            ),
        ],
    )


class ResponseGenerationEngineTests(unittest.TestCase):
    def test_mandatory_vague_rule_applies_to_every_request(self):
        rules = [
            r for r in built_in_rules_for("SI")
            if r.id == "always_vague_ambiguous_overbroad"
        ]

        state = generate_review_state(_parsed(), rules, context_text="")

        self.assertEqual(len(state.requests), 2)
        for review in state.requests:
            self.assertIn("vague, ambiguous", review.proposed_objections)
            self.assertEqual(
                review.selected_rule_ids,
                ["always_vague_ambiguous_overbroad"],
            )

    def test_conditional_rule_uses_defined_term(self):
        rules = [
            r for r in built_in_rules_for("SI")
            if r.id == "ambiguous_term_when_unclear"
        ]

        state = generate_review_state(_parsed(), rules, context_text="")

        self.assertEqual(state.requests[0].proposed_objections, "")
        self.assertIn('"INCIDENT"', state.requests[1].proposed_objections)
        self.assertIn("ambiguous_term_when_unclear", state.requests[1].selected_rule_ids)

    def test_callback_can_control_conditional_rule_and_substantive_text(self):
        all_rules = built_in_rules_for("SI")
        conditional = [
            r for r in all_rules
            if r.id == "privilege_when_potentially_privileged"
        ]
        instructions = [r for r in all_rules if r.category == RuleCategory.SUBSTANTIVE]

        def decide(rule, request, parsed, context_text):
            return ConditionalRuleDecision(applies=request.number == "1")

        def draft(request, parsed, context_text, selected_rules):
            instruction_names = [r.name for r in selected_rules if r.category == RuleCategory.SUBSTANTIVE]
            return f"Answer {request.number}. {instruction_names[0]}"

        state = generate_review_state(
            _parsed(),
            conditional + instructions,
            context_text="case facts",
            callbacks=DraftCallbacks(
                should_apply_rule=decide,
                draft_substantive=draft,
            ),
        )

        self.assertIn("privilege", state.requests[0].proposed_objections.lower())
        self.assertEqual(state.requests[1].proposed_objections, "")
        self.assertIn("Answer 1.", state.requests[0].proposed_substantive_response)
        self.assertIn("as few words as possible", state.requests[0].proposed_substantive_response)

    def test_fi_fixed_mode_uses_number_specific_objections_and_responses(self):
        parsed = _parsed("FI")
        parsed.requests = [
            ParsedRequest(number="1.1", text="State your name."),
            ParsedRequest(number="15.1", text="Identify denials and defenses."),
            ParsedRequest(number="17.1", text="Is your response to each RFA unqualified?"),
        ]
        response_rules = ResponseRules(fi_1_1_response="Test Firm 555")

        state = generate_review_state(
            parsed,
            selected_rules=[],
            context_text="",
            response_rules=response_rules,
            fi_mode="fixed",
        )

        self.assertEqual(state.requests[0].proposed_objections, "")
        self.assertIn("Test Firm", state.requests[0].proposed_substantive_response)
        self.assertIn("vague and ambiguous", state.requests[1].proposed_objections)
        self.assertIn("general denial", state.requests[1].proposed_substantive_response.lower())
        self.assertIn("PENDING", state.requests[2].proposed_substantive_response)

    def test_fi_fixed_mode_splits_16_series_objections_from_response(self):
        parsed = _parsed("FI")
        parsed.requests = [
            ParsedRequest(number="16.9", text="Do you contend damages were not caused?"),
        ]

        state = generate_review_state(
            parsed,
            selected_rules=[],
            context_text="",
            response_rules=ResponseRules(),
            fi_mode="fixed",
        )

        review = state.requests[0]
        self.assertIn("calls for an expert opinion", review.proposed_objections)
        self.assertIn("section 16.0 should not be used", review.proposed_objections)
        self.assertEqual(review.proposed_substantive_response, "")

    def test_fi_fixed_mode_uses_requested_fixed_rules_over_stale_saved_rules(self):
        parsed = _parsed("FI")
        parsed.requests = [
            ParsedRequest(number="3.7", text="Within the past five years..."),
            ParsedRequest(number="7.1", text="Do you attribute damages to incident?"),
            ParsedRequest(number="7.2", text="List each injury."),
            ParsedRequest(number="7.3", text="Do you still have complaints?"),
            ParsedRequest(number="12.1", text="State each witness."),
            ParsedRequest(number="15.1", text="Identify denials and defenses."),
            ParsedRequest(number="16.1", text="Do you contend injuries were not caused?"),
        ]
        response_rules = ResponseRules.from_dict(
            {
                "fi_15_1_response": "Old saved 15.1 response.",
                "fi_16_response": "Old saved 16 response.",
                "fi_objections_by_number": {
                    "12.1": "Old saved 12.1 objection.",
                    "15.1": "Old saved 15.1 objection.",
                    "16.*": "Old saved 16 objection.",
                },
            }
        )

        state = generate_review_state(
            parsed,
            selected_rules=[],
            context_text="",
            response_rules=response_rules,
            fi_mode="fixed",
        )
        reviews = {item.number: item for item in state.requests}

        self.assertIn("customary licenses necessary", reviews["3.7"].proposed_substantive_response)
        self.assertEqual(
            reviews["7.1"].proposed_substantive_response,
            "Not Applicable. Responding Party is not making a claim for damages in this action.",
        )
        self.assertEqual(reviews["7.2"].proposed_substantive_response, "Not Applicable.")
        self.assertEqual(reviews["7.3"].proposed_substantive_response, "Not Applicable.")
        self.assertIn('"witnessed," "knowledge," and "statement"', reviews["12.1"].proposed_objections)
        self.assertNotIn("Old saved 12.1", reviews["12.1"].proposed_objections)
        self.assertEqual(
            reviews["15.1"].proposed_substantive_response,
            (
                "A general denial is interposed as a matter of right based in part on "
                "California Code of Civil Procedure § 431.30. As to affirmative defenses, "
                "this interrogatory is premature at this time."
            ),
        )
        self.assertIn("vague and ambiguous as to the term", reviews["15.1"].proposed_objections)
        self.assertIn("pursuant to instruction 2(d)", reviews["16.1"].proposed_objections)
        self.assertEqual(reviews["16.1"].proposed_substantive_response, "")

    def test_structured_proposal_ignores_model_objection_text(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
            apply_structured_proposal,
        )

        parsed = _parsed()
        rules = built_in_rules_for("SI")
        proposal = StructuredProposal(
            request_number="2",
            conditional_objection_rule_ids=["ambiguous_term_when_unclear"],
            applied_custom_rule_ids=[],
            applied_instruction_rule_ids=["minimal_direct_answer"],
            ambiguous_term="INCIDENT",
            proposed_objections="MODEL SHOULD NOT CONTROL OBJECTION TEXT",
            proposed_substantive_response="No additional facts known.",
            needs_review=True,
            review_reason="No specific context found.",
        )

        review = apply_structured_proposal(parsed.requests[1], parsed, rules, proposal)

        self.assertIn('"INCIDENT"', review.proposed_objections)
        self.assertNotIn("MODEL SHOULD NOT CONTROL", review.proposed_objections)
        self.assertEqual(review.proposed_substantive_response, "No additional facts known.")
        self.assertTrue(review.needs_review)
        self.assertEqual(review.review_reason, "No specific context found.")

    def test_structured_proposal_unknown_rule_ids_preserve_mandatory_rules(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
            apply_structured_proposal,
        )

        parsed = _parsed()
        rules = built_in_rules_for("SI")
        proposal = StructuredProposal(
            request_number="1",
            conditional_objection_rule_ids=["missing_rule"],
            applied_custom_rule_ids=[],
            applied_instruction_rule_ids=[],
            proposed_substantive_response="Unknown.",
        )

        review = apply_structured_proposal(parsed.requests[0], parsed, rules, proposal)

        self.assertIn("vague, ambiguous", review.proposed_objections)
        self.assertIn("always_vague_ambiguous_overbroad", review.selected_rule_ids)
        self.assertNotIn("missing_rule", review.selected_rule_ids)
        self.assertTrue(review.needs_review)
        self.assertIn("Unknown rule ID", review.review_reason)

    def test_structured_proposal_rejects_ungrounded_ambiguous_term(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
            apply_structured_proposal,
        )

        parsed = _parsed()
        rules = built_in_rules_for("SI")
        proposal = StructuredProposal(
            request_number="2",
            conditional_objection_rule_ids=["ambiguous_term_when_unclear"],
            ambiguous_term="BAD TERM WITH EXTRA OBJECTION",
            proposed_substantive_response="No additional facts known.",
        )

        review = apply_structured_proposal(parsed.requests[1], parsed, rules, proposal)

        self.assertNotIn("BAD TERM WITH EXTRA OBJECTION", review.proposed_objections)
        self.assertIn('"INCIDENT"', review.proposed_objections)

    def test_structured_proposal_request_number_mismatch_needs_review(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
            apply_structured_proposal,
        )

        parsed = _parsed()
        rules = built_in_rules_for("SI")
        proposal = StructuredProposal(
            request_number="2",
            conditional_objection_rule_ids=[],
            applied_instruction_rule_ids=["minimal_direct_answer"],
            proposed_substantive_response="No additional facts known.",
        )

        review = apply_structured_proposal(parsed.requests[0], parsed, rules, proposal)

        self.assertTrue(review.needs_review)
        self.assertIn("request number mismatch", review.review_reason.lower())
        self.assertEqual(review.proposed_substantive_response, "No additional facts known.")

    def test_parse_structured_proposal_json_extracts_schema(self):
        from icharlotte_core.discovery.response_generation_engine import (
            parse_structured_proposal_response,
        )

        proposal = parse_structured_proposal_response(
            """
            ```json
            {
              "request_number": "1",
              "conditional_objection_rule_ids": ["privilege_when_potentially_privileged"],
              "applied_custom_rule_ids": [],
              "applied_instruction_rule_ids": ["minimal_direct_answer"],
              "ambiguous_term": "",
              "proposed_objections": "ignored",
              "proposed_substantive_response": "No privileged documents will be produced.",
              "needs_review": true,
              "review_reason": "Possible privilege issue."
            }
            ```
            """
        )

        self.assertEqual(proposal.request_number, "1")
        self.assertEqual(
            proposal.conditional_objection_rule_ids,
            ["privilege_when_potentially_privileged"],
        )
        self.assertTrue(proposal.needs_review)
        self.assertEqual(proposal.review_reason, "Possible privilege issue.")

    def test_generate_review_state_can_use_structured_proposal_callback(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
            generate_review_state,
        )

        rules = built_in_rules_for("SI")

        def propose(request, parsed, context_text, selected_rules, response_rules):
            return StructuredProposal(
                request_number=request.number,
                conditional_objection_rule_ids=["privilege_when_potentially_privileged"]
                if request.number == "1"
                else [],
                applied_instruction_rule_ids=["minimal_direct_answer"],
                proposed_substantive_response=f"Response for {request.number}.",
                needs_review=request.number == "2",
                review_reason="No specific context found." if request.number == "2" else "",
            )

        state = generate_review_state(
            _parsed(),
            rules,
            context_text="case facts",
            callbacks=DraftCallbacks(structured_proposal=propose),
        )

        self.assertIn("privilege", state.requests[0].proposed_objections.lower())
        self.assertEqual(state.requests[0].proposed_substantive_response, "Response for 1.")
        self.assertTrue(state.requests[1].needs_review)
        self.assertEqual(state.requests[1].review_reason, "No specific context found.")

    def test_build_structured_proposal_prompt_contains_schema_and_packet(self):
        from icharlotte_core.discovery.response_generation_engine import (
            build_structured_proposal_prompt,
        )

        parsed = _parsed("RFA")
        prompt = build_structured_proposal_prompt(
            parsed.requests[0],
            parsed,
            context_packet="[status.txt #1]\nNo witnesses identified.",
            selected_rules=built_in_rules_for("RFA"),
            response_rules=ResponseRules(),
        )

        self.assertIn("Return ONLY a JSON object", prompt)
        self.assertIn("conditional_objection_rule_ids", prompt)
        self.assertIn("[status.txt #1]", prompt)
        self.assertIn("admit only when", prompt.lower())

    def test_fallback_proposal_marks_empty_context_for_review(self):
        from icharlotte_core.discovery.response_generation_engine import (
            build_fallback_structured_proposal,
        )

        parsed = _parsed("SI")
        proposal = build_fallback_structured_proposal(
            parsed.requests[0],
            parsed,
            context_packet="",
        )

        self.assertEqual(proposal.request_number, "1")
        self.assertTrue(proposal.needs_review)
        self.assertIn("No specific context found", proposal.review_reason)

    def test_rfa_fallback_uses_insufficient_information(self):
        from icharlotte_core.discovery.response_generation_engine import (
            build_fallback_structured_proposal,
        )

        parsed = _parsed("RFA")
        proposal = build_fallback_structured_proposal(
            parsed.requests[0],
            parsed,
            context_packet="",
        )

        self.assertIn("insufficient to enable Responding Party to admit", proposal.proposed_substantive_response)
        self.assertTrue(proposal.needs_review)

    def test_rpd_fallback_uses_unable_to_comply_when_context_empty(self):
        from icharlotte_core.discovery.response_generation_engine import (
            build_fallback_structured_proposal,
        )

        parsed = _parsed("RPD")
        proposal = build_fallback_structured_proposal(
            parsed.requests[0],
            parsed,
            context_packet="",
        )

        self.assertIn("unable to comply", proposal.proposed_substantive_response.lower())
        self.assertTrue(proposal.needs_review)


class StructuredProposalPromptTests(unittest.TestCase):
    def _build_prompt(self):
        parsed = _parsed()
        request = parsed.requests[0]
        rules = built_in_rules_for("SI")
        return build_structured_proposal_prompt(
            request=request,
            parsed=parsed,
            context_packet="",
            selected_rules=rules,
            response_rules=ResponseRules(),
        )

    def test_prompt_omits_proposed_objections_from_schema(self):
        prompt = self._build_prompt()
        # The dead "proposed_objections" field used to live inside the JSON
        # schema block. We still mention objection RULES in the schema,
        # but the OBJECTIONS-TEXT field is gone.
        self.assertNotIn('"proposed_objections"', prompt)

    def test_prompt_explicitly_forbids_drafting_objection_text(self):
        prompt = self._build_prompt()
        self.assertIn("Do not draft objection text.", prompt)

    def test_parse_handles_payload_without_proposed_objections(self):
        # Old clients won't send the field. The parser must tolerate that.
        proposal = parse_structured_proposal_response(
            '{"request_number": "1", "proposed_substantive_response": "ok"}'
        )
        self.assertEqual(proposal.request_number, "1")
        self.assertEqual(proposal.proposed_substantive_response, "ok")
        self.assertEqual(proposal.proposed_objections, "")

    def test_parse_still_tolerates_legacy_proposed_objections_field(self):
        # Old persisted JSON may still carry it; we just ignore the value.
        proposal = parse_structured_proposal_response(
            '{"request_number": "1", "proposed_objections": "legacy", '
            '"proposed_substantive_response": "ok"}'
        )
        self.assertEqual(proposal.proposed_objections, "legacy")


class PendingReviewStateTests(unittest.TestCase):
    def _fi_parsed(self):
        return ParsedDiscovery(
            discovery_type="FI",
            propounding_party="P",
            responding_party="D",
            set_number=1,
            set_word="ONE",
            case_number="1",
            requests=[
                ParsedRequest(number="A", text="Applicable, no fixed response."),
                ParsedRequest(number="F", text="Has a fixed response."),
                ParsedRequest(number="N", text="Inapplicable form interrogatory."),
            ],
        )

    def test_pending_state_marks_non_skipped_requests_as_pending(self):
        from icharlotte_core.discovery.response_generation_engine import (
            _build_pending_review_state,
        )
        from icharlotte_core.discovery.response_rules import ResponseRules

        parsed = _parsed()  # SI, two requests
        state = _build_pending_review_state(parsed, ResponseRules(), "custom")
        self.assertEqual(len(state.requests), 2)
        for review in state.requests:
            self.assertTrue(review.is_pending)
            self.assertEqual(review.proposed_substantive_response, "")
            self.assertEqual(review.proposed_objections, "")
            self.assertEqual(review.review_reason, "Generating...")

    def test_pending_state_fills_fixed_fi_responses_immediately(self):
        from icharlotte_core.discovery import response_generation_engine as eng
        from icharlotte_core.discovery.response_rules import ResponseRules

        # Mock both authoritative lookups so the test fully controls the
        # branches without depending on which numbers happen to be in
        # response_drafter's default _INAPPLICABLE_ or fi_responses_by_number.
        rules = ResponseRules()
        # Avoid the FI rules dict from steering the lookup -- leave default.

        original_detect = eng.detect_inapplicable_fi if hasattr(
            eng, "detect_inapplicable_fi"
        ) else None
        original_fixed = eng.get_fi_fixed_response if hasattr(
            eng, "get_fi_fixed_response"
        ) else None

        # These helpers are imported function-locally inside
        # _build_pending_review_state, so we patch the module they come
        # from rather than this engine module.
        from icharlotte_core.discovery import response_drafter as rd
        original_rd_detect = rd.detect_inapplicable_fi
        original_rd_fixed = rd.get_fi_fixed_response
        original_rd_objections = rd.get_fi_fixed_objections
        original_rd_strip = rd.strip_fi_objections_from_fixed_response

        rd.detect_inapplicable_fi = lambda number: number == "N"
        rd.get_fi_fixed_response = lambda number, _rules: (
            "Fixed answer body." if number == "F" else None
        )
        rd.get_fi_fixed_objections = lambda number, _rules: ""
        rd.strip_fi_objections_from_fixed_response = lambda _num, text, _rules: text

        try:
            parsed = self._fi_parsed()
            state = eng._build_pending_review_state(parsed, rules, "fixed")
        finally:
            rd.detect_inapplicable_fi = original_rd_detect
            rd.get_fi_fixed_response = original_rd_fixed
            rd.get_fi_fixed_objections = original_rd_objections
            rd.strip_fi_objections_from_fixed_response = original_rd_strip

        rows = {r.number: r for r in state.requests}
        # F has a fixed response -- not pending, body filled.
        self.assertFalse(rows["F"].is_pending)
        self.assertEqual(rows["F"].proposed_substantive_response, "Fixed answer body.")
        # N is inapplicable -- not pending, gets the inapplicable response.
        self.assertFalse(rows["N"].is_pending)
        self.assertIn("not applicable", rows["N"].proposed_substantive_response.lower())
        # A is applicable with no fixed response -- pending.
        self.assertTrue(rows["A"].is_pending)
        self.assertEqual(rows["A"].review_reason, "Generating...")
    def test_apply_proposal_overwrites_pending_row(self):
        from icharlotte_core.discovery.response_generation_engine import (
            _apply_proposal_to_review_state,
            _build_pending_review_state,
            StructuredProposal,
        )
        from icharlotte_core.discovery.response_rules import ResponseRules

        parsed = _parsed()
        state = _build_pending_review_state(parsed, ResponseRules(), "custom")
        proposal = StructuredProposal(
            request_number="1",
            proposed_substantive_response="Drafted answer.",
        )

        review = _apply_proposal_to_review_state(
            review_state=state,
            req_number="1",
            proposal=proposal,
            parsed=parsed,
            selected_rules=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )

        self.assertIsNotNone(review)
        self.assertFalse(review.is_pending)
        self.assertEqual(review.proposed_substantive_response, "Drafted answer.")
        # The state object itself reflects the change.
        self.assertFalse(state.requests[0].is_pending)

    def test_apply_proposal_returns_none_for_unknown_request_number(self):
        from icharlotte_core.discovery.response_generation_engine import (
            _apply_proposal_to_review_state,
            _build_pending_review_state,
            StructuredProposal,
        )
        from icharlotte_core.discovery.response_rules import ResponseRules

        parsed = _parsed()
        state = _build_pending_review_state(parsed, ResponseRules(), "custom")
        result = _apply_proposal_to_review_state(
            review_state=state,
            req_number="999",
            proposal=StructuredProposal(request_number="999"),
            parsed=parsed,
            selected_rules=[],
            response_rules=ResponseRules(),
            fi_mode="custom",
        )
        self.assertIsNone(result)


if __name__ == "__main__":
    unittest.main()
