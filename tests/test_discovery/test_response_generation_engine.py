import unittest

from icharlotte_core.discovery.response_generation_engine import (
    ConditionalRuleDecision,
    DraftCallbacks,
    generate_review_state,
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

    def test_structured_proposal_ignores_unknown_rule_ids(self):
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

        self.assertEqual(review.proposed_objections, "")
        self.assertTrue(review.needs_review)
        self.assertIn("Unknown rule ID", review.review_reason)

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


if __name__ == "__main__":
    unittest.main()
