import unittest

from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_review_state import (
    RequestReview,
    ReviewState,
    insert_quick_objection,
    remove_quick_objection,
)
from icharlotte_core.discovery.response_rule_library import get_quick_objection_rules
from icharlotte_core.discovery.response_rules import ResponseRules


class ResponseReviewStateTests(unittest.TestCase):
    def test_insert_quick_objection_adds_text_once(self):
        review = RequestReview(number="1", request_text="Identify witnesses.")
        quick = next(r for r in get_quick_objection_rules() if r.id == "quick_vague")

        insert_quick_objection(review, quick)
        insert_quick_objection(review, quick)

        self.assertEqual(review.proposed_objections.count(quick.output_text), 1)
        self.assertEqual(review.selected_quick_objection_ids, ["quick_vague"])

    def test_insert_quick_objection_substitutes_term(self):
        review = RequestReview(number="1", request_text="Define INCIDENT.")
        quick = next(
            r for r in get_quick_objection_rules() if r.id == "quick_undefined_term"
        )

        insert_quick_objection(review, quick, term="INCIDENT")

        self.assertIn('"INCIDENT"', review.proposed_objections)
        self.assertNotIn("{term}", review.proposed_objections)

    def test_remove_quick_objection_removes_matching_text(self):
        review = RequestReview(number="1", request_text="Identify witnesses.")
        quick = next(r for r in get_quick_objection_rules() if r.id == "quick_vague")
        insert_quick_objection(review, quick)

        remove_quick_objection(review, quick)

        self.assertEqual(review.proposed_objections, "")
        self.assertEqual(review.selected_quick_objection_ids, [])

    def test_review_state_round_trips(self):
        state = ReviewState(
            requests=[
                RequestReview(
                    number="1",
                    request_text="Q",
                    proposed_objections="O",
                    proposed_substantive_response="R",
                    selected_rule_ids=["always_vague_ambiguous_overbroad"],
                    approved=True,
                )
            ]
        )

        loaded = ReviewState.from_dict(state.to_dict())

        self.assertEqual(loaded.requests, state.requests)
        self.assertTrue(loaded.all_approved())

    def test_review_warning_fields_round_trip(self):
        state = ReviewState(
            requests=[
                RequestReview(
                    number="1",
                    request_text="Identify witnesses.",
                    proposed_substantive_response="Unknown at this time.",
                    needs_review=True,
                    review_reason="No specific context found.",
                )
            ]
        )

        loaded = ReviewState.from_dict(state.to_dict())

        self.assertTrue(loaded.requests[0].needs_review)
        self.assertEqual(
            loaded.requests[0].review_reason,
            "No specific context found.",
        )

    def test_review_warning_fields_are_optional_for_old_state(self):
        loaded = RequestReview.from_dict(
            {
                "number": "1",
                "request_text": "Identify witnesses.",
                "proposed_objections": "",
                "proposed_substantive_response": "",
                "selected_rule_ids": [],
                "selected_quick_objection_ids": [],
                "approved": False,
            }
        )

        self.assertFalse(loaded.needs_review)
        self.assertEqual(loaded.review_reason, "")

    def test_to_plain_text_uses_current_review_values(self):
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[ParsedRequest(number="1", text="Identify witnesses.")],
        )
        state = ReviewState(
            requests=[
                RequestReview(
                    number="1",
                    request_text="Identify witnesses.",
                    proposed_objections="Objection.",
                    proposed_substantive_response="No witnesses known.",
                    approved=True,
                )
            ]
        )

        text = state.to_plain_text(parsed, ResponseRules())

        self.assertIn("SPECIAL INTERROGATORY NO. 1:", text)
        self.assertIn("Objection.", text)
        self.assertIn("No witnesses known.", text)


if __name__ == "__main__":
    unittest.main()
