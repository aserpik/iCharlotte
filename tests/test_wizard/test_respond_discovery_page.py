import unittest
import tempfile
from dataclasses import asdict
from pathlib import Path
from unittest.mock import patch

import pytest

pytest.importorskip("pytestqt")
from PySide6.QtWidgets import QApplication

from icharlotte_core.discovery.response_generation_engine import StructuredProposal
from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
from icharlotte_core.discovery.response_review_state import RequestReview, ReviewState
from icharlotte_core.discovery.response_rules import ResponseRules
import icharlotte_core.ui.wizard.pages.respond_discovery_page as respond_discovery_page
from icharlotte_core.ui.wizard.pages.respond_discovery_page import (
    RespondDiscoverySettingsPage,
    RespondDiscoveryWorker,
    _build_structured_proposal_map,
    _draft_substantive_response_map,
    _normalize_and_filter_parsed_discovery,
    load_respond_response_rules,
)


class RespondDiscoverySettingsPageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def test_si_rules_show_user_defined_labels(self):
        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
        )

        labels = page.visible_rule_labels()

        self.assertIn(
            'ALWAYS include "vague, ambiguous, overbroad" objections',
            labels,
        )
        self.assertTrue(page.has_substantive_rules())

    def test_rfa_rules_do_not_show_substantive_rules(self):
        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\rfa.pdf",
            detected_type="RFA",
        )

        self.assertGreater(page.visible_rule_count(), 0)
        self.assertFalse(page.has_substantive_rules())

    def test_fi_fixed_hides_rule_cards_until_custom_selected(self):
        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\frogg.pdf",
            detected_type="FI",
        )

        self.assertEqual(page.visible_rule_count(), 0)
        page.set_fi_mode("custom")
        self.assertGreater(page.visible_rule_count(), 0)
        self.assertTrue(page.has_substantive_rules())

    def test_restored_fi_mode_keeps_radio_buttons_in_sync(self):
        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\frogg.pdf",
            detected_type="FI",
        )

        page.from_dict({"fi_mode": "custom"})

        self.assertTrue(page.rb_fi_custom.isChecked())
        page.rb_fi_fixed.click()
        self.assertEqual(page.fi_mode, "fixed")

    @patch("icharlotte_core.ui.wizard.pages.respond_discovery_page.load_respond_response_rules")
    def test_fixed_radio_after_restore_generates_fixed_fi_review(self, mock_load_rules):
        mock_load_rules.return_value.fi_objections_by_number = {
            "15.1": "Saved 15.1 objections."
        }
        mock_load_rules.return_value.fi_15_1_response = "Saved 15.1 response."
        parsed = ParsedDiscovery(
            discovery_type="FI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[ParsedRequest(number="15.1", text="Identify denials and defenses.")],
        )
        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\frogg.pdf",
            detected_type="FI",
            parsed_discovery=parsed,
        )
        page.from_dict({"fi_mode": "custom", "parsed_discovery": asdict(parsed)})

        page.rb_fi_fixed.click()
        page._generate_review_from_parsed()

        review = page.review_state.requests[0]
        self.assertEqual(review.proposed_objections, "Saved 15.1 objections.")
        self.assertEqual(review.proposed_substantive_response, "Saved 15.1 response.")

    def test_to_dict_includes_review_state_when_loaded(self):
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
            [RequestReview(number="1", request_text="Identify witnesses.", approved=True)]
        )
        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
            parsed_discovery=parsed,
            review_state=state,
        )

        data = page.to_dict()

        self.assertEqual(data["detected_type"], "SI")
        self.assertEqual(data["parsed_discovery"]["requests"][0]["number"], "1")
        self.assertTrue(data["review_state"]["requests"][0]["approved"])

    def test_to_dict_includes_save_as_defaults(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            parsed = ParsedDiscovery(
                discovery_type="SI",
                propounding_party="Plaintiff Smith",
                responding_party="Defendant Jones",
                set_number=1,
                set_word="ONE",
                case_number="123",
                requests=[ParsedRequest(number="1", text="Identify witnesses.")],
            )
            page = RespondDiscoverySettingsPage(
                case_root=tmp,
                file_number="3000.075",
                discovery_file=str(Path(tmp) / "DISCOVERY" / "PROPOUNDED" / "srogg.pdf"),
                detected_type="SI",
                parsed_discovery=parsed,
                review_state=ReviewState(
                    [RequestReview(number="1", request_text="Identify witnesses.", approved=True)]
                ),
            )

            data = page.to_dict()

            self.assertEqual(
                data["save_default_dir"],
                str(Path(tmp) / "DISCOVERY" / "RESPONSES"),
            )
            self.assertEqual(
                data["suggested_filename"],
                "Def Jones's Resp to SI(1).docx",
            )

    @patch(
        "icharlotte_core.ui.wizard.pages.respond_discovery_page.extract_selected_form_interrogatory_numbers",
        return_value=["1.1", "3.1"],
    )
    def test_filters_fi_parse_to_checked_boxes(self, _mock_selected):
        parsed = ParsedDiscovery(
            discovery_type="Form Interrogatories",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[
                ParsedRequest(number="1.1", text="State who answered."),
                ParsedRequest(number="3.1", text="Are you a corporation?"),
                ParsedRequest(number="6.1", text="Do you claim injuries?"),
            ],
        )

        filtered = _normalize_and_filter_parsed_discovery(
            parsed,
            detected_type="FI",
            discovery_file=r"C:\case\frogs.pdf",
        )

        self.assertEqual(filtered.discovery_type, "FI")
        self.assertEqual([req.number for req in filtered.requests], ["1.1", "3.1"])

    @patch(
        "icharlotte_core.ui.wizard.pages.respond_discovery_page.extract_selected_form_interrogatory_numbers",
        return_value=["1.1"],
    )
    def test_empty_fi_parse_still_creates_reviewable_checked_rows(self, _mock_selected):
        parsed = ParsedDiscovery(
            discovery_type="Form Interrogatories",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[],
        )

        completed = _normalize_and_filter_parsed_discovery(
            parsed,
            detected_type="FI",
            discovery_file=r"C:\case\frogs.pdf",
        )

        self.assertEqual(completed.discovery_type, "FI")
        self.assertEqual(len(completed.requests), 1)
        self.assertEqual(completed.requests[0].number, "1.1")

    @patch("icharlotte_core.llm_config.call_llm")
    def test_fi_fixed_mode_does_not_draft_already_fixed_responses(self, mock_call):
        parsed = ParsedDiscovery(
            discovery_type="FI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[ParsedRequest(number="1.1", text="State who answered.")],
        )

        responses = _draft_substantive_response_map(
            parsed,
            context_text="",
            fi_mode="fixed",
        )

        self.assertEqual(responses, {})
        mock_call.assert_not_called()

    @patch("icharlotte_core.llm_config.call_llm")
    def test_structured_proposal_map_uses_request_specific_context(self, mock_call):
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[
                ParsedRequest(number="1", text="Identify all witnesses."),
                ParsedRequest(number="2", text="Identify all documents."),
            ],
        )
        mock_call.side_effect = [
            '{"request_number":"1","conditional_objection_rule_ids":[],"applied_custom_rule_ids":[],"applied_instruction_rule_ids":["minimal_direct_answer"],"ambiguous_term":"","proposed_objections":"","proposed_substantive_response":"John Smith.","needs_review":false,"review_reason":""}',
            '{"request_number":"2","conditional_objection_rule_ids":[],"applied_custom_rule_ids":[],"applied_instruction_rule_ids":["minimal_direct_answer"],"ambiguous_term":"","proposed_objections":"","proposed_substantive_response":"Photos and repair invoices.","needs_review":false,"review_reason":""}',
        ]

        proposals = _build_structured_proposal_map(
            parsed=parsed,
            selected_rules=[],
            context_text_by_path={
                "status.txt": "Witnesses\nJohn Smith saw the collision.\n\nDocuments\nPhotos and repair invoices exist."
            },
            response_rules=ResponseRules(),
        )

        self.assertEqual(proposals["1"].proposed_substantive_response, "John Smith.")
        self.assertEqual(proposals["2"].proposed_substantive_response, "Photos and repair invoices.")
        first_prompt = mock_call.call_args_list[0].args[0]
        second_prompt = mock_call.call_args_list[1].args[0]
        self.assertIn("John Smith saw the collision", first_prompt)
        self.assertIn("Photos and repair invoices exist", second_prompt)

    def test_fixed_fi_non_fixed_request_uses_proposal_substantive_callback(self):
        parsed = ParsedDiscovery(
            discovery_type="FI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[
                ParsedRequest(
                    number="2.1",
                    text="State your name and contact information.",
                )
            ],
        )
        proposal = StructuredProposal(
            request_number="2.1",
            proposed_substantive_response="Defendant Jones resides at 100 Main Street.",
            needs_review=True,
            review_reason="No specific context found.",
        )

        review_state = respond_discovery_page._generate_review_state_from_proposals(
            parsed,
            selected_rules=[],
            response_rules=ResponseRules(),
            proposal_map={"2.1": proposal},
            fi_mode="fixed",
        )

        self.assertEqual(
            review_state.requests[0].proposed_substantive_response,
            "Defendant Jones resides at 100 Main Street.",
        )
        self.assertTrue(review_state.requests[0].needs_review)
        self.assertEqual(
            review_state.requests[0].review_reason,
            "No specific context found.",
        )

    @patch("icharlotte_core.llm_config.call_llm")
    def test_generate_review_from_parsed_uses_structured_proposals(self, mock_call):
        mock_call.return_value = (
            '{"request_number":"1",'
            '"conditional_objection_rule_ids":[],'
            '"applied_custom_rule_ids":[],'
            '"applied_instruction_rule_ids":[],'
            '"ambiguous_term":"",'
            '"proposed_objections":"",'
            '"proposed_substantive_response":"Jane Roe witnessed the collision.",'
            '"needs_review":false,'
            '"review_reason":""}'
        )
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[ParsedRequest(number="1", text="Identify all witnesses.")],
        )

        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            context_file = Path(tmp) / "status.txt"
            context_file.write_text(
                "Witnesses section: Jane Roe saw the collision.",
                encoding="utf-8",
            )
            page = RespondDiscoverySettingsPage(
                case_root=tmp,
                file_number="",
                discovery_file=str(Path(tmp) / "srogg.txt"),
                detected_type="SI",
                parsed_discovery=parsed,
            )
            page.context_files = [str(context_file)]

            page._generate_review_from_parsed()

        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "Jane Roe witnessed the collision.",
        )
        self.assertIn("Jane Roe saw the collision", mock_call.call_args.args[0])

    @patch("Scripts.case_data_manager.CaseDataManager")
    def test_loads_advanced_mode_response_rules(self, mock_manager_cls):
        mock_manager_cls.return_value.get_value.return_value = {
            "fi_objections": "Saved FI objections.",
            "fi_1_1_response": "Saved 1.1 response.",
        }

        rules = load_respond_response_rules("1234.001")

        self.assertEqual(rules.fi_objections, "Saved FI objections.")
        self.assertEqual(rules.fi_1_1_response, "Saved 1.1 response.")
        mock_manager_cls.return_value.get_value.assert_called_once_with(
            "1234.001",
            "respond_rules",
        )

    def test_next_marks_current_response_approved_without_checkbox(self):
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[
                ParsedRequest(number="1", text="Identify witnesses."),
                ParsedRequest(number="2", text="Identify documents."),
            ],
        )
        state = ReviewState(
            [
                RequestReview(number="1", request_text="Identify witnesses."),
                RequestReview(number="2", request_text="Identify documents."),
            ]
        )
        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
            parsed_discovery=parsed,
            review_state=state,
        )

        self.assertFalse(hasattr(page, "approved_cb"))
        page.next_review_btn.click()

        self.assertTrue(page.review_state.requests[0].approved)
        self.assertFalse(page.review_state.requests[1].approved)

    def test_review_warning_label_shows_current_request_reason(self):
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
            [
                RequestReview(
                    number="1",
                    request_text="Identify witnesses.",
                    proposed_substantive_response="Unknown.",
                    needs_review=True,
                    review_reason="No specific context found.",
                )
            ]
        )

        page = RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
            parsed_discovery=parsed,
            review_state=state,
        )

        self.assertFalse(page.review_warning_label.isHidden())
        self.assertIn("No specific context found", page.review_warning_label.text())

    def test_context_file_picker_starts_in_status_folder_next_to_discovery_file(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            propounded = Path(tmp) / "DISCOVERY" / "PROPOUNDED"
            status = propounded / "STATUS"
            status.mkdir(parents=True)
            discovery_file = propounded / "srogg.pdf"
            discovery_file.write_text("SPECIAL INTERROGATORIES")

            page = RespondDiscoverySettingsPage(
                case_root=tmp,
                file_number="1234.001",
                discovery_file=str(discovery_file),
                detected_type="SI",
            )

            with patch.object(page, "_generate_proposals") as mock_generate:
                with patch(
                    "icharlotte_core.ui.wizard.pages.respond_discovery_page.QFileDialog.getOpenFileNames",
                    return_value=([], ""),
                ) as mock_dialog:
                    page._on_select_context_files()

            self.assertEqual(mock_dialog.call_args.args[2], str(status))
            mock_generate.assert_called_once()

    def test_context_file_picker_prefers_case_status_folder(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            case_status = Path(tmp) / "STATUS"
            case_status.mkdir()
            propounded = Path(tmp) / "DISCOVERY" / "PROPOUNDED"
            propounded.mkdir(parents=True)
            discovery_file = propounded / "srogg.pdf"
            discovery_file.write_text("SPECIAL INTERROGATORIES")

            page = RespondDiscoverySettingsPage(
                case_root=tmp,
                file_number="3000.075",
                discovery_file=str(discovery_file),
                detected_type="SI",
            )

            with patch.object(page, "_generate_proposals") as mock_generate:
                with patch(
                    "icharlotte_core.ui.wizard.pages.respond_discovery_page.QFileDialog.getOpenFileNames",
                    return_value=([], ""),
                ) as mock_dialog:
                    page._on_select_context_files()

            self.assertEqual(mock_dialog.call_args.args[2], str(case_status))
            mock_generate.assert_called_once()

    def test_si_quick_response_refer_to_document_inserts_text(self):
        page = self._review_page_for_type("SI")

        self.assertEqual(page.quick_response_labels(), ["Refer to Document"])
        page.quick_response_button("Refer to Document").click()

        self.assertIn("Code of Civil Procedure section 2030.230", page.response_edit.toPlainText())
        self.assertIn("documents produced concurrently herewith", page.response_edit.toPlainText())

    def test_fi_quick_response_refer_to_document_inserts_text(self):
        page = self._review_page_for_type("FI")

        self.assertEqual(page.quick_response_labels(), ["Refer to Document"])
        page.quick_response_button("Refer to Document").click()

        self.assertIn("Code of Civil Procedure section 2030.230", page.response_edit.toPlainText())

    def test_rpd_quick_responses_insert_production_language(self):
        page = self._review_page_for_type("RPD")

        self.assertEqual(page.quick_response_labels(), ["Will produce", "Wont produce"])
        page.quick_response_button("Will produce").click()
        self.assertIn("will comply with this request", page.response_edit.toPlainText())

        page.quick_response_button("Wont produce").click()
        self.assertIn("unable to comply", page.response_edit.toPlainText())

    def test_rfa_quick_responses_insert_admission_language(self):
        page = self._review_page_for_type("RFA")

        self.assertEqual(page.quick_response_labels(), ["Admit", "Deny", "Cant Admit/Deny"])
        page.quick_response_button("Admit").click()
        self.assertEqual(page.response_edit.toPlainText(), "Admit.")

        page.quick_response_button("Deny").click()
        self.assertEqual(page.response_edit.toPlainText(), "Deny.")

        page.quick_response_button("Cant Admit/Deny").click()
        self.assertIn("insufficient to enable Responding Party to admit", page.response_edit.toPlainText())

    def _review_page_for_type(self, discovery_type: str) -> RespondDiscoverySettingsPage:
        parsed = ParsedDiscovery(
            discovery_type=discovery_type,
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[ParsedRequest(number="1", text="Request text.")],
        )
        state = ReviewState(
            [RequestReview(number="1", request_text="Request text.")]
        )
        return RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=fr"C:\case\{discovery_type.lower()}.pdf",
            detected_type=discovery_type,
            parsed_discovery=parsed,
            review_state=state,
        )


class RespondDiscoveryWorkerTests(unittest.TestCase):
    def test_worker_rejects_unapproved_requests(self):
        worker = RespondDiscoveryWorker(
            case_path="",
            file_number="1234.001",
            settings={
                "review_state": {
                    "requests": [
                        {
                            "number": "1",
                            "request_text": "Q",
                            "proposed_objections": "O",
                            "proposed_substantive_response": "R",
                            "approved": False,
                        }
                    ]
                }
            },
        )

        ok, message = worker.validate_settings()

        self.assertFalse(ok)
        self.assertIn("approved", message.lower())

    @patch(
        "icharlotte_core.discovery.assembler.DiscoveryAssembler.find_caption_page",
        return_value=r"C:\case\caption.docx",
    )
    @patch("icharlotte_core.discovery.response_assembler.ResponseAssembler")
    @patch("icharlotte_core.word_validator.validate_discovery_response_docx")
    def test_worker_assembles_approved_review(
        self,
        mock_validate,
        mock_assembler_cls,
        _mock_caption,
    ):
        class Validation:
            has_errors = False

        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="Plaintiff Smith",
            responding_party="Defendant Jones",
            set_number=1,
            set_word="ONE",
            case_number="123",
            requests=[ParsedRequest(number="1", text="Identify witnesses.")],
        )
        review_state = ReviewState(
            [
                RequestReview(
                    number="1",
                    request_text="Identify witnesses.",
                    proposed_objections="Objection.",
                    proposed_substantive_response="No witnesses known.",
                    approved=True,
                )
            ]
        )
        mock_validate.return_value = Validation()
        mock_assembler_cls.return_value.assemble.return_value = r"C:\case\out.docx"
        worker = RespondDiscoveryWorker(
            case_path=r"C:\case",
            file_number="1234.001",
            settings={
                "parsed_discovery": asdict(parsed),
                "review_state": review_state.to_dict(),
            },
        )
        results = []
        worker.finished_result.connect(lambda ok, payload: results.append((ok, payload)))

        worker.run()

        self.assertTrue(results[0][0], results)
        self.assertIn(".icharlotte", results[0][1])
        self.assertNotIn("DISCOVERY\\RESPONSES", results[0][1])
        assembled_text = mock_assembler_cls.return_value.assemble.call_args.kwargs["response_text"]
        self.assertIn("Objection.", assembled_text)
        self.assertIn("No witnesses known.", assembled_text)
        mock_validate.assert_called_once()


if __name__ == "__main__":
    unittest.main()
