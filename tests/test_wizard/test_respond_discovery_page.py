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
import icharlotte_core.ui.wizard.pages.respond_discovery_page as respond_discovery_page
from icharlotte_core.ui.wizard.pages.respond_discovery_page import (
    RespondDiscoverySettingsPage,
    RespondDiscoveryWorker,
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
        page._open_review_screen_with_pending(parsed)

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

    def test_review_warning_label_shows_generic_message_without_reason(self):
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
                    review_reason="",
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
        self.assertEqual(page.review_warning_label.text(), "Needs review.")

    def test_review_warning_label_updates_when_moving_between_requests(self):
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
                RequestReview(
                    number="1",
                    request_text="Identify witnesses.",
                    proposed_substantive_response="Unknown.",
                    needs_review=True,
                    review_reason="No specific context found.",
                ),
                RequestReview(
                    number="2",
                    request_text="Identify documents.",
                    proposed_substantive_response="Documents will be produced.",
                    needs_review=False,
                    review_reason="",
                ),
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
        self.assertEqual(
            page.review_warning_label.text(),
            "Needs review: No specific context found.",
        )

        page.next_review_btn.click()

        self.assertTrue(page.review_warning_label.isHidden())
        self.assertEqual(page.review_warning_label.text(), "")

        page.prev_btn.click()

        self.assertFalse(page.review_warning_label.isHidden())
        self.assertEqual(
            page.review_warning_label.text(),
            "Needs review: No specific context found.",
        )

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
                    "icharlotte_core.ui.wizard.pages.respond_discovery_page.ContextFilesDialog.get_files",
                    return_value=[],
                ) as mock_dialog:
                    page._on_select_context_files()

            self.assertEqual(mock_dialog.call_args.kwargs["start_dir"], str(status))
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
                    "icharlotte_core.ui.wizard.pages.respond_discovery_page.ContextFilesDialog.get_files",
                    return_value=[],
                ) as mock_dialog:
                    page._on_select_context_files()

            self.assertEqual(mock_dialog.call_args.kwargs["start_dir"], str(case_status))
            mock_generate.assert_called_once()

    def test_context_file_picker_cancel_aborts_generation(self):
        with tempfile.TemporaryDirectory(dir="C:\\geminiterminal2") as tmp:
            propounded = Path(tmp) / "DISCOVERY" / "PROPOUNDED"
            propounded.mkdir(parents=True)
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
                    "icharlotte_core.ui.wizard.pages.respond_discovery_page.ContextFilesDialog.get_files",
                    return_value=None,
                ):
                    page._on_select_context_files()

            mock_generate.assert_not_called()
            self.assertEqual(page.context_files, [])

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


from icharlotte_core.discovery.response_generation_engine import (
    StructuredProposal,
)


class _MakeParsedSI:
    def __call__(self, n=2):
        from icharlotte_core.discovery.response_parser import ParsedDiscovery, ParsedRequest
        return ParsedDiscovery(
            discovery_type="SI",
            propounding_party="P",
            responding_party="D",
            set_number=1,
            set_word="ONE",
            case_number="1",
            requests=[
                ParsedRequest(number=str(i + 1), text=f"Request {i + 1}.")
                for i in range(n)
            ],
        )


class StreamingReviewStateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _make_page(self):
        return RespondDiscoverySettingsPage(
            case_root="",
            file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
        )

    def test_review_screen_opens_with_all_requests_pending(self):
        page = self._make_page()
        parsed = _MakeParsedSI()(n=3)
        page._open_review_screen_with_pending(parsed)

        self.assertIsNotNone(page.review_state)
        self.assertEqual(len(page.review_state.requests), 3)
        for review in page.review_state.requests:
            self.assertTrue(review.is_pending)

    def test_apply_proposal_clears_pending_on_matching_row(self):
        page = self._make_page()
        parsed = _MakeParsedSI()(n=2)
        page._open_review_screen_with_pending(parsed)
        page.parsed_discovery = parsed

        proposal = StructuredProposal(
            request_number="1",
            proposed_substantive_response="Drafted answer for 1.",
        )
        page._on_proposal_ready("1", proposal)

        self.assertFalse(page.review_state.requests[0].is_pending)
        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "Drafted answer for 1.",
        )
        # The second row is still pending.
        self.assertTrue(page.review_state.requests[1].is_pending)

    def test_proposal_does_not_overwrite_user_edits(self):
        page = self._make_page()
        parsed = _MakeParsedSI()(n=2)
        page._open_review_screen_with_pending(parsed)
        page.parsed_discovery = parsed

        # User edits the first row's response while it's still pending.
        page.review_state.requests[0].proposed_substantive_response = "USER TYPED THIS"

        proposal = StructuredProposal(
            request_number="1",
            proposed_substantive_response="Drafted answer.",
        )
        page._on_proposal_ready("1", proposal)

        # The user text survives; the new draft is stashed in pending_replacement.
        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "USER TYPED THIS",
        )
        self.assertIsNotNone(page.review_state.requests[0].pending_replacement)
        self.assertEqual(
            page.review_state.requests[0].pending_replacement.proposed_substantive_response,
            "Drafted answer.",
        )
        # is_pending is cleared either way.
        self.assertFalse(page.review_state.requests[0].is_pending)

    def test_finalize_disabled_until_no_pending_remain(self):
        page = self._make_page()
        parsed = _MakeParsedSI()(n=2)
        page._open_review_screen_with_pending(parsed)
        page.parsed_discovery = parsed
        page._show_review()

        # No proposals delivered yet -> Finalize is disabled.
        self.assertFalse(page.finalize_btn.isEnabled())

        # Deliver both proposals.
        page._on_proposal_ready(
            "1",
            StructuredProposal(request_number="1", proposed_substantive_response="A1"),
        )
        page._on_proposal_ready(
            "2",
            StructuredProposal(request_number="2", proposed_substantive_response="A2"),
        )
        page._on_coordinator_all_done()
        # All rows arrived; Finalize is enabled once user approves them via UI.
        # We only check the "no longer disabled by pending" half here.
        self.assertTrue(page.finalize_btn.isEnabled())


class ReviewScreenStatusIndicatorTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _make_page_with_pending(self, n=3):
        from icharlotte_core.discovery.response_parser import (
            ParsedDiscovery, ParsedRequest,
        )
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="P",
            responding_party="D",
            set_number=1, set_word="ONE", case_number="1",
            requests=[
                ParsedRequest(number=str(i + 1), text=f"Request {i + 1}.")
                for i in range(n)
            ],
        )
        page = RespondDiscoverySettingsPage(
            case_root="", file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
        )
        page.parsed_discovery = parsed
        page._open_review_screen_with_pending(parsed)
        return page, parsed

    def test_pending_row_shows_generating_placeholder_in_response_pane(self):
        page, _ = self._make_page_with_pending()
        page._current_review_index = 0
        page._load_current_review()
        self.assertEqual(page.response_edit.toPlainText(), "")
        self.assertTrue(page.response_edit.isReadOnly())
        self.assertIn("Generating", page.review_warning_label.text())

    def test_completed_row_enables_edits(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page, parsed = self._make_page_with_pending()
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1", proposed_substantive_response="Done."
            ),
        )
        page._current_review_index = 0
        page._load_current_review()
        self.assertFalse(page.response_edit.isReadOnly())
        self.assertEqual(page.response_edit.toPlainText(), "Done.")

    def test_status_bar_reflects_progress(self):
        page, parsed = self._make_page_with_pending(n=4)
        page._on_coordinator_progress(2, 4)
        self.assertIn("2", page.review_status_label.text())
        self.assertIn("4", page.review_status_label.text())

    def test_request_number_header_includes_status_icon(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page, parsed = self._make_page_with_pending(n=2)
        page._current_review_index = 0
        page._load_current_review()
        self.assertIn("⏳", page.request_label.text())  # hourglass U+23F3
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1",
                proposed_substantive_response="Done.",
                needs_review=True,
                review_reason="weak context",
            ),
        )
        page._load_current_review()
        self.assertIn("⚠", page.request_label.text())  # warning U+26A0


class RegenerateAndConflictTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _make_page(self, n=2):
        from icharlotte_core.discovery.response_parser import (
            ParsedDiscovery, ParsedRequest,
        )
        parsed = ParsedDiscovery(
            discovery_type="SI",
            propounding_party="P",
            responding_party="D",
            set_number=1, set_word="ONE", case_number="1",
            requests=[
                ParsedRequest(number=str(i + 1), text=f"Request {i + 1}.")
                for i in range(n)
            ],
        )
        page = RespondDiscoverySettingsPage(
            case_root="", file_number="1234.001",
            discovery_file=r"C:\case\srogg.pdf",
            detected_type="SI",
        )
        page.parsed_discovery = parsed
        page._open_review_screen_with_pending(parsed)
        return page, parsed

    def test_regenerate_button_calls_coordinator_with_current_request(self):
        page, parsed = self._make_page()
        # Install a real coordinator with a no-op stub.
        regen_calls = []
        class _Stub:
            def regenerate(self, request_number, **kwargs):
                regen_calls.append((request_number, kwargs.get("override_instruction", "")))
                return True
            def cancel(self): pass
            def start(self, **kwargs): pass
            def is_done(self): return False
        page._coordinator = _Stub()
        page._context_chunks = []
        page._loaded_response_rules = None

        page._current_review_index = 0
        page.regenerate_instruction_edit.setText("more privilege")
        page._on_regenerate_clicked()

        self.assertEqual(regen_calls, [("1", "more privilege")])

    def test_regenerate_marks_row_pending_until_proposal_arrives(self):
        page, parsed = self._make_page()
        class _Stub:
            def regenerate(self, **kwargs): return True
            def cancel(self): pass
            def start(self, **kwargs): pass
            def is_done(self): return False
        page._coordinator = _Stub()
        page._context_chunks = []
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1", proposed_substantive_response="Initial."
            ),
        )
        self.assertFalse(page.review_state.requests[0].is_pending)

        page._current_review_index = 0
        page.regenerate_instruction_edit.setText("")
        page._on_regenerate_clicked()
        self.assertTrue(page.review_state.requests[0].is_pending)

    def test_conflict_banner_view_apply_replaces_user_text(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page, parsed = self._make_page()
        page.review_state.requests[0].proposed_substantive_response = "USER TEXT"
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1", proposed_substantive_response="DRAFT"
            ),
        )
        page._current_review_index = 0
        page._load_current_review()
        self.assertTrue(page.conflict_banner.isVisibleTo(page))

        page._apply_pending_replacement()
        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "DRAFT",
        )
        self.assertIsNone(page.review_state.requests[0].pending_replacement)
        self.assertFalse(page.conflict_banner.isVisibleTo(page))

    def test_conflict_banner_discard_keeps_user_text(self):
        from icharlotte_core.discovery.response_generation_engine import (
            StructuredProposal,
        )
        page, parsed = self._make_page()
        page.review_state.requests[0].proposed_substantive_response = "USER TEXT"
        page._on_proposal_ready(
            "1",
            StructuredProposal(
                request_number="1", proposed_substantive_response="DRAFT"
            ),
        )
        page._current_review_index = 0
        page._load_current_review()
        page._discard_pending_replacement()
        self.assertEqual(
            page.review_state.requests[0].proposed_substantive_response,
            "USER TEXT",
        )
        self.assertIsNone(page.review_state.requests[0].pending_replacement)


if __name__ == "__main__":
    unittest.main()
