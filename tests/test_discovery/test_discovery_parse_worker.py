import unittest
from unittest.mock import patch, MagicMock

import pytest

pytest.importorskip("pytestqt")
from PySide6.QtWidgets import QApplication

from icharlotte_core.discovery.discovery_parse_worker import DiscoveryParseWorker
from icharlotte_core.discovery.response_parser import ParsedDiscovery


class DiscoveryParseWorkerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._app = QApplication.instance() or QApplication([])

    def _run_synchronously(self, worker):
        # Invoke the body directly instead of starting a thread.
        worker.run()

    def test_emits_parsed_discovery_on_success(self):
        emitted = []

        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            return_value="some discovery text",
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.call_llm",
            return_value="ignored",
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.parse_llm_response",
            return_value=ParsedDiscovery(
                discovery_type="SI",
                propounding_party="P",
                responding_party="D",
                set_number=1,
                set_word="ONE",
                case_number="1",
                requests=[],
            ),
        ):
            worker = DiscoveryParseWorker(
                discovery_file="x.pdf",
                detected_type="SI",
            )
            worker.parse_finished.connect(
                lambda ok, payload: emitted.append((ok, payload))
            )
            self._run_synchronously(worker)

        self.assertEqual(len(emitted), 1)
        ok, parsed = emitted[0]
        self.assertTrue(ok)
        self.assertIsInstance(parsed, ParsedDiscovery)
        self.assertEqual(parsed.discovery_type, "SI")

    def test_passes_confirmed_fi_numbers_to_filter(self):
        captured = {}

        def fake_normalize(parsed, detected_type, discovery_file, selected_fi_numbers=None):
            captured["selected"] = selected_fi_numbers
            return parsed

        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            return_value="FORM INTERROGATORIES text",
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.call_llm",
            return_value="ignored",
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.parse_llm_response",
            return_value=ParsedDiscovery(
                discovery_type="FI",
                propounding_party="P",
                responding_party="D",
                set_number=1,
                set_word="ONE",
                case_number="1",
                requests=[],
            ),
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.normalize_and_filter_parsed_discovery",
            side_effect=fake_normalize,
        ):
            worker = DiscoveryParseWorker(
                discovery_file="frog.pdf",
                detected_type="FI",
                selected_fi_numbers=["1.1", "2.4"],
            )
            worker.parse_finished.connect(lambda *_: None)
            self._run_synchronously(worker)

        self.assertEqual(captured["selected"], ["1.1", "2.4"])

    def test_emits_error_when_file_is_empty(self):
        emitted = []
        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            return_value="",
        ):
            worker = DiscoveryParseWorker(
                discovery_file="x.pdf", detected_type="SI",
            )
            worker.parse_finished.connect(
                lambda ok, payload: emitted.append((ok, payload))
            )
            self._run_synchronously(worker)
        ok, payload = emitted[0]
        self.assertFalse(ok)
        self.assertIn("Could not read text", payload)

    def test_emits_error_when_llm_returns_nothing(self):
        emitted = []
        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            return_value="text",
        ), patch(
            "icharlotte_core.discovery.discovery_parse_worker.call_llm",
            return_value="",
        ):
            worker = DiscoveryParseWorker(
                discovery_file="x.pdf", detected_type="SI",
            )
            worker.parse_finished.connect(
                lambda ok, payload: emitted.append((ok, payload))
            )
            self._run_synchronously(worker)
        ok, payload = emitted[0]
        self.assertFalse(ok)
        self.assertIn("parser did not return", payload)

    def test_emits_error_when_unexpected_exception(self):
        emitted = []
        with patch(
            "icharlotte_core.discovery.discovery_parse_worker.read_document_text",
            side_effect=RuntimeError("disk gone"),
        ):
            worker = DiscoveryParseWorker(
                discovery_file="x.pdf", detected_type="SI",
            )
            worker.parse_finished.connect(
                lambda ok, payload: emitted.append((ok, payload))
            )
            self._run_synchronously(worker)
        ok, payload = emitted[0]
        self.assertFalse(ok)
        self.assertIn("disk gone", payload)
