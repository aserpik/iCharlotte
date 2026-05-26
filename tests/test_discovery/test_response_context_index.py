import unittest

from icharlotte_core.discovery.response_context_index import (
    ContextChunk,
    build_context_chunks,
    format_context_packet,
    select_context_packet,
)
from icharlotte_core.discovery.response_parser import ParsedRequest


class ResponseContextIndexTests(unittest.TestCase):
    def test_build_context_chunks_splits_headings_and_paragraphs(self):
        chunks = build_context_chunks(
            {
                r"C:\case\status.txt": (
                    "Witnesses\n"
                    "John Smith saw the impact.\n\n"
                    "Damages\n"
                    "Plaintiff claims neck pain."
                )
            }
        )

        self.assertGreaterEqual(len(chunks), 2)
        self.assertEqual(chunks[0].source_path, r"C:\case\status.txt")
        self.assertEqual(chunks[0].sequence, 1)
        self.assertEqual(chunks[0].heading, "Witnesses")
        self.assertEqual(chunks[1].heading, "Damages")
        self.assertIn("John Smith", "\n".join(chunk.text for chunk in chunks))

    def test_build_context_chunks_does_not_treat_sentence_start_as_heading(self):
        chunks = build_context_chunks(
            {
                r"C:\case\status.txt": (
                    "John Smith saw the impact and reported it.\n"
                    "Plaintiff later claimed neck pain."
                )
            }
        )

        self.assertEqual(len(chunks), 1)
        self.assertEqual(chunks[0].heading, "")

    def test_build_context_chunks_does_not_treat_unpunctuated_soft_wrap_as_heading(self):
        chunks = build_context_chunks(
            {
                r"C:\case\status.txt": (
                    "John Smith saw the impact and reported it\n"
                    "Plaintiff later claimed neck pain."
                )
            }
        )

        self.assertEqual(len(chunks), 1)
        self.assertEqual(chunks[0].heading, "")

    def test_select_context_packet_prefers_request_terms(self):
        chunks = [
            ContextChunk("status.txt", 1, "Witnesses\nJohn Smith saw the impact.", "Witnesses"),
            ContextChunk("status.txt", 2, "Damages\nPlaintiff claims neck pain.", "Damages"),
        ]
        request = ParsedRequest(number="1", text="Identify all witnesses to the INCIDENT.")

        selected = select_context_packet(request, chunks, max_chunks=1)

        self.assertEqual(len(selected), 1)
        self.assertIn("John Smith", selected[0].text)

    def test_select_context_packet_returns_empty_for_no_signal(self):
        chunks = [
            ContextChunk("status.txt", 1, "Billing notes only.", ""),
        ]
        request = ParsedRequest(number="1", text="Identify all witnesses.")

        selected = select_context_packet(request, chunks, max_chunks=3)

        self.assertEqual(selected, [])

    def test_format_context_packet_includes_source_labels(self):
        text = format_context_packet(
            [ContextChunk("status.txt", 2, "John Smith saw the impact.", "Witnesses")]
        )

        self.assertIn("[status.txt #2]", text)
        self.assertIn("John Smith saw the impact.", text)


if __name__ == "__main__":
    unittest.main()
