"""Tests for the chat markdown renderer."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

from icharlotte_core.chat.markdown_render import render_markdown, CHAT_MARKDOWN_CSS


class TestRenderMarkdown(unittest.TestCase):
    def test_pipe_table_becomes_bordered_table(self):
        md = (
            "| Col A | Col B |\n"
            "| ----- | ----- |\n"
            "| a1    | b1    |\n"
            "| a2    | b2    |\n"
        )
        html = render_markdown(md)
        # Must produce a real <table> with visible border attributes so Qt renders gridlines
        self.assertIn("<table", html)
        self.assertIn('border="1"', html)
        self.assertIn('cellpadding="4"', html)
        self.assertIn('cellspacing="0"', html)
        # Header and body cells preserved
        self.assertIn("<th>Col A</th>", html)
        self.assertIn("<td>a1</td>", html)

    def test_fenced_code_block_preserved(self):
        md = "```python\nprint('hi')\n```"
        html = render_markdown(md)
        self.assertIn("<pre>", html)
        self.assertIn("<code", html)
        self.assertIn("print(", html)

    def test_inline_emphasis(self):
        html = render_markdown("This is **bold** and *italic* and `code`.")
        self.assertIn("<strong>bold</strong>", html)
        self.assertIn("<em>italic</em>", html)
        self.assertIn("<code>code</code>", html)

    def test_newlines_become_breaks_gfm_style(self):
        # nl2br: single newline inside a paragraph -> <br>
        html = render_markdown("line one\nline two")
        self.assertIn("<br", html)

    def test_bulleted_list(self):
        md = "- one\n- two\n- three"
        html = render_markdown(md)
        self.assertIn("<ul>", html)
        self.assertEqual(html.count("<li>"), 3)

    def test_blockquote(self):
        html = render_markdown("> quoted text")
        self.assertIn("<blockquote>", html)

    def test_empty_input(self):
        self.assertEqual(render_markdown(""), "")
        self.assertEqual(render_markdown(None), "")

    def test_table_with_inline_formatting(self):
        md = (
            "| Name | Note |\n"
            "| ---- | ---- |\n"
            "| **A** | `x` |\n"
        )
        html = render_markdown(md)
        self.assertIn("<table", html)
        self.assertIn("<strong>A</strong>", html)
        self.assertIn("<code>x</code>", html)

    def test_falls_back_on_invalid_input(self):
        # Even if extensions misbehave, we should always return a string
        html = render_markdown("plain text with no markdown")
        self.assertIsInstance(html, str)
        self.assertIn("plain text", html)

    def test_css_block_defines_table_borders(self):
        # Sanity: the shared stylesheet covers table borders and code styling
        self.assertIn("table", CHAT_MARKDOWN_CSS)
        self.assertIn("border", CHAT_MARKDOWN_CSS)
        self.assertIn("pre", CHAT_MARKDOWN_CSS)


if __name__ == "__main__":
    unittest.main()
