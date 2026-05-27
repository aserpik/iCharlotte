"""Regression test: Claude CLI subprocess must use UTF-8 encoding.

Without explicit encoding="utf-8", subprocess.Popen(..., text=True) defaults to
the system locale on Windows (cp1252), which cannot encode characters in the
Unicode Private Use Area (e.g. , the Symbol-font bullet used by Word).
When a user attaches a .docx with bulleted lists, those PUA chars end up in
the combined prompt and crash proc.stdin.write with:
    'charmap' codec can't encode character '\\uf0b7' in position N
"""
import unittest
from unittest.mock import patch, MagicMock


PUA_BULLET = ""  # Word Symbol-font bullet (Private Use Area)


def _settings(stream):
    return {"stream": stream}


class FakePopen:
    """Stand-in for subprocess.Popen capturing kwargs and exposing minimal API."""
    last_kwargs = None
    last_stdin_written = None

    def __init__(self, args, **kwargs):
        FakePopen.last_kwargs = kwargs
        self.stdin = MagicMock()
        self.stdin.write = self._record_stdin
        self.stdout = iter([])  # no events -> generator exits cleanly
        self.stderr = MagicMock()
        self.stderr.read = lambda: ""
        self.returncode = 0

    @classmethod
    def _record_stdin(cls, data):
        cls.last_stdin_written = data

    def wait(self):
        return 0


class ClaudeStreamingUnicodeTests(unittest.TestCase):
    def setUp(self):
        FakePopen.last_kwargs = None
        FakePopen.last_stdin_written = None

    def test_streaming_popen_uses_utf8_encoding(self):
        from icharlotte_core import llm
        with patch.object(llm.subprocess, "Popen", FakePopen):
            gen = llm.LLMHandler.generate(
                provider="Claude",
                model="sonnet",
                system_prompt="",
                user_prompt=f"Summarize these bullets:\n{PUA_BULLET} item one",
                file_contents="",
                settings=_settings(stream=True),
            )
            # Drain the generator so Popen actually runs
            list(gen)

        kwargs = FakePopen.last_kwargs
        self.assertIsNotNone(kwargs, "Popen was not called")
        self.assertEqual(
            kwargs.get("encoding"),
            "utf-8",
            "Claude CLI subprocess must specify encoding='utf-8' to handle "
            "Unicode (e.g. Word PUA bullets) on Windows; otherwise cp1252 "
            "crashes proc.stdin.write.",
        )

    def test_streaming_prompt_with_pua_bullet_reaches_stdin(self):
        from icharlotte_core import llm
        with patch.object(llm.subprocess, "Popen", FakePopen):
            gen = llm.LLMHandler.generate(
                provider="Claude",
                model="sonnet",
                system_prompt="",
                user_prompt=f"Bullet: {PUA_BULLET}",
                file_contents="",
                settings=_settings(stream=True),
            )
            list(gen)

        written = FakePopen.last_stdin_written
        self.assertIsNotNone(written, "Nothing was written to stdin")
        self.assertIn(PUA_BULLET, written,
                      "PUA bullet was stripped/mangled before reaching stdin")

    def test_nonstreaming_run_uses_utf8_encoding(self):
        from icharlotte_core import llm
        fake_result = MagicMock(returncode=0, stdout="ok", stderr="")

        captured = {}
        def fake_run(*args, **kwargs):
            captured["kwargs"] = kwargs
            captured["input"] = kwargs.get("input")
            return fake_result

        with patch.object(llm.subprocess, "run", side_effect=fake_run):
            result = llm.LLMHandler.generate(
                provider="Claude",
                model="sonnet",
                system_prompt="",
                user_prompt=f"Bullet: {PUA_BULLET}",
                file_contents="",
                settings=_settings(stream=False),
            )

        self.assertEqual(result, "ok")
        self.assertEqual(
            captured["kwargs"].get("encoding"),
            "utf-8",
            "Non-streaming Claude CLI subprocess.run must specify "
            "encoding='utf-8' to handle Unicode on Windows.",
        )
        self.assertIn(PUA_BULLET, captured["input"])


if __name__ == "__main__":
    unittest.main()
