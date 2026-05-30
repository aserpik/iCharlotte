import os
import json
import tempfile
import unittest
from unittest.mock import patch, MagicMock

from icharlotte_core import llm


class MapModelTests(unittest.TestCase):
    def test_thinking_maps_to_codex(self):
        self.assertEqual(llm._map_openai_model_to_codex("gpt-5.2-thinking"), "gpt-5.2-codex")

    def test_instant_maps_to_codex(self):
        self.assertEqual(llm._map_openai_model_to_codex("gpt-5.2-instant"), "gpt-5.2-codex")

    def test_codex_id_passthrough(self):
        self.assertEqual(llm._map_openai_model_to_codex("gpt-5-codex"), "gpt-5-codex")

    def test_unknown_gpt5_maps_to_generic_codex(self):
        self.assertEqual(llm._map_openai_model_to_codex("gpt-5.9-foo"), "gpt-5-codex")

    def test_gpt4o_unsupported(self):
        self.assertIsNone(llm._map_openai_model_to_codex("gpt-4o"))

    def test_o1_unsupported(self):
        self.assertIsNone(llm._map_openai_model_to_codex("o1"))

    def test_empty_unsupported(self):
        self.assertIsNone(llm._map_openai_model_to_codex(""))
        self.assertIsNone(llm._map_openai_model_to_codex(None))


class CodexAvailableTests(unittest.TestCase):
    def test_available_when_on_path_and_logged_in(self):
        with patch.object(llm.shutil, "which", return_value="C:/codex.exe"), \
             patch.object(llm.os.path, "isfile", return_value=True):
            self.assertTrue(llm.codex_available())

    def test_unavailable_when_not_on_path(self):
        with patch.object(llm.shutil, "which", return_value=None), \
             patch.object(llm.os.path, "isfile", return_value=True):
            self.assertFalse(llm.codex_available())

    def test_unavailable_when_not_logged_in(self):
        with patch.object(llm.shutil, "which", return_value="C:/codex.exe"), \
             patch.object(llm.os.path, "isfile", return_value=False):
            self.assertFalse(llm.codex_available())


class SubscriptionEnabledTests(unittest.TestCase):
    def _write_prefs(self, payload):
        fd, path = tempfile.mkstemp(suffix=".json")
        os.close(fd)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f)
        self.addCleanup(lambda: os.remove(path))
        return path

    def test_enabled_when_flag_true(self):
        path = self._write_prefs({"openai_use_subscription": True})
        with patch.object(llm, "_subscription_prefs_path", return_value=path):
            self.assertTrue(llm.openai_subscription_enabled())

    def test_disabled_when_flag_false(self):
        path = self._write_prefs({"openai_use_subscription": False})
        with patch.object(llm, "_subscription_prefs_path", return_value=path):
            self.assertFalse(llm.openai_subscription_enabled())

    def test_default_true_when_key_missing(self):
        path = self._write_prefs({"version": "2.1"})
        with patch.object(llm, "_subscription_prefs_path", return_value=path):
            self.assertTrue(llm.openai_subscription_enabled())

    def test_default_true_when_file_missing(self):
        with patch.object(llm, "_subscription_prefs_path",
                          return_value="C:/nonexistent/does-not-exist.json"):
            self.assertTrue(llm.openai_subscription_enabled())


class CodexGenerateTests(unittest.TestCase):
    def _fake_run_writes(self, text, returncode=0, stderr=""):
        def fake_run(cmd, **kwargs):
            idx = cmd.index("--output-last-message")
            out_path = cmd[idx + 1]
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(text)
            return MagicMock(returncode=returncode, stdout="BANNER NOISE", stderr=stderr)
        return fake_run

    def test_returns_last_message_not_stdout(self):
        with patch.object(llm.subprocess, "run", side_effect=self._fake_run_writes("Hello from Codex")):
            out = llm._generate_openai_codex_cli(
                "gpt-5.2-thinking", "sys", "hi", "", None, do_stream=False)
        self.assertEqual(out, "Hello from Codex")

    def test_streaming_returns_iterator(self):
        with patch.object(llm.subprocess, "run", side_effect=self._fake_run_writes("chunked")):
            gen = llm._generate_openai_codex_cli(
                "gpt-5.2-thinking", "sys", "hi", "", None, do_stream=True)
        self.assertEqual(list(gen), ["chunked"])

    def test_unsupported_model_raises(self):
        with self.assertRaises(ValueError):
            llm._generate_openai_codex_cli("gpt-4o", "sys", "hi", "", None, do_stream=False)

    def test_nonzero_exit_raises(self):
        def fake_run(cmd, **kwargs):
            return MagicMock(returncode=1, stdout="", stderr="boom")
        with patch.object(llm.subprocess, "run", side_effect=fake_run):
            with self.assertRaises(Exception):
                llm._generate_openai_codex_cli(
                    "gpt-5.2-thinking", "sys", "hi", "", None, do_stream=False)

    def test_command_uses_readonly_sandbox_and_model(self):
        captured = {}
        def fake_run(cmd, **kwargs):
            captured["cmd"] = cmd
            idx = cmd.index("--output-last-message")
            with open(cmd[idx + 1], "w", encoding="utf-8") as f:
                f.write("ok")
            return MagicMock(returncode=0, stdout="", stderr="")
        with patch.object(llm.subprocess, "run", side_effect=fake_run):
            llm._generate_openai_codex_cli(
                "gpt-5.2-thinking", "sys", "hi", "", None, do_stream=False)
        cmd = captured["cmd"]
        self.assertEqual(cmd[:2], ["codex", "exec"])
        self.assertIn("--sandbox", cmd)
        self.assertIn("read-only", cmd)
        self.assertIn("gpt-5.2-codex", cmd)
        self.assertIn("--ask-for-approval", cmd)
        # the value follows the flag
        self.assertEqual(cmd[cmd.index("--ask-for-approval") + 1], "never")

    def test_openai_api_key_removed_from_subprocess_env(self):
        captured = {}
        def fake_run(cmd, **kwargs):
            captured["env"] = kwargs.get("env")
            idx = cmd.index("--output-last-message")
            with open(cmd[idx + 1], "w", encoding="utf-8") as f:
                f.write("ok")
            return MagicMock(returncode=0, stdout="", stderr="")
        with patch.dict(os.environ, {"OPENAI_API_KEY": "sk-should-be-removed"}, clear=False), \
             patch.object(llm.subprocess, "run", side_effect=fake_run):
            llm._generate_openai_codex_cli(
                "gpt-5.2-thinking", "sys", "hi", "", None, do_stream=False)
        self.assertIsNotNone(captured["env"])
        self.assertNotIn("OPENAI_API_KEY", captured["env"])


def _settings(stream=False):
    return {"stream": stream}


class RoutingTests(unittest.TestCase):
    def test_uses_codex_when_enabled_available_supported(self):
        with patch.object(llm, "openai_subscription_enabled", return_value=True), \
             patch.object(llm, "codex_available", return_value=True), \
             patch.object(llm, "_generate_openai_codex_cli", return_value="CODEX") as gen, \
             patch.dict(llm.API_KEYS, {"OpenAI": "sk-test"}, clear=False):
            out = llm.LLMHandler.generate(
                "OpenAI", "gpt-5.2-thinking", "sys", "hi", "", _settings())
        self.assertEqual(out, "CODEX")
        gen.assert_called_once()

    def test_falls_back_to_api_when_codex_raises(self):
        fake_resp = MagicMock(status_code=200)
        fake_resp.json.return_value = {"choices": [{"message": {"content": "API"}}]}
        with patch.object(llm, "openai_subscription_enabled", return_value=True), \
             patch.object(llm, "codex_available", return_value=True), \
             patch.object(llm, "_generate_openai_codex_cli", side_effect=Exception("boom")), \
             patch.object(llm.requests, "post", return_value=fake_resp), \
             patch.dict(llm.API_KEYS, {"OpenAI": "sk-test"}, clear=False):
            out = llm.LLMHandler.generate(
                "OpenAI", "gpt-4o", "sys", "hi", "", _settings())
        self.assertEqual(out, "API")

    def test_unsupported_model_skips_codex_uses_api(self):
        fake_resp = MagicMock(status_code=200)
        fake_resp.json.return_value = {"choices": [{"message": {"content": "API"}}]}
        with patch.object(llm, "openai_subscription_enabled", return_value=True), \
             patch.object(llm, "codex_available", return_value=True), \
             patch.object(llm, "_generate_openai_codex_cli") as gen, \
             patch.object(llm.requests, "post", return_value=fake_resp), \
             patch.dict(llm.API_KEYS, {"OpenAI": "sk-test"}, clear=False):
            out = llm.LLMHandler.generate(
                "OpenAI", "gpt-4o", "sys", "hi", "", _settings())
        gen.assert_not_called()
        self.assertEqual(out, "API")

    def test_no_api_key_with_subscription_does_not_raise(self):
        with patch.object(llm, "openai_subscription_enabled", return_value=True), \
             patch.object(llm, "codex_available", return_value=True), \
             patch.object(llm, "_generate_openai_codex_cli", return_value="CODEX"), \
             patch.dict(llm.API_KEYS, {"OpenAI": None}, clear=False):
            out = llm.LLMHandler.generate(
                "OpenAI", "gpt-5.2-thinking", "sys", "hi", "", _settings())
        self.assertEqual(out, "CODEX")

    def test_no_api_key_and_codex_raises_gives_clear_error(self):
        with patch.object(llm, "openai_subscription_enabled", return_value=True), \
             patch.object(llm, "codex_available", return_value=True), \
             patch.object(llm, "_generate_openai_codex_cli", side_effect=Exception("boom")), \
             patch.dict(llm.API_KEYS, {"OpenAI": None}, clear=False):
            with self.assertRaises(RuntimeError) as ctx:
                llm.LLMHandler.generate(
                    "OpenAI", "gpt-5.2-thinking", "sys", "hi", "", _settings())
            self.assertIn("no OPENAI_API_KEY", str(ctx.exception))

    def test_no_api_key_and_unsupported_model_gives_clear_error(self):
        with patch.object(llm, "openai_subscription_enabled", return_value=True), \
             patch.object(llm, "codex_available", return_value=True), \
             patch.dict(llm.API_KEYS, {"OpenAI": None}, clear=False):
            with self.assertRaises(RuntimeError) as ctx:
                llm.LLMHandler.generate(
                    "OpenAI", "gpt-4o", "sys", "hi", "", _settings())
            self.assertIn("no OPENAI_API_KEY", str(ctx.exception))


class ModelFilterTests(unittest.TestCase):
    def test_filters_to_codex_supported(self):
        ids = ["gpt-5.2-thinking", "gpt-5.2-instant", "gpt-4o", "o1"]
        self.assertEqual(
            llm.subscription_supported_openai_model_ids(ids),
            ["gpt-5.2-thinking", "gpt-5.2-instant"],
        )


if __name__ == "__main__":
    unittest.main()
