# OpenAI via ChatGPT Subscription (Codex CLI routing) — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Route OpenAI model calls through the user's ChatGPT subscription via the official Codex CLI, app-wide, with automatic fallback to the existing `OPENAI_API_KEY` path.

**Architecture:** Mirror the existing Claude branch in `icharlotte_core/llm.py` (which shells out to the `claude` CLI for the Max subscription). Add an OpenAI analog that shells out to `codex exec`. All new logic lives inside `LLMHandler.generate()` so every caller (chat tab, 28 agents, Word assistant) inherits it. The existing `api.openai.com` code is left intact as the fallback.

**Tech Stack:** Python 3, `subprocess`, OpenAI `codex` CLI (Node/npm), `unittest` + `unittest.mock`.

**Spec:** `docs/superpowers/specs/2026-05-29-openai-chatgpt-subscription-design.md`

**Key files:**
- Modify: `icharlotte_core/llm.py` (new module-level helpers + hook in `generate()`)
- Modify: `config/llm_preferences.json` (add `openai_use_subscription` flag)
- Create: `tests/test_openai_codex_subscription.py`

---

## Task 1: Prerequisite — install & log in to Codex CLI (manual, one-time)

**Files:** none (environment setup)

- [ ] **Step 1: Install the Codex CLI**

Run: `npm install -g @openai/codex`
Expected: installs without error (npm 10.9.2 / node v22.14.0 already verified present).

- [ ] **Step 2: Verify it is on PATH**

Run (PowerShell): `(Get-Command codex).Source; codex --version`
Expected: a path and a version string.

- [ ] **Step 3: Log in with ChatGPT**

Run: `codex login`
Expected: opens a browser; after signing in with the ChatGPT account, the terminal confirms login. An auth file appears at `~/.codex/auth.json`.

- [ ] **Step 4: Confirm the exec flags this plan relies on**

Run: `codex exec --help`
Confirm these flags exist (names below are used throughout this plan): `--sandbox`, `--skip-git-repo-check`, `--output-last-message`, `-m`/`--model`.
If the installed version names any of them differently, note the exact spelling — it is used verbatim in Task 5's `_build_codex_command`. Adjust that one function accordingly.

- [ ] **Step 5: One-shot smoke check (proves subscription billing works)**

Run: `codex exec --sandbox read-only --skip-git-repo-check -m gpt-5.2-codex "Reply with the single word: ok"`
Expected: prints `ok` (or similar), with no API-key configured — confirming it used the ChatGPT login.

No commit (environment only).

---

## Task 2: Model mapping `_map_openai_model_to_codex`

Maps an app OpenAI model id to a Codex model name, or `None` when the model is not available on the subscription (caller falls back to the API key).

**Files:**
- Modify: `icharlotte_core/llm.py` (imports + new function near the other `_openai_*` helpers, after line ~34)
- Test: `tests/test_openai_codex_subscription.py`

- [ ] **Step 1: Add imports**

In `icharlotte_core/llm.py`, add `shutil` and `tempfile` to the import block at the top (after `import threading` on line 7):

```python
import shutil
import tempfile
```

- [ ] **Step 2: Write the failing test**

Create `tests/test_openai_codex_subscription.py`:

```python
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


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 3: Run test to verify it fails**

Run: `python -m pytest tests/test_openai_codex_subscription.py::MapModelTests -v`
Expected: FAIL with `AttributeError: module 'icharlotte_core.llm' has no attribute '_map_openai_model_to_codex'`.

- [ ] **Step 4: Write minimal implementation**

In `icharlotte_core/llm.py`, after `_openai_reasoning_effort` (line ~34):

```python
# App OpenAI model id -> Codex model name for the ChatGPT subscription.
# Anything that maps to None is not available via Codex; the caller falls back
# to the OPENAI_API_KEY path.
_CODEX_MODEL_OVERRIDES = {
    "gpt-5.2-thinking": "gpt-5.2-codex",
    "gpt-5.2-instant": "gpt-5.2-codex",
    "gpt-5.1-thinking": "gpt-5.1-codex",
    "gpt-5.1-instant": "gpt-5.1-codex",
}


def _map_openai_model_to_codex(model):
    """Map an app OpenAI model id to a Codex model name, or None if the model is
    not available on the ChatGPT subscription."""
    if not model:
        return None
    model_id = model.strip().lower()
    if model_id in _CODEX_MODEL_OVERRIDES:
        return _CODEX_MODEL_OVERRIDES[model_id]
    if model_id.startswith("gpt-5"):
        return model_id if "codex" in model_id else "gpt-5-codex"
    return None
```

- [ ] **Step 5: Run test to verify it passes**

Run: `python -m pytest tests/test_openai_codex_subscription.py::MapModelTests -v`
Expected: PASS (7 tests).

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/llm.py tests/test_openai_codex_subscription.py
git commit -m "feat(llm): map OpenAI model ids to Codex subscription models"
```

---

## Task 3: Detection helpers `codex_available`

**Files:**
- Modify: `icharlotte_core/llm.py` (new functions after Task 2's code)
- Test: `tests/test_openai_codex_subscription.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_openai_codex_subscription.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_openai_codex_subscription.py::CodexAvailableTests -v`
Expected: FAIL with `AttributeError: ... has no attribute 'codex_available'`.

- [ ] **Step 3: Write minimal implementation**

In `icharlotte_core/llm.py`, after the model-map code:

```python
def _codex_on_path():
    return shutil.which("codex") is not None


def _codex_auth_path():
    return os.path.join(os.path.expanduser("~"), ".codex", "auth.json")


def _codex_logged_in():
    return os.path.isfile(_codex_auth_path())


def codex_available():
    """True when the Codex CLI is installed AND a ChatGPT login exists."""
    return _codex_on_path() and _codex_logged_in()
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_openai_codex_subscription.py::CodexAvailableTests -v`
Expected: PASS (3 tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/llm.py tests/test_openai_codex_subscription.py
git commit -m "feat(llm): detect Codex CLI install + ChatGPT login state"
```

---

## Task 4: Enablement flag `openai_subscription_enabled`

**Files:**
- Modify: `icharlotte_core/llm.py` (new functions)
- Modify: `config/llm_preferences.json` (add top-level flag)
- Test: `tests/test_openai_codex_subscription.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_openai_codex_subscription.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_openai_codex_subscription.py::SubscriptionEnabledTests -v`
Expected: FAIL with `AttributeError: ... has no attribute '_subscription_prefs_path'`.

- [ ] **Step 3: Write minimal implementation**

In `icharlotte_core/llm.py`, after the detection helpers:

```python
def _subscription_prefs_path():
    """Path to llm_preferences.json (lazy import avoids circular import)."""
    try:
        from .llm_config import CONFIG_FILE
        return CONFIG_FILE
    except Exception:
        return os.path.join(os.getcwd(), "config", "llm_preferences.json")


def openai_subscription_enabled():
    """Whether OpenAI calls should try the ChatGPT subscription first.
    Reads top-level 'openai_use_subscription' from llm_preferences.json
    (default True)."""
    try:
        with open(_subscription_prefs_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
        return bool(data.get("openai_use_subscription", True))
    except Exception:
        return True
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_openai_codex_subscription.py::SubscriptionEnabledTests -v`
Expected: PASS (4 tests).

- [ ] **Step 5: Add the flag to the real prefs file**

In `config/llm_preferences.json`, add a top-level key right after `"version": "2.1",` (line 2):

```json
  "openai_use_subscription": true,
```

(Default True is already the code behavior; this makes the setting discoverable/editable.)

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/llm.py tests/test_openai_codex_subscription.py config/llm_preferences.json
git commit -m "feat(llm): openai_use_subscription preference flag"
```

---

## Task 5: Core generator `_generate_openai_codex_cli`

Builds the constrained `codex exec` command, runs it in a neutral temp cwd, and returns the model's final message (read from `--output-last-message`, which avoids CLI banner noise).

**Files:**
- Modify: `icharlotte_core/llm.py` (new functions)
- Test: `tests/test_openai_codex_subscription.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_openai_codex_subscription.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_openai_codex_subscription.py::CodexGenerateTests -v`
Expected: FAIL with `AttributeError: ... has no attribute '_generate_openai_codex_cli'`.

- [ ] **Step 3: Write minimal implementation**

In `icharlotte_core/llm.py`, after the enablement helpers:

```python
_CODEX_PERSONA_PREAMBLE = (
    "You are a general-purpose assistant answering the user's request directly. "
    "Do NOT write code, run commands, create or modify files, or use tools "
    "unless the user explicitly asks for that. Respond in plain prose."
)


def _build_codex_command(codex_model, last_message_path):
    return [
        "codex", "exec",
        "--sandbox", "read-only",
        "--skip-git-repo-check",
        "--output-last-message", last_message_path,
        "-m", codex_model,
    ]


def _build_codex_prompt(system_prompt, user_prompt, file_contents, history):
    parts = [_CODEX_PERSONA_PREAMBLE]
    if system_prompt:
        parts.append(f"[SYSTEM INSTRUCTIONS]:\n{system_prompt}")
    if history:
        for msg in history:
            role_label = "User" if msg.get("role") == "user" else "Assistant"
            parts.append(f"[{role_label}]: {msg.get('content', '')}")
    parts.append(f"[User]: {user_prompt}")
    if file_contents:
        parts.append("\n\n[ATTACHED FILES]:\n" + file_contents)
    return "\n\n".join(parts)


def _generate_openai_codex_cli(model, system_prompt, user_prompt, file_contents,
                               history, do_stream, timeout=300):
    """Generate via the Codex CLI (ChatGPT subscription). Raises on unsupported
    model or CLI failure so the caller can fall back to the API key."""
    codex_model = _map_openai_model_to_codex(model)
    if codex_model is None:
        raise ValueError(f"Model '{model}' is not available via ChatGPT subscription")

    prompt = _build_codex_prompt(system_prompt, user_prompt, file_contents, history)
    cwd = tempfile.gettempdir()  # never a client folder
    env = os.environ.copy()
    env.pop("CLAUDECODE", None)
    creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

    fd, last_msg_path = tempfile.mkstemp(suffix=".txt", prefix="codex_out_")
    os.close(fd)
    try:
        cmd = _build_codex_command(codex_model, last_msg_path)
        log_event(f"Sending request to Codex CLI (Model: {codex_model})")
        result = subprocess.run(
            cmd, input=prompt, capture_output=True, text=True,
            encoding="utf-8", errors="replace", cwd=cwd, env=env,
            timeout=timeout, creationflags=creationflags,
        )
        if result.returncode != 0:
            raise Exception(f"Codex CLI Error (exit {result.returncode}): {result.stderr}")
        try:
            with open(last_msg_path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read().strip()
        except OSError:
            text = (result.stdout or "").strip()
    finally:
        try:
            os.remove(last_msg_path)
        except OSError:
            pass

    return iter([text] if text else []) if do_stream else text
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_openai_codex_subscription.py::CodexGenerateTests -v`
Expected: PASS (5 tests).

- [ ] **Step 5: Commit**

```bash
git add icharlotte_core/llm.py tests/test_openai_codex_subscription.py
git commit -m "feat(llm): _generate_openai_codex_cli core generator"
```

---

## Task 6: Wire into `generate()` (routing + api-key guard fix)

**Files:**
- Modify: `icharlotte_core/llm.py` (lines ~178-180 and the OpenAI branch at line ~332)
- Test: `tests/test_openai_codex_subscription.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_openai_codex_subscription.py`:

```python
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
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_openai_codex_subscription.py::RoutingTests -v`
Expected: FAIL — `test_no_api_key_with_subscription_does_not_raise` raises `ValueError: API Key for OpenAI not found`, and the codex tests fail because routing isn't wired yet.

- [ ] **Step 3: Fix the api-key guard**

In `icharlotte_core/llm.py`, replace the guard at lines ~178-180:

```python
        api_key = API_KEYS.get(provider)
        if not api_key and provider != "Claude":
            raise ValueError(f"API Key for {provider} not found.")
```

with:

```python
        api_key = API_KEYS.get(provider)
        if not api_key and provider != "Claude":
            _openai_sub = (
                provider == "OpenAI"
                and openai_subscription_enabled()
                and codex_available()
            )
            if not _openai_sub:
                raise ValueError(f"API Key for {provider} not found.")
```

- [ ] **Step 4: Add the routing hook**

In `icharlotte_core/llm.py`, find the OpenAI branch (line ~332, `elif provider == "OpenAI":`). Insert the subscription attempt as the FIRST thing inside the branch, before `use_responses_api = ...`:

```python
        elif provider == "OpenAI":
            if openai_subscription_enabled() and codex_available():
                if _map_openai_model_to_codex(model) is not None:
                    try:
                        return _generate_openai_codex_cli(
                            model, system_prompt, user_prompt, file_contents,
                            history, do_stream,
                        )
                    except Exception as e:
                        log_event(
                            f"Codex subscription path failed, falling back to "
                            f"OpenAI API key: {e}", "warning")
                else:
                    log_event(
                        f"Model '{model}' not available on ChatGPT subscription; "
                        f"using OpenAI API key", "info")

            use_responses_api = _openai_uses_responses_api(model)
            # ... rest of existing OpenAI branch unchanged ...
```

- [ ] **Step 5: Run test to verify it passes**

Run: `python -m pytest tests/test_openai_codex_subscription.py::RoutingTests -v`
Expected: PASS (4 tests).

- [ ] **Step 6: Run the whole new test file**

Run: `python -m pytest tests/test_openai_codex_subscription.py -v`
Expected: PASS (all classes).

- [ ] **Step 7: Commit**

```bash
git add icharlotte_core/llm.py tests/test_openai_codex_subscription.py
git commit -m "feat(llm): route OpenAI through ChatGPT subscription with API-key fallback"
```

---

## Task 7 (optional polish): filter model picker to Codex-supported models

Correctness does not depend on this — unsupported models already fall back to the API key. This only hides non-subscription models from the chat picker when the subscription is active.

**Files:**
- Modify: `icharlotte_core/llm.py` (new helper)
- Test: `tests/test_openai_codex_subscription.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_openai_codex_subscription.py`:

```python
class ModelFilterTests(unittest.TestCase):
    def test_filters_to_codex_supported(self):
        ids = ["gpt-5.2-thinking", "gpt-5.2-instant", "gpt-4o", "o1"]
        self.assertEqual(
            llm.subscription_supported_openai_model_ids(ids),
            ["gpt-5.2-thinking", "gpt-5.2-instant"],
        )
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest tests/test_openai_codex_subscription.py::ModelFilterTests -v`
Expected: FAIL with `AttributeError: ... has no attribute 'subscription_supported_openai_model_ids'`.

- [ ] **Step 3: Write minimal implementation**

In `icharlotte_core/llm.py`, after `_map_openai_model_to_codex`:

```python
def subscription_supported_openai_model_ids(model_ids):
    """Subset of *model_ids* usable on the ChatGPT subscription (Codex)."""
    return [m for m in model_ids if _map_openai_model_to_codex(m) is not None]
```

- [ ] **Step 4: Run test to verify it passes**

Run: `python -m pytest tests/test_openai_codex_subscription.py::ModelFilterTests -v`
Expected: PASS.

- [ ] **Step 5: Wire into the chat model picker (optional)**

Where the chat settings dropdown populates OpenAI models (search: `grep -rn "OpenAI" icharlotte_core/ui/chat_dialogs.py icharlotte_core/ui/tabs.py`), when `llm.openai_subscription_enabled() and llm.codex_available()`, pass the OpenAI id list through `llm.subscription_supported_openai_model_ids(...)` before display. Keep it behind that condition so API-key users see the full list.

- [ ] **Step 6: Commit**

```bash
git add icharlotte_core/llm.py tests/test_openai_codex_subscription.py
git commit -m "feat(ui): filter OpenAI model picker to subscription models"
```

---

## Task 8: Full verification

**Files:** none (verification)

- [ ] **Step 1: Run the full test suite**

Run: `python -m pytest tests/ -q`
Expected: no new failures versus the pre-existing baseline. (Note any pre-existing failures unrelated to this change.)

- [ ] **Step 2: Manual smoke — chat tab uses subscription**

Launch: `python iCharlotte.py`. In the Chat tab, select an OpenAI gpt-5.x model and send a message.
Verify: a normal response returns. Check `icharlotte_Activity.log` for `Sending request to Codex CLI (Model: gpt-5.2-codex)` — confirming the subscription path, not the API.

- [ ] **Step 3: Manual smoke — agent uses subscription**

Run one agent that uses OpenAI (or temporarily point a task's `model_sequence` at an OpenAI gpt-5 model). Confirm the same `Codex CLI` log line appears.

- [ ] **Step 4: Manual smoke — fallback works**

In the Chat tab, select `gpt-4o`. Send a message. Verify the log shows `not available on ChatGPT subscription; using OpenAI API key` and the response still returns via the API key.

- [ ] **Step 5: Update memory**

Add a one-line entry to `MEMORY.md` index pointing to a new topic note `openai_chatgpt_subscription.md` summarizing: Codex CLI routing mirrors the Claude branch; flag `openai_use_subscription` in llm_preferences.json; gpt-5.x only; fallback to API key.

---

## Self-Review notes

- **Spec coverage:** architecture/hook (Task 6), Codex invocation + neutral cwd + persona neutralization (Task 5), streaming-as-buffered-emit (Task 5, `do_stream` → single-chunk iterator, matching the spec's "buffer and emit once" fallback), enablement flag (Task 4), model set + mapping (Task 2), model picker filter (Task 7), fallback/error handling (Tasks 5+6), testing (Tasks 2-8). Prerequisite install/login (Task 1).
- **Deferred vs spec:** JSONL token-by-token streaming is intentionally deferred — v1 buffers and emits once (spec-allowed) for robustness. Full UI picker filtering is Task 7 (optional) since fallback makes it non-essential.
- **Type consistency:** `_map_openai_model_to_codex` (str|None), `codex_available`/`openai_subscription_enabled` (bool), `_generate_openai_codex_cli(model, system_prompt, user_prompt, file_contents, history, do_stream, timeout=300)` — signature matches the call site in Task 6.
- **Flag-name risk:** Task 1 Step 4 verifies the exact `codex exec` flag spellings; only `_build_codex_command` (Task 5) consumes them.
