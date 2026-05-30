import requests
import json
import datetime
import subprocess
import os
import re
import shutil
import tempfile
import threading
from .config import API_KEYS
from .utils import log_event
from . import model_catalog

OPENAI_REASONING_MODEL_RE = re.compile(r"^o\d(?:[-.]|$)", re.IGNORECASE)
OPENAI_REASONING_EFFORTS = {"none", "minimal", "low", "medium", "high", "xhigh"}
GEMINI_MODELS_ENDPOINT = model_catalog.GEMINI_MODELS_ENDPOINT
CURATED_GEMINI_MODELS = model_catalog.CURATED_GEMINI_MODELS


def _openai_uses_responses_api(model):
    """Return True for OpenAI reasoning models that should use Responses."""
    model_id = (model or "").lower()
    return model_id.startswith("gpt-5") or bool(OPENAI_REASONING_MODEL_RE.match(model_id))


def _openai_model_supports_streaming(model):
    return not (model or "").lower().startswith("gpt-5.5-pro")


def _openai_reasoning_effort(model, thinking_level):
    level = (thinking_level or "").strip().lower()
    if level == "none" and (model or "").lower().startswith("gpt-5.5-pro"):
        return None
    if level in OPENAI_REASONING_EFFORTS:
        return level
    return None


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


def subscription_supported_openai_model_ids(model_ids):
    """Subset of *model_ids* usable on the ChatGPT subscription (Codex)."""
    return [m for m in model_ids if _map_openai_model_to_codex(m) is not None]


def _codex_on_path():
    return shutil.which("codex") is not None


def _codex_auth_path():
    return os.path.join(os.path.expanduser("~"), ".codex", "auth.json")


def _codex_logged_in():
    return os.path.isfile(_codex_auth_path())


def codex_available():
    """True when the Codex CLI is installed AND a ChatGPT login exists."""
    return _codex_on_path() and _codex_logged_in()


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


def _curated_gemini_models():
    return list(model_catalog.CURATED_GEMINI_MODELS)


def _normalize_gemini_model_name(name):
    return model_catalog._normalize_gemini_model_name(name)


def _gemini_model_methods(model):
    return model_catalog._gemini_model_methods(model)


def _gemini_supports_generate_content(model):
    return model_catalog._gemini_supports_generate_content(model)


def _gemini_model_name(model):
    return model_catalog._gemini_model_id(model)


def _gemini_models_from_entries(entries):
    models = [entry.model_id for entry in model_catalog.gemini_entries_from_models(entries)]
    return sorted(set(models), reverse=True)


def _fetch_gemini_models_over_rest(api_key):
    return [
        entry.model_id
        for entry in model_catalog.fetch_provider_models(
            "Gemini",
            api_key=api_key,
            request_get=requests.get,
        )
    ]


def _openai_user_content(user_prompt, file_contents):
    full_user_content = user_prompt
    if file_contents:
        full_user_content += "\n\n[ATTACHED FILES]:\n" + file_contents
    return full_user_content


def _openai_responses_input(history, full_user_content):
    input_items = []
    for msg in history or []:
        role = msg.get("role", "user")
        if role not in {"user", "assistant", "developer", "system"}:
            role = "user"
        input_items.append({"role": role, "content": msg.get("content", "")})
    input_items.append({"role": "user", "content": full_user_content})
    return input_items


def _extract_openai_responses_text(data):
    if isinstance(data.get("output_text"), str):
        return data["output_text"]

    text_parts = []
    for item in data.get("output", []):
        content = item.get("content", [])
        if isinstance(content, str):
            text_parts.append(content)
            continue

        for part in content:
            if not isinstance(part, dict):
                continue
            text = part.get("text")
            if isinstance(text, str) and part.get("type") in {"output_text", "text", None}:
                text_parts.append(text)

    return "".join(text_parts)


# Try to import PySide6 for worker classes (optional - needed for UI integration)
try:
    from PySide6.QtCore import QThread, Signal
    PYSIDE6_AVAILABLE = True
except ImportError:
    PYSIDE6_AVAILABLE = False
    QThread = object
    Signal = lambda *args: property(lambda self: None)

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None
    types = None

class LLMHandler:
    @staticmethod
    def create_cache(provider, model, content, system_instruction=None, ttl_minutes=60):
        """Creates a cache for Gemini and returns the cache name."""
        if provider != "Gemini":
            return None
            
        if not genai:
            raise ImportError("google.genai not installed")
            
        api_key = API_KEYS.get(provider)
        client = genai.Client(api_key=api_key)
        
        try:
            # Create config for cache
            cache_config = {
                "contents": [content],
                "ttl": f"{ttl_minutes*60}s"
            }
            if system_instruction:
                cache_config["system_instruction"] = system_instruction

            cache = client.caches.create(
                model=model,
                config=types.CreateCachedContentConfig(**cache_config)
            )
            return cache.name
        except Exception as e:
            print(f"Cache creation failed: {e}")
            return None

    @staticmethod
    def generate(provider, model, system_prompt, user_prompt, file_contents, settings, history=None, media_files=None):
        """
        Generic generator.
        If settings['stream'] is True, yields chunks (str).
        Otherwise returns full text (str).

        media_files: Optional list of file paths for audio/video that should be
                     uploaded via the Gemini Files API (Gemini provider only).

        history: Optional list of dicts [{'role': 'user'|'assistant', 'content': str}, ...]
        """
        temp = settings.get('temperature', 1.0)
        top_p = settings.get('top_p', 0.95)
        max_tokens = settings.get('max_tokens', -1)
        thinking_level = settings.get('thinking_level', "None")
        do_stream = settings.get('stream', False)
        cache_name = settings.get('cache_name', None)
        
        api_key = API_KEYS.get(provider)
        if not api_key and provider != "Claude":
            raise ValueError(f"API Key for {provider} not found.")

        if provider == "Gemini":
            if not genai:
                raise ImportError("google.genai not installed")

            client = genai.Client(api_key=api_key)
            api_level = thinking_level.upper()
            
            config_params = {
                "temperature": temp,
                "top_p": top_p,
            }
            
            # Only add system_instruction if NOT using cache (cache has it)
            if not cache_name:
                config_params["system_instruction"] = system_prompt

            if max_tokens > 0:
                config_params["max_output_tokens"] = max_tokens
            if thinking_level != "None" and ("gemini-3" in model.lower() or "thinking" in model.lower()):
                config_params["thinking_config"] = {
                    "include_thoughts": True,
                    "thinking_level": api_level
                }
            if cache_name:
                    config_params["cached_content"] = cache_name

            # Upload media files via Gemini Files API if provided
            uploaded_files = []
            if media_files:
                for media_path in media_files:
                    try:
                        log_event(f"Uploading media file to Gemini: {media_path}")
                        uploaded = client.files.upload(file=media_path)
                        uploaded_files.append(uploaded)
                        log_event(f"Uploaded: {uploaded.name} ({uploaded.mime_type})")
                    except Exception as upload_err:
                        log_event(f"Media upload failed for {media_path}: {upload_err}", "error")

            # Prompt construction
            if history:
                # Convert generic history to Gemini contents
                gemini_contents = []
                for h in history:
                    role = "user" if h['role'] == "user" else "model"
                    gemini_contents.append(types.Content(role=role, parts=[types.Part(text=h['content'])]))

                # Current message — build parts list with text + uploaded media
                full_prompt = user_prompt
                if file_contents and not cache_name:
                    full_prompt += "\n\n[ATTACHED FILES]:\n" + file_contents

                parts = [types.Part(text=full_prompt)]
                for uf in uploaded_files:
                    parts.append(types.Part.from_uri(file_uri=uf.uri, mime_type=uf.mime_type))
                gemini_contents.append(types.Content(role="user", parts=parts))

                final_contents = gemini_contents
            else:
                # Legacy/Simple mode
                full_prompt = user_prompt
                if file_contents and not cache_name:
                    full_prompt += "\n\n[ATTACHED FILES]:\n" + file_contents
                if uploaded_files:
                    # Mixed content: text + media files
                    final_contents = [full_prompt] + uploaded_files
                else:
                    final_contents = full_prompt
                
            # Helper to run generation
            def _run_gen(cfg):
                if do_stream:
                    return client.models.generate_content_stream(
                        model=model,
                        contents=final_contents,
                        config=types.GenerateContentConfig(**cfg)
                    )
                else:
                    return client.models.generate_content(
                        model=model,
                        contents=final_contents,
                        config=types.GenerateContentConfig(**cfg)
                    )

            try:
                log_event(f"Sending request to Gemini (Model: {model})")
                response = _run_gen(config_params)
                log_event("Received response from Gemini")
            except Exception as e:
                log_event(f"Initial Gemini request failed: {e}", "warning")
                # Fallback retry without thinking if failed
                if "thinking_config" in config_params:
                    log_event("Retrying without thinking_config...", "warning")
                    del config_params["thinking_config"]
                    try:
                        response = _run_gen(config_params)
                    except Exception as e2:
                        log_event(f"Fallback request failed: {e2}", "error")
                        raise
                else:
                    raise

            def _cleanup_uploads(client_ref, files_to_delete):
                for uf in files_to_delete:
                    try:
                        client_ref.files.delete(name=uf.name)
                    except Exception:
                        pass

            if do_stream:
                def stream_wrapper(client_ref, stream_resp, uploads):
                    try:
                        for chunk in stream_resp:
                            yield chunk.text
                    except Exception as e:
                        log_event(f"Stream error: {e}", "error")
                        yield f"[Stream Error: {e}]"
                    finally:
                        _cleanup_uploads(client_ref, uploads)
                return stream_wrapper(client, response, uploaded_files)
            else:
                try:
                    # Check for safety blocks or empty responses
                    if not response.candidates:
                            log_event("Gemini returned no candidates (blocked?)", "error")
                            return "[Error: The AI response was blocked or empty.]"

                    candidate = response.candidates[0]
                    finish_reason = getattr(candidate, 'finish_reason', None)

                    # Check for blocked/empty candidate
                    candidate_content = getattr(candidate, 'content', None)
                    candidate_parts = getattr(candidate_content, 'parts', None) if candidate_content else None
                    if not candidate_content or not candidate_parts or len(candidate_parts) == 0:
                            log_event(f"Gemini candidate has no content (finish_reason={finish_reason})", "error")
                            return f"[Error: The AI response was blocked or empty (reason: {finish_reason}).]"

                    # Try the safe .text accessor first, fall back to manual
                    try:
                            return response.text
                    except (ValueError, AttributeError):
                            return candidate_parts[0].text
                except Exception as e:
                    log_event(f"Error accessing response text: {e}", "error")
                    try:
                        log_event(f"Response dump: {response}", "info")
                    except: pass
                    raise
                finally:
                    _cleanup_uploads(client, uploaded_files)

        elif provider == "OpenAI":
            use_responses_api = _openai_uses_responses_api(model)
            request_stream = do_stream and _openai_model_supports_streaming(model)
            url = (
                "https://api.openai.com/v1/responses"
                if use_responses_api
                else "https://api.openai.com/v1/chat/completions"
            )
            headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

            full_user_content = _openai_user_content(user_prompt, file_contents)

            if use_responses_api:
                payload = {
                    "model": model,
                    "input": _openai_responses_input(history, full_user_content),
                }
                if system_prompt:
                    payload["instructions"] = system_prompt
                if max_tokens > 0:
                    payload["max_output_tokens"] = max_tokens
                effort = _openai_reasoning_effort(model, thinking_level)
                if effort:
                    payload["reasoning"] = {"effort": effort}
            else:
                messages = [{"role": "system", "content": system_prompt}]
                if history:
                    messages.extend(history)
                messages.append({"role": "user", "content": full_user_content})

                payload = {
                    "model": model, "messages": messages,
                    "temperature": temp, "top_p": top_p
                }
                if max_tokens > 0: payload["max_tokens"] = max_tokens

            if do_stream:
                # Streaming mode
                if not request_stream:
                    api_name = "OpenAI Responses" if use_responses_api else "OpenAI"
                    log_event(
                        f"Sending non-streaming request to {api_name} "
                        f"(Model: {model}, streaming not supported)"
                    )
                    resp = requests.post(url, headers=headers, json=payload)
                    if resp.status_code != 200:
                        raise Exception(f"OpenAI Error {resp.status_code}: {resp.text}")

                    data = resp.json()
                    if use_responses_api:
                        text = _extract_openai_responses_text(data)
                    else:
                        text = data['choices'][0]['message']['content']
                    return iter([text] if text else [])

                payload["stream"] = True
                api_name = "OpenAI Responses" if use_responses_api else "OpenAI"
                log_event(f"Sending streaming request to {api_name} (Model: {model})")
                resp = requests.post(url, headers=headers, json=payload, stream=True)

                if resp.status_code != 200:
                    raise Exception(f"OpenAI Error {resp.status_code}: {resp.text}")

                def openai_stream():
                    try:
                        for line in resp.iter_lines():
                            if line:
                                line_str = line.decode('utf-8')
                                if line_str.startswith('data: '):
                                    data_str = line_str[6:]
                                    if data_str.strip() == '[DONE]':
                                        break
                                    try:
                                        data = json.loads(data_str)
                                        if use_responses_api:
                                            event_type = data.get("type")
                                            if event_type == "response.output_text.delta":
                                                delta = data.get("delta", "")
                                                if delta:
                                                    yield delta
                                            elif event_type in {"error", "response.failed"}:
                                                err = data.get("error") or data.get("response", {}).get("error") or data
                                                raise Exception(err)
                                        else:
                                            delta = data.get('choices', [{}])[0].get('delta', {})
                                            if 'content' in delta and delta['content']:
                                                yield delta['content']
                                    except json.JSONDecodeError:
                                        pass
                    except Exception as e:
                        log_event(f"OpenAI stream error: {e}", "error")
                        yield f"[Stream Error: {e}]"

                return openai_stream()
            else:
                # Non-streaming mode
                api_name = "OpenAI Responses" if use_responses_api else "OpenAI"
                log_event(f"Sending request to {api_name} (Model: {model})")
                resp = requests.post(url, headers=headers, json=payload)
                if resp.status_code == 200:
                    data = resp.json()
                    if use_responses_api:
                        return _extract_openai_responses_text(data)
                    return data['choices'][0]['message']['content']
                raise Exception(f"OpenAI Error {resp.status_code}: {resp.text}")

        elif provider == "Claude":
            # Route through Claude Code CLI to use Max subscription
            full_prompt_parts = []
            if history:
                for msg in history:
                    role_label = "User" if msg['role'] == 'user' else "Assistant"
                    full_prompt_parts.append(f"[{role_label}]: {msg['content']}")
                full_prompt_parts.append("[User]: ")

            full_prompt_parts.append(user_prompt)
            if file_contents:
                full_prompt_parts.append("\n\n[ATTACHED FILES]:\n" + file_contents)

            combined_prompt = "\n".join(full_prompt_parts)

            # Environment: remove CLAUDECODE to allow nested invocation,
            # and remove ANTHROPIC_API_KEY so CLI uses OAuth (Max subscription)
            # instead of the (possibly empty) API key
            env = os.environ.copy()
            env.pop("CLAUDECODE", None)
            env.pop("ANTHROPIC_API_KEY", None)

            # Map model aliases for CLI compatibility
            model_arg = model
            if "opus" in model.lower():
                model_arg = "opus"
            elif "sonnet" in model.lower():
                model_arg = "sonnet"
            elif "haiku" in model.lower():
                model_arg = "haiku"

            log_event(f"Sending request to Claude Code CLI (Model: {model_arg}, stream={do_stream})")

            # Base CLI args: pipe mode, disable all tools so it acts as
            # pure text generation (not an agent that creates files)
            base_args = ["claude", "-p", "--model", model_arg,
                         "--allowedTools", "none",
                         "--no-session-persistence"]
            if system_prompt:
                base_args.extend(["--system-prompt", system_prompt])

            if do_stream:
                cmd = base_args + ["--output-format", "stream-json", "--verbose",
                                   "--include-partial-messages"]

                def claude_cli_stream():
                    try:
                        creationflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                        proc = subprocess.Popen(
                            cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, env=env, text=True,
                            encoding="utf-8", errors="replace",
                            creationflags=creationflags
                        )
                        proc.stdin.write(combined_prompt)
                        proc.stdin.close()

                        got_stream_tokens = False
                        for line in proc.stdout:
                            line = line.strip()
                            if not line:
                                continue
                            try:
                                event = json.loads(line)
                                evt_type = event.get("type", "")

                                if evt_type == "stream_event":
                                    # Token-by-token streaming
                                    inner = event.get("event", {})
                                    if inner.get("type") == "content_block_delta":
                                        delta = inner.get("delta", {})
                                        if delta.get("type") == "text_delta":
                                            text = delta.get("text", "")
                                            if text:
                                                got_stream_tokens = True
                                                yield text
                                elif evt_type == "result":
                                    # Only use result text if we never got stream tokens
                                    if not got_stream_tokens:
                                        result_text = event.get("result", "")
                                        if result_text:
                                            yield result_text
                                    break
                                # Skip 'assistant' events — they duplicate stream tokens
                            except json.JSONDecodeError:
                                continue

                        proc.wait()
                        if proc.returncode != 0:
                            stderr_out = proc.stderr.read()
                            if stderr_out:
                                log_event(f"Claude CLI stderr: {stderr_out}", "warning")
                    except Exception as e:
                        log_event(f"Claude CLI stream error: {e}", "error")
                        yield f"[Stream Error: {e}]"

                return claude_cli_stream()
            else:
                cmd = base_args + ["--output-format", "text"]

                try:
                    creationflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                    result = subprocess.run(
                        cmd, input=combined_prompt, capture_output=True,
                        text=True, encoding="utf-8", errors="replace",
                        env=env, timeout=300,
                        creationflags=creationflags
                    )
                    if result.returncode == 0:
                        return result.stdout.strip()
                    raise Exception(f"Claude CLI Error (exit {result.returncode}): {result.stderr}")
                except subprocess.TimeoutExpired:
                    raise Exception("Claude CLI timed out after 300 seconds")
        
        return "Provider not implemented."

# Worker classes - only available when PySide6 is installed
if PYSIDE6_AVAILABLE:
    class LLMWorker(QThread):
        finished = Signal(str)
        new_token = Signal(str)  # For streaming
        error = Signal(str)

        def __init__(self, provider, model, system, user, files, settings, history=None, media_files=None):
            super().__init__()
            self.provider = provider
            self.model = model
            self.system = system
            self.user = user
            self.files = files
            self.settings = settings
            self.history = history
            self.media_files = media_files
            self._stop_requested = False

        def request_stop(self):
            """Request the worker to stop processing."""
            self._stop_requested = True

        def run(self):
            try:
                result = LLMHandler.generate(
                    self.provider, self.model, self.system,
                    self.user, self.files, self.settings,
                    history=self.history, media_files=self.media_files
                )

                if self.settings.get('stream', False):
                    full_text = ""
                    for chunk in result:
                        if self._stop_requested:
                            log_event("Stream stopped by user request")
                            break
                        if chunk:
                            full_text += chunk
                            self.new_token.emit(chunk)
                    self.finished.emit(full_text)
                else:
                    self.finished.emit(result)

            except Exception as e:
                if not self._stop_requested:
                    self.error.emit(str(e))

    class ModelFetcher(QThread):
        finished = Signal(str, list)
        error = Signal(str)

        def __init__(self, provider, api_key):
            super().__init__()
            self.provider = provider
            self.api_key = api_key

        def run(self):
            try:
                entries = model_catalog.list_provider_models(
                    self.provider,
                    api_key=self.api_key,
                    request_get=requests.get,
                    allow_fallback=True,
                )
                models = [entry.model_id for entry in entries]

                if self.provider == "Gemini":
                    models.sort(reverse=True)

                self.finished.emit(self.provider, models)

            except Exception as e:
                self.error.emit(str(e))
else:
    # Placeholder classes when PySide6 is not available
    LLMWorker = None
    ModelFetcher = None
