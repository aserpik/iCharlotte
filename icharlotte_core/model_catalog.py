"""Dynamic model discovery for iCharlotte LLM providers."""

from __future__ import annotations

from dataclasses import dataclass
import re
from typing import Callable, Dict, Iterable, List, Optional

import requests


PROVIDER_ORDER = ("Gemini", "OpenAI", "Claude")

GEMINI_MODELS_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models"
OPENAI_MODELS_ENDPOINT = "https://api.openai.com/v1/models"
ANTHROPIC_MODELS_ENDPOINT = "https://api.anthropic.com/v1/models"
ANTHROPIC_VERSION = "2023-06-01"

REQUEST_TIMEOUT_SECONDS = 20


@dataclass(frozen=True)
class ModelCatalogEntry:
    provider: str
    model_id: str
    display_name: str


FALLBACK_MODELS: Dict[str, List[ModelCatalogEntry]] = {
    "Gemini": [
        ModelCatalogEntry("Gemini", "gemini-3.1-pro-preview", "Gemini 3.1 Pro Preview"),
        ModelCatalogEntry("Gemini", "gemini-3.5-flash", "Gemini 3.5 Flash"),
        ModelCatalogEntry("Gemini", "gemini-3.1-flash-lite", "Gemini 3.1 Flash-Lite"),
    ],
    "Claude": [
        ModelCatalogEntry("Claude", "claude-opus-4-6", "Claude Opus 4.6"),
        ModelCatalogEntry("Claude", "claude-sonnet-4-6", "Claude Sonnet 4.6"),
        ModelCatalogEntry("Claude", "claude-haiku-4-5-20251001", "Claude Haiku 4.5"),
        ModelCatalogEntry("Claude", "claude-sonnet-4-20250514", "Claude Sonnet 4"),
        ModelCatalogEntry("Claude", "claude-haiku-4-20250514", "Claude Haiku 4"),
    ],
    "OpenAI": [
        ModelCatalogEntry("OpenAI", "gpt-5.2-thinking", "GPT-5.2 Thinking"),
        ModelCatalogEntry("OpenAI", "gpt-5.2-instant", "GPT-5.2 Instant"),
        ModelCatalogEntry("OpenAI", "gpt-5.1-thinking", "GPT-5.1 Thinking"),
        ModelCatalogEntry("OpenAI", "gpt-5.1-instant", "GPT-5.1 Instant"),
        ModelCatalogEntry("OpenAI", "gpt-4o", "GPT-4o"),
        ModelCatalogEntry("OpenAI", "gpt-4o-mini", "GPT-4o Mini"),
        ModelCatalogEntry("OpenAI", "gpt-4-turbo", "GPT-4 Turbo"),
        ModelCatalogEntry("OpenAI", "o1", "O1"),
        ModelCatalogEntry("OpenAI", "o1-mini", "O1 Mini"),
    ],
}

CURATED_GEMINI_MODELS = [entry.model_id for entry in FALLBACK_MODELS["Gemini"]]

OPENAI_GENERATIVE_ID_RE = re.compile(r"^(?:gpt-|chatgpt-|o\d(?:[-.]|$))", re.IGNORECASE)
OPENAI_NON_TEXT_MARKERS = (
    "audio",
    "dall-e",
    "embedding",
    "image",
    "moderation",
    "realtime",
    "speech",
    "transcribe",
    "tts",
    "whisper",
)


class ModelCatalogError(Exception):
    """Raised when live model discovery fails and fallback is disabled."""


def fallback_models_for_provider(provider: str) -> List[ModelCatalogEntry]:
    """Return the local fallback model list for *provider*."""
    return list(FALLBACK_MODELS.get(provider, []))


def fallback_models_by_provider() -> Dict[str, List[ModelCatalogEntry]]:
    return {provider: fallback_models_for_provider(provider) for provider in PROVIDER_ORDER}


def as_dialog_models(entries: Iterable[ModelCatalogEntry]) -> List[tuple]:
    """Convert entries to ``(model_id, display_name)`` tuples for settings UIs."""
    return [(entry.model_id, entry.display_name) for entry in entries]


def dialog_models_by_provider(
    models_by_provider: Optional[Dict[str, Iterable[ModelCatalogEntry]]] = None,
) -> Dict[str, List[tuple]]:
    source = models_by_provider or fallback_models_by_provider()
    return {provider: as_dialog_models(entries) for provider, entries in source.items()}


def flatten_word_models(
    models_by_provider: Optional[Dict[str, Iterable[ModelCatalogEntry]]] = None,
) -> List[tuple]:
    """Convert entries to ``(display_name, provider, model_id)`` for Word popup."""
    source = models_by_provider or fallback_models_by_provider()
    flattened = []
    for provider in PROVIDER_ORDER:
        for entry in source.get(provider, []):
            flattened.append((entry.display_name, entry.provider, entry.model_id))
    return flattened


def display_name_for_model(provider: str, model_id: str, display_name: str = "") -> str:
    if display_name:
        return display_name

    fallback = {entry.model_id: entry.display_name for entry in fallback_models_for_provider(provider)}
    if model_id in fallback:
        return fallback[model_id]

    if provider == "OpenAI":
        return model_id.upper() if model_id.startswith("o") and len(model_id) <= 3 else model_id

    if provider == "Claude":
        return model_id.replace("-", " ").title()

    if provider == "Gemini":
        return model_id.replace("-", " ").title()

    return model_id


def _require_api_key(provider: str, api_key: Optional[str]) -> str:
    if not api_key:
        raise ModelCatalogError(f"API key for {provider} is not configured")
    return api_key


def _raise_for_status(response, provider: str):
    if response.status_code != 200:
        raise ModelCatalogError(f"{provider} model list error: {response.status_code} {response.text}")


def _normalize_gemini_model_name(name: str) -> str:
    name = (name or "").strip()
    if name.startswith("models/"):
        return name[7:]
    return name


def _gemini_model_methods(model) -> Optional[Iterable[str]]:
    if isinstance(model, dict):
        return (
            model.get("supportedGenerationMethods")
            or model.get("supported_actions")
            or model.get("supportedActions")
        )

    for attr in ("supported_actions", "supported_generation_methods", "supportedGenerationMethods"):
        methods = getattr(model, attr, None)
        if methods is not None:
            return methods
    return None


def _gemini_supports_generate_content(model) -> bool:
    methods = _gemini_model_methods(model)
    if methods is None:
        return True
    return "generateContent" in methods


def _gemini_model_id(model) -> str:
    if isinstance(model, dict):
        return _normalize_gemini_model_name(model.get("name") or model.get("baseModelId"))
    return _normalize_gemini_model_name(getattr(model, "name", None) or str(model))


def _gemini_display_name(model, model_id: str) -> str:
    if isinstance(model, dict):
        return model.get("displayName") or display_name_for_model("Gemini", model_id)
    return getattr(model, "display_name", None) or display_name_for_model("Gemini", model_id)


def gemini_entries_from_models(models: Iterable) -> List[ModelCatalogEntry]:
    entries = []
    seen = set()
    for model in models:
        model_id = _gemini_model_id(model)
        if not model_id or "gemini" not in model_id.lower():
            continue
        if not _gemini_supports_generate_content(model):
            continue
        if model_id in seen:
            continue
        seen.add(model_id)
        entries.append(
            ModelCatalogEntry(
                provider="Gemini",
                model_id=model_id,
                display_name=_gemini_display_name(model, model_id),
            )
        )
    return entries


def _fetch_gemini_models(api_key: str, request_get: Callable) -> List[ModelCatalogEntry]:
    api_key = _require_api_key("Gemini", api_key)
    entries: List[ModelCatalogEntry] = []
    page_token = None

    while True:
        params = {"key": api_key, "pageSize": 1000}
        if page_token:
            params["pageToken"] = page_token

        response = request_get(GEMINI_MODELS_ENDPOINT, params=params, timeout=REQUEST_TIMEOUT_SECONDS)
        _raise_for_status(response, "Gemini")
        payload = response.json()
        entries.extend(gemini_entries_from_models(payload.get("models", [])))

        page_token = payload.get("nextPageToken")
        if not page_token:
            break

    return _dedupe_entries(entries)


def _is_openai_generation_model(model_id: str) -> bool:
    lowered = (model_id or "").lower()
    if not OPENAI_GENERATIVE_ID_RE.match(lowered):
        return False
    return not any(marker in lowered for marker in OPENAI_NON_TEXT_MARKERS)


def _fetch_openai_models(api_key: str, request_get: Callable) -> List[ModelCatalogEntry]:
    api_key = _require_api_key("OpenAI", api_key)
    headers = {"Authorization": f"Bearer {api_key}"}
    response = request_get(OPENAI_MODELS_ENDPOINT, headers=headers, timeout=REQUEST_TIMEOUT_SECONDS)
    _raise_for_status(response, "OpenAI")

    models = response.json().get("data", [])
    models = [model for model in models if _is_openai_generation_model(model.get("id", ""))]
    models.sort(key=lambda model: (model.get("created") or 0, model.get("id") or ""), reverse=True)
    return [
        ModelCatalogEntry(
            provider="OpenAI",
            model_id=model["id"],
            display_name=display_name_for_model("OpenAI", model["id"]),
        )
        for model in models
    ]


def _fetch_claude_models(api_key: str, request_get: Callable) -> List[ModelCatalogEntry]:
    api_key = _require_api_key("Claude", api_key)
    headers = {
        "x-api-key": api_key,
        "anthropic-version": ANTHROPIC_VERSION,
    }
    entries: List[ModelCatalogEntry] = []
    after_id = None

    while True:
        params = {"limit": 1000}
        if after_id:
            params["after_id"] = after_id

        response = request_get(
            ANTHROPIC_MODELS_ENDPOINT,
            headers=headers,
            params=params,
            timeout=REQUEST_TIMEOUT_SECONDS,
        )
        _raise_for_status(response, "Claude")
        payload = response.json()

        for model in payload.get("data", []):
            model_id = model.get("id", "")
            if model.get("type") == "model" and model_id.startswith("claude-"):
                entries.append(
                    ModelCatalogEntry(
                        provider="Claude",
                        model_id=model_id,
                        display_name=display_name_for_model(
                            "Claude",
                            model_id,
                            model.get("display_name") or model.get("displayName") or "",
                        ),
                    )
                )

        if not payload.get("has_more"):
            break
        after_id = payload.get("last_id")
        if not after_id:
            break

    return _dedupe_entries(entries)


def _dedupe_entries(entries: Iterable[ModelCatalogEntry]) -> List[ModelCatalogEntry]:
    deduped = []
    seen = set()
    for entry in entries:
        if entry.model_id in seen:
            continue
        seen.add(entry.model_id)
        deduped.append(entry)
    return deduped


def fetch_provider_models(
    provider: str,
    api_key: Optional[str] = None,
    request_get: Optional[Callable] = None,
) -> List[ModelCatalogEntry]:
    """Fetch live model metadata for a provider.

    The three providers expose list endpoints:
    Gemini returns model capabilities, OpenAI returns accessible model IDs, and
    Anthropic returns Claude model metadata. Callers that need graceful UI
    behavior should use :func:`list_provider_models`.
    """
    request_get = request_get or requests.get

    if provider == "Gemini":
        return _fetch_gemini_models(api_key, request_get)
    if provider == "OpenAI":
        return _fetch_openai_models(api_key, request_get)
    if provider == "Claude":
        return _fetch_claude_models(api_key, request_get)
    raise ModelCatalogError(f"Unsupported provider: {provider}")


def list_provider_models(
    provider: str,
    api_key: Optional[str] = None,
    request_get: Optional[Callable] = None,
    allow_fallback: bool = True,
) -> List[ModelCatalogEntry]:
    """Fetch provider models, falling back to curated local entries on failure."""
    try:
        entries = fetch_provider_models(provider, api_key=api_key, request_get=request_get)
        return entries or fallback_models_for_provider(provider)
    except Exception:
        if allow_fallback:
            return fallback_models_for_provider(provider)
        raise


def model_ids_for_provider(
    provider: str,
    api_key: Optional[str] = None,
    request_get: Optional[Callable] = None,
    allow_fallback: bool = True,
) -> List[str]:
    return [
        entry.model_id
        for entry in list_provider_models(
            provider,
            api_key=api_key,
            request_get=request_get,
            allow_fallback=allow_fallback,
        )
    ]
