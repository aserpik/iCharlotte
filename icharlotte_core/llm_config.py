"""
Config-Driven LLM Preferences for iCharlotte

Provides centralized configuration for LLM model selection with
multi-provider fallback support (Gemini, Claude, OpenAI).

Supports per-agent/function model configuration with fallback to defaults.
"""

import os
import json
from dataclasses import dataclass, field, asdict
from typing import List, Optional, Dict, Any
from enum import Enum


# =============================================================================
# Environment Variable Normalization
# =============================================================================

# The google-genai SDK checks both GOOGLE_API_KEY and GEMINI_API_KEY env vars.
# When both are set, it logs a warning on every client creation.
# Normalize to only GEMINI_API_KEY to prevent the warning.
def _normalize_gemini_api_key():
    """Normalize Gemini API key environment variables to prevent SDK warnings."""
    gemini_key = os.environ.get("GEMINI_API_KEY")
    google_key = os.environ.get("GOOGLE_API_KEY")

    if gemini_key and google_key:
        # Both are set - remove GOOGLE_API_KEY to prevent warning
        del os.environ["GOOGLE_API_KEY"]
    elif google_key and not gemini_key:
        # Only GOOGLE_API_KEY is set - copy to GEMINI_API_KEY and remove original
        os.environ["GEMINI_API_KEY"] = google_key
        del os.environ["GOOGLE_API_KEY"]

# Run normalization at module load time
_normalize_gemini_api_key()


# =============================================================================
# Configuration File Path
# =============================================================================

CONFIG_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'config')
CONFIG_FILE = os.path.join(CONFIG_DIR, 'llm_preferences.json')


# =============================================================================
# Agent/Function Definitions
# =============================================================================

# All agents and functions that use LLMs
# Format: (id, display_name, description, default_task_type)
AGENT_DEFINITIONS = [
    # Document Processing Agents
    ("agent_separate", "Separate", "PDF separation and document indexing", "extraction"),
    ("agent_summarize", "Summarize", "Document summarization with verification", "summary"),
    ("agent_sum_disc", "Summarize Discovery", "Discovery response extraction and summarization", "extraction"),
    ("agent_sum_depo", "Summarize Deposition", "Deposition transcript analysis", "extraction"),
    ("agent_med_rec", "Medical Records", "Medical record processing", "extraction"),
    ("agent_med_chron", "Medical Chronology", "Medical chronology generation", "extraction"),
    ("agent_organize", "Organize", "Document organization and categorization", "quick"),
    ("agent_timeline", "Timeline", "Date and event extraction for timelines", "extraction"),
    ("agent_contradict", "Contradictions", "Detect factual contradictions across documents", "cross_check"),

    # Case Agents
    ("agent_docket", "Docket", "Docket download and processing", "extraction"),
    ("agent_complaint", "Complaint", "Complaint document analysis", "extraction"),
    ("agent_subpoena", "Subpoena Tracker", "Subpoena tracking and analysis", "extraction"),
    ("agent_liability", "Liability Script", "Liability analysis script", "extraction"),
    ("agent_exposure", "Exposure", "Exposure calculation and analysis", "extraction"),

    # UI Functions
    ("func_chat", "Chat Tab", "Interactive AI chat", "general"),
    ("func_email_intelligence", "Email Intelligence", "AI-powered email analysis", "general"),
    ("func_email_compose", "Email Compose", "AI-powered email composition", "general"),
    ("func_liability_tab", "Liability Tab", "Liability and exposure analysis tab", "general"),
    ("func_sent_monitor", "Sent Items Monitor", "Monitor sent emails for todos", "quick"),
    ("func_attachment_classifier", "Attachment Classifier", "Classify legal document attachments", "classification"),
]

# Create lookup dict
AGENT_LOOKUP = {agent[0]: {"name": agent[1], "description": agent[2], "default_task": agent[3]}
                for agent in AGENT_DEFINITIONS}


# =============================================================================
# Data Classes
# =============================================================================

class Provider(Enum):
    """Supported LLM providers."""
    GEMINI = "Gemini"
    CLAUDE = "Claude"
    OPENAI = "OpenAI"


@dataclass
class ModelSpec:
    """Specification for a single LLM model."""
    provider: str
    model: str
    max_tokens: int = 8192
    supports_thinking: bool = False
    supports_streaming: bool = True
    cost_tier: str = "standard"  # "low", "standard", "high"

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> 'ModelSpec':
        # Handle legacy data that might have extra fields
        valid_fields = {'provider', 'model', 'max_tokens', 'supports_thinking',
                        'supports_streaming', 'cost_tier'}
        filtered = {k: v for k, v in data.items() if k in valid_fields}
        return cls(**filtered)


@dataclass
class AgentConfig:
    """Configuration for a specific agent or function."""
    agent_id: str
    display_name: str
    description: str = ""
    model_sequence: List[ModelSpec] = field(default_factory=list)
    max_retries: int = 3
    timeout_seconds: int = 120
    use_default: bool = True  # If True, use default task-type config instead

    def to_dict(self) -> dict:
        return {
            "agent_id": self.agent_id,
            "display_name": self.display_name,
            "description": self.description,
            "model_sequence": [m.to_dict() for m in self.model_sequence],
            "max_retries": self.max_retries,
            "timeout_seconds": self.timeout_seconds,
            "use_default": self.use_default
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'AgentConfig':
        return cls(
            agent_id=data["agent_id"],
            display_name=data.get("display_name", data["agent_id"]),
            description=data.get("description", ""),
            model_sequence=[ModelSpec.from_dict(m) for m in data.get("model_sequence", [])],
            max_retries=data.get("max_retries", 3),
            timeout_seconds=data.get("timeout_seconds", 120),
            use_default=data.get("use_default", True)
        )


@dataclass
class TaskConfig:
    """Configuration for a specific task type (used as defaults)."""
    name: str
    model_sequence: List[ModelSpec] = field(default_factory=list)
    max_retries: int = 3
    timeout_seconds: int = 120

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "model_sequence": [m.to_dict() for m in self.model_sequence],
            "max_retries": self.max_retries,
            "timeout_seconds": self.timeout_seconds
        }

    @classmethod
    def from_dict(cls, data: dict) -> 'TaskConfig':
        return cls(
            name=data["name"],
            model_sequence=[ModelSpec.from_dict(m) for m in data.get("model_sequence", [])],
            max_retries=data.get("max_retries", 3),
            timeout_seconds=data.get("timeout_seconds", 120)
        )


# =============================================================================
# Default Model Configurations
# =============================================================================

# Default model sequence for general tasks
DEFAULT_MODEL_SEQUENCE = [
    ModelSpec(
        provider="Gemini",
        model="gemini-2.5-pro",
        max_tokens=8192,
        supports_thinking=True,
        cost_tier="high"
    ),
    ModelSpec(
        provider="Gemini",
        model="gemini-2.5-flash",
        max_tokens=8192,
        supports_thinking=True,
        cost_tier="standard"
    ),
    ModelSpec(
        provider="Claude",
        model="claude-sonnet-4-20250514",
        max_tokens=8192,
        supports_thinking=False,
        cost_tier="high"
    ),
    ModelSpec(
        provider="OpenAI",
        model="gpt-4o",
        max_tokens=8192,
        supports_thinking=False,
        cost_tier="high"
    ),
]

# Fast model sequence for quick operations
FAST_MODEL_SEQUENCE = [
    ModelSpec(
        provider="Gemini",
        model="gemini-2.5-flash",
        max_tokens=4096,
        supports_thinking=False,
        cost_tier="low"
    ),
    ModelSpec(
        provider="Claude",
        model="claude-haiku-4-20250514",
        max_tokens=4096,
        supports_thinking=False,
        cost_tier="low"
    ),
    ModelSpec(
        provider="OpenAI",
        model="gpt-4o-mini",
        max_tokens=4096,
        supports_thinking=False,
        cost_tier="low"
    ),
]

# Gemini 3 model sequence for cross-check/contradiction tasks
GEMINI3_MODEL_SEQUENCE = [
    ModelSpec(
        provider="Gemini",
        model="gemini-3-pro-preview",
        max_tokens=8192,
        supports_thinking=True,
        cost_tier="high"
    ),
    ModelSpec(
        provider="Gemini",
        model="gemini-3-flash-preview",
        max_tokens=8192,
        supports_thinking=True,
        cost_tier="standard"
    ),
    ModelSpec(
        provider="Claude",
        model="claude-sonnet-4-20250514",
        max_tokens=8192,
        supports_thinking=False,
        cost_tier="high"
    ),
]

# Default task configurations
DEFAULT_TASK_CONFIGS = {
    "general": TaskConfig(
        name="general",
        model_sequence=DEFAULT_MODEL_SEQUENCE,
        max_retries=3,
        timeout_seconds=120
    ),
    "extraction": TaskConfig(
        name="extraction",
        model_sequence=DEFAULT_MODEL_SEQUENCE,
        max_retries=3,
        timeout_seconds=180
    ),
    "summary": TaskConfig(
        name="summary",
        model_sequence=DEFAULT_MODEL_SEQUENCE,
        max_retries=3,
        timeout_seconds=180
    ),
    "cross_check": TaskConfig(
        name="cross_check",
        model_sequence=GEMINI3_MODEL_SEQUENCE,
        max_retries=2,
        timeout_seconds=120
    ),
    "classification": TaskConfig(
        name="classification",
        model_sequence=FAST_MODEL_SEQUENCE,
        max_retries=2,
        timeout_seconds=60
    ),
    "quick": TaskConfig(
        name="quick",
        model_sequence=FAST_MODEL_SEQUENCE,
        max_retries=2,
        timeout_seconds=60
    ),
}


# =============================================================================
# LLM Configuration Manager
# =============================================================================

class LLMConfig:
    """
    Centralized LLM configuration manager.

    Loads configuration from JSON file with defaults fallback.
    Supports agent-specific and task-specific model sequences with multi-provider fallback.

    Usage:
        config = LLMConfig()

        # Get config for a specific agent
        models = config.get_model_sequence_for_agent("agent_summarize")

        # Get config for a task type (fallback)
        models = config.get_model_sequence("summary")
    """

    _instance: Optional['LLMConfig'] = None

    def __new__(cls):
        """Singleton pattern for global configuration."""
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return

        self._initialized = True
        self._config: Dict[str, Any] = {}
        self._task_configs: Dict[str, TaskConfig] = {}
        self._agent_configs: Dict[str, AgentConfig] = {}

        # Load configuration
        self._load_config()

    def _load_config(self):
        """Load configuration from file or use defaults."""
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                self._config = data

                # Parse task configurations
                for task_name, task_data in data.get("tasks", {}).items():
                    self._task_configs[task_name] = TaskConfig.from_dict(task_data)

                # Parse agent configurations
                for agent_id, agent_data in data.get("agents", {}).items():
                    self._agent_configs[agent_id] = AgentConfig.from_dict(agent_data)

            except (json.JSONDecodeError, KeyError) as e:
                print(f"Warning: Failed to load LLM config: {e}. Using defaults.")
                self._use_defaults()
        else:
            self._use_defaults()
            # Create default config file
            self._save_config()

    def _use_defaults(self):
        """Use default configuration."""
        self._task_configs = {k: TaskConfig(
            name=v.name,
            model_sequence=list(v.model_sequence),
            max_retries=v.max_retries,
            timeout_seconds=v.timeout_seconds
        ) for k, v in DEFAULT_TASK_CONFIGS.items()}

        # Initialize agent configs with use_default=True
        self._agent_configs = {}
        for agent_id, info in AGENT_LOOKUP.items():
            self._agent_configs[agent_id] = AgentConfig(
                agent_id=agent_id,
                display_name=info["name"],
                description=info["description"],
                model_sequence=[],
                use_default=True
            )

        self._config = {
            "version": "2.0",
            "default_provider": "Gemini",
            "tasks": {name: config.to_dict() for name, config in self._task_configs.items()},
            "agents": {agent_id: config.to_dict() for agent_id, config in self._agent_configs.items()}
        }

    def _save_config(self):
        """Save current configuration to file."""
        try:
            os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)

            # Update config dict before saving
            self._config["tasks"] = {
                name: config.to_dict() for name, config in self._task_configs.items()
            }
            self._config["agents"] = {
                agent_id: config.to_dict() for agent_id, config in self._agent_configs.items()
            }

            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, indent=2)
        except IOError as e:
            print(f"Warning: Failed to save LLM config: {e}")

    def reload(self):
        """Reload configuration from file."""
        self._initialized = False
        self.__init__()

    # =========================================================================
    # Agent-Specific API
    # =========================================================================

    def get_all_agents(self) -> List[str]:
        """Get list of all agent IDs."""
        return list(AGENT_LOOKUP.keys())

    def get_agent_info(self, agent_id: str) -> Dict[str, str]:
        """Get display info for an agent."""
        return AGENT_LOOKUP.get(agent_id, {"name": agent_id, "description": "", "default_task": "general"})

    def get_agent_config(self, agent_id: str) -> AgentConfig:
        """Get configuration for a specific agent."""
        if agent_id in self._agent_configs:
            return self._agent_configs[agent_id]

        # Create default config
        info = self.get_agent_info(agent_id)
        return AgentConfig(
            agent_id=agent_id,
            display_name=info["name"],
            description=info["description"],
            use_default=True
        )

    def get_model_sequence_for_agent(self, agent_id: str, fallback_task: str = None) -> List[ModelSpec]:
        """
        Get the model sequence for a specific agent.

        Args:
            agent_id: Agent identifier (e.g., "agent_summarize", "func_chat")
            fallback_task: Optional task type to use if agent uses default

        Returns:
            List of ModelSpec objects in order of preference.
        """
        agent_config = self._agent_configs.get(agent_id)

        if agent_config and not agent_config.use_default and agent_config.model_sequence:
            return agent_config.model_sequence

        # Fall back to task type
        if fallback_task:
            task_config = self._task_configs.get(fallback_task)
            if task_config:
                return task_config.model_sequence

        # Use agent's default task type
        info = self.get_agent_info(agent_id)
        default_task = info.get("default_task", "general")
        task_config = self._task_configs.get(default_task)
        if task_config:
            return task_config.model_sequence

        return DEFAULT_MODEL_SEQUENCE

    def update_agent_config(self, agent_id: str, model_sequence: List[ModelSpec] = None,
                            use_default: bool = None, max_retries: int = None,
                            timeout_seconds: int = None):
        """
        Update configuration for a specific agent.

        Args:
            agent_id: Agent identifier.
            model_sequence: New model sequence (if use_default is False).
            use_default: Whether to use default task-type config.
            max_retries: Optional new max retries.
            timeout_seconds: Optional new timeout.
        """
        if agent_id not in self._agent_configs:
            info = self.get_agent_info(agent_id)
            self._agent_configs[agent_id] = AgentConfig(
                agent_id=agent_id,
                display_name=info["name"],
                description=info["description"],
                use_default=True
            )

        config = self._agent_configs[agent_id]

        if use_default is not None:
            config.use_default = use_default
        if model_sequence is not None:
            config.model_sequence = model_sequence
        if max_retries is not None:
            config.max_retries = max_retries
        if timeout_seconds is not None:
            config.timeout_seconds = timeout_seconds

        self._save_config()

    # =========================================================================
    # Task-Type API (for backwards compatibility and defaults)
    # =========================================================================

    def get_all_task_types(self) -> List[str]:
        """Get list of all configured task types."""
        return list(self._task_configs.keys())

    def get_model_sequence(self, task_type: str = "general") -> List[ModelSpec]:
        """
        Get the model sequence for a specific task type.

        Args:
            task_type: Type of task (general, extraction, summary, etc.)

        Returns:
            List of ModelSpec objects in order of preference.
        """
        task_config = self._task_configs.get(task_type)
        if task_config:
            return task_config.model_sequence

        # Fallback to general
        general_config = self._task_configs.get("general")
        if general_config:
            return general_config.model_sequence

        return DEFAULT_MODEL_SEQUENCE

    def get_task_config(self, task_type: str = "general") -> TaskConfig:
        """
        Get full task configuration.

        Args:
            task_type: Type of task.

        Returns:
            TaskConfig object.
        """
        return self._task_configs.get(task_type,
                                      self._task_configs.get("general",
                                                             DEFAULT_TASK_CONFIGS["general"]))

    def update_task_config(self, task_type: str, model_sequence: List[ModelSpec],
                           max_retries: int = None, timeout_seconds: int = None):
        """
        Update configuration for a specific task type.

        Args:
            task_type: Type of task to update.
            model_sequence: New model sequence.
            max_retries: Optional new max retries.
            timeout_seconds: Optional new timeout.
        """
        if task_type not in self._task_configs:
            self._task_configs[task_type] = TaskConfig(
                name=task_type,
                model_sequence=model_sequence,
                max_retries=max_retries or 3,
                timeout_seconds=timeout_seconds or 120
            )
        else:
            existing = self._task_configs[task_type]
            existing.model_sequence = model_sequence
            if max_retries is not None:
                existing.max_retries = max_retries
            if timeout_seconds is not None:
                existing.timeout_seconds = timeout_seconds

        self._save_config()

    # =========================================================================
    # Provider API
    # =========================================================================

    def get_fallback_providers(self) -> List[str]:
        """Get list of all configured providers in fallback order."""
        providers = []
        seen = set()

        for model in self.get_model_sequence("general"):
            if model.provider not in seen:
                providers.append(model.provider)
                seen.add(model.provider)

        return providers

    def get_default_provider(self) -> str:
        """Get the default/primary provider."""
        return self._config.get("default_provider", "Gemini")

    def set_default_provider(self, provider: str):
        """Set the default provider."""
        self._config["default_provider"] = provider
        self._save_config()

    def get_api_key(self, provider: str) -> Optional[str]:
        """Get API key for a provider from environment."""
        env_vars = {
            "Gemini": "GEMINI_API_KEY",
            "Claude": "ANTHROPIC_API_KEY",
            "OpenAI": "OPENAI_API_KEY",
        }

        env_var = env_vars.get(provider)
        if env_var:
            return os.environ.get(env_var)
        return None

    def is_provider_available(self, provider: str) -> bool:
        """Check if a provider has an API key configured."""
        return self.get_api_key(provider) is not None

    def get_available_providers(self) -> List[str]:
        """Get list of providers with configured API keys."""
        return [p for p in self.get_fallback_providers() if self.is_provider_available(p)]


# =============================================================================
# LLM Caller with Fallback
# =============================================================================

class LLMCaller:
    """
    Unified LLM caller with automatic fallback support.

    Uses LLMConfig to try models in sequence until one succeeds.

    Usage:
        caller = LLMCaller()
        result = caller.call("Summarize this document", text, agent_id="agent_summarize")
    """

    def __init__(self, config: LLMConfig = None, logger=None):
        """
        Initialize the LLM caller.

        Args:
            config: LLMConfig instance. Uses singleton if not provided.
            logger: Optional logger for output.
        """
        self.config = config or LLMConfig()
        self.logger = logger

        # Provider-specific clients (lazy initialized)
        self._clients: Dict[str, Any] = {}

    def _log(self, message: str, level: str = "info"):
        """Log a message."""
        if self.logger:
            if hasattr(self.logger, level):
                getattr(self.logger, level)(message)
            else:
                self.logger.info(message)
        else:
            print(message, flush=True)

    def _get_gemini_client(self):
        """Get or create Gemini client."""
        if "Gemini" not in self._clients:
            try:
                from google import genai
                api_key = self.config.get_api_key("Gemini")
                if api_key:
                    self._clients["Gemini"] = genai.Client(api_key=api_key)
            except ImportError:
                pass
        return self._clients.get("Gemini")

    def _detect_provider(self, model: str) -> Optional[str]:
        """Detect the provider from a model name."""
        model_lower = model.lower()
        if "gemini" in model_lower:
            return "Gemini"
        elif "claude" in model_lower or "haiku" in model_lower or "sonnet" in model_lower or "opus" in model_lower:
            return "Claude"
        elif "gpt" in model_lower or "o1" in model_lower or "o3" in model_lower:
            return "OpenAI"
        return None

    def _call_gemini(self, model: str, prompt: str, text: str) -> Optional[str]:
        """Call Gemini API."""
        client = self._get_gemini_client()
        if not client:
            raise Exception("Gemini client not available")

        full_prompt = f"{prompt}\n\nDOCUMENT CONTENT:\n{text}"

        response = client.models.generate_content(
            model=model,
            contents=full_prompt
        )

        if response and response.text:
            return response.text
        return None

    def _call_claude(self, model: str, prompt: str, text: str) -> Optional[str]:
        """Call Claude API."""
        import requests

        api_key = self.config.get_api_key("Claude")
        if not api_key:
            raise Exception("Claude API key not available")

        full_content = f"{prompt}\n\nDOCUMENT CONTENT:\n{text}"

        response = requests.post(
            "https://api.anthropic.com/v1/messages",
            headers={
                "x-api-key": api_key,
                "anthropic-version": "2023-06-01",
                "content-type": "application/json"
            },
            json={
                "model": model,
                "max_tokens": 8192,
                "messages": [{"role": "user", "content": full_content}]
            },
            timeout=120
        )

        if response.status_code == 200:
            data = response.json()
            if data.get("content"):
                return data["content"][0].get("text", "")

        response.raise_for_status()
        return None

    def _call_openai(self, model: str, prompt: str, text: str) -> Optional[str]:
        """Call OpenAI API."""
        import requests

        api_key = self.config.get_api_key("OpenAI")
        if not api_key:
            raise Exception("OpenAI API key not available")

        full_content = f"{prompt}\n\nDOCUMENT CONTENT:\n{text}"

        response = requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json"
            },
            json={
                "model": model,
                "messages": [{"role": "user", "content": full_content}],
                "max_tokens": 8192
            },
            timeout=120
        )

        if response.status_code == 200:
            data = response.json()
            if data.get("choices"):
                return data["choices"][0].get("message", {}).get("content", "")

        response.raise_for_status()
        return None

    def call(self, prompt: str, text: str, task_type: str = "general",
             agent_id: str = None, model_override: str = None) -> Optional[str]:
        """
        Call LLM with automatic fallback through configured models.

        Args:
            prompt: The instruction/prompt.
            text: The document content.
            task_type: Type of task for model selection (fallback).
            agent_id: Specific agent ID for model selection (priority).
            model_override: Specific model to use (e.g., "gemini-2.5-pro").
                           Bypasses model sequence and uses only this model.

        Returns:
            LLM response text or None if all models fail.
        """
        # Handle model override - create a single-model sequence
        if model_override:
            provider = self._detect_provider(model_override)
            if provider:
                model_sequence = [ModelSpec(provider=provider, model=model_override)]
                max_retries = 3
                self._log(f"Using model override: {model_override} ({provider})")
            else:
                self._log(f"Unknown provider for model: {model_override}, falling back to default", "warning")
                model_override = None

        # Get model sequence - prefer agent-specific, fallback to task type
        if not model_override:
            if agent_id:
                model_sequence = self.config.get_model_sequence_for_agent(agent_id, task_type)
                agent_config = self.config.get_agent_config(agent_id)
                max_retries = agent_config.max_retries if not agent_config.use_default else 3
            else:
                model_sequence = self.config.get_model_sequence(task_type)
                task_config = self.config.get_task_config(task_type)
                max_retries = task_config.max_retries

        last_error = None

        for model_spec in model_sequence:
            # Skip providers without API keys
            if not self.config.is_provider_available(model_spec.provider):
                continue

            self._log(f"Attempting {model_spec.provider} model: {model_spec.model}")

            for attempt in range(max_retries):
                try:
                    if model_spec.provider == "Gemini":
                        result = self._call_gemini(model_spec.model, prompt, text)
                    elif model_spec.provider == "Claude":
                        result = self._call_claude(model_spec.model, prompt, text)
                    elif model_spec.provider == "OpenAI":
                        result = self._call_openai(model_spec.model, prompt, text)
                    else:
                        continue

                    if result:
                        self._log(f"Success with {model_spec.provider}/{model_spec.model}")
                        return result

                except Exception as e:
                    last_error = e
                    self._log(f"Attempt {attempt + 1} failed with {model_spec.model}: {e}", "warning")

                    if attempt < max_retries - 1:
                        import time
                        time.sleep(2 ** attempt)  # Exponential backoff

        self._log(f"All models failed. Last error: {last_error}", "error")
        return None


# =============================================================================
# Convenience Functions
# =============================================================================

def get_model_sequence(task_type: str = "general") -> List[ModelSpec]:
    """Get model sequence for a task type (convenience function)."""
    return LLMConfig().get_model_sequence(task_type)


def get_model_sequence_for_agent(agent_id: str) -> List[ModelSpec]:
    """Get model sequence for an agent (convenience function)."""
    return LLMConfig().get_model_sequence_for_agent(agent_id)


def get_api_key(provider: str) -> Optional[str]:
    """Get API key for a provider (convenience function)."""
    return LLMConfig().get_api_key(provider)


def call_llm(prompt: str, text: str, task_type: str = "general",
             agent_id: str = None, logger=None) -> Optional[str]:
    """
    Call LLM with automatic fallback (convenience function).

    Args:
        prompt: The instruction/prompt.
        text: The document content.
        task_type: Type of task (fallback).
        agent_id: Specific agent ID (priority).
        logger: Optional logger.

    Returns:
        LLM response or None.
    """
    caller = LLMCaller(logger=logger)
    return caller.call(prompt, text, task_type, agent_id)
