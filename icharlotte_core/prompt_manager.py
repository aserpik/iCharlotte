"""
Prompt Manager for iCharlotte

Manages prompt versioning and retrieval for agent scripts.

Features:
- Versioned prompt storage
- Current/active version tracking
- Prompt history and rollback
- A/B testing support
"""

import os
import shutil
import json
from datetime import datetime
from typing import Optional, List, Dict
from dataclasses import dataclass, asdict


# =============================================================================
# Configuration
# =============================================================================

PROMPTS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "Scripts", "prompts")
PROMPT_REGISTRY = os.path.join(PROMPTS_DIR, "registry.json")


# =============================================================================
# Data Classes
# =============================================================================

@dataclass
class PromptVersion:
    """Metadata for a prompt version."""
    version: str
    created: str
    description: str = ""
    author: str = ""
    is_current: bool = False
    performance_score: float = 0.0
    usage_count: int = 0

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class PromptInfo:
    """Information about a prompt."""
    agent: str
    pass_name: str
    current_version: str = ""
    versions: List[PromptVersion] = None

    def __post_init__(self):
        if self.versions is None:
            self.versions = []


# =============================================================================
# Prompt Manager
# =============================================================================

class PromptManager:
    """
    Manages versioned prompts for agent scripts.

    Usage:
        manager = PromptManager()

        # Get current prompt
        prompt = manager.get_prompt("discovery", "extraction")

        # Get specific version
        prompt = manager.get_prompt("discovery", "extraction", version="v2")

        # Create new version
        manager.create_version("discovery", "extraction", new_content, description="Improved date extraction")

        # Set current version
        manager.set_current("discovery", "extraction", "v2")
    """

    def __init__(self, prompts_dir: str = None):
        """
        Initialize the prompt manager.

        Args:
            prompts_dir: Directory for prompt storage. Defaults to Scripts/prompts.
        """
        self.prompts_dir = prompts_dir or PROMPTS_DIR
        self._ensure_directory_structure()
        self._registry = self._load_registry()

    def _ensure_directory_structure(self):
        """Create the prompts directory structure if needed."""
        if not os.path.exists(self.prompts_dir):
            os.makedirs(self.prompts_dir)

        # Create agent subdirectories
        for agent in ['summarize', 'discovery', 'deposition', 'timeline', 'contradiction']:
            agent_dir = os.path.join(self.prompts_dir, agent)
            if not os.path.exists(agent_dir):
                os.makedirs(agent_dir)

    def _load_registry(self) -> Dict:
        """Load the prompt registry."""
        if os.path.exists(PROMPT_REGISTRY):
            try:
                with open(PROMPT_REGISTRY, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"prompts": {}}

    def _save_registry(self):
        """Save the prompt registry."""
        try:
            with open(PROMPT_REGISTRY, "w", encoding="utf-8") as f:
                json.dump(self._registry, f, indent=2)
        except Exception as e:
            print(f"Warning: Could not save prompt registry: {e}")

    def _get_prompt_key(self, agent: str, pass_name: str) -> str:
        """Get the registry key for a prompt."""
        return f"{agent}:{pass_name}"

    def _get_prompt_path(self, agent: str, pass_name: str, version: str) -> str:
        """Get the file path for a prompt version."""
        return os.path.join(self.prompts_dir, agent, f"{pass_name}_{version}.txt")

    def _get_current_path(self, agent: str, pass_name: str) -> str:
        """Get the path for the current version symlink/copy."""
        return os.path.join(self.prompts_dir, agent, f"{pass_name}_current.txt")

    def get_prompt(
        self,
        agent: str,
        pass_name: str,
        version: str = "current"
    ) -> Optional[str]:
        """
        Get a prompt by agent, pass name, and version.

        Args:
            agent: Agent name (e.g., "discovery", "deposition").
            pass_name: Pass name (e.g., "extraction", "cross_check").
            version: Version string or "current" for active version.

        Returns:
            Prompt text or None if not found.
        """
        if version == "current":
            # Try current file first
            current_path = self._get_current_path(agent, pass_name)
            if os.path.exists(current_path):
                with open(current_path, "r", encoding="utf-8") as f:
                    return f.read()

            # Fall back to registry
            key = self._get_prompt_key(agent, pass_name)
            if key in self._registry.get("prompts", {}):
                version = self._registry["prompts"][key].get("current_version", "v1")

        # Get specific version
        prompt_path = self._get_prompt_path(agent, pass_name, version)
        if os.path.exists(prompt_path):
            with open(prompt_path, "r", encoding="utf-8") as f:
                return f.read()

        # Try legacy paths in Scripts folder
        legacy_paths = [
            os.path.join(os.path.dirname(self.prompts_dir), f"{agent.upper()}_{pass_name.upper()}_PROMPT.txt"),
            os.path.join(os.path.dirname(self.prompts_dir), f"SUMMARIZE_{pass_name.upper()}_PROMPT.txt"),
            os.path.join(os.path.dirname(self.prompts_dir), f"{pass_name.upper()}_PROMPT.txt"),
        ]

        for legacy_path in legacy_paths:
            if os.path.exists(legacy_path):
                with open(legacy_path, "r", encoding="utf-8") as f:
                    return f.read()

        return None

    def list_versions(self, agent: str, pass_name: str) -> List[PromptVersion]:
        """
        List all versions of a prompt.

        Args:
            agent: Agent name.
            pass_name: Pass name.

        Returns:
            List of PromptVersion objects.
        """
        key = self._get_prompt_key(agent, pass_name)
        prompt_info = self._registry.get("prompts", {}).get(key, {})

        versions = []
        for v_data in prompt_info.get("versions", []):
            versions.append(PromptVersion(**v_data))

        return versions

    def create_version(
        self,
        agent: str,
        pass_name: str,
        content: str,
        version: str = None,
        description: str = "",
        author: str = "",
        set_as_current: bool = False
    ) -> str:
        """
        Create a new version of a prompt.

        Args:
            agent: Agent name.
            pass_name: Pass name.
            content: Prompt content.
            version: Version string (auto-generated if None).
            description: Description of changes.
            author: Author of the version.
            set_as_current: Whether to set this as the current version.

        Returns:
            The version string.
        """
        key = self._get_prompt_key(agent, pass_name)

        # Initialize prompt info if needed
        if key not in self._registry.get("prompts", {}):
            if "prompts" not in self._registry:
                self._registry["prompts"] = {}
            self._registry["prompts"][key] = {
                "agent": agent,
                "pass_name": pass_name,
                "current_version": "",
                "versions": []
            }

        prompt_info = self._registry["prompts"][key]

        # Generate version number if not provided
        if version is None:
            existing = [v["version"] for v in prompt_info.get("versions", [])]
            version_num = 1
            while f"v{version_num}" in existing:
                version_num += 1
            version = f"v{version_num}"

        # Create version metadata
        version_meta = PromptVersion(
            version=version,
            created=datetime.now().isoformat(),
            description=description,
            author=author,
            is_current=set_as_current
        )

        # Save prompt file
        prompt_path = self._get_prompt_path(agent, pass_name, version)
        os.makedirs(os.path.dirname(prompt_path), exist_ok=True)
        with open(prompt_path, "w", encoding="utf-8") as f:
            f.write(content)

        # Update registry
        prompt_info["versions"].append(version_meta.to_dict())

        if set_as_current:
            self.set_current(agent, pass_name, version)
        else:
            self._save_registry()

        return version

    def set_current(self, agent: str, pass_name: str, version: str) -> bool:
        """
        Set a version as the current/active version.

        Args:
            agent: Agent name.
            pass_name: Pass name.
            version: Version to set as current.

        Returns:
            True if successful.
        """
        key = self._get_prompt_key(agent, pass_name)

        # Check version exists
        prompt_path = self._get_prompt_path(agent, pass_name, version)
        if not os.path.exists(prompt_path):
            return False

        # Update registry
        if key in self._registry.get("prompts", {}):
            self._registry["prompts"][key]["current_version"] = version

            # Update is_current flags
            for v in self._registry["prompts"][key].get("versions", []):
                v["is_current"] = (v["version"] == version)

        # Copy to current file
        current_path = self._get_current_path(agent, pass_name)
        shutil.copy2(prompt_path, current_path)

        self._save_registry()
        return True

    def get_info(self, agent: str, pass_name: str) -> Optional[PromptInfo]:
        """
        Get information about a prompt.

        Args:
            agent: Agent name.
            pass_name: Pass name.

        Returns:
            PromptInfo object or None.
        """
        key = self._get_prompt_key(agent, pass_name)
        data = self._registry.get("prompts", {}).get(key)

        if not data:
            return None

        return PromptInfo(
            agent=data.get("agent", agent),
            pass_name=data.get("pass_name", pass_name),
            current_version=data.get("current_version", ""),
            versions=[PromptVersion(**v) for v in data.get("versions", [])]
        )

    def record_usage(self, agent: str, pass_name: str, version: str = "current"):
        """Record that a prompt was used."""
        key = self._get_prompt_key(agent, pass_name)

        if key in self._registry.get("prompts", {}):
            if version == "current":
                version = self._registry["prompts"][key].get("current_version", "")

            for v in self._registry["prompts"][key].get("versions", []):
                if v["version"] == version:
                    v["usage_count"] = v.get("usage_count", 0) + 1
                    break

            self._save_registry()

    def record_performance(
        self,
        agent: str,
        pass_name: str,
        score: float,
        version: str = "current"
    ):
        """
        Record a performance score for a prompt version.

        Args:
            agent: Agent name.
            pass_name: Pass name.
            score: Performance score (0.0 to 1.0).
            version: Version string or "current".
        """
        key = self._get_prompt_key(agent, pass_name)

        if key in self._registry.get("prompts", {}):
            if version == "current":
                version = self._registry["prompts"][key].get("current_version", "")

            for v in self._registry["prompts"][key].get("versions", []):
                if v["version"] == version:
                    # Rolling average
                    old_score = v.get("performance_score", 0.0)
                    usage = v.get("usage_count", 1)
                    v["performance_score"] = (old_score * (usage - 1) + score) / usage
                    break

            self._save_registry()

    def migrate_legacy_prompts(self):
        """
        Migrate existing prompts from Scripts folder to versioned storage.

        This should be run once to import existing prompts.
        """
        scripts_dir = os.path.dirname(self.prompts_dir)

        # Map of legacy files to agent/pass
        legacy_map = {
            "SUMMARIZE_PROMPT.txt": ("summarize", "summary"),
            "SUMMARIZE_CROSS_CHECK_PROMPT.txt": ("summarize", "cross_check"),
            "CONSOLIDATE_DISCOVERY_PROMPT.txt": ("discovery", "extraction"),
            "CROSS_CHECK_PROMPT.txt": ("discovery", "cross_check"),
            "SUMMARIZE_DEPOSITION_PROMPT.txt": ("deposition", "summary"),
            "DEPOSITION_EXTRACTION_PROMPT.txt": ("deposition", "extraction"),
            "DEPOSITION_CROSS_CHECK_PROMPT.txt": ("deposition", "cross_check"),
            "TIMELINE_EXTRACTION_PROMPT.txt": ("timeline", "extraction"),
            "CONTRADICTION_DETECTION_PROMPT.txt": ("contradiction", "detection"),
        }

        migrated = 0
        for filename, (agent, pass_name) in legacy_map.items():
            legacy_path = os.path.join(scripts_dir, filename)
            if os.path.exists(legacy_path):
                with open(legacy_path, "r", encoding="utf-8") as f:
                    content = f.read()

                self.create_version(
                    agent, pass_name, content,
                    version="v1",
                    description="Migrated from legacy location",
                    set_as_current=True
                )
                migrated += 1

        return migrated


# =============================================================================
# Global Instance
# =============================================================================

_default_manager: Optional[PromptManager] = None


def get_prompt_manager() -> PromptManager:
    """Get the global prompt manager instance."""
    global _default_manager
    if _default_manager is None:
        _default_manager = PromptManager()
    return _default_manager


def get_prompt(agent: str, pass_name: str, version: str = "current") -> Optional[str]:
    """Convenience function to get a prompt."""
    return get_prompt_manager().get_prompt(agent, pass_name, version)
