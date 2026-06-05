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
        self._registry_path = os.path.join(self.prompts_dir, "registry.json")
        self._ensure_directory_structure()
        self._registry = self._load_registry()

    def _ensure_directory_structure(self):
        """Create the prompts directory structure if needed."""
        if not os.path.exists(self.prompts_dir):
            os.makedirs(self.prompts_dir)

        # Create agent subdirectories
        for agent in ['summarize', 'discovery', 'deposition', 'timeline', 'contradiction',
                      'word_assistant', 'legal_research', 'mediation_brief',
                      'oppose_motion']:
            agent_dir = os.path.join(self.prompts_dir, agent)
            if not os.path.exists(agent_dir):
                os.makedirs(agent_dir)

    def _load_registry(self) -> Dict:
        """Load the prompt registry."""
        if os.path.exists(self._registry_path):
            try:
                with open(self._registry_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"prompts": {}}

    def _save_registry(self):
        """Save the prompt registry."""
        try:
            with open(self._registry_path, "w", encoding="utf-8") as f:
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

    def update_version(
        self,
        agent: str,
        pass_name: str,
        version: str,
        content: str,
    ) -> bool:
        """
        Overwrite an existing version's content in place.

        Unlike create_version(), this does NOT create a new version or change
        any version metadata. If the given version is the active/current one,
        the current pointer file is refreshed so the runtime picks up the edit.

        Args:
            agent: Agent name.
            pass_name: Pass name.
            version: Version to overwrite (must already exist).
            content: New prompt content.

        Returns:
            True if the version existed and was updated, False otherwise.
        """
        key = self._get_prompt_key(agent, pass_name)
        entry = self._registry.get("prompts", {}).get(key)
        if not entry:
            return False

        known = any(v.get("version") == version for v in entry.get("versions", []))
        if not known:
            return False

        # Overwrite the version file in place.
        prompt_path = self._get_prompt_path(agent, pass_name, version)
        os.makedirs(os.path.dirname(prompt_path), exist_ok=True)
        with open(prompt_path, "w", encoding="utf-8") as f:
            f.write(content)

        # If this is the active version, refresh the current pointer file.
        if entry.get("current_version") == version:
            current_path = self._get_current_path(agent, pass_name)
            shutil.copy2(prompt_path, current_path)

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

    def seed_pipeline_prompts(self):
        """Seed all pipeline prompts (word assistant, legal research, mediation brief).

        Writes each hardcoded default as v1 if no version exists yet.
        Idempotent — safe to call multiple times; never overwrites user edits.
        """
        from icharlotte_core.word_hotkey import (
            DEFAULT_WORD_SYSTEM_PROMPT,
            DEFAULT_WORD_REDLINE_SYSTEM_PROMPT,
            EMAIL_SYSTEM_PROMPT,
            DEFAULT_REDLINE_PREFIX,
            DEFAULT_INSERTION_INSTRUCTIONS,
            DEFAULT_SELECTION_INSTRUCTIONS,
        )
        from icharlotte_core.legal_research.prompts import (
            QUERY_PLANNING_PROMPT,
            QUERY_EXTRACTION_PROMPT,
            SYNTHESIS_PROMPT,
            VERIFICATION_PROMPT,
            RELEVANCE_RANKING_PROMPT,
            RESEARCH_FRAMING_INSTRUCTION,
            CITATION_INSTRUCTION,
        )
        from icharlotte_core.mediation_brief import MediationBriefGenerator
        from icharlotte_core.opposition import prompts as oppose_prompts

        # ── Migration: prune deprecated word_assistant entries ──────────────
        # placeholder_instructions and cursor_instructions were consolidated
        # into the templated insertion_instructions entry. Remove the
        # orphaned registry entries + files so the workbench UI stops
        # showing them. Safe to run multiple times (no-op if already pruned).
        deprecated = [
            ("word_assistant", "placeholder_instructions"),
            ("word_assistant", "cursor_instructions"),
        ]
        pruned = 0
        for agent, pass_name in deprecated:
            key = self._get_prompt_key(agent, pass_name)
            if key not in self._registry.get("prompts", {}):
                continue
            # Delete all version files + the current pointer for this prompt
            entry = self._registry["prompts"][key]
            for v_meta in entry.get("versions", []):
                version_id = v_meta.get("version") if isinstance(v_meta, dict) else None
                if not version_id:
                    continue
                vpath = self._get_prompt_path(agent, pass_name, version_id)
                if os.path.exists(vpath):
                    try:
                        os.remove(vpath)
                    except OSError as e:
                        print(f"[PromptManager] Could not remove {vpath}: {e}")
            current_path = self._get_current_path(agent, pass_name)
            if os.path.exists(current_path):
                try:
                    os.remove(current_path)
                except OSError as e:
                    print(f"[PromptManager] Could not remove {current_path}: {e}")
            del self._registry["prompts"][key]
            pruned += 1
        if pruned:
            self._save_registry()
            print(f"[PromptManager] Pruned {pruned} deprecated word_assistant prompts")

        seeds = [
            ("word_assistant", "system_prompt", DEFAULT_WORD_SYSTEM_PROMPT, "Word free-form system prompt (preamble + Word delta)"),
            ("word_assistant", "redline_system_prompt", DEFAULT_WORD_REDLINE_SYSTEM_PROMPT, "Redline system prompt (preamble + redline delta)"),
            ("word_assistant", "email_system_prompt", EMAIL_SYSTEM_PROMPT, "Outlook email system prompt (preamble + email delta)"),
            ("word_assistant", "redline_prefix", DEFAULT_REDLINE_PREFIX, "Prefix appended at end of user message in redline mode"),
            ("word_assistant", "insertion_instructions", DEFAULT_INSERTION_INSTRUCTIONS, "Shared insertion-instructions template (placeholder + cursor). Uses {anchor_label}, {anchor_short}, {extra_rules}."),
            ("word_assistant", "selection_instructions", DEFAULT_SELECTION_INSTRUCTIONS, "Instructions for transforming a selection with full-doc context"),
            ("legal_research", "query_planning", QUERY_PLANNING_PROMPT, "Structured JSON query generation"),
            ("legal_research", "query_extraction", QUERY_EXTRACTION_PROMPT, "Extract queries from litigation prompt"),
            ("legal_research", "synthesis", SYNTHESIS_PROMPT, "Synthesize authorities into memo"),
            ("legal_research", "verification", VERIFICATION_PROMPT, "Citation verification"),
            ("legal_research", "relevance_ranking", RELEVANCE_RANKING_PROMPT, "Case relevance ranking"),
            ("legal_research", "research_framing", RESEARCH_FRAMING_INSTRUCTION, "Citation requirements for user prompt"),
            ("legal_research", "citation_instruction", CITATION_INSTRUCTION, "Strict citation rules"),
            ("mediation_brief", "style_guide", MediationBriefGenerator.STYLE_GUIDE, "Defense writing style/tone guide"),
            ("mediation_brief", "formatting_rules", MediationBriefGenerator.FORMATTING_RULES, "Structural formatting rules"),
            ("oppose_motion", "analyze_motion", oppose_prompts.ANALYZE_MOTION_PROMPT, "Motion analysis: extract metadata + principal arguments"),
            ("oppose_motion", "generate_outline", oppose_prompts.GENERATE_OUTLINE_PROMPT, "Outline generation from analyzed metadata"),
            ("oppose_motion", "research_queries", oppose_prompts.RESEARCH_QUERIES_PROMPT, "Per-argument CourtListener search query generation"),
            ("oppose_motion", "rerank_select", oppose_prompts.RERANK_SELECT_PROMPT, "Re-rank + select best authorities with verbatim passage"),
            ("oppose_motion", "draft_memorandum", oppose_prompts.DRAFT_MEMORANDUM_PROMPT, "Drafter prompt (no pre-draft research; uses style exemplars)"),
            ("oppose_motion", "verify_citation", oppose_prompts.VERIFY_CITATION_PROMPT, "Per-citation verifier: case + statute"),
            ("oppose_motion", "find_replacement", oppose_prompts.FIND_REPLACEMENT_PROMPT, "Optional replacement-case search on red verdicts"),
        ]

        # Depo Prep prompts (Scripts/depo_prep_lib). Guarded so a Scripts import
        # problem can't break seeding of the in-process agents above.
        try:
            from Scripts.depo_prep_lib import prompts as _depo_prompts
            for _pass, _tmpl in _depo_prompts.DEPO_PREP_PROMPT_DEFAULTS.items():
                _desc = _depo_prompts.DEPO_PREP_PROMPT_DESCRIPTIONS.get(_pass, "")
                seeds.append(("depo_prep", _pass, _tmpl, _desc))
        except Exception as e:
            print(f"[PromptManager] Could not seed depo_prep prompts: {e}")

        seeded = 0
        for agent, pass_name, content, description in seeds:
            key = self._get_prompt_key(agent, pass_name)
            current_path = self._get_current_path(agent, pass_name)
            # Skip if registry entry exists AND the file is actually on disk
            if key in self._registry.get("prompts", {}) and os.path.exists(current_path):
                continue
            # Clean orphaned registry entry (key exists but file missing)
            if key in self._registry.get("prompts", {}):
                del self._registry["prompts"][key]
            self.create_version(
                agent, pass_name, content.strip(),
                version="v1",
                description=description,
                author="system",
                set_as_current=True,
            )
            seeded += 1

        if seeded:
            print(f"[PromptManager] Seeded {seeded} pipeline prompts")
        return seeded


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
