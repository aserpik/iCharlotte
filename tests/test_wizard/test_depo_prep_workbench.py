"""Prompt Workbench integration for Depo Prep.

Verifies the wiring that makes Depo Prep appear in the workbench with editable
prompts + a resolvable model-config id, without invoking the full (heavy)
seed_pipeline_prompts or constructing the QDialog.
"""
import pytest


def test_depo_prep_prompt_defaults_cover_all_passes():
    from Scripts.depo_prep_lib.prompts import (
        DEPO_PREP_PROMPT_DEFAULTS, DEPO_PREP_PROMPT_DESCRIPTIONS,
    )
    expected = {"source_digest", "topic_clustering", "topic_questions", "dedup", "polish"}
    assert set(DEPO_PREP_PROMPT_DEFAULTS) == expected
    # Every pass has a non-trivial default template and a description.
    for pass_name, tmpl in DEPO_PREP_PROMPT_DEFAULTS.items():
        assert isinstance(tmpl, str) and len(tmpl) > 50
        assert DEPO_PREP_PROMPT_DESCRIPTIONS.get(pass_name)


def test_depo_prep_prompts_seed_and_retrieve(tmp_path):
    """The defaults round-trip through PromptManager (the mechanism
    seed_pipeline_prompts uses), so the workbench can list + edit them."""
    from icharlotte_core.prompt_manager import PromptManager
    from Scripts.depo_prep_lib.prompts import (
        DEPO_PREP_PROMPT_DEFAULTS, DEPO_PREP_PROMPT_DESCRIPTIONS,
    )
    pm = PromptManager(prompts_dir=str(tmp_path / "prompts"))
    for pass_name, tmpl in DEPO_PREP_PROMPT_DEFAULTS.items():
        pm.create_version(
            "depo_prep", pass_name, tmpl.strip(), version="v1",
            description=DEPO_PREP_PROMPT_DESCRIPTIONS.get(pass_name, ""),
            set_as_current=True,
        )
    for pass_name in DEPO_PREP_PROMPT_DEFAULTS:
        assert pm.get_prompt("depo_prep", pass_name), f"{pass_name} not retrievable"
    # The optional-field placeholder survives so the builder can fill it.
    assert "{optional_fields_block}" in pm.get_prompt("depo_prep", "topic_questions")


def test_builder_uses_workbench_edited_template(monkeypatch):
    """When the store has an edited template, the builder uses it (proving
    workbench edits take effect at runtime) and substitutes placeholders."""
    import Scripts.depo_prep_lib.prompts as P
    monkeypatch.setattr(
        P, "_load_template",
        lambda pass_name: "EDITED {deponent_name} | {deponent_role} | {source_filename}",
    )
    prompt, payload = P.build_per_source_digest_prompt(
        deponent_name="Jane Doe", deponent_role="Plaintiff",
        source_text="raw text", source_filename="depo.pdf",
    )
    assert prompt == "EDITED Jane Doe | Plaintiff | depo.pdf"
    assert payload == "raw text"


def test_builder_falls_back_to_default_when_store_unavailable(monkeypatch):
    """If the store lookup raises, the builder still returns the default."""
    import Scripts.depo_prep_lib.prompts as P

    def boom(agent, pass_name, version="current"):
        raise RuntimeError("store down")

    # Patch the symbol the lazy import resolves to.
    import icharlotte_core.prompt_manager as pmmod
    monkeypatch.setattr(pmmod, "get_prompt", boom)
    prompt, _ = P.build_polish_prompt(outline_text="x")
    assert "Do not add any new questions." in prompt  # default content


def test_workbench_maps_depo_prep_to_DepoPrep_agent():
    from icharlotte_core.ui.dialogs import WORKBENCH_TO_AGENT_ID
    assert WORKBENCH_TO_AGENT_ID.get("depo_prep") == "DepoPrep"
