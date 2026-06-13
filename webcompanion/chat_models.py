"""Curated provider/model list for the chat compose picker.

Derived from the real iCharlotte model sequences so the phone offers the same
models the desktop is configured for.
"""
from icharlotte_core.llm_config import DEFAULT_MODEL_SEQUENCE, FAST_MODEL_SEQUENCE


def available_models() -> list:
    """Return [{'provider','model','label'}], de-duplicated, order preserved."""
    out, seen = [], set()
    for spec in list(DEFAULT_MODEL_SEQUENCE) + list(FAST_MODEL_SEQUENCE):
        key = (spec.provider, spec.model)
        if key in seen:
            continue
        seen.add(key)
        out.append({"provider": spec.provider, "model": spec.model,
                    "label": f"{spec.provider} {spec.model}"})
    return out
