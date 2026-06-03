"""Editable registry of motion types for the Generate Motion task.

The registry persists to ``Scripts/prompts/generate_motion/motion_types.json``
and is the runtime source of truth once it exists. When the file is absent the
registry seeds (in memory) from ``config.BUILTIN_SEED`` — reading never writes,
so generating motions has no side effects; the file is created on the first
explicit ``save()`` (e.g. a Workbench edit).
"""
from __future__ import annotations

import json
import os
from typing import Dict, List, Optional

from .config import BUILTIN_SEED, MotionTypeConfig


class MotionTypeRegistry:
    def __init__(self, path: str, types: Optional[List[MotionTypeConfig]] = None) -> None:
        self.path = path
        # Ordered mapping by type_id (dict preserves insertion order).
        self._types: Dict[str, MotionTypeConfig] = {}
        for cfg in types or []:
            self._types[cfg.type_id] = cfg

    # ---- Load / save ---------------------------------------------------- #

    @classmethod
    def load(cls, path: str) -> "MotionTypeRegistry":
        if path and os.path.isfile(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                raw = data.get("types") if isinstance(data, dict) else None
                if isinstance(raw, list):
                    types = [MotionTypeConfig.from_dict(d) for d in raw if isinstance(d, dict)]
                    if types:
                        return cls(path, types)
            except (OSError, ValueError, json.JSONDecodeError):
                pass
        # Seed from the built-ins (in memory; not written until save()).
        return cls(path, list(BUILTIN_SEED.values()))

    def save(self) -> None:
        directory = os.path.dirname(self.path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        payload = {"types": [cfg.to_dict() for cfg in self._types.values()]}
        tmp = self.path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        os.replace(tmp, self.path)

    # ---- Query ---------------------------------------------------------- #

    def list_types(self) -> List[MotionTypeConfig]:
        return list(self._types.values())

    def get(self, type_id: Optional[str]) -> MotionTypeConfig:
        if type_id and type_id in self._types:
            return self._types[type_id]
        if "generic" in self._types:
            return self._types["generic"]
        return BUILTIN_SEED["generic"]

    # ---- Mutate (caller is responsible for save()) ---------------------- #

    def add(self, config: MotionTypeConfig) -> None:
        self._types[config.type_id] = config

    def update(self, type_id: str, **fields) -> bool:
        existing = self._types.get(type_id)
        if existing is None:
            return False
        merged = {**existing.to_dict(), **fields, "type_id": type_id}
        self._types[type_id] = MotionTypeConfig.from_dict(merged)
        return True

    def remove(self, type_id: str) -> bool:
        if type_id in self._types:
            del self._types[type_id]
            return True
        return False

    def restore_defaults(self) -> None:
        self._types = {cfg.type_id: cfg for cfg in BUILTIN_SEED.values()}

    def restore_default(self, type_id: str) -> bool:
        if type_id in BUILTIN_SEED:
            self._types[type_id] = BUILTIN_SEED[type_id]
            return True
        return False
