"""Configuration for Word redline functionality."""

import os
import json
from typing import Dict, Any

# Default redline settings
DEFAULT_REDLINE_SETTINGS = {
    "redline_mode_default": False,
    "auto_enable_track_changes": True,
    "redline_fallback_notify": True,
    "max_redline_text_length": 50000
}

def load_redline_settings(config_dir: str) -> Dict[str, Any]:
    """Load redline settings from JSON file.

    Args:
        config_dir: Directory containing configuration files

    Returns:
        Dictionary of redline settings
    """
    settings_path = os.path.join(config_dir, "redline_settings.json")

    if os.path.exists(settings_path):
        try:
            with open(settings_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error loading redline settings: {e}")

    return DEFAULT_REDLINE_SETTINGS.copy()

def save_redline_settings(config_dir: str, settings: Dict[str, Any]) -> bool:
    """Save redline settings to JSON file.

    Args:
        config_dir: Directory containing configuration files
        settings: Dictionary of redline settings to save

    Returns:
        True if saved successfully, False otherwise
    """
    settings_path = os.path.join(config_dir, "redline_settings.json")

    try:
        os.makedirs(config_dir, exist_ok=True)
        with open(settings_path, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2)
        return True
    except Exception as e:
        print(f"Error saving redline settings: {e}")
        return False
