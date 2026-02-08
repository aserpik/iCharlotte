"""Tests for redline configuration."""

import os
import json
import tempfile
import unittest
from icharlotte_core.redline_config import (
    load_redline_settings,
    save_redline_settings,
    DEFAULT_REDLINE_SETTINGS
)

class TestRedlineConfig(unittest.TestCase):

    def setUp(self):
        self.temp_dir = tempfile.mkdtemp()

    def test_load_default_settings(self):
        """Test loading default settings when no file exists."""
        settings = load_redline_settings(self.temp_dir)
        self.assertEqual(settings, DEFAULT_REDLINE_SETTINGS)

    def test_save_and_load_settings(self):
        """Test saving and loading custom settings."""
        custom_settings = {
            "redline_mode_default": True,
            "auto_enable_track_changes": False,
            "redline_fallback_notify": True,
            "max_redline_text_length": 100000
        }

        # Save
        result = save_redline_settings(self.temp_dir, custom_settings)
        self.assertTrue(result)

        # Load
        loaded = load_redline_settings(self.temp_dir)
        self.assertEqual(loaded, custom_settings)

if __name__ == '__main__':
    unittest.main()
