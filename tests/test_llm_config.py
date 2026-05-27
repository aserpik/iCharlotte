import unittest
from unittest.mock import patch


class LLMConfigDiscoveryResponseTests(unittest.TestCase):
    def setUp(self):
        from icharlotte_core.llm_config import LLMConfig
        self._saved_instance = LLMConfig._instance
        LLMConfig._instance = None

    def tearDown(self):
        from icharlotte_core.llm_config import LLMConfig
        LLMConfig._instance = self._saved_instance

    def test_default_cap_is_eight_when_key_absent(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        with patch.object(config, "_config", {}):
            self.assertEqual(config.discovery_response_max_concurrent(), 8)

    def test_reads_override_from_config_dict(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        with patch.object(
            config,
            "_config",
            {"discovery_response": {"max_concurrent_proposals": 3}},
        ):
            self.assertEqual(config.discovery_response_max_concurrent(), 3)

    def test_invalid_value_falls_back_to_default(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        with patch.object(
            config,
            "_config",
            {"discovery_response": {"max_concurrent_proposals": "bad"}},
        ):
            self.assertEqual(config.discovery_response_max_concurrent(), 8)

    def test_zero_or_negative_clamps_to_one(self):
        from icharlotte_core.llm_config import LLMConfig
        config = LLMConfig()
        with patch.object(
            config,
            "_config",
            {"discovery_response": {"max_concurrent_proposals": 0}},
        ):
            self.assertEqual(config.discovery_response_max_concurrent(), 1)
        with patch.object(
            config,
            "_config",
            {"discovery_response": {"max_concurrent_proposals": -5}},
        ):
            self.assertEqual(config.discovery_response_max_concurrent(), 1)
