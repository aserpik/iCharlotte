import unittest

from icharlotte_core import llm
from icharlotte_core import model_catalog


class FakeResponse:
    def __init__(self, payload, status_code=200, text="OK"):
        self._payload = payload
        self.status_code = status_code
        self.text = text

    def json(self):
        return self._payload


class ModelCatalogTests(unittest.TestCase):
    def test_gemini_fetch_filters_to_generate_content_models(self):
        calls = []

        def fake_get(url, headers=None, params=None, timeout=None):
            calls.append((url, headers, params, timeout))
            return FakeResponse(
                {
                    "models": [
                        {
                            "name": "models/gemini-3.1-pro-preview",
                            "displayName": "Gemini 3.1 Pro Preview",
                            "supportedGenerationMethods": ["generateContent"],
                        },
                        {
                            "name": "models/gemini-embedding-only",
                            "displayName": "Gemini Embedding",
                            "supportedGenerationMethods": ["embedContent"],
                        },
                    ]
                }
            )

        entries = model_catalog.fetch_provider_models(
            "Gemini", "gemini-key", request_get=fake_get
        )

        self.assertEqual(calls[0][0], model_catalog.GEMINI_MODELS_ENDPOINT)
        self.assertEqual(calls[0][2]["key"], "gemini-key")
        self.assertEqual([entry.model_id for entry in entries], ["gemini-3.1-pro-preview"])
        self.assertEqual(entries[0].display_name, "Gemini 3.1 Pro Preview")

    def test_openai_fetch_filters_to_text_generation_models(self):
        def fake_get(url, headers=None, params=None, timeout=None):
            return FakeResponse(
                {
                    "data": [
                        {"id": "text-embedding-3-large", "created": 10},
                        {"id": "gpt-5.5-pro", "created": 30},
                        {"id": "o4-mini", "created": 20},
                        {"id": "gpt-realtime", "created": 40},
                        {"id": "chatgpt-4o-latest", "created": 15},
                    ]
                }
            )

        entries = model_catalog.fetch_provider_models(
            "OpenAI", "openai-key", request_get=fake_get
        )

        self.assertEqual(
            [entry.model_id for entry in entries],
            ["gpt-5.5-pro", "o4-mini", "chatgpt-4o-latest"],
        )

    def test_claude_fetch_uses_anthropic_models_api_and_paginates(self):
        calls = []
        pages = [
            {
                "data": [
                    {
                        "id": "claude-opus-4-6",
                        "display_name": "Claude Opus 4.6",
                        "type": "model",
                    }
                ],
                "has_more": True,
                "last_id": "claude-opus-4-6",
            },
            {
                "data": [
                    {
                        "id": "claude-sonnet-4-6",
                        "display_name": "Claude Sonnet 4.6",
                        "type": "model",
                    }
                ],
                "has_more": False,
                "last_id": "claude-sonnet-4-6",
            },
        ]

        def fake_get(url, headers=None, params=None, timeout=None):
            calls.append((url, headers, params, timeout))
            return FakeResponse(pages[len(calls) - 1])

        entries = model_catalog.fetch_provider_models(
            "Claude", "anthropic-key", request_get=fake_get
        )

        self.assertEqual(calls[0][0], model_catalog.ANTHROPIC_MODELS_ENDPOINT)
        self.assertEqual(calls[0][1]["x-api-key"], "anthropic-key")
        self.assertEqual(calls[0][1]["anthropic-version"], "2023-06-01")
        self.assertEqual(calls[1][2]["after_id"], "claude-opus-4-6")
        self.assertEqual(
            [entry.model_id for entry in entries],
            ["claude-opus-4-6", "claude-sonnet-4-6"],
        )

    def test_list_provider_models_falls_back_without_api_key(self):
        entries = model_catalog.list_provider_models("OpenAI", api_key=None)

        self.assertIn("gpt-4o-mini", [entry.model_id for entry in entries])

    def test_model_fetcher_delegates_to_shared_catalog(self):
        if llm.ModelFetcher is None:
            self.skipTest("PySide6 is not available")

        original = model_catalog.fetch_provider_models

        def fake_fetch(provider, api_key=None, request_get=None):
            self.assertEqual(provider, "Claude")
            self.assertEqual(api_key, "anthropic-key")
            return [
                model_catalog.ModelCatalogEntry(
                    provider="Claude",
                    model_id="claude-new-future-model",
                    display_name="Claude New Future Model",
                )
            ]

        model_catalog.fetch_provider_models = fake_fetch
        try:
            finished = []
            errors = []
            fetcher = llm.ModelFetcher("Claude", "anthropic-key")
            fetcher.finished.connect(lambda provider, models: finished.append((provider, models)))
            fetcher.error.connect(errors.append)

            fetcher.run()

            self.assertEqual(errors, [])
            self.assertEqual(finished, [("Claude", ["claude-new-future-model"])])
        finally:
            model_catalog.fetch_provider_models = original

    def test_future_openai_reasoning_models_use_responses_api(self):
        self.assertTrue(llm._openai_uses_responses_api("o5-mini"))
        self.assertTrue(llm._openai_uses_responses_api("gpt-5.7-pro"))
        self.assertFalse(llm._openai_uses_responses_api("gpt-4o"))


if __name__ == "__main__":
    unittest.main()
