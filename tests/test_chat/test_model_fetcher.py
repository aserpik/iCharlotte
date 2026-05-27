import unittest

from icharlotte_core import llm


class ModelFetcherTests(unittest.TestCase):
    def test_gemini_models_are_fetched_over_rest_when_sdk_missing(self):
        if llm.ModelFetcher is None:
            self.skipTest("PySide6 is not available")

        original_genai = llm.genai
        original_get = llm.requests.get
        llm.genai = None

        class FakeResponse:
            status_code = 200
            text = "OK"

            def json(self):
                return {
                    "models": [
                        {
                            "name": "models/gemini-3.1-pro-preview",
                            "supportedGenerationMethods": ["generateContent"],
                        },
                        {
                            "name": "models/gemini-3.5-flash",
                            "supportedGenerationMethods": ["generateContent"],
                        },
                        {
                            "name": "models/gemini-new-live-model",
                            "supportedGenerationMethods": ["generateContent"],
                        },
                        {
                            "name": "models/gemini-embedding-only",
                            "supportedGenerationMethods": ["embedContent"],
                        },
                    ]
                }

        calls = []

        def fake_get(url, params=None, timeout=None):
            calls.append((url, params, timeout))
            return FakeResponse()

        llm.requests.get = fake_get
        try:
            finished = []
            errors = []

            fetcher = llm.ModelFetcher("Gemini", "dummy-key")
            fetcher.finished.connect(lambda provider, models: finished.append((provider, models)))
            fetcher.error.connect(errors.append)

            fetcher.run()

            self.assertEqual(errors, [])
            self.assertEqual(len(finished), 1)
            provider, models = finished[0]
            self.assertEqual(provider, "Gemini")
            self.assertIn("gemini-3.1-pro-preview", models)
            self.assertIn("gemini-3.5-flash", models)
            self.assertIn("gemini-new-live-model", models)
            self.assertNotIn("gemini-embedding-only", models)
            self.assertFalse(any(model.startswith("Error:") for model in models))
            self.assertEqual(calls[0][0], llm.GEMINI_MODELS_ENDPOINT)
            self.assertEqual(calls[0][1]["key"], "dummy-key")
        finally:
            llm.genai = original_genai
            llm.requests.get = original_get

    def test_gemini_models_fall_back_to_curated_list_when_live_fetch_fails(self):
        if llm.ModelFetcher is None:
            self.skipTest("PySide6 is not available")

        original_genai = llm.genai
        original_get = llm.requests.get
        llm.genai = None
        llm.requests.get = lambda *args, **kwargs: (_ for _ in ()).throw(Exception("network down"))
        try:
            finished = []
            errors = []

            fetcher = llm.ModelFetcher("Gemini", "dummy-key")
            fetcher.finished.connect(lambda provider, models: finished.append((provider, models)))
            fetcher.error.connect(errors.append)

            fetcher.run()

            self.assertEqual(errors, [])
            self.assertEqual(len(finished), 1)
            provider, models = finished[0]
            self.assertEqual(provider, "Gemini")
            self.assertEqual(models, sorted(llm.CURATED_GEMINI_MODELS, reverse=True))
        finally:
            llm.genai = original_genai
            llm.requests.get = original_get


if __name__ == "__main__":
    unittest.main()
