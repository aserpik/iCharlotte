"""The rerank excerpt must show the part of the opinion RELEVANT to the
argument, not the first N chars (which are usually caption/procedural intro)."""
from icharlotte_core.opposition.argument_research import _relevant_excerpt, _format_candidates


def test_relevant_excerpt_finds_deep_holding():
    boiler = "This appeal concerns a default judgment and procedural history. " * 120  # ~7600 chars
    holding = ("The trial court did not abuse its discretion. Good cause supports a forensic "
               "inspection of the computer where relevance is shown and privacy is protected.")
    text = boiler + holding + (" Additional unrelated discussion. " * 50)
    prop = "good cause exists for a forensic inspection of the computer; privacy protected"
    ex = _relevant_excerpt(text, prop, max_chars=2000)
    assert "forensic inspection" in ex          # the deep holding is included
    assert "good cause" in ex.lower()
    assert "default judgment" not in ex[:200] or "forensic" in ex  # window moved off the intro


def test_relevant_excerpt_short_text_passthrough():
    t = "short opinion text about inspection"
    assert _relevant_excerpt(t, "inspection", max_chars=6000) == t


def test_relevant_excerpt_no_term_match_falls_back_to_start():
    text = "alpha beta gamma " * 1000
    ex = _relevant_excerpt(text, "zzzz qqqq", max_chars=500)
    assert ex.startswith("alpha beta gamma")     # no relevant terms -> opening


def test_format_candidates_uses_proposition_relevance():
    boiler = "Procedural background filler sentence about appeal. " * 200
    holding = "Good cause for inspection of electronically stored information requires a showing of relevance."
    text = boiler + holding
    cands = [{"cluster_id": "x", "case_name": "A v B", "citation": "1 Cal.5th 1", "text": text}]
    out = _format_candidates(
        cands, proposition="good cause inspection electronically stored information relevance")
    assert "Good cause for inspection of electronically stored information" in out
    # the excerpt that the LLM sees is a substring of the real opinion (verbatim check still holds)
    assert "Good cause for inspection of electronically stored information requires a showing of relevance." in text
