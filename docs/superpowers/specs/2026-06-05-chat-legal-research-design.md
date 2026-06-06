# Chat Legal Research Source Selection Design

## Scope

Replace the current Chat tab Legal Research checkbox behavior only. This design
does not change the Word assistant legal research checkbox, the Oppose a Motion
wizard, or the Generate Motion wizard.

The Chat tab should let a user run grounded California legal research against
any selected combination of:

- firm/sample-motion authority;
- the local California case-law corpus;
- the live CourtListener API.

When research mode is enabled, the final Chat answer must show its research
basis: what was searched, why authorities were selected, which source found
them, and short verbatim quotes from the cited case law so the user can verify
the answer.

## Existing Context

The Chat tab currently has a `Legal Research` checkbox in
`icharlotte_core/ui/tabs.py`. The existing path attempts to run
`LegalResearchEngine` before the streaming LLM response, injects the resulting
authority block into the system prompt, and later appends a source list.

The repo already has useful components that should be reused:

- `icharlotte_core/legal_research/local_corpus/corpus.py` provides
  `LocalCaseCorpus.search_opinions(...)`, with keyword and semantic search over
  the local California case-law corpus.
- `icharlotte_core/firm_briefs/provider.py` and
  `icharlotte_core/firm_briefs/index.py` provide firm/sample-motion authority
  candidates and proposition-level semantic reranking.
- `icharlotte_core/legal_research/sources/courtlistener.py` provides live
  CourtListener search, citation lookup, opinion fetch, and California reporter
  citation preference.
- `icharlotte_core/opposition/argument_research.py` contains useful patterns
  for hybrid search, relevant excerpt selection, LLM reranking, and rejecting
  fabricated quotes.

The new Chat feature should use those pieces, but should not directly reuse the
motion-specific opposition workflow as the public Chat API.

## User Interface

The Chat toolbar keeps the existing `Legal Research` checkbox as the mode
toggle. Add a compact source selector beside it. The selector should expose:

- `Firm/sample-motion authority`: checkbox.
- `Local California corpus`: checkbox.
- `CourtListener API`: three-option mode:
  - `Off`: never call CourtListener.
  - `Fallback/current-law`: call CourtListener only when selected local sources
    are missing, stale, thin, or the user's prompt asks for current authority.
  - `Always search`: include CourtListener in every researched query.

The first-run default should be:

- firm/sample-motion authority on;
- local California corpus on;
- CourtListener fallback/current-law.

After the user changes these source settings, persist the changed choices and
reuse them in later Chat sessions.

If both local sources are off, the UI should not leave CourtListener in
fallback/current-law mode because there is nothing to fall back from. In that
case, CourtListener should be either `Off` or `Always search`.

## Settings Persistence

Persist the Chat legal research source choices separately from the Word
assistant legal research setting. A small settings structure is sufficient:

```json
{
  "chat_legal_research": {
    "enabled_default": false,
    "firm_authority": true,
    "local_corpus": true,
    "courtlistener_mode": "fallback_current_law"
  }
}
```

The exact storage file can follow the Chat tab's existing settings pattern. The
important requirement is that changing Chat source choices does not affect the
Word assistant or wizard legal research behavior.

## Architecture

Add a Chat-specific backend service, for example
`icharlotte_core/chat/legal_research.py`. This service owns the checkbox
behavior and returns a structured research packet for prompt injection and UI
display.

The public API should accept:

- user message text;
- attached file context text;
- selected source settings;
- current model/provider callback for LLM query extraction and reranking;
- optional progress callback for Chat status messages.

It should return a `ChatResearchPacket` with:

- extracted research questions or propositions;
- source settings actually used;
- searches run;
- selected authorities;
- source warnings;
- prompt-ready authority block;
- display-ready research basis summary.

The Chat tab should remain responsible for UI state, progress display, and
passing the research packet into the final LLM call. The orchestrator should
remain Qt-free so it can be tested directly.

## Retrieval Flow

When Legal Research is enabled:

1. Build research input from the user message and attached context.
2. Use the selected LLM to extract 1 to 5 focused legal propositions or search
   questions.
3. For each proposition, search only the selected sources:
   - firm/sample-motion authority if enabled;
   - local California corpus if enabled;
   - CourtListener according to its selected mode.
4. Merge candidates by stable case identity where possible:
   - CourtListener cluster id;
   - local corpus case uid;
   - normalized reporter citation;
   - normalized case name plus year as a fallback.
5. Preserve all source provenance on merged candidates.
6. Fetch opinion text or source text for candidates when available.
7. Select relevant excerpts from actual source text.
8. Ask the LLM to select the best authorities and explain the selection.
9. Reject any selected authority whose quote is not found verbatim in the
   retrieved source text.
10. Build a prompt-ready packet that the final Chat answer may cite.

## Source Semantics

### Firm/sample-motion authority

Use the existing firm brief/sample-motion index to retrieve proposition-level
authority candidates. These results are valuable because they show what the
firm has previously cited for similar points.

Firm authority must carry provenance:

- source brief path or label;
- proposition text from the sample;
- quoted passage harvested from the sample;
- whether the citation was verified against the local corpus or CourtListener.

If a firm authority citation cannot be resolved to actual case text, it may be
shown as unverified firm authority, but the final Chat answer should not treat
it as a verified citation unless another selected source verifies it.

### Local California corpus

Use `LocalCaseCorpus.search_opinions(..., semantic=True)` and keyword search
patterns already present in the repo. Natural-language proposition text should
lead semantic search. Boolean or CourtListener-style search syntax can still be
used for keyword searches where helpful.

The local corpus should be the preferred source for speed and reduced API usage
when it is selected and available.

### CourtListener API

CourtListener behavior depends on the user-selected mode:

- `Off`: never call CourtListener, including as fallback.
- `Fallback/current-law`: call CourtListener only when selected local sources
  are unavailable, stale, thin, or the prompt asks for current law.
- `Always search`: call CourtListener for each researched proposition.

When CourtListener is selected in any calling mode and no API token is
available, the feature should report that limitation clearly. If no other
selected source can provide verified authority, research mode should stop before
the final LLM answer.

## Reranking And Quote Requirements

The reranker prompt should require each selected authority to return:

- candidate id;
- reason selected;
- supported proposition;
- verbatim quote copied from the provided source excerpt;
- any caveat or limitation.

The service must verify that the returned quote appears in the actual retrieved
source text after whitespace normalization. If the quote does not match, drop
that selection rather than passing it into the final answer.

The selected authority model should include:

- case name;
- citation;
- year/date;
- court;
- source list;
- URL or local source reference;
- reason selected;
- supported proposition;
- verbatim quote;
- verification status.

## Prompt Injection And Final Answer

When research succeeds, inject a structured authority packet into the final Chat
system prompt. The prompt should tell the model:

- cite only authorities in the packet;
- do not invent or recall citations from memory;
- state when selected sources do not support a requested proposition;
- include a concise `Research Basis` section in the answer.

The `Research Basis` section should include:

- searches run;
- sources searched;
- cases cited;
- why each cited case was selected;
- short supporting quotes;
- warnings such as stale corpus, missing token, thin results, or unverified
  firm authority.

The existing deterministic citation check can still run after the streamed
answer, but it should be treated as a backstop rather than the primary guard.

## Failure Handling

Research mode should fail closed. In this feature, that means the Chat tab
should not produce a research-backed legal answer when the selected sources
cannot produce a verified research basis.

Examples:

- If Legal Research is enabled but no selected source is usable, show an error
  and do not start the final LLM answer.
- If only CourtListener is selected and no API token exists, show an error and
  do not start the final LLM answer.
- If selected local sources are missing but another selected source is usable,
  warn and continue with the usable selected source.
- If the LLM selects a quote that is not present in the source text, reject that
  authority.
- If no selected verified authority supports the proposition, the final answer
  should say the selected sources did not provide support rather than inventing
  a case.

This does not mean the application crashes or disables ordinary Chat. It means
the legal research mode will not pretend to have verified authority when it
does not.

## Tests

Add focused tests for the new Chat legal research service and the Chat tab
wiring.

Service tests should cover:

- source settings normalization, including the rule that CourtListener fallback
  is invalid when both local sources are off;
- firm-only search;
- local-corpus-only search;
- CourtListener-off never calls the live client;
- CourtListener fallback calls the live client only for thin, stale, missing,
  or current-law conditions;
- CourtListener always-search calls the live client for every proposition;
- missing-token behavior;
- merge and provenance preservation;
- rejection of fabricated or non-verbatim quotes;
- prompt packet construction with Research Basis content.

Chat tab tests should cover:

- source selector defaults;
- persistence of changed source selections;
- research packet injection into the final LLM call;
- visible failure when no selected source is usable;
- source warnings shown in the Chat transcript.

Verification should prefer focused tests around:

- `tests/test_legal_research/test_courtlistener.py`;
- `tests/test_legal_research/test_local_corpus`;
- relevant `tests/test_firm_briefs`;
- new Chat legal research tests.

Avoid broad unrelated legal-research suites if the runtime still lacks optional
dependencies used by unrelated CA courts or LegInfo tests.

## Non-Goals

This design does not:

- modify Word assistant legal research behavior;
- modify wizard motion drafting behavior;
- add a full citation-review panel;
- build or rebuild the local corpus or firm-brief index;
- guarantee Shepard's or KeyCite treatment beyond the available local and
  CourtListener signals.

## Approval State

The user approved:

- Chat tab only;
- a new Chat-specific legal research orchestrator;
- persistent source selection;
- user-controllable CourtListener modes;
- no CourtListener fallback when CourtListener is set to off;
- fail-closed behavior for unsupported research mode.
