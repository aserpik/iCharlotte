# Chat Legal Research Output Mode Design

## Goal

Add a selectable output mode for Chat Legal Research so the user can choose between the current fast conversational answer and a more polished legal research memo format. The change must preserve the existing verified-authority guardrails: Chat may cite only authorities selected by the legal research packet, and the audit trail must remain available for review.

## User Experience

The Chat legal research controls will include an output selector near the existing source settings:

- Quick Answer: current-style conversational answer, optimized for speed and directness.
- Research Memo: polished report-style answer, organized like a concise legal research memo.

Quick Answer will remain the default to avoid changing existing behavior unexpectedly.

Research Memo output will use this structure unless the user's prompt clearly asks for a different structure:

1. Summary
2. Governing Rule
3. Best Supporting Cases
4. Limitations / Adverse Authority
5. Suggested Argument Framing
6. Research Basis

The main answer should read like attorney work product, not a retrieval transcript. Raw searches, source labels, quote verification, and selection reasons should remain in the appended Legal Research Basis block.

## Architecture

Add a small enum to `icharlotte_core.chat.legal_research`:

- `ChatResearchOutputMode.QUICK_ANSWER`
- `ChatResearchOutputMode.RESEARCH_MEMO`

Add the selected mode to the research packet or chat send payload so the final answer prompt can vary by mode without changing retrieval. Retrieval, candidate selection, quote verification, CourtListener fallback behavior, and deterministic citation checking remain shared.

The UI should persist the selected output mode using the same settings pattern as the existing firm authority, local corpus, and CourtListener mode controls.

## Prompting

Current legal research mode already appends a verified authority block and instructs the model not to cite outside it. That guardrail stays.

The prompt change should separate two responsibilities:

- Citation guardrail: always cite only authorities in `[CHAT LEGAL RESEARCH AUTHORITY]`; do not add citations from memory.
- Presentation style: only Research Memo mode receives the memo-format instruction.

The current instruction requiring the answer model to include a "Research Basis" section should be removed from the shared guardrail prompt because the UI already appends a Legal Research Basis block. Research Memo may include a short "Research Basis" heading in the polished answer only if it summarizes methodology without duplicating the appendix.

## Authority Roles

Research Memo mode should classify selected authorities for organization. Initial roles can be derived from existing selection data and simple heuristics:

- foundational: older Supreme Court or seminal rule cases.
- direct: cases that squarely support the requested proposition.
- limiting: cases that refine or narrow the requested proposition.
- adverse: cases that cut against the requested proposition or must be distinguished.
- background: useful context that is not central to the argument.

This role field is presentation metadata only. It must not make an unverified authority citeable.

## Data Flow

1. User enables Chat Legal Research and selects source settings.
2. User selects Quick Answer or Research Memo.
3. Existing research service extracts propositions, searches selected sources, verifies quotes, and returns selected authorities.
4. Chat builds the augmented system prompt:
   - Shared citation guardrails.
   - Verified authority block.
   - Optional Research Memo formatting instruction.
5. The model drafts the answer.
6. Existing deterministic citation check runs against known selected case names.
7. UI appends the Legal Research Basis block for auditability.
8. The conversation saves the answer plus appended basis, as it does today.

## Error Handling

If no verified authorities are found, behavior remains fail-closed. The selected output mode must not cause Chat to answer from memory.

If role classification fails, Research Memo should still produce a memo using the selected authorities in citation order. The failure should be non-fatal and should not affect the Legal Research Basis block.

If the user explicitly asks for a format that conflicts with Research Memo mode, the user prompt should control the visible format, but citation guardrails still apply.

## Testing

Add or update focused tests for:

- The new output mode enum and default value.
- UI persistence for Quick Answer vs Research Memo.
- `_run_chat_legal_research` or the send payload carrying the selected output mode.
- Quick Answer preserving the existing guardrail prompt without memo-format instructions.
- Research Memo adding the memo-format instruction while still including the verified authority block.
- The UI not duplicating a full Research Basis in the main answer and appended Legal Research Basis block.
- Deterministic citation checking still running after Research Memo answers.

Manual smoke test:

1. Run iCharlotte.
2. Enable Chat Legal Research.
3. Ask the nuisance/trespass occupancy question in Quick Answer mode.
4. Ask the same question in Research Memo mode.
5. Confirm Research Memo is more polished, cites only selected authorities, and still shows the appended Legal Research Basis with View links.

## Non-Goals

- Do not change retrieval ranking, CourtListener fallback, local corpus indexing, or quote verification in this feature.
- Do not add new legal data sources.
- Do not remove the appended Legal Research Basis audit trail.
- Do not make Research Memo the default yet.
