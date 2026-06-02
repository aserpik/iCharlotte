# Wizard Launcher — Card Grid Categories + Search

**Date:** 2026-06-02
**Status:** Approved design, pending implementation plan
**Scope:** Wizard Mode launcher only (`icharlotte_core/ui/wizard/`). No changes to task execution, settings pages, or persistence.

## Problem

The Wizard launcher (`wizard_tab.py`) renders all tasks as a flat 3-wide grid of cards. With 11 tasks today and more planned (see the "new task types" track), a flat grid is hard to scan and offers no way to find a task by the legal term a user has in mind ("RFP", "IME") rather than the task's official name.

## Goals

- Group task cards under categories so the catalog is scannable.
- Add a search box that filters cards live, matching legal jargon as well as task names.
- Keep the change scoped: Recent Tasks section and overall page chrome stay as they are.

## Non-Goals

- No drag-to-reorder or user-customizable categories.
- No persistence of search/filter state across launcher visits.
- No changes to how tasks run, or to the Settings/Status/Output flow.

## Chosen Approach — Layout "A" (Grouped Sections)

Considered three layouts: (A) grouped sections with category headers, (B) filter chips + single grid, (C) left category rail + grid. Selected **A**: every task is visible at a glance under its category header, which suits a browse-the-catalog mental model, and search collapses it to the relevant subset.

## Data Model — `registry.py`

Add two optional fields to `TaskSpec`:

```python
category: str = "General"          # must be a value in CATEGORY_ORDER
keywords: List[str] = field(default_factory=list)   # search aliases
```

Add a module-level ordering constant:

```python
CATEGORY_ORDER = ["Discovery", "Medical", "Motions & Drafting", "General"]
```

Defaults keep existing/edge cases safe: a spec with no `category` lands in "General"; no `keywords` means search falls back to title + description only for that task.

### Category + keyword assignments

| Task | Category | Keyword aliases (starter set) |
|------|----------|-------------------------------|
| Summarize Discovery | Discovery | responses, RFP, RFA, interrogatory, form interrogatories, special interrogatories, production |
| Summarize Depositions | Discovery | depo, transcript, testimony, witness |
| Depo Prep | Discovery | depo, outline, questions, examination, prepare |
| Respond to Discovery | Discovery | RFP, RFA, interrogatory, form interrogatories, objections, propounded, responses |
| Subpoena Tracker | Discovery | subpoena, SDT, records, deposition subpoena, tracker |
| Medical Records Review | Medical | medical, records, chronology, IME, billing, MRI, treatment |
| Med Chron Analysis | Medical | chronology, chron, analysis, medical, gaps, billing |
| Med Record Extractor | Medical | extract, pages, PDF, records, Bates, exhibit |
| Oppose a Motion | Motions & Drafting | motion, opposition, oppose, MSJ, brief, memorandum, demurrer |
| Summarize Documents | General | summary, summarize, document, general |
| Chat | General | chat, ask, assistant, question |

(Keyword lists are curated, not exhaustive; easy to extend per task.)

## Layout — `wizard_tab.py`

- A search `QLineEdit` is pinned at the top of the card area (above the grid, below the existing page header/subtitle).
- Cards render under category headers in `CATEGORY_ORDER`. Each header shows a count, e.g. `Discovery · 5`. Categories with zero (matching) tasks are not rendered.
- Recent Tasks stays exactly where it is: the collapsible bottom splitter section, unchanged.
- Each category section reuses the existing `TaskCard` widget; only the grouping/headers are new.

## Search Behavior

- **Live filter** as the user types, with a light debounce.
- **Match:** case-insensitive substring against the union of `title + description + keywords`.
- **Grouping during search:** matches stay under their category headers; headers with no matches disappear; counts update to reflect matches.
- **Clearing** the box restores the full grouped view.
- **Empty result:** a single muted line — `No tasks match "{query}".`
- **Keyboard:** the search box autofocuses when the Wizard launcher is shown; **Esc** clears it. Filter state resets on each return to the launcher (not persisted).

## Components / Boundaries

- `registry.py` — owns the taxonomy (categories, order, keyword aliases). Single source of truth.
- `wizard_tab.py` — owns rendering: builds grouped sections from `list_tasks()` + `CATEGORY_ORDER`, owns the search box and the filter function `_matches(spec, query) -> bool`.
- A small pure helper (e.g. `filter_tasks(specs, query) -> dict[str, list[TaskSpec]]`) keeps the match/group logic testable without Qt.

## Testing

- **Pure logic (no Qt):**
  - Every `TaskSpec.category` is in `CATEGORY_ORDER`.
  - `filter_tasks` returns expected tasks for representative queries: `"rfp"` → Respond to Discovery + Summarize Discovery; `"ime"` → Medical Records Review; `""` → all, grouped in order.
  - Empty/whitespace query returns the full grouped set; nonsense query returns no groups.
  - Category counts equal the number of cards rendered.
- **pytest-qt** (mirrors `tests/test_wizard/test_theme_and_scaffold.py`):
  - Typing into the search box updates visible cards; Esc clears; autofocus on show.

## Risks / Notes

- Keyword lists are subjective; treat the table above as a starting point, tune over time.
- `"Motions & Drafting"` has one card today; kept deliberately as the home for future drafting tasks.
- Emoji icon glyphs are unchanged.
