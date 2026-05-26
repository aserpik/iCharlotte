# Oppose a Motion Wizard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a guided Wizard Mode task that drafts a California civil opposition memorandum, verifies legal citations through CourtListener-backed checks, and lets the user save the final Word document only through Save As.

**Architecture:** Build a new `icharlotte_core/opposition/` service layer for state, extraction, outline handling, legal research orchestration, citation verification, drafting, and Word assembly. Add a custom Wizard task tab for the multi-step UI because the standard `InProcessTaskTab` only passes a string output path to its output page, while this task needs a richer draft/citation state for the right-side source drawer. Keep real LLM and CourtListener calls behind injectable callables so automated tests stay deterministic.

**Tech Stack:** Python, PySide6, python-docx, existing `DocumentProcessor`, existing `LLMConfig`/`call_llm`, existing CourtListener client, existing Wizard Mode registry/routing/persistence.

---

## Scope Notes

The current checkout has many unrelated modified and deleted files. Do not revert, stage, or commit unrelated work. This plan touches only the files listed below.

The visual companion created `.superpowers/` files during brainstorming. They are not part of this feature implementation.

## File Structure

Create these files:

- `icharlotte_core/opposition/__init__.py` - package marker and public imports.
- `icharlotte_core/opposition/models.py` - serializable dataclasses for motion metadata, outline nodes, section plan items, draft documents, and citation verification records.
- `icharlotte_core/opposition/extraction.py` - file type validation and extraction wrappers around `DocumentProcessor`.
- `icharlotte_core/opposition/outline.py` - outline tree operations and selected-section-plan conversion.
- `icharlotte_core/opposition/motion_analyzer.py` - prompt construction and strict JSON parsing for motion metadata and outline proposals.
- `icharlotte_core/opposition/drafter.py` - prompt construction and draft parsing from selected outline and authority records.
- `icharlotte_core/opposition/citation_verifier.py` - CourtListener lookup mapping, support-passage checks, and replacement suggestion records.
- `icharlotte_core/opposition/assembler.py` - caption/template-aware Word preview assembly and validation wrapper.
- `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` - confirmation, outline editor, output preview, source drawer, and worker classes.
- `tests/test_opposition/__init__.py`
- `tests/test_opposition/test_models.py`
- `tests/test_opposition/test_extraction.py`
- `tests/test_opposition/test_outline.py`
- `tests/test_opposition/test_motion_analyzer.py`
- `tests/test_opposition/test_citation_verifier.py`
- `tests/test_opposition/test_assembler.py`
- `tests/test_wizard/test_oppose_motion_registry.py`
- `tests/test_wizard/test_oppose_motion_page.py`

Modify these files:

- `icharlotte_core/ui/wizard/registry.py` - add `oppose_motion` task.
- `icharlotte_core/ui/wizard/task_routing.py` - map `oppose_motion` to the custom builder and skip the generic initial picker.
- `icharlotte_core/ui/wizard/in_process_task_tab.py` - add `build_oppose_motion_tab`.
- `icharlotte_core/legal_research/sources/courtlistener.py` - add citation lookup and opinion metadata helpers.
- `icharlotte_core/word_validator.py` - add lightweight `validate_opposition_docx`.
- `iCharlotte.py` - restore/reopen in-process task tabs through their builders when needed.

---

### Task 1: Register The Wizard Task And Route To A Builder

**Files:**
- Test: `tests/test_wizard/test_oppose_motion_registry.py`
- Modify: `icharlotte_core/ui/wizard/registry.py`
- Modify: `icharlotte_core/ui/wizard/task_routing.py`
- Modify: `icharlotte_core/ui/wizard/in_process_task_tab.py`

- [ ] **Step 1: Write the failing registry/routing test**

Create `tests/test_wizard/test_oppose_motion_registry.py`:

```python
import unittest

from icharlotte_core.ui.wizard.registry import get_task, list_tasks
from icharlotte_core.ui.wizard.task_routing import (
    get_in_process_task_builder_name,
    requires_initial_file_picker,
)


class OpposeMotionRegistryTests(unittest.TestCase):
    def test_task_registered(self):
        ids = {task.task_id for task in list_tasks()}
        self.assertIn("oppose_motion", ids)
        spec = get_task("oppose_motion")
        self.assertEqual(spec.title, "Oppose a Motion")
        self.assertEqual(
            spec.default_folders,
            ["MOTIONS", "PLEADINGS", "DISCOVERY"],
        )
        self.assertEqual(spec.script_name, "")

    def test_task_uses_custom_in_process_builder_without_generic_picker(self):
        self.assertEqual(
            get_in_process_task_builder_name("oppose_motion"),
            "build_oppose_motion_tab",
        )
        self.assertFalse(requires_initial_file_picker("oppose_motion"))


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_registry.py -q
```

Expected: FAIL because `oppose_motion` is not registered.

- [ ] **Step 3: Add the task spec**

In `icharlotte_core/ui/wizard/registry.py`, add this entry before `"chat"`:

```python
    "oppose_motion": TaskSpec(
        task_id="oppose_motion",
        title="Oppose a Motion",
        description="Draft and verify an opposition memorandum for a California civil motion.",
        icon_glyph="\U0001F4DD",
        script_name="",
        default_folders=["MOTIONS", "PLEADINGS", "DISCOVERY"],
    ),
```

- [ ] **Step 4: Add the route**

In `icharlotte_core/ui/wizard/task_routing.py`, extend `_IN_PROCESS_TASK_BUILDERS`:

```python
_IN_PROCESS_TASK_BUILDERS = {
    "subpoena_tracker": "build_subpoena_tab",
    "respond_to_discovery": "build_respond_to_discovery_tab",
    "oppose_motion": "build_oppose_motion_tab",
}
```

- [ ] **Step 5: Add a temporary builder that imports the real page**

At the bottom of `icharlotte_core/ui/wizard/in_process_task_tab.py`, add:

```python
def build_oppose_motion_tab(
    spec,
    case_path: str,
    file_number: str,
    parent: QWidget | None,
):
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
        build_oppose_motion_tab as _build,
    )

    return _build(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        parent=parent,
    )
```

This import will fail until Task 9 creates the page. Keep the test in this task focused on registry/routing, not invoking the builder.

- [ ] **Step 6: Run the registry test**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_registry.py -q
```

Expected: PASS.

- [ ] **Step 7: Commit**

```powershell
git add tests/test_wizard/test_oppose_motion_registry.py icharlotte_core/ui/wizard/registry.py icharlotte_core/ui/wizard/task_routing.py icharlotte_core/ui/wizard/in_process_task_tab.py
git commit -m "feat(wizard): register oppose motion task"
```

---

### Task 2: Add Opposition Models And File Extraction

**Files:**
- Create: `icharlotte_core/opposition/__init__.py`
- Create: `icharlotte_core/opposition/models.py`
- Create: `icharlotte_core/opposition/extraction.py`
- Test: `tests/test_opposition/__init__.py`
- Test: `tests/test_opposition/test_models.py`
- Test: `tests/test_opposition/test_extraction.py`

- [ ] **Step 1: Write model serialization tests**

Create `tests/test_opposition/__init__.py` as an empty file.

Create `tests/test_opposition/test_models.py`:

```python
from icharlotte_core.opposition.models import (
    CitationVerification,
    DraftDocument,
    MotionMetadata,
    OutlineNode,
)


def test_motion_metadata_round_trip():
    meta = MotionMetadata(
        motion_type="Motion for Summary Judgment",
        moving_party="Defendant City",
        opposing_party="Plaintiff Doe",
        relief_requested="summary judgment",
        hearing_date="2026-06-12",
        opposition_due_date="2026-05-29",
        procedural_posture="pretrial",
        principal_arguments=["no duty", "no causation"],
        opposition_posture="triable issues exist",
    )

    restored = MotionMetadata.from_dict(meta.to_dict())

    assert restored == meta


def test_outline_node_round_trip_preserves_children_and_selected_state():
    root = OutlineNode(
        id="n1",
        text="The motion fails because triable issues exist",
        selected=True,
        children=[
            OutlineNode(
                id="n2",
                text="Defendant ignores conflicting testimony",
                selected=False,
            )
        ],
    )

    restored = OutlineNode.from_dict(root.to_dict())

    assert restored.id == "n1"
    assert restored.children[0].id == "n2"
    assert restored.children[0].selected is False


def test_draft_document_round_trip_keeps_citation_records():
    draft = DraftDocument(
        title="Opposition to Motion for Summary Judgment",
        body_text="Argument with Rowland v. Christian (1968) 69 Cal.2d 108.",
        citations=[
            CitationVerification(
                citation_text="Rowland v. Christian (1968) 69 Cal.2d 108",
                normalized_citation="69 Cal.2d 108",
                status="verified",
                case_name="Rowland v. Christian",
                court="California Supreme Court",
                date="1968-08-08",
                opinion_url="https://www.courtlistener.com/opinion/...",
                supporting_passage="Everyone is responsible...",
                warning="",
                replacement_candidates=[],
            )
        ],
    )

    restored = DraftDocument.from_dict(draft.to_dict())

    assert restored.citations[0].status == "verified"
    assert restored.body_text.startswith("Argument with")
```

- [ ] **Step 2: Write extraction tests**

Create `tests/test_opposition/test_extraction.py`:

```python
from pathlib import Path

from icharlotte_core.document_processor import ExtractionMethod, ExtractResult
from icharlotte_core.opposition.extraction import (
    SUPPORTED_CONTEXT_EXTENSIONS,
    SUPPORTED_MOTION_EXTENSIONS,
    extract_document_text,
    is_supported_context_file,
    is_supported_motion_file,
)


class FakeProcessor:
    def extract_text(self, file_path, ocr_enabled=True):
        return ExtractResult(
            text=f"text from {Path(file_path).name}",
            page_count=1,
            extraction_method=ExtractionMethod.NATIVE,
            char_count=10,
            file_path=file_path,
        )


def test_supported_motion_file_types():
    assert SUPPORTED_MOTION_EXTENSIONS == {".pdf", ".docx"}
    assert is_supported_motion_file("motion.pdf")
    assert is_supported_motion_file("motion.DOCX")
    assert not is_supported_motion_file("motion.txt")


def test_supported_context_file_types():
    assert SUPPORTED_CONTEXT_EXTENSIONS == {".pdf", ".docx", ".txt", ".msg"}
    assert is_supported_context_file("facts.msg")
    assert is_supported_context_file("facts.txt")
    assert not is_supported_context_file("facts.xlsx")


def test_extract_document_text_uses_injected_processor(tmp_path):
    source = tmp_path / "motion.pdf"
    source.write_text("unused")

    result = extract_document_text(str(source), processor=FakeProcessor())

    assert result.text == "text from motion.pdf"
    assert result.file_path == str(source)
```

- [ ] **Step 3: Run the failing tests**

Run:

```powershell
pytest tests/test_opposition/test_models.py tests/test_opposition/test_extraction.py -q
```

Expected: FAIL because the `opposition` package does not exist.

- [ ] **Step 4: Create package exports**

Create `icharlotte_core/opposition/__init__.py`:

```python
"""Opposition memorandum drafting services for Wizard Mode."""

from .models import (
    CitationVerification,
    DraftDocument,
    MotionMetadata,
    OutlineNode,
    SectionPlanItem,
)

__all__ = [
    "CitationVerification",
    "DraftDocument",
    "MotionMetadata",
    "OutlineNode",
    "SectionPlanItem",
]
```

- [ ] **Step 5: Create serializable models**

Create `icharlotte_core/opposition/models.py`:

```python
from __future__ import annotations

from dataclasses import asdict, dataclass, field


@dataclass
class MotionMetadata:
    motion_type: str = ""
    moving_party: str = ""
    opposing_party: str = ""
    relief_requested: str = ""
    hearing_date: str = ""
    opposition_due_date: str = ""
    procedural_posture: str = ""
    principal_arguments: list[str] = field(default_factory=list)
    opposition_posture: str = ""

    def required_missing(self) -> list[str]:
        missing: list[str] = []
        if not self.motion_type.strip():
            missing.append("motion_type")
        if not self.relief_requested.strip():
            missing.append("relief_requested")
        if not [arg for arg in self.principal_arguments if arg.strip()]:
            missing.append("principal_arguments")
        return missing

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict | None) -> "MotionMetadata":
        data = dict(data or {})
        args = data.get("principal_arguments")
        if not isinstance(args, list):
            data["principal_arguments"] = []
        return cls(**{k: data.get(k, getattr(cls(), k)) for k in cls().__dataclass_fields__})


@dataclass
class OutlineNode:
    id: str
    text: str
    selected: bool = True
    children: list["OutlineNode"] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "text": self.text,
            "selected": self.selected,
            "children": [child.to_dict() for child in self.children],
        }

    @classmethod
    def from_dict(cls, data: dict) -> "OutlineNode":
        return cls(
            id=str(data.get("id", "")),
            text=str(data.get("text", "")),
            selected=bool(data.get("selected", True)),
            children=[
                cls.from_dict(child)
                for child in data.get("children", [])
                if isinstance(child, dict)
            ],
        )


@dataclass
class SectionPlanItem:
    id: str
    path: list[str]
    level: int

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict) -> "SectionPlanItem":
        return cls(
            id=str(data.get("id", "")),
            path=[str(item) for item in data.get("path", [])],
            level=int(data.get("level", 1)),
        )


@dataclass
class CitationVerification:
    citation_text: str
    normalized_citation: str = ""
    status: str = "exists_support_unconfirmed"
    case_name: str = ""
    court: str = ""
    date: str = ""
    opinion_url: str = ""
    supporting_passage: str = ""
    warning: str = ""
    replacement_candidates: list[dict] = field(default_factory=list)

    def to_dict(self) -> dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict | None) -> "CitationVerification":
        data = dict(data or {})
        return cls(
            citation_text=str(data.get("citation_text", "")),
            normalized_citation=str(data.get("normalized_citation", "")),
            status=str(data.get("status", "exists_support_unconfirmed")),
            case_name=str(data.get("case_name", "")),
            court=str(data.get("court", "")),
            date=str(data.get("date", "")),
            opinion_url=str(data.get("opinion_url", "")),
            supporting_passage=str(data.get("supporting_passage", "")),
            warning=str(data.get("warning", "")),
            replacement_candidates=list(data.get("replacement_candidates", [])),
        )


@dataclass
class DraftDocument:
    title: str = ""
    body_text: str = ""
    citations: list[CitationVerification] = field(default_factory=list)
    preview_path: str = ""

    def to_dict(self) -> dict:
        return {
            "title": self.title,
            "body_text": self.body_text,
            "citations": [citation.to_dict() for citation in self.citations],
            "preview_path": self.preview_path,
        }

    @classmethod
    def from_dict(cls, data: dict | None) -> "DraftDocument":
        data = dict(data or {})
        return cls(
            title=str(data.get("title", "")),
            body_text=str(data.get("body_text", "")),
            citations=[
                CitationVerification.from_dict(item)
                for item in data.get("citations", [])
                if isinstance(item, dict)
            ],
            preview_path=str(data.get("preview_path", "")),
        )
```

- [ ] **Step 6: Create extraction helpers**

Create `icharlotte_core/opposition/extraction.py`:

```python
from __future__ import annotations

import os

from icharlotte_core.document_processor import DocumentProcessor, ExtractResult


SUPPORTED_MOTION_EXTENSIONS = {".pdf", ".docx"}
SUPPORTED_CONTEXT_EXTENSIONS = {".pdf", ".docx", ".txt", ".msg"}


def _extension(path: str) -> str:
    return os.path.splitext(path or "")[1].lower()


def is_supported_motion_file(path: str) -> bool:
    return _extension(path) in SUPPORTED_MOTION_EXTENSIONS


def is_supported_context_file(path: str) -> bool:
    return _extension(path) in SUPPORTED_CONTEXT_EXTENSIONS


def extract_document_text(
    path: str,
    *,
    processor: DocumentProcessor | None = None,
    ocr_enabled: bool = True,
) -> ExtractResult:
    if not path:
        raise ValueError("Document path is required.")
    processor = processor or DocumentProcessor()
    return processor.extract_text(path, ocr_enabled=ocr_enabled)


def extract_context_bundle(
    paths: list[str],
    *,
    processor: DocumentProcessor | None = None,
) -> tuple[str, list[str]]:
    processor = processor or DocumentProcessor()
    parts: list[str] = []
    warnings: list[str] = []
    for path in paths:
        if not is_supported_context_file(path):
            warnings.append(f"Unsupported context file type: {path}")
            continue
        result = processor.extract_text(path)
        if result.success:
            parts.append(f"[CONTEXT FILE: {os.path.basename(path)}]\n{result.text}")
        else:
            warnings.append(result.error or f"Could not extract text from {path}")
    return "\n\n---\n\n".join(parts), warnings
```

- [ ] **Step 7: Run the tests**

Run:

```powershell
pytest tests/test_opposition/test_models.py tests/test_opposition/test_extraction.py -q
```

Expected: PASS.

- [ ] **Step 8: Commit**

```powershell
git add icharlotte_core/opposition/__init__.py icharlotte_core/opposition/models.py icharlotte_core/opposition/extraction.py tests/test_opposition/__init__.py tests/test_opposition/test_models.py tests/test_opposition/test_extraction.py
git commit -m "feat(opposition): add models and extraction helpers"
```

---

### Task 3: Implement Outline Tree Operations

**Files:**
- Create: `icharlotte_core/opposition/outline.py`
- Test: `tests/test_opposition/test_outline.py`

- [ ] **Step 1: Write outline behavior tests**

Create `tests/test_opposition/test_outline.py`:

```python
from icharlotte_core.opposition.models import OutlineNode
from icharlotte_core.opposition.outline import (
    add_child,
    delete_node,
    move_node,
    selected_section_plan,
    update_node_text,
)


def _sample_outline():
    return [
        OutlineNode(
            id="a",
            text="Plaintiff met the burden for opposition",
            selected=True,
            children=[
                OutlineNode(
                    id="a1",
                    text="Evidence creates disputed facts",
                    selected=True,
                    children=[
                        OutlineNode(id="a1x", text="Witness testimony conflicts", selected=True)
                    ],
                )
            ],
        ),
        OutlineNode(id="b", text="The motion is procedurally defective", selected=False),
    ]


def test_selected_section_plan_includes_three_levels():
    plan = selected_section_plan(_sample_outline())

    assert [item.id for item in plan] == ["a", "a1", "a1x"]
    assert plan[2].path == [
        "Plaintiff met the burden for opposition",
        "Evidence creates disputed facts",
        "Witness testimony conflicts",
    ]
    assert plan[2].level == 3


def test_deselected_parent_excludes_selected_children():
    outline = _sample_outline()
    outline[0].selected = False

    assert selected_section_plan(outline) == []


def test_update_add_delete_and_move_node():
    outline = _sample_outline()
    update_node_text(outline, "a1", "Triable issues exist")
    add_child(outline, "a", OutlineNode(id="a2", text="Expert testimony supports opposition"))
    move_node(outline, "a2", direction=-1)
    delete_node(outline, "a1x")

    assert outline[0].children[0].id == "a2"
    assert outline[0].children[1].text == "Triable issues exist"
    assert outline[0].children[1].children == []
```

- [ ] **Step 2: Run the failing tests**

Run:

```powershell
pytest tests/test_opposition/test_outline.py -q
```

Expected: FAIL because `outline.py` does not exist.

- [ ] **Step 3: Implement outline operations**

Create `icharlotte_core/opposition/outline.py`:

```python
from __future__ import annotations

from collections.abc import Iterable

from .models import OutlineNode, SectionPlanItem


def _walk(nodes: list[OutlineNode]) -> Iterable[tuple[OutlineNode, list[OutlineNode]]]:
    for node in nodes:
        yield node, nodes
        yield from _walk(node.children)


def find_node(nodes: list[OutlineNode], node_id: str) -> OutlineNode | None:
    for node, _siblings in _walk(nodes):
        if node.id == node_id:
            return node
    return None


def selected_section_plan(nodes: list[OutlineNode]) -> list[SectionPlanItem]:
    plan: list[SectionPlanItem] = []

    def visit(node: OutlineNode, path: list[str], ancestors_selected: bool) -> None:
        active = ancestors_selected and node.selected
        current_path = path + [node.text]
        if active:
            plan.append(
                SectionPlanItem(
                    id=node.id,
                    path=current_path,
                    level=len(current_path),
                )
            )
            for child in node.children:
                visit(child, current_path, True)

    for root in nodes:
        visit(root, [], True)
    return plan


def update_node_text(nodes: list[OutlineNode], node_id: str, text: str) -> bool:
    node = find_node(nodes, node_id)
    if node is None:
        return False
    node.text = text
    return True


def add_child(nodes: list[OutlineNode], parent_id: str, child: OutlineNode) -> bool:
    parent = find_node(nodes, parent_id)
    if parent is None:
        return False
    parent.children.append(child)
    return True


def delete_node(nodes: list[OutlineNode], node_id: str) -> bool:
    for node, siblings in _walk(nodes):
        if node.id == node_id:
            siblings.remove(node)
            return True
    return False


def move_node(nodes: list[OutlineNode], node_id: str, direction: int) -> bool:
    for node, siblings in _walk(nodes):
        if node.id != node_id:
            continue
        idx = siblings.index(node)
        target = idx + (-1 if direction < 0 else 1)
        if target < 0 or target >= len(siblings):
            return False
        siblings[idx], siblings[target] = siblings[target], siblings[idx]
        return True
    return False
```

- [ ] **Step 4: Run the tests**

Run:

```powershell
pytest tests/test_opposition/test_outline.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/opposition/outline.py tests/test_opposition/test_outline.py
git commit -m "feat(opposition): add outline tree operations"
```

---

### Task 4: Add CourtListener Citation Lookup Helpers

**Files:**
- Modify: `icharlotte_core/legal_research/sources/courtlistener.py`
- Test: `tests/test_legal_research/test_courtlistener.py`

- [ ] **Step 1: Add tests for citation lookup and opinion text preference**

Append these tests to `tests/test_legal_research/test_courtlistener.py`:

```python
class TestCitationLookup(unittest.TestCase):
    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.post")
    def test_lookup_citations_posts_text(self, mock_post):
        resp = MagicMock()
        resp.status_code = 200
        resp.json.return_value = [
            {
                "citation": "69 Cal.2d 108",
                "normalized_citations": ["69 Cal.2d 108"],
                "status": 200,
                "clusters": [{"id": 111, "case_name": "Rowland v. Christian"}],
            }
        ]
        resp.raise_for_status = MagicMock()
        mock_post.return_value = resp

        client = CourtListenerClient(token="tok")
        data = client.lookup_citations("Rowland v. Christian (1968) 69 Cal.2d 108")

        self.assertEqual(data[0]["status"], 200)
        self.assertIn("citation-lookup", mock_post.call_args.args[0])
        self.assertEqual(
            mock_post.call_args.kwargs["data"]["text"],
            "Rowland v. Christian (1968) 69 Cal.2d 108",
        )

    @patch("icharlotte_core.legal_research.sources.courtlistener.requests.get")
    def test_get_opinion_text_prefers_html_with_citations(self, mock_get):
        cluster_resp = MagicMock()
        cluster_resp.status_code = 200
        cluster_resp.json.return_value = {
            "id": 12345,
            "sub_opinions": [{"id": 999}],
        }
        cluster_resp.raise_for_status = MagicMock()

        opinion_resp = MagicMock()
        opinion_resp.status_code = 200
        opinion_resp.json.return_value = {
            "plain_text": "plain text fallback",
            "html_with_citations": "<p>html text with <em>citations</em></p>",
        }
        opinion_resp.raise_for_status = MagicMock()
        mock_get.side_effect = [cluster_resp, opinion_resp]

        client = CourtListenerClient(token="tok")
        text = client.get_opinion_text(12345)

        self.assertEqual(text, "html text with citations")
```

- [ ] **Step 2: Run the failing tests**

Run:

```powershell
pytest tests/test_legal_research/test_courtlistener.py::TestCitationLookup -q
```

Expected: FAIL because `lookup_citations` does not exist and `get_opinion_text` does not prefer HTML.

- [ ] **Step 3: Implement `lookup_citations`**

In `icharlotte_core/legal_research/sources/courtlistener.py`, add `import html` near the other imports:

```python
import html
```

Add this method to `CourtListenerClient`:

```python
    def lookup_citations(self, text: str) -> list[dict]:
        """Look up legal citations in a text blob through CourtListener v4."""
        try:
            resp = requests.post(
                f"{BASE_URL}/citation-lookup/",
                headers=self._headers(),
                data={"text": text or ""},
                timeout=REQUEST_TIMEOUT,
            )
            resp.raise_for_status()
            data = resp.json()
            return data if isinstance(data, list) else []
        except Exception:
            logger.warning("CourtListener citation lookup failed", exc_info=True)
            return []
```

- [ ] **Step 4: Prefer `html_with_citations` in `get_opinion_text`**

In `get_opinion_text`, replace the final text extraction block with:

```python
            html_text = opinion_data.get("html_with_citations") or ""
            if html_text:
                text = re.sub(r"<[^>]+>", "", html_text)
                text = html.unescape(text)
            else:
                text = opinion_data.get("plain_text") or ""

            return text if text else None
```

- [ ] **Step 5: Run CourtListener tests**

Run:

```powershell
pytest tests/test_legal_research/test_courtlistener.py -q
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/legal_research/sources/courtlistener.py tests/test_legal_research/test_courtlistener.py
git commit -m "feat(legal-research): add CourtListener citation lookup"
```

---

### Task 5: Implement Citation Verification Mapping And Support Checks

**Files:**
- Create: `icharlotte_core/opposition/citation_verifier.py`
- Test: `tests/test_opposition/test_citation_verifier.py`

- [ ] **Step 1: Write verifier tests**

Create `tests/test_opposition/test_citation_verifier.py`:

```python
from icharlotte_core.opposition.citation_verifier import (
    find_supporting_passage,
    map_lookup_record,
    verify_citations,
)


class FakeCourtListener:
    def lookup_citations(self, text):
        return [
            {
                "citation": "69 Cal.2d 108",
                "normalized_citations": ["69 Cal.2d 108"],
                "status": 200,
                "clusters": [
                    {
                        "id": 10,
                        "case_name": "Rowland v. Christian",
                        "date_filed": "1968-08-08",
                        "absolute_url": "/opinion/10/rowland/",
                        "court": "cal",
                    }
                ],
            },
            {
                "citation": "1 Fake 200",
                "normalized_citations": ["1 Fake 200"],
                "status": 404,
                "error_message": "Citation not found",
                "clusters": [],
            },
        ]

    def get_opinion_text(self, cluster_id):
        assert cluster_id == 10
        return (
            "Everyone is responsible for an injury occasioned to another by "
            "his or her want of ordinary care or skill."
        )

    def search_opinions(self, query, max_results=5):
        return []


def test_map_lookup_record_found():
    verification = map_lookup_record(
        {
            "citation": "69 Cal.2d 108",
            "normalized_citations": ["69 Cal.2d 108"],
            "status": 200,
            "clusters": [{"case_name": "Rowland v. Christian"}],
        }
    )

    assert verification.status == "exists_support_unconfirmed"
    assert verification.case_name == "Rowland v. Christian"
    assert verification.normalized_citation == "69 Cal.2d 108"


def test_find_supporting_passage_uses_overlap():
    passage = find_supporting_passage(
        "ordinary care responsibility injury",
        (
            "Intro sentence. Everyone is responsible for an injury occasioned "
            "to another by his or her want of ordinary care or skill. Closing."
        ),
    )

    assert "ordinary care" in passage


def test_verify_citations_marks_verified_only_when_support_found():
    draft = (
        "California law imposes ordinary care duties. "
        "(Rowland v. Christian (1968) 69 Cal.2d 108.) "
        "Another rule. (Madeup v. Case (2020) 1 Fake 200.)"
    )

    result = verify_citations(
        draft,
        citation_propositions={"69 Cal.2d 108": "ordinary care responsibility injury"},
        courtlistener=FakeCourtListener(),
    )

    assert result[0].status == "verified"
    assert result[0].supporting_passage
    assert result[1].status == "not_found"
```

- [ ] **Step 2: Run the failing tests**

Run:

```powershell
pytest tests/test_opposition/test_citation_verifier.py -q
```

Expected: FAIL because the verifier module does not exist.

- [ ] **Step 3: Implement the verifier**

Create `icharlotte_core/opposition/citation_verifier.py`:

```python
from __future__ import annotations

import re

from .models import CitationVerification


STATUS_MAP = {
    404: "not_found",
    400: "invalid",
    300: "ambiguous",
    429: "throttled",
}


def _cluster_case_name(cluster: dict) -> str:
    return str(cluster.get("case_name") or cluster.get("caseName") or "")


def _cluster_url(cluster: dict) -> str:
    absolute = cluster.get("absolute_url") or cluster.get("absolute_url") or ""
    return f"https://www.courtlistener.com{absolute}" if absolute else ""


def _cluster_id(cluster: dict) -> int | None:
    raw = cluster.get("id") or cluster.get("cluster_id")
    try:
        return int(raw)
    except (TypeError, ValueError):
        return None


def map_lookup_record(record: dict) -> CitationVerification:
    normalized = ""
    normalized_values = record.get("normalized_citations") or []
    if normalized_values:
        normalized = str(normalized_values[0])
    status_code = int(record.get("status") or 0)
    clusters = record.get("clusters") or []
    cluster = clusters[0] if clusters and isinstance(clusters[0], dict) else {}
    if status_code == 200:
        status = "exists_support_unconfirmed"
    else:
        status = STATUS_MAP.get(status_code, "exists_support_unconfirmed")
    return CitationVerification(
        citation_text=str(record.get("citation", "")),
        normalized_citation=normalized,
        status=status,
        case_name=_cluster_case_name(cluster),
        court=str(cluster.get("court") or ""),
        date=str(cluster.get("date_filed") or cluster.get("dateFiled") or ""),
        opinion_url=_cluster_url(cluster),
        warning=str(record.get("error_message") or ""),
    )


def _sentences(text: str) -> list[str]:
    parts = re.split(r"(?<=[.!?])\s+", text or "")
    return [part.strip() for part in parts if part.strip()]


def find_supporting_passage(proposition: str, opinion_text: str) -> str:
    terms = {
        token.lower()
        for token in re.findall(r"[A-Za-z][A-Za-z\-]{3,}", proposition or "")
    }
    if not terms:
        return ""
    best = ""
    best_score = 0
    for sentence in _sentences(opinion_text):
        sentence_terms = {
            token.lower()
            for token in re.findall(r"[A-Za-z][A-Za-z\-]{3,}", sentence)
        }
        score = len(terms & sentence_terms)
        if score > best_score:
            best = sentence
            best_score = score
    minimum = max(2, min(4, len(terms)))
    return best if best_score >= minimum else ""


def verify_citations(
    draft_text: str,
    *,
    citation_propositions: dict[str, str],
    courtlistener,
) -> list[CitationVerification]:
    records = courtlistener.lookup_citations(draft_text)
    verifications: list[CitationVerification] = []
    for record in records:
        verification = map_lookup_record(record)
        if verification.status != "exists_support_unconfirmed":
            verifications.append(verification)
            continue
        clusters = record.get("clusters") or []
        cluster = clusters[0] if clusters and isinstance(clusters[0], dict) else {}
        cluster_id = _cluster_id(cluster)
        proposition = (
            citation_propositions.get(verification.normalized_citation)
            or citation_propositions.get(verification.citation_text)
            or ""
        )
        opinion_text = courtlistener.get_opinion_text(cluster_id) if cluster_id else ""
        passage = find_supporting_passage(proposition, opinion_text or "")
        if passage:
            verification.status = "verified"
            verification.supporting_passage = passage
        verifications.append(verification)
    return verifications
```

- [ ] **Step 4: Run verifier tests**

Run:

```powershell
pytest tests/test_opposition/test_citation_verifier.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/opposition/citation_verifier.py tests/test_opposition/test_citation_verifier.py
git commit -m "feat(opposition): add citation verification states"
```

---

### Task 6: Add Motion Analysis And Drafting Services With Injectable LLMs

**Files:**
- Create: `icharlotte_core/opposition/motion_analyzer.py`
- Create: `icharlotte_core/opposition/drafter.py`
- Test: `tests/test_opposition/test_motion_analyzer.py`

- [ ] **Step 1: Write motion analyzer and drafter tests**

Create `tests/test_opposition/test_motion_analyzer.py`:

```python
import json

from icharlotte_core.opposition.drafter import draft_memorandum
from icharlotte_core.opposition.models import MotionMetadata, SectionPlanItem
from icharlotte_core.opposition.motion_analyzer import (
    analyze_motion,
    generate_outline,
)


def test_analyze_motion_parses_json_metadata():
    def llm(system_prompt, user_prompt):
        return json.dumps(
            {
                "motion_type": "Motion for Summary Judgment",
                "moving_party": "Defendant City",
                "opposing_party": "Plaintiff Doe",
                "relief_requested": "summary judgment",
                "hearing_date": "2026-06-12",
                "opposition_due_date": "",
                "procedural_posture": "pretrial",
                "principal_arguments": ["no duty", "no causation"],
                "opposition_posture": "triable issues of material fact",
            }
        )

    metadata = analyze_motion("NOTICE OF MOTION", llm_callback=llm)

    assert metadata.motion_type == "Motion for Summary Judgment"
    assert metadata.required_missing() == []


def test_generate_outline_creates_selected_three_level_tree():
    def llm(system_prompt, user_prompt):
        return json.dumps(
            {
                "outline": [
                    {
                        "text": "Triable issues defeat summary judgment",
                        "children": [
                            {
                                "text": "The moving evidence is incomplete",
                                "children": [
                                    {"text": "Witness testimony conflicts"}
                                ],
                            }
                        ],
                    }
                ]
            }
        )

    nodes = generate_outline(MotionMetadata(motion_type="MSJ"), "motion", "context", llm_callback=llm)

    assert nodes[0].selected is True
    assert nodes[0].children[0].children[0].selected is True


def test_draft_memorandum_uses_context_without_record_citations():
    def llm(system_prompt, user_prompt):
        assert "CASE CONTEXT" in user_prompt
        return (
            "I. INTRODUCTION\n"
            "The evidence creates triable disputes.\n\n"
            "II. ARGUMENT\n"
            "California law requires summary judgment to be denied when material facts are disputed. "
            "(Aguilar v. Atlantic Richfield Co. (2001) 25 Cal.4th 826.)"
        )

    draft = draft_memorandum(
        metadata=MotionMetadata(motion_type="Motion for Summary Judgment"),
        section_plan=[
            SectionPlanItem(
                id="s1",
                path=["Triable issues defeat summary judgment"],
                level=1,
            )
        ],
        motion_text="Motion text",
        context_text="Context facts from selected documents",
        authority_block="Aguilar v. Atlantic Richfield Co. (2001) 25 Cal.4th 826",
        llm_callback=llm,
    )

    assert "Context facts" not in draft.body_text
    assert "Aguilar" in draft.body_text
```

- [ ] **Step 2: Run the failing tests**

Run:

```powershell
pytest tests/test_opposition/test_motion_analyzer.py -q
```

Expected: FAIL because analyzer and drafter modules do not exist.

- [ ] **Step 3: Implement motion analyzer**

Create `icharlotte_core/opposition/motion_analyzer.py`:

```python
from __future__ import annotations

import json
import uuid

from .models import MotionMetadata, OutlineNode


def _loads_json(text: str) -> dict:
    cleaned = (text or "").strip()
    if cleaned.startswith("```"):
        cleaned = cleaned.split("\n", 1)[1] if "\n" in cleaned else cleaned[3:]
        if cleaned.endswith("```"):
            cleaned = cleaned[:-3]
    try:
        data = json.loads(cleaned.strip())
        return data if isinstance(data, dict) else {}
    except json.JSONDecodeError:
        return {}


def analyze_motion(motion_text: str, *, llm_callback) -> MotionMetadata:
    system_prompt = (
        "You are a California civil litigation attorney. Extract structured "
        "metadata from the motion to oppose. Return only valid JSON."
    )
    user_prompt = (
        "Return JSON with keys: motion_type, moving_party, opposing_party, "
        "relief_requested, hearing_date, opposition_due_date, procedural_posture, "
        "principal_arguments, opposition_posture.\n\n"
        f"MOTION TEXT:\n{motion_text[:200000]}"
    )
    return MotionMetadata.from_dict(_loads_json(llm_callback(system_prompt, user_prompt)))


def _node_from_raw(raw: dict, depth: int = 1) -> OutlineNode:
    children = raw.get("children") if isinstance(raw.get("children"), list) else []
    return OutlineNode(
        id=str(raw.get("id") or f"outline_{uuid.uuid4().hex[:10]}"),
        text=str(raw.get("text", "")).strip(),
        selected=True,
        children=[
            _node_from_raw(child, depth + 1)
            for child in children
            if isinstance(child, dict) and depth < 3
        ],
    )


def generate_outline(
    metadata: MotionMetadata,
    motion_text: str,
    context_text: str,
    *,
    llm_callback,
) -> list[OutlineNode]:
    system_prompt = (
        "You are planning a California civil opposition memorandum. Return a "
        "three-level outline as JSON. Every node must have text and optional children."
    )
    user_prompt = (
        f"MOTION TYPE: {metadata.motion_type}\n"
        f"RELIEF REQUESTED: {metadata.relief_requested}\n"
        f"PRINCIPAL ARGUMENTS: {metadata.principal_arguments}\n\n"
        f"MOTION TEXT:\n{motion_text[:120000]}\n\n"
        f"CASE CONTEXT:\n{context_text[:120000]}\n\n"
        'Return JSON: {"outline": [{"text": "...", "children": []}]}'
    )
    raw = _loads_json(llm_callback(system_prompt, user_prompt))
    outline = raw.get("outline") if isinstance(raw.get("outline"), list) else []
    return [
        _node_from_raw(item)
        for item in outline
        if isinstance(item, dict) and str(item.get("text", "")).strip()
    ]
```

- [ ] **Step 4: Implement drafter**

Create `icharlotte_core/opposition/drafter.py`:

```python
from __future__ import annotations

from .models import DraftDocument, MotionMetadata, SectionPlanItem


def _format_section_plan(section_plan: list[SectionPlanItem]) -> str:
    lines: list[str] = []
    for item in section_plan:
        prefix = "#" * item.level
        lines.append(f"{prefix} {' > '.join(item.path)}")
    return "\n".join(lines)


def draft_memorandum(
    *,
    metadata: MotionMetadata,
    section_plan: list[SectionPlanItem],
    motion_text: str,
    context_text: str,
    authority_block: str,
    llm_callback,
) -> DraftDocument:
    system_prompt = (
        "You are drafting a comprehensive and persuasive California civil "
        "opposition memorandum. Cite only authorities in LEGAL AUTHORITY. "
        "Use case-context facts for grounding but do not cite context documents. "
        "Do not include a citation verification appendix."
    )
    user_prompt = (
        f"MOTION TYPE: {metadata.motion_type}\n"
        f"MOVING PARTY: {metadata.moving_party}\n"
        f"OPPOSING PARTY: {metadata.opposing_party}\n"
        f"RELIEF REQUESTED: {metadata.relief_requested}\n"
        f"OPPOSITION POSTURE: {metadata.opposition_posture}\n\n"
        f"SELECTED OUTLINE:\n{_format_section_plan(section_plan)}\n\n"
        f"LEGAL AUTHORITY:\n{authority_block}\n\n"
        f"MOTION TEXT:\n{motion_text[:200000]}\n\n"
        f"CASE CONTEXT:\n{context_text[:200000]}"
    )
    body = llm_callback(system_prompt, user_prompt) or ""
    title = f"Opposition to {metadata.motion_type}".strip()
    return DraftDocument(title=title, body_text=body.strip())
```

- [ ] **Step 5: Run analyzer tests**

Run:

```powershell
pytest tests/test_opposition/test_motion_analyzer.py -q
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/opposition/motion_analyzer.py icharlotte_core/opposition/drafter.py tests/test_opposition/test_motion_analyzer.py
git commit -m "feat(opposition): add motion analysis and drafting prompts"
```

---

### Task 7: Add Word Assembly And Opposition Validation

**Files:**
- Create: `icharlotte_core/opposition/assembler.py`
- Modify: `icharlotte_core/word_validator.py`
- Test: `tests/test_opposition/test_assembler.py`

- [ ] **Step 1: Write assembly tests**

Create `tests/test_opposition/test_assembler.py`:

```python
from docx import Document

from icharlotte_core.opposition.assembler import assemble_opposition_preview
from icharlotte_core.opposition.models import DraftDocument
from icharlotte_core.word_validator import validate_opposition_docx


def test_assemble_body_only_preview_and_validate(tmp_path):
    output_path = tmp_path / "opposition_preview.docx"
    draft = DraftDocument(
        title="Opposition to Motion for Summary Judgment",
        body_text="I. INTRODUCTION\nThe motion should be denied.\n\nII. ARGUMENT\nTriable issues exist.",
    )

    result = assemble_opposition_preview(
        draft=draft,
        output_path=str(output_path),
        caption_path="",
    )

    assert result == str(output_path)
    validation = validate_opposition_docx(str(output_path))
    assert validation.has_errors is False
    doc = Document(str(output_path))
    assert "Triable issues exist" in "\n".join(p.text for p in doc.paragraphs)


def test_assemble_reuses_caption_docx_when_available(tmp_path):
    caption_path = tmp_path / "Caption Page.docx"
    caption = Document()
    caption.add_paragraph("CAPTION PAGE")
    caption.save(str(caption_path))
    output_path = tmp_path / "preview.docx"

    assemble_opposition_preview(
        draft=DraftDocument(title="Opposition", body_text="Argument body"),
        output_path=str(output_path),
        caption_path=str(caption_path),
    )

    doc = Document(str(output_path))
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Opposition" in text
    assert "Argument body" in text
```

- [ ] **Step 2: Run the failing tests**

Run:

```powershell
pytest tests/test_opposition/test_assembler.py -q
```

Expected: FAIL because `assembler.py` and `validate_opposition_docx` do not exist.

- [ ] **Step 3: Add Word validation wrapper**

At the end of `icharlotte_core/word_validator.py`, add:

```python
def validate_opposition_docx(doc_path: str) -> ValidationResult:
    """Lightweight validation for generated opposition memorandum .docx files."""
    from docx import Document

    result = ValidationResult(
        context=f"Opposition memorandum: {os.path.basename(doc_path)}"
    )
    if not os.path.exists(doc_path):
        result.findings.append(Finding("ERROR", "file", f"File not found: {doc_path}"))
        return result
    try:
        doc = Document(doc_path)
    except Exception as exc:
        result.findings.append(
            Finding("ERROR", "open_docx", f"Could not open generated docx: {exc}")
        )
        return result
    text = "\n".join(p.text for p in doc.paragraphs if p.text and p.text.strip())
    if not text.strip():
        result.findings.append(
            Finding("ERROR", "content", "Generated document contains no paragraph text")
        )
    else:
        result.findings.append(
            Finding("PASS", "content", "Generated document contains paragraph text")
        )
    if "OPPOSITION" in text.upper():
        result.findings.append(
            Finding("PASS", "opposition_marker", "Opposition title detected")
        )
    else:
        result.findings.append(
            Finding("WARN", "opposition_marker", "No opposition title detected")
        )
    return result
```

- [ ] **Step 4: Add assembler**

Create `icharlotte_core/opposition/assembler.py`:

```python
from __future__ import annotations

import os

from docx import Document
from docx.shared import Pt

from .models import DraftDocument


def _add_paragraphs(doc, body_text: str) -> None:
    for raw in (body_text or "").splitlines():
        line = raw.strip()
        if not line:
            doc.add_paragraph("")
            continue
        paragraph = doc.add_paragraph()
        run = paragraph.add_run(line)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12)
        if line.isupper() or line.startswith(("I.", "II.", "III.", "IV.", "V.")):
            run.bold = True


def _replace_caption_marker(doc, title: str) -> bool:
    for paragraph in doc.paragraphs:
        if "CAPTION PAGE" in paragraph.text.upper():
            paragraph.clear()
            run = paragraph.add_run(title)
            run.bold = True
            run.font.name = "Times New Roman"
            run.font.size = Pt(12)
            return True
    return False


def assemble_opposition_preview(
    *,
    draft: DraftDocument,
    output_path: str,
    caption_path: str = "",
) -> str:
    if caption_path and os.path.isfile(caption_path):
        doc = Document(caption_path)
        if not _replace_caption_marker(doc, draft.title or "Opposition"):
            title_para = doc.add_paragraph()
            title_run = title_para.add_run(draft.title or "Opposition")
            title_run.bold = True
    else:
        doc = Document()
        title_para = doc.add_paragraph()
        title_run = title_para.add_run(draft.title or "Opposition")
        title_run.bold = True
        title_run.font.name = "Times New Roman"
        title_run.font.size = Pt(12)
    _add_paragraphs(doc, draft.body_text)
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    doc.save(output_path)
    return output_path
```

- [ ] **Step 5: Run assembly tests**

Run:

```powershell
pytest tests/test_opposition/test_assembler.py -q
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/opposition/assembler.py icharlotte_core/word_validator.py tests/test_opposition/test_assembler.py
git commit -m "feat(opposition): add Word preview assembly"
```

---

### Task 8: Build The Custom Oppose Motion Wizard Page Skeleton

**Files:**
- Create: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
- Test: `tests/test_wizard/test_oppose_motion_page.py`

- [ ] **Step 1: Write UI skeleton tests**

Create `tests/test_wizard/test_oppose_motion_page.py`:

```python
import pytest

pytest.importorskip("pytestqt")

from icharlotte_core.opposition.models import MotionMetadata, OutlineNode
from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
    OpposeMotionSettingsPage,
    OpposeMotionTaskTab,
    SETTINGS_PAGE_CONFIRM,
)


def test_confirmation_blocks_missing_required_fields(qtbot):
    page = OpposeMotionSettingsPage(
        case_root="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(page)
    page.set_metadata(MotionMetadata())

    assert page.can_continue_to_outline() is False

    page.set_metadata(
        MotionMetadata(
            motion_type="Motion for Summary Judgment",
            relief_requested="summary judgment",
            principal_arguments=["no duty"],
        )
    )

    assert page.can_continue_to_outline() is True


def test_outline_items_start_checked(qtbot):
    page = OpposeMotionSettingsPage(
        case_root="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(page)
    page.set_outline(
        [
            OutlineNode(
                id="a",
                text="Main",
                children=[OutlineNode(id="a1", text="Sub")],
            )
        ]
    )

    assert page.outline_tree.topLevelItem(0).checkState(0).value == 2
    assert page.outline_tree.topLevelItem(0).child(0).checkState(0).value == 2


def test_task_tab_starts_on_confirmation_page(qtbot):
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)

    assert tab.settings_page.currentIndex() == SETTINGS_PAGE_CONFIRM
```

- [ ] **Step 2: Run the failing UI tests**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_page.py -q
```

Expected: FAIL because the page module does not exist.

- [ ] **Step 3: Create the settings page and task tab skeleton**

Create `icharlotte_core/ui/wizard/pages/oppose_motion_page.py` with these class and signal names:

```python
from __future__ import annotations

import os

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QStackedWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from icharlotte_core.opposition.models import (
    DraftDocument,
    MotionMetadata,
    OutlineNode,
)
from icharlotte_core.ui.wizard.pages.status_page import StatusPage


SETTINGS_PAGE_CONFIRM = 0
SETTINGS_PAGE_OUTLINE = 1
TASK_PAGE_SETTINGS = 0
TASK_PAGE_STATUS = 1
TASK_PAGE_OUTPUT = 2


class OpposeMotionSettingsPage(QStackedWidget):
    run_requested = Signal(dict)

    def __init__(
        self,
        case_root: str,
        file_number: str,
        motion_file: str,
        context_files: list[str],
        parent: QWidget | None = None,
    ):
        super().__init__(parent)
        self.case_root = case_root
        self.file_number = file_number
        self.motion_file = motion_file
        self.context_files = list(context_files)
        self.metadata = MotionMetadata()
        self.outline: list[OutlineNode] = []
        self._build_confirm_page()
        self._build_outline_page()

    def _build_confirm_page(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.addWidget(QLabel(f"Motion: {os.path.basename(self.motion_file)}"))
        self.motion_type_edit = QLineEdit()
        self.motion_type_edit.setPlaceholderText("Motion type")
        layout.addWidget(self.motion_type_edit)
        self.relief_edit = QLineEdit()
        self.relief_edit.setPlaceholderText("Relief requested")
        layout.addWidget(self.relief_edit)
        self.arguments_edit = QPlainTextEdit()
        self.arguments_edit.setPlaceholderText("Principal arguments, one per line")
        layout.addWidget(self.arguments_edit, 1)
        row = QHBoxLayout()
        row.addStretch()
        self.continue_btn = QPushButton("Generate Outline")
        self.continue_btn.clicked.connect(self._on_continue_to_outline)
        row.addWidget(self.continue_btn)
        layout.addLayout(row)
        self.addWidget(page)

    def _build_outline_page(self) -> None:
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.addWidget(QLabel("Opposition Outline"))
        self.outline_tree = QTreeWidget()
        self.outline_tree.setHeaderLabels(["Include / Heading"])
        layout.addWidget(self.outline_tree, 1)
        row = QHBoxLayout()
        self.add_heading_btn = QPushButton("Add Heading")
        row.addWidget(self.add_heading_btn)
        row.addStretch()
        self.generate_btn = QPushButton("Generate Draft")
        self.generate_btn.clicked.connect(self._emit_run_requested)
        row.addWidget(self.generate_btn)
        layout.addLayout(row)
        self.addWidget(page)

    def set_metadata(self, metadata: MotionMetadata) -> None:
        self.metadata = metadata
        self.motion_type_edit.setText(metadata.motion_type)
        self.relief_edit.setText(metadata.relief_requested)
        self.arguments_edit.setPlainText("\n".join(metadata.principal_arguments))

    def current_metadata(self) -> MotionMetadata:
        metadata = MotionMetadata.from_dict(self.metadata.to_dict())
        metadata.motion_type = self.motion_type_edit.text().strip()
        metadata.relief_requested = self.relief_edit.text().strip()
        metadata.principal_arguments = [
            line.strip()
            for line in self.arguments_edit.toPlainText().splitlines()
            if line.strip()
        ]
        return metadata

    def can_continue_to_outline(self) -> bool:
        return self.current_metadata().required_missing() == []

    def _on_continue_to_outline(self) -> None:
        if not self.can_continue_to_outline():
            QMessageBox.warning(
                self,
                "Missing required fields",
                "Motion type, relief requested, and principal arguments are required.",
            )
            return
        self.setCurrentIndex(SETTINGS_PAGE_OUTLINE)

    def set_outline(self, outline: list[OutlineNode]) -> None:
        self.outline = list(outline)
        self.outline_tree.clear()
        for node in self.outline:
            self.outline_tree.addTopLevelItem(self._item_from_node(node))
        self.outline_tree.expandAll()

    def _item_from_node(self, node: OutlineNode) -> QTreeWidgetItem:
        item = QTreeWidgetItem([node.text])
        item.setData(0, Qt.ItemDataRole.UserRole, node.id)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsEditable | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(
            0,
            Qt.CheckState.Checked if node.selected else Qt.CheckState.Unchecked,
        )
        for child in node.children:
            item.addChild(self._item_from_node(child))
        return item

    def to_dict(self) -> dict:
        return {
            "motion_file": self.motion_file,
            "context_files": list(self.context_files),
            "metadata": self.current_metadata().to_dict(),
            "outline": [node.to_dict() for node in self.outline],
        }

    def from_dict(self, data: dict) -> None:
        self.motion_file = data.get("motion_file", self.motion_file)
        self.context_files = list(data.get("context_files", self.context_files))
        self.set_metadata(MotionMetadata.from_dict(data.get("metadata")))
        self.set_outline([
            OutlineNode.from_dict(item)
            for item in data.get("outline", [])
            if isinstance(item, dict)
        ])

    def _emit_run_requested(self) -> None:
        self.run_requested.emit(self.to_dict())


class OpposeMotionOutputPage(QWidget):
    rerun_requested = Signal()
    edit_settings_requested = Signal()

    def __init__(self, parent: QWidget | None = None):
        super().__init__(parent)
        self.draft = DraftDocument()
        layout = QHBoxLayout(self)
        self.editor = QPlainTextEdit()
        layout.addWidget(self.editor, 2)
        self.source_drawer = QPlainTextEdit()
        self.source_drawer.setReadOnly(True)
        layout.addWidget(self.source_drawer, 1)

    def show_result(self, draft: DraftDocument) -> None:
        self.draft = draft
        self.editor.setPlainText(draft.body_text)
        if draft.citations:
            first = draft.citations[0]
            self.source_drawer.setPlainText(
                f"{first.citation_text}\n{first.status}\n\n{first.supporting_passage}"
            )


class OpposeMotionTaskTab(QStackedWidget):
    task_completed = Signal(dict)

    def __init__(self, spec, case_path: str, file_number: str, motion_file: str, context_files: list[str], parent=None):
        super().__init__(parent)
        self._spec = spec
        self._case_path = case_path
        self._file_number = file_number
        self._files = [motion_file] + list(context_files)
        self._worker = None
        self.settings_page = OpposeMotionSettingsPage(case_path, file_number, motion_file, context_files)
        self.status_page = StatusPage()
        self.output_page = OpposeMotionOutputPage()
        self.addWidget(self.settings_page)
        self.addWidget(self.status_page)
        self.addWidget(self.output_page)
        self.settings_page.run_requested.connect(self._on_run)

    @property
    def spec(self):
        return self._spec

    @property
    def files(self) -> list[str]:
        return list(self._files)

    def _on_run(self, settings: dict) -> None:
        self.status_page.reset()
        self.status_page.on_status("Drafting opposition memorandum...")
        self.setCurrentIndex(TASK_PAGE_STATUS)


def build_oppose_motion_tab(spec, case_path: str, file_number: str, parent: QWidget | None):
    motion_file, _ = QFileDialog.getOpenFileName(
        parent,
        "Select motion to oppose",
        case_path,
        "Motion files (*.pdf *.docx)",
    )
    if not motion_file:
        return None
    context_files, _ = QFileDialog.getOpenFileNames(
        parent,
        "Select context document(s)",
        os.path.dirname(motion_file) or case_path,
        "Context files (*.pdf *.docx *.txt *.msg);;All files (*.*)",
    )
    return OpposeMotionTaskTab(
        spec=spec,
        case_path=case_path,
        file_number=file_number,
        motion_file=motion_file,
        context_files=list(context_files or []),
        parent=parent,
    )
```

- [ ] **Step 4: Run UI skeleton tests**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_page.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_wizard/test_oppose_motion_page.py
git commit -m "feat(wizard): add oppose motion page skeleton"
```

---

### Task 9: Wire The Worker Pipeline Into The Custom Task Tab

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
- Test: `tests/test_wizard/test_oppose_motion_page.py`

- [ ] **Step 1: Add worker success-path test**

Append to `tests/test_wizard/test_oppose_motion_page.py`:

```python
from icharlotte_core.opposition.models import CitationVerification, DraftDocument
from icharlotte_core.ui.wizard.pages.oppose_motion_page import TASK_PAGE_OUTPUT


def test_task_tab_loads_draft_result(qtbot):
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)
    draft = DraftDocument(
        title="Opposition",
        body_text="Argument text",
        citations=[
            CitationVerification(
                citation_text="69 Cal.2d 108",
                status="verified",
                supporting_passage="ordinary care",
            )
        ],
    )

    tab._on_worker_finished(True, draft)

    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert "Argument text" in tab.output_page.editor.toPlainText()
    assert "ordinary care" in tab.output_page.source_drawer.toPlainText()
```

- [ ] **Step 2: Run the failing test**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_page.py::test_task_tab_loads_draft_result -q
```

Expected: FAIL because `_on_worker_finished` does not exist.

- [ ] **Step 3: Add `OpposeMotionWorker`**

In `oppose_motion_page.py`, add imports:

```python
from icharlotte_core.discovery.assembler import DiscoveryAssembler
from icharlotte_core.legal_research.sources.courtlistener import CourtListenerClient
from icharlotte_core.opposition.assembler import assemble_opposition_preview
from icharlotte_core.opposition.citation_verifier import verify_citations
from icharlotte_core.opposition.drafter import draft_memorandum
from icharlotte_core.opposition.extraction import extract_context_bundle, extract_document_text
from icharlotte_core.opposition.motion_analyzer import analyze_motion, generate_outline
from icharlotte_core.opposition.outline import selected_section_plan
from icharlotte_core.word_validator import validate_opposition_docx
```

Add this worker class:

```python
class OpposeMotionWorker(QThread):
    progress = Signal(str)
    finished_result = Signal(bool, object)

    def __init__(self, case_path: str, file_number: str, settings: dict, parent=None):
        super().__init__(parent)
        self.case_path = case_path
        self.file_number = file_number
        self.settings = dict(settings or {})

    def run(self) -> None:
        try:
            from icharlotte_core.config import COURTLISTENER_API_TOKEN
            from icharlotte_core.llm_config import call_llm

            self.progress.emit("Extracting motion text...")
            motion_file = self.settings.get("motion_file", "")
            motion_result = extract_document_text(motion_file)
            if not motion_result.success:
                self.finished_result.emit(False, motion_result.error or "Could not read motion.")
                return

            self.progress.emit("Extracting context documents...")
            context_text, warnings = extract_context_bundle(self.settings.get("context_files", []))

            self.progress.emit("Preparing metadata and outline...")
            metadata = MotionMetadata.from_dict(self.settings.get("metadata"))
            outline = [
                OutlineNode.from_dict(item)
                for item in self.settings.get("outline", [])
                if isinstance(item, dict)
            ]
            plan = selected_section_plan(outline)

            self.progress.emit("Researching and drafting memorandum...")
            def llm(system_prompt, user_prompt):
                return call_llm(user_prompt, system_prompt, task_type="general", agent_id="agent_chat")

            draft = draft_memorandum(
                metadata=metadata,
                section_plan=plan,
                motion_text=motion_result.text,
                context_text=context_text,
                authority_block="",
                llm_callback=llm,
            )

            if COURTLISTENER_API_TOKEN:
                self.progress.emit("Verifying citations...")
                client = CourtListenerClient(COURTLISTENER_API_TOKEN)
                draft.citations = verify_citations(
                    draft.body_text,
                    citation_propositions={},
                    courtlistener=client,
                )
            else:
                self.progress.emit("CourtListener token missing. Citations remain unverified.")

            preview_dir = os.path.join(
                self.case_path,
                "NOTES",
                "AI OUTPUT",
                ".icharlotte",
                "wizard_previews",
                "oppose_motion",
            )
            preview_path = os.path.join(preview_dir, "Opposition Preview.docx")
            caption_path = DiscoveryAssembler.find_caption_page(self.case_path) or ""
            assemble_opposition_preview(draft=draft, output_path=preview_path, caption_path=caption_path)
            validation = validate_opposition_docx(preview_path)
            if validation.has_errors:
                self.finished_result.emit(False, "Word validation failed for opposition preview.")
                return
            draft.preview_path = preview_path
            self.finished_result.emit(True, draft)
        except Exception as exc:
            self.finished_result.emit(False, str(exc))
```

- [ ] **Step 4: Wire task tab worker completion**

In `OpposeMotionTaskTab`, replace `_on_run` with:

```python
    def _on_run(self, settings: dict) -> None:
        self.status_page.reset()
        self.status_page.on_status("Drafting opposition memorandum...")
        self.status_page.progress_bar.setRange(0, 0)
        self.setCurrentIndex(TASK_PAGE_STATUS)
        worker = OpposeMotionWorker(
            case_path=self._case_path,
            file_number=self._file_number,
            settings=settings,
            parent=self,
        )
        worker.progress.connect(self.status_page.on_status)
        worker.finished_result.connect(self._on_worker_finished)
        self._worker = worker
        worker.start()

    def _on_worker_finished(self, success: bool, payload: object) -> None:
        from datetime import datetime

        self._worker = None
        if not success:
            self.status_page.on_status(f"FAILED: {payload}")
            return
        draft = payload if isinstance(payload, DraftDocument) else DraftDocument()
        self.output_page.show_result(draft)
        self.setCurrentIndex(TASK_PAGE_OUTPUT)
        self.task_completed.emit(
            {
                "task_id": self._spec.task_id,
                "title": self._spec.title,
                "files": list(self._files),
                "settings": self.settings_page.to_dict(),
                "output_path": draft.preview_path,
                "completed_at": datetime.now().isoformat(timespec="seconds"),
            }
        )
```

- [ ] **Step 5: Run UI tests**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_page.py -q
```

Expected: PASS.

- [ ] **Step 6: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_wizard/test_oppose_motion_page.py
git commit -m "feat(wizard): run oppose motion draft worker"
```

---

### Task 10: Add Source Drawer Citation Click And Save As Behavior

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
- Test: `tests/test_wizard/test_oppose_motion_page.py`

- [ ] **Step 1: Add source drawer and Save As tests**

Append to `tests/test_wizard/test_oppose_motion_page.py`:

```python
from unittest.mock import patch


def test_output_page_show_citation_updates_drawer(qtbot):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    draft = DraftDocument(
        title="Opposition",
        body_text="Text",
        citations=[
            CitationVerification(
                citation_text="69 Cal.2d 108",
                normalized_citation="69 Cal.2d 108",
                status="verified",
                case_name="Rowland v. Christian",
                supporting_passage="ordinary care language",
            )
        ],
    )
    page.show_result(draft)
    page.show_citation(0)

    assert "Rowland v. Christian" in page.source_drawer.toPlainText()
    assert "ordinary care language" in page.source_drawer.toPlainText()


def test_save_as_uses_dialog_and_does_not_save_when_cancelled(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    source = tmp_path / "preview.docx"
    source.write_bytes(b"docx-bytes")
    page.show_result(DraftDocument(title="Opposition", body_text="Text", preview_path=str(source)))

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=("", ""),
    ) as dialog:
        page.save_as()

    assert dialog.called
```

- [ ] **Step 2: Run the failing tests**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_page.py::test_output_page_show_citation_updates_drawer tests/test_wizard/test_oppose_motion_page.py::test_save_as_uses_dialog_and_does_not_save_when_cancelled -q
```

Expected: FAIL because `show_citation` and `save_as` do not exist.

- [ ] **Step 3: Add output page buttons and methods**

In `OpposeMotionOutputPage.__init__`, replace the simple layout with:

```python
        outer = QVBoxLayout(self)
        body = QHBoxLayout()
        self.editor = QPlainTextEdit()
        body.addWidget(self.editor, 2)
        self.source_drawer = QPlainTextEdit()
        self.source_drawer.setReadOnly(True)
        body.addWidget(self.source_drawer, 1)
        outer.addLayout(body, 1)

        row = QHBoxLayout()
        row.addStretch()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_as)
        row.addWidget(self.save_btn)
        outer.addLayout(row)
```

Add methods:

```python
    def show_citation(self, index: int) -> None:
        if index < 0 or index >= len(self.draft.citations):
            return
        citation = self.draft.citations[index]
        self.source_drawer.setPlainText(
            "\n".join(
                [
                    citation.citation_text,
                    f"Normalized: {citation.normalized_citation}",
                    f"Status: {citation.status}",
                    f"Case: {citation.case_name}",
                    f"Court: {citation.court}",
                    f"Date: {citation.date}",
                    f"Opinion: {citation.opinion_url}",
                    "",
                    "Supporting passage:",
                    citation.supporting_passage or "(support not confirmed)",
                    "",
                    citation.warning,
                ]
            ).strip()
        )

    def save_as(self) -> None:
        import shutil

        if not self.draft.preview_path:
            QMessageBox.warning(self, "No preview", "No generated opposition preview is available.")
            return
        target, _ = QFileDialog.getSaveFileName(
            self,
            "Save Opposition Memorandum",
            os.path.join(os.path.dirname(self.draft.preview_path), self.draft.title + ".docx"),
            "Word Documents (*.docx);;All files (*.*)",
        )
        if not target:
            return
        if not target.lower().endswith(".docx"):
            target += ".docx"
        shutil.copyfile(self.draft.preview_path, target)
        QMessageBox.information(self, "Saved", f"Saved:\n{target}")
```

- [ ] **Step 4: Run UI tests**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_page.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_wizard/test_oppose_motion_page.py
git commit -m "feat(wizard): add opposition source drawer and save as"
```

---

### Task 11: Restore And Reopen In-Process Task Tabs Correctly

**Files:**
- Modify: `iCharlotte.py`
- Test: `tests/test_wizard/test_in_process_task_tab.py` or new `tests/test_wizard/test_oppose_motion_restore.py`

- [ ] **Step 1: Add a pure helper for in-process task restoration**

Because `iCharlotte.py` is large and difficult to instantiate in unit tests, first add a small helper in `icharlotte_core/ui/wizard/task_routing.py`:

```python
def is_in_process_task(task_id: str) -> bool:
    return get_in_process_task_builder_name(task_id) is not None
```

Add to `tests/test_wizard/test_task_routing.py`:

```python
from icharlotte_core.ui.wizard.task_routing import is_in_process_task


def test_oppose_motion_is_in_process_task():
    assert is_in_process_task("oppose_motion") is True
    assert is_in_process_task("summarize_documents") is False
```

- [ ] **Step 2: Run routing tests**

Run:

```powershell
pytest tests/test_wizard/test_task_routing.py tests/test_wizard/test_oppose_motion_registry.py -q
```

Expected: PASS after adding the helper.

- [ ] **Step 3: Modify `_restore_task_tabs_for_case` and `_on_reopen_recent_task`**

In `iCharlotte.py`, update both restore paths so if `get_in_process_task_builder_name(task_id)` returns `"build_oppose_motion_tab"`, the code reopens the saved task by constructing `OpposeMotionTaskTab` directly and applying the persisted settings. The launch builder remains only for starting a new task from the Wizard card, where file dialogs are expected. Use:

```python
from icharlotte_core.ui.wizard.task_routing import get_in_process_task_builder_name
from icharlotte_core.ui.wizard import in_process_task_tab

builder_name = get_in_process_task_builder_name(task_id)
if builder_name == "build_oppose_motion_tab":
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
        OpposeMotionTaskTab,
    )
    settings = entry.get("settings") or {}
    task_tab = OpposeMotionTaskTab(
        spec=spec,
        case_path=self.case_path,
        file_number=self.file_number,
        motion_file=settings.get("motion_file", ""),
        context_files=settings.get("context_files", []),
        parent=self,
    )
    task_tab.settings_page.from_dict(settings)
else:
    task_tab = TaskTab(
        spec=spec,
        files=files_abs,
        case_path=self.case_path,
        file_number=self.file_number,
        parent=self,
    )
```

Apply this pattern in both `_restore_task_tabs_for_case` and `_on_reopen_recent_task` for `oppose_motion`. Leave existing generic `TaskTab` behavior unchanged for standard subprocess tasks.

- [ ] **Step 4: Run focused wizard routing tests**

Run:

```powershell
pytest tests/test_wizard/test_task_routing.py tests/test_wizard/test_oppose_motion_registry.py tests/test_wizard/test_oppose_motion_page.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add iCharlotte.py icharlotte_core/ui/wizard/task_routing.py tests/test_wizard/test_task_routing.py
git commit -m "feat(wizard): restore oppose motion task tabs"
```

---

### Task 12: Add Real Motion And Context Picker Validation

**Files:**
- Modify: `icharlotte_core/ui/wizard/pages/oppose_motion_page.py`
- Test: `tests/test_wizard/test_oppose_motion_page.py`

- [ ] **Step 1: Add builder picker tests**

Append to `tests/test_wizard/test_oppose_motion_page.py`:

```python
def test_builder_rejects_cancelled_motion_picker(qtbot):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=("", ""),
    ):
        tab = build_oppose_motion_tab(spec, "/tmp/case", "0000.000", None)
    assert tab is None


def test_builder_rejects_unsupported_motion_file(qtbot):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=("/tmp/motion.txt", ""),
    ):
        with patch("icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.warning") as warning:
            tab = build_oppose_motion_tab(spec, "/tmp/case", "0000.000", None)
    assert tab is None
    assert warning.called
```

- [ ] **Step 2: Run failing tests**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_page.py::test_builder_rejects_cancelled_motion_picker tests/test_wizard/test_oppose_motion_page.py::test_builder_rejects_unsupported_motion_file -q
```

Expected: FAIL if builder does not validate extensions.

- [ ] **Step 3: Add validation to builder**

In `oppose_motion_page.py`, import:

```python
from icharlotte_core.opposition.extraction import (
    is_supported_context_file,
    is_supported_motion_file,
)
```

In `build_oppose_motion_tab`, after motion picker:

```python
    if not is_supported_motion_file(motion_file):
        QMessageBox.warning(
            parent,
            "Unsupported motion file",
            "Select a PDF or DOCX motion.",
        )
        return None
```

After context picker:

```python
    context_files = [
        path for path in (context_files or [])
        if is_supported_context_file(path)
    ]
```

- [ ] **Step 4: Run page tests**

Run:

```powershell
pytest tests/test_wizard/test_oppose_motion_page.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/ui/wizard/pages/oppose_motion_page.py tests/test_wizard/test_oppose_motion_page.py
git commit -m "feat(wizard): validate oppose motion file pickers"
```

---

### Task 13: Add Replacement Candidate Suggestions

**Files:**
- Modify: `icharlotte_core/opposition/citation_verifier.py`
- Test: `tests/test_opposition/test_citation_verifier.py`

- [ ] **Step 1: Add replacement suggestion test**

Append to `tests/test_opposition/test_citation_verifier.py`:

```python
def test_replacement_candidates_are_added_for_unconfirmed_support():
    class ReplacementClient(FakeCourtListener):
        def lookup_citations(self, text):
            return [
                {
                    "citation": "69 Cal.2d 108",
                    "normalized_citations": ["69 Cal.2d 108"],
                    "status": 200,
                    "clusters": [{"id": 10, "case_name": "Rowland v. Christian"}],
                }
            ]

        def get_opinion_text(self, cluster_id):
            return "No matching support terms here."

        def search_opinions(self, query, max_results=5):
            from icharlotte_core.legal_research.models import CaseResult
            return [
                CaseResult(
                    name="Aguilar v. Atlantic Richfield Co.",
                    citation="25 Cal.4th 826",
                    date="2001-07-12",
                    court="cal",
                    snippet="summary judgment burden and triable issue rule",
                    url="https://www.courtlistener.com/opinion/aguilar/",
                    cluster_id=99,
                )
            ]

    result = verify_citations(
        "Text (Rowland v. Christian (1968) 69 Cal.2d 108.)",
        citation_propositions={"69 Cal.2d 108": "summary judgment triable issue burden"},
        courtlistener=ReplacementClient(),
    )

    assert result[0].status == "exists_support_unconfirmed"
    assert result[0].replacement_candidates[0]["case_name"] == "Aguilar v. Atlantic Richfield Co."
```

- [ ] **Step 2: Run failing test**

Run:

```powershell
pytest tests/test_opposition/test_citation_verifier.py::test_replacement_candidates_are_added_for_unconfirmed_support -q
```

Expected: FAIL because replacements are not added.

- [ ] **Step 3: Add replacement helper**

In `citation_verifier.py`, add:

```python
def replacement_candidates(proposition: str, courtlistener) -> list[dict]:
    if not proposition.strip():
        return []
    results = courtlistener.search_opinions(proposition, max_results=5)
    candidates: list[dict] = []
    for case in results:
        candidates.append(
            {
                "case_name": case.name,
                "citation": case.formatted_citation,
                "court": case.court,
                "date": case.date,
                "opinion_url": case.url,
                "reason": case.snippet,
            }
        )
    return candidates
```

In `verify_citations`, after support passage check:

```python
        if passage:
            verification.status = "verified"
            verification.supporting_passage = passage
        else:
            verification.replacement_candidates = replacement_candidates(
                proposition,
                courtlistener,
            )
```

- [ ] **Step 4: Run verifier tests**

Run:

```powershell
pytest tests/test_opposition/test_citation_verifier.py -q
```

Expected: PASS.

- [ ] **Step 5: Commit**

```powershell
git add icharlotte_core/opposition/citation_verifier.py tests/test_opposition/test_citation_verifier.py
git commit -m "feat(opposition): suggest citation replacements"
```

---

### Task 14: Run Focused Verification Suite And Manual Smoke Check

**Files:**
- No planned source edits unless verification fails.

- [ ] **Step 1: Compile changed Python modules**

Run:

```powershell
python -m py_compile `
  icharlotte_core/opposition/models.py `
  icharlotte_core/opposition/extraction.py `
  icharlotte_core/opposition/outline.py `
  icharlotte_core/opposition/motion_analyzer.py `
  icharlotte_core/opposition/drafter.py `
  icharlotte_core/opposition/citation_verifier.py `
  icharlotte_core/opposition/assembler.py `
  icharlotte_core/ui/wizard/pages/oppose_motion_page.py `
  icharlotte_core/ui/wizard/registry.py `
  icharlotte_core/ui/wizard/task_routing.py `
  icharlotte_core/legal_research/sources/courtlistener.py `
  icharlotte_core/word_validator.py
```

Expected: no output and exit code 0.

- [ ] **Step 2: Run opposition and wizard tests**

Run:

```powershell
pytest tests/test_opposition tests/test_wizard/test_oppose_motion_registry.py tests/test_wizard/test_oppose_motion_page.py tests/test_wizard/test_task_routing.py -q
```

Expected: PASS.

- [ ] **Step 3: Run legal research client tests**

Run:

```powershell
pytest tests/test_legal_research/test_courtlistener.py -q
```

Expected: PASS.

- [ ] **Step 4: Manual app smoke check**

Run the app:

```powershell
python iCharlotte.py
```

Manual steps:

1. Open a test case in Wizard Mode.
2. Click **Oppose a Motion**.
3. Select a PDF or DOCX motion.
4. Select context files or continue with none and verify the warning behavior.
5. Confirm metadata fields.
6. Confirm outline items are selected by default.
7. Click **Generate Draft**.
8. Verify the output screen shows draft text and the right-side source drawer.
9. Click **Save** and verify a Save As dialog appears.
10. Cancel Save As and confirm no final user-selected file is created.
11. Do not close any Microsoft Word windows during this check.

- [ ] **Step 5: Commit verification fixes if any**

If any verification step required a source edit:

```powershell
git add <changed files>
git commit -m "fix(opposition): resolve verification issues"
```

If no source edits were required, do not create an empty commit.

---

## Plan Self-Review Checklist

- Spec coverage:
  - Wizard card and routing: Task 1.
  - Motion/context selection and type validation: Tasks 8 and 12.
  - Motion confirmation and outline review: Tasks 6 and 8.
  - Three-level outline selected by default: Tasks 3 and 8.
  - Drafting from selected outline and factual context without context citations: Task 6.
  - CourtListener citation lookup and verification states: Tasks 4 and 5.
  - Replacement candidates without automatic replacement: Task 13.
  - Right-side source drawer: Task 10.
  - Caption/template preview and Word validation: Task 7.
  - Save As only: Task 10.
  - Wizard-state persistence and reopen: Task 11.
  - No separate verification report or appendix: enforced by Tasks 7, 10, and 11.
- Placeholder scan:
  - This plan contains no unresolved placeholders or unspecified “add tests” steps.
- Type consistency:
  - `MotionMetadata`, `OutlineNode`, `SectionPlanItem`, `CitationVerification`, and `DraftDocument` are introduced in Task 2 and reused by name in later tasks.
  - The task id is consistently `oppose_motion`.
  - The builder name is consistently `build_oppose_motion_tab`.
