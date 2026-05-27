"""Pure outline tree helpers for opposition drafting."""

from __future__ import annotations

from icharlotte_core.opposition.models import OutlineNode, SectionPlanItem


MAX_OUTLINE_DEPTH = 3


def normalize_outline(nodes: list[OutlineNode]) -> list[OutlineNode]:
    """Return a cleaned copy of an outline tree."""
    reserved_existing_ids = _existing_nonblank_ids(nodes)
    used_ids: set[str] = set()

    def normalize_node(
        node: OutlineNode,
        depth: int,
        path: tuple[int, ...],
    ) -> list[OutlineNode]:
        text = node.text.strip()
        if not text:
            promoted_children: list[OutlineNode] = []
            for child_index, child in enumerate(node.children):
                promoted_children.extend(
                    normalize_node(child, depth, (*path, child_index + 1))
                )
            return promoted_children

        existing_id = node.id.strip() if isinstance(node.id, str) else ""
        if existing_id and existing_id not in used_ids:
            node_id = existing_id
        else:
            node_id = _generated_id(path, used_ids | reserved_existing_ids)
        used_ids.add(node_id)

        children: list[OutlineNode] = []
        if depth < MAX_OUTLINE_DEPTH:
            for child_index, child in enumerate(node.children):
                children.extend(
                    normalize_node(child, depth + 1, (*path, child_index + 1))
                )

        return [
            OutlineNode(
                id=node_id,
                text=text,
                selected=node.selected,
                children=children,
            )
        ]

    normalized: list[OutlineNode] = []
    for index, node in enumerate(nodes):
        normalized.extend(normalize_node(node, 1, (index + 1,)))
    return normalized


def flatten_outline(nodes: list[OutlineNode]) -> list[OutlineNode]:
    """Return outline nodes in preorder."""
    flattened: list[OutlineNode] = []

    def visit(node: OutlineNode) -> None:
        flattened.append(node)
        for child in node.children:
            visit(child)

    for node in nodes:
        visit(node)
    return flattened


def selected_section_plan(nodes: list[OutlineNode]) -> list[SectionPlanItem]:
    """Convert selected outline nodes into section plan items."""
    plan: list[SectionPlanItem] = []

    def visit(node: OutlineNode, ancestors: list[str]) -> None:
        text = node.text.strip()
        if not text:
            return

        path = [*ancestors, text]
        if node.selected:
            plan.append(SectionPlanItem(id=node.id, path=path, text=text))

        for child in node.children:
            visit(child, path)

    for node in nodes:
        visit(node, [])
    return plan


def _existing_nonblank_ids(nodes: list[OutlineNode]) -> set[str]:
    ids: set[str] = set()
    for node in nodes:
        if isinstance(node.id, str) and node.id.strip():
            ids.add(node.id.strip())
        ids.update(_existing_nonblank_ids(node.children))
    return ids


def _generated_id(path: tuple[int, ...], used_ids: set[str]) -> str:
    base_id = "outline-" + "-".join(str(part) for part in path)
    candidate = base_id
    suffix = 2
    while candidate in used_ids:
        candidate = f"{base_id}-{suffix}"
        suffix += 1
    return candidate
