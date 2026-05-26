from icharlotte_core.opposition.models import OutlineNode
from icharlotte_core.opposition.outline import (
    MAX_OUTLINE_DEPTH,
    flatten_outline,
    normalize_outline,
    selected_section_plan,
)


def test_normalize_outline_keeps_three_levels_max():
    nodes = [
        OutlineNode(
            id="main",
            text="Main",
            children=[
                OutlineNode(
                    id="sub",
                    text="Sub",
                    children=[
                        OutlineNode(
                            id="detail",
                            text="Detail",
                            children=[
                                OutlineNode(id="too-deep", text="Too Deep"),
                            ],
                        ),
                    ],
                ),
            ],
        )
    ]

    normalized = normalize_outline(nodes)

    assert MAX_OUTLINE_DEPTH == 3
    assert normalized[0].children[0].children[0].text == "Detail"
    assert normalized[0].children[0].children[0].children == []
    assert len(flatten_outline(normalized)) == 3


def test_normalize_outline_drops_blank_text_nodes():
    nodes = [
        OutlineNode(id="blank", text="   "),
        OutlineNode(id="kept", text="  Kept  "),
    ]

    normalized = normalize_outline(nodes)

    assert [node.text for node in normalized] == ["Kept"]
    assert [node.id for node in normalized] == ["kept"]


def test_normalize_outline_assigns_missing_ids():
    nodes = [
        OutlineNode(text="Main", children=[OutlineNode(id="", text="Sub")]),
    ]

    normalized = normalize_outline(nodes)

    assert normalized[0].id
    assert normalized[0].children[0].id
    assert normalized[0].id != normalized[0].children[0].id


def test_normalize_outline_assigns_unique_ids_for_duplicate_existing_ids():
    nodes = [
        OutlineNode(id="same", text="First"),
        OutlineNode(id="same", text="Second"),
        OutlineNode(id="same", text="Third"),
    ]

    normalized = normalize_outline(nodes)

    ids = [node.id for node in normalized]
    assert ids[0] == "same"
    assert len(ids) == len(set(ids))
    assert ids[1:] == ["outline-2", "outline-3"]


def test_normalize_outline_strips_whitespace_from_existing_ids():
    nodes = [
        OutlineNode(id="  existing-id  ", text="Main"),
    ]

    normalized = normalize_outline(nodes)

    assert normalized[0].id == "existing-id"


def test_normalize_outline_promotes_nonblank_children_from_blank_parent():
    nodes = [
        OutlineNode(
            id="blank-parent",
            text=" ",
            children=[
                OutlineNode(id="child", text=" Child ", selected=False),
                OutlineNode(text="", children=[OutlineNode(id="grandchild", text="Grandchild")]),
            ],
        ),
        OutlineNode(id="main", text="Main"),
    ]

    normalized = normalize_outline(nodes)

    assert [node.text for node in normalized] == ["Child", "Grandchild", "Main"]
    assert [node.id for node in normalized] == ["child", "grandchild", "main"]
    assert normalized[0].selected is False


def test_selected_section_plan_includes_all_selected_nodes_by_default():
    nodes = normalize_outline(
        [
            OutlineNode(
                id="main",
                text="Main",
                children=[OutlineNode(id="sub", text="Sub")],
            )
        ]
    )

    plan = selected_section_plan(nodes)

    assert [item.id for item in plan] == ["main", "sub"]
    assert [item.path for item in plan] == [["Main"], ["Main", "Sub"]]
    assert [item.text for item in plan] == ["Main", "Sub"]


def test_selected_child_under_unselected_parent_keeps_parent_in_path():
    nodes = normalize_outline(
        [
            OutlineNode(
                id="main",
                text="Main",
                selected=False,
                children=[OutlineNode(id="sub", text="Sub", selected=True)],
            )
        ]
    )

    plan = selected_section_plan(nodes)

    assert [item.id for item in plan] == ["sub"]
    assert [item.path for item in plan] == [["Main", "Sub"]]
