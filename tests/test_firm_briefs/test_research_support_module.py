from icharlotte_core.ui.wizard.pages import _motion_research_support as mrs
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp
from icharlotte_core.opposition.models import MotionMetadata, SectionPlanItem


def test_shared_module_exposes_helpers():
    for name in ("make_firm_provider", "make_local_corpus", "research_targets",
                 "firm_style_exemplars"):
        assert hasattr(mrs, name)


def test_pages_use_shared_helpers():
    # Both pages re-export / reference the shared helpers (no private cross-import).
    assert omp._make_firm_provider is mrs.make_firm_provider
    assert gmp._make_firm_provider is mrs.make_firm_provider
    assert gmp._firm_style_exemplars is mrs.firm_style_exemplars


def test_research_targets_skip_structural_and_near_duplicate_points():
    metadata = MotionMetadata(
        principal_arguments=[
            "The First Amended Complaint fails to state facts sufficient to constitute a cause of action.",
            "The fraud claim is not pleaded with specificity.\x00",
        ]
    )
    plan = [
        SectionPlanItem(path=["Argument"], text="Argument"),
        SectionPlanItem(path=["Argument", "Legal Standard"], text="Legal Standard"),
        SectionPlanItem(
            path=["Argument", "Failure to State Facts"],
            text="A. The First Amended Complaint Fails to State Facts Sufficient to Constitute a Cause of Action",
        ),
        SectionPlanItem(
            path=["Argument", "Fraud"],
            text="2. The Third Cause of Action for Fraud Is Not Pleaded With Specificity",
        ),
        SectionPlanItem(
            path=["Argument", "UCL"],
            text="The UCL claim fails without a predicate unlawful, unfair, or fraudulent practice.",
        ),
    ]

    targets = mrs.research_targets(metadata, plan)
    lowered = [target.lower() for target in targets]

    assert "argument" not in lowered
    assert "legal standard" not in lowered
    assert sum("facts sufficient" in target for target in lowered) == 1
    assert sum("fraud" in target and "specificity" in target for target in lowered) == 1
    assert any("ucl" in target for target in lowered)
    assert all("\x00" not in target for target in targets)
