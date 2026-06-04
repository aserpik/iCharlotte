from icharlotte_core.ui.wizard.pages import _motion_research_support as mrs
from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_shared_module_exposes_helpers():
    for name in ("make_firm_provider", "make_local_corpus", "research_targets",
                 "firm_style_exemplars"):
        assert hasattr(mrs, name)


def test_pages_use_shared_helpers():
    # Both pages re-export / reference the shared helpers (no private cross-import).
    assert omp._make_firm_provider is mrs.make_firm_provider
    assert gmp._make_firm_provider is mrs.make_firm_provider
    assert gmp._firm_style_exemplars is mrs.firm_style_exemplars
