from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp
from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_pages_expose_attach_provenance():
    assert hasattr(omp, "attach_firm_provenance")
    assert hasattr(gmp, "attach_firm_provenance")
