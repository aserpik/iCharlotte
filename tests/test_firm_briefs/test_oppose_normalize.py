from icharlotte_core.ui.wizard.pages import oppose_motion_page as omp


def test_module_exposes_normalizer():
    # The page must import the canonical normalizer for match-time use.
    assert omp.normalize_motion_type("Defendant's Motion for Summary Judgment") == "msj"
    assert omp.normalize_motion_type("Motion to Compel Further Responses") == "compel"
