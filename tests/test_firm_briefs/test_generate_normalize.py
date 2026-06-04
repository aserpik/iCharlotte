from icharlotte_core.ui.wizard.pages import generate_motion_page as gmp


def test_module_exposes_normalizer():
    assert gmp.normalize_motion_type("MSJ") == "msj"
    assert gmp.normalize_motion_type("Motion for Good Faith Settlement") == "gfs"
