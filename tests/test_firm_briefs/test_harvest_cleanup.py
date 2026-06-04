from icharlotte_core.firm_briefs.citation_harvest import clean_case_name

def test_strips_signal_words():
    assert clean_case_name("See Townsend v. Superior Court") == "Townsend v. Superior Court"
    assert clean_case_name("See, e.g., Blank v. Kirwan") == "Blank v. Kirwan"
    assert clean_case_name("In Beckstead v. Superior Court") == "Beckstead v. Superior Court"
    assert clean_case_name("Cf. Ellis v. Toshiba") == "Ellis v. Toshiba"
    assert clean_case_name("Accord Sangster v. Paetkau") == "Sangster v. Paetkau"

def test_collapses_whitespace_and_newlines():
    assert clean_case_name("North \nCoast Business Park v. Nielsen") == "North Coast Business Park v. Nielsen"

def test_leaves_clean_names_untouched():
    assert clean_case_name("Dore v. Arnold Worldwide, Inc.") == "Dore v. Arnold Worldwide, Inc."
    assert clean_case_name("") == ""
