import pytest
from icharlotte_core.firm_briefs.motion_taxonomy import normalize_motion_type, display_name

@pytest.mark.parametrize("text,expected", [
    ("Defendant's Notice of Motion and Motion for Summary Judgment", "msj"),
    ("MSJ", "msj"),
    ("Motion for Summary Adjudication", "msj"),
    ("Motion to Compel Further Responses to RFP, Set One", "compel"),
    ("MTC", "compel"),
    ("Demurrer to First Amended Complaint", "demurrer"),
    ("Motion to Strike Punitive Damages", "strike"),
    ("Motion in Limine No. 3", "in_limine"),
    ("Motion to Quash Service of Summons", "quash"),
    ("Motion for Terminating Sanctions", "sanctions"),
    ("Motion to be Relieved as Counsel", "relieve_counsel"),
    ("Motion to Continue Trial and All Related Dates", "continue_trial"),
    ("Motion for Trial Preference", "continue_trial"),
    ("Ex Parte Application to Advance Hearing", "ex_parte"),
    ("Ex Parte Application to Continue Trial", "ex_parte"),   # ex parte wins over continue
    ("Motion for Leave to Amend Complaint", "leave"),
    ("Motion for Leave to File Cross-Complaint", "leave"),
    ("Motion for Leave to Conduct IME of Plaintiff", "ime"),  # IME wins over leave
    ("Notice of Motion for Independent Medical Examination", "ime"),
    ("Motion for Determination of Good Faith Settlement", "gfs"),
    ("Defendant's Motion GFS", "gfs"),
    ("Motion to Dismiss for Forum Non Conveniens", "dismiss"),
    ("Motion to Consolidate Related Cases", "consolidate"),
    ("Motion for Reconsideration", "reconsider"),
    ("Motion for Protective Order", "protective_order"),
    ("Motion to Set Aside Default", "set_aside_default"),
    ("Motion to Tax Costs", "other"),          # real but unmapped -> other
    ("", "other"),
    (None, "other"),
])
def test_normalize(text, expected):
    assert normalize_motion_type(text) == expected

def test_display_name():
    assert display_name("msj") == "Motion for Summary Judgment/Adjudication"
    assert display_name("ime") == "Motion for Leave to Conduct IME"
    assert display_name("nonexistent") == "nonexistent"
