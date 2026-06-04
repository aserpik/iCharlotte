# tests/test_firm_briefs/test_path_meta.py
from icharlotte_core.firm_briefs.path_meta import meta_for_path

ROOT = r"C:\lib\5800_AMTRUST_Pleadings_PDFs"


def test_moving_motion_folder():
    p = ROOT + r"\Motion - Summary Judgment\013 - Hall__msj.pdf"
    assert meta_for_path(p, ROOT) == ("msj", "moving")


def test_opposition_subfolder():
    p = ROOT + r"\Oppositions\Motion to Compel\008 - Rosas__opp.pdf"
    assert meta_for_path(p, ROOT) == ("compel", "opposition")


def test_reply_subfolder():
    p = ROOT + r"\Replies\Demurrer\072 - Forney__reply.pdf"
    assert meta_for_path(p, ROOT) == ("demurrer", "reply")


def test_pleading_folder():
    p = ROOT + r"\Pleadings - Answer\002 - Campos__answer.pdf"
    assert meta_for_path(p, ROOT) == ("answer", "pleading")


def test_ex_parte_is_moving():
    p = ROOT + r"\Ex Parte Applications\Continue Trial\x.pdf"
    assert meta_for_path(p, ROOT) == ("ex_parte", "moving")


def test_support_and_other_return_none():
    assert meta_for_path(ROOT + r"\_Support - Notices\x.pdf", ROOT) is None
    assert meta_for_path(ROOT + r"\_Other\x.pdf", ROOT) is None
