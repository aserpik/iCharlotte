from icharlotte_core.opposition import argument_research


class _StrIdClient:
    """Mimics LocalCaseCorpus: get_opinion_text takes a STRING uid."""
    def __init__(self):
        self.calls = []
    def get_opinion_text(self, uid):
        self.calls.append(uid)
        assert isinstance(uid, str)        # must NOT be int()-cast
        return "opinion text" if uid == "cap:1" else ""


def test_opinion_text_passes_string_uid_through(tmp_path):
    client = _StrIdClient()
    text = argument_research._opinion_text(client, str(tmp_path), "cap:1")
    assert text == "opinion text"
    assert client.calls == ["cap:1"]
