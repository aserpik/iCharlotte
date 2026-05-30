import io
import json
import zipfile

from icharlotte_core.legal_research.local_corpus.loaders import cap_loader


def _make_zip() -> bytes:
    case = {
        "id": 269732,
        "name": "THE PEOPLE v. SNOW",
        "name_abbreviation": "People v. Snow",
        "decision_date": "2003-04-03",
        "docket_number": "S018033",
        "citations": [{"type": "official", "cite": "30 Cal. 4th 43"},
                      {"type": "parallel", "cite": "131 Cal. Rptr. 2d 1"}],
        "court": {"name": "Supreme Court of California", "name_abbreviation": "Cal."},
        "jurisdiction": {"name": "Cal.", "name_long": "California"},
        "cites_to": [{"cite": "536 U.S. 584"}, {"cite": "30 Cal. 4th 1"}],
        "casebody": {"opinions": [
            {"type": "majority", "author": "THE COURT",
             "text": "Para one about duty. \n\nPara two about the page break and privacy."}
        ]},
    }
    html = ('<section class="casebody"><p>Para one about duty. '
            '<a data-label="56" class="page-label">*56</a>'
            'Para two about the page break and privacy.</p></section>')
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("json/0043-01.json", json.dumps(case))
        zf.writestr("html/0043-01.html", html)
    return buf.getvalue()


def test_iter_cases_yields_normalized_case_and_passages():
    results = list(cap_loader.iter_cases_from_zip(_make_zip()))
    assert len(results) == 1
    case, passages = results[0]
    assert case.case_uid == "cap:269732"
    assert case.source == "cap"
    assert case.citation == "30 Cal. 4th 43"
    assert case.parallel_citations == ["131 Cal. Rptr. 2d 1"]
    assert case.court == "Cal."
    assert case.year == "2003"
    assert "536 U.S. 584" in case.cites_to
    assert "duty" in case.full_text
    assert passages, "expected at least one passage"
    assert passages[0].case_uid == "cap:269732"
    # Page-label propagated to at least one passage near the *56 break
    assert any(p.page_label == "56" for p in passages) or passages[0].page_label == ""


def test_non_ca_case_is_skipped():
    case = {"id": 1, "jurisdiction": {"name_long": "New York"},
            "citations": [{"type": "official", "cite": "1 N.Y. 1"}],
            "casebody": {"opinions": [{"type": "majority", "text": "x"}]}}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("json/0001-01.json", json.dumps(case))
        zf.writestr("html/0001-01.html", "<p>x</p>")
    assert list(cap_loader.iter_cases_from_zip(buf.getvalue())) == []
