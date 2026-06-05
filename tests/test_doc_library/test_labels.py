from icharlotte_core.doc_library.labels import auto_label


def test_deposition_with_party():
    lbl = auto_label("summarize_depositions", [r"Z:\x\depo.pdf"],
                     {"party": "Plaintiff"}, existing_labels=[])
    assert lbl == "Plaintiff's Deposition Transcript"


def test_deposition_without_party_drops_possessive():
    lbl = auto_label("summarize_depositions", [r"Z:\x\depo.pdf"],
                     {}, existing_labels=[])
    assert lbl == "Deposition Transcript"


def test_deposition_with_deponent_name():
    # Deponent name takes priority over party -> "Deposition of <name>".
    lbl = auto_label("summarize_depositions", [r"Z:\x\depo.pdf"],
                     {"name": "Joe Smith", "party": "Plaintiff"}, existing_labels=[])
    assert lbl == "Deposition of Joe Smith"


def test_discovery_with_party():
    lbl = auto_label("summarize_discovery", [r"Z:\x\rfp.pdf"],
                     {"party": "Defendant"}, existing_labels=[])
    assert lbl == "Defendant's Discovery Responses"


def test_medical_with_name():
    lbl = auto_label("medical_records", [r"Z:\x\recs.pdf"],
                     {"name": "Brier Buchalter"}, existing_labels=[])
    assert lbl == "Medical Records — Brier Buchalter"


def test_summarize_documents_uses_clean_filename():
    lbl = auto_label("summarize_documents",
                     [r"Z:\x\TRAFFIC_COLLISION_REPORT.pdf"],
                     {}, existing_labels=[])
    assert lbl == "Traffic Collision Report"


def test_multi_file_summarize_documents_suffix():
    lbl = auto_label("summarize_documents",
                     [r"Z:\x\a.pdf", r"Z:\x\b.pdf", r"Z:\x\c.pdf"],
                     {}, existing_labels=[])
    assert lbl == "A +2 more"


def test_collision_appends_numeric_suffix():
    existing = ["Deposition Transcript"]
    lbl = auto_label("summarize_depositions", [r"Z:\x\d.pdf"], {}, existing)
    assert lbl == "Deposition Transcript (2)"
