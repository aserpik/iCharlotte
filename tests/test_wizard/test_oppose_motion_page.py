import pytest

pytest.importorskip("pytestqt")

from unittest.mock import patch

from icharlotte_core.opposition.models import (
    CitationVerification,
    DraftDocument,
    MotionMetadata,
    OutlineNode,
    SectionPlanItem,
)
from icharlotte_core.ui.wizard.pages.oppose_motion_page import (
    OpposeMotionOutputPage,
    OpposeMotionSettingsPage,
    OpposeMotionTaskTab,
    OpposeMotionWorker,
    SETTINGS_PAGE_CONFIRM,
    SETTINGS_PAGE_OUTLINE,
    TASK_PAGE_SETTINGS,
    TASK_PAGE_OUTPUT,
    _research_targets,
    build_oppose_motion_tab,
)


def test_research_targets_unions_args_and_section_leaves_skipping_structural():
    """Research must cover every subsection proposition, not just the top-level
    arguments — otherwise sub-points (meet-and-confer, cutoff, ...) get no
    authority and the drafter emits '[no case authority retrieved for this
    point]'. Structural sections are skipped; duplicates collapse."""
    md = MotionMetadata(principal_arguments=["Forensic imaging is irrelevant"])
    plan = [
        SectionPlanItem(id="i", path=["I"], text="Introduction"),
        SectionPlanItem(id="a", path=["II", "A"], text="Discovery on the privacy claims is closed"),
        SectionPlanItem(id="b", path=["II", "B"], text="Plaintiff failed to meet and confer"),
        SectionPlanItem(id="b2", path=["II", "B"], text="Plaintiff failed to meet and confer"),  # dup
        SectionPlanItem(id="c", path=["VIII"], text="Conclusion"),
    ]
    targets = _research_targets(md, plan)
    assert "Forensic imaging is irrelevant" in targets       # principal argument kept
    assert "Discovery on the privacy claims is closed" in targets  # subsection leaf added
    assert "Plaintiff failed to meet and confer" in targets
    assert "Introduction" not in targets and "Conclusion" not in targets  # structural skipped
    assert len(targets) == 3                                   # deduped
    assert len(targets) <= 24                                  # capped


def test_confirmation_blocks_missing_required_fields(qtbot):
    page = OpposeMotionSettingsPage(
        case_root="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(page)
    page.set_metadata(MotionMetadata())

    assert page.can_continue_to_outline() is False

    page.set_metadata(
        MotionMetadata(
            motion_type="Motion for Summary Judgment",
            relief_requested="summary judgment",
            principal_arguments=["no duty"],
        )
    )

    assert page.can_continue_to_outline() is True


def test_outline_items_start_checked(qtbot):
    page = OpposeMotionSettingsPage(
        case_root="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(page)
    page.set_outline(
        [
            OutlineNode(
                id="a",
                text="Main",
                children=[OutlineNode(id="a1", text="Sub")],
            )
        ]
    )

    assert page.outline_tree.topLevelItem(0).checkState(0).value == 2
    assert page.outline_tree.topLevelItem(0).child(0).checkState(0).value == 2


def test_task_tab_starts_on_confirmation_page(qtbot):
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)

    assert tab.settings_page.currentIndex() == SETTINGS_PAGE_CONFIRM


def test_task_tab_auto_analysis_populates_confirmation_and_outline(qtbot, monkeypatch):
    started = []

    class FakeSignal:
        def __init__(self):
            self._slots = []

        def connect(self, slot):
            self._slots.append(slot)

        def emit(self, *args):
            for slot in self._slots:
                slot(*args)

    class FakeAnalysisWorker:
        def __init__(self, settings, parent=None):
            self.settings = settings
            self.progress = FakeSignal()
            self.finished_analysis = FakeSignal()
            self.finished = FakeSignal()

        def start(self):
            started.append(self.settings)
            self.finished_analysis.emit(
                True,
                {
                    "metadata": MotionMetadata(
                        motion_type="Motion for Summary Judgment",
                        relief_requested="summary judgment",
                        principal_arguments=["no triable issue"],
                    ),
                    "outline": [
                        OutlineNode(
                            id="arg-1",
                            text="There Are Triable Issues of Material Fact",
                            children=[
                                OutlineNode(
                                    id="arg-1-a",
                                    text="The moving party ignores disputed testimony",
                                )
                            ],
                        )
                    ],
                },
            )
            self.finished.emit()

        def deleteLater(self):
            pass

        def isRunning(self):
            return False

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.OpposeMotionAnalysisWorker",
        FakeAnalysisWorker,
        raising=False,
    )

    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=["/tmp/facts.pdf"],
        auto_analyze=True,
    )
    qtbot.addWidget(tab)

    assert started[0]["motion_file"] == "/tmp/motion.pdf"
    assert started[0]["context_files"] == ["/tmp/facts.pdf"]
    assert tab.currentIndex() == TASK_PAGE_SETTINGS
    # After analysis, jump straight to the outline screen — the user no longer
    # needs to click through a confirmation step that just shows extracted metadata.
    assert tab.settings_page.currentIndex() == SETTINGS_PAGE_OUTLINE
    # Metadata is still populated on the (hidden) confirmation widgets so it
    # flows through to the drafter when the user clicks Generate Draft.
    assert tab.settings_page.motion_type_edit.text() == "Motion for Summary Judgment"
    assert tab.settings_page.relief_edit.text() == "summary judgment"
    assert "no triable issue" in tab.settings_page.arguments_edit.toPlainText()
    assert tab.settings_page.outline_tree.topLevelItem(0).text(0) == (
        "There Are Triable Issues of Material Fact"
    )
    assert tab.settings_page.outline_tree.topLevelItem(0).child(0).text(0) == (
        "The moving party ignores disputed testimony"
    )


def test_task_tab_loads_draft_result(qtbot):
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)
    draft = DraftDocument(
        title="Opposition",
        body_text="Argument text",
        citations=[
            CitationVerification(
                citation_text="69 Cal.2d 108",
                verdict="SUPPORTED",
                evidence="ordinary care",
            )
        ],
    )

    tab._on_worker_finished(True, draft)

    assert tab.currentIndex() == TASK_PAGE_OUTPUT
    assert "Argument text" in tab.output_page.editor.toPlainText()
    # Detail panel shows the verdict-aware HTML rendering — evidence is
    # rendered under the "Verified holding:" heading for SUPPORTED.
    assert "ordinary care" in tab.output_page.detail_panel.body_html


def test_settings_from_dict_accepts_missing_state_and_refreshes_motion_label(qtbot):
    page = OpposeMotionSettingsPage(
        case_root="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/old.pdf",
        context_files=[],
    )
    qtbot.addWidget(page)

    page.from_dict(None)
    assert "old.pdf" in page.motion_label.text()

    page.from_dict({"motion_file": "/tmp/new_motion.pdf"})
    assert page.motion_file == "/tmp/new_motion.pdf"
    assert "new_motion.pdf" in page.motion_label.text()


def test_task_tab_stores_last_settings_when_run_is_requested(qtbot, monkeypatch):
    started = []

    class FakeWorker:
        def __init__(self, case_path, file_number, settings, parent=None):
            self.settings = settings
            self.progress = type("Signal", (), {"connect": lambda self, slot: None})()
            self.finished_result = type("Signal", (), {"connect": lambda self, slot: None})()

        def start(self):
            started.append(self.settings)

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.OpposeMotionWorker",
        FakeWorker,
    )
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)

    tab._on_run({"motion_file": "/tmp/motion.pdf", "outline": []})

    assert tab._last_settings["motion_file"] == "/tmp/motion.pdf"
    assert started[0]["motion_file"] == "/tmp/motion.pdf"


def test_worker_reports_empty_draft_as_failure(monkeypatch, tmp_path):
    results = []
    assembled = []

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.extract_document_text",
        lambda _path: type("ExtractResult", (), {"success": True, "text": "motion", "error": ""})(),
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.extract_context_bundle",
        lambda _paths: ("context", []),
    )
    monkeypatch.setenv("COURTLISTENER_API_TOKEN", "tok")
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.draft_memorandum",
        lambda **_kwargs: DraftDocument(
            title="Opposition",
            body_text="",
            rejection_reason="LLM returned an empty body_text.",
        ),
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.assemble_opposition_preview",
        lambda **_kwargs: assembled.append(_kwargs),
    )

    worker = OpposeMotionWorker(
        case_path=str(tmp_path),
        file_number="0000.000",
        settings={
            "motion_file": "/tmp/motion.pdf",
            "context_files": [],
            "metadata": MotionMetadata(motion_type="Motion to Compel").to_dict(),
            "outline": [],
        },
    )
    worker.finished_result.connect(lambda success, payload: results.append((success, payload)))

    worker.run()

    assert results
    assert results[0][0] is False
    # The drafter's rejection_reason now flows into the worker's failure message.
    assert "LLM returned an empty body_text" in results[0][1]
    assert assembled == []


def test_worker_drafts_then_verifies_parsed_citations(monkeypatch, tmp_path):
    """End-to-end happy path: drafter -> citation parser -> verifier."""
    from unittest.mock import MagicMock

    results = []
    draft_calls = []

    # Exercise the CourtListener-API verifier path deterministically; the local
    # corpus path is covered by test_oppose_motion_local_corpus.py. Without this
    # the test's outcome depends on whether corpus.db exists on the machine.
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page._corpus_available",
        lambda: False,
    )
    monkeypatch.setenv("COURTLISTENER_API_TOKEN", "tok")
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.extract_document_text",
        lambda _path: type("ExtractResult", (), {"success": True, "text": "motion", "error": ""})(),
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.extract_context_bundle",
        lambda _paths: ("context", []),
    )

    def fake_draft_memorandum(**kwargs):
        draft_calls.append(kwargs)
        return DraftDocument(
            title="Opposition",
            body_text=(
                "Argument. *Aguilar v. Atlantic Richfield Co.* (2001) "
                "25 Cal.4th 826 controls the analysis."
            ),
        )

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.draft_memorandum",
        fake_draft_memorandum,
    )

    verifier = MagicMock()
    verifier.verify_all.return_value = [
        CitationVerification(
            citation_text="Aguilar v. Atlantic Richfield Co. 25 Cal.4th 826",
            normalized_citation="Aguilar v. Atlantic Richfield Co. 25 Cal.4th 826",
            verdict="SUPPORTED",
            kind="case",
            evidence="summary judgment burden passage",
        )
    ]
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.build_opposition_verifier",
        lambda **_kwargs: verifier,
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.assemble_opposition_preview",
        lambda **_kwargs: None,
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.validate_opposition_docx",
        lambda _path: type("Validation", (), {"has_errors": False})(),
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.DiscoveryAssembler.find_caption_page",
        lambda _case_path: "",
    )

    worker = OpposeMotionWorker(
        case_path=str(tmp_path),
        file_number="0000.000",
        settings={
            "motion_file": "/tmp/motion.pdf",
            "context_files": ["/tmp/context.docx"],
            "metadata": MotionMetadata(motion_type="Motion for Summary Judgment").to_dict(),
            "outline": [OutlineNode(id="a", text="Triable Issues").to_dict()],
        },
    )
    worker.finished_result.connect(lambda success, payload: results.append((success, payload)))

    worker.run()

    assert results and results[0][0] is True
    # Drafter now gets style_exemplars (not the old authority_block).
    assert "style_exemplars" in draft_calls[0]
    assert "authority_block" not in draft_calls[0]
    # Verifier was invoked with the parsed citation list.
    verifier.verify_all.assert_called_once()
    parsed_citations = verifier.verify_all.call_args.args[0]
    assert parsed_citations and parsed_citations[0].kind == "case"
    # The verdict-bearing citation flows through to the returned draft.
    assert results[0][1].citations[0].verdict == "SUPPORTED"


def test_worker_skips_verifier_when_draft_has_no_citations(
    monkeypatch, tmp_path
):
    """When the drafted body contains no citations, the verifier is not built."""
    from unittest.mock import MagicMock

    results = []
    draft_calls = []
    progress_msgs = []
    bov_calls: list = []

    monkeypatch.setenv("COURTLISTENER_API_TOKEN", "tok")
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.extract_document_text",
        lambda _path: type("ExtractResult", (), {"success": True, "text": "motion", "error": ""})(),
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.extract_context_bundle",
        lambda _paths: ("context", []),
    )

    def fake_draft_memorandum(**kwargs):
        draft_calls.append(kwargs)
        # Body has no citations the parser can detect.
        return DraftDocument(title="Opposition", body_text="Plaintiff opposes the motion.")

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.draft_memorandum",
        fake_draft_memorandum,
    )

    def fake_build_verifier(**kwargs):
        bov_calls.append(kwargs)
        return MagicMock()

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.build_opposition_verifier",
        fake_build_verifier,
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.assemble_opposition_preview",
        lambda **_kwargs: None,
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.validate_opposition_docx",
        lambda _path: type("Validation", (), {"has_errors": False})(),
    )
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.DiscoveryAssembler.find_caption_page",
        lambda _case_path: "",
    )

    worker = OpposeMotionWorker(
        case_path=str(tmp_path),
        file_number="0000.000",
        settings={
            "motion_file": "/tmp/motion.pdf",
            "context_files": [],
            "metadata": MotionMetadata(motion_type="Demurrer").to_dict(),
            "outline": [],
        },
    )
    worker.progress.connect(lambda msg: progress_msgs.append(msg))
    worker.finished_result.connect(lambda success, payload: results.append((success, payload)))

    worker.run()

    # Success even with no citations.
    assert results and results[0][0] is True
    # Drafter was called with style_exemplars (not authority_block).
    assert draft_calls and "style_exemplars" in draft_calls[0]
    assert "authority_block" not in draft_calls[0]
    # User was warned about the missing citations.
    assert any("No citations detected" in msg for msg in progress_msgs)
    # Verifier is not constructed when there is nothing to verify.
    assert bov_calls == []
    # Returned draft carries no citations.
    assert results[0][1].citations == []


def test_task_tab_does_not_start_second_worker_while_running(qtbot, monkeypatch):
    started = []

    class FakeWorker:
        def __init__(self, case_path, file_number, settings, parent=None):
            self.settings = settings
            self.progress = type("Signal", (), {"connect": lambda self, slot: None})()
            self.finished_result = type("Signal", (), {"connect": lambda self, slot: None})()

        def start(self):
            started.append(self.settings)

    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.OpposeMotionWorker",
        FakeWorker,
    )
    tab = OpposeMotionTaskTab(
        spec=type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})(),
        case_path="/tmp/case",
        file_number="0000.000",
        motion_file="/tmp/motion.pdf",
        context_files=[],
    )
    qtbot.addWidget(tab)

    tab._on_run({"motion_file": "/tmp/first.pdf"})
    tab._on_run({"motion_file": "/tmp/second.pdf"})

    assert len(started) == 1
    assert started[0]["motion_file"] == "/tmp/first.pdf"


def test_output_page_show_citation_updates_detail_panel(qtbot):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    draft = DraftDocument(
        title="Opposition",
        body_text="Text",
        citations=[
            CitationVerification(
                citation_text="69 Cal.2d 108",
                normalized_citation="69 Cal.2d 108",
                case_name="Rowland v. Christian",
                verdict="SUPPORTED",
                proposition="A duty of ordinary care applies.",
                evidence="ordinary care language",
                note="Direct support.",
                opinion_url="https://www.courtlistener.com/opinion/123/",
            )
        ],
    )
    page.show_result(draft)
    page.show_citation(0)

    # The detail panel renders verdict-aware content rather than a raw dump.
    header_text = page.detail_panel.header_label.text()
    body_html = page.detail_panel.body_html
    assert "SUPPORTED" in header_text
    assert "Rowland v. Christian" in header_text
    assert "ordinary care language" in body_html
    assert "ordinary care applies" in body_html
    # The "Open in CourtListener" button is shown (citation has an
    # opinion_url). Use isHidden() since the page isn't show()'n in headless tests.
    assert not page.detail_panel.open_btn.isHidden()


def test_output_editor_is_read_only(qtbot):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)

    assert page.editor.isReadOnly() is True


def test_output_renders_citations_as_clickable_anchors(qtbot):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    draft = DraftDocument(
        title="Opposition",
        body_text=(
            "In *Sinaiko Healthcare Consulting, Inc. v. Pacific Healthcare Consultants* "
            "(2007) 148 Cal. App. 4th 390, the court explained the rule. "
            "See also 195 Cal. App. 4th 1275."
        ),
        citations=[
            CitationVerification(
                citation_text="148 Cal. App. 4th 390",
                case_name="Sinaiko Healthcare Consulting, Inc. v. Pacific Healthcare Consultants",
                status="exists_support_unconfirmed",
            ),
            CitationVerification(
                citation_text="195 Cal. App. 4th 1275",
                case_name="Simke, Chodos, Silberfeld & Anteau, Inc. v. Athans",
                status="exists_support_unconfirmed",
            ),
        ],
    )
    page.show_result(draft)

    html_text = page.editor.toHtml()
    # Each citation gets a clickable anchor with index in the URL.
    assert "citation:0" in html_text
    assert "citation:1" in html_text
    # The case name was italicized via *...* markdown syntax (QTextBrowser
    # normalizes <i> into inline font-style:italic).
    assert "font-style:italic" in html_text


def test_output_anchor_click_updates_detail_panel_without_popup(qtbot, monkeypatch):
    # Anchor clicks must update the right-side detail panel in place; the
    # legacy popup dialog must NOT auto-open.
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    draft = DraftDocument(
        title="Opposition",
        body_text="See 148 Cal. App. 4th 390.",
        citations=[
            CitationVerification(
                citation_text="148 Cal. App. 4th 390",
                case_name="Sinaiko Healthcare",
                opinion_url="https://example.com/op/1",
                verdict="SUPPORTED",
                evidence="Sample evidence quote.",
            )
        ],
    )
    page.show_result(draft)

    # Guard: if the popup ever sneaks back in, we want the test to fail loudly.
    dialog_opens: list[object] = []
    from icharlotte_core.ui.wizard.pages import oppose_motion_page as page_mod

    monkeypatch.setattr(
        page_mod.CitationDetailDialog,
        "exec",
        lambda self: dialog_opens.append(self) or 0,
    )

    from PySide6.QtCore import QUrl

    page._on_anchor_clicked(QUrl("citation:0"))

    assert dialog_opens == []  # no popup
    assert "Sinaiko Healthcare" in page.detail_panel.header_label.text()
    assert "Sample evidence quote" in page.detail_panel.body_html


def test_citation_detail_dialog_shows_supporting_passage(qtbot):
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import CitationDetailDialog

    citation = CitationVerification(
        citation_text="148 Cal. App. 4th 390",
        case_name="Sinaiko Healthcare Consulting, Inc. v. Pacific Healthcare Consultants",
        court="California Court of Appeal",
        date="2007-03-08",
        opinion_url="https://www.courtlistener.com/opinion/2271540/",
        verdict="SUPPORTED",
        evidence="The party fails to make a timely response, the propounding party may move.",
    )
    dialog = CitationDetailDialog(citation)
    qtbot.addWidget(dialog)

    assert "timely response" in dialog.body_html


def test_citation_detail_dialog_explains_unverified_verdict(qtbot):
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import CitationDetailDialog

    citation = CitationVerification(
        citation_text="148 Cal. App. 4th 390",
        case_name="Sinaiko Healthcare",
        opinion_url="https://example.com",
        verdict="UNVERIFIED",
    )
    dialog = CitationDetailDialog(citation)
    qtbot.addWidget(dialog)

    # No passage; user sees an explanatory note for the unverified verdict.
    assert "verifier" in dialog.body_html.lower()


def test_citation_detail_dialog_omits_open_button_without_url(qtbot):
    from icharlotte_core.ui.wizard.pages.oppose_motion_page import CitationDetailDialog

    citation = CitationVerification(
        citation_text="148 Cal. App. 4th 390",
        case_name="Sinaiko Healthcare",
        opinion_url="",
        verdict="NOT_FOUND",
    )
    dialog = CitationDetailDialog(citation)
    qtbot.addWidget(dialog)

    # The new dialog omits the "Open in CourtListener" button entirely when
    # there is no opinion_url, so no enabled-state needs to be checked.
    from PySide6.QtWidgets import QPushButton

    open_buttons = [
        b for b in dialog.findChildren(QPushButton)
        if "open" in b.text().lower()
    ]
    assert open_buttons == []


def test_save_as_uses_dialog_and_does_not_save_when_cancelled(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    source = tmp_path / "preview.docx"
    source.write_bytes(b"docx-bytes")
    page.show_result(
        DraftDocument(
            title="Opposition",
            body_text="Text",
            preview_path=str(source),
        )
    )

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=("", ""),
    ) as dialog, patch(
        "icharlotte_core.ui.wizard.pages.citation_review.shutil.copyfile"
    ) as copyfile, patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.information"
    ) as info:
        page.save_as()

    assert dialog.called
    copyfile.assert_not_called()
    info.assert_not_called()


def test_save_as_defaults_outside_internal_preview_folder(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    preview = (
        tmp_path
        / "NOTES"
        / "AI OUTPUT"
        / ".icharlotte"
        / "wizard_previews"
        / "oppose_motion"
        / "preview.docx"
    )
    preview.parent.mkdir(parents=True)
    preview.write_bytes(b"docx-bytes")
    page.show_result(
        DraftDocument(
            title="Opposition",
            body_text="Text",
            preview_path=str(preview),
        )
    )

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=("", ""),
    ) as dialog:
        page.save_as()

    suggested_path = dialog.call_args.args[2]
    assert ".icharlotte" not in suggested_path
    assert suggested_path.endswith("Opposition.docx")


def test_save_as_appends_docx_and_copies_preview(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    source = tmp_path / "preview.docx"
    source.write_bytes(b"docx-bytes")
    target = tmp_path / "saved-opposition"
    page.show_result(
        DraftDocument(
            title="Opposition",
            body_text="Text",
            preview_path=str(source),
        )
    )

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=(str(target), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.information"
    ) as info:
        page.save_as()

    assert (tmp_path / "saved-opposition.docx").read_bytes() == b"docx-bytes"
    assert info.called


def test_save_as_reports_copy_errors(qtbot, tmp_path):
    page = OpposeMotionOutputPage()
    qtbot.addWidget(page)
    source = tmp_path / "preview.docx"
    source.write_bytes(b"docx-bytes")
    target = tmp_path / "saved.docx"
    page.show_result(
        DraftDocument(
            title="Opposition",
            body_text="Text",
            preview_path=str(source),
        )
    )

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getSaveFileName",
        return_value=(str(target), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.citation_review.shutil.copyfile",
        side_effect=OSError("locked"),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.critical"
    ) as critical:
        page.save_as()

    assert critical.called


def test_builder_rejects_cancelled_motion_picker(qtbot):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=("", ""),
    ):
        tab = build_oppose_motion_tab(spec, "/tmp/case", "0000.000", None)
    assert tab is None


def test_builder_rejects_unsupported_motion_file(qtbot):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=("/tmp/motion.txt", ""),
    ):
        with patch(
            "icharlotte_core.ui.wizard.pages.oppose_motion_page.QMessageBox.warning"
        ) as warning:
            tab = build_oppose_motion_tab(spec, "/tmp/case", "0000.000", None)
    assert tab is None
    assert warning.called


def test_builder_filters_unsupported_context_files(qtbot, tmp_path, monkeypatch):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    motion = tmp_path / "motion.pdf"
    motion.write_bytes(b"")
    good_context = tmp_path / "facts.txt"
    bad_context = tmp_path / "notes.xlsx"
    good_context.write_text("facts")
    bad_context.write_text("spreadsheet")
    monkeypatch.setattr(OpposeMotionTaskTab, "_start_analysis", lambda self: None)

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=(str(motion), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.ContextFilesDialog.get_files",
        return_value=[str(good_context), str(bad_context)],
    ):
        tab = build_oppose_motion_tab(spec, str(tmp_path), "0000.000", None)

    qtbot.addWidget(tab)
    assert tab.settings_page.context_files == [str(good_context)]


def test_builder_accumulates_multifolder_context_files(qtbot, tmp_path, monkeypatch):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    motion = tmp_path / "motion.pdf"
    motion.write_bytes(b"")
    folder_a = tmp_path / "DISCOVERY"
    folder_b = tmp_path / "RECORDS"
    folder_a.mkdir()
    folder_b.mkdir()
    ctx_a = folder_a / "depo.pdf"
    ctx_b = folder_b / "records.pdf"
    ctx_a.write_text("a")
    ctx_b.write_text("b")
    monkeypatch.setattr(OpposeMotionTaskTab, "_start_analysis", lambda self: None)

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=(str(motion), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.ContextFilesDialog.get_files",
        return_value=[str(ctx_a), str(ctx_b)],
    ):
        tab = build_oppose_motion_tab(spec, str(tmp_path), "0000.000", None)

    qtbot.addWidget(tab)
    assert tab.settings_page.context_files == [str(ctx_a), str(ctx_b)]


def test_builder_handles_cancelled_context_picker(qtbot, tmp_path, monkeypatch):
    spec = type("Spec", (), {"task_id": "oppose_motion", "title": "Oppose a Motion"})()
    motion = tmp_path / "motion.pdf"
    motion.write_bytes(b"")
    monkeypatch.setattr(OpposeMotionTaskTab, "_start_analysis", lambda self: None)

    with patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.QFileDialog.getOpenFileName",
        return_value=(str(motion), ""),
    ), patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.ContextFilesDialog.get_files",
        return_value=None,
    ):
        tab = build_oppose_motion_tab(spec, str(tmp_path), "0000.000", None)

    qtbot.addWidget(tab)
    assert tab.settings_page.context_files == []


def test_worker_calls_verifier_with_parsed_citations(tmp_path, monkeypatch):
    """Worker plumbing test: OpposeMotionWorker invokes the new pipeline."""
    from unittest.mock import MagicMock, patch as _patch

    from icharlotte_core.opposition.models import DraftDocument, MotionMetadata

    monkeypatch.setenv("COURTLISTENER_API_TOKEN", "dummy-token")
    # Exercise the CourtListener-API verifier path deterministically; the local
    # corpus path is covered by test_oppose_motion_local_corpus.py.
    monkeypatch.setattr(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page._corpus_available",
        lambda: False,
    )

    motion_pdf = tmp_path / "motion.pdf"
    motion_pdf.write_bytes(b"%PDF-1.4 dummy")

    fake_motion_text = MagicMock(success=True, text="Motion text body.", error="")
    fake_metadata = MotionMetadata(
        motion_type="Motion to Compel",
        relief_requested="An order compelling responses",
        principal_arguments=["Late responses"],
    )
    fake_draft = DraftDocument(
        title="Opposition to MTC",
        body_text="The court held in *Smith v. Jones* (2010) 50 Cal.4th 100 that ...",
    )

    with _patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.extract_document_text",
        return_value=fake_motion_text,
    ), _patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.extract_context_bundle",
        return_value=("ctx text", []),
    ), _patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.research_arguments",
        return_value=[],
    ), _patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.draft_memorandum",
        return_value=fake_draft,
    ) as draft_fn, _patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.build_opposition_verifier"
    ) as bov, _patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.assemble_opposition_preview"
    ), _patch(
        "icharlotte_core.ui.wizard.pages.oppose_motion_page.validate_opposition_docx"
    ) as validate:
        verifier = MagicMock()
        verifier.verify_all.return_value = []
        bov.return_value = verifier
        validate.return_value = MagicMock(has_errors=False)

        worker = OpposeMotionWorker(
            case_path=str(tmp_path),
            file_number="X",
            settings={
                "motion_file": str(motion_pdf),
                "context_files": [],
                "metadata": fake_metadata.to_dict(),
                "outline": [],
            },
        )

        # Run the worker body directly (not via QThread.start) so we can assert.
        results: list = []
        worker.finished_result.connect(lambda ok, payload: results.append((ok, payload)))
        worker.run()

        # Drafter received style_exemplars (not authority_block).
        assert draft_fn.call_args is not None
        call_kwargs = draft_fn.call_args.kwargs
        assert "style_exemplars" in call_kwargs
        assert "authority_block" not in call_kwargs

        # Verifier was constructed + called with parsed citations.
        bov.assert_called_once()
        verifier.verify_all.assert_called_once()


def test_worker_grounds_draft_and_runs_pool_check(monkeypatch, tmp_path):
    import icharlotte_core.ui.wizard.pages.oppose_motion_page as page
    from icharlotte_core.opposition.models import (
        DraftDocument, MotionMetadata, RetrievedAuthority,
    )
    from icharlotte_core.opposition.citation_parser import Citation

    # The research branch only runs when a CourtListener token is present.
    monkeypatch.setenv("COURTLISTENER_API_TOKEN", "test-token")

    # Stub document extraction so no real file is read.
    monkeypatch.setattr(page, "extract_document_text",
                        lambda p: type("R", (), {"success": True, "text": "motion text", "error": ""})())
    monkeypatch.setattr(page, "extract_context_bundle", lambda files: ("ctx", []))

    captured = {}

    def fake_research(arguments, **kwargs):
        captured["arguments"] = list(arguments)
        if kwargs.get("on_progress"):
            kwargs["on_progress"]("  arg — 1 case(s) found")
        return [RetrievedAuthority(argument_text="arg one", cluster_id="1",
                                   case_name="A v. B", citation="226 Cal.App.4th 401",
                                   supports="s", passage="p")]
    monkeypatch.setattr(page, "research_arguments", fake_research)

    def fake_draft(**kwargs):
        captured["authorities"] = kwargs.get("retrieved_authorities")
        return DraftDocument(title="Opposition", body_text="*A v. B* (2014) 226 Cal.App.4th 401 controls.")
    monkeypatch.setattr(page, "draft_memorandum", fake_draft)

    # Verifier returns whatever it is asked to verify, marked SUPPORTED.
    class FakeVerifier:
        def verify_all(self, cites, on_progress=None):
            from icharlotte_core.opposition.models import CitationVerification
            return [CitationVerification(citation_text=c.raw_text, kind=c.kind,
                                         verdict="SUPPORTED", body_offset=c.body_offset) for c in cites]
    monkeypatch.setattr(page, "build_opposition_verifier", lambda **kw: FakeVerifier())

    # Skip real assembly/validation.
    monkeypatch.setattr(page, "assemble_opposition_preview", lambda **kw: None)
    monkeypatch.setattr(page, "validate_opposition_docx",
                        lambda p: type("V", (), {"has_errors": False})())
    monkeypatch.setattr(page.DiscoveryAssembler, "find_caption_page", staticmethod(lambda p: ""))

    settings = {
        "motion_file": "m.pdf", "context_files": [],
        "metadata": MotionMetadata(motion_type="MTC", relief_requested="x",
                                   principal_arguments=["arg one", "arg two"]).to_dict(),
        "outline": [],
    }
    worker = page.OpposeMotionWorker(case_path=str(tmp_path), file_number="123", settings=settings)

    results = {}
    worker.finished_result.connect(lambda ok, payload: results.update(ok=ok, payload=payload))
    worker.run()   # run synchronously (do not start the thread)

    assert results["ok"] is True
    assert captured["arguments"] == ["arg one", "arg two"]
    # The pool was forwarded to the drafter.
    assert captured["authorities"] and captured["authorities"][0].cluster_id == "1"
