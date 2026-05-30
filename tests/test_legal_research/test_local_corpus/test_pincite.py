from icharlotte_core.legal_research.local_corpus.pincite import (
    page_label_map, page_label_for_offset,
)

SAMPLE_HTML = """
<section class="casebody">
  <p id="b1">Opening text on page fifty-five.
     <a data-label="56" class="page-label">*56</a>Now we are on page 56 and continue.
     <a data-label="57" class="page-label">*57</a>Page 57 content here.</p>
</section>
"""


def test_page_label_map_extracts_ordered_breaks():
    breaks = page_label_map(SAMPLE_HTML)
    labels = [lbl for _off, lbl in breaks]
    assert labels == ["56", "57"]
    # offsets strictly increasing
    offs = [off for off, _lbl in breaks]
    assert offs == sorted(offs)


def test_page_label_for_offset_returns_enclosing_page():
    breaks = page_label_map(SAMPLE_HTML)
    first_break_off = breaks[0][0]
    # Just before the first *56 marker -> no preceding label -> empty string.
    assert page_label_for_offset(breaks, max(0, first_break_off - 5)) == ""
    # At/after the *56 marker -> "56"
    assert page_label_for_offset(breaks, first_break_off + 1) == "56"
    # After the *57 marker -> "57"
    assert page_label_for_offset(breaks, breaks[1][0] + 1) == "57"
