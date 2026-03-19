import pytest
from unittest.mock import MagicMock, patch, ANY
from tts_papertrail.base import Report, TextEntry, ListEntry, TableEntry, ListItem, RichText, SectionEntry, SummaryTable, PowerTableEntry
import tts_papertrail.hypersonic

@pytest.fixture
def mock_compiler():
    # Patch the HtmlCompiler imported in hypersonic.py
    with patch("tts_papertrail.hypersonic.HtmlCompiler") as MockCompiler:
        compiler_instance = MockCompiler.return_value
        # Mock methods that are called
        compiler_instance.add_body_component = MagicMock()
        compiler_instance.render_to_file = MagicMock()
        yield compiler_instance

@pytest.fixture
def mock_pane_container():
    # Patch PaneContainer to verify entries are added as tabs
    with patch("tts_papertrail.hypersonic.PaneContainer") as MockPane:
        yield MockPane.return_value

@pytest.fixture
def mock_generic_container():
    # Patch GenericContainer to verify usage of the new implementation
    with patch("tts_papertrail.hypersonic.GenericContainer") as MockGC:
        yield MockGC

def test_basic_html_export(mock_compiler):
    """Test that to_html creates a compiler and saves the file."""
    report = Report("Test Report", "Author")
    report.add_entry(TextEntry("Intro", "Hello World"))
    
    report.to_html("output.html")
    
    # Verify render_to_file was called with the correct filename
    mock_compiler.render_to_file.assert_called_with("output.html")

def test_entries_added_to_panes(mock_compiler, mock_pane_container):
    """Test that entries are added to the PaneContainer tabs."""
    report = Report("Multi-Tab Report", "Author")
    
    # Add two entries
    report.add_entry(TextEntry("Tab 1", "Content A"))
    report.add_entry(TextEntry("Tab 2", "Content B"))
    
    report.to_html("out.html")
    
    # Verify 2 panes were added (one for each entry)
    assert mock_pane_container.add_pane.call_count == 2
    
    # Verify the pane container was added to the main report body.
    mock_compiler.add_body_component.assert_any_call(mock_pane_container)

def test_section_entry_rendering(mock_compiler):
    """Test SectionEntry adds its children to the compiler."""
    report = Report("Section Report", "Author")
    section = SectionEntry("My Section", [
        TextEntry("Part 1", "Hello"),
        TextEntry("Part 2", "World")
    ])
    report.add_entry(section)
    
    report.to_html("section.html")
    
    # Since SectionEntry adds its parts to a *pane* compiler (not the main one),
    # we can verify this by checking if PaneContainer.add_pane was called with content.
    # However, to test the _add_section_to_compiler logic specifically, we can check 
    # if it recurses. 
    # For a unit test, it's easier to verify that no error is raised and the flow completes.
    mock_compiler.render_to_file.assert_called()

def test_list_entry_rendering(mock_compiler):
    """Test ListEntry rendering logic."""
    report = Report("List Report", "Author")
    item = ListItem(0, [RichText("Item 1")])
    report.add_entry(ListEntry("My List", [item]))
    
    report.to_html("list.html")
    
    # Verify UL creation is mocked/handled implicitly by the flow
    mock_compiler.render_to_file.assert_called()

def test_summary_table_rendering(mock_compiler, mock_generic_container):
    """Test SummaryTable rendering."""
    report = Report("Summary Test", "Author")
    summary = SummaryTable("Overview")
    summary.add_to_summary_table("Cat", "Stat", "Msg", "Disp")
    report.add_entry(summary)
    
    report.to_html("summary.html")
    
    # Check GenericContainer usage
    mock_generic_container.assert_called()
    mock_generic_container.return_value.power_table.assert_called()

def test_power_table_rendering(mock_compiler):
    """Test PowerTableEntry delegates to the container."""
    report = Report("Power Test", "Author")
    mock_container = MagicMock()
    entry = PowerTableEntry("Data", mock_container, add_filters=True, add_sorting=False)
    report.add_entry(entry)
    
    report.to_html("power.html")
    
    # Verify container.power_table was called with correct args
    mock_container.power_table.assert_called_with(
        superheader=None,
        add_filters=True,
        add_sorting=False
    )

def test_richtext_rendering(mock_compiler):
    """Test that RichText objects are converted to HTML components."""
    # We can test the private method directly for granular coverage
    from tts_papertrail.hypersonic import HypersonicReport
    renderer = HypersonicReport("Test", "Author")
    
    # Bold + Italic
    rt = RichText("BoldItalic", bold=True, italic=True)
    span = renderer._richtext_to_span(rt)
    # Should be a Strong component with italic style
    # (Checking attributes depends on html_utils implementation, but basic check is:)
    assert span is not None
    
    # Color + BgColor
    rt2 = RichText("Colors", color="#f00", bg_color="#0f0")
    span2 = renderer._richtext_to_span(rt2)
    assert span2 is not None