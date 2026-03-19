import pytest
from unittest.mock import MagicMock, patch, ANY
from tts_papertrail.base import Report, TableEntry, TextEntry, RichText, SectionEntry, ListEntry, ListItem, PowerTableEntry

@pytest.fixture
def mock_workbook():
    with patch("xlsxwriter.Workbook") as MockWB:
        wb_instance = MockWB.return_value
        # Mock add_worksheet to return a mock worksheet
        wb_instance.add_worksheet.return_value = MagicMock()
        # Mock add_format to return a dummy object
        wb_instance.add_format.return_value = MagicMock()
        yield wb_instance

def test_sheet_name_sanitization(mock_workbook):
    report = Report("Test", "Author")
    # Title > 31 chars and invalid chars
    long_title = "A" * 35 + "[:]"
    entry = TableEntry(long_title, [])
    report.add_entry(entry)
    
    report.to_excel("out.xlsx")
    
    # Verify add_worksheet was called with truncated, clean name
    expected_name = "A" * 31
    mock_workbook.add_worksheet.assert_called_with(expected_name)

def test_table_entry_writing(mock_workbook):
    report = Report("Test", "Author")
    headers = ["Col1"]
    data = [["Val1"]]
    report.add_entry(TableEntry("Tab1", data, headers))
    
    report.to_excel("out.xlsx")
    
    ws = mock_workbook.add_worksheet.return_value
    # Verify header write (0, 0, "Col1", <format_object>)
    ws.write.assert_any_call(0, 0, "Col1", ANY)
    # Verify data write
    ws.write.assert_any_call(1, 0, "Val1")

def test_section_entry_writing(mock_workbook):
    report = Report("Test", "Author")
    section = SectionEntry("My Section", [
        TextEntry("Intro", "Hello"),
        TextEntry("Body", "World")
    ])
    report.add_entry(section)
    
    report.to_excel("out.xlsx")
    
    ws = mock_workbook.add_worksheet.return_value
    # It should write to the same sheet multiple times
    # Checking that it wrote the titles and content
    assert ws.write.call_count >= 4  # 2 titles + 2 content blocks

def test_list_entry_rendering(mock_workbook):
    """Test ListEntry rendering to Excel."""
    report = Report("List Test", "Author")
    item = ListItem(0, [RichText("Item 1")])
    entry = ListEntry("My List", [item])
    report.add_entry(entry)

    report.to_excel("out.xlsx")

    ws = mock_workbook.add_worksheet.return_value
    # Should call write_rich_string or write with indentation
    assert ws.write.called or ws.write_rich_string.called

def test_rich_text_formatting(mock_workbook):
    """Test RichText conversion to Excel format."""
    report = Report("Rich Test", "Author")
    rt = RichText("Bold", bold=True, color="#FF0000")
    entry = TextEntry("Rich", [rt])
    report.add_entry(entry)

    report.to_excel("out.xlsx")
    
    # Check that add_format was called with correct props
    mock_workbook.add_format.assert_any_call({'bold': True, 'font_color': '#FF0000'})

def test_power_table_rendering(mock_workbook):
    """Test PowerTableEntry rendering including HTML cell conversion."""
    report = Report("Power Test", "Author")
    
    # Mock DataItem
    mock_item = MagicMock()
    mock_item.printable_values = {"ColA": "<b>Bold</b>"}
    mock_item.default_html_cell_styles = {"ColA": {"color": "red"}} # color name handled by str() fallback or ignored if simple
    
    mock_container = MagicMock()
    mock_container.__iter__.return_value = [mock_item]
    mock_container.repr_cols = ["ColA"]
    
    entry = PowerTableEntry("Power", mock_container)
    report.add_entry(entry)

    report.to_excel("out.xlsx")
    
    ws = mock_workbook.add_worksheet.return_value
    # Should write data row
    assert ws.write.called or ws.write_rich_string.called