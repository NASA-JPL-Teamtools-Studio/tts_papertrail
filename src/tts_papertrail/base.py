import abc
import re
from typing import List, Optional, Union, Any
from datetime import datetime, timezone
from bs4 import BeautifulSoup, NavigableString  # pip install beautifulsoup4

class RichText:
    """
    The atomic unit of style for Papertrail reports.
    
    This class represents a span of text with specific formatting attributes 
    (bold, italic, color). It serves as a universal intermediate representation 
    that can be translated into Excel formats, Word runs, PowerPoint runs, or HTML spans.

    Attributes:
        text (str): The actual text content.
        bold (bool): Whether the text is bold.
        italic (bool): Whether the text is italicized.
        color (Optional[str]): Hex color code (e.g., "#FF0000").
        bg_color (Optional[str]): Hex color code for the background (highlighting).
        url (Optional[str]): A hyperlink URL (if applicable).
    """
    
    def __init__(self, text: str, bold: bool = False, italic: bool = False, 
                 color: Optional[str] = None, bg_color: Optional[str] = None,
                 url: Optional[str] = None):
        self.text = text
        self.bold = bold
        self.italic = italic
        self.color = color      
        self.bg_color = bg_color 
        self.url = url

    def __repr__(self):
        return f"<RichText '{self.text}' bold={self.bold} color={self.color}>"

    @staticmethod
    def _parse_soup_element(element) -> List['RichText']:
        """
        Recursively parses a BeautifulSoup element into a list of RichText objects.
        
        Args:
            element: A BeautifulSoup Tag or NavigableString.

        Returns:
            List[RichText]: A flattened list of RichText objects representing the element.
        """
        parts = []
        if isinstance(element, (NavigableString, str)):
            # Collapse whitespace (browser-like behavior) but preserve single spaces
            text = re.sub(r'\s+', ' ', str(element))
            if text: 
                return [RichText(text)]
            return []

        # Extract inline styles (e.g., <span style="font-weight:bold">)
        style = element.get('style', '').lower()
        
        # Determine styles for this level
        is_bold = (element.name in ['b', 'strong'] or 'font-weight:bold' in style or 'font-weight: bold' in style)
        is_italic = (element.name in ['i', 'em'] or 'font-style:italic' in style or 'font-style: italic' in style)
        
        color = None
        if 'color:' in style:
            try:
                # Simplistic parse: "color: #ff0000;" -> "#ff0000"
                val = style.split('color:')[1].split(';')[0].strip()
                if val.startswith('#'): 
                    color = val
            except IndexError: 
                pass

        # Recurse into children
        for child in element.contents:
            child_parts = RichText._parse_soup_element(child)
            # Apply parent styles to children (additive styling)
            for part in child_parts:
                if is_bold: part.bold = True
                if is_italic: part.italic = True
                if color and not part.color: part.color = color
            parts.extend(child_parts)
            
        return parts

    @staticmethod
    def from_html(html_fragment: str) -> List['RichText']:
        """
        Parses a raw HTML string into a list of RichText objects.
        
        This is a "best-effort" parser designed to ingest HTML reports from other tools.
        It supports basic tags (<b>, <i>, <strong>, <em>) and inline CSS for 
        bold, italic, and hex color codes.

        Args:
            html_fragment (str): The HTML string to parse.

        Returns:
            List[RichText]: The sequence of styled text objects.
        """
        soup = BeautifulSoup(html_fragment, 'html.parser')
        return RichText._parse_soup_element(soup)

class Entry(abc.ABC):
    """
    Abstract base class for all report sections.
    
    Attributes:
        title (Optional[str]): The title of this section. Usage varies by renderer
                               (e.g., Worksheet name in Excel, Slide Title in PPT).
    """
    def __init__(self, title: Optional[str] = None):
        self.title = title

class TextEntry(Entry):
    """
    Represents a block of text, which may contain multiple paragraphs or styling.

    Attributes:
        content (List[Union[str, RichText]]): The textual content.
    """
    def __init__(self, title: Optional[str], content: Union[str, RichText, List[Union[str, RichText]]]):
        super().__init__(title)
        if isinstance(content, list): 
            self.content = content
        elif isinstance(content, (str, RichText)): 
            self.content = [content]
        else: 
            raise ValueError("Content must be string, RichText, or list.")

class ListItem:
    """
    Helper class for ListEntry representing a single bullet point.

    Attributes:
        level (int): The indentation level (0-based).
        content (List[RichText]): The text content of the list item.
    """
    def __init__(self, level: int, content: List[RichText]):
        self.level = level
        self.content = content

class ListEntry(Entry):
    """
    Represents a hierarchical (nested) list.

    Attributes:
        items (List[ListItem]): The sequence of list items.
    """
    def __init__(self, title: Optional[str], items: List[ListItem]):
        super().__init__(title)
        self.items = items

    @staticmethod
    def from_html(title: Optional[str], html_string: str) -> 'ListEntry':
        """
        Parses an HTML <ul> or <ol> into a hierarchical ListEntry.
        
        Handles nested lists by recursively calculating the indentation level.
        Preserves formatting (bold, color) within the list items.

        Args:
            title (Optional[str]): The title for the list entry.
            html_string (str): The HTML string containing <ul> or <ol> tags.

        Returns:
            ListEntry: The populated list entry.
        """
        soup = BeautifulSoup(html_string, 'html.parser')
        root_list = soup.find(['ul', 'ol'])
        if not root_list: 
            return ListEntry(title, [])

        parsed_items = []
        def traverse_list(list_node, level):
            # recursive=False ensures we only process direct children <li>
            for li in list_node.find_all('li', recursive=False):
                # Clone <li> to manipulate it (remove nested lists from text extraction)
                li_clone = BeautifulSoup(str(li), 'html.parser').find('li')
                for nested in li_clone.find_all(['ul', 'ol']): 
                    nested.extract()
                
                rich_content = RichText._parse_soup_element(li_clone)
                if rich_content: 
                    parsed_items.append(ListItem(level, rich_content))

                # Recurse into original nested lists
                for nested in li.find_all(['ul', 'ol'], recursive=False):
                    traverse_list(nested, level + 1)

        traverse_list(root_list, 0)
        return ListEntry(title, parsed_items)

class SectionEntry(Entry):
    """
    A container for grouping multiple entries together.
    
    This is useful for keeping related content (e.g., a preamble paragraph and 
    a list) visually grouped.
    - PowerPoint: Renders all parts on a single slide.
    - Excel: Stacks all parts on a single worksheet.
    - Word: Renders parts sequentially under one heading.

    Attributes:
        parts (List[Entry]): The list of child entries (TextEntry, ListEntry, etc.).
    """
    def __init__(self, title: str, parts: List[Entry]):
        super().__init__(title)
        self.parts = parts

class TableEntry(Entry):
    """
    Represents a 2D grid of data.

    Attributes:
        data (List[List[Any]]): 2D list of cell data. Cells can be strings, 
                                numbers, or RichText objects.
        headers (Optional[List[str]]): List of column headers.
    """
    def __init__(self, title: str, data: List[List[Any]], headers: Optional[List[str]] = None):
        super().__init__(title)
        self.data = data
        self.headers = headers

class PowerTableEntry(Entry):
    """
    Represents a table backed by a DataContainer with power_table() support.
    
    This enables interactive HTML tables with filtering, sorting, and rich disposition
    styling when rendered with HypersonicReport. For other formats (Excel, Word, PPT),
    the data is extracted from the container.

    Attributes:
        container (DataContainer): A DataContainer object (e.g., EhaContainer, EvrContainer).
        columns (Optional[List[str]]): List of column names to include. If None, uses all repr_cols.
        add_filters (bool): Enable client-side filtering in HTML output.
        add_sorting (bool): Enable client-side sorting in HTML output.
    """
    def __init__(self, title: str, container, columns: Optional[List[str]] = None, 
                 add_filters: bool = True, add_sorting: bool = True):
        super().__init__(title)
        self.container = container
        self.columns = columns
        self.add_filters = add_filters
        self.add_sorting = add_sorting

class SummaryTable(Entry):
    """
    A summary table for collecting test/check results.
    
    This entry type is designed to accumulate rows of summary information,
    typically used for high-level status reporting.
    
    Attributes:
        rows (List[dict]): The accumulated summary rows, each containing:
            - category (str): The category or type of check
            - status (str): The overall status
            - message (str): A descriptive message
            - disposition_summary (str): Summary of dispositions
    """
    def __init__(self, title: str):
        super().__init__(title)
        self.rows = []
    
    def add_to_summary_table(self, category: str, status: str, message: str, disposition_summary: str):
        """
        Adds a row to the summary table.
        
        Args:
            category (str): The category or type of check
            status (str): The overall status (e.g., "PASS", "FAIL", "WARNING")
            message (str): A descriptive message
            disposition_summary (str): Summary of dispositions
        """
        self.rows.append({
            'category': category,
            'status': status,
            'message': message,
            'disposition_summary': disposition_summary
        })

class Report:
    """
    A unified report builder that can export to multiple formats.

    Attributes:
        name (str): The name of the report.
        author (str): The author's name.
        created_at (datetime): The UTC timestamp of creation.
        entries (List[Entry]): The ordered list of content entries.
    """
    def __init__(self, name: str, author: str):
        self.name = name
        self.author = author
        self.created_at = datetime.now(timezone.utc)
        self.entries: List[Entry] = []

    def add_entry(self, entry: Entry):
        """Adds a content entry to the report."""
        self.entries.append(entry)

    def to_html(self, filename: str):
        """
        Exports the report to HTML format.
        
        Args:
            filename (str): Output filename (will append .html if missing).
        """
        from tts_papertrail.hypersonic import HypersonicReport
        renderer = HypersonicReport(self.name, self.author)
        renderer.entries = self.entries
        renderer.created_at = self.created_at
        renderer.save(filename)
    
    def to_excel(self, filename: str):
        """
        Exports the report to Excel format.
        
        Args:
            filename (str): Output filename (will append .xlsx if missing).
        """
        from tts_papertrail.excelerate import ExcelReport
        renderer = ExcelReport(self.name, self.author)
        renderer.entries = self.entries
        renderer.created_at = self.created_at
        renderer.save(filename)
    
    def to_word(self, filename: str):
        """
        Exports the report to Word format.
        
        Args:
            filename (str): Output filename (will append .docx if missing).
        """
        from tts_papertrail.wordsmith import WordReport
        renderer = WordReport(self.name, self.author)
        renderer.entries = self.entries
        renderer.created_at = self.created_at
        renderer.save(filename)
    
    def to_powerpoint(self, filename: str):
        """
        Exports the report to PowerPoint format.
        
        Args:
            filename (str): Output filename (will append .pptx if missing).
        """
        from tts_papertrail.slidekick import PowerPointReport
        renderer = PowerPointReport(self.name, self.author)
        renderer.entries = self.entries
        renderer.created_at = self.created_at
        renderer.save(filename)