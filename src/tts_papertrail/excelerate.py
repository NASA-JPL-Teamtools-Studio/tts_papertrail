import xlsxwriter
from tts_papertrail.base import Report, TableEntry, PowerTableEntry, SummaryTable, TextEntry, ListEntry, SectionEntry, RichText

class ExcelReport:
    """
    Renders reports to Microsoft Excel (.xlsx) files using `xlsxwriter`.
    
    Idiomatic Features:
    - **Worksheets:** Each TableEntry generally gets its own worksheet.
    - **Header Formatting:** Headers are automatically frozen (freeze_panes) and styled.
    - **Rich Text:** RichText objects are converted to cell formats. Multiple styles 
      within a single cell are supported via `write_rich_string`.
    - **Sanitization:** Enforces Excel's 31-character sheet name limit and removes illegal characters.
    """
    
    def __init__(self, name: str, author: str):
        self.name = name
        self.author = author
        self.created_at = None  # Will be set by Report.to_excel()
        self.entries = []

    def _get_valid_sheet_name(self, title: str) -> str:
        """
        Sanitizes a string to be a valid Excel worksheet name.
        
        Excel rules: Max 31 chars, no chars: [ ] : * ? / \
        """
        if not title: return "Sheet"
        clean = title
        for char in ['/', '\\', '?', '*', ':', '[', ']']: 
            clean = clean.replace(char, '')
        return clean[:31]

    def save(self, filename: str):
        """
        Generates the .xlsx file.
        
        Args:
            filename (str): Output filename. Appends .xlsx if missing.
        """
        if not filename.endswith('.xlsx'): 
            filename += '.xlsx'
        workbook = xlsxwriter.Workbook(filename)
        
        # --- Internal Helpers ---
        def _build_cell_html(value, cell_style):
            """Helper to build HTML string from value and CSS styling."""
            if not value:
                return ''
            
            style_parts = []
            if 'color' in cell_style:
                style_parts.append(f"color:{cell_style['color']}")
            if 'background-color' in cell_style:
                style_parts.append(f"background-color:{cell_style['background-color']}")
            if cell_style.get('font-weight') == 'bold':
                style_parts.append('font-weight:bold')
            if cell_style.get('font-style') == 'italic':
                style_parts.append('font-style:italic')
            
            if style_parts:
                style_attr = '; '.join(style_parts)
                return f'<span style="{style_attr}">{value}</span>'
            return str(value)
        
        format_cache = {}
        def get_format(rt: RichText):
            """Converts RichText attributes to an xlsxwriter format dictionary."""
            props = {}
            if rt.bold: props['bold'] = True
            if rt.italic: props['italic'] = True
            if rt.color: props['font_color'] = rt.color
            if rt.bg_color: props['bg_color'] = rt.bg_color
            
            # Cache formats to prevent hitting Excel's internal format count limit
            key = tuple(sorted(props.items()))
            if key not in format_cache: 
                format_cache[key] = workbook.add_format(props)
            return format_cache[key]

        def write_rich_line(ws, row, col, content_list):
            """Writes a list of mixed Strings/RichText into a single cell."""
            fragments = []
            for item in content_list:
                text = item.text if isinstance(item, RichText) else str(item)
                fmt = get_format(item) if isinstance(item, RichText) else None
                if fmt: fragments.append(fmt)
                fragments.append(text)
            
            if not fragments: return
            # xlsxwriter requires >1 fragment for rich strings, or just 1 string for normal
            if len(fragments) == 1 and isinstance(fragments[0], str): 
                ws.write(row, col, fragments[0])
            else: 
                ws.write_rich_string(row, col, *fragments)

        def render_text(ws, row, entry):
            """Renders a TextEntry to a specific row in the worksheet."""
            if entry.title:
                ws.write(row, 0, entry.title, workbook.add_format({'bold': True, 'underline': True}))
                row += 1
            write_rich_line(ws, row, 0, entry.content)
            return row + 1

        def render_list(ws, row, entry):
            """Renders a ListEntry with each item on its own row."""
            if entry.title:
                ws.write(row, 0, entry.title, workbook.add_format({'bold': True, 'underline': True}))
                row += 1
            for item in entry.items:
                # Simulate visual indentation with spaces
                indent = "    " * item.level + "• "
                
                # Prepend indent to content for rich string writing
                content_with_indent = [indent] + item.content
                
                # Write each list item on its own row
                write_rich_line(ws, row, 0, content_with_indent)
                row += 1
            return row

        # --- Main Rendering Loop ---
        header_fmt = workbook.add_format({'bold': True, 'bottom': 1, 'bg_color': '#D3D3D3'})

        for entry in self.entries:
            # Create a new sheet for the entry (or set of entries)
            sheet_title = entry.title or "Untitled"
            ws = workbook.add_worksheet(self._get_valid_sheet_name(sheet_title))
            curr_row = 0

            if isinstance(entry, SummaryTable):
                # Render SummaryTable with headers
                headers = ['Category', 'Status', 'Message', 'Disposition Summary']
                for c, h in enumerate(headers):
                    ws.write(0, c, str(h), header_fmt)
                ws.freeze_panes(1, 0)
                curr_row = 1
                
                # Write data rows
                for row_data in entry.rows:
                    ws.write(curr_row, 0, row_data['category'])
                    
                    # Handle HTML in status field
                    status = row_data['status']
                    if '<' in status and '>' in status:
                        rich_parts = RichText.from_html(status)
                        if rich_parts and len(rich_parts) == 1:
                            ws.write(curr_row, 1, rich_parts[0].text, get_format(rich_parts[0]))
                        else:
                            ws.write(curr_row, 1, status)
                    else:
                        ws.write(curr_row, 1, status)
                    
                    ws.write(curr_row, 2, row_data['message'])
                    ws.write(curr_row, 3, row_data['disposition_summary'])
                    curr_row += 1
                ws.autofit()
            
            elif isinstance(entry, PowerTableEntry):
                # Extract data from DataContainer with full HTML styling
                container = entry.container.with_cols(entry.columns) if entry.columns else entry.container
                headers = entry.columns if entry.columns else container.repr_cols
                
                # Write headers
                for c, h in enumerate(headers):
                    ws.write(0, c, str(h), header_fmt)
                ws.freeze_panes(1, 0)
                curr_row = 1
                
                # Write data rows
                for item in container:
                    # Get formatted values and cell styles from the DataItem
                    printable_vals = item.printable_values
                    cell_styles = item.default_html_cell_styles
                    
                    for c, col in enumerate(headers):
                        value = printable_vals.get(col, '')
                        cell_style = cell_styles.get(col, {})
                        
                        # Build HTML representation of the cell
                        cell_html = _build_cell_html(value, cell_style)
                        
                        # Parse HTML and apply formatting
                        if '<' in cell_html and '>' in cell_html:
                            rich_parts = RichText.from_html(cell_html)
                            if rich_parts:
                                # Use write_rich_string for multiple formatted parts
                                if len(rich_parts) == 1:
                                    # Single part - simple write
                                    ws.write(curr_row, c, rich_parts[0].text, get_format(rich_parts[0]))
                                else:
                                    # Multiple parts - use rich string
                                    fragments = []
                                    for part in rich_parts:
                                        fragments.append(get_format(part))
                                        fragments.append(part.text)
                                    ws.write_rich_string(curr_row, c, *fragments)
                            else:
                                ws.write(curr_row, c, str(value) if value else '')
                        else:
                            ws.write(curr_row, c, cell_html if cell_html else '')
                    curr_row += 1
                ws.autofit()
            
            elif isinstance(entry, TableEntry):
                # Render Table with Headers
                if entry.headers:
                    for c, h in enumerate(entry.headers): 
                        ws.write(0, c, str(h), header_fmt)
                    ws.freeze_panes(1, 0) # Freeze top row
                    curr_row = 1
                
                for r_data in entry.data:
                    for c, cell in enumerate(r_data):
                        if isinstance(cell, RichText): 
                            ws.write(curr_row, c, cell.text, get_format(cell))
                        else: 
                            ws.write(curr_row, c, cell)
                    curr_row += 1
                ws.autofit()

            elif isinstance(entry, SectionEntry):
                # Stack Section parts vertically on one sheet
                for part in entry.parts:
                    if isinstance(part, TextEntry):
                        curr_row = render_text(ws, curr_row, part)
                    elif isinstance(part, ListEntry):
                        curr_row = render_list(ws, curr_row, part)
                    curr_row += 1 # Visual gap between parts

            elif isinstance(entry, TextEntry):
                render_text(ws, 0, entry)
            
            elif isinstance(entry, ListEntry):
                render_list(ws, 0, entry)

        workbook.close()
        print(f"✅ Excel report saved to {filename}")