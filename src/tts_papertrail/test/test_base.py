import pytest
from tts_papertrail.base import RichText, ListEntry, ListItem, TextEntry

class TestRichText:
    def test_plain_string(self):
        rt = RichText.from_html("Hello World")
        assert len(rt) == 1
        assert rt[0].text == "Hello World"
        assert not rt[0].bold

    def test_bold_tag(self):
        rt = RichText.from_html("<b>Bold</b>")
        assert len(rt) == 1
        assert rt[0].text == "Bold"
        assert rt[0].bold

    def test_style_parsing(self):
        html = '<span style="color: #FF0000; font-weight: bold;">Error</span>'
        rt = RichText.from_html(html)
        assert rt[0].text == "Error"
        # Parser normalizes styles to lowercase
        assert rt[0].color == "#ff0000"
        assert rt[0].bold

    def test_mixed_content(self):
        # "Prefix " (plain) + "Bold" (bold) + " Suffix" (plain)
        html = "Prefix <b>Bold</b> Suffix"
        rt = RichText.from_html(html)
        assert len(rt) == 3
        assert rt[0].text == "Prefix "
        assert rt[1].text == "Bold" and rt[1].bold
        assert rt[2].text == " Suffix"

    def test_whitespace_collapsing(self):
        # Browser behavior: newlines become single space
        html = "   Line1\n   Line2   "
        rt = RichText.from_html(html)
        assert rt[0].text == " Line1 Line2 "

class TestListEntry:
    def test_simple_list(self):
        html = "<ul><li>Item 1</li><li>Item 2</li></ul>"
        entry = ListEntry.from_html("My List", html)
        assert len(entry.items) == 2
        assert entry.items[0].content[0].text == "Item 1"
        assert entry.items[0].level == 0

    def test_nested_list(self):
        # Note: The newline after Parent will be collapsed to a space by the parser
        html = """
        <ul>
            <li>Parent
                <ul>
                    <li>Child</li>
                </ul>
            </li>
        </ul>
        """
        entry = ListEntry.from_html(None, html)
        assert len(entry.items) == 2
        
        # Parent
        assert entry.items[0].level == 0
        # Expect trailing space due to newline in HTML string
        assert entry.items[0].content[0].text == "Parent "
        
        # Child
        assert entry.items[1].level == 1
        assert entry.items[1].content[0].text == "Child"

    def test_rich_list_item(self):
        # Use hex color for explicit styling test
        html = "<ul><li><span style='color:#ff0000'>Alert</span></li></ul>"
        entry = ListEntry.from_html(None, html)
        item = entry.items[0]
        assert item.content[0].text == "Alert"
        assert item.content[0].color == "#ff0000"