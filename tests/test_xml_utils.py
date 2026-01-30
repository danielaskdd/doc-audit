"""
Tests for utils module - XML sanitization functions
"""

import sys
from pathlib import Path

# Add scripts directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent / "skills" / "doc-audit" / "scripts"))

from utils import sanitize_xml_string  # type: ignore


class TestSanitizeXmlString:
    """Tests for sanitize_xml_string function"""

    def test_empty_string(self):
        """Empty string returns empty string"""
        assert sanitize_xml_string("") == ""

    def test_none_returns_none(self):
        """None input returns None"""
        assert sanitize_xml_string(None) is None

    def test_non_string_returns_unchanged(self):
        """Non-string input returns unchanged"""
        assert sanitize_xml_string(123) == 123
        assert sanitize_xml_string([1, 2, 3]) == [1, 2, 3]

    def test_normal_text_unchanged(self):
        """Normal text without control characters passes through"""
        text = "Hello, World! This is normal text."
        assert sanitize_xml_string(text) == text

    def test_preserves_allowed_whitespace(self):
        """Tab, LF, and CR are preserved"""
        text = "Line1\tTabbed\nLine2\rLine3"
        assert sanitize_xml_string(text) == text

    def test_removes_null_character(self):
        """Null character (0x00) is removed"""
        assert sanitize_xml_string("Hello\x00World") == "HelloWorld"

    def test_removes_bell_character(self):
        """Bell character (0x07) is removed"""
        assert sanitize_xml_string("Hello\x07World") == "HelloWorld"

    def test_removes_backspace(self):
        """Backspace (0x08) is removed"""
        assert sanitize_xml_string("Hello\x08World") == "HelloWorld"

    def test_removes_vertical_tab(self):
        """Vertical tab (0x0B) is removed"""
        assert sanitize_xml_string("Hello\x0BWorld") == "HelloWorld"

    def test_removes_form_feed(self):
        """Form feed (0x0C) is removed"""
        assert sanitize_xml_string("Hello\x0CWorld") == "HelloWorld"

    def test_removes_control_chars_0x0E_to_0x1F(self):
        """Control characters 0x0E-0x1F are removed"""
        # Test a few representative ones
        assert sanitize_xml_string("Hello\x0EWorld") == "HelloWorld"
        assert sanitize_xml_string("Hello\x1FWorld") == "HelloWorld"
        assert sanitize_xml_string("Hello\x10World") == "HelloWorld"

    def test_removes_multiple_control_chars(self):
        """Multiple control characters are all removed"""
        text = "A\x00B\x01C\x02D\x07E\x0BF"
        assert sanitize_xml_string(text) == "ABCDEF"

    def test_only_control_chars_returns_empty(self):
        """String with only control characters returns empty string"""
        text = "\x00\x01\x02\x03"
        assert sanitize_xml_string(text) == ""

    def test_unicode_preserved(self):
        """Unicode characters above 0x1F are preserved"""
        text = "Hello 世界 🌍 café"
        assert sanitize_xml_string(text) == text

    def test_mixed_content(self):
        """Mixed valid and invalid characters are handled correctly"""
        text = "Hello\x00\tWorld\x07\n你好\x0B!"
        expected = "Hello\tWorld\n你好!"
        assert sanitize_xml_string(text) == expected

    def test_real_world_llm_output(self):
        """Simulates control chars that might appear in LLM output"""
        # Sometimes LLM output may contain stray control characters
        text = "The violation is:\x00 missing currency specification"
        expected = "The violation is: missing currency specification"
        assert sanitize_xml_string(text) == expected
