#!/usr/bin/env python3
"""
ABOUTME: Tests for mixed body/table content handling in apply_audit_edits.py
ABOUTME: Covers extract_longest_segment() and _process_item() <table> tag detection
"""

import sys
from pathlib import Path
from lxml import etree

# Add scripts directory to path
_scripts_dir = Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'
sys.path.insert(0, str(_scripts_dir))

from _apply_audit_edits_helpers import (  # noqa: E402
    create_mock_applier,
    create_paragraph_xml,
    create_edit_item,
    create_table_with_cells,
    NSMAP,
)
from apply_audit_edits import (  # type: ignore[import-not-found]  # noqa: E402
    extract_longest_segment,
    NS,
)


# ============================================================
# Tests: extract_longest_segment()
# ============================================================

class TestExtractLongestSegment:
    """Tests for the extract_longest_segment helper function."""

    def test_no_table_tags_returns_none(self):
        """Plain text without table tags returns None."""
        assert extract_longest_segment("Normal text without tags") is None

    def test_empty_string_returns_none(self):
        assert extract_longest_segment("") is None

    def test_body_before_table(self):
        """Body text before a table: returns the longer segment."""
        text = '表73　 配套元器件规格型号及厂家信息\n<table>[["序号", "名称"]]</table>'
        result = extract_longest_segment(text)
        # The heading text (16 chars) is longer than JSON (14 chars)
        assert result == '表73　 配套元器件规格型号及厂家信息'

    def test_body_after_table(self):
        """Body text after a table: returns the longer segment."""
        text = '<table>[["A"]]</table>\nThis is a much longer paragraph of body text after the table.'
        result = extract_longest_segment(text)
        assert result == 'This is a much longer paragraph of body text after the table.'

    def test_only_table_tag_with_content(self):
        """Text that is entirely a table block."""
        text = '<table>[["cell1", "cell2"]]</table>'
        result = extract_longest_segment(text)
        assert result == '[["cell1", "cell2"]]'

    def test_empty_segments(self):
        """All segments empty after stripping returns None."""
        text = '<table></table>'
        result = extract_longest_segment(text)
        assert result is None

    def test_multiple_segments(self):
        """Multiple segments: picks the longest."""
        text = 'short\n<table>[["a very long table content here"]]</table>\nmedium text'
        result = extract_longest_segment(text)
        assert result == '[["a very long table content here"]]'

    def test_only_opening_tag(self):
        """Only opening <table> tag (malformed) still triggers splitting."""
        text = 'Heading text\n<table>[["data with longer content"]]'
        result = extract_longest_segment(text)
        assert result == '[["data with longer content"]]'

    def test_only_closing_tag(self):
        """Only closing </table> tag (malformed) still triggers splitting."""
        text = '[["data"]]</table>\nBody text after'
        result = extract_longest_segment(text)
        assert result == 'Body text after'

    def test_heading_longer_than_table(self):
        """When heading text is longer than table content, returns heading."""
        text = 'This is a very long heading about the specification requirements\n<table>[["A"]]</table>'
        result = extract_longest_segment(text)
        assert result == 'This is a very long heading about the specification requirements'


# ============================================================
# Tests: _process_item with mixed body/table content
# ============================================================

class TestProcessItemMixedContent:
    """Tests for _process_item handling of <table> tags in violation_text."""

    def _build_body_with_para_and_table(self):
        """Build a body element with a heading paragraph followed by a table."""
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)

        # Heading paragraph
        heading_para = create_paragraph_xml("表73　 配套元器件规格型号及厂家信息", para_id="AAA")
        body.append(heading_para)

        # Table with one row
        table = create_table_with_cells(
            [["序号", "名称", "型号"]],
            row_para_ids=[["T01", "T02", "T03"]]
        )
        body.append(table)

        # End paragraph (for uuid_end)
        end_para = create_paragraph_xml("End content", para_id="BBB")
        body.append(end_para)

        return body

    def test_delete_with_table_tag_falls_back(self):
        """delete action with <table> in violation_text should produce fallback comment."""
        applier = create_mock_applier()
        applier.body_elem = self._build_body_with_para_and_table()
        applier._init_para_order()

        violation = '表73　 配套元器件规格型号及厂家信息\n<table>[["序号", "名称", "型号"]]</table>'
        item = create_edit_item(
            uuid="AAA", uuid_end="BBB",
            violation_text=violation,
            fix_action="delete",
            revised_text=""
        )

        result = applier._process_item(item)

        assert result.success is True
        assert result.warning is True
        assert "Mixed body/table content" in result.error_message
        # Should have generated a fallback comment
        assert len(applier.comments) == 1
        assert "[FALLBACK]" in applier.comments[0]['text']

    def test_replace_with_table_tag_falls_back(self):
        """replace action with <table> in violation_text should produce fallback comment."""
        applier = create_mock_applier()
        applier.body_elem = self._build_body_with_para_and_table()
        applier._init_para_order()

        violation = '表73　 配套元器件规格型号及厂家信息\n<table>[["序号", "名称", "型号"]]</table>'
        item = create_edit_item(
            uuid="AAA", uuid_end="BBB",
            violation_text=violation,
            fix_action="replace",
            revised_text="fixed text"
        )

        result = applier._process_item(item)

        assert result.success is True
        assert result.warning is True
        assert "Mixed body/table content" in result.error_message

    def test_manual_with_table_tag_extracts_longest_segment(self):
        """manual action should extract longest segment and search for it."""
        applier = create_mock_applier()
        applier.body_elem = self._build_body_with_para_and_table()
        applier._init_para_order()

        # The heading text IS in the document, so if we extract it as the longest
        # segment, the search should find it and successfully place a comment.
        heading = "表73　 配套元器件规格型号及厂家信息"
        violation = f'{heading}\n<table>[["序号", "名称", "型号"]]</table>'
        item = create_edit_item(
            uuid="AAA", uuid_end="BBB",
            violation_text=violation,
            fix_action="manual",
            revised_text="suggestion"
        )

        result = applier._process_item(item)

        assert result.success is True
        # Should have generated a comment (either selected or error—both OK)
        assert len(applier.comments) >= 1

    def test_manual_with_long_table_json_double_bracket(self):
        """manual action where table JSON is longer than heading (double bracket [[ case)."""
        applier = create_mock_applier()
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)

        # Short heading paragraph
        heading_para = create_paragraph_xml("表73　 配套信息", para_id="AAA")
        body.append(heading_para)

        # Table with many columns (JSON representation is longer than heading)
        table = create_table_with_cells(
            [["序号", "名称", "型号", "数量", "质量等级", "工作温度范围", "质量等级引用标准"]],
            row_para_ids=[["T01", "T02", "T03", "T04", "T05", "T06", "T07"]]
        )
        body.append(table)

        end_para = create_paragraph_xml("End", para_id="BBB")
        body.append(end_para)

        applier.body_elem = body
        applier._init_para_order()

        # Table JSON is much longer than heading, so extract_longest_segment picks it
        # This tests the [[ double bracket fix in normalize_table_json and Fallback 3
        violation = '表73　 配套信息\n<table>[["序号", "名称", "型号", "数量", "质量等级", "工作温度范围", "质量等级引用标准"]]'
        item = create_edit_item(
            uuid="AAA", uuid_end="BBB",
            violation_text=violation,
            fix_action="manual",
            revised_text="请核对并补充厂家信息列"
        )

        result = applier._process_item(item)

        assert result.success is True
        assert len(applier.comments) >= 1

    def test_manual_no_table_tag_unchanged(self):
        """manual action without <table> should proceed normally (no segment extraction)."""
        applier = create_mock_applier()
        applier.body_elem = self._build_body_with_para_and_table()
        applier._init_para_order()

        item = create_edit_item(
            uuid="AAA", uuid_end="BBB",
            violation_text="表73　 配套元器件规格型号及厂家信息",
            fix_action="manual",
            revised_text="suggestion"
        )

        result = applier._process_item(item)

        assert result.success is True
        assert len(applier.comments) >= 1

    def test_delete_without_table_tag_not_affected(self):
        """delete action without <table> is not affected by the new logic."""
        applier = create_mock_applier()
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        p = create_paragraph_xml("Remove this text please", para_id="AAA")
        body.append(p)
        applier.body_elem = body
        applier._init_para_order()

        item = create_edit_item(
            uuid="AAA", uuid_end="AAA",
            violation_text="this text",
            fix_action="delete",
            revised_text=""
        )

        result = applier._process_item(item)

        assert result.success is True
        # Should NOT have "Mixed body/table content" in error
        if result.error_message:
            assert "Mixed body/table content" not in result.error_message
