#!/usr/bin/env python3
"""
ABOUTME: Unit tests for apply_audit_edits.py
"""

import sys
import json
import tempfile
from pathlib import Path

TESTS_DIR = Path(__file__).parent
sys.path.insert(0, str(TESTS_DIR))

import pytest
from lxml import etree
from unittest.mock import patch

import _apply_audit_edits_helpers as helpers
from _apply_audit_edits_helpers import (
    apply_module, AuditEditApplier, NS, DRAWING_PATTERN,
    strip_auto_numbering, EditItem, EditResult, NSMAP,
    create_paragraph_xml, create_paragraph_with_inline_image,
    create_paragraph_with_anchor_image, create_paragraph_with_track_changes,
    create_mock_applier, create_edit_item, get_test_author,
    create_mock_body_with_paragraphs, create_table_cell_with_paragraphs,
    create_table_with_cells, create_multi_row_table,
)


class TestCollectRunsInfoAcrossParagraphs:
    """Tests for _collect_runs_info_across_paragraphs method"""

    def test_single_paragraph_returns_correct_flag(self):
        """Single paragraph should return is_cross_paragraph=False"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA'])
        applier.body_elem = body

        start_para = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(start_para, 'AAA')

        assert boundary_error is None
        assert is_cross_paragraph is False
        assert 'Paragraph AAA' in combined_text

    def test_multiple_paragraphs_returns_correct_flag(self):
        """Multiple paragraphs should return is_cross_paragraph=True"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB', 'CCC'])
        applier.body_elem = body

        start_para = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(start_para, 'CCC')

        assert boundary_error is None
        assert is_cross_paragraph is True
        assert 'Paragraph AAA' in combined_text
        assert 'Paragraph CCC' in combined_text

    def test_paragraph_boundaries_as_newlines(self):
        """Paragraph boundaries should be converted to \\n"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB'])
        applier.body_elem = body

        start_para = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(start_para, 'BBB')

        assert boundary_error is None
        # Should have newline between paragraphs
        assert '\n' in combined_text
        assert combined_text == 'Paragraph AAA\nParagraph BBB'

    def test_runs_contain_para_elem_reference(self):
        """Runs should contain para_elem reference"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB'])
        applier.body_elem = body

        start_para = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(start_para, 'BBB')

        assert boundary_error is None
        # Filter out boundary markers
        real_runs = [r for r in runs_info if not r.get('is_para_boundary', False)]

        # All real runs should have para_elem
        for run in real_runs:
            assert 'para_elem' in run
            assert run['para_elem'] is not None

    def test_para_boundary_markers_flagged(self):
        """Paragraph boundary runs should be flagged with is_para_boundary=True"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB', 'CCC'])
        applier.body_elem = body

        start_para = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(start_para, 'CCC')

        assert boundary_error is None
        # Find boundary markers
        boundary_runs = [r for r in runs_info if r.get('is_para_boundary', False)]

        # Should have 2 boundary markers (AAA->BBB, BBB->CCC)
        assert len(boundary_runs) == 2

        # Boundary markers should all be '\n'
        for run in boundary_runs:
            assert run['text'] == '\n'

    def test_no_boundary_after_last_paragraph(self):
        """No boundary marker should be added after the last paragraph (uuid_end)"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB'])
        applier.body_elem = body

        start_para = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(start_para, 'BBB')

        assert boundary_error is None
        # Should have exactly 1 boundary marker (AAA->BBB)
        boundary_runs = [r for r in runs_info if r.get('is_para_boundary', False)]
        assert len(boundary_runs) == 1


# ============================================================
# Tests: Cross-paragraph search
# ============================================================


class TestCrossParagraphSearch:
    """Tests for cross-paragraph text search in _process_item"""
    
    def test_single_para_match_found_first(self):
        """Single paragraph match should be found first (no cross-para search)"""
        applier = create_mock_applier()
        
        # Create body with multiple paragraphs
        body = create_mock_body_with_paragraphs(['AAA', 'BBB', 'CCC'])
        applier.body_elem = body
        
        # Modify BBB to have specific content
        para_bbb = body.find(f'.//w:p[@w14:paraId="BBB"]', NSMAP)
        para_bbb.find('.//w:t', NSMAP).text = "This is a violation"
        
        # Create edit item
        item = create_edit_item(
            uuid='AAA',
            uuid_end='CCC',
            violation_text='This is a violation',
            fix_action='manual'
        )
        
        # Process item (should find in single paragraph)
        result = applier._process_item(item)
        
        # Should succeed without using cross-paragraph mode
        assert result.success is True


# ============================================================
# Tests: Cross-paragraph manual comments
# ============================================================


class TestApplyManualCrossParagraph:
    """Tests for _apply_manual with cross-paragraph content"""
    
    def test_manual_comment_across_paragraphs(self):
        """Cross-paragraph comment should correctly insert commentRangeStart/End"""
        applier = create_mock_applier()
        
        # Create table cell with multiple paragraphs
        tc = create_table_cell_with_paragraphs(
            ['Line 1', 'Line 2', 'Line 3'],
            ['AAA', 'BBB', 'CCC']
        )
        
        # Create mock body and add cell paragraphs to it
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        for p in tc.findall('.//w:p', NSMAP):
            body.append(p)
        applier.body_elem = body
        
        # Get paragraphs
        para_aaa = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        # Collect runs across paragraphs
        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(para_aaa, 'CCC')

        assert boundary_error is None
        # Verify combined text
        assert combined_text == 'Line 1\nLine 2\nLine 3'
        assert is_cross_paragraph is True

        # Apply manual comment to text spanning all 3 paragraphs
        violation_text = 'Line 1\nLine 2\nLine 3'
        match_start = 0
        
        result = applier._apply_manual(
            para_aaa, violation_text,
            "Cross-paragraph issue", "Fix suggestion",
            runs_info, match_start,
            get_test_author(applier),
            is_cross_paragraph=True
        )
        
        assert result == 'success'
        
        # Verify comment markers were created
        all_range_starts = body.findall('.//w:commentRangeStart', NSMAP)
        all_range_ends = body.findall('.//w:commentRangeEnd', NSMAP)
        
        assert len(all_range_starts) == 1
        assert len(all_range_ends) == 1
        
        # Verify comment was recorded
        assert len(applier.comments) == 1
    
    def test_manual_filters_para_boundary_markers(self):
        """Should filter paragraph boundary markers, only process real runs"""
        applier = create_mock_applier()
        
        # Create paragraphs
        body = create_mock_body_with_paragraphs(['AAA', 'BBB'])
        applier.body_elem = body
        
        para_aaa = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        # Collect runs (will include boundary marker)
        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(para_aaa, 'BBB')

        assert boundary_error is None
        # Verify boundary marker exists
        boundary_runs = [r for r in runs_info if r.get('is_para_boundary', False)]
        assert len(boundary_runs) == 1

        # Apply manual - should filter out boundary marker
        violation_text = combined_text  # Full text including \n
        match_start = 0
        
        result = applier._apply_manual(
            para_aaa, violation_text,
            "Test", "Test",
            runs_info, match_start,
            get_test_author(applier),
            is_cross_paragraph=True
        )
        
        # Should succeed (boundary marker filtered)
        assert result == 'success'


# ============================================================
# Tests: Cross-paragraph fallback for delete/replace
# ============================================================


class TestCrossParagraphFallback:
    """Tests for cross-paragraph fallback behavior in delete/replace"""
    
    def test_delete_with_single_para_content_proceeds(self):
        """Search cross-para but match in single para → normal delete"""
        applier = create_mock_applier()
        
        # Create multiple paragraphs, but violation only in first
        body = create_mock_body_with_paragraphs(['AAA', 'BBB'])
        applier.body_elem = body
        
        # Modify AAA to have violation text
        para_aaa = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)
        para_aaa.find('.//w:t', NSMAP).text = "Delete this text"

        # Collect runs across paragraphs
        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(para_aaa, 'BBB')

        assert boundary_error is None
        assert is_cross_paragraph is True  # Search is cross-para
        
        # But actual match is only in AAA
        violation_text = "Delete this"
        match_start = combined_text.find(violation_text)
        assert match_start == 0
        
        # Find affected runs
        match_end = match_start + len(violation_text)
        affected = applier._find_affected_runs(runs_info, match_start, match_end)
        real_runs = [r for r in affected if not r.get('is_para_boundary', False)]
        
        # Check actual paragraphs spanned
        para_elems = set(r.get('para_elem') for r in real_runs if r.get('para_elem') is not None)
        
        # Should only span 1 paragraph (AAA)
        assert len(para_elems) == 1
        
        # Delete should proceed (not fallback)
        # This simulates the logic in _process_item for delete operation
        if len(para_elems) == 1:
            target_para = real_runs[0].get('para_elem')
            result = applier._apply_delete(
                target_para, violation_text,
                "Test reason",
                runs_info, match_start,
                get_test_author(applier)
            )
            assert result == 'success'

    def test_replace_with_single_para_content_proceeds(self):
        """Search cross-para but match in single para → normal replace"""
        applier = create_mock_applier()

        # Create multiple paragraphs
        body = create_mock_body_with_paragraphs(['AAA', 'BBB'])
        applier.body_elem = body

        # Modify BBB to have violation text
        para_bbb = body.find(f'.//w:p[@w14:paraId="BBB"]', NSMAP)
        para_bbb.find('.//w:t', NSMAP).text = "Bad text here"

        para_aaa = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        # Collect runs across paragraphs
        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(para_aaa, 'BBB')

        assert boundary_error is None
        assert is_cross_paragraph is True

        # Match is only in BBB (after \n)
        violation_text = "Bad text"
        match_start = combined_text.find(violation_text)

        # Find affected runs
        match_end = match_start + len(violation_text)
        affected = applier._find_affected_runs(runs_info, match_start, match_end)
        real_runs = [r for r in affected if not r.get('is_para_boundary', False)]

        # Check actual paragraphs spanned
        para_elems = set(r.get('para_elem') for r in real_runs if r.get('para_elem') is not None)

        # Should only span 1 paragraph (BBB)
        assert len(para_elems) == 1

        # Replace should proceed
        if len(para_elems) == 1:
            target_para = real_runs[0].get('para_elem')
            result = applier._apply_replace(
                target_para, violation_text, "Good text",
                "Test reason",
                runs_info, match_start,
                get_test_author(applier)
            )
            assert result == 'success'
    
    def test_delete_with_cross_para_content_fallbacks(self):
        """Match actually spans multiple paragraphs → fallback to comment"""
        applier = create_mock_applier()
        
        # Create paragraphs
        tc = create_table_cell_with_paragraphs(
            ['First line', 'Second line'],
            ['AAA', 'BBB']
        )
        
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        for p in tc.findall('.//w:p', NSMAP):
            body.append(p)
        applier.body_elem = body
        
        para_aaa = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)

        # Collect runs
        runs_info, combined_text, is_cross_paragraph, boundary_error = applier._collect_runs_info_across_paragraphs(para_aaa, 'BBB')

        assert boundary_error is None
        # Match spans both paragraphs (includes \n)
        violation_text = "First line\nSecond"
        match_start = 0
        match_end = match_start + len(violation_text)
        
        # Find affected runs
        affected = applier._find_affected_runs(runs_info, match_start, match_end)
        real_runs = [r for r in affected if not r.get('is_para_boundary', False)]
        
        # Check actual paragraphs spanned
        para_elems = set(r.get('para_elem') for r in real_runs if r.get('para_elem') is not None)
        
        # Should span 2 paragraphs
        assert len(para_elems) == 2
        
        # This would trigger fallback in actual code
        # We verify the detection logic works correctly

