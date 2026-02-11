#!/usr/bin/env python3
"""
ABOUTME: Unit tests for apply_audit_edits.py
"""

import sys
import json
from pathlib import Path
import pytest
from lxml import etree
from unittest.mock import patch

from _apply_audit_edits_helpers import (
    apply_module, AuditEditApplier, NS, DRAWING_PATTERN,
    strip_auto_numbering, strip_markup_tags, EditResult, NSMAP,
    create_paragraph_xml, create_paragraph_with_inline_image,
    create_paragraph_with_anchor_image, create_paragraph_with_track_changes,
    create_mock_applier, create_edit_item, get_test_author,
    create_mock_body_with_paragraphs
)

TESTS_DIR = Path(__file__).parent
sys.path.insert(0, str(TESTS_DIR))

class TestCollectRunsInfoOriginal:
    """Tests for _collect_runs_info_original method"""
    
    def test_simple_text(self):
        """Test paragraph with simple text"""
        applier = create_mock_applier()
        para = create_paragraph_xml("Hello World")
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        assert combined_text == "Hello World"
        assert len(runs_info) == 1
        assert runs_info[0]['text'] == "Hello World"
        assert runs_info[0]['start'] == 0
        assert runs_info[0]['end'] == 11
    
    def test_inline_image(self):
        """Test paragraph with inline image"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Before ",
            img_id="1",
            img_name="图片 1",
            text_after=" After"
        )
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        expected_img_str = '<drawing id="1" name="图片 1" />'
        assert expected_img_str in combined_text
        assert combined_text == f'Before {expected_img_str} After'
        
        # Find image run
        img_runs = [r for r in runs_info if r.get('is_drawing')]
        assert len(img_runs) == 1
        assert img_runs[0]['text'] == expected_img_str
    
    def test_anchor_image_supported(self):
        """Test that floating (anchor) images are extracted as drawing placeholders"""
        applier = create_mock_applier()
        para = create_paragraph_with_anchor_image(
            text_content="Text with floating image",
            img_id="1",
            img_name="Float Image"
        )

        runs_info, combined_text = applier._collect_runs_info_original(para)

        expected_img_str = '<drawing id="1" name="Float Image" />'
        assert combined_text == f"Text with floating image{expected_img_str}"

        # One drawing run
        img_runs = [r for r in runs_info if r.get('is_drawing')]
        assert len(img_runs) == 1
        assert img_runs[0]['text'] == expected_img_str
    
    def test_track_changes_deleted_text(self):
        """Test paragraph with deleted text (should be included in original)"""
        applier = create_mock_applier()
        para = create_paragraph_with_track_changes(
            text_before="Hello ",
            deleted_text="old",
            inserted_text="new",
            text_after=" World"
        )
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        # Original text should include deleted text, exclude inserted text
        assert combined_text == "Hello old World"
        assert "new" not in combined_text


# ============================================================
# Tests: _apply_delete
# ============================================================


class TestApplyDelete:
    """Tests for _apply_delete method"""
    
    def test_delete_simple_text(self):
        """Test deleting simple text"""
        applier = create_mock_applier()
        para = create_paragraph_xml("Hello World to delete")

        runs_info, _ = applier._collect_runs_info_original(para)

        result = applier._apply_delete(para, "to delete", "Test reason", runs_info, 12, get_test_author(applier))

        assert result == 'success'
        # Verify w:del element was created
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1
        assert len(applier.comments) == 1
        assert applier.comments[0]['text'] == "Test reason"

    def test_delete_text_not_found(self):
        """Test deleting text that doesn't match position"""
        applier = create_mock_applier()
        para = create_paragraph_xml("Hello World")

        runs_info, _ = applier._collect_runs_info_original(para)

        # Position 50 is beyond text length
        result = applier._apply_delete(para, "missing", "Test reason", runs_info, 50, get_test_author(applier))

        assert result == 'fallback'


# ============================================================
# Tests: _apply_replace with Images
# ============================================================


class TestApplyReplaceWithImages:
    """Tests for _apply_replace method with image handling"""
    
    def test_replace_simple_text(self):
        """Test simple text replacement"""
        applier = create_mock_applier()
        para = create_paragraph_xml("Hello World")

        runs_info, combined_text = applier._collect_runs_info_original(para)

        # Simple text replacement
        violation_text = "Hello"
        revised_text = "Hi"

        match_start = combined_text.find(violation_text)
        result = applier._apply_replace(para, violation_text, revised_text, "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'
        # Verify w:del and w:ins elements were created
        del_elems = para.findall('.//w:del', NSMAP)
        ins_elems = para.findall('.//w:ins', NSMAP)
        assert len(del_elems) == 1
        assert len(ins_elems) == 1
        assert len(applier.comments) == 1
        assert applier.comments[0]['text'] == "Test reason"

    def test_replace_insert_image_fallback(self):
        """Test that inserting images via replace triggers fallback"""
        applier = create_mock_applier()
        para = create_paragraph_xml("Hello World")

        runs_info, combined_text = applier._collect_runs_info_original(para)

        # Try to insert image in revised text
        violation_text = "Hello World"
        revised_text = 'Hello <drawing id="2" name="New Image" /> World'

        result = applier._apply_replace(para, violation_text, revised_text, "Test reason", runs_info, 0, get_test_author(applier))

        assert result == 'fallback'

    def test_replace_delete_image(self):
        """Test that deleting images via replace triggers fallback"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Hello ",
            img_id="1",
            img_name="图片 1",
            text_after=" World"
        )

        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="图片 1" />'

        # Delete image from content
        violation_text = f"Hello {img_str} World"
        revised_text = "Hello World"

        match_start = combined_text.find(violation_text)
        result = applier._apply_replace(para, violation_text, revised_text, "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'fallback'


# ============================================================
# Tests: _apply_replace Equal Portions
# ============================================================


class TestApplyReplaceEqualPortions:
    """Tests for equal portion handling in _apply_replace"""
    
    def test_equal_portion_preserves_image(self):
        """Image in equal portion should be preserved (deepcopy)"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Hello ",
            img_id="1",
            img_name="图片 1",
            text_after=" World"
        )

        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="图片 1" />'

        # Replace "Hello" with "Hi", keep " <img> World" as equal
        violation_text = f"Hello {img_str} World"
        revised_text = f"Hi {img_str} World"

        match_start = combined_text.find(violation_text)
        result = applier._apply_replace(para, violation_text, revised_text, "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'

        # Verify image element still exists
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1

        # Verify inline docPr preserved
        inline = para.find('.//wp:inline', NSMAP)
        assert inline is not None
        doc_pr = inline.find('wp:docPr', NSMAP)
        assert doc_pr is not None
        assert doc_pr.get('id') == '1'
        assert doc_pr.get('name') == '图片 1'

    def test_equal_at_start_with_delete_at_end(self):
        """Equal portion at start, delete at end"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Keep ",
            img_id="1",
            img_name="图片 1",
            text_after=" Delete"
        )

        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="图片 1" />'

        # Delete " Delete" at end, keep "Keep <img>" as equal
        violation_text = f"Keep {img_str} Delete"
        revised_text = f"Keep {img_str}"

        match_start = combined_text.find(violation_text)
        result = applier._apply_replace(para, violation_text, revised_text, "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'

        # Verify image preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1

        # Verify w:del created for " Delete"
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1

    def test_equal_text_before_image_in_equal(self):
        """Equal portion has text before image, both should be preserved"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Delete ",
            img_id="1",
            img_name="图片 1",
            text_after=" End"
        )

        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="图片 1" />'

        # Delete "Delete", keep " <img> End" (space + image + text)
        violation_text = f"Delete {img_str} End"
        revised_text = f"{img_str} End"

        match_start = combined_text.find(violation_text)
        result = applier._apply_replace(para, violation_text, revised_text, "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'

        # Verify image preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1

        # Verify w:del created
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1


def create_paragraph_with_multiple_images(
    parts: list,
    para_id: str = "12345678"
) -> etree.Element:
    """
    Create a paragraph with multiple text/image parts.
    
    Args:
        parts: List of dicts, each with either 'text' or 'img' key
               e.g. [{'text': 'A '}, {'img': ('1', 'Img1')}, {'text': ' B'}]
        para_id: w14:paraId attribute value
    
    Returns:
        lxml Element representing <w:p>
    """
    p = etree.Element(f'{{{NS["w"]}}}p', nsmap=NSMAP)
    p.set(f'{{{NS["w14"]}}}paraId', para_id)
    
    for part in parts:
        if 'text' in part:
            r = etree.SubElement(p, f'{{{NS["w"]}}}r')
            t = etree.SubElement(r, f'{{{NS["w"]}}}t')
            t.text = part['text']
        elif 'img' in part:
            img_id, img_name = part['img']
            r_img = etree.SubElement(p, f'{{{NS["w"]}}}r')
            drawing = etree.SubElement(r_img, f'{{{NS["w"]}}}drawing')
            inline = etree.SubElement(drawing, f'{{{NS["wp"]}}}inline')
            doc_pr = etree.SubElement(inline, f'{{{NS["wp"]}}}docPr')
            doc_pr.set('id', img_id)
            doc_pr.set('name', img_name)
    
    return p



class TestApplyReplaceMultipleImages:
    """Tests for multiple images in equal portions"""

    def test_two_images_both_preserved(self):
        """Two images in equal portion should both be preserved"""
        applier = create_mock_applier()
        para = create_paragraph_with_multiple_images([
            {'text': 'A '},
            {'img': ('1', 'Img1')},
            {'text': ' B '},
            {'img': ('2', 'Img2')},
            {'text': ' C'}
        ])

        runs_info, combined_text = applier._collect_runs_info_original(para)
        img1_str = '<drawing id="1" name="Img1" />'
        img2_str = '<drawing id="2" name="Img2" />'

        # Replace "A" with "X", keep rest
        violation_text = f"A {img1_str} B {img2_str} C"
        revised_text = f"X {img1_str} B {img2_str} C"

        match_start = combined_text.find(violation_text)
        result = applier._apply_replace(para, violation_text, revised_text, "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'

        # Verify both images preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 2


# ============================================================
# Tests: _apply_manual (Comment)
# ============================================================


class TestApplyManual:
    """Tests for _apply_manual method"""
    
    def test_manual_comment_on_text(self):
        """Test adding comment to text"""
        applier = create_mock_applier()
        para = create_paragraph_xml("This is problematic text here")
        
        runs_info, _ = applier._collect_runs_info_original(para)
        
        result = applier._apply_manual(
            para, "problematic text",
            "This text is wrong", "Fix suggestion",
            runs_info, 8,
            get_test_author(applier)
        )
        
        assert result == 'success'
        # Verify comment elements
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1

    def test_manual_comment_with_fallback_reason(self):
        """Test fallback prefix is added when fallback_reason is provided"""
        applier = create_mock_applier()
        para = create_paragraph_xml("This is problematic text here")

        runs_info, _ = applier._collect_runs_info_original(para)

        result = applier._apply_manual(
            para, "problematic text",
            "This text is wrong", "Fix suggestion",
            runs_info, 8,
            get_test_author(applier),
            fallback_reason="Cross-cell track change not supported, fallback to comment"
        )

        assert result == 'success'
        assert len(applier.comments) == 1
        assert applier.comments[0]['text'].startswith(
            "[FALLBACK]Cross-cell"
        )


class TestApplyErrorComment:
    """Tests for _apply_error_comment fallback prefix support."""

    def test_error_comment_with_fallback_reason_prefix(self):
        applier = create_mock_applier()
        para = create_paragraph_xml("Some text")
        item = create_edit_item(
            violation_text="bad text",
            violation_reason="reason",
            revised_text="suggestion",
        )

        ok = applier._apply_error_comment(
            para,
            item,
            fallback_reason="FB_TEST: fallback path",
        )

        assert ok is True
        assert len(applier.comments) == 1
        assert applier.comments[0]["text"].startswith(
            "[FALLBACK]FB_TEST: fallback path\n{WHY}reason"
        )


# ============================================================
# Tests: _apply_delete with Images
# ============================================================


class TestApplyDeleteWithImages:
    """Tests for _apply_delete method with image content"""

    def test_delete_text_before_image(self):
        """Delete text appearing before an inline image, image should be preserved"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Delete this ",
            img_id="1",
            img_name="图片 1",
            text_after=" keep"
        )

        runs_info, combined_text = applier._collect_runs_info_original(para)

        # Delete "Delete this "
        result = applier._apply_delete(para, "Delete this ", "Test reason", runs_info, 0, get_test_author(applier))

        assert result == 'success'
        # Verify w:del element was created
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1
        # Verify image is preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1

    def test_delete_text_after_image(self):
        """Delete text appearing after an inline image, image should be preserved"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Keep ",
            img_id="1",
            img_name="图片 1",
            text_after=" delete this"
        )

        runs_info, combined_text = applier._collect_runs_info_original(para)

        # Find position of " delete this" after the image
        match_start = combined_text.find(" delete this")
        result = applier._apply_delete(para, " delete this", "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'
        # Verify w:del element was created
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1
        # Verify image is preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1

    def test_delete_image_placeholder(self):
        """Delete the image placeholder string itself"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Before ",
            img_id="1",
            img_name="图片 1",
            text_after=" After"
        )

        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="图片 1" />'

        # Delete the image placeholder
        match_start = combined_text.find(img_str)
        result = applier._apply_delete(para, img_str, "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'
        # Verify w:del element was created
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1

    def test_delete_text_with_image_in_middle(self):
        """Delete text in a paragraph where image is in the middle"""
        applier = create_mock_applier()
        para = create_paragraph_with_multiple_images([
            {'text': 'Start delete_me '},
            {'img': ('1', 'Img1')},
            {'text': ' End'}
        ])

        runs_info, combined_text = applier._collect_runs_info_original(para)

        # Delete "delete_me " before the image
        match_start = combined_text.find("delete_me ")
        result = applier._apply_delete(para, "delete_me ", "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'
        # Verify w:del element was created
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1
        # Verify image is still present
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1

    def test_delete_text_between_two_images(self):
        """Delete text between two images"""
        applier = create_mock_applier()
        para = create_paragraph_with_multiple_images([
            {'img': ('1', 'Img1')},
            {'text': ' delete me '},
            {'img': ('2', 'Img2')}
        ])

        runs_info, combined_text = applier._collect_runs_info_original(para)

        # Delete " delete me " between the images
        match_start = combined_text.find(" delete me ")
        result = applier._apply_delete(para, " delete me ", "Test reason", runs_info, match_start, get_test_author(applier))

        assert result == 'success'
        # Verify w:del element was created
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1
        # Verify both images are preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 2


# ============================================================
# Tests: _apply_manual with Images
# ============================================================


class TestApplyManualWithImages:
    """Tests for _apply_manual method with image content"""
    
    def test_manual_comment_on_text_before_image(self):
        """Add comment to text before an inline image"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Mark this text ",
            img_id="1",
            img_name="图片 1",
            text_after=" after"
        )
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        result = applier._apply_manual(
            para, "Mark this",
            "This text needs review", "Suggestion here",
            runs_info, 0,
            get_test_author(applier)
        )
        
        assert result == 'success'
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1
        # Verify image is preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1
        # Verify comment was recorded
        assert len(applier.comments) == 1
    
    def test_manual_comment_on_text_after_image(self):
        """Add comment to text after an inline image"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Before ",
            img_id="1",
            img_name="图片 1",
            text_after=" mark this text"
        )
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        # Find position of "mark this" after the image
        match_start = combined_text.find("mark this")
        result = applier._apply_manual(
            para, "mark this",
            "This text needs fixing", "Fix suggestion",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1
        # Verify image is preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1
    
    def test_manual_comment_on_image_placeholder(self):
        """Add comment directly on the image placeholder string"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Before ",
            img_id="1",
            img_name="图片 1",
            text_after=" After"
        )
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="图片 1" />'
        
        # Add comment to the image placeholder
        match_start = combined_text.find(img_str)
        result = applier._apply_manual(
            para, img_str,
            "This image needs replacement", "Use a better image",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1
        # Verify comment was recorded
        assert len(applier.comments) == 1
        assert "This image needs replacement" in applier.comments[0]['text']
    
    def test_manual_comment_spanning_text_and_image(self):
        """Add comment to text that includes an image placeholder"""
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="See ",
            img_id="1",
            img_name="图片 1",
            text_after=" above"
        )
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="图片 1" />'
        
        # Add comment spanning "See <image> above"
        violation_text = f"See {img_str} above"
        match_start = combined_text.find(violation_text)
        result = applier._apply_manual(
            para, violation_text,
            "This section needs work", "Improve the description",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1
    
    def test_manual_comment_with_multiple_images(self):
        """Add comment in a paragraph with multiple images"""
        applier = create_mock_applier()
        para = create_paragraph_with_multiple_images([
            {'text': 'Text A '},
            {'img': ('1', 'Img1')},
            {'text': ' mark this '},
            {'img': ('2', 'Img2')},
            {'text': ' Text B'}
        ])
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        # Add comment to "mark this" between the images
        match_start = combined_text.find("mark this")
        result = applier._apply_manual(
            para, "mark this",
            "This needs attention", "Suggestion",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1
        # Verify both images are preserved
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 2
        # Verify comment was recorded
        assert len(applier.comments) == 1
    
    def test_manual_comment_on_image_between_text(self):
        """Add comment on image that is between text runs"""
        applier = create_mock_applier()
        para = create_paragraph_with_multiple_images([
            {'text': 'Start '},
            {'img': ('1', 'ProblemImage')},
            {'text': ' End'}
        ])
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="ProblemImage" />'
        
        # Add comment to the image
        match_start = combined_text.find(img_str)
        result = applier._apply_manual(
            para, img_str,
            "Image quality is low", "Replace with high-res version",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        # Verify comment was recorded with correct content
        assert len(applier.comments) == 1
        assert "Image quality is low" in applier.comments[0]['text']
        assert "Replace with high-res version" in applier.comments[0]['text']


# ============================================================
# Tests: DRAWING_PATTERN regex
# ============================================================


class TestDrawingPattern:
    """Tests for DRAWING_PATTERN regex"""
    
    def test_matches_valid_drawing(self):
        """Test pattern matches valid drawing placeholder"""
        valid_cases = [
            '<drawing id="1" name="图片 1" />',
            '<drawing id="123" name="Image" />',
            '<drawing id="0" name="" />',
            '<drawing id="8" name="Image 8" path="sample_blocks.image/image8.emf" format="emf" />',
        ]
        for case in valid_cases:
            assert DRAWING_PATTERN.fullmatch(case), f"Should match: {case}"
    
    def test_no_match_invalid(self):
        """Test pattern doesn't match invalid formats"""
        invalid_cases = [
            '<drawing id="1" name="图片 1">content</drawing>',  # has content
            '<draw id="1" name="test" />',  # wrong tag
            'text <drawing id="1" name="Test" /> more',  # mixed with text
        ]
        for case in invalid_cases:
            assert not DRAWING_PATTERN.fullmatch(case), f"Should not fullmatch: {case}"
    
    def test_search_finds_in_text(self):
        """Test search finds drawing in mixed text"""
        text = 'Hello <drawing id="1" name="图片 1" /> World'
        assert DRAWING_PATTERN.search(text)


# ============================================================
# Tests: Auto-numbering stripping
# ============================================================


class TestStripAutoNumbering:
    """Tests for strip_auto_numbering helper"""
    
    def test_strip_numeric(self):
        """Test stripping numeric prefixes"""
        assert strip_auto_numbering("1. Introduction") == ("Introduction", True)
        assert strip_auto_numbering("1.1 Details") == ("Details", True)
        assert strip_auto_numbering("1) First") == ("First", True)
    
    def test_strip_alphabetic(self):
        """Test stripping alphabetic prefixes"""
        assert strip_auto_numbering("a. First item") == ("First item", True)
        assert strip_auto_numbering("A) Section") == ("Section", True)
    
    def test_strip_bullet(self):
        """Test stripping bullet"""
        assert strip_auto_numbering("• Note") == ("Note", True)
    
    def test_no_strip_normal(self):
        """Test no stripping for normal text"""
        assert strip_auto_numbering("Normal text") == ("Normal text", False)


class TestStripMarkupTags:
    """Tests for strip_markup_tags helper."""

    def test_strip_sup_sub_tags(self):
        stripped, changed = strip_markup_tags("H<sup>2</sup>O")
        assert changed is True
        assert stripped == "H2O"

    def test_strip_mixed_tags(self):
        text = (
            "A<drawing id=\"1\" name=\"img\" />"
            "x<sub>2</sub>"
            "<equation>x^2</equation>"
            "<table>[\"1\"]</table>"
            "B"
        )
        stripped, changed = strip_markup_tags(text)
        assert changed is True
        assert stripped == (
            "A<drawing id=\"1\" name=\"img\" />"
            "x2"
            "<equation>x^2</equation>"
            "<table>[\"1\"]</table>"
            "B"
        )


class TestLocateItemMatchNumbering:
    """Regression tests for numbering stripping during item locate."""

    def test_boundary_body_match_strips_revised_text(self):
        """When violation_text is stripped by mode, revised_text should be stripped too."""
        applier = create_mock_applier()
        para = create_paragraph_xml(
            "三防漆涂覆\n依据文件ADYU0.005.000《三防漆涂覆通用工艺规程》和"
            "ADYU2.088.1192ZD6《信息处理组件三防漆涂覆作业指导书》中要求执行。",
            para_id="AAAA1111",
        )
        item = create_edit_item(
            uuid="AAAA1111",
            uuid_end="AAAA1111",
            fix_action="replace",
            violation_text=(
                "h) 三防漆涂覆\n依据文件ADYU0.005.000《三防漆涂覆通用工艺规程》和"
                "ADYU2.088.1192ZD6《信息处理组件三防漆涂覆作业指导书》中要求执行。"
            ),
            revised_text=(
                "h) 三防漆涂覆\n依据文件ADYU0.005.000《三防漆涂覆通用工艺规程》和"
                "ADYU2.088.1192ZD5《信息处理组件三防漆涂覆作业指导书》中要求执行。"
            ),
        )

        # Force fallback route:
        # single-para original miss -> numbering variant miss -> boundary_crossed ->
        # body segment original miss -> body segment numbering variant hit
        search_results = iter([(-1, None), (-1, None), (-1, None), (0, None)])
        applier._find_in_runs_with_normalization = lambda runs, text: next(search_results)
        applier._iter_paragraphs_in_range = lambda start_para, uuid_end: iter([para])
        applier._collect_runs_info_across_paragraphs = (
            lambda start_para, uuid_end: ([], "", False, "boundary_crossed")
        )
        applier._find_tables_in_range = lambda start_para, uuid_end: []
        applier._is_paragraph_in_table = lambda _para: False
        applier._apply_fallback_comment = lambda *args, **kwargs: None

        context = applier._locate_item_match(
            item=item,
            anchor_para=para,
            violation_text=item.violation_text,
            revised_text=item.revised_text,
        )

        assert context["target_para"] is para
        assert context["matched_start"] == 0
        assert context["numbering_stripped"] is True
        assert context["violation_text"].startswith("三防漆涂覆")
        assert context["revised_text"].startswith("三防漆涂覆")
        assert not context["revised_text"].startswith("h) ")
        assert "ZD5" in context["revised_text"]


class TestNotFoundMarkupRetry:
    """Tests for not-found fallback retry with markup stripping."""

    @staticmethod
    def _setup_body(applier, text: str, para_id: str = "AAA"):
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        para = create_paragraph_xml(text, para_id=para_id)
        body.append(para)
        applier.body_elem = body
        applier._init_para_order()
        return para

    @staticmethod
    def _create_para_with_subscript(para_id: str = "AAA"):
        para = etree.Element(f'{{{NS["w"]}}}p', nsmap=NSMAP)
        para.set(f'{{{NS["w14"]}}}paraId', para_id)

        r1 = etree.SubElement(para, f'{{{NS["w"]}}}r')
        t1 = etree.SubElement(r1, f'{{{NS["w"]}}}t')
        t1.text = "π"

        r2 = etree.SubElement(para, f'{{{NS["w"]}}}r')
        r2pr = etree.SubElement(r2, f'{{{NS["w"]}}}rPr')
        va = etree.SubElement(r2pr, f'{{{NS["w"]}}}vertAlign')
        va.set(f'{{{NS["w"]}}}val', "subscript")
        t2 = etree.SubElement(r2, f'{{{NS["w"]}}}t')
        t2.text = "C"

        r3 = etree.SubElement(para, f'{{{NS["w"]}}}r')
        t3 = etree.SubElement(r3, f'{{{NS["w"]}}}t')
        t3.text = "＝0.65(L为印制板层数）"
        return para

    def test_not_found_c_delete_retries_to_range_comment(self):
        applier = create_mock_applier()
        para = self._setup_body(applier, "H2O sample")

        item = create_edit_item(
            uuid="AAA",
            uuid_end="AAA",
            fix_action="delete",
            violation_text="H<sup>2</sup>O",
            revised_text="",
        )

        # Force C-branch: boundary crossed + no table/body fallback hit.
        applier._collect_runs_info_across_paragraphs = (
            lambda start_para, uuid_end: ([], "", False, "boundary_crossed")
        )
        applier._find_tables_in_range = lambda start_para, uuid_end: []
        applier._is_paragraph_in_table = lambda _para: False

        result = applier._process_item(item)

        assert result.success is True
        assert result.warning is True
        assert result.error_message == "Violation text not found(C)"
        assert len(para.findall('.//w:commentRangeStart', NSMAP)) == 1
        assert len(para.findall('.//w:commentRangeEnd', NSMAP)) == 1
        assert len(para.findall('.//w:del', NSMAP)) == 0
        assert len(para.findall('.//w:ins', NSMAP)) == 0
        assert applier.comments[0]['text'].startswith("[FALLBACK]Violation text not found(C)")

    def test_not_found_s_manual_retries_to_range_comment(self):
        applier = create_mock_applier()
        para = self._setup_body(applier, "H2O sample")

        item = create_edit_item(
            uuid="AAA",
            uuid_end="AAA",
            fix_action="manual",
            violation_text="H<sup>2</sup>O",
            revised_text="Use plain text",
        )

        # Force S-branch: one table segment considered, still not found on first pass.
        fake_table = etree.Element(f'{{{NS["w"]}}}tbl', nsmap=NSMAP)
        applier._find_tables_in_range = lambda start_para, uuid_end: [fake_table]

        result = applier._process_item(item)

        assert result.success is True
        assert result.warning is False
        assert result.error_message is None
        assert len(para.findall('.//w:commentRangeStart', NSMAP)) == 1
        assert len(para.findall('.//w:commentRangeEnd', NSMAP)) == 1
        assert not applier.comments[0]['text'].startswith("[FALLBACK]")

    def test_not_found_m_replace_retries_to_range_comment(self):
        applier = create_mock_applier()
        para = self._setup_body(applier, "H2O sample")

        item = create_edit_item(
            uuid="AAA",
            uuid_end="AAA",
            fix_action="replace",
            violation_text="H<sup>2</sup>O",
            revised_text="Water",
        )

        result = applier._process_item(item)

        assert result.success is True
        assert result.warning is True
        assert result.error_message == "Violation text not found(M)"
        assert len(para.findall('.//w:commentRangeStart', NSMAP)) == 1
        assert len(para.findall('.//w:commentRangeEnd', NSMAP)) == 1
        assert len(para.findall('.//w:del', NSMAP)) == 0
        assert len(para.findall('.//w:ins', NSMAP)) == 0
        assert applier.comments[0]['text'].startswith("[FALLBACK]Violation text not found(M)")

    def test_retry_strips_sup_sub_in_both_violation_and_document(self):
        applier = create_mock_applier()
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        para = self._create_para_with_subscript("AAA")
        body.append(para)
        applier.body_elem = body
        applier._init_para_order()

        item = create_edit_item(
            uuid="AAA",
            uuid_end="AAA",
            fix_action="manual",
            violation_text="πC＝0.65(L为印制板层数）",
            revised_text="建议说明变量定义",
        )

        result = applier._process_item(item)

        assert result.success is True
        assert result.warning is False
        assert result.error_message is None
        assert len(para.findall('.//w:commentRangeStart', NSMAP)) == 1
        assert len(para.findall('.//w:commentRangeEnd', NSMAP)) == 1
        assert not applier.comments[0]['text'].startswith("[FALLBACK]")

    @pytest.mark.parametrize(
        "reason_case,fix_action,violation_text,revised_text,expect_success,expect_warning",
        [
            ("C", "delete", "Missing target", "", True, True),
            ("S", "manual", "Missing target", "Suggestion", True, True),
            ("M", "replace", "Missing target", "Fixed", False, False),
        ],
    )
    def test_not_found_retry_miss_keeps_reference_comment_and_prefix(
        self,
        reason_case,
        fix_action,
        violation_text,
        revised_text,
        expect_success,
        expect_warning,
    ):
        applier = create_mock_applier()
        para = self._setup_body(applier, "H2O sample")

        item = create_edit_item(
            uuid="AAA",
            uuid_end="AAA",
            fix_action=fix_action,
            violation_text=violation_text,
            revised_text=revised_text,
        )

        if reason_case == "C":
            applier._collect_runs_info_across_paragraphs = (
                lambda start_para, uuid_end: ([], "", False, "boundary_crossed")
            )
            applier._find_tables_in_range = lambda start_para, uuid_end: []
            applier._is_paragraph_in_table = lambda _para: False
        elif reason_case == "S":
            fake_table = etree.Element(f'{{{NS["w"]}}}tbl', nsmap=NSMAP)
            applier._find_tables_in_range = lambda start_para, uuid_end: [fake_table]

        result = applier._process_item(item)

        assert result.success is expect_success
        assert result.warning is expect_warning
        assert result.error_message == f"Violation text not found({reason_case})"
        assert len(para.findall('.//w:commentRangeStart', NSMAP)) == 0
        assert len(para.findall('.//w:commentRangeEnd', NSMAP)) == 0
        assert len(para.findall('.//w:commentReference', NSMAP)) == 1
        assert applier.comments[0]['text'].startswith(
            f"[FALLBACK]Violation text not found({reason_case})"
        )


# ============================================================
# Tests: _find_revision_ancestor
# ============================================================


class TestFindRevisionAncestor:
    """Tests for _find_revision_ancestor helper method"""
    
    def test_find_revision_ancestor_normal_run(self):
        """Normal run (not in revision) should return None"""
        applier = create_mock_applier()
        para = create_paragraph_xml("Hello World")
        
        # Get the run element
        run = para.find('.//w:r', NSMAP)
        
        result = applier._find_revision_ancestor(run, para)
        assert result is None
    
    def test_find_revision_ancestor_inside_del(self):
        """Run inside w:del should return the del element"""
        applier = create_mock_applier()
        para = create_paragraph_with_track_changes(
            text_before="Before ",
            deleted_text="deleted",
            inserted_text="",
            text_after=" after"
        )
        
        # Find the w:del element and its contained run
        del_elem = para.find('.//w:del', NSMAP)
        run_in_del = del_elem.find('.//w:r', NSMAP)
        
        result = applier._find_revision_ancestor(run_in_del, para)
        assert result is not None
        assert result.tag == f'{{{NS["w"]}}}del'
        assert result is del_elem
    
    def test_find_revision_ancestor_inside_ins(self):
        """Run inside w:ins should return the ins element"""
        applier = create_mock_applier()
        para = create_paragraph_with_track_changes(
            text_before="Before ",
            deleted_text="",
            inserted_text="inserted",
            text_after=" after"
        )
        
        # Find the w:ins element and its contained run
        ins_elem = para.find('.//w:ins', NSMAP)
        run_in_ins = ins_elem.find('.//w:r', NSMAP)
        
        result = applier._find_revision_ancestor(run_in_ins, para)
        assert result is not None
        assert result.tag == f'{{{NS["w"]}}}ins'
        assert result is ins_elem


# ============================================================
# Tests: _apply_manual with Revision Content (New Tests)
# ============================================================


class TestApplyManualWithRevisions:
    """Tests for _apply_manual method preserving revision structure"""
    
    def test_manual_comment_on_deleted_text(self):
        """
        Add comment to text that was deleted by a previous rule.
        Comment markers should be inserted outside the w:del container.
        """
        applier = create_mock_applier()
        para = create_paragraph_with_track_changes(
            text_before="Before ",
            deleted_text="mark this deleted text",
            inserted_text="replacement",
            text_after=" after"
        )
        
        # Collect original text (includes deleted text, excludes inserted)
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        # Verify original text includes the deleted content
        assert "mark this deleted text" in combined_text
        assert "replacement" not in combined_text
        
        # Find position of "mark this" in the deleted text
        match_start = combined_text.find("mark this deleted text")
        
        result = applier._apply_manual(
            para, "mark this deleted text",
            "This deleted text needs review", "Suggestion",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1
        
        # Verify w:del element is still present and intact
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1
        
        # Verify commentRangeStart is BEFORE w:del (sibling, not inside)
        del_elem = del_elems[0]
        del_idx = list(para).index(del_elem)
        range_start_elem = range_start[0]
        range_start_idx = list(para).index(range_start_elem)
        assert range_start_idx < del_idx, "commentRangeStart should be before w:del"
        
        # Verify comment was recorded
        assert len(applier.comments) == 1
    
    def test_manual_comment_spanning_normal_and_deleted(self):
        """
        Add comment spanning normal text and deleted text.
        Start marker should be in normal run, end marker after w:del container.
        """
        applier = create_mock_applier()
        para = create_paragraph_with_track_changes(
            text_before="mark this normal ",
            deleted_text="and deleted",
            inserted_text="",
            text_after=" after"
        )
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        # Comment spans "this normal and deleted"
        violation_text = "this normal and deleted"
        match_start = combined_text.find(violation_text)
        
        result = applier._apply_manual(
            para, violation_text,
            "This span needs review", "Fix suggestion",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1
        
        # Verify w:del element is still present
        del_elems = para.findall('.//w:del', NSMAP)
        assert len(del_elems) == 1
        
        # Verify comment was recorded
        assert len(applier.comments) == 1
    
    def test_manual_preserves_multiple_runs_in_range(self):
        """
        Add comment to text spanning multiple runs.
        All runs should be preserved, not replaced.
        """
        applier = create_mock_applier()
        # Create paragraph with multiple runs
        para = create_paragraph_with_multiple_images([
            {'text': 'Run1 '},
            {'text': 'Run2 '},
            {'text': 'Run3'}
        ])
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        
        # Count runs before
        runs_before = len(para.findall('.//w:r', NSMAP))
        
        # Comment spans "Run2"
        match_start = combined_text.find("Run2")
        result = applier._apply_manual(
            para, "Run2",
            "Middle run needs attention", "Suggestion",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1
        
        # Verify runs are preserved (may have additional runs for split, but original structure intact)
        runs_after = para.findall('.//w:r', NSMAP)
        # Should have at least the original runs (may have more due to splitting)
        assert len(runs_after) >= runs_before
        
        # Verify comment was recorded
        assert len(applier.comments) == 1
    
    def test_manual_with_image_in_range_preserves_image(self):
        """
        Add comment to text that includes an image.
        Image run should be preserved, not replaced with text.
        """
        applier = create_mock_applier()
        para = create_paragraph_with_inline_image(
            text_before="Before ",
            img_id="1",
            img_name="TestImg",
            text_after=" After"
        )
        
        runs_info, combined_text = applier._collect_runs_info_original(para)
        img_str = '<drawing id="1" name="TestImg" />'
        
        # Verify image is in original text
        assert img_str in combined_text
        
        # Comment spans text including image
        violation_text = f"Before {img_str} After"
        match_start = combined_text.find(violation_text)
        
        result = applier._apply_manual(
            para, violation_text,
            "This section includes image", "Suggestion",
            runs_info, match_start,
            get_test_author(applier)
        )
        
        assert result == 'success'
        
        # Verify image element is still present
        img_elems = para.findall('.//w:drawing', NSMAP)
        assert len(img_elems) == 1
        
        # Verify inline structure preserved
        inline = para.find('.//wp:inline', NSMAP)
        assert inline is not None
        doc_pr = inline.find('wp:docPr', NSMAP)
        assert doc_pr is not None
        assert doc_pr.get('id') == '1'
        assert doc_pr.get('name') == 'TestImg'
        
        # Verify comment elements were created
        range_start = para.findall('.//w:commentRangeStart', NSMAP)
        range_end = para.findall('.//w:commentRangeEnd', NSMAP)
        assert len(range_start) == 1
        assert len(range_end) == 1


# ============================================================
# Tests: _load_jsonl strict uuid_end validation
# ============================================================


class TestLoadJsonlStrictValidation:
    """Tests for _load_jsonl strict uuid_end validation"""
    
    def test_load_jsonl_with_uuid_end_flat_format(self, tmp_path):
        """Test loading flat format JSONL with uuid_end field"""
        jsonl_path = tmp_path / "flat_uuid_end.jsonl"
        with jsonl_path.open(mode='w', encoding='utf-8') as f:
            # Write meta line
            meta = {'type': 'meta', 'source_file': '/tmp/test.docx', 'source_hash': 'sha256:abc123'}
            json.dump(meta, f)
            f.write('\n')
            
            # Write edit item with uuid_end
            item = {
                'uuid': 'AAAAAAAA',
                'uuid_end': 'BBBBBBBB',
                'violation_text': 'bad text',
                'violation_reason': 'wrong',
                'fix_action': 'delete',
                'revised_text': '',
                'category': 'test',
                'rule_id': 'R001'
            }
            json.dump(item, f)
            f.write('\n')
            
            f.flush()
            
            # Create applier with mocked dependencies
            with patch.object(AuditEditApplier, '__init__', lambda x, *args, **kwargs: None):
                applier = AuditEditApplier.__new__(AuditEditApplier)
                applier.jsonl_path = jsonl_path
                
                meta_loaded, items_loaded = applier._load_jsonl()
                
                assert len(items_loaded) == 1
                assert items_loaded[0].uuid == 'AAAAAAAA'
                assert items_loaded[0].uuid_end == 'BBBBBBBB'
    
    def test_load_jsonl_missing_uuid_end_flat_format_raises(self, tmp_path):
        """Test that missing uuid_end in flat format raises ValueError"""
        jsonl_path = tmp_path / "flat_missing_uuid_end.jsonl"
        with jsonl_path.open(mode='w', encoding='utf-8') as f:
            # Write meta line
            meta = {'type': 'meta', 'source_file': '/tmp/test.docx', 'source_hash': 'sha256:abc123'}
            json.dump(meta, f)
            f.write('\n')
            
            # Write edit item WITHOUT uuid_end
            item = {
                'uuid': 'AAAAAAAA',
                # 'uuid_end' is missing
                'violation_text': 'bad text',
                'violation_reason': 'wrong',
                'fix_action': 'delete',
                'revised_text': '',
                'category': 'test',
                'rule_id': 'R001'
            }
            json.dump(item, f)
            f.write('\n')
            
            f.flush()
            
            # Create applier with mocked dependencies
            with patch.object(AuditEditApplier, '__init__', lambda x, *args, **kwargs: None):
                applier = AuditEditApplier.__new__(AuditEditApplier)
                applier.jsonl_path = jsonl_path
                
                with pytest.raises(ValueError, match="Missing 'uuid_end' field"):
                    applier._load_jsonl()
    
    def test_load_jsonl_with_uuid_end_nested_format(self, tmp_path):
        """Test loading nested format JSONL (from run_audit.py) with uuid_end"""
        jsonl_path = tmp_path / "nested_uuid_end.jsonl"
        with jsonl_path.open(mode='w', encoding='utf-8') as f:
            # Write meta line
            meta = {'type': 'meta', 'source_file': '/tmp/test.docx', 'source_hash': 'sha256:abc123'}
            json.dump(meta, f)
            f.write('\n')
            
            # Write paragraph with violations (nested format)
            para = {
                'uuid': 'AAAAAAAA',
                'uuid_end': 'BBBBBBBB',  # Block-level uuid_end
                'p_heading': 'Section 1',
                'p_content': 'Content text',
                'violations': [
                    {
                        'violation_text': 'bad text',
                        'violation_reason': 'wrong',
                        'fix_action': 'replace',
                        'revised_text': 'good text',
                        'category': 'test',
                        'rule_id': 'R001'
                        # uuid_end inherited from paragraph level
                    }
                ]
            }
            json.dump(para, f)
            f.write('\n')
            
            f.flush()
            
            # Create applier with mocked dependencies
            with patch.object(AuditEditApplier, '__init__', lambda x, *args, **kwargs: None):
                applier = AuditEditApplier.__new__(AuditEditApplier)
                applier.jsonl_path = jsonl_path
                
                meta_loaded, items_loaded = applier._load_jsonl()
                
                assert len(items_loaded) == 1
                assert items_loaded[0].uuid == 'AAAAAAAA'
                assert items_loaded[0].uuid_end == 'BBBBBBBB'
                assert items_loaded[0].heading == 'Section 1'
    
    def test_load_jsonl_missing_uuid_end_nested_format_raises(self, tmp_path):
        """Test that missing uuid_end in nested format raises ValueError"""
        jsonl_path = tmp_path / "nested_missing_uuid_end.jsonl"
        with jsonl_path.open(mode='w', encoding='utf-8') as f:
            # Write meta line
            meta = {'type': 'meta', 'source_file': '/tmp/test.docx', 'source_hash': 'sha256:abc123'}
            json.dump(meta, f)
            f.write('\n')
            
            # Write paragraph with violations WITHOUT uuid_end
            para = {
                'uuid': 'AAAAAAAA',
                # 'uuid_end' is missing at both levels
                'p_heading': 'Section 1',
                'p_content': 'Content text',
                'violations': [
                    {
                        'violation_text': 'bad text',
                        'violation_reason': 'wrong',
                        'fix_action': 'replace',
                        'revised_text': 'good text',
                        'category': 'test',
                        'rule_id': 'R001'
                    }
                ]
            }
            json.dump(para, f)
            f.write('\n')
            
            f.flush()
            
            # Create applier with mocked dependencies
            with patch.object(AuditEditApplier, '__init__', lambda x, *args, **kwargs: None):
                applier = AuditEditApplier.__new__(AuditEditApplier)
                applier.jsonl_path = jsonl_path
                
                with pytest.raises(ValueError, match="Missing 'uuid_end' field"):
                    applier._load_jsonl()


# ============================================================
# Tests: _iter_paragraphs_in_range
# ============================================================
class TestIterParagraphsInRange:
    """Tests for _iter_paragraphs_in_range method"""
    
    def test_iter_single_paragraph(self):
        """Test iteration over a single paragraph (uuid == uuid_end)"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB', 'CCC'])
        applier.body_elem = body
        
        # Find start paragraph
        start_para = body.find(f'.//w:p[@w14:paraId="BBB"]', NSMAP)
        
        # Iterate with same uuid_end (single paragraph)
        paras = list(applier._iter_paragraphs_in_range(start_para, 'BBB'))
        
        assert len(paras) == 1
        assert paras[0].get(f'{{{NS["w14"]}}}paraId') == 'BBB'
    
    def test_iter_multiple_paragraphs(self):
        """Test iteration over multiple paragraphs"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB', 'CCC', 'DDD', 'EEE'])
        applier.body_elem = body
        
        # Find start paragraph
        start_para = body.find(f'.//w:p[@w14:paraId="BBB"]', NSMAP)
        
        # Iterate from BBB to DDD
        paras = list(applier._iter_paragraphs_in_range(start_para, 'DDD'))
        
        assert len(paras) == 3
        para_ids = [p.get(f'{{{NS["w14"]}}}paraId') for p in paras]
        assert para_ids == ['BBB', 'CCC', 'DDD']
    
    def test_iter_stops_at_uuid_end(self):
        """Test that iteration stops at uuid_end (inclusive)"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB', 'CCC', 'DDD', 'EEE'])
        applier.body_elem = body
        
        # Find start paragraph
        start_para = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)
        
        # Iterate from AAA to CCC - should not include DDD, EEE
        paras = list(applier._iter_paragraphs_in_range(start_para, 'CCC'))
        
        assert len(paras) == 3
        para_ids = [p.get(f'{{{NS["w14"]}}}paraId') for p in paras]
        assert para_ids == ['AAA', 'BBB', 'CCC']
        assert 'DDD' not in para_ids
        assert 'EEE' not in para_ids
    
    def test_iter_with_nonexistent_uuid_end(self):
        """Test iteration when uuid_end doesn't exist (iterates to end)"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB', 'CCC'])
        applier.body_elem = body
        
        # Find start paragraph
        start_para = body.find(f'.//w:p[@w14:paraId="AAA"]', NSMAP)
        
        # uuid_end doesn't exist - should iterate all remaining paragraphs
        paras = list(applier._iter_paragraphs_in_range(start_para, 'ZZZZZ'))
        
        assert len(paras) == 3
        para_ids = [p.get(f'{{{NS["w14"]}}}paraId') for p in paras]
        assert para_ids == ['AAA', 'BBB', 'CCC']
    
    def test_iter_start_not_in_body(self):
        """Test iteration when start_node is not in body"""
        applier = create_mock_applier()
        body = create_mock_body_with_paragraphs(['AAA', 'BBB', 'CCC'])
        applier.body_elem = body
        
        # Create a detached paragraph (not in body)
        detached_para = create_paragraph_xml("Detached", "XXX")
        
        # Should return empty list
        paras = list(applier._iter_paragraphs_in_range(detached_para, 'BBB'))
        
        assert len(paras) == 0


# ============================================================
# Tests: main exit code behavior
# ============================================================


class TestMainExitCodeBehavior:
    """Tests for main() exit codes on warnings/failures."""
    
    class DummyApplier:
        results = []
        
        def __init__(self, jsonl_file, output_path=None, author=None, initials=None, skip_hash=False, verbose=False):
            self.source_path = Path(jsonl_file)
            self.output_path = Path(output_path) if output_path else Path("out.docx")
            self.edit_items = [1, 2]
        
        def apply(self):
            return self.__class__.results
        
        def save(self, dry_run=False):
            return None
        
        def save_failed_items(self):
            return None
    
    def test_main_returns_zero_with_failures(self, monkeypatch):
        """main() should return 0 even when warnings or failures exist."""
        item = create_edit_item()
        self.DummyApplier.results = [
            EditResult(True, item, "fallback", warning=True),
            EditResult(False, item, "failed"),
        ]
        
        monkeypatch.setattr(apply_module, "AuditEditApplier", self.DummyApplier)
        monkeypatch.setattr(sys, "argv", ["apply_audit_edits.py", "input.jsonl", "--dry-run"])
        
        exit_code = apply_module.main()
        assert exit_code == 0


# ============================================================
# Tests: _collect_runs_info_across_paragraphs
# ============================================================


def _create_body_with_conflict_runs(text_before: str, between: str, text_after: str) -> etree.Element:
    """Create body with one paragraph where both target matches are in revision runs."""
    body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
    p = etree.SubElement(body, f'{{{NS["w"]}}}p')
    p.set(f'{{{NS["w14"]}}}paraId', "AAA")

    r1 = etree.SubElement(p, f'{{{NS["w"]}}}r')
    t1 = etree.SubElement(r1, f'{{{NS["w"]}}}t')
    t1.text = text_before

    del1 = etree.SubElement(p, f'{{{NS["w"]}}}del')
    del1.set(f'{{{NS["w"]}}}id', "1")
    del1.set(f'{{{NS["w"]}}}author', "Test")
    del1_run = etree.SubElement(del1, f'{{{NS["w"]}}}r')
    del1_t = etree.SubElement(del1_run, f'{{{NS["w"]}}}delText')
    del1_t.text = "foo"

    r2 = etree.SubElement(p, f'{{{NS["w"]}}}r')
    t2 = etree.SubElement(r2, f'{{{NS["w"]}}}t')
    t2.text = between

    del2 = etree.SubElement(p, f'{{{NS["w"]}}}del')
    del2.set(f'{{{NS["w"]}}}id', "2")
    del2.set(f'{{{NS["w"]}}}author', "Test")
    del2_run = etree.SubElement(del2, f'{{{NS["w"]}}}r')
    del2_t = etree.SubElement(del2_run, f'{{{NS["w"]}}}delText')
    del2_t.text = "foo"

    r3 = etree.SubElement(p, f'{{{NS["w"]}}}r')
    t3 = etree.SubElement(r3, f'{{{NS["w"]}}}t')
    t3.text = text_after
    return body


class TestProcessItemConflictRetry:
    """Tests for conflict retry behavior in _process_item."""

    def test_replace_conflict_retries_next_match_and_succeeds(self):
        """If first match conflicts, should retry next match and apply edit."""
        applier = create_mock_applier()
        body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        p = etree.SubElement(body, f'{{{NS["w"]}}}p')
        p.set(f'{{{NS["w14"]}}}paraId', "AAA")

        r1 = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t1 = etree.SubElement(r1, f'{{{NS["w"]}}}t')
        t1.text = "prefix "

        del_elem = etree.SubElement(p, f'{{{NS["w"]}}}del')
        del_elem.set(f'{{{NS["w"]}}}id', "1")
        del_elem.set(f'{{{NS["w"]}}}author', "Test")
        del_run = etree.SubElement(del_elem, f'{{{NS["w"]}}}r')
        del_t = etree.SubElement(del_run, f'{{{NS["w"]}}}delText')
        del_t.text = "foo"

        r2 = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t2 = etree.SubElement(r2, f'{{{NS["w"]}}}t')
        t2.text = " middle foo suffix"

        applier.body_elem = body

        item = create_edit_item(
            uuid="AAA",
            uuid_end="AAA",
            violation_text="foo",
            fix_action="replace",
            revised_text="bar",
        )

        result = applier._process_item(item)

        assert result.success is True
        assert result.warning is False
        assert result.error_message is None
        assert not any(
            c["text"].startswith("[FALLBACK]Multiple changes overlap")
            for c in applier.comments
        )

    def test_conflict_fallback_comment_uses_last_matched_text(self):
        """When all matches conflict, fallback comment should use last matched text."""
        applier = create_mock_applier()
        applier.body_elem = _create_body_with_conflict_runs(
            text_before="prefix ",
            between=" middle ",
            text_after=" suffix",
        )

        item = create_edit_item(
            uuid="AAA",
            uuid_end="AAA",
            violation_text=" foo ",
            fix_action="delete",
            revised_text="",
        )

        result = applier._process_item(item)

        assert result.success is True
        assert result.warning is True
        assert result.error_message == "Multiple changes overlap"

        conflict_comments = [
            c for c in applier.comments
            if c["text"].startswith("[FALLBACK]Multiple changes overlap")
        ]
        assert len(conflict_comments) == 1

        where_part = conflict_comments[0]["text"].split("{WHERE}", 1)[1].split("{SUGGEST}", 1)[0]
        assert where_part == "foo "


class TestStatusReasonLatch:
    """Tests for one-shot status reason write/consume behavior."""

    def test_consume_requires_matching_status(self):
        applier = create_mock_applier()
        applier._set_status_reason(
            "cross_cell_fallback",
            "CC_SAMPLE",
            "sample reason",
        )

        # Mismatch should not consume.
        assert applier._consume_status_reason("fallback") is None
        matched = applier._consume_status_reason("cross_cell_fallback")
        assert matched == "CC_SAMPLE: sample reason"

        # One-shot: consumed reason should be cleared.
        assert applier._consume_status_reason("cross_cell_fallback") is None

    def test_set_is_last_write_wins_not_queue(self):
        applier = create_mock_applier()
        applier._set_status_reason("fallback", "FB_A", "first")
        applier._set_status_reason("fallback", "FB_B", "second")
        assert applier._consume_status_reason("fallback") == "FB_B: second"
