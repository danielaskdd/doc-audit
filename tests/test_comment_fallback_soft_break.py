#!/usr/bin/env python3
"""
ABOUTME: Tests fallback comment newline rendering in comments.xml.
ABOUTME: Ensures first newline becomes soft break and trailing newlines are escaped.
"""

import sys
from pathlib import Path

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

TESTS_DIR = Path(__file__).parent
sys.path.insert(0, str(TESTS_DIR))

from _apply_audit_edits_helpers import NS, NSMAP, create_mock_applier  # noqa: E402


def _save_and_get_comment(comment_text: str):
    """Save one comment through _save_comments and return (applier, comment_xml_elem)."""
    applier = create_mock_applier()
    applier.doc = Document()
    applier.comments = [{
        'id': 0,
        'text': comment_text,
        'author': applier.author,
    }]

    applier._save_comments()

    comments_part = applier.doc.part.part_related_by(RT.COMMENTS)
    comments_xml = etree.fromstring(comments_part.blob)
    comment_elems = comments_xml.findall(f'{{{NS["w"]}}}comment')
    assert len(comment_elems) == 1
    return applier, comment_elems[0]


def test_fallback_single_newline_uses_soft_break():
    original = "[FALLBACK]Reason\n{WHY}abc"
    applier, comment_elem = _save_and_get_comment(original)

    br_elems = comment_elem.findall('.//w:br', NSMAP)
    assert len(br_elems) == 1

    text_nodes = [t.text or "" for t in comment_elem.findall('.//w:t', NSMAP)]
    assert text_nodes == ["[FALLBACK]Reason", "{WHY}abc"]

    # Keep in-memory comments unchanged.
    assert applier.comments[0]['text'] == original


def test_fallback_multi_newline_escapes_tail_newlines():
    original = "[FALLBACK]Reason\nLine1\nLine2"
    applier, comment_elem = _save_and_get_comment(original)

    br_elems = comment_elem.findall('.//w:br', NSMAP)
    assert len(br_elems) == 1

    text_nodes = [t.text or "" for t in comment_elem.findall('.//w:t', NSMAP)]
    assert text_nodes == ["[FALLBACK]Reason", "Line1\\nLine2"]
    assert all('\n' not in text for text in text_nodes[1:])

    # Keep in-memory comments unchanged.
    assert applier.comments[0]['text'] == original


def test_non_fallback_newline_does_not_create_soft_break():
    original = "A\nB"
    applier, comment_elem = _save_and_get_comment(original)

    br_elems = comment_elem.findall('.//w:br', NSMAP)
    assert len(br_elems) == 0

    text_nodes = [t.text or "" for t in comment_elem.findall('.//w:t', NSMAP)]
    assert text_nodes == [original]

    # Keep in-memory comments unchanged.
    assert applier.comments[0]['text'] == original
