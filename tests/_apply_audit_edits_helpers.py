#!/usr/bin/env python3
"""
ABOUTME: Shared helpers for apply_audit_edits.py tests.
"""

import sys
import importlib
from pathlib import Path
from lxml import etree
from unittest.mock import patch

# Add skills/doc-audit/scripts directory to path (must be before import)
_scripts_dir = Path(__file__).parent.parent / 'skills' / 'doc-audit' / 'scripts'
sys.path.insert(0, str(_scripts_dir))

apply_module = importlib.import_module("apply_audit_edits")
AuditEditApplier = apply_module.AuditEditApplier
NS = apply_module.NS
EditItem = apply_module.EditItem

_common_module = importlib.import_module("docx_edit.common")
DRAWING_PATTERN = _common_module.DRAWING_PATTERN
strip_auto_numbering = _common_module.strip_auto_numbering
strip_markup_tags = _common_module.strip_markup_tags
EditResult = _common_module.EditResult

# ============================================================
# XML Namespace Constants
# ============================================================

NSMAP = {
    'w': NS['w'],
    'w14': NS['w14'],
    'wp': NS['wp'],
}


# ============================================================
# Mock Helper Functions
# ============================================================

def create_paragraph_xml(text_content: str, para_id: str = "12345678") -> etree.Element:
    """
    Create a simple paragraph element with text.

    Args:
        text_content: Text to include in the paragraph
        para_id: w14:paraId attribute value

    Returns:
        lxml Element representing <w:p>
    """
    p = etree.Element(f'{{{NS["w"]}}}p', nsmap=NSMAP)
    p.set(f'{{{NS["w14"]}}}paraId', para_id)

    # Create run with text
    r = etree.SubElement(p, f'{{{NS["w"]}}}r')
    t = etree.SubElement(r, f'{{{NS["w"]}}}t')
    t.text = text_content

    return p


def create_paragraph_with_br(text_content: str, para_id: str = "12345678") -> etree.Element:
    """
    Create a paragraph element with soft line breaks represented by <w:br/>.

    Args:
        text_content: Text with "\n" indicating soft line breaks
        para_id: w14:paraId attribute value

    Returns:
        lxml Element representing <w:p>
    """
    p = etree.Element(f'{{{NS["w"]}}}p', nsmap=NSMAP)
    p.set(f'{{{NS["w14"]}}}paraId', para_id)

    parts = text_content.split('\n')
    r = etree.SubElement(p, f'{{{NS["w"]}}}r')
    for idx, part in enumerate(parts):
        if part:
            t = etree.SubElement(r, f'{{{NS["w"]}}}t')
            t.text = part
        if idx < len(parts) - 1:
            etree.SubElement(r, f'{{{NS["w"]}}}br')

    return p


def create_paragraph_with_inline_image(
    text_before: str,
    img_id: str,
    img_name: str,
    text_after: str = "",
    para_id: str = "12345678"
) -> etree.Element:
    """
    Create a paragraph with text and inline image.

    Args:
        text_before: Text before the image
        img_id: Image id attribute
        img_name: Image name attribute
        text_after: Text after the image
        para_id: w14:paraId attribute value

    Returns:
        lxml Element representing <w:p>
    """
    p = etree.Element(f'{{{NS["w"]}}}p', nsmap=NSMAP)
    p.set(f'{{{NS["w14"]}}}paraId', para_id)

    # Run with text before image
    if text_before:
        r1 = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t1 = etree.SubElement(r1, f'{{{NS["w"]}}}t')
        t1.text = text_before

    # Run with inline image
    r_img = etree.SubElement(p, f'{{{NS["w"]}}}r')
    drawing = etree.SubElement(r_img, f'{{{NS["w"]}}}drawing')
    inline = etree.SubElement(drawing, f'{{{NS["wp"]}}}inline')
    doc_pr = etree.SubElement(inline, f'{{{NS["wp"]}}}docPr')
    doc_pr.set('id', img_id)
    doc_pr.set('name', img_name)

    # Run with text after image
    if text_after:
        r2 = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t2 = etree.SubElement(r2, f'{{{NS["w"]}}}t')
        t2.text = text_after

    return p


def create_paragraph_with_anchor_image(
    text_content: str,
    img_id: str,
    img_name: str,
    para_id: str = "12345678"
) -> etree.Element:
    """
    Create a paragraph with text and floating (anchor) image.
    Anchor images should be extracted as drawing placeholders.

    Args:
        text_content: Text content
        img_id: Image id attribute
        img_name: Image name attribute
        para_id: w14:paraId attribute value

    Returns:
        lxml Element representing <w:p>
    """
    p = etree.Element(f'{{{NS["w"]}}}p', nsmap=NSMAP)
    p.set(f'{{{NS["w14"]}}}paraId', para_id)

    # Run with text
    r1 = etree.SubElement(p, f'{{{NS["w"]}}}r')
    t1 = etree.SubElement(r1, f'{{{NS["w"]}}}t')
    t1.text = text_content

    # Run with floating/anchor image (should be ignored)
    r_img = etree.SubElement(p, f'{{{NS["w"]}}}r')
    drawing = etree.SubElement(r_img, f'{{{NS["w"]}}}drawing')
    anchor = etree.SubElement(drawing, f'{{{NS["wp"]}}}anchor')  # anchor, not inline
    doc_pr = etree.SubElement(anchor, f'{{{NS["wp"]}}}docPr')
    doc_pr.set('id', img_id)
    doc_pr.set('name', img_name)

    return p


def create_paragraph_with_track_changes(
    text_before: str,
    deleted_text: str,
    inserted_text: str,
    text_after: str,
    para_id: str = "12345678"
) -> etree.Element:
    """
    Create a paragraph with track changes (w:del and w:ins).

    Args:
        text_before: Text before changes
        deleted_text: Text marked as deleted
        inserted_text: Text marked as inserted
        text_after: Text after changes
        para_id: w14:paraId attribute value

    Returns:
        lxml Element representing <w:p>
    """
    p = etree.Element(f'{{{NS["w"]}}}p', nsmap=NSMAP)
    p.set(f'{{{NS["w14"]}}}paraId', para_id)

    # Run with text before
    if text_before:
        r1 = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t1 = etree.SubElement(r1, f'{{{NS["w"]}}}t')
        t1.text = text_before

    # Deleted text
    if deleted_text:
        del_elem = etree.SubElement(p, f'{{{NS["w"]}}}del')
        del_elem.set(f'{{{NS["w"]}}}id', "0")
        del_elem.set(f'{{{NS["w"]}}}author', "Test")
        r_del = etree.SubElement(del_elem, f'{{{NS["w"]}}}r')
        del_text = etree.SubElement(r_del, f'{{{NS["w"]}}}delText')
        del_text.text = deleted_text

    # Inserted text
    if inserted_text:
        ins_elem = etree.SubElement(p, f'{{{NS["w"]}}}ins')
        ins_elem.set(f'{{{NS["w"]}}}id', "1")
        ins_elem.set(f'{{{NS["w"]}}}author', "Test")
        r_ins = etree.SubElement(ins_elem, f'{{{NS["w"]}}}r')
        ins_t = etree.SubElement(r_ins, f'{{{NS["w"]}}}t')
        ins_t.text = inserted_text

    # Run with text after
    if text_after:
        r2 = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t2 = etree.SubElement(r2, f'{{{NS["w"]}}}t')
        t2.text = text_after

    return p


def create_mock_applier():
    """
    Create a mock AuditEditApplier for testing internal methods.

    Returns:
        AuditEditApplier instance with mocked dependencies
    """
    with patch.object(AuditEditApplier, '_load_jsonl') as mock_load:
        mock_load.return_value = (
            {'source_file': '/tmp/test.docx', 'source_hash': 'sha256:abc123'},
            []
        )
        applier = AuditEditApplier.__new__(AuditEditApplier)
        applier.meta = {'source_file': '/tmp/test.docx', 'source_hash': 'sha256:abc123'}
        applier.edit_items = []
        applier.verbose = True
        applier.author = 'Test'
        applier.initials = 'Te'
        applier.next_change_id = 0
        applier.next_comment_id = 0
        applier.comments = []
        applier.body_elem = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
        applier._para_list = []
        applier._para_order = {}
        applier._para_id_list = []
        # Set a fixed timestamp for tests
        applier.operation_timestamp = '2025-01-01T00:00:00Z'
        return applier


def create_edit_item(
    uuid: str = "AAA",
    uuid_end: str = "BBB",
    violation_text: str = "Bad text",
    violation_reason: str = "Reason",
    fix_action: str = "replace",
    revised_text: str = "Good text",
    category: str = "semantic",
    rule_id: str = "R001"
) -> EditItem:
    """Create a minimal EditItem for CLI tests."""
    return EditItem(
        uuid=uuid,
        uuid_end=uuid_end,
        violation_text=violation_text,
        violation_reason=violation_reason,
        fix_action=fix_action,
        revised_text=revised_text,
        category=category,
        rule_id=rule_id
    )


def get_test_author(applier, category: str = "semantic") -> str:
    """Build test author using category suffix (matches production logic)."""
    return applier._author_for_item(create_edit_item(category=category))


def create_mock_body_with_paragraphs(para_ids: list) -> etree.Element:
    """Create a mock <w:body> element with paragraphs having paraId."""
    body = etree.Element(f'{{{NS["w"]}}}body', nsmap=NSMAP)
    for para_id in para_ids:
        p = etree.SubElement(body, f'{{{NS["w"]}}}p')
        p.set(f'{{{NS["w14"]}}}paraId', para_id)
        r = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t = etree.SubElement(r, f'{{{NS["w"]}}}t')
        t.text = f"Paragraph {para_id}"
    return body


def create_table_cell_with_paragraphs(para_contents: list, para_ids: list) -> etree.Element:
    """Create a table cell <w:tc> with paragraphs and specified paraIds."""
    tc = etree.Element(f'{{{NS["w"]}}}tc', nsmap=NSMAP)
    for content, para_id in zip(para_contents, para_ids):
        p = etree.SubElement(tc, f'{{{NS["w"]}}}p')
        p.set(f'{{{NS["w14"]}}}paraId', para_id)
        r = etree.SubElement(p, f'{{{NS["w"]}}}r')
        t = etree.SubElement(r, f'{{{NS["w"]}}}t')
        t.text = content
    return tc


def _flatten_para_ids(para_ids: list) -> list:
    """Flatten nested paraId lists into a single list (row-major order)."""
    if not para_ids:
        return []
    flattened = []
    stack = list(para_ids)
    while stack:
        item = stack.pop(0)
        if isinstance(item, list):
            stack = item + stack
        else:
            flattened.append(item)
    return flattened


def create_table_with_cells(cell_contents: list, row_para_ids: list = None) -> etree.Element:
    """
    Create a single-row table with given cell contents and optional paraIds.

    cell_contents: list of cells, where each cell is a list of paragraph strings.
    row_para_ids: optional list (or nested list) of paraIds assigned in order
                  of paragraphs as they are created left-to-right.
    """
    tbl = etree.Element(f'{{{NS["w"]}}}tbl', nsmap=NSMAP)

    tr = etree.SubElement(tbl, f'{{{NS["w"]}}}tr')
    para_iter = iter(_flatten_para_ids(row_para_ids))

    for cell in cell_contents:
        tc = etree.SubElement(tr, f'{{{NS["w"]}}}tc')
        for para_text in cell:
            p = etree.SubElement(tc, f'{{{NS["w"]}}}p')
            try:
                para_id = next(para_iter)
            except StopIteration:
                para_id = None
            if para_id:
                p.set(f'{{{NS["w14"]}}}paraId', para_id)
            r = etree.SubElement(p, f'{{{NS["w"]}}}r')
            t = etree.SubElement(r, f'{{{NS["w"]}}}t')
            t.text = para_text
    return tbl


def create_table_with_cells_with_br(cell_contents: list, row_para_ids: list = None) -> etree.Element:
    """
    Create a single-row table with given cell contents using <w:br/> for soft breaks.

    cell_contents: list of cells, where each cell is a list of paragraph strings.
                   Each paragraph string may include "\n" for soft line breaks.
    row_para_ids: optional list (or nested list) of paraIds assigned in order
                  of paragraphs as they are created left-to-right.
    """
    tbl = etree.Element(f'{{{NS["w"]}}}tbl', nsmap=NSMAP)

    tr = etree.SubElement(tbl, f'{{{NS["w"]}}}tr')
    para_iter = iter(_flatten_para_ids(row_para_ids))

    for cell in cell_contents:
        tc = etree.SubElement(tr, f'{{{NS["w"]}}}tc')
        for para_text in cell:
            try:
                para_id = next(para_iter)
            except StopIteration:
                para_id = None
            if para_id:
                p = create_paragraph_with_br(para_text, para_id=para_id)
            else:
                p = create_paragraph_with_br(para_text)
            tc.append(p)

    return tbl


def create_multi_row_table(rows_data: list, para_ids: list = None) -> etree.Element:
    """
    Create a table with multiple rows.

    rows_data: list of rows, each row is a list of cells, each cell is a list of paragraph strings.
    para_ids: optional list (or nested list) of paraIds assigned sequentially
              as paragraphs are created in row-major order.
    """
    tbl = etree.Element(f'{{{NS["w"]}}}tbl', nsmap=NSMAP)
    para_iter = iter(_flatten_para_ids(para_ids))

    for row_cells in rows_data:
        tr = etree.SubElement(tbl, f'{{{NS["w"]}}}tr')
        for cell in row_cells:
            tc = etree.SubElement(tr, f'{{{NS["w"]}}}tc')
            for para_text in cell:
                p = etree.SubElement(tc, f'{{{NS["w"]}}}p')
                try:
                    para_id = next(para_iter)
                except StopIteration:
                    para_id = None
                if para_id:
                    p.set(f'{{{NS["w14"]}}}paraId', para_id)
                r = etree.SubElement(p, f'{{{NS["w"]}}}r')
                t = etree.SubElement(r, f'{{{NS["w"]}}}t')
                t.text = para_text
    return tbl
