#!/usr/bin/env python3
"""
ABOUTME: Applies audit results to Word documents with track changes and comments
ABOUTME: Reads JSONL export from audit report and modifies the source document
"""

import argparse
import copy
import hashlib
import json
import sys
import re
import difflib
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Generator

from docx import Document
from docx.opc.part import Part
from docx.opc.packuri import PackURI
from lxml import etree

# ============================================================
# Constants
# ============================================================

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
}

COMMENTS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"

# Auto-numbering pattern for detecting and stripping list prefixes
# Matches: "1. ", "1.1 ", "1) ", "1）", "a. ", "A) ", "• ", etc.
AUTO_NUMBERING_PATTERN = re.compile(
    r'^(?:'
    r'\d+(?:[\.\d)）]+)\s+'  # Numeric: 1. 1.1 1) 1）
    r'|'
    r'[a-zA-Z][.)）]\s+'     # Alphabetic: a. A) b）
    r'|'
    r'•\s*'                   # Bullet: • (optional space)
    r')'
)

# Drawing pattern for detecting inline image placeholders
# Matches: <drawing id="1" name="图片 1" />
DRAWING_PATTERN = re.compile(r'<drawing\s+id="[^"]*"\s+name="[^"]*"\s*/>')

# ============================================================
# Data Classes
# ============================================================

@dataclass
class EditItem:
    """Single edit item from JSONL export"""
    uuid: str                    # Start paragraph ID (w14:paraId)
    uuid_end: str                # End paragraph ID (w14:paraId) - required
    violation_text: str          # Text to find
    violation_reason: str        # Reason for violation  
    fix_action: str              # delete | replace | manual
    revised_text: str            # Replacement text or suggestion
    category: str                # Category
    rule_id: str                 # Rule ID
    heading: str = ''            # Violation heading/title

@dataclass
class EditResult:
    """Result of processing an edit item"""
    success: bool
    item: EditItem
    error_message: Optional[str] = None
    warning: bool = False  # Warning flag for expected fallback cases (e.g., manual text not found)

# ============================================================
# Helper Functions
# ============================================================

def sanitize_xml_string(text: str) -> str:
    """
    Remove control characters that are illegal in XML 1.0.

    XML 1.0 allows: #x9 (tab), #xA (LF), #xD (CR), and #x20-#xD7FF, #xE000-#xFFFD, #x10000-#x10FFFF
    This function removes all other control characters (0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F).

    Args:
        text: Text that may contain control characters

    Returns:
        Sanitized text safe for XML
    """
    if not text:
        return text
    # Build a translation table to remove illegal control characters
    # Keep: \t (0x09), \n (0x0A), \r (0x0D)
    # Remove: 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F
    illegal_chars = ''.join(
        chr(c) for c in range(0x20)
        if c not in (0x09, 0x0A, 0x0D)
    )
    return text.translate(str.maketrans('', '', illegal_chars))


def format_text_preview(text: str, max_len: int = 30) -> str:
    """
    Format text for log output: remove newlines and truncate.

    Args:
        text: Text to format
        max_len: Maximum length before truncation

    Returns:
        Clean, truncated text with "..." suffix if truncated
    """
    clean = text.replace('\n', ' ').replace('\r', '').replace('\t', ' ')
    # Collapse multiple spaces
    while '  ' in clean:
        clean = clean.replace('  ', ' ')
    clean = clean.strip()
    if len(clean) > max_len:
        return clean[:max_len] + "..."
    return clean


def strip_auto_numbering(text: str) -> Tuple[str, bool]:
    """
    Strip auto-numbering prefix from text.
    
    Examples:
        "1. Introduction" -> ("Introduction", True)
        "a) First item" -> ("First item", True)
        "• Important note" -> ("Important note", True)
        "Normal text" -> ("Normal text", False)
    
    Returns:
        (stripped_text, was_stripped)
    """
    match = AUTO_NUMBERING_PATTERN.match(text)
    if match:
        return text[match.end():], True
    return text, False

# ============================================================
# Main Class: AuditEditApplier
# ============================================================

class AuditEditApplier:
    """
    Applies audit edit results to Word documents.
    
    Supports three fix_action types:
    - delete: Remove text with track changes
    - replace: Replace text using diff-based minimal changes
    - manual: Add Word comment on the text
    """
    
    def __init__(self, jsonl_path: str, output_path: str = None,
                 skip_hash: bool = False, verbose: bool = False,
                 author: str = 'AI', initials: str = 'AI'):
        self.jsonl_path = Path(jsonl_path)
        self.skip_hash = skip_hash
        self.verbose = verbose
        self.author = author
        self.initials = initials if initials else author[:2] if len(author) >= 2 else author
        # Different author names for revisions vs comments
        self.revision_author = f"{author}-fixed"
        self.comment_author = f"{author}-comment"
        
        # Load JSONL
        self.meta, self.edit_items = self._load_jsonl()
        
        # Determine paths
        self.source_path = Path(self.meta['source_file'])
        self.output_path = Path(output_path) if output_path else \
            self.source_path.with_stem(self.source_path.stem + '_edited')
        
        # Document objects (lazy loaded)
        self.doc: Document = None
        self.body_elem = None
        
        # Comment management
        self.next_comment_id = 0
        self.comments: List[Dict] = []
        
        # Track change ID management
        self.next_change_id = 0
        
        # Results tracking
        self.results: List[EditResult] = []
    
    # ==================== JSONL & Hash ====================
    
    def _load_jsonl(self) -> Tuple[Dict, List[EditItem]]:
        """
        Load JSONL export file.
        
        Supports two formats:
        1. Flat format (from html report export): Each line is a single violation with all fields
        2. Nested format (from run_audit.py output): Each line contains multiple violations per paragraph
        
        STRICT MODE: uuid_end is required for all violations. If missing, raises ValueError.
        """
        meta = {}
        items = []
        
        with open(self.jsonl_path, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line:
                    continue
                data = json.loads(line)
                
                if data.get('type') == 'meta':
                    meta = data
                elif 'violations' in data:
                    # Nested format from run_audit.py
                    # Each paragraph has multiple violations or empty array
                    violations = data.get('violations', [])
                    if not violations:
                        # Skip paragraphs with no violations
                        continue
                    
                    # Extract common fields for this paragraph
                    uuid = data.get('uuid', '')
                    heading = data.get('p_heading', '')  # Map from p_heading
                    
                    # Flatten violations array into separate EditItems
                    for v in violations:
                        # STRICT: uuid_end is required (from violation or paragraph-level)
                        uuid_end = v.get('uuid_end', data.get('uuid_end', ''))
                        if not uuid_end:
                            raise ValueError(
                                f"Missing 'uuid_end' field in violation at line {line_num}.\n"
                                f"Rule: {v.get('rule_id', 'N/A')}\n"
                                f"Text: {v.get('violation_text', '')[:50]}...\n"
                                f"Please re-run the audit with the latest parse_document.py and run_audit.py"
                            )
                        
                        items.append(EditItem(
                            uuid=v.get('uuid', uuid),  # Use violation-level uuid if present, else paragraph
                            uuid_end=uuid_end,
                            violation_text=v.get('violation_text', ''),
                            violation_reason=v.get('violation_reason', ''),
                            fix_action=v.get('fix_action', 'manual'),
                            revised_text=v.get('revised_text', ''),
                            category=v.get('category', ''),
                            rule_id=v.get('rule_id', ''),
                            heading=heading
                        ))
                else:
                    # Flat format (existing format for backward compatibility)
                    # STRICT: uuid_end is required
                    uuid_end = data.get('uuid_end', '')
                    if not uuid_end:
                        raise ValueError(
                            f"Missing 'uuid_end' field in edit item at line {line_num}.\n"
                            f"Rule: {data.get('rule_id', 'N/A')}\n"
                            f"Text: {data.get('violation_text', '')[:50]}...\n"
                            f"Please re-run the audit with the latest parse_document.py and run_audit.py"
                        )
                    
                    items.append(EditItem(
                        uuid=data.get('uuid', ''),
                        uuid_end=uuid_end,
                        violation_text=data.get('violation_text', ''),
                        violation_reason=data.get('violation_reason', ''),
                        fix_action=data.get('fix_action', 'manual'),
                        revised_text=data.get('revised_text', ''),
                        category=data.get('category', ''),
                        rule_id=data.get('rule_id', ''),
                        heading=data.get('heading', '')
                    ))
        
        if not meta:
            raise ValueError("JSONL file missing meta line")
        
        return meta, items
    
    def _verify_hash(self) -> bool:
        """Verify document hash matches expected value"""
        expected_hash = self.meta.get('source_hash', '')
        if not expected_hash:
            raise ValueError(
                "JSONL file missing source_hash field.\n"
                "Cannot verify document integrity. Use --skip-hash to bypass verification."
            )
        
        sha256 = hashlib.sha256()
        with open(self.source_path, 'rb') as f:
            for chunk in iter(lambda: f.read(4096), b""):
                sha256.update(chunk)
        
        actual_hash = f"sha256:{sha256.hexdigest()}"
        return actual_hash == expected_hash
    
    # ==================== Paragraph Location ====================
    
    def _xpath(self, elem, expr: str):
        """
        Execute XPath expression with proper namespace handling.
        
        python-docx's BaseOxmlElement has namespaces pre-registered,
        while pure lxml elements (used in tests) require explicit namespaces.
        This helper method handles both cases.
        
        Args:
            elem: Element to query (BaseOxmlElement or lxml.etree.Element)
            expr: XPath expression using namespace prefixes (e.g., './/w:p')
        
        Returns:
            List of matching elements
        """
        try:
            # First try without explicit namespaces (python-docx BaseOxmlElement)
            return elem.xpath(expr)
        except etree.XPathEvalError:
            # Fallback to explicit namespaces (pure lxml elements in tests)
            return elem.xpath(expr, namespaces=NS)
    
    def _find_para_node_by_id(self, para_id: str):
        """
        Find paragraph by w14:paraId using XPath deep search.
        Handles paragraphs nested in tables.
        """
        xpath_expr = f'.//w:p[@w14:paraId="{para_id}"]'
        nodes = self._xpath(self.body_elem, xpath_expr)
        return nodes[0] if nodes else None
    
    def _iter_paragraphs_following(self, start_node) -> Generator:
        """
        Generator: iterate paragraphs from start_node in document order.
        Handles transitions from table cells to main body and vice versa.
        
        DEPRECATED: Use _iter_paragraphs_in_range() with uuid_end boundary instead.
        """
        all_paras = self._xpath(self.body_elem, './/w:p')
        
        try:
            start_index = all_paras.index(start_node)
            for p in all_paras[start_index:]:
                yield p
        except ValueError:
            return
    
    def _iter_paragraphs_in_range(self, start_node, uuid_end: str) -> Generator:
        """
        Generator: iterate paragraphs from start_node to uuid_end (inclusive).
        
        This restricts the search range to within a specific text block,
        preventing accidental modifications to content in other blocks.
        
        Args:
            start_node: Starting paragraph element (from _find_para_node_by_id)
            uuid_end: End boundary paraId (w14:paraId) - iteration stops after this paragraph
        
        Yields:
            Paragraph elements in document order, from start_node to uuid_end (inclusive)
        """
        all_paras = self._xpath(self.body_elem, './/w:p')
        
        try:
            start_index = all_paras.index(start_node)
            for p in all_paras[start_index:]:
                yield p
                # Stop after reaching the end boundary (inclusive)
                para_id = p.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                if para_id == uuid_end:
                    return
        except ValueError:
            return
    
    # ==================== Run Processing ====================

    def _collect_runs_info_original(self, para_elem) -> Tuple[List[Dict], str]:
        """
        Collect run info representing ORIGINAL text (before track changes).
        This is used for searching violation text in documents that have been edited.
        
        Logic:
        - Include: <w:delText> in <w:del> elements (deleted text was part of original)
        - Include: Normal <w:t>, <w:tab>, <w:br> NOT inside <w:ins> or <w:del> elements
        - Exclude: <w:t> inside <w:ins> elements (inserted text didn't exist in original)
        
        Returns:
            Tuple of (runs_info, combined_text)
            runs_info: [{'text': str, 'start': int, 'end': int}, ...]
            combined_text: Full text string
        """
        runs_info = []
        pos = 0
        
        # Process direct children to handle track changes correctly
        for child in para_elem:
            if child.tag == f'{{{NS["w"]}}}del':
                # Deleted text = part of original document
                for run in child.findall('.//w:r', NS):
                    rPr = run.find('w:rPr', NS)
                    # Look for w:delText, w:tab, w:br elements
                    for elem in run:
                        if elem.tag == f'{{{NS["w"]}}}delText':
                            text = elem.text or ''
                            if text:
                                runs_info.append({
                                    'text': text,
                                    'start': pos,
                                    'end': pos + len(text),
                                    'elem': run,
                                    'rPr': rPr
                                })
                                pos += len(text)
                        elif elem.tag == f'{{{NS["w"]}}}tab':
                            runs_info.append({
                                'text': '\t',
                                'start': pos,
                                'end': pos + 1,
                                'elem': run,
                                'rPr': rPr
                            })
                            pos += 1
                        elif elem.tag == f'{{{NS["w"]}}}br':
                            runs_info.append({
                                'text': '\n',
                                'start': pos,
                                'end': pos + 1,
                                'elem': run,
                                'rPr': rPr
                            })
                            pos += 1
                            
            elif child.tag == f'{{{NS["w"]}}}ins':
                # Inserted text = NOT part of original, skip completely
                pass
                
            elif child.tag == f'{{{NS["w"]}}}r':
                # Normal run (not in revision markup)
                rPr = child.find('w:rPr', NS)
                for elem in child:
                    if elem.tag == f'{{{NS["w"]}}}t':
                        text = elem.text or ''
                        if text:
                            runs_info.append({
                                'text': text,
                                'start': pos,
                                'end': pos + len(text),
                                'elem': child,
                                'rPr': rPr
                            })
                            pos += len(text)
                    elif elem.tag == f'{{{NS["w"]}}}tab':
                        runs_info.append({
                            'text': '\t',
                            'start': pos,
                            'end': pos + 1,
                            'elem': child,
                            'rPr': rPr
                        })
                        pos += 1
                    elif elem.tag == f'{{{NS["w"]}}}br':
                        runs_info.append({
                            'text': '\n',
                            'start': pos,
                            'end': pos + 1,
                            'elem': child,
                            'rPr': rPr
                        })
                        pos += 1
                    elif elem.tag == f'{{{NS["w"]}}}drawing':
                        # Handle inline images (ignore floating/anchor images)
                        inline = elem.find(f'{{{NS["wp"]}}}inline')
                        if inline is not None:
                            doc_pr = inline.find(f'{{{NS["wp"]}}}docPr')
                            if doc_pr is not None:
                                img_id = doc_pr.get('id', '')
                                img_name = doc_pr.get('name', '')
                                img_str = f'<drawing id="{img_id}" name="{img_name}" />'
                                runs_info.append({
                                    'text': img_str,
                                    'start': pos,
                                    'end': pos + len(img_str),
                                    'elem': child,
                                    'rPr': rPr,
                                    'is_drawing': True
                                })
                                pos += len(img_str)
        
        combined_text = ''.join(r['text'] for r in runs_info)
        return runs_info, combined_text
    
    def _get_combined_text(self, runs_info: List[Dict]) -> str:
        """Get combined text from runs"""
        return ''.join(r['text'] for r in runs_info)
    
    def _find_affected_runs(self, runs_info: List[Dict], 
                           match_start: int, match_end: int) -> List[Dict]:
        """Find runs that overlap with the target text range"""
        affected = []
        for info in runs_info:
            if info['end'] > match_start and info['start'] < match_end:
                affected.append(info)
        return affected
    
    def _get_rPr_xml(self, rPr_elem) -> str:
        """Convert rPr element to XML string"""
        if rPr_elem is None:
            return ''
        return etree.tostring(rPr_elem, encoding='unicode')
    
    # ==================== Diff Calculation ====================
    
    def _calculate_diff(self, old_text: str, new_text: str) -> List[Tuple[str, str]]:
        """
        Calculate minimal diff between two texts.
        Returns: [('equal', text), ('delete', text), ('insert', text), ...]
        """
        matcher = difflib.SequenceMatcher(None, old_text, new_text)
        operations = []
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                operations.append(('equal', old_text[i1:i2]))
            elif tag == 'delete':
                operations.append(('delete', old_text[i1:i2]))
            elif tag == 'insert':
                operations.append(('insert', new_text[j1:j2]))
            elif tag == 'replace':
                operations.append(('delete', old_text[i1:i2]))
                operations.append(('insert', new_text[j1:j2]))
        
        return operations
    
    # ==================== ID Management ====================
    
    def _init_comment_id(self):
        """Initialize next comment ID by scanning existing comments"""
        max_id = -1
        
        # Check comments.xml via OPC API
        try:
            from docx.opc.constants import RELATIONSHIP_TYPE as RT
            comments_part = self.doc.part.part_related_by(RT.COMMENTS)
            comments_xml = etree.fromstring(comments_part.blob)
            
            for comment in comments_xml.findall(f'{{{NS["w"]}}}comment'):
                cid = comment.get(f'{{{NS["w"]}}}id')
                if cid:
                    try:
                        max_id = max(max_id, int(cid))
                    except ValueError:
                        pass
        except (KeyError, AttributeError):
            pass
        
        # Also check document.xml for comment references
        for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
            for elem in self.body_elem.iter(f'{{{NS["w"]}}}{tag}'):
                cid = elem.get(f'{{{NS["w"]}}}id')
                if cid:
                    try:
                        max_id = max(max_id, int(cid))
                    except ValueError:
                        pass
        
        self.next_comment_id = max_id + 1
    
    def _init_change_id(self):
        """Initialize next track change ID by scanning existing changes"""
        max_id = -1
        
        for tag in ('ins', 'del'):
            for elem in self.body_elem.iter(f'{{{NS["w"]}}}{tag}'):
                cid = elem.get(f'{{{NS["w"]}}}id')
                if cid:
                    try:
                        max_id = max(max_id, int(cid))
                    except ValueError:
                        pass
        
        self.next_change_id = max_id + 1
    
    def _get_next_change_id(self) -> str:
        """Get next track change ID and increment counter"""
        cid = str(self.next_change_id)
        self.next_change_id += 1
        return cid
    
    # ==================== XML Helpers ====================
    
    def _escape_xml(self, text: str) -> str:
        """Escape XML special characters and remove illegal control characters"""
        text = sanitize_xml_string(text)
        return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))
    
    def _create_run(self, text: str, rPr_xml: str = '') -> etree.Element:
        """Create a new w:r element with text"""
        run_xml = f'<w:r xmlns:w="{NS["w"]}">{rPr_xml}<w:t>{self._escape_xml(text)}</w:t></w:r>'
        return etree.fromstring(run_xml)
    
    def _replace_runs(self, para_elem, affected_runs: List[Dict], 
                     new_elements: List[etree.Element]):
        """Replace affected runs with new elements in the paragraph"""
        if not affected_runs or not new_elements:
            return
        
        first_run = affected_runs[0]['elem']
        parent = first_run.getparent()
        
        # Safety check: if element is no longer in DOM, skip
        if parent is None:
            return
        
        try:
            insert_idx = list(parent).index(first_run)
        except ValueError:
            # Element no longer in parent
            return
        
        # Remove old runs
        for info in affected_runs:
            try:
                if info['elem'].getparent() is parent:
                    parent.remove(info['elem'])
            except ValueError:
                pass  # Already removed
        
        # Insert new elements
        for i, elem in enumerate(new_elements):
            parent.insert(insert_idx + i, elem)
    
    def _check_overlap_with_revisions(self, affected_runs: List[Dict]) -> bool:
        """
        Check if any affected run is inside revision markup (w:del or w:ins).
        
        This indicates the text has been modified by a previous rule,
        and we should fallback to comment annotation.
        
        Returns:
            True if overlap detected (should fallback), False otherwise
        """
        for info in affected_runs:
            elem = info.get('elem')
            if elem is None:
                continue
            parent = elem.getparent()
            if parent is not None and parent.tag in (
                f'{{{NS["w"]}}}del',
                f'{{{NS["w"]}}}ins'
            ):
                return True
        return False
    
    def _find_revision_ancestor(self, elem, para_elem) -> Optional[etree.Element]:
        """
        Find the outermost revision container (w:del or w:ins) for an element.
        
        Walks up the tree from elem to para_elem, looking for revision markup.
        Returns the outermost revision container if found, None otherwise.
        
        Args:
            elem: The element to check (usually a w:r run element)
            para_elem: The paragraph element (stopping point for traversal)
        
        Returns:
            The outermost w:del or w:ins element, or None if not inside revision
        """
        revision_tags = (f'{{{NS["w"]}}}del', f'{{{NS["w"]}}}ins')
        outermost_revision = None
        
        current = elem
        while current is not None and current is not para_elem:
            parent = current.getparent()
            if parent is not None and parent.tag in revision_tags:
                outermost_revision = parent
            current = parent
        
        return outermost_revision
    
    # ==================== Delete Operation ====================

    def _apply_delete(self, para_elem, violation_text: str,
                     orig_runs_info: List[Dict],
                     orig_match_start: int) -> str:
        """
        Apply delete operation with track changes.
        
        Args:
            para_elem: Paragraph element
            violation_text: Text to delete
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
        
        Returns:
            'success': Deletion applied successfully
            'fallback': Should fallback to comment annotation
            'conflict': Text overlaps with previous rule modification
        """
        # Use original text position directly
        match_start = orig_match_start
        match_end = match_start + len(violation_text)
        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)
        
        if not affected:
            return 'fallback'
        
        # Check if text overlaps with previous modifications
        if self._check_overlap_with_revisions(affected):
            return 'conflict'
        
        rPr_xml = self._get_rPr_xml(affected[0]['rPr'])
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        change_id = self._get_next_change_id()
        
        # Calculate split points
        first_run = affected[0]
        last_run = affected[-1]
        before_text = first_run['text'][:match_start - first_run['start']]
        after_text = last_run['text'][match_end - last_run['start']:]
        
        new_elements = []
        
        # Before text (unchanged)
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))
        
        # Deleted text
        del_xml = f'''<w:del xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.revision_author}" w:date="{timestamp}">
            <w:r>{rPr_xml}<w:delText>{self._escape_xml(violation_text)}</w:delText></w:r>
        </w:del>'''
        new_elements.append(etree.fromstring(del_xml))
        
        # After text (unchanged)
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))
        
        self._replace_runs(para_elem, affected, new_elements)
        return 'success'
    
    # ==================== Replace Operation ====================
    
    def _apply_replace(self, para_elem, violation_text: str, 
                      revised_text: str,
                      orig_runs_info: List[Dict],
                      orig_match_start: int) -> str:
        """
        Apply replace operation with diff-based track changes.
        
        Strategy: Build all elements in a single pass, preserving original elements
        for 'equal' portions (to keep formatting, images, etc.) and creating
        track changes for delete/insert portions.
        
        Args:
            para_elem: Paragraph element
            violation_text: Text to replace
            revised_text: New text
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
        
        Returns:
            'success': Replace applied (may be partial)
            'fallback': Should fallback to comment annotation
            'conflict': Modifying text overlaps with previous rule modification
        """
        # Use original text position directly
        match_start = orig_match_start
        match_end = match_start + len(violation_text)
        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)
        
        if not affected:
            return 'fallback'
        
        # Calculate diff
        diff_ops = self._calculate_diff(violation_text, revised_text)
        
        # Check for image handling issues: cannot insert images via revision markup
        for op, text in diff_ops:
            if op == 'insert' and DRAWING_PATTERN.search(text):
                if self.verbose:
                    print(f"  [Fallback] Cannot insert images via revision markup")
                return 'fallback'
        
        # Check for conflicts only on delete operations
        current_pos = match_start
        for op, text in diff_ops:
            if op == 'delete':
                del_end = current_pos + len(text)
                del_affected = self._find_affected_runs(orig_runs_info, current_pos, del_end)
                if self._check_overlap_with_revisions(del_affected):
                    return 'conflict'
                current_pos = del_end
            elif op == 'equal':
                current_pos += len(text)
            # insert doesn't consume original text position
        
        rPr_xml = self._get_rPr_xml(affected[0]['rPr'])
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        # Calculate split points for before/after text
        first_run = affected[0]
        last_run = affected[-1]
        before_text = first_run['text'][:match_start - first_run['start']]
        after_text = last_run['text'][match_end - last_run['start']:]
        
        # Build new elements in a single pass
        new_elements = []
        
        # Before text (unchanged part before the match)
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))
        
        # Process diff operations
        violation_pos = 0  # Position within violation_text
        
        for op, text in diff_ops:
            if op == 'equal':
                # For equal portions, try to preserve original elements (especially images)
                equal_start = match_start + violation_pos
                equal_end = equal_start + len(text)
                equal_runs = self._find_affected_runs(orig_runs_info, equal_start, equal_end)
                
                if equal_runs:
                    # Check if this is a single image run that matches exactly
                    if (len(equal_runs) == 1 and 
                        equal_runs[0].get('is_drawing') and
                        equal_runs[0]['start'] == equal_start and
                        equal_runs[0]['end'] == equal_end):
                        # Copy original image element
                        new_elements.append(copy.deepcopy(equal_runs[0]['elem']))
                    else:
                        # Extract text from the equal portion
                        # Handle partial runs at boundaries
                        for i, eq_run in enumerate(equal_runs):
                            run_text = eq_run['text']
                            
                            # Calculate the portion of this run that's in our range
                            run_start_in_range = max(0, equal_start - eq_run['start'])
                            run_end_in_range = min(len(run_text), equal_end - eq_run['start'])
                            portion = run_text[run_start_in_range:run_end_in_range]
                            
                            if eq_run.get('is_drawing'):
                                # Image run - copy entire element if fully contained
                                if run_start_in_range == 0 and run_end_in_range == len(run_text):
                                    new_elements.append(copy.deepcopy(eq_run['elem']))
                            elif portion:
                                new_elements.append(self._create_run(portion, rPr_xml))
                else:
                    # No runs found, create text directly
                    new_elements.append(self._create_run(text, rPr_xml))
                
                violation_pos += len(text)
                
            elif op == 'delete':
                change_id = self._get_next_change_id()
                del_xml = f'''<w:del xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.revision_author}" w:date="{timestamp}">
                    <w:r>{rPr_xml}<w:delText>{self._escape_xml(text)}</w:delText></w:r>
                </w:del>'''
                new_elements.append(etree.fromstring(del_xml))
                violation_pos += len(text)
                
            elif op == 'insert':
                change_id = self._get_next_change_id()
                ins_xml = f'''<w:ins xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.revision_author}" w:date="{timestamp}">
                    <w:r>{rPr_xml}<w:t>{self._escape_xml(text)}</w:t></w:r>
                </w:ins>'''
                new_elements.append(etree.fromstring(ins_xml))
                # insert doesn't consume violation_pos
        
        # After text (unchanged part after the match)
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))
        
        # Single DOM operation to replace all affected runs
        self._replace_runs(para_elem, affected, new_elements)
        return 'success'
    
    # ==================== Manual (Comment) Operation ====================
    
    def _apply_error_comment(self, para_elem, item: EditItem, author_override: str = None) -> bool:
        """
        Insert an unselected comment at the end of paragraph for failed items.
        
        Comment format: {WHY}<violation_reason>  {WHERE}<violation_text>{SUGGEST}<revised_text>
        Author: author_override if provided, otherwise {self.author}-notfound
        
        Args:
            para_elem: Paragraph element to attach comment
            item: EditItem with violation details
            author_override: Optional author name to use instead of default
        """
        comment_id = self.next_comment_id
        self.next_comment_id += 1
        
        # Insert only commentReference at the end of paragraph (no range selection)
        ref_xml = f'''<w:r xmlns:w="{NS['w']}">
            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            <w:commentReference w:id="{comment_id}"/>
        </w:r>'''
        para_elem.append(etree.fromstring(ref_xml))
        
        # Record comment content with custom format and author
        comment_text = f"{{WHY}}{item.violation_reason}  {{WHERE}}{item.violation_text}{{SUGGEST}}{item.revised_text}"
        comment_author = author_override if author_override else f"{self.author}-notfound"
        self.comments.append({
            'id': comment_id,
            'text': comment_text,
            'author': comment_author
        })
        
        return True
    
    def _apply_fallback_comment(self, para_elem, item: EditItem, reason: str = "") -> bool:
        """
        Insert a fallback comment when delete/replace/manual operation cannot be applied.
        
        This is different from _apply_error_comment:
        - Error comment: Unexpected failure
        - Fallback comment: Expected fallback (text was modified by previous rule)
        
        Args:
            para_elem: Paragraph element
            item: Edit item
            reason: Reason for fallback (e.g., "Text modified by previous rule")
        
        Returns:
            True (always succeeds)
        """
        comment_id = self.next_comment_id
        self.next_comment_id += 1
        
        # Insert commentReference at end of paragraph
        ref_xml = f'''<w:r xmlns:w="{NS['w']}">
            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            <w:commentReference w:id="{comment_id}"/>
        </w:r>'''
        para_elem.append(etree.fromstring(ref_xml))
        
        # Format: [FALLBACK] reason | {WHY} ... {WHERE} ... {SUGGEST} ...
        comment_text = f"[FALLBACK] {reason}\n{{WHY}}{item.violation_reason}  {{WHERE}}{item.violation_text}{{SUGGEST}}{item.revised_text}"
        
        self.comments.append({
            'id': comment_id,
            'text': comment_text,
            'author': f"{self.author}-conflict"
        })
        
        return True
    
    def _apply_manual(self, para_elem, violation_text: str,
                     violation_reason: str, revised_text: str,
                     orig_runs_info: List[Dict],
                     orig_match_start: int) -> str:
        """
        Apply manual operation by adding a Word comment.
        
        This method preserves internal run structure (including images, revision 
        markup, etc.) by inserting comment markers at element boundaries rather
        than replacing the content.
        
        Strategy:
        1. Find the precise start/end positions in the original text
        2. For each boundary (start/end), check if the run is inside revision markup
        3. If inside revision: insert marker outside the revision container
        4. If not inside revision: split run if needed, insert marker at precise position
        
        Args:
            para_elem: Paragraph element
            violation_text: Text to mark with comment
            violation_reason: Reason to show in comment
            revised_text: Suggestion to show in comment
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
        
        Returns:
            'success': Comment added successfully
            'fallback': Should fallback to error comment
        """
        # Use original text position directly
        match_start = orig_match_start
        match_end = match_start + len(violation_text)
        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)
        
        if not affected:
            return 'fallback'
        
        comment_id = self.next_comment_id
        self.next_comment_id += 1
        
        first_run_info = affected[0]
        last_run_info = affected[-1]
        first_run = first_run_info['elem']
        last_run = last_run_info['elem']
        rPr_xml = self._get_rPr_xml(first_run_info['rPr'])
        
        # Check if start/end runs are inside revision markup
        start_revision = self._find_revision_ancestor(first_run, para_elem)
        end_revision = self._find_revision_ancestor(last_run, para_elem)
        
        # Calculate text split points
        before_text = first_run_info['text'][:match_start - first_run_info['start']]
        after_text = last_run_info['text'][match_end - last_run_info['start']:]
        
        # Create comment markers
        range_start = etree.fromstring(
            f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        )
        range_end = etree.fromstring(
            f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        )
        comment_ref = etree.fromstring(f'''<w:r xmlns:w="{NS['w']}">
            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            <w:commentReference w:id="{comment_id}"/>
        </w:r>''')
        
        # === Handle START position ===
        if start_revision is not None:
            # Start is inside revision: insert commentRangeStart before revision container
            parent = start_revision.getparent()
            if parent is not None:
                idx = list(parent).index(start_revision)
                parent.insert(idx, range_start)
        else:
            # Start is in normal run: may need to split
            parent = first_run.getparent()
            if parent is None:
                return 'fallback'
            
            try:
                idx = list(parent).index(first_run)
            except ValueError:
                return 'fallback'
            
            if before_text:
                # Need to split: create run for before_text, insert before first_run
                before_run = self._create_run(before_text, rPr_xml)
                parent.insert(idx, before_run)
                idx += 1
                
                # Update first_run's text content (remove before_text portion)
                t_elem = first_run.find('w:t', NS)
                if t_elem is not None and t_elem.text:
                    t_elem.text = t_elem.text[len(before_text):]
            
            # Insert commentRangeStart before the (possibly modified) first_run
            parent.insert(idx, range_start)
        
        # === Handle END position ===
        if end_revision is not None:
            # End is inside revision: insert commentRangeEnd after revision container
            parent = end_revision.getparent()
            if parent is not None:
                idx = list(parent).index(end_revision)
                parent.insert(idx + 1, range_end)
                parent.insert(idx + 2, comment_ref)
        else:
            # End is in normal run: may need to split
            parent = last_run.getparent()
            if parent is None:
                return 'fallback'
            
            try:
                idx = list(parent).index(last_run)
            except ValueError:
                return 'fallback'
            
            if after_text:
                # Need to split: update last_run to remove after_text, create new run for after_text
                t_elem = last_run.find('w:t', NS)
                if t_elem is not None and t_elem.text:
                    original_text = t_elem.text
                    # Keep only the portion up to match_end
                    keep_len = len(original_text) - len(after_text)
                    t_elem.text = original_text[:keep_len]
                
                # Insert commentRangeEnd after last_run
                parent.insert(idx + 1, range_end)
                # Insert commentReference after range_end
                parent.insert(idx + 2, comment_ref)
                # Create after_run and insert after comment_ref
                after_run = self._create_run(after_text, rPr_xml)
                parent.insert(idx + 3, after_run)
            else:
                # No split needed: insert commentRangeEnd and reference after last_run
                parent.insert(idx + 1, range_end)
                parent.insert(idx + 2, comment_ref)
        
        # Record comment content
        comment_text = violation_reason
        if revised_text:
            comment_text += f"\nSuggestion: {revised_text}"
        
        self.comments.append({
            'id': comment_id,
            'text': comment_text
        })
        
        return 'success'
    
    # ==================== Comment Saving ====================
    
    def _save_comments(self):
        """Save comments to comments.xml using OPC API"""
        if not self.comments:
            return
        
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        # Try to get existing comments.xml
        try:
            comments_part = self.doc.part.part_related_by(RT.COMMENTS)
            comments_xml = etree.fromstring(comments_part.blob)
        except (KeyError, AttributeError):
            # Create new comments.xml
            comments_xml = etree.fromstring(
                f'<w:comments xmlns:w="{NS["w"]}" xmlns:w14="{NS["w14"]}"/>'
            )
        
        # Add comments
        for comment in self.comments:
            comment_elem = etree.SubElement(
                comments_xml, f'{{{NS["w"]}}}comment'
            )
            comment_elem.set(f'{{{NS["w"]}}}id', str(comment['id']))
            # Support independent author for each comment (e.g., AI-notfound, AI-conflict)
            # Default author is comment_author (author-comment suffix)
            comment_author = comment.get('author', self.comment_author)
            comment_elem.set(f'{{{NS["w"]}}}author', comment_author)
            comment_elem.set(f'{{{NS["w"]}}}date', timestamp)
            # Use self.initials for all comments with author prefix matching self.author
            # This includes: AI-comment, AI-notfound, AI-conflict (all share same initials)
            if comment_author.startswith(self.author):
                comment_initials = self.initials
            else:
                comment_initials = comment_author[:2] if len(comment_author) >= 2 else comment_author
            comment_elem.set(f'{{{NS["w"]}}}initials', comment_initials)

            # Add paragraph with text
            p = etree.SubElement(comment_elem, f'{{{NS["w"]}}}p')
            r = etree.SubElement(p, f'{{{NS["w"]}}}r')
            t = etree.SubElement(r, f'{{{NS["w"]}}}t')
            t.text = sanitize_xml_string(comment['text'])
        
        # Save via OPC
        blob = etree.tostring(
            comments_xml, xml_declaration=True, encoding='UTF-8'
        )
        
        try:
            comments_part = self.doc.part.part_related_by(RT.COMMENTS)
            comments_part._blob = blob
        except (KeyError, AttributeError):
            # Create new Part
            comments_part = Part(
                PackURI('/word/comments.xml'),
                COMMENTS_CONTENT_TYPE,
                blob,
                self.doc.part.package
            )
            self.doc.part.relate_to(comments_part, RT.COMMENTS)
    
    # ==================== Main Processing ====================
    
    def _process_item(self, item: EditItem) -> EditResult:
        """Process a single edit item"""
        anchor_para = None
        try:
            # Strip leading/trailing whitespace from search and replacement text
            # to prevent matching failures caused by whitespace in JSONL data
            violation_text = item.violation_text.strip()
            revised_text = item.revised_text.strip()
            
            # 1. Find anchor paragraph by ID
            anchor_para = self._find_para_node_by_id(item.uuid)
            if anchor_para is None:
                return EditResult(False, item,
                    f"Paragraph ID {item.uuid} not found (may be in header/footer or ID changed)")
            
            # 2. Search for text from anchor paragraph using ORIGINAL text (before revisions)
            # Store match results to pass to apply methods (avoid double matching)
            # IMPORTANT: Search is restricted to uuid -> uuid_end range to prevent
            # accidental modifications to content in other text blocks
            target_para = None
            matched_runs_info = None
            matched_start = -1
            numbering_stripped = False
            
            # Try original text first (using revision-free view)
            for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                runs_info_orig, combined_orig = self._collect_runs_info_original(para)
                
                pos = combined_orig.find(violation_text)
                if pos != -1:
                    target_para = para
                    matched_runs_info = runs_info_orig
                    matched_start = pos
                    break
            
            # Fallback: Try stripping auto-numbering if original match failed
            if target_para is None:
                stripped_violation, was_stripped = strip_auto_numbering(violation_text)
                
                if was_stripped:
                    if self.verbose:
                        print(f"  [Fallback] Trying without numbering: {stripped_violation[:30]}...")
                    
                    for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                        runs_info_orig, combined_orig = self._collect_runs_info_original(para)
                        
                        pos = combined_orig.find(stripped_violation)
                        if pos != -1:
                            target_para = para
                            matched_runs_info = runs_info_orig
                            matched_start = pos
                            numbering_stripped = True
                            violation_text = stripped_violation
                            
                            # Handle revised_text for replace operation
                            if item.fix_action == 'replace':
                                stripped_revised, revised_has_numbering = strip_auto_numbering(revised_text)
                                
                                if revised_has_numbering:
                                    # Both have numbering: strip both
                                    revised_text = stripped_revised
                                    if self.verbose:
                                        print(f"  [Fallback] Stripped numbering from both texts")
                                # else: Only violation_text had numbering, keep revised_text as-is
                            
                            break
            
            if target_para is None:
                # For manual fix_action, text not found is expected (not an error)
                # Use AI-comment author instead of AI-notfound, mark as warning
                if item.fix_action == 'manual':
                    self._apply_error_comment(anchor_para, item, author_override=self.comment_author)
                    return EditResult(
                        success=True,
                        item=item,
                        error_message="Target missing, comment on heading instead: ",
                        warning=True
                    )
                else:
                    # For delete/replace, text not found is an error
                    self._apply_error_comment(anchor_para, item)
                    return EditResult(False, item,
                        f"Text not found after anchor: ")
            
            # 3. Apply operation based on fix_action
            # Pass matched_runs_info and matched_start to avoid double matching
            success_status = None
            
            if item.fix_action == 'delete':
                success_status = self._apply_delete(
                    target_para, violation_text,
                    matched_runs_info, matched_start
                )
            elif item.fix_action == 'replace':
                success_status = self._apply_replace(
                    target_para, violation_text, revised_text,
                    matched_runs_info, matched_start
                )
            elif item.fix_action == 'manual':
                success_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start
                )
            else:
                # Insert error comment for unknown action type
                self._apply_error_comment(anchor_para, item)
                return EditResult(False, item, f"Unknown action type: {item.fix_action}")
            
            # Handle results
            if success_status == 'success':
                if numbering_stripped and self.verbose:
                    print(f"  [Success] Matched after stripping auto-numbering")
                return EditResult(True, item)
            elif success_status == 'conflict':
                # Text overlaps with previous rule modification - mark as warning
                reason = "Text overlaps with previous rule modification"
                self._apply_fallback_comment(target_para, item, reason)
                if self.verbose:
                    print(f"  [Conflict] {reason}")
                return EditResult(
                    success=True,
                    item=item,
                    error_message=reason,
                    warning=True
                )
            elif success_status == 'fallback':
                # Fallback to comment annotation - mark as warning for all fix_actions
                # For manual fix_action, use AI-comment author (expected behavior)
                # For delete/replace, use AI-conflict author
                reason = "Text not found in current document state"
                if item.fix_action == 'manual':
                    self._apply_error_comment(target_para, item, author_override=self.comment_author)
                else:
                    self._apply_fallback_comment(target_para, item, reason)
                if self.verbose:
                    print(f"  [Fallback] {reason}")
                return EditResult(
                    success=True,
                    item=item,
                    error_message=reason,
                    warning=True
                )
            else:
                # Unexpected return value or old boolean False
                self._apply_error_comment(anchor_para, item)
                return EditResult(False, item, "Operation failed")
                
        except Exception as e:
            # Insert error comment on exception if anchor paragraph exists
            if anchor_para is not None:
                self._apply_error_comment(anchor_para, item)
            return EditResult(False, item, str(e))
    
    def apply(self) -> List[EditResult]:
        """Execute all edit operations"""
        # 1. Verify hash
        if not self.skip_hash:
            if not self._verify_hash():
                raise ValueError(
                    f"Document hash mismatch\n"
                    f"Expected: {self.meta.get('source_hash', 'N/A')}\n"
                    f"Document may have been modified. Use --skip-hash to bypass."
                )
        
        # 2. Load document
        self.doc = Document(str(self.source_path))
        self.body_elem = self.doc._element.body
        
        # 3. Initialize IDs
        self._init_comment_id()
        self._init_change_id()
        
        # 4. Process each item
        for i, item in enumerate(self.edit_items):
            if self.verbose:
                print(f"[{i+1}/{len(self.edit_items)}] {item.fix_action}: "
                      f"{item.violation_text[:40]}...")
            
            result = self._process_item(item)
            self.results.append(result)
            
            if self.verbose:
                status = "✓" if result.success else "✗"
                print(f"  [{status}]", end="")
                if not result.success:
                    print(f" {result.error_message}")
                else:
                    print()
        
        # 5. Save comments
        self._save_comments()
        
        return self.results
    
    def save(self, dry_run: bool = False):
        """Save modified document"""
        if dry_run:
            print(f"[DRY RUN] Would save to: {self.output_path}")
            return
        
        self.doc.save(str(self.output_path))
        print(f"Saved to: {self.output_path}")
    
    def save_failed_items(self) -> Optional[Path]:
        """
        Save failed edit items to JSONL file for retry.
        
        Returns:
            Path to failed items file if any failures exist, None otherwise
        """
        failed_results = [r for r in self.results if not r.success]
        
        if not failed_results:
            return None  # No failures, no file created
        
        # Generate output path: <input>_fail.jsonl
        fail_path = self.jsonl_path.with_stem(self.jsonl_path.stem + '_fail')
        
        with open(fail_path, 'w', encoding='utf-8') as f:
            # Write enhanced meta line
            meta_line = {
                **self.meta,  # Include all original meta fields
                'type': 'meta',  # Explicitly ensure type field exists (required for retry)
                'original_export': self.jsonl_path.name,
                'generated_at': datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%S%z'),
                'failed_count': len(failed_results),
                'total_count': len(self.edit_items)
            }
            json.dump(meta_line, f, ensure_ascii=False)
            f.write('\n')
            
            # Write failed items with error information (matching HTML export field order)
            # Note: content field removed - not needed for apply_audit_edits.py processing
            for result in failed_results:
                item = result.item
                data = {
                    'category': item.category,
                    'fix_action': item.fix_action,
                    'violation_reason': item.violation_reason,
                    'violation_text': item.violation_text,
                    'revised_text': item.revised_text,
                    'rule_id': item.rule_id,
                    'uuid': item.uuid,
                    'uuid_end': item.uuid_end,  # Required for retry
                    'heading': item.heading,
                    '_error': result.error_message  # Add error info for debugging
                }
                json.dump(data, f, ensure_ascii=False)
                f.write('\n')
        
        return fail_path

# ============================================================
# Main Function
# ============================================================

def main() -> int:
    parser = argparse.ArgumentParser(
        description="Apply audit results to Word document"
    )
    parser.add_argument('jsonl_file', help='Audit export file (JSONL format)')
    parser.add_argument('-o', '--output', help='Output file path')
    parser.add_argument('--author', default='AI', 
                       help='Author name for track changes/comments (default: AI)')
    parser.add_argument('--initials', 
                       help='Author initials for comments (default: first 2 chars of author)')
    parser.add_argument('--skip-hash', action='store_true', 
                       help='Skip hash verification')
    parser.add_argument('--dry-run', action='store_true', 
                       help='Validate only, do not save')
    parser.add_argument('-v', '--verbose', action='store_true', 
                       help='Verbose output')
    
    args = parser.parse_args()
    
    try:
        applier = AuditEditApplier(
            args.jsonl_file,
            output_path=args.output,
            author=args.author,
            initials=args.initials,
            skip_hash=args.skip_hash,
            verbose=args.verbose
        )
        
        print(f"Source file: {applier.source_path}")
        print(f"Output to: {applier.output_path}")
        print(f"Edit items: {len(applier.edit_items)}")
        print("-" * 50)
        
        results = applier.apply()
        
        # Statistics
        success_count = sum(1 for r in results if r.success and not r.warning)
        warning_count = sum(1 for r in results if r.success and r.warning)
        fail_count = sum(1 for r in results if not r.success)
        
        if warning_count > 0:
            print("\nWarning items (fallback to comment):")
            for r in results:
                if r.success and r.warning:
                    text_preview = format_text_preview(r.item.violation_text)
                    print(f"  - [{r.item.rule_id}] {r.error_message}: {text_preview}")
        
        if fail_count > 0:
            print("\nFailed items:")
            for r in results:
                if not r.success:
                    text_preview = format_text_preview(r.item.violation_text)
                    print(f"  - [{r.item.rule_id}] {r.error_message}: {text_preview}")

        print("-" * 50)
        print(f"Completed: {success_count} succeeded, {warning_count} warnings, {fail_count} failed")

        if not args.dry_run:
            applier.save()
        else:
            applier.save(dry_run=True)
        
        # Save failed items for retry
        fail_file = applier.save_failed_items()
        if fail_file:
            print(f"\n{'=' * 50}")
            print(f"Failed items saved to: {fail_file}")
            print(f"  → You can modify and retry with this file")
            print(f"  → Command: python {sys.argv[0]} {fail_file} --skip-hash")
            print(f"{'=' * 50}")
            
        return 0
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

if __name__ == '__main__':
    sys.exit(main())
