#!/usr/bin/env python3
"""
ABOUTME: Applies audit results to Word documents with track changes and comments
ABOUTME: Reads JSONL export from audit report and modifies the source document
"""

import argparse
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

# ============================================================
# Data Classes
# ============================================================

@dataclass
class EditItem:
    """Single edit item from JSONL export"""
    uuid: str                    # Paragraph ID (w14:paraId)
    violation_text: str          # Text to find
    violation_reason: str        # Reason for violation  
    fix_action: str              # delete | replace | manual
    revised_text: str            # Replacement text or suggestion
    category: str                # Category
    rule_id: str                 # Rule ID
    heading: str = ''            # Violation heading/title
    content: str = ''            # Original paragraph content

@dataclass
class EditResult:
    """Result of processing an edit item"""
    success: bool
    item: EditItem
    error_message: Optional[str] = None

# ============================================================
# Helper Functions
# ============================================================

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
                 author: str = 'AI-Assistant', initials: str = 'AI'):
        self.jsonl_path = Path(jsonl_path)
        self.skip_hash = skip_hash
        self.verbose = verbose
        self.author = author
        self.initials = initials if initials else author[:2] if len(author) >= 2 else author
        
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
        2. Nested format (from run_audit.py outpu): Each line contains multiple violations per paragraph
        """
        meta = {}
        items = []
        
        with open(self.jsonl_path, 'r', encoding='utf-8') as f:
            for line in f:
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
                    content = data.get('p_content', '')  # Map from p_content
                    
                    # Flatten violations array into separate EditItems
                    for v in violations:
                        items.append(EditItem(
                            uuid=uuid,
                            violation_text=v.get('violation_text', ''),
                            violation_reason=v.get('violation_reason', ''),
                            fix_action=v.get('fix_action', 'manual'),
                            revised_text=v.get('revised_text', ''),
                            category=v.get('category', ''),
                            rule_id=v.get('rule_id', ''),
                            heading=heading,
                            content=content
                        ))
                else:
                    # Flat format (existing format for backward compatibility)
                    items.append(EditItem(
                        uuid=data.get('uuid', ''),
                        violation_text=data.get('violation_text', ''),
                        violation_reason=data.get('violation_reason', ''),
                        fix_action=data.get('fix_action', 'manual'),
                        revised_text=data.get('revised_text', ''),
                        category=data.get('category', ''),
                        rule_id=data.get('rule_id', ''),
                        heading=data.get('heading', ''),
                        content=data.get('content', '')
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
    
    def _find_para_node_by_id(self, para_id: str):
        """
        Find paragraph by w14:paraId using XPath deep search.
        Handles paragraphs nested in tables.
        """
        xpath_expr = f'.//w:p[@w14:paraId="{para_id}"]'
        nodes = self.body_elem.xpath(xpath_expr)
        return nodes[0] if nodes else None
    
    def _iter_paragraphs_following(self, start_node) -> Generator:
        """
        Generator: iterate paragraphs from start_node in document order.
        Handles transitions from table cells to main body and vice versa.
        """
        all_paras = self.body_elem.xpath('.//w:p')
        
        try:
            start_index = all_paras.index(start_node)
            for p in all_paras[start_index:]:
                yield p
        except ValueError:
            return
    
    # ==================== Run Processing ====================
    
    def _collect_runs_info(self, para_elem) -> List[Dict]:
        """
        Collect information about all runs in a paragraph.
        Returns: [{'elem': run, 'text': str, 'start': int, 'end': int, 'rPr': element}, ...]
        """
        runs_info = []
        pos = 0
        
        for run in para_elem.findall('.//w:r', NS):
            text_elems = run.findall('w:t', NS)
            if not text_elems:
                continue
            
            run_text = ''.join(t.text or '' for t in text_elems)
            if not run_text:
                continue
                
            rPr = run.find('w:rPr', NS)
            
            runs_info.append({
                'elem': run,
                'text': run_text,
                'start': pos,
                'end': pos + len(run_text),
                'rPr': rPr
            })
            pos += len(run_text)
        
        return runs_info
    
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
        """Escape XML special characters"""
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
        insert_idx = list(parent).index(first_run)
        
        # Remove old runs
        for info in affected_runs:
            parent.remove(info['elem'])
        
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
        del_xml = f'''<w:del xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.author}" w:date="{timestamp}">
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
        
        # Calculate diff first to check only the parts that will be modified
        diff_ops = self._calculate_diff(violation_text, revised_text)
        
        # Check if any 'delete' operation overlaps with previous modifications
        # Only the deleted parts need to be checked, not 'equal' or 'insert' parts
        current_pos = match_start
        for op, text in diff_ops:
            if op == 'delete':
                # Find runs for this delete operation
                del_end = current_pos + len(text)
                del_affected = self._find_affected_runs(orig_runs_info, current_pos, del_end)
                if self._check_overlap_with_revisions(del_affected):
                    return 'conflict'
                current_pos = del_end
            elif op == 'equal':
                # Equal text consumes position but doesn't need conflict check
                current_pos += len(text)
            # 'insert' doesn't consume original text position
        
        rPr_xml = self._get_rPr_xml(affected[0]['rPr'])
        timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
        
        # Calculate split points
        first_run = affected[0]
        last_run = affected[-1]
        before_text = first_run['text'][:match_start - first_run['start']]
        after_text = last_run['text'][match_end - last_run['start']:]
        
        new_elements = []
        
        # Before text
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))
        
        # Apply diff operations
        for op, text in diff_ops:
            if op == 'equal':
                new_elements.append(self._create_run(text, rPr_xml))
            elif op == 'delete':
                change_id = self._get_next_change_id()
                del_xml = f'''<w:del xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.author}" w:date="{timestamp}">
                    <w:r>{rPr_xml}<w:delText>{self._escape_xml(text)}</w:delText></w:r>
                </w:del>'''
                new_elements.append(etree.fromstring(del_xml))
            elif op == 'insert':
                change_id = self._get_next_change_id()
                ins_xml = f'''<w:ins xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{self.author}" w:date="{timestamp}">
                    <w:r>{rPr_xml}<w:t>{self._escape_xml(text)}</w:t></w:r>
                </w:ins>'''
                new_elements.append(etree.fromstring(ins_xml))
        
        # After text
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))
        
        self._replace_runs(para_elem, affected, new_elements)
        return 'success'
    
    # ==================== Manual (Comment) Operation ====================
    
    def _apply_error_comment(self, para_elem, item: EditItem) -> bool:
        """
        Insert an unselected comment at the end of paragraph for failed items.
        
        Comment format: {WHY}<violation_reason>  {WHERE}<violation_text>{SUGGEST}<revised_text>
        Author: AI-Error
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
        self.comments.append({
            'id': comment_id,
            'text': comment_text,
            'author': 'AI-Error'
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
            'author': 'AI-Fallback'
        })
        
        return True
    
    def _apply_manual(self, para_elem, violation_text: str,
                     violation_reason: str, revised_text: str,
                     orig_runs_info: List[Dict],
                     orig_match_start: int) -> str:
        """
        Apply manual operation by adding a Word comment.
        
        Args:
            para_elem: Paragraph element
            violation_text: Text to mark with comment
            violation_reason: Reason to show in comment
            revised_text: Suggestion to show in comment
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
        
        Uses original text position directly to support commenting on text
        that may have been deleted by previous rules (Word displays comments
        on deleted text correctly).
        
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
        
        rPr_xml = self._get_rPr_xml(affected[0]['rPr'])
        
        # Calculate split points
        first_run = affected[0]
        last_run = affected[-1]
        before_text = first_run['text'][:match_start - first_run['start']]
        after_text = last_run['text'][match_end - last_run['start']:]
        
        new_elements = []
        
        # Before text
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))
        
        # Comment range start
        range_start = etree.fromstring(
            f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        )
        new_elements.append(range_start)
        
        # Commented text
        new_elements.append(self._create_run(violation_text, rPr_xml))
        
        # Comment range end
        range_end = etree.fromstring(
            f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        )
        new_elements.append(range_end)
        
        # Comment reference
        ref_xml = f'''<w:r xmlns:w="{NS['w']}">
            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            <w:commentReference w:id="{comment_id}"/>
        </w:r>'''
        new_elements.append(etree.fromstring(ref_xml))
        
        # After text
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))
        
        self._replace_runs(para_elem, affected, new_elements)
        
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
            # Support independent author for each comment (e.g., AI-Error for failed items)
            comment_author = comment.get('author', self.author)
            comment_elem.set(f'{{{NS["w"]}}}author', comment_author)
            comment_elem.set(f'{{{NS["w"]}}}date', timestamp)
            # Use author initials or default
            comment_initials = comment_author[:2] if len(comment_author) >= 2 else comment_author
            comment_elem.set(f'{{{NS["w"]}}}initials', comment_initials)

            # Add paragraph with text
            p = etree.SubElement(comment_elem, f'{{{NS["w"]}}}p')
            r = etree.SubElement(p, f'{{{NS["w"]}}}r')
            t = etree.SubElement(r, f'{{{NS["w"]}}}t')
            t.text = comment['text']
        
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
            # 1. Find anchor paragraph by ID
            anchor_para = self._find_para_node_by_id(item.uuid)
            if anchor_para is None:
                return EditResult(False, item,
                    f"Paragraph ID {item.uuid} not found (may be in header/footer or ID changed)")
            
            # 2. Search for text from anchor paragraph using ORIGINAL text (before revisions)
            # Store match results to pass to apply methods (avoid double matching)
            target_para = None
            matched_runs_info = None
            matched_start = -1
            violation_text_to_use = item.violation_text
            revised_text_to_use = item.revised_text
            numbering_stripped = False
            
            # Try original text first (using revision-free view)
            for para in self._iter_paragraphs_following(anchor_para):
                runs_info_orig, combined_orig = self._collect_runs_info_original(para)
                
                pos = combined_orig.find(item.violation_text)
                if pos != -1:
                    target_para = para
                    matched_runs_info = runs_info_orig
                    matched_start = pos
                    break
            
            # Fallback: Try stripping auto-numbering if original match failed
            if target_para is None:
                stripped_violation, was_stripped = strip_auto_numbering(item.violation_text)
                
                if was_stripped:
                    if self.verbose:
                        print(f"  [Fallback] Trying without numbering: {stripped_violation[:30]}...")
                    
                    for para in self._iter_paragraphs_following(anchor_para):
                        runs_info_orig, combined_orig = self._collect_runs_info_original(para)
                        
                        pos = combined_orig.find(stripped_violation)
                        if pos != -1:
                            target_para = para
                            matched_runs_info = runs_info_orig
                            matched_start = pos
                            numbering_stripped = True
                            violation_text_to_use = stripped_violation
                            
                            # Handle revised_text for replace operation
                            if item.fix_action == 'replace':
                                stripped_revised, revised_has_numbering = strip_auto_numbering(item.revised_text)
                                
                                if revised_has_numbering:
                                    # Both have numbering: strip both
                                    revised_text_to_use = stripped_revised
                                    if self.verbose:
                                        print(f"  [Fallback] Stripped numbering from both texts")
                                else:
                                    # Only violation_text has numbering: use original revised_text
                                    revised_text_to_use = item.revised_text
                                    if self.verbose:
                                        print(f"  [Fallback] Only violation_text had numbering")
                            
                            break
            
            if target_para is None:
                # Insert error comment immediately on failure
                self._apply_error_comment(anchor_para, item)
                return EditResult(False, item,
                    f"Text not found after anchor: {item.violation_text[:30]}...")
            
            # 3. Apply operation based on fix_action
            # Pass matched_runs_info and matched_start to avoid double matching
            success_status = None
            
            if item.fix_action == 'delete':
                success_status = self._apply_delete(
                    target_para, violation_text_to_use,
                    matched_runs_info, matched_start
                )
            elif item.fix_action == 'replace':
                success_status = self._apply_replace(
                    target_para, violation_text_to_use, revised_text_to_use,
                    matched_runs_info, matched_start
                )
            elif item.fix_action == 'manual':
                success_status = self._apply_manual(
                    target_para, violation_text_to_use,
                    item.violation_reason, revised_text_to_use,
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
                # Text overlaps with previous rule modification
                reason = "Text overlaps with previous rule modification"
                self._apply_fallback_comment(target_para, item, reason)
                if self.verbose:
                    print(f"  [Conflict] {reason}")
                return EditResult(True, item, f"Fallback to comment: {reason}")
            elif success_status == 'fallback':
                # Fallback to comment annotation
                reason = "Text not found in current document state"
                self._apply_fallback_comment(target_para, item, reason)
                if self.verbose:
                    print(f"  [Fallback] {reason}")
                return EditResult(True, item, f"Fallback to comment: {reason}")
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
                    'heading': item.heading,
                    'content': item.content,
                    '_error': result.error_message  # Add error info for debugging
                }
                json.dump(data, f, ensure_ascii=False)
                f.write('\n')
        
        return fail_path

# ============================================================
# Main Function
# ============================================================

def main():
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
        success_count = sum(1 for r in results if r.success)
        fail_count = len(results) - success_count
        
        print("-" * 50)
        print(f"Completed: {success_count} succeeded, {fail_count} failed")
        
        if fail_count > 0:
            print("\nFailed items:")
            for r in results:
                if not r.success:
                    print(f"  - [{r.item.rule_id}] {r.error_message}")
        
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
            
        sys.exit(0 if fail_count == 0 else 1)
        
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main()
