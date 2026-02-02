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

from utils import sanitize_xml_string

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
# Matches: "1. ", "1.1 ", "1) ", "1）", "a. ", "A) ", "• ", "表1 ", "图2 ", etc.
AUTO_NUMBERING_PATTERN = re.compile(
    r'^(?:'
    r'\d+(?:[\.\d)）]+)\s+'  # Numeric: 1. 1.1 1) 1）
    r'|'
    r'[a-zA-Z][.)）]\s+'     # Alphabetic: a. A) b）
    r'|'
    r'•\s*'                   # Bullet: • (optional space)
    r'|'
    r'[表图]\s*\d+\s*'        # Table/Figure: 表1 图2 表 3 图 4
    r')'
)

# Table row numbering pattern for detecting numeric first cell in JSON format
# Matches: ["1", ["2", ["10", etc. (first cell is a number)
TABLE_ROW_NUMBERING_PATTERN = re.compile(
    r'^\["\d+",\s*'
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

def json_escape(text: str) -> str:
    """
    Escape text for JSON format (without outer quotes).

    This matches the escaping that json.dumps() applies to string content,
    which is necessary for matching violation_text from LLM against
    reconstructed table content.

    Args:
        text: Raw text to escape

    Returns:
        JSON-escaped text (without surrounding quotes)
    """
    # json.dumps adds surrounding quotes, strip them
    return json.dumps(text, ensure_ascii=False)[1:-1]


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


def strip_table_row_numbering(text: str) -> Tuple[str, bool]:
    """
    Replaces leading table row numbering with empty string to match actual table structure.
    
    During parse phase, Word auto-numbering shows as "1", "2", "3" etc.
    During apply phase, the same cells contain empty strings "" because auto-numbering
    is not stored in the cell content. This function replaces the number with empty string
    to align with the actual table structure.
    
    Args:
        text: Text that may start with table row numbering pattern like '["1", '
        
    Returns:
        Tuple of (processed_text, was_stripped):
        - processed_text: Text with number replaced by "" if found, original otherwise
        - was_stripped: True if numbering was replaced, False otherwise
        
    Examples:
        '["1", "content"]' -> ('["", "content"]', True)
        '["content"]' -> ('["content"]', False)
    """
    match = TABLE_ROW_NUMBERING_PATTERN.match(text)
    if match:
        # Replace row number with empty string to match actual table structure
        # During parse: '["1", "content"]', during apply: '["", "content"]'
        return '["", ' + text[match.end():], True
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

        # Unified timestamp for all track changes and comments in one apply() run
        self.operation_timestamp: str = None

        # Results tracking
        self.results: List[EditResult] = []

    # ==================== Author Helpers ====================

    def _category_suffix(self, item: EditItem) -> str:
        """Normalize category suffix for author; default to 'uncategorized'."""
        category = (item.category or '').strip()
        return category if category else 'uncategorized'

    def _author_for_item(self, item: EditItem) -> str:
        """Build author name using base author + category suffix."""
        return f"{self.author}-{self._category_suffix(item)}"
    
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

    # ==================== Table Detection Helpers ====================

    def _find_ancestor(self, elem, tag: str):
        """
        Find ancestor element with specified tag.

        Args:
            elem: Starting element
            tag: Full tag name including namespace (e.g., '{http://...}tbl')

        Returns:
            Ancestor element if found, None otherwise
        """
        parent = elem.getparent()
        while parent is not None:
            if parent.tag == tag:
                return parent
            parent = parent.getparent()
        return None

    def _find_ancestor_table(self, para_elem):
        """Find the table (w:tbl) containing this paragraph, if any."""
        return self._find_ancestor(para_elem, f'{{{NS["w"]}}}tbl')

    def _find_ancestor_cell(self, para_elem):
        """Find the table cell (w:tc) containing this paragraph, if any."""
        return self._find_ancestor(para_elem, f'{{{NS["w"]}}}tc')

    def _find_ancestor_row(self, para_elem):
        """Find the table row (w:tr) containing this paragraph, if any."""
        return self._find_ancestor(para_elem, f'{{{NS["w"]}}}tr')

    def _is_paragraph_in_table(self, para_elem) -> bool:
        """Check if paragraph is inside a table."""
        return self._find_ancestor_table(para_elem) is not None

    def _find_para_by_uuid(self, uuid: str):
        """Find paragraph element by its w14:paraId."""
        return self._find_para_node_by_id(uuid)

    def _normalize_text_for_search(self, text: str) -> str:
        """
        Normalize text for search matching by removing trailing whitespace from each line.
        
        This matches parse_document.py behavior where para_text.strip() is called,
        preventing mismatch due to trailing spaces in XML runs.
        
        Args:
            text: Text to normalize
        
        Returns:
            Normalized text with trailing whitespace removed from each line
        
        TODO: Known limitation - index misalignment with trailing whitespace
            The normalized string is shorter than the original when lines have trailing
            whitespace. Currently, matched_start from normalized.find() is used directly
            against runs_info_orig built from unmodified text. If any earlier line has
            trailing spaces before '\\n', the normalized position points earlier than the
            real position in the original runs, causing _find_affected_runs to select the
            wrong span. This can lead to edits/comments landing in incorrect locations.
            
            Future improvement: Return (normalized_text, index_mapping) where
            index_mapping[normalized_pos] = original_pos, allowing accurate translation
            of match positions back to original text coordinates.
            
            Current risk: LOW - trailing whitespace in paragraphs is rare in practice,
            and incorrect edits would be caught during track changes review.
        """
        if not text:
            return text
        lines = text.split('\n')
        return '\n'.join(line.rstrip() for line in lines)

    def _find_tables_in_range(self, start_para, uuid_end: str) -> List:
        """
        Find all tables within uuid → uuid_end range.
        
        Args:
            start_para: Starting paragraph element
            uuid_end: End boundary paraId (inclusive)
        
        Returns:
            List of table elements (w:tbl) found in the range
        """
        tables = []
        seen_tables = set()
        
        for para in self._iter_paragraphs_in_range(start_para, uuid_end):
            # Check if this paragraph is in a table
            table = self._find_ancestor_table(para)
            if table is not None:
                # Use element ID to avoid duplicates
                table_id = id(table)
                if table_id not in seen_tables:
                    tables.append(table)
                    seen_tables.add(table_id)
        
        return tables

    def _search_in_table_cell_raw(self, table_elem, violation_text: str,
                                   start_para, uuid_end: str) -> Optional[Tuple]:
        """
        Search for raw text in table cells (non-JSON mode).
        
        Each cell is searched independently to prevent cross-cell matching.
        Multi-paragraph content within a cell is joined with real '\\n'.
        
        This addresses the case where:
        - violation_text is plain text (not JSON format)
        - Target content is in a table cell with multiple paragraphs
        - LLM has decoded JSON-escaped '\\\\n' to real '\\n'
        
        Args:
            table_elem: Table element (w:tbl) to search in
            violation_text: Plain text to find (with real newlines)
            start_para: Starting paragraph (unused - kept for API compatibility)
            uuid_end: End boundary paraId
        
        Returns:
            Tuple of (target_para, runs_info, match_start, matched_text) or None if not found
            - matched_text: The actual text that was matched (may be normalized if fallback was used)
        
        Note:
            The in_range gate has been removed because _find_tables_in_range already
            constrains the table to the uuid range. The anchor paragraph (start_para)
            is often outside the table (e.g., heading before table), so checking for
            it would cause all cells to be skipped.
        """
        # Iterate cells in document order
        for tc in table_elem.iter(f'{{{NS["w"]}}}tc'):
            cell_paras = tc.findall(f'{{{NS["w"]}}}p')
            if not cell_paras:
                continue
            
            # Build cell content with run tracking
            cell_runs = []
            pos = 0
            cell_first_para = cell_paras[0]
            
            for para_idx, para in enumerate(cell_paras):
                para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                
                # Add paragraph boundary (not for first paragraph in cell)
                if para_idx > 0 and cell_runs:
                    cell_runs.append({
                        'text': '\n',  # Real newline
                        'start': pos,
                        'end': pos + 1,
                        'para_elem': para,
                        'cell_elem': tc,
                        'is_para_boundary': True
                    })
                    pos += 1
                
                # Collect paragraph runs (original text, no JSON escaping)
                para_runs, _ = self._collect_runs_info_original(para)
                for run in para_runs:
                    run_copy = dict(run)
                    run_copy['para_elem'] = para
                    run_copy['cell_elem'] = tc
                    run_copy['start'] = run['start'] + pos
                    run_copy['end'] = run['end'] + pos
                    cell_runs.append(run_copy)
                
                if cell_runs:
                    pos = cell_runs[-1]['end']
                
                # Stop at uuid_end
                if para_id == uuid_end:
                    break
            
            if not cell_runs:
                continue
            
            # Search within this cell only (prevents cross-cell matching)
            cell_text = ''.join(r['text'] for r in cell_runs)
            # Normalize to match parse_document.py behavior (removes trailing whitespace)
            cell_normalized = self._normalize_text_for_search(cell_text)
            match_pos = cell_normalized.find(violation_text)
            
            # Fallback: If violation_text contains \\n literal (LLM didn't decode JSON escape),
            # try converting to real newline
            matched_text = violation_text  # Track the actual matched text
            if match_pos == -1 and '\\n' in violation_text:
                normalized_violation = violation_text.replace('\\n', '\n')
                match_pos = cell_normalized.find(normalized_violation)
                if match_pos != -1:
                    matched_text = normalized_violation  # Use normalized text for consistent span calculation
                    if self.verbose:
                        print(f"  [Fallback] Matched after converting \\\\n to real newline")
            
            if match_pos != -1:
                # Found! Return match info including the matched text
                if self.verbose:
                    print(f"  [Success] Found in cell (raw text mode): '{cell_text[:50]}...'")
                return (cell_first_para, cell_runs, match_pos, matched_text)
        
        return None

    def _get_cell_merge_properties(self, tcPr):
        """
        Get cell merge properties (gridSpan and vMerge).
        
        Args:
            tcPr: Cell properties element (w:tcPr)
        
        Returns:
            Tuple of (grid_span, vmerge_type)
            - grid_span: Number of columns this cell spans (default 1)
            - vmerge_type: 'restart' | 'continue' | None
        """
        grid_span = 1
        vmerge_type = None
        
        if tcPr is not None:
            # Check gridSpan (horizontal merge)
            gs = tcPr.find(f'{{{NS["w"]}}}gridSpan')
            if gs is not None:
                try:
                    grid_span = int(gs.get(f'{{{NS["w"]}}}val'))
                except (ValueError, TypeError):
                    grid_span = 1
            
            # Check vMerge (vertical merge)
            vmerge_elem = tcPr.find(f'{{{NS["w"]}}}vMerge')
            if vmerge_elem is not None:
                vmerge_val = vmerge_elem.get(f'{{{NS["w"]}}}val')
                if vmerge_val == 'restart':
                    vmerge_type = 'restart'
                else:
                    # None or 'continue' both mean continue
                    vmerge_type = 'continue'
        
        return grid_span, vmerge_type

    # ==================== Run Processing ====================

    def _collect_runs_info_across_paragraphs(self, start_para, uuid_end: str) -> Tuple[List[Dict], str, bool, Optional[str]]:
        """
        Collect run info across multiple paragraphs (uuid → uuid_end range).

        Behavior:
        - Body text: Paragraph boundaries are represented as '\\n'
        - Table content (same row): JSON format ["cell1", "cell2"] with '", "' between cells
        - Cross body/table boundary: Returns boundary_error
        - Cross table row boundary: Returns row_boundary_error (Word doesn't support cross-row comments)

        Args:
            start_para: Starting paragraph element
            uuid_end: End boundary paraId (inclusive)

        Returns:
            Tuple of (runs_info, combined_text, is_cross_paragraph, boundary_error)
            - runs_info: List of run info dicts with 'text', 'start', 'end', 'para_elem', etc.
            - combined_text: Full text string
            - is_cross_paragraph: True if text spans multiple paragraphs
            - boundary_error: None if OK, or error type string:
              - 'boundary_crossed': Crossed body/table or different tables
              - 'row_boundary_crossed': Crossed table row boundary
        """
        # 1. Detect if start paragraph is in a table
        start_in_table = self._is_paragraph_in_table(start_para)
        start_table = self._find_ancestor_table(start_para) if start_in_table else None

        # 2. Find end paragraph and check its context
        end_para = self._find_para_by_uuid(uuid_end)
        end_in_table = self._is_paragraph_in_table(end_para) if end_para is not None else False
        end_table = self._find_ancestor_table(end_para) if end_in_table else None

        # 3. Check boundary consistency
        if start_in_table != end_in_table:
            # Crossed body/table boundary
            return [], '', False, 'boundary_crossed'

        if start_in_table and start_table is not end_table:
            # Crossed different tables
            return [], '', False, 'boundary_crossed'

        # Note: Row boundary check is NOT done here.
        # For multi-row table blocks, we collect content row by row and let
        # the caller check if the actual match spans multiple rows.

        # 4. Dispatch to appropriate collector
        if start_in_table:
            return self._collect_runs_info_in_table(start_para, uuid_end, start_table)
        else:
            return self._collect_runs_info_in_body(start_para, uuid_end)

    def _collect_runs_info_in_body(self, start_para, uuid_end: str) -> Tuple[List[Dict], str, bool, Optional[str]]:
        """
        Collect run info across paragraphs in document body (not in table).
        Paragraph boundaries are represented as '\\n'.

        Args:
            start_para: Starting paragraph element
            uuid_end: End boundary paraId (inclusive)

        Returns:
            Tuple of (runs_info, combined_text, is_cross_paragraph, boundary_error)
        """
        runs_info = []
        pos = 0
        para_count = 0

        for para in self._iter_paragraphs_in_range(start_para, uuid_end):
            para_count += 1

            # Collect runs for this paragraph
            para_runs, para_text = self._collect_runs_info_original(para)

            # Skip empty paragraphs to match parse_document.py behavior
            if not para_text.strip():
                continue

            # Add paragraph element reference to each run
            for run in para_runs:
                run['para_elem'] = para
                run['start'] += pos
                run['end'] += pos
                runs_info.append(run)

            pos += len(para_text)

            # Add paragraph boundary (except after last paragraph)
            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
            if para_id != uuid_end:
                # Not the last paragraph - add boundary marker
                runs_info.append({
                    'text': '\n',
                    'start': pos,
                    'end': pos + 1,
                    'para_elem': para,
                    'is_para_boundary': True  # Mark as paragraph boundary
                })
                pos += 1

        combined_text = ''.join(r['text'] for r in runs_info)
        is_cross_paragraph = para_count > 1

        return runs_info, combined_text, is_cross_paragraph, None

    def _collect_runs_info_in_table(self, start_para, uuid_end: str, table_elem) -> Tuple[List[Dict], str, bool, Optional[str]]:
        """
        Collect run info within a table using JSON format, handling multiple rows and merged cells.

        For each row, format matches parse_document.py: ["cell1", "cell2", "cell3"]
        - Cell boundaries within row: '", "'
        - Paragraph boundaries within cell: '\\n' (JSON-escaped as '\\n')
        - Row boundaries: '"], ["' (close previous row, open new row)
        - Content is JSON-escaped to match LLM output
        
        Merged cell handling (consistent with TableExtractor):
        - Horizontal merge (gridSpan): Content in first cell only, spans multiple columns
        - Vertical merge (vMerge):
          - restart: Start of merge region, content recorded in vmerge_content
          - continue: Copy content from vmerge_content if available (within uuid range)
          - If restart is outside uuid range, continue cells remain empty
        
        Note: This method collects content across multiple rows. The caller should
        check if the actual match spans multiple rows using _check_cross_row_boundary().

        Args:
            start_para: Starting paragraph element
            uuid_end: End boundary paraId (inclusive)
            table_elem: The table element (w:tbl)

        Returns:
            Tuple of (runs_info, combined_text, is_cross_paragraph, boundary_error)
            - runs_info includes 'cell_elem' and 'row_elem' fields
            - runs_info includes 'original_text' for actual content (before JSON escaping)
        """
        # Get number of columns from tblGrid
        tbl_grid = table_elem.find(f'{{{NS["w"]}}}tblGrid')
        num_cols = 0
        if tbl_grid is not None:
            num_cols = len(tbl_grid.findall(f'{{{NS["w"]}}}gridCol'))
        
        if num_cols == 0:
            # No columns defined, fallback to paragraph iteration
            return self._collect_runs_info_in_table_legacy(start_para, uuid_end, table_elem)
        
        runs_info = []
        pos = 0
        in_range = False
        first_row_in_range = True
        vmerge_content = {}  # {grid_col: {'runs': [...], 'text': str}}
        last_para = start_para
        reached_end = False  # Flag to stop when uuid_end is found
        
        # Iterate rows (w:tr elements)
        for tr in table_elem.findall(f'{{{NS["w"]}}}tr'):
            if reached_end:
                break
            
            grid_col = 0
            row_data = []  # List of (cell_runs, cell_text) for this row
            
            # Iterate cells (w:tc elements) in this row
            for tc in tr.findall(f'{{{NS["w"]}}}tc'):
                if reached_end:
                    break
                # Get cell properties
                tcPr = tc.find(f'{{{NS["w"]}}}tcPr')
                grid_span, vmerge_type = self._get_cell_merge_properties(tcPr)
                
                # Collect all paragraphs in this cell
                cell_paras = tc.findall(f'{{{NS["w"]}}}p')
                
                # Check if any paragraph in this cell is in range
                cell_in_range = False
                for para in cell_paras:
                    para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                    if para is start_para:
                        in_range = True
                        cell_in_range = True
                    if in_range and para_id:
                        cell_in_range = True
                        last_para = para
                        if para_id == uuid_end:
                            break
                
                # Collect cell content if in range
                cell_runs = []
                cell_text = ''
                
                if cell_in_range:
                    # Determine cell content based on vMerge type
                    if vmerge_type == 'restart':
                        # Merge restart: collect content and store in vmerge_content
                        for para_idx, para in enumerate(cell_paras):
                            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                            if not in_range and para is not start_para:
                                continue
                            if para is start_para:
                                in_range = True
                            
                            if para_idx > 0:
                                # Add paragraph boundary marker
                                cell_runs.append({
                                    'text': '\\n',
                                    'is_json_escape': True,
                                    'is_para_boundary': True,
                                    'para_elem': para,
                                    'cell_elem': tc,
                                    'row_elem': tr
                                })
                                cell_text += '\\n'
                            
                            para_runs, _ = self._collect_runs_info_original(para)
                            for run in para_runs:
                                original_text = run['text']
                                escaped_text = json_escape(original_text)
                                cell_runs.append({
                                    **run,
                                    'original_text': original_text,
                                    'text': escaped_text,
                                    'para_elem': para,
                                    'cell_elem': tc,
                                    'row_elem': tr
                                })
                                cell_text += escaped_text
                            
                            if para_id == uuid_end:
                                break
                        
                        # Store in vmerge_content for future continue cells
                        vmerge_content[grid_col] = {
                            'runs': cell_runs,
                            'text': cell_text
                        }
                    
                    elif vmerge_type == 'continue':
                        # Merge continue: copy from vmerge_content if available
                        if grid_col in vmerge_content:
                            # Deep copy runs to avoid reference issues
                            cell_runs = [dict(r) for r in vmerge_content[grid_col]['runs']]
                            cell_text = vmerge_content[grid_col]['text']
                            # Update row_elem to current row
                            for run in cell_runs:
                                run['row_elem'] = tr
                        # else: restart outside range, keep empty
                    
                    else:
                        # Normal cell (no vMerge)
                        for para_idx, para in enumerate(cell_paras):
                            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                            if not in_range and para is not start_para:
                                continue
                            if para is start_para:
                                in_range = True
                            
                            if para_idx > 0:
                                cell_runs.append({
                                    'text': '\\n',
                                    'is_json_escape': True,
                                    'is_para_boundary': True,
                                    'para_elem': para,
                                    'cell_elem': tc,
                                    'row_elem': tr
                                })
                                cell_text += '\\n'
                            
                            para_runs, _ = self._collect_runs_info_original(para)
                            for run in para_runs:
                                original_text = run['text']
                                escaped_text = json_escape(original_text)
                                cell_runs.append({
                                    **run,
                                    'original_text': original_text,
                                    'text': escaped_text,
                                    'para_elem': para,
                                    'cell_elem': tc,
                                    'row_elem': tr
                                })
                                cell_text += escaped_text
                            
                            if para_id == uuid_end:
                                break
                        
                        # Check if empty cell inherits from vMerge (TableExtractor logic)
                        if not cell_text and grid_col in vmerge_content:
                            cell_runs = [dict(r) for r in vmerge_content[grid_col]['runs']]
                            cell_text = vmerge_content[grid_col]['text']
                            for run in cell_runs:
                                run['row_elem'] = tr
                        elif cell_text:
                            # Non-empty normal cell ends vMerge region
                            vmerge_content.pop(grid_col, None)
                
                # Store cell data - only if cell is in range or has content
                # This prevents adding empty cells before the start of the range
                if cell_in_range or cell_runs:
                    row_data.append((cell_runs, cell_text))
                else:
                    # Cell not in range and empty - don't add to row_data
                    # But we still need to account for it in the row structure
                    # So we add a placeholder that will be filtered later
                    pass
                
                # Advance grid column by gridSpan
                grid_col += grid_span
                
                # Check if we've reached the end - stop immediately when uuid_end is found
                if cell_in_range:
                    for para in cell_paras:
                        para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                        if para_id == uuid_end:
                            reached_end = True
                            break
            
            # After processing all cells in row, add to runs_info if row is in range
            if in_range and any(runs for runs, _ in row_data):
                # Add row opening marker (first row gets '["', subsequent get '"], ["')
                if first_row_in_range:
                    runs_info.append({
                        'text': '["',
                        'start': pos,
                        'end': pos + 2,
                        'para_elem': last_para,
                        'is_json_boundary': True
                    })
                    pos += 2
                    first_row_in_range = False
                else:
                    runs_info.append({
                        'text': '"], ["',
                        'start': pos,
                        'end': pos + 6,
                        'para_elem': last_para,
                        'is_json_boundary': True,
                        'is_row_boundary': True
                    })
                    pos += 6
                
                # Add cell data for this row
                for cell_idx, (cell_runs, cell_text) in enumerate(row_data):
                    if cell_idx > 0:
                        # Add cell boundary marker
                        runs_info.append({
                            'text': '", "',
                            'start': pos,
                            'end': pos + 4,
                            'para_elem': last_para,
                            'is_json_boundary': True,
                            'is_cell_boundary': True
                        })
                        pos += 4
                    
                    # Add cell runs
                    for run in cell_runs:
                        run['start'] = pos
                        run['end'] = pos + len(run['text'])
                        runs_info.append(run)
                        pos += len(run['text'])
            
        
        # Add closing '"]'
        if in_range:
            runs_info.append({
                'text': '"]',
                'start': pos,
                'end': pos + 2,
                'para_elem': last_para,
                'is_json_boundary': True
            })
            pos += 2
        
        combined_text = ''.join(r['text'] for r in runs_info)
        is_cross_paragraph = True  # Table mode is always treated as cross-paragraph
        
        return runs_info, combined_text, is_cross_paragraph, None
    
    def _collect_runs_info_in_table_legacy(self, start_para, uuid_end: str, table_elem) -> Tuple[List[Dict], str, bool, Optional[str]]:
        """
        Legacy table collection method (paragraph iteration without merge handling).
        
        Used as fallback when tblGrid is not available.
        """
        runs_info = []
        pos = 0
        current_cell = None
        current_row = None
        in_range = False
        last_para = start_para

        # Add opening '["'
        runs_info.append({
            'text': '["',
            'start': pos,
            'end': pos + 2,
            'para_elem': start_para,
            'is_json_boundary': True
        })
        pos += 2

        # Iterate all paragraphs in the table (in document order)
        for para in table_elem.iter(f'{{{NS["w"]}}}p'):
            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')

            # Check if we've entered the range
            if para is start_para:
                in_range = True

            if not in_range:
                continue

            last_para = para

            # Get cell and row for this paragraph
            cell = self._find_ancestor_cell(para)
            row = self._find_ancestor_row(para)

            # Handle row transition
            if row is not current_row:
                if current_row is not None:
                    # New row: close previous row and open new row '"], ["'
                    runs_info.append({
                        'text': '"], ["',
                        'start': pos,
                        'end': pos + 6,
                        'para_elem': para,
                        'is_json_boundary': True,
                        'is_row_boundary': True
                    })
                    pos += 6
                    current_cell = None  # Reset cell tracking for new row
                current_row = row

            # Handle cell transition (within same row)
            if cell is not current_cell:
                if current_cell is not None:
                    # New cell in same row: add '", "'
                    runs_info.append({
                        'text': '", "',
                        'start': pos,
                        'end': pos + 4,
                        'para_elem': para,
                        'is_json_boundary': True,
                        'is_cell_boundary': True
                    })
                    pos += 4
                current_cell = cell
            elif current_cell is not None:
                # Same cell, new paragraph: add '\n' (which is \\n in JSON)
                runs_info.append({
                    'text': '\\n',
                    'start': pos,
                    'end': pos + 2,
                    'para_elem': para,
                    'is_para_boundary': True,
                    'is_json_escape': True
                })
                pos += 2

            # Collect paragraph content (with JSON escaping)
            para_runs, _ = self._collect_runs_info_original(para)

            for run in para_runs:
                original_text = run['text']
                escaped_text = json_escape(original_text)

                run['original_text'] = original_text
                run['text'] = escaped_text
                run['para_elem'] = para
                run['cell_elem'] = cell
                run['row_elem'] = row
                run['start'] = pos
                run['end'] = pos + len(escaped_text)
                runs_info.append(run)

                pos += len(escaped_text)

            # Check if we've reached the end
            if para_id == uuid_end:
                break

        # Add closing '"]'
        runs_info.append({
            'text': '"]',
            'start': pos,
            'end': pos + 2,
            'para_elem': last_para,
            'is_json_boundary': True
        })
        pos += 2

        combined_text = ''.join(r['text'] for r in runs_info)
        is_cross_paragraph = True  # Table mode is always treated as cross-paragraph

        return runs_info, combined_text, is_cross_paragraph, None

    def _check_cross_cell_boundary(self, affected_runs: List[Dict]) -> bool:
        """
        Check if affected runs span multiple table cells.

        Args:
            affected_runs: List of run info dicts from _find_affected_runs

        Returns:
            True if runs span multiple cells, False otherwise
        """
        cells = set()
        for info in affected_runs:
            cell = info.get('cell_elem')
            if cell is not None:
                cells.add(id(cell))
        return len(cells) > 1

    def _check_cross_row_boundary(self, affected_runs: List[Dict]) -> bool:
        """
        Check if affected runs span multiple table rows.

        Args:
            affected_runs: List of run info dicts from _find_affected_runs

        Returns:
            True if runs span multiple rows, False otherwise
        """
        rows = set()
        for info in affected_runs:
            row = info.get('row_elem')
            if row is not None:
                rows.add(id(row))
        return len(rows) > 1

    def _try_extract_single_cell_edit(self, violation_text: str, revised_text: str,
                                       affected_runs: List[Dict], match_start: int) -> Optional[Dict]:
        """
        Try to extract single-cell edit when text spans multiple cells.

        This method analyzes the diff between violation_text and revised_text to
        determine if all changes are confined to a single table cell. If so, it
        extracts the cell-specific violation and revised text.

        Args:
            violation_text: Original violation text (may span multiple cells)
            revised_text: Revised text (may span multiple cells)
            affected_runs: List of run info dicts that would be affected
            match_start: Absolute position where violation_text starts in document

        Returns:
            Dict with keys: 'cell_violation', 'cell_revised', 'cell_elem', 'cell_runs'
            Or None if changes span multiple cells
        """
        # Calculate diff to find changed positions
        diff_ops = self._calculate_diff(violation_text, revised_text)
        
        # Track positions of all changes (delete/insert operations)
        change_positions = []
        current_pos = 0
        
        for op, text in diff_ops:
            if op == 'delete':
                # Record the position range of deleted text
                change_positions.append((current_pos, current_pos + len(text)))
                current_pos += len(text)
            elif op == 'insert':
                # Record the insertion point (zero-length range)
                change_positions.append((current_pos, current_pos))
                # insert doesn't consume original text position
            elif op == 'equal':
                current_pos += len(text)
        
        if not change_positions:
            # No changes detected
            return None
        
        # Find the cell(s) containing all changes
        cells_with_changes = set()
        
        # Use match_start as base offset for coordinate translation
        # change_positions are relative to violation_text (0-based)
        # match_start is the absolute position where violation_text starts in document
        base_offset = match_start
        
        for change_idx, (change_start_rel, change_end_rel) in enumerate(change_positions):
            # Translate relative positions to absolute positions
            change_start = base_offset + change_start_rel
            change_end = base_offset + change_end_rel
            
            # Determine if this is an insert (zero-length) or delete operation
            is_insert = (change_start == change_end)
            # Find runs affected by this change
            for run_idx, run in enumerate(affected_runs):
                # Skip synthetic boundary markers
                if (run.get('is_json_boundary') or 
                    run.get('is_json_escape') or 
                    run.get('is_para_boundary')):
                    continue
                
                # Check if this run is affected by the change (using absolute positions)
                is_affected = False
                if is_insert:
                    # Insert operation: check if insertion point is within run boundaries
                    # A run contains the insertion point if start <= pos <= end
                    if run['start'] <= change_start <= run['end']:
                        is_affected = True
                else:
                    # Delete operation: check if run overlaps with the deletion range
                    if run['end'] > change_start and run['start'] < change_end:
                        is_affected = True
                
                if is_affected:
                    cell = run.get('cell_elem')
                    if cell is not None:
                        cells_with_changes.add(id(cell))
        if len(cells_with_changes) != 1:
            # Changes span multiple cells or no cell found
            return None
        
        # All changes are in a single cell - extract cell-specific content
        # Find the target cell element
        target_cell_id = next(iter(cells_with_changes))
        target_cell = None
        cell_runs = []
        
        for run in affected_runs:
            if run.get('is_json_boundary') or run.get('is_json_escape') or run.get('is_para_boundary'):
                continue
            cell = run.get('cell_elem')
            if cell is not None and id(cell) == target_cell_id:
                target_cell = cell
                cell_runs.append(run)
        
        if target_cell is None or not cell_runs:
            return None
        
        # Extract the portion of violation_text and revised_text within this cell
        cell_start = min(r['start'] for r in cell_runs)
        cell_end = max(r['end'] for r in cell_runs)
        
        # Convert absolute positions to relative offsets within violation_text
        # cell_start/cell_end are absolute positions in combined_text
        # violation_text starts at match_start, so we need relative offsets
        relative_start = cell_start - match_start
        relative_end = cell_end - match_start
        
        # Extract cell_violation using relative offsets
        # Account for JSON escaping if in table mode
        if self._is_table_mode(affected_runs):
            cell_violation_escaped = violation_text[relative_start:relative_end]
            cell_violation = self._decode_json_escaped(cell_violation_escaped)
        else:
            cell_violation = violation_text[relative_start:relative_end]
        
        # Build cell_revised by applying diff operations to cell_violation
        # This handles cases where insertion/deletion changes text length
        cell_revised = cell_violation
        
        # Apply diff operations that fall within the cell range
        violation_pos = 0
        revised_accumulator = ''
        
        for op, text in diff_ops:
            if op == 'equal':
                chunk_start = violation_pos
                chunk_end = violation_pos + len(text)
                
                # Check if this chunk overlaps with cell range [relative_start, relative_end)
                if chunk_end > relative_start and chunk_start < relative_end:
                    # Calculate overlap
                    overlap_start = max(0, relative_start - chunk_start)
                    overlap_end = min(len(text), relative_end - chunk_start)
                    revised_accumulator += text[overlap_start:overlap_end]
                
                violation_pos += len(text)
            
            elif op == 'delete':
                chunk_start = violation_pos
                chunk_end = violation_pos + len(text)
                
                # Skip deleted text if it falls within cell range
                # (already handled by not including it in revised_accumulator)
                violation_pos += len(text)
            
            elif op == 'insert':
                # Insert operations don't have a position in violation_text
                # Check if insertion point is within cell range
                if relative_start <= violation_pos < relative_end:
                    revised_accumulator += text
        
        cell_revised = revised_accumulator if revised_accumulator else cell_violation
        
        return {
            'cell_violation': cell_violation,
            'cell_revised': cell_revised,
            'cell_elem': target_cell,
            'cell_runs': cell_runs
        }

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
                            # Handle line breaks - textWrapping or no type = soft line break
                            br_type = elem.get(f'{{{NS["w"]}}}type')
                            if br_type in (None, 'textWrapping'):
                                runs_info.append({
                                    'text': '\n',
                                    'start': pos,
                                    'end': pos + 1,
                                    'elem': run,
                                    'rPr': rPr
                                })
                                pos += 1
                            # Skip page and column breaks (layout elements)
                            
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
                        # Handle line breaks - textWrapping or no type = soft line break
                        br_type = elem.get(f'{{{NS["w"]}}}type')
                        if br_type in (None, 'textWrapping'):
                            runs_info.append({
                                'text': '\n',
                                'start': pos,
                                'end': pos + 1,
                                'elem': child,
                                'rPr': rPr
                            })
                            pos += 1
                        # Skip page and column breaks (layout elements)
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

    def _filter_real_runs(self, runs: List[Dict]) -> List[Dict]:
        """
        Filter out synthetic boundary runs (JSON boundaries, paragraph boundaries).

        Synthetic runs are injected for text matching but don't have actual
        document elements (elem/rPr). This method filters them out before
        applying document modifications.

        Args:
            runs: List of run info dicts

        Returns:
            List of runs that have actual document elements
        """
        return [r for r in runs
                if not r.get('is_json_boundary', False)
                and not r.get('is_json_escape', False)
                and not r.get('is_para_boundary', False)]

    def _get_run_original_text(self, run: Dict) -> str:
        """
        Get the original (unescaped) text for a run.

        In table mode, runs have 'original_text' with unescaped content
        and 'text' with JSON-escaped content. For document mutations,
        we always want the original text.

        Args:
            run: Run info dict

        Returns:
            Original text content (unescaped)
        """
        return run.get('original_text', run['text'])

    def _is_table_mode(self, runs: List[Dict]) -> bool:
        """
        Check if we're operating in table mode (runs have JSON-escaped text).

        Args:
            runs: List of run info dicts

        Returns:
            True if any run has 'original_text' field (indicating table mode)
        """
        return any(r.get('original_text') is not None for r in runs)

    def _decode_json_escaped(self, text: str) -> str:
        """
        Decode JSON-escaped text back to original.

        In table mode, violation_text and revised_text from the LLM contain
        JSON escape sequences (like \" for quotes, \\n for newlines).
        This method decodes them back to the original characters.

        Args:
            text: JSON-escaped text string

        Returns:
            Decoded original text
        """
        if not text:
            return text
        try:
            # json.loads expects a complete JSON string with quotes
            return json.loads('"' + text + '"')
        except json.JSONDecodeError:
            # Fallback: if decode fails, return original text
            return text

    def _translate_escaped_offset(self, run: Dict, escaped_offset: int) -> int:
        """
        Translate an offset from escaped text space to original text space.

        In table mode, run['text'] contains JSON-escaped content where
        characters like " become \\". This method translates offsets
        from the escaped space to the original text space.

        Args:
            run: Run info dict with 'text' and possibly 'original_text'
            escaped_offset: Offset within run['text']

        Returns:
            Corresponding offset in original text
        """
        if 'original_text' not in run:
            return escaped_offset  # No escaping, offset is the same

        escaped_text = run['text']
        original_text = run['original_text']

        if escaped_offset <= 0:
            return 0
        if escaped_offset >= len(escaped_text):
            return len(original_text)

        # Try to decode the escaped prefix as JSON string content
        escaped_prefix = escaped_text[:escaped_offset]
        try:
            # json.loads expects a complete JSON string
            original_prefix = json.loads('"' + escaped_prefix + '"')
            return len(original_prefix)
        except json.JSONDecodeError:
            # Partial escape sequence - find the last valid boundary
            for i in range(escaped_offset - 1, -1, -1):
                try:
                    original_prefix = json.loads('"' + escaped_text[:i] + '"')
                    return len(original_prefix)
                except json.JSONDecodeError:
                    continue
            return 0

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
    
    def _replace_runs(self, _para_elem, affected_runs: List[Dict],
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
                     violation_reason: str,
                     orig_runs_info: List[Dict],
                     orig_match_start: int,
                     author: str) -> str:
        """
        Apply delete operation with track changes and comment annotation.

        Args:
            para_elem: Paragraph element
            violation_text: Text to delete
            violation_reason: Reason for violation (used as comment text)
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
            author: Track change author (base author + category suffix)

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

        # Filter out synthetic boundary runs (JSON boundaries, para boundaries)
        real_runs = self._filter_real_runs(affected)
        if not real_runs:
            return 'fallback'

        # Check if text overlaps with previous modifications
        if self._check_overlap_with_revisions(real_runs):
            return 'conflict'

        rPr_xml = self._get_rPr_xml(real_runs[0].get('rPr'))
        change_id = self._get_next_change_id()

        # Allocate comment ID for wrapping the deleted text
        comment_id = self.next_comment_id
        self.next_comment_id += 1

        # Calculate split points using real runs only
        # Use original text for mutations (not JSON-escaped text)
        first_run = real_runs[0]
        last_run = real_runs[-1]
        first_orig_text = self._get_run_original_text(first_run)
        last_orig_text = self._get_run_original_text(last_run)

        # Translate offsets from escaped space to original space
        before_offset = self._translate_escaped_offset(first_run, max(0, match_start - first_run['start']))
        after_offset = self._translate_escaped_offset(last_run, max(0, match_end - last_run['start']))

        before_text = first_orig_text[:before_offset]
        after_text = last_orig_text[after_offset:]

        new_elements = []

        # Before text (unchanged)
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))

        # Comment range start (before deleted text)
        comment_start_xml = f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        new_elements.append(etree.fromstring(comment_start_xml))

        # Decode violation_text if in table mode (JSON-escaped)
        if self._is_table_mode(real_runs):
            del_text = self._decode_json_escaped(violation_text)
        else:
            del_text = violation_text

        # Deleted text
        del_xml = f'''<w:del xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{author}" w:date="{self.operation_timestamp}">
            <w:r>{rPr_xml}<w:delText>{self._escape_xml(del_text)}</w:delText></w:r>
        </w:del>'''
        new_elements.append(etree.fromstring(del_xml))

        # Comment range end and reference (after deleted text)
        comment_end_xml = f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        new_elements.append(etree.fromstring(comment_end_xml))

        comment_ref_xml = f'''<w:r xmlns:w="{NS['w']}">
            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            <w:commentReference w:id="{comment_id}"/>
        </w:r>'''
        new_elements.append(etree.fromstring(comment_ref_xml))

        # After text (unchanged)
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))

        self._replace_runs(para_elem, real_runs, new_elements)

        # Record comment with violation_reason as content
        # Use "-R" suffix to distinguish comment author from track change author
        self.comments.append({
            'id': comment_id,
            'text': violation_reason,
            'author': f"{author}-R"
        })

        return 'success'

    # ==================== Replace Operation ====================
    
    def _apply_replace(self, para_elem, violation_text: str,
                      revised_text: str,
                      violation_reason: str,
                      orig_runs_info: List[Dict],
                      orig_match_start: int,
                      author: str) -> str:
        """
        Apply replace operation with diff-based track changes and comment annotation.

        Strategy: Build all elements in a single pass, preserving original elements
        for 'equal' portions (to keep formatting, images, etc.) and creating
        track changes for delete/insert portions.

        Args:
            para_elem: Paragraph element
            violation_text: Text to replace
            revised_text: New text
            violation_reason: Reason for violation (used as comment text)
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
            author: Track change author (base author + category suffix)

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

        # Filter out synthetic boundary runs (JSON boundaries, para boundaries)
        real_runs = self._filter_real_runs(affected)
        if not real_runs:
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
                del_real = self._filter_real_runs(del_affected)
                if self._check_overlap_with_revisions(del_real):
                    return 'conflict'
                current_pos = del_end
            elif op == 'equal':
                current_pos += len(text)
            # insert doesn't consume original text position

        rPr_xml = self._get_rPr_xml(real_runs[0].get('rPr'))

        # Allocate comment ID for wrapping the replaced text
        comment_id = self.next_comment_id
        self.next_comment_id += 1

        # Calculate split points for before/after text using real runs
        # Use original text for mutations (not JSON-escaped text)
        first_run = real_runs[0]
        last_run = real_runs[-1]
        first_orig_text = self._get_run_original_text(first_run)
        last_orig_text = self._get_run_original_text(last_run)

        # Translate offsets from escaped space to original space
        before_offset = self._translate_escaped_offset(first_run, max(0, match_start - first_run['start']))
        after_offset = self._translate_escaped_offset(last_run, max(0, match_end - last_run['start']))

        before_text = first_orig_text[:before_offset]
        after_text = last_orig_text[after_offset:]

        # Build new elements in a single pass
        new_elements = []

        # Before text (unchanged part before the match)
        if before_text:
            new_elements.append(self._create_run(before_text, rPr_xml))

        # Comment range start (before replaced text)
        comment_start_xml = f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        new_elements.append(etree.fromstring(comment_start_xml))

        # Check if we're in table mode (need to decode JSON-escaped text)
        is_table_mode = self._is_table_mode(real_runs)

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
                        # Use original text for document mutations
                        for eq_run in equal_runs:
                            escaped_text = eq_run['text']
                            orig_text = self._get_run_original_text(eq_run)

                            # Calculate offsets in escaped space, then translate
                            escaped_start = max(0, equal_start - eq_run['start'])
                            escaped_end = min(len(escaped_text), equal_end - eq_run['start'])

                            # Translate to original space
                            orig_start = self._translate_escaped_offset(eq_run, escaped_start)
                            orig_end = self._translate_escaped_offset(eq_run, escaped_end)
                            portion = orig_text[orig_start:orig_end]

                            if eq_run.get('is_drawing'):
                                # Image run - copy entire element if fully contained
                                if escaped_start == 0 and escaped_end == len(escaped_text):
                                    new_elements.append(copy.deepcopy(eq_run['elem']))
                            elif portion:
                                new_elements.append(self._create_run(portion, rPr_xml))
                else:
                    # No runs found, create text directly
                    # Decode if in table mode
                    equal_text = self._decode_json_escaped(text) if is_table_mode else text
                    new_elements.append(self._create_run(equal_text, rPr_xml))

                violation_pos += len(text)

            elif op == 'delete':
                change_id = self._get_next_change_id()
                # Decode if in table mode
                del_text = self._decode_json_escaped(text) if is_table_mode else text
                del_xml = f'''<w:del xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{author}" w:date="{self.operation_timestamp}">
                    <w:r>{rPr_xml}<w:delText>{self._escape_xml(del_text)}</w:delText></w:r>
                </w:del>'''
                new_elements.append(etree.fromstring(del_xml))
                violation_pos += len(text)

            elif op == 'insert':
                change_id = self._get_next_change_id()
                # Decode if in table mode
                ins_text = self._decode_json_escaped(text) if is_table_mode else text
                ins_xml = f'''<w:ins xmlns:w="{NS['w']}" w:id="{change_id}" w:author="{author}" w:date="{self.operation_timestamp}">
                    <w:r>{rPr_xml}<w:t>{self._escape_xml(ins_text)}</w:t></w:r>
                </w:ins>'''
                new_elements.append(etree.fromstring(ins_xml))
                # insert doesn't consume violation_pos

        # Comment range end and reference (after replaced text)
        comment_end_xml = f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        new_elements.append(etree.fromstring(comment_end_xml))

        comment_ref_xml = f'''<w:r xmlns:w="{NS['w']}">
            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
            <w:commentReference w:id="{comment_id}"/>
        </w:r>'''
        new_elements.append(etree.fromstring(comment_ref_xml))

        # After text (unchanged part after the match)
        if after_text:
            new_elements.append(self._create_run(after_text, rPr_xml))

        # Single DOM operation to replace all real runs (not boundary markers)
        self._replace_runs(para_elem, real_runs, new_elements)

        # Record comment with violation_reason as content
        # Use "-R" suffix to distinguish comment author from track change author
        self.comments.append({
            'id': comment_id,
            'text': violation_reason,
            'author': f"{author}-R"
        })

        return 'success'

    # ==================== Manual (Comment) Operation ====================
    
    def _apply_error_comment(self, para_elem, item: EditItem, author_override: str = None) -> bool:
        """
        Insert an unselected comment at the end of paragraph for failed items.
        
        Comment format: {WHY}<violation_reason>  {WHERE}<violation_text>{SUGGEST}<revised_text>
        Author: author_override if provided, otherwise {self.author}-{category}
        
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
        comment_author = author_override if author_override else self._author_for_item(item)
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
            'author': self._author_for_item(item)
        })
        
        return True
    
    def _apply_manual(self, para_elem, violation_text: str,
                     violation_reason: str, revised_text: str,
                     orig_runs_info: List[Dict],
                     orig_match_start: int,
                     author: str,
                     is_cross_paragraph: bool = False) -> str:
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
        
        Cross-paragraph support:
        - When is_cross_paragraph=True, commentRangeStart/End can span multiple paragraphs
        - Uses 'para_elem' field from runs to determine which paragraph to insert markers into
        
        Args:
            para_elem: Paragraph element (anchor paragraph, may not be used in cross-para mode)
            violation_text: Text to mark with comment
            violation_reason: Reason to show in comment
            revised_text: Suggestion to show in comment
            orig_runs_info: Pre-computed original runs info from _process_item
            orig_match_start: Pre-computed match position in original text
            author: Comment author (base author + category suffix)
            is_cross_paragraph: If True, handle cross-paragraph comment range
        
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

        # Filter out all synthetic boundary markers (JSON and paragraph boundaries)
        real_runs = self._filter_real_runs(affected)

        if not real_runs:
            return 'fallback'

        comment_id = self.next_comment_id
        self.next_comment_id += 1

        first_run_info = real_runs[0]
        last_run_info = real_runs[-1]
        first_run = first_run_info['elem']
        last_run = last_run_info['elem']
        rPr_xml = self._get_rPr_xml(first_run_info.get('rPr'))

        # Get parent paragraphs (may be different in cross-paragraph mode)
        if is_cross_paragraph:
            first_para = first_run_info.get('para_elem', para_elem)
            last_para = last_run_info.get('para_elem', para_elem)
        else:
            first_para = para_elem
            last_para = para_elem

        # Check if start/end runs are inside revision markup
        start_revision = self._find_revision_ancestor(first_run, first_para)
        end_revision = self._find_revision_ancestor(last_run, last_para)

        # Calculate text split points using real runs
        # Use original text for mutations (not JSON-escaped text)
        first_orig_text = self._get_run_original_text(first_run_info)
        last_orig_text = self._get_run_original_text(last_run_info)

        # Translate offsets from escaped space to original space
        before_offset = self._translate_escaped_offset(first_run_info, max(0, match_start - first_run_info['start']))
        after_offset = self._translate_escaped_offset(last_run_info, max(0, match_end - last_run_info['start']))

        before_text = first_orig_text[:before_offset]
        after_text = last_orig_text[after_offset:]
        
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
            'text': comment_text,
            'author': author
        })
        
        return 'success'
    
    # ==================== Comment Saving ====================
    
    def _save_comments(self):
        """Save comments to comments.xml using OPC API"""
        if not self.comments:
            return

        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        
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
            # Support independent author for each comment (author-category suffix)
            comment_author = comment.get('author', self.author)
            comment_elem.set(f'{{{NS["w"]}}}author', comment_author)
            comment_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
            # Use self.initials for all comments with author prefix matching self.author
            # This includes: AI-<category> (all share same initials)
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
            
            item_author = self._author_for_item(item)

            # 2. Search for text from anchor paragraph using ORIGINAL text (before revisions)
            # Store match results to pass to apply methods (avoid double matching)
            # IMPORTANT: Search is restricted to uuid -> uuid_end range to prevent
            # accidental modifications to content in other text blocks
            target_para = None
            matched_runs_info = None
            matched_start = -1
            numbering_stripped = False
            
            # Strategy: Try single-paragraph search first, then cross-paragraph if needed
            is_cross_paragraph = False
            
            # Try original text first (using revision-free view) - single paragraph
            for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                runs_info_orig, combined_orig = self._collect_runs_info_original(para)
                
                # Normalize to match parse_document.py behavior (removes trailing whitespace)
                combined_normalized = self._normalize_text_for_search(combined_orig)
                pos = combined_normalized.find(violation_text)
                if pos != -1:
                    target_para = para
                    matched_runs_info = runs_info_orig
                    matched_start = pos
                    break
            
            # Fallback 1: Try stripping auto-numbering if original match failed
            if target_para is None:
                stripped_violation, was_stripped = strip_auto_numbering(violation_text)
                
                if was_stripped:
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
                                # else: Only violation_text had numbering, keep revised_text as-is
                            
                            break
            
            # Fallback 2: Try cross-paragraph search (within uuid → uuid_end range)
            # This also handles table content with JSON format matching
            boundary_error = None
            if target_para is None:
                # Collect runs across all paragraphs in range
                cross_runs, cross_text, is_multi_para, boundary_error = self._collect_runs_info_across_paragraphs(
                    anchor_para, item.uuid_end
                )

                if boundary_error:
                    # Boundary error detected - will handle below
                    if self.verbose:
                        print(f"  [Boundary] {boundary_error}")
                elif is_multi_para:
                    # Only use cross-paragraph mode if there are actually multiple paragraphs
                    # Normalize to match parse_document.py behavior (removes trailing whitespace)
                    cross_normalized = self._normalize_text_for_search(cross_text)
                    pos = cross_normalized.find(violation_text)
                    if pos != -1:
                        # Found match across paragraphs
                        target_para = anchor_para  # Use anchor as reference
                        matched_runs_info = cross_runs
                        matched_start = pos
                        is_cross_paragraph = True

                        if self.verbose:
                            print(f"  [Success] Found in cross-paragraph mode")
            
            # Fallback 3: Try table search if violation_text looks like JSON array
            if target_para is None and violation_text.startswith('["'):
                # Find all tables in range
                tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)
                
                # Try with original violation_text first, then with row numbering stripped
                search_attempts = [
                    (violation_text, False),  # Original text, not stripped
                ]
                
                # If violation_text starts with row number, add stripped version
                stripped_table_text, was_stripped = strip_table_row_numbering(violation_text)
                if was_stripped:
                    search_attempts.append((stripped_table_text, True))
                
                for search_text, is_stripped in search_attempts:
                    for table_elem in tables_in_range:
                        # Get the first and last paragraph in this table
                        table_paras = list(table_elem.iter(f'{{{NS["w"]}}}p'))
                        if not table_paras:
                            continue
                        
                        first_table_para = table_paras[0]
                        last_table_para = table_paras[-1]
                        
                        # Get their paraIds
                        first_para_id = first_table_para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                        last_para_id = last_table_para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                        
                        if not first_para_id or not last_para_id:
                            continue
                        
                        # Collect table content in JSON format
                        try:
                            table_runs, table_text, _, _ = self._collect_runs_info_in_table(
                                first_table_para, last_para_id, table_elem
                            )
                            
                            # Search for search_text in table content
                            pos = table_text.find(search_text)
                            if pos != -1:
                                # Found match in this table!
                                target_para = first_table_para  # Use first para as anchor
                                matched_runs_info = table_runs
                                matched_start = pos
                                is_cross_paragraph = True  # Table mode is always cross-paragraph
                                
                                # Update violation_text to the matched version
                                violation_text = search_text
                                
                                # For replace operations, also strip row numbering from revised_text
                                if is_stripped and item.fix_action == 'replace':
                                    stripped_revised, revised_was_stripped = strip_table_row_numbering(revised_text)
                                    if revised_was_stripped:
                                        revised_text = stripped_revised
                                
                                if self.verbose:
                                    if is_stripped:
                                        print(f"  [Success] Found in table after stripping row numbering")
                                    else:
                                        print(f"  [Success] Found in table (JSON format)")
                                break
                        except Exception:
                            # If table processing fails, continue to next table
                            continue
                    
                    # If found, break outer loop
                    if target_para is not None:
                        break
            
            # Fallback 2.5: Try non-JSON table search (raw text mode)
            # For plain text violation_text that may be in table cells with multiple paragraphs
            if target_para is None and not violation_text.startswith('["'):
                # Find all tables in range
                tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)
                
                for table_elem in tables_in_range:
                    # Search for raw text in each cell independently
                    result = self._search_in_table_cell_raw(
                        table_elem, violation_text, anchor_para, item.uuid_end
                    )
                    
                    if result:
                        target_para, matched_runs_info, matched_start, matched_text = result
                        # Update violation_text with the actual matched text (handles fallback normalization)
                        violation_text = matched_text
                        # Cell content is always treated as single-paragraph for now
                        is_cross_paragraph = False
                        
                        if self.verbose:
                            print(f"  [Success] Found in table cell (plain text mode)")
                        break
            
            if target_para is None:
                # Check if we have a boundary error from table/row crossing
                if boundary_error:
                    # Special handling for boundary_crossed: try searching in tables first
                    if boundary_error == 'boundary_crossed':
                        # Find all tables in range
                        tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)
                        
                        # Try searching raw text in each table cell
                        for table_elem in tables_in_range:
                            # Iterate all cells in this table
                            for tc in table_elem.iter(f'{{{NS["w"]}}}tc'):
                                # Collect all paragraphs in this cell
                                cell_paras = tc.findall(f'{{{NS["w"]}}}p')
                                if not cell_paras:
                                    continue
                                
                                # Build cell's raw text content (with \n between paragraphs)
                                cell_text_parts = []
                                cell_para_runs_map = {}  # Map from para to runs_info
                                
                                for cell_para in cell_paras:
                                    para_runs, para_text = self._collect_runs_info_original(cell_para)
                                    cell_text_parts.append(para_text)
                                    cell_para_runs_map[id(cell_para)] = (para_runs, para_text)
                                
                                # Join with \n (actual newline, not JSON-escaped)
                                cell_combined_text = '\n'.join(cell_text_parts)
                                
                                # Normalize to match parse_document.py behavior (removes trailing whitespace)
                                cell_normalized = self._normalize_text_for_search(cell_combined_text)
                                
                                # Search for violation_text in cell's raw text
                                pos = cell_normalized.find(violation_text)
                                if pos != -1:
                                    # Found match in this cell!
                                    if self.verbose:
                                        print(f"  [Success] Found in table cell (raw text match)")
                                    
                                    # Determine which paragraph(s) contain the match
                                    # For simplicity, if match is within first paragraph, use it
                                    # Otherwise, this is a cross-paragraph match within the cell
                                    current_offset = 0
                                    matched_para = None
                                    matched_para_runs = None
                                    matched_para_start = -1
                                    
                                    for cell_para in cell_paras:
                                        para_id_obj = id(cell_para)
                                        if para_id_obj not in cell_para_runs_map:
                                            continue
                                        
                                        para_runs, para_text = cell_para_runs_map[para_id_obj]
                                        para_len = len(para_text)
                                        
                                        # Check if match starts in this paragraph
                                        if current_offset <= pos < current_offset + para_len:
                                            matched_para = cell_para
                                            matched_para_runs = para_runs
                                            matched_para_start = pos - current_offset
                                            break
                                        
                                        current_offset += para_len + 1  # +1 for \n separator
                                    
                                    if matched_para is not None:
                                        # Use the matched paragraph
                                        target_para = matched_para
                                        matched_runs_info = matched_para_runs
                                        matched_start = matched_para_start
                                        is_cross_paragraph = False  # Single para in cell
                                        break
                            
                            # If found, break table loop
                            if target_para is not None:
                                break
                    
                    # If still not found after table search, try body text search
                    if target_para is None:
                        if self.verbose:
                            print(f"  [Boundary] Table search failed, trying body text...")
                        
                        # Collect ALL body paragraphs in range, grouped by continuity
                        # This handles: table→body, body→table, and interleaved scenarios
                        body_segments = []  # List of segments, each segment = [(para, runs, text), ...]
                        current_segment = []
                        
                        for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                            if self._is_paragraph_in_table(para):
                                # Encountered table paragraph - end current body segment
                                if current_segment:
                                    body_segments.append(current_segment)
                                    current_segment = []
                                continue
                            
                            # Body paragraph
                            para_runs, para_text = self._collect_runs_info_original(para)
                            # Skip empty paragraphs to match parse_document.py behavior
                            if not para_text.strip():
                                continue
                            
                            current_segment.append((para, para_runs, para_text))
                        
                        # Don't forget the last segment
                        if current_segment:
                            body_segments.append(current_segment)
                        
                        # Search in each body segment (try both original and stripped numbering)
                        for segment_idx, body_paras_data in enumerate(body_segments):
                            # Build combined text with \n separator (like _collect_runs_info_in_body)
                            all_runs = []
                            pos = 0
                            
                            for i, (para, para_runs, para_text) in enumerate(body_paras_data):
                                # Add paragraph runs with updated positions
                                for run in para_runs:
                                    run_copy = dict(run)
                                    run_copy['para_elem'] = para
                                    run_copy['start'] = run['start'] + pos
                                    run_copy['end'] = run['end'] + pos
                                    all_runs.append(run_copy)
                                
                                pos += len(para_text)
                                
                                # Add paragraph boundary (except after last)
                                if i < len(body_paras_data) - 1:
                                    all_runs.append({
                                        'text': '\n',
                                        'start': pos,
                                        'end': pos + 1,
                                        'para_elem': para,
                                        'is_para_boundary': True
                                    })
                                    pos += 1
                            
                            combined_body_text = ''.join(r['text'] for r in all_runs)
                            
                            # Normalize to match parse_document.py behavior (removes trailing whitespace)
                            combined_normalized = self._normalize_text_for_search(combined_body_text)
                            
                            # Try multiple search patterns (original and stripped numbering)
                            search_attempts = [(violation_text, False)]
                            stripped_v, was_stripped = strip_auto_numbering(violation_text)
                            if was_stripped:
                                search_attempts.append((stripped_v, True))
                            
                            match_pos = -1
                            matched_text = violation_text
                            
                            for search_text, is_stripped in search_attempts:
                                match_pos = combined_normalized.find(search_text)
                                if match_pos != -1:
                                    matched_text = search_text
                                    break
                            
                            if match_pos != -1:
                                # Found match in this segment!
                                target_para = body_paras_data[0][0]  # Use first para as anchor
                                matched_runs_info = all_runs
                                matched_start = match_pos
                                is_cross_paragraph = len(body_paras_data) > 1
                                violation_text = matched_text  # Update violation_text to matched version
                                numbering_stripped = (matched_text != item.violation_text.strip())
                                
                                if self.verbose:
                                    if is_cross_paragraph:
                                        print(f"  [Success] Found in body segment {segment_idx + 1} (cross-paragraph)")
                                    else:
                                        print(f"  [Success] Found in body segment {segment_idx + 1}")
                                break  # Stop searching other segments
                    
                    # If still not found after segmented body text search, apply fallback
                    if target_para is None:
                        reason = ""
                        if boundary_error == 'boundary_crossed':
                            reason = "Text crosses body/table boundary (not supported)"
                        elif boundary_error == 'row_boundary_crossed':
                            reason = "Text crosses table row boundary (Word doesn't support cross-row comments)"
                        else:
                            reason = f"Boundary error: {boundary_error}"

                        self._apply_fallback_comment(anchor_para, item, reason)
                        if self.verbose:
                            print(f"  [Boundary] {reason}")
                        return EditResult(
                            success=True,
                            item=item,
                            error_message=reason,
                            warning=True
                        )
                
                # Only proceed with fallback if target_para is still None
                # (boundary_error handling may have found a match)
                if target_para is None:
                    # For manual fix_action, text not found is expected (not an error)
                    if item.fix_action == 'manual':
                        self._apply_error_comment(anchor_para, item)
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
                # Check for cross-paragraph: only fallback if ACTUAL match spans multiple paragraphs
                if is_cross_paragraph:
                    # Get affected runs in the matched range
                    match_end = matched_start + len(violation_text)
                    affected = self._find_affected_runs(matched_runs_info, matched_start, match_end)

                    # Filter out boundary markers (paragraph and JSON boundaries)
                    real_runs = [r for r in affected
                                 if not r.get('is_para_boundary', False)
                                 and not r.get('is_json_boundary', False)
                                 and not r.get('is_json_escape', False)]

                    # Check if actual match spans multiple paragraphs
                    para_elems = set(r.get('para_elem') for r in real_runs if r.get('para_elem') is not None)

                    # Check if match spans multiple table rows (most restrictive)
                    if self._check_cross_row_boundary(real_runs):
                        # Cross-row delete not supported - fallback to comment
                        if self.verbose:
                            print(f"  [Cross-row] delete spans multiple rows, fallback to comment")
                        success_status = 'cross_row_fallback'
                    # Check if match spans multiple table cells (within same row)
                    elif self._check_cross_cell_boundary(real_runs):
                        # Cross-cell delete not supported - fallback to comment
                        if self.verbose:
                            print(f"  [Cross-cell] delete spans multiple cells, fallback to comment")
                        success_status = 'cross_cell_fallback'
                    elif len(para_elems) > 1:
                        # Actually spans multiple paragraphs - fallback to comment
                        if self.verbose:
                            print(f"  [Cross-paragraph] delete spans {len(para_elems)} paragraphs, fallback to comment")
                        success_status = 'cross_paragraph_fallback'
                    else:
                        # All content is in single paragraph - safe to delete
                        if real_runs:
                            target_para = real_runs[0].get('para_elem', target_para)
                        success_status = self._apply_delete(
                            target_para, violation_text,
                            item.violation_reason,
                            matched_runs_info, matched_start,
                            item_author
                        )
                else:
                    success_status = self._apply_delete(
                        target_para, violation_text,
                        item.violation_reason,
                        matched_runs_info, matched_start,
                        item_author
                    )
            elif item.fix_action == 'replace':
                # Check for cross-paragraph: only fallback if ACTUAL match spans multiple paragraphs
                if is_cross_paragraph:
                    # Get affected runs in the matched range
                    match_end = matched_start + len(violation_text)
                    affected = self._find_affected_runs(matched_runs_info, matched_start, match_end)

                    # Filter out boundary markers (paragraph and JSON boundaries)
                    real_runs = [r for r in affected
                                 if not r.get('is_para_boundary', False)
                                 and not r.get('is_json_boundary', False)
                                 and not r.get('is_json_escape', False)]

                    # Check if actual match spans multiple paragraphs
                    para_elems = set(r.get('para_elem') for r in real_runs if r.get('para_elem') is not None)

                    # Check if match spans multiple table rows (most restrictive)
                    if self._check_cross_row_boundary(real_runs):
                        # Cross-row replace not supported - fallback to comment
                        if self.verbose:
                            print(f"  [Cross-row] replace spans multiple rows, fallback to comment")
                        success_status = 'cross_row_fallback'
                    # Check if match spans multiple table cells (within same row)
                    elif self._check_cross_cell_boundary(real_runs):
                        # Try to extract single-cell edit
                        if self.verbose:
                            print(f"  [Cross-cell] Detected cross-cell match, trying single-cell extraction...")
                        
                        single_cell = self._try_extract_single_cell_edit(
                            violation_text, revised_text, affected, matched_start
                        )
                        
                        if single_cell:
                            # Successfully extracted single-cell edit - apply it
                            # Find the paragraph containing this cell
                            cell_para = None
                            for run in single_cell['cell_runs']:
                                if run.get('para_elem') is not None:
                                    cell_para = run['para_elem']
                                    break
                            
                            if cell_para is not None:
                                # Collect runs for this single cell only
                                cell_runs_info, cell_text = self._collect_runs_info_original(cell_para)
                                cell_pos = cell_text.find(single_cell['cell_violation'])
                                
                                if cell_pos != -1:
                                    # Apply replace to the single cell
                                    success_status = self._apply_replace(
                                        cell_para,
                                        single_cell['cell_violation'],
                                        single_cell['cell_revised'],
                                        item.violation_reason,
                                        cell_runs_info,
                                        cell_pos,
                                        item_author
                                    )
                                    
                                    if self.verbose and success_status == 'success':
                                        print(f"  [Single-cell] Successfully applied track change to single cell")
                                else:
                                    # Fallback if we can't find the text in the cell
                                    success_status = 'cross_cell_fallback'
                            else:
                                # Fallback if we can't find the paragraph
                                success_status = 'cross_cell_fallback'
                        else:
                            # Changes span multiple cells - fallback to comment
                            if self.verbose:
                                print(f"  [Cross-cell] Changes span multiple cells, fallback to comment")
                            success_status = 'cross_cell_fallback'
                    elif len(para_elems) > 1:
                        # Actually spans multiple paragraphs - fallback to comment
                        if self.verbose:
                            print(f"  [Cross-paragraph] replace spans {len(para_elems)} paragraphs, fallback to comment")
                        success_status = 'cross_paragraph_fallback'
                    else:
                        # All content is in single paragraph - safe to replace
                        if real_runs:
                            target_para = real_runs[0].get('para_elem', target_para)
                        success_status = self._apply_replace(
                            target_para, violation_text, revised_text,
                            item.violation_reason,
                            matched_runs_info, matched_start,
                            item_author
                        )
                else:
                    success_status = self._apply_replace(
                        target_para, violation_text, revised_text,
                        item.violation_reason,
                        matched_runs_info, matched_start,
                        item_author
                    )
            elif item.fix_action == 'manual':
                success_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph  # Pass cross-paragraph flag
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
            elif success_status == 'cross_paragraph_fallback':
                # Cross-paragraph delete/replace not supported - fallback to manual comment
                reason = "Cross-paragraph delete/replace not supported (Phase 1 limitation)"
                # Apply manual comment instead
                manual_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph
                )
                if manual_status == 'success':
                    if self.verbose:
                        print(f"  [Cross-paragraph] Applied comment instead")
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=reason,
                        warning=True
                    )
                else:
                    # Manual comment also failed - use fallback comment
                    self._apply_fallback_comment(target_para, item, reason)
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=f"{reason} (comment also failed)",
                        warning=True
                    )
            elif success_status == 'cross_row_fallback':
                # Cross-row delete/replace not supported - Word doesn't support cross-row comments
                reason = "Cross-row track change not supported (Word limitation)"
                # Use fallback comment (non-selected) since cross-row comments are not supported
                self._apply_fallback_comment(target_para, item, reason)
                if self.verbose:
                    print(f"  [Cross-row] Applied fallback comment")
                return EditResult(
                    success=True,
                    item=item,
                    error_message=reason,
                    warning=True
                )
            elif success_status == 'cross_cell_fallback':
                # Cross-cell delete/replace not supported - fallback to manual comment
                reason = "Cross-cell track change not supported, fallback to comment"
                # Apply manual comment instead (same row comment is supported)
                manual_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph
                )
                if manual_status == 'success':
                    if self.verbose:
                        print(f"  [Cross-cell] Applied comment instead")
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=reason,
                        warning=True
                    )
                else:
                    # Manual comment also failed - use fallback comment
                    self._apply_fallback_comment(target_para, item, reason)
                    return EditResult(
                        success=True,
                        item=item,
                        error_message=f"{reason} (comment also failed)",
                        warning=True
                    )
            elif success_status == 'fallback':
                # Fallback to comment annotation - mark as warning for all fix_actions
                reason = "Text not found in current document state"
                if item.fix_action == 'manual':
                    self._apply_error_comment(target_para, item)
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

        # 4. Set unified timestamp for all operations in this run
        self.operation_timestamp = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')

        # 5. Process each item
        for i, item in enumerate(self.edit_items):
            if self.verbose:
                print(f"[{i+1}/{len(self.edit_items)}] {item.fix_action}: "
                      f"{item.violation_text[:40]}...")
            
            result = self._process_item(item)
            self.results.append(result)
            
            if self.verbose:
                status = "✓" if result.success else "✗"
                if not result.success:
                    print(f"  [{status}]", end="")
                    print(f" {result.error_message}")
        
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
        if args.verbose:
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
