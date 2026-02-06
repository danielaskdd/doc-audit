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

# Set to a specific string from the origin content to WATCH for debuge
DEBUG_MARKER = ""

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
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

# Equation pattern for detecting LaTeX equation placeholders
# Matches: <equation>latex_content</equation>
EQUATION_PATTERN = re.compile(r'<equation>.*?</equation>', re.DOTALL)

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


def strip_auto_numbering_lines(text: str) -> Tuple[str, bool]:
    """
    Strip auto-numbering prefix from each line in a multi-line string.

    This is needed when violation_text spans multiple numbered list items
    (e.g., "e) ...\\nf) ...") while Word stores numbering separately.

    Args:
        text: Text that may contain multiple numbered lines

    Returns:
        (stripped_text, was_stripped)
    """
    if '\n' not in text:
        return strip_auto_numbering(text)

    lines = text.split('\n')
    new_lines = []
    was_stripped = False
    for line in lines:
        stripped, stripped_flag = strip_auto_numbering(line)
        if stripped_flag:
            was_stripped = True
        new_lines.append(stripped)
    return '\n'.join(new_lines), was_stripped


def build_numbering_variants(text: str) -> List[Tuple[str, str]]:
    """
    Build search variants with auto-numbering stripped.

    Returns list of (variant_text, mode) where mode is "prefix" or "lines".
    """
    variants: List[Tuple[str, str]] = []
    stripped_prefix, was_prefix = strip_auto_numbering(text)
    if was_prefix:
        variants.append((stripped_prefix, "prefix"))

    stripped_lines, was_lines = strip_auto_numbering_lines(text)
    if was_lines and stripped_lines != stripped_prefix:
        variants.append((stripped_lines, "lines"))

    return variants


def strip_numbering_by_mode(text: str, mode: Optional[str]) -> Tuple[str, bool]:
    """
    Strip numbering from text based on a specific mode.
    """
    if mode == "prefix":
        return strip_auto_numbering(text)
    if mode == "lines":
        return strip_auto_numbering_lines(text)
    return text, False


def strip_table_row_number_only(text: str) -> Tuple[str, bool]:
    """
    Replace only the first cell row number in table JSON text.

    This function strips row numbering from the first cell in each row,
    but does NOT strip auto-numbering inside other cells. This is needed
    when cell content like "a) ..." is actual text rather than Word numbering.

    Args:
        text: Text that may start with table row numbering pattern like '["1", '

    Returns:
        Tuple of (processed_text, was_stripped):
        - processed_text: Text with first cell row numbers stripped if found
        - was_stripped: True if any numbering was stripped, False otherwise
    """
    stripped_text = text.strip()
    if not stripped_text.startswith('['):
        return text, False

    was_modified = False

    def process_row(row: list) -> list:
        nonlocal was_modified
        if not row:
            return row

        first_cell = row[0]
        if isinstance(first_cell, str):
            first_cell_stripped = first_cell.strip()
            stripped_first, stripped_flag = strip_auto_numbering(first_cell_stripped)
            if stripped_flag:
                was_modified = True
                row = list(row)
                row[0] = stripped_first.strip()
            elif re.fullmatch(r'\d+(?:[.\d)）]+)?', first_cell_stripped):
                row = list(row)
                row[0] = ""
                was_modified = True

        return row

    rows = None
    mode = None  # "single", "multi", "full"

    try:
        parsed = json.loads(stripped_text)
        if isinstance(parsed, list):
            if parsed and isinstance(parsed[0], list):
                rows = parsed
                mode = "full"
            else:
                rows = [parsed]
                mode = "single"
    except json.JSONDecodeError:
        rows = None

    if rows is None and stripped_text.startswith('["'):
        try:
            parsed = json.loads(f'[{stripped_text}]')
            if isinstance(parsed, list) and (not parsed or isinstance(parsed[0], list)):
                rows = parsed
                mode = "multi"
        except json.JSONDecodeError:
            rows = None

    if rows is None:
        return text, False

    new_rows = []
    for row in rows:
        if isinstance(row, list):
            new_rows.append(process_row(row))
        else:
            new_rows.append(row)

    if mode == "full":
        row_string = ', '.join(json.dumps(row, ensure_ascii=False) for row in new_rows)
        if was_modified:
            return row_string, True
        if stripped_text != row_string:
            return row_string, True
        return text, False

    if not was_modified:
        return text, False
    if mode == "single":
        return json.dumps(new_rows[0], ensure_ascii=False), True
    return ', '.join(json.dumps(row, ensure_ascii=False) for row in new_rows), True

def strip_table_row_numbering(text: str) -> Tuple[str, bool]:
    """
    Replaces leading table row numbering and strips auto-numbering from all cells.
    
    During parse phase, Word auto-numbering shows as "1", "2", "3" etc.
    During apply phase, the same cells contain empty strings "" because auto-numbering
    is not stored in the cell content. This function:
    1. Replaces first cell number with empty string
    2. Strips auto-numbering prefix from ALL cell contents (e.g., "1. ", "a) ", "• ")
    3. Handles multi-paragraph cells by processing each paragraph separately
    
    Args:
        text: Text that may start with table row numbering pattern like '["1", '
        
    Returns:
        Tuple of (processed_text, was_stripped):
        - processed_text: Text with numbering stripped from all cells if found
        - was_stripped: True if any numbering was stripped, False otherwise
        
    Examples:
        '["1", "1. Intro", "2. Body"]' -> ('["", "Intro", "Body"]', True)
        '["", "a) First\\nb) Second"]' -> ('["", "First\\nSecond"]', True)
        '["content"]' -> ('["content"]', False)
        '["1", "A"], ["2", "B"]' -> ('["", "A"], ["", "B"]', True)
        '[[\"1\", \"A\"], [\"2\", \"B\"]]' -> ('[\"\", \"A\"], [\"\", \"B\"]', True)
    """
    stripped_text = text.strip()
    if not stripped_text.startswith('['):
        return text, False
    
    was_modified = False
    
    def process_row(row: list) -> list:
        nonlocal was_modified
        if not row:
            return row

        # 1. First cell: Replace row number with empty string
        first_cell = row[0]
        if isinstance(first_cell, str):
            first_cell_stripped = first_cell.strip()
            stripped_first, stripped_flag = strip_auto_numbering(first_cell_stripped)
            if stripped_flag:
                was_modified = True
                row = list(row)
                row[0] = stripped_first.strip()
            elif re.fullmatch(r'\d+(?:[.\d)）]+)?', first_cell_stripped):
                # Pure row number (e.g., "9", "9.", "9)", "9.1")
                row = list(row)
                row[0] = ""
                was_modified = True

        # 2. Strip auto-numbering from each cell and paragraph
        new_cells = []
        for cell in row:
            if isinstance(cell, str):
                # Split cell by newlines (handles both \n and literal \\n after JSON decode)
                # After json.loads, \\n in JSON becomes \n in string
                paragraphs = cell.split('\n')

                # Process each paragraph
                new_paragraphs = []
                for para in paragraphs:
                    stripped, stripped_flag = strip_auto_numbering(para)
                    if stripped_flag:
                        was_modified = True
                    new_paragraphs.append(stripped)

                # Rejoin with newline
                new_cells.append('\n'.join(new_paragraphs))
            else:
                new_cells.append(cell)

        return new_cells

    rows = None
    mode = None  # "single", "multi", "full"

    # 1) Try direct JSON parse (single row or full table array)
    try:
        parsed = json.loads(stripped_text)
        if isinstance(parsed, list):
            if parsed and isinstance(parsed[0], list):
                rows = parsed
                mode = "full"
            else:
                rows = [parsed]
                mode = "single"
    except json.JSONDecodeError:
        rows = None

    # 2) Try wrapping multi-row JSON (e.g., '["1",...], ["2",...]')
    if rows is None and stripped_text.startswith('["'):
        try:
            parsed = json.loads(f'[{stripped_text}]')
            if isinstance(parsed, list) and (not parsed or isinstance(parsed[0], list)):
                rows = parsed
                mode = "multi"
        except json.JSONDecodeError:
            rows = None

    if rows is None:
        # Fallback: return original text if JSON parsing fails
        return text, False

    new_rows = []
    for row in rows:
        if isinstance(row, list):
            new_rows.append(process_row(row))
        else:
            new_rows.append(row)

    if mode == "full":
        # Convert to row-string to match table_text format
        row_string = ', '.join(json.dumps(row, ensure_ascii=False) for row in new_rows)
        if was_modified:
            return row_string, True
        # If no numbering was stripped, still normalize full-table JSON for matching
        if stripped_text != row_string:
            return row_string, True
        return text, False

    if not was_modified:
        return text, False
    if mode == "single":
        return json.dumps(new_rows[0], ensure_ascii=False), True
    # mode == "multi"
    return ', '.join(json.dumps(row, ensure_ascii=False) for row in new_rows), True


def normalize_table_json(text: str) -> str:
    """
    Normalize table JSON by removing duplicate brackets at boundaries.
    
    LLM may incorrectly include extra brackets when referencing table rows:
    - First row: '[["...' instead of '["...'
    - Last row: '..."]]' instead of '..."]'
    
    This function cleans up these artifacts by removing duplicate brackets.
    
    Args:
        text: Table JSON text that may have duplicate brackets
        
    Returns:
        Normalized text with duplicate brackets removed
        
    Examples:
        '[["cell1", "cell2"]' -> '["cell1", "cell2"]'
        '["cell1", "cell2"]]' -> '["cell1", "cell2"]'
        '[["cell1", "cell2"]]' -> '["cell1", "cell2"]'
        '["cell1", "cell2"]' -> '["cell1", "cell2"]' (no change)
    """
    if not text.startswith('["'):
        return text
    
    result = text
    # Remove leading duplicate bracket
    if result.startswith('[["'):
        result = result[1:]
    # Remove trailing duplicate bracket
    if result.endswith('"]]'):
        result = result[:-1]
    
    return result

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

        # Paragraph order cache (initialized in apply())
        self._para_list: List = []
        self._para_order: Dict[int, int] = {}
        self._para_id_list: List[Optional[str]] = []

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

    def _find_last_para_with_id_in_table(self, table_elem, uuid_end: str, start_para=None):
        """
        Find the last paragraph in a table (optionally after start_para) with the given paraId.

        This is used to handle vertical-merge cases where Word duplicates paraId
        across multiple rows.
        """
        if table_elem is None or not uuid_end:
            return None

        found_start = start_para is None
        end_para = None

        for para in table_elem.iter(f'{{{NS["w"]}}}p'):
            if not found_start:
                if para is start_para:
                    found_start = True
                else:
                    continue
            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
            if para_id == uuid_end:
                end_para = para

        return end_para

    def _resolve_end_para(self, start_para, uuid_end: str):
        """
        Resolve the end paragraph element for a uuid range.

        - If start is in a table, pick the LAST matching paraId in the same table
          after start (to handle vertical merge duplicate paraIds).
        - Otherwise, pick the FIRST matching paraId after start in document order.
        """
        if start_para is None or not uuid_end:
            return None

        if self._is_paragraph_in_table(start_para):
            table = self._find_ancestor_table(start_para)
            end_para = self._find_last_para_with_id_in_table(table, uuid_end, start_para)
            if end_para is not None:
                return end_para

        all_paras = self._xpath(self.body_elem, './/w:p')
        try:
            start_index = all_paras.index(start_para)
        except ValueError:
            start_index = 0

        for para in all_paras[start_index:]:
            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
            if para_id == uuid_end:
                return para

        return None
    
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
        end_para = self._resolve_end_para(start_node, uuid_end)

        try:
            start_index = all_paras.index(start_node)
            for p in all_paras[start_index:]:
                yield p
                # Stop after reaching the end boundary (inclusive)
                if end_para is not None:
                    if p is end_para:
                        return
                    continue
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

    def _strip_runs_whitespace(self, runs: List[Dict]) -> List[Dict]:
        """
        Strip leading/trailing whitespace across a paragraph's runs.

        This mirrors table_extractor.py behavior where each paragraph is
        `para_text.strip()` before being joined into cell text. It removes
        leading/trailing whitespace (including tabs) across the entire paragraph
        while preserving internal whitespace.

        Args:
            runs: List of run dicts from _collect_runs_info_original

        Returns:
            List of runs with adjusted 'text' values and empty runs removed
        """
        if not runs:
            return []

        combined = ''.join(r.get('text', '') for r in runs)
        if not combined:
            return []

        lead = len(combined) - len(combined.lstrip())
        trail = len(combined) - len(combined.rstrip())

        if lead == 0 and trail == 0:
            return runs

        # Trim leading whitespace
        idx = 0
        while idx < len(runs) and lead > 0:
            text = runs[idx].get('text', '')
            if not text:
                idx += 1
                continue
            if lead >= len(text):
                runs[idx]['text'] = ''
                lead -= len(text)
                idx += 1
            else:
                runs[idx]['text'] = text[lead:]
                lead = 0

        # Trim trailing whitespace
        idx = len(runs) - 1
        while idx >= 0 and trail > 0:
            text = runs[idx].get('text', '')
            if not text:
                idx -= 1
                continue
            if trail >= len(text):
                runs[idx]['text'] = ''
                trail -= len(text)
                idx -= 1
            else:
                runs[idx]['text'] = text[:-trail]
                trail = 0

        return [r for r in runs if r.get('text')]

    def _normalize_table_text_for_search(self, runs_info: List[Dict]) -> Tuple[str, List[int]]:
        """
        Build a normalized table text for search and a mapping back to original positions.

        This mirrors table_extractor.py behavior by stripping leading/trailing whitespace
        from each paragraph within a cell. It returns:
        - normalized_text: text with per-paragraph whitespace stripped
        - norm_to_orig: list mapping each normalized char index -> original char index

        The mapping allows translating match positions from normalized text
        back to the original combined table text for accurate run selection.
        """
        normalized_chars: List[str] = []
        norm_to_orig: List[int] = []
        para_chars: List[Tuple[str, int]] = []  # (char, orig_index)
        orig_pos = 0

        def flush_paragraph():
            nonlocal para_chars
            if not para_chars:
                return
            para_text = ''.join(ch for ch, _ in para_chars)
            if not para_text:
                para_chars = []
                return
            # Strip leading/trailing whitespace like para_text.strip()
            start = 0
            end = len(para_text)
            while start < end and para_text[start].isspace():
                start += 1
            while end > start and para_text[end - 1].isspace():
                end -= 1
            for ch, orig_idx in para_chars[start:end]:
                normalized_chars.append(ch)
                norm_to_orig.append(orig_idx)
            para_chars = []

        def iter_original_with_escaped_indices(original_text: str) -> List[Tuple[str, int]]:
            """
            Yield (char, escaped_index_start) for original_text.

            This maps original characters to positions in the JSON-escaped
            string stored in runs_info['text'].
            """
            escaped_pos = 0
            for ch in original_text:
                escaped_chunk = json_escape(ch)
                yield ch, escaped_pos
                escaped_pos += len(escaped_chunk)

        for run in runs_info:
            text = run.get('text', '')
            run_len = len(text)

            if run.get('is_para_boundary'):
                flush_paragraph()
                for i, ch in enumerate(text):
                    normalized_chars.append(ch)
                    norm_to_orig.append(orig_pos + i)
                orig_pos += run_len
                continue

            if run.get('is_json_boundary') or run.get('is_cell_boundary') or run.get('is_row_boundary'):
                flush_paragraph()
                for i, ch in enumerate(text):
                    normalized_chars.append(ch)
                    norm_to_orig.append(orig_pos + i)
                orig_pos += run_len
                continue

            # Content run (part of a paragraph)
            original_text = run.get('original_text', text)
            if 'original_text' in run:
                for ch, offset in iter_original_with_escaped_indices(original_text):
                    para_chars.append((ch, orig_pos + offset))
            else:
                for i, ch in enumerate(original_text):
                    para_chars.append((ch, orig_pos + i))
            orig_pos += run_len

        # Flush remaining paragraph at end
        flush_paragraph()

        return ''.join(normalized_chars), norm_to_orig

    def _normalize_body_runs_for_search(self, runs_info: List[Dict]) -> Tuple[str, List[int], str]:
        """
        Normalize body runs for search by stripping trailing whitespace on each line.

        Returns:
            normalized_text: text with per-line rstrip applied
            norm_to_orig: list mapping normalized index -> original index
            combined_text: original combined text
        """
        normalized_chars: List[str] = []
        norm_to_orig: List[int] = []
        line_chars: List[Tuple[str, int]] = []  # (char, orig_index)
        combined_parts: List[str] = []
        orig_pos = 0

        def flush_line():
            nonlocal line_chars
            if not line_chars:
                return
            end = len(line_chars)
            while end > 0 and line_chars[end - 1][0].isspace():
                end -= 1
            for ch, idx in line_chars[:end]:
                normalized_chars.append(ch)
                norm_to_orig.append(idx)
            line_chars = []

        for run in runs_info:
            text = run.get('text', '')
            combined_parts.append(text)
            for ch in text:
                if ch == '\n':
                    flush_line()
                    normalized_chars.append(ch)
                    norm_to_orig.append(orig_pos)
                else:
                    line_chars.append((ch, orig_pos))
                orig_pos += 1

        flush_line()

        return ''.join(normalized_chars), norm_to_orig, ''.join(combined_parts)

    def _find_in_runs_with_normalization(self, runs_info: List[Dict], search_text: str) -> Tuple[int, Optional[str]]:
        """
        Find search_text in runs_info with trailing-whitespace normalization fallback.

        Returns:
            (match_start, matched_text_override)
            - match_start: position in original combined text, or -1 if not found
            - matched_text_override: original substring that matches the normalized hit,
              or None if a direct match was found without normalization.
        """
        combined_text = ''.join(r.get('text', '') for r in runs_info)
        if not search_text:
            return -1, None

        pos_raw = combined_text.find(search_text)
        if pos_raw != -1:
            return pos_raw, None

        normalized_text, norm_to_orig, _ = self._normalize_body_runs_for_search(runs_info)
        pos_norm = normalized_text.find(search_text)
        if pos_norm == -1:
            return -1, None

        norm_end = pos_norm + len(search_text) - 1
        if norm_end >= len(norm_to_orig):
            return -1, None

        orig_start = norm_to_orig[pos_norm]
        orig_end = norm_to_orig[norm_end] + 1
        return orig_start, combined_text[orig_start:orig_end]

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
            Tuple of (target_para, runs_info, match_start, matched_text, strip_mode) or None if not found
            - matched_text: The actual text that was matched (may be normalized if fallback was used)
            - strip_mode: "prefix" | "lines" | None (if auto-numbering was stripped)
        
        Note:
            The in_range gate has been removed because _find_tables_in_range already
            constrains the table to the uuid range. The anchor paragraph (start_para)
            is often outside the table (e.g., heading before table), so checking for
            it would cause all cells to be skipped.
        """
        end_para = self._find_last_para_with_id_in_table(table_elem, uuid_end)

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
                para_runs = self._strip_runs_whitespace(para_runs)
                for run in para_runs:
                    run_copy = dict(run)
                    run_copy['para_elem'] = para
                    run_copy['cell_elem'] = tc
                    run_copy['start'] = run['start'] + pos
                    run_copy['end'] = run['end'] + pos
                    cell_runs.append(run_copy)
                
                if cell_runs:
                    pos = cell_runs[-1]['end']
                
                # Stop at uuid_end (last occurrence in table if duplicated)
                if end_para is not None:
                    if para is end_para:
                        break
                else:
                    if para_id == uuid_end:
                        break
            
            if not cell_runs:
                continue
            
            # Build cell text for debug logging
            cell_text = ''.join(r['text'] for r in cell_runs)

            # Search within this cell only (prevents cross-cell matching)
            # Use normalization + mapping to handle trailing whitespace safely
            # Build search attempts (original + numbering-stripped + newline-normalized)
            search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
            search_attempts.extend(build_numbering_variants(violation_text))

            if '\\n' in violation_text:
                newline_text = violation_text.replace('\\n', '\n')
                search_attempts.append((newline_text, None))
                search_attempts.extend(build_numbering_variants(newline_text))

            # De-duplicate search attempts while preserving order
            seen = set()
            deduped_attempts: List[Tuple[str, Optional[str]]] = []
            for text, mode in search_attempts:
                key = (text, mode)
                if key in seen:
                    continue
                seen.add(key)
                deduped_attempts.append((text, mode))

            match_pos = -1
            matched_text = violation_text
            matched_override = None
            matched_strip_mode: Optional[str] = None

            for search_text, strip_mode in deduped_attempts:
                match_pos, matched_override = self._find_in_runs_with_normalization(
                    cell_runs, search_text
                )
                if match_pos != -1:
                    matched_text = matched_override or search_text
                    matched_strip_mode = strip_mode
                    if self.verbose and '\\n' in violation_text and search_text != violation_text:
                        print(f"  [Fallback] Matched after converting \\\\n or stripping numbering")
                    break
            
            # Debug: log snippet from marker on non-JSON match failure
            if match_pos == -1:
                if DEBUG_MARKER:
                    marker_idx = cell_text.find(DEBUG_MARKER)
                    if DEBUG_MARKER and marker_idx != -1:
                        snippet_len = len(DEBUG_MARKER) + 60
                        snippet = cell_text[marker_idx:marker_idx + snippet_len]
                        print(f"  [DEBUG] Non-JSON cell content from marker: {repr(snippet)}")

            if match_pos != -1:
                # Found! Return match info including the matched text
                if self.verbose:
                    print(f"  [Success] Found in cell (raw text mode): '{cell_text[:50]}...'")
                return (cell_first_para, cell_runs, match_pos, matched_text, matched_strip_mode)
        
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
        """
        # 1. Detect if start paragraph is in a table
        start_in_table = self._is_paragraph_in_table(start_para)
        start_table = self._find_ancestor_table(start_para) if start_in_table else None

        # 2. Find end paragraph and check its context
        end_para = self._resolve_end_para(start_para, uuid_end)
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
            return self._collect_runs_info_in_table(start_para, uuid_end, start_table, end_para=end_para)
        else:
            return self._collect_runs_info_in_body(start_para, uuid_end, end_para=end_para)

    def _collect_runs_info_in_body(self, start_para, uuid_end: str, end_para=None) -> Tuple[List[Dict], str, bool, Optional[str]]:
        """
        Collect run info across paragraphs in document body (not in table).
        Paragraph boundaries are represented as '\\n'.

        Args:
            start_para: Starting paragraph element
            uuid_end: End boundary paraId (inclusive)

        Returns:
            Tuple of (runs_info, combined_text, is_cross_paragraph, boundary_error)
        """
        if end_para is None:
            end_para = self._resolve_end_para(start_para, uuid_end)
        runs_info = []
        pos = 0
        para_count = 0

        for para in self._iter_paragraphs_in_range(start_para, uuid_end):
            para_count += 1

            # Collect runs for this paragraph
            para_runs, para_text = self._collect_runs_info_original(para)
            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')

            # Skip empty paragraphs to match parse_document.py behavior
            if not para_text.strip():
                continue

            # Add paragraph element reference to each run
            for run in para_runs:
                run['para_elem'] = para
                run['host_para_elem'] = para
                run['host_para_id'] = para_id
                run['start'] += pos
                run['end'] += pos
                runs_info.append(run)

            pos += len(para_text)

            # Add paragraph boundary (except after last paragraph)
            is_last_para = False
            if end_para is not None:
                is_last_para = para is end_para
            else:
                para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                is_last_para = (para_id == uuid_end)
            if not is_last_para:
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

    def _collect_runs_info_in_table(self, start_para, uuid_end: str, table_elem, end_para=None) -> Tuple[List[Dict], str, bool, Optional[str]]:
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
        if end_para is None:
            end_para = self._resolve_end_para(start_para, uuid_end)

        def is_end_para(para) -> bool:
            if end_para is not None:
                return para is end_para
            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
            return para_id == uuid_end

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
        cols_in_range = set()  # Track which columns are within uuid range
        
        # Iterate rows (w:tr elements)
        for tr in table_elem.findall(f'{{{NS["w"]}}}tr'):
            if reached_end:
                break
            
            grid_col = 0
            row_data = [None] * num_cols  # Pre-fill with None to match TableExtractor behavior
            
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
                    if in_range and not reached_end and para_id:
                        cell_in_range = True
                        last_para = para
                        if is_end_para(para):
                            reached_end = True
                            break
                
                # Collect cell content if in range
                cell_runs = []
                cell_text = ''
                
                if cell_in_range:
                    # Mark all columns spanned by this cell as in range
                    cols_in_range.update(range(grid_col, grid_col + grid_span))
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
                            
                            para_runs, para_orig_text = self._collect_runs_info_original(para)
                            # Strip paragraph text to match table_extractor.py behavior (line 144/163)
                            # table_extractor.py calls para_text.strip() before checking if empty
                            para_orig_text = para_orig_text.strip()
                            
                            # Skip empty paragraphs (match table_extractor.py behavior at line 144/163)
                            # table_extractor.py strips each paragraph and skips if empty
                            if not para_orig_text:
                                continue
                            
                            para_runs = self._strip_runs_whitespace(para_runs)
                            for run in para_runs:
                                original_text = run['text']
                                escaped_text = json_escape(original_text)
                                cell_runs.append({
                                    **run,
                                    'original_text': original_text,
                                    'text': escaped_text,
                                    'para_elem': para,
                                    'host_para_elem': para,
                                    'host_para_id': para_id,
                                    'cell_elem': tc,
                                    'row_elem': tr
                                })
                                cell_text += escaped_text
                            
                            if is_end_para(para):
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
                                # Override host para to current row cell para
                                if cell_paras:
                                    host_para = cell_paras[0]
                                    host_para_id = host_para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                                    run['host_para_elem'] = host_para
                                    run['host_para_id'] = host_para_id
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
                            
                            para_runs, para_orig_text = self._collect_runs_info_original(para)
                            # Strip paragraph text to match table_extractor.py behavior (line 144/163)
                            # table_extractor.py calls para_text.strip() before checking if empty
                            para_orig_text = para_orig_text.strip()
                            
                            # Skip empty paragraphs (match table_extractor.py behavior at line 144/163)
                            # table_extractor.py strips each paragraph and skips if empty
                            if not para_orig_text:
                                continue
                            
                            para_runs = self._strip_runs_whitespace(para_runs)
                            for run in para_runs:
                                original_text = run['text']
                                escaped_text = json_escape(original_text)
                                cell_runs.append({
                                    **run,
                                    'original_text': original_text,
                                    'text': escaped_text,
                                    'para_elem': para,
                                    'host_para_elem': para,
                                    'host_para_id': para_id,
                                    'cell_elem': tc,
                                    'row_elem': tr
                                })
                                cell_text += escaped_text
                            
                            if is_end_para(para):
                                break
                        
                        # Check if empty cell inherits from vMerge (TableExtractor logic)
                        if not cell_text and grid_col in vmerge_content:
                            cell_runs = [dict(r) for r in vmerge_content[grid_col]['runs']]
                            cell_text = vmerge_content[grid_col]['text']
                            for run in cell_runs:
                                run['row_elem'] = tr
                                # Override host para to current row cell para
                                if cell_paras:
                                    host_para = cell_paras[0]
                                    host_para_id = host_para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                                    run['host_para_elem'] = host_para
                                    run['host_para_id'] = host_para_id
                        elif cell_text:
                            # Non-empty normal cell ends vMerge region
                            vmerge_content.pop(grid_col, None)
                
                # Store cell data at grid_col position (matching TableExtractor behavior)
                # gridSpan cells only occupy the starting position, skipped columns remain None
                if grid_col < num_cols and (cell_in_range or cell_runs):
                    row_data[grid_col] = (cell_runs, cell_text)
                
                # Advance grid column by gridSpan
                grid_col += grid_span
                
                # Check if we've reached the end - stop immediately when uuid_end is found
                if cell_in_range:
                    for para in cell_paras:
                        if is_end_para(para):
                            reached_end = True
                            break
                
                # If we've reached the end, stop processing subsequent cells in this row
                if reached_end:
                    break
            
            # After processing all cells in row, add to runs_info if row is in range
            if in_range and any(cell_tuple is not None and cell_tuple[0] for cell_tuple in row_data):
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
                
                # Output cells based on cols_in_range to preserve gridSpan behavior
                # For each column in range:
                #   - If row_data[col] has data: output cell content
                #   - If row_data[col] is None but col in cols_in_range: output "" (gridSpan placeholder)
                #   - If col not in cols_in_range: skip (out of range)
                output_col_count = 0
                for col_idx in range(num_cols):
                    if col_idx not in cols_in_range:
                        continue  # Skip columns outside uuid range
                    
                    if output_col_count > 0:
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
                    
                    if row_data[col_idx] is not None:
                        # Cell has content - output runs
                        cell_runs, cell_text = row_data[col_idx]
                        for run in cell_runs:
                            run['start'] = pos
                            run['end'] = pos + len(run['text'])
                            runs_info.append(run)
                            pos += len(run['text'])
                    # else: gridSpan placeholder - output nothing (empty string between ", ")
                    
                    output_col_count += 1
            
        
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
        end_para = self._resolve_end_para(start_para, uuid_end)

        def is_end_para(para) -> bool:
            if end_para is not None:
                return para is end_para
            para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
            return para_id == uuid_end

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
            para_runs = self._strip_runs_whitespace(para_runs)

            for run in para_runs:
                original_text = run['text']
                escaped_text = json_escape(original_text)

                run['original_text'] = original_text
                run['text'] = escaped_text
                run['para_elem'] = para
                run['host_para_elem'] = para
                run['host_para_id'] = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                run['cell_elem'] = cell
                run['row_elem'] = row
                run['start'] = pos
                run['end'] = pos + len(escaped_text)
                runs_info.append(run)

                pos += len(escaped_text)

            # Check if we've reached the end
            if is_end_para(para):
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

    def _extract_text_in_range_from_table(self, table_elem, uuid_start: str, uuid_end: str) -> str:
        """
        Extract table text between uuid_start and uuid_end in JSON row format.

        Returns a string like:
          ["cell1", "cell2"], ["cell3", "cell4"]

        This mirrors parse_document-style row formatting and handles gridSpan/vMerge:
        - gridSpan: repeats cell content across spanned columns
        - vMerge restart/continue: repeats content only when restart is within range
        """
        rows_text = []
        in_range = False
        reached_end = False
        vmerge_content = {}  # {grid_col: text}
        end_para = self._find_last_para_with_id_in_table(table_elem, uuid_end)

        for tr in table_elem.findall(f'{{{NS["w"]}}}tr'):
            if reached_end:
                break
            row_cells = []
            row_in_range = False
            grid_col = 0

            for tc in tr.findall(f'{{{NS["w"]}}}tc'):
                if reached_end:
                    break

                tcPr = tc.find(f'{{{NS["w"]}}}tcPr')
                grid_span, vmerge_type = self._get_cell_merge_properties(tcPr)
                cell_paras = tc.findall(f'{{{NS["w"]}}}p')

                cell_text_parts = []
                cell_has_range = False

                for para in cell_paras:
                    para_id = para.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
                    if para_id == uuid_start:
                        in_range = True

                    if in_range:
                        cell_has_range = True
                        _, para_text = self._collect_runs_info_original(para)
                        if para_text:
                            cell_text_parts.append(para_text)

                    if in_range:
                        if end_para is not None:
                            if para is end_para:
                                reached_end = True
                                break
                        elif para_id == uuid_end:
                            reached_end = True
                            break

                cell_text = '\n'.join(cell_text_parts).replace('\x07', '')

                if vmerge_type == 'restart':
                    if cell_has_range:
                        for col in range(grid_col, grid_col + grid_span):
                            vmerge_content[col] = cell_text
                elif vmerge_type == 'continue':
                    if grid_col in vmerge_content:
                        cell_text = vmerge_content.get(grid_col, '')
                else:
                    if not cell_text and grid_col in vmerge_content:
                        cell_text = vmerge_content.get(grid_col, '')
                    elif cell_text:
                        for col in range(grid_col, grid_col + grid_span):
                            vmerge_content.pop(col, None)

                if in_range or cell_has_range:
                    row_in_range = True
                    escaped = json_escape(cell_text)
                    for _ in range(grid_span):
                        row_cells.append(escaped)

                grid_col += grid_span

            if row_in_range:
                rows_text.append('["' + '", "'.join(row_cells) + '"]')

        return ', '.join(rows_text)

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
    
    def _group_runs_by_cell(self, runs: List[Dict]) -> Dict[Tuple, List[Dict]]:
        """
        Group runs by (para_elem, cell_elem) pairs.
        
        This allows processing runs cell-by-cell for multi-cell operations
        like delete/replace across table cells.
        
        IMPORTANT: This method assumes runs are already in document order
        (as provided by _collect_runs_info_* methods). The grouping preserves
        this order within each cell group.
        
        Args:
            runs: List of run info dicts (must be in document order)
        
        Returns:
            Dict mapping (para_elem, cell_elem) -> list of runs in that cell
        """
        groups = {}
        for run in runs:
            para = run.get('para_elem')
            cell = run.get('cell_elem')
            key = (id(para) if para is not None else None, 
                   id(cell) if cell is not None else None)
            if key not in groups:
                groups[key] = []
            groups[key].append(run)
        return groups

    def _group_runs_by_paragraph(self, runs: List[Dict]) -> List[Dict]:
        """
        Group runs by paragraph in document order.

        Args:
            runs: List of run info dicts (must be in document order)

        Returns:
            List of dicts: [{'para_elem': para, 'runs': [...]}]
        """
        groups = {}
        order = []
        for run in runs:
            para = run.get('para_elem')
            if para is None:
                continue
            key = id(para)
            if key not in groups:
                groups[key] = {'para_elem': para, 'runs': []}
                order.append(key)
            groups[key]['runs'].append(run)
        return [groups[k] for k in order]

    def _init_para_order(self):
        """Initialize paragraph order cache for document-order comparisons."""
        if self.body_elem is None:
            return
        self._para_list = list(self._xpath(self.body_elem, './/w:p'))
        self._para_order = {id(p): i for i, p in enumerate(self._para_list)}
        self._para_id_list = [
            p.get('{http://schemas.microsoft.com/office/word/2010/wordml}paraId')
            for p in self._para_list
        ]

    def _get_run_doc_key(self, run_elem, para_elem,
                         para_order: Optional[Dict[int, int]] = None) -> Optional[Tuple[int, int]]:
        """
        Get a stable document-order key for a run element.

        Returns (para_index, run_index_in_para) or None if unavailable.
        This is used to detect inverted ordering when runs are reused
        across rows (e.g., vMerge continue cells).
        """
        if run_elem is None or para_elem is None:
            return None
        if para_order is None:
            para_order = self._para_order
        if not para_order:
            self._init_para_order()
            para_order = self._para_order
        para_idx = para_order.get(id(para_elem))
        if para_idx is None:
            # Refresh cache in case body_elem was replaced in tests
            self._init_para_order()
            para_order = self._para_order
            para_idx = para_order.get(id(para_elem))
        if para_idx is None:
            return None
        try:
            run_list = list(para_elem.iter(f'{{{NS["w"]}}}r'))
            run_idx = run_list.index(run_elem)
        except ValueError:
            run_idx = 0
        return (para_idx, run_idx)
    
    def _is_multi_row_json(self, text: str) -> bool:
        """
        Check if text is multi-row JSON format.
        
        Multi-row JSON format: '["cell1", "cell2"], ["cell3", "cell4"]'
        Single-row JSON format: '["cell1", "cell2"]'
        
        Args:
            text: Text to check
        
        Returns:
            True if text contains row boundary marker '"], ["'
        """
        return text.startswith('["') and '"], ["' in text

    def _apply_delete_multi_cell(self, real_runs: List[Dict], violation_text: str,
                                  violation_reason: str, author: str, match_start: int,
                                  item: EditItem) -> str:
        """
        Apply delete operation across multiple cells/rows.
        
        Strategy: Delete content in each cell independently by wrapping runs
        with <w:del> tags, properly handling partial run deletion when the
        match doesn't align to run boundaries.
        
        Args:
            real_runs: List of real runs (boundaries filtered out)
            violation_text: Text being deleted (may be JSON format)
            violation_reason: Reason for deletion
            author: Track change author
            match_start: Absolute position where violation_text starts
            item: Original EditItem (for fallback comment)
        
        Returns:
            'success' if at least one cell deleted, 'fallback' if all failed
        """
        if not real_runs:
            return 'fallback'

        match_end = match_start + len(violation_text)

        # Group runs by cell
        cell_groups = self._group_runs_by_cell(real_runs)

        # Track success/failure for each cell
        success_count = 0
        failed_cells = []  # List of (para_elem, cell_violation, error_reason)
        first_success_para = None

        # Process each cell independently
        for (para_id, cell_id), cell_runs in cell_groups.items():
            if not cell_runs:
                continue

            # De-duplicate runs by their underlying elem reference
            # This prevents duplicates when vMerge='continue' cells copy runs from restart cells
            seen_elems = set()
            unique_cell_runs = []
            for run in cell_runs:
                elem_id = id(run.get('elem'))
                if elem_id not in seen_elems:
                    seen_elems.add(elem_id)
                    unique_cell_runs.append(run)
            cell_runs = unique_cell_runs

            if not cell_runs:
                continue

            # Get the paragraph element from first run
            para_elem = cell_runs[0].get('para_elem')
            if para_elem is None:
                failed_cells.append((None, "", "No paragraph found"))
                continue

            # Find cell boundaries in the match range
            cell_start = min(r['start'] for r in cell_runs)
            cell_end = max(r['end'] for r in cell_runs)

            # Calculate intersection with match range
            del_start_in_cell = max(cell_start, match_start)
            del_end_in_cell = min(cell_end, match_end)

            if del_start_in_cell >= del_end_in_cell:
                # No overlap with this cell
                continue

            # Identify affected runs in this cell
            affected_cell_runs = [r for r in cell_runs
                                  if r['end'] > del_start_in_cell and r['start'] < del_end_in_cell]

            if not affected_cell_runs:
                failed_cells.append((para_elem, "", "No affected runs found"))
                continue

            # Extract cell violation text for error reporting
            cell_violation_parts = []
            first_run = affected_cell_runs[0]
            last_run = affected_cell_runs[-1]
            before_offset = self._translate_escaped_offset(
                first_run, max(0, del_start_in_cell - first_run['start']))
            after_offset = self._translate_escaped_offset(
                last_run, max(0, del_end_in_cell - last_run['start']))

            for run in affected_cell_runs:
                if run.get('is_drawing'):
                    cell_violation_parts.append(run['text'])
                    continue
                orig_text = self._get_run_original_text(run)
                if run is first_run and run is last_run:
                    cell_violation_parts.append(orig_text[before_offset:after_offset])
                elif run is first_run:
                    cell_violation_parts.append(orig_text[before_offset:])
                elif run is last_run:
                    cell_violation_parts.append(orig_text[:after_offset])
                else:
                    cell_violation_parts.append(orig_text)
            cell_violation = ''.join(cell_violation_parts)

            # Check if any affected run overlaps with existing revisions
            if self._check_overlap_with_revisions(affected_cell_runs):
                failed_cells.append((para_elem, cell_violation, "Overlaps with existing revision"))
                continue

            if not cell_violation:
                failed_cells.append((para_elem, "", "No text to delete"))
                continue

            para_groups = self._group_runs_by_paragraph(affected_cell_runs)
            if not para_groups:
                failed_cells.append((para_elem, cell_violation, "No paragraph groups found"))
                continue

            prepared = self._prepare_deletion_items(
                para_groups, del_start_in_cell, del_end_in_cell
            )
            if not prepared:
                failed_cells.append((para_elem, cell_violation, "No paragraphs prepared for deletion"))
                continue

            shared_change_id = self._get_next_change_id()
            success_paragraphs = self._delete_paragraphs_in_unit(
                prepared, shared_change_id, author, comment_id=None
            )

            if success_paragraphs == 0:
                failed_cells.append((para_elem, cell_violation, "Failed to replace runs"))
                continue

            success_count += 1
            if first_success_para is None:
                first_success_para = para_elem
        
        # Handle results based on success/failure counts
        if success_count > 0:
            # At least one cell deleted - add overall comment
            if first_success_para is not None:
                comment_id = self.next_comment_id
                self.next_comment_id += 1
                
                # Insert commentReference at end of first successful paragraph
                ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                    <w:commentReference w:id="{comment_id}"/>
                </w:r>'''
                first_success_para.append(etree.fromstring(ref_xml))
                
                # Record comment with -R suffix author
                self.comments.append({
                    'id': comment_id,
                    'text': violation_reason,
                    'author': f"{author}-R"
                })
            
            # Add individual comments for failed cells
            for para_elem, cell_violation, error_reason in failed_cells:
                if para_elem is not None:
                    self._apply_cell_fallback_comment(
                        para_elem,
                        cell_violation,
                        "",  # No revised text for delete operation
                        error_reason,
                        item
                    )
            
            if self.verbose:
                if failed_cells:
                    print(f"  [Multi-cell delete] Partial success: {success_count}/{len(cell_groups)} cells processed, {len(failed_cells)} failed")
                else:
                    print(f"  [Multi-cell delete] Successfully deleted from all {len(cell_groups)} cells")
            elif failed_cells:
                # Summary logging for partial failures in non-verbose mode
                print(f"  [Warning] {len(failed_cells)} cell(s) failed during delete operation")
            
            return 'success'
        else:
            # All cells failed - fallback to overall comment
            return 'fallback'
    
    def _try_extract_multi_cell_edits(self, violation_text: str, revised_text: str,
                                        affected_runs: List[Dict], match_start: int) -> Optional[List[Dict]]:
        """
        Try to extract multi-cell edits when text spans multiple cells.
        
        This method analyzes the diff to determine if changes are distributed across
        multiple cells, with each change staying within its own cell boundary.
        
        Args:
            violation_text: Original violation text (may span multiple cells)
            revised_text: Revised text (may span multiple cells)
            affected_runs: List of run info dicts that would be affected
            match_start: Absolute position where violation_text starts in document
        
        Returns:
            List of dicts, each with keys: 'cell_violation', 'cell_revised', 'cell_elem', 'cell_runs'
            Or None if any change crosses cell boundaries
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
        
        # Find which cell each change affects
        # Map: cell_id -> list of change indices
        cell_to_changes = {}
        
        base_offset = match_start
        
        for change_idx, (change_start_rel, change_end_rel) in enumerate(change_positions):
            # Translate relative positions to absolute positions
            change_start = base_offset + change_start_rel
            change_end = base_offset + change_end_rel
            
            is_insert = (change_start == change_end)
            
            # Find which cell contains this change
            change_cell = None
            for run in affected_runs:
                if (run.get('is_json_boundary') or 
                    run.get('is_json_escape') or 
                    run.get('is_para_boundary')):
                    continue
                
                # Check if this run is affected by the change
                is_affected = False
                if is_insert:
                    if run['start'] <= change_start <= run['end']:
                        is_affected = True
                else:
                    if run['end'] > change_start and run['start'] < change_end:
                        is_affected = True
                
                if is_affected:
                    cell = run.get('cell_elem')
                    if cell is not None:
                        if change_cell is None:
                            change_cell = cell
                        elif id(change_cell) != id(cell):
                            # Change spans multiple cells
                            return None
            
            if change_cell is not None:
                cell_id = id(change_cell)
                if cell_id not in cell_to_changes:
                    cell_to_changes[cell_id] = []
                cell_to_changes[cell_id].append(change_idx)
            else:
                # Change landed on JSON boundary marker (e.g., ", " or "], [")
                # Cannot be correctly mapped to a cell - this indicates a cross-cell
                # operation like merge/split that requires fallback to comment
                return None
        
        if not cell_to_changes:
            return None
        
        # Group affected runs by cell
        cell_runs_map = {}
        for run in affected_runs:
            if (run.get('is_json_boundary') or 
                run.get('is_json_escape') or 
                run.get('is_para_boundary')):
                continue
            
            cell = run.get('cell_elem')
            if cell is not None:
                cell_id = id(cell)
                if cell_id not in cell_runs_map:
                    cell_runs_map[cell_id] = {'cell': cell, 'runs': []}
                cell_runs_map[cell_id]['runs'].append(run)
        
        # Extract cell-specific edits
        result_edits = []
        
        for cell_id, change_indices in cell_to_changes.items():
            if cell_id not in cell_runs_map:
                continue
            
            cell_info = cell_runs_map[cell_id]
            cell_runs = cell_info['runs']
            
            # Find cell boundaries in violation_text
            cell_start = min(r['start'] for r in cell_runs)
            cell_end = max(r['end'] for r in cell_runs)
            
            relative_start = cell_start - match_start
            relative_end = cell_end - match_start
            
            # Extract cell_violation
            if self._is_table_mode(affected_runs):
                cell_violation_escaped = violation_text[relative_start:relative_end]
                cell_violation = self._decode_json_escaped(cell_violation_escaped)
            else:
                cell_violation = violation_text[relative_start:relative_end]

            # Build cell_revised by applying diff operations within cell range
            violation_pos = 0
            revised_accumulator = ''
            has_changes_in_cell = False  # Track whether any changes affect this cell
            
            for op, text in diff_ops:
                if op == 'equal':
                    chunk_start = violation_pos
                    chunk_end = violation_pos + len(text)
                    
                    # Check if this chunk overlaps with cell range
                    if chunk_end > relative_start and chunk_start < relative_end:
                        overlap_start = max(0, relative_start - chunk_start)
                        overlap_end = min(len(text), relative_end - chunk_start)
                        revised_accumulator += text[overlap_start:overlap_end]
                    
                    violation_pos += len(text)
                
                elif op == 'delete':
                    chunk_start = violation_pos
                    chunk_end = violation_pos + len(text)
                    
                    # Check if this deletion affects the current cell
                    if chunk_end > relative_start and chunk_start < relative_end:
                        has_changes_in_cell = True
                    
                    violation_pos += len(text)
                
                elif op == 'insert':
                    # Check if insertion point is within cell range (exclude right boundary)
                    # Insertions at exact cell boundary (violation_pos == relative_end) are 
                    # ambiguous and should trigger cross-cell fallback
                    if relative_start <= violation_pos < relative_end:
                        revised_accumulator += text
                        has_changes_in_cell = True
            
            # Use accumulator if there were changes (even if empty), otherwise keep original
            cell_revised = revised_accumulator if has_changes_in_cell else cell_violation
            
            # Fix: Decode cell_revised in table mode (same as cell_violation)
            # cell_revised is accumulated from escaped diff chunks and needs decoding
            # to match cell_violation which is already decoded
            if self._is_table_mode(affected_runs) and cell_revised != cell_violation:
                _cell_revised = cell_revised
                cell_revised = self._decode_json_escaped(cell_revised)
                if _cell_revised != cell_revised:
                    print(f"  [Multi-cell] Decoded cell_revised to avoid JSON artifacts: {format_text_preview(_cell_revised, 60)}")

            result_edits.append({
                'cell_violation': cell_violation,
                'cell_revised': cell_revised,
                'cell_elem': cell_info['cell'],
                'cell_runs': cell_runs
            })
        
        return result_edits if result_edits else None

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
                if relative_start <= violation_pos <= relative_end:
                    revised_accumulator += text
        
        cell_revised = revised_accumulator if revised_accumulator else cell_violation

        # Fix: Decode cell_revised in table mode (same as cell_violation)
        # cell_revised is accumulated from escaped diff chunks and needs decoding
        # to avoid injecting JSON boundary characters into the cell text.
        if self._is_table_mode(affected_runs) and cell_revised != cell_violation:
            _cell_revised = cell_revised
            cell_revised = self._decode_json_escaped(cell_revised)
            if _cell_revised != cell_revised:
                print("  [Single-cell] Decoded cell_revised to avoid JSON artifacts: {format_text_preview(_cell_revised, 60)}")

        return {
            'cell_violation': cell_violation,
            'cell_revised': cell_revised,
            'cell_elem': target_cell,
            'cell_runs': cell_runs
        }

    def _extract_text_from_run_element(self, run_elem, rPr) -> str:
        """
        Extract text from a single run element with superscript/subscript markup.
        
        This matches parse_document.py's extract_text_from_run() behavior to ensure
        violation_text from LLM matches the text extracted during apply phase.
        
        Superscript/subscript detection:
        - Check w:rPr/w:vertAlign[@w:val="superscript|subscript"]
        - Wrap text with <sup>...</sup> or <sub>...</sub> markup
        
        Args:
            run_elem: The w:r run element to extract text from
            rPr: Pre-fetched w:rPr element (may be None)
        
        Returns:
            Extracted text with superscript/subscript markup if applicable
        """
        # Check for vertical alignment (superscript/subscript)
        vert_align = None
        if rPr is not None:
            vert_align_elem = rPr.find('w:vertAlign', NS)
            if vert_align_elem is not None:
                vert_align = vert_align_elem.get(f'{{{NS["w"]}}}val')
        
        # Extract text content from run
        text_parts = []
        for elem in run_elem:
            if elem.tag == f'{{{NS["w"]}}}t':
                text = elem.text or ''
                if text:
                    text_parts.append(text)
            elif elem.tag == f'{{{NS["w"]}}}delText':
                # For deleted text in revision markup
                text = elem.text or ''
                if text:
                    text_parts.append(text)
            elif elem.tag == f'{{{NS["w"]}}}tab':
                text_parts.append('\t')
            elif elem.tag == f'{{{NS["w"]}}}br':
                # Handle line breaks - textWrapping or no type = soft line break
                br_type = elem.get(f'{{{NS["w"]}}}type')
                if br_type in (None, 'textWrapping'):
                    text_parts.append('\n')
                # Skip page and column breaks (layout elements)
        
        combined_text = ''.join(text_parts)
        
        # Apply superscript/subscript markup
        if combined_text and vert_align in ('superscript', 'subscript'):
            if vert_align == 'superscript':
                return f'<sup>{combined_text}</sup>'
            else:  # subscript
                return f'<sub>{combined_text}</sub>'
        
        return combined_text

    def _collect_runs_info_original(self, para_elem) -> Tuple[List[Dict], str]:
        """
        Collect run info representing ORIGINAL text (before track changes).
        This is used for searching violation text in documents that have been edited.
        
        Logic:
        - Include: <w:delText> in <w:del> elements (deleted text was part of original)
        - Include: Normal <w:t>, <w:tab>, <w:br> NOT inside <w:ins> or <w:del> elements
        - Exclude: <w:t> inside <w:ins> elements (inserted text didn't exist in original)
        - Superscript/subscript: Wrapped with <sup>/<sub> tags for LLM matching
        
        Returns:
            Tuple of (runs_info, combined_text)
            runs_info: [{'text': str, 'start': int, 'end': int}, ...]
            combined_text: Full text string with <sup>/<sub> markup
        """
        runs_info = []
        pos = 0
        w_ns = NS['w']
        wp_ns = NS['wp']
        m_ns = NS['m']
        w_ins_tag = f'{{{w_ns}}}ins'
        w_r_tag = f'{{{w_ns}}}r'
        m_omath_tag = f'{{{m_ns}}}oMath'
        m_omathpara_tag = f'{{{m_ns}}}oMathPara'

        def append_equation(omath_elem, host_run=None, host_rPr=None) -> None:
            """Append a synthetic run entry for an OMML equation."""
            nonlocal pos
            try:
                from omml import convert_omml_to_latex
                latex = convert_omml_to_latex(omath_elem)
            except Exception:
                return
            if not latex:
                return
            eq_text = f'<equation>{latex}</equation>'
            runs_info.append({
                'text': eq_text,
                'start': pos,
                'end': pos + len(eq_text),
                'elem': host_run,
                'rPr': host_rPr,
                'is_equation': True
            })
            pos += len(eq_text)

        def append_run(run_elem) -> None:
            """Append run entries in order, splitting drawings and equations."""
            nonlocal pos
            rPr = run_elem.find('w:rPr', NS)

            # Check for vertical alignment (superscript/subscript)
            vert_align = None
            if rPr is not None:
                vert_align_elem = rPr.find('w:vertAlign', NS)
                if vert_align_elem is not None:
                    vert_align = vert_align_elem.get(f'{{{w_ns}}}val')

            # Buffer for accumulating text before a drawing/equation
            text_buffer = []

            def flush_text_buffer() -> None:
                """Helper to flush accumulated text as a separate entry."""
                nonlocal pos
                if not text_buffer:
                    return

                combined_text = ''.join(text_buffer)

                # Apply superscript/subscript markup if needed
                if vert_align in ('superscript', 'subscript'):
                    if vert_align == 'superscript':
                        combined_text = f'<sup>{combined_text}</sup>'
                    else:
                        combined_text = f'<sub>{combined_text}</sub>'

                runs_info.append({
                    'text': combined_text,
                    'start': pos,
                    'end': pos + len(combined_text),
                    'elem': run_elem,
                    'rPr': rPr,
                    'is_drawing': False
                })
                pos += len(combined_text)
                text_buffer.clear()

            # Process children in order to preserve text/image/equation positions
            for elem in run_elem:
                if elem.tag == f'{{{w_ns}}}t':
                    text = elem.text or ''
                    if text:
                        text_buffer.append(text)
                elif elem.tag == f'{{{w_ns}}}delText':
                    text = elem.text or ''
                    if text:
                        text_buffer.append(text)
                elif elem.tag == f'{{{w_ns}}}tab':
                    text_buffer.append('\t')
                elif elem.tag == f'{{{w_ns}}}br':
                    br_type = elem.get(f'{{{w_ns}}}type')
                    if br_type in (None, 'textWrapping'):
                        text_buffer.append('\n')
                elif elem.tag == f'{{{w_ns}}}drawing':
                    # Flush any accumulated text before the drawing
                    flush_text_buffer()

                    # Create separate entry for drawing
                    inline = elem.find(f'{{{wp_ns}}}inline')
                    if inline is not None:
                        doc_pr = inline.find(f'{{{wp_ns}}}docPr')
                        if doc_pr is not None:
                            img_id = doc_pr.get('id', '')
                            img_name = doc_pr.get('name', '')
                            drawing_text = f'<drawing id="{img_id}" name="{img_name}" />'

                            runs_info.append({
                                'text': drawing_text,
                                'start': pos,
                                'end': pos + len(drawing_text),
                                'elem': run_elem,
                                'rPr': rPr,
                                'is_drawing': True,
                                'drawing_elem': elem  # Store reference to drawing element
                            })
                            pos += len(drawing_text)
                elif elem.tag == m_omath_tag:
                    # Flush any accumulated text before the equation
                    flush_text_buffer()
                    append_equation(elem, run_elem, rPr)
                elif elem.tag == m_omathpara_tag:
                    # Flush any accumulated text before the equation block
                    flush_text_buffer()
                    for omath in elem:
                        if omath.tag == m_omath_tag:
                            append_equation(omath, run_elem, rPr)

            # Flush any remaining text after all elements
            flush_text_buffer()

        def walk(node) -> None:
            tag = node.tag
            if tag == w_ins_tag:
                # Inserted text = NOT part of original, skip completely
                return
            if tag == w_r_tag:
                append_run(node)
                return
            if tag == m_omath_tag:
                append_equation(node)
                return
            if tag == m_omathpara_tag:
                for child in node:
                    if child.tag == m_omath_tag:
                        append_equation(child)
                    else:
                        walk(child)
                return
            for child in node:
                walk(child)

        # Walk paragraph content in document order
        for child in para_elem:
            walk(child)
        
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

    def _filter_real_runs(self, runs: List[Dict], include_equations: bool = False) -> List[Dict]:
        """
        Filter out synthetic boundary runs (JSON boundaries, paragraph boundaries).

        Synthetic runs are injected for text matching but don't have actual
        document elements (elem/rPr). This method filters them out before
        applying document modifications.

        Args:
            runs: List of run info dicts
            include_equations: If True, keep synthetic equation runs for comment anchoring

        Returns:
            List of runs that have actual document elements
        """
        return [r for r in runs
                if not r.get('is_json_boundary', False)
                and not r.get('is_json_escape', False)
                and not r.get('is_para_boundary', False)
                and (include_equations or not r.get('is_equation', False))]

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

    def _prepare_deletion_items(self, para_groups: List[Dict],
                                match_start: int, match_end: int) -> List[Dict]:
        """
        Build prepared deletion items from paragraph-grouped runs.

        Shared preparation logic for _apply_delete_cross_paragraph and
        _apply_delete_multi_cell. Computes before_text, del_text, after_text
        for each paragraph in the deletion range.

        Args:
            para_groups: List of {'para_elem': ..., 'runs': [...]} dicts
                         (from _group_runs_by_paragraph)
            match_start: Start position of the matched violation text
            match_end: End position of the matched violation text

        Returns:
            List of dicts with keys:
            - para_elem, affected_runs, rPr_xml, before_text, del_text, after_text
        """
        prepared = []
        for group in para_groups:
            para_elem = group['para_elem']
            para_runs = group['runs']

            # Identify affected runs within this paragraph
            first_run = None
            last_run = None
            affected_runs_in_para = []
            for run in para_runs:
                if run['end'] > match_start and run['start'] < match_end:
                    if first_run is None:
                        first_run = run
                    last_run = run
                    affected_runs_in_para.append(run)

            if first_run is None or last_run is None:
                if self.verbose:
                    print(f"  [Prepare] Skipping paragraph: no affected runs")
                continue

            rPr_xml = self._get_rPr_xml(first_run.get('rPr'))

            before_offset = self._translate_escaped_offset(
                first_run, max(0, match_start - first_run['start']))
            after_offset = self._translate_escaped_offset(
                last_run, max(0, match_end - last_run['start']))

            first_orig_text = self._get_run_original_text(first_run)
            last_orig_text = self._get_run_original_text(last_run)

            before_text = first_orig_text[:before_offset]
            after_text = last_orig_text[after_offset:]

            # Build deleted text from affected runs in this paragraph
            del_parts = []
            for run in affected_runs_in_para:
                if run.get('is_drawing'):
                    del_parts.append(run['text'])
                    continue
                orig_text = self._get_run_original_text(run)
                if run is first_run and run is last_run:
                    del_parts.append(orig_text[before_offset:after_offset])
                elif run is first_run:
                    del_parts.append(orig_text[before_offset:])
                elif run is last_run:
                    del_parts.append(orig_text[:after_offset])
                else:
                    del_parts.append(orig_text)

            del_text = ''.join(del_parts)
            if not del_text:
                if self.verbose:
                    print(f"  [Prepare] Skipping paragraph: no text to delete")
                continue

            prepared.append({
                'para_elem': para_elem,
                'affected_runs': affected_runs_in_para,
                'rPr_xml': rPr_xml,
                'before_text': before_text,
                'del_text': del_text,
                'after_text': after_text,
            })

        return prepared

    def _delete_paragraphs_in_unit(self, prepared: List[Dict],
                                    shared_change_id: str, author: str,
                                    comment_id: Optional[int] = None) -> int:
        """
        Delete content across multiple paragraphs in a single unit.

        Shared core loop for _apply_delete_cross_paragraph and _apply_delete_multi_cell.
        Handles building deletion elements, replacing runs, and adding paragraph mark
        deletion for non-last paragraphs so Word displays a unified deletion.

        Args:
            prepared: List from _prepare_deletion_items()
            shared_change_id: Change ID shared across all paragraphs in this unit
            author: Track change author
            comment_id: If set, insert commentRangeStart before first deletion
                        and commentRangeEnd+ref after last deletion

        Returns:
            Number of paragraphs successfully processed
        """
        if not prepared:
            return 0

        success_count = 0
        is_first_para = True

        for item_idx, item in enumerate(prepared):
            para_elem = item['para_elem']
            affected_runs = item['affected_runs']
            rPr_xml = item['rPr_xml']
            before_text = item['before_text']
            del_text = item['del_text']
            after_text = item['after_text']

            is_last_para = (item_idx == len(prepared) - 1)

            # Build new elements for this paragraph
            new_elements = []

            if before_text:
                run_or_container = self._create_run(before_text, rPr_xml)
                if run_or_container.tag == 'container':
                    new_elements.extend(list(run_or_container))
                else:
                    new_elements.append(run_or_container)

            # Comment range start (only for first paragraph, if comment_id provided)
            if is_first_para and comment_id is not None:
                comment_start_xml = (
                    f'<w:commentRangeStart xmlns:w="{NS["w"]}" '
                    f'w:id="{comment_id}"/>'
                )
                new_elements.append(etree.fromstring(comment_start_xml))

            # Deleted text with shared change_id
            self._append_del_elements(
                new_elements, del_text, rPr_xml, author,
                change_id=shared_change_id
            )

            # Comment range end and reference (only for last paragraph)
            if is_last_para and comment_id is not None:
                comment_end_xml = (
                    f'<w:commentRangeEnd xmlns:w="{NS["w"]}" '
                    f'w:id="{comment_id}"/>'
                )
                new_elements.append(etree.fromstring(comment_end_xml))
                comment_ref_xml = (
                    f'<w:r xmlns:w="{NS["w"]}">'
                    f'<w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
                    f'<w:commentReference w:id="{comment_id}"/>'
                    f'</w:r>'
                )
                new_elements.append(etree.fromstring(comment_ref_xml))

            if after_text:
                run_or_container = self._create_run(after_text, rPr_xml)
                if run_or_container.tag == 'container':
                    new_elements.extend(list(run_or_container))
                else:
                    new_elements.append(run_or_container)

            # Replace runs and check success
            replace_success = self._replace_runs(
                para_elem, affected_runs, new_elements
            )

            if replace_success:
                # For non-last paragraphs, mark paragraph mark (¶) as deleted
                # so Word merges consecutive deleted paragraphs into one unified
                # deletion. Without this, Word displays each deletion separately.
                if not is_last_para:
                    pPr = para_elem.find(f'{{{NS["w"]}}}pPr')
                    if pPr is None:
                        pPr = etree.SubElement(para_elem, f'{{{NS["w"]}}}pPr')
                        # pPr must be the first child of w:p
                        para_elem.insert(0, pPr)
                    rPr_in_pPr = pPr.find(f'{{{NS["w"]}}}rPr')
                    if rPr_in_pPr is None:
                        rPr_in_pPr = etree.SubElement(pPr, f'{{{NS["w"]}}}rPr')
                    del_mark = etree.SubElement(
                        rPr_in_pPr, f'{{{NS["w"]}}}del'
                    )
                    del_mark.set(f'{{{NS["w"]}}}id', shared_change_id)
                    del_mark.set(f'{{{NS["w"]}}}author', author)
                    del_mark.set(f'{{{NS["w"]}}}date', self.operation_timestamp)

                success_count += 1
            else:
                if self.verbose:
                    print(f"  [Delete unit] Failed to replace runs in paragraph")

            is_first_para = False

        return success_count

    def _append_del_elements(self, new_elements: List[etree.Element],
                             del_text: str, rPr_xml: str, author: str,
                             change_id: Optional[str] = None):
        """Append deletion elements for del_text into new_elements.
        
        Args:
            new_elements: List to append deletion elements to
            del_text: Text to delete
            rPr_xml: Run properties XML
            author: Author name
            change_id: Optional pre-generated change ID. If None, generates new ID.
        """
        if not del_text:
            return

        change_id = change_id or self._get_next_change_id()
        run_or_container = self._create_run(del_text, rPr_xml)

        if run_or_container.tag == 'container':
            # Multiple runs (has markup) - wrap each in w:del
            for del_run in run_or_container:
                t_elem = del_run.find(f'{{{NS["w"]}}}t')
                if t_elem is not None:
                    t_elem.tag = f'{{{NS["w"]}}}delText'

                del_elem = etree.Element(f'{{{NS["w"]}}}del')
                del_elem.set(f'{{{NS["w"]}}}id', change_id)
                del_elem.set(f'{{{NS["w"]}}}author', author)
                del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                del_elem.append(del_run)
                new_elements.append(del_elem)
        else:
            t_elem = run_or_container.find(f'{{{NS["w"]}}}t')
            if t_elem is not None:
                t_elem.tag = f'{{{NS["w"]}}}delText'

            del_elem = etree.Element(f'{{{NS["w"]}}}del')
            del_elem.set(f'{{{NS["w"]}}}id', change_id)
            del_elem.set(f'{{{NS["w"]}}}author', author)
            del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
            del_elem.append(run_or_container)
            new_elements.append(del_elem)

    def _build_ins_elements_with_breaks(self, text: str, rPr_xml: str, author: str) -> List[etree.Element]:
        """
        Build insertion elements for text, converting '\\n' to <w:br/>.
        Returns a list of w:ins elements.
        """
        if not text:
            return []

        normalized = text.replace('\r\n', '\n').replace('\r', '\n')
        lines = normalized.split('\n')
        ins_elements = []
        change_id = self._get_next_change_id()

        for idx, line in enumerate(lines):
            if line:
                run_or_container = self._create_run(line, rPr_xml)
                if run_or_container.tag == 'container':
                    for ins_run in run_or_container:
                        ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                        ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                        ins_elem.set(f'{{{NS["w"]}}}author', author)
                        ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                        ins_elem.append(ins_run)
                        ins_elements.append(ins_elem)
                else:
                    ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                    ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                    ins_elem.set(f'{{{NS["w"]}}}author', author)
                    ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                    ins_elem.append(run_or_container)
                    ins_elements.append(ins_elem)

            # Insert line break between lines
            if idx < len(lines) - 1:
                br_run = etree.Element(f'{{{NS["w"]}}}r')
                etree.SubElement(br_run, f'{{{NS["w"]}}}br')
                ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                ins_elem.set(f'{{{NS["w"]}}}author', author)
                ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                ins_elem.append(br_run)
                ins_elements.append(ins_elem)

        return ins_elements
    
    # ==================== Diff Calculation ====================
    
    def _check_special_element_modification(self, violation_text: str, diff_ops: List, has_markup: bool = False) -> Tuple[bool, str]:
        """
        Check if any diff operation modifies content inside special elements (<drawing> or <equation>).
        
        Strategy:
        1. Find all special element position ranges in violation_text
        2. If has_markup=True, map ranges to plain-text coordinate space (diff ops work on plain text)
        3. Track position through diff ops
        4. If any delete/insert operation overlaps with special element ranges, reject
        
        Args:
            violation_text: Original violation text (may contain special elements and markup)
            diff_ops: List of diff operations from _calculate_diff or _calculate_markup_aware_diff
            has_markup: If True, diff_ops work on plain text (markup stripped), need coordinate mapping
        
        Returns:
            Tuple of (should_reject, reason)
            - should_reject: True if modification involves special element content
            - reason: Description of why rejection is needed
        """
        # Find all special element position ranges in violation_text (original coordinates)
        special_ranges_orig = []  # [(start, end, element_type), ...]
        
        for match in DRAWING_PATTERN.finditer(violation_text):
            special_ranges_orig.append((match.start(), match.end(), 'drawing'))
        
        for match in EQUATION_PATTERN.finditer(violation_text):
            special_ranges_orig.append((match.start(), match.end(), 'equation'))
        
        # If no special elements, only check if inserting new ones
        if not special_ranges_orig:
            # Rebuild complete inserted text from consecutive insert operations
            # This handles the case where markup-aware diff splits insertions containing
            # <equation> tags into multiple chunks (e.g., '<equation>H', '2', 'O</equation>')
            # where individual chunks don't match the full pattern
            full_insert_text = ''.join(
                op_tuple[1] for op_tuple in diff_ops 
                if op_tuple[0] == 'insert'
            )
            
            if full_insert_text:
                if DRAWING_PATTERN.search(full_insert_text):
                    return True, "Cannot insert drawing via revision markup"
                if EQUATION_PATTERN.search(full_insert_text):
                    return True, "Cannot insert equation via revision markup"
            return False, ""
        
        # Check if diff operations would modify content inside special elements
        # Strategy: Track position through diff ops and check overlap with special element ranges
        
        # Map special_ranges to plain-text coordinates if markup is present
        if has_markup and special_ranges_orig:
            # Build mapping from original position to plain-text position
            segments = self._parse_formatted_text(violation_text)
            orig_to_plain = {}  # Map original char index -> plain char index
            plain_pos = 0
            orig_pos = 0
            
            for text, vert_align in segments:
                # Original text may have <sup>text</sup> (11 chars for "text")
                # Plain text has just "text" (4 chars)
                if vert_align == 'superscript':
                    # Original: <sup>text</sup>
                    markup_before = '<sup>'
                    markup_after = '</sup>'
                elif vert_align == 'subscript':
                    # Original: <sub>text</sub>
                    markup_before = '<sub>'
                    markup_after = '</sub>'
                else:
                    # No markup
                    markup_before = ''
                    markup_after = ''
                
                # Skip markup_before in original, map content
                orig_pos += len(markup_before)
                for char in text:
                    orig_to_plain[orig_pos] = plain_pos
                    orig_pos += 1
                    plain_pos += 1
                orig_pos += len(markup_after)
            
            # Transform special_ranges to plain-text coordinates
            special_ranges = []
            for elem_start_orig, elem_end_orig, elem_type in special_ranges_orig:
                # Find plain-text positions for element boundaries
                # Use the first and last content positions within the element
                elem_start_plain = None
                elem_end_plain = None
                
                # Find first content position in element
                for orig_idx in range(elem_start_orig, elem_end_orig):
                    if orig_idx in orig_to_plain:
                        elem_start_plain = orig_to_plain[orig_idx]
                        break
                
                # Find last content position in element
                for orig_idx in range(elem_end_orig - 1, elem_start_orig - 1, -1):
                    if orig_idx in orig_to_plain:
                        elem_end_plain = orig_to_plain[orig_idx] + 1  # +1 for exclusive end
                        break
                
                # If element is entirely within markup tags (no content mapped), skip it
                # This shouldn't happen for <drawing> or <equation> but handle gracefully
                if elem_start_plain is not None and elem_end_plain is not None:
                    special_ranges.append((elem_start_plain, elem_end_plain, elem_type))
        else:
            # No markup: use original coordinates directly
            special_ranges = special_ranges_orig
        
        # Pre-check: Rebuild full insert text and check for new special elements
        # This must be done BEFORE the main loop to catch markup-split insertions
        # (e.g., '<equation>H', '2', 'O</equation>' from markup-aware diff)
        full_insert_text = ''.join(
            op_tuple[1] for op_tuple in diff_ops 
            if op_tuple[0] == 'insert'
        )
        
        if full_insert_text:
            if DRAWING_PATTERN.search(full_insert_text):
                return True, "Cannot insert drawing via revision markup"
            if EQUATION_PATTERN.search(full_insert_text):
                return True, "Cannot insert equation via revision markup"
        
        # Track cumulative deletion coverage for each special element
        # This handles cases where markup-aware diff splits a complete deletion into multiple delete ops
        # (e.g., deleting <equation>H<sub>2</sub>O</equation> produces separate deletes for non-markup and subscript segments)
        elem_deleted_ranges = {i: [] for i in range(len(special_ranges))}
        
        # First pass: collect all delete operations and their coverage of each special element
        # Also check if any special element would survive in "equal" segments when has_markup=True
        current_pos = 0
        for op_tuple in diff_ops:
            # Handle both 2-tuple and 3-tuple formats
            op = op_tuple[0]
            text = op_tuple[1]

            if op == 'equal' and has_markup:
                # When has_markup=True, equal segments preserve content unchanged
                # No need to check special elements - they remain intact
                current_pos += len(text)

            elif op == 'delete':
                del_start = current_pos
                del_end = current_pos + len(text)

                # Record overlap with each special element
                for idx, (elem_start, elem_end, elem_type) in enumerate(special_ranges):
                    if del_end > elem_start and del_start < elem_end:
                        # Calculate overlap relative to element start
                        overlap_start = max(del_start, elem_start) - elem_start
                        overlap_end = min(del_end, elem_end) - elem_start
                        elem_deleted_ranges[idx].append((overlap_start, overlap_end))

                current_pos += len(text)

            elif op == 'equal':
                current_pos += len(text)

            elif op == 'insert':
                # Check if inserting at a position inside special element
                for elem_start, elem_end, elem_type in special_ranges:
                    if elem_start < current_pos < elem_end:
                        return True, f"Cannot insert inside {elem_type}"
                
                # Note: Full insert text is already checked before loop (see pre-check above)
                # Individual chunk checks are redundant and buggy - removed

        # Second pass: check if any special element was partially deleted
        for idx, (elem_start, elem_end, elem_type) in enumerate(special_ranges):
            deleted_ranges = elem_deleted_ranges[idx]
            
            if not deleted_ranges:
                continue  # Element not affected by deletions
            
            elem_len = elem_end - elem_start
            
            # Merge overlapping ranges and calculate total coverage
            merged_ranges = []
            for start, end in sorted(deleted_ranges):
                if merged_ranges and start <= merged_ranges[-1][1]:
                    # Overlapping or adjacent - merge
                    merged_ranges[-1] = (merged_ranges[-1][0], max(merged_ranges[-1][1], end))
                else:
                    merged_ranges.append((start, end))
            
            total_deleted = sum(end - start for start, end in merged_ranges)
            
            if total_deleted == elem_len:
                # Complete deletion - reject (cannot track change delete special elements)
                return True, f"Cannot delete {elem_type} via revision markup"
            else:
                # Partial deletion - reject
                return True, f"Cannot partially modify {elem_type} content"

        return False, ""
    
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
    
    def _calculate_markup_aware_diff(self, old_text: str, new_text: str) -> List[Tuple[str, str, Optional[str]]]:
        """
        Calculate diff with markup awareness for <sup>/<sub> tags.
        
        This method parses both texts into segments with formatting info,
        then performs diff on the text content only (ignoring markup tags).
        
        Args:
            old_text: Original text with possible <sup>/<sub> markup
            new_text: New text with possible <sup>/<sub> markup
        
        Returns:
            List of (operation, text, vert_align) tuples where:
            - operation: 'equal' | 'delete' | 'insert'
            - text: The actual text content (without markup tags)
            - vert_align: 'superscript' | 'subscript' | None
        
        Examples:
            old: "x<sup>2</sup>"  new: "x<sup>3</sup>"
            → [('equal', 'x', None), ('delete', '2', 'superscript'), ('insert', '3', 'superscript')]
        """
        # Parse both texts into segments
        old_segments = self._parse_formatted_text(old_text)
        new_segments = self._parse_formatted_text(new_text)
        
        # Build plain text versions for diffing (text content only, no markup)
        old_plain = ''.join(text for text, _ in old_segments)
        new_plain = ''.join(text for text, _ in new_segments)
        
        # Perform character-level diff on plain text
        matcher = difflib.SequenceMatcher(None, old_plain, new_plain)
        operations = []
        
        for tag, i1, i2, j1, j2 in matcher.get_opcodes():
            if tag == 'equal':
                # Find segments in old_text that cover this range
                ops = self._map_range_to_segments(old_segments, i1, i2, 'equal')
                operations.extend(ops)
            elif tag == 'delete':
                # Find segments in old_text
                ops = self._map_range_to_segments(old_segments, i1, i2, 'delete')
                operations.extend(ops)
            elif tag == 'insert':
                # Find segments in new_text
                ops = self._map_range_to_segments(new_segments, j1, j2, 'insert')
                operations.extend(ops)
            elif tag == 'replace':
                # Delete from old, insert from new
                ops_del = self._map_range_to_segments(old_segments, i1, i2, 'delete')
                ops_ins = self._map_range_to_segments(new_segments, j1, j2, 'insert')
                operations.extend(ops_del)
                operations.extend(ops_ins)
        
        return operations
    
    def _map_range_to_segments(self, segments: List[Tuple[str, Optional[str]]], 
                                start: int, end: int, operation: str) -> List[Tuple[str, str, Optional[str]]]:
        """
        Map a character range to segments with formatting info.
        
        Args:
            segments: List of (text, vert_align) tuples
            start: Start position in plain text
            end: End position in plain text
            operation: 'equal' | 'delete' | 'insert'
        
        Returns:
            List of (operation, text, vert_align) tuples
        """
        result = []
        pos = 0
        
        for text, vert_align in segments:
            segment_start = pos
            segment_end = pos + len(text)
            
            # Check if this segment overlaps with [start, end)
            if segment_end <= start:
                # Before the range
                pos = segment_end
                continue
            if segment_start >= end:
                # After the range
                break
            
            # Calculate overlap
            overlap_start = max(0, start - segment_start)
            overlap_end = min(len(text), end - segment_start)
            
            if overlap_start < overlap_end:
                chunk = text[overlap_start:overlap_end]
                result.append((operation, chunk, vert_align))
            
            pos = segment_end
        
        return result
    
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
    
    def _parse_formatted_text(self, text: str) -> List[Tuple[str, Optional[str]]]:
        """
        Parse text with <sup>/<sub> markup into segments with format info.
        
        This function splits text containing superscript/subscript markup into
        segments, each with its associated vertical alignment type.
        
        Args:
            text: Text possibly containing <sup>...</sup> or <sub>...</sub> markup
        
        Returns:
            List of (text_content, vert_align) tuples where:
            - text_content: The actual text without markup tags
            - vert_align: 'superscript' | 'subscript' | None
        
        Examples:
            "x<sup>2</sup>" -> [("x", None), ("2", "superscript")]
            "H<sub>2</sub>O" -> [("H", None), ("2", "subscript"), ("O", None)]
            "normal text" -> [("normal text", None)]
        """
        if '<sup>' not in text and '<sub>' not in text:
            # Fast path: no markup
            return [(text, None)]
        
        segments = []
        pos = 0
        
        # Pattern to match <sup>...</sup> or <sub>...</sub>
        # Non-greedy match to handle multiple tags correctly
        pattern = re.compile(r'<(sup|sub)>(.*?)</\1>', re.DOTALL)
        
        for match in pattern.finditer(text):
            # Add text before this tag (if any)
            if match.start() > pos:
                segments.append((text[pos:match.start()], None))
            
            # Add the tagged content
            tag_type = match.group(1)  # 'sup' or 'sub'
            tag_content = match.group(2)
            vert_align = 'superscript' if tag_type == 'sup' else 'subscript'
            segments.append((tag_content, vert_align))
            
            pos = match.end()
        
        # Add remaining text after last tag (if any)
        if pos < len(text):
            segments.append((text[pos:], None))
        
        return segments
    
    def _create_run(self, text: str, rPr_xml: str = '') -> etree.Element:
        """
        Create new w:r element(s) with text, supporting <sup>/<sub> markup.
        
        If text contains superscript/subscript markup, returns a document fragment
        with multiple runs (each with appropriate w:vertAlign in w:rPr).
        
        Args:
            text: Text to insert, may contain <sup>...</sup> or <sub>...</sub>
            rPr_xml: Base run properties XML (will be modified to add w:vertAlign)
        
        Returns:
            Single w:r element or document fragment containing multiple w:r elements
        """
        segments = self._parse_formatted_text(text)
        
        if len(segments) == 1 and segments[0][1] is None:
            # Simple case: no formatting, return single run
            run_xml = f'<w:r xmlns:w="{NS["w"]}">{rPr_xml}<w:t>{self._escape_xml(text)}</w:t></w:r>'
            return etree.fromstring(run_xml)
        
        # Complex case: need multiple runs with different formatting
        # Parse base rPr to potentially modify it
        if rPr_xml:
            # Check if rPr_xml is already a complete <w:rPr> element
            # _get_rPr_xml() returns the full element, so parse it directly
            if rPr_xml.strip().startswith('<w:rPr') or rPr_xml.strip().startswith('<rPr'):
                # Already a complete element, parse directly
                base_rPr = etree.fromstring(rPr_xml)
            else:
                # Just inner content, wrap in <w:rPr>
                base_rPr = etree.fromstring(f'<w:rPr xmlns:w="{NS["w"]}">{rPr_xml}</w:rPr>')
        else:
            base_rPr = None
        
        # Create a container for multiple runs
        container = etree.Element('container')
        
        for segment_text, vert_align in segments:
            if not segment_text:
                continue  # Skip empty segments
            
            # Start with a copy of base rPr
            if base_rPr is not None:
                run_rPr = etree.fromstring(etree.tostring(base_rPr, encoding='unicode'))
            else:
                run_rPr = etree.Element(f'{{{NS["w"]}}}rPr')
            
            # Add or update w:vertAlign if needed
            if vert_align:
                # Remove existing w:vertAlign if present
                for existing in run_rPr.findall('w:vertAlign', NS):
                    run_rPr.remove(existing)
                
                # Add new w:vertAlign
                vert_align_elem = etree.SubElement(run_rPr, f'{{{NS["w"]}}}vertAlign')
                vert_align_elem.set(f'{{{NS["w"]}}}val', vert_align)
            else:
                # For normal segments (vert_align=None), strip any inherited vertAlign
                # to prevent normal text from being incorrectly rendered as super/subscript
                for existing in run_rPr.findall('w:vertAlign', NS):
                    run_rPr.remove(existing)
            
            # Create run with modified rPr
            run = etree.Element(f'{{{NS["w"]}}}r')
            run.append(run_rPr)
            
            t_elem = etree.SubElement(run, f'{{{NS["w"]}}}t')
            t_elem.text = segment_text
            
            container.append(run)
        
        # If only one run was created, return it directly
        if len(container) == 1:
            return container[0]
        
        # Otherwise return the container (caller will need to extract children)
        return container
    
    def _replace_runs(self, _para_elem, affected_runs: List[Dict],
                     new_elements: List[etree.Element]) -> bool:
        """
        Replace affected runs with new elements in the paragraph.
        
        Returns:
            True if replacement succeeded, False if DOM operation failed
        """
        if not affected_runs or not new_elements:
            return False
        
        first_run = affected_runs[0]['elem']
        parent = first_run.getparent()
        
        # Safety check: if element is no longer in DOM, skip
        if parent is None:
            if self.verbose:
                print("  [Warning] _replace_runs failed: first_run has no parent")
            return False
        
        try:
            insert_idx = list(parent).index(first_run)
        except ValueError:
            # Element no longer in parent
            if self.verbose:
                print("  [Warning] _replace_runs failed: first_run not in parent")
            return False
        
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
        
        return True
    
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

    def _apply_delete_cross_paragraph(self, orig_runs_info: List[Dict],
                                      orig_match_start: int,
                                      violation_text: str,
                                      violation_reason: str,
                                      author: str) -> str:
        """
        Apply delete across multiple paragraphs (body text only).

        Strategy:
        - Delete ranges per paragraph with track changes (w:del)
        - Preserve paragraph structure (do not remove w:p)
        - Add a comment for each paragraph segment
        """
        match_start = orig_match_start
        match_end = match_start + len(violation_text)

        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)
        if not affected:
            if self.verbose:
                print("  [Cross-paragraph delete] No affected runs found")
            return 'cross_paragraph_fallback'

        real_runs = self._filter_real_runs(affected)
        if not real_runs:
            if self.verbose:
                print("  [Cross-paragraph delete] No real runs after filtering")
            return 'cross_paragraph_fallback'

        # Do not apply cross-paragraph delete in table mode
        if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
            if self.verbose:
                print("  [Cross-paragraph delete] Table mode detected, fallback")
            return 'cross_paragraph_fallback'

        # Check overlap with existing revisions
        if self._check_overlap_with_revisions(real_runs):
            if self.verbose:
                print("  [Cross-paragraph delete] Overlap with existing revisions")
            return 'conflict'

        para_groups = self._group_runs_by_paragraph(real_runs)
        if not para_groups:
            if self.verbose:
                print("  [Cross-paragraph delete] No paragraph groups")
            return 'cross_paragraph_fallback'

        if self.verbose:
            print(f"  [Cross-paragraph delete] Processing {len(para_groups)} paragraph(s)")

        # Use shared helper to build prepared deletion items
        prepared = self._prepare_deletion_items(para_groups, match_start, match_end)

        if not prepared:
            if self.verbose:
                print("  [Cross-paragraph delete] No paragraphs prepared for deletion")
            return 'cross_paragraph_fallback'

        # Pre-generate shared change_id and comment_id for unified deletion
        shared_change_id = self._get_next_change_id()
        comment_id = self.next_comment_id
        self.next_comment_id += 1

        # Use shared helper to apply deletions with paragraph mark merging
        success_count = self._delete_paragraphs_in_unit(
            prepared, shared_change_id, author, comment_id=comment_id
        )

        if success_count == 0:
            if self.verbose:
                print(f"  [Cross-paragraph delete] All paragraphs failed, fallback to comment")
            return 'cross_paragraph_fallback'
        
        # Record single comment for the unified deletion
        self.comments.append({
            'id': comment_id,
            'text': violation_reason,
            'author': f"{author}-R"
        })
        
        if success_count < len(prepared):
            if self.verbose:
                print(f"  [Cross-paragraph delete] Partial success: {success_count}/{len(prepared)} paragraphs")
        else:
            if self.verbose:
                print(f"  [Cross-paragraph delete] All {success_count} paragraphs succeeded with unified deletion")

        return 'success'

    def _apply_replace_cross_paragraph(self, orig_runs_info: List[Dict],
                                       orig_match_start: int,
                                       violation_text: str,
                                       revised_text: str,
                                       violation_reason: str,
                                       author: str) -> str:
        """
        Apply replace across multiple paragraphs with diff-level edits.

        Strategy:
        - Compute diff between violation_text and revised_text
        - Apply per-paragraph replace only on changed segments
        - If diff deletes paragraph boundary ('\\n'), merge paragraphs by removing
          the boundary and appending following paragraph content to the previous one
        """
        match_start = orig_match_start
        match_end = match_start + len(violation_text)

        affected = self._find_affected_runs(orig_runs_info, match_start, match_end)
        if not affected:
            return 'cross_paragraph_fallback'

        real_runs = self._filter_real_runs(affected)
        if not real_runs:
            return 'cross_paragraph_fallback'

        # Do not apply cross-paragraph replace in table mode
        if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
            return 'cross_paragraph_fallback'

        # Check overlap with existing revisions
        if self._check_overlap_with_revisions(real_runs):
            return 'conflict'

        # Compute diff ops
        has_markup = ('<sup>' in violation_text or '<sub>' in violation_text or
                      '<sup>' in revised_text or '<sub>' in revised_text)
        if has_markup:
            diff_ops = self._calculate_markup_aware_diff(violation_text, revised_text)
        else:
            plain_diff = self._calculate_diff(violation_text, revised_text)
            diff_ops = [(op, text, None) for op, text in plain_diff]

        # Check for special element modification (drawing/equation)
        should_reject, reject_reason = self._check_special_element_modification(violation_text, diff_ops, has_markup)
        if should_reject:
            if self.verbose:
                print(f"  [Fallback] {reject_reason}")
            return 'cross_paragraph_fallback'

        # Detect deleted paragraph boundaries
        boundary_positions = [idx for idx, ch in enumerate(violation_text) if ch == '\n']
        boundary_pos_to_idx = {pos: i for i, pos in enumerate(boundary_positions)}
        deleted_boundary_indices = set()
        orig_pos = 0
        for op_tuple in diff_ops:
            op, text, _ = op_tuple if len(op_tuple) == 3 else (*op_tuple, None)
            if op == 'delete' and '\n' in text:
                for i, ch in enumerate(text):
                    if ch == '\n':
                        bpos = orig_pos + i
                        if bpos in boundary_pos_to_idx:
                            deleted_boundary_indices.add(boundary_pos_to_idx[bpos])
            if op in ('equal', 'delete'):
                orig_pos += len(text)

        # Build combined text for the match range
        combined_text = ''.join(r.get('text', '') for r in orig_runs_info)
        match_text = combined_text[match_start:match_end]

        # Build paragraph groups in order
        para_groups = self._group_runs_by_paragraph(real_runs)
        if not para_groups:
            return 'cross_paragraph_fallback'

        # Build paragraph ranges within combined_text
        para_ranges = []
        for group in para_groups:
            runs = group['runs']
            para_start = min(r['start'] for r in runs)
            para_end = max(r['end'] for r in runs)
            overlap_start = max(match_start, para_start)
            overlap_end = min(match_end, para_end)
            if overlap_start < overlap_end:
                para_ranges.append({
                    'para_elem': group['para_elem'],
                    'overlap_start': overlap_start,
                    'overlap_end': overlap_end,
                    'para_start': para_start
                })

        if not para_ranges:
            return 'cross_paragraph_fallback'

        # Helper: extract revised segment for an original range
        def extract_revised_segment(seg_start: int, seg_end: int) -> str:
            orig_pos = 0
            parts = []
            for op_tuple in diff_ops:
                op, text, _ = op_tuple if len(op_tuple) == 3 else (*op_tuple, None)
                if op == 'equal':
                    op_start = orig_pos
                    op_end = orig_pos + len(text)
                    # overlap with [seg_start, seg_end)
                    if op_end > seg_start and op_start < seg_end:
                        take_start = max(seg_start, op_start) - op_start
                        take_end = min(seg_end, op_end) - op_start
                        parts.append(text[take_start:take_end])
                    orig_pos += len(text)
                elif op == 'delete':
                    orig_pos += len(text)
                elif op == 'insert':
                    # insert occurs at current orig_pos
                    if seg_start <= orig_pos <= seg_end:
                        parts.append(text)
            return ''.join(parts)

        # Apply per-paragraph replace
        any_applied = False
        for pr in para_ranges:
            para_elem = pr['para_elem']
            seg_start = pr['overlap_start'] - match_start
            seg_end = pr['overlap_end'] - match_start
            orig_segment = match_text[seg_start:seg_end]
            revised_segment = extract_revised_segment(seg_start, seg_end)

            if orig_segment == revised_segment:
                continue

            para_runs_info, _ = self._collect_runs_info_original(para_elem)
            match_start_in_para = pr['overlap_start'] - pr['para_start']
            status = self._apply_replace(
                para_elem,
                orig_segment,
                revised_segment,
                violation_reason,
                para_runs_info,
                match_start_in_para,
                author
            )

            if status == 'conflict':
                return 'conflict'
            if status not in ('success',):
                return 'cross_paragraph_fallback'
            any_applied = True

        # Merge paragraphs if boundary deleted
        if deleted_boundary_indices:
            # Collect paragraph elements in order
            para_elems = [pr['para_elem'] for pr in para_ranges]
            # Merge from right to left to avoid index shifts
            for b_idx in sorted(deleted_boundary_indices, reverse=True):
                if b_idx < 0 or b_idx + 1 >= len(para_elems):
                    continue
                prev_para = para_elems[b_idx]
                next_para = para_elems[b_idx + 1]
                # Move children (except pPr) from next to prev
                for child in list(next_para):
                    if child.tag == f'{{{NS["w"]}}}pPr':
                        continue
                    next_para.remove(child)
                    prev_para.append(child)
                parent = next_para.getparent()
                if parent is not None:
                    try:
                        parent.remove(next_para)
                    except ValueError:
                        pass
                # Remove from list to keep indices consistent
                para_elems.pop(b_idx + 1)

        return 'success' if any_applied else 'success'
    
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
            'equation_fallback': Equation-only content cannot be edited
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
            if any(r.get('is_equation', False) for r in affected):
                return 'equation_fallback'
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
            run_or_container = self._create_run(before_text, rPr_xml)
            if run_or_container.tag == 'container':
                new_elements.extend(list(run_or_container))
            else:
                new_elements.append(run_or_container)

        # Comment range start (before deleted text)
        comment_start_xml = f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
        new_elements.append(etree.fromstring(comment_start_xml))

        # Decode violation_text if in table mode (JSON-escaped)
        if self._is_table_mode(real_runs):
            del_text = self._decode_json_escaped(violation_text)
        else:
            del_text = violation_text

        # Deleted text - use _create_run to handle <sup>/<sub> markup
        change_id = self._get_next_change_id()
        
        # Create runs with proper vertAlign formatting
        run_or_container = self._create_run(del_text, rPr_xml)
        
        if run_or_container.tag == 'container':
            # Multiple runs (has markup) - wrap each in w:del
            for del_run in run_or_container:
                # Change w:t to w:delText
                t_elem = del_run.find(f'{{{NS["w"]}}}t')
                if t_elem is not None:
                    t_elem.tag = f'{{{NS["w"]}}}delText'
                
                # Wrap in w:del
                del_elem = etree.Element(f'{{{NS["w"]}}}del')
                del_elem.set(f'{{{NS["w"]}}}id', change_id)
                del_elem.set(f'{{{NS["w"]}}}author', author)
                del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                del_elem.append(del_run)
                new_elements.append(del_elem)
        else:
            # Single run - wrap in w:del
            t_elem = run_or_container.find(f'{{{NS["w"]}}}t')
            if t_elem is not None:
                t_elem.tag = f'{{{NS["w"]}}}delText'
            
            del_elem = etree.Element(f'{{{NS["w"]}}}del')
            del_elem.set(f'{{{NS["w"]}}}id', change_id)
            del_elem.set(f'{{{NS["w"]}}}author', author)
            del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
            del_elem.append(run_or_container)
            new_elements.append(del_elem)

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
            run_or_container = self._create_run(after_text, rPr_xml)
            if run_or_container.tag == 'container':
                new_elements.extend(list(run_or_container))
            else:
                new_elements.append(run_or_container)

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
                      author: str,
                      skip_comment: bool = False) -> str:
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
            skip_comment: If True, skip adding comment (used for multi-cell operations)

        Returns:
            'success': Replace applied (may be partial)
            'fallback': Should fallback to comment annotation
            'equation_fallback': Equation-only content cannot be edited
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
            if any(r.get('is_equation', False) for r in affected):
                return 'equation_fallback'
            return 'fallback'

        # Check if we need markup-aware diff (detect <sup>/<sub> tags)
        has_markup = ('<sup>' in violation_text or '<sub>' in violation_text or 
                      '<sup>' in revised_text or '<sub>' in revised_text)
        
        if has_markup:
            # Use markup-aware diff that preserves formatting
            diff_ops = self._calculate_markup_aware_diff(violation_text, revised_text)
        else:
            # Use standard character-level diff for plain text
            # Convert to markup-aware format for consistent handling
            plain_diff = self._calculate_diff(violation_text, revised_text)
            diff_ops = [(op, text, None) for op, text in plain_diff]

        # Check for special element modification (drawing/equation)
        should_reject, reject_reason = self._check_special_element_modification(violation_text, diff_ops, has_markup)
        if should_reject:
            if self.verbose:
                print(f"  [Fallback] {reject_reason}")
            return 'fallback'

        # Check for conflicts only on delete operations
        current_pos = match_start
        for op_tuple in diff_ops:
            # Extract operation and text (ignore vert_align for conflict checking)
            if len(op_tuple) == 3:
                op, text, _ = op_tuple
            else:
                op, text = op_tuple
            
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

        # Allocate comment ID for wrapping the replaced text (unless skipped)
        comment_id = None
        if not skip_comment:
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
            run_or_container = self._create_run(before_text, rPr_xml)
            if run_or_container.tag == 'container':
                new_elements.extend(list(run_or_container))
            else:
                new_elements.append(run_or_container)

        # Comment range start (before replaced text) - only if not skipped
        if not skip_comment:
            comment_start_xml = f'<w:commentRangeStart xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
            new_elements.append(etree.fromstring(comment_start_xml))

        # Check if we're in table mode (need to decode JSON-escaped text)
        is_table_mode = self._is_table_mode(real_runs)

        # Process diff operations
        violation_pos = 0  # Position within violation_text (plain text, no markup)

        for op_tuple in diff_ops:
            # Extract components - handle both 2-tuple and 3-tuple formats
            if len(op_tuple) == 3:
                op, text, vert_align = op_tuple
            else:
                op, text = op_tuple
                vert_align = None
            if op == 'equal':
                # When has_markup=True, position mapping between plain text (violation_pos)
                # and combined_text (orig_runs_info positions) is incorrect due to <sup>/<sub> tags.
                # Skip run preservation and just recreate the text.
                if has_markup:
                    # Decode if in table mode
                    equal_text = self._decode_json_escaped(text) if is_table_mode else text
                    
                    # Wrap with markup if vert_align is specified
                    if vert_align == 'superscript':
                        equal_text = f'<sup>{equal_text}</sup>'
                    elif vert_align == 'subscript':
                        equal_text = f'<sub>{equal_text}</sub>'
                    
                    # Create run with proper vertAlign formatting
                    run_or_container = self._create_run(equal_text, rPr_xml)
                    if run_or_container.tag == 'container':
                        new_elements.extend(list(run_or_container))
                    else:
                        new_elements.append(run_or_container)
                else:
                    # No markup: preserve original elements (especially images)
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
                                    run_or_container = self._create_run(portion, rPr_xml)
                                    if run_or_container.tag == 'container':
                                        new_elements.extend(list(run_or_container))
                                    else:
                                        new_elements.append(run_or_container)
                    else:
                        # No runs found, create text directly
                        # Decode if in table mode and use _create_run to handle markup
                        equal_text = self._decode_json_escaped(text) if is_table_mode else text
                        run_or_container = self._create_run(equal_text, rPr_xml)
                        if run_or_container.tag == 'container':
                            new_elements.extend(list(run_or_container))
                        else:
                            new_elements.append(run_or_container)

                violation_pos += len(text)

            elif op == 'delete':
                # Decode if in table mode
                del_text = self._decode_json_escaped(text) if is_table_mode else text
                
                # Wrap with markup if vert_align is specified
                if vert_align == 'superscript':
                    del_text = f'<sup>{del_text}</sup>'
                elif vert_align == 'subscript':
                    del_text = f'<sub>{del_text}</sub>'
                
                change_id = self._get_next_change_id()
                
                # Create runs with proper vertAlign formatting
                run_or_container = self._create_run(del_text, rPr_xml)
                
                if run_or_container.tag == 'container':
                    # Multiple runs (has markup) - wrap each in w:del
                    for del_run in run_or_container:
                        # Change w:t to w:delText
                        t_elem = del_run.find(f'{{{NS["w"]}}}t')
                        if t_elem is not None:
                            t_elem.tag = f'{{{NS["w"]}}}delText'
                        
                        # Wrap in w:del
                        del_elem = etree.Element(f'{{{NS["w"]}}}del')
                        del_elem.set(f'{{{NS["w"]}}}id', change_id)
                        del_elem.set(f'{{{NS["w"]}}}author', author)
                        del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                        del_elem.append(del_run)
                        new_elements.append(del_elem)
                else:
                    # Single run - wrap in w:del
                    t_elem = run_or_container.find(f'{{{NS["w"]}}}t')
                    if t_elem is not None:
                        t_elem.tag = f'{{{NS["w"]}}}delText'
                    
                    del_elem = etree.Element(f'{{{NS["w"]}}}del')
                    del_elem.set(f'{{{NS["w"]}}}id', change_id)
                    del_elem.set(f'{{{NS["w"]}}}author', author)
                    del_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                    del_elem.append(run_or_container)
                    new_elements.append(del_elem)
                
                violation_pos += len(text)

            elif op == 'insert':
                # Decode if in table mode
                ins_text = self._decode_json_escaped(text) if is_table_mode else text
                
                # Wrap with markup if vert_align is specified
                if vert_align == 'superscript':
                    ins_text = f'<sup>{ins_text}</sup>'
                elif vert_align == 'subscript':
                    ins_text = f'<sub>{ins_text}</sub>'

                # Handle soft line breaks for inserted text
                if '\n' in ins_text:
                    ins_elements = self._build_ins_elements_with_breaks(ins_text, rPr_xml, author)
                    new_elements.extend(ins_elements)
                    # insert doesn't consume violation_pos
                    continue
                
                change_id = self._get_next_change_id()
                
                # Create runs with proper vertAlign formatting
                run_or_container = self._create_run(ins_text, rPr_xml)
                
                if run_or_container.tag == 'container':
                    # Multiple runs (has markup) - wrap each in w:ins
                    for ins_run in run_or_container:
                        # Wrap in w:ins
                        ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                        ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                        ins_elem.set(f'{{{NS["w"]}}}author', author)
                        ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                        ins_elem.append(ins_run)
                        new_elements.append(ins_elem)
                else:
                    # Single run - wrap in w:ins
                    ins_elem = etree.Element(f'{{{NS["w"]}}}ins')
                    ins_elem.set(f'{{{NS["w"]}}}id', change_id)
                    ins_elem.set(f'{{{NS["w"]}}}author', author)
                    ins_elem.set(f'{{{NS["w"]}}}date', self.operation_timestamp)
                    ins_elem.append(run_or_container)
                    new_elements.append(ins_elem)
                # insert doesn't consume violation_pos

        # Comment range end and reference (after replaced text) - only if not skipped
        if not skip_comment:
            comment_end_xml = f'<w:commentRangeEnd xmlns:w="{NS["w"]}" w:id="{comment_id}"/>'
            new_elements.append(etree.fromstring(comment_end_xml))

            comment_ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                <w:commentReference w:id="{comment_id}"/>
            </w:r>'''
            new_elements.append(etree.fromstring(comment_ref_xml))

        # After text (unchanged part after the match)
        if after_text:
            run_or_container = self._create_run(after_text, rPr_xml)
            if run_or_container.tag == 'container':
                new_elements.extend(list(run_or_container))
            else:
                new_elements.append(run_or_container)

        # Single DOM operation to replace all real runs (not boundary markers)
        self._replace_runs(para_elem, real_runs, new_elements)

        # Record comment with violation_reason as content (only if not skipped)
        # Use "-R" suffix to distinguish comment author from track change author
        if not skip_comment:
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
        comment_text = f"[FALLBACK]{reason} {{WHY}}{item.violation_reason}  {{WHERE}}{item.violation_text} {{SUGGEST}}{item.revised_text}"
        
        self.comments.append({
            'id': comment_id,
            'text': comment_text,
            'author': self._author_for_item(item)
        })
        
        return True
    
    def _apply_cell_fallback_comment(self, para_elem, cell_violation: str, 
                                      cell_revised: str, reason: str, 
                                      item: EditItem) -> bool:
        """
        Insert a fallback comment for a single failed cell in multi-cell operation.
        
        Args:
            para_elem: Paragraph element in the failed cell
            cell_violation: Cell-specific violation text
            cell_revised: Cell-specific revised text
            reason: Reason for failure (e.g., "Text not found")
            item: Original edit item (for rule_id and violation_reason)
        
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
        
        # Format: [CELL FAILED] reason | {WHY} ... {WHERE} ... {SUGGEST} ...
        comment_text = f"[CELL FAILED] {reason}\n{{WHY}}{item.violation_reason}  {{WHERE}}{cell_violation}{{SUGGEST}}{cell_revised}"
        
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
                     is_cross_paragraph: bool = False,
                     fallback_reason: Optional[str] = None) -> str:
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
        real_runs = self._filter_real_runs(affected, include_equations=True)

        if not real_runs:
            return 'fallback'

        comment_id = self.next_comment_id
        self.next_comment_id += 1

        first_run_info = real_runs[0]
        last_run_info = real_runs[-1]

        def _doc_key_for_run(run_info: Dict) -> Optional[Tuple[int, int]]:
            """Compute document-order key using host para if available, else para_elem."""
            host_para = run_info.get('host_para_elem')
            para = run_info.get('para_elem', para_elem)
            key = None
            if host_para is not None:
                key = self._get_run_doc_key(run_info.get('elem'), host_para)
            if key is None and para is not None and para is not host_para:
                key = self._get_run_doc_key(run_info.get('elem'), para)
            return key

        start_key = _doc_key_for_run(first_run_info)
        end_key = _doc_key_for_run(last_run_info)

        if start_key is not None and end_key is None:
            # No valid end run to anchor range - fallback to reference-only comment
            fallback_para = (first_run_info.get('host_para_elem') or
                             first_run_info.get('para_elem') or
                             para_elem)
            if fallback_para is None:
                print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                return 'fallback'
            ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                <w:commentReference w:id="{comment_id}"/>
            </w:r>'''
            fallback_para.append(etree.fromstring(ref_xml))
            comment_text = f"[FALLBACK]Reference-only comment: {violation_reason}"
            if revised_text:
                comment_text += f"\nSuggestion: {revised_text}"
            self.comments.append({
                'id': comment_id,
                'text': comment_text,
                'author': author
            })
            return 'success'

        if start_key is not None and end_key is not None and end_key < start_key:
            candidate = None
            for run_info in reversed(real_runs):
                cand_key = _doc_key_for_run(run_info)
                if cand_key is not None and cand_key >= start_key:
                    candidate = run_info
                    break
            if candidate is None:
                fallback_para = (first_run_info.get('host_para_elem') or
                                 first_run_info.get('para_elem') or
                                 para_elem)
                if fallback_para is None:
                    print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                    return 'fallback'
                ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                    <w:commentReference w:id="{comment_id}"/>
                </w:r>'''
                fallback_para.append(etree.fromstring(ref_xml))
                comment_text = f"[FALLBACK]Reference-only comment: {violation_reason}"
                if revised_text:
                    comment_text += f"\nSuggestion: {revised_text}"
                self.comments.append({
                    'id': comment_id,
                    'text': comment_text,
                    'author': author
                })
                if self.verbose:
                    print("  [Debug] ParaId order inverted; fallback to reference-only comment")
                return 'success'
            if self.verbose:
                print("  [Debug] ParaId order inverted; adjusted end run")
            last_run_info = candidate
            # Clamp match_end to avoid splitting beyond the adjusted end run
            match_end = min(match_end, last_run_info.get('end', match_end))

        first_run = first_run_info.get('elem')
        last_run = last_run_info.get('elem')
        rPr_xml = self._get_rPr_xml(first_run_info.get('rPr'))
        no_split_start = first_run_info.get('is_equation', False)
        no_split_end = last_run_info.get('is_equation', False)

        # Host paragraph anchors (for vMerge continue cases where runs are reused)
        start_host_para = first_run_info.get('host_para_elem')
        end_host_para = last_run_info.get('host_para_elem')
        start_use_host = start_host_para is not None and start_host_para is not first_run_info.get('para_elem')
        end_use_host = end_host_para is not None and end_host_para is not last_run_info.get('para_elem')

        # If any run is a host mismatch (vMerge continue reuse), prefer host-para anchors
        has_host_mismatch = any(
            r.get('host_para_elem') is not None and r.get('host_para_elem') is not r.get('para_elem')
            for r in real_runs
        )
        if has_host_mismatch:
            start_use_host = True
            end_use_host = True

        if first_run is None:
            start_use_host = True
            no_split_start = True
        if last_run is None:
            end_use_host = True
            no_split_end = True

        if start_use_host and start_host_para is None:
            start_host_para = first_run_info.get('para_elem', para_elem)
        if end_use_host and end_host_para is None:
            end_host_para = last_run_info.get('para_elem', para_elem)

        # Choose a reference paragraph that actually has visible content.
        reference_para = None
        for run_info in reversed(real_runs):
            para_candidate = run_info.get('host_para_elem')
            if para_candidate is None:
                para_candidate = run_info.get('para_elem')
            if para_candidate is None:
                continue
            try:
                _, para_text = self._collect_runs_info_original(para_candidate)
                if para_text.strip():
                    reference_para = para_candidate
                    break
            except Exception:
                continue
        if reference_para is None:
            reference_para = end_host_para if end_host_para is not None else start_host_para
            if reference_para is None:
                reference_para = para_elem

        # If host anchors are inverted, fallback to reference-only comment
        if start_use_host and end_use_host and start_host_para is not None and reference_para is not None:
            self._init_para_order()
            start_idx = self._para_order.get(id(start_host_para))
            end_idx = self._para_order.get(id(reference_para))
            if start_idx is not None and end_idx is not None and end_idx < start_idx:
                fallback_para = start_host_para or para_elem
                if fallback_para is None:
                    print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                    return 'fallback'
                ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                    <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                    <w:commentReference w:id="{comment_id}"/>
                </w:r>'''
                fallback_para.append(etree.fromstring(ref_xml))
                comment_text = f"[FALLBACK]Reference-only comment: {violation_reason}"
                if revised_text:
                    comment_text += f"\nSuggestion: {revised_text}"
                self.comments.append({
                    'id': comment_id,
                    'text': comment_text,
                    'author': author
                })
                if self.verbose:
                    print("  [Debug] ParaId order inverted; fallback to reference-only comment")
                return 'success'

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
        if start_use_host:
            # Insert commentRangeStart at start of host paragraph (no run split)
            target_para = start_host_para
            if target_para is None:
                print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                return 'fallback'
            insert_idx = 0
            for i, child in enumerate(list(target_para)):
                if child.tag == f'{{{NS["w"]}}}pPr':
                    continue
                insert_idx = i
                break
            target_para.insert(insert_idx, range_start)
        elif start_revision is not None:
            # Start is inside revision: insert commentRangeStart before revision container
            parent = start_revision.getparent()
            if parent is not None:
                idx = list(parent).index(start_revision)
                parent.insert(idx, range_start)
        elif no_split_start:
            # Start is in a non-text run (equation): insert range start before the run
            parent = first_run.getparent() if first_run is not None else None
            if parent is None:
                return 'fallback'
            try:
                idx = list(parent).index(first_run)
            except ValueError:
                return 'fallback'
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
                run_or_container = self._create_run(before_text, rPr_xml)
                if run_or_container.tag == 'container':
                    # Unwrap container: insert each child run
                    for child_run in run_or_container:
                        parent.insert(idx, child_run)
                        idx += 1
                else:
                    parent.insert(idx, run_or_container)
                    idx += 1
                
                # Update first_run's text content (remove before_text portion)
                t_elem = first_run.find('w:t', NS)
                if t_elem is not None and t_elem.text:
                    t_elem.text = t_elem.text[len(before_text):]
            
            # Insert commentRangeStart before the (possibly modified) first_run
            parent.insert(idx, range_start)
        
        # === Handle END position ===
        if end_use_host:
            # Insert commentRangeEnd and reference at end of host paragraph (no run split)
            target_para = reference_para if reference_para is not None else end_host_para
            if target_para is None:
                print("  [Warning] Reference-only fallback failed: no anchor paragraph")
                return 'fallback'
            if self.verbose and reference_para is not None and end_host_para is not None and reference_para is not end_host_para:
                print("  [Debug] End anchor moved to previous non-empty paragraph")
            target_para.append(range_end)
            target_para.append(comment_ref)
        elif end_revision is not None:
            # End is inside revision: insert commentRangeEnd after revision container
            parent = end_revision.getparent()
            if parent is not None:
                idx = list(parent).index(end_revision)
                parent.insert(idx + 1, range_end)
                parent.insert(idx + 2, comment_ref)
        elif no_split_end:
            # End is in a non-text run (equation): insert range end after the run
            parent = last_run.getparent() if last_run is not None else None
            if parent is None:
                return 'fallback'
            try:
                idx = list(parent).index(last_run)
            except ValueError:
                return 'fallback'
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
                run_or_container = self._create_run(after_text, rPr_xml)
                if run_or_container.tag == 'container':
                    # Unwrap container: insert each child run
                    insert_pos = idx + 3
                    for child_run in run_or_container:
                        parent.insert(insert_pos, child_run)
                        insert_pos += 1
                else:
                    parent.insert(idx + 3, run_or_container)
            else:
                # No split needed: insert commentRangeEnd and reference after last_run
                parent.insert(idx + 1, range_end)
                parent.insert(idx + 2, comment_ref)

        # Record comment content
        if fallback_reason:
            comment_text = f"[FALLBACK]{fallback_reason} {violation_reason}"
        else:
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

            # Add paragraph with formatted text (handle <sup>/<sub> markup)
            p = etree.SubElement(comment_elem, f'{{{NS["w"]}}}p')
            
            # Parse text for superscript/subscript markup
            # sanitize_xml_string first to remove illegal control characters
            sanitized_text = sanitize_xml_string(comment['text'])
            segments = self._parse_formatted_text(sanitized_text)
            
            for segment_text, vert_align in segments:
                if not segment_text:
                    continue
                
                r = etree.SubElement(p, f'{{{NS["w"]}}}r')
                
                # Add run properties with vertAlign if needed
                if vert_align:
                    rPr = etree.SubElement(r, f'{{{NS["w"]}}}rPr')
                    vert_elem = etree.SubElement(rPr, f'{{{NS["w"]}}}vertAlign')
                    vert_elem.set(f'{{{NS["w"]}}}val', vert_align)
                
                t = etree.SubElement(r, f'{{{NS["w"]}}}t')
                # Preserve whitespace (Word drops leading/trailing spaces without this)
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                t.text = segment_text

        if self.verbose:
            try:
                total = len(comments_xml.findall(f'{{{NS["w"]}}}comment'))
                print(f"[Comments] comments.xml count after merge: {total}")
            except Exception:
                pass
        
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

        if self.verbose:
            try:
                print(f"[Comments] Saved {len(self.comments)} comment(s) to comments.xml")
            except Exception:
                pass
    
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
            numbering_variants = build_numbering_variants(violation_text)
            
            # Try original text first (using revision-free view) - single paragraph
            for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                runs_info_orig, _ = self._collect_runs_info_original(para)

                pos, matched_override = self._find_in_runs_with_normalization(
                    runs_info_orig, violation_text
                )
                if pos != -1:
                    target_para = para
                    matched_runs_info = runs_info_orig
                    matched_start = pos
                    if matched_override is not None:
                        violation_text = matched_override
                    break
            
            # Fallback 1: Try stripping auto-numbering if original match failed
            if target_para is None:
                for stripped_violation, strip_mode in numbering_variants:
                    for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                        runs_info_orig, _ = self._collect_runs_info_original(para)

                        pos, matched_override = self._find_in_runs_with_normalization(
                            runs_info_orig, stripped_violation
                        )
                        if pos != -1:
                            target_para = para
                            matched_runs_info = runs_info_orig
                            matched_start = pos
                            numbering_stripped = True
                            violation_text = matched_override or stripped_violation

                            # Handle revised_text for replace operation
                            if item.fix_action == 'replace':
                                stripped_revised, revised_has_numbering = strip_numbering_by_mode(revised_text, strip_mode)
                                if revised_has_numbering:
                                    # Both have numbering: strip both
                                    revised_text = stripped_revised
                                # else: Only violation_text had numbering, keep revised_text as-is

                            break
                    if target_para is not None:
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
                    # Build search attempts: original, numbering-stripped variants, and newline-unescaped
                    search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                    for stripped_violation, strip_mode in numbering_variants:
                        search_attempts.append((stripped_violation, strip_mode))
                    if '\\n' in violation_text:
                        search_attempts.append((violation_text.replace('\\n', '\n'), None))
                    if self._is_table_mode(cross_runs) and violation_text.startswith('["'):
                        stripped_row_only, was_row_stripped = strip_table_row_number_only(violation_text)
                        if was_row_stripped:
                            search_attempts.append((stripped_row_only, "table_row_number"))
                        stripped_table_text, was_table_stripped = strip_table_row_numbering(violation_text)
                        if was_table_stripped and stripped_table_text != stripped_row_only:
                            search_attempts.append((stripped_table_text, "table_row"))

                    for search_text, strip_mode in search_attempts:
                        if self._is_table_mode(cross_runs):
                            pos = cross_text.find(search_text)
                            matched_override = None
                            if pos == -1 and search_text:
                                normalized_table_text, norm_to_orig = self._normalize_table_text_for_search(cross_runs)
                                pos_norm = normalized_table_text.find(search_text)
                                if pos_norm != -1:
                                    norm_end = pos_norm + len(search_text) - 1
                                    if 0 <= pos_norm < len(norm_to_orig) and 0 <= norm_end < len(norm_to_orig):
                                        orig_start = norm_to_orig[pos_norm]
                                        orig_end = norm_to_orig[norm_end] + 1
                                        pos = orig_start
                                        matched_override = cross_text[orig_start:orig_end]
                        else:
                            pos, matched_override = self._find_in_runs_with_normalization(
                                cross_runs, search_text
                            )

                        if pos != -1:
                            # Found match across paragraphs
                            target_para = anchor_para  # Use anchor as reference
                            matched_runs_info = cross_runs
                            matched_start = pos
                            is_cross_paragraph = True
                            if matched_override is not None:
                                violation_text = matched_override
                            else:
                                violation_text = search_text

                            if strip_mode:
                                numbering_stripped = True
                                if item.fix_action == 'replace':
                                    if strip_mode == "table_row_number":
                                        stripped_revised, revised_has_numbering = strip_table_row_number_only(revised_text)
                                    elif strip_mode == "table_row":
                                        stripped_revised, revised_has_numbering = strip_table_row_numbering(revised_text)
                                    else:
                                        stripped_revised, revised_has_numbering = strip_numbering_by_mode(revised_text, strip_mode)
                                    if revised_has_numbering:
                                        revised_text = stripped_revised

                            if self.verbose:
                                print(f"  [Success] Found in cross-paragraph mode")
                            break
            
            # Fallback 3: Try table search if violation_text looks like JSON array
            if target_para is None and violation_text.startswith('["'):
                # Normalize to remove duplicate brackets at boundaries (LLM artifacts)
                violation_text = normalize_table_json(violation_text)
                if item.fix_action == 'replace':
                    revised_text = normalize_table_json(revised_text)
                
                # Find all tables in range
                tables_in_range = self._find_tables_in_range(anchor_para, item.uuid_end)
                
                # Try with original violation_text first, then with row numbering stripped
                search_attempts = [
                    (violation_text, None),  # Original text, not stripped
                ]

                stripped_row_only, was_row_stripped = strip_table_row_number_only(violation_text)
                if was_row_stripped:
                    search_attempts.append((stripped_row_only, "table_row_number"))

                stripped_table_text, was_table_stripped = strip_table_row_numbering(violation_text)
                if was_table_stripped and stripped_table_text != stripped_row_only:
                    search_attempts.append((stripped_table_text, "table_row"))
                
                for search_text, strip_mode in search_attempts:
                    for table_idx, table_elem in enumerate(tables_in_range):
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
                            matched_text_override = None
                            if pos == -1 and search_text:
                                # Fallback: normalize table text by stripping per-paragraph whitespace
                                # to match parse_document.py behavior, and map back to original indices.
                                normalized_table_text, norm_to_orig = self._normalize_table_text_for_search(table_runs)
                                pos_norm = normalized_table_text.find(search_text)
                                if pos_norm != -1:
                                    norm_end = pos_norm + len(search_text) - 1
                                    if 0 <= pos_norm < len(norm_to_orig) and 0 <= norm_end < len(norm_to_orig):
                                        orig_start = norm_to_orig[pos_norm]
                                        orig_end = norm_to_orig[norm_end] + 1
                                        pos = orig_start
                                        matched_text_override = table_text[orig_start:orig_end]
                            
                            # Debug logging for specific content
                            if pos == -1 and DEBUG_MARKER and (DEBUG_MARKER in table_text or DEBUG_MARKER in search_text):
                                print(f"\n  [DEBUG] Table matching failed for row containing '{DEBUG_MARKER}':")
                                
                                # Extract only the row containing DEBUG_MARKER from table_text
                                def extract_matching_row(text, marker):
                                    """Extract the row containing the marker from table JSON format."""
                                    # Table format: ["cell1", "cell2"], ["cell3", "cell4"]
                                    # Split by row boundaries '], ['
                                    if not text.startswith('["'):
                                        return None
                                    
                                    # Remove outer brackets and split by row separator
                                    rows_text = text[2:-2] if text.endswith('"]') else text[2:]
                                    rows = rows_text.split('"], ["')
                                    
                                    for row in rows:
                                        if marker in row:
                                            return '["' + row + '"]'
                                    return None
                                
                                table_row = extract_matching_row(table_text, DEBUG_MARKER)
                                search_row = extract_matching_row(search_text, DEBUG_MARKER)
                                
                                if table_row:
                                    print(f"  [DEBUG] Table row content:")
                                    print(f"    {table_row}")
                                else:
                                    print(f"  [DEBUG] Table content (marker '{DEBUG_MARKER}' not found in individual row):")
                                    print(f"    {table_text[:300]}...")
                                
                                if search_row:
                                    print(f"  [DEBUG] Searching for:")
                                    print(f"    {search_row}")
                                else:
                                    print(f"  [DEBUG] Search content (marker '{DEBUG_MARKER}' not found in individual row):")
                                    print(f"    {search_text[:300]}...")
                                
                                # Show character-level diff if both rows found
                                if table_row and search_row:
                                    min_len = min(len(table_row), len(search_row))
                                    for i in range(min_len):
                                        if table_row[i] != search_row[i]:
                                            print(f"  [DEBUG] First difference at position {i}:")
                                            print(f"    Table: ...{repr(table_row[max(0,i-10):i+30])}...")
                                            print(f"    Search: ...{repr(search_row[max(0,i-10):i+30])}...")
                                            break
                                    else:
                                        # No difference found in common length, check length difference
                                        if len(table_row) != len(search_row):
                                            print(f"  [DEBUG] Length mismatch: table={len(table_row)}, search={len(search_row)}")
                            
                            if pos != -1:
                                # Found match in this table!
                                target_para = first_table_para  # Use first para as anchor
                                matched_runs_info = table_runs
                                matched_start = pos
                                is_cross_paragraph = True  # Table mode is always cross-paragraph
                                
                                # Update violation_text to the matched version(stripped or not)
                                violation_text = matched_text_override or search_text
                                
                                # For replace operations, also strip row numbering from revised_text
                                if strip_mode and item.fix_action == 'replace':
                                    if strip_mode == "table_row_number":
                                        stripped_revised, revised_was_stripped = strip_table_row_number_only(revised_text)
                                    else:
                                        stripped_revised, revised_was_stripped = strip_table_row_numbering(revised_text)
                                    if revised_was_stripped:
                                        revised_text = stripped_revised
                                
                                if self.verbose:
                                    if strip_mode:
                                        print(f"  [Success] Found in table after stripping row numbering")
                                    else:
                                        print(f"  [Success] Found in table (JSON format)")
                                break

                        except (ValueError, KeyError, IndexError, AttributeError) as e:
                            # If table processing fails, continue to next table
                            if self.verbose:
                                print(f"  [Warning] Skipping table: {e}")
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
                        target_para, matched_runs_info, matched_start, matched_text, strip_mode = result
                        # Update violation_text with the actual matched text (handles fallback normalization)
                        violation_text = matched_text
                        # Cell content is always treated as single-paragraph for now
                        is_cross_paragraph = False
                        if strip_mode:
                            numbering_stripped = True
                            if item.fix_action == 'replace':
                                stripped_revised, revised_has_numbering = strip_numbering_by_mode(revised_text, strip_mode)
                                if revised_has_numbering:
                                    revised_text = stripped_revised
                        
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
                                
                                # Build search attempts for raw text in cell
                                search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                                search_attempts.extend(build_numbering_variants(violation_text))
                                if '\\n' in violation_text:
                                    newline_text = violation_text.replace('\\n', '\n')
                                    search_attempts.append((newline_text, None))
                                    search_attempts.extend(build_numbering_variants(newline_text))

                                # De-duplicate while preserving order
                                seen = set()
                                deduped_attempts: List[Tuple[str, Optional[str]]] = []
                                for text, mode in search_attempts:
                                    key = (text, mode)
                                    if key in seen:
                                        continue
                                    seen.add(key)
                                    deduped_attempts.append((text, mode))

                                match_pos = -1
                                matched_search_text = violation_text
                                matched_strip_mode: Optional[str] = None
                                for search_text, strip_mode in deduped_attempts:
                                    match_pos = cell_normalized.find(search_text)
                                    if match_pos != -1:
                                        matched_search_text = search_text
                                        matched_strip_mode = strip_mode
                                        break

                                if match_pos != -1:
                                    # Found match in this cell!
                                    if self.verbose:
                                        print(f"  [Success] Found in table cell (raw text match)")
                                    if matched_strip_mode:
                                        numbering_stripped = True
                                        violation_text = matched_search_text
                                        if item.fix_action == 'replace':
                                            stripped_revised, revised_has_numbering = strip_numbering_by_mode(revised_text, matched_strip_mode)
                                            if revised_has_numbering:
                                                revised_text = stripped_revised
                                    else:
                                        violation_text = matched_search_text
                                    
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
                                        if current_offset <= match_pos < current_offset + para_len:
                                            matched_para = cell_para
                                            matched_para_runs = para_runs
                                            matched_para_start = match_pos - current_offset
                                            break
                                        
                                        current_offset += para_len + 1  # +1 for \n separator
                                    
                                    if matched_para is not None:
                                        # Use the matched paragraph
                                        target_para = matched_para
                                        matched_runs_info = matched_para_runs
                                        matched_start = matched_para_start
                                        is_cross_paragraph = False  # Single para in cell
                                        break
                                else:
                                    # Debug: log snippet from marker on non-JSON match failure
                                    marker_idx = cell_combined_text.find(DEBUG_MARKER)
                                    if DEBUG_MARKER and marker_idx != -1:
                                        snippet_len = len(DEBUG_MARKER) + 60
                                        snippet = cell_combined_text[marker_idx:marker_idx + snippet_len]
                                        print(f"  [DEBUG] Non-JSON cell content from marker: {repr(snippet)}")

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
                            
                            # Try multiple search patterns (original and stripped numbering)
                            # Use build_numbering_variants() to support multi-line numbering removal
                            search_attempts: List[Tuple[str, Optional[str]]] = [(violation_text, None)]
                            numbering_variants = build_numbering_variants(violation_text)
                            search_attempts.extend(numbering_variants)
                            
                            match_pos = -1
                            matched_text = violation_text
                            matched_override = None
                            matched_is_stripped = False
                            
                            for search_text, is_stripped in search_attempts:
                                match_pos, matched_override = self._find_in_runs_with_normalization(
                                    all_runs, search_text
                                )
                                if match_pos != -1:
                                    matched_text = matched_override or search_text
                                    matched_is_stripped = is_stripped
                                    break
                            
                            if match_pos != -1:
                                # Found match in this segment!
                                target_para = body_paras_data[0][0]  # Use first para as anchor
                                matched_runs_info = all_runs
                                matched_start = match_pos
                                is_cross_paragraph = len(body_paras_data) > 1
                                violation_text = matched_text  # Update violation_text to matched version
                                numbering_stripped = matched_is_stripped
                                
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
                            reason = "Violation text not found(C)"  # Table/body boundary crossed
                        else:
                            reason = f"Boundary error: {boundary_error}"

                        # Debug: log body paragraph content from marker on boundary failure
                        if DEBUG_MARKER:
                            try:
                                for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                                    if self._is_paragraph_in_table(para):
                                        continue
                                    _, para_text = self._collect_runs_info_original(para)
                                    marker_idx = para_text.find(DEBUG_MARKER)
                                    if marker_idx != -1:
                                        snippet_len = len(DEBUG_MARKER) + 60
                                        snippet = para_text[marker_idx:marker_idx + snippet_len]
                                        print(f"  [DEBUG] Body search target: {repr(violation_text)}")
                                        break
                            except Exception:
                                # Debug logging should never break apply flow
                                pass

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
                    # Debug: log body paragraph content from marker on overall match failure
                    if DEBUG_MARKER:
                        try:
                            for para in self._iter_paragraphs_in_range(anchor_para, item.uuid_end):
                                if self._is_paragraph_in_table(para):
                                    continue
                                _, para_text = self._collect_runs_info_original(para)
                                marker_idx = para_text.find(DEBUG_MARKER)
                                if marker_idx != -1:
                                    snippet_len = len(DEBUG_MARKER) + 60
                                    snippet = para_text[marker_idx:marker_idx + snippet_len]
                                    print(f"  [DEBUG] Body search target: {repr(violation_text)}")
                                    break
                        except Exception:
                            # Debug logging should never break apply flow
                            pass

                    # Calculate total segments for error message
                    # Variables are initialized in their respective code paths above
                    # If not set, default to empty lists
                    tables_count = len(tables_in_range) if 'tables_in_range' in dir() else 0
                    body_count = len(body_segments) if 'body_segments' in dir() else 0
                    total_segments = tables_count + body_count
                    
                    if total_segments == 1:
                        reason = "Violation text not found(S)"  # Single segment
                    else:
                        reason = f"Violation text not found(M)"  # Multiple segments
                    
                    # For manual fix_action, text not found is expected (not an error)
                    if item.fix_action == 'manual':
                        self._apply_error_comment(anchor_para, item)
                        return EditResult(
                            success=True,
                            item=item,
                            error_message=reason,
                            warning=True
                        )
                    else:
                        # For delete/replace, text not found is an error
                        self._apply_error_comment(anchor_para, item)
                        return EditResult(False, item, reason)
            
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

                    # Check if match spans multiple table rows
                    if self._check_cross_row_boundary(real_runs):
                        # Multi-row delete: use cell-by-cell deletion
                        if self.verbose:
                            print(f"  [Multi-row] Applying cell-by-cell deletion")
                        success_status = self._apply_delete_multi_cell(
                            real_runs, violation_text,
                            item.violation_reason, item_author, matched_start,
                            item
                        )
                    # Check if match spans multiple table cells (within same row)
                    elif self._check_cross_cell_boundary(real_runs):
                        # Multi-cell delete: use cell-by-cell deletion
                        if self.verbose:
                            print(f"  [Multi-cell] Applying cell-by-cell deletion")
                        success_status = self._apply_delete_multi_cell(
                            real_runs, violation_text,
                            item.violation_reason, item_author, matched_start,
                            item
                        )
                    elif len(para_elems) > 1:
                        # Actually spans multiple paragraphs
                        if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
                            if self.verbose:
                                print(f"  [Cross-paragraph] delete spans {len(para_elems)} paragraphs, fallback to comment")
                            success_status = 'cross_paragraph_fallback'
                        else:
                            if self.verbose:
                                print(f"  [Cross-paragraph] delete spans {len(para_elems)} paragraphs, applying cross-paragraph delete")
                            success_status = self._apply_delete_cross_paragraph(
                                matched_runs_info, matched_start, violation_text,
                                item.violation_reason, item_author
                            )
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

                    # Check if match spans multiple table rows
                    is_cross_row = self._check_cross_row_boundary(real_runs)

                    # Detect JSON cell boundaries even if only one cell has text runs.
                    # This handles cases where empty cells (e.g., stripped row numbers)
                    # produce boundary markers but no runs with cell_elem.
                    has_cell_boundary = any(r.get('is_cell_boundary') for r in affected)
                    is_cross_cell = self._check_cross_cell_boundary(real_runs) or has_cell_boundary or is_cross_row

                    if is_cross_row and self.verbose:
                        print(f"  [Cross-row] replace spans multiple rows, trying cell-by-cell extraction...")
                    if self.verbose and has_cell_boundary and not self._check_cross_cell_boundary(real_runs):
                        print(f"  [Cross-cell] Boundary markers detected (empty cells), forcing cell extraction")

                    if success_status is None and is_cross_cell:
                        # Try to extract single-cell edit first
                        if self.verbose:
                            print(f"  [Cross-cell] Detected cross-cell match, trying cell-by-cell extraction...")
                        
                        single_cell = self._try_extract_single_cell_edit(
                            violation_text, revised_text, affected, matched_start
                        )
                        
                        if single_cell:
                            # Successfully extracted single-cell edit - all changes in one cell
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
                                        print(f"  [Single-cell] All changes in one cell, applied track change")
                                else:
                                    success_status = 'cross_cell_fallback'
                            else:
                                success_status = 'cross_cell_fallback'
                        else:
                            # Try multi-cell extraction - changes distributed across cells
                            multi_cells = self._try_extract_multi_cell_edits(
                                violation_text, revised_text, affected, matched_start
                            )
                            
                            if multi_cells:
                                # Successfully extracted multi-cell edits - each change within its own cell
                                if self.verbose:
                                    print(f"  [Multi-cell] Found {len(multi_cells)} cells with changes, applying track changes...")
                                
                                success_count = 0
                                failed_cells = []  # List of (cell_edit, cell_para, error_reason)
                                first_success_para = None
                                
                                for cell_edit in multi_cells:
                                    # Collect all paragraphs in this cell (handle multi-paragraph cells)
                                    cell_paras = set()
                                    for run in cell_edit['cell_runs']:
                                        para = run.get('para_elem')
                                        if para is not None:
                                            cell_paras.add(para)
                                    
                                    if not cell_paras:
                                        failed_cells.append((cell_edit, None, "No paragraph found"))
                                        continue  # Continue processing next cell
                                    
                                    # Build combined runs/text from all paragraphs in this cell
                                    cell_runs_info = []
                                    cell_text_parts = []
                                    pos = 0
                                    
                                    # Sort paragraphs by document order (use first run's start position as key)
                                    para_list = []
                                    for para in cell_paras:
                                        # Find first run from this paragraph to get its position
                                        first_run_pos = None
                                        for run in cell_edit['cell_runs']:
                                            if run.get('para_elem') is para:
                                                first_run_pos = run.get('start', 0)
                                                break
                                        para_list.append((first_run_pos or 0, para))
                                    para_list.sort(key=lambda x: x[0])
                                    
                                    first_para = None
                                    for para_idx, (_, para) in enumerate(para_list):
                                        if first_para is None:
                                            first_para = para
                                        
                                        # Collect runs for this paragraph
                                        para_runs, para_text = self._collect_runs_info_original(para)
                                        
                                        # Add paragraph boundary (not for first paragraph)
                                        if para_idx > 0 and cell_runs_info:
                                            cell_runs_info.append({
                                                'text': '\n',
                                                'start': pos,
                                                'end': pos + 1,
                                                'para_elem': para,
                                                'is_para_boundary': True
                                            })
                                            cell_text_parts.append('\n')
                                            pos += 1
                                        
                                        # Add paragraph runs with adjusted positions
                                        for run in para_runs:
                                            run_copy = dict(run)
                                            run_copy['para_elem'] = para
                                            run_copy['start'] = run['start'] + pos
                                            run_copy['end'] = run['end'] + pos
                                            cell_runs_info.append(run_copy)
                                        
                                        cell_text_parts.append(para_text)
                                        pos += len(para_text)
                                    
                                    cell_text = ''.join(cell_text_parts)
                                    cell_pos = cell_text.find(cell_edit['cell_violation'])
                                    
                                    if cell_pos == -1:
                                        failed_cells.append((cell_edit, first_para, "Text not found in cell"))
                                        continue  # Continue processing next cell
                                    
                                    # Apply replace to this cell WITHOUT comment (skip_comment=True)
                                    cell_status = self._apply_replace(
                                        first_para,
                                        cell_edit['cell_violation'],
                                        cell_edit['cell_revised'],
                                        item.violation_reason,
                                        cell_runs_info,
                                        cell_pos,
                                        item_author,
                                        skip_comment=True  # Skip comment for individual cells
                                    )
                                    
                                    if cell_status != 'success':
                                        failed_cells.append((cell_edit, first_para, f"Apply failed: {cell_status}"))
                                        continue  # Continue processing next cell
                                    
                                    # Success - track for overall comment
                                    success_count += 1
                                    if first_success_para is None:
                                        first_success_para = first_para
                                
                                # Handle results based on success/failure counts
                                if success_count > 0:
                                    # At least one cell succeeded - add overall comment
                                    if first_success_para is not None:
                                        comment_id = self.next_comment_id
                                        self.next_comment_id += 1
                                        
                                        # Insert commentReference at end of first successful paragraph
                                        ref_xml = f'''<w:r xmlns:w="{NS['w']}">
                                            <w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>
                                            <w:commentReference w:id="{comment_id}"/>
                                        </w:r>'''
                                        first_success_para.append(etree.fromstring(ref_xml))
                                        
                                        # Record comment with -R suffix author
                                        self.comments.append({
                                            'id': comment_id,
                                            'text': item.violation_reason,
                                            'author': f"{item_author}-R"
                                        })
                                    
                                    # Add individual comments for failed cells
                                    for cell_edit, cell_para, error_reason in failed_cells:
                                        if cell_para is not None:
                                            self._apply_cell_fallback_comment(
                                                cell_para,
                                                cell_edit['cell_violation'],
                                                cell_edit['cell_revised'],
                                                error_reason,
                                                item
                                            )
                                    
                                    success_status = 'success'
                                    if self.verbose:
                                        if failed_cells:
                                            print(f"  [Multi-cell] Partial success: {success_count}/{len(multi_cells)} cells processed, {len(failed_cells)} failed")
                                        else:
                                            print(f"  [Multi-cell] Successfully applied track changes to all {len(multi_cells)} cells")
                                else:
                                    # All cells failed - fallback to overall comment
                                    success_status = 'cross_cell_fallback'
                            else:
                                # Changes cross cell boundaries - fallback to comment
                                if self.verbose:
                                    print(f"  [Cross-cell] Changes cross cell boundaries, fallback to comment")
                                success_status = 'cross_cell_fallback'
                        if success_status is None and is_cross_row:
                            # Cell extraction failed for cross-row content
                            if self.verbose:
                                print(f"  [Cross-row] Cell extraction failed, fallback to comment")
                            success_status = 'cross_row_fallback'
                    elif len(para_elems) > 1:
                        # Actually spans multiple paragraphs
                        if self._is_table_mode(real_runs) or any(r.get('cell_elem') is not None for r in real_runs):
                            if self.verbose:
                                print(f"  [Cross-paragraph] replace spans {len(para_elems)} paragraphs, fallback to comment")
                            # Debug: log snippet from marker on cross-paragraph fallback
                            try:
                                if DEBUG_MARKER:
                                    combined_text = ''.join(r.get('text', '') for r in matched_runs_info)
                                    marker_idx = combined_text.find(DEBUG_MARKER)
                                    if marker_idx != -1:
                                        snippet_len = len(DEBUG_MARKER) + 60
                                        snippet = combined_text[marker_idx:marker_idx + snippet_len]
                                        print(f"  [DEBUG] Cross-paragraph content from marker: {repr(snippet)}")
                            except Exception:
                                # Debug logging should never break apply flow
                                pass
                            success_status = 'cross_paragraph_fallback'
                        else:
                            if self.verbose:
                                print(f"  [Cross-paragraph] replace spans {len(para_elems)} paragraphs, applying cross-paragraph replace")
                            success_status = self._apply_replace_cross_paragraph(
                                matched_runs_info, matched_start, violation_text,
                                revised_text, item.violation_reason, item_author
                            )
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
                reason = "Multiple changes overlap."
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
                reason = (
                    "Cross-paragraph delete/replace not supported"
                )
                # Apply manual comment instead
                manual_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph,
                    fallback_reason=reason
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
                reason = "Cross-row edit not supported"
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
                reason = "Cross-cell edit not supported"
                # Apply manual comment instead (same row comment is supported)
                manual_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph,
                    fallback_reason=reason
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
            elif success_status == 'equation_fallback':
                # Equation-only content cannot be edited - fallback to manual comment
                reason = "Equation cannot be edited"
                manual_status = self._apply_manual(
                    target_para, violation_text,
                    item.violation_reason, revised_text,
                    matched_runs_info, matched_start,
                    item_author,
                    is_cross_paragraph,
                    fallback_reason=reason
                )
                if manual_status == 'success':
                    if self.verbose:
                        print(f"  [Equation-only] Applied comment instead")
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
                reason = "No editable runs found"
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

        # 2.5 Initialize paragraph order cache
        self._init_para_order()
        
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

        # 5. Debug: count comment markers in document
        if self.verbose:
            try:
                rs = self._xpath(self.body_elem, './/w:commentRangeStart')
                re = self._xpath(self.body_elem, './/w:commentRangeEnd')
                rf = self._xpath(self.body_elem, './/w:commentReference')
                print(
                    f"[Comments] Markers in document.xml: "
                    f"rangeStart={len(rs)} rangeEnd={len(re)} reference={len(rf)}"
                )
            except Exception:
                pass

        # 6. Save comments
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
