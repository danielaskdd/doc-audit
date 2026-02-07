#!/usr/bin/env python3
"""
ABOUTME: Applies audit results to Word documents with track changes and comments
ABOUTME: Reads JSONL export from audit report and modifies the source document
"""

import json
import re
from dataclasses import dataclass
from typing import List, Tuple, Optional


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

# Table tag pattern for detecting mixed body/table content from parse_document.py
# Matches: <table> and </table> tags that wrap table JSON content
TABLE_TAG_PATTERN = re.compile(r'</?table>')

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


def extract_longest_segment(text: str) -> Optional[str]:
    """
    Split text by <table>/</table> tags and return the longest non-empty segment.

    When violation_text from the LLM contains mixed body text and table content
    (e.g., "Heading text\\n<table>JSON data</table>"), this function extracts
    the longest segment for use as the search target in manual (comment) operations.

    Args:
        text: Text that may contain <table>/</table> tags

    Returns:
        The longest non-empty stripped segment, or None if no table tags found
        or all segments are empty after stripping.

    Examples:
        "Title\\n<table>[[...long JSON...]]</table>" -> '[[...long JSON...]]'
        "<table>JSON</table>\\nBody text here" -> 'Body text here' (if longer)
        "No table tags" -> None
        "<table></table>" -> None
    """
    if '<table>' not in text and '</table>' not in text:
        return None

    segments = TABLE_TAG_PATTERN.split(text)
    clean = [s.strip() for s in segments if s.strip()]
    if not clean:
        return None
    return max(clean, key=len)


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
    if not text.startswith('["') and not text.startswith('[["'):
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

