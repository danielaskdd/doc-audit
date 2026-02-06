"""Audit edit applier composed from focused mixins."""

import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Generator, Iterator, List, Optional, Tuple

from docx import Document
from lxml import etree

from .common import (
    COMMENTS_CONTENT_TYPE,
    COMMENTS_REL_TYPE,
    NS,
    EditItem,
    EditResult,
    build_numbering_variants,
    extract_longest_segment,
    format_text_preview,
    normalize_table_json,
    strip_numbering_by_mode,
    strip_table_row_number_only,
    strip_table_row_numbering,
)
from .navigation import DocxNavigationMixin
from .table_ops import DocxTableMixin
from .edit_actions import DocxEditActionsMixin


class AuditEditApplier(DocxNavigationMixin, DocxTableMixin, DocxEditActionsMixin):
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

        def _category_suffix(self, item: EditItem) -> str:
            """Normalize category suffix for author; default to 'uncategorized'."""
            category = (item.category or '').strip()
            return category if category else 'uncategorized'

        def _author_for_item(self, item: EditItem) -> str:
            """Build author name using base author + category suffix."""
            return f"{self.author}-{self._category_suffix(item)}"

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

                # Handle mixed body/table content (violation_text contains <table> tags)
                has_table_tag = '<table>' in violation_text or '</table>' in violation_text
                if has_table_tag:
                    if item.fix_action in ('delete', 'replace'):
                        reason = "Mixed body/table content is invalid"
                        self._apply_fallback_comment(anchor_para, item, reason)
                        if self.verbose:
                            print(f"  [Mixed content] {reason}")
                        return EditResult(success=True, item=item, error_message=reason, warning=True)
                    else:  # manual
                        longest = extract_longest_segment(violation_text)
                        if longest:
                            if self.verbose:
                                print(f"  [Mixed content] Extracted longest segment for comment: '{format_text_preview(longest)}'")
                            violation_text = longest

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
                if target_para is None and (violation_text.startswith('["') or violation_text.startswith('[["')):
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
                if target_para is None and not violation_text.startswith('["') and not violation_text.startswith('[["'):
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
