from .common import (
    NS, EditItem, EditResult, json_escape, build_numbering_variants, DEBUG_MARKER
)
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Generator, Iterator
import json
import hashlib
from docx import Document
from lxml import etree


class NavigationMixin:
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

        def iter_original_with_escaped_indices(original_text: str) -> Iterator[Tuple[str, int]]:
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
