"""Navigation and search helpers for Word XML trees."""

import re
from typing import Dict, Generator, Iterator, List, Optional, Tuple

from lxml import etree

from .common import NS, build_numbering_variants, json_escape


class DocxNavigationMixin:
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
